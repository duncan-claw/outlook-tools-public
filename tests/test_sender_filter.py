"""Synthetic sender-filter tests. All addresses/names are invented fixtures."""
import test_support
import json
from collections import Counter
import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path
from unittest.mock import patch

from spam_guard import sender_filter as sf
from spam_guard.group_calendar_guard import GuardError
from spam_guard.sender_identity import score_sender

SPAM = ("xqvobzzj13@example.com", "Sample Person")
FRIEND = "friend.random9x8q@example.com"  # suspicious-looking, but a contact


class IdentityTests(unittest.TestCase):
    def test_example_is_suspicious_with_reasons(self):
        result = score_sender(*SPAM)
        self.assertTrue(result["suspicious"])
        self.assertEqual(result["signals"], ["NAME_MISMATCH", "RANDOM_LOCAL_PART", "TRAILING_DIGITS"])

    def test_no_single_signal_is_proof(self):
        self.assertFalse(score_sender("xqvobzzj13@example.com", None)["suspicious"])      # random + digits only
        self.assertFalse(score_sender("sample.person84@example.com", "Sample Person")["suspicious"])  # digits
        self.assertFalse(score_sender("margaret.smith@example.com", "Mum")["suspicious"])    # mismatch
        self.assertFalse(score_sender("abc@example.com", "x@bank.example")["suspicious"])    # spoof alone

    def test_aliases_and_initials(self):
        for address in ("sperson@example.com", "sample.p@example.com", "samper@example.com", "person@example.com"):
            self.assertFalse(score_sender(address, "Sample Person")["suspicious"], address)

    def test_transliteration_and_non_latin(self):
        self.assertFalse(score_sender("xiaoqing.zhou@example.com", "Xiaoqing Zhou")["suspicious"])
        self.assertFalse(score_sender("jose.mueller@example.com", "José Müller")["suspicious"])
        result = score_sender("qiang.wang88@example.com", "王强")
        self.assertFalse(result["suspicious"])
        self.assertIn("NAME_NOT_ASSESSABLE", result["notes"])

    def test_business_and_automated_senders(self):
        for address, name in [("noreply@ses.amazonaws.example", "Amazon Web Services"),
                              ("bounce-a8f3k2q9@mail.vendor.example", "Vendor Notifications"),
                              ("x7f3k9qz@em.shop.example", "Sarah from Shop"),
                              ("notifications@code.example", "GitHub")]:
            self.assertFalse(score_sender(address, name)["suspicious"], address)

    def test_random_alnum_with_mismatch(self):
        self.assertTrue(score_sender("jx7k2m9q@example.com", "Bob")["suspicious"])

    def test_malformed(self):
        for address in (None, "", "no-at-sign", "a@b", "a b@example.com"):
            self.assertIn("MALFORMED_ADDRESS", score_sender(address, "Sample Person")["notes"])


def msg(mid, address, name=None, folder="INBOX"):
    return {"id": mid, "internetMessageId": "<" + mid + "@synthetic>", "parentFolderId": folder,
            "receivedDateTime": "2000-01-18T00:00:00Z",
            "from": {"emailAddress": {"address": address, "name": name}}}


class FakeClient:
    def __init__(self, messages, contacts=(FRIEND,), contacts_fail=False, account=sf.ACCOUNT):
        self.messages = {m["id"]: m for m in messages}
        self.contacts = list(contacts)
        self.contacts_fail = contacts_fail
        self.account = account
        self.moves = []

    def request(self, method, path, body=None):
        if path.startswith("/me/?"):
            return {"mail": self.account}
        if path.startswith("/me/mailFolders/inbox?"):
            return {"id": "INBOX"}
        if path.startswith("/me/mailFolders/junkemail?"):
            return {"id": "JUNK"}
        if path.startswith("/me/mailFolders/deleteditems?"):
            return {"id": "DELETED"}
        if path.startswith("/me/mailFolders/recoverableitemsdeletions?"):
            return {"id": "RECOVERABLE"}
        if path.startswith("/me/mailFolders/recoverableitems"):
            raise GuardError("HTTP_404")
        mid =path.split("/me/messages/", 1)[1].split("?")[0].split("/")[0]
        if mid not in self.messages:
            raise GuardError("HTTP_404")
        if method == "POST":
            self.moves.append((mid, body["destinationId"]))
            self.messages[mid]["parentFolderId"] = body["destinationId"]
            return {"id": mid}  # immutable ID survives the move
        return dict(self.messages[mid])

    def pages(self, path):
        if path.startswith("/me/contactFolders"):
            return iter([])
        if path.startswith("/me/contacts"):
            if self.contacts_fail:
                raise GuardError("HTTP_403")
            return iter([{"emailAddresses": [{"address": a}]} for a in self.contacts])
        if path.startswith("/me/mailFolders/inbox/messages"):
            return iter([dict(m) for m in self.messages.values() if m["parentFolderId"] == "INBOX"])
        if path.startswith("/me/mailFolders/junkemail/messages"):
            return iter([dict(m) for m in self.messages.values() if m["parentFolderId"] == "JUNK"])
        raise AssertionError(path)

    def add(self, message):
        self.messages[message["id"]] = message


class FilterTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.patch = patch.object(sf, "STATE", Path(self.tmp.name))
        self.patch.start()
        self.sources = ("contacts", "local")

    def tearDown(self):
        self.patch.stop()
        self.tmp.cleanup()

    def inbox(self):
        return [msg("spam", *SPAM),
                msg("contact", FRIEND, "Sample Person"),
                msg("self", sf.ACCOUNT, "Example Owner"),     # account owner / self-sent mail
                msg("unknown", "stranger@example.org", "Example Recipient"),
                msg("broken", None)]

    def run_scan(self, client, **kw):
        kw.setdefault("sources", self.sources)
        return sf.scan(client, **kw)

    def test_dry_run_counts_both_policies_without_mutation(self):
        client = FakeClient(self.inbox(), contacts=["Friend.Random9x8q@example.com"])
        result = self.run_scan(client)
        self.assertEqual(client.moves, [])
        self.assertEqual(result["would_junk_scored"], 1)
        self.assertEqual(result["would_junk_strict"], 2)
        self.assertEqual(result["safe"], 2)
        self.assertEqual(result["malformed"], 1)

    def test_output_contains_counts_and_codes_only(self):
        result = json.dumps(self.run_scan(FakeClient(self.inbox())))
        for secret in ("qzxwobate", "Sample Person", "stranger", "spam"):
            self.assertNotIn(secret, result)

    def test_scored_apply_moves_only_suspicious_to_junk(self):
        client = FakeClient(self.inbox(), contacts=["friend.random9x8q@example.com"])
        result = self.run_scan(client, mode="scored", apply=True)
        self.assertEqual(client.moves, [("spam", "JUNK")])
        self.assertEqual(result["moved"], 1)
        report = [json.loads(l) for l in (sf.STATE / "moves.jsonl").read_text().splitlines()]
        self.assertIn("SCORED_SUSPICIOUS", report[0]["codes"])
        self.assertIn("RANDOM_LOCAL_PART", report[0]["codes"])

    def test_strict_moves_everything_not_safe_listed(self):
        client = FakeClient(self.inbox(), contacts=["friend.random9x8q@example.com"])
        self.run_scan(client, mode="strict", apply=True)
        self.assertEqual(sorted(m for m, _ in client.moves), ["spam", "unknown"])
        self.assertTrue(all(dest == "JUNK" for _, dest in client.moves))

    def test_local_allowlist_address_and_domain(self):
        (sf.STATE / "allowlist.txt").write_text("# comment\nXQVOBZZJ13@example.com\n@example.org\n@example.com\n")
        client = FakeClient(self.inbox())
        result = self.run_scan(client, mode="strict", apply=True)
        self.assertEqual(client.moves, [])
        self.assertEqual(result["allowlist_ignored_entries"], 1)  # free-mail domain refused

    def test_domain_entry_does_not_match_lookalike_suffix(self):
        safe = sf.SafeList(domains=["example.org"])
        self.assertEqual(safe.check("a@mail.example.org"), "SAFE_LOCAL_DOMAIN")
        self.assertIsNone(safe.check("a@badexample.org"))
        self.assertIsNone(safe.check("a@example.org.evil.example"))

    def test_correspondent_hash_protects(self):
        now = datetime.now(timezone.utc)
        (sf.STATE / "correspondents.json").write_text(json.dumps(
            {"refreshed_at": now.isoformat(), "hashes": [sf.digest(SPAM[0])]}))
        client = FakeClient(self.inbox())
        self.run_scan(client, mode="scored", apply=True, sources=("correspondents",))
        self.assertNotIn(("spam", "JUNK"), client.moves)

    def test_stale_or_missing_correspondents_fail_closed(self):
        client = FakeClient(self.inbox())
        with self.assertRaises(GuardError):
            self.run_scan(client, mode="strict", apply=True, sources=("correspondents",))
        old = datetime.now(timezone.utc) - timedelta(days=4)
        (sf.STATE / "correspondents.json").write_text(json.dumps({"refreshed_at": old.isoformat(), "hashes": []}))
        with self.assertRaises(GuardError):
            self.run_scan(client, mode="strict", apply=True, sources=("correspondents",))
        self.assertEqual(client.moves, [])

    def test_contacts_failure_prevents_moves(self):
        client = FakeClient(self.inbox(), contacts_fail=True)
        with self.assertRaises(GuardError):
            self.run_scan(client, mode="strict", apply=True)
        self.assertEqual(client.moves, [])

    def test_account_mismatch(self):
        with self.assertRaises(GuardError):
            self.run_scan(FakeClient(self.inbox(), account="other@example.com"))

    def test_missing_account_configuration_fails_closed(self):
        with patch.object(sf, "ACCOUNT", ""):
            with self.assertRaisesRegex(GuardError, "ACCOUNT_NOT_CONFIGURED"):
                self.run_scan(FakeClient(self.inbox()))

    def test_apply_requires_policy(self):
        with self.assertRaises(GuardError):
            self.run_scan(FakeClient(self.inbox()), mode="dry-run", apply=True)

    def test_ledger_prevents_repeat_and_respects_manual_restore(self):
        client = FakeClient(self.inbox())
        self.run_scan(client, mode="scored", apply=True)
        client.messages["spam"]["parentFolderId"] = "INBOX"   # the account owner moves it back
        result = self.run_scan(client, mode="scored", apply=True)
        self.assertEqual(len(client.moves), 1)
        self.assertEqual(result["already_processed"], 1)

    def test_restore_moves_back_and_is_never_refiltered(self):
        client = FakeClient(self.inbox())
        self.run_scan(client, mode="scored", apply=True)
        key = json.loads((sf.STATE / "moves.jsonl").read_text().splitlines()[0])["key"]
        result = sf.restore(client, keys=[key[:12]])
        self.assertEqual(result["restored"], 1)
        self.assertEqual(client.messages["spam"]["parentFolderId"], "INBOX")
        self.run_scan(client, mode="scored", apply=True)
        self.assertEqual(len(client.moves), 2)  # one move, one restore; no re-junk
        self.assertEqual(sf.restore(client, keys=[key])["selected"], 0)

    def test_restore_skips_message_no_longer_in_junk(self):
        client = FakeClient(self.inbox())
        self.run_scan(client, mode="scored", apply=True)
        client.messages["spam"]["parentFolderId"] = "ELSEWHERE"
        self.assertEqual(sf.restore(client, since="2000-01-01")["not_in_junk"], 1)

    def test_reread_prevents_move_if_already_moved(self):
        client = FakeClient(self.inbox())
        real = client.request

        def moved_meanwhile(method, path, body=None):
            if method == "GET" and path.startswith("/me/messages/spam"):
                client.messages["spam"]["parentFolderId"] = "ARCHIVE"
            return real(method, path, body)
        client.request = moved_meanwhile
        self.assertEqual(self.run_scan(client, mode="scored", apply=True)["moved"], 0)

    def test_move_cap(self):
        messages = [msg("s%d" % i, "zqxkwv%d@example.com" % (10 + i), "Amy Lee") for i in range(5)]
        client = FakeClient(messages)
        result = self.run_scan(client, mode="scored", apply=True, max_moves=2)
        self.assertEqual(result["moved"], 2)
        self.assertEqual(result["move_cap_reached"], 1)

    def test_scheduled_off_by_default_and_live_requires_since(self):
        self.assertEqual(sf.scheduled(FakeClient(self.inbox())), {"mode": "off"})
        (sf.STATE / "config.json").write_text(json.dumps({"mode": "strict", "sources": ["local"]}))
        with self.assertRaises(GuardError):
            sf.scheduled(FakeClient(self.inbox()))

    def test_scheduled_dry_run_never_moves(self):
        (sf.STATE / "config.json").write_text(json.dumps({"mode": "dry-run", "sources": ["local"]}))
        client = FakeClient(self.inbox())
        self.assertFalse(sf.scheduled(client)["applied"])
        self.assertEqual(client.moves, [])

    def test_freemail_stranger_moved_even_with_matching_name(self):
        client = FakeClient([msg("plain", "sample.person84@example.com", "Sample Person"),
                             msg("biz", "orders@shop.example", "Shop")])
        result = self.run_scan(client, mode="scored", apply=True)
        self.assertEqual(client.moves, [("plain", "JUNK")])
        self.assertEqual(result["reasons"]["SCORED_FREEMAIL_UNKNOWN"], 1)

    def test_rescue_teaches_safe_for_next_message(self):
        client = FakeClient([msg("spam", *SPAM)])
        self.run_scan(client, mode="scored", apply=True)
        client.messages["spam"]["parentFolderId"] = "INBOX"        # the account owner moves it back
        client.add(msg("next", *SPAM))
        result = self.run_scan(client, mode="scored", apply=True)
        self.assertEqual(result["learned_safe"], 1)
        self.assertEqual(len(client.moves), 1)
        self.assertEqual(client.messages["next"]["parentFolderId"], "INBOX")

    def test_user_junk_teaches_block_even_for_contact(self):
        client = FakeClient([msg("a", FRIEND, "Sample Person")])
        self.run_scan(client, mode="scored", apply=True)            # seen in Inbox, kept
        client.messages["a"]["parentFolderId"] = "JUNK"             # the account owner moves it to Junk
        client.add(msg("b", FRIEND, "Sample Person"))
        result = self.run_scan(client, mode="scored", apply=True)
        self.assertEqual(result["learned_block"], 1)
        self.assertEqual(client.moves, [("b", "JUNK")])
        self.assertIn("USER_BLOCKED_SENDER", result["reasons"])

    def test_latest_user_action_wins(self):
        client = FakeClient([msg("a", FRIEND, "Sample Person")])
        self.run_scan(client, mode="scored", apply=True)
        client.messages["a"]["parentFolderId"] = "JUNK"             # junk by mistake
        self.run_scan(client, mode="scored", apply=True)
        client.messages["a"]["parentFolderId"] = "INBOX"            # and back again
        client.add(msg("b", FRIEND, "Sample Person"))
        result = self.run_scan(client, mode="scored", apply=True)
        self.assertEqual(result["learned_safe"], 1)
        self.assertEqual(client.moves, [])
        self.assertEqual(self.run_scan(client, mode="scored", apply=True)["learned_safe"], 0)  # learned once

    def test_outlook_delivered_junk_and_own_moves_teach_nothing(self):
        client = FakeClient([msg("spam", *SPAM), msg("outlook", "stranger@example.org", "Example Recipient", folder="JUNK")])
        self.run_scan(client, mode="scored", apply=True)
        result = self.run_scan(client, mode="scored", apply=True)
        self.assertEqual((result["learned_safe"], result["learned_block"]), (0, 0))

    def test_rescuing_outlook_junk_teaches_safe(self):
        client = FakeClient([msg("outlook", *SPAM, folder="JUNK")])
        self.run_scan(client, mode="scored", apply=True)
        client.messages["outlook"]["parentFolderId"] = "INBOX"
        client.add(msg("next", *SPAM))
        result = self.run_scan(client, mode="scored", apply=True)
        self.assertEqual(result["learned_safe"], 1)
        self.assertEqual(client.moves, [])

    def test_rescue_into_other_folder_teaches_safe(self):
        client = FakeClient([msg("spam", *SPAM)])
        self.run_scan(client, mode="scored", apply=True)
        self.run_scan(client, mode="scored", apply=True)            # seen in Junk
        client.messages["spam"]["parentFolderId"] = "PERSONAL"      # dragged straight to a folder
        client.add(msg("next", *SPAM))
        result = self.run_scan(client, mode="scored", apply=True)
        self.assertEqual(result["learned_safe"], 1)
        self.assertEqual(client.messages["next"]["parentFolderId"], "INBOX")

    def test_deleting_from_junk_or_purge_teaches_nothing(self):
        client = FakeClient([msg("spam", *SPAM), msg("purged", "zqxkwv55@example.com", "Amy Lee")])
        self.run_scan(client, mode="scored", apply=True)
        self.run_scan(client, mode="scored", apply=True)
        client.messages["spam"]["parentFolderId"] = "DELETED"
        del client.messages["purged"]
        result = self.run_scan(client, mode="scored", apply=True)
        self.assertEqual(result["learned_safe"], 0)

    def test_outlook_emptying_junk_teaches_nothing(self):
        client = FakeClient([msg("spam", *SPAM)])
        self.run_scan(client, mode="scored", apply=True)
        self.run_scan(client, mode="scored", apply=True)            # seen in Junk
        client.messages["spam"]["parentFolderId"] = "RECOVERABLE"   # Junk emptied
        client.add(msg("next", *SPAM))
        result = self.run_scan(client, mode="scored", apply=True)
        self.assertEqual(result["learned_safe"], 0)
        self.assertEqual(client.messages["next"]["parentFolderId"], "JUNK")

    def test_restore_teaches_safe(self):
        client = FakeClient([msg("spam", *SPAM)])
        self.run_scan(client, mode="scored", apply=True)
        sf.restore(client, since="2000-01-01")
        client.add(msg("next", *SPAM))
        self.run_scan(client, mode="scored", apply=True)
        self.assertEqual(client.messages["next"]["parentFolderId"], "INBOX")

    def test_allow_rejects_freemail_domain(self):
        with self.assertRaises(GuardError):
            sf.add_allow("@example.com")
        sf.add_allow("example.owner@example.org")
        self.assertIn("example.owner@example.org", (sf.STATE / "allowlist.txt").read_text())


if __name__ == "__main__":
    unittest.main()


class OutlookSafeSendersTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.patch = patch.object(sf, "STATE", Path(self.tmp.name))
        self.patch.start()

    def tearDown(self):
        self.patch.stop()
        self.tmp.cleanup()

    def snapshot(self, entries, **extra):
        from spam_guard import outlook_safe_senders as oss
        payload = {"Body": {"Options": {"TrustedSendersAndDomains": entries, "ContactsTrusted": True,
                                        "BlockedSendersAndDomains": ["blocked@example.net"]}}}
        data = oss.parse(payload)
        data.update(extra)
        (sf.STATE / oss.SNAPSHOT).write_text(json.dumps(data))
        return data

    def test_parse_hashes_and_refuses_freemail_domains(self):
        data = self.snapshot(["Sample.Person84@example.com", "@shop.example", "school.example", "example.com", "@example.com", 7])
        self.assertEqual((len(data["addresses"]), len(data["domains"]), data["ignored"]), (1, 2, 3))
        self.assertTrue(data["contacts_trusted"])
        self.assertNotIn("shop.example", json.dumps(data))   # hashes only

    def test_parse_requires_list(self):
        from spam_guard import outlook_safe_senders as oss
        with self.assertRaises(oss.SyncError):
            oss.parse({"Body": {}})

    def test_safe_senders_and_domains_protect(self):
        self.snapshot(["sample.person84@example.com", "@shop.example"])
        safe = sf.build_safe_list(None, ("outlook_safe",), Counter())
        self.assertEqual(safe.check("sample.person84@example.com"), "SAFE_OUTLOOK_SAFE_SENDER")
        self.assertEqual(safe.check("news@mail.shop.example"), "SAFE_OUTLOOK_SAFE_DOMAIN")
        self.assertIsNone(safe.check("x@badshop.example"))
        self.assertIsNone(safe.check("x@shop.example.evil.example"))

    def test_missing_snapshot_fails_closed_stale_still_protects(self):
        with self.assertRaises(GuardError):
            sf.build_safe_list(None, ("outlook_safe",), Counter())
        old = (datetime.now(timezone.utc) - timedelta(days=30)).isoformat()
        self.snapshot(["sample.person84@example.com"], synced_at=old)
        summary = Counter()
        safe = sf.build_safe_list(None, ("outlook_safe",), summary)
        self.assertEqual(summary["outlook_safe_stale"], 1)
        self.assertEqual(safe.check("sample.person84@example.com"), "SAFE_OUTLOOK_SAFE_SENDER")

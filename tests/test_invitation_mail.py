import test_support
import unittest
from spam_guard.invitation_mail import eligible

class MailTests(unittest.TestCase):
    def message(self, address, kind='#microsoft.graph.eventMessageRequest'):
        return {'@odata.type':kind,'from':{'emailAddress':{'address':address}}}
    def test_groups_invite(self): self.assertTrue(eligible(self.message('synthetic@GROUPS.OUTLOOK.COM')))
    def test_ordinary_mail_preserved(self): self.assertFalse(eligible(self.message('synthetic@groups.outlook.com','#microsoft.graph.message')))
    def test_unknown_invite_preserved(self): self.assertFalse(eligible(self.message('synthetic@example.com')))
    def test_suffix_attack(self): self.assertFalse(eligible(self.message('synthetic@groups.outlook.com.evil.example')))


class MoveTests(unittest.TestCase):
    def test_missing_account_configuration_fails_closed(self):
        from unittest.mock import patch
        from spam_guard import invitation_mail as im
        with patch.object(im.config, 'ACCOUNT', ''):
            with self.assertRaisesRegex(RuntimeError, 'ACCOUNT_NOT_CONFIGURED'):
                im.scan(apply=True)

    def test_matches_move_to_junk_not_deleted_items(self):
        from unittest.mock import patch
        from spam_guard import invitation_mail as im
        invite={'id':'m1','@odata.type':'#microsoft.graph.eventMessageRequest','from':{'emailAddress':{'address':'synthetic@groups.outlook.com'}}}
        calls=[]
        def fake(method,path,body=None,query=None):
            calls.append((method,path,body))
            if path=='/me': return {'mail':'owner@example.com'}
            if path.startswith('/me/mailFolders/inbox/messages'): return {'value':[invite]}
            if method=='POST': return {'id':'moved'}
            return invite
        with patch.object(im.graph,'graph_request',side_effect=fake):
            self.assertEqual(im.scan(apply=True)['moved'],1)
        self.assertIn(('POST','/me/messages/m1/move',{'destinationId':'junkemail'}),calls)

"""Deterministic sender-identity scoring from address and display name only.

Both inputs are attacker-controlled. Nothing here reads subjects, bodies, links
or attachments, and nothing is sent to an agent or LLM. Every decision carries
fixed reason codes so a human can see why a sender was (or was not) flagged.

No single signal is treated as proof: a random-looking address, trailing digits
or a name mismatch each score below the threshold on their own. Legitimate
aliases, businesses, automated senders and transliterated names are
accommodated by (a) weighting randomness heavily only on free-mail domains,
(b) suppressing randomness when the address matches the claimed name, and
(c) not assessing name mismatch for organisation-like or non-Latin names.
"""
from __future__ import annotations

import re
import unicodedata

THRESHOLD = 4

FREEMAIL = frozenset("""
gmail.com googlemail.com outlook.com hotmail.com live.com msn.com yahoo.com
ymail.com icloud.com me.com mac.com aol.com proton.me protonmail.com gmx.com
gmx.net mail.com yandex.com yandex.ru mail.ru qq.com 163.com 126.com zoho.com
hotmail.co.uk yahoo.co.uk outlook.com.au hotmail.com.au yahoo.com.au live.com.au
bigpond.com bigpond.net.au optusnet.com.au
""".split())

# Words that make a display name organisation-like rather than a person's name.
ORG_WORDS = frozenset("""
team support info news newsletter service services notification notifications
alert alerts billing accounts account admin store shop bank official customer
care help noreply no-reply reply update updates security office group pty ltd
inc llc co company club community digest mail delivery order orders
""".split())

# Common English/romanised-name letter pairs. A local part made mostly of pairs
# outside this set reads as generated rather than typed.
COMMON_BIGRAMS = frozenset("""
th he in er an re on at en nd ti es or te of ed is it al ar st to nt ng se ha
as ou io le ve co me de hi ri ro ic ne ea ra ce li ch ll be ma si om ur ca el
ta la ns di fo ho pe ec pr no ct us ac ot il tr ly nc et ut ss so rs un lo wa
ge ie wh ee wi em ad ol rt po we na ul ni ts mo ow pa im mi ai sh ir su id os
iv ia am fi ci vi pl ig tu ev ld ry mp fe bl ab gh ty op wo sa ay ex ke fr oo
av ag if ap gr od bo sp rd do uc bu ei ov by rm ep tt oc fa ef cu rn sc gi da
yo cr cl du ga qu ue ff ba ey ls va um pp ua up lu go ht ru ug ds lt pi rc rr
eg au ck ew mu br bi pt ak pu ui rg ib tl ny ki rk ys ob mm fu ph og ms ye ud
mb ip ub oi rl gu dr hr cc tw ft wn nu af hu nn eo vo rv nf xp gn sm fl iz ok
nl my gl aw sy oa ks yt ja jo je ju ka ko ku da ya yu za zi zo ze ky ry kr nk
""".split())

VOWELS = set("aeiouy")


def _ascii(text: str) -> str:
    """Fold accents to ASCII; German-style digraphs are normalised separately."""
    folded = unicodedata.normalize("NFKD", text)
    return "".join(c for c in folded if c.isascii()).lower()


def _digraphs(text: str) -> str:
    return text.replace("ae", "a").replace("oe", "o").replace("ue", "u")


def split_address(address) -> tuple[str, str] | None:
    if not isinstance(address, str):
        return None
    match = re.fullmatch(r"([^\s@<>\"]{1,64})@([a-z0-9.-]{1,253}\.[a-z]{2,63})", address.strip().lower())
    return (match.group(1), match.group(2)) if match else None


def name_tokens(display) -> list[str] | None:
    """Return person-name tokens, or None when the name cannot be assessed."""
    if not isinstance(display, str) or not display.strip():
        return None
    raw = display.strip()
    letters = [c for c in raw if c.isalpha()]
    if not letters:
        return None
    # Non-Latin scripts (CJK, Cyrillic, Arabic...) cannot be compared with an
    # ASCII local part without real transliteration; do not guess.
    if sum(1 for c in letters if not _ascii(c)) > len(letters) / 3:
        return None
    if any(c.isdigit() or c in "@/:" for c in raw):
        return None
    tokens = [t for t in re.split(r"[^a-z]+", _ascii(raw)) if t]
    if not 1 <= len(tokens) <= 4 or any(t in ORG_WORDS for t in tokens):
        return None
    return tokens


def name_matches(local: str, tokens: list[str]) -> bool:
    letters = _digraphs(re.sub(r"[^a-z]", "", local))
    norm = [_digraphs(t) for t in tokens]
    if any(len(t) >= 3 and t in letters for t in norm):
        return True
    if len(letters) >= 3 and any(t.startswith(letters) for t in norm):
        return True
    # Initial + surname / name + initial: "tcooper", "toric".
    if len(norm) >= 2 and len(norm[-1]) >= 3:
        if (norm[0][0] + norm[-1]) in letters or (norm[0] + norm[-1][0]) in letters:
            return True
    # Several 3-letter prefixes: "torcoo".
    prefixes = [t[:3] for t in norm if len(t) >= 3]
    return len(prefixes) >= 2 and all(p in letters for p in prefixes)


def looks_random(local: str) -> bool:
    letters = re.sub(r"[^a-z]", "", local)
    if len(letters) < 6:
        return False
    pairs = [letters[i:i + 2] for i in range(len(letters) - 1)]
    rare = sum(1 for p in pairs if p not in COMMON_BIGRAMS) / len(pairs)
    vowels = sum(1 for c in letters if c in VOWELS) / len(letters)
    return rare >= 0.4 or vowels < 0.15


def alnum_scramble(local: str) -> bool:
    """Digits interleaved inside letters, e.g. a8f3k2q."""
    return len(re.findall(r"[a-z][0-9]+(?=[a-z])", local)) >= 2


def embedded_address_spoof(display, address: str) -> bool:
    if not isinstance(display, str):
        return False
    found = re.findall(r"[^\s@<>\"'()]+@[^\s@<>\"'()]+\.[a-z]{2,}", display.lower())
    return any(f.strip(".,;") != address for f in found)


def score_sender(address, display) -> dict:
    """Score one sender. Returns {'score', 'signals', 'notes', 'suspicious'}.

    signals: codes that add to the score; notes: protective/informational codes.
    """
    parts = split_address(address)
    if parts is None:
        return {"score": 0, "signals": [], "notes": ["MALFORMED_ADDRESS"], "suspicious": False}
    local, domain = parts
    address = local + "@" + domain
    freemail = domain in FREEMAIL
    signals: dict[str, int] = {}
    notes: list[str] = ["FREEMAIL_DOMAIN"] if freemail else []

    tokens = name_tokens(display)
    matched = tokens is not None and name_matches(local, tokens)
    if matched:
        notes.append("NAME_MATCHES_ADDRESS")
    elif tokens is None and isinstance(display, str) and display.strip():
        notes.append("NAME_NOT_ASSESSABLE")

    if looks_random(local) and not matched:
        signals["RANDOM_LOCAL_PART"] = 2 if freemail else 1
    if alnum_scramble(local) and not matched:
        signals["ALNUM_SCRAMBLE"] = 2 if freemail else 1
    if re.search(r"[a-z][0-9]{2,}$", local):
        signals["TRAILING_DIGITS"] = 1
    if tokens is not None and not matched and freemail:
        signals["NAME_MISMATCH"] = 2
    if embedded_address_spoof(display, address):
        signals["DISPLAY_NAME_ADDRESS_SPOOF"] = 3

    score = sum(signals.values())
    # Require two independent signals as well as the threshold.
    suspicious = score >= THRESHOLD and len(signals) >= 2
    return {"score": score, "signals": sorted(signals), "notes": notes, "suspicious": suspicious}

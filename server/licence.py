"""TQ Excel Connector - licence check.

A licence is a small JSON file signed by Techno Q:

    {"customer": "World of Hospitality", "account": "1234567", "users": 5,
     "expires": "2027-12-31", "issued": "2026-11-17", "signature": "<base64>"}

The signature covers the other fields, so a licence cannot be edited or moved to another NetSuite
account. Techno Q makes licences with scripts/make_licence.py and the private key it keeps; the
connector only holds the public key (licence_key.py). Sample data never needs a licence.

"users" is enforced by seats (updates.Seats and tq_seats.json in NetSuite): per NetSuite user with sign-in, per PC with tokens.
"""

import base64
import datetime
import json
import os

import tqcrypto

FIELDS = ("customer", "account", "users", "expires", "issued")


def _public_key():
    try:
        from licence_key import PUBLIC_KEY
    except ImportError:
        return None
    return PUBLIC_KEY


def payload(licence):
    """The bytes the signature covers."""
    return json.dumps({k: licence.get(k) for k in FIELDS}, sort_keys=True, separators=(",", ":")).encode()


def normal_account(account):
    return str(account or "").strip().upper().replace("-", "_")


def parse(text):
    """Licence text (as pasted or read from the file) -> dict. Raises ValueError."""
    try:
        licence = json.loads(text)
    except ValueError:
        raise ValueError("The licence is not in the right form. Paste the whole text you received from Techno Q.")
    if not isinstance(licence, dict) or not licence.get("signature"):
        raise ValueError("The licence is not in the right form. Paste the whole text you received from Techno Q.")
    return licence


def read(path):
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as handle:
        return parse(handle.read())


def check(licence, account, today=None):
    """Returns {"ok", "message", and the licence's fields}. account: the NetSuite account in use."""
    today = today or datetime.date.today()
    if not licence:
        return {"ok": False, "message": "No licence yet. Paste the licence from Techno Q on the setup page."}
    info = {k: licence.get(k) for k in FIELDS}
    public = _public_key()
    try:
        signature = base64.b64decode(str(licence.get("signature")), validate=True)
    except ValueError:
        signature = b""
    if not public or not tqcrypto.verify(public, payload(licence), signature):
        return dict(info, ok=False, message="The licence is not valid (its signature does not match). "
                                            "Ask Techno Q for a new one.")
    if normal_account(licence.get("account")) != normal_account(account):
        return dict(info, ok=False, message="The licence is for NetSuite account %s, not %s."
                                            % (licence.get("account"), account))
    try:
        expires = datetime.date.fromisoformat(str(licence.get("expires")))
    except ValueError:
        return dict(info, ok=False, message="The licence has no valid expiry date.")
    if expires < today:
        return dict(info, ok=False, message="The licence ended on %s. Ask Techno Q to renew it." % expires)
    return dict(info, ok=True, message="Licensed to %s for %s users until %s."
                                       % (licence.get("customer"), licence.get("users"), expires))


def make(key, customer, account, users, expires, today=None):
    """A signed licence (used by Techno Q's scripts/make_licence.py)."""
    licence = {"customer": customer, "account": normal_account(account), "users": int(users),
               "expires": str(expires), "issued": str(today or datetime.date.today())}
    licence["signature"] = base64.b64encode(tqcrypto.sign(key, payload(licence))).decode()
    return licence

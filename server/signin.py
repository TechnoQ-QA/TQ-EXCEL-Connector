"""TQ Excel Connector - "Sign in with NetSuite" (OAuth 2.0, authorization code grant with PKCE).

Each person signs in on NetSuite's own page in their browser (with their two-factor check, if they have one) and
picks a role there. The connector never sees the password. NetSuite sends the browser back to
https://localhost:3000/oauth/callback with a one-time code, which is exchanged for:
  - an access token (one hour), sent as "Authorization: Bearer ..." with every RESTlet call, and
  - a refresh token, used to get the next access token without asking again; its lifetime and rotation are set on
    the integration record (up to 30 days), after which the person signs in again.

The integration record is a public client (no secret on the PC; PKCE proves the code was asked for here). Its
client ID goes in config.json as "clientId". The tokens are kept in data/session.json, encrypted for this Windows
user with DPAPI, so another user of the PC (or a copy of the file) cannot use them.

Token-based authentication (the four keys in config.json) still works for accounts not moved yet; NetSuite is
retiring it by 2027.
"""

import base64
import hashlib
import json
import os
import secrets
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request

REDIRECT_URI = "https://localhost:3000/oauth/callback"
SCOPE = "restlets"
PENDING_SECONDS = 600       # a sign-in started more than ten minutes ago is not accepted
REFRESH_EARLY = 120         # renew the access token two minutes before it ends


class SignInNeeded(Exception):
    """No valid session: the person has to sign in (again)."""


# ------------------------------------------------------------------ DPAPI --

def _dpapi(data, protect):
    """Windows CryptProtectData / CryptUnprotectData for the signed-in Windows user."""
    import ctypes
    from ctypes import wintypes

    class Blob(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    crypt32, kernel32 = ctypes.windll.crypt32, ctypes.windll.kernel32
    source = ctypes.create_string_buffer(data, len(data))
    blob_in, blob_out = Blob(len(data), ctypes.cast(source, ctypes.POINTER(ctypes.c_char))), Blob()
    call = crypt32.CryptProtectData if protect else crypt32.CryptUnprotectData
    if not call(ctypes.byref(blob_in), None, None, None, None, 0x1, ctypes.byref(blob_out)):   # 0x1: no UI
        raise OSError("Windows could not %s the NetSuite session" % ("protect" if protect else "read"))
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        kernel32.LocalFree(blob_out.pbData)


def protect(text):
    raw = text.encode()
    if sys.platform == "win32":
        return "dpapi:" + base64.b64encode(_dpapi(raw, True)).decode()
    return "plain:" + base64.b64encode(raw).decode()   # tests on other systems only


def unprotect(text):
    kind, _, data = str(text).partition(":")
    raw = base64.b64decode(data)
    return (_dpapi(raw, False) if kind == "dpapi" else raw).decode()


# ---------------------------------------------------------------- session --

def host(account):
    return str(account).strip().lower().replace("_", "-")


def authorize_url(cfg):
    return "https://%s.app.netsuite.com/app/login/oauth2/authorize.nl" % host(cfg["accountId"])


def token_url(cfg):
    return "https://%s.suitetalk.api.netsuite.com/services/rest/auth/oauth2/v1/token" % host(cfg["accountId"])


def revoke_url(cfg):
    return "https://%s.suitetalk.api.netsuite.com/services/rest/auth/oauth2/v1/revoke" % host(cfg["accountId"])


def _post_form(url, fields):
    request = urllib.request.Request(url, data=urllib.parse.urlencode(fields).encode(), method="POST",
                                     headers={"Content-Type": "application/x-www-form-urlencoded",
                                              "Accept": "application/json"})
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            return response.status, response.read().decode()
    except urllib.error.HTTPError as err:
        return err.code, err.read().decode(errors="replace")


class Session:
    """The signed-in person's tokens for one NetSuite account, kept in data/session.json."""

    def __init__(self, cfg, data_dir, post=None):
        self.cfg, self.path = cfg, os.path.join(data_dir, "session.json")
        self.post = post or _post_form          # tests replace the network call
        self.pending = {}                       # state -> (code verifier, when)
        self.lock = threading.Lock()
        self.access, self.expires = None, 0.0

    # ---- reading and writing the file
    def _load(self):
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                saved = json.load(f)
        except (OSError, ValueError):
            return {}
        if not isinstance(saved, dict) or saved.get("account") != self.cfg["accountId"]:
            return {}   # another account's session is not used
        return saved

    def _save(self, saved):
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(saved, f)

    def info(self):
        """What the panel shows: signed in or not, and as whom."""
        saved = self._load()
        if not saved.get("refresh"):
            return {"mode": "oauth2", "signedIn": False}
        return {"mode": "oauth2", "signedIn": True, "user": saved.get("user", ""), "role": saved.get("role", ""),
                "since": saved.get("since", "")}

    # ---- signing in
    def start(self):
        """The NetSuite address to open in the browser; remembers the PKCE verifier for the callback."""
        verifier = secrets.token_urlsafe(64)
        challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).decode().rstrip("=")
        state = secrets.token_urlsafe(24)
        now = time.time()
        with self.lock:
            self.pending = {s: v for s, v in self.pending.items() if now - v[1] < PENDING_SECONDS}
            self.pending[state] = (verifier, now)
        return authorize_url(self.cfg) + "?" + urllib.parse.urlencode({
            "response_type": "code", "client_id": self.cfg["clientId"], "redirect_uri": REDIRECT_URI,
            "scope": SCOPE, "state": state, "code_challenge": challenge, "code_challenge_method": "S256"})

    def finish(self, code, state):
        """The callback: checks state, exchanges the code, saves the tokens. Raises ValueError with a reason."""
        with self.lock:
            found = self.pending.pop(state or "", None)
        if not found or time.time() - found[1] > PENDING_SECONDS:
            raise ValueError("This sign-in was not started from the TQ Connector on this PC, or it took too long. "
                             "Start it again from the panel.")
        status, text = self.post(token_url(self.cfg), {
            "grant_type": "authorization_code", "code": code, "redirect_uri": REDIRECT_URI,
            "client_id": self.cfg["clientId"], "code_verifier": found[0]})
        self._store(status, text, {})

    def _store(self, status, text, saved):
        try:
            answer = json.loads(text)
        except ValueError:
            answer = {}
        if status >= 400 or not answer.get("access_token"):
            raise ValueError("NetSuite did not accept the sign-in: %s" % (answer.get("error_description")
                                                                          or answer.get("error") or text[:200]))
        self.access = answer["access_token"]
        self.expires = time.time() + int(answer.get("expires_in") or 3600)
        refresh = answer.get("refresh_token") or (unprotect(saved["refresh"]) if saved.get("refresh") else "")
        saved = dict(saved, account=self.cfg["accountId"], refresh=protect(refresh) if refresh else "",
                     since=saved.get("since") or time.strftime("%Y-%m-%d %H:%M"))
        self._save(saved)

    def remember_user(self, user, role):
        saved = self._load()
        if saved:
            saved.update(user=str(user or "")[:80], role=str(role or "")[:80])
            self._save(saved)

    # ---- using it
    def token(self, renew=False):
        """A valid access token, renewed with the refresh token when needed. Raises SignInNeeded."""
        with self.lock:
            if self.access and not renew and time.time() < self.expires - REFRESH_EARLY:
                return self.access
            saved = self._load()
            if not saved.get("refresh"):
                raise SignInNeeded("Sign in with NetSuite first (TQ Connector panel, Sign in).")
            try:
                refresh = unprotect(saved["refresh"])
            except (OSError, ValueError):
                raise SignInNeeded("The saved NetSuite session cannot be read on this PC. Sign in again.")
            status, text = self.post(token_url(self.cfg), {"grant_type": "refresh_token", "refresh_token": refresh,
                                                           "client_id": self.cfg["clientId"]})
            try:
                self._store(status, text, saved)
            except ValueError:
                self.access = None
                self._save({})   # the refresh token is spent or revoked: sign in again
                raise SignInNeeded("Your NetSuite sign-in has ended. Sign in again (TQ Connector panel, Sign in).")
            return self.access

    def sign_out(self):
        saved = self._load()
        if saved.get("refresh"):
            try:
                self.post(revoke_url(self.cfg), {"token": unprotect(saved["refresh"]), "client_id": self.cfg["clientId"]})
            except (OSError, ValueError, urllib.error.URLError):
                pass   # signed out here either way
        self.access, self.expires = None, 0.0
        self._save({})

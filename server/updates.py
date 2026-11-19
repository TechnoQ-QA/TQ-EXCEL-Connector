"""TQ Excel Connector - seats per PC and automatic updates.

Seats. The licence allows a number of users. Each PC's connector has its own id (machine.json, made once) and
registers it with the NetSuite script, which keeps the list (tq_seats.json in the File Cabinet). A PC that is
not on the list is refused once the list is full; a PC not seen for 30 days frees its seat.

Updates. Techno Q publishes an update manifest at a web address (config "updateUrl"):
    {"version": "1.1.0", "url": "https://.../TQConnectorSetup.exe", "sha256": "...", "notes": "...",
     "signature": "<base64>"}
The signature (made with Techno Q's licence key, scripts/sign_update.py) covers the other fields, so a manifest
cannot point at another file. The downloaded setup file must match sha256 before it is run.
"""

import base64
import hashlib
import json
import os
import socket
import subprocess
import sys
import tempfile
import threading
import time
import urllib.request
import uuid

import licence as licence_module
import tqcrypto

FIELDS = ("version", "url", "sha256", "notes")
SEAT_RECHECK = 6 * 3600       # seconds between seat checks while allowed
SEAT_RETRY = 300              # ... and while refused
MAX_MANIFEST = 64 * 1024
MAX_SETUP = 300 * 1024 * 1024


# ------------------------------------------------------------------ seats --

def machine(data_dir):
    """This PC's id and name, made once and kept in machine.json."""
    path = os.path.join(data_dir, "machine.json")
    try:
        with open(path, "r", encoding="utf-8") as f:
            found = json.load(f)
        if isinstance(found, dict) and found.get("id"):
            return found
    except (OSError, ValueError):
        pass
    found = {"id": uuid.uuid4().hex, "label": socket.gethostname()[:60]}
    try:
        os.makedirs(data_dir, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(found, f)
    except OSError:
        pass   # a read-only folder: the id lasts until the connector restarts
    return found


class Seats:
    """Wraps the NetSuite call: before data is read, this PC must hold one of the licence's seats."""

    def __init__(self, backend, pc, users, by_user=False):
        self.backend, self.pc, self.users, self.by_user = backend, pc, int(users or 1), by_user
        self.checked, self.state = 0.0, {"counted": False, "allowed": True, "message": "Not checked yet."}
        self.lock = threading.Lock()

    def _check(self):
        answer = self.backend({"action": "seat", "machine": self.pc["id"], "label": self.pc.get("label", ""),
                               "users": self.users, "byUser": self.by_user})
        if not answer.get("ok"):
            error = str(answer.get("error") or "")
            if "unknown action" in error.lower():   # NetSuite scripts older than version 4: nothing to count with
                return {"counted": False, "allowed": True, "message": "Seats are not counted (NetSuite scripts "
                                                                      "before version 4)."}
            return dict(self.state, message="Seat check failed: %s" % error)   # keep the last answer
        used, users = answer.get("used"), answer.get("users")
        if "allowed" not in answer:   # an answer that does not know about seats
            return {"counted": False, "allowed": True, "message": "Seats are not counted."}
        if answer.get("allowed"):
            return {"counted": True, "allowed": True, "used": used, "users": users,
                    "message": "This PC holds a seat: %s of %s in use." % (used, users)}
        names = ", ".join("%s (%s, last %s)" % (m.get("label"), m.get("user"), m.get("last"))
                          for m in answer.get("machines") or [])
        return {"counted": True, "allowed": False, "used": used, "users": users,
                "message": "All %s seats of the licence are in use: %s. A PC not used for 30 days frees its seat; "
                           "or ask Techno Q for more users." % (users, names)}

    def __call__(self, payload):
        # the connection test and the seat list itself pass, so a full licence can still be seen and a seat freed
        if payload.get("action") not in ("ping", "seat", "seatList", "seatRelease"):
            with self.lock:
                wait = SEAT_RECHECK if self.state.get("allowed") else SEAT_RETRY
                if time.time() - self.checked > wait:
                    self.checked = time.time()
                    self.state = self._check()
            if not self.state.get("allowed"):
                return {"ok": False, "error": "Licence: " + self.state["message"]}
        return self.backend(payload)


# ---------------------------------------------------------------- updates --

def manifest_payload(manifest):
    return json.dumps({k: manifest.get(k) for k in FIELDS}, sort_keys=True, separators=(",", ":")).encode()


def sign_manifest(key, version, url, sha256, notes=""):
    manifest = {"version": version, "url": url, "sha256": sha256.lower(), "notes": notes}
    manifest["signature"] = base64.b64encode(tqcrypto.sign(key, manifest_payload(manifest))).decode()
    return manifest


def _version(text):
    try:
        return tuple(int(p) for p in str(text).strip().split("."))
    except ValueError:
        return (0,)


def _fetch(url, limit):
    if not str(url).lower().startswith("https://"):
        raise ValueError("Updates are only fetched over https.")
    with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "TQConnector"}), timeout=60) as r:
        data = r.read(limit + 1)
    if len(data) > limit:
        raise ValueError("The download is larger than expected.")
    return data


def verify(manifest):
    """Raises ValueError unless the manifest is signed with Techno Q's key."""
    public = licence_module._public_key()
    try:
        signature = base64.b64decode(str(manifest.get("signature")), validate=True)
    except ValueError:
        signature = b""
    if not public or not tqcrypto.verify(public, manifest_payload(manifest), signature):
        raise ValueError("The update is not signed by Techno Q, so it is not used.")
    if not str(manifest.get("url", "")).lower().startswith("https://") or len(str(manifest.get("sha256", ""))) != 64:
        raise ValueError("The update description is incomplete.")


def check(update_url, current):
    """{"available", "version", "notes", "current"}; raises ValueError for a bad or unreachable manifest."""
    if not update_url:
        return {"available": False, "current": current, "message": "No update address is set up."}
    try:
        manifest = json.loads(_fetch(update_url, MAX_MANIFEST).decode())
    except ValueError:
        raise
    except Exception as err:   # noqa: BLE001 - network errors of any kind
        raise ValueError("Could not reach the update address: %s" % err)
    if not isinstance(manifest, dict):
        raise ValueError("The update description is not in the right form.")
    verify(manifest)
    newer = _version(manifest["version"]) > _version(current)
    return {"available": newer, "current": current, "version": manifest["version"], "notes": manifest.get("notes", ""),
            "manifest": manifest,
            "message": ("Version %s is available." % manifest["version"]) if newer else "This is the latest version."}


class Watcher:
    """Checks the update address a minute after the connector starts, then once a week. The panel shows a notice
    when a newer, signed version is there. config "autoUpdateCheck": false turns the weekly check off (an
    administrator who rolls out updates centrally); Check for an update on the setup page still works."""

    WEEK = 7 * 24 * 3600

    def __init__(self, url_of, current, enabled_of):
        self.url_of, self.current, self.enabled_of = url_of, current, enabled_of
        self.state = {"available": False}
        self.stop = threading.Event()

    def check(self):
        url = self.url_of()
        if not url or not self.enabled_of():
            self.state = {"available": False}
            return self.state
        try:
            found = check(url, self.current)
            found.pop("manifest", None)
            self.state = dict(found, checked=time.strftime("%Y-%m-%d %H:%M"))
        except ValueError as err:
            self.state = {"available": False, "error": str(err), "checked": time.strftime("%Y-%m-%d %H:%M")}
        return self.state

    def run(self):
        if not self.stop.wait(60):
            while True:
                self.check()
                if self.stop.wait(self.WEEK):
                    break

    def start(self):
        threading.Thread(target=self.run, daemon=True).start()
        return self


def install(update_url, current, run=True):
    """Downloads the newer setup file, checks it against the signed sha256 and starts it. The setup stops this
    connector, replaces its files and starts it again."""
    found = check(update_url, current)
    if not found["available"]:
        return {"ok": True, "started": False, "message": found["message"]}
    manifest = found["manifest"]
    data = _fetch(manifest["url"], MAX_SETUP)
    if hashlib.sha256(data).hexdigest() != manifest["sha256"].lower():
        raise ValueError("The downloaded setup file does not match the signed update, so it is not run.")
    path = os.path.join(tempfile.gettempdir(), "TQConnectorSetup-%s.exe" % manifest["version"])
    with open(path, "wb") as f:
        f.write(data)
    if run and sys.platform == "win32":
        subprocess.Popen([path], close_fds=True,
                         creationflags=getattr(subprocess, "DETACHED_PROCESS", 0) |
                         getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0))
    return {"ok": True, "started": bool(run), "path": path,
            "message": "Installing version %s: the connector restarts by itself in a minute. Then close and "
                       "reopen the TQ Connector panel." % manifest["version"]}

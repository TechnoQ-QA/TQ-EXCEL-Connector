"""TQ Excel Connector - keeping it running and moving it between PCs.

    watch(argv)              TQConnector.exe --watch: starts the connector and starts it again if it stops, so
                             Excel's add-in (which the connector serves) never stays broken after a crash.
    cert_days_left(path)     days until this PC's localhost certificate expires.
    renew_certificate(...)   a new certificate 60 days before the old one expires, trusted for this Windows user
                             (Windows asks to confirm), then the old one removed. Files are swapped only once
                             the new one is trusted, so a "No" leaves everything working as it was.
    diagnostics(...)         a zip for Techno Q support: versions, status, settings with every secret removed,
                             the end of the log. Saved in Documents\\TQ Connector.
    Prefs                    small per-PC preferences (saved filter sets), data\\prefs.json.
    export_bundle / import_bundle
                             "Move to another PC": the connection settings (never a secret), the licence, the
                             logo, saved filter sets and the workbook's layouts in one .json file.
"""

import base64
import datetime
import json
import os
import platform
import re
import ssl
import subprocess
import sys
import threading
import time
import urllib.request
import zipfile

RENEW_DAYS = 60
SECRET_WORDS = re.compile(r"secret|token|password|apikey|aikey|key$", re.I)
BUNDLE_KIND = "tq-connector-setup"


def _log(text):
    print("%s %s" % (time.strftime("%Y-%m-%d %H:%M:%S"), text))
    try:
        sys.stdout.flush()
    except (AttributeError, OSError, ValueError):
        pass


# ------------------------------------------------------------------ keep it running --

def _answering(port):
    """True when a connector already answers on the port (then there is nothing to restart)."""
    try:
        ctx = ssl.create_default_context()
        ctx.check_hostname, ctx.verify_mode = False, ssl.CERT_NONE
        with urllib.request.urlopen("https://localhost:%d/api/status" % port, timeout=4, context=ctx) as r:
            return r.status == 200
    except Exception:   # noqa: BLE001 - any failure means "not answering"
        return False


def watch(argv, frozen, script, port=3000):
    """Runs the connector as a child process and starts it again whenever it stops. A child that stops at
    once (another copy has the port, a broken setting) is retried less and less often, up to once a minute."""
    child_args = [a for a in argv if a != "--watch"]
    cmd = [sys.executable] + ([] if frozen else [script]) + child_args
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    wait = 3
    while True:
        if _answering(port):
            time.sleep(60)      # someone else's copy is serving: check again later
            continue
        started = time.time()
        try:
            child = subprocess.Popen(cmd, creationflags=flags)
            code = child.wait()
        except OSError as err:
            code = "could not start: %s" % err
        ran = time.time() - started
        wait = 3 if ran > 60 else min(60, wait * 2)
        _log("Connector stopped (%s) after %d s; starting it again in %d s." % (code, ran, wait))
        time.sleep(wait)


# ------------------------------------------------------------- the localhost certificate --

def cert_expires(cert_path):
    """The certificate's expiry as a date, or None when it cannot be read."""
    try:
        info = ssl._ssl._test_decode_cert(cert_path)   # noqa: SLF001 - the standard library's own PEM reader
        return datetime.datetime.strptime(info["notAfter"], "%b %d %H:%M:%S %Y %Z").date()
    except Exception:   # noqa: BLE001
        try:   # made by this connector: valid 825 days from when the file was written
            return (datetime.datetime.fromtimestamp(os.path.getmtime(cert_path)) + datetime.timedelta(days=825)).date()
        except OSError:
            return None


def cert_days_left(cert_path):
    expires = cert_expires(cert_path)
    return (expires - datetime.date.today()).days if expires else None


def _certutil(args):
    system = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "certutil.exe")
    try:
        return subprocess.run([system] + args, capture_output=True, timeout=600,
                              creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)).returncode
    except (OSError, subprocess.SubprocessError):
        return -1


def renew_certificate(cert_path, key_path, make, days=RENEW_DAYS):
    """When the certificate expires within `days`: make a new one beside it, ask Windows to trust it, and swap
    the files only if that worked. Returns a short note, or "" when nothing was needed."""
    left = cert_days_left(cert_path)
    if left is None or left > days or os.name != "nt":
        return ""
    new_cert, new_key = cert_path + ".new", key_path + ".new"
    old_thumb = ""
    try:
        with open(cert_path + ".sha1", encoding="ascii") as f:
            old_thumb = f.read().strip()
    except OSError:
        pass
    make(new_cert, new_key)
    if _certutil(["-user", "-addstore", "Root", new_cert]) != 0:
        for p in (new_cert, new_key, new_cert + ".sha1"):
            try:
                os.remove(p)
            except OSError:
                pass
        return "The certificate expires in %d days; the new one was not trusted (Windows was answered No). " \
               "It will be asked again next start." % left
    os.replace(new_cert, cert_path)
    os.replace(new_key, key_path)
    if os.path.exists(new_cert + ".sha1"):
        os.replace(new_cert + ".sha1", cert_path + ".sha1")
    with open(cert_path + ".trusted", "w", encoding="ascii") as f:
        f.write("yes")
    if len(old_thumb) == 40:
        _certutil(["-user", "-delstore", "Root", old_thumb])
    return "The localhost certificate was renewed (the old one had %d days left)." % left


# ----------------------------------------------------------------- support and files --

def documents_dir():
    """Documents\\TQ Connector, where diagnostics and exported setups are saved."""
    home = os.path.expanduser("~")
    docs = os.path.join(home, "Documents")
    path = os.path.join(docs if os.path.isdir(docs) else home, "TQ Connector")
    os.makedirs(path, exist_ok=True)
    return path


def show_in_folder(path):
    """Opens Explorer with the file selected (Windows), so the person can attach or copy it."""
    if os.name != "nt":
        return
    try:
        subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
    except OSError:
        pass


def redact(value):
    """A copy of the settings with every secret replaced (keys, tokens, passwords, the AI key)."""
    if isinstance(value, dict):
        return {k: ("***" if SECRET_WORDS.search(str(k)) and value[k] not in ("", None) else redact(value[k]))
                for k in value if not str(k).startswith("_")}
    if isinstance(value, list):
        return [redact(v) for v in value]
    return value


def diagnostics(data_dir, version, status, config, cert_path, open_folder=True):
    """Writes TQ diagnostics <date>.zip and returns its path. Nothing secret goes in: settings are redacted,
    the NetSuite sign-in (session.json), the certificate's key and the licence key are never added."""
    stamp = time.strftime("%Y-%m-%d %H%M")
    path = os.path.join(documents_dir(), "TQ diagnostics %s.zip" % stamp)
    log_tail = ""
    try:
        with open(os.path.join(data_dir, "server.log"), encoding="utf-8", errors="replace") as f:
            log_tail = "".join(f.readlines()[-3000:])
    except OSError:
        log_tail = "(no log)"
    files = []
    try:
        files = sorted("%s  %d bytes" % (n, os.path.getsize(os.path.join(data_dir, n))) for n in os.listdir(data_dir)
                       if os.path.isfile(os.path.join(data_dir, n)))
    except OSError:
        pass
    about = {
        "connector": version, "made": stamp, "windows": platform.platform(), "python": sys.version.split()[0],
        "certificateExpires": str(cert_expires(cert_path) or "unknown"), "dataFiles": files,
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("about.json", json.dumps(about, indent=2))
        z.writestr("status.json", json.dumps(redact(status), indent=2, default=str))
        z.writestr("settings.json", json.dumps(redact(config or {}), indent=2, default=str))
        z.writestr("server.log", log_tail)
    if open_folder:
        show_in_folder(path)
    return path


class Prefs:
    """data\\prefs.json: this PC's saved filter sets (and room for more), written whole on each change."""

    def __init__(self, data_dir):
        self.path = os.path.join(data_dir, "prefs.json")
        self.lock = threading.Lock()

    def read(self):
        try:
            with open(self.path, encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except (OSError, ValueError):
            return {}

    def write(self, data):
        with self.lock:
            os.makedirs(os.path.dirname(self.path), exist_ok=True)
            tmp = self.path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=1)
            os.replace(tmp, self.path)

    def handle(self, payload):
        """{op: get} | {op: saveSet, name, cells} | {op: removeSet, name}."""
        data = self.read()
        sets = [s for s in data.get("filterSets") or [] if isinstance(s, dict) and s.get("name")]
        op = str(payload.get("op") or "get")
        if op == "saveSet":
            name = str(payload.get("name") or "").strip()[:60]
            cells = payload.get("cells") if isinstance(payload.get("cells"), dict) else {}
            if not name:
                return {"ok": False, "error": "Give the filter set a name."}
            clean = {k: str(v)[:400] for k, v in cells.items()
                     if k in ("subsidiary", "department", "class", "location", "filters") and str(v or "").strip()}
            if not clean:
                return {"ok": False, "error": "The report has no filters to save: every amount is shown."}
            sets = [s for s in sets if s["name"].lower() != name.lower()] + [{"name": name, "cells": clean}]
        elif op == "removeSet":
            name = str(payload.get("name") or "").strip().lower()
            sets = [s for s in sets if s["name"].lower() != name]
        elif op != "get":
            return {"ok": False, "error": "Unknown request"}
        if op != "get":
            data["filterSets"] = sorted(sets, key=lambda s: s["name"].lower())
            self.write(data)
        return {"ok": True, "filterSets": sorted(sets, key=lambda s: s["name"].lower())}


# --------------------------------------------------------------- move to another PC --

MOVED_SETTINGS = ("accountId", "restletUrl", "clientId", "authMode", "fiscalYearStartMonth", "updateUrl",
                  "autoUpdateCheck")


def export_bundle(config, licence_text, logo, prefs, layouts, queries, version, name=""):
    """Writes "TQ setup <name> <date>.json" in Documents\\TQ Connector and returns (path, summary)."""
    bundle = {
        "kind": BUNDLE_KIND, "version": 1, "connector": version, "made": time.strftime("%Y-%m-%d %H:%M"),
        "name": str(name or "")[:60],
        "settings": {k: config[k] for k in MOVED_SETTINGS if config and k in config},
        "licence": licence_text or "",
        "logo": {"type": logo[1], "data": base64.b64encode(logo[0]).decode()} if logo else None,
        "filterSets": (prefs or {}).get("filterSets") or [],
        "layouts": [l for l in layouts or [] if isinstance(l, dict) and isinstance(l.get("lines"), list)][:50],
        "queries": [q for q in queries or [] if isinstance(q, dict)][:100],
    }
    safe = re.sub(r"[^\w\- ]+", "", bundle["name"]).strip()
    path = os.path.join(documents_dir(), "TQ setup %s%s.json" % (safe + " " if safe else "", time.strftime("%Y-%m-%d %H%M")))
    with open(path, "w", encoding="utf-8") as f:
        json.dump(bundle, f, indent=1)
    summary = {"settings": bool(bundle["settings"]), "licence": bool(bundle["licence"]), "logo": bool(bundle["logo"]),
               "filterSets": len(bundle["filterSets"]), "layouts": len(bundle["layouts"]), "queries": len(bundle["queries"])}
    return path, summary


def read_bundle(bundle):
    """Checks a setup file: returns it, or raises ValueError with a plain reason."""
    if not isinstance(bundle, dict) or bundle.get("kind") != BUNDLE_KIND:
        raise ValueError("That is not a TQ Connector setup file (made with Move to another PC > Export).")
    for key, kind in (("settings", dict), ("filterSets", list), ("layouts", list), ("queries", list)):
        if not isinstance(bundle.get(key) or kind(), kind):
            raise ValueError("The setup file is damaged (%s)." % key)
    return bundle

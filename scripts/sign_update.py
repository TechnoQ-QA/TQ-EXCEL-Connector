"""Techno Q only: publish a new version of the connector.

    python scripts/sign_update.py build/TQConnectorSetup.exe --version 1.1.0 \
        --url https://downloads.example.com/tq/TQConnectorSetup-1.1.0.exe --notes "Monthly pack, budgets"

writes update.json next to the setup file. Upload both: the setup file at --url, update.json at the address the
customers have under Updates on their setup page. Each connector checks update.json, verifies the signature
with the licence key and the setup file's sha256, and only then runs it.

Signed with the same private key as the licences (%USERPROFILE%\\TQ Licence Key\\private-key.json).
"""

import argparse
import hashlib
import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "server"))
sys.path.insert(0, os.path.join(ROOT, "scripts"))
import updates  # noqa: E402
from make_licence import load_key  # noqa: E402


def main():
    parser = argparse.ArgumentParser(description="Sign a TQ Connector update")
    parser.add_argument("setup", help="the setup file, e.g. build/TQConnectorSetup.exe")
    parser.add_argument("--version", required=True, help="its version, e.g. 1.1.0 (must be higher than installed)")
    parser.add_argument("--url", required=True, help="https address the setup file will be downloaded from")
    parser.add_argument("--notes", default="", help="one line saying what is new")
    args = parser.parse_args()
    if not args.url.lower().startswith("https://"):
        sys.exit("--url must start with https://")
    with open(args.setup, "rb") as f:
        digest = hashlib.sha256(f.read()).hexdigest()
    manifest = updates.sign_manifest(load_key(), args.version, args.url, digest, args.notes)
    out = os.path.join(os.path.dirname(os.path.abspath(args.setup)), "update.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=1)
    print("Wrote %s (sha256 %s)" % (out, digest))


if __name__ == "__main__":
    main()

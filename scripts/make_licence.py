"""Techno Q only: make a licence for a customer.

    python scripts/make_licence.py --customer "World of Hospitality" --account 1234567 --users 5 --expires 2027-12-31

writes licence-<account>.json, to send to the customer (they paste it on the connector's setup page).

The private key that signs licences is kept OUTSIDE the project, in
%USERPROFILE%\\TQ Licence Key\\private-key.json. Keep a safe copy of that file: without it no new
licence can be made, and anyone who has it can make licences. It is never shipped to customers.

    python scripts/make_licence.py --new-key

makes that key the first time and writes its public half into server/licence_key.py. Making a
new key later makes every licence issued so far invalid.
"""

import argparse
import datetime
import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "server"))
import licence  # noqa: E402
import tqcrypto  # noqa: E402

KEY_DIR = os.path.join(os.path.expanduser("~"), "TQ Licence Key")
KEY_PATH = os.path.join(KEY_DIR, "private-key.json")
PUBLIC_PATH = os.path.join(ROOT, "server", "licence_key.py")


def new_key():
    if os.path.exists(KEY_PATH):
        sys.exit("A key already exists at %s. Move it away first if you really want a new one "
                 "(every licence made with the old key stops working)." % KEY_PATH)
    key = tqcrypto.generate_key(3072)
    os.makedirs(KEY_DIR, exist_ok=True)
    with open(KEY_PATH, "w", encoding="utf-8") as handle:
        json.dump({k: str(v) for k, v in key.items()}, handle, indent=1)
    with open(PUBLIC_PATH, "w", encoding="utf-8", newline="\n") as handle:
        handle.write('"""Public key for checking licences (made by scripts/make_licence.py --new-key)."""\n\n'
                     "PUBLIC_KEY = {\n    \"n\": %d,\n    \"e\": %d,\n}\n" % (key["n"], key["e"]))
    print("Private key: %s  (keep a safe copy; never share it)" % KEY_PATH)
    print("Public key written to %s" % PUBLIC_PATH)


def load_key():
    if not os.path.exists(KEY_PATH):
        sys.exit("No private key at %s. Run with --new-key once, or copy the key file there." % KEY_PATH)
    with open(KEY_PATH, "r", encoding="utf-8") as handle:
        return {k: int(v) for k, v in json.load(handle).items()}


def main():
    parser = argparse.ArgumentParser(description="Make a TQ Connector licence")
    parser.add_argument("--new-key", action="store_true", help="make the signing key (once)")
    parser.add_argument("--customer")
    parser.add_argument("--account", help="NetSuite account ID, e.g. 1234567 or 1234567_SB1")
    parser.add_argument("--users", type=int, default=5)
    parser.add_argument("--expires", help="last day, e.g. 2027-12-31")
    parser.add_argument("--out", help="file to write (default licence-<account>.json)")
    args = parser.parse_args()
    if args.new_key:
        new_key()
        return
    if not (args.customer and args.account and args.expires):
        parser.error("--customer, --account and --expires are needed")
    expires = datetime.date.fromisoformat(args.expires)
    made = licence.make(load_key(), args.customer, args.account, args.users, expires)
    check = licence.check(made, made["account"])
    if not check["ok"]:
        sys.exit("The licence does not check against server/licence_key.py: %s" % check["message"])
    out = args.out or "licence-%s.json" % made["account"]
    with open(out, "w", encoding="utf-8") as handle:
        json.dump(made, handle, indent=1)
    print("%s  (%s)" % (out, check["message"]))


if __name__ == "__main__":
    main()

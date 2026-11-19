"""Checks the NetSuite connection and prints balances to compare with the
NetSuite Trial Balance.

    python scripts\\check_netsuite.py                  (uses server\\config.json)
    python scripts\\check_netsuite.py --period "Oct 2026"
    python scripts\\check_netsuite.py --mock           (sample data)

Prints results only, never the credentials.
"""

import argparse
import datetime
import functools
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "server"))
import server  # noqa: E402


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--mock", action="store_true")
    parser.add_argument("--period", help="period name; default is the latest period that has ended")
    parser.add_argument("--rows", type=int, default=25, help="how many balances to print")
    args = parser.parse_args()

    if args.mock:
        call = server.MockBackend().handle
    else:
        try:
            cfg = server.load_config()
        except (ValueError, OSError) as err:
            sys.exit("Config problem: %s" % err)
        if cfg is None:
            sys.exit("server\\config.json not found. Copy config.example.json and fill it in.")
        call = functools.partial(server.call_netsuite, cfg)

    def step(label, payload):
        out = call(payload)
        if not out.get("ok"):
            sys.exit("FAILED at %s: %s" % (label, out.get("error")))
        return out

    ping = step("ping", {"action": "ping"})
    print("1. Connected: account %s, user %s, role %s" % (ping["account"], ping["user"], ping["role"]))

    periods = step("periods", {"action": "periods"})["periods"]
    print("2. Periods found: %d" % len(periods))
    if args.period:
        period = args.period
    else:
        today = datetime.date.today().isoformat()
        ended = [p for p in periods if p["end"] <= today]
        if not ended:
            sys.exit("No period has ended yet; pass --period.")
        period = max(ended, key=lambda p: p["end"])["name"]

    accounts = step("accounts", {"action": "accounts"})["accounts"]
    print("3. Accounts with a number: %d" % len(accounts))
    if not accounts:
        sys.exit("No numbered accounts returned; check the role's permissions.")

    results = []
    for start in range(0, len(accounts), 500):
        requests = [{"account": a["number"], "fromPeriod": period} for a in accounts[start:start + 500]]
        results += step("balances", {"action": "balances", "requests": requests})["results"]
    errors = [(a, r["error"]) for a, r in zip(accounts, results) if "error" in r]
    nonzero = [(a, r["value"]) for a, r in zip(accounts, results) if r.get("value")]
    print("4. Balances for %s: %d non-zero, %d zero, %d errors"
          % (period, len(nonzero), len(accounts) - len(nonzero) - len(errors), len(errors)))
    for account, error in errors[:5]:
        print("   error  %-10s %s" % (account["number"], error))

    print("\n   Compare these with the NetSuite Trial Balance for %s" % period)
    print("   (balance sheet = closing balance; income/expense = activity in the period):\n")
    for account, value in nonzero[:args.rows]:
        print("   %-10s %-40.40s %-14s %16s" % (account["number"], account["name"], account["type"],
                                                  format(value, ",.2f")))

    if nonzero:
        account, value = nonzero[0]
        drill = step("transactions", {"action": "transactions",
                                      "request": {"account": account["number"], "fromPeriod": period}})
        total = round(sum(r["amount"] for r in drill["rows"]), 2)
        if drill["truncated"]:
            verdict = "list cut at 1,000 lines, total not compared"
        else:
            verdict = "matches balance" if abs(total - value) < 0.005 else "DOES NOT MATCH balance %.2f" % value
        print("\n5. Drill-down on %s: %d lines, total %s (%s)"
              % (account["number"], len(drill["rows"]), format(total, ",.2f"), verdict))
    print("\nDone.")


if __name__ == "__main__":
    main()

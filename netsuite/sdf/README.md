# TQ Excel Connector: the NetSuite package

Everything the connector needs inside a customer's NetSuite account, as one SuiteCloud (SDF) project that is
deployed in one step instead of being set up by hand.

| In the package | What it is |
|---|---|
| `FileCabinet/SuiteScripts/TQ Excel Connector/` | The three scripts: `tq_excel_core.js` (the logic), `tq_excel_restlet.js` (what the add-in calls), `tq_excel_elevated.js` |
| `customscript_tq_excel_restlet` | The RESTlet and its deployment (released, all roles; it runs as the caller's role) |
| `customscript_tq_excel_elevated` | The Suitelet the RESTlet calls for balances, deployed to **run as Administrator**, not available without login. Who may ask is checked in `tq_excel_core.js` (the caller needs a role that may run financial reports) |
| `customrole_tq_excel_user` | Role **TQ Excel User**: financial statements, trial balance and general ledger reports; log in with access tokens (TBA) and OAuth 2.0; view accounts, customers, vendors, items; find transactions and searches |

The seat list (`tq_seats.json`) is made by the script itself in the same folder the first time a PC connects.

## Prepare the package (Techno Q)

```
python scripts/package_netsuite.py
```

copies the scripts from `netsuite/` into the package and checks it: every object is well-formed, points at a
script in the package, and the Suitelet has the IDs the RESTlet calls and runs as Administrator. The test suite
runs the same check. Always edit the scripts in `netsuite/`, never the copies in the package.

## Deploy to an account

One-time tools on the PC that deploys (free):

- Node.js 22 LTS: `winget install --id OpenJS.NodeJS.LTS -e`
- Oracle JDK 17 or 21 (the SuiteCloud CLI needs Java)
- SuiteCloud CLI: `npm install -g @oracle/suitecloud-cli`

In the customer's NetSuite (an administrator, once):

1. Setup > Company > Enable Features > SuiteCloud: tick **Server SuiteScript**, **SuiteCloud Development
   Framework**, **Token-Based Authentication** and **OAuth 2.0**.
2. The deploying user needs a role with the **SuiteCloud Development Framework** permission (Administrator has it).

Then, in `netsuite/sdf`:

```
suitecloud account:setup          (signs in to that NetSuite account in the browser, once per account)
suitecloud project:validate --server
suitecloud project:deploy
```

After the deploy:

1. Give each Excel user the role **TQ Excel User** (Lists > Employees > the employee > Access > Roles), next to
   their own roles. They keep their own role for NetSuite itself.
2. Copy the RESTlet's **External URL**: Customization > Scripting > Script Deployments >
   TQ Excel Connector RESTlet. The connector's setup page needs it.
3. Until per-user sign-in is built: Setup > Integration > Manage Integrations > New (tick Token-Based
   Authentication), and an access token for each user (Setup > Users/Roles > Access Tokens). This step goes away
   with "Sign in with NetSuite".

**An account that was set up by hand** (as the Techno Q sandbox was): the deploy updates the Suitelet in place,
because it has the same IDs. If the old RESTlet had another script ID, the deploy adds this one next to it. Point
the connector at the new RESTlet's URL, then make the old script record inactive. Also delete any stray copy of
`tq_excel_restlet.js` left directly in SuiteScripts.

## Deployed so far

- **Script 9, 23 Nov 2026:** sandbox 65987_SB1 (30 of 30 steps) and Techno Q's account 12253116, which feeds bundle
  715956 (30 of 30 steps). Deploy to the sandbox by setting `defaultAuthId` in `project.json` to `tq-sandbox` for
  the run (the CLI has no account option), then back to `tq-publisher`.

- **Techno Q sandbox (65987_SB1), Nov 2026.** Validated by NetSuite (21 of 21 steps, 0 errors) and deployed (29 of
  29 steps). The deploy created the role TQ Excel User and updated the hand-made RESTlet and Suitelet in place
  (they already had these IDs, so nothing was duplicated); the script files kept their File Cabinet IDs. The
  connector answered as before afterwards. A stray older `tq_excel_restlet.js` directly in SuiteScripts (file
  14736) is unused and can be deleted by hand.
- **Nov 19 2026:** tq_excel_core.js version 6 (seats per NetSuite user, seat list and release) deployed the
  same way, 0 errors; ping answers version 6.
- **Nov 19 2026:** version 7 (company report library, tq_library.json) deployed the same way, 0 errors; ping
  answers version 7, and a library entry was saved, read and removed again.

On this PC the CLI runs from a terminal as `%APPDATA%\npm\suitecloud.cmd` (Node.js and the npm folder must be on
the PATH). The sandbox sign-in is saved under the authentication ID `tq-sandbox`; `netsuite/sdf/project.json`
records it and is not committed.

## Later: a SuiteApp

This is an account customization project, deployed into each customer's account. Listing it as a SuiteApp
(installed by the customer from the SuiteApp marketplace, and updated from there) needs Techno Q to join Oracle's
SuiteCloud Developer Network. The same objects carry over; only the project type and the publisher ID change.

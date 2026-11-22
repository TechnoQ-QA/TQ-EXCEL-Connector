// TQ Connector for NetSuite: Techno Q Delivery Guide (internal). node delivery.js <images dir> <output.docx>
// Everything Techno Q does for a new client or a test: licence, delivery pack, test install, renewals, releases.
const c = require('./common');
const { P, H1, H2, BR, bullets, steps, table, box, spacer } = c;
const OUT = process.argv[3];

const BUNDLE = '715956';
const PROJECT = 'C:\\Users\\<you>\\Desktop\\Development\\netsuite-excel-connector';

const body = [
  ...c.cover('Techno Q Delivery Guide', 'Licences, the delivery pack, test installations, renewals and releases',
    ['Version 1.5  ·  November 2026', 'Internal: Techno Q only. Do not send to clients.',
      'NetSuite bundle ' + BUNDLE + '  ·  Connector 1.5.3']),
  ...c.toc(),

  // ------------------------------------------------------------------ 1
  H1('1. Who does what'),
  P('A new client needs three things from Techno Q: a **licence** for their NetSuite account, the **setup file**, and the **guides**. Everything else the client does themselves with the Installation Guide.'),
  table([
    ['Step', 'Who', 'Where it is described'],
    ['Get the client’s NetSuite account ID(s), number of users, end date', 'Techno Q (sales / project)', '3.1'],
    ['Make the licence', 'Techno Q, on the PC with the licence key', '3'],
    ['Send the delivery pack', 'Techno Q', '2'],
    ['Install bundle ' + BUNDLE + ', give the role', 'Client’s NetSuite administrator', 'Installation Guide, section 2'],
    ['Run the setup on each PC and sign in', 'Client’s users or IT', 'Installation Guide, section 3'],
    ['Acceptance checklist and sign-off', 'Client and Techno Q', 'Installation Guide, section 4'],
    ['Renewals, more users, updates', 'Techno Q', '5 and 6'],
  ], [4300, 2900, 2438], { header: true }),
  spacer(),
  box('warn', 'What never leaves Techno Q', [
    'The **project folder** (source code, the build folder, `server\\config.json` with Techno Q’s own NetSuite settings).',
    'The **licence key**: `%USERPROFILE%\\TQ Licence Key\\private-key.json`. Anyone who has it can make licences and sign updates.',
    'The **client secret** of Techno Q’s NetSuite integration. The connector never uses it.',
  ]),

  // ------------------------------------------------------------------ 2
  BR(),
  H1('2. The delivery pack'),
  P('Do **not** zip the project folder and send it. The client needs only the files below. Put them in one folder (or zip that folder) and send it by email, Teams or a USB stick.'),
  table([
    ['File', 'Where to find it', 'Notes'],
    ['`TQConnectorSetup.exe`', '`build\\TQConnectorSetup.exe` in the project folder', 'The same file for every client and every PC. Signed by Techno Q.'],
    ['`licence-<account>.json`', 'Made in section 3', 'One per NetSuite account. The client’s administrator chooses it once, during the setup.'],
    ['`TQ Connector - Installation Guide.docx`', '`docs\\guides`', 'For the client’s administrator, users and IT.'],
    ['`TQ Connector - User Guide.docx`', '`docs\\guides`', 'For the Excel users.'],
    ['`TechnoQ.cer` (optional)', '`installer\\TechnoQ.cer`', 'Only if the client’s IT wants Windows to show Techno Q as the publisher (Installation Guide 5.1).'],
  ], [3300, 3100, 3238], { header: true }),
  H2('2.1 Check the setup file before sending'),
  steps([
    'Right-click `TQConnectorSetup.exe` > **Properties** > **Details**: **Product version** must be the version you mean to deliver (now **1.5.3**).',
    'Its fingerprint, for the client’s IT if they ask: in PowerShell `Get-FileHash .\\TQConnectorSetup.exe` (the build prints the same SHA-256).',
    'If the build folder is older than the last change, rebuild first: `powershell -ExecutionPolicy Bypass -File installer\\build.ps1` in the project folder (about 2 minutes).',
  ]),
  box('tip', 'Antivirus on Techno Q PCs', 'Kaspersky may flag a freshly built setup (behaviour detection, “PDM:Trojan.Win32.Generic”). It is a false positive: never choose “Disinfect and restart”. Ask IT to exclude the project’s `build` folder and `%LOCALAPPDATA%\\TQ Connector`. A code-signing certificate from a recognised authority is the lasting fix.'),

  // ------------------------------------------------------------------ 3
  BR(),
  H1('3. Making a licence'),
  P('A licence is a small signed text file. It names the client, **one** NetSuite account, the number of users and the last day it is valid. Because it is signed, nobody can change those values or use it on another account.'),
  H2('3.1 What to get from the client'),
  table([
    ['Item', 'Where the client finds it', 'Example'],
    ['Company name (shown in Excel)', '', 'Leylaty'],
    ['NetSuite account ID', 'Setup > Company > Company Information > **Account ID**', '`1234567` (production), `1234567_SB1` (sandbox)'],
    ['Number of users', 'Each NetSuite user who signs in from Excel takes one seat', '5'],
    ['Last day', 'End of the subscription', '2027-12-31'],
  ], [2800, 4100, 2738], { header: true }),
  P('A client with a sandbox and production needs **two licences**, one per account ID.'),

  H2('3.2 Run the script'),
  P('Only on the Techno Q PC that holds `%USERPROFILE%\\TQ Licence Key\\private-key.json` (with Python 3 and the project folder). Open PowerShell and run:'),
  table([
    ['PowerShell'],
    ['`cd ' + PROJECT + '`'],
    ['`python scripts\\make_licence.py --customer "Leylaty" --account 1234567 --users 5 --expires 2027-12-31`'],
  ], [9638], { header: true }),
  steps([
    'The script signs the licence, checks it against the key built into the connector, and prints for example `licence-1234567.json  (Licensed to Leylaty for 5 users until 2027-12-31.)`.',
    'The file is in the project folder. Do not edit it: any change breaks the signature. It is never committed to Git.',
    'Move it to a client folder outside the project, for example `%USERPROFILE%\\TQ Licence Key\\issued\\Leylaty\\`, and add a line to Techno Q’s licence register (client, account, users, end date, issued date).',
  ]),
  table([
    ['Option', 'Meaning'],
    ['`--users N`', 'Seats. Default 5.'],
    ['`--expires YYYY-MM-DD`', 'The last day the licence works. A date in the past is refused.'],
    ['`--out <file>`', 'Write the file somewhere else (for example straight into the client folder).'],
  ], [3000, 6638], { firstCol: true }),

  H2('3.3 How the licence travels'),
  bullets([
    'The client’s **administrator** chooses the file on the setup’s **Licence** page (or pastes it later on Setup > Licence) and signs in.',
    'The connector checks it and stores it in NetSuite (`tq_licence.json` in the File Cabinet).',
    'Every other user’s PC fetches it from NetSuite when the user signs in. Nobody copies the file to other PCs.',
  ]),

  // ------------------------------------------------------------------ 4
  BR(),
  H1('4. Test installation: new PC, new account'),
  P('Use this to test a release, or to show the product, on a PC that never had TQ Connector and a NetSuite account that never had the bundle (for example a client’s sandbox or a NetSuite test account).'),
  H2('4.1 Prepare (Techno Q PC)'),
  steps([
    'Get the test account’s **account ID** and an **administrator login** for it.',
    'Make a short test licence: `python scripts\\make_licence.py --customer "Techno Q test" --account <ID> --users 2 --expires <today + 30 days>`.',
    'Copy to the test PC (USB stick, Teams or OneDrive): `TQConnectorSetup.exe` and the licence file. Nothing else.',
  ]),
  H2('4.2 NetSuite (test account, as administrator)'),
  steps([
    '**Setup > Company > Enable Features > SuiteCloud**: Server SuiteScript and OAuth 2.0 ticked.',
    '**Customization > SuiteBundler > Search & Install Bundles**: search `' + BUNDLE + '`, install **TQ Excel Connector**, wait for status **Installed**.',
    'Give your test user the role **TQ Excel User** (an administrator can also sign in with Administrator).',
  ]),
  H2('4.3 The test PC'),
  steps([
    'Sign in to Windows as the person who will use Excel. Close Excel.',
    'Run `TQConnectorSetup.exe`. SmartScreen: **More info > Run anyway**.',
    'Type the test **account ID**. On the Licence page choose the licence file.',
    'Windows asks to install the **localhost** certificate: **Yes**.',
    '**Finish**: the setup page opens. Click **Sign in with NetSuite**, sign in as the administrator, **Allow**.',
    'Open Excel > **TQ Connector** tab > **Open panel**.',
  ]),
  H2('4.4 What to check'),
  table([
    ['☐', 'Check', 'Expected'],
    ['☐', 'Setup page > Overview', 'Signed in as you; licence “Techno Q test, 2 users, until <date>”'],
    ['☐', 'Financial year', 'Year-to-date figures in the demo pack match NetSuite (the start month is read from NetSuite; nobody typed it)'],
    ['☐', 'Excel panel welcome checks', 'Connected, signed in, licensed'],
    ['☐', 'Gear menu > Build a demo pack', 'Figures match NetSuite’s Income Statement and Balance Sheet'],
    ['☐', 'Drill down on a figure', 'Transactions listed; opening one goes to NetSuite'],
    ['☐', 'File Cabinet in NetSuite', '`tq_licence.json` and `tq_seats.json` exist'],
    ['☐', 'Second user on a second PC: setup with **no** licence file, then sign in', 'Licence fetched from NetSuite; Setup > Seats shows 2 of 2'],
    ['☐', 'Third user (licence for 2)', 'Refused with “All 2 seats of the licence are in use”'],
  ], [500, 4700, 4438], { header: true }),
  H2('4.5 Test a renewal'),
  steps([
    'Make a second licence for the same account with more users, for example `--users 3`, and the same or a later date.',
    'On the administrator’s PC: Setup > **Licence**, paste the new file’s text, Save.',
    'Other PCs take it within 6 hours, or at once after the connector restarts (sign out of Windows and in again). Setup > Overview then shows 3 users.',
  ]),
  H2('4.6 Clean up'),
  bullets([
    'Test PC: Windows Settings > Apps > **TQ Connector for NetSuite** > Uninstall.',
    'NetSuite: Setup > **Seats** > release the test seats; uninstall the bundle if the account is not staying (Customization > SuiteBundler > Search & Install Bundles > List).',
    'If `tq_licence.json` and `tq_seats.json` remain in the File Cabinet, delete them.',
  ]),

  // ------------------------------------------------------------------ 5
  BR(),
  H1('5. Licence life: renewals, users, expiry'),
  table([
    ['Situation', 'What Techno Q does', 'What the client sees'],
    ['Licence valid', 'Nothing', 'Setup > Overview shows “until <date>”. Seats list on Setup > Seats.'],
    ['Last 30 days', 'Send the renewal.', 'A yellow reminder in the Excel panel (“ends in N days”); in the last 7 days it cannot be hidden.'],
    ['Renewal (new end date)', 'Make a new licence, same account, new `--expires`. Send it.', 'Administrator pastes it on Setup > Licence. Every PC takes it within 6 hours.'],
    ['More users', 'New licence with the higher `--users` (keep the end date).', 'Same as a renewal. New users can sign in once it has reached NetSuite.'],
    ['Fewer users', 'New licence with the lower `--users`.', 'Extra seats stop working; an administrator releases the unused ones on Setup > Seats.'],
    ['Licence ended', 'Make the renewal (better: before the end date).', 'From the day after the end date, NetSuite data stops: “The licence ended on <date>. Ask Techno Q to renew it.” Workbooks stay; sample data still works. Once the renewal is pasted, every PC picks it up within about a minute.'],
    ['Client moves from sandbox to production', 'A licence for the production account ID.', 'The administrator installs the bundle in production and chooses the new licence.'],
    ['Sandbox refreshed from production', 'Nothing, if the sandbox licence still fits.', 'The bundle may need installing again in the sandbox (Installation Guide, section 6).'],
    ['Licence for the wrong account', 'Make it again with the right account ID.', '“The licence is for NetSuite account X, not Y”; nothing is saved.'],
  ], [2400, 3400, 3838], { header: true }),
  spacer(),
  box('note', 'Reminder before the end date', 'From 30 days before the end date every user sees a yellow reminder at the top of the Excel panel, and the setup page shows the days left. Still keep the end dates in Techno Q’s licence register and send renewals before the reminder starts.'),

  // ------------------------------------------------------------------ 6
  BR(),
  H1('6. Releasing a new version'),
  H2('6.1 The connector (setup file)'),
  steps([
    'Change the code, raise `VERSION` in `server\\server.py` and `<Version>` in `addin\\manifest.xml`, run the tests: `python -m unittest discover -s tests`.',
    'Build: `powershell -ExecutionPolicy Bypass -File installer\\build.ps1`. Result: `build\\TQConnectorSetup.exe`, signed.',
    'Send the new setup file to clients (or publish it for automatic updates: `scripts\\sign_update.py`, see the Technical Reference 6.11). Running it over an older version keeps settings, licence and sign-in.',
  ]),
  H2('6.2 The NetSuite scripts (bundle)'),
  steps([
    'Deploy the SDF package to Techno Q’s account 12253116 (auth ID `tq-publisher`).',
    'Customization > SuiteBundler > Create Bundle > List > **' + BUNDLE + '** > **Push** to the client accounts. Clients can also update from their Installed Bundles list.',
  ]),

  // ------------------------------------------------------------------ 7
  H1('7. Keeping the key safe'),
  bullets([
    'Keep an offline copy of `private-key.json` (an encrypted USB stick or Techno Q’s password manager).',
    'Lost key: no new licence can be made and no update signed. `make_licence.py --new-key` makes a new pair, but every licence already issued stops working.',
    'New Techno Q PC for licences: copy `private-key.json` into `%USERPROFILE%\\TQ Licence Key\\` on that PC, with the project folder from Git.',
  ]),

  H1('8. Troubleshooting (Techno Q side)'),
  table([
    ['What you see', 'What to do'],
    ['make_licence: “No private key at …”', 'This PC has no key. Use the licence PC, or copy the key file from the safe copy.'],
    ['make_licence: “does not check against server/licence_key.py”', 'The key and the connector do not belong together. Do not send the file; check you use the right key.'],
    ['Client: “The licence is not valid (its signature does not match)”', 'The file was edited or cut short. Send it again as an attachment, not pasted in an email body.'],
    ['Client: “The licence is for NetSuite account X, not Y”', 'Wrong account ID (often sandbox vs production). Make it again with the ID the client typed in the setup.'],
    ['Client cannot find bundle ' + BUNDLE, 'Check it is still **Public** (Customization > SuiteBundler > Create Bundle > List > Set Availability) in account 12253116.'],
    ['Client: “No licence” on a user’s PC', 'The administrator has not signed in with the licence yet, or is not an administrator in NetSuite.'],
  ], [4300, 5338], { firstCol: true }),
];

c.save(c.doc('Techno Q Delivery Guide', body), OUT);

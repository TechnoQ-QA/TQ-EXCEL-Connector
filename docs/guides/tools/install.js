// TQ Connector for NetSuite: Installation Guide (first client). node install.js <images dir> <output.docx>
const path = require('path');
const c = require('./common');
const { P, H1, H2, H3, BR, bullets, steps, table, box, picture, pictures2, spacer } = c;
const IMG = process.argv[2];
const OUT = process.argv[3];
const img = (n) => path.join(IMG, n + '.png');
const W = c.TEXT_W;

const body = [
  ...c.cover('Installation Guide', 'Setting up a new client: NetSuite, the connector and every Excel user',
    ['Version 1.4  ·  November 2026', 'For: Techno Q consultants, the client’s NetSuite administrator and IT',
      'NetSuite script version 7  ·  Connector 1.4.0']),
  ...c.toc(),

  // ------------------------------------------------------------------ 1
  H1('1. About this guide'),
  P('This guide takes a new client from nothing to every finance user working with live NetSuite figures in Excel. Follow the parts in order. Each part says **who** does it and roughly **how long** it takes.'),
  table([
    ['Part', 'What', 'Who', 'Time'],
    ['A', 'Set up NetSuite: features, the TQ package, the sign-in integration, user roles', 'Techno Q consultant with the client’s NetSuite administrator', '45 min'],
    ['B', 'Prepare the client’s files: licence and settings', 'Techno Q', '10 min'],
    ['C', 'Trust Techno Q’s certificate, allow the files in antivirus', 'Client IT (once for the company)', '15 min'],
    ['D', 'Install on each PC, sign in, first test', 'Each user (or IT), with Techno Q for the first one', '5 min per PC'],
    ['E', 'Acceptance checklist and sign-off', 'Techno Q and the client', '15 min'],
    ['F', 'After go-live: users, seats, new PCs, updates, support', 'Client administrator, Techno Q', 'as needed'],
  ], [700, 4300, 3300, 1338]),
  spacer(),
  box('key', 'In one sentence', 'A small TQ package is deployed into the client’s NetSuite; a small program, the TQ Connector, is installed on each PC and connects Excel to NetSuite; each user signs in with their own NetSuite login, so NetSuite’s roles decide what they see.'),

  H2('1.1 What is installed where'),
  table([
    ['Where', 'What', 'Notes'],
    ['Client’s NetSuite', 'TQ Excel Connector package: three scripts (RESTlet, Suitelet, library), the role **TQ Excel User**, and a sign-in integration record', 'Deployed once per NetSuite account (sandbox and production separately)'],
    ['Each PC (per Windows user)', '**TQ Connector**: a small program in the user’s own folder that Excel’s add-in talks to, and the Excel add-in itself (a **TQ Connector** tab on the ribbon)', 'No administrator rights needed. Desktop Excel on Windows.'],
    ['Techno Q', 'The licence for the client’s NetSuite account (number of users, expiry date)', 'Made by Techno Q after the account ID is known'],
  ], [2000, 4700, 2938]),

  H2('1.2 What Techno Q sends the client'),
  table([
    ['File', 'What it is'],
    ['`TQConnectorSetup.exe`', 'The setup file, run on each PC. Signed by Techno Q.'],
    ['`TQConnectorSetup.exe.sha256.txt`', 'Its checksum, to check the file is genuine.'],
    ['`TechnoQ.cer`', 'Techno Q’s public code-signing certificate, trusted once by IT.'],
    ['`TQConnector.settings.json`', 'The client’s NetSuite connection details, so nobody types them on each PC (Part B).'],
    ['`licence.json`', 'The client’s licence (Part B).'],
    ['This guide, the user guide', 'For IT, the administrator and the users.'],
  ], [3300, 6338]),

  // ------------------------------------------------------------------ 2
  H1('2. Before you start'),
  H2('2.1 Requirements'),
  table([
    ['Item', 'Requirement'],
    ['PCs', 'Windows 10 or 11, 64-bit. About 100 MB free disk space per user.'],
    ['Excel', 'Microsoft 365 (desktop), Excel 2021 or 2024 for Windows. Excel on the web and on Mac are not supported.'],
    ['Windows account', 'Each person installs under their own Windows sign-in. No administrator rights are needed.'],
    ['NetSuite', 'Any edition. An administrator for Part A. OneWorld, multi-currency and multi-book are supported but not required.'],
    ['Network', 'The PCs reach NetSuite over the internet (HTTPS). Nothing else is opened: the connector only listens on the PC itself.'],
    ['Techno Q consultant’s PC (Part A)', 'Node.js 22 LTS, Oracle Java 17 or 21, SuiteCloud CLI (`npm install -g @oracle/suitecloud-cli`), and the TQ package (`netsuite\\sdf` in Techno Q’s repository).'],
  ], [2600, 7038], { firstCol: true }),

  H2('2.2 Information to collect from the client'),
  table([
    ['Item', 'Example', 'Used in'],
    ['NetSuite account ID (production and sandbox)', '1234567 and 1234567_SB1', 'A, B'],
    ['Users who will use Excel (names, NetSuite employee records)', 'Finance Manager, two accountants', 'A6, licence'],
    ['First month of the fiscal year', 'January', 'B2'],
    ['Who may upload budgets and journal entries', 'Finance Manager only', 'A6'],
    ['Company logo for reports (optional; else the NetSuite logo)', 'PNG, up to 1 MB', 'F'],
    ['How IT distributes software and trusts certificates', 'Group Policy, Intune, or by hand', 'C'],
  ], [4200, 3200, 2238]),
  box('tip', 'Sandbox first', 'Set everything up in the client’s **sandbox** account first, test, then repeat Parts A and B for **production** (a separate deploy, integration record and licence).'),

  // ------------------------------------------------------------------ A
  BR(),
  H1('Part A. NetSuite setup'),
  P('**Who:** Techno Q consultant, signed in with the client’s NetSuite administrator.   **Time:** about 45 minutes.'),

  H2('A1. Enable the features'),
  P('In NetSuite: **Setup > Company > Enable Features > SuiteCloud**. Tick, then Save:'),
  table([
    ['Feature', 'Section', 'Why'],
    ['Server SuiteScript', 'SuiteScript', 'Runs the TQ scripts'],
    ['SuiteCloud Development Framework', 'SuiteCloud Development Framework', 'Lets the package be deployed in one step'],
    ['OAuth 2.0', 'Manage Authentication', 'Sign in with NetSuite (each user’s own login)'],
    ['Token-Based Authentication', 'Manage Authentication', 'Only for the deploy tool’s own sign-in, if NetSuite asks for it'],
  ], [3200, 3300, 3138]),
  P('Accept the SuiteCloud terms if NetSuite shows them.'),

  H2('A2. Deploy the TQ package'),
  P('On the consultant’s PC, in a terminal, in the package folder (`netsuite\\sdf`):'),
  steps([
    'Sign in to the client’s account (once per account). A browser opens; the administrator signs in and allows it. Give the authentication an ID such as `client-sb` or `client-prod`:',
    ['`suitecloud account:setup`'],
    'Check the package against the account (no changes are made):',
    ['`suitecloud project:validate --server`', 'Expected: 0 errors.'],
    'Deploy:',
    ['`suitecloud project:deploy`', 'Expected: every step ticked, ending with the two script deployments. Two warnings about "allroles" are normal.'],
  ]),
  P('The package creates, in **SuiteScripts > TQ Excel Connector**: `tq_excel_core.js`, `tq_excel_restlet.js`, `tq_excel_elevated.js`; the script records **TQ Excel Connector RESTlet** and **TQ Excel Connector elevated** with their deployments; and the role **TQ Excel User**.'),
  box('note', 'If the account already had TQ scripts set up by hand', 'The deploy updates them in place when they have the same IDs. If an old RESTlet had another ID, the new one is added next to it: use the new RESTlet’s URL and make the old script record inactive.'),

  H2('A3. Check the two script deployments'),
  P('**Customization > Scripting > Script Deployments**, filter on "TQ Excel":'),
  table([
    ['Deployment', 'Status', 'Execute as role', 'Audience'],
    ['TQ Excel Connector RESTlet', 'Released', '(the current role)', 'All roles'],
    ['TQ Excel Connector elevated (Suitelet)', 'Released', '**Administrator**', 'All roles; Available without login: **No**'],
  ], [3400, 1500, 2100, 2638]),
  box('warn', 'Important', 'The Suitelet must run as **Administrator**. The add-in then sees every posting, so balances match NetSuite’s reports, while the script itself checks that each caller may run financial reports. If it does not run as Administrator, users see "Deploy the elevated Suitelet with Execute As Role = Administrator".'),

  H2('A4. Copy the RESTlet address'),
  P('Open the **TQ Excel Connector RESTlet** deployment and copy its **External URL**. It looks like:'),
  P('`https://1234567.restlets.api.netsuite.com/app/site/hosting/restlet.nl?script=101&deploy=1`'),
  P('Keep it for Part B.'),

  H2('A5. Create the sign-in integration'),
  P('**Setup > Integration > Manage Integrations > New**. Fill in exactly:'),
  table([
    ['Field', 'Value'],
    ['Name', 'TQ Excel Connector'],
    ['State', 'Enabled'],
    ['Token-Based Authentication (TBA)', 'Not ticked (not needed with OAuth 2.0)'],
    ['OAuth 2.0 > Authorization Code Grant', '**Ticked**'],
    ['OAuth 2.0 > Public Client', '**Ticked** (no secret is stored on the PCs)'],
    ['Redirect URI', '`https://localhost:3000/oauth/callback` (exactly)'],
    ['Scope > RESTlets', '**Ticked** (nothing else)'],
    ['Refresh token validity (hours)', '720 (30 days)'],
    ['Maximum time for token rotation (hours)', '720'],
  ], [4000, 5638], { firstCol: true }),
  P('Save. NetSuite shows the **Client ID** (and a client secret) once. Copy the **Client ID** for Part B.'),
  box('warn', 'The client secret', 'TQ Connector does not use the client secret. Do not copy it, store it or send it anywhere. If it was exposed, use Reset Credentials on the record later; the Client ID stays the same.'),

  H2('A6. Give the users access'),
  P('For each Excel user: **Lists > Employees > (employee) > Access > Roles**, add **TQ Excel User** next to their existing roles, Save. They keep their usual role for NetSuite itself; when they sign in from Excel they pick TQ Excel User (or any role with the same report permissions).'),
  table([
    ['To do this in Excel', 'The user needs'],
    ['See balances, reports, dashboards, drill down', 'TQ Excel User (or a role with Financial Statements, Trial Balance or General Ledger report permission)'],
    ['Upload budgets', 'Administrator, or a role with Budget permission at Edit or higher'],
    ['Upload journal entries', 'Administrator, or a role with Make Journal Entry at Edit or higher'],
    ['Release other users’ seats, manage the company report library', 'Administrator'],
    ['Data tables (saved searches, SuiteQL)', 'What their own role can see in NetSuite'],
  ], [4300, 5338]),

  // ------------------------------------------------------------------ B
  BR(),
  H1('Part B. Prepare the client’s files'),
  P('**Who:** Techno Q.   **Time:** about 10 minutes.'),
  H2('B1. Make the licence'),
  P('On the Techno Q PC that holds the licence key, in the repository folder:'),
  P('`python scripts\\make_licence.py --customer "Client Name" --account 1234567 --users 5 --expires 2027-12-31`'),
  P('This writes `licence-1234567.json`. Rename it to `licence.json` for the client. The licence is bound to that NetSuite account and cannot be edited. **users** is the number of people who may use the add-in at the same time (seats).'),
  H2('B2. Make the settings file'),
  P('Create `TQConnector.settings.json` (a plain text file) with the details from Part A:'),
  ...['{', '  "accountId": "1234567",', '  "restletUrl": "https://1234567.restlets.api.netsuite.com/app/site/hosting/restlet.nl?script=101&deploy=1",',
    '  "clientId": "the Client ID from A5",', '  "fiscalYearStartMonth": 1', '}'].map((l) => P('`' + l + '`', { para: { spacing: { after: 0 }, indent: { left: 360 } } })),
  spacer(),
  P('With this file next to the setup file, the installer skips its NetSuite and licence pages and copies both files in. Without it, the installer asks for the same details (Part D2).'),
  H2('B3. Put the files together'),
  P('One folder for the client: `TQConnectorSetup.exe`, `TQConnectorSetup.exe.sha256.txt`, `TechnoQ.cer` (all three from Techno Q’s build folder), `TQConnector.settings.json`, `licence.json`, this guide and the user guide. Share it through the client’s file share or Teams; avoid a web download link (Windows SmartScreen warns about downloaded programs).'),

  // ------------------------------------------------------------------ C
  BR(),
  H1('Part C. Client IT: once for the company'),
  P('**Who:** the client’s IT.   **Time:** about 15 minutes.'),
  H2('C1. Check the setup file is genuine'),
  P('In PowerShell, in the folder with the files:'),
  P('`Get-FileHash .\\TQConnectorSetup.exe -Algorithm SHA256`'),
  P('The result must match `TQConnectorSetup.exe.sha256.txt` and the checksum Techno Q gives separately (email or phone).'),
  H2('C2. Trust Techno Q’s certificate'),
  P('The setup file and the connector are signed by **Techno Q Code Signing** (thumbprint `F264951F57972A7C749F4163724D4B2B2F2DD4AB`, valid until November 2031). Once the certificate is trusted, Windows shows Techno Q as the verified publisher, and any tampered file shows as invalid. The certificate can only sign programs; it is not valid for websites or email.'),
  P('Check it first: `certutil -dump .\\TechnoQ.cer` must show that thumbprint. Then choose one way:'),
  table([
    ['Scope', 'How'],
    ['Whole company: Group Policy', 'Computer Configuration > Policies > Windows Settings > Security Settings > Public Key Policies: import TechnoQ.cer into **Trusted Root Certification Authorities** and into **Trusted Publishers**.'],
    ['Whole company: Intune', 'Devices > Configuration > Create > Templates > Trusted certificate: TechnoQ.cer to the computer store Root (and Trusted Publisher if offered).'],
    ['One PC, all users (administrator PowerShell)', '`certutil -addstore Root .\\TechnoQ.cer` then `certutil -addstore TrustedPublisher .\\TechnoQ.cer`'],
    ['One PC, one user (no administrator)', '`certutil -user -addstore Root .\\TechnoQ.cer` then `certutil -user -addstore TrustedPublisher .\\TechnoQ.cer`'],
  ], [3000, 6638], { firstCol: true }),
  P('Check: right-click TQConnectorSetup.exe > Properties > Digital Signatures > Details shows **This digital signature is OK**.'),
  H2('C3. Antivirus'),
  P('If the antivirus stops the setup file or TQConnector.exe, allow **Techno Q** as a trusted publisher (by its certificate), or allow the files by their SHA-256. The program lives in `%LOCALAPPDATA%\\TQ Connector` of each user.'),
  H2('C4. Silent installation (optional)'),
  P('IT can install for a user without any questions, while signed in **as that user** (not as SYSTEM, because it installs per user):'),
  P('`TQConnectorSetup.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART`'),
  P('with `TQConnector.settings.json` and `licence.json` in the same folder. Windows still asks the user once to trust the PC’s own localhost certificate (D2, step 4).'),

  // ------------------------------------------------------------------ D
  BR(),
  H1('Part D. Install on each PC'),
  P('**Who:** each user (or IT). Do the first PC together with Techno Q.   **Time:** about 5 minutes per PC.'),
  H2('D1. Before'),
  bullets(['Sign in to Windows as the person who will use Excel.', 'Close Excel (or restart it after the installation).', 'Have the folder from Part B open.']),
  H2('D2. Run the setup'),
  steps([
    'Double-click **TQConnectorSetup.exe**. Windows shows the publisher **Techno Q** (if IT trusted the certificate). Click **Yes** if Windows asks to allow it.',
    'Welcome page: **Next**.',
    'Only without the prepared files: the **Connect to NetSuite** page asks for the account ID, the RESTlet URL and the client ID, then the **Licence** page for the licence file. Both can be left empty and filled in later on the setup page.',
    'Windows asks: **"Do you want to install this certificate?"** for **localhost**. Click **Yes**. This is the PC’s own certificate for the connection between Excel and the connector; it is valid only for this PC.',
    'The setup finishes, starts the connector and offers to open the setup page. Leave it ticked and click **Finish**.',
  ]),
  box('note', 'What the setup did', ['Installed TQ Connector in %LOCALAPPDATA%\\TQ Connector (no administrator rights).', 'Registered the add-in with Excel, made and trusted this PC’s localhost certificate.', 'Started the connector, and set it to start at every Windows sign-in and to restart by itself if it ever stops.', 'Added a Start menu shortcut **TQ Connector**.']),

  H2('D3. Check the setup page and sign in'),
  P('The setup page opens in the browser (also reachable later from the panel’s gear menu > Setup).'),
  pictures2([img('setup-overview'), 'Setup > Overview: connection, signed-in user, licence, version'], [img('setup-connection'), 'Setup > NetSuite connection: Sign in with NetSuite'], 290),
  steps([
    'On **Overview**, check the account ID and the licence (users and expiry date).',
    'Open **NetSuite connection** and click **Sign in with NetSuite**.',
    'NetSuite’s own sign-in page opens: sign in (with the two-factor code if asked), choose the role **TQ Excel User**, and click **Allow**.',
    'Back on the setup page: **"Signed in to NetSuite as (your name)"**. The password never reaches TQ Connector.',
  ]),
  box('tip', 'Sign-in lasts 30 days', 'NetSuite keeps the user signed in for up to 30 days of use (the refresh token on the integration record). After that, or after Sign out, the panel shows a **Sign in with NetSuite** link.'),

  H2('D4. Open Excel'),
  steps([
    'Start Excel and open a blank workbook.',
    'The ribbon has a **TQ Connector** tab, and the Home tab a **TQ Connector** button. Click **Open panel**.',
    'The first time, the panel shows the welcome checks: connected, signed in, licensed.',
  ]),
  pictures2([img('panel-welcome'), 'The first time: the welcome checks'], [img('panel-home'), 'The panel’s Home screen'], 290),
  box('note', 'If the TQ Connector tab does not appear', 'Close Excel completely and start it again. Still missing: File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs must not block developer add-ins (company policy); ask IT.'),

  H2('D5. First test (with the first user)'),
  steps([
    'In the panel: gear menu > **Build a demo pack (3 sheets)**. It makes a monthly P&L, a balance sheet and an executive dashboard for last month.',
    'Compare two or three figures with NetSuite (Reports > Financial > Income Statement and Balance Sheet, same month and subsidiary).',
    'Right-click a number > **TQ Connector > Drill down**: the transactions behind it appear.',
    'Change the month with the **‹ ›** arrows on the report screen: every column rolls.',
    'Press **Refresh all**: the message line says how many numbers came from NetSuite and how long it took.',
  ]),

  // ------------------------------------------------------------------ E
  BR(),
  H1('Part E. Acceptance checklist'),
  P('Tick each line with the client before handing over.'),
  table([
    ['☐', 'Check', 'Result'],
    ['☐', 'Package deployed; RESTlet and Suitelet released; Suitelet runs as Administrator', ''],
    ['☐', 'Integration record: Authorization Code Grant, Public Client, RESTlets, redirect URI exact', ''],
    ['☐', 'Every Excel user has the TQ Excel User role (or equivalent)', ''],
    ['☐', 'Setup file checksum checked; TechnoQ.cer trusted; signature shows "OK"', ''],
    ['☐', 'Connector installed on each PC; TQ Connector tab visible in Excel', ''],
    ['☐', 'Each user signed in with their own NetSuite login', ''],
    ['☐', 'Licence shows the right client, account, users and expiry', ''],
    ['☐', 'Demo pack figures match NetSuite’s Income Statement and Balance Sheet', ''],
    ['☐', 'Drill down shows transactions and opens them in NetSuite', ''],
    ['☐', 'Budget / journal upload: Check works for the authorised users (if used)', ''],
    ['☐', 'Company logo on PDF pack (Setup > Company logo, or the NetSuite logo)', ''],
    ['☐', 'Users have the user guide; the administrator has this guide', ''],
  ], [500, 6600, 2538], { header: true }),
  spacer(),
  table([
    ['Sign-off', 'Name', 'Signature', 'Date'],
    ['Client', '', '', ''],
    ['Techno Q', '', '', ''],
  ], [1800, 3000, 3000, 1838]),

  // ------------------------------------------------------------------ F
  BR(),
  H1('Part F. After go-live'),
  H2('F1. Seats: who uses the licence'),
  P('Setup > **Seats** lists who holds a seat and when it was last used. A seat belongs to a person (whichever PC they use) and frees itself after 30 days without use. An administrator can **Release** a seat at once, for example when someone leaves.'),
  ...picture(img('setup-seats'), 300, 'Setup > Seats'),
  H2('F2. A new user'),
  steps(['Add the role TQ Excel User to their employee record (A6).', 'Install on their PC (Part D).', 'If all seats are taken: release an unused one, or ask Techno Q to extend the licence.']),
  H2('F3. A new or replaced PC'),
  P('Install on the new PC (Part D). To take the user’s setup along (saved filter sets, logo, the layouts of a workbook): on the old PC, panel gear menu > **Move to another PC** > **Export**; on the new PC, the same screen > **Import**. Passwords and keys are never in that file; the user signs in with NetSuite again on the new PC.'),
  ...picture(img('panel-transfer'), 300, 'Move to another PC: export and import'),
  H2('F4. Updates'),
  P('Techno Q sends a new setup file: run it on each PC (or silently, C4). Settings, licence, sign-in and seats are kept. When Techno Q gives an update address, enter it on Setup > **Updates**: the panel then offers new versions itself, and installs only versions signed by Techno Q.'),
  ...picture(img('setup-updates'), 300, 'Setup > Updates'),
  H2('F5. NetSuite changes'),
  bullets(['When Techno Q releases a new TQ script version, deploy the package again (A2). The panel shows a notice when the NetSuite script is older than the add-in needs.',
    'A new subsidiary, department or account needs nothing: the add-in reads them from NetSuite.',
    'Sandbox refreshed from production: deploy the package to the sandbox again and re-create its integration record if needed.']),
  H2('F6. Support'),
  P('If something does not work: Setup > **Save diagnostics**. It saves a small zip in **Documents\\TQ Connector** with versions, status, settings (every password, token and key removed) and the connector’s log. Email it to Techno Q support.'),
  H2('F7. Uninstall'),
  P('Windows Settings > Apps > **TQ Connector for NetSuite** > Uninstall. It stops the connector and removes the add-in, the shortcuts and this PC’s localhost certificate. The data folder (settings, licence) is kept for a later reinstall; delete `%LOCALAPPDATA%\\TQ Connector` to remove everything.'),

  // ------------------------------------------------------------------ G
  BR(),
  H1('Troubleshooting'),
  table([
    ['What you see', 'What to do'],
    ['Excel: "add-in error" or the panel stays blank', 'The connector is not running. Start menu > **TQ Connector**, wait 10 seconds, reopen the panel. It then keeps itself running.'],
    ['"TQ Connector is not running on this PC"', 'Same as above.'],
    ['No TQ Connector tab in Excel', 'Restart Excel. Check that the setup finished. Company policy may block developer add-ins: ask IT.'],
    ['Windows asks again to trust a localhost certificate', 'Normal about every two years: the certificate renews itself. Click **Yes**.'],
    ['"Not signed in"', 'Click **Sign in with NetSuite** (panel header or Setup > NetSuite connection).'],
    ['"Invalid Redirect URI" on the NetSuite page', 'The integration’s Redirect URI must be exactly `https://localhost:3000/oauth/callback` (A5).'],
    ['"None of your NetSuite roles may run financial reports"', 'Give the user the TQ Excel User role (A6) and sign in with it.'],
    ['"Deploy the elevated Suitelet with Execute As Role = Administrator"', 'Fix the Suitelet deployment (A3).'],
    ['"NetSuite has TQ script version N"', 'Deploy the TQ package again (A2).'],
    ['All seats are in use', 'Setup > Seats: release unused seats, or extend the licence.'],
    ['"Licence" error', 'The licence must be for this NetSuite account and not expired: Setup > Licence; ask Techno Q.'],
    ['Figures differ from NetSuite', 'Check the same month, subsidiary and filters on both sides; drill down on the figure; send Techno Q the diagnostics.'],
    ['Anything else', 'Setup > Save diagnostics, email the zip to Techno Q.'],
  ], [3800, 5838], { firstCol: true }),
];

c.save(c.doc('Installation Guide', body), OUT);

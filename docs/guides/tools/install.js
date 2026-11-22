// TQ Connector for NetSuite: Installation Guide (short version: bundle + account ID). node install.js <images dir> <output.docx>
const path = require('path');
const c = require('./common');
const { P, H1, H2, BR, bullets, steps, table, box, picture, pictures2, spacer } = c;
const IMG = process.argv[2];
const OUT = process.argv[3];
const img = (n) => path.join(IMG, n + '.png');

const BUNDLE = '715956';

const body = [
  ...c.cover('Installation Guide', 'A new client in three steps: install the bundle, run the setup, sign in',
    ['Version 1.5  ·  November 2026', 'For: the client’s NetSuite administrator, IT and users',
      'NetSuite bundle ' + BUNDLE + '  ·  Connector 1.5.3']),
  ...c.toc(),

  // ------------------------------------------------------------------ 1
  H1('1. Overview'),
  box('key', 'The whole installation', [
    '**Administrator, once (about 10 minutes):** install the TQ bundle in NetSuite and give users the TQ Excel User role.',
    '**Each PC (about 3 minutes):** run TQConnectorSetup.exe, type the NetSuite account ID, sign in with NetSuite.',
    'Nothing else: no scripts to deploy, no integration record to create, no keys or addresses to copy.',
  ]),
  spacer(),
  table([
    ['Step', 'What', 'Who', 'Time'],
    ['1', 'Install bundle ' + BUNDLE + ' and assign the role', 'Client’s NetSuite administrator', '10 min'],
    ['2', 'Run the setup on each PC, type the account ID', 'Each user (or IT)', '3 min per PC'],
    ['3', 'Sign in with NetSuite (the administrator first: this stores the licence)', 'Each user', '1 min'],
  ], [700, 5100, 2600, 1238]),

  H2('1.1 What happens automatically'),
  table([
    ['Before (version 1.4)', 'Now (version 1.5)'],
    ['Deploy three scripts with the SuiteCloud CLI', 'The bundle installs the scripts, their deployments and the role'],
    ['Create an integration record, copy its Client ID', 'Techno Q’s sign-in integration comes with the bundle (NetSuite adds it at the first sign-in too)'],
    ['Copy the RESTlet address', 'Worked out from the account ID'],
    ['Type the first month of the fiscal year', 'Read from NetSuite’s accounting periods'],
    ['Copy the licence file to every PC', 'Stored once in NetSuite; every PC fetches it when the user signs in'],
    ['Prepare a settings file for every PC', 'Only the account ID, typed in the setup'],
  ], [4800, 4838], { header: true }),

  H2('1.2 What Techno Q sends'),
  table([
    ['File', 'What it is'],
    ['`TQConnectorSetup.exe`', 'The setup, run on each PC. Signed by Techno Q.'],
    ['`licence.json`', 'The licence for the client’s NetSuite account (users and expiry date). Used once, on the administrator’s PC.'],
    ['`TechnoQ.cer` (optional)', 'Techno Q’s code-signing certificate, if IT wants Windows to show Techno Q as the verified publisher (section 5).'],
    ['This guide and the user guide', ''],
  ], [3300, 6338]),
  P('Techno Q needs the client’s NetSuite **account ID** (production, and sandbox if used) to make the licence and to make the bundle available to that account.'),

  H2('1.3 Requirements'),
  table([
    ['Item', 'Requirement'],
    ['PCs', 'Windows 10 or 11, 64-bit. No administrator rights needed.'],
    ['Excel', 'Microsoft 365 (desktop), Excel 2021 or 2024 for Windows.'],
    ['NetSuite', 'Any edition. OneWorld, multi-currency and multi-book are supported, not required.'],
    ['Network', 'The PCs reach NetSuite over HTTPS. Nothing else is opened.'],
  ], [2400, 7238], { firstCol: true }),

  // ------------------------------------------------------------------ 2
  BR(),
  H1('2. NetSuite administrator: once'),
  H2('2.1 Check two features'),
  P('**Setup > Company > Enable Features > SuiteCloud**. Both are usually already ticked; tick them if not, then Save:'),
  table([
    ['Feature', 'Section', 'Why'],
    ['Server SuiteScript', 'SuiteScript', 'The bundle’s scripts run in NetSuite'],
    ['OAuth 2.0', 'Manage Authentication', 'Users sign in with their own NetSuite login'],
  ], [3000, 3000, 3638]),

  H2('2.2 Install the TQ bundle'),
  steps([
    '**Customization > SuiteBundler > Search & Install Bundles**.',
    'In **Keywords** type `' + BUNDLE + '` and click **Search**.',
    'Click **TQ Excel Connector**, then **Install**.',
    'On the preview page click **Install Bundle**, then **OK**.',
    'Wait until **Customization > SuiteBundler > Search & Install Bundles > List** shows the bundle with status **Installed** (a minute or two).',
  ]),
  box('note', 'What the bundle installs', [
    'The scripts **TQ Excel Connector RESTlet** and **TQ Excel Connector elevated**, released, with the right settings (the elevated one runs as Administrator and is not available without login).',
    'The role **TQ Excel User**.',
    'Techno Q’s sign-in integration **TQ Excel Connector** (no secret is stored on the PCs).',
    'The script code is hidden and locked; Techno Q releases updates through the same bundle.',
  ]),

  H2('2.3 Give users the role'),
  P('For each Excel user: **Lists > Employees > (employee) > Access > Roles**, add **TQ Excel User**, Save. They keep their usual roles; when they sign in from Excel they choose TQ Excel User (or any role with the same report permissions).'),
  table([
    ['To do this in Excel', 'The user needs'],
    ['Balances, reports, dashboards, drill down', 'TQ Excel User (or Financial Statements / Trial Balance / General Ledger report permission)'],
    ['Upload budgets', 'Administrator, or Budget permission at Edit or higher'],
    ['Upload journal entries', 'Administrator, or Make Journal Entry at Edit or higher'],
    ['Release seats, manage the company report library', 'Administrator'],
  ], [4300, 5338]),

  // ------------------------------------------------------------------ 3
  BR(),
  H1('3. Each PC'),
  H2('3.1 Run the setup'),
  steps([
    'Sign in to Windows as the person who will use Excel, and close Excel.',
    'Double-click **TQConnectorSetup.exe**. If Windows SmartScreen asks, click **More info > Run anyway** (or see section 5 to avoid this).',
    'Type the **NetSuite account ID** (for example `1234567`, or `1234567_SB1` for a sandbox). A NetSuite address pasted here works too: the setup keeps only the account ID.',
    'Licence page: on the **administrator’s** PC choose `licence.json`; on every other PC leave it empty and click **Next**.',
    'Windows asks **"Do you want to install this certificate?"** for **localhost**: click **Yes**. It is this PC’s own certificate for Excel and the connector.',
    'Click **Finish**. The setup page opens in the browser.',
  ]),
  box('note', 'What the setup did', ['Installed TQ Connector in %LOCALAPPDATA%\\TQ Connector, registered the Excel add-in and started the connector.', 'The connector starts at every Windows sign-in and restarts itself if it ever stops.']),

  H2('3.2 Sign in with NetSuite'),
  pictures2([img('setup-connection'), 'Setup > NetSuite connection'], [img('setup-overview'), 'Setup > Overview after signing in'], 290),
  steps([
    'On the setup page, click **Sign in with NetSuite**.',
    'NetSuite’s own page opens: sign in, choose the role **TQ Excel User**, click **Allow**.',
    'Back on the setup page: **"Signed in to NetSuite as (your name)"**, with the licence shown.',
  ]),
  box('key', 'The administrator signs in first', 'When the administrator signs in with the licence chosen in 3.1, the licence is stored in NetSuite. Every other user’s PC then fetches it at sign-in: nothing to copy. A licence can also be pasted later on Setup > Licence (administrator).'),
  box('tip', 'Sign-in lasts 30 days', 'NetSuite keeps the user signed in for up to 30 days of use. After that the panel shows a **Sign in with NetSuite** link.'),

  H2('3.3 Open Excel'),
  steps([
    'Start Excel, open a blank workbook.',
    'On the **TQ Connector** tab click **Open panel**: the welcome checks show connected, signed in, licensed.',
    'First test: gear menu > **Build a demo pack**, and compare two or three figures with NetSuite’s Income Statement and Balance Sheet.',
  ]),
  pictures2([img('panel-welcome'), 'The first time: the welcome checks'], [img('panel-home'), 'The panel’s Home screen'], 290),

  // ------------------------------------------------------------------ 4
  BR(),
  H1('4. Acceptance checklist'),
  table([
    ['☐', 'Check', 'Result'],
    ['☐', 'Bundle ' + BUNDLE + ' shows Installed in NetSuite', ''],
    ['☐', 'Every Excel user has the TQ Excel User role', ''],
    ['☐', 'TQ Connector installed on each PC; TQ Connector tab in Excel', ''],
    ['☐', 'Each user signed in with their own NetSuite login', ''],
    ['☐', 'Licence shows the right client, users and expiry on every PC', ''],
    ['☐', 'Demo pack figures match NetSuite’s Income Statement and Balance Sheet', ''],
    ['☐', 'Drill down shows transactions and opens them in NetSuite', ''],
  ], [500, 6600, 2538], { header: true }),
  spacer(),
  table([
    ['Sign-off', 'Name', 'Signature', 'Date'],
    ['Client', '', '', ''],
    ['Techno Q', '', '', ''],
  ], [1800, 3000, 3000, 1838]),

  // ------------------------------------------------------------------ 5
  BR(),
  H1('5. For IT (optional)'),
  H2('5.1 Show Techno Q as the verified publisher'),
  P('Trusting `TechnoQ.cer` once removes the SmartScreen question and makes Windows show **Techno Q** as publisher. Check the thumbprint first (`certutil -dump .\\TechnoQ.cer` shows `F264951F57972A7C749F4163724D4B2B2F2DD4AB`), then:'),
  table([
    ['Scope', 'How'],
    ['Whole company', 'Group Policy or Intune: import TechnoQ.cer into **Trusted Root Certification Authorities** and **Trusted Publishers**.'],
    ['One PC (administrator PowerShell)', '`certutil -addstore Root .\\TechnoQ.cer` and `certutil -addstore TrustedPublisher .\\TechnoQ.cer`'],
  ], [3000, 6638], { firstCol: true }),
  H2('5.2 Silent installation'),
  P('Signed in as the user: `TQConnectorSetup.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /ACCOUNT=1234567`. The user then only signs in. On the administrator’s PC add `/LICENCE="C:\\path\\licence.json"`.'),
  H2('5.3 Antivirus'),
  P('If the antivirus stops the setup or TQConnector.exe, allow Techno Q as a trusted publisher. The program lives in `%LOCALAPPDATA%\\TQ Connector`.'),

  // ------------------------------------------------------------------ 6
  H1('6. After go-live'),
  table([
    ['Task', 'How'],
    ['New user', 'Add the TQ Excel User role, run the setup on their PC, sign in.'],
    ['New or replaced PC', 'Run the setup and sign in. To take a user’s filter sets and layouts along: panel gear menu > **Move to another PC** > Export, then Import on the new PC.'],
    ['Seats', 'Setup > **Seats** shows who uses the licence. A seat frees itself after 30 days unused; an administrator can release one at once.'],
    ['Licence renewal', 'From 30 days before the end date the Excel panel shows a reminder. Techno Q sends a new licence.json: the administrator pastes it on Setup > Licence. NetSuite’s copy is replaced at once; every other PC takes it within 6 hours.'],
    ['NetSuite script update', 'Techno Q pushes it through the bundle; nothing to do. The panel says if NetSuite’s script is older than the add-in needs.'],
    ['Connector update', 'Run the new setup file, or accept the update the panel offers. Settings, licence and sign-in are kept.'],
    ['Sandbox refresh', 'Install bundle ' + BUNDLE + ' in the sandbox again if it is missing after the refresh.'],
    ['Support', 'Setup > **Save diagnostics** saves a zip (no passwords or keys) in Documents\\TQ Connector: email it to Techno Q.'],
    ['Uninstall', 'Windows Settings > Apps > **TQ Connector for NetSuite** > Uninstall.'],
  ], [2600, 7038], { firstCol: true }),

  H1('7. Troubleshooting'),
  table([
    ['What you see', 'What to do'],
    ['Excel panel blank, or "TQ Connector is not running"', 'Start menu > **TQ Connector**, wait 10 seconds, reopen the panel.'],
    ['No TQ Connector tab in Excel', 'Restart Excel. Company policy may block add-ins: ask IT.'],
    ['"Not signed in"', 'Click **Sign in with NetSuite**.'],
    ['NetSuite says the integration is not found or disabled', 'Check OAuth 2.0 is ticked (2.1) and the bundle is installed (2.2).'],
    ['"None of your NetSuite roles may run financial reports"', 'Give the user the TQ Excel User role and sign in with it.'],
    ['"No licence" on a user’s PC', 'The administrator signs in once with the licence (3.2), or pastes it on Setup > Licence.'],
    ['All seats are in use', 'Setup > Seats: release unused seats, or ask Techno Q to extend the licence.'],
    ['Figures differ from NetSuite', 'Check the same month, subsidiary and filters; drill down; send Techno Q the diagnostics.'],
  ], [3800, 5838], { firstCol: true }),
];

c.save(c.doc('Installation Guide', body), OUT);

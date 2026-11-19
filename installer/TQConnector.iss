; TQ Connector for NetSuite - the setup file a customer runs on each PC (Inno Setup 6).
; Built by installer\build.ps1, which passes /DAppVersion=... and /DStage=... (the staged app folder).
;
; What it does, for the signed-in Windows user only (no administrator rights):
;   - stops a running connector and replaces its files in %LOCALAPPDATA%\TQ Connector; the data folder
;     (settings, licence, this PC's certificate and seat) is kept
;   - makes this PC's certificate for https://localhost the first time and asks Windows to trust it
;   - registers the add-in with Excel and starts the connector at every sign-in (and again if it ever stops)
;   - starts the connector; on a first install it offers to open the setup page
; Uninstall (Windows Settings > Apps) undoes all of it and keeps the data folder.
; To sign the setup file, build.ps1 adds SignTool=... once Techno Q's code-signing certificate is set up.

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#ifndef Stage
  #define Stage "..\build\stage\app"
#endif
#define AddinId "6f1c2a9e-3b7d-4e58-9a41-c2d07b5e8f13"

[Setup]
AppId={{A7C3E1D2-5B64-4F0A-9C21-7E8D3B9F4A10}
AppName=TQ Connector for NetSuite
AppVersion={#AppVersion}
AppVerName=TQ Connector for NetSuite {#AppVersion}
AppPublisher=Techno Q
AppPublisherURL=https://www.technoq.com
VersionInfoVersion={#AppVersion}
VersionInfoCompany=Techno Q
VersionInfoDescription=TQ Connector for NetSuite setup
DefaultDirName={localappdata}\TQ Connector
DisableDirPage=yes
DisableProgramGroupPage=yes
DisableReadyPage=yes
PrivilegesRequired=lowest
OutputDir=..\build
OutputBaseFilename=TQConnectorSetup
SetupIconFile=..\addin\assets\tq.ico
UninstallDisplayIcon={app}\TQConnector.exe
UninstallDisplayName=TQ Connector for NetSuite
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=no
RestartApplications=no
#ifdef SignTool
SignTool={#SignTool}
SignedUninstaller=yes
#endif

[Messages]
WelcomeLabel2=This installs TQ Connector for NetSuite [ver] for you on this PC.%n%nIt runs a small local server that Excel's TQ Connector add-in talks to. Excel can stay open, but restart it afterwards to see the add-in.

[Files]
Source: "{#Stage}\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Dirs]
Name: "{app}\data"; Flags: uninsneveruninstall

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\Developer"; ValueType: string; ValueName: "{#AddinId}"; ValueData: "{app}\addin\manifest.xml"; Flags: uninsdeletevalue

[Icons]
; --watch: the connector is started again by itself if it ever stops, so Excel's add-in never stays broken
Name: "{userstartup}\TQ Connector"; Filename: "{app}\TQConnector.exe"; Parameters: "--watch"; WorkingDir: "{app}"; Comment: "TQ Connector for NetSuite (local server for Excel)"
Name: "{userprograms}\TQ Connector"; Filename: "{app}\TQConnector.exe"; Parameters: "--watch"; WorkingDir: "{app}"; Comment: "Start TQ Connector for NetSuite (if Excel says the add-in cannot load)"

[Run]
Filename: "https://localhost:3000/setup.html"; Description: "Open the setup page (NetSuite details and licence)"; Flags: shellexec postinstall skipifsilent nowait; Check: FirstInstall

[UninstallDelete]
Type: filesandordirs; Name: "{app}\_internal"
Type: filesandordirs; Name: "{app}\addin"

[Code]
var
  WasSetUp: Boolean;
  InfoPage: TInputQueryWizardPage;
  LicencePage: TInputFileWizardPage;

// A prepared settings file next to the setup file (for rolling out to several PCs): TQConnector.settings.json is
// copied to the data folder as config.json, and licence.json as the licence. Then nobody types anything.
function PreparedSettings(): String;
begin
  Result := ExpandConstant('{src}\TQConnector.settings.json');
end;

function JsonText(Value: String): String;
begin
  Result := Trim(Value);
  StringChangeEx(Result, '\', '\\', True);
  StringChangeEx(Result, '"', '\"', True);
end;

procedure InitializeWizard();
begin
  InfoPage := CreateInputQueryPage(wpWelcome, 'Connect to NetSuite',
    'Your NetSuite details (you can also enter or change them later on the setup page).',
    'From your NetSuite administrator: the account ID, the TQ Excel Connector RESTlet''s External URL, and the client ID of ' +
    'the TQ Excel Connector (sign-in) integration. Leave them empty to enter them later.');
  InfoPage.Add('Account ID (e.g. 1234567 or 1234567_SB1):', False);
  InfoPage.Add('RESTlet External URL:', False);
  InfoPage.Add('Client ID (Sign in with NetSuite):', False);
  LicencePage := CreateInputFilePage(InfoPage.ID, 'Licence', 'The licence file from Techno Q.',
    'Choose the licence file you received (you can also paste it later on the setup page).');
  LicencePage.Add('Licence file:', 'Licence files (*.json)|*.json|All files|*.*', '.json');
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := False;
  if (PageID = InfoPage.ID) or (PageID = LicencePage.ID) then
    Result := FileExists(ExpandConstant('{localappdata}\TQ Connector\data\config.json')) or FileExists(PreparedSettings());
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  Url: String;
begin
  Result := True;
  if CurPageID = InfoPage.ID then begin
    Url := Lowercase(Trim(InfoPage.Values[1]));
    if (Url <> '') and ((Pos('https://', Url) <> 1) or (Pos('.restlets.api.netsuite.com', Url) = 0)) then begin
      MsgBox('The RESTlet URL should look like https://1234567.restlets.api.netsuite.com/app/site/hosting/restlet.nl?script=...', mbError, MB_OK);
      Result := False;
    end;
  end;
end;

// Writes the details from the installer pages (or the prepared file) into the data folder, on a first install only.
procedure SaveSettings();
var
  Data, Config, Licence: String;
begin
  Data := ExpandConstant('{app}\data');
  Config := Data + '\config.json';
  ForceDirectories(Data);
  if not FileExists(Config) then begin
    if FileExists(PreparedSettings()) then
      FileCopy(PreparedSettings(), Config, False)
    else if Trim(InfoPage.Values[0]) <> '' then
      SaveStringToFile(Config, '{"accountId": "' + JsonText(InfoPage.Values[0]) + '", "restletUrl": "' +
        JsonText(InfoPage.Values[1]) + '", "clientId": "' + JsonText(InfoPage.Values[2]) +
        '", "fiscalYearStartMonth": 1}', False);
  end;
  Licence := Data + '\licence.json';
  if not FileExists(Licence) then begin
    if FileExists(ExpandConstant('{src}\licence.json')) then
      FileCopy(ExpandConstant('{src}\licence.json'), Licence, False)
    else if (LicencePage.Values[0] <> '') and FileExists(LicencePage.Values[0]) then
      FileCopy(LicencePage.Values[0], Licence, False);
  end;
end;

function CertFile(): String;
begin
  Result := ExpandConstant('{app}\data\certs\localhost.crt');
end;

procedure StopConnector();
var
  Code: Integer;
begin
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM TQConnector.exe', '', SW_HIDE, ewWaitUntilTerminated, Code);
  Sleep(700);
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  WasSetUp := FileExists(ExpandConstant('{app}\data\config.json'));
  StopConnector();
  // files of an older version (the data folder stays)
  DelTree(ExpandConstant('{app}\_internal'), True, True, True);
  DelTree(ExpandConstant('{app}\addin'), True, True, True);
  // the uninstall entry of the earlier, script-based installer: this one has its own
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\TQConnector');
  Result := '';
end;

function FirstInstall(): Boolean;
begin
  Result := not WasSetUp;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  Code: Integer;
  Marker: String;
begin
  if CurStep <> ssPostInstall then
    exit;
  SaveSettings();
  if not FileExists(CertFile()) then begin
    WizardForm.StatusLabel.Caption := 'Making this PC''s certificate for https://localhost...';
    Exec(ExpandConstant('{app}\TQConnector.exe'), '--make-cert', ExpandConstant('{app}'), SW_HIDE, ewWaitUntilTerminated, Code);
  end;
  Marker := CertFile() + '.trusted';
  if FileExists(CertFile()) and not FileExists(Marker) then begin
    WizardForm.StatusLabel.Caption := 'Trusting the certificate for this Windows user (answer Yes in the Windows window)...';
    if Exec(ExpandConstant('{sys}\certutil.exe'), '-user -addstore Root "' + CertFile() + '"', '', SW_HIDE,
            ewWaitUntilTerminated, Code) and (Code = 0) then
      SaveStringToFile(Marker, 'yes', False)
    else
      MsgBox('The certificate was not trusted, so Excel cannot load the add-in yet. Run the setup again and answer Yes in the Windows window.', mbError, MB_OK);
  end;
  WizardForm.StatusLabel.Caption := 'Starting the connector...';
  Exec(ExpandConstant('{app}\TQConnector.exe'), '--watch', ExpandConstant('{app}'), SW_HIDE, ewNoWait, Code);
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  Code: Integer;
  Thumbprint: AnsiString;
begin
  if CurUninstallStep = usUninstall then begin
    StopConnector();
    // exactly this PC's certificate, by its thumbprint (written by the connector when it made it)
    if LoadStringFromFile(CertFile() + '.sha1', Thumbprint) and (Length(Trim(String(Thumbprint))) = 40) then
      Exec(ExpandConstant('{sys}\certutil.exe'), '-user -delstore Root ' + Trim(String(Thumbprint)), '', SW_HIDE,
           ewWaitUntilTerminated, Code);
    DeleteFile(CertFile() + '.trusted');
  end;
end;

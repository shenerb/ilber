[Setup]
AppId={{8F9E2B4A-5C1D-4E8B-9A3C-7D6E5F4A3B2C}
AppName=ILBERTech Platform v2.0
AppVersion=2.0.0
AppPublisher=ILBERTech AI
DefaultDirName={autopf}\ILBERTech_Platform
DefaultGroupName=ILBERTech Platform
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
SetupIconFile=C:\Users\lenovo\Desktop\ilber\ilber_logo.ico
UninstallDisplayIcon={app}\ilber_logo.ico
OutputDir=C:\Users\lenovo\Desktop\ilber
OutputBaseFilename=ILBERTech_Platform_Setup_v2
Compression=lzma2/fast
SolidCompression=no
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "turkish"; MessagesFile: "compiler:Languages\Turkish.isl"
Name: "german"; MessagesFile: "compiler:Languages\German.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\ILBERTech_Platform.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\lenovo\Desktop\ilber\dist\ILBERTech_Platform\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\ILBERTech Platform v2.0"; Filename: "{app}\ILBERTech_Platform.exe"; IconFilename: "{app}\ilber_logo.ico"
Name: "{autodesktop}\ILBERTech Platform v2.0"; Filename: "{app}\ILBERTech_Platform.exe"; IconFilename: "{app}\ilber_logo.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\ILBERTech_Platform.exe"; Description: "{cm:LaunchProgram,ILBERTech Platform v2.0}"; Flags: nowait postinstall skipifsilent

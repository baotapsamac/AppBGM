[Setup]
AppName=AppBGM
AppVersion=1.4
DefaultDirName={autopf}\AppBGM
DefaultGroupName=AppBGM
OutputBaseFilename=AppBGM_v1.4_Setup_x64
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
SetupIconFile=app_icon.ico

[Tasks]
Name: "desktopicon"; Description: "Tạo biểu tượng ngoài màn hình Desktop"; GroupDescription: "Additional icons:"

[Files]
Source: "dist\app\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "LibreOffice\*"; DestDir: "{app}\LibreOffice"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\AppBGM"; Filename: "{app}\app.exe"
Name: "{autodesktop}\AppBGM"; Filename: "{app}\app.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\app.exe"; Description: "Khởi chạy AppBGM"; Flags: nowait postinstall skipifsilent

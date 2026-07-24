[Setup]
AppName=AppBGM Editor
AppVersion=1.5
DefaultDirName={autopf}\AppBGM
DefaultGroupName=AppBGM
OutputBaseFilename=AppBGM_Editor_v1.5_Setup_x64
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
Name: "{autoprograms}\AppBGM Editor"; Filename: "{app}\app.exe"
Name: "{autodesktop}\AppBGM Editor"; Filename: "{app}\app.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\app.exe"; Description: "Khởi chạy AppBGM Editor"; Flags: nowait postinstall skipifsilent

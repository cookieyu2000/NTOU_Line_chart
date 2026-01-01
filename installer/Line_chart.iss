[Setup]
AppName=NTOU_Tools
AppVersion=1.0
AppPublisher=NTOU_Tools
DefaultDirName={autopf}\NTOU_Tools
DefaultGroupName=NTOU_Tools
OutputBaseFilename=NTOU_Tools_Setup
OutputDir=dist\installer
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
SetupIconFile=..\build\icons\NTOU_Tools.ico

[Files]
Source: "..\dist\NTOU_Tools.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\NTOU_Tools"; Filename: "{app}\NTOU_Tools.exe"
Name: "{commondesktop}\NTOU_Tools"; Filename: "{app}\NTOU_Tools.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop icon"; GroupDescription: "Additional icons:"

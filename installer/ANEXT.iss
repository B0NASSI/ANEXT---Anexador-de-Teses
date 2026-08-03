#define MyAppName "ANEXT — Anexador de Teses"
#define MyAppVersion "3.4"
#define MyAppPublisher "Rodriguez & Sousa Advogados Associados"
#define MyAppExeName "ANEXT Launcher.exe"
#define SourceDir "C:\Users\Pichau\Desktop\Petição Inicial\CODE\CLAUDE CODE\ANEXT\ANEXT - Instalador"

[Setup]
AppId={{6F2C9E1A-8B3D-4E7F-9A2C-5D1B6F0A3C8E}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppVerName={#MyAppName} {#MyAppVersion}
DefaultDirName={localappdata}\ANEXT
DefaultGroupName=ANEXT
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=output
OutputBaseFilename=ANEXT Setup v3.4
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName={#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}
DisableWelcomePage=no

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar um atalho na área de trabalho"; GroupDescription: "Atalhos adicionais:"

[Files]
Source: "{#SourceDir}\ANEXT.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\_internal\*"; DestDir: "{app}\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#SourceDir}\ANEXT Launcher.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\_internal_launcher\*"; DestDir: "{app}\_internal_launcher"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#SourceDir}\versao.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autodesktop}\ANEXT"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir o ANEXT agora"; Flags: nowait postinstall skipifsilent

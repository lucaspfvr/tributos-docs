[Setup]
AppId=TributosDocs.Biquinhas
AppName=Tributos Docs
AppVersion=0.2.0
AppPublisher=Setor de Tributos e Arrecadação
DefaultDirName={localappdata}\Programs\Tributos Docs
DefaultGroupName=Tributos Docs
PrivilegesRequired=lowest
OutputDir=..\release
OutputBaseFilename=Tributos-Docs-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\TributosDocs.exe

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Files]
Source: "..\dist\TributosDocs\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\Tributos Docs"; Filename: "{app}\TributosDocs.exe"

[Run]
Filename: "{app}\TributosDocs.exe"; Description: "Abrir Tributos Docs"; Flags: nowait postinstall skipifsilent

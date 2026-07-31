[Setup]
AppName=Green Email Verifier
AppVersion=1.0
AppPublisher=Green Email Verifier
AppPublisherURL=https://verificador-emails.vercel.app
AppSupportURL=https://verificador-emails.vercel.app
AppUpdatesURL=https://verificador-emails.vercel.app
DefaultDirName={autopf}\GreenEmailVerifier
DefaultGroupName=Green Email Verifier
OutputDir=Output
OutputBaseFilename=VerificadorEmails_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\VerificadorEmails.exe
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear acceso directo en el escritorio"; GroupDescription: "Iconos adicionales:"; Flags: unchecked
Name: "quicklaunchicon"; Description: "Crear acceso directo en el menú inicio"; GroupDescription: "Iconos adicionales:"

[Files]
Source: "dist\VerificadorEmails.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Green Email Verifier"; Filename: "{app}\VerificadorEmails.exe"
Name: "{group}\Desinstalar Green Email Verifier"; Filename: "{uninstallexe}"
Name: "{commondesktop}\Green Email Verifier"; Filename: "{app}\VerificadorEmails.exe"; Tasks: desktopicon
Name: "{commonstartmenu}\Green Email Verifier"; Filename: "{app}\VerificadorEmails.exe"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\VerificadorEmails.exe"; Description: "Ejecutar Green Email Verifier ahora"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

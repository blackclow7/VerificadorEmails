[Setup]
AppName=Correo Certificado
AppVersion=1.0
AppPublisher=Correo Certificado
AppPublisherURL=https://verificador-emails.vercel.app
AppSupportURL=https://verificador-emails.vercel.app
AppUpdatesURL=https://verificador-emails.vercel.app
DefaultDirName={autopf}\CorreoCertificado
DefaultGroupName=Correo Certificado
OutputDir=Output
OutputBaseFilename=VerificadorEmails_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile=
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
Name: "{group}\Correo Certificado"; Filename: "{app}\VerificadorEmails.exe"
Name: "{group}\Desinstalar Correo Certificado"; Filename: "{uninstallexe}"
Name: "{commondesktop}\Correo Certificado"; Filename: "{app}\VerificadorEmails.exe"; Tasks: desktopicon
Name: "{commonstartmenu}\Correo Certificado"; Filename: "{app}\VerificadorEmails.exe"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\VerificadorEmails.exe"; Description: "Ejecutar Correo Certificado ahora"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

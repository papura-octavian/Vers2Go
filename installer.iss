; Script Inno Setup pentru installerul Vers2Go.
; Compilare locală:  ISCC.exe installer.iss
; În CI (GitHub Actions), versiunea vine din tag: ISCC.exe /DMyAppVersion=1.2.3 installer.iss

#ifndef MyAppVersion
  #define MyAppVersion "1.0.0"
#endif

[Setup]
; AppId identifică unic aplicația pentru upgrade-uri/dezinstalare — nu-l schimba între versiuni
AppId={{B7E6A1D4-3C52-4F8E-9A17-D2E84C0F5B29}
AppName=Vers2Go
AppVersion={#MyAppVersion}
AppPublisher=SkipIt Labs
AppPublisherURL=https://github.com/SkipIt-Labs/Vers2Go
DefaultDirName={autopf}\Vers2Go
DisableProgramGroupPage=yes
OutputDir=installer_output
OutputBaseFilename=Vers2Go-Setup-{#MyAppVersion}
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\Vers2Go.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
; Instalare per-utilizator (fără drepturi de administrator); utilizatorul
; poate alege instalarea pentru tot sistemul din dialog dacă are drepturi
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "dist\Vers2Go.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\Vers2Go"; Filename: "{app}\Vers2Go.exe"
Name: "{autodesktop}\Vers2Go"; Filename: "{app}\Vers2Go.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Vers2Go.exe"; Description: "{cm:LaunchProgram,Vers2Go}"; Flags: nowait postinstall skipifsilent

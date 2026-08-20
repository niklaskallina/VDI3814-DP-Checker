; Inno-Setup-Beschreibung fuer den Windows-Installer.
; Erzeugt eine einzelne Setup-Datei, die alles Noetige mitbringt:
; Programm, Python-Laufzeit und die Texterkennung fuer Scans.
; Es wird nichts nachinstalliert und keine Administratorrechte benoetigt.

#define AppName "VDI 3814 DP-Checker"
#define AppExe "VDI3814-DP-Checker.exe"
#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif

[Setup]
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=VDI 3814 DP-Checker
DefaultDirName={localappdata}\VDI3814-DP-Checker
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\Output
OutputBaseFilename=VDI3814-DP-Checker-Setup
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayName={#AppName}
UninstallDisplayIcon={app}\{#AppExe}

[Languages]
Name: "deutsch"; MessagesFile: "compiler:Languages\German.isl"

[Tasks]
Name: "desktopicon"; Description: "Verknüpfung auf dem Desktop anlegen"; GroupDescription: "Zusätzliche Verknüpfungen"

[Files]
Source: "..\dist\VDI3814-DP-Checker\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{group}\{#AppName} deinstallieren"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "{#AppName} jetzt starten"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}\out"

; PDFree Inno Setup Script
; Build with Inno Setup 6: https://jrsoftware.org/isinfo.php
; Output: PDFree_Setup.exe

#define AppName "PDFree"
#define AppVersion "1.0.0"
#define AppPublisher "Fioerd"
#define AppURL "https://github.com/Fioerd/PDFree"
#define AppExeName "PDFree.exe"

; ============================================================
; STEP 1 — Digital signature + UAC
; ============================================================
; PrivilegesRequired=admin triggers a UAC prompt so the installer
; can write to Program Files and the system registry.
; To code-sign the output Setup .exe (removes SmartScreen warning):
;   signtool sign /fd sha256 /tr http://timestamp.digicert.com /td sha256 /f cert.pfx PDFree_Setup.exe
[Setup]
AppId={{A3F2C1D4-8E7B-4F9A-BC23-1D5E6F7A8B9C}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}/issues
AppUpdatesURL={#AppURL}/releases
PrivilegesRequired=admin
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes
LicenseFile=LICENSE
OutputDir=dist
OutputBaseFilename=PDFree_Setup
SetupIconFile=LOGO.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
UsedUserAreasWarning=no
; Minimum Windows version: Windows 10 (6.2 = Windows 8, 10.0 = Windows 10)
MinVersion=10.0

; ============================================================
; STEP 3 — Setup wizard pages
; ============================================================
; Inno Setup shows: Welcome → License → Directory → Components
; → Tasks (shortcuts) → Ready → Installing → Finish
; The pages below are shown automatically based on the sections defined.

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

; ============================================================
; STEP 7 — Dependencies / runtimes
; ============================================================
; PyInstaller bundles Python and all pip packages, so no separate
; Python runtime is needed. The Microsoft Visual C++ Redistributable
; DLLs are also bundled inside _internal/. Nothing extra to install.
; If a future version requires an external runtime, add it here:
;   [Run]
;   Filename: "vc_redist.x64.exe"; Parameters: "/quiet /norestart"; ...

; ============================================================
; STEP 3 — Components (optional feature selection in wizard)
; ============================================================
[Components]
Name: "main";        Description: "PDFree Application (required)"; Types: full compact custom; Flags: fixed
Name: "desktopicon"; Description: "Desktop shortcut";              Types: full

; ============================================================
; STEP 3 — Tasks (user choices shown in wizard)
; ============================================================
[Tasks]
Name: "desktopicon";    Description: "Create a &desktop shortcut";              GroupDescription: "Additional shortcuts:"; Components: desktopicon
Name: "fileassoc";      Description: "Open .pdf files with PDFree by default";  GroupDescription: "File associations:";    Flags: unchecked

; ============================================================
; STEP 4 — Copy files to install directory
; ============================================================
; The entire dist\PDFree\ folder (exe + _internal\) is written
; to {autopf}\PDFree\ (e.g. C:\Program Files\PDFree\)
[Files]
Source: "dist\PDFree\{#AppExeName}"; DestDir: "{app}";           Flags: ignoreversion; Components: main
Source: "dist\PDFree\_internal\*";   DestDir: "{app}\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: main
Source: "LOGO.ico";                  DestDir: "{app}";           Flags: ignoreversion; Components: main

; ============================================================
; STEP 6 — Shortcuts (Start Menu + Desktop)
; ============================================================
[Icons]
; Start Menu
Name: "{group}\{#AppName}";              Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\LOGO.ico"; Comment: "Open PDFree PDF Toolbox"
Name: "{group}\Uninstall {#AppName}";    Filename: "{uninstallexe}"

; Desktop (only if task selected)
Name: "{userdesktop}\{#AppName}";        Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\LOGO.ico"; Tasks: desktopicon

; ============================================================
; STEP 5 — Windows Registry
; ============================================================
[Registry]
; --- Uninstall entry (Add/Remove Programs) ---
; Inno Setup writes this automatically under:
;   HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{AppId}
; The entries below supplement it with display icon and estimated size.
Root: HKLM; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{{A3F2C1D4-8E7B-4F9A-BC23-1D5E6F7A8B9C}_is1"; \
  ValueType: string; ValueName: "DisplayIcon"; ValueData: "{app}\LOGO.ico"; Flags: uninsdeletevalue

; --- File association: .pdf → PDFree (only if task selected) ---
; Registers PDFree as a capable handler for .pdf without forcing it
; as the default (the user can always change it in Windows Settings).
Root: HKCR; Subkey: ".pdf\OpenWithProgids";       ValueType: string; ValueName: "PDFree.Document"; ValueData: ""; Flags: uninsdeletevalue; Tasks: fileassoc
Root: HKCR; Subkey: "PDFree.Document";             ValueType: string; ValueName: "";                ValueData: "PDF Document";              Flags: uninsdeletekey;  Tasks: fileassoc
Root: HKCR; Subkey: "PDFree.Document\DefaultIcon"; ValueType: string; ValueName: "";                ValueData: "{app}\LOGO.ico,0";          Tasks: fileassoc
Root: HKCR; Subkey: "PDFree.Document\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExeName}"" ""%1"""; Tasks: fileassoc

; Notify Windows Shell that file associations changed
Root: HKCR; Subkey: ".pdf"; ValueType: string; ValueName: ""; ValueData: "PDFree.Document"; Flags: uninsdeletevalue; Tasks: fileassoc

; ============================================================
; STEP 2 — Self-extraction / first run after install
; ============================================================
; Inno Setup handles extraction internally. The [Run] section
; optionally launches the app immediately after installation.
[Run]
Filename: "{app}\{#AppExeName}"; Description: "Launch {#AppName}"; Flags: nowait postinstall skipifsilent

; ============================================================
; STEP 8 — PATH / environment variables
; ============================================================
; PDFree is a GUI application so PATH modification is not needed.
; If a CLI companion tool is added in the future, add it here:
;   Root: HKLM; Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment";
;   ValueType: expandsz; ValueName: "Path"; ValueData: "{olddata};{app}"

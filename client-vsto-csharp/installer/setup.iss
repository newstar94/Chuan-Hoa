; Inno Setup Script for VietDoc Standardizer VSTO Add-in
; Generates VietDocStandardizer_Setup.exe for Microsoft Word 2013 & 2016

#define MyAppName "Chuẩn Hóa Thể Thức Văn Bản (VSTO Add-in)"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "VietDoc Standardizer Team"
#define MyAppURL "https://chuanhoathethuc.vn"
#define MyAppExeName "VietDocStandardizer.dll"

[Setup]
AppId={{D3B07384-D113-4F44-972A-60589A19D826}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={autopf}\VietDocStandardizer
DisableProgramGroupPage=yes
OutputBaseFilename=VietDocStandardizer_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest

[Languages]
Name: "vietnamese"; MessagesFile: "compiler:Languages\Vietnamese.isl"

[Files]
Source: "..\bin\Release\VietDocStandardizer.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\bin\Release\VietDocStandardizer.vsto"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\bin\Release\VietDocStandardizer.dll.manifest"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; Register Word 2013 Addin
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\VietDocStandardizer"; ValueType: string; ValueName: "Description"; ValueData: "Add-in Chuẩn hóa thể thức văn bản hành chính (NĐ 30, Đảng, Viettel)"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\VietDocStandardizer"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Chuẩn Hóa Thể Thức"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\VietDocStandardizer"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\VietDocStandardizer"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\VietDocStandardizer.vsto|vstolocal"; Flags: uninsdeletekey

[Run]
Filename: "{sys}\regsvr32.exe"; Parameters: "/s ""{app}\{#MyAppExeName}"""; Flags: nowait postinstall skipifsilent

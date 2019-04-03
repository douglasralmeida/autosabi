; Script para o instalador do Automatizador do SABI
; requer InnoSetup

#define MyAppName "Automatizador do SABI"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Douglas R. Almeida"
#define MyAppURL "https://github.com/douglasralmeida/autosabi"

[Setup]
AppId={{1867F404-B82F-4E50-B316-3F493A1C2D29}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
AllowNoIcons=yes
ChangesEnvironment=true
Compression=lzma
DefaultDirName={pf}\Automatizador do SABI
DefaultGroupName=Automatizador do SABI
DisableWelcomePage=False
MinVersion=0,5.01sp3
OutputBaseFilename=autosabiinstala
SetupIconFile=..\res\setupicone.ico
SolidCompression=yes
ShowLanguageDialog=no
UninstallDisplayName=Automatizador do SABI
UninstallDisplayIcon={app}\automasabi.exe
VersionInfoVersion=1.0.0
VersionInfoProductVersion=1.0
UninstallDisplaySize=5000000
OutputDir=output
DisableReadyPage=True

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Files]
; NOTE: Don't use "Flags: ignoreversion" on any shared system files
Source: "..\bin\automasabi.exe"; DestDir: "{app}"
Source: "..\bin\msvbvm60.dll"; DestDir: "{app}"

[Icons]
Name: "{group}\Automatizador do SABI"; Filename: "{app}\automasabi.exe"; WorkingDir: "{app}"; IconFilename: "{app}\automasabi.exe"; IconIndex: 0
[Setup]
AppName=Cast/Crew Copy&Paste
AppId=CastCrewCopyPaste
AppVerName=Cast/Crew Copy&Paste 1.2.1.0
AppCopyright=Copyright © Doena Soft. 2020 - 2025
AppPublisher=Doena Soft.
AppPublisherURL=http://doena-journal.net/en/dvd-profiler-tools/
DefaultDirName={commonpf32}\Doena Soft.\CastCrewCopyPaste
; DefaultGroupName=Doena Soft.
DirExistsWarning=No
SourceDir=..\CastCrewCopyPaste\bin\x86\Release\net481
Compression=zip/9
AppMutex=InvelosDVDPro
OutputBaseFilename=CastCrewCopyPasteSetup
OutputDir=..\Setup\CastCrewCopyPaste
MinVersion=0,6.1sp1
PrivilegesRequired=admin
WizardStyle=modern
DisableReadyPage=yes
ShowLanguageDialog=no
VersionInfoCompany=Doena Soft.
VersionInfoCopyright=2020 - 2025
VersionInfoDescription=Plugin Template Setup
VersionInfoVersion=1.2.1.0
UninstallDisplayIcon={app}\djdsoft.ico

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
WinVersionTooLowError=This program requires Windows XP or above to be installed.%n%nWindows 9x, NT and 2000 are not supported.

[Types]
Name: "full"; Description: "Full installation"

[Files]
Source: "djdsoft.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "DoenaSoft.DVDProfiler.Helper.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "DoenaSoft.DVDProfiler.Xml.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "DoenaSoft.AbstractionLayer.UI.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "DoenaSoft.AbstractionLayer.Web.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "DoenaSoft.AbstractionLayer.WinForms.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "DoenaSoft.ToolBox.dll"; DestDir: "{app}"; Flags: ignoreversion

Source: "System.Buffers.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "System.Memory.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "System.Numerics.Vectors.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "System.Resources.Extensions.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "System.Runtime.CompilerServices.Unsafe.dll"; DestDir: "{app}"; Flags: ignoreversion

Source: "DoenaSoft.CastCrewCopyPaste.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "DoenaSoft.CastCrewCopyPaste.pdb"; DestDir: "{app}"; Flags: ignoreversion

Source: "de\DoenaSoft.DVDProfiler.Helper.resources.dll"; DestDir: "{app}\de"; Flags: ignoreversion
Source: "de\DoenaSoft.CastCrewCopyPaste.resources.dll"; DestDir: "{app}\de"; Flags: ignoreversion

Source: "ReadMe\ReadMe.html"; DestDir: "{app}\ReadMe"; Flags: ignoreversion

; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Run]
Filename: "{win}\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"; Parameters: "/codebase ""{app}\DoenaSoft.CastCrewCopyPaste.dll"""; Flags: runhidden

;[UninstallDelete]

[UninstallRun]
Filename: "{win}\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"; Parameters: "/u ""{app}\DoenaSoft.CastCrewCopyPaste.dll"""; Flags: runhidden

[Registry]
; Register - Cleanup ahead of time in case the user didn't uninstall the previous version.
Root: HKCR; Subkey: "CLSID\{{FF16CAA3-A77A-41F8-8DAF-19922C695211}"; Flags: dontcreatekey deletekey
Root: HKCR; Subkey: "DoenaSoft.DVDProfiler.CastCrewCopyPaste.Plugin"; Flags: dontcreatekey deletekey
Root: HKCU; Subkey: "Software\Invelos Software\DVD Profiler\Plugins\Identified"; ValueType: none; ValueName: "{{FF16CAA3-A77A-41F8-8DAF-19922C695211}"; ValueData: "0"; Flags: deletevalue
Root: HKLM; Subkey: "Software\Classes\CLSID\{{FF16CAA3-A77A-41F8-8DAF-19922C695211}"; Flags: dontcreatekey deletekey
Root: HKLM; Subkey: "Software\Classes\DoenaSoft.DVDProfiler.CastCrewCopyPaste.Plugin"; Flags: dontcreatekey deletekey
; Unregister
Root: HKCR; Subkey: "CLSID\{{FF16CAA3-A77A-41F8-8DAF-19922C695211}"; Flags: dontcreatekey uninsdeletekey
Root: HKCR; Subkey: "DoenaSoft.DVDProfiler.CastCrewCopyPaste.Plugin"; Flags: dontcreatekey uninsdeletekey
Root: HKCU; Subkey: "Software\Invelos Software\DVD Profiler\Plugins\Identified"; ValueType: none; ValueName: "{{FF16CAA3-A77A-41F8-8DAF-19922C695211}"; ValueData: "0"; Flags: uninsdeletevalue
Root: HKLM; Subkey: "Software\Classes\CLSID\{{FF16CAA3-A77A-41F8-8DAF-19922C695211}"; Flags: dontcreatekey uninsdeletekey
Root: HKLM; Subkey: "Software\Classes\DoenaSoft.DVDProfiler.CastCrewCopyPaste.Plugin"; Flags: dontcreatekey uninsdeletekey

[Code]
function IsDotNET4Detected(): boolean;
// Function to detect dotNet framework version 4
// Returns true if it is available, false it's not.
var
dotNetStatus: boolean;
begin
dotNetStatus := RegKeyExists(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4');
Result := dotNetStatus;
end;

function InitializeSetup(): Boolean;
// Called at the beginning of the setup package.
begin

if not IsDotNET4Detected then
begin
MsgBox( 'The Microsoft .NET Framework version 4 is not installed. Please install it and try again.', mbInformation, MB_OK );
Result := false;
end
else
Result := true;
end;
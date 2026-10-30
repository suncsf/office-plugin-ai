; AI 助手（Office 加载项）安装程序脚本
; 编译：ISCC.exe installer\setup.iss

#define MyAppName "AI 助手（Office 加载项）"
#define MyAppVersion "1.0.1"
#define AddinId "a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d"

[Setup]
AppId={{B7E3A9C2-5D14-4F6E-9A8B-3C2D1E0F4A55}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=office-plugin-ai
DefaultDirName={localappdata}\Programs\OfficeAI
PrivilegesRequired=lowest
MinVersion=10.0
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
OutputDir=output
OutputBaseFilename=AI助手-Setup-{#MyAppVersion}
UninstallDisplayName={#MyAppName}
UninstallDisplayIcon={app}\OfficeAIConsole.exe
DisableProgramGroupPage=yes
; 用户数据目录（%LOCALAPPDATA%\OfficeAI\data）不属于 {app}，卸载不会删除

[Messages]
WelcomeLabel2=本向导将安装 [name/ver]。%n%n内置 Node.js 运行时，无需额外安装环境。%n%n系统要求：Windows 10 或更高版本；Microsoft Office 2019 或更高版本（Word / Excel / PowerPoint）。%n%n注：Office 2019 使用 IE11 兼容界面，部分高级功能（AI 直接执行代码、PPT 形状级操作）不可用；Office 2021 及以上体验完整。%n%n安装完成后，请在打开的“控制台”中确认服务已启动。
SelectTasksLabel2=选择要执行的附加任务：
UninstallAppFullTitle=卸载 %1
ConfirmUninstall=确定要完全移除 %1 吗？%n%n您的对话与附件数据（%LOCALAPPDATA%\OfficeAI\data）将被保留。

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式(&D)"; GroupDescription: "附加任务："
Name: "autostart"; Description: "开机自动启动服务（后台最小化）(&A)"; GroupDescription: "附加任务："; Flags: unchecked
Name: "registeraddin"; Description: "注册到 Word / Excel / PowerPoint（旁加载加载项）(&R)"; GroupDescription: "附加任务："; Flags: checkedonce

[Files]
Source: "staging\app\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Dirs]
; 用户数据目录（卸载不删除）
Name: "{localappdata}\OfficeAI\data"

[Registry]
Root: HKCU; Subkey: "SOFTWARE\Microsoft\Office\16.0\Wef\Developer"; ValueType: string; ValueName: "{#AddinId}"; ValueData: "{app}\manifest.xml"; Tasks: registeraddin; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "OfficeAI助手"; ValueData: """{app}\OfficeAIConsole.exe"" --min"; Tasks: autostart; Flags: uninsdeletevalue

[Icons]
Name: "{autoprograms}\AI 助手\AI 助手控制台"; Filename: "{app}\OfficeAIConsole.exe"; Comment: "启动/停止 AI 助手服务、安装证书、注册加载项"
Name: "{autodesktop}\AI 助手控制台"; Filename: "{app}\OfficeAIConsole.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\OfficeAIConsole.exe"; Description: "启动 AI 助手控制台"; Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "{cmd}"; Parameters: "/C taskkill /F /IM OfficeAIConsole.exe /T"; Flags: runhidden; RunOnceId: "KillConsole"

[Code]
function NextButtonClick(CurPageID: Integer): Boolean;
var
  hasWord: Boolean;
begin
  Result := True;
  if CurPageID = wpReady then
  begin
    hasWord := RegKeyExists(HKEY_CLASSES_ROOT, 'Word.Application') or
               RegKeyExists(HKEY_CLASSES_ROOT, 'Excel.Application') or
               RegKeyExists(HKEY_CLASSES_ROOT, 'PowerPoint.Application');
    if not hasWord then
    begin
      if MsgBox('未检测到 Microsoft Office。本加载项需要 Office 2019 或更高版本（Word/Excel/PowerPoint 至少其一）。' + #13#10 + '仍要继续安装吗？',
        mbConfirmation, MB_YESNO) = IDNO then
        Result := False;
    end;
  end;
end;

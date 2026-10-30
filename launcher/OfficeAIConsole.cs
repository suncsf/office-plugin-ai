// AI 助手 · 控制台（WinForms，.NET Framework 4.x，C#5 兼容）
// 职责：启动/停止服务、安装证书、注册旁加载、启动 Office、开机自启、日志展示；最小化隐藏到系统托盘（不占任务栏）
// 编译：csc /target:winexe /r:System.Windows.Forms.dll /r:System.Drawing.dll OfficeAIConsole.cs
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Win32;

namespace OfficeAI
{
    public class MainForm : Form
    {
        const string SideloadKey = @"SOFTWARE\Microsoft\Office\16.0\Wef\Developer";
        const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
        const string RunValueName = "OfficeAI助手";
        internal const string RestoreEventName = "OfficeAI-Console-Restore";

        readonly string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        readonly string nodeExe;
        // 用户数据与程序分离：卸载/升级不丢失对话与附件
        readonly string dataDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeAI", "data");
        readonly string logFile;
        Process serverProc;
        IntPtr jobHandle;
        Label lblStatus;
        Button btnStart, btnStop, btnCert, btnRegister, btnUnregister;
        CheckBox chkAutoStart;
        TextBox txtLog;
        System.Windows.Forms.Timer timer;
        NotifyIcon tray;
        ContextMenuStrip trayMenu;
        System.Threading.EventWaitHandle restoreEvent;
        bool trayHintShown;

        public MainForm(bool startMinimized)
        {
            nodeExe = Path.Combine(baseDir, "runtime", "node.exe");
            logFile = Path.Combine(dataDir, "server.log");
            Text = "AI 助手 · 控制台（Office 加载项）";
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft YaHei UI", 9F);
            ClientSize = new Size(560, 360);

            lblStatus = new Label();
            lblStatus.Text = "状态：检查中…";
            lblStatus.Location = new Point(14, 14);
            lblStatus.AutoSize = true;
            lblStatus.Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold);
            Controls.Add(lblStatus);

            chkAutoStart = new CheckBox();
            chkAutoStart.Text = "开机自动启动（最小化到后台）";
            chkAutoStart.Location = new Point(360, 12);
            chkAutoStart.AutoSize = true;
            chkAutoStart.Checked = RegistryGetValue(RunKey, RunValueName) != null;
            chkAutoStart.CheckedChanged += delegate { ToggleAutoStart(); };
            Controls.Add(chkAutoStart);

            btnStart = MKBtn("▶ 启动服务", new Point(14, 46), delegate { StartServer(); });
            btnStop = MKBtn("■ 停止服务", new Point(134, 46), delegate { StopServer(); });
            btnCert = MKBtn("🔑 安装证书", new Point(254, 46), delegate { InstallCert(); });
            btnRegister = MKBtn("📌 注册加载项", new Point(374, 46), delegate { RegisterAddin(true); });
            btnUnregister = MKBtn("🗑 取消注册", new Point(14, 86), delegate { RegisterAddin(false); });
            Controls.Add(btnStart); Controls.Add(btnStop); Controls.Add(btnCert);
            Controls.Add(btnRegister); Controls.Add(btnUnregister);
            Controls.Add(MKBtn("📘 启动 Word", new Point(134, 86), delegate { LaunchOffice("WINWORD.EXE"); }));
            Controls.Add(MKBtn("📗 启动 Excel", new Point(254, 86), delegate { LaunchOffice("EXCEL.EXE"); }));
            Controls.Add(MKBtn("📙 启动 PPT", new Point(374, 86), delegate { LaunchOffice("POWERPNT.EXE"); }));

            Button btnData = MKBtn("📁 数据目录", new Point(14, 126), delegate { Process.Start("explorer.exe", dataDir); });
            Button btnLog = MKBtn("📄 服务日志", new Point(134, 126), delegate { OpenLogFile(); });
            Controls.Add(btnData); Controls.Add(btnLog);

            Label hint = new Label();
            hint.Text = "首次使用：① 安装证书 ② 注册加载项 ③ 启动 Word/Excel/PPT →【开始】选项卡 → AI 助手";
            hint.Location = new Point(14, 132);
            hint.AutoSize = true;
            hint.ForeColor = Color.DimGray;
            hint.SendToBack();
            Controls.Add(hint);
            btnData.Location = new Point(14, 158); btnLog.Location = new Point(134, 158);
            Controls.Add(MKBtn("ℹ 关于", new Point(254, 158), delegate { ShowAbout(); }));

            txtLog = new TextBox();
            txtLog.Multiline = true;
            txtLog.ReadOnly = true;
            txtLog.ScrollBars = ScrollBars.Vertical;
            txtLog.BackColor = Color.White;
            txtLog.Font = new Font("Consolas", 8.5F);
            txtLog.Location = new Point(14, 196);
            txtLog.Size = new Size(532, 148);
            Controls.Add(txtLog);

            restoreEvent = CreateRestoreEvent();

            timer = new System.Windows.Forms.Timer();
            timer.Interval = 2000;
            timer.Tick += delegate
            {
                try
                {
                    if (restoreEvent != null && restoreEvent.WaitOne(TimeSpan.Zero))
                        ShowFromTray(); // 第二实例请求唤起
                }
                catch { }
                UpdateStatus();
            };
            timer.Start();

            /* ---------------- 托盘：最小化即隐藏，任务栏不留按钮 ---------------- */
            trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("🖥 打开界面", null, delegate { ShowFromTray(); });
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add("📘 启动 Word", null, delegate { LaunchOffice("WINWORD.EXE"); });
            trayMenu.Items.Add("📗 启动 Excel", null, delegate { LaunchOffice("EXCEL.EXE"); });
            trayMenu.Items.Add("📙 启动 PPT", null, delegate { LaunchOffice("POWERPNT.EXE"); });
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add("ℹ 关于", null, delegate { ShowAbout(); });
            trayMenu.Items.Add("✖ 关闭", null, delegate { Close(); });
            tray = new NotifyIcon();
            tray.Text = "AI 助手控制台";
            tray.Icon = GetAppIcon();
            tray.ContextMenuStrip = trayMenu;
            tray.Click += delegate { ShowFromTray(); };
            tray.Visible = true;
            Resize += delegate
            {
                if (WindowState == FormWindowState.Minimized)
                {
                    Hide();
                    if (!trayHintShown)
                    {
                        trayHintShown = true;
                        tray.BalloonTipTitle = "AI 助手";
                        tray.BalloonTipText = "已最小化到系统托盘，点击托盘图标恢复窗口";
                        try { tray.ShowBalloonTip(3000); } catch { }
                    }
                }
            };

            Load += delegate
            {
                Directory.CreateDirectory(dataDir);
                Log("控制台已启动：" + baseDir);
                EnsureReady();
                if (startMinimized) WindowState = FormWindowState.Minimized;
            };
            FormClosing += delegate(object s, FormClosingEventArgs e)
            {
                if (serverProc != null && !serverProc.HasExited)
                {
                    if (MessageBox.Show("服务正在运行，退出控制台将一并停止服务。确定退出？",
                        "确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK)
                    { e.Cancel = true; return; }
                }
                StopServer();
                if (tray != null)
                {
                    tray.Visible = false; // 退出前移除托盘图标
                    tray.Dispose();
                }
            };
        }

        Button MKBtn(string text, Point pt, EventHandler onClick)
        {
            Button b = new Button();
            b.Text = text;
            b.Location = pt;
            b.Size = new Size(112, 30);
            b.Click += onClick;
            return b;
        }

        /* ---------------- 自动就绪 ---------------- */
        void EnsureReady()
        {
            bool certsOk = File.Exists(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                @".office-addin-dev-certs\localhost.crt"));
            if (!certsOk)
            {
                Log("未检测到开发证书，自动安装…");
                InstallCert(true);
            }
            RegisterAddin(true);
            StartServer();
        }

        /* ---------------- 服务 ---------------- */
        void StartServer()
        {
            if (serverProc != null && !serverProc.HasExited) { Log("服务已在运行中"); return; }
            if (!File.Exists(nodeExe)) { Log("❌ 未找到内置 Node 运行时：" + nodeExe); return; }
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo();
                psi.FileName = nodeExe;
                psi.Arguments = "server/index.js";
                psi.WorkingDirectory = baseDir;
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;
                psi.RedirectStandardOutput = true;
                psi.RedirectStandardError = true;
                psi.EnvironmentVariables["OFFICE_AI_DATA_DIR"] = dataDir;
                serverProc = new Process();
                serverProc.StartInfo = psi;
                serverProc.OutputDataReceived += delegate(object s, DataReceivedEventArgs e) { if (e.Data != null) Log(e.Data); };
                serverProc.ErrorDataReceived += delegate(object s, DataReceivedEventArgs e) { if (e.Data != null) Log("[err] " + e.Data); };
                serverProc.Start();
                serverProc.BeginOutputReadLine();
                serverProc.BeginErrorReadLine();
                AssignToJob(serverProc);
                Log("✅ 服务已启动（进程 " + serverProc.Id + "）https://localhost:3000");
                UpdateStatus();
            }
            catch (Exception ex)
            {
                Log("❌ 启动服务失败：" + ex.Message);
            }
        }

        void StopServer()
        {
            if (serverProc != null && !serverProc.HasExited)
            {
                try { serverProc.Kill(); Log("■ 服务已停止"); }
                catch (Exception ex) { Log("停止失败：" + ex.Message); }
            }
            serverProc = null;
            UpdateStatus();
        }

        void UpdateStatus()
        {
            bool running = serverProc != null && !serverProc.HasExited;
            lblStatus.Text = running ? "状态：● 服务运行中（https://localhost:3000）" : "状态：○ 服务未运行";
            lblStatus.ForeColor = running ? Color.FromArgb(22, 140, 70) : Color.Firebrick;
        }

        /* ---------------- 托盘 ---------------- */
        void ShowFromTray()
        {
            Show();
            WindowState = FormWindowState.Normal;
            TopMost = true;  // 确保从托盘唤起时置前可见
            TopMost = false;
            Activate();
        }

        Icon GetAppIcon()
        {
            try { return Icon.ExtractAssociatedIcon(Application.ExecutablePath); }
            catch { return SystemIcons.Application; }
        }

        System.Threading.EventWaitHandle CreateRestoreEvent()
        {
            try
            {
                return new System.Threading.EventWaitHandle(
                    false, System.Threading.EventResetMode.AutoReset, RestoreEventName);
            }
            catch { return null; }
        }

        /* ---------------- 证书 ---------------- */
        void InstallCert() { InstallCert(false); }
        void InstallCert(bool quiet)
        {
            RunNodeTool("office-addin-dev-certs", "cli.js install", "安装证书", quiet);
        }

        void RunNodeTool(string pkg, string args, string label, bool quiet)
        {
            try
            {
                string script = Path.Combine(baseDir, "node_modules", pkg, args.Split(' ')[0]);
                if (!File.Exists(script)) { Log("❌ 缺少组件：" + pkg); return; }
                ProcessStartInfo psi = new ProcessStartInfo();
                psi.FileName = nodeExe;
                psi.Arguments = "\"" + script + "\" " + args.Substring(args.IndexOf(' ') + 1);
                psi.WorkingDirectory = baseDir;
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;
                psi.RedirectStandardOutput = true;
                psi.RedirectStandardError = true;
                Process p = Process.Start(psi);
                string output = p.StandardOutput.ReadToEnd() + p.StandardError.ReadToEnd();
                p.WaitForExit(120000);
                if (!quiet || p.ExitCode != 0) Log("【" + label + "】exit=" + p.ExitCode + "\r\n" + output.Trim());
                else Log("【" + label + "】完成");
            }
            catch (Exception ex)
            {
                Log("❌ " + label + "失败：" + ex.Message);
            }
        }

        /* ---------------- 旁加载注册 ---------------- */
        void RegisterAddin(bool register)
        {
            try
            {
                string manifest = Path.Combine(baseDir, "manifest.xml");
                string id = null;
                if (File.Exists(manifest))
                {
                    Match m = Regex.Match(File.ReadAllText(manifest), "<Id>([^<]+)</Id>");
                    if (m.Success) id = m.Groups[1].Value.Trim();
                }
                if (id == null) { Log("❌ 未找到 manifest.xml 或其 <Id>"); return; }
                if (register)
                {
                    Run("reg.exe", "add \"HKCU\\" + SideloadKey + "\" /v \"" + id + "\" /t REG_SZ /d \"" + manifest + "\" /f");
                    Log("✅ 加载项已注册（重启 Word/Excel/PPT 后在【开始】选项卡点击 AI 助手）");
                }
                else
                {
                    Run("reg.exe", "delete \"HKCU\\" + SideloadKey + "\" /v \"" + id + "\" /f");
                    Log("✅ 已取消注册加载项");
                }
            }
            catch (Exception ex)
            {
                Log("❌ 注册操作失败：" + ex.Message);
            }
        }

        /* ---------------- 其它 ---------------- */
        void LaunchOffice(string exe)
        {
            try
            {
                Process.Start(new ProcessStartInfo { FileName = exe, UseShellExecute = true });
                Log("已启动 " + exe);
            }
            catch (Exception ex)
            {
                Log("❌ 启动 " + exe + " 失败（未安装 Office？）：" + ex.Message);
            }
        }

        void ToggleAutoStart()
        {
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(RunKey))
            {
                if (chkAutoStart.Checked)
                    k.SetValue(RunValueName, "\"" + Application.ExecutablePath + "\" --min");
                else if (k.GetValue(RunValueName) != null)
                    k.DeleteValue(RunValueName);
            }
            Log(chkAutoStart.Checked ? "✅ 已开启开机自启" : "已关闭开机自启");
        }

        void OpenLogFile()
        {
            if (!File.Exists(logFile)) { MessageBox.Show("暂无日志"); return; }
            Process.Start(new ProcessStartInfo { FileName = "notepad.exe", Arguments = "\"" + logFile + "\"", UseShellExecute = true });
        }

        /* ---------------- 关于 ---------------- */
        void ShowAbout()
        {
            string version = "";
            try
            {
                string pkg = Path.Combine(baseDir, "package.json");
                if (File.Exists(pkg))
                {
                    Match m = Regex.Match(File.ReadAllText(pkg), "\"version\"\\s*:\\s*\"([^\"]+)\"");
                    if (m.Success) version = m.Groups[1].Value.Trim();
                }
            }
            catch { }
            string developer = "", department = "";
            try
            {
                // 隐私脱敏：开发者/部门信息不入仓库；运行目录可选 about-info.local.txt 本地注入（首行开发者、次行部门）
                string infoFile = Path.Combine(baseDir, "about-info.local.txt");
                if (File.Exists(infoFile))
                {
                    string[] lines = File.ReadAllLines(infoFile, Encoding.UTF8);
                    if (lines.Length > 0) developer = lines[0].Trim();
                    if (lines.Length > 1) department = lines[1].Trim();
                }
            }
            catch { }
            string detail = "";
            if (developer != "") detail += "开发人员：" + developer + Environment.NewLine;
            if (department != "") detail += "部门：" + department + Environment.NewLine;
            MessageBox.Show(this,
                "AI 助手 · Office 加载项控制台" + (version != "" ? "（版本 " + version + "）" : "") +
                (detail != "" ? Environment.NewLine + Environment.NewLine + detail.TrimEnd() : ""),
                "关于 AI 助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        void Log(string msg)
        {
            string line = "[" + DateTime.Now.ToString("MM-dd HH:mm:ss") + "] " + msg;
            try
            {
                File.AppendAllText(logFile, line + Environment.NewLine, Encoding.UTF8);
            }
            catch { }
            if (txtLog.IsHandleCreated)
            {
                txtLog.BeginInvoke((MethodInvoker)delegate
                {
                    txtLog.AppendText(line + Environment.NewLine);
                });
            }
        }

        string RegistryGetValue(string key, string name)
        {
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(key))
            {
                return k == null ? null : (string)k.GetValue(name);
            }
        }

        void Run(string file, string args)
        {
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = file;
            psi.Arguments = args;
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            Process.Start(psi).WaitForExit(30000);
        }

        /* ---------------- JobObject：主程序退出时一并结束子进程 ---------------- */
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr CreateJobObject(IntPtr a, string lpName);
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool SetInformationJobObject(IntPtr hJob, int infoType, IntPtr lpInfo, int cbInfo);
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);

        void AssignToJob(Process p)
        {
            try
            {
                if (jobHandle == IntPtr.Zero)
                {
                    jobHandle = CreateJobObject(IntPtr.Zero, null);
                    IntPtr infoPtr = System.Runtime.InteropServices.Marshal.AllocHGlobal(Marshal.SizeOf(typeof(ExtendedLimitInfo)));
                    try
                    {
                        ExtendedLimitInfo info = new ExtendedLimitInfo();
                        info.Basic.LimitFlags = 0x2000; // JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
                        System.Runtime.InteropServices.Marshal.StructureToPtr(info, infoPtr, false);
                        SetInformationJobObject(jobHandle, 9, infoPtr, Marshal.SizeOf(typeof(ExtendedLimitInfo)));
                    }
                    finally { System.Runtime.InteropServices.Marshal.FreeHGlobal(infoPtr); }
                }
                AssignProcessToJobObject(jobHandle, p.Handle);
            }
            catch { /* 尽力而为 */ }
        }

        [StructLayout(LayoutKind.Sequential)]
        struct BasicLimitInfo
        {
            public long PerProcessUserTimeLimit;
            public long PerJobUserTimeLimit;
            public uint LimitFlags;
            public UIntPtr MinimumWorkingSetSize;
            public UIntPtr MaximumWorkingSetSize;
            public uint ActiveProcessLimit;
            public UIntPtr Affinity;
            public uint PriorityClass;
            public uint SchedulingClass;
        }
        [StructLayout(LayoutKind.Sequential)]
        struct ExtendedLimitInfo
        {
            public BasicLimitInfo Basic;
            public ulong IoInfo1; public ulong IoInfo2; public ulong IoInfo3;
            public ulong IoInfo4; public ulong IoInfo5; public ulong IoInfo6;
            public UIntPtr ProcessMemoryLimit;
            public UIntPtr JobMemoryLimit;
            public UIntPtr PeakProcessMemoryUsed;
            public UIntPtr PeakJobMemoryUsed;
        }
    }

    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            bool created;
            System.Threading.Mutex mutex = new System.Threading.Mutex(true, "OfficeAI-Console-SingleInstance", out created);
            if (!created)
            {
                // 已在运行（可能正隐藏于托盘）：通知其唤起窗口，而非弹框
                try
                {
                    using (System.Threading.EventWaitHandle ev =
                        System.Threading.EventWaitHandle.OpenExisting(MainForm.RestoreEventName))
                        ev.Set();
                }
                catch { }
                return;
            }
            bool minimized = args != null && Array.IndexOf(args, "--min") >= 0;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm(minimized));
            GC.KeepAlive(mutex);
        }
    }
}

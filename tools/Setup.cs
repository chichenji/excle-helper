using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("EXCLE助手 安装程序")]
[assembly: AssemblyDescription("EXCLE助手 一键安装与卸载工具（支持 Excel 与 WPS 32/64 位）")]
[assembly: AssemblyProduct("EXCLE助手")]
[assembly: AssemblyVersion("1.1.0.0")]
[assembly: AssemblyFileVersion("1.1.0.0")]

namespace ExcleHelperSetup
{
    internal static class Program
    {
        internal const string VersionStr = "1.1.0";
        internal const string AsmVersion = "1.1.0.0";
        internal const string GuidStr = "{5E1C7A93-2B48-4D16-9F77-8C3A5B1E6D24}";
        internal const string NetCatGuid = "{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}";
        internal const string ProgId = "ExcleHelper.Connect";
        internal const string AsmFullName = "ExcleHelper, Version=" + AsmVersion + ", Culture=neutral, PublicKeyToken=null";

        internal static string InstallDir
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcleHelper");
            }
        }

        internal static string TargetDll { get { return Path.Combine(InstallDir, "ExcleHelper.dll"); } }

        [STAThread]
        private static int Main(string[] args)
        {
            if (args != null && args.Length > 0)
            {
                string cmd = (args[0] ?? "").Trim().ToLowerInvariant();
                string msg;
                if (cmd == "--install" || cmd == "/install")
                    return DoInstall(out msg) ? 0 : 1;
                if (cmd == "--uninstall" || cmd == "/uninstall")
                    return DoUninstall(out msg) ? 0 : 1;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new SetupForm());
            return 0;
        }

        internal static bool IsInstalled(out string currentCodeBase)
        {
            currentCodeBase = "";
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(@"Software\Classes\CLSID\" + GuidStr + @"\InprocServer32"))
                {
                    if (k == null) return false;
                    currentCodeBase = Convert.ToString(k.GetValue("CodeBase") ?? "");
                    return currentCodeBase.Length > 0;
                }
            }
            catch { return false; }
        }

        internal static string DetectRunningHosts()
        {
            bool xl = Process.GetProcessesByName("EXCEL").Length > 0;
            bool wps = Process.GetProcessesByName("et").Length > 0;
            if (xl && wps) return "Excel 与 WPS 表格";
            if (xl) return "Excel";
            if (wps) return "WPS 表格";
            return null;
        }

        internal static bool DoInstall(out string message)
        {
            try
            {
                Directory.CreateDirectory(InstallDir);
                string dllPath = TargetDll;

                // 若旧 DLL 正在被宿主占用，先安全挪走旧文件再释放新文件
                if (File.Exists(dllPath))
                {
                    try { File.Delete(dllPath); }
                    catch
                    {
                        string parked = Path.Combine(InstallDir, "ExcleHelper.old-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".dll");
                        File.Move(dllPath, parked);
                    }
                }

                Assembly asm = Assembly.GetExecutingAssembly();
                using (Stream src = asm.GetManifestResourceStream("ExcleHelper.dll"))
                {
                    if (src == null)
                    {
                        message = "安装包内未找到嵌入的 ExcleHelper.dll。";
                        return false;
                    }
                    using (FileStream dst = new FileStream(dllPath, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        src.CopyTo(dst);
                    }
                }

                // 清理历史遗留的 .old-*.dll
                try
                {
                    foreach (string f in Directory.GetFiles(InstallDir, "ExcleHelper.old-*.dll"))
                    {
                        try { File.Delete(f); } catch { }
                    }
                }
                catch { }

                string codeBase = new Uri(dllPath).AbsoluteUri;
                RegisterClsid(@"Software\Classes\CLSID", codeBase);
                RegisterClsid(@"Software\Classes\WOW6432Node\CLSID", codeBase);

                using (RegistryKey pk = Registry.CurrentUser.CreateSubKey(@"Software\Classes\" + ProgId))
                {
                    pk.SetValue("", ProgId);
                    using (RegistryKey ck = pk.CreateSubKey("CLSID")) ck.SetValue("", GuidStr);
                }

                RegisterAddin(@"Software\Microsoft\Office\Excel\Addins\" + ProgId);
                RegisterAddin(@"Software\Kingsoft\Office\6.0\et\Addins\" + ProgId);

                using (RegistryKey wl = Registry.CurrentUser.CreateSubKey(@"Software\Kingsoft\Office\ET\AddinsWL"))
                {
                    wl.SetValue(ProgId, "", RegistryValueKind.String);
                }
                try
                {
                    using (RegistryKey bl = Registry.CurrentUser.OpenSubKey(@"Software\Kingsoft\Office\ET\AddinsBL", true))
                    {
                        if (bl != null) bl.DeleteValue(ProgId, false);
                    }
                }
                catch { }

                // 清理旧版 GridKit 原型注册项，避免功能区重复
                string[] legacy =
                {
                    @"Software\Microsoft\Office\Excel\Addins\GridKit.Connect",
                    @"Software\Kingsoft\Office\6.0\et\Addins\GridKit.Connect",
                    @"Software\Classes\CLSID\{A7C4E1B2-9F3D-4E8A-B6C1-2D5F7A9E3B41}",
                    @"Software\Classes\WOW6432Node\CLSID\{A7C4E1B2-9F3D-4E8A-B6C1-2D5F7A9E3B41}",
                    @"Software\Classes\GridKit.Connect"
                };
                foreach (string lk in legacy)
                {
                    try { Registry.CurrentUser.DeleteSubKeyTree(lk, false); } catch { }
                }

                message = "安装成功！\n\n"
                    + "安装路径：" + dllPath + "\n"
                    + "已注册：Excel (32/64位) + WPS 表格 (32/64位 + 白名单)\n\n"
                    + "请重启 Excel 或 WPS 表格，即可看到「EXCLE助手」「格式与数据」「工具与输出」三个选项卡。\n"
                    + "（WPS 用户请确认：选项 → 信任中心 → 已勾选「启用第三方 COM 加载项」）";
                return true;
            }
            catch (Exception ex)
            {
                message = "安装失败：" + ex.Message;
                return false;
            }
        }

        private static void RegisterClsid(string root, string codeBase)
        {
            string basePath = root + @"\" + GuidStr;
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(basePath))
            {
                k.SetValue("", ProgId);
                using (RegistryKey inp = k.CreateSubKey("InprocServer32"))
                {
                    inp.SetValue("", "mscoree.dll");
                    inp.SetValue("ThreadingModel", "Both");
                    inp.SetValue("Assembly", AsmFullName);
                    inp.SetValue("Class", ProgId);
                    inp.SetValue("RuntimeVersion", "v4.0.30319");
                    inp.SetValue("CodeBase", codeBase);
                    using (RegistryKey ver = inp.CreateSubKey(AsmVersion))
                    {
                        ver.SetValue("Assembly", AsmFullName);
                        ver.SetValue("Class", ProgId);
                        ver.SetValue("RuntimeVersion", "v4.0.30319");
                        ver.SetValue("CodeBase", codeBase);
                    }
                }
                using (RegistryKey p = k.CreateSubKey("ProgId")) p.SetValue("", ProgId);
                using (RegistryKey cat = k.CreateSubKey(@"Implemented Categories\" + NetCatGuid)) { }
            }
        }

        private static void RegisterAddin(string path)
        {
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(path))
            {
                k.SetValue("FriendlyName", "EXCLE助手", RegistryValueKind.String);
                k.SetValue("Description", "Excel / WPS 表格效率工具箱（全功能免费，纯本地运算）", RegistryValueKind.String);
                k.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                k.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);
            }
        }

        internal static bool DoUninstall(out string message)
        {
            try
            {
                string[] keys =
                {
                    @"Software\Microsoft\Office\Excel\Addins\" + ProgId,
                    @"Software\Kingsoft\Office\6.0\et\Addins\" + ProgId,
                    @"Software\Classes\" + ProgId,
                    @"Software\Classes\CLSID\" + GuidStr,
                    @"Software\Classes\WOW6432Node\CLSID\" + GuidStr
                };
                foreach (string k in keys)
                {
                    try { Registry.CurrentUser.DeleteSubKeyTree(k, false); } catch { }
                }
                try
                {
                    using (RegistryKey wl = Registry.CurrentUser.OpenSubKey(@"Software\Kingsoft\Office\ET\AddinsWL", true))
                    {
                        if (wl != null) wl.DeleteValue(ProgId, false);
                    }
                }
                catch { }

                bool dllLocked = false;
                if (File.Exists(TargetDll))
                {
                    try { File.Delete(TargetDll); }
                    catch { dllLocked = true; }
                }

                message = "已卸载 EXCLE助手 的 Excel 与 WPS 注册项。"
                    + (dllLocked ? "\n（DLL 正被打开的表格程序占用，关闭后可手动删除：" + TargetDll + "）" : "")
                    + "\n\n请重启 Excel 或 WPS 表格生效。";
                return true;
            }
            catch (Exception ex)
            {
                message = "卸载失败：" + ex.Message;
                return false;
            }
        }
    }

    internal sealed class SetupForm : Form
    {
        private readonly Color BrandDark = Color.FromArgb(29, 45, 109);
        private readonly Color BrandBlue = Color.FromArgb(37, 99, 235);
        private Label _status;
        private Button _btnInstall, _btnUninstall, _btnClose;

        internal SetupForm()
        {
            Text = "EXCLE助手 v" + Program.VersionStr + " 安装程序";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = true;
            ClientSize = new Size(520, 320);
            BackColor = Color.White;
            Font = new Font("Microsoft YaHei UI", 9.5f);

            Label desc = new Label();
            desc.Text = "• 支持 Microsoft Excel (32/64位) 与 WPS 表格 ET (32/64位)\n"
                      + "• 320 个本地功能 + AI 助手，全功能免费、无需管理员权限\n"
                      + "• 安装位置：" + Program.InstallDir + "\n"
                      + "• WPS 用户请在「选项 → 信任中心」勾选「启用第三方 COM 加载项」";
            desc.SetBounds(24, 96, 472, 92);
            desc.ForeColor = Color.FromArgb(51, 65, 85);
            Controls.Add(desc);

            _status = new Label();
            _status.SetBounds(24, 198, 472, 46);
            _status.ForeColor = Color.FromArgb(15, 23, 42);
            _status.BackColor = Color.FromArgb(248, 250, 252);
            _status.Padding = new Padding(10, 8, 10, 8);
            Controls.Add(_status);

            _btnInstall = new Button();
            _btnInstall.Text = "一键安装 / 更新";
            _btnInstall.SetBounds(24, 264, 168, 36);
            _btnInstall.BackColor = BrandBlue;
            _btnInstall.ForeColor = Color.White;
            _btnInstall.FlatStyle = FlatStyle.Flat;
            _btnInstall.FlatAppearance.BorderSize = 0;
            _btnInstall.Font = new Font("Microsoft YaHei UI", 9.5f, FontStyle.Bold);
            _btnInstall.Cursor = Cursors.Hand;
            _btnInstall.Click += OnInstallClick;
            Controls.Add(_btnInstall);

            _btnUninstall = new Button();
            _btnUninstall.Text = "卸载";
            _btnUninstall.SetBounds(206, 264, 100, 36);
            _btnUninstall.FlatStyle = FlatStyle.Flat;
            _btnUninstall.ForeColor = Color.FromArgb(71, 85, 105);
            _btnUninstall.Cursor = Cursors.Hand;
            _btnUninstall.Click += OnUninstallClick;
            Controls.Add(_btnUninstall);

            _btnClose = new Button();
            _btnClose.Text = "关闭";
            _btnClose.SetBounds(406, 264, 90, 36);
            _btnClose.FlatStyle = FlatStyle.Flat;
            _btnClose.Cursor = Cursors.Hand;
            _btnClose.Click += delegate { Close(); };
            Controls.Add(_btnClose);

            RefreshStatus();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle header = new Rectangle(0, 0, ClientSize.Width, 78);
            using (LinearGradientBrush b = new LinearGradientBrush(header, BrandDark, BrandBlue, 18f))
                g.FillRectangle(b, header);
            using (Font tf = new Font("Microsoft YaHei UI", 16f, FontStyle.Bold))
            using (Font sf = new Font("Microsoft YaHei UI", 8.5f))
            using (SolidBrush w = new SolidBrush(Color.White))
            using (SolidBrush w2 = new SolidBrush(Color.FromArgb(205, 219, 234, 254)))
            {
                g.DrawString("EXCLE助手  v" + Program.VersionStr, tf, w, 24, 14);
                g.DrawString("Excel / WPS 表格效率工具箱  ·  单文件零依赖插件", sf, w2, 26, 48);
            }
        }

        private void RefreshStatus()
        {
            string cb;
            bool installed = Program.IsInstalled(out cb);
            string running = Program.DetectRunningHosts();
            if (installed)
                _status.Text = "当前状态：已注册\n位置：" + cb;
            else
                _status.Text = "当前状态：尚未安装" + (running != null ? "\n提示：检测到正在运行 " + running + "，安装后请重启生效。" : "");
        }

        private void OnInstallClick(object sender, EventArgs e)
        {
            string running = Program.DetectRunningHosts();
            if (running != null)
            {
                DialogResult r = MessageBox.Show(
                    "检测到正在运行 " + running + "。\n建议先保存并关闭后再安装，也可继续安装后重启生效。\n\n是否继续安装？",
                    Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (r != DialogResult.Yes) return;
            }
            string msg;
            bool ok = Program.DoInstall(out msg);
            RefreshStatus();
            MessageBox.Show(msg, Text, MessageBoxButtons.OK, ok ? MessageBoxIcon.Information : MessageBoxIcon.Error);
        }

        private void OnUninstallClick(object sender, EventArgs e)
        {
            if (MessageBox.Show("确认卸载 EXCLE助手（移除当前用户的 Excel 与 WPS 注册项）？",
                Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;
            string msg;
            bool ok = Program.DoUninstall(out msg);
            RefreshStatus();
            MessageBox.Show(msg, Text, MessageBoxButtons.OK, ok ? MessageBoxIcon.Information : MessageBoxIcon.Error);
        }
    }
}

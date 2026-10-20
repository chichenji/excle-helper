using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcleHelper
{
    /// <summary>
    /// L0 host shell: a plain, VSTO-free COM add-in.
    /// Excel calls OnConnection -> GetCustomUI -> OnButton (late bound IDispatch).
    /// No VSTO runtime, no PIA, no administrator rights, no license layer.
    /// </summary>
    [ComVisible(true)]
    [Guid("5E1C7A93-2B48-4D16-9F77-8C3A5B1E6D24")]
    [ProgId("ExcleHelper.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        // ---- load diagnostics (opt-in: set EXCLE_TRACE=1) ----
        internal static void Trace(string s)
        {
            try
            {
                if (Environment.GetEnvironmentVariable("EXCLE_TRACE") != "1") return;
                string f = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "exclehelper-load.log");
                System.IO.File.AppendAllText(f,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  [" +
                    System.Diagnostics.Process.GetCurrentProcess().Id + "]  " + s + "\r\n");
            }
            catch { }
        }

        public Connect()
        {
            Trace("ctor: instance created");
        }

        // ---- IDTExtensibility2 ----
        public void OnConnection(object Application, ext_ConnectMode ConnectMode,
                                 object AddInInst, ref Array custom)
        {
            Trace("OnConnection: enter mode=" + (int)ConnectMode);
            Xl.App = Application;
            // quiet mode for scripted / unattended runs: never block on a message box
            try
            {
                if (Environment.GetEnvironmentVariable("EXCLE_SILENT") == "1")
                {
                    Xl.Silent = true;
                    Trace("OnConnection: silent mode from EXCLE_SILENT");
                }
            }
            catch { }
            bool published = Rot.Publish(this);
            Trace("OnConnection: exit, published=" + published);
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Trace("OnDisconnection: enter mode=" + (int)RemoveMode);
            Rot.Unpublish();
            Xl.App = null;
        }

        public void OnAddInsUpdate(ref Array custom) { Trace("OnAddInsUpdate"); }

        public void OnStartupComplete(ref Array custom) { Trace("OnStartupComplete"); }

        public void OnBeginShutdown(ref Array custom) { Trace("OnBeginShutdown"); }

        // ---- IRibbonExtensibility ----
        public string GetCustomUI(string RibbonID)
        {
            Trace("GetCustomUI: " + RibbonID);
            Assembly asm = Assembly.GetExecutingAssembly();
            using (Stream s = asm.GetManifestResourceStream("ExcleHelper.ribbon.xml"))
            {
                if (s == null) return "";
                using (StreamReader r = new StreamReader(s, Encoding.UTF8))
                {
                    return r.ReadToEnd();
                }
            }
        }

        // ---- ribbon callbacks ----
        // Office reaches these through IDispatch on the class's AutoDispatch
        // interface. A parameter typed as a COM interface (IRibbonControl /
        // IRibbonUI) is not always expressible there, in which case Office can
        // not bind the name and silently drops the callback. Taking `object`
        // keeps the signature VARIANT-compatible; the members are then read by
        // IDispatch reflection.
        internal static string ComProp(object o, string name)
        {
            if (o == null) return "";
            try
            {
                object v = o.GetType().InvokeMember(
                    name,
                    BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null, o, null);
                return v == null ? "" : Convert.ToString(v);
            }
            catch { return ""; }
        }

        // ---- ribbon callback: per-control icon (our own set, no imageMso) ----
        private static int _imgCalls;

        public object OnGetImage(object control)
        {
            try
            {
                _imgCalls++;
                if (_imgCalls <= 2 || _imgCalls % 25 == 0) Trace("OnGetImage #" + _imgCalls + " id=" + ComProp(control, "Id"));
                string id = ComProp(control, "Id");
                Feature f = Catalog.ById(id);
                if (f == null) return Icons.ToPictureDisp(Icons.Get("·", "ui."));
                return Icons.ToPictureDisp(Icons.Get(f.Label, f.Tag));
            }
            catch { return null; }
        }
        // fired once the custom UI is instantiated
        public void OnRibbonLoad(object ribbon)
        {
            Trace("OnRibbonLoad: custom UI accepted by Office");
        }

        // every button lands here
        public void OnButton(object control)
        {
            string tag = ComProp(control, "Tag");
            string id = ComProp(control, "Id");
            Trace("OnButton: id=" + id + " tag=" + tag);
            if (string.IsNullOrEmpty(tag)) tag = id;
            RunOp(tag);
        }

        /// <summary>
        /// Public entry point: also callable from VBA / PowerShell / tests.
        /// e.g. Application.COMAddIns("ExcleHelper.Connect").Object.RunOp("t.keep.cn")
        /// </summary>
        public void RunOp(string tag)
        {
            if (!Xl.Ready()) { Trace("RunOp: no application"); return; }
            if (string.IsNullOrEmpty(tag))
            {
                Trace("RunOp: empty tag");
                Xl.Bad("未识别的按钮。");
                return;
            }
            Ops.Init();
            if (!Ops.Has(tag))
            {
                Trace("RunOp: NOT REGISTERED " + tag);
                Xl.Bad("该功能尚未实现：" + tag);
                return;
            }
            Trace("RunOp: begin " + tag);
            try
            {
                Ops.Run(tag);
                Trace("RunOp: done " + tag + " | " + Xl.Log.ToString().Replace("\r\n", " ; "));
                Xl.Log.Length = 0;
            }
            catch (Exception ex)
            {
                Trace("RunOp: EXCEPTION " + tag + " -> " + ex.GetType().Name + ": " + ex.Message);
                Xl.Log.Length = 0;
                Xl.Bad("操作失败：" + ex.Message);
            }
            finally
            {
                try { Xl.App.ScreenUpdating = true; } catch { }
                try { Xl.App.EnableEvents = true; } catch { }
            }
            if (!string.Equals(tag, "x.repeat", StringComparison.OrdinalIgnoreCase))
                Xl.LastTag = tag;
        }

        // ---- introspection helpers (used by the installer / tests) ----
        public string Ping()
        {
            return "EXCLE助手 OK / ops=" + Ops.Count + " / connected=" + (Xl.App != null);
        }

        /// <summary>导出图标总览 PNG（设计评审用，不需要 Excel）。</summary>
        public string ExportIconSheet() { return Icons.RenderSheet(); }

        /// <summary>Suppress message boxes; notifications accumulate in the log (scripted runs).</summary>
        public void SetSilent(bool on)
        {
            Xl.Silent = on;
        }

        /// <summary>Return and clear accumulated notifications.</summary>
        public string DrainLog()
        {
            string s = Xl.Log.ToString();
            Xl.Log.Length = 0;
            return s;
        }

        public string Version
        {
            get { return "1.0.0"; }
        }

        public string LastOperation
        {
            get { return Xl.LastTag; }
        }

        public string ListOps()
        {
            return "ops=" + Ops.Count;
        }
    }
}

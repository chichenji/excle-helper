using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>AI 界面层：进度窗、结果预览、设置、对话面板，以及"后台请求 + UI 线程回填"的调度器。</summary>
    internal static class AiRunner
    {
        /// <summary>
        /// 在后台线程执行 work（只允许纯网络/字符串操作），UI 线程显示进度并可取消。
        /// work 的入参是一个"是否已取消"的委托。
        /// </summary>
        internal static bool Run(string title, Func<Func<bool>, string> work, out string text, out string error)
        {
            string result = null, err = null;
            bool cancelFlag = false;
            AiProgress dlg = new AiProgress(title);
            dlg.CancelRequested += delegate { cancelFlag = true; };
            dlg.Show();
            dlg.Refresh();
            System.Threading.Tasks.Task task = System.Threading.Tasks.Task.Factory.StartNew(delegate
            {
                try { result = work(delegate { return cancelFlag; }); }
                catch (Exception ex) { err = AiHttp.Sanitize(ex.Message); }
            });
            while (!task.IsCompleted)
            {
                Application.DoEvents();
                System.Threading.Thread.Sleep(30);
                if (dlg.IsDisposed) cancelFlag = true;
            }
            try { dlg.Close(); dlg.Dispose(); } catch { }
            text = result;
            error = err;
            return err == null && result != null;
        }
    }
}

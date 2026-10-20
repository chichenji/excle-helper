using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>回退相关操作（undo.*）。</summary>
    internal static class UndoOps
    {
        internal static void Register()
        {
            Ops.Reg("undo.last", UndoLast);
            Ops.Reg("undo.redo", RedoLast);
            Ops.Reg("undo.history", History);
            Ops.Reg("undo.clear", Clear);
        }

        private static bool UndoLast() { return Journal.UndoLast(); }
        private static bool RedoLast() { return Journal.RedoLast(); }
        private static bool Clear() { Journal.Clear(); return true; }

        private static bool History()
        {
            List<string> lines = Journal.HistoryLines();
            if (lines.Count == 0) { Xl.Say("还没有可回退的操作记录。"); return true; }
            int picked = HistoryForm.Pick(lines, Journal.UndoCount, Journal.RedoCount);
            if (picked < 0) return true;
            Journal.RollbackTo(picked);
            return true;
        }
    }

    /// <summary>历史记录窗：列出最近操作，双击或点「回退到此前」执行多步回退。</summary>
    internal sealed class HistoryForm : Form
    {
        private ListBox _lb;
        private int _picked = -1;

        internal static int Pick(List<string> lines, int undoCount, int redoCount)
        {
            using (HistoryForm f = new HistoryForm(lines, undoCount, redoCount))
            {
                f.ShowDialog();
                return f._picked;
            }
        }

        private HistoryForm(List<string> lines, int undoCount, int redoCount)
        {
            Text = "回退历史";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(560, 420);
            Font = new Font("Microsoft YaHei UI", 9.5f);
            KeyPreview = true;

            Label tip = new Label();
            tip.Text = "最近 " + undoCount + " 步可回退（可重做 " + redoCount + " 步）   ·   双击某步 = 回退到它之前";
            tip.SetBounds(14, 10, 530, 20);
            tip.ForeColor = Color.FromArgb(100, 116, 139);
            Controls.Add(tip);

            _lb = new ListBox();
            _lb.SetBounds(14, 34, 532, 330);
            _lb.IntegralHeight = false;
            foreach (string s in lines) _lb.Items.Add(s);
            _lb.DoubleClick += delegate { Accept(); };
            Controls.Add(_lb);

            Button ok = new Button();
            ok.Text = "回退到此前";
            ok.SetBounds(320, 374, 108, 32);
            ok.Click += delegate { Accept(); };
            Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "取消";
            cancel.SetBounds(438, 374, 108, 32);
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);

            _lb.KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter) { Accept(); e.SuppressKeyPress = true; }
            };
            KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Escape) { DialogResult = DialogResult.Cancel; Close(); }
            };
            Load += delegate { if (_lb.Items.Count > 0) _lb.SelectedIndex = 0; _lb.Focus(); };
        }

        private void Accept()
        {
            _picked = _lb.SelectedIndex;
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}

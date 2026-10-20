using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>
    /// ONE generic parameter dialog for every operation, built from a label list.
    /// This replaces the original add-in's "one WinForms form per feature"
    /// approach (300+ designer forms) with a single reusable component.
    /// </summary>
    internal static class Ask
    {
        private sealed class Dlg : Form
        {
            internal readonly List<TextBox> Boxes = new List<TextBox>();

            internal Dlg(string title, string[] labels, string[] defaults)
            {
                Text = "EXCLE助手 - " + title;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                StartPosition = FormStartPosition.CenterScreen;
                MaximizeBox = false;
                MinimizeBox = false;
                ShowInTaskbar = false;
                Font = new Font("Microsoft YaHei UI", 9f);
                AutoScaleMode = AutoScaleMode.Dpi;

                int pad = 14;
                int labelW = 132;
                int boxW = 250;
                int rowH = 30;
                int top = pad;

                for (int i = 0; i < labels.Length; i++)
                {
                    Label lb = new Label();
                    lb.Text = labels[i];
                    lb.AutoSize = false;
                    lb.TextAlign = ContentAlignment.MiddleLeft;
                    lb.SetBounds(pad, top + 4, labelW, 22);
                    Controls.Add(lb);

                    TextBox tb = new TextBox();
                    tb.SetBounds(pad + labelW + 6, top, boxW, 24);
                    tb.Text = (defaults != null && i < defaults.Length) ? defaults[i] : "";
                    Controls.Add(tb);
                    Boxes.Add(tb);

                    top += rowH;
                }

                int btnTop = top + 8;
                Button ok = new Button();
                ok.Text = "确定";
                ok.SetBounds(pad + labelW + 6, btnTop, 92, 30);
                ok.DialogResult = DialogResult.OK;
                Controls.Add(ok);

                Button cancel = new Button();
                cancel.Text = "取消";
                cancel.SetBounds(pad + labelW + 6 + 100, btnTop, 92, 30);
                cancel.DialogResult = DialogResult.Cancel;
                Controls.Add(cancel);

                AcceptButton = ok;
                CancelButton = cancel;

                ClientSize = new Size(pad * 2 + labelW + 6 + boxW, btnTop + 30 + pad);
            }
        }

        /// <summary>Show a multi-field prompt. Returns null when cancelled.</summary>
        internal static string[] Show(string title, string[] labels, string[] defaults)
        {
            using (Dlg d = new Dlg(title, labels, defaults))
            {
                if (d.ShowDialog() != DialogResult.OK) return null;
                string[] result = new string[d.Boxes.Count];
                for (int i = 0; i < d.Boxes.Count; i++) result[i] = d.Boxes[i].Text;
                return result;
            }
        }

        internal static string One(string title, string label, string def)
        {
            string[] r = Show(title, new string[] { label }, new string[] { def });
            return r == null ? null : r[0];
        }

        internal static bool Int(string title, string label, string def, out int value)
        {
            value = int.MinValue;
            string s = One(title, label, def);
            if (s == null) return false;
            s = s.Trim();
            if (!int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
            {
                Xl.Bad("请输入整数。");
                return false;
            }
            return true;
        }

        internal static bool Dbl(string title, string label, string def, out double value)
        {
            value = 0d;
            string s = One(title, label, def);
            if (s == null) return false;
            s = s.Trim();
            if (!double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
            {
                Xl.Bad("请输入数字。");
                return false;
            }
            return true;
        }

        internal static bool Text(string title, string label, string def, out string value)
        {
            value = One(title, label, def);
            return value != null;
        }
    }
}

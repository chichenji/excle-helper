using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Text;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>
    /// AI 对话面板：品牌化窗口 + 流式输出 + 引用选区 + 设置入口。
    /// 与功能面板、AI 弹窗共用同一套视觉语言（渐变头部 / 扁平按钮 / 圆角输入框）。
    /// </summary>
    internal sealed class AiChatPanel : Form
    {
        private static AiChatPanel _instance;

        private RichTextBox _log;
        private FieldBox _input;
        private CheckBox _useSel;
        private FlatButton _send, _stop;
        private Label _modelLabel, _selLabel, _hint;
        private Rectangle _closeRect;
        private bool _closeHot;
        private Point _dragOffset;
        private bool _dragging;

        private readonly List<AiMessage> _history = new List<AiMessage>();
        private volatile bool _cancelled;
        private bool _busy;

        private const int HeaderH = 64;
        private const int FooterH = 86;

        internal static void ShowPanel()
        {
            if (_instance != null && !_instance.IsDisposed)
            {
                _instance.Activate();
                _instance.BringToFront();
                return;
            }
            _instance = new AiChatPanel();
            _instance.Show();
        }

        private AiChatPanel()
        {
            Text = "EXCLE助手 · AI 助手";
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(780, 600);
            MinimumSize = new Size(620, 460);
            BackColor = Color.White;
            ShowInTaskbar = false;
            KeyPreview = true;
            Font = UiKit.FontOf(9.5f, FontStyle.Regular);
            DoubleBuffered = true;
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint
                     | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);

            _log = new RichTextBox();
            _log.ReadOnly = true;
            _log.BorderStyle = BorderStyle.None;
            _log.BackColor = Color.White;
            _log.Font = UiKit.FontOf(9.5f, FontStyle.Regular);
            _log.DetectUrls = false;
            Controls.Add(_log);

            _input = new FieldBox("", true);
            _input.Inner.KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter && !e.Shift)
                {
                    e.SuppressKeyPress = true;
                    Send();
                }
            };
            Controls.Add(_input);

            _useSel = new CheckBox();
            _useSel.Text = "引用当前选区作为上下文";
            _useSel.Checked = true;
            _useSel.ForeColor = Color.FromArgb(51, 65, 85);
            _useSel.FlatStyle = FlatStyle.Flat;
            Controls.Add(_useSel);

            _modelLabel = new Label();
            _modelLabel.ForeColor = UiKit.Muted;
            _modelLabel.Font = UiKit.FontOf(8.5f, FontStyle.Regular);
            Controls.Add(_modelLabel);

            _selLabel = new Label();
            _selLabel.ForeColor = UiKit.Muted;
            _selLabel.Font = UiKit.FontOf(8.5f, FontStyle.Regular);
            Controls.Add(_selLabel);

            _hint = new Label();
            _hint.ForeColor = UiKit.Faint;
            _hint.Font = UiKit.FontOf(8.5f, FontStyle.Regular);
            _hint.Text = "Enter 发送 · Shift+Enter 换行 · 写入可回退";
            Controls.Add(_hint);

            _send = new FlatButton("发送", FlatButton.Kind.Primary);
            _send.Click += delegate { Send(); };
            Controls.Add(_send);

            _stop = new FlatButton("停止", FlatButton.Kind.Secondary);
            _stop.Enabled = false;
            _stop.Click += delegate { _cancelled = true; };
            Controls.Add(_stop);

            FlatButton set = new FlatButton("设置", FlatButton.Kind.Ghost);
            set.Click += delegate { AiSettingsForm.Open(); RefreshModelLabel(); };
            Controls.Add(set);

            FlatButton clear = new FlatButton("清空", FlatButton.Kind.Ghost);
            clear.Click += delegate { _history.Clear(); _log.Clear(); AppendTip(); };
            Controls.Add(clear);

            _settingsBtn = set;
            _clearBtn = clear;

            Load += delegate { RoundCorners(); LayoutChildren(); RefreshModelLabel(); AppendTip(); _input.Inner.Focus(); };
            Resize += delegate { RoundCorners(); LayoutChildren(); };
        }

        private FlatButton _settingsBtn;
        private FlatButton _clearBtn;

        private void RoundCorners()
        {
            try { Region = new Region(UiKit.Rounded(new Rectangle(0, 0, Width, Height), 12)); }
            catch { }
        }

        private void LayoutChildren()
        {
            int w = ClientSize.Width, h = ClientSize.Height;
            int pad = 18;
            _log.SetBounds(pad, HeaderH + 40, w - pad * 2, h - HeaderH - FooterH - 56);
            _input.SetBounds(pad, h - FooterH + 6, w - pad * 2 - 200, FooterH - 44);
            _send.SetBounds(w - pad - 96, h - FooterH + 10, 96, 34);
            _stop.SetBounds(w - pad - 96, h - FooterH + 50, 96, 30);
            _useSel.SetBounds(pad, h - FooterH + 50, 268, 22);
            _hint.SetBounds(pad + 276, h - FooterH + 52, w - pad * 2 - 394, 20);
            _settingsBtn.SetBounds(w - 244, 18, 62, 28);
            _clearBtn.SetBounds(w - 172, 18, 62, 28);
            _modelLabel.SetBounds(pad, HeaderH + 12, w - pad * 2, 18);
            _selLabel.Visible = false;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            _closeRect = new Rectangle(Width - 42, 16, 28, 28);
            UiKit.PaintHeader(g, new Rectangle(0, 0, Width, HeaderH), "AI 助手",
                              "", _closeRect, _closeHot, "✦");
            base.OnPaint(e);
            UiKit.PaintFooterLine(g, Height - FooterH, Width);
            UiKit.PaintBorder(g, new Rectangle(0, 0, Width - 1, Height - 1));
        }

        private void RefreshModelLabel()
        {
            AiConfig c = AiConfig.Load();
            string sel = "未引用选区";
            try
            {
                dynamic s = Xl.Sel;
                if (s != null) sel = "当前选区 " + (string)s.Address(false, false);
            }
            catch { }
            _modelLabel.Text = c.Model + "   ·   " + (AiSecret.Has ? "已配置密钥" : "未配置密钥（请点设置）")
                             + "   ·   " + sel;
        }

        private void AppendTip()
        {
            _log.SelectionStart = _log.TextLength;
            _log.SelectionLength = 0;
            _log.SelectionColor = UiKit.Faint;
            _log.AppendText("可以直接提问，例如：\n");
            _log.AppendText("  · 按地区汇总销售额要怎么写公式？\n");
            _log.AppendText("  · 这列数据有什么异常？\n");
            _log.AppendText("  · 帮我把选区里的地址拆成省市区。\n\n");
        }

        private void Append(string text, Color color, bool bold)
        {
            _log.SelectionStart = _log.TextLength;
            _log.SelectionLength = 0;
            _log.SelectionColor = color;
            if (bold) _log.SelectionFont = new Font(_log.Font, FontStyle.Bold);
            _log.AppendText(text);
            if (bold) _log.SelectionFont = _log.Font;
            _log.ScrollToCaret();
        }

        private void Send()
        {
            if (_busy) return;
            string q = _input.Inner.Text.Trim();
            if (q.Length == 0) return;

            AiConfig cfg = AiConfig.Load();
            string key = Environment.GetEnvironmentVariable("EXCLE_AI_KEY");
            if (string.IsNullOrEmpty(key)) key = AiSecret.Load();
            if (key.Length == 0)
            {
                Append("（还没有配置 API Key，请点右上角「设置」）\n\n", UiKit.Danger, false);
                return;
            }

            string digest = null;
            if (_useSel.Checked)
            {
                try
                {
                    dynamic sel = Xl.Sel;
                    if (sel != null)
                    {
                        int rows, cols;
                        string addr;
                        digest = AiTasks.Digest(sel, cfg, out rows, out cols, out addr);
                    }
                }
                catch { }
            }

            _input.Inner.Clear();
            Append("我：", UiKit.BrandBlue, true);
            Append(q + "\n", UiKit.Ink, false);
            Append("AI：", Color.FromArgb(5, 150, 105), true);

            List<AiMessage> msgs = AiTasks.ChatTask(q, digest, _history);
            _history.Add(new AiMessage("user", q));
            _busy = true;
            _cancelled = false;
            _send.Enabled = false;
            _stop.Enabled = true;

            System.Threading.Tasks.Task.Factory.StartNew(delegate
            {
                string err;
                string full = AiHttp.Chat(cfg, key, msgs, delegate(string piece)
                {
                    try
                    {
                        BeginInvoke((MethodInvoker)delegate
                        {
                            _log.SelectionStart = _log.TextLength;
                            _log.SelectionLength = 0;
                            _log.SelectionColor = UiKit.Ink;
                            _log.AppendText(piece);
                            _log.ScrollToCaret();
                        });
                    }
                    catch { }
                }, delegate { return _cancelled; }, out err);

                string done = full;
                string e2 = err;
                try
                {
                    BeginInvoke((MethodInvoker)delegate
                    {
                        _busy = false;
                        _send.Enabled = true;
                        _stop.Enabled = false;
                        if (done == null) Append("\n（请求失败：" + (e2 ?? "未知错误") + "）\n\n", UiKit.Danger, false);
                        else
                        {
                            Append("\n\n", UiKit.Ink, false);
                            _history.Add(new AiMessage("assistant", done));
                        }
                        RefreshModelLabel();
                    });
                }
                catch { }
            });
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (_closeRect.Contains(e.Location)) { Close(); return; }
            if (e.Y < HeaderH && e.Button == MouseButtons.Left) { _dragging = true; _dragOffset = e.Location; }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (_dragging && e.Button == MouseButtons.Left)
            {
                Point p = PointToScreen(e.Location);
                Location = new Point(p.X - _dragOffset.X, p.Y - _dragOffset.Y);
            }
            bool hot = _closeRect.Contains(e.Location);
            if (hot != _closeHot) { _closeHot = hot; Invalidate(_closeRect); }
            Cursor = hot ? Cursors.Hand : Cursors.Default;
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            _dragging = false;
        }

        /// <summary>离屏渲染面板（设计评审用）。</summary>
        internal static void RenderShot()
        {
            AiChatPanel f = new AiChatPanel();
            f.StartPosition = FormStartPosition.Manual;
            f.Location = new Point(-4000, -4000);
            f.Show();
            Application.DoEvents();
            f._log.Clear();
            f.Append("我：", UiKit.BrandBlue, true);
            f.Append("按地区汇总销售额，并给出公式\n", UiKit.Ink, false);
            f.Append("AI：", Color.FromArgb(5, 150, 105), true);
            f.Append("用 SUMIF 按地区汇总即可：\n", UiKit.Ink, false);
            f.Append("  =SUMIF(B2:B100,\"华东\",D2:D100)\n", UiKit.Ink, false);
            f.Append("说明：对 B 列等于「华东」的行，汇总 D 列销售额。\n", UiKit.Ink, false);
            f.Append("注意：区域请按实际数据范围调整；多条件请改用 SUMIFS。\n\n", UiKit.Ink, false);
            Application.DoEvents();
            using (Bitmap bmp = new Bitmap(f.Width, f.Height))
            {
                f.DrawToBitmap(bmp, new Rectangle(0, 0, f.Width, f.Height));
                bmp.Save(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "exclehelper-ai-chat.png"),
                         System.Drawing.Imaging.ImageFormat.Png);
            }
            f.Close();
        }
    }
}

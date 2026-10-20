using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>
    /// 功能面板：品牌化、可搜索的操作台。
    /// 目录与功能区同源（Catalog 解析 ribbon.xml）；搜索区整块自绘，支持直接打字、
    /// 中文子串与拼音首字母检索。无子控件，渲染与主题完全一致。
    /// </summary>
    internal sealed class Launcher : Form
    {
        private const int HeaderH = 80;
        private const int FooterH = 42;
        private const int RailW = 184;
        private const int SearchH = 42;
        private const string Placeholder = "搜索功能：输入中文关键词，或拼音首字母（如 wb 找文本）";

        private static Launcher _instance;

        private readonly Color BrandDark = Color.FromArgb(29, 45, 109);
        private readonly Color BrandBlue = Color.FromArgb(37, 99, 235);
        private readonly Color Ink = Color.FromArgb(15, 23, 42);
        private readonly Color Muted = Color.FromArgb(100, 116, 139);
        private readonly Color Line = Color.FromArgb(226, 232, 240);
        private readonly Color RowHover = Color.FromArgb(248, 250, 252);
        private readonly Color RowSel = Color.FromArgb(239, 246, 255);
        private readonly Color ChipIdle = Color.FromArgb(241, 245, 249);
        private readonly Color SearchBg = Color.FromArgb(248, 250, 252);

        private ListBox _cats;
        private ListBox _items;
        private List<Feature> _view = new List<Feature>();
        private string _cat = "全部";
        private string _query = "";
        private int _hover = -1;
        private Rectangle _closeRect;
        private Rectangle _searchRect;
        private Point _dragOffset;

        internal static void ShowPanel()
        {
            if (_instance != null && !_instance.IsDisposed)
            {
                _instance.Activate();
                _instance.BringToFront();
                return;
            }
            _instance = new Launcher();
            _instance.Show();
            _instance.Activate();
        }

        /// <summary>离屏渲染成 PNG，用于设计评审（不依赖可见桌面）。</summary>
        internal static void RenderShot()
        {
            Launcher f = new Launcher();
            f.StartPosition = FormStartPosition.Manual;
            f.Location = new Point(-4000, -4000);
            f.ShowInTaskbar = false;
            f.Show();
            Application.DoEvents();
            Application.DoEvents();
            using (Bitmap bmp = new Bitmap(f.Width, f.Height))
            {
                f.DrawToBitmap(bmp, new Rectangle(0, 0, f.Width, f.Height));
                string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "exclehelper-panel.png");
                bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                f.Close();
                Xl.Say("面板预览已保存：\n" + path);
            }
        }

        internal Launcher()
        {
            Text = "EXCLE助手";
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(920, 600);
            MinimumSize = new Size(760, 480);
            BackColor = Color.White;
            ShowInTaskbar = false;
            KeyPreview = true;
            DoubleBuffered = true;
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint
                     | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
            Font = new Font("Microsoft YaHei UI", 9f);

            BuildCategories();
            BuildItems();

            Load += delegate { RoundCorners(); LayoutChildren(); Reload(); _items.Focus(); };
            Resize += delegate { LayoutChildren(); RoundCorners(); };
            KeyDown += OnKeyDown;
            KeyPress += OnKeyPress;
        }

        // ---------------- layout ----------------
        private void BuildCategories()
        {
            _cats = new ListBox();
            _cats.BorderStyle = BorderStyle.None;
            _cats.DrawMode = DrawMode.OwnerDrawFixed;
            _cats.ItemHeight = 32;
            _cats.IntegralHeight = false;
            _cats.BackColor = Color.FromArgb(248, 250, 252);
            _cats.ForeColor = Ink;
            _cats.Font = new Font("Microsoft YaHei UI", 9.5f);
            _cats.DrawItem += DrawCategory;
            _cats.SelectedIndexChanged += delegate
            {
                if (_cats.SelectedItem == null) return;
                _cat = _cats.SelectedItem.ToString();
                Reload();
            };
            Controls.Add(_cats);
        }

        private void BuildItems()
        {
            _items = new ListBox();
            _items.BorderStyle = BorderStyle.None;
            _items.DrawMode = DrawMode.OwnerDrawFixed;
            _items.ItemHeight = 56;
            _items.IntegralHeight = false;
            _items.BackColor = Color.White;
            _items.Font = new Font("Microsoft YaHei UI", 10f);
            _items.DrawItem += DrawItemRow;
            _items.DoubleClick += delegate { RunSelected(); };
            _items.MouseMove += delegate(object s, MouseEventArgs e)
            {
                int idx = _items.IndexFromPoint(e.Location);
                if (idx != _hover) { _hover = idx; _items.Invalidate(); }
            };
            _items.MouseLeave += delegate { _hover = -1; _items.Invalidate(); };
            Controls.Add(_items);
        }

        private void LayoutChildren()
        {
            int w = ClientSize.Width, h = ClientSize.Height;
            _searchRect = new Rectangle(RailW + 16, HeaderH + 14, w - RailW - 32, SearchH);
            _cats.SetBounds(0, HeaderH, RailW, h - HeaderH - FooterH);
            _items.SetBounds(RailW, HeaderH + 64, w - RailW - 2, h - HeaderH - FooterH - 64);
        }

        private void RoundCorners()
        {
            using (GraphicsPath p = Rounded(new Rectangle(0, 0, Width, Height), 12))
                Region = new Region(p);
        }

        private static GraphicsPath Rounded(Rectangle r, int radius)
        {
            GraphicsPath p = new GraphicsPath();
            int d = radius * 2;
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        // ---------------- data ----------------
        private void Reload()
        {
            _view = new List<Feature>();
            foreach (Feature f in Catalog.All)
            {
                if (_cat != "全部" && f.Group != _cat) continue;
                if (!Catalog.Match(f, _query)) continue;
                _view.Add(f);
            }
            _items.BeginUpdate();
            _items.Items.Clear();
            foreach (Feature f in _view) _items.Items.Add(f.Label);
            _items.EndUpdate();
            if (_items.Items.Count > 0) _items.SelectedIndex = 0;
            _items.Invalidate();

            if (_cats.Items.Count == 0)
            {
                _cats.Items.Add("全部");
                foreach (string c in Catalog.Categories) _cats.Items.Add(c);
                _cats.SelectedIndex = 0;
            }
            Invalidate();
        }

        private void RunSelected()
        {
            int i = _items.SelectedIndex;
            if (i < 0 || i >= _view.Count) return;
            string tag = _view[i].Tag;
            Hide();
            try
            {
                Ops.Init();
                if (Ops.Has(tag)) Ops.Run(tag);
                else Xl.Bad("该功能尚未注册：" + tag);
            }
            catch (Exception ex)
            {
                Xl.Bad("执行失败：" + ex.Message);
            }
            finally
            {
                Close();
            }
        }

        // ---------------- painting ----------------
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            using (LinearGradientBrush b = new LinearGradientBrush(
                       new Rectangle(0, 0, Width, HeaderH), BrandDark, BrandBlue, 16f))
                g.FillRectangle(b, new Rectangle(0, 0, Width, HeaderH));

            using (Font tf = new Font("Microsoft YaHei UI", 18f, FontStyle.Bold))
            using (Font sf = new Font("Microsoft YaHei UI", 8.5f))
            using (SolidBrush w = new SolidBrush(Color.White))
            using (SolidBrush w2 = new SolidBrush(Color.FromArgb(198, 219, 234, 254)))
            {
                g.DrawString("EXCLE助手", tf, w, 26, 15);
                g.DrawString("表格效率工具箱   ·   全功能免费   ·   纯本地运算 · 不上传数据", sf, w2, 28, 50);
            }

            _closeRect = new Rectangle(Width - 46, 16, 28, 28);
            using (SolidBrush cb = new SolidBrush(Color.FromArgb(34, 255, 255, 255)))
                g.FillEllipse(cb, _closeRect);
            using (Pen cp = new Pen(Color.White, 1.7f))
            {
                int cx = _closeRect.X + 9, cy = _closeRect.Y + 9, s = 10;
                g.DrawLine(cp, cx, cy, cx + s, cy + s);
                g.DrawLine(cp, cx + s, cy, cx, cy + s);
            }

            using (GraphicsPath sp = Rounded(_searchRect, 10))
            using (SolidBrush sb = new SolidBrush(SearchBg))
            using (Pen spen = new Pen(Line))
            {
                g.FillPath(sb, sp);
                g.DrawPath(spen, sp);
            }
            using (Font gf = new Font("Segoe UI Symbol", 11.5f))
            using (SolidBrush gb = new SolidBrush(Color.FromArgb(148, 163, 184)))
                g.DrawString("\u2315", gf, gb, _searchRect.X + 20, _searchRect.Y + 8);

            using (Font qf = new Font("Microsoft YaHei UI", 10.5f))
            {
                float tx = _searchRect.X + 52;
                float ty = _searchRect.Y + 10;
                if (_query.Length == 0)
                {
                    using (SolidBrush pb = new SolidBrush(Color.FromArgb(148, 163, 184)))
                        g.DrawString(Placeholder, qf, pb, tx, ty);
                }
                else
                {
                    using (SolidBrush qb = new SolidBrush(Ink))
                        g.DrawString(_query, qf, qb, tx, ty);
                    SizeF sz = g.MeasureString(_query, qf);
                    using (Pen caret = new Pen(BrandBlue, 1.6f))
                        g.DrawLine(caret, tx + sz.Width + 1, ty + 1, tx + sz.Width + 1, ty + 19);
                }
            }

            using (Pen dp = new Pen(Line))
            {
                g.DrawLine(dp, RailW, HeaderH + 64, Width, HeaderH + 64);
                g.DrawLine(dp, RailW, HeaderH, RailW, Height - FooterH);
                g.DrawLine(dp, 0, Height - FooterH, Width, Height - FooterH);
            }

            using (Font ff = new Font("Microsoft YaHei UI", 8.5f))
            using (SolidBrush fb = new SolidBrush(Muted))
            {
                string left = "共 " + Catalog.All.Count + " 个功能   ·   " + _view.Count
                              + " 项匹配   ·   双击或回车执行   ·   Esc 关闭";
                string right = "v1.1   EXCLE助手";
                g.DrawString(left, ff, fb, 22, Height - FooterH + 13);
                SizeF rs = g.MeasureString(right, ff);
                g.DrawString(right, ff, fb, Width - rs.Width - 22, Height - FooterH + 13);
            }

            if (_view.Count == 0)
            {
                using (Font ef = new Font("Microsoft YaHei UI", 10f))
                using (SolidBrush eb = new SolidBrush(Color.FromArgb(148, 163, 184)))
                {
                    Rectangle area = new Rectangle(RailW, HeaderH + 64, Width - RailW, Height - HeaderH - FooterH - 64);
                    StringFormat f = new StringFormat();
                    f.Alignment = StringAlignment.Center;
                    f.LineAlignment = StringAlignment.Center;
                    g.DrawString("没有匹配的功能，换个关键词试试", ef, eb, area, f);
                }
            }

            using (GraphicsPath bp = Rounded(new Rectangle(0, 0, Width - 1, Height - 1), 12))
            using (Pen pen = new Pen(Color.FromArgb(212, 220, 232)))
                g.DrawPath(pen, bp);
        }

        private void DrawCategory(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            bool sel = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            Rectangle r = new Rectangle(10, e.Bounds.Y + 3, RailW - 36, e.Bounds.Height - 6);
            if (sel)
            {
                using (GraphicsPath p = Rounded(r, 7))
                using (SolidBrush b = new SolidBrush(Color.FromArgb(226, 236, 254)))
                    g.FillPath(b, p);
                using (SolidBrush bar = new SolidBrush(BrandBlue))
                    g.FillRectangle(bar, r.X, r.Y + 5, 3, r.Height - 10);
            }
            string text = _cats.Items[e.Index].ToString();
            using (Font f = new Font("Microsoft YaHei UI", 9.5f, sel ? FontStyle.Bold : FontStyle.Regular))
            using (SolidBrush b = new SolidBrush(sel ? BrandDark : Color.FromArgb(51, 65, 85)))
                g.DrawString(text, f, b, r.X + 15, r.Y + (r.Height - f.Height) / 2 - 1);
        }

        private void DrawItemRow(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _view.Count) return;
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            Feature f = _view[e.Index];
            bool sel = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            bool hov = e.Index == _hover;
            int rowW = _items.ClientSize.Width - 30;
            Rectangle r = new Rectangle(10, e.Bounds.Y + 3, rowW, e.Bounds.Height - 6);

            if (sel || hov)
            {
                using (GraphicsPath p = Rounded(r, 8))
                using (SolidBrush b = new SolidBrush(sel ? RowSel : RowHover))
                    g.FillPath(b, p);
                if (sel)
                {
                    using (GraphicsPath p2 = Rounded(r, 8))
                    using (Pen pen = new Pen(Color.FromArgb(191, 219, 254)))
                        g.DrawPath(pen, p2);
                }
            }

            // 图形图标（与功能区同一套图形，选中时加浅色底）
            Rectangle chip = new Rectangle(r.X + 12, r.Y + 11, 34, 34);
            if (sel)
            {
                using (GraphicsPath p = Rounded(chip, 9))
                using (SolidBrush b = new SolidBrush(Color.FromArgb(30, BrandBlue)))
                    g.FillPath(b, p);
            }
            using (Bitmap ic = Glyphs.Render(f.Tag, 26))
                g.DrawImage(ic, new Rectangle(chip.X + 4, chip.Y + 4, 26, 26));

            StringFormat ell = new StringFormat();
            ell.Trimming = StringTrimming.EllipsisCharacter;
            ell.FormatFlags = StringFormatFlags.NoWrap;

            using (Font nf = new Font("Microsoft YaHei UI", 10f, FontStyle.Bold))
            using (SolidBrush nb = new SolidBrush(Ink))
                g.DrawString(f.Label, nf, nb, new RectangleF(r.X + 56, r.Y + 9, r.Width - 72, 20), ell);

            string sub = f.Tab + "  /  " + f.Group;
            if (!string.IsNullOrEmpty(f.Hint)) sub += "   ·   " + f.Hint;
            using (Font sf2 = new Font("Microsoft YaHei UI", 8f))
            using (SolidBrush sb2 = new SolidBrush(Muted))
                g.DrawString(sub, sf2, sb2, new RectangleF(r.X + 56, r.Y + 31, r.Width - 72, 18), ell);
        }

        // ---------------- input ----------------
        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                if (_query.Length > 0) { _query = ""; Reload(); }
                else Close();
                e.SuppressKeyPress = true;
                return;
            }
            if (e.KeyCode == Keys.Enter) { RunSelected(); e.SuppressKeyPress = true; return; }
            if (e.KeyCode == Keys.Back)
            {
                if (_query.Length > 0) { _query = _query.Substring(0, _query.Length - 1); Reload(); }
                e.SuppressKeyPress = true;
                return;
            }
            if (e.KeyCode == Keys.Down && _items.Focused)
            {
                if (_items.SelectedIndex < _items.Items.Count - 1) _items.SelectedIndex++;
                e.SuppressKeyPress = true;
                return;
            }
            if (e.KeyCode == Keys.Up && _items.Focused)
            {
                if (_items.SelectedIndex > 0) _items.SelectedIndex--;
                e.SuppressKeyPress = true;
                return;
            }
            if (e.Control && e.KeyCode == Keys.V)
            {
                try { if (Clipboard.ContainsText()) { _query += Clipboard.GetText().Trim(); Reload(); } }
                catch { }
                e.SuppressKeyPress = true;
            }
        }

        private void OnKeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsControl(e.KeyChar)) return;
            _query += e.KeyChar;
            Reload();
            e.Handled = true;
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (_closeRect.Contains(e.Location)) { Close(); return; }
            if (e.Y < HeaderH && e.Button == MouseButtons.Left) _dragOffset = e.Location;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (e.Button == MouseButtons.Left && _dragOffset != Point.Empty && e.Y < HeaderH)
            {
                Point p = PointToScreen(e.Location);
                Location = new Point(p.X - _dragOffset.X, p.Y - _dragOffset.Y);
            }
            Cursor = _closeRect.Contains(e.Location) ? Cursors.Hand : Cursors.Default;
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            _dragOffset = Point.Empty;
        }

        protected override void OnPaintBackground(PaintEventArgs e) { }
    }
}

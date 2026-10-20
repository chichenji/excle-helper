using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>
    /// 统一 UI 组件：品牌化窗口基类 + 扁平按钮 + 绘制工具。
    /// 设计基调与功能面板一致：深蓝→亮蓝渐变头部、白色主体、圆角、细描边、雅黑字体。
    /// </summary>
    internal static class UiKit
    {
        internal static readonly Color BrandDark = Color.FromArgb(29, 45, 109);
        internal static readonly Color BrandBlue = Color.FromArgb(37, 99, 235);
        internal static readonly Color BrandSoft = Color.FromArgb(239, 246, 255);
        internal static readonly Color Ink = Color.FromArgb(15, 23, 42);
        internal static readonly Color Muted = Color.FromArgb(100, 116, 139);
        internal static readonly Color Faint = Color.FromArgb(148, 163, 184);
        internal static readonly Color Line = Color.FromArgb(226, 232, 240);
        internal static readonly Color Bg = Color.FromArgb(248, 250, 252);
        internal static readonly Color Danger = Color.FromArgb(220, 38, 38);

        /// <summary>执行 UI 动作并把异常写入 %TEMP%\exclehelper-ui.log（设计稿渲染用，避免影响主流程）。</summary>
        internal static void Safe(Action a, string name)
        {
            try { a(); }
            catch (Exception ex)
            {
                try
                {
                    string log = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "exclehelper-ui.log");
                    System.IO.File.AppendAllText(log, name + " 失败：" + ex.GetType().Name + " " + ex.Message + "\r\n");
                }
                catch { }
            }
        }

        internal static Font FontOf(float size, FontStyle style)
        {
            return new Font("Microsoft YaHei UI", size, style);
        }

        internal static GraphicsPath Rounded(Rectangle r, int radius)
        {
            GraphicsPath p = new GraphicsPath();
            int d = radius * 2;
            if (d <= 0) { p.AddRectangle(r); return p; }
            if (d > r.Height) d = r.Height;
            if (d > r.Width) d = r.Width;
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        /// <summary>渐变头部 + 标题/副标题 + 关闭按钮 + 圆角边框。</summary>
        internal static void PaintHeader(Graphics g, Rectangle header, string title, string subtitle,
                                         Rectangle closeRect, bool closeHot, string glyph)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (LinearGradientBrush b = new LinearGradientBrush(header, BrandDark, BrandBlue, 16f))
                g.FillRectangle(b, header);

            int x = header.X + 20;
            if (!string.IsNullOrEmpty(glyph))
            {
                using (Font gf = FontOf(15f, FontStyle.Bold))
                using (SolidBrush gb = new SolidBrush(Color.FromArgb(90, 255, 255, 255)))
                    g.DrawString(glyph, gf, gb, x - 2, header.Y + 15);
                x += 34;
            }
            using (Font tf = FontOf(13f, FontStyle.Bold))
            using (SolidBrush tb = new SolidBrush(Color.White))
                g.DrawString(title, tf, tb, x, header.Y + 13);

            if (!string.IsNullOrEmpty(subtitle))
            {
                using (Font sf = FontOf(8.5f, FontStyle.Regular))
                using (SolidBrush sb = new SolidBrush(Color.FromArgb(198, 219, 234, 254)))
                    g.DrawString(subtitle, sf, sb, x, header.Y + 36);
            }

            using (SolidBrush cb = new SolidBrush(Color.FromArgb(closeHot ? 70 : 34, 255, 255, 255)))
                g.FillEllipse(cb, closeRect);
            using (Pen cp = new Pen(Color.White, 1.7f))
            {
                int cx = closeRect.X + 9, cy = closeRect.Y + 9, s = 10;
                g.DrawLine(cp, cx, cy, cx + s, cy + s);
                g.DrawLine(cp, cx + s, cy, cx, cy + s);
            }
        }

        internal static void PaintFooterLine(Graphics g, int y, int width)
        {
            using (Pen p = new Pen(Line))
                g.DrawLine(p, 0, y, width, y);
        }

        internal static void PaintBorder(Graphics g, Rectangle bounds)
        {
            using (GraphicsPath p = Rounded(bounds, 12))
            using (Pen pen = new Pen(Color.FromArgb(203, 213, 225)))
                g.DrawPath(pen, p);
        }
    }

    /// <summary>品牌化对话框基类：无系统边框、圆角、可拖动、自带关闭按钮。</summary>
    internal class BrandForm : Form
    {
        protected int HeaderH = 64;
        protected string HeadTitle = "";
        protected string HeadSubtitle = "";
        protected string HeadGlyph = "";
        internal bool AllowClose = true;

        private Rectangle _closeRect;
        private bool _closeHot;
        private Point _dragOffset;
        private bool _dragging;

        internal BrandForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.White;
            ShowInTaskbar = false;
            Font = UiKit.FontOf(9.5f, FontStyle.Regular);
            DoubleBuffered = true;
            KeyPreview = true;
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint
                     | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
            Load += delegate { RoundCorners(); };
            Resize += delegate { RoundCorners(); };
        }

        protected void RoundCorners()
        {
            try { Region = new Region(UiKit.Rounded(new Rectangle(0, 0, Width, Height), 12)); }
            catch { }
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            e.Graphics.Clear(BackColor);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Rectangle header = new Rectangle(0, 0, Width, HeaderH);
            _closeRect = new Rectangle(Width - 46, 16, 28, 28);
            UiKit.PaintHeader(g, header, HeadTitle, HeadSubtitle, _closeRect, _closeHot, HeadGlyph);
            base.OnPaint(e);
            UiKit.PaintBorder(g, new Rectangle(0, 0, Width - 1, Height - 1));
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (AllowClose && _closeRect.Contains(e.Location)) { Close(); return; }
            if (e.Y < HeaderH && e.Button == MouseButtons.Left)
            {
                _dragging = true;
                _dragOffset = e.Location;
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (_dragging && e.Button == MouseButtons.Left)
            {
                Point p = PointToScreen(e.Location);
                Location = new Point(p.X - _dragOffset.X, p.Y - _dragOffset.Y);
            }
            bool hot = AllowClose && _closeRect.Contains(e.Location);
            if (hot != _closeHot) { _closeHot = hot; Invalidate(_closeRect); }
            Cursor = hot ? Cursors.Hand : Cursors.Default;
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            _dragging = false;
        }
    }

    /// <summary>扁平按钮：Primary（品牌蓝实心）/ Secondary（白底描边）/ Ghost（无底）。</summary>
    internal sealed class FlatButton : Control
    {
        internal enum Kind { Primary, Secondary, Ghost }

        private Kind _kind = Kind.Secondary;
        private bool _hot, _down;

        internal FlatButton(string text, Kind kind)
        {
            Text = text;
            _kind = kind;
            Size = new Size(96, 32);
            Cursor = Cursors.Hand;
            Font = UiKit.FontOf(9.5f, FontStyle.Regular);
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint
                     | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle r = new Rectangle(0, 0, Width - 1, Height - 1);
            Color fill, border, ink;

            if (_kind == Kind.Primary)
            {
                fill = _down ? Color.FromArgb(29, 78, 216) : (_hot ? Color.FromArgb(59, 130, 246) : UiKit.BrandBlue);
                border = fill;
                ink = Color.White;
            }
            else if (_kind == Kind.Secondary)
            {
                fill = _down ? Color.FromArgb(241, 245, 249) : (_hot ? UiKit.BrandSoft : Color.White);
                border = _hot ? Color.FromArgb(147, 197, 253) : UiKit.Line;
                ink = UiKit.Ink;
            }
            else
            {
                fill = _hot ? Color.FromArgb(241, 245, 249) : Color.White;
                border = Color.White;
                ink = UiKit.Muted;
            }

            using (GraphicsPath p = UiKit.Rounded(r, 7))
            using (SolidBrush b = new SolidBrush(fill))
            using (Pen pen = new Pen(border))
            {
                g.FillPath(b, p);
                if (_kind != Kind.Ghost) g.DrawPath(pen, p);
            }
            using (SolidBrush tb = new SolidBrush(Enabled ? ink : UiKit.Faint))
            {
                StringFormat sf = new StringFormat();
                sf.Alignment = StringAlignment.Center;
                sf.LineAlignment = StringAlignment.Center;
                g.DrawString(Text, Font, tb, new RectangleF(0, 0, Width, Height), sf);
            }
        }

        protected override void OnMouseEnter(EventArgs e) { _hot = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hot = false; _down = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { _down = true; Invalidate(); base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { _down = false; Invalidate(); base.OnMouseUp(e); }
        protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }
    }

    /// <summary>带圆角描边的输入框容器（内部放标准 TextBox）。</summary>
    internal sealed class FieldBox : Panel
    {
        internal TextBox Inner;

        internal FieldBox(string text, bool multiline)
        {
            BackColor = Color.White;
            Padding = new Padding(10, 6, 10, 6);
            Inner = new TextBox();
            Inner.BorderStyle = BorderStyle.None;
            Inner.Font = UiKit.FontOf(9.5f, FontStyle.Regular);
            Inner.ForeColor = UiKit.Ink;
            Inner.Multiline = multiline;
            Inner.Text = text ?? "";
            Inner.Dock = DockStyle.Fill;
            if (multiline) Inner.ScrollBars = ScrollBars.Vertical;
            Controls.Add(Inner);
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint
                     | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (GraphicsPath p = UiKit.Rounded(new Rectangle(0, 0, Width - 1, Height - 1), 7))
            using (SolidBrush b = new SolidBrush(Color.White))
            using (Pen pen = new Pen(Inner.Focused ? Color.FromArgb(147, 197, 253) : UiKit.Line))
            {
                e.Graphics.FillPath(b, p);
                e.Graphics.DrawPath(pen, p);
            }
        }
    }
}

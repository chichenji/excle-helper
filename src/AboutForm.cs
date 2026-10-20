using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Reflection;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>品牌化「关于」窗口。</summary>
    internal sealed class AboutForm : Form
    {
        private static AboutForm _instance;
        private readonly Color BrandDark = Color.FromArgb(30, 58, 138);
        private readonly Color BrandBlue = Color.FromArgb(37, 99, 235);

        internal static void ShowAbout()
        {
            if (_instance != null && !_instance.IsDisposed) { _instance.Activate(); return; }
            _instance = new AboutForm();
            _instance.Show();
            _instance.Activate();
        }

private static int TabCount()
        {
            System.Collections.Generic.List<string> t = new System.Collections.Generic.List<string>();
            foreach (Feature f in Catalog.All) if (!t.Contains(f.Tab)) t.Add(f.Tab);
            return t.Count;
        }

        internal static void RenderShot()
        {
            AboutForm f = new AboutForm();
            f.StartPosition = FormStartPosition.Manual;
            f.Location = new Point(-4000, -4000);
            f.ShowInTaskbar = false;
            f.Show();
            Application.DoEvents();
            Application.DoEvents();
            using (Bitmap bmp = new Bitmap(f.Width, f.Height))
            {
                f.DrawToBitmap(bmp, new Rectangle(0, 0, f.Width, f.Height));
                string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "exclehelper-about.png");
                bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                f.Close();
                Xl.Say("关于窗口预览已保存：\n" + path);
            }
        }

        internal AboutForm()
        {
            Text = "关于 EXCLE助手";
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(560, 360);
            ShowInTaskbar = false;
            DoubleBuffered = true;
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint
                     | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
            Font = new Font("Microsoft YaHei UI", 9f);
            KeyPreview = true;
            KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.Escape) Close(); };
            Click += delegate { Close(); };
        }

        protected override void OnPaintBackground(PaintEventArgs e) { }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            Rectangle all = new Rectangle(0, 0, Width, Height);
            using (GraphicsPath p = Rounded(all, 14))
            using (SolidBrush b = new SolidBrush(Color.White))
            {
                g.FillPath(b, p);
                using (Pen pen = new Pen(Color.FromArgb(226, 232, 240)))
                    g.DrawPath(pen, p);
                Region = new Region(p);
            }

            Rectangle header = new Rectangle(0, 0, Width, 96);
            using (GraphicsPath p = Rounded(new Rectangle(-2, -2, Width + 4, 100), 16))
            using (LinearGradientBrush b = new LinearGradientBrush(header, BrandDark, BrandBlue, 20f))
                g.FillPath(b, p);

            using (Font tf = new Font("Microsoft YaHei UI", 19f, FontStyle.Bold))
            using (Font sf = new Font("Microsoft YaHei UI", 9f))
            using (SolidBrush w = new SolidBrush(Color.White))
            using (SolidBrush w2 = new SolidBrush(Color.FromArgb(205, 219, 234, 254)))
            {
                g.DrawString("EXCLE助手", tf, w, 26, 18);
                g.DrawString("表格效率工具箱  ·  单文件插件", sf, w2, 28, 56);
            }

            // 免费徽标
            Rectangle badge = new Rectangle(Width - 150, 24, 126, 30);
            using (GraphicsPath p = Rounded(badge, 14))
            using (SolidBrush b = new SolidBrush(Color.FromArgb(52, 255, 255, 255)))
                g.FillPath(b, p);
            using (Font bf = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold))
            using (SolidBrush bb = new SolidBrush(Color.White))
            {
                StringFormat f = new StringFormat();
                f.Alignment = StringAlignment.Center;
                f.LineAlignment = StringAlignment.Center;
                g.DrawString("全功能免费", bf, bb, badge, f);
            }

            int y = 116;
            using (Font kf = new Font("Microsoft YaHei UI", 9.5f, FontStyle.Bold))
            using (Font vf = new Font("Microsoft YaHei UI", 9.5f))
            using (SolidBrush kb = new SolidBrush(Color.FromArgb(15, 23, 42)))
            using (SolidBrush vb = new SolidBrush(Color.FromArgb(71, 85, 105)))
            {
                g.DrawString("版本", kf, kb, 28, y);
                g.DrawString("1.1.0   单 DLL / .NET Framework 4.8 / AnyCPU", vf, vb, 96, y); y += 26;
                g.DrawString("功能", kf, kb, 28, y);
                g.DrawString(Ops.Count + " 个操作，" + TabCount() + " 个功能选项卡，可搜索 · 可脚本调用", vf, vb, 96, y); y += 26;
                g.DrawString("架构", kf, kb, 28, y);
                g.DrawString("经典 COM 加载项 + 声明式 Ribbon + 操作注册表", vf, vb, 96, y); y += 26;
                g.DrawString("依赖", kf, kb, 28, y);
                g.DrawString("无（不需要 VSTO / PIA / Aspose / 联网）", vf, vb, 96, y); y += 26;
                g.DrawString("授权", kf, kb, 28, y);
                g.DrawString("无账户、无会员、无门禁、无计费接口", vf, vb, 96, y);
            }

            using (Pen pen = new Pen(Color.FromArgb(226, 232, 240)))
                g.DrawLine(pen, 24, Height - 56, Width - 24, Height - 56);

            using (Font ff = new Font("Microsoft YaHei UI", 8.5f))
            using (SolidBrush fb = new SolidBrush(Color.FromArgb(100, 116, 139)))
            {
                g.DrawString("点击任意处或按 Esc 关闭", ff, fb, 28, Height - 40);
                string right = "纯本地运算 · 不上传数据";
                SizeF rs = g.MeasureString(right, ff);
                g.DrawString(right, ff, fb, Width - rs.Width - 28, Height - 40);
            }
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
    }
}

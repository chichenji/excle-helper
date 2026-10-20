using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>
    /// 统一图标系统：按领域着色的圆角方块 + 白色首字，32×32 渲染后由 Office 缩放。
    /// 与功能面板的 chip 语言一致，保证「每个控件都有图标、且风格统一」，
    /// 不再依赖 imageMso（本机 51 个 id 里有 32 个是无效的，会渲染成空白）。
    /// </summary>
    internal static class Icons
    {
        private static readonly Dictionary<string, Bitmap> Cache = new Dictionary<string, Bitmap>();

        // 领域配色（同一类功能同一色，读起来成体系）
        private static readonly Color Blue = Color.FromArgb(37, 99, 235);     // 文本
        private static readonly Color Sky = Color.FromArgb(14, 165, 233);     // 提取 / 查找
        private static readonly Color Violet = Color.FromArgb(124, 58, 237);  // 数值
        private static readonly Color Teal = Color.FromArgb(13, 148, 136);    // 编辑
        private static readonly Color Pink = Color.FromArgb(219, 39, 119);    // 去重
        private static readonly Color Indigo = Color.FromArgb(79, 70, 229);   // 合并 / 拆分
        private static readonly Color Amber = Color.FromArgb(245, 158, 11);   // 分析
        private static readonly Color Emerald = Color.FromArgb(5, 150, 105);  // 格式
        private static readonly Color Orange = Color.FromArgb(217, 119, 6);   // 对象
        private static readonly Color Red = Color.FromArgb(220, 38, 38);      // 校验
        private static readonly Color Purple = Color.FromArgb(147, 51, 234);  // 中文专项
        private static readonly Color Slate = Color.FromArgb(71, 85, 105);    // 其他 / 界面
        private static readonly Color Cyan = Color.FromArgb(8, 145, 178);     // 日期时间
        private static readonly Color Green = Color.FromArgb(22, 163, 74);    // 录入验证
        private static readonly Color Stone = Color.FromArgb(87, 83, 78);     // 打印
        private static readonly Color Rose = Color.FromArgb(225, 29, 72);     // 图片
        private static readonly Color Brown = Color.FromArgb(146, 64, 14);    // 模板
        private static readonly Color Dark = Color.FromArgb(55, 65, 81);      // 安全
        private static readonly Color Gold = Color.FromArgb(202, 138, 4);     // 条件格式

        private static Color ColorOf(string tag)
        {
            if (string.IsNullOrEmpty(tag)) return Slate;
            if (tag.StartsWith("dt.", StringComparison.Ordinal)) return Cyan;
            if (tag.StartsWith("vd.", StringComparison.Ordinal)) return Green;
            if (tag.StartsWith("fm.", StringComparison.Ordinal)) return Slate;
            if (tag.StartsWith("pr.", StringComparison.Ordinal)) return Stone;
            if (tag.StartsWith("pic.", StringComparison.Ordinal)) return Rose;
            if (tag.StartsWith("st.", StringComparison.Ordinal)) return Amber;
            if (tag.StartsWith("cf.", StringComparison.Ordinal)) return Gold;
            if (tag.StartsWith("sec.", StringComparison.Ordinal)) return Dark;
            if (tag.StartsWith("tpl.", StringComparison.Ordinal)) return Brown;
            if (tag.StartsWith("t.ext", StringComparison.Ordinal)) return Sky;
            if (tag.StartsWith("t.", StringComparison.Ordinal)) return Blue;
            if (tag.StartsWith("n.", StringComparison.Ordinal)) return Violet;
            if (tag.StartsWith("d.dup", StringComparison.Ordinal)) return Pink;
            if (tag.StartsWith("d.merge", StringComparison.Ordinal)) return Indigo;
            if (tag.StartsWith("d.transpose", StringComparison.Ordinal)) return Indigo;
            if (tag.StartsWith("d.dim", StringComparison.Ordinal)) return Indigo;
            if (tag.StartsWith("d.subtotal", StringComparison.Ordinal)) return Amber;
            if (tag.StartsWith("d.sum", StringComparison.Ordinal)) return Amber;
            if (tag.StartsWith("d.rand", StringComparison.Ordinal)) return Amber;
            if (tag.StartsWith("d.cmp", StringComparison.Ordinal)) return Amber;
            if (tag.StartsWith("d.", StringComparison.Ordinal)) return Teal;
            if (tag.StartsWith("e.", StringComparison.Ordinal)) return Teal;
            if (tag.StartsWith("m.", StringComparison.Ordinal)) return Sky;
            if (tag.StartsWith("f.", StringComparison.Ordinal)) return Emerald;
            if (tag.StartsWith("o.", StringComparison.Ordinal)) return Orange;
            if (tag.StartsWith("v.", StringComparison.Ordinal)) return Red;
            if (tag.StartsWith("c.", StringComparison.Ordinal)) return Purple;
            if (tag.StartsWith("sm.", StringComparison.Ordinal)) return Indigo;
            if (tag.StartsWith("s.", StringComparison.Ordinal)) return Indigo;
            if (tag.StartsWith("w.", StringComparison.Ordinal)) return Indigo;
            return Slate;
        }

        private static Color Lighten(Color c, int delta)
        {
            return Color.FromArgb(
                Math.Min(255, c.R + delta),
                Math.Min(255, c.G + delta),
                Math.Min(255, c.B + delta));
        }

        private static GraphicsPath Rounded(Rectangle r, int radius)
        {
            GraphicsPath p = new GraphicsPath();
            int d = radius * 2;
            if (d > r.Height) d = r.Height;
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        /// <summary>32×32 品牌图标：领域色渐变圆角块 + 白色首字。</summary>
        /// <summary>32×32 图形图标（细线条，无文字）。</summary>
        internal static Bitmap Get(string label, string tag)
        {
            return Glyphs.Render(tag, 32);
        }
        /// <summary>把全部图标画成一张总览图，供设计评审。</summary>
        internal static string RenderSheet()
        {
            System.Collections.Generic.List<Feature> all = Catalog.All;
            int cols = 5, cellW = 262, cellH = 74;
            int rows = (int)Math.Ceiling(all.Count / (double)cols);
            int h = rows * cellH + 56;
            using (Bitmap bmp = new Bitmap(cols * cellW + 24, h))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                g.Clear(Color.White);
                using (Font hf = new Font("Microsoft YaHei UI", 14f, FontStyle.Bold))
                using (SolidBrush hb = new SolidBrush(Color.FromArgb(15, 23, 42)))
                    g.DrawString("EXCLE助手 · 图形图标总览（" + all.Count + " 个）", hf, hb, 16, 12);
                using (Font lf = new Font("Microsoft YaHei UI", 10f))
                using (Font gf2 = new Font("Microsoft YaHei UI", 7f))
                {
                    for (int k = 0; k < all.Count; k++)
                    {
                        Feature f = all[k];
                        int cx = 16 + (k % cols) * cellW;
                        int cy = 48 + (k / cols) * cellH;
                        try
                        {
                            using (Bitmap ic = Get(f.Label, f.Tag))
                                g.DrawImage(ic, new Rectangle(cx, cy + 6, 56, 56));
                        }
                        catch (Exception ex)
                        {
                            try
                            {
                                System.IO.File.AppendAllText(
                                    System.IO.Path.Combine(System.IO.Path.GetTempPath(), "exclehelper-ui.log"),
                                    "图标生成失败 tag=" + f.Tag + " → " + ex.ToString() + "\r\n");
                            }
                            catch { }
                        }
                        using (SolidBrush lb = new SolidBrush(Color.FromArgb(30, 41, 59)))
                            g.DrawString(f.Label, lf, lb, cx + 66, cy + 12);
                        using (SolidBrush gb2 = new SolidBrush(Color.FromArgb(148, 163, 184)))
                            g.DrawString(f.Tag, gf2, gb2, cx + 66, cy + 34);
                    }
                }
                string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "exclehelper-icons.png");
                bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                return path;
            }
        }
        /// <summary>把 Bitmap 转成 Office 需要的 IPictureDisp。</summary>
        internal static object ToPictureDisp(Image img)
        {
            return PictureHost.ToPictureDisp(img);
        }

        private sealed class PictureHost : AxHost
        {
            private PictureHost() : base("7B5C4E2A-1F3D-4A6B-9C8E-2D4F6A8B0C1E") { }
            internal static object ToPictureDisp(Image img)
            {
                return GetIPictureDispFromPicture(img);
            }
        }
    }
}

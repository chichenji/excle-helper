using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>
    /// 图片与图形域（批次 3）：批量插图、统一尺寸、对齐单元格、导出、删除、水印。
    /// 二维码/截图需要可扫码/可截屏的验证环境，暂缓（见计划文档）。
    /// </summary>
    internal static class PicOps
    {
        internal static void Register()
        {
            Ops.Reg("pic.insert", Insert);
            Ops.Reg("pic.size", Size);
            Ops.Reg("pic.anchor", Anchor);
            Ops.Reg("pic.export", Export);
            Ops.Reg("pic.del", Delete);
            Ops.Reg("pic.watermark", Watermark);
            Ops.Reg("pic.qrcode", QrCodeOp);
            Ops.Reg("pic.snip", Snip);
        }

        // ---------------- 7. 生成二维码 ----------------
        private static bool QrCodeOp()
        {
            dynamic r = Xl.SelRange("要生成二维码的单元格（取该单元格文本作为内容）");
            if (r == null) return true;
            dynamic cell = r.Cells[1, 1];
            string text = Xl.S(cell.Value2).Trim();
            if (text.Length == 0)
            {
                string input = Ask.One("生成二维码", "二维码内容", "");
                if (input == null) return true;
                text = input.Trim();
                if (text.Length == 0) return true;
            }

            int version = QrCode.PickVersion(System.Text.Encoding.UTF8.GetByteCount(text));
            if (version == 0)
            {
                Xl.Bad("内容过长：本编码器支持约 270 字节（中文约 90 字），请缩短后再试。");
                return true;
            }

            using (System.Drawing.Bitmap bmp = QrCode.Render(text, 6, 4))
            {
                if (bmp == null) { Xl.Bad("二维码生成失败。"); return true; }
                string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                                                     "exclehelper-qr-" + Guid.NewGuid().ToString("N") + ".png");
                bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                try
                {
                    double side = Math.Max(80, (double)cell.Height * 1.4);
                    double left = (double)cell.Left + ((double)cell.Width - side) / 2;
                    double top = (double)cell.Top + 1;
                    dynamic sh = cell.Worksheet.Shapes.AddPicture(path, false, true, left, top, side, side);
                    sh.Placement = 1;
                    Xl.Say("二维码已插入（版本 " + version + "，纠错等级 L）。\n内容："
                           + (text.Length > 40 ? text.Substring(0, 40) + "…" : text));
                }
                catch (Exception ex) { Xl.Bad("插入图片失败：" + ex.Message); }
            }
            return true;
        }

        // ---------------- 8. 截图插入 ----------------
        private static bool Snip()
        {
            if (!Xl.Ready()) return true;
            Xl.Say("3 秒后截取当前屏幕，并插入到活动单元格。\n请先切换到你需要的画面。");
            System.Threading.Thread.Sleep(3000);
            try
            {
                System.Drawing.Rectangle b = System.Windows.Forms.SystemInformation.VirtualScreen;
                using (System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(b.Width, b.Height))
                {
                    using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(bmp))
                        g.CopyFromScreen(b.Left, b.Top, 0, 0, bmp.Size);
                    string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                                                         "exclehelper-snip-" + Guid.NewGuid().ToString("N") + ".png");
                    bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                    dynamic cell = Xl.Sel;
                    double w = Math.Max(160, (double)cell.Width);
                    double h = w * bmp.Height / Math.Max(1, bmp.Width);
                    dynamic sh = cell.Worksheet.Shapes.AddPicture(path, false, true,
                                                                 (double)cell.Left + 1, (double)cell.Top + 1, w, h);
                    sh.Placement = 1;
                }
                Xl.Say("已插入屏幕截图（可拖动/缩放调整）。");
            }
            catch (Exception ex) { Xl.Bad("截图失败：" + ex.Message); }
            return true;
        }

        private static string PickFolder(string desc)
        {
            using (FolderBrowserDialog d = new FolderBrowserDialog())
            {
                d.Description = desc;
                d.ShowNewFolderButton = false;
                return d.ShowDialog() == DialogResult.OK ? d.SelectedPath : null;
            }
        }

        // ---------------- 1. 批量插图 ----------------
        private static bool Insert()
        {
            dynamic r = Xl.SelRange("含图片文件名（不含扩展名）的单元格区域");
            if (r == null) return true;
            string folder = PickFolder("选择图片所在文件夹");
            if (folder == null) return true;
            string[] exts = { ".png", ".jpg", ".jpeg", ".gif", ".bmp" };

            int rows = (int)r.Rows.Count, cols = (int)r.Columns.Count;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            dynamic ws = r.Worksheet;
            int ok = 0, miss = 0, err = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        dynamic cell = ws.Cells[firstRow + i, firstCol + j];
                        string name = Xl.S(cell.Value2).Trim();
                        if (name.Length == 0) continue;
                        string path = null;
                        foreach (string e in exts)
                        {
                            string p = Path.Combine(folder, name + e);
                            if (File.Exists(p)) { path = p; break; }
                            p = Path.Combine(folder, name);
                            if (File.Exists(p)) { path = p; break; }
                        }
                        if (path == null) { miss++; continue; }
                        try
                        {
                            double left = (double)cell.Left + 1;
                            double top = (double)cell.Top + 1;
                            double w = (double)cell.Width - 2;
                            double h = (double)cell.Height - 2;
                            dynamic sh = ws.Shapes.AddPicture(path, false, true, left, top, w, h);
                            sh.Placement = 1;    // xlMoveAndSize
                            ok++;
                        }
                        catch { err++; }
                    }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已插入 " + ok + " 张图片（按文件名匹配）"
                   + (miss > 0 ? "，" + miss + " 个未找到对应图片" : "")
                   + (err > 0 ? "，" + err + " 个插入失败" : "") + "。");
            return true;
        }

        // ---------------- 2. 统一尺寸 ----------------
        private static bool Size()
        {
            dynamic r = Xl.SelRange("任意单元格（作用于本表图片/形状）");
            if (r == null) return true;
            dynamic ws = r.Worksheet;
            string[] p = Ask.Show("统一图片尺寸",
                new string[] { "宽度（厘米，留空不改）", "高度（厘米，留空不改）", "保持比例(1/0)" },
                new string[] { "3", "", "1" });
            if (p == null) return true;
            double wcm = 0, hcm = 0;
            if (p[0].Trim().Length > 0 && !double.TryParse(p[0], NumberStyles.Any, CultureInfo.InvariantCulture, out wcm))
            { Xl.Bad("宽度请输入数字。"); return true; }
            if (p[1].Trim().Length > 0 && !double.TryParse(p[1], NumberStyles.Any, CultureInfo.InvariantCulture, out hcm))
            { Xl.Bad("高度请输入数字。"); return true; }
            bool keepRatio = p[2].Trim() != "0";
            if (wcm <= 0 && hcm <= 0) { Xl.Bad("宽度和高度至少填一个。"); return true; }

            int n = 0;
            Xl.Turn(true);
            try
            {
                int count = 0;
                try { count = (int)ws.Shapes.Count; } catch { }
                for (int i = 1; i <= count; i++)
                {
                    dynamic sh = null;
                    try { sh = ws.Shapes.Item(i); } catch { continue; }
                    int type = 0;
                    try { type = (int)sh.Type; } catch { }
                    if (type != 13 && type != 1) continue;   // 图片 / 自动形状
                    try
                    {
                        if (keepRatio)
                        {
                            if (wcm > 0) { double ratio = sh.Height / sh.Width; sh.Width = wcm * 28.3465; sh.Height = sh.Width * ratio; }
                            else { double ratio = sh.Width / sh.Height; sh.Height = hcm * 28.3465; sh.Width = sh.Height * ratio; }
                        }
                        else
                        {
                            if (wcm > 0) sh.Width = wcm * 28.3465;
                            if (hcm > 0) sh.Height = hcm * 28.3465;
                        }
                        n++;
                    }
                    catch { }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已调整 {n} 个对象尺寸。", n));
            return true;
        }

        // ---------------- 3. 对齐单元格 ----------------
        private static bool Anchor()
        {
            dynamic r = Xl.SelRange("要对齐的图片所在区域");
            if (r == null) return true;
            dynamic ws = r.Worksheet;
            int rows = (int)r.Rows.Count, cols = (int)r.Columns.Count;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;

            // 只处理完全落在该区域内的图片：按图片中心点归属单元格
            int n = 0;
            Xl.Turn(true);
            try
            {
                int count = 0;
                try { count = (int)ws.Shapes.Count; } catch { }
                for (int i = 1; i <= count; i++)
                {
                    dynamic sh = null;
                    try { sh = ws.Shapes.Item(i); } catch { continue; }
                    try
                    {
                        dynamic tl = null, br = null;
                        try { tl = sh.TopLeftCell; br = sh.BottomRightCell; } catch { }
                        if (tl == null || br == null) continue;
                        int tr = (int)tl.Row, tc = (int)tl.Column;
                        int brr = (int)br.Row, bcc = (int)br.Column;
                        // 与选区不相交则跳过，避免"区域外对象空扫"
                        if (brr < firstRow || tr > firstRow + rows - 1 || bcc < firstCol || tc > firstCol + cols - 1) continue;
                        dynamic hit = ws.Cells[tr, tc];
                        sh.Left = (double)hit.Left + 1;
                        sh.Top = (double)hit.Top + 1;
                        sh.Width = (double)hit.Width - 2;
                        sh.Height = (double)hit.Height - 2;
                        sh.Placement = 1;
                        n++;
                    }
                    catch { }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已把 {n} 个对象贴合到所在单元格。", n));
            return true;
        }

        // ---------------- 4. 导出图片 ----------------
        private static bool Export()
        {
            dynamic r = Xl.SelRange("图片所在区域");
            if (r == null) return true;
            string folder = PickFolder("选择图片导出文件夹");
            if (folder == null) return true;
            dynamic ws = r.Worksheet;
            int n = 0;
            Xl.Turn(true);
            try
            {
                int count = 0;
                try { count = (int)ws.Shapes.Count; } catch { }
                for (int i = 1; i <= count; i++)
                {
                    dynamic sh = null;
                    try { sh = ws.Shapes.Item(i); } catch { continue; }
                    int type = 0;
                    try { type = (int)sh.Type; } catch { }
                    if (type != 13) continue;
                    dynamic co = null;
                    try
                    {
                        sh.Copy();
                        co = ws.ChartObjects().Add(0, 0, 200, 200);
                        dynamic ch = co.Chart;
                        ch.Paste();
                        string name = (string)sh.Name;
                        string path = Path.Combine(folder, name + ".png");
                        ch.Export(path, "PNG");
                        n++;
                    }
                    catch { }
                    finally
                    {
                        try { if (co != null) co.Delete(); } catch { }   // 失败也回收，避免堆积空图表
                        try { Xl.App.CutCopyMode = 0; } catch { }
                    }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已导出 " + n + " 张图片到：\n" + folder);
            return true;
        }

        // ---------------- 5. 删除图片 ----------------
        private static bool Delete()
        {
            dynamic r = Xl.SelRange("任意单元格（作用于本表）");
            if (r == null) return true;
            dynamic ws = r.Worksheet;
            int count = 0, n = 0;
            try { count = (int)ws.Shapes.Count; } catch { }
            Xl.Turn(true);
            try
            {
                for (int i = count; i >= 1; i--)
                {
                    dynamic sh = null;
                    try { sh = ws.Shapes.Item(i); } catch { continue; }
                    int type = 0;
                    try { type = (int)sh.Type; } catch { }
                    if (type != 13) continue;
                    try { sh.Delete(); n++; } catch { }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(n == 0 ? "当前表没有图片。" : Xl.Fmt("已删除 {n} 张图片。", n));
            return true;
        }

        // ---------------- 6. 水印 ----------------
        private static bool Watermark()
        {
            dynamic r = Xl.SelRange("任意单元格（作用于本表）");
            if (r == null) return true;
            dynamic ws = r.Worksheet;
            string text = Ask.One("添加水印", "水印文字", "内部资料");
            if (text == null || text.Trim().Length == 0) return true;

            Xl.Turn(true);
            try
            {
                dynamic used = ws.UsedRange;
                double left = (double)used.Left;
                double top = (double)used.Top;
                double w = Math.Max((double)used.Width, 400);
                double h = Math.Max((double)used.Height, 200);

                dynamic sh = ws.Shapes.AddTextbox(1, left + w / 4, top + h / 4, w / 2, h / 2);
                sh.TextFrame.Characters().Text = text;
                sh.TextFrame.Characters().Font.Size = 48;
                sh.TextFrame.Characters().Font.Bold = true;
                sh.TextFrame.Characters().Font.Color = 0xC0C0C0;  // 浅灰
                sh.TextFrame.HorizontalAlignment = -4108;
                sh.TextFrame.VerticalAlignment = -4108;
                sh.Fill.Visible = 0;
                sh.Line.Visible = 0;
                sh.Rotation = 315;
                sh.Placement = 3;        // xlMove（不随单元格缩放）
                try { sh.Name = "_watermark"; } catch { }
                Xl.Say("水印已添加（浅灰、倾斜 45°、不遮挡单元格编辑）。\n需要删除时选中后按 Delete。");
            }
            catch (Exception ex) { Xl.Bad("添加水印失败：" + ex.Message); }
            finally { Xl.Turn(false); }
            return true;
        }
    }
}

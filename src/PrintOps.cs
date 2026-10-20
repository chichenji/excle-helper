using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>打印与页面设置域（批次 3）。</summary>
    internal static class PrintOps
    {
        internal static void Register()
        {
            Ops.Reg("pr.area", SetArea);
            Ops.Reg("pr.title", SetTitle);
            Ops.Reg("pr.pagenum", PageNum);
            Ops.Reg("pr.header", BatchHeader);
            Ops.Reg("pr.pagebreak", PageBreak);
            Ops.Reg("pr.fit", Fit);
            Ops.Reg("pr.center", Center);
            Ops.Reg("pr.grid", Grid);
            Ops.Reg("pr.batch", BatchPrint);
            Ops.Reg("pr.preview", PreviewAll);
        }

        private static bool Ready()
        {
            if (Xl.Book == null) { Xl.Bad("请先打开工作簿。"); return false; }
            return true;
        }

        private static bool SetArea()
        {
            dynamic r = Xl.SelRange("要作为打印区域的区域");
            if (r == null) return true;
            dynamic ws = r.Worksheet;
            ws.Activate();
            try { ws.PageSetup.PrintArea = (string)r.Address(true, true); }
            catch (Exception ex) { Xl.Bad("设置失败：" + ex.Message); return true; }
            Xl.Say("已把 " + Xl.Address(r) + " 设为打印区域。");
            return true;
        }

        private static bool SetTitle()
        {
            dynamic r = Xl.SelRange("要作为每页标题的行（整行）");
            if (r == null) return true;
            dynamic ws = r.Worksheet;
            int r1 = (int)r.Row, r2 = r1 + (int)r.Rows.Count - 1;
            ws.Activate();
            try { ws.PageSetup.PrintTitleRows = "$" + r1 + ":$" + r2; }
            catch (Exception ex) { Xl.Bad("设置失败：" + ex.Message); return true; }
            Xl.Say("已设置顶端标题行：第 " + r1 + "-" + r2 + " 行。");
            return true;
        }

        private static bool PageNum()
        {
            if (!Ready()) return true;
            dynamic ws = Xl.Act;
            string pos = Ask.One("统一页码", "位置：1=页脚居中  2=页脚右侧  3=页眉右侧", "1");
            if (pos == null) return true;
            string text = "第 &P 页 / 共 &N 页";
            Xl.Turn(true);
            try
            {
                ws.Activate();
                if (pos.Trim() == "2") ws.PageSetup.RightFooter = text;
                else if (pos.Trim() == "3") ws.PageSetup.RightHeader = text;
                else ws.PageSetup.CenterFooter = text;
            }
            catch (Exception ex) { Xl.Bad("设置失败：" + ex.Message); return true; }
            finally { Xl.Turn(false); }
            Xl.Say("已设置页码：" + text);
            return true;
        }

        private static bool BatchHeader()
        {
            if (!Ready()) return true;
            dynamic wb = Xl.Book;
            string[] p = Ask.Show("批量页眉页脚",
                new string[] { "页眉（留空不改）", "页脚（留空不改）" },
                new string[] { "", "" });
            if (p == null) return true;
            if (p[0].Length == 0 && p[1].Length == 0) return true;

            int n = (int)wb.Worksheets.Count, ok = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 1; i <= n; i++)
                {
                    dynamic ws = wb.Worksheets[i];
                    try
                    {
                        if (p[0].Length > 0) ws.PageSetup.CenterHeader = p[0];
                        if (p[1].Length > 0) ws.PageSetup.CenterFooter = p[1];
                        ok++;
                    }
                    catch { }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已为 " + ok + " 张工作表套用页眉页脚。");
            return true;
        }

        private static bool PageBreak()
        {
            dynamic r = Xl.SelRange("从哪一行开始分页");
            if (r == null) return true;
            string mode = Ask.One("插入分页符", "1 = 在选区上方插入一个分页符\n2 = 每隔 N 行插入（按选区范围）", "1");
            if (mode == null) return true;

            dynamic ws = r.Worksheet;
            int first = (int)r.Row;
            int rows = (int)r.Rows.Count;
            int n = 0;
            Xl.Turn(true);
            try
            {
                ws.Activate();
                (Xl.App.ActiveWindow).View = 2;       // xlPageBreakPreview
                if (mode.Trim() == "2")
                {
                    bool clear = Xl.Ask("插入前先清除本表已有的手工分页符？");
                    if (clear)
                    {
                        try { ws.ResetAllPageBreaks(); } catch { }
                    }
                    string s = Ask.One("间隔分页", "每隔几行分一页", "40");
                    if (s == null) { Xl.Turn(false); return true; }
                    int step;
                    if (!int.TryParse(s, out step) || step < 1) { Xl.Bad("请输入正整数。"); Xl.Turn(false); return true; }
                    for (int i = first + step; i < first + rows; i += step)
                    {
                        try { ws.HPageBreaks.Add(ws.Rows[i]); n++; } catch { }
                    }
                }
                else
                {
                    try { ws.HPageBreaks.Add(ws.Rows[first]); n++; } catch { }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已插入 {n} 个水平分页符。", n));
            return true;
        }

        private static bool Fit()
        {
            if (!Ready()) return true;
            dynamic wb = Xl.Book;
            bool all = Xl.Ask("应用到所有工作表？\n选「否」只处理当前表。");
            int n = all ? (int)wb.Worksheets.Count : 1;
            Xl.Turn(true);
            try
            {
                for (int i = 1; i <= n; i++)
                {
                    dynamic ws = all ? wb.Worksheets[i] : Xl.Act;
                    try
                    {
                        ws.PageSetup.Zoom = false;
                        ws.PageSetup.FitToPagesWide = 1;
                        ws.PageSetup.FitToPagesTall = false;
                    }
                    catch { }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已设置为「所有列缩放到一页宽」。");
            return true;
        }

        private static bool Center()
        {
            if (!Ready()) return true;
            dynamic ws = Xl.Act;
            string p = Ask.One("页面居中", "1 = 水平居中  2 = 垂直居中  3 = 两者都居中", "1");
            if (p == null) return true;
            bool ok = false;
            Xl.Turn(true);
            try
            {
                ws.Activate();
                ws.PageSetup.CenterHorizontally = (p.Trim() == "1" || p.Trim() == "3");
                ws.PageSetup.CenterVertically = (p.Trim() == "2" || p.Trim() == "3");
                ok = true;
            }
            catch (Exception ex) { Xl.Bad("设置失败：" + ex.Message); }
            finally { Xl.Turn(false); }
            if (ok) Xl.Say("页面居中已设置。");
            return true;
        }

        private static bool Grid()
        {
            if (!Ready()) return true;
            dynamic ws = Xl.Act;
            bool want = true, ok = false;
            Xl.Turn(true);
            try
            {
                ws.Activate();
                want = !(bool)ws.PageSetup.PrintGridlines;
                ws.PageSetup.PrintGridlines = want;
                ok = true;
            }
            catch (Exception ex) { Xl.Bad("设置失败：" + ex.Message); }
            finally { Xl.Turn(false); }
            if (ok) Xl.Say(want ? "已开启打印网格线。" : "已关闭打印网格线。");
            return true;
        }

        private static bool BatchPrint()
        {
            if (!Ready()) return true;
            dynamic wb = Xl.Book;
            int n = (int)wb.Worksheets.Count;
            string mode = Ask.One("批量输出",
                "1 = 每张工作表导出为一个 PDF\n2 = 整簿导出为一个 PDF\n3 = 直接打印全部工作表", "1");
            if (mode == null) return true;

            if (mode.Trim() == "3")
            {
                if (!Xl.Ask("将按当前打印机设置打印 " + n + " 张工作表，继续？")) return true;
                wb.PrintOut();
                Xl.Say("已提交打印任务。");
                return true;
            }

            string folder;
            using (FolderBrowserDialog d = new FolderBrowserDialog())
            {
                d.Description = "选择 PDF 输出文件夹";
                if (d.ShowDialog() != DialogResult.OK) return true;
                folder = d.SelectedPath;
            }

            Xl.Turn(true);
            int ok = 0;
            try
            {
                if (mode.Trim() == "2")
                {
                    string path = Path.Combine(folder, SafeName((string)wb.Name) + ".pdf");
                    wb.ExportAsFixedFormat(0, path);
                    ok = 1;
                }
                else
                {
                    for (int i = 1; i <= n; i++)
                    {
                        dynamic ws = wb.Worksheets[i];
                        try
                        {
                            string path = Path.Combine(folder, SafeName((string)ws.Name) + ".pdf");
                            ws.ExportAsFixedFormat(0, path);
                            ok++;
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex) { Xl.Bad("导出失败：" + ex.Message); }
            finally { Xl.Turn(false); }
            Xl.Say("已导出 " + ok + " 个 PDF 到：\n" + folder);
            return true;
        }

        private static string SafeName(string s)
        {
            foreach (char c in new char[] { ':', '\\', '/', '?', '*', '[', ']', '"', '<', '>', '|' })
                s = s.Replace(c, '_');
            return s;
        }

        private static bool PreviewAll()
        {
            if (!Ready()) return true;
            dynamic wb = Xl.Book;
            int n = (int)wb.Worksheets.Count;
            Xl.Turn(true);
            try
            {
                for (int i = 1; i <= n; i++)
                {
                    try
                    {
                        wb.Worksheets[i].Activate();
                        (Xl.App.ActiveWindow).View = 2;   // xlPageBreakPreview
                        dynamic ws = wb.Worksheets[i];
                        int pages = 1;
                        try { pages = (int)ws.PageSetup.Pages.Count; } catch { }
                        if (pages > 40) { Xl.Say("第 " + i + " 张表分页过多（" + pages + " 页），请检查缩放。"); }
                    }
                    catch { }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已切换到分页预览，可逐表检查分页情况。");
            return true;
        }
    }
}

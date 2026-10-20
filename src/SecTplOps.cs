using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>
    /// 批次 4 收尾：条件格式（cf.*）、安全与结构（sec.*）、简繁与全半角（cn.*、t.*）、
    /// 套打模板（tpl.*）。简繁转换走 Windows NLS（LCMapStringEx），不内置字表。
    /// </summary>
    internal static class SecTplOps
    {
        internal static void Register()
        {
            Ops.Reg("cf.value", CfValue);
            Ops.Reg("cf.bars", CfBars);
            Ops.Reg("cf.scale", CfScale);

            Ops.Reg("sec.structure", SecStructure);
            Ops.Reg("sec.hide.formula", SecHideFormula);
            Ops.Reg("sec.lock", SecLock);
            Ops.Reg("sec.unlink", SecUnlink);
            Ops.Reg("sec.meta", SecMeta);

            Ops.Reg("cn.t2s", ToSimplified);
            Ops.Reg("cn.s2t", ToTraditional);
            Ops.Reg("t.width.toggle", WidthToggle);

            Ops.Reg("tpl.label", TplLabel);
            Ops.Reg("tpl.payslip", TplPayslip);
            Ops.Reg("tpl.book", TplBook);
        }

        // ================= 条件格式 =================
        private static bool CfValue()
        {
            dynamic r = Xl.SelRange("要设置条件格式的区域");
            if (r == null) return true;
            string[] p = Ask.Show("条件标色",
                new string[] { "比较方式：1=大于 2=小于 3=等于 4=介于 5=重复值", "阈值（介于用逗号分隔两个值）", "底色：1=浅红 2=浅黄 3=浅绿 4=浅蓝" },
                new string[] { "1", "100", "1" });
            if (p == null) return true;
            int op, color;
            if (!int.TryParse(p[0], out op) || !int.TryParse(p[2], out color)) { Xl.Bad("请输入数字选项。"); return true; }
            int bg = color == 2 ? 0x9CF2FF : color == 3 ? 0xC6F0CE : color == 4 ? 0xF5E2CD : 0xC1C1FF;

            bool ok = false;
            Xl.Turn(true);
            try
            {
                r.FormatConditions.Delete();
                dynamic fc;
                if (op == 5)
                {
                    string addr = (string)r.Cells[1, 1].Address(true, true);
                    string f = "=COUNTIF(" + (string)r.Address(true, true) + "," + addr + ")>1";
                    fc = r.FormatConditions.Add(2, Type.Missing, f, Type.Missing);   // xlExpression: 表达式在 Formula1
                }
                else if (op == 1) fc = r.FormatConditions.Add(1, 5, "=" + p[1].Trim());
                else if (op == 2) fc = r.FormatConditions.Add(1, 6, "=" + p[1].Trim());
                else if (op == 3) fc = r.FormatConditions.Add(1, 3, "=" + p[1].Trim());
                else
                {
                    string[] two = p[1].Split(new char[] { ',', '，' });
                    if (two.Length < 2) { Xl.Bad("介于需要两个值。"); return true; }
                    fc = r.FormatConditions.Add(1, 1, "=" + two[0].Trim(), "=" + two[1].Trim());
                }
                fc.Interior.Color = bg;
                fc.Font.Bold = true;
                ok = true;
            }
            catch (Exception ex) { Xl.Bad("设置失败：" + ex.Message); }
            finally { Xl.Turn(false); }
            if (ok) Xl.Say("条件格式已应用。");
            return true;
        }

        private static bool CfBars()
        {
            dynamic r = Xl.SelRange("要加数据条的区域");
            if (r == null) return true;
            bool ok = false;
            Xl.Turn(true);
            try
            {
                r.FormatConditions.Delete();
                r.FormatConditions.AddDatabar();
                r.FormatConditions.Item(1).BarColor.Color = 0xEB6325;
                r.FormatConditions.Item(1).ShowValue = true;
                ok = true;
            }
            catch (Exception ex) { Xl.Bad("设置失败（需要 Excel 2010+）：" + ex.Message); }
            finally { Xl.Turn(false); }
            if (ok) Xl.Say("已添加数据条。");
            return true;
        }

        private static bool CfScale()
        {
            dynamic r = Xl.SelRange("要加色阶的区域");
            if (r == null) return true;
            bool ok = false;
            Xl.Turn(true);
            try
            {
                r.FormatConditions.Delete();
                r.FormatConditions.AddColorScale(3);
                dynamic cs = r.FormatConditions.Item(1);
                cs.ColorScaleCriteria.Item(1).Type = 1;   // xlConditionValueLowestValue
                cs.ColorScaleCriteria.Item(1).FormatColor.Color = 0xF8696B;
                cs.ColorScaleCriteria.Item(2).Type = 3;   // xlConditionValuePercent
                cs.ColorScaleCriteria.Item(2).Value = 50;
                cs.ColorScaleCriteria.Item(2).FormatColor.Color = 0xFFEB84;
                cs.ColorScaleCriteria.Item(3).Type = 2;   // xlConditionValueHighestValue
                cs.ColorScaleCriteria.Item(3).FormatColor.Color = 0x63BE7B;
                ok = true;
            }
            catch (Exception ex) { Xl.Bad("设置失败：" + ex.Message); }
            finally { Xl.Turn(false); }
            if (ok) Xl.Say("已添加三色阶。");
            return true;
        }

        // ================= 安全与结构 =================
        private static bool SecStructure()
        {
            if (Xl.Book == null) { Xl.Bad("请先打开工作簿。"); return true; }
            dynamic wb = Xl.Book;
            bool protectedNow;
            try { protectedNow = (bool)wb.ProtectStructure; } catch { protectedNow = false; }

            string pwd = Ask.One(protectedNow ? "取消结构保护" : "保护工作簿结构",
                                 "密码（可留空）", "");
            if (pwd == null) return true;
            try
            {
                if (protectedNow)
                {
                    if (pwd.Length == 0) wb.Unprotect();
                    else wb.Unprotect(pwd);
                    Xl.Say("已取消工作簿结构保护。");
                }
                else
                {
                    if (pwd.Length == 0) wb.Protect(Type.Missing, true, true);
                    else wb.Protect(pwd, true, true);
                    Xl.Say("已保护工作簿结构（禁止增删/重命名工作表）。");
                }
            }
            catch (Exception ex) { Xl.Bad("操作失败：" + ex.Message); }
            return true;
        }

        private static bool SecHideFormula()
        {
            dynamic r = Xl.SelRange("要隐藏公式的区域");
            if (r == null) return true;
            string pwd = Ask.One("隐藏公式", "保护密码（可留空）", "");
            if (pwd == null) return true;
            dynamic ws = r.Worksheet;
            Xl.Turn(true);
            try
            {
                r.FormulaHidden = true;
                r.Locked = true;
                ws.Activate();
                if (pwd.Length == 0) ws.Protect(Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                else ws.Protect(pwd, Type.Missing, Type.Missing, Type.Missing, true);
            }
            catch (Exception ex) { Xl.Bad("操作失败：" + ex.Message); }
            finally { Xl.Turn(false); }
            Xl.Say("已隐藏公式并保护工作表（公式仍在计算，只是不可见）。");
            return true;
        }

        private static bool SecLock()
        {
            dynamic r = Xl.SelRange("要锁定的区域");
            if (r == null) return true;
            bool protect = Xl.Ask("锁定后是否立即保护工作表？\n选「否」只标记锁定状态。");
            Xl.Turn(true);
            try
            {
                r.Locked = true;
                if (protect)
                {
                    dynamic ws = r.Worksheet;
                    ws.Activate();
                    // 参数位：5=UserInterfaceOnly，之后为各类 Allow*（全部保持默认 False）
                    ws.Protect(Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                }
            }
            catch (Exception ex) { Xl.Bad("操作失败：" + ex.Message); }
            finally { Xl.Turn(false); }
            Xl.Say(protect ? "已锁定并保护工作表。" : "已标记为锁定（保护工作表后生效）。");
            return true;
        }

        private static bool SecUnlink()
        {
            if (Xl.Book == null) { Xl.Bad("请先打开工作簿。"); return true; }
            dynamic wb = Xl.Book;
            if (!Xl.Ask("将断开所有外部链接并把外链公式转成当前值，继续？")) return true;
            int links = 0;
            Xl.Turn(true);
            try
            {
                object src = null;
                try { src = wb.LinkSources(1); } catch { }
                if (src is Array)
                {
                    Array arr = (Array)src;
                    for (int i = arr.Length; i >= 1; i--)
                    {
                        try { wb.BreakLink((string)arr.GetValue(i), 1); links++; }
                        catch { }
                    }
                }
                // 兜底：只把「含外链标记 [文件名]」的公式转成当前值，普通公式一律不动。
                // 用 SpecialCells 一次定位公式格，避免逐格 COM 往返。
                int converted = 0, noFormula = 0;
                int n = (int)wb.Worksheets.Count;
                for (int i = 1; i <= n; i++)
                {
                    dynamic ws = wb.Worksheets[i];
                    dynamic cells = null;
                    try { cells = ws.UsedRange.SpecialCells(-4123); }   // xlCellTypeFormulas
                    catch { noFormula++; continue; }
                    if (cells == null) { noFormula++; continue; }
                    int areas = 1;
                    try { areas = (int)cells.Areas.Count; } catch { }
                    for (int a = 1; a <= areas; a++)
                    {
                        dynamic area = (areas == 1) ? cells : cells.Areas[a];
                        try
                        {
                            Array fs = area.Formula as Array;
                            Array vs = area.Value2 as Array;
                            if (fs == null || vs == null) continue;
                            int lb0 = fs.GetLowerBound(0), lb1 = fs.GetLowerBound(1);
                            int rr = fs.GetUpperBound(0) - lb0 + 1;
                            int cc = fs.GetUpperBound(1) - lb1 + 1;
                            int r0 = (int)area.Row, c0 = (int)area.Column;
                            for (int x = 0; x < rr; x++)
                            {
                                for (int y = 0; y < cc; y++)
                                {
                                    string f = fs.GetValue(lb0 + x, lb1 + y) as string;
                                    if (f == null || f.IndexOf('[') < 0) continue;
                                    try
                                    {
                                        ws.Cells[r0 + x, c0 + y].Value2 = vs.GetValue(lb0 + x, lb1 + y);
                                        converted++;
                                    }
                                    catch { }
                                }
                            }
                        }
                        catch { }
                    }
                }
                Xl.Say("已断开 " + links + " 个外部链接，另有 " + converted + " 个外链公式已转为当前值。"
                       + (noFormula > 0 ? "\n（" + noFormula + " 张工作表没有公式，已跳过）" : ""));
                return true;
            }
            finally { Xl.Turn(false); }
        }

        private static bool SecMeta()
        {
            if (Xl.Book == null) { Xl.Bad("请先打开工作簿。"); return true; }
            dynamic wb = Xl.Book;
            if (!Xl.Ask("将清空作者、公司、标题等文档属性，继续？")) return true;
            string[] props = { "Author", "Company", "Title", "Subject", "Comments", "Keywords", "Manager", "Category", "LastAuthor" };
            int ok = 0;
            Xl.Turn(true);
            try
            {
                foreach (string k in props)
                {
                    try
                    {
                        wb.BuiltinDocumentProperties[k].Value = "";
                        ok++;
                    }
                    catch { }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已清空 " + ok + " 项文档属性。");
            return true;
        }

        // ================= 简繁 / 全半角（Windows NLS） =================
        private const uint LcmapHalfwidth = 0x00000004;
        private const uint LcmapFullwidth = 0x00000008;
        private const uint LcmapSimplified = 0x02000000;
        private const uint LcmapTraditional = 0x04000000;

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int LCMapStringEx(
            string lpLocaleName, uint dwMapFlags, string lpSrcStr, int cchSrc,
            [Out] char[] lpDestStr, int cchDest,
            IntPtr lpVersionInformation, IntPtr lpReserved, IntPtr sortHandle);

        private static string MapText(string s, uint flag, string locale)
        {
            if (string.IsNullOrEmpty(s)) return s;
            char[] buf = new char[s.Length * 2 + 2];
            int n = LCMapStringEx(locale, flag, s, s.Length, buf, buf.Length,
                                  IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            if (n <= 0) return s;
            return new string(buf, 0, Math.Min(n, buf.Length));
        }

        private static bool ToSimplified()
        {
            Xl.MapSel(delegate(string s) { return MapText(s, LcmapSimplified, "zh-CN"); }, "要转简体中文的区域");
            return true;
        }

        private static bool ToTraditional()
        {
            Xl.MapSel(delegate(string s) { return MapText(s, LcmapTraditional, "zh-TW"); }, "要转繁体中文的区域");
            return true;
        }

        private static bool WidthToggle()
        {
            string mode = Ask.One("全角/半角", "1 = 转全角   2 = 转半角", "2");
            if (mode == null) return true;
            uint flag = mode.Trim() == "1" ? LcmapFullwidth : LcmapHalfwidth;
            Xl.MapSel(delegate(string s) { return MapText(s, flag, ""); }, "要转换全半角的区域");
            return true;
        }

        // ================= 套打模板 =================
        private static bool TplLabel()
        {
            dynamic r = Xl.SelRange("单列标签数据");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列。"); return true; }
            string[] p = Ask.Show("标签套打排版",
                new string[] { "每行放几个", "每个标签之间空几列", "标题(可留空)" },
                new string[] { "3", "0", "" });
            if (p == null) return true;
            int perRow, gap;
            if (!int.TryParse(p[0], out perRow) || perRow < 1 || perRow > 20) { Xl.Bad("每行个数需在 1-20。"); return true; }
            if (!int.TryParse(p[1], out gap) || gap < 0) gap = 0;

            int totalCols = perRow + gap * (perRow - 1);
            int outRows = (rows + perRow - 1) / perRow;
            int startRow = p[2].Length > 0 ? 2 : 1;
            object[,] outv = new object[outRows + startRow - 1, totalCols];
            if (p[2].Length > 0) outv[0, 0] = p[2];
            for (int i = 0; i < rows; i++)
            {
                int rr = startRow - 1 + i / perRow;
                int cc = (i % perRow) * (1 + gap);
                outv[rr, cc] = a[i, 0];
            }
            dynamic sh = Xl.AddSheet("标签排版");
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("已排成 " + perRow + " 列标签 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        private static bool TplPayslip()
        {
            dynamic r = Xl.SelRange("含标题行的工资表");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (rows < 2) { Xl.Bad("至少需要标题行 + 一行数据。"); return true; }

            int outRows = rows * 2;
            object[,] outv = new object[outRows, cols];
            int k = 0;
            for (int i = 1; i < rows; i++)
            {
                for (int j = 0; j < cols; j++) outv[k, j] = a[0, j];
                k++;
                for (int j = 0; j < cols; j++) outv[k, j] = a[i, j];
                k++;
            }
            dynamic sh = Xl.AddSheet("工资条");
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("已生成 " + (rows - 1) + " 条工资条（标题+数据交替）→ 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        private static bool TplBook()
        {
            if (Xl.Book == null) { Xl.Bad("请先打开工作簿。"); return true; }
            dynamic wb = Xl.Book;
            string tpl = Ask.One("批量生成工作簿",
                "模板工作表名（留空=当前表）", "");
            if (tpl == null) return true;
            dynamic ws;
            try { ws = tpl.Trim().Length == 0 ? Xl.Act : wb.Worksheets[tpl.Trim()]; }
            catch { Xl.Bad("找不到工作表：" + tpl); return true; }

            dynamic data = Xl.PickRange("请选择【数据区域】（第一行是列名，列名对应模板里的 {列名}）");
            if (data == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(data, out rows, out cols);
            if (rows < 2) { Xl.Bad("数据区域至少需要标题行 + 一行。"); return true; }

            string folder;
            using (FolderBrowserDialog d = new FolderBrowserDialog())
            {
                d.Description = "选择生成的工作簿保存文件夹";
                if (d.ShowDialog() != DialogResult.OK) return true;
                folder = d.SelectedPath;
            }
            string keyCol = Ask.One("批量生成工作簿", "用哪一列作为文件名（第几列）", "1");
            if (keyCol == null) return true;
            int keyIdx;
            if (!int.TryParse(keyCol, out keyIdx) || keyIdx < 1 || keyIdx > cols) { Xl.Bad("列号无效。"); return true; }

            if (!Xl.Ask("将按 " + (rows - 1) + " 行数据生成 " + (rows - 1) + " 个工作簿，继续？")) return true;

            int ok = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 1; i < rows; i++)
                {
                    ws.Copy();
                    dynamic nb = Xl.App.ActiveWorkbook;
                    dynamic nws = nb.Worksheets[1];
                    try
                    {
                        dynamic ur = nws.UsedRange;
                        int rr = (int)ur.Rows.Count, cc = (int)ur.Columns.Count;
                        object[,] cells = Xl.Vals(ur, out rr, out cc);
                        int ur0 = (int)ur.Row, uc0 = (int)ur.Column;
                        for (int x = 0; x < rr; x++)
                        {
                            for (int y = 0; y < cc; y++)
                            {
                                string s = Xl.S(cells[x, y]);
                                if (s.IndexOf('{') < 0) continue;
                                string orig = s;
                                for (int h = 0; h < cols; h++)
                                {
                                    string token = "{" + Xl.S(a[0, h]) + "}";
                                    if (s.IndexOf(token, StringComparison.Ordinal) < 0) continue;
                                    s = s.Replace(token, Xl.S(a[i, h]));
                                }
                                if (s == orig) continue;
                                // 只回写真正发生替换的单元格，模板里的公式保持为公式
                                nws.Cells[ur0 + x, uc0 + y].Value2 = s;
                            }
                        }
                        string name = SafeName(Xl.S(a[i, keyIdx - 1]));
                        if (name.Length == 0) name = "book" + i;
                        nb.SaveAs(Path.Combine(folder, name + ".xlsx"));
                        nb.Close(false);
                        ok++;
                    }
                    catch { try { Xl.App.ActiveWorkbook.Close(false); } catch { } }
                }
            }
            catch (Exception ex) { Xl.Bad("生成失败：" + ex.Message); }
            finally { Xl.Turn(false); }
            ws.Activate();
            Xl.Say("已生成 " + ok + " 个工作簿到：\n" + folder);
            return true;
        }

        private static string SafeName(string s)
        {
            foreach (char c in new char[] { ':', '\\', '/', '?', '*', '[', ']', '"', '<', '>', '|' })
                s = s.Replace(c, '_');
            if (s.Length > 100) s = s.Substring(0, 100);
            return s.Trim();
        }
    }
}

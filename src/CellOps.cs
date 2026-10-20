using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcleHelper
{
    /// <summary>对象与校验域：批注、超链接、格式校验、重复行与空值标记。</summary>
    internal static class CellOps
    {
        internal static void Register()
        {
            Ops.Reg("o.comment.add", CommentAdd);
            Ops.Reg("o.comment.del", CommentDel);
            Ops.Reg("o.comment.export", CommentExport);
            Ops.Reg("o.link.add", LinkAdd);
            Ops.Reg("o.link.del", LinkDel);
            Ops.Reg("v.email.check", EmailCheck);
            Ops.Reg("v.null.highlight", NullHighlight);
            Ops.Reg("v.dup.row", DupRow);
            Ops.Reg("v.space.check", SpaceCheck);
        }

        // ---------------- comments ----------------
        private static bool CommentAdd()
        {
            dynamic r = Xl.SelRange("要加批注的区域");
            if (r == null) return true;
            string text = Ask.One("批量插入批注", "批注内容（{addr} 会替换为单元格地址）", "");
            if (text == null) return true;

            int rows = (int)r.Rows.Count, cols = (int)r.Columns.Count;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            dynamic ws = r.Worksheet;
            int ok = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        dynamic cell = ws.Cells[firstRow + i, firstCol + j];
                        string addr = (string)cell.Address(false, false);
                        string body = text.Replace("{addr}", addr);
                        try
                        {
                            if (cell.Comment != null) cell.Comment.Delete();
                            cell.AddComment(body);
                            ok++;
                        }
                        catch { }
                    }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已写入 {n} 条批注。", ok));
            return true;
        }

        private static bool CommentDel()
        {
            dynamic r = Xl.SelRange("要删除批注的区域");
            if (r == null) return true;
            try { r.ClearComments(); }
            catch { Xl.Bad("删除批注失败。"); return true; }
            Xl.Say("已清空选区批注。");
            return true;
        }

        private static bool CommentExport()
        {
            dynamic r = Xl.SelRange("要导出批注的区域");
            if (r == null) return true;
            int rows = (int)r.Rows.Count, cols = (int)r.Columns.Count;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            dynamic ws = r.Worksheet;
            List<object[]> rowsData = new List<object[]>();
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        dynamic cell = ws.Cells[firstRow + i, firstCol + j];
                        string c = "";
                        try { if (cell.Comment != null) c = (string)cell.Comment.Text(); }
                        catch { }
                        if (c.Length == 0) continue;
                        rowsData.Add(new object[] { (string)cell.Address(false, false), cell.Value2, c });
                    }
                }
            }
            finally { Xl.Turn(false); }
            if (rowsData.Count == 0) { Xl.Say("选区没有批注。"); return true; }
            object[,] outv = new object[rowsData.Count + 1, 3];
            outv[0, 0] = "单元格"; outv[0, 1] = "内容"; outv[0, 2] = "批注";
            for (int i = 0; i < rowsData.Count; i++)
                for (int j = 0; j < 3; j++) outv[i + 1, j] = rowsData[i][j];
            dynamic sh = Xl.AddSheet("批注导出");
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("已导出 " + rowsData.Count + " 条批注 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        // ---------------- hyperlinks ----------------
        private static bool LinkAdd()
        {
            dynamic r = Xl.SelRange("含网址/路径的单列区域");
            if (r == null) return true;
            int rows = (int)r.Rows.Count, cols = (int)r.Columns.Count;
            if (cols != 1) { Xl.Bad("请只选择一列。"); return true; }
            int firstRow = (int)r.Row;
            dynamic ws = r.Worksheet;
            int ok = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    dynamic cell = ws.Cells[firstRow + i, (int)r.Column];
                    string addr = Xl.S(cell.Value2).Trim();
                    if (addr.Length == 0) continue;
                    try
                    {
                        if (cell.Hyperlinks != null && (int)cell.Hyperlinks.Count > 0) continue;
                        ws.Hyperlinks.Add(cell, addr, Type.Missing, Type.Missing, addr);
                        ok++;
                    }
                    catch { }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已生成 {n} 个超链接。", ok));
            return true;
        }

        private static bool LinkDel()
        {
            dynamic r = Xl.SelRange("要删除超链接的区域");
            if (r == null) return true;
            int rows = (int)r.Rows.Count, cols = (int)r.Columns.Count;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            dynamic ws = r.Worksheet;
            int ok = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        dynamic cell = ws.Cells[firstRow + i, firstCol + j];
                        try
                        {
                            if (cell.Hyperlinks != null && (int)cell.Hyperlinks.Count > 0)
                            {
                                cell.Hyperlinks.Delete();
                                ok++;
                            }
                        }
                        catch { }
                    }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已删除 {n} 个超链接（保留文本）。", ok));
            return true;
        }

        // ---------------- validation / marking ----------------
        private static bool EmailCheck()
        {
            dynamic r = Xl.SelRange("含邮箱的单列区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列。"); return true; }
            Regex rx = new Regex(@"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$");
            object[,] outv = new object[rows, 1];
            int good = 0;
            for (int i = 0; i < rows; i++)
            {
                string s = Xl.S(a[i, 0]).Trim();
                if (s.Length == 0) { outv[i, 0] = ""; continue; }
                bool ok = rx.IsMatch(s);
                outv[i, 0] = ok ? "正确" : "错误";
                if (ok) good++;
            }
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, 1, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("邮箱校验完成：正确 " + good + " 个。");
            return true;
        }

        private static bool NullHighlight()
        {
            dynamic r = Xl.SelRange("要检查空值的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            dynamic ws = r.Worksheet;
            int n = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        if (Xl.S(a[i, j]).Trim().Length > 0) continue;
                        ws.Cells[firstRow + i, firstCol + j].Interior.Color = 65535; // 黄色
                        n++;
                    }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(n == 0 ? "没有空单元格。" : Xl.Fmt("已标记 {n} 个空单元格（黄色）。", n));
            return true;
        }

        private static bool DupRow()
        {
            dynamic r = Xl.SelRange("要检查重复行的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            dynamic ws = r.Worksheet;
            Dictionary<string, int> cnt = new Dictionary<string, int>(StringComparer.Ordinal);
            string[] keys = new string[rows];
            for (int i = 0; i < rows; i++)
            {
                StringBuilder sb = new StringBuilder();
                for (int j = 0; j < cols; j++) { sb.Append(Xl.S(a[i, j])); sb.Append('\u0001'); }
                keys[i] = sb.ToString();
                if (cnt.ContainsKey(keys[i])) cnt[keys[i]] = cnt[keys[i]] + 1;
                else cnt[keys[i]] = 1;
            }
            int n = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    if (cnt[keys[i]] <= 1) continue;
                    dynamic s1 = ws.Cells[firstRow + i, firstCol];
                    dynamic s2 = ws.Cells[firstRow + i, firstCol + cols - 1];
                    ws.Range(s1, s2).Interior.Color = 13551615; // 浅红
                    n++;
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(n == 0 ? "没有重复行。" : Xl.Fmt("已标色 {n} 个重复行。", n));
            return true;
        }

        private static bool SpaceCheck()
        {
            dynamic r = Xl.SelRange("要检查首尾空格的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            dynamic ws = r.Worksheet;
            int n = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        string s = Xl.S(a[i, j]);
                        if (s.Length == 0) continue;
                        if (s == s.Trim()) continue;
                        ws.Cells[firstRow + i, firstCol + j].Interior.Color = 65535;
                        n++;
                    }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(n == 0 ? "没有首尾空格问题。" : Xl.Fmt("已标记 {n} 个含首尾空格的单元格（黄色）。", n));
            return true;
        }
    }
}

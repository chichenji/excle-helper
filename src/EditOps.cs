using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcleHelper
{
    /// <summary>
    /// 编辑域：复制粘贴、清空与删除、合并转换、行列重排。
    /// </summary>
    internal static class EditOps
    {
        internal static void Register()
        {
            // 复制粘贴
            Ops.Reg("e.copy.to", CopyVisibleTo);
            Ops.Reg("e.copy.values", CopyValuesTo);
            Ops.Reg("e.copy.n", CopyTimes);
            Ops.Reg("e.copy.all", CopyAllWithSize);
            Ops.Reg("e.copy.sheets", CopyToSheets);
            Ops.Reg("e.copy.merge", CopyToMerged);
            // 清空与删除
            Ops.Reg("e.clear.unselected", ClearUnselected);
            Ops.Reg("e.clear.given", ClearGiven);
            Ops.Reg("e.del.blankcells", DelBlankCells);
            Ops.Reg("e.del.objects", DelObjects);
            Ops.Reg("e.del.textbox", DelTextBox);
            Ops.Reg("e.del.cf", DelConditionFormat);
            // 合并转换
            Ops.Reg("e.merge.rowcat", MergeRowConcat);
            Ops.Reg("e.merge.colcat", MergeColConcat);
            Ops.Reg("e.merge.blank", MergeBlankUp);
            Ops.Reg("e.flip.rows", FlipRows);
            Ops.Reg("e.combo", Combination);
            Ops.Reg("e.split.row", SplitRow);
        }

        // ---------------- 复制粘贴 ----------------
        private static bool CopyVisibleTo()
        {
            dynamic src = Xl.SelRange("要复制的区域");
            if (src == null) return true;
            dynamic dst = Xl.PickRange("请选择【粘贴目标】的左上角单元格");
            if (dst == null) return true;
            Xl.Turn(true);
            try
            {
                dynamic vis;
                try { vis = src.SpecialCells(12); }   // xlCellTypeVisible
                catch { Xl.Bad("选区没有可见单元格。"); return true; }
                vis.Copy(dst);
            }
            finally { Xl.Turn(false); }
            Xl.Say("已复制可见区域到 " + Xl.Address(dst) + "。");
            return true;
        }

        private static bool CopyValuesTo()
        {
            dynamic src = Xl.SelRange("要复制的区域");
            if (src == null) return true;
            dynamic dst = Xl.PickRange("请选择【只粘贴数值】的目标单元格");
            if (dst == null) return true;
            Xl.Turn(true);
            try
            {
                src.Copy();
                dst.PasteSpecial(-4163, -4142, false, false);   // xlPasteValues
                Xl.App.CutCopyMode = 0;
            }
            finally { Xl.Turn(false); }
            Xl.Say("已只粘贴数值到 " + Xl.Address(dst) + "。");
            return true;
        }

        private static bool CopyTimes()
        {
            dynamic src = Xl.SelRange("要复制的区域");
            if (src == null) return true;
            int n;
            if (!Ask.Int("复制 N 次", "向下复制几次", "3", out n)) return true;
            if (n < 1 || n > 500) { Xl.Bad("次数需在 1-500。"); return true; }
            int rows = (int)src.Rows.Count;
            dynamic ws = src.Worksheet;
            int firstRow = (int)src.Row, firstCol = (int)src.Column;
            int cols = (int)src.Columns.Count;
            Xl.Turn(true);
            try
            {
                for (int k = 1; k <= n; k++)
                {
                    dynamic dst = ws.Cells[firstRow + rows * k, firstCol];
                    src.Copy(dst);
                }
                Xl.App.CutCopyMode = 0;
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已复制 {n} 次，共 " + (rows * (n + 1)) + " 行。", n));
            return true;
        }

        private static bool CopyAllWithSize()
        {
            dynamic src = Xl.SelRange("要复制的区域");
            if (src == null) return true;
            dynamic dst = Xl.PickRange("请选择目标左上角单元格");
            if (dst == null) return true;

            int rows = (int)src.Rows.Count, cols = (int)src.Columns.Count;
            int dRow = (int)dst.Row, dCol = (int)dst.Column;
            dynamic ws = dst.Worksheet;
            int[] widths = new int[cols];
            int[] heights = new int[rows];
            Xl.Turn(true);
            try
            {
                for (int j = 0; j < cols; j++) widths[j] = (int)src.Columns[j + 1].ColumnWidth;
                for (int i = 0; i < rows; i++) heights[i] = (int)src.Rows[i + 1].RowHeight;
                src.Copy(dst);
                Xl.App.CutCopyMode = 0;
                for (int j = 0; j < cols; j++)
                {
                    if (widths[j] > 0) ws.Columns[dCol + j].ColumnWidth = widths[j];
                }
                for (int i = 0; i < rows; i++)
                {
                    if (heights[i] > 0) ws.Rows[dRow + i].RowHeight = heights[i];
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已连同行高列宽一起复制（" + rows + " 行 × " + cols + " 列）。");
            return true;
        }

        private static bool CopyToSheets()
        {
            dynamic src = Xl.SelRange("要复制的区域");
            if (src == null) return true;
            dynamic wb = Xl.Book;
            string spec = Ask.One("复制到多表", "目标工作表(逗号分隔，留空=除当前表外的全部)",
                                  "");
            if (spec == null) return true;

            List<string> names = new List<string>();
            if (spec.Trim().Length == 0)
            {
                string cur = (string)src.Worksheet.Name;
                int n = (int)wb.Worksheets.Count;
                for (int i = 1; i <= n; i++)
                {
                    string nm = (string)wb.Worksheets[i].Name;
                    if (nm != cur) names.Add(nm);
                }
            }
            else
            {
                foreach (string s in spec.Split(new char[] { ',', '，', '|' }))
                    if (s.Trim().Length > 0) names.Add(s.Trim());
            }

            int r = (int)src.Row, c = (int)src.Column;
            int ok = 0;
            Xl.Turn(true);
            try
            {
                foreach (string nm in names)
                {
                    dynamic t;
                    try { t = wb.Worksheets[nm]; } catch { continue; }
                    try
                    {
                        src.Copy(t.Cells[r, c]);
                        ok++;
                    }
                    catch { }
                }
                Xl.App.CutCopyMode = 0;
            }
            finally { Xl.Turn(false); }
            Xl.Say("已复制到 " + ok + " 张工作表的同一位置。");
            return true;
        }

        private static bool CopyToMerged()
        {
            dynamic src = Xl.SelRange("要填充的数据(单列)");
            if (src == null) return true;
            dynamic dst = Xl.PickRange("请选择【含合并单元格】的目标列");
            if (dst == null) return true;

            int rows, cols;
            object[,] a = Xl.Vals(src, out rows, out cols);
            if (cols != 1) { Xl.Bad("数据请只选一列。"); return true; }

            dynamic ws = dst.Worksheet;
            int firstRow = (int)dst.Row;
            int total = (int)dst.Rows.Count;
            int k = 0;
            int filled = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < total && k < rows; i++)
                {
                    dynamic cell = ws.Cells[firstRow + i, (int)dst.Column];
                    bool merged = false;
                    try { merged = (bool)cell.MergeCells; } catch { }
                    if (!merged) continue;
                    cell.Value2 = a[k, 0];
                    k++;
                    filled++;
                    i += (int)cell.MergeArea.Rows.Count - 1;
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已把 " + filled + " 个值填入合并单元格。");
            return true;
        }

        // ---------------- 清空与删除 ----------------
        private static bool ClearUnselected()
        {
            dynamic sel = Xl.SelRange("要保留的区域");
            if (sel == null) return true;
            dynamic ws = sel.Worksheet;
            dynamic ur = ws.UsedRange;
            if (!Xl.Ask("将清空「已用区域中选区之外」的内容（保留选区），继续？")) return true;

            int ur1 = (int)ur.Row, uc1 = (int)ur.Column;
            int ur2 = ur1 + (int)ur.Rows.Count - 1, uc2 = uc1 + (int)ur.Columns.Count - 1;
            int s1 = (int)sel.Row, sc1 = (int)sel.Column;
            int s2 = s1 + (int)sel.Rows.Count - 1, sc2 = sc1 + (int)sel.Columns.Count - 1;

            Xl.Turn(true);
            try
            {
                // 上方 / 下方 / 左侧 / 右侧 四块批量清空
                if (s1 > ur1) ws.Range(ws.Cells[ur1, uc1], ws.Cells[s1 - 1, uc2]).ClearContents();
                if (s2 < ur2) ws.Range(ws.Cells[s2 + 1, uc1], ws.Cells[ur2, uc2]).ClearContents();
                if (sc1 > uc1) ws.Range(ws.Cells[s1, uc1], ws.Cells[s2, sc1 - 1]).ClearContents();
                if (sc2 < uc2) ws.Range(ws.Cells[s1, sc2 + 1], ws.Cells[s2, uc2]).ClearContents();
            }
            finally { Xl.Turn(false); }
            Xl.Say("已清空选区之外的内容。");
            return true;
        }

        private static bool ClearGiven()
        {
            dynamic r = Xl.PickRange("请选择要清空内容的区域");
            if (r == null) return true;
            if (!Xl.Ask("将清空 " + Xl.Address(r) + " 的内容（保留格式），继续？")) return true;
            r.ClearContents();
            Xl.Say("已清空 " + Xl.Address(r) + "。");
            return true;
        }

        private static bool DelBlankCells()
        {
            dynamic r = Xl.SelRange("要删除空单元格的区域");
            if (r == null) return true;
            Xl.Turn(true);
            try
            {
                dynamic blanks;
                try { blanks = r.SpecialCells(4); }      // xlCellTypeBlanks
                catch { Xl.Say("选区没有空单元格。"); return true; }
                int n = (int)blanks.Count;
                blanks.Delete(-4162);                    // xlShiftUp
                Xl.Say(Xl.Fmt("已删除 {n} 个空单元格（下方内容上移）。", n));
            }
            finally { Xl.Turn(false); }
            return true;
        }

        private static bool DelObjects()
        {
            dynamic r = Xl.SelRange("任意单元格");
            if (r == null) return true;
            dynamic ws = r.Worksheet;
            int n = 0;
            try { n = (int)ws.Shapes.Count; } catch { }
            if (n == 0) { Xl.Say("当前工作表没有图形对象。"); return true; }
            if (!Xl.Ask("将删除当前工作表上的 " + n + " 个图形对象（图片/形状/图表框），继续？")) return true;
            Xl.Turn(true);
            try
            {
                for (int i = n; i >= 1; i--)
                {
                    try { ws.Shapes.Item(i).Delete(); } catch { }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已删除图形对象。");
            return true;
        }

        private static bool DelTextBox()
        {
            dynamic r = Xl.SelRange("任意单元格");
            if (r == null) return true;
            dynamic ws = r.Worksheet;
            int n = 0, ok = 0;
            try { n = (int)ws.Shapes.Count; } catch { }
            Xl.Turn(true);
            try
            {
                for (int i = n; i >= 1; i--)
                {
                    dynamic sh = null;
                    try { sh = ws.Shapes.Item(i); } catch { continue; }
                    int type = 0;
                    try { type = (int)sh.Type; } catch { }
                    if (type != 17) continue;    // msoTextBox
                    try { sh.Delete(); ok++; } catch { }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已删除 {n} 个文本框。", ok));
            return true;
        }

        private static bool DelConditionFormat()
        {
            dynamic r = Xl.SelRange("要删除条件格式的区域");
            if (r == null) return true;
            int n = 0;
            try { n = (int)r.FormatConditions.Count; } catch { }
            if (n == 0) { Xl.Say("选区没有条件格式。"); return true; }
            if (!Xl.Ask("将删除选区的 " + n + " 条条件格式，继续？")) return true;
            try { r.FormatConditions.Delete(); }
            catch { Xl.Bad("删除条件格式失败。"); return true; }
            Xl.Say("已删除条件格式。");
            return true;
        }

        // ---------------- 合并转换 ----------------
        private static bool MergeRowConcat()
        {
            dynamic r = Xl.SelRange("要按行合并的区域");
            if (r == null) return true;
            string sep = Ask.One("按行合并", "连接符(可留空)", "");
            if (sep == null) return true;

            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            dynamic ws = r.Worksheet;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    StringBuilder sb = new StringBuilder();
                    for (int j = 0; j < cols; j++)
                    {
                        string s = Xl.S(a[i, j]);
                        if (s.Length == 0) continue;
                        if (sb.Length > 0) sb.Append(sep);
                        sb.Append(s);
                    }
                    dynamic s1 = ws.Cells[firstRow + i, firstCol];
                    dynamic s2 = ws.Cells[firstRow + i, firstCol + cols - 1];
                    ws.Range(s1, s2).Merge();
                    s1.Value2 = sb.ToString();
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已按行合并并拼接内容。");
            return true;
        }

        private static bool MergeColConcat()
        {
            dynamic r = Xl.SelRange("要按列合并的区域");
            if (r == null) return true;
            string sep = Ask.One("按列合并", "连接符(可留空)", "");
            if (sep == null) return true;

            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            dynamic ws = r.Worksheet;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            Xl.Turn(true);
            try
            {
                for (int j = 0; j < cols; j++)
                {
                    StringBuilder sb = new StringBuilder();
                    for (int i = 0; i < rows; i++)
                    {
                        string s = Xl.S(a[i, j]);
                        if (s.Length == 0) continue;
                        if (sb.Length > 0) sb.Append(sep);
                        sb.Append(s);
                    }
                    dynamic s1 = ws.Cells[firstRow, firstCol + j];
                    dynamic s2 = ws.Cells[firstRow + rows - 1, firstCol + j];
                    ws.Range(s1, s2).Merge();
                    s1.Value2 = sb.ToString();
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已按列合并并拼接内容。");
            return true;
        }

        private static bool MergeBlankUp()
        {
            dynamic r = Xl.SelRange("要处理空值的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            dynamic ws = r.Worksheet;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            int merged = 0;
            Xl.Turn(true);
            try
            {
                for (int j = 0; j < cols; j++)
                {
                    for (int i = 1; i < rows; i++)
                    {
                        if (Xl.S(a[i, j]).Trim().Length > 0) continue;
                        if (Xl.S(a[i - 1, j]).Trim().Length == 0) continue;
                        dynamic s1 = ws.Cells[firstRow + i - 1, firstCol + j];
                        dynamic s2 = ws.Cells[firstRow + i, firstCol + j];
                        try { ws.Range(s1, s2).Merge(); merged++; } catch { }
                    }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已把 {n} 处空白并入上方单元格。", merged));
            return true;
        }

        private static bool FlipRows()
        {
            dynamic r = Xl.SelRange("要上下倒置的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (rows < 2) { Xl.Say("至少需要两行。"); return true; }
            object[,] outv = new object[rows, cols];
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                    outv[i, j] = a[rows - 1 - i, j];
            Xl.Turn(true);
            try { r.Value2 = outv; }
            finally { Xl.Turn(false); }
            Xl.Say("已把选区行的顺序上下倒置（仅保留值）。");
            return true;
        }

        private static bool Combination()
        {
            if (!Xl.Ready()) return true;
            dynamic a1 = Xl.PickRange("请选择【第一组】值（单列）");
            if (a1 == null) return true;
            dynamic a2 = Xl.PickRange("请选择【第二组】值（单列）");
            if (a2 == null) return true;
            string sep = Ask.One("排列组合", "连接符(留空则分两列输出)", "");
            if (sep == null) return true;

            int r1 = (int)a1.Rows.Count, c1 = (int)a1.Columns.Count;
            int r2 = (int)a2.Rows.Count, c2 = (int)a2.Columns.Count;
            if (c1 != 1 || c2 != 1) { Xl.Bad("请分别只选一列。"); return true; }
            int x, y;
            object[,] v1 = Xl.Vals(a1, out x, out y);
            object[,] v2 = Xl.Vals(a2, out x, out y);

            int total = r1 * r2;
            if (total > 200000) { Xl.Bad("组合结果过多（" + total + " 行）。"); return true; }
            object[,] outv = sep.Length > 0 ? new object[total, 1] : new object[total, 2];
            int k = 0;
            for (int i = 0; i < r1; i++)
            {
                for (int j = 0; j < r2; j++)
                {
                    string s1 = Xl.S(v1[i, 0]);
                    string s2 = Xl.S(v2[j, 0]);
                    if (sep.Length > 0) outv[k, 0] = s1 + sep + s2;
                    else { outv[k, 0] = s1; outv[k, 1] = s2; }
                    k++;
                }
            }
            dynamic sh = Xl.AddSheet("排列组合");
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("已生成 " + total + " 行组合 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        private static bool SplitRow()
        {
            dynamic r = Xl.SelRange("含标题行的数据区域");
            if (r == null) return true;
            string[] p = Ask.Show("整行拆分",
                new string[] { "按第几列拆分", "分隔符", "第一行是标题(1=是)" },
                new string[] { "1", ",", "1" });
            if (p == null) return true;
            int keyCol;
            if (!int.TryParse(p[0], out keyCol)) { Xl.Bad("列号必须是整数。"); return true; }
            string sep = p[1].Replace("\\t", "\t");
            if (sep.Length == 0) { Xl.Bad("分隔符不能为空。"); return true; }
            bool header = p[2].Trim() != "0";

            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (keyCol < 1 || keyCol > cols) { Xl.Bad("列号超出范围。"); return true; }

            List<object[]> outRows = new List<object[]>();
            int start = header ? 1 : 0;
            if (header)
            {
                object[] h = new object[cols];
                for (int j = 0; j < cols; j++) h[j] = a[0, j];
                outRows.Add(h);
            }
            for (int i = start; i < rows; i++)
            {
                string s = Xl.S(a[i, keyCol - 1]);
                string[] parts = s.Split(new string[] { sep }, StringSplitOptions.None);
                if (parts.Length <= 1)
                {
                    object[] one = new object[cols];
                    for (int j = 0; j < cols; j++) one[j] = a[i, j];
                    outRows.Add(one);
                    continue;
                }
                foreach (string part in parts)
                {
                    object[] row = new object[cols];
                    for (int j = 0; j < cols; j++) row[j] = a[i, j];
                    row[keyCol - 1] = part.Trim();
                    outRows.Add(row);
                }
            }
            object[,] outv = new object[outRows.Count, cols];
            for (int i = 0; i < outRows.Count; i++)
                for (int j = 0; j < cols; j++) outv[i, j] = outRows[i][j];
            dynamic sh = Xl.AddSheet("整行拆分");
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("已拆分为 " + outRows.Count + " 行 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcleHelper
{
    /// <summary>Table domain: deleting, clearing, inserting, de-duplicating, merging, reshaping, statistics.</summary>
    internal static class DataOps
    {
        // ---------------- shared ----------------
        private static bool Blank(object v)
        {
            return Xl.S(v).Trim().Length == 0;
        }

        private static void DeleteMask(dynamic r, bool[] keep, bool byRow)
        {
            dynamic sh = r.Worksheet;
            int first = byRow ? (int)r.Row : (int)r.Column;
            int i = keep.Length - 1;
            Xl.Turn(true);
            try
            {
                while (i >= 0)
                {
                    if (keep[i]) { i--; continue; }
                    int end = i;
                    while (i >= 0 && !keep[i]) i--;
                    int start = i + 1;
                    if (byRow)
                    {
                        dynamic a = sh.Rows[first + start];
                        dynamic b = sh.Rows[first + end];
                        sh.Range(a, b).Delete();
                    }
                    else
                    {
                        dynamic a = sh.Columns[first + start];
                        dynamic b = sh.Columns[first + end];
                        sh.Range(a, b).Delete();
                    }
                }
            }
            finally { Xl.Turn(false); }
        }

        private static void SaveToSheet(string sheetName, object[,] data, string note)
        {
            dynamic sh = Xl.AddSheet(sheetName);
            Xl.Dump(sh, 1, 1, data);
            Xl.Say(note + "\n已生成工作表「" + (string)sh.Name + "」。");
        }

        internal static void Register()
        {
            Ops.Reg("d.del.blankrows", () => DeleteBlank(true));
            Ops.Reg("d.del.blankcols", () => DeleteBlank(false));
            Ops.Reg("d.del.blankrowcells", DelRowsWithBlank);
            Ops.Reg("d.del.hidden", DelHidden);

            Ops.Reg("d.clear.zero", ClearZero);
            Ops.Reg("d.clear.nonnum", ClearNonNum);
            Ops.Reg("d.clear.err", ClearErr);

            Ops.Reg("d.ins.rows", InsRows);
            Ops.Reg("d.ins.cols", InsCols);
            Ops.Reg("d.del.every", DelEvery);

            Ops.Reg("d.dup.del", DupDel);
            Ops.Reg("d.dup.uniq", DupUniq);
            Ops.Reg("d.dup.highlight", DupHighlight);
            Ops.Reg("d.dup.count", DupCount);

            Ops.Reg("d.merge.same", MergeSame);
            Ops.Reg("d.merge.unsame", MergeUnsame);
            Ops.Reg("d.merge.keep", MergeKeep);

            Ops.Reg("d.transpose", Transpose);
            Ops.Reg("d.dim.flat", Flat);
            Ops.Reg("d.dim.matrix", Matrix);

            Ops.Reg("d.rand.num", RandNum);
            Ops.Reg("d.rand.text", RandText);
            Ops.Reg("d.rand.sort", RandSort);
            Ops.Reg("d.rand.pick", RandPick);

            Ops.Reg("d.subtotal", SubTotal);
            Ops.Reg("d.sum.color", SumColor);
            Ops.Reg("d.sum.blank", SumBlank);
            Ops.Reg("d.cmp.two", CmpTwo);
        }

        // ---------------- delete ----------------
        private static bool DeleteBlank(bool byRow)
        {
            dynamic r = Xl.SelRange(byRow ? "要检查空行的区域" : "要检查空列的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int n = byRow ? rows : cols;
            bool[] keep = new bool[n];
            int del = 0;
            for (int k = 0; k < n; k++)
            {
                bool blank = true;
                if (byRow) { for (int j = 0; j < cols; j++) if (!Blank(a[k, j])) { blank = false; break; } }
                else { for (int i = 0; i < rows; i++) if (!Blank(a[i, k])) { blank = false; break; } }
                keep[k] = !blank;
                if (blank) del++;
            }
            if (del == 0) { Xl.Say("没有空" + (byRow ? "行" : "列") + "。"); return true; }
            if (!Xl.Ask("将删除 " + del.ToString(CultureInfo.InvariantCulture) + " 个空"
                        + (byRow ? "行" : "列") + "，继续？")) return true;
            DeleteMask(r, keep, byRow);
            Xl.Say(Xl.Fmt("已删除 {n} 个空" + (byRow ? "行" : "列") + "。", del));
            return true;
        }

        private static bool DelRowsWithBlank()
        {
            dynamic r = Xl.SelRange("要检查的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            bool[] keep = new bool[rows];
            int del = 0;
            for (int i = 0; i < rows; i++)
            {
                bool hasBlank = false;
                for (int j = 0; j < cols; j++) if (Blank(a[i, j])) { hasBlank = true; break; }
                keep[i] = !hasBlank;
                if (hasBlank) del++;
            }
            if (del == 0) { Xl.Say("没有包含空单元格的行。"); return true; }
            if (!Xl.Ask("有 " + del.ToString(CultureInfo.InvariantCulture)
                        + " 行含空单元格，将整行删除。继续？")) return true;
            DeleteMask(r, keep, true);
            Xl.Say(Xl.Fmt("已删除 {n} 行。", del));
            return true;
        }

        private static bool DelHidden()
        {
            dynamic r = Xl.SelRange("要检查的区域");
            if (r == null) return true;
            int rows = (int)r.Rows.Count;
            int cols = (int)r.Columns.Count;
            bool[] keepR = new bool[rows];
            int delR = 0;
            for (int i = 0; i < rows; i++)
            {
                bool hidden;
                try { hidden = (bool)r.Rows[i + 1].Hidden; } catch { hidden = false; }
                keepR[i] = !hidden;
                if (hidden) delR++;
            }
            if (delR > 0)
            {
                if (!Xl.Ask("将删除 " + delR.ToString(CultureInfo.InvariantCulture) + " 个隐藏行，继续？")) return true;
                DeleteMask(r, keepR, true);
            }
            bool[] keepC = new bool[cols];
            int delC = 0;
            for (int j = 0; j < cols; j++)
            {
                bool hidden;
                try { hidden = (bool)r.Columns[j + 1].Hidden; } catch { hidden = false; }
                keepC[j] = !hidden;
                if (hidden) delC++;
            }
            if (delC > 0)
            {
                if (Xl.Ask("还发现 " + delC.ToString(CultureInfo.InvariantCulture) + " 个隐藏列，一并删除？"))
                    DeleteMask(r, keepC, false);
            }
            Xl.Say("隐藏行/列处理完成。");
            return true;
        }

        // ---------------- clear ----------------
        private static void MapClear(Func<object, bool> shouldClear, string what)
        {
            dynamic r = Xl.SelRange(what);
            if (r == null) return;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int n = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        object v = a[i, j];
                        if (Xl.IsFormula(v)) continue;
                        if (!shouldClear(v)) continue;
                        a[i, j] = null;
                        n++;
                    }
                }
                if (n > 0) r.Value2 = a;
            }
            finally { Xl.Turn(false); }
            if (n == 0) Xl.Say("没有符合条件的内容。");
            else Xl.Say(Xl.Fmt("已清空 {n} 个单元格。", n));
        }

        private static bool ClearZero()
        {
            MapClear(delegate(object v) { return Xl.IsNumber(v) && Math.Abs(Xl.D(v)) < 1e-12; }, "要清空 0 值的区域");
            return true;
        }

        private static bool ClearNonNum()
        {
            MapClear(delegate(object v)
            {
                if (Blank(v)) return false;
                return !Xl.IsNumber(v);
            }, "要清空非数字的区域");
            return true;
        }

        private static bool ClearErr()
        {
            dynamic r = Xl.SelRange("要清空错误值的区域");
            if (r == null) return true;
            try
            {
                dynamic cells = r.SpecialCells(2, 16); // constants, errors
                int n = (int)cells.Count;
                cells.ClearContents();
                Xl.Say(Xl.Fmt("已清空 {n} 个错误值。", n));
            }
            catch
            {
                Xl.Say("没有错误值。");
            }
            return true;
        }

        // ---------------- insert ----------------
        private static bool InsRows()
        {
            dynamic r = Xl.SelRange("插入位置的起始行");
            if (r == null) return true;
            int n;
            if (!Ask.Int("插入行", "插入几行", "1", out n)) return true;
            if (n < 1) { Xl.Bad("行数必须 >= 1。"); return true; }
            dynamic ws = r.Worksheet;
            dynamic a = ws.Rows[(int)r.Row];
            dynamic b = ws.Rows[(int)r.Row + n - 1];
            Xl.Turn(true);
            try { ws.Range(a, b).Insert(); }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已在第 " + (int)r.Row + " 行上方插入 {n} 行。", n));
            return true;
        }

        private static bool InsCols()
        {
            dynamic r = Xl.SelRange("插入位置的起始列");
            if (r == null) return true;
            int n;
            if (!Ask.Int("插入列", "插入几列", "1", out n)) return true;
            if (n < 1) { Xl.Bad("列数必须 >= 1。"); return true; }
            dynamic ws = r.Worksheet;
            dynamic a = ws.Columns[(int)r.Column];
            dynamic b = ws.Columns[(int)r.Column + n - 1];
            Xl.Turn(true);
            try { ws.Range(a, b).Insert(); }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已插入 {n} 列。", n));
            return true;
        }

        private static bool DelEvery()
        {
            dynamic r = Xl.SelRange("要处理的行区域");
            if (r == null) return true;
            int n;
            if (!Ask.Int("间隔删除行", "每隔几行删除 1 行（如 2 = 删第2、4、6…行）", "2", out n)) return true;
            if (n < 2) { Xl.Bad("间隔必须 >= 2。"); return true; }
            int rows = (int)r.Rows.Count;
            bool[] keep = new bool[rows];
            int del = 0;
            for (int i = 0; i < rows; i++)
            {
                bool drop = ((i + 1) % n) == 0;
                keep[i] = !drop;
                if (drop) del++;
            }
            if (del == 0) { Xl.Say("行数不足，无需删除。"); return true; }
            if (!Xl.Ask("将删除 " + del.ToString(CultureInfo.InvariantCulture) + " 行，继续？")) return true;
            DeleteMask(r, keep, true);
            Xl.Say(Xl.Fmt("已删除 {n} 行。", del));
            return true;
        }

        // ---------------- duplicates ----------------
        private static bool DupDel()
        {
            dynamic r = Xl.SelRange("要去重的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
            bool[] keep = new bool[rows];
            int del = 0;
            for (int i = 0; i < rows; i++)
            {
                StringBuilder key = new StringBuilder();
                for (int j = 0; j < cols; j++) { key.Append(Xl.S(a[i, j])); key.Append('\u0001'); }
                string k = key.ToString();
                if (seen.Contains(k)) { keep[i] = false; del++; }
                else { seen.Add(k); keep[i] = true; }
            }
            if (del == 0) { Xl.Say("没有重复行。"); return true; }
            DeleteMask(r, keep, true);
            Xl.Say(Xl.Fmt("已删除 {n} 个重复行。", del));
            return true;
        }

        private static bool DupUniq()
        {
            dynamic r = Xl.SelRange("要提取唯一值的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            List<string> order = new List<string>();
            HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    string s = Xl.S(a[i, j]);
                    if (s.Trim().Length == 0) continue;
                    if (seen.Add(s)) order.Add(s);
                }
            }
            if (order.Count == 0) { Xl.Say("没有可用数据。"); return true; }
            object[,] outv = new object[order.Count, 1];
            for (int i = 0; i < order.Count; i++) outv[i, 0] = order[i];
            SaveToSheet("唯一值", outv, "共 " + order.Count.ToString(CultureInfo.InvariantCulture) + " 个唯一值。");
            return true;
        }

        private static bool DupHighlight()
        {
            dynamic r = Xl.SelRange("要高亮重复值的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            Dictionary<string, int> cnt = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    string s = Xl.S(a[i, j]);
                    if (s.Trim().Length == 0) continue;
                    if (cnt.ContainsKey(s)) cnt[s] = cnt[s] + 1; else cnt[s] = 1;
                }
            int n = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        string s = Xl.S(a[i, j]);
                        if (s.Trim().Length == 0) continue;
                        if (cnt[s] <= 1) continue;
                        r.Cells[i + 1, j + 1].Interior.Color = 0x0000FF; // red
                        n++;
                    }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(n == 0 ? "没有重复值。" : Xl.Fmt("已高亮 {n} 个重复单元格。", n));
            return true;
        }

        private static bool DupCount()
        {
            dynamic r = Xl.SelRange("要统计重复次数的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            Dictionary<string, int> cnt = new Dictionary<string, int>(StringComparer.Ordinal);
            List<string> order = new List<string>();
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    string s = Xl.S(a[i, j]);
                    if (s.Trim().Length == 0) continue;
                    if (cnt.ContainsKey(s)) cnt[s] = cnt[s] + 1;
                    else { cnt[s] = 1; order.Add(s); }
                }
            if (order.Count == 0) { Xl.Say("没有可用数据。"); return true; }
            order.Sort(delegate(string x, string y)
            {
                int c = cnt[y].CompareTo(cnt[x]);
                return c != 0 ? c : string.CompareOrdinal(x, y);
            });
            object[,] outv = new object[order.Count + 1, 2];
            outv[0, 0] = "值"; outv[0, 1] = "出现次数";
            for (int i = 0; i < order.Count; i++) { outv[i + 1, 0] = order[i]; outv[i + 1, 1] = cnt[order[i]]; }
            SaveToSheet("重复统计", outv, "共统计 " + order.Count.ToString(CultureInfo.InvariantCulture) + " 个值。");
            return true;
        }

        // ---------------- merge ----------------
        private static bool MergeSame()
        {
            dynamic r = Xl.SelRange("要合并相同单元格的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            dynamic ws = r.Worksheet;
            int firstRow = (int)r.Row;
            int firstCol = (int)r.Column;
            int merged = 0;
            Xl.Turn(true);
            try
            {
                for (int j = 0; j < cols; j++)
                {
                    int i = 0;
                    while (i < rows)
                    {
                        string v = Xl.S(a[i, j]);
                        int k = i;
                        while (k + 1 < rows && Xl.S(a[k + 1, j]) == v) k++;
                        if (v.Trim().Length > 0 && k > i)
                        {
                            dynamic c1 = ws.Cells[firstRow + i, firstCol + j];
                            dynamic c2 = ws.Cells[firstRow + k, firstCol + j];
                            ws.Range(c1, c2).Merge();
                            merged++;
                        }
                        i = k + 1;
                    }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已合并 {n} 处相同单元格。", merged));
            return true;
        }

        private static bool MergeUnsame()
        {
            dynamic r = Xl.SelRange("要取消合并的区域");
            if (r == null) return true;
            dynamic ws = r.Worksheet;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            int rows = (int)r.Rows.Count, cols = (int)r.Columns.Count;
            int n = 0;
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
                            if (!(bool)cell.MergeCells) continue;
                            object v = cell.Value2;
                            dynamic area = cell.MergeArea;
                            int ar = (int)area.Rows.Count, ac = (int)area.Columns.Count;
                            area.UnMerge();
                            object[,] fill = new object[ar, ac];
                            for (int x = 0; x < ar; x++) for (int y = 0; y < ac; y++) fill[x, y] = v;
                            area.Value2 = fill;
                            n++;
                        }
                        catch { }
                    }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已拆解 {n} 处合并单元格并填充内容。", n));
            return true;
        }

        private static bool MergeKeep()
        {
            dynamic r = Xl.SelRange("要合并并拼接内容的区域");
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
                    int i = 0;
                    while (i < rows)
                    {
                        string v = Xl.S(a[i, j]);
                        int k = i;
                        while (k + 1 < rows && Xl.S(a[k + 1, j]) == v) k++;
                        if (v.Trim().Length > 0 && k > i)
                        {
                            List<string> vals = new List<string>();
                            for (int x = i; x <= k; x++)
                            {
                                string s = Xl.S(a[x, j]);
                                if (s.Length > 0 && !vals.Contains(s)) vals.Add(s);
                            }
                            dynamic c1 = ws.Cells[firstRow + i, firstCol + j];
                            dynamic c2 = ws.Cells[firstRow + k, firstCol + j];
                            dynamic rng = ws.Range(c1, c2);
                            rng.Merge();
                            c1.Value2 = string.Join("\n", vals.ToArray());
                            merged++;
                        }
                        i = k + 1;
                    }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已合并 {n} 处并保留内容。", merged));
            return true;
        }

        // ---------------- reshape ----------------
        private static bool Transpose()
        {
            dynamic r = Xl.SelRange("要转置的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            object[,] outv = new object[cols, rows];
            for (int i = 0; i < rows; i++) for (int j = 0; j < cols; j++) outv[j, i] = a[i, j];
            SaveToSheet("转置结果", outv, "已按 " + rows + "×" + cols + " → " + cols + "×" + rows + " 转置。");
            return true;
        }

        private static bool Flat()
        {
            dynamic r = Xl.SelRange("要一维化的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            List<object> list = new List<object>();
            if (rows == 1)
            {
                for (int j = 0; j < cols; j++) if (Xl.S(a[0, j]).Length > 0) list.Add(a[0, j]);
            }
            else
            {
                for (int j = 0; j < cols; j++)
                    for (int i = 0; i < rows; i++)
                        if (Xl.S(a[i, j]).Length > 0) list.Add(a[i, j]);
            }
            if (list.Count == 0) { Xl.Say("没有数据。"); return true; }
            object[,] outv = new object[list.Count, 1];
            for (int i = 0; i < list.Count; i++) outv[i, 0] = list[i];
            SaveToSheet("一维化", outv, "共 " + list.Count.ToString(CultureInfo.InvariantCulture) + " 个值。");
            return true;
        }

        private static bool Matrix()
        {
            dynamic r = Xl.SelRange("要二维化的一列数据");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            List<object> list = new List<object>();
            for (int i = 0; i < rows; i++) for (int j = 0; j < cols; j++) list.Add(a[i, j]);
            int per;
            if (!Ask.Int("一维转二维", "每行放几个值", "5", out per)) return true;
            if (per < 1) { Xl.Bad("每行个数必须 >= 1。"); return true; }
            int outRows = (list.Count + per - 1) / per;
            if (outRows < 1) outRows = 1;
            object[,] outv = new object[outRows, per];
            for (int i = 0; i < list.Count; i++) outv[i / per, i % per] = list[i];
            SaveToSheet("二维化", outv, "已排成 " + outRows + " 行 × " + per + " 列。");
            return true;
        }

        // ---------------- random ----------------
        private static bool RandNum()
        {
            dynamic r = Xl.SelRange("要填入随机数的区域");
            if (r == null) return true;
            string[] v = Ask.Show("生成随机数", new string[] { "最小值", "最大值", "是否整数(1/0)" },
                                   new string[] { "1", "100", "1" });
            if (v == null) return true;
            double lo, hi;
            if (!double.TryParse(v[0], NumberStyles.Any, CultureInfo.InvariantCulture, out lo) ||
                !double.TryParse(v[1], NumberStyles.Any, CultureInfo.InvariantCulture, out hi))
            {
                Xl.Bad("请输入数字。");
                return true;
            }
            bool isInt = v[2].Trim() != "0";
            if (hi < lo) { double t = lo; lo = hi; hi = t; }
            int rows, cols;
            Xl.Vals(r, out rows, out cols);
            Random rnd = new Random(Guid.NewGuid().GetHashCode());
            object[,] outv = new object[rows, cols];
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                    outv[i, j] = isInt
                        ? (object)rnd.Next((int)Math.Round(lo), (int)Math.Round(hi) + 1)
                        : (object)(lo + rnd.NextDouble() * (hi - lo));
            Xl.Turn(true);
            try { r.Value2 = outv; }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已生成 {n} 个随机数。", rows * cols));
            return true;
        }

        private static bool RandText()
        {
            dynamic r = Xl.SelRange("要填入随机文本的区域");
            if (r == null) return true;
            int len;
            if (!Ask.Int("生成随机文本", "每个单元格的长度", "8", out len)) return true;
            if (len < 1) { Xl.Bad("长度必须 >= 1。"); return true; }
            const string chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
            int rows, cols;
            Xl.Vals(r, out rows, out cols);
            Random rnd = new Random(Guid.NewGuid().GetHashCode());
            object[,] outv = new object[rows, cols];
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    StringBuilder sb = new StringBuilder(len);
                    for (int k = 0; k < len; k++) sb.Append(chars[rnd.Next(chars.Length)]);
                    outv[i, j] = sb.ToString();
                }
            }
            Xl.Turn(true);
            try { r.Value2 = outv; }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已生成 {n} 个随机文本。", rows * cols));
            return true;
        }

        private static bool RandSort()
        {
            dynamic r = Xl.SelRange("要随机排序的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            Random rnd = new Random(Guid.NewGuid().GetHashCode());
            int[] idx = new int[rows];
            for (int i = 0; i < rows; i++) idx[i] = i;
            for (int i = rows - 1; i > 0; i--)
            {
                int k = rnd.Next(i + 1);
                int t = idx[i]; idx[i] = idx[k]; idx[k] = t;
            }
            object[,] outv = new object[rows, cols];
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                    outv[i, j] = a[idx[i], j];
            Xl.Turn(true);
            try { r.Value2 = outv; }
            finally { Xl.Turn(false); }
            Xl.Say("已随机排序（仅保留值，原格式不再保留）。");
            return true;
        }

        private static bool RandPick()
        {
            dynamic r = Xl.SelRange("要抽样的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int n;
            if (!Ask.Int("随机抽取", "抽取多少行", "10", out n)) return true;
            if (n < 1 || n > rows) { Xl.Bad("抽取行数需在 1-" + rows + " 之间。"); return true; }
            int[] idx = new int[rows];
            for (int i = 0; i < rows; i++) idx[i] = i;
            Random rnd = new Random(Guid.NewGuid().GetHashCode());
            for (int i = rows - 1; i > 0; i--)
            {
                int k = rnd.Next(i + 1);
                int t = idx[i]; idx[i] = idx[k]; idx[k] = t;
            }
            object[,] outv = new object[n, cols];
            for (int i = 0; i < n; i++)
                for (int j = 0; j < cols; j++)
                    outv[i, j] = a[idx[i], j];
            SaveToSheet("随机抽取", outv, "已抽取 " + n.ToString(CultureInfo.InvariantCulture) + " 行。");
            return true;
        }

        // ---------------- statistics ----------------
        private static bool SubTotal()
        {
            dynamic r = Xl.SelRange("含标题行的数据区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (rows < 2) { Xl.Bad("至少需要标题行和一行数据。"); return true; }
            string[] v = Ask.Show("分类汇总", new string[] { "分类列(第几列)", "汇总列(第几列)" },
                                   new string[] { "1", "2" });
            if (v == null) return true;
            int keyCol, valCol;
            if (!int.TryParse(v[0], out keyCol) || !int.TryParse(v[1], out valCol)) { Xl.Bad("请输入列号。"); return true; }
            if (keyCol < 1 || keyCol > cols || valCol < 1 || valCol > cols) { Xl.Bad("列号超出范围。"); return true; }

            List<string> keys = new List<string>();
            List<double> sums = new List<double>();
            List<int> counts = new List<int>();
            string cur = null;
            for (int i = 1; i < rows; i++)
            {
                string k = Xl.S(a[i, keyCol - 1]);
                if (cur != k)
                {
                    if (cur != null) { keys.Add(cur); sums.Add(0); counts.Add(0); }
                    cur = k;
                }
                if (keys.Count == 0) { keys.Add(cur); sums.Add(0); counts.Add(0); }
                int li = keys.Count - 1;
                if (Xl.IsNumber(a[i, valCol - 1])) sums[li] = sums[li] + Xl.D(a[i, valCol - 1]);
                counts[li] = counts[li] + 1;
            }
            if (keys.Count == 0) { Xl.Say("没有可汇总的数据。"); return true; }

            object[,] outv = new object[keys.Count + 1, 3];
            outv[0, 0] = Xl.S(a[0, keyCol - 1]);
            outv[0, 1] = "汇总(" + Xl.S(a[0, valCol - 1]) + ")";
            outv[0, 2] = "行数";
            for (int i = 0; i < keys.Count; i++)
            {
                outv[i + 1, 0] = keys[i];
                outv[i + 1, 1] = sums[i];
                outv[i + 1, 2] = counts[i];
            }
            SaveToSheet("分类汇总", outv, "已按相邻分组汇总。");
            return true;
        }

        private static bool SumColor()
        {
            dynamic r = Xl.SelRange("要按颜色统计的区域");
            if (r == null) return true;
            int rows, cols;
            Xl.Vals(r, out rows, out cols);
            if (rows * cols > 50000)
            {
                Xl.Bad("区域过大（" + (rows * cols) + " 格），请缩小范围。");
                return true;
            }
            Dictionary<int, double> sums = new Dictionary<int, double>();
            Dictionary<int, int> counts = new Dictionary<int, int>();
            List<int> order = new List<int>();
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        dynamic cell = r.Cells[i + 1, j + 1];
                        int color = 0;
                        try { color = (int)cell.Interior.Color; } catch { }
                        if (!sums.ContainsKey(color)) { sums[color] = 0; counts[color] = 0; order.Add(color); }
                        counts[color] = counts[color] + 1;
                        object v = cell.Value2;
                        if (Xl.IsNumber(v)) sums[color] = sums[color] + Xl.D(v);
                    }
                }
            }
            finally { Xl.Turn(false); }
            object[,] outv = new object[order.Count + 1, 4];
            outv[0, 0] = "背景色"; outv[0, 1] = "数值合计"; outv[0, 2] = "单元格数"; outv[0, 3] = "色值";
            for (int i = 0; i < order.Count; i++)
            {
                int c = order[i];
                outv[i + 1, 0] = c == 0 || c == 0xFFFFFF ? "无填充" : "";
                outv[i + 1, 1] = sums[c];
                outv[i + 1, 2] = counts[c];
                outv[i + 1, 3] = c;
            }
            dynamic sh = Xl.AddSheet("颜色统计");
            Xl.Dump(sh, 1, 1, outv);
            for (int i = 0; i < order.Count; i++)
            {
                int c = order[i];
                if (c == 0 || c == 0xFFFFFF) continue;
                try { sh.Cells[i + 2, 1].Interior.Color = c; } catch { }
            }
            Xl.Say("已生成工作表「" + (string)sh.Name + "」，共 " + order.Count.ToString(CultureInfo.InvariantCulture) + " 种颜色。");
            return true;
        }

        private static bool SumBlank()
        {
            dynamic r = Xl.SelRange("要求和的单列区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列。"); return true; }
            int filled = 0;
            int i = 0;
            while (i < rows)
            {
                if (!Xl.IsNumber(a[i, 0])) { i++; continue; }
                double sum = 0;
                int k = i;
                while (k < rows && Xl.IsNumber(a[k, 0])) { sum += Xl.D(a[k, 0]); k++; }
                if (k < rows && Blank(a[k, 0]))
                {
                    a[k, 0] = sum;
                    filled++;
                    i = k + 1;
                }
                else i = k;
            }
            if (filled == 0) { Xl.Say("没有找到「数字块 + 空白」的组合。"); return true; }
            Xl.Turn(true);
            try { r.Value2 = a; }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已在 {n} 个空白格填入上方数字之和。", filled));
            return true;
        }

        private static bool CmpTwo()
        {
            if (!Xl.Ready()) return true;
            dynamic a1 = Xl.PickRange("请选择【区域1】");
            if (a1 == null) return true;
            dynamic a2 = Xl.PickRange("请选择【区域2】（大小需与区域1 相同）");
            if (a2 == null) return true;
            int r1 = (int)a1.Rows.Count, c1 = (int)a1.Columns.Count;
            int r2 = (int)a2.Rows.Count, c2 = (int)a2.Columns.Count;
            if (r1 != r2 || c1 != c2)
            {
                Xl.Bad("两个区域大小不一致：区域1 为 " + r1 + "×" + c1 + "，区域2 为 " + r2 + "×" + c2 + "。");
                return true;
            }
            int rows, cols;
            object[,] x = Xl.Vals(a1, out rows, out cols);
            object[,] y = Xl.Vals(a2, out rows, out cols);
            List<object[]> diffs = new List<object[]>();
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    if (string.Equals(Xl.S(x[i, j]), Xl.S(y[i, j]), StringComparison.Ordinal)) continue;
                    diffs.Add(new object[] { i + 1, j + 1, x[i, j], y[i, j] });
                }
            if (diffs.Count == 0) { Xl.Say("两个区域完全一致。"); return true; }
            object[,] outv = new object[diffs.Count + 1, 4];
            outv[0, 0] = "行"; outv[0, 1] = "列"; outv[0, 2] = "区域1"; outv[0, 3] = "区域2";
            for (int i = 0; i < diffs.Count; i++)
                for (int j = 0; j < 4; j++) outv[i + 1, j] = diffs[i][j];
            SaveToSheet("对比结果", outv, "共 " + diffs.Count.ToString(CultureInfo.InvariantCulture) + " 处不同。");
            return true;
        }
    }
}

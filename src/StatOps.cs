using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcleHelper
{
    /// <summary>
    /// 统计与分析域（批次 4）：占比、环比、同比、排名、描述统计、频数分布、分组汇总、相关系数。
    /// </summary>
    internal static class StatOps
    {
        internal static void Register()
        {
            Ops.Reg("st.share", Share);
            Ops.Reg("st.mom", Mom);
            Ops.Reg("st.yoy", Yoy);
            Ops.Reg("st.rank", Rank);
            Ops.Reg("st.describe", Describe);
            Ops.Reg("st.freq", Freq);
            Ops.Reg("st.group", Group);
            Ops.Reg("st.corr", Corr);
        }

        private static void AppendColumn(dynamic r, object[,] data, string fmt)
        {
            int cols = (int)r.Columns.Count;
            TextOps.WriteBlock(r, 0, cols, data);
            if (fmt != null)
            {
                dynamic ws = r.Worksheet;
                int firstRow = (int)r.Row;
                int rowCount = data.GetLength(0);
                dynamic c1 = ws.Cells[firstRow, (int)r.Column + cols];
                dynamic c2 = ws.Cells[firstRow + rowCount - 1, (int)r.Column + cols];
                try { ws.Range(c1, c2).NumberFormat = fmt; } catch { }
            }
        }

        // ---------------- 占比 ----------------
        private static bool Share()
        {
            dynamic r = Xl.SelRange("数值列");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列数值。"); return true; }
            double total = 0;
            for (int i = 0; i < rows; i++) if (Xl.IsNumber(a[i, 0])) total += Xl.D(a[i, 0]);
            if (Math.Abs(total) < 1e-12) { Xl.Bad("合计为 0，无法计算占比。"); return true; }
            object[,] outv = new object[rows, 1];
            for (int i = 0; i < rows; i++)
                outv[i, 0] = Xl.IsNumber(a[i, 0]) ? (object)(Xl.D(a[i, 0]) / total) : null;
            Xl.Turn(true);
            try { AppendColumn(r, outv, "0.00%"); }
            finally { Xl.Turn(false); }
            Xl.Say("已计算占比（写右侧，百分比格式）。");
            return true;
        }

        // ---------------- 环比 ----------------
        private static bool Mom()
        {
            dynamic r = Xl.SelRange("按时间顺序排列的数值列");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列数值。"); return true; }
            object[,] outv = new object[rows, 1];
            for (int i = 1; i < rows; i++)
            {
                if (!Xl.IsNumber(a[i, 0]) || !Xl.IsNumber(a[i - 1, 0])) { outv[i, 0] = null; continue; }
                double cur = Xl.D(a[i, 0]), prev = Xl.D(a[i - 1, 0]);
                outv[i, 0] = Math.Abs(prev) < 1e-12 ? (object)null : (cur - prev) / prev;
            }
            Xl.Turn(true);
            try { AppendColumn(r, outv, "0.00%"); }
            finally { Xl.Turn(false); }
            Xl.Say("已计算环比（本期/上期-1，写右侧）。");
            return true;
        }

        // ---------------- 同比 ----------------
        private static bool Yoy()
        {
            dynamic r = Xl.SelRange("数值列");
            if (r == null) return true;
            int lag;
            if (!Ask.Int("同比", "与多少期之前对比（年度数据通常填 12）", "12", out lag)) return true;
            if (lag < 1) { Xl.Bad("期数必须 >= 1。"); return true; }
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列数值。"); return true; }
            object[,] outv = new object[rows, 1];
            for (int i = lag; i < rows; i++)
            {
                if (!Xl.IsNumber(a[i, 0]) || !Xl.IsNumber(a[i - lag, 0])) { outv[i, 0] = null; continue; }
                double cur = Xl.D(a[i, 0]), prev = Xl.D(a[i - lag, 0]);
                outv[i, 0] = Math.Abs(prev) < 1e-12 ? (object)null : (cur - prev) / prev;
            }
            Xl.Turn(true);
            try { AppendColumn(r, outv, "0.00%"); }
            finally { Xl.Turn(false); }
            Xl.Say("已计算同比（与 " + lag + " 期前对比，写右侧）。");
            return true;
        }

        // ---------------- 排名 ----------------
        private static bool Rank()
        {
            dynamic r = Xl.SelRange("数值列");
            if (r == null) return true;
            string mode = Ask.One("排名", "1 = 从大到小（降序）  2 = 从小到大（升序）", "1");
            if (mode == null) return true;
            bool desc = mode.Trim() != "2";
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列数值。"); return true; }

            List<double> vals = new List<double>();
            for (int i = 0; i < rows; i++) if (Xl.IsNumber(a[i, 0])) vals.Add(Xl.D(a[i, 0]));
            vals.Sort();
            if (desc) vals.Reverse();   // desc = 从大到小：升序后需反转

            object[,] outv = new object[rows, 2];
            for (int i = 0; i < rows; i++)
            {
                if (!Xl.IsNumber(a[i, 0])) { outv[i, 0] = null; outv[i, 1] = null; continue; }
                double v = Xl.D(a[i, 0]);
                int idx = vals.IndexOf(v);
                int sameCount = 0;
                foreach (double x in vals) if (Math.Abs(x - v) < 1e-12) sameCount++;
                outv[i, 0] = idx + 1;
                outv[i, 1] = sameCount > 1 ? "并列" : "";
            }
            Xl.Turn(true);
            try { AppendColumn(r, outv, null); }
            finally { Xl.Turn(false); }
            Xl.Say("已计算排名（并列取最小名次，写右侧）。");
            return true;
        }

        // ---------------- 描述统计 ----------------
        private static bool Describe()
        {
            dynamic r = Xl.SelRange("数值列");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列数值。"); return true; }

            List<double> v = new List<double>();
            for (int i = 0; i < rows; i++) if (Xl.IsNumber(a[i, 0])) v.Add(Xl.D(a[i, 0]));
            if (v.Count == 0) { Xl.Bad("没有数值。"); return true; }
            v.Sort();
            double sum = 0;
            foreach (double x in v) sum += x;
            double mean = sum / v.Count;
            double median = Percentile(v, 0.5);
            double variance = 0;
            foreach (double x in v) variance += (x - mean) * (x - mean);
            variance = v.Count > 1 ? variance / (v.Count - 1) : 0;

            object[,] outv = new object[11, 2];
            string[] names = { "计数", "合计", "均值", "中位数", "标准差", "最小值", "下四分位", "上四分位", "最大值", "极差", "变异系数" };
            double sd = Math.Sqrt(variance);
            object[] values =
            {
                v.Count, sum, mean, median, sd, v[0],
                Percentile(v, 0.25), Percentile(v, 0.75), v[v.Count - 1],
                v[v.Count - 1] - v[0],
                Math.Abs(mean) < 1e-12 ? 0d : sd / mean
            };
            for (int i = 0; i < names.Length; i++) { outv[i, 0] = names[i]; outv[i, 1] = values[i]; }
            dynamic sh = Xl.AddSheet("描述统计");
            Xl.Dump(sh, 1, 1, outv);
            try { sh.Columns[1].ColumnWidth = 12; sh.Columns[2].ColumnWidth = 16; } catch { }
            Xl.Say("已生成描述统计（" + v.Count + " 个样本）→ 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        private static double Percentile(List<double> sorted, double p)
        {
            if (sorted.Count == 0) return 0;
            double pos = p * (sorted.Count - 1);
            int lo = (int)Math.Floor(pos);
            int hi = (int)Math.Ceiling(pos);
            if (lo == hi) return sorted[lo];
            return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
        }

        // ---------------- 频数分布 ----------------
        private static bool Freq()
        {
            dynamic r = Xl.SelRange("数值列");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列数值。"); return true; }
            int bins;
            if (!Ask.Int("频数分布", "分成几组", "10", out bins)) return true;
            if (bins < 2 || bins > 100) { Xl.Bad("组数需在 2-100。"); return true; }

            List<double> v = new List<double>();
            for (int i = 0; i < rows; i++) if (Xl.IsNumber(a[i, 0])) v.Add(Xl.D(a[i, 0]));
            if (v.Count == 0) { Xl.Bad("没有数值。"); return true; }
            v.Sort();
            double lo = v[0], hi = v[v.Count - 1];
            double step = (hi - lo) / bins;
            if (step <= 0) step = 1;

            int[] counts = new int[bins];
            foreach (double x in v)
            {
                int idx = (int)Math.Floor((x - lo) / step);
                if (idx >= bins) idx = bins - 1;
                if (idx < 0) idx = 0;
                counts[idx]++;
            }
            object[,] outv = new object[bins + 1, 3];
            outv[0, 0] = "区间"; outv[0, 1] = "上限"; outv[0, 2] = "频数";
            for (int i = 0; i < bins; i++)
            {
                double from = lo + step * i, to = lo + step * (i + 1);
                outv[i + 1, 0] = string.Format(CultureInfo.InvariantCulture, "{0:0.##} ~ {1:0.##}", from, to);
                outv[i + 1, 1] = to;
                outv[i + 1, 2] = counts[i];
            }
            dynamic sh = Xl.AddSheet("频数分布");
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("已生成 " + bins + " 组频数分布 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        // ---------------- 分组汇总 ----------------
        private static bool Group()
        {
            dynamic r = Xl.SelRange("含标题行的数据区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (rows < 2) { Xl.Bad("至少需要标题行 + 一行数据。"); return true; }
            string[] p = Ask.Show("分组汇总",
                new string[] { "分组列（第几列，可多列用逗号）", "汇总列（第几列）", "汇总方式：1=求和 2=计数 3=去重计数 4=平均" },
                new string[] { "1", "2", "1" });
            if (p == null) return true;
            int valCol, mode;
            if (!int.TryParse(p[1], out valCol) || !int.TryParse(p[2], out mode)) { Xl.Bad("请输入整数。"); return true; }
            List<int> keyCols = new List<int>();
            foreach (string s in p[0].Split(new char[] { ',', '，' }))
            {
                int k;
                if (int.TryParse(s.Trim(), out k) && k >= 1 && k <= cols) keyCols.Add(k);
            }
            if (keyCols.Count == 0) { Xl.Bad("分组列无效。"); return true; }
            if (valCol < 1 || valCol > cols) { Xl.Bad("汇总列超出范围。"); return true; }

            List<string> keys = new List<string>();
            Dictionary<string, double> sums = new Dictionary<string, double>(StringComparer.Ordinal);
            Dictionary<string, int> counts = new Dictionary<string, int>(StringComparer.Ordinal);
            Dictionary<string, HashSet<string>> uniq = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
            for (int i = 1; i < rows; i++)
            {
                System.Text.StringBuilder kb = new System.Text.StringBuilder();
                foreach (int k in keyCols) { kb.Append(Xl.S(a[i, k - 1])); kb.Append('\u0001'); }
                string key = kb.ToString();
                if (!sums.ContainsKey(key))
                {
                    sums[key] = 0; counts[key] = 0;
                    uniq[key] = new HashSet<string>(StringComparer.Ordinal);
                    keys.Add(key);
                }
                counts[key]++;
                if (Xl.IsNumber(a[i, valCol - 1])) sums[key] += Xl.D(a[i, valCol - 1]);
                uniq[key].Add(Xl.S(a[i, valCol - 1]));
            }

            int w = keyCols.Count + 2;
            object[,] outv = new object[keys.Count + 1, w];
            for (int j = 0; j < keyCols.Count; j++) outv[0, j] = Xl.S(a[0, keyCols[j] - 1]);
            outv[0, keyCols.Count] = "汇总";
            outv[0, keyCols.Count + 1] = "条数";
            for (int i = 0; i < keys.Count; i++)
            {
                string[] parts = keys[i].Split('\u0001');
                for (int j = 0; j < keyCols.Count; j++) outv[i + 1, j] = parts[j];
                string key = keys[i];
                outv[i + 1, keyCols.Count] =
                    mode == 2 ? (object)counts[key] :
                    mode == 3 ? (object)uniq[key].Count :
                    mode == 4 ? (counts[key] == 0 ? 0d : sums[key] / counts[key]) :
                                (object)sums[key];
                outv[i + 1, keyCols.Count + 1] = counts[key];
            }
            dynamic sh = Xl.AddSheet("分组汇总");
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("已按 " + keyCols.Count + " 列分组汇总 " + keys.Count + " 组 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        // ---------------- 相关系数 ----------------
        private static bool Corr()
        {
            dynamic r = Xl.SelRange("两列数值");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 2) { Xl.Bad("请选择两列数值。"); return true; }

            List<double> x = new List<double>(), y = new List<double>();
            for (int i = 0; i < rows; i++)
            {
                if (!Xl.IsNumber(a[i, 0]) || !Xl.IsNumber(a[i, 1])) continue;
                x.Add(Xl.D(a[i, 0]));
                y.Add(Xl.D(a[i, 1]));
            }
            if (x.Count < 3) { Xl.Bad("有效样本不足（至少 3 组）。"); return true; }
            double mx = 0, my = 0;
            for (int i = 0; i < x.Count; i++) { mx += x[i]; my += y[i]; }
            mx /= x.Count; my /= x.Count;
            double sxy = 0, sxx = 0, syy = 0;
            for (int i = 0; i < x.Count; i++)
            {
                double dx = x[i] - mx, dy = y[i] - my;
                sxy += dx * dy; sxx += dx * dx; syy += dy * dy;
            }
            double corr = (sxx <= 0 || syy <= 0) ? 0 : sxy / Math.Sqrt(sxx * syy);
            dynamic sh = Xl.AddSheet("相关系数");
            object[,] outv = new object[3, 2];
            outv[0, 0] = "样本数"; outv[0, 1] = x.Count;
            outv[1, 0] = "相关系数 r"; outv[1, 1] = Math.Round(corr, 4);
            outv[2, 0] = "强度"; outv[2, 1] = Strength(corr);
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("皮尔逊相关系数 r = " + Math.Round(corr, 4) + "（" + Strength(corr) + "）→ 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        private static string Strength(double r)
        {
            double a = Math.Abs(r);
            if (a >= 0.8) return "强相关";
            if (a >= 0.5) return "中等相关";
            if (a >= 0.3) return "弱相关";
            return "几乎无关";
        }
    }
}

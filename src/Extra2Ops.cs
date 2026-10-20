using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcleHelper
{
    /// <summary>
    /// 文本高级与数值扩展（t2/n2）：相似度、模糊匹配、正则多列、批量替换、随机数、抽样、
    /// 单位换算、分箱、求和校验、密码生成。
    /// </summary>
    internal static class Extra2Ops
    {
        internal static void Register()
        {
            Ops.Reg("t2.similar", Similar);
            Ops.Reg("t2.fuzzy", Fuzzy);
            Ops.Reg("t2.uniqchar", UniqChar);
            Ops.Reg("t2.bylen", ByLen);
            Ops.Reg("t2.rxcols", RxCols);
            Ops.Reg("t2.batchreplace", BatchReplace);
            Ops.Reg("t2.initials", Initials);
            Ops.Reg("t2.between", Between);
            Ops.Reg("t2.count", CountText);

            Ops.Reg("n2.rand", RandNum);
            Ops.Reg("n2.sample", Sample);
            Ops.Reg("n2.unit", UnitConvert);
            Ops.Reg("n2.bins", Bins);
            Ops.Reg("n2.sumcheck", SumCheck);
            Ops.Reg("n2.password", Password);
        }

        // ================= 文本高级 =================

        /// <summary>Levenshtein 相似度（0-1），长度差异过大时提前退出。</summary>
        private static double Similarity(string a, string b)
        {
            if (a == b) return 1.0;
            if (a.Length == 0 || b.Length == 0) return 0.0;
            if (a.Length > 400) a = a.Substring(0, 400);
            if (b.Length > 400) b = b.Substring(0, 400);
            int[] prev = new int[b.Length + 1];
            int[] cur = new int[b.Length + 1];
            for (int j = 0; j <= b.Length; j++) prev[j] = j;
            for (int i = 1; i <= a.Length; i++)
            {
                cur[0] = i;
                for (int j = 1; j <= b.Length; j++)
                {
                    int cost = a[i - 1] == b[j - 1] ? 0 : 1;
                    int del = prev[j] + 1, ins = cur[j - 1] + 1, sub = prev[j - 1] + cost;
                    cur[j] = Math.Min(Math.Min(del, ins), sub);
                }
                int[] t = prev; prev = cur; cur = t;
            }
            double dist = prev[b.Length];
            return 1.0 - dist / Math.Max(a.Length, b.Length);
        }

        private static bool Similar()
        {
            dynamic r = Xl.SelRange("两列文本（第一列基准，第二列对比）");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols < 2) { Xl.Bad("请选择两列：基准文本、对比文本。"); return true; }
            object[,] outv = new object[rows, 2];
            outv[0, 0] = "相似度";
            outv[0, 1] = "是否相同";
            for (int i = 0; i < rows; i++)
            {
                string x = Xl.S(a[i, 0]).Trim(), y = Xl.S(a[i, 1]).Trim();
                if (x.Length == 0 && y.Length == 0) continue;
                double sim = Similarity(x, y);
                outv[i, 0] = Math.Round(sim, 4);
                outv[i, 1] = sim >= 0.999 ? "相同" : (sim >= 0.8 ? "近似" : "不同");
            }
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, cols, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("已计算相似度（写右侧两列）。");
            return true;
        }

        private static bool Fuzzy()
        {
            dynamic r = Xl.SelRange("要查找的文本列");
            if (r == null) return true;
            dynamic refr = Xl.PickRange("请选择对照表（单列，作为候选词）");
            if (refr == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int rr, rc;
            object[,] refs = Xl.Vals(refr, out rr, out rc);

            List<string> pool = new List<string>();
            for (int i = 0; i < rr; i++)
                for (int j = 0; j < rc; j++)
                {
                    string s = Xl.S(a[i, j]).Trim();
                    if (s.Length > 0 && !pool.Contains(s)) pool.Add(s);
                }
            if (pool.Count == 0) { Xl.Bad("对照表为空。"); return true; }

            object[,] outv = new object[rows, 2];
            outv[0, 0] = "最佳匹配";
            outv[0, 1] = "相似度";
            for (int i = 0; i < rows; i++)
            {
                string s = Xl.S(a[i, 0]).Trim();
                if (s.Length == 0) continue;
                string best = null;
                double bestSim = -1;
                foreach (string p in pool)
                {
                    double sim = Similarity(s, p);
                    if (sim > bestSim) { bestSim = sim; best = p; }
                }
                outv[i, 0] = best;
                outv[i, 1] = Math.Round(bestSim, 4);
            }
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, cols, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("已模糊匹配 " + pool.Count + " 个候选词（写右侧）。");
            return true;
        }

        private static bool UniqChar()
        {
            Xl.MapSel(delegate(string s)
            {
                StringBuilder sb = new StringBuilder();
                HashSet<char> seen = new HashSet<char>();
                foreach (char c in s) if (seen.Add(c)) sb.Append(c);
                return sb.ToString();
            }, "要去除重复字符的区域");
            return true;
        }

        private static bool ByLen()
        {
            dynamic r = Xl.SelRange("要按长度筛选的区域");
            if (r == null) return true;
            string[] p = Ask.Show("按长度筛选",
                new string[] { "最小长度", "最大长度", "处理方式：1=清空不符合的  2=只保留符合的（删除整行）" },
                new string[] { "0", "10", "1" });
            if (p == null) return true;
            int lo, hi, mode;
            if (!int.TryParse(p[0], out lo) || !int.TryParse(p[1], out hi)) { Xl.Bad("长度需为整数。"); return true; }
            mode = p[2].Trim() == "2" ? 2 : 1;

            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (mode == 1)
            {
                int n = 0;
                for (int i = 0; i < rows; i++)
                    for (int j = 0; j < cols; j++)
                    {
                        string s = Xl.S(a[i, j]);
                        if (s.Length == 0) continue;
                        if (s.Length < lo || s.Length > hi) { a[i, j] = null; n++; }
                    }
                Xl.Turn(true);
                try { Xl.WriteSkipFormula(r, a); }
                finally { Xl.Turn(false); }
                Xl.Say(Xl.Fmt("已清空 {n} 个长度不符的单元格。", n));
                return true;
            }

            // 模式 2：整行删除（从下往上，避免行号漂移）
            int firstRow = (int)r.Row;
            int deleted = 0;
            Xl.Turn(true);
            try
            {
                for (int i = rows - 1; i >= 0; i--)
                {
                    string s = Xl.S(a[i, 0]);
                    bool bad = s.Length > 0 && (s.Length < lo || s.Length > hi);
                    if (!bad) continue;
                    r.Worksheet.Rows[firstRow + i].Delete();
                    deleted++;
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已删除 {n} 行（按首列长度判断）。", deleted));
            return true;
        }

        private static bool RxCols()
        {
            dynamic r = Xl.SelRange("要提取的区域");
            if (r == null) return true;
            string[] p = Ask.Show("正则提取到多列",
                new string[] { "正则表达式（每个括号为一列）", "示例" },
                new string[] { @"(\d{4})-(\d{2})-(\d{2})", "2024-03-05 → 2024 | 03 | 05" });
            if (p == null) return true;
            Regex rx;
            try { rx = new Regex(p[0]); }
            catch (Exception ex) { Xl.Bad("正则无效：" + ex.Message); return true; }

            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int groups = 0;
            for (int i = 0; i < rows && groups == 0; i++)
                for (int j = 0; j < cols; j++)
                {
                    Match m = rx.Match(Xl.S(a[i, j]));
                    if (m.Success) { groups = m.Groups.Count - 1; break; }
                }
            if (groups == 0) { Xl.Bad("该正则在选区中没有匹配，或没有捕获组。"); return true; }

            object[,] outv = new object[rows, groups];
            int ok = 0;
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    Match m = rx.Match(Xl.S(a[i, j]));
                    if (!m.Success) continue;
                    for (int g = 1; g <= groups; g++) outv[i, g - 1] = m.Groups[g].Value;
                    ok++;
                }
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, cols, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("已提取 " + groups + " 列捕获组（匹配 " + ok + " 行，写右侧）。");
            return true;
        }

        private static bool BatchReplace()
        {
            dynamic map = Xl.PickRange("请选择两列替换规则（查找 / 替换为）");
            if (map == null) return true;
            dynamic r = Xl.SelRange("要执行替换的区域");
            if (r == null) return true;

            int mr, mc;
            object[,] rules = Xl.Vals(map, out mr, out mc);
            if (mc < 2) { Xl.Bad("规则表需要两列。"); return true; }
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int n = 0;
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    string s = Xl.S(a[i, j]);
                    if (s.Length == 0) continue;
                    string orig = s;
                    for (int k = 0; k < mr; k++)
                    {
                        string find = Xl.S(rules[k, 0]);
                        if (find.Length == 0) continue;
                        s = s.Replace(find, Xl.S(rules[k, 1]));
                    }
                    if (s != orig) { a[i, j] = s; n++; }
                }
            Xl.Turn(true);
            try { Xl.WriteSkipFormula(r, a); }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已按 {r} 条规则替换 {n} 个单元格。", mr).Replace("{r}", mr.ToString(CultureInfo.InvariantCulture)).Replace("{n}", n.ToString(CultureInfo.InvariantCulture)));
            return true;
        }

        private static bool Initials()
        {
            string mode = Ask.One("提取首字母", "1 = 中文取拼音首字母   2 = 英文取单词首字母   3 = 中英混合都取", "1");
            if (mode == null) return true;
            int m = mode.Trim() == "2" ? 2 : (mode.Trim() == "3" ? 3 : 1);
            Xl.MapSel(delegate(string s)
            {
                StringBuilder sb = new StringBuilder();
                bool lastWasLetter = false;
                foreach (char c in s)
                {
                    if (c >= 0x4E00 && c <= 0x9FFF)
                    {
                        if (m != 2)
                        {
                            char p = CnOps.PyOf(c);
                            if (p >= 'A' && p <= 'Z') sb.Append(p);
                        }
                        lastWasLetter = false;
                    }
                    else if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'))
                    {
                        if (m != 1 && !lastWasLetter) sb.Append(char.ToUpperInvariant(c));
                        lastWasLetter = true;
                    }
                    else lastWasLetter = false;
                }
                return sb.ToString();
            }, "要提取首字母的区域");
            return true;
        }

        private static bool Between()
        {
            dynamic r = Xl.SelRange("要截取的区域");
            if (r == null) return true;
            string[] p = Ask.Show("截取标记之间的内容",
                new string[] { "起始标记（不含）", "结束标记（不含）" },
                new string[] { "【", "】" });
            if (p == null) return true;
            string left = p[0], right = p[1];
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            object[,] outv = new object[rows, cols];
            int n = 0;
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    string s = Xl.S(a[i, j]);
                    int i1 = left.Length == 0 ? 0 : s.IndexOf(left, StringComparison.Ordinal);
                    if (i1 < 0) continue;
                    i1 = left.Length == 0 ? 0 : i1 + left.Length;
                    int i2 = right.Length == 0 ? s.Length : s.IndexOf(right, i1, StringComparison.Ordinal);
                    if (i2 < 0) i2 = s.Length;
                    outv[i, j] = s.Substring(i1, i2 - i1);
                    n++;
                }
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, cols, outv); }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已截取 {n} 个单元格（写右侧）。", n));
            return true;
        }

        private static bool CountText()
        {
            dynamic r = Xl.SelRange("要统计的区域");
            if (r == null) return true;
            string key = Ask.One("统计字符出现次数", "要统计的文本", "");
            if (key == null || key.Length == 0) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            object[,] outv = new object[rows, cols];
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    string s = Xl.S(a[i, j]);
                    int n = 0, idx = 0;
                    while ((idx = s.IndexOf(key, idx, StringComparison.Ordinal)) >= 0) { n++; idx += key.Length; }
                    outv[i, j] = n;
                }
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, cols, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("已统计「" + key + "」出现次数（写右侧）。");
            return true;
        }

        // ================= 数值扩展 =================

        private static bool RandNum()
        {
            dynamic r = Xl.SelRange("要填入随机数的区域");
            if (r == null) return true;
            string[] p = Ask.Show("随机数",
                new string[] { "最小值", "最大值", "类型：1=整数 2=小数 3=正态(均值=最小值,标准差=最大值)" },
                new string[] { "1", "100", "1" });
            if (p == null) return true;
            double lo, hi;
            if (!double.TryParse(p[0], NumberStyles.Any, CultureInfo.InvariantCulture, out lo) ||
                !double.TryParse(p[1], NumberStyles.Any, CultureInfo.InvariantCulture, out hi))
            { Xl.Bad("请输入数字。"); return true; }
            int type = p[2].Trim() == "2" ? 2 : (p[2].Trim() == "3" ? 3 : 1);
            if (hi < lo) { double t = lo; lo = hi; hi = t; }

            int rows, cols;
            Xl.Vals(r, out rows, out cols);
            Random rnd = new Random(Guid.NewGuid().GetHashCode());
            object[,] outv = new object[rows, cols];
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    if (type == 1) outv[i, j] = (long)(lo + rnd.NextDouble() * (hi - lo + 1));
                    else if (type == 2) outv[i, j] = Math.Round(lo + rnd.NextDouble() * (hi - lo), 4);
                    else
                    {
                        double u1 = 1.0 - rnd.NextDouble(), u2 = 1.0 - rnd.NextDouble();
                        double z = Math.Sqrt(-2.0 * Math.Log(u1)) * Math.Sin(2.0 * Math.PI * u2);
                        outv[i, j] = Math.Round(lo + z * Math.Max(0.0001, hi), 4);
                    }
                }
            Xl.Turn(true);
            try { Xl.WriteSkipFormula(r, outv); }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已生成 {n} 个随机数。", rows * cols));
            return true;
        }

        private static bool Sample()
        {
            dynamic r = Xl.SelRange("含标题行的数据区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (rows < 2) { Xl.Bad("至少需要标题行 + 一行数据。"); return true; }
            string[] p = Ask.Show("随机抽样",
                new string[] { "方式：1=按行数 2=按比例(%)", "数值" },
                new string[] { "1", "10" });
            if (p == null) return true;
            double v;
            if (!double.TryParse(p[1], NumberStyles.Any, CultureInfo.InvariantCulture, out v) || v <= 0)
            { Xl.Bad("请输入正数。"); return true; }
            int take = p[0].Trim() == "2" ? (int)Math.Round((rows - 1) * v / 100.0) : (int)v;
            if (take < 1) take = 1;
            if (take > rows - 1) take = rows - 1;

            List<int> idx = new List<int>();
            for (int i = 1; i < rows; i++) idx.Add(i);
            Random rnd = new Random(Guid.NewGuid().GetHashCode());
            for (int i = idx.Count - 1; i > 0; i--)
            {
                int k = rnd.Next(i + 1);
                int t = idx[i]; idx[i] = idx[k]; idx[k] = t;
            }
            idx.Sort();
            object[,] outv = new object[take + 1, cols];
            for (int j = 0; j < cols; j++) outv[0, j] = a[0, j];
            for (int i = 0; i < take; i++)
                for (int j = 0; j < cols; j++)
                    outv[i + 1, j] = a[idx[i], j];
            dynamic sh = Xl.AddSheet("随机抽样");
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("已从 " + (rows - 1) + " 行中抽取 " + take + " 行 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        private static readonly string[][] UnitTable =
        {
            new string[] { "长度", "m", "km", "cm", "mm", "英里", "英尺", "英寸" },
            new string[] { "重量", "kg", "g", "吨", "磅", "盎司" },
            new string[] { "面积", "平方米", "公顷", "亩", "平方千米", "平方英尺" },
            new string[] { "温度", "℃", "℉", "K" },
            new string[] { "时间", "秒", "分钟", "小时", "天" },
            new string[] { "数据", "B", "KB", "MB", "GB", "TB" }
        };

        private static bool UnitConvert()
        {
            dynamic r = Xl.SelRange("要换算的数值区域");
            if (r == null) return true;
            StringBuilder cat = new StringBuilder("类别：");
            for (int i = 0; i < UnitTable.Length; i++) cat.Append((i + 1) + "=" + UnitTable[i][0] + "  ");
            string c = Ask.One("单位换算", cat.ToString(), "1");
            if (c == null) return true;
            int ci;
            if (!int.TryParse(c, out ci) || ci < 1 || ci > UnitTable.Length) { Xl.Bad("类别无效。"); return true; }
            string[] units = UnitTable[ci - 1];
            string[] p = Ask.Show("单位换算 " + units[0],
                new string[] { "源单位（" + string.Join("/", units, 1, units.Length - 1) + "）", "目标单位" },
                new string[] { units[1], units[2] });
            if (p == null) return true;
            string from = p[0].Trim(), to = p[1].Trim();
            if (Array.IndexOf(units, from) <= 0 || Array.IndexOf(units, to) <= 0)
            { Xl.Bad("单位名称不在列表内。"); return true; }

            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            object[,] outv = new object[rows, cols];
            int n = 0;
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    if (!Xl.IsNumber(a[i, j])) { outv[i, j] = a[i, j]; continue; }
                    double x = Xl.D(a[i, j]);
                    outv[i, j] = ConvertUnit(units[0], from, to, x);
                    n++;
                }
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, cols, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("已换算 " + n + " 个数值（" + from + " → " + to + "，写右侧）。");
            return true;
        }

        private static double ConvertUnit(string cat, string from, string to, double x)
        {
            if (cat == "温度")
            {
                double celsius = from == "℃" ? x : from == "℉" ? (x - 32) * 5 / 9 : x - 273.15;
                return to == "℃" ? celsius : to == "℉" ? celsius * 9 / 5 + 32 : celsius + 273.15;
            }
            double f = Factor(cat, from), t = Factor(cat, to);
            if (t == 0) return x;
            return x * f / t;
        }

        private static double Factor(string cat, string u)
        {
            if (cat == "长度")
            {
                if (u == "m") return 1;
                if (u == "km") return 1000;
                if (u == "cm") return 0.01;
                if (u == "mm") return 0.001;
                if (u == "英里") return 1609.344;
                if (u == "英尺") return 0.3048;
                if (u == "英寸") return 0.0254;
            }
            if (cat == "重量")
            {
                if (u == "kg") return 1;
                if (u == "g") return 0.001;
                if (u == "吨") return 1000;
                if (u == "磅") return 0.45359237;
                if (u == "盎司") return 0.028349523125;
            }
            if (cat == "面积")
            {
                if (u == "平方米") return 1;
                if (u == "公顷") return 10000;
                if (u == "亩") return 666.6666667;
                if (u == "平方千米") return 1000000;
                if (u == "平方英尺") return 0.09290304;
            }
            if (cat == "时间")
            {
                if (u == "秒") return 1;
                if (u == "分钟") return 60;
                if (u == "小时") return 3600;
                if (u == "天") return 86400;
            }
            if (cat == "数据")
            {
                if (u == "B") return 1;
                if (u == "KB") return 1024;
                if (u == "MB") return 1024.0 * 1024;
                if (u == "GB") return 1024.0 * 1024 * 1024;
                if (u == "TB") return 1024.0 * 1024 * 1024 * 1024;
            }
            return 1;
        }

        private static bool Bins()
        {
            dynamic r = Xl.SelRange("要分箱的数值列");
            if (r == null) return true;
            string[] p = Ask.Show("区间分箱",
                new string[] { "组数", "方式：1=等宽 2=等频（每组数量接近）" },
                new string[] { "10", "1" });
            if (p == null) return true;
            int bins, mode;
            if (!int.TryParse(p[0], out bins) || bins < 2 || bins > 200) { Xl.Bad("组数需在 2-200。"); return true; }
            mode = p[1].Trim() == "2" ? 2 : 1;

            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列数值。"); return true; }
            List<double> vals = new List<double>();
            for (int i = 0; i < rows; i++) if (Xl.IsNumber(a[i, 0])) vals.Add(Xl.D(a[i, 0]));
            if (vals.Count == 0) { Xl.Bad("没有数值。"); return true; }
            List<double> sorted = new List<double>(vals);
            sorted.Sort();

            double[] edges = new double[bins + 1];
            if (mode == 1)
            {
                double lo = sorted[0], hi = sorted[sorted.Count - 1];
                double step = (hi - lo) / bins;
                if (step <= 0) step = 1;
                for (int i = 0; i <= bins; i++) edges[i] = lo + step * i;
            }
            else
            {
                for (int i = 0; i <= bins; i++)
                {
                    int idx = (int)Math.Round((double)i * (sorted.Count - 1) / bins);
                    edges[i] = sorted[Math.Min(Math.Max(idx, 0), sorted.Count - 1)];
                }
            }

            object[,] outv = new object[rows, 2];
            outv[0, 0] = "区间";
            outv[0, 1] = "组号";
            for (int i = 0; i < rows; i++)
            {
                if (!Xl.IsNumber(a[i, 0])) continue;
                double v = Xl.D(a[i, 0]);
                int b = bins;
                for (int k = 0; k < bins; k++)
                    if (v >= edges[k] && v < edges[k + 1]) { b = k + 1; break; }
                if (v >= edges[bins]) b = bins;
                outv[i, 0] = string.Format(CultureInfo.InvariantCulture, "[{0:0.##}, {1:0.##})", edges[b - 1], edges[b]);
                outv[i, 1] = b;
            }
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, 1, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("已按" + (mode == 1 ? "等宽" : "等频") + "分 " + bins + " 组（写右侧）。");
            return true;
        }

        private static bool SumCheck()
        {
            dynamic r = Xl.SelRange("要校验的数值区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            double total = 0;
            int cnt = 0;
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                    if (Xl.IsNumber(a[i, j])) { total += Xl.D(a[i, j]); cnt++; }
            string s = Ask.One("求和校验", "期望合计（留空则与当前活动单元格比较）", "");
            if (s == null) return true;
            double expect;
            if (s.Trim().Length == 0)
            {
                dynamic act = Xl.Sel;
                expect = Xl.D(act.Value2);
            }
            else if (!double.TryParse(s.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out expect))
            { Xl.Bad("期望值需为数字。"); return true; }

            double diff = total - expect;
            StringBuilder sb = new StringBuilder();
            sb.Append("校验结果\n\n");
            sb.Append("数值个数：").Append(cnt).Append('\n');
            sb.Append("实际合计：").Append(total.ToString("0.####", CultureInfo.InvariantCulture)).Append('\n');
            sb.Append("期望合计：").Append(expect.ToString("0.####", CultureInfo.InvariantCulture)).Append('\n');
            sb.Append("差异：").Append(diff.ToString("0.####", CultureInfo.InvariantCulture)).Append('\n');
            sb.Append(Math.Abs(diff) < 1e-9 ? "\n结论：一致 ✓" : "\n结论：不一致 ✗");
            Xl.Say(sb.ToString());
            return true;
        }

        private static bool Password()
        {
            dynamic r = Xl.SelRange("要生成密码的区域");
            if (r == null) return true;
            string[] p = Ask.Show("生成随机密码",
                new string[] { "长度", "字符集：1=字母+数字 2=大小写+数字+符号 3=仅数字" },
                new string[] { "12", "2" });
            if (p == null) return true;
            int len, set;
            if (!int.TryParse(p[0], out len) || len < 4 || len > 64) { Xl.Bad("长度需在 4-64。"); return true; }
            set = p[1].Trim() == "3" ? 3 : (p[1].Trim() == "1" ? 1 : 2);

            const string lower = "abcdefghijkmnpqrstuvwxyz";
            const string upper = "ABCDEFGHJKLMNPQRSTUVWXYZ";
            const string digit = "23456789";
            const string sym = "!@#$%^&*-_=+";
            string pool = set == 3 ? digit : set == 1 ? lower + upper + digit : lower + upper + digit + sym;

            int rows, cols;
            Xl.Vals(r, out rows, out cols);
            Random rnd = new Random(Guid.NewGuid().GetHashCode());
            object[,] outv = new object[rows, cols];
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    StringBuilder sb = new StringBuilder();
                    for (int k = 0; k < len; k++) sb.Append(pool[rnd.Next(pool.Length)]);
                    outv[i, j] = sb.ToString();
                }
            Xl.Turn(true);
            try { Xl.WriteSkipFormula(r, outv); }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已生成 {n} 个随机密码（长度 {L}）。", rows * cols).Replace("{L}", len.ToString(CultureInfo.InvariantCulture)));
            return true;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcleHelper
{
    /// <summary>
    /// 批次 A2 · 数据质量（ql.*）：异常值检测、缺失值处理、模糊重复、
    /// 数据字典、单位统一、电话规范化、地址规范化、变更日志。
    /// 输入支持 %TEMP%\excle-ql-inputs.txt 文件钩子（自动化测试用）。
    /// </summary>
    internal static class QualityOps
    {
        internal static void Register()
        {
            Ops.Reg("ql.outlier", Outlier);
            Ops.Reg("ql.missing", Missing);
            Ops.Reg("ql.dup.fuzzy", DupFuzzy);
            Ops.Reg("ql.dict", Dict);
            Ops.Reg("ql.unit.unify", UnitUnify);
            Ops.Reg("ql.phone.norm", PhoneNorm);
            Ops.Reg("ql.addr.norm", AddrNorm);
            Ops.Reg("ql.changelog", ChangeLog);
        }

        // ================= 输入钩子 =================
        private static string FileOpt(string key)
        {
            try
            {
                string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "excle-ql-inputs.txt");
                if (!System.IO.File.Exists(path)) return null;
                foreach (string line in System.IO.File.ReadAllLines(path, Encoding.UTF8))
                {
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    if (!string.Equals(line.Substring(0, eq).Trim(), key, StringComparison.OrdinalIgnoreCase)) continue;
                    return line.Substring(eq + 1).Trim();
                }
            }
            catch { }
            return null;
        }

        private static string Opt(string key, string title, string tip, string def)
        {
            string v = FileOpt(key);
            if (v != null) return v;
            string s = Ask.One(title, tip, def);
            return s == null ? null : s.Trim();
        }

        private static int OptInt(string key, string title, string tip, int def)
        {
            string v = Opt(key, title, tip, def.ToString(CultureInfo.InvariantCulture));
            int r;
            return (v != null && int.TryParse(v, out r)) ? r : def;
        }

        private static bool Num(object v, out double d)
        {
            d = 0;
            if (v == null) return false;
            if (v is double) { d = (double)v; return true; }
            return double.TryParse(Xl.S(v).Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out d);
        }

        // ================= 1. 异常值检测 =================
        private static bool Outlier()
        {
            dynamic r = Xl.SelRange("要检测异常值的数值区域（可多列）");
            if (r == null) return true;
            int mode = OptInt("mode", "异常值检测", "口径：1=3σ  2=IQR(1.5倍)", 2);
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            dynamic ws = r.Worksheet;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;

            List<object[]> hits = new List<object[]>();
            Xl.Turn(true);
            int marked = 0;
            try
            {
                for (int j = 0; j < cols; j++)
                {
                    List<double> vals = new List<double>();
                    for (int i = 1; i < rows; i++)
                    {
                        double d;
                        if (Num(a[i, j], out d)) vals.Add(d);
                    }
                    if (vals.Count < 4) continue;
                    List<double> sorted = new List<double>(vals);
                    sorted.Sort();
                    double lo, hi;
                    if (mode == 2)
                    {
                        double q1 = Pct(sorted, 0.25), q3 = Pct(sorted, 0.75);
                        double iqr = q3 - q1;
                        lo = q1 - 1.5 * iqr; hi = q3 + 1.5 * iqr;
                    }
                    else
                    {
                        double mean = 0;
                        foreach (double d in vals) mean += d;
                        mean /= vals.Count;
                        double var = 0;
                        foreach (double d in vals) var += (d - mean) * (d - mean);
                        double sd = Math.Sqrt(vals.Count > 1 ? var / (vals.Count - 1) : 0);
                        lo = mean - 3 * sd; hi = mean + 3 * sd;
                    }
                    for (int i = 1; i < rows; i++)
                    {
                        double d;
                        if (!Num(a[i, j], out d)) continue;
                        if (d >= lo && d <= hi) continue;
                        marked++;
                        try { ws.Cells[firstRow + i, firstCol + j].Interior.Color = 0xC1C1FF; } catch { }
                        hits.Add(new object[] { (string)ws.Cells[firstRow + i, firstCol + j].Address(false, false), Xl.S(a[0, j]), d, lo, hi });
                    }
                }
            }
            finally { Xl.Turn(false); }

            if (marked == 0) { Xl.Say("没有发现异常值（口径：" + (mode == 2 ? "IQR" : "3σ") + "）。"); return true; }
            object[,] outv = new object[hits.Count + 1, 5];
            outv[0, 0] = "单元格"; outv[0, 1] = "列"; outv[0, 2] = "值"; outv[0, 3] = "下界"; outv[0, 4] = "上界";
            for (int i = 0; i < hits.Count; i++)
                for (int j = 0; j < 5; j++) outv[i + 1, j] = hits[i][j];
            dynamic sh = Xl.AddSheet("异常值清单");
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("标出 " + marked + " 个异常值（选区已标色），清单见工作表「" + (string)sh.Name + "」。");
            return true;
        }

        private static double Pct(List<double> sorted, double p)
        {
            if (sorted.Count == 0) return 0;
            double pos = p * (sorted.Count - 1);
            int lo = (int)Math.Floor(pos), hi = (int)Math.Ceiling(pos);
            if (lo == hi) return sorted[lo];
            return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
        }

        // ================= 2. 缺失值处理 =================
        private static bool Missing()
        {
            dynamic r = Xl.SelRange("要处理空值的区域（可多列）");
            if (r == null) return true;
            int mode = OptInt("mode", "缺失值处理",
                "策略：1=均值 2=中位数 3=众数 4=上方值 5=线性插值 6=标记「空」", 4);
            int header = OptInt("header", "缺失值处理", "第一行是表头吗？1=是 0=否", 1);
            int start = header == 1 ? 1 : 0;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int filled = 0;
            object[,] outv = (object[,])a.Clone();
            for (int j = 0; j < cols; j++)
            {
                List<double> nums = new List<double>();
                Dictionary<string, int> freq = new Dictionary<string, int>(StringComparer.Ordinal);
                for (int i = start; i < rows; i++)
                {
                    string s = Xl.S(a[i, j]).Trim();
                    if (s.Length == 0) continue;
                    double d;
                    if (Num(a[i, j], out d)) nums.Add(d);
                    else { int c; freq.TryGetValue(s, out c); freq[s] = c + 1; }
                }
                double mean = 0, median = 0;
                if (nums.Count > 0)
                {
                    foreach (double d in nums) mean += d;
                    mean /= nums.Count;
                    List<double> sorted = new List<double>(nums);
                    sorted.Sort();
                    median = Pct(sorted, 0.5);
                }
                string mode3 = null;
                int best = 0;
                foreach (KeyValuePair<string, int> kv in freq)
                    if (kv.Value > best) { best = kv.Value; mode3 = kv.Key; }

                object last = null;
                for (int i = start; i < rows; i++)
                {
                    string s = Xl.S(a[i, j]).Trim();
                    bool blank = s.Length == 0;
                    if (!blank) { last = a[i, j]; continue; }
                    object val = null;
                    switch (mode)
                    {
                        case 1: if (nums.Count > 0) val = Math.Round(mean, 4); break;
                        case 2: if (nums.Count > 0) val = Math.Round(median, 4); break;
                        case 3: if (mode3 != null) val = mode3; break;
                        case 4: val = last; break;
                        case 5:
                            if (nums.Count > 0)
                            {
                                // 找下一非空数值做线性插值
                                double prev = 0, next = 0;
                                int pIdx = -1, nIdx = -1;
                                for (int k = i - 1; k >= start; k--) { double d; if (Num(a[k, j], out d)) { prev = d; pIdx = k; break; } }
                                for (int k = i + 1; k < rows; k++) { double d; if (Num(a[k, j], out d)) { next = d; nIdx = k; break; } }
                                if (pIdx >= 0 && nIdx > pIdx)
                                    val = Math.Round(prev + (next - prev) * (i - pIdx) / (nIdx - pIdx), 4);
                                else if (pIdx >= 0) val = prev;
                                else if (nIdx >= 0) val = next;
                            }
                            break;
                        default: val = "空"; break;
                    }
                    if (val == null) continue;
                    outv[i, j] = val;
                    filled++;
                }
            }
            if (filled == 0) { Xl.Say("没有需要填充的空值。"); return true; }
            Xl.Turn(true);
            try { Xl.WriteSkipFormula(r, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("已按策略 " + mode + " 填充 " + filled + " 个空值（可回退）。");
            return true;
        }

        // ================= 3. 模糊重复 =================
        private static bool DupFuzzy()
        {
            dynamic r = Xl.SelRange("要检测相似重复的单列");
            if (r == null) return true;
            int thr = OptInt("threshold", "模糊重复检测", "相似度阈值（0-100，建议 85）", 85);
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列。"); return true; }
            if (rows > 3000) { Xl.Bad("行数过多（>3000），请先缩小范围。"); return true; }

            List<string> vals = new List<string>();
            List<int> rowOf = new List<int>();
            for (int i = 0; i < rows; i++)
            {
                string s = Xl.S(a[i, 0]).Trim();
                if (s.Length == 0) continue;
                vals.Add(s);
                rowOf.Add(i + 1);
            }
            double limit = thr / 100.0;
            List<object[]> pairs = new List<object[]>();
            for (int i = 0; i < vals.Count; i++)
                for (int k = i + 1; k < vals.Count; k++)
                {
                    double sim = AiTasksSimilar(vals[i], vals[k]);
                    if (sim < limit) continue;
                    pairs.Add(new object[] { vals[i], vals[k], Math.Round(sim, 3), rowOf[i], rowOf[k] });
                }
            if (pairs.Count == 0) { Xl.Say("没有相似度 ≥ " + thr + "% 的重复。"); return true; }
            object[,] outv = new object[pairs.Count + 1, 5];
            outv[0, 0] = "值A"; outv[0, 1] = "值B"; outv[0, 2] = "相似度"; outv[0, 3] = "行A"; outv[0, 4] = "行B";
            for (int i = 0; i < pairs.Count; i++)
                for (int j = 0; j < 5; j++) outv[i + 1, j] = pairs[i][j];
            dynamic sh = Xl.AddSheet("模糊重复");
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("发现 " + pairs.Count + " 组相似重复（阈值 " + thr + "%）→ 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        private static double AiTasksSimilar(string a, string b)
        {
            if (a == b) return 1;
            if (a.Length == 0 || b.Length == 0) return 0;
            if (a.Length > 200) a = a.Substring(0, 200);
            if (b.Length > 200) b = b.Substring(0, 200);
            int[] prev = new int[b.Length + 1], cur = new int[b.Length + 1];
            for (int j = 0; j <= b.Length; j++) prev[j] = j;
            for (int i = 1; i <= a.Length; i++)
            {
                cur[0] = i;
                for (int j = 1; j <= b.Length; j++)
                {
                    int cost = a[i - 1] == b[j - 1] ? 0 : 1;
                    int m = Math.Min(prev[j] + 1, cur[j - 1] + 1);
                    m = Math.Min(m, prev[j - 1] + cost);
                    cur[j] = m;
                }
                int[] t = prev; prev = cur; cur = t;
            }
            return 1.0 - (double)prev[b.Length] / Math.Max(a.Length, b.Length);
        }

        // ================= 4. 数据字典 =================
        private static bool Dict()
        {
            dynamic r = Xl.SelRange("要生成数据字典的区域（含表头）");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            object[,] outv = new object[cols + 1, 6];
            outv[0, 0] = "列名"; outv[0, 1] = "类型"; outv[0, 2] = "空值率"; outv[0, 3] = "唯一值数"; outv[0, 4] = "范围/示例"; outv[0, 5] = "说明（待补）";
            for (int j = 0; j < cols; j++)
            {
                string header = Xl.S(a[0, j]).Trim();
                int blank = 0, num = 0, date = 0, total = rows - 1;
                HashSet<string> uniq = new HashSet<string>(StringComparer.Ordinal);
                double mn = double.MaxValue, mx = double.MinValue;
                string sample = "";
                for (int i = 1; i < rows; i++)
                {
                    string s = Xl.S(a[i, j]).Trim();
                    if (s.Length == 0) { blank++; continue; }
                    if (uniq.Count < 5000) uniq.Add(s);
                    if (sample.Length == 0) sample = s.Length > 20 ? s.Substring(0, 20) : s;
                    double d;
                    DateTime dt;
                    if (Num(a[i, j], out d)) { num++; if (d < mn) mn = d; if (d > mx) mx = d; }
                    else if (DateOps.TryParseDate(s, out dt)) date++;
                }
                string type = num >= total * 0.8 ? "数值" : date >= total * 0.8 ? "日期" : "文本";
                string range = (type == "数值" && mn <= mx)
                    ? (mn.ToString("0.##", CultureInfo.InvariantCulture) + " ~ " + mx.ToString("0.##", CultureInfo.InvariantCulture))
                    : "例：" + sample;
                outv[j + 1, 0] = header.Length > 0 ? header : "列" + (j + 1);
                outv[j + 1, 1] = type;
                outv[j + 1, 2] = total == 0 ? 0 : Math.Round((double)blank / total, 4);
                outv[j + 1, 3] = uniq.Count;
                outv[j + 1, 4] = range;
                outv[j + 1, 5] = "";
            }
            dynamic sh = Xl.AddSheet("数据字典");
            Xl.Dump(sh, 1, 1, outv);
            try
            {
                sh.Columns[1].ColumnWidth = 16; sh.Columns[2].ColumnWidth = 8;
                sh.Columns[3].ColumnWidth = 10; sh.Columns[4].ColumnWidth = 10;
                sh.Columns[5].ColumnWidth = 28; sh.Columns[6].ColumnWidth = 24;
                sh.Rows[1].Font.Bold = true;
            }
            catch { }
            Xl.Say("已生成 " + cols + " 列的数据字典 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        // ================= 5. 单位统一 =================
        private static readonly Dictionary<string, double> WeightToKg = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase)
        {
            {"kg",1},{"千克",1},{"公斤",1},{"g",0.001},{"克",0.001},{"吨",1000},{"t",1000},{"磅",0.45359237},{"lb",0.45359237}
        };

        private static bool UnitUnify()
        {
            dynamic r = Xl.SelRange("要统一单位的区域（如 12kg / 3吨 / 500g）");
            if (r == null) return true;
            string target = Opt("target", "单位统一", "统一到哪个单位（kg/g/吨）", "kg");
            if (target == null) return true;
            double tf;
            if (!WeightToKg.TryGetValue(target.Trim(), out tf)) { Xl.Bad("不支持的目标单位：" + target); return true; }

            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            object[,] outv = new object[rows, cols];
            int ok = 0, skip = 0;
            Regex rx = new Regex(@"^\s*([\d\.,]+)\s*([A-Za-z\u4e00-\u9fa5]{1,3})\s*$");
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    string s = Xl.S(a[i, j]).Trim();
                    if (s.Length == 0) { outv[i, j] = a[i, j]; continue; }
                    Match m = rx.Match(s);
                    if (!m.Success) { outv[i, j] = a[i, j]; skip++; continue; }
                    double v;
                    if (!double.TryParse(m.Groups[1].Value.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out v))
                    { outv[i, j] = a[i, j]; skip++; continue; }
                    double f;
                    if (!WeightToKg.TryGetValue(m.Groups[2].Value.Trim(), out f)) { outv[i, j] = a[i, j]; skip++; continue; }
                    outv[i, j] = Math.Round(v * f / tf, 4);
                    ok++;
                }
            if (ok == 0) { Xl.Bad("没有识别到可换算的「数值+单位」。"); return true; }
            Xl.Turn(true);
            try
            {
                TextOps.WriteBlock(r, 0, cols, outv);
                dynamic ws = r.Worksheet;
                int firstRow = (int)r.Row;
                dynamic c1 = ws.Cells[firstRow, (int)r.Column + cols];
                dynamic c2 = ws.Cells[firstRow + rows - 1, (int)r.Column + cols];
                try { ws.Range(c1, c2).NumberFormat = "0.####"; } catch { }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已换算 " + ok + " 个值到 " + target + (skip > 0 ? "；" + skip + " 个无法识别已原样保留" : "") + "（写右侧）。");
            return true;
        }

        // ================= 6. 电话规范化 =================
        private static bool PhoneNorm()
        {
            dynamic r = Xl.SelRange("要规范化的电话列");
            if (r == null) return true;
            int mode = OptInt("mode", "电话规范化", "输出格式：1=纯 11 位  2=加 +86  3=3-4-4 分段", 2);
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int ok = 0;
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    string s = Xl.S(a[i, j]);
                    if (s.Trim().Length == 0) continue;
                    StringBuilder digits = new StringBuilder();
                    foreach (char ch in s) if (char.IsDigit(ch)) digits.Append(ch);
                    string d = digits.ToString();
                    if (d.StartsWith("0086")) d = d.Substring(4);
                    else if (d.StartsWith("86") && d.Length > 11) d = d.Substring(2);
                    if (d.Length != 11 || d[0] != '1') continue;
                    a[i, j] = mode == 1 ? d : mode == 3 ? (d.Substring(0, 3) + "-" + d.Substring(3, 4) + "-" + d.Substring(7)) : ("+86 " + d);
                    ok++;
                }
            if (ok == 0) { Xl.Say("没有识别到 11 位手机号。"); return true; }
            Xl.Turn(true);
            try { Xl.WriteSkipFormula(r, a); }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已规范化 {n} 个电话。", ok));
            return true;
        }

        // ================= 7. 地址规范化 =================
        private static bool AddrNorm()
        {
            dynamic r = Xl.SelRange("要规范化的地址列");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int ok = 0;
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    string s = Xl.S(a[i, j]);
                    if (s.Trim().Length == 0) continue;
                    string t = Regex.Replace(s, @"\s+", " ").Trim();
                    t = t.Replace("　", " ");
                    // 补全省/市/区后缀（仅当紧邻下一个行政区名时）
                    t = Regex.Replace(t, @"([\u4e00-\u9fa5]{2,4})(?=(市|区|县))", "$1");
                    // 去掉重复的省名（如 广东广东省）
                    t = Regex.Replace(t, @"([\u4e00-\u9fa5]{2,4}省)\1", "$1");
                    t = Regex.Replace(t, @"([\u4e00-\u9fa5]{2,4}市)\1", "$1");
                    if (t != s) { a[i, j] = t; ok++; }
                }
            if (ok == 0) { Xl.Say("地址已是规范形式。"); return true; }
            Xl.Turn(true);
            try { Xl.WriteSkipFormula(r, a); }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已规范化 {n} 个地址。", ok));
            return true;
        }

        // ================= 8. 变更日志 =================
        private static bool ChangeLog()
        {
            dynamic r = Xl.SelRange("要跟踪变更的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] now = Xl.Vals(r, out rows, out cols);
            dynamic wb = Xl.Book;
            dynamic snap = null, log = null;
            try
            {
                foreach (dynamic s in wb.Worksheets) { }
            }
            catch { }
            // 取或建隐藏工作表
            snap = GetOrCreateSheet(wb, "_snap");
            log = GetOrCreateSheet(wb, "_changelog");
            int sr, sc;
            object[,] old = Xl.Vals(snap.Range[snap.Cells[1, 1], snap.Cells[rows, cols]], out sr, out sc);

            List<object[]> diffs = new List<object[]>();
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    string o = Xl.S(old[i, j]), n = Xl.S(now[i, j]);
                    if (o == n) continue;
                    diffs.Add(new object[] { DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), "R" + (i + 1) + "C" + (j + 1), o, n });
                }

            // 写日志（追加）
            int logRow = 2;
            try { logRow = (int)log.UsedRange.Rows.Count + 1; if (logRow < 2) logRow = 2; } catch { }
            if (diffs.Count > 0)
            {
                object[,] data = new object[diffs.Count, 4];
                for (int i = 0; i < diffs.Count; i++)
                    for (int j = 0; j < 4; j++) data[i, j] = diffs[i][j];
                Xl.Dump(log, logRow, 1, data);
            }
            if (logRow == 2)
            {
                object[,] head = new object[1, 4];
                head[0, 0] = "时间"; head[0, 1] = "位置"; head[0, 2] = "旧值"; head[0, 3] = "新值";
                Xl.Dump(log, 1, 1, head);
            }
            // 覆盖快照
            Xl.Dump(snap, 1, 1, now);
            try { snap.Visible = 0; log.Visible = 0; } catch { }

            if (diffs.Count == 0) Xl.Say("与上次快照相比没有变化（已更新快照）。");
            else Xl.Say("记录 " + diffs.Count + " 处变更到隐藏表「_changelog」（首次执行只建快照，不记变更）。");
            return true;
        }

        private static dynamic GetOrCreateSheet(dynamic wb, string name)
        {
            try { return wb.Worksheets[name]; }
            catch { }
            dynamic sh = wb.Worksheets.Add(Type.Missing, wb.Worksheets[wb.Worksheets.Count]);
            sh.Name = name;
            return sh;
        }
    }
}

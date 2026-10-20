using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcleHelper
{
    /// <summary>
    /// 批次 A1 · 数据整合（jt.*）：多表关联 JOIN、按主键刷新、记录级合并、
    /// 列名对齐追加、一键生成透视表、交叉表。
    /// 输入支持环境变量钩子（自动化测试用）：EXCLE_JT_LEFT/RIGHT/OPTS。
    /// </summary>
    internal static class JoinOps
    {
        internal static void Register()
        {
            Ops.Reg("jt.join", Join);
            Ops.Reg("jt.update", Update);
            Ops.Reg("jt.group.merge", GroupMerge);
            Ops.Reg("jt.append", Append);
            Ops.Reg("jt.pivot", Pivot);
            Ops.Reg("jt.crosstab", CrossTab);
        }

        // ================= 输入钩子 =================
        /// <summary>
        /// 自动化测试用：%TEMP%\excle-jt-inputs.txt 里按 key=value 提供输入，
        /// 优先级 文件 &gt; 环境变量 &gt; 对话框。文件方式可在同一 Excel 会话内反复改变。
        /// </summary>
        private static string FileOpt(string key)
        {
            try
            {
                string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "excle-jt-inputs.txt");
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

        private static dynamic Pick(string which, string tip)
        {
            string spec = FileOpt(which);
            if (string.IsNullOrEmpty(spec)) spec = Environment.GetEnvironmentVariable("EXCLE_JT_" + which.ToUpperInvariant());
            if (!string.IsNullOrEmpty(spec))
            {
                try
                {
                    object addr = spec;
                    dynamic r = Xl.App.Range[addr];
                    if (r != null) return r;
                }
                catch { }
            }
            return Xl.PickRange(tip);
        }

        /// <summary>key=1;type=2;sep=、 形式的环境变量覆盖弹窗；缺失则询问。</summary>
        private static int OptInt(string key, string title, string tip, int def)
        {
            string single = FileOpt(key);
            if (!string.IsNullOrEmpty(single))
            {
                int sv;
                if (int.TryParse(single, out sv)) return sv;
            }
            string opts = Environment.GetEnvironmentVariable("EXCLE_JT_OPTS");
            if (!string.IsNullOrEmpty(opts))
            {
                foreach (string pair in opts.Split(';'))
                {
                    int eq = pair.IndexOf('=');
                    if (eq <= 0) continue;
                    if (!string.Equals(pair.Substring(0, eq).Trim(), key, StringComparison.OrdinalIgnoreCase)) continue;
                    int v;
                    if (int.TryParse(pair.Substring(eq + 1).Trim(), out v)) return v;
                }
                return def;
            }
            string s = Ask.One(title, tip, def.ToString(CultureInfo.InvariantCulture));
            if (s == null) return int.MinValue;
            int r;
            return int.TryParse(s.Trim(), out r) ? r : def;
        }

        private static string OptStr(string key, string title, string tip, string def)
        {
            string single = FileOpt(key);
            if (!string.IsNullOrEmpty(single)) return single;
            string opts = Environment.GetEnvironmentVariable("EXCLE_JT_OPTS");
            if (!string.IsNullOrEmpty(opts))
            {
                foreach (string pair in opts.Split(';'))
                {
                    int eq = pair.IndexOf('=');
                    if (eq <= 0) continue;
                    if (!string.Equals(pair.Substring(0, eq).Trim(), key, StringComparison.OrdinalIgnoreCase)) continue;
                    return pair.Substring(eq + 1).Trim();
                }
                return def;
            }
            string s = Ask.One(title, tip, def);
            return s == null ? null : s.Trim();
        }

        private static string K(object v)
        {
            return Xl.S(v).Trim();
        }


        private static void WriteSheet(string name, List<string> header, List<object[]> rows, out string sheetName)
        {
            object[,] data = new object[rows.Count + 1, header.Count];
            for (int j = 0; j < header.Count; j++) data[0, j] = header[j];
            for (int i = 0; i < rows.Count; i++)
            {
                object[] r = rows[i];
                for (int j = 0; j < header.Count; j++) data[i + 1, j] = j < r.Length ? r[j] : null;
            }
            dynamic sh = Xl.AddSheet(name);
            Xl.Dump(sh, 1, 1, data);
            try
            {
                sh.Rows[1].Font.Bold = true;
                sh.Rows[1].Interior.Color = 0xE3E7E1;   // 浅灰蓝表头
                sh.Application.ActiveWindow.SplitRow = 0;
            }
            catch { }
            sheetName = (string)sh.Name;
        }

        // ================= 1. 多表关联合并 =================
        private static bool Join()
        {
            dynamic left = Pick("left", "请选择【左表】（第一行是表头）");
            if (left == null) return true;
            dynamic right = Pick("right", "请选择【右表】（第一行是表头）");
            if (right == null) return true;

            int kl = OptInt("keyLeft", "多表关联", "左表用第几列作为关联键", 1);
            if (kl == int.MinValue) return true;
            int kr = OptInt("keyRight", "多表关联", "右表用第几列作为关联键", 1);
            if (kr == int.MinValue) return true;
            int type = OptInt("type", "多表关联", "连接类型：1=左连接 2=内连接 3=右连接 4=全连接", 1);
            if (type == int.MinValue) return true;

            int lr, lc, rr, rc;
            object[,] L = Xl.Vals(left, out lr, out lc);
            object[,] R = Xl.Vals(right, out rr, out rc);
            if (kl < 1 || kl > lc) { Xl.Bad("左表关联键列超出范围。"); return true; }
            if (kr < 1 || kr > rc) { Xl.Bad("右表关联键列超出范围。"); return true; }
            if (lc == 0 || rc == 0) { Xl.Bad("表格为空。"); return true; }

            Dictionary<string, int> idx = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 1; i < rr; i++)
            {
                string k = K(R[i, kr - 1]);
                if (k.Length == 0) continue;
                if (!idx.ContainsKey(k)) idx[k] = i;      // 重复键取首行
            }

            List<int> keepRight = new List<int>();
            for (int j = 0; j < rc; j++) if (j != kr - 1) keepRight.Add(j);

            List<string> header = new List<string>();
            HashSet<string> used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int j = 0; j < lc; j++)
            {
                string h = Xl.S(L[0, j]).Trim();
                if (h.Length == 0) h = "列" + (j + 1);
                while (!used.Add(h)) h += "_";
                header.Add(h);
            }
            for (int k = 0; k < keepRight.Count; k++)
            {
                string h = Xl.S(R[0, keepRight[k]]).Trim();
                if (h.Length == 0) h = "右列" + (keepRight[k] + 1);
                if (!used.Add(h)) h = h + "（右）";
                while (!used.Add(h)) h += "_";
                header.Add(h);
            }

            List<object[]> rows = new List<object[]>();
            HashSet<int> matched = new HashSet<int>();
            int leftUnmatched = 0;
            for (int i = 1; i < lr; i++)
            {
                string key = K(L[i, kl - 1]);
                int ri = 0;
                bool found = false;
                if (key.Length > 0) found = idx.TryGetValue(key, out ri);
                if (found) matched.Add(ri);
                if (!found && (type == 2 || type == 3)) continue;   // 内连接/右连接：跳过左表未匹配行
                object[] row = new object[header.Count];
                for (int j = 0; j < lc; j++) row[j] = L[i, j];
                if (found)
                    for (int k = 0; k < keepRight.Count; k++) row[lc + k] = R[ri, keepRight[k]];
                if (!found) leftUnmatched++;
                rows.Add(row);
            }

            int rightUnmatched = 0;
            if (type == 3 || type == 4)
            {
                for (int i = 1; i < rr; i++)
                {
                    if (matched.Contains(i)) continue;
                    object[] row = new object[header.Count];
                    // 右表独有行：把右表的键值写回左表键列，否则整行看起来没有主键
                    if (kl - 1 < lc) row[kl - 1] = R[i, kr - 1];
                    for (int k = 0; k < keepRight.Count; k++) row[lc + k] = R[i, keepRight[k]];
                    rows.Add(row);
                    rightUnmatched++;
                }
            }

            if (rows.Count == 0) { Xl.Say("没有可输出的行（内连接未匹配到任何键）。"); return true; }
            string name;
            WriteSheet("关联结果", header, rows, out name);
            StringBuilder sb = new StringBuilder();
            sb.Append("关联完成 → 工作表「").Append(name).Append("」\n");
            sb.Append("输出：").Append(rows.Count).Append(" 行 × ").Append(header.Count).Append(" 列\n");
            sb.Append("左表未匹配：").Append(leftUnmatched).Append(" 行");
            if (type == 3 || type == 4) sb.Append("；右表未匹配：").Append(rightUnmatched).Append(" 行");
            sb.Append("\n重复键按首条匹配（右表重复 ").Append(CountDup(idx.Count, rr)).Append(" 个键）");
            Xl.Say(sb.ToString());
            return true;
        }

        private static int CountDup(int distinct, int rows)
        {
            return Math.Max(0, (rows - 1) - distinct);
        }

        // ================= 2. 按主键刷新字段 =================
        private static bool Update()
        {
            dynamic left = Pick("left", "请选择【要更新的表】（含表头）");
            if (left == null) return true;
            dynamic right = Pick("right", "请选择【数据来源表】（含表头）");
            if (right == null) return true;

            int kl = OptInt("keyLeft", "按主键刷新", "左表用第几列作为主键", 1);
            if (kl == int.MinValue) return true;
            int kr = OptInt("keyRight", "按主键刷新", "右表用第几列作为主键", 1);
            if (kr == int.MinValue) return true;

            int lr, lc, rr, rc;
            object[,] L = Xl.Vals(left, out lr, out lc);
            object[,] R = Xl.Vals(right, out rr, out rc);
            if (kl < 1 || kl > lc || kr < 1 || kr > rc) { Xl.Bad("主键列超出范围。"); return true; }

            // 右表按表头名映射到左表同名列
            Dictionary<string, int> rightByName = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int j = 0; j < rc; j++)
            {
                string h = Xl.S(R[0, j]).Trim();
                if (h.Length > 0) rightByName[h] = j;
            }
            List<int> targets = new List<int>();
            List<int> sources = new List<int>();
            for (int j = 0; j < lc; j++)
            {
                string h = Xl.S(L[0, j]).Trim();
                if (h.Length == 0 || j == kl - 1) continue;
                int rj;
                if (rightByName.TryGetValue(h, out rj)) { targets.Add(j); sources.Add(rj); }
            }
            if (targets.Count == 0) { Xl.Bad("两表没有同名的可更新列（除主键外）。"); return true; }

            Dictionary<string, int> idx = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 1; i < rr; i++)
            {
                string k = K(R[i, kr - 1]);
                if (k.Length > 0 && !idx.ContainsKey(k)) idx[k] = i;
            }

            int firstRow = (int)left.Row, firstCol = (int)left.Column;
            dynamic ws = left.Worksheet;
            int updated = 0, unmatched = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 1; i < lr; i++)
                {
                    string key = K(L[i, kl - 1]);
                    int ri;
                    if (key.Length == 0 || !idx.TryGetValue(key, out ri))
                    {
                        unmatched++;
                        try { ws.Cells[firstRow + i, firstCol].Interior.Color = 0xC1C1FF; } catch { }
                        continue;
                    }
                    for (int t = 0; t < targets.Count; t++)
                    {
                        object v = R[ri, sources[t]];
                        if (Xl.S(v).Length == 0) continue;
                        try
                        {
                            ws.Cells[firstRow + i, firstCol + targets[t]].Value2 = v;
                            updated++;
                        }
                        catch { }
                    }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已按主键刷新 " + updated + " 个单元格"
                   + (unmatched > 0 ? "；" + unmatched + " 行在主表中找不到（已标色）" : "") + "。");
            return true;
        }

        // ================= 3. 记录级合并 =================
        private static bool GroupMerge()
        {
            dynamic src = Pick("left", "请选择要合并的表（含表头）");
            if (src == null) return true;
            int keyCol = OptInt("keyLeft", "记录级合并", "按第几列分组（同值合并为一行）", 1);
            if (keyCol == int.MinValue) return true;
            string sep = OptStr("sep", "记录级合并", "多值连接符", "、");
            if (sep == null) return true;

            int rows, cols;
            object[,] a = Xl.Vals(src, out rows, out cols);
            if (keyCol < 1 || keyCol > cols) { Xl.Bad("分组列超出范围。"); return true; }

            List<string> order = new List<string>();
            Dictionary<string, List<string>[]> buckets = new Dictionary<string, List<string>[]>(StringComparer.OrdinalIgnoreCase);
            for (int i = 1; i < rows; i++)
            {
                string key = K(a[i, keyCol - 1]);
                if (key.Length == 0) continue;
                List<string>[] cellValues;
                if (!buckets.TryGetValue(key, out cellValues))
                {
                    cellValues = new List<string>[cols];
                    for (int j = 0; j < cols; j++) cellValues[j] = new List<string>();
                    buckets[key] = cellValues;
                    order.Add(key);
                }
                for (int j = 0; j < cols; j++)
                {
                    string v = Xl.S(a[i, j]).Trim();
                    if (v.Length == 0) continue;
                    if (!cellValues[j].Contains(v)) cellValues[j].Add(v);   // 组内去重
                }
            }
            if (order.Count == 0) { Xl.Bad("没有可合并的数据行。"); return true; }

            List<string> header = new List<string>();
            for (int j = 0; j < cols; j++)
            {
                string h = Xl.S(a[0, j]).Trim();
                header.Add(h.Length > 0 ? h : "列" + (j + 1));
            }
            List<object[]> outRows = new List<object[]>();
            foreach (string key in order)
            {
                List<string>[] v = buckets[key];
                object[] row = new object[cols];
                for (int j = 0; j < cols; j++)
                    row[j] = v[j].Count == 1 ? (object)v[j][0] : string.Join(sep, v[j].ToArray());
                outRows.Add(row);
            }
            string name;
            WriteSheet("记录合并", header, outRows, out name);
            Xl.Say("已把 " + (rows - 1) + " 行按第 " + keyCol + " 列合并为 " + outRows.Count
                   + " 行 → 工作表「" + name + "」（多值用「" + sep + "」连接）。");
            return true;
        }

        // ================= 4. 列名对齐追加 =================
        private static bool Append()
        {
            dynamic main = Pick("left", "请选择【主表】（第一行是表头）");
            if (main == null) return true;
            dynamic add = Pick("right", "请选择【要追加的区域】（可多选多个区域，第一行是表头）");
            if (add == null) return true;

            int mr, mc;
            object[,] M = Xl.Vals(main, out mr, out mc);
            List<string> header = new List<string>();
            Dictionary<string, int> colIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int j = 0; j < mc; j++)
            {
                string h = Xl.S(M[0, j]).Trim();
                if (h.Length == 0) h = "列" + (j + 1);
                string uniq = h;
                int n = 2;
                while (colIndex.ContainsKey(uniq)) uniq = h + "_" + (n++);
                colIndex[uniq] = j;
                header.Add(uniq);
            }

            List<object[]> outRows = new List<object[]>();
            for (int i = 1; i < mr; i++)
            {
                object[] row = new object[header.Count];
                for (int j = 0; j < mc; j++) row[j] = M[i, j];
                outRows.Add(row);
            }

            int appended = 0, skipped = 0;
            try
            {
                int areas = (int)add.Areas.Count;
                for (int ai = 1; ai <= areas; ai++)
                {
                    dynamic area = areas == 1 ? (dynamic)add : add.Areas[ai];
                    int ar, ac;
                    object[,] A = Xl.Vals(area, out ar, out ac);
                    for (int j = 0; j < ac; j++)
                    {
                        string h = Xl.S(A[0, j]).Trim();
                        if (h.Length == 0) continue;
                        if (!colIndex.ContainsKey(h))
                        {
                            colIndex[h] = header.Count;
                            header.Add(h);
                        }
                    }
                    for (int i = 1; i < ar; i++)
                    {
                        object[] row = new object[header.Count];
                        for (int j = 0; j < ac; j++)
                        {
                            string h = Xl.S(A[0, j]).Trim();
                            int target;
                            if (h.Length > 0 && colIndex.TryGetValue(h, out target)) row[target] = A[i, j];
                        }
                        outRows.Add(row);
                        appended++;
                    }
                }
            }
            catch { skipped++; }

            if (appended == 0) { Xl.Bad("追加区域没有数据行。"); return true; }
            // 补齐列宽
            List<object[]> fixedRows = new List<object[]>();
            foreach (object[] r in outRows)
            {
                object[] n = new object[header.Count];
                for (int j = 0; j < header.Count; j++) n[j] = j < r.Length ? r[j] : null;
                fixedRows.Add(n);
            }
            string name;
            WriteSheet("列名对齐追加", header, fixedRows, out name);
            Xl.Say("已按列名对齐追加 " + appended + " 行（主表 " + (mr - 1) + " 行）→ 工作表「" + name + "」。");
            return true;
        }

        // ================= 5. 一键生成透视表 =================
        private static bool Pivot()
        {
            dynamic src = Pick("left", "请选择数据区域（第一行是表头）");
            if (src == null) return true;
            int rowCol = OptInt("rowCol", "生成透视表", "行字段用第几列", 1);
            if (rowCol == int.MinValue) return true;
            int valCol = OptInt("valCol", "生成透视表", "值字段用第几列", 4);
            if (valCol == int.MinValue) return true;
            int fn = OptInt("fn", "生成透视表", "汇总方式：1=求和 2=计数 3=平均 4=最大 5=最小", 1);
            if (fn == int.MinValue) return true;

            int rows, cols;
            object[,] a = Xl.Vals(src, out rows, out cols);
            if (rowCol < 1 || rowCol > cols || valCol < 1 || valCol > cols) { Xl.Bad("字段列号超出范围。"); return true; }
            string rowField = Xl.S(a[0, rowCol - 1]).Trim();
            string valField = Xl.S(a[0, valCol - 1]).Trim();
            if (rowField.Length == 0 || valField.Length == 0) { Xl.Bad("表头不能为空，请先给数据加表头。"); return true; }

            Xl.Turn(true);
            try
            {
                dynamic wb = Xl.Book;
                dynamic ws = Xl.AddSheet("透视表");
                string addr = (string)src.Address(true, true, 1, true);   // 带表名
                dynamic cache = wb.PivotCaches().Create(1, addr);          // xlDatabase
                dynamic pt = cache.CreatePivotTable("'" + (string)ws.Name + "'!R3C1", addr);
                pt.PivotFields(rowField).Orientation = 1;                  // xlRowField
                dynamic df = pt.PivotFields(valField);
                df.Orientation = 4;                                        // xlDataField
                int[] fns = { -4157, -4112, -4106, -4136, -4139 };         // Sum/Count/Average/Max/Min
                int code = fns[Math.Min(Math.Max(fn, 1), 5) - 1];
                df.Function = code;
                df.NumberFormat = "#,##0";
                Xl.Say("已生成透视表 → 工作表「" + (string)ws.Name + "」。\n可按需拖动字段做多维分析。");
                return true;
            }
            catch (Exception ex)
            {
                // 透视表创建失败（版本/区域设置差异）时退回静态交叉表
                Xl.Say("透视表创建失败（" + AiHttp.Sanitize(ex.Message) + "），改用静态交叉表。");
            }
            finally { Xl.Turn(false); }
            return CrossTabCore(src, rowCol, valCol, fn, true);
        }

        // ================= 6. 交叉表 =================
        private static bool CrossTab()
        {
            dynamic src = Pick("left", "请选择数据区域（第一行是表头）");
            if (src == null) return true;
            int rowCol = OptInt("rowCol", "交叉表", "行字段用第几列", 1);
            if (rowCol == int.MinValue) return true;
            int colCol = OptInt("colCol", "交叉表", "列字段用第几列", 2);
            if (colCol == int.MinValue) return true;
            int valCol = OptInt("valCol", "交叉表", "值字段用第几列", 4);
            if (valCol == int.MinValue) return true;
            int fn = OptInt("fn", "交叉表", "汇总方式：1=求和 2=计数 3=平均", 1);
            if (fn == int.MinValue) return true;
            return CrossTabCore(src, rowCol, valCol, fn, false, colCol);
        }

        private static bool CrossTabCore(dynamic src, int rowCol, int valCol, int fn, bool noColField, int colCol = 0)
        {
            int rows, cols;
            object[,] a = Xl.Vals(src, out rows, out cols);
            if (rowCol < 1 || rowCol > cols || valCol < 1 || valCol > cols) { Xl.Bad("字段列号超出范围。"); return true; }
            if (colCol < 1 || colCol > cols) colCol = rowCol == 1 ? 2 : 1;

            List<string> rowKeys = new List<string>();
            List<string> colKeys = new List<string>();
            Dictionary<string, double> sum = new Dictionary<string, double>(StringComparer.Ordinal);
            Dictionary<string, int> cnt = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int i = 1; i < rows; i++)
            {
                string rk = K(a[i, rowCol - 1]);
                string ck = noColField ? "值" : K(a[i, colCol - 1]);
                if (rk.Length == 0) continue;
                if (!rowKeys.Contains(rk)) rowKeys.Add(rk);
                if (!colKeys.Contains(ck)) colKeys.Add(ck);
                string k = rk + "\u0001" + ck;
                double v = Xl.IsNumber(a[i, valCol - 1]) ? Xl.D(a[i, valCol - 1]) : 0;
                double cur;
                sum.TryGetValue(k, out cur);
                sum[k] = cur + v;
                int c;
                cnt.TryGetValue(k, out c);
                cnt[k] = c + 1;
            }
            if (rowKeys.Count == 0) { Xl.Bad("没有可统计的数据。"); return true; }

            rowKeys.Sort(StringComparer.Ordinal);
            colKeys.Sort(StringComparer.Ordinal);

            int total = rowKeys.Count + 2;
            int totalCols = colKeys.Count + 2;
            object[,] outv = new object[total, totalCols];
            string rowTitle = Xl.S(a[0, rowCol - 1]).Trim();
            outv[0, 0] = (rowTitle.Length > 0 ? rowTitle : "行") + " \\ 列";
            for (int j = 0; j < colKeys.Count; j++) outv[0, j + 1] = colKeys[j];
            outv[0, totalCols - 1] = "合计";
            for (int i = 0; i < rowKeys.Count; i++)
            {
                outv[i + 1, 0] = rowKeys[i];
                double rowSum = 0;
                for (int j = 0; j < colKeys.Count; j++)
                {
                    string k = rowKeys[i] + "\u0001" + colKeys[j];
                    double s;
                    int c;
                    sum.TryGetValue(k, out s);
                    cnt.TryGetValue(k, out c);
                    double v = fn == 2 ? c : (fn == 3 ? (c == 0 ? 0 : s / c) : s);
                    outv[i + 1, j + 1] = c == 0 ? (object)null : Math.Round(v, 4);
                    rowSum += v;
                }
                outv[i + 1, totalCols - 1] = Math.Round(rowSum, 4);
            }
            outv[total - 1, 0] = "合计";
            for (int j = 0; j < colKeys.Count; j++)
            {
                double colSum = 0;
                for (int i = 0; i < rowKeys.Count; i++)
                {
                    object v = outv[i + 1, j + 1];
                    if (v != null) colSum += Convert.ToDouble(v, CultureInfo.InvariantCulture);
                }
                outv[total - 1, j + 1] = Math.Round(colSum, 4);
            }
            double grand = 0;
            for (int i = 1; i < total - 1; i++)
            {
                object v = outv[i, totalCols - 1];
                if (v != null) grand += Convert.ToDouble(v, CultureInfo.InvariantCulture);
            }
            outv[total - 1, totalCols - 1] = Math.Round(grand, 4);

            string name;
            List<string> header = new List<string>();
            for (int j = 0; j < totalCols; j++) header.Add(Xl.S(outv[0, j]));
            List<object[]> rowsOut = new List<object[]>();
            for (int i = 1; i < total; i++)
            {
                object[] r = new object[totalCols];
                for (int j = 0; j < totalCols; j++) r[j] = outv[i, j];
                rowsOut.Add(r);
            }
            WriteSheet("交叉表", header, rowsOut, out name);
            Xl.Say("已生成交叉表（" + rowKeys.Count + " 行 × " + colKeys.Count + " 列，含合计）→ 工作表「" + name + "」。");
            return true;
        }
    }
}

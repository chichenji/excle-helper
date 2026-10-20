using System;
using System.Collections.Generic;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcleHelper
{
    /// <summary>
    /// 第二批扩展（p2/x2/q2）：隐私与地址、颜色与对比、数据体检与结构。
    /// 覆盖隐私与地址、颜色与对比、数据体检与结构三类能力。
    /// </summary>
    internal static class Extra3Ops
    {
        internal static void Register()
        {
            Ops.Reg("p2.addr.split", AddrSplit);
            Ops.Reg("p2.mask.batch", MaskBatch);
            Ops.Reg("p2.encrypt", EncryptOp);
            Ops.Reg("p2.decrypt", DecryptOp);
            Ops.Reg("p2.dup.cross", DupCross);

            Ops.Reg("x2.color.sum", ColorSum);
            Ops.Reg("x2.color.extract", ColorExtract);
            Ops.Reg("x2.color.replace", ColorReplace);
            Ops.Reg("x2.table.diff", TableDiff);
            Ops.Reg("x2.cell.diff", CellDiff);

            Ops.Reg("q2.audit", Audit);
            Ops.Reg("q2.seq.group", SeqGroup);
            Ops.Reg("q2.subtotal", Subtotal);
            Ops.Reg("q2.colmerge", ColMerge);
            Ops.Reg("q2.table.toggle", TableToggle);
            Ops.Reg("q2.blank.insert", BlankInsert);
        }

        // ================= 地址与隐私 =================

        private static readonly Regex AddrRx = new Regex(
            @"(?<prov>[^省市区县]{2,8}?(?:省|自治区|特别行政区))?(?<city>[^省市区县]{1,10}?(?:市|自治州|地区|盟))?(?<dist>[^省市区县]{1,10}?(?:区|县|旗|市辖区))?(?<rest>.*)",
            RegexOptions.Compiled);

        private static bool AddrSplit()
        {
            dynamic r = Xl.SelRange("要拆分的地址列");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列地址。"); return true; }
            object[,] outv = new object[rows, 4];
            int ok = 0;
            for (int i = 0; i < rows; i++)
            {
                string s = Xl.S(a[i, 0]).Trim();
                if (s.Length == 0) continue;
                Match m = AddrRx.Match(s);
                if (!m.Success) { outv[i, 3] = s; continue; }
                outv[i, 0] = m.Groups["prov"].Value;
                outv[i, 1] = m.Groups["city"].Value;
                outv[i, 2] = m.Groups["dist"].Value;
                outv[i, 3] = m.Groups["rest"].Value.Trim();
                ok++;
            }
            if (ok == 0) { Xl.Bad("没有识别出地址结构。"); return true; }
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, 1, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("已拆分为 省 / 市 / 区县 / 详细地址 四列（写右侧）。");
            return true;
        }

        private static string MaskOne(string s)
        {
            string t = s.Trim();
            if (t.Length == 0) return s;
            // 邮箱
            Match mail = Regex.Match(t, @"^([\w\.\-]+)@([\w\-]+\.[\w\.\-]+)$");
            if (mail.Success)
            {
                string name = mail.Groups[1].Value;
                string head = name.Length <= 1 ? name : name.Substring(0, 1);
                return head + "***@" + mail.Groups[2].Value;
            }
            // 身份证 18 位（末位可为 X）
            if (Regex.IsMatch(t, @"^\d{17}[\dXx]$")) return t.Substring(0, 6) + "********" + t.Substring(14);
            // 手机号 11 位
            if (Regex.IsMatch(t, @"^1[3-9]\d{9}$")) return t.Substring(0, 3) + "****" + t.Substring(7);
            // 银行卡 16-19 位
            if (Regex.IsMatch(t, @"^\d{16,19}$")) return new string('*', t.Length - 4) + t.Substring(t.Length - 4);
            // 中文姓名 2-4 字
            if (Regex.IsMatch(t, @"^[\u4e00-\u9fa5]{2,4}$"))
            {
                if (t.Length == 2) return t.Substring(0, 1) + "*";
                return t.Substring(0, 1) + new string('*', t.Length - 2) + t.Substring(t.Length - 1);
            }
            return s;
        }

        private static bool MaskBatch()
        {
            dynamic r = Xl.SelRange("要脱敏的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int n = 0;
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    string s = Xl.S(a[i, j]);
                    if (s.Trim().Length == 0) continue;
                    string m = MaskOne(s);
                    if (m != s) { a[i, j] = m; n++; }
                }
            if (n == 0) { Xl.Say("没有识别到需要脱敏的内容。"); return true; }
            Xl.Turn(true);
            try { Xl.WriteSkipFormula(r, a); }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已脱敏 {n} 个单元格（手机号/身份证/银行卡/邮箱/姓名自动识别）。", n));
            return true;
        }

        private static string Crypt(string text, string pwd, bool encrypt)
        {
            if (encrypt)
            {
                byte[] salt = new byte[16];
                using (RNGCryptoServiceProvider rng = new RNGCryptoServiceProvider()) rng.GetBytes(salt);
                using (Rfc2898DeriveBytes k = new Rfc2898DeriveBytes(pwd, salt, 1000))
                using (AesManaged aes = new AesManaged())
                {
                    aes.KeySize = 256;
                    aes.Mode = CipherMode.CBC;
                    aes.Padding = PaddingMode.PKCS7;
                    aes.Key = k.GetBytes(32);
                    aes.GenerateIV();
                    byte[] data = Encoding.UTF8.GetBytes(text);
                    using (ICryptoTransform enc = aes.CreateEncryptor())
                    {
                        byte[] cipher = enc.TransformFinalBlock(data, 0, data.Length);
                        byte[] all = new byte[32 + cipher.Length];
                        Buffer.BlockCopy(salt, 0, all, 0, 16);
                        Buffer.BlockCopy(aes.IV, 0, all, 16, 16);
                        Buffer.BlockCopy(cipher, 0, all, 32, cipher.Length);
                        return "ENC:" + Convert.ToBase64String(all);
                    }
                }
            }
            if (!text.StartsWith("ENC:", StringComparison.Ordinal)) return null;
            byte[] raw = Convert.FromBase64String(text.Substring(4));
            if (raw.Length < 33) return null;
            byte[] s2 = new byte[16], iv = new byte[16], cipher2 = new byte[raw.Length - 32];
            Buffer.BlockCopy(raw, 0, s2, 0, 16);
            Buffer.BlockCopy(raw, 16, iv, 0, 16);
            Buffer.BlockCopy(raw, 32, cipher2, 0, cipher2.Length);
            using (Rfc2898DeriveBytes k = new Rfc2898DeriveBytes(pwd, s2, 1000))
            using (AesManaged aes = new AesManaged())
            {
                aes.KeySize = 256;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;
                aes.Key = k.GetBytes(32);
                aes.IV = iv;
                using (ICryptoTransform dec = aes.CreateDecryptor())
                {
                    byte[] plain = dec.TransformFinalBlock(cipher2, 0, cipher2.Length);
                    return Encoding.UTF8.GetString(plain);
                }
            }
        }

        private static bool EncryptOp()
        {
            dynamic r = Xl.SelRange("要加密的区域");
            if (r == null) return true;
            string pwd = Ask.One("文本加密", "口令（请自行保存，丢失无法还原）", "");
            if (pwd == null || pwd.Length == 0) { Xl.Bad("口令不能为空。"); return true; }
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int n = 0;
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    string s = Xl.S(a[i, j]);
                    if (s.Length == 0 || s.StartsWith("ENC:", StringComparison.Ordinal)) continue;
                    a[i, j] = Crypt(s, pwd, true);
                    n++;
                }
            Xl.Turn(true);
            try { Xl.WriteSkipFormula(r, a); }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已加密 {n} 个单元格（AES-256-CBC，密文以 ENC: 开头）。", n));
            return true;
        }

        private static bool DecryptOp()
        {
            dynamic r = Xl.SelRange("要解密的区域");
            if (r == null) return true;
            string pwd = Ask.One("文本解密", "口令", "");
            if (pwd == null || pwd.Length == 0) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int n = 0, bad = 0;
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    string s = Xl.S(a[i, j]);
                    if (!s.StartsWith("ENC:", StringComparison.Ordinal)) continue;
                    try
                    {
                        string plain = Crypt(s, pwd, false);
                        if (plain == null) { bad++; continue; }
                        a[i, j] = plain;
                        n++;
                    }
                    catch { bad++; }
                }
            Xl.Turn(true);
            try { Xl.WriteSkipFormula(r, a); }
            finally { Xl.Turn(false); }
            Xl.Say("已解密 " + n + " 个单元格" + (bad > 0 ? "（" + bad + " 个口令不符或格式损坏）" : "") + "。");
            return true;
        }

        private static bool DupCross()
        {
            dynamic r = Xl.SelRange("主表的键值列");
            if (r == null) return true;
            dynamic other = Xl.PickRange("请选择要对比的另一列（可来自其他工作表）");
            if (other == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int orr, occ;
            object[,] b = Xl.Vals(other, out orr, out occ);

            HashSet<string> pool = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < orr; i++)
                for (int j = 0; j < occ; j++)
                {
                    string s = Xl.S(a[i, j]).Trim();
                    if (s.Length > 0) pool.Add(s);
                }
            object[,] outv = new object[rows, 2];
            outv[0, 0] = "是否重复";
            outv[0, 1] = "说明";
            int hit = 0;
            for (int i = 0; i < rows; i++)
            {
                string s = Xl.S(a[i, 0]).Trim();
                if (s.Length == 0) continue;
                bool dup = pool.Contains(s);
                outv[i, 0] = dup ? "重复" : "";
                outv[i, 1] = dup ? "在对比区域中存在" : "";
                if (dup) hit++;
            }
            Xl.Turn(true);
            try
            {
                TextOps.WriteBlock(r, 0, cols, outv);
                dynamic ws = r.Worksheet;
                int firstRow = (int)r.Row;
                dynamic c1 = ws.Cells[firstRow, (int)r.Column + cols];
                dynamic c2 = ws.Cells[firstRow + rows - 1, (int)r.Column + cols];
                ws.Range(c1, c2).FormatConditions.Delete();
                dynamic fc = ws.Range(c1, c2).FormatConditions.Add(1, 3, "=\"重复\"");
                fc.Interior.Color = 0xC1C1FF;
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("对比完成：{n} 个键值在另一区域中重复（写右侧并标色）。", hit));
            return true;
        }

        // ================= 颜色与对比 =================

        private static int TargetColor()
        {
            try
            {
                dynamic c = Xl.Act;
                object v = c.Interior.Color;
                if (v == null) return -1;
                return Convert.ToInt32(v, CultureInfo.InvariantCulture);
            }
            catch { return -1; }
        }

        private static bool ColorSum()
        {
            int target = TargetColor();
            if (target < 0) { Xl.Bad("无法读取活动单元格的填充色。"); return true; }
            dynamic r = Xl.SelRange("要按颜色统计的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            dynamic ws = r.Worksheet;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            int cnt = 0, colored = 0;
            double sum = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                    for (int j = 0; j < cols; j++)
                    {
                        dynamic cell = ws.Cells[firstRow + i, firstCol + j];
                        int color;
                        try { color = Convert.ToInt32(cell.Interior.Color, CultureInfo.InvariantCulture); }
                        catch { continue; }
                        if (color != target) continue;
                        colored++;
                        if (Xl.IsNumber(a[i, j])) { sum += Xl.D(a[i, j]); cnt++; }
                    }
            }
            finally { Xl.Turn(false); }
            Xl.Say("按颜色统计（以活动单元格颜色为准）\n\n"
                   + "同色单元格：" + colored + " 个\n"
                   + "其中数值：" + cnt + " 个\n"
                   + "数值合计：" + sum.ToString("0.####", CultureInfo.InvariantCulture));
            return true;
        }

        private static bool ColorExtract()
        {
            int target = TargetColor();
            if (target < 0) { Xl.Bad("无法读取活动单元格的填充色。"); return true; }
            dynamic r = Xl.SelRange("要按颜色提取的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            dynamic ws = r.Worksheet;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            List<object> picked = new List<object>();
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                    for (int j = 0; j < cols; j++)
                    {
                        dynamic cell = ws.Cells[firstRow + i, firstCol + j];
                        int color;
                        try { color = Convert.ToInt32(cell.Interior.Color, CultureInfo.InvariantCulture); }
                        catch { continue; }
                        if (color == target) picked.Add(a[i, j]);
                    }
            }
            finally { Xl.Turn(false); }
            if (picked.Count == 0) { Xl.Say("没有同色单元格。"); return true; }
            object[,] outv = new object[picked.Count, 1];
            for (int i = 0; i < picked.Count; i++) outv[i, 0] = picked[i];
            dynamic sh = Xl.AddSheet("同色提取");
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("已提取 " + picked.Count + " 个同色单元格 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        private static bool ColorReplace()
        {
            int target = TargetColor();
            if (target < 0) { Xl.Bad("无法读取活动单元格的填充色。"); return true; }
            string v = Ask.One("按颜色替换内容", "把这些同色单元格替换为（留空=清空内容）", "");
            if (v == null) return true;
            dynamic r = Xl.SelRange("要处理的范围");
            if (r == null) return true;
            dynamic ws = r.Worksheet;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            int rows = (int)r.Rows.Count, cols = (int)r.Columns.Count;
            int n = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                    for (int j = 0; j < cols; j++)
                    {
                        dynamic cell = ws.Cells[firstRow + i, firstCol + j];
                        int color;
                        try { color = Convert.ToInt32(cell.Interior.Color, CultureInfo.InvariantCulture); }
                        catch { continue; }
                        if (color != target) continue;
                        try
                        {
                            if (v.Length == 0) cell.ClearContents();
                            else cell.Value2 = v;
                            n++;
                        }
                        catch { }
                    }
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已替换 {n} 个同色单元格。", n));
            return true;
        }

        private static bool TableDiff()
        {
            dynamic left = Xl.PickRange("请选择【左表】区域（第一行是表头）");
            if (left == null) return true;
            dynamic right = Xl.PickRange("请选择【右表】区域（第一行是表头）");
            if (right == null) return true;
            string k = Ask.One("两表差异对比", "用第几列作为主键（列号）", "1");
            if (k == null) return true;
            int keyIdx;
            if (!int.TryParse(k, out keyIdx) || keyIdx < 1) { Xl.Bad("主键列号无效。"); return true; }

            int lr, lc, rr, rc;
            object[,] A = Xl.Vals(left, out lr, out lc);
            object[,] B = Xl.Vals(right, out rr, out rc);
            if (keyIdx > lc || keyIdx > rc) { Xl.Bad("主键列超出范围。"); return true; }

            Dictionary<string, int> bIdx = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 1; i < rr; i++)
            {
                string key = Xl.S(B[i, keyIdx - 1]).Trim();
                if (key.Length > 0 && !bIdx.ContainsKey(key)) bIdx[key] = i;
            }
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            List<object[]> rowsOut = new List<object[]>();
            int maxCols = Math.Max(lc, rc);
            for (int i = 1; i < lr; i++)
            {
                string key = Xl.S(A[i, keyIdx - 1]).Trim();
                if (key.Length == 0) continue;
                seen.Add(key);
                int bi;
                if (!bIdx.TryGetValue(key, out bi))
                {
                    rowsOut.Add(new object[] { key, "仅左表有", JoinRow(A, i, lc), "" });
                    continue;
                }
                string lv = JoinRow(A, i, lc), rv = JoinRow(B, bi, rc);
                if (!string.Equals(lv, rv, StringComparison.Ordinal))
                    rowsOut.Add(new object[] { key, "值不同", lv, rv });
            }
            for (int i = 1; i < rr; i++)
            {
                string key = Xl.S(B[i, keyIdx - 1]).Trim();
                if (key.Length == 0 || seen.Contains(key)) continue;
                rowsOut.Add(new object[] { key, "仅右表有", "", JoinRow(B, i, rc) });
            }
            if (rowsOut.Count == 0) { Xl.Say("两表完全一致（按主键逐行比对）。"); return true; }
            object[,] outv = new object[rowsOut.Count + 1, 4];
            outv[0, 0] = "主键"; outv[0, 1] = "差异类型"; outv[0, 2] = "左表"; outv[0, 3] = "右表";
            for (int i = 0; i < rowsOut.Count; i++)
                for (int j = 0; j < 4; j++) outv[i + 1, j] = rowsOut[i][j];
            dynamic sh = Xl.AddSheet("两表差异");
            Xl.Dump(sh, 1, 1, outv);
            try { sh.Columns[1].ColumnWidth = 14; sh.Columns[2].ColumnWidth = 12; sh.Columns[3].ColumnWidth = 40; sh.Columns[4].ColumnWidth = 40; } catch { }
            Xl.Say("发现 " + rowsOut.Count + " 处差异 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        private static string JoinRow(object[,] a, int row, int cols)
        {
            StringBuilder sb = new StringBuilder();
            for (int j = 0; j < cols; j++)
            {
                if (j > 0) sb.Append(" | ");
                sb.Append(Xl.S(a[row, j]));
            }
            return sb.ToString();
        }

        private static bool CellDiff()
        {
            dynamic c1 = Xl.PickRange("请选择第一个单元格");
            if (c1 == null) return true;
            dynamic c2 = Xl.PickRange("请选择第二个单元格");
            if (c2 == null) return true;
            string a = Xl.S(c1.Value2), b = Xl.S(c2.Value2);
            StringBuilder sb = new StringBuilder();
            sb.Append("单元格差异\n\n");
            sb.Append("A：").Append(a.Length > 60 ? a.Substring(0, 60) + "…" : a).Append('\n');
            sb.Append("B：").Append(b.Length > 60 ? b.Substring(0, 60) + "…" : b).Append('\n');
            if (a == b) { sb.Append("\n内容完全相同 ✓"); Xl.Say(sb.ToString()); return true; }
            int n = Math.Min(a.Length, b.Length), first = -1, last = -1;
            for (int i = 0; i < n; i++) if (a[i] != b[i]) { if (first < 0) first = i; last = i; }
            if (first < 0) { first = n; last = n; }
            sb.Append("\n长度：").Append(a.Length).Append(" vs ").Append(b.Length);
            sb.Append("\n首个不同位置：第 ").Append(first + 1).Append(" 个字符");
            if (last >= first)
                sb.Append("\nA 片段：").Append(Slice(a, first)).Append("\nB 片段：").Append(Slice(b, first));
            Xl.Say(sb.ToString());
            return true;
        }

        private static string Slice(string s, int from)
        {
            if (from >= s.Length) return "（无）";
            int len = Math.Min(24, s.Length - from);
            return s.Substring(from, len) + (s.Length - from > len ? "…" : "");
        }

        // ================= 数据体检与结构 =================

        private static bool Audit()
        {
            dynamic r = Xl.SelRange("要体检的数据区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (rows < 2) { Xl.Bad("至少需要一行数据。"); return true; }

            int blanks = 0, spaces = 0, errs = 0, dupCells = 0, longText = 0, numAsText = 0;
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    object v = a[i, j];
                    string s = Xl.S(v);
                    if (v == null || s.Length == 0) { blanks++; continue; }
                    if (s.Trim().Length == 0) { blanks++; continue; }
                    if (s != s.Trim()) spaces++;
                    if (s.IndexOf("#", StringComparison.Ordinal) == 0 && s.Length <= 8) errs++;
                    if (s.Length > 100) longText++;
                    if (!Xl.IsNumber(v) && Regex.IsMatch(s, @"^-?\d+(\.\d+)?$")) numAsText++;
                }
            Dictionary<string, int> freq = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int i = 0; i < rows; i++)
            {
                StringBuilder key = new StringBuilder();
                for (int j = 0; j < cols; j++) { key.Append(Xl.S(a[i, j])); key.Append('\u0001'); }
                string kk = key.ToString();
                int c;
                freq.TryGetValue(kk, out c);
                freq[kk] = c + 1;
            }
            foreach (KeyValuePair<string, int> kv in freq) if (kv.Value > 1) dupCells += kv.Value;

            object[,] outv = new object[9, 3];
            string[] names = { "总行数", "总列数", "空单元格", "首尾空格", "疑似错误值", "重复行涉及行数", "超长文本(>100)", "数字存成文本", "结论" };
            object[] vals =
            {
                rows, cols, blanks, spaces, errs, dupCells, longText, numAsText,
                (blanks + spaces + errs + dupCells + longText + numAsText) == 0 ? "未发现明显问题 ✓" : "存在需要处理的项目"
            };
            for (int i = 0; i < names.Length; i++) { outv[i, 0] = names[i]; outv[i, 1] = vals[i]; }
            dynamic sh = Xl.AddSheet("数据体检");
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("体检完成 → 工作表「" + (string)sh.Name + "」（" + rows + " 行 × " + cols + " 列）。");
            return true;
        }

        private static bool SeqGroup()
        {
            dynamic r = Xl.SelRange("按第几列分组？请选择分组列（单列）");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列作为分组依据。"); return true; }
            object[,] outv = new object[rows, 2];
            outv[0, 0] = "组内序号";
            outv[0, 1] = "组号";
            int group = 0;
            string prev = null;
            int seq = 0;
            for (int i = 0; i < rows; i++)
            {
                string key = Xl.S(a[i, 0]).Trim();
                if (i == 0 || key != prev) { group++; seq = 1; prev = key; }
                else seq++;
                outv[i, 0] = seq;
                outv[i, 1] = group;
            }
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, 1, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("已生成组内序号与组号（" + group + " 组，写右侧）。");
            return true;
        }

        private static bool Subtotal()
        {
            dynamic r = Xl.SelRange("含标题行的数据区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (rows < 2) { Xl.Bad("至少需要标题行 + 一行数据。"); return true; }
            string k = Ask.One("插入小计行", "按第几列分组（列号）", "1");
            if (k == null) return true;
            int keyIdx;
            if (!int.TryParse(k, out keyIdx) || keyIdx < 1 || keyIdx > cols) { Xl.Bad("列号无效。"); return true; }

            List<object[]> outRows = new List<object[]>();
            outRows.Add(RowOf(a, 0, cols));
            double[] sums = new double[cols];
            string prev = null;
            int n = 0;
            for (int i = 1; i < rows; i++)
            {
                string key = Xl.S(a[i, keyIdx - 1]).Trim();
                if (prev != null && key != prev)
                {
                    outRows.Add(SumRow(a, 0, cols, sums, prev, keyIdx));
                    sums = new double[cols];
                    n++;
                }
                outRows.Add(RowOf(a, i, cols));
                for (int j = 0; j < cols; j++) if (Xl.IsNumber(a[i, j])) sums[j] += Xl.D(a[i, j]);
                prev = key;
            }
            outRows.Add(SumRow(a, 0, cols, sums, prev, keyIdx));
            n++;

            object[,] outv = new object[outRows.Count, cols];
            for (int i = 0; i < outRows.Count; i++)
                for (int j = 0; j < cols; j++) outv[i, j] = outRows[i][j];
            dynamic sh = Xl.AddSheet("小计");
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("已插入 " + n + " 个小计行 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        private static object[] RowOf(object[,] a, int row, int cols)
        {
            object[] o = new object[cols];
            for (int j = 0; j < cols; j++) o[j] = a[row, j];
            return o;
        }

        private static object[] SumRow(object[,] a, int headerRow, int cols, double[] sums, string key, int keyIdx)
        {
            object[] o = new object[cols];
            o[keyIdx - 1] = key + " 小计";
            for (int j = 0; j < cols; j++) if (j != keyIdx - 1 && Math.Abs(sums[j]) > 1e-12) o[j] = sums[j];
            return o;
        }

        private static bool ColMerge()
        {
            dynamic r = Xl.SelRange("含表头的数据区域（多个同名列会合并为一列）");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (rows < 2) { Xl.Bad("至少需要标题行 + 一行数据。"); return true; }

            List<string> order = new List<string>();
            Dictionary<string, List<object>> buckets = new Dictionary<string, List<object>>(StringComparer.Ordinal);
            for (int j = 0; j < cols; j++)
            {
                string name = Xl.S(a[0, j]).Trim();
                if (name.Length == 0) name = "列" + (j + 1);
                List<object> list;
                if (!buckets.TryGetValue(name, out list)) { list = new List<object>(); buckets[name] = list; order.Add(name); }
                for (int i = 1; i < rows; i++)
                {
                    object v = a[i, j];
                    if (v != null && Xl.S(v).Length > 0) list.Add(v);
                }
            }
            int maxLen = 0;
            foreach (string name in order) maxLen = Math.Max(maxLen, buckets[name].Count);
            object[,] outv = new object[maxLen + 1, order.Count];
            for (int j = 0; j < order.Count; j++)
            {
                outv[0, j] = order[j];
                List<object> list = buckets[order[j]];
                for (int i = 0; i < list.Count; i++) outv[i + 1, j] = list[i];
            }
            dynamic sh = Xl.AddSheet("同名列合并");
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("已把 " + cols + " 列按表头合并为 " + order.Count + " 列 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        private static bool TableToggle()
        {
            dynamic r = Xl.SelRange("要转换的区域（表格 ↔ 普通区域）");
            if (r == null) return true;
            dynamic ws = r.Worksheet;
            dynamic lo = null;
            try
            {
                int n = (int)ws.ListObjects.Count;
                for (int i = 1; i <= n; i++)
                {
                    dynamic t = ws.ListObjects.Item(i);
                    dynamic tr = t.Range;
                    if ((int)tr.Row == (int)r.Row && (int)tr.Column == (int)r.Column) { lo = t; break; }
                }
            }
            catch { }
            Xl.Turn(true);
            try
            {
                if (lo != null)
                {
                    lo.Unlist();
                    Xl.Say("已把表格转为普通区域（格式保留）。");
                }
                else
                {
                    ws.ListObjects.Add(1, r, Type.Missing, 1);   // xlSrcRange, HasHeaders
                    Xl.Say("已把区域转为表格（自动筛选 + 条纹样式）。");
                }
            }
            catch (Exception ex) { Xl.Bad("转换失败：" + ex.Message); }
            finally { Xl.Turn(false); }
            return true;
        }

        private static bool BlankInsert()
        {
            dynamic r = Xl.SelRange("要插入空白行的区域");
            if (r == null) return true;
            string[] p = Ask.Show("隔行插入空白行",
                new string[] { "每隔几行插入", "每次插入几行" },
                new string[] { "1", "1" });
            if (p == null) return true;
            int every, count;
            if (!int.TryParse(p[0], out every) || every < 1) { Xl.Bad("间隔需为正整数。"); return true; }
            if (!int.TryParse(p[1], out count) || count < 1) { Xl.Bad("插入行数需为正整数。"); return true; }
            int rows = (int)r.Rows.Count;
            int firstRow = (int)r.Row;
            dynamic ws = r.Worksheet;
            int n = 0;
            Xl.Turn(true);
            try
            {
                for (int i = rows; i > 0; i -= every)
                {
                    if (i == rows) continue;    // 末尾不加
                    ws.Rows[firstRow + i].Resize[count].Insert();
                    n++;
                }
            }
            catch (Exception ex) { Xl.Bad("插入失败：" + ex.Message); }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已在 {n} 处插入空白行。", n));
            return true;
        }
    }
}

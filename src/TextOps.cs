using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcleHelper
{
    /// <summary>Text domain: character filtering, insertion, cutting, case, regex, splitting.</summary>
    internal static class TextOps
    {
        // ---------------- character classes ----------------
        internal static bool IsCn(char c)
        {
            return (c >= 0x4E00 && c <= 0x9FFF) || (c >= 0x3400 && c <= 0x4DBF) || (c >= 0xF900 && c <= 0xFAFF);
        }

        internal static bool IsEn(char c)
        {
            return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
        }

        internal static bool IsDig(char c)
        {
            return c >= '0' && c <= '9';
        }

        internal static bool IsSym(char c)
        {
            return !char.IsLetterOrDigit(c) && !IsBlank(c);
        }

        internal static bool IsBlank(char c)
        {
            return c == ' ' || c == '\t' || c == '\r' || c == '\n' || c == '\u3000' || char.IsWhiteSpace(c);
        }

        internal static bool IsInvisible(char c)
        {
            if (c < 32) return true;
            if (c == 127 || c == 0xA0 || c == 0x3000) return true;
            if (c == 0x200B || c == 0x200C || c == 0x200D || c == 0xFEFF) return true;
            if (c >= 0x2000 && c <= 0x200F) return true;
            return false;
        }

        internal static string Keep(string s, Func<char, bool> p)
        {
            StringBuilder sb = new StringBuilder(s.Length);
            foreach (char c in s) if (p(c)) sb.Append(c);
            return sb.ToString();
        }

        internal static string Drop(string s, Func<char, bool> p)
        {
            StringBuilder sb = new StringBuilder(s.Length);
            foreach (char c in s) if (!p(c)) sb.Append(c);
            return sb.ToString();
        }

        private static bool Apply(Func<string, string> f)
        {
            Xl.MapSel(f, "要处理的单元格区域");
            return true;
        }

        /// <summary>Write a block anchored at (anchor + dRow, anchor + dCol), 0-based offsets.</summary>
        internal static void WriteBlock(dynamic anchor, int dRow, int dCol, object[,] data)
        {
            int r = data.GetLength(0), c = data.GetLength(1);
            if (r <= 0 || c <= 0) return;
            dynamic ws = anchor.Worksheet;
            int row = (int)anchor.Row + dRow;
            int col = (int)anchor.Column + dCol;
            if (r == 1 && c == 1)
            {
                ws.Cells[row, col].Value2 = data[0, 0];
                return;
            }
            dynamic a = ws.Cells[row, col];
            dynamic b = ws.Cells[row + r - 1, col + c - 1];
            dynamic rng = ws.Range(a, b);
            rng.Value2 = data;
        }

        internal static void Register()
        {
            // ---- keep / delete by character class ----
            Ops.Reg("t.keep.cn", () => Apply(s => Keep(s, IsCn)));
            Ops.Reg("t.keep.en", () => Apply(s => Keep(s, IsEn)));
            Ops.Reg("t.keep.num", () => Apply(s => Keep(s, IsDig)));
            Ops.Reg("t.keep.sym", () => Apply(s => Keep(s, IsSym)));
            Ops.Reg("t.del.cn", () => Apply(s => Drop(s, IsCn)));
            Ops.Reg("t.del.en", () => Apply(s => Drop(s, IsEn)));
            Ops.Reg("t.del.num", () => Apply(s => Drop(s, IsDig)));
            Ops.Reg("t.del.sym", () => Apply(s => Drop(s, IsSym)));

            // ---- whitespace ----
            Ops.Reg("t.trim.all", () => Apply(s => Drop(s, IsBlank)));
            Ops.Reg("t.trim.head", () => Apply(s => s.TrimStart(' ', '\t', '\u3000', '\r', '\n')));
            Ops.Reg("t.trim.tail", () => Apply(s => s.TrimEnd(' ', '\t', '\u3000', '\r', '\n')));
            Ops.Reg("t.del.invisible", () => Apply(s => Drop(s, IsInvisible)));

            Ops.Reg("t.del.text", DelText);

            // ---- insert ----
            Ops.Reg("t.ins.head", InsHead);
            Ops.Reg("t.ins.tail", InsTail);
            Ops.Reg("t.ins.mid", InsMid);

            // ---- cut ----
            Ops.Reg("t.cut.head", CutHead);
            Ops.Reg("t.cut.tail", CutTail);
            Ops.Reg("t.cut.mid", CutMid);

            // ---- case ----
            Ops.Reg("t.case.upper", () => Apply(s => s.ToUpperInvariant()));
            Ops.Reg("t.case.lower", () => Apply(s => s.ToLowerInvariant()));
            Ops.Reg("t.case.proper", () => Apply(Proper));
            Ops.Reg("t.case.invert", () => Apply(InvertCase));

            // ---- more ----
            Ops.Reg("t.reverse", () => Apply(Reverse));
            Ops.Reg("t.len", Len);
            Ops.Reg("t.freq", Freq);
            Ops.Reg("t.rx.extract", RxExtract);
            Ops.Reg("t.rx.replace", RxReplace);

            // ---- content extraction ----
            Ops.Reg("t.ext.phone", () => Apply(s => FirstMatch(s, @"1[3-9]\d{9}")));
            Ops.Reg("t.ext.email", () => Apply(s => FirstMatch(s, @"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")));
            Ops.Reg("t.ext.url", () => Apply(s => FirstMatch(s, @"(?:https?://|www\.)[^\s，。；、]+")));
            Ops.Reg("t.ext.num", () => Apply(ExtractNum));
            Ops.Reg("t.ext.date", () => Apply(ExtractDate));

            // ---- split / join ----
            Ops.Reg("t.split", Split);
            Ops.Reg("t.join", Join);
        }

        // ---------------- implementations ----------------
        private static bool DelText()
        {
            string[] r = Ask.Show("删除指定文本", new string[] { "要删除的文本(多个用 | 分隔)" }, new string[] { "" });
            if (r == null) return true;
            string spec = r[0];
            if (string.IsNullOrEmpty(spec)) return true;
            string[] parts = spec.Split('|');
            return Apply(delegate(string s)
            {
                foreach (string p in parts)
                {
                    if (p.Length == 0) continue;
                    s = s.Replace(p, "");
                }
                return s;
            });
        }

        private static bool InsHead()
        {
            string v = Ask.One("插入文本", "插入到开头的文本", "");
            if (v == null || v.Length == 0) return true;
            return Apply(s => v + s);
        }

        private static bool InsTail()
        {
            string v = Ask.One("插入文本", "插入到末尾的文本", "");
            if (v == null || v.Length == 0) return true;
            return Apply(s => s + v);
        }

        private static bool InsMid()
        {
            string[] r = Ask.Show("插入文本", new string[] { "插入位置(第几个字符后)", "要插入的文本" },
                                   new string[] { "1", "" });
            if (r == null) return true;
            int pos;
            if (!int.TryParse(r[0], out pos) || pos < 0) { Xl.Bad("位置必须是 0 或正整数。"); return true; }
            string v = r[1];
            if (v.Length == 0) return true;
            return Apply(delegate(string s)
            {
                int p = pos > s.Length ? s.Length : pos;
                return s.Insert(p, v);
            });
        }

        private static bool CutHead()
        {
            int n;
            if (!Ask.Int("截取文本", "保留开头多少个字符", "1", out n)) return true;
            if (n < 0) { Xl.Bad("数量不能为负。"); return true; }
            return Apply(s => s.Length <= n ? s : s.Substring(0, n));
        }

        private static bool CutTail()
        {
            int n;
            if (!Ask.Int("截取文本", "保留末尾多少个字符", "1", out n)) return true;
            if (n < 0) { Xl.Bad("数量不能为负。"); return true; }
            return Apply(s => s.Length <= n ? s : s.Substring(s.Length - n));
        }

        private static bool CutMid()
        {
            string[] r = Ask.Show("截取文本", new string[] { "起始位置(从1开始)", "截取长度" },
                                   new string[] { "1", "1" });
            if (r == null) return true;
            int start, len;
            if (!int.TryParse(r[0], out start) || !int.TryParse(r[1], out len)) { Xl.Bad("请输入整数。"); return true; }
            if (start < 1 || len < 1) { Xl.Bad("起始位置与长度必须 >= 1。"); return true; }
            return Apply(delegate(string s)
            {
                int p = start - 1;
                if (p >= s.Length) return "";
                int n = len;
                if (p + n > s.Length) n = s.Length - p;
                return s.Substring(p, n);
            });
        }

        private static string Proper(string s)
        {
            StringBuilder sb = new StringBuilder(s.Length);
            bool start = true;
            foreach (char c in s)
            {
                if (char.IsLetter(c))
                {
                    sb.Append(start ? char.ToUpperInvariant(c) : char.ToLowerInvariant(c));
                    start = false;
                }
                else
                {
                    sb.Append(c);
                    start = true;
                }
            }
            return sb.ToString();
        }

        private static string InvertCase(string s)
        {
            StringBuilder sb = new StringBuilder(s.Length);
            foreach (char c in s)
            {
                if (c >= 'A' && c <= 'Z') sb.Append(char.ToLowerInvariant(c));
                else if (c >= 'a' && c <= 'z') sb.Append(char.ToUpperInvariant(c));
                else sb.Append(c);
            }
            return sb.ToString();
        }

        private static string Reverse(string s)
        {
            char[] a = s.ToCharArray();
            Array.Reverse(a);
            return new string(a);
        }

        private static bool Len()
        {
            dynamic r = Xl.SelRange("要计算长度的单列区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1)
            {
                Xl.Bad("请只选择一列。");
                return true;
            }
            object[,] outv = new object[rows, 1];
            for (int i = 0; i < rows; i++) outv[i, 0] = Xl.S(a[i, 0]).Length;
            Xl.Turn(true);
            try { WriteBlock(r, 0, 1, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("已把文本长度写入右侧一列。");
            return true;
        }

        private static bool Freq()
        {
            dynamic r = Xl.SelRange("要统计词频的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            Dictionary<string, int> map = new Dictionary<string, int>(StringComparer.Ordinal);
            List<string> order = new List<string>();
            Regex rx = new Regex(@"[^\p{L}\p{Nd}]+", RegexOptions.Compiled);
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    string s = Xl.S(a[i, j]);
                    if (s.Length == 0) continue;
                    foreach (string w in rx.Split(s))
                    {
                        if (w.Length == 0) continue;
                        if (map.ContainsKey(w)) map[w] = map[w] + 1;
                        else { map[w] = 1; order.Add(w); }
                    }
                }
            }
            if (order.Count == 0) { Xl.Say("没有可统计的词。"); return true; }
            order.Sort(delegate(string x, string y)
            {
                int c = map[y].CompareTo(map[x]);
                return c != 0 ? c : string.CompareOrdinal(x, y);
            });
            object[,] outv = new object[order.Count + 1, 2];
            outv[0, 0] = "词";
            outv[0, 1] = "次数";
            for (int i = 0; i < order.Count; i++)
            {
                outv[i + 1, 0] = order[i];
                outv[i + 1, 1] = map[order[i]];
            }
            dynamic sh = Xl.AddSheet("词频统计");
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("已生成工作表「" + (string)sh.Name + "」，共 " + order.Count.ToString(CultureInfo.InvariantCulture) + " 个词。");
            return true;
        }

        private static bool RxExtract()
        {
            string pattern = Ask.One("正则提取", "正则表达式（提取匹配到的内容）", @"\d+");
            if (pattern == null || pattern.Length == 0) return true;
            Regex rx;
            try { rx = new Regex(pattern); }
            catch (Exception ex) { Xl.Bad("正则无效：" + ex.Message); return true; }
            return Apply(delegate(string s)
            {
                MatchCollection ms = rx.Matches(s);
                if (ms.Count == 0) return s;
                StringBuilder sb = new StringBuilder();
                foreach (Match m in ms)
                {
                    if (sb.Length > 0) sb.Append(' ');
                    sb.Append(m.Value);
                }
                return sb.ToString();
            });
        }

        private static bool RxReplace()
        {
            string[] r = Ask.Show("正则替换", new string[] { "正则表达式", "替换为(可用 $1 $2)" },
                                   new string[] { @"\s+", "" });
            if (r == null) return true;
            Regex rx;
            try { rx = new Regex(r[0]); }
            catch (Exception ex) { Xl.Bad("正则无效：" + ex.Message); return true; }
            string rep = r[1];
            return Apply(s => rx.Replace(s, rep));
        }

        private static string FirstMatch(string s, string pattern)
        {
            Match m = Regex.Match(s, pattern);
            return m.Success ? m.Value : s;
        }

        private static string ExtractNum(string s)
        {
            StringBuilder sb = new StringBuilder();
            bool dot = false;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c >= '0' && c <= '9') { sb.Append(c); continue; }
                if ((c == '-' || c == '－') && sb.Length == 0) { sb.Append('-'); continue; }
                if ((c == '.' || c == '．') && !dot && sb.Length > 0) { sb.Append('.'); dot = true; continue; }
            }
            string v = sb.ToString();
            if (v.Length == 0 || v == "-") return s;
            return v;
        }

        private static string ExtractDate(string s)
        {
            Match m = Regex.Match(s, @"(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})");
            if (!m.Success) return s;
            int y = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            int mo = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            int d = int.Parse(m.Groups[3].Value, CultureInfo.InvariantCulture);
            if (mo < 1 || mo > 12 || d < 1 || d > 31) return s;
            return string.Format(CultureInfo.InvariantCulture, "{0:0000}-{1:00}-{2:00}", y, mo, d);
        }

        private static bool Split()
        {
            dynamic r = Xl.SelRange("要分列的单列区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1)
            {
                Xl.Bad("请只选择一列数据。");
                return true;
            }
            string sep = Ask.One("按分隔符分列", "分隔符(支持 \\t 表示制表符)", ",");
            if (sep == null || sep.Length == 0) return true;
            sep = sep.Replace("\\t", "\t");

            List<string[]> parts = new List<string[]>(rows);
            int max = 1;
            for (int i = 0; i < rows; i++)
            {
                string s = Xl.S(a[i, 0]);
                string[] p = s.Length == 0 ? new string[] { "" } : s.Split(new string[] { sep }, StringSplitOptions.None);
                parts.Add(p);
                if (p.Length > max) max = p.Length;
            }
            if (max <= 1) { Xl.Say("没有找到分隔符「" + sep + "」。"); return true; }

            object[,] outv = new object[rows, max];
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < max; j++)
                    outv[i, j] = j < parts[i].Length ? parts[i][j] : "";
            }
            Xl.Turn(true);
            try { WriteBlock(r, 0, 1, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("已拆分为 " + max.ToString(CultureInfo.InvariantCulture) + " 列，写入区域右侧。");
            return true;
        }

        private static bool Join()
        {
            dynamic r = Xl.SelRange("要合并的多列区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols < 2)
            {
                Xl.Bad("请选择至少两列。");
                return true;
            }
            string sep = Ask.One("多列合并", "连接符(可留空)", "");
            if (sep == null) return true;

            object[,] outv = new object[rows, 1];
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
                outv[i, 0] = sb.ToString();
            }
            Xl.Turn(true);
            try { WriteBlock(r, 0, cols, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("已合并写入选区右侧。");
            return true;
        }
    }
}

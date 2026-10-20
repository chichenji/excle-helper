using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcleHelper
{
    /// <summary>Number domain: sequences, rounding, number/Chinese/RMB conversion, batch math.</summary>
    internal static class NumOps
    {
        // ---------------- shared plumbing ----------------
        private static void MapV(Func<object, object> f, string what, string doneMsg)
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
                        object old = a[i, j];
                        if (Xl.IsFormula(old)) continue;
                        object nv = f(old);
                        if (Equals(nv, old)) continue;
                        a[i, j] = nv;
                        n++;
                    }
                }
                if (n > 0) r.Value2 = a;
            }
            finally { Xl.Turn(false); }

            if (n == 0) Xl.Say("没有可处理的内容。");
            else Xl.Say(Xl.Fmt(doneMsg, n));
        }

        private static void MapNum(Func<double, double> f, string what)
        {
            MapV(delegate(object v)
            {
                if (!Xl.IsNumber(v)) return v;
                double d = Xl.D(v);
                double nd = f(d);
                if (Math.Abs(nd - d) < 1e-12) return v;
                return nd;
            }, what, "已处理 {n} 个数字。");
        }

        // ---------------- sequences ----------------
        private static void FillSeq(Func<int, object> gen, string what)
        {
            dynamic r = Xl.SelRange(what);
            if (r == null) return;
            int rows, cols;
            Xl.Vals(r, out rows, out cols);
            object[,] outv = new object[rows, cols];
            int k = 0;
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                    outv[i, j] = gen(k++);
            Xl.Turn(true);
            try { r.Value2 = outv; }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已填充 {n} 个单元格。", rows * cols));
        }

        internal static string Roman(int n)
        {
            if (n <= 0 || n > 3999) return n.ToString(CultureInfo.InvariantCulture);
            int[] val = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
            string[] sym = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < val.Length; i++)
            {
                while (n >= val[i]) { sb.Append(sym[i]); n -= val[i]; }
            }
            return sb.ToString();
        }

        internal static string Letter(int n)
        {
            if (n <= 0) return "";
            StringBuilder sb = new StringBuilder();
            while (n > 0)
            {
                int m = (n - 1) % 26;
                sb.Insert(0, (char)('A' + m));
                n = (n - 1) / 26;
            }
            return sb.ToString();
        }

        private static string[] CnDigits = { "零", "一", "二", "三", "四", "五", "六", "七", "八", "九" };
        private static string[] CnUnits = { "", "十", "百", "千" };
        private static string[] CnGroups = { "", "万", "亿", "万亿" };

        internal static string NumToCn(long v)
        {
            if (v == 0) return "零";
            bool neg = v < 0;
            v = Math.Abs(v);
            List<int> parts = new List<int>();
            long t = v;
            while (t > 0)
            {
                parts.Add((int)(t % 10000));
                t /= 10000;
            }
            StringBuilder sb = new StringBuilder();
            bool zero = false;
            for (int i = parts.Count - 1; i >= 0; i--)
            {
                int part = parts[i];
                if (part == 0)
                {
                    if (sb.Length > 0) zero = true;
                    continue;
                }
                if (sb.Length > 0 && (zero || part < 1000))
                {
                    if (!sb.ToString().EndsWith("零", StringComparison.Ordinal)) sb.Append("零");
                }
                sb.Append(PartCn(part, CnDigits));
                sb.Append(CnGroups[i]);
                zero = false;
            }
            string s = sb.ToString();
            s = Regex.Replace(s, "零+", "零");
            if (s.StartsWith("一十", StringComparison.Ordinal)) s = s.Substring(1);
            return (neg ? "负" : "") + s;
        }

        private static string PartCn(int part, string[] digits)
        {
            StringBuilder sb = new StringBuilder();
            string s = part.ToString(CultureInfo.InvariantCulture);
            int len = s.Length;
            bool zero = false;
            for (int i = 0; i < len; i++)
            {
                int d = s[i] - '0';
                int unit = len - 1 - i;
                if (d == 0)
                {
                    if (sb.Length > 0) zero = true;
                    continue;
                }
                if (zero) { sb.Append("零"); zero = false; }
                sb.Append(digits[d]);
                sb.Append(CnUnits[unit]);
            }
            return sb.ToString();
        }

        /// <summary>Parse Chinese numerals (一二三…十百千万亿, 两). Returns null when not a numeral.</summary>
        internal static double? CnToNum(string s)
        {
            if (string.IsNullOrEmpty(s)) return null;
            long total = 0, section = 0, number = 0;
            bool any = false;
            foreach (char c in s.Trim())
            {
                int d = Array.IndexOf(CnDigits, c.ToString());
                if (d < 0 && c == '两') d = 2;
                if (d >= 0) { number = d; any = true; continue; }

                int mul = 0;
                if (c == '十') mul = 10;
                else if (c == '百') mul = 100;
                else if (c == '千') mul = 1000;

                if (mul > 0)
                {
                    if (number == 0) number = 1;
                    section += number * mul;
                    number = 0;
                    any = true;
                    continue;
                }
                if (c == '万')
                {
                    total += (section + number) * 10000L;
                    section = 0; number = 0; any = true;
                    continue;
                }
                if (c == '亿')
                {
                    total = (total + section + number) * 100000000L;
                    section = 0; number = 0; any = true;
                    continue;
                }
                if (c == '零' || c == '〇') { number = 0; any = true; continue; }
                return null;
            }
            if (!any) return null;
            return (double)(total + section + number);
        }

        // ---------------- RMB ----------------
        private static readonly string[] RmbDigits = { "零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖" };
        private static readonly string[] RmbUnits = { "", "拾", "佰", "仟" };
        private static readonly string[] RmbGroups = { "", "万", "亿", "万亿" };

        internal static string RmbUpper(decimal money)
        {
            if (money == 0m) return "零元整";
            bool neg = money < 0m;
            money = Math.Abs(money);
            money = Math.Round(money, 2, MidpointRounding.AwayFromZero);
            long yuan = (long)Math.Floor(money);
            int cents = (int)Math.Round((money - yuan) * 100m);
            if (cents >= 100) { yuan++; cents -= 100; }

            StringBuilder sb = new StringBuilder();
            sb.Append(yuan == 0 ? "零" : RmbInt(yuan));
            sb.Append("元");

            if (cents == 0)
            {
                sb.Append("整");
            }
            else
            {
                int jiao = cents / 10;
                int fen = cents % 10;
                if (jiao > 0) { sb.Append(RmbDigits[jiao]); sb.Append("角"); }
                else if (yuan > 0 && fen > 0) sb.Append("零");
                if (fen > 0) { sb.Append(RmbDigits[fen]); sb.Append("分"); }
            }
            return (neg ? "负" : "") + sb.ToString();
        }

        private static string RmbInt(long v)
        {
            List<int> parts = new List<int>();
            long t = v;
            while (t > 0) { parts.Add((int)(t % 10000)); t /= 10000; }
            StringBuilder sb = new StringBuilder();
            bool zero = false;
            for (int i = parts.Count - 1; i >= 0; i--)
            {
                int part = parts[i];
                if (part == 0) { if (sb.Length > 0) zero = true; continue; }
                if (sb.Length > 0 && (zero || part < 1000))
                {
                    if (!sb.ToString().EndsWith("零", StringComparison.Ordinal)) sb.Append("零");
                }
                sb.Append(PartCn(part, RmbDigits));
                sb.Append(RmbGroups[i]);
                zero = false;
            }
            string s = sb.ToString();
            s = Regex.Replace(s, "零+", "零");
            return s;
        }

        // ---------------- fraction ----------------
        internal static string ToFraction(double v, long maxDen)
        {
            if (Math.Abs(v) < 1e-12) return "0";
            bool neg = v < 0;
            v = Math.Abs(v);
            long h1 = 1, h2 = 0, k1 = 0, k2 = 1;
            double b = v;
            for (int i = 0; i < 64; i++)
            {
                long a = (long)Math.Floor(b);
                long h = a * h1 + h2;
                long k = a * k1 + k2;
                if (k > maxDen) break;
                h2 = h1; h1 = h; k2 = k1; k1 = k;
                double frac = b - a;
                if (frac < 1e-12) break;
                b = 1.0 / frac;
            }
            if (k1 == 0) return v.ToString("0.########", CultureInfo.InvariantCulture);
            if (k1 == 1) return (neg ? "-" : "") + h1.ToString(CultureInfo.InvariantCulture);
            return (neg ? "-" : "") + h1.ToString(CultureInfo.InvariantCulture) + "/" + k1.ToString(CultureInfo.InvariantCulture);
        }

        // ---------------- registration ----------------
        internal static void Register()
        {
            Ops.Reg("n.seq.num", SeqNum);
            Ops.Reg("n.seq.letter", () => { FillSeq(Letter, "要填充序列的区域"); return true; });
            Ops.Reg("n.seq.week", SeqWeek);
            Ops.Reg("n.seq.month", SeqMonth);
            Ops.Reg("n.seq.cnnum", () => { FillSeq(i => NumToCn(i + 1), "要填充序列的区域"); return true; });
            Ops.Reg("n.seq.roman", () => { FillSeq(i => Roman(i + 1), "要填充序列的区域"); return true; });

            Ops.Reg("n.round", RoundN);
            Ops.Reg("n.roundup", RoundUp);
            Ops.Reg("n.rounddown", RoundDown);
            Ops.Reg("n.int", () => { MapNum(Math.Floor, "要取整的数字区域"); return true; });
            Ops.Reg("n.frac", Frac);

            Ops.Reg("n.math", MathBatch);
            Ops.Reg("n.totext", ToNumber);
            Ops.Reg("n.cntoar", CnToAr);
            Ops.Reg("n.artoCN", ArToCn);
            Ops.Reg("n.rmb", Rmb);
            Ops.Reg("n.display", Display);
            Ops.Reg("n.fillblank", FillBlank);
        }

        private static bool SeqNum()
        {
            string[] r = Ask.Show("录入数字序列", new string[] { "起始值", "步长" }, new string[] { "1", "1" });
            if (r == null) return true;
            double start, step;
            if (!double.TryParse(r[0], NumberStyles.Any, CultureInfo.InvariantCulture, out start) ||
                !double.TryParse(r[1], NumberStyles.Any, CultureInfo.InvariantCulture, out step))
            {
                Xl.Bad("请输入数字。");
                return true;
            }
            FillSeq(i => (object)(start + step * i), "要填充序列的区域");
            return true;
        }

        private static bool SeqWeek()
        {
            string[] names = { "星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日" };
            FillSeq(i => names[i % 7], "要填充序列的区域");
            return true;
        }

        private static bool SeqMonth()
        {
            FillSeq(i => (i % 12 + 1).ToString(CultureInfo.InvariantCulture) + "月", "要填充序列的区域");
            return true;
        }

        private static bool RoundN()
        {
            int n;
            if (!Ask.Int("四舍五入", "保留小数位数", "2", out n)) return true;
            if (n < 0 || n > 15) { Xl.Bad("位数需在 0-15 之间。"); return true; }
            MapNum(d => Math.Round(d, n, MidpointRounding.AwayFromZero), "要四舍五入的数字区域");
            return true;
        }

        private static bool RoundUp()
        {
            int n;
            if (!Ask.Int("向上舍入", "保留小数位数", "0", out n)) return true;
            double p = Math.Pow(10, n);
            MapNum(d => Math.Ceiling(d * p) / p, "要向上舍入的数字区域");
            return true;
        }

        private static bool RoundDown()
        {
            int n;
            if (!Ask.Int("向下舍入", "保留小数位数", "0", out n)) return true;
            double p = Math.Pow(10, n);
            MapNum(d => Math.Floor(d * p) / p, "要向下舍入的数字区域");
            return true;
        }

        private static bool Frac()
        {
            int den;
            if (!Ask.Int("小数转分数", "最大分母", "1000", out den)) return true;
            if (den < 1) { Xl.Bad("分母必须为正整数。"); return true; }
            MapV(delegate(object v)
            {
                if (!Xl.IsNumber(v)) return v;
                double d = Xl.D(v);
                if (Math.Abs(d - Math.Round(d)) < 1e-12) return v;
                return ToFraction(d, den);
            }, "要转换的小数区域", "已转换 {n} 个数字。");
            return true;
        }

        private static bool MathBatch()
        {
            string[] r = Ask.Show("批量运算", new string[] { "运算(+ - * /)", "数值" }, new string[] { "+", "1" });
            if (r == null) return true;
            string op = r[0].Trim();
            double v;
            if (!double.TryParse(r[1], NumberStyles.Any, CultureInfo.InvariantCulture, out v))
            {
                Xl.Bad("请输入数值。");
                return true;
            }
            Func<double, double> f;
            if (op == "+") f = d => d + v;
            else if (op == "-") f = d => d - v;
            else if (op == "*") f = d => d * v;
            else if (op == "/") { if (v == 0) { Xl.Bad("不能除以 0。"); return true; } f = d => d / v; }
            else { Xl.Bad("运算只能是 + - * / 。"); return true; }
            MapNum(f, "要运算的数字区域");
            return true;
        }

        private static bool ToNumber()
        {
            MapV(delegate(object v)
            {
                if (v == null || v is DBNull) return v;
                if (v is double || v is int || v is long || v is decimal) return v;
                string s = Xl.S(v).Trim();
                if (s.Length == 0) return v;
                double d;
                if (!double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d)) return v;
                return d;
            }, "要转换的文本型数字区域", "已转换 {n} 个单元格。");
            return true;
        }

        private static bool CnToAr()
        {
            MapV(delegate(object v)
            {
                string s = Xl.S(v).Trim();
                if (s.Length == 0) return v;
                double? d = CnToNum(s);
                if (!d.HasValue) return v;
                return d.Value;
            }, "要转换的中文数字区域", "已转换 {n} 个单元格。");
            return true;
        }

        private static bool ArToCn()
        {
            MapV(delegate(object v)
            {
                if (!Xl.IsNumber(v)) return v;
                double d = Xl.D(v);
                if (Math.Abs(d - Math.Round(d)) > 1e-9) return v;
                long n = (long)Math.Round(d);
                if (Math.Abs(n) > 999999999999999L) return v;
                return NumToCn(n);
            }, "要转换的数字区域", "已转换 {n} 个单元格。");
            return true;
        }

        private static bool Rmb()
        {
            MapV(delegate(object v)
            {
                if (!Xl.IsNumber(v)) return v;
                decimal m;
                try { m = Convert.ToDecimal(Xl.D(v)); } catch { return v; }
                return RmbUpper(m);
            }, "要转换金额的区域", "已转换 {n} 个金额。");
            return true;
        }

        private static bool Display()
        {
            dynamic r = Xl.SelRange("要保留显示值的区域");
            if (r == null) return true;
            int rows, cols;
            Xl.Vals(r, out rows, out cols);
            if (rows * cols > 200000)
            {
                if (!Xl.Ask("选区有 " + (rows * cols).ToString(CultureInfo.InvariantCulture)
                            + " 个单元格，可能较慢。继续？")) return true;
            }
            object[,] outv = new object[rows, cols];
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        dynamic cell = r.Cells[i + 1, j + 1];
                        outv[i, j] = (string)cell.Text;
                    }
                }
                r.Value2 = outv;
            }
            finally { Xl.Turn(false); }
            Xl.Say("已用显示值替换原内容（公式已转为文本）。");
            return true;
        }

        private static bool FillBlank()
        {
            dynamic r = Xl.SelRange("要填充空白的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int n = 0;
            for (int j = 0; j < cols; j++)
            {
                object last = null;
                for (int i = 0; i < rows; i++)
                {
                    object v = a[i, j];
                    string s = Xl.S(v);
                    if (s.Length == 0)
                    {
                        if (last != null) { a[i, j] = last; n++; }
                    }
                    else last = v;
                }
            }
            if (n == 0) { Xl.Say("没有空白单元格。"); return true; }
            Xl.Turn(true);
            try { r.Value2 = a; }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已填充 {n} 个空白单元格。", n));
            return true;
        }
    }
}

using System;
using System.Globalization;
using System.Text;

namespace ExcleHelper
{
    /// <summary>补充能力：补零、英文金额、文本打码、压缩空格、百分比、千分位、银行家舍入、随机分摊。</summary>
    internal static class ExtrasOps
    {
        internal static void Register()
        {
            Ops.Reg("t.padleft", PadLeft);
            Ops.Reg("t.num2en", NumToEnglish);
            Ops.Reg("t.mask.mid", MaskMiddle);
            Ops.Reg("t.squeeze", Squeeze);
            Ops.Reg("n.percent", Percent);
            Ops.Reg("n.thousand", Thousand);
            Ops.Reg("n.bankers", Bankers);
            Ops.Reg("n.randsplit", RandSplit);
        }

        // ---------------- text ----------------
        private static bool PadLeft()
        {
            string[] p = Ask.Show("补零（定长编码）", new string[] { "目标长度", "填充字符" }, new string[] { "6", "0" });
            if (p == null) return true;
            int len;
            if (!int.TryParse(p[0], out len) || len < 1) { Xl.Bad("长度必须是正整数。"); return true; }
            string pad = p[1].Length == 0 ? "0" : p[1].Substring(0, 1);
            Xl.MapSel(delegate(string s)
            {
                s = s.Trim();
                if (s.Length >= len) return s;
                StringBuilder sb = new StringBuilder();
                for (int i = s.Length; i < len; i++) sb.Append(pad);
                return sb.ToString() + s;
            }, "要补零的区域");
            return true;
        }

        private static bool Squeeze()
        {
            Xl.MapSel(delegate(string s)
            {
                StringBuilder sb = new StringBuilder(s.Length);
                bool prevSpace = false;
                foreach (char c in s)
                {
                    bool isSpace = (c == ' ' || c == '\t' || c == '\u3000');
                    if (isSpace)
                    {
                        if (prevSpace) continue;
                        sb.Append(' ');
                        prevSpace = true;
                    }
                    else
                    {
                        sb.Append(c);
                        prevSpace = false;
                    }
                }
                return sb.ToString();
            }, "要压缩空格的多余区域");
            return true;
        }

        private static bool MaskMiddle()
        {
            string[] p = Ask.Show("文本中间打码", new string[] { "保留前几位", "保留后几位", "打码字符" },
                                   new string[] { "1", "1", "*" });
            if (p == null) return true;
            int head, tail;
            if (!int.TryParse(p[0], out head) || !int.TryParse(p[1], out tail) || head < 0 || tail < 0)
            {
                Xl.Bad("请输入非负整数。");
                return true;
            }
            string mask = p[2].Length == 0 ? "*" : p[2].Substring(0, 1);
            Xl.MapSel(delegate(string s)
            {
                s = s.Trim();
                if (s.Length <= head + tail) return s;
                StringBuilder sb = new StringBuilder();
                sb.Append(s.Substring(0, head));
                for (int i = head; i < s.Length - tail; i++) sb.Append(mask);
                sb.Append(s.Substring(s.Length - tail));
                return sb.ToString();
            }, "要打码的区域");
            return true;
        }

        // ---------------- number ----------------
        private static bool Percent()
        {
            string d = Ask.One("转百分比", "保留小数位", "2");
            if (d == null) return true;
            int n;
            if (!int.TryParse(d, out n) || n < 0) { Xl.Bad("小数位必须 >= 0。"); return true; }
            Xl.MapSel(delegate(string s)
            {
                double v;
                if (!double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out v)) return s;
                return (v * 100).ToString("F" + n, CultureInfo.InvariantCulture) + "%";
            }, "要转百分比的小数区域");
            return true;
        }

        private static bool Thousand()
        {
            Xl.MapSel(delegate(string s)
            {
                double v;
                if (!double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out v)) return s;
                if (Math.Abs(v - Math.Round(v)) < 1e-9)
                    return ((long)Math.Round(v)).ToString("N0", CultureInfo.InvariantCulture);
                return v.ToString("N2", CultureInfo.InvariantCulture);
            }, "要加千分位的数字区域");
            return true;
        }

        private static bool Bankers()
        {
            string d = Ask.One("四舍六入五成双", "保留小数位", "0");
            if (d == null) return true;
            int n;
            if (!int.TryParse(d, out n) || n < 0 || n > 15) { Xl.Bad("小数位需在 0-15。"); return true; }
            Xl.MapSel(delegate(string s)
            {
                double v;
                if (!double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out v)) return s;
                return Math.Round(v, n, MidpointRounding.ToEven).ToString(CultureInfo.InvariantCulture);
            }, "要按银行家舍入的数字区域");
            return true;
        }

        private static bool RandSplit()
        {
            dynamic r = Xl.SelRange("要分摊的金额区域");
            if (r == null) return true;
            int parts;
            if (!Ask.Int("随机分摊", "每笔拆成几份", "3", out parts)) return true;
            if (parts < 2 || parts > 50) { Xl.Bad("份数需在 2-50 之间。"); return true; }

            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列金额。"); return true; }
            object[,] outv = new object[rows, parts];
            Random rnd = new Random(Guid.NewGuid().GetHashCode());
            for (int i = 0; i < rows; i++)
            {
                if (!Xl.IsNumber(a[i, 0])) continue;
                double total = Math.Round(Xl.D(a[i, 0]), 2);
                double remain = total;
                double[] weights = new double[parts];
                double sum = 0;
                for (int k = 0; k < parts; k++) { weights[k] = rnd.NextDouble() + 0.01; sum += weights[k]; }
                for (int k = 0; k < parts; k++)
                {
                    double v;
                    if (k == parts - 1) v = Math.Round(remain, 2);
                    else { v = Math.Round(total * weights[k] / sum, 2); remain -= v; }
                    outv[i, k] = v;
                }
            }
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, 1, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("已随机分摊为 " + parts + " 份，写入选区右侧（末份为差额，合计不变）。");
            return true;
        }

        // ---------------- number -> english amount ----------------
        private static readonly string[] Ones =
        {
            "", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine",
            "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen",
            "seventeen", "eighteen", "nineteen"
        };

        private static readonly string[] Tens =
        { "", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety" };

        private static string Say(long n)
        {
            if (n < 20) return Ones[n];
            if (n < 100) return Tens[n / 10] + (n % 10 > 0 ? "-" + Ones[n % 10] : "");
            if (n < 1000)
                return Ones[n / 100] + " hundred" + (n % 100 > 0 ? " and " + Say(n % 100) : "");
            long[] units = { 1000000000L, 1000000L, 1000L };
            string[] names = { "billion", "million", "thousand" };
            for (int i = 0; i < units.Length; i++)
            {
                if (n >= units[i])
                    return Say(n / units[i]) + " " + names[i]
                           + (n % units[i] > 0 ? " " + Say(n % units[i]) : "");
            }
            return n.ToString(CultureInfo.InvariantCulture);
        }

        internal static string EnglishAmount(decimal money)
        {
            bool neg = money < 0m;
            money = Math.Abs(money);
            money = Math.Round(money, 2, MidpointRounding.AwayFromZero);
            long dollars = (long)Math.Floor(money);
            int cents = (int)Math.Round((money - dollars) * 100m);
            if (cents >= 100) { dollars++; cents -= 100; }
            StringBuilder sb = new StringBuilder();
            sb.Append(dollars == 0 ? "zero" : Say(dollars));
            sb.Append(dollars == 1 ? " dollar" : " dollars");
            if (cents > 0)
            {
                sb.Append(" and ");
                sb.Append(Say(cents));
                sb.Append(cents == 1 ? " cent" : " cents");
            }
            else sb.Append(" only");
            string s = sb.ToString();
            return (neg ? "minus " : "") + char.ToUpperInvariant(s[0]) + s.Substring(1);
        }

        private static bool NumToEnglish()
        {
            Xl.MapSel(delegate(string s)
            {
                decimal m;
                if (!decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out m)) return s;
                return EnglishAmount(m);
            }, "要转英文金额的数字区域");
            return true;
        }
    }
}

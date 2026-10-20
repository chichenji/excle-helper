using System;
using System.Globalization;
using System.Text;

namespace ExcleHelper
{
    /// <summary>Chinese-specific domain: ID card, phone, name, pinyin initial, bank/USCI checks.</summary>
    internal static class CnOps
    {
        // ---------------- ID card ----------------
        private static readonly int[] IdW = { 7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2 };
        private const string IdCheck = "10X98765432";

        internal static bool ValidId(string s)
        {
            if (s == null || s.Length != 18) return false;
            for (int i = 0; i < 17; i++) if (s[i] < '0' || s[i] > '9') return false;
            char last = char.ToUpperInvariant(s[17]);
            int sum = 0;
            for (int i = 0; i < 17; i++) sum += (s[i] - '0') * IdW[i];
            return last == IdCheck[sum % 11];
        }

        internal static string Up15To18(string s)
        {
            if (s == null || s.Length != 15) return s;
            for (int i = 0; i < 15; i++) if (s[i] < '0' || s[i] > '9') return s;
            string t = s.Substring(0, 6) + "19" + s.Substring(6);
            int sum = 0;
            for (int i = 0; i < 17; i++) sum += (t[i] - '0') * IdW[i];
            return t + IdCheck[sum % 11];
        }

        private static string BirthOf(string s)
        {
            string t = s.Length == 15 ? Up15To18(s) : s;
            if (t.Length != 18) return "";
            string y = t.Substring(6, 4), m = t.Substring(10, 2), d = t.Substring(12, 2);
            int iy, im, id;
            if (!int.TryParse(y, out iy) || !int.TryParse(m, out im) || !int.TryParse(d, out id)) return "";
            if (im < 1 || im > 12 || id < 1 || id > 31) return "";
            return string.Format(CultureInfo.InvariantCulture, "{0}-{1}-{2}", y, m, d);
        }

        private static string SexOf(string s)
        {
            string t = s.Length == 15 ? Up15To18(s) : s;
            if (t.Length != 18) return "";
            int d = t[16] - '0';
            if (d < 0 || d > 9) return "";
            return (d % 2 == 1) ? "男" : "女";
        }

        private static int AgeOf(string birth)
        {
            DateTime dt;
            if (!DateTime.TryParse(birth, out dt)) return -1;
            DateTime now = DateTime.Now;
            int a = now.Year - dt.Year;
            if (now.Month < dt.Month || (now.Month == dt.Month && now.Day < dt.Day)) a--;
            return a < 0 ? -1 : a;
        }

        // ---------------- registration ----------------
        internal static void Register()
        {
            Ops.Reg("c.id.info", IdInfo);
            Ops.Reg("c.id.check", IdCheckOp);
            Ops.Reg("c.id.15to18", IdUp);
            Ops.Reg("c.id.mask", IdMask);

            Ops.Reg("c.phone.seg", PhoneSeg);
            Ops.Reg("c.phone.mask", PhoneMask);

            Ops.Reg("c.name.split", NameSplit);
            Ops.Reg("c.name.mask", NameMask);

            Ops.Reg("c.py.first", PyFirstOp);
            Ops.Reg("c.bank.check", BankCheck);
            Ops.Reg("c.usci.check", UsciCheck);
        }

        private static bool IdInfo()
        {
            dynamic r = Xl.SelRange("含身份证号的单列区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列身份证号。"); return true; }
            object[,] outv = new object[rows, 3];
            int ok = 0;
            for (int i = 0; i < rows; i++)
            {
                string s = Xl.S(a[i, 0]).Trim().ToUpperInvariant();
                if (s.Length != 15 && s.Length != 18) continue;
                string birth = BirthOf(s);
                string sex = SexOf(s);
                int age = AgeOf(birth);
                outv[i, 0] = sex;
                outv[i, 1] = birth;
                outv[i, 2] = age < 0 ? (object)"" : age;
                ok++;
            }
            if (ok == 0) { Xl.Bad("没有识别到合法长度的身份证号。"); return true; }
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, 1, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("已解析 " + ok.ToString(CultureInfo.InvariantCulture)
                   + " 个号码：右侧依次写入 性别 / 出生日期 / 年龄。");
            return true;
        }

        private static bool IdCheckOp()
        {
            dynamic r = Xl.SelRange("含身份证号的单列区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列身份证号。"); return true; }
            object[,] outv = new object[rows, 1];
            int good = 0;
            for (int i = 0; i < rows; i++)
            {
                string s = Xl.S(a[i, 0]).Trim().ToUpperInvariant();
                bool ok = ValidId(s);
                outv[i, 0] = ok ? "正确" : "错误";
                if (ok) good++;
            }
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, 1, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("校验完成：" + good.ToString(CultureInfo.InvariantCulture) + " 个正确，"
                   + (rows - good).ToString(CultureInfo.InvariantCulture) + " 个错误。");
            return true;
        }

        private static bool IdUp()
        {
            dynamic r = Xl.SelRange("含 15 位身份证号的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int n = 0;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    string s = Xl.S(a[i, j]).Trim();
                    if (s.Length != 15) continue;
                    string t = Up15To18(s);
                    if (t == s) continue;
                    a[i, j] = t;
                    n++;
                }
            }
            if (n == 0) { Xl.Say("没有 15 位身份证号。"); return true; }
            Xl.Turn(true);
            try { r.Value2 = a; }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已升级 {n} 个号码。", n));
            return true;
        }

        private static bool IdMask()
        {
            dynamic r = Xl.SelRange("含身份证号的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int n = 0;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    string s = Xl.S(a[i, j]).Trim();
                    if (s.Length != 15 && s.Length != 18) continue;
                    string m = s.Substring(0, 6) + new string('*', s.Length - 10) + s.Substring(s.Length - 4);
                    a[i, j] = m;
                    n++;
                }
            }
            if (n == 0) { Xl.Say("没有可处理的身份证号。"); return true; }
            Xl.Turn(true);
            try { r.Value2 = a; }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已处理 {n} 个号码。", n));
            return true;
        }

        // ---------------- phone ----------------
        private static bool PhoneSeg()
        {
            dynamic r = Xl.SelRange("含手机号的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int n = 0;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    string s = DigitsOnly(Xl.S(a[i, j]));
                    if (s.Length != 11) continue;
                    a[i, j] = s.Substring(0, 3) + " " + s.Substring(3, 4) + " " + s.Substring(7);
                    n++;
                }
            }
            if (n == 0) { Xl.Say("没有 11 位手机号。"); return true; }
            Xl.Turn(true);
            try { r.Value2 = a; }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已分段 {n} 个号码。", n));
            return true;
        }

        private static bool PhoneMask()
        {
            dynamic r = Xl.SelRange("含手机号的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int n = 0;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    string s = DigitsOnly(Xl.S(a[i, j]));
                    if (s.Length != 11) continue;
                    a[i, j] = s.Substring(0, 3) + "****" + s.Substring(7);
                    n++;
                }
            }
            if (n == 0) { Xl.Say("没有 11 位手机号。"); return true; }
            Xl.Turn(true);
            try { r.Value2 = a; }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已打码 {n} 个号码。", n));
            return true;
        }

        private static string DigitsOnly(string s)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in s) if (c >= '0' && c <= '9') sb.Append(c);
            return sb.ToString();
        }

        // ---------------- name ----------------
        private static readonly string[] DoubleSurname =
        {
            "欧阳","司马","上官","诸葛","东方","皇甫","尉迟","公孙","慕容","长孙","宇文",
            "司徒","夏侯","令狐","端木","独孤","南宫","西门","百里","呼延","轩辕","澹台",
            "公冶","申屠","太史","闻人","万俟","濮阳","淳于","单于","赫连","皇甫"
        };

        private static bool NameSplit()
        {
            dynamic r = Xl.SelRange("含姓名的单列区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列姓名。"); return true; }
            object[,] outv = new object[rows, 2];
            for (int i = 0; i < rows; i++)
            {
                string s = Xl.S(a[i, 0]).Trim();
                if (s.Length == 0) { outv[i, 0] = ""; outv[i, 1] = ""; continue; }
                int take = 1;
                if (s.Length >= 3)
                {
                    string two = s.Length >= 2 ? s.Substring(0, 2) : "";
                    if (Array.IndexOf(DoubleSurname, two) >= 0) take = 2;
                }
                if (s.Length == 2 && Array.IndexOf(DoubleSurname, s) >= 0) take = 2;
                if (take >= s.Length) { outv[i, 0] = s; outv[i, 1] = ""; continue; }
                outv[i, 0] = s.Substring(0, take);
                outv[i, 1] = s.Substring(take);
            }
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, 1, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("已拆分为 姓 / 名 两列（右侧）。");
            return true;
        }

        private static bool NameMask()
        {
            dynamic r = Xl.SelRange("含姓名的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int n = 0;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    string s = Xl.S(a[i, j]).Trim();
                    if (s.Length == 0) continue;
                    a[i, j] = s.Substring(0, 1) + new string('*', s.Length - 1);
                    n++;
                }
            }
            if (n == 0) { Xl.Say("没有可处理的数据。"); return true; }
            Xl.Turn(true);
            try { r.Value2 = a; }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已打码 {n} 个姓名。", n));
            return true;
        }

        // ---------------- pinyin initial (GB2312 level-1 ranges) ----------------
        private static readonly int[] PyStart =
        {
            0xB0A1,0xB0C5,0xB2C1,0xB4EE,0xB6EA,0xB7A2,0xB8C1,0xB9FE,0xBBF7,0xBFA6,
            0xC0AC,0xC2E8,0xC4C3,0xC5B6,0xC5BE,0xC6DA,0xC8BB,0xC8F6,0xCBFA,0xCDDA,
            0xCEF4,0xD1B9,0xD4D1,0xD7F9
        };

        private const string PyLetters = "ABCDEFGHJKLMNOPQRSTWXYZ";

        internal static char PyOf(char c)
        {
            if (!(c >= 0x4E00 && c <= 0x9FFF)) return c;
            byte[] b;
            try { b = Encoding.GetEncoding(936).GetBytes(c.ToString()); }
            catch { return c; }
            if (b.Length < 2) return c;
            int code = b[0] * 256 + b[1];
            if (code < PyStart[0] || code > PyStart[PyStart.Length - 1]) return c;
            for (int i = PyStart.Length - 2; i >= 0; i--)
            {
                if (code >= PyStart[i]) return PyLetters[i];
            }
            return c;
        }

        private static bool PyFirstOp()
        {
            dynamic r = Xl.SelRange("要去拼音首字母的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int n = 0;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    string s = Xl.S(a[i, j]);
                    if (s.Length == 0) continue;
                    StringBuilder sb = new StringBuilder(s.Length);
                    foreach (char c in s) sb.Append(PyOf(c));
                    a[i, j] = sb.ToString();
                    n++;
                }
            }
            if (n == 0) { Xl.Say("没有可处理的内容。"); return true; }
            Xl.Turn(true);
            try { r.Value2 = a; }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已转换 {n} 个单元格。", n));
            return true;
        }

        // ---------------- bank card (Luhn) ----------------
        private static bool BankCheck()
        {
            dynamic r = Xl.SelRange("含银行卡号的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int good = 0, total = 0;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    string s = DigitsOnly(Xl.S(a[i, j]));
                    if (s.Length < 12 || s.Length > 19) continue;
                    total++;
                    if (Luhn(s)) { a[i, j] = s + " 正确"; good++; }
                    else a[i, j] = s + " 错误";
                }
            }
            if (total == 0) { Xl.Say("没有 12-19 位的卡号。"); return true; }
            Xl.Turn(true);
            try { r.Value2 = a; }
            finally { Xl.Turn(false); }
            Xl.Say("校验完成：" + good.ToString(CultureInfo.InvariantCulture) + " 个正确，"
                   + (total - good).ToString(CultureInfo.InvariantCulture) + " 个错误。");
            return true;
        }

        internal static bool Luhn(string digits)
        {
            int sum = 0;
            bool dbl = false;
            for (int i = digits.Length - 1; i >= 0; i--)
            {
                int d = digits[i] - '0';
                if (dbl)
                {
                    d *= 2;
                    if (d > 9) d -= 9;
                }
                sum += d;
                dbl = !dbl;
            }
            return sum % 10 == 0;
        }

        // ---------------- unified social credit code ----------------
        private const string UsciChars = "0123456789ABCDEFGHJKLMNPQRTUWXY";
        private static readonly int[] UsciW = { 1, 3, 9, 27, 19, 26, 16, 17, 20, 29, 25, 13, 8, 24, 10, 30, 28 };

        internal static bool ValidUsci(string s)
        {
            if (s == null || s.Length != 18) return false;
            s = s.ToUpperInvariant();
            int sum = 0;
            for (int i = 0; i < 17; i++)
            {
                int idx = UsciChars.IndexOf(s[i]);
                if (idx < 0) return false;
                sum += idx * UsciW[i];
            }
            int mod = 31 - (sum % 31);
            if (mod == 31) mod = 0;
            return UsciChars[mod] == s[17];
        }

        private static bool UsciCheck()
        {
            dynamic r = Xl.SelRange("含统一社会信用代码的区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int good = 0, total = 0;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    string s = Xl.S(a[i, j]).Trim();
                    if (s.Length != 18) continue;
                    total++;
                    if (ValidUsci(s)) { a[i, j] = s + " 正确"; good++; }
                    else a[i, j] = s + " 错误";
                }
            }
            if (total == 0) { Xl.Say("没有 18 位信用代码。"); return true; }
            Xl.Turn(true);
            try { r.Value2 = a; }
            finally { Xl.Turn(false); }
            Xl.Say("校验完成：" + good.ToString(CultureInfo.InvariantCulture) + " 个正确，"
                   + (total - good).ToString(CultureInfo.InvariantCulture) + " 个错误。");
            return true;
        }
    }
}

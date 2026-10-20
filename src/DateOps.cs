using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcleHelper
{
    /// <summary>
    /// 日期时间域（批次 1）：标准化、格式、拆分合并、差值、工作日、年龄、时间戳、
    /// 公历农历互转、区间统计、随机日期、清除非日期。
    /// 农历用 .NET 内置 ChineseLunisolarCalendar，不依赖自备数据表。
    /// </summary>
    internal static class DateOps
    {
        private static readonly ChineseLunisolarCalendar Lunar = new ChineseLunisolarCalendar();
        private static readonly string[] WeekNames = { "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六" };

        internal static void Register()
        {
            Ops.Reg("dt.normalize", Normalize);
            Ops.Reg("dt.fmt.date", FmtDate);
            Ops.Reg("dt.fmt.time", FmtTime);
            Ops.Reg("dt.split", Split);
            Ops.Reg("dt.combine", Combine);
            Ops.Reg("dt.diff", Diff);
            Ops.Reg("dt.add", Add);
            Ops.Reg("dt.workdays", WorkDays);
            Ops.Reg("dt.workday.after", WorkDayAfter);
            Ops.Reg("dt.age", Age);
            Ops.Reg("dt.timestamp", Timestamp);
            Ops.Reg("dt.lunar", LunarOp);
            Ops.Reg("dt.range.stat", RangeStat);
            Ops.Reg("dt.rand", RandDate);
            Ops.Reg("dt.clear.invalid", ClearInvalid);
        }

        // ---------------- parsing ----------------
        /// <summary>把单元格内容尽量解析成日期（支持中文年月日、点号、斜杠、8 位数字、Excel 序列号）。</summary>
        internal static bool TryParseDate(object v, out DateTime dt)
        {
            dt = DateTime.MinValue;
            if (v == null || v is DBNull) return false;
            if (v is DateTime) { dt = (DateTime)v; return true; }
            if (v is double || v is int || v is long)
            {
                double d = Convert.ToDouble(v, CultureInfo.InvariantCulture);
                if (d >= 1 && d < 80000)
                {
                    try { dt = DateTime.FromOADate(d); return true; }
                    catch { }
                }
                if (d >= 19000101 && d <= 29991231 && Math.Abs(d - Math.Floor(d)) < 1e-9)
                {
                    string s8 = ((long)d).ToString(CultureInfo.InvariantCulture);
                    if (DateTime.TryParseExact(s8, "yyyyMMdd", CultureInfo.InvariantCulture,
                                               DateTimeStyles.None, out dt)) return true;
                }
                return false;
            }
            string t = Convert.ToString(v, CultureInfo.InvariantCulture).Trim();
            if (t.Length == 0) return false;

            // 中文：2023年5月6日 / 2023年5月6日 13:20
            Match m = Regex.Match(t, @"(\d{4})\s*[年\-\./]\s*(\d{1,2})\s*[月\-\./]\s*(\d{1,2})\s*日?(?:\s*(\d{1,2})[:时](\d{1,2}))?");
            if (m.Success)
            {
                int y = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
                int mo = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
                int d = int.Parse(m.Groups[3].Value, CultureInfo.InvariantCulture);
                int hh = m.Groups[4].Success ? int.Parse(m.Groups[4].Value, CultureInfo.InvariantCulture) : 0;
                int mi = m.Groups[5].Success ? int.Parse(m.Groups[5].Value, CultureInfo.InvariantCulture) : 0;
                if (y >= 1900 && y <= 2999 && mo >= 1 && mo <= 12 && d >= 1 && d <= 31)
                {
                    try { dt = new DateTime(y, mo, d, hh, mi, 0); return true; }
                    catch { }
                }
            }
            // 20230506
            if (Regex.IsMatch(t, @"^\d{8}$"))
            {
                if (DateTime.TryParseExact(t, "yyyyMMdd", CultureInfo.InvariantCulture,
                                           DateTimeStyles.None, out dt)) return true;
            }
            // 2023-05-06 / 2023/5/6 / 05/06/2023
            if (DateTime.TryParse(t, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt)) return true;
            if (DateTime.TryParse(t, CultureInfo.CurrentCulture, DateTimeStyles.None, out dt)) return true;
            return false;
        }

        private static string Iso(DateTime d)
        {
            return d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        private static void WriteDates(dynamic rng, object[,] dates, string fmt)
        {
            Xl.WriteSkipFormula(rng, dates);      // 公式格不覆盖
            try { rng.NumberFormat = fmt; } catch { }
        }

        // ---------------- 1. 标准化 ----------------
        private static bool Normalize()
        {
            dynamic r = Xl.SelRange("要标准化的日期区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            object[,] outv = new object[rows, cols];
            int ok = 0;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    DateTime d;
                    if (TryParseDate(a[i, j], out d)) { outv[i, j] = d.ToOADate(); ok++; }
                    else outv[i, j] = a[i, j];
                }
            }
            if (ok == 0) { Xl.Bad("没有识别到日期。"); return true; }
            Xl.Turn(true);
            try { WriteDates(r, outv, "yyyy-mm-dd"); }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已标准化 {n} 个日期（yyyy-mm-dd）。", ok));
            return true;
        }

        // ---------------- 2/3. 格式 ----------------
        private static bool FmtDate()
        {
            dynamic r = Xl.SelRange("要设置日期格式的区域");
            if (r == null) return true;
            string[] p = Ask.Show("日期格式",
                new string[] { "输入格式代码", "常用：yyyy-mm-dd / yyyy/m/d / yyyy年m月d日 / mm-dd / m月d日 / aaaa" },
                new string[] { "yyyy-mm-dd", "" });
            if (p == null) return true;
            string fmt = p[0].Trim();
            if (fmt.Length == 0) return true;
            Xl.Turn(true);
            try { r.NumberFormatLocal = fmt; }
            catch { Xl.Bad("格式代码可能无效。"); return true; }
            finally { Xl.Turn(false); }
            Xl.Say("已应用格式：" + fmt);
            return true;
        }

        private static bool FmtTime()
        {
            dynamic r = Xl.SelRange("要设置时间格式的区域");
            if (r == null) return true;
            string[] p = Ask.Show("时间格式",
                new string[] { "输入格式代码", "常用：hh:mm / hh:mm:ss / hh时mm分 / [h]:mm:ss（累计时长）" },
                new string[] { "hh:mm:ss", "" });
            if (p == null) return true;
            string fmt = p[0].Trim();
            if (fmt.Length == 0) return true;
            Xl.Turn(true);
            try { r.NumberFormatLocal = fmt; }
            catch { Xl.Bad("格式代码可能无效。"); return true; }
            finally { Xl.Turn(false); }
            Xl.Say("已应用格式：" + fmt);
            return true;
        }

        // ---------------- 4. 拆分 ----------------
        private static bool Split()
        {
            dynamic r = Xl.SelRange("要拆分的日期列");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列日期。"); return true; }

            object[,] outv = new object[rows, 6];
            int ok = 0;
            for (int i = 0; i < rows; i++)
            {
                DateTime d;
                if (!TryParseDate(a[i, 0], out d)) continue;
                outv[i, 0] = d.Year;
                outv[i, 1] = d.Month;
                outv[i, 2] = d.Day;
                outv[i, 3] = WeekNames[(int)d.DayOfWeek];
                outv[i, 4] = "第" + ((d.Month - 1) / 3 + 1) + "季度";
                outv[i, 5] = ISOWeek(d);
                ok++;
            }
            if (ok == 0) { Xl.Bad("没有可拆分的日期。"); return true; }
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, 1, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("已拆分为 年/月/日/星期/季度/周数 六列（写右侧）。");
            return true;
        }

        private static int ISOWeek(DateTime d)
        {
            try { return ISOWeekOfYear(d); }
            catch { return Lunar.GetWeekOfYear(d, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday); }
        }

        private static int ISOWeekOfYear(DateTime d)
        {
            DayOfWeek dow = d.DayOfWeek;
            int day = (dow == DayOfWeek.Sunday) ? 7 : (int)dow;
            DateTime thursday = d.AddDays(4 - day);
            DateTime firstThursday = new DateTime(thursday.Year, 1, 1);
            int offset = ((int)firstThursday.DayOfWeek + 6) % 7;
            return 1 + (thursday.DayOfYear - 1 - offset + 7) / 7;
        }

        // ---------------- 5. 合并 ----------------
        private static bool Combine()
        {
            dynamic r = Xl.SelRange("年/月/日 三列区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols < 3) { Xl.Bad("请选择三列：年、月、日（按顺序）。"); return true; }
            object[,] outv = new object[rows, 1];
            int ok = 0;
            for (int i = 0; i < rows; i++)
            {
                int y, m, d;
                if (!int.TryParse(Xl.S(a[i, 0]).Trim(), out y)) continue;
                if (!int.TryParse(Xl.S(a[i, 1]).Trim(), out m)) continue;
                if (!int.TryParse(Xl.S(a[i, 2]).Trim(), out d)) continue;
                try { outv[i, 0] = new DateTime(y, m, d).ToOADate(); ok++; }
                catch { outv[i, 0] = ""; }
            }
            if (ok == 0) { Xl.Bad("没有可合并的行。"); return true; }
            Xl.Turn(true);
            try
            {
                TextOps.WriteBlock(r, 0, cols, outv);
                dynamic tgt = r.Worksheet.Cells[(int)r.Row, (int)r.Column + cols];
                dynamic block = r.Worksheet.Range(tgt, tgt.Offset[rows - 1, 0]);
                block.NumberFormat = "yyyy-mm-dd";
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已合并 {n} 行为标准日期（写右侧）。", ok));
            return true;
        }

        // ---------------- 6. 差值 ----------------
        private static bool Diff()
        {
            dynamic r = Xl.SelRange("起始日期与结束日期两列");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 2) { Xl.Bad("请选择两列：起始日期、结束日期。"); return true; }
            string unit = Ask.One("日期差", "单位：1=天  2=月  3=年", "1");
            if (unit == null) return true;
            int mode;
            if (!int.TryParse(unit, out mode)) mode = 1;

            object[,] outv = new object[rows, 1];
            int ok = 0;
            for (int i = 0; i < rows; i++)
            {
                DateTime d1, d2;
                if (!TryParseDate(a[i, 0], out d1) || !TryParseDate(a[i, 1], out d2)) continue;
                if (mode == 2) outv[i, 0] = MonthsBetween(d1, d2);
                else if (mode == 3) outv[i, 0] = Math.Floor(MonthsBetween(d1, d2) / 12.0);
                else outv[i, 0] = (d2.Date - d1.Date).Days;
                ok++;
            }
            if (ok == 0) { Xl.Bad("没有成对的日期。"); return true; }
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, 2, outv); }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已计算 {n} 组日期差（写右侧）。", ok));
            return true;
        }

        private static int MonthsBetween(DateTime a, DateTime b)
        {
            int months = (b.Year - a.Year) * 12 + (b.Month - a.Month);
            if (b.Day < a.Day) months--;
            return months;
        }

        // ---------------- 7. 加减 ----------------
        private static bool Add()
        {
            dynamic r = Xl.SelRange("要加减的日期区域");
            if (r == null) return true;
            string[] p = Ask.Show("日期加减",
                new string[] { "数量（可负数）", "单位：1=天 2=月 3=年 4=工作日后" },
                new string[] { "1", "1" });
            if (p == null) return true;
            int n, unit;
            if (!int.TryParse(p[0], out n) || !int.TryParse(p[1], out unit)) { Xl.Bad("请输入整数。"); return true; }

            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            object[,] outv = new object[rows, cols];
            int ok = 0;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    DateTime d;
                    if (!TryParseDate(a[i, j], out d)) { outv[i, j] = a[i, j]; continue; }
                    DateTime nd;
                    if (unit == 2) nd = d.AddMonths(n);
                    else if (unit == 3) nd = d.AddYears(n);
                    else if (unit == 4) nd = AddWorkDays(d, n, null);
                    else nd = d.AddDays(n);
                    outv[i, j] = nd.ToOADate();
                    ok++;
                }
            }
            if (ok == 0) { Xl.Bad("没有可用日期。"); return true; }
            Xl.Turn(true);
            try { WriteDates(r, outv, "yyyy-mm-dd"); }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已对 {n} 个日期完成加减。", ok));
            return true;
        }

        // ---------------- 8/9. 工作日 ----------------
        private static DateTime AddWorkDays(DateTime start, int n, HashSet<DateTime> holidays)
        {
            DateTime d = start;
            int step = n >= 0 ? 1 : -1;
            int left = Math.Abs(n);
            while (left > 0)
            {
                d = d.AddDays(step);
                if (d.DayOfWeek == DayOfWeek.Saturday || d.DayOfWeek == DayOfWeek.Sunday) continue;
                if (holidays != null && holidays.Contains(d.Date)) continue;
                left--;
            }
            return d;
        }

        private static HashSet<DateTime> PickHolidays()
        {
            HashSet<DateTime> set = new HashSet<DateTime>();
            if (!Xl.Ask("是否指定节假日区域（用于排除法定假日）？\n选「否」则只按周末计算。")) return set;
            dynamic hr = Xl.PickRange("请选择节假日日期区域（可多列）");
            if (hr == null) return set;
            int rows, cols;
            object[,] a = Xl.Vals(hr, out rows, out cols);
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    DateTime d;
                    if (TryParseDate(a[i, j], out d)) set.Add(d.Date);
                }
            return set;
        }

        private static bool WorkDays()
        {
            dynamic r = Xl.SelRange("起止日期两列");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 2) { Xl.Bad("请选择两列：开始日期、结束日期。"); return true; }
            HashSet<DateTime> holidays = PickHolidays();

            object[,] outv = new object[rows, 1];
            int ok = 0;
            for (int i = 0; i < rows; i++)
            {
                DateTime d1, d2;
                if (!TryParseDate(a[i, 0], out d1) || !TryParseDate(a[i, 1], out d2)) continue;
                if (d2 < d1) { DateTime t = d1; d1 = d2; d2 = t; }
                int count = 0;
                for (DateTime d = d1.Date; d <= d2.Date; d = d.AddDays(1))
                {
                    if (d.DayOfWeek == DayOfWeek.Saturday || d.DayOfWeek == DayOfWeek.Sunday) continue;
                    if (holidays.Contains(d)) continue;
                    count++;
                }
                outv[i, 0] = count;
                ok++;
            }
            if (ok == 0) { Xl.Bad("没有成对日期。"); return true; }
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, 2, outv); }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已计算 {n} 组工作日天数（含首尾，写右侧）。", ok));
            return true;
        }

        private static bool WorkDayAfter()
        {
            dynamic r = Xl.SelRange("起始日期区域");
            if (r == null) return true;
            int n;
            if (!Ask.Int("N 个工作日后", "往后推几个工作日（可负数）", "5", out n)) return true;
            HashSet<DateTime> holidays = PickHolidays();

            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            object[,] outv = new object[rows, cols];
            int ok = 0;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    DateTime d;
                    if (!TryParseDate(a[i, j], out d)) { outv[i, j] = a[i, j]; continue; }
                    outv[i, j] = AddWorkDays(d, n, holidays).ToOADate();
                    ok++;
                }
            }
            if (ok == 0) { Xl.Bad("没有可用日期。"); return true; }
            Xl.Turn(true);
            try { WriteDates(r, outv, "yyyy-mm-dd"); }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已推算 {n} 个日期。", ok));
            return true;
        }

        // ---------------- 10. 年龄 ----------------
        private static bool Age()
        {
            dynamic r = Xl.SelRange("出生日期列");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列出生日期。"); return true; }
            object[,] outv = new object[rows, 2];
            DateTime now = DateTime.Today;
            for (int i = 0; i < rows; i++)
            {
                DateTime d;
                if (!TryParseDate(a[i, 0], out d)) { outv[i, 0] = ""; outv[i, 1] = ""; continue; }
                int yrs = now.Year - d.Year;
                int mos = now.Month - d.Month;
                int days = now.Day - d.Day;
                if (days < 0) { mos--; days += DateTime.DaysInMonth(now.AddMonths(-1).Year, now.AddMonths(-1).Month); }
                if (mos < 0) { yrs--; mos += 12; }
                outv[i, 0] = yrs;
                outv[i, 1] = yrs + "岁" + mos + "个月";
            }
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, 1, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("已计算周岁与「X岁Y个月」（写右侧两列）。");
            return true;
        }

        // ---------------- 11. 时间戳 ----------------
        private static bool Timestamp()
        {
            dynamic r = Xl.SelRange("时间戳或日期区域");
            if (r == null) return true;
            string mode = Ask.One("时间戳与日期互转", "1 = 时间戳转日期   2 = 日期转时间戳", "1");
            if (mode == null) return true;
            bool toDate = mode.Trim() != "2";

            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            object[,] outv = new object[rows, cols];
            int ok = 0;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    object v = a[i, j];
                    if (toDate)
                    {
                        string s = Xl.S(v).Trim();
                        long ts;
                        if (!long.TryParse(s, out ts)) { outv[i, j] = v; continue; }
                        if (s.Length == 13) ts /= 1000;
                        DateTime utc = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(ts);
                        outv[i, j] = utc.ToLocalTime().ToOADate();
                        ok++;
                    }
                    else
                    {
                        DateTime d;
                        if (!TryParseDate(v, out d)) { outv[i, j] = v; continue; }
                        outv[i, j] = (long)(d.ToUniversalTime() - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds;
                        ok++;
                    }
                }
            }
            if (ok == 0) { Xl.Bad("没有可转换的值。"); return true; }
            Xl.Turn(true);
            try
            {
                Xl.WriteSkipFormula(r, outv);
                if (toDate) r.NumberFormat = "yyyy-mm-dd hh:mm:ss";
                else r.NumberFormat = "0";
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已转换 {n} 个值。", ok));
            return true;
        }

        // ---------------- 12. 农历 ----------------
        private static bool LunarOp()
        {
            dynamic r = Xl.SelRange("公历日期区域");
            if (r == null) return true;
            string mode = Ask.One("公历 ↔ 农历", "1 = 公历转农历   2 = 农历转公历（请输入 2024-01-01 形式的农历年月日）", "1");
            if (mode == null) return true;
            bool toLunar = mode.Trim() != "2";

            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            object[,] outv = new object[rows, 2];
            int ok = 0;
            for (int i = 0; i < rows; i++)
            {
                if (toLunar)
                {
                    DateTime d;
                    if (!TryParseDate(a[i, 0], out d)) continue;
                    if (d.Year < 1901 || d.Year > 2100) continue;
                    int ly = Lunar.GetYear(d);
                    int lm = Lunar.GetMonth(d);
                    int ld = Lunar.GetDayOfMonth(d);
                    int leap = Lunar.GetLeapMonth(ly);
                    bool isLeap = leap > 0 && lm == leap;
                    int realMonth = (leap > 0 && lm >= leap) ? lm - 1 : lm;
                    outv[i, 0] = Ly(ly) + "年" + (isLeap ? "闰" : "") + Lm(realMonth) + "月" + Ld(ld);
                    outv[i, 1] = Lunar.GetSexagenaryYear(d);
                    ok++;
                }
                else
                {
                    string s = Xl.S(a[i, 0]).Trim();
                    Match m = Regex.Match(s, @"(\d{4})\D+(\d{1,2})\D+(\d{1,2})");
                    if (!m.Success) continue;
                    int y = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
                    int mo = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
                    int dd = int.Parse(m.Groups[3].Value, CultureInfo.InvariantCulture);
                    try
                    {
                        DateTime g = Lunar.ToDateTime(y, mo, dd, 0, 0, 0, 0);
                        outv[i, 0] = g.ToOADate();
                        ok++;
                    }
                    catch { }
                }
            }
            if (ok == 0) { Xl.Bad("没有可转换的日期（农历支持 1901-2100）。"); return true; }
            Xl.Turn(true);
            try
            {
                TextOps.WriteBlock(r, 0, 1, outv);
                if (!toLunar)
                {
                    dynamic tgt = r.Worksheet.Cells[(int)r.Row, (int)r.Column + 1];
                    tgt.NumberFormat = "yyyy-mm-dd";
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已转换 {n} 个日期（写右侧）。", ok));
            return true;
        }

        private static string Ly(int y) { return y.ToString(CultureInfo.InvariantCulture); }
        private static readonly string[] LmNames = { "", "正", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二" };
        private static string Lm(int m) { return (m >= 1 && m <= 12) ? LmNames[m] : m.ToString(CultureInfo.InvariantCulture); }
        private static readonly string[] LdNames =
        {
            "", "初一","初二","初三","初四","初五","初六","初七","初八","初九","初十",
            "十一","十二","十三","十四","十五","十六","十七","十八","十九","二十",
            "廿一","廿二","廿三","廿四","廿五","廿六","廿七","廿八","廿九","三十"
        };
        private static string Ld(int d) { return (d >= 1 && d <= 30) ? LdNames[d] : d.ToString(CultureInfo.InvariantCulture); }

        // ---------------- 13. 区间统计 ----------------
        private static bool RangeStat()
        {
            dynamic r = Xl.SelRange("日期列（可含一列数值）");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols > 2) { Xl.Bad("最多两列：日期、数值(可选)。"); return true; }
            bool hasValue = cols == 2;
            string g = Ask.One("区间统计", "按什么分组：1=年  2=月  3=季度  4=周", "2");
            if (g == null) return true;
            int mode;
            if (!int.TryParse(g, out mode)) mode = 2;

            List<string> keys = new List<string>();
            Dictionary<string, int> counts = new Dictionary<string, int>(StringComparer.Ordinal);
            Dictionary<string, double> sums = new Dictionary<string, double>(StringComparer.Ordinal);
            for (int i = 0; i < rows; i++)
            {
                DateTime d;
                if (!TryParseDate(a[i, 0], out d)) continue;
                string k;
                if (mode == 1) k = d.Year.ToString(CultureInfo.InvariantCulture);
                else if (mode == 3) k = d.Year + "-Q" + ((d.Month - 1) / 3 + 1);
                else if (mode == 4) k = d.Year + "-W" + ISOWeek(d).ToString("00", CultureInfo.InvariantCulture);
                else k = d.ToString("yyyy-MM", CultureInfo.InvariantCulture);
                if (!counts.ContainsKey(k)) { counts[k] = 0; sums[k] = 0; keys.Add(k); }
                counts[k] = counts[k] + 1;
                if (hasValue && Xl.IsNumber(a[i, 1])) sums[k] = sums[k] + Xl.D(a[i, 1]);
            }
            if (keys.Count == 0) { Xl.Bad("没有可用日期。"); return true; }
            keys.Sort(StringComparer.Ordinal);

            int w = hasValue ? 3 : 2;
            object[,] outv = new object[keys.Count + 1, w];
            outv[0, 0] = "区间"; outv[0, 1] = "条数";
            if (hasValue) outv[0, 2] = "合计";
            for (int i = 0; i < keys.Count; i++)
            {
                outv[i + 1, 0] = keys[i];
                outv[i + 1, 1] = counts[keys[i]];
                if (hasValue) outv[i + 1, 2] = sums[keys[i]];
            }
            dynamic sh = Xl.AddSheet("区间统计");
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("已生成 " + keys.Count + " 个区间 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        // ---------------- 14. 随机日期 ----------------
        private static bool RandDate()
        {
            dynamic r = Xl.SelRange("要填入随机日期的区域");
            if (r == null) return true;
            string[] p = Ask.Show("随机日期",
                new string[] { "起始日期", "结束日期", "是否含时间(1/0)" },
                new string[] { "2024-01-01", "2024-12-31", "0" });
            if (p == null) return true;
            DateTime d1, d2;
            if (!TryParseDate(p[0], out d1) || !TryParseDate(p[1], out d2))
            {
                Xl.Bad("日期格式无法识别。");
                return true;
            }
            if (d2 < d1) { DateTime t = d1; d1 = d2; d2 = t; }
            bool withTime = p[2].Trim() != "0";
            int rows, cols;
            Xl.Vals(r, out rows, out cols);
            Random rnd = new Random(Guid.NewGuid().GetHashCode());
            object[,] outv = new object[rows, cols];
            double span = (d2 - d1).TotalSeconds;
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                    outv[i, j] = d1.AddSeconds(rnd.NextDouble() * span).ToOADate();
            Xl.Turn(true);
            try { Xl.WriteSkipFormula(r, outv); }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已生成 {n} 个随机日期。", rows * cols));
            return true;
        }

        // ---------------- 15. 清除非日期 ----------------
        private static bool ClearInvalid()
        {
            dynamic r = Xl.SelRange("要检查的日期区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            int n = 0;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (Xl.S(a[i, j]).Trim().Length == 0) continue;
                    DateTime d;
                    if (TryParseDate(a[i, j], out d)) continue;
                    a[i, j] = null;
                    n++;
                }
            }
            if (n == 0) { Xl.Say("没有非法日期。"); return true; }
            Xl.Turn(true);
            try { Xl.WriteSkipFormula(r, a); }      // 只清非法值，公式格保持不动
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已清空 {n} 个无法识别的值。", n));
            return true;
        }
    }
}

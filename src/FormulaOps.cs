using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcleHelper
{
    /// <summary>
    /// 公式工具域（批次 2）：公式↔文本、引用方式切换、公式内替换、显示全部公式、函数速查。
    /// </summary>
    internal static class FormulaOps
    {
        internal static void Register()
        {
            Ops.Reg("fm.to.text", ToText);
            Ops.Reg("fm.to.formula", ToFormula);
            Ops.Reg("fm.ref.toggle", RefToggle);
            Ops.Reg("fm.replace", ReplaceInFormula);
            Ops.Reg("fm.show.all", ShowAll);
            Ops.Reg("fm.help", Help);
        }

        // ---------------- 公式 -> 文本 ----------------
        private static bool ToText()
        {
            dynamic r = Xl.SelRange("要转成文本的公式区域");
            if (r == null) return true;
            int rows = (int)r.Rows.Count, cols = (int)r.Columns.Count;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            dynamic ws = r.Worksheet;
            int n = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        dynamic cell = ws.Cells[firstRow + i, firstCol + j];
                        string f = null;
                        try { if ((bool)cell.HasFormula) f = (string)cell.Formula; } catch { }
                        if (string.IsNullOrEmpty(f)) continue;
                        cell.NumberFormat = "@";
                        cell.Value2 = f;      // 保留等号，显示为公式文本
                        n++;
                    }
                }
            }
            finally { Xl.Turn(false); }
            if (n == 0) { Xl.Say("选区没有公式。"); return true; }
            Xl.Say(Xl.Fmt("已把 {n} 个公式转为文本（可直接复制查看）。", n));
            return true;
        }

        // ---------------- 文本 -> 公式 ----------------
        private static bool ToFormula()
        {
            dynamic r = Xl.SelRange("要转成公式的文本区域");
            if (r == null) return true;
            int rows = (int)r.Rows.Count, cols = (int)r.Columns.Count;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            dynamic ws = r.Worksheet;
            int n = 0, bad = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        dynamic cell = ws.Cells[firstRow + i, firstCol + j];
                        string s = Xl.S(cell.Value2).Trim();
                        if (s.Length < 2 || s[0] != '=') continue;
                        try
                        {
                            cell.NumberFormat = "General";
                            cell.Formula = s;
                            n++;
                        }
                        catch { bad++; }
                    }
                }
            }
            finally { Xl.Turn(false); }
            if (n == 0) { Xl.Say("没有以 = 开头的文本。"); return true; }
            Xl.Say("已还原 " + n + " 个公式" + (bad > 0 ? "（" + bad + " 条语法错误已跳过）" : "") + "。");
            return true;
        }

        // ---------------- 引用方式切换 ----------------
        private static bool RefToggle()
        {
            dynamic r = Xl.SelRange("含公式的区域");
            if (r == null) return true;
            string p = Ask.One("引用方式切换",
                "1 = 全绝对 $A$1   2 = 全相对 A1\n3 = 行绝对 A$1   4 = 列绝对 $A1", "1");
            if (p == null) return true;
            int mode;
            if (!int.TryParse(p, out mode) || mode < 1 || mode > 4) { Xl.Bad("请输入 1-4。"); return true; }

            int rows = (int)r.Rows.Count, cols = (int)r.Columns.Count;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            dynamic ws = r.Worksheet;
            int n = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        dynamic cell = ws.Cells[firstRow + i, firstCol + j];
                        string f = null;
                        try { if ((bool)cell.HasFormula) f = (string)cell.Formula; } catch { }
                        if (string.IsNullOrEmpty(f)) continue;
                        string nf = RefRewrite(f, mode);
                        if (nf == f) continue;
                        try { cell.Formula = nf; n++; }
                        catch { }
                    }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已切换 {n} 个单元格的引用方式。", n));
            return true;
        }

        /// <summary>
        /// 改写公式里的 A1 引用。只处理**双引号之外**的片段（不动字符串字面量），
        /// 并用 (?![\dA-Za-z_(]) 排除 LOG10、SUM1( 这类函数名。
        /// </summary>
        private static readonly Regex RefRx =
            new Regex(@"(?<![A-Za-z0-9_$])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![\dA-Za-z_(])");

        private static string RefRewrite(string formula, int mode)
        {
            StringBuilder sb = new StringBuilder(formula.Length + 8);
            bool inStr = false;
            int start = 0;
            for (int i = 0; i < formula.Length; i++)
            {
                if (formula[i] != '"') continue;
                if (inStr && i + 1 < formula.Length && formula[i + 1] == '"') { i++; continue; }  // "" 转义
                string seg = formula.Substring(start, i - start);
                sb.Append(inStr ? seg : RewriteSegment(seg, mode));
                sb.Append('"');
                start = i + 1;
                inStr = !inStr;
            }
            string tail = formula.Substring(start);
            sb.Append(inStr ? tail : RewriteSegment(tail, mode));
            return sb.ToString();
        }

        private static string RewriteSegment(string seg, int mode)
        {
            if (seg.IndexOf('$') < 0 && !Regex.IsMatch(seg, "[A-Za-z]")) return seg;
            return RefRx.Replace(seg, delegate(Match m)
            {
                string colPart = m.Groups[2].Value;
                string rowPart = m.Groups[4].Value;
                if (mode == 1) return "$" + colPart + "$" + rowPart;
                if (mode == 2) return colPart + rowPart;
                if (mode == 3) return colPart + "$" + rowPart;
                return "$" + colPart + rowPart;
            });
        }

        // ---------------- 公式内替换 ----------------
        private static bool ReplaceInFormula()
        {
            dynamic r = Xl.SelRange("要替换的公式区域");
            if (r == null) return true;
            string[] p = Ask.Show("公式内替换", new string[] { "查找内容", "替换为" }, new string[] { "", "" });
            if (p == null || p[0].Length == 0) return true;

            int rows = (int)r.Rows.Count, cols = (int)r.Columns.Count;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            dynamic ws = r.Worksheet;
            int n = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        dynamic cell = ws.Cells[firstRow + i, firstCol + j];
                        string f = null;
                        try { if ((bool)cell.HasFormula) f = (string)cell.Formula; } catch { }
                        if (string.IsNullOrEmpty(f)) continue;
                        if (f.IndexOf(p[0], StringComparison.OrdinalIgnoreCase) < 0) continue;
                        string nf = f.Replace(p[0], p[1]);
                        try { cell.Formula = nf; n++; }
                        catch { }
                    }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已在 {n} 个公式中替换（普通文本不受影响）。", n));
            return true;
        }

        // ---------------- 显示全部公式 ----------------
        private static bool ShowAll()
        {
            if (!Xl.Ready()) return true;
            // DisplayFormulas 是 Window 上的属性（不是 Application）
            bool on;
            try { on = !(bool)Xl.App.ActiveWindow.DisplayFormulas; }
            catch { on = true; }
            try { Xl.App.ActiveWindow.DisplayFormulas = on; }
            catch (Exception ex) { Xl.Bad("切换失败：" + ex.Message); return true; }
            Xl.Say(on ? "已进入「显示公式」模式（再执行一次可恢复）。" : "已恢复显示计算结果。");
            return true;
        }

        // ---------------- 函数速查 ----------------
        private static readonly string[,] Funcs =
        {
            {"SUM",        "SUM(数值1,[数值2]…)", "求和；忽略文本与空值"},
            {"SUMIF",      "SUMIF(条件区域,条件,[求和区域])", "单条件求和"},
            {"SUMIFS",     "SUMIFS(求和区域,条件区域1,条件1,…)", "多条件求和"},
            {"SUMPRODUCT", "SUMPRODUCT(数组1,[数组2]…)", "乘积求和；常用于多条件计数"},
            {"COUNT",      "COUNT(值1,…)", "数字个数"},
            {"COUNTA",     "COUNTA(值1,…)", "非空单元格个数"},
            {"COUNTIF",    "COUNTIF(区域,条件)", "单条件计数；条件支持 >10、\"*文*\"、\"<>\"&A1"},
            {"COUNTIFS",   "COUNTIFS(区域1,条件1,…)", "多条件计数"},
            {"AVERAGE",    "AVERAGE(数值1,…)", "平均值；忽略空与文本"},
            {"AVERAGEIF",  "AVERAGEIF(条件区域,条件,[平均值区域])", "条件平均"},
            {"MAX/MIN",    "MAX(数值1,…) / MIN(数值1,…)", "最大/最小值"},
            {"LARGE/SMALL","LARGE(数组,k)", "第 k 大/小值"},
            {"MEDIAN",     "MEDIAN(数值1,…)", "中位数"},
            {"STDEV.P/S",  "STDEV.S(数值1,…)", "样本/总体标准差"},
            {"VLOOKUP",    "VLOOKUP(查找值,表格区域,返回列号,[是否近似])", "纵向查找；末参写 FALSE/0 精确匹配"},
            {"HLOOKUP",    "HLOOKUP(查找值,区域,行号,[近似])", "横向查找"},
            {"INDEX",      "INDEX(区域,行号,[列号])", "按行列取值，常与 MATCH 组合"},
            {"MATCH",      "MATCH(查找值,区域,[匹配类型])", "返回位置；0 = 精确"},
            {"XLOOKUP",    "XLOOKUP(查找值,查找数组,返回数组,[未找到],[匹配模式])", "新版查找，替代 VLOOKUP"},
            {"FILTER",     "FILTER(数组,条件,[空值])", "按条件筛选（动态数组）"},
            {"UNIQUE",     "UNIQUE(数组)", "去重（动态数组）"},
            {"SORT",       "SORT(数组,[排序索引],[顺序])", "排序（动态数组）"},
            {"TEXTJOIN",   "TEXTJOIN(分隔符,是否忽略空,文本1,…)", "多单元格拼接，替代 &"},
            {"CONCAT",     "CONCAT(文本1,…)", "拼接（无分隔符）"},
            {"TEXT",       "TEXT(值,格式代码)", "格式化数字/日期为文本，如 TEXT(A1,\"yyyy-mm-dd\")"},
            {"LEFT/RIGHT", "LEFT(文本,字符数) / RIGHT(文本,字符数)", "截取首/尾"},
            {"MID",        "MID(文本,起始位置,字符数)", "按位置截取"},
            {"LEN",        "LEN(文本)", "字符长度"},
            {"TRIM",       "TRIM(文本)", "去掉首尾空格并压缩中间多余空格"},
            {"SUBSTITUTE", "SUBSTITUTE(文本,旧文本,新文本,[第几个])", "替换指定内容"},
            {"REPLACE",    "REPLACE(文本,起始,字符数,新文本)", "按位置替换"},
            {"FIND/SEARCH","FIND(查找文本,源文本,[起始])", "SEARCH 支持通配符且不区分大小写"},
            {"IF",         "IF(条件,真值,[假值])", "条件判断；嵌套建议改用 IFS"},
            {"IFS",        "IFS(条件1,值1,条件2,值2,…)", "多条件分支"},
            {"IFERROR",    "IFERROR(值,出错时返回)", "错误兜底"},
            {"AND/OR",     "AND(条件1,…) / OR(条件1,…)", "逻辑与/或"},
            {"ISNUMBER",   "ISNUMBER(值)", "判断是否为数字"},
            {"ROUND",      "ROUND(数值,位数)", "四舍五入；ROUNDUP/ROUNDDOWN 为上/下"},
            {"INT",        "INT(数值)", "向下取整"},
            {"MOD",        "MOD(数值,除数)", "取余"},
            {"ABS",        "ABS(数值)", "绝对值"},
            {"RANK.EQ",    "RANK.EQ(数值,引用区域,[顺序])", "排名"},
            {"PERCENTILE", "PERCENTILE.INC(数组,k)", "分位数"},
            {"DATE",       "DATE(年,月,日)", "构造日期"},
            {"TODAY/NOW",  "TODAY() / NOW()", "当天 / 当前时刻"},
            {"YEAR/MONTH/DAY", "YEAR(日期)", "取年/月/日"},
            {"WEEKDAY",    "WEEKDAY(日期,[类型])", "星期几；类型 2 = 周一为 1"},
            {"NETWORKDAYS","NETWORKDAYS(开始,结束,[节假日])", "工作日天数"},
            {"WORKDAY",    "WORKDAY(开始,天数,[节假日])", "N 个工作日后的日期"},
            {"EOMONTH",    "EOMONTH(开始,月数)", "月末日期"},
            {"DATEDIF",    "DATEDIF(开始,结束,\"Y\"/\"M\"/\"D\")", "年/月/日差（隐藏函数）"},
            {"INDIRECT",   "INDIRECT(文本引用)", "把文本变成引用，常用于二级联动"},
            {"OFFSET",     "OFFSET(基准,行偏移,列偏移,[高],[宽])", "动态区域；易拖慢计算"}
        };

        private static bool Help()
        {
            string q = Ask.One("函数速查", "输入函数名或其一部分（留空查看全部）", "");
            if (q == null) return true;
            string key = q.Trim().ToUpperInvariant();

            StringBuilder sb = new StringBuilder();
            int count = 0;
            for (int i = 0; i < Funcs.GetLength(0); i++)
            {
                string name = Funcs[i, 0];
                if (key.Length > 0 && !name.ToUpperInvariant().Contains(key)) continue;
                sb.Append(name).Append("\n    语法：").Append(Funcs[i, 1])
                  .Append("\n    说明：").Append(Funcs[i, 2]).Append("\n\n");
                count++;
            }
            if (count == 0)
            {
                Xl.Say("没有匹配的函数。可用函数：" + Funcs.GetLength(0) + " 个。");
                return true;
            }
            if (sb.Length > 1600) sb.Length = 1600;
            Xl.Say("函数速查（" + count + " 个）：\n\n" + sb.ToString());
            return true;
        }
    }
}

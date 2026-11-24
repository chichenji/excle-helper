using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>
    /// The ONLY place in the add-in that abstracts host COM differences between
    /// Microsoft Excel and WPS Spreadsheets (ET).
    /// Everything goes through IDispatch (dynamic) so no Office/WPS PIA is needed
    /// at build time or run time.
    /// </summary>
    internal static class Xl
    {
        internal static dynamic App;
        internal static bool IsWps;
        internal static string HostName = "Excel";
        internal static string LastTag = "";

        /// <summary>When true, notifications are captured instead of shown (scripted / headless runs).</summary>
        internal static bool Silent;
        internal static readonly System.Text.StringBuilder Log = new System.Text.StringBuilder();

        internal const string Caption = "EXCLE助手";

        internal static void Attach(object app)
        {
            App = app;
            IsWps = DetectWps(app);
            HostName = IsWps ? "WPS 表格" : "Excel";
        }

        internal static void Detach()
        {
            App = null;
            IsWps = false;
            HostName = "Excel";
        }

        private static bool DetectWps(object app)
        {
            try
            {
                string proc = Process.GetCurrentProcess().ProcessName ?? "";
                if (proc.Equals("et", StringComparison.OrdinalIgnoreCase)
                    || proc.Equals("wps", StringComparison.OrdinalIgnoreCase)
                    || proc.Equals("wpsoffice", StringComparison.OrdinalIgnoreCase)
                    || proc.IndexOf("kingsoft", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            catch { }
            try
            {
                if (app != null)
                {
                    dynamic d = app;
                    string name = Convert.ToString(d.Name, CultureInfo.InvariantCulture) ?? "";
                    if (name.IndexOf("WPS", StringComparison.OrdinalIgnoreCase) >= 0
                        || name.IndexOf("Kingsoft", StringComparison.OrdinalIgnoreCase) >= 0
                        || name.Equals("ET", StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            catch { }
            return false;
        }

        // ================= application state =================
        internal static bool Ready()
        {
            if (App == null)
            {
                Bad("插件尚未连接到表格程序（Excel / WPS）。请重启或在「加载项」设置中重新勾选。");
                return false;
            }
            return true;
        }

        internal static dynamic Act
        {
            get { try { return App.ActiveSheet; } catch { return null; } }
        }

        internal static dynamic Book
        {
            get { try { return App.ActiveWorkbook; } catch { return null; } }
        }

        internal static dynamic Sel
        {
            get { try { return App.Selection; } catch { return null; } }
        }

        internal static dynamic ActiveCell
        {
            get { try { return App.ActiveCell; } catch { return null; } }
        }

        /// <summary>Total cell count of a range without 32-bit overflow on whole-column/sheet selections.</summary>
        internal static long CellCount(dynamic rng)
        {
            if (rng == null) return 0;
            try { return Convert.ToInt64(rng.CountLarge, CultureInfo.InvariantCulture); }
            catch { }
            try { return Convert.ToInt64(rng.Count, CultureInfo.InvariantCulture); }
            catch { return 0; }
        }

        /// <summary>Selection if and only if it is a cell range.</summary>
        internal static dynamic SelRange(string what)
        {
            if (!Ready()) return null;
            dynamic s = Sel;
            if (s == null) { Bad("请先选择" + what + "。"); return null; }
            try
            {
                object cells = s.Cells;
                long n = CellCount(s);
                if (cells == null || n <= 0) { Bad("请先选择" + what + "。"); return null; }
            }
            catch
            {
                Bad("当前选中的不是单元格区域，请重新选择" + what + "。");
                return null;
            }
            return s;
        }

        // ================= alert suppression & safe COM wrappers =================
        private sealed class AlertScope : IDisposable
        {
            private readonly bool _hadPrev;
            private readonly bool _prev;

            internal AlertScope()
            {
                try
                {
                    _prev = (bool)App.DisplayAlerts;
                    _hadPrev = true;
                }
                catch { _hadPrev = false; _prev = true; }
                try { App.DisplayAlerts = false; } catch { }
            }

            public void Dispose()
            {
                if (_hadPrev)
                {
                    try { App.DisplayAlerts = _prev; } catch { }
                }
            }
        }

        /// <summary>Temporarily suppress Excel / WPS confirmation popups and restore previous DisplayAlerts.</summary>
        internal static IDisposable SuppressAlerts()
        {
            return new AlertScope();
        }

        internal static void DeleteSheetSilent(dynamic ws)
        {
            if (ws == null) return;
            using (SuppressAlerts())
            {
                ws.Delete();
            }
        }

        internal static void MergeSilent(dynamic rng)
        {
            if (rng == null) return;
            using (SuppressAlerts())
            {
                rng.Merge();
            }
        }

        /// <summary>
        /// Save workbook as .xlsx (FileFormat=51 xlOpenXMLWorkbook) across both Excel and WPS,
        /// suppressing overwrite prompts and handling source books that were .xls/.et/.csv.
        /// </summary>
        internal static void SaveWorkbookAsXlsx(dynamic wb, string path)
        {
            using (SuppressAlerts())
            {
                string ext = (Path.GetExtension(path) ?? "").ToLowerInvariant();
                if (ext == ".xlsx")
                {
                    try
                    {
                        wb.SaveAs(path, 51);
                        return;
                    }
                    catch { }
                }
                else if (ext == ".xlsm")
                {
                    try
                    {
                        wb.SaveAs(path, 52);
                        return;
                    }
                    catch { }
                }
                else if (ext == ".xls")
                {
                    try
                    {
                        wb.SaveAs(path, 56);
                        return;
                    }
                    catch { }
                }
                wb.SaveAs(path);
            }
        }

        /// <summary>Apply thin borders safely for 1-row / 1-column / multi-cell ranges in both Excel and WPS.</summary>
        internal static void SetThinBorders(dynamic rng, int color)
        {
            dynamic b = rng.Borders;
            int rows = 1, cols = 1;
            try { rows = (int)rng.Rows.Count; } catch { }
            try { cols = (int)rng.Columns.Count; } catch { }

            // 7=xlEdgeLeft, 8=xlEdgeTop, 9=xlEdgeBottom, 10=xlEdgeRight
            for (int i = 7; i <= 10; i++)
            {
                try
                {
                    dynamic edge = b[i];
                    edge.LineStyle = 1; // xlContinuous
                    edge.Weight = 2;    // xlThin
                    edge.Color = color;
                }
                catch { }
            }
            if (cols > 1)
            {
                try
                {
                    dynamic edge = b[11]; // xlInsideVertical
                    edge.LineStyle = 1;
                    edge.Weight = 2;
                    edge.Color = color;
                }
                catch { }
            }
            if (rows > 1)
            {
                try
                {
                    dynamic edge = b[12]; // xlInsideHorizontal
                    edge.LineStyle = 1;
                    edge.Weight = 2;
                    edge.Color = color;
                }
                catch { }
            }
        }

        /// <summary>
        /// Create a chart using ChartObjects().Add (compatible with both Excel 2007-365 and WPS ET),
        /// falling back to Shapes.AddChart2 if needed. Returns the Chart COM object.
        /// </summary>
        internal static dynamic AddChartCompat(dynamic ws, int chartType, double left, double top, double width, double height)
        {
            try
            {
                dynamic co = ws.ChartObjects().Add(left, top, width, height);
                dynamic ch = co.Chart;
                try { ch.ChartType = chartType; } catch { }
                return ch;
            }
            catch
            {
                dynamic shape = ws.Shapes.AddChart2(-1, chartType, left, top, width, height);
                return shape.Chart;
            }
        }

        // ================= messages =================
        internal static void Say(string msg)
        {
            if (Silent) { Log.AppendLine("INFO " + msg.Replace("\n", " ")); return; }
            MessageBox.Show(msg, Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        internal static void Bad(string msg)
        {
            if (Silent) { Log.AppendLine("WARN " + msg.Replace("\n", " ")); return; }
            MessageBox.Show(msg, Caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        internal static bool Ask(string msg)
        {
            return MessageBox.Show(msg, Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                   == DialogResult.Yes;
        }

        internal static string Fmt(string template, int n)
        {
            return template.Replace("{n}", n.ToString(CultureInfo.InvariantCulture));
        }

        // ================= range picking =================
        internal static dynamic PickRange(string prompt)
        {
            if (!Ready()) return null;
            object r;
            try
            {
                r = App.InputBox(prompt, Caption, Type.Missing, Type.Missing,
                                 Type.Missing, Type.Missing, Type.Missing, 8);
            }
            catch { return null; }
            if (r == null || r is bool) return null;
            return r;
        }

        // ================= busy / repaint =================
        internal static void Turn(bool busy)
        {
            try { App.ScreenUpdating = !busy; } catch { }
            try { App.EnableEvents = !busy; } catch { }
            if (!busy)
            {
                try { App.Calculate(); } catch { }
                try { App.ScreenUpdating = true; } catch { }
            }
        }

        // ================= error cell values (COM VT_ERROR) =================
        /// <summary>
        /// Detect Excel / WPS cell error values (returned by COM Value2 as negative Int32 CVErr codes,
        /// as well as literal "#..." error text).
        /// </summary>
        internal static bool IsErrorValue(object v, out string code)
        {
            code = null;
            if (v == null || v is DBNull) return false;
            if (v is int)
            {
                int iv = (int)v;
                switch (iv)
                {
                    case -2146826288: code = "#NULL!"; return true;
                    case -2146826281: code = "#DIV/0!"; return true;
                    case -2146826273: code = "#VALUE!"; return true;
                    case -2146826265: code = "#REF!"; return true;
                    case -2146826259: code = "#NAME?"; return true;
                    case -2146826252: code = "#NUM!"; return true;
                    case -2146826246: code = "#N/A"; return true;
                    case -2146826245: code = "#GETTING_DATA"; return true;
                }
                if (iv <= -2146826200 && iv >= -2146826290)
                {
                    code = "#ERROR!";
                    return true;
                }
                return false;
            }
            string s = v as string;
            if (s != null && s.Length >= 4 && s.Length <= 14 && s[0] == '#')
            {
                string u = s.ToUpperInvariant();
                if (u == "#N/A" || u == "#DIV/0!" || u == "#VALUE!" || u == "#REF!"
                    || u == "#NAME?" || u == "#NUM!" || u == "#NULL!" || u == "#SPILL!"
                    || u == "#CALC!" || u == "#FIELD!" || u == "#GETTING_DATA")
                {
                    code = s;
                    return true;
                }
            }
            return false;
        }

        // ================= values & formulas =================
        private static object[,] Normalize2D(object raw, out int rows, out int cols)
        {
            Array arr = raw as Array;
            if (arr != null && arr.Rank == 2)
            {
                int r0 = arr.GetLowerBound(0);
                int c0 = arr.GetLowerBound(1);
                rows = arr.GetLength(0);
                cols = arr.GetLength(1);
                object[,] a = new object[rows, cols];
                for (int i = 0; i < rows; i++)
                    for (int j = 0; j < cols; j++)
                        a[i, j] = arr.GetValue(r0 + i, c0 + j);
                return a;
            }

            rows = 1;
            cols = 1;
            object[,] one = new object[1, 1];
            one[0, 0] = raw;
            return one;
        }

        /// <summary>
        /// Read a range's Value2 as a guaranteed 0-based 2D array.
        /// Excel/WPS hand back SAFEARRAYs whose lower bound is 1, and marshaling may
        /// preserve that, so the raw cast is not safe to index with a[i, j].
        /// </summary>
        internal static object[,] Vals(dynamic rng, out int rows, out int cols)
        {
            return Normalize2D(rng.Value2, out rows, out cols);
        }

        /// <summary>
        /// Read a range's Formula as a guaranteed 0-based 2D array (works for both 1-cell and multi-cell ranges).
        /// </summary>
        internal static object[,] Formulas(dynamic rng, out int rows, out int cols)
        {
            return Normalize2D(rng.Formula, out rows, out cols);
        }

        internal static string S(object v)
        {
            if (v == null || v is DBNull) return "";
            string errCode;
            if (v is int && IsErrorValue(v, out errCode)) return errCode;
            // Excel stores long digit strings as doubles; Convert.ToString would then
            // hand back "4.11111111111111E+15" and destroy identifier/card numbers.
            if (v is double || v is float || v is decimal)
            {
                double d = Convert.ToDouble(v, CultureInfo.InvariantCulture);
                if (Math.Abs(d) < 1e18 && Math.Abs(d - Math.Floor(d)) < 1e-9)
                    return ((long)Math.Round(d)).ToString(CultureInfo.InvariantCulture);
                return d.ToString("R", CultureInfo.InvariantCulture);
            }
            return Convert.ToString(v, CultureInfo.InvariantCulture);
        }

        internal static bool IsNumber(object v)
        {
            if (v == null || v is DBNull) return false;
            string errCode;
            if (v is int && IsErrorValue(v, out errCode)) return false;
            if (v is double || v is int || v is long || v is decimal || v is float || v is short) return true;
            double d;
            return double.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture),
                                   NumberStyles.Any, CultureInfo.InvariantCulture, out d);
        }

        internal static double D(object v)
        {
            if (v == null || v is DBNull) return 0d;
            string errCode;
            if (v is int && IsErrorValue(v, out errCode)) return 0d;
            double d;
            if (double.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture),
                                NumberStyles.Any, CultureInfo.InvariantCulture, out d)) return d;
            try { return Convert.ToDouble(v, CultureInfo.InvariantCulture); }
            catch { return 0d; }
        }

        internal static bool IsFormula(object v)
        {
            string s = v as string;
            return s != null && s.Length > 1 && s[0] == '=';
        }

        /// <summary>
        /// 把二维数组写回选区，但**跳过公式单元格**（避免把公式静默固化成常量）。
        /// 同时兼容单单元格与多单元格选区；无公式时走整块写入的快路径；返回实际写入的单元格数。
        /// </summary>
        internal static int WriteSkipFormula(dynamic rng, object[,] outv)
        {
            int rows = outv.GetLength(0), cols = outv.GetLength(1);
            int firstRow = (int)rng.Row, firstCol = (int)rng.Column;
            dynamic ws = rng.Worksheet;

            object[,] formulas = null;
            bool anyFormula = false;
            try
            {
                int fr, fc;
                formulas = Formulas(rng, out fr, out fc);
                if (fr == rows && fc == cols)
                {
                    for (int i = 0; i < rows && !anyFormula; i++)
                        for (int j = 0; j < cols; j++)
                            if (IsFormula(formulas[i, j])) { anyFormula = true; break; }
                }
                else
                {
                    formulas = null;
                }
            }
            catch { formulas = null; }

            if (!anyFormula)
            {
                try
                {
                    if (rows == 1 && cols == 1) rng.Value2 = outv[0, 0];
                    else rng.Value2 = outv;
                    return rows * cols;
                }
                catch { return 0; }
            }

            int written = 0;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (formulas != null && IsFormula(formulas[i, j])) continue; // 公式格保持不动
                    try
                    {
                        ws.Cells[firstRow + i, firstCol + j].Value2 = outv[i, j];
                        written++;
                    }
                    catch { }
                }
            }
            return written;
        }

        /// <summary>
        /// Transform every non-formula cell of the selection. Formulas are truly detected
        /// via Range.Formula and skipped so formulas stay intact.
        /// </summary>
        internal static void MapSel(Func<string, string> f, string what)
        {
            dynamic r = SelRange(what);
            if (r == null) return;
            int rows, cols;
            object[,] a = Vals(r, out rows, out cols);
            object[,] fs = null;
            try
            {
                int fr, fc;
                fs = Formulas(r, out fr, out fc);
                if (fr != rows || fc != cols) fs = null;
            }
            catch { fs = null; }

            int changed = 0;
            Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        if (fs != null && IsFormula(fs[i, j])) continue;
                        object old = a[i, j];
                        string src = S(old);
                        if (src.Length == 0) continue;
                        string dst = f(src);
                        if (string.Equals(dst, src, StringComparison.Ordinal)) continue;
                        a[i, j] = dst;
                        changed++;
                    }
                }
                if (changed > 0) WriteSkipFormula(r, a);
            }
            finally { Turn(false); }

            if (changed == 0) Say("没有需要处理的内容。");
            else Say(Fmt("已处理 {n} 个单元格。", changed));
        }

        // ================= sheets =================
        internal static dynamic AddSheet(string name)
        {
            dynamic wb = Book;
            dynamic sh = wb.Worksheets.Add(Type.Missing, wb.Worksheets[wb.Worksheets.Count]);
            sh.Name = SafeName(wb, name);
            return sh;
        }

        internal static string SafeName(dynamic wb, string name)
        {
            if (string.IsNullOrEmpty(name)) name = "结果";
            foreach (char c in new char[] { ':', '\\', '/', '?', '*', '[', ']' })
                name = name.Replace(c, '_');
            if (name.Length > 31) name = name.Substring(0, 31);
            string baseName = name;
            int k = 2;
            while (Exists(wb, name))
            {
                string suffix = "_" + k.ToString(CultureInfo.InvariantCulture);
                int keep = 31 - suffix.Length;
                if (keep < 1) keep = 1;
                name = (baseName.Length > keep ? baseName.Substring(0, keep) : baseName) + suffix;
                k++;
            }
            return name;
        }

        private static bool Exists(dynamic wb, string name)
        {
            try { dynamic s = wb.Worksheets[name]; return s != null; }
            catch { return false; }
        }

        /// <summary>Write a 2D array to a sheet starting at (row, col), 1-based.</summary>
        internal static void Dump(dynamic sh, int row, int col, object[,] data)
        {
            int r = data.GetLength(0), c = data.GetLength(1);
            if (r <= 0 || c <= 0) return;
            dynamic start = sh.Cells[row, col];
            if (r == 1 && c == 1) { start.Value2 = data[0, 0]; return; }
            dynamic end = sh.Cells[row + r - 1, col + c - 1];
            dynamic rng = sh.Range(start, end);
            rng.Value2 = data;
        }

        /// <summary>Delete the rows of a keep-mask in batched contiguous runs (bottom-up).</summary>
        internal static void DeleteRuns(dynamic sh, int firstRow, bool[] keep)
        {
            int i = keep.Length - 1;
            while (i >= 0)
            {
                if (keep[i]) { i--; continue; }
                int end = i;
                while (i >= 0 && !keep[i]) i--;
                int start = i + 1;
                dynamic a = sh.Rows[firstRow + start];
                dynamic b = sh.Rows[firstRow + end];
                sh.Range(a, b).Delete();
            }
        }

        // ================= misc =================
        internal static int UsedRows(dynamic sh)
        {
            try { return (int)sh.UsedRange.Rows.Count; } catch { return 0; }
        }

        internal static int UsedCols(dynamic sh)
        {
            try { return (int)sh.UsedRange.Columns.Count; } catch { return 0; }
        }

        internal static string Address(dynamic rng)
        {
            try { return (string)rng.Address(false, false); } catch { return "?"; }
        }
    }
}

using System;
using System.Globalization;
using System.Text;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>
    /// The ONLY place in the add-in that talks to Excel.
    /// Everything goes through IDispatch (dynamic) so no Office PIA is needed
    /// at build time or run time.
    /// </summary>
    internal static class Xl
    {
        internal static dynamic App;
        internal static string LastTag = "";

        /// <summary>When true, notifications are captured instead of shown (scripted / headless runs).</summary>
        internal static bool Silent;
        internal static readonly System.Text.StringBuilder Log = new System.Text.StringBuilder();

        internal const string Caption = "EXCLE助手";

        // ================= application state =================
        internal static bool Ready()
        {
            if (App == null)
            {
                Bad("插件尚未连接到 Excel。请重启 Excel 或在「文件-选项-加载项」中重新勾选。");
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

        /// <summary>Selection if and only if it is a cell range.</summary>
        internal static dynamic SelRange(string what)
        {
            if (!Ready()) return null;
            dynamic s = Sel;
            if (s == null) { Bad("请先选择" + what + "。"); return null; }
            try
            {
                object cells = s.Cells;
                int n = (int)s.Count;
                if (cells == null || n <= 0) { Bad("请先选择" + what + "。"); return null; }
            }
            catch
            {
                Bad("当前选中的不是单元格区域，请重新选择" + what + "。");
                return null;
            }
            return s;
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
            if (r is bool) return null;
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

        // ================= values =================
        /// <summary>
        /// Read a range as a guaranteed 0-based 2D array.
        /// Excel hands back SAFEARRAYs whose lower bound is 1, and marshaling may
        /// preserve that, so the raw cast is not safe to index with a[i, j].
        /// </summary>
        internal static object[,] Vals(dynamic rng, out int rows, out int cols)
        {
            object v = rng.Value2;
            Array arr = v as Array;
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
            one[0, 0] = v;
            return one;
        }

        internal static string S(object v)
        {
            if (v == null || v is DBNull) return "";
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
            if (v is double || v is int || v is long || v is decimal || v is float || v is short) return true;
            double d;
            return double.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture),
                                   NumberStyles.Any, CultureInfo.InvariantCulture, out d);
        }

        internal static double D(object v)
        {
            if (v == null || v is DBNull) return 0d;
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
        /// 无公式时走整块写入的快路径；返回实际写入的单元格数。
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
                Array arr = rng.Formula as Array;
                if (arr != null)
                {
                    int lb0 = arr.GetLowerBound(0), lb1 = arr.GetLowerBound(1);
                    formulas = new object[rows, cols];
                    for (int i = 0; i < rows; i++)
                    {
                        for (int j = 0; j < cols; j++)
                        {
                            object f = arr.GetValue(lb0 + i, lb1 + j);
                            formulas[i, j] = f;
                            string fs = f as string;
                            if (fs != null && fs.Length > 0 && fs[0] == '=') anyFormula = true;
                        }
                    }
                }
            }
            catch { }

            if (!anyFormula)
            {
                try { rng.Value2 = outv; return rows * cols; }
                catch { return 0; }
            }

            int written = 0;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    string f = formulas[i, j] as string;
                    if (f != null && f.Length > 0 && f[0] == '=') continue;   // 公式格保持不动
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
        /// Transform every cell of the selection. Formulas are skipped and only
        /// changed cells are written back, so other cells keep formulas/formats.
        /// </summary>
        internal static void MapSel(Func<string, string> f, string what)
        {
            dynamic r = SelRange(what);
            if (r == null) return;
            int rows, cols;
            object[,] a = Vals(r, out rows, out cols);
            int changed = 0;
            Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        object old = a[i, j];
                        if (IsFormula(old)) continue;
                        string src = S(old);
                        if (src.Length == 0) continue;
                        string dst = f(src);
                        if (string.Equals(dst, src, StringComparison.Ordinal)) continue;
                        a[i, j] = dst;
                        changed++;
                    }
                }
                if (changed > 0) r.Value2 = a;
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
                Turn(false);
                Turn(true);
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

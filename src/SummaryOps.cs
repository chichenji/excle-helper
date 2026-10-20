using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>
    /// 汇总拆分域：合并多表/多簿、按位置汇总、拆分工作簿、按列值拆分到多文件。
    /// 汇总拆分：合并多表/多簿、按位置汇总、拆分工作表与工作簿（拆分保留原格式）。
    /// </summary>
    internal static class SummaryOps
    {
        internal static void Register()
        {
            Ops.Reg("sm.merge.two", MergeTwo);
            Ops.Reg("sm.merge.multi", MergeMulti);
            Ops.Reg("sm.merge.books", MergeBooks);
            Ops.Reg("sm.sum.sheets", SumSheets);
            Ops.Reg("sm.sum.books", SumBooks);
            Ops.Reg("sm.sum.master", SumMaster);
            Ops.Reg("sm.split.book", SplitBook);
            Ops.Reg("sm.split.values", SplitValuesToFiles);
        }

        // ---------------- dialogs ----------------
        private static string PickFolder(string desc)
        {
            using (FolderBrowserDialog d = new FolderBrowserDialog())
            {
                d.Description = desc;
                d.ShowNewFolderButton = true;
                return d.ShowDialog() == DialogResult.OK ? d.SelectedPath : null;
            }
        }

        private static bool NoBook()
        {
            if (Xl.Book == null) { Xl.Bad("请先打开一个工作簿。"); return true; }
            return false;
        }

        private static string SheetList(dynamic wb)
        {
            StringBuilder sb = new StringBuilder();
            int n = (int)wb.Worksheets.Count;
            for (int i = 1; i <= n; i++)
            {
                if (sb.Length > 0) sb.Append("、");
                sb.Append((string)wb.Worksheets[i].Name);
            }
            return sb.ToString();
        }

        private static List<string> ParseNames(string spec, dynamic wb)
        {
            List<string> list = new List<string>();
            if (string.IsNullOrEmpty(spec) || spec.Trim().Length == 0)
            {
                int n = (int)wb.Worksheets.Count;
                for (int i = 1; i <= n; i++) list.Add((string)wb.Worksheets[i].Name);
                return list;
            }
            foreach (string s in spec.Split(new char[] { ',', '，', '|', '、' }))
            {
                string t = s.Trim();
                if (t.Length > 0) list.Add(t);
            }
            return list;
        }

        // ---------------- merge ----------------
        private static bool MergeTwo()
        {
            if (NoBook()) return true;
            dynamic wb = Xl.Book;
            string[] p = Ask.Show("合并两表",
                new string[] { "表1 名称", "表2 名称", "第一行是否为标题(1=是)" },
                new string[] { "", "", "1" });
            if (p == null) return true;
            if (p[0].Trim().Length == 0 || p[1].Trim().Length == 0)
            {
                Xl.Bad("请输入两张表的名称。\n现有：" + SheetList(wb));
                return true;
            }
            bool header = p[2].Trim() != "0";

            List<string> rows = new List<string>();
            List<object[]> data = new List<object[]>();
            int width = 0;
            bool headerAdded = false;
            foreach (string name in new string[] { p[0].Trim(), p[1].Trim() })
            {
                dynamic ws;
                try { ws = wb.Worksheets[name]; }
                catch { Xl.Bad("找不到工作表：" + name); return true; }
                dynamic ur = ws.UsedRange;
                int r = (int)ur.Rows.Count, c = (int)ur.Columns.Count;
                object[,] a = Xl.Vals(ur, out r, out c);
                if (c > width) width = c;
                for (int i = 0; i < r; i++)
                {
                    if (header && !headerAdded && i == 0)
                    {
                        object[] h = new object[width];
                        for (int j = 0; j < c; j++) h[j] = a[0, j];
                        data.Add(h);
                        continue;
                    }
                    object[] row = new object[width];
                    for (int j = 0; j < c; j++) row[j] = a[i, j];
                    data.Add(row);
                }
                if (header) headerAdded = true;
            }
            if (data.Count == 0) { Xl.Bad("没有可合并的数据。"); return true; }
            object[,] outv = ToMatrix(data, width);
            dynamic sh = Xl.AddSheet("合并结果");
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("已合并 2 张表，共 " + data.Count + " 行 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        private static object[,] ToMatrix(List<object[]> data, int width)
        {
            object[,] m = new object[data.Count, width];
            for (int i = 0; i < data.Count; i++)
                for (int j = 0; j < width; j++)
                    m[i, j] = j < data[i].Length ? data[i][j] : null;
            return m;
        }

        private static bool MergeMulti()
        {
            if (NoBook()) return true;
            dynamic wb = Xl.Book;
            string spec = Ask.One("合并多表",
                "要合并的工作表(逗号分隔，留空=全部)\n现有：" + SheetList(wb), "");
            if (spec == null) return true;
            List<string> names = ParseNames(spec, wb);
            if (names.Count < 2) { Xl.Bad("至少需要两张表。"); return true; }

            List<object[]> data = new List<object[]>();
            int width = 0;
            bool headerAdded = false;
            int used = 0;
            foreach (string name in names)
            {
                dynamic ws;
                try { ws = wb.Worksheets[name]; } catch { continue; }
                dynamic ur = ws.UsedRange;
                int r = (int)ur.Rows.Count, c = (int)ur.Columns.Count;
                if (r < 1 || c < 1) continue;
                object[,] a = Xl.Vals(ur, out r, out c);
                if (c > width) width = c;
                for (int i = 0; i < r; i++)
                {
                    if (!headerAdded && i == 0)
                    {
                        object[] h = new object[width];
                        for (int j = 0; j < c; j++) h[j] = a[0, j];
                        data.Add(h);
                        continue;
                    }
                    object[] row = new object[width];
                    for (int j = 0; j < c; j++) row[j] = a[i, j];
                    data.Add(row);
                }
                headerAdded = true;
                used++;
            }
            if (data.Count == 0) { Xl.Bad("没有读到数据。"); return true; }
            dynamic sh = Xl.AddSheet("合并结果");
            Xl.Dump(sh, 1, 1, ToMatrix(data, width));
            Xl.Say("已合并 " + used + " 张表，共 " + data.Count + " 行 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        private static bool MergeBooks()
        {
            string folder = PickFolder("选择要合并的 Excel 文件所在文件夹");
            if (folder == null) return true;
            string[] files = Directory.GetFiles(folder, "*.xlsx");
            if (files.Length == 0) files = Directory.GetFiles(folder, "*.xls*");
            if (files.Length == 0) { Xl.Bad("该文件夹下没有 Excel 文件。"); return true; }

            string mode = Ask.One("合并多簿",
                "1 = 每个文件的所有工作表都合并\n2 = 只取每个文件的第一张表", "1");
            if (mode == null) return true;
            bool allSheets = mode.Trim() != "2";
            if (!Xl.Ask("将合并 " + files.Length + " 个文件，继续？")) return true;

            List<object[]> data = new List<object[]>();
            int width = 0;
            bool headerAdded = false;
            int okFiles = 0;

            Xl.Turn(true);
            try
            {
                foreach (string f in files)
                {
                    dynamic wb = null;
                    try
                    {
                        wb = Xl.App.Workbooks.Open(f, 0, true);
                        int sn = allSheets ? (int)wb.Worksheets.Count : 1;
                        for (int k = 1; k <= sn; k++)
                        {
                            dynamic ws = wb.Worksheets[k];
                            dynamic ur = ws.UsedRange;
                            int r = (int)ur.Rows.Count, c = (int)ur.Columns.Count;
                            if (r < 1 || c < 1) continue;
                            object[,] a = Xl.Vals(ur, out r, out c);
                            if (c > width) width = c;
                            for (int i = 0; i < r; i++)
                            {
                                if (!headerAdded && i == 0)
                                {
                                    object[] h = new object[width];
                                    for (int j = 0; j < c; j++) h[j] = a[0, j];
                                    data.Add(h);
                                    continue;
                                }
                                object[] row = new object[width];
                                for (int j = 0; j < c; j++) row[j] = a[i, j];
                                data.Add(row);
                            }
                            headerAdded = true;
                        }
                        okFiles++;
                    }
                    catch { }
                    finally { if (wb != null) { try { wb.Close(false); } catch { } } }
                }
            }
            finally { Xl.Turn(false); }

            if (data.Count == 0) { Xl.Bad("没有读到数据。"); return true; }
            dynamic sh = Xl.AddSheet("多簿合并");
            Xl.Dump(sh, 1, 1, ToMatrix(data, width));
            Xl.Say("已合并 " + okFiles + " 个文件，共 " + data.Count + " 行 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        // ---------------- sum by position ----------------
        private static object[,] SumMatrix(List<object[,]> blocks, out int rows, out int cols)
        {
            rows = 0; cols = 0;
            foreach (object[,] b in blocks)
            {
                if (b.GetLength(0) > rows) rows = b.GetLength(0);
                if (b.GetLength(1) > cols) cols = b.GetLength(1);
            }
            object[,] acc = new object[rows, cols];
            foreach (object[,] b in blocks)
            {
                for (int i = 0; i < b.GetLength(0); i++)
                {
                    for (int j = 0; j < b.GetLength(1); j++)
                    {
                        object v = b[i, j];
                        if (v == null || v is DBNull) continue;
                        if (Xl.IsNumber(v))
                        {
                            double cur = Xl.IsNumber(acc[i, j]) ? Xl.D(acc[i, j]) : 0d;
                            acc[i, j] = cur + Xl.D(v);
                        }
                        else if (acc[i, j] == null)
                        {
                            acc[i, j] = v;
                        }
                    }
                }
            }
            return acc;
        }

        private static bool SumSheets()
        {
            if (NoBook()) return true;
            dynamic wb = Xl.Book;
            string spec = Ask.One("汇总多表（按位置相加）",
                "要汇总的工作表(逗号分隔，留空=全部)\n现有：" + SheetList(wb), "");
            if (spec == null) return true;
            List<string> names = ParseNames(spec, wb);
            if (names.Count < 2) { Xl.Bad("至少需要两张表。"); return true; }

            List<object[,]> blocks = new List<object[,]>();
            foreach (string name in names)
            {
                dynamic ws;
                try { ws = wb.Worksheets[name]; } catch { continue; }
                int r, c;
                blocks.Add(Xl.Vals(ws.UsedRange, out r, out c));
            }
            if (blocks.Count == 0) { Xl.Bad("没有读到数据。"); return true; }
            int rows, cols;
            object[,] acc = SumMatrix(blocks, out rows, out cols);
            dynamic sh = Xl.AddSheet("汇总结果");
            Xl.Dump(sh, 1, 1, acc);
            Xl.Say("已按位置汇总 " + blocks.Count + " 张表（" + rows + " 行 × " + cols
                   + " 列）→ 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        private static bool SumBooks()
        {
            string folder = PickFolder("选择要汇总的 Excel 文件所在文件夹");
            if (folder == null) return true;
            string[] files = Directory.GetFiles(folder, "*.xlsx");
            if (files.Length == 0) files = Directory.GetFiles(folder, "*.xls*");
            if (files.Length == 0) { Xl.Bad("该文件夹下没有 Excel 文件。"); return true; }
            string sheetName = Ask.One("汇总多簿",
                "汇总每个文件的哪张表(留空=第一张)", "");
            if (sheetName == null) return true;
            if (!Xl.Ask("将汇总 " + files.Length + " 个文件，继续？")) return true;

            List<object[,]> blocks = new List<object[,]>();
            int used = 0;
            Xl.Turn(true);
            try
            {
                foreach (string f in files)
                {
                    dynamic wb = null;
                    try
                    {
                        wb = Xl.App.Workbooks.Open(f, 0, true);
                        dynamic ws = null;
                        if (sheetName.Trim().Length > 0)
                        {
                            try { ws = wb.Worksheets[sheetName.Trim()]; } catch { ws = null; }
                        }
                        if (ws == null) ws = wb.Worksheets[1];
                        int r, c;
                        blocks.Add(Xl.Vals(ws.UsedRange, out r, out c));
                        used++;
                    }
                    catch { }
                    finally { if (wb != null) { try { wb.Close(false); } catch { } } }
                }
            }
            finally { Xl.Turn(false); }

            if (blocks.Count == 0) { Xl.Bad("没有读到数据。"); return true; }
            int rows, cols;
            object[,] acc = SumMatrix(blocks, out rows, out cols);
            dynamic sh = Xl.AddSheet("多簿汇总");
            Xl.Dump(sh, 1, 1, acc);
            Xl.Say("已汇总 " + used + " 个文件（" + rows + " 行 × " + cols + " 列）→ 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        /// <summary>汇总大师：整文件夹（所有文件的可见工作表）按位置累加。</summary>
        private static bool SumMaster()
        {
            string folder = PickFolder("汇总大师：选择数据文件夹");
            if (folder == null) return true;
            if (!Xl.Ask("将遍历该文件夹下所有 Excel 文件的全部可见工作表，按位置累加（文本取首个非空值）。继续？")) return true;

            string[] files = Directory.GetFiles(folder, "*.xlsx");
            if (files.Length == 0) files = Directory.GetFiles(folder, "*.xls*");
            if (files.Length == 0) { Xl.Bad("该文件夹下没有 Excel 文件。"); return true; }

            List<object[,]> blocks = new List<object[,]>();
            int used = 0;
            Xl.Turn(true);
            try
            {
                foreach (string f in files)
                {
                    dynamic wb = null;
                    try
                    {
                        wb = Xl.App.Workbooks.Open(f, 0, true);
                        int n = (int)wb.Worksheets.Count;
                        for (int k = 1; k <= n; k++)
                        {
                            dynamic ws = wb.Worksheets[k];
                            try { if ((int)ws.Visible != -1) continue; } catch { }
                            int r, c;
                            object[,] a = Xl.Vals(ws.UsedRange, out r, out c);
                            if (r > 1 && c > 1) blocks.Add(a);
                        }
                        used++;
                    }
                    catch { }
                    finally { if (wb != null) { try { wb.Close(false); } catch { } } }
                }
            }
            finally { Xl.Turn(false); }

            if (blocks.Count == 0) { Xl.Bad("没有读到可用工作表。"); return true; }
            int rows, cols;
            object[,] acc = SumMatrix(blocks, out rows, out cols);
            dynamic sh = Xl.AddSheet("汇总大师");
            Xl.Dump(sh, 1, 1, acc);
            Xl.Say("共 " + used + " 个文件 / " + blocks.Count + " 张表 → 汇总表（" + rows + " 行 × " + cols + " 列）。");
            return true;
        }

        // ---------------- split ----------------
        private static bool SplitBook()
        {
            if (NoBook()) return true;
            dynamic wb = Xl.Book;
            string folder = PickFolder("选择拆分后的保存文件夹");
            if (folder == null) return true;
            int n = (int)wb.Worksheets.Count;
            if (!Xl.Ask("将把 " + n + " 个工作表分别另存为独立工作簿，继续？")) return true;

            int ok = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 1; i <= n; i++)
                {
                    dynamic ws = wb.Worksheets[i];
                    string name = (string)ws.Name;
                    try
                    {
                        ws.Copy();
                        dynamic nb = Xl.App.ActiveWorkbook;
                        string path = Path.Combine(folder, Safe(name) + ".xlsx");
                        nb.SaveAs(path);
                        nb.Close(false);
                        ok++;
                    }
                    catch { }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已拆分为 " + ok + " 个文件，保存到：\n" + folder);
            return true;
        }

        private static string Safe(string name)
        {
            foreach (char c in new char[] { ':', '\\', '/', '?', '*', '[', ']', '"', '<', '>', '|' })
                name = name.Replace(c, '_');
            return name.Length > 100 ? name.Substring(0, 100) : name;
        }

        internal static void DeleteMaskedRows(dynamic ws, int firstRow, bool[] keep)
        {
            Xl.DeleteRuns(ws, firstRow, keep);
        }

        private static bool SplitValuesToFiles()
        {
            if (NoBook()) return true;
            dynamic sh = Xl.Act;
            if (sh == null) { Xl.Bad("请先激活要拆分的工作表。"); return true; }
            dynamic ur = sh.UsedRange;
            int rows = (int)ur.Rows.Count;
            int cols = (int)ur.Columns.Count;
            int firstRow = (int)ur.Row;
            if (rows < 2) { Xl.Bad("数据不足。"); return true; }

            string[] p = Ask.Show("按列值拆分到多文件",
                new string[] { "依据列(第几列)", "第一行是标题(1=是)", "最多拆分数(防误操作)" },
                new string[] { "1", "1", "100" });
            if (p == null) return true;
            int keyCol, maxN;
            if (!int.TryParse(p[0], out keyCol) || !int.TryParse(p[2], out maxN)) { Xl.Bad("请输入整数。"); return true; }
            if (keyCol < 1 || keyCol > cols) { Xl.Bad("列号超出范围。"); return true; }
            bool header = p[1].Trim() != "0";

            string folder = PickFolder("选择拆分文件的保存文件夹");
            if (folder == null) return true;

            object[,] a = Xl.Vals(ur, out rows, out cols);
            int dataStart = header ? 1 : 0;
            List<string> keys = new List<string>();
            Dictionary<string, List<int>> groups = new Dictionary<string, List<int>>(StringComparer.Ordinal);
            for (int i = dataStart; i < rows; i++)
            {
                string k = Xl.S(a[i, keyCol - 1]).Trim();
                if (k.Length == 0) k = "空白";
                List<int> list;
                if (!groups.TryGetValue(k, out list)) { list = new List<int>(); groups[k] = list; keys.Add(k); }
                list.Add(i);
            }
            if (keys.Count == 0) { Xl.Bad("没有可拆分的数据。"); return true; }
            if (keys.Count > maxN) { Xl.Bad("将产生 " + keys.Count + " 个文件，超过上限 " + maxN + "。"); return true; }
            if (!Xl.Ask("将拆分为 " + keys.Count + " 个文件（保留原格式），继续？")) return true;

            dynamic wb = Xl.Book;
            int ok = 0;
            Xl.Turn(true);
            try
            {
                sh.Activate();
                foreach (string k in keys)
                {
                    List<int> list = groups[k];
                    HashSet<int> keepSet = new HashSet<int>(list);
                    dynamic copy = sh.Copy(Type.Missing, wb.Worksheets[wb.Worksheets.Count]);
                    dynamic nb = Xl.App.ActiveWorkbook;
                    dynamic wsn = nb.Worksheets[1];
                    bool[] keep = new bool[rows];
                    for (int i = 0; i < rows; i++) keep[i] = (i < dataStart) || keepSet.Contains(i);
                    if (header && dataStart == 1) keep[0] = true;
                    Xl.DeleteRuns(wsn, firstRow, keep);
                    string path = Path.Combine(folder, Safe(k) + ".xlsx");
                    nb.SaveAs(path);
                    nb.Close(false);
                    ok++;
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已生成 " + ok + " 个文件，保存到：\n" + folder);
            return true;
        }
    }
}

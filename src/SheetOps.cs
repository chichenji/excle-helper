using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>Workbook / worksheet domain: index, batch create, split, merge, export.</summary>
    internal static class SheetOps
    {
        internal static void Register()
        {
            Ops.Reg("s.index", Index);
            Ops.Reg("s.rename.batch", RenameBatch);
            Ops.Reg("s.create.batch", CreateBatch);
            Ops.Reg("s.delete", DeleteSheets);
            Ops.Reg("s.sort", SortSheets);
            Ops.Reg("s.split", SplitSheet);
            Ops.Reg("s.merge", MergeSheets);
            Ops.Reg("s.saveone", SaveOne);
            Ops.Reg("s.pdf", ToPdf);
            Ops.Reg("s.txt", ToTxt);
            Ops.Reg("w.merge", MergeWorkbooks);
            Ops.Reg("w.slim", Slim);
        }

        // ---------------- dialogs ----------------
        private static string SavePath(string filter, string defName)
        {
            using (SaveFileDialog d = new SaveFileDialog())
            {
                d.Filter = filter;
                d.FileName = defName;
                d.OverwritePrompt = true;
                return d.ShowDialog() == DialogResult.OK ? d.FileName : null;
            }
        }

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
            if (Xl.Book == null)
            {
                Xl.Bad("请先打开一个工作簿。");
                return true;
            }
            return false;
        }

        // ---------------- worksheet index ----------------
        private static bool Index()
        {
            if (NoBook()) return true;
            dynamic wb = Xl.Book;
            int n = (int)wb.Worksheets.Count;
            object[,] data = new object[n + 1, 3];
            data[0, 0] = "序号"; data[0, 1] = "工作表名称"; data[0, 2] = "可见性";
            for (int i = 1; i <= n; i++)
            {
                dynamic ws = wb.Worksheets[i];
                data[i, 0] = i;
                data[i, 1] = (string)ws.Name;
                int vis = (int)ws.Visible;
                data[i, 2] = vis == -1 ? "可见" : (vis == 0 ? "隐藏" : "深度隐藏");
            }
            dynamic sh = Xl.AddSheet("表名目录");
            Xl.Dump(sh, 1, 1, data);
            try
            {
                for (int i = 1; i <= n; i++)
                {
                    string name = Convert.ToString(data[i, 1], CultureInfo.InvariantCulture);
                    dynamic cell = sh.Cells[i + 1, 2];
                    sh.Hyperlinks.Add(cell, "", "'" + name + "'!A1", Type.Missing, name);
                }
            }
            catch { }
            Xl.Say("已生成工作表目录，共 " + n.ToString(CultureInfo.InvariantCulture) + " 个工作表（名称可点击跳转）。");
            return true;
        }

        // ---------------- batch rename ----------------
        private static bool RenameBatch()
        {
            if (NoBook()) return true;
            string t = Ask.One("批量重命名工作表",
                "命名模板：{n} 为序号，{name} 为原名，可加前缀后缀", "表_{n}");
            if (t == null || t.Length == 0) return true;
            dynamic wb = Xl.Book;
            int n = (int)wb.Worksheets.Count;
            if (!Xl.Ask("将按模板「" + t + "」重命名全部 " + n + " 个工作表，继续？")) return true;

            Xl.Turn(true);
            int ok = 0;
            try
            {
                List<string> names = new List<string>();
                for (int i = 1; i <= n; i++) names.Add((string)wb.Worksheets[i].Name);
                for (int i = 0; i < names.Count; i++)
                {
                    string nn = t.Replace("{n}", (i + 1).ToString(CultureInfo.InvariantCulture))
                                 .Replace("{name}", names[i]);
                    nn = Xl.SafeName(wb, nn);
                    try { wb.Worksheets[i + 1].Name = nn; ok++; }
                    catch { }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已重命名 " + ok.ToString(CultureInfo.InvariantCulture) + " 个工作表。");
            return true;
        }

        // ---------------- batch create ----------------
        private static bool CreateBatch()
        {
            if (NoBook()) return true;
            dynamic r = Xl.SelRange("含工作表名称的单列区域");
            if (r == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列。"); return true; }

            List<string> names = new List<string>();
            for (int i = 0; i < rows; i++)
            {
                string s = Xl.S(a[i, 0]).Trim();
                if (s.Length > 0) names.Add(s);
            }
            if (names.Count == 0) { Xl.Say("没有可用名称。"); return true; }
            if (!Xl.Ask("将创建 " + names.Count + " 个工作表，继续？")) return true;

            dynamic wb = Xl.Book;
            int ok = 0;
            Xl.Turn(true);
            try
            {
                foreach (string s in names)
                {
                    try
                    {
                        dynamic sh = wb.Worksheets.Add(Type.Missing, wb.Worksheets[wb.Worksheets.Count]);
                        sh.Name = Xl.SafeName(wb, s);
                        ok++;
                    }
                    catch { }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已创建 " + ok.ToString(CultureInfo.InvariantCulture) + " 个工作表。");
            return true;
        }

        // ---------------- delete sheets ----------------
        private static bool DeleteSheets()
        {
            if (NoBook()) return true;
            dynamic wb = Xl.Book;
            int n = (int)wb.Worksheets.Count;
            StringBuilder sb = new StringBuilder();
            for (int i = 1; i <= n; i++)
            {
                if (sb.Length > 0) sb.Append("、");
                sb.Append((string)wb.Worksheets[i].Name);
            }
            string spec = Ask.One("删除工作表", "输入要删除的工作表名（多个用 | 分隔）\n现有：" + sb, "");
            if (spec == null || spec.Trim().Length == 0) return true;
            string[] wanted = spec.Split('|');

            List<string> hit = new List<string>();
            for (int i = 1; i <= n; i++)
            {
                string name = (string)wb.Worksheets[i].Name;
                foreach (string w in wanted)
                {
                    if (string.Equals(name, w.Trim(), StringComparison.Ordinal)) { hit.Add(name); break; }
                }
            }
            if (hit.Count == 0) { Xl.Say("没有匹配的工作表名。"); return true; }
            if (hit.Count >= n) { Xl.Bad("不能删除全部工作表（工作簿至少要保留一个可见表）。"); return true; }
            if (!Xl.Ask("将删除：" + string.Join("、", hit.ToArray()) + "\n继续？")) return true;

            int ok = 0;
            Xl.Turn(true);
            try
            {
                for (int i = (int)wb.Worksheets.Count; i >= 1; i--)
                {
                    string name = (string)wb.Worksheets[i].Name;
                    if (!hit.Contains(name)) continue;
                    try { wb.Worksheets[i].Delete(); ok++; }
                    catch { }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已删除 " + ok.ToString(CultureInfo.InvariantCulture) + " 个工作表。");
            return true;
        }

        // ---------------- sort sheets ----------------
        private static bool SortSheets()
        {
            if (NoBook()) return true;
            string p = Ask.One("工作表排序", "1 = 按名称升序，2 = 按名称降序", "1");
            if (p == null) return true;
            bool asc = p.Trim() != "2";
            dynamic wb = Xl.Book;
            int n = (int)wb.Worksheets.Count;
            List<string> names = new List<string>();
            for (int i = 1; i <= n; i++) names.Add((string)wb.Worksheets[i].Name);
            names.Sort(StringComparer.CurrentCulture);
            if (!asc) names.Reverse();

            Xl.Turn(true);
            try
            {
                for (int i = 0; i < names.Count; i++)
                {
                    try { wb.Worksheets[names[i]].Move(wb.Worksheets[i + 1]); }
                    catch { }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已按名称" + (asc ? "升序" : "降序") + "排列工作表。");
            return true;
        }

        // ---------------- split ----------------
        private static bool SplitSheet()
        {
            if (NoBook()) return true;
            dynamic sh = Xl.Act;
            if (sh == null) { Xl.Bad("请先激活一个工作表。"); return true; }
            dynamic ur = sh.UsedRange;
            int rows = (int)ur.Rows.Count;
            int cols = (int)ur.Columns.Count;
            int firstRow = (int)ur.Row;
            if (rows < 2) { Xl.Bad("当前工作表数据不足。"); return true; }

            string[] p = Ask.Show("拆分工作表",
                new string[] { "按第几列拆分(从1开始)", "第一行是否为标题(1=是)", "最多拆分为多少个表(防误操作)" },
                new string[] { "1", "1", "100" });
            if (p == null) return true;
            int keyCol, hasHeader, maxSheets;
            if (!int.TryParse(p[0], out keyCol) || !int.TryParse(p[1], out hasHeader) || !int.TryParse(p[2], out maxSheets))
            {
                Xl.Bad("请输入整数。");
                return true;
            }
            if (keyCol < 1 || keyCol > cols) { Xl.Bad("列号超出数据范围（共 " + cols + " 列）。"); return true; }
            if (maxSheets < 1) maxSheets = 100;
            bool header = hasHeader == 1;

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
            if (keys.Count > maxSheets)
            {
                Xl.Bad("将产生 " + keys.Count + " 个工作表，超过上限 " + maxSheets + "。请调整列或上限。");
                return true;
            }
            if (!Xl.Ask("将按第 " + keyCol + " 列拆分为 " + keys.Count + " 个工作表，继续？")) return true;

            dynamic wb = Xl.Book;
            int done = 0;
            Xl.Turn(true);
            try
            {
                sh.Activate();
                foreach (string k in keys)
                {
                    List<int> list = groups[k];
                    HashSet<int> keepSet = new HashSet<int>(list);
                    dynamic copy = sh.Copy(Type.Missing, wb.Worksheets[wb.Worksheets.Count]);
                    dynamic ws = Xl.App.ActiveSheet;

                    bool[] keep = new bool[rows];
                    for (int i = 0; i < rows; i++)
                        keep[i] = (i < dataStart) || keepSet.Contains(i);

                    if (header && dataStart == 1) keep[0] = true;
                    Xl.DeleteRuns(ws, firstRow, keep);
                    try { ws.Name = Xl.SafeName(wb, k); } catch { }
                    done++;
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已拆分为 " + done.ToString(CultureInfo.InvariantCulture) + " 个工作表（保留原格式）。");
            return true;
        }

        // ---------------- merge sheets ----------------
        private static bool MergeSheets()
        {
            if (NoBook()) return true;
            dynamic wb = Xl.Book;
            int n = (int)wb.Worksheets.Count;
            StringBuilder sb = new StringBuilder();
            for (int i = 1; i <= n; i++)
            {
                if (sb.Length > 0) sb.Append("、");
                sb.Append((string)wb.Worksheets[i].Name);
            }
            string spec = Ask.One("合并工作表",
                "要合并的工作表名(逗号分隔，留空=全部可见表)\n现有：" + sb, "");
            if (spec == null) return true;

            List<string> targets = new List<string>();
            if (spec.Trim().Length == 0)
            {
                for (int i = 1; i <= n; i++) targets.Add((string)wb.Worksheets[i].Name);
            }
            else
            {
                foreach (string s in spec.Split(new char[] { ',', '，', '|' }))
                {
                    string t = s.Trim();
                    if (t.Length > 0) targets.Add(t);
                }
            }

            List<object[]> rowsData = new List<object[]>();
            int width = 0;
            bool headerDone = false;
            foreach (string name in targets)
            {
                dynamic ws;
                try { ws = wb.Worksheets[name]; }
                catch { continue; }
                dynamic ur = ws.UsedRange;
                int r = (int)ur.Rows.Count, c = (int)ur.Columns.Count;
                if (r < 1) continue;
                object[,] a = Xl.Vals(ur, out r, out c);
                if (c > width) width = c;
                for (int i = 0; i < r; i++)
                {
                    if (!headerDone && i == 0)
                    {
                        object[] h = new object[width];
                        for (int j = 0; j < c; j++) h[j] = a[0, j];
                        rowsData.Add(h);
                        continue;
                    }
                    object[] row = new object[width];
                    for (int j = 0; j < c; j++) row[j] = a[i, j];
                    rowsData.Add(row);
                }
                headerDone = true;
            }
            if (rowsData.Count == 0) { Xl.Bad("没有读到数据。"); return true; }

            object[,] outv = new object[rowsData.Count, width];
            for (int i = 0; i < rowsData.Count; i++)
                for (int j = 0; j < width; j++) outv[i, j] = rowsData[i][j];

            dynamic sh = Xl.AddSheet("合并结果");
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("已合并 " + targets.Count + " 个工作表，共 " + rowsData.Count + " 行 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        // ---------------- save one sheet ----------------
        private static bool SaveOne()
        {
            if (NoBook()) return true;
            dynamic sh = Xl.Act;
            string path = SavePath("Excel 工作簿|*.xlsx", ((string)sh.Name).Replace("/", "_") + ".xlsx");
            if (path == null) return true;
            Xl.Turn(true);
            try
            {
                sh.Copy();
                dynamic nb = Xl.App.ActiveWorkbook;
                nb.SaveAs(path);
                nb.Close(false);
            }
            finally { Xl.Turn(false); }
            Xl.Say("已另存为：\n" + path);
            return true;
        }

        // ---------------- export ----------------
        private static bool ToPdf()
        {
            if (NoBook()) return true;
            dynamic sh = Xl.Act;
            string path = SavePath("PDF 文件|*.pdf", ((string)sh.Name).Replace("/", "_") + ".pdf");
            if (path == null) return true;
            try
            {
                if (File.Exists(path)) File.Delete(path);
            }
            catch { }
            Xl.Turn(true);
            try { sh.ExportAsFixedFormat(0, path); }
            finally { Xl.Turn(false); }
            Xl.Say("已输出 PDF：\n" + path);
            return true;
        }

        private static bool ToTxt()
        {
            if (NoBook()) return true;
            dynamic sh = Xl.Act;
            string path = SavePath("文本文件|*.txt", ((string)sh.Name).Replace("/", "_") + ".txt");
            if (path == null) return true;
            dynamic ur = sh.UsedRange;
            int rows, cols;
            object[,] a = Xl.Vals(ur, out rows, out cols);
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (j > 0) sb.Append('\t');
                    sb.Append(Xl.S(a[i, j]));
                }
                sb.Append("\r\n");
            }
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
            Xl.Say("已输出文本：\n" + path);
            return true;
        }

        // ---------------- merge workbooks ----------------
        private static bool MergeWorkbooks()
        {
            string folder = PickFolder("选择包含要合并的 Excel 文件的文件夹");
            if (folder == null) return true;
            string[] files = Directory.GetFiles(folder, "*.xlsx");
            if (files.Length == 0) files = Directory.GetFiles(folder, "*.xls*");
            if (files.Length == 0) { Xl.Bad("该文件夹下没有 Excel 文件。"); return true; }
            if (!Xl.Ask("将合并 " + files.Length + " 个文件（每个文件取其第一张工作表），继续？")) return true;

            List<object[]> rowsData = new List<object[]>();
            int width = 0;
            bool headerDone = false;
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
                        dynamic ws = wb.Worksheets[1];
                        dynamic ur = ws.UsedRange;
                        int r = (int)ur.Rows.Count, c = (int)ur.Columns.Count;
                        if (r < 1) { wb.Close(false); continue; }
                        object[,] a = Xl.Vals(ur, out r, out c);
                        if (c > width) width = c;
                        for (int i = 0; i < r; i++)
                        {
                            if (!headerDone && i == 0)
                            {
                                object[] h = new object[width];
                                for (int j = 0; j < c; j++) h[j] = a[0, j];
                                rowsData.Add(h);
                                continue;
                            }
                            object[] row = new object[width];
                            for (int j = 0; j < c; j++) row[j] = a[i, j];
                            rowsData.Add(row);
                        }
                        headerDone = true;
                        used++;
                    }
                    catch { }
                    finally
                    {
                        if (wb != null) { try { wb.Close(false); } catch { } }
                    }
                }
            }
            finally { Xl.Turn(false); }

            if (rowsData.Count == 0) { Xl.Bad("没有读到数据。"); return true; }
            object[,] outv = new object[rowsData.Count, width];
            for (int i = 0; i < rowsData.Count; i++)
                for (int j = 0; j < width; j++) outv[i, j] = rowsData[i][j];

            dynamic sh = Xl.AddSheet("多簿合并");
            Xl.Dump(sh, 1, 1, outv);
            Xl.Say("已合并 " + used + " 个文件，共 " + rowsData.Count + " 行 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        // ---------------- slim ----------------
        private static bool Slim()
        {
            if (NoBook()) return true;
            if (!Xl.Ask("工作簿瘦身会把所有公式转为值、清除多余格式引用，且不可撤销。\n建议先保存副本。继续？")) return true;
            dynamic wb = Xl.Book;
            int n = (int)wb.Worksheets.Count;
            Xl.Turn(true);
            try
            {
                for (int i = 1; i <= n; i++)
                {
                    dynamic ws = wb.Worksheets[i];
                    dynamic ur = ws.UsedRange;
                    int r = (int)ur.Rows.Count, c = (int)ur.Columns.Count;
                    if (r < 1 || c < 1) continue;
                    object[,] a = Xl.Vals(ur, out r, out c);
                    ur.Value2 = a;
                }
                try { wb.LinkSources(1); wb.LinkSources(2); } catch { }
            }
            finally { Xl.Turn(false); }
            Xl.Say("瘦身完成（公式已转值）。请另存为新文件。");
            return true;
        }
    }
}

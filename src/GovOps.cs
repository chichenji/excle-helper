using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace ExcleHelper
{
    /// <summary>批次 A4 · 工作簿治理（wb.snapshot / wb.diff / fm.audit / fm.volatile / wb.names.clean / wb.links / wb.health）。</summary>
    internal static class GovOps
    {
        internal static void Register()
        {
            Ops.Reg("wb.snapshot", Snapshot);
            Ops.Reg("wb.diff", Diff);
            Ops.Reg("fm.audit", FormulaAudit);
            Ops.Reg("fm.volatile", Volatile);
            Ops.Reg("wb.names.clean", NamesClean);
            Ops.Reg("wb.links", Links);
            Ops.Reg("wb.health", Health);
        }

        private static string SnapDir
        {
            get
            {
                string d = Path.Combine(AiConfig.Dir, "snap");
                if (!Directory.Exists(d)) Directory.CreateDirectory(d);
                return d;
            }
        }

        private static string SnapPath
        {
            get
            {
                string name = "snapshot.xlsx";
                try { name = System.Text.RegularExpressions.Regex.Replace((string)Xl.Book.Name, @"\.\w+$", "") + ".snapshot.xlsx"; }
                catch { }
                return Path.Combine(SnapDir, name);
            }
        }

        // ================= 1. 快照 =================
        private static bool Snapshot()
        {
            if (Xl.Book == null) { Xl.Bad("请先打开工作簿。"); return true; }
            dynamic wb = Xl.Book;
            Xl.Turn(true);
            try
            {
                string path = SnapPath;
                wb.SaveCopyAs(path);
                Xl.Say("已保存当前状态快照：\n" + path + "\n\n之后可用「与快照对比」查看改了什么。");
            }
            catch (Exception ex) { Xl.Bad("保存快照失败：" + AiHttp.Sanitize(ex.Message)); }
            finally { Xl.Turn(false); }
            return true;
        }

        // ================= 2. 与快照对比 =================
        private static bool Diff()
        {
            if (Xl.Book == null) { Xl.Bad("请先打开工作簿。"); return true; }
            string path = SnapPath;
            if (!File.Exists(path)) { Xl.Bad("还没有快照，请先执行「保存当前快照」。\n预期位置：" + path); return true; }

            dynamic cur = Xl.Book;
            dynamic other = null;
            List<object[]> diffs = new List<object[]>();
            int sheetsCompared = 0;
            Xl.Turn(true);
            try
            {
                other = Xl.App.Workbooks.Open(path, false, true);
                int n = (int)cur.Worksheets.Count;
                for (int i = 1; i <= n; i++)
                {
                    dynamic s1 = cur.Worksheets[i];
                    string name = (string)s1.Name;
                    dynamic s2 = null;
                    try { s2 = other.Worksheets[name]; }
                    catch { diffs.Add(new object[] { name, "-", "（新增工作表）", "" }); continue; }
                    sheetsCompared++;
                    int r1, c1, r2, c2;
                    object[,] a = Xl.Vals(s1.UsedRange, out r1, out c1);
                    object[,] b = Xl.Vals(s2.UsedRange, out r2, out c2);
                    int rows = Math.Max(r1, r2), cols = Math.Max(c1, c2);
                    for (int x = 0; x < rows; x++)
                    {
                        for (int y = 0; y < cols; y++)
                        {
                            string va = x < r1 && y < c1 ? Xl.S(a[x, y]) : "";
                            string vb = x < r2 && y < c2 ? Xl.S(b[x, y]) : "";
                            if (va == vb) continue;
                            diffs.Add(new object[] { name, Addr(x + 1, y + 1), vb, va });
                            if (diffs.Count > 5000) break;
                        }
                        if (diffs.Count > 5000) break;
                    }
                    if (diffs.Count > 5000) break;
                }
                other.Close(false);
            }
            catch (Exception ex) { Xl.Bad("对比失败：" + AiHttp.Sanitize(ex.Message)); }
            finally
            {
                try { if (other != null) other.Close(false); } catch { }
                Xl.Turn(false);
            }

            if (diffs.Count == 0) { Xl.Say("与快照一致（对比 " + sheetsCompared + " 张工作表），没有差异。"); return true; }
            object[,] data = new object[diffs.Count + 1, 4];
            data[0, 0] = "工作表"; data[0, 1] = "位置"; data[0, 2] = "快照值"; data[0, 3] = "当前值";
            for (int i = 0; i < diffs.Count; i++)
                for (int j = 0; j < 4; j++) data[i + 1, j] = diffs[i][j];
            dynamic sh = Xl.AddSheet("快照差异");
            Xl.Dump(sh, 1, 1, data);
            try { sh.Columns[1].ColumnWidth = 16; sh.Columns[2].ColumnWidth = 10; sh.Columns[3].ColumnWidth = 30; sh.Columns[4].ColumnWidth = 30; } catch { }
            Xl.Say("发现 " + diffs.Count + " 处差异 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        private static string Addr(int row, int col)
        {
            StringBuilder sb = new StringBuilder();
            int n = col;
            while (n > 0) { int m = (n - 1) % 26; sb.Insert(0, (char)('A' + m)); n = (n - 1) / 26; }
            return sb.ToString() + row.ToString(CultureInfo.InvariantCulture);
        }

        // ================= 3. 公式依赖审计 =================
        private static bool FormulaAudit()
        {
            dynamic cell = Xl.Sel;
            if (cell == null) { Xl.Bad("请先选中一个单元格。"); return true; }
            dynamic ws = cell.Worksheet;
            StringBuilder sb = new StringBuilder();
            string addr = "";
            try { addr = (string)cell.Address(false, false); } catch { }
            sb.Append("单元格：").Append(ws.Name).Append('!').Append(addr).Append("\n\n");

            try
            {
                string f = "";
                try { if ((bool)cell.HasFormula) f = (string)cell.Formula; } catch { }
                sb.Append("公式：").Append(f.Length > 0 ? f : "（无公式）").Append("\n\n");

                List<string> pre = new List<string>();
                try
                {
                    dynamic p = cell.Precedents;
                    int n = (int)p.Areas.Count;
                    for (int i = 1; i <= n; i++) pre.Add((string)p.Areas[i].Address(false, false));
                }
                catch { }
                sb.Append("本格引用了（上游 ").Append(pre.Count).Append(" 个区域）：\n");
                if (pre.Count == 0) sb.Append("  （无）\n");
                foreach (string s in pre) sb.Append("  · ").Append(s).Append('\n');

                List<string> dep = new List<string>();
                try
                {
                    dynamic d = cell.Dependents;
                    int n = (int)d.Areas.Count;
                    for (int i = 1; i <= n; i++) dep.Add((string)d.Areas[i].Address(false, false));
                }
                catch { }
                sb.Append("\n引用了本格的位置（下游 ").Append(dep.Count).Append(" 个区域）：\n");
                if (dep.Count == 0) sb.Append("  （无）\n");
                foreach (string s in dep) sb.Append("  · ").Append(s).Append('\n');

                sb.Append("\n提示：修改本格会影响上面列出的下游单元格。");
            }
            catch (Exception ex) { sb.Append("\n读取依赖失败：").Append(AiHttp.Sanitize(ex.Message)); }
            AiPreview.Show("公式依赖审计", sb.ToString(), false);
            return true;
        }

        // ================= 4. 易失函数统计 =================
        private static bool Volatile()
        {
            dynamic ws = Xl.Act;
            if (ws == null) return true;
            string[] funcs = { "NOW(", "TODAY(", "OFFSET(", "INDIRECT(", "RAND(", "RANDBETWEEN(", "INFO(", "CELL(" };
            int[] counts = new int[funcs.Length];
            int total = 0;
            Xl.Turn(true);
            try
            {
                dynamic ur = ws.UsedRange;
                int rows = (int)ur.Rows.Count, cols = (int)ur.Columns.Count;
                if ((long)rows * cols > 300000) { Xl.Bad("区域过大，请先缩小。"); return true; }
                object[,] a = Xl.Vals(ur, out rows, out cols);
                for (int i = 0; i < rows; i++)
                    for (int j = 0; j < cols; j++)
                    {
                        string s = a[i, j] as string;
                        if (s == null || s.Length < 2 || s[0] != '=') continue;
                        string u = s.ToUpperInvariant();
                        bool hit = false;
                        for (int k = 0; k < funcs.Length; k++)
                            if (u.Contains(funcs[k])) { counts[k]++; hit = true; }
                        if (hit) total++;
                    }
            }
            finally { Xl.Turn(false); }

            StringBuilder sb = new StringBuilder();
            sb.Append("工作表：").Append((string)ws.Name).Append("\n易失函数会每次重算，数量多时明显拖慢速度。\n\n");
            for (int k = 0; k < funcs.Length; k++)
                if (counts[k] > 0) sb.Append("  ").Append(funcs[k].TrimEnd('(')).Append("：").Append(counts[k]).Append(" 处\n");
            sb.Append("\n含易失函数的单元格共 ").Append(total).Append(" 个。\n");
            if (total == 0) sb.Append("未发现易失函数，性能友好 ✓");
            else sb.Append("建议：用 INDEX 替代 OFFSET、用直接引用替代 INDIRECT，NOW/TODAY 收敛到单个单元格。");
            AiPreview.Show("易失函数统计", sb.ToString(), false);
            return true;
        }

        // ================= 5. 命名管理 =================
        private static bool NamesClean()
        {
            if (Xl.Book == null) return true;
            dynamic wb = Xl.Book;
            List<object[]> rows = new List<object[]>();
            try
            {
                dynamic names = wb.Names;
                int n = (int)names.Count;
                for (int i = 1; i <= n; i++)
                {
                    dynamic nm = names.Item(i);
                    string name = "", refers = "", visible = "";
                    try { name = (string)nm.Name; } catch { }
                    try { refers = (string)nm.RefersTo; } catch { }
                    try { visible = (bool)nm.Visible ? "是" : "否"; } catch { }
                    string status = "正常";
                    if (refers != null && refers.Contains("#REF!")) status = "引用失效";
                    else if (name != null && name.Contains("_xlnm.")) status = "系统名称";
                    rows.Add(new object[] { name, refers, visible, status });
                }
            }
            catch (Exception ex) { Xl.Bad("读取名称失败：" + AiHttp.Sanitize(ex.Message)); return true; }

            if (rows.Count == 0) { Xl.Say("本簿没有定义名称。"); return true; }
            object[,] data = new object[rows.Count + 1, 4];
            data[0, 0] = "名称"; data[0, 1] = "引用"; data[0, 2] = "可见"; data[0, 3] = "状态";
            for (int i = 0; i < rows.Count; i++)
                for (int j = 0; j < 4; j++) data[i + 1, j] = rows[i][j];
            dynamic sh = Xl.AddSheet("名称清单");
            Xl.Dump(sh, 1, 1, data);
            try { sh.Columns[1].ColumnWidth = 22; sh.Columns[2].ColumnWidth = 46; } catch { }
            int bad = 0;
            foreach (object[] r in rows) if ((string)r[3] == "引用失效") bad++;
            Xl.Say("共 " + rows.Count + " 个名称（失效 " + bad + " 个）→ 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        // ================= 6. 外部链接清单 =================
        private static bool Links()
        {
            if (Xl.Book == null) return true;
            dynamic wb = Xl.Book;
            List<object[]> rows = new List<object[]>();
            try
            {
                object src = wb.LinkSources(1);
                Array arr = src as Array;
                if (arr != null)
                {
                    for (int i = arr.GetLowerBound(0); i <= arr.GetUpperBound(0); i++)
                    {
                        string s = Convert.ToString(arr.GetValue(i), CultureInfo.InvariantCulture);
                        if (!string.IsNullOrEmpty(s)) rows.Add(new object[] { s, "工作表链接" });
                    }
                }
            }
            catch { }
            try
            {
                object src2 = wb.LinkSources(2);   // Excel links
                Array arr2 = src2 as Array;
                if (arr2 != null)
                {
                    for (int i = arr2.GetLowerBound(0); i <= arr2.GetUpperBound(0); i++)
                    {
                        string s = Convert.ToString(arr2.GetValue(i), CultureInfo.InvariantCulture);
                        if (!string.IsNullOrEmpty(s)) rows.Add(new object[] { s, "Excel 链接" });
                    }
                }
            }
            catch { }

            if (rows.Count == 0) { Xl.Say("本簿没有外部链接。"); return true; }
            object[,] data = new object[rows.Count + 1, 2];
            data[0, 0] = "链接来源"; data[0, 1] = "类型";
            for (int i = 0; i < rows.Count; i++) { data[i + 1, 0] = rows[i][0]; data[i + 1, 1] = rows[i][1]; }
            dynamic sh = Xl.AddSheet("外部链接");
            Xl.Dump(sh, 1, 1, data);
            Xl.Say("共 " + rows.Count + " 个外部链接 → 工作表「" + (string)sh.Name + "」。\n需要断开请用「隐私与安全 → 断开外部链接」。");
            return true;
        }

        // ================= 7. 工作簿体检 =================
        private static bool Health()
        {
            if (Xl.Book == null) return true;
            dynamic wb = Xl.Book;
            int sheets = 0, visible = 0, merged = 0, formulas = 0, cells = 0, names = 0, links = 0;
            long bytes = 0;
            string path = "";
            Xl.Turn(true);
            try
            {
                sheets = (int)wb.Worksheets.Count;
                for (int i = 1; i <= sheets; i++)
                {
                    dynamic ws = wb.Worksheets[i];
                    try { if ((int)ws.Visible == -1) visible++; } catch { }
                    try { if ((bool)ws.UsedRange.MergeCells) merged++; } catch { }
                    try
                    {
                        int r = (int)ws.UsedRange.Rows.Count, c = (int)ws.UsedRange.Columns.Count;
                        cells += r * c;
                        if ((long)r * c <= 200000)
                        {
                            dynamic f = ws.UsedRange.SpecialCells(-4123);   // formulas
                            if (f != null) formulas += (int)f.Count;
                        }
                    }
                    catch { }
                }
                try { names = (int)wb.Names.Count; } catch { }
                try
                {
                    object src = wb.LinkSources(1);
                    Array arr = src as Array;
                    if (arr != null) links = arr.Length;
                }
                catch { }
                try
                {
                    path = (string)wb.FullName;
                    if (File.Exists(path)) bytes = new FileInfo(path).Length;
                }
                catch { }
            }
            finally { Xl.Turn(false); }

            object[,] data = new object[9, 2];
            data[0, 0] = "工作表数"; data[0, 1] = sheets;
            data[1, 0] = "可见工作表"; data[1, 1] = visible;
            data[2, 0] = "已用单元格数"; data[2, 1] = cells;
            data[3, 0] = "公式数（抽样）"; data[3, 1] = formulas;
            data[4, 0] = "定义名称"; data[4, 1] = names;
            data[5, 0] = "外部链接"; data[5, 1] = links;
            data[6, 0] = "含合并单元格的表"; data[6, 1] = merged;
            data[7, 0] = "文件大小(KB)"; data[7, 1] = bytes / 1024;
            data[8, 0] = "文件路径"; data[8, 1] = path.Length > 0 ? path : "（未保存）";

            dynamic sh = Xl.AddSheet("工作簿体检");
            Xl.Dump(sh, 1, 1, data);
            try { sh.Columns[1].ColumnWidth = 20; sh.Columns[2].ColumnWidth = 50; } catch { }
            Xl.Say("体检完成 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>
    /// 工作表域：
    /// 建表、复制、移动、保护、表名、目录索引、文件与文件夹维护、导出 CSV。
    /// </summary>
    internal static class WsOps
    {
        internal static void Register()
        {
            Ops.Reg("w.new", NewSheet);
            Ops.Reg("w.copy", CopySheet);
            Ops.Reg("w.move", MoveSheet);
            Ops.Reg("w.protect", ProtectSheet);
            Ops.Reg("w.names.extract", ExtractNames);
            Ops.Reg("w.smart.index", SmartIndex);
            Ops.Reg("w.rename.one", RenameOne);
            Ops.Reg("w.hide", HideSheets);
            Ops.Reg("w.unhide", UnhideSheets);
            Ops.Reg("w.book.rename", RenameBook);
            Ops.Reg("w.file.rename", RenameFiles);
            Ops.Reg("w.file.list", FileList);
            Ops.Reg("w.folder.create", CreateFolder);
            Ops.Reg("w.folder.empty", EmptyFolders);
            Ops.Reg("w.export.csv", ExportCsv);
        }

        private static bool NoBook()
        {
            if (Xl.Book == null) { Xl.Bad("请先打开一个工作簿。"); return true; }
            return false;
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

        private static string SheetNamesString(dynamic wb)
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

        // ---------------- 建表 / 复制 / 移动 ----------------
        private static bool NewSheet()
        {
            if (NoBook()) return true;
            string name = Ask.One("创建工作表", "工作表名称（留空自动命名）", "");
            if (name == null) return true;
            dynamic wb = Xl.Book;
            int before = (int)wb.Worksheets.Count;
            dynamic sh = wb.Worksheets.Add(Type.Missing, wb.Worksheets[before]);
            if (name.Trim().Length > 0) sh.Name = Xl.SafeName(wb, name.Trim());
            sh.Activate();
            Xl.Say("已创建工作表「" + (string)sh.Name + "」。");
            return true;
        }

        private static bool CopySheet()
        {
            if (NoBook()) return true;
            dynamic wb = Xl.Book;
            dynamic cur = Xl.Act;
            string[] p = Ask.Show("复制工作表",
                new string[] { "当前表：" + (string)cur.Name + "，复制几份", "新表名后缀" },
                new string[] { "1", "_副本" });
            if (p == null) return true;
            int n;
            if (!int.TryParse(p[0], out n) || n < 1 || n > 50) { Xl.Bad("份数需在 1-50。"); return true; }
            string suffix = p[1];
            string baseName = (string)cur.Name;
            Xl.Turn(true);
            try
            {
                for (int k = 1; k <= n; k++)
                {
                    dynamic copy = cur.Copy(Type.Missing, wb.Worksheets[wb.Worksheets.Count]);
                    dynamic made = Xl.App.ActiveSheet;
                    try { made.Name = Xl.SafeName(wb, baseName + suffix + (n > 1 ? k.ToString(CultureInfo.InvariantCulture) : "")); }
                    catch { }
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已复制 {n} 份工作表。", n));
            return true;
        }

        private static bool MoveSheet()
        {
            if (NoBook()) return true;
            dynamic wb = Xl.Book;
            dynamic cur = Xl.Act;
            string p = Ask.One("移动工作表",
                "移动到第几位（从 1 开始）\n现有：" + SheetNamesString(wb), "1");
            if (p == null) return true;
            int idx;
            if (!int.TryParse(p, out idx)) { Xl.Bad("请输入整数。"); return true; }
            int n = (int)wb.Worksheets.Count;
            if (idx < 1) idx = 1;
            if (idx > n) idx = n;
            Xl.Turn(true);
            try
            {
                if (idx == 1) cur.Move(wb.Worksheets[2]);
                else cur.Move(Type.Missing, wb.Worksheets[idx]);
            }
            finally { Xl.Turn(false); }
            Xl.Say("已把「" + (string)cur.Name + "」移动到第 " + idx + " 位。");
            return true;
        }

        private static bool ProtectSheet()
        {
            if (NoBook()) return true;
            dynamic ws = Xl.Act;
            bool protecting;
            try { protecting = !(bool)ws.ProtectContents; }
            catch { protecting = true; }

            if (protecting)
            {
                string pwd = Ask.One("保护工作表", "密码（留空则无密码）", "");
                if (pwd == null) return true;
                if (pwd.Length == 0) ws.Protect(Type.Missing, true, true, true);
                else ws.Protect(pwd, true, true, true);
                Xl.Say("已保护工作表「" + (string)ws.Name + "」。");
            }
            else
            {
                string pwd = Ask.One("取消保护", "密码（若无密码直接确定）", "");
                if (pwd == null) return true;
                try
                {
                    if (pwd.Length == 0) ws.Unprotect();
                    else ws.Unprotect(pwd);
                    Xl.Say("已取消保护。");
                }
                catch { Xl.Bad("取消保护失败：密码可能不正确。"); }
            }
            return true;
        }

        // ---------------- 表名 / 索引 ----------------
        private static bool ExtractNames()
        {
            if (NoBook()) return true;
            dynamic r = Xl.SelRange("写入表名的起始单元格");
            if (r == null) return true;
            dynamic wb = Xl.Book;
            int n = (int)wb.Worksheets.Count;
            object[,] outv = new object[n, 1];
            for (int i = 1; i <= n; i++) outv[i - 1, 0] = (string)wb.Worksheets[i].Name;
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, 0, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("已把 " + n + " 个工作表名写入 " + Xl.Address(r) + " 起。");
            return true;
        }

        private static bool SmartIndex()
        {
            if (NoBook()) return true;
            dynamic wb = Xl.Book;
            int n = (int)wb.Worksheets.Count;
            object[,] data = new object[n + 1, 4];
            data[0, 0] = "序号"; data[0, 1] = "工作表；点击跳转"; data[0, 2] = "已用行数"; data[0, 3] = "状态";
            for (int i = 1; i <= n; i++)
            {
                dynamic ws = wb.Worksheets[i];
                data[i, 0] = i;
                data[i, 1] = (string)ws.Name;
                int rows = 0;
                try { rows = (int)ws.UsedRange.Rows.Count; } catch { }
                data[i, 2] = rows;
                int vis = -1;
                try { vis = (int)ws.Visible; } catch { }
                data[i, 3] = vis == -1 ? "可见" : (vis == 0 ? "隐藏" : "深度隐藏");
            }
            dynamic sh = Xl.AddSheet("index");
            try { sh.Name = Xl.SafeName(wb, "目录"); } catch { }
            Xl.Dump(sh, 1, 1, data);
            try
            {
                for (int i = 1; i <= n; i++)
                {
                    string nm = Convert.ToString(data[i, 1], CultureInfo.InvariantCulture);
                    dynamic cell = sh.Cells[i + 1, 2];
                    sh.Hyperlinks.Add(cell, "", "'" + nm + "'!A1", Type.Missing, nm);
                }
                sh.Columns[2].ColumnWidth = 26;
                sh.Columns[1].ColumnWidth = 6;
                dynamic head = sh.Range(sh.Cells[1, 1], sh.Cells[1, 4]);
                head.Font.Bold = true;
                head.Interior.Color = 0xEB6325;   // 品牌蓝
                head.Font.Color = 0xFFFFFF;
            }
            catch { }
            sh.Activate();
            Xl.Say("已生成智能索引（" + n + " 张表，点击表名可跳转）。");
            return true;
        }

        private static bool RenameOne()
        {
            if (NoBook()) return true;
            dynamic ws = Xl.Act;
            string name = Ask.One("重命名工作表", "新名称", (string)ws.Name);
            if (name == null || name.Trim().Length == 0) return true;
            try { ws.Name = Xl.SafeName(Xl.Book, name.Trim()); }
            catch (Exception ex) { Xl.Bad("重命名失败：" + ex.Message); return true; }
            Xl.Say("已重命名为「" + (string)ws.Name + "」。");
            return true;
        }

        private static bool HideSheets()
        {
            if (NoBook()) return true;
            dynamic wb = Xl.Book;
            string spec = Ask.One("隐藏工作表", "要隐藏的表(逗号分隔，留空=当前表)\n现有：" + SheetNamesString(wb), "");
            if (spec == null) return true;
            List<string> names = new List<string>();
            if (spec.Trim().Length == 0) names.Add((string)Xl.Act.Name);
            else foreach (string s in spec.Split(new char[] { ',', '，', '|' }))
                    if (s.Trim().Length > 0) names.Add(s.Trim());

            int n = (int)wb.Worksheets.Count;
            int visible = 0;
            for (int i = 1; i <= n; i++) { try { if ((int)wb.Worksheets[i].Visible == -1) visible++; } catch { } }
            if (visible - names.Count < 1) { Xl.Bad("不能隐藏全部工作表。"); return true; }

            int ok = 0;
            foreach (string nm in names)
            {
                try { wb.Worksheets[nm].Visible = 0; ok++; } catch { }
            }
            Xl.Say(Xl.Fmt("已隐藏 {n} 张表。", ok));
            return true;
        }

        private static bool UnhideSheets()
        {
            if (NoBook()) return true;
            dynamic wb = Xl.Book;
            int n = (int)wb.Worksheets.Count;
            int ok = 0;
            for (int i = 1; i <= n; i++)
            {
                try
                {
                    if ((int)wb.Worksheets[i].Visible != -1) { wb.Worksheets[i].Visible = -1; ok++; }
                }
                catch { }
            }
            Xl.Say(ok == 0 ? "没有被隐藏的工作表。" : Xl.Fmt("已取消隐藏 {n} 张表。", ok));
            return true;
        }

        // ---------------- 工作簿 / 文件 ----------------
        private static bool RenameBook()
        {
            if (NoBook()) return true;
            dynamic wb = Xl.Book;
            string cur = "";
            try { cur = Path.GetFileNameWithoutExtension((string)wb.FullName); } catch { }
            string name = Ask.One("重命名本簿", "新的文件名（不含扩展名）", cur);
            if (name == null || name.Trim().Length == 0) return true;
            try
            {
                string folder = Path.GetDirectoryName((string)wb.FullName);
                string ext = Path.GetExtension((string)wb.FullName);
                if (ext.Length == 0) ext = ".xlsx";
                string target = Path.Combine(folder, name.Trim() + ext);
                if (File.Exists(target)) { Xl.Bad("目标文件已存在：" + target); return true; }
                wb.SaveAs(target);
            }
            catch (Exception ex) { Xl.Bad("重命名失败：" + ex.Message); return true; }
            Xl.Say("已保存为新文件名。");
            return true;
        }

        private static bool RenameFiles()
        {
            string folder = PickFolder("选择要重命名的文件所在文件夹");
            if (folder == null) return true;
            string[] files = Directory.GetFiles(folder, "*.xls*");
            if (files.Length == 0) { Xl.Bad("该文件夹下没有 Excel 文件。"); return true; }
            string[] p = Ask.Show("规则重命名文件",
                new string[] { "新前缀（可留空）", "序号起始", "保留原名(1=是)" },
                new string[] { "", "1", "1" });
            if (p == null) return true;
            int start;
            if (!int.TryParse(p[1], out start)) { Xl.Bad("起始序号必须是整数。"); return true; }
            bool keep = p[2].Trim() != "0";
            if (!Xl.Ask("将重命名 " + files.Length + " 个文件，继续？")) return true;

            int ok = 0;
            Array.Sort(files);
            for (int i = 0; i < files.Length; i++)
            {
                try
                {
                    string dir = Path.GetDirectoryName(files[i]);
                    string ext = Path.GetExtension(files[i]);
                    string oldName = Path.GetFileNameWithoutExtension(files[i]);
                    string nn = p[0] + (keep ? oldName : "") + (start + i).ToString(CultureInfo.InvariantCulture);
                    string target = Path.Combine(dir, nn + ext);
                    if (string.Equals(target, files[i], StringComparison.OrdinalIgnoreCase)) continue;
                    if (File.Exists(target)) continue;
                    File.Move(files[i], target);
                    ok++;
                }
                catch { }
            }
            Xl.Say("已重命名 " + ok + " 个文件。");
            return true;
        }

        private static bool FileList()
        {
            string folder = PickFolder("选择要列出文件的文件夹");
            if (folder == null) return true;
            string[] files = Directory.GetFiles(folder);
            if (files.Length == 0) { Xl.Bad("该文件夹为空。"); return true; }
            object[,] data = new object[files.Length + 1, 4];
            data[0, 0] = "文件名"; data[0, 1] = "大小(KB)"; data[0, 2] = "修改时间"; data[0, 3] = "完整路径";
            Array.Sort(files, StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < files.Length; i++)
            {
                FileInfo fi = new FileInfo(files[i]);
                data[i + 1, 0] = fi.Name;
                data[i + 1, 1] = Math.Round(fi.Length / 1024.0, 1);
                data[i + 1, 2] = fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
                data[i + 1, 3] = fi.FullName;
            }
            dynamic sh = Xl.AddSheet("文件目录");
            Xl.Dump(sh, 1, 1, data);
            sh.Columns[1].ColumnWidth = 40;
            sh.Columns[4].ColumnWidth = 60;
            Xl.Say("已列出 " + files.Length + " 个文件 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        private static bool CreateFolder()
        {
            string folder = PickFolder("选择在哪一层新建文件夹");
            if (folder == null) return true;
            string name = Ask.One("创建文件夹", "文件夹名称（可写多级，如 2026\\一月）", "");
            if (name == null || name.Trim().Length == 0) return true;
            try
            {
                string target = Path.Combine(folder, name.Trim());
                Directory.CreateDirectory(target);
                Xl.Say("已创建：" + target);
            }
            catch (Exception ex) { Xl.Bad("创建失败：" + ex.Message); }
            return true;
        }

        private static bool EmptyFolders()
        {
            string folder = PickFolder("选择要检查空文件夹的目录");
            if (folder == null) return true;
            List<string> empties = new List<string>();
            try
            {
                foreach (string d in Directory.GetDirectories(folder, "*", SearchOption.AllDirectories))
                {
                    try
                    {
                        if (Directory.GetFiles(d).Length == 0 && Directory.GetDirectories(d).Length == 0)
                            empties.Add(d);
                    }
                    catch { }
                }
            }
            catch (Exception ex) { Xl.Bad("扫描失败：" + ex.Message); return true; }

            if (empties.Count == 0) { Xl.Say("没有空文件夹。"); return true; }
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < empties.Count && i < 15; i++) sb.AppendLine(empties[i]);
            if (empties.Count > 15) sb.AppendLine("… 共 " + empties.Count + " 个");
            if (!Xl.Ask("发现 " + empties.Count + " 个空文件夹：\n" + sb + "\n将把它们移到回收站（不是永久删除），继续？"))
                return true;

            int ok = 0;
            foreach (string d in empties)
            {
                try
                {
                    Microsoft.VisualBasic.FileIO.FileSystem.DeleteDirectory(
                        d,
                        Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
                        Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);
                    ok++;
                }
                catch { }
            }
            Xl.Say("已把 " + ok + " 个空文件夹移到回收站。");
            return true;
        }

        private static bool ExportCsv()
        {
            if (NoBook()) return true;
            dynamic ws = Xl.Act;
            string path;
            using (SaveFileDialog d = new SaveFileDialog())
            {
                d.Filter = "CSV 文件|*.csv";
                d.FileName = ((string)ws.Name).Replace("/", "_") + ".csv";
                if (d.ShowDialog() != DialogResult.OK) return true;
                path = d.FileName;
            }
            int rows, cols;
            object[,] a = Xl.Vals(ws.UsedRange, out rows, out cols);
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (j > 0) sb.Append(',');
                    string s = Xl.S(a[i, j]);
                    if (s.IndexOf(',') >= 0 || s.IndexOf('"') >= 0 || s.IndexOf('\n') >= 0)
                        sb.Append('"').Append(s.Replace("\"", "\"\"")).Append('"');
                    else sb.Append(s);
                }
                sb.Append("\r\n");
            }
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
            Xl.Say("已输出 CSV：\n" + path);
            return true;
        }
    }
}

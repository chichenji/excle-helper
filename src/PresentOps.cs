using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>
    /// 批次 A3 · 呈现与效率（cf.manage / fmt.* / ch.set / pic.barcode / pic.export.range / ui.recent / ui.favorites）。
    /// 条码用 Code39（实现简单且可靠，支持 0-9 A-Z - . 空格 $ / + %）。
    /// </summary>
    internal static class PresentOps
    {
        internal static void Register()
        {
            Ops.Reg("cf.manage", CfManage);
            Ops.Reg("fmt.sparkline", Sparkline);
            Ops.Reg("fmt.style", Style);
            Ops.Reg("ch.set", ChartSet);
            Ops.Reg("pic.barcode", Barcode);
            Ops.Reg("pic.export.range", ExportRange);
            Ops.Reg("ui.recent", Recent);
            Ops.Reg("ui.favorites", Favorites);
        }

        // ================= 输入钩子（自动化测试用） =================
        private static string PfOpt(string key, string title, string tip, string def)
        {
            try
            {
                string path = Path.Combine(Path.GetTempPath(), "excle-pf-inputs.txt");
                if (File.Exists(path))
                {
                    foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
                    {
                        int eq = line.IndexOf('=');
                        if (eq <= 0) continue;
                        if (!string.Equals(line.Substring(0, eq).Trim(), key, StringComparison.OrdinalIgnoreCase)) continue;
                        return line.Substring(eq + 1).Trim();
                    }
                }
            }
            catch { }
            string s = Ask.One(title, tip, def);
            return s == null ? null : s.Trim();
        }

        // ================= 1. 条件格式管理器 =================
        private static bool CfManage()
        {
            dynamic ws = Xl.Act;
            if (ws == null) { Xl.Bad("请先打开工作簿。"); return true; }
            List<object[]> rows = new List<object[]>();
            try
            {
                dynamic ur = ws.UsedRange;
                int r = (int)ur.Rows.Count, c = (int)ur.Columns.Count;
                if ((long)r * c > 200000) { Xl.Bad("区域过大，请先缩小数据范围。"); return true; }
                for (int i = 1; i <= r; i++)
                {
                    for (int j = 1; j <= c; j++)
                    {
                        dynamic cell = ws.Cells[i, j];
                        int n = 0;
                        try { n = (int)cell.FormatConditions.Count; } catch { }
                        if (n == 0) continue;
                        for (int k = 1; k <= n; k++)
                        {
                            dynamic fc = cell.FormatConditions[k];
                            string type = "";
                            try { type = ((int)fc.Type).ToString(CultureInfo.InvariantCulture); } catch { }
                            string f1 = "";
                            try { f1 = (string)fc.Formula1; } catch { }
                            rows.Add(new object[] { (string)cell.Address(false, false), TypeName(type), f1 });
                        }
                    }
                }
            }
            catch (Exception ex) { Xl.Bad("读取条件格式失败：" + AiHttp.Sanitize(ex.Message)); return true; }

            if (rows.Count == 0) { Xl.Say("当前工作表没有条件格式规则。"); return true; }
            string mode = PfOpt("cfmode", "条件格式管理器",
                "共 " + rows.Count + " 条规则。\n1 = 只列清单（到新表）\n2 = 全部删除",
                "1");
            if (mode == null) return true;

            if (mode.Trim() == "2")
            {
                if (!Xl.Ask("确认删除本表全部条件格式？")) return true;
                int del = 0;
                try
                {
                    dynamic ur = ws.UsedRange;
                    del = (int)ur.FormatConditions.Count;
                    ur.FormatConditions.Delete();
                }
                catch { }
                Xl.Say("已删除本表条件格式（" + del + " 条规则集合）。可用「撤销上次操作」尝试回退格式。");
                return true;
            }

            object[,] data = new object[rows.Count + 1, 3];
            data[0, 0] = "单元格"; data[0, 1] = "规则类型"; data[0, 2] = "公式/条件";
            for (int i = 0; i < rows.Count; i++)
                for (int j = 0; j < 3; j++) data[i + 1, j] = rows[i][j];
            dynamic sh = Xl.AddSheet("条件格式清单");
            Xl.Dump(sh, 1, 1, data);
            try { sh.Columns[1].ColumnWidth = 14; sh.Columns[2].ColumnWidth = 16; sh.Columns[3].ColumnWidth = 40; } catch { }
            Xl.Say("共 " + rows.Count + " 条规则 → 工作表「" + (string)sh.Name + "」。");
            return true;
        }

        private static string TypeName(string code)
        {
            switch (code)
            {
                case "1": return "单元格值";
                case "2": return "表达式";
                case "3": return "色阶";
                case "4": return "数据条";
                case "5": return "图标集";
                case "6": return "前10项";
                default: return "类型" + code;
            }
        }

        // ================= 2. 迷你图 =================
        private static bool Sparkline()
        {
            dynamic r = Xl.SelRange("每行是一组数据的区域（迷你图会插到每行右侧）");
            if (r == null) return true;
            int kind = 1;
            string k = Ask.One("批量插入迷你图", "类型：1=折线 2=柱形 3=盈亏", "1");
            if (k == null) return true;
            int.TryParse(k.Trim(), out kind);
            if (kind < 1 || kind > 3) kind = 1;

            int rows = (int)r.Rows.Count, cols = (int)r.Columns.Count;
            dynamic ws = r.Worksheet;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            int made = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    dynamic src = ws.Range[ws.Cells[firstRow + i, firstCol],
                                           ws.Cells[firstRow + i, firstCol + cols - 1]];
                    dynamic target = ws.Cells[firstRow + i, firstCol + cols];
                    try
                    {
                        target.SparklineGroups.Add(kind, src);
                        made++;
                    }
                    catch { }
                }
            }
            catch (Exception ex) { Xl.Bad("插入迷你图失败（需要 Excel 2010+）：" + AiHttp.Sanitize(ex.Message)); }
            finally { Xl.Turn(false); }
            Xl.Say("已插入 " + made + " 个迷你图（写在每行右侧一列）。");
            return true;
        }

        // ================= 3. 单元格样式库 =================
        private static bool Style()
        {
            dynamic r = Xl.SelRange("要套用样式的区域");
            if (r == null) return true;
            string pick = PfOpt("style", "单元格样式库",
                "1=标题（深蓝底白字） 2=表头（灰底加粗） 3=金额（千分位） 4=百分比 5=日期\n" +
                "6=警示（浅红） 7=成功（浅绿） 8=斑马纹（隔行浅灰）", "2");
            if (pick == null) return true;
            int s;
            if (!int.TryParse(pick.Trim(), out s)) s = 2;

            Xl.Turn(true);
            try
            {
                switch (s)
                {
                    case 1:
                        r.Font.Bold = true; r.Font.Size = 12; r.Font.Color = 0xFFFFFF;
                        r.Interior.Color = 0x6D2D1D; r.HorizontalAlignment = -4108; break;
                    case 2:
                        r.Font.Bold = true; r.Interior.Color = 0xE3E7E1; r.HorizontalAlignment = -4108; break;
                    case 3:
                        r.NumberFormat = "#,##0.00"; r.HorizontalAlignment = -4152; break;
                    case 4:
                        r.NumberFormat = "0.00%"; r.HorizontalAlignment = -4152; break;
                    case 5:
                        r.NumberFormat = "yyyy-mm-dd"; r.HorizontalAlignment = -4108; break;
                    case 6:
                        r.Interior.Color = 0xC1C1FF; break;
                    case 7:
                        r.Interior.Color = 0xC6F0CE; break;
                    default:
                        int rows = (int)r.Rows.Count, cols = (int)r.Columns.Count;
                        int firstRow = (int)r.Row, firstCol = (int)r.Column;
                        dynamic ws = r.Worksheet;
                        for (int i = 0; i < rows; i += 2)
                        {
                            dynamic row = ws.Range[ws.Cells[firstRow + i, firstCol],
                                                   ws.Cells[firstRow + i, firstCol + cols - 1]];
                            row.Interior.Color = 0xF2F2F2;
                        }
                        break;
                }
            }
            catch (Exception ex) { Xl.Bad("套用样式失败：" + AiHttp.Sanitize(ex.Message)); }
            finally { Xl.Turn(false); }
            Xl.Say("样式已应用。");
            return true;
        }

        // ================= 4. 图表批处理 =================
        private static bool ChartSet()
        {
            dynamic ws = Xl.Act;
            if (ws == null) { Xl.Bad("请先打开工作簿。"); return true; }
            string[] p = Ask.Show("图表批处理",
                new string[] { "统一字体大小（留空不改）", "统一宽度(像素)", "统一高度(像素)", "显示数据标签(1/0)" },
                new string[] { "9", "", "", "0" });
            if (p == null) return true;
            int size = 0, w = 0, h = 0;
            int.TryParse(p[0].Trim(), out size);
            int.TryParse(p[1].Trim(), out w);
            int.TryParse(p[2].Trim(), out h);
            bool labels = p[3].Trim() == "1";

            int n = 0;
            Xl.Turn(true);
            try
            {
                int count = (int)ws.ChartObjects().Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic co = ws.ChartObjects(i);
                    dynamic ch = co.Chart;
                    try
                    {
                        if (size > 0) { ch.ChartArea.Font.Size = size; ch.ChartArea.Font.Name = "Microsoft YaHei UI"; }
                        if (w > 0) co.Width = w;
                        if (h > 0) co.Height = h;
                        if (labels)
                        {
                            int sc = (int)ch.SeriesCollection().Count;
                            for (int k = 1; k <= sc; k++)
                            {
                                try { ch.SeriesCollection(k).HasDataLabels = true; } catch { }
                            }
                        }
                        n++;
                    }
                    catch { }
                }
            }
            catch (Exception ex) { Xl.Bad("处理图表失败：" + AiHttp.Sanitize(ex.Message)); }
            finally { Xl.Turn(false); }
            Xl.Say(n == 0 ? "当前表没有图表。" : Xl.Fmt("已处理 {n} 个图表。", n));
            return true;
        }

        // ================= 5. 条码（Code39） =================
        private static readonly Dictionary<char, string> Code39 = new Dictionary<char, string>
        {
            {'0',"101001101101"},{'1',"110100101011"},{'2',"101100101011"},{'3',"110110010101"},
            {'4',"101001101011"},{'5',"110100110101"},{'6',"101100110101"},{'7',"101001011011"},
            {'8',"110100101101"},{'9',"101100101101"},{'A',"110101001011"},{'B',"101101001011"},
            {'C',"110110100101"},{'D',"101011001011"},{'E',"110101100101"},{'F',"101101100101"},
            {'G',"101010011011"},{'H',"110101001101"},{'I',"101101001101"},{'J',"101011001101"},
            {'K',"110101010011"},{'L',"101101010011"},{'M',"110110101001"},{'N',"101011010011"},
            {'O',"110101101001"},{'P',"101101101001"},{'Q',"101010110011"},{'R',"110101011001"},
            {'S',"101101011001"},{'T',"101011011001"},{'U',"110010101011"},{'V',"100110101011"},
            {'W',"110011010101"},{'X',"100101101011"},{'Y',"110010110101"},{'Z',"100110110101"},
            {'-',"100101011011"},{'.',"110010101101"},{' ',"100110101101"},{'$',"100100100101"},
            {'/',"100100101001"},{'+',"100101001001"},{'%',"101001001001"},{'*',"100101101101"}
        };

        private static bool Barcode()
        {
            dynamic r = Xl.SelRange("要生成条码的单元格（取文本，仅支持 0-9 A-Z - . $ / + %）");
            if (r == null) return true;
            dynamic cell = r.Cells[1, 1];
            string text = Xl.S(cell.Value2).Trim().ToUpperInvariant();
            if (text.Length == 0)
            {
                string input = Ask.One("生成条码", "条码内容", "");
                if (input == null) return true;
                text = input.Trim().ToUpperInvariant();
            }
            if (text.Length == 0) return true;
            foreach (char c in text)
                if (!Code39.ContainsKey(c)) { Xl.Bad("Code39 不支持字符：" + c); return true; }
            if (text.Length > 30) { Xl.Bad("内容过长（最多 30 个字符）。"); return true; }

            StringBuilder bits = new StringBuilder();
            bits.Append(Code39['*']);
            foreach (char c in text) bits.Append(Code39[c]);
            bits.Append(Code39['*']);

            // 渲染：每个 bit 3px 宽，含静区
            int bitW = 3, height = 60, quiet = 10;
            int width = bits.Length * bitW + quiet * 2;
            using (Bitmap bmp = new Bitmap(width, height))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.White);
                    using (SolidBrush b = new SolidBrush(Color.Black))
                    {
                        for (int i = 0; i < bits.Length; i++)
                        {
                            if (bits[i] != '1') continue;
                            g.FillRectangle(b, quiet + i * bitW, 4, bitW, height - 8);
                        }
                    }
                }
                string path = Path.Combine(Path.GetTempPath(), "exclehelper-barcode-" + Guid.NewGuid().ToString("N") + ".png");
                bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                try
                {
                    double left = (double)cell.Left + 2;
                    double top = (double)cell.Top + 2;
                    double w = Math.Min(width, 260);
                    double h = w * height / width;
                    dynamic sh = cell.Worksheet.Shapes.AddPicture(path, false, true, left, top, w, h);
                    sh.Placement = 1;
                    Xl.Say("已插入 Code39 条码：" + text + "（" + bits.Length + " 模块）。");
                }
                catch (Exception ex) { Xl.Bad("插入条码失败：" + AiHttp.Sanitize(ex.Message)); }
            }
            return true;
        }

        // ================= 6. 区域导出图片 =================
        private static bool ExportRange()
        {
            dynamic r = Xl.SelRange("要导出为图片的区域");
            if (r == null) return true;
            string path;
            using (SaveFileDialog d = new SaveFileDialog())
            {
                d.Title = "导出选区为图片";
                d.Filter = "PNG 图片|*.png";
                d.FileName = "选区.png";
                if (d.ShowDialog() != DialogResult.OK) return true;
                path = d.FileName;
            }
            Xl.Turn(true);
            try
            {
                r.CopyPicture(1, 2);   // xlScreen, xlBitmap
                dynamic co = r.Worksheet.ChartObjects().Add(0, 0, Math.Max(120, (double)r.Width), Math.Max(80, (double)r.Height));
                dynamic ch = co.Chart;
                ch.Paste();
                ch.Export(path, "PNG");
                co.Delete();
                Xl.Say("已导出到：\n" + path);
            }
            catch (Exception ex) { Xl.Bad("导出失败：" + AiHttp.Sanitize(ex.Message)); }
            finally
            {
                try { Xl.App.CutCopyMode = 0; } catch { }
                Xl.Turn(false);
            }
            return true;
        }

        // ================= 7/8. 最近使用 / 收藏夹 =================
        private static string UiFile { get { return Path.Combine(AiConfig.Dir, "usage.ini"); } }

        internal static void RecordUse(string tag)
        {
            try
            {
                List<string> list = ReadList("recent");
                list.Remove(tag);
                list.Insert(0, tag);
                while (list.Count > 12) list.RemoveAt(list.Count - 1);
                WriteList("recent", list);
            }
            catch { }
        }

        internal static List<string> ReadList(string key)
        {
            List<string> list = new List<string>();
            try
            {
                if (!File.Exists(UiFile)) return list;
                foreach (string line in File.ReadAllLines(UiFile, Encoding.UTF8))
                {
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    if (!string.Equals(line.Substring(0, eq).Trim(), key, StringComparison.OrdinalIgnoreCase)) continue;
                    foreach (string p in line.Substring(eq + 1).Split(','))
                    {
                        string t = p.Trim();
                        if (t.Length > 0) list.Add(t);
                    }
                }
            }
            catch { }
            return list;
        }

        internal static void WriteList(string key, List<string> list)
        {
            try
            {
                List<string> lines = new List<string>();
                if (File.Exists(UiFile)) lines.AddRange(File.ReadAllLines(UiFile, Encoding.UTF8));
                string newLine = key + "=" + string.Join(",", list.ToArray());
                bool replaced = false;
                for (int i = 0; i < lines.Count; i++)
                {
                    int eq = lines[i].IndexOf('=');
                    if (eq > 0 && string.Equals(lines[i].Substring(0, eq).Trim(), key, StringComparison.OrdinalIgnoreCase))
                    {
                        lines[i] = newLine;
                        replaced = true;
                        break;
                    }
                }
                if (!replaced) lines.Add(newLine);
                File.WriteAllLines(UiFile, lines.ToArray(), Encoding.UTF8);
            }
            catch { }
        }

        private static bool Recent()
        {
            List<string> list = ReadList("recent");
            if (list.Count == 0) { Xl.Say("还没有使用记录。"); return true; }
            return ShowList("最近使用", list, false);
        }

        private static bool Favorites()
        {
            List<string> list = ReadList("fav");
            return ShowList("收藏夹", list, true);
        }

        private static bool ShowList(string title, List<string> tags, bool editable)
        {
            using (UsageForm f = new UsageForm(title, tags, editable))
            {
                f.ShowDialog();
                return true;
            }
        }
    }

    /// <summary>最近使用 / 收藏夹 列表窗：双击执行，收藏夹可移除。</summary>
    internal sealed class UsageForm : Form
    {
        private ListBox _lb;
        private readonly bool _editable;
        private readonly List<string> _tags;

        internal UsageForm(string title, List<string> tags, bool editable)
        {
            _tags = new List<string>(tags);
            _editable = editable;
            Text = "EXCLE助手 · " + title;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            ClientSize = new Size(420, 380);
            Font = new Font("Microsoft YaHei UI", 9.5f);

            Label tip = new Label();
            tip.Text = editable ? "双击执行；点「移除」从收藏里删除" : "双击执行";
            tip.SetBounds(12, 10, 380, 20);
            tip.ForeColor = Color.FromArgb(100, 116, 139);
            Controls.Add(tip);

            _lb = new ListBox();
            _lb.SetBounds(12, 34, 396, 290);
            _lb.IntegralHeight = false;
            _lb.DoubleClick += delegate { Run(); };
            Controls.Add(_lb);
            Fill();

            Button run = new Button();
            run.Text = "执行";
            run.SetBounds(214, 336, 92, 30);
            run.Click += delegate { Run(); };
            Controls.Add(run);

            if (editable)
            {
                Button remove = new Button();
                remove.Text = "移除";
                remove.SetBounds(118, 336, 92, 30);
                remove.Click += delegate
                {
                    if (_lb.SelectedIndex < 0) return;
                    _tags.RemoveAt(_lb.SelectedIndex);
                    PresentOps.WriteList("fav", _tags);
                    Fill();
                };
                Controls.Add(remove);
            }

            Button close = new Button();
            close.Text = "关闭";
            close.SetBounds(314, 336, 92, 30);
            close.DialogResult = DialogResult.Cancel;
            Controls.Add(close);
        }

        private void Fill()
        {
            _lb.Items.Clear();
            foreach (string t in _tags)
            {
                string label = Journal.LabelOf(t);
                _lb.Items.Add(label + "    (" + t + ")");
            }
        }

        private void Run()
        {
            int i = _lb.SelectedIndex;
            if (i < 0 || i >= _tags.Count) return;
            string tag = _tags[i];
            Hide();
            try
            {
                Ops.Init();
                if (Ops.Has(tag)) Ops.Run(tag);
                else Xl.Bad("功能已不存在：" + tag);
            }
            catch (Exception ex) { Xl.Bad("执行失败：" + ex.Message); }
            finally { Close(); }
        }
    }
}

using System;
using System.Globalization;

namespace ExcleHelper
{
    /// <summary>格式美化域：行高列宽、边框、隔行填色、表头、冻结、格式复制。</summary>
    internal static class FormatOps
    {
        private static int Bgr(int r, int g, int b) { return b * 65536 + g * 256 + r; }

        internal static void Register()
        {
            Ops.Reg("f.rowheight", RowHeight);
            Ops.Reg("f.colwidth", ColWidth);
            Ops.Reg("f.autofit", AutoFit);
            Ops.Reg("f.border", Border);
            Ops.Reg("f.stripe", Stripe);
            Ops.Reg("f.headstyle", HeadStyle);
            Ops.Reg("f.freeze", Freeze);
            Ops.Reg("f.clearfmt", ClearFmt);
            Ops.Reg("f.copymulti", CopyFormatMulti);
        }

        private static bool RowHeight()
        {
            dynamic r = Xl.SelRange("要设置行高的区域");
            if (r == null) return true;
            double h;
            if (!Ask.Dbl("统一行高", "行高(磅，留空自动)", "18", out h)) return true;
            Xl.Turn(true);
            try
            {
                if (h <= 0) { r.Rows.AutoFit(); Xl.Say("已自动调整行高。"); return true; }
                r.RowHeight = h;
            }
            finally { Xl.Turn(false); }
            Xl.Say("已把选中区域行高设为 " + h.ToString(CultureInfo.InvariantCulture) + " 磅。");
            return true;
        }

        private static bool ColWidth()
        {
            dynamic r = Xl.SelRange("要设置列宽的区域");
            if (r == null) return true;
            double w;
            if (!Ask.Dbl("统一列宽", "列宽(字符数，留空自动)", "12", out w)) return true;
            Xl.Turn(true);
            try
            {
                if (w <= 0) { r.Columns.AutoFit(); Xl.Say("已自动调整列宽。"); return true; }
                r.ColumnWidth = w;
            }
            finally { Xl.Turn(false); }
            Xl.Say("已把选中区域列宽设为 " + w.ToString(CultureInfo.InvariantCulture) + "。");
            return true;
        }

        private static bool AutoFit()
        {
            dynamic r = Xl.SelRange("要自动调整的区域");
            if (r == null) return true;
            Xl.Turn(true);
            try
            {
                r.EntireColumn.AutoFit();
                r.EntireRow.AutoFit();
            }
            finally { Xl.Turn(false); }
            Xl.Say("已自动调整行高与列宽。");
            return true;
        }

        private static bool Border()
        {
            dynamic r = Xl.SelRange("要加边框的区域");
            if (r == null) return true;
            Xl.Turn(true);
            try
            {
                dynamic b = r.Borders;
                // 7..12 = 左上右下 + 内部横竖
                for (int i = 7; i <= 12; i++)
                {
                    dynamic edge = b[i];
                    edge.LineStyle = 1;      // xlContinuous
                    edge.Weight = 2;         // xlThin
                    edge.Color = Bgr(180, 190, 200);
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("已为选区加上细边框。");
            return true;
        }

        private static bool Stripe()
        {
            dynamic r = Xl.SelRange("要隔行填色的区域");
            if (r == null) return true;
            int rows = (int)r.Rows.Count;
            int cols = (int)r.Columns.Count;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            dynamic ws = r.Worksheet;
            string step = Ask.One("隔行填色", "每隔几行填一次色", "2");
            if (step == null) return true;
            int n;
            if (!int.TryParse(step, out n) || n < 2) { Xl.Bad("间隔需 >= 2。"); return true; }

            int light = Bgr(240, 244, 249);
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    if ((i + 1) % n != 0) continue;
                    dynamic a = ws.Cells[firstRow + i, firstCol];
                    dynamic e = ws.Cells[firstRow + i, firstCol + cols - 1];
                    ws.Range(a, e).Interior.Color = light;
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已为 {n} 行填色。", rows / n));
            return true;
        }

        private static bool HeadStyle()
        {
            dynamic r = Xl.SelRange("表头区域");
            if (r == null) return true;
            bool first = Xl.Ask("把选区【第一行】作为表头美化？\n选「否」则美化整个选区。");
            dynamic target = first ? r.Rows[1] : r;
            Xl.Turn(true);
            try
            {
                target.Font.Bold = true;
                target.Font.Color = Bgr(255, 255, 255);
                target.Interior.Color = Bgr(37, 99, 235);
                target.HorizontalAlignment = -4108;   // xlCenter
                target.VerticalAlignment = -4108;
                target.RowHeight = 22;
                dynamic b = target.Borders;
                for (int i = 7; i <= 12; i++)
                {
                    b[i].LineStyle = 1;
                    b[i].Weight = 2;
                    b[i].Color = Bgr(37, 99, 235);
                }
            }
            finally { Xl.Turn(false); }
            Xl.Say("表头样式已应用（深蓝底 / 白字 / 加粗 / 居中 / 边框）。");
            return true;
        }

        private static bool Freeze()
        {
            dynamic r = Xl.SelRange("冻结位置的下方单元格");
            if (r == null) return true;
            dynamic ws = r.Worksheet;
            int row = (int)r.Row, col = (int)r.Column;
            Xl.Turn(true);
            try
            {
                ws.Activate();
                dynamic win = Xl.App.ActiveWindow;
                win.SplitRow = row - 1;
                win.SplitColumn = col - 1;
                win.FreezePanes = true;
            }
            finally { Xl.Turn(false); }
            Xl.Say("已冻结：第 " + (row - 1) + " 行以上 / 第 " + (col - 1) + " 列以左。");
            return true;
        }

        private static bool ClearFmt()
        {
            dynamic r = Xl.SelRange("要清除格式的区域");
            if (r == null) return true;
            if (!Xl.Ask("将清除选区的格式（保留内容），继续？")) return true;
            try { r.ClearFormats(); }
            catch { Xl.Bad("清除格式失败。"); return true; }
            Xl.Say("已清除选区格式。");
            return true;
        }

        private static bool CopyFormatMulti()
        {
            if (!Xl.Ready()) return true;
            dynamic src = Xl.PickRange("请选择【样板区域】（格式来源）");
            if (src == null) return true;
            dynamic ws = Xl.App.ActiveSheet;
            string spec = Ask.One("格式套用到多表", "目标工作表名(逗号分隔，留空=除当前表外的全部)", "");
            if (spec == null) return true;

            dynamic wb = Xl.Book;
            int n = (int)wb.Worksheets.Count;
            string[] names;
            if (spec.Trim().Length == 0)
            {
                System.Collections.Generic.List<string> list = new System.Collections.Generic.List<string>();
                string cur = (string)ws.Name;
                for (int i = 1; i <= n; i++)
                {
                    string nm = (string)wb.Worksheets[i].Name;
                    if (nm != cur) list.Add(nm);
                }
                names = list.ToArray();
            }
            else
            {
                names = spec.Split(new char[] { ',', '，', '|' });
            }
            int ok = 0;
            Xl.Turn(true);
            try
            {
                src.Copy();
                foreach (string raw in names)
                {
                    string nm = raw.Trim();
                    if (nm.Length == 0) continue;
                    dynamic t;
                    try { t = wb.Worksheets[nm]; } catch { continue; }
                    try
                    {
                        t.Activate();
                        dynamic cell = t.Cells[(int)src.Row, (int)src.Column];
                        cell.PasteSpecial(-4122, -4142, false, false);  // 只粘贴格式
                        ok++;
                    }
                    catch { }
                }
                Xl.App.CutCopyMode = 0;
            }
            finally { Xl.Turn(false); }
            Xl.Say("已把样板格式套用到 " + ok + " 张工作表。");
            return true;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>
    /// 录入与验证域（批次 2）：批量下拉、二级联动、拒绝重复、数值范围、
    /// 复选框/单选框、快速录入词库。
    /// </summary>
    internal static class InputOps
    {
        internal static void Register()
        {
            Ops.Reg("vd.list", List);
            Ops.Reg("vd.list.cascade", Cascade);
            Ops.Reg("vd.no.dup", NoDup);
            Ops.Reg("vd.range", RangeLimit);
            Ops.Reg("vd.checkbox", CheckBoxes);
            Ops.Reg("vd.radio", Radios);
            Ops.Reg("vd.quick", QuickInput);
        }

        // ---------------- 1. 批量下拉 ----------------
        private static bool List()
        {
            dynamic r = Xl.SelRange("要应用下拉的区域");
            if (r == null) return true;
            string mode = Ask.One("批量插入下拉", "1 = 手动输入选项（逗号分隔）\n2 = 引用某个区域作为选项", "1");
            if (mode == null) return true;

            string formula;
            if (mode.Trim() == "2")
            {
                dynamic src = Xl.PickRange("请选择选项来源区域");
                if (src == null) return true;
                // External:=true 带上工作表名，来源在别的表时也能正确引用
                string addr = (string)src.Address(true, true, 1, true);
                formula = "=" + addr;
            }
            else
            {
                string opts = Ask.One("批量插入下拉", "选项（逗号分隔）", "是,否");
                if (opts == null || opts.Trim().Length == 0) return true;
                formula = "=" + Quote(opts);
            }

            Xl.Turn(true);
            try
            {
                r.Validation.Delete();
                r.Validation.Add(3, 1, 1, formula);   // xlValidateList, xlValidAlertStop, xlBetween
                r.Validation.IgnoreBlank = true;
                r.Validation.InCellDropdown = true;
                r.Validation.ShowError = true;
                r.Validation.ErrorTitle = "输入无效";
                r.Validation.ErrorMessage = "请从下拉列表中选择。";
            }
            catch (Exception ex) { Xl.Bad("添加数据验证失败：" + ex.Message); return true; }
            finally { Xl.Turn(false); }
            Xl.Say("已为 " + Xl.Address(r) + " 添加下拉列表。");
            return true;
        }

        private static string Quote(string csv)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append('"');
            foreach (string s in csv.Split(new char[] { ',', '，', '|' }))
            {
                string t = s.Trim();
                if (t.Length == 0) continue;
                if (sb.Length > 1) sb.Append(',');
                sb.Append(t.Replace("\"", ""));
            }
            sb.Append('"');
            return sb.ToString();
        }

        // ---------------- 2. 二级联动 ----------------
        private static bool Cascade()
        {
            if (Xl.Book == null) { Xl.Bad("请先打开工作簿。"); return true; }
            dynamic map = Xl.PickRange("请选择【父类 + 子类】两列数据（不含标题）");
            if (map == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(map, out rows, out cols);
            if (cols != 2) { Xl.Bad("请选择两列：第一列父类、第二列子类。"); return true; }

            dynamic wb = Xl.Book;
            // 建立字典：父类 -> 子类集合
            List<string> parents = new List<string>();
            Dictionary<string, List<string>> kids = new Dictionary<string, List<string>>(StringComparer.Ordinal);
            for (int i = 0; i < rows; i++)
            {
                string p = Xl.S(a[i, 0]).Trim();
                string c = Xl.S(a[i, 1]).Trim();
                if (p.Length == 0 || c.Length == 0) continue;
                List<string> list;
                if (!kids.TryGetValue(p, out list)) { list = new List<string>(); kids[p] = list; parents.Add(p); }
                if (!list.Contains(c)) list.Add(c);
            }
            if (parents.Count == 0) { Xl.Bad("没有读到有效的父子数据。"); return true; }

            Xl.Turn(true);
            try
            {
                // 隐藏字典表：第 1 行父类名，下面是对应子项
                dynamic lk = wb.Worksheets.Add(Type.Missing, wb.Worksheets[wb.Worksheets.Count]);
                lk.Name = Xl.SafeName(wb, "_lk");
                int col = 1;
                System.Collections.Generic.HashSet<string> usedNames = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (string p in parents)
                {
                    lk.Cells[1, col].Value2 = p;
                    List<string> list = kids[p];
                    for (int k = 0; k < list.Count; k++) lk.Cells[k + 2, col].Value2 = list[k];
                    // 定义名称（父类名 -> 该列子项）
                    string safe = MakeName(p);
                    while (usedNames.Contains(safe)) safe = safe + "_" + col;   // 防止不同父类撞名
                    usedNames.Add(safe);
                    try
                    {
                        // RefersTo 必须带工作表名，否则取决于 Names.Add 时的活动表
                        wb.Names.Add(safe, "='" + (string)lk.Name + "'!" + (string)lk.Cells[2, col].Address(true, true)
                                          + ":" + (string)lk.Cells[list.Count + 1, col].Address(true, true));
                    }
                    catch { }
                    col++;
                }
                lk.Visible = 0;

                // 在用户指定的两个单元格上分别建父/子下拉
                string[] p2 = Ask.Show("二级联动",
                    new string[] { "父类下拉放在哪个单元格（如 B2）", "子类下拉放在哪个单元格（如 C2）" },
                    new string[] { "B2", "C2" });
                if (p2 == null) { Xl.Turn(false); return true; }
                dynamic ws = Xl.Act;
                string parentAddr = p2[0].Trim();
                string childAddr = p2[1].Trim();
                dynamic pc = ws.Range[parentAddr];
                dynamic cc = ws.Range[childAddr];

                pc.Validation.Delete();
                pc.Validation.Add(3, 1, 1, "=" + string.Join(",", parents.ToArray()));
                pc.Validation.InCellDropdown = true;

                cc.Validation.Delete();
                cc.Validation.Add(3, 1, 1, "=INDIRECT(" + (string)pc.Address(false, false).Replace("$", "") + ")");
                cc.Validation.InCellDropdown = true;

                Xl.Say("二级联动已建立：\n父类 " + parentAddr + "（" + parents.Count
                       + " 项）\n子类 " + childAddr + "（随父类变化）\n字典表「" + (string)lk.Name + "」已隐藏。");
            }
            catch (Exception ex)
            {
                Xl.Bad("建立联动失败：" + ex.Message);
            }
            finally { Xl.Turn(false); }
            return true;
        }

        private static string MakeName(string p)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in p)
            {
                if (char.IsLetterOrDigit(c) || c == '_') sb.Append(c);
                else sb.Append('_');
            }
            if (sb.Length == 0) sb.Append("p");
            if (sb.Length > 20) sb.Length = 20;
            if (char.IsDigit(sb[0])) sb.Insert(0, "p_");
            return sb.ToString();
        }

        // ---------------- 3. 拒绝重复 ----------------
        private static bool NoDup()
        {
            dynamic r = Xl.SelRange("要防止重复输入的区域");
            if (r == null) return true;
            string addr = (string)r.Address(false, false);
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            string col = ColLetter(firstCol);
            string formula = "=COUNTIF($" + col + "$" + firstRow + ":$" + col + "$1048576,"
                             + (string)r.Cells[1, 1].Address(false, false) + ")=1";
            Xl.Turn(true);
            try
            {
                r.Validation.Delete();
                r.Validation.Add(0, 1, 1, formula);   // xlValidateCustom
                r.Validation.IgnoreBlank = true;
                r.Validation.ShowError = true;
                r.Validation.ErrorTitle = "重复输入";
                r.Validation.ErrorMessage = "该列已存在相同内容，请更换。";
            }
            catch (Exception ex) { Xl.Bad("设置失败：" + ex.Message); return true; }
            finally { Xl.Turn(false); }
            Xl.Say("已禁止在 " + addr + " 中输入重复值。");
            return true;
        }

        private static string ColLetter(int col)
        {
            StringBuilder sb = new StringBuilder();
            while (col > 0)
            {
                int m = (col - 1) % 26;
                sb.Insert(0, (char)('A' + m));
                col = (col - 1) / 26;
            }
            return sb.ToString();
        }

        // ---------------- 4. 数值范围 ----------------
        private static bool RangeLimit()
        {
            dynamic r = Xl.SelRange("要限制范围的区域");
            if (r == null) return true;
            string[] p = Ask.Show("数值范围限制",
                new string[] { "最小值", "最大值", "超出提示" },
                new string[] { "0", "100", "请输入 0-100 之间的数字" });
            if (p == null) return true;
            double lo, hi;
            if (!double.TryParse(p[0], NumberStyles.Any, CultureInfo.InvariantCulture, out lo) ||
                !double.TryParse(p[1], NumberStyles.Any, CultureInfo.InvariantCulture, out hi))
            {
                Xl.Bad("请输入数字。");
                return true;
            }
            Xl.Turn(true);
            try
            {
                r.Validation.Delete();
                r.Validation.Add(2, 1, 1, lo, hi);   // xlValidateDecimal, xlBetween
                r.Validation.IgnoreBlank = true;
                r.Validation.ShowError = true;
                r.Validation.ErrorTitle = "超出范围";
                r.Validation.ErrorMessage = p[2];
            }
            catch (Exception ex) { Xl.Bad("设置失败：" + ex.Message); return true; }
            finally { Xl.Turn(false); }
            Xl.Say("已限制 " + Xl.Address(r) + " 只能输入 " + lo + " ~ " + hi + "。");
            return true;
        }

        // ---------------- 5. 复选框 ----------------
        private static bool CheckBoxes()
        {
            dynamic r = Xl.SelRange("每行要放复选框的单元格区域");
            if (r == null) return true;
            int rows = (int)r.Rows.Count, cols = (int)r.Columns.Count;
            if (rows * cols > 500) { Xl.Bad("一次最多 500 个，请缩小选区。"); return true; }
            dynamic ws = r.Worksheet;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            int ok = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        dynamic cell = ws.Cells[firstRow + i, firstCol + j];
                        double left = (double)cell.Left + 2;
                        double top = (double)cell.Top + 2;
                        dynamic sh = ws.Shapes.AddFormControl(1, left, top, 16, 16);   // xlCheckBox
                        try
                        {
                            sh.LinkedCell = (string)ws.Cells[firstRow + i, firstCol + j].Address(false, false);
                            sh.OnAction = "";
                        }
                        catch { }
                        ok++;
                    }
                }
            }
            catch (Exception ex) { Xl.Bad("插入复选框失败：" + ex.Message); }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已插入 {n} 个复选框（勾选结果写回所在单元格 TRUE/FALSE）。", ok));
            return true;
        }

        // ---------------- 6. 单选框 ----------------
        private static bool Radios()
        {
            dynamic r = Xl.SelRange("每行放一组的单选按钮区域");
            if (r == null) return true;
            int rows = (int)r.Rows.Count, cols = (int)r.Columns.Count;
            if (rows > 100) { Xl.Bad("一次最多 100 行。"); return true; }
            dynamic ws = r.Worksheet;
            int firstRow = (int)r.Row, firstCol = (int)r.Column;
            int groups = 0, unbound = 0;
            Xl.Turn(true);
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    dynamic c1 = ws.Cells[firstRow + i, firstCol];
                    dynamic c2 = ws.Cells[firstRow + i, firstCol + cols - 1];
                    double left = (double)c1.Left;
                    double top = (double)c1.Top;
                    double width = (double)c2.Left + (double)c2.Width - left;
                    ws.Shapes.AddFormControl(4, left, top, width, 18);                 // xlGroupBox
                    for (int j = 0; j < cols; j++)
                    {
                        dynamic cell = ws.Cells[firstRow + i, firstCol + j];
                        double bl = (double)cell.Left + 2;
                        double bt = (double)cell.Top + 3;
                        dynamic rb = ws.Shapes.AddFormControl(7, bl, bt, 14, 14);      // xlOptionButton
                        try { rb.LinkedCell = (string)ws.Cells[firstRow + i, firstCol + cols].Address(false, false); }
                        catch { unbound++; }
                    }
                    groups++;
                }
            }
            catch (Exception ex) { Xl.Bad("插入单选框失败：" + ex.Message); }
            finally { Xl.Turn(false); }
            Xl.Say(Xl.Fmt("已为 {n} 行各建一组单选按钮（选中项编号写入行尾单元格）。", groups)
                   + (unbound > 0 ? "其中 " + unbound + " 个控件未能绑定单元格。" : ""));
            return true;
        }

        // ---------------- 7. 快速录入词库 ----------------
        private static bool QuickInput()
        {
            dynamic src = Xl.PickRange("请选择词库区域（单列或多列，每格一个词条）");
            if (src == null) return true;
            int rows, cols;
            object[,] a = Xl.Vals(src, out rows, out cols);
            List<string> words = new List<string>();
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                {
                    string s = Xl.S(a[i, j]).Trim();
                    if (s.Length > 0 && !words.Contains(s)) words.Add(s);
                }
            if (words.Count == 0) { Xl.Bad("词库为空。"); return true; }

            string picked = QuickPick.Pick(words);
            if (picked == null) return true;
            dynamic target = Xl.Sel;
            Xl.Turn(true);
            try { target.Value2 = picked; }
            finally { Xl.Turn(false); }
            Xl.Say("已写入：" + picked);
            return true;
        }
    }

    /// <summary>快速录入词库选择窗（紧凑列表，双击或回车写入）。</summary>
    internal sealed class QuickPick : Form
    {
        private ListBox _lb;
        private string _result;

        internal static string Pick(List<string> words)
        {
            using (QuickPick f = new QuickPick(words))
            {
                return f.ShowDialog() == DialogResult.OK ? f._result : null;
            }
        }

        private QuickPick(List<string> words)
        {
            Text = "快速录入";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new System.Drawing.Size(360, 420);
            Font = new System.Drawing.Font("Microsoft YaHei UI", 9.5f);
            KeyPreview = true;

            Label tip = new Label();
            tip.Text = "双击或回车写入当前单元格";
            tip.SetBounds(12, 10, 330, 20);
            tip.ForeColor = System.Drawing.Color.FromArgb(100, 116, 139);
            Controls.Add(tip);

            _lb = new ListBox();
            _lb.SetBounds(12, 34, 336, 336);
            _lb.IntegralHeight = false;
            _lb.DoubleClick += delegate { Accept(); };
            foreach (string w in words) _lb.Items.Add(w);
            Controls.Add(_lb);

            Button ok = new Button();
            ok.Text = "写入";
            ok.SetBounds(168, 376, 88, 30);
            ok.Click += delegate { Accept(); };
            Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "取消";
            cancel.SetBounds(260, 376, 88, 30);
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);

            _lb.KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter) { Accept(); e.SuppressKeyPress = true; }
            };
            KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Escape) { DialogResult = DialogResult.Cancel; Close(); }
            };
            Load += delegate { if (_lb.Items.Count > 0) _lb.SelectedIndex = 0; _lb.Focus(); };
        }

        private void Accept()
        {
            if (_lb.SelectedItem == null) return;
            _result = _lb.SelectedItem.ToString();
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}

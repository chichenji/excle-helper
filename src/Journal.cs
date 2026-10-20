using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcleHelper
{
    /// <summary>
    /// 回退系统：Excel 原生撤销栈捕获不到 COM 写入（这是同类插件的通病），
    /// 这里在每次操作前对「当前工作表 + 余量区域」做一次快照（公式/值 + 数字格式），
    /// 并提供多步回退、重做与历史记录。快照在内存中保留最近 20 步。
    /// </summary>
    internal static class Journal
    {
        private sealed class Snap
        {
            internal string Tag;
            internal string Label;
            internal DateTime At;
            internal string Book;
            internal string Sheet;
            internal int R0, C0, Rows, Cols;
            internal object[,] Formula;     // 公式或值（统一用 Formula 读，赋值时公式仍是公式）
            internal object[,] Format;      // NumberFormat
            internal List<string> Sheets;   // 快照时的全部工作表名
            internal int ActiveIdx;
            internal bool Skipped;
        }

        private static readonly List<Snap> UndoStack = new List<Snap>();
        private static readonly List<Snap> RedoStack = new List<Snap>();
        private static Snap _pending;
        private static bool _busy;

        private const int MaxDepth = 20;
        private const long MaxCells = 250000;
        private const int MarginCols = 12;
        private const int MarginRows = 60;

        internal static int UndoCount { get { return UndoStack.Count; } }
        internal static int RedoCount { get { return RedoStack.Count; } }

        internal static string LabelOf(string tag)
        {
            try
            {
                foreach (Feature f in Catalog.All)
                    if (string.Equals(f.Tag, tag, StringComparison.OrdinalIgnoreCase)) return f.Label;
            }
            catch { }
            return tag;
        }

        internal static List<string> HistoryLines()
        {
            List<string> list = new List<string>();
            for (int i = UndoStack.Count - 1; i >= 0; i--)
            {
                Snap s = UndoStack[i];
                list.Add(string.Format(CultureInfo.InvariantCulture, "{0:HH:mm:ss}  {1}  （{2}）",
                                       s.At, s.Label, s.Sheet));
            }
            return list;
        }

        /// <summary>操作执行前调用：抓取快照。回退/重做自身不记录。</summary>
        internal static void Begin(string tag)
        {
            _pending = null;
            if (_busy) return;
            if (tag != null && tag.StartsWith("undo.", StringComparison.OrdinalIgnoreCase)) return;
            try
            {
                if (Xl.Book == null) return;
                dynamic ws = Xl.Act;
                if (ws == null) return;

                Snap s = new Snap();
                s.Tag = tag;
                s.Label = LabelOf(tag);
                s.At = DateTime.Now;
                s.Book = (string)Xl.Book.Name;
                s.Sheet = (string)ws.Name;

                dynamic ur = ws.UsedRange;
                s.R0 = (int)ur.Row;
                s.C0 = (int)ur.Column;
                int rows = (int)ur.Rows.Count + MarginRows;
                int cols = (int)ur.Columns.Count + MarginCols;
                if (s.R0 + rows - 1 > 1048576) rows = 1048576 - s.R0 + 1;
                if (s.C0 + cols - 1 > 16384) cols = 16384 - s.C0 + 1;

                if ((long)rows * cols > MaxCells)
                {
                    rows = (int)Math.Max(1, MaxCells / Math.Max(1, cols));
                    if ((long)rows * cols > MaxCells) { s.Skipped = true; }
                }
                s.Rows = rows;
                s.Cols = cols;

                if (!s.Skipped)
                {
                    dynamic block = ws.Range(ws.Cells[s.R0, s.C0], ws.Cells[s.R0 + rows - 1, s.C0 + cols - 1]);
                    s.Formula = ToZeroBased(block.Formula, rows, cols);
                    s.Format = ToZeroBased(block.NumberFormat, rows, cols);
                }

                s.Sheets = new List<string>();
                dynamic sheets = Xl.Book.Worksheets;
                int n = (int)sheets.Count;
                for (int i = 1; i <= n; i++) s.Sheets.Add((string)sheets[i].Name);
                try { s.ActiveIdx = (int)ws.Index; } catch { s.ActiveIdx = 1; }

                _pending = s;
            }
            catch { _pending = null; }
        }

        /// <summary>操作成功后调用：把快照压入撤销栈。</summary>
        internal static void Commit()
        {
            if (_pending == null) return;
            Snap s = _pending;
            _pending = null;
            if (s.Skipped) return;               // 区域过大未快照，本次不提供回退
            UndoStack.Add(s);
            RedoStack.Clear();
            while (UndoStack.Count > MaxDepth) UndoStack.RemoveAt(0);
        }

        internal static void Abort()
        {
            _pending = null;
        }

        private static object[,] ToZeroBased(object raw, int rows, int cols)
        {
            object[,] a = new object[rows, cols];
            Array arr = raw as Array;
            if (arr == null) { a[0, 0] = raw; return a; }
            int lb0 = arr.GetLowerBound(0), lb1 = arr.GetLowerBound(1);
            int rMax = Math.Min(rows, arr.GetUpperBound(0) - lb0 + 1);
            int cMax = Math.Min(cols, arr.GetUpperBound(1) - lb1 + 1);
            for (int i = 0; i < rMax; i++)
                for (int j = 0; j < cMax; j++)
                    a[i, j] = arr.GetValue(lb0 + i, lb1 + j);
            return a;
        }

        private static object[,] ToOneBased(object[,] a)
        {
            int rows = a.GetLength(0), cols = a.GetLength(1);
            object[,] o = new object[rows, cols];
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                    o[i, j] = a[i, j];
            // C# 数组本身是 0 基，COM 赋值接受 0 基 object[,]；这里直接返回
            return o;
        }

        // ---------------- 回退 / 重做 ----------------
        internal static bool UndoLast()
        {
            if (UndoStack.Count == 0) { Xl.Say("没有可回退的操作。"); return true; }
            Snap s = UndoStack[UndoStack.Count - 1];
            Snap current = CaptureCurrent("__redo__");
            if (!Restore(s)) return true;
            UndoStack.RemoveAt(UndoStack.Count - 1);
            if (current != null) RedoStack.Add(current);
            Xl.Say("已回退：" + s.Label + "\n（工作簿 " + s.Book + " · 工作表 " + s.Sheet + "）");
            return true;
        }

        internal static bool RedoLast()
        {
            if (RedoStack.Count == 0) { Xl.Say("没有可重做的操作。"); return true; }
            Snap s = RedoStack[RedoStack.Count - 1];
            Snap current = CaptureCurrent("__undo__");
            if (!Restore(s)) return true;
            RedoStack.RemoveAt(RedoStack.Count - 1);
            if (current != null) UndoStack.Add(current);
            Xl.Say("已重做：" + s.Label);
            return true;
        }

        /// <summary>回退到「第 index 步之前」：index 从 0 开始按 HistoryLines 的顺序（最新在前）。</summary>
        internal static bool RollbackTo(int indexFromNewest)
        {
            if (indexFromNewest < 0 || indexFromNewest >= UndoStack.Count) return false;
            int target = UndoStack.Count - 1 - indexFromNewest;
            Snap s = UndoStack[target];
            Snap current = CaptureCurrent("__redo__");
            if (!Restore(s)) return false;
            for (int i = UndoStack.Count - 1; i >= target; i--) UndoStack.RemoveAt(i);
            if (current != null) RedoStack.Add(current);
            Xl.Say("已回退到：" + s.Label + " 之前");
            return true;
        }

        internal static void Clear()
        {
            UndoStack.Clear();
            RedoStack.Clear();
            Xl.Say("回退历史已清空。");
        }

        private static Snap CaptureCurrent(string tag)
        {
            Snap keep = _pending;
            _pending = null;
            bool busy = _busy;
            _busy = false;
            try
            {
                Begin(tag);
                Snap s = _pending;
                _pending = keep;
                _busy = busy;
                return s;
            }
            catch
            {
                _pending = keep;
                _busy = busy;
                return null;
            }
        }

        private static bool Restore(Snap s)
        {
            _busy = true;
            Xl.Turn(true);
            try
            {
                dynamic wb = Xl.Book;
                if (wb == null) { Xl.Bad("请先打开工作簿。"); return false; }
                if (!string.Equals((string)wb.Name, s.Book, StringComparison.Ordinal))
                {
                    Xl.Bad("该步骤属于工作簿「" + s.Book + "」，请先切换到它再回退。");
                    return false;
                }

                // 1) 工作表结构：删除本次新增的表（重命名的情况不删，只提示）
                List<string> now = new List<string>();
                int n = (int)wb.Worksheets.Count;
                for (int i = 1; i <= n; i++) now.Add((string)wb.Worksheets[i].Name);
                List<string> added = new List<string>();
                foreach (string name in now) if (!s.Sheets.Contains(name)) added.Add(name);
                List<string> missing = new List<string>();
                foreach (string name in s.Sheets) if (!now.Contains(name)) missing.Add(name);

                bool renameLikely = added.Count > 0 && added.Count == missing.Count;
                foreach (string name in added)
                {
                    if (renameLikely) break;
                    try { wb.Worksheets[name].Delete(); } catch { }
                }

                // 2) 单元格内容与格式
                dynamic ws = null;
                try { ws = wb.Worksheets[s.Sheet]; }
                catch { ws = wb.Worksheets[Math.Min(s.ActiveIdx, (int)wb.Worksheets.Count)]; }
                ws.Activate();

                if (!s.Skipped && s.Formula != null)
                {
                    dynamic block = ws.Range(ws.Cells[s.R0, s.C0],
                                             ws.Cells[s.R0 + s.Rows - 1, s.C0 + s.Cols - 1]);
                    block.Formula = ToOneBased(s.Formula);
                    if (s.Format != null)
                    {
                        try { block.NumberFormat = ToOneBased(s.Format); } catch { }
                    }

                    // 3) 清掉快照区域之外、由该操作产生的多余内容
                    try
                    {
                        dynamic ur = ws.UsedRange;
                        int urEnd = (int)ur.Row + (int)ur.Rows.Count - 1;
                        int ucEnd = (int)ur.Column + (int)ur.Columns.Count - 1;
                        int snapEndR = s.R0 + s.Rows - 1;
                        int snapEndC = s.C0 + s.Cols - 1;
                        if (urEnd > snapEndR)
                        {
                            dynamic tail = ws.Range(ws.Cells[snapEndR + 1, 1], ws.Cells[urEnd, ucEnd]);
                            if ((long)(urEnd - snapEndR) * ucEnd < 200000) tail.ClearContents();
                        }
                        if (ucEnd > snapEndC)
                        {
                            dynamic tail = ws.Range(ws.Cells[1, snapEndC + 1], ws.Cells[urEnd, ucEnd]);
                            if ((long)urEnd * (ucEnd - snapEndC) < 200000) tail.ClearContents();
                        }
                    }
                    catch { }
                }

                if (renameLikely)
                    Xl.Say("注意：本次操作包含工作表重命名，已跳过自动改名（内容已回退）。");
                return true;
            }
            catch (Exception ex)
            {
                Xl.Bad("回退失败：" + ex.Message);
                return false;
            }
            finally
            {
                Xl.Turn(false);
                _busy = false;
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcleHelper
{
    /// <summary>Lookup domain: single-key / multi-key matching and existence checks.</summary>
    internal static class MatchOps
    {
        internal static void Register()
        {
            Ops.Reg("m.lookup.one", LookupOne);
            Ops.Reg("m.lookup.multi", LookupMulti);
            Ops.Reg("m.lookup.exists", Exists);
        }

        private static string Key(object[,] a, int row, int fromCol, int count)
        {
            if (count <= 1) return Xl.S(a[row, fromCol]);
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            for (int i = 0; i < count; i++)
            {
                sb.Append(Xl.S(a[row, fromCol + i]));
                sb.Append('\u0001');
            }
            return sb.ToString();
        }

        private static bool LookupOne()
        {
            if (!Xl.Ready()) return true;
            dynamic keys = Xl.PickRange("请选择【要查找的键值列】（单列）");
            if (keys == null) return true;
            dynamic table = Xl.PickRange("请选择【查找表】（第一行为标题）");
            if (table == null) return true;

            int kr = (int)keys.Rows.Count, kc = (int)keys.Columns.Count;
            int tr = (int)table.Rows.Count, tc = (int)table.Columns.Count;
            if (kc != 1) { Xl.Bad("键值区域请只选一列。"); return true; }
            if (tr < 2) { Xl.Bad("查找表至少需要标题行 + 一行数据。"); return true; }

            string[] p = Ask.Show("单条件查找",
                new string[] { "查找表中匹配列(第几列)", "返回第几列", "结果写入哪一列(键值右侧第几列)" },
                new string[] { "1", "2", "1" });
            if (p == null) return true;
            int keyCol, retCol, outOff;
            if (!int.TryParse(p[0], out keyCol) || !int.TryParse(p[1], out retCol) || !int.TryParse(p[2], out outOff))
            {
                Xl.Bad("请输入列号。");
                return true;
            }
            if (keyCol < 1 || keyCol > tc || retCol < 1 || retCol > tc)
            {
                Xl.Bad("列号超出查找表范围（共 " + tc + " 列）。");
                return true;
            }
            if (outOff < 1) outOff = 1;

            int rows, cols;
            object[,] ka = Xl.Vals(keys, out rows, out cols);
            object[,] ta = Xl.Vals(table, out tr, out tc);

            Dictionary<string, object> map = new Dictionary<string, object>(StringComparer.Ordinal);
            for (int i = 1; i < tr; i++)
            {
                string k = Key(ta, i, keyCol - 1, 1);
                if (!map.ContainsKey(k)) map[k] = ta[i, retCol - 1];
            }

            object[,] outv = new object[rows, 1];
            int hit = 0;
            for (int i = 0; i < rows; i++)
            {
                object v;
                string k = Xl.S(ka[i, 0]);
                if (map.TryGetValue(k, out v)) { outv[i, 0] = v; hit++; }
                else outv[i, 0] = "#N/A";
            }
            Xl.Turn(true);
            try { TextOps.WriteBlock(keys, 0, outOff, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("匹配完成：命中 " + hit.ToString(CultureInfo.InvariantCulture)
                   + " 行，未命中 " + (rows - hit).ToString(CultureInfo.InvariantCulture) + " 行（标记 #N/A）。");
            return true;
        }

        private static bool LookupMulti()
        {
            if (!Xl.Ready()) return true;
            dynamic keys = Xl.PickRange("请选择【要查找的多条件键值】（两列：条件1、条件2）");
            if (keys == null) return true;
            dynamic table = Xl.PickRange("请选择【查找表】（前两列为条件1、条件2，第一行为标题）");
            if (table == null) return true;

            int kr = (int)keys.Rows.Count, kc = (int)keys.Columns.Count;
            int tr = (int)table.Rows.Count, tc = (int)table.Columns.Count;
            if (kc != 2) { Xl.Bad("多条件键值区域请选择两列。"); return true; }
            if (tc < 3) { Xl.Bad("查找表至少需要三列（条件1、条件2、返回值）。"); return true; }

            string p = Ask.One("多条件查找", "返回第几列", "3");
            if (p == null) return true;
            int retCol;
            if (!int.TryParse(p, out retCol) || retCol < 1 || retCol > tc)
            {
                Xl.Bad("返回列号超出范围。");
                return true;
            }

            int rows, cols;
            object[,] ka = Xl.Vals(keys, out rows, out cols);
            object[,] ta = Xl.Vals(table, out tr, out tc);

            Dictionary<string, object> map = new Dictionary<string, object>(StringComparer.Ordinal);
            for (int i = 1; i < tr; i++)
            {
                string k = Key(ta, i, 0, 2);
                if (!map.ContainsKey(k)) map[k] = ta[i, retCol - 1];
            }

            object[,] outv = new object[rows, 1];
            int hit = 0;
            for (int i = 0; i < rows; i++)
            {
                string k = Key(ka, i, 0, 2);
                object v;
                if (map.TryGetValue(k, out v)) { outv[i, 0] = v; hit++; }
                else outv[i, 0] = "#N/A";
            }
            Xl.Turn(true);
            try { TextOps.WriteBlock(keys, 0, 2, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("多条件匹配完成：命中 " + hit.ToString(CultureInfo.InvariantCulture)
                   + " 行，未命中 " + (rows - hit).ToString(CultureInfo.InvariantCulture) + " 行。");
            return true;
        }

        private static bool Exists()
        {
            if (!Xl.Ready()) return true;
            dynamic keys = Xl.PickRange("请选择【要检查的键值列】");
            if (keys == null) return true;
            dynamic table = Xl.PickRange("请选择【对照表】（第一行为标题）");
            if (table == null) return true;

            int tr = (int)table.Rows.Count, tc = (int)table.Columns.Count;
            string p = Ask.One("查是否存在", "对照表中比对第几列", "1");
            if (p == null) return true;
            int keyCol;
            if (!int.TryParse(p, out keyCol) || keyCol < 1 || keyCol > tc)
            {
                Xl.Bad("列号超出对照表范围。");
                return true;
            }

            int rows, cols;
            object[,] ka = Xl.Vals(keys, out rows, out cols);
            object[,] ta = Xl.Vals(table, out tr, out tc);
            HashSet<string> set = new HashSet<string>(StringComparer.Ordinal);
            for (int i = 1; i < tr; i++) set.Add(Xl.S(ta[i, keyCol - 1]));

            object[,] outv = new object[rows, 1];
            int hit = 0;
            for (int i = 0; i < rows; i++)
            {
                bool ok = set.Contains(Xl.S(ka[i, 0]));
                outv[i, 0] = ok ? "是" : "否";
                if (ok) hit++;
            }
            Xl.Turn(true);
            try { TextOps.WriteBlock(keys, 0, 1, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("检查完成：存在 " + hit.ToString(CultureInfo.InvariantCulture)
                   + " 个，不存在 " + (rows - hit).ToString(CultureInfo.InvariantCulture) + " 个。");
            return true;
        }
    }
}

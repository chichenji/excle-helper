using System;

namespace ExcleHelper
{
    /// <summary>Housekeeping: help, about, repeat-last.</summary>
    internal static class MiscOps
    {
        internal static void Register()
        {
            Ops.Reg("x.about", About);
            Ops.Reg("x.help", Help);
            Ops.Reg("x.repeat", Repeat);
        }

        private static bool About()
        {
            string msg =
                "EXCLE助手  " + "1.0.0" + "\r\n" +
                "──────────────────────────\r\n" +
                "一个独立实现的 Excel 表格效率插件。\r\n" +
                "• 注册功能数：" + Ops.Count + "\r\n" +
                "• 架构：单 DLL 经典 COM 加载项 + 声明式 Ribbon + 操作注册表\r\n" +
                "• 依赖：无（不需要 VSTO / PIA / Aspose / NPOI / 联网）\r\n" +
                "• 授权：无账户、无会员、无功能门禁，全部功能直接可用\r\n" +
                "• 所有运算在本机完成，不上传任何数据\r\n" +
                "──────────────────────────\r\n" +
                "提示：选中区域后点按钮即可；参数在工作流内按需询问。";
            Xl.Say(msg);
            return true;
        }

        private static bool Help()
        {
            string msg =
                "使用方法\r\n" +
                "──────────────────────────\r\n" +
                "1) 先在 Excel 中选中要处理的数据区域。\r\n" +
                "2) 在「EXCLE助手」选项卡里点对应按钮。\r\n" +
                "3) 需要参数的按钮会弹出一个小对话框，填完点确定。\r\n" +
                "4) 需要框选区域时，会弹出一个「区域选择」输入框，\r\n" +
                "   用鼠标直接拖选即可（和 Excel 原生选择一致）。\r\n" +
                "\r\n" +
                "结果去向\r\n" +
                "• 文本/数值类：就地修改原选区（跳过公式）。\r\n" +
                "• 分列、合并、查找匹配、身份证解析：写入选区右侧。\r\n" +
                "• 汇总、去重清单、对比结果：生成新的工作表。\r\n" +
                "\r\n" +
                "常用组合\r\n" +
                "• 清洗：删除所有空格 → 删除不可见字符 → 只保留中文\r\n" +
                "• 拆分：选一列做依据 → 拆分工作表（保留原格式）\r\n" +
                "• 匹配：选键值列 → 选查找表 → 单条件查找\r\n" +
                "• 报表：分类汇总 / 按颜色统计 / 两区域对比\r\n" +
                "\r\n" +
                "所有类别均无试用限制。";
            Xl.Say(msg);
            return true;
        }

        private static bool Repeat()
        {
            if (string.IsNullOrEmpty(Xl.LastTag))
            {
                Xl.Say("还没有执行过任何功能。");
                return true;
            }
            if (string.Equals(Xl.LastTag, "x.repeat", StringComparison.OrdinalIgnoreCase))
            {
                Xl.Say("上一次操作就是「重复上次」。");
                return true;
            }
            Ops.Run(Xl.LastTag);
            return true;
        }
    }
}

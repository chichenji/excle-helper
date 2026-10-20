using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcleHelper
{
    /// <summary>
    /// 批次 C · AI 三期：报告生成、多表关联建议、整表公式纠错、批量翻译、合成测试数据、异常解释。
    /// 复用 AiExtOps 的调用与校验链路；所有落盘都走 Ops 注册表（自动可回退）。
    /// </summary>
    internal static class AiAdvOps
    {
        internal static void Register()
        {
            Ops.Reg("ai.report", Report);
            Ops.Reg("ai.join", JoinAdvice);
            Ops.Reg("ai.fix.all", FixAll);
            Ops.Reg("ai.translate", Translate);
            Ops.Reg("ai.synth", Synth);
            Ops.Reg("ai.explain.outlier", ExplainOutlier);
        }

        // ================= 1. 报告生成 =================
        private static bool Report()
        {
            AiConfig cfg = AiConfig.Load();
            string key = AiExtOps.ApiKey();
            if (!AiExtOps.Ready(cfg, key)) return true;
            dynamic r = Xl.SelRange("要做成报告的数据区域（含表头）");
            if (r == null) return true;

            int rows, cols;
            string addr;
            string digest = AiTasks.Digest(r, cfg, out rows, out cols, out addr);
            List<AiMessage> msgs = new List<AiMessage>();
            msgs.Add(new AiMessage("system",
                "你是数据分析报告撰写助手。基于摘要写一份中文报告，包含四段：\n" +
                "【概览】一段话说明数据范围与口径；\n" +
                "【关键指标】3-6 条带数字的要点；\n" +
                "【发现的模式或异常】；\n" +
                "【建议行动】。\n" +
                "每段独立成行，不要输出 Markdown 表格。不要编造摘要里没有的数据。"));
            msgs.Add(new AiMessage("user", digest));

            string err;
            string text = AiExtOps.Call(cfg, key, msgs, "AI 正在撰写报告", out err);
            if (text == null) { Xl.Bad("请求失败：" + (err ?? "未知错误")); return true; }
            if (!AiPreview.Show("AI 分析报告（" + addr + "）", text, true)) return true;

            string[] lines = text.Replace("\r", "").Split('\n');
            List<object[]> outRows = new List<object[]>();
            outRows.Add(new object[] { "AI 分析报告 · " + addr });
            outRows.Add(new object[] { "生成时间：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm") });
            outRows.Add(new object[] { "" });
            foreach (string line in lines)
            {
                string t = line.Trim();
                if (t.Length == 0) { outRows.Add(new object[] { "" }); continue; }
                outRows.Add(new object[] { t });
            }
            object[,] data = new object[outRows.Count, 1];
            for (int i = 0; i < outRows.Count; i++) data[i, 0] = outRows[i][0];
            dynamic sh = Xl.AddSheet("AI报告");
            Xl.Dump(sh, 1, 1, data);
            try
            {
                sh.Columns[1].ColumnWidth = 110;
                sh.Columns[1].WrapText = true;
                sh.Rows[1].Font.Bold = true;
                sh.Rows[1].Font.Size = 13;
            }
            catch { }
            Xl.Say("报告已写入工作表「" + (string)sh.Name + "」（可复制到 Word/邮件）。");
            return true;
        }

        // ================= 2. 多表关联建议 =================
        private static bool JoinAdvice()
        {
            AiConfig cfg = AiConfig.Load();
            string key = AiExtOps.ApiKey();
            if (!AiExtOps.Ready(cfg, key)) return true;

            dynamic left = Xl.PickRange("请选择【左表】（含表头）");
            if (left == null) return true;
            dynamic right = Xl.PickRange("请选择【右表】（含表头）");
            if (right == null) return true;

            int lr, lc, rr, rc;
            string addr1, addr2;
            string d1 = AiTasks.Digest(left, cfg, out lr, out lc, out addr1);
            string d2 = AiTasks.Digest(right, cfg, out rr, out rc, out addr2);

            List<AiMessage> msgs = new List<AiMessage>();
            msgs.Add(new AiMessage("system",
                "你是数据建模顾问。判断两张表应该按哪些列关联、用什么连接方式。\n" +
                "只输出 JSON：{\"keyLeft\":\"左表列名\",\"keyRight\":\"右表列名\",\"type\":\"left|inner|right|full\",\"reason\":\"理由\",\"risks\":[\"风险提示\"]}\n" +
                "列名必须来自给定表头。"));
            msgs.Add(new AiMessage("user", "【左表】\n" + d1 + "\n【右表】\n" + d2));

            string err;
            string text = AiExtOps.Call(cfg, key, msgs, "AI 正在分析关联关系", out err);
            if (text == null) { Xl.Bad("请求失败：" + (err ?? "未知错误")); return true; }
            Dictionary<string, object> res = AiTasks.ParseResult(text, out err);
            if (res == null) { Xl.Bad(err); return true; }

            StringBuilder sb = new StringBuilder();
            sb.Append("建议关联键：").Append(Json.Str(res, "keyLeft")).Append("  =  ").Append(Json.Str(res, "keyRight")).Append('\n');
            sb.Append("连接方式：").Append(Json.Str(res, "type")).Append("\n\n");
            sb.Append("理由：").Append(Json.Str(res, "reason")).Append("\n\n");
            List<object> risks = Json.ArrOf(res, "risks");
            if (risks.Count > 0)
            {
                sb.Append("风险提示：\n");
                foreach (object o in risks) sb.Append("  · ").Append(Json.StrOf(o)).Append('\n');
            }
            sb.Append("\n下一步：用「汇总统计 → 多表关联 → 多表关联合并（JOIN）」按上述键执行。");
            AiPreview.Show("AI 关联建议", sb.ToString(), false);
            return true;
        }

        // ================= 3. 整表公式纠错 =================
        private static bool FixAll()
        {
            AiConfig cfg = AiConfig.Load();
            string key = AiExtOps.ApiKey();
            if (!AiExtOps.Ready(cfg, key)) return true;
            dynamic ws = Xl.Act;
            if (ws == null) return true;

            List<string> bad = new List<string>();
            List<string> addrs = new List<string>();
            Xl.Turn(true);
            try
            {
                dynamic ur = ws.UsedRange;
                int rows = (int)ur.Rows.Count, cols = (int)ur.Columns.Count;
                if ((long)rows * cols > 150000) { Xl.Bad("数据过大，请先缩小范围。"); return true; }
                object[,] a = Xl.Vals(ur, out rows, out cols);
                int firstRow = (int)ur.Row, firstCol = (int)ur.Column;
                for (int i = 0; i < rows && bad.Count < 20; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        string s = a[i, j] as string;
                        if (s == null) continue;
                        if (s.Length > 1 && s[0] == '=') continue;   // 正常公式
                        if (s.StartsWith("#") && s.Length <= 12)
                        {
                            bad.Add(s);
                            addrs.Add(Addrs(firstRow + i, firstCol + j));
                        }
                    }
                }
            }
            finally { Xl.Turn(false); }

            if (bad.Count == 0) { Xl.Say("当前表没有错误值，公式健康 ✓"); return true; }

            StringBuilder q = new StringBuilder();
            q.Append("以下单元格显示为错误值，请判断原因并给出修复建议（能修就给完整公式）：\n");
            for (int i = 0; i < bad.Count; i++)
                q.Append(addrs[i]).Append(" → ").Append(bad[i]).Append('\n');
            q.Append("\n只输出 JSON：{\"items\":[{\"cell\":\"地址\",\"cause\":\"原因\",\"fixed\":\"修复后的公式或空\",\"hint\":\"操作建议\"}]}");

            List<AiMessage> msgs = new List<AiMessage>();
            msgs.Add(new AiMessage("system",
                "你是 Excel 公式排错助手。#N/A 多为查找不到、#REF! 为引用失效、#VALUE! 为类型不匹配、#DIV/0! 为除零、#NAME? 为函数名拼错。\n" +
                "只输出 JSON，不要输出解释文字。"));
            msgs.Add(new AiMessage("user", q.ToString()));

            string err;
            string text = AiExtOps.Call(cfg, key, msgs, "AI 正在排查错误", out err);
            if (text == null) { Xl.Bad("请求失败：" + (err ?? "未知错误")); return true; }
            Dictionary<string, object> res = AiTasks.ParseResult(text, out err);
            if (res == null) { Xl.Bad(err); return true; }
            List<object> items = Json.ArrOf(res, "items");
            if (items.Count == 0) { Xl.Say("模型没有给出结论。"); return true; }

            StringBuilder sb = new StringBuilder();
            sb.Append("发现 ").Append(bad.Count).Append(" 个错误值：\n\n");
            foreach (object o in items)
            {
                sb.Append("· ").Append(Json.Str(o, "cell")).Append(" —— ").Append(Json.Str(o, "cause")).Append('\n');
                string fixedF = Json.Str(o, "fixed");
                if (!string.IsNullOrEmpty(fixedF)) sb.Append("   修复：").Append(fixedF).Append('\n');
                string hint = Json.Str(o, "hint");
                if (!string.IsNullOrEmpty(hint)) sb.Append("   建议：").Append(hint).Append('\n');
            }
            sb.Append("\n（本功能只给建议，不会自动改公式；需要改请复制后手动替换或用「公式纠错」逐个处理）");
            AiPreview.Show("AI 整表公式体检", sb.ToString(), false);
            return true;
        }

        private static string Addrs(int row, int col)
        {
            StringBuilder sb = new StringBuilder();
            int n = col;
            while (n > 0) { int m = (n - 1) % 26; sb.Insert(0, (char)('A' + m)); n = (n - 1) / 26; }
            return sb.ToString() + row.ToString(CultureInfo.InvariantCulture);
        }

        // ================= 4. 批量翻译 =================
        private static bool Translate()
        {
            AiConfig cfg = AiConfig.Load();
            string key = AiExtOps.ApiKey();
            if (!AiExtOps.Ready(cfg, key)) return true;
            dynamic r = Xl.SelRange("要翻译的文本列");
            if (r == null) return true;
            string lang = AiExtOps.Prompt("ask", "批量翻译", "目标语言（例：英文 / 日文 / 繁体中文）", "英文");
            if (string.IsNullOrEmpty(lang)) return true;

            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            if (cols != 1) { Xl.Bad("请只选择一列。"); return true; }

            const int batch = 30;
            List<string> result = new List<string>();
            int done = 0;
            while (done < rows)
            {
                int take = Math.Min(batch, rows - done);
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < take; i++) sb.Append(i + 1).Append(". ").Append(Xl.S(a[done + i, 0])).Append('\n');
                List<AiMessage> msgs = new List<AiMessage>();
                msgs.Add(new AiMessage("system",
                    "你是翻译助手。把每条文本翻译成" + lang + "。\n" +
                    "只输出 JSON：{\"values\":[\"译文1\",\"译文2\"]}，长度与顺序必须与输入一致。\n" +
                    "保留数字与专有名词；空行返回空字符串。"));
                msgs.Add(new AiMessage("user", "共 " + take + " 条：\n" + sb));

                string err;
                string text = AiExtOps.Call(cfg, key, msgs, "AI 正在翻译（" + (done + take) + "/" + rows + "）", out err);
                if (text == null) { Xl.Bad("请求失败：" + (err ?? "未知错误")); return true; }
                Dictionary<string, object> res = AiTasks.ParseResult(text, out err);
                if (res == null) { Xl.Bad(err); return true; }
                List<string> got = new List<string>();
                foreach (object o in Json.ArrOf(res, "values")) got.Add(Json.StrOf(o) ?? "");
                for (int i = 0; i < take; i++) result.Add(i < got.Count ? got[i] : "");
                done += take;
            }

            StringBuilder prev = new StringBuilder();
            prev.Append("目标语言：").Append(lang).Append("\n共 ").Append(result.Count).Append(" 条\n\n预览：\n");
            for (int i = 0; i < Math.Min(8, result.Count); i++)
                prev.Append("  ").Append(Xl.S(a[i, 0])).Append("  →  ").Append(result[i]).Append('\n');
            if (!AiPreview.Show("AI 批量翻译", prev.ToString(), true)) return true;

            object[,] outv = new object[result.Count, 1];
            for (int i = 0; i < result.Count; i++) outv[i, 0] = result[i];
            Xl.Turn(true);
            try { TextOps.WriteBlock(r, 0, 1, outv); }
            finally { Xl.Turn(false); }
            Xl.Say("已翻译 " + result.Count + " 条（写入右侧一列）。");
            return true;
        }

        // ================= 5. 合成测试数据 =================
        private static bool Synth()
        {
            AiConfig cfg = AiConfig.Load();
            string key = AiExtOps.ApiKey();
            if (!AiExtOps.Ready(cfg, key)) return true;
            dynamic r = Xl.SelRange("请选择【表头】所在行（只需表头）");
            if (r == null) return true;
            string nStr = AiExtOps.Prompt("ask", "合成测试数据", "生成多少行？", "20");
            int n;
            if (!int.TryParse(nStr, out n) || n < 1 || n > 500) n = 20;

            int rows, cols;
            object[,] a = Xl.Vals(r, out rows, out cols);
            List<string> header = new List<string>();
            for (int j = 0; j < cols; j++) header.Add(Xl.S(a[0, j]));

            List<AiMessage> msgs = new List<AiMessage>();
            msgs.Add(new AiMessage("system",
                "你是测试数据生成器。根据表头生成合规的假数据，不要使用真实个人信息。\n" +
                "只输出 JSON：{\"rows\":[{\"列名\":值}]}，共 " + n + " 行，列名必须与表头一致。\n" +
                "数值列给合理范围的数字，日期列给 yyyy-mm-dd。"));
            msgs.Add(new AiMessage("user", "表头：" + string.Join(",", header.ToArray()) + "\n行数：" + n));

            string err;
            string text = AiExtOps.Call(cfg, key, msgs, "AI 正在生成测试数据", out err);
            if (text == null) { Xl.Bad("请求失败：" + (err ?? "未知错误")); return true; }
            Dictionary<string, object> res = AiTasks.ParseResult(text, out err);
            if (res == null) { Xl.Bad(err); return true; }
            List<object> rowsArr = Json.ArrOf(res, "rows");
            if (rowsArr.Count == 0) { Xl.Bad("模型没有返回数据。"); return true; }

            object[,] data = new object[rowsArr.Count + 1, cols];
            for (int j = 0; j < cols; j++) data[0, j] = header[j];
            for (int i = 0; i < rowsArr.Count; i++)
            {
                Dictionary<string, object> o = Json.Obj(rowsArr[i]);
                for (int j = 0; j < cols; j++) data[i + 1, j] = Json.Str(o, header[j]) ?? "";
            }
            dynamic sh = Xl.AddSheet("合成数据");
            Xl.Dump(sh, 1, 1, data);
            try { sh.Rows[1].Font.Bold = true; sh.UsedRange.Columns.AutoFit(); } catch { }
            Xl.Say("已生成 " + rowsArr.Count + " 行合成数据 → 工作表「" + (string)sh.Name + "」（不含真实个人信息）。");
            return true;
        }

        // ================= 6. 异常解释 =================
        private static bool ExplainOutlier()
        {
            AiConfig cfg = AiConfig.Load();
            string key = AiExtOps.ApiKey();
            if (!AiExtOps.Ready(cfg, key)) return true;
            dynamic r = Xl.SelRange("异常清单或可疑数据区域");
            if (r == null) return true;

            int rows, cols;
            string addr;
            string digest = AiTasks.Digest(r, cfg, out rows, out cols, out addr);
            List<AiMessage> msgs = new List<AiMessage>();
            msgs.Add(new AiMessage("system",
                "你是数据质量顾问。针对表格里可能异常的数据，给出可能原因与核实建议。\n" +
                "输出中文要点，每条不超过两行；区分「业务上合理」与「疑似录入错误」。不要编造看不到的数字。"));
            msgs.Add(new AiMessage("user", digest));

            string err;
            string text = AiExtOps.Call(cfg, key, msgs, "AI 正在分析异常", out err);
            if (text == null) { Xl.Bad("请求失败：" + (err ?? "未知错误")); return true; }
            AiPreview.Show("AI 异常解释（" + addr + "）", text, false);
            return true;
        }
    }
}

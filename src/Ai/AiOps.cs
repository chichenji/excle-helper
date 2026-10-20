using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>AI 操作（ai.*）：设置、连接测试、生成公式、解释公式、纠错、对话面板。</summary>
    internal static class AiOps
    {
        internal static void Register()
        {
            Ops.Reg("ai.settings", Settings);
            Ops.Reg("ai.test", Test);
            Ops.Reg("ai.panel", Panel);
            Ops.Reg("ai.formula", Formula);
            Ops.Reg("ai.explain", Explain);
            Ops.Reg("ai.fix", Fix);
        }

        private static bool Settings() { AiSettingsForm.Open(); return true; }

        /// <summary>取 API Key：自动化测试可用 EXCLE_AI_KEY 覆盖。</summary>
        private static string ApiKey()
        {
            string k = Environment.GetEnvironmentVariable("EXCLE_AI_KEY");
            return string.IsNullOrEmpty(k) ? AiSecret.Load() : k;
        }

        private static bool Test()
        {
            AiConfig cfg = AiConfig.Load();
            string key = ApiKey();
            if (key.Length == 0) { Xl.Bad("还没有配置 API Key，请先执行「AI 设置」。"); return true; }
            string text, err;
            List<AiMessage> msgs = AiTasks.PingTask();
            bool ok = AiRunner.Run("测试 AI 连接", delegate(Func<bool> c) { return AiHttp.Chat(cfg, key, msgs, null, c, out err); },
                                   out text, out err);
            if (ok) Xl.Say("连接正常。\n模型：" + cfg.Model + "\n返回：" + AiHttp.Sanitize(text));
            else Xl.Bad("连接失败：" + (err ?? "无响应"));
            return true;
        }

        private static bool Panel()
        {
            if (Environment.GetEnvironmentVariable("EXCLE_SHOT") == "1") { AiChatPanel.RenderShot(); return true; }
            AiChatPanel.ShowPanel();
            return true;
        }

        // ---------------- 生成公式 ----------------
        private static bool Formula()
        {
            AiConfig cfg = AiConfig.Load();
            string key = ApiKey();
            if (key.Length == 0) { Xl.Bad("还没有配置 API Key，请先执行「AI 设置」。"); return true; }

            dynamic r = Xl.SelRange("要分析的数据区域（建议包含表头）");
            if (r == null) return true;
            // 自动化测试钩子：设置 EXCLE_AI_ASK 时跳过输入框
            string req = AiTasks.FileOpt("ask");
            if (string.IsNullOrEmpty(req)) req = Environment.GetEnvironmentVariable("EXCLE_AI_ASK");
            if (string.IsNullOrEmpty(req))
                req = Ask.One("AI 生成公式", "用一句话描述你要算什么（例：按地区汇总销售额）", "");
            if (req == null || req.Trim().Length == 0) return true;

            int rows, cols;
            string addr;
            string digest = AiTasks.Digest(r, cfg, out rows, out cols, out addr);   // UI 线程内构造
            string suggest = AiTasks.SuggestTarget(r);
            List<AiMessage> msgs = AiTasks.FormulaTask(req.Trim(), digest);

            string text, err;
            bool ok = AiRunner.Run("AI 正在生成公式", delegate(Func<bool> c)
            {
                return AiHttp.Chat(cfg, key, msgs, null, c, out err);
            }, out text, out err);
            if (!ok) { Xl.Bad("请求失败：" + (err ?? "未知错误")); return true; }

            Dictionary<string, object> res = AiTasks.ParseResult(text, out err);
            if (res == null) { Xl.Bad(err); return true; }

            string formula = (Json.Str(res, "formula") ?? "").Trim();
            if (formula.Length == 0)
            {
                Xl.Bad("模型没有给出公式。原始返回：\n" + AiHttp.Sanitize(text));
                return true;
            }
            string risk = AiTasks.CheckFormulaRisk(formula);
            string note;
            bool valid = AiTasks.ValidateFormula(formula, out note);
            string target = Json.Str(res, "target");
            if (string.IsNullOrEmpty(target)) target = suggest;

            StringBuilder sb = new StringBuilder();
            sb.Append("公式：\n").Append(formula).Append("\n\n");
            sb.Append("写入位置：").Append(target).Append("\n\n");
            sb.Append("说明：\n").Append(Json.Str(res, "explanation") ?? "（无）").Append("\n\n");
            List<object> asm = Json.ArrOf(res, "assumptions");
            if (asm.Count > 0)
            {
                sb.Append("模型假设：\n");
                foreach (object a in asm) sb.Append("  · ").Append(Json.StrOf(a)).Append('\n');
                sb.Append('\n');
            }
            sb.Append("置信度：").Append(Json.Num(res, "confidence", 0).ToString("0.00")).Append("\n\n");
            sb.Append("本地校验：").Append(valid ? "公式语法通过" : "未通过").Append(note != null ? "（" + note + "）" : "").Append('\n');
            if (risk != null) sb.Append("风险提示：").Append(risk).Append('\n');
            sb.Append("\n数据区域：").Append(addr).Append("（").Append(rows).Append(" 行 × ").Append(cols).Append(" 列）\n");
            sb.Append(cfg.Redact ? "已开启发送前脱敏\n" : "未开启脱敏\n");

            bool canApply = valid && risk == null;
            if (AiPreview.Show("AI 生成的公式", sb.ToString(), canApply))
            {
                string msg;
                if (AiTasks.ApplyFormula(formula, target, out msg)) Xl.Say(msg + "\n（可用「撤销上次操作」回退）");
                else Xl.Bad(msg);
            }
            return true;
        }

        // ---------------- 解释公式 ----------------
        private static bool Explain()
        {
            AiConfig cfg = AiConfig.Load();
            string key = ApiKey();
            if (key.Length == 0) { Xl.Bad("还没有配置 API Key，请先执行「AI 设置」。"); return true; }

            string f = CurrentFormula();
            if (f == null) { Xl.Bad("请先选中一个包含公式的单元格。"); return true; }

            string digest = SmallDigest(cfg);
            List<AiMessage> msgs = AiTasks.ExplainTask(f, digest);
            string text, err;
            bool ok = AiRunner.Run("AI 正在解释公式", delegate(Func<bool> c) { return AiHttp.Chat(cfg, key, msgs, null, c, out err); },
                                   out text, out err);
            if (!ok) { Xl.Bad("请求失败：" + (err ?? "未知错误")); return true; }

            Dictionary<string, object> res = AiTasks.ParseResult(text, out err);
            if (res == null) { Xl.Bad(err); return true; }

            StringBuilder sb = new StringBuilder();
            sb.Append("公式：\n").Append(f).Append("\n\n");
            sb.Append("解释：\n").Append(Json.Str(res, "explanation") ?? "（无）").Append("\n\n");
            List<object> pit = Json.ArrOf(res, "pitfalls");
            if (pit.Count > 0)
            {
                sb.Append("易错点：\n");
                foreach (object a in pit) sb.Append("  · ").Append(Json.StrOf(a)).Append('\n');
                sb.Append('\n');
            }
            List<object> eq = Json.ArrOf(res, "equivalent");
            if (eq.Count > 0)
            {
                sb.Append("等价写法：\n");
                foreach (object a in eq) sb.Append("  · ").Append(Json.StrOf(a)).Append('\n');
            }
            AiPreview.Show("AI 公式解释", sb.ToString(), false);
            return true;
        }

        // ---------------- 公式纠错 ----------------
        private static bool Fix()
        {
            AiConfig cfg = AiConfig.Load();
            string key = ApiKey();
            if (key.Length == 0) { Xl.Bad("还没有配置 API Key，请先执行「AI 设置」。"); return true; }

            string f = CurrentFormula();
            if (f == null) { Xl.Bad("请先选中一个包含公式的单元格。"); return true; }

            string digest = SmallDigest(cfg);
            List<AiMessage> msgs = AiTasks.FixTask(f, digest);
            string text, err;
            bool ok = AiRunner.Run("AI 正在检查公式", delegate(Func<bool> c) { return AiHttp.Chat(cfg, key, msgs, null, c, out err); },
                                   out text, out err);
            if (!ok) { Xl.Bad("请求失败：" + (err ?? "未知错误")); return true; }

            Dictionary<string, object> res = AiTasks.ParseResult(text, out err);
            if (res == null) { Xl.Bad(err); return true; }

            string fixedFormula = (Json.Str(res, "fixed") ?? "").Trim();
            string risk = AiTasks.CheckFormulaRisk(fixedFormula);
            string note = null;
            bool valid = fixedFormula.Length > 0 && AiTasks.ValidateFormula(fixedFormula, out note);

            StringBuilder sb = new StringBuilder();
            sb.Append("原公式：\n").Append(f).Append("\n\n");
            sb.Append("修正后：\n").Append(fixedFormula.Length > 0 ? fixedFormula : "（模型认为无需修改）").Append("\n\n");
            sb.Append("原因：\n").Append(Json.Str(res, "reason") ?? "（无）").Append("\n\n");
            sb.Append("本地校验：").Append(valid ? "通过" : "未通过").Append(note != null ? "（" + note + "）" : "").Append('\n');
            if (risk != null) sb.Append("风险提示：").Append(risk).Append('\n');

            bool canApply = valid && risk == null && !string.Equals(fixedFormula, f, StringComparison.Ordinal);
            if (AiPreview.Show("AI 公式纠错", sb.ToString(), canApply))
            {
                try
                {
                    dynamic cell = Xl.Sel;
                    cell.Formula = fixedFormula;
                    Xl.Say("已写回修正后的公式。");
                }
                catch (Exception ex) { Xl.Bad("写入失败：" + AiHttp.Sanitize(ex.Message)); }
            }
            return true;
        }

        // ---------------- 工具 ----------------
        private static string CurrentFormula()
        {
            try
            {
                dynamic cell = Xl.Sel;
                if (cell == null) return null;
                if ((bool)cell.HasFormula) return (string)cell.Formula;
            }
            catch { }
            return null;
        }

        /// <summary>解释/纠错时给一个轻量上下文：当前表已用区域的前几行。</summary>
        private static string SmallDigest(AiConfig cfg)
        {
            try
            {
                dynamic ws = Xl.Act;
                dynamic ur = ws.UsedRange;
                int rows, cols;
                string addr;
                return AiTasks.Digest(ur, cfg, out rows, out cols, out addr);
            }
            catch { return null; }
        }
    }
}

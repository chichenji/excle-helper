using System;
using System.Collections.Generic;

namespace ExcleHelper
{
    /// <summary>
    /// Declarative operation registry: ribbon tag -> Func&lt;bool&gt;.
    /// Adding a feature is one Reg() line, not a new WinForms project.
    /// </summary>
    internal static class Ops
    {
        private static readonly Dictionary<string, Func<bool>> Map =
            new Dictionary<string, Func<bool>>(StringComparer.OrdinalIgnoreCase);

        private static bool _ready;

        internal static void Reg(string tag, Func<bool> fn)
        {
            Map[tag] = fn;
        }

        internal static void Init()
        {
            if (_ready) return;
            _ready = true;
            UndoOps.Register();
            AiOps.Register();
            AiExtOps.Register();
            AiAdvOps.Register();
            JoinOps.Register();
            QualityOps.Register();
            PresentOps.Register();
            GovOps.Register();
            Extra2Ops.Register();
            Extra3Ops.Register();
            StatOps.Register();
            SecTplOps.Register();
            PrintOps.Register();
            PicOps.Register();
            DateOps.Register();
            InputOps.Register();
            FormulaOps.Register();
            TextOps.Register();
            NumOps.Register();
            DataOps.Register();
            MatchOps.Register();
            SheetOps.Register();
            CnOps.Register();
            EditOps.Register();
            WsOps.Register();
            SummaryOps.Register();
            FormatOps.Register();
            CellOps.Register();
            ExtrasOps.Register();
            UIOps.Register();
            MiscOps.Register();
        }

        internal static bool Has(string tag)
        {
            Init();
            return Map.ContainsKey(tag);
        }

        internal static bool Run(string tag)
        {
            Init();
            Func<bool> fn;
            if (!Map.TryGetValue(tag, out fn)) return false;
            PresentOps.RecordUse(tag);          // 记录最近使用（面板/收藏夹用）
            Journal.Begin(tag);                 // 执行前快照，供「撤销上次」使用
            bool ok;
            try
            {
                ok = fn();
            }
            catch
            {
                Journal.Abort();
                throw;
            }
            if (ok) Journal.Commit();
            else Journal.Abort();
            return ok;
        }

        internal static int Count
        {
            get { Init(); return Map.Count; }
        }
    }
}

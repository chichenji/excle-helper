using System;

namespace ExcleHelper
{
    /// <summary>界面层操作：功能面板、关于窗口。</summary>
    internal static class UIOps
    {
        internal static void Register()
        {
            Ops.Reg("ui.panel", Panel);
            Ops.Reg("ui.about", About);
            Ops.Reg("ui.shoticons", ShotIcons);
            Ops.Reg("ui.shotai", ShotAi);
        }

        private static bool Shot()
        {
            return Environment.GetEnvironmentVariable("EXCLE_SHOT") == "1";
        }

        private static bool Panel()
        {
            if (!Xl.Ready()) return true;
            if (Shot()) { Launcher.RenderShot(); return true; }
            Launcher.ShowPanel();
            return true;
        }

        /// <summary>离屏渲染三个 AI 弹窗，便于设计评审。</summary>
        private static bool ShotAi()
        {
            AiProgress.RenderShot();
            AiPreview.RenderShot();
            AiSettingsForm.RenderShot();
            Xl.Say("AI 弹窗预览已保存到 %TEMP%：ai-progress / ai-preview / ai-settings.png");
            return true;
        }

        private static bool ShotIcons()
        {
            string p = Icons.RenderSheet();
            Xl.Say("图标总览已保存：\n" + p);
            return true;
        }

        private static bool About()
        {
            if (Shot())
            {
                UiKit.Safe(delegate { AboutForm.RenderShot(); }, "AboutForm");
                UiKit.Safe(delegate { Icons.RenderSheet(); }, "Icons");
                UiKit.Safe(delegate { QrCode.RenderDemo(); }, "QrCode");
                UiKit.Safe(delegate { AiProgress.RenderShot(); }, "AiProgress");
                UiKit.Safe(delegate { AiPreview.RenderShot(); }, "AiPreview");
                UiKit.Safe(delegate { AiSettingsForm.RenderShot(); }, "AiSettingsForm");
                UiKit.Safe(delegate { AiChatPanel.RenderShot(); }, "AiChatPanel");
                return true;
            }
            AboutForm.ShowAbout();
            return true;
        }
    }
}

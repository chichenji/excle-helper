using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcleHelper
{
    /// <summary>请求中的进度窗（品牌化，可取消）。</summary>
    internal sealed class AiProgress : BrandForm
    {
        internal event EventHandler CancelRequested;

        private readonly Timer _t = new Timer();
        private int _dots;
        private int _phase;
        private readonly string _stage;
        private FlatButton _cancel;

        internal AiProgress(string title)
        {
            HeadTitle = "AI 正在处理";
            HeadSubtitle = title;
            HeadGlyph = "✦";
            AllowClose = false;
            ClientSize = new Size(440, 206);

            Label msg = new Label();
            msg.Text = "正在请求模型，请稍候…";
            msg.SetBounds(24, 88, 392, 24);
            msg.Font = UiKit.FontOf(10f, FontStyle.Regular);
            msg.ForeColor = UiKit.Ink;
            Controls.Add(msg);
            _stage = "";

            Label tip = new Label();
            tip.Text = "可随时取消；结果会先预览，确认后才写入。";
            tip.SetBounds(24, 114, 392, 20);
            tip.ForeColor = UiKit.Muted;
            tip.Font = UiKit.FontOf(8.5f, FontStyle.Regular);
            Controls.Add(tip);

            _cancel = new FlatButton("取消请求", FlatButton.Kind.Secondary);
            _cancel.SetBounds(324, 162, 96, 30);
            _cancel.Click += delegate
            {
                _t.Stop();
                EventHandler h = CancelRequested;
                if (h != null) h(this, EventArgs.Empty);
                _cancel.Text = "正在取消";
                _cancel.Enabled = false;
                Invalidate();
            };
            Controls.Add(_cancel);

            _t.Interval = 260;
            _t.Tick += delegate { _dots = (_dots + 1) % 4; _phase = (_phase + 3) % 100; Invalidate(); };
            _t.Start();
            FormClosed += delegate { _t.Stop(); };
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            // 进度条：细长轨道 + 移动的高亮段
            Rectangle track = new Rectangle(24, 146, 392, 4);
            using (GraphicsPath p = UiKit.Rounded(track, 2))
            using (SolidBrush b = new SolidBrush(Color.FromArgb(226, 232, 240)))
                g.FillPath(b, p);

            int segW = 120;
            int x = track.X + (_phase * (track.Width + segW) / 100) - segW;
            Rectangle seg = new Rectangle(Math.Max(track.X, x), track.Y,
                                          Math.Min(segW, track.Right - Math.Max(track.X, x)), track.Height);
            if (seg.Width > 4)
            {
                using (GraphicsPath p = UiKit.Rounded(seg, 2))
                using (LinearGradientBrush b = new LinearGradientBrush(seg,
                           Color.FromArgb(147, 197, 253), UiKit.BrandBlue, 0f))
                    g.FillPath(b, p);
            }

            // 省略号动画
            using (Font f = UiKit.FontOf(10f, FontStyle.Regular))
            using (SolidBrush b = new SolidBrush(UiKit.BrandBlue))
                g.DrawString(new string('.', _dots), f, b, 234, 89);
        }

        internal static void RenderShot()
        {
            AiProgress f = new AiProgress("AI 正在生成公式");
            f.StartPosition = FormStartPosition.Manual;
            f.Location = new Point(-4000, -4000);
            f.Show();
            Application.DoEvents();
            using (Bitmap bmp = new Bitmap(f.Width, f.Height))
            {
                f.DrawToBitmap(bmp, new Rectangle(0, 0, f.Width, f.Height));
                bmp.Save(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "exclehelper-ai-progress.png"),
                         System.Drawing.Imaging.ImageFormat.Png);
            }
            f.Close();
        }
    }

    /// <summary>AI 结果预览：品牌化窗口，等宽正文，先看后落盘。</summary>
    internal sealed class AiPreview : BrandForm
    {
        internal bool Applied;
        private FieldBox _body;

        internal static bool Show(string title, string text, bool withApply)
        {
            if (Environment.GetEnvironmentVariable("EXCLE_AI_AUTOAPPLY") == "1") return withApply;
            using (AiPreview f = new AiPreview(title, text, withApply))
            {
                f.ShowDialog();
                return f.Applied;
            }
        }

        private AiPreview(string title, string text, bool withApply)
        {
            HeadTitle = title;
            HeadSubtitle = "AI 结果 · 确认后才会写入表格";
            HeadGlyph = "✦";
            ClientSize = new Size(720, 504);

            _body = new FieldBox(text, true);
            _body.SetBounds(20, 84, 680, 340);
            _body.Inner.ReadOnly = true;
            _body.Inner.Font = new Font("Consolas", 9.5f);
            _body.Inner.BackColor = Color.White;
            _body.Inner.WordWrap = false;
            _body.Inner.ScrollBars = ScrollBars.Both;
            Controls.Add(_body);

            Label hint = new Label();
            hint.Text = withApply ? "写入后可用「撤销上次操作」回退。" : "本结果不会修改表格。";
            hint.SetBounds(22, 434, 360, 20);
            hint.ForeColor = UiKit.Muted;
            hint.Font = UiKit.FontOf(8.5f, FontStyle.Regular);
            Controls.Add(hint);

            FlatButton copy = new FlatButton("复制", FlatButton.Kind.Ghost);
            copy.SetBounds(424, 456, 76, 32);
            copy.Click += delegate { try { Clipboard.SetText(_body.Inner.Text); copy.Text = "已复制"; } catch { } };
            Controls.Add(copy);

            FlatButton cancel = new FlatButton(withApply ? "取消" : "关闭", FlatButton.Kind.Secondary);
            cancel.SetBounds(508, 456, 80, 32);
            cancel.Click += delegate { Close(); };
            Controls.Add(cancel);

            if (withApply)
            {
                FlatButton apply = new FlatButton("应用到表格", FlatButton.Kind.Primary);
                apply.SetBounds(596, 456, 104, 32);
                apply.Click += delegate { Applied = true; Close(); };
                Controls.Add(apply);
            }

            KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.Escape) Close(); };
            Load += delegate { _body.Inner.Focus(); _body.Inner.SelectionStart = 0; };
        }

        internal static void RenderShot()
        {
            string sample =
                "公式：\r\n" +
                "=SUMIFS(D:D,B:B,\"华东\",A:A,\">=\"&DATE(2024,1,1))\r\n\r\n" +
                "写入位置：F2\r\n\r\n" +
                "说明：\r\n" +
                "统计 B 列为「华东」且 A 列日期不早于 2024-01-01 的销售额合计。\r\n\r\n" +
                "模型假设：\r\n" +
                "  · B 列是地区\r\n" +
                "  · D 列是销售额\r\n\r\n" +
                "置信度：0.92\r\n\r\n" +
                "本地校验：公式语法通过\r\n" +
                "数据区域：A1:D51（50 行 × 4 列）\r\n" +
                "已开启发送前脱敏";
            AiPreview f = new AiPreview("AI 生成的公式", sample, true);
            f.StartPosition = FormStartPosition.Manual;
            f.Location = new Point(-4000, -4000);
            f.Show();
            Application.DoEvents();
            using (Bitmap bmp = new Bitmap(f.Width, f.Height))
            {
                f.DrawToBitmap(bmp, new Rectangle(0, 0, f.Width, f.Height));
                bmp.Save(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "exclehelper-ai-preview.png"),
                         System.Drawing.Imaging.ImageFormat.Png);
            }
            f.Close();
        }
    }

    /// <summary>AI 设置：分区布局 + 品牌化控件。</summary>
    internal sealed class AiSettingsForm : BrandForm
    {
        private ComboBox _provider;
        private FieldBox _url, _model, _key, _temp, _maxTok, _timeout, _rows;
        private CheckBox _stream, _jsonMode, _redact;
        private Label _state;
        private bool _suppressPreset;

        internal static void Open()
        {
            using (AiSettingsForm f = new AiSettingsForm())
                f.ShowDialog();
        }

        private AiSettingsForm()
        {
            HeadTitle = "AI 设置";
            HeadSubtitle = "供应商 · 密钥（DPAPI 加密保存）· 上下文与隐私";
            HeadGlyph = "✦";
            ClientSize = new Size(660, 520);
            AiConfig c = AiConfig.Load();

            Section("连接", 20, 84);

            Label l1 = Field("供应商", 20, 116);
            _provider = new ComboBox();
            _provider.DropDownStyle = ComboBoxStyle.DropDownList;
            _provider.FlatStyle = FlatStyle.Flat;
            _provider.Font = UiKit.FontOf(9.5f, FontStyle.Regular);
            _provider.SetBounds(104, 114, 150, 24);
            _provider.Items.AddRange(new object[] { "DeepSeek", "OpenAI 兼容", "本地 Ollama", "自定义" });
            _provider.SelectedIndex = 0;
            _provider.SelectedIndexChanged += delegate { ApplyPreset(); };
            Controls.Add(_provider);

            Label l2 = Field("模型名", 276, 116);
            _model = Box(c.Model, 356, 112, 284);

            Label l3 = Field("接口地址", 20, 156);
            _url = Box(c.BaseUrl, 104, 152, 536);

            Label l4 = Field("API Key", 20, 196);
            _key = Box(AiSecret.Load(), 104, 192, 536);
            _key.Inner.UseSystemPasswordChar = true;

            Section("参数", 20, 244);
            Label l5 = Field("温度", 20, 276);
            _temp = Box(c.Temperature.ToString("0.##"), 104, 272, 70);
            Label l6 = Field("最大输出", 190, 276);
            _maxTok = Box(c.MaxTokens.ToString(), 274, 272, 70);
            Label l7 = Field("超时(秒)", 360, 276);
            _timeout = Box(c.TimeoutSeconds.ToString(), 444, 272, 70);
            Label l8 = Field("样本行数", 500, 276);
            _rows = Box(c.ContextRows.ToString(), 580, 272, 60);

            _stream = Check("流式输出（首字更快）", 104, 312, c.Streaming);
            _jsonMode = Check("要求 JSON 输出（不支持时自动降级）", 104, 338, c.JsonMode);
            _redact = Check("发送前脱敏（手机号 / 身份证 / 银行卡 / 邮箱 / 姓名）", 104, 364, c.Redact);

            _state = new Label();
            _state.SetBounds(20, 430, 400, 22);
            _state.ForeColor = UiKit.Muted;
            _state.Font = UiKit.FontOf(8.5f, FontStyle.Regular);
            _state.Text = AiSecret.Has ? "已保存 API Key（DPAPI 加密，仅本机可解）" : "尚未保存 API Key";
            Controls.Add(_state);

            FlatButton test = new FlatButton("测试连接", FlatButton.Kind.Ghost);
            test.SetBounds(300, 460, 90, 34);
            test.Click += delegate { TestConnection(); };
            Controls.Add(test);

            FlatButton cancel = new FlatButton("取消", FlatButton.Kind.Secondary);
            cancel.SetBounds(456, 460, 82, 34);
            cancel.Click += delegate { Close(); };
            Controls.Add(cancel);

            FlatButton save = new FlatButton("保存", FlatButton.Kind.Primary);
            save.SetBounds(546, 460, 94, 34);
            save.Click += delegate { Save(); };
            Controls.Add(save);

            Load += delegate { ApplyProviderGuess(); };
        }

        private void Section(string title, int x, int y)
        {
            Label l = new Label();
            l.Text = title;
            l.SetBounds(x, y, 200, 22);
            l.Font = UiKit.FontOf(10f, FontStyle.Bold);
            l.ForeColor = UiKit.BrandDark;
            Controls.Add(l);
            Label bar = new Label();
            bar.SetBounds(x, y + 22, 620, 1);
            bar.BackColor = UiKit.Line;
            Controls.Add(bar);
        }

        private Label Field(string text, int x, int y)
        {
            Label l = new Label();
            l.Text = text;
            l.SetBounds(x, y + 3, 90, 20);
            l.ForeColor = Color.FromArgb(51, 65, 85);
            Controls.Add(l);
            return l;
        }

        private FieldBox Box(string value, int x, int y, int w)
        {
            FieldBox b = new FieldBox(value, false);
            b.SetBounds(x, y, w, 30);
            Controls.Add(b);
            return b;
        }

        private CheckBox Check(string text, int x, int y, bool v)
        {
            CheckBox c = new CheckBox();
            c.Text = text;
            c.SetBounds(x, y, 520, 22);
            c.Checked = v;
            c.ForeColor = Color.FromArgb(51, 65, 85);
            Controls.Add(c);
            return c;
        }

        private void ApplyProviderGuess()
        {
            _suppressPreset = true;
            try
            {
                string u = (_url.Inner.Text ?? "").ToLowerInvariant();
                if (u.Contains("deepseek")) _provider.SelectedIndex = 0;
                else if (u.Contains("localhost") || u.Contains("127.0.0.1")) _provider.SelectedIndex = 2;
                else if (u.Contains("openai")) _provider.SelectedIndex = 1;
                else _provider.SelectedIndex = 3;
            }
            finally { _suppressPreset = false; }
        }

        private void ApplyPreset()
        {
            if (_suppressPreset) return;
            if (_provider.SelectedIndex == 0)
            {
                _url.Inner.Text = "https://api.deepseek.com/v1";
                _model.Inner.Text = "deepseek-chat";
            }
            else if (_provider.SelectedIndex == 2)
            {
                _url.Inner.Text = "http://localhost:11434/v1";
                if (_model.Inner.Text.Trim().Length == 0 || _model.Inner.Text.Contains("deepseek"))
                    _model.Inner.Text = "qwen2.5:7b";
            }
        }

        private AiConfig Collect()
        {
            AiConfig c = new AiConfig();
            c.Provider = _provider.SelectedIndex == 0 ? "deepseek"
                       : _provider.SelectedIndex == 1 ? "openai"
                       : _provider.SelectedIndex == 2 ? "ollama" : "custom";
            c.BaseUrl = _url.Inner.Text.Trim();
            c.Model = _model.Inner.Text.Trim();
            double d;
            if (double.TryParse(_temp.Inner.Text.Trim(), out d)) c.Temperature = d;
            int n;
            if (int.TryParse(_maxTok.Inner.Text.Trim(), out n) && n > 0) c.MaxTokens = n;
            if (int.TryParse(_timeout.Inner.Text.Trim(), out n) && n >= 5) c.TimeoutSeconds = n;
            if (int.TryParse(_rows.Inner.Text.Trim(), out n) && n >= 0 && n <= 200) c.ContextRows = n;
            c.Streaming = _stream.Checked;
            c.JsonMode = _jsonMode.Checked;
            c.Redact = _redact.Checked;
            return c;
        }

        private void Save()
        {
            AiConfig c = Collect();
            if (c.Model.Length == 0) { _state.Text = "模型名不能为空。"; return; }
            c.Save();
            AiSecret.Save(_key.Inner.Text.Trim());
            Xl.Say("AI 设置已保存。\n配置目录：" + AiConfig.Dir);
            Close();
        }

        private void TestConnection()
        {
            AiConfig c = Collect();
            string key = _key.Inner.Text.Trim();
            _state.Text = "正在测试连接…";
            string text, err;
            List<AiMessage> msgs = AiTasks.PingTask();
            bool ok = AiRunner.Run("测试 AI 连接", delegate(Func<bool> cancelled)
            {
                return AiHttp.Chat(c, key, msgs, null, cancelled, out err);
            }, out text, out err);
            _state.Text = ok && !string.IsNullOrEmpty(text)
                ? "连接成功，模型返回：" + (text.Length > 60 ? text.Substring(0, 60) : text)
                : "连接失败：" + (err ?? "无响应");
        }

        internal static void RenderShot()
        {
            AiSettingsForm f = new AiSettingsForm();
            f.StartPosition = FormStartPosition.Manual;
            f.Location = new Point(-4000, -4000);
            f.Show();
            Application.DoEvents();
            using (Bitmap bmp = new Bitmap(f.Width, f.Height))
            {
                f.DrawToBitmap(bmp, new Rectangle(0, 0, f.Width, f.Height));
                bmp.Save(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "exclehelper-ai-settings.png"),
                         System.Drawing.Imaging.ImageFormat.Png);
            }
            f.Close();
        }
    }
}

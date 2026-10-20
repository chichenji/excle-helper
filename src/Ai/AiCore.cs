using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;

namespace ExcleHelper
{
    /// <summary>AI 设置（存 %APPDATA%\ExcleHelper\ai.ini，纯文本键值，不含密钥）。</summary>
    internal sealed class AiConfig
    {
        internal string Provider = "deepseek";
        internal string BaseUrl = "https://api.deepseek.com/v1";
        internal string Model = "deepseek-chat";
        internal double Temperature = 0.2;
        internal int MaxTokens = 2048;
        internal int TimeoutSeconds = 60;
        internal bool Streaming = true;
        internal bool JsonMode = true;        // 请求 response_format=json_object
        internal bool Redact = true;          // 发送前脱敏敏感列
        internal int ContextRows = 5;         // 摘要模式样本行数

        internal static string Dir
        {
            get
            {
                string d = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcleHelper");
                if (!Directory.Exists(d)) Directory.CreateDirectory(d);
                return d;
            }
        }

        private static string FilePath { get { return Path.Combine(Dir, "ai.ini"); } }

        internal static AiConfig Load()
        {
            AiConfig c = new AiConfig();
            try
            {
                if (!File.Exists(FilePath)) return c;
                foreach (string line in File.ReadAllLines(FilePath, Encoding.UTF8))
                {
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    string k = line.Substring(0, eq).Trim();
                    string v = line.Substring(eq + 1).Trim();
                    switch (k)
                    {
                        case "provider": c.Provider = v; break;
                        case "baseurl": c.BaseUrl = v; break;
                        case "model": c.Model = v; break;
                        case "temperature": c.Temperature = ToD(v, c.Temperature); break;
                        case "maxtokens": c.MaxTokens = (int)ToD(v, c.MaxTokens); break;
                        case "timeout": c.TimeoutSeconds = (int)ToD(v, c.TimeoutSeconds); break;
                        case "streaming": c.Streaming = v == "1"; break;
                        case "jsonmode": c.JsonMode = v == "1"; break;
                        case "redact": c.Redact = v == "1"; break;
                        case "contextrows": c.ContextRows = (int)ToD(v, c.ContextRows); break;
                    }
                }
            }
            catch { }
            return c;
        }

        private static double ToD(string s, double def)
        {
            double d;
            return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d) ? d : def;
        }

        internal void Save()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("provider=" + Provider);
            sb.AppendLine("baseurl=" + BaseUrl);
            sb.AppendLine("model=" + Model);
            sb.AppendLine("temperature=" + Temperature.ToString("0.##", CultureInfo.InvariantCulture));
            sb.AppendLine("maxtokens=" + MaxTokens.ToString(CultureInfo.InvariantCulture));
            sb.AppendLine("timeout=" + TimeoutSeconds.ToString(CultureInfo.InvariantCulture));
            sb.AppendLine("streaming=" + (Streaming ? "1" : "0"));
            sb.AppendLine("jsonmode=" + (JsonMode ? "1" : "0"));
            sb.AppendLine("redact=" + (Redact ? "1" : "0"));
            sb.AppendLine("contextrows=" + ContextRows.ToString(CultureInfo.InvariantCulture));
            File.WriteAllText(FilePath, sb.ToString(), Encoding.UTF8);
        }

        internal string NormalizedBase
        {
            get
            {
                string b = (BaseUrl ?? "").Trim().TrimEnd('/');
                if (b.Length == 0) b = "https://api.deepseek.com/v1";
                if (!b.EndsWith("/chat/completions", StringComparison.OrdinalIgnoreCase))
                    b = b + "/chat/completions";
                return b;
            }
        }
    }

    /// <summary>API Key 用 DPAPI（当前用户作用域）加密保存，明文只在内存中短暂存在。</summary>
    internal static class AiSecret
    {
        private static string FilePath { get { return Path.Combine(AiConfig.Dir, "ai.bin"); } }

        internal static string Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return "";
                byte[] blob = File.ReadAllBytes(FilePath);
                byte[] plain = ProtectedData.Unprotect(blob, null, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(plain);
            }
            catch { return ""; }
        }

        internal static bool Save(string key)
        {
            try
            {
                if (string.IsNullOrEmpty(key))
                {
                    if (File.Exists(FilePath)) File.Delete(FilePath);
                    return true;
                }
                byte[] blob = ProtectedData.Protect(Encoding.UTF8.GetBytes(key), null, DataProtectionScope.CurrentUser);
                File.WriteAllBytes(FilePath, blob);
                return true;
            }
            catch { return false; }
        }

        internal static bool Has { get { return Load().Length > 0; } }
    }

    internal sealed class AiMessage
    {
        internal string Role;
        internal string Content;
        internal AiMessage(string role, string content) { Role = role; Content = content; }
    }

    /// <summary>OpenAI 兼容协议的调用客户端（DeepSeek / 自定义 / 本地 Ollama 通用）。</summary>
    internal static class AiHttp
    {
        private static bool _tlsReady;

        internal static void EnsureTls()
        {
            if (_tlsReady) return;
            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            }
            catch { }
            _tlsReady = true;
        }

        /// <summary>把 API Key、手机号、身份证、银行卡从文本中抹掉，便于安全地展示错误信息。</summary>
        internal static string Sanitize(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            string key = AiSecret.Load();
            if (key.Length > 8) s = s.Replace(key, "***");
            s = System.Text.RegularExpressions.Regex.Replace(s, @"sk-[A-Za-z0-9_\-]{8,}", "sk-***");
            s = System.Text.RegularExpressions.Regex.Replace(s, @"\b1[3-9]\d{9}\b", "1**********");
            s = System.Text.RegularExpressions.Regex.Replace(s, @"\b\d{17}[\dXx]\b", "******************");
            if (s.Length > 600) s = s.Substring(0, 600) + "…";
            return s;
        }

        /// <summary>
        /// 调用 /chat/completions。返回完整文本；流式时逐段回调 onDelta。
        /// 全部网络操作都在调用方线程（由调用方放到后台线程执行）。
        /// </summary>
        internal static string Chat(AiConfig cfg, string apiKey, List<AiMessage> messages,
                                    Action<string> onDelta, Func<bool> cancelled, out string error)
        {
            error = null;
            EnsureTls();

            bool jsonMode = cfg.JsonMode;
            string body = BuildBody(cfg, messages, cfg.Streaming, jsonMode);
            string text;
            string err;
            if (TryPost(cfg, apiKey, body, onDelta, cancelled, out text, out err)) return text;

            // JSON 模式不被支持时自动降级重试一次
            if (jsonMode && err != null && err.IndexOf("response_format", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                jsonMode = false;
                body = BuildBody(cfg, messages, cfg.Streaming, false);
                if (TryPost(cfg, apiKey, body, onDelta, cancelled, out text, out err)) return text;
            }
            error = err;
            return null;
        }

        private static string BuildBody(AiConfig cfg, List<AiMessage> messages, bool stream, bool jsonMode)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("{\"model\":").Append(Json.Quote(cfg.Model));
            sb.Append(",\"messages\":[");
            for (int i = 0; i < messages.Count; i++)
            {
                if (i > 0) sb.Append(',');
                sb.Append("{\"role\":").Append(Json.Quote(messages[i].Role))
                  .Append(",\"content\":").Append(Json.Quote(messages[i].Content)).Append('}');
            }
            sb.Append(']');
            sb.Append(",\"temperature\":").Append(cfg.Temperature.ToString("0.##", CultureInfo.InvariantCulture));
            sb.Append(",\"max_tokens\":").Append(cfg.MaxTokens.ToString(CultureInfo.InvariantCulture));
            sb.Append(",\"stream\":").Append(stream ? "true" : "false");
            if (jsonMode) sb.Append(",\"response_format\":{\"type\":\"json_object\"}");
            sb.Append('}');
            return sb.ToString();
        }

        private static bool TryPost(AiConfig cfg, string apiKey, string body,
                                    Action<string> onDelta, Func<bool> cancelled,
                                    out string text, out string error)
        {
            text = null; error = null;
            HttpWebRequest req = null;
            try
            {
                req = (HttpWebRequest)WebRequest.Create(cfg.NormalizedBase);
                req.Method = "POST";
                req.ContentType = "application/json";
                req.Accept = cfg.Streaming ? "text/event-stream" : "application/json";
                req.Timeout = Math.Max(5, cfg.TimeoutSeconds) * 1000;
                req.ReadWriteTimeout = Math.Max(10, cfg.TimeoutSeconds * 2) * 1000;
                req.UserAgent = "ExcleHelper/1.1";
                if (!string.IsNullOrEmpty(apiKey))
                    req.Headers["Authorization"] = "Bearer " + apiKey;

                byte[] data = Encoding.UTF8.GetBytes(body);
                req.ContentLength = data.Length;
                using (Stream rs = req.GetRequestStream()) rs.Write(data, 0, data.Length);

                using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                using (Stream stream = resp.GetResponseStream())
                using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                {
                    if (cfg.Streaming)
                    {
                        StringBuilder all = new StringBuilder();
                        string line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            if (cancelled != null && cancelled()) { error = "已取消"; return false; }
                            if (line.Length == 0) continue;
                            if (!line.StartsWith("data:", StringComparison.Ordinal)) continue;
                            string payload = line.Substring(5).Trim();
                            if (payload == "[DONE]") break;
                            object o = Json.Parse(payload);
                            List<object> choices = Json.ArrOf(o, "choices");
                            if (choices.Count == 0) continue;
                            object delta = Json.Obj(choices[0]);
                            Dictionary<string, object> d0 = Json.Obj(choices[0]);
                            object dObj = null;
                            if (d0 != null) d0.TryGetValue("delta", out dObj);
                            string piece = Json.Str(dObj, "content");
                            if (string.IsNullOrEmpty(piece))
                            {
                                // 少数实现把增量放在 message.content
                                object mObj = null;
                                if (d0 != null) d0.TryGetValue("message", out mObj);
                                piece = Json.Str(mObj, "content");
                            }
                            if (!string.IsNullOrEmpty(piece))
                            {
                                all.Append(piece);
                                if (onDelta != null) onDelta(piece);
                            }
                            else if (delta == null) { }
                        }
                        text = all.ToString();
                        return true;
                    }

                    string raw = reader.ReadToEnd();
                    object parsed = Json.Parse(raw);
                    List<object> ch = Json.ArrOf(parsed, "choices");
                    if (ch.Count == 0)
                    {
                        error = "响应中没有 choices：" + Sanitize(raw);
                        return false;
                    }
                    Dictionary<string, object> c0 = Json.Obj(ch[0]);
                    object msg = null;
                    if (c0 != null) c0.TryGetValue("message", out msg);
                    text = Json.Str(msg, "content");
                    if (text == null) text = Json.Str(c0, "text");
                    if (text == null) { error = "响应格式无法识别：" + Sanitize(raw); return false; }
                    return true;
                }
            }
            catch (WebException we)
            {
                string detail = "";
                try
                {
                    if (we.Response != null)
                    {
                        using (Stream s = we.Response.GetResponseStream())
                        using (StreamReader r = new StreamReader(s, Encoding.UTF8))
                            detail = r.ReadToEnd();
                    }
                }
                catch { }
                error = "网络错误：" + we.Message + (detail.Length > 0 ? " | " + Sanitize(detail) : "");
                return false;
            }
            catch (Exception ex)
            {
                error = "请求失败：" + Sanitize(ex.Message);
                return false;
            }
            finally
            {
                try { if (req != null) req.Abort(); } catch { }
            }
        }
    }
}

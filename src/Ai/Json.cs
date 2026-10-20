using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcleHelper
{
    /// <summary>
    /// 极简 JSON 解析/构造（自写，保持"零第三方依赖"）。
    /// Parse 返回：Dictionary&lt;string,object&gt; / List&lt;object&gt; / string / double / bool / null。
    /// </summary>
    internal static class Json
    {
        // ---------------- 解析 ----------------
        internal static object Parse(string s)
        {
            if (string.IsNullOrEmpty(s)) return null;
            int i = 0;
            object v = ParseValue(s, ref i);
            return v;
        }

        private static object ParseValue(string s, ref int i)
        {
            SkipWs(s, ref i);
            if (i >= s.Length) return null;
            char c = s[i];
            if (c == '{') return ParseObject(s, ref i);
            if (c == '[') return ParseArray(s, ref i);
            if (c == '"') return ParseString(s, ref i);
            if (c == 't') { i += 4; return true; }
            if (c == 'f') { i += 5; return false; }
            if (c == 'n') { i += 4; return null; }
            return ParseNumber(s, ref i);
        }

        private static Dictionary<string, object> ParseObject(string s, ref int i)
        {
            Dictionary<string, object> d = new Dictionary<string, object>(StringComparer.Ordinal);
            i++;   // {
            while (true)
            {
                SkipWs(s, ref i);
                if (i >= s.Length) break;
                if (s[i] == '}') { i++; break; }
                if (s[i] == ',') { i++; continue; }
                string key = ParseString(s, ref i);
                SkipWs(s, ref i);
                if (i < s.Length && s[i] == ':') i++;
                object val = ParseValue(s, ref i);
                d[key] = val;
            }
            return d;
        }

        private static List<object> ParseArray(string s, ref int i)
        {
            List<object> list = new List<object>();
            i++;   // [
            while (true)
            {
                SkipWs(s, ref i);
                if (i >= s.Length) break;
                if (s[i] == ']') { i++; break; }
                if (s[i] == ',') { i++; continue; }
                list.Add(ParseValue(s, ref i));
            }
            return list;
        }

        private static string ParseString(string s, ref int i)
        {
            StringBuilder sb = new StringBuilder();
            if (i < s.Length && s[i] == '"') i++;
            while (i < s.Length)
            {
                char c = s[i++];
                if (c == '"') break;
                if (c != '\\') { sb.Append(c); continue; }
                if (i >= s.Length) break;
                char e = s[i++];
                switch (e)
                {
                    case 'n': sb.Append('\n'); break;
                    case 't': sb.Append('\t'); break;
                    case 'r': sb.Append('\r'); break;
                    case 'b': sb.Append('\b'); break;
                    case 'f': sb.Append('\f'); break;
                    case 'u':
                        if (i + 4 <= s.Length)
                        {
                            int code;
                            if (int.TryParse(s.Substring(i, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out code))
                                sb.Append((char)code);
                            i += 4;
                        }
                        break;
                    default: sb.Append(e); break;
                }
            }
            return sb.ToString();
        }

        private static object ParseNumber(string s, ref int i)
        {
            int start = i;
            while (i < s.Length && (char.IsDigit(s[i]) || s[i] == '-' || s[i] == '+' || s[i] == '.' || s[i] == 'e' || s[i] == 'E'))
                i++;
            double d;
            if (double.TryParse(s.Substring(start, i - start), NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                return d;
            return null;
        }

        private static void SkipWs(string s, ref int i)
        {
            while (i < s.Length && (s[i] == ' ' || s[i] == '\t' || s[i] == '\n' || s[i] == '\r')) i++;
        }

        // ---------------- 取值 ----------------
        internal static Dictionary<string, object> Obj(object o)
        {
            return o as Dictionary<string, object>;
        }

        internal static List<object> Arr(object o)
        {
            return o as List<object>;
        }

        internal static string Str(object container, string key)
        {
            Dictionary<string, object> d = container as Dictionary<string, object>;
            if (d == null) return null;
            object v;
            if (!d.TryGetValue(key, out v) || v == null) return null;
            if (v is string) return (string)v;
            if (v is double) return ((double)v).ToString(CultureInfo.InvariantCulture);
            if (v is bool) return ((bool)v) ? "true" : "false";
            return v.ToString();
        }

        internal static string StrOf(object v)
        {
            if (v == null) return null;
            if (v is string) return (string)v;
            if (v is double) return ((double)v).ToString(CultureInfo.InvariantCulture);
            if (v is bool) return ((bool)v) ? "true" : "false";
            return v.ToString();
        }

        internal static double Num(object container, string key, double def)
        {
            Dictionary<string, object> d = container as Dictionary<string, object>;
            object v;
            if (d == null || !d.TryGetValue(key, out v) || v == null) return def;
            if (v is double) return (double)v;
            double r;
            if (double.TryParse(StrOf(v), NumberStyles.Any, CultureInfo.InvariantCulture, out r)) return r;
            return def;
        }

        internal static List<object> ArrOf(object container, string key)
        {
            Dictionary<string, object> d = container as Dictionary<string, object>;
            object v;
            if (d == null || !d.TryGetValue(key, out v)) return new List<object>();
            List<object> l = v as List<object>;
            return l ?? new List<object>();
        }

        /// <summary>从可能带 ```json 围栏的文本里提取 JSON 主体。</summary>
        internal static string ExtractJson(string text)
        {
            if (string.IsNullOrEmpty(text)) return null;
            string s = text.Trim();
            int fence = s.IndexOf("```", StringComparison.Ordinal);
            if (fence >= 0)
            {
                int nl = s.IndexOf('\n', fence);
                if (nl > 0)
                {
                    int end = s.IndexOf("```", nl, StringComparison.Ordinal);
                    if (end > nl) s = s.Substring(nl + 1, end - nl - 1).Trim();
                }
            }
            int a = s.IndexOf('{');
            int b = s.LastIndexOf('}');
            if (a >= 0 && b > a) return s.Substring(a, b - a + 1);
            return s;
        }

        // ---------------- 构造 ----------------
        internal static string Quote(string s)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append('"');
            if (s != null)
            {
                foreach (char c in s)
                {
                    switch (c)
                    {
                        case '"': sb.Append("\\\""); break;
                        case '\\': sb.Append("\\\\"); break;
                        case '\n': sb.Append("\\n"); break;
                        case '\r': sb.Append("\\r"); break;
                        case '\t': sb.Append("\\t"); break;
                        default:
                            if (c < ' ') sb.Append("\\u").Append(((int)c).ToString("x4", CultureInfo.InvariantCulture));
                            else sb.Append(c);
                            break;
                    }
                }
            }
            sb.Append('"');
            return sb.ToString();
        }

        internal static string NumStr(double d)
        {
            return d.ToString("0.####", CultureInfo.InvariantCulture);
        }
    }
}

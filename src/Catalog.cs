using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;

namespace ExcleHelper
{
    internal sealed class Feature
    {
        internal string Id;
        internal string Tag;
        internal string Label;
        internal string Hint;
        internal string Tab;
        internal string Group;
    }

    /// <summary>
    /// 功能目录：直接解析内嵌 ribbon.xml 得到 tab/group/label/tag，
    /// 因此面板与功能区永远同源，不会出现「界面有、注册表没有」的脱节。
    /// </summary>
    internal static class Catalog
    {
        private static List<Feature> _all;
        private static List<string> _categories;
        private static Dictionary<string, Feature> _byId;

        internal static List<Feature> All
        {
            get { if (_all == null) Load(); return _all; }
        }

        internal static List<string> Categories
        {
            get { if (_categories == null) Load(); return _categories; }
        }

        private static string Attr(XmlNode n, string name)
        {
            if (n.Attributes == null) return "";
            XmlAttribute a = n.Attributes[name];
            return a == null ? "" : a.Value;
        }

        /// <summary>getImage 回调按控件 id 反查标签与 tag（用于生成图标）。</summary>
        internal static Feature ById(string id)
        {
            if (_byId == null) Load();
            Feature f;
            if (_byId != null && _byId.TryGetValue(id, out f)) return f;
            return null;
        }

        private static void Load()
        {
            _all = new List<Feature>();
            _byId = new Dictionary<string, Feature>(StringComparer.Ordinal);
            _categories = new List<string>();
            try
            {
                Assembly asm = Assembly.GetExecutingAssembly();
                using (Stream s = asm.GetManifestResourceStream("ExcleHelper.ribbon.xml"))
                {
                    if (s == null) return;
                    using (StreamReader r = new StreamReader(s, Encoding.UTF8))
                    {
                        XmlDocument doc = new XmlDocument();
                        doc.LoadXml(r.ReadToEnd());
                        XmlNamespaceManager ns = new XmlNamespaceManager(doc.NameTable);
                        string uri = doc.DocumentElement.NamespaceURI;
                        if (!string.IsNullOrEmpty(uri)) ns.AddNamespace("m", uri);
                        string tabPath = string.IsNullOrEmpty(uri) ? "//tab" : "//m:tab";
                        string groupPath = string.IsNullOrEmpty(uri) ? "group" : "m:group";
                        string buttonPath = string.IsNullOrEmpty(uri) ? ".//button" : ".//m:button";

                        foreach (XmlNode tab in doc.SelectNodes(tabPath, ns))
                        {
                            string tabLabel = Attr(tab, "label");
                            foreach (XmlNode g in tab.SelectNodes(groupPath, ns))
                            {
                                string gLabel = Attr(g, "label");
                                foreach (XmlNode b in g.SelectNodes(buttonPath, ns))
                                {
                                    string tag = Attr(b, "tag");
                                    if (tag.Length == 0) continue;
                                    Feature f = new Feature();
                                    f.Id = Attr(b, "id");
                                    f.Tag = tag;
                                    f.Label = Attr(b, "label");
                                    string st = Attr(b, "screentip");
                                    if (st.Length == 0) st = Attr(b, "supertip");
                                    f.Hint = st;
                                    f.Tab = tabLabel;
                                    f.Group = gLabel;
                                    _all.Add(f);
                                    if (!string.IsNullOrEmpty(f.Id)) _byId[f.Id] = f;
                                }
                                if (gLabel.Length > 0 && !_categories.Contains(gLabel)) _categories.Add(gLabel);
                            }
                        }
                    }
                }
            }
            catch { }
        }

        /// <summary>匹配：中文子串，或目标文字拼音首字母序列。</summary>
        internal static bool Match(Feature f, string query)
        {
            if (string.IsNullOrEmpty(query)) return true;
            string q = query.Trim().ToLowerInvariant();
            if (q.Length == 0) return true;
            if (f.Label.ToLowerInvariant().Contains(q)) return true;
            if (f.Group != null && f.Group.ToLowerInvariant().Contains(q)) return true;
            if (f.Hint != null && f.Hint.ToLowerInvariant().Contains(q)) return true;

            StringBuilder initials = new StringBuilder();
            foreach (char c in f.Label)
            {
                char p = CnOps.PyOf(c);
                if (p >= 'A' && p <= 'Z') initials.Append(char.ToLowerInvariant(p));
            }
            if (initials.Length > 0 && initials.ToString().Contains(q)) return true;
            return false;
        }
    }
}

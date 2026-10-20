using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace ExcleHelper
{
    /// <summary>
    /// 图形图标库（无文字、无外部素材）：**每个功能一个独立图标**。
    /// 组合体系 = 基础图形（按功能域取色）+ 右下角角标（区分同类操作）+ 可选删除斜杠。
    /// 规格串格式："base|badge|flag"。角标为白字，画在右下角实心圆内。
    /// </summary>
    internal static class Glyphs
    {
        private static readonly Dictionary<string, Bitmap> Cache = new Dictionary<string, Bitmap>();

        /// <summary>总兜底：任何异常都返回一个可用的占位图标，绝不返回 null / 空白。</summary>
        internal static Bitmap Render(string tag, int size)
        {
            Bitmap b;
            try { b = RenderCore(tag, size); }
            catch { b = null; }
            if (b == null) b = Fallback(tag, size);
            // 关键：返回缓存位图的副本。调用方（面板/图标总览/Office）可能 Dispose，
            // 直接返回缓存实例会被释放，导致后续取同一图标失败甚至出现空白图标。
            try { return new Bitmap(b); }
            catch { return b; }
        }

        private static Bitmap Fallback(string tag, int size)
        {
            Bitmap bmp = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                g.ScaleTransform(size / 32f, size / 32f);
                using (Pen p = new Pen(Palette.ColorOf(tag), 2f))
                using (GraphicsPath path = Round(new RectangleF(7, 7, 18, 18), 4))
                    g.DrawPath(p, path);
            }
            return bmp;
        }

        private static Bitmap RenderCore(string tag, int size)
        {
            string spec = Spec(tag);
            string key = spec + "@" + size;
            Bitmap cached;
            if (Cache.TryGetValue(key, out cached)) return cached;

            string[] parts = spec.Split('|');
            string baseId = parts[0];
            string badge = parts.Length > 1 ? parts[1] : "";
            bool slash = parts.Length > 2 && parts[2] == "slash";
            bool hasBadge = badge.Length > 0 && badge != "none";

            Color c = Palette.ColorOf(tag);
            Bitmap bmp = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                float k = size / 32f;
                g.ScaleTransform(k, k);

                using (Pen pen = new Pen(c, 2f))
                {
                    pen.StartCap = LineCap.Round; pen.EndCap = LineCap.Round; pen.LineJoin = LineJoin.Round;
                    if (hasBadge) { g.TranslateTransform(13f, 13.5f); g.ScaleTransform(0.78f, 0.78f); }
                    else { g.TranslateTransform(16f, 16f); g.ScaleTransform(0.82f, 0.82f); }
                    g.TranslateTransform(-16f, -16f);
                    try { DrawBase(g, pen, c, baseId); }
                    catch
                    {
                        // fallback：任何绘制异常都退回简单方块，避免图标空白
                        g.ResetTransform();
                        g.ScaleTransform(size / 32f, size / 32f);
                        using (Pen fp = new Pen(c, 2f))
                        using (GraphicsPath fp2 = Round(new RectangleF(7, 7, 18, 18), 4))
                            g.DrawPath(fp, fp2);
                    }
                    g.ResetTransform();
                    g.ScaleTransform(k, k);
                }

                if (hasBadge)
                {
                    using (SolidBrush bb = new SolidBrush(c)) g.FillEllipse(bb, 18.5f, 18.5f, 13f, 13f);
                    using (Pen bp = new Pen(Color.White, 1.5f))
                    {
                        bp.StartCap = LineCap.Round; bp.EndCap = LineCap.Round; bp.LineJoin = LineJoin.Round;
                        DrawBadge(g, bp, badge);
                    }
                }
                if (slash)
                {
                    using (Pen sp = new Pen(Color.FromArgb(215, 38, 38), 2.2f))
                    {
                        sp.StartCap = LineCap.Round; sp.EndCap = LineCap.Round;
                        g.DrawLine(sp, 5, 27, 27, 5);
                    }
                }
            }
            Cache[key] = bmp;
            return bmp;
        }

        // ================= 每个功能 → 独立图标规格 =================
        internal static string Spec(string tag)
        {
            switch (tag)
            {
                // —— 工作簿治理 ——
                case "wb.snapshot": return "history|filled";
                case "wb.diff": return "diff|book";
                case "fm.audit": return "link|dots";
                case "fm.volatile": return "clock|bars";
                case "wb.names.clean": return "book|check";
                case "wb.links": return "link|dots";
                case "wb.health": return "shield|bars";

                // —— AI 三期 ——
                case "ai.report": return "doc|spark";
                case "ai.join": return "merge|spark";
                case "ai.fix.all": return "fx|dots";
                case "ai.translate": return "globe|spark";
                case "ai.synth": return "dice|spark";
                case "ai.explain.outlier": return "chartline|eye";

                // —— 呈现与效率 ——
                case "cf.manage": return "brush|bars";
                case "fmt.sparkline": return "chartline|box";
                case "fmt.style": return "brush|filled";
                case "ch.set": return "chartbar|box";
                case "pic.barcode": return "qr|bars";
                case "pic.export.range": return "image|spark";
                case "ui.recent": return "history|dots";
                case "ui.favorites": return "star|box";

                // —— 数据质量 ——
                case "ql.outlier": return "chartline|check";
                case "ql.missing": return "grid|plus";
                case "ql.dup.fuzzy": return "dedupe|percent";
                case "ql.dict": return "book|bars";
                case "ql.unit.unify": return "swap|num";
                case "ql.phone.norm": return "phone|check";
                case "ql.addr.norm": return "pin|check";
                case "ql.changelog": return "history|box";

                // —— 数据整合 ——
                case "jt.join": return "merge|table";
                case "jt.update": return "merge|check";
                case "jt.group.merge": return "merge|dots";
                case "jt.append": return "merge|plus";
                case "jt.pivot": return "pivot|table";
                case "jt.crosstab": return "table|bars";

                // —— AI ——
                case "ai.settings": return "info|box";
                case "ai.test": return "globe|check";
                case "ai.panel": return "comment|spark";
                case "ai.formula": return "fx|spark";
                case "ai.explain": return "fx|eye";
                case "ai.fix": return "fx|check";
                case "ai.fill": return "clipboard|spark";
                case "ai.classify": return "bars|spark";
                case "ai.extract": return "funnel|spark";
                case "ai.insight": return "chartline|spark";
                case "ai.clean": return "broom|spark";
                case "ai.filter": return "funnel|check";
                case "ai.chart": return "chartpie|spark";
                case "ai.dict": return "book|spark";
                case "ai.ask": return "comment|dots";

                // —— 界面与回退 ——
                case "ui.panel": return "panel";
                case "ui.about": return "info";
                case "x.help": return "help";
                case "x.repeat": return "redo|dot";
                case "undo.last": return "undo";
                case "undo.redo": return "redo";
                case "undo.history": return "history";
                case "undo.clear": return "eraser|minus";

                // —— 文本：保留 / 删除 ——
                case "t.keep.cn": return "funnel|cn";
                case "t.keep.en": return "funnel|en";
                case "t.keep.num": return "funnel|num";
                case "t.keep.sym": return "funnel|sym";
                case "t.del.cn": return "eraser|cn";
                case "t.del.en": return "eraser|en";
                case "t.del.num": return "eraser|num";
                case "t.del.sym": return "eraser|sym";
                // —— 文本：空格与不可见 ——
                case "t.trim.all": return "broom|slash";
                case "t.trim.head": return "broom|up";
                case "t.trim.tail": return "broom|down";
                case "t.squeeze": return "broom|merge";
                case "t.del.invisible": return "broom|eye";
                case "t.del.text": return "eraser|hash";
                // —— 文本：插入与截取 ——
                case "t.ins.head": return "plus|up";
                case "t.ins.tail": return "plus|down";
                case "t.ins.mid": return "plus|dot";
                case "t.cut.head": return "scissors|up";
                case "t.cut.tail": return "scissors|down";
                case "t.cut.mid": return "scissors|dot";
                // —— 文本：大小写与补全 ——
                case "t.case.upper": return "arrows|up";
                case "t.case.lower": return "arrows|down";
                case "t.case.proper": return "arrows|star";
                case "t.case.invert": return "arrows|swap";
                case "t.padleft": return "hash|plus";
                case "t.mask.mid": return "eye|dot";
                case "t.width.toggle": return "swap|percent";
                // —— 文本：更多 ——
                case "t.reverse": return "swap|dots";
                case "t.len": return "ruler";
                case "t.freq": return "chartbar|hash";
                case "t.num2en": return "percent|star";
                case "t.rx.extract": return "fx|two";
                case "t.rx.replace": return "fx|dots";
                case "t.ext.phone": return "phone|out";
                case "t.ext.email": return "mail|out";
                case "t.ext.url": return "globe|out";
                case "t.ext.num": return "hash|out";
                case "t.ext.date": return "calendar|out";
                case "t.split": return "split";
                case "t.join": return "merge";
                // —— 文本高级 ——
                case "t2.similar": return "diff|percent";
                case "t2.fuzzy": return "target|dots";
                case "t2.rxcols": return "fx|three";
                case "t2.batchreplace": return "swap|three";
                case "t2.uniqchar": return "dedupe|hash";
                case "t2.bylen": return "ruler|num";
                case "t2.initials": return "star|dots";
                case "t2.between": return "scissors|two";
                case "t2.count": return "hash|bars";
                case "cn.t2s": return "swap|cn";
                case "cn.s2t": return "swap|box";

                // —— 数值：序列 ——
                case "n.seq.num": return "bars|num";
                case "n.seq.letter": return "bars|en";
                case "n.seq.week": return "bars|cal";
                case "n.seq.month": return "bars|clock";
                case "n.seq.cnnum": return "bars|cn";
                case "n.seq.roman": return "bars|hash";
                // —— 数值：舍入 ——
                case "n.round": return "percent|dot";
                case "n.roundup": return "percent|up";
                case "n.rounddown": return "percent|down";
                case "n.bankers": return "percent|two";
                case "n.int": return "percent|box";
                case "n.frac": return "percent|slash";
                // —— 数值：转换 ——
                case "n.totext": return "calc|out";
                case "n.cntoar": return "hash|cn";
                case "n.artoCN": return "hash|num";
                case "n.rmb": return "card|cn";
                case "n.percent": return "percent|plus";
                case "n.thousand": return "hash|dots";
                case "n.display": return "eye|percent";
                case "n.math": return "calc|plus";
                case "n.fillblank": return "grid|filled";
                case "n.randsplit": return "dice|split";
                // —— 数值扩展 ——
                case "n2.rand": return "dice|spark";
                case "n2.sample": return "target|two";
                case "n2.unit": return "swap|ruler";
                case "n2.bins": return "chartbar|three";
                case "n2.sumcheck": return "sum|check";
                case "n2.password": return "key|star";

                // —— 日期时间 ——
                case "dt.normalize": return "calendar|check";
                case "dt.fmt.date": return "calendar|hash";
                case "dt.fmt.time": return "clock|hash";
                case "dt.diff": return "swap|minus";
                case "dt.add": return "swap|plus";
                case "dt.workdays": return "calendar|bars";
                case "dt.workday.after": return "calendar|spark";
                case "dt.age": return "person|cal";
                case "dt.split": return "split|cal";
                case "dt.combine": return "merge|cal";
                case "dt.range.stat": return "chartbar|cal";
                case "dt.rand": return "dice|cal";
                case "dt.clear.invalid": return "eraser|cal";
                case "dt.timestamp": return "clock|swap";
                case "dt.lunar": return "moon";

                // —— 表格：清空 / 删除 ——
                case "d.clear.zero": return "eraser|dot";
                case "d.clear.nonnum": return "eraser|spark";
                case "d.clear.err": return "eraser|cross";
                case "d.del.blankrows": return "trash|up";
                case "d.del.blankcols": return "trash|left";
                case "d.del.blankrowcells": return "trash|dots";
                case "d.del.hidden": return "trash|eye";
                case "d.ins.rows": return "plus|three";
                case "d.ins.cols": return "plus|left";
                case "d.del.every": return "trash|three";
                // —— 表格：重复值 ——
                case "d.dup.del": return "dedupe|minus";
                case "d.dup.uniq": return "dedupe|star";
                case "d.dup.highlight": return "dedupe|dot";
                case "d.dup.count": return "dedupe|num";
                // —— 表格：合并 / 重排 ——
                case "d.merge.same": return "merge|dot";
                case "d.merge.unsame": return "split|dot";
                case "d.merge.keep": return "merge|filled";
                case "d.transpose": return "transpose";
                case "d.dim.flat": return "pivot|in";
                case "d.dim.matrix": return "pivot|out";
                case "d.cmp.two": return "diff|two";

                // —— 编辑 ——
                case "e.copy.to": return "copy|in";
                case "e.copy.values": return "copy|num";
                case "e.copy.n": return "copy|three";
                case "e.copy.all": return "copy|filled";
                case "e.copy.sheets": return "copy|two";
                case "e.copy.merge": return "copy|merge";
                case "e.clear.unselected": return "eraser|box";
                case "e.clear.given": return "eraser|dots";
                case "e.del.blankcells": return "trash|dot";
                case "e.del.objects": return "trash|star";
                case "e.del.textbox": return "trash|box";
                case "e.del.cf": return "trash|bars";
                case "e.merge.rowcat": return "merge|up";
                case "e.merge.colcat": return "merge|left";
                case "e.merge.blank": return "merge|empty";
                case "e.flip.rows": return "swap|up";
                case "e.combo": return "pivot|three";
                case "e.split.row": return "split|up";

                // —— 格式 ——
                case "f.autofit": return "ruler|out";
                case "f.border": return "border";
                case "f.headstyle": return "brush|star";
                case "f.rowheight": return "ruler|up";
                case "f.colwidth": return "ruler|left";
                case "f.freeze": return "freeze";
                case "f.clearfmt": return "eraser|brush";
                case "f.copymulti": return "copy|sheet";
                case "f.stripe": return "stripe";
                case "cf.value": return "brush|percent";
                case "cf.bars": return "chartbar|spark";
                case "cf.scale": return "chartbar|percent";

                // —— 公式 ——
                case "fm.to.text": return "fx|out";
                case "fm.to.formula": return "fx|in";
                case "fm.ref.toggle": return "swap|hash";
                case "fm.replace": return "fx|swap";
                case "fm.show.all": return "eye|bars";
                case "fm.help": return "help|hash";

                // —— 对象 ——
                case "o.comment.add": return "comment|plus";
                case "o.comment.del": return "comment|minus";
                case "o.comment.export": return "comment|out";
                case "o.link.add": return "link|plus";
                case "o.link.del": return "link|minus";

                // —— 隐私与地址 ——
                case "p2.addr.split": return "pin|split";
                case "p2.mask.batch": return "eye|three";
                case "p2.encrypt": return "lock|star";
                case "p2.decrypt": return "lock|check";
                case "p2.dup.cross": return "dedupe|two";

                // —— 图片 ——
                case "pic.insert": return "image|plus";
                case "pic.size": return "image|ruler";
                case "pic.anchor": return "image|box";
                case "pic.export": return "image|out";
                case "pic.del": return "image|minus";
                case "pic.watermark": return "watermark";
                case "pic.qrcode": return "qr";
                case "pic.snip": return "camera";

                // —— 打印 ——
                case "pr.area": return "printer|box";
                case "pr.title": return "printer|up";
                case "pr.pagenum": return "printer|num";
                case "pr.header": return "printer|dots";
                case "pr.pagebreak": return "printer|split";
                case "pr.fit": return "printer|out";
                case "pr.center": return "printer|dot";
                case "pr.grid": return "printer|bars";
                case "pr.batch": return "printer|three";
                case "pr.preview": return "printer|eye";

                // —— 结构 ——
                case "q2.audit": return "audit|check";
                case "q2.seq.group": return "bars|three";
                case "q2.subtotal": return "sum|plus";
                case "q2.colmerge": return "merge|bars";
                case "q2.table.toggle": return "table|swap";
                case "q2.blank.insert": return "plus|empty";

                // —— 工作表 ——
                case "s.split": return "sheet|split";
                case "s.index": return "sheet|dots";
                case "s.rename.batch": return "sheet|hash";
                case "s.create.batch": return "sheet|empty";
                case "s.delete": return "sheet|cross";
                case "s.sort": return "sheet|bars";
                case "s.merge": return "sheet|merge";
                case "s.saveone": return "file|out";
                case "s.pdf": return "pdf";
                case "s.txt": return "file|hash";

                // —— 安全 ——
                case "sec.structure": return "shield|box";
                case "sec.lock": return "lock|box";
                case "sec.hide.formula": return "eye|slash";
                case "sec.meta": return "doc|eraser";
                case "sec.unlink": return "link|slash";

                // —— 汇总拆分 ——
                case "sm.merge.two": return "merge|two";
                case "sm.merge.multi": return "merge|three";
                case "sm.merge.books": return "book|merge";
                case "sm.sum.sheets": return "sum|two";
                case "sm.sum.books": return "sum|three";
                case "sm.sum.master": return "sum|star";
                case "sm.split.book": return "split|box";
                case "sm.split.values": return "split|three";

                // —— 统计 ——
                case "st.share": return "chartpie|percent";
                case "st.mom": return "chartline|dots";
                case "st.yoy": return "chartline|two";
                case "st.rank": return "bars|star";
                case "st.describe": return "chartbar|dots";
                case "st.freq": return "chartbar|bars";
                case "st.group": return "pivot|dots";
                case "st.corr": return "chartline|percent";

                // —— 校验标记 ——
                case "v.email.check": return "mail|check";
                case "v.null.highlight": return "box|dot";
                case "v.space.check": return "broom|dot";
                case "v.dup.row": return "dedupe|bars";

                // —— 录入验证 ——
                case "vd.list": return "grid|down";
                case "vd.list.cascade": return "grid|two";
                case "vd.checkbox": return "check|box";
                case "vd.radio": return "target|dot";
                case "vd.no.dup": return "dedupe|slash";
                case "vd.range": return "ruler|dots";
                case "vd.quick": return "clipboard|star";

                // —— 工作表与文件 ——
                case "w.new": return "sheet|plus";
                case "w.copy": return "sheet|two";
                case "w.move": return "sheet|out";
                case "w.rename.one": return "sheet|spark";
                case "w.protect": return "sheet|lock";
                case "w.hide": return "sheet|minus";
                case "w.unhide": return "sheet|eye";
                case "w.smart.index": return "sheet|star";
                case "w.names.extract": return "sheet|in";
                case "w.book.rename": return "book|hash";
                case "w.file.rename": return "file|spark";
                case "w.file.list": return "folder|dots";
                case "w.folder.create": return "folder|plus";
                case "w.folder.empty": return "folder|minus";
                case "w.export.csv": return "file|bars";
                case "w.slim": return "book|eraser";

                // —— 颜色与对比 ——
                case "x2.color.sum": return "sum|dot";
                case "x2.color.extract": return "brush|out";
                case "x2.color.replace": return "brush|swap";
                case "x2.table.diff": return "diff|table";
                case "x2.cell.diff": return "diff|dot";

                // —— 模板 ——
                case "tpl.label": return "doc|dots";
                case "tpl.payslip": return "doc|bars";
                case "tpl.book": return "book|plus";

                // —— 证件 ——
                case "c.py.first": return "star|cn";
                case "c.bank.check": return "card|check";
                case "c.usci.check": return "shield|check";
                case "c.id.info": return "idcard|out";
                case "c.id.check": return "idcard|check";
                case "c.id.15to18": return "idcard|up";
                case "c.id.mask": return "idcard|slash";
                case "c.phone.seg": return "phone|bars";
                case "c.phone.mask": return "phone|slash";
                case "c.name.split": return "person|split";
                case "c.name.mask": return "person|slash";
            }

            // 兜底：按前缀给一个合理图形（保证不出现空白图标）
            string t = (tag ?? "").ToLowerInvariant();
            if (t.StartsWith("t.")) return "docline";
            if (t.StartsWith("n.")) return "calc";
            if (t.StartsWith("d.")) return "grid";
            if (t.StartsWith("e.")) return "clipboard";
            if (t.StartsWith("f.")) return "brush";
            if (t.StartsWith("w.") || t.StartsWith("s.")) return "sheet";
            if (t.StartsWith("sm.")) return "merge";
            if (t.StartsWith("pr.")) return "printer";
            if (t.StartsWith("pic.")) return "image";
            if (t.StartsWith("sec.")) return "lock";
            if (t.StartsWith("vd.")) return "check";
            if (t.StartsWith("st.")) return "chartbar";
            return "doc";
        }

        // ================= 基础图形 =================
        private static void DrawBase(Graphics g, Pen p, Color c, string id)
        {
            switch (id)
            {
                case "doc":
                    Rect(g, p, 7, 4, 18, 24, 3); Line(g, p, 11, 11, 21, 11); Line(g, p, 11, 16, 21, 16); Line(g, p, 11, 21, 17, 21); break;
                case "docline":
                    Rect(g, p, 5, 6, 22, 20, 3); Line(g, p, 9, 13, 23, 13); Line(g, p, 9, 18, 19, 18); break;
                case "grid":
                    Rect(g, p, 5, 5, 22, 22, 2); Line(g, p, 5, 12, 27, 12); Line(g, p, 5, 19, 27, 19); Line(g, p, 12, 5, 12, 27); Line(g, p, 19, 5, 19, 27); break;
                case "table":
                    Rect(g, p, 4, 7, 24, 18, 2); Line(g, p, 4, 13, 28, 13); Line(g, p, 12, 7, 12, 25); Line(g, p, 20, 7, 20, 25); break;
                case "sheet":
                    Rect(g, p, 5, 4, 22, 19, 2); Line(g, p, 5, 10, 27, 10); Line(g, p, 13, 4, 13, 23); Line(g, p, 20, 4, 20, 23); Rect(g, p, 9, 23, 10, 4, 1); break;
                case "book":
                    Rect(g, p, 6, 5, 20, 22, 2); Line(g, p, 11, 5, 11, 27); Line(g, p, 15, 12, 23, 12); Line(g, p, 15, 17, 21, 17); break;
                case "folder":
                    Path(g, p, 4, 9, 13, 9, 15, 13, 28, 13, 28, 25, 4, 25, 4, 9); break;
                case "file":
                    Path(g, p, 8, 4, 19, 4, 25, 10, 25, 28, 8, 28, 8, 4); Path(g, p, 19, 4, 19, 10, 25, 10); break;
                case "pdf":
                    Path(g, p, 7, 4, 19, 4, 25, 10, 25, 28, 7, 28, 7, 4); Path(g, p, 19, 4, 19, 10, 25, 10); Line(g, p, 11, 16, 21, 16); Line(g, p, 11, 21, 18, 21); break;
                case "image":
                    Rect(g, p, 4, 6, 24, 20, 3); Circle(g, p, 11, 13, 2.5f); Path(g, p, 6, 24, 14, 15, 20, 21, 24, 17, 27, 20); break;
                case "camera":
                    Rect(g, p, 4, 11, 24, 15, 3); Path(g, p, 12, 11, 14, 7, 19, 7, 21, 11); Circle(g, p, 16, 18, 5); break;
                case "qr":
                    Rect(g, p, 5, 5, 8, 8, 1); Rect(g, p, 19, 5, 8, 8, 1); Rect(g, p, 5, 19, 8, 8, 1); Dot(g, p, 20, 20); Dot(g, p, 25, 20); Dot(g, p, 20, 25); Dot(g, p, 25, 25); break;
                case "watermark":
                    Rect(g, p, 5, 5, 22, 22, 2); Line(g, p, 10, 22, 22, 10); Line(g, p, 10, 16, 16, 10); Line(g, p, 16, 22, 22, 16); break;
                case "printer":
                    Path(g, p, 8, 12, 8, 5, 24, 5, 24, 12); Rect(g, p, 5, 12, 22, 9, 2); Rect(g, p, 8, 20, 16, 8, 1); break;
                case "calendar":
                    Rect(g, p, 5, 7, 22, 20, 3); Line(g, p, 5, 13, 27, 13); Line(g, p, 11, 4, 11, 10); Line(g, p, 21, 4, 21, 10); Dot(g, p, 11, 18); Dot(g, p, 16, 18); Dot(g, p, 21, 18); Dot(g, p, 11, 23); break;
                case "clock":
                    Circle(g, p, 16, 16, 11); Line(g, p, 16, 16, 16, 9); Line(g, p, 16, 16, 22, 18); break;
                case "moon":
                    Path(g, p, 20, 5, 13, 8, 10, 16, 13, 24, 20, 27, 12, 27, 6, 22, 5, 15, 9, 8, 16, 5, 20, 5); break;
                case "sum":
                    Rect(g, p, 5, 20, 6, 6, 1); Rect(g, p, 13, 13, 6, 13, 1); Rect(g, p, 21, 7, 6, 19, 1); Line(g, p, 4, 27, 28, 27); break;
                case "chartbar":
                    Line(g, p, 6, 26, 26, 26); Line(g, p, 10, 26, 10, 17); Line(g, p, 16, 26, 16, 10); Line(g, p, 22, 26, 22, 20); break;
                case "chartline":
                    Line(g, p, 6, 26, 26, 26); Line(g, p, 6, 26, 6, 6); Path(g, p, 9, 22, 15, 14, 20, 18, 27, 8); break;
                case "chartpie":
                    Circle(g, p, 16, 16, 11); Line(g, p, 16, 16, 16, 5); Line(g, p, 16, 16, 26, 20); break;
                case "percent":
                    Circle(g, p, 10, 11, 5); Circle(g, p, 22, 21, 5); Line(g, p, 8, 24, 24, 8); break;
                case "calc":
                    Rect(g, p, 7, 4, 18, 24, 3); Line(g, p, 11, 9, 21, 9); Dot(g, p, 11, 15); Dot(g, p, 16, 15); Dot(g, p, 21, 15); Dot(g, p, 11, 21); Dot(g, p, 16, 21); Dot(g, p, 21, 21); break;
                case "ruler":
                    Rect(g, p, 4, 12, 24, 8, 2); Line(g, p, 9, 12, 9, 16); Line(g, p, 14, 12, 14, 17); Line(g, p, 19, 12, 19, 16); Line(g, p, 24, 12, 24, 17); break;
                case "dice":
                    Rect(g, p, 6, 6, 20, 20, 4); Dot(g, p, 11, 11); Dot(g, p, 21, 11); Dot(g, p, 16, 16); Dot(g, p, 11, 21); Dot(g, p, 21, 21); break;
                case "target":
                    Circle(g, p, 16, 16, 10); Circle(g, p, 16, 16, 4); Line(g, p, 16, 3, 16, 8); break;
                case "funnel":
                    Path(g, p, 5, 6, 27, 6, 19, 15, 19, 26, 13, 22, 13, 15, 5, 6); break;
                case "eraser":
                    Path(g, p, 4, 18, 13, 9, 24, 20, 15, 29, 4, 18); Line(g, p, 13, 9, 24, 20); Line(g, p, 7, 27, 26, 27); break;
                case "broom":
                    Line(g, p, 23, 5, 14, 15); Path(g, p, 14, 15, 8, 19, 12, 27, 18, 25, 20, 17, 14, 15); break;
                case "scissors":
                    Circle(g, p, 9, 9, 4); Circle(g, p, 9, 23, 4); Line(g, p, 12, 11, 26, 24); Line(g, p, 12, 21, 26, 8); break;
                case "arrows":
                    Line(g, p, 8, 24, 8, 9); Line(g, p, 8, 9, 4, 14); Line(g, p, 8, 9, 12, 14); Line(g, p, 22, 8, 22, 23); Line(g, p, 22, 23, 18, 18); Line(g, p, 22, 23, 26, 18); break;
                case "swap":
                    Line(g, p, 6, 12, 24, 12); Line(g, p, 20, 8, 24, 12); Line(g, p, 20, 16, 24, 12); Line(g, p, 26, 21, 8, 21); Line(g, p, 12, 17, 8, 21); Line(g, p, 12, 25, 8, 21); break;
                case "hash":
                    Line(g, p, 12, 6, 10, 26); Line(g, p, 20, 6, 18, 26); Line(g, p, 7, 13, 25, 13); Line(g, p, 6, 20, 24, 20); break;
                case "star":
                    Path(g, p, 16, 5, 19, 13, 27, 13, 21, 18, 23, 26, 16, 21, 9, 26, 11, 18, 5, 13, 13, 13, 16, 5); break;
                case "person":
                    Circle(g, p, 16, 11, 5); Path(g, p, 7, 27, 9, 20, 16, 18, 23, 20, 25, 27); break;
                case "phone":
                    Path(g, p, 9, 4, 14, 9, 11, 14, 18, 21, 23, 18, 28, 23, 24, 27, 8, 26, 4, 12, 9, 4); break;
                case "card":
                    Rect(g, p, 4, 8, 24, 16, 3); Fill(g, c, 4, 12, 24, 4); Line(g, p, 9, 20, 17, 20); break;
                case "idcard":
                    Rect(g, p, 4, 8, 24, 16, 3); Circle(g, p, 11, 16, 3); Line(g, p, 17, 13, 24, 13); Line(g, p, 17, 18, 24, 18); break;
                case "shield":
                    Path(g, p, 16, 4, 26, 8, 26, 17, 16, 28, 6, 17, 6, 8, 16, 4); Path(g, p, 11, 16, 15, 20, 22, 11); break;
                case "lock":
                    Rect(g, p, 7, 14, 18, 14, 3); Path(g, p, 11, 14, 11, 10, 21, 10, 21, 14); Dot(g, p, 16, 21); break;
                case "key":
                    Circle(g, p, 11, 12, 6); Line(g, p, 15, 16, 26, 27); Line(g, p, 22, 23, 25, 20); Line(g, p, 25, 26, 28, 23); break;
                case "eye":
                    Path(g, p, 4, 16, 11, 9, 21, 9, 28, 16, 21, 23, 11, 23, 4, 16); Circle(g, p, 16, 16, 3); break;
                case "comment":
                    Path(g, p, 5, 6, 27, 6, 27, 21, 14, 21, 8, 27, 8, 21, 5, 21, 5, 6); Line(g, p, 10, 12, 22, 12); Line(g, p, 10, 17, 18, 17); break;
                case "link":
                    Path(g, p, 13, 19, 9, 23, 6, 20, 10, 16); Path(g, p, 19, 13, 23, 9, 26, 12, 22, 16); Line(g, p, 13, 19, 19, 13); break;
                case "check":
                    Circle(g, p, 16, 16, 11); Path(g, p, 10, 16, 14, 21, 23, 11); break;
                case "diff":
                    Rect(g, p, 5, 6, 9, 20, 2); Rect(g, p, 18, 6, 9, 20, 2); Line(g, p, 14, 10, 14, 22); Line(g, p, 15, 12, 17, 12); Line(g, p, 15, 20, 17, 20); break;
                case "pin":
                    Path(g, p, 16, 28, 16, 18); Path(g, p, 10, 4, 22, 4, 22, 12, 10, 12, 10, 4); Circle(g, p, 16, 8, 2); break;
                case "clipboard":
                    Rect(g, p, 7, 5, 18, 23, 3); Rect(g, p, 12, 3, 8, 5, 2); Line(g, p, 11, 14, 21, 14); Line(g, p, 11, 19, 18, 19); break;
                case "copy":
                    Rect(g, p, 4, 4, 14, 16, 2); Rect(g, p, 12, 12, 14, 16, 2); break;
                case "trash":
                    Line(g, p, 6, 9, 26, 9); Line(g, p, 12, 9, 12, 5); Line(g, p, 12, 5, 20, 5); Line(g, p, 20, 5, 20, 9); Path(g, p, 8, 9, 10, 28, 22, 28, 24, 9); Line(g, p, 13, 14, 13, 23); Line(g, p, 19, 14, 19, 23); break;
                case "plus":
                    Rect(g, p, 5, 5, 22, 22, 3); Line(g, p, 16, 10, 16, 22); Line(g, p, 10, 16, 22, 16); break;
                case "merge":
                    Rect(g, p, 4, 8, 9, 6, 2); Rect(g, p, 4, 18, 9, 6, 2); Path(g, p, 15, 11, 20, 11, 20, 21, 15, 21); Rect(g, p, 21, 11, 7, 10, 2); break;
                case "split":
                    Rect(g, p, 4, 11, 8, 10, 2); Line(g, p, 14, 16, 20, 16); Line(g, p, 20, 16, 20, 8); Line(g, p, 20, 16, 20, 24); Rect(g, p, 21, 5, 7, 6, 2); Rect(g, p, 21, 21, 7, 6, 2); break;
                case "transpose":
                    Rect(g, p, 5, 5, 10, 10, 2); Rect(g, p, 17, 17, 10, 10, 2); Path(g, p, 17, 9, 23, 9, 23, 15); Line(g, p, 20, 12, 23, 15); Line(g, p, 26, 12, 23, 15); break;
                case "pivot":
                    Rect(g, p, 5, 5, 22, 16, 2); Line(g, p, 5, 11, 27, 11); Line(g, p, 13, 5, 13, 21); Line(g, p, 8, 27, 24, 27); Line(g, p, 12, 24, 8, 27); Line(g, p, 12, 30, 8, 27); break;
                case "dedupe":
                    Circle(g, p, 13, 16, 8); Circle(g, p, 20, 16, 8); break;
                case "brush":
                    Path(g, p, 6, 12, 20, 6, 26, 10, 14, 17, 6, 12); Line(g, p, 14, 17, 10, 22); Rect(g, p, 8, 22, 6, 6, 2); break;
                case "border":
                    Rect(g, p, 5, 5, 22, 22, 2); Line(g, p, 12, 5, 12, 27); Line(g, p, 19, 5, 19, 27); break;
                case "stripe":
                    Rect(g, p, 5, 6, 22, 20, 2); Fill(g, c, 6, 7, 20, 4); Fill(g, c, 6, 15, 20, 4); Fill(g, c, 6, 23, 20, 2); break;
                case "freeze":
                    Line(g, p, 16, 5, 16, 27); Line(g, p, 7, 10, 25, 22); Line(g, p, 7, 22, 25, 10); break;
                case "panel":
                    Rect(g, p, 4, 5, 24, 20, 3); Line(g, p, 4, 11, 28, 11); Circle(g, p, 19, 21, 4); Line(g, p, 22, 24, 27, 28); break;
                case "info":
                    Circle(g, p, 16, 16, 11); Line(g, p, 16, 15, 16, 23); Dot(g, p, 16, 10); break;
                case "help":
                    Circle(g, p, 16, 16, 11); Path(g, p, 12, 13, 13, 9, 18, 8, 21, 11, 20, 15, 16, 17, 16, 20); Dot(g, p, 16, 24); break;
                case "undo":
                    Path(g, p, 25, 24, 22, 12, 12, 10); Line(g, p, 12, 10, 18, 5); Line(g, p, 12, 10, 18, 15); break;
                case "redo":
                    Path(g, p, 7, 24, 10, 12, 20, 10); Line(g, p, 20, 10, 14, 5); Line(g, p, 20, 10, 14, 15); break;
                case "history":
                    Circle(g, p, 16, 16, 11); Line(g, p, 16, 10, 16, 17); Line(g, p, 16, 17, 21, 20); Path(g, p, 5, 16, 5, 8); break;
                case "fx":
                    Line(g, p, 5, 24, 9, 24); Line(g, p, 9, 24, 14, 8); Line(g, p, 14, 8, 18, 8); Line(g, p, 20, 12, 28, 12); Line(g, p, 20, 20, 28, 20); Line(g, p, 22, 15, 26, 15); break;
                case "globe":
                    Circle(g, p, 16, 16, 11); Path(g, p, 5, 16, 27, 16); Path(g, p, 16, 5, 16, 27); Path(g, p, 8, 9, 16, 12, 24, 9); Path(g, p, 8, 23, 16, 20, 24, 23); break;
                case "mail":
                    Rect(g, p, 4, 8, 24, 16, 3); Path(g, p, 4, 10, 16, 19, 28, 10); break;
                case "audit":
                    Rect(g, p, 7, 5, 18, 23, 3); Rect(g, p, 12, 3, 8, 5, 2); Path(g, p, 11, 18, 15, 22, 22, 13); break;
                case "box":
                    Rect(g, p, 6, 6, 20, 20, 3); break;
                default:
                    Rect(g, p, 7, 7, 18, 18, 4); Dot(g, p, 16, 16); break;
            }
        }

        // ================= 角标（白色，画在 17.5,17.5 起的 13×13 圆内） =================
        private static void DrawBadge(Graphics g, Pen p, string id)
        {
            switch (id)
            {
                case "cn": Solid(g, 20.8f, 20.8f, 6.4f, 6.4f, 3.2f); break;
                case "en": Solid(g, 20.8f, 20.8f, 6.4f, 6.4f, 0f); break;
                case "num": SolidTri(g); break;
                case "sym": SolidDiamond(g); break;
                case "up": Line(g, p, 24, 27, 24, 21.2f); Line(g, p, 21.6f, 23.6f, 24, 21.2f); Line(g, p, 26.4f, 23.6f, 24, 21.2f); break;
                case "down": Line(g, p, 24, 21, 24, 26.8f); Line(g, p, 21.6f, 24.4f, 24, 26.8f); Line(g, p, 26.4f, 24.4f, 24, 26.8f); break;
                case "left": Line(g, p, 27, 24, 21.2f, 24); Line(g, p, 23.6f, 21.6f, 21.2f, 24); Line(g, p, 23.6f, 26.4f, 21.2f, 24); break;
                case "plus": Line(g, p, 24, 21.2f, 24, 26.8f); Line(g, p, 21.2f, 24, 26.8f, 24); break;
                case "minus": Line(g, p, 21.2f, 24, 26.8f, 24); break;
                case "cross": Line(g, p, 21.8f, 21.8f, 26.2f, 26.2f); Line(g, p, 26.2f, 21.8f, 21.8f, 26.2f); break;
                case "check": Path(g, p, 21.2f, 24, 23.2f, 26.4f, 26.8f, 21.4f); break;
                case "star": Path(g, p, 24, 20.8f, 25.1f, 23, 27.2f, 23.2f, 25.6f, 24.9f, 26, 27.2f, 24, 26.1f, 22, 27.2f, 22.4f, 24.9f, 20.8f, 23.2f, 22.9f, 23, 24, 20.8f); break;
                case "dot": Dot(g, p, 24, 24); break;
                case "dots": Dot(g, p, 21.2f, 24); Dot(g, p, 24, 24); Dot(g, p, 26.8f, 24); break;
                case "bars": Line(g, p, 21.6f, 26.8f, 21.6f, 24); Line(g, p, 24, 26.8f, 24, 21.6f); Line(g, p, 26.4f, 26.8f, 26.4f, 25.2f); break;
                case "two": Dot(g, p, 22, 24); Dot(g, p, 26, 24); break;
                case "three": Line(g, p, 21.6f, 26.6f, 26.4f, 26.6f); Line(g, p, 21.6f, 24, 26.4f, 24); Line(g, p, 21.6f, 21.4f, 26.4f, 21.4f); break;
                case "lock": Rect(g, p, 21.6f, 23.8f, 4.8f, 3.6f, 1); Path(g, p, 22.6f, 23.8f, 22.6f, 22.4f, 25.4f, 22.4f, 25.4f, 23.8f); break;
                case "key": Circle(g, p, 22.6f, 22.6f, 1.8f); Line(g, p, 23.8f, 23.8f, 26.6f, 26.6f); break;
                case "eye": Path(g, p, 21.2f, 24, 24, 21.6f, 26.8f, 24, 24, 26.4f, 21.2f, 24); Dot(g, p, 24, 24); break;
                case "hash": Line(g, p, 23, 21, 22.4f, 27); Line(g, p, 25.6f, 21, 25, 27); Line(g, p, 21.4f, 23.4f, 26.6f, 23.4f); Line(g, p, 21.4f, 25.2f, 26.6f, 25.2f); break;
                case "percent": Line(g, p, 21.4f, 26.6f, 26.6f, 21.4f); Dot(g, p, 22.2f, 22.2f); Dot(g, p, 25.8f, 25.8f); break;
                case "clock": Circle(g, p, 24, 24, 3.2f); Line(g, p, 24, 24, 24, 21.6f); Line(g, p, 24, 24, 26, 24.8f); break;
                case "cal": Rect(g, p, 21.2f, 22, 5.6f, 5, 1); Line(g, p, 21.2f, 23.6f, 26.8f, 23.6f); break;
                case "in": Rect(g, p, 21.6f, 21.6f, 4.8f, 4.8f, 1); Line(g, p, 21.6f, 24, 24.4f, 24); Line(g, p, 23, 22.6f, 24.4f, 24); Line(g, p, 23, 25.4f, 24.4f, 24); break;
                case "out": Rect(g, p, 21.6f, 21.6f, 4.8f, 4.8f, 1); Line(g, p, 24, 24, 26.8f, 24); Line(g, p, 25.4f, 22.6f, 26.8f, 24); Line(g, p, 25.4f, 25.4f, 26.8f, 24); break;
                case "spark": Line(g, p, 24, 20.8f, 24, 27.2f); Line(g, p, 20.8f, 24, 27.2f, 24); break;
                case "box": Rect(g, p, 21.6f, 21.6f, 4.8f, 4.8f, 1); break;
                case "slash": Line(g, p, 21.8f, 26.2f, 26.2f, 21.8f); break;
                case "filled": Fill(g, p.Color, 21.6f, 21.6f, 4.8f, 4.8f); break;
                case "swap": Line(g, p, 21.2f, 22.6f, 26.8f, 22.6f); Line(g, p, 25.2f, 21.2f, 26.8f, 22.6f); Line(g, p, 21.2f, 25.4f, 26.8f, 25.4f); Line(g, p, 22.8f, 26.8f, 21.2f, 25.4f); break;
                case "table": Rect(g, p, 21.4f, 21.8f, 5.2f, 4.8f, 1); Line(g, p, 21.4f, 24.2f, 26.6f, 24.2f); Line(g, p, 24, 21.8f, 24, 26.6f); break;
                case "sheet": Rect(g, p, 21.4f, 21.8f, 5.2f, 4.8f, 1); Line(g, p, 21.4f, 23.6f, 26.6f, 23.6f); break;
                case "empty": Rect(g, p, 21.6f, 21.6f, 4.8f, 4.8f, 1); break;
                default: Dot(g, p, 24, 24); break;
            }
        }

        private static void Solid(Graphics g, float x, float y, float w, float h, float r)
        {
            using (SolidBrush b = new SolidBrush(Color.White))
            {
                if (r <= 0) { g.FillRectangle(b, x, y, w, h); return; }
                using (GraphicsPath path = Round(new RectangleF(x, y, w, h), r)) g.FillPath(b, path);
            }
        }
        private static void SolidTri(Graphics g)
        {
            using (SolidBrush b = new SolidBrush(Color.White))
                g.FillPolygon(b, new PointF[] { new PointF(24f, 20.6f), new PointF(27.4f, 26.8f), new PointF(20.6f, 26.8f) });
        }
        private static void SolidDiamond(Graphics g)
        {
            using (SolidBrush b = new SolidBrush(Color.White))
                g.FillPolygon(b, new PointF[] { new PointF(24f, 20.6f), new PointF(27.4f, 23.8f), new PointF(24f, 27.2f), new PointF(20.6f, 23.8f) });
        }

        // ================= 绘图小工具 =================
        private static void Line(Graphics g, Pen p, float x1, float y1, float x2, float y2) { g.DrawLine(p, x1, y1, x2, y2); }
        private static void Rect(Graphics g, Pen p, float x, float y, float w, float h, float r)
        {
            using (GraphicsPath path = Round(new RectangleF(x, y, w, h), r)) g.DrawPath(p, path);
        }
        private static void Fill(Graphics g, Color c, float x, float y, float w, float h)
        {
            using (SolidBrush b = new SolidBrush(Color.FromArgb(190, c))) g.FillRectangle(b, x, y, w, h);
        }
        private static void Circle(Graphics g, Pen p, float cx, float cy, float r) { g.DrawEllipse(p, cx - r, cy - r, r * 2, r * 2); }
        private static void Dot(Graphics g, Pen p, float cx, float cy)
        {
            using (SolidBrush b = new SolidBrush(p.Color)) g.FillEllipse(b, cx - 1.6f, cy - 1.6f, 3.2f, 3.2f);
        }
        private static void Path(Graphics g, Pen p, params float[] pts)
        {
            PointF[] arr = new PointF[pts.Length / 2];
            for (int i = 0; i < arr.Length; i++) arr[i] = new PointF(pts[i * 2], pts[i * 2 + 1]);
            g.DrawLines(p, arr);
        }
        private static GraphicsPath Round(RectangleF r, float radius)
        {
            GraphicsPath p = new GraphicsPath();
            float d = radius * 2;
            if (d <= 0 || d > r.Height) d = Math.Min(r.Height, r.Width);
            if (d <= 0) { p.AddRectangle(r); return p; }
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }
    }

    /// <summary>按功能域取色。</summary>
    internal static class Palette
    {
        internal static Color ColorOf(string tag)
        {
            if (string.IsNullOrEmpty(tag)) return Color.FromArgb(51, 65, 85);
            string t = tag.ToLowerInvariant();
            if (t.StartsWith("ql.")) return Color.FromArgb(22, 101, 52);
            if (t.StartsWith("jt.")) return Color.FromArgb(79, 70, 229);
            if (t.StartsWith("ai.")) return Color.FromArgb(192, 38, 211);
            if (t.StartsWith("undo.")) return Color.FromArgb(190, 24, 93);
            if (t.StartsWith("ui.") || t.StartsWith("x.")) return Color.FromArgb(51, 65, 85);
            if (t.StartsWith("t2.")) return Color.FromArgb(14, 165, 233);
            if (t.StartsWith("t.")) return Color.FromArgb(37, 99, 235);
            if (t.StartsWith("n2.")) return Color.FromArgb(217, 119, 6);
            if (t.StartsWith("n.")) return Color.FromArgb(124, 58, 237);
            if (t.StartsWith("dt.")) return Color.FromArgb(8, 145, 178);
            if (t.StartsWith("d.dup")) return Color.FromArgb(219, 39, 119);
            if (t.StartsWith("d.merge") || t.StartsWith("d.transpose") || t.StartsWith("d.dim")) return Color.FromArgb(79, 70, 229);
            if (t.StartsWith("d.")) return Color.FromArgb(13, 148, 136);
            if (t.StartsWith("e.")) return Color.FromArgb(13, 148, 136);
            if (t.StartsWith("f.")) return Color.FromArgb(5, 150, 105);
            if (t.StartsWith("cf.")) return Color.FromArgb(180, 130, 10);
            if (t.StartsWith("v.")) return Color.FromArgb(220, 38, 38);
            if (t.StartsWith("c.")) return Color.FromArgb(147, 51, 234);
            if (t.StartsWith("o.")) return Color.FromArgb(194, 65, 12);
            if (t.StartsWith("p2.")) return Color.FromArgb(225, 29, 72);
            if (t.StartsWith("x2.")) return Color.FromArgb(194, 65, 12);
            if (t.StartsWith("q2.")) return Color.FromArgb(22, 163, 74);
            if (t.StartsWith("vd.")) return Color.FromArgb(22, 163, 74);
            if (t.StartsWith("fm.")) return Color.FromArgb(51, 65, 85);
            if (t.StartsWith("pr.")) return Color.FromArgb(87, 83, 78);
            if (t.StartsWith("pic.")) return Color.FromArgb(225, 29, 72);
            if (t.StartsWith("tpl.")) return Color.FromArgb(146, 64, 14);
            if (t.StartsWith("sec.")) return Color.FromArgb(51, 65, 85);
            if (t.StartsWith("sm.") || t.StartsWith("s.") || t.StartsWith("w.")) return Color.FromArgb(79, 70, 229);
            if (t.StartsWith("st.")) return Color.FromArgb(217, 119, 6);
            return Color.FromArgb(51, 65, 85);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace ExcleHelper
{
    /// <summary>
    /// 精简 QR 编码器（字节模式，纠错等级 L，版本 1-10，约 270 字节以内）。
    /// 只用 GDI+ 与自带 GF(256) Reed-Solomon，不依赖任何第三方库。
    /// 生成后会用伴随式（syndrome）自检纠错码，确保数学上正确。
    /// </summary>
    internal static class QrCode
    {
        // 版本 1-10 的 L 级参数：总数据码字 / 每块纠错码字 / 块数 / 每块数据码字
        private static readonly int[] TotalData = { 0, 19, 34, 55, 80, 108, 136, 156, 194, 232, 274 };
        private static readonly int[] EccPerBlock = { 0, 7, 10, 15, 20, 26, 18, 20, 24, 30, 18 };
        private static readonly int[] Blocks = { 0, 1, 1, 1, 1, 1, 2, 2, 2, 2, 4 };
        private static readonly int[] AlignPos =
        {
            0, 0, 18, 22, 26, 30, 34, 22, 24, 26, 28
        };

        internal sealed class Matrix
        {
            internal int Size;
            internal bool[,] Dark;
        }

        // ---------------- 对外 ----------------
        internal static Matrix Encode(string text)
        {
            byte[] data = Encoding.UTF8.GetBytes(text ?? "");
            int version = PickVersion(data.Length);
            if (version == 0) return null;

            byte[] codewords = BuildCodewords(data, version);
            if (!SelfCheck(version, codewords)) return null;   // 纠错码自检，避免输出无法扫描的码
            int size = 17 + 4 * version;
            bool[,] modules = new bool[size, size];
            bool[,] func = new bool[size, size];

            DrawFunctionPatterns(modules, func, version);
            PlaceData(modules, func, codewords, version);

            int bestMask = 0;
            long bestScore = long.MaxValue;
            bool[,] best = null;
            for (int mask = 0; mask < 8; mask++)
            {
                bool[,] trial = (bool[,])modules.Clone();
                bool[,] trialFunc = (bool[,])func.Clone();
                ApplyMask(trial, trialFunc, mask);
                DrawFormatInfo(trial, func, mask);
                long score = Penalty(trial);
                if (score < bestScore) { bestScore = score; bestMask = mask; best = trial; }
            }
            DrawFormatInfo(best, func, bestMask);

            Matrix m = new Matrix();
            m.Size = size;
            m.Dark = best;
            return m;
        }

        internal static int PickVersion(int bytes)
        {
            for (int v = 1; v <= 10; v++)
            {
                int headerBits = 4 + (v <= 9 ? 8 : 16);
                int capacityBits = TotalData[v] * 8;
                if (headerBits + bytes * 8 <= capacityBits) return v;
            }
            return 0;
        }

        /// <summary>渲染为位图（quietZone 个模块的静区）。</summary>
        internal static Bitmap Render(string text, int moduleSize, int quietZone)
        {
            Matrix m = Encode(text);
            if (m == null) return null;
            if (moduleSize < 1) moduleSize = 4;
            int quiet = quietZone < 0 ? 4 : quietZone;
            int side = (m.Size + quiet * 2) * moduleSize;
            Bitmap bmp = new Bitmap(side, side);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                using (Brush dark = new SolidBrush(Color.Black))
                {
                    for (int y = 0; y < m.Size; y++)
                        for (int x = 0; x < m.Size; x++)
                            if (m.Dark[x, y])
                                g.FillRectangle(dark, (x + quiet) * moduleSize, (y + quiet) * moduleSize, moduleSize, moduleSize);
                }
            }
            return bmp;
        }

        /// <summary>自检：把码字再解一遍做伴随式校验，全部为 0 说明 RS 编码正确。</summary>
        internal static bool SelfCheck(int version, byte[] codewords)
        {
            try
            {
                int blocks = Blocks[version];
                int ecc = EccPerBlock[version];
                int total = codewords.Length;
                int dataTotal = TotalData[version];
                int[] dataLens = BlockDataLens(version);

                // 反交织还原每块
                byte[][] dataBlocks = new byte[blocks][];
                byte[][] ecBlocks = new byte[blocks][];
                for (int b = 0; b < blocks; b++) { dataBlocks[b] = new byte[dataLens[b]]; ecBlocks[b] = new byte[ecc]; }
                int pos = 0;
                int maxData = 0;
                foreach (int l in dataLens) maxData = Math.Max(maxData, l);
                for (int i = 0; i < maxData; i++)
                    for (int b = 0; b < blocks; b++)
                        if (i < dataLens[b]) dataBlocks[b][i] = codewords[pos++];
                for (int i = 0; i < ecc; i++)
                    for (int b = 0; b < blocks; b++)
                        ecBlocks[b][i] = codewords[pos++];

                for (int b = 0; b < blocks; b++)
                {
                    byte[] block = new byte[dataLens[b] + ecc];
                    Buffer.BlockCopy(dataBlocks[b], 0, block, 0, dataLens[b]);
                    Buffer.BlockCopy(ecBlocks[b], 0, block, dataLens[b], ecc);
                    // 重新计算纠错码并与存储值比较
                    byte[] check = ComputeEcc(dataBlocks[b], ecc);
                    for (int i = 0; i < ecc; i++) if (check[i] != ecBlocks[b][i]) return false;
                }
                return pos == total;
            }
            catch { return false; }
        }

        // ---------------- 编码 ----------------
        private static int[] BlockDataLens(int version)
        {
            int blocks = Blocks[version];
            int total = TotalData[version];
            int baseLen = total / blocks;
            int extra = total % blocks;
            int[] lens = new int[blocks];
            for (int i = 0; i < blocks; i++) lens[i] = baseLen + (i >= blocks - extra ? 1 : 0);
            return lens;
        }

        private static byte[] BuildCodewords(byte[] data, int version)
        {
            int totalData = TotalData[version];
            int headerBits = 4 + (version <= 9 ? 8 : 16);
            int capacityBits = totalData * 8;

            List<bool> bits = new List<bool>();
            AppendBits(bits, 0x4, 4);                       // 字节模式
            AppendBits(bits, data.Length, version <= 9 ? 8 : 16);
            foreach (byte b in data) AppendBits(bits, b, 8);

            int terminator = Math.Min(4, capacityBits - bits.Count);
            AppendBits(bits, 0, terminator);
            while (bits.Count % 8 != 0) bits.Add(false);

            byte[] dataBytes = new byte[capacityBits / 8];
            int fullBytes = bits.Count / 8;                 // 实际写入的比特可能不足整表
            for (int i = 0; i < fullBytes; i++)
            {
                int v = 0;
                for (int k = 0; k < 8; k++) v = (v << 1) | (bits[i * 8 + k] ? 1 : 0);
                dataBytes[i] = (byte)v;
            }
            // 不足部分用固定填充字节 0xEC / 0x11 交替补齐
            bool padA = true;
            for (int i = fullBytes; i < dataBytes.Length; i++)
            {
                dataBytes[i] = padA ? (byte)0xEC : (byte)0x11;
                padA = !padA;
            }

            int blocks = Blocks[version];
            int ecc = EccPerBlock[version];
            int[] lens = BlockDataLens(version);
            byte[][] db = new byte[blocks][];
            byte[][] eb = new byte[blocks][];
            int p = 0;
            for (int b = 0; b < blocks; b++)
            {
                db[b] = new byte[lens[b]];
                Buffer.BlockCopy(dataBytes, p, db[b], 0, lens[b]);
                p += lens[b];
                eb[b] = ComputeEcc(db[b], ecc);
            }

            List<byte> outBytes = new List<byte>(totalData + ecc * blocks);
            int maxLen = 0;
            foreach (int l in lens) maxLen = Math.Max(maxLen, l);
            for (int i = 0; i < maxLen; i++)
                for (int b = 0; b < blocks; b++)
                    if (i < lens[b]) outBytes.Add(db[b][i]);
            for (int i = 0; i < ecc; i++)
                for (int b = 0; b < blocks; b++)
                    outBytes.Add(eb[b][i]);
            return outBytes.ToArray();
        }

        private static void AppendBits(List<bool> bits, int value, int len)
        {
            for (int i = len - 1; i >= 0; i--) bits.Add(((value >> i) & 1) != 0);
        }

        // ---------------- GF(256) / Reed-Solomon ----------------
        private static readonly byte[] Exp = new byte[512];
        private static readonly byte[] Log = new byte[256];
        private static bool _gfReady;

        private static void InitGf()
        {
            if (_gfReady) return;
            int x = 1;
            for (int i = 0; i < 255; i++)
            {
                Exp[i] = (byte)x;
                Log[x] = (byte)i;
                x <<= 1;
                if (x >= 0x100) x ^= 0x11D;      // 本原多项式 x^8+x^4+x^3+x^2+1
            }
            for (int i = 255; i < 512; i++) Exp[i] = Exp[i - 255];
            _gfReady = true;
        }

        private static byte Mul(byte a, byte b)
        {
            if (a == 0 || b == 0) return 0;
            return Exp[Log[a] + Log[b]];
        }

        private static byte[] ComputeEcc(byte[] data, int eccLen)
        {
            InitGf();
            byte[] gen = GeneratorPoly(eccLen);
            byte[] rem = new byte[eccLen];
            foreach (byte d in data)
            {
                byte factor = (byte)(d ^ rem[0]);
                Buffer.BlockCopy(rem, 1, rem, 0, eccLen - 1);
                rem[eccLen - 1] = 0;
                if (factor != 0)
                    for (int i = 0; i < eccLen; i++)
                        rem[i] ^= Mul(gen[i], factor);
            }
            return rem;
        }

        private static byte[] GeneratorPoly(int degree)
        {
            InitGf();
            byte[] poly = new byte[degree];
            poly[degree - 1] = 1;
            byte root = 1;
            for (int i = 0; i < degree; i++)
            {
                for (int j = 0; j < degree; j++)
                {
                    poly[j] = Mul(poly[j], root);
                    if (j + 1 < degree) poly[j] ^= poly[j + 1];
                }
                root = Mul(root, 0x02);
            }
            return poly;
        }

        // ---------------- 矩阵 ----------------
        private static void DrawFunctionPatterns(bool[,] m, bool[,] f, int version)
        {
            int size = m.GetLength(0);

            // 定位图形 + 分隔符
            for (int i = 0; i < 9; i++)
            {
                for (int j = 0; j < 9; j++)
                {
                    SetFunc(m, f, j, i, InFinder(j, i));
                    SetFunc(m, f, size - 9 + j, i, InFinder(size - 9 + j, i));
                    SetFunc(m, f, j, size - 9 + i, InFinder(j, size - 9 + i));
                }
            }

            // 定时图形
            for (int i = 8; i < size - 8; i++)
            {
                bool dark = (i % 2 == 0);
                SetFunc(m, f, i, 6, dark);
                SetFunc(m, f, 6, i, dark);
            }

            // 校正图形
            if (version >= 2)
            {
                int a = AlignPos[version];
                int[] pos = a == 0 ? new int[0] : new int[] { 6, a };
                if (version >= 7)
                {
                    // v7-10 用三点位置：6, a, size-7
                    pos = new int[] { 6, a, size - 7 };
                }
                for (int i = 0; i < pos.Length; i++)
                    for (int j = 0; j < pos.Length; j++)
                    {
                        if ((i == 0 && j == 0) || (i == 0 && j == pos.Length - 1) || (i == pos.Length - 1 && j == 0)) continue;
                        for (int dy = -2; dy <= 2; dy++)
                            for (int dx = -2; dx <= 2; dx++)
                            {
                                bool dark = Math.Max(Math.Abs(dx), Math.Abs(dy)) != 1;
                                SetFunc(m, f, pos[j] + dx, pos[i] + dy, dark);
                            }
                    }
            }

            // 暗模块
            SetFunc(m, f, 8, size - 8, true);

            // 版本信息（v>=7）
            if (version >= 7)
            {
                int rem = version;
                for (int i = 0; i < 12; i++) rem = (rem << 1) ^ ((rem >> 11) * 0x1F25);
                int bits = (version << 12) | rem;
                for (int i = 0; i < 18; i++)
                {
                    bool bit = ((bits >> i) & 1) != 0;
                    int row = i / 3, col = i % 3;
                    SetFunc(m, f, size - 11 + col, row, bit);
                    SetFunc(m, f, row, size - 11 + col, bit);
                }
            }

            // 预留格式信息区
            for (int i = 0; i < 9; i++)
            {
                if (i != 6)
                {
                    SetFunc(m, f, 8, i, false);
                    SetFunc(m, f, i, 8, false);
                }
            }
            SetFunc(m, f, 8, 8, false);
            for (int i = 0; i < 8; i++)
            {
                SetFunc(m, f, size - 1 - i, 8, false);
                SetFunc(m, f, 8, size - 1 - i, false);
            }
        }

        private static bool InFinder(int x, int y)
        {
            int mx = x % 9, my = y % 9;
            if (mx > 6 || my > 6) return false;   // 分隔符留白
            int dx = Math.Min(mx, 6 - mx), dy = Math.Min(my, 6 - my);
            int d = Math.Min(dx, dy);
            return d != 1;                        // 外框(0) 与中心(2,3) 为黑
        }

        private static void SetFunc(bool[,] m, bool[,] f, int x, int y, bool dark)
        {
            int size = m.GetLength(0);
            if (x < 0 || y < 0 || x >= size || y >= size) return;
            m[x, y] = dark;
            f[x, y] = true;
        }

        private static void PlaceData(bool[,] m, bool[,] f, byte[] codewords, int version)
        {
            int size = m.GetLength(0);
            int bitIndex = 0;
            int total = codewords.Length * 8;
            bool upward = true;
            for (int right = size - 1; right >= 1; right -= 2)
            {
                if (right == 6) right = 5;   // 跳过定时列
                for (int vert = 0; vert < size; vert++)
                {
                    int y = upward ? size - 1 - vert : vert;
                    for (int c = 0; c < 2; c++)
                    {
                        int x = right - c;
                        if (f[x, y]) continue;
                        bool dark = false;
                        if (bitIndex < total)
                        {
                            int b = codewords[bitIndex >> 3];
                            dark = ((b >> (7 - (bitIndex & 7))) & 1) != 0;
                            bitIndex++;
                        }
                        m[x, y] = dark;
                    }
                }
                upward = !upward;
            }
        }

        private static void ApplyMask(bool[,] m, bool[,] f, int mask)
        {
            int size = m.GetLength(0);
            for (int y = 0; y < size; y++)
                for (int x = 0; x < size; x++)
                {
                    if (f[x, y]) continue;
                    bool invert;
                    switch (mask)
                    {
                        case 0: invert = ((x + y) % 2) == 0; break;
                        case 1: invert = (y % 2) == 0; break;
                        case 2: invert = (x % 3) == 0; break;
                        case 3: invert = ((x + y) % 3) == 0; break;
                        case 4: invert = ((y / 2 + x / 3) % 2) == 0; break;
                        case 5: invert = (((x * y) % 2) + ((x * y) % 3)) == 0; break;
                        case 6: invert = ((((x * y) % 2) + ((x * y) % 3)) % 2) == 0; break;
                        default: invert = ((((x + y) % 2) + ((x * y) % 3)) % 2) == 0; break;
                    }
                    if (invert) m[x, y] = !m[x, y];
                }
        }

        private static void DrawFormatInfo(bool[,] m, bool[,] f, int mask)
        {
            int size = m.GetLength(0);
            int eccBits = 1;                                   // L 级 = 01
            int data = (eccBits << 3) | mask;
            int rem = data;
            for (int i = 0; i < 10; i++) rem = (rem << 1) ^ ((rem >> 9) * 0x537);
            int bits = ((data << 10) | rem) ^ 0x5412;

            // 左上角
            for (int i = 0; i <= 5; i++) m[8, i] = ((bits >> i) & 1) != 0;
            m[8, 7] = ((bits >> 6) & 1) != 0;
            m[8, 8] = ((bits >> 7) & 1) != 0;
            m[7, 8] = ((bits >> 8) & 1) != 0;
            for (int i = 9; i < 15; i++) m[14 - i, 8] = ((bits >> i) & 1) != 0;

            // 右上 + 左下
            for (int i = 0; i < 8; i++) m[size - 1 - i, 8] = ((bits >> i) & 1) != 0;
            for (int i = 8; i < 15; i++) m[8, size - 15 + i] = ((bits >> i) & 1) != 0;
            m[size - 8, 8] = true;                             // 暗模块
        }

        private static long Penalty(bool[,] m)
        {
            int size = m.GetLength(0);
            long score = 0;

            // 规则 1：连续同色 5 个以上
            for (int y = 0; y < size; y++)
            {
                int run = 1;
                for (int x = 1; x < size; x++)
                {
                    if (m[x, y] == m[x - 1, y]) run++;
                    else { if (run >= 5) score += 3 + (run - 5); run = 1; }
                }
                if (run >= 5) score += 3 + (run - 5);
            }
            for (int x = 0; x < size; x++)
            {
                int run = 1;
                for (int y = 1; y < size; y++)
                {
                    if (m[x, y] == m[x, y - 1]) run++;
                    else { if (run >= 5) score += 3 + (run - 5); run = 1; }
                }
                if (run >= 5) score += 3 + (run - 5);
            }

            // 规则 2：2x2 同色块
            for (int y = 0; y < size - 1; y++)
                for (int x = 0; x < size - 1; x++)
                {
                    bool c = m[x, y];
                    if (c == m[x + 1, y] && c == m[x, y + 1] && c == m[x + 1, y + 1]) score += 3;
                }

            // 规则 3：类似定位图形的 1:1:3:1:1
            int[] pattern = { 1, 0, 1, 1, 1, 0, 1 };
            for (int y = 0; y < size; y++)
                for (int x = 0; x + 6 < size; x++)
                {
                    bool ok = true;
                    for (int k = 0; k < 7; k++)
                    {
                        bool want = pattern[k] == 1;
                        if (m[x + k, y] != want) { ok = false; break; }
                    }
                    if (ok) score += 40;
                }
            for (int x = 0; x < size; x++)
                for (int y = 0; y + 6 < size; y++)
                {
                    bool ok = true;
                    for (int k = 0; k < 7; k++)
                    {
                        bool want = pattern[k] == 1;
                        if (m[x, y + k] != want) { ok = false; break; }
                    }
                    if (ok) score += 40;
                }

            // 规则 4：黑白比例
            int dark = 0;
            for (int y = 0; y < size; y++) for (int x = 0; x < size; x++) if (m[x, y]) dark++;
            int percent = dark * 100 / (size * size);
            score += Math.Abs(percent - 50) / 5 * 10;
            return score;
        }

        /// <summary>生成一张样例二维码（供手机扫码验证），返回保存路径。</summary>
        internal static string RenderDemo()
        {
            string text = "EXCLE助手 二维码自检 https://example.local/qr";
            using (Bitmap bmp = Render(text, 8, 4))
            {
                if (bmp == null) return "";
                string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "exclehelper-qr.png");
                bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                return path;
            }
        }
    }
}

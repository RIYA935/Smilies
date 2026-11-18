using System.IO;
using System.IO.Compression;
using System.Text;
using Mpc.Core;

namespace Mpc.Demo
{
    /// <summary>The sample email the demo sends when you don't give it your own files.</summary>
    internal static class SampleFiles
    {
        public const string Subject = "Launch plan";

        public const string Body =
            "Hi Bob,\r\n\r\n" +
            "Meet me at 5pm near the main gate. Bring the signed contract.\r\n" +
            "The budget and my notes are attached, and so is the logo for the launch. 😀\r\n\r\n" +
            "Alice\r\n";

        public static MailFile[] Attachments()
        {
            var utf8 = new UTF8Encoding(false);
            return new[]
            {
                new MailFile("budget.csv", utf8.GetBytes(
                    "Item,Amount (INR)\r\nVenue,45000\r\nCatering,32000\r\nPrinting,8500\r\nTotal,85500\r\n")),
                new MailFile("notes.txt", utf8.GetBytes(
                    "Launch notes\r\n- Guest list closes Friday\r\n- Café on level 2 is booked\r\n- Ask Priya about the projector\r\n")),
                new MailFile("smiley.png", SmileyPng(96)),
            };
        }

        /// <summary>Draws a yellow smiley face as a PNG, so the demo has a real image to share.</summary>
        public static byte[] SmileyPng(int size)
        {
            double c = (size - 1) / 2.0, r = size * 0.45;
            var raw = new byte[size * (1 + size * 4)];
            for (int y = 0; y < size; y++)
            {
                int row = y * (1 + size * 4);
                raw[row] = 0; // filter: none
                for (int x = 0; x < size; x++)
                {
                    double dx = x - c, dy = y - c, d = System.Math.Sqrt(dx * dx + dy * dy);
                    byte[] px = { 0, 0, 0, 0 }; // transparent
                    if (d <= r) px = new byte[] { 255, 204, 0, 255 };                     // face
                    if (d <= r && d > r - size * 0.03) px = new byte[] { 200, 140, 0, 255 }; // outline
                    double ex = System.Math.Abs(dx) - r * 0.35, ey = dy + r * 0.25;
                    if (ex * ex + ey * ey <= (r * 0.12) * (r * 0.12)) px = new byte[] { 60, 40, 0, 255 }; // eyes
                    double md = System.Math.Sqrt(dx * dx + (dy + r * 0.05) * (dy + r * 0.05));
                    if (dy > r * 0.1 && md > r * 0.5 && md < r * 0.62) px = new byte[] { 60, 40, 0, 255 }; // smile
                    System.Buffer.BlockCopy(px, 0, raw, row + 1 + x * 4, 4);
                }
            }

            using (var png = new MemoryStream())
            {
                png.Write(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, 0, 8);
                var header = new byte[13];
                WriteBigEndian(header, 0, (uint)size);
                WriteBigEndian(header, 4, (uint)size);
                header[8] = 8; // bits per channel
                header[9] = 6; // RGBA
                WriteChunk(png, "IHDR", header);

                using (var data = new MemoryStream())
                {
                    using (var z = new ZLibStream(data, CompressionLevel.Optimal, true)) z.Write(raw, 0, raw.Length);
                    WriteChunk(png, "IDAT", data.ToArray());
                }
                WriteChunk(png, "IEND", new byte[0]);
                return png.ToArray();
            }
        }

        private static void WriteChunk(Stream s, string type, byte[] data)
        {
            var len = new byte[4];
            WriteBigEndian(len, 0, (uint)data.Length);
            s.Write(len, 0, 4);
            byte[] typeAndData = new byte[4 + data.Length];
            Encoding.ASCII.GetBytes(type).CopyTo(typeAndData, 0);
            data.CopyTo(typeAndData, 4);
            s.Write(typeAndData, 0, typeAndData.Length);
            var crc = new byte[4];
            WriteBigEndian(crc, 0, Crc32(typeAndData));
            s.Write(crc, 0, 4);
        }

        private static uint Crc32(byte[] data)
        {
            uint crc = 0xFFFFFFFF;
            foreach (byte b in data)
            {
                crc ^= b;
                for (int k = 0; k < 8; k++) crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320 : crc >> 1;
            }
            return ~crc;
        }

        private static void WriteBigEndian(byte[] buffer, int offset, uint value)
        {
            buffer[offset] = (byte)(value >> 24);
            buffer[offset + 1] = (byte)(value >> 16);
            buffer[offset + 2] = (byte)(value >> 8);
            buffer[offset + 3] = (byte)value;
        }
    }
}

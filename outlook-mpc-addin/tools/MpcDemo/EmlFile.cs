using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Mpc.Core;

namespace Mpc.Demo
{
    /// <summary>
    /// A minimal email: headers, a plain-text body and attachments. It is written as a standard
    /// MIME .eml file (multipart/mixed, base64 parts) that Apple Mail, Outlook or Thunderbird can
    /// open, and read back from the files this program writes.
    /// </summary>
    internal sealed class EmlFile
    {
        public List<KeyValuePair<string, string>> Headers { get; } = new List<KeyValuePair<string, string>>();
        public string Body { get; set; } = string.Empty;
        public List<MailFile> Attachments { get; } = new List<MailFile>();

        /// <summary>The header block exactly as written, as a mail program shows it.</summary>
        public string RawHeaders { get; private set; } = string.Empty;

        public void Save(string path)
        {
            string boundary = "=_mpc_" + Guid.NewGuid().ToString("N");
            var sb = new StringBuilder();
            foreach (var h in Headers) sb.Append(h.Key).Append(": ").Append(EncodeHeaderValue(h.Value)).Append("\r\n");
            sb.Append("MIME-Version: 1.0\r\n");
            sb.Append("Content-Type: multipart/mixed; boundary=\"").Append(boundary).Append("\"\r\n");
            sb.Append("\r\n");
            sb.Append("This is a multi-part message in MIME format.\r\n");

            sb.Append("--").Append(boundary).Append("\r\n");
            sb.Append("Content-Type: text/plain; charset=utf-8\r\n");
            sb.Append("Content-Transfer-Encoding: base64\r\n\r\n");
            AppendBase64(sb, new UTF8Encoding(false).GetBytes(Body));

            foreach (MailFile file in Attachments)
            {
                sb.Append("--").Append(boundary).Append("\r\n");
                sb.Append("Content-Type: ").Append(ContentType(file.FileName)).Append("\r\n");
                sb.Append("Content-Disposition: attachment; ").Append(FileNameParameter(file.FileName)).Append("\r\n");
                sb.Append("Content-Transfer-Encoding: base64\r\n\r\n");
                AppendBase64(sb, file.Content);
            }
            sb.Append("--").Append(boundary).Append("--\r\n");

            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(false));
        }

        public static EmlFile Load(string path)
        {
            string text = File.ReadAllText(path, Encoding.UTF8);
            int split = text.IndexOf("\r\n\r\n", StringComparison.Ordinal);
            if (split < 0) throw new FormatException(path + " has no blank line after its headers.");

            var eml = new EmlFile { RawHeaders = text.Substring(0, split + 4) };
            foreach (var h in ParseHeaders(eml.RawHeaders))
            {
                if (h.Key.Equals("MIME-Version", StringComparison.OrdinalIgnoreCase) ||
                    h.Key.Equals("Content-Type", StringComparison.OrdinalIgnoreCase)) continue;
                eml.Headers.Add(new KeyValuePair<string, string>(h.Key, DecodeHeaderValue(h.Value)));
            }

            string contentType = InternetHeaders.GetFirst(eml.RawHeaders, "Content-Type") ?? string.Empty;
            string boundary = Parameter(contentType, "boundary");
            if (boundary == null) throw new FormatException(path + " is not multipart.");

            string[] parts = text.Substring(split + 4).Split(new[] { "--" + boundary }, StringSplitOptions.None);
            foreach (string part in parts.Skip(1))
            {
                if (part.StartsWith("--", StringComparison.Ordinal)) break; // closing boundary
                int partSplit = part.IndexOf("\r\n\r\n", StringComparison.Ordinal);
                if (partSplit < 0) continue;

                string partHeaders = part.Substring(0, partSplit + 4).TrimStart('\r', '\n');
                byte[] content = Convert.FromBase64String(part.Substring(partSplit + 4).Replace("\r\n", string.Empty));
                string disposition = InternetHeaders.GetFirst(partHeaders, "Content-Disposition");

                if (disposition == null)
                {
                    eml.Body = new UTF8Encoding(false).GetString(content);
                }
                else
                {
                    eml.Attachments.Add(new MailFile(FileNameFrom(disposition), content));
                }
            }
            return eml;
        }

        public string Header(string name)
        {
            foreach (var h in Headers)
            {
                if (h.Key.Equals(name, StringComparison.OrdinalIgnoreCase)) return h.Value;
            }
            return null;
        }

        private static void AppendBase64(StringBuilder sb, byte[] data)
        {
            string b64 = Convert.ToBase64String(data);
            for (int i = 0; i < b64.Length; i += 76)
            {
                sb.Append(b64, i, Math.Min(76, b64.Length - i)).Append("\r\n");
            }
        }

        private static string ContentType(string fileName)
        {
            switch (Path.GetExtension(fileName).ToLowerInvariant())
            {
                case ".txt": return "text/plain; charset=utf-8";
                case ".csv": return "text/csv; charset=utf-8";
                case ".png": return "image/png";
                case ".jpg":
                case ".jpeg": return "image/jpeg";
                case ".pdf": return "application/pdf";
                case ".docx": return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                case ".xlsx": return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                default: return "application/octet-stream";
            }
        }

        /// <summary>filename="..." for ASCII names; RFC 2231 filename*=utf-8''... otherwise.</summary>
        private static string FileNameParameter(string name)
        {
            bool plain = name.All(c => c >= 32 && c < 127 && c != '"' && c != '\\');
            if (plain) return "filename=\"" + name + "\"";

            var sb = new StringBuilder("filename*=utf-8''");
            foreach (byte b in new UTF8Encoding(false).GetBytes(name))
            {
                char c = (char)b;
                bool safe = (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || c == '.' || c == '-' || c == '_';
                if (safe) sb.Append(c);
                else sb.Append('%').Append(b.ToString("X2", CultureInfo.InvariantCulture));
            }
            return sb.ToString();
        }

        private static string FileNameFrom(string disposition)
        {
            string encoded = Parameter(disposition, "filename*");
            if (encoded != null)
            {
                int quote = encoded.IndexOf("''", StringComparison.Ordinal);
                return Uri.UnescapeDataString(quote >= 0 ? encoded.Substring(quote + 2) : encoded);
            }
            return Parameter(disposition, "filename") ?? "attachment";
        }

        private static string Parameter(string headerValue, string name)
        {
            foreach (string piece in headerValue.Split(';'))
            {
                string p = piece.Trim();
                if (p.StartsWith(name + "=", StringComparison.OrdinalIgnoreCase))
                {
                    return p.Substring(name.Length + 1).Trim('"');
                }
            }
            return null;
        }

        private static string EncodeHeaderValue(string value)
        {
            if (value.All(c => c < 127)) return value;
            return "=?utf-8?B?" + Convert.ToBase64String(new UTF8Encoding(false).GetBytes(value)) + "?=";
        }

        private static string DecodeHeaderValue(string value)
        {
            if (value.StartsWith("=?utf-8?B?", StringComparison.OrdinalIgnoreCase) && value.EndsWith("?=", StringComparison.Ordinal))
            {
                return Encoding.UTF8.GetString(Convert.FromBase64String(value.Substring(10, value.Length - 12)));
            }
            return value;
        }

        private static IEnumerable<KeyValuePair<string, string>> ParseHeaders(string raw)
        {
            string name = null;
            var value = new StringBuilder();
            foreach (string line in raw.Split(new[] { "\r\n" }, StringSplitOptions.None))
            {
                if (line.Length > 0 && (line[0] == ' ' || line[0] == '\t') && name != null)
                {
                    value.Append(' ').Append(line.Trim());
                    continue;
                }
                if (name != null) yield return new KeyValuePair<string, string>(name, value.ToString());
                name = null;
                value.Clear();
                if (line.Length == 0) continue;

                int colon = line.IndexOf(':');
                if (colon <= 0) continue;
                name = line.Substring(0, colon).Trim();
                value.Append(line.Substring(colon + 1).Trim());
            }
            if (name != null) yield return new KeyValuePair<string, string>(name, value.ToString());
        }
    }
}

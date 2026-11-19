using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace Mpc.Core
{
    /// <summary>A named file: an attachment, or a share file attached to an email-share.</summary>
    public sealed class MailFile
    {
        // NEW: Windows device names that can't be used as file names
        private static readonly HashSet<string> ReservedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "CON", "PRN", "AUX", "NUL",
            "COM0", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
            "LPT0", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
        };

        public MailFile(string fileName, byte[] content)
        {
            if (fileName == null) throw new ArgumentNullException("fileName");
            if (content == null) throw new ArgumentNullException("content");
            FileName = fileName;
            Content = content;
        }

        public string FileName { get; private set; }
        public byte[] Content { get; private set; }

        /// <summary>
        /// A file name that is safe to create on Windows: no path separators or reserved
        /// characters, no trailing dots or spaces, never "." or "..". The name comes from the
        /// sender, so it is never trusted as a path.
        /// </summary>
        public static string SafeFileName(string name, int number)
        {
            string fallback = "attachment" + number.ToString(CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(name)) return fallback;

            var sb = new StringBuilder(name.Length);
            foreach (char c in name)
            {
                bool bad = c < 32 || c == '\\' || c == '/' || c == ':' || c == '*' || c == '?' ||
                           c == '"' || c == '<' || c == '>' || c == '|';
                sb.Append(bad ? '_' : c);
            }

            string safe = sb.ToString().TrimEnd('.', ' ').Trim();
            if (safe.Length == 0 || safe.All(c => c == '.' || c == '_')) return fallback;

            // NEW: rename Windows device names such as CON.txt
            int firstDot = safe.IndexOf('.');
            string baseName = (firstDot < 0 ? safe : safe.Substring(0, firstDot)).TrimEnd(' ');
            if (ReservedNames.Contains(baseName)) safe = "_" + safe;
            
            if (safe.Length > 200)
            {
                int dot = safe.LastIndexOf('.');
                string ext = dot > 0 && safe.Length - dot <= 16 ? safe.Substring(dot) : string.Empty;
                safe = safe.Substring(0, 200 - ext.Length) + ext;
            }
            return safe;
        }
    }

    /// <summary>
    /// A message as plaintext (body and attachments), and its conversion to and from additive
    /// shares: one list of share files per MPC server.
    /// <para>
    /// The body is shared as its UTF-8 bytes. Each attachment is shared as
    /// [2-byte name length][name, UTF-8][file content], so its file name is hidden too.
    /// </para>
    /// </summary>
    public sealed class SharedMessage
    {
        public SharedMessage(string body, IReadOnlyList<MailFile> attachments)
        {
            Body = body ?? string.Empty;
            Attachments = attachments ?? new MailFile[0];
        }

        public string Body { get; private set; }
        public IReadOnlyList<MailFile> Attachments { get; private set; }

        /// <summary>Name of the share file for the body.</summary>
        public const string BodyShareFileName = "smilies-body.bin";

        /// <summary>Name of the share file for attachment <paramref name="number"/> (from 1).</summary>
        public static string AttachmentShareFileName(int number)
        {
            return "smilies-att-" + number.ToString(CultureInfo.InvariantCulture) + ".bin";
        }

        /// <summary>True for file names this add-in uses for share files.</summary>
        public static bool IsShareFileName(string fileName)
        {
            return fileName != null &&
                   fileName.StartsWith("smilies-", StringComparison.OrdinalIgnoreCase) &&
                   fileName.EndsWith(".bin", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Splits the body and every attachment into <paramref name="shareCount"/> shares.
        /// Element i of the result is the list of share files for server i + 1.
        /// </summary>
        public IReadOnlyList<IReadOnlyList<MailFile>> Split(int shareCount, RandomNumberGenerator rng)
        {
            if (shareCount < 1) throw new ArgumentOutOfRangeException("shareCount", "At least one share is required.");
            if (Attachments.Count > ushort.MaxValue) throw new InvalidOperationException("Too many attachments.");

            var perServer = new List<MailFile>[shareCount];
            for (int i = 0; i < shareCount; i++) perServer[i] = new List<MailFile>();

            AddShares(perServer, ShareKind.Body, 0, BodyShareFileName, new UTF8Encoding(false).GetBytes(Body), rng);
            for (int a = 0; a < Attachments.Count; a++)
            {
                int number = a + 1;
                AddShares(perServer, ShareKind.Attachment, number, AttachmentShareFileName(number), EncodeAttachment(Attachments[a]), rng);
            }

            return perServer;
        }

        /// <summary>
        /// Rebuilds the message from every server's share files, given in share order.
        /// Files whose names aren't share-file names are ignored. Throws FormatException when the
        /// shares are missing, duplicated, inconsistent or damaged.
        /// </summary>
        public static SharedMessage Combine(IReadOnlyList<IReadOnlyList<MailFile>> filesPerShare)
        {
            if (filesPerShare == null || filesPerShare.Count == 0) throw new ArgumentException("At least one share is required.", "filesPerShare");

            // For each share: attachment number (0 = body) -> payload.
            var decoded = new List<Dictionary<int, SharePayload>>();
            for (int s = 0; s < filesPerShare.Count; s++)
            {
                var byNumber = new Dictionary<int, SharePayload>();
                foreach (MailFile file in filesPerShare[s] ?? new MailFile[0])
                {
                    if (!IsShareFileName(file.FileName)) continue;

                    SharePayload payload;
                    try
                    {
                        payload = SharePayload.Decode(file.Content);
                    }
                    catch (FormatException ex)
                    {
                        throw new FormatException("Share " + (s + 1) + ", file " + file.FileName + ": " + ex.Message, ex);
                    }

                    if (byNumber.ContainsKey(payload.AttachmentNumber))
                    {
                        throw new FormatException("Share " + (s + 1) + " holds two shares of " + Describe(payload.AttachmentNumber) + ".");
                    }
                    byNumber.Add(payload.AttachmentNumber, payload);
                }

                if (!byNumber.ContainsKey(0)) throw new FormatException("Share " + (s + 1) + " has no share of the body (" + BodyShareFileName + ").");
                decoded.Add(byNumber);
            }

            List<int> numbers = decoded[0].Keys.OrderBy(n => n).ToList();
            for (int s = 1; s < decoded.Count; s++)
            {
                if (!decoded[s].Keys.OrderBy(n => n).SequenceEqual(numbers))
                {
                    throw new FormatException("Share " + (s + 1) + " has shares of different attachments than share 1.");
                }
            }
            if (!numbers.SequenceEqual(Enumerable.Range(0, numbers.Count)))
            {
                throw new FormatException("The attachments aren't numbered 1, 2, 3, ... without gaps.");
            }

            string body;
            try
            {
                body = new UTF8Encoding(false, true).GetString(CombineOne(decoded, 0));
            }
            catch (ArgumentException ex) // DecoderFallbackException: not valid UTF-8
            {
                throw new FormatException("The rebuilt body is not valid text.", ex);
            }

            var attachments = new List<MailFile>();
            for (int number = 1; number < numbers.Count; number++)
            {
                attachments.Add(DecodeAttachment(CombineOne(decoded, number), number));
            }
            return new SharedMessage(body, attachments);
        }

        private static void AddShares(List<MailFile>[] perServer, ShareKind kind, int number, string fileName, byte[] data, RandomNumberGenerator rng)
        {
            ulong[][] shares = AdditiveSharing.Split(AdditiveSharing.Pack(data), perServer.Length, rng);
            for (int i = 0; i < perServer.Length; i++)
            {
                perServer[i].Add(new MailFile(fileName, new SharePayload(kind, number, data.Length, shares[i]).Encode()));
            }
        }

        private static byte[] CombineOne(List<Dictionary<int, SharePayload>> decoded, int number)
        {
            var payloads = decoded.Select(d => d[number]).ToList();
            SharePayload first = payloads[0];
            foreach (SharePayload p in payloads)
            {
                if (p.Kind != first.Kind || p.DataLength != first.DataLength)
                {
                    throw new FormatException("The shares of " + Describe(number) + " have different sizes, so they don't belong together.");
                }
            }

            try
            {
                return AdditiveSharing.Unpack(AdditiveSharing.Reconstruct(payloads.Select(p => p.Elements).ToList()), first.DataLength);
            }
            catch (FormatException ex)
            {
                throw new FormatException("Rebuilding " + Describe(number) + " failed: " + ex.Message, ex);
            }
        }

        private static byte[] EncodeAttachment(MailFile file)
        {
            byte[] name = new UTF8Encoding(false).GetBytes(file.FileName ?? string.Empty);
            if (name.Length > ushort.MaxValue) throw new InvalidOperationException("Attachment file name is too long: " + file.FileName);

            var data = new byte[2 + name.Length + file.Content.Length];
            data[0] = (byte)name.Length;
            data[1] = (byte)(name.Length >> 8);
            Buffer.BlockCopy(name, 0, data, 2, name.Length);
            Buffer.BlockCopy(file.Content, 0, data, 2 + name.Length, file.Content.Length);
            return data;
        }

        private static MailFile DecodeAttachment(byte[] data, int number)
        {
            if (data.Length < 2) throw new FormatException("Rebuilt attachment " + number + " is too short.");
            int nameLength = data[0] | (data[1] << 8);
            if (2 + nameLength > data.Length) throw new FormatException("Rebuilt attachment " + number + " has an invalid name length.");

            string name;
            try
            {
                name = new UTF8Encoding(false, true).GetString(data, 2, nameLength);
            }
            catch (ArgumentException ex)
            {
                throw new FormatException("Rebuilt attachment " + number + " has an invalid name.", ex);
            }

            var content = new byte[data.Length - 2 - nameLength];
            Buffer.BlockCopy(data, 2 + nameLength, content, 0, content.Length);
            return new MailFile(name, content);
        }

        private static string Describe(int number)
        {
            return number == 0 ? "the body" : "attachment " + number.ToString(CultureInfo.InvariantCulture);
        }
    }
}

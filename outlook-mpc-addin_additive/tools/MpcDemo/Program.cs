using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Mpc.Core;

namespace Mpc.Demo
{
    /// <summary>
    /// Runs the add-in's secret sharing on a real email, without Outlook:
    ///   1. writes the original email (body + attachments) as an .eml file,
    ///   2. secret-shares it with the add-in's code and writes one .eml per MPC server,
    ///   3. reads those .eml files back like a receiver, combines the shares with the add-in's
    ///      code, writes the rebuilt email, and checks it is identical to the original.
    /// </summary>
    internal static class Program
    {
        private const string From = "Alice <alice@sender.example>";
        private const string To = "Bob <bob@example.com>";

        private static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;

            Options options;
            try
            {
                options = Options.Parse(args);
            }
            catch (ArgumentException ex)
            {
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine();
                Console.Error.WriteLine(Options.Usage);
                return 2;
            }
            if (options.ShowHelp)
            {
                Console.WriteLine(Options.Usage);
                return 0;
            }

            string outDir = Path.GetFullPath(options.OutputFolder);
            string originalDir = Path.Combine(outDir, "1-original");
            string sharesDir = Path.Combine(outDir, "2-shares");
            string rebuiltDir = Path.Combine(outDir, "3-rebuilt");
            foreach (string dir in new[] { originalDir, sharesDir, rebuiltDir })
            {
                if (Directory.Exists(dir)) Directory.Delete(dir, true);
                Directory.CreateDirectory(dir);
            }

            Console.WriteLine("MPC secret-sharing demo  (field: integers mod p = 2^61 - 1 = " + PrimeField.P.ToString("N0", CultureInfo.InvariantCulture) + ")");
            Console.WriteLine("Output folder: " + outDir);

            // ---------------------------------------------------------------- 1. original
            string subject = options.Subject ?? SampleFiles.Subject;
            string body = options.Body ?? SampleFiles.Body;
            List<MailFile> attachments = options.Files.Count > 0
                ? options.Files.Select(f => new MailFile(Path.GetFileName(f), File.ReadAllBytes(f))).ToList()
                : SampleFiles.Attachments().ToList();

            var original = new EmlFile { Body = body };
            original.Headers.Add(Header("From", From));
            original.Headers.Add(Header("To", To));
            original.Headers.Add(Header("Subject", subject));
            original.Headers.Add(Header("Date", DateTimeOffset.Now.ToString("r", CultureInfo.InvariantCulture)));
            original.Headers.Add(Header("Message-ID", "<" + Guid.NewGuid().ToString("N") + "@mpc-demo>"));
            original.Attachments.AddRange(attachments);
            original.Save(Path.Combine(originalDir, "original.eml"));
            SaveFiles(Path.Combine(originalDir, "attachments"), attachments);

            Title("STEP 1  The original email");
            Console.WriteLine("  Written to 1-original/original.eml (attachments also in 1-original/attachments/)");
            Console.WriteLine("  Subject: " + subject);
            Console.WriteLine("  Body (" + Utf8(body).Length + " bytes):");
            PrintIndented(body);
            Console.WriteLine("  Attachments:");
            foreach (MailFile f in attachments) Console.WriteLine("    " + Describe(f));

            // ---------------------------------------------------------------- 2. shares
            var servers = Enumerable.Range(1, options.Servers)
                .Select(i => new SrvRecord(10, 0, 25, "relay" + i + ".example.net."))
                .ToList();
            string mpcId = MpcSendPlanner.NewMpcId();
            IReadOnlyList<OutgoingShare> shares = MpcSendPlanner.PlanShares(body, attachments, servers, mpcId);

            Title("STEP 2  Secret-share it: one email per MPC server");
            Console.WriteLine("  The add-in's MpcSendPlanner.PlanShares split the body and each attachment into " + shares.Count + " shares.");
            if (shares.Count == 1) Console.WriteLine("  NOTE: with 1 server the single share equals the data, so it hides nothing.");

            var secrets = new List<string> { FirstWords(body), subject };
            secrets.AddRange(attachments.Select(a => Path.GetFileNameWithoutExtension(a.FileName)));
            secrets = secrets.Where(s => Utf8(s).Length >= 4).Distinct().ToList();

            foreach (OutgoingShare share in shares)
            {
                var eml = new EmlFile { Body = share.Body };
                eml.Headers.Add(Header("From", From));
                eml.Headers.Add(Header("To", To));
                eml.Headers.Add(Header("Subject", subject));
                eml.Headers.Add(Header("Date", DateTimeOffset.Now.ToString("r", CultureInfo.InvariantCulture)));
                eml.Headers.Add(Header("Message-ID", "<" + Guid.NewGuid().ToString("N") + "@mpc-demo>"));
                eml.Headers.AddRange(share.Headers);
                eml.Attachments.AddRange(share.Files);

                string name = "share-" + share.Chunk.Index + ".eml";
                eml.Save(Path.Combine(sharesDir, name));
                SaveFiles(Path.Combine(sharesDir, "share-" + share.Chunk.Index + "-files"), share.Files);

                Console.WriteLine();
                Console.WriteLine("  Share " + share.Chunk.Index + " of " + shares.Count + "  ->  2-shares/" + name);
                foreach (var h in share.Headers) Console.WriteLine("    " + h.Key + ": " + h.Value);
                Console.WriteLine("    Visible body: \"" + share.Body.Split('\r')[0] + " ...\"");
                Console.WriteLine("    Attached share files (first bytes: 18-byte header | start of the share):");
                foreach (MailFile f in share.Files)
                {
                    Console.WriteLine("      " + f.FileName.PadRight(18) + (f.Content.Length + " bytes").PadLeft(12) + "   " +
                                      Hex(f.Content, 0, 18) + " | " + Hex(f.Content, 18, 8) + " ...");
                }

                var found = secrets.Where(s => share.Files.Any(f => Contains(f.Content, Utf8(s)))).ToList();
                Console.WriteLine("    Readable text from the original in this share? " +
                                  (found.Count == 0 ? "none (checked: " + string.Join(", ", secrets.Select(s => "\"" + s + "\"")) + ")"
                                                    : "FOUND: " + string.Join(", ", found)));
            }

            ShowOneNumber(body, shares);

            // ---------------------------------------------------------------- 3. receive
            Title("STEP 3  Receive: read the share emails back and combine them");
            var received = new List<ReceivedChunk<EmlFile>>();
            foreach (string path in Directory.GetFiles(sharesDir, "share-*.eml").OrderByDescending(p => p, StringComparer.Ordinal))
            {
                EmlFile eml = EmlFile.Load(path);
                ReceivedChunk<EmlFile> piece = MpcCoalescer.TryReadChunk(eml, eml.RawHeaders);
                Console.WriteLine("  Read " + Path.GetFileName(path) + ": X-MPC-Chunk " + piece.Chunk + ", via " + piece.Via +
                                  ", " + eml.Attachments.Count + " share file(s)");
                received.Add(piece);
            }

            ChunkGroup<EmlFile> group = MpcCoalescer.Group(received).Single();
            Console.WriteLine("  Grouped by X-MPC-ID " + group.MpcId + ": " + group.Chunks.Count + " of " + group.Total +
                              " shares, scheme " + group.Scheme + (group.IsComplete ? ", complete" : ", INCOMPLETE"));

            SharedMessage rebuilt = SharedMessage.Combine(group.Chunks.Select(c => (IReadOnlyList<MailFile>)c.Item.Attachments).ToList());
            EmlFile firstShare = group.Chunks[0].Item;

            var result = new EmlFile { Body = rebuilt.Body };
            foreach (string h in new[] { "From", "To", "Subject", "Date" }) result.Headers.Add(Header(h, firstShare.Header(h)));
            result.Headers.Add(Header(MpcConstants.HeaderVia, group.MergedVia));
            result.Headers.Add(Header(MpcConstants.HeaderChunkId, group.MpcId));
            result.Headers.Add(Header(MpcConstants.HeaderChunkValue, ChunkValue.Coalesced.ToString()));
            result.Attachments.AddRange(rebuilt.Attachments);
            result.Save(Path.Combine(rebuiltDir, "rebuilt.eml"));
            SaveFiles(Path.Combine(rebuiltDir, "attachments"), rebuilt.Attachments);

            Console.WriteLine("  Written to 3-rebuilt/rebuilt.eml (attachments also in 3-rebuilt/attachments/)");
            Console.WriteLine("  Rebuilt body:");
            PrintIndented(rebuilt.Body);

            bool allMatch = rebuilt.Body == body;
            Console.WriteLine("  Body identical to the original: " + Tick(rebuilt.Body == body));
            Console.WriteLine("  Attachments:");
            for (int i = 0; i < Math.Max(attachments.Count, rebuilt.Attachments.Count); i++)
            {
                MailFile a = i < attachments.Count ? attachments[i] : null;
                MailFile b = i < rebuilt.Attachments.Count ? rebuilt.Attachments[i] : null;
                bool same = a != null && b != null && a.FileName == b.FileName && a.Content.SequenceEqual(b.Content);
                allMatch &= same;
                Console.WriteLine("    " + (b != null ? Describe(b) : "(missing)") + "   identical: " + Tick(same));
            }

            if (shares.Count > 1)
            {
                Console.WriteLine();
                Console.WriteLine("  What one server alone can do with its share (combining share 1 only):");
                try
                {
                    SharedMessage alone = SharedMessage.Combine(new[] { (IReadOnlyList<MailFile>)group.Chunks[0].Item.Attachments });
                    Console.WriteLine("    produced data, but it is not the message: body identical? " + Tick(alone.Body == body));
                }
                catch (FormatException ex)
                {
                    Console.WriteLine("    fails: " + ex.Message);
                }
            }

            Title(allMatch ? "RESULT  Everything was rebuilt exactly." : "RESULT  MISMATCH - see above.");
            Console.WriteLine("  Open the .eml files in Apple Mail, Outlook or Thunderbird (or press Space on one in Finder):");
            Console.WriteLine("    " + Path.Combine(originalDir, "original.eml"));
            foreach (OutgoingShare s in shares) Console.WriteLine("    " + Path.Combine(sharesDir, "share-" + s.Chunk.Index + ".eml"));
            Console.WriteLine("    " + Path.Combine(rebuiltDir, "rebuilt.eml"));
            return allMatch ? 0 : 1;
        }

        /// <summary>Shows the arithmetic for the first number of the body: share 1 + ... + share N = secret (mod p).</summary>
        private static void ShowOneNumber(string body, IReadOnlyList<OutgoingShare> shares)
        {
            byte[] bytes = Utf8(body);
            if (bytes.Length == 0) return;

            ulong secret = AdditiveSharing.Pack(bytes)[0];
            var parts = shares.Select(s => SharePayload.Decode(s.Files.First(f => f.FileName == SharedMessage.BodyShareFileName).Content).Elements[0]).ToList();
            ulong sum = 0;
            foreach (ulong part in parts) sum = PrimeField.Add(sum, part);

            string text = Encoding.UTF8.GetString(bytes, 0, Math.Min(7, bytes.Length)).Replace("\r", "\\r").Replace("\n", "\\n");
            Console.WriteLine();
            Console.WriteLine("  How one number is shared: the body's first 7 bytes, \"" + text + "\", packed into one number");
            Console.WriteLine("    secret          = " + secret.ToString("N0", CultureInfo.InvariantCulture).PadLeft(27));
            for (int i = 0; i < parts.Count; i++)
            {
                string note = i < parts.Count - 1 ? "random" : (parts.Count == 1 ? "the secret itself" : "secret - (the others), mod p");
                Console.WriteLine(("    share " + (i + 1)).PadRight(20) + "= " + parts[i].ToString("N0", CultureInfo.InvariantCulture).PadLeft(27) + "   (" + note + ")");
            }
            Console.WriteLine("    sum of shares mod p = " + sum.ToString("N0", CultureInfo.InvariantCulture).PadLeft(23) + "   " + Tick(sum == secret) + " equals the secret");
        }

        private static KeyValuePair<string, string> Header(string name, string value)
        {
            return new KeyValuePair<string, string>(name, value ?? string.Empty);
        }

        private static void SaveFiles(string dir, IEnumerable<MailFile> files)
        {
            Directory.CreateDirectory(dir);
            int n = 0;
            foreach (MailFile f in files)
            {
                n++;
                string path = Path.Combine(dir, MailFile.SafeFileName(f.FileName, n));
                if (File.Exists(path)) path = Path.Combine(dir, n + "-" + MailFile.SafeFileName(f.FileName, n));
                File.WriteAllBytes(path, f.Content);
            }
        }

        private static string Describe(MailFile f)
        {
            string hash;
            using (var sha = SHA256.Create()) hash = Hex(sha.ComputeHash(f.Content), 0, 8).Replace(" ", string.Empty).ToLowerInvariant();
            return f.FileName.PadRight(18) + (f.Content.Length.ToString("N0", CultureInfo.InvariantCulture) + " bytes").PadLeft(14) + "   SHA-256 " + hash + "...";
        }

        private static void PrintIndented(string text)
        {
            foreach (string line in text.Replace("\r\n", "\n").TrimEnd('\n').Split('\n')) Console.WriteLine("    | " + line);
        }

        private static string FirstWords(string body)
        {
            string line = body.Replace("\r\n", "\n").Split('\n').FirstOrDefault(l => l.Trim().Length >= 4) ?? body;
            return string.Join(" ", line.Trim().Split(' ').Take(2));
        }

        private static void Title(string text)
        {
            Console.WriteLine();
            Console.WriteLine(text);
            Console.WriteLine(new string('-', text.Length));
        }

        private static string Tick(bool ok)
        {
            return ok ? "yes ✓" : "NO ✗";
        }

        private static byte[] Utf8(string s)
        {
            return new UTF8Encoding(false).GetBytes(s);
        }

        private static string Hex(byte[] data, int start, int count)
        {
            var sb = new StringBuilder();
            for (int i = start; i < Math.Min(data.Length, start + count); i++)
            {
                if (sb.Length > 0) sb.Append(' ');
                sb.Append(data[i].ToString("X2", CultureInfo.InvariantCulture));
            }
            return sb.ToString();
        }

        private static bool Contains(byte[] haystack, byte[] needle)
        {
            if (needle.Length == 0) return false;
            for (int i = 0; i + needle.Length <= haystack.Length; i++)
            {
                int k = 0;
                while (k < needle.Length && haystack[i + k] == needle[k]) k++;
                if (k == needle.Length) return true;
            }
            return false;
        }

        private sealed class Options
        {
            public const string Usage =
                "Usage: dotnet run --project tools/MpcDemo -- [options] [attachment files...]\n" +
                "\n" +
                "With no files, a sample email with budget.csv, notes.txt and smiley.png is used.\n" +
                "\n" +
                "Options:\n" +
                "  --servers N       number of MPC servers, 1 to 5 (default 2)\n" +
                "  --subject TEXT    subject line\n" +
                "  --body TEXT       body text\n" +
                "  --body-file PATH  read the body from a text file\n" +
                "  --out FOLDER      where to write the .eml files (default ./demo-output)\n" +
                "  --help            show this help";

            public int Servers = 2;
            public string Subject;
            public string Body;
            public string OutputFolder = "demo-output";
            public bool ShowHelp;
            public readonly List<string> Files = new List<string>();

            public static Options Parse(string[] args)
            {
                var o = new Options();
                for (int i = 0; i < args.Length; i++)
                {
                    string a = args[i];
                    Func<string> next = () =>
                    {
                        if (i + 1 >= args.Length) throw new ArgumentException(a + " needs a value.");
                        return args[++i];
                    };

                    switch (a)
                    {
                        case "--help":
                        case "-h":
                            o.ShowHelp = true;
                            break;
                        case "--servers":
                            int n;
                            if (!int.TryParse(next(), out n) || n < 1 || n > 5) throw new ArgumentException("--servers must be 1 to 5.");
                            o.Servers = n;
                            break;
                        case "--subject":
                            o.Subject = next();
                            break;
                        case "--body":
                            o.Body = next().Replace("\\n", "\r\n");
                            break;
                        case "--body-file":
                            o.Body = File.ReadAllText(next());
                            break;
                        case "--out":
                            o.OutputFolder = next();
                            break;
                        default:
                            if (a.StartsWith("--", StringComparison.Ordinal)) throw new ArgumentException("Unknown option " + a + ".");
                            if (!File.Exists(a)) throw new ArgumentException("File not found: " + a);
                            o.Files.Add(a);
                            break;
                    }
                }
                return o;
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace Mpc.Core.Tests
{
    /// <summary>Random bytes from a seeded System.Random, so a test can repeat the same "randomness".</summary>
    internal sealed class SeededRng : RandomNumberGenerator
    {
        private readonly Random _random;

        public SeededRng(int seed)
        {
            _random = new Random(seed);
        }

        public override void GetBytes(byte[] data)
        {
            _random.NextBytes(data);
        }
    }

    /// <summary>Hands out a fixed byte sequence first, then zeros.</summary>
    internal sealed class ScriptedRng : RandomNumberGenerator
    {
        private readonly byte[] _script;
        private int _used;

        public ScriptedRng(byte[] script)
        {
            _script = script;
        }

        public override void GetBytes(byte[] data)
        {
            for (int i = 0; i < data.Length; i++) data[i] = _used < _script.Length ? _script[_used++] : (byte)0;
        }
    }

    public static class PrimeFieldTests
    {
        [Test]
        public static void AddsAndSubtractsModP()
        {
            ulong p = PrimeField.P;
            Assert.Equal(2305843009213693951UL, p, "p = 2^61 - 1");
            Assert.Equal(0UL, PrimeField.Add(p - 1, 1), "wraps to 0");
            Assert.Equal(p - 2, PrimeField.Add(p - 1, p - 1), "largest sum");
            Assert.Equal(12UL, PrimeField.Add(5, 7));
            Assert.Equal(2UL, PrimeField.Subtract(5, 3));
            Assert.Equal(p - 1, PrimeField.Subtract(0, 1), "wraps below 0");
            Assert.True(PrimeField.IsElement(p - 1), "p - 1 is in the field");
            Assert.False(PrimeField.IsElement(p), "p is not");
            Assert.Throws<FormatException>(() => PrimeField.CheckElement(p), "p rejected");
        }
    }

    public static class AdditiveSharingTests
    {
        [Test]
        public static void PacksSevenBytesPerElement()
        {
            ulong[] packed = AdditiveSharing.Pack(new byte[] { 1, 2, 3, 4, 5, 6, 7, 8 });
            Assert.Sequence(new[] { 0x07060504030201UL, 0x08UL }, packed, "little-endian, 7 bytes each");

            for (int length = 0; length <= 30; length++)
            {
                byte[] data = Enumerable.Range(0, length).Select(i => (byte)(255 - i)).ToArray();
                ulong[] elements = AdditiveSharing.Pack(data);
                Assert.Equal((length + 6) / 7, elements.Length, "element count for " + length);
                Assert.True(elements.All(e => e < (1UL << 56)), "below 2^56");
                Assert.Sequence(data, AdditiveSharing.Unpack(elements, length), "round trip " + length);
            }
        }

        [Test]
        public static void UnpackRejectsDataThatPackCannotMake()
        {
            Assert.Throws<FormatException>(() => AdditiveSharing.Unpack(new[] { 1UL << 56 }, 7), "element of 8 bytes");
            Assert.Throws<FormatException>(() => AdditiveSharing.Unpack(new[] { 0x0100UL }, 1), "non-zero padding");
            Assert.Throws<FormatException>(() => AdditiveSharing.Unpack(new[] { 1UL, 2UL }, 7), "too many elements");
        }

        [Test]
        public static void SharesAddBackToTheSecret()
        {
            var rng = new SeededRng(1);
            ulong[] secret = { 0, 1, 42, PrimeField.P - 1, 123456789012345UL };
            for (int n = 1; n <= 5; n++)
            {
                ulong[][] shares = AdditiveSharing.Split(secret, n, rng);
                Assert.Equal(n, shares.Length, "share count");
                Assert.True(shares.All(s => s.Length == secret.Length && s.All(PrimeField.IsElement)), "shares are field elements");
                Assert.Sequence(secret, AdditiveSharing.Reconstruct(shares), "reconstructed with " + n);
            }

            Assert.Sequence(secret, AdditiveSharing.Split(secret, 1, rng)[0], "one share is the secret itself");
        }

        [Test]
        public static void FirstSharesDependOnlyOnTheRandomness()
        {
            // With the same randomness, share 1 of two different secrets is identical: it is
            // pure randomness, so on its own it says nothing about the secret.
            ulong[] hello = AdditiveSharing.Pack(Encoding.UTF8.GetBytes("hello, world!"));
            ulong[] other = AdditiveSharing.Pack(Encoding.UTF8.GetBytes("transfer $1000"));
            ulong[][] a = AdditiveSharing.Split(hello, 2, new SeededRng(7));
            ulong[][] b = AdditiveSharing.Split(other, 2, new SeededRng(7));
            Assert.Sequence(a[0], b[0], "share 1 identical");
            Assert.False(a[1].SequenceEqual(b[1]), "share 2 differs");
        }

        [Test]
        public static void DrawsAgainWhenRandomBitsHitP()
        {
            // The first 8 random bytes are all 0xFF: their low 61 bits are exactly p, which is not
            // a field element, so they must be skipped and the next 8 bytes used.
            byte[] script = Enumerable.Repeat((byte)0xFF, 8).Concat(new byte[] { 9, 0, 0, 0, 0, 0, 0, 0 }).ToArray();
            ulong[][] shares = AdditiveSharing.Split(new ulong[] { 100 }, 2, new ScriptedRng(script));
            Assert.Equal(9UL, shares[0][0], "the 0xFF block was skipped");
            Assert.Equal(91UL, shares[1][0], "100 - 9");
        }

        [Test]
        public static void ReconstructRejectsMismatchedShares()
        {
            Assert.Throws<FormatException>(() => AdditiveSharing.Reconstruct(new[] { new ulong[2], new ulong[3] }), "lengths differ");
            Assert.Throws<FormatException>(() => AdditiveSharing.Reconstruct(new[] { new[] { PrimeField.P } }), "not a field element");
        }
    }

    public static class SharePayloadTests
    {
        [Test]
        public static void EncodesTheDocumentedLayout()
        {
            var payload = new SharePayload(ShareKind.Attachment, 3, 9, new ulong[] { 0x0102030405060708UL, 5UL });
            byte[] b = payload.Encode();

            Assert.Equal(18 + 16, b.Length, "header + 2 elements");
            Assert.Equal("SMLS", Encoding.ASCII.GetString(b, 0, 4), "magic");
            Assert.Sequence(new byte[] { 1, 2, 1, 0 }, b.Skip(4).Take(4), "version, kind, ring, form");
            Assert.Sequence(new byte[] { 3, 0 }, b.Skip(8).Take(2), "attachment number");
            Assert.Sequence(new byte[] { 9, 0, 0, 0, 0, 0, 0, 0 }, b.Skip(10).Take(8), "data length");
            Assert.Sequence(new byte[] { 8, 7, 6, 5, 4, 3, 2, 1 }, b.Skip(18).Take(8), "element 1, little-endian");

            SharePayload back = SharePayload.Decode(b);
            Assert.Equal(ShareKind.Attachment, back.Kind);
            Assert.Equal(3, back.AttachmentNumber);
            Assert.Equal(9L, back.DataLength);
            Assert.Sequence(payload.Elements, back.Elements, "elements");
        }

        [Test]
        public static void RejectsFilesThatAreNotValidShares()
        {
            byte[] good = new SharePayload(ShareKind.Body, 0, 7, new ulong[] { 1 }).Encode();
            Func<int, byte, byte[]> with = (index, value) =>
            {
                byte[] copy = (byte[])good.Clone();
                copy[index] = value;
                return copy;
            };

            SharePayload.Decode(good);
            Assert.Throws<FormatException>(() => SharePayload.Decode(with(0, (byte)'X')), "bad magic");
            Assert.Throws<FormatException>(() => SharePayload.Decode(with(4, 2)), "unknown version");
            Assert.Throws<FormatException>(() => SharePayload.Decode(with(5, 3)), "unknown kind");
            Assert.Throws<FormatException>(() => SharePayload.Decode(with(6, 2)), "unknown field");
            Assert.Throws<FormatException>(() => SharePayload.Decode(with(7, 1)), "seed form not supported");
            Assert.Throws<FormatException>(() => SharePayload.Decode(with(8, 5)), "body with an attachment number");
            Assert.Throws<FormatException>(() => SharePayload.Decode(with(10, 8)), "length needs 2 elements, file has 1");
            Assert.Throws<FormatException>(() => SharePayload.Decode(good.Take(good.Length - 1).ToArray()), "cut short");
            Assert.Throws<FormatException>(() => SharePayload.Decode(with(25, 0x20)), "element not below p");
        }
    }

    public static class SharedMessageTests
    {
        private static SharedMessage Sample()
        {
            return new SharedMessage(
                "Meet me at 5pm near the main gate.\r\nBring the contract. ünïcödé 😀\r\n",
                new[]
                {
                    new MailFile("contract.pdf", Enumerable.Range(0, 1000).Select(i => (byte)(i * 7)).ToArray()),
                    new MailFile("empty.txt", new byte[0]),
                    new MailFile("résumé 日本.docx", new byte[] { 0, 255, 0, 255 }),
                    new MailFile("contract.pdf", new byte[] { 42 }), // same name twice
                });
        }

        [Test]
        public static void RoundTripsBodyAndAttachments()
        {
            SharedMessage original = Sample();
            for (int n = 1; n <= 4; n++)
            {
                IReadOnlyList<IReadOnlyList<MailFile>> shares = original.Split(n, new SeededRng(n));
                Assert.Equal(n, shares.Count, "one list per server");
                foreach (IReadOnlyList<MailFile> files in shares)
                {
                    Assert.Sequence(
                        new[] { "smilies-body.bin", "smilies-att-1.bin", "smilies-att-2.bin", "smilies-att-3.bin", "smilies-att-4.bin" },
                        files.Select(f => f.FileName), "share file names");
                }

                SharedMessage back = SharedMessage.Combine(shares);
                Assert.Equal(original.Body, back.Body, "body with " + n);
                Assert.Sequence(original.Attachments.Select(a => a.FileName), back.Attachments.Select(a => a.FileName), "names with " + n);
                for (int a = 0; a < original.Attachments.Count; a++)
                {
                    Assert.Sequence(original.Attachments[a].Content, back.Attachments[a].Content, "attachment " + (a + 1) + " with " + n);
                }
            }
        }

        [Test]
        public static void NoShareContainsTheText()
        {
            SharedMessage original = Sample();
            IReadOnlyList<IReadOnlyList<MailFile>> shares = original.Split(2, new SeededRng(3));
            byte[][] secrets =
            {
                Encoding.UTF8.GetBytes("Meet me"),
                Encoding.UTF8.GetBytes("contract"),
                Encoding.UTF8.GetBytes("résumé"),
            };

            foreach (MailFile file in shares.SelectMany(s => s))
            {
                foreach (byte[] secret in secrets)
                {
                    Assert.False(Contains(file.Content, secret), file.FileName + " contains \"" + Encoding.UTF8.GetString(secret) + "\"");
                }
            }
        }

        [Test]
        public static void RejectsSharesThatDontBelongTogether()
        {
            SharedMessage original = Sample();
            IReadOnlyList<IReadOnlyList<MailFile>> a = original.Split(2, new SeededRng(10));
            IReadOnlyList<IReadOnlyList<MailFile>> b = original.Split(2, new SeededRng(11));

            Assert.Throws<FormatException>(() => SharedMessage.Combine(new[] { a[0], b[1] }), "shares of two different sends");
            Assert.Throws<FormatException>(() => SharedMessage.Combine(new[] { a[0], (IReadOnlyList<MailFile>)a[1].Skip(1).ToList() }), "body share missing");
            Assert.Throws<FormatException>(() => SharedMessage.Combine(new[] { a[0], (IReadOnlyList<MailFile>)a[1].Take(3).ToList() }), "attachment share missing");
            Assert.Throws<FormatException>(() => SharedMessage.Combine(new[] { a[0], (IReadOnlyList<MailFile>)a[1].Concat(a[1].Take(1)).ToList() }), "duplicate share file");

            var other = new SharedMessage("different", new[] { new MailFile("x", new byte[] { 1 }) }).Split(2, new SeededRng(12));
            Assert.Throws<FormatException>(() => SharedMessage.Combine(new[] { a[0], other[1] }), "another message");
        }

        [Test]
        public static void IgnoresOtherAttachments()
        {
            SharedMessage original = Sample();
            IReadOnlyList<IReadOnlyList<MailFile>> shares = original.Split(2, new SeededRng(5));
            var withExtra = new[]
            {
                (IReadOnlyList<MailFile>)shares[0].Concat(new[] { new MailFile("footer-logo.png", new byte[] { 1, 2, 3 }) }).ToList(),
                shares[1],
            };
            Assert.Equal(original.Body, SharedMessage.Combine(withExtra).Body, "a server-added file is ignored");
        }

        [Test]
        public static void MakesFileNamesSafe()
        {
            Assert.Equal("report.pdf", MailFile.SafeFileName("report.pdf", 1));
            Assert.Equal(".._.._evil.exe", MailFile.SafeFileName("..\\../evil.exe", 1), "no path separators");
            Assert.Equal("attachment2", MailFile.SafeFileName("..", 2), "never a parent folder");
            Assert.Equal("attachment3", MailFile.SafeFileName("   ", 3), "blank");
            Assert.Equal("a_b_c_d", MailFile.SafeFileName("a:b*c?d", 4), "reserved characters");
            Assert.Equal("notes", MailFile.SafeFileName("notes. . ", 5), "no trailing dots or spaces");
            string longName = new string('x', 300) + ".pdf";
            string safe = MailFile.SafeFileName(longName, 6);
            Assert.Equal(200, safe.Length, "shortened");
            Assert.True(safe.EndsWith(".pdf", StringComparison.Ordinal), "keeps the extension");
        }

        private static bool Contains(byte[] haystack, byte[] needle)
        {
            for (int i = 0; i + needle.Length <= haystack.Length; i++)
            {
                int k = 0;
                while (k < needle.Length && haystack[i + k] == needle[k]) k++;
                if (k == needle.Length) return true;
            }
            return false;
        }
    }

    public static class ShareSendAndReceiveTests
    {
        private static readonly SrvRecord[] Servers =
        {
            new SrvRecord(0, 0, 25, "relay1.example.net."),
            new SrvRecord(0, 0, 25, "relay2.example.org"),
        };

        [Test]
        public static void PlansOneEmailSharePerServer()
        {
            string id = MpcSendPlanner.NewMpcId();
            IReadOnlyList<OutgoingShare> shares = MpcSendPlanner.PlanShares(
                "Meet me at 5pm", new[] { new MailFile("plan.txt", Encoding.UTF8.GetBytes("secret plan")) }, Servers, id, new SeededRng(1));

            Assert.Equal(2, shares.Count, "shares");
            Assert.Sequence(
                new[] { "X-MPC-Via", "X-MPC-ID", "X-MPC-Chunk", "X-MPC-Result", "X-MPC-Scheme" },
                shares[0].Headers.Select(h => h.Key), "header names");
            Assert.Equal("relay1.example.net", shares[0].Via, "trailing dot removed");
            Assert.Equal("relay2.example.org", shares[1].Headers[0].Value);
            Assert.Equal(id, shares[1].Headers[1].Value);
            Assert.Equal("[001/002]", shares[0].Headers[2].Value);
            Assert.Equal("[002/002]", shares[1].Headers[2].Value);
            Assert.Equal(" ", shares[0].Headers[3].Value);
            Assert.Equal("additive-p61", shares[0].Headers[4].Value);

            foreach (OutgoingShare share in shares)
            {
                Assert.Sequence(new[] { "smilies-body.bin", "smilies-att-1.bin" }, share.Files.Select(f => f.FileName), "files");
                Assert.False(share.Body.Contains("5pm") || share.Body.Contains("plan"), "placeholder body has none of the content");
            }
        }

        [Test]
        public static void ReadsTheSchemeHeader()
        {
            string raw = "X-MPC-Via: relay1\r\nX-MPC-ID: A\r\nX-MPC-Chunk: [001/002]\r\nX-MPC-Scheme: additive-p61\r\n\r\n";
            ReceivedChunk<string> share = MpcCoalescer.TryReadChunk("s1", raw);
            Assert.Equal("additive-p61", share.Scheme);

            ReceivedChunk<string> legacy = MpcCoalescer.TryReadChunk("c2", "X-MPC-Via: relay2\r\nX-MPC-ID: A\r\nX-MPC-Chunk: [002/002]\r\n\r\n");
            Assert.Null(legacy.Scheme, "mailutils chunk has no scheme");

            ChunkGroup<string> mixed = MpcCoalescer.Group(new[] { share, legacy }).Single();
            Assert.False(mixed.HasOneScheme, "mixed schemes detected");

            ChunkGroup<string> shares = MpcCoalescer.Group(new[] { share, MpcCoalescer.TryReadChunk("s2", raw.Replace("[001/002]", "[002/002]")) }).Single();
            Assert.True(shares.HasOneScheme && shares.IsAdditiveShares, "additive shares");
        }

        [Test]
        public static void RoundTripsThroughSendAndReceive()
        {
            string body = "Line one\r\nLine two with ünïcödé and 😀\r\n";
            var attachments = new[] { new MailFile("photo.jpg", Enumerable.Range(0, 5000).Select(i => (byte)i).ToArray()) };
            IReadOnlyList<OutgoingShare> sent = MpcSendPlanner.PlanShares(body, attachments, Servers, MpcSendPlanner.NewMpcId());

            // Deliver in reverse order, rendering each share's headers as a receiving server would.
            var received = sent.Reverse().Select(s =>
            {
                string raw = string.Concat(s.Headers.Select(h => h.Key + ": " + h.Value + "\r\n")) + "\r\n";
                return MpcCoalescer.TryReadChunk(s.Files, raw);
            }).ToList();

            ChunkGroup<IReadOnlyList<MailFile>> g = MpcCoalescer.Group(received).Single();
            Assert.True(g.IsComplete && g.IsAdditiveShares, "complete set of shares");

            SharedMessage back = SharedMessage.Combine(g.Chunks.Select(c => c.Item).ToList());
            Assert.Equal(body, back.Body, "body");
            Assert.Equal("photo.jpg", back.Attachments.Single().FileName, "attachment name");
            Assert.Sequence(attachments[0].Content, back.Attachments.Single().Content, "attachment content");
            Assert.Equal("relay1.example.net, relay2.example.org", g.MergedVia, "via");
        }
    }
}

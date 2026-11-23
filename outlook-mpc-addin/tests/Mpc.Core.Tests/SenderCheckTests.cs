using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Mpc.Core.Tests
{
    /// <summary>
    /// Stage 1 (sender checks). The Authentication-Results lines below are the ones Postfix with
    /// postfix-policyd-spf-python, OpenDKIM and OpenDMARC (Ubuntu 24.04, as in the lab) added to a
    /// chunk signed for green.mpc, configured as lab/aws-opentofu/templates/bootstrap.sh.tftpl does.
    /// </summary>
    public static class SenderCheckTests
    {
        private const string RelayA = "relay.provider-a.test";
        private const string RelayB = "relay.provider-b.test";

        private static string[] Passed(string relay)
        {
            return new[]
            {
                relay + "; dmarc=pass (p=reject dis=none) header.from=green.mpc",
                relay + ";\r\n\tdkim=pass (2048-bit key; unprotected) header.d=green.mpc header.i=@green.mpc header.a=rsa-sha256 header.s=mpc header.b=zhixlIZt;\r\n\tdkim-atps=neutral",
                relay + "; spf=pass (sender SPF authorized) smtp.mailfrom=green.mpc (client-ip=10.20.1.10; helo=mail.green.mpc; envelope-from=alice@green.mpc; receiver=blue.mpc)",
            };
        }

        /// <summary>A forged chunk: unsigned, sent from an address not in green.mpc's SPF record.</summary>
        private static string[] Forged(string relay)
        {
            return new[]
            {
                relay + "; dmarc=fail (p=reject dis=none) header.from=green.mpc",
                relay + "; dkim=none;\r\n\tdkim-atps=neutral",
                relay + "; spf=fail (SPF fail - not authorized) smtp.mailfrom=green.mpc (client-ip=10.20.9.9; helo=evil; envelope-from=alice@green.mpc; receiver=blue.mpc)",
            };
        }

        private static string Headers(int index, string via, string from, IEnumerable<string> authenticationResults)
        {
            var sb = new StringBuilder();
            sb.Append("Received: from relay.provider-c.test by mail.blue.mpc\r\n");
            foreach (string ar in authenticationResults) sb.Append("Authentication-Results: ").Append(ar).Append("\r\n");
            sb.Append("From: ").Append(from).Append("\r\n");
            sb.Append("To: Bob <bob@blue.mpc>\r\n");
            sb.Append("Subject: Project plan\r\n");
            sb.Append("X-MPC-Via: ").Append(via).Append("\r\n");
            sb.Append("X-MPC-ID: 3F2504E0-4F89-11D3-9A0C-0305E82C3301\r\n");
            sb.Append("X-MPC-Chunk: [00").Append(index).Append("/002]\r\n");
            sb.Append("X-MPC-Result:  \r\n");
            sb.Append("X-MPC-Scheme: additive-p61\r\n\r\n");
            return sb.ToString();
        }

        private static SenderCheckResult Check(string from1, IEnumerable<string> results1, string from2, IEnumerable<string> results2)
        {
            var chunks = new[]
            {
                MpcCoalescer.TryReadChunk("m1", Headers(1, RelayA, from1, results1)),
                MpcCoalescer.TryReadChunk("m2", Headers(2, RelayB, from2, results2)),
            };
            IReadOnlyList<ChunkGroup<string>> groups = MpcCoalescer.Group(chunks);
            Assert.Equal(1, groups.Count, "one message");
            Assert.True(groups[0].IsComplete, "complete");
            return SenderChecks.Evaluate(groups[0]);
        }

        private const string Alice = "Alice <alice@green.mpc>";

        [Test]
        public static void ParsesTheLabsHeaders()
        {
            AuthenticationResults dmarc, dkim, spf;
            Assert.True(AuthenticationResults.TryParse(Passed(RelayA)[0], out dmarc), "dmarc");
            Assert.True(AuthenticationResults.TryParse(Passed(RelayA)[1].Replace("\r\n", ""), out dkim), "dkim");
            Assert.True(AuthenticationResults.TryParse(Passed(RelayA)[2], out spf), "spf");

            Assert.Equal(RelayA, dmarc.AuthservId);
            Assert.Equal("dmarc", dmarc.Results[0].Method);
            Assert.Equal("pass", dmarc.Results[0].Result);
            Assert.Equal("green.mpc", dmarc.Results[0].GetProperty("header.from"));

            Assert.Sequence(new[] { "dkim=pass", "dkim-atps=neutral" }, dkim.Results.Select(r => r.ToString()), "dkim results");
            Assert.Equal("green.mpc", dkim.Results[0].GetProperty("header.d"));
            Assert.Equal("mpc", dkim.Results[0].GetProperty("header.s"));

            // The comment holds semicolons; they must not split the result.
            Assert.Equal(1, spf.Results.Count, "one SPF result");
            Assert.Equal("pass", spf.Results[0].Result);
            Assert.Equal("green.mpc", spf.Results[0].GetProperty("smtp.mailfrom"));
        }

        [Test]
        public static void ParsesOtherForms()
        {
            AuthenticationResults ar;
            // The SPF checker's line when it skips a relay: no method at all.
            Assert.True(AuthenticationResults.TryParse(RelayA + "; none (SPF check N/A for local connections - client-ip=10.20.5.10; helo=relay.provider-c.test)", out ar), "none");
            Assert.Equal(0, ar.Results.Count, "no results");

            // Version number, method version, quoted value with a semicolon, nested comment.
            Assert.True(AuthenticationResults.TryParse("mx.example.com 1; dkim/1=pass (good (very)) header.d=\"a;b.example\" reason=\"ok\"; dmarc=FAIL", out ar), "mixed");
            Assert.Equal("mx.example.com", ar.AuthservId);
            Assert.Sequence(new[] { "dkim=pass", "dmarc=fail" }, ar.Results.Select(r => r.ToString()), "results");
            Assert.Equal("a;b.example", ar.Results[0].GetProperty("header.d"));
            Assert.Equal("ok", ar.Results[0].GetProperty("reason"));

            Assert.False(AuthenticationResults.TryParse("   ", out ar), "empty");
        }

        [Test]
        public static void ReadsFromAndResultsWithTheChunk()
        {
            ReceivedChunk<string> chunk = MpcCoalescer.TryReadChunk("m", Headers(1, RelayA, Alice, Passed(RelayA)));
            Assert.Equal(Alice, chunk.From);
            Assert.Equal(3, chunk.AuthenticationResults.Count, "three headers");
            Assert.True(chunk.AuthenticationResults.All(a => a.AuthservId == RelayA), "all from relay-a");
        }

        [Test]
        public static void PassesWhenEveryChunkPassedAtItsRelay()
        {
            SenderCheckResult result = Check(Alice, Passed(RelayA), Alice, Passed(RelayB));
            Assert.True(result.Passed, "passed: " + string.Join("; ", result.Problems));
        }

        [Test]
        public static void FailsAForgedChunk()
        {
            SenderCheckResult result = Check(Alice, Passed(RelayA), Alice, Forged(RelayB));
            Assert.False(result.Passed, "forged chunk 2");
            Assert.True(result.Problems.Any(p => p.StartsWith("chunk 2: DMARC fail", StringComparison.Ordinal)), "DMARC: " + string.Join("; ", result.Problems));
            Assert.True(result.Problems.Any(p => p.StartsWith("chunk 2: no valid DKIM signature from green.mpc", StringComparison.Ordinal)), "DKIM");
            Assert.False(result.Problems.Any(p => p.StartsWith("chunk 1", StringComparison.Ordinal)), "chunk 1 is fine");
        }

        [Test]
        public static void TrustsOnlyTheRelayInXMpcVia()
        {
            // Chunk 2 went via relay-b, but carries only results claiming to be from relay-a
            // (for example written by the sender): they don't count.
            SenderCheckResult result = Check(Alice, Passed(RelayA), Alice, Passed(RelayA));
            Assert.False(result.Passed, "wrong relay");
            Assert.Equal("chunk 2: no sender-check results from " + RelayB, result.Problems.Single());

            // Results from relay-b in another letter case, or with a trailing dot, do count.
            Assert.True(Check(Alice, Passed(RelayA), Alice, Passed("Relay.Provider-B.test.")).Passed, "case and trailing dot");
        }

        [Test]
        public static void FailsWithoutResults()
        {
            SenderCheckResult result = Check(Alice, Passed(RelayA), Alice, new string[0]);
            Assert.Equal("chunk 2: no sender-check results from " + RelayB, result.Problems.Single());
        }

        [Test]
        public static void SpfAloneIsNotEnough()
        {
            // DMARC passes on SPF, but the chunk isn't signed: it may have been changed.
            var spfOnly = new[] { RelayB + "; dmarc=pass header.from=green.mpc", RelayB + "; dkim=none", RelayB + "; spf=pass smtp.mailfrom=green.mpc" };
            SenderCheckResult result = Check(Alice, Passed(RelayA), Alice, spfOnly);
            Assert.Equal("chunk 2: no valid DKIM signature from green.mpc (at " + RelayB + ")", result.Problems.Single());
        }

        [Test]
        public static void DkimMustBeFromTheFromDomain()
        {
            var otherSigner = new[] { RelayB + "; dmarc=pass header.from=green.mpc", RelayB + "; dkim=pass header.d=evil.example", RelayB + "; spf=pass smtp.mailfrom=green.mpc" };
            Assert.False(Check(Alice, Passed(RelayA), Alice, otherSigner).Passed, "signed by another domain");

            var dmarcForAnotherDomain = new[] { RelayB + "; dmarc=pass header.from=evil.example", RelayB + "; dkim=pass header.d=green.mpc" };
            Assert.False(Check(Alice, Passed(RelayA), Alice, dmarcForAnotherDomain).Passed, "DMARC for another From domain");
        }

        [Test]
        public static void ChunksMustHaveTheSameSender()
        {
            SenderCheckResult result = Check(Alice, Passed(RelayA), "Mallory <mallory@green.mpc>", Passed(RelayB));
            Assert.Equal("the chunks have different From: addresses (alice@green.mpc, mallory@green.mpc)", result.Problems.Single());

            // The display name and letter case don't matter.
            Assert.True(Check(Alice, Passed(RelayA), "ALICE@Green.MPC", Passed(RelayB)).Passed, "same address");
        }

        [Test]
        public static void FailsWithoutAFromAddress()
        {
            SenderCheckResult result = Check(Alice, Passed(RelayA), "undisclosed-recipients:;", Passed(RelayB));
            Assert.Equal("chunk 2: no From: address", result.Problems.Single());
        }

        [Test]
        public static void GetsTheAddress()
        {
            Assert.Equal("alice@green.mpc", EmailAddress.GetAddress("Alice <Alice@Green.MPC>"));
            Assert.Equal("alice@green.mpc", EmailAddress.GetAddress(" alice@green.mpc "));
            Assert.Null(EmailAddress.GetAddress("Alice"), "no @");
            Assert.Null(EmailAddress.GetAddress("@green.mpc"), "no local part");
            Assert.Null(EmailAddress.GetAddress(null), "null");
        }

        [Test]
        public static void RequireSenderChecksSetting()
        {
            Assert.True(new MpcSettings().RequireSenderChecks, "on by default");

            string path = Path.Combine(Path.GetTempPath(), "mpc-settings-" + Guid.NewGuid().ToString("N") + ".ini");
            try
            {
                File.WriteAllText(path, "RequireSenderChecks=false\n");
                Assert.False(MpcSettings.Load(path).RequireSenderChecks, "turned off");

                var settings = new MpcSettings();
                settings.Save(path);
                Assert.True(File.ReadAllText(path).Contains("RequireSenderChecks=true"), "saved");
                Assert.True(MpcSettings.Load(path).RequireSenderChecks, "read back");
            }
            finally
            {
                File.Delete(path);
            }
        }
    }
}

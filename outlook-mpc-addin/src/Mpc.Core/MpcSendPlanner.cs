using System;
using System.Collections.Generic;

namespace Mpc.Core
{
    /// <summary>One message to send: a slice of the body plus its X-MPC-* headers.</summary>
    public sealed class OutgoingChunk
    {
        public OutgoingChunk(ChunkValue chunk, string body, string via, string mpcId)
        {
            Chunk = chunk;
            Body = body;
            Via = via;
            MpcId = mpcId;

            // Same headers, values and order as mpc_mail_compose_send__() in send.c.
            Headers = new[]
            {
                new KeyValuePair<string, string>(MpcConstants.HeaderVia, via),
                new KeyValuePair<string, string>(MpcConstants.HeaderChunkId, mpcId),
                new KeyValuePair<string, string>(MpcConstants.HeaderChunkValue, chunk.ToString()),
                new KeyValuePair<string, string>(MpcConstants.HeaderResult, MpcConstants.BlankResultValue),
            };
        }

        public ChunkValue Chunk { get; private set; }
        public string Body { get; private set; }
        public string Via { get; private set; }
        public string MpcId { get; private set; }

        /// <summary>As in send.c, attachments travel only with the last chunk.</summary>
        public bool CarriesAttachments
        {
            get { return Chunk.Index == Chunk.Total; }
        }

        public IReadOnlyList<KeyValuePair<string, string>> Headers { get; private set; }
    }

    /// <summary>
    /// One email-share to send: the X-MPC-* headers, a short plain-text body explaining what
    /// the message is, and this server's share files (body and every attachment).
    /// </summary>
    public sealed class OutgoingShare
    {
        public OutgoingShare(ChunkValue chunk, string via, string mpcId, IReadOnlyList<MailFile> files)
        {
            Chunk = chunk;
            Via = via;
            MpcId = mpcId;
            Files = files;

            // The four headers of send.c, in its order, then the scheme.
            Headers = new[]
            {
                new KeyValuePair<string, string>(MpcConstants.HeaderVia, via),
                new KeyValuePair<string, string>(MpcConstants.HeaderChunkId, mpcId),
                new KeyValuePair<string, string>(MpcConstants.HeaderChunkValue, chunk.ToString()),
                new KeyValuePair<string, string>(MpcConstants.HeaderResult, MpcConstants.BlankResultValue),
                new KeyValuePair<string, string>(MpcConstants.HeaderScheme, MpcConstants.SchemeAdditiveP61),
            };

            Body = string.Format(
                System.Globalization.CultureInfo.InvariantCulture,
                "This is share {0} of {1} of a message sent with MPC secret sharing.\r\n" +
                "The text and attachments are in the attached .bin files, which can only be read\r\n" +
                "once all {1} shares are combined by the MPC add-in.\r\n",
                chunk.Index, chunk.Total);
        }

        /// <summary>Share number and total, sent as X-MPC-Chunk.</summary>
        public ChunkValue Chunk { get; private set; }
        public string Via { get; private set; }
        public string MpcId { get; private set; }

        /// <summary>Visible body of the email-share. It contains none of the real text.</summary>
        public string Body { get; private set; }

        /// <summary>The share files to attach.</summary>
        public IReadOnlyList<MailFile> Files { get; private set; }

        public IReadOnlyList<KeyValuePair<string, string>> Headers { get; private set; }
    }

    public static class MpcSendPlanner
    {
        /// <summary>
        /// Secret-shares the body and attachments: one email-share per server, share n going via
        /// the n-th SRV target. Pass null for <paramref name="rng"/> to use the system's
        /// cryptographic random number generator.
        /// </summary>
        public static IReadOnlyList<OutgoingShare> PlanShares(
            string body, IReadOnlyList<MailFile> attachments, IReadOnlyList<SrvRecord> servers, string mpcId,
            System.Security.Cryptography.RandomNumberGenerator rng = null)
        {
            if (servers == null || servers.Count == 0) throw new ArgumentException("At least one MPC server is required.", "servers");
            if (string.IsNullOrWhiteSpace(mpcId)) throw new ArgumentException("An MPC ID is required.", "mpcId");

            var message = new SharedMessage(body, attachments);
            IReadOnlyList<IReadOnlyList<MailFile>> files;
            if (rng != null)
            {
                files = message.Split(servers.Count, rng);
            }
            else
            {
                using (var systemRng = System.Security.Cryptography.RandomNumberGenerator.Create())
                {
                    files = message.Split(servers.Count, systemRng);
                }
            }

            var shares = new List<OutgoingShare>(servers.Count);
            for (int i = 0; i < servers.Count; i++)
            {
                shares.Add(new OutgoingShare(new ChunkValue(i + 1, servers.Count), servers[i].Target.TrimEnd('.'), mpcId, files[i]));
            }
            return shares;
        }

        /// <summary>
        /// A new X-MPC-ID. send.c uses uuid_generate_random + uuid_unparse_upper, which gives
        /// the same upper-case 8-4-4-4-12 text as this.
        /// </summary>
        public static string NewMpcId()
        {
            return Guid.NewGuid().ToString("D").ToUpperInvariant();
        }

        /// <summary>
        /// The mailutils format: splits <paramref name="body"/> into one plain-text slice per
        /// server (chunk n goes via the n-th SRV target). The add-in now sends with
        /// <see cref="PlanShares"/>; this is kept to document and test compatibility with send.c.
        /// </summary>
        public static IReadOnlyList<OutgoingChunk> Plan(string body, IReadOnlyList<SrvRecord> servers, string mpcId)
        {
            if (servers == null || servers.Count == 0) throw new ArgumentException("At least one MPC server is required.", "servers");
            if (string.IsNullOrWhiteSpace(mpcId)) throw new ArgumentException("An MPC ID is required.", "mpcId");

            IReadOnlyList<string> parts = BodyChunker.Split(body, servers.Count);
            var chunks = new List<OutgoingChunk>(parts.Count);
            for (int i = 0; i < parts.Count; i++)
            {
                chunks.Add(new OutgoingChunk(
                    new ChunkValue(i + 1, parts.Count),
                    parts[i],
                    servers[i].Target.TrimEnd('.'),
                    mpcId));
            }
            return chunks;
        }
    }
}

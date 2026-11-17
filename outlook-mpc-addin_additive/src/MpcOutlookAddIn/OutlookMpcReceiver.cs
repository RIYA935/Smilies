using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Mpc.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MpcOutlookAddIn
{
    /// <summary>What one rejoin pass did.</summary>
    internal sealed class RejoinReport
    {
        public int Scanned;
        public int ChunksFound;
        public int Rejoined;
        public int Waiting;
        public int Failed;
        public readonly List<string> WaitingDetails = new List<string>();

        public override string ToString()
        {
            if (ChunksFound == 0) return "No MPC chunks were found in this folder.";

            var lines = new List<string>();
            lines.Add("Rejoined " + Rejoined + " message(s) from " + ChunksFound + " chunk(s).");
            if (Waiting > 0)
            {
                lines.Add(Waiting + " message(s) are still waiting for chunks:");
                lines.AddRange(WaitingDetails.Take(10).Select(d => "  • " + d));
            }
            if (Failed > 0) lines.Add(Failed + " message(s) could not be rejoined; see " + AddInLog.LogPath + ".");
            return string.Join(Environment.NewLine, lines);
        }
    }

    /// <summary>
    /// The receiving side of MPC, based on maildir_mpc_coaelesce_mails_unlocked() in mail.c.
    /// <para>
    /// Pieces are found by reading the X-MPC-* fields from each message's Internet headers and
    /// grouped by X-MPC-ID. When every piece of a message has arrived, a copy of piece 1 gets the
    /// rebuilt content, X-MPC-Chunk "[000/000]" and the combined X-MPC-Via; the pieces are then
    /// moved to Deleted Items. The content is rebuilt in one of two ways:
    /// </para>
    /// <list type="bullet">
    /// <item>X-MPC-Scheme "additive-p61" (sent by this add-in): the .bin share files of every
    /// piece are added up mod p to rebuild the body and the attachments.</item>
    /// <item>No X-MPC-Scheme (sent by mailutils): the body slices are joined in order and the
    /// attachments copied from the other pieces.</item>
    /// </list>
    /// </summary>
    internal sealed class OutlookMpcReceiver
    {
        private readonly Func<MpcSettings> _settings;

        public OutlookMpcReceiver(Func<MpcSettings> settings)
        {
            _settings = settings;
        }

        public RejoinReport RejoinFolder(Outlook.MAPIFolder folder)
        {
            var report = new RejoinReport();
            var chunks = new List<ReceivedChunk<Outlook.MailItem>>();

            Outlook.Items items = folder.Items;
            Outlook.Items scope = Restrict(items);
            try
            {
                foreach (object entry in scope)
                {
                    report.Scanned++;
                    var mail = entry as Outlook.MailItem;
                    ReceivedChunk<Outlook.MailItem> chunk = mail == null ? null : ReadChunk(mail);
                    if (chunk == null)
                    {
                        OutlookMpc.Release(entry);
                        continue;
                    }
                    chunks.Add(chunk);
                }

                report.ChunksFound = chunks.Count;

                foreach (ChunkGroup<Outlook.MailItem> group in MpcCoalescer.Group(chunks))
                {
                    if (!group.IsComplete)
                    {
                        report.Waiting++;
                        report.WaitingDetails.Add(
                            "\"" + group.Chunks[0].Item.Subject + "\": missing chunk " +
                            string.Join(", ", group.MissingChunkNumbers) + " of " + group.Total);
                        continue;
                    }

                    try
                    {
                        Rejoin(group);
                        report.Rejoined++;
                    }
                    catch (Exception ex)
                    {
                        report.Failed++;
                        AddInLog.Write("Rejoining X-MPC-ID " + group.MpcId + " failed", ex);
                    }
                }
            }
            finally
            {
                foreach (ReceivedChunk<Outlook.MailItem> c in chunks) OutlookMpc.Release(c.Item);
                if (!ReferenceEquals(scope, items)) OutlookMpc.Release(scope);
                OutlookMpc.Release(items);
            }

            AddInLog.Write("Rejoin pass on \"" + folder.Name + "\": scanned " + report.Scanned + ", chunks " + report.ChunksFound +
                           ", rejoined " + report.Rejoined + ", waiting " + report.Waiting + ", failed " + report.Failed + ".");
            return report;
        }

        private static ReceivedChunk<Outlook.MailItem> ReadChunk(Outlook.MailItem mail)
        {
            if (OutlookMpc.IsRejoined(mail)) return null;

            string headers = OutlookMpc.GetTransportHeaders(mail);
            if (string.IsNullOrEmpty(headers) ||
                headers.IndexOf(MpcConstants.HeaderChunkId, StringComparison.OrdinalIgnoreCase) < 0)
            {
                return null;
            }

            return MpcCoalescer.TryReadChunk(mail, headers);
        }

        private void Rejoin(ChunkGroup<Outlook.MailItem> group)
        {
            IReadOnlyList<ReceivedChunk<Outlook.MailItem>> parts = group.Chunks;
            Outlook.MailItem first = parts[0].Item;

            if (!group.HasOneScheme)
            {
                throw new InvalidOperationException("Its pieces have different X-MPC-Scheme values.");
            }

            // For shares, rebuild everything before touching the mailbox: if the shares don't
            // add up, nothing is created and the pieces stay where they are.
            SharedMessage rebuilt = null;
            if (group.IsAdditiveShares)
            {
                var filesPerShare = parts
                    .Select(p => (IReadOnlyList<MailFile>)OutlookMpc.ReadAttachments(p.Item, SharedMessage.IsShareFileName))
                    .ToList();
                rebuilt = SharedMessage.Combine(filesPerShare);
            }
            else if (group.Scheme != null)
            {
                throw new NotSupportedException("Unknown X-MPC-Scheme \"" + group.Scheme + "\".");
            }

            // Copy() keeps the sender, recipients, dates and folder of piece 1.
            var merged = (Outlook.MailItem)first.Copy();
            try
            {
                merged.BodyFormat = Outlook.OlBodyFormat.olFormatPlain;

                if (rebuilt != null)
                {
                    // The copy carries piece 1's share files; replace them with the real files.
                    OutlookMpc.RemoveAttachmentsFrom(merged, 1);
                    merged.Body = rebuilt.Body;
                    OutlookMpc.AddAttachments(merged, rebuilt.Attachments);
                }
                else
                {
                    merged.Body = MpcCoalescer.MergeBodies(parts.Select(p => p.Item.Body ?? string.Empty));

                    // send.c attaches files to the last chunk; the copy of chunk 1 needs them.
                    foreach (ReceivedChunk<Outlook.MailItem> part in parts.Skip(1))
                    {
                        OutlookMpc.AddAttachments(merged, OutlookMpc.ReadAttachments(part.Item, null));
                    }
                }

                merged.UnRead = parts.Any(p => p.Item.UnRead);

                // As mail.c does: mark the message as rejoined and record every relay.
                string headers = OutlookMpc.GetTransportHeaders(merged) ?? OutlookMpc.GetTransportHeaders(first) ?? string.Empty;
                headers = InternetHeaders.SetValue(headers, MpcConstants.HeaderChunkValue, ChunkValue.Coalesced.ToString());
                headers = InternetHeaders.SetValue(headers, MpcConstants.HeaderVia, group.MergedVia);
                OutlookMpc.TrySetTransportHeaders(merged, headers);
                OutlookMpc.SetInternetHeader(merged, MpcConstants.HeaderChunkValue, ChunkValue.Coalesced.ToString());
                OutlookMpc.SetInternetHeader(merged, MpcConstants.HeaderVia, group.MergedVia);
                OutlookMpc.MarkRejoined(merged);

                merged.Save();
            }
            catch (Exception)
            {
                try
                {
                    merged.Delete();
                }
                catch (Exception ex)
                {
                    AddInLog.Write("Could not remove a partly rejoined message", ex);
                }
                throw;
            }
            finally
            {
                OutlookMpc.Release(merged);
            }

            AddInLog.Write("Rejoined X-MPC-ID " + group.MpcId + " (" + parts.Count + (rebuilt != null ? " shares" : " chunks") +
                           " via " + group.MergedVia + ").");

            
            bool delete = _settings().DeleteChunksAfterRejoin;
            foreach (ReceivedChunk<Outlook.MailItem> part in parts.Concat(group.Extras))
            {
                try
                {
                    // Mark every piece so a later pass never rebuilds this message again,
                    // even when the pieces are kept (DeleteChunksAfterRejoin=false).
                    OutlookMpc.MarkRejoined(part.Item);
                    part.Item.Save();

                    if (delete) part.Item.Delete(); // moves to Deleted Items
                }
                catch (Exception ex)
                {
                    AddInLog.Write("Could not mark or delete an MPC chunk after rejoining", ex);
                }
            }
            
        }

        /// <summary>Limits the scan to recently received messages (RejoinLookbackDays).</summary>
        private Outlook.Items Restrict(Outlook.Items items)
        {
            int days = _settings().RejoinLookbackDays;
            if (days <= 0) return items;

            // Outlook's [ReceivedTime] filter expects a date in the user's locale format.
            string since = DateTime.Now.AddDays(-days).ToString("g", CultureInfo.CurrentCulture);
            try
            {
                return items.Restrict("[ReceivedTime] >= '" + since + "'");
            }
            catch (Exception ex)
            {
                AddInLog.Write("Could not filter by received date (" + since + "); scanning the whole folder", ex);
                return items;
            }
        }
    }
}

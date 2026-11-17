using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Mpc.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MpcOutlookAddIn
{
    /// <summary>
    /// The sending side of MPC, based on mpc_mail_compose_send__() in send.c.
    /// <para>
    /// For a message whose first To recipient's domain publishes _mpc1._tcp / _mpc2._tcp SRV
    /// records, the plain-text body and every attachment are split into additive secret shares
    /// over Z_p (p = 2^61 - 1), one share per server. Each email-share carries a share of the
    /// body and of every attachment as .bin files, with a placeholder body. Shares 1..N-1 are
    /// sent as new messages, and the message being sent becomes share N. Every share carries
    /// X-MPC-Via, X-MPC-ID, X-MPC-Chunk, X-MPC-Result and X-MPC-Scheme.
    /// </para>
    /// </summary>
    internal sealed class OutlookMpcSender
    {
        private const string Caption = "MPC";

        private readonly Outlook.Application _application;
        private readonly Func<MpcSettings> _settings;
        private readonly Func<ISrvResolver> _resolverFactory;

        // True while this class sends shares itself, so its own ItemSend calls pass through.
        private bool _sendingChunks;

        public OutlookMpcSender(Outlook.Application application, Func<MpcSettings> settings, Func<ISrvResolver> resolverFactory)
        {
            _application = application;
            _settings = settings;
            _resolverFactory = resolverFactory;
        }

        private enum Outcome
        {
            /// <summary>No MPC servers for this recipient: send the message unchanged.</summary>
            SendNormally,

            /// <summary>Shares 1..N-1 are sent and the message is now share N: let it go.</summary>
            SendAsLastChunk,

            /// <summary>Do not send.</summary>
            Cancel,
        }

        /// <summary>Application.ItemSend handler: splits automatically when the setting is on.</summary>
        public void OnItemSend(object item, ref bool cancel)
        {
            if (_sendingChunks || cancel) return;

            var mail = item as Outlook.MailItem;
            if (mail == null) return;

            try
            {
                if (!_settings().AutoSplitOnSend) return;
                if (OutlookMpc.GetInternetHeader(mail, MpcConstants.HeaderChunkId) != null) return; // already a share

                if (Prepare(mail, false) == Outcome.Cancel) cancel = true;
            }
            catch (Exception ex)
            {
                AddInLog.Write("MPC processing failed in ItemSend", ex);
                DialogResult answer = MessageBox.Show(
                    "MPC could not process this message:\n\n" + ex.Message + "\n\nSend it as a normal message?",
                    Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (answer != DialogResult.Yes) cancel = true;
            }
        }

        /// <summary>The "Send via MPC" ribbon button: always splits, or explains why it can't.</summary>
        public void SendFromRibbon(Outlook.MailItem mail)
        {
            try
            {
                if (OutlookMpc.GetInternetHeader(mail, MpcConstants.HeaderChunkId) != null)
                {
                    MessageBox.Show("This message is already an MPC share.", Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (Prepare(mail, true) != Outcome.SendAsLastChunk) return;

                _sendingChunks = true;
                try
                {
                    ((Outlook._MailItem)mail).Send();
                }
                finally
                {
                    _sendingChunks = false;
                }
            }
            catch (Exception ex)
            {
                AddInLog.Write("Send via MPC failed", ex);
                MessageBox.Show("MPC could not send this message:\n\n" + ex.Message, Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Outcome Prepare(Outlook.MailItem mail, bool forced)
        {
            // send.c uses the domain of the To header (its first address).
            string to = FirstToAddress(mail);
            string domain = EmailAddress.GetDomain(to);
            if (domain == null)
            {
                if (forced)
                {
                    MessageBox.Show("Add a To recipient with an email address first.", Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return Outcome.Cancel;
                }
                return Outcome.SendNormally;
            }

            IReadOnlyList<SrvRecord> servers;
            try
            {
                servers = MpcServerDiscovery.DiscoverServers(_resolverFactory(), domain);
            }
            catch (DnsLookupException ex)
            {
                AddInLog.Write("SRV lookup for " + domain + " failed", ex);
                if (forced)
                {
                    MessageBox.Show("Could not look up the MPC servers for " + domain + ":\n\n" + ex.Message,
                        Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return Outcome.Cancel;
                }

                DialogResult answer = MessageBox.Show(
                    "Could not look up the MPC servers for " + domain + ":\n\n" + ex.Message +
                    "\n\nSend this as a normal (unshared) message?",
                    Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                return answer == DialogResult.Yes ? Outcome.SendNormally : Outcome.Cancel;
            }

            if (servers.Count == 0)
            {
                AddInLog.Write("No _mpcN._tcp SRV records for " + domain + "; sending normally.");
                if (forced)
                {
                    MessageBox.Show(
                        domain + " does not publish MPC servers (no SRV record for " + MpcConstants.SrvName(1, domain) +
                        "), so the message can't be sent via MPC.",
                        Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return Outcome.Cancel;
                }
                return Outcome.SendNormally;
            }

            if (!ConfirmOtherDomains(mail, domain)) return Outcome.Cancel;

            // 0. Read the text and attachments, and secret-share them. Nothing is changed yet.
            string mpcId = MpcSendPlanner.NewMpcId();
            List<MailFile> attachments = OutlookMpc.ReadAttachments(mail, null);
            IReadOnlyList<OutgoingShare> shares = MpcSendPlanner.PlanShares(mail.Body ?? string.Empty, attachments, servers, mpcId);
            if (shares.Count == 1)
            {
                AddInLog.Write("Only one MPC server for " + domain + ": its single share equals the message, so that server can read it.");
            }
            AddInLog.Write("Secret-sharing message for " + domain + " into " + shares.Count + " share(s) with " +
                           attachments.Count + " attachment(s), X-MPC-ID " + mpcId +
                           ", via " + string.Join(", ", shares.Select(c => c.Via)));

            // 1. Build shares 1..N-1 as new messages. Nothing has been sent yet, so on failure
            //    they are discarded and the original message is left untouched.
            var created = new List<Outlook.MailItem>();
            try
            {
                foreach (OutgoingShare share in shares.Take(shares.Count - 1))
                {
                    created.Add(CreateShareMessage(mail, share));
                }
            }
            catch (Exception)
            {
                foreach (Outlook.MailItem item in created) Discard(item);
                throw;
            }

            // 2. Send them.
            int sent = 0;
            _sendingChunks = true;
            try
            {
                foreach (Outlook.MailItem item in created)
                {
                    ((Outlook._MailItem)item).Send();
                    sent++;
                }
            }
            catch (Exception ex)
            {
                AddInLog.Write("Sending MPC share " + (sent + 1) + " failed", ex);
                foreach (Outlook.MailItem item in created.Skip(sent)) Discard(item);
                MessageBox.Show(
                    "Sending MPC share " + (sent + 1) + " of " + shares.Count + " failed:\n\n" + ex.Message +
                    (sent > 0 ? "\n\n" + sent + " share(s) were already sent. " : "\n\n") +
                    "Your message has not been sent.",
                    Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return Outcome.Cancel;
            }
            finally
            {
                _sendingChunks = false;
                foreach (Outlook.MailItem item in created) OutlookMpc.Release(item);
            }

            // 3. Turn the message being sent into the last share. The other shares are already
            //    out, so on failure stop here rather than offering to send the full text.
            try
            {
                ApplyShare(mail, shares[shares.Count - 1]);
            }
            catch (Exception ex)
            {
                AddInLog.Write("Preparing the last MPC share failed", ex);
                MessageBox.Show(
                    "Shares 1 to " + (shares.Count - 1) + " of " + shares.Count + " were sent, but the last share " +
                    "could not be prepared:\n\n" + ex.Message + "\n\nYour message has not been sent.",
                    Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return Outcome.Cancel;
            }
            return Outcome.SendAsLastChunk;
        }

        private Outlook.MailItem CreateShareMessage(Outlook.MailItem original, OutgoingShare share)
        {
            var item = (Outlook.MailItem)_application.CreateItem(Outlook.OlItemType.olMailItem);

            Outlook.Account account = original.SendUsingAccount;
            if (account != null)
            {
                item.SendUsingAccount = account;
                OutlookMpc.Release(account);
            }

            if (!string.IsNullOrEmpty(original.SentOnBehalfOfName)) item.SentOnBehalfOfName = original.SentOnBehalfOfName;
            item.Subject = original.Subject;
            item.Importance = original.Importance;
            item.Sensitivity = original.Sensitivity;

            Outlook.Recipients sourceRecipients = original.Recipients;
            Outlook.Recipients targetRecipients = item.Recipients;
            try
            {
                foreach (Outlook.Recipient source in sourceRecipients)
                {
                    Outlook.Recipient copy = targetRecipients.Add(OutlookMpc.GetSmtpAddress(source));
                    copy.Type = source.Type;
                    OutlookMpc.Release(copy);
                    OutlookMpc.Release(source);
                }

                if (!targetRecipients.ResolveAll())
                {
                    throw new InvalidOperationException("Outlook could not resolve all recipients for MPC share " + share.Chunk + ".");
                }
            }
            finally
            {
                OutlookMpc.Release(sourceRecipients);
                OutlookMpc.Release(targetRecipients);
            }

            try
            {
                ApplyShare(item, share);
            }
            catch (Exception)
            {
                Discard(item);
                throw;
            }
            return item;
        }

        /// <summary>
        /// Makes <paramref name="mail"/> an email-share: its attachments are replaced by this
        /// server's share files, its body by a placeholder, and the X-MPC-* headers are set.
        /// </summary>
        private static void ApplyShare(Outlook.MailItem mail, OutgoingShare share)
        {
            // Add the share files first. If that fails, remove what was added so the message
            // (and its original attachments) are as the user left them.
            int originalCount = OutlookMpc.AttachmentCount(mail);
            try
            {
                OutlookMpc.AddAttachments(mail, share.Files);
            }
            catch (Exception)
            {
                try
                {
                    OutlookMpc.RemoveAttachmentsFrom(mail, originalCount + 1);
                }
                catch (Exception ex)
                {
                    AddInLog.Write("Could not remove share files after a failure", ex);
                }
                throw;
            }

            // The originals are now in the shares; none of their content may go out in the clear.
            for (int i = originalCount; i >= 1; i--)
            {
                RemoveAttachment(mail, i);
            }

            // The body is shared as plain text, and the visible body is only a placeholder.
            mail.BodyFormat = Outlook.OlBodyFormat.olFormatPlain;
            mail.Body = share.Body;

            foreach (KeyValuePair<string, string> header in share.Headers)
            {
                OutlookMpc.SetInternetHeader(mail, header.Key, header.Value);
            }
        }

        private static void RemoveAttachment(Outlook.MailItem mail, int index)
        {
            Outlook.Attachments attachments = mail.Attachments;
            try
            {
                attachments.Remove(index);
            }
            finally
            {
                OutlookMpc.Release(attachments);
            }
        }

        private static string FirstToAddress(Outlook.MailItem mail)
        {
            Outlook.Recipients recipients = mail.Recipients;
            try
            {
                foreach (Outlook.Recipient r in recipients)
                {
                    try
                    {
                        if (r.Type == (int)Outlook.OlMailRecipientType.olTo) return OutlookMpc.GetSmtpAddress(r);
                    }
                    finally
                    {
                        OutlookMpc.Release(r);
                    }
                }
                return null;
            }
            finally
            {
                OutlookMpc.Release(recipients);
            }
        }

        /// <summary>
        /// Every share goes to every recipient, routed by the first To recipient's MPC servers
        /// (as in send.c). Ask before doing that for recipients in other domains.
        /// </summary>
        private static bool ConfirmOtherDomains(Outlook.MailItem mail, string mpcDomain)
        {
            var others = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
            Outlook.Recipients recipients = mail.Recipients;
            try
            {
                foreach (Outlook.Recipient r in recipients)
                {
                    string d = EmailAddress.GetDomain(OutlookMpc.GetSmtpAddress(r));
                    if (d != null && !string.Equals(d, mpcDomain, StringComparison.OrdinalIgnoreCase)) others.Add(d);
                    OutlookMpc.Release(r);
                }
            }
            finally
            {
                OutlookMpc.Release(recipients);
            }

            if (others.Count == 0) return true;

            return MessageBox.Show(
                "This message will be secret-shared using the MPC servers of " + mpcDomain + ".\n\n" +
                "Recipients at " + string.Join(", ", others) + " will receive the separate shares too, " +
                "and can only read the message if their mail program combines MPC shares.\n\nContinue?",
                Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
        }

        private static void Discard(Outlook.MailItem item)
        {
            try
            {
                item.Delete();
            }
            catch (Exception ex)
            {
                AddInLog.Write("Could not discard an unsent MPC share", ex);
            }
        }
    }
}

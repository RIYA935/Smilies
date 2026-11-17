using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Mpc.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MpcOutlookAddIn
{
    /// <summary>Reading and writing X-MPC-* headers and other MAPI properties on Outlook items.</summary>
    internal static class OutlookMpc
    {
        /// <summary>
        /// PS_INTERNET_HEADERS. A string property in this namespace becomes an Internet header
        /// of the same name when Outlook sends the message.
        /// </summary>
        public const string InternetHeaderNamespace =
            "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/";

        /// <summary>PR_TRANSPORT_MESSAGE_HEADERS: the raw header block of a received message.</summary>
        public const string TransportHeadersProperty = "http://schemas.microsoft.com/mapi/proptag/0x007D001F";

        /// <summary>PR_SMTP_ADDRESS on an address entry.</summary>
        public const string SmtpAddressProperty = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F";

        /// <summary>Private flag this add-in sets on a message it has rejoined.</summary>
        public const string RejoinedMarkerProperty =
            "http://schemas.microsoft.com/mapi/string/{6F1C3A52-7D4B-4E0B-9C1E-5A2B8D3F4E61}/MpcRejoined";

        public static void SetInternetHeader(Outlook.MailItem mail, string name, string value)
        {
            // NEW: a line break in a header value could add extra headers (header injection)
            if (value != null && (value.IndexOf('\r') >= 0 || value.IndexOf('\n') >= 0))
                throw new ArgumentException("Header value must not contain line breaks.", "value");
            
            Outlook.PropertyAccessor accessor = mail.PropertyAccessor;
            try
            {
                accessor.SetProperty(InternetHeaderNamespace + name, value ?? string.Empty);
            }
            finally
            {
                Release(accessor);
            }
        }

        /// <summary>Value of an X- header property on an item, or null when it is not set.</summary>
        public static string GetInternetHeader(Outlook.MailItem mail, string name)
        {
            // Exchange stores received headers under lower-case names; Outlook keeps the
            // original case on items it sends.
            string value = GetStringProperty(mail, InternetHeaderNamespace + name);
            if (value == null) value = GetStringProperty(mail, InternetHeaderNamespace + name.ToLowerInvariant());
            return value;
        }

        public static string GetTransportHeaders(Outlook.MailItem mail)
        {
            return GetStringProperty(mail, TransportHeadersProperty);
        }

        public static bool TrySetTransportHeaders(Outlook.MailItem mail, string headers)
        {
            Outlook.PropertyAccessor accessor = mail.PropertyAccessor;
            try
            {
                accessor.SetProperty(TransportHeadersProperty, headers);
                return true;
            }
            catch (Exception ex)
            {
                AddInLog.Write("Could not update the transport headers of the rejoined message: " + ex.Message);
                return false;
            }
            finally
            {
                Release(accessor);
            }
        }

        public static bool IsRejoined(Outlook.MailItem mail)
        {
            Outlook.PropertyAccessor accessor = mail.PropertyAccessor;
            try
            {
                object value = accessor.GetProperty(RejoinedMarkerProperty);
                return value is bool && (bool)value;
            }
            catch (Exception)
            {
                return false; // property not set
            }
            finally
            {
                Release(accessor);
            }
        }

        public static void MarkRejoined(Outlook.MailItem mail)
        {
            Outlook.PropertyAccessor accessor = mail.PropertyAccessor;
            try
            {
                accessor.SetProperty(RejoinedMarkerProperty, true);
            }
            finally
            {
                Release(accessor);
            }
        }

        /// <summary>The SMTP address of a recipient, including Exchange users and groups.</summary>
        public static string GetSmtpAddress(Outlook.Recipient recipient)
        {
            Outlook.AddressEntry entry = null;
            try
            {
                entry = recipient.AddressEntry;
                if (entry != null && string.Equals(entry.Type, "EX", StringComparison.OrdinalIgnoreCase))
                {
                    Outlook.ExchangeUser user = entry.GetExchangeUser();
                    if (user != null)
                    {
                        string smtp = user.PrimarySmtpAddress;
                        Release(user);
                        if (!string.IsNullOrEmpty(smtp)) return smtp;
                    }

                    Outlook.ExchangeDistributionList list = entry.GetExchangeDistributionList();
                    if (list != null)
                    {
                        string smtp = list.PrimarySmtpAddress;
                        Release(list);
                        if (!string.IsNullOrEmpty(smtp)) return smtp;
                    }

                    Outlook.PropertyAccessor accessor = entry.PropertyAccessor;
                    try
                    {
                        string smtp = accessor.GetProperty(SmtpAddressProperty) as string;
                        if (!string.IsNullOrEmpty(smtp)) return smtp;
                    }
                    catch (Exception)
                    {
                        // fall back to the address below
                    }
                    finally
                    {
                        Release(accessor);
                    }
                }
            }
            catch (Exception ex)
            {
                AddInLog.Write("Could not read the SMTP address of " + recipient.Name + ": " + ex.Message);
            }
            finally
            {
                Release(entry);
            }

            return recipient.Address;
        }

        /// <summary>
        /// Reads attachments into memory, in order. <paramref name="include"/> filters by file
        /// name (null reads them all). Each file is saved to a private temp folder, read and
        /// deleted straight away.
        /// </summary>
        public static List<MailFile> ReadAttachments(Outlook.MailItem mail, Func<string, bool> include)
        {
            var files = new List<MailFile>();
            Outlook.Attachments attachments = mail.Attachments;
            string tempDir = NewTempDir();
            try
            {
                int count = attachments.Count;
                for (int i = 1; i <= count; i++)
                {
                    Outlook.Attachment attachment = attachments[i];
                    try
                    {
                        string name = AttachmentName(attachment, i);
                        if (include != null && !include(name)) continue;

                        string dir = Path.Combine(tempDir, i.ToString(CultureInfo.InvariantCulture));
                        Directory.CreateDirectory(dir);
                        string path = Path.Combine(dir, MailFile.SafeFileName(name, i));
                        attachment.SaveAsFile(path);
                        files.Add(new MailFile(name, File.ReadAllBytes(path)));
                        File.Delete(path);
                    }
                    finally
                    {
                        Release(attachment);
                    }
                }
            }
            finally
            {
                Release(attachments);
                DeleteTempDir(tempDir);
            }
            return files;
        }

        /// <summary>
        /// Adds files as attachments. Names are made safe (they may come from another sender);
        /// each file goes through its own temp folder so Outlook keeps the name exactly.
        /// </summary>
        public static void AddAttachments(Outlook.MailItem mail, IReadOnlyList<MailFile> files)
        {
            if (files == null || files.Count == 0) return;

            Outlook.Attachments attachments = mail.Attachments;
            string tempDir = NewTempDir();
            try
            {
                for (int i = 0; i < files.Count; i++)
                {
                    string name = MailFile.SafeFileName(files[i].FileName, i + 1);
                    string dir = Path.Combine(tempDir, (i + 1).ToString(CultureInfo.InvariantCulture));
                    Directory.CreateDirectory(dir);
                    string path = Path.Combine(dir, name);
                    File.WriteAllBytes(path, files[i].Content);

                    Outlook.Attachment added = attachments.Add(path, Outlook.OlAttachmentType.olByValue, Type.Missing, name);
                    Release(added);
                    File.Delete(path);
                }
            }
            finally
            {
                Release(attachments);
                DeleteTempDir(tempDir);
            }
        }

        public static int AttachmentCount(Outlook.MailItem mail)
        {
            Outlook.Attachments attachments = mail.Attachments;
            try
            {
                return attachments.Count;
            }
            finally
            {
                Release(attachments);
            }
        }

        /// <summary>Removes attachments number <paramref name="firstIndex"/> (1-based) to the last.</summary>
        public static void RemoveAttachmentsFrom(Outlook.MailItem mail, int firstIndex)
        {
            Outlook.Attachments attachments = mail.Attachments;
            try
            {
                for (int i = attachments.Count; i >= Math.Max(1, firstIndex); i--)
                {
                    attachments.Remove(i);
                }
            }
            finally
            {
                Release(attachments);
            }
        }

        private static string AttachmentName(Outlook.Attachment attachment, int index)
        {
            string name = null;
            try
            {
                name = attachment.FileName;
            }
            catch (Exception)
            {
                // some attachment types (OLE objects) have no file name
            }
            if (string.IsNullOrWhiteSpace(name))
            {
                try
                {
                    name = attachment.DisplayName;
                }
                catch (Exception)
                {
                }
            }
            return string.IsNullOrWhiteSpace(name) ? "attachment" + index.ToString(CultureInfo.InvariantCulture) : name;
        }

        private static string NewTempDir()
        {
            string dir = Path.Combine(Path.GetTempPath(), "MpcOutlookAddIn", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            return dir;
        }

        private static void DeleteTempDir(string dir)
        {
            try
            {
                if (Directory.Exists(dir)) Directory.Delete(dir, true);
            }
            catch (IOException ex)
            {
                AddInLog.Write("Could not delete temp folder " + dir + ": " + ex.Message);
            }
            catch (UnauthorizedAccessException ex)
            {
                AddInLog.Write("Could not delete temp folder " + dir + ": " + ex.Message);
            }
        }

        public static void Release(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }

        private static string GetStringProperty(Outlook.MailItem mail, string schemaName)
        {
            Outlook.PropertyAccessor accessor = mail.PropertyAccessor;
            try
            {
                return accessor.GetProperty(schemaName) as string;
            }
            catch (Exception)
            {
                return null; // property not set
            }
            finally
            {
                Release(accessor);
            }
        }
    }
}

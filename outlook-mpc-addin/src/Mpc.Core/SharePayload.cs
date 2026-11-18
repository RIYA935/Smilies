using System;
using System.Text;

namespace Mpc.Core
{
    /// <summary>What a share file holds a share of.</summary>
    public enum ShareKind : byte
    {
        Body = 1,
        Attachment = 2,
    }

    /// <summary>
    /// One share of the body or of one attachment, as stored in a .bin file attached to an
    /// email-share. Layout (all integers little-endian):
    /// <code>
    /// offset  size  field
    ///   0      4    magic "SMLS"
    ///   4      1    format version (1)
    ///   5      1    kind: 1 = body, 2 = attachment
    ///   6      1    ring: 1 = Z_p, p = 2^61 - 1
    ///   7      1    form: 0 = full share
    ///   8      2    attachment number (1, 2, ...; 0 for the body)
    ///  10      8    length of the shared data in bytes
    ///  18      8n   the share: n = ceil(length / 7) field elements, 8 bytes each
    /// </code>
    /// </summary>
    public sealed class SharePayload
    {
        public const byte FormatVersion = 1;
        public const byte RingPrimeP61 = 1;
        public const byte FormFull = 0;
        public const int HeaderLength = 18;

        private static readonly byte[] Magic = Encoding.ASCII.GetBytes("SMLS");

        public SharePayload(ShareKind kind, int attachmentNumber, long dataLength, ulong[] elements)
        {
            if (kind != ShareKind.Body && kind != ShareKind.Attachment) throw new ArgumentOutOfRangeException("kind");
            if (attachmentNumber < 0 || attachmentNumber > ushort.MaxValue) throw new ArgumentOutOfRangeException("attachmentNumber");
            if (kind == ShareKind.Body && attachmentNumber != 0) throw new ArgumentException("The body has attachment number 0.", "attachmentNumber");
            if (kind == ShareKind.Attachment && attachmentNumber == 0) throw new ArgumentException("Attachments are numbered from 1.", "attachmentNumber");
            if (elements == null) throw new ArgumentNullException("elements");
            if (AdditiveSharing.ElementCount(dataLength) != elements.Length)
            {
                throw new ArgumentException(elements.Length + " elements can't hold exactly " + dataLength + " bytes.", "elements");
            }

            Kind = kind;
            AttachmentNumber = attachmentNumber;
            DataLength = dataLength;
            Elements = elements;
        }

        public ShareKind Kind { get; private set; }
        public int AttachmentNumber { get; private set; }
        public long DataLength { get; private set; }
        public ulong[] Elements { get; private set; }

        public byte[] Encode()
        {
            var bytes = new byte[HeaderLength + 8L * Elements.Length];
            Buffer.BlockCopy(Magic, 0, bytes, 0, 4);
            bytes[4] = FormatVersion;
            bytes[5] = (byte)Kind;
            bytes[6] = RingPrimeP61;
            bytes[7] = FormFull;
            WriteLittleEndian(bytes, 8, (ulong)AttachmentNumber, 2);
            WriteLittleEndian(bytes, 10, (ulong)DataLength, 8);
            for (int e = 0; e < Elements.Length; e++)
            {
                WriteLittleEndian(bytes, HeaderLength + 8 * e, Elements[e], 8);
            }
            return bytes;
        }

        /// <summary>Reads a share file, checking every field. Throws FormatException if it isn't one.</summary>
        public static SharePayload Decode(byte[] bytes)
        {
            if (bytes == null || bytes.Length < HeaderLength) throw new FormatException("Too short to be an MPC share file.");
            for (int k = 0; k < 4; k++)
            {
                if (bytes[k] != Magic[k]) throw new FormatException("Not an MPC share file (missing SMLS marker).");
            }
            if (bytes[4] != FormatVersion) throw new FormatException("Unsupported MPC share format version " + bytes[4] + ".");

            var kind = (ShareKind)bytes[5];
            if (kind != ShareKind.Body && kind != ShareKind.Attachment) throw new FormatException("Unknown share kind " + bytes[5] + ".");
            if (bytes[6] != RingPrimeP61) throw new FormatException("Unsupported field " + bytes[6] + " (only p = 2^61 - 1 is supported).");
            if (bytes[7] != FormFull) throw new FormatException("Unsupported share form " + bytes[7] + " (only full shares are supported).");

            int number = (int)ReadLittleEndian(bytes, 8, 2);
            if ((kind == ShareKind.Body) != (number == 0)) throw new FormatException("Share kind and attachment number don't match.");

            ulong length = ReadLittleEndian(bytes, 10, 8);
            if (length > int.MaxValue) throw new FormatException("Shared data length " + length + " is too large.");

            long count = AdditiveSharing.ElementCount((long)length);
            if (bytes.Length != HeaderLength + 8 * count)
            {
                throw new FormatException("Share file is " + bytes.Length + " bytes; " + count + " elements need " + (HeaderLength + 8 * count) + ".");
            }

            var elements = new ulong[count];
            for (long e = 0; e < count; e++)
            {
                ulong value = ReadLittleEndian(bytes, HeaderLength + 8 * (int)e, 8);
                PrimeField.CheckElement(value);
                elements[e] = value;
            }

            return new SharePayload(kind, number, (long)length, elements);
        }

        private static void WriteLittleEndian(byte[] buffer, long offset, ulong value, int size)
        {
            for (int k = 0; k < size; k++) buffer[offset + k] = (byte)(value >> (8 * k));
        }

        private static ulong ReadLittleEndian(byte[] buffer, int offset, int size)
        {
            ulong value = 0;
            for (int k = size - 1; k >= 0; k--) value = (value << 8) | buffer[offset + k];
            return value;
        }
    }
}

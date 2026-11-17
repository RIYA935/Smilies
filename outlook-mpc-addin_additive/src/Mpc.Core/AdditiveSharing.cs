using System;
using System.Collections.Generic;
using System.Security.Cryptography;

namespace Mpc.Core
{
    /// <summary>
    /// Additive secret sharing over Z_p (see <see cref="PrimeField"/>).
    /// <para>
    /// A secret x is split into n shares s_1..s_n with s_1..s_{n-1} uniformly random and
    /// s_n = x - (s_1 + ... + s_{n-1}) mod p. Any n - 1 shares are uniformly random and say
    /// nothing about x; adding all n gives x back. Each element of a vector is shared on its own.
    /// </para>
    /// </summary>
    public static class AdditiveSharing
    {
        /// <summary>Bytes packed into one field element. 7 bytes stay below 2^56, well below p.</summary>
        public const int BytesPerElement = 7;

        /// <summary>
        /// Splits a vector of field elements into <paramref name="shareCount"/> share vectors.
        /// With one share there is nothing to hide it among, so that share equals the secret.
        /// </summary>
        public static ulong[][] Split(ulong[] secret, int shareCount, RandomNumberGenerator rng)
        {
            if (secret == null) throw new ArgumentNullException("secret");
            if (shareCount < 1) throw new ArgumentOutOfRangeException("shareCount", "At least one share is required.");
            if (rng == null) throw new ArgumentNullException("rng");

            var shares = new ulong[shareCount][];
            for (int i = 0; i < shareCount; i++) shares[i] = new ulong[secret.Length];

            var random = new FieldRandom(rng);
            for (int j = 0; j < secret.Length; j++)
            {
                PrimeField.CheckElement(secret[j]);
                ulong sum = 0;
                for (int i = 0; i < shareCount - 1; i++)
                {
                    ulong r = random.Next();
                    shares[i][j] = r;
                    sum = PrimeField.Add(sum, r);
                }
                shares[shareCount - 1][j] = PrimeField.Subtract(secret[j], sum);
            }
            return shares;
        }

        /// <summary>Adds the share vectors element by element, mod p.</summary>
        public static ulong[] Reconstruct(IReadOnlyList<ulong[]> shares)
        {
            if (shares == null || shares.Count == 0) throw new ArgumentException("At least one share is required.", "shares");

            int length = shares[0].Length;
            foreach (ulong[] share in shares)
            {
                if (share == null || share.Length != length)
                {
                    throw new FormatException("The shares have different lengths, so they are not shares of the same data.");
                }
            }

            var result = new ulong[length];
            foreach (ulong[] share in shares)
            {
                for (int j = 0; j < length; j++)
                {
                    PrimeField.CheckElement(share[j]);
                    result[j] = PrimeField.Add(result[j], share[j]);
                }
            }
            return result;
        }

        /// <summary>Number of field elements needed for <paramref name="byteCount"/> bytes.</summary>
        public static long ElementCount(long byteCount)
        {
            if (byteCount < 0) throw new ArgumentOutOfRangeException("byteCount");
            return (byteCount + BytesPerElement - 1) / BytesPerElement;
        }

        /// <summary>
        /// Packs bytes into field elements, 7 bytes per element, little-endian. The last element
        /// is padded with zero bytes.
        /// </summary>
        public static ulong[] Pack(byte[] data)
        {
            if (data == null) throw new ArgumentNullException("data");

            var elements = new ulong[ElementCount(data.Length)];
            for (int e = 0; e < elements.Length; e++)
            {
                int start = e * BytesPerElement;
                int end = Math.Min(start + BytesPerElement, data.Length);
                ulong value = 0;
                for (int k = end - 1; k >= start; k--) value = (value << 8) | data[k];
                elements[e] = value;
            }
            return elements;
        }

        /// <summary>
        /// Turns field elements back into <paramref name="length"/> bytes. Throws when an element
        /// can't have come from <see cref="Pack"/> (too big, or non-zero padding), which is what
        /// adding shares from different messages, or damaged shares, almost always produces.
        /// </summary>
        public static byte[] Unpack(ulong[] elements, long length)
        {
            if (elements == null) throw new ArgumentNullException("elements");
            if (length < 0 || length > int.MaxValue) throw new FormatException("Invalid data length " + length + ".");
            if (ElementCount(length) != elements.Length)
            {
                throw new FormatException(elements.Length + " elements can't hold exactly " + length + " bytes.");
            }

            var data = new byte[length];
            for (int e = 0; e < elements.Length; e++)
            {
                ulong value = elements[e];
                if (value >> (8 * BytesPerElement) != 0)
                {
                    throw new FormatException("The rebuilt data is invalid: the shares don't belong together or are damaged.");
                }

                int start = e * BytesPerElement;
                for (int k = 0; k < BytesPerElement; k++)
                {
                    byte b = (byte)(value >> (8 * k));
                    if (start + k < length)
                    {
                        data[start + k] = b;
                    }
                    else if (b != 0)
                    {
                        throw new FormatException("The rebuilt data is invalid: the shares don't belong together or are damaged.");
                    }
                }
            }
            return data;
        }

        /// <summary>Uniformly random field elements from a cryptographic random number generator.</summary>
        private sealed class FieldRandom
        {
            private readonly RandomNumberGenerator _rng;
            private readonly byte[] _buffer = new byte[8 * 4096];
            private int _position;

            public FieldRandom(RandomNumberGenerator rng)
            {
                _rng = rng;
                _position = _buffer.Length;
            }

            public ulong Next()
            {
                while (true)
                {
                    if (_position == _buffer.Length)
                    {
                        _rng.GetBytes(_buffer);
                        _position = 0;
                    }

                    ulong raw = 0;
                    for (int k = 7; k >= 0; k--) raw = (raw << 8) | _buffer[_position + k];
                    _position += 8;

                    // Keep the low 61 bits: a uniform value in [0, 2^61 - 1]. Only 2^61 - 1
                    // itself is outside the field, so draw again in that (2^-61) case.
                    ulong value = raw & PrimeField.P;
                    if (value != PrimeField.P) return value;
                }
            }
        }
    }
}

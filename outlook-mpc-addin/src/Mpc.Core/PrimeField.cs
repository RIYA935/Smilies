using System;

namespace Mpc.Core
{
    /// <summary>
    /// Arithmetic in the prime field Z_p with p = 2^61 - 1 (a Mersenne prime).
    /// Every element is a ulong in [0, p). Two elements add to less than 2^62, so no step
    /// overflows a ulong.
    /// </summary>
    public static class PrimeField
    {
        /// <summary>p = 2^61 - 1 = 2,305,843,009,213,693,951.</summary>
        public const ulong P = (1UL << 61) - 1;

        public static bool IsElement(ulong value)
        {
            return value < P;
        }

        /// <summary>(a + b) mod p, for a and b already in [0, p).</summary>
        public static ulong Add(ulong a, ulong b)
        {
            ulong s = a + b;
            return s >= P ? s - P : s;
        }

        /// <summary>(a - b) mod p, for a and b already in [0, p).</summary>
        public static ulong Subtract(ulong a, ulong b)
        {
            return a >= b ? a - b : a + (P - b);
        }

        /// <summary>Throws unless <paramref name="value"/> is in [0, p).</summary>
        public static void CheckElement(ulong value)
        {
            if (value >= P) throw new FormatException("Value " + value + " is not an element of the field (it is not below 2^61 - 1).");
        }
    }
}

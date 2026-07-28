using System;
using System.Security.Cryptography;
using System.Text;

namespace QYachtMaster.Utils
{
    public static class SecurityHelper
    {
        private const int SaltSize = 16; // 128-bit
        private const int KeySize = 32;  // 256-bit
        private const int Iterations = 10000;

        /// <summary>
        /// Hashes a plain-text password using PBKDF2 (Modern implementation).
        /// </summary>
        public static string HashPassword(string password)
        {
            var salt = RandomNumberGenerator.GetBytes(SaltSize);
            var key = Rfc2898DeriveBytes.Pbkdf2(
                password,
                salt,
                Iterations,
                HashAlgorithmName.SHA256,
                KeySize);

            var bytes = new byte[SaltSize + KeySize];
            Buffer.BlockCopy(salt, 0, bytes, 0, SaltSize);
            Buffer.BlockCopy(key, 0, bytes, SaltSize, KeySize);

            return Convert.ToBase64String(bytes);
        }

        /// <summary>
        /// Verifies a password against an existing hash.
        /// </summary>
        public static bool VerifyPassword(string password, string hashedPassword)
        {
            try
            {
                var bytes = Convert.FromBase64String(hashedPassword);
                var salt = new byte[SaltSize];
                var key = new byte[KeySize];

                Buffer.BlockCopy(bytes, 0, salt, 0, SaltSize);
                Buffer.BlockCopy(bytes, SaltSize, key, 0, KeySize);

                var testKey = Rfc2898DeriveBytes.Pbkdf2(
                    password,
                    salt,
                    Iterations,
                    HashAlgorithmName.SHA256,
                    KeySize);

                // Cryptographic fixed-time comparison
                return CryptographicOperations.FixedTimeEquals(key, testKey);
            }
            catch
            {
                return false;
            }
        }
    }
}
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    /// <summary>Small DPAPI-backed store for account/session material. Raw secrets never go to logs or plaintext files.</summary>
    internal sealed class ProtectedSessionStore
    {
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("ChuanHoa.AccountSession.v1");
        private readonly string _path;

        public ProtectedSessionStore(string directory, string fileName = "account-session.dat")
        {
            Directory.CreateDirectory(directory);
            _path = Path.Combine(directory, fileName);
        }

        public void Save(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("SESSION_REQUIRED", nameof(value));
            var clear = Encoding.UTF8.GetBytes(value);
            var protectedBytes = ProtectedData.Protect(clear, Entropy, DataProtectionScope.CurrentUser);
            var temp = _path + ".new";
            File.WriteAllBytes(temp, protectedBytes);
            if (File.Exists(_path)) File.Replace(temp, _path, null); else File.Move(temp, _path);
        }

        public string? Load()
        {
            if (!File.Exists(_path)) return null;
            try
            {
                var clear = ProtectedData.Unprotect(File.ReadAllBytes(_path), Entropy, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(clear);
            }
            catch (CryptographicException) { return null; }
            catch (IOException) { return null; }
        }

        public void Clear()
        {
            if (File.Exists(_path)) File.Delete(_path);
        }
    }
}

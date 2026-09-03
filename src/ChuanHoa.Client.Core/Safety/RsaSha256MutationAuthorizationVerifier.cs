using System;
using System.Security.Cryptography;

namespace ChuanHoa.Client.Core.Safety
{
    public interface IMutationAuthorizationPublicKeyProvider
    {
        bool TryGetPublicKey(string keyId, out RSAParameters publicKey);
    }

    public sealed class RsaSha256MutationAuthorizationVerifier : IMutationAuthorizationVerifier
    {
        private const string SupportedAlgorithm = "RS256";
        private readonly IMutationAuthorizationPublicKeyProvider _keyProvider;

        public RsaSha256MutationAuthorizationVerifier(IMutationAuthorizationPublicKeyProvider keyProvider)
        {
            _keyProvider = keyProvider ?? throw new ArgumentNullException(nameof(keyProvider));
        }

        public bool IsAuthentic(MutationAuthorization authorization)
        {
            if (authorization == null ||
                string.IsNullOrWhiteSpace(authorization.KeyId) ||
                !string.Equals(authorization.SignatureAlgorithm, SupportedAlgorithm, StringComparison.Ordinal) ||
                string.IsNullOrWhiteSpace(authorization.Signature) ||
                string.IsNullOrWhiteSpace(authorization.SubjectUserId) ||
                string.IsNullOrWhiteSpace(authorization.DeviceId) ||
                string.IsNullOrWhiteSpace(authorization.DocumentRevision) ||
                string.IsNullOrWhiteSpace(authorization.Nonce))
            {
                return false;
            }

            RSAParameters publicKey;
            if (!_keyProvider.TryGetPublicKey(authorization.KeyId, out publicKey))
            {
                return false;
            }

            byte[] signature;
            try
            {
                signature = DecodeBase64Url(authorization.Signature);
            }
            catch (FormatException)
            {
                return false;
            }

            try
            {
                using (var rsa = RSA.Create())
                {
                    rsa.ImportParameters(publicKey);
                    return rsa.VerifyData(
                        MutationAuthorizationCanonicalizer.GetCanonicalBytes(authorization),
                        signature,
                        HashAlgorithmName.SHA256,
                        RSASignaturePadding.Pkcs1);
                }
            }
            catch (CryptographicException)
            {
                return false;
            }
        }

        private static byte[] DecodeBase64Url(string value)
        {
            var normalized = value.Replace('-', '+').Replace('_', '/');
            switch (normalized.Length % 4)
            {
                case 0:
                    break;
                case 2:
                    normalized += "==";
                    break;
                case 3:
                    normalized += "=";
                    break;
                default:
                    throw new FormatException("Invalid base64url signature length.");
            }

            return Convert.FromBase64String(normalized);
        }
    }
}

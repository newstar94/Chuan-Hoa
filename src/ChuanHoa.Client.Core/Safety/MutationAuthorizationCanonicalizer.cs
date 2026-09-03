using System;
using System.IO;
using System.Text;

namespace ChuanHoa.Client.Core.Safety
{
    public static class MutationAuthorizationCanonicalizer
    {
        private static readonly byte[] Header = Encoding.ASCII.GetBytes("chuanhoa-mutation-authorization-v1");

        public static byte[] GetCanonicalBytes(MutationAuthorization authorization)
        {
            if (authorization == null)
            {
                throw new ArgumentNullException(nameof(authorization));
            }

            using (var stream = new MemoryStream())
            using (var writer = new BinaryWriter(stream, Encoding.UTF8, true))
            {
                WriteBytes(writer, Header);
                WriteString(writer, authorization.Kind.ToString());
                WriteString(writer, authorization.Schema);
                WriteString(writer, authorization.CommandId);
                WriteString(writer, authorization.DocumentFingerprint);
                WriteString(writer, authorization.Scope);
                writer.Write(authorization.NotBeforeUtc.ToUniversalTime().Ticks);
                writer.Write(authorization.ExpiresAtUtc.ToUniversalTime().Ticks);
                WriteString(writer, authorization.Jti);
                WriteString(writer, authorization.KeyId);
                WriteString(writer, authorization.SignatureAlgorithm);
                WriteString(writer, authorization.SubjectUserId);
                WriteString(writer, authorization.DeviceId);
                WriteString(writer, authorization.DocumentRevision);
                WriteString(writer, authorization.Nonce);
                writer.Flush();
                return stream.ToArray();
            }
        }

        private static void WriteString(BinaryWriter writer, string value)
        {
            WriteBytes(writer, Encoding.UTF8.GetBytes(value ?? string.Empty));
        }

        private static void WriteBytes(BinaryWriter writer, byte[] value)
        {
            writer.Write(value.Length);
            writer.Write(value);
        }
    }
}

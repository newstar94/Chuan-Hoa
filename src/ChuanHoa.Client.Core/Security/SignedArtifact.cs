using System;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace ChuanHoa.Client.Core.Security
{
    public sealed class SignedArtifact
    {
        public SignedArtifact(string kind, string keyId, byte[] payload, byte[] signature)
        {
            Kind = Require(kind, nameof(kind));
            KeyId = Require(keyId, nameof(keyId));
            Payload = payload ?? throw new ArgumentNullException(nameof(payload));
            Signature = signature ?? throw new ArgumentNullException(nameof(signature));
        }

        public string Kind { get; }
        public string KeyId { get; }
        public byte[] Payload { get; }
        public byte[] Signature { get; }

        private static string Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A value is required.", name);
            return value;
        }
    }

    public static class SignedArtifactCodec
    {
        public const string EnvelopeSchema = "chuanhoa.signed-artifact.v1";

        public static string Encode(SignedArtifact artifact)
        {
            if (artifact == null) throw new ArgumentNullException(nameof(artifact));
            return new XDocument(
                new XElement("signedArtifact",
                    new XAttribute("schema", EnvelopeSchema),
                    new XAttribute("kind", artifact.Kind),
                    new XAttribute("keyId", artifact.KeyId),
                    new XAttribute("algorithm", "RS256"),
                    new XElement("payload", Convert.ToBase64String(artifact.Payload)),
                    new XElement("signature", Convert.ToBase64String(artifact.Signature))))
                .ToString(SaveOptions.DisableFormatting);
        }

        public static SignedArtifact Decode(string encoded)
        {
            if (string.IsNullOrWhiteSpace(encoded)) throw new FormatException("Signed artifact is empty.");
            var root = XDocument.Parse(encoded, LoadOptions.PreserveWhitespace).Root;
            if (root == null || root.Name.LocalName != "signedArtifact") throw new FormatException("Signed artifact root is invalid.");
            if ((string)root.Attribute("schema") != EnvelopeSchema) throw new FormatException("Signed artifact schema is unsupported.");
            if ((string)root.Attribute("algorithm") != "RS256") throw new FormatException("Signed artifact algorithm is unsupported.");
            EnsureOnly(root, "payload", "signature");
            var payload = RequiredElement(root, "payload");
            var signature = RequiredElement(root, "signature");
            return new SignedArtifact(
                RequiredAttribute(root, "kind"),
                RequiredAttribute(root, "keyId"),
                Convert.FromBase64String(payload.Value),
                Convert.FromBase64String(signature.Value));
        }

        public static byte[] CanonicalPayload(XElement payload)
        {
            if (payload == null) throw new ArgumentNullException(nameof(payload));
            return Encoding.UTF8.GetBytes(payload.ToString(SaveOptions.DisableFormatting));
        }

        private static void EnsureOnly(XElement root, params string[] names)
        {
            foreach (var element in root.Elements())
            {
                if (Array.IndexOf(names, element.Name.LocalName) < 0)
                    throw new FormatException("Unknown signed artifact element: " + element.Name.LocalName + ".");
            }
        }

        private static XElement RequiredElement(XElement root, string name)
        {
            var elements = root.Elements(name);
            XElement? found = null;
            foreach (var element in elements)
            {
                if (found != null) throw new FormatException("Duplicate element: " + name + ".");
                found = element;
            }
            return found ?? throw new FormatException("Missing element: " + name + ".");
        }

        internal static string RequiredAttribute(XElement element, string name)
        {
            var value = (string)element.Attribute(name);
            if (string.IsNullOrWhiteSpace(value)) throw new FormatException("Missing attribute: " + name + ".");
            return value;
        }
    }

    public sealed class RsaSha256ArtifactVerifier
    {
        private readonly string _keyId;
        private readonly RSAParameters _publicKey;

        public RsaSha256ArtifactVerifier(string keyId, string publicKeyXml)
        {
            _keyId = string.IsNullOrWhiteSpace(keyId) ? throw new ArgumentException("Key id is required.", nameof(keyId)) : keyId;
            _publicKey = ParsePublicKey(publicKeyXml);
        }

        public byte[] Verify(string encoded, string expectedKind)
        {
            var artifact = SignedArtifactCodec.Decode(encoded);
            if (!string.Equals(artifact.Kind, expectedKind, StringComparison.Ordinal))
                throw new CryptographicException("Signed artifact kind does not match.");
            if (!string.Equals(artifact.KeyId, _keyId, StringComparison.Ordinal))
                throw new CryptographicException("Signed artifact key id is not trusted.");
            using (var rsa = RSA.Create())
            {
                rsa.ImportParameters(_publicKey);
                if (!rsa.VerifyData(artifact.Payload, artifact.Signature, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1))
                    throw new CryptographicException("Signed artifact signature is invalid.");
            }
            return artifact.Payload;
        }

        public static string ExportPublicKeyXml(RSA rsa)
        {
            if (rsa == null) throw new ArgumentNullException(nameof(rsa));
            var key = rsa.ExportParameters(false);
            return new XElement("RSAKeyValue",
                new XElement("Modulus", Convert.ToBase64String(key.Modulus)),
                new XElement("Exponent", Convert.ToBase64String(key.Exponent)))
                .ToString(SaveOptions.DisableFormatting);
        }

        private static RSAParameters ParsePublicKey(string xml)
        {
            if (string.IsNullOrWhiteSpace(xml)) throw new FormatException("Public key is empty.");
            var root = XElement.Parse(xml, LoadOptions.None);
            if (root.Name.LocalName != "RSAKeyValue") throw new FormatException("Public key root is invalid.");
            return new RSAParameters
            {
                Modulus = Convert.FromBase64String(root.Element("Modulus")?.Value ?? throw new FormatException("Public key modulus is missing.")),
                Exponent = Convert.FromBase64String(root.Element("Exponent")?.Value ?? throw new FormatException("Public key exponent is missing."))
            };
        }
    }
}

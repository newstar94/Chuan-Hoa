using System;

namespace ChuanHoa.Client.Core.Safety
{
    public sealed class MutationSafetyException : Exception
    {
        public MutationSafetyException(string code, string message)
            : base(message)
        {
            Code = code;
        }

        public MutationSafetyException(string code, string message, Exception innerException)
            : base(message, innerException)
        {
            Code = code;
        }

        public string Code { get; }
    }
}

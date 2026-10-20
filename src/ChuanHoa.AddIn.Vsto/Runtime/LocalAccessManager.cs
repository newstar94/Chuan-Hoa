using System;
using System.Linq;
using ChuanHoa.Client.Core.Licensing;
using ChuanHoa.Client.Core.Rules;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    /// <summary>Local rule availability, not user authorization. No server or session.</summary>
    public sealed class LocalAccessManager : IDisposable
    {
        public const string FormatFeature = "FORMAT_SCAN";
        public const string SpellingFeature = "SPELLING_SCAN";
        public const string DocumentToolsFeature = "DOCUMENT_TOOLS";
        public const string AutoFixFeature = "AUTOFIX";
        public const string TableImageToolsFeature = "TABLE_IMAGE_TOOLS";
        public const string FullAccessFeature = "FULL_ACCESS";
        private readonly Lazy<LocalRulePack> _rules;
        private bool _disposed;

        public LocalAccessManager(string clientReleaseId)
        {
            _rules = new Lazy<LocalRulePack>(() => BundledRulePack.Load(clientReleaseId));
        }

        public event EventHandler? CacheStateChanged;
        public bool IsRefreshInProgress => false;
        public bool CanUseCommand(string commandId) => !_disposed &&
            (ProductFeaturePolicy.RequiredFeature(commandId).Length == 0 ||
             HasCachedFeature(ProductFeaturePolicy.RequiredFeature(commandId)));

        public bool HasCachedFeature(string feature) =>
            !_disposed && ProductFeaturePolicy.FreeFeatures.Contains(feature);

        public void RequireCommand(string commandId)
        {
            var feature = ProductFeaturePolicy.RequiredFeature(commandId);
            if (feature.Length != 0) GetRulePack(feature);
        }

        public LocalRulePack GetRulePack(string feature)
        {
            if (_disposed) throw new ObjectDisposedException(nameof(LocalAccessManager));
            if (!HasCachedFeature(feature)) throw new InvalidOperationException("UNKNOWN_PRODUCT_FEATURE");
            return _rules.Value;
        }

        public LocalRulePack GetTableImageRulePack() => GetRulePack(TableImageToolsFeature);
        public string DescribeStatus() => "Chuẩn Hóa miễn phí · Không cần tài khoản · Bộ quy tắc đi kèm ứng dụng.";
        public void WarmUp()
        {
            GetRulePack(FormatFeature);
            CacheStateChanged?.Invoke(this, EventArgs.Empty);
        }
        public void Dispose() { _disposed = true; }
    }
}

using System;
using ChuanHoa.AddIn.Vsto.Runtime;

namespace ChuanHoa.DevelopmentAccessSmoke
{
    internal static class Program
    {
        private static int Main(string[] args)
        {
            try
            {
                using (var access = new LocalAccessManager(
                    typeof(LocalAccessManager).Assembly.GetName().Version.ToString()))
                {
                    foreach (var feature in new[] { LocalAccessManager.FormatFeature,
                        LocalAccessManager.SpellingFeature, LocalAccessManager.AutoFixFeature,
                        LocalAccessManager.DocumentToolsFeature, LocalAccessManager.TableImageToolsFeature })
                    {
                        var pack = access.GetRulePack(feature);
                        if (pack.Lexicon.Count < 1000 || pack.ExpiresAtUtc != DateTimeOffset.MaxValue)
                            throw new InvalidOperationException("Bundled free rule pack is incomplete.");
                    }
                    if (access.HasCachedFeature("UNKNOWN"))
                        throw new InvalidOperationException("Unknown feature accepted.");
                    Console.WriteLine("DEVELOPMENT_ACCESS_SMOKE_PASS FREE_OFFLINE_BUNDLED_RULES");
                }
                return 0;
            }
            catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        }
    }
}

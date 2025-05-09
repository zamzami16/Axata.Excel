using System.Text;

namespace Axata.Excel;

[ExcludeFromCodeCoverage]
internal static class EncodingInitialization
{
    private static readonly Lazy<bool> _isInitialized = new(() =>
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return true;
    });

    public static void EnsureInitialized()
    {
        _ = _isInitialized.Value;
    }
}

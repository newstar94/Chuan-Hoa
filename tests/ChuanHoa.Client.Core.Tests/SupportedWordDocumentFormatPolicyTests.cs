using ChuanHoa.Client.Core.Safety;

namespace ChuanHoa.Client.Core.Tests;

public sealed class SupportedWordDocumentFormatPolicyTests
{
    [Theory]
    [InlineData(@"C:\VanBan\quyet-dinh.doc", 0)]
    [InlineData(@"C:\VanBan\QUYET-DINH.DOC", 0)]
    [InlineData(@"C:\VanBan\quyet-dinh.docx", 12)]
    [InlineData(@"C:\VanBan\QUYET-DINH.DOCX", 12)]
    [InlineData(@"C:\VanBan\quyet-dinh.docx", 24)]
    public void Accepts_saved_doc_and_docx_documents(string fullName, int saveFormat)
    {
        Assert.True(SupportedWordDocumentFormatPolicy.IsSupported(fullName, saveFormat));
    }

    [Theory]
    [InlineData(@"C:\VanBan\quyet-dinh.docm", 13)]
    [InlineData(@"C:\VanBan\quyet-dinh.dotx", 14)]
    [InlineData(@"C:\VanBan\quyet-dinh.rtf", 6)]
    [InlineData(@"C:\VanBan\quyet-dinh.doc", 12)]
    [InlineData(@"C:\VanBan\quyet-dinh.docx", 0)]
    [InlineData("", 12)]
    public void Rejects_unsupported_or_mismatched_document_formats(string fullName, int saveFormat)
    {
        Assert.False(SupportedWordDocumentFormatPolicy.IsSupported(fullName, saveFormat));
    }
}

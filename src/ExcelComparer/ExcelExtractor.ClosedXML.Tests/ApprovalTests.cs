
using ApprovalTests;
using ApprovalTests.Reporters;

namespace ExcelExtractor.ClosedXML.Tests;

[UseReporter(typeof(VisualStudioCodeReporter))]
public class ApprovalTests
{
    [Fact]
    public void ExtractFile()
    {
        var fileName = "C:\\Users\\Phong.NguyenDoan\\Downloads\\Temp\\xxx.xlsx";

        using var fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
        var extractor = new ClosedXmlExcelExtractor();
        var workbook = extractor.Extract(fileStream);
        Approvals.Verify(workbook.ToJson());
    }

    [Fact]
    public void Compare2Files()
    {
        var fileName1 = "C:\\Users\\Phong.NguyenDoan\\Downloads\\Temp\\xxx.xlsx";
        var fileName2 = "C:\\Users\\Phong.NguyenDoan\\Downloads\\Temp\\xxx1.xlsx";

        using var fileStream1 = new FileStream(fileName1, FileMode.Open, FileAccess.Read);
        var extractor1 = new ClosedXmlExcelExtractor();
        var workbook1 = extractor1.Extract(fileStream1);


        using var fileStream2 = new FileStream(fileName2, FileMode.Open, FileAccess.Read);
        var extractor2 = new ClosedXmlExcelExtractor();
        var workbook2 = extractor2.Extract(fileStream2);

        Approvals.AssertText(workbook1.ToJson(), workbook2.ToJson());
    }
}
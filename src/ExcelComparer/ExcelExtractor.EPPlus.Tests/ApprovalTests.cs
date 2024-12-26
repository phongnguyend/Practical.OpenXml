using ApprovalTests.Reporters;
using ApprovalTests;

namespace ExcelExtractor.EPPlus.Tests;

[UseReporter(typeof(VisualStudioCodeReporter))]
public class ApprovalTests
{
    [Fact]
    public void Extract()
    {
        var fileName = "C:\\Users\\Phong.NguyenDoan\\Downloads\\xxx.xlsx";

        using var fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
        var extractor = new EPPlusExcelExtractor();
        var workbook = extractor.Extract(fileStream);
        Approvals.Verify(workbook.ToJson());
    }
}
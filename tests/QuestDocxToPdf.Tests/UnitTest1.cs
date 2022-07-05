using System.Diagnostics;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using QuestDocxToPdf.Core;
using QuestPDF.Fluent;
using Xunit;

namespace QuestDocxToPdf.Tests;

public class UnitTest1
{
    [Fact]
    public void Test1()
    {
        using var doc = WordprocessingDocument.Open("mismatch.docx", false);
        var a = new DocXDocument(doc, new DocXGenerationOptions());
        var stopwatch = Stopwatch.StartNew();
        var z = a.GeneratePdf();
        //var zz = a.GenerateImages();
        stopwatch.Stop();
        var b = stopwatch.Elapsed.Milliseconds;
        File.WriteAllBytes("../../../../../test.pdf", z);
        // File.WriteAllBytes("../../../../../image.png", zz.First());
    }
}

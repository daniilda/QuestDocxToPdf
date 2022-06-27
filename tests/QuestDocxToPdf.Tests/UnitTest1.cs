using System.IO;
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
        using var doc = WordprocessingDocument.Open("test.docx", false);
        var q = doc.MainDocumentPart.Document.Body.ChildElements;
        foreach (var el in q)
        {
            
        }
        var a = new DocXDocument(doc, new DocXGenerationOptions());
        var z = a.GeneratePdf();
        File.WriteAllBytes("../../../../../test.pdf", z);
    }
}
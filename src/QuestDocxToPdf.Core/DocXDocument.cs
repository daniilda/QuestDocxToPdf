using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using QuestPDF.Drawing;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;

namespace QuestDocxToPdf.Core;

public class DocXDocument : IDocument
{
    public DocXDocument(WordprocessingDocument document, DocXGenerationOptions options)
    {
        Document = document;
        Options = options;
    } 

    public WordprocessingDocument Document { get; }
    public DocXGenerationOptions Options { get; }

    public DocumentMetadata GetMetadata() => DocumentMetadata.Default;

    public void Compose(IDocumentContainer container)
    {
        ComposeHeader(container);
    }

    private void ComposeHeader(IDocumentContainer container)
    {
        container.Page(page =>
        {
            var header = page.Header();
            var headerPart = Document.MainDocumentPart.HeaderParts.FirstOrDefault();
            var q = headerPart!.Header.GetAttributes();
            var headerElements = headerPart!.Header;
        });
    }

    private void ResolveNodes(IContainer container, IEnumerable<OpenXmlElement> nodes)
    {
        foreach (var node in nodes)
            ResolveNode(container, node);
    }

    private void ResolveNode(IContainer container, OpenXmlElement node)
    {
        switch (node.LocalName)
        {
            case "tblPr":
            {
                var xmlTableProperties = (TableProperties) node;
                break;
            }
            case "tblGrid":
            {
                var xmlTableGrid = (TableGrid) node;
                break;
            }
            case "tr":
            {
                var xmlTableRow = (TableRow) node;
                break;
            }
            case "p":
            {
                break;
            }
        }
    }
}
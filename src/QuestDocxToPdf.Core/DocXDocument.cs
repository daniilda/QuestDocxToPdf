using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using QuestPDF.Drawing;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
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
        container.Page(
            page =>
            {
                page.Header().Column(ComposeHeader);
                page.Content().Column(ComposeBody);
                page.Footer().Column(ComposeFooter);
            });
    }

    private void ComposeHeader(ColumnDescriptor descriptor)
    {
        var xmlHeader = Document.MainDocumentPart.HeaderParts.FirstOrDefault()?.Header;
        if (xmlHeader is null)
            return;
        ResolveNode(descriptor, xmlHeader.ChildElements);
    }

    private void ComposeFooter(ColumnDescriptor descriptor)
    {
        var xmlFooter = Document.MainDocumentPart.FooterParts.FirstOrDefault()?.Footer;
        if (xmlFooter is null)
            return;
        ResolveNode(descriptor, xmlFooter.ChildElements);
    }

    private void ComposeBody(ColumnDescriptor descriptor)
    {
        var xmlBody = Document.MainDocumentPart.Document.Body;
        if (xmlBody is null)
            return;
        ResolveNode(descriptor, xmlBody.ChildElements);
    }

    private void ResolveNode(ColumnDescriptor descriptor, OpenXmlElementList nodes)
    {
        foreach (var node in nodes)
            ResolveNode(descriptor, node);
    }

    private void ResolveNode(ColumnDescriptor descriptor, OpenXmlElement node)
    {
        switch (node.LocalName)
        {
            // case "tbl":
            // {
            //     var xmlTable = (Table)node;
            //
            //     descriptor.Item().Table(
            //         table =>
            //         {
            //             ResolveTableNode(table, xmlTable.ChildElements);
            //         });
            //     break;
            // }
            case "p":
            {
                var xmlParagraph = (Paragraph)node;
                ResolveNode(descriptor, xmlParagraph.ChildElements);
                break;
            }
            case "pPr":
            {
                var xmlParagraphProperties = (ParagraphProperties)node;
                break;
            }
            case "r":
            {
                var xmlRun = (Run)node;
                descriptor.Item().Text(
                    text =>
                    {
                        text.Line(xmlRun.InnerText);
                    });
                break;
            }
        }
    }

    private void ResolveTableNode(TableDescriptor table, OpenXmlElementList nodes)
    {
        foreach (var node in nodes)
            ResolveTableNode(table, node);
    }

    private void ResolveTableNode(TableDescriptor table, OpenXmlElement node)
    {
        switch (node.LocalName)
        {
            case "tblPr":
            {
                var xmlTableProperties = (TableProperties)node;
                break;
            }
            case "tblGrid":
            {
                var xmlTableGrid = (TableGrid)node;
                break;
            }
            case "tr":
            {
                var xmlTableRow = (TableRow)node;
                break;
            }
            case "trPr":
            {
                var xmlTableRowProperties = (TableRowProperties)node;
                break;
            }
            case "tc":
            {
                var xmlTableCell = (TableCell)node;
                break;
            }
            case "tcPr":
            {
                var xmlTableCellProperties = (TableCellProperties)node;

                break;
            }
        }
        ResolveTableNode(table, node.ChildElements);
    }
}

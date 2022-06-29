using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using QuestPDF.Drawing;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using Colors = QuestPDF.Helpers.Colors;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using GridColumn = DocumentFormat.OpenXml.Wordprocessing.GridColumn;
using PageSize = DocumentFormat.OpenXml.Wordprocessing.PageSize;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Picture = DocumentFormat.OpenXml.Wordprocessing.Picture;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableGrid = DocumentFormat.OpenXml.Wordprocessing.TableGrid;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace QuestDocxToPdf.Core;

public class DocXDocument : IDocument
{
    private readonly DocumentMetadata _metadata;
    private const float DocxToQuestPDFScale = 19.99832045683574f;
    private const float DocxToQuestPDFFontScale = 2f;

    public DocXDocument(WordprocessingDocument document, DocXGenerationOptions options)
    {
        Document = document;
        Options = options;
        _metadata = new DocumentMetadata()
        {
            DocumentLayoutExceptionThreshold = 100000
        };
    }

    public WordprocessingDocument Document { get; }

    public DocXGenerationOptions Options { get; }

    public DocumentMetadata GetMetadata() => _metadata;

    public void Compose(IDocumentContainer container)
    {
        container.Page(
            page =>
            {
                var fillcolor = Document.MainDocumentPart.Document.DocumentBackground?.Background?.Fillcolor;
                var a = Document.MainDocumentPart.Document.Body.ChildElements.FirstOrDefault(
                    x => x.LocalName == "sectPr");
                var b = (SectionProperties?)a;
                var xmlMargins = b?.ChildElements.FirstOrDefault(x => x.LocalName == "pgMar");
                var xmlSize = b?.ChildElements.FirstOrDefault(x => x.LocalName == "pgSz");
                var size = (PageSize?)xmlSize;
                var aa = new QuestPDF.Helpers.PageSize(size!.Width!/DocxToQuestPDFScale, size!.Height!/DocxToQuestPDFScale);
                var d = (PageMargin?) xmlMargins;
                page.Background().Background(Colors.White);
                page.Size(aa);
                page.MarginLeft(d?.Left/DocxToQuestPDFScale);
                page.MarginRight(d?.Right/DocxToQuestPDFScale);
                page.MarginTop(d?.Top/DocxToQuestPDFScale);
                page.MarginBottom(d?.Bottom/DocxToQuestPDFScale);
                page.Header().Column(ComposeHeader);
                page.Content().Column(ComposeBody);
                page.Footer().Column(ComposeFooter);
            });
        var a = 1;
        Console.WriteLine(a);
    }

    private void ComposeHeader(ColumnDescriptor descriptor)
    {
        descriptor.Spacing(0);
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
            case "tbl":
            {
                var xmlTable = (Table)node;
                var xmlTableGrid = node.ChildElements.OfType<TableGrid>().FirstOrDefault();
                var columns = xmlTableGrid!.ChildElements.OfType<GridColumn>();
                var enumerable = columns as GridColumn[] ?? columns.ToArray();
                uint rowC = 1;
                uint columnC = 1;
                descriptor.Item().Table(
                    table =>
                    {
                        table.ColumnsDefinition(x =>
                        {
                            foreach (var column in enumerable)
                            {
                                x.RelativeColumn(float.Parse(column.Width));
                            }
                        });
                        var tableRows = xmlTable.ChildElements.OfType<TableRow>();
                        foreach (var tableRow in tableRows)
                        {
                            var tableRowCells = tableRow.ChildElements.OfType<TableCell>();
                            foreach (var tableCell in tableRowCells)
                            {
                                var gripSpan = (uint) (tableCell.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
                                var columnSpanValue = tableCell.TableCellProperties?.VerticalMerge?.Val?.Value;
                                table.Cell().Row(rowC).Column(columnC).ColumnSpan(gripSpan).Border(0.01f).Column(x =>
                                {
                                    ResolveNode(x, tableCell.ChildElements);
                                });

                                columnC += gripSpan;
                            }
                            columnC = 1;
                            rowC++;
                        }
                    });
                descriptor.Spacing(1);
                break;
            }
            case "p":
            {
                var xmlParagraph = (Paragraph)node;
                descriptor.Item().Text(
                    text
                        =>
                    {
                        ResolveTextNode(text, xmlParagraph.ChildElements);
                    });
                break;
            }
        }
    }

    private void ResolveTextNode(TextDescriptor text, OpenXmlElement node)
    {
        switch (node.LocalName)
        {
            case "pPr":
            {
                var xmlParagraphProperties = (ParagraphProperties)node;
                var justification = xmlParagraphProperties.Justification?.Val?.Value;
                switch (justification)
                {
                    case JustificationValues.Center:
                        text.AlignCenter();
                        break;
                    case JustificationValues.Left:
                        text.AlignLeft();
                        break;
                    case JustificationValues.Right:
                        text.AlignRight();
                        break;
                    case JustificationValues.Both: // Я Хз как это реализовать, а оно вообще надо?)
                        text.AlignLeft();
                        break;
                }

                break;
            }
            case "r":
            {
                var xmlRun = (Run)node;
                if (xmlRun.ChildElements.FirstOrDefault(x => x.LocalName == "t") is not null)
                {
                    if (xmlRun.RunProperties?.Spacing?.Val is not null)
                    {
                        var val = xmlRun.RunProperties.Spacing.Val;
                        text.ParagraphSpacing(val);
                    }
                    var spanDescriptor =
                        text.Span(xmlRun.ChildElements.FirstOrDefault(x => x.LocalName == "t")!.InnerText)
                            .WrapAnywhere();
                    if (xmlRun.RunProperties?.RunFonts?.Ascii?.Value is not null)
                        spanDescriptor.FontFamily(xmlRun.RunProperties.RunFonts.Ascii.Value);
                    if (xmlRun.RunProperties?.Color?.Val.Value is not null)
                    {
                        var val = xmlRun.RunProperties.Color.Val.Value;
                        if (val == "auto")
                            break;
                        spanDescriptor.FontColor(val);
                    }

                    if (xmlRun.RunProperties?.FontSize?.Val is not null)
                    {
                        var val = xmlRun.RunProperties.FontSize.Val;
                        if (val == "auto")
                            break;
                        spanDescriptor.FontSize(float.Parse(val) / DocxToQuestPDFFontScale);
                    }
                }

                if (xmlRun.ChildElements.FirstOrDefault(x => x.LocalName == "drawing") is not null)
                {
                    var drawing = (Drawing)xmlRun.ChildElements.First(x => x.LocalName == "drawing");
                    var picture = drawing.Inline.Graphic.GraphicData.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().FirstOrDefault();
                    var blip = picture?.BlipFill.Blip?.Embed?.Value;
                    if (blip is null)
                        break;
                    var img = (ImagePart)Document.MainDocumentPart.GetPartById(blip);
                    using var imageStream = img.GetStream();
                    text.Element().Image(imageStream);
                }

                ResolveTextNode(text, xmlRun.ChildElements);
                break;
            }
            case "br":
            {
                text.EmptyLine();
                break;
            }
        }
    }

    private void ResolveTextNode(TextDescriptor text, OpenXmlElementList nodes)
    {
        foreach (var node in nodes)
            ResolveTextNode(text, node);
    }
}

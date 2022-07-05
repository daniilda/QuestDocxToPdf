using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using QuestDocxToPdf.Core.DescriptorResolvers;
using QuestPDF.Drawing;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using BottomBorder = DocumentFormat.OpenXml.Drawing.BottomBorder;
using Colors = QuestPDF.Helpers.Colors;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using GridColumn = DocumentFormat.OpenXml.Wordprocessing.GridColumn;
using PageSize = DocumentFormat.OpenXml.Wordprocessing.PageSize;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Picture = DocumentFormat.OpenXml.Wordprocessing.Picture;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableGrid = DocumentFormat.OpenXml.Wordprocessing.TableGrid;
using TableProperties = DocumentFormat.OpenXml.Wordprocessing.TableProperties;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using TableStyle = DocumentFormat.OpenXml.Wordprocessing.TableStyle;
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
                var b = (SectionProperties?) a;
                var xmlMargins = b?.ChildElements.FirstOrDefault(x => x.LocalName == "pgMar");
                var xmlSize = b?.ChildElements.FirstOrDefault(x => x.LocalName == "pgSz");
                var size = (PageSize?) xmlSize;
                var aa = new QuestPDF.Helpers.PageSize(size!.Width! / DocxToQuestPDFScale,
                    size!.Height! / DocxToQuestPDFScale);
                var d = (PageMargin?) xmlMargins;
                page.Background().Background(Colors.White);
                page.Size(aa);
                page.MarginLeft(d?.Left / DocxToQuestPDFScale);
                page.MarginRight(d?.Right / DocxToQuestPDFScale);
                // page.MarginTop(d?.Top/DocxToQuestPDFScale);
                // page.MarginBottom(d?.Bottom/DocxToQuestPDFScale);
                page.Header().Column(ComposeHeader);
                page.Content().Column(ComposeBody);
                page.Footer().Column(ComposeFooter);
            });
        var a = 1;
        Console.WriteLine(a);
    }

    private void ComposeHeader(ColumnDescriptor descriptor)
    {
        descriptor.Item().Height(10);
        var xmlHeader = Document.MainDocumentPart.HeaderParts.FirstOrDefault()?.Header;
        if (xmlHeader is null)
            return;
        ResolveNode(descriptor, xmlHeader.ChildElements);
        descriptor.Item().Height(10);
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

    private void ResolveNode(ColumnDescriptor descriptor, OpenXmlElementList nodes,
        (float? bottom, float? top, float? right, float? left)? padding = null)
    {
        foreach (var node in nodes)
            ResolveNode(descriptor, node, padding);
    }

    private void ResolveNode(ColumnDescriptor descriptor, OpenXmlElement node,
        (float? bottom, float? top, float? right, float? left)? padding = null)
    {
        switch (node.LocalName)
        {
            case "tbl":
            {
                var xmlTable = (Table) node;
                var props = xmlTable.ChildElements.OfType<TableProperties>().FirstOrDefault();
                var aaa = props?.TableCellMarginDefault;
                var xmlTableGrid = node.ChildElements.OfType<TableGrid>().FirstOrDefault();
                var columns = xmlTableGrid!.ChildElements.OfType<GridColumn>();
                var enumerable = columns as GridColumn[] ?? columns.ToArray();
                uint rowC = 1;
                uint columnC = 1;

                descriptor.Item().Container().EnsureSpace().Table(
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
                        uint aaa = 1;
                        uint bbb = 1;
                        var count = tableRows.First().OfType<TableCell>().Count();
                        var ab = new int[tableRows.Count(), count];
                        foreach (var tableRow in tableRows)
                        {
                            var tableRowProperties = tableRow.TableRowProperties;
                            var height = tableRowProperties?.ChildElements.OfType<TableRowHeight>().FirstOrDefault();
                            var tableRowCells = tableRow.ChildElements.OfType<TableCell>();
                            var runProperties = tableRow.Descendants<RunProperties>();
                            var maxFontInTheRow = runProperties.MaxBy(x => float.Parse(x.FontSize.Val.Value))?.FontSize
                                .Val.Value;


                            foreach (var tableCell in tableRowCells)
                            {
                                var gripSpan = (uint) (tableCell.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
                                var columnSpanValue = tableCell.TableCellProperties?.VerticalMerge?.Val?.Value;
                                if (columnSpanValue == MergedCellValues.Restart)
                                {
                                    // var i = ab.TryGetValue(new CellCoordinates());
                                }

                                aaa += gripSpan;
                            }

                            foreach (var tableCell in tableRowCells)
                            {
                                var gripSpan = (uint) (tableCell.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
                                var columnSpanValue = tableCell.TableCellProperties?.VerticalMerge?.Val?.Value;
                                (float bottom, float top, float right, float left) a = (
                                    float.Parse(tableCell.TableCellProperties?.TableCellMargin?.BottomMargin?.Width ??
                                                "0"),
                                    float.Parse(tableCell.TableCellProperties?.TableCellMargin?.TopMargin?.Width ??
                                                "0"),
                                    float.Parse(tableCell.TableCellProperties?.TableCellMargin?.RightMargin?.Width ??
                                                "0"),
                                    float.Parse(
                                        tableCell.TableCellProperties?.TableCellMargin?.LeftMargin?.Width ?? "0"));
                                table.Cell().Row(rowC).Column(columnC).ColumnSpan(gripSpan).Border(0.1f)
                                    .MinHeight(height?.Val is null
                                        ? (maxFontInTheRow is null
                                            ? 0
                                            : float.Parse(maxFontInTheRow) / DocxToQuestPDFScale)
                                        : height.Val / DocxToQuestPDFScale).Column(x =>
                                    {
                                        ResolveNode(x, tableCell.ChildElements, a);
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
                var xmlParagraph = (Paragraph) node;
                descriptor.Item()
                    .PaddingLeft(padding?.left ?? 0)
                    .PaddingBottom(padding?.bottom ?? 0)
                    .PaddingTop(padding?.top ?? 0)
                    .PaddingRight(padding?.bottom ?? 0).Text(
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
                var xmlParagraphProperties = (ParagraphProperties) node;
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
                var xmlRun = (Run) node;
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
                    var drawing = (Drawing) xmlRun.ChildElements.First(x => x.LocalName == "drawing");
                    var picture = drawing.Inline.Graphic.GraphicData
                        .Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().FirstOrDefault();
                    var blip = picture?.BlipFill.Blip?.Embed?.Value;
                    if (blip is null)
                        break;
                    var img = (ImagePart) Document.MainDocumentPart.GetPartById(blip);
                    using var imageStream = img.GetStream();
                    text.Element().AlignMiddle().Image(imageStream);
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
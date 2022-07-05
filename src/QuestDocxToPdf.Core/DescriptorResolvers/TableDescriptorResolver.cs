using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using QuestPDF.Fluent;

namespace QuestDocxToPdf.Core.DescriptorResolvers;

public class TableDescriptorResolver
{
    private TableDescriptor _descriptor;
    private Table _node;
    private IDictionary<CellCoordinates, int> _verticalSpan = new Dictionary<CellCoordinates, int>();

    public TableDescriptorResolver(TableDescriptor descriptor, Table node)
    {
        _descriptor = descriptor;
        _node = node;
    }

    public void Resolve()
    {
        
    }

    private void CreateTableRowColumnMap()
    {
        var rows = _node.OfType<TableRow>();
        foreach (var row in rows)
        {
            
        }
    }
}
namespace QuestDocxToPdf.Core;

public class DocXGenerationOptions
{
    private DocXGenerationOptions? _instance;

    public DocXGenerationOptions GetInstance()
        => _instance ??= new DocXGenerationOptions();

    public float DefaultCellHorizontalMargins { get; set; } = 2;
}
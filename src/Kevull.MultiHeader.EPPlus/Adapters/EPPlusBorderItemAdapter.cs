using System.Drawing;
using Kevull.MultiHeader.Core.Abstractions;
using OfficeOpenXml.Style;

namespace Kevull.MultiHeader.EPPlus.Adapters;

internal class EPPlusBorderItemAdapter : IExcelBorderItem
{
    private readonly ExcelBorderItem _borderItem;

    public EPPlusBorderItemAdapter(ExcelBorderItem borderItem)
    {
        _borderItem = borderItem;
    }

    public object Style
    {
        get => _borderItem.Style;
        set => _borderItem.Style = (ExcelBorderStyle)value;
    }

    public Color Color
    {
        get => _borderItem.Color.ToColor();
        set => _borderItem.Color.SetColor(value);
    }
}

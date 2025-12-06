using Kevull.MultiHeader.Core.Abstractions;
using OfficeOpenXml.Style;

namespace Kevull.MultiHeader.EPPlus.Adapters;

internal class EPPlusBorderAdapter : IExcelBorder
{
    private readonly Border _border;

    public EPPlusBorderAdapter(Border border)
    {
        _border = border;
    }

    public IExcelBorderItem Top => new EPPlusBorderItemAdapter(_border.Top);
    public IExcelBorderItem Bottom => new EPPlusBorderItemAdapter(_border.Bottom);
    public IExcelBorderItem Left => new EPPlusBorderItemAdapter(_border.Left);
    public IExcelBorderItem Right => new EPPlusBorderItemAdapter(_border.Right);
}

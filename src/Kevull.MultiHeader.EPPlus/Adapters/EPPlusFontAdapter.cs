using System.Drawing;
using Kevull.MultiHeader.Core.Abstractions;
using OfficeOpenXml.Style;

namespace Kevull.MultiHeader.EPPlus.Adapters;

internal class EPPlusFontAdapter : IExcelFont
{
    private readonly ExcelFont _font;

    public EPPlusFontAdapter(ExcelFont font)
    {
        _font = font;
    }

    public string Name
    {
        get => _font.Name;
        set => _font.Name = value;
    }

    public float Size
    {
        get => _font.Size;
        set => _font.Size = value;
    }

    public bool Bold
    {
        get => _font.Bold;
        set => _font.Bold = value;
    }

    public bool Italic
    {
        get => _font.Italic;
        set => _font.Italic = value;
    }

    public bool Strike
    {
        get => _font.Strike;
        set => _font.Strike = value;
    }

    public bool UnderLine
    {
        get => _font.UnderLine;
        set => _font.UnderLine = value;
    }

    public Color Color
    {
        get => Color.Empty.FromHexString(_font.Color.Rgb);
        set => _font.Color.SetColor(value);
    }
}

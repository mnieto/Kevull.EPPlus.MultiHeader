using System.Drawing;

namespace Kevull.MultiHeader.Core.Abstractions;

public interface IExcelFont
{
    string Name { get; set; }
    float Size { get; set; }
    bool Bold { get; set; }
    bool Italic { get; set; }
    bool Strike { get; set; }
    bool UnderLine { get; set; }
    Color Color { get; set; }
}

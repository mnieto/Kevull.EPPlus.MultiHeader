using System.Drawing;

namespace Kevull.MultiHeader.Core.Abstractions;

public interface IExcelFill
{
    object PatternType { get; set; }
    Color BackgroundColor { get; set; }
    Color PatternColor { get; set; }
}

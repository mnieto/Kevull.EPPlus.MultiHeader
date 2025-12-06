using System.Drawing;

namespace Kevull.MultiHeader.Core.Abstractions;

public interface IExcelBorderItem
{
    object Style { get; set; }
    Color Color { get; set; }
}

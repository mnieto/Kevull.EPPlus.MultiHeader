namespace Kevull.MultiHeader.Core.Abstractions;

public interface IExcelBorder
{
    IExcelBorderItem Top { get; }
    IExcelBorderItem Bottom { get; }
    IExcelBorderItem Left { get; }
    IExcelBorderItem Right { get; }
}

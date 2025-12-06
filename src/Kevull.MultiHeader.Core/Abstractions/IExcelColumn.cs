namespace Kevull.MultiHeader.Core.Abstractions;

public interface IExcelColumn
{
    public bool Hidden { get; set; }
    void AutoFit(double MinimumWidth, double MaximumWidth);
    double Width { get; set; }
}

namespace Kevull.MultiHeader.Core.Abstractions;
public interface IExcelWorkbook
{
    IExcelWorksheet AddWorksheet(string name);
    void Calculate();
    void SaveAs(string filePath);
}
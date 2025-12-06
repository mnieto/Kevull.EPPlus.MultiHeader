using Kevull.MultiHeader.Core;
using Kevull.MultiHeader.Core.Abstractions;
using Kevull.MultiHeader.EPPlus.Adapters;
using OfficeOpenXml;

namespace Kevull.MultiHeader.EPPlus;

/// <summary>
/// EPPlus-specific implementation of ConfigurationBuilder
/// </summary>
/// <typeparam name="T"></typeparam>
public class EPPlusConfigurationBuilder<T> : ConfigurationBuilder<T>
{
    private readonly ExcelPackage _xls;

    public EPPlusConfigurationBuilder(ExcelPackage xls)
    {
        _xls = xls;
    }

    public override ConfigurationBuilder<T> AddNamedStyle(string name, Action<IExcelStyle> style)
    {
        var namedStyle = _xls.Workbook.Styles.CreateNamedStyle(name);
        var adapter = new EPPlusStyleAdapter(namedStyle.Style);
        style?.Invoke(adapter);
        return this;
    }

    public override ConfigurationBuilder<T> SetStartingAddress(string address)
    {
        var cellAddress = new ExcelCellAddress(address);
        return SetStartingAddres(cellAddress.Row, cellAddress.Column);
    }

    public override ConfigurationBuilder<T> SetStartingAddres(int row, int column)
    {
        var cellAddress = new EPPlusCellAddress(row, column);
        SetStartingAddressInternal(cellAddress);
        return this;
    }

    private class EPPlusCellAddress : IExcelCellAddress
    {
        public EPPlusCellAddress(int row, int column)
        {
            Row = row;
            Column = column;
        }

        public int Row { get; }
        public int Column { get; }
    }
}

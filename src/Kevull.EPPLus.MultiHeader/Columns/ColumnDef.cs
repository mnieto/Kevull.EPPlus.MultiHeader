using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kevull.EPPLus.MultiHeader.Columns
{
    /// <summary>
    /// Column common properties
    /// </summary>
    public class ColumnDef
    {
        internal ColumnDef() { }

        private int? _order;

        /// <summary>
        /// Human friendly name. If it is not provided, it will use <see cref="ColumnInfo.Name"/>
        /// </summary>
        public string? DisplayName { get; set; }

        /// <summary>
        /// Is this column rendered but hidden?
        /// </summary>
        public bool? Hidden { get; set; }

        /// <summary>
        /// Diplay order. Order is relative to the other columns. Columns that has no order are added after those that have it. Order starts from 1
        /// </summary>
        public int? Order
        {
            get => _order;
            set
            {
                if (value != null && value <= 0)
                    throw new ArgumentOutOfRangeException(nameof(Order), "Value must be null or be greater or equals to 1");
                _order = value;
            }

        }

        /// <summary>
        /// Name of a style defined in the Excel workbook
        /// </summary>
        /// <remarks>
        /// Style names are not checked at configuration time, but in the <see cref="MultiHeaderReport{T}.GenerateReport(IEnumerable{T})"/> method
        /// You can assign the style name during the column creation or use any existing Style in the Excel file. 
        /// The <see cref="ConfigurationBuilder{T}.AddNamedStyle(string, Action{OfficeOpenXml.Style.ExcelStyle})"/> is a handy method that wraps the EPPlus <see cref="OfficeOpenXml.ExcelStyles.CreateNamedStyle(string)"/> method
        /// </remarks>
        public string? StyleName { get; set; }

        /// <summary>
        /// Allows to configure the column's width behaviour: Default, Custom, Auto, Hidden
        /// </summary>
        public ColumnWidth ColumnWidth { get; private set; } = new ColumnWidth();

    }    
}

using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EPPLus.MultiHeader.Columns
{
    /// <summary>
    /// Add a formula column. That is, each time the report will render a value for this column, it will use the specified Excel formula
    /// </summary>
    public class ColumnFormula : ColumnInfo
    {
        private readonly string _formula;

        public override bool IsDynamic => true;

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="name">name of the property. In this case, it cannot be infered from the source Type</param>
        /// <param name="formula">Excel Formula used for this column. Be sure to use the correct absulte/relative references in the formula</param>
        public ColumnFormula(string name, string formula) : base(name)
        {
            _formula = formula;
        }

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="name">name of the property. In this case, it cannot be infered from the source Type</param>
        /// <param name="formula">Excel Formula used for this column. Be sure to use the correct absulte/relative references in the formula</param>
        /// <param name="order">Diplay order. Order is relative to the other columns. Columns that has no <paramref name="order"/> are added after those that have it</param>
        /// <param name="displayName">Human friendly name for the column. If not specified, the property Name is used</param>
        /// <param name="hidden">Column is written to the Excel, but it's hidden</param>
        public ColumnFormula(string name, string formula, int? order=null, string? displayName = null, bool hidden = false) : base(name, order, displayName, hidden)
        {
            _formula = formula;
        }

        internal override void WriteCell(ExcelRange cell, Dictionary<string, PropertyInfo> properties, object? obj)
        {
            cell.Formula = _formula;
        }
    }


}

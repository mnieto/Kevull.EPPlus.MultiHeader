using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EPPLus.MultiHeader.Columns
{
    public class ColumnFormula : ColumnInfo
    {
        private readonly string _formula;

        public override bool IsDynamic => true;
        public ColumnFormula(string name, string formula) : base(name)
        {
            _formula = formula;
        }

        public ColumnFormula(string name, string formula, int? order=null, string? displayName = null, bool hidden = false) : base(name, order, displayName, hidden)
        {
            _formula = formula;
        }

        public override void WriteCell(ExcelRange cell, Dictionary<string, PropertyInfo> properties, object? obj)
        {
            cell.Formula = _formula;
        }
    }


}

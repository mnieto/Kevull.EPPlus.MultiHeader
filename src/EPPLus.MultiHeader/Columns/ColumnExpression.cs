using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EPPLus.MultiHeader.Columns
{
    public class ColumnExpression<T> : ColumnInfo
    {
        private Func<T, object?> _expression;
        public override bool IsDynamic => true;
        public ColumnExpression(string name, Func<T, object?> expression) : base(name)
        {
            _expression = expression ?? throw new ArgumentNullException(nameof(expression));
        }

        public ColumnExpression(string name, Func<T, object?> expression, int? order = null, string? displayName = null, bool hidden = false) : base(name, order, displayName, hidden)
        {
            _expression = expression ?? throw new ArgumentNullException(nameof(expression));
        }

        public override void WriteCell(ExcelRange cell, Dictionary<string, PropertyInfo> properties, object? obj)
        {
            if (obj is null)
                return;

            cell.Value = _expression((T)obj);
        }
    }

}

using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EPPLus.MultiHeader.Columns
{
    public class ColumnInfo
    {
        protected string? _displayName;
        protected int? _order;
        public bool Hidden { get; set; }
        public string Name { get; set; }
        public string DisplayName { get => _displayName ?? Name; set => _displayName = value; }

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

        public bool Ignore { get; set; }

        public ColumnInfo(string name, bool ignore)
        {
            Name = name;
            Ignore = ignore;
        }
        public ColumnInfo(string name, int? order = null, string? displayName = null, bool hidden = false)
        {
            Hidden = hidden;
            Name = name;
            Order = order;
            _displayName = displayName;
        }

        public virtual void WriteCell(ExcelRange cell, Dictionary<string, PropertyInfo> properties, object obj)
        {
            cell.Value = properties[Name].GetValue(obj);
        }

    }

    public class ColumnInfo<T> : ColumnInfo
    {
        public ColumnInfo(Expression<Func<T, object?>> columnSelector) : base(GetPropertyName(columnSelector)) { }
        public ColumnInfo(Expression<Func<T, object?>> columnSelector, bool ignore) : base(GetPropertyName(columnSelector), ignore) { }
        public ColumnInfo(Expression<Func<T, object?>> columnSelector, int? order = null, string? displayName = null, bool hidden = false)
            : base(GetPropertyName(columnSelector), order, displayName, hidden) { }

        internal static string GetPropertyName(Expression<Func<T, object?>> columnSelector)
        {
            var memberExpr = columnSelector.Body as MemberExpression;
            var unaryExpr = columnSelector.Body as UnaryExpression;
            if (memberExpr == null && unaryExpr == null)
                throw new InvalidCastException(columnSelector.Body.ToString());

            return (memberExpr ?? (unaryExpr!.Operand as MemberExpression)!).Member.Name;
        }
    }
}

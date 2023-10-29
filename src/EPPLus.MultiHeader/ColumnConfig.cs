using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EPPLus.MultiHeader
{
    public class ColumnConfig
    {
        private string? _displayName;
        public bool Hidden { get; set; }
        public string Name { get; set; }
        public int? Order { get; set; }
        public bool Ignore { get; set; }
        public string DisplayName { get => _displayName ?? Name; set => _displayName = value; }

        public ColumnConfig(string name, bool ignore)
        {
            Name = name;
            Ignore = ignore;
        }
        public ColumnConfig(string name, int? order = null, string? displayName = null, bool hidden = false)
        {
            Hidden = hidden;
            Name = name;
            Order = order;
            _displayName = displayName;
        }

    }

    public class ColumnConfig<T> : ColumnConfig
    {
        public ColumnConfig(Expression<Func<T, object>> columnSelector) : base(GetPropertyName(columnSelector)) { }
        public ColumnConfig(Expression<Func<T, object>> columnSelector, bool ignore) : base(GetPropertyName(columnSelector), ignore) { }
        public ColumnConfig(Expression<Func<T, object>> columnSelector, int? order = null, string? displayName = null, bool hidden = false)
            : base(GetPropertyName(columnSelector), order, displayName, hidden) { }

        private static string GetPropertyName(Expression<Func<T, object>> columnSelector)
        {
            var memberExpr = columnSelector.Body as MemberExpression;
            var unaryExpr = columnSelector.Body as UnaryExpression;
            if (memberExpr == null && unaryExpr == null)
                throw new InvalidCastException(columnSelector.Body.ToString());

            return (memberExpr ?? (unaryExpr!.Operand as MemberExpression)!).Member.Name;
        }
    }
}

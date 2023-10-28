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
        public string Name { get; set; }
        public int? Order { get; set; }
        public bool Ignore { get; set; }
        public string DisplayName { get => _displayName ?? Name; set => _displayName = value; }

        public ColumnConfig(string name) : this(name, false) { }
        public ColumnConfig(string name, bool ignore)
        {
            Name = name;
            Ignore = ignore;
        }
        public ColumnConfig(string name, int order, string? displayName = null)
        {
            Name = name;
            Order = order;
            _displayName = displayName;
        }

    }

    public class ColumnConfig<T> : ColumnConfig
    {
        public ColumnConfig(Expression<Func<T, object>> columnSelector) : base(GetPropertyName(columnSelector)) { }
        public ColumnConfig(Expression<Func<T, object>> columnSelector, bool ignore) : base(GetPropertyName(columnSelector), ignore) { }
        public ColumnConfig(Expression<Func<T, object>> columnSelector, int order, string? displayName = null)
            : base(GetPropertyName(columnSelector), order, displayName) { }

        private static string GetPropertyName(Expression<Func<T, object>> columnSelector)
        {
            if (columnSelector.Body is MemberExpression memberExpr)
            {
                return memberExpr.Member.Name;
            }
            else if (columnSelector.Body is UnaryExpression body)
            {
                body = (UnaryExpression)columnSelector.Body;
                return ((MemberExpression)body.Operand).Member.Name;
            }
            else
            {
                throw new InvalidCastException(columnSelector.Body.ToString());
            }
        }
    }
}

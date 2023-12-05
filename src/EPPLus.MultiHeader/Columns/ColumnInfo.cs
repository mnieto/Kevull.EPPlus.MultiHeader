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
        public string FullName { get; set; }
        public string? ParentName { get; protected set; }
        public Type? ParentType { get; protected set; }
        public string DisplayName { get => _displayName ?? Name; set => _displayName = value; }

        public int Index { get; set; } 
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

        public virtual bool IsDynamic => false;
        public int Deep => FullName.Split('.').Length;
        public virtual bool IsMultiValue => false;
        public HeaderManager? Header { get; set; }
        public bool HasChildren => Header != null && Header.Columns.Count > 0;
        public virtual int Width => Header == null ? 1 : Header!.Columns.Sum(c  => c.Width);


        public ColumnInfo(string name, bool ignore)
        {
            FullName = name;
            Name = GetName(name);
            Ignore = ignore;
        }
        public ColumnInfo(string name, int? order = null, string? displayName = null, bool hidden = false)
        {
            Hidden = hidden;
            FullName = name;
            Name = GetName(name);
            Order = order;
            _displayName = displayName;
        }

        internal ColumnInfo(PropertyNames names, bool ignore)
        {
            FullName = names.FullName;
            Name = names.Name;
            ParentName = names.ParentName;
            ParentType = names.ParentType;
            Ignore = ignore;
        }

        internal ColumnInfo(PropertyNames names, int? order = null, string? displayName = null, bool hidden = false)
        {
            Hidden = hidden;
            FullName = names.FullName;
            Name = names.Name;
            ParentName = names.ParentName;
            ParentType = names.ParentType;
            Order = order;
            _displayName = displayName;
        }

        public virtual void WriteCell(ExcelRange cell, Dictionary<string, PropertyInfo> properties, object? obj)
        {
            if (obj != null)
                cell.Value = properties[Name].GetValue(obj);
        }

        public virtual void WriteHeader(ExcelRange cell)
        {
            cell.Value = DisplayName;
        }

        private string GetName(string fullName)
        {
            int pos = fullName.IndexOf('.');
            return pos == -1 ? fullName : fullName.Substring(0, pos);
        }

    }

    public class ColumnInfo<T> : ColumnInfo
    {
        public ColumnInfo(Expression<Func<T, object?>> columnSelector) : base(GetPropertyName(columnSelector)) { }
        public ColumnInfo(Expression<Func<T, object?>> columnSelector, bool ignore) : base(GetPropertyName(columnSelector), ignore) { }
        public ColumnInfo(Expression<Func<T, object?>> columnSelector, int? order = null, string? displayName = null, bool hidden = false)
            : base(GetPropertyName(columnSelector), order, displayName, hidden) { }

        internal static PropertyNames GetPropertyName(Expression<Func<T, object?>> columnSelector)
        {
            return new PropertyNameBuilder<T>().Build(columnSelector);
        }
    }
}

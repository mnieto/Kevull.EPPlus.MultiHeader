using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Kevull.EPPLus.MultiHeader.Columns
{
    /// <summary>
    /// Base class for columns
    /// </summary>
    [DebuggerDisplay("{Name}")]
    public class ColumnInfo
    {
        /// <summary>
        /// Human friendly name. If it is not provided, it will use <see cref="Name"/>
        /// </summary>
        protected string? _displayName;

        /// <summary>
        /// Diplay order. Order is relative to the other columns. Columns that has no order are added after those that have it. Order starts from 1
        /// </summary>
        protected int? _order;

        /// <summary>
        /// Is this column rendered but hidden?
        /// </summary>
        public bool Hidden { get; set; }

        /// <summary>
        /// Column name. This will match with the property name, except those columns that are Dynamic (<see cref="ColumnExpression{T}"/> and <see cref="ColumnFormula"/>).
        /// </summary>
        public string Name { get; set; }
        
        /// <summary>
        /// Full name for nested properties. For first level properties, <see cref="FullName"/> and <see cref="Name"/> will be the same
        /// </summary>
        internal string FullName { get; set; }

        /// <summary>
        /// Parent property name. If the column has a parent property
        /// </summary>
        public string? ParentName { get; protected set; }
        
        /// <summary>
        /// Parent property Type. If the column has a parent property
        /// </summary>
        protected internal Type? ParentType { get; set; }

        /// <summary>
        /// Allows to configure the colum with and get the with properties
        /// </summary>
        public ColumnWidth ColumnWidth { get; set; } = new ColumnWidth();

        /// <summary>
        /// Human friendly name. If it is not provided, it will use <see cref="Name"/>
        /// </summary>
        public string DisplayName { get => _displayName ?? Name; set => _displayName = value; }

        /// <summary>
        /// Excel column index where render the data. Do not confuse with <see cref="Order"/>. Intended for internal use purposes.
        /// </summary>
        internal int Index { get; set; }

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
        /// Ignore this property. This column will not be rendered
        /// </summary>
        public bool Ignore { get; set; }

        /// <summary>
        /// Data content is rendered from the source object or calculated
        /// </summary>
        internal virtual bool IsDynamic => false;
        
        /// <summary>
        /// Number of child levels below this
        /// </summary>
        internal int Deep => FullName.Split('.').Length;
        
        /// <summary>
        /// Is it a property with a single value or is it a <see cref="IDictionary{TKey, TValue}"/> or <see cref="IEnumerable{T}"/>.
        /// </summary>
        internal virtual bool IsMultiValue => false;

        /// <summary>
        /// If this column's Type is a complex object, this property will store the child headers
        /// </summary>
        internal HeaderManager? Header { get; set; }
        
        /// <summary>
        /// Has child columns. That is, is it a complex object?
        /// </summary>
        internal bool HasChildren => Header != null && Header.Columns.Count > 0;
        
        /// <summary>
        /// Number of Excel columns needed to render this property (and all its children)
        /// </summary>
        internal virtual int Width => Header == null ? 1 : Header!.Columns.Sum(c  => c.Width);

        /// <summary>
        /// Name of a style defined in the Excel workbook
        /// </summary>
        /// <remarks>
        /// Style names are not checked at configuration time, but in the <see cref="MultiHeaderReport{T}.GenerateReport(IEnumerable{T})"/> method
        /// You can assign the style name during the column creation or use any existing Style in the Excel file. 
        /// The <see cref="ConfigurationBuilder{T}.AddNamedStyle(string, Action{OfficeOpenXml.Style.ExcelStyle})"/> is a handy method
        /// that wraps the EPPlus <see cref="ExcelStyles.CreateNamedStyle(string)"/> method
        /// </remarks>
        public string? StyleName { get; set; }

        /// <summary>
        /// Ctor. Used internally in nested properties and for testing purposes. Use <see cref="ColumnInfo{T}"/>
        /// </summary>
        internal ColumnInfo(string name, bool ignore)
        {
            FullName = name;
            Name = GetName(name);
            Ignore = ignore;
        }

        /// <summary>
        /// Ctor. Used internally in nested properties and for testing purposes. Use <see cref="ColumnInfo{T}"/>
        /// </summary>
        internal ColumnInfo(string name, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null)
        {
            Hidden = hidden;
            FullName = name;
            Name = GetName(name);
            Order = order;
            _displayName = displayName;
            StyleName = styleName;
        }

        internal ColumnInfo(string name, Action<ColumnDef> configAction)
        {
            var config = new ColumnDef();
            configAction.Invoke(config);
            FullName = name;
            Name = GetName(name);
        }

        /// <summary>
        /// Ctor. Used internally in nested properties and for testing purposes. Use <see cref="ColumnInfo{T}"/>
        /// </summary>
        internal ColumnInfo(PropertyNames names, Action<ColumnDef> configAction)
        {
            var config = new ColumnDef();
            configAction.Invoke(config);

            FullName = names.FullName;
            Name = names.Name;
            ParentName = names.ParentName;
            ParentType = names.ParentType;

            //Hidden = config.Hidden;
            Order = config.Order;
            _displayName = config.DisplayName;
            StyleName = config.StyleName;
            ColumnWidth = config.ColumnWidth;
        }

        /// <summary>
        /// Ctor. Used internally in nested properties and for testing purposes. Use <see cref="ColumnInfo{T}"/>
        /// </summary>
        internal ColumnInfo(PropertyNames names, bool ignore)
        {
            FullName = names.FullName;
            Name = names.Name;
            ParentName = names.ParentName;
            ParentType = names.ParentType;
            Ignore = ignore;
        }

        /// <summary>
        /// Ctor. Used internally in nested properties and for testing purposes. Use <see cref="ColumnInfo{T}"/>
        /// </summary>
        internal ColumnInfo(PropertyNames names, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null)
        {
            Hidden = hidden;
            FullName = names.FullName;
            Name = names.Name;
            ParentName = names.ParentName;
            ParentType = names.ParentType;
            Order = order;
            StyleName= styleName;
            _displayName = displayName;
        }

        internal virtual void FormatHeader(ExcelRange cell, int height)
        {
            cell.Offset(0, 0, height, Width).Merge = true;
        }

        internal virtual void WriteCell(ExcelRange cell, Dictionary<string, PropertyInfo> properties, object? obj)
        {
            if (obj != null)
                cell.Value = properties[Name].GetValue(obj);
        }

        internal virtual void WriteHeader(ExcelRange cell)
        {
            cell.Value = DisplayName;
        }

        private string GetName(string fullName)
        {
            int pos = fullName.IndexOf('.');
            return pos == -1 ? fullName : fullName.Substring(0, pos);
        }

    }

    /// <summary>
    /// Base class for columns
    /// </summary>
    public class ColumnInfo<T> : ColumnInfo
    {
        /// <summary>
        /// Simple Ctor
        /// </summary>
        /// <param name="columnSelector">Lambda expression to specify the property</param>
        public ColumnInfo(Expression<Func<T, object?>> columnSelector) : base(GetPropertyName(columnSelector)) { }

        /// <summary>
        /// Ctor for ignore use case
        /// </summary>
        /// <param name="columnSelector">Lambda expression to specify the property</param>
        /// <param name="ignore">Ignore this property. This column will not be rendered</param>
        public ColumnInfo(Expression<Func<T, object?>> columnSelector, bool ignore) : base(GetPropertyName(columnSelector), ignore) { }

        /// <summary>
        /// General use Ctor
        /// </summary>
        /// <param name="columnSelector">Lambda expression to specify the property</param>
        /// <param name="order">Ignore this property. This column will not be rendered</param>
        /// <param name="displayName">Human friendly name. If it is not provided, it will use <see cref="ColumnInfo.Name"/></param>
        /// <param name="hidden">Is this column rendered but hidden?</param>
        /// <param name="styleName">Name of a style defined in the Excel workbook</param>
        public ColumnInfo(Expression<Func<T, object?>> columnSelector, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null)
            : base(GetPropertyName(columnSelector), order, displayName, hidden, styleName) { }

        /// <summary>
        /// General use Ctor
        /// </summary>
        /// <param name="columnSelector">Lambda expression to specify the property</param>
        /// <param name="cfg"> Action that will be invoked to configure the ColumnInfo properties using a <see cref="ColumnDef"/> object</param>
        public ColumnInfo(Expression<Func<T, object?>> columnSelector, Action<ColumnDef> cfg) : base(GetPropertyName(columnSelector), cfg) { }

        /// <summary>
        /// General use Ctor
        /// </summary>
        /// <param name="name">name for this column</param>
        /// <param name="cfg"> Action that will be invoked to configure the ColumnInfo properties using a <see cref="ColumnDef"/> object</param>
        public ColumnInfo(string name, Action<ColumnDef> cfg) : base(name, cfg) { }

        /// <summary>
        /// Ctor. Used internally in nested properties and for testing purposes. Use <see cref="ColumnInfo{T}"/>
        /// </summary>
        internal ColumnInfo(string name, bool ignore) : base(name, ignore) { }

        /// <summary>
        /// Ctor. Used internally in nested properties and for testing purposes. Use <see cref="ColumnInfo{T}"/>
        /// </summary>
        internal ColumnInfo(string name, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null)
            : base(name, order, displayName, hidden, styleName) { }

        /// <summary>
        /// Ctor. Used internally in nested properties and for testing purposes. Use <see cref="ColumnInfo{T}"/>
        /// </summary>
        internal ColumnInfo(PropertyNames names, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null)
            : base(names, order, displayName, hidden, styleName) { }

        /// <summary>
        /// Ctor. Used internally in nested properties and for testing purposes. Use <see cref="ColumnInfo{T}"/>
        /// </summary>
        internal ColumnInfo(PropertyNames names, bool ignore) : base(names, ignore) { }

        internal static PropertyNames GetPropertyName(Expression<Func<T, object?>> columnSelector)
        {
            return new PropertyNameBuilder<T>().Build(columnSelector);
        }
    }
}

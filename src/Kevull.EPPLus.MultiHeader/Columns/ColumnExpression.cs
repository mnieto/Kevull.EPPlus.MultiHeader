using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Kevull.EPPLus.MultiHeader.Columns
{
    /// <summary>
    /// Add an expression column. That is, each time the report will render a value for this column, it will invoke a lambda expression.
    /// </summary>
    public class ColumnExpression<T> : ColumnInfo<T>
    {
        private Func<T, object?> _expression;

        /// <summary>
        /// Data content is rendered from the source object or calculated
        /// </summary>
        public override bool IsDynamic => true;

        /// <summary>
        /// Ctor. Used ineternaly in nested properties and for testing purposes. Use <see cref="ColumnExpression{T}"/>
        /// </summary>
        /// <param name="name">name of the property. In this case, it cannot be infered from the source Type</param>
        /// <param name="expression">Lambda expression to be evaluated to render the column value each row</param>
        internal ColumnExpression(string name, Func<T, object?> expression) : base(name)
        {
            _expression = expression ?? throw new ArgumentNullException(nameof(expression));
        }

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="name">name of the property. In this case, it cannot be infered from the source Type</param>
        /// <param name="expression">Lambda expression to be evaluated to render the column value each row</param>
        /// <param name="order">Diplay order. Order is relative to the other columns. Columns that has no <paramref name="order"/> are added after those that have it</param>
        /// <param name="displayName">Human friendly name for the column. If not specified, the property Name is used</param>
        /// <param name="hidden">Column is written to the Excel, but it's hidden</param>
        public ColumnExpression(string name, Func<T, object?> expression, int? order = null, string? displayName = null, bool hidden = false) : base(name, order, displayName, hidden)
        {
            _expression = expression ?? throw new ArgumentNullException(nameof(expression));
        }

        internal override void WriteCell(ExcelRange cell, Dictionary<string, PropertyInfo> properties, object? obj)
        {
            if (obj is null)
                return;

            cell.Value = _expression((T)obj);
        }
    }

}

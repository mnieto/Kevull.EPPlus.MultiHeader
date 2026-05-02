using Kevull.MultiHeader.Core;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Kevull.MultiHeader.Core.Columns
{
    /// <summary>
    /// Specialized <see cref="ColumnInfo"/> that renders data from a <see cref="IDictionary{TKey, TValue}"/> or <see cref="IEnumerable{T}"/>.
    /// </summary>
    public class ColumnEnumeration<T> : ColumnInfo<T>
    {
        private readonly Dictionary<string, int> _keyValues;

        /// <summary>
        /// Number of Excel columns needed to render this property (and all its children)
        /// </summary>
        internal override int Width => Header == null ? _keyValues.Count : Header!.Columns.Sum(c => c.Width) * _keyValues.Count;

        /// <summary>
        /// Is it a property with a single value or is it a <see cref="IDictionary{TKey, TValue}"/> or <see cref="IEnumerable{T}"/>.
        /// </summary>
        internal override bool IsMultiValue => true;

        /// <summary>
        /// Gets the allowed values for the child columns
        /// </summary>
        public List<string> Keys => _keyValues.Keys.Cast<string>().ToList();

        /// <summary>
        /// General use Ctor
        /// </summary>
        /// <param name="columnSelector">Lambda expression to specify the property</param>
        /// <param name="keyValues">Allowed key values. This is used to allocate a specific number of columns</param>
        /// <param name="order">Ignore this property. This column will not be rendered</param>
        /// <param name="displayName">Human friendly name. If it is not provided, it will use <see cref="ColumnInfo.Name"/></param>
        /// <param name="hidden">Is this column rendered but hidden?</param>
        /// <param name="styleName">Name of a style defined in the Excel workbook</param>
        public ColumnEnumeration(Expression<Func<T, object?>> columnSelector, IEnumerable<string> keyValues, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null) :
            base(ColumnInfo<T>.GetPropertyName(columnSelector), order, displayName, hidden, styleName)
        {
            _keyValues = AddKeyValues(keyValues);
        }

        /// <summary>
        /// Simple Ctor
        /// </summary>
        /// <param name="columnSelector">Lambda expression to specify the property</param>
        /// <param name="keyValues">Allowed key values. This is used to allocate a specific number of columns</param>
        /// <param name="cfg"> Action that will be invoked to configure the ColumnInfo properties using a <see cref="ColumnDef"/> object</param>
        public ColumnEnumeration(Expression<Func<T, object?>> columnSelector, IEnumerable<string> keyValues, Action<ColumnDef> cfg) :
            base(columnSelector, cfg)
        {
            _keyValues = AddKeyValues(keyValues);
        }

        /// <summary>
        /// Create a Column based on a <see cref="IDictionary{TKey, TValue}"/> or <see cref="IEnumerable{T}"/>.
        /// </summary>
        /// <param name="name">Name of the column</param>
        /// <param name="keyValues">Allowed column names for the child columns</param>
        /// <param name="order">In which position show the column</param>
        /// <param name="displayName">A column display name. If null, <paramref name="name"/> will be used</param>
        /// <param name="hidden">Hide this column</param>
        /// <param name="styleName">Name of a style defined in the Excel workbook</param>
        internal ColumnEnumeration(string name, IEnumerable<string> keyValues, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null) 
            : base(name, order, displayName, hidden, styleName)
        {
            _keyValues = AddKeyValues(keyValues);
        }

        public override void FormatHeader(IExcelWriter writer, int row, int col, int height)
        {
            // Merge the parent header across all columns
            writer.Merge(row, col, row, col + Width - 1);

            // Merge each child column header vertically
            foreach (var kvp in _keyValues)
            {
                int offset = kvp.Value;
                writer.Merge(row + 1, col + offset, row + height - 1, col + offset);
            }
        }

        public override void WriteCell(IExcelWriter writer, int row, int col, Dictionary<string, PropertyInfo> properties, object? obj)
        {
            if (obj == null)
                return;

            var collection = properties[Name].GetValue(obj)!;
            if (collection is IDictionary dictionary)
            {
                var enumerator = dictionary.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    string key = enumerator.Key.ToString()!;
                    int offset = _keyValues[key];   //this will throw if key is not in the initialized keyValues. This is intentional
                    writer.WriteCell(row, col + offset, enumerator.Value);
                }
            }
            else if (collection is IEnumerable enumerable)
            {
                foreach (object item in enumerable)
                {
                    string key = item.ToString()!;
                    int offset = _keyValues[key];   //this will throw if key is not in the initialized keyValues. This is intentional
                    writer.WriteCell(row, col + offset, item);
                }
            }
            else
            {
                throw new NotSupportedException($"only {nameof(IEnumerable)} or {nameof(IDictionary)} are supported");
            }
        }

        public override void WriteHeader(IExcelWriter writer, int row, int col)
        {
            // Write parent header
            writer.WriteCell(row, col, DisplayName);

            // Write child headers
            foreach (var kvp in _keyValues)
            {
                string key = kvp.Key;
                int offset = kvp.Value;
                writer.WriteCell(row + 1, col + offset, key);
            }
        }

        private Dictionary<string, int> AddKeyValues(IEnumerable<string> keyValues)
        {
            int i = 0;
            return keyValues.ToDictionary(x => x, _ => i++);
        }
    }
}

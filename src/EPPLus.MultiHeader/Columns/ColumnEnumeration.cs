using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EPPLus.MultiHeader.Columns
{
    public class ColumnEnumeration : ColumnInfo
    {
        private readonly Dictionary<string, int> _keyValues;

        public override int Width => Header == null ? _keyValues.Count : Header!.Columns.Sum(c => c.Width) * _keyValues.Count;
        public override bool IsMultiValue => true;
        public List<string> Keys => _keyValues.Keys.Cast<string>().ToList();

        public ColumnEnumeration(string name, IEnumerable<string> keyValues, bool ignore) : base(name, ignore)
        {
            int i = 0;
            _keyValues = keyValues.ToDictionary(x => x, _ => i++);
        }

        public ColumnEnumeration(string name, IEnumerable<string> keyValues, int? order = null, string? displayName = null, bool hidden = false) : base(name, order, displayName, hidden)
        {
            int i = 0;
            _keyValues = keyValues.ToDictionary(x => x, _ => i++);
        }

        internal ColumnEnumeration(PropertyNames names, IEnumerable<string> keyValues, bool ignore) : base(names, ignore)
        {
            int i = 0;
            _keyValues = keyValues.ToDictionary(x => x, _ => i++);
        }

        internal ColumnEnumeration(PropertyNames names, IEnumerable<string> keyValues, int? order = null, string? displayName = null, bool hidden = false) : base(names, order, displayName, hidden)
        {
            int i = 0;
            _keyValues = keyValues.ToDictionary(x => x, _ => i++);
        }

        public override void WriteCell(ExcelRange cell, Dictionary<string, PropertyInfo> properties, object? obj)
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
                    cell.Offset(0, offset).Value = enumerator.Value;
                }
            }
            else if (collection is IEnumerable enumerable)
            {
                foreach (object item in enumerable)
                {
                    string key = item.ToString()!;
                    int offset = _keyValues[key];   //this will throw if key is not in the initialized keyValues. This is intentional
                    cell.Offset(0, offset).Value = item;
                }
            }
            else
            {
                throw new NotSupportedException($"only {nameof(IEnumerable)} or {nameof(IDictionary)} are supported");
            }
        }
    }

    public class ColumnEnumeration<T> : ColumnEnumeration
    {
        public ColumnEnumeration(Expression<Func<T, object?>> columnSelector, IEnumerable<string> keyValues, bool ignore) :
            base(ColumnInfo<T>.GetPropertyName(columnSelector), keyValues, ignore)
        { }

        public ColumnEnumeration(Expression<Func<T, object?>> columnSelector, IEnumerable<string> keyValues, int? order = null, string? displayName = null, bool hidden = false) :
            base(ColumnInfo<T>.GetPropertyName(columnSelector), keyValues, order, displayName, hidden)
        { }
    }
}

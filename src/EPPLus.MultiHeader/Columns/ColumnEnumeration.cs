using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EPPLus.MultiHeader.Columns
{
    internal class ColumnEnumeration : ColumnInfo
    {
        private readonly Dictionary<string, int> _keyValues;
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

        public override void WriteCell(ExcelRange cell, Dictionary<string, PropertyInfo> properties, object? obj)
        {
            if (obj == null)
                return;

            var collection = properties[Name].GetValue(obj)!;
            if (collection is IDictionary dictionary)
            {
                var enumerator = dictionary.GetEnumerator();
                while(enumerator.MoveNext())
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
}

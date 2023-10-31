using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EPPLus.MultiHeader.Columns
{
    public class ColumnHyperLilnk : ColumnConfig
    {
        public string UrlPropertyName { get; set; }
        public bool IgnoreLinkErrors { get; set; } = true;

        public ColumnHyperLilnk(string name, string urlPropertyName) : base(name) {
            UrlPropertyName = urlPropertyName;
        }

        public ColumnHyperLilnk(string name, string urlPropertyName, int? order = null, string? displayName = null, bool hidden = false) :
            base(name, order, displayName, hidden)
        {
            UrlPropertyName = urlPropertyName;
        }

        public override void WriteCell(ExcelRange cell, Dictionary<string, PropertyInfo> properties, object obj)
        {
            cell.Value = properties[Name].GetValue(obj);
            object? url = properties[UrlPropertyName].GetValue(obj);
            if (url != null)
            {
                try
                {
                    cell.Hyperlink = new Uri(url.ToString()!);
                }   
                catch (Exception)
                {
                    if (!IgnoreLinkErrors)
                        throw;
                }
            }
        }
    }

    public class ColumnHyperLilnk<T> : ColumnHyperLilnk 
    {
        public ColumnHyperLilnk(Expression<Func<T, object?>> columnSelector, string urlPropertyName) : 
            base(ColumnConfig<T>.GetPropertyName(columnSelector), urlPropertyName) { }

        public ColumnHyperLilnk(Expression<Func<T, object?>> columnSelector, Expression<Func<T, object?>> urlColumnSelector, int? order = null, string? displayName = null, bool hidden = false)
            : base(ColumnConfig <T>.GetPropertyName(columnSelector), ColumnConfig<T>.GetPropertyName(urlColumnSelector), order, displayName, hidden) { }
    }
}

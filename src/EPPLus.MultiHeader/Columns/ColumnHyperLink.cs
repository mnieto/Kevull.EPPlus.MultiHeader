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
    /// <summary>
    /// Add a column with hyperlink. That is, the Excel column is associated to 2 fields: the url and the display content
    /// </summary>
    public class ColumnHyperLink : ColumnInfo
    {
        /// <summary>
        /// Object property that contains the Url
        /// </summary>
        public string UrlPropertyName { get; set; }

        /// <summary>
        /// If <c>false</c> throws an error if the url is not valid. Otherwise ignores malformed url
        /// </summary>
        public bool IgnoreLinkErrors { get; set; } = true;

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="name">name of the property. This will match with the property name</param>
        /// <param name="urlPropertyName">property that contains the url</param>
        public ColumnHyperLink(string name, string urlPropertyName) : base(name)
        {
            UrlPropertyName = urlPropertyName;
        }

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="name">name of the property. This will match with the property name</param>
        /// <param name="urlPropertyName">property that contains the url</param>
        /// <param name="order">Diplay order. Order is relative to the other columns. Columns that has no <paramref name="order"/> are added after those that have it</param>
        /// <param name="displayName">Human friendly name for the column. If not specified, the property Name is used</param>
        /// <param name="hidden">Column is written to the Excel, but it's hidden</param>
        public ColumnHyperLink(string name, string urlPropertyName, int? order = null, string? displayName = null, bool hidden = false) :
            base(name, order, displayName, hidden)
        {
            UrlPropertyName = urlPropertyName;
        }

        /// <summary>
        /// Ctor. Used internaly to build nested objects
        /// </summary>
        /// <param name="names">name of the property. This will match with the property name</param>
        /// <param name="urlPropertyName">property that contains the url</param>
        internal ColumnHyperLink(PropertyNames names, string urlPropertyName) : base(names)
        {
            UrlPropertyName = urlPropertyName;
        }

        /// <summary>
        /// Ctor. Used internaly to build nested objects
        /// </summary>
        /// <param name="names">name of the property. This will match with the property name</param>
        /// <param name="urlPropertyName">property that contains the url</param>
        /// <param name="order">Diplay order. Order is relative to the other columns. Columns that has no <paramref name="order"/> are added after those that have it</param>
        /// <param name="displayName">Human friendly name for the column. If not specified, the property Name is used</param>
        /// <param name="hidden">Column is written to the Excel, but it's hidden</param>
        internal ColumnHyperLink(PropertyNames names, string urlPropertyName, int? order = null, string? displayName = null, bool hidden = false) :
            base(names, order, displayName, hidden)
        {
            UrlPropertyName = urlPropertyName;
        }


        internal override void WriteCell(ExcelRange cell, Dictionary<string, PropertyInfo> properties, object? obj)
        {
            if (obj == null)
                return;
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

    /// <summary>
    /// Add a column with hyperlink. That is, the Excel column is associated to 2 fields: the url and the display content
    /// </summary>
    public class ColumnHyperLink<T> : ColumnHyperLink 
    {
        /// <summary>
        /// Add a column with hyperlink. That is, the Excel column is associated to 2 fields: the url and the display content
        /// </summary>
        /// <param name="columnSelector">Property associated to the column</param>
        /// <param name="urlPropertyName">Property that contains the Url</param>
        public ColumnHyperLink(Expression<Func<T, object?>> columnSelector, string urlPropertyName) : 
            base(ColumnInfo<T>.GetPropertyName(columnSelector), urlPropertyName) { }

        /// <summary>
        /// Add a column with hyperlink. That is, the Excel column is associated to 2 fields: the url and the display content
        /// </summary>
        /// <param name="columnSelector">Property associated to the column</param>
        /// <param name="urlColumnSelector">Property that contains the Url</param>
        /// <param name="order">Diplay order. Order is relative to the other columns. Columns that has no <paramref name="order"/> are added after those that have it</param>
        /// <param name="displayName"></param>
        /// <param name="hidden"></param>
        public ColumnHyperLink(Expression<Func<T, object?>> columnSelector, Expression<Func<T, object?>> urlColumnSelector, int? order = null, string? displayName = null, bool hidden = false)
            : base(ColumnInfo<T>.GetPropertyName(columnSelector), ColumnInfo<T>.GetPropertyName(urlColumnSelector).Name, order, displayName, hidden) { }
    }
}

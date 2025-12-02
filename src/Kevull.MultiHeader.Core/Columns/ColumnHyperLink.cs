using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Kevull.EPPLus.MultiHeader.Columns
{
    /// <summary>
    /// Add a column with hyperlink. That is, the Excel column is associated to 2 fields: the url and the display content
    /// </summary>
    public class ColumnHyperLink<T> : ColumnInfo<T>
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
        /// Add a column with hyperlink. That is, the Excel column is associated to 2 fields: the url and the display content
        /// </summary>
        /// <param name="columnSelector">Property associated to the column</param>
        /// <param name="urlColumnSelector">Property that contains the Url</param>
        /// <param name="order">Diplay order. Order is relative to the other columns. Columns that has no <paramref name="order"/> are added after those that have it</param>
        /// <param name="displayName"></param>
        /// <param name="hidden"></param>
        /// <param name="styleName">Name of a style defined in the Excel workbook</param>
        public ColumnHyperLink(Expression<Func<T, object?>> columnSelector, Expression<Func<T, object?>> urlColumnSelector, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null)
            : base(GetPropertyName(columnSelector), order, displayName, hidden, styleName)
        {
            UrlPropertyName = GetPropertyName(urlColumnSelector).Name;
        }

        /// <summary>
        /// Add a column with hyperlink. That is, the Excel column is associated to 2 fields: the url and the display content
        /// </summary>
        /// <param name="columnSelector">Property associated to the column</param>
        /// <param name="urlColumnSelector">Property that contains the Url</param>
        /// <param name="cfg"> Action that will be invoked to configure the ColumnInfo properties using a <see cref="ColumnDef"/> object</param>
        public ColumnHyperLink(Expression<Func<T, object?>> columnSelector, Expression<Func<T, object?>> urlColumnSelector, Action<ColumnDef> cfg)
            : base(columnSelector, cfg)
        {
            UrlPropertyName = GetPropertyName(urlColumnSelector).Name;
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
}

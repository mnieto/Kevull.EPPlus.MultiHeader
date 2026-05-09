using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using Kevull.MultiHeader.Core;
using Kevull.MultiHeader.Core.Columns;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Text.RegularExpressions;

namespace Kevull.MultiHeader.ClosedXml
{
    /// <summary>
    /// Helper class to configure the report and column options.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ConfigurationBuilder<T> : IConfigurationBuilder<T>
    {
        private readonly List<ColumnInfo> columns;
        private readonly XLWorkbook xls;
        private CellAddress StartingAddress { get; set; } = new CellAddress(1, 1);

        /// <summary>
        /// Collection of named styles available for column and header formatting.
        /// </summary>
        public Dictionary<string, CellFormat> NamedStyles = new Dictionary<string, CellFormat>();

        /// <inheritdoc />
        public bool AutoFilter { get; set; } = true;

        /// <inheritdoc />
        public bool AppendToExistingReport { get; set; }

        /// <inheritdoc />
        public bool AutoFreezePanes { get; set; } = true;

        /// <inheritdoc />
        public int TopRow => StartingAddress.Row;

        /// <inheritdoc />
        public int LeftColumn => StartingAddress.Column;

        /// <summary>
        /// Ctor invoked to get default configuration at first step.
        /// </summary>
        /// <param name="xls">Excel reference.</param>
        public ConfigurationBuilder(XLWorkbook xls) : this(xls, new List<ColumnInfo>()) { }

        /// <summary>
        /// Ctor intended for testing purposes.
        /// </summary>
        /// <param name="xls">Excel reference.</param>
        /// <param name="columns">List of column configurations.</param>
        internal ConfigurationBuilder(XLWorkbook xls, params ColumnInfo[] columns) : this(xls, columns.ToList()) { }

        /// <summary>
        /// Ctor.
        /// </summary>
        /// <param name="xls">Excel reference.</param>
        /// <param name="columns">List of column configurations.</param>
        public ConfigurationBuilder(XLWorkbook xls, IEnumerable<ColumnInfo> columns)
        {
            this.xls = xls;
            this.columns = columns.ToList();
        }

        /// <inheritdoc />
        public IConfigurationBuilder<T> AddColumn(Expression<Func<T, object?>> columnSelector)
        {
            columns.Add(new ColumnInfo<T>(columnSelector));
            return this;
        }

        /// <inheritdoc />
        public IConfigurationBuilder<T> AddColumn(Expression<Func<T, object?>> columnSelector, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null)
        {
            columns.Add(new ColumnInfo<T>(columnSelector, order, displayName, hidden, styleName));
            return this;
        }

        /// <inheritdoc />
        public IConfigurationBuilder<T> AddColumn(Expression<Func<T, object?>> columnSelector, Action<ColumnDef> cfg)
        {
            columns.Add(new ColumnInfo<T>(columnSelector, cfg));
            return this;
        }

        /// <inheritdoc />
        public IConfigurationBuilder<T> AddEnumeration(Expression<Func<T, object?>> columnSelector, IEnumerable<string> keyValues, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null)
        {
            columns.Add(new ColumnEnumeration<T>(columnSelector, keyValues, order, displayName, hidden, styleName));
            return this;
        }

        /// <inheritdoc />
        public IConfigurationBuilder<T> AddEnumeration(Expression<Func<T, object?>> columnSelector, IEnumerable<string> keyValues, Action<ColumnDef> cfg)
        {
            columns.Add(new ColumnEnumeration<T>(columnSelector, keyValues, cfg));
            return this;
        }

        /// <inheritdoc />
        public IConfigurationBuilder<T> AddExpression(string name, Func<T, object?> expression, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null)
        {
            columns.Add(new ColumnExpression<T>(name, expression, order, displayName, hidden, styleName));
            return this;
        }

        /// <inheritdoc />
        public IConfigurationBuilder<T> AddExpression(string name, Func<T, object?> expression, Action<ColumnDef> cfg)
        {
            columns.Add(new ColumnExpression<T>(name, expression, cfg));
            return this;
        }

        /// <inheritdoc />
        public IConfigurationBuilder<T> AddFormula(string name, string formula, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null)
        {
            columns.Add(new ColumnFormula(name, formula, order, displayName, hidden, styleName));
            return this;
        }

        /// <inheritdoc />
        public IConfigurationBuilder<T> AddFormula(string name, string formula, Action<ColumnDef> cfg)
        {
            columns.Add(new ColumnFormula(name, formula, cfg));
            return this;
        }

        /// <inheritdoc />
        public IConfigurationBuilder<T> AddHyperLinkColumn(Expression<Func<T, object?>> columnSelector, Expression<Func<T, object?>> urlColumnSelector, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null)
        {
            columns.Add(new ColumnHyperLink<T>(columnSelector, urlColumnSelector, order, displayName, hidden, styleName));
            return this;
        }

        /// <inheritdoc />
        public IConfigurationBuilder<T> AddHyperLinkColumn(Expression<Func<T, object?>> columnSelector, Expression<Func<T, object?>> urlColumnSelector, Action<ColumnDef> cfg)
        {
            columns.Add(new ColumnHyperLink<T>(columnSelector, urlColumnSelector, cfg));
            return this;
        }

        /// <inheritdoc />
        public IConfigurationBuilder<T> IgnoreColumn(Expression<Func<T, object?>> columnSelector)
        {
            columns.Add(new ColumnInfo<T>(columnSelector, true));
            return this;
        }

        /// <inheritdoc />
        public IConfigurationBuilder<T> AddHeaderStyle(Action<CellFormat> style)
        {
            return AddNamedStyle(StyleNames.HeaderStyleName, style);
        }

        /// <inheritdoc />
        public IConfigurationBuilder<T> AddNamedStyle(string name, Action<CellFormat> style)
        {
            NamedStyles.Add(name, new CellFormat());
            style?.Invoke(NamedStyles[name]);
            return this;
        }

        /// <inheritdoc />
        public IConfigurationBuilder<T> SetStartingAddress(string address)
        {
            StartingAddress = ParseAddress(address);
            return this;
        }

        /// <inheritdoc />
        public IConfigurationBuilder<T> SetStartingAddres(int row, int column)
        {
            StartingAddress = new CellAddress(row, column);
            return this;
        }

        /// <inheritdoc />
        public HeaderManager<T> Build()
        {
            var headerManager = new HeaderManager<T>(columns);
            headerManager.AutoFilter = AutoFilter;
            headerManager.AutoFreezePanes = AutoFreezePanes;
            headerManager.FirstRow = StartingAddress.Row;
            headerManager.FirstColumn = StartingAddress.Column;
            headerManager.AppendToExistingReport = AppendToExistingReport;
            return headerManager;
        }

        private static CellAddress ParseAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
                throw new ArgumentNullException(nameof(address));

            var match = Regex.Match(address.Trim(), "^(?<col>[A-Za-z]+)(?<row>[1-9][0-9]*)$");
            if (!match.Success)
                throw new FormatException($"Invalid cell address: {address}");

            var colLetters = match.Groups["col"].Value.ToUpperInvariant();
            var row = int.Parse(match.Groups["row"].Value, CultureInfo.InvariantCulture);

            int col = 0;
            foreach (var ch in colLetters)
            {
                col = (col * 26) + (ch - 'A' + 1);
            }

            return new CellAddress(row, col);
        }

        private sealed class CellAddress
        {
            public int Row { get; }
            public int Column { get; }

            public CellAddress(int row, int column)
            {
                if (row <= 0) throw new ArgumentOutOfRangeException(nameof(row));
                if (column <= 0) throw new ArgumentOutOfRangeException(nameof(column));
                Row = row;
                Column = column;
            }
        }
    }
}

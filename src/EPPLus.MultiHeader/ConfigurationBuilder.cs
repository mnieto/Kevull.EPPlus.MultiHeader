using EPPLus.MultiHeader.Columns;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Linq.Expressions;
using System.Reflection;

namespace EPPLus.MultiHeader
{
    /// <summary>
    /// Helper class to configure the report and column options
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ConfigurationBuilder<T>
    {
        private List<ColumnInfo> columns;

        /// <summary>
        /// Shows or not autofilter on last header row
        /// </summary>
        public bool AutoFilter { get; set; } = true;

        /// <summary>
        /// Ctor invoked to get default configuration at first step
        /// </summary>
        public ConfigurationBuilder() : this(new List<ColumnInfo>()) { }

        /// <summary>
        /// Ctor intended for testing purposes
        /// </summary>
        /// <param name="columns">List of column configurations</param>
        internal ConfigurationBuilder(params ColumnInfo[] columns) : this(columns.ToList()) { }

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="columns">List of column configurations</param>
        public ConfigurationBuilder(IEnumerable<ColumnInfo> columns)
        {
            this.columns = columns.ToList();
        }

        /// <summary>
        /// Adds a column with default configuration
        /// </summary>
        public ConfigurationBuilder<T> AddColumn(Expression<Func<T, object?>> columnSelector)
        {
            columns.Add(new ColumnInfo<T>(columnSelector));
            return this;
        }

        /// <summary>
        /// Add a column
        /// </summary>
        /// <param name="columnSelector">Allows specify the column name</param>
        /// <param name="order">Diplay order. Order is relative to the other columns. Columns that has no <paramref name="order"/> are added after those that have it</param>
        /// <param name="displayName">Human friendly name for the column. If not specified, the property Name is used</param>
        /// <param name="hidden">Column is written to the Excel, but it's hidden</param>
        public ConfigurationBuilder<T> AddColumn(Expression<Func<T, object?>> columnSelector, int? order = null, string? displayName = null, bool hidden = false)
        {
            columns.Add(new ColumnInfo<T>(columnSelector, order, displayName, hidden));
            return this;
        }

        /// <summary>
        /// Add a column whose type implements <see cref="IDictionary{TKey, TValue}" or <see cref="IEnumerable{T}"/> where Tkey is always invoked with <see cref="Object.ToString()"/> />
        /// </summary>
        /// <param name="columnSelector">Allows specify the column name</param>
        /// <param name="keyValues">Allowed key values. This is used to allocate a specific number of columns</param>
        /// <param name="order">Diplay order. Order is relative to the other columns. Columns that has no <paramref name="order"/> are added after those that have it</param>
        /// <param name="displayName">Human friendly name for the column. If not specified, the property Name is used</param>
        /// <param name="hidden">Column is written to the Excel, but it's hidden</param>
        public ConfigurationBuilder<T> AddEnumeration(Expression<Func<T, object?>> columnSelector, IEnumerable<string> keyValues, int? order = null, string? displayName = null, bool hidden = false)
        {
            columns.Add(new ColumnEnumeration<T>(columnSelector, keyValues, order, displayName, hidden));
            return this;
        }

        /// <summary>
        /// Add an expression column. That is, each time the report will render a value for this column, it will invoke the <paramref name="expression"/> lambda.
        /// </summary>
        /// <param name="name">name of the property. In this case, it cannot be infered from the source Type</param>
        /// <param name="expression">Lambda expression to be evaluated to render the column value each row</param>
        /// <param name="order">Diplay order. Order is relative to the other columns. Columns that has no <paramref name="order"/> are added after those that have it</param>
        /// <param name="displayName">Human friendly name for the column. If not specified, the property Name is used</param>
        /// <param name="hidden">Column is written to the Excel, but it's hidden</param>
        public ConfigurationBuilder<T> AddExpression(string name, Func<T, object?> expression, int? order = null, string? displayName = null, bool hidden = false)
        {
            columns.Add(new ColumnExpression<T>(name, expression, order, displayName, hidden));
            return this;
        }

        /// <summary>
        /// Add a formula column. That is, each time the report will render a value for this column, it will use the specified <paramref name="formula"/>.
        /// </summary>
        /// <param name="name">name of the property. In this case, it cannot be infered from the source Type</param>
        /// <param name="formula">Formula used for this column. Be sure to use the correct absulte/relative references in the formula</param>
        /// <param name="order">Diplay order. Order is relative to the other columns. Columns that has no <paramref name="order"/> are added after those that have it</param>
        /// <param name="displayName">Human friendly name for the column. If not specified, the property Name is used</param>
        /// <param name="hidden">Column is written to the Excel, but it's hidden</param>
        public ConfigurationBuilder<T> AddFormula(string name, string formula, int? order = null, string? displayName = null, bool hidden = false)
        {
            columns.Add(new ColumnFormula(name, formula, order, displayName, hidden));
            return this;
        }

        /// <summary>
        /// Add a column whose cells will contain a hyperlink
        /// </summary>
        /// <param name="columnSelector">Allows specify the column name</param>
        /// <param name="urlColumnSelector">Allows specify the column wich will contain the url</param>
        /// <param name="order">Diplay order. Order is relative to the other columns. Columns that has no <paramref name="order"/> are added after those that have it</param>
        /// <param name="displayName">Human friendly name for the column. If not specified, the property Name is used</param>
        /// <param name="hidden">Column is written to the Excel, but it's hidden</param>
        public ConfigurationBuilder<T> AddHyperLinkColumn(Expression<Func<T, object?>> columnSelector, Expression<Func<T, object?>> urlColumnSelector, int? order = null, string? displayName = null, bool hidden = false)
        {
            columns.Add(new ColumnHyperLink<T>(columnSelector, urlColumnSelector, order, displayName, hidden));
            return this;
        }

        /// <summary>
        /// Ignore this property. This column will not be rendered
        /// </summary>
        /// <param name="columnSelector">Allows specify the column name</param>
        public ConfigurationBuilder<T> IgnoreColumn(Expression<Func<T, object?>> columnSelector)
        {
            columns.Add(new ColumnInfo<T>(columnSelector, true));
            return this;
        }

        /// <summary>
        /// Adds a custom header style. If not specified, a default one will be applyed
        /// </summary>
        /// <param name="xls"><see cref="ExcelPackage"/> where the style will be defined</param>
        /// <param name="style">Lambda expresion to define the style</param>
        /// <remarks>The default style is defined as below</remarks>
        /// <example>
        /// <code>
        /// var namedStyle = _xls.Workbook.Styles.CreateNamedStyle("Headers");
        /// namedStyle.Style.Border.Left.Style = ExcelBorderStyle.Thin;
        /// namedStyle.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        /// namedStyle.Style.Border.Top.Style = ExcelBorderStyle.Thin;
        /// namedStyle.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        /// namedStyle.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        /// namedStyle.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        /// namedStyle.Style.Fill.SetBackground(Color.LightGray, ExcelFillStyle.Solid);
        /// namedStyle.Style.Font.Bold = true;
        /// </code>
        /// </example>
        public ConfigurationBuilder<T> AddHeaderStyle(ExcelPackage xls, Action<ExcelStyle> style)
        {
            var namedStyle = xls.Workbook.Styles.CreateNamedStyle(MultiHeaderReport<T>.HeaderStyleName);
            style(namedStyle.Style);
            return this;
        }

        /// <summary>
        /// Garther all the provided information to generate the needed internal structures
        /// </summary>
        public HeaderManager<T> Build()
        {
            var headerManager = new HeaderManager<T>(columns);
            headerManager.AutoFilter = AutoFilter;
            return headerManager;
        }

    }
}
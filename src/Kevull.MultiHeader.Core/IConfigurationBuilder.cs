using Kevull.MultiHeader.Core;
using Kevull.MultiHeader.Core.Columns;
using System.Linq.Expressions;

namespace Kevull.MultiHeader.Core
{
    /// <summary>
    /// Defines the fluent configuration contract for building a multi-header report definition.
    /// </summary>
    /// <typeparam name="T">The type of the source data used to generate report rows.</typeparam>
    public interface IConfigurationBuilder<T>
    {
        /// <summary>
        /// Gets or sets a value indicating whether the generated output should be appended to an existing report.
        /// </summary>
        bool AppendToExistingReport { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether auto-filter should be enabled for the generated header row.
        /// </summary>
        bool AutoFilter { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether panes should be frozen automatically at the starting position.
        /// </summary>
        bool AutoFreezePanes { get; set; }

        /// <summary>
        /// Gets the starting left column index for writing the report.
        /// </summary>
        int LeftColumn { get; }

        /// <summary>
        /// Gets the starting top row index for writing the report.
        /// </summary>
        int TopRow { get; }

        /// <summary>
        /// Adds a column mapped to the specified model property.
        /// </summary>
        /// <param name="columnSelector">The expression selecting the model property to map as a column.</param>
        /// <returns>The current configuration builder instance.</returns>
        IConfigurationBuilder<T> AddColumn(Expression<Func<T, object?>> columnSelector);

        /// <summary>
        /// Adds a column mapped to the specified model property with optional metadata.
        /// </summary>
        /// <param name="columnSelector">The expression selecting the model property to map as a column.</param>
        /// <param name="order">The zero-based display order of the column.</param>
        /// <param name="displayName">The display name shown in the header.</param>
        /// <param name="hidden">A value indicating whether the column is hidden.</param>
        /// <param name="styleName">The name of a style to apply to the column.</param>
        /// <returns>The current configuration builder instance.</returns>
        IConfigurationBuilder<T> AddColumn(Expression<Func<T, object?>> columnSelector, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null);

        /// <summary>
        /// Adds a column mapped to the specified model property and configures it using a custom action.
        /// </summary>
        /// <param name="columnSelector">The expression selecting the model property to map as a column.</param>
        /// <param name="cfg">The configuration action for the created column definition.</param>
        /// <returns>The current configuration builder instance.</returns>
        IConfigurationBuilder<T> AddColumn(Expression<Func<T, object?>> columnSelector, Action<ColumnDef> cfg);

        /// <summary>
        /// Adds an enumeration-based column mapped to the specified model property.
        /// </summary>
        /// <param name="columnSelector">The expression selecting the model property to map as a column.</param>
        /// <param name="keyValues">The ordered list of allowed values for the enumeration column.</param>
        /// <param name="order">The zero-based display order of the column.</param>
        /// <param name="displayName">The display name shown in the header.</param>
        /// <param name="hidden">A value indicating whether the column is hidden.</param>
        /// <param name="styleName">The name of a style to apply to the column.</param>
        /// <returns>The current configuration builder instance.</returns>
        IConfigurationBuilder<T> AddEnumeration(Expression<Func<T, object?>> columnSelector, IEnumerable<string> keyValues, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null);

        /// <summary>
        /// Adds an enumeration-based column mapped to the specified model property and configures it using a custom action.
        /// </summary>
        /// <param name="columnSelector">The expression selecting the model property to map as a column.</param>
        /// <param name="keyValues">The ordered list of allowed values for the enumeration column.</param>
        /// <param name="cfg">The configuration action for the created column definition.</param>
        /// <returns>The current configuration builder instance.</returns>
        IConfigurationBuilder<T> AddEnumeration(Expression<Func<T, object?>> columnSelector, IEnumerable<string> keyValues, Action<ColumnDef> cfg);

        /// <summary>
        /// Adds a calculated expression column.
        /// </summary>
        /// <param name="name">The internal column name.</param>
        /// <param name="expression">The value expression evaluated for each data item.</param>
        /// <param name="order">The zero-based display order of the column.</param>
        /// <param name="displayName">The display name shown in the header.</param>
        /// <param name="hidden">A value indicating whether the column is hidden.</param>
        /// <param name="styleName">The name of a style to apply to the column.</param>
        /// <returns>The current configuration builder instance.</returns>
        IConfigurationBuilder<T> AddExpression(string name, Func<T, object?> expression, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null);

        /// <summary>
        /// Adds a calculated expression column and configures it using a custom action.
        /// </summary>
        /// <param name="name">The internal column name.</param>
        /// <param name="expression">The value expression evaluated for each data item.</param>
        /// <param name="cfg">The configuration action for the created column definition.</param>
        /// <returns>The current configuration builder instance.</returns>
        IConfigurationBuilder<T> AddExpression(string name, Func<T, object?> expression, Action<ColumnDef> cfg);

        /// <summary>
        /// Adds a formula column.
        /// </summary>
        /// <param name="name">The internal column name.</param>
        /// <param name="formula">The Excel formula to apply.</param>
        /// <param name="order">The zero-based display order of the column.</param>
        /// <param name="displayName">The display name shown in the header.</param>
        /// <param name="hidden">A value indicating whether the column is hidden.</param>
        /// <param name="styleName">The name of a style to apply to the column.</param>
        /// <returns>The current configuration builder instance.</returns>
        IConfigurationBuilder<T> AddFormula(string name, string formula, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null);

        /// <summary>
        /// Adds a formula column and configures it using a custom action.
        /// </summary>
        /// <param name="name">The internal column name.</param>
        /// <param name="formula">The Excel formula to apply.</param>
        /// <param name="cfg">The configuration action for the created column definition.</param>
        /// <returns>The current configuration builder instance.</returns>
        IConfigurationBuilder<T> AddFormula(string name, string formula, Action<ColumnDef> cfg);

        /// <summary>
        /// Adds a style to be applied to header cells.
        /// </summary>
        /// <param name="style">The action that configures header cell formatting.</param>
        /// <returns>The current configuration builder instance.</returns>
        IConfigurationBuilder<T> AddHeaderStyle(Action<CellFormat> style);

        /// <summary>
        /// Adds a hyperlink column using one property for display text and another for URL.
        /// </summary>
        /// <param name="columnSelector">The expression selecting the model property used as hyperlink display text.</param>
        /// <param name="urlColumnSelector">The expression selecting the model property used as hyperlink URL.</param>
        /// <param name="order">The zero-based display order of the column.</param>
        /// <param name="displayName">The display name shown in the header.</param>
        /// <param name="hidden">A value indicating whether the column is hidden.</param>
        /// <param name="styleName">The name of a style to apply to the column.</param>
        /// <returns>The current configuration builder instance.</returns>
        IConfigurationBuilder<T> AddHyperLinkColumn(Expression<Func<T, object?>> columnSelector, Expression<Func<T, object?>> urlColumnSelector, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null);

        /// <summary>
        /// Adds a hyperlink column and configures it using a custom action.
        /// </summary>
        /// <param name="columnSelector">The expression selecting the model property used as hyperlink display text.</param>
        /// <param name="urlColumnSelector">The expression selecting the model property used as hyperlink URL.</param>
        /// <param name="cfg">The configuration action for the created column definition.</param>
        /// <returns>The current configuration builder instance.</returns>
        IConfigurationBuilder<T> AddHyperLinkColumn(Expression<Func<T, object?>> columnSelector, Expression<Func<T, object?>> urlColumnSelector, Action<ColumnDef> cfg);

        /// <summary>
        /// Adds a named reusable style.
        /// </summary>
        /// <param name="name">The style name.</param>
        /// <param name="style">The action that configures the style format.</param>
        /// <returns>The current configuration builder instance.</returns>
        IConfigurationBuilder<T> AddNamedStyle(string name, Action<CellFormat> style);

        /// <summary>
        /// Marks the specified model property to be ignored during report generation.
        /// </summary>
        /// <param name="columnSelector">The expression selecting the model property to ignore.</param>
        /// <returns>The current configuration builder instance.</returns>
        IConfigurationBuilder<T> IgnoreColumn(Expression<Func<T, object?>> columnSelector);

        /// <summary>
        /// Sets the starting position used to place the report in the worksheet.
        /// </summary>
        /// <param name="row">The starting row index.</param>
        /// <param name="column">The starting column index.</param>
        /// <returns>The current configuration builder instance.</returns>
        IConfigurationBuilder<T> SetStartingAddres(int row, int column);

        /// <summary>
        /// Sets the starting position used to place the report in the worksheet.
        /// </summary>
        /// <param name="address">The Excel address (for example, <c>A1</c>) for the top-left report cell.</param>
        /// <returns>The current configuration builder instance.</returns>
        IConfigurationBuilder<T> SetStartingAddress(string address);

        /// <summary>
        /// Builds the configured header manager used for report generation.
        /// </summary>
        /// <returns>The built <see cref="HeaderManager{T}"/> instance.</returns>
        HeaderManager<T> Build();
    }
}
using System;
using System.Collections.Generic;
using System.Text;

namespace Kevull.MultiHeader.Core
{

    /// <summary>
    /// Interface for writing and formatting Excel cells in a library-agnostic way
    /// </summary>
    public interface IExcelWriter
    {

        /// <summary>
        /// Writes a value to a specific cell
        /// </summary>
        /// <param name="row">Row number (1-based)</param>
        /// <param name="col">Column number (1-based)</param>
        /// <param name="value">Value to write to the cell (can be null for empty cells)</param>
        void WriteCell(int row, int col, object? value);

        /// <summary>
        /// Writes a value to a range of cells
        /// </summary>
        /// <param name="fromRow">Starting row number (1-based)</param>
        /// <param name="fromCol">Starting column number (1-based)</param>
        /// <param name="toRow">Ending row number (1-based)</param>
        /// <param name="toCol">Ending column number (1-based)</param>
        /// <param name="value">Value to write to all cells in the range (can be null for empty cells)</param>
        void WriteCell(int fromRow, int fromCol, int toRow, int toCol, object? value);

        /// <summary>
        /// Writes a value with a hyperlink to a specific cell
        /// </summary>
        /// <param name="row">Row number (1-based)</param>
        /// <param name="col">Column number (1-based)</param>
        /// <param name="value">Value to display in the cell (can be null)</param>
        /// <param name="url">URL for the hyperlink</param>
        void WriteCellWithHyperlink(int row, int col, object? value, string url);

        /// <summary>
        /// Writes a formula to a specific cell</summary>
        /// <param name="row">Row number (1-based)</param>
        /// <param name="col">Column number (1-based)</param>
        /// <param name="formula">Excel formula (without the leading '=')</param>
        void WriteFormula(int row, int col, string formula);

        /// <summary>
        /// Writes a formula to a range of cells
        /// </summary>
        /// <param name="fromRow">Starting row number (1-based)</param>
        /// <param name="fromCol">Starting column number (1-based)</param>
        /// <param name="toRow">Ending row number (1-based)</param>
        /// <param name="toCol">Ending column number (1-based)</param>
        /// <param name="formula">Excel formula (without the leading '=')</param>
        void WriteFormula(int fromRow, int fromCol, int toRow, int toCol, string formula);

        /// <summary>
        /// Applies formatting to a specific cell using the library-agnostic <see cref="CellFormat"/>
        /// </summary>
        /// <param name="row">Row number (1-based)</param>
        /// <param name="col">Column number (1-based)</param>
        /// <param name="format">Cell format to apply</param>
        void ApplyFormat(int row, int col, CellFormat format);

        /// <summary>
        /// Applies formatting to a range of cells using the library-agnostic <see cref="CellFormat"/>
        /// </summary>
        /// <param name="fromRow">Starting row number (1-based)</param>
        /// <param name="fromCol">Starting column number (1-based)</param>
        /// <param name="toRow">Ending row number (1-based)</param>
        /// <param name="toCol">Ending column number (1-based)</param>
        /// <param name="format">Cell format to apply</param>
        void ApplyFormat(int fromRow, int fromCol, int toRow, int toCol, CellFormat format);

        /// <summary>
        /// Applies native library-specific formatting to a specific cell
        /// </summary>
        /// <param name="row">Row number (1-based)</param>
        /// <param name="col">Column number (1-based)</param>
        /// <param name="format">Action that receives the native cell/range object for direct manipulation</param>
        void ApplyNativeFormat(int row, int col, Action<object> format);

        /// <summary>
        /// Applies native library-specific formatting to a range of cells
        /// </summary>
        /// <param name="fromRow">Starting row number (1-based)</param>
        /// <param name="fromCol">Starting column number (1-based)</param>
        /// <param name="toRow">Ending row number (1-based)</param>
        /// <param name="toCol">Ending column number (1-based)</param>
        /// <param name="format">Action that receives the native cell/range object for direct manipulation</param>
        void ApplyNativeFormat(int fromRow, int fromCol, int toRow, int toCol, Action<object> format);

        /// <summary>
        /// Creates a named style that can be reused across multiple cells
        /// </summary>
        /// <param name="styleName">Unique name for the style</param>
        /// <param name="format">Cell format defining the style properties</param>
        void CreateNamedStyle(string styleName, CellFormat format);

        /// <summary>
        /// Applies a previously created named style to a specific cell
        /// </summary>
        /// <param name="row">Row number (1-based)</param>
        /// <param name="col">Column number (1-based)</param>
        /// <param name="styleName">Name of the style to apply</param>
        void ApplyNamedStyle(int row, int col, string styleName);

        /// <summary>
        /// Applies a previously created named style to a range of cells
        /// </summary>
        /// <param name="fromRow">Starting row number (1-based)</param>
        /// <param name="fromCol">Starting column number (1-based)</param>
        /// <param name="toRow">Ending row number (1-based)</param>
        /// <param name="toCol">Ending column number (1-based)</param>
        /// <param name="styleName">Name of the style to apply</param>
        void ApplyNamedStyle(int fromRow, int fromCol, int toRow, int toCol, string styleName);

        /// <summary>
        /// Checks if a named style exists
        /// </summary>
        /// <param name="styleName">Name of the style to check</param>
        /// <returns>True if the style exists, false otherwise</returns>
        bool NamedStyleExists(string styleName);

        /// <summary>
        /// Merges a range of cells into a single cell
        /// </summary>
        /// <param name="fromRow">Starting row number (1-based)</param>
        /// <param name="fromCol">Starting column number (1-based)</param>
        /// <param name="toRow">Ending row number (1-based)</param>
        /// <param name="toCol">Ending column number (1-based)</param>
        void Merge(int fromRow, int fromCol, int toRow, int toCol);

        /// <summary>
        /// Auto-fits a column width to its content
        /// </summary>
        /// <param name="col">Column number (1-based)</param>
        void AutoFitColumn(int col);

        /// <summary>
        /// Auto-fits a column width to its content with minimum and maximum width constraints
        /// </summary>
        /// <param name="col">Column number (1-based)</param>
        /// <param name="minWidth">Minimum width for the column</param>
        /// <param name="maxWidth">Maximum width for the column</param>
        void AutoFitColumn(int col, double minWidth, double maxWidth);

        /// <summary>
        /// Sets the width of a column
        /// </summary>
        /// <param name="col">Column number (1-based)</param>
        /// <param name="width">Width value for the column</param>
        void SetColumnWidth(int col, double width);

        /// <summary>
        /// Hides or shows a column
        /// </summary>
        /// <param name="col">Column number (1-based)</param>
        /// <param name="hidden">True to hide the column, false to show it</param>
        void SetColumnHidden(int col, bool hidden);

        /// <summary>
        /// Enables or disables AutoFilter for a range of cells
        /// </summary>
        /// <param name="fromRow">Starting row number (1-based)</param>
        /// <param name="fromCol">Starting column number (1-based)</param>
        /// <param name="toRow">Ending row number (1-based)</param>
        /// <param name="toCol">Ending column number (1-based)</param>
        /// <param name="autoFilter">True to enable AutoFilter, false to disable it</param>
        void SetAutoFilter(int fromRow, int fromCol, int toRow, int toCol, bool autoFilter);

        /// <summary>
        /// Freezes panes at the specified position
        /// </summary>
        /// <param name="row">Row number where the freeze should occur (1-based)</param>
        /// <param name="col">Column number where the freeze should occur (1-based)</param>
        void FreezePanes(int row, int col);

        /// <summary>
        /// Forces recalculation of all formulas in the worksheet
        /// </summary>
        void Recalculate();

    }
}

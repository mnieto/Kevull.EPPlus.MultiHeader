using Kevull.MultiHeader.Core;
using Kevull.MultiHeader.Core.Columns;
using System.Linq.Expressions;

namespace Kevull.MultiHeader.Core
{
    public interface IConfigurationBuilder<T>
    {
        bool AppendToExistingReport { get; set; }
        bool AutoFilter { get; set; }
        bool AutoFreezePanes { get; set; }
        int LeftColumn { get; }
        int TopRow { get; }

        IConfigurationBuilder<T> AddColumn(Expression<Func<T, object?>> columnSelector);
        IConfigurationBuilder<T> AddColumn(Expression<Func<T, object?>> columnSelector, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null);
        IConfigurationBuilder<T> AddColumn(Expression<Func<T, object?>> columnSelector, Action<ColumnDef> cfg);
        IConfigurationBuilder<T> AddEnumeration(Expression<Func<T, object?>> columnSelector, IEnumerable<string> keyValues, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null);
        IConfigurationBuilder<T> AddEnumeration(Expression<Func<T, object?>> columnSelector, IEnumerable<string> keyValues, Action<ColumnDef> cfg);
        IConfigurationBuilder<T> AddExpression(string name, Func<T, object?> expression, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null);
        IConfigurationBuilder<T> AddExpression(string name, Func<T, object?> expression, Action<ColumnDef> cfg);
        IConfigurationBuilder<T> AddFormula(string name, string formula, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null);
        IConfigurationBuilder<T> AddFormula(string name, string formula, Action<ColumnDef> cfg);
        IConfigurationBuilder<T> AddHeaderStyle(Action<CellFormat> style);
        IConfigurationBuilder<T> AddHyperLinkColumn(Expression<Func<T, object?>> columnSelector, Expression<Func<T, object?>> urlColumnSelector, int? order = null, string? displayName = null, bool hidden = false, string? styleName = null);
        IConfigurationBuilder<T> AddHyperLinkColumn(Expression<Func<T, object?>> columnSelector, Expression<Func<T, object?>> urlColumnSelector, Action<ColumnDef> cfg);
        IConfigurationBuilder<T> AddNamedStyle(string name, Action<CellFormat> style);
        HeaderManager<T> Build();
        IConfigurationBuilder<T> IgnoreColumn(Expression<Func<T, object?>> columnSelector);
        IConfigurationBuilder<T> SetStartingAddres(int row, int column);
        IConfigurationBuilder<T> SetStartingAddress(string address);
    }
}
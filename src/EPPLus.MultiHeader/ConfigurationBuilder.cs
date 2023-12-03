using EPPLus.MultiHeader.Columns;
using OfficeOpenXml;
using System.Linq.Expressions;
using System.Reflection;

namespace EPPLus.MultiHeader
{
    public class ConfigurationBuilder<T>
    {
        private List<ColumnInfo> columns;

        public ConfigurationBuilder() : this(new List<ColumnInfo>()) { }
        public ConfigurationBuilder(params ColumnInfo[] config): this(config.ToList()) { }

        public ConfigurationBuilder(IEnumerable<ColumnInfo> columns)
        {
            this.columns = columns.ToList();
        }

        public ConfigurationBuilder<T> AddColumn(Expression<Func<T, object?>> columnSelector)
        {
            columns.Add(new ColumnInfo<T>(columnSelector));
            return this;
        }

        public ConfigurationBuilder<T> AddColumn(Expression<Func<T, object?>> columnSelector, int? order = null, string? displayName = null, bool hidden = false)
        {
            columns.Add(new ColumnInfo<T>(columnSelector, order, displayName, hidden));
            return this;
        }

        public ConfigurationBuilder<T> AddEnumeration(Expression<Func<T, object?>> columnSelector, IEnumerable<string> keyValues, int? order = null, string? displayName = null, bool hidden = false)
        {
            columns.Add(new ColumnEnumeration<T>(columnSelector, keyValues, order, displayName, hidden));
            return this;
        }

        public ConfigurationBuilder<T> AddExpression(string name, Func<T, object?> expression, int? order = null, string? displayName = null, bool hidden = false)
        {
            columns.Add(new ColumnExpression<T>(name, expression, order, displayName, hidden));
            return this;
        }

        public ConfigurationBuilder<T> AddFormula(string name, string formula, int? order = null, string? displayName = null, bool hidden = false)
        {
            columns.Add(new ColumnFormula(name, formula, order, displayName, hidden));
            return this;
        }

        public ConfigurationBuilder<T> AddHyperLinkColumn(Expression<Func<T, object?>> columnSelector, Expression<Func<T, object?>> urlColumnSelector, int? order = null, string? displayName = null, bool hidden = false)
        {
            columns.Add(new ColumnHyperLilnk<T>(columnSelector, urlColumnSelector, order, displayName, hidden));
            return this;
        }

        public ConfigurationBuilder<T> IgnoreColumn(Expression<Func<T, object?>> columnSelector)
        {
            columns.Add(new ColumnInfo<T>(columnSelector, true));
            return this;
        }

        public List<ColumnInfo> Build()
        {
            return columns;
        }

    }
}
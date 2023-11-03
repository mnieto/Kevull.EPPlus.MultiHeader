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
            var result = new List<ColumnInfo>();
            var properties = typeof(T).GetTypeInfo().GetProperties();
            foreach (var property in properties)
            {
                var columConfig = columns.FirstOrDefault(x => x.Name.Equals(property.Name));
                if (columConfig == null)
                {
                    result.Add(new ColumnInfo(property.Name));
                } else if (ShouldAddColumn(columConfig)) { 
                    result.Add(columConfig);
                }
            }
            result.AddRange(columns.Where(x => x.IsDynamic));
            return SetupColumnsOrder(result);
        }

        private List<ColumnInfo> SetupColumnsOrder(List<ColumnInfo> columns)
        {
            int c = 0;
            int previous = 0;
            var tempList = columns.Where(x => x.Order.HasValue).OrderBy(x => x.Order).ToList();
            tempList.AddRange(columns.Where(x => x.Order == null));
            for (int i = 0; i < tempList.Count; i++)
            {
                var item = tempList[i];
                if (item.Order.HasValue)
                {
                    c = item.Order.Value;
                    if (i == 0) {
                        previous = c;
                    } else if (c == previous)
                    {
                        throw new InvalidOperationException($"Repeated order for columns {tempList[i].Name} and {tempList[i - 1].Name}");
                    }
                }
                else
                {
                    item.Order = ++c;
                }
            }
            return tempList;
        }

        private bool ShouldAddColumn(ColumnInfo columConfig)
        {
            return !columConfig.Ignore;
        }
    }
}
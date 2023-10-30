using System.Linq.Expressions;
using System.Reflection;

namespace EPPLus.MultiHeader
{
    public class ConfigurationBuilder<T>
    {
        private List<ColumnConfig> columns;

        public ConfigurationBuilder() : this(new List<ColumnConfig>()) { }
        public ConfigurationBuilder(params ColumnConfig[] config): this(config.ToList()) { }

        public ConfigurationBuilder(IEnumerable<ColumnConfig> columns)
        {
            this.columns = columns.ToList();
        }

        public ConfigurationBuilder<T> AddColumn(Expression<Func<T, object?>> columnSelector)
        {
            columns.Add(new ColumnConfig<T>(columnSelector));
            return this;
        }

        public ConfigurationBuilder<T> AddColumn(Expression<Func<T, object?>> columnSelector, int? order = null, string? displayName = null, bool hidden = false)
        {
            columns.Add(new ColumnConfig<T>(columnSelector, order, displayName, hidden));
            return this;
        }

        public ConfigurationBuilder<T> IgnoreColumn(Expression<Func<T, object?>> columnSelector)
        {
            columns.Add(new ColumnConfig<T>(columnSelector, true));
            return this;
        }

        public List<ColumnConfig> Build()
        {
            var result = new List<ColumnConfig>();
            var properties = typeof(T).GetTypeInfo().GetProperties();
            foreach (var property in properties)
            {
                var columConfig = columns.FirstOrDefault(x => x.Name.Equals(property.Name));
                if (columConfig == null)
                {
                    result.Add(new ColumnConfig(property.Name));
                } else if (ShouldAddColumn(columConfig)) { 
                    result.Add(columConfig);
                }
            }
            //Add dynamic columns here
            return SetupColumnsOrder(result);
        }

        private List<ColumnConfig> SetupColumnsOrder(List<ColumnConfig> columns)
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

        private bool ShouldAddColumn(ColumnConfig columConfig)
        {
            return !columConfig.Ignore;
        }
    }
}
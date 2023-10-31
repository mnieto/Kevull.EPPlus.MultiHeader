using EPPLus.MultiHeader.Columns;
using System.Reflection;

namespace EPPLus.MultiHeader
{
    public class HeaderManager<T>
    {
        public List<ColumnConfig> Columns { get; set; }
        public Dictionary<string, PropertyInfo> Properties { get; set; }
 

        public HeaderManager() {
            (Columns, Properties) = BuildHeaders();
        }

        public HeaderManager(List<ColumnConfig> columns)
        {
            Columns = columns;
            var properties = typeof(T).GetTypeInfo().GetProperties();
            Properties = properties.ToDictionary(x => x.Name, x => x);
        }

        private (List<ColumnConfig>, Dictionary<string, PropertyInfo>) BuildHeaders()
        {
            var result = new List<ColumnConfig>();
            var properties = typeof(T).GetTypeInfo().GetProperties();
            foreach (var property in properties)
            {
                result.Add(new ColumnConfig(property.Name));
            }
            return (result, properties.ToDictionary(x => x.Name, x => x));
        }

    }
}
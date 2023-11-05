using EPPLus.MultiHeader.Columns;
using System.Reflection;

namespace EPPLus.MultiHeader
{
    public class HeaderManager
    {
        public List<ColumnInfo> Columns { get; set; }
        public Dictionary<string, PropertyInfo> Properties { get; set; }

        public int Height { get; set; }

        protected Type ObjectType { get; private set; }

#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider declaring as nullable.
        public HeaderManager(Type type)
        {
            ObjectType = type;
            BuildHeaders();
        }
#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider declaring as nullable.

        public HeaderManager(Type type, List<ColumnInfo> columns)
        {
            ObjectType = type;
            Columns = columns;
            var properties = ObjectType.GetTypeInfo().GetProperties();
            Properties = properties.ToDictionary(x => x.Name, x => x);
            BuildHeadersFromColumns();
        }

        private void BuildHeaders(int index = 1, int deep = 1)
        {
            Height = deep;
            var result = new List<ColumnInfo>();
            var properties = ObjectType.GetTypeInfo().GetProperties();
            int order = 1;
            foreach (var property in properties)
            {
                var colInfo = new ColumnInfo(property.Name);
                colInfo.Index = index;
                colInfo.Order = order++;
                result.Add(colInfo);
                if (IsNestedObject(property.PropertyType))
                {
                    colInfo.Header = new HeaderManager(property.PropertyType);
                    colInfo.Header.BuildHeaders(index, deep + 1);
                    Height = Math.Max(Height, colInfo.Header.Height);
                }
                index += colInfo.Width;

            }
            Columns = result;
            Properties = properties.ToDictionary(x => x.Name, x => x);
            
        }

        private void BuildHeadersFromColumns(int index = 1, int deep = 1)
        {
            Height = deep;
            foreach(var colInfo in Columns)
            {
                if (colInfo.IsDynamic)
                    continue;
                var property = Properties[colInfo.Name];
                if (IsNestedObject(property.PropertyType))
                {
                    colInfo.Header = new HeaderManager(property.PropertyType);
                    colInfo.Header.BuildHeaders(index, deep + 1);
                    Height = Math.Max(Height, colInfo.Header.Height);
                    index += colInfo.Width;
                }

            }
        }

        private bool IsNestedObject(Type type)
        {
            return type != typeof(string) &&
                   (type.IsClass || type.IsInterface) &&
                   type != typeof(Uri);
        }
    }

    public class HeaderManager<T> : HeaderManager
    {

        public HeaderManager() : base(typeof(T)) { }

        public HeaderManager(List<ColumnInfo> columns): base(typeof(T), columns) { }

    }
}
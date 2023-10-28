using System.Reflection;

namespace EPPLus.MultiHeader
{
    public class HeaderManager<T>
    {
        public List<ColumnInfo> Columns { get; set; }
 

        public HeaderManager() {
            Columns = BuildHeaders();
        }

        private List<ColumnInfo> BuildHeaders()
        {
            var result = new List<ColumnInfo>();
            var properties = typeof(T).GetTypeInfo().GetProperties();
            foreach (var property in properties)
            {
                result.Add(new ColumnInfo(property));
            }
            return result;
        }

    }
}
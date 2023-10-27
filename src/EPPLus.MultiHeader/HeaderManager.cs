using System.Reflection;

namespace EPPLus.MultiHeader
{
    public class HeaderManager<T>
    {
        public Dictionary<string, HeaderInfo> Columns { get; set; }
 

        public HeaderManager() {
            Columns = BuildHeaders();
        }

        private Dictionary<string, HeaderInfo> BuildHeaders()
        {
            var result = new Dictionary<string, HeaderInfo>();
            var properties = typeof(T).GetTypeInfo().GetProperties();
            foreach (var property in properties)
            {
                result.Add(property.Name, new HeaderInfo(property));
            }
            return result;
        }
    }
}
using System.Reflection;

namespace EPPLus.MultiHeader
{
    public class HeaderInfo
    {
        private string? _displayName;
        public PropertyInfo Property { get; set; }
        public string DisplayName { get => _displayName ?? Property.Name; set => _displayName = value; }

        public HeaderInfo(PropertyInfo property)
        {
            Property = property;
        }

    }
}
using System.Reflection;

namespace EPPLus.MultiHeader
{
    public class ColumnInfo
    {
        private string? _displayName;
        public PropertyInfo Property { get; set; }
        public string DisplayName { get => _displayName ?? Property.Name; set => _displayName = value; }

        public int? Order { get; set; }
        public bool Ignore { get; set; }

        public ColumnInfo(PropertyInfo property)
        {
            Property = property;
        }

    }
}
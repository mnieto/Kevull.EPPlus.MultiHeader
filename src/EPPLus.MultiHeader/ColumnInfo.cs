using System.Reflection;

namespace EPPLus.MultiHeader
{
    public class ColumnInfo
    {
        private string? _displayName;
        private int? _order;
        public PropertyInfo Property { get; set; }
        public string Name { get => Property.Name; }
        public string DisplayName { get => _displayName ?? Property.Name; set => _displayName = value; }

        public int? Order
        {
            get => _order;
            set
            {
                if (value != null && value <= 0)
                    throw new ArgumentOutOfRangeException(nameof(Order), "Value must be null or be greater or equals to 1");
                _order = value;
            }
        }

        public bool Ignore { get; set; }

        public ColumnInfo(PropertyInfo property, ColumnConfig? config = null)
        {
            Property = property;
            if (config != null)
            {
                DisplayName = config.DisplayName;
                Ignore = config.Ignore;
                Order = config.Order;
            }
        }

    }
}
namespace Kevull.EPPLus.MultiHeader
{
    /// <summary>
    /// Allows to configure the colum with
    /// </summary>
    public class ColumnWidth
    {
        /// <summary>
        /// Gets the column width behaviour
        /// </summary>
        public WidthType Type { get; private set; }

        /// <summary>
        /// Gets the column's width if <see cref="Type"/> is <see cref="WidthType.Custom"/> otherwise <c>null</c>
        /// </summary>
        public double? Width { get; private set; }

        /// <summary>
        /// Minimum column width in case the width type is <see cref="WidthType.Auto"/>. Default will be ExcelWorksheet.DefaultColWidth.
        /// </summary>
        public double MinimumWidth { get; private set; } = double.MinValue;

        /// <summary>
        /// Maximum column width in case the width type is <see cref="WidthType.Auto"/>.
        /// </summary>
        public double MaximunWidth { get; private set; } = double.MaxValue;

        /// <summary>
        /// Sets the column with type
        /// </summary>
        /// <param name="type">Column width behaviour other than <see cref="WidthType.Custom"/></param>
        /// <param name="minimumWidth">Minimum column width. This parameter is ignored if <paramref name="type"/> is not <see cref="WidthType.Auto"/></param>
        /// <param name="maximumWidth">Maximum column width. This parameter is ignored if <paramref name="type"/> is not <see cref="WidthType.Auto"/></param>
        /// <exception cref="ArgumentException">If <paramref name="type"/> is <see cref="WidthType.Custom"/> it will throw an exception</exception>
        public ColumnWidth SetWidth(WidthType type, double minimumWidth = double.MinValue, double maximumWidth = double.MaxValue)
        {
            if (type == WidthType.Custom)
                throw new ArgumentException("Use the SetWidth(int) overload to set a custom width");
            if (type == WidthType.Auto)
            {
                MinimumWidth = minimumWidth;
                MaximunWidth = maximumWidth;
            }
            Type = type;
            Width = null;
            return this;
        }

        /// <summary>
        /// Sets a custom with for the column
        /// </summary>
        /// <param name="with"></param>
        public ColumnWidth SetWidth(double with)
        {
            Type = WidthType.Custom;
            Width = with;
            return this;
        }
    }

    /// <summary>
    /// Column width behaviour
    /// </summary>
    public enum WidthType
    {
        /// <summary>Default column width: No action taken</summary>
        Default = 0,
        /// <summary><see cref="ColumnWidth.Width"/> be set</summary>
        Custom = 1,
        /// <summary>Excel autho width</summary>
        Auto = 2,
        /// <summary>Column will be hidden</summary>
        Hidden = 3
    }
}

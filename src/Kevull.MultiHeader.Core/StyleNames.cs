namespace Kevull.MultiHeader.Core
{
    /// <summary>
    /// Contains names and number formats used by default report styles.
    /// </summary>
    public class StyleNames
    {
        /// <summary>
        /// Name of the default header style.
        /// </summary>
        public const string HeaderStyleName =  "__Headers__";

        /// <summary>
        /// Name of the default date style.
        /// </summary>
        public const string DateStyleName = "__date__";

        /// <summary>
        /// Name of the default time style.
        /// </summary>
        public const string TimeStyleName = "__time__";

        /// <summary>
        /// Default number format for time values.
        /// </summary>
        /// <remarks>
        /// This format depends on local system settings.
        /// </remarks>
        public const string TimeFormat = "[$-x-systime]h:mm:ss AM/PM";

        /// <summary>
        /// Default number format for date values.
        /// </summary>
        /// <remarks>
        /// This format depends on local system settings.
        /// </remarks>
        public const string DateFormat = "mm-dd-yy";
    }
}
namespace Kevull.MultiHeader.Core
{
    /// <summary>
    /// Represents the formatting properties for an Excel cell
    /// </summary>
    public class CellFormat
    {
        /// <summary>
        /// Border style for the left edge
        /// </summary>
        public BorderStyle? LeftBorder { get; set; }

        /// <summary>
        /// Border style for the right edge
        /// </summary>
        public BorderStyle? RightBorder { get; set; }

        /// <summary>
        /// Border style for the top edge
        /// </summary>
        public BorderStyle? TopBorder { get; set; }

        /// <summary>
        /// Border style for the bottom edge
        /// </summary>
        public BorderStyle? BottomBorder { get; set; }

        /// <summary>
        /// Vertical alignment of cell content
        /// </summary>
        public VerticalAlignment? VerticalAlignment { get; set; }

        /// <summary>
        /// Horizontal alignment of cell content
        /// </summary>
        public HorizontalAlignment? HorizontalAlignment { get; set; }

        /// <summary>
        /// Background color of the cell
        /// </summary>
        public ExcelColor? BackgroundColor { get; set; }

        /// <summary>
        /// Fill style pattern
        /// </summary>
        public FillStyle? FillStyle { get; set; }

        /// <summary>
        /// Whether the font should be bold
        /// </summary>
        public bool? Bold { get; set; }

        /// <summary>
        /// Whether the font should be italic
        /// </summary>
        public bool? Italic { get; set; }

        /// <summary>
        /// Font family name
        /// </summary>
        public string? FontName { get; set; }

        /// <summary>
        /// Font size in points
        /// </summary>
        public float? FontSize { get; set; }

        /// <summary>
        /// Font color
        /// </summary>
        public ExcelColor? FontColor { get; set; }

        /// <summary>
        /// Number format string (e.g., "mm-dd-yy", "0.00")
        /// </summary>
        public string? NumberFormat { get; set; }

        /// <summary>
        /// Text rotation in degrees
        /// </summary>
        public int? TextRotation { get; set; }

        /// <summary>
        /// Whether text should wrap in the cell
        /// </summary>
        public bool? WrapText { get; set; }

        /// <summary>
        /// Merges the current format with another format.
        /// Any non-null value in <paramref name="other" /> overwrites the current value.
        /// </summary>
        /// <param name="other">Format to merge into the current instance.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="other" /> is <c>null</c>.</exception>
        public void Merge(CellFormat other)
        {
            if (other is null)
            {
                throw new ArgumentNullException(nameof(other));
            }

            LeftBorder = other.LeftBorder ?? LeftBorder;
            RightBorder = other.RightBorder ?? RightBorder;
            TopBorder = other.TopBorder ?? TopBorder;
            BottomBorder = other.BottomBorder ?? BottomBorder;
            VerticalAlignment = other.VerticalAlignment ?? VerticalAlignment;
            HorizontalAlignment = other.HorizontalAlignment ?? HorizontalAlignment;
            BackgroundColor = other.BackgroundColor ?? BackgroundColor;
            FillStyle = other.FillStyle ?? FillStyle;
            Bold = other.Bold ?? Bold;
            Italic = other.Italic ?? Italic;
            FontName = other.FontName ?? FontName;
            FontSize = other.FontSize ?? FontSize;
            FontColor = other.FontColor ?? FontColor;
            NumberFormat = other.NumberFormat ?? NumberFormat;
            TextRotation = other.TextRotation ?? TextRotation;
            WrapText = other.WrapText ?? WrapText;
        }
    }
}

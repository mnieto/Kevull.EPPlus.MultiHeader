using System.Drawing;
using OfficeOpenXml.Style;

namespace Kevull.MultiHeader.EPPlus;

internal static class ColorExtensions
{
    /// <summary>
    /// Creates a Color from a hex string representing RGB values.
    /// Supports formats: "RRGGBB", "#RRGGBB", "AARRGGBB", or "#AARRGGBB"
    /// </summary>
    /// <param name="color">The Color type (used for extension method)</param>
    /// <param name="hexString">Hex string representing RGB or ARGB values</param>
    /// <returns>A Color instance</returns>
    public static Color FromHexString(this Color color, string hexString)
    {
        if (string.IsNullOrWhiteSpace(hexString))
        {
            return Color.Empty;
        }

        // Remove # if present
        hexString = hexString.TrimStart('#');

        // Parse based on length
        return hexString.Length switch
        {
            6 => Color.FromArgb(
                Convert.ToInt32(hexString.Substring(0, 2), 16),
                Convert.ToInt32(hexString.Substring(2, 2), 16),
                Convert.ToInt32(hexString.Substring(4, 2), 16)
            ),
            8 => Color.FromArgb(
                Convert.ToInt32(hexString.Substring(0, 2), 16),
                Convert.ToInt32(hexString.Substring(2, 2), 16),
                Convert.ToInt32(hexString.Substring(4, 2), 16),
                Convert.ToInt32(hexString.Substring(6, 2), 16)
            ),
            _ => throw new ArgumentException($"Invalid hex color string: {hexString}. Expected 6 or 8 characters.", nameof(hexString))
        };
    }

    /// <summary>
    /// Converts an ExcelColor to System.Drawing.Color
    /// </summary>
    /// <param name="excelColor">The ExcelColor to convert</param>
    /// <returns>A System.Drawing.Color instance</returns>
    public static Color ToColor(this ExcelColor excelColor)
    {
        if (excelColor == null || string.IsNullOrWhiteSpace(excelColor.Rgb))
        {
            return Color.Empty;
        }

        return Color.Empty.FromHexString(excelColor.Rgb);
    }
}

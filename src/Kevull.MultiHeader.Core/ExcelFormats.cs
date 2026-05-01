using System;
using System.Collections.Generic;
using System.Text;

namespace Kevull.MultiHeader.Core
{
    /// <summary>
    /// Represents a color using ARGB values, library-agnostic
    /// </summary>
    public struct ExcelColor
    {
        /// <summary>
        /// Alpha component (transparency) of the color (0-255)
        /// </summary>
        public byte A { get; set; }

        /// <summary>
        /// Red component of the color (0-255)
        /// </summary>
        public byte R { get; set; }

        /// <summary>
        /// Green component of the color (0-255)
        /// </summary>
        public byte G { get; set; }

        /// <summary>
        /// Blue component of the color (0-255)
        /// </summary>
        public byte B { get; set; }

        /// <summary>
        /// Creates a new ExcelColor with full opacity (alpha = 255)
        /// </summary>
        /// <param name="r">Red component (0-255)</param>
        /// <param name="g">Green component (0-255)</param>
        /// <param name="b">Blue component (0-255)</param>
        public ExcelColor(byte r, byte g, byte b) : this(255, r, g, b) { }

        /// <summary>
        /// Creates a new ExcelColor with specified ARGB values
        /// </summary>
        /// <param name="a">Alpha component (0-255, where 0 is transparent and 255 is opaque)</param>
        /// <param name="r">Red component (0-255)</param>
        /// <param name="g">Green component (0-255)</param>
        /// <param name="b">Blue component (0-255)</param>
        public ExcelColor(byte a, byte r, byte g, byte b)
        {
            A = a;
            R = r;
            G = g;
            B = b;
        }

        /// <summary>
        /// Creates a color from a hex string (e.g., "#FF0000" or "FF0000")
        /// </summary>
        /// <param name="hex">Hex color string in format "#RRGGBB", "RRGGBB", "#AARRGGBB", or "AARRGGBB"</param>
        /// <exception cref="ArgumentException">Thrown when hex string is not in a valid format</exception>
        /// <exception cref="ArgumentNullException">Thrown when hex string is null or empty</exception>
        public ExcelColor(string hex)
        {
            if (string.IsNullOrEmpty(hex))
                throw new ArgumentNullException(nameof(hex));

            hex = hex.TrimStart('#');
            try
            {
                if (hex.Length == 6)
                {
                    A = 255;
                    R = Convert.ToByte(hex.Substring(0, 2), 16);
                    G = Convert.ToByte(hex.Substring(2, 2), 16);
                    B = Convert.ToByte(hex.Substring(4, 2), 16);
                    return;
                }
                else if (hex.Length == 8)
                {
                    A = Convert.ToByte(hex.Substring(0, 2), 16);
                    R = Convert.ToByte(hex.Substring(2, 2), 16);
                    G = Convert.ToByte(hex.Substring(4, 2), 16);
                    B = Convert.ToByte(hex.Substring(6, 2), 16);
                    return;
                }
            }
            catch (FormatException ex)
            {
                throw new ArgumentException($"Invalid hex color format: {hex}", nameof(hex), ex);
            }
            throw new ArgumentException($"Invalid hex color format: {hex}", nameof(hex));
        }


        /// <summary>
        /// Converts the color to a hex string (e.g., "#RRGGBB")
        /// </summary>
        public string Rgb => $"{R:X2}{G:X2}{B:X2}";

        /// <summary>
        /// Gets the color value as an ARGB hexadecimal string in the format AARRGGBB.
        /// </summary>
        public string Argb => $"{A:X2}{R:X2}{G:X2}{B:X2}";

        /// <summary>
        /// Gets a predefined black color (RGB: 0, 0, 0)
        /// </summary>
        public static ExcelColor Black => new ExcelColor(0, 0, 0);

        /// <summary>
        /// Gets a predefined white color (RGB: 255, 255, 255)
        /// </summary>
        public static ExcelColor White => new ExcelColor(255, 255, 255);

        /// <summary>
        /// Gets a predefined red color (RGB: 255, 0, 0)
        /// </summary>
        public static ExcelColor Red => new ExcelColor(255, 0, 0);

        /// <summary>
        /// Gets a predefined green color (RGB: 0, 255, 0)
        /// </summary>
        public static ExcelColor Green => new ExcelColor(0, 255, 0);

        /// <summary>
        /// Gets a predefined blue color (RGB: 0, 0, 255)
        /// </summary>
        public static ExcelColor Blue => new ExcelColor(0, 0, 255);

        /// <summary>
        /// Gets a predefined yellow color (RGB: 255, 255, 0)
        /// </summary>
        public static ExcelColor Yellow => new ExcelColor(255, 255, 0);

        /// <summary>
        /// Gets a predefined light gray color (RGB: 211, 211, 211)
        /// </summary>
        public static ExcelColor LightGray => new ExcelColor(211, 211, 211);

        /// <summary>
        /// Gets a predefined gray color (RGB: 128, 128, 128)
        /// </summary>
        public static ExcelColor Gray => new ExcelColor(128, 128, 128);

        /// <summary>
        /// Gets a predefined dark gray color (RGB: 169, 169, 169)
        /// </summary>
        public static ExcelColor DarkGray => new ExcelColor(169, 169, 169);
    }

    /// <summary>
    /// Border styles for cell edges
    /// </summary>
    public enum BorderStyle
    {
        /// <summary>
        /// No border
        /// </summary>
        None = 0,

        /// <summary>
        /// Thin border line
        /// </summary>
        Thin = 1,

        /// <summary>
        /// Medium thickness border line
        /// </summary>
        Medium = 2,

        /// <summary>
        /// Thick border line
        /// </summary>
        Thick = 3,

        /// <summary>
        /// Double border line
        /// </summary>
        Double = 4,

        /// <summary>
        /// Dotted border line
        /// </summary>
        Dotted = 5,

        /// <summary>
        /// Dashed border line
        /// </summary>
        Dashed = 6,

        /// <summary>
        /// Dash-dot pattern border line
        /// </summary>
        DashDot = 7,

        /// <summary>
        /// Dash-dot-dot pattern border line
        /// </summary>
        DashDotDot = 8
    }

    /// <summary>
    /// Vertical alignment options for cell content
    /// </summary>
    public enum VerticalAlignment
    {
        /// <summary>
        /// Align content to the top of the cell
        /// </summary>
        Top = 0,

        /// <summary>
        /// Align content to the center (middle) of the cell
        /// </summary>
        Center = 1,

        /// <summary>
        /// Align content to the bottom of the cell
        /// </summary>
        Bottom = 2,

        /// <summary>
        /// Justify content vertically with equal spacing
        /// </summary>
        Justify = 3,

        /// <summary>
        /// Distribute content vertically with equal spacing between lines
        /// </summary>
        Distributed = 4
    }

    /// <summary>
    /// Horizontal alignment options for cell content
    /// </summary>
    public enum HorizontalAlignment
    {
        /// <summary>
        /// General (default) alignment based on content type
        /// </summary>
        General = 0,

        /// <summary>
        /// Align content to the left of the cell
        /// </summary>
        Left = 1,

        /// <summary>
        /// Align content to the center of the cell
        /// </summary>
        Center = 2,

        /// <summary>
        /// Align content to the right of the cell
        /// </summary>
        Right = 3,

        /// <summary>
        /// Fill the cell by repeating the content
        /// </summary>
        Fill = 4,

        /// <summary>
        /// Justify content horizontally with equal spacing
        /// </summary>
        Justify = 5,

        /// <summary>
        /// Center content across selection (multiple cells)
        /// </summary>
        CenterContinuous = 6,

        /// <summary>
        /// Distribute content horizontally with equal spacing
        /// </summary>
        Distributed = 7
    }

    /// <summary>
    /// Fill style patterns for cell backgrounds
    /// </summary>
    public enum FillStyle
    {
        /// <summary>
        /// No fill pattern
        /// </summary>
        None = 0,

        /// <summary>
        /// Solid color fill
        /// </summary>
        Solid = 1,

        /// <summary>
        /// Dark gray pattern fill
        /// </summary>
        DarkGray = 2,

        /// <summary>
        /// Medium gray pattern fill
        /// </summary>
        MediumGray = 3,

        /// <summary>
        /// Light gray pattern fill
        /// </summary>
        LightGray = 4,

        /// <summary>
        /// 12.5% gray pattern fill
        /// </summary>
        Gray125 = 5,

        /// <summary>
        /// 6.25% gray pattern fill
        /// </summary>
        Gray0625 = 6
    }
}

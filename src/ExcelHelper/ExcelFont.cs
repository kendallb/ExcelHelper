/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System.Drawing;

namespace ExcelHelper
{
    /// <summary>
    /// Class to manage the properties for fonts portably
    /// </summary>
    public class ExcelFont
    {
        /// <summary>
        /// Excel font initialization
        /// </summary>
        /// <param name="fontName">The font name, by default Arial.</param>
        /// <param name="fontSize">The font size in points, by default 10 pt.</param>
        public ExcelFont(
            string fontName,
            float fontSize)
            : this(fontName, fontSize, ExcelFontStyle.Regular, Color.FromArgb(0, 0, 0, 0))
        {
        }

        /// <summary>
        /// Excel font initialization
        /// </summary>
        /// <param name="fontName">The font name, by default Arial.</param>
        /// <param name="fontSize">The font size in points, by default 10 pt.</param>
        /// <param name="style">The font style</param>
        public ExcelFont(
            string fontName,
            float fontSize,
            ExcelFontStyle style)
            : this(fontName, fontSize, style, Color.FromArgb(0, 0, 0, 0))
        {
        }

        /// <summary>
        /// Excel font initialization
        /// </summary>
        /// <param name="fontName">The font name, by default Arial.</param>
        /// <param name="fontSize">The font size in points, by default 10 pt.</param>
        /// <param name="style">The font style</param>
        /// <param name="color">The foreground color of the font, by default <b>Black</b>.</param>
        public ExcelFont(
            string fontName,
            float fontSize,
            ExcelFontStyle style,
            Color color)
        {
            FontName = fontName;
            FontSize = fontSize;
            Style = style;
            Color = color;
        }

        /// <summary>
        /// Gets Excel font name (font family)
        /// </summary>
        public string FontName { get; }

        /// <summary>
        /// Gets Excel font size in points
        /// </summary>
        public float FontSize { get; }

        /// <summary>
        /// Gets the bold flag of the Excel font
        /// </summary>
        public ExcelFontStyle Style { get; }

        /// <summary>
        /// Gets the color of the Excel font
        /// </summary>
        public Color Color { get; }
    }
}
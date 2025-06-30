/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;

namespace ExcelHelper
{
    /// <summary>
    /// Specifies how to align cell content horizontally within a cell.
    /// </summary>
    [Flags]
    public enum ExcelFontStyle
    {
        /// <summary>Normal text.</summary>
        Regular = 0,

        /// <summary>Bold text.</summary>
        Bold = 1,

        /// <summary>Italic text.</summary>
        Italic = 2,

        /// <summary>Underlined text.</summary>
        Underline = 4,

        /// <summary>Text with a line through the middle.</summary>
        Strikeout = 8,
    }
}
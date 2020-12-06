/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

namespace ExcelHelper
{
    /// <summary>
    /// Specifies how to align cell content horizontally within a cell.
    /// </summary>
    public enum ExcelAlignHorizontal
    {
        /// <summary>
        /// Not specified (use default).
        /// </summary>
        Undefined = -1,

        /// <summary>
        /// Align strings to the left, numbers to the right.
        /// </summary>
        General = 0,

        /// <summary>
        /// Align to cell left.
        /// </summary>
        Left = 1,

        /// <summary>
        /// Align to cell center.
        /// </summary>
        Center = 2,

        /// <summary>
        /// Align to cell right.
        /// </summary>
        Right = 3,

        /// <summary>
        /// Fill cell, repeating content as necessary.
        /// </summary>
        Fill = 4,

        /// <summary>
        /// Justify content horizontally to span the whole cell width.
        /// </summary>
        Justify = 5,
    }
}
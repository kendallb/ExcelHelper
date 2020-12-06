/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

namespace ExcelHelper
{
    /// <summary>
    /// Specifies how to align cell content vertically within a cell.
    /// </summary>
    public enum ExcelAlignVertical
    {
        /// <summary>
        /// Not specified (use default).
        /// </summary>
        Undefined = -1,

        /// <summary>
        /// Align to cell top.
        /// </summary>
        Top = 0,

        /// <summary>
        /// Align to cell center.
        /// </summary>
        Center = 1,

        /// <summary>
        /// Align to cell bottom.
        /// </summary>
        Bottom = 2,

        /// <summary>
        /// Justify content vertically to span the whole cell height.
        /// </summary>
        Justify = 3,
    }
}
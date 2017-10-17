/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 */

using System;
using System.Collections.Generic;

namespace ExcelHelper.Configuration
{
    /// <summary>
    /// Used to compare <see cref="ExcelPropertyMap"/>s.
    /// The order is by field index ascending. Any
    /// fields that don't have an index are pushed
    /// to the bottom.
    /// </summary>
    internal class ExcelPropertyMapComparer : IComparer<ExcelPropertyMap>
    {
        /// <summary>
        /// Compares two objects and returns a value indicating whether one is less than, equal to, or greater than the other.
        /// </summary>
        /// <returns>
        /// Value 
        ///                     Condition 
        ///                     Less than zero 
        ///                 <paramref name="x"/> is less than <paramref name="y"/>. 
        ///                     Zero 
        ///                 <paramref name="x"/> equals <paramref name="y"/>. 
        ///                     Greater than zero 
        ///                 <paramref name="x"/> is greater than <paramref name="y"/>. 
        /// </returns>
        /// <param name="x">The first object to compare. 
        ///                 </param><param name="y">The second object to compare. 
        ///                 </param><exception cref="T:System.ArgumentException">Neither <paramref name="x"/> nor <paramref name="y"/> implements the <see cref="T:System.IComparable"/> interface.
        ///                     -or- 
        ///                 <paramref name="x"/> and <paramref name="y"/> are of different types and neither one can handle comparisons with the other. 
        ///                 </exception><filterpriority>2</filterpriority>
        public virtual int Compare(
            object x,
            object y)
        {
            var xProperty = x as ExcelPropertyMap;
            var yProperty = y as ExcelPropertyMap;
            return Compare(xProperty, yProperty);
        }

        /// <summary>
        /// Compares two objects and returns a value indicating whether one is less than, equal to, or greater than the other.
        /// </summary>
        /// <returns>
        /// Value 
        ///                     Condition 
        ///                     Less than zero
        ///                 <paramref name="x"/> is less than <paramref name="y"/>.
        ///                     Zero
        ///                 <paramref name="x"/> equals <paramref name="y"/>.
        ///                     Greater than zero
        ///                 <paramref name="x"/> is greater than <paramref name="y"/>.
        /// </returns>
        /// <param name="x">The first object to compare.
        ///                 </param><param name="y">The second object to compare.
        ///                 </param>
        public virtual int Compare(
            ExcelPropertyMap x,
            ExcelPropertyMap y)
        {
            if (x == null) {
                throw new ArgumentNullException(nameof(x));
            }
            if (y == null) {
                throw new ArgumentNullException(nameof(y));
            }

            return x.Data.Index.CompareTo(y.Data.Index);
        }
    }
}
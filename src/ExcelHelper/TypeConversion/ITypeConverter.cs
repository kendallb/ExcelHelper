/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 */

using System;

namespace ExcelHelper.TypeConversion
{
    /// <summary>
    /// Converts objects to and from Excel compatible values. Internally Excel will store values as any
    /// of the following types:
    ///     - Null (empty cell)
    ///     - Strings
    ///     - Numeric (stored as a double)
    ///     - Boolean
    ///     - DateTime (stored as a double value)
    /// </summary>
    public interface ITypeConverter
    {
        /// <summary>
        /// True if the type converter will pass the native type through to Excel on writing, false
        /// if conversion is required.
        /// </summary>
        bool AcceptsNativeType { get; }

        /// <summary>
        /// Returns the type that we are converting from
        /// </summary>
        Type ConvertedType { get; }

        /// <summary>
        /// Converts the object to an Excel value. This is not called if Excel supports the type natively.
        /// </summary>
        /// <param name="options">The options to use when converting.</param>
        /// <param name="value">The object to convert to an Excel value.</param>
        /// <returns>The Excel value representation of the object.</returns>
        object ConvertToExcel(
            TypeConverterOptions options,
            object value);

        /// <summary>
        /// Converts an Excel value to an object.
        /// </summary>
        /// <param name="options">The options to use when converting.</param>
        /// <param name="excelValue">The Excel value to convert to an object.</param>
        /// <returns>The object created from the Excel value.</returns>
        object ConvertFromExcel(
            TypeConverterOptions options,
            object excelValue);

#if USE_C1_EXCEL
        /// <summary>
        /// Return the Excel type formatting string for the current options (null if not defined)
        /// </summary>
        /// <param name="options">The options to use when converting.</param>
        /// <returns>The Excel formatting string for the object, null to use default formatting.</returns>
        string ExcelFormatString(
            TypeConverterOptions options);
#endif
    }
}
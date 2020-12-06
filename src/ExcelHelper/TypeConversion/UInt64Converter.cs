/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using System.Globalization;

namespace ExcelHelper.TypeConversion
{
    /// <summary>
    /// Converts a UInt64 to and from an Excel value.
    /// </summary>
    public class UInt64Converter : DefaultTypeConverter
    {
        /// <summary>
        /// Constructor for the type converter
        /// </summary>
        public UInt64Converter()
            : base(true, typeof(ulong))
        {
        }

        /// <summary>
        /// Converts an Excel value to an object.
        /// </summary>
        /// <param name="options">The options to use when converting.</param>
        /// <param name="excelValue">The Excel value to convert to an object.</param>
        /// <returns>The object created from the Excel value.</returns>
        public override object ConvertFromExcel(
            TypeConverterOptions options,
            object excelValue)
        {
            if (excelValue is string text) {
                var numberStyle = options.NumberStyle ?? NumberStyles.Integer;

                if (ulong.TryParse(text, numberStyle, options.CultureInfo, out var ul)) {
                    return ul;
                }
            }
            try {
                return Convert.ToUInt64(excelValue);
            } catch (Exception e) {
                throw new ExcelTypeConverterException(ValueIsNotANumber, e);
            }
        }
    }
}
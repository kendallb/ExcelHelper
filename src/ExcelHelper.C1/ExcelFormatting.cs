/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using C1.C1Excel;
using ExcelHelper.TypeConversion;

namespace ExcelHelper
{
    /// <summary>
    /// Class to handle converting to Excel format strings during conversion
    /// </summary>
    public static class ExcelFormatting
    {
        /// <summary>
        /// Return the Excel type formatting string for the current options (null if not defined)
        /// </summary>
        /// <param name="options">The options to use when converting.</param>
        /// <param name="acceptsNativeType">True if the type converter will pass the native type through to Excel on writing, false if conversion is required.</param>
        /// <param name="convertedType">The type that we are converting from</param>
        /// <returns>The Excel formatting string for the object, null to use default formatting.</returns>
        public static string DefaultFormatString(
            TypeConverterOptions options,
            bool acceptsNativeType,
            Type convertedType)
        {
            if (acceptsNativeType) {
                if (options.NumberFormat != null) {
                    var format = XLStyle.FormatDotNetToXL(options.NumberFormat, convertedType, options.CultureInfo);
                    if (!string.IsNullOrEmpty(format)) {
                        return format;
                    }
                } else if (options.DateFormat != null) {
                    var format = XLStyle.FormatDotNetToXL(options.DateFormat, convertedType, options.CultureInfo);
                    if (!string.IsNullOrEmpty(format)) {
                        return format;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Return the Excel type formatting string for the current options (null if not defined)
        /// </summary>
        /// <param name="options">The options to use when converting.</param>
        /// <returns>The Excel formatting string for the object, null to use default formatting.</returns>
        public static string DateTimeFormatString(
            TypeConverterOptions options)
        {
            // Always use the general date format for storing dates in Excel, so they don't look like doubles to the user
            return XLStyle.FormatDotNetToXL(options.DateFormat ?? "G", typeof(DateTime), options.CultureInfo);
        }
    }
}
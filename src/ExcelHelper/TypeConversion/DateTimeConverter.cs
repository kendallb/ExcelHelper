/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using System.Globalization;
#if USE_C1_EXCEL
using C1.C1Excel;
#endif

namespace ExcelHelper.TypeConversion
{
    /// <summary>
    /// Converts a DateTime to and from an Excel value.
    /// </summary>
    public class DateTimeConverter : DefaultTypeConverter
    {
        /// <summary>
        /// Constructor for the type converter
        /// </summary>
        public DateTimeConverter()
            : base(true, typeof(DateTime))
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
            if (excelValue != null) {
                // Excel stores dates natively as doubles in the OLE Automation format for older formats
                // but it also comes in as native DateTime values when using XLSX files.
                if (excelValue.GetType() == typeof(double)) {
                    return DateTime.FromOADate((double)excelValue);
                }
                if (excelValue.GetType() == typeof(DateTime)) {
                    return excelValue;
                }

                // Try to parse the date as a string if it comes in that way
                var text = excelValue as string;
                if (text != null) {
                    if (string.IsNullOrWhiteSpace(text)) {
                        return DateTime.MinValue;
                    }

                    var formatProvider = (IFormatProvider)options.CultureInfo.GetFormat(typeof(DateTimeFormatInfo)) ?? options.CultureInfo;
                    var dateTimeStyle = options.DateTimeStyle ?? DateTimeStyles.None;
                    return DateTime.Parse(text, formatProvider, dateTimeStyle);
                }

                // Fail if we cannot parse the value
                return base.ConvertFromExcel(options, excelValue);
            }

            // Return DateTime.MinValue if the entry is null
            return DateTime.MinValue;
        }

#if USE_C1_EXCEL
        /// <summary>
        /// Return the Excel type formatting string for the current options (null if not defined)
        /// </summary>
        /// <param name="options">The options to use when converting.</param>
        /// <returns>The Excel formatting string for the object, null to use default formatting.</returns>
        public override string ExcelFormatString(
            TypeConverterOptions options)
        {
            // Always use the general date format for storing dates in Excel, so they don't look like doubles to the user
            return XLStyle.FormatDotNetToXL(options.DateFormat ?? "G", typeof(DateTime), options.CultureInfo);
        }
#endif
    }
}
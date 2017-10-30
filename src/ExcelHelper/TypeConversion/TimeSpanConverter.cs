/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;

namespace ExcelHelper.TypeConversion
{
    /// <summary>
    /// Converts a TimeSpan to and from an Excel value.
    /// </summary>
    public class TimeSpanConverter : DefaultTypeConverter
    {
        /// <summary>
        /// Constructor for the type converter
        /// </summary>
        public TimeSpanConverter()
            : base(false, typeof(TimeSpan))
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
                // ClosedXML can store time spans natively in Excel
                var type = excelValue.GetType();
                if (type == typeof(TimeSpan)) {
                    return excelValue;
                } else if (type == typeof(DateTime)) {
                    // ExcelDataReader reads TimeSpans as DateTime values
                    var dt = (DateTime)excelValue;
                    return new TimeSpan(dt.DayOfYear, dt.Hour, dt.Minute, dt.Second, dt.Millisecond);
                }

                // Try to parse the timespan as a string if it comes in that way
                var text = excelValue as string;
                if (text != null) {
                    var formatProvider = (IFormatProvider)options.CultureInfo;
                    TimeSpan span;
                    if (TimeSpan.TryParse(text, formatProvider, out span)) {
                        return span;
                    }
                }
            }

            // Fail if we cannot parse the value
            return base.ConvertFromExcel(options, excelValue);
        }
    }
}
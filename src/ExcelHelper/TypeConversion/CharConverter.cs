/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

namespace ExcelHelper.TypeConversion
{
    /// <summary>
    /// Converts a Char to and from an Excel value.
    /// </summary>
    public class CharConverter : DefaultTypeConverter
    {
        /// <summary>
        /// Constructor for the type converter
        /// </summary>
        public CharConverter()
            : base(true, typeof(char))
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
                if (char.TryParse(text.Trim(), out var c)) {
                    return c;
                }
            }
            return base.ConvertFromExcel(options, excelValue);
        }
    }
}
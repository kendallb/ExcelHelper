/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System.Collections;

namespace ExcelHelper.TypeConversion
{
    /// <summary>
    /// Throws an exception when used. This is here so that it's apparent
    /// that there is no support for IEnumerable type conversion. A custom
    /// converter will need to be created to have a field convert to and 
    /// from an IEnumerable.
    /// </summary>
    public class EnumerableConverter : DefaultTypeConverter
    {
        private const string Message = 
            "Converting IEnumerable types is not supported for a single field. " +
            "If you want to do this, create your own ITypeConverter and register " +
            "it in the TypeConverterFactory by calling AddConverter.";

        /// <summary>
        /// Constructor for the type converter
        /// </summary>
        public EnumerableConverter()
            : base(false, typeof(IEnumerable))
        {
        }

        /// <summary>
        /// Converts the object to an Excel value. This is not called if Excel supports the type natively.
        /// </summary>
        /// <param name="options">The options to use when converting.</param>
        /// <param name="value">The object to convert to an Excel value.</param>
        /// <returns>The Excel value representation of the object.</returns>
        public override object ConvertToExcel(
            TypeConverterOptions options,
            object value)
        {
            throw new ExcelTypeConverterException(Message);
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
            throw new ExcelTypeConverterException(Message);
        }
    }
}
/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
#if USE_C1_EXCEL
using C1.C1Excel;
#endif

namespace ExcelHelper.TypeConversion
{
    /// <summary>
    /// Converts an object to and from a an Excel value.
    /// </summary>
    public abstract class DefaultTypeConverter : ITypeConverter
    {
        private readonly Type _convertedType;

        /// <summary>
        /// Invalid conversion message
        /// </summary>
        protected const string ConversionCannotBePerformed = "The conversion cannot be performed.";

        /// <summary>
        /// Value is not a number conversion message
        /// </summary>
        protected const string ValueIsNotANumber = "The value is not an number.";

        /// <summary>
        /// Constructor for the default type converter base class
        /// </summary>
        /// <param name="acceptsNativeType">True if Excel handles this type natively, false if not</param>
        /// <param name="convertedType">Type that we are converting</param>
        protected DefaultTypeConverter(
            bool acceptsNativeType,
            Type convertedType)
        {
            AcceptsNativeType = acceptsNativeType;
            _convertedType = convertedType;
        }

        /// <summary>
        /// True if the type converter will pass the native type through to Excel on writing, false
        /// if conversion is required.
        /// </summary>
        public bool AcceptsNativeType { get; protected set; }

        /// <summary>
        /// Returns the type that we are converting from
        /// </summary>
        public Type ConvertedType => _convertedType;

        /// <summary>
        /// Converts the object to an Excel value. This is not called if Excel supports the type natively.
        /// </summary>
        /// <param name="options">The options to use when converting.</param>
        /// <param name="value">The object to convert to an Excel value.</param>
        /// <returns>The Excel value representation of the object.</returns>
        public virtual object ConvertToExcel(
            TypeConverterOptions options,
            object value)
        {
            return value?.ToString();
        }

        /// <summary>
        /// Converts an Excel value to an object.
        /// </summary>
        /// <param name="options">The options to use when converting.</param>
        /// <param name="excelValue">The Excel value to convert to an object.</param>
        /// <returns>The object created from the Excel value.</returns>
        public virtual object ConvertFromExcel(
            TypeConverterOptions options,
            object excelValue)
        {
            throw new ExcelTypeConverterException(ConversionCannotBePerformed);
        }

#if USE_C1_EXCEL
        /// <summary>
        /// Return the Excel type formatting string for the current options (null if not defined)
        /// </summary>
        /// <param name="options">The options to use when converting.</param>
        /// <returns>The Excel formatting string for the object, null to use default formatting.</returns>
        public virtual string ExcelFormatString(
            TypeConverterOptions options)
        {
            if (AcceptsNativeType) {
                if (options.NumberFormat != null) {
                    var format = XLStyle.FormatDotNetToXL(options.NumberFormat, _convertedType, options.CultureInfo);
                    if (!string.IsNullOrEmpty(format)) {
                        return format;
                    }
                } else if (options.DateFormat != null) {
                    var format = XLStyle.FormatDotNetToXL(options.DateFormat, _convertedType, options.CultureInfo);
                    if (!string.IsNullOrEmpty(format)) {
                        return format;
                    }
                }
            }
            return null;
        }
#endif
    }
}
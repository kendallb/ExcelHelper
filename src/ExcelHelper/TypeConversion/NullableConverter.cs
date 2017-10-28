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
    /// Converts a Nullable to and from an Excel value.
    /// </summary>
    public class NullableConverter : DefaultTypeConverter
    {
        /// <summary>
        /// Gets the type of the nullable.
        /// </summary>
        /// <value>
        /// The type of the nullable.
        /// </value>
        public Type NullableType { get; }

        /// <summary>
        /// Gets the underlying type of the nullable.
        /// </summary>
        /// <value>
        /// The underlying type.
        /// </value>
        public Type UnderlyingType { get; }

        /// <summary>
        /// Gets the type converter for the underlying type.
        /// </summary>
        /// <value>
        /// The type converter.
        /// </value>
        public ITypeConverter UnderlyingTypeConverter { get; }

        /// <summary>
        /// Creates a new <see cref="NullableConverter"/> for the given <see cref="Nullable{T}"/> <see cref="Type"/>.
        /// </summary>
        /// <param name="type">The nullable type.</param>
        /// <exception cref="System.ArgumentException">type is not a nullable type.</exception>
        public NullableConverter(
            Type type)
            : base(false, type)
        {
            NullableType = type;
            UnderlyingType = Nullable.GetUnderlyingType(type);
            if (UnderlyingType == null) {
                throw new ArgumentException("type is not a nullable type.");
            }

            // Only accept the native type if the underlying type also does
            UnderlyingTypeConverter = TypeConverterFactory.GetConverter(UnderlyingType);
            AcceptsNativeType = UnderlyingTypeConverter.AcceptsNativeType;
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
            // Excel can natively store null values, so we can write them in here
            if (value == null) {
                return null;
            }
            if (UnderlyingTypeConverter.AcceptsNativeType) {
                return value;
            }
            return UnderlyingTypeConverter.ConvertToExcel(options, value);
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
            // Handle empty cells by returning a null value
            if (excelValue == null || excelValue as string == "") {
                return null;
            }
            return UnderlyingTypeConverter.ConvertFromExcel(options, excelValue);
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
            return UnderlyingTypeConverter.ExcelFormatString(options);
        }
#endif
    }
}
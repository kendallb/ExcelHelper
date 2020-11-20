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
    /// Converts an Enum to and from an Excel value.
    /// </summary>
    public class EnumConverter : DefaultTypeConverter
    {
        private readonly Type _type;

        /// <summary>
        /// Creates a new <see cref="EnumConverter"/> for the given <see cref="Enum"/> <see cref="Type"/>.
        /// </summary>
        /// <param name="type">The type of the Enum.</param>
        public EnumConverter(
            Type type)
            : base(false, type)
        {
            if (!typeof(Enum).IsAssignableFrom(type)) {
                throw new ArgumentException($"'{type.FullName}' is not an Enum.");
            }
            _type = type;
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
            var text = excelValue as string;
            if (text != null) {
                try {
                    return Enum.Parse(_type, text, true);
                } catch (Exception e) {
                    throw new ExcelTypeConverterException($"The value is not a valid value for {_type.Name}", e);
                }
            }
            return base.ConvertFromExcel(options, excelValue);
        }
    }
}
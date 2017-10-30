/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System.Globalization;
using System.Linq;

namespace ExcelHelper.TypeConversion
{
    /// <summary>
    /// Converts a Boolean to and from an Excel value.
    /// </summary>
    public class BooleanConverter : DefaultTypeConverter
    {
        /// <summary>
        /// Constructor for the type converter
        /// </summary>
        public BooleanConverter()
            : base(false, typeof(bool))
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
            if ((bool)value) {
                var result = options.BooleanTrueValues.First();
                if (result == "1") {
                    return 1;
                }
                return result;
            } else {
                var result = options.BooleanFalseValues.First();
                if (result == "0") {
                    return 0;
                }
                return result;
            }
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
            // Return the value directly if it is already a boolean
            if (excelValue != null) {
                if (excelValue.GetType() == typeof(bool)) {
                    return excelValue;
                }
                if (excelValue.GetType() == typeof(double)) {
                    return (double)excelValue != 0.0;
                }

                var text = excelValue as string;
                if (text != null) {
                    // Try parsing the strings true/false
                    bool b;
                    if (bool.TryParse(text, out b)) {
                        return b;
                    }

                    // Try parsing as 0 or 1
                    short sh;
                    if (short.TryParse(text, out sh)) {
                        if (sh == 0) {
                            return false;
                        }
                        if (sh == 1) {
                            return true;
                        }
                    }

                    // Try parsing true values from the options (usually yes/y)
                    var t = (text ?? string.Empty).Trim();
                    foreach (var trueValue in options.BooleanTrueValues) {
                        if (options.CultureInfo.CompareInfo.Compare(trueValue, t, CompareOptions.IgnoreCase) == 0) {
                            return true;
                        }
                    }

                    // Try parsing false values from the options (usually no/n)
                    foreach (var falseValue in options.BooleanFalseValues) {
                        if (options.CultureInfo.CompareInfo.Compare(falseValue, t, CompareOptions.IgnoreCase) == 0) {
                            return false;
                        }
                    }
                }
            } else if (options.BooleanFalseValues.Contains("")) {
                // If we support blank to mean false, also convert null to false as well
                return false;
            }
            return base.ConvertFromExcel(options, excelValue);
        }
    }
}
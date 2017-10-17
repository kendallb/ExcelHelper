/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 */

using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelHelper.TypeConversion
{
    /// <summary>
    /// Options used when doing type conversion.
    /// </summary>
    public class TypeConverterOptions
    {
        private readonly List<string> _booleanTrueValues = new List<string> { "true", "yes", "y" };
        private readonly List<string> _booleanFalseValues = new List<string> { "false", "no", "n" };

        /// <summary>
        /// Gets or sets the culture info.
        /// </summary>
        public CultureInfo CultureInfo { get; set; }

        /// <summary>
        /// Gets or sets the date time style.
        /// </summary>
        public DateTimeStyles? DateTimeStyle { get; set; }

        /// <summary>
        /// Gets or sets the number style.
        /// </summary>
        public NumberStyles? NumberStyle { get; set; }

        /// <summary>
        /// Gets or sets the string format for numbers. These are Excel formatting strings. You can find a reference to them with this link:
        /// 
        /// https://support.office.com/en-us/article/Number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68
        /// </summary>
        public string NumberFormat { get; set; }

        /// <summary>
        /// Gets or sets the string format for DateTimes. These are Excel formatting strings. You can find a reference to them with this link:
        /// 
        /// https://support.office.com/en-us/article/Number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68
        /// </summary>
        public string DateFormat { get; set; }

        /// <summary>
        /// Gets the list of values that can be
        /// used to represent a boolean of true.
        /// </summary>
        public List<string> BooleanTrueValues => _booleanTrueValues;

        /// <summary>
        /// Gets the list of values that can be
        /// used to represent a boolean of false.
        /// </summary>
        public List<string> BooleanFalseValues => _booleanFalseValues;

        /// <summary>
        /// Indicates that the column should be considered a formula
        /// </summary>
        public bool IsFormula { get; set; }

        /// <summary>
        /// Merges TypeConverterOptions by applying the values of sources in order to a
        /// new TypeConverterOptions instance.
        /// </summary>
        /// <param name="sources">The sources that will be applied.</param>
        /// <returns>A new instance of TypeConverterOptions with the source applied to it.</returns>
        public static TypeConverterOptions Merge(
            params TypeConverterOptions[] sources)
        {
            var options = new TypeConverterOptions();
            foreach (var source in sources) {
                if (source == null) {
                    continue;
                }
                if (source.CultureInfo != null) {
                    options.CultureInfo = source.CultureInfo;
                }
                if (source.DateTimeStyle != null) {
                    options.DateTimeStyle = source.DateTimeStyle;
                }
                if (source.NumberStyle != null) {
                    options.NumberStyle = source.NumberStyle;
                }
                if (source.NumberFormat != null) {
                    options.NumberFormat = source.NumberFormat;
                }
                if (source.DateFormat != null) {
                    options.DateFormat = source.DateFormat;
                }
                if (!options._booleanTrueValues.SequenceEqual(source._booleanTrueValues)) {
                    options._booleanTrueValues.Clear();
                    options._booleanTrueValues.AddRange(source._booleanTrueValues);
                }
                if (!options._booleanFalseValues.SequenceEqual(source._booleanFalseValues)) {
                    options._booleanFalseValues.Clear();
                    options._booleanFalseValues.AddRange(source._booleanFalseValues);
                }
                if (source.IsFormula) {
                    options.IsFormula = true;
                }
            }
            return options;
        }
    }
}
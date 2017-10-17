/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 */

using System;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using ExcelHelper.TypeConversion;

namespace ExcelHelper.Configuration
{
    /// <summary>
    /// Mapping info for a property to a Excel field.
    /// </summary>
    [DebuggerDisplay("Names = {string.Join(\",\", Data.Names)}, Index = {Data.Index}, Ignore = {Data.Ignore}, Property = {Data.Property}, TypeConverter = {Data.TypeConverter}")]
    public class ExcelPropertyMap
    {
        private readonly ExcelPropertyMapData _data;

        /// <summary>
        /// Creates a new <see cref="ExcelPropertyMap"/> instance using the specified property.
        /// </summary>
        public ExcelPropertyMap(
            PropertyInfo property)
        {
            _data = new ExcelPropertyMapData(property);
        }

        /// <summary>
        /// Property map data.
        /// </summary>
        public ExcelPropertyMapData Data => _data;

        /// <summary>
        /// When reading, is used to get the field at the index of the name if there was a header specified. 
        /// It will look for the first name match in the order listed. When writing, sets the name of the 
        /// field in the header record. The first name will be used.
        /// </summary>
        /// <param name="names">The possible names of the Excel field.</param>
        public ExcelPropertyMap Name(
            params string[] names)
        {
            if (names == null || names.Length == 0) {
                throw new ArgumentNullException(nameof(names));
            }
            _data.Names.Clear();
            _data.Names.AddRange(names);
            return this;
        }

        /// <summary>
        /// When reading, is used to get the index of the name used when there are multiple names that are the same.
        /// </summary>
        /// <param name="index">The index of the name.</param>
        public ExcelPropertyMap NameIndex(
            int index)
        {
            _data.NameIndex = index;
            return this;
        }

        /// <summary>
        /// When reading, is used to get the field at the given index. When writing, the fields will be 
        /// written in the order of the field indexes.
        /// </summary>
        /// <param name="index">The index of the Excel field.</param>
        public ExcelPropertyMap Index(
            int index)
        {
            _data.Index = index;
            _data.IsIndexSet = true;
            return this;
        }

        /// <summary>
        /// Ignore the property when reading and writing.
        /// </summary>
        public ExcelPropertyMap Ignore()
        {
            _data.Ignore = true;
            return this;
        }

        /// <summary>
        /// Property is only used when writing and ignored when reading
        /// </summary>
        public ExcelPropertyMap WriteOnly()
        {
            _data.WriteOnly = true;
            return this;
        }

        /// <summary>
        /// Property is written but is is optional and may be missing when reading. If a property
        /// is missing, the property is left as the default value for that type.
        /// </summary>
        public ExcelPropertyMap OptionalRead()
        {
            _data.OptionalRead = true;
            return this;
        }

        /// <summary>
        /// The default value that will be used when reading when the Excel field is empty.
        /// </summary>
        /// <param name="defaultValue">The default value.</param>
        public ExcelPropertyMap Default(
            object defaultValue)
        {
            _data.Default = defaultValue;
            return this;
        }

        /// <summary>
        /// Specifies the <see cref="TypeConverter"/> to use when converting the property to and from a Excel field.
        /// </summary>
        /// <param name="typeConverter">The TypeConverter to use.</param>
        public ExcelPropertyMap TypeConverter(
            ITypeConverter typeConverter)
        {
            _data.TypeConverter = typeConverter;
            return this;
        }

        /// <summary>
        /// Specifies the <see cref="TypeConverter"/> to use when converting the property to and from a Excel field.
        /// </summary>
        /// <typeparam name="T">The <see cref="Type"/> of the 
        /// <see cref="TypeConverter"/> to use.</typeparam>
        public ExcelPropertyMap TypeConverter<T>()
            where T : ITypeConverter
        {
            TypeConverter(ReflectionHelper.CreateInstance<T>());
            return this;
        }

        /// <summary>
        /// The <see cref="CultureInfo"/> used when type converting. This will override the 
        /// global <see cref="ExcelConfiguration.CultureInfo"/> setting.
        /// </summary>
        /// <param name="cultureInfo">The culture info.</param>
        public ExcelPropertyMap Culture(
            CultureInfo cultureInfo)
        {
            _data.TypeConverterOptions.CultureInfo = cultureInfo;
            return this;
        }

        /// <summary>
        /// The <see cref="DateTimeStyles"/> to use when type converting strings to <see cref="DateTime"/>.
        /// </summary>
        /// <param name="dateTimeStyle">The date time style.</param>
        public ExcelPropertyMap DateTimeStyle(
            DateTimeStyles dateTimeStyle)
        {
            _data.TypeConverterOptions.DateTimeStyle = dateTimeStyle;
            return this;
        }

        /// <summary>
        /// The <see cref="NumberStyles"/> to use when type converting. This is used when doing any number conversions.
        /// </summary>
        /// <param name="numberStyle"></param>
        public ExcelPropertyMap NumberStyle(
            NumberStyles numberStyle)
        {
            _data.TypeConverterOptions.NumberStyle = numberStyle;
            if (numberStyle == NumberStyles.Currency) {
                // Turn on currency formatting if the currency format is requested
                _data.TypeConverterOptions.NumberFormat = "$#,##0.00;($#,##0.00)";
            }
            return this;
        }

        /// <summary>
        /// The string format to be used when type converting numbers.
        /// </summary>
        /// <param name="format">The format.</param>
        public ExcelPropertyMap NumberFormat(
            string format)
        {
            _data.TypeConverterOptions.NumberFormat = format;
            return this;
        }

        /// <summary>
        /// The string format to be used when type converting DateTimes.
        /// </summary>
        /// <param name="format">The format.</param>
        public ExcelPropertyMap DateFormat(
            string format)
        {
            _data.TypeConverterOptions.DateFormat = format;
            return this;
        }

        /// <summary>
        /// Sets the boolean type to be numeric 0 and 1 values. Values get written as 0 and 1 to the Excel file
        /// </summary>
        public ExcelPropertyMap BooleanStyleNumeric()
        {
            return BooleanStyle(true, true, "1", "true", "yes", "y").BooleanStyle(false, true, "0", "false", "no", "n");
        }

        /// <summary>
        /// Sets the boolean type to the string values "yes" and "no"
        /// </summary>
        public ExcelPropertyMap BooleanStyleYesNo()
        {
            return BooleanStyle(true, true, "yes", "true", "y").BooleanStyle(false, true, "no", "false", "n");
        }

        /// <summary>
        /// Sets the boolean type to the string values "yes" and ""
        /// </summary>
        public ExcelPropertyMap BooleanStyleYesBlank()
        {
            return BooleanStyle(true, true, "yes", "true", "y").BooleanStyle(false, true, "", "no", "false", "n");
        }

        /// <summary>
        /// The string values used to represent a boolean when converting. If you are overriding the default
        /// boolean string values, the first value in the list is used to write the resulting Excel file.
        /// </summary>
        /// <param name="isTrue">A value indicating whether true values or false values are being set.</param>
        /// <param name="booleanValues">The string boolean values.</param>
        public ExcelPropertyMap BooleanStyle(
            bool isTrue,
            params string[] booleanValues)
        {
            return BooleanStyle(isTrue, true, booleanValues);
        }

        /// <summary>
        /// The string values used to represent a boolean when converting. If you are overriding the default
        /// boolean string values, the first value in the list is used to write the resulting Excel file.
        /// </summary>
        /// <param name="isTrue">A value indicating whether true values or false values are being set.</param>
        /// <param name="clearValues">A value indication if the current values should be cleared before adding the new ones.</param>
        /// <param name="booleanValues">The string boolean values.</param>
        public ExcelPropertyMap BooleanStyle(
            bool isTrue,
            bool clearValues,
            params string[] booleanValues)
        {
            if (isTrue) {
                if (clearValues) {
                    _data.TypeConverterOptions.BooleanTrueValues.Clear();
                }
                _data.TypeConverterOptions.BooleanTrueValues.AddRange(booleanValues);
            } else {
                if (clearValues) {
                    _data.TypeConverterOptions.BooleanFalseValues.Clear();
                }
                _data.TypeConverterOptions.BooleanFalseValues.AddRange(booleanValues);
            }
            return this;
        }

        /// <summary>
        /// Tells the converter that this column is a formula
        /// </summary>
        public ExcelPropertyMap IsFormula()
        {
            _data.TypeConverterOptions.IsFormula = true;
            return this;
        }
    }
}
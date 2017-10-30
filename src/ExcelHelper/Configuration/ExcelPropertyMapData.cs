/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System.Collections.Generic;
using System.Reflection;
using ExcelHelper.TypeConversion;

namespace ExcelHelper.Configuration
{
    /// <summary>
    /// The configured data for the property map.
    /// </summary>
    public class ExcelPropertyMapData
    {
        private readonly PropertyInfo _property;
        private readonly List<string> _names;
        private bool _isDefaultSet;
        private object _defaultValue;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelPropertyMapData"/> class.
        /// </summary>
        /// <param name="property">The property.</param>
        public ExcelPropertyMapData(
            PropertyInfo property)
        {
            _property = property;
            _names = new List<string> { property.Name };
            Index = -1;
            TypeConverter = TypeConverterFactory.GetConverter(property.PropertyType);
            TypeConverterOptions = new TypeConverterOptions();
        }

        /// <summary>
        /// Gets the <see cref="PropertyInfo"/> that the data
        /// is associated with.
        /// </summary>
        public PropertyInfo Property => _property;

        /// <summary>
        /// Gets the list of column names.
        /// </summary>
        public List<string> Names => _names;

        /// <summary>
        /// Gets or sets the index of the name. This is used if there are multiple columns with the same names.
        /// </summary>
        public int NameIndex { get; set; }

        /// <summary>
        /// Gets or sets the column index.
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// Gets or sets a value indicating if the index was explicitly set. True if it was explicitly set, otherwise false.
        /// </summary>
        public bool IsIndexSet { get; set; }

        /// <summary>
        /// Gets or sets the type converter.
        /// </summary>
        public ITypeConverter TypeConverter { get; set; }

        /// <summary>
        /// Gets the type converter options.
        /// </summary>
        public TypeConverterOptions TypeConverterOptions { get; private set; }

        /// <summary>
        /// Gets or sets a value indicating whether the field should be ignored.
        /// </summary>
        public bool Ignore { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the field should be ignored on reads
        /// </summary>
        public bool WriteOnly { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the field should be can be missing on reads. If a field
        /// is missing, the field is left as the default value for that type.
        /// </summary>
        public bool OptionalRead { get; set; }

        /// <summary>
        /// Gets or sets the default value used when a Excel field is empty.
        /// </summary>
        public object Default
        {
            get { return _defaultValue; }
            set
            {
                _defaultValue = value;
                _isDefaultSet = true;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is default value set.
        /// the default value was explicitly set. True if it was explicitly set, otherwise false.
        /// </summary>
        public bool IsDefaultSet => _isDefaultSet;
    }
}
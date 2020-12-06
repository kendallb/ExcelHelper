/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using System.Globalization;

namespace ExcelHelper.Configuration
{
    /// <summary>
    /// Configuration used for reading and writing Excel data.
    /// </summary>
    public class ExcelConfiguration : IExcelConfiguration
    {
        private bool _headerIsBold = true;
        private bool _autoSizeColumns = true;
        private double _maxColumnWidth = double.MaxValue;
        private bool _willThrowOnMissingHeader = true;
        private bool _isHeaderCaseSensitive = true;
        private CultureInfo _cultureInfo = CultureInfo.CurrentCulture;
        private readonly ExcelClassMapCollection _maps = new ExcelClassMapCollection();

        /// <summary>
        /// The configured <see cref="ExcelClassMapBase"/>s.
        /// </summary>
        public ExcelClassMapCollection Maps => _maps;

        /// <summary>
        /// Gets or sets whether we should return blank strings or not. Excel stores empty
        /// cells as null values in the file, so if you set this to true empty cells will
        /// become blank strings when read into a string field, rather than being null.
        /// </summary>
        public bool ReadEmptyStrings { get; set; }

        /// <summary>
        /// Gets or sets a value indicating if the Excel file header row should be bold or not.
        /// Default is true.
        /// </summary>
        public bool HeaderIsBold
        {
            get => _headerIsBold;
            set => _headerIsBold = value;
        }

        /// <summary>
        /// Gets or sets a value indicating if the Excel file columns should be auto sized.
        /// Default is true.
        /// </summary>
        public bool AutoSizeColumns
        {
            get => _autoSizeColumns;
            set => _autoSizeColumns = value;
        }

        /// <summary>
        /// Gets or sets a value indicating the maximum column width for auto column sizing in twips
        /// </summary>
        public double MaxColumnWidth
        {
            get => _maxColumnWidth;
            set => _maxColumnWidth = value;
        }

        /// <summary>
        /// Gets or sets a value indicating if an exception will be thrown if a field defined in a mapping is missing.
        /// True to throw an exception, otherwise false. Default is true.
        /// </summary>
        public bool WillThrowOnMissingHeader
        {
            get => _willThrowOnMissingHeader;
            set => _willThrowOnMissingHeader = value;
        }

        /// <summary>
        /// Gets or sets a value indicating whether matching header column names is case sensitive. True for case sensitive
        /// matching, otherwise false. Default is true.
        /// </summary>
        public bool IsHeaderCaseSensitive
        {
            get => _isHeaderCaseSensitive;
            set => _isHeaderCaseSensitive = value;
        }

        /// <summary>
        /// Gets or sets a value indicating whether matcher header column names will ignore white space. True to ignore
        /// white space, otherwise false. Default is false.
        /// </summary>
        public bool IgnoreHeaderWhiteSpace { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether headers should be trimmed. True to trim headers,
        /// otherwise false. Default is false.
        /// </summary>
        public bool TrimHeaders { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether fields should be trimmed. True to trim fields,
        /// otherwise false. Default is false.
        /// </summary>
        public bool TrimFields { get; set; }

        /// <summary>
        /// Gets or sets the culture info used to read an write Excel files.
        /// </summary>
        public CultureInfo CultureInfo
        {
            get => _cultureInfo;
            set => _cultureInfo = value;
        }

        /// <summary>
        /// Gets or sets a value indicating if private get and set property accessor should be
        /// ignored when reading and writing. True to ignore, otherwise false. Default is false.
        /// </summary>
        public bool IgnorePrivateAccessor { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether exceptions that occur during reading should be
        /// ignored. True to ignore exceptions, otherwise false. Default is false. This is only
        /// applicable when during <see cref="IExcelReader.GetRecords{T}"/>.
        /// </summary>
        public bool IgnoreReadingExceptions { get; set; }

        /// <summary>
        /// True to ignore empty rows and move to the next record. False to finish reading when an empty
        /// row is reached. False is the default.
        /// </summary>
        public bool IgnoreEmptyRows { get; set; }

        /// <summary>
        /// Gets or sets the callback that is called when a reading exception occurs. This will only happen when
        /// <see cref="IgnoreReadingExceptions"/> is true, and when calling <see cref="IExcelReader.GetRecords{T}"/>.
        /// </summary>
        public Action<Exception, ExcelReadErrorDetails> ReadingExceptionCallback { get; set; }

        /// <summary>
        /// Use a <see cref="ExcelClassMap{T}" /> to configure mappings. When using a class map, no properties
        /// are mapped by default. Only properties specified in the mapping are used.
        /// </summary>
        /// <typeparam name="TMap">The type of mapping class to use.</typeparam>
        public void RegisterClassMap<TMap>()
            where TMap : ExcelClassMapBase
        {
            var map = ReflectionHelper.CreateInstance<TMap>();
            RegisterClassMap(map);
        }

        /// <summary>
        /// Use a <see cref="ExcelClassMap{T}" /> to configure mappings. When using a class map, no
        /// properties are mapped by default. Only properties specified in the mapping are used.
        /// </summary>
        /// <param name="classMapType">The type of mapping class to use.</param>
        public void RegisterClassMap(
            Type classMapType)
        {
            if (!typeof(ExcelClassMapBase).IsAssignableFrom(classMapType)) {
                throw new ArgumentException("The class map type must inherit from ExcelClassMap.");
            }

            var map = (ExcelClassMapBase)ReflectionHelper.CreateInstance(classMapType);
            RegisterClassMap(map);
        }

        /// <summary>
        /// Registers the class map.
        /// </summary>
        /// <param name="mapBase">The class map to register.</param>
        public void RegisterClassMap(
            ExcelClassMapBase mapBase)
        {
            if (mapBase.Constructor == null && mapBase.PropertyMaps.Count == 0 && mapBase.ReferenceMaps.Count == 0) {
                throw new ExcelConfigurationException("No mappings were specified in the ExcelClassMap.");
            }

            Maps.Add(mapBase);
        }

        /// <summary>
        /// Unregisters the class map.
        /// </summary>
        /// <typeparam name="TMap">The map type to unregister.</typeparam>
        public void UnregisterClassMap<TMap>()
            where TMap : ExcelClassMapBase
        {
            UnregisterClassMap(typeof(TMap));
        }

        /// <summary>
        /// Unregisters the class map.
        /// </summary>
        /// <param name="classMapType">The map type to unregister.</param>
        public void UnregisterClassMap(
            Type classMapType)
        {
            _maps.Remove(classMapType);
        }

        /// <summary>
        /// Unregisters all class maps.
        /// </summary>
        public void UnregisterClassMap()
        {
            _maps.Clear();
        }

        /// <summary>
        /// Generates a <see cref="ExcelClassMapBase"/> for the type.
        /// </summary>
        /// <typeparam name="T">The type to generate the map for.</typeparam>
        /// <returns>The generate map.</returns>
        public ExcelClassMapBase AutoMap<T>()
        {
            return AutoMap(typeof(T));
        }

        /// <summary>
        /// Generates a <see cref="ExcelClassMapBase"/> for the type.
        /// </summary>
        /// <param name="type">The type to generate for the map.</param>
        /// <returns>The generate map.</returns>
        public ExcelClassMapBase AutoMap(
            Type type)
        {
            var mapType = typeof(DefaultExcelClassMap<>).MakeGenericType(type);
            var map = (ExcelClassMapBase)ReflectionHelper.CreateInstance(mapType);
            map.AutoMap();
            return map;
        }
    }
}
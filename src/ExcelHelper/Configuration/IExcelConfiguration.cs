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
    /// Interface to configuration used for reading and writing Excel data.
    /// </summary>
    public interface IExcelConfiguration
    {
        /// <summary>
        /// The configured <see cref="ExcelClassMap"/>s.
        /// </summary>
        ExcelClassMapCollection Maps { get; }

        /// <summary>
        /// Gets or sets a value indicating if the Excel file header row should be bold or not.
        /// Default is true.
        /// </summary>
        bool HeaderIsBold { get; set; }

        /// <summary>
        /// Gets or sets a value indicating if the Excel file columns should be auto sized.
        /// Default is false.
        /// </summary>
        bool AutoSizeColumns { get; set; }

        /// <summary>
        /// Gets or sets a value indicating the maximum column widht for auto column sizing in twips
        /// </summary>
        double MaxColumnWidth { get; set; }

        /// <summary>
        /// Gets or sets a value indicating if an exception will be thrown if a field defined in a mapping is missing.
        /// True to throw an exception, otherwise false. Default is true.
        /// </summary>
        bool WillThrowOnMissingHeader { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether matching header column names is case sensitive. True for case sensitive
        /// matching, otherwise false. Default is true.
        /// </summary>
        bool IsHeaderCaseSensitive { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether matcher header column names will ignore white space. True to ignore
        /// white space, otherwise false. Default is false.
        /// </summary>
        bool IgnoreHeaderWhiteSpace { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether headers should be trimmed. True to trim headers,
        /// otherwise false. Default is false.
        /// </summary>
        bool TrimHeaders { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether fields should be trimmed. True to trim fields,
        /// otherwise false. Default is false.
        /// </summary>
        bool TrimFields { get; set; }

        /// <summary>
        /// Gets or sets the culture info used to read an write Excel files.
        /// </summary>
        CultureInfo CultureInfo { get; set; }

        /// <summary>
        /// Gets or sets a value indicating if private get and set property accessor should be 
        /// ignored when reading and writing. True to ignore, otherwise false. Default is false.
        /// </summary>
        bool IgnorePrivateAccessor { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether exceptions that occur during reading should be 
        /// ignored. True to ignore exceptions, otherwise false. Default is false. This is only 
        /// applicable when during <see cref="IExcelReader.GetRecords{T}"/>.
        /// </summary>
        bool IgnoreReadingExceptions { get; set; }

        /// <summary>
        /// True to ignore empty rows and move to the next record. False to finish reading when an empty
        /// row is reached. False is the default.
        /// </summary>
        bool IgnoreEmptyRows { get; set; }

        /// <summary>
        /// Gets or sets the callback that is called when a reading exception occurs. This will only happen when
        /// <see cref="IgnoreReadingExceptions"/> is true, and when calling <see cref="IExcelReader.GetRecords{T}"/>.
        /// </summary>
        Action<Exception, ExcelReadErrorDetails> ReadingExceptionCallback { get; set; }

        /// <summary>
        /// Use a <see cref="ExcelClassMap{T}" /> to configure mappings. When using a class map, no properties 
        /// are mapped by default. Only properties specified in the mapping are used.
        /// </summary>
        /// <typeparam name="TMap">The type of mapping class to use.</typeparam>
        void RegisterClassMap<TMap>()
            where TMap : ExcelClassMap;

        /// <summary>
        /// Use a <see cref="ExcelClassMap{T}" /> to configure mappings. When using a class map, no 
        /// properties are mapped by default. Only properties specified in the mapping are used.
        /// </summary>
        /// <param name="classMapType">The type of mapping class to use.</param>
        void RegisterClassMap(
            Type classMapType);

        /// <summary>
        /// Registers the class map.
        /// </summary>
        /// <param name="map">The class map to register.</param>
        void RegisterClassMap(
            ExcelClassMap map);

        /// <summary>
        /// Unregisters the class map.
        /// </summary>
        /// <typeparam name="TMap">The map type to unregister.</typeparam>
        void UnregisterClassMap<TMap>()
            where TMap : ExcelClassMap;

        /// <summary>
        /// Unregisters the class map.
        /// </summary>
        /// <param name="classMapType">The map type to unregister.</param>
        void UnregisterClassMap(
            Type classMapType);

        /// <summary>
        /// Unregisters all class maps.
        /// </summary>
        void UnregisterClassMap();

        /// <summary>
        /// Generates a <see cref="ExcelClassMap"/> for the type.
        /// </summary>
        /// <typeparam name="T">The type to generate the map for.</typeparam>
        /// <returns>The generate map.</returns>
        ExcelClassMap AutoMap<T>();

        /// <summary>
        /// Generates a <see cref="ExcelClassMap"/> for the type.
        /// </summary>
        /// <param name="type">The type to generate for the map.</param>
        /// <returns>The generate map.</returns>
        ExcelClassMap AutoMap(
            Type type);
    }
}
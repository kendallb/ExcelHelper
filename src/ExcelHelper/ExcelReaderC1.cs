/*
 * Copyright (C) 2004-2013 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

#if USE_C1_EXCEL
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Linq;
using System.Linq.Expressions;
using C1.C1Excel;
using ExcelHelper.Configuration;
using ExcelHelper.TypeConversion;

namespace ExcelHelper
{
    /// <summary>
    /// Used to read original non-XML Excel files.
    /// </summary>
    public class ExcelReaderC1 : IExcelReader
    {
        private bool _disposed;
        private C1XLBook _book;
        private XLSheet _sheet;
        private int _row;
        private int _currentIndex = -1;
        private int _columnCount;
        private readonly Dictionary<string, List<int>> _namedIndexes = new Dictionary<string, List<int>>();
        private readonly List<PropertyInfo> _importedColumns = new List<PropertyInfo>();
        private readonly Dictionary<Type, Delegate> _recordFuncs = new Dictionary<Type, Delegate>();
        private readonly ExcelConfiguration _configuration;

        /// <summary>
        /// Gets the configuration.
        /// </summary>
        public IExcelConfiguration Configuration => _configuration;

        /// <summary>
        /// Creates a new Excel stream using the given <see cref="Stream"/>.
        /// </summary>
        /// <param name="stream">The stream.</param>
        public ExcelReaderC1(
            Stream stream)
            : this(stream, new ExcelConfiguration())
        {
        }

        /// <summary>
        /// Creates a new Excel stream using the given <see cref="Stream"/> and <see cref="ExcelConfiguration"/>.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelReaderC1(
            Stream stream,
            ExcelConfiguration configuration)
        {
            if (stream == null) {
                throw new ArgumentNullException(nameof(stream));
            }

            if (configuration == null) {
                throw new ArgumentNullException(nameof(configuration));
            }
            _configuration = configuration;
            _book = new C1XLBook();
            try {
                _book.Load(stream, FileFormat.OpenXml);
            } catch {
                _book.Load(stream, FileFormat.Biff8);
            }
            ChangeSheet(0);
        }

        /// <summary>
        /// Returns the total number of columns
        /// </summary>
        public int TotalColumns => _sheet.Columns.Count;

        /// <summary>
        /// Returns the total number of rows
        /// </summary>
        public int TotalRows => _sheet.Rows.Count;

        /// <summary>
        /// Returns the total number of sheets in the Excel file
        /// </summary>
        public int TotalSheets => _book.Sheets.Count;

        /// <summary>
        /// Returns the name of the current sheet
        /// </summary>
        public string SheetName => _sheet?.Name;

        /// <summary>
        /// Changes to using the passed in sheet. Note that changing to a new sheet automatically resets the
        /// internal row counter used by GetRecords.
        /// </summary>
        /// <param name="sheet">Sheet to change to (0 to TotalSheets - 1)</param>
        /// <returns>True on success, false if the sheet is out of range</returns>
        public bool ChangeSheet(
            int sheet)
        {
            var sheets = _book.Sheets;
            if (sheet >= sheets.Count) {
                return false;
            }
            _sheet = sheets[sheet];
            _row = -1;
            return true;
        }

        /// <summary>
        /// Skip over the given number of rows. Useful for cases where the header columns are not in the first row.
        /// </summary>
        /// <param name="count">The number of rows to skip</param>
        public void SkipRows(
            int count)
        {
            _row += count;
            if (_row > TotalRows) {
                _row = TotalRows;
            }
        }

        /// <summary>
        /// Reads the next row in the Excel file to you can get the cell values using the GetCell() function
        /// </summary>
        /// <returns>True if there is another row, false if not</returns>
        public bool ReadRow()
        {
            if (_row >= TotalRows) {
                return false;
            }
            _row++;
            return true;
        }

        /// <summary>
        /// Reads a cell from the Excel file.
        /// </summary>
        /// <typeparam name="T">The type of the field.</typeparam>
        /// <param name="index">Column index to reead the value from.</param>
        /// <returns>The value from the column converted to the specific type</returns>
        public T GetColumn<T>(
            int index)
        {
            var type = typeof(T);
            var converter = TypeConverterFactory.GetConverter(type);
            var typeConverterOptions = TypeConverterOptionsFactory.GetOptions(type, _configuration.CultureInfo);
            return (T)converter.ConvertFromExcel(typeConverterOptions, _sheet[_row, index].Value);
        }

        /// <summary>
        /// Gets the raw field at position (column) index.
        /// </summary>
        /// <param name="index">The zero based index of the field.</param>
        /// <returns>The raw field.</returns>
        protected object GetField(
            int index)
        {
            // Set the current index being used so we have more information if an error occurs when reading records.
            _currentIndex = index;

            // Get the field value from the Excel file
            var field = _sheet[_row, index].Value;

            // Trim string fields if the option is set
            if (_configuration.TrimFields && field.GetType() == typeof(string)) {
                field = ((string)field).Trim();
            }
            return field;
        }

        /// <summary>
        /// Parses the named indexes from the header record.
        /// </summary>
        private void ParseHeaderRecord()
        {
            // We assume all columns contain headers by default
            _columnCount = _sheet.Columns.Count;

            // First make sure we have a header record
            if (IsEmptyRecord()) {
                throw new ExcelReaderException("No header record was found.");
            }

            // Process each column in the header row
            for (var i = 0; i < _columnCount; i++) {
                // Get the header name
                var name = _sheet[_row, i].Value as string;
                if (string.IsNullOrWhiteSpace(name)) {
                    // Header is null or empty, so we are done. This can happen if the file has more total columns
                    // in it than header rows, which can happen if some white space ends up in a right column
                    // or there are extra rows below the records. Also check that the next column is
                    // also empty and throw and exception if not.
                    if (i + 1 < _columnCount && !string.IsNullOrWhiteSpace(_sheet[_row, i + 1].Value as string)) {
                        throw new Exception($"Blank column header found at column {ColumnHelper.ColumnIndexToColumnLetter(i + 1)}");
                    }
                    _columnCount = i;
                    break;
                }

                // Now store the named index for later for this header column
                if (!_configuration.IsHeaderCaseSensitive) {
                    name = name.ToLower();
                }
                if (_namedIndexes.ContainsKey(name)) {
                    _namedIndexes[name].Add(i);
                } else {
                    _namedIndexes[name] = new List<int> { i };
                }
            }

            // Move to the next row
            _row++;
        }

        /// <summary>
        /// Determines if the record at the current line is empty or not
        /// </summary>
        /// <returns>True if record is empty, false if not</returns>
        private bool IsEmptyRecord()
        {
            for (var i = 0; i < _columnCount; i++) {
                var o = _sheet[_row, i].Value;
                if (o != null) {
                    if (o.GetType() == typeof(string)) {
                        // Make sure string fields are not empty strings
                        if (!string.IsNullOrWhiteSpace((string)o)) {
                            return false;
                        }
                    } else {
                        // Non-null, non-string fields are not empty
                        return false;
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// Gets all the records in the Excel file and converts each to <see cref="Type"/> T.
        /// </summary>
        /// <typeparam name="T">The <see cref="Type"/> of the record.</typeparam>
        /// <returns>An <see cref="IEnumerable{T}" /> of records.</returns>
        public IEnumerable<T> GetRecords<T>()
        {
            // Get the type of all the records
            var type = typeof(T);

            // Make sure it is mapped
            if (_configuration.Maps[type] == null) {
                _configuration.Maps.Add(_configuration.AutoMap(type));
            }

            // First read the header record and parse it
            ReadRow();
            ParseHeaderRecord();

            // Create the function to read the records outside the inner loop
            Delegate parseRecord;
            try {
                parseRecord = GetParseRecordFunc(type);
            } catch (Exception ex) {
                ExceptionHelper.AddExceptionDataMessage(ex, type);
                throw;
            }

            // Read each record one at a time and yield it
            var ignoreEmptyRows = Configuration.IgnoreEmptyRows;
            while (_row < TotalRows) {
                // If we are not ignoring empty records, bail when we reach one. Otherwise we process all rows in the file.
                if (IsEmptyRecord()) {
                    if (ignoreEmptyRows) {
                        _row++;
                        continue;
                    } else {
                        break;
                    }
                }

                // Now process the record
                T record;
                try {
                    _currentIndex = -1;
                    record = (T)parseRecord.DynamicInvoke();
                    _row++;
                } catch (Exception ex) {
                    // Build the details about the error so it can be logged
                    var details = new ExcelReadErrorDetails {
                        Row = _row + 1,                                 // Show 1 based row to the user
                        Column = _currentIndex + 1,                     // Show 1 based column to the user
                        FieldName = (from pair in _namedIndexes
                                     from index in pair.Value
                                     where index == _currentIndex
                                     select pair.Key).SingleOrDefault(),
                        FieldValue = _sheet[_row, _currentIndex].Value,
                    };

                    // Add the details to the exception
                    ExceptionHelper.AddExceptionDataMessage(ex, type, details);

                    // If we are ignoring errors, optionally call the callback and continue
                    if (_configuration.IgnoreReadingExceptions) {
                        _configuration.ReadingExceptionCallback?.Invoke(ex, details);
                        _row++;
                        continue;
                    }
                    throw;
                }
                yield return record;
            }
        }

        /// <summary>
        /// Gets all the records in the Excel file and converts each to dictionary of strings to strings.
        /// </summary>
        /// <returns>An enumeration of dictionaries.</returns>
        public IEnumerable<Dictionary<string, string>> GetRecordsAsDictionary()
        {
            // We assume all columns contain headers by default
            _columnCount = _sheet.Columns.Count;

            // First make sure we have a header record
            ReadRow();
            if (IsEmptyRecord()) {
                throw new ExcelReaderException("No header record was found.");
            }

            // Process each column in the header row
            var headers = new List<string>();
            for (var i = 0; i < _columnCount; i++) {
                // Get the header name
                var name = _sheet[_row, i].Value as string;
                if (string.IsNullOrWhiteSpace(name)) {
                    // Header is null or empty, so we are done. This can happen if the file has more total columns
                    // in it than header rows, which can happen if some white space ends up in a right column
                    // or there are extra rows below the records. Also check that the next column is
                    // also empty and throw and exception if not.
                    if (i + 1 < _columnCount && !string.IsNullOrWhiteSpace(_sheet[_row, i + 1].Value as string)) {
                        throw new Exception($"Blank column header found at column {ColumnHelper.ColumnIndexToColumnLetter(i + 1)}");
                    }
                    _columnCount = i;
                    break;
                }

                // Now store the named index for later for this header column
                if (!_configuration.IsHeaderCaseSensitive) {
                    name = name.ToLower();
                }
                headers.Add(name.Trim());
            }

            // Move to the next row
            _row++;

            // Read each record one at a time and yield it
            while (!IsEmptyRecord()) {
                var record = new Dictionary<string, string>();
                for (var i = 0; i < _columnCount; i++) {
                    try {
                        var cell = _sheet[_row, i];
                        var value = cell.Value;
                        if (value != null) {
                            // Format value if cell has a style with format set
                            string text;
                            var style = cell.Style ?? (_sheet.Rows[_row].Style ?? _sheet.Columns[i].Style);
                            if (value.GetType() == typeof(bool)) {
                                // By default Excel formats boolean as uppercase
                                if ((bool)value) {
                                    text = "TRUE";
                                } else {
                                    text = "FALSE";
                                }
                            } else {
                                if (style != null && style.Format.Length > 0 && value is IFormattable) {
                                    var fmt = XLStyle.FormatXLToDotNet(style.Format.ToUpperInvariant());
                                    if (!string.IsNullOrWhiteSpace(fmt)) {
                                        text = ((IFormattable)value).ToString(fmt, CultureInfo.CurrentCulture);
                                    } else {
                                        text = value.ToString();
                                    }
                                } else {
                                    text = value.ToString();
                                }
                            }
                            record.Add(headers[i], text);
                        } else {
                            // Always return empty strings, not nulls
                            record.Add(headers[i], "");
                        }
                    } catch (Exception ex) {
                        // Build the details about the error so it can be logged
                        var details = new ExcelReadErrorDetails {
                            Row = _row + 1,
                            Column = i + 1,
                            FieldName = headers[i],
                            FieldValue = _sheet[_row, i].Value,
                        };

                        // Add the details to the exception
                        ExceptionHelper.AddExceptionDataMessage(ex, null, details);

                        // If we are ignoring errors, optionally call the callback and continue
                        if (_configuration.IgnoreReadingExceptions) {
                            _configuration.ReadingExceptionCallback?.Invoke(ex, details);
                            _row++;
                            continue;
                        }
                        throw;
                    }
                }
                _row++;
                yield return record;
            }
        }

        /// <summary>
        /// Gets a list of all the properties for columns that are found in the import. This can only be called
        /// after first calling GetRecords()
        /// </summary>
        /// <returns>List of properties for columns found in the Excel file.</returns>
        public List<PropertyInfo> GetImportedColumns()
        {
            return _importedColumns;
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        /// <filterpriority>2</filterpriority>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        /// <param name="disposing">True if the instance needs to be disposed of.</param>
        protected virtual void Dispose(
            bool disposing)
        {
            if (_disposed) {
                return;
            }
            if (disposing) {
                _book?.Dispose();
                _book = null;
                _sheet = null;
            }
            _disposed = true;
        }

        /// <summary>
        /// Gets the index of the field at name if found.
        /// </summary>
        /// <param name="optionalRead">True if the field is optional on read</param>
        /// <param name="names">The possible names of the field to get the index for.</param>
        /// <param name="index">The index of the field if there are multiple fields with the same name.</param>
        /// <returns>The index of the field if found, otherwise -1.</returns>
        private int GetFieldIndex(
            bool optionalRead,
            string[] names,
            int index)
        {
            var compareOptions = !_configuration.IsHeaderCaseSensitive ? CompareOptions.IgnoreCase : CompareOptions.None;
            string name = null;
            foreach (var pair in _namedIndexes) {
                // Get the header name we will match against
                var namedIndex = pair.Key;
                if (_configuration.IgnoreHeaderWhiteSpace) {
                    namedIndex = Regex.Replace(namedIndex, "\\s", string.Empty);
                } else if (_configuration.TrimHeaders) {
                    namedIndex = namedIndex.Trim();
                }

                // Check if this index matches one of the names passed in
                foreach (var n in names) {
                    if (_configuration.CultureInfo.CompareInfo.Compare(namedIndex, n, compareOptions) == 0) {
                        name = pair.Key;
                        break;
                    }
                }
                if (name != null) {
                    break;
                }
            }

            // Handle the situation where we could not map this field. We may want to allow for record fields to be missing
            // so they will end up with the default values rather than requiring all the fields to be present
            if (name == null) {
                if (!optionalRead && _configuration.WillThrowOnMissingHeader) {
                    // If we're in strict reading mode or the field is not optional and the named index isn't found, throw an exception.
                    var namesJoined = $"'{string.Join("', '", names)}'";
                    if (names.Length > 1) {
                        throw new ExcelMissingFieldException($"Fields {namesJoined} do not exist in the Excel file.");
                    } else {
                        throw new ExcelMissingFieldException($"Field {namesJoined} does not exist in the Excel file.");
                    }
                }
                return -1;
            }

            // Found the field index, so return it's offset
            return _namedIndexes[name][index];
        }

        /// <summary>
        /// Gets the function delegate used to populate a custom class object with data from the reader.
        /// </summary>
        /// <param name="type">The type of object that is created and populated.</param>
        /// <returns>The function delegate.</returns>
        private Delegate GetParseRecordFunc(
            Type type)
        {
            if (!_recordFuncs.ContainsKey(type)) {
                // Build binding functions for all the properties in the record
                var bindings = new List<MemberBinding>();
                CreatePropertyBindingsForMapping(_configuration.Maps[type], bindings);
                if (bindings.Count == 0) {
                    throw new ExcelReaderException($"No properties are mapped for type '{type.FullName}'.");
                }

                // Build the expression to construct the class and execute all the bindings and compile it
                var constructorExpression = _configuration.Maps[type].Constructor ?? Expression.New(type);
                var body = Expression.MemberInit(constructorExpression, bindings);
                var funcType = typeof(Func<>).MakeGenericType(type);
                _recordFuncs[type] = Expression.Lambda(funcType, body).Compile();
            }
            return _recordFuncs[type];
        }

        /// <summary>
        /// Creates the property bindings for the given <see cref="ExcelClassMap"/>.
        /// </summary>
        /// <param name="mapping">The mapping to create the bindings for.</param>
        /// <param name="bindings">The bindings that will be added to from the mapping.</param>
        private void CreatePropertyBindingsForMapping(
            ExcelClassMap mapping,
            List<MemberBinding> bindings)
        {
            // First bind all the regular properties for this record
            AddPropertyBindings(mapping.PropertyMaps, bindings);

            // Now process each reference map to map embedded classes
            foreach (var referenceMap in mapping.ReferenceMaps) {
                // Ignore any maps we cannot read
                if (!CanRead(referenceMap)) {
                    continue;
                }

                // Now map all the properties in this reference map
                var referenceBindings = new List<MemberBinding>();
                CreatePropertyBindingsForMapping(referenceMap.Mapping, referenceBindings);
                var referenceBody = Expression.MemberInit(Expression.New(referenceMap.Property.PropertyType), referenceBindings);
                bindings.Add(Expression.Bind(referenceMap.Property, referenceBody));
            }
        }

        /// <summary>
        /// Adds a <see cref="MemberBinding"/> for each property for it's field.
        /// </summary>
        /// <param name="properties">The properties to add bindings for.</param>
        /// <param name="bindings">The bindings that will be added to from the properties.</param>
        private void AddPropertyBindings(
            ExcelPropertyMapCollection properties,
            List<MemberBinding> bindings)
        {
            foreach (var propertyMap in properties) {
                // Ignore properties that are not read
                if (!CanRead(propertyMap)) {
                    continue;
                }

                // Find the index of this field in the row
                var index = -1;
                var data = propertyMap.Data;
                if (data.IsIndexSet) {
                    // If an index was explicitly set, use it.
                    index = data.Index;
                } else {
                    // Fallback to the default name.
                    index = GetFieldIndex(data.OptionalRead, data.Names.ToArray(), data.NameIndex);
                }

                // Skip if the index was not found. This can happen if not all fields are included in the
                // import file, and we are not in strict reading mode or the field was marked as optional read.
                // Very useful if you want missing fields to be imported with default values. The optional read mode
                // is useful to make sure critical fields are always present.
                if (index == -1) {
                    continue;
                }

                // Get the field using the field index
                var method = GetType().GetMethod("GetField", BindingFlags.NonPublic | BindingFlags.Instance);
                Expression fieldExpression = Expression.Call(Expression.Constant(this), method, Expression.Constant(index, typeof(int)));

                // Get the type conversion information we need
                var property = data.Property;
                var propertyType = property.PropertyType;
                var typeConverterExpression = Expression.Constant(data.TypeConverter);
                var typeConverterOptions = TypeConverterOptions.Merge(
                    TypeConverterOptionsFactory.GetOptions(propertyType, _configuration.CultureInfo),
                    data.TypeConverterOptions);
                var typeConverterOptionsExpression = Expression.Constant(typeConverterOptions);

                // Store the mapped property in our list of properties
                _importedColumns.Add(property);

                // If a default value is set, check for an empty record and set the field to the default if it is empty
                Expression expression;
                if (data.IsDefaultSet) {
                    // Creating an expression to hold the local field variable
                    var field = Expression.Parameter(typeof(object), "field");

                    // Handle strings differently, so we can compare strings to the empty string as well as null
                    Expression checkFieldEmptyExpression = propertyType == typeof(string) ?
                        checkFieldEmptyExpression = Expression.Call(typeof(string), "IsNullOrEmpty", null, Expression.Convert(field, typeof(string))) :
                        checkFieldEmptyExpression = Expression.Equal(field, Expression.Constant(null));

                    // Expression to assign the default value
                    var defaultValueExpression = Expression.Assign(field, Expression.Convert(Expression.Constant(data.Default), typeof(object)));

                    // Expression to convert the field value and store it back in the variable
                    var convertExpression = Expression.Assign(field, Expression.Call(typeConverterExpression, "ConvertFromExcel", null, typeConverterOptionsExpression, field));

                    // Create a block to execute so GetField won't be called twice
                    expression = Expression.Block(
                        // Local variable
                        new[] { field },

                        // Assign the result of GetField() to a local variable
                        Expression.Assign(field, fieldExpression),

                        // Conditionally set the field to the default value, or the converted value
                        Expression.IfThenElse(checkFieldEmptyExpression, defaultValueExpression, convertExpression),

                        // Finally convert the field local variable and return it
                        Expression.Convert(field, propertyType));
                } else {
                    // Convert the field from Excel format to the native type directly
                    expression = Expression.Convert(
                        Expression.Call(typeConverterExpression, "ConvertFromExcel", null, typeConverterOptionsExpression, fieldExpression),
                        propertyType);
                }

                // Now add the binding to bind the expression result to the property
                bindings.Add(Expression.Bind(property, expression));
            }
        }

        /// <summary>
        /// Determines if the property for the <see cref="ExcelPropertyMap"/>
        /// can be read.
        /// </summary>
        /// <param name="propertyMap">The property map.</param>
        /// <returns>A value indicating of the property can be read. True if it can, otherwise false.</returns>
        private bool CanRead(
            ExcelPropertyMap propertyMap)
        {
            var cantRead =
                // Write only properties.
                propertyMap.Data.WriteOnly ||
                // Ignored properties.
                propertyMap.Data.Ignore ||
                // Properties that don't have a public setter and we are honoring the accessor modifier.
                propertyMap.Data.Property.GetSetMethod() == null && !_configuration.IgnorePrivateAccessor ||
                // Properties that don't have a setter at all.
                propertyMap.Data.Property.GetSetMethod(true) == null;
            return !cantRead;
        }

        /// <summary>
        /// Determines if the property for the <see cref="ExcelPropertyReferenceMap"/>
        /// can be read.
        /// </summary>
        /// <param name="propertyReferenceMap">The reference map.</param>
        /// <returns>A value indicating of the property can be read. True if it can, otherwise false.</returns>
        private bool CanRead(
            ExcelPropertyReferenceMap propertyReferenceMap)
        {
            var cantRead =
                // Properties that don't have a public setter and we are honoring the accessor modifier.
                propertyReferenceMap.Property.GetSetMethod() == null && !_configuration.IgnorePrivateAccessor ||
                // Properties that don't have a setter at all.
                propertyReferenceMap.Property.GetSetMethod(true) == null;
            return !cantRead;
        }
    }
}
#endif
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
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using C1.C1Excel;
using ExcelHelper.Configuration;
using ExcelHelper.TypeConversion;

namespace ExcelHelper
{
    /// <summary>
    /// Used to write Excel files.
    /// </summary>
    public class ExcelWriterC1 : IExcelWriter
    {
        private bool _disposed;
        private readonly Stream _stream;
        private C1XLBook _book;
        private XLSheet _sheet;
        private Graphics _graphics;
        private int _row;
        private int _col;
        private float[] _colWidths;
        private readonly Dictionary<Type, Delegate> _typeActions = new Dictionary<Type, Delegate>();
        private readonly ExcelConfiguration _configuration;

        /// <summary>
        /// Gets the configuration.
        /// </summary>
        public IExcelConfiguration Configuration => _configuration;

        /// <summary>
        /// Creates a new Excel writer using the given <see cref="Stream"/> and
        /// a default <see cref="ExcelConfiguration"/>.
        /// </summary>
        /// <param name="stream">The <see cref="Stream"/> used to write the Excel file.</param>
        public ExcelWriterC1(
            Stream stream)
            : this(stream, new ExcelConfiguration())
        {
        }

        /// <summary>
        /// Creates a new Excel writer using the given <see cref="Stream"/> and <see cref="ExcelConfiguration"/>.
        /// </summary>
        /// <param name="stream">The <see cref="Stream"/> used to write the Excel file.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelWriterC1(
            Stream stream,
            ExcelConfiguration configuration)
        {
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            _stream = stream ?? throw new ArgumentNullException(nameof(stream));
            _book = new C1XLBook();
            _book.CompatibilityMode = CompatibilityMode.Excel2007;
            ChangeSheet(0);
            if (_configuration.AutoSizeColumns) {
                _graphics = Graphics.FromHwnd(IntPtr.Zero);
            }

            // Set the default font to Calibri 11, which is the default in newer versions of office
            DefaultFont = new Font("Calibri", 11, FontStyle.Regular);
        }

        /// <summary>
        /// Gets or sets the default font for the Excel file
        /// </summary>
        public Font DefaultFont
        {
            get => _book.DefaultFont;
            set => _book.DefaultFont = value;
        }

        /// <summary>
        /// Update an Excel cell style with the passed in attributes
        /// </summary>
        /// <param name="cellStyle">Place to store the cell style</param>
        /// <param name="format">Optional formatting string for the cell</param>
        /// <param name="fontStyle">Optional font style for the cell</param>
        /// <param name="fontSize">Optional font size for the cell</param>
        /// <param name="fontName">Optional font name for the cell</param>
        /// <param name="horizontalAlign">Optional horizontal alignment</param>
        /// <param name="verticalAlign">Optional vertical alignment</param>
        private void UpdateStyle(
            ref XLStyle cellStyle,
            string format,
            FontStyle? fontStyle,
            float? fontSize,
            string fontName,
            ExcelAlignHorizontal? horizontalAlign,
            ExcelAlignVertical? verticalAlign)
        {
            // Set up native formatting if provided
            if (!string.IsNullOrEmpty(format)) {
                if (cellStyle == null) {
                    cellStyle = new XLStyle(_book);
                }
                cellStyle.Format = format;
            }

            // Apply font styling if defined
            if (fontStyle != null || fontSize != null || fontName != null) {
                if (cellStyle == null) {
                    cellStyle = new XLStyle(_book);
                }
                var defaultFont = _book.DefaultFont;
                var name = fontName ?? defaultFont.Name;
                var size = fontSize ?? defaultFont.SizeInPoints;
                var style = fontStyle ?? defaultFont.Style;
                cellStyle.Font = new Font(name, size, style);
            }

            // Apply the horizontal alignment if defined
            if (horizontalAlign != null) {
                if (cellStyle == null) {
                    cellStyle = new XLStyle(_book);
                }
                switch (horizontalAlign) {
                    case ExcelAlignHorizontal.General:
                        cellStyle.AlignHorz = XLAlignHorzEnum.General;
                        break;
                    case ExcelAlignHorizontal.Left:
                        cellStyle.AlignHorz = XLAlignHorzEnum.Left;
                        break;
                    case ExcelAlignHorizontal.Center:
                        cellStyle.AlignHorz = XLAlignHorzEnum.Center;
                        break;
                    case ExcelAlignHorizontal.Right:
                        cellStyle.AlignHorz = XLAlignHorzEnum.Right;
                        break;
                    case ExcelAlignHorizontal.Fill:
                        cellStyle.AlignHorz = XLAlignHorzEnum.Fill;
                        break;
                    case ExcelAlignHorizontal.Justify:
                        cellStyle.AlignHorz = XLAlignHorzEnum.Justify;
                        break;
                }
            }

            // Apply the vertical alignment if defined
            if (verticalAlign !=  null) {
                if (cellStyle == null) {
                    cellStyle = new XLStyle(_book);
                }
                switch (verticalAlign) {
                    case ExcelAlignVertical.Top:
                        cellStyle.AlignVert = XLAlignVertEnum.Top;
                        break;
                    case ExcelAlignVertical.Center:
                        cellStyle.AlignVert = XLAlignVertEnum.Center;
                        break;
                    case ExcelAlignVertical.Bottom:
                        cellStyle.AlignVert = XLAlignVertEnum.Bottom;
                        break;
                    case ExcelAlignVertical.Justify:
                        cellStyle.AlignVert = XLAlignVertEnum.Justify;
                        break;
                }
            }
        }

        /// <summary>
        /// Changes to using the passed in sheet. Note that changing to a new sheet automatically resets the
        /// internal row and column counter used by WriteRecords.
        /// </summary>
        /// <param name="sheet">Sheet to change to</param>
        public void ChangeSheet(
            int sheet)
        {
            // Perform any column resizing for the current sheet before we change it
            PerformColumnResize();

            // Insert all the sheets up to the index we need if the count is less
            var sheets = _book.Sheets;
            if (sheet >= sheets.Count) {
                for (var i = sheets.Count; i <= sheet; i++) {
                    sheets.Insert(i);
                }
            }
            _sheet = sheets[sheet];
            _row = _col = 0;
        }

        /// <summary>
        /// Writes a cell to the Excel file.
        /// </summary>
        /// <typeparam name="T">The type of the field.</typeparam>
        /// <param name="row">Row to write the field to.</param>
        /// <param name="col">Column to write the field to.</param>
        /// <param name="field">The field to write.</param>
        /// <param name="numberFormat">Optional number formatting string for the cell</param>
        /// <param name="dateFormat">Optional DateTime formatting string for the cell</param>
        /// <param name="fontStyle">Optional font style for the cell</param>
        /// <param name="fontSize">Optional font size for the cell</param>
        /// <param name="fontName">Optional font name for the cell</param>
        /// <param name="horizontalAlign">Optional horizontal alignment</param>
        /// <param name="verticalAlign">Optional vertical alignment</param>
        public void WriteCell<T>(
            int row,
            int col,
            T field,
            string numberFormat = null,
            string dateFormat = null,
            FontStyle? fontStyle = null,
            float? fontSize = null,
            string fontName = null,
            ExcelAlignHorizontal? horizontalAlign = ExcelAlignHorizontal.Undefined,
            ExcelAlignVertical? verticalAlign = ExcelAlignVertical.Undefined)
        {
            // Find the type conversion options
            var type = typeof(T);
            var converter = TypeConverterFactory.GetConverter(type);
            var options = TypeConverterOptions.Merge(TypeConverterOptionsFactory.GetOptions(type, _configuration.CultureInfo));

            // Set the formatting options to override the defaults
            var format = numberFormat ?? dateFormat;
            if (converter.AcceptsNativeType) {
                // Convert the options to Excel format
                if (format != null) {
                    format = XLStyle.FormatDotNetToXL(format, converter.ConvertedType, options.CultureInfo);
                } else {
                    // If no formatting is provided, see if the native type requires it (mostly for DateTime)
                    format = converter.ExcelFormatString(options);
                }
            } else {
                // Override the formatting for the formatter, and do not format the Excel cell
                if (numberFormat != null) {
                    options.NumberFormat = format;
                } else if (dateFormat != null) {
                    options.DateFormat = format;
                }
                format = null;
            }

            // Find the default style to use for this cell based on the row and column styles
            var cellStyle = _sheet.Rows[row].Style ?? _sheet.Columns[col].Style;

            // Clone the style so it does not modify the entire row or column
            cellStyle = cellStyle?.Clone();

            // Set up cell formatting for this cell
            UpdateStyle(ref cellStyle, format, fontStyle, fontSize, fontName, horizontalAlign, verticalAlign);

            // Apply the style to this cell if defined
            if (cellStyle != null) {
                _sheet[row, col].Style = cellStyle;
            }

            // Now write the cell contents
            object value;
            if (converter.AcceptsNativeType) {
                value = field;
            } else {
                value = converter.ConvertToExcel(options, field);
            }

            if (value is string s && s.StartsWith("=")) {
                // Write as a formula if it starts with an equals sign
                _sheet[row, col].Value = "";
                _sheet[row, col].Formula = s;
            } else {
                _sheet[row, col].Value = value;
            }
        }

        /// <summary>
        /// Set an entire column to a specific format. By default Excel defines the
        /// cell style in the following order; cell, row, column, worksheet default
        /// </summary>
        /// <param name="col">Column to set the format for</param>
        /// <param name="numberFormat">Optional number formatting string for the cell</param>
        /// <param name="dateFormat">Optional DateTime formatting string for the cell</param>
        /// <param name="fontStyle">Optional font style for the cell</param>
        /// <param name="fontSize">Optional font size for the cell</param>
        /// <param name="fontName">Optional font name for the cell</param>
        /// <param name="horizontalAlign">Optional horizontal alignment</param>
        /// <param name="verticalAlign">Optional vertical alignment</param>
        public void SetColumnFormat(
            int col,
            string numberFormat = null,
            string dateFormat = null,
            FontStyle? fontStyle = null,
            float? fontSize = null,
            string fontName = null,
            ExcelAlignHorizontal? horizontalAlign = ExcelAlignHorizontal.Undefined,
            ExcelAlignVertical? verticalAlign = ExcelAlignVertical.Undefined)
        {
            var format = XLStyle.FormatDotNetToXL(numberFormat ?? dateFormat);
            XLStyle style = null;
            UpdateStyle(ref style, format, fontStyle, fontSize, fontName, horizontalAlign, verticalAlign);
            _sheet.Columns[col].Style = style;
        }

        /// <summary>
        /// Set an entire row to a specific format. By default Excel defines the
        /// cell style in the following order; cell, row, column, worksheet default
        /// </summary>
        /// <param name="row">Row to set the format for</param>
        /// <param name="numberFormat">Optional number formatting string for the cell</param>
        /// <param name="dateFormat">Optional DateTime formatting string for the cell</param>
        /// <param name="fontStyle">Optional font style for the cell</param>
        /// <param name="fontSize">Optional font size for the cell</param>
        /// <param name="fontName">Optional font name for the cell</param>
        /// <param name="horizontalAlign">Optional horizontal alignment</param>
        /// <param name="verticalAlign">Optional vertical alignment</param>
        public void SetRowFormat(
            int row,
            string numberFormat = null,
            string dateFormat = null,
            FontStyle? fontStyle = null,
            float? fontSize = null,
            string fontName = null,
            ExcelAlignHorizontal? horizontalAlign = ExcelAlignHorizontal.Undefined,
            ExcelAlignVertical? verticalAlign = ExcelAlignVertical.Undefined)
        {
            var format = XLStyle.FormatDotNetToXL(numberFormat ?? dateFormat);
            XLStyle style = null;
            UpdateStyle(ref style, format, fontStyle, fontSize, fontName, horizontalAlign, verticalAlign);
            _sheet.Rows[row].Style = style;
        }

        /// <summary>
        /// Auto size the columns based on a specific row. If any columns in this row
        /// are wider than the current maximum, the column sizes are increased.
        /// </summary>
        /// <param name="row">Row to auto size the columns for</param>
        /// <param name="maxWidth">Maximum allowed width</param>
        private void AutoSizeColumnsForRow(
            int row,
            float maxWidth)
        {
            // Allocate or resize the column sizes based on the column count
            var count = _sheet.Columns.Count;
            if (_colWidths == null) {
                _colWidths = new float[count];
            } else if (_colWidths.Length < count) {
                Array.Resize(ref _colWidths, count);
            }

            // Now process each column in turn and measure it
            for (var i = 0; i < count; i++) {
                var cell = _sheet[row, i];
                var value = cell.Value;
                if (value != null) {
                    // Format value if cell has a style with format set
                    string text;
                    var style = cell.Style ?? (_sheet.Rows[row].Style ?? _sheet.Columns[i].Style);

                    if (value.GetType() == typeof(bool)) {
                        // By default Excel formats boolean as uppercase
                        if ((bool)value) {
                            text = "TRUE";
                        } else {
                            text = "FALSE";
                        }
                    } else {
                        if (style != null && style.Format.Length > 0 && value is IFormattable formattable) {
                            var fmt = XLStyle.FormatXLToDotNet(style.Format.ToUpperInvariant());
                            if (!string.IsNullOrEmpty(fmt)) {
                                text = formattable.ToString(fmt, CultureInfo.CurrentCulture);
                            } else {
                                text = formattable.ToString();
                            }
                        } else {
                            text = value.ToString();
                        }
                    }

                    // Get font (default or style)
                    var font = _book.DefaultFont;
                    if (style?.Font != null) {
                        font = style.Font;
                    }

                    // Measure string (with a little tolerance)
                    var size = _graphics.MeasureString(text + ".", font);

                    // Keep widest so far, capped to the maximum
                    if (size.Width > _colWidths[i]) {
                        _colWidths[i] = Math.Min(size.Width, maxWidth);
                    }
                }
            }
        }

        /// <summary>
        /// Adjusts all the column widths to match the content
        /// </summary>
        /// <param name="minWidth">Minimum width in twips</param>
        /// <param name="maxWidth">Maximum width in twips</param>
        public void AdjustColumnsToContent(
            double minWidth,
            double maxWidth)
        {
            for (var i = 0; i <= _row; i++) {
                AutoSizeColumnsForRow(_row, (float)maxWidth);
            }
        }

        /// <summary>
        /// Adjusts all the column widths to match the content
        /// </summary>
        /// <param name="minHeight">Minimum height in twips</param>
        /// <param name="maxHeight">Maximum height in twips</param>
        public void AdjustRowsToContent(
            double minHeight,
            double maxHeight)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Set the size of a specific column in twips
        /// </summary>
        /// <param name="col">Column to set the width for</param>
        /// <param name="width">Width of the column in twips</param>
        public void SetColumnWidth(
            int col,
            double width)
        {
            if (_configuration.AutoSizeColumns && _colWidths != null && col < _colWidths.Length) {
                _colWidths[col] = C1XLBook.TwipsToPixels(width);
            } else {
                _sheet.Columns[col].Width = (int)width;
            }
        }

        /// <summary>
        /// Set the height of a specific row in twips
        /// </summary>
        /// <param name="row">Row to set the height for</param>
        /// <param name="height">Height of the column in twips</param>
        public void SetRowHeight(
            int row,
            double height)
        {
            _sheet.Rows[row].Height = (int)height;
        }

        /// <summary>
        /// Writes a field to the Excel file natively
        /// </summary>
        /// <param name="field">The field object to write.</param>
        protected void WriteFieldNative(
            object field)
        {
            _sheet[_row, _col++].Value = field;
        }

        /// <summary>
        /// Writes a field to the Excel file
        /// </summary>
        /// <param name="field">The field object to write.</param>
        /// <param name="converter">Type converter to use</param>
        /// <param name="typeConverterOptions">Type converter options to use</param>
        protected void WriteFieldConverted(
            object field,
            ITypeConverter converter,
            TypeConverterOptions typeConverterOptions)
        {
            _sheet[_row, _col++].Value = converter.ConvertToExcel(typeConverterOptions, field);
        }

        /// <summary>
        /// Writes a field to the Excel file as a formula. We assume the incoming value
        /// is a string
        /// </summary>
        /// <param name="field">The field object to write.</param>
        protected void WriteFieldFormula(
            object field)
        {
            if (field is string s && s.StartsWith("=")) {
                _sheet[_row, _col].Value = "";
                _sheet[_row, _col++].Formula = s;
            } else {
                _sheet[_row, _col++].Value = field;
            }
        }

        /// <summary>
        /// Ends writing of the current record and starts a new record. This is used
        /// when manually writing records with WriteField.
        /// </summary>
        private void NextRecord()
        {
            if (_configuration.AutoSizeColumns) {
                AutoSizeColumnsForRow(_row, (float)_configuration.MaxColumnWidth);
            }
            _col = 0;
            _row++;
        }

        /// <summary>
        /// Writes the header record from the given properties.
        /// </summary>
        /// <param name="properties">The properties for the records.</param>
        private void WriteHeader(
            ExcelPropertyMapCollection properties)
        {
            // Write the header fields
            foreach (var property in properties) {
                if (CanWrite(property)) {
                    _sheet[_row, _col++].Value = property.Data.Names.FirstOrDefault();
                }
            }

            // Set the style for the header to bold if desired
            if (_configuration.HeaderIsBold) {
                _sheet.Rows[_row].Style = new XLStyle(_book) {
                    Font = new Font(_book.DefaultFont, FontStyle.Bold),
                };
            }

            // Move to the next record
            NextRecord();
        }

        /// <summary>
        /// Writes out the column styles for the record
        /// </summary>
        /// <param name="properties">Properties for the record</param>
        private void WriteColumnStyles(
            ExcelPropertyMapCollection properties)
        {
            // Write the column styles for all the columns
            for (var col = 0; col < properties.Count; col++) {
                // Determine if this property is written
                var propertyMap = properties[col];
                if (!CanWrite(propertyMap)) {
                    continue;
                }

                // Now get the property converter and options
                var data = propertyMap.Data;
                var typeConverterOptions = TypeConverterOptions.Merge(
                    TypeConverterOptionsFactory.GetOptions(data.Property.PropertyType, _configuration.CultureInfo),
                    data.TypeConverterOptions);

                // Write the cell formatting style if defined for this type
                var format = data.TypeConverter.ExcelFormatString(typeConverterOptions);
                if (format != null) {
                    _sheet.Columns[col].Style = new XLStyle(_book) {
                        Format = format,
                    };
                }
            }
        }

        /// <summary>
        /// Writes the list of typed records to the Excel file.
        /// </summary>
        /// <param name="records">The list of records to write.</param>
        /// <param name="writeHeader">True to write the header, false to not write the header</param>
        public void WriteRecords<T>(
            IEnumerable<T> records,
            bool writeHeader = true)
            where T : class
        {
            // Get the type of all the records
            var type = typeof(T);

            // Make sure it is mapped
            if (_configuration.Maps[type] == null) {
                _configuration.Maps.Add(_configuration.AutoMap(type));
            }

            // Get a list of all the properties so they will be sorted properly.
            var properties = new ExcelPropertyMapCollection();
            AddProperties(properties, _configuration.Maps[type]);
            if (properties.Count == 0) {
                throw new ExcelWriterException($"No properties are mapped for type '{type.FullName}'.");
            }


            // Only write the header the first time we are called (allows for paginating results into a file)
            if (writeHeader) {
                WriteHeader(properties);
            }

            // Write all the column styles
            WriteColumnStyles(properties);

            // Get the action method for writing the records out
            Delegate writeRecord = null;
            try {
                writeRecord = GetWriteRecordAction(type, properties);
            } catch (Exception ex) {
                ExceptionHelper.AddExceptionDataMessage(ex, type);
                throw;
            }

            // Now process each record
            foreach (var record in records) {
                writeRecord.DynamicInvoke(record);
                NextRecord();
            }
        }

        /// <summary>
        /// Closes the writer and saves the written data to the stream. Automatically called
        /// when disposed.
        /// </summary>
        public void Close()
        {
            if (_book != null) {
                // Set the column widths if we are doing auto sizing
                PerformColumnResize();

                // Now save the Excel file to the output stream
                _book.Save(_stream, FileFormat.OpenXml);

                // Clean up and dispose of everything
                _book?.Dispose();
                _graphics?.Dispose();
                _sheet = null;
                _book = null;
                _graphics = null;
            }
        }

        /// <summary>
        /// Perform the column sizing for the sheet, and then clear out the column widths
        /// </summary>
        private void PerformColumnResize()
        {
            if (_configuration.AutoSizeColumns && _colWidths != null) {
                for (var i = 0; i < _colWidths.Length; i++) {
                    _sheet.Columns[i].Width = C1XLBook.PixelsToTwips(_colWidths[i]);
                }
            }
            _colWidths = null;
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
                Close();
            }
            _disposed = true;
        }

        /// <summary>
        /// Adds the properties from the mapping. This will recursively
        /// traverse the mapping tree and add all properties for
        /// reference maps.
        /// </summary>
        /// <param name="properties">The properties to be added to.</param>
        /// <param name="mapping">The mapping where the properties are added from.</param>
        protected void AddProperties(
            ExcelPropertyMapCollection properties,
            ExcelClassMapBase mapping)
        {
            properties.AddRange(mapping.PropertyMaps);
            foreach (var refMap in mapping.ReferenceMaps) {
                AddProperties(properties, refMap.Mapping);
            }
        }

        /// <summary>
        /// Creates a property expression for the given property on the record.
        /// This will recursively traverse the mapping to find the property
        /// and create a safe property accessor for each level as it goes.
        /// </summary>
        /// <param name="recordExpression">The current property expression.</param>
        /// <param name="mapping">The mapping to look for the property to map on.</param>
        /// <param name="propertyMap">The property map to look for on the mapping.</param>
        /// <returns>An Expression to access the given property.</returns>
        protected Expression CreatePropertyExpression(
            Expression recordExpression,
            ExcelClassMapBase mapping,
            ExcelPropertyMap propertyMap)
        {
            // Handle the simple case where the property is on this level.
            if (mapping.PropertyMaps.Any(pm => pm == propertyMap)) {
                return Expression.Property(recordExpression, propertyMap.Data.Property);
            }

            // The property isn't on this level of the mapping. We need to search down through the reference maps.
            foreach (var refMap in mapping.ReferenceMaps) {
                // Recursively find the property access expression for this property
                var wrapped = Expression.Property(recordExpression, refMap.Property);
                var propertyExpression = CreatePropertyExpression(wrapped, refMap.Mapping, propertyMap);
                if (propertyExpression == null) {
                    // Not in this reference map, try the next one
                    continue;
                }

                // Build an expression that looks like this for value types:
                //
                // (record.RefMap == null) ? return new type() : record.RefMap.Property
                //
                // and like this for nullable types:
                //
                // (record.RefMap == null) ? return null : record.RefMap.Property
                //
                // So that the properties of the reference mapped object will be written as null or a default value
                // if the reference map is not present in the record being written
                var nullCheckExpression = Expression.Equal(wrapped, Expression.Constant(null));
                var isValueType = propertyMap.Data.Property.PropertyType.IsValueType;
                var defaultValueExpression = isValueType
                    ? (Expression)Expression.New(propertyMap.Data.Property.PropertyType)
                    : Expression.Constant(null, propertyMap.Data.Property.PropertyType);
                var conditionExpression = Expression.Condition(nullCheckExpression, defaultValueExpression, propertyExpression);
                return conditionExpression;
            }

            // We get here if the property did not match anything in this mapping
            return null;
        }

        /// <summary>
        /// Gets the action delegate used to write the custom
        /// class object to the writer.
        /// </summary>
        /// <param name="type">The type of the custom class being written.</param>
        /// <param name="properties">Properties for the record</param>
        /// <returns>The action delegate.</returns>
        private Delegate GetWriteRecordAction(
            Type type,
            ExcelPropertyMapCollection properties)
        {
            if (!_typeActions.ContainsKey(type)) {
                // Define the parameter to the action to pass in the record
                var recordParameter = Expression.Parameter(type, "record");

                // Build delegates to write every property out
                var delegates = new List<Delegate>();
                foreach (var propertyMap in properties) {
                    // Ignore properties that are not written
                    if (!CanWrite(propertyMap)) {
                        continue;
                    }

                    // Get the type converter and converter options for this type
                    var data = propertyMap.Data;
                    var typeConverter = data.TypeConverter;
                    var typeConverterOptions = TypeConverterOptions.Merge(
                        TypeConverterOptionsFactory.GetOptions(data.Property.PropertyType, _configuration.CultureInfo),
                        data.TypeConverterOptions);

                    // Create an expression to extract the field from the record
                    var fieldExpression = CreatePropertyExpression(recordParameter, _configuration.Maps[type], propertyMap);

                    Expression actionExpression;
                    if (typeConverterOptions.IsFormula) {
                        // Define an expression to call WriteFieldFormula(property)
                        actionExpression = Expression.Call(
                            Expression.Constant(this),
                            GetType().GetMethod("WriteFieldFormula", BindingFlags.NonPublic | BindingFlags.Instance),
                            Expression.Convert(fieldExpression, typeof(object)));
                    } else if (typeConverter.AcceptsNativeType) {
                        // Define an expression to call WriteFieldNative(property)
                        actionExpression = Expression.Call(
                            Expression.Constant(this),
                            GetType().GetMethod("WriteFieldNative", BindingFlags.NonPublic | BindingFlags.Instance),
                            Expression.Convert(fieldExpression, typeof(object)));
                    } else {
                        // Define an expression to call WriteFieldConverted(property, typeConverter, typeConverterOptions)
                        actionExpression = Expression.Call(
                            Expression.Constant(this),
                            GetType().GetMethod("WriteFieldConverted", BindingFlags.NonPublic | BindingFlags.Instance),
                            Expression.Convert(fieldExpression, typeof(object)),
                            Expression.Constant(typeConverter),
                            Expression.Constant(typeConverterOptions));
                    }

                    // Now create a lambda expression and compile it
                    var actionType = typeof(Action<>).MakeGenericType(type);
                    delegates.Add(Expression.Lambda(actionType, actionExpression, recordParameter).Compile());
                }

                // Combine all the delegates together so they are executed in order
                _typeActions[type] = Delegate.Combine(delegates.ToArray());
            }
            return _typeActions[type];
        }

        /// <summary>
        /// Checks if the property can be written.
        /// </summary>
        /// <param name="propertyMap">The property map that we are checking.</param>
        /// <returns>A value indicating if the property can be written.
        /// True if the property can be written, otherwise false.</returns>
        protected bool CanWrite(
            ExcelPropertyMap propertyMap)
        {
            var cantWrite =
                // Ignored properties
                propertyMap.Data.Ignore ||
                // Properties that don't have a public getter and we are honoring the accessor modifier
                propertyMap.Data.Property.GetGetMethod() == null && !_configuration.IgnorePrivateAccessor ||
                // Properties that don't have a getter at all
                propertyMap.Data.Property.GetGetMethod(true) == null;
            return !cantWrite;
        }
    }
}
#endif
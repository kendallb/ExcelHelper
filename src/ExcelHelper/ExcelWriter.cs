/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ClosedXML.Excel;
using ExcelHelper.Configuration;
using ExcelHelper.TypeConversion;

namespace ExcelHelper
{
    /// <summary>
    /// Used to write Excel files.
    /// </summary>
    public class ExcelWriter : IExcelWriter
    {
        private bool _disposed;
        private bool _closed;
        private readonly Stream _stream;
        private XLWorkbook _book;
        private IXLWorksheet _sheet;
        private ExcelFont _defaultFont;
        private int _row;
        private int _col;
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
        public ExcelWriter(
            Stream stream)
            : this(stream, new ExcelConfiguration())
        {
        }

        /// <summary>
        /// Creates a new Excel writer using the given <see cref="Stream"/> and <see cref="ExcelConfiguration"/>.
        /// </summary>
        /// <param name="stream">The <see cref="Stream"/> used to write the Excel file.</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelWriter(
            Stream stream,
            ExcelConfiguration configuration)
        {
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            _stream = stream ?? throw new ArgumentNullException(nameof(stream));
            _book = new XLWorkbook(XLEventTracking.Disabled);
            ChangeSheet(0);

            // Set the default font to Calibri 11, which is the default in newer versions of office
            DefaultFont = new ExcelFont("Calibri", 11, ExcelFontStyle.Regular);
        }

        /// <summary>
        /// Gets or sets the default font for the Excel file
        /// </summary>
        public ExcelFont DefaultFont
        {
            get => _defaultFont;
            set
            {
                var font = _book.Style.Font;
                font.Bold = (value.Style & ExcelFontStyle.Bold) != 0;
                font.Italic = (value.Style & ExcelFontStyle.Italic) != 0;
                font.Underline = (value.Style & ExcelFontStyle.Underline) != 0 ? XLFontUnderlineValues.Single : XLFontUnderlineValues.None;
                font.Strikethrough = (value.Style & ExcelFontStyle.Strikeout) != 0;
                font.FontSize = value.FontSize;
                font.FontName = value.FontName;
                _defaultFont = value;
            }
        }

        /// <summary>
        /// Update an Excel cell style with the passed in attributes
        /// </summary>
        /// <param name="cellStyle">IXLStyle to update</param>
        /// <param name="numberFormat">Optional number formatting string for the cell</param>
        /// <param name="dateFormat">Optional DateTime formatting string for the cell</param>
        /// <param name="fontStyle">Optional font style for the cell</param>
        /// <param name="fontSize">Optional font size for the cell</param>
        /// <param name="fontName">Optional font name for the cell</param>
        /// <param name="horizontalAlign">Optional horizontal alignment</param>
        /// <param name="verticalAlign">Optional vertical alignment</param>
        private void UpdateStyle(
            IXLStyle cellStyle,
            string numberFormat,
            string dateFormat,
            ExcelFontStyle? fontStyle,
            float? fontSize,
            string fontName,
            ExcelAlignHorizontal? horizontalAlign,
            ExcelAlignVertical? verticalAlign)
        {
            // Set up native formatting if provided
            if (!string.IsNullOrEmpty(numberFormat)) {
                cellStyle.NumberFormat.SetFormat(numberFormat);
            }
            if (!string.IsNullOrEmpty(dateFormat)) {
                cellStyle.DateFormat.SetFormat(dateFormat);
            }

            // Apply font styling if defined
            if (fontStyle != null || fontSize != null || fontName != null) {
                var defaultFont = _defaultFont;
                var style = fontStyle ?? defaultFont.Style;
                var font = cellStyle.Font;
                font.Bold = (style & ExcelFontStyle.Bold) != 0;
                font.Italic = (style & ExcelFontStyle.Italic) != 0;
                font.Underline = (style & ExcelFontStyle.Underline) != 0 ? XLFontUnderlineValues.Single : XLFontUnderlineValues.None;
                font.Strikethrough = (style & ExcelFontStyle.Strikeout) != 0;
                font.FontSize = fontSize ?? defaultFont.FontSize;
                font.FontName = fontName ?? defaultFont.FontName;
            }

            // Apply the horizontal alignment if defined
            if (horizontalAlign != null) {
                switch (horizontalAlign) {
                    case ExcelAlignHorizontal.General:
                        cellStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.General;
                        break;
                    case ExcelAlignHorizontal.Left:
                        cellStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        break;
                    case ExcelAlignHorizontal.Center:
                        cellStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        break;
                    case ExcelAlignHorizontal.Right:
                        cellStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        break;
                    case ExcelAlignHorizontal.Fill:
                        cellStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Fill;
                        break;
                    case ExcelAlignHorizontal.Justify:
                        cellStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Justify;
                        break;
                }
            }

            // Apply the vertical alignment if defined
            if (verticalAlign !=  null) {
                switch (verticalAlign) {
                    case ExcelAlignVertical.Top:
                        cellStyle.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                        break;
                    case ExcelAlignVertical.Center:
                        cellStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        break;
                    case ExcelAlignVertical.Bottom:
                        cellStyle.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                        break;
                    case ExcelAlignVertical.Justify:
                        cellStyle.Alignment.Vertical = XLAlignmentVerticalValues.Justify;
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
            var sheets = _book.Worksheets;
            if (sheet >= sheets.Count) {
                for (var i = sheets.Count; i <= sheet; i++) {
                    sheets.Add($"Sheet{i+1}");
                }
            }

            // Toss the old sheet and reference the new one
            _sheet = sheets.Worksheet(sheet + 1);
            _row = _col = 1;
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
            ExcelFontStyle? fontStyle = null,
            float? fontSize = null,
            string fontName = null,
            ExcelAlignHorizontal? horizontalAlign = null,
            ExcelAlignVertical? verticalAlign = null)
        {
            // Clear the cell if the field is null
            var cell = _sheet.Cell(row + 1, col + 1);
            if (field == null) {
                cell.SetValue((object)null);
                return;
            }

            // Find the type conversion options
            var type = typeof(T);
            var converter = TypeConverterFactory.GetConverter(type);
            var options = TypeConverterOptions.Merge(TypeConverterOptionsFactory.GetOptions(type, _configuration.CultureInfo));

            // Set the formatting options to override the defaults
            numberFormat = numberFormat ?? options.NumberFormat;
            dateFormat = dateFormat ?? options.DateFormat;

            // Apply the style to this cell if defined
            if (numberFormat != null || dateFormat != null || fontStyle != null || fontSize != null || fontName != null || horizontalAlign != null || verticalAlign != null) {
                UpdateStyle(cell.Style, numberFormat, dateFormat, fontStyle, fontSize, fontName, horizontalAlign, verticalAlign);
            }

            // Now write the cell contents
            if (field.GetType() == typeof(string)) {
                var s = (string)(object)field;
                if (s != null && s.StartsWith("=")) {
                    // Write as a formula if it starts with an equals sign
                    cell.FormulaA1 = s;
                } else {
                    cell.SetValue(s);
                }
            } else if (converter.AcceptsNativeType) {
                cell.SetValue(field);
            } else {
                cell.SetValue(converter.ConvertToExcel(options, field));
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
            ExcelFontStyle? fontStyle = null,
            float? fontSize = null,
            string fontName = null,
            ExcelAlignHorizontal? horizontalAlign = null,
            ExcelAlignVertical? verticalAlign = null)
        {
            var xlColumn = _sheet.Column(col + 1);
            UpdateStyle(xlColumn.Style, numberFormat, dateFormat, fontStyle, fontSize, fontName, horizontalAlign, verticalAlign);
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
            ExcelFontStyle? fontStyle = null,
            float? fontSize = null,
            string fontName = null,
            ExcelAlignHorizontal? horizontalAlign = null,
            ExcelAlignVertical? verticalAlign = null)
        {
            var xlRow = _sheet.Row(row + 1);
            UpdateStyle(xlRow.Style, numberFormat, dateFormat, fontStyle, fontSize, fontName, horizontalAlign, verticalAlign);
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
            var columns = _sheet.Columns();
            columns.AdjustToContents(minWidth, maxWidth);
        }

        /// <summary>
        /// Adjusts all the column widths to match the content for specific rows
        /// </summary>
        /// <param name="startRow">The row to start calculating the column width</param>
        /// <param name="minWidth">Minimum width in twips</param>
        /// <param name="maxWidth">Maximum width in twips</param>
        public void AdjustColumnsToContent(
            int startRow,
            double minWidth,
            double maxWidth)
        {
            var columns = _sheet.Columns();
            columns.AdjustToContents(startRow + 1, minWidth, maxWidth);
        }

        /// <summary>
        /// Adjusts all the column widths to match the content for specific rows
        /// </summary>
        /// <param name="startRow">The row to start calculating the column width</param>
        /// <param name="endRow">The row to end calculating the column width (inclusive)</param>
        /// <param name="minWidth">Minimum width in twips</param>
        /// <param name="maxWidth">Maximum width in twips</param>
        public void AdjustColumnsToContent(
            int startRow,
            int endRow,
            double minWidth,
            double maxWidth)
        {
            var columns = _sheet.Columns();
            columns.AdjustToContents(startRow + 1, endRow + 1, minWidth, maxWidth);
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
            var rows = _sheet.Rows();
            rows.AdjustToContents(minHeight, maxHeight);
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
            var xlColumn = _sheet.Column(col + 1);
            xlColumn.Width = width;
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
            var xlRow = _sheet.Row(row + 1);
            xlRow.Height = height;
        }

        /// <summary>
        /// Writes a field to the Excel file natively
        /// </summary>
        /// <param name="field">The field object to write.</param>
        protected void WriteFieldNative(
            object field)
        {
            _sheet.Cell(_row, _col++).SetValue(field);
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
            _sheet.Cell(_row, _col++).SetValue(converter.ConvertToExcel(typeConverterOptions, field));
        }

        /// <summary>
        /// Writes a field to the Excel file as a formula. We assume the incoming value
        /// is a string
        /// </summary>
        /// <param name="field">The field object to write.</param>
        protected void WriteFieldFormula(
            object field)
        {
            var cell = _sheet.Cell(_row, _col++);
            if (field.GetType() == typeof(string)) {
                var s = (string)field;
                if (s != null && s.StartsWith("=")) {
                    cell.FormulaA1 = s;
                } else {
                    cell.SetValue(field);
                }
            } else {
                cell.SetValue(field);
            }
        }

        /// <summary>
        /// Ends writing of the current record and starts a new record. This is used
        /// when manually writing records with WriteField.
        /// </summary>
        private void NextRecord()
        {
            _col = 1;
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
                    _sheet.Cell(_row, _col++).SetValue(property.Data.Names.FirstOrDefault());
                }
            }

            // Set the style for the header to bold if desired
            if (_configuration.HeaderIsBold) {
                var xlRow = _sheet.Row(_row);
                xlRow.Style.Font.SetBold(true);
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
                var isDate = data.TypeConverter.ConvertedType == typeof(DateTime);
                var format = isDate ? typeConverterOptions.DateFormat : typeConverterOptions.NumberFormat;
                if (format != null) {
                    var xlColumn = _sheet.Column(col + 1);
                    if (isDate) {
                        xlColumn.Style.DateFormat.Format = format;
                    } else {
                        xlColumn.Style.NumberFormat.Format = format;
                    }
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
                // Write the header
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
            if (_book != null && !_closed) {
                // Set the column widths if we are doing auto sizing
                PerformColumnResize();

                // Now save the Excel file to the output stream
                _book.SaveAs(_stream);

                // Mark us as now closed
                _closed = true;
            }
        }

        /// <summary>
        /// Perform the column sizing for the sheet, and then clear out the column widths
        /// </summary>
        private void PerformColumnResize()
        {
            // ReSharper disable once UseNullPropagationWhenPossible
            if (_configuration.AutoSizeColumns && _sheet != null) {
                var columns = _sheet.Columns();
                columns.AdjustToContents(0, _configuration.MaxColumnWidth);
            }
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

            // Clean up and dispose of everything. We do it during the dispose as it makes it easier for us to measure the memory usage
            _book?.Dispose();
            _sheet = null;
            _book = null;
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
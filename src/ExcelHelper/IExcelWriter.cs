/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using ExcelHelper.Configuration;

namespace ExcelHelper
{
    /// <summary>
    /// Defines methods used to write to a Excel file.
    /// </summary>
    public interface IExcelWriter : IDisposable
    {
        /// <summary>
        /// Gets the configuration.
        /// </summary>
        IExcelConfiguration Configuration { get; }

        /// <summary>
        /// Gets or sets the default font for the Excel file
        /// </summary>
        Font DefaultFont { get; set; }

        /// <summary>
        /// Changes to using the passed in sheet. Note that changing to a new sheet automatically resets the
        /// internal row and column counter used by WriteRecords.
        /// </summary>
        /// <param name="sheet">Sheet to change to</param>
        void ChangeSheet(
            int sheet);

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
        void WriteCell<T>(
            int row,
            int col,
            T field,
            string numberFormat = null,
            string dateFormat = null,
            FontStyle? fontStyle = null,
            float? fontSize = null,
            string fontName = null);

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
        void SetColumnFormat(
            int col,
            string numberFormat = null,
            string dateFormat = null,
            FontStyle? fontStyle = null,
            float? fontSize = null,
            string fontName = null);

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
        void SetRowFormat(
            int row,
            string numberFormat = null,
            string dateFormat = null,
            FontStyle? fontStyle = null,
            float? fontSize = null,
            string fontName = null);

        /// <summary>
        /// Adjusts all the column widths to match the content
        /// </summary>
        /// <param name="minWidth">Minimum width in twips</param>
        /// <param name="maxWidth">Maximum width in twips</param>
        void AdjustColumnsToContent(
            double minWidth = 0,
            double maxWidth = double.MaxValue);

        /// <summary>
        /// Adjusts all the column widths to match the content
        /// </summary>
        /// <param name="minHeight">Minimum height in twips</param>
        /// <param name="maxHeight">Maximum height in twips</param>
        void AdjustRowsToContent(
            double minHeight = 0,
            double maxHeight = double.MaxValue);

        /// <summary>
        /// Set the width of a specific column in twips
        /// </summary>
        /// <param name="col">Column to set the width for</param>
        /// <param name="width">Width of the column in twips</param>
        void SetColumnWidth(
            int col,
            double width);

        /// <summary>
        /// Set the height of a specific row in twips
        /// </summary>
        /// <param name="row">Row to set the height for</param>
        /// <param name="height">Height of the column in twips</param>
        void SetRowHeight(
            int row,
            double height);

        /// <summary>
        /// Writes the list of typed records to the Excel file.
        /// </summary>
        /// <param name="records">The list of records to write.</param>
        /// <param name="writeHeader">True to write the header, false to not write the header</param>
        void WriteRecords<T>(
            IEnumerable<T> records,
            bool writeHeader = true)
            where T : class;

        /// <summary>
        /// Closes the writer and saves the written data to the stream. Automatically called
        /// when disposed.
        /// </summary>
        void Close();
    }
}
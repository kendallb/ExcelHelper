/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using System.Collections.Generic;
using System.Reflection;
using ExcelHelper.Configuration;

namespace ExcelHelper
{
    /// <summary>
    /// Defines methods used to read parsed data from a Excel file.
    /// </summary>
    public interface IExcelReader : IDisposable
    {
        /// <summary>
        /// Gets or sets the configuration.
        /// </summary>
        IExcelConfiguration Configuration { get; }

        /// <summary>
        /// Returns the total number of columns
        /// </summary>
        int TotalColumns { get; }

        /// <summary>
        /// Returns the total number of sheets in the Excel file
        /// </summary>
        int TotalSheets { get; }

        /// <summary>
        /// Returns the name of the current sheet
        /// </summary>
        string SheetName { get; }

        /// <summary>
        /// Changes to using a specific sheet in the Excel file. Note that changing to a new sheet automatically resets the 
        /// internal row counter used by GetRecords.
        /// </summary>
        /// <param name="sheet">Sheet to change to (0 to TotalSheets - 1)</param>
        /// <returns>True on success, false if the sheet is out of range</returns>
        bool ChangeSheet(
            int sheet);

        /// <summary>
        /// Skip over the given number of rows. Useful for cases where the header columns are not in the first row.
        /// </summary>
        /// <param name="count">The number of rows to skip</param>
        void SkipRows(
            int count);

        /// <summary>
        /// Moves to the next row in the Excel file when using the GetCell() function
        /// </summary>
        /// <returns>True if there is another row, false if not</returns>
        bool ReadRow();

        /// <summary>
        /// Reads a cell from the Excel file at the current row
        /// </summary>
        /// <typeparam name="T">The type of the field.</typeparam>
        /// <param name="index">Column to write the field to.</param>
        /// <returns>The value from the cell converted to the specific type</returns>
        T GetColumn<T>(
            int index);

        /// <summary>
        /// Gets all the records in the Excel file and converts each to <see cref="Type"/> T.
        /// </summary>
        /// <typeparam name="T">The <see cref="Type"/> of the record.</typeparam>
        /// <returns>An <see cref="IEnumerable{T}" /> of records.</returns>
        IEnumerable<T> GetRecords<T>();

        /// <summary>
        /// Gets all the records in the Excel file and converts each to dictionary of strings to strings.
        /// </summary>
        /// <returns>An enumeration of dictionaries.</returns>
        IEnumerable<Dictionary<string, string>> GetRecordsAsDictionary();

        /// <summary>
        /// Gets a list of all the properties for columns that are found in the import. This can only be called
        /// after first calling GetRecords()
        /// </summary>
        /// <returns>List of properties for columns found in the Excel file.</returns>
        List<PropertyInfo> GetImportedColumns();
    }
}
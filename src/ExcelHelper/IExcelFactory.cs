/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System.IO;
using ExcelHelper.Configuration;

namespace ExcelHelper
{
    /// <summary>
    /// Defines methods used to create
    /// ExcelHelper classes.
    /// </summary>
    public interface IExcelFactory
    {
        /// <summary>
        /// Creates an <see cref="IExcelReader"/>.
        /// </summary>
        /// <param name="stream">The text stream to use for the excel stream.</param>
        /// <returns>The created stream.</returns>
        IExcelReader CreateReader(
            Stream stream);

        /// <summary>
        /// Creates an <see cref="IExcelReader"/>.
        /// </summary>
        /// <param name="stream">The text stream to use for the excel stream.</param>
        /// <param name="configuration">The configuration to use for the stream.</param>
        /// <returns>The created stream.</returns>
        IExcelReader CreateReader(
            Stream stream,
            ExcelConfiguration configuration);

        /// <summary>
        /// Creates an <see cref="IExcelWriter"/>.
        /// </summary>
        /// <param name="stream">The stream used to write the Excel file.</param>
        /// <returns>The created writer.</returns>
        IExcelWriter CreateWriter(
            Stream stream);

        /// <summary>
        /// Creates an <see cref="IExcelWriter"/>.
        /// </summary>
        /// <param name="stream">The stream used to write the Excel file.</param>
        /// <param name="configuration">The configuration to use for the writer.</param>
        /// <returns>The created writer.</returns>
        IExcelWriter CreateWriter(
            Stream stream,
            ExcelConfiguration configuration);
    }
}
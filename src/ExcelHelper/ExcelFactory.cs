/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 */

using System.IO;
using ExcelHelper.Configuration;

namespace ExcelHelper
{
    /// <summary>
    /// Creates ExcelHelper classes.
    /// </summary>
    public class ExcelFactory : IExcelFactory
    {
        /// <summary>
        /// Creates an <see cref="IExcelReader"/>.
        /// </summary>
        /// <param name="stream">The text stream to use for the excel stream.</param>
        /// <returns>The created stream.</returns>
        public virtual IExcelReader CreateReader(
            Stream stream)
        {
#if BIFF8_SUPPORT
            try {
                return new ExcelReader(stream);
            } catch (FileFormatException) {
                return new ExcelReaderBiff8(stream);
            }
#else
            return new ExcelReader(stream);
#endif
        }

        /// <summary>
        /// Creates an <see cref="IExcelReader"/>.
        /// </summary>
        /// <param name="stream">The text stream to use for the excel stream.</param>
        /// <param name="configuration">The configuration to use for the stream.</param>
        /// <returns>The created stream.</returns>
        public virtual IExcelReader CreateReader(
            Stream stream,
            ExcelConfiguration configuration)
        {
#if BIFF8_SUPPORT
            try {
                return new ExcelReader(stream, configuration);
            } catch (FileFormatException) {
                return new ExcelReaderBiff8(stream, configuration);
            }
#else
            return new ExcelReader(stream, configuration);
#endif
        }

        /// <summary>
        /// Creates an <see cref="IExcelWriter"/>.
        /// </summary>
        /// <param name="stream">The stream used to write the Excel file.</param>
        /// <returns>The created writer.</returns>
        public virtual IExcelWriter CreateWriter(
            Stream stream)
        {
            return new ExcelWriter(stream);
        }

        /// <summary>
        /// Creates an <see cref="IExcelWriter"/>.
        /// </summary>
        /// <param name="stream">The stream used to write the Excel file.</param>
        /// <param name="configuration">The configuration to use for the writer.</param>
        /// <returns>The created writer.</returns>
        public virtual IExcelWriter CreateWriter(
            Stream stream,
            ExcelConfiguration configuration)
        {
            return new ExcelWriter(stream, configuration);
        }
    }
}
/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;

namespace ExcelHelper
{
    /// <summary>
    /// Represents errors that occur due to bad data.
    /// </summary>
    public class ExcelBadDataException : ExcelHelperException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelBadDataException"/> class.
        /// </summary>
        public ExcelBadDataException()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelBadDataException"/> class
        /// with a specified error message.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public ExcelBadDataException(
            string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelBadDataException"/> class
        /// with a specified error message and a reference to the inner exception that 
        /// is the cause of this exception.
        /// </summary>
        /// <param name="message">The error message that explains the reason for the exception.</param>
        /// <param name="innerException">The exception that is the cause of the current exception, or a null reference (Nothing in Visual Basic) if no inner exception is specified.</param>
        public ExcelBadDataException(
            string message,
            Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
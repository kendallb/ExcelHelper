/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using System.Text;

namespace ExcelHelper
{
    /// <summary>
    /// Common exception tasks.
    /// </summary>
    internal static class ExceptionHelper
    {
        /// <summary>
        /// Adds ExcelHelper specific information to <see cref="Exception.Data"/>.
        /// </summary>
        /// <param name="exception">The exception to add the info to.</param>
        /// <param name="type">The type of object that was being created in the ExcelReader.</param>
        /// <param name="details">The details of the parsing error.</param>
        public static void AddExceptionDataMessage(
            Exception exception,
            Type type,
            ExcelReadErrorDetails details = null)
        {
            // An error could occur in the parser and get this message set on it, then occur in the
            // reader and have it set again. This is ok because when the reader calls this method,
            // it will have extra info to be added.
            try {
                exception.Data["ExcelHelper"] = GetErrorMessage(type, details);
            } catch (Exception ex) {
                var exString = new StringBuilder();
                exString.AppendLine("An error occurred while creating exception details.");
                exString.AppendLine();
                exString.AppendLine(ex.ToString());
                exception.Data["ExcelHelper"] = exString.ToString();
            }
        }

        /// <summary>
        /// Gets ExcelHelper information to be added to an exception.
        /// </summary>
        /// <param name="type">The type of object that was being created in the ExcelReader.</param>
        /// <param name="details">The details of the parsing error.</param>
        /// <returns>The ExcelHelper information.</returns>
        public static string GetErrorMessage(
            Type type,
            ExcelReadErrorDetails details = null)
        {
            var messageInfo = new StringBuilder();
            if (type != null) {
                messageInfo.AppendFormat("Type: '{0}'", type.FullName).AppendLine();
            }
            if (details != null) {
                messageInfo.AppendFormat("Row: '{0}' (1 based)", details.Row).AppendLine();
                messageInfo.AppendFormat("Column: '{0}' (1 based)", details.Column).AppendLine();
                messageInfo.AppendFormat("Field Name: '{0}'", details.FieldName).AppendLine();
                messageInfo.AppendFormat("Field Value: '{0}'", details.FieldValue).AppendLine();
            }
            return messageInfo.ToString();
        }
    }
}
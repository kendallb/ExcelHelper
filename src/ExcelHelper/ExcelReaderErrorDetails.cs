/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 */

namespace ExcelHelper
{
    /// <summary>
    /// Defines details about an error while parsing the file
    /// </summary>
    public class ExcelReadErrorDetails
    {
        /// <summary>
        /// Current row within the file where the error occurred
        /// </summary>
        public int Row { get; set; }

        /// <summary>
        /// Current column within the file where the error occurred
        /// </summary>
        public int Column { get; set; }

        /// <summary>
        /// Current field name of the column where the error occurred
        /// </summary>
        public string FieldName { get; set; }

        /// <summary>
        /// Actual value from the cell that caused the error
        /// </summary>
        public object FieldValue { get; set; }
    }
}
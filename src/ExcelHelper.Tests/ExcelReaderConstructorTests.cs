/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

#if !USE_C1_EXCEL
using System.IO;
using ClosedXML.Excel;
using ExcelHelper.Configuration;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelHelper.Tests
{
    [TestClass]
    public class ExcelReaderConstructorTests
    {
        [TestMethod]
        public void EnsureInternalsAreSetupWhenPasingReaderAndConfigTest()
        {
            using (var stream = new MemoryStream()) {
                // Make sure the stream is a valid Excel file
                using (var book = new XLWorkbook()) {
                    book.AddWorksheet("Sheet 1");
                    book.SaveAs(stream);
                }

                stream.Position = 0;
                var config = new ExcelConfiguration();
                using (var excel = new ExcelReader(stream, config)) {
                    Assert.AreSame(config, excel.Configuration);
                }
            }
        }
    }
}
#endif
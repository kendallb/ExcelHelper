/*
 * Copyright (C) 2004-2013 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 */

#if USE_C1_EXCEL
using System.IO;
using C1.C1Excel;
using ExcelHelper.Configuration;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelHelper.Tests
{
    [TestClass]
    public class ExcelReaderC1ConstructorTests
    {
        [TestMethod]
        public void EnsureInternalsAreSetupWhenPasingReaderAndConfigTest()
        {
            using (var stream = new MemoryStream()) {
                // Make sure the stream is a valid Excel file
                using (var book = new C1XLBook()) {
                    book.Save(stream);
                }

                stream.Position = 0;
                var config = new ExcelConfiguration();
                using (var excel = new ExcelReaderC1(stream, config)) {
                    Assert.AreSame(config, excel.Configuration);
                }
            }
        }
    }
}
#endif
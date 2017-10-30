/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System.IO;
using ExcelHelper.Configuration;
using Microsoft.VisualStudio.TestTools.UnitTesting;

#if USE_C1_EXCEL
namespace ExcelHelper.Tests
{
    [TestClass]
    public class ExcelWriterC1ConstructorTests
    {
        [TestMethod]
        public void EnsureInternalsAreSetupWhenPasingWriterAndConfigTest()
        {
            using (var stream = new MemoryStream()) {
                var config = new ExcelConfiguration();
                using (var excel = new ExcelWriterC1(stream, config)) {
                    Assert.AreSame(config, excel.Configuration);
                }
            }
        }
    }
}
#endif
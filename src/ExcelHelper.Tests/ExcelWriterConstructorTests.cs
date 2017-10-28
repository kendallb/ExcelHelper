/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 */

#if !USE_C1_EXCEL
using System.IO;
using ExcelHelper.Configuration;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelHelper.Tests
{
    [TestClass]
    public class ExcelWriterConstructorTests
    {
        [TestMethod]
        public void EnsureInternalsAreSetupWhenPasingWriterAndConfigTest()
        {
            using (var stream = new MemoryStream()) {
                var config = new ExcelConfiguration();
                using (var excel = new ExcelWriter(stream, config)) {
                    Assert.AreSame(config, excel.Configuration);
                }
            }
        }
    }
}
#endif
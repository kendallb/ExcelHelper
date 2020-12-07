/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System.IO;
using ExcelHelper.Configuration;
using NUnit.Framework;

namespace ExcelHelper.Tests
{
    [TestFixture]
    public class ExcelWriterConstructorTests
    {
        [Test]
        public void EnsureInternalsAreSetupWhenPassingWriterAndConfigTest()
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
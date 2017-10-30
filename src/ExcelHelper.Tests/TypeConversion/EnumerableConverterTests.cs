/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelHelper.TypeConversion;

namespace ExcelHelper.Tests.TypeConversion
{
    [TestClass]
    public class EnumerableConverterTests
    {
        [TestMethod]
        public void ConvertTest()
        {
            var converter = new EnumerableConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            try {
                converter.ConvertFromExcel(typeConverterOptions, "");
                Assert.Fail();
            } catch (ExcelTypeConverterException) {
            }
            try {
                converter.ConvertToExcel(typeConverterOptions, 5);
                Assert.Fail();
            } catch (ExcelTypeConverterException) {
            }
        }
    }
}
/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using System.Globalization;
using ExcelHelper.TypeConversion;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelHelper.Tests.TypeConversion
{
    [TestClass]
    public class GuidConverterTests
    {
        [TestMethod]
        public void PropertiesTest()
        {
            var converter = new GuidConverter();
            Assert.AreEqual(false, converter.AcceptsNativeType);
            Assert.AreEqual(typeof(Guid), converter.ConvertedType);
        }

        [TestMethod]
        public void ConvertToExcelTest()
        {
            var converter = new GuidConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            Assert.AreEqual("79f1a554-babd-41a1-8caf-185250a1fc21", converter.ConvertToExcel(typeConverterOptions, new Guid("79f1a554-babd-41a1-8caf-185250a1fc21")));
            Assert.AreEqual(null, converter.ConvertToExcel(typeConverterOptions, null));
        }

        [TestMethod]
        public void ConvertFromExcelTest()
        {
            var converter = new GuidConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            var guid = new Guid("79f1a554-babd-41a1-8caf-185250a1fc21");
            Assert.AreEqual(guid, converter.ConvertFromExcel(typeConverterOptions, "79f1a554-babd-41a1-8caf-185250a1fc21"));
            Assert.AreEqual(guid, converter.ConvertFromExcel(typeConverterOptions, " 79f1a554-babd-41a1-8caf-185250a1fc21 "));

            try {
                converter.ConvertFromExcel(typeConverterOptions, null);
                Assert.Fail();
            } catch (ExcelTypeConverterException) {
            }
        }
    }
}
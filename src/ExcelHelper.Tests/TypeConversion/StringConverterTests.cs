/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System.Globalization;
using ExcelHelper.TypeConversion;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelHelper.Tests.TypeConversion
{
    [TestClass]
    public class StringConverterTests
    {
        [TestMethod]
        public void PropertiesTest()
        {
            var converter = new StringConverter();
            Assert.AreEqual(true, converter.AcceptsNativeType);
            Assert.AreEqual(typeof(string), converter.ConvertedType);
        }

        [TestMethod]
        public void ConvertToExcelTest()
        {
            var converter = new StringConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture,
            };

            Assert.AreEqual("123", converter.ConvertToExcel(typeConverterOptions, "123"));
            Assert.AreEqual(null, converter.ConvertToExcel(typeConverterOptions, null));
        }

        [TestMethod]
        public void ConvertFromExcelTest()
        {
            var converter = new StringConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture,
            };

            Assert.AreEqual("12.3", converter.ConvertFromExcel(typeConverterOptions, 12.3));
            Assert.AreEqual("123", converter.ConvertFromExcel(typeConverterOptions, "123"));
            Assert.AreEqual("123", converter.ConvertFromExcel(typeConverterOptions, " 123 "));
            Assert.AreEqual("", converter.ConvertFromExcel(typeConverterOptions, null));
        }
    }
}
/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System.Globalization;
using ExcelHelper.TypeConversion;
using NUnit.Framework;

namespace ExcelHelper.Tests.TypeConversion
{
    [TestFixture]
    public class SingleConverterTests
    {
        [Test]
        public void PropertiesTest()
        {
            var converter = new SingleConverter();
            Assert.AreEqual(true, converter.AcceptsNativeType);
            Assert.AreEqual(typeof(float), converter.ConvertedType);
        }

        [Test]
        public void ConvertFromExcelTest()
        {
            var converter = new SingleConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            Assert.AreEqual((float)12.3, converter.ConvertFromExcel(typeConverterOptions, 12.3));
            Assert.AreEqual((float)12.3, converter.ConvertFromExcel(typeConverterOptions, "12.3"));
            Assert.AreEqual((float)12.3, converter.ConvertFromExcel(typeConverterOptions, " 12.3 "));
            Assert.AreEqual((float)0, converter.ConvertFromExcel(typeConverterOptions, null));

            try {
                converter.ConvertFromExcel(typeConverterOptions, "");
                Assert.Fail();
            } catch (ExcelTypeConverterException) {
            }
        }
    }
}
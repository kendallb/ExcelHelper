﻿/*
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
    public class DoubleConverterTests
    {
        [Test]
        public void PropertiesTest()
        {
            var converter = new DoubleConverter();
            Assert.AreEqual(true, converter.AcceptsNativeType);
            Assert.AreEqual(typeof(double), converter.ConvertedType);
        }

        [Test]
        public void ConvertFromExcelTest()
        {
            var converter = new DoubleConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            Assert.AreEqual((double)12.3m, converter.ConvertFromExcel(typeConverterOptions, 12.3));
            Assert.AreEqual(12.3, converter.ConvertFromExcel(typeConverterOptions, "12.3"));
            Assert.AreEqual(12.3, converter.ConvertFromExcel(typeConverterOptions, " 12.3 "));
            Assert.AreEqual((double)0, converter.ConvertFromExcel(typeConverterOptions, null));

            try {
                converter.ConvertFromExcel(typeConverterOptions, "");
                Assert.Fail();
            } catch (ExcelTypeConverterException) {
            }
        }
    }
}
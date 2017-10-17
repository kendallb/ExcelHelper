/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 */

using System;
using System.Globalization;
using ExcelHelper.TypeConversion;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelHelper.Tests.TypeConversion
{
    [TestClass]
    public class DateTimeConverterTests
    {
        [TestMethod]
        public void PropertiesTest()
        {
            var converter = new DateTimeConverter();
            Assert.AreEqual(true, converter.AcceptsNativeType);
            Assert.AreEqual(typeof(DateTime), converter.ConvertedType);
        }

        [TestMethod]
        public void ConvertFromExcelTest()
        {
            var converter = new DateTimeConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            var dateTime = DateTime.Now;
            var dateString = dateTime.ToString();

            // Valid conversions.
            Assert.AreEqual(dateString, converter.ConvertFromExcel(typeConverterOptions, dateTime).ToString());
            Assert.AreEqual(dateString, converter.ConvertFromExcel(typeConverterOptions, dateTime.ToOADate()).ToString());
            Assert.AreEqual(dateString, converter.ConvertFromExcel(typeConverterOptions, dateTime.ToString()).ToString());
            Assert.AreEqual(dateString, converter.ConvertFromExcel(typeConverterOptions, dateTime.ToString("o")).ToString());
            Assert.AreEqual(dateString, converter.ConvertFromExcel(typeConverterOptions, " " + dateTime + " ").ToString());

            // Empty conversions.
            dateTime = DateTime.MinValue;
            dateString = dateTime.ToString();
            Assert.AreEqual(dateString, converter.ConvertFromExcel(typeConverterOptions, " ").ToString());
            Assert.AreEqual(dateString, converter.ConvertFromExcel(typeConverterOptions, null).ToString());

            // Invalid conversions.
            try {
                converter.ConvertFromExcel(typeConverterOptions, 1);
                Assert.Fail();
            } catch (ExcelTypeConverterException) {
            }
        }

        [TestMethod]
        public void ComponentModelCompatibilityTest()
        {
            var converter = new DateTimeConverter();
            var cmConverter = new System.ComponentModel.DateTimeConverter();

            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            var val = (DateTime)cmConverter.ConvertFromString("");
            Assert.AreEqual(DateTime.MinValue, val);

            val = (DateTime)converter.ConvertFromExcel(typeConverterOptions, "");
            Assert.AreEqual(DateTime.MinValue, val);

            val = (DateTime)converter.ConvertFromExcel(typeConverterOptions, null);
            Assert.AreEqual(DateTime.MinValue, val);

            try {
                cmConverter.ConvertFromString("blah");
                Assert.Fail();
            } catch (FormatException) {
            }

            try {
                converter.ConvertFromExcel(typeConverterOptions, "blah");
            } catch (FormatException) {
            }
        }
    }
}
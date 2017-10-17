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
    public class TimeSpanConverterTests
    {
        [TestMethod]
        public void PropertiesTest()
        {
            var converter = new TimeSpanConverter();
            Assert.AreEqual(false, converter.AcceptsNativeType);
            Assert.AreEqual(typeof(TimeSpan), converter.ConvertedType);
        }

        [TestMethod]
        public void ConvertToExcelTest()
        {
            var converter = new TimeSpanConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            var dateTime = DateTime.Now;
            var timeSpan = new TimeSpan(dateTime.Hour, dateTime.Minute, dateTime.Second, dateTime.Millisecond);

            // Valid conversions.
            Assert.AreEqual(timeSpan.ToString(), converter.ConvertToExcel(typeConverterOptions, timeSpan));

            // Invalid conversions.
            Assert.AreEqual("1", converter.ConvertToExcel(typeConverterOptions, 1));
            Assert.AreEqual(null, converter.ConvertToExcel(typeConverterOptions, null));
        }

        [TestMethod]
        public void ConvertFromExcelTest()
        {
            var converter = new TimeSpanConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            var dateTime = DateTime.Now;
            var timeSpan = new TimeSpan(dateTime.Hour, dateTime.Minute, dateTime.Second, dateTime.Millisecond);
            var timeString = timeSpan.ToString();

            // Valid conversions.
            Assert.AreEqual(timeString, converter.ConvertFromExcel(typeConverterOptions, timeSpan.ToString()).ToString());
            Assert.AreEqual(timeString, converter.ConvertFromExcel(typeConverterOptions, " " + timeSpan + " ").ToString());

            // Invalid conversions.
            try {
                converter.ConvertFromExcel(typeConverterOptions, null);
                Assert.Fail();
            } catch (ExcelTypeConverterException) {
            }
        }

        [TestMethod]
        public void ComponentModelCompatibilityTest()
        {
            var converter = new TimeSpanConverter();
            var cmConverter = new System.ComponentModel.TimeSpanConverter();

            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            try {
                cmConverter.ConvertFromString("");
                Assert.Fail();
            } catch (FormatException) {
            }

            try {
                converter.ConvertFromExcel(typeConverterOptions, "");
                Assert.Fail();
            } catch (ExcelTypeConverterException) {
            }

            try {
                cmConverter.ConvertFromString(null);
                Assert.Fail();
            } catch (NotSupportedException) {
            }

            try {
                converter.ConvertFromExcel(typeConverterOptions, null);
                Assert.Fail();
            } catch (ExcelTypeConverterException) {
            }
        }
    }
}

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
using NUnit.Framework;
using NUnit.Framework.Legacy;

namespace ExcelHelper.Tests.TypeConversion
{
    [TestFixture]
    public class DateTimeConverterTests
    {
        [Test]
        public void PropertiesTest()
        {
            var converter = new DateTimeConverter();
            ClassicAssert.AreEqual(true, converter.AcceptsNativeType);
            ClassicAssert.AreEqual(typeof(DateTime), converter.ConvertedType);
        }

        [Test]
        public void ConvertFromExcelTest()
        {
            var converter = new DateTimeConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            var dateTime = DateTime.Now;
            var dateString = dateTime.ToString();

            // Valid conversions.
            ClassicAssert.AreEqual(dateString, converter.ConvertFromExcel(typeConverterOptions, dateTime).ToString());
            ClassicAssert.AreEqual(dateString, converter.ConvertFromExcel(typeConverterOptions, dateTime.ToOADate()).ToString());
            ClassicAssert.AreEqual(dateString, converter.ConvertFromExcel(typeConverterOptions, dateTime.ToString()).ToString());
            ClassicAssert.AreEqual(dateString, converter.ConvertFromExcel(typeConverterOptions, dateTime.ToString("o")).ToString());
            ClassicAssert.AreEqual(dateString, converter.ConvertFromExcel(typeConverterOptions, " " + dateTime + " ").ToString());

            // Empty conversions.
            dateTime = DateTime.MinValue;
            dateString = dateTime.ToString();
            ClassicAssert.AreEqual(dateString, converter.ConvertFromExcel(typeConverterOptions, " ").ToString());
            ClassicAssert.AreEqual(dateString, converter.ConvertFromExcel(typeConverterOptions, null).ToString());

            // Invalid conversions.
            try {
                converter.ConvertFromExcel(typeConverterOptions, 1);
                Assert.Fail();
            } catch (ExcelTypeConverterException) {
            }
        }

        [Test]
        public void ComponentModelCompatibilityTest()
        {
            var converter = new DateTimeConverter();
            var cmConverter = new System.ComponentModel.DateTimeConverter();

            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            var val = (DateTime)cmConverter.ConvertFromString("");
            ClassicAssert.AreEqual(DateTime.MinValue, val);

            val = (DateTime)converter.ConvertFromExcel(typeConverterOptions, "");
            ClassicAssert.AreEqual(DateTime.MinValue, val);

            val = (DateTime)converter.ConvertFromExcel(typeConverterOptions, null);
            ClassicAssert.AreEqual(DateTime.MinValue, val);

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
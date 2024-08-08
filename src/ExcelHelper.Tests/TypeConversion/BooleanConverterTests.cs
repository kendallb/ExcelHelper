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
using NUnit.Framework.Legacy;

namespace ExcelHelper.Tests.TypeConversion
{
    [TestFixture]
    public class BooleanConverterTests
    {
        [Test]
        public void PropertiesTest()
        {
            var converter = new BooleanConverter();
            ClassicAssert.AreEqual(false, converter.AcceptsNativeType);
            ClassicAssert.AreEqual(typeof(bool), converter.ConvertedType);
        }

        [Test]
        public void ConvertToExcelTest()
        {
            var converter = new BooleanConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            // Valid conversions with default formatting
            ClassicAssert.AreEqual("true", converter.ConvertToExcel(typeConverterOptions, true));
            ClassicAssert.AreEqual("false", converter.ConvertToExcel(typeConverterOptions, false));

            // Valid conversions with non-default formatting as numbers
            typeConverterOptions.BooleanTrueValues.Clear();
            typeConverterOptions.BooleanTrueValues.Add("1");
            typeConverterOptions.BooleanFalseValues.Clear();
            typeConverterOptions.BooleanFalseValues.Add("0");
            ClassicAssert.AreEqual(1, converter.ConvertToExcel(typeConverterOptions, true));
            ClassicAssert.AreEqual(0, converter.ConvertToExcel(typeConverterOptions, false));
        }

        [Test]
        public void ConvertFromExcelTest()
        {
            var converter = new BooleanConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture,
            };

            ClassicAssert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, true));
            ClassicAssert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, "true"));
            ClassicAssert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, "True"));
            ClassicAssert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, "TRUE"));
            ClassicAssert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, 1.0));
            ClassicAssert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, "1"));
            ClassicAssert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, "yes"));
            ClassicAssert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, "YES"));
            ClassicAssert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, "y"));
            ClassicAssert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, "Y"));
            ClassicAssert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, " true "));
            ClassicAssert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, " yes "));
            ClassicAssert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, " y "));

            ClassicAssert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, false));
            ClassicAssert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, "false"));
            ClassicAssert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, "False"));
            ClassicAssert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, "FALSE"));
            ClassicAssert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, 0.0));
            ClassicAssert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, "0"));
            ClassicAssert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, "no"));
            ClassicAssert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, "NO"));
            ClassicAssert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, "n"));
            ClassicAssert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, "N"));
            ClassicAssert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, " false "));
            ClassicAssert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, " 0 "));
            ClassicAssert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, " no "));
            ClassicAssert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, " n "));

            try {
                converter.ConvertFromExcel(typeConverterOptions, null);
                Assert.Fail();
            } catch (ExcelTypeConverterException) {
            }

            try {
                converter.ConvertFromExcel(typeConverterOptions, "unknown");
                Assert.Fail();
            } catch (ExcelTypeConverterException) {
            }

            // Make sure null converts properly if we allow blank
            typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture,
                BooleanFalseValues = { "", "no", "false", "n" },
            };
            ClassicAssert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, ""));
            ClassicAssert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, null));
        }
    }
}
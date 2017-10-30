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
    public class BooleanConverterTests
    {
        [TestMethod]
        public void PropertiesTest()
        {
            var converter = new BooleanConverter();
            Assert.AreEqual(false, converter.AcceptsNativeType);
            Assert.AreEqual(typeof(bool), converter.ConvertedType);
        }

        [TestMethod]
        public void ConvertToExcelTest()
        {
            var converter = new BooleanConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            // Valid conversions with default formatting
            Assert.AreEqual("true", converter.ConvertToExcel(typeConverterOptions, true));
            Assert.AreEqual("false", converter.ConvertToExcel(typeConverterOptions, false));

            // Valid conversions with non-default formatting as numbers
            typeConverterOptions.BooleanTrueValues.Clear();
            typeConverterOptions.BooleanTrueValues.Add("1");
            typeConverterOptions.BooleanFalseValues.Clear();
            typeConverterOptions.BooleanFalseValues.Add("0");
            Assert.AreEqual(1, converter.ConvertToExcel(typeConverterOptions, true));
            Assert.AreEqual(0, converter.ConvertToExcel(typeConverterOptions, false));
        }

        [TestMethod]
        public void ConvertFromExcelTest()
        {
            var converter = new BooleanConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture,
            };

            Assert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, true));
            Assert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, "true"));
            Assert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, "True"));
            Assert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, "TRUE"));
            Assert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, 1.0));
            Assert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, "1"));
            Assert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, "yes"));
            Assert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, "YES"));
            Assert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, "y"));
            Assert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, "Y"));
            Assert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, " true "));
            Assert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, " yes "));
            Assert.IsTrue((bool)converter.ConvertFromExcel(typeConverterOptions, " y "));

            Assert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, false));
            Assert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, "false"));
            Assert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, "False"));
            Assert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, "FALSE"));
            Assert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, 0.0));
            Assert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, "0"));
            Assert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, "no"));
            Assert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, "NO"));
            Assert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, "n"));
            Assert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, "N"));
            Assert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, " false "));
            Assert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, " 0 "));
            Assert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, " no "));
            Assert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, " n "));

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
            Assert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, ""));
            Assert.IsFalse((bool)converter.ConvertFromExcel(typeConverterOptions, null));
        }
    }
}
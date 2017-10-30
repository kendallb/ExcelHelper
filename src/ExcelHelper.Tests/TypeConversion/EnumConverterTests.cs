/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelHelper.TypeConversion;
// ReSharper disable ObjectCreationAsStatement

namespace ExcelHelper.Tests.TypeConversion
{
    [TestClass]
    public class EnumConverterTests
    {
        [TestMethod]
        public void ConstructorTest()
        {
            try {
                new EnumConverter(typeof(string));
                Assert.Fail();
            } catch (ArgumentException ex) {
                Assert.AreEqual("'System.String' is not an Enum.", ex.Message);
            }
        }

        [TestMethod]
        public void PropertiesTest()
        {
            var converter = new EnumConverter(typeof(TestEnum));
            Assert.AreEqual(false, converter.AcceptsNativeType);
            Assert.AreEqual(typeof(TestEnum), converter.ConvertedType);
        }

        [TestMethod]
        public void ConvertToExcelTest()
        {
            var converter = new EnumConverter(typeof(TestEnum));
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            Assert.AreEqual("None", converter.ConvertToExcel(typeConverterOptions, (TestEnum)0));
            Assert.AreEqual("None", converter.ConvertToExcel(typeConverterOptions, TestEnum.None));
            Assert.AreEqual("One", converter.ConvertToExcel(typeConverterOptions, (TestEnum)1));
            Assert.AreEqual("One", converter.ConvertToExcel(typeConverterOptions, TestEnum.One));
        }

        [TestMethod]
        public void ConvertFromExcelTest()
        {
            var converter = new EnumConverter(typeof(TestEnum));
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            Assert.AreEqual(TestEnum.One, converter.ConvertFromExcel(typeConverterOptions, "One"));
            Assert.AreEqual(TestEnum.One, converter.ConvertFromExcel(typeConverterOptions, "one"));
            Assert.AreEqual(TestEnum.One, converter.ConvertFromExcel(typeConverterOptions, "1"));
            try {
                Assert.AreEqual(TestEnum.One, converter.ConvertFromExcel(typeConverterOptions, ""));
                Assert.Fail();
            } catch (ExcelTypeConverterException) {
            }

            try {
                Assert.AreEqual(TestEnum.One, converter.ConvertFromExcel(typeConverterOptions, null));
                Assert.Fail();
            } catch (ExcelTypeConverterException) {
            }
        }

        private enum TestEnum
        {
            None = 0,
            One = 1,
        }
    }
}
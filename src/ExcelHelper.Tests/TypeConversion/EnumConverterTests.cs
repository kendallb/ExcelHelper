/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using System.Globalization;
using NUnit.Framework;
using ExcelHelper.TypeConversion;
using NUnit.Framework.Legacy;

// ReSharper disable ObjectCreationAsStatement

namespace ExcelHelper.Tests.TypeConversion
{
    [TestFixture]
    public class EnumConverterTests
    {
        [Test]
        public void ConstructorTest()
        {
            try {
                new EnumConverter(typeof(string));
                Assert.Fail();
            } catch (ArgumentException ex) {
                ClassicAssert.AreEqual("'System.String' is not an Enum.", ex.Message);
            }
        }

        [Test]
        public void PropertiesTest()
        {
            var converter = new EnumConverter(typeof(TestEnum));
            ClassicAssert.AreEqual(false, converter.AcceptsNativeType);
            ClassicAssert.AreEqual(typeof(TestEnum), converter.ConvertedType);
        }

        [Test]
        public void ConvertToExcelTest()
        {
            var converter = new EnumConverter(typeof(TestEnum));
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            ClassicAssert.AreEqual("None", converter.ConvertToExcel(typeConverterOptions, (TestEnum)0));
            ClassicAssert.AreEqual("None", converter.ConvertToExcel(typeConverterOptions, TestEnum.None));
            ClassicAssert.AreEqual("One", converter.ConvertToExcel(typeConverterOptions, (TestEnum)1));
            ClassicAssert.AreEqual("One", converter.ConvertToExcel(typeConverterOptions, TestEnum.One));
        }

        [Test]
        public void ConvertFromExcelTest()
        {
            var converter = new EnumConverter(typeof(TestEnum));
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            ClassicAssert.AreEqual(TestEnum.One, converter.ConvertFromExcel(typeConverterOptions, "One"));
            ClassicAssert.AreEqual(TestEnum.One, converter.ConvertFromExcel(typeConverterOptions, "one"));
            ClassicAssert.AreEqual(TestEnum.One, converter.ConvertFromExcel(typeConverterOptions, "1"));
            try {
                ClassicAssert.AreEqual(TestEnum.One, converter.ConvertFromExcel(typeConverterOptions, ""));
                Assert.Fail();
            } catch (ExcelTypeConverterException) {
            }

            try {
                ClassicAssert.AreEqual(TestEnum.One, converter.ConvertFromExcel(typeConverterOptions, null));
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
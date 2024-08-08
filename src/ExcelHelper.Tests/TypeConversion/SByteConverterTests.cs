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
using NUnit.Framework.Legacy;

namespace ExcelHelper.Tests.TypeConversion
{
    [TestFixture]
    public class SByteConverterTests
    {
        [Test]
        public void PropertiesTest()
        {
            var converter = new SByteConverter();
            ClassicAssert.AreEqual(true, converter.AcceptsNativeType);
            ClassicAssert.AreEqual(typeof(sbyte), converter.ConvertedType);
        }

        [Test]
        public void ConvertFromExcelTest()
        {
            var converter = new SByteConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            ClassicAssert.AreEqual((sbyte)123, converter.ConvertFromExcel(typeConverterOptions, (double)123));
            ClassicAssert.AreEqual((sbyte)123, converter.ConvertFromExcel(typeConverterOptions, "123"));
            ClassicAssert.AreEqual((sbyte)123, converter.ConvertFromExcel(typeConverterOptions, " 123 "));
            ClassicAssert.AreEqual((sbyte)0, converter.ConvertFromExcel(typeConverterOptions, null));

            typeConverterOptions.NumberStyle = NumberStyles.HexNumber;
            ClassicAssert.AreEqual((sbyte)0x12, converter.ConvertFromExcel(typeConverterOptions, "12"));

            try {
                converter.ConvertFromExcel(typeConverterOptions, "");
                Assert.Fail();
            } catch (ExcelTypeConverterException) {
            }
        }
    }
}
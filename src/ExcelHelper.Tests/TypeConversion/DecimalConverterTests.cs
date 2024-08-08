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
    public class DecimalConverterTests
    {
        [Test]
        public void PropertiesTest()
        {
            var converter = new DecimalConverter();
            ClassicAssert.AreEqual(true, converter.AcceptsNativeType);
            ClassicAssert.AreEqual(typeof(decimal), converter.ConvertedType);
        }

        [Test]
        public void ConvertFromExcelTest()
        {
            var converter = new DecimalConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            ClassicAssert.AreEqual(12.3m, converter.ConvertFromExcel(typeConverterOptions, 12.3));
            ClassicAssert.AreEqual(12.3m, converter.ConvertFromExcel(typeConverterOptions, "12.3"));
            ClassicAssert.AreEqual(12.3m, converter.ConvertFromExcel(typeConverterOptions, " 12.3 "));
            ClassicAssert.AreEqual((decimal)0, converter.ConvertFromExcel(typeConverterOptions, null));

            try {
                converter.ConvertFromExcel(typeConverterOptions, "");
                Assert.Fail();
            } catch (ExcelTypeConverterException) {
            }
        }
    }
}
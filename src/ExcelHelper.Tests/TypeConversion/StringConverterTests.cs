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
    public class StringConverterTests
    {
        [Test]
        public void PropertiesTest()
        {
            var converter = new StringConverter();
            ClassicAssert.AreEqual(true, converter.AcceptsNativeType);
            ClassicAssert.AreEqual(typeof(string), converter.ConvertedType);
        }

        [Test]
        public void ConvertToExcelTest()
        {
            var converter = new StringConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture,
            };

            ClassicAssert.AreEqual("123", converter.ConvertToExcel(typeConverterOptions, "123"));
            ClassicAssert.AreEqual(null, converter.ConvertToExcel(typeConverterOptions, null));
        }

        [Test]
        public void ConvertFromExcelTest()
        {
            var converter = new StringConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture,
            };

            ClassicAssert.AreEqual("12.3", converter.ConvertFromExcel(typeConverterOptions, 12.3));
            ClassicAssert.AreEqual("123", converter.ConvertFromExcel(typeConverterOptions, "123"));
            ClassicAssert.AreEqual("123", converter.ConvertFromExcel(typeConverterOptions, " 123 "));
            ClassicAssert.AreEqual("", converter.ConvertFromExcel(typeConverterOptions, null));
        }
    }
}
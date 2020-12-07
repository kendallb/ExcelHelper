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

namespace ExcelHelper.Tests.TypeConversion
{
    [TestFixture]
    public class NullableConverterTests
    {
        [Test]
        public void PropertiesTest()
        {
            var converter = new NullableConverter(typeof(decimal?));
            Assert.AreEqual(true, converter.AcceptsNativeType);
            Assert.AreEqual(typeof(decimal?), converter.ConvertedType);
        }

        [Test]
        public void ConvertToExcelTest()
        {
            var converter = new NullableConverter(typeof(decimal?));
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            Assert.AreEqual(12.3m, converter.ConvertToExcel(typeConverterOptions, (decimal?)12.3m));
            Assert.AreEqual(null, converter.ConvertToExcel(typeConverterOptions, null));
        }

        [Test]
        public void ConvertFromExcelTest()
        {
            var converter = new NullableConverter(typeof(decimal?));
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            Assert.AreEqual((decimal?)12.3m, converter.ConvertFromExcel(typeConverterOptions, 12.3));
            Assert.AreEqual((decimal?)12.3m, converter.ConvertFromExcel(typeConverterOptions, "12.3"));
            Assert.AreEqual((decimal?)12.3m, converter.ConvertFromExcel(typeConverterOptions, "12.3"));
            Assert.AreEqual((decimal?)12.3m, converter.ConvertFromExcel(typeConverterOptions, " 12.3 "));
            Assert.AreEqual(null, converter.ConvertFromExcel(typeConverterOptions, ""));
        }
    }
}
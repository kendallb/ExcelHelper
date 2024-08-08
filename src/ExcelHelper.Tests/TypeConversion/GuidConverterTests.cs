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
    public class GuidConverterTests
    {
        [Test]
        public void PropertiesTest()
        {
            var converter = new GuidConverter();
            ClassicAssert.AreEqual(false, converter.AcceptsNativeType);
            ClassicAssert.AreEqual(typeof(Guid), converter.ConvertedType);
        }

        [Test]
        public void ConvertToExcelTest()
        {
            var converter = new GuidConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            ClassicAssert.AreEqual("79f1a554-babd-41a1-8caf-185250a1fc21", converter.ConvertToExcel(typeConverterOptions, new Guid("79f1a554-babd-41a1-8caf-185250a1fc21")));
            ClassicAssert.AreEqual(null, converter.ConvertToExcel(typeConverterOptions, null));
        }

        [Test]
        public void ConvertFromExcelTest()
        {
            var converter = new GuidConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            var guid = new Guid("79f1a554-babd-41a1-8caf-185250a1fc21");
            ClassicAssert.AreEqual(guid, converter.ConvertFromExcel(typeConverterOptions, "79f1a554-babd-41a1-8caf-185250a1fc21"));
            ClassicAssert.AreEqual(guid, converter.ConvertFromExcel(typeConverterOptions, " 79f1a554-babd-41a1-8caf-185250a1fc21 "));

            try {
                converter.ConvertFromExcel(typeConverterOptions, null);
                Assert.Fail();
            } catch (ExcelTypeConverterException) {
            }
        }
    }
}
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
    public class ByteConverterTests
    {
        [TestMethod]
        public void PropertiesTest()
        {
            var converter = new ByteConverter();
            Assert.AreEqual(true, converter.AcceptsNativeType);
            Assert.AreEqual(typeof(byte), converter.ConvertedType);
        }
        
        [TestMethod]
        public void ConvertFromExcelTest()
        {
            var converter = new ByteConverter();
            var typeConverterOptions = new TypeConverterOptions {
                CultureInfo = CultureInfo.CurrentCulture
            };

            Assert.AreEqual((byte)123, converter.ConvertFromExcel(typeConverterOptions, (double)123));
            Assert.AreEqual((byte)123, converter.ConvertFromExcel(typeConverterOptions, "123"));
            Assert.AreEqual((byte)123, converter.ConvertFromExcel(typeConverterOptions, " 123 "));
            Assert.AreEqual((byte)0, converter.ConvertFromExcel(typeConverterOptions, null));

            typeConverterOptions.NumberStyle = NumberStyles.HexNumber;
            Assert.AreEqual((byte)0x12, converter.ConvertFromExcel(typeConverterOptions, "12"));

            try {
                converter.ConvertFromExcel(typeConverterOptions, "");
                Assert.Fail();
            } catch (ExcelTypeConverterException) {
            }
        }
    }
}
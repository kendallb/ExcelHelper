/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using ExcelHelper.TypeConversion;
using NUnit.Framework;
using NUnit.Framework.Legacy;

namespace ExcelHelper.Tests.TypeConversion
{
    [TestFixture]
    public class DefaultTypeConverterTests
    {
        [Test]
        public void PropertiesTest()
        {
            var converter = new TestConverter();
            ClassicAssert.AreEqual(true, converter.AcceptsNativeType);
            ClassicAssert.AreEqual(typeof(double), converter.ConvertedType);
        }

        private class TestConverter : DefaultTypeConverter
        {
            public TestConverter()
                : base(true, typeof(double))
            {
            }
        }
    }
}
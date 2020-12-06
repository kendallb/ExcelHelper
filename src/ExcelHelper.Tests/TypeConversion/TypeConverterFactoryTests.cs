/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using ExcelHelper.TypeConversion;
using Microsoft.VisualStudio.TestTools.UnitTesting;
// ReSharper disable UnusedMember.Local

namespace ExcelHelper.Tests.TypeConversion
{
    [TestClass]
    public class TypeConverterFactoryTests
    {
        [TestMethod]
        public void GetConverterForUnknownTypeTest()
        {
            try {
                TypeConverterFactory.GetConverter(typeof(TestUnknownClass));
            } catch (ExcelTypeConverterException) {
            }
        }

        //[TestMethod]
        //public void GetConverterForKnownTypeTest()
        //{
        //    try {
        //        TypeConverterFactory.GetConverter<TestKnownClass>();
        //    } catch (ExcelTypeConverterException) {
        //    }

        //    TypeConverterFactory.AddConverter<TestKnownClass>(new TestKnownConverter());
        //    var converter = TypeConverterFactory.GetConverter<TestKnownClass>();

        //    Assert.IsInstanceOfType(converter, typeof(TestKnownConverter));
        //}

        //[TestMethod]
        //public void RemoveConverterForUnknownTypeTest()
        //{
        //    TypeConverterFactory.RemoveConverter<TestUnknownClass>();
        //    TypeConverterFactory.RemoveConverter(typeof(TestUnknownClass));
        //}

        [TestMethod]
        public void GetConverterForBooleanTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(bool));

            Assert.IsInstanceOfType(converter, typeof(BooleanConverter));
            Assert.IsFalse(converter.AcceptsNativeType);
        }

        [TestMethod]
        public void GetConverterForByteTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(byte));

            Assert.IsInstanceOfType(converter, typeof(ByteConverter));
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [TestMethod]
        public void GetConverterForCharTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(char));

            Assert.IsInstanceOfType(converter, typeof(CharConverter));
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [TestMethod]
        public void GetConverterForDateTimeTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(DateTime));

            Assert.IsInstanceOfType(converter, typeof(DateTimeConverter));
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [TestMethod]
        public void GetConverterForDecimalTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(decimal));

            Assert.IsInstanceOfType(converter, typeof(DecimalConverter));
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [TestMethod]
        public void GetConverterForDoubleTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(double));

            Assert.IsInstanceOfType(converter, typeof(DoubleConverter));
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [TestMethod]
        public void GetConverterForFloatTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(float));

            Assert.IsInstanceOfType(converter, typeof(SingleConverter));
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [TestMethod]
        public void GetConverterForGuidTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(Guid));

            Assert.IsInstanceOfType(converter, typeof(GuidConverter));
            Assert.IsFalse(converter.AcceptsNativeType);
        }

        [TestMethod]
        public void GetConverterForInt16Test()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(short));

            Assert.IsInstanceOfType(converter, typeof(Int16Converter));
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [TestMethod]
        public void GetConverterForInt32Test()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(int));

            Assert.IsInstanceOfType(converter, typeof(Int32Converter));
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [TestMethod]
        public void GetConverterForInt64Test()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(long));

            Assert.IsInstanceOfType(converter, typeof(Int64Converter));
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [TestMethod]
        public void GetConverterForNullableTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(int?));

            Assert.IsInstanceOfType(converter, typeof(NullableConverter));
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [TestMethod]
        public void GetConverterForSByteTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(sbyte));

            Assert.IsInstanceOfType(converter, typeof(SByteConverter));
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [TestMethod]
        public void GetConverterForStringTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(string));

            Assert.IsInstanceOfType(converter, typeof(StringConverter));
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [TestMethod]
        public void GetConverterForTimeSpanTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(TimeSpan));

            Assert.IsInstanceOfType(converter, typeof(TimeSpanConverter));
            Assert.IsFalse(converter.AcceptsNativeType);
        }

        [TestMethod]
        public void GetConverterForUInt16Test()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(ushort));

            Assert.IsInstanceOfType(converter, typeof(UInt16Converter));
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [TestMethod]
        public void GetConverterForUInt32Test()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(uint));

            Assert.IsInstanceOfType(converter, typeof(UInt32Converter));
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [TestMethod]
        public void GetConverterForUInt64Test()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(ulong));

            Assert.IsInstanceOfType(converter, typeof(UInt64Converter));
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [TestMethod]
        public void GetConverterForEnumTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(TestEnum));

            Assert.IsInstanceOfType(converter, typeof(EnumConverter));
            Assert.IsFalse(converter.AcceptsNativeType);
        }

        [TestMethod]
        public void GetConverterForEnumerableTypesTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(IEnumerable));
            Assert.IsInstanceOfType(converter, typeof(EnumerableConverter));

            converter = TypeConverterFactory.GetConverter(typeof(IList));
            Assert.IsInstanceOfType(converter, typeof(EnumerableConverter));

            converter = TypeConverterFactory.GetConverter(typeof(List<int>));
            Assert.IsInstanceOfType(converter, typeof(EnumerableConverter));

            converter = TypeConverterFactory.GetConverter(typeof(ICollection));
            Assert.IsInstanceOfType(converter, typeof(EnumerableConverter));

            converter = TypeConverterFactory.GetConverter(typeof(Collection<int>));
            Assert.IsInstanceOfType(converter, typeof(EnumerableConverter));

            converter = TypeConverterFactory.GetConverter(typeof(IDictionary));
            Assert.IsInstanceOfType(converter, typeof(EnumerableConverter));

            converter = TypeConverterFactory.GetConverter(typeof(Dictionary<int, string>));
            Assert.IsInstanceOfType(converter, typeof(EnumerableConverter));

            converter = TypeConverterFactory.GetConverter(typeof(Array));
            Assert.IsInstanceOfType(converter, typeof(EnumerableConverter));
        }

        //[TestMethod]
        //public void GetConverterForCustomListConverterThatIsNotEnumerableConverterTest()
        //{
        //    TypeConverterFactory.AddConverter<List<string>>(new TestListConverter());
        //    var converter = TypeConverterFactory.GetConverter(typeof(List<string>));
        //    Assert.IsInstanceOfType(converter, typeof(TestListConverter));

        //    converter = TypeConverterFactory.GetConverter(typeof(List<int>));
        //    Assert.IsInstanceOfType(converter, typeof(EnumerableConverter));

        //    converter = TypeConverterFactory.GetConverter(typeof(Array));
        //    Assert.IsInstanceOfType(converter, typeof(EnumerableConverter));
        //}

        private class TestListConverter : DefaultTypeConverter
        {
            public TestListConverter()
                : base(false, typeof(object))
            {
            }
        }

        private class TestUnknownClass
        {
        }

        private class TestKnownClass
        {
        }

        private class TestKnownConverter : DefaultTypeConverter
        {
            public TestKnownConverter()
                : base(false, typeof(object))
            {
            }
        }

        private enum TestEnum
        {
        }
    }
}
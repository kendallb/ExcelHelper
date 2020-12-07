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
using NUnit.Framework;
// ReSharper disable UnusedMember.Local

namespace ExcelHelper.Tests.TypeConversion
{
    [TestFixture]
    public class TypeConverterFactoryTests
    {
        [Test]
        public void GetConverterForUnknownTypeTest()
        {
            try {
                TypeConverterFactory.GetConverter(typeof(TestUnknownClass));
            } catch (ExcelTypeConverterException) {
            }
        }

        //[Test]
        //public void GetConverterForKnownTypeTest()
        //{
        //    try {
        //        TypeConverterFactory.GetConverter<TestKnownClass>();
        //    } catch (ExcelTypeConverterException) {
        //    }

        //    TypeConverterFactory.AddConverter<TestKnownClass>(new TestKnownConverter());
        //    var converter = TypeConverterFactory.GetConverter<TestKnownClass>();

        //    Assert.IsInstanceOf<TestKnownConverter>(converter);
        //}

        //[Test]
        //public void RemoveConverterForUnknownTypeTest()
        //{
        //    TypeConverterFactory.RemoveConverter<TestUnknownClass>();
        //    TypeConverterFactory.RemoveConverter(typeof(TestUnknownClass));
        //}

        [Test]
        public void GetConverterForBooleanTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(bool));

            Assert.IsInstanceOf<BooleanConverter>(converter);
            Assert.IsFalse(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForByteTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(byte));

            Assert.IsInstanceOf<ByteConverter>(converter);
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForCharTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(char));

            Assert.IsInstanceOf<CharConverter>(converter);
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForDateTimeTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(DateTime));

            Assert.IsInstanceOf<DateTimeConverter>(converter);
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForDecimalTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(decimal));

            Assert.IsInstanceOf<DecimalConverter>(converter);
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForDoubleTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(double));

            Assert.IsInstanceOf<DoubleConverter>(converter);
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForFloatTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(float));

            Assert.IsInstanceOf<SingleConverter>(converter);
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForGuidTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(Guid));

            Assert.IsInstanceOf<GuidConverter>(converter);
            Assert.IsFalse(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForInt16Test()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(short));

            Assert.IsInstanceOf<Int16Converter>(converter);
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForInt32Test()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(int));

            Assert.IsInstanceOf<Int32Converter>(converter);
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForInt64Test()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(long));

            Assert.IsInstanceOf<Int64Converter>(converter);
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForNullableTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(int?));

            Assert.IsInstanceOf<NullableConverter>(converter);
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForSByteTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(sbyte));

            Assert.IsInstanceOf<SByteConverter>(converter);
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForStringTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(string));

            Assert.IsInstanceOf<StringConverter>(converter);
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForTimeSpanTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(TimeSpan));

            Assert.IsInstanceOf<TimeSpanConverter>(converter);
            Assert.IsFalse(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForUInt16Test()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(ushort));

            Assert.IsInstanceOf<UInt16Converter>(converter);
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForUInt32Test()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(uint));

            Assert.IsInstanceOf<UInt32Converter>(converter);
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForUInt64Test()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(ulong));

            Assert.IsInstanceOf<UInt64Converter>(converter);
            Assert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForEnumTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(TestEnum));

            Assert.IsInstanceOf<EnumConverter>(converter);
            Assert.IsFalse(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForEnumerableTypesTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(IEnumerable));
            Assert.IsInstanceOf<EnumerableConverter>(converter);

            converter = TypeConverterFactory.GetConverter(typeof(IList));
            Assert.IsInstanceOf<EnumerableConverter>(converter);

            converter = TypeConverterFactory.GetConverter(typeof(List<int>));
            Assert.IsInstanceOf<EnumerableConverter>(converter);

            converter = TypeConverterFactory.GetConverter(typeof(ICollection));
            Assert.IsInstanceOf<EnumerableConverter>(converter);

            converter = TypeConverterFactory.GetConverter(typeof(Collection<int>));
            Assert.IsInstanceOf<EnumerableConverter>(converter);

            converter = TypeConverterFactory.GetConverter(typeof(IDictionary));
            Assert.IsInstanceOf<EnumerableConverter>(converter);

            converter = TypeConverterFactory.GetConverter(typeof(Dictionary<int, string>));
            Assert.IsInstanceOf<EnumerableConverter>(converter);

            converter = TypeConverterFactory.GetConverter(typeof(Array));
            Assert.IsInstanceOf<EnumerableConverter>(converter);
        }

        //[Test]
        //public void GetConverterForCustomListConverterThatIsNotEnumerableConverterTest()
        //{
        //    TypeConverterFactory.AddConverter<List<string>>(new TestListConverter());
        //    var converter = TypeConverterFactory.GetConverter(typeof(List<string>));
        //    Assert.IsInstanceOf<TestListConverter>(converter);

        //    converter = TypeConverterFactory.GetConverter(typeof(List<int>));
        //    Assert.IsInstanceOf<EnumerableConverter>(converter);

        //    converter = TypeConverterFactory.GetConverter(typeof(Array));
        //    Assert.IsInstanceOf<EnumerableConverter>(converter);
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
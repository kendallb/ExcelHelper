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
using NUnit.Framework.Legacy;

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

        //    ClassicAssert.IsInstanceOf<TestKnownConverter>(converter);
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

            ClassicAssert.IsInstanceOf<BooleanConverter>(converter);
            ClassicAssert.IsFalse(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForByteTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(byte));

            ClassicAssert.IsInstanceOf<ByteConverter>(converter);
            ClassicAssert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForCharTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(char));

            ClassicAssert.IsInstanceOf<CharConverter>(converter);
            ClassicAssert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForDateTimeTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(DateTime));

            ClassicAssert.IsInstanceOf<DateTimeConverter>(converter);
            ClassicAssert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForDecimalTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(decimal));

            ClassicAssert.IsInstanceOf<DecimalConverter>(converter);
            ClassicAssert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForDoubleTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(double));

            ClassicAssert.IsInstanceOf<DoubleConverter>(converter);
            ClassicAssert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForFloatTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(float));

            ClassicAssert.IsInstanceOf<SingleConverter>(converter);
            ClassicAssert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForGuidTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(Guid));

            ClassicAssert.IsInstanceOf<GuidConverter>(converter);
            ClassicAssert.IsFalse(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForInt16Test()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(short));

            ClassicAssert.IsInstanceOf<Int16Converter>(converter);
            ClassicAssert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForInt32Test()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(int));

            ClassicAssert.IsInstanceOf<Int32Converter>(converter);
            ClassicAssert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForInt64Test()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(long));

            ClassicAssert.IsInstanceOf<Int64Converter>(converter);
            ClassicAssert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForNullableTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(int?));

            ClassicAssert.IsInstanceOf<NullableConverter>(converter);
            ClassicAssert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForSByteTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(sbyte));

            ClassicAssert.IsInstanceOf<SByteConverter>(converter);
            ClassicAssert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForStringTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(string));

            ClassicAssert.IsInstanceOf<StringConverter>(converter);
            ClassicAssert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForTimeSpanTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(TimeSpan));

            ClassicAssert.IsInstanceOf<TimeSpanConverter>(converter);
            ClassicAssert.IsFalse(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForUInt16Test()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(ushort));

            ClassicAssert.IsInstanceOf<UInt16Converter>(converter);
            ClassicAssert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForUInt32Test()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(uint));

            ClassicAssert.IsInstanceOf<UInt32Converter>(converter);
            ClassicAssert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForUInt64Test()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(ulong));

            ClassicAssert.IsInstanceOf<UInt64Converter>(converter);
            ClassicAssert.IsTrue(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForEnumTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(TestEnum));

            ClassicAssert.IsInstanceOf<EnumConverter>(converter);
            ClassicAssert.IsFalse(converter.AcceptsNativeType);
        }

        [Test]
        public void GetConverterForEnumerableTypesTest()
        {
            var converter = TypeConverterFactory.GetConverter(typeof(IEnumerable));
            ClassicAssert.IsInstanceOf<EnumerableConverter>(converter);

            converter = TypeConverterFactory.GetConverter(typeof(IList));
            ClassicAssert.IsInstanceOf<EnumerableConverter>(converter);

            converter = TypeConverterFactory.GetConverter(typeof(List<int>));
            ClassicAssert.IsInstanceOf<EnumerableConverter>(converter);

            converter = TypeConverterFactory.GetConverter(typeof(ICollection));
            ClassicAssert.IsInstanceOf<EnumerableConverter>(converter);

            converter = TypeConverterFactory.GetConverter(typeof(Collection<int>));
            ClassicAssert.IsInstanceOf<EnumerableConverter>(converter);

            converter = TypeConverterFactory.GetConverter(typeof(IDictionary));
            ClassicAssert.IsInstanceOf<EnumerableConverter>(converter);

            converter = TypeConverterFactory.GetConverter(typeof(Dictionary<int, string>));
            ClassicAssert.IsInstanceOf<EnumerableConverter>(converter);

            converter = TypeConverterFactory.GetConverter(typeof(Array));
            ClassicAssert.IsInstanceOf<EnumerableConverter>(converter);
        }

        //[Test]
        //public void GetConverterForCustomListConverterThatIsNotEnumerableConverterTest()
        //{
        //    TypeConverterFactory.AddConverter<List<string>>(new TestListConverter());
        //    var converter = TypeConverterFactory.GetConverter(typeof(List<string>));
        //    ClassicAssert.IsInstanceOf<TestListConverter>(converter);

        //    converter = TypeConverterFactory.GetConverter(typeof(List<int>));
        //    ClassicAssert.IsInstanceOf<EnumerableConverter>(converter);

        //    converter = TypeConverterFactory.GetConverter(typeof(Array));
        //    ClassicAssert.IsInstanceOf<EnumerableConverter>(converter);
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
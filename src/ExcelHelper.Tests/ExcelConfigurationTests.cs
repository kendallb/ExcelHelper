/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using ExcelHelper.Configuration;
using NUnit.Framework;
// ReSharper disable UnusedAutoPropertyAccessor.Local

namespace ExcelHelper.Tests
{
    [TestFixture]
    public class ExcelConfigurationTests
    {
        [Test]
        public void AddingMappingsWithGenericMethod1Test()
        {
            var config = new ExcelConfiguration();
            config.RegisterClassMap<TestClassMappings>();

            Assert.AreEqual(2, config.Maps[typeof(TestClass)].PropertyMaps.Count);
        }

        [Test]
        public void AddingMappingsWithGenericMethod2Test()
        {
            var config = new ExcelConfiguration();
            config.RegisterClassMap<TestClassMappings>();

            Assert.AreEqual(2, config.Maps[typeof(TestClass)].PropertyMaps.Count);
        }

        [Test]
        public void AddingMappingsWithNonGenericMethodTest()
        {
            var config = new ExcelConfiguration();
            config.RegisterClassMap(typeof(TestClassMappings));

            Assert.AreEqual(2, config.Maps[typeof(TestClass)].PropertyMaps.Count);
        }

        [Test]
        public void AddingMappingsWithInstanceMethodTest()
        {
            var config = new ExcelConfiguration();
            config.RegisterClassMap(new TestClassMappings());

            Assert.AreEqual(2, config.Maps[typeof(TestClass)].PropertyMaps.Count);
        }

        [Test]
        public void RegisterClassMapGenericTest()
        {
            var config = new ExcelConfiguration();

            Assert.IsNull(config.Maps[typeof(TestClass)]);
            config.RegisterClassMap<TestClassMappings>();
            Assert.IsNotNull(config.Maps[typeof(TestClass)]);
        }

        [Test]
        public void RegisterClassMapNonGenericTest()
        {
            var config = new ExcelConfiguration();

            Assert.IsNull(config.Maps[typeof(TestClass)]);
            config.RegisterClassMap(typeof(TestClassMappings));
            Assert.IsNotNull(config.Maps[typeof(TestClass)]);
        }

        [Test]
        public void RegisterClassInstanceTest()
        {
            var config = new ExcelConfiguration();

            Assert.IsNull(config.Maps[typeof(TestClass)]);
            config.RegisterClassMap(new TestClassMappings());
            Assert.IsNotNull(config.Maps[typeof(TestClass)]);
        }

        [Test]
        public void UnregisterClassMapGenericTest()
        {
            var config = new ExcelConfiguration();

            Assert.IsNull(config.Maps[typeof(TestClass)]);
            config.RegisterClassMap<TestClassMappings>();
            Assert.IsNotNull(config.Maps[typeof(TestClass)]);

            config.UnregisterClassMap<TestClassMappings>();
            Assert.IsNull(config.Maps[typeof(TestClass)]);
        }

        [Test]
        public void UnregisterClassNonMapGenericTest()
        {
            var config = new ExcelConfiguration();

            Assert.IsNull(config.Maps[typeof(TestClass)]);
            config.RegisterClassMap(typeof(TestClassMappings));
            Assert.IsNotNull(config.Maps[typeof(TestClass)]);

            config.UnregisterClassMap(typeof(TestClassMappings));
            Assert.IsNull(config.Maps[typeof(TestClass)]);
        }

        [Test]
        public void AddingMappingsWithNonGenericMethodThrowsWhenNotAExcelClassMap()
        {
            try {
                new ExcelConfiguration().RegisterClassMap(typeof(TestClass));
                Assert.Fail();
            } catch (ArgumentException) {
            }
        }

        private class TestClass
        {
            public string StringColumn { get; set; }
            public int IntColumn { get; set; }
        }

        private sealed class TestClassMappings : ExcelClassMap<TestClass>
        {
            public TestClassMappings()
            {
                Map(c => c.StringColumn);
                Map(c => c.IntColumn);
            }
        }
    }
}
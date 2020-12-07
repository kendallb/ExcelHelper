/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using NUnit.Framework;

namespace ExcelHelper.Tests
{
    [TestFixture]
    public class ReflectionHelperTests
    {
        [Test]
        public void CreateInstanceTests()
        {
            var test = ReflectionHelper.CreateInstance<Test>();

            Assert.IsNotNull(test);
            Assert.AreEqual("name", Test.Name);

            test = (Test)ReflectionHelper.CreateInstance(typeof(Test));
            Assert.IsNotNull(test);
            Assert.AreEqual("name", Test.Name);
        }

        private class Test
        {
            public static string Name => "name";
        }
    }
}
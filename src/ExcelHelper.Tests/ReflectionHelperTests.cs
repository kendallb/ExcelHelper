/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelHelper.Tests
{
    [TestClass]
    public class ReflectionHelperTests
    {
        [TestMethod]
        public void CreateInstanceTests()
        {
            var test = ReflectionHelper.CreateInstance<Test>();

            Assert.IsNotNull(test);
            Assert.AreEqual("name", test.Name);

            test = (Test)ReflectionHelper.CreateInstance(typeof(Test));
            Assert.IsNotNull(test);
            Assert.AreEqual("name", test.Name);
        }

        private class Test
        {
            public string Name => "name";
        }
    }
}
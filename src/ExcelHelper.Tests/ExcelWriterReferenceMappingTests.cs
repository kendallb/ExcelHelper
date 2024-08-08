/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using ExcelHelper.Configuration;
using NUnit.Framework;
using NUnit.Framework.Legacy;

// ReSharper disable ClassNeverInstantiated.Local

namespace ExcelHelper.Tests
{
    [TestFixture]
    public class ExcelWriterReferenceMappingTests
    {
        [Test]
        public void NestedReferencesTest()
        {
            var records = new List<A>();
            for (var i = 0; i < 2; i++) {
                var row = i + 1;
                records.Add(
                    new A {
                        Id = "a" + row,
                        B = new B {
                            Id = "b" + row,
                            C = new C {
                                Id = "c" + row,
                                D = new D {
                                    Id = "d" + row
                                }
                            }
                        }
                    });
            }

            using (var stream = new MemoryStream()) {
                using (var excel = new ExcelWriter(stream)) {
                    excel.Configuration.RegisterClassMap<AMap>();
                    excel.WriteRecords(records);
                    excel.Close();

                    stream.Position = 0;
                    using (var book = new XLWorkbook(stream)) {
                        var sheet = book.Worksheets.Worksheet(1);

                        // Check the header row
                        ClassicAssert.AreEqual("AId", sheet.Cell(1, 1).Value);
                        ClassicAssert.AreEqual("BId", sheet.Cell(1, 2).Value);
                        ClassicAssert.AreEqual("CId", sheet.Cell(1, 3).Value);
                        ClassicAssert.AreEqual("DId", sheet.Cell(1, 4).Value);

                        // Check the first record
                        ClassicAssert.AreEqual("a1", sheet.Cell(2, 1).Value);
                        ClassicAssert.AreEqual("b1", sheet.Cell(2, 2).Value);
                        ClassicAssert.AreEqual("c1", sheet.Cell(2, 3).Value);
                        ClassicAssert.AreEqual("d1", sheet.Cell(2, 4).Value);

                        // Check the second record
                        ClassicAssert.AreEqual("a2", sheet.Cell(3, 1).Value);
                        ClassicAssert.AreEqual("b2", sheet.Cell(3, 2).Value);
                        ClassicAssert.AreEqual("c2", sheet.Cell(3, 3).Value);
                        ClassicAssert.AreEqual("d2", sheet.Cell(3, 4).Value);
                    }
                }
            }
        }

        [Test]
        public void NullReferenceTest()
        {
            var records = new List<A> {
                new A {
                    Id = "1",
                }
            };

            using (var stream = new MemoryStream()) {
                using (var excel = new ExcelWriter(stream)) {
                    excel.Configuration.RegisterClassMap<AMap>();
                    excel.WriteRecords(records);
                    excel.Close();

                    stream.Position = 0;
                    using (var book = new XLWorkbook(stream)) {
                        var sheet = book.Worksheets.Worksheet(1);

                        // Check the header row
                        ClassicAssert.AreEqual("AId", sheet.Cell(1, 1).Value);
                        ClassicAssert.AreEqual("BId", sheet.Cell(1, 2).Value);
                        ClassicAssert.AreEqual("CId", sheet.Cell(1, 3).Value);
                        ClassicAssert.AreEqual("DId", sheet.Cell(1, 4).Value);

                        // Check the first record
                        ClassicAssert.AreEqual("1", sheet.Cell(2, 1).Value);
                        ClassicAssert.AreEqual("", sheet.Cell(2, 2).Value);
                        ClassicAssert.AreEqual("", sheet.Cell(2, 3).Value);
                        ClassicAssert.AreEqual("", sheet.Cell(2, 4).Value);
                    }
                }
            }
        }

        private class A
        {
            public string Id { get; set; }
            public B B { get; set; }
        }

        private class B
        {
            public string Id { get; set; }
            public C C { get; set; }
        }

        private class C
        {
            public string Id { get; set; }
            public D D { get; set; }
        }

        private class D
        {
            public string Id { get; set; }
        }

        private sealed class AMap : ExcelClassMap<A>
        {
            public AMap()
            {
                Map(m => m.Id).Name("AId");
                References<BMap>(m => m.B);
            }
        }

        private sealed class BMap : ExcelClassMap<B>
        {
            public BMap()
            {
                Map(m => m.Id).Name("BId");
                References<CMap>(m => m.C);
            }
        }

        private sealed class CMap : ExcelClassMap<C>
        {
            public CMap()
            {
                Map(m => m.Id).Name("CId");
                References<DMap>(m => m.D);
            }
        }

        private sealed class DMap : ExcelClassMap<D>
        {
            public DMap()
            {
                Map(m => m.Id).Name("DId");
            }
        }
    }
}
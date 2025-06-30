/*
 * Copyright (C) 2004-2013 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System.IO;
using System.Linq;
using C1.Excel;
using NUnit.Framework;
using ExcelHelper.Configuration;
using NUnit.Framework.Legacy;

// ReSharper disable ClassNeverInstantiated.Local
// ReSharper disable UnusedAutoPropertyAccessor.Local

namespace ExcelHelper.Tests
{
    [TestFixture]
    public class ExcelReaderReferenceMappingTests
    {
        private ExcelFactory _factory;

        [SetUp]
        public void SetUp()
        {
            _factory = new ExcelFactory();
        }

        [Test]
        public void NestedReferencesClassMappingTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];

                    // Write the header fields
                    sheet[0, 0].Value = "AId";
                    sheet[0, 1].Value = "BId";
                    sheet[0, 2].Value = "CId";
                    sheet[0, 3].Value = "DId";

                    // Write out the first record
                    sheet[1, 0].Value = "a1";
                    sheet[1, 1].Value = "b1";
                    sheet[1, 2].Value = "c1";
                    sheet[1, 3].Value = "d1";

                    // Write out the second record
                    sheet[2, 0].Value = "a2";
                    sheet[2, 1].Value = "b2";
                    sheet[2, 2].Value = "c2";
                    sheet[2, 3].Value = "d2";

                    // Save it to the stream
                    book.Save(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = _factory.CreateReader(stream)) {
                    excel.Configuration.RegisterClassMap<AMap>();
                    var records = excel.GetRecords<A>().ToList();

                    // Make sure we got our records
                    ClassicAssert.AreEqual(2, records.Count);
                    for (var i = 0; i < records.Count; i++) {
                        var rowId = i + 1;
                        var row = records[i];
                        ClassicAssert.AreEqual("a" + rowId, row.Id);
                        ClassicAssert.AreEqual("b" + rowId, row.B.Id);
                        ClassicAssert.AreEqual("c" + rowId, row.B.C.Id);
                        ClassicAssert.AreEqual("d" + rowId, row.B.C.D.Id);
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
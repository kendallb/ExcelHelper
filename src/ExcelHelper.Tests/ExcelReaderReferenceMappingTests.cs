/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 */

using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelHelper.Configuration;
// ReSharper disable ClassNeverInstantiated.Local
// ReSharper disable UnusedAutoPropertyAccessor.Local

namespace ExcelHelper.Tests
{
    [TestClass]
    public class ExcelReaderReferenceMappingTests
    {
        private ExcelFactory _factory;

        [TestInitialize]
        public void SetUp()
        {
            _factory = new ExcelFactory();
        }

        [TestMethod]
        public void NestedReferencesClassMappingTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");

                    // Write the header fields
                    sheet.Cell(1, 1).SetValue("AId");
                    sheet.Cell(1, 2).SetValue("BId");
                    sheet.Cell(1, 3).SetValue("CId");
                    sheet.Cell(1, 4).SetValue("DId");

                    // Write out the first record
                    sheet.Cell(2, 1).SetValue("a1");
                    sheet.Cell(2, 2).SetValue("b1");
                    sheet.Cell(2, 3).SetValue("c1");
                    sheet.Cell(2, 4).SetValue("d1");

                    // Write out the second record
                    sheet.Cell(3, 1).SetValue("a2");
                    sheet.Cell(3, 2).SetValue("b2");
                    sheet.Cell(3, 3).SetValue("c2");
                    sheet.Cell(3, 4).SetValue("d2");

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = _factory.CreateReader(stream)) {
                    excel.Configuration.RegisterClassMap<AMap>();
                    var records = excel.GetRecords<A>().ToList();

                    // Make sure we got our records
                    Assert.AreEqual(2, records.Count);
                    for (var i = 0; i < records.Count; i++) {
                        var rowId = i + 1;
                        var row = records[i];
                        Assert.AreEqual("a" + rowId, row.Id);
                        Assert.AreEqual("b" + rowId, row.B.Id);
                        Assert.AreEqual("c" + rowId, row.B.C.Id);
                        Assert.AreEqual("d" + rowId, row.B.C.D.Id);
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
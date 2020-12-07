﻿/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using C1.C1Excel;
using ExcelHelper.Configuration;
using ExcelHelper.TypeConversion;
using NUnit.Framework;
// ReSharper disable UnusedAutoPropertyAccessor.Local
// ReSharper disable ClassNeverInstantiated.Local

namespace ExcelHelper.Tests
{
    [TestFixture]
    public class ExcelWriterTests
    {
        [Test]
        public void WriteCellTest()
        {
            using (var stream = new MemoryStream()) {
                using (var excel = new ExcelWriter(stream)) {
                    // Set up our row and column formats first
                    excel.SetRowFormat(1, fontStyle: FontStyle.Bold, fontSize: 16);
                    excel.SetColumnFormat(7, fontStyle: FontStyle.Italic, fontSize: 24);

                    var date = DateTime.Today;
                    var guid = new Guid("bfb9c599-bc9e-4f97-ae59-25f2ca09cfdf");
                    excel.WriteCell(0, 0, "one");
                    excel.WriteCell(0, 1, "one, two", fontStyle: FontStyle.Bold);
                    excel.WriteCell(0, 2, "one \"two\" three", fontSize: 18);
                    excel.WriteCell(0, 3, " one ", fontName: "Times");
                    excel.WriteCell(0, 4, date);
                    excel.WriteCell(0, 5, date, null, "d");
                    excel.WriteCell(0, 6, date, null, "D", FontStyle.Bold, horizontalAlign: ExcelAlignHorizontal.Right, verticalAlign: ExcelAlignVertical.Center);
                    excel.WriteCell(0, 7, (byte)1);
                    excel.WriteCell(0, 8, (short)2);
                    excel.WriteCell(0, 9, 3);
                    excel.WriteCell(0, 10, "=1+2");

                    excel.WriteCell(1, 0, (long)4);
                    excel.WriteCell(1, 1, (float)5);
                    excel.WriteCell(1, 2, (double)6);
                    excel.WriteCell(1, 3, (decimal)123.456, "C");
                    excel.WriteCell(1, 4, guid);
                    excel.WriteCell(1, 5, true);
                    excel.WriteCell(1, 6, false);
                    excel.WriteCell(1, 7, (string)null);
                    excel.WriteCell(1, 8, new TimeSpan(1, 2, 3));
                    excel.WriteCell(1, 9, new TimeSpan(1, 2, 3), "g");
                    excel.WriteCell(1, 10, "=2*3");

                    // Auto size the columns
                    excel.AdjustColumnsToContent(0, 10000);

                    // Override some column and row sizes
                    excel.SetRowHeight(1, 600);
                    excel.SetColumnWidth(1, 500);
                    excel.SetColumnWidth(11, 700);

                    // Change to third sheet and write a cell
                    excel.ChangeSheet(2);
                    excel.WriteCell(0, 0, "third sheet");

                    excel.Close();

                    stream.Position = 0;
                    using (var book = new C1XLBook()) {
                        book.Load(stream, FileFormat.OpenXml);
                        var sheet = book.Sheets[0];

                        // Verify row and column styles
                        Assert.AreEqual(FontStyle.Bold, sheet.Rows[1].Style.Font.Style);
                        Assert.AreEqual(16.0, sheet.Rows[1].Style.Font.SizeInPoints);
                        Assert.AreEqual(FontStyle.Italic, sheet.Columns[7].Style.Font.Style);
                        Assert.AreEqual(24.0, sheet.Columns[7].Style.Font.SizeInPoints);

                        // Check some automatically sizes column widths
                        Assert.AreEqual(2655, sheet.Columns[2].Width);
                        Assert.AreEqual(5541, sheet.Columns[4].Width);

                        // Verify the overridden row and column sizes
                        Assert.AreEqual(600, sheet.Rows[1].Height);
                        Assert.AreEqual(495, sheet.Columns[1].Width);
                        Assert.AreEqual(700, sheet.Columns[11].Width);

                        // Verify first row
                        Assert.AreEqual("one", sheet[0, 0].Value);
                        Assert.AreEqual("", sheet[0, 0].Style.Format);
                        Assert.AreEqual("one, two", sheet[0, 1].Value);
                        Assert.AreEqual(FontStyle.Bold, sheet[0, 1].Style.Font.Style);
                        Assert.AreEqual("one \"two\" three", sheet[0, 2].Value);
                        Assert.AreEqual(18.0, sheet[0, 2].Style.Font.SizeInPoints);
                        Assert.AreEqual(" one ", sheet[0, 3].Value);
                        Assert.AreEqual("Times New Roman", sheet[0, 3].Style.Font.Name);
                        Assert.AreEqual(date, sheet[0, 4].Value);
                        Assert.AreEqual(@"m\/D\/YYYY\ H:mm:ss\ AM/PM", sheet[0, 4].Style.Format);
                        Assert.AreEqual(date, sheet[0, 5].Value);
                        Assert.AreEqual(@"m\/D\/YYYY", sheet[0, 5].Style.Format);
                        Assert.AreEqual(date, sheet[0, 6].Value);
                        Assert.AreEqual(@"DDDD,\ mmmm\ D,\ YYYY", sheet[0, 6].Style.Format);
                        Assert.AreEqual(FontStyle.Bold, sheet[0, 6].Style.Font.Style);
                        Assert.AreEqual(XLAlignHorzEnum.Right, sheet[0, 6].Style.AlignHorz);
                        Assert.AreEqual(XLAlignVertEnum.Center, sheet[0, 6].Style.AlignVert);
                        Assert.AreEqual((double)1, sheet[0, 7].Value);
                        Assert.AreEqual((double)2, sheet[0, 8].Value);
                        Assert.AreEqual((double)3, sheet[0, 9].Value);
                        Assert.AreEqual("1+2", sheet[0, 10].Formula);
                        Assert.AreEqual(null, sheet[0, 10].Value);

                        // Verify second row
                        Assert.AreEqual((double)4, sheet[1, 0].Value);
                        Assert.AreEqual("", sheet[1, 0].Style.Format);
                        Assert.AreEqual((double)5, sheet[1, 1].Value);
                        Assert.AreEqual((double)6, sheet[1, 2].Value);
                        Assert.AreEqual(123.456, sheet[1, 3].Value);
                        Assert.AreEqual(@"$#,##0.00;($#,##0.00)", sheet[1, 3].Style.Format);
                        Assert.AreEqual("bfb9c599-bc9e-4f97-ae59-25f2ca09cfdf", sheet[1, 4].Value);
                        Assert.AreEqual("true", sheet[1, 5].Value);
                        Assert.AreEqual("false", sheet[1, 6].Value);
                        Assert.AreEqual(null, sheet[1, 7].Value);
                        Assert.AreEqual("01:02:03", sheet[1, 8].Value);
                        Assert.AreEqual("", sheet[1, 8].Style.Format);
                        Assert.AreEqual("01:02:03", sheet[1, 9].Value);
                        Assert.AreEqual("", sheet[1, 9].Style.Format);
                        Assert.AreEqual("2*3", sheet[1, 10].Formula);
                        Assert.AreEqual(null, sheet[1, 10].Value);

                        // Verify third sheet
                        Assert.AreEqual(3, book.Sheets.Count);
                        sheet = book.Sheets[2];
                        Assert.AreEqual("third sheet", sheet[0, 0].Value);
                    }
                }
            }
        }

        [Test]
        public void LargeFileTest()
        {
            using (var stream = new MemoryStream()) {
                using (var excel = new ExcelWriter(stream)) {
                    // Write out 66K rows
                    for (var i = 0; i < 66000; i++) {
                        excel.WriteCell(i, 0, i.ToString());
                    }
                    excel.Close();

                    // Now read it back
                    stream.Position = 0;
                    using (var book = new C1XLBook()) {
                        book.Load(stream, FileFormat.OpenXml);
                        var sheet = book.Sheets[0];

                        // Verify 66K rows
                        for (var i = 0; i < 66000; i++) {
                            Assert.AreEqual(i.ToString(), sheet[i, 0].Value);
                        }
                    }
                }
            }
        }

        [Test]
        public void WriteRecordsTest()
        {
            var date = DateTime.Today;
            var yesterday = DateTime.Today.AddDays(-1);
            var records = new List<TestRecord> {
                new TestRecord {
                    IntColumn = 1,
                    StringColumn = "string column",
                    IgnoredColumn = "ignored column",
                    FirstColumn = "first column",
                    TypeConvertedColumn = "written as test",
                    BoolColumn = true,
                    DoubleColumn = 12.34,
                    DateTimeColumn = date,
                    NullStringColumn = null,
                    FormulaColumn = "=1+2",
                },
                new TestRecord {
                    IntColumn = 2,
                    StringColumn = "string column 2",
                    IgnoredColumn = "ignored column 2",
                    FirstColumn = "first column 2",
                    TypeConvertedColumn = "written as test",
                    BoolColumn = false,
                    DoubleColumn = 43.21,
                    DateTimeColumn = yesterday,
                    NullStringColumn = null,
                    FormulaColumn = "not a formula",
                },
            };

            using (var stream = new MemoryStream()) {
                using (var excel = new ExcelWriter(stream)) {
                    excel.Configuration.RegisterClassMap<TestRecordMap>();
                    excel.WriteRecords(records);
                    excel.ChangeSheet(2);
                    excel.WriteRecords(records, false);
                    excel.Close();

                    stream.Position = 0;
                    using (var book = new C1XLBook()) {
                        book.Load(stream, FileFormat.OpenXml);
                        var sheet = book.Sheets[0];
                        CheckRecords(sheet, date, yesterday);
                        Assert.AreEqual(3, book.Sheets.Count);
                        sheet = book.Sheets[2];
                        CheckRecords(sheet, date, yesterday, false);
                    }
                }
            }
        }

        /// <summary>
        /// Checks the records in the sheet
        /// </summary>
        /// <param name="sheet">Sheet to check</param>
        /// <param name="date">Current date</param>
        /// <param name="yesterday">Yesterdays date</param>
        /// <param name="checkHeader">True to check the header, false to not</param>
        private static void CheckRecords(
            XLSheet sheet,
            DateTime date,
            DateTime yesterday,
            bool checkHeader = true)
        {
            var row = 0;
            if (checkHeader) {
                // Check the header is bold
                Assert.AreEqual(FontStyle.Bold, sheet.Rows[0].Style.Font.Style);

                // Check the header row
                Assert.AreEqual("FirstColumn", sheet[row, 0].Value);
                Assert.AreEqual("Int Column", sheet[row, 1].Value);
                Assert.AreEqual("StringColumn", sheet[row, 2].Value);
                Assert.AreEqual("TypeConvertedColumn", sheet[row, 3].Value);
                Assert.AreEqual("BoolColumn", sheet[row, 4].Value);
                Assert.AreEqual("DoubleColumn", sheet[row, 5].Value);
                Assert.AreEqual("DateTimeColumn", sheet[row, 6].Value);
                Assert.AreEqual("NullStringColumn", sheet[row, 7].Value);
                Assert.AreEqual("FormulaColumn", sheet[row, 8].Value);
                row++;
            }

            // Check the first record
            Assert.AreEqual("first column", sheet[row, 0].Value);
            Assert.AreEqual((double)1, sheet[row, 1].Value);
            Assert.AreEqual("string column", sheet[row, 2].Value);
            Assert.AreEqual("test", sheet[row, 3].Value);
            Assert.AreEqual("true", sheet[row, 4].Value);
            Assert.AreEqual(12.34, sheet[row, 5].Value);
            Assert.AreEqual(date, sheet[row, 6].Value);
            Assert.AreEqual(@"m\/D\/YYYY\ H:mm:ss\ AM/PM", sheet[row, 6].Style.Format);
            Assert.AreEqual(null, sheet[row, 7].Value);
            Assert.AreEqual("1+2", sheet[row, 8].Formula);
            Assert.AreEqual(null, sheet[row, 8].Value);
            row++;

            // Check the second record
            Assert.AreEqual("first column 2", sheet[row, 0].Value);
            Assert.AreEqual((double)2, sheet[row, 1].Value);
            Assert.AreEqual("string column 2", sheet[row, 2].Value);
            Assert.AreEqual("test", sheet[row, 3].Value);
            Assert.AreEqual("false", sheet[row, 4].Value);
            Assert.AreEqual(43.21, sheet[row, 5].Value);
            Assert.AreEqual(yesterday, sheet[row, 6].Value);
            Assert.AreEqual(@"m\/D\/YYYY\ H:mm:ss\ AM/PM", sheet[row, 6].Style.Format);
            Assert.AreEqual(null, sheet[row, 7].Value);
            Assert.AreEqual("not a formula", sheet[row, 8].Value);
        }

        [Test]
        public void WriteRecordsNoIndexesTest()
        {
            var records = new List<TestRecordNoIndexes> {
                new TestRecordNoIndexes {
                    IntColumn = 1,
                    StringColumn = "string column",
                    IgnoredColumn = "ignored column",
                    FirstColumn = "first column",
                    TypeConvertedColumn = "written as test",
                },
            };

            using (var stream = new MemoryStream()) {
                using (var excel = new ExcelWriter(stream)) {
                    excel.Configuration.RegisterClassMap<TestRecordNoIndexesMap>();
                    excel.WriteRecords(records);
                    excel.Close();

                    stream.Position = 0;
                    using (var book = new C1XLBook()) {
                        book.Load(stream, FileFormat.OpenXml);
                        var sheet = book.Sheets[0];

                        // Check the header row
                        Assert.AreEqual("Int Column", sheet[0, 0].Value);
                        Assert.AreEqual("StringColumn", sheet[0, 1].Value);
                        Assert.AreEqual("FirstColumn", sheet[0, 2].Value);
                        Assert.AreEqual("TypeConvertedColumn", sheet[0, 3].Value);

                        // Check the first record
                        Assert.AreEqual((double)1, sheet[1, 0].Value);
                        Assert.AreEqual("string column", sheet[1, 1].Value);
                        Assert.AreEqual("first column", sheet[1, 2].Value);
                        Assert.AreEqual("test", sheet[1, 3].Value);
                    }
                }
            }
        }

        [Test]
        public void WriteRecordsWithReferencesTest()
        {
            var records = new List<Person> {
                new Person {
                    FirstName = "First Name",
                    LastName = "Last Name",
                    HomeAddress = new Address {
                        Street = "Home Street",
                        City = "Home City",
                        State = "Home State",
                        Zip = "Home Zip",
                        ID = 2,
                    },
                    WorkAddress = new Address {
                        Street = "Work Street",
                        City = "Work City",
                        State = "Work State",
                        Zip = "Work Zip",
                        ID = 3,
                    },
                    NullAddress = null,
                },
            };

            using (var stream = new MemoryStream()) {
                using (var excel = new ExcelWriter(stream)) {
                    excel.Configuration.RegisterClassMap<PersonMap>();
                    excel.WriteRecords(records);
                    excel.Close();

                    stream.Position = 0;
                    using (var book = new C1XLBook()) {
                        book.Load(stream, FileFormat.OpenXml);
                        var sheet = book.Sheets[0];

                        // Check the header row
                        Assert.AreEqual("FirstName", sheet[0, 0].Value);
                        Assert.AreEqual("LastName", sheet[0, 1].Value);
                        Assert.AreEqual("HomeStreet", sheet[0, 2].Value);
                        Assert.AreEqual("HomeCity", sheet[0, 3].Value);
                        Assert.AreEqual("HomeState", sheet[0, 4].Value);
                        Assert.AreEqual("HomeZip", sheet[0, 5].Value);
                        Assert.AreEqual("HomeID", sheet[0, 6].Value);
                        Assert.AreEqual("WorkStreet", sheet[0, 7].Value);
                        Assert.AreEqual("WorkCity", sheet[0, 8].Value);
                        Assert.AreEqual("WorkState", sheet[0, 9].Value);
                        Assert.AreEqual("WorkZip", sheet[0, 10].Value);
                        Assert.AreEqual("WorkID", sheet[0, 11].Value);
                        Assert.AreEqual("NullStreet", sheet[0, 12].Value);
                        Assert.AreEqual("NullCity", sheet[0, 13].Value);
                        Assert.AreEqual("NullState", sheet[0, 14].Value);
                        Assert.AreEqual("NullZip", sheet[0, 15].Value);
                        Assert.AreEqual("NullID", sheet[0, 16].Value);

                        // Check the record
                        Assert.AreEqual("First Name", sheet[1, 0].Value);
                        Assert.AreEqual("Last Name", sheet[1, 1].Value);
                        Assert.AreEqual("Home Street", sheet[1, 2].Value);
                        Assert.AreEqual("Home City", sheet[1, 3].Value);
                        Assert.AreEqual("Home State", sheet[1, 4].Value);
                        Assert.AreEqual("Home Zip", sheet[1, 5].Value);
                        Assert.AreEqual(2.0, sheet[1, 6].Value);
                        Assert.AreEqual("Work Street", sheet[1, 7].Value);
                        Assert.AreEqual("Work City", sheet[1, 8].Value);
                        Assert.AreEqual("Work State", sheet[1, 9].Value);
                        Assert.AreEqual("Work Zip", sheet[1, 10].Value);
                        Assert.AreEqual(3.0, sheet[1, 11].Value);
                        Assert.AreEqual(null, sheet[1, 12].Value);
                        Assert.AreEqual(null, sheet[1, 13].Value);
                        Assert.AreEqual(null, sheet[1, 14].Value);
                        Assert.AreEqual(null, sheet[1, 15].Value);
                        Assert.AreEqual(0.0, sheet[1, 16].Value);
                    }
                }
            }
        }

        [Test]
        public void WriteNoGetterTest()
        {
            using (var stream = new MemoryStream()) {
                using (var excel = new ExcelWriter(stream)) {
                    var list = new List<TestPrivateGet> {
                        new TestPrivateGet {
                            ID = 1,
                            Name = "one"
                        }
                    };
                    excel.WriteRecords(list);
                    excel.Close();

                    stream.Position = 0;
                    using (var book = new C1XLBook()) {
                        book.Load(stream, FileFormat.OpenXml);
                        var sheet = book.Sheets[0];
                        Assert.AreEqual("ID", sheet[0, 0].Value);
                        Assert.AreEqual(null, sheet[0, 1].Value);
                        Assert.AreEqual((double)1, sheet[1, 0].Value);
                        Assert.AreEqual(null, sheet[1, 1].Value);
                    }
                }
            }
        }

        [Test]
        public void WriteMultipleNamesTest()
        {
            var records = new List<MultipleNamesClass> {
                new MultipleNamesClass {
                    IntColumn = 1,
                    StringColumn = "one"
                },
                new MultipleNamesClass {
                    IntColumn = 2,
                    StringColumn = "two"
                }
            };

            using (var stream = new MemoryStream()) {
                using (var excel = new ExcelWriter(stream)) {
                    excel.Configuration.RegisterClassMap<MultipleNamesClassMap>();
                    excel.WriteRecords(records);
                    excel.Close();

                    stream.Position = 0;
                    using (var book = new C1XLBook()) {
                        book.Load(stream, FileFormat.OpenXml);
                        var sheet = book.Sheets[0];

                        // Check the header row
                        Assert.AreEqual("int1", sheet[0, 0].Value);
                        Assert.AreEqual("string1", sheet[0, 1].Value);

                        // Check the first record
                        Assert.AreEqual((double)1, sheet[1, 0].Value);
                        Assert.AreEqual("one", sheet[1, 1].Value);

                        // Check the second record
                        Assert.AreEqual((double)2, sheet[2, 0].Value);
                        Assert.AreEqual("two", sheet[2, 1].Value);
                    }
                }
            }
        }

        [Test]
        public void SameNameMultipleTimesTest()
        {
            var records = new List<SameNameMultipleTimesClass> {
                new SameNameMultipleTimesClass {
                    Name1 = "1",
                    Name2 = "2",
                    Name3 = "3"
                }
            };

            using (var stream = new MemoryStream()) {
                using (var excel = new ExcelWriter(stream)) {
                    excel.Configuration.RegisterClassMap<SameNameMultipleTimesClassMap>();
                    excel.WriteRecords(records);
                    excel.Close();

                    stream.Position = 0;
                    using (var book = new C1XLBook()) {
                        book.Load(stream, FileFormat.OpenXml);
                        var sheet = book.Sheets[0];

                        // Check the header row
                        Assert.AreEqual("ColumnName", sheet[0, 0].Value);
                        Assert.AreEqual("ColumnName", sheet[0, 1].Value);
                        Assert.AreEqual("ColumnName", sheet[0, 2].Value);

                        // Check the first record
                        Assert.AreEqual("1", sheet[1, 0].Value);
                        Assert.AreEqual("2", sheet[1, 1].Value);
                        Assert.AreEqual("3", sheet[1, 2].Value);
                    }
                }
            }
        }

        private class TestPrivateGet
        {
            public int ID { get; set; }
            public string Name { private get; set; }
        }

        private class TestRecord
        {
            public int IntColumn { get; set; }
            public string StringColumn { get; set; }
            public string IgnoredColumn { get; set; }
            public string FirstColumn { get; set; }
            public string TypeConvertedColumn { get; set; }
            public bool BoolColumn { get; set; }
            public double DoubleColumn { get; set; }
            public DateTime DateTimeColumn { get; set; }
            public string NullStringColumn { get; set; }
            public string FormulaColumn { get; set; }
        }

        private sealed class TestRecordMap : ExcelClassMap<TestRecord>
        {
            public TestRecordMap()
            {
                Map(m => m.IntColumn).Name("Int Column").Index(1).TypeConverter<Int32Converter>();
                Map(m => m.StringColumn);
                Map(m => m.FirstColumn).Index(0);
                Map(m => m.TypeConvertedColumn).TypeConverter<TestTypeConverter>();
                Map(m => m.BoolColumn);
                Map(m => m.DoubleColumn);
                Map(m => m.DateTimeColumn);
                Map(m => m.NullStringColumn);
                Map(m => m.FormulaColumn).IsFormula();
            }
        }

        private class TestRecordNoIndexes
        {
            public int IntColumn { get; set; }
            public string StringColumn { get; set; }
            public string IgnoredColumn { get; set; }
            public string FirstColumn { get; set; }
            public string TypeConvertedColumn { get; set; }
        }

        private sealed class TestRecordNoIndexesMap : ExcelClassMap<TestRecordNoIndexes>
        {
            public TestRecordNoIndexesMap()
            {
                Map(m => m.IntColumn).Name("Int Column").TypeConverter<Int32Converter>();
                Map(m => m.StringColumn);
                Map(m => m.FirstColumn);
                Map(m => m.TypeConvertedColumn).TypeConverter<TestTypeConverter>();
            }
        }

        private class TestTypeConverter : ITypeConverter
        {
            public bool AcceptsNativeType => false;

            public Type ConvertedType => typeof(object);

            public object ConvertToExcel(
                TypeConverterOptions options,
                object value)
            {
                return "test";
            }

            public object ConvertFromExcel(
                TypeConverterOptions options,
                object excelValue)
            {
                throw new NotImplementedException();
            }

            public string ExcelFormatString(
                TypeConverterOptions options)
            {
                return null;
            }
        }

        private class Person
        {
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public Address HomeAddress { get; set; }
            public Address WorkAddress { get; set; }
            public Address NullAddress { get; set; }
        }

        private class Address
        {
            public string Street { get; set; }
            public string City { get; set; }
            public string State { get; set; }
            public string Zip { get; set; }
            public int ID { get; set; }
        }

        private sealed class PersonMap : ExcelClassMap<Person>
        {
            public PersonMap()
            {
                Map(m => m.FirstName);
                Map(m => m.LastName);
                References<HomeAddressMap>(m => m.HomeAddress);
                References<WorkAddressMap>(m => m.WorkAddress);
                References<NullAddressMap>(m => m.NullAddress);
            }
        }

        private sealed class HomeAddressMap : ExcelClassMap<Address>
        {
            public HomeAddressMap()
            {
                Map(m => m.Street).Name("HomeStreet");
                Map(m => m.City).Name("HomeCity");
                Map(m => m.State).Name("HomeState");
                Map(m => m.Zip).Name("HomeZip");
                Map(m => m.ID).Name("HomeID");
            }
        }

        private sealed class WorkAddressMap : ExcelClassMap<Address>
        {
            public WorkAddressMap()
            {
                Map(m => m.Street).Name("WorkStreet");
                Map(m => m.City).Name("WorkCity");
                Map(m => m.State).Name("WorkState");
                Map(m => m.Zip).Name("WorkZip");
                Map(m => m.ID).Name("WorkID");
            }
        }

        private sealed class NullAddressMap : ExcelClassMap<Address>
        {
            public NullAddressMap()
            {
                Map(m => m.Street).Name("NullStreet");
                Map(m => m.City).Name("NullCity");
                Map(m => m.State).Name("NullState");
                Map(m => m.Zip).Name("NullZip");
                Map(m => m.ID).Name("NullID");
            }
        }

        private class SameNameMultipleTimesClass
        {
            public string Name1 { get; set; }
            public string Name2 { get; set; }
            public string Name3 { get; set; }
        }

        private sealed class SameNameMultipleTimesClassMap : ExcelClassMap<SameNameMultipleTimesClass>
        {
            public SameNameMultipleTimesClassMap()
            {
                Map(m => m.Name1).Name("ColumnName").NameIndex(1);
                Map(m => m.Name2).Name("ColumnName").NameIndex(2);
                Map(m => m.Name3).Name("ColumnName").NameIndex(0);
            }
        }

        private class MultipleNamesClass
        {
            public int IntColumn { get; set; }
            public string StringColumn { get; set; }
        }

        private sealed class MultipleNamesClassMap : ExcelClassMap<MultipleNamesClass>
        {
            public MultipleNamesClassMap()
            {
                Map(m => m.IntColumn).Name("int1", "int2", "int3");
                Map(m => m.StringColumn).Name("string1", "string2", "string3");
            }
        }
    }
}
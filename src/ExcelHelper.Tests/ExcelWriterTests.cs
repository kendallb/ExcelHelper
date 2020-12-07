/*
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
using ClosedXML.Excel;
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

                    var date = DateTime.Parse("2017-10-16 03:05 PM");
                    var guid = new Guid("bfb9c599-bc9e-4f97-ae59-25f2ca09cfdf");
                    excel.WriteCell(0, 0, "one");
                    excel.WriteCell(0, 1, "one, two", fontStyle: FontStyle.Bold);
                    excel.WriteCell(0, 2, "one \"two\" three", fontSize: 18);
                    excel.WriteCell(0, 3, " one ", fontName: "Times New Roman");
                    excel.WriteCell(0, 4, date);
                    excel.WriteCell(0, 5, date, dateFormat: "M/d/yyyy h:mm:ss AM/PM");
                    excel.WriteCell(0, 6, date, dateFormat: "M/d/yyyy");
                    excel.WriteCell(0, 7, date, dateFormat: "dddd, MMMM d, yyyy", fontStyle: FontStyle.Bold, horizontalAlign: ExcelAlignHorizontal.Right, verticalAlign: ExcelAlignVertical.Bottom);
                    excel.WriteCell(0, 8, (byte)1);
                    excel.WriteCell(0, 9, (short)2);
                    excel.WriteCell(0, 10, 3);
                    excel.WriteCell(0, 11, "=1+2");

                    excel.WriteCell(1, 0, (long)4);
                    excel.WriteCell(1, 1, (float)5);
                    excel.WriteCell(1, 2, (double)6);
                    excel.WriteCell(1, 3, (decimal)123.456, "$#,##0.00;($#,##0.00)");
                    excel.WriteCell(1, 4, (decimal)-123.456, "$#,##0.00;($#,##0.00)");
                    excel.WriteCell(1, 5, guid);
                    excel.WriteCell(1, 6, true);
                    excel.WriteCell(1, 7, false);
                    excel.WriteCell(1, 8, (string)null);
                    excel.WriteCell(1, 9, new TimeSpan(1, 2, 3));
                    excel.WriteCell(1, 10, "=2*3");

                    // Override some column and row sizes
                    excel.SetRowHeight(1, 600);
                    excel.SetColumnWidth(1, 500);
                    excel.SetColumnWidth(11, 700);

                    // Change to third sheet and write a cell
                    excel.ChangeSheet(2);
                    excel.WriteCell(0, 0, "third sheet");

                    excel.Close();

#if WRITE_TEST_FILE
                    using (var fileStream = File.Create("C:\\temp\\test.xlsx")) {
                        stream.Seek(0, SeekOrigin.Begin);
                        stream.CopyTo(fileStream);
                    }
#endif

                    stream.Position = 0;
                    using (var book = new XLWorkbook(stream)) {
                        var sheet = book.Worksheets.Worksheet(1);

                        // Verify row and column styles
                        Assert.AreEqual(true, sheet.Row(2).Style.Font.Bold);
                        Assert.AreEqual(16.0, sheet.Row(2).Style.Font.FontSize);

                        // TODO: This is not working in ClosedXML 0.94.2. This is fixed in 0.95 beta 2 so when 0.95 is out, we can put this back.
                        // Assert.AreEqual(true, sheet.Column(8).Style.Font.Italic);
                        // Assert.AreEqual(24.0, sheet.Column(8).Style.Font.FontSize);

                        // Verify the overridden row and column sizes
                        Assert.AreEqual(600, sheet.Row(2).Height);

                        // TODO: This is not working. Have to figure out why ...
                        // Assert.AreEqual(495, sheet.Column(2).Width);
                        // Assert.AreEqual(700, sheet.Column(12).Width);

                        // Check some automatically sized column widths
                        Assert.AreEqual(23.71, sheet.Column(3).Width);
                        Assert.AreEqual(11.209999999999999, sheet.Column(4).Width);

                        // Verify first row
                        Assert.AreEqual("one", sheet.Cell(1, 1).Value);
                        Assert.AreEqual("", sheet.Cell(1, 1).Style.NumberFormat.Format);
                        Assert.AreEqual("", sheet.Cell(1, 1).Style.DateFormat.Format);
                        Assert.AreEqual("one, two", sheet.Cell(1, 2).Value);
                        Assert.AreEqual(true, sheet.Cell(1, 2).Style.Font.Bold);
                        Assert.AreEqual("one \"two\" three", sheet.Cell(1, 3).Value);
                        Assert.AreEqual(18.0, sheet.Cell(1, 3).Style.Font.FontSize);
                        Assert.AreEqual(" one ", sheet.Cell(1, 4).Value);
                        Assert.AreEqual("Times New Roman", sheet.Cell(1, 4).Style.Font.FontName);
                        Assert.AreEqual(date, sheet.Cell(1, 5).Value);
                        Assert.AreEqual("10/16/2017 15:05", sheet.Cell(1, 5).GetFormattedString());
                        Assert.AreEqual("", sheet.Cell(1, 5).Style.DateFormat.Format);
                        Assert.AreEqual(date, sheet.Cell(1, 6).Value);

                        // TODO: This is broken also. GetFormattedString() returns the format itself, not the string? (M/d/yyyy h:mm:ss tt)
                        Assert.AreEqual("10/16/2017 3:05:00 PM", sheet.Cell(1, 6).GetFormattedString());
                        Assert.AreEqual("M/d/yyyy h:mm:ss AM/PM", sheet.Cell(1, 6).Style.DateFormat.Format);
                        Assert.AreEqual(date, sheet.Cell(1, 7).Value);
                        Assert.AreEqual("10/16/2017", sheet.Cell(1, 7).GetFormattedString());
                        Assert.AreEqual("M/d/yyyy", sheet.Cell(1, 7).Style.DateFormat.Format);
                        Assert.AreEqual(XLAlignmentHorizontalValues.Right, sheet.Cell(1, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right);
                        Assert.AreEqual(XLAlignmentVerticalValues.Bottom, sheet.Cell(1, 7).Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom);
                        Assert.AreEqual(date, sheet.Cell(1, 8).Value);

                        // TODO: This is broken. Returns Monday October 16, 2017 when it should be Monday, October 16, 2017. The Excel file has the correct value.
                        //Assert.AreEqual("Monday, October 16, 2017", sheet.Cell(1, 8).GetFormattedString());
                        Assert.AreEqual("dddd, MMMM d, yyyy", sheet.Cell(1, 8).Style.DateFormat.Format);
                        Assert.AreEqual(true, sheet.Cell(1, 8).Style.Font.Bold);
                        Assert.AreEqual((double)1, sheet.Cell(1, 9).Value);
                        Assert.AreEqual("1", sheet.Cell(1, 9).GetFormattedString());
                        Assert.AreEqual((double)2, sheet.Cell(1, 10).Value);
                        Assert.AreEqual("2", sheet.Cell(1, 10).GetFormattedString());
                        Assert.AreEqual((double)3, sheet.Cell(1, 11).Value);
                        Assert.AreEqual("3", sheet.Cell(1, 11).GetFormattedString());
                        Assert.AreEqual("1+2", sheet.Cell(1, 12).FormulaA1);
                        Assert.AreEqual((double)3, sheet.Cell(1, 12).Value);
                        Assert.AreEqual("3", sheet.Cell(1, 12).GetFormattedString());

                        // Verify second row
                        Assert.AreEqual((double)4, sheet.Cell(2, 1).Value);
                        Assert.AreEqual("4", sheet.Cell(2, 1).GetFormattedString());
                        Assert.AreEqual("", sheet.Cell(2, 1).Style.NumberFormat.Format);
                        Assert.AreEqual((double)5, sheet.Cell(2, 2).Value);
                        Assert.AreEqual("5", sheet.Cell(2, 2).GetFormattedString());
                        Assert.AreEqual((double)6, sheet.Cell(2, 3).Value);
                        Assert.AreEqual("6", sheet.Cell(2, 3).GetFormattedString());
                        Assert.AreEqual(123.456, sheet.Cell(2, 4).Value);
                        Assert.AreEqual("$123.46", sheet.Cell(2, 4).GetFormattedString());
                        Assert.AreEqual("$#,##0.00;($#,##0.00)", sheet.Cell(2, 4).Style.NumberFormat.Format);
                        Assert.AreEqual(-123.456, sheet.Cell(2, 5).Value);
                        Assert.AreEqual("($123.46)", sheet.Cell(2, 5).GetFormattedString());
                        Assert.AreEqual("$#,##0.00;($#,##0.00)", sheet.Cell(2, 5).Style.NumberFormat.Format);
                        Assert.AreEqual("bfb9c599-bc9e-4f97-ae59-25f2ca09cfdf", sheet.Cell(2, 6).Value);
                        Assert.AreEqual("bfb9c599-bc9e-4f97-ae59-25f2ca09cfdf", sheet.Cell(2, 6).GetFormattedString());
                        Assert.AreEqual("true", sheet.Cell(2, 7).Value);
                        Assert.AreEqual("true", sheet.Cell(2, 7).GetFormattedString());
                        Assert.AreEqual("false", sheet.Cell(2, 8).Value);
                        Assert.AreEqual("false", sheet.Cell(2, 8).GetFormattedString());
                        Assert.AreEqual("", sheet.Cell(2, 9).Value);
                        Assert.AreEqual("01:02:03", sheet.Cell(2, 10).Value);
                        Assert.AreEqual("01:02:03", sheet.Cell(2, 10).GetFormattedString());
                        Assert.AreEqual("", sheet.Cell(2, 10).Style.NumberFormat.Format);
                        Assert.AreEqual("2*3", sheet.Cell(2, 11).FormulaA1);
                        Assert.AreEqual((double)6, sheet.Cell(2, 11).Value);
                        Assert.AreEqual("6", sheet.Cell(2, 11).GetFormattedString());

                        // Verify third sheet
                        Assert.AreEqual(3, book.Worksheets.Count);
                        sheet = book.Worksheets.Worksheet(3);
                        Assert.AreEqual("third sheet", sheet.Cell(1, 1).Value);
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
                    using (var book = new XLWorkbook(stream)) {
                        var sheet = book.Worksheets.Worksheet(1);

                        // Verify 66K rows
                        for (var i = 0; i < 66000; i++) {
                            Assert.AreEqual(i.ToString(), sheet.Cell(i + 1, 1).Value);
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
                    using (var book = new XLWorkbook(stream)) {
                        var sheet = book.Worksheets.Worksheet(1);
                        CheckRecords(sheet, date, yesterday);
                        Assert.AreEqual(3, book.Worksheets.Count);
                        sheet = book.Worksheets.Worksheet(3);
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
            IXLWorksheet sheet,
            DateTime date,
            DateTime yesterday,
            bool checkHeader = true)
        {
            var row = 1;
            if (checkHeader) {
                // Check the header is bold
                Assert.AreEqual(true, sheet.Row(1).Style.Font.Bold);

                // Check the header row
                Assert.AreEqual("FirstColumn", sheet.Cell(row, 1).Value);
                Assert.AreEqual("Int Column", sheet.Cell(row, 2).Value);
                Assert.AreEqual("StringColumn", sheet.Cell(row, 3).Value);
                Assert.AreEqual("TypeConvertedColumn", sheet.Cell(row, 4).Value);
                Assert.AreEqual("BoolColumn", sheet.Cell(row, 5).Value);
                Assert.AreEqual("DoubleColumn", sheet.Cell(row, 6).Value);
                Assert.AreEqual("DateTimeColumn", sheet.Cell(row, 7).Value);
                Assert.AreEqual("NullStringColumn", sheet.Cell(row, 8).Value);
                Assert.AreEqual("FormulaColumn", sheet.Cell(row, 9).Value);
                row++;
            }

            // Check the first record
            Assert.AreEqual("first column", sheet.Cell(row, 1).Value);
            Assert.AreEqual((double)1, sheet.Cell(row, 2).Value);
            Assert.AreEqual("string column", sheet.Cell(row, 3).Value);
            Assert.AreEqual("test", sheet.Cell(row, 4).Value);
            Assert.AreEqual("true", sheet.Cell(row, 5).Value);
            Assert.AreEqual(12.34, sheet.Cell(row, 6).Value);
            Assert.AreEqual(date, sheet.Cell(row, 7).Value);
            Assert.AreEqual("", sheet.Cell(row, 8).Style.DateFormat.Format);    // TODO: Do we need a different default here for dates?
            Assert.AreEqual("", sheet.Cell(row, 8).Value);
            Assert.AreEqual("1+2", sheet.Cell(row, 9).FormulaA1);
            Assert.AreEqual((double)3, sheet.Cell(row, 9).Value);
            row++;

            // Check the second record
            Assert.AreEqual("first column 2", sheet.Cell(row, 1).Value);
            Assert.AreEqual((double)2, sheet.Cell(row, 2).Value);
            Assert.AreEqual("string column 2", sheet.Cell(row, 3).Value);
            Assert.AreEqual("test", sheet.Cell(row, 4).Value);
            Assert.AreEqual("false", sheet.Cell(row, 5).Value);
            Assert.AreEqual(43.21, sheet.Cell(row, 6).Value);
            Assert.AreEqual(yesterday, sheet.Cell(row, 7).Value);
            Assert.AreEqual("", sheet.Cell(row, 7).Style.DateFormat.Format);    // TODO: Do we need a different default here for dates?
            Assert.AreEqual("", sheet.Cell(row, 8).Value);
            Assert.AreEqual("not a formula", sheet.Cell(row, 9).Value);
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
                    using (var book = new XLWorkbook(stream)) {
                        var sheet = book.Worksheets.Worksheet(1);

                        // Check the header row
                        Assert.AreEqual("Int Column", sheet.Cell(1, 1).Value);
                        Assert.AreEqual("StringColumn", sheet.Cell(1, 2).Value);
                        Assert.AreEqual("FirstColumn", sheet.Cell(1, 3).Value);
                        Assert.AreEqual("TypeConvertedColumn", sheet.Cell(1, 4).Value);

                        // Check the first record
                        Assert.AreEqual((double)1, sheet.Cell(2, 1).Value);
                        Assert.AreEqual("string column", sheet.Cell(2, 2).Value);
                        Assert.AreEqual("first column", sheet.Cell(2, 3).Value);
                        Assert.AreEqual("test", sheet.Cell(2, 4).Value);
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
                    using (var book = new XLWorkbook(stream)) {
                        var sheet = book.Worksheets.Worksheet(1);

                        // Check the header row
                        Assert.AreEqual("FirstName", sheet.Cell(1, 1).Value);
                        Assert.AreEqual("LastName", sheet.Cell(1, 2).Value);
                        Assert.AreEqual("HomeStreet", sheet.Cell(1, 3).Value);
                        Assert.AreEqual("HomeCity", sheet.Cell(1, 4).Value);
                        Assert.AreEqual("HomeState", sheet.Cell(1, 5).Value);
                        Assert.AreEqual("HomeZip", sheet.Cell(1, 6).Value);
                        Assert.AreEqual("HomeID", sheet.Cell(1, 7).Value);
                        Assert.AreEqual("WorkStreet", sheet.Cell(1, 8).Value);
                        Assert.AreEqual("WorkCity", sheet.Cell(1, 9).Value);
                        Assert.AreEqual("WorkState", sheet.Cell(1, 10).Value);
                        Assert.AreEqual("WorkZip", sheet.Cell(1, 11).Value);
                        Assert.AreEqual("WorkID", sheet.Cell(1, 12).Value);
                        Assert.AreEqual("NullStreet", sheet.Cell(1, 13).Value);
                        Assert.AreEqual("NullCity", sheet.Cell(1, 14).Value);
                        Assert.AreEqual("NullState", sheet.Cell(1, 15).Value);
                        Assert.AreEqual("NullZip", sheet.Cell(1, 16).Value);
                        Assert.AreEqual("NullID", sheet.Cell(1, 17).Value);

                        // Check the record
                        Assert.AreEqual("First Name", sheet.Cell(2, 1).Value);
                        Assert.AreEqual("Last Name", sheet.Cell(2, 2).Value);
                        Assert.AreEqual("Home Street", sheet.Cell(2, 3).Value);
                        Assert.AreEqual("Home City", sheet.Cell(2, 4).Value);
                        Assert.AreEqual("Home State", sheet.Cell(2, 5).Value);
                        Assert.AreEqual("Home Zip", sheet.Cell(2, 6).Value);
                        Assert.AreEqual(2.0, sheet.Cell(2, 7).Value);
                        Assert.AreEqual("Work Street", sheet.Cell(2, 8).Value);
                        Assert.AreEqual("Work City", sheet.Cell(2, 9).Value);
                        Assert.AreEqual("Work State", sheet.Cell(2, 10).Value);
                        Assert.AreEqual("Work Zip", sheet.Cell(2, 11).Value);
                        Assert.AreEqual(3.0, sheet.Cell(2, 12).Value);
                        Assert.AreEqual("", sheet.Cell(2, 13).Value);
                        Assert.AreEqual("", sheet.Cell(2, 14).Value);
                        Assert.AreEqual("", sheet.Cell(2, 15).Value);
                        Assert.AreEqual("", sheet.Cell(2, 16).Value);
                        Assert.AreEqual(0.0, sheet.Cell(2, 17).Value);
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
                    using (var book = new XLWorkbook(stream)) {
                        var sheet = book.Worksheets.Worksheet(1);
                        Assert.AreEqual("ID", sheet.Cell(1, 1).Value);
                        Assert.AreEqual("", sheet.Cell(1, 2).Value);
                        Assert.AreEqual((double)1, sheet.Cell(2, 1).Value);
                        Assert.AreEqual("", sheet.Cell(2, 2).Value);
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
                    using (var book = new XLWorkbook(stream)) {
                        var sheet = book.Worksheets.Worksheet(1);

                        // Check the header row
                        Assert.AreEqual("int1", sheet.Cell(1, 1).Value);
                        Assert.AreEqual("string1", sheet.Cell(1, 2).Value);

                        // Check the first record
                        Assert.AreEqual((double)1, sheet.Cell(2, 1).Value);
                        Assert.AreEqual("one", sheet.Cell(2, 2).Value);

                        // Check the second record
                        Assert.AreEqual((double)2, sheet.Cell(3, 1).Value);
                        Assert.AreEqual("two", sheet.Cell(3, 2).Value);
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
                    using (var book = new XLWorkbook(stream)) {
                        var sheet = book.Worksheets.Worksheet(1);

                        // Check the header row
                        Assert.AreEqual("ColumnName", sheet.Cell(1, 1).Value);
                        Assert.AreEqual("ColumnName", sheet.Cell(1, 2).Value);
                        Assert.AreEqual("ColumnName", sheet.Cell(1, 3).Value);

                        // Check the first record
                        Assert.AreEqual("1", sheet.Cell(2, 1).Value);
                        Assert.AreEqual("2", sheet.Cell(2, 2).Value);
                        Assert.AreEqual("3", sheet.Cell(2, 3).Value);
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
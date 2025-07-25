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
using ClosedXML.Excel;
using ExcelHelper.Configuration;
using ExcelHelper.TypeConversion;
using NUnit.Framework;
using NUnit.Framework.Legacy;

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
                    excel.SetRowFormat(1, fontStyle: ExcelFontStyle.Bold, fontSize: 16);
                    excel.SetColumnFormat(7, fontStyle: ExcelFontStyle.Italic, fontSize: 24);

                    var date = DateTime.Parse("2017-10-16 03:05 PM");
                    var guid = new Guid("bfb9c599-bc9e-4f97-ae59-25f2ca09cfdf");
                    excel.WriteCell(0, 0, "one");
                    excel.WriteCell(0, 1, "one, two", fontStyle: ExcelFontStyle.Bold);
                    excel.WriteCell(0, 2, "one \"two\" three", fontSize: 18);
                    excel.WriteCell(0, 3, " one ", fontName: "Times New Roman");
                    excel.WriteCell(0, 4, date);
                    excel.WriteCell(0, 5, date, dateFormat: "M/d/yyyy h:mm:ss AM/PM");
                    excel.WriteCell(0, 6, date, dateFormat: "M/d/yyyy");
                    excel.WriteCell(0, 7, date, dateFormat: "dddd, MMMM d, yyyy", fontStyle: ExcelFontStyle.Bold, horizontalAlign: ExcelAlignHorizontal.Right, verticalAlign: ExcelAlignVertical.Bottom);
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
                        ClassicAssert.AreEqual(true, sheet.Row(2).Style.Font.Bold);
                        ClassicAssert.AreEqual(16.0, sheet.Row(2).Style.Font.FontSize);
                        ClassicAssert.AreEqual(true, sheet.Column(8).Style.Font.Italic);
                        ClassicAssert.AreEqual(24.0, sheet.Column(8).Style.Font.FontSize);

                        // Check some automatically sized column widths
                        ClassicAssert.AreEqual(23.71, sheet.Column(3).Width);
                        ClassicAssert.AreEqual(11.209999999999999, sheet.Column(4).Width);

                        // Verify the overridden row and column sizes
                        ClassicAssert.AreEqual(600, sheet.Row(2).Height);
                        ClassicAssert.AreEqual(8.8599999999999994, sheet.Column(2).Width);
                        ClassicAssert.AreEqual(2.1899999999999999, sheet.Column(12).Width);

                        // Verify first row
                        ClassicAssert.AreEqual("one", sheet.Cell(1, 1).Value);
                        ClassicAssert.AreEqual("", sheet.Cell(1, 1).Style.NumberFormat.Format);
                        ClassicAssert.AreEqual("", sheet.Cell(1, 1).Style.DateFormat.Format);
                        ClassicAssert.AreEqual("one, two", sheet.Cell(1, 2).Value);
                        ClassicAssert.AreEqual(true, sheet.Cell(1, 2).Style.Font.Bold);
                        ClassicAssert.AreEqual("one \"two\" three", sheet.Cell(1, 3).Value);
                        ClassicAssert.AreEqual(18.0, sheet.Cell(1, 3).Style.Font.FontSize);
                        ClassicAssert.AreEqual(" one ", sheet.Cell(1, 4).Value);
                        ClassicAssert.AreEqual("Times New Roman", sheet.Cell(1, 4).Style.Font.FontName);
                        ClassicAssert.AreEqual(date, sheet.Cell(1, 5).Value);
                        ClassicAssert.AreEqual("10/16/2017 15:05", sheet.Cell(1, 5).GetFormattedString());
                        ClassicAssert.AreEqual("", sheet.Cell(1, 5).Style.DateFormat.Format);
                        ClassicAssert.AreEqual(date, sheet.Cell(1, 6).Value);
                        ClassicAssert.AreEqual("10/16/2017 3:05:00 PM", sheet.Cell(1, 6).GetFormattedString());
                        ClassicAssert.AreEqual("M/d/yyyy h:mm:ss AM/PM", sheet.Cell(1, 6).Style.DateFormat.Format);
                        ClassicAssert.AreEqual(date, sheet.Cell(1, 7).Value);
                        ClassicAssert.AreEqual("10/16/2017", sheet.Cell(1, 7).GetFormattedString());
                        ClassicAssert.AreEqual("M/d/yyyy", sheet.Cell(1, 7).Style.DateFormat.Format);
                        ClassicAssert.AreEqual(XLAlignmentHorizontalValues.Right, sheet.Cell(1, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right);
                        ClassicAssert.AreEqual(XLAlignmentVerticalValues.Bottom, sheet.Cell(1, 7).Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom);
                        ClassicAssert.AreEqual(date, sheet.Cell(1, 8).Value);
                        ClassicAssert.AreEqual("Monday, October 16, 2017", sheet.Cell(1, 8).GetFormattedString());
                        ClassicAssert.AreEqual("dddd, MMMM d, yyyy", sheet.Cell(1, 8).Style.DateFormat.Format);
                        ClassicAssert.AreEqual(true, sheet.Cell(1, 8).Style.Font.Bold);
                        ClassicAssert.AreEqual((double)1, sheet.Cell(1, 9).Value);
                        ClassicAssert.AreEqual("1", sheet.Cell(1, 9).GetFormattedString());
                        ClassicAssert.AreEqual((double)2, sheet.Cell(1, 10).Value);
                        ClassicAssert.AreEqual("2", sheet.Cell(1, 10).GetFormattedString());
                        ClassicAssert.AreEqual((double)3, sheet.Cell(1, 11).Value);
                        ClassicAssert.AreEqual("3", sheet.Cell(1, 11).GetFormattedString());
                        ClassicAssert.AreEqual("1+2", sheet.Cell(1, 12).FormulaA1);
                        ClassicAssert.AreEqual((double)3, sheet.Cell(1, 12).Value);
                        ClassicAssert.AreEqual("3", sheet.Cell(1, 12).GetFormattedString());

                        // Verify second row
                        ClassicAssert.AreEqual((double)4, sheet.Cell(2, 1).Value);
                        ClassicAssert.AreEqual("4", sheet.Cell(2, 1).GetFormattedString());
                        ClassicAssert.AreEqual("", sheet.Cell(2, 1).Style.NumberFormat.Format);
                        ClassicAssert.AreEqual((double)5, sheet.Cell(2, 2).Value);
                        ClassicAssert.AreEqual("5", sheet.Cell(2, 2).GetFormattedString());
                        ClassicAssert.AreEqual((double)6, sheet.Cell(2, 3).Value);
                        ClassicAssert.AreEqual("6", sheet.Cell(2, 3).GetFormattedString());
                        ClassicAssert.AreEqual(123.456, sheet.Cell(2, 4).Value);
                        ClassicAssert.AreEqual("$123.46", sheet.Cell(2, 4).GetFormattedString());
                        ClassicAssert.AreEqual("$#,##0.00;($#,##0.00)", sheet.Cell(2, 4).Style.NumberFormat.Format);
                        ClassicAssert.AreEqual(-123.456, sheet.Cell(2, 5).Value);
                        ClassicAssert.AreEqual("($123.46)", sheet.Cell(2, 5).GetFormattedString());
                        ClassicAssert.AreEqual("$#,##0.00;($#,##0.00)", sheet.Cell(2, 5).Style.NumberFormat.Format);
                        ClassicAssert.AreEqual("bfb9c599-bc9e-4f97-ae59-25f2ca09cfdf", sheet.Cell(2, 6).Value);
                        ClassicAssert.AreEqual("bfb9c599-bc9e-4f97-ae59-25f2ca09cfdf", sheet.Cell(2, 6).GetFormattedString());
                        ClassicAssert.AreEqual("true", sheet.Cell(2, 7).Value);
                        ClassicAssert.AreEqual("true", sheet.Cell(2, 7).GetFormattedString());
                        ClassicAssert.AreEqual("false", sheet.Cell(2, 8).Value);
                        ClassicAssert.AreEqual("false", sheet.Cell(2, 8).GetFormattedString());
                        ClassicAssert.AreEqual("", sheet.Cell(2, 9).Value);
                        ClassicAssert.AreEqual("01:02:03", sheet.Cell(2, 10).Value);
                        ClassicAssert.AreEqual("01:02:03", sheet.Cell(2, 10).GetFormattedString());
                        ClassicAssert.AreEqual("", sheet.Cell(2, 10).Style.NumberFormat.Format);
                        ClassicAssert.AreEqual("2*3", sheet.Cell(2, 11).FormulaA1);
                        ClassicAssert.AreEqual((double)6, sheet.Cell(2, 11).Value);
                        ClassicAssert.AreEqual("6", sheet.Cell(2, 11).GetFormattedString());

                        // Verify third sheet
                        ClassicAssert.AreEqual(3, book.Worksheets.Count);
                        sheet = book.Worksheets.Worksheet(3);
                        ClassicAssert.AreEqual("third sheet", sheet.Cell(1, 1).Value);
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
                            ClassicAssert.AreEqual(i.ToString(), sheet.Cell(i + 1, 1).Value);
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
                        ClassicAssert.AreEqual(3, book.Worksheets.Count);
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
                ClassicAssert.AreEqual(true, sheet.Row(1).Style.Font.Bold);

                // Check the header row
                ClassicAssert.AreEqual("FirstColumn", sheet.Cell(row, 1).Value);
                ClassicAssert.AreEqual("Int Column", sheet.Cell(row, 2).Value);
                ClassicAssert.AreEqual("StringColumn", sheet.Cell(row, 3).Value);
                ClassicAssert.AreEqual("TypeConvertedColumn", sheet.Cell(row, 4).Value);
                ClassicAssert.AreEqual("BoolColumn", sheet.Cell(row, 5).Value);
                ClassicAssert.AreEqual("DoubleColumn", sheet.Cell(row, 6).Value);
                ClassicAssert.AreEqual("DateTimeColumn", sheet.Cell(row, 7).Value);
                ClassicAssert.AreEqual("NullStringColumn", sheet.Cell(row, 8).Value);
                ClassicAssert.AreEqual("FormulaColumn", sheet.Cell(row, 9).Value);
                row++;
            }

            // Check the first record
            ClassicAssert.AreEqual("first column", sheet.Cell(row, 1).Value);
            ClassicAssert.AreEqual((double)1, sheet.Cell(row, 2).Value);
            ClassicAssert.AreEqual("string column", sheet.Cell(row, 3).Value);
            ClassicAssert.AreEqual("test", sheet.Cell(row, 4).Value);
            ClassicAssert.AreEqual("true", sheet.Cell(row, 5).Value);
            ClassicAssert.AreEqual(12.34, sheet.Cell(row, 6).Value);
            ClassicAssert.AreEqual(date, sheet.Cell(row, 7).Value);
            ClassicAssert.AreEqual("", sheet.Cell(row, 8).Style.DateFormat.Format);    // TODO: Do we need a different default here for dates?
            ClassicAssert.AreEqual("", sheet.Cell(row, 8).Value);
            ClassicAssert.AreEqual("1+2", sheet.Cell(row, 9).FormulaA1);
            ClassicAssert.AreEqual((double)3, sheet.Cell(row, 9).Value);
            row++;

            // Check the second record
            ClassicAssert.AreEqual("first column 2", sheet.Cell(row, 1).Value);
            ClassicAssert.AreEqual((double)2, sheet.Cell(row, 2).Value);
            ClassicAssert.AreEqual("string column 2", sheet.Cell(row, 3).Value);
            ClassicAssert.AreEqual("test", sheet.Cell(row, 4).Value);
            ClassicAssert.AreEqual("false", sheet.Cell(row, 5).Value);
            ClassicAssert.AreEqual(43.21, sheet.Cell(row, 6).Value);
            ClassicAssert.AreEqual(yesterday, sheet.Cell(row, 7).Value);
            ClassicAssert.AreEqual("", sheet.Cell(row, 7).Style.DateFormat.Format);    // TODO: Do we need a different default here for dates?
            ClassicAssert.AreEqual("", sheet.Cell(row, 8).Value);
            ClassicAssert.AreEqual("not a formula", sheet.Cell(row, 9).Value);
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
                        ClassicAssert.AreEqual("Int Column", sheet.Cell(1, 1).Value);
                        ClassicAssert.AreEqual("StringColumn", sheet.Cell(1, 2).Value);
                        ClassicAssert.AreEqual("FirstColumn", sheet.Cell(1, 3).Value);
                        ClassicAssert.AreEqual("TypeConvertedColumn", sheet.Cell(1, 4).Value);

                        // Check the first record
                        ClassicAssert.AreEqual((double)1, sheet.Cell(2, 1).Value);
                        ClassicAssert.AreEqual("string column", sheet.Cell(2, 2).Value);
                        ClassicAssert.AreEqual("first column", sheet.Cell(2, 3).Value);
                        ClassicAssert.AreEqual("test", sheet.Cell(2, 4).Value);
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
                        ClassicAssert.AreEqual("FirstName", sheet.Cell(1, 1).Value);
                        ClassicAssert.AreEqual("LastName", sheet.Cell(1, 2).Value);
                        ClassicAssert.AreEqual("HomeStreet", sheet.Cell(1, 3).Value);
                        ClassicAssert.AreEqual("HomeCity", sheet.Cell(1, 4).Value);
                        ClassicAssert.AreEqual("HomeState", sheet.Cell(1, 5).Value);
                        ClassicAssert.AreEqual("HomeZip", sheet.Cell(1, 6).Value);
                        ClassicAssert.AreEqual("HomeID", sheet.Cell(1, 7).Value);
                        ClassicAssert.AreEqual("WorkStreet", sheet.Cell(1, 8).Value);
                        ClassicAssert.AreEqual("WorkCity", sheet.Cell(1, 9).Value);
                        ClassicAssert.AreEqual("WorkState", sheet.Cell(1, 10).Value);
                        ClassicAssert.AreEqual("WorkZip", sheet.Cell(1, 11).Value);
                        ClassicAssert.AreEqual("WorkID", sheet.Cell(1, 12).Value);
                        ClassicAssert.AreEqual("NullStreet", sheet.Cell(1, 13).Value);
                        ClassicAssert.AreEqual("NullCity", sheet.Cell(1, 14).Value);
                        ClassicAssert.AreEqual("NullState", sheet.Cell(1, 15).Value);
                        ClassicAssert.AreEqual("NullZip", sheet.Cell(1, 16).Value);
                        ClassicAssert.AreEqual("NullID", sheet.Cell(1, 17).Value);

                        // Check the record
                        ClassicAssert.AreEqual("First Name", sheet.Cell(2, 1).Value);
                        ClassicAssert.AreEqual("Last Name", sheet.Cell(2, 2).Value);
                        ClassicAssert.AreEqual("Home Street", sheet.Cell(2, 3).Value);
                        ClassicAssert.AreEqual("Home City", sheet.Cell(2, 4).Value);
                        ClassicAssert.AreEqual("Home State", sheet.Cell(2, 5).Value);
                        ClassicAssert.AreEqual("Home Zip", sheet.Cell(2, 6).Value);
                        ClassicAssert.AreEqual(2.0, sheet.Cell(2, 7).Value);
                        ClassicAssert.AreEqual("Work Street", sheet.Cell(2, 8).Value);
                        ClassicAssert.AreEqual("Work City", sheet.Cell(2, 9).Value);
                        ClassicAssert.AreEqual("Work State", sheet.Cell(2, 10).Value);
                        ClassicAssert.AreEqual("Work Zip", sheet.Cell(2, 11).Value);
                        ClassicAssert.AreEqual(3.0, sheet.Cell(2, 12).Value);
                        ClassicAssert.AreEqual("", sheet.Cell(2, 13).Value);
                        ClassicAssert.AreEqual("", sheet.Cell(2, 14).Value);
                        ClassicAssert.AreEqual("", sheet.Cell(2, 15).Value);
                        ClassicAssert.AreEqual("", sheet.Cell(2, 16).Value);
                        ClassicAssert.AreEqual(0.0, sheet.Cell(2, 17).Value);
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
                        ClassicAssert.AreEqual("ID", sheet.Cell(1, 1).Value);
                        ClassicAssert.AreEqual("", sheet.Cell(1, 2).Value);
                        ClassicAssert.AreEqual((double)1, sheet.Cell(2, 1).Value);
                        ClassicAssert.AreEqual("", sheet.Cell(2, 2).Value);
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
                        ClassicAssert.AreEqual("int1", sheet.Cell(1, 1).Value);
                        ClassicAssert.AreEqual("string1", sheet.Cell(1, 2).Value);

                        // Check the first record
                        ClassicAssert.AreEqual((double)1, sheet.Cell(2, 1).Value);
                        ClassicAssert.AreEqual("one", sheet.Cell(2, 2).Value);

                        // Check the second record
                        ClassicAssert.AreEqual((double)2, sheet.Cell(3, 1).Value);
                        ClassicAssert.AreEqual("two", sheet.Cell(3, 2).Value);
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
                        ClassicAssert.AreEqual("ColumnName", sheet.Cell(1, 1).Value);
                        ClassicAssert.AreEqual("ColumnName", sheet.Cell(1, 2).Value);
                        ClassicAssert.AreEqual("ColumnName", sheet.Cell(1, 3).Value);

                        // Check the first record
                        ClassicAssert.AreEqual("1", sheet.Cell(2, 1).Value);
                        ClassicAssert.AreEqual("2", sheet.Cell(2, 2).Value);
                        ClassicAssert.AreEqual("3", sheet.Cell(2, 3).Value);
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
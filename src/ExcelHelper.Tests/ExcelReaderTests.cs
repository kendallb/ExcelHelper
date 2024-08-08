/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ExcelHelper.Configuration;
using ExcelHelper.TypeConversion;
using NUnit.Framework;
using NUnit.Framework.Legacy;

// ReSharper disable ReturnValueOfPureMethodIsNotUsed
// ReSharper disable UnusedAutoPropertyAccessor.Local
// ReSharper disable ClassNeverInstantiated.Local
// ReSharper disable UnusedMember.Local

// TODO: Use the C1 libraries for unit testing to convert from OpenXML test file data to BIFF8 in memory ...

namespace ExcelHelper.Tests
{
    [TestFixture]
    public class ExcelReaderTests
    {
        private ExcelFactory _factory;

        [SetUp]
        public void SetUp()
        {
            _factory = new ExcelFactory();
        }

        [Test]
        public void ReadCellTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                const double n = 1.2;
                const int nsi = 3;
                const double ns = 2.1;
                var d = DateTime.Today;
                const char c = 'c';
                var guid = Guid.NewGuid();
                var ts = new TimeSpan(45, 2, 3, 4, 5);
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");
                    sheet.Cell(1, 1).SetValue(n);
                    sheet.Cell(1, 2).SetValue(nsi.ToString());
                    sheet.Cell(1, 3).SetValue(ns.ToString());
                    sheet.Cell(1, 4).SetValue(d);
                    sheet.Cell(1, 5).SetValue(d.ToString());
                    sheet.Cell(1, 6).SetValue(true);
                    sheet.Cell(1, 7).SetValue("true");
                    sheet.Cell(1, 8).SetValue("yes");
                    sheet.Cell(1, 9).SetValue(c);
                    sheet.Cell(1, 10).SetValue((object)null);
                    sheet.Cell(1, 11).SetValue(guid.ToString());
                    sheet.Cell(1, 12).SetValue(ts);
                    sheet.Cell(1, 13).SetValue(ts.ToString());
                    sheet.Cell(1, 14).SetValue(ts.ToString());
                    book.AddWorksheet("Sheet 2");
                    sheet = book.AddWorksheet("Sheet 3");
                    sheet.Cell(1, 1).SetValue("third sheet");
                    book.SaveAs(stream);
                }

                // Now parse the Excel file as all available types
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    // Check the column and row counts are correct
                    ClassicAssert.AreEqual(14, excel.TotalColumns);

                    // Test all number conversions
                    if (!excel.ReadRow()) {
                        throw new ArgumentException();
                    }
                    ClassicAssert.AreEqual((sbyte)n, excel.GetColumn<sbyte>(0));
                    ClassicAssert.AreEqual((short)n, excel.GetColumn<short>(0));
                    ClassicAssert.AreEqual((int)n, excel.GetColumn<int>(0));
                    ClassicAssert.AreEqual((long)n, excel.GetColumn<long>(0));
                    ClassicAssert.AreEqual((byte)n, excel.GetColumn<byte>(0));
                    ClassicAssert.AreEqual((ushort)n, excel.GetColumn<ushort>(0));
                    ClassicAssert.AreEqual((uint)n, excel.GetColumn<uint>(0));
                    ClassicAssert.AreEqual((ulong)n, excel.GetColumn<ulong>(0));
                    ClassicAssert.AreEqual((float)n, excel.GetColumn<float>(0));
                    ClassicAssert.AreEqual(n, excel.GetColumn<double>(0));
                    ClassicAssert.AreEqual((decimal)n, excel.GetColumn<decimal>(0));
                    ClassicAssert.AreEqual(n.ToString(), excel.GetColumn<string>(0));

                    // Test all number conversions with a string cell
                    ClassicAssert.AreEqual((sbyte)nsi, excel.GetColumn<sbyte>(1));
                    ClassicAssert.AreEqual((short)nsi, excel.GetColumn<short>(1));
                    ClassicAssert.AreEqual(nsi, excel.GetColumn<int>(1));
                    ClassicAssert.AreEqual(nsi, excel.GetColumn<long>(1));
                    ClassicAssert.AreEqual((byte)nsi, excel.GetColumn<byte>(1));
                    ClassicAssert.AreEqual((ushort)nsi, excel.GetColumn<ushort>(1));
                    ClassicAssert.AreEqual((uint)nsi, excel.GetColumn<uint>(1));
                    ClassicAssert.AreEqual((ulong)nsi, excel.GetColumn<ulong>(1));
                    ClassicAssert.AreEqual((float)ns, excel.GetColumn<float>(2));
                    ClassicAssert.AreEqual(ns, excel.GetColumn<double>(2));
                    ClassicAssert.AreEqual((decimal)ns, excel.GetColumn<decimal>(2));
                    ClassicAssert.AreEqual(nsi.ToString(), excel.GetColumn<string>(1));
                    ClassicAssert.AreEqual(ns.ToString(), excel.GetColumn<string>(2));

                    // Test dates
                    ClassicAssert.AreEqual(d, excel.GetColumn<DateTime>(3));
                    ClassicAssert.AreEqual(d, excel.GetColumn<DateTime>(4));

                    // Test boolean
                    ClassicAssert.AreEqual(true, excel.GetColumn<bool>(5));
                    ClassicAssert.AreEqual("True", excel.GetColumn<string>(5));
                    ClassicAssert.AreEqual(true, excel.GetColumn<bool>(6));
                    ClassicAssert.AreEqual("true", excel.GetColumn<string>(6));
                    ClassicAssert.AreEqual(true, excel.GetColumn<bool>(7));
                    ClassicAssert.AreEqual("yes", excel.GetColumn<string>(7));

                    // Test character
                    ClassicAssert.AreEqual('c', excel.GetColumn<char>(8));
                    ClassicAssert.AreEqual("c", excel.GetColumn<string>(8));

                    // Test null
                    ClassicAssert.AreEqual("", excel.GetColumn<string>(9));
                    ClassicAssert.AreEqual(null, excel.GetColumn<int?>(9));
                    ClassicAssert.AreEqual(DateTime.MinValue, excel.GetColumn<DateTime>(9));

                    // Test guid
                    ClassicAssert.AreEqual(guid, excel.GetColumn<Guid>(10));
                    ClassicAssert.AreEqual(guid.ToString(), excel.GetColumn<string>(10));

                    // Test TimeSpan
                    ClassicAssert.AreEqual(ts, excel.GetColumn<TimeSpan>(11));
                    // TODO: This won't work until ExcelDataReader is changed to natively parse TimeSpans
                    //Assert.AreEqual(ts.ToString(), excel.GetColumn<string>(11));
                    ClassicAssert.AreEqual(ts, excel.GetColumn<TimeSpan>(12));
                    ClassicAssert.AreEqual(ts.ToString(), excel.GetColumn<string>(12));
                    ClassicAssert.AreEqual(ts, excel.GetColumn<TimeSpan>(13));
                    ClassicAssert.AreEqual(ts.ToString(), excel.GetColumn<string>(13));

                    // Test the third sheet
                    ClassicAssert.AreEqual(3, excel.TotalSheets);
                    ClassicAssert.IsTrue(excel.ChangeSheet(2));
                    ClassicAssert.IsFalse(excel.ChangeSheet(3));
                    if (!excel.ReadRow()) {
                        throw new ArgumentException();
                    }
                    ClassicAssert.AreEqual("third sheet", excel.GetColumn<string>(0));
                }
            }
        }

        [Test]
        public void LargeFileTest()
        {
            using (var stream = new MemoryStream()) {
                // Write out 66K rows
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");
                    for (var i = 0; i < 66000; i++) {
                        sheet.Cell(i + 1, 1).SetValue(i);
                    }
                    book.SaveAs(stream);
                }

                // Now read it back
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    // Check the column and row counts are correct
                    ClassicAssert.AreEqual(1, excel.TotalColumns);

                    // Verify 66K rows
                    for (var i = 0; i < 66000; i++) {
                        if (!excel.ReadRow()) {
                            throw new ArgumentException();
                        }
                        ClassicAssert.AreEqual(i, excel.GetColumn<int>(0));
                    }
                }
            }
        }

        [Test]
        public void GetRecordsMissingFieldsThrowsErrorTest()
        {
            using (var stream = new MemoryStream()) {
                // Create an empty book with only one field
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");
                    sheet.Cell(1, 1).SetValue("FirstColumn");
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    excel.Configuration.RegisterClassMap<TestRecordMap>();
                    try {
                        excel.GetRecords<TestRecord>().ToList();
                        throw new Exception("We should not get here!");
                    } catch (ExcelMissingFieldException) {
                        // We should get here as there are no fields found in the mapping
                    }
                }
            }
        }

        [Test]
        public void GetRecordsTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                var date = DateTime.Today;
                var guid = Guid.NewGuid();
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");
                    WriteRecords(sheet, guid, date);
                    book.AddWorksheet("Sheet 2");
                    sheet = book.AddWorksheet("Sheet 3");
                    WriteRecords(sheet, guid, date);

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    excel.Configuration.RegisterClassMap<TestRecordMap>();
                    ValidateRecords(excel, guid, date);
                    ClassicAssert.AreEqual(3, excel.TotalSheets);
                    ClassicAssert.IsTrue(excel.ChangeSheet(2));
                    ClassicAssert.IsFalse(excel.ChangeSheet(3));
                    ValidateRecords(excel, guid, date);
                }
            }
        }

        [Test]
        public void GetRecordsWithEmptyRowTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                var date = DateTime.Today;
                var guid = Guid.NewGuid();
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");
                    WriteRecords(sheet, guid, date, includeBlankRow: true);
                    book.AddWorksheet("Sheet 2");
                    sheet = book.AddWorksheet("Sheet 3");
                    WriteRecords(sheet, guid, date, includeBlankRow: true);

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    excel.Configuration.RegisterClassMap<TestRecordMap>();
                    ValidateRecords(excel, guid, date, ignoreEmptyRows: true);
                    ClassicAssert.AreEqual(3, excel.TotalSheets);
                    ClassicAssert.IsTrue(excel.ChangeSheet(2));
                    ClassicAssert.IsFalse(excel.ChangeSheet(3));
                    ValidateRecords(excel, guid, date, ignoreEmptyRows: true);
                }
            }
        }

        [Test]
        public void SkipRowsTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                var date = DateTime.Today;
                var guid = Guid.NewGuid();
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");
                    WriteRecords(sheet, guid, date, null, 2);
                    book.AddWorksheet("Sheet 2");
                    sheet = book.AddWorksheet("Sheet 3");
                    WriteRecords(sheet, guid, date);

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    excel.Configuration.RegisterClassMap<TestRecordMap>();
                    excel.SkipRows(2);
                    ValidateRecords(excel, guid, date);
                    ClassicAssert.AreEqual(3, excel.TotalSheets);
                    ClassicAssert.IsTrue(excel.ChangeSheet(2));
                    ClassicAssert.IsFalse(excel.ChangeSheet(3));
                    ValidateRecords(excel, guid, date);
                }
            }
        }

        [Test]
        public void SheetNameTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                var date = DateTime.Today;
                var guid = Guid.NewGuid();
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");
                    sheet.Name = "Test Sheet 1";
                    WriteRecords(sheet, guid, date);
                    book.AddWorksheet("Sheet 2");
                    sheet = book.AddWorksheet("Sheet 3");
                    sheet.Name = "Test Sheet 2";
                    WriteRecords(sheet, guid, date);

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    excel.Configuration.RegisterClassMap<TestRecordMap>();
                    ValidateRecords(excel, guid, date);
                    ClassicAssert.AreEqual(excel.SheetName, "Test Sheet 1");
                    ClassicAssert.AreEqual(3, excel.TotalSheets);
                    ClassicAssert.IsTrue(excel.ChangeSheet(2));
                    ClassicAssert.AreEqual(excel.SheetName, "Test Sheet 2");
                    ClassicAssert.IsFalse(excel.ChangeSheet(3));
                    ValidateRecords(excel, guid, date);
                }
            }
        }

        [Test]
        public void GetRecordsOptionalFieldTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                var date = DateTime.Today;
                var guid = Guid.NewGuid();
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");
                    WriteRecords(sheet, guid, date, "optional");
                    book.AddWorksheet("Sheet 2");
                    sheet = book.AddWorksheet("Sheet 3");
                    WriteRecords(sheet, guid, date, "optional");

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    excel.Configuration.RegisterClassMap<TestRecordMap>();
                    ValidateRecords(excel, guid, date, "optional");
                    ClassicAssert.AreEqual(3, excel.TotalSheets);
                    ClassicAssert.IsTrue(excel.ChangeSheet(2));
                    ClassicAssert.IsFalse(excel.ChangeSheet(3));
                    ValidateRecords(excel, guid, date, "optional");
                }
            }
        }

        [Test]
        public void GetRecordsMissingFieldTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                var date = DateTime.Today;
                var guid = Guid.NewGuid();
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");
                    WriteRecords(sheet, guid, date);
                    book.AddWorksheet("Sheet 2");
                    sheet = book.AddWorksheet("Sheet 3");
                    WriteRecords(sheet, guid, date);

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    excel.Configuration.WillThrowOnMissingHeader = false;
                    excel.Configuration.RegisterClassMap<TestRecordMapMissingField>();
                    ValidateRecords(excel, guid, date);
                    ClassicAssert.AreEqual(3, excel.TotalSheets);
                    ClassicAssert.IsTrue(excel.ChangeSheet(2));
                    ClassicAssert.IsFalse(excel.ChangeSheet(3));
                    ValidateRecords(excel, guid, date);
                }
            }
        }

        /// <summary>
        /// Writes some test record data to the Excel file
        /// </summary>
        /// <param name="sheet">Sheet to write to</param>
        /// <param name="guid">GUID for the test</param>
        /// <param name="date">Date for the test</param>
        /// <param name="optionalReadValue">Value to put into the optional read column, null to not include it</param>
        /// <param name="firstRow">The first row number</param>
        /// <param name="includeBlankRow">True to include a blank row</param>
        private static void WriteRecords(
            IXLWorksheet sheet,
            Guid guid,
            DateTime date,
            string optionalReadValue = null,
            int firstRow = 0,
            bool includeBlankRow = false)
        {
            // Write the header fields
            var row = firstRow + 1;
            sheet.Cell(row, 1).SetValue("FirstColumn");
            sheet.Cell(row, 2).SetValue("TypeConvertedColumn");
            sheet.Cell(row, 3).SetValue("IntColumn");
            sheet.Cell(row, 4).SetValue("String Column");
            sheet.Cell(row, 5).SetValue("GuidColumn");
            sheet.Cell(row, 6).SetValue("BoolColumn");
            sheet.Cell(row, 7).SetValue("DoubleColumn");
            sheet.Cell(row, 8).SetValue("DateTimeColumn");
            sheet.Cell(row, 9).SetValue("NullStringColumn");
            if (optionalReadValue != null) {
                sheet.Cell(row, 10).SetValue("OptionalReadColumn");
            }

            // Write the first record
            row++;
            sheet.Cell(row, 1).SetValue(1);
            sheet.Cell(row, 2).SetValue("converts to test");
            sheet.Cell(row, 3).SetValue(1 * 2);
            sheet.Cell(row, 4).SetValue("string column 1");
            sheet.Cell(row, 5).SetValue(guid.ToString());
            sheet.Cell(row, 6).SetValue(true);
            sheet.Cell(row, 7).SetValue(1 * 3.0);
            sheet.Cell(row, 8).SetValue(date.AddDays(1));
            sheet.Cell(row, 9).SetValue((object)null);
            if (optionalReadValue != null) {
                sheet.Cell(row, 10).SetValue(optionalReadValue);
            }

            // Include a blank row in the middle
            if (includeBlankRow) {
                row++;
            }

            // Write the second record
            row++;
            sheet.Cell(row, 1).SetValue(2);
            sheet.Cell(row, 2).SetValue("converts to test");
            sheet.Cell(row, 3).SetValue(2 * 2);
            sheet.Cell(row, 4).SetValue("string column 2");
            sheet.Cell(row, 5).SetValue(guid.ToString());
            sheet.Cell(row, 6).SetValue(false);
            sheet.Cell(row, 7).SetValue(2 * 3.0);
            sheet.Cell(row, 8).SetValue(date.AddDays(2));
            sheet.Cell(row, 9).SetValue((object)null);
            if (optionalReadValue != null) {
                sheet.Cell(row, 10).SetValue(optionalReadValue);
            }

            // Write a blank field outside of the header count. To make sure we only
            // process the columns up to the header count width
            if (!includeBlankRow) {
                sheet.Cell(2, optionalReadValue == null ? 10 : 11).SetValue("");
            }
        }

        /// <summary>
        /// Validate the records read from the sheet
        /// </summary>
        /// <param name="excel">Excel reader to read from</param>
        /// <param name="guid">GUID for the test</param>
        /// <param name="date">Date for the test</param>
        /// <param name="optionalReadValue">Value to expect in the optional read column</param>
        /// <param name="ignoreEmptyRows">True to ignore empty rows</param>
        private static void ValidateRecords(
            IExcelReader excel,
            Guid guid,
            DateTime date,
            string optionalReadValue = null,
            bool ignoreEmptyRows = false)
        {
            // Set the ignore empty rows field
            excel.Configuration.IgnoreEmptyRows = ignoreEmptyRows;

            // Make sure we got two records
            var records = excel.GetRecords<TestRecord>().ToList();
            ClassicAssert.AreEqual(2, records.Count);

            // Verify the records are what we expect
            for (var i = 1; i <= records.Count; i++) {
                var record = records[i - 1];
                ClassicAssert.AreEqual(i, record.FirstColumn);
                ClassicAssert.AreEqual(i * 2, record.IntColumn);
                ClassicAssert.AreEqual("string column " + i, record.StringColumn);
                ClassicAssert.AreEqual("test", record.TypeConvertedColumn);
                ClassicAssert.AreEqual(guid, record.GuidColumn);
                ClassicAssert.AreEqual(0, record.NoMatchingField);
                ClassicAssert.AreEqual(i == 1, record.BoolColumn);
                ClassicAssert.AreEqual(i * 3.0, record.DoubleColumn);
                ClassicAssert.AreEqual(date.AddDays(i), record.DateTimeColumn);
                ClassicAssert.AreEqual("", record.NullStringColumn);
                ClassicAssert.AreEqual(optionalReadValue, record.OptionalReadColumn);
            }

            // Validate the mapped columns
            var columns = excel.GetImportedColumns();

            // Make sure we have the column count we expect
            ClassicAssert.AreEqual(optionalReadValue == null ? 9 : 10, columns.Count);
            ClassicAssert.AreEqual("TestRecord", columns[0].DeclaringType.Name);
            ClassicAssert.AreEqual("IntColumn", columns[0].Name);
            ClassicAssert.AreEqual("FirstColumn", columns[1].Name);
            ClassicAssert.AreEqual("StringColumn", columns[2].Name);
            ClassicAssert.AreEqual("TypeConvertedColumn", columns[3].Name);
            ClassicAssert.AreEqual("GuidColumn", columns[4].Name);
            ClassicAssert.AreEqual("BoolColumn", columns[5].Name);
            ClassicAssert.AreEqual("DoubleColumn", columns[6].Name);
            ClassicAssert.AreEqual("DateTimeColumn", columns[7].Name);
            ClassicAssert.AreEqual("NullStringColumn", columns[8].Name);
            if (optionalReadValue != null) {
                ClassicAssert.AreEqual("OptionalReadColumn", columns[9].Name);
            }
        }

        [Test]
        public void GetRecordsWithReferencesTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");

                    // Write the header fields
                    sheet.Cell(1, 1).SetValue("FirstName");
                    sheet.Cell(1, 2).SetValue("LastName");
                    sheet.Cell(1, 3).SetValue("HomeStreet");
                    sheet.Cell(1, 4).SetValue("HomeCity");
                    sheet.Cell(1, 5).SetValue("HomeState");
                    sheet.Cell(1, 6).SetValue("HomeZip");
                    sheet.Cell(1, 7).SetValue("WorkStreet");
                    sheet.Cell(1, 8).SetValue("WorkCity");
                    sheet.Cell(1, 9).SetValue("WorkState");
                    sheet.Cell(1, 10).SetValue("WorkZip");

                    // Write out a record
                    sheet.Cell(2, 1).SetValue("First Name");
                    sheet.Cell(2, 2).SetValue("Last Name");
                    sheet.Cell(2, 3).SetValue("Home Street");
                    sheet.Cell(2, 4).SetValue("Home City");
                    sheet.Cell(2, 5).SetValue("Home State");
                    sheet.Cell(2, 6).SetValue("Home Zip");
                    sheet.Cell(2, 7).SetValue("Work Street");
                    sheet.Cell(2, 8).SetValue("Work City");
                    sheet.Cell(2, 9).SetValue("Work State");
                    sheet.Cell(2, 10).SetValue("Work Zip");

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    excel.Configuration.RegisterClassMap<PersonMap>();
                    var records = excel.GetRecords<Person>().ToList();

                    // Make sure we got our record
                    ClassicAssert.AreEqual(1, records.Count);
                    var record = records[0];
                    ClassicAssert.AreEqual("First Name", record.FirstName);
                    ClassicAssert.AreEqual("Last Name", record.LastName);
                    ClassicAssert.AreEqual("Home Street", record.HomeAddress.Street);
                    ClassicAssert.AreEqual("Home City", record.HomeAddress.City);
                    ClassicAssert.AreEqual("Home State", record.HomeAddress.State);
                    ClassicAssert.AreEqual("Home Zip", record.HomeAddress.Zip);
                    ClassicAssert.AreEqual("Work Street", record.WorkAddress.Street);
                    ClassicAssert.AreEqual("Work City", record.WorkAddress.City);
                    ClassicAssert.AreEqual("Work State", record.WorkAddress.State);
                    ClassicAssert.AreEqual("Work Zip", record.WorkAddress.Zip);
                }
            }
        }

        [Test]
        public void GetRecordsFailsWithMissingHeadersTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");

                    // Write the header fields
                    sheet.Cell(1, 1).SetValue("FirstColumn");
                    sheet.Cell(2, 1).SetValue(1);
                    sheet.Cell(3, 1).SetValue(2);

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    try {
                        excel.GetRecords<TestRecord>().ToList();
                        Assert.Fail();
                    } catch (ExcelReaderException ex) {
                        ClassicAssert.AreEqual("Field 'IntColumn' does not exist in the Excel file.", ex.Message);
                    }
                }
            }
        }

        [Test]
        public void GetRecordsWithDuplicateHeaderNames()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");

                    // Write the header fields
                    sheet.Cell(1, 1).SetValue("Column");
                    sheet.Cell(1, 2).SetValue("Column");
                    sheet.Cell(1, 3).SetValue("Column");

                    // Write the first record
                    sheet.Cell(2, 1).SetValue("one");
                    sheet.Cell(2, 2).SetValue("two");
                    sheet.Cell(2, 3).SetValue("three");

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    excel.Configuration.RegisterClassMap<TestRecordDuplicateHeaderNamesMap>();
                    var records = excel.GetRecords<TestRecordDuplicateHeaderNames>().ToList();

                    // Make sure we got the correct data
                    ClassicAssert.AreEqual(1, records.Count);
                    var record = records[0];
                    ClassicAssert.AreEqual("one", record.Column1);
                    ClassicAssert.AreEqual("two", record.Column2);
                    ClassicAssert.AreEqual("three", record.Column3);
                }
            }
        }

        [Test]
        public void GetRecordsWithMultipleHeaderNames()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");

                    // Write the header fields
                    sheet.Cell(1, 1).SetValue("int2");
                    sheet.Cell(1, 2).SetValue("string3");

                    // Write the first record
                    sheet.Cell(2, 1).SetValue(1);
                    sheet.Cell(2, 2).SetValue("one");

                    // Write the second record
                    sheet.Cell(3, 1).SetValue(2);
                    sheet.Cell(3, 2).SetValue("two");

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    excel.Configuration.RegisterClassMap<MultipleNamesClassMap>();
                    var records = excel.GetRecords<MultipleNamesClass>().ToList();

                    // Make sure we got the correct data
                    ClassicAssert.AreEqual(2, records.Count);
                    ClassicAssert.AreEqual(1, records[0].IntColumn);
                    ClassicAssert.AreEqual("one", records[0].StringColumn);
                    ClassicAssert.AreEqual(2, records[1].IntColumn);
                    ClassicAssert.AreEqual("two", records[1].StringColumn);
                }
            }
        }

        [Test]
        public void GetRecordEmptyFileFailsTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                using (var book = new XLWorkbook()) {
                    book.AddWorksheet("Sheet 1");
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    try {
                        excel.GetRecords<TestRecord>().ToList();
                        Assert.Fail();
                    } catch (ExcelReaderException ex) {
                        ClassicAssert.AreEqual("No header record was found.", ex.Message);
                    }
                }
            }
        }

        [Test]
        public void GetRecordEmptyValuesNullableTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                var guid = Guid.NewGuid();
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");

                    // Write the header fields
                    sheet.Cell(1, 1).SetValue("StringColumn");
                    sheet.Cell(1, 2).SetValue("IntColumn");
                    sheet.Cell(1, 3).SetValue("GuidColumn");

                    // Write the first record
                    sheet.Cell(2, 1).SetValue("string");
                    sheet.Cell(2, 2).SetValue((object)null);
                    sheet.Cell(2, 3).SetValue((object)null);

                    // Write the second record
                    sheet.Cell(3, 1).SetValue((object)null);
                    sheet.Cell(3, 2).SetValue(2);
                    sheet.Cell(3, 3).SetValue(guid);

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file. Note that we are unable to write NULL strings with ClosedXML so they are always blank.
                // If we wish to test this we should create some real Excel files with nulls in them using the C1 library and save
                // them to disk to use directly rather than building them on the fly here with ClosedXML.
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    var records = excel.GetRecords<TestNullable>().ToList();

                    // Make sure we got two records
                    ClassicAssert.AreEqual(2, records.Count);

                    // Verify the records are what we expect
                    var record = records[0];
                    ClassicAssert.AreEqual("string", record.StringColumn);
                    ClassicAssert.AreEqual(null, record.IntColumn);
                    ClassicAssert.AreEqual(null, record.GuidColumn);

                    record = records[1];
                    ClassicAssert.AreEqual("", record.StringColumn);
                    ClassicAssert.AreEqual(2, record.IntColumn);
                    ClassicAssert.AreEqual(guid, record.GuidColumn);
                }
            }
        }

        [Test]
        public void CaseInsensitiveHeaderMatchingTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");

                    // Write the header fields
                    sheet.Cell(1, 1).SetValue("stringcolumn");
                    sheet.Cell(1, 2).SetValue("intcolumn");

                    // Write the first record
                    sheet.Cell(2, 1).SetValue("string");
                    sheet.Cell(2, 2).SetValue(1);

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    excel.Configuration.WillThrowOnMissingHeader = false;
                    excel.Configuration.IsHeaderCaseSensitive = false;
                    var records = excel.GetRecords<TestRecord>().ToList();

                    // Verify the records are what we expect
                    ClassicAssert.AreEqual(1, records.Count);
                    var record = records[0];
                    ClassicAssert.AreEqual("string", record.StringColumn);
                    ClassicAssert.AreEqual(1, record.IntColumn);
                }
            }
        }

        [Test]
        public void SpacesInHeaderTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");

                    // Write the header fields
                    sheet.Cell(1, 1).SetValue(" String Column ");
                    sheet.Cell(1, 2).SetValue(" Int Column ");

                    // Write the first record
                    sheet.Cell(2, 1).SetValue("string");
                    sheet.Cell(2, 2).SetValue(1);

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    excel.Configuration.WillThrowOnMissingHeader = false;
                    excel.Configuration.IgnoreHeaderWhiteSpace = true;
                    var records = excel.GetRecords<TestRecord>().ToList();

                    // Verify the records are what we expect
                    ClassicAssert.AreEqual(1, records.Count);
                    var record = records[0];
                    ClassicAssert.AreEqual("string", record.StringColumn);
                    ClassicAssert.AreEqual(1, record.IntColumn);
                }
            }
        }

        [Test]
        public void TrimHeadersTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");

                    // Write the header fields
                    sheet.Cell(1, 1).SetValue(" IntColumn ");
                    sheet.Cell(1, 2).SetValue(" String Column ");

                    // Write the first record
                    sheet.Cell(2, 1).SetValue(1);
                    sheet.Cell(2, 2).SetValue("string");

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    excel.Configuration.WillThrowOnMissingHeader = false;
                    excel.Configuration.TrimHeaders = true;
                    excel.Configuration.RegisterClassMap<TestRecordMapMissingField>();
                    var records = excel.GetRecords<TestRecord>().ToList();

                    // Verify the records are what we expect
                    ClassicAssert.AreEqual(1, records.Count);
                    var record = records[0];
                    ClassicAssert.AreEqual(1, record.IntColumn);
                    ClassicAssert.AreEqual("string", record.StringColumn);
                }
            }
        }

        [Test]
        public void DefaultValueTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");

                    // Write the header fields
                    sheet.Cell(1, 1).SetValue("IntColumn");
                    sheet.Cell(1, 2).SetValue("StringColumn");

                    // Write the first record
                    sheet.Cell(2, 1).SetValue((object)null);
                    sheet.Cell(2, 2).SetValue("some string");

                    // Write the second record
                    sheet.Cell(3, 1).SetValue(1);
                    sheet.Cell(3, 2).SetValue((object)null);

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    excel.Configuration.RegisterClassMap<TestDefaultValuesMap>();
                    var records = excel.GetRecords<TestDefaultValues>().ToList();

                    // Verify the records are what we expect
                    ClassicAssert.AreEqual(2, records.Count);
                    var record = records[0];
                    ClassicAssert.AreEqual(-1, record.IntColumn);
                    ClassicAssert.AreEqual("some string", record.StringColumn);
                    record = records[1];
                    ClassicAssert.AreEqual(1, record.IntColumn);
                    ClassicAssert.AreEqual("default string", record.StringColumn);
                }
            }
        }

        [Test]
        public void BooleanTypeConverterTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");

                    // Write the header fields
                    sheet.Cell(1, 1).SetValue("BoolColumn");

                    // Write the test records
                    sheet.Cell(2, 1).SetValue(true);
                    sheet.Cell(3, 1).SetValue("true");
                    sheet.Cell(4, 1).SetValue("True");
                    sheet.Cell(5, 1).SetValue("yes");
                    sheet.Cell(6, 1).SetValue("y");
                    sheet.Cell(7, 1).SetValue(false);
                    sheet.Cell(8, 1).SetValue("false");
                    sheet.Cell(9, 1).SetValue("False");
                    sheet.Cell(10, 1).SetValue("no");
                    sheet.Cell(11, 1).SetValue("n");

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    var records = excel.GetRecords<TestBoolean>().ToList();

                    // Verify the records are what we expect
                    ClassicAssert.AreEqual(10, records.Count);
                    for (var i = 0; i < 5; i++) {
                        ClassicAssert.AreEqual(true, records[i].BoolColumn);
                    }
                    for (var i = 5; i < 10; i++) {
                        ClassicAssert.AreEqual(false, records[i].BoolColumn);
                    }
                }
            }
        }

        [Test]
        public void IgnoreExceptionsTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                var date = DateTime.Today;
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");

                    // Write the header fields
                    sheet.Cell(1, 1).SetValue("StringColumn");
                    sheet.Cell(1, 2).SetValue("IntColumn");
                    sheet.Cell(1, 3).SetValue("BoolColumn");
                    sheet.Cell(1, 4).SetValue("DateTimeColumn");

                    // Write the first record
                    sheet.Cell(2, 1).SetValue("string");
                    sheet.Cell(2, 2).SetValue(1);
                    sheet.Cell(2, 3).SetValue("bullshit");
                    sheet.Cell(2, 4).SetValue("bullshit");

                    // Write the second record
                    sheet.Cell(3, 1).SetValue("string");
                    sheet.Cell(3, 2).SetValue(1);
                    sheet.Cell(3, 3).SetValue(true);
                    sheet.Cell(3, 4).SetValue("bullshit");

                    // Write the third record
                    sheet.Cell(4, 1).SetValue("string");
                    sheet.Cell(4, 2).SetValue(1);
                    sheet.Cell(4, 3).SetValue(true);
                    sheet.Cell(4, 4).SetValue(date);

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    excel.Configuration.WillThrowOnMissingHeader = false;
                    excel.Configuration.IgnoreReadingExceptions = true;
                    var allDetails = new List<ExcelReadErrorDetails>();
                    var exceptions = new List<Exception>();
                    excel.Configuration.ReadingExceptionCallback = (ex, d) => {
                        exceptions.Add(ex);
                        allDetails.Add(d);
                    };
                    var records = excel.GetRecords<TestRecord>().ToList();

                    // Verify the records are what we expect
                    ClassicAssert.AreEqual(1, records.Count);
                    var record = records[0];
                    ClassicAssert.AreEqual("string", record.StringColumn);
                    ClassicAssert.AreEqual(1, record.IntColumn);
                    ClassicAssert.AreEqual(true, record.BoolColumn);
                    ClassicAssert.AreEqual(date, record.DateTimeColumn);

                    // Check we got the information we need about the parse errors
                    ClassicAssert.AreEqual(2, allDetails.Count);
                    var details = allDetails[0];
                    ClassicAssert.AreEqual(2, details.Row);
                    ClassicAssert.AreEqual(3, details.Column);
                    ClassicAssert.AreEqual("BoolColumn", details.FieldName);
                    ClassicAssert.AreEqual("bullshit", details.FieldValue);
                    details = allDetails[1];
                    ClassicAssert.AreEqual(3, details.Row);
                    ClassicAssert.AreEqual(4, details.Column);
                    ClassicAssert.AreEqual("DateTimeColumn", details.FieldName);
                    ClassicAssert.AreEqual("bullshit", details.FieldValue);

                    // Check the exception details are what we expect
                    ClassicAssert.AreEqual(2, exceptions.Count);
                    var message =
                        @"Type: 'ExcelHelper.Tests.ExcelReaderTests+TestRecord'" + "\r\n" +
                        @"Row: '2' (1 based)" + "\r\n" +
                        @"Column: '3' (1 based)" + "\r\n" +
                        @"Field Name: 'BoolColumn'" + "\r\n" +
                        @"Field Value: 'bullshit'" + "\r\n";
                    ClassicAssert.AreEqual(message, exceptions[0].Data["ExcelHelper"]);
                    message =
                        @"Type: 'ExcelHelper.Tests.ExcelReaderTests+TestRecord'" + "\r\n" +
                        @"Row: '3' (1 based)" + "\r\n" +
                        @"Column: '4' (1 based)" + "\r\n" +
                        @"Field Name: 'DateTimeColumn'" + "\r\n" +
                        @"Field Value: 'bullshit'" + "\r\n";
                    ClassicAssert.AreEqual(message, exceptions[1].Data["ExcelHelper"]);
                }
            }
        }

        [Test]
        public void ReadStructRecordsTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");

                    // Write the header fields
                    sheet.Cell(1, 1).SetValue("ID");
                    sheet.Cell(1, 2).SetValue("Name");

                    // Write the first record
                    sheet.Cell(2, 1).SetValue(1);
                    sheet.Cell(2, 2).SetValue("a name");

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    var records = excel.GetRecords<TestStruct>().ToList();

                    // Verify the records are what we expect
                    ClassicAssert.AreEqual(1, records.Count);
                    var record = records[0];
                    ClassicAssert.AreEqual(1, record.ID);
                    ClassicAssert.AreEqual("a name", record.Name);
                }
            }
        }

        [Test]
        public void TrimFieldsTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");

                    // Write the header fields
                    sheet.Cell(1, 1).SetValue("IntColumn");
                    sheet.Cell(1, 2).SetValue("StringColumn");

                    // Write the first record
                    sheet.Cell(2, 1).SetValue(1);
                    sheet.Cell(2, 2).SetValue(" string ");

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    excel.Configuration.WillThrowOnMissingHeader = false;
                    excel.Configuration.TrimFields = true;
                    var records = excel.GetRecords<TestRecord>().ToList();

                    // Verify the records are what we expect
                    ClassicAssert.AreEqual(1, records.Count);
                    var record = records[0];
                    ClassicAssert.AreEqual(1, record.IntColumn);
                    ClassicAssert.AreEqual("string", record.StringColumn);
                }
            }
        }

        [Test]
        public void GetRecordsAsDictionaryTest()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                var now = DateTime.Now;
                var date = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, now.Second, now.Millisecond, DateTimeKind.Unspecified);
                var guid = Guid.NewGuid();
                using (var book = new XLWorkbook()) {
                    var sheet = book.AddWorksheet("Sheet 1");
                    WriteRecords(sheet, guid, date);
                    book.AddWorksheet("Sheet 2");
                    sheet = book.AddWorksheet("Sheet 3");
                    WriteRecords(sheet, guid, date);

                    // Save it to the stream
                    book.SaveAs(stream);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReader(stream)) {
                    ValidateRecordsAsDictionary(excel, guid, date);
                    ClassicAssert.AreEqual(3, excel.TotalSheets);
                    ClassicAssert.IsTrue(excel.ChangeSheet(2));
                    ClassicAssert.IsFalse(excel.ChangeSheet(3));
                    ValidateRecordsAsDictionary(excel, guid, date);
                }
            }
        }

        /// <summary>
        /// Validate the records read from the sheet
        /// </summary>
        /// <param name="excel">Excel reader to read from</param>
        /// <param name="guid">GUID for the test</param>
        /// <param name="date">Date for the test</param>
        private static void ValidateRecordsAsDictionary(
            IExcelReader excel,
            Guid guid,
            DateTime date)
        {
            // Make sure we got two records
            var records = excel.GetRecordsAsDictionary().ToList();
            ClassicAssert.AreEqual(2, records.Count);

            // Verify the records are what we expect
            for (var i = 1; i <= records.Count; i++) {
                var record = records[i - 1];
                ClassicAssert.AreEqual(i.ToString(), record["FirstColumn"]);
                ClassicAssert.AreEqual((i * 2).ToString(), record["IntColumn"]);
                ClassicAssert.AreEqual("string column " + i, record["String Column"]);
                ClassicAssert.AreEqual("converts to test", record["TypeConvertedColumn"]);
                ClassicAssert.AreEqual(guid.ToString(), record["GuidColumn"]);
                ClassicAssert.AreEqual((i == 1).ToString().ToUpperInvariant(), record["BoolColumn"]);
                ClassicAssert.AreEqual((i * 3.0).ToString(), record["DoubleColumn"]);
                ClassicAssert.AreEqual(date.AddDays(i).ToString("o"), record["DateTimeColumn"]);
                ClassicAssert.AreEqual("", record["NullStringColumn"]);
            }
        }

        private struct TestStruct
        {
            public int ID { get; set; }
            public string Name { get; set; }
        }

        private class TestBoolean
        {
            public bool BoolColumn { get; set; }
        }

        private class TestDefaultValues
        {
            public int IntColumn { get; set; }
            public string StringColumn { get; set; }
        }

        private sealed class TestDefaultValuesMap : ExcelClassMap<TestDefaultValues>
        {
            public TestDefaultValuesMap()
            {
                Map(m => m.IntColumn).Default(-1);
                Map(m => m.StringColumn).Default("default string");
            }
        }

        private class TestNullable
        {
            public int? IntColumn { get; set; }
            public string StringColumn { get; set; }
            public Guid? GuidColumn { get; set; }
        }

        private class TestRecord
        {
            public int IntColumn { get; set; }
            public string StringColumn { get; set; }
            public string IgnoredColumn { get; set; }
            public string TypeConvertedColumn { get; set; }
            public int FirstColumn { get; set; }
            public Guid GuidColumn { get; set; }
            public int NoMatchingField { get; set; }
            public bool BoolColumn { get; set; }
            public double DoubleColumn { get; set; }
            public DateTime DateTimeColumn { get; set; }
            public string NullStringColumn { get; set; }
            public string OptionalReadColumn { get; set; }
        }

        private class TestRecordMap : ExcelClassMap<TestRecord>
        {
            protected TestRecordMap()
            {
                Map(m => m.IntColumn).TypeConverter<Int32Converter>();
                Map(m => m.StringColumn).Name("String Column");
                Map(m => m.TypeConvertedColumn).Index(1).TypeConverter<TestTypeConverter>();
                Map(m => m.FirstColumn).Index(0);
                Map(m => m.GuidColumn);
                Map(m => m.BoolColumn);
                Map(m => m.DoubleColumn);
                Map(m => m.DateTimeColumn);
                Map(m => m.NullStringColumn);
                Map(m => m.OptionalReadColumn).OptionalRead();
            }
        }

        private sealed class TestRecordMapMissingField : TestRecordMap
        {
            public TestRecordMapMissingField()
            {
                Map(m => m.NoMatchingField);
            }
        }

        private class TestRecordDuplicateHeaderNames
        {
            public string Column1 { get; set; }
            public string Column2 { get; set; }
            public string Column3 { get; set; }
        }

        private sealed class TestRecordDuplicateHeaderNamesMap : ExcelClassMap<TestRecordDuplicateHeaderNames>
        {
            public TestRecordDuplicateHeaderNamesMap()
            {
                Map(m => m.Column1).Name("Column").NameIndex(0);
                Map(m => m.Column2).Name("Column").NameIndex(1);
                Map(m => m.Column3).Name("Column").NameIndex(2);
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
                throw new NotImplementedException();
            }

            public object ConvertFromExcel(
                TypeConverterOptions options,
                object excelValue)
            {
                return "test";
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
        }

        private class Address
        {
            public string Street { get; set; }
            public string City { get; set; }
            public string State { get; set; }
            public string Zip { get; set; }
        }

        private sealed class PersonMap : ExcelClassMap<Person>
        {
            public PersonMap()
            {
                Map(m => m.FirstName);
                Map(m => m.LastName);
                References<HomeAddressMap>(m => m.HomeAddress);
                References<WorkAddressMap>(m => m.WorkAddress);
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
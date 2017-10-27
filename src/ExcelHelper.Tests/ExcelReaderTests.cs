/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ExcelHelper.Configuration;
using ExcelHelper.TypeConversion;
using Microsoft.VisualStudio.TestTools.UnitTesting;
// ReSharper disable ReturnValueOfPureMethodIsNotUsed
// ReSharper disable UnusedAutoPropertyAccessor.Local
// ReSharper disable ClassNeverInstantiated.Local
// ReSharper disable UnusedMember.Local

// TODO: Use the C1 libraries for unit testing to convert from OpenXML test file data to BIFF8 in memory ...    

namespace ExcelHelper.Tests
{
    [TestClass]
    public class ExcelReaderTests
    {
        private ExcelFactory _factory;

        [TestInitialize]
        public void SetUp()
        {
            _factory = new ExcelFactory();
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    // Check the column and row counts are correct
                    Assert.AreEqual(14, excel.TotalColumns);

                    // Test all number conversions
                    if (!excel.ReadRow()) {
                        throw new ArgumentException();
                    }
                    Assert.AreEqual((sbyte)n, excel.GetColumn<sbyte>(0));
                    Assert.AreEqual((short)n, excel.GetColumn<short>(0));
                    Assert.AreEqual((int)n, excel.GetColumn<int>(0));
                    Assert.AreEqual((long)n, excel.GetColumn<long>(0));
                    Assert.AreEqual((byte)n, excel.GetColumn<byte>(0));
                    Assert.AreEqual((ushort)n, excel.GetColumn<ushort>(0));
                    Assert.AreEqual((uint)n, excel.GetColumn<uint>(0));
                    Assert.AreEqual((ulong)n, excel.GetColumn<ulong>(0));
                    Assert.AreEqual((float)n, excel.GetColumn<float>(0));
                    Assert.AreEqual(n, excel.GetColumn<double>(0));
                    Assert.AreEqual((decimal)n, excel.GetColumn<decimal>(0));
                    Assert.AreEqual(n.ToString(), excel.GetColumn<string>(0));

                    // Test all number conversions with a string cell
                    Assert.AreEqual((sbyte)nsi, excel.GetColumn<sbyte>(1));
                    Assert.AreEqual((short)nsi, excel.GetColumn<short>(1));
                    Assert.AreEqual(nsi, excel.GetColumn<int>(1));
                    Assert.AreEqual(nsi, excel.GetColumn<long>(1));
                    Assert.AreEqual((byte)nsi, excel.GetColumn<byte>(1));
                    Assert.AreEqual((ushort)nsi, excel.GetColumn<ushort>(1));
                    Assert.AreEqual((uint)nsi, excel.GetColumn<uint>(1));
                    Assert.AreEqual((ulong)nsi, excel.GetColumn<ulong>(1));
                    Assert.AreEqual((float)ns, excel.GetColumn<float>(2));
                    Assert.AreEqual(ns, excel.GetColumn<double>(2));
                    Assert.AreEqual((decimal)ns, excel.GetColumn<decimal>(2));
                    Assert.AreEqual(nsi.ToString(), excel.GetColumn<string>(1));
                    Assert.AreEqual(ns.ToString(), excel.GetColumn<string>(2));

                    // Test dates
                    Assert.AreEqual(d, excel.GetColumn<DateTime>(3));
                    Assert.AreEqual(d, excel.GetColumn<DateTime>(4));

                    // Test boolean
                    Assert.AreEqual(true, excel.GetColumn<bool>(5));
                    Assert.AreEqual("True", excel.GetColumn<string>(5));
                    Assert.AreEqual(true, excel.GetColumn<bool>(6));
                    Assert.AreEqual("true", excel.GetColumn<string>(6));
                    Assert.AreEqual(true, excel.GetColumn<bool>(7));
                    Assert.AreEqual("yes", excel.GetColumn<string>(7));

                    // Test character
                    Assert.AreEqual('c', excel.GetColumn<char>(8));
                    Assert.AreEqual("c", excel.GetColumn<string>(8));

                    // Test null
                    Assert.AreEqual("", excel.GetColumn<string>(9));
                    Assert.AreEqual(null, excel.GetColumn<int?>(9));
                    Assert.AreEqual(DateTime.MinValue, excel.GetColumn<DateTime>(9));

                    // Test guid
                    Assert.AreEqual(guid, excel.GetColumn<Guid>(10));
                    Assert.AreEqual(guid.ToString(), excel.GetColumn<string>(10));

                    // Test TimeSpan
                    Assert.AreEqual(ts, excel.GetColumn<TimeSpan>(11));
                    // TODO: This won't work until ExcelDataReader is changed to natively parse TimeSpans
                    //Assert.AreEqual(ts.ToString(), excel.GetColumn<string>(11));
                    Assert.AreEqual(ts, excel.GetColumn<TimeSpan>(12));
                    Assert.AreEqual(ts.ToString(), excel.GetColumn<string>(12));
                    Assert.AreEqual(ts, excel.GetColumn<TimeSpan>(13));
                    Assert.AreEqual(ts.ToString(), excel.GetColumn<string>(13));

                    // Test the third sheet
                    Assert.AreEqual(3, excel.TotalSheets);
                    Assert.IsTrue(excel.ChangeSheet(2));
                    Assert.IsFalse(excel.ChangeSheet(3));
                    if (!excel.ReadRow()) {
                        throw new ArgumentException();
                    }
                    Assert.AreEqual("third sheet", excel.GetColumn<string>(0));
                }
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    // Check the column and row counts are correct
                    Assert.AreEqual(1, excel.TotalColumns);

                    // Verify 66K rows
                    for (var i = 0; i < 66000; i++) {
                        if (!excel.ReadRow()) {
                            throw new ArgumentException();
                        }
                        Assert.AreEqual(i, excel.GetColumn<int>(0));
                    }
                }
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
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

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    excel.Configuration.RegisterClassMap<TestRecordMap>();
                    ValidateRecords(excel, guid, date);
                    Assert.AreEqual(3, excel.TotalSheets);
                    Assert.IsTrue(excel.ChangeSheet(2));
                    Assert.IsFalse(excel.ChangeSheet(3));
                    ValidateRecords(excel, guid, date);
                }
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    excel.Configuration.RegisterClassMap<TestRecordMap>();
                    excel.SkipRows(2);
                    ValidateRecords(excel, guid, date);
                    Assert.AreEqual(3, excel.TotalSheets);
                    Assert.IsTrue(excel.ChangeSheet(2));
                    Assert.IsFalse(excel.ChangeSheet(3));
                    ValidateRecords(excel, guid, date);
                }
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    excel.Configuration.RegisterClassMap<TestRecordMap>();
                    ValidateRecords(excel, guid, date);
                    Assert.AreEqual(excel.SheetName, "Test Sheet 1");
                    Assert.AreEqual(3, excel.TotalSheets);
                    Assert.IsTrue(excel.ChangeSheet(2));
                    Assert.AreEqual(excel.SheetName, "Test Sheet 2");
                    Assert.IsFalse(excel.ChangeSheet(3));
                    ValidateRecords(excel, guid, date);
                }
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    excel.Configuration.RegisterClassMap<TestRecordMap>();
                    ValidateRecords(excel, guid, date, "optional");
                    Assert.AreEqual(3, excel.TotalSheets);
                    Assert.IsTrue(excel.ChangeSheet(2));
                    Assert.IsFalse(excel.ChangeSheet(3));
                    ValidateRecords(excel, guid, date, "optional");
                }
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    excel.Configuration.WillThrowOnMissingHeader = false;
                    excel.Configuration.RegisterClassMap<TestRecordMapMissingField>();
                    ValidateRecords(excel, guid, date);
                    Assert.AreEqual(3, excel.TotalSheets);
                    Assert.IsTrue(excel.ChangeSheet(2));
                    Assert.IsFalse(excel.ChangeSheet(3));
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
        private static void WriteRecords(
            IXLWorksheet sheet,
            Guid guid,
            DateTime date,
            string optionalReadValue = null,
            int firstRow = 0)
        {
            // Write the header fields
            sheet.Cell(firstRow + 1, 1).SetValue("FirstColumn");
            sheet.Cell(firstRow + 1, 2).SetValue("TypeConvertedColumn");
            sheet.Cell(firstRow + 1, 3).SetValue("IntColumn");
            sheet.Cell(firstRow + 1, 4).SetValue("String Column");
            sheet.Cell(firstRow + 1, 5).SetValue("GuidColumn");
            sheet.Cell(firstRow + 1, 6).SetValue("BoolColumn");
            sheet.Cell(firstRow + 1, 7).SetValue("DoubleColumn");
            sheet.Cell(firstRow + 1, 8).SetValue("DateTimeColumn");
            sheet.Cell(firstRow + 1, 9).SetValue("NullStringColumn");
            if (optionalReadValue != null) {
                sheet.Cell(firstRow + 1, 10).SetValue("OptionalReadColumn");
            }

            // Write the first record
            sheet.Cell(firstRow + 2, 1).SetValue(1);
            sheet.Cell(firstRow + 2, 2).SetValue("converts to test");
            sheet.Cell(firstRow + 2, 3).SetValue(1 * 2);
            sheet.Cell(firstRow + 2, 4).SetValue("string column 1");
            sheet.Cell(firstRow + 2, 5).SetValue(guid.ToString());
            sheet.Cell(firstRow + 2, 6).SetValue(true);
            sheet.Cell(firstRow + 2, 7).SetValue(1 * 3.0);
            sheet.Cell(firstRow + 2, 8).SetValue(date.AddDays(1));
            sheet.Cell(firstRow + 2, 9).SetValue((object)null);
            if (optionalReadValue != null) {
                sheet.Cell(firstRow + 2, 10).SetValue(optionalReadValue);
            }

            // Write the second record
            sheet.Cell(firstRow + 3, 1).SetValue(2);
            sheet.Cell(firstRow + 3, 2).SetValue("converts to test");
            sheet.Cell(firstRow + 3, 3).SetValue(2 * 2);
            sheet.Cell(firstRow + 3, 4).SetValue("string column 2");
            sheet.Cell(firstRow + 3, 5).SetValue(guid.ToString());
            sheet.Cell(firstRow + 3, 6).SetValue(false);
            sheet.Cell(firstRow + 3, 7).SetValue(2 * 3.0);
            sheet.Cell(firstRow + 3, 8).SetValue(date.AddDays(2));
            sheet.Cell(firstRow + 3, 9).SetValue((object)null);
            if (optionalReadValue != null) {
                sheet.Cell(firstRow + 3, 10).SetValue(optionalReadValue);
            }

            // Write a blank field outside of the header count. To make sure we only
            // process the columns up to the header count width
            sheet.Cell(2, optionalReadValue == null ? 10 : 11).SetValue("");
        }

        /// <summary>
        /// Validate the records read from the sheet
        /// </summary>
        /// <param name="excel">Excel reader to read from</param>
        /// <param name="guid">GUID for the test</param>
        /// <param name="date">Date for the test</param>
        /// <param name="optionalReadValue">Value to expect in the optional read column</param>
        private static void ValidateRecords(
            IExcelReader excel,
            Guid guid,
            DateTime date,
            string optionalReadValue = null)
        {
            // Make sure we got two records
            var records = excel.GetRecords<TestRecord>().ToList();
            Assert.AreEqual(2, records.Count);

            // Verify the records are what we expect
            for (var i = 1; i <= records.Count; i++) {
                var record = records[i - 1];
                Assert.AreEqual(i, record.FirstColumn);
                Assert.AreEqual(i * 2, record.IntColumn);
                Assert.AreEqual("string column " + i, record.StringColumn);
                Assert.AreEqual("test", record.TypeConvertedColumn);
                Assert.AreEqual(guid, record.GuidColumn);
                Assert.AreEqual(0, record.NoMatchingField);
                Assert.AreEqual(i == 1, record.BoolColumn);
                Assert.AreEqual(i * 3.0, record.DoubleColumn);
                Assert.AreEqual(date.AddDays(i), record.DateTimeColumn);
                Assert.AreEqual("", record.NullStringColumn);
                Assert.AreEqual(optionalReadValue, record.OptionalReadColumn);
            }

            // Validate the mapped columns
            var columns = excel.GetImportedColumns();

            // Make sure we have the column count we expect
            Assert.AreEqual(optionalReadValue == null ? 9 : 10, columns.Count);
            Assert.AreEqual("TestRecord", columns[0].DeclaringType.Name);
            Assert.AreEqual("IntColumn", columns[0].Name);
            Assert.AreEqual("FirstColumn", columns[1].Name);
            Assert.AreEqual("StringColumn", columns[2].Name);
            Assert.AreEqual("TypeConvertedColumn", columns[3].Name);
            Assert.AreEqual("GuidColumn", columns[4].Name);
            Assert.AreEqual("BoolColumn", columns[5].Name);
            Assert.AreEqual("DoubleColumn", columns[6].Name);
            Assert.AreEqual("DateTimeColumn", columns[7].Name);
            Assert.AreEqual("NullStringColumn", columns[8].Name);
            if (optionalReadValue != null) {
                Assert.AreEqual("OptionalReadColumn", columns[9].Name);
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    excel.Configuration.RegisterClassMap<PersonMap>();
                    var records = excel.GetRecords<Person>().ToList();

                    // Make sure we got our record
                    Assert.AreEqual(1, records.Count);
                    var record = records[0];
                    Assert.AreEqual("First Name", record.FirstName);
                    Assert.AreEqual("Last Name", record.LastName);
                    Assert.AreEqual("Home Street", record.HomeAddress.Street);
                    Assert.AreEqual("Home City", record.HomeAddress.City);
                    Assert.AreEqual("Home State", record.HomeAddress.State);
                    Assert.AreEqual("Home Zip", record.HomeAddress.Zip);
                    Assert.AreEqual("Work Street", record.WorkAddress.Street);
                    Assert.AreEqual("Work City", record.WorkAddress.City);
                    Assert.AreEqual("Work State", record.WorkAddress.State);
                    Assert.AreEqual("Work Zip", record.WorkAddress.Zip);
                }
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    try {
                        excel.GetRecords<TestRecord>().ToList();
                        Assert.Fail();
                    } catch (ExcelReaderException ex) {
                        Assert.AreEqual("Field 'IntColumn' does not exist in the Excel file.", ex.Message);
                    }
                }
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    excel.Configuration.RegisterClassMap<TestRecordDuplicateHeaderNamesMap>();
                    var records = excel.GetRecords<TestRecordDuplicateHeaderNames>().ToList();

                    // Make sure we got the correct data
                    Assert.AreEqual(1, records.Count);
                    var record = records[0];
                    Assert.AreEqual("one", record.Column1);
                    Assert.AreEqual("two", record.Column2);
                    Assert.AreEqual("three", record.Column3);
                }
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    excel.Configuration.RegisterClassMap<MultipleNamesClassMap>();
                    var records = excel.GetRecords<MultipleNamesClass>().ToList();

                    // Make sure we got the correct data
                    Assert.AreEqual(2, records.Count);
                    Assert.AreEqual(1, records[0].IntColumn);
                    Assert.AreEqual("one", records[0].StringColumn);
                    Assert.AreEqual(2, records[1].IntColumn);
                    Assert.AreEqual("two", records[1].StringColumn);
                }
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    try {
                        excel.GetRecords<TestRecord>().ToList();
                        Assert.Fail();
                    } catch (ExcelReaderException ex) {
                        Assert.AreEqual("No header record was found.", ex.Message);
                    }
                }
            }
        }

        [TestMethod]
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
                // If we wish to test this we should create some real Excel files with null's in them using the C1 library and save
                // them to disk to use directly rather than building them on the fly here with ClosedXML.
                stream.Position = 0;
                using (var excel = _factory.CreateReader(stream)) {
                    var records = excel.GetRecords<TestNullable>().ToList();

                    // Make sure we got two records
                    Assert.AreEqual(2, records.Count);

                    // Verify the records are what we expect
                    var record = records[0];
                    Assert.AreEqual("string", record.StringColumn);
                    Assert.AreEqual(null, record.IntColumn);
                    Assert.AreEqual(null, record.GuidColumn);

                    record = records[1];
                    Assert.AreEqual("", record.StringColumn);
                    Assert.AreEqual(2, record.IntColumn);
                    Assert.AreEqual(guid, record.GuidColumn);
                }
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    excel.Configuration.WillThrowOnMissingHeader = false;
                    excel.Configuration.IsHeaderCaseSensitive = false;
                    var records = excel.GetRecords<TestRecord>().ToList();

                    // Verify the records are what we expect
                    Assert.AreEqual(1, records.Count);
                    var record = records[0];
                    Assert.AreEqual("string", record.StringColumn);
                    Assert.AreEqual(1, record.IntColumn);
                }
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    excel.Configuration.WillThrowOnMissingHeader = false;
                    excel.Configuration.IgnoreHeaderWhiteSpace = true;
                    var records = excel.GetRecords<TestRecord>().ToList();

                    // Verify the records are what we expect
                    Assert.AreEqual(1, records.Count);
                    var record = records[0];
                    Assert.AreEqual("string", record.StringColumn);
                    Assert.AreEqual(1, record.IntColumn);
                }
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    excel.Configuration.WillThrowOnMissingHeader = false;
                    excel.Configuration.TrimHeaders = true;
                    excel.Configuration.RegisterClassMap<TestRecordMapMissingField>();
                    var records = excel.GetRecords<TestRecord>().ToList();

                    // Verify the records are what we expect
                    Assert.AreEqual(1, records.Count);
                    var record = records[0];
                    Assert.AreEqual(1, record.IntColumn);
                    Assert.AreEqual("string", record.StringColumn);
                }
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    excel.Configuration.RegisterClassMap<TestDefaultValuesMap>();
                    var records = excel.GetRecords<TestDefaultValues>().ToList();

                    // Verify the records are what we expect
                    Assert.AreEqual(2, records.Count);
                    var record = records[0];
                    Assert.AreEqual(-1, record.IntColumn);
                    Assert.AreEqual("some string", record.StringColumn);
                    record = records[1];
                    Assert.AreEqual(1, record.IntColumn);
                    Assert.AreEqual("default string", record.StringColumn);
                }
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    var records = excel.GetRecords<TestBoolean>().ToList();

                    // Verify the records are what we expect
                    Assert.AreEqual(10, records.Count);
                    for (var i = 0; i < 5; i++) {
                        Assert.AreEqual(true, records[i].BoolColumn);
                    }
                    for (var i = 5; i < 10; i++) {
                        Assert.AreEqual(false, records[i].BoolColumn);
                    }
                }
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
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
                    Assert.AreEqual(1, records.Count);
                    var record = records[0];
                    Assert.AreEqual("string", record.StringColumn);
                    Assert.AreEqual(1, record.IntColumn);
                    Assert.AreEqual(true, record.BoolColumn);
                    Assert.AreEqual(date, record.DateTimeColumn);

                    // Check we got the information we need about the parse errors
                    Assert.AreEqual(2, allDetails.Count);
                    var details = allDetails[0];
                    Assert.AreEqual(2, details.Row);
                    Assert.AreEqual(3, details.Column);
                    Assert.AreEqual("BoolColumn", details.FieldName);
                    Assert.AreEqual("bullshit", details.FieldValue);
                    details = allDetails[1];
                    Assert.AreEqual(3, details.Row);
                    Assert.AreEqual(4, details.Column);
                    Assert.AreEqual("DateTimeColumn", details.FieldName);
                    Assert.AreEqual("bullshit", details.FieldValue);

                    // Check the exception details are what we expect
                    Assert.AreEqual(2, exceptions.Count);
                    var message =
                        @"Type: 'ExcelHelper.Tests.ExcelReaderTests+TestRecord'" + "\r\n" +
                        @"Row: '2' (1 based)" + "\r\n" +
                        @"Column: '3' (1 based)" + "\r\n" +
                        @"Field Name: 'BoolColumn'" + "\r\n" +
                        @"Field Value: 'bullshit'" + "\r\n";
                    Assert.AreEqual(message, exceptions[0].Data["ExcelHelper"]);
                    message =
                        @"Type: 'ExcelHelper.Tests.ExcelReaderTests+TestRecord'" + "\r\n" +
                        @"Row: '3' (1 based)" + "\r\n" +
                        @"Column: '4' (1 based)" + "\r\n" +
                        @"Field Name: 'DateTimeColumn'" + "\r\n" +
                        @"Field Value: 'bullshit'" + "\r\n";
                    Assert.AreEqual(message, exceptions[1].Data["ExcelHelper"]);
                }
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    var records = excel.GetRecords<TestStruct>().ToList();

                    // Verify the records are what we expect
                    Assert.AreEqual(1, records.Count);
                    var record = records[0];
                    Assert.AreEqual(1, record.ID);
                    Assert.AreEqual("a name", record.Name);
                }
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    excel.Configuration.WillThrowOnMissingHeader = false;
                    excel.Configuration.TrimFields = true;
                    var records = excel.GetRecords<TestRecord>().ToList();

                    // Verify the records are what we expect
                    Assert.AreEqual(1, records.Count);
                    var record = records[0];
                    Assert.AreEqual(1, record.IntColumn);
                    Assert.AreEqual("string", record.StringColumn);
                }
            }
        }

        [TestMethod]
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
                using (var excel = _factory.CreateReader(stream)) {
                    ValidateRecordsAsDictionary(excel, guid, date);
                    Assert.AreEqual(3, excel.TotalSheets);
                    Assert.IsTrue(excel.ChangeSheet(2));
                    Assert.IsFalse(excel.ChangeSheet(3));
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
            Assert.AreEqual(2, records.Count);

            // Verify the records are what we expect
            for (var i = 1; i <= records.Count; i++) {
                var record = records[i - 1];
                Assert.AreEqual(i.ToString(), record["FirstColumn"]);
                Assert.AreEqual((i * 2).ToString(), record["IntColumn"]);
                Assert.AreEqual("string column " + i, record["String Column"]);
                Assert.AreEqual("converts to test", record["TypeConvertedColumn"]);
                Assert.AreEqual(guid.ToString(), record["GuidColumn"]);
                Assert.AreEqual((i == 1).ToString().ToUpperInvariant(), record["BoolColumn"]);
                Assert.AreEqual((i * 3.0).ToString(), record["DoubleColumn"]);
                Assert.AreEqual(date.AddDays(i).ToString("o"), record["DateTimeColumn"]);
                Assert.AreEqual("", record["NullStringColumn"]);
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
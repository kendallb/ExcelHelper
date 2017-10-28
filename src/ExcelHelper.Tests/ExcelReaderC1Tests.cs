/*
 * Copyright (C) 2004-2013 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 */

#if USE_C1_EXCEL
using System;
using System.Collections.Generic;
using System.IO;
using C1.C1Excel;
using System.Linq;
using ExcelHelper.Configuration;
using ExcelHelper.TypeConversion;
using Microsoft.VisualStudio.TestTools.UnitTesting;
// ReSharper disable ReturnValueOfPureMethodIsNotUsed
// ReSharper disable UnusedAutoPropertyAccessor.Local
// ReSharper disable ClassNeverInstantiated.Local
// ReSharper disable UnusedMember.Local

namespace ExcelHelper.Tests
{
    [TestClass]
    public class ExcelReaderC1Tests
    {
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
                var ts = new TimeSpan(1, 2, 3, 4, 5);
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];
                    sheet[0, 0].Value = n;
                    sheet[0, 1].Value = nsi.ToString();
                    sheet[0, 2].Value = ns.ToString();
                    sheet[0, 3].Value = d;
                    sheet[0, 4].Value = d.ToString();
                    sheet[0, 5].Value = true;
                    sheet[0, 6].Value = "true";
                    sheet[0, 7].Value = "yes";
                    sheet[0, 8].Value = c;
                    sheet[0, 9].Value = null;
                    sheet[0, 10].Value = guid.ToString();
                    sheet[0, 11].Value = ts;
                    sheet[0, 12].Value = ts.ToString();
                    sheet[0, 13].Value = ts.ToString();
                    book.Sheets.Insert(1);
                    sheet = book.Sheets.Insert(2);
                    sheet[0, 0].Value = "third sheet";
                    book.Save(stream, FileFormat.Biff8);
                }

                // Now parse the Excel file as all available types
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
                    // Check the column and row counts are correct
                    Assert.AreEqual(14, excel.TotalColumns);
                    Assert.AreEqual(1, excel.TotalRows);

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
                    Assert.AreEqual(ts.ToString(), excel.GetColumn<string>(11));
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
        public void ReadBiff8Test()
        {
            using (var stream = new MemoryStream()) {
                // Create some test data to parse
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];
                    sheet[0, 0].Value = "one";
                    book.Save(stream, FileFormat.Biff8);
                }

                // Now parse the Excel file as all available types
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
                    // Test all number conversions
                    if (!excel.ReadRow()) {
                        throw new ArgumentException();
                    }
                    Assert.AreEqual("one", excel.GetColumn<string>(0));
                }
            }
        }

        [TestMethod]
        public void GetRecordsMissingFieldsThrowsErrorTest()
        {
            using (var stream = new MemoryStream()) {
                // Create an empty book with only one field
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];
                    sheet[0, 0].Value = "FirstColumn";
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];
                    WriteRecords(sheet, guid, date);
                    book.Sheets.Insert(1);
                    sheet = book.Sheets.Insert(2);
                    WriteRecords(sheet, guid, date);

                    // Save it to the stream
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];
                    WriteRecords(sheet, guid, date, null, 2);
                    book.Sheets.Insert(1);
                    sheet = book.Sheets.Insert(2);
                    WriteRecords(sheet, guid, date);

                    // Save it to the stream
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];
                    sheet.Name = "Test Sheet 1";
                    WriteRecords(sheet, guid, date);
                    book.Sheets.Insert(1);
                    sheet = book.Sheets.Insert(2);
                    sheet.Name = "Test Sheet 2";
                    WriteRecords(sheet, guid, date);

                    // Save it to the stream
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];
                    WriteRecords(sheet, guid, date, "optional");
                    book.Sheets.Insert(1);
                    sheet = book.Sheets.Insert(2);
                    WriteRecords(sheet, guid, date, "optional");

                    // Save it to the stream
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];
                    WriteRecords(sheet, guid, date);
                    book.Sheets.Insert(1);
                    sheet = book.Sheets.Insert(2);
                    WriteRecords(sheet, guid, date);

                    // Save it to the stream
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
            XLSheet sheet,
            Guid guid,
            DateTime date,
            string optionalReadValue = null,
            int firstRow = 0)
        {
            // Write the header fields
            sheet[firstRow, 0].Value = "FirstColumn";
            sheet[firstRow, 1].Value = "TypeConvertedColumn";
            sheet[firstRow, 2].Value = "IntColumn";
            sheet[firstRow, 3].Value = "String Column";
            sheet[firstRow, 4].Value = "GuidColumn";
            sheet[firstRow, 5].Value = "BoolColumn";
            sheet[firstRow, 6].Value = "DoubleColumn";
            sheet[firstRow, 7].Value = "DateTimeColumn";
            sheet[firstRow, 8].Value = "NullStringColumn";
            if (optionalReadValue != null) {
                sheet[firstRow, 9].Value = "OptionalReadColumn";
            }

            // Write the first record
            sheet[firstRow + 1, 0].Value = 1;
            sheet[firstRow + 1, 1].Value = "converts to test";
            sheet[firstRow + 1, 2].Value = 1 * 2;
            sheet[firstRow + 1, 3].Value = "string column 1";
            sheet[firstRow + 1, 4].Value = guid.ToString();
            sheet[firstRow + 1, 5].Value = true;
            sheet[firstRow + 1, 6].Value = 1 * 3.0;
            sheet[firstRow + 1, 7].Value = date.AddDays(1);
            sheet[firstRow + 1, 8].Value = null;
            if (optionalReadValue != null) {
                sheet[firstRow + 1, 9].Value = optionalReadValue;
            }

            // Write the second record
            sheet[firstRow + 2, 0].Value = 2;
            sheet[firstRow + 2, 1].Value = "converts to test";
            sheet[firstRow + 2, 2].Value = 2 * 2;
            sheet[firstRow + 2, 3].Value = "string column 2";
            sheet[firstRow + 2, 4].Value = guid.ToString();
            sheet[firstRow + 2, 5].Value = false;
            sheet[firstRow + 2, 6].Value = 2 * 3.0;
            sheet[firstRow + 2, 7].Value = date.AddDays(2);
            sheet[firstRow + 2, 8].Value = null;
            if (optionalReadValue != null) {
                sheet[firstRow + 2, 9].Value = optionalReadValue;
            }

            // Write a blank field outside of the header count. To make sure we only
            // process the columns up to the header count width
            sheet[1, optionalReadValue == null ? 9 : 10].Value = "";
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
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];

                    // Write the header fields
                    sheet[0, 0].Value = "FirstName";
                    sheet[0, 1].Value = "LastName";
                    sheet[0, 2].Value = "HomeStreet";
                    sheet[0, 3].Value = "HomeCity";
                    sheet[0, 4].Value = "HomeState";
                    sheet[0, 5].Value = "HomeZip";
                    sheet[0, 6].Value = "WorkStreet";
                    sheet[0, 7].Value = "WorkCity";
                    sheet[0, 8].Value = "WorkState";
                    sheet[0, 9].Value = "WorkZip";

                    // Write out a record
                    sheet[1, 0].Value = "First Name";
                    sheet[1, 1].Value = "Last Name";
                    sheet[1, 2].Value = "Home Street";
                    sheet[1, 3].Value = "Home City";
                    sheet[1, 4].Value = "Home State";
                    sheet[1, 5].Value = "Home Zip";
                    sheet[1, 6].Value = "Work Street";
                    sheet[1, 7].Value = "Work City";
                    sheet[1, 8].Value = "Work State";
                    sheet[1, 9].Value = "Work Zip";

                    // Save it to the stream
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];

                    // Write the header fields
                    sheet[0, 0].Value = "FirstColumn";
                    sheet[1, 0].Value = 1;
                    sheet[2, 0].Value = 2;

                    // Save it to the stream
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];

                    // Write the header fields
                    sheet[0, 0].Value = "Column";
                    sheet[0, 1].Value = "Column";
                    sheet[0, 2].Value = "Column";

                    // Write the first record
                    sheet[1, 0].Value = "one";
                    sheet[1, 1].Value = "two";
                    sheet[1, 2].Value = "three";

                    // Save it to the stream
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file 
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];

                    // Write the header fields
                    sheet[0, 0].Value = "int2";
                    sheet[0, 1].Value = "string3";

                    // Write the first record
                    sheet[1, 0].Value = 1;
                    sheet[1, 1].Value = "one";

                    // Write the second record
                    sheet[2, 0].Value = 2;
                    sheet[2, 1].Value = "two";

                    // Save it to the stream
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file 
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                using (var book = new C1XLBook()) {
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file 
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];

                    // Write the header fields
                    sheet[0, 0].Value = "StringColumn";
                    sheet[0, 1].Value = "IntColumn";
                    sheet[0, 2].Value = "GuidColumn";

                    // Write the first record
                    sheet[1, 0].Value = "string";
                    sheet[1, 1].Value = null;
                    sheet[1, 2].Value = null;

                    // Write the second record
                    sheet[2, 0].Value = null;
                    sheet[2, 1].Value = 2;
                    sheet[2, 2].Value = guid.ToString();

                    // Save it to the stream
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];

                    // Write the header fields
                    sheet[0, 0].Value = "stringcolumn";
                    sheet[0, 1].Value = "intcolumn";

                    // Write the first record
                    sheet[1, 0].Value = "string";
                    sheet[1, 1].Value = 1;

                    // Save it to the stream
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];

                    // Write the header fields
                    sheet[0, 0].Value = " String Column ";
                    sheet[0, 1].Value = " Int Column ";

                    // Write the first record
                    sheet[1, 0].Value = "string";
                    sheet[1, 1].Value = 1;

                    // Save it to the stream
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];

                    // Write the header fields
                    sheet[0, 0].Value = " IntColumn ";
                    sheet[0, 1].Value = " String Column ";

                    // Write the first record
                    sheet[1, 0].Value = 1;
                    sheet[1, 1].Value = "string";

                    // Save it to the stream
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];

                    // Write the header fields
                    sheet[0, 0].Value = "IntColumn";
                    sheet[0, 1].Value = "StringColumn";

                    // Write the first record
                    sheet[1, 0].Value = null;
                    sheet[1, 1].Value = "some string";

                    // Write the second record
                    sheet[2, 0].Value = 1;
                    sheet[2, 1].Value = null;

                    // Save it to the stream
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];

                    // Write the header fields
                    sheet[0, 0].Value = "BoolColumn";

                    // Write the test records
                    sheet[1, 0].Value = true;
                    sheet[2, 0].Value = "true";
                    sheet[3, 0].Value = "True";
                    sheet[4, 0].Value = "yes";
                    sheet[5, 0].Value = "y";
                    sheet[6, 0].Value = false;
                    sheet[7, 0].Value = "false";
                    sheet[8, 0].Value = "False";
                    sheet[9, 0].Value = "no";
                    sheet[10, 0].Value = "n";

                    // Save it to the stream
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];

                    // Write the header fields
                    sheet[0, 0].Value = "StringColumn";
                    sheet[0, 1].Value = "IntColumn";
                    sheet[0, 2].Value = "BoolColumn";
                    sheet[0, 3].Value = "DateTimeColumn";

                    // Write the first record
                    sheet[1, 0].Value = "string";
                    sheet[1, 1].Value = 1;
                    sheet[1, 2].Value = "bullshit";
                    sheet[1, 3].Value = "bullshit";

                    // Write the second record
                    sheet[2, 0].Value = "string";
                    sheet[2, 1].Value = 1;
                    sheet[2, 2].Value = true;
                    sheet[2, 3].Value = "bullshit";

                    // Write the third record
                    sheet[3, 0].Value = "string";
                    sheet[3, 1].Value = 1;
                    sheet[3, 2].Value = true;
                    sheet[3, 3].Value = date;

                    // Save it to the stream
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                        @"Type: 'ExcelHelper.Tests.ExcelReaderC1Tests+TestRecord'" + "\r\n" +
                        @"Row: '2' (1 based)" + "\r\n" +
                        @"Column: '3' (1 based)" + "\r\n" +
                        @"Field Name: 'BoolColumn'" + "\r\n" +
                        @"Field Value: 'bullshit'" + "\r\n";
                    Assert.AreEqual(message, exceptions[0].Data["ExcelHelper"]);
                    message =
                        @"Type: 'ExcelHelper.Tests.ExcelReaderC1Tests+TestRecord'" + "\r\n" +
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
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];

                    // Write the header fields
                    sheet[0, 0].Value = "ID";
                    sheet[0, 1].Value = "Name";

                    // Write the first record
                    sheet[1, 0].Value = 1;
                    sheet[1, 1].Value = "a name";

                    // Save it to the stream
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];

                    // Write the header fields
                    sheet[0, 0].Value = "IntColumn";
                    sheet[0, 1].Value = "StringColumn";

                    // Write the first record
                    sheet[1, 0].Value = 1;
                    sheet[1, 1].Value = " string ";

                    // Save it to the stream
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                var date = DateTime.Today;
                var guid = Guid.NewGuid();
                using (var book = new C1XLBook()) {
                    var sheet = book.Sheets[0];
                    WriteRecords(sheet, guid, date);
                    book.Sheets.Insert(1);
                    sheet = book.Sheets.Insert(2);
                    WriteRecords(sheet, guid, date);

                    // Save it to the stream
                    book.Save(stream, FileFormat.OpenXml);
                }

                // Now parse the Excel file
                stream.Position = 0;
                using (var excel = new ExcelReaderC1(stream)) {
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
                Assert.AreEqual(date.AddDays(i).ToOADate().ToString(), record["DateTimeColumn"]);
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
#endif
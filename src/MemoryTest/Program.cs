/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 *
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using ExcelHelper;
using ExcelHelper.Configuration;

namespace MemoryTest
{
    class Program
    {
        private static long _maxMemory;

        public class ExcelRecord
        {
            public string String1 { get; set; }
            public string String2 { get; set; }
            public string String3 { get; set; }
            public string String4 { get; set; }
            public string String5 { get; set; }
            public string String6 { get; set; }
            public string String7 { get; set; }
            public bool Bool1 { get; set; }
            public bool Bool2 { get; set; }
            public bool Bool3 { get; set; }
            public decimal Decimal1 { get; set; }
            public decimal Decimal2 { get; set; }
            public decimal Decimal3 { get; set; }
            public decimal Decimal4 { get; set; }
            public decimal Decimal5 { get; set; }
            public decimal Decimal6 { get; set; }
            public decimal Decimal7 { get; set; }
            public DateTime? DateTime1 { get; set; }
            public DateTime? DateTime2 { get; set; }
            public DateTime? DateTime3 { get; set; }
            public DateTime? DateTime4 { get; set; }
            public DateTime? DateTime5 { get; set; }
            public DateTime? DateTime6 { get; set; }
            public DateTime? DateTime7 { get; set; }
            public int Int1 { get; set; }
            public int Int2 { get; set; }
            public int Int3 { get; set; }
            public int Int4 { get; set; }
            public int Int5 { get; set; }
            public int Int6 { get; set; }
            public int Int7 { get; set; }
            public double Double1 { get; set; }
            public double Double2 { get; set; }
            public double Double3 { get; set; }
            public double Double4 { get; set; }
            public double Double5 { get; set; }
            public double Double6 { get; set; }
            public double Double7 { get; set; }
        }

        public sealed class ClassMap : ExcelClassMap<ExcelRecord>
        {
            public ClassMap()
            {
                Map(m => m.String1);
                Map(m => m.String2);
                Map(m => m.String3);
                Map(m => m.String4);
                Map(m => m.String5);
                Map(m => m.String6);
                Map(m => m.String7);
                Map(m => m.Bool1).BooleanStyleNumeric();
                Map(m => m.Bool2).BooleanStyleYesNo();
                Map(m => m.Bool3).BooleanStyleYesBlank();
                Map(m => m.Decimal1).NumberStyle(NumberStyles.Currency);
                Map(m => m.Decimal2).NumberStyle(NumberStyles.Currency);
                Map(m => m.Decimal3).NumberStyle(NumberStyles.Currency);
                Map(m => m.Decimal4).NumberStyle(NumberStyles.Currency);
                Map(m => m.Decimal5).NumberStyle(NumberStyles.Currency);
                Map(m => m.Decimal6).NumberStyle(NumberStyles.Currency);
                Map(m => m.Decimal7).NumberStyle(NumberStyles.Currency);
                Map(m => m.DateTime1).DateTimeStyle(DateTimeStyles.AssumeLocal);
                Map(m => m.DateTime2).DateTimeStyle(DateTimeStyles.AssumeLocal);
                Map(m => m.DateTime3).DateTimeStyle(DateTimeStyles.AssumeLocal);
                Map(m => m.DateTime4).DateTimeStyle(DateTimeStyles.AssumeLocal);
                Map(m => m.DateTime5).DateTimeStyle(DateTimeStyles.AssumeLocal);
                Map(m => m.DateTime6).DateTimeStyle(DateTimeStyles.AssumeLocal);
                Map(m => m.DateTime7).DateTimeStyle(DateTimeStyles.AssumeLocal);
                Map(m => m.Int1);
                Map(m => m.Int2);
                Map(m => m.Int3);
                Map(m => m.Int4);
                Map(m => m.Int5);
                Map(m => m.Int6);
                Map(m => m.Int7);
                Map(m => m.Double1);
                Map(m => m.Double2);
                Map(m => m.Double3);
                Map(m => m.Double4);
                Map(m => m.Double5);
                Map(m => m.Double6);
                Map(m => m.Double7);
            }
        }

        private static IEnumerable<ExcelRecord> GenerateRecords()
        {
            var now = DateTime.Now;
            for (var i = 0; i < 40000; i++) {
                yield return new ExcelRecord {
                    String1 = "This is a long string " + i,
                    String2 = "This is a long string " + i,
                    String3 = "This is a long string " + i,
                    String4 = "This is a long string " + i,
                    String5 = "This is a long string " + i,
                    String6 = "This is a long string " + i,
                    String7 = "This is a long string " + i,
                    Bool1 = i % 1 == 0,
                    Bool2 = i % 1 != 0,
                    Bool3 = i % 1 == 0,
                    Decimal1 = 1.2345m * i,
                    Decimal2 = 12.345m * i,
                    Decimal3 = 123.45m * i,
                    Decimal4 = 1234.56m * i,
                    Decimal5 = 12345.67m * i,
                    Decimal6 = 123456.78m * i,
                    Decimal7 = 1234578.90m * i,
                    DateTime1 = null,
                    DateTime2 = now.AddSeconds(i),
                    DateTime3 = now.AddMinutes(i),
                    DateTime4 = now.AddHours(i),
                    DateTime5 = now.AddDays(i),
                    DateTime6 = now.AddMonths(i),
                    DateTime7 = now,
                    Int1 = i,
                    Int2 = i + 1,
                    Int3 = i + 2,
                    Int4 = i + 3,
                    Int5 = i + 4,
                    Int6 = i + 5,
                    Int7 = i + 6,
                    Double1 = 1.2345 * i,
                    Double2 = 12.345 * i,
                    Double3 = 123.45 * i,
                    Double4 = 1234.56 * i,
                    Double5 = 12345.67 * i,
                    Double6 = 123456.78 * i,
                    Double7 = 1234578.90 * i,
                };
            }
        }

        private static void Main()
        {
            // Save the time stamp and starting memory
            Console.WriteLine(@"Starting ...");
            var start = DateTime.Now;
            var startMemory = _maxMemory = Process.GetCurrentProcess().WorkingSet64;

            // Profile the excel generation
            var factory = new ExcelFactory();
            using (var ms = new MemoryStream()) {
                using (var excel = factory.CreateWriter(ms)) {
                    excel.WriteRecords(GenerateRecords());

                    // Call close before we dispose of our classes, so we can measure the total memory used
                    excel.Close();
                    _maxMemory = Math.Max(Process.GetCurrentProcess().WorkingSet64, _maxMemory);
                }
            }

            // Write out how long it took and how much memory was used
            var usedMemory = _maxMemory - startMemory;
            var elapsed = DateTime.Now - start;
            Console.WriteLine(@"Done!");
            Console.WriteLine(@"Process took: {0} minutes and {1} seconds", Math.Floor(elapsed.TotalMinutes), elapsed.Seconds);
            Console.WriteLine(@"Process used: {0:F1}MB", usedMemory / 1048576.0);
        }
    }
}

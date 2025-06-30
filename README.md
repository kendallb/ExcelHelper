# ExcelHelper

A library for reading and writing Excel files. Extremely fast, flexible, and easy to use. Supports reading and writing of custom class objects.

This library is based on the awesome library CsvHelper for reading and writing CSV files by Josh Close. We use the awesome
library for reading Excel files in all formats (including the binary BIFF8 and earlier formats as well as the modern OpenXML formats), 
ExcelDataReader. Lastly we use the awesome library for writing native OpenXML based Excel files ClosedXML, which is layered on top of
the Microsoft OpenXML libraries. Since we rely on OpenXML for writing Excel files, this library can only write Excel files in the modern
OpenXML (.xlsx) formats. But it can read them in any format.

This library is re-write of the CsvHelper library to support Excel file reading and writing so it has diverged somewhat from the original 
CsvHelper API in order to support the features we need with better support for Excel file reading and writing. You can find the awesome
CsvHelper library for processing CSV files here:

https://github.com/JoshClose/CsvHelper

You can find the awesome Open Source libraries we rely upon for Excel reading and writing here:

https://github.com/ExcelDataReader/ExcelDataReader

https://github.com/ClosedXML/ClosedXML

## Install

To install ExcelHelper, run the following command in the Package Manager Console

    PM> Install-Package AMain.ExcelHelper

## Documentation

TODO: This needs to be done

## License

Dual licensed

Microsoft Public License (MS-PL)

http://www.opensource.org/licenses/MS-PL

Apache License, Version 2.0

http://opensource.org/licenses/Apache-2.0

Commercial license for C1.Excel can be generated for builds as explained here (you need an active license to do that):

https://developer.mescius.com/componentone/docs/license/online-license/license-user-controls.html

## Contribution

Want to contribute? Great! Here are a few guidelines.

1. If you want to do a feature, post an issue about the feature first. Some features are intentionally left out, some features may already be in the works, or I may have some advice on how I think it should be done. I would feel bad if time was spent on some code that won't be used.
2. If you want to do a bug fix, it might not be a bad idea to post about it too. I've had the same bug fixed by multiple people at the same time before.
3. All code should have a unit test. If you make a feature, there should be significant tests around the feature. If you do a bug fix, there should be a test specific to that bug so it doesn't happen again.
4. Pull requests should have a single commit. If you have multiple commits, squash them into a single commit before requesting a pull.
5. Try and follow the code styling already in place. If you have ReSharper there is a dotsettings file included and things should automatically be formatted for you.

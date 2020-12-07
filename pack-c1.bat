@echo off
rmdir /S /Q NuGet
src\.nuget\NuGet pack src\ExcelHelper.C1 -Symbols -Properties Configuration=Release -OutputDirectory NuGet
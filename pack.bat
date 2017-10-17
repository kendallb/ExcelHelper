@echo off
rmdir /S /Q NuGet
src\.nuget\NuGet pack src\ExcelHelper -Symbols -Properties Configuration=Release -OutputDirectory NuGet
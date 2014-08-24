@echo off

rem Create NuGet package and output to local NuGet path.
for /r %%x in (*.nuspec) do nuget pack "%%x" -o C:\NuGet\

pause
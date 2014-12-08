@echo off

set _nuget="..\packages\NuGet.CommandLine.2.8.3\tools\NuGet.exe"

rem Create NuGet package and output to local NuGet path.
for /r %%x in (*.nuspec) do %_nuget% pack "%%x" -o C:\NuGet\

pause
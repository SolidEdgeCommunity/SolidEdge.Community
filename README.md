SolidEdge.Community Repostiory
================
This is the central repository for the Solid Edge Community.

---

## [SolidEdge.Community NuGet Package](https://www.nuget.org/packages/SolidEdge.Community)
SolidEdge.Community.dll is a single .NET 4.0 assembly that makes automating Solid Edge easy. The core feature of this assembly are [Extension Methods](http://msdn.microsoft.com/library/bb383977.aspx) to the base Solid Edge API. The assembly also has helper classes where needed.

SolidEdge.Community.xml is also provided alongside the .dll and is the primary source of documentation.

As part of the build process, [Interop.SolidEdge.dll](https://www.nuget.org/packages/Interop.SolidEdge) is merged into the assembly using [ILRepack](https://www.nuget.org/packages/ILRepack), thus removing the dependency on [Interop.SolidEdge.dll](https://www.nuget.org/packages/Interop.SolidEdge) and simplifying development\deployment. Note that all documentation from Interop.SolidEdge.dll is also merged.

---

## [SolidEdge.Community.AddIn.Tools NuGet Package](https://www.nuget.org/packages/SolidEdge.Community.AddIn.Tools)

This NuGet package does not contain any assemblies but does depends on the [SolidEdge.Community NuGet Package](https://www.nuget.org/packages/SolidEdge.Community) which will automatically get installed when you install this package.

For example usage, please download and review the AddInDemo project available at the [Samples for Solid Edge on CodePlex](https://solidedgesamples.codeplex.com).

### MSBUILD custom task
This package installs a MSBUILD custom task named SolidEdge.Community.AddIn.Tools.targets into your project. Upon the AfterBuild event, an executable named EmbedResources.exe will scan your assembly for native Win32 resources and embed them into your assembly. The reason this is necessary is because the Solid Edge AddIn API only allows Win32 resources. This requirement does not work well with .NET projects so the executable was created to aide in the process.

### Package Manager Console
This package installs a PowerShell script named SolidEdgeAddIn.psm1. This script adds two console commands that aide in registering\unregistering your addin.  You can control the registration of your addin with Solid Edge directly inside Visual Studio. Navigate to Tools -> NuGet Package Manager -> Package Manager Console. 

* **Register-SolidEdgeAddIn** - Registers the addin for Solid Edge x86 and Solid Edge  x64 on the development machine.
* **Unregister-SolidEdgeAddIn** - Unregisters the addin for Solid Edge x86 and Solid Edge x64 on the development machine.

---

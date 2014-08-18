SolidEdge.Community Repostiory
================
This is the central repository for the [SolidEdge.Community](http://www.nuget.org/packages/SolidEdge.Community) NuGet package.

---

## Overview
The core functionality of this package is the SolidEdge.Community.dll. This assembly extends the [Interop.SolidEdge](http://www.nuget.org/packages/Interop.SolidEdge) NuGet package and enhances the development experience. For examples of how to use the package, please download the latest release of [Samples for Solid Edge](https://solidedgesamples.codeplex.com/) on CodePlex.

---

## Installation via Nuget Package Manager
The [Nuget Package Manager](http://docs.nuget.org/docs/start-here/managing-nuget-packages-using-the-dialog) provides a GUI interface for interacting with NuGet. Note that the steps will vary depending on your version of Visual Studio.

![](https://raw.githubusercontent.com/SolidEdgeCommunity/SolidEdge.Community/master/media/Install.png)

## Installation via Nuget Package Manager Console
The [Nuget Package Manager Console](http://docs.nuget.org/docs/start-here/using-the-package-manager-console) provides a command line style interface for interacting with NuGet. Note that the steps will vary depending on your version of Visual Studio.

![](https://raw.githubusercontent.com/SolidEdgeCommunity/SolidEdge.Community/master/media/InstallCommandLine.png)

You can also install a specific version of the package by appending the version to the command.

![](https://raw.githubusercontent.com/SolidEdgeCommunity/SolidEdge.Community/master/media/InstallCommandLineVersion.png)

---

## Extension Methods
[Extension methods](http://msdn.microsoft.com/library/bb383977.aspx) make working with the Solid Edge API much easier. These are community provided methods that extend the base Solid Edge API.

![](https://raw.githubusercontent.com/SolidEdgeCommunity/SolidEdge.Community/master/media/ExtensionMethods.png)

**C# Syntax**

    // The following statement enables all SolidEdgeCommunity extension methods.
    using SolidEdgeCommunity.Extensions;
    
or
    
    // Extension methods can also be called directly.
    SolidEdgeCommunity.Extensions.ApplicationExtensions.StartCommand(application, SolidEdgeConstants.PartCommandConstants.PartEditCopy);

**Visual Basic Syntax**

    ' The following statement enables all SolidEdgeCommunity extension methods.
    Imports SolidEdgeCommunity.Extensions

or

    ' Extension methods can also be called directly.
    SolidEdgeCommunity.Extensions.ApplicationExtensions.StartCommand(objApplication, SolidEdgeConstants.PartCommandConstants.PartEditCopy)
    
### Intellisense
When extension methods are enabled, they will appear next to the standard APIs. They have **different icons** and are marked **(extension)** in their description.

![](https://raw.githubusercontent.com/SolidEdgeCommunity/SolidEdge.Community/master/media/ExtensionMethodIntellisense.png)

---

## NuGet Package Manager Console Commands
The [SolidEdge.Community](http://www.nuget.org/packages/SolidEdge.Community) package imports a PowerShell script named SolidEdge.Community.psm1 into your project on startup. This script adds several console commands that aide in Solid Edge development.  To access the console inside Visual Studio. Navigate to Tools -> NuGet Package Manager -> Package Manager Console. 

### Register-SolidEdgeAddIn
Registers the addin for Solid Edge x86 and Solid Edge  x64 on the development machine.

### Unregister-SolidEdgeAddIn
Unregisters the addin for Solid Edge x86 and Solid Edge x64 on the development machine.

### Enable-SolidEdgeCommunityBuildTarget
Modifies your project to include a custom [MSBuild target](http://msdn.microsoft.com/library/ms171462.aspx). Currently, the build target will execute EmbedNativeResources.exe against your assembly during the **AfterBuild** event. As the executable name implies, it will embed native Win32 resources into your assembly after building the project. Native Win32 resources are necessary when you want to show custom images in Solid Edge.
    
### Disable-SolidEdgeCommunityBuildTarget
Modifies your project and removes the custom [MSBuild target](http://msdn.microsoft.com/library/ms171462.aspx) added by a previous Enable-SolidEdgeCommunityBuildTarget command.

### Install-SolidEdgeAddInRibbonSchema
Adds a Ribbon.xsd to your project. This XSD contains definitions that validate any user created Ribbon XML. The validation happens real-time in Visual Studio. If you later decide that you do not want the XSD, simply delete it from your project.

### Start-SolidEdge
Starts a new instance of Solid Edge and makes it visible to the user.

### Stop-SolidEdge
Stops a running instance of Solid Edge.
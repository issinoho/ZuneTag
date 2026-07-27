//------------------------------------------------------------------
// Zune Meta Tag Editor
// Assembly Info
//
// <copyright file="AssemblyInfo.cs" company="The Drunken Bakery">
//     Copyright (c) 2009 The Drunken Bakery. All rights reserved.
// </copyright>
//
// Editor to update WMV meta tags for the Zune
// Assembly-level attributes.
//
// Author: IRS
// $Revision: 1.2 $
//------------------------------------------------------------------

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

[assembly: InternalsVisibleTo("ZuneTag.Tests")]

// General Information about an assembly is controlled through the following
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("Zune Tag Editor")]
[assembly: AssemblyDescription("Editor to update WMV meta tags for the Zune")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("The Drunken Bakery")]
[assembly: AssemblyProduct("ZuneTag")]
[assembly: AssemblyCopyright("Copyright © The Drunken Bakery 2009-")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible
// to COM components.  If you need to access a type in this assembly from
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("71b56e16-a7be-4d6a-871d-ab8035f7f1bb")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version
//      Build Number
//      Revision
//
// This is a fixed placeholder for local/dev builds. The GitHub Actions release
// workflow (.github/workflows/build-windows.yml) overwrites both of these with
// the next release's actual version number before building, so a released
// ZuneTag.exe's assembly version always matches its GitHub release tag.
[assembly: AssemblyVersion("1.1.16.0")]
[assembly: AssemblyFileVersion("1.1.16.0")]

Imports System
Imports System.Reflection
Imports System.Resources
Imports System.Runtime.InteropServices
Imports System.Runtime.Versioning

#If NETCOREAPP Then
<Assembly: SupportedOSPlatform("windows")>
#End If

' General Information about an assembly is controlled through the following
' set of attributes. Change these attribute values to modify the information
' associated with an assembly.

' Review the values of the assembly attributes
<Assembly: AssemblyCompany("David Parvin")>

#If DEBUG Then
<Assembly: AssemblyConfiguration("Debug")>
#Else
<Assembly: AssemblyConfiguration("Release")>
#End If

<Assembly: AssemblyCopyright("Copyright © 2025")>
<Assembly: AssemblyDescription("EventLogHelper.Elevator")>
<Assembly: AssemblyProduct("EventLogHelper.Elevator")>
<Assembly: AssemblyTrademark("")>

<Assembly: NeutralResourcesLanguage("en-US", UltimateResourceFallbackLocation.MainAssembly)>

#If NET8_0_OR_GREATER Then
<Assembly: System.Runtime.Versioning.SupportedOSPlatform("windows")>
#End If

#If NET35 Then
<Assembly: AssemblyTitle("Event Log Helper Elevator (.NET Framework 3.5)")>
#End If

#If NET462 Then
<Assembly: AssemblyTitle("Event Log Helper Elevator (.NET Framework 4.6.2)")>
#End If

#If NET472 Then
<Assembly: AssemblyTitle("Event Log Helper Elevator (.NET Framework 4.7.2)")>
#End If

#If NET481 Then
<Assembly: AssemblyTitle("Event Log Helper Elevator (.NET Framework 4.8.1)")>
#End If

#If NET8_0 Then
<Assembly: AssemblyTitle("Event Log Helper Elevator (.NET 8.0)")>
#End If

#If NET9_0 Then
<Assembly: AssemblyTitle("Event Log Helper Elevator (.NET 9.0)")>
#End If

<Assembly: ComVisible(False)>

'The following GUID is for the ID of the typelib if this project is exposed to COM
<Assembly: Guid("6494a13a-d11f-4bb7-b946-29f3f4b48cb2")>

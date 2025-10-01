Imports System.Reflection
Imports System.Resources
Imports System.Runtime.InteropServices
Imports System.Runtime.Versioning

' General Information about an assembly is controlled through the following
' set of attributes. Change these attribute values to modify the information
' associated with an assembly.
<Assembly: AssemblyCompany("David Parvin")>

#If DEBUG Then
<Assembly: AssemblyConfiguration("Debug")>
#Else
<Assembly: AssemblyConfiguration("Release")>
#End If

<Assembly: AssemblyCopyright("Copyright © 2025")>
<Assembly: AssemblyProduct("Event Log Helper")>
<Assembly: AssemblyDescription("EventLogHelper.Sample")>
<Assembly: NeutralResourcesLanguage("en-US", UltimateResourceFallbackLocation.MainAssembly)>

#If NET8_0_OR_GREATER Then
<Assembly: System.Runtime.Versioning.SupportedOSPlatform("windows")>
#End If

#If NET35 Then
<Assembly: AssemblyTitle("Event Log Helper Documentation Generator (.NET Framework 3.5)")>
#End If

#If NET462 Then
<Assembly: AssemblyTitle("Event Log Helper Documentation Generator (.NET Framework 4.6.2)")>
#End If

#If NET472 Then
<Assembly: AssemblyTitle("Event Log Helper Documentation Generator (.NET Framework 4.7.2)")>
#End If

#If NET481 Then
<Assembly: AssemblyTitle("Event Log Helper Documentation Generator (.NET Framework 4.8.1)")>
#End If

#If NET8_0 Then
<Assembly: AssemblyTitle("Event Log Helper Documentation Generator (.NET 8.0)")>
#End If

#If NET9_0 Then
<Assembly: AssemblyTitle("Event Log Helper Documentation Generator (.NET 9.0)")>
#End If

' Setting ComVisible to false makes the types in this assembly not visible
' to COM components.  If you need to access a type in this assembly from
' COM, set the ComVisible attribute to true on that type.
<Assembly: ComVisible(False)>

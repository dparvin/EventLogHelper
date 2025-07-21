Imports System.Reflection
Imports System.Resources
Imports System.Runtime.InteropServices

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
<Assembly: AssemblyDescription("Event Log Helper Unit Tests")>
<Assembly: AssemblyProduct("Event Log Helper")>
<Assembly: AssemblyTitle("EventLogHelper.Tests")>
<Assembly: NeutralResourcesLanguage("en-US", UltimateResourceFallbackLocation.MainAssembly)>

#If NET8_0_OR_GREATER Then
<Assembly: Runtime.Versioning.SupportedOSPlatform("windows")>
#End If

' Setting ComVisible to false makes the types in this assembly not visible
' to COM components.  If you need to access a type in this assembly from
' COM, set the ComVisible attribute to true on that type.
<Assembly: ComVisible(False)>

Imports System.IO
Imports Microsoft.Win32
Imports Xunit.Abstractions
Imports EventLogHelper.Interfaces

#If NET35 Then
<Diagnostics.ExcludeFromCodeCoverage>
Public Class TestRegistryReader
#Else
<CodeAnalysis.ExcludeFromCodeCoverage>
Public Class TestRegistryReader
#End If

    Implements IRegistryReader

    Private ReadOnly _registryData As Dictionary(Of String, String)

#If NET35 Then
    ''' <summary>
    ''' Initializes a new instance of the <see cref="TestRegistryReader"/> class.
    ''' </summary>
    Public Sub New()
#Else
    ''' <summary>
    ''' Initializes a new instance of the <see cref="TestRegistryReader"/> class.
    ''' </summary>
    Public Sub New(
            ByRef output As ITestOutputHelper)

        OutputHelper = output
#End If

        _registryData = New Dictionary(Of String, String) From {
            {"HKEY_LOCAL_MACHINE\SOFTWARE\MyApp\Setting1", "Value1"},
            {"HKEY_LOCAL_MACHINE\SOFTWARE\MyApp\Setting2", "Value2"}
        }

    End Sub

#If NET35 Then
#Else
    ''' <summary>
    ''' Gets or sets the output helper.
    ''' </summary>
    ''' <value>
    ''' The output helper.
    ''' </value>
    Private Property OutputHelper As ITestOutputHelper
#End If

    ''' <summary>
    ''' Gets the value.
    ''' </summary>
    ''' <param name="key">The key.</param>
    ''' <returns></returns>
    Friend Function GetValue(key As String) As String

        Dim results As String = String.Empty
        Return If(_registryData.TryGetValue(key, results), results, Nothing)

    End Function

    ''' <summary>
    ''' Sets the value.
    ''' </summary>
    ''' <param name="key">The key.</param>
    ''' <param name="value">The value.</param>
    Friend Sub SetValue(key As String, value As String)

        If KeyExists(key) Then
            _registryData(key) = value
        Else
            _registryData.Add(key, value)
        End If

    End Sub

    ''' <summary>
    ''' Keys the exists.
    ''' </summary>
    ''' <param name="key">The key.</param>
    ''' <returns></returns>
    Private Function KeyExists(key As String) As Boolean

        Output($"Checking if key exists: {key} and it returned {_registryData.ContainsKey(key)}")

        Return _registryData.ContainsKey(key)

    End Function

    ''' <summary>
    ''' Gets or sets the default event log source result.
    ''' </summary>
    ''' <value>
    ''' The default event log source result.
    ''' </value>
    Public Property DefaultEventLogSourceResult As String = "Application"

    ''' <summary>
    ''' Determines whether a specific subkey exists within the given registry hive on the specified machine.
    ''' </summary>
    ''' <param name="hive">The root <see cref="T:Microsoft.Win32.RegistryHive" /> (e.g., <c>RegistryHive.LocalMachine</c>, <c>RegistryHive.CurrentUser</c>) where the lookup begins.</param>
    ''' <param name="machineName">The name of the target machine. Use <c>"."</c> for the local machine.</param>
    ''' <param name="path">The registry path under the specified hive to check for existence (e.g., <c>"SYSTEM\CurrentControlSet\Services\EventLog\Application"</c>).</param>
    ''' <returns>
    ''' <c>true</c> if the registry subkey exists; otherwise, <c>false</c>.
    ''' </returns>
    Public Function SubKeyExists(
            ByVal hive As RegistryHive,
            ByVal machineName As String,
            ByVal path As String) As Boolean Implements IRegistryReader.SubKeyExists

        ' Simulate checking for a subkey in the test registry data
        Dim fullPath As String = $"{hive}\{path}"
        Output($"Checking if subkey exists at hive: {hive}, machineName: '{machineName}', path: '{path}' returning {_registryData.ContainsKey(fullPath)}")

        Return _registryData.ContainsKey(fullPath)

    End Function

    ''' <summary>
    ''' Determines whether the current user has permission to access a specific registry location, with an option to check for write access.
    ''' </summary>
    ''' <param name="hive">The root <see cref="T:Microsoft.Win32.RegistryHive" /> to open (e.g., <c>RegistryHive.LocalMachine</c>).</param>
    ''' <param name="machineName">The name of the machine to access. Use <c>"."</c> for the local machine.</param>
    ''' <param name="registryPath">The path of the registry key relative to the specified hive (e.g., <c>"SYSTEM\CurrentControlSet\Services\EventLog\Application"</c>).</param>
    ''' <param name="writeAccess">If <c>true</c>, checks for write permission; otherwise, checks for read access.</param>
    ''' <returns>
    ''' <c>true</c> if access is permitted; otherwise, <c>false</c>.
    ''' </returns>
    Public Function HasRegistryAccess(
            hive As RegistryHive,
            machineName As String,
            registryPath As String,
            writeAccess As Boolean) As Boolean Implements IRegistryReader.HasRegistryAccess

        ' Simulate checking for a subkey in the test registry data
        Dim fullPath As String = $"{hive}\{registryPath}"
        Output($"Checking that the user has Registry Access at hive: {hive}, machineName: '{machineName}', registryPath: '{registryPath}', writeAccess: {writeAccess} returning {_registryData.ContainsKey(fullPath)}")

        Return _registryData.ContainsKey(fullPath)

    End Function
    Public Function GetDefaultEventLogSource(
            ByVal logName As String,
            ByVal machineName As String) As String Implements IRegistryReader.GetDefaultEventLogSource

        Output($"Getting default source for event log: log name: '{logName}' on machine: '{machineName}' and returns '{DefaultEventLogSourceResult}'")

        Return DefaultEventLogSourceResult

    End Function

    ''' <summary>
    ''' Outputs the specified message.
    ''' </summary>
    ''' <param name="message">The message.</param>
#If NET35 Then
    Private Sub Output(ByVal message As String)

        Console.WriteLine(message)

    End Sub
#Else
    Private Sub Output(ByVal message As String)

        OutputHelper.WriteLine(message)

    End Sub

#End If

End Class

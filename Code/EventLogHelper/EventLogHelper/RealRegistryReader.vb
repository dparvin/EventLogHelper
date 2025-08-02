Imports System.IO
Imports System.Security
Imports Microsoft.Win32

''' <summary>
''' Provides a way to read and validate access to the Windows Registry.
''' </summary>
''' <seealso cref="EventLogHelper.IRegistryReader" />
#If NET35 Then
<Diagnostics.ExcludeFromCodeCoverage>
Friend Class RealRegistryReader
#Else
<CodeAnalysis.ExcludeFromCodeCoverage>
Friend Class RealRegistryReader
#End If

    Implements IRegistryReader

    ''' <summary>
    ''' Initializes a new instance of the <see cref="RealRegistryReader"/> class.
    ''' </summary>
    Friend Sub New()
    End Sub

    ''' <summary>
    ''' Determines whether a specific subkey exists within the given registry hive on the specified machine.
    ''' </summary>
    ''' <param name="hive">The root <see cref="RegistryHive" /> (e.g., <c>RegistryHive.LocalMachine</c>, <c>RegistryHive.CurrentUser</c>) where the lookup begins.</param>
    ''' <param name="machineName">The name of the target machine. Use <c>"."</c> for the local machine.</param>
    ''' <param name="registryPath"></param>
    ''' <returns>
    '''   <c>true</c> if the registry subkey exists; otherwise, <c>false</c>.
    ''' </returns>
    Friend Function SubKeyExists(
            ByVal hive As RegistryHive,
            ByVal machineName As String,
            ByVal registryPath As String) As Boolean Implements IRegistryReader.SubKeyExists

        Try
            Dim baseKey As RegistryKey
            If machineName = "." OrElse String.Equals(machineName, Environment.MachineName, StringComparison.OrdinalIgnoreCase) Then
#If NET35 Then
                baseKey = OpenBaseKey(hive)
#Else
                baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Default)
#End If
            Else
                baseKey = RegistryKey.OpenRemoteBaseKey(hive, machineName)
            End If

            Using eventLogKey As RegistryKey = baseKey.OpenSubKey(registryPath)
                Return eventLogKey IsNot Nothing
            End Using
        Catch ex As Exception
            ' You may want to log or handle this differently
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Determines whether the current user has permission to access a specific registry location, with an option to check for write access.
    ''' </summary>
    ''' <param name="hive">The root <see cref="RegistryHive" /> to open (e.g., <c>RegistryHive.LocalMachine</c>).</param>
    ''' <param name="machineName">The name of the machine to access. Use <c>"."</c> for the local machine.</param>
    ''' <param name="registryPath">The path of the registry key relative to the specified hive (e.g., <c>"SYSTEM\CurrentControlSet\Services\EventLog\Application"</c>).</param>
    ''' <param name="writeAccess">If <c>true</c>, checks for write permission; otherwise, checks for read access.</param>
    ''' <returns>
    ''' <c>true</c> if access is permitted; otherwise, <c>false</c>.
    ''' </returns>
    Friend Function HasRegistryAccess(
            ByVal hive As RegistryHive,
            ByVal machineName As String,
            ByVal registryPath As String,
            ByVal writeAccess As Boolean) As Boolean Implements IRegistryReader.HasRegistryAccess

        Try
            Dim baseKey As RegistryKey
            If machineName = "." OrElse String.Equals(machineName, Environment.MachineName, StringComparison.OrdinalIgnoreCase) Then
#If NET35 Then
                baseKey = OpenBaseKey(hive)
#Else
                baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Default)
#End If
            Else
                baseKey = RegistryKey.OpenRemoteBaseKey(hive, machineName)
            End If

            ' Attempt to open the sub-key with desired access
            Using subKey As RegistryKey = baseKey.OpenSubKey(registryPath, writable:=writeAccess)
                Return subKey IsNot Nothing
            End Using

        Catch ex As UnauthorizedAccessException
            Return False
        Catch ex As Security.SecurityException
            Return False
        Catch ex As Exception
            ' Handle other errors like key not found, etc.
            Return False
        End Try

    End Function

#If NET35 Then
    ''' <summary>
    ''' Opens the base key.
    ''' </summary>
    ''' <param name="hive">The hive.</param>
    ''' <returns></returns>
    ''' <exception cref="System.NotSupportedException">Unsupported hive: {hive}</exception>
    Private Shared Function OpenBaseKey(
            ByVal hive As RegistryHive) As RegistryKey

        Dim baseKey As RegistryKey

        Select Case hive
            Case RegistryHive.ClassesRoot
                baseKey = Registry.ClassesRoot
            Case RegistryHive.CurrentUser
                baseKey = Registry.CurrentUser
            Case RegistryHive.LocalMachine
                baseKey = Registry.LocalMachine
            Case RegistryHive.Users
                baseKey = Registry.Users
            Case RegistryHive.CurrentConfig
                baseKey = Registry.CurrentConfig
            Case Else
                Throw New NotSupportedException($"Unsupported hive: {hive}")
        End Select

        Return baseKey

    End Function
#End If

    ''' <summary>
    ''' Attempts to determine a default source name for the given event log by inspecting the registry.
    ''' </summary>
    ''' <param name="logName">The name of the event log (e.g., "Application", "System").</param>
    ''' <param name="machineName">The target machine name, or "." for the local machine.</param>
    ''' <returns>The name of a registered event source, or an empty string if none are found.</returns>
    Public Function GetDefaultEventLogSource(
            ByVal logName As String,
            ByVal machineName As String) As String Implements IRegistryReader.GetDefaultEventLogSource

        If String.IsNullOrEmpty(logName) Then Return String.Empty

        Try
            Dim baseKey As RegistryKey =
                If(machineName = "."c,
                   Registry.LocalMachine,
                   RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, machineName))

            Using reg As RegistryKey = baseKey.OpenSubKey($"SYSTEM\CurrentControlSet\Services\EventLog\{logName}", False)
                If reg IsNot Nothing Then
                    Dim sources As String() = reg.GetSubKeyNames()
                    If sources.Length > 0 Then
                        If sources.Contains(logName) Then Return logName
                        Return sources(0)
                    End If
                End If
            End Using
        Catch ex As SecurityException
            ' Handle or log if needed
        Catch ex As UnauthorizedAccessException
        Catch ex As IOException
        End Try

        Return String.Empty

    End Function

End Class

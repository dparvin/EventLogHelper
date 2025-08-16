Imports Microsoft.Win32

''' <summary>
''' Provides an abstraction layer for reading from and validating access to the Windows Registry.
''' </summary>
''' <remarks>
''' <para>
''' The <see cref="IRegistryReader"/> interface decouples registry operations from 
''' direct calls to <see cref="Microsoft.Win32.Registry"/>. This enables safer, more 
''' testable code by allowing registry access to be mocked or replaced in unit tests 
''' without requiring actual machine registry modifications.
''' </para>
''' 
''' <para>
''' Implementations typically wrap <see cref="Microsoft.Win32.RegistryKey"/> APIs, 
''' but custom implementations may return simulated results for testing security 
''' scenarios (e.g., missing permissions, missing keys).
''' </para>
''' </remarks>
Public Interface IRegistryReader

    ''' <summary>
    ''' Checks whether a specific registry subkey exists under the given hive and path.
    ''' </summary>
    ''' <param name="hive">
    ''' The root <see cref="RegistryHive"/> (e.g., <c>RegistryHive.LocalMachine</c>, 
    ''' <c>RegistryHive.CurrentUser</c>) where the lookup begins.
    ''' </param>
    ''' <param name="machineName">
    ''' The target machine name. Use <c>"."</c> to indicate the local machine.
    ''' </param>
    ''' <param name="path">
    ''' The registry subkey path relative to the hive 
    ''' (e.g., <c>"SYSTEM\CurrentControlSet\Services\EventLog\Application"</c>).
    ''' </param>
    ''' <returns><c>True</c> if the subkey exists; otherwise, <c>False</c>.</returns>
    ''' <remarks>
    ''' Useful for checking if required event log configuration keys exist before attempting 
    ''' to create or write to an event source.
    ''' </remarks>
    Function SubKeyExists(
        ByVal hive As RegistryHive,
        ByVal machineName As String,
        ByVal path As String) As Boolean

    ''' <summary>
    ''' Verifies whether the current user has sufficient permissions to access a registry key.
    ''' </summary>
    ''' <param name="hive">
    ''' The root <see cref="RegistryHive"/> (e.g., <c>RegistryHive.LocalMachine</c>).
    ''' </param>
    ''' <param name="machineName">
    ''' The target machine name. Use <c>"."</c> for the local machine.
    ''' </param>
    ''' <param name="registryPath">
    ''' The registry key path relative to the specified hive.
    ''' </param>
    ''' <param name="writeAccess">
    ''' If <c>True</c>, checks whether write access is allowed; if <c>False</c>, checks for read access.
    ''' </param>
    ''' <returns><c>True</c> if the user has the requested access; otherwise, <c>False</c>.</returns>
    ''' <remarks>
    ''' This method is critical when creating event sources, since writing registry entries 
    ''' in <c>SYSTEM\CurrentControlSet\Services\EventLog</c> requires administrative privileges.
    ''' </remarks>
    Function HasRegistryAccess(
        ByVal hive As RegistryHive,
        ByVal machineName As String,
        ByVal registryPath As String,
        ByVal writeAccess As Boolean) As Boolean

    ''' <summary>
    ''' Attempts to determine a default event source for the specified event log.
    ''' </summary>
    ''' <param name="logName">
    ''' The name of the event log (e.g., "Application", "System", or a custom log).
    ''' </param>
    ''' <param name="machineName">
    ''' The target machine name. Use <c>"."</c> for the local machine.
    ''' </param>
    ''' <returns>
    ''' The name of a registered event source, or an empty string if none are found.
    ''' </returns>
    ''' <remarks>
    ''' This can be used as a fallback when an explicit source name has not been provided.
    ''' </remarks>
    Function GetDefaultEventLogSource(
        ByVal logName As String,
        ByVal machineName As String) As String

End Interface

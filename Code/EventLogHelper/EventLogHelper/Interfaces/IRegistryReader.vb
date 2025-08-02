Imports Microsoft.Win32

''' <summary>
''' Provides an abstraction for reading and validating access to the Windows Registry.
''' This interface allows for decoupled implementations, enabling unit testing and mocking of registry access.
''' </summary>
Public Interface IRegistryReader

    ''' <summary>
    ''' Determines whether a specific subkey exists within the given registry hive on the specified machine.
    ''' </summary>
    ''' <param name="hive">
    ''' The root <see cref="RegistryHive" /> (e.g., <c>RegistryHive.LocalMachine</c>, <c>RegistryHive.CurrentUser</c>) where the lookup begins.
    ''' </param>
    ''' <param name="machineName">
    ''' The name of the target machine. Use <c>"."</c> for the local machine.
    ''' </param>
    ''' <param name="path">
    ''' The registry path under the specified hive to check for existence (e.g., <c>"SYSTEM\CurrentControlSet\Services\EventLog\Application"</c>).
    ''' </param>
    ''' <returns>
    ''' <c>true</c> if the registry subkey exists; otherwise, <c>false</c>.
    ''' </returns>
    Function SubKeyExists(
        ByVal hive As RegistryHive,
        ByVal machineName As String,
        ByVal path As String) As Boolean

    ''' <summary>
    ''' Determines whether the current user has permission to access a specific registry location, with an option to check for write access.
    ''' </summary>
    ''' <param name="hive">
    ''' The root <see cref="RegistryHive" /> to open (e.g., <c>RegistryHive.LocalMachine</c>).
    ''' </param>
    ''' <param name="machineName">
    ''' The name of the machine to access. Use <c>"."</c> for the local machine.
    ''' </param>
    ''' <param name="registryPath">
    ''' The path of the registry key relative to the specified hive (e.g., <c>"SYSTEM\CurrentControlSet\Services\EventLog\Application"</c>).
    ''' </param>
    ''' <param name="writeAccess">
    ''' If <c>true</c>, checks for write permission; otherwise, checks for read access.
    ''' </param>
    ''' <returns>
    ''' <c>true</c> if access is permitted; otherwise, <c>false</c>.
    ''' </returns>
    Function HasRegistryAccess(
        ByVal hive As RegistryHive,
        ByVal machineName As String,
        ByVal registryPath As String,
        ByVal writeAccess As Boolean) As Boolean

    ''' <summary>
    ''' Attempts to determine a default source name for the given event log by inspecting the registry.
    ''' </summary>
    ''' <param name="logName">The name of the event log (e.g., "Application", "System").</param>
    ''' <param name="machineName">The target machine name, or "." for the local machine.</param>
    ''' <returns>The name of a registered event source, or an empty string if none are found.</returns>
    Function GetDefaultEventLogSource(
        ByVal logName As String,
        ByVal machineName As String) As String

End Interface

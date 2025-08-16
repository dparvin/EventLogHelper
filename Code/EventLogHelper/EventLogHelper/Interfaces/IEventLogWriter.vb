''' <summary>
''' Defines the contract for writing to and managing Windows Event Logs.
''' </summary>
''' <remarks>
''' <para>
''' The <see cref="IEventLogWriter"/> interface abstracts all direct interactions 
''' with <see cref="System.Diagnostics.EventLog"/>. This allows event log operations 
''' (such as creating sources, checking for existence, and writing entries) 
''' to be performed in a testable and mockable way.
''' </para>
''' 
''' <para>
''' The default implementation wraps the .NET Framework <see cref="EventLog"/> API, 
''' but custom implementations can be supplied for unit testing, redirection 
''' (e.g., writing to a file), or specialized logging behavior.
''' </para>
''' </remarks>
Public Interface IEventLogWriter

    ''' <summary>
    ''' Determines whether the specified event log exists on the given machine.
    ''' </summary>
    ''' <param name="logName">The name of the event log to check (e.g., "Application", "System", or a custom log).</param>
    ''' <param name="machineName">The name of the target machine (use "." for the local machine).</param>
    ''' <returns><c>True</c> if the log exists; otherwise, <c>False</c>.</returns>
    ''' <remarks>
    ''' Wraps <see cref="EventLog.Exists(String)"/> to enable testable checks for log existence.
    ''' </remarks>
    Function Exists(
            logName As String,
            machineName As String) As Boolean

    ''' <summary>
    ''' Determines whether the specified event source is registered on the given machine.
    ''' </summary>
    ''' <param name="sourceName">The event source name to check.</param>
    ''' <param name="machineName">The name of the target machine (use "." for the local machine).</param>
    ''' <returns><c>True</c> if the source exists; otherwise, <c>False</c>.</returns>
    ''' <remarks>
    ''' Wraps <see cref="EventLog.SourceExists(String, String)"/> to allow testable verification of event sources.
    ''' </remarks>
    Function SourceExists(
            sourceName As String,
            machineName As String) As Boolean

    ''' <summary>
    ''' Retrieves or initializes an <see cref="EventLog"/> instance with the specified configuration.
    ''' </summary>
    ''' <param name="machineName">The name of the machine (use "." for the local machine).</param>
    ''' <param name="logName">The name of the log (e.g., "Application" or a custom log).</param>
    ''' <param name="sourceName">The name of the event source to use or create.</param>
    ''' <param name="maxKilobytes">Maximum size of the log in kilobytes.</param>
    ''' <param name="retentionDays">Number of days to retain log entries before overwriting.</param>
    ''' <param name="writeInitEntry">If <c>True</c>, an initialization entry is written upon setup.</param>
    ''' <returns>An <see cref="EventLog"/> instance configured with the specified parameters.</returns>
    Function GetLog(
            machineName As String,
            logName As String,
            sourceName As String,
            maxKilobytes As Integer,
            retentionDays As Integer,
            writeInitEntry As Boolean) As EventLog


    ''' <summary>
    ''' Creates a new event source for the specified log on the given machine, if it does not already exist.
    ''' </summary>
    ''' <param name="sourceName">The name of the event source.</param>
    ''' <param name="logName">The log with which to associate the source (e.g., "Application").</param>
    ''' <param name="machineName">The name of the machine (use "." for the local machine).</param>
    ''' <remarks>
    ''' Wraps <see cref="EventLog.CreateEventSource(String, String, String)"/>. 
    ''' If the source already exists, no action is taken.
    ''' </remarks>
    Sub CreateEventSource(
            sourceName As String,
            logName As String,
            machineName As String)

    ''' <summary>
    ''' Writes an entry to the specified event log using explicit log and source details.
    ''' </summary>
    ''' <param name="machineName">The name of the machine where the log entry will be written.</param>
    ''' <param name="logName">The name of the log to write to.</param>
    ''' <param name="sourceName">The event source name (must already be registered with the log).</param>
    ''' <param name="message">The message text to log. If too long, it will be truncated.</param>
    ''' <param name="eventType">The event type (Information, Warning, Error, etc.).</param>
    ''' <param name="eventId">Application-defined identifier for the event.</param>
    ''' <param name="category">Application-defined category number for the event.</param>
    ''' <param name="rawData">Optional binary data associated with the entry (can be <c>Nothing</c>).</param>
    ''' <param name="maxKilobytes">Maximum log size in kilobytes.</param>
    ''' <param name="retentionDays">Retention policy in days.</param>
    ''' <param name="writeInitEntry">If <c>True</c>, an initialization entry is written if the log is first created.</param>
    Sub WriteEntry(
            machineName As String,
            logName As String,
            sourceName As String,
            message As String,
            eventType As EventLogEntryType,
            eventId As Integer,
            category As Short,
            rawData As Byte(),
            maxKilobytes As Integer,
            retentionDays As Integer,
            writeInitEntry As Boolean)

    ''' <summary>
    ''' Writes an entry directly to an existing <see cref="EventLog"/> instance.
    ''' </summary>
    ''' <param name="log">The target <see cref="EventLog"/> instance.</param>
    ''' <param name="message">The message to log.</param>
    ''' <param name="eventType">The event type.</param>
    ''' <param name="eventID">The event identifier.</param>
    ''' <param name="category">The category number.</param>
    ''' <param name="rawData">Optional binary data associated with the entry.</param>
    Sub WriteEntry(
        ByRef log As EventLog,
        ByVal message As String,
        ByVal eventType As EventLogEntryType,
        ByVal eventID As Integer,
        ByVal category As Short,
        ByRef rawData As Byte())

End Interface

''' <summary>
''' Interface for writing to the event log.
''' </summary>
Public Interface IEventLogWriter

    ''' <summary>
    ''' Determines whether the specified event log exists on the local computer.
    ''' </summary>
    ''' <param name="logName">The name of the event log to check (e.g., "Application", "System", or a custom log name).</param>
    ''' <param name="machineName">Name of the machine.</param>
    ''' <returns>
    '''   <c>True</c> if the specified event log exists; otherwise, <c>False</c>.
    ''' </returns><remarks>
    ''' This method wraps <see cref="EventLog.Exists(String)" /> to allow the check to be abstracted
    ''' and testable via <see cref="IEventLogWriter" />.
    ''' </remarks>
    Function Exists(
            logName As String,
            machineName As String) As Boolean

    ''' <summary>
    ''' Determines whether the specified event source is registered with the given event log on the local computer.
    ''' </summary>
    ''' <param name="sourceName">The name of the event source to check.</param>
    ''' <param name="machineName">Name of the machine.</param>
    ''' <returns>
    '''   <c>True</c> if the specified event source exists and is registered with the given log; otherwise, <c>False</c>.
    ''' </returns><remarks>
    ''' This method wraps <see cref="EventLog.SourceExists(String, String)" /> and is used to allow
    ''' testable verification of the event source.
    ''' </remarks>
    Function SourceExists(
            sourceName As String,
            machineName As String) As Boolean

    ''' <summary>
    ''' Gets the log.
    ''' </summary>
    ''' <param name="machineName">Name of the machine.</param>
    ''' <param name="logName">Name of the log.</param>
    ''' <param name="sourceName">Name of the source.</param>
    ''' <param name="maxKilobytes">The maximum kilobytes.</param>
    ''' <param name="retentionDays">The retention days.</param>
    ''' <param name="writeInitEntry">if set to <c>true</c> [write initialize entry].</param>
    ''' <returns></returns>
    Function GetLog(
            machineName As String,
            logName As String,
            sourceName As String,
            maxKilobytes As Integer,
            retentionDays As Integer,
            writeInitEntry As Boolean) As EventLog


    ''' <summary>
    ''' Creates a new event source and associates it with the specified event log, if it does not already exist.
    ''' </summary>
    ''' <param name="sourceName">The name of the event source to create.</param>
    ''' <param name="logName">The name of the event log to associate the source with (e.g., "Application").</param>
    ''' <param name="machineName">Name of the machine.</param><remarks>
    ''' This method wraps <see cref="EventLog.CreateEventSource(String, String, String)" /> to facilitate testable
    ''' logging behavior. If the source already exists, no action is taken.
    ''' </remarks>
    Sub CreateEventSource(
            sourceName As String,
            logName As String,
            machineName As String)

    ''' <summary>
    ''' Writes an entry to the specified event log using the given parameters.
    ''' </summary>
    ''' <param name="machineName">The name of the machine where this is getting logged to.</param>
    ''' <param name="logName">The name of the event log to write to (e.g., "Application" or "CustomLog").</param>
    ''' <param name="sourceName">The source of the log entry. This must be registered with the specified log.</param>
    ''' <param name="message">The message text to log. If too long, it will be truncated to the maximum allowed length.</param>
    ''' <param name="eventType">The type of event to log (e.g., <see cref="EventLogEntryType.Information" />, <see cref="EventLogEntryType.Error" />).</param>
    ''' <param name="eventId">The application-defined identifier for the event.</param>
    ''' <param name="category">An application-defined category number for the event.</param>
    ''' <param name="rawData">Optional raw binary data to associate with the entry. Can be <c>Nothing</c> if not used.</param><remarks>
    ''' This method creates a new <see cref="EventLog" /> instance targeting the specified log and writes the entry
    ''' with the given parameters. The source must be associated with the log prior to calling this method.
    ''' </remarks>
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
    ''' Writes the entry.
    ''' </summary>
    ''' <param name="log">The log.</param>
    ''' <param name="message">The message.</param>
    ''' <param name="eventType">The type.</param>
    ''' <param name="eventID">The event identifier.</param>
    ''' <param name="category">The category.</param>
    ''' <param name="rawData">The raw data.</param>
    Sub WriteEntry(
        ByRef log As EventLog,
        ByVal message As String,
        ByVal eventType As EventLogEntryType,
        ByVal eventID As Integer,
        ByVal category As Short,
        ByRef rawData As Byte())

End Interface

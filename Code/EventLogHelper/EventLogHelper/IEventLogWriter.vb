''' <summary>
''' Interface for writing to the event log.
''' </summary>
Public Interface IEventLogWriter

    ''' <summary>
    ''' Determines whether the specified event log exists on the local computer.
    ''' </summary>
    ''' <param name="logName">
    ''' The name of the event log to check (e.g., "Application", "System", or a custom log name).
    ''' </param>
    ''' <returns>
    ''' <c>True</c> if the specified event log exists; otherwise, <c>False</c>.
    ''' </returns>
    ''' <remarks>
    ''' This method wraps <see cref="EventLog.Exists(String)"/> to allow the check to be abstracted 
    ''' and testable via <see cref="IEventLogWriter"/>.
    ''' </remarks>
    Function Exists(logName As String) As Boolean

    ''' <summary>
    ''' Determines whether the specified event source is registered with the given event log on the local computer.
    ''' </summary>
    ''' <param name="source">The name of the event source to check.</param>
    ''' <param name="logName">The name of the event log in which to look for the source (e.g., "Application").</param>
    ''' <returns>
    ''' <c>True</c> if the specified event source exists and is registered with the given log; otherwise, <c>False</c>.
    ''' </returns>
    ''' <remarks>
    ''' This method wraps <see cref="EventLog.SourceExists(String, String)"/> and is used to allow 
    ''' testable verification of the event source.
    ''' </remarks>
    Function SourceExists(source As String, logName As String) As Boolean

    ''' <summary>
    ''' Creates a new event source and associates it with the specified event log, if it does not already exist.
    ''' </summary>
    ''' <param name="source">The name of the event source to create.</param>
    ''' <param name="logName">The name of the event log to associate the source with (e.g., "Application").</param>
    ''' <remarks>
    ''' This method wraps <see cref="EventLog.CreateEventSource(String, String)"/> to facilitate testable 
    ''' logging behavior. If the source already exists, no action is taken.
    ''' </remarks>
    Sub CreateEventSource(source As String, logName As String)

    ''' <summary>
    ''' Writes an entry to the specified event log using the given parameters.
    ''' </summary>
    ''' <param name="log">The name of the event log to write to (e.g., "Application" or "CustomLog").</param>
    ''' <param name="source">The source of the log entry. This must be registered with the specified log.</param>
    ''' <param name="message">The message text to log. If too long, it will be truncated to the maximum allowed length.</param>
    ''' <param name="eventType">The type of event to log (e.g., <see cref="EventLogEntryType.Information"/>, <see cref="EventLogEntryType.Error"/>).</param>
    ''' <param name="eventId">The application-defined identifier for the event.</param>
    ''' <param name="category">An application-defined category number for the event.</param>
    ''' <param name="rawData">Optional raw binary data to associate with the entry. Can be <c>Nothing</c> if not used.</param>
    ''' <remarks>
    ''' This method creates a new <see cref="EventLog"/> instance targeting the specified log and writes the entry
    ''' with the given parameters. The source must be associated with the log prior to calling this method.
    ''' </remarks>
    Sub WriteEntry(
            log As String,
            source As String,
            message As String,
            eventType As EventLogEntryType,
            eventId As Integer,
            category As Short,
            rawData As Byte())

End Interface

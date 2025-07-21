''' <summary>
''' Provides a real implementation of the <see cref="IEventLogWriter"/> interface 
''' that logs events directly to the Windows Event Log using the <see cref="System.Diagnostics.EventLog"/> class.
''' </summary>
''' <remarks>
''' This class is intended for production use where actual event logging is required. It encapsulates the logic for:
''' <list type="bullet">
'''   <item><description>Writing entries to the event log</description></item>
'''   <item><description>Checking if a log or source exists</description></item>
'''   <item><description>Creating event log sources if needed</description></item>
''' </list>
''' 
''' In unit test scenarios, consider using a test or stub implementation of <see cref="IEventLogWriter"/> 
''' to avoid writing to the system event log.
''' </remarks>
#If NET35 Then
<Diagnostics.ExcludeFromCodeCoverage>
Public Class RealEventLogWriter
#Else
<CodeAnalysis.ExcludeFromCodeCoverage>
Public Class RealEventLogWriter
#End If

    Implements IEventLogWriter

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
    Public Sub WriteEntry(
            ByVal log As String,
            ByVal source As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventId As Integer,
            ByVal category As Short,
            ByVal rawData As Byte()) Implements IEventLogWriter.WriteEntry

        Using logInstance As New EventLog(log)
            logInstance.Source = source
            logInstance.WriteEntry(
                message,
                eventType,
                eventId,
                category,
                rawData)
        End Using

    End Sub

    ''' <summary>
    ''' Creates a new event source and associates it with the specified event log, if it does not already exist.
    ''' </summary>
    ''' <param name="source">The name of the event source to create.</param>
    ''' <param name="logName">The name of the event log to associate the source with (e.g., "Application").</param><remarks>
    ''' This method wraps <see cref="EventLog.CreateEventSource(String, String)" /> to facilitate testable logging behavior.
    ''' If the source already exists, no action is taken.
    ''' </remarks>
    Public Sub CreateEventSource(
            ByVal source As String,
            ByVal logName As String) Implements IEventLogWriter.CreateEventSource

        If Not SourceExists(source, logName) Then
            EventLog.CreateEventSource(source, logName)
        End If

    End Sub

    ''' <summary>
    ''' Determines whether the specified event log exists on the local computer.
    ''' </summary>
    ''' <param name="logName">
    ''' The name of the event log to check (e.g., "Application", "System", or a custom log name).
    ''' </param>
    ''' <returns>
    ''' <c>True</c> if the specified event log exists; otherwise, <c>False</c>.
    ''' </returns><remarks>
    ''' This method wraps <see cref="EventLog.Exists(String)" /> to allow the check to be abstracted
    ''' and testable via <see cref="IEventLogWriter" />.
    ''' </remarks>
    Public Function Exists(
            ByVal logName As String) As Boolean Implements IEventLogWriter.Exists

        Return EventLog.Exists(logName)

    End Function

    ''' <summary>
    ''' Determines whether the specified event source is registered with the given event log on the local computer.
    ''' </summary>
    ''' <param name="source">The name of the event source to check.</param>
    ''' <param name="logName">The name of the event log in which to look for the source (e.g., "Application").</param>
    ''' <returns>
    ''' <c>True</c> if the specified event source exists and is registered with the given log; otherwise, <c>False</c>.
    ''' </returns><remarks>
    ''' This method wraps <see cref="EventLog.SourceExists(String, String)" /> and is used to allow
    ''' testable verification of the event source.
    ''' </remarks>
    Public Function SourceExists(
            ByVal source As String,
            ByVal logName As String) As Boolean Implements IEventLogWriter.SourceExists

        Return EventLog.SourceExists(source, logName)

    End Function

End Class

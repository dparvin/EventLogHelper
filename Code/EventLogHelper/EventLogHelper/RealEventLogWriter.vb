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
Friend Class RealEventLogWriter
#Else
<CodeAnalysis.ExcludeFromCodeCoverage>
Friend Class RealEventLogWriter
#End If

    Implements IEventLogWriter

    ''' <summary>
    ''' Writes an entry to the specified event log using the given parameters.
    ''' </summary>
    ''' <param name="logName">The name of the event log to write to (e.g., "Application" or "CustomLog").</param>
    ''' <param name="sourceName">The source of the log entry. This must be registered with the specified log.</param>
    ''' <param name="message">The message text to log. If too long, it will be truncated to the maximum allowed length.</param>
    ''' <param name="eventType">The type of event to log (e.g., <see cref="EventLogEntryType.Information"/>, <see cref="EventLogEntryType.Error"/>).</param>
    ''' <param name="eventId">The application-defined identifier for the event.</param>
    ''' <param name="category">An application-defined category number for the event.</param>
    ''' <param name="rawData">Optional raw binary data to associate with the entry. Can be <c>Nothing</c> if not used.</param>
    ''' <remarks>
    ''' This method creates a new <see cref="EventLog"/> instance targeting the specified log and writes the entry
    ''' with the given parameters. The source must be associated with the log prior to calling this method.
    ''' </remarks>
    Friend Sub WriteEntry(
            ByVal machineName As String,
            ByVal logName As String,
            ByVal sourceName As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventId As Integer,
            ByVal category As Short,
            ByVal rawData As Byte(),
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer,
            ByVal writeInitEntry As Boolean) Implements IEventLogWriter.WriteEntry

        Using logInstance As EventLog = GetLog(
                machineName,
                logName,
                sourceName,
                maxKilobytes,
                retentionDays,
                writeInitEntry)
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
    ''' <param name="sourceName">The name of the event source to create.</param>
    ''' <param name="logName">The name of the event log to associate the source with (e.g., "Application").</param><remarks>
    ''' This method wraps <see cref="EventLog.CreateEventSource(String, String, String)" /> to facilitate testable logging behavior.
    ''' If the source already exists, no action is taken.
    ''' </remarks>
    Friend Sub CreateEventSource(
            ByVal sourceName As String,
            ByVal logName As String,
            ByVal machineName As String) Implements IEventLogWriter.CreateEventSource

        If Not SourceExists(sourceName, machineName) Then
            Dim obj As New EventSourceCreationData(sourceName, logName) With {
                .MachineName = machineName
            }
            EventLog.CreateEventSource(obj)
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
    Friend Function Exists(
            ByVal logName As String,
            ByVal machineName As String) As Boolean Implements IEventLogWriter.Exists

        Return EventLog.Exists(logName, machineName)

    End Function

    ''' <summary>
    ''' Determines whether the specified event source is registered with the given event log on the local computer.
    ''' </summary>
    ''' <param name="source">The name of the event source to check.</param>
    ''' <returns>
    ''' <c>True</c> if the specified event source exists and is registered with the given log; otherwise, <c>False</c>.
    ''' </returns><remarks>
    ''' This method wraps <see cref="EventLog.SourceExists(String, String)" /> and is used to allow
    ''' testable verification of the event source.
    ''' </remarks>
    Friend Function SourceExists(
            ByVal source As String,
            ByVal machineName As String) As Boolean Implements IEventLogWriter.SourceExists

        Return EventLog.SourceExists(source, machineName)

    End Function

    Friend Function GetLog(
            machineName As String,
            logName As String,
            sourceName As String,
            maxKilobytes As Integer,
            retentionDays As Integer,
            writeInitEntry As Boolean) As EventLog Implements IEventLogWriter.GetLog

        If String.IsNullOrEmpty(logName) Then Throw New ArgumentNullException(NameOf(logName))
        If String.IsNullOrEmpty(sourceName) Then Throw New ArgumentNullException(NameOf(sourceName))
        If String.IsNullOrEmpty(machineName) Then machineName = "."

        Try
            ' Source registration can only be done on the local machine
            If machineName = "." OrElse String.Equals(machineName, Environment.MachineName, StringComparison.OrdinalIgnoreCase) Then
                If Not EventLog.SourceExists(sourceName) Then
                    Dim sourceData As New EventSourceCreationData(sourceName, logName)
                    EventLog.CreateEventSource(sourceData)
                    ' Optionally wait here or notify user that restart might be needed
                End If
            End If

            Dim el As New EventLog(logName, machineName, sourceName)

            If writeInitEntry Then
                el.WriteEntry($"Initialized log '{logName}' with source '{sourceName}' on machine '{machineName}'.", EventLogEntryType.Information)
            End If

            ' Only applies to local machine logs
            el.MaximumKilobytes = maxKilobytes
            el.ModifyOverflowPolicy(OverflowAction.OverwriteOlder, retentionDays)

            Return el

        Catch ex As Exception
            Throw New ApplicationException($"Failed to initialize event log '{logName}' with source '{sourceName}' on machine '{machineName}'.", ex)
        End Try

    End Function

    ''' <summary>
    ''' Writes the entry.
    ''' </summary>
    ''' <param name="eventLog">The log.</param>
    ''' <param name="message">The message.</param>
    ''' <param name="eventType">The type.</param>
    ''' <param name="eventID">The event identifier.</param>
    ''' <param name="category">The category.</param>
    ''' <param name="rawData">The raw data.</param>
    Friend Sub WriteEntry(
            ByRef eventLog As EventLog,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByRef rawData() As Byte) Implements IEventLogWriter.WriteEntry

        eventLog.WriteEntry(
            message,
            eventType,
            eventID,
            category,
            rawData)

    End Sub

End Class

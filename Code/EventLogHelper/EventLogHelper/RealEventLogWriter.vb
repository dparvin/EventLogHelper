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

    ''' <summary>
    ''' Retrieves the event log name that the specified source is registered under 
    ''' on the given machine.
    ''' </summary>
    ''' <param name="sourceName">
    ''' The event source to check. Must not be <c>null</c> or empty.
    ''' </param>
    ''' <param name="machineName">
    ''' The target machine name, or <c>"."</c> for the local machine.
    ''' </param>
    ''' <returns>
    ''' The log name (e.g., "Application", "System") associated with the source.
    ''' </returns>
    ''' <exception cref="ArgumentNullException">
    ''' Thrown if <paramref name="sourceName"/> is null or empty.
    ''' </exception>
    ''' <exception cref="ApplicationException">
    ''' Thrown if the log name cannot be retrieved for the specified source 
    ''' on the machine, wrapping the original exception for context.
    ''' </exception>
    ''' <remarks>
    ''' This implementation calls <see cref="EventLog.LogNameFromSourceName"/> 
    ''' internally and wraps any unexpected failures in an <see cref="ApplicationException"/>.
    ''' </remarks>
    Friend Function GetLogForSource(
            ByVal sourceName As String,
            ByVal machineName As String) As String Implements IEventLogWriter.GetLogForSource

        If String.IsNullOrEmpty(sourceName) Then Throw New ArgumentNullException(NameOf(sourceName))
        Try
            Dim logName As String = EventLog.LogNameFromSourceName(sourceName, machineName)
            Return logName
        Catch ex As Exception
            Throw New ApplicationException($"Failed to retrieve log name for source '{sourceName}' on machine '{machineName}'.", ex)
        End Try

    End Function

    ''' <summary>
    ''' Retrieves or initializes an <see cref="EventLog" /> instance with the specified configuration.
    ''' </summary>
    ''' <param name="machineName">The name of the machine (use "." for the local machine).</param>
    ''' <param name="logName">The name of the log (e.g., "Application" or a custom log).</param>
    ''' <param name="sourceName">The name of the event source to use or create.</param>
    ''' <param name="maxKilobytes">Maximum size of the log in kilobytes.</param>
    ''' <param name="retentionDays">Number of days to retain log entries before overwriting.</param>
    ''' <param name="writeInitEntry">If <c>True</c>, an initialization entry is written upon setup.</param>
    ''' <returns>
    ''' An <see cref="EventLog" /> instance configured with the specified parameters.
    ''' </returns>
    ''' <exception cref="ArgumentNullException">
    ''' logName
    ''' or
    ''' sourceName
    ''' </exception>
    ''' <exception cref="InvalidOperationException">
    ''' The source '{sourceToUse}' is registered under a different log than '{logName}' on machine '{machineName}'.
    ''' or
    ''' No default source could be determined for log '{logName}' on machine '{machineName}'. " &amp;
    '''                             "Please ensure the log is properly initialized with a valid source.
    ''' </exception>
    ''' <exception cref="ApplicationException">Failed to initialize event log '{logName}' with source '{sourceToUse}' on machine '{machineName}'.</exception>
    Friend Function GetLog(
            ByVal machineName As String,
            ByVal logName As String,
            ByVal sourceName As String,
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer,
            ByVal writeInitEntry As Boolean) As EventLog Implements IEventLogWriter.GetLog

        If String.IsNullOrEmpty(logName) Then Throw New ArgumentNullException(NameOf(logName))
        If String.IsNullOrEmpty(sourceName) Then Throw New ArgumentNullException(NameOf(sourceName))
        If String.IsNullOrEmpty(machineName) Then machineName = "."
        Dim sourceToUse As String = sourceName ' this is used to tell the process which source to actually use if we can't use the one we want to use.

        Try
            If Not Exists(logName, machineName) OrElse Not SourceExists(sourceToUse, machineName) Then
                CreateEventSource(sourceName, logName, machineName)
            ElseIf Not IsSourceRegisteredToLog(sourceToUse, logName, machineName) Then
                Select Case SmartEventLogger.SourceResolutionBehavior
                    Case SourceResolutionBehavior.Strict
                        ' Fail immediately with a clear exception message.
                        Throw New InvalidOperationException(
                        $"The source '{sourceToUse}' is registered under a different log than '{logName}' on machine '{machineName}'.")

                    Case SourceResolutionBehavior.UseSourceLog
                        ' Adjust the log name to match the one the source is already tied to.
                        logName = GetLogForSource(sourceToUse, machineName)

                    Case SourceResolutionBehavior.UseLogsDefaultSource
                        sourceToUse = DefaultSource(logName, machineName)
                        If String.IsNullOrEmpty(sourceToUse) Then
                            Throw New InvalidOperationException(
                            $"No default source could be determined for log '{logName}' on machine '{machineName}'. " &
                            "Please ensure the log is properly initialized with a valid source.")
                        End If
                End Select
            End If

            Dim el As New EventLog(logName, machineName, sourceToUse)

            If writeInitEntry Then
                el.WriteEntry($"Initialized log '{logName}' with source '{sourceToUse}' on machine '{machineName}'.", EventLogEntryType.Information)
            End If

            ' Only applies to local machine logs
            el.MaximumKilobytes = maxKilobytes
            el.ModifyOverflowPolicy(OverflowAction.OverwriteOlder, retentionDays)

            Return el

        Catch ex As Exception
            Throw New ApplicationException($"Failed to initialize event log '{logName}' with source '{sourceToUse}' on machine '{machineName}'.", ex)
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

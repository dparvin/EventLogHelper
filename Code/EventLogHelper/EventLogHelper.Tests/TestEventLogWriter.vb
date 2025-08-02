#If NET35 Then
#Else
Imports Xunit.Abstractions
#End If

''' <summary>
''' Test implementation of the <see cref="IEventLogWriter" /> interface for unit testing purposes.
''' This class simulates writing to an event log without actually performing any logging operations.
''' It can be used to verify that the code interacts with the event log correctly without needing
''' access to the actual Windows Event Log.
''' </summary>
''' <seealso cref="IEventLogWriter" />
#If NET35 Then
<Diagnostics.ExcludeFromCodeCoverage>
Public Class TestEventLogWriter
#Else
<CodeAnalysis.ExcludeFromCodeCoverage>
Public Class TestEventLogWriter
#End If

    Implements IEventLogWriter

#If NET35 Then
    ''' <summary>
    ''' Initializes a new instance of the <see cref="TestEventLogWriter"/> class.
    ''' </summary>
    ''' <param name="logExists">if set to <c>true</c> [log exists].</param>
    ''' <param name="sourceExists">if set to <c>true</c> [source exists].</param>
    Public Sub New(
            Optional ByVal logExists As Boolean = True,
            Optional ByVal sourceExists As Boolean = True)
#Else
    ''' <summary>
    ''' Initializes a new instance of the <see cref="TestEventLogWriter"/> class.
    ''' </summary>
    ''' <param name="output"></param>
    ''' <param name="logExists">if set to <c>true</c> [log exists].</param>
    ''' <param name="sourceExists">if set to <c>true</c> [source exists].</param>
    Public Sub New(
            ByRef output As ITestOutputHelper,
            Optional ByVal logExists As Boolean = True,
            Optional ByVal sourceExists As Boolean = True)

        OutputHelper = output
#End If
        ReturnLogExists = logExists
        ReturnSourceExists = sourceExists

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
    ''' Gets or sets a value indicating whether to simulate that the log exists.
    ''' </summary>
    ''' <value>
    '''   <c>true</c> if simulating that the log exists; otherwise, <c>false</c>.
    ''' </value>
    Public ReadOnly Property ReturnLogExists As Boolean

    ''' <summary>
    ''' Gets or sets a value indicating whether to simulate that the source exists.
    ''' </summary>
    ''' <value>
    '''   <c>true</c> if simulating that the source exists; otherwise, <c>false</c>.
    ''' </value>
    Public ReadOnly Property ReturnSourceExists As Boolean

    ''' <summary>
    ''' Gets or sets a value indicating whether the create event source method was called.
    ''' </summary>
    ''' <value>
    '''   <c>true</c> if the create event source method was called; otherwise, <c>false</c>.
    ''' </value>
    Public Property CreateEventSourceCalled As Boolean

    ''' <summary>
    ''' Gets or sets a value indicating whether the write entry method was called.
    ''' </summary>
    ''' <value>
    '''   <c>true</c> if the write entry method was called; otherwise, <c>false</c>.
    ''' </value>
    Public Property WriteEntryCalled As Boolean

    ''' <summary>
    ''' Gets or sets a value indicating whether the raw data length is set to 0.
    ''' </summary>
    ''' <value>
    '''   <c>true</c> if the raw data length is set to 0; otherwise, <c>false</c>.
    ''' </value>
    Public Property MessageLength As Integer

    ''' <summary>
    ''' Gets or sets the last message.
    ''' </summary>
    ''' <value>
    ''' The last message.
    ''' </value>
    Public Property LastMessage As String

    ''' <summary>
    ''' Gets or sets a value indicating whether the raw data length is set to 0.
    ''' </summary>
    ''' <value>
    '''   <c>true</c> if the raw data length is set to 0; otherwise, <c>false</c>.
    ''' </value>
    Public Property SourceLength As Integer

    ''' <summary>
    ''' Creates a new event source and associates it with the specified event log, if it does not already exist.
    ''' </summary>
    ''' <param name="source">The name of the event source to create.</param>
    ''' <param name="logName">The name of the event log to associate the source with (e.g., "Application").</param><remarks>
    ''' This method wraps <see cref="M:System.Diagnostics.EventLog.CreateEventSource(System.String,System.String)" /> to 
    ''' facilitate testable logging behavior. If the source already exists, no action is taken.
    ''' </remarks>
    Public Sub CreateEventSource(
            ByVal source As String,
            ByVal logName As String,
            ByVal MachineName As String) Implements IEventLogWriter.CreateEventSource

        ' Simulate creating an event source
        Output($"Creating event source: {source} for log: {logName}")
        ' In a real implementation, you would create the event source here
        ' For example:
        ' If Not EventLog.SourceExists(source, logName) Then
        '     EventLog.CreateEventSource(source, logName)
        ' End If
        CreateEventSourceCalled = True

    End Sub

    ''' <summary>
    ''' Writes an entry to the specified event log using the given parameters.
    ''' </summary>
    ''' <param name="logName">The name of the event log to write to (e.g., "Application" or "CustomLog").</param>
    ''' <param name="sourceName">The source of the log entry. This must be registered with the specified log.</param>
    ''' <param name="message">
    ''' The message text to log. If too long, it will be truncated to the maximum allowed length.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log (e.g., <see cref="F:System.Diagnostics.EventLogEntryType.Information" />, 
    ''' <see cref="F:System.Diagnostics.EventLogEntryType.Error" />).
    ''' </param>
    ''' <param name="eventId">The application-defined identifier for the event.</param>
    ''' <param name="category">An application-defined category number for the event.</param>
    ''' <param name="rawData">
    ''' Optional raw binary data to associate with the entry. Can be <c>Nothing</c> if not used.
    ''' </param>
    ''' <remarks>
    ''' This method creates a new <see cref="T:System.Diagnostics.EventLog" /> instance targeting the 
    ''' specified log and writes the entry with the given parameters. The source must be associated with 
    ''' the log prior to calling this method.
    ''' </remarks>
    Public Sub WriteEntry(
            ByVal machineName As String,
            ByVal logName As String,
            ByVal sourceName As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventId As Integer,
            ByVal category As Short,
            ByVal rawData() As Byte,
            ByVal maxKilobytes As Integer,
            ByVal retentionDays As Integer,
            ByVal writeInitEntry As Boolean) Implements IEventLogWriter.WriteEntry

        ' Simulate writing an entry to the event log
        Output(String.Join(Environment.NewLine, {
            $"Machine Name: {machineName}",
            $"Log Name: {logName}",
            $"Source Name: {sourceName}",
            $"Message: {message}",
            $"EventType: {eventType}",
            $"EventID: {eventId}",
            $"Category: {category}",
            $"RawData Length: {If(rawData IsNot Nothing, rawData.Length, 0)}",
            $"MaxKilobytes: {maxKilobytes}",
            $"RetentionDays: {retentionDays}",
            $"WriteInitEntry: {writeInitEntry}"
        }))

        LastMessage = message
        MessageLength = message.Length
        SourceLength = sourceName.Length
        WriteEntryCalled = True

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
    ''' This method wraps <see cref="M:System.Diagnostics.EventLog.Exists(System.String)" /> to allow the check 
    ''' to be abstracted and testable via <see cref="T:EventLogHelper.IEventLogWriter" />.
    ''' </remarks>
    Public Function Exists(
            ByVal logName As String,
            ByVal machineName As String) As Boolean Implements IEventLogWriter.Exists

        Output($"Checking if log exists: {logName} on machine {machineName}, Returning {ReturnLogExists}")
        Return ReturnLogExists

    End Function

    ''' <summary>
    ''' Determines whether the specified event source is registered with the given event log on the local computer.
    ''' </summary>
    ''' <param name="source">The name of the event source to check.</param>
    ''' <param name="machineName">The name of the event log in which to look for the source (e.g., "Application").</param>
    ''' <returns>
    ''' <c>True</c> if the specified event source exists and is registered with the given log; otherwise, <c>False</c>.
    ''' </returns><remarks>
    ''' This method wraps <see cref="M:System.Diagnostics.EventLog.SourceExists(System.String,System.String)" /> and is used to allow
    ''' testable verification of the event source.
    ''' </remarks>
    Public Function SourceExists(
            source As String,
            machineName As String) As Boolean Implements IEventLogWriter.SourceExists

        Output($"Checking if source exists: {source} on machine: {machineName}, Returning {ReturnSourceExists}")
        Return ReturnSourceExists

    End Function

    Public Function GetLog(
            machineName As String,
            logName As String,
            sourceName As String,
            maxKilobytes As Integer,
            retentionDays As Integer,
            writeInitEntry As Boolean) As EventLog Implements IEventLogWriter.GetLog

        Output(String.Join(Environment.NewLine, {
            $"Machine Name: {machineName}",
            $"Log Name: {logName}",
            $"Source Name: {sourceName}",
            $"MaxKilobytes: {maxKilobytes}",
            $"RetentionDays: {retentionDays}",
            $"WriteInitEntry: {writeInitEntry}"
        }))

        Return New EventLog()

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

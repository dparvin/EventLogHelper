''' <summary>
''' This module provides functionality for logging events in a smart way.
''' It can be used to log various types of events, such as errors, warnings, and informational messages.
''' The implementation details are not provided here, but this module serves as a placeholder for future enhancements.
''' </summary>
Public Module SmartEventLogger

    ''' <summary>
    ''' The writer
    ''' </summary>
    Private _writer As IEventLogWriter = New RealEventLogWriter()

    ''' <summary>
    ''' Sets the writer.
    ''' </summary>
    ''' <param name="writer">The writer.</param>
    Public Sub SetWriter(writer As IEventLogWriter)
        _writer = writer
    End Sub

    ''' <summary>
    ''' Logs the specified log.
    ''' </summary>
    ''' <param name="log">The log.</param>
    ''' <param name="source">The source.</param>
    ''' <param name="message">The message.</param>
    ''' <param name="eventType">Type of the event.</param>
    ''' <param name="eventID">The event identifier.</param>
    ''' <param name="category">The category.</param>
    ''' <param name="rawData">The raw data.</param>
    Public Sub Log(
            ByVal log As String,
            ByVal source As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByVal rawData As Byte())

        Dim defaultLog As String = If(String.IsNullOrEmpty(log), "Application", log.Trim())
        Dim defaultSource As String = If(String.IsNullOrEmpty(source), "SmartEventLogger", source.Trim())

        If Not EventLog.Exists(defaultLog) OrElse Not EventLog.SourceExists(defaultSource, defaultLog) Then
            EventLog.CreateEventSource(defaultSource, defaultLog)
        End If
        Dim finalMessage As String = If(String.IsNullOrEmpty(message), "No message provided.", message.Trim())
        If Not finalMessage.EndsWith("."c) Then finalMessage &= "."
        If Not finalMessage.StartsWith("["c) Then finalMessage = $"[{defaultSource}] {finalMessage}"
        If finalMessage.Length > 32766 Then finalMessage = finalMessage.Substring(0, 32766)
        Select Case eventType
            Case EventLogEntryType.Error, EventLogEntryType.Warning, EventLogEntryType.Information
                ' Valid
            Case Else
                eventType = EventLogEntryType.Information
        End Select
        Using logInstance As New EventLog(defaultLog)
            _writer.WriteEntry(defaultLog, defaultSource, finalMessage, eventType, eventID, category, rawData)
        End Using

    End Sub

End Module

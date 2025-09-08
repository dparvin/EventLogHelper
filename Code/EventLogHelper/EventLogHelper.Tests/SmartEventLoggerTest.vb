Imports Xunit
Imports System.ComponentModel
Imports System.Reflection


#If NET35 Then
Imports System.Diagnostics
Imports Xunit.Extensions
#Else
Imports Xunit.Abstractions
#End If

#If NET35 Then
Namespace net35
#ElseIf NET452 Then
Namespace net452
#ElseIf NET462 Then
Namespace net462
#ElseIf NET472 Then
Namespace net472
#ElseIf NET481 Then
Namespace net481
#ElseIf NET8_0 Then
Namespace net80
#ElseIf NET9_0 Then
Namespace net90
#End If

    ''' <summary>
    ''' Test class for <see cref="SmartEventLogger" />.
    ''' This class is used to test the SmartEventLogger functionality.
    ''' It is designed to be run with xUnit and outputs results using ITestOutputHelper.
    ''' </summary>
#If NET5_0_OR_GREATER Then
    <Runtime.Versioning.SupportedOSPlatform("windows")>
    Public Class SmartEventLoggerTest
#Else
    Public Class SmartEventLoggerTest
#End If

#If NET35 Then
        Sub New()

            SmartEventLogger.Reset()

        End Sub
#Else
        Private OutputHelper As ITestOutputHelper

        ''' <summary>
        ''' Initializes a new instance of the <see cref="SmartEventLoggerTest"/> class.
        ''' </summary>
        ''' <param name="output">The output.</param>
        Sub New(ByVal output As ITestOutputHelper)

            OutputHelper = output
            SmartEventLogger.Reset()

        End Sub
#End If

#Region " Test Methods ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ "

        ''' <summary>
        ''' Logs the message.
        ''' </summary>
        <Theory>
        <InlineData(LoggingSeverity.Critical, True)>
        <InlineData(LoggingSeverity.Error, True)>
        <InlineData(LoggingSeverity.Info, True)>
        <InlineData(LoggingSeverity.Warning, True)>
        <InlineData(LoggingSeverity.Verbose, False)>
        Public Sub Log_Message_Severity(
                ByVal Severity As LoggingSeverity,
                ByVal WriteEntry As Boolean)

            ' Arrange
            Dim ll As LoggingLevel = CurrentLoggingLevel
            CurrentLoggingLevel = LoggingLevel.Normal
#If NET35 Then
            Dim testWriter As New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            Dim testWriter As New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim message As String = "Test message"

            ' Act
            Log(message, Severity)

            ' Assert
            CurrentLoggingLevel = ll ' Restore original logging level
            Assert.Equal(WriteEntry, testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the message.
        ''' </summary>
        <Fact>
        Public Sub Log_Message()

            ' Arrange
#If NET35 Then
            Dim testWriter As New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            Dim testWriter As New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim message As String = "Test message"

            ' Act
            Log(message)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the source name message.
        ''' </summary>
        <Fact>
        Public Sub Log_SourceName_Message_Severity()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim sourceName As String = "CustomSource"
            Dim message As String = "Test message"

            ' Act
            Log(sourceName, message, LoggingSeverity.Error)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal($"[{sourceName}] {message}.", testWriter.LastMessage)

        End Sub

        ''' <summary>
        ''' Logs the source name message.
        ''' </summary>
        <Fact>
        Public Sub Log_SourceName_Message()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim sourceName As String = "CustomSource"
            Dim message As String = "Test message"

            ' Act
            Log(sourceName, message)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal($"[{sourceName}] {message}.", testWriter.LastMessage)

        End Sub

        ''' <summary>
        ''' Logs the log name source name message.
        ''' </summary>
        <Fact>
        Public Sub Log_LogName_SourceName_Message()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim logName As String = "CustomLog"
            Dim sourceName As String = "CustomSource"
            Dim message As String = "Test message"

            ' Act
            Log(logName, sourceName, message)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal($"[{sourceName}] {message}.", testWriter.LastMessage)

        End Sub

        ''' <summary>
        ''' Logs the type of the message event.
        ''' </summary>
        <Fact>
        Public Sub Log_Message_EventType_EntrySeverity()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim message As String = "Test message"
            Dim eventType As EventLogEntryType = EventLogEntryType.Warning

            ' Act
            Log(message, eventType, LoggingSeverity.Error)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the type of the message event.
        ''' </summary>
        <Fact>
        Public Sub Log_Message_EventType()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim message As String = "Test message"
            Dim eventType As EventLogEntryType = EventLogEntryType.Warning

            ' Act
            Log(message, eventType)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the type of the source name message event.
        ''' </summary>
        <Fact>
        Public Sub Log_SourceName_Message_EventType()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim sourceName As String = "CustomSource"
            Dim message As String = "Test message"
            Dim eventType As EventLogEntryType = EventLogEntryType.Warning

            ' Act
            Log(sourceName, message, eventType)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal($"[{sourceName}] {message}.", testWriter.LastMessage)

        End Sub

        ''' <summary>
        ''' Logs the log name source name message event type event identifier.
        ''' </summary>
        <Fact>
        Public Sub Log_LogName_SourceName_Message_EventType()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim logName As String = "CustomLog"
            Dim sourceName As String = "CustomSource"
            Dim message As String = "Test message"
            Dim eventType As EventLogEntryType = EventLogEntryType.Warning

            ' Act
            Log(logName, sourceName, message, eventType)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal($"[{sourceName}] {message}.", testWriter.LastMessage)

        End Sub

        ''' <summary>
        ''' Logs the message event type event identifier.
        ''' </summary>
        <Fact>
        Public Sub Log_Message_EventType_EventId_EntrySeverity()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim message As String = "Test message"
            Dim eventType As EventLogEntryType = EventLogEntryType.Warning
            Dim eventId As Integer = 42

            ' Act
            Log(message, eventType, eventId, LoggingSeverity.Error)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the message event type event identifier.
        ''' </summary>
        <Fact>
        Public Sub Log_Message_EventType_EventId()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim message As String = "Test message"
            Dim eventType As EventLogEntryType = EventLogEntryType.Warning
            Dim eventId As Integer = 42

            ' Act
            Log(message, eventType, eventId)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the message event type event identifier category.
        ''' </summary>
        <Fact>
        Public Sub Log_Message_EventType_EventId_Category_EntrySeverity()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim message As String = "Test message"
            Dim eventType As EventLogEntryType = EventLogEntryType.Warning
            Dim eventId As Integer = 42
            Dim category As Short = 1

            ' Act
            Log(message, eventType, eventId, category, LoggingSeverity.Error)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the message event type event identifier category.
        ''' </summary>
        <Fact>
        Public Sub Log_Message_EventType_EventId_Category()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim message As String = "Test message"
            Dim eventType As EventLogEntryType = EventLogEntryType.Warning
            Dim eventId As Integer = 42
            Dim category As Short = 1

            ' Act
            Log(message, eventType, eventId, category)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the log name source name message event type event identifier.
        ''' </summary>
        <Fact>
        Public Sub Log_LogName_SourceName_Message_EventType_EventId()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim logName As String = "CustomLog"
            Dim sourceName As String = "CustomSource"
            Dim message As String = "Test message"
            Dim eventType As EventLogEntryType = EventLogEntryType.Warning
            Dim eventId As Integer = 42

            ' Act
            Log(logName, sourceName, message, eventType, eventId)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal($"[{sourceName}] {message}.", testWriter.LastMessage)

        End Sub

        ''' <summary>
        ''' Logs the log name source name message event type event identifier category.
        ''' </summary>
        <Fact>
        Public Sub Log_SourceName_Message_EventType_EventId_Category()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim sourceName As String = "CustomSource"
            Dim message As String = "Test message"
            Dim eventType As EventLogEntryType = EventLogEntryType.Warning
            Dim eventId As Integer = 42
            Dim category As Short = 1

            ' Act
            Log(sourceName, message, eventType, eventId, category)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal($"[{sourceName}] {message}.", testWriter.LastMessage)

        End Sub

        ''' <summary>
        ''' Logs the log name source name message event type event identifier category.
        ''' </summary>
        <Fact>
        Public Sub Log_LogName_SourceName_Message_EventType_EventId_Category_RawData_EntrySeverity()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim logName As String = "CustomLog"
            Dim sourceName As String = "CustomSource"
            Dim message As String = "Test message"
            Dim eventType As EventLogEntryType = EventLogEntryType.Warning
            Dim eventId As Integer = 42
            Dim category As Short = 1

            ' Act
            Log(logName, sourceName, message, eventType, eventId, category, Nothing, LoggingSeverity.Error)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal($"[{sourceName}] {message}.", testWriter.LastMessage)

        End Sub

        ''' <summary>
        ''' Logs the log name source name message event type event identifier category.
        ''' </summary>
        <Fact>
        Public Sub Log_LogName_SourceName_Message_EventType_EventId_Category()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim logName As String = "CustomLog"
            Dim sourceName As String = "CustomSource"
            Dim message As String = "Test message"
            Dim eventType As EventLogEntryType = EventLogEntryType.Warning
            Dim eventId As Integer = 42
            Dim category As Short = 1

            ' Act
            Log(logName, sourceName, message, eventType, eventId, category)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal($"[{sourceName}] {message}.", testWriter.LastMessage)

        End Sub

        ''' <summary>
        ''' Logs the message raw data.
        ''' </summary>
        <Fact>
        Public Sub Log_Message_RawData_EventSeverity()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim message As String = "Test message"
            Dim rawData As Byte() = Nothing

            ' Act
            Log(message, rawData, LoggingSeverity.Error)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the message raw data.
        ''' </summary>
        <Fact>
        Public Sub Log_Message_RawData()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim message As String = "Test message"
            Dim rawData As Byte() = Nothing

            ' Act
            Log(message, rawData)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the message event type raw data.
        ''' </summary>
        <Fact>
        Public Sub Log_Message_EventType_RawData()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim message As String = "Test message"
            Dim eventType As EventLogEntryType = EventLogEntryType.Warning
            Dim rawData As Byte() = Nothing

            ' Act
            Log(message, eventType, rawData)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the message event type event identifier raw data.
        ''' </summary>
        <Fact>
        Public Sub Log_Message_EventType_EventId_RawData()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim message As String = "Test message"
            Dim eventType As EventLogEntryType = EventLogEntryType.Warning
            Dim eventId As Integer = 42
            Dim rawData As Byte() = Nothing

            ' Act
            Log(message, eventType, eventId, rawData)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the message event type event identifier category raw data.
        ''' </summary>
        <Fact>
        Public Sub Log_Message_EventType_EventId_Category_RawData()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim message As String = "Test message"
            Dim eventType As EventLogEntryType = EventLogEntryType.Warning
            Dim eventId As Integer = 42
            Dim category As Short = 1
            Dim rawData As Byte() = Nothing

            ' Act
            Log(message, eventType, eventId, category, rawData)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the message event type event identifier category raw data.
        ''' </summary>
        <Fact>
        Public Sub Log_Message_maxKilobytes_retentionDays_EntrySeverity()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim message As String = "Test message"
            Dim maxKilobytes As Integer = 1024 * 1024 ' 1 MB
            Dim retentionDays As Integer = 7

            ' Act
            Log(message, maxKilobytes, retentionDays, LoggingSeverity.Error)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the message event type event identifier category raw data.
        ''' </summary>
        <Fact>
        Public Sub Log_Message_maxKilobytes_retentionDays()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim message As String = "Test message"
            Dim maxKilobytes As Integer = 1024 * 1024 ' 1 MB
            Dim retentionDays As Integer = 7

            ' Act
            Log(message, maxKilobytes, retentionDays)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the message event type event identifier category raw data.
        ''' </summary>
        <Fact>
        Public Sub Log_Message_maxKilobytes_retentionDays_writeInitEntry_EntrySeverity()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)
            MachineName = ""
            TruncationMarker = ""

            Dim message As String = "Test message"
            Dim maxKilobytes As Integer = 1024 * 1024 ' 1 MB
            Dim retentionDays As Integer = 7
            Dim writeInitEntry As Boolean = True

            ' Act
            Log(message, maxKilobytes, retentionDays, writeInitEntry, LoggingSeverity.Error)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the message event type event identifier category raw data.
        ''' </summary>
        <Fact>
        Public Sub Log_Message_maxKilobytes_retentionDays_writeInitEntry()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)
            MachineName = ""
            TruncationMarker = ""

            Dim message As String = "Test message"
            Dim maxKilobytes As Integer = 1024 * 1024 ' 1 MB
            Dim retentionDays As Integer = 7
            Dim writeInitEntry As Boolean = True

            ' Act
            Log(message, maxKilobytes, retentionDays, writeInitEntry)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the source name message event type event identifier category raw data.
        ''' </summary>
        <Fact>
        Public Sub Log_SourceName_Message_EventType_EventId_Category_RawData_EntrySeverity()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)
            MachineName = "SomeMachineName"
            Dim sourceName As String = "CustomSource"
            Dim message As String = "Test message"
            Dim eventType As EventLogEntryType = EventLogEntryType.Warning
            Dim eventId As Integer = 42
            Dim category As Short = 1
            Dim rawData As Byte() = Nothing

            ' Act
            Log(sourceName, message, eventType, eventId, category, rawData, LoggingSeverity.Error)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal($"[{sourceName}] {message}.", testWriter.LastMessage)

        End Sub

        ''' <summary>
        ''' Logs the source name message event type event identifier category raw data.
        ''' </summary>
        <Fact>
        Public Sub Log_SourceName_Message_EventType_EventId_Category_RawData()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)
            MachineName = "SomeMachineName"
            Dim sourceName As String = "CustomSource"
            Dim message As String = "Test message"
            Dim eventType As EventLogEntryType = EventLogEntryType.Warning
            Dim eventId As Integer = 42
            Dim category As Short = 1
            Dim rawData As Byte() = Nothing

            ' Act
            Log(sourceName, message, eventType, eventId, category, rawData)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal($"[{sourceName}] {message}.", testWriter.LastMessage)

        End Sub

        ''' <summary>
        ''' Logs the creates event source when log or source does not exist.
        ''' </summary>
        <Fact>
        Public Sub Log_CreatesEventSource_WhenLogOrSourceDoesNotExist()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim logName As String = "CustomLog"
            Dim sourceName As String = "CustomSource"
            Dim message As String = "Test message"
            Dim eventType As EventLogEntryType = EventLogEntryType.Warning
            Dim eventId As Integer = 42
            Dim category As Short = 1
            Dim rawData As Byte() = Nothing

            ' Act
            Log(logName, sourceName, message, eventType, eventId, category, rawData)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal($"[{sourceName}] {message}.", testWriter.LastMessage)

        End Sub

        ''' <summary>
        ''' Logs the creates event source when log and source does not provided.
        ''' </summary>
        <Fact>
        Public Sub Log_CreatesEventSource_WhenLogandSourceAreNotProvided()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim logName As String = ""
            Dim sourceName As String = ""
            Dim message As String = ""
            Dim eventType As EventLogEntryType = EventLogEntryType.FailureAudit
            Dim eventId As Integer = 27
            Dim category As Short = 3
            Dim rawData As Byte() = Nothing

            ' Act
            Log(logName, sourceName, message, eventType, eventId, category, rawData)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the exists and source exist.
        ''' </summary>
        ''' <param name="logExists">if set to <c>true</c> log exists.</param>
        ''' <param name="sourceExists">if set to <c>true</c> source exists.</param>
        <Theory>
        <InlineData(False, False)>
        <InlineData(False, True)>
        <InlineData(True, False)>
        <InlineData(True, True)>
        Public Sub Log_Exists_And_Source_Exist(logExists As Boolean, sourceExists As Boolean)

            ' Arrange
#If NET35 Then
            Dim testWriter As New TestEventLogWriter(logExists:=logExists, sourceExists:=sourceExists)
            Dim mockReader As New TestRegistryReader()
#Else
            Dim testWriter As New TestEventLogWriter(OutputHelper, logExists:=logExists, sourceExists:=sourceExists)
            Dim mockReader As New TestRegistryReader(OutputHelper)
#End If
            SetWriter(testWriter)
            SetRegistryReader(mockReader)

            Dim logName As String = ""
            Dim sourceName As String = ""
            Dim message As String = ""
            Dim eventType As EventLogEntryType = EventLogEntryType.FailureAudit
            Dim eventId As Integer = 27
            Dim category As Short = 3
            Dim rawData As Byte() = Nothing
            Dim currentSource As String = Source(sourceName)

            Dim registryPath As String = $"LocalMachine\SYSTEM\CurrentControlSet\Services\EventLog\{SmartEventLogger.LogName}\{currentSource}"

            mockReader.SetValue(registryPath, currentSource)

            ' Act
            Log(logName, sourceName, message, eventType, eventId, category, rawData)

            ' Assert
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the creates event source when log and source does not provided.
        ''' </summary>
        <Fact>
        Public Sub Log_Message_Length_Is_Shortened_WhenToLong()

            ' Arrange
#If NET35 Then
            Dim testReader As New TestRegistryReader()
            Dim testWriter As New TestEventLogWriter()
#Else
            Dim testReader As New TestRegistryReader(OutputHelper)
            Dim testWriter As New TestEventLogWriter(OutputHelper)
#End If
            SetRegistryReader(testReader)
            SetWriter(testWriter)

            Dim logName As String = ""
            Dim sourceName As New String("s"c, 300) ' 300 characters long
            Dim message As New String("X"c, 40000) ' 40,000 characters long
            Dim eventType As EventLogEntryType = EventLogEntryType.FailureAudit
            Dim eventId As Integer = 27
            Dim category As Short = 3
            Dim rawData As Byte() = Nothing
            Dim currentSource As String = Source(sourceName)

            Dim registryPath As String = $"LocalMachine\SYSTEM\CurrentControlSet\Services\EventLog\{SmartEventLogger.LogName}\{currentSource}"

            testReader.SetValue(registryPath, currentSource)

            ' Act
            Log(logName, sourceName, message, eventType, eventId, category, rawData)

            ' Assert
            Assert.False(testWriter.CreateEventSourceCalled)
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal(32766, testWriter.MessageLength) ' Maximum length for event log message
            Assert.Equal(211, testWriter.SourceLength) ' Source is truncated to 254 max characters

        End Sub

        ''' <summary>
        ''' Logs the creates event source when log and source does not provided.
        ''' </summary>
        <Fact>
        Public Sub Log_Message_Multiple_Entries_WhenToLong()

            ' Arrange
#If NET35 Then
            Dim testWriter As New TestEventLogWriter()
            Dim testReader As New TestRegistryReader()
#Else
            Dim testWriter As New TestEventLogWriter(OutputHelper)
            Dim testReader As New TestRegistryReader(OutputHelper)
#End If
            SetWriter(testWriter)
            SetRegistryReader(testReader)

            Dim logName As String = ""
            Dim sourceName As New String("s"c, 300) ' 300 characters long
            Dim message As New String("X"c, 40000) ' 40,000 characters long
            Dim eventType As EventLogEntryType = EventLogEntryType.FailureAudit
            Dim eventId As Integer = 27
            Dim category As Short = 3
            Dim rawData As Byte() = Nothing
            AllowMultiEntryMessages = True
            ContinuationMarker = String.Empty
            Dim currentSource As String = Source(sourceName)

            Dim registryPath As String = $"LocalMachine\SYSTEM\CurrentControlSet\Services\EventLog\{SmartEventLogger.LogName}\{currentSource}"

            testReader.SetValue(registryPath, currentSource)

            ' Act
            Log(logName, sourceName, message, eventType, eventId, category, rawData)

            AllowMultiEntryMessages = False
            ' Assert
            Assert.False(testWriter.CreateEventSourceCalled)
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal(7691, testWriter.MessageLength) ' Maximum length for event log message
            Assert.Equal(211, testWriter.SourceLength) ' Source is truncated to 254 max characters

        End Sub

        ''' <summary>
        ''' Logs the creates event source when log and source does not provided.
        ''' </summary>
        <Fact>
        Public Sub Fluent_Log_Message_Length_Is_Shortened_WhenToLong()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter()
#Else
            testWriter = New TestEventLogWriter(OutputHelper)
#End If
            SetWriter(testWriter)

            Dim logName As String = ""
            Dim sourceName As New String("s"c, 300) ' 300 characters long
            Dim message As New String("X"c, 40000) ' 40,000 characters long
            Dim eventType As EventLogEntryType = EventLogEntryType.FailureAudit
            Dim eventId As Integer = 27
            Dim category As Short = 3
            Dim rawData As Byte() = Nothing
            AllowMultiEntryMessages = False

            ' Act
            GetLog(logName, sourceName).LogEntry(message, eventType, eventId, category, rawData)

            ' Assert
            Assert.False(testWriter.CreateEventSourceCalled)
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal(32766, testWriter.MessageLength) ' Maximum length for event log message
            Assert.Equal(211, testWriter.SourceLength) ' Source is truncated to 254 max characters

        End Sub

        ''' <summary>
        ''' Fluents the log message.
        ''' </summary>
        <Theory>
        <InlineData(LoggingSeverity.Critical, True)>
        <InlineData(LoggingSeverity.Error, True)>
        <InlineData(LoggingSeverity.Info, True)>
        <InlineData(LoggingSeverity.Warning, True)>
        <InlineData(LoggingSeverity.Verbose, False)>
        Public Sub Fluent_Log_Message_Severity(
                ByVal Severity As LoggingSeverity,
                ByVal WriteEntry As Boolean)

            ' Arrange
            Dim ll As LoggingLevel = CurrentLoggingLevel
            CurrentLoggingLevel = LoggingLevel.Normal
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter()
#Else
            testWriter = New TestEventLogWriter(OutputHelper)
#End If
            SetWriter(testWriter)

            Dim logName As String = ""
            Dim sourceName As New String("s"c, 300) ' 300 characters long
            Dim message As New String("X"c, 40000) ' 40,000 characters long
            logName = "TestLog"
            AllowMultiEntryMessages = False

            ' Act
            GetLog(logName, sourceName).LogEntry(message, Severity)

            ' Assert
            CurrentLoggingLevel = ll ' Restore original logging level
            Assert.False(testWriter.CreateEventSourceCalled)
            Assert.Equal(WriteEntry, testWriter.WriteEntryCalled)
            Assert.Equal(If(WriteEntry, 32766, 0), testWriter.MessageLength) ' Maximum length for event log message
            Assert.Equal(211, testWriter.SourceLength) ' Source is truncated to 211 max characters

        End Sub

        ''' <summary>
        ''' Fluents the log message.
        ''' </summary>
        <Fact>
        Public Sub Fluent_Log_Message()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter()
#Else
            testWriter = New TestEventLogWriter(OutputHelper)
#End If
            SetWriter(testWriter)

            Dim logName As String = ""
            Dim sourceName As New String("s"c, 300) ' 300 characters long
            Dim message As New String("X"c, 40000) ' 40,000 characters long
            logName = "TestLog"
            AllowMultiEntryMessages = False

            ' Act
            GetLog(logName, sourceName).LogEntry(message)

            ' Assert
            Assert.False(testWriter.CreateEventSourceCalled)
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal(32766, testWriter.MessageLength) ' Maximum length for event log message
            Assert.Equal(211, testWriter.SourceLength) ' Source is truncated to 254 max characters

        End Sub

        ''' <summary>
        ''' Fluents the log message event identifier.
        ''' </summary>
        <Fact>
        Public Sub Fluent_Log_Message_EventId_EntrySeverity()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter()
#Else
            testWriter = New TestEventLogWriter(OutputHelper)
#End If
            SetWriter(testWriter)

            Dim logName As String = ""
            Dim sourceName As New String("s"c, 300) ' 300 characters long
            Dim message As New String("X"c, 40000) ' 40,000 characters long
            Dim eventId As Integer = 42
            logName = "TestLog"
            AllowMultiEntryMessages = False

            ' Act
            GetLog(logName, sourceName).LogEntry(message, eventId, LoggingSeverity.Error)

            ' Assert
            Assert.False(testWriter.CreateEventSourceCalled)
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal(32766, testWriter.MessageLength) ' Maximum length for event log message
            Assert.Equal(211, testWriter.SourceLength) ' Source is truncated to 254 max characters

        End Sub

        ''' <summary>
        ''' Fluents the log message event identifier.
        ''' </summary>
        <Fact>
        Public Sub Fluent_Log_Message_EventId()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter()
#Else
            testWriter = New TestEventLogWriter(OutputHelper)
#End If
            SetWriter(testWriter)

            Dim logName As String = ""
            Dim sourceName As New String("s"c, 300) ' 300 characters long
            Dim message As New String("X"c, 40000) ' 40,000 characters long
            Dim eventId As Integer = 42
            logName = "TestLog"
            AllowMultiEntryMessages = False

            ' Act
            GetLog(logName, sourceName).LogEntry(message, eventId)

            ' Assert
            Assert.False(testWriter.CreateEventSourceCalled)
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal(32766, testWriter.MessageLength) ' Maximum length for event log message
            Assert.Equal(211, testWriter.SourceLength) ' Source is truncated to 254 max characters

        End Sub

        ''' <summary>
        ''' Fluents the log message category.
        ''' </summary>
        <Fact>
        Public Sub Fluent_Log_Message_Category_EntrySeverity()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter()
#Else
            testWriter = New TestEventLogWriter(OutputHelper)
#End If
            SetWriter(testWriter)

            Dim logName As String = ""
            Dim sourceName As New String("s"c, 300) ' 300 characters long
            Dim message As New String("X"c, 40000) ' 40,000 characters long
            Dim category As Short = 3
            logName = "TestLog"
            AllowMultiEntryMessages = False

            ' Act
            GetLog(logName, sourceName).LogEntry(message, category, LoggingSeverity.Error)

            ' Assert
            Assert.False(testWriter.CreateEventSourceCalled)
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal(32766, testWriter.MessageLength) ' Maximum length for event log message
            Assert.Equal(211, testWriter.SourceLength) ' Source is truncated to 254 max characters

        End Sub

        ''' <summary>
        ''' Fluents the log message category.
        ''' </summary>
        <Fact>
        Public Sub Fluent_Log_Message_Category()

            ' Arrange
#If NET35 Then
            Dim testWriter As New TestEventLogWriter()
#Else
            Dim testWriter As New TestEventLogWriter(OutputHelper)
#End If
            SetWriter(testWriter)

            Dim logName As String = ""
            Dim sourceName As New String("s"c, 300) ' 300 characters long
            Dim message As New String("X"c, 40000) ' 40,000 characters long
            Dim category As Short = 3
            logName = "TestLog"

            ' Act
            GetLog(logName, sourceName).LogEntry(message, category)

            ' Assert
            Assert.False(testWriter.CreateEventSourceCalled)
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal(32766, testWriter.MessageLength) ' Maximum length for event log message
            Assert.Equal(211, testWriter.SourceLength) ' Source is truncated to 254 max characters

        End Sub

        ''' <summary>
        ''' Fluents the log message raw data entry severity.
        ''' </summary>
        <Fact>
        Public Sub Fluent_Log_Message_RawData_EntrySeverity()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter()
#Else
            testWriter = New TestEventLogWriter(OutputHelper)
#End If
            SetWriter(testWriter)

            Dim logName As String = ""
            Dim sourceName As New String("s"c, 300) ' 300 characters long
            Dim message As New String("X"c, 40000) ' 40,000 characters long
            Dim rawData As Byte() = Nothing
            logName = "TestLog"
            AllowMultiEntryMessages = False

            ' Act
            GetLog(logName, sourceName).LogEntry(message, rawData, LoggingSeverity.Error)

            ' Assert
            Assert.False(testWriter.CreateEventSourceCalled)
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal(32766, testWriter.MessageLength) ' Maximum length for event log message
            Assert.Equal(211, testWriter.SourceLength) ' Source is truncated to 254 max characters

        End Sub

        ''' <summary>
        ''' Fluents the log message raw data.
        ''' </summary>
        <Fact>
        Public Sub Fluent_Log_Message_RawData()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter()
#Else
            testWriter = New TestEventLogWriter(OutputHelper)
#End If
            SetWriter(testWriter)

            Dim logName As String = ""
            Dim sourceName As New String("s"c, 300) ' 300 characters long
            Dim message As New String("X"c, 40000) ' 40,000 characters long
            Dim rawData As Byte() = Nothing
            logName = "TestLog"
            AllowMultiEntryMessages = False

            ' Act
            GetLog(logName, sourceName).LogEntry(message, rawData)

            ' Assert
            Assert.False(testWriter.CreateEventSourceCalled)
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal(32766, testWriter.MessageLength) ' Maximum length for event log message
            Assert.Equal(211, testWriter.SourceLength) ' Source is truncated to 254 max characters

        End Sub

        ''' <summary>
        ''' Fluents the log long message.
        ''' </summary>
        <Fact>
        Public Sub Fluent_Log_LongMessage()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter()
#Else
            testWriter = New TestEventLogWriter(OutputHelper)
#End If
            SetWriter(testWriter)

            Dim logName As String = ""
            Dim sourceName As New String("s"c, 300) ' 300 characters long
            Dim message As New String("X"c, 40000) ' 40,000 characters long
            Dim rawData As Byte() = Nothing
            logName = "TestLog"
            AllowMultiEntryMessages = True
            ContinuationMarker = " ... <Continued>"

            ' Act
            GetLog(logName, sourceName).LogEntry(message, rawData)

            AllowMultiEntryMessages = False
            ' Assert
            Assert.False(testWriter.CreateEventSourceCalled)
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal(7703, testWriter.MessageLength) ' Maximum length for event log message
            Assert.Equal(211, testWriter.SourceLength) ' Source is truncated to 254 max characters

        End Sub

        ''' <summary>
        ''' Defaults the source returns expected result when source exists.
        ''' </summary>
        <Fact>
        Public Sub DefaultSource_ReturnsExpectedResult_WhenSourceExists()

            ' Arrange
            Dim OldMachine As String = MachineName
            Dim expectedSource As String = "ExpectedSource"

#If NET35 Then
            Dim mockReader As New TestRegistryReader()
#Else
            Dim mockReader As New TestRegistryReader(OutputHelper)
#End If
            mockReader.DefaultEventLogSourceResult = expectedSource

            SetRegistryReader(mockReader)
            MachineName = "SomeOtherMachine"

            ' Act
            Dim result As String = DefaultSource("CustomLog")

            ' Assert
            Assert.Equal(expectedSource, result)
            MachineName = OldMachine

        End Sub

        ''' <summary>
        ''' Maximums the source returns proper source when given blank source.
        ''' </summary>
        <Fact>
        Public Sub MaxSource_ReturnsProperSource_WhenGivenBlank_Source()

            ' Arrange
            SourceName = String.Empty

            ' Act
            Dim result As String = MaxSource(String.Empty)
            Dim Initialized As Boolean = IsInitialized
            ' Assert
            Output($"result = {result}")
            Assert.True(Initialized)
#If NET35 Then
            Assert.True(result.StartsWith("EventLogHelper.Tests.net"))
            Assert.True(result.EndsWith(".SmartEventLoggerTest.MaxSource_ReturnsProperSource_WhenGivenBlank_Source"))
#Else
            Assert.StartsWith("EventLogHelper.Tests.net", result)
            Assert.EndsWith(".SmartEventLoggerTest.MaxSource_ReturnsProperSource_WhenGivenBlank_Source", result)
#End If

        End Sub

        ''' <summary>
        ''' Maximums the source returns given source when given blank source.
        ''' </summary>
        <Fact>
        Public Sub MaxSource_ReturnsGivenSource_WhenGivenBlank_Source()

            ' Arrange
            Dim OldSource As String = SourceName
            SourceName = "CustomSource"
            ' Act
            Dim result As String = MaxSource(String.Empty)

            ' Assert
            Output($"result = {result}")
#If NET35 Then
            Assert.True(result.Equals("CustomSource"))
#Else
            Assert.Equal("CustomSource", result)
#End If
            SourceName = OldSource

        End Sub

        ''' <summary>
        ''' Determines whether is source register to log.
        ''' </summary>
        <Fact>
        Public Sub IsSourceRegisterToLog()

            ' Arrange
#If NET35 Then
            Dim mockReader As New TestRegistryReader()
#Else
            Dim mockReader As New TestRegistryReader(OutputHelper)
#End If
            SetRegistryReader(mockReader)

            Dim CurrentSource As String = "CustomSource"
            Dim CurrentLog As String = "CustomLog"
            Dim CurrentMachine As String = "SomeMachineSomeplace"

            Dim registryPath As String = $"LocalMachine\SYSTEM\CurrentControlSet\Services\EventLog\{CurrentLog}\{CurrentSource}"

            mockReader.SetValue(registryPath, CurrentSource)

            ' Act
            Dim result As Boolean = IsSourceRegisteredToLog(CurrentSource, CurrentLog, CurrentMachine)

            ' Assert
            Output($"result = {result}")
            Assert.True(result)

        End Sub

        ''' <summary>
        ''' Determines whether this instance can read registry for log and source test.
        ''' </summary>
        <Fact>
        Public Sub CanReadRegistryForLogAndSourceTest()

            ' Arrange
#If NET35 Then
            Dim mockReader As New TestRegistryReader()
#Else
            Dim mockReader As New TestRegistryReader(OutputHelper)
#End If
            SmartEventLogger.SetRegistryReader(mockReader)

            Dim CurrentSource As String = "CustomSource"
            Dim CurrentLog As String = "CustomLog"
            Dim CurrentMachine As String = "SomeMachineSomeplace"

            Dim registryPath As String = $"LocalMachine\SYSTEM\CurrentControlSet\Services\EventLog\{CurrentLog}\{CurrentSource}"

            mockReader.SetValue(registryPath, CurrentSource)

            ' Act
            Dim result As Boolean = CanReadRegistryForLogAndSource(CurrentSource, CurrentLog, CurrentMachine)

            ' Assert
            Output($"result = {result}")
            Assert.True(result)

        End Sub

        ''' <summary>
        ''' Determines whether this instance can read registry for log test.
        ''' </summary>
        <Fact>
        Public Sub CanReadRegistryForLogTest()

            ' Arrange
#If NET35 Then
            Dim mockReader As New TestRegistryReader()
#Else
            Dim mockReader As New TestRegistryReader(OutputHelper)
#End If
            SetRegistryReader(mockReader)

            Dim CurrentLog As String = "CustomLog"
            Dim CurrentMachine As String = "SomeMachineSomeplace"

            Dim registryPath As String = $"LocalMachine\SYSTEM\CurrentControlSet\Services\EventLog\{CurrentLog}"

            mockReader.SetValue(registryPath, CurrentLog)

            ' Act
            Dim result As Boolean = CanReadRegistryForLog(CurrentLog, CurrentMachine)

            ' Assert
            Output($"result = {result}")
            Assert.True(result)

        End Sub

        ''' <summary>
        ''' test of GetLog.
        ''' </summary>
        <Fact>
        Public Sub GetLogTest()

            ' Arrange
#If NET35 Then
            Dim testWriter As New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            Dim testWriter As New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            InitializeConfiguration()

            MachineName = "SomeMachineSomeplace"
            LogName = "CustomLog"
            SourceName = "CustomSource"
            MaxKilobytes = 1024 * 1024 ' 1 MB
            RetentionDays = 7
            WriteInitEntry = True

            ' Act
            Dim result As EventLog = GetLog()

            ' Assert
            Output($"result = {(result.GetType())}")
            Assert.NotNull(result)

        End Sub

        ''' <summary>
        ''' Gets the application setting fails with type not enum.
        ''' </summary>
        <Fact>
        Public Sub GetAppSetting_FailsWithTypeNotEnum()

            ' Arrange
            Dim value As Integer = 42

            ' Act & Assert
            Dim ex As ArgumentException =
                Assert.Throws(Of ArgumentException)(
                    Sub()
                        Dim result As Integer = GetAppSetting("TestSetting", value)
                    End Sub)

            Assert.Contains("must be an Enum type", ex.Message)

        End Sub

        ''' <summary>
        ''' Gets the application setting fails with type not enum.
        ''' </summary>
        <Fact>
        Public Sub GetAppSetting_ReturnsDefaultWhenInvaid()

            ' Arrange
            Dim value As LoggingSeverity = LoggingSeverity.Critical

            ' Act 
            Dim result As LoggingSeverity = GetAppSetting("Test.InvalidItem", value)

            ' Assert
            Assert.Equal(value, result)

        End Sub

#If NET5_0_OR_GREATER Then
        ''' <summary>
        ''' Gets the application setting fails with type not enum.
        ''' </summary>
        <Fact>
        Public Sub GetAppSetting_ReturnsRightValueWhenValid()

            ' Arrange
            Dim value As LoggingSeverity = LoggingSeverity.Critical

            ' Act 
            Dim result As LoggingSeverity = GetAppSetting("Test.Error", value)

            ' Assert
            Assert.Equal(LoggingSeverity.Error, result)

        End Sub
#End If

#If NET5_0_OR_GREATER Then
        ''' <summary>
        ''' Gets the application setting fails with type not enum.
        ''' </summary>
        <Fact>
        Public Sub GetAppSetting_triggersAnOverFlowException()

            ' Arrange
            Dim value As LoggingSeverity = LoggingSeverity.Critical

            ' Act & Assert
            Dim result As LoggingSeverity = GetAppSetting("Test.Overflow", value)

            Assert.Equal(value, result)

        End Sub
#End If

#If NET5_0_OR_GREATER Then
        <Fact>
        Public Sub GetAppSetting_ReadsValuesFromJson()

            ' Arrange

            ' Act
            InitializeConfiguration()

            ' Assert
            Assert.Equal("TestLog", LogName)
            Assert.Equal(42, RetentionDays)
            Assert.True(WriteInitEntry)

        End Sub
#End If

        ''' <summary>
        ''' Logs the exists and source exist.
        ''' </summary>
        <Fact>
        Public Sub LogExists_And_SourceExist_ButNotTogether_UseSourceLog()

            ' Arrange
#If NET35 Then
            Dim testWriter As New TestEventLogWriter(logExists:=True, sourceExists:=True)
            Dim mockReader As New TestRegistryReader()
#Else
            Dim testWriter As New TestEventLogWriter(OutputHelper, logExists:=True, sourceExists:=True)
            Dim mockReader As New TestRegistryReader(OutputHelper)
#End If
            SetWriter(testWriter)
            SetRegistryReader(mockReader)

            Dim logName As String = ""
            Dim sourceName As String = ""
            Dim message As String = ""
            Dim eventType As EventLogEntryType = EventLogEntryType.FailureAudit
            Dim eventId As Integer = 27
            Dim category As Short = 3
            Dim rawData As Byte() = Nothing

            SmartEventLogger.SourceResolutionBehavior = SourceResolutionBehavior.UseSourceLog

            ' Act
            Log(logName, sourceName, message, eventType, eventId, category, rawData)

            ' Assert
            Assert.False(testWriter.CreateEventSourceCalled)
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the exists and source exist.
        ''' </summary>
        <Fact>
        Public Sub LogExists_And_SourceExist_ButNotTogether_UseLogDefaultSource()

            ' Arrange
#If NET35 Then
            Dim testWriter As New TestEventLogWriter(logExists:=True, sourceExists:=True)
            Dim mockReader As New TestRegistryReader()
#Else
            Dim testWriter As New TestEventLogWriter(OutputHelper, logExists:=True, sourceExists:=True)
            Dim mockReader As New TestRegistryReader(OutputHelper)
#End If
            SetWriter(testWriter)
            SetRegistryReader(mockReader)

            Dim logName As String = ""
            Dim sourceName As String = ""
            Dim message As String = ""
            Dim eventType As EventLogEntryType = EventLogEntryType.FailureAudit
            Dim eventId As Integer = 27
            Dim category As Short = 3
            Dim rawData As Byte() = Nothing

            SmartEventLogger.SourceResolutionBehavior = SourceResolutionBehavior.UseLogsDefaultSource
            mockReader.DefaultEventLogSourceResult = "DefaultSource"

            ' Act
            Log(logName, sourceName, message, eventType, eventId, category, rawData)

            ' Assert
            Assert.False(testWriter.CreateEventSourceCalled)
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

#If NETCOREAPP Then
        ''' <summary>
        ''' Returns the value from application settings when no application settings json.
        ''' </summary>
        <Fact>
        Public Sub ReturnsValueFromAppSettings_WhenNoAppSettingsJson()

            ' Arrange: Pretend no appsettings.json file
            FileSystem = New FakeFileSystem(Array.Empty(Of String)())

            ' Act
            Dim result As String = GetAppSetting("SomeKey", "DefaultValue")

            ' Assert
            Assert.Equal("DefaultValue", result)

        End Sub

        ''' <summary>
        ''' Uses the application settings json when file exists.
        ''' </summary>
        <Fact>
        Public Sub UsesAppSettingsJson_WhenFileExists()

            ' Arrange: Pretend appsettings.json exists
            Dim fakePath As String = IO.Path.Combine(AppContext.BaseDirectory, "appsettings.json")
            FileSystem = New FakeFileSystem({fakePath})

            ' Act
            Dim result As String = GetAppSetting("SomeKey", "DefaultValue")

            ' Assert
            ' (you may need to also stub config here — this just ensures the path check hits)
            Assert.NotNull(result)

        End Sub

        ''' <summary>
        ''' Uses the custom json file when overridden.
        ''' </summary>
        <Fact>
        Public Sub UsesCustomJsonFile_WhenOverridden()

            ' Arrange
            JsonSettingsFile = "customsettings.json"

            ' Act
            Dim result As String = GetAppSetting("SomeKey", "DefaultValue")

            ' Assert
            Assert.NotNull(result)

        End Sub

        ''' <summary>
        ''' Falls back to default when configuration throws.
        ''' </summary>
        <Fact>
        Public Sub FallsBackToDefault_WhenConfigThrows()

            ' Arrange
            FileSystem = New FakeFileSystem({}) ' no json file
            Config = New ThrowingConfigShim()

            ' Act
            Dim result As String = GetAppSetting("SomeKey", "DefaultValue")

            ' Assert
            Assert.Equal("DefaultValue", result)

        End Sub
#End If

#End Region

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

End Namespace
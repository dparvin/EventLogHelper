Imports Xunit
Imports System.ComponentModel

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
    ''' Test class for <see cref="SmartEventLogger"/>.
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
#Else
        Private OutputHelper As ITestOutputHelper

        ''' <summary>
        ''' Initializes a new instance of the <see cref="SmartEventLoggerTest"/> class.
        ''' </summary>
        ''' <param name="output">The output.</param>
        Sub New(ByVal output As ITestOutputHelper)

            OutputHelper = output

        End Sub
#End If

#Region " Test Methods ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ "

        ''' <summary>
        ''' Logs the message.
        ''' </summary>
        <Fact>
        Public Sub Log_Message()

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=False, sourceExists:=False)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=False, sourceExists:=False)
#End If
            SetWriter(testWriter)

            Dim message As String = "Test message"

            ' Act
            Log(message)

            ' Assert
            Assert.True(testWriter.CreateEventSourceCalled)
            Assert.True(testWriter.WriteEntryCalled)

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
            Assert.True(testWriter.CreateEventSourceCalled)
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
            Assert.True(testWriter.CreateEventSourceCalled)
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal($"[{sourceName}] {message}.", testWriter.LastMessage)

        End Sub

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
            Assert.True(testWriter.CreateEventSourceCalled)
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
            Assert.True(testWriter.CreateEventSourceCalled)
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
            Assert.True(testWriter.CreateEventSourceCalled)
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal($"[{sourceName}] {message}.", testWriter.LastMessage)

        End Sub

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
            Assert.True(testWriter.CreateEventSourceCalled)
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
            Assert.True(testWriter.CreateEventSourceCalled)
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
            Assert.True(testWriter.CreateEventSourceCalled)
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
            Assert.True(testWriter.CreateEventSourceCalled)
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
            Assert.True(testWriter.CreateEventSourceCalled)
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal($"[{sourceName}] {message}.", testWriter.LastMessage)

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
            Assert.True(testWriter.CreateEventSourceCalled)
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
            Assert.True(testWriter.CreateEventSourceCalled)
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
            Assert.True(testWriter.CreateEventSourceCalled)
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
            Assert.True(testWriter.CreateEventSourceCalled)
            Assert.True(testWriter.WriteEntryCalled)

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

            Dim sourceName As String = "CustomSource"
            Dim message As String = "Test message"
            Dim eventType As EventLogEntryType = EventLogEntryType.Warning
            Dim eventId As Integer = 42
            Dim category As Short = 1
            Dim rawData As Byte() = Nothing

            ' Act
            Log(sourceName, message, eventType, eventId, category, rawData)

            ' Assert
            Assert.True(testWriter.CreateEventSourceCalled)
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
            Assert.True(testWriter.CreateEventSourceCalled)
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
            Assert.True(testWriter.CreateEventSourceCalled)
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
        Public Sub Log_Exists_And_SourceExist(logExists As Boolean, sourceExists As Boolean)

            ' Arrange
            Dim testWriter As TestEventLogWriter
#If NET35 Then
            testWriter = New TestEventLogWriter(logExists:=logExists, sourceExists:=sourceExists)
#Else
            testWriter = New TestEventLogWriter(OutputHelper, logExists:=logExists, sourceExists:=sourceExists)
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
            If logExists AndAlso sourceExists Then
                Assert.False(testWriter.CreateEventSourceCalled)
            Else
                Assert.True(testWriter.CreateEventSourceCalled)
            End If
            Assert.True(testWriter.WriteEntryCalled)

        End Sub

        ''' <summary>
        ''' Logs the creates event source when log and source does not provided.
        ''' </summary>
        <Fact>
        Public Sub Log_Message_Length_Is_Shortened_WhenToLong()

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

            ' Act
            Log(logName, sourceName, message, eventType, eventId, category, rawData)

            ' Assert
            Assert.False(testWriter.CreateEventSourceCalled)
            Assert.True(testWriter.WriteEntryCalled)
            Assert.Equal(32766, testWriter.MessageLength) ' Maximum length for event log message
            Assert.Equal(254, testWriter.SourceLength) ' Source is truncated to 254 max characters

        End Sub


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
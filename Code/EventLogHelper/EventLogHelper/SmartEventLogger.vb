Imports System.Reflection

''' <summary>
''' Provides high-level utilities for logging to the Windows Event Log in a simplified and customizable way.
''' </summary>
''' <remarks>
''' <para>
''' This module abstracts common patterns for writing entries to the Windows Event Log. It simplifies usage by handling:
''' </para>
''' <list type="bullet">
'''   <item>Automatic validation and creation of logs and sources.</item>
'''   <item>Falls back when <c>log</c> or <c>source</c> parameters are not provided.</item>
'''   <item>Automatic source naming based on the calling method and class context.</item>
'''   <item>Safe handling and truncation of long messages to comply with system limits.</item>
''' </list>
''' <para>
''' The default event log is <c>Application</c>. If no source is provided, a fully qualified method name
''' (including namespace and class) is used as the source identifier.
''' </para>
''' <para>
''' This module is suitable for both production use (via <see cref="RealEventLogWriter"/>) and unit testing 
''' (via a test implementation such as <c>EventLogHelper.Tests.TestEventLogWriter</c> in a separate test project).
''' </para>
''' </remarks>
Public Module SmartEventLogger

    ''' <summary>
    ''' The current event log writer instance used by the logger.
    ''' </summary>
    ''' <remarks>
    ''' By default, this is initialized to an instance of <see cref="RealEventLogWriter"/>, which writes to the actual
    ''' Windows Event Log. This field is typically replaced during testing with a mock or stub implementation via the
    ''' <see cref="SetWriter"/> method.
    ''' </remarks>
    Private _writer As IEventLogWriter = New RealEventLogWriter()

    ''' <summary>
    ''' Sets the <see cref="IEventLogWriter"/> implementation used for logging operations.
    ''' </summary>
    ''' <param name="writer">
    ''' The custom writer instance to use. This allows swapping out the default <see cref="RealEventLogWriter"/>
    ''' for testing or alternate logging implementations.
    ''' </param>
    ''' <remarks>
    ''' This method is primarily intended for use in unit tests or advanced scenarios where log output should be redirected
    ''' or suppressed. For example, a test implementation may capture and inspect log data without writing to the real event log.
    ''' </remarks>
    Public Sub SetWriter(writer As IEventLogWriter)
        _writer = writer
    End Sub

#Region " Log Methods ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ "

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String)

        SmartEventLogger.Log("", "", message)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="source">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal source As String,
            ByVal message As String)

        SmartEventLogger.Log("", source, message)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType)

        SmartEventLogger.Log("", "", message, eventType)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="log">
    ''' The name of the event log to write to (e.g., "Application", "System", or a custom log name).
    ''' If <c>null</c> or empty, "Application" is used by default.
    ''' </param>
    ''' <param name="source">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal log As String,
            ByVal source As String,
            ByVal message As String)

        SmartEventLogger.Log(log, source, message, EventLogEntryType.Information)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="source">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal source As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType)

        SmartEventLogger.Log("", source, message, eventType)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="log">
    ''' The name of the event log to write to (e.g., "Application", "System", or a custom log name).
    ''' If <c>null</c> or empty, "Application" is used by default.
    ''' </param>
    ''' <param name="source">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal log As String,
            ByVal source As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType)

        SmartEventLogger.Log(log, source, message, eventType, 0)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer)

        SmartEventLogger.Log("", message, eventType, eventID)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="source">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal source As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer)

        SmartEventLogger.Log("", source, message, eventType, eventID)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="log">
    ''' The name of the event log to write to (e.g., "Application", "System", or a custom log name).
    ''' If <c>null</c> or empty, "Application" is used by default.
    ''' </param>
    ''' <param name="source">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal log As String,
            ByVal source As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer)

        SmartEventLogger.Log(log, source, message, eventType, eventID, 0)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short)

        SmartEventLogger.Log("", message, eventType, eventID, category)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="source">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal source As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short)

        SmartEventLogger.Log("", source, message, eventType, eventID, category)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="log">
    ''' The name of the event log to write to (e.g., "Application", "System", or a custom log name).
    ''' If <c>null</c> or empty, "Application" is used by default.
    ''' </param>
    ''' <param name="source">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal log As String,
            ByVal source As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short)

        SmartEventLogger.Log(log, source, message, eventType, eventID, category, Nothing)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal rawData As Byte())

        SmartEventLogger.Log(message, EventLogEntryType.Information, rawData)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal rawData As Byte())

        SmartEventLogger.Log(message, eventType, 0, rawData)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal rawData As Byte())

        SmartEventLogger.Log(message, eventType, eventID, 0, rawData)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByVal rawData As Byte())

        SmartEventLogger.Log("", "", message, eventType, eventID, category, rawData)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="source">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal source As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByVal rawData As Byte())

        SmartEventLogger.Log("", source, message, eventType, eventID, category, rawData)

    End Sub

    ''' <summary>
    ''' Writes an entry to the Windows Event Log using the provided parameters.
    ''' </summary>
    ''' <param name="log">
    ''' The name of the event log to write to (e.g., "Application", "System", or a custom log name).
    ''' If <c>null</c> or empty, "Application" is used by default.
    ''' </param>
    ''' <param name="source">
    ''' The source of the log entry. If <c>null</c> or empty, a default source is generated using the calling method's
    ''' namespace, class, and method name. If the source exceeds 254 characters, it is automatically truncated.
    ''' </param>
    ''' <param name="message">
    ''' The message to log. If <c>null</c> or empty, a default message is used. Messages longer than 32,766 characters
    ''' are truncated and suffixed with "… &lt;truncated&gt;" unless multi-entry logging is enabled.
    ''' </param>
    ''' <param name="eventType">
    ''' The type of event to log. Valid values are <see cref="EventLogEntryType.Error" />,
    ''' <see cref="EventLogEntryType.Warning" />, and <see cref="EventLogEntryType.Information" />.
    ''' If an invalid value is passed, the event is logged as <c>Information</c>.
    ''' </param>
    ''' <param name="eventID">
    ''' The numeric event identifier for the log entry. This is application-defined and may be used to group similar events.
    ''' </param>
    ''' <param name="category">
    ''' The category for the event, if applicable. This is typically used in categorized event views.
    ''' </param>
    ''' <param name="rawData">
    ''' Optional binary data to include with the log entry. Can be <c>null</c> if not used.
    ''' </param>
    ''' <remarks>
    ''' If the specified log or source does not exist, they are created automatically. The message is automatically wrapped
    ''' in square brackets with the source name if it does not already include them, and ends with a period.
    ''' </remarks>
    Public Sub Log(
            ByVal log As String,
            ByVal source As String,
            ByVal message As String,
            ByVal eventType As EventLogEntryType,
            ByVal eventID As Integer,
            ByVal category As Short,
            ByVal rawData As Byte())

        Dim defaultLog As String = If(String.IsNullOrEmpty(log), "Application", log.Trim())
        Dim defaultSource As String = GetDefaultEventSourceName(source)

        If Not _writer.Exists(defaultLog) OrElse Not _writer.SourceExists(defaultSource, defaultLog) Then
            _writer.CreateEventSource(defaultSource, defaultLog)
        End If
        Dim finalMessage As String = If(String.IsNullOrEmpty(message), "No message provided.", message.Trim())
        If Not finalMessage.EndsWith("."c) Then finalMessage &= "."
        If Not finalMessage.StartsWith("["c) Then finalMessage = $"[{defaultSource}] {finalMessage}"

        Const maxLength As Integer = 32766
        Const truncationMarker As String = "... [TRUNCATED]"

        If finalMessage.Length > maxLength Then
            ' Leave room for the truncation marker
            Dim maxWithoutMarker As Integer = maxLength - truncationMarker.Length
            finalMessage = $"{finalMessage.Substring(0, maxWithoutMarker)}{truncationMarker}"
        End If
        Select Case eventType
            Case EventLogEntryType.Error, EventLogEntryType.Warning, EventLogEntryType.Information
                ' Valid
            Case Else
                eventType = EventLogEntryType.Information
        End Select
        _writer.WriteEntry(defaultLog, defaultSource, finalMessage, eventType, eventID, category, rawData)

    End Sub

#End Region

    ''' <summary>
    ''' Gets the default name of the event source.
    ''' </summary>
    ''' <returns></returns>
    Private Function GetDefaultEventSourceName(ByVal source As String) As String

        ' Final fallback if nothing else is available
        Dim strReturn As String = "SmartEventLogger"

        If String.IsNullOrEmpty(source) Then
            ' Try to get the calling assembly (e.g., the one calling your library)
            Dim thisAssembly As Assembly = Assembly.GetExecutingAssembly()

            Dim stackTrace As New StackTrace()
            For i As Integer = 1 To stackTrace.FrameCount - 1
                Dim frame As StackFrame = stackTrace.GetFrame(i)
                Dim method As MethodBase = frame.GetMethod()
                Dim type As Type = method.DeclaringType
                If type.Assembly IsNot thisAssembly Then
                    strReturn = $"{type.FullName}.{method.Name}"
                    Exit For
                End If
            Next
        Else
            strReturn = source.Trim()
        End If

        ' Windows event source names are limited to 254 characters
        If strReturn.Length > 254 Then strReturn = strReturn.Substring(0, 254)

        Return strReturn

    End Function

End Module

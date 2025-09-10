Namespace Enums

    ''' <summary>
    ''' Represents the severity of an individual log entry.
    ''' </summary>
    ''' <remarks>
    ''' <para>
    ''' Each log entry specifies a <see cref="LoggingSeverity"/> that reflects its 
    ''' importance or urgency. Whether the entry is actually written to the event log 
    ''' depends on the current <see cref="LoggingLevel"/>: only entries whose severity 
    ''' is greater than or equal to the active logging level are recorded.
    ''' </para>
    ''' 
    ''' <para>
    ''' The severity values are ordered from least to most important:
    ''' <list type="bullet">
    '''   <item><term>Verbose (0)</term><description>Detailed diagnostic or tracing information.</description></item>
    '''   <item><term>Info (1)</term><description>General informational messages about normal operation.</description></item>
    '''   <item><term>Warning (2)</term><description>Non-critical issues or unexpected conditions.</description></item>
    '''   <item><term>Error (3)</term><description>Failures or serious problems that need attention.</description></item>
    '''   <item><term>Critical (4)</term><description>Severe failures that require immediate action.</description></item>
    ''' </list>
    ''' </para>
    ''' </remarks>
    Public Enum LoggingSeverity
        ''' <summary>
        ''' Detailed diagnostic or tracing information (lowest importance).
        ''' </summary>
        Verbose = 0

        ''' <summary>
        ''' General informational messages about normal application flow.
        ''' </summary>
        Info = 1

        ''' <summary>
        ''' Non-critical issues, recoverable problems, or unexpected conditions.
        ''' </summary>
        Warning = 2

        ''' <summary>
        ''' Errors or failures that prevent normal operation.
        ''' </summary>
        [Error] = 3

        ''' <summary>
        ''' Severe failures requiring immediate attention (highest importance).
        ''' </summary>
        Critical = 4
    End Enum

End Namespace
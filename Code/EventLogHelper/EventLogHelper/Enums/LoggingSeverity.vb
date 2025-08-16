''' <summary>
''' Represents the severity levels for logging events.
''' This enumeration is used to categorize log messages based on their importance and urgency.
''' </summary>
''' <remarks>
''' The severity levels range from Verbose (least severe) to Critical (most severe).
''' Log entries with this enumeration will only be written to the event log when if 
''' the <see cref="LoggingLevel"/> is low enough for the item getting logged.
''' </remarks>
Public Enum LoggingSeverity
    ''' <summary>
    ''' The verbose
    ''' </summary>
    Verbose = 0
    ''' <summary>
    ''' The information
    ''' </summary>
    Info = 1
    ''' <summary>
    ''' The warning
    ''' </summary>
    Warning = 2
    ''' <summary>
    ''' The error
    ''' </summary>
    [Error] = 3
    ''' <summary>
    ''' The critical
    ''' </summary>
    Critical = 4
End Enum
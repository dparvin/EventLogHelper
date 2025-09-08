''' <summary>
''' Defines exit codes for the EventLogHelperCreate.exe utility.
''' </summary>
Friend Enum ElevatorExitCode
    ''' <summary>
    ''' The success
    ''' </summary>
    Success = 0
    ''' <summary>
    ''' The no arguments
    ''' </summary>
    NoArguments = 1
    ''' <summary>
    ''' The not elevated
    ''' </summary>
    NotElevated = 2
    ''' <summary>
    ''' The bad arguments
    ''' </summary>
    BadArguments = 3
    ''' <summary>
    ''' The error creating
    ''' </summary>
    ErrorCreating = 4
End Enum

Namespace Enums

    ''' <summary>
    ''' Specifies how to resolve situations where an event source is already registered
    ''' to a different log than the one requested. This determines whether logging
    ''' should fail, adapt, or create a new source.
    ''' </summary>
    Public Enum SourceResolutionBehavior

        ''' <summary>
        ''' Require an exact match. If the source is registered under a different log
        ''' than the one requested, logging fails with an error.
        ''' </summary>
        Strict

        ''' <summary>
        ''' If the source is registered under a different log than requested,
        ''' automatically switch to that log and proceed with logging.
        ''' </summary>
        UseSourceLog

        ''' <summary>
        ''' If the source is not registered to the specified log, use the log's default
        ''' source (the log name itself if available, otherwise the first registered source).
        ''' </summary>
        UseLogsDefaultSource
    End Enum

End Namespace
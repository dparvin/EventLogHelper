Namespace Interfaces

    ''' <summary>
    ''' Defines a contract for file system operations, allowing for abstraction and easier testing.
    ''' </summary>
    Public Interface IFileSystemShim

        ''' <summary>
        ''' Files the exists.
        ''' </summary>
        ''' <param name="path">The path.</param>
        ''' <returns></returns>
        Function FileExists(path As String) As Boolean

    End Interface

End Namespace

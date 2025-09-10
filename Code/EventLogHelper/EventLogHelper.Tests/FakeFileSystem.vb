Imports EventLogHelper.Interfaces

''' <summary>
''' A fake implementation of <see cref="IFileSystemShim"/> for testing purposes.
''' </summary>
''' <seealso cref="IFileSystemShim" />
Friend Class FakeFileSystem

    Implements IFileSystemShim

    ''' <summary>
    ''' The files
    ''' </summary>
    Private ReadOnly _files As HashSet(Of String)

    ''' <summary>
    ''' Initializes a new instance of the <see cref="FakeFileSystem"/> class.
    ''' </summary>
    ''' <param name="files">The files.</param>
    Public Sub New(files As IEnumerable(Of String))

        _files = New HashSet(Of String)(files, StringComparer.OrdinalIgnoreCase)

    End Sub

    ''' <summary>
    ''' Files the exists.
    ''' </summary>
    ''' <param name="path">The path.</param>
    ''' <returns></returns>
    Public Function FileExists(path As String) As Boolean Implements IFileSystemShim.FileExists

        Return _files.Contains(path)

    End Function

End Class

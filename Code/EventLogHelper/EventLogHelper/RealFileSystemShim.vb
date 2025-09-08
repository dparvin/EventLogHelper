''' <summary>
''' A real implementation of <see cref="EventLogHelper.IFileSystemShim"/> that uses the actual file system.
''' </summary>
''' <seealso cref="EventLogHelper.IFileSystemShim" />
#If NET35 Then
<Diagnostics.ExcludeFromCodeCoverage>
Friend Class RealFileSystemShim
#Else
<CodeAnalysis.ExcludeFromCodeCoverage>
Friend Class RealFileSystemShim
#End If

    Implements IFileSystemShim

    ''' <summary>
    ''' Files the exists.
    ''' </summary>
    ''' <param name="path">The path.</param>
    ''' <returns></returns>
    Public Function FileExists(path As String) As Boolean Implements IFileSystemShim.FileExists

        Return IO.File.Exists(path)

    End Function

End Class

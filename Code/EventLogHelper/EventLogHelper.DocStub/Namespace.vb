Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Xml.Linq

''' <summary>
''' This class is used to fix the namespaces in the documentation.
''' </summary>
Friend NotInheritable Class [Namespace]

    Public Sub New()
    End Sub

    Public Property NamespaceName As String
    Public Property DocumentationPath As String

    Private summaryValue As String
    Public Property Summary As String
        Get
            Return summaryValue
        End Get
        Set(value As String)
            summaryValue = SanitizeMultilineText(value)
        End Set
    End Property

    Private remarksValue As String
    Public Property Remarks As String
        Get
            Return remarksValue
        End Get
        Set(value As String)
            remarksValue = SanitizeMultilineText(value)
        End Set
    End Property

    Public Property Examples As String

    Public Sub UpdateNamespaceDocumentationFile()
        If NamespaceName Is Nothing Then
            Throw New ArgumentNullException(NameOf(NamespaceName), "Namespace cannot be null.")
        End If
        If DocumentationPath Is Nothing Then
            Throw New ArgumentNullException(NameOf(DocumentationPath), "DocumentationPath cannot be null.")
        End If

        Dim markdownFile = Path.Combine(DocumentationPath, $"{NamespaceName}Namespace.md")

        If File.Exists(markdownFile) Then
            Dim lines = File.ReadAllLines(markdownFile).ToList()
            Dim insertIndex As Integer = -1

            For i = 0 To lines.Count - 1
                If lines(i).Trim().StartsWith("## ", StringComparison.OrdinalIgnoreCase) AndAlso
                       lines(i).Contains($"{NamespaceName} namespace") Then
                    insertIndex = i + 1
                    Exit For
                End If
            Next

            If insertIndex = -1 Then
                Console.WriteLine($"⚠ Could not find header for namespace '{NamespaceName}' in {markdownFile}")
            Else
                Dim insertLines As New List(Of String)()

                If Not String.IsNullOrWhiteSpace(Summary) Then
                    insertLines.Add("")
                    insertLines.Add("## Summary")
                    insertLines.Add(Summary.Trim())
                End If

                If Not String.IsNullOrWhiteSpace(Remarks) Then
                    insertLines.Add("")
                    insertLines.Add("## Remarks")
                    insertLines.Add(Remarks.Trim())
                End If

                Dim example = ConvertExampleElementToMarkdown()
                If Not String.IsNullOrEmpty(example) Then
                    insertLines.Add("")
                    insertLines.Add("## Examples")
                    insertLines.Add(example.Trim())
                End If

                insertLines.Add("")
                insertLines.Add("## Public Types")

                lines.InsertRange(insertIndex, insertLines)
                File.WriteAllLines(markdownFile, lines)
                Console.WriteLine($"✔ Updated {markdownFile}" & vbCrLf & vbCrLf)
            End If
        End If
    End Sub

    Private Shared Function SanitizeMultilineText(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then Return String.Empty

        Dim lines = text.
                Replace(Environment.NewLine, vbLf).
                Split(New String() {vbLf}, StringSplitOptions.None).
                Select(Function(line) line.Trim())

        Dim result As New List(Of String)()
        Dim paragraphBuilder As New List(Of String)()

        For Each line In lines
            If String.IsNullOrWhiteSpace(line) Then
                If paragraphBuilder.Count > 0 Then
                    result.Add(String.Join(" ", paragraphBuilder))
                    paragraphBuilder.Clear()
                End If
                result.Add(String.Empty)
            Else
                paragraphBuilder.Add(line)
            End If
        Next

        If paragraphBuilder.Count > 0 Then
            result.Add(String.Join(" ", paragraphBuilder))
        End If

        Return String.Join(vbLf, result)
    End Function

    Public Function ConvertExampleElementToMarkdown() As String
        If String.IsNullOrEmpty(Examples) Then Return String.Empty

        Dim sb As New StringBuilder()
        Dim exampleElement = XElement.Parse(Examples)

        For Each node In exampleElement.Nodes()
            If TypeOf node Is XText Then
                sb.Append(DirectCast(node, XText).Value.TrimEnd())
            ElseIf TypeOf node Is XElement Then
                Dim el = DirectCast(node, XElement)
                Select Case el.Name.LocalName.ToLowerInvariant()
                    Case "c"
                        sb.Append($"`{el.Value.Trim()}`")
                    Case "code"
                        sb.AppendLine()
                        sb.AppendLine()
                        Dim language = el.Attribute("language")?.Value
                        If String.IsNullOrEmpty(language) Then language = "csharp"

                        Dim lines = el.Value.
                                Replace(vbCrLf, vbLf).
                                Replace(vbCr, vbLf).
                                Split(New String() {vbLf}, StringSplitOptions.None).
                                Select(Function(line) line.TrimEnd()).
                                ToList()

                        While lines.Count > 0 AndAlso String.IsNullOrWhiteSpace(lines(0))
                            lines.RemoveAt(0)
                        End While

                        Dim minIndent = lines.
                                Where(Function(l) Not String.IsNullOrWhiteSpace(l)).
                                Select(Function(l) l.TakeWhile(AddressOf Char.IsWhiteSpace).Count()).
                                DefaultIfEmpty(0).
                                Min()

                        lines = lines.Select(Function(line) If(line.Length >= minIndent, line.Substring(minIndent), line)).ToList()

                        sb.AppendLine($"```{language}")
                        For Each line In lines
                            sb.AppendLine(line)
                        Next
                        sb.AppendLine("```")
                        sb.AppendLine()
                    Case Else
                        sb.Append(el.Value)
                End Select
            End If
            sb.Append(" ")
        Next

        Return sb.ToString().Trim()
    End Function

    Public Overrides Function ToString() As String
        Return $"{NameOf(NamespaceName)} = {NamespaceName}"
    End Function

End Class

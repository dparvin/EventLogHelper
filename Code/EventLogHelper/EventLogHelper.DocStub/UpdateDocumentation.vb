Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Xml.Linq

Friend NotInheritable Class UpdateDocumentation

    Public Sub New(args As String())
        If args.Length < 2 Then
            Console.Error.WriteLine("Usage: <ToolName> <AssemblyPath.dll> <OutputFolder> [options]")
            Environment.Exit(1)
        End If

        ' 1. Path to the DLL
        DllPath = Path.GetFullPath(args(0))

        If Not File.Exists(DllPath) Then
            Console.Error.WriteLine($"Error: Assembly not found: {DllPath}")
            Environment.Exit(1)
        End If

        ' 2. Path to the XML (same folder, same base name)
        XmlPath = Path.ChangeExtension(DllPath, ".xml")

        If Not File.Exists(XmlPath) Then
            Console.Error.WriteLine($"Warning: XML documentation file not found: {XmlPath}")
        End If

        ' 3. Output directory for generated markdown files
        OutputDir = Path.GetFullPath(args(1))

        Console.WriteLine($"✔ DLL Path:     {DllPath}")
        Console.WriteLine($"✔ XML Path:     {XmlPath}")
        Console.WriteLine($"✔ Output Path:  {OutputDir}")
    End Sub

    Public Sub UpdateFiles()
        Dim doc As XDocument = XDocument.Load(XmlPath)

        Dim namespaceDocs = doc.Descendants("member") _
                .Where(Function(m)
                           Dim name As String = CType(m.Attribute("name"), String)
                           Return Not String.IsNullOrEmpty(name) AndAlso
                                  name.StartsWith("T:", StringComparison.OrdinalIgnoreCase) AndAlso
                                  name.EndsWith(".NamespaceDoc", StringComparison.OrdinalIgnoreCase)
                       End Function) _
                .Select(Function(m)
                            Dim fullTypeName As String = CType(m.Attribute("name"), String)
                            Dim ns As String = fullTypeName.Substring(2, fullTypeName.Length - "T:NamespaceDoc".Length - 1)

                            Return New [Namespace] With {
                                .NamespaceName = ns,
                                .Summary = m.Element("summary")?.Value.Trim(),
                                .Remarks = m.Element("remarks")?.Value.Trim(),
                                .Examples = m.Element("example")?.ToString(),
                                .DocumentationPath = OutputDir
                            }
                        End Function) _
                .OrderBy(Function(n) n.NamespaceName) _
                .ToList()

        ' Print for verification
        For Each nsDoc In namespaceDocs
            Console.WriteLine($"✔ Namespace: {nsDoc.NamespaceName}")
            Console.WriteLine($"  Summary: {nsDoc.Summary}")
            Console.WriteLine($"  Remarks: {nsDoc.Remarks}" & vbLf)
            If Not String.IsNullOrWhiteSpace(nsDoc.Examples) Then
                Console.WriteLine($"  Examples: {nsDoc.Examples}")
            End If

            If Not String.Equals(nsDoc.NamespaceName, DllName, StringComparison.OrdinalIgnoreCase) Then
                nsDoc.UpdateNamespaceDocumentationFile()
            Else
                MakeMainPage(namespaceDocs, nsDoc)
            End If
        Next
    End Sub

    Private Sub MakeMainPage(namespaceDocs As List(Of [Namespace]), MainItem As [Namespace])
        Dim markdownFile As String = Path.Combine(MainItem.DocumentationPath, $"{MainItem.NamespaceName}.md")

        Dim insertLines As New List(Of String) From {
                $"# {MainItem.NamespaceName} assembly"
            }

        If Not String.IsNullOrWhiteSpace(MainItem.Summary) Then
            insertLines.Add("")
            insertLines.Add("## Summary")
            insertLines.Add(MainItem.Summary.Trim())
        End If

        If Not String.IsNullOrWhiteSpace(MainItem.Remarks) Then
            insertLines.Add("")
            insertLines.Add("## Remarks")
            insertLines.Add(MainItem.Remarks.Trim())
        End If

        insertLines.Add("")
        insertLines.Add("## Namespaces")
        insertLines.Add("| Name | description |")
        insertLines.Add("| --- | --- |")

        For Each nsDoc In namespaceDocs.
                    Where(Function(n) Not String.Equals(n.NamespaceName, DllName, StringComparison.OrdinalIgnoreCase)).
                    OrderBy(Function(n) n.NamespaceName)

            Dim summaryEscaped = If(nsDoc.Summary, "").Replace("|", "\|").Trim()
            insertLines.Add($"| [{nsDoc.NamespaceName}]({nsDoc.NamespaceName}Namespace.md) | {summaryEscaped} |")
        Next

        insertLines.Add("") ' Final newline

        File.WriteAllLines(markdownFile, insertLines)
        Console.WriteLine($"✔ Updated {markdownFile}")
    End Sub

    Public Property DllPath As String
    Public Property XmlPath As String
    Public Property OutputDir As String

    Public ReadOnly Property DllName As String
        Get
            Return Path.GetFileNameWithoutExtension(DllPath)
        End Get
    End Property

End Class

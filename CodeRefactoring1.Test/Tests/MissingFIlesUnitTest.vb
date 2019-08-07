Imports System.IO
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.MSBuild

<TestClass()> Public Class MissingFileTest
    <TestMethod()>
    Public Sub CheckForMissingFiles()
        Dim slnPath As String = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "..", "..", ".."))

        For Each filename As String In Directory.GetFiles(slnPath, "*.FileListAbsolute.txt", IO.SearchOption.AllDirectories)
            If filename.Contains("\.vs\") Then Continue For
            Using reader As StreamReader = File.OpenText(filename)
                While (reader.Peek() <> -1)
                    Dim FileToFind As String = reader.ReadLine()
                    If FileToFind.ToLower.EndsWith(".dll") Then
                        Assert.IsTrue(File.Exists(FileToFind), $"File {FileToFind} is missing")
                    End If
                End While
            End Using
        Next
        Using workspace As MSBuildWorkspace = MSBuildWorkspace.Create()
            Dim solution As Solution = workspace.OpenSolutionAsync(Path.Combine(slnPath, "CodeRefactoring1.sln")).Result
        End Using
    End Sub

End Class
Imports System.IO

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace AllFiles.UnitTest

    <TestClass()> Public Class AllFilesUnitTest
        Const StartDir As String = "C:\Users\PaulM\Source\Repos\roslyn\src\Compilers\VisualBasic\Portable"
        Private ErrorCount As Integer = 0
        Private SuccessCount As Integer = 0

        Private Sub ProcessOneDirectory(DirName As String)
            For Each Dir As String In Directory.EnumerateDirectories(DirName)
                Me.ProcessOneDirectory(Dir)
            Next
            For Each FileName As String In Directory.EnumerateFiles(DirName, "*.vb")
                Me.ProcessOneFile(FileName)
            Next
        End Sub

        Private Sub ProcessOneFile(fileName As String)
            Dim Text As String
            Text = My.Computer.FileSystem.ReadAllText(fileName)
            Text = "
Option Infer On
Option Strict On
" & Text
            Dim comp As VisualBasicCompilation = CreateCompilationWithMscorlibAndVBRuntime(Text, "a.vb")

            Dim tree As SyntaxTree = comp.SyntaxTrees(0)

            Dim model As SemanticModel = comp.GetSemanticModel(tree)
            Dim nodes As IEnumerable(Of SyntaxNode) = tree.GetCompilationUnitRoot().DescendantNodes()
            For Each node As ForEachStatementSyntax In nodes.OfType(Of ForEachStatementSyntax)()
                Dim foreachInfo As ForEachStatementInfo = model.GetForEachStatementInfo(node)
                Dim elementType As ITypeSymbol = foreachInfo.ElementType
                If foreachInfo.ElementType Is Nothing Then
                    Me.ErrorCount += 1
                Else
                    Me.SuccessCount += 1
                End If
            Next
        End Sub

        <TestMethod()>
        <Ignore()>
        Public Sub ElementTypeUnitTest()
            Me.ProcessOneDirectory(StartDir)
            Xunit.Assert.True(Me.ErrorCount = Me.SuccessCount, $"Errors = {Me.ErrorCount} Success = {Me.SuccessCount}")
        End Sub

    End Class

End Namespace
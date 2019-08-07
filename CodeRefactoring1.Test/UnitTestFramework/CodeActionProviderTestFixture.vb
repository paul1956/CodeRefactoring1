Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Formatting
Imports Microsoft.CodeAnalysis.Simplification
Imports Microsoft.CodeAnalysis.Text
Imports Xunit

Namespace Roslyn.UnitTestFramework
    Public MustInherit Class CodeActionProviderTestFixture
        Protected Function CreateDocument(ByVal code As String) As Document
            Dim fileExtension As String = If(Me.LanguageName = LanguageNames.CSharp, ".cs", ".vb")

            Dim projectId As ProjectId = ProjectId.CreateNewId(debugName:="TestProject")
            Dim documentId As DocumentId = DocumentId.CreateNewId(projectId, debugName:="Test" & fileExtension)

            Using NewAdhocWorkspace As AdhocWorkspace = New AdhocWorkspace()
                Return NewAdhocWorkspace.CurrentSolution.AddProject(projectId, "TestProject", "TestProject", Me.LanguageName).AddMetadataReferences(projectId, SharedReferences.References).AddDocument(documentId, "Test" & fileExtension, SourceText.From(code)).GetDocument(documentId)
            End Using
        End Function

        Protected Sub VerifyDocument(ByVal expected As String, ByVal compareTokens As Boolean, ByVal document As Document)
            If compareTokens Then
                Me.VerifyTokens(expected, Me.Format(document).ToString())
            Else
                Me.VerifyText(expected, document)
            End If
        End Sub

        Private Function Format(ByVal document As Document) As SyntaxNode
            Dim updatedDocument As Document = document.WithSyntaxRoot(document.GetSyntaxRootAsync().Result)
            Return Formatter.FormatAsync(Simplifier.ReduceAsync(updatedDocument, Simplifier.Annotation).Result, Formatter.Annotation).Result.GetSyntaxRootAsync().Result
        End Function

        Private Function ParseTokens(ByVal text As String) As IList(Of SyntaxToken)
            Return Me.ParseTokens(text).Select(Function(t As SyntaxToken) CType(t, SyntaxToken)).ToList()
        End Function

        Private Function VerifyTokens(ByVal expected As String, ByVal actual As String) As Boolean
            Dim expectedNewTokens As IList(Of SyntaxToken) = Me.ParseTokens(expected)
            Dim actualNewTokens As IList(Of SyntaxToken) = Me.ParseTokens(actual)

            For i As Integer = 0 To Math.Min(expectedNewTokens.Count, actualNewTokens.Count) - 1
                Assert.Equal(expectedNewTokens(i).ToString(), actualNewTokens(i).ToString())
            Next i

            If expectedNewTokens.Count <> actualNewTokens.Count Then
                Dim expectedDisplay As String = String.Join(" ", expectedNewTokens.Select(Function(t As SyntaxToken) t.ToString()))
                Dim actualDisplay As String = String.Join(" ", actualNewTokens.Select(Function(t As SyntaxToken) t.ToString()))
                Assert.True(False, String.Format("Wrong token count. Expected '{0}', Actual '{1}', Expected Text: '{2}', Actual Text: '{3}'", expectedNewTokens.Count, actualNewTokens.Count, expectedDisplay, actualDisplay))
            End If

            Return True
        End Function

        Private Function VerifyText(ByVal expected As String, ByVal document As Document) As Boolean
            Dim actual As String = Me.Format(document).ToString()
            Assert.Equal(expected, actual, ignoreWhiteSpaceDifferences:=True)
            Return True
        End Function

        Protected MustOverride ReadOnly Property LanguageName() As String
    End Class
End Namespace

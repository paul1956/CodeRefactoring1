Imports System.Collections.Generic
Imports System.Collections.Immutable
Imports System.Linq
Imports System.Threading.Tasks
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeActions
Imports Microsoft.CodeAnalysis.CodeFixes

''' <summary>
''' Provides a base class to write a <see cref="FixAllProvider"/> that fixes documents independently.
''' </summary>
Friend MustInherit Class DocumentBasedFixAllProvider
    Inherits FixAllProvider

    Protected MustOverride ReadOnly Property CodeActionTitle() As String

    Public Overrides Function GetFixAsync(ByVal fixAllContext As FixAllContext) As Task(Of CodeAction)
        Dim fixAction As CodeAction
        Select Case fixAllContext.Scope
            Case FixAllScope.Document
                fixAction = CodeAction.Create(Me.CodeActionTitle, Function(c As CancellationToken) Me.GetDocumentFixesAsync(fixAllContext.WithCancellationToken(c)), NameOf(DocumentBasedFixAllProvider))
            Case FixAllScope.Project
                fixAction = CodeAction.Create(Me.CodeActionTitle, Function(c As CancellationToken) Me.GetProjectFixesAsync(fixAllContext.WithCancellationToken(c), fixAllContext.Project), NameOf(DocumentBasedFixAllProvider))
            Case FixAllScope.Solution
                fixAction = CodeAction.Create(Me.CodeActionTitle, Function(c As CancellationToken) Me.GetSolutionFixesAsync(fixAllContext.WithCancellationToken(c)), NameOf(DocumentBasedFixAllProvider))
            Case Else
                fixAction = Nothing
        End Select
        Return Task.FromResult(fixAction)
    End Function

    ''' <summary>
    ''' Fixes all occurrences of a diagnostic in a specific document.
    ''' </summary>
    ''' <param name="fixAllContext">The context for the Fix All operation.</param>
    ''' <param name="document">The document to fix.</param>
    ''' <param name="diagnostics">The diagnostics to fix in the document.</param>
    ''' <returns>
    ''' <para>The new <see cref="SyntaxNode"/> representing the root of the fixed document.</para>
    ''' <para>-or-</para>
    ''' <para><see langword="null"/>, if no changes were made to the document.</para>
    ''' </returns>
    Protected MustOverride Function FixAllInDocumentAsync(ByVal fixAllContext As FixAllContext, ByVal document As Document, ByVal diagnostics As ImmutableArray(Of Diagnostic)) As Task(Of SyntaxNode)

    Private Async Function GetDocumentFixesAsync(ByVal fixAllContext As FixAllContext) As Task(Of Document)
        Dim documentDiagnosticsToFix As ImmutableDictionary(Of Document, ImmutableArray(Of Diagnostic)) = Await FixAllContextHelper.GetDocumentDiagnosticsToFixAsync(fixAllContext).ConfigureAwait(False)
        Dim diagnostics As ImmutableArray(Of Diagnostic) = Nothing
        If Not documentDiagnosticsToFix.TryGetValue(fixAllContext.Document, diagnostics) Then
            Return fixAllContext.Document
        End If

        Dim newRoot As SyntaxNode = Await Me.FixAllInDocumentAsync(fixAllContext, fixAllContext.Document, diagnostics).ConfigureAwait(False)
        If newRoot Is Nothing Then
            Return fixAllContext.Document
        End If

        Return fixAllContext.Document.WithSyntaxRoot(newRoot)
    End Function

    Private Async Function GetSolutionFixesAsync(ByVal fixAllContext As FixAllContext, ByVal documents As ImmutableArray(Of Document)) As Task(Of Solution)
        Dim documentDiagnosticsToFix As ImmutableDictionary(Of Document, ImmutableArray(Of Diagnostic)) = Await FixAllContextHelper.GetDocumentDiagnosticsToFixAsync(fixAllContext).ConfigureAwait(False)

        Dim solution_Renamed As Solution = fixAllContext.Solution
        Dim newDocuments As New List(Of Task(Of SyntaxNode))(documents.Length)
        For Each document As Document In documents
            Dim diagnostics As ImmutableArray(Of Diagnostic) = Nothing
            If Not documentDiagnosticsToFix.TryGetValue(document, diagnostics) Then
                newDocuments.Add(document.GetSyntaxRootAsync(fixAllContext.CancellationToken))
                Continue For
            End If

            newDocuments.Add(Me.FixAllInDocumentAsync(fixAllContext, document, diagnostics))
        Next document

        For i As Integer = 0 To documents.Length - 1
            Dim newDocumentRoot As SyntaxNode = Await newDocuments(i).ConfigureAwait(False)
            If newDocumentRoot Is Nothing Then
                Continue For
            End If

            solution_Renamed = solution_Renamed.WithDocumentSyntaxRoot(documents(i).Id, newDocumentRoot)
        Next i

        Return solution_Renamed
    End Function

    Private Function GetProjectFixesAsync(ByVal fixAllContext As FixAllContext, ByVal project As Project) As Task(Of Solution)
        Return Me.GetSolutionFixesAsync(fixAllContext, project.Documents.ToImmutableArray())
    End Function

    Private Function GetSolutionFixesAsync(ByVal fixAllContext As FixAllContext) As Task(Of Solution)
        Dim documents As ImmutableArray(Of Document) = fixAllContext.Solution.Projects.SelectMany(Function(i As Project) i.Documents).ToImmutableArray()
        Return Me.GetSolutionFixesAsync(fixAllContext, documents)
    End Function
End Class

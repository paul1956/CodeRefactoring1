Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.CodeFixes

<ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(ByValAnalyzer_FixCodeFixProvider)), [Shared]>
Public Class ByValAnalyzer_FixCodeFixProvider
    Inherits CodeFixProvider

    Private Const title As String = "Remove ByVal"

    Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String)
        Get
            Return ImmutableArray.Create(ByValAnalyzer_FixAnalyzer.DiagnosticId)
        End Get
    End Property

    Private Async Function RemoveOccuranceAsync(document As Document, token As SyntaxToken, cancellationToken As CancellationToken) As Task(Of Document)
        Dim oldRoot As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
        Dim rewriter As Rewriter = New Rewriter(Function(t As SyntaxToken) t = token)
        Dim newRoot As SyntaxNode = rewriter.Visit(oldRoot)
        Return document.WithSyntaxRoot(newRoot)
    End Function

    Public NotOverridable Overrides Function GetFixAllProvider() As FixAllProvider
        ' See https://github.com/dotnet/roslyn/blob/master/docs/analyzers/FixAllProvider.md for more information on Fix All Providers
        Return WellKnownFixAllProviders.BatchFixer
    End Function

    Public NotOverridable Overrides Async Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
        Dim diagnostic As Diagnostic = context.Diagnostics.First()
        Dim root As SyntaxNode = Await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(False)

        Dim diagnosticSpan As TextSpan = diagnostic.Location.SourceSpan

        Dim textSpan As TextSpan = context.Span
        Dim token As SyntaxToken = root.FindToken(textSpan.Start)

        ' Register a code action that will invoke the fix.
        Dim CodeAction As CodeAction = CodeAction.Create(
                title:=title,
                createChangedDocument:=Function(c) Me.RemoveOccuranceAsync(context.Document, token, c),
                equivalenceKey:=title)
        context.RegisterCodeFix(CodeAction, diagnostic)
    End Function

    Private Class RemoveByValCodeAction
        Inherits CodeAction

        Private ReadOnly _title As String
        Private ReadOnly createChangedDocument As Func(Of Object, Task(Of Document))

        Public Sub New(title As String, createChangedDocument As Func(Of Object, Task(Of Document)))
            Me._title = title
            Me.createChangedDocument = createChangedDocument
        End Sub

        Public Overrides ReadOnly Property Title As String
            Get
                Return Me._title
            End Get
        End Property

        Protected Overrides Function GetChangedDocumentAsync(cancellationToken As CancellationToken) As Task(Of Document)
            Return Me.createChangedDocument(cancellationToken)
        End Function

    End Class

    Private Class Rewriter
        Inherits VisualBasicSyntaxRewriter

        Private ReadOnly _predicate As Func(Of SyntaxToken, Boolean)

        Public Sub New(predicate As Func(Of SyntaxToken, Boolean))
            Me._predicate = predicate
        End Sub

        Public Overrides Function VisitToken(token As SyntaxToken) As SyntaxToken
            If token.Kind = SyntaxKind.ByValKeyword AndAlso Me._predicate(token) Then
                Dim NewTrailingTrivia As New List(Of SyntaxTrivia)
                Dim TrailingTrivia As SyntaxTriviaList = token.TrailingTrivia
                Dim TriviaUBound As Integer = TrailingTrivia.Count - 1
                Dim FirstContinuation As Boolean = True
                If TriviaUBound > 1 Then
                    For i As Integer = 0 To TriviaUBound
                        Dim Trivia As SyntaxTrivia = TrailingTrivia(i)
                        Dim NextTrivia As SyntaxTrivia = If(i < TriviaUBound, TrailingTrivia(i + 1), Nothing)
                        If Trivia.IsKind(SyntaxKind.WhitespaceTrivia) AndAlso NextTrivia.IsKind(SyntaxKind.LineContinuationTrivia) Then
                            If FirstContinuation Then
                                i += 2
                                FirstContinuation = False
                                Continue For
                            End If
                            If Trivia.IsKind(SyntaxKind.EndOfLineTrivia) Then
                                FirstContinuation = False
                            End If
                        End If
                        NewTrailingTrivia.Add(Trivia)
                    Next
                End If
                Return SyntaxFactory.Token(token.LeadingTrivia, SyntaxKind.EmptyToken, NewTrailingTrivia.ToSyntaxTriviaList)
            End If

            Return token
        End Function

    End Class

End Class
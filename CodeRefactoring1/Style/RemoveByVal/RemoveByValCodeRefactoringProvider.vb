Namespace Style

    <ExportCodeRefactoringProvider(LanguageNames.VisualBasic, Name:="RemoveByValVB"), [Shared]>
    Class RemoveByValCodeRefactoringProvider
        Inherits CodeRefactoringProvider

        Private Async Function RemoveAllOccurancesAsync(document As Document, cancellationToken As CancellationToken) As Task(Of Document)
            Dim oldRoot As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim rewriter As Rewriter = New Rewriter(Function(current As SyntaxToken) True)
            Dim newRoot As SyntaxNode = rewriter.Visit(oldRoot)
            Return document.WithSyntaxRoot(newRoot)
        End Function

        Private Async Function RemoveOccuranceAsync(document As Document, token As SyntaxToken, cancellationToken As CancellationToken) As Task(Of Document)
            Dim oldRoot As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim rewriter As Rewriter = New Rewriter(Function(t As SyntaxToken) t = token)
            Dim newRoot As SyntaxNode = rewriter.Visit(oldRoot)
            Return document.WithSyntaxRoot(newRoot)
        End Function

        Public NotOverridable Overrides Async Function ComputeRefactoringsAsync(context As CodeRefactoringContext) As Task
            Dim document As Document = context.Document
            Dim textSpan As TextSpan = context.Span
            Dim cancellationToken As CancellationToken = context.CancellationToken

            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim token As SyntaxToken = root.FindToken(textSpan.Start)

            If token.Kind = SyntaxKind.ByValKeyword AndAlso token.Span.IntersectsWith(textSpan.Start) Then
                context.RegisterRefactoring(New RemoveByValCodeAction("Remove unnecessary ByVal keyword",
                                                                  CType(Function(c As CancellationToken) Me.RemoveOccuranceAsync(document, token, c), Func(Of Object, Task(Of Document)))))
                context.RegisterRefactoring(New RemoveByValCodeAction("Remove all occurrences of unnecessary ByVal keywords",
                                                                  CType(Function(c As CancellationToken) Me.RemoveAllOccurancesAsync(document, c), Func(Of Object, Task(Of Document)))))
            End If
        End Function

        Class RemoveByValCodeAction
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

        Class Rewriter
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

End Namespace
Namespace Style
    <ExportCodeRefactoringProvider(LanguageNames.VisualBasic, Name:="RemoveByValVB"), [Shared]>
    Class RemoveByValCodeRefactoringProvider
        Inherits CodeRefactoringProvider

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

        Private Async Function RemoveOccuranceAsync(document As Document, token As SyntaxToken, cancellationToken As CancellationToken) As Task(Of Document)
            Dim oldRoot As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim rewriter As Rewriter = New Rewriter(Function(t As SyntaxToken) t = token)
            Dim newRoot As SyntaxNode = rewriter.Visit(oldRoot)
            Return document.WithSyntaxRoot(newRoot)
        End Function

        Private Async Function RemoveAllOccurancesAsync(document As Document, cancellationToken As CancellationToken) As Task(Of Document)
            Dim oldRoot As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim rewriter As Rewriter = New Rewriter(Function(current As SyntaxToken) True)
            Dim newRoot As SyntaxNode = rewriter.Visit(oldRoot)
            Return document.WithSyntaxRoot(newRoot)
        End Function

        Class Rewriter
            Inherits VisualBasicSyntaxRewriter

            Private ReadOnly _predicate As Func(Of SyntaxToken, Boolean)

            Public Sub New(predicate As Func(Of SyntaxToken, Boolean))
                Me._predicate = predicate
            End Sub

            Public Overrides Function VisitToken(token As SyntaxToken) As SyntaxToken
                If token.Kind = SyntaxKind.ByValKeyword AndAlso Me._predicate(token) Then
                    Return SyntaxFactory.Token(token.LeadingTrivia, SyntaxKind.ByValKeyword, Nothing, String.Empty)
                End If

                Return token
            End Function
        End Class

        Class RemoveByValCodeAction
            Inherits CodeAction

            Private ReadOnly createChangedDocument As Func(Of Object, Task(Of Document))
            Private ReadOnly _title As String

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
    End Class
End Namespace
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Formatting
Imports System.Collections.Immutable


Namespace Design
    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(EmptyCatchBlockCodeFixProvider)), Composition.Shared>
    Public Class EmptyCatchBlockCodeFixProvider
        Inherits CodeFixProvider
        Friend Shared ReadOnly FixInsertExceptionClass As New LocalizableResourceString(NameOf(My.Resources.EmptyCatchBlockCodeFixProvider_InsertException), My.Resources.ResourceManager, GetType(My.Resources.Resources))

        Friend Shared ReadOnly FixRemoveEmptyCatchBlock As New LocalizableResourceString(NameOf(My.Resources.EmptyCatchBlockCodeFixProvider_Remove), My.Resources.ResourceManager, GetType(My.Resources.Resources))
        Friend Shared ReadOnly FixRemoveTry As New LocalizableResourceString(NameOf(My.Resources.EmptyCatchBlockCodeFixProvider_RemoveTry), My.Resources.ResourceManager, GetType(My.Resources.Resources))

        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(DiagnosticIds.EmptyCatchBlockDiagnosticId)

        Public Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Overrides Async Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim root As SyntaxNode = Await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(False)
            Dim diag As Diagnostic = context.Diagnostics.First
            Dim diagSpan As Text.TextSpan = diag.Location.SourceSpan
            Dim declaration As CatchBlockSyntax = root.FindToken(diagSpan.Start).Parent.AncestorsAndSelf.OfType(Of CatchBlockSyntax).First

            Dim tryBlock As VisualBasic.Syntax.TryBlockSyntax = DirectCast(declaration.Parent, TryBlockSyntax)
            If tryBlock.CatchBlocks.Count > 1 Then
                context.RegisterCodeFix(CodeAction.Create(FixRemoveEmptyCatchBlock.ToString(), Function(c As CancellationToken) Me.RemoveCatch(context.Document, declaration, c), NameOf(EmptyCatchBlockCodeFixProvider) & NameOf(RemoveTry)), diag)
            Else
                context.RegisterCodeFix(CodeAction.Create(FixRemoveTry.ToString(), Function(c As CancellationToken) Me.RemoveTry(context.Document, declaration, c), NameOf(EmptyCatchBlockCodeFixProvider) & NameOf(RemoveTry)), diag)
            End If
            context.RegisterCodeFix(CodeAction.Create(FixInsertExceptionClass.ToString(), Function(c As CancellationToken) Me.InsertExceptionClassCommentAsync(context.Document, declaration, c), NameOf(EmptyCatchBlockCodeFixProvider) & NameOf(InsertExceptionClassCommentAsync)), diag)
        End Function

        Private Async Function InsertExceptionClassCommentAsync(document As Document, catchBlock As CatchBlockSyntax, cancellationToken As CancellationToken) As Task(Of Document)
            Dim statements As SyntaxList(Of SyntaxNode) = New SyntaxList(Of SyntaxNode)().Add(SyntaxFactory.ThrowStatement())

            Dim catchStatement As VisualBasic.Syntax.CatchStatementSyntax = SyntaxFactory.CatchStatement(
            SyntaxFactory.IdentifierName("ex"),
            SyntaxFactory.SimpleAsClause(SyntaxFactory.IdentifierName(NameOf(Exception))),
            Nothing)

            Dim catchClause As CatchBlockSyntax = SyntaxFactory.CatchBlock(catchStatement, statements).
            WithLeadingTrivia(catchBlock.GetLeadingTrivia).
            WithTrailingTrivia(catchBlock.GetTrailingTrivia).
            WithAdditionalAnnotations(Formatter.Annotation)

            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim newRoot As SyntaxNode = root.ReplaceNode(catchBlock, catchClause)
            Dim newDocument As Document = document.WithSyntaxRoot(newRoot)
            Return newDocument

        End Function

        Private Async Function RemoveCatch(document As Document, catchBlock As CatchBlockSyntax, cancellationToken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)

            Dim newRoot As SyntaxNode = root.RemoveNode(catchBlock, SyntaxRemoveOptions.KeepNoTrivia)

            Dim newDocument As Document = document.WithSyntaxRoot(newRoot)
            Return newDocument
        End Function

        Private Async Function RemoveTry(document As Document, catchBlock As CatchBlockSyntax, cancellationToken As CancellationToken) As Task(Of Document)
            Dim tryBlock As VisualBasic.Syntax.TryBlockSyntax = DirectCast(catchBlock.Parent, TryBlockSyntax)
            Dim statements As SyntaxList(Of VisualBasic.Syntax.StatementSyntax) = tryBlock.Statements
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)

            Dim newRoot As SyntaxNode = root.ReplaceNode(catchBlock.Parent,
                                       statements.Select(Function(s As StatementSyntax) s.
                                            WithLeadingTrivia(catchBlock.Parent.GetLeadingTrivia()).
                                            WithTrailingTrivia(catchBlock.Parent.GetTrailingTrivia())))

            Dim newDocument As Document = document.WithSyntaxRoot(newRoot)
            Return newDocument
        End Function

    End Class
End Namespace
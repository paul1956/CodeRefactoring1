Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Formatting

Namespace Usage
    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(DisposablesShouldCallSuppressFinalizeCodeFixProvider)), Composition.Shared>
    Public Class DisposablesShouldCallSuppressFinalizeCodeFixProvider
        Inherits CodeFixProvider

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim diagnostic As Diagnostic = context.Diagnostics.First
            context.RegisterCodeFix(CodeAction.Create("Call GC.SuppressFinalize", Function(ct As CancellationToken) Me.AddSuppressFinalizeAsync(context.Document, diagnostic, ct), NameOf(DisposablesShouldCallSuppressFinalizeCodeFixProvider)), diagnostic)
            Return Task.FromResult(0)
        End Function

        Public NotOverridable Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(DiagnosticIds.DisposablesShouldCallSuppressFinalizeDiagnosticId)

        Public Async Function AddSuppressFinalizeAsync(document As Document, diagnostic As Diagnostic, cancellationToken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim span As TextSpan = diagnostic.Location.SourceSpan
            Dim method As MethodBlockSyntax = root.FindToken(span.Start).Parent.FirstAncestorOrSelf(Of MethodBlockSyntax)()

            Dim suppressInvocation As ExpressionStatementSyntax =
                    SyntaxFactory.ExpressionStatement(
                        SyntaxFactory.InvocationExpression(
                            SyntaxFactory.SimpleMemberAccessExpression(
                                    SyntaxFactory.IdentifierName(NameOf(GC)),
                                    SyntaxFactory.IdentifierName("SuppressFinalize"))).
                            WithArgumentList(SyntaxFactory.ArgumentList().AddArguments(SyntaxFactory.SimpleArgument(SyntaxFactory.MeExpression)))).
                            WithTrailingTrivia(SyntaxFactory.CarriageReturnLineFeed)

            Dim newMethod As MethodBlockSyntax = SyntaxFactory.SubBlock(method.SubOrFunctionStatement).
                WithStatements(method.Statements.Add(suppressInvocation)).
                WithAdditionalAnnotations(Formatter.Annotation)

            Return document.
            WithSyntaxRoot((Await document.GetSyntaxRootAsync(cancellationToken)).
                           ReplaceNode(method, newMethod).
                           WithAdditionalAnnotations(Formatter.Annotation))

        End Function
    End Class
End Namespace
Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Formatting

Namespace Design

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(NameOfCodeFixProvider)), Composition.Shared>
    Public Class NameOfCodeFixProvider
        Inherits CodeFixProvider

        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(NameOfDiagnosticId, NameOf_ExternalDiagnosticId)

        Private Async Function MakeNameOf(document As Document, diagnostic As Diagnostic, cancellationToken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim diagnosticspan As Text.TextSpan = diagnostic.Location.SourceSpan
            Dim stringLiteral As LiteralExpressionSyntax = root.FindToken(diagnosticspan.Start).Parent.AncestorsAndSelf.OfType(Of LiteralExpressionSyntax).FirstOrDefault

            Dim newNameof As ExpressionSyntax = SyntaxFactory.ParseExpression($"NameOf({stringLiteral.Token.ToString().Replace("""", "")})").
                WithLeadingTrivia(stringLiteral.GetLeadingTrivia).
                WithTrailingTrivia(stringLiteral.GetTrailingTrivia).
                WithAdditionalAnnotations(Formatter.Annotation)

            Dim newRoot As SyntaxNode = root.ReplaceNode(stringLiteral, newNameof)
            Return document.WithSyntaxRoot(newRoot)
        End Function

        Public Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim diagnostic As Diagnostic = context.Diagnostics.First
            context.RegisterCodeFix(CodeAction.Create(My.Resources.NameOfAnalyzer_Title, Function(c As CancellationToken) Me.MakeNameOf(context.Document, diagnostic, c), NameOf(NameOfCodeFixProvider)), diagnostic)
            Return Task.FromResult(0)
        End Function

    End Class

End Namespace
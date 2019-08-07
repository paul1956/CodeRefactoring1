Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.CodeFixes

Namespace Usage

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(MustInheritClassShouldNotHavePublicConstructorsCodeFixProvider)), Composition.Shared>
    Public Class MustInheritClassShouldNotHavePublicConstructorsCodeFixProvider
        Inherits CodeFixProvider

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim diag As Diagnostic = context.Diagnostics.First()
            context.RegisterCodeFix(CodeAction.Create("Use 'Friend' instead of 'Public'", Function(c As CancellationToken) Me.ReplacePublicWithProtectedAsync(context.Document, diag, c), NameOf(MustInheritClassShouldNotHavePublicConstructorsCodeFixProvider)), diag)
            Return Task.FromResult(0)
        End Function

        Public NotOverridable Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(DiagnosticIds.AbstractClassShouldNotHavePublicCtorsDiagnosticId)

        Private Async Function ReplacePublicWithProtectedAsync(document As Document, diagnostic As Diagnostic, cancellationToken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim span As TextSpan = diagnostic.Location.SourceSpan

            Dim constructor As SubNewStatementSyntax = root.FindToken(span.Start).Parent.FirstAncestorOrSelf(Of SubNewStatementSyntax)

            Dim [public] As SyntaxToken = constructor.Modifiers.First(Function(m As SyntaxToken) m.IsKind(SyntaxKind.PublicKeyword))
            Dim [protected] As SyntaxToken = SyntaxFactory.Token([public].LeadingTrivia, SyntaxKind.ProtectedKeyword, [public].TrailingTrivia)

            Dim newModifiers As SyntaxTokenList = constructor.Modifiers.Replace([public], [protected])
            Dim newConstructor As SubNewStatementSyntax = constructor.WithModifiers(newModifiers)
            Dim newRoot As SyntaxNode = root.ReplaceNode(constructor, newConstructor)
            Dim newDocumnent As Document = document.WithSyntaxRoot(newRoot)
            Return newDocumnent
        End Function

    End Class

End Namespace
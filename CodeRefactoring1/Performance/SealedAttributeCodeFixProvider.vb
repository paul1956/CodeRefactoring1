Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.CodeFixes

Namespace Performance
    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(SealedAttributeCodeFixProvider)), Composition.Shared>
    Public Class SealedAttributeCodeFixProvider
        Inherits CodeFixProvider

        Public Overrides NotOverridable ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(SealedAttributeAnalyzer.Id)

        Public Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Overrides Async Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim root As SyntaxNode = Await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(False)
            Dim diag As Diagnostic = context.Diagnostics.First()
            Dim sourceSpan As Text.TextSpan = diag.Location.SourceSpan
            Dim type As VisualBasic.Syntax.ClassStatementSyntax = root.FindToken(sourceSpan.Start).Parent.AncestorsAndSelf().OfType(Of ClassStatementSyntax)().First()
            context.RegisterCodeFix(CodeAction.Create("Mark as NotInheritable", Function(ct As CancellationToken) Me.MarkClassAsSealed(context.Document, diag, ct), NameOf(SealedAttributeCodeFixProvider)), diag)
        End Function

        Private Async Function MarkClassAsSealed(document As Document, diagnostic As Diagnostic, cancellationToken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim sourceSpan As Text.TextSpan = diagnostic.Location.SourceSpan
            Dim type As VisualBasic.Syntax.ClassStatementSyntax = root.FindToken(sourceSpan.Start).Parent.AncestorsAndSelf().OfType(Of ClassStatementSyntax)().First()

            Return document.WithSyntaxRoot(
                (Await document.GetSyntaxRootAsync(cancellationToken)).
                ReplaceNode(type, type.WithModifiers(type.Modifiers.
                Add(SyntaxFactory.Token(SyntaxKind.NotInheritableKeyword))).
                WithAdditionalAnnotations(Formatting.Formatter.Annotation)))
        End Function
    End Class
End Namespace
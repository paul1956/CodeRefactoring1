' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Formatting

Namespace Performance

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(MakeLocalVariableConstWhenPossibleCodeFixProvider)), Composition.Shared>
    Public Class MakeLocalVariableConstWhenPossibleCodeFixProvider
        Inherits CodeFixProvider

        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(MakeLocalVariableConstWhenPossibleAnalyzer.Id)

        Public Shared Async Function MakeConstantAsync(document As Document, diagnostic As Diagnostic, cancellationToken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim diagnosticSpan As Text.TextSpan = diagnostic.Location.SourceSpan
            Dim localDeclaration As VisualBasic.Syntax.LocalDeclarationStatementSyntax = root.FindToken(diagnosticSpan.Start).Parent.AncestorsAndSelf().OfType(Of LocalDeclarationStatementSyntax).First()

            Dim declaration As VisualBasic.Syntax.VariableDeclaratorSyntax = localDeclaration.Declarators.First

            Dim dimModifier As SyntaxToken = localDeclaration.Modifiers.First()

            Dim constant As SyntaxToken = SyntaxFactory.Token(SyntaxKind.ConstKeyword).
            WithLeadingTrivia(dimModifier.LeadingTrivia).
            WithTrailingTrivia(dimModifier.TrailingTrivia).
            WithAdditionalAnnotations(Formatter.Annotation)

            Dim modifiers As SyntaxTokenList = localDeclaration.Modifiers.Replace(dimModifier, constant)

            Dim newLocalDeclaration As VisualBasic.Syntax.LocalDeclarationStatementSyntax = localDeclaration.
            WithModifiers(modifiers).
            WithLeadingTrivia(localDeclaration.GetLeadingTrivia()).
            WithTrailingTrivia(localDeclaration.GetTrailingTrivia()).
            WithAdditionalAnnotations(Formatter.Annotation)

            Dim newRoot As SyntaxNode = root.ReplaceNode(localDeclaration, newLocalDeclaration)
            Return document.WithSyntaxRoot(newRoot)
        End Function

        Public Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim diagnostic As Diagnostic = context.Diagnostics.First()
            Const message As String = "Make constant"
            context.RegisterCodeFix(CodeAction.Create(message, Function(c As CancellationToken) MakeConstantAsync(context.Document, diagnostic, c), NameOf(MakeLocalVariableConstWhenPossibleCodeFixProvider)), diagnostic)
            Return Task.FromResult(0)
        End Function

    End Class

End Namespace

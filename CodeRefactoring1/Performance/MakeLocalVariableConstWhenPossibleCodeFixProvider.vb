﻿Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Formatting
Imports System.Collections.Immutable

Namespace Performance
    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(MakeLocalVariableConstWhenPossibleCodeFixProvider)), Composition.Shared>
    Public Class MakeLocalVariableConstWhenPossibleCodeFixProvider
        Inherits CodeFixProvider

        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(MakeLocalVariableConstWhenPossibleAnalyzer.Id)

        Public Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim diagnostic As Diagnostic = context.Diagnostics.First()
            Const message As String = "Make constant"
            context.RegisterCodeFix(CodeAction.Create(message, Function(c As CancellationToken) Me.MakeConstantAsync(context.Document, diagnostic, c), NameOf(MakeLocalVariableConstWhenPossibleCodeFixProvider)), diagnostic)
            Return Task.FromResult(0)
        End Function

        Public Async Function MakeConstantAsync(document As Document, diagnostic As Diagnostic, cancellationToken As CancellationToken) As Task(Of Document)
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
    End Class
End Namespace
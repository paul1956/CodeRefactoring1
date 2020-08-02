' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeActions
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Formatting
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Design

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(CatchEmptyCodeFixProvider)), [Shared]>
    Public Class CatchEmptyCodeFixProvider
        Inherits CodeFixProvider

        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(CatchEmptyAnalyzer.Id)

        Public Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim diag As Diagnostic = context.Diagnostics.First()
            context.RegisterCodeFix(CodeAction.Create("Add an Exception class", Function(c As CancellationToken) Me.MakeCatchEmptyAsync(context.Document, diag, c), NameOf(CatchEmptyCodeFixProvider)), diag)
            Return Task.FromResult(0)
        End Function

        Private Async Function MakeCatchEmptyAsync(document As Document, diag As Diagnostic, cancellationtoken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationtoken).ConfigureAwait(False)
            Dim diagSpan As Text.TextSpan = diag.Location.SourceSpan
            Dim catchStatement As CatchBlockSyntax = root.FindToken(diagSpan.Start).Parent.AncestorsAndSelf.OfType(Of CatchBlockSyntax).First()
            Dim semanticModel As SemanticModel = Await document.GetSemanticModelAsync(cancellationtoken)

            Dim newCatch As CatchBlockSyntax = SyntaxFactory.CatchBlock(
            SyntaxFactory.CatchStatement(
                SyntaxFactory.IdentifierName("ex"),
                SyntaxFactory.SimpleAsClause(SyntaxFactory.IdentifierName(NameOf(Exception))),
                Nothing)).
                WithStatements(catchStatement.Statements).
                WithLeadingTrivia(catchStatement.GetLeadingTrivia).
                WithTrailingTrivia(catchStatement.GetTrailingTrivia).
                WithAdditionalAnnotations(Formatter.Annotation)

            Dim newRoot As SyntaxNode = root.ReplaceNode(catchStatement, newCatch)
            Dim newDoc As Document = document.WithSyntaxRoot(newRoot)
            Return newDoc
        End Function

    End Class

End Namespace

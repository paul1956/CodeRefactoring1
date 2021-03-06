﻿' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeActions
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Usage

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(MustInheritClassShouldNotHavePublicConstructorsCodeFixProvider)), [Shared]>
    Public Class MustInheritClassShouldNotHavePublicConstructorsCodeFixProvider
        Inherits CodeFixProvider

        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(AbstractClassShouldNotHavePublicCtorsDiagnosticId)

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

        Public NotOverridable Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim diag As Diagnostic = context.Diagnostics.First()
            context.RegisterCodeFix(CodeAction.Create("Use 'Friend' instead of 'Public'", Function(c As CancellationToken) Me.ReplacePublicWithProtectedAsync(context.Document, diag, c), NameOf(MustInheritClassShouldNotHavePublicConstructorsCodeFixProvider)), diag)
            Return Task.FromResult(0)
        End Function

    End Class

End Namespace

' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Formatting

Namespace Refactoring

    Public MustInherit Class BaseAllowMembersOrderingCodeFixProvider
        Inherits CodeFixProvider

        Protected Sub New(codeActionDescription As String)
            Me.CodeActionDescription = codeActionDescription
        End Sub

        Protected ReadOnly Property CodeActionDescription As String
        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(AllowMembersOrderingAnalyzer.Id)

        Private Shared Function TryReplaceTypeMembers(typeBlock As TypeBlockSyntax, membersDeclaration As IEnumerable(Of DeclarationStatementSyntax), sortedMembers As IEnumerable(Of DeclarationStatementSyntax), ByRef orderedType As TypeBlockSyntax) As Boolean
            Dim sortedMembersQueue As New Queue(Of DeclarationStatementSyntax)(sortedMembers)
            Dim orderChanged As Boolean = False
            orderedType = typeBlock.ReplaceNodes(membersDeclaration,
                                                 Function(original As DeclarationStatementSyntax, rewritten As DeclarationStatementSyntax)
                                                     Dim newMember As DeclarationStatementSyntax = sortedMembersQueue.Dequeue()
                                                     If Not orderChanged And Not original.Equals(newMember) Then
                                                         orderChanged = True
                                                     End If
                                                     Return newMember
                                                 End Function)
            Return orderChanged
        End Function

        Private Async Function AllowMembersOrderingAsync(document As Document, typeBlock As TypeBlockSyntax, cancellationToken As CancellationToken) As Task(Of Document)
            Dim membersDeclaration As IEnumerable(Of DeclarationStatementSyntax) = typeBlock.Members.OfType(Of DeclarationStatementSyntax)
            Dim root As CompilationUnitSyntax = DirectCast(Await document.GetSyntaxRootAsync(cancellationToken), CompilationUnitSyntax)

            Dim newTypeBlock As TypeBlockSyntax = Nothing
            Dim orderChanged As Boolean = TryReplaceTypeMembers(typeBlock, membersDeclaration, membersDeclaration.OrderBy(Function(member As DeclarationStatementSyntax) member, GetMemberDeclarationComparer(document, cancellationToken)), newTypeBlock)
            If Not orderChanged Then Return Nothing

            Dim newDocument As Document = document.WithSyntaxRoot(root.
                                                      ReplaceNode(typeBlock, newTypeBlock).
                                                      WithAdditionalAnnotations(Formatter.Annotation))
            Return newDocument
        End Function

        Protected MustOverride Function GetMemberDeclarationComparer(document As Document, cancellationToken As CancellationToken) As IComparer(Of DeclarationStatementSyntax)

        Public Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Overrides Async Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim root As SyntaxNode = Await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(False)
            Dim diagnostic As Diagnostic = context.Diagnostics.First()
            Dim diagnosticSpan As TextSpan = diagnostic.Location.SourceSpan

            Dim typeBlock As TypeBlockSyntax = DirectCast(root.FindToken(diagnosticSpan.Start).Parent.FirstAncestorOrSelfOfType(GetType(ClassBlockSyntax), GetType(StructureBlockSyntax), GetType(ModuleBlockSyntax)), TypeBlockSyntax)
            Dim newDocument As Document = Await AllowMembersOrderingAsync(context.Document, typeBlock, context.CancellationToken)
            If newDocument IsNot Nothing Then
                context.RegisterCodeFix(action:=CodeAction.Create(title:=String.Format(CodeActionDescription, typeBlock), createChangedDocument:=Function(ct As CancellationToken)
                                                                                                                                                     Return Task.FromResult(result:=newDocument)
                                                                                                                                                 End Function, equivalenceKey:="BaseAllowMembersOrderingCodeFixProvider"), diagnostic:=diagnostic)
            End If
        End Function

    End Class

End Namespace

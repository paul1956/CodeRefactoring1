' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Formatting
Imports Microsoft.CodeAnalysis.Simplification
Imports Microsoft.CodeAnalysis.VisualBasic.VisualBasicExtensions

Namespace Usage

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(DisposableFieldNotDisposedCodeFixProvider)), Composition.Shared>
    Public Class DisposableFieldNotDisposedCodeFixProvider
        Inherits CodeFixProvider

        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(DisposableFieldNotDisposed_CreatedDiagnosticId, DisposableFieldNotDisposed_ReturnedDiagnosticId)

        Private Shared Function AddDisposeDeclarationToDisposeMethod(variableDeclarator As VariableDeclaratorSyntax, type As ClassBlockSyntax, typeSymbol As INamedTypeSymbol) As TypeBlockSyntax
            Dim disposableMethod As IMethodSymbol = typeSymbol.GetMembers("Dispose").OfType(Of IMethodSymbol).FirstOrDefault(Function(d As IMethodSymbol) d.Arity = 0)
            Dim disposeStatment As StatementSyntax = SyntaxFactory.ParseExecutableStatement(variableDeclarator.Names.First().ToString() & ".Dispose()").
                WithTrailingTrivia(SyntaxFactory.CarriageReturnLineFeed)

            Dim newType As TypeBlockSyntax
            If disposableMethod Is Nothing Then
                Dim disposeMethod As MethodBlockSyntax = SyntaxFactory.SubBlock(SyntaxFactory.SubStatement("Dispose").
                    WithParameterList(SyntaxFactory.ParameterList()).
                    WithImplementsClause(SyntaxFactory.ImplementsClause(SyntaxFactory.QualifiedName(SyntaxFactory.IdentifierName(NameOf(IDisposable)), SyntaxFactory.IdentifierName("Dispose")))).
                    WithModifiers(SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.PublicKeyword))),
                    New SyntaxList(Of StatementSyntax)().Add(disposeStatment)).
                    WithAdditionalAnnotations(Formatter.Annotation)

                newType = type.AddMembers(disposeMethod)
            Else
                Dim existingDisposeMethod As MethodBlockSyntax = TryCast(disposableMethod.DeclaringSyntaxReferences.FirstOrDefault?.GetSyntax().Parent, MethodBlockSyntax)
                If type.Members.Contains(existingDisposeMethod) Then
                    Dim newDisposeMethod As MethodBlockSyntax = existingDisposeMethod.AddStatements(disposeStatment).
                    WithAdditionalAnnotations(Formatter.Annotation)

                    ' Ensure the dispose method includes the implements clause
                    Dim disposeStatement As MethodStatementSyntax = newDisposeMethod.SubOrFunctionStatement
                    Dim disposeStatementTrailingTrivia As SyntaxTriviaList = disposeStatement.GetTrailingTrivia()
                    If disposeStatement.ImplementsClause Is Nothing Then
                        disposeStatement = disposeStatement.
                        WithoutTrailingTrivia().
                        WithImplementsClause(SyntaxFactory.ImplementsClause(SyntaxFactory.QualifiedName(SyntaxFactory.IdentifierName(NameOf(IDisposable)), SyntaxFactory.IdentifierName("Dispose")))).
                        NormalizeWhitespace(" ").
                        WithTrailingTrivia(disposeStatementTrailingTrivia).
                        WithAdditionalAnnotations(Formatter.Annotation)

                        newDisposeMethod = newDisposeMethod.ReplaceNode(newDisposeMethod.SubOrFunctionStatement, disposeStatement)
                    End If
                    newType = type.ReplaceNode(existingDisposeMethod, newDisposeMethod)
                Else
                    Dim fieldDeclaration As SyntaxNode = variableDeclarator.Parent
                    Dim newFieldDeclaration As SyntaxNode = fieldDeclaration.
                        WithTrailingTrivia(SyntaxFactory.CommentTrivia("' Add " & disposeStatment.ToString() & " to the Dispose method on the partial file." & vbCrLf))

                    newType = type.ReplaceNode(fieldDeclaration, newFieldDeclaration)
                End If
            End If
            Return newType
        End Function

        Private Shared Function AddIDisposableImplementationToType(type As ClassBlockSyntax, typeSymbol As INamedTypeSymbol) As ClassBlockSyntax
            Dim iDisposableInterface As INamedTypeSymbol = typeSymbol.AllInterfaces.FirstOrDefault(Function(i As INamedTypeSymbol) i.ToString.EndsWith(NameOf(IDisposable)))
            If iDisposableInterface IsNot Nothing Then Return type
            Dim implementIdisposable As ImplementsStatementSyntax = SyntaxFactory.ImplementsStatement(SyntaxFactory.ParseName("System.IDisposable")).
                WithTrailingTrivia(SyntaxFactory.CarriageReturnLineFeed).
                WithAdditionalAnnotations(Simplifier.Annotation)

            Dim newImplementsList As SyntaxList(Of ImplementsStatementSyntax) = If(Not type.Implements.Any(),
            type.Implements.Add(implementIdisposable),
            New SyntaxList(Of ImplementsStatementSyntax)().
                              Add(implementIdisposable))

            Dim newType As ClassBlockSyntax = type.WithImplements(newImplementsList).
            WithAdditionalAnnotations(Formatter.Annotation)
            Return newType
        End Function

        Private Async Function DisposeField(document As Document, diagnostic As Diagnostic, cancellationToken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim span As TextSpan = diagnostic.Location.SourceSpan
            Dim variableDeclarator As VariableDeclaratorSyntax = root.FindToken(span.Start).Parent.FirstAncestorOrSelf(Of VariableDeclaratorSyntax)()

            Dim semanticModel As SemanticModel = Await document.GetSemanticModelAsync(cancellationToken)
            Dim type As ClassBlockSyntax = variableDeclarator.FirstAncestorOrSelf(Of ClassBlockSyntax)
            Dim typeSymbol As INamedTypeSymbol = semanticModel.GetDeclaredSymbol(type)
            Dim newTypeImplementingIDisposable As ClassBlockSyntax = AddIDisposableImplementationToType(type, typeSymbol)
            Dim newTypeWithDisposeMethod As TypeBlockSyntax = AddDisposeDeclarationToDisposeMethod(variableDeclarator, newTypeImplementingIDisposable, typeSymbol)
            Dim newRoot As SyntaxNode = root.ReplaceNode(type, newTypeWithDisposeMethod)
            Dim newDocument As Document = document.WithSyntaxRoot(newRoot)
            Return newDocument
        End Function

        Public Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim diagnostic As Diagnostic = context.Diagnostics.First
            context.RegisterCodeFix(CodeAction.Create($"Dispose field '{diagnostic.Properties!variableIdentifier}",
                                              Function(c As CancellationToken) DisposeField(context.Document, diagnostic, c),
                                              NameOf(DisposableFieldNotDisposedCodeFixProvider)),
                            diagnostic)
            Return Task.FromResult(0)
        End Function

    End Class

End Namespace

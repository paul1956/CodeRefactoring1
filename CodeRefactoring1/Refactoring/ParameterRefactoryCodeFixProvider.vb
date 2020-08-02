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
Imports Microsoft.CodeAnalysis.Text
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Refactoring

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(ParameterRefactoryCodeFixProvider)), [Shared]>
    Public Class ParameterRefactoryCodeFixProvider
        Inherits CodeFixProvider

        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(ParameterRefactoryDiagnosticId)

        Private Shared Function FirstLetterToLower(input As String) As String
            Return String.Concat(input.Replace(input(0).ToString(), input(0).ToString().ToLower()))
        End Function

        Private Shared Function FirstLetterToUpper(input As String) As String
            Return String.Concat(input.Replace(input(0).ToString(), input(0).ToString().ToUpper()))
        End Function

        Private Shared Function NewClassFactory(className As String, oldClass As ClassBlockSyntax, oldMethod As MethodBlockSyntax) As ClassBlockSyntax
            Dim newParameter As ParameterSyntax = SyntaxFactory.Parameter(identifier:=SyntaxFactory.ModifiedIdentifier(FirstLetterToLower(className))).
                WithAsClause(SyntaxFactory.SimpleAsClause(SyntaxFactory.IdentifierName(className))).
                NormalizeWhitespace(" ")

            Dim parameters As ParameterListSyntax = SyntaxFactory.ParameterList(SyntaxFactory.SeparatedList(Of ParameterSyntax).Add(newParameter)).
                WithAdditionalAnnotations(Formatter.Annotation)

            Dim methodStatement As MethodStatementSyntax = If(oldMethod.Kind = SyntaxKind.SubBlock,
                SyntaxFactory.SubStatement(oldMethod.SubOrFunctionStatement.Identifier.Text).
                    WithModifiers(oldMethod.SubOrFunctionStatement.Modifiers).
                    WithParameterList(parameters).
                    WithTrailingTrivia(SyntaxFactory.CarriageReturnLineFeed).
                    WithAdditionalAnnotations(Formatter.Annotation),
                SyntaxFactory.FunctionStatement(oldMethod.SubOrFunctionStatement.Identifier.Text).
                    WithModifiers(oldMethod.SubOrFunctionStatement.Modifiers).
                    WithParameterList(parameters).
                    WithAsClause(oldMethod.SubOrFunctionStatement.AsClause).
                    WithTrailingTrivia(SyntaxFactory.CarriageReturnLineFeed).
                    WithAdditionalAnnotations(Formatter.Annotation))

            Dim newMethod As MethodBlockSyntax = SyntaxFactory.MethodBlock(oldMethod.Kind, methodStatement, oldMethod.Statements, oldMethod.EndSubOrFunctionStatement).
                WithAdditionalAnnotations(Formatter.Annotation)
            Dim newClass As ClassBlockSyntax = oldClass.ReplaceNode(oldMethod, newMethod)
            Return newClass
        End Function

        Private Shared Function NewClassParameterFactory(newClassName As String, properties As List(Of PropertyStatementSyntax)) As ClassBlockSyntax
            Dim declaration As ClassStatementSyntax = SyntaxFactory.ClassStatement(newClassName).
                WithModifiers(SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.PublicKeyword)))

            Return SyntaxFactory.ClassBlock(declaration).
                WithMembers(SyntaxFactory.List(Of StatementSyntax)(properties)).
                WithTrailingTrivia(SyntaxFactory.CarriageReturnLineFeed).
                WithAdditionalAnnotations(Formatter.Annotation)
        End Function

        Private Shared Function NewCompilationFactory(oldCompilation As CompilationUnitSyntax, oldClass As ClassBlockSyntax, oldMethod As MethodBlockSyntax) As CompilationUnitSyntax
            Dim newNamespace As CompilationUnitSyntax = oldCompilation
            Dim className As String = "NewClass" & oldMethod.SubOrFunctionStatement.Identifier.Text
            Dim oldMemberNamespace As StatementSyntax = oldCompilation.Members.FirstOrDefault(Function(member As StatementSyntax) member.Equals(oldClass))
            newNamespace = oldCompilation.ReplaceNode(oldMemberNamespace, NewClassFactory(className, oldClass, oldMethod))
            Dim newParameterClass As ClassBlockSyntax = NewClassParameterFactory(className, NewPropertyClassFactory(oldMethod))

            Return newNamespace.WithMembers(newNamespace.Members.Add(newParameterClass)).
                WithAdditionalAnnotations(Formatter.Annotation)
        End Function

        Private Shared Function NewNamespaceFactory(oldNamespace As NamespaceBlockSyntax, oldClass As ClassBlockSyntax, oldMethod As MethodBlockSyntax) As NamespaceBlockSyntax
            Dim newNamespace As NamespaceBlockSyntax = oldNamespace
            Dim className As String = "NewClass" & oldMethod.SubOrFunctionStatement.Identifier.Text
            Dim memberNamespaceOld As StatementSyntax = oldNamespace.Members.FirstOrDefault(Function(member As StatementSyntax) member.Equals(oldClass))
            newNamespace = oldNamespace.ReplaceNode(memberNamespaceOld, NewClassFactory(className, oldClass, oldMethod))
            Dim newParameterClass As ClassBlockSyntax = NewClassParameterFactory(className, NewPropertyClassFactory(oldMethod))
            newNamespace = newNamespace.
                WithMembers(newNamespace.Members.Add(newParameterClass)).
                WithAdditionalAnnotations(Formatter.Annotation)
            Return newNamespace
        End Function

        Private Shared Function NewPropertyClassFactory(methodOld As MethodBlockSyntax) As List(Of PropertyStatementSyntax)
            Dim properties As List(Of PropertyStatementSyntax) = New List(Of PropertyStatementSyntax)
            For Each param As ParameterSyntax In methodOld.SubOrFunctionStatement.ParameterList.Parameters
                Dim newProperty As PropertyStatementSyntax = SyntaxFactory.PropertyStatement(FirstLetterToUpper(param.Identifier.GetText().ToString())).
                    WithModifiers(SyntaxFactory.TokenList(SyntaxFactory.Token(SyntaxKind.PublicKeyword))).
                    WithAsClause(param.AsClause).
                    WithAdditionalAnnotations(Formatter.Annotation)

                properties.Add(newProperty)
            Next
            Return properties
        End Function

        Private Async Function NewClassAsync(document As Document, diagnostic As Diagnostic, cancellationToken As CancellationToken) As Task(Of Document)
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim diagnosticSpan As TextSpan = diagnostic.Location.SourceSpan
            Dim declarationClass As ClassBlockSyntax = root.FindToken(diagnosticSpan.Start).Parent.FirstAncestorOfType(Of ClassBlockSyntax)
            Dim declarationNamespace As NamespaceBlockSyntax = root.FindToken(diagnosticSpan.Start).Parent.FirstAncestorOfType(Of NamespaceBlockSyntax)
            Dim declarationMethod As MethodBlockSyntax = root.FindToken(diagnosticSpan.Start).Parent.FirstAncestorOfType(Of MethodBlockSyntax)

            Dim newRootParameter As SyntaxNode
            If declarationNamespace Is Nothing Then
                Dim newCompilation As CompilationUnitSyntax = NewCompilationFactory(DirectCast(declarationClass.Parent, CompilationUnitSyntax), declarationClass, declarationMethod)
                newRootParameter = root.ReplaceNode(declarationClass.Parent, newCompilation)
                Return document.WithSyntaxRoot(newRootParameter)
            End If
            Dim newNamespace As NamespaceBlockSyntax = NewNamespaceFactory(declarationNamespace, declarationClass, declarationMethod)
            newRootParameter = root.ReplaceNode(declarationNamespace, newNamespace)
            Return document.WithSyntaxRoot(newRootParameter)
        End Function

        Public Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim diagnostic As Diagnostic = context.Diagnostics.First()
            context.RegisterCodeFix(CodeAction.Create("Change to new Class", Function(c As CancellationToken) Me.NewClassAsync(context.Document, diagnostic, c), NameOf(ParameterRefactoryCodeFixProvider)), diagnostic)
            Return Task.FromResult(0)
        End Function

    End Class

End Namespace

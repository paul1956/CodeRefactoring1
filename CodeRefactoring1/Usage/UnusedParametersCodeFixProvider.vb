' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.FindSymbols

Namespace Usage

    Public Structure DocumentIdAndRoot
        Friend DocumentId As DocumentId
        Friend Root As SyntaxNode

        Public Shared Operator <>(left As DocumentIdAndRoot, right As DocumentIdAndRoot) As Boolean
            Return Not left = right
        End Operator

        Public Shared Operator =(left As DocumentIdAndRoot, right As DocumentIdAndRoot) As Boolean
            Return left.Equals(right)
        End Operator

        Public Overrides Function Equals(obj As Object) As Boolean
            Throw New NotImplementedException()
        End Function

        Public Overrides Function GetHashCode() As Integer
            Throw New NotImplementedException()
        End Function

    End Structure

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(UnusedParametersCodeFixProvider)), Composition.Shared>
    Public Class UnusedParametersCodeFixProvider
        Inherits CodeFixProvider

        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String) = ImmutableArray.Create(UnusedParametersDiagnosticId)

        Private Shared Async Function RemoveParameterAsync(document As Document, diagnostic As Diagnostic, cancellationToken As CancellationToken) As Task(Of Solution)
            Dim solution As Solution = document.Project.Solution
            Dim newSolution As Solution = solution
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(cancellationToken).ConfigureAwait(False)
            Dim parameter As ParameterSyntax = root.FindToken(diagnostic.Location.SourceSpan.Start).Parent.FirstAncestorOrSelf(Of ParameterSyntax)
            Dim docs As List(Of DocumentIdAndRoot) = Await RemoveParameterAsync(document, parameter, root, cancellationToken)
            For Each doc As DocumentIdAndRoot In docs
                newSolution = newSolution.WithDocumentSyntaxRoot(doc.DocumentId, doc.Root)
            Next
            Return newSolution
        End Function

        Public Shared Async Function RemoveParameterAsync(document As Document, parameter As ParameterSyntax, root As SyntaxNode, cancellationToken As CancellationToken) As Task(Of List(Of DocumentIdAndRoot))
            Dim solution As Solution = document.Project.Solution
            Dim parameterList As ParameterListSyntax = DirectCast(parameter.Parent, ParameterListSyntax)
            Dim parameterPosition As Integer = parameterList.Parameters.IndexOf(parameter)
            Dim newParameterList As ParameterListSyntax = parameterList.WithParameters(parameterList.Parameters.Remove(parameter))
            Dim foundDocument As Boolean = False
            Dim semanticModel As SemanticModel = Await document.GetSemanticModelAsync(cancellationToken).ConfigureAwait(False)
            Dim method As SyntaxNode = parameter.FirstAncestorOfType(GetType(SubNewStatementSyntax), GetType(MethodBlockSyntax))
            Dim methodSymbol As ISymbol = semanticModel.GetDeclaredSymbol(method)
            Dim references As IEnumerable(Of FindSymbols.ReferencedSymbol) = Await SymbolFinder.FindReferencesAsync(methodSymbol, solution, cancellationToken)
            Dim documentGroups As IEnumerable(Of IGrouping(Of Document, FindSymbols.ReferenceLocation)) = references.SelectMany(Function(r) r.Locations).GroupBy(Function(loc) loc.Document)
            Dim docs As List(Of DocumentIdAndRoot) = New List(Of DocumentIdAndRoot)
            For Each documentGroup As IGrouping(Of Document, FindSymbols.ReferenceLocation) In documentGroups
                Dim referencingDocument As Document = documentGroup.Key
                Dim locRoot As SyntaxNode
                Dim locSemanticModel As SemanticModel
                Dim replacingArgs As Dictionary(Of SyntaxNode, SyntaxNode) = New Dictionary(Of SyntaxNode, SyntaxNode)
                If referencingDocument.Equals(document) Then
                    locSemanticModel = semanticModel
                    locRoot = root
                    replacingArgs.Add(parameterList, newParameterList)
                    foundDocument = True
                Else
                    locSemanticModel = Await referencingDocument.GetSemanticModelAsync(cancellationToken)
                    locRoot = Await locSemanticModel.SyntaxTree.GetRootAsync(cancellationToken)
                End If
                For Each loc As ReferenceLocation In documentGroup
                    Dim methodIdentifier As SyntaxNode = locRoot.FindNode(loc.Location.SourceSpan)
                    Dim objectCreation As ObjectCreationExpressionSyntax = TryCast(methodIdentifier.Parent, ObjectCreationExpressionSyntax)
                    Dim arguments As ArgumentListSyntax = If(objectCreation IsNot Nothing,
                        objectCreation.ArgumentList,
                        methodIdentifier.FirstAncestorOfType(Of InvocationExpressionSyntax).ArgumentList)

                    ' Attempt to find the parameter as a named argument.  Named arguments can only appear once in the argument list.
                    Dim namedArg As SimpleArgumentSyntax = arguments.Arguments.Where(Function(arg) arg.IsNamed).OfType(Of SimpleArgumentSyntax).SingleOrDefault(Function(arg) arg.NameColonEquals.Name.Identifier.Text = parameter.Identifier.Identifier.Text)
                    If namedArg IsNot Nothing Then
                        Dim newArguments As ArgumentListSyntax = arguments.WithArguments(arguments.Arguments.Remove(namedArg))
                        replacingArgs.Add(arguments, newArguments)
                    ElseIf parameter.Modifiers.Any(Function(m) m.IsKind(SyntaxKind.ParamArrayKeyword)) Then
                        Dim newArguments As ArgumentListSyntax = arguments
                        While parameterPosition < newArguments.Arguments.Count
                            newArguments = newArguments.WithArguments(newArguments.Arguments.RemoveAt(parameterPosition))
                        End While
                        replacingArgs.Add(arguments, newArguments)
                    ElseIf parameterPosition < arguments.Arguments.Count Then
                        Dim newArguments As ArgumentListSyntax = arguments.WithArguments(arguments.Arguments.RemoveAt(parameterPosition))
                        replacingArgs.Add(arguments, newArguments)
                    End If
                Next
                Dim newLocRoot As SyntaxNode = locRoot.ReplaceNodes(replacingArgs.Keys, Function(original, rewritten) replacingArgs(original))
                docs.Add(New DocumentIdAndRoot With {.DocumentId = referencingDocument.Id, .Root = newLocRoot})
            Next
            If Not foundDocument Then
                Dim newRoot As SyntaxNode = root.ReplaceNode(parameterList, newParameterList)
                Dim newDocument As Document = document.WithSyntaxRoot(newRoot)
                docs.Add(New DocumentIdAndRoot With {.DocumentId = document.Id, .Root = newRoot})
            End If
            Return docs
        End Function

        Public Overrides Function GetFixAllProvider() As FixAllProvider
            Return UnusedParametersCodeFixAllProvider.Instance
        End Function

        Public Overrides Function RegisterCodeFixesAsync(context As CodeFixContext) As Task
            Dim Diagnostic As Diagnostic = context.Diagnostics.First()
            context.RegisterCodeFix(CodeAction.Create(
                                    String.Format("Remove unused parameter: '{0}'", Diagnostic.Properties("identifier")),
                                    Function(c) RemoveParameterAsync(context.Document, Diagnostic, c),
                                    NameOf(UnusedParametersCodeFixProvider)),
                                    Diagnostic)
            Return Task.FromResult(0)
        End Function

    End Class

End Namespace

' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports CodeRefactoring1.Usage.MethodAnalyzers

Imports Microsoft.CodeAnalysis.Diagnostics

Public Class MethodChecker

    Private ReadOnly _context As SyntaxNodeAnalysisContext
    Private ReadOnly _diagnosticDescriptor As DiagnosticDescriptor

    Public Sub New(context As SyntaxNodeAnalysisContext, diagnosticDescriptor As DiagnosticDescriptor)
        _context = context
        _diagnosticDescriptor = diagnosticDescriptor
    End Sub

    Public Sub AnalyzeConstructor(methodInformation As MethodInformation)
        If ConstructorNameNotFound(methodInformation) OrElse MethodFullNameNotFound(methodInformation.MethodFullDefinition) Then Exit Sub
        Dim argumentList As ArgumentListSyntax = TryCast(_context.Node, ObjectCreationExpressionSyntax).ArgumentList
        Dim arguments As List(Of Object) = GetArguments(argumentList)
        Execute(methodInformation, arguments, argumentList)
    End Sub

    Private Function ConstructorNameNotFound(methodInformation As MethodInformation) As Boolean
        Return AbreviatedConstructorNameNotFound(methodInformation) AndAlso QualifiedConstructorNameNotFound(methodInformation)
    End Function

    Private Function AbreviatedConstructorNameNotFound(methodInformation As MethodInformation) As Boolean
        Dim objectCreationExpressionSyntax As ObjectCreationExpressionSyntax = DirectCast(_context.Node, ObjectCreationExpressionSyntax)
        Dim identifier As IdentifierNameSyntax = TryCast(objectCreationExpressionSyntax.Type, IdentifierNameSyntax)
        Return identifier?.Identifier.ValueText <> methodInformation.MethodName
    End Function

    Private Function QualifiedConstructorNameNotFound(methodInformation As MethodInformation) As Boolean
        Dim objectCreationExpressionSyntax As ObjectCreationExpressionSyntax = DirectCast(_context.Node, ObjectCreationExpressionSyntax)
        Dim identifier As QualifiedNameSyntax = TryCast(objectCreationExpressionSyntax.Type, QualifiedNameSyntax)
        Return identifier?.Right.ToString() <> methodInformation.MethodName
    End Function

    Public Sub AnalyzeMethod(methodInformation As MethodInformation)
        If MethodNameNotFound(methodInformation) OrElse MethodFullNameNotFound(methodInformation.MethodFullDefinition) Then
            Exit Sub
        End If
        Dim argumentList As ArgumentListSyntax = DirectCast(_context.Node, InvocationExpressionSyntax).ArgumentList
        Dim arguments As List(Of Object) = GetArguments(argumentList)
        Execute(methodInformation, arguments, argumentList)
    End Sub

    Private Function MethodNameNotFound(methodInformation As MethodInformation) As Boolean
        Dim invocationExpression As InvocationExpressionSyntax = DirectCast(_context.Node, InvocationExpressionSyntax)
        Dim memberExpression As MemberAccessExpressionSyntax = TryCast(invocationExpression.Expression, MemberAccessExpressionSyntax)
        Return memberExpression?.Name?.Identifier.ValueText <> methodInformation.MethodName
    End Function

    Private Function MethodFullNameNotFound(methodDefinition As String) As Boolean
        Dim memberSymbol As ISymbol = _context.SemanticModel.GetSymbolInfo(_context.Node).Symbol
        Return memberSymbol?.ToString <> methodDefinition
    End Function

    Private Sub Execute(methodInformation As MethodInformation, arguments As List(Of Object), argumentList As ArgumentListSyntax)
        If Not argumentList.Arguments.Any Then
            Exit Sub
        End If
        Try
            methodInformation.MethodAction.Invoke(arguments)
        Catch ex As Exception
            While (ex.InnerException IsNot Nothing)
                ex = ex.InnerException
            End While
            Dim diag As Diagnostic = Diagnostic.Create(_diagnosticDescriptor, argumentList.Arguments(methodInformation.ArgumentIndex).GetLocation(), ex.Message)
            _context.ReportDiagnostic(diag)
        End Try
    End Sub

    Private Function GetArguments(argumentList As ArgumentListSyntax) As List(Of Object)
        Return argumentList.Arguments.
            Select(Function(a) a.GetExpression()).
            Select(Function(l) If(l Is Nothing, Nothing, _context.SemanticModel.GetConstantValue(l).Value)).
            ToList()
    End Function

End Class

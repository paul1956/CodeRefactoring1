' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Performance

    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class RemoveWhereWhenItIsPossibleAnalyzer
        Inherits DiagnosticAnalyzer

        Private Const Category As String = SupportedCategories.Performance
        Private Const Description As String = "When a LINQ operator supports a predicate parameter it should be used instead of using 'Where' followed by the operator"
        Private Const MessageFormat As String = "You can remove 'Where' moving the predicate to '{0}'."
        Private Const Title As String = "You should remove the 'Where' invocation when it is possible."
        Shared ReadOnly supportedMethods() As String = {"First", "FirstOrDefault", "Last", "LastOrDefault", "Any", "Single", "SingleOrDefault", "Count"}

        Private Shared ReadOnly Rule As New DiagnosticDescriptor(
                            RemoveWhereWhenItIsPossibleDiagnosticId,
                            Title,
                            MessageFormat,
                            Category,
                            DiagnosticSeverity.Warning,
                            isEnabledByDefault:=True,
                            Description,
                            helpLinkUri:=ForDiagnostic(RemoveWhereWhenItIsPossibleDiagnosticId),
                            Array.Empty(Of String))

        Public Shared ReadOnly Id As String = RemoveWhereWhenItIsPossibleDiagnosticId
        Public Overrides ReadOnly Property SupportedDiagnostics() As ImmutableArray(Of DiagnosticDescriptor) = ImmutableArray.Create(Rule)

        Private Sub AnalyzeNode(context As SyntaxNodeAnalysisContext)
            If (context.Node.IsGenerated()) Then Return
            Dim whereInvoke As InvocationExpressionSyntax = DirectCast(context.Node, InvocationExpressionSyntax)
            If GetNameOfTheInvokeMethod(whereInvoke) <> "Where" Then Exit Sub

            Dim nextMethodInvoke As InvocationExpressionSyntax = whereInvoke.Parent.FirstAncestorOrSelf(Of InvocationExpressionSyntax)()

            Dim candidate As String = GetNameOfTheInvokeMethod(nextMethodInvoke)
            If Not supportedMethods.Contains(candidate) Then Exit Sub

            If nextMethodInvoke.ArgumentList.Arguments.Any Then Return
            Dim props As ImmutableDictionary(Of String, String) = New Dictionary(Of String, String) From {{"methodName", candidate}}.ToImmutableDictionary()
            Dim diag As Diagnostic = Diagnostic.Create(Rule, GetNameExpressionOfTheInvokedMethod(whereInvoke).GetLocation(), props, candidate)

            context.ReportDiagnostic(diag)
        End Sub

        Friend Shared Function GetNameExpressionOfTheInvokedMethod(invoke As InvocationExpressionSyntax) As SimpleNameSyntax
            If invoke Is Nothing Then Return Nothing

            Dim memberAccess As MemberAccessExpressionSyntax = invoke.ChildNodes.
            OfType(Of MemberAccessExpressionSyntax)().
            FirstOrDefault()

            Return memberAccess?.Name

        End Function

        Friend Shared Function GetNameOfTheInvokeMethod(invoke As InvocationExpressionSyntax) As String
            If (invoke Is Nothing) Then Return Nothing
            Dim memberAccess As MemberAccessExpressionSyntax = invoke.ChildNodes.
            OfType(Of MemberAccessExpressionSyntax).
            FirstOrDefault()

            Return GetNameExpressionOfTheInvokedMethod(invoke)?.ToString()
        End Function

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf AnalyzeNode, SyntaxKind.InvocationExpression)
        End Sub

    End Class

End Namespace

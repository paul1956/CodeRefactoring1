' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Refactoring

    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class ChangeAnyToAllAnalyzer
        Inherits DiagnosticAnalyzer

        Private Const Category As String = SupportedCategories.Refactoring
        Private Const MessageAll As String = "Change All to Any"
        Private Const MessageAny As String = "Change Any to All"
        Private Const TitleAll As String = MessageAll
        Private Const TitleAny As String = MessageAny

        Protected Shared RuleAll As New DiagnosticDescriptor(
                                ChangeAllToAnyDiagnosticId,
                                TitleAll,
                                MessageAll,
                                Category,
                                DiagnosticSeverity.Hidden,
                                isEnabledByDefault:=True,
                                description:="",
                                helpLinkUri:=ForDiagnostic(ChangeAllToAnyDiagnosticId),
                                Array.Empty(Of String))

        Protected Shared RuleAny As New DiagnosticDescriptor(
                                ChangeAnyToAllDiagnosticId,
                                TitleAny,
                                MessageAny,
                                Category,
                                DiagnosticSeverity.Hidden,
                                isEnabledByDefault:=True,
                                description:="",
                                helpLinkUri:=ForDiagnostic(ChangeAnyToAllDiagnosticId),
                                Array.Empty(Of String))

        Public Shared ReadOnly AllName As IdentifierNameSyntax = SyntaxFactory.IdentifierName(NameOf(Enumerable.All))

        Public Shared ReadOnly AnyName As IdentifierNameSyntax = SyntaxFactory.IdentifierName(NameOf(Enumerable.Any))

        Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(RuleAll, RuleAny)
            End Get
        End Property

        Private Shared Function GetCorrespondingDiagnostic(model As SemanticModel, invocation As InvocationExpressionSyntax) As DiagnosticDescriptor
            Dim methodName As String = TryCast(invocation?.Expression, MemberAccessExpressionSyntax)?.Name?.ToString()
            Dim nameToCheck As IdentifierNameSyntax = If(methodName = NameOf(Enumerable.Any), AllName,
                If(methodName = NameOf(Enumerable.All), AnyName, Nothing))
            If nameToCheck Is Nothing Then Return Nothing
            Dim invocationSymbol As IMethodSymbol = TryCast(model.GetSymbolInfo(invocation).Symbol, IMethodSymbol)
            If invocationSymbol?.Parameters.Length <> 1 Then Return Nothing
            If Not IsLambdaWithEmptyOrSimpleBody(invocation) Then Return Nothing
            Dim otherInvocation As InvocationExpressionSyntax = invocation.WithExpression(DirectCast(invocation.Expression, MemberAccessExpressionSyntax).WithName(nameToCheck))
            Dim otherInvocationSymbol As SymbolInfo = model.GetSpeculativeSymbolInfo(invocation.SpanStart, otherInvocation, SpeculativeBindingOption.BindAsExpression)
            If otherInvocationSymbol.Symbol Is Nothing Then Return Nothing
            Return If(methodName = NameOf(Enumerable.Any), RuleAny, RuleAll)
        End Function

        Private Shared Function IsLambdaWithEmptyOrSimpleBody(invocation As InvocationExpressionSyntax) As Boolean
            Dim arg As ArgumentSyntax = invocation.ArgumentList?.Arguments.FirstOrDefault()
            If arg Is Nothing Then Return True ' Empty body
            Return arg.GetExpression().Kind = SyntaxKind.SingleLineFunctionLambdaExpression
        End Function

        Private Sub AnalyzeInvocation(context As SyntaxNodeAnalysisContext)
            If context.Node.IsGenerated() Then Exit Sub
            Dim invocation As InvocationExpressionSyntax = DirectCast(context.Node, InvocationExpressionSyntax)
            If invocation.Parent.IsKind(SyntaxKind.ExpressionStatement) Then Exit Sub
            Dim diagnosticToRaise As DiagnosticDescriptor = GetCorrespondingDiagnostic(context.SemanticModel, invocation)
            If diagnosticToRaise Is Nothing Then Exit Sub
            Dim diag As Diagnostic = Diagnostic.Create(diagnosticToRaise, DirectCast(invocation.Expression, MemberAccessExpressionSyntax).Name.GetLocation())
            context.ReportDiagnostic(diag)
        End Sub

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf AnalyzeInvocation, SyntaxKind.InvocationExpression)
        End Sub

    End Class

End Namespace

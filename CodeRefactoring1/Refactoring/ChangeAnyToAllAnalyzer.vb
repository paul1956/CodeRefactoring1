Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Refactoring
    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class ChangeAnyToAllAnalyzer
        Inherits DiagnosticAnalyzer

        Friend Const MessageAny As String = "Change Any to All"
        Friend Const MessageAll As String = "Change All to Any"
        Friend Const TitleAny As String = MessageAny
        Friend Const TitleAll As String = MessageAll
        Friend Const Category As String = SupportedCategories.Refactoring

        Friend Shared RuleAny As New DiagnosticDescriptor(
            DiagnosticIds.ChangeAnyToAllDiagnosticId,
            TitleAny,
            MessageAny,
            Category,
            DiagnosticSeverity.Hidden,
            isEnabledByDefault:=True,
            helpLinkUri:=HelpLink.ForDiagnostic(DiagnosticIds.ChangeAnyToAllDiagnosticId))
        Friend Shared RuleAll As New DiagnosticDescriptor(
            DiagnosticIds.ChangeAllToAnyDiagnosticId,
            TitleAll,
            MessageAll,
            Category,
            DiagnosticSeverity.Hidden,
            isEnabledByDefault:=True,
            helpLinkUri:=HelpLink.ForDiagnostic(DiagnosticIds.ChangeAllToAnyDiagnosticId))

        Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(RuleAll, RuleAny)
            End Get
        End Property

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf Me.AnalyzeInvocation, SyntaxKind.InvocationExpression)
        End Sub

        Public Shared ReadOnly AllName As IdentifierNameSyntax = SyntaxFactory.IdentifierName(NameOf(Enumerable.All))
        Public Shared ReadOnly AnyName As IdentifierNameSyntax = SyntaxFactory.IdentifierName(NameOf(Enumerable.Any))

        Private Sub AnalyzeInvocation(context As SyntaxNodeAnalysisContext)
            If context.Node.IsGenerated() Then Exit Sub
            Dim invocation As InvocationExpressionSyntax = DirectCast(context.Node, InvocationExpressionSyntax)
            If invocation.Parent.IsKind(SyntaxKind.ExpressionStatement) Then Exit Sub
            Dim diagnosticToRaise As DiagnosticDescriptor = GetCorrespondingDiagnostic(context.SemanticModel, invocation)
            If diagnosticToRaise Is Nothing Then Exit Sub
            Dim diag As Diagnostic = Diagnostic.Create(diagnosticToRaise, DirectCast(invocation.Expression, MemberAccessExpressionSyntax).Name.GetLocation())
            context.ReportDiagnostic(diag)
        End Sub

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
    End Class
End Namespace

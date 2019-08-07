﻿Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Reliability
    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class UseConfigureAwaitFalseAnalyzer
        Inherits DiagnosticAnalyzer

        Friend Const Title As String = "Use ConfigureAwait(False) on awaited task."
        Friend Const MessageFormat As String = "Consider using ConfigureAwait(False) on the awaited task."
        Friend Const Category As String = SupportedCategories.Reliability

        Friend Shared Rule As New DiagnosticDescriptor(
        DiagnosticIds.UseConfigureAwaitFalseDiagnosticId,
        Title,
        MessageFormat,
        Category,
        DiagnosticSeverity.Hidden,
        isEnabledByDefault:=True,
        helpLinkUri:=HelpLink.ForDiagnostic(DiagnosticIds.UseConfigureAwaitFalseDiagnosticId))

        Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(Rule)
            End Get
        End Property

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf AnalyzeNode, SyntaxKind.AwaitExpression)
        End Sub

        Private Shared Sub AnalyzeNode(context As SyntaxNodeAnalysisContext)
            If (context.Node.IsGenerated()) Then Return
            Dim awaitExpression As AwaitExpressionSyntax = DirectCast(context.Node, AwaitExpressionSyntax)
            Dim awaitedExpression As ExpressionSyntax = awaitExpression.Expression
            If Not IsTask(awaitedExpression, context) Then Exit Sub

            Dim diag As Diagnostic = Diagnostic.Create(Rule, awaitExpression.GetLocation())
            context.ReportDiagnostic(diag)
        End Sub

        Private Shared Function IsTask(expression As ExpressionSyntax, context As SyntaxNodeAnalysisContext) As Boolean
            Dim type As INamedTypeSymbol = TryCast(context.SemanticModel.GetTypeInfo(expression).Type, INamedTypeSymbol)
            If type Is Nothing Then Return False
            Dim taskType As INamedTypeSymbol
            If type.IsGenericType Then
                type = type.ConstructedFrom
                taskType = context.SemanticModel.Compilation.GetTypeByMetadataName("System.Threading.Tasks.Task`1")
            Else
                taskType = context.SemanticModel.Compilation.GetTypeByMetadataName("System.Threading.Tasks.Task")
            End If
            Return type.Equals(taskType)
        End Function
    End Class
End Namespace
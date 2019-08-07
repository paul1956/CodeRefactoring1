Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Refactoring
    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class ParameterRefactoryAnalyzer
        Inherits DiagnosticAnalyzer

        Friend Const Title As String = "You should use a class."
        Friend Const MessageFormat As String = "When the method has more than three parameters, use a class."
        Friend Const Category As String = SupportedCategories.Refactoring

        Friend Shared Rule As New DiagnosticDescriptor(
        DiagnosticIds.ParameterRefactoryDiagnosticId,
        Title,
        MessageFormat,
        Category,
        DiagnosticSeverity.Hidden,
        isEnabledByDefault:=True,
        helpLinkUri:=HelpLink.ForDiagnostic(DiagnosticIds.ParameterRefactoryDiagnosticId)
    )

        Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(Rule)
            End Get
        End Property

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf Me.AnalyzeNode, SyntaxKind.SubBlock, SyntaxKind.FunctionBlock)
        End Sub

        Public Sub AnalyzeNode(context As SyntaxNodeAnalysisContext)
            If (context.Node.IsGenerated()) Then Return
            Dim method As MethodBlockSyntax = DirectCast(context.Node, MethodBlockSyntax)
            If method.SubOrFunctionStatement.Modifiers.Any(SyntaxKind.FriendKeyword) Then Exit Sub

            ' Check for extension method
            For Each attributeList As AttributeListSyntax In method.SubOrFunctionStatement.AttributeLists
                For Each attribute As AttributeSyntax In attributeList.Attributes
                    If attribute.Name.ToString().Contains("Extension") Then Exit Sub
                Next
            Next

            Dim contentParameter As ParameterListSyntax = method.SubOrFunctionStatement.ParameterList

            If contentParameter Is Nothing OrElse contentParameter.Parameters.Count <= 3 Then Exit Sub
            If method.Statements.Any() Then Exit Sub

            For Each parameter As ParameterSyntax In contentParameter.Parameters
                For Each modifier As SyntaxToken In parameter.Modifiers
                    If modifier.IsKind(SyntaxKind.ByRefKeyword) OrElse
                        modifier.IsKind(SyntaxKind.ParamArrayKeyword) Then
                        Exit Sub
                    End If
                Next
            Next

            Dim diag As Diagnostic = Diagnostic.Create(Rule, contentParameter.GetLocation())
            context.ReportDiagnostic(diag)
        End Sub
    End Class
End Namespace




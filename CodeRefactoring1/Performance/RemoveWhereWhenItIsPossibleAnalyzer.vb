Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On
Imports Microsoft.CodeAnalysis.Diagnostics
Imports System.Collections.Immutable

Namespace Performance
    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class RemoveWhereWhenItIsPossibleAnalyzer
        Inherits DiagnosticAnalyzer

        Public Shared ReadOnly Id As String = DiagnosticIds.RemoveWhereWhenItIsPossibleDiagnosticId
        Public Const Title As String = "You should remove the 'Where' invocation when it is possible."
        Public Const MessageFormat As String = "You can remove 'Where' moving the predicate to '{0}'."
        Public Const Category As String = SupportedCategories.Performance
        Public Const Description As String = "When a LINQ operator supports a predicate parameter it should be used instead of using 'Where' followed by the operator"
        Protected Shared Rule As DiagnosticDescriptor = New DiagnosticDescriptor(
            DiagnosticIds.RemoveWhereWhenItIsPossibleDiagnosticId,
            Title,
            MessageFormat,
            Category,
            DiagnosticSeverity.Warning,
            isEnabledByDefault:=True,
            description:=Description,
            helpLinkUri:=HelpLink.ForDiagnostic(DiagnosticIds.RemoveWhereWhenItIsPossibleDiagnosticId))

        Public Overrides ReadOnly Property SupportedDiagnostics() As ImmutableArray(Of DiagnosticDescriptor) = ImmutableArray.Create(Rule)

        Shared ReadOnly supportedMethods() As String = {"First", "FirstOrDefault", "Last", "LastOrDefault", "Any", "Single", "SingleOrDefault", "Count"}

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf Me.AnalyzeNode, SyntaxKind.InvocationExpression)
        End Sub

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

        Friend Shared Function GetNameOfTheInvokeMethod(invoke As InvocationExpressionSyntax) As String
            If (invoke Is Nothing) Then Return Nothing
            Dim memberAccess As VisualBasic.Syntax.MemberAccessExpressionSyntax = invoke.ChildNodes.
            OfType(Of MemberAccessExpressionSyntax).
            FirstOrDefault()

            Return GetNameExpressionOfTheInvokedMethod(invoke)?.ToString()
        End Function

        Friend Shared Function GetNameExpressionOfTheInvokedMethod(invoke As InvocationExpressionSyntax) As SimpleNameSyntax
            If invoke Is Nothing Then Return Nothing

            Dim memberAccess As VisualBasic.Syntax.MemberAccessExpressionSyntax = invoke.ChildNodes.
            OfType(Of MemberAccessExpressionSyntax)().
            FirstOrDefault()

            Return memberAccess?.Name

        End Function

    End Class
End Namespace
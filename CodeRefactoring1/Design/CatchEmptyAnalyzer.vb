Imports Microsoft.CodeAnalysis.Diagnostics
Imports System.Collections.Immutable

Namespace Design
    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class CatchEmptyAnalyzer
        Inherits DiagnosticAnalyzer

        Public Shared ReadOnly Id As String = DiagnosticIds.CatchEmptyDiagnosticId
        Public Const Title As String = "Your catch should include an Exception"
        Public Const MessageFormat As String = "{0}"
        Public Const Category As String = SupportedCategories.Design
        Protected Shared Rule As DiagnosticDescriptor = New DiagnosticDescriptor(
            DiagnosticIds.CatchEmptyDiagnosticId,
            Title,
            MessageFormat,
            Category,
            DiagnosticSeverity.Warning,
            isEnabledByDefault:=True,
            helpLinkUri:=HelpLink.ForDiagnostic(DiagnosticIds.CatchEmptyDiagnosticId))

        Public Overrides ReadOnly Property SupportedDiagnostics() As ImmutableArray(Of DiagnosticDescriptor) = ImmutableArray.Create(Rule)

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf Me.Analyzer, SyntaxKind.CatchStatement)
        End Sub

        Private Sub Analyzer(context As SyntaxNodeAnalysisContext)
            If (context.Node.IsGenerated()) Then Return
            Dim catchStatement As CatchStatementSyntax = DirectCast(context.Node, CatchStatementSyntax)
            If catchStatement Is Nothing Then Exit Sub

            If catchStatement.IdentifierName Is Nothing Then
                Dim diag As Diagnostic = Diagnostic.Create(Rule, catchStatement.GetLocation(), "Consider adding an Exception to the catch.")
                context.ReportDiagnostic(diag)
            End If
        End Sub

    End Class
End Namespace
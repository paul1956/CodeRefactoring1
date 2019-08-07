Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Design
    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class EmptyCatchBlockAnalyzer
        Inherits DiagnosticAnalyzer

        Public Const Title As String = "Catch block cannot be empty"
        Public Const MessageFormat As String = "{0}"
        Public Const Category As String = SupportedCategories.Design
        Public Const Description As String = "An empty catch block suppresses all errors and shouldn't be used.
If the error is expected, consider logging it or changing the control flow such that it is explicit."
        Protected Shared Rule As DiagnosticDescriptor = New DiagnosticDescriptor(
             DiagnosticIds.EmptyCatchBlockDiagnosticId,
            Title,
            MessageFormat,
            Category,
            DiagnosticSeverity.Warning,
            isEnabledByDefault:=True,
            description:=Description,
            helpLinkUri:=HelpLink.ForDiagnostic(DiagnosticIds.EmptyCatchBlockDiagnosticId))

        Public Overrides ReadOnly Property SupportedDiagnostics() As ImmutableArray(Of DiagnosticDescriptor) = ImmutableArray.Create(Rule)

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf Me.Analyzer, SyntaxKind.CatchBlock)
        End Sub

        Private Sub Analyzer(context As SyntaxNodeAnalysisContext)
            If (context.Node.IsGenerated()) Then Return
            Dim catchBlock As VisualBasic.Syntax.CatchBlockSyntax = DirectCast(context.Node, CatchBlockSyntax)
            If (catchBlock.Statements.Count <> 0) Then Exit Sub
            Dim diag As Diagnostic = Diagnostic.Create(Rule, catchBlock.GetLocation(), "Empty Catch Block.")
            context.ReportDiagnostic(diag)
        End Sub

    End Class
End Namespace

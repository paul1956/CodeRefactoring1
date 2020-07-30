Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.Diagnostics

<DiagnosticAnalyzer(LanguageNames.VisualBasic)>
Public Class ByValAnalyzer_FixAnalyzer
    Inherits DiagnosticAnalyzer

    Private Const Category As String = "Style"
    Private Shared ReadOnly Description As LocalizableString = New LocalizableResourceString(NameOf(My.Resources.RemoveByValDescription), My.Resources.ResourceManager, GetType(My.Resources.Resources))
    Private Shared ReadOnly MessageFormat As LocalizableString = New LocalizableResourceString(NameOf(My.Resources.RemoveByValMessageFormat), My.Resources.ResourceManager, GetType(My.Resources.Resources))
    Private Shared ReadOnly Title As LocalizableString = New LocalizableResourceString(NameOf(My.Resources.RemoveByValTitle), My.Resources.ResourceManager, GetType(My.Resources.Resources))

    Private Shared ReadOnly Rule As New DiagnosticDescriptor(
                                DiagnosticId,
                                Title,
                                MessageFormat,
                                Category,
                                DiagnosticSeverity.Info,
                                isEnabledByDefault:=True,
                                Description,
                                helpLinkUri:="",
                                Array.Empty(Of String)
                                )

    ' You can change these strings in the Resources.resx file. If you do not want your analyzer to be localize-able, you can use regular strings for Title and MessageFormat.
    ' See https://github.com/dotnet/roslyn/blob/master/docs/analyzers/Localizing%20Analyzers.md for more on localization

    Public Const DiagnosticId As String = "ByValAnalyzer_Fix"

    Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
        Get
            Return ImmutableArray.Create(Rule)
        End Get
    End Property

    Private Sub AnalyzeMethod(ByVal context As SyntaxNodeAnalysisContext)
        Dim SubFunctionStatement As MethodStatementSyntax = CType(context.Node, MethodStatementSyntax)
        For i As Integer = 0 To SubFunctionStatement.ParameterList.Parameters.Count - 1
            Dim Parameter As ParameterSyntax = SubFunctionStatement.ParameterList.Parameters(i)
            For Each Modifier As SyntaxToken In Parameter.Modifiers
                If Modifier.IsKind(SyntaxKind.ByValKeyword) Then
                    Dim diag As Diagnostic = Diagnostic.Create(Rule, Modifier.GetLocation, i)
                    context.ReportDiagnostic(diag)
                End If
            Next
        Next
    End Sub

    Public Overrides Sub Initialize(context As AnalysisContext)
        ' TODO: Consider registering other actions that act on syntax instead of or in addition to symbols
        ' See https://github.com/dotnet/roslyn/blob/master/docs/analyzers/Analyzer%20Actions%20Semantics.md for more information
        context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None)
        context.EnableConcurrentExecution()
        context.RegisterSyntaxNodeAction(AddressOf Me.AnalyzeMethod, SyntaxKind.FunctionStatement, SyntaxKind.SubStatement)
    End Sub

End Class
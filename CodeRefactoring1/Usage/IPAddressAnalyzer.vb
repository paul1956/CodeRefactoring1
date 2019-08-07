Imports System.Collections.Immutable
Imports System.Reflection
Imports CodeRefactoring1.Usage.MethodAnalyzers
Imports Microsoft.CodeAnalysis.Diagnostics

<DiagnosticAnalyzer(LanguageNames.VisualBasic)>
Public Class IPAddressAnalyzer
    Inherits DiagnosticAnalyzer

    Friend Const Title As String = "Your IP Address syntax is wrong."
    Friend Const MessageFormat As String = "{0}"
    Private Const Description As String = "This diagnostic checks the IP Address string and triggers if the parsing will fail by throwing an exception."

    Friend Shared Rule As New DiagnosticDescriptor(
        DiagnosticIds.IPAddressDiagnosticId,
        Title,
        MessageFormat,
        SupportedCategories.Usage,
        DiagnosticSeverity.Error,
        isEnabledByDefault:=True,
        description:=Description,
        helpLinkUri:=HelpLink.ForDiagnostic(DiagnosticIds.IPAddressDiagnosticId))

    Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
        Get
            Return ImmutableArray.Create(Rule)
        End Get
    End Property

    Public Overrides Sub Initialize(context As AnalysisContext)
        context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
        context.EnableConcurrentExecution()
        context.RegisterSyntaxNodeAction(AddressOf Me.Analyzer, SyntaxKind.InvocationExpression)
    End Sub

    Private Sub Analyzer(context As SyntaxNodeAnalysisContext)
        If (context.Node.IsGenerated()) Then Return
        Dim method As New MethodInformation("Parse",
                                            "Public Shared Overloads Function Parse(ipString As String) As System.Net.IPAddress",
                                            Sub(args) parseMethodInfo.Value.Invoke(Nothing, {args(0).ToString()}))
        Dim checker As MethodChecker = New MethodChecker(context, Rule)
        checker.AnalyzeMethod(method)
    End Sub

    Private Shared ReadOnly objectType As New Lazy(Of Type)(Function() System.Type.GetType("System.Net.IPAddress, System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"))
    Private Shared ReadOnly parseMethodInfo As New Lazy(Of MethodInfo)(Function() objectType.Value.GetRuntimeMethod("Parse", {GetType(String)}))
End Class

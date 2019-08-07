Imports Microsoft.CodeAnalysis.Diagnostics
Imports System.Collections.Immutable
Imports CodeRefactoring1.Usage.MethodAnalyzers

Namespace Usage
    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class UriAnalyzer
        Inherits DiagnosticAnalyzer
        Friend Const Title As String = "Your Uri syntax is wrong."
        Friend Const MessageFormat As String = "{0}"
        Friend Const Category As String = SupportedCategories.Usage

        Private Const Description As String = "This diagnostic checks the Uri string and triggers if the parsing fail " + "by throwing an exception."

        Friend Shared Rule As New DiagnosticDescriptor(DiagnosticIds.UriDiagnosticId, Title, MessageFormat, Category, DiagnosticSeverity.[Error], isEnabledByDefault:=True,
            description:=Description, helpLinkUri:=HelpLink.ForDiagnostic(DiagnosticIds.UriDiagnosticId))

        Public Overrides ReadOnly Property SupportedDiagnostics() As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(Rule)
            End Get
        End Property

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf Me.Analyzer, SyntaxKind.ObjectCreationExpression)
        End Sub

        Private Sub Analyzer(context As SyntaxNodeAnalysisContext)
            If context.Node.IsGenerated() Then Return
            Dim mainConstrutor As MethodInformation = New MethodInformation(NameOf(Uri),
                                                       "Public Overloads Sub New(uriString As String)",
                                                       Sub(args)
                                                           If args(0) Is Nothing Then Return
                                                           Dim a As Uri = New Uri(args(0).ToString())
                                                       End Sub)
            Dim constructorWithUriKind As MethodInformation = New MethodInformation(NameOf(Uri),
                                                               "Public Overloads Sub New(uriString As String, uriKind As System.UriKind)",
                                                               Sub(args)
                                                                   If args(0) Is Nothing Then Return
                                                                   Dim a As Uri = New Uri(args(0).ToString(), DirectCast(args(1), UriKind))
                                                               End Sub)
            Dim checker As MethodChecker = New MethodChecker(context, Rule)
            checker.AnalyzeConstructor(mainConstrutor)
            checker.AnalyzeConstructor(constructorWithUriKind)
        End Sub
    End Class
End Namespace
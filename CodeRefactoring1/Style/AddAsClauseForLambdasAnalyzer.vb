Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Style

    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class AddAsClauseForLambdasAnalyzer
        Inherits DiagnosticAnalyzer

        Private Const Category As String = SupportedCategories.Style
        Private Const Description As String = "Option Strict On requires all Lambda declarations to have an 'As' clause."
        Private Const MessageFormat As String = Description
        Private Const Title As String = Description
        Private Shared ReadOnly Rule As New DiagnosticDescriptor(
            DiagnosticIds.AddAsClauseForLambdaDiagnosticId,
            Title,
            MessageFormat,
            Category,
            DiagnosticSeverity.Error,
            isEnabledByDefault:=True,
            description:=Description,
            helpLinkUri:=ForDiagnostic(DiagnosticIds.AddAsClauseDiagnosticId)
        )

        Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(Rule)
            End Get
        End Property

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf AnalyzeLambdaVariable, SyntaxKind.FunctionLambdaHeader, SyntaxKind.SubLambdaHeader)
        End Sub

        Private Shared Sub AnalyzeLambdaVariable(context As SyntaxNodeAnalysisContext)
            Try
                Dim model As SemanticModel = context.SemanticModel
                If model.OptionStrict = OptionStrict.On AndAlso model.OptionInfer = True Then
                    Exit Sub
                End If

                Dim diag As Diagnostic = Nothing
                Select Case context.Node.Kind
                    Case SyntaxKind.FunctionLambdaHeader, SyntaxKind.SubLambdaHeader
                        Dim _LambdaHeaderSyntax As LambdaHeaderSyntax = DirectCast(context.Node, LambdaHeaderSyntax)
                        If _LambdaHeaderSyntax.ParameterList Is Nothing Then
                            Exit Sub
                        End If
                        For Each param As ParameterSyntax In _LambdaHeaderSyntax.ParameterList.Parameters
                            If param.AsClause Is Nothing Then
                                diag = Diagnostic.Create(Rule, param.GetLocation(), param.GetType.ToString)
                                context.ReportDiagnostic(diag)
                            End If
                        Next
                        Exit Sub
                    Case Else
                        Stop
                End Select
                context.ReportDiagnostic(diag)
            Catch ex As Exception When ex.HResult <> (New OperationCanceledException).HResult
                Stop
                Throw
            End Try
        End Sub
    End Class

End Namespace
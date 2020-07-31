' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Design

    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class CatchEmptyAnalyzer
        Inherits DiagnosticAnalyzer

        Private Const Category As String = SupportedCategories.Design

        Private Const MessageFormat As String = "{0}"

        Private Const Title As String = "Your catch should include an Exception"

        Protected Shared Rule As New DiagnosticDescriptor(
                            Id,
                            Title,
                            MessageFormat,
                            Category,
                            DiagnosticSeverity.Warning,
                            isEnabledByDefault:=True,
                            description:="",
                            helpLinkUri:=ForDiagnostic(Id),
                            Array.Empty(Of String)
                            )

        Public Const Id As String = CatchEmptyDiagnosticId
        Public Overrides ReadOnly Property SupportedDiagnostics() As ImmutableArray(Of DiagnosticDescriptor) = ImmutableArray.Create(Rule)

        Private Sub Analyzer(context As SyntaxNodeAnalysisContext)
            If (context.Node.IsGenerated()) Then Return
            Dim catchStatement As CatchStatementSyntax = DirectCast(context.Node, CatchStatementSyntax)
            If catchStatement Is Nothing Then Exit Sub

            If catchStatement.IdentifierName Is Nothing Then
                Dim diag As Diagnostic = Diagnostic.Create(Rule, catchStatement.GetLocation(), "Consider adding an Exception to the catch.")
                context.ReportDiagnostic(diag)
            End If
        End Sub

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf Analyzer, SyntaxKind.CatchStatement)
        End Sub

    End Class

End Namespace

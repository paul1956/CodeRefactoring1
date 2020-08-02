' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Diagnostics
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Namespace Design

    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class EmptyCatchBlockAnalyzer
        Inherits DiagnosticAnalyzer

        Private Const Category As String = SupportedCategories.Design

        Private Const Description As String = "An empty catch block suppresses all errors and shouldn't be used.
If the error is expected, consider logging it or changing the control flow such that it is explicit."

        Private Const MessageFormat As String = "{0}"
        Private Const Title As String = "Catch block cannot be empty"

        Protected Shared Rule As New DiagnosticDescriptor(
                            EmptyCatchBlockDiagnosticId,
                            Title,
                            MessageFormat,
                            Category,
                            DiagnosticSeverity.Warning,
                            isEnabledByDefault:=True,
                            Description,
                            helpLinkUri:=ForDiagnostic(EmptyCatchBlockDiagnosticId),
                            Array.Empty(Of String)
                            )

        Public Overrides ReadOnly Property SupportedDiagnostics() As ImmutableArray(Of DiagnosticDescriptor) = ImmutableArray.Create(Rule)

        Private Sub Analyzer(context As SyntaxNodeAnalysisContext)
            If context.Node.IsGenerated() Then Return
            Dim catchBlock As VisualBasic.Syntax.CatchBlockSyntax = DirectCast(context.Node, CatchBlockSyntax)
            If catchBlock.Statements.Count <> 0 Then Exit Sub
            Dim diag As Diagnostic = Diagnostic.Create(Rule, catchBlock.GetLocation(), "Empty Catch Block.")
            context.ReportDiagnostic(diag)
        End Sub

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf Me.Analyzer, SyntaxKind.CatchBlock)
        End Sub

    End Class

End Namespace

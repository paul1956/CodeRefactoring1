' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Design

    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class StaticConstructorExceptionAnalyzer
        Inherits DiagnosticAnalyzer

        Private Const Category As String = SupportedCategories.Design

        Private Const Description As String = "Static constructor are called before the first time a class is used but the caller doesn't control when exactly.
Exception thrown in this context forces callers to use 'try' block around any usage of the class and should be avoided."

        Private Const MessageFormat As String = "Don't throw exceptions inside static constructors."
        Private Const Title As String = "Don't throw exception inside static constructors."

        Protected Shared Rule As New DiagnosticDescriptor(
                            StaticConstructorExceptionDiagnosticId,
                            Title,
                            MessageFormat,
                            Category,
                            DiagnosticSeverity.Warning,
                            isEnabledByDefault:=True,
                            Description,
                            helpLinkUri:=ForDiagnostic(StaticConstructorExceptionDiagnosticId),
                            Array.Empty(Of String)
                            )

        Public Shared ReadOnly Id As String = StaticConstructorExceptionDiagnosticId
        Public Overrides ReadOnly Property SupportedDiagnostics() As ImmutableArray(Of DiagnosticDescriptor) = ImmutableArray.Create(Rule)

        Private Sub Analyzer(context As SyntaxNodeAnalysisContext)
            If (context.Node.IsGenerated()) Then Return
            Dim ctor As SubNewStatementSyntax = DirectCast(context.Node, SubNewStatementSyntax)
            If Not ctor.Modifiers.Any(SyntaxKind.SharedKeyword) Then Exit Sub

            Dim constructorBlock As Syntax.ConstructorBlockSyntax = DirectCast(ctor.Parent, Syntax.ConstructorBlockSyntax)
            If Not constructorBlock.Statements.Any() Then Exit Sub

            Dim throwBlock As ThrowStatementSyntax = constructorBlock.ChildNodes.OfType(Of Syntax.ThrowStatementSyntax).FirstOrDefault()
            If throwBlock Is Nothing Then Exit Sub

            context.ReportDiagnostic(Diagnostic.Create(Rule, throwBlock.GetLocation, Title))
        End Sub

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf Analyzer, SyntaxKind.SubNewStatement)
        End Sub

    End Class

End Namespace

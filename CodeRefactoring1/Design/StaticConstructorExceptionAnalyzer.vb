﻿Imports System.Collections.Immutable
Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Design
    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class StaticConstructorExceptionAnalyzer
        Inherits DiagnosticAnalyzer

        Public Shared ReadOnly Id As String = DiagnosticIds.StaticConstructorExceptionDiagnosticId
        Public Const Title As String = "Don't throw exception inside static constructors."
        Public Const MessageFormat As String = "Don't throw exceptions inside static constructors."
        Public Const Category As String = SupportedCategories.Design
        Public Const Description As String = "Static constructor are called before the first time a class is used but the caller doesn't control when exactly.
Exception thrown in this context forces callers to use 'try' block around any useage of the class and should be avoided."
        Protected Shared Rule As DiagnosticDescriptor = New DiagnosticDescriptor(
                DiagnosticIds.StaticConstructorExceptionDiagnosticId,
                Title,
                MessageFormat,
                Category,
                DiagnosticSeverity.Warning,
                isEnabledByDefault:=True,
                description:=Description,
                helpLinkUri:=HelpLink.ForDiagnostic(DiagnosticIds.StaticConstructorExceptionDiagnosticId))

        Public Overrides ReadOnly Property SupportedDiagnostics() As ImmutableArray(Of DiagnosticDescriptor) = ImmutableArray.Create(Rule)

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf Me.Analyzer, SyntaxKind.SubNewStatement)
        End Sub

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
    End Class
End Namespace

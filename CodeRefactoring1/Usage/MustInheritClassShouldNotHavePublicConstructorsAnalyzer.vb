' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Usage

    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class MustInheritClassShouldNotHavePublicConstructorsAnalyzer
        Inherits DiagnosticAnalyzer

        Private Const MessageFormat As String = "Constructor should not be public."
        Private Const Title As String = "MustInherit class should not have public constructors."

        Friend Shared Rule As New DiagnosticDescriptor(
                        AbstractClassShouldNotHavePublicCtorsDiagnosticId,
                        Title,
                        MessageFormat,
                        SupportedCategories.Usage,
                        DiagnosticSeverity.Warning,
                        isEnabledByDefault:=True,
                        description:=MessageFormat,
                        helpLinkUri:=ForDiagnostic(AbstractClassShouldNotHavePublicCtorsDiagnosticId),
                        Array.Empty(Of String))

        Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(Rule)
            End Get
        End Property

        Private Sub AnalyzeNode(context As SyntaxNodeAnalysisContext)
            If (context.Node.IsGenerated()) Then Return
            Dim constructor As SubNewStatementSyntax = DirectCast(context.Node, SubNewStatementSyntax)
            If Not constructor.Modifiers.Any(Function(m) m.IsKind(SyntaxKind.PublicKeyword)) Then Exit Sub

            Dim classDeclaration As ClassBlockSyntax = constructor.FirstAncestorOfType(Of ClassBlockSyntax)
            If classDeclaration Is Nothing Then Exit Sub
            If Not classDeclaration.ClassStatement.Modifiers.Any(Function(m) m.IsKind(SyntaxKind.MustInheritKeyword)) Then Exit Sub

            Dim diag As Diagnostic = Diagnostic.Create(Rule, constructor.GetLocation())
            context.ReportDiagnostic(diag)
        End Sub

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf AnalyzeNode, SyntaxKind.SubNewStatement)
        End Sub

    End Class

End Namespace

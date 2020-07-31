' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Usage

    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class NeverUsedAnalyzer
        Inherits DiagnosticAnalyzer
        Private Const Description As String = "When a Public method is declared but not used, remove it to avoid confusion."
        Private Const Message As String = "Pubic Method is not used."

        Private Const Title As String = "Unused Public Method"

        Private Shared ReadOnly Rule As New DiagnosticDescriptor(
                        RemovePublicMethodNeverUsedDiagnosticId,
                        Title,
                        Message,
                        SupportedCategories.Usage,
                        DiagnosticSeverity.Info,
                        isEnabledByDefault:=True,
                        Description,
                        helpLinkUri:=ForDiagnostic(RemovePublicMethodNeverUsedDiagnosticId),
                        Array.Empty(Of String))

        Public Overrides ReadOnly Property SupportedDiagnostics As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(Rule)
            End Get
        End Property

        Private Shared Function IsMethodUsed(methodTarget As MethodStatementSyntax, SyntaxTrees As IEnumerable(Of SyntaxTree)) As Boolean
            For Each tree As SyntaxTree In SyntaxTrees
                Dim root As SyntaxNode = Nothing
                If Not tree.TryGetRoot(root) Then
                    Continue For
                End If

                Dim hasIdentifier As IEnumerable(Of IdentifierNameSyntax) = root?.DescendantNodes()?.OfType(Of IdentifierNameSyntax)()
                If hasIdentifier Is Nothing OrElse Not hasIdentifier.Any() Then
                    Continue For
                End If
                For Each IdentifierName As IdentifierNameSyntax In hasIdentifier
                    If IdentifierName.Identifier.ValueText = methodTarget.Identifier.ValueText Then
                        If IdentifierName.Identifier.SpanStart = methodTarget.Identifier.SpanStart Then
                            Stop
                            Continue For
                        End If
                        Return True
                    End If
                Next
                'Return hasIdentifier.Any(Function(a) a IsNot Nothing AndAlso a.Identifier.ValueText.Equals(methodTarget?.Identifier.ValueText))
            Next
            Return False
        End Function

        Private Sub AnalyzeNode(context As SyntaxNodeAnalysisContext)
            Dim methodStatement As MethodStatementSyntax = DirectCast(context.Node, MethodStatementSyntax)
            If methodStatement.HandlesClause IsNot Nothing Then Exit Sub
            If Not methodStatement.Modifiers.Any(Function(a As SyntaxToken) a.ValueText = SyntaxFactory.Token(SyntaxKind.PublicKeyword).ValueText) Then Exit Sub
            If methodStatement.Modifiers.Any(Function(a As SyntaxToken) a.ValueText = SyntaxFactory.Token(SyntaxKind.OverridesKeyword).ValueText) Then Exit Sub
            If methodStatement.HasLeadingTrivia Then
                If methodStatement.GetLeadingTrivia.Any(Function(a As SyntaxTrivia) a.Kind = SyntaxKind.IfDirectiveTrivia) Then
                    Exit Sub
                End If
            End If
            For Each Attrib As AttributeListSyntax In methodStatement.GetAttributes
                If Attrib.ToString.Contains({"TestMethod", "Fact"}) Then Exit Sub
            Next

            If IsMethodUsed(methodStatement, context.Compilation.SyntaxTrees) Then Exit Sub
            Dim props As ImmutableDictionary(Of String, String) = New Dictionary(Of String, String) From {{"identifier", methodStatement.Identifier.Text}}.ToImmutableDictionary()
            Dim diag As Diagnostic = Diagnostic.Create(Rule, methodStatement.GetLocation(), props)
            context.ReportDiagnostic(diag)
        End Sub

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf AnalyzeNode, SyntaxKind.SubStatement, SyntaxKind.FunctionStatement)
        End Sub

    End Class

End Namespace

' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable

Imports CodeRefactoring1.Usage.MethodAnalyzers

Imports Microsoft.CodeAnalysis.Diagnostics

Namespace Usage

    <DiagnosticAnalyzer(LanguageNames.VisualBasic)>
    Public Class UriAnalyzer
        Inherits DiagnosticAnalyzer
        Private Const Category As String = SupportedCategories.Usage
        Private Const Description As String = "This diagnostic checks the Uri string and triggers if the parsing fail " + "by throwing an exception."
        Private Const MessageFormat As String = "{0}"
        Private Const Title As String = "Your Uri syntax is wrong."

        Protected Shared Rule As New DiagnosticDescriptor(
                                UriDiagnosticId,
                                Title,
                                MessageFormat,
                                Category,
                                DiagnosticSeverity.[Error],
                                isEnabledByDefault:=True,
                                Description,
                                helpLinkUri:=ForDiagnostic(UriDiagnosticId),
                                Array.Empty(Of String)
                                )

        Public Overrides ReadOnly Property SupportedDiagnostics() As ImmutableArray(Of DiagnosticDescriptor)
            Get
                Return ImmutableArray.Create(Rule)
            End Get
        End Property

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

        Public Overrides Sub Initialize(context As AnalysisContext)
            context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.Analyze Or GeneratedCodeAnalysisFlags.ReportDiagnostics)
            context.EnableConcurrentExecution()
            context.RegisterSyntaxNodeAction(AddressOf Analyzer, SyntaxKind.ObjectCreationExpression)
        End Sub

    End Class

End Namespace

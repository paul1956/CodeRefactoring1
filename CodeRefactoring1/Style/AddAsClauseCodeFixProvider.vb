Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Collections.Immutable

Imports Microsoft.CodeAnalysis.CodeFixes

Namespace Style

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(AddAsClauseCodeFixProvider)), Composition.Shared>
    Public Class AddAsClauseCodeFixProvider
        Inherits CodeFixProvider
        Private Const Title As String = "Add As Clause"

        'Public Shared ReadOnly Property Instance() As FixAllProvider = New AddAsClauseCodeFixAllProvider()
        Public NotOverridable Overrides ReadOnly Property FixableDiagnosticIds As ImmutableArray(Of String)
            Get
                Return ImmutableArray.Create(
                                            AddAsClauseDiagnosticId,
                                            AddAsClauseForLambdaDiagnosticId,
                                            ChangeAsObjectToMoreSpecificDiagnosticId,
                                            ERR_ClassNotExpression1DiagnosticId,
                                            ERR_EnumNotExpression1DiagnosticId,
                                            ERR_InterfaceNotExpression1DiagnosticId,
                                            ERR_NameNotDeclared1DiagnosticId,
                                            ERR_NamespaceNotExpression1DiagnosticId,
                                            ERR_StrictDisallowImplicitObjectDiagnosticId,
                                            ERR_StrictDisallowsImplicitProcDiagnosticId,
                                            ERR_StructureNotExpression1DiagnosticId,
                                            ERR_TypeNotExpression1DiagnosticId
                                            )
            End Get
        End Property

        Public NotOverridable Overrides Function GetFixAllProvider() As FixAllProvider
            Return WellKnownFixAllProviders.BatchFixer
        End Function

        Public NotOverridable Overrides Async Function RegisterCodeFixesAsync(Context As CodeFixContext) As Task
            Try
                Dim Root As SyntaxNode = Await Context.Document.GetSyntaxRootAsync(Context.CancellationToken).ConfigureAwait(False)
                Dim Model As SemanticModel = Await Context.Document.GetSemanticModelAsync(Context.CancellationToken)
                Dim FirstDiagnostic As Diagnostic = Context.Diagnostics.First()
                Dim DiagnosticSpanStart As Integer = FirstDiagnostic.Location.SourceSpan.Start
                Dim VariableDeclaration As SyntaxNode = Root.FindToken(DiagnosticSpanStart).Parent.FirstAncestorOrSelfOfType(GetType(VariableDeclaratorSyntax))
                If VariableDeclaration IsNot Nothing Then
                    For Each DiagnosticEntry As Diagnostic In Model.GetDiagnostics(VariableDeclaration.Span)
                        ' Does Diagnostic contain any of the errors I care about?
                        If Me.FixableDiagnosticIds.Any(Function(m As String) m.ToString = DiagnosticEntry.Id.ToString) Then
                            Exit For
                        End If
                        Exit Function
                    Next
                    Dim PossibleVariableDeclaratorSyntax As VariableDeclaratorSyntax = CType(VariableDeclaration, VariableDeclaratorSyntax)
                    If PossibleVariableDeclaratorSyntax IsNot Nothing Then
                        If PossibleVariableDeclaratorSyntax.AsClause Is Nothing OrElse PossibleVariableDeclaratorSyntax.AsClause.GetText.ToString.Trim = "As Object" Then
                            Context.RegisterCodeFix(CodeAction.Create(Title,
                                                            createChangedDocument:=Function(c As CancellationToken) AddAsClauseAsync(Context.Document, DirectCast(VariableDeclaration, VariableDeclaratorSyntax), c),
                                                            equivalenceKey:=Title),
                                                            FirstDiagnostic)
                            Exit Function
                        End If
                    End If
                End If
                Dim LambdaHeader As SyntaxNode = Root.FindToken(DiagnosticSpanStart).Parent.FirstAncestorOrSelfOfType(GetType(LambdaHeaderSyntax))
                If LambdaHeader IsNot Nothing Then
                    Context.RegisterCodeFix(CodeAction.Create(title:=Title,
                                                                createChangedDocument:=Function(c As CancellationToken) AddAsClauseAsync(Context.Document, DirectCast(LambdaHeader, LambdaHeaderSyntax), c),
                                                                equivalenceKey:=Title),
                                                                FirstDiagnostic)
                    Exit Function
                End If
                Dim ForStatement As SyntaxNode = Root.FindToken(DiagnosticSpanStart).Parent.FirstAncestorOrSelfOfType(GetType(ForStatementSyntax))
                If ForStatement IsNot Nothing Then
                    Context.RegisterCodeFix(CodeAction.Create(title:=Title,
                                                            createChangedDocument:=Function(c As CancellationToken) AddAsClauseAsync(Context.Document, DirectCast(ForStatement, ForStatementSyntax), c),
                                                            equivalenceKey:=Title),
                                                            FirstDiagnostic)
                    Exit Function
                End If
                Dim ForEachStatement As SyntaxNode = Root.FindToken(DiagnosticSpanStart).Parent.FirstAncestorOrSelfOfType(GetType(ForEachStatementSyntax))
                If ForEachStatement IsNot Nothing Then
                    Context.RegisterCodeFix(CodeAction.Create(title:=Title,
                                                                createChangedDocument:=Function(c As CancellationToken) AddAsClauseAsync(Context.Document, DirectCast(ForEachStatement, ForEachStatementSyntax), c),
                                                                equivalenceKey:=Title),
                                                                FirstDiagnostic)
                    Exit Function
                End If
            Catch ex As Exception When ex.HResult <> (New OperationCanceledException).HResult
                Stop
                Throw
            End Try
        End Function

    End Class

End Namespace
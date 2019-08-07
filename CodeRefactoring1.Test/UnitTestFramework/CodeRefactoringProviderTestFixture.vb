Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

Imports System.Threading
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeActions
Imports Microsoft.CodeAnalysis.CodeRefactorings
Imports Microsoft.CodeAnalysis.Text
Imports Xunit

Namespace Roslyn.UnitTestFramework
    Public MustInherit Class CodeRefactoringProviderTestFixture
        Inherits CodeActionProviderTestFixture

        Private Function GetRefactoring(ByVal document As Document, ByVal span As TextSpan) As IEnumerable(Of CodeAction)
            Dim provider As CodeRefactoringProvider = Me.CreateCodeRefactoringProvider()
            Dim actions As New List(Of CodeAction)()
            Dim context As New CodeRefactoringContext(document, span, Sub(a) actions.Add(a), CancellationToken.None)
            provider.ComputeRefactoringsAsync(context).Wait()
            Return actions
        End Function

        Protected Sub TestNoActions(ByVal markup As String)
            If Not markup.Contains(ControlChars.Cr) Then
                markup = markup.Replace(vbLf, vbCrLf)
            End If

            Dim code As String = Nothing
            Dim span As TextSpan = Nothing
            MarkupTestFile.GetSpan(markup, code, span)

            Dim document As Document = Me.CreateDocument(code)
            Dim actions As IEnumerable(Of CodeAction) = Me.GetRefactoring(document, span)

            Assert.True(actions Is Nothing OrElse actions.Count() = 0)
        End Sub

        Protected Sub Test(ByVal markup As String, ByVal expected As String, Optional ByVal actionIndex As Integer = 0, Optional ByVal compareTokens As Boolean = False)
            If Not markup.Contains(ControlChars.Cr) Then
                markup = markup.Replace(vbLf, vbCrLf)
            End If

            If Not expected.Contains(ControlChars.Cr) Then
                expected = expected.Replace(vbLf, vbCrLf)
            End If

            Dim code As String = Nothing
            Dim span As TextSpan = Nothing
            MarkupTestFile.GetSpan(markup, code, span)

            Dim document As Document = Me.CreateDocument(code)
            Dim actions As IEnumerable(Of CodeAction) = Me.GetRefactoring(document, span)

            Assert.NotNull(actions)

            Dim action As CodeAction = actions.ElementAt(actionIndex)
            Assert.NotNull(action)

            Dim edit As ApplyChangesOperation = action.GetOperationsAsync(CancellationToken.None).Result.OfType(Of ApplyChangesOperation)().First()
            Me.VerifyDocument(expected, compareTokens, edit.ChangedSolution.GetDocument(document.Id))
        End Sub

        Protected MustOverride Function CreateCodeRefactoringProvider() As CodeRefactoringProvider
    End Class
End Namespace

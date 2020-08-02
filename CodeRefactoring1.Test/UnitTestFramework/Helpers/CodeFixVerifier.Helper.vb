' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Collections.Immutable
Imports System.Threading
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeActions
Imports Microsoft.CodeAnalysis.Formatting
Imports Microsoft.CodeAnalysis.Simplification

Namespace TestHelper

    ' Diagnostic Producer class with extra methods dealing with applying codefixes
    ' All methods are shared
    Partial Public MustInherit Class CodeFixVerifier
        Inherits DiagnosticVerifier

        ''' <summary>
        ''' Apply the inputted CodeAction to the inputted document.
        ''' Meant to be used to apply codefixes.
        ''' </summary>
        ''' <param name="document">The Document to apply the fix on</param>
        ''' <param name="codeAction">A CodeAction that will be applied to the Document.</param>
        ''' <returns>A Document with the changes from the CodeAction</returns>
        Private Shared Async Function ApplyFix(document As Document, codeAction As CodeAction) As Task(Of Document)
            Dim operations As ImmutableArray(Of CodeActionOperation) = Await codeAction.GetOperationsAsync(CancellationToken.None)
            Dim solution As Solution = operations.OfType(Of ApplyChangesOperation).Single.ChangedSolution
            Dim Project1 As Project = solution.GetProject(solution.ProjectIds(0))
            Return solution.GetDocument(document.Id)
        End Function

        ''' <summary>
        ''' Compare two collections of Diagnostics, and return a list of any New diagnostics that appear only in the second collection.
        ''' Note: Considers Diagnostics to be the same if they have the same Ids.  In the case of multiple diagnostics With the same Id In a row,
        ''' this method may not necessarily return the new one.
        ''' </summary>
        ''' <param name="diagnostics">The Diagnostics that existed in the code before the CodeFix was applied</param>
        ''' <param name="newDiagnostics">The Diagnostics that exist in the code after the CodeFix was applied</param>
        ''' <returns>A list of Diagnostics that only surfaced in the code after the CodeFix was applied</returns>
        Private Shared Iterator Function GetNewDiagnostics(diagnostics As IEnumerable(Of Diagnostic), newDiagnostics As IEnumerable(Of Diagnostic)) As IEnumerable(Of Diagnostic)

            Dim oldArray As Diagnostic() = diagnostics.OrderBy(Function(d) d.Location.SourceSpan.Start).ToArray()
            Dim newArray As Diagnostic() = newDiagnostics.OrderBy(Function(d) d.Location.SourceSpan.Start).ToArray()

            Dim oldIndex As Integer = 0
            Dim newIndex As Integer = 0

            While newIndex < newArray.Length

                If oldIndex < oldArray.Length AndAlso oldArray(oldIndex).Id = newArray(newIndex).Id Then
                    oldIndex += 1
                    newIndex += 1
                Else
                    Yield newArray(newIndex)
                    newIndex += 1
                End If
            End While

        End Function

        ''' <summary>
        ''' Get the existing compiler diagnostics on the inputted document.
        ''' </summary>
        ''' <param name="document">The Document to run the compiler diagnostic analyzers on</param>
        ''' <returns>The compiler diagnostics that were found in the code</returns>
        Private Shared Function GetCompilerDiagnostics(document As Document) As IEnumerable(Of Diagnostic)
            Return document.GetSemanticModelAsync().Result.GetDiagnostics()
        End Function

        ''' <summary>
        ''' Given a Document, turn it into a string based on the syntax root
        ''' </summary>
        ''' <param name="document">The Document to be converted to a string</param>
        ''' <returns>A string containing the syntax of the Document after formatting</returns>
        Private Shared Function GetStringFromDocument(document As Document) As String
            Dim simplifiedDoc As Document = Simplifier.ReduceAsync(document, Simplifier.Annotation).Result
            Dim root As SyntaxNode = simplifiedDoc.GetSyntaxRootAsync().Result
            root = Formatter.Format(root, Formatter.Annotation, simplifiedDoc.Project.Solution.Workspace)
            Return root.GetText().ToString()
        End Function

    End Class

End Namespace

' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.CodeActions
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports VBRefactorings.Usage.UnusedParametersCodeFixProvider

Namespace Usage

    Public NotInheritable Class UnusedParametersCodeFixAllProvider
        Inherits FixAllProvider

        Private Const Message As String = "Remove unused parameter"

        Public Shared Instance As UnusedParametersCodeFixAllProvider = New UnusedParametersCodeFixAllProvider

        Private Sub New()
            MyBase.New

        End Sub

        Private Shared Async Function GetDiagnosticsInDocAsync(fixAllContext As FixAllContext, document As Document) As Task(Of DiagnosticsInDoc)
            Dim diagnostics As Immutable.ImmutableArray(Of Diagnostic) = Await fixAllContext.GetDocumentDiagnosticsAsync(document).ConfigureAwait(False)
            If Not diagnostics.Any Then
                Return DiagnosticsInDoc.Empty
            End If
            Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(fixAllContext.CancellationToken).ConfigureAwait(False)
            Dim doc As DiagnosticsInDoc = DiagnosticsInDoc.Create(document.Id, diagnostics, root)
            Return doc
        End Function

        Private Shared Async Function GetFixedSolutionAsync(fixAllContext As FixAllContext, sol As SolutionWithDocs) As Task(Of Solution)
            Dim newSolution As Solution = sol.Solution
            For Each doc As DiagnosticsInDoc In sol.Docs
                For Each node As SyntaxNode In doc.Nodes
                    Dim document As Document = newSolution.GetDocument(doc.DocumentId)
                    Dim root As SyntaxNode = Await document.GetSyntaxRootAsync(fixAllContext.CancellationToken).ConfigureAwait(False)
                    Dim trackedNode As SyntaxNode = root.GetCurrentNode(node)
                    Dim parameter As ParameterSyntax = trackedNode.AncestorsAndSelf().OfType(Of ParameterSyntax).First()
                    Dim docResults As List(Of DocumentIdAndRoot) = Await RemoveParameterAsync(document, parameter, root, fixAllContext.CancellationToken)
                    For Each docResult As DocumentIdAndRoot In docResults
                        newSolution = newSolution.WithDocumentSyntaxRoot(docResult._documentId, docResult._root)
                    Next
                Next
            Next
            Return newSolution
        End Function

        Private Overloads Shared Async Function GetSolutionWithDocsAsync(fixAllContext As FixAllContext, solution As Solution) As Task(Of SolutionWithDocs)
            Dim docs As List(Of DiagnosticsInDoc) = New List(Of DiagnosticsInDoc)
            Dim sol As SolutionWithDocs = New SolutionWithDocs() With {.Docs = docs, .Solution = solution}
            For Each pId As ProjectId In solution.Projects.Select(Function(p As Project) p.Id)
                Dim project As Project = sol.Solution.GetProject(pId)
                Dim newSol As SolutionWithDocs = Await GetSolutionWithDocsAsync(fixAllContext, project).ConfigureAwait(False)
                sol.Merge(newSol)
            Next
            Return sol
        End Function

        Private Overloads Shared Async Function GetSolutionWithDocsAsync(fixAllContext As FixAllContext, project As Project) As Task(Of SolutionWithDocs)
            Dim docs As List(Of DiagnosticsInDoc) = New List(Of DiagnosticsInDoc)
            Dim newSolution As Solution = project.Solution
            For Each document As Document In project.Documents
                Dim doc As DiagnosticsInDoc = Await GetDiagnosticsInDocAsync(fixAllContext, document)
                If doc.Equals(DiagnosticsInDoc.Empty) Then Continue For
                docs.Add(doc)
                newSolution = newSolution.WithDocumentSyntaxRoot(document.Id, doc.TrackedRoot)
            Next
            Dim sol As SolutionWithDocs = New SolutionWithDocs() With {.Docs = docs, .Solution = newSolution}
            Return sol
        End Function

        Private Overloads Shared Async Function GetSolutionWithDocsAsync(fixAllContext As FixAllContext, document As Document) As Task(Of SolutionWithDocs)
            Dim docs As List(Of DiagnosticsInDoc) = New List(Of DiagnosticsInDoc)
            Dim doc As DiagnosticsInDoc = Await GetDiagnosticsInDocAsync(fixAllContext, document)
            docs.Add(doc)
            Dim newSolution As Solution = document.Project.Solution.WithDocumentSyntaxRoot(document.Id, doc.TrackedRoot)
            Dim sol As SolutionWithDocs = New SolutionWithDocs() With {.Docs = docs, .Solution = newSolution}
            Return sol
        End Function

        Public Overrides Function GetFixAsync(fixAllContext As FixAllContext) As Task(Of CodeAction)
            Select Case fixAllContext.Scope
                Case FixAllScope.Document
                    Return Task.FromResult(CodeAction.Create(Message, Async Function(ct As CancellationToken) Await GetFixedSolutionAsync(fixAllContext, (Await GetSolutionWithDocsAsync(fixAllContext, fixAllContext.Document)))))
                Case FixAllScope.Project
                    Return Task.FromResult(CodeAction.Create(Message, Async Function(ct As CancellationToken) Await GetFixedSolutionAsync(fixAllContext, (Await GetSolutionWithDocsAsync(fixAllContext, fixAllContext.Project)))))
                Case FixAllScope.Solution
                    Return Task.FromResult(CodeAction.Create(Message, Async Function(ct As CancellationToken) Await GetFixedSolutionAsync(fixAllContext, (Await GetSolutionWithDocsAsync(fixAllContext, fixAllContext.Solution)))))
            End Select

            Return Nothing
        End Function

        Private Structure DiagnosticsInDoc
            Public DocumentId As DocumentId

            Public Nodes As List(Of SyntaxNode)

            Public TrackedRoot As SyntaxNode

            Public Shared Property Empty As DiagnosticsInDoc = New DiagnosticsInDoc()

            Public Shared Function Create(documentId As DocumentId, diagnostics As IList(Of Diagnostic), root As SyntaxNode) As DiagnosticsInDoc
                Dim nodes As List(Of SyntaxNode) = diagnostics.Select(Function(d As Diagnostic) root.FindNode(d.Location.SourceSpan)).Where(Function(n As SyntaxNode) Not n.IsMissing).ToList()
                Dim diagnosticsInDoc As DiagnosticsInDoc = New DiagnosticsInDoc() With {.DocumentId = documentId, .TrackedRoot = root.TrackNodes(nodes), .Nodes = nodes}
                Return diagnosticsInDoc
            End Function

        End Structure

        Private Structure SolutionWithDocs

            Public Docs As List(Of DiagnosticsInDoc)
            Public Solution As Solution

            Public Sub Merge(sol As SolutionWithDocs)
                Solution = sol.Solution
                Docs.AddRange(sol.Docs)
            End Sub

        End Structure

    End Class

End Namespace

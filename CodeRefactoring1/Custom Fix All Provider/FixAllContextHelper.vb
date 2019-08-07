Option Infer On
Option Explicit Off
Option Strict Off

Imports System.Collections.Concurrent
Imports System.Collections.Immutable
Imports System.Reflection
Imports Microsoft.CodeAnalysis.CodeFixes
Imports Microsoft.CodeAnalysis.Diagnostics

Friend Class FixAllContextHelper
    Private Shared ReadOnly DiagnosticAnalyzers As ImmutableDictionary(Of String, ImmutableArray(Of Type)) = GetAllAnalyzers()
    Private Shared ReadOnly GetAnalyzerCompilationDiagnosticsAsync As Func(Of CompilationWithAnalyzers, ImmutableArray(Of DiagnosticAnalyzer), CancellationToken, Task(Of ImmutableArray(Of Diagnostic)))
    Private Shared ReadOnly GetAnalyzerSemanticDiagnosticsAsync As Func(Of CompilationWithAnalyzers, SemanticModel, TextSpan?, CancellationToken, Task(Of ImmutableArray(Of Diagnostic)))
    Private Shared ReadOnly GetAnalyzerSyntaxDiagnosticsAsync As Func(Of CompilationWithAnalyzers, SyntaxTree, CancellationToken, Task(Of ImmutableArray(Of Diagnostic)))

    Public Shared Async Function GetAllDiagnosticsAsync(ByVal compilation As Compilation, ByVal compilationWithAnalyzers As CompilationWithAnalyzers, ByVal analyzers As ImmutableArray(Of DiagnosticAnalyzer), ByVal documents As IEnumerable(Of Document), ByVal includeCompilerDiagnostics As Boolean, ByVal cancellationToken As CancellationToken) As Task(Of ImmutableArray(Of Diagnostic))
        If GetAnalyzerSyntaxDiagnosticsAsync Is Nothing OrElse GetAnalyzerSemanticDiagnosticsAsync Is Nothing Then
            Dim allDiagnostics As ImmutableArray(Of Diagnostic) = Await compilationWithAnalyzers.GetAllDiagnosticsAsync().ConfigureAwait(False)
            Return allDiagnostics
        End If
        compilationWithAnalyzers.Compilation.GetDeclarationDiagnostics(cancellationToken)
        ' Note that the following loop to obtain syntax and semantic diagnostics for each document cannot operate
        ' on parallel due to our use of a single CompilationWithAnalyzers instance.
        Dim diagnostics As ImmutableArray(Of Diagnostic).Builder = ImmutableArray.CreateBuilder(Of Diagnostic)()
        For Each Document As Document In documents
            Dim syntaxTree As SyntaxTree = Await Document.GetSyntaxTreeAsync(cancellationToken).ConfigureAwait(False)
            Dim syntaxDiagnostics As ImmutableArray(Of Diagnostic) = Await GetAnalyzerSyntaxDiagnosticsAsync(compilationWithAnalyzers, syntaxTree, cancellationToken).ConfigureAwait(False)
            diagnostics.AddRange(syntaxDiagnostics)
            Dim semanticModel As SemanticModel = Await Document.GetSemanticModelAsync(cancellationToken).ConfigureAwait(False)
            Dim semanticDiagnostics As ImmutableArray(Of Diagnostic) = Await GetAnalyzerSemanticDiagnosticsAsync(compilationWithAnalyzers, semanticModel, Nothing, cancellationToken).ConfigureAwait(False)
            diagnostics.AddRange(semanticDiagnostics)
        Next Document
        For Each analyzer As DiagnosticAnalyzer In analyzers
            diagnostics.AddRange(Await GetAnalyzerCompilationDiagnosticsAsync(compilationWithAnalyzers, ImmutableArray.Create(analyzer), cancellationToken).ConfigureAwait(False))
        Next analyzer
        If includeCompilerDiagnostics Then
            ' This is the special handling for cases where code fixes operate on warnings produced by the C#
            ' compiler, as opposed to being created by specific analyzers.
            Dim compilerDiagnostics As ImmutableArray(Of Diagnostic) = compilation.GetDiagnostics(cancellationToken)
            diagnostics.AddRange(compilerDiagnostics)
        End If
        Return diagnostics.ToImmutable()
    End Function

    ' SyncLock workaround implemented
    Public Shared Async Function GetDocumentDiagnosticsToFixAsync(ByVal fixAllContext As CodeFixes.FixAllContext) As Task(Of ImmutableDictionary(Of Document, ImmutableArray(Of Diagnostic)))
        Dim allDiagnostics As ImmutableArray(Of Diagnostic) = ImmutableArray(Of Diagnostic).Empty
        Dim projectsToFix As ImmutableArray(Of Project) = ImmutableArray(Of Project).Empty
        Dim document As Document = fixAllContext.Document
        Dim project As Project = fixAllContext.Project
        Select Case fixAllContext.Scope
            Case FixAllScope.Document
                If document IsNot Nothing Then
                    Dim documentDiagnostics As ImmutableArray(Of Diagnostic) = Await fixAllContext.GetDocumentDiagnosticsAsync(document).ConfigureAwait(False)
                    Return ImmutableDictionary(Of Document, ImmutableArray(Of Diagnostic)).Empty.SetItem(document, documentDiagnostics)
                End If
            Case FixAllScope.Project
                projectsToFix = ImmutableArray.Create(project)
                allDiagnostics = Await GetAllDiagnosticsAsync(fixAllContext, project).ConfigureAwait(False)
            Case FixAllScope.Solution
                projectsToFix = project.Solution.Projects.Where(Function(p As Project) p.Language = project.Language).ToImmutableArray()
                Dim diagnostics As ConcurrentDictionary(Of ProjectId, ImmutableArray(Of Diagnostic)) = New ConcurrentDictionary(Of ProjectId, ImmutableArray(Of Diagnostic))()
                Dim tasks As Task() = New Task(projectsToFix.Length - 1) {}
                For i As Integer = 0 To projectsToFix.Length - 1
                    fixAllContext.CancellationToken.ThrowIfCancellationRequested()
                    Dim projectToFix As Project = projectsToFix(i)
                    tasks(i) = Task.Run(Async Function()
                                            Dim projectDiagnostics As ImmutableArray(Of Diagnostic) = Await GetAllDiagnosticsAsync(fixAllContext, projectToFix).ConfigureAwait(False)
                                            diagnostics.TryAdd(projectToFix.Id, projectDiagnostics)
                                        End Function, fixAllContext.CancellationToken)
                Next i
                Await Task.WhenAll(tasks).ConfigureAwait(False)
                allDiagnostics = allDiagnostics.AddRange(diagnostics.SelectMany(Function(i As KeyValuePair(Of ProjectId, ImmutableArray(Of Diagnostic))) i.Value.Where(Function(x As Diagnostic) fixAllContext.DiagnosticIds.Contains(x.Id))))
        End Select
        If allDiagnostics.IsEmpty Then
            Return ImmutableDictionary(Of Document, ImmutableArray(Of Diagnostic)).Empty
        End If
        Return Await GetDocumentDiagnosticsToFixAsync(allDiagnostics, projectsToFix, fixAllContext.CancellationToken).ConfigureAwait(False)
    End Function

    'Public Shared Async Function GetProjectDiagnosticsToFixAsync(ByVal fixAllContext As CodeFixes.FixAllContext) As Task(Of ImmutableDictionary(Of Project, ImmutableArray(Of Diagnostic)))
    '    Dim project As Project = fixAllContext.Project
    '    If project IsNot Nothing Then
    '        Select Case fixAllContext.Scope
    '            Case FixAllScope.Project
    '                Dim diagnostics As ImmutableArray(Of Diagnostic) = Await fixAllContext.GetProjectDiagnosticsAsync(project).ConfigureAwait(False)
    '                Return ImmutableDictionary(Of Project, ImmutableArray(Of Diagnostic)).Empty.SetItem(project, diagnostics)
    '            Case FixAllScope.Solution
    '                Dim projectsAndDiagnostics As ConcurrentDictionary(Of Project, ImmutableArray(Of Diagnostic)) = New ConcurrentDictionary(Of Project, ImmutableArray(Of Diagnostic))()
    '                Dim options As ParallelOptions = New ParallelOptions() With {.CancellationToken = fixAllContext.CancellationToken}
    '                Parallel.ForEach(project.Solution.Projects, options, Sub(proj As Project)
    '                                                                         fixAllContext.CancellationToken.ThrowIfCancellationRequested()
    '                                                                         Dim projectDiagnosticsTask As Task(Of ImmutableArray(Of Diagnostic)) = fixAllContext.GetProjectDiagnosticsAsync(proj)
    '                                                                         projectDiagnosticsTask.Wait(fixAllContext.CancellationToken)
    '                                                                         Dim projectDiagnostics As ImmutableArray(Of Diagnostic) = projectDiagnosticsTask.Result
    '                                                                         If projectDiagnostics.Any() Then
    '                                                                             projectsAndDiagnostics.TryAdd(proj, projectDiagnostics)
    '                                                                         End If
    '                                                                     End Sub)
    '                Return projectsAndDiagnostics.ToImmutableDictionary()
    '        End Select
    '    End If
    '    Return ImmutableDictionary(Of Project, ImmutableArray(Of Diagnostic)).Empty
    'End Function
    Private Shared Function GetAllAnalyzers() As ImmutableDictionary(Of String, ImmutableArray(Of Type))
        Dim assembly As Assembly = GetType(NoCodeFixAttribute).GetTypeInfo().Assembly
        Dim diagnosticAnalyzerType As Type = GetType(DiagnosticAnalyzer)
        Dim analyzers As ImmutableDictionary(Of String, ImmutableArray(Of Type)).Builder = ImmutableDictionary.CreateBuilder(Of String, ImmutableArray(Of Type))()
        For Each Type As Reflection.TypeInfo In assembly.DefinedTypes
            If Type.IsSubclassOf(diagnosticAnalyzerType) AndAlso Not Type.IsAbstract Then
                Dim analyzerType As Type = Type.AsType()
                Dim analyzer As DiagnosticAnalyzer = DirectCast(Activator.CreateInstance(analyzerType), DiagnosticAnalyzer)
                For Each descriptor As DiagnosticDescriptor In analyzer.SupportedDiagnostics
                    Dim types As ImmutableArray(Of Type) = Nothing
                    types = If(analyzers.TryGetValue(descriptor.Id, types), types.Add(analyzerType), ImmutableArray.Create(analyzerType))
                    analyzers(descriptor.Id) = types
                Next descriptor
            End If
        Next Type
        Return analyzers.ToImmutable()
    End Function

    ''' <summary>
    ''' Gets all <see cref="Diagnostic"/> instances within a specific <see cref="Project"/> which are relevant to a
    ''' <see cref="CodeFixes.FixAllContext"/>.
    ''' </summary>
    ''' <param name="fixAllContext">The context for the Fix All operation.</param>
    ''' <param name="project">The project.</param>
    ''' <returns>A <see cref="Task{TResult}"/> representing the asynchronous operation. When the task completes
    ''' successfully, the <see cref="Task{TResult}.Result"/> will contain the requested diagnostics.</returns>
    Private Shared Async Function GetAllDiagnosticsAsync(ByVal fixAllContext As CodeFixes.FixAllContext, ByVal project As Project) As Task(Of ImmutableArray(Of Diagnostic))
        If GetAnalyzerSyntaxDiagnosticsAsync Is Nothing OrElse GetAnalyzerSemanticDiagnosticsAsync Is Nothing Then
            Return Await fixAllContext.GetAllDiagnosticsAsync(project).ConfigureAwait(False)
        End If
        '
        '             * The rest of this method is workaround code for issues with Roslyn 1.1...
        '
        Dim analyzers As ImmutableArray(Of DiagnosticAnalyzer) = GetDiagnosticAnalyzersForContext(fixAllContext)
        ' Most code fixes in this project operate on diagnostics reported by analyzers in this project. However, a
        ' few code fixes also operate on standard warnings produced by the C# compiler. Special handling is
        ' required for the latter case since these warnings are not considered "analyzer diagnostics".
        Dim includeCompilerDiagnostics As Boolean = fixAllContext.DiagnosticIds.Any(Function(x As String) x.StartsWith("BC", StringComparison.Ordinal))
        ' Use a single CompilationWithAnalyzers for the entire operation. This allows us to use the
        ' GetDeclarationDiagnostics workaround for dotnet/roslyn#7446 a single time, rather than once per document.
        Dim compilation As Compilation = Await project.GetCompilationAsync(fixAllContext.CancellationToken).ConfigureAwait(False)
        Dim compilationWithAnalyzers As CompilationWithAnalyzers = compilation.WithAnalyzers(analyzers, project.AnalyzerOptions, fixAllContext.CancellationToken)
        Dim diagnostics As ImmutableArray(Of Diagnostic) = Await GetAllDiagnosticsAsync(compilation, compilationWithAnalyzers, analyzers, project.Documents, includeCompilerDiagnostics, fixAllContext.CancellationToken).ConfigureAwait(False)
        ' Make sure to filter the results to the set requested for the Fix All operation, since analyzers can
        ' report diagnostics with different IDs.
        diagnostics = diagnostics.RemoveAll(Function(x As Diagnostic) Not fixAllContext.DiagnosticIds.Contains(x.Id))
        Return diagnostics
    End Function

    Private Shared Function GetDiagnosticAnalyzersForContext(ByVal fixAllContext As CodeFixes.FixAllContext) As ImmutableArray(Of DiagnosticAnalyzer)
        Return DiagnosticAnalyzers.Where(Function(x As KeyValuePair(Of String, ImmutableArray(Of Type))) fixAllContext.DiagnosticIds.Contains(x.Key)).SelectMany(Function(x As KeyValuePair(Of String, ImmutableArray(Of Type))) x.Value).Distinct().Select(Function(type As Type) DirectCast(Activator.CreateInstance(type), DiagnosticAnalyzer)).ToImmutableArray()
    End Function
    Private Shared Async Function GetDocumentDiagnosticsToFixAsync(ByVal diagnostics As ImmutableArray(Of Diagnostic), ByVal projects As ImmutableArray(Of Project), ByVal cancellationToken As CancellationToken) As Task(Of ImmutableDictionary(Of Document, ImmutableArray(Of Diagnostic)))
        Dim treeToDocumentMap As ImmutableDictionary(Of SyntaxTree, Document) = Await GetTreeToDocumentMapAsync(projects, cancellationToken).ConfigureAwait(False)
        Dim builder As ImmutableDictionary(Of Document, ImmutableArray(Of Diagnostic)).Builder = ImmutableDictionary.CreateBuilder(Of Document, ImmutableArray(Of Diagnostic))()
        For Each documentAndDiagnostics As IGrouping(Of Document, Diagnostic) In diagnostics.GroupBy(Function(d As Diagnostic) GetReportedDocument(d, treeToDocumentMap))
            cancellationToken.ThrowIfCancellationRequested()
            Dim document As Document = documentAndDiagnostics.Key
            Dim diagnosticsForDocument As ImmutableArray(Of Diagnostic) = documentAndDiagnostics.ToImmutableArray()
            builder.Add(document, diagnosticsForDocument)
        Next documentAndDiagnostics
        Return builder.ToImmutable()
    End Function

    Private Shared Function GetReportedDocument(ByVal diagnostic As Diagnostic, ByVal treeToDocumentsMap As ImmutableDictionary(Of SyntaxTree, Document)) As Document
        Dim tree As SyntaxTree = diagnostic.Location.SourceTree
        If tree IsNot Nothing Then
            Dim document As Document = Nothing
            If treeToDocumentsMap.TryGetValue(tree, document) Then
                Return document
            End If
        End If
        Return Nothing
    End Function

    Private Shared Async Function GetTreeToDocumentMapAsync(ByVal projects As ImmutableArray(Of Project), ByVal cancellationToken As CancellationToken) As Task(Of ImmutableDictionary(Of SyntaxTree, Document))
        Dim builder As ImmutableDictionary(Of SyntaxTree, Document).Builder = ImmutableDictionary.CreateBuilder(Of SyntaxTree, Document)()
        For Each Project As Project In projects
            cancellationToken.ThrowIfCancellationRequested()
            For Each Document As Document In Project.Documents
                cancellationToken.ThrowIfCancellationRequested()
                Dim tree As SyntaxTree = Await Document.GetSyntaxTreeAsync(cancellationToken).ConfigureAwait(False)
                builder.Add(tree, Document)
            Next Document
        Next Project
        Return builder.ToImmutable()
    End Function
End Class
﻿Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On
Imports System.Collections.Immutable
Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions

Public Module GeneratedCodeAnalysisExtensions
    Public Cache As New WeakReference(Of ImmutableDictionary(Of SyntaxTree, Boolean))(ImmutableDictionary(Of SyntaxTree, Boolean).Empty)

    Private Function ContainsAutogeneratedComment(tree As SyntaxTree, Optional cancellationToken As CancellationToken = Nothing) As Boolean
        Dim root As SyntaxNode = tree.GetRoot(cancellationToken)
        If root Is Nothing Then
            Return False
        End If
        Dim firstToken As SyntaxToken = root.GetFirstToken()
        If Not firstToken.HasLeadingTrivia Then
            Return False
        End If

        For Each trivia As SyntaxTrivia In firstToken.LeadingTrivia.Where(Function(t) t.IsKind(SyntaxKind.CommentTrivia)).Take(2)
            Dim str As String = trivia.ToString()
            If str = "// This file has been generated by the GUI designer. Do not modify." OrElse
                str = "// <auto-generated>" OrElse
                str = "// <autogenerated>" Then
                Return True
            End If
        Next
        Return False
    End Function

    <Extension>
    Public Function IsFromGeneratedCode(semanticModel As SemanticModel, cancellationToken As CancellationToken) As Boolean
        Dim table As ImmutableDictionary(Of SyntaxTree, Boolean) = Nothing
        Dim tree As SyntaxTree = semanticModel.SyntaxTree

        If Cache.TryGetTarget(table) Then
            If table.ContainsKey(tree) Then
                Return table(tree)
            End If
        Else
            table = ImmutableDictionary(Of SyntaxTree, Boolean).Empty
        End If

        Dim result As Boolean = tree.FilePath.IsOnGeneratedFile OrElse ContainsAutogeneratedComment(tree, cancellationToken)
        Cache.SetTarget(table.Add(tree, result))

        Return result
    End Function

    ''' <summary>
    ''' This function tries to figure out if a file is generate and should not be processed
    ''' </summary>
    ''' <param name="filePath"></param>
    ''' <returns>True if file is generated</returns>
    <Extension>
    Public Function IsOnGeneratedFile(filePath As String) As Boolean
        Return Regex.IsMatch(filePath, "(\\service|\\TemporaryGeneratedFile_.*|\\assemblyinfo|\\assemblyattributes|\.(g\.i|g|designer|generated|assemblyattributes))\.(cs|vb)$", RegexOptions.IgnoreCase)
    End Function
End Module

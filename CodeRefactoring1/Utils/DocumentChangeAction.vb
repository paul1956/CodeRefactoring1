' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Public NotInheritable Class DocumentChangeAction
    Inherits CodeAction

    Private ReadOnly m_CreateChangedDocument As Func(Of CancellationToken, Task(Of Document))
    Private ReadOnly m_Title As String
    Private m_Severity As DiagnosticSeverity
    Private m_TextSpan As TextSpan

    Public Sub New(ByVal _TextSpan As TextSpan, ByVal _Severity As DiagnosticSeverity, ByVal title As String, ByVal createChangedDocument As Func(Of CancellationToken, Task(Of Document)))
        m_Severity = _Severity
        m_TextSpan = _TextSpan
        m_Title = title
        m_CreateChangedDocument = createChangedDocument
    End Sub

    Public Property Severity() As DiagnosticSeverity
        Get
            Return m_Severity
        End Get
        Private Set(ByVal value As DiagnosticSeverity)
            m_Severity = value
        End Set
    End Property

    Public Property TextSpan() As TextSpan
        Get
            Return m_TextSpan
        End Get
        Private Set(ByVal value As TextSpan)
            m_TextSpan = value
        End Set
    End Property

    Public Overrides ReadOnly Property Title() As String
        Get
            Return m_Title
        End Get
    End Property

    Protected Overrides Function GetChangedDocumentAsync(ByVal cancellationToken As CancellationToken) As Task(Of Document)
        If m_CreateChangedDocument Is Nothing Then
            Return MyBase.GetChangedDocumentAsync(cancellationToken)
        End If

        Dim task As Task(Of Document) = m_CreateChangedDocument.Invoke(cancellationToken)
        Return task
    End Function

End Class

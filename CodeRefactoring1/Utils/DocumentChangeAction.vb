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
        Me.m_Severity = _Severity
        Me.m_TextSpan = _TextSpan
        Me.m_Title = title
        Me.m_CreateChangedDocument = createChangedDocument
    End Sub

    Public Property Severity() As DiagnosticSeverity
        Get
            Return Me.m_Severity
        End Get
        Private Set(ByVal value As DiagnosticSeverity)
            Me.m_Severity = value
        End Set
    End Property

    Public Property TextSpan() As TextSpan
        Get
            Return Me.m_TextSpan
        End Get
        Private Set(ByVal value As TextSpan)
            Me.m_TextSpan = value
        End Set
    End Property

    Public Overrides ReadOnly Property Title() As String
        Get
            Return Me.m_Title
        End Get
    End Property

    Protected Overrides Function GetChangedDocumentAsync(ByVal cancellationToken As CancellationToken) As Task(Of Document)
        If Me.m_CreateChangedDocument Is Nothing Then
            Return MyBase.GetChangedDocumentAsync(cancellationToken)
        End If

        Dim task As Task(Of Document) = Me.m_CreateChangedDocument.Invoke(cancellationToken)
        Return task
    End Function

End Class
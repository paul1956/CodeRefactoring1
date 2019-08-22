' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.
Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Namespace TestHelper
    ''' <summary>
    ''' Location where the diagnostic appears, as determined by path, line number, And column number.
    ''' </summary>
    Public Class DiagnosticResultLocation

        Public Sub New(path As String, line As Integer, column As Integer)
            If line < -1 Then
                Throw New ArgumentOutOfRangeException(NameOf(line), "line must be >= -1")
            End If

            If column < -1 Then
                Throw New ArgumentOutOfRangeException(NameOf(column), "column must be >= -1")
            End If

            Me.Path = path
            Me.Line = line
            Me.Column = column

        End Sub

        Public Property Column As Integer
        Public Property Line As Integer
        Public Property Path As String
    End Class

End Namespace
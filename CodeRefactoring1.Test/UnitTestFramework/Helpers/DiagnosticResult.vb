' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.
Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Imports Microsoft.CodeAnalysis

Namespace TestHelper

    ''' <summary>
    ''' Struct that stores information about a Diagnostic appearing in a source
    ''' </summary>
    Public Structure DiagnosticResult
        Private innerlocations As DiagnosticResultLocation()

        Public ReadOnly Property Column As Integer
            Get
                Return If(Me.Locations.Length > 0, Me.Locations(0).Column, -1)
            End Get
        End Property

        Public Property Id As String

        Public ReadOnly Property Line As Integer
            Get
                Return If(Me.Locations.Length > 0, Me.Locations(0).Line, -1)
            End Get
        End Property

        Public Property Locations As DiagnosticResultLocation()
            Get
                If Me.innerlocations Is Nothing Then
                    Me.innerlocations = {}
                End If
                Return Me.innerlocations
            End Get
            Set
                Me.innerlocations = Value
            End Set
        End Property

        Public Property Message As String

        Public ReadOnly Property Path As String
            Get
                Return If(Me.Locations.Length > 0, Me.Locations(0).Path, "")
            End Get
        End Property

        Public Property Severity As DiagnosticSeverity
    End Structure

    ''' <summary>
    ''' Location where the diagnostic appears, as determined by path, line number, And column number.
    ''' </summary>
    Public Structure DiagnosticResultLocation

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
    End Structure

End Namespace
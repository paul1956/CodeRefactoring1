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
                    Me.innerlocations = Array.Empty(Of DiagnosticResultLocation)()
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

        Public Shared Operator =(left As DiagnosticResult, right As DiagnosticResult) As Boolean
            Return left.Equals(right)
        End Operator

        Public Shared Operator <>(left As DiagnosticResult, right As DiagnosticResult) As Boolean
            Return Not left = right
        End Operator

        Public Overrides Function Equals(obj As Object) As Boolean
            Throw New NotImplementedException()
        End Function

        Public Overrides Function GetHashCode() As Integer
            Throw New NotImplementedException()
        End Function
    End Structure

End Namespace
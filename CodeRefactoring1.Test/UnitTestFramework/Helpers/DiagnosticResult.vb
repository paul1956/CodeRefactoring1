' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

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
        Private _innerlocations As DiagnosticResultLocation()

        Public ReadOnly Property Column As Integer
            Get
                Return If(Locations.Length > 0, Locations(0).Column, -1)
            End Get
        End Property

        Public Property Id As String

        Public ReadOnly Property Line As Integer
            Get
                Return If(Locations.Length > 0, Locations(0).Line, -1)
            End Get
        End Property

#Disable Warning CA1819 ' Properties should not return arrays
        Public Property Locations As DiagnosticResultLocation()
#Enable Warning CA1819 ' Properties should not return arrays
            Get
                If _innerlocations Is Nothing Then
                    _innerlocations = Array.Empty(Of DiagnosticResultLocation)()
                End If
                Return _innerlocations
            End Get
            Set
                _innerlocations = Value
            End Set
        End Property

        Public Property Message As String

        Public ReadOnly Property Path As String
            Get
                Return If(Locations.Length > 0, Locations(0).Path, "")
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

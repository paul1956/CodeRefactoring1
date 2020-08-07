' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis.Text

Public Module CompareGenerticObject

    <Extension>
    Public Function CompareT(Of T)(Left As T, Right As T, ByRef MismatchedPropertyName As String) As Boolean
        Dim properties() As PropertyInfo = Left.GetType().GetProperties(BindingFlags.Public Or BindingFlags.Instance)
        For Each _property As PropertyInfo In properties
            Try
                If _property.GetIndexParameters.Length = 0 Then
                    If _property.PropertyType.Name.EndsWith(NameOf(Collection)) Then
                        Throw New NotImplementedException()
                    Else
                        Select Case _property.PropertyType.Name
                            Case "Boolean", "DateTime", "Double", "GUID", "Int16", "Int32", "Integer", "String"
                                If _property.GetValue(Left, Nothing) Is _property.GetValue(Right, Nothing) Then
                                    MismatchedPropertyName = _property.Name
                                    Return False
                                End If
                            Case "List`1"
                                Throw New NotImplementedException()
                            Case NameOf(TextSpan)
                                If Not DirectCast(_property.GetValue(Left, Nothing), TextSpan).CompareTextSpan(DirectCast(_property.GetValue(Left, Nothing), TextSpan), MismatchedPropertyName) Then
                                    Return False
                                End If
                            Case Else
                                Throw New ApplicationException($"Unknown _property.Name where _property.GetIndexParameters = 0 {_property.Name} Type {_property.PropertyType.Name})")
                        End Select
                    End If
                Else
                    Select Case _property.Name
                        Case "Item"
                        Case Else
                            Throw New ApplicationException($"Unknown _property.Name where _property.GetIndexParameters > 0 {_property.Name} Type {_property.PropertyType.Name})")
                    End Select
                End If
            Catch ex As Exception When ex.HResult <> (New OperationCanceledException).HResult
                Stop
                Throw
            End Try
        Next
        Return True
    End Function

    <Extension>
    Public Function CompareTextSpan(Left As TextSpan, Right As TextSpan, ByRef MismatchedPropertyName As String) As Boolean
        Dim properties() As PropertyInfo = Left.GetType().GetProperties(BindingFlags.Public Or BindingFlags.Instance)
        For Each _property As PropertyInfo In properties
            Try
                If _property.GetIndexParameters.Length = 0 Then
                    If _property.PropertyType.Name.EndsWith(NameOf(Collection)) Then
                        Throw New NotImplementedException()
                    Else
                        Select Case _property.PropertyType.Name
                            Case "Boolean", "DateTime", "Double", "GUID", "Int16", "Int32", "Integer", "String"
                                If _property.GetValue(Left, Nothing) Is _property.GetValue(Right, Nothing) Then
                                    MismatchedPropertyName = _property.Name
                                    Return False
                                End If
                            Case "List`1"
                                Throw New NotImplementedException()
                            Case Else
                                Throw New ApplicationException($"Unknown _property.Name where _property.GetIndexParameters = 0 {_property.Name} Type {_property.PropertyType.Name})")
                        End Select
                    End If
                Else
                    Select Case _property.Name
                        Case "Item"
                        Case Else
                            Throw New ApplicationException($"Unknown _property.Name where _property.GetIndexParameters > 0 {_property.Name} Type {_property.PropertyType.Name})")
                    End Select
                End If
            Catch ex As Exception When ex.HResult <> (New OperationCanceledException).HResult
                Stop
                Throw
            End Try
        Next
        Return True
    End Function

End Module

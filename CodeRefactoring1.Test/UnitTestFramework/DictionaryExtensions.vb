Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Runtime.CompilerServices

' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

Namespace Roslyn.UnitTestFramework
    Friend Module DictionaryExtensions

        ' Copied from ConcurrentDictionary since IDictionary doesn't have this useful method
        <Extension>
        Public Function GetOrAdd(Of TKey, TValue)(ByVal dictionary As IDictionary(Of TKey, TValue), ByVal key As TKey, ByVal [function] As Func(Of TKey, TValue)) As TValue
            Dim value As TValue = Nothing
            If Not dictionary.TryGetValue(key, value) Then
                value = [function](key)
                dictionary.Add(key, value)
            End If

            Return value
        End Function

        <Extension>
        Public Function GetOrAdd(Of TKey, TValue)(ByVal dictionary As IDictionary(Of TKey, TValue), ByVal key As TKey, ByVal [function] As Func(Of TValue)) As TValue
            Return dictionary.GetOrAdd(key, Function(underscore As TKey) [function]())
        End Function

    End Module
End Namespace

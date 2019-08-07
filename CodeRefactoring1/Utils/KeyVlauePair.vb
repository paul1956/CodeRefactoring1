﻿Imports System.Collections.Generic
Imports System.Runtime.CompilerServices

Namespace Roslyn.Utilities
    Public Module KeyValuePair
        Public Function Create(Of K, V)(ByVal key As K, ByVal value As V) As KeyValuePair(Of K, V)
            Return New KeyValuePair(Of K, V)(key, value)
        End Function

        <Extension>
        Public Sub Deconstruct(Of TKey, TValue)(ByVal keyValuePair_Renamed As KeyValuePair(Of TKey, TValue), <System.Runtime.InteropServices.Out()> ByRef key As TKey, <System.Runtime.InteropServices.Out()> ByRef value As TValue)
            key = keyValuePair_Renamed.Key
            value = keyValuePair_Renamed.Value
        End Sub
    End Module
End Namespace

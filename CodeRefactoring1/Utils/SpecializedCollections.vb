Partial NotInheritable Class SpecializedCollections
    Private Sub New()
    End Sub
    Public Shared ReadOnly EmptyBytes As Byte() = EmptyArray(Of Byte)()
    Public Shared ReadOnly EmptyObjects As Object() = EmptyArray(Of Object)()

    Public Shared Function EmptyArray(Of T)() As T()
        Return Empty.Array(Of T).Instance
    End Function

    Public Shared Function EmptyEnumerator(Of T)() As IEnumerator(Of T)
        Return Empty.Enumerator(Of T).Instance
    End Function

    Public Shared Function EmptyEnumerable(Of T)() As IEnumerable(Of T)
        Return Empty.List(Of T).Instance
    End Function

    Public Shared Function EmptyCollection(Of T)() As ICollection(Of T)
        Return Empty.List(Of T).Instance
    End Function

    Public Shared Function EmptyList(Of T)() As IList(Of T)
        Return Empty.List(Of T).Instance
    End Function

    Public Shared Function EmptyReadOnlyList(Of T)() As IReadOnlyList(Of T)
        Return Empty.List(Of T).Instance
    End Function

    Public Shared Function EmptySet(Of T)() As ISet(Of T)
        Return Empty.Set(Of T).Instance
    End Function

    Public Shared Function EmptyDictionary(Of TKey, TValue)() As IDictionary(Of TKey, TValue)
        Return Empty.Dictionary(Of TKey, TValue).Instance
    End Function

    Public Shared Function SingletonEnumerable(Of T)(value As T) As IEnumerable(Of T)
        Return New Singleton.Collection(Of T)(value)
    End Function

    Public Shared Function SingletonCollection(Of T)(value As T) As ICollection(Of T)
        Return New Singleton.Collection(Of T)(value)
    End Function

    Public Shared Function SingletonEnumerator(Of T)(value As T) As IEnumerator(Of T)
        Return New Singleton.Enumerator(Of T)(value)
    End Function

    Public Shared Function ReadOnlyEnumerable(Of T)(values As IEnumerable(Of T)) As IEnumerable(Of T)
        Return New [ReadOnly].Enumerable(Of IEnumerable(Of T), T)(values)
    End Function

    Public Shared Function ReadOnlyCollection(Of T)(collection As ICollection(Of T)) As ICollection(Of T)
        Return If(collection Is Nothing OrElse collection.Count = 0, EmptyCollection(Of T)(), New [ReadOnly].Collection(Of ICollection(Of T), T)(collection))
    End Function

    Public Shared Function ReadOnlySet(Of T)([set] As ISet(Of T)) As ISet(Of T)
        Return If([set] Is Nothing OrElse [set].Count = 0, EmptySet(Of T)(), New [ReadOnly].Set(Of ISet(Of T), T)([set]))
    End Function

    Public Shared Function ReadOnlySet(Of T)(values As IEnumerable(Of T)) As ISet(Of T)
        If TypeOf values Is ISet(Of T) Then
            Return ReadOnlySet(DirectCast(values, ISet(Of T)))
        End If

        Dim result As HashSet(Of T) = Nothing
        For Each item As T In values
            result = If(result, New HashSet(Of T)())
            result.Add(item)
        Next

        Return ReadOnlySet(result)
    End Function

    Partial Private Class Empty
        Friend Class Array(Of T)
            Public Shared ReadOnly Instance As T() = New T(-1) {}
        End Class

        Friend Class Collection(Of T)
            Inherits Enumerable(Of T)
            Implements ICollection(Of T)

            Public Shared ReadOnly Instance As ICollection(Of T) = New Collection(Of T)()

            Protected Sub New()
            End Sub

#Disable Warning IDE0060 ' Remove unused parameter
            Public Sub Add(item As T)
                Throw New NotSupportedException()
            End Sub

            Public Sub Clear()
            End Sub

            Public Function Contains(item As T) As Boolean
                Return False
            End Function

            Public Sub CopyTo(array As T(), arrayIndex As Integer)
            End Sub

            Public ReadOnly Property Count() As Integer
                Get
                    Return 0
                End Get
            End Property

            Public ReadOnly Property IsReadOnly() As Boolean
                Get
                    Return True
                End Get
            End Property

            Private ReadOnly Property ICollection_Count As Integer Implements ICollection(Of T).Count
                Get
                    Throw New NotImplementedException()
                End Get
            End Property

            Private ReadOnly Property ICollection_IsReadOnly As Boolean Implements ICollection(Of T).IsReadOnly
                Get
                    Throw New NotImplementedException()
                End Get
            End Property

            Public Function Remove(item As T) As Boolean
                Throw New NotSupportedException()
            End Function

            Private Sub ICollection_Add(item As T) Implements ICollection(Of T).Add
                Throw New NotImplementedException()
            End Sub

            Private Sub ICollection_Clear() Implements ICollection(Of T).Clear
                Throw New NotImplementedException()
            End Sub

            Private Function ICollection_Contains(item As T) As Boolean Implements ICollection(Of T).Contains
                Throw New NotImplementedException()
            End Function

            Private Sub ICollection_CopyTo(array() As T, arrayIndex As Integer) Implements ICollection(Of T).CopyTo
                Throw New NotImplementedException()
            End Sub

            Private Function ICollection_Remove(item As T) As Boolean Implements ICollection(Of T).Remove
                Throw New NotImplementedException()
            End Function
#Enable Warning IDE0060 ' Remove unused parameter
        End Class

        Friend Class Dictionary(Of TKey, TValue)
            Inherits Collection(Of KeyValuePair(Of TKey, TValue))
            Implements IDictionary(Of TKey, TValue)

            Public Shared Shadows ReadOnly Instance As IDictionary(Of TKey, TValue) = New Dictionary(Of TKey, TValue)()

            Private Sub New()
            End Sub

#Disable Warning IDE0060 ' Remove unused parameter
            Public Overloads Sub Add(key As TKey, value As TValue)
                Throw New NotSupportedException()
            End Sub

            Public Function ContainsKey(key As TKey) As Boolean
                Return False
            End Function

            Public ReadOnly Property Keys() As ICollection(Of TKey)
                Get
                    Return Collection(Of TKey).Instance
                End Get
            End Property

            Public Overloads Function Remove(key As TKey) As Boolean
                Throw New NotSupportedException()
            End Function

            Public Function TryGetValue(key As TKey, ByRef value As TValue) As Boolean
                value = Nothing
                Return False
            End Function

            Private Function IDictionary_ContainsKey(key As TKey) As Boolean Implements IDictionary(Of TKey, TValue).ContainsKey
                Throw New NotImplementedException()
            End Function

            Private Sub IDictionary_Add(key As TKey, value As TValue) Implements IDictionary(Of TKey, TValue).Add
                Throw New NotImplementedException()
            End Sub

            Private Function IDictionary_Remove(key As TKey) As Boolean Implements IDictionary(Of TKey, TValue).Remove
                Throw New NotImplementedException()
            End Function

            Private Function IDictionary_TryGetValue(key As TKey, ByRef value As TValue) As Boolean Implements IDictionary(Of TKey, TValue).TryGetValue
                Throw New NotImplementedException()
            End Function

            Public ReadOnly Property Values() As ICollection(Of TValue)
                Get
                    Return Collection(Of TValue).Instance
                End Get
            End Property

            Default Public Property Item(key As TKey) As TValue
                Get
                    Throw New NotSupportedException()
                End Get

                Set
                    Throw New NotSupportedException()
                End Set
            End Property

            Property IDictionary_Item(key As TKey) As TValue Implements IDictionary(Of TKey, TValue).Item
                Get
                    Throw New NotImplementedException()
                End Get
                Set(value As TValue)
                    Throw New NotImplementedException()
                End Set
            End Property

            Private ReadOnly Property IDictionary_Keys As ICollection(Of TKey) Implements IDictionary(Of TKey, TValue).Keys
                Get
                    Throw New NotImplementedException()
                End Get
            End Property

            Private ReadOnly Property IDictionary_Values As ICollection(Of TValue) Implements IDictionary(Of TKey, TValue).Values
                Get
                    Throw New NotImplementedException()
                End Get
            End Property
        End Class

        Friend Class Enumerable(Of T)
            Implements IEnumerable(Of T)

            ' PERF: cache the instance of enumerator.
            ' accessing a generic static field is kinda slow from here,
            ' but since empty enumerable are singletons, there is no harm in having
            ' one extra instance field
            Private ReadOnly _enumerator As IEnumerator(Of T) = Enumerator(Of T).Instance

            Public Function GetEnumerator() As IEnumerator(Of T)
                Return Me._enumerator
            End Function

            Private Function System_Collections_IEnumerable_GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
                Return Me.GetEnumerator()
            End Function

            Private Function IEnumerable_GetEnumerator() As IEnumerator(Of T) Implements IEnumerable(Of T).GetEnumerator
                Throw New NotImplementedException()
            End Function
#Enable Warning IDE0060 ' Remove unused parameter
        End Class

        Friend Class Enumerator
            Implements IEnumerator

            Public Shared ReadOnly Instance As IEnumerator = New Enumerator()

            Protected Sub New()
            End Sub

            Public ReadOnly Property Current() As Object
                Get
                    Throw New InvalidOperationException()
                End Get
            End Property

            Private ReadOnly Property IEnumerator_Current As Object Implements IEnumerator.Current
                Get
                    Throw New NotImplementedException()
                End Get
            End Property

            Public Function MoveNext() As Boolean
                Return False
            End Function

            Public Sub Reset()
                Throw New InvalidOperationException()
            End Sub

            Private Function IEnumerator_MoveNext() As Boolean Implements IEnumerator.MoveNext
                Throw New NotImplementedException()
            End Function

            Private Sub IEnumerator_Reset() Implements IEnumerator.Reset
                Throw New NotImplementedException()
            End Sub
        End Class

        Friend Class Enumerator(Of T)
            Inherits Enumerator
            Implements IEnumerator(Of T)

            Public Shared Shadows ReadOnly Instance As IEnumerator(Of T) = New Enumerator(Of T)()

            Protected Sub New()
            End Sub

            Public Shadows ReadOnly Property Current() As T
                Get
                    Throw New InvalidOperationException()
                End Get
            End Property

            Private ReadOnly Property IEnumerator_Current As T Implements IEnumerator(Of T).Current
                Get
                    Throw New NotImplementedException()
                End Get
            End Property

            Public Sub Dispose()
            End Sub

            Private Sub IDisposable_Dispose() Implements IDisposable.Dispose
                Throw New NotImplementedException()
            End Sub
        End Class

        Friend Class List(Of T)
            Inherits Collection(Of T)
            Implements IList(Of T)
            Implements IReadOnlyList(Of T)

            Public Shared Shadows ReadOnly Instance As New List(Of T)()

            Protected Sub New()
            End Sub

#Disable Warning IDE0060 ' Remove unused parameter
            Public Function IndexOf(item As T) As Integer
                Return -1
            End Function

            Public Sub Insert(index As Integer, item As T)
                Throw New NotSupportedException()
            End Sub

            Public Sub RemoveAt(index As Integer)
                Throw New NotSupportedException()
            End Sub

            Private Function IList_IndexOf(item As T) As Integer Implements IList(Of T).IndexOf
                Throw New NotImplementedException()
            End Function

            Private Sub IList_Insert(index As Integer, item As T) Implements IList(Of T).Insert
                Throw New NotImplementedException()
            End Sub

            Private Sub IList_RemoveAt(index As Integer) Implements IList(Of T).RemoveAt
                Throw New NotImplementedException()
            End Sub

            Default Public Property Item(index As Integer) As T
                Get
                    Throw New ArgumentOutOfRangeException(NameOf(index))
                End Get

                Set
                    Throw New NotSupportedException()
                End Set
            End Property

            ReadOnly Property IReadOnlyList_Item(index As Integer) As T Implements IReadOnlyList(Of T).Item
                Get
                    Throw New NotImplementedException()
                End Get
            End Property

            Private ReadOnly Property IReadOnlyCollection_Count As Integer Implements IReadOnlyCollection(Of T).Count
                Get
                    Throw New NotImplementedException()
                End Get
            End Property

            Property IList_Item(index As Integer) As T Implements IList(Of T).Item
                Get
                    Throw New NotImplementedException()
                End Get
                Set(value As T)
                    Throw New NotImplementedException()
                End Set
            End Property
#Enable Warning IDE0060 ' Remove unused parameter
        End Class

        Friend Class [Set](Of T)
            Inherits Collection(Of T)
            Implements ISet(Of T)

            Public Shared Shadows ReadOnly Instance As ISet(Of T) = New [Set](Of T)()

            Protected Sub New()
            End Sub

#Disable Warning IDE0060 ' Remove unused parameter
            Public Shadows Function Add(item As T) As Boolean
                Throw New NotImplementedException()
            End Function

            Public Sub ExceptWith(other As IEnumerable(Of T))
                Throw New NotImplementedException()
            End Sub

            Public Sub IntersectWith(other As IEnumerable(Of T))
                Throw New NotImplementedException()
            End Sub

            Public Function IsProperSubsetOf(other As IEnumerable(Of T)) As Boolean
                Throw New NotImplementedException()
            End Function

            Public Function IsProperSupersetOf(other As IEnumerable(Of T)) As Boolean
                Throw New NotImplementedException()
            End Function

            Public Function IsSubsetOf(other As IEnumerable(Of T)) As Boolean
                Throw New NotImplementedException()
            End Function

            Public Function IsSupersetOf(other As IEnumerable(Of T)) As Boolean
                Throw New NotImplementedException()
            End Function

            Public Function Overlaps(other As IEnumerable(Of T)) As Boolean
                Throw New NotImplementedException()
            End Function

            Public Function SetEquals(other As IEnumerable(Of T)) As Boolean
                Throw New NotImplementedException()
            End Function

            Public Sub SymmetricExceptWith(other As IEnumerable(Of T))
                Throw New NotImplementedException()
            End Sub

            Public Sub UnionWith(other As IEnumerable(Of T))
                Throw New NotImplementedException()
            End Sub

            Public Shadows Function GetEnumerator() As System.Collections.IEnumerator
                Return [Set](Of T).Instance.GetEnumerator()
            End Function

            Private Function ISet_Add(item As T) As Boolean Implements ISet(Of T).Add
                Throw New NotImplementedException()
            End Function

            Private Sub ISet_UnionWith(other As IEnumerable(Of T)) Implements ISet(Of T).UnionWith
                Throw New NotImplementedException()
            End Sub

            Private Sub ISet_IntersectWith(other As IEnumerable(Of T)) Implements ISet(Of T).IntersectWith
                Throw New NotImplementedException()
            End Sub

            Private Sub ISet_ExceptWith(other As IEnumerable(Of T)) Implements ISet(Of T).ExceptWith
                Throw New NotImplementedException()
            End Sub

            Private Sub ISet_SymmetricExceptWith(other As IEnumerable(Of T)) Implements ISet(Of T).SymmetricExceptWith
                Throw New NotImplementedException()
            End Sub

            Private Function ISet_IsSubsetOf(other As IEnumerable(Of T)) As Boolean Implements ISet(Of T).IsSubsetOf
                Throw New NotImplementedException()
            End Function

            Private Function ISet_IsSupersetOf(other As IEnumerable(Of T)) As Boolean Implements ISet(Of T).IsSupersetOf
                Throw New NotImplementedException()
            End Function

            Private Function ISet_IsProperSupersetOf(other As IEnumerable(Of T)) As Boolean Implements ISet(Of T).IsProperSupersetOf
                Throw New NotImplementedException()
            End Function

            Private Function ISet_IsProperSubsetOf(other As IEnumerable(Of T)) As Boolean Implements ISet(Of T).IsProperSubsetOf
                Throw New NotImplementedException()
            End Function

            Private Function ISet_Overlaps(other As IEnumerable(Of T)) As Boolean Implements ISet(Of T).Overlaps
                Throw New NotImplementedException()
            End Function

            Private Function ISet_SetEquals(other As IEnumerable(Of T)) As Boolean Implements ISet(Of T).SetEquals
                Throw New NotImplementedException()
            End Function
#Enable Warning IDE0060 ' Remove unused parameter
        End Class
    End Class

    Partial Private NotInheritable Class [ReadOnly]
        Private Sub New()
        End Sub
        Friend Class Collection(Of TUnderlying As ICollection(Of T), T)
            Inherits Enumerable(Of TUnderlying, T)
            Implements ICollection(Of T)

            Public Sub New(underlying As TUnderlying)
                MyBase.New(underlying)
            End Sub

#Disable Warning IDE0060 ' Remove unused parameter
            Public Sub Add(item As T)
                Throw New NotSupportedException()
            End Sub

            Public Sub Clear()
                Throw New NotSupportedException()
            End Sub

            Public Function Contains(item As T) As Boolean
                Return Me.Underlying.Contains(item)
            End Function

            Public Sub CopyTo(array As T(), arrayIndex As Integer)
                Me.Underlying.CopyTo(array, arrayIndex)
            End Sub

            Public ReadOnly Property Count() As Integer
                Get
                    Return Me.Underlying.Count
                End Get
            End Property

            Public ReadOnly Property IsReadOnly() As Boolean
                Get
                    Return True
                End Get
            End Property

            Private ReadOnly Property ICollection_Count As Integer Implements ICollection(Of T).Count
                Get
                    Throw New NotImplementedException()
                End Get
            End Property

            Private ReadOnly Property ICollection_IsReadOnly As Boolean Implements ICollection(Of T).IsReadOnly
                Get
                    Throw New NotImplementedException()
                End Get
            End Property

            Public Function Remove(item As T) As Boolean
                Throw New NotSupportedException()
            End Function

            Private Sub ICollection_Add(item As T) Implements ICollection(Of T).Add
                Throw New NotImplementedException()
            End Sub

            Private Sub ICollection_Clear() Implements ICollection(Of T).Clear
                Throw New NotImplementedException()
            End Sub

            Private Function ICollection_Contains(item As T) As Boolean Implements ICollection(Of T).Contains
                Throw New NotImplementedException()
            End Function

            Private Sub ICollection_CopyTo(array() As T, arrayIndex As Integer) Implements ICollection(Of T).CopyTo
                Throw New NotImplementedException()
            End Sub

            Private Function ICollection_Remove(item As T) As Boolean Implements ICollection(Of T).Remove
                Throw New NotImplementedException()
            End Function
#Enable Warning IDE0060 ' Remove unused parameter
        End Class

        Friend Class Enumerable(Of TUnderlying As IEnumerable)
            Implements IEnumerable

            Protected ReadOnly Underlying As TUnderlying

            Public Sub New(underlying As TUnderlying)
                Me.Underlying = underlying
            End Sub

            Public Function GetEnumerator() As IEnumerator
                Return Me.Underlying.GetEnumerator()
            End Function

            Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
                Throw New NotImplementedException()
            End Function
        End Class

        Friend Class Enumerable(Of TUnderlying As IEnumerable(Of T), T)
            Inherits Enumerable(Of TUnderlying)
            Implements IEnumerable(Of T)

            Public Sub New(underlying As TUnderlying)
                MyBase.New(underlying)
            End Sub

            Public Shadows Function GetEnumerator() As IEnumerator(Of T)
                Return Me.Underlying.GetEnumerator()
            End Function

            Private Function IEnumerable_GetEnumerator() As IEnumerator(Of T) Implements IEnumerable(Of T).GetEnumerator
                Throw New NotImplementedException()
            End Function
        End Class

        Friend Class [Set](Of TUnderlying As ISet(Of T), T)
            Inherits Collection(Of TUnderlying, T)
            Implements ISet(Of T)

            Public Sub New(underlying As TUnderlying)
                MyBase.New(underlying)
            End Sub

#Disable Warning IDE0060 ' Remove unused parameter
            Public Shadows Function Add(item As T) As Boolean
                Throw New NotSupportedException()
            End Function

            Public Sub ExceptWith(other As IEnumerable(Of T))
                Throw New NotSupportedException()
            End Sub

            Public Sub IntersectWith(other As IEnumerable(Of T))
                Throw New NotSupportedException()
            End Sub
#Enable Warning IDE0060 ' Remove unused parameter

            Public Function IsProperSubsetOf(other As IEnumerable(Of T)) As Boolean
                Return Me.Underlying.IsProperSubsetOf(other)
            End Function

            Public Function IsProperSupersetOf(other As IEnumerable(Of T)) As Boolean
                Return Me.Underlying.IsProperSupersetOf(other)
            End Function

            Public Function IsSubsetOf(other As IEnumerable(Of T)) As Boolean
                Return Me.Underlying.IsSubsetOf(other)
            End Function

            Public Function IsSupersetOf(other As IEnumerable(Of T)) As Boolean
                Return Me.Underlying.IsSupersetOf(other)
            End Function

            Public Function Overlaps(other As IEnumerable(Of T)) As Boolean
                Return Me.Underlying.Overlaps(other)
            End Function

            Public Function SetEquals(other As IEnumerable(Of T)) As Boolean
                Return Me.Underlying.SetEquals(other)
            End Function

#Disable Warning IDE0060 ' Remove unused parameter
            Public Sub SymmetricExceptWith(other As IEnumerable(Of T))
                Throw New NotSupportedException()
            End Sub

            Public Sub UnionWith(other As IEnumerable(Of T))
                Throw New NotSupportedException()
            End Sub

            Private Function ISet_Add(item As T) As Boolean Implements ISet(Of T).Add
                Throw New NotImplementedException()
            End Function

            Private Sub ISet_UnionWith(other As IEnumerable(Of T)) Implements ISet(Of T).UnionWith
                Throw New NotImplementedException()
            End Sub

            Private Sub ISet_IntersectWith(other As IEnumerable(Of T)) Implements ISet(Of T).IntersectWith
                Throw New NotImplementedException()
            End Sub

            Private Sub ISet_ExceptWith(other As IEnumerable(Of T)) Implements ISet(Of T).ExceptWith
                Throw New NotImplementedException()
            End Sub

            Private Sub ISet_SymmetricExceptWith(other As IEnumerable(Of T)) Implements ISet(Of T).SymmetricExceptWith
                Throw New NotImplementedException()
            End Sub

            Private Function ISet_IsSubsetOf(other As IEnumerable(Of T)) As Boolean Implements ISet(Of T).IsSubsetOf
                Throw New NotImplementedException()
            End Function

            Private Function ISet_IsSupersetOf(other As IEnumerable(Of T)) As Boolean Implements ISet(Of T).IsSupersetOf
                Throw New NotImplementedException()
            End Function

            Private Function ISet_IsProperSupersetOf(other As IEnumerable(Of T)) As Boolean Implements ISet(Of T).IsProperSupersetOf
                Throw New NotImplementedException()
            End Function

            Private Function ISet_IsProperSubsetOf(other As IEnumerable(Of T)) As Boolean Implements ISet(Of T).IsProperSubsetOf
                Throw New NotImplementedException()
            End Function

            Private Function ISet_Overlaps(other As IEnumerable(Of T)) As Boolean Implements ISet(Of T).Overlaps
                Throw New NotImplementedException()
            End Function

            Private Function ISet_SetEquals(other As IEnumerable(Of T)) As Boolean Implements ISet(Of T).SetEquals
                Throw New NotImplementedException()
            End Function
        End Class
#Enable Warning IDE0060 ' Remove unused parameter

    End Class

    Partial Private NotInheritable Class Singleton
        Private Sub New()
        End Sub
        Friend NotInheritable Class Collection(Of T)
            Implements ICollection(Of T)
            Implements IReadOnlyCollection(Of T)

            Private ReadOnly _loneValue As T

            Public Sub New(value As T)
                Me._loneValue = value
            End Sub

#Disable Warning IDE0060 ' Remove unused parameter
            Public Sub Add(item As T)
                Throw New NotSupportedException()
            End Sub

            Public Sub Clear()
                Throw New NotSupportedException()
            End Sub
#Enable Warning IDE0060 ' Remove unused parameter

            Public Function Contains(item As T) As Boolean
                Return EqualityComparer(Of T).[Default].Equals(Me._loneValue, item)
            End Function

            Public Sub CopyTo(array As T(), arrayIndex As Integer)
                array(arrayIndex) = Me._loneValue
            End Sub

            Public ReadOnly Property Count() As Integer
                Get
                    Return 1
                End Get
            End Property

            Public ReadOnly Property IsReadOnly() As Boolean
                Get
                    Return True
                End Get
            End Property

            Private ReadOnly Property ICollection_Count As Integer Implements ICollection(Of T).Count
                Get
                    Throw New NotImplementedException()
                End Get
            End Property

            Private ReadOnly Property ICollection_IsReadOnly As Boolean Implements ICollection(Of T).IsReadOnly
                Get
                    Throw New NotImplementedException()
                End Get
            End Property

            Private ReadOnly Property IReadOnlyCollection_Count As Integer Implements IReadOnlyCollection(Of T).Count
                Get
                    Throw New NotImplementedException()
                End Get
            End Property

#Disable Warning IDE0060 ' Remove unused parameter
            Public Function Remove(item As T) As Boolean
#Enable Warning IDE0060 ' Remove unused parameter
                Throw New NotSupportedException()
            End Function

            Public Function GetEnumerator() As IEnumerator(Of T)
                Return New Enumerator(Of T)(Me._loneValue)
            End Function

            Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
                Return Me.GetEnumerator()
            End Function

            Private Sub ICollection_Add(item As T) Implements ICollection(Of T).Add
                Throw New NotImplementedException()
            End Sub

            Private Sub ICollection_Clear() Implements ICollection(Of T).Clear
                Throw New NotImplementedException()
            End Sub

            Private Function ICollection_Contains(item As T) As Boolean Implements ICollection(Of T).Contains
                Throw New NotImplementedException()
            End Function

            Private Sub ICollection_CopyTo(array() As T, arrayIndex As Integer) Implements ICollection(Of T).CopyTo
                Throw New NotImplementedException()
            End Sub

            Private Function ICollection_Remove(item As T) As Boolean Implements ICollection(Of T).Remove
                Throw New NotImplementedException()
            End Function

            Private Function IEnumerable_GetEnumerator1() As IEnumerator(Of T) Implements IEnumerable(Of T).GetEnumerator
                Throw New NotImplementedException()
            End Function
        End Class
        Friend Class Enumerator(Of T)
            Implements IEnumerator(Of T)

            Private _moveNextCalled As Boolean

            Public Sub New(value As T)
                Me.Current = value
                Me._moveNextCalled = False
            End Sub

            Public ReadOnly Property Current() As T

            Private ReadOnly Property IEnumerator_Current() As Object Implements IEnumerator.Current
                Get
                    Return Me.Current
                End Get
            End Property

            Private ReadOnly Property IEnumerator_Current1 As T Implements IEnumerator(Of T).Current
                Get
                    Throw New NotImplementedException()
                End Get
            End Property

            Public Sub Dispose()
            End Sub

            Public Function MoveNext() As Boolean
                If Not Me._moveNextCalled Then
                    Me._moveNextCalled = True
                    Return True
                End If

                Return False
            End Function

            Public Sub Reset()
                Me._moveNextCalled = False
            End Sub

            Private Sub IDisposable_Dispose() Implements IDisposable.Dispose
                Throw New NotImplementedException()
            End Sub

            Private Function IEnumerator_MoveNext() As Boolean Implements IEnumerator.MoveNext
                Throw New NotImplementedException()
            End Function

            Private Sub IEnumerator_Reset() Implements IEnumerator.Reset
                Throw New NotImplementedException()
            End Sub
        End Class

    End Class

End Class

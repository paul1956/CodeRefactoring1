Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Public NotInheritable Class InsertionResult

    ''' <summary>
    ''' Gets the context the insertion is invoked at.
    ''' </summary>
    Private privateContext As CodeRefactoringContext

    ''' <summary>
    ''' Gets the location of the type part the node should be inserted to.
    ''' </summary>
    Private privateLocation As Location

    ''' <summary>
    ''' Gets the node that should be inserted.
    ''' </summary>
    Private privateNode As SyntaxNode

    ''' <summary>
    ''' Gets the type the node should be inserted to.
    ''' </summary>
    Private privateType As INamedTypeSymbol

    Public Sub New(ByVal context As CodeRefactoringContext, ByVal node As SyntaxNode, ByVal type As INamedTypeSymbol, ByVal location As Location)
        Me.Context = context
        Me.Node = node
        Me.Type = type
        Me.Location = location
    End Sub

    Public Property Context() As CodeRefactoringContext
        Get
            Return Me.privateContext
        End Get
        Private Set(ByVal value As CodeRefactoringContext)
            Me.privateContext = value
        End Set
    End Property

    Public Property Location() As Location
        Get
            Return Me.privateLocation
        End Get
        Private Set(ByVal value As Location)
            Me.privateLocation = value
        End Set
    End Property

    Public Property Node() As SyntaxNode
        Get
            Return Me.privateNode
        End Get
        Private Set(ByVal value As SyntaxNode)
            Me.privateNode = value
        End Set
    End Property

    Public Property Type() As INamedTypeSymbol
        Get
            Return Me.privateType
        End Get
        Private Set(ByVal value As INamedTypeSymbol)
            Me.privateType = value
        End Set
    End Property

    Public Shared Function GuessCorrectLocation(ByVal context As CodeRefactoringContext, ByVal locations As Immutable.ImmutableArray(Of Location)) As Location
        For Each Loc As Location In locations
            If context.Document.FilePath = Loc.SourceTree.FilePath Then
                Return Loc
            End If
        Next
        Return locations(0)
    End Function

End Class
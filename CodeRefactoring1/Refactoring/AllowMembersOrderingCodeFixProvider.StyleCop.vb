' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Imports Microsoft.CodeAnalysis.CodeFixes

Namespace Refactoring

    <ExportCodeFixProvider(LanguageNames.VisualBasic, Name:=NameOf(StyleCopAllowMembersOrderingCodeFixProvider)), Composition.Shared>
    Public Class StyleCopAllowMembersOrderingCodeFixProvider
        Inherits BaseAllowMembersOrderingCodeFixProvider

        Public Sub New()
            MyBase.New("Order {0}'s members following StyleCop patterns")
        End Sub

        Protected Overrides Function GetMemberDeclarationComparer(document As Document, cancellationToken As CancellationToken) As IComparer(Of DeclarationStatementSyntax)
            Return New StyleCopMembersComparer
        End Function

        Friend Class StyleCopMembersComparer
            Implements IComparer(Of DeclarationStatementSyntax)

            ReadOnly accessLevelRank As New Dictionary(Of SyntaxKind, Integer) From
                {
                    {SyntaxKind.PublicKeyword, -4},
                    {SyntaxKind.FriendKeyword, -2},
                    {SyntaxKind.ProtectedKeyword, 1},
                    {SyntaxKind.PrivateKeyword, 2}
                }

            ReadOnly specialModifierRank As New Dictionary(Of SyntaxKind, Integer) From
                {
                    {SyntaxKind.ConstKeyword, 1},
                    {SyntaxKind.SharedKeyword, 2}
                }

            ReadOnly typeRank As New Dictionary(Of Type, Integer) From
                                        {
                    {GetType(FieldDeclarationSyntax), 1},
                    {GetType(ConstructorBlockSyntax), 2},
                    {GetType(DelegateStatementSyntax), 4},
                    {GetType(EventStatementSyntax), 5},
                    {GetType(EventBlockSyntax), 6},
                    {GetType(EnumBlockSyntax), 7},
                    {GetType(InterfaceBlockSyntax), 8},
                    {GetType(PropertyBlockSyntax), 9},
                    {GetType(PropertyStatementSyntax), 9}, ' For VB Auto Properties
                    {GetType(OperatorBlockSyntax), 11},
                    {GetType(MethodBlockSyntax), 12},
                    {GetType(StructureBlockSyntax), 13},
                    {GetType(ClassBlockSyntax), 14}
                }

            Private Shared Function GetFieldName(declarations As SeparatedSyntaxList(Of VariableDeclaratorSyntax)) As String
                Dim names As IEnumerable = From declaration In declarations
                                           From name In declaration.Names
                                           Select name.Identifier.Text
                Return String.Join("", names)
            End Function

            Private Shared Function SumRankPoints(tokens As SyntaxTokenList, rank As Dictionary(Of SyntaxKind, Integer), defaultSumValue As Integer) As Integer
                Dim points As Integer = tokens.Sum(Function(t As SyntaxToken) If(rank.ContainsKey(t.Kind), rank(t.Kind), 0))

                Return If(points = 0, defaultSumValue, points)
            End Function

            Private Function GetAccessLevelPoints(tokens As SyntaxTokenList) As Integer
                Return SumRankPoints(tokens, accessLevelRank, accessLevelRank(SyntaxKind.PrivateKeyword))
            End Function

            Private Function GetName(node As SyntaxNode) As String
                If TypeOf node Is FieldDeclarationSyntax Then
                    Return GetFieldName(DirectCast(node, FieldDeclarationSyntax).Declarators)
                End If
                If TypeOf node Is PropertyStatementSyntax Then
                    Return DirectCast(node, PropertyStatementSyntax).Identifier.Text
                End If
                If TypeOf node Is PropertyBlockSyntax Then
                    Return GetName(DirectCast(node, PropertyBlockSyntax).PropertyStatement)
                End If
                If TypeOf node Is MethodBlockSyntax Then
                    Return DirectCast(node, MethodBlockSyntax).SubOrFunctionStatement.Identifier.Text
                End If
                If TypeOf node Is SubNewStatementSyntax Then
                    Return "New"
                End If
                If TypeOf node Is EnumBlockSyntax Then
                    Return DirectCast(node, EnumBlockSyntax).EnumStatement.Identifier.Text
                End If
                If TypeOf node Is InterfaceBlockSyntax Then
                    Return DirectCast(node, InterfaceBlockSyntax).InterfaceStatement.Identifier.Text
                End If
                If TypeOf node Is EventStatementSyntax Then
                    Return DirectCast(node, EventStatementSyntax).Identifier.Text
                End If
                If TypeOf node Is EventBlockSyntax Then
                    Return GetName(DirectCast(node, EventBlockSyntax).EventStatement)
                End If
                If TypeOf node Is OperatorBlockSyntax Then
                    Return DirectCast(node, OperatorBlockSyntax).OperatorStatement.OperatorToken.Text
                End If
                If TypeOf node Is DelegateStatementSyntax Then
                    Return DirectCast(node, DelegateStatementSyntax).Identifier.Text
                End If
                If TypeOf node Is ClassBlockSyntax Then
                    Return DirectCast(node, ClassBlockSyntax).ClassStatement.Identifier.Text
                End If
                If TypeOf node Is StructureBlockSyntax Then
                    Return DirectCast(node, StructureBlockSyntax).StructureStatement.Identifier.Text
                End If
                Return ""
            End Function

            Private Function GetRankPoints(node As DeclarationStatementSyntax) As Integer
                Dim points As Integer = 0
                If Not typeRank.TryGetValue(node.GetType(), points) Then
                    Return 0
                End If
                Return points
            End Function

            Private Function GetSpecialModifierPoints(tokens As SyntaxTokenList) As Integer
                Return SumRankPoints(tokens, specialModifierRank, 100)
            End Function

            Public Function Compare(x As DeclarationStatementSyntax, y As DeclarationStatementSyntax) As Integer Implements IComparer(Of DeclarationStatementSyntax).Compare
                If x Is Nothing AndAlso y Is Nothing Then Return 0
                If x Is Nothing Then Return 1
                If y Is Nothing Then Return -1
                If x.Equals(y) Then Return 0

                Dim comparedPoints As Integer = GetRankPoints(x).CompareTo(GetRankPoints(y))
                If comparedPoints <> 0 Then Return comparedPoints

                Dim xModifiers As SyntaxTokenList = x.GetModifiers
                Dim yModifiers As SyntaxTokenList = y.GetModifiers
                comparedPoints = GetAccessLevelPoints(xModifiers).CompareTo(GetAccessLevelPoints(yModifiers))
                If comparedPoints <> 0 Then Return comparedPoints

                comparedPoints = GetSpecialModifierPoints(xModifiers).CompareTo(GetSpecialModifierPoints(yModifiers))
                If comparedPoints <> 0 Then Return comparedPoints

                Return GetName(x).CompareTo(GetName(y))
            End Function

        End Class

    End Class

End Namespace

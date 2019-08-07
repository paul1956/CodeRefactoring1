Imports System.Runtime.CompilerServices

Public Module IAssemblySymbolExtensions

    <Extension>
    Public Function IsSameAssemblyOrHasFriendAccessTo(ByVal assembly As IAssemblySymbol, ByVal toAssembly As IAssemblySymbol) As Boolean
        Return Equals(assembly, toAssembly) OrElse (assembly.IsInteractive AndAlso toAssembly.IsInteractive) OrElse toAssembly.GivesAccessTo(assembly)
    End Function

End Module
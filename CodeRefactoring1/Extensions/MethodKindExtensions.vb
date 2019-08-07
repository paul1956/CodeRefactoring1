' Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

'Namespace Microsoft.CodeAnalysis.Shared.Extensions
Imports System.Runtime.CompilerServices

Public Module MethodKindExtensions
    <Extension>
    Public Function IsPropertyAccessor(ByVal kind As MethodKind) As Boolean
        Return kind = MethodKind.PropertyGet OrElse kind = MethodKind.PropertySet
    End Function
End Module
'End Namespace

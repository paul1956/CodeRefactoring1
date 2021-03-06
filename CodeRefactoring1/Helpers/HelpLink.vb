﻿' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Public Module HelpLink

    Public Function ForDiagnostic(diagnosticId As String) As String
        Return $"https://code-cracker.github.io/diagnostics/{diagnosticId}.html"
    End Function

End Module

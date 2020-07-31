' Licensed to the .NET Foundation under one or more agreements.
' The .NET Foundation licenses this file to you under the MIT license.
' See the LICENSE file in the project root for more information.

Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Diagnostics.CodeAnalysis

Imports ConfOxide

Namespace EditorExtensions.Settings

    Public NotInheritable Class VBSettings
        Inherits SettingsBase(Of VBSettings)

        <SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")>
        Public Shared ReadOnly Instance As New VBSettings()

        Private m_General As GeneralSettings

        Public Property General() As GeneralSettings
            Get
                Return m_General
            End Get
            Private Set
                m_General = Value
            End Set
        End Property

    End Class

    Public NotInheritable Class GeneralSettings
        Inherits SettingsBase(Of GeneralSettings)
        Implements INameOfSettings

        Public Property AllMessagesToOutputWindow As Boolean = True

        <SuppressMessage("Style", "IDE0032:Use auto property", Justification:="<Pending>")>
        Private _csvIgnoreList As String = "Microsoft,System,MS"

        Public Property CsvIgnoreList As String Implements INameOfSettings.CsvIgnoreList
            Get
                Return _csvIgnoreList
            End Get
            Set(value As String)
                _csvIgnoreList = value
            End Set
        End Property

    End Class

    Public Interface INameOfSettings
        Property CsvIgnoreList As String
    End Interface

End Namespace

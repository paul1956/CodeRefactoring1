Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On


Imports ConfOxide
Imports System.Diagnostics.CodeAnalysis

Namespace EditorExtensions.Settings
    Public NotInheritable Class VBSettings
        Inherits SettingsBase(Of VBSettings)
        <SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes")>
        Public Shared ReadOnly Instance As New VBSettings()
        Private m_General As GeneralSettings

        Public Property General() As GeneralSettings
            Get
                Return Me.m_General
            End Get
            Private Set
                Me.m_General = Value
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
                Return Me._csvIgnoreList
            End Get
            Set(value As String)
                Me._csvIgnoreList = value
            End Set
        End Property

    End Class

    Public Interface INameOfSettings
        Property CsvIgnoreList As String
    End Interface

End Namespace

Option Compare Text
Option Explicit On
Option Infer Off
Option Strict On

Public NotInheritable Class AnalyzerSettings
    Public Shared ReadOnly Instance As New AnalyzerSettings()
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
    Implements INameOfSettings

#Disable Warning IDE0032 ' Use auto property
    Private _CVS_IgnoreList As String = "Microsoft,System,MS"
#Enable Warning IDE0032 ' Use auto property

    Private Interface INameOfSettings
        Property CVS_IgnoreList As String
    End Interface

    Public Property CVS_IgnoreList As String Implements INameOfSettings.CVS_IgnoreList
        Get
            Return Me._CVS_IgnoreList
        End Get
        Set(value As String)
            Me._CVS_IgnoreList = value
        End Set
    End Property

End Class
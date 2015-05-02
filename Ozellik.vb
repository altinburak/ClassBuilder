Public Class Ozellik
    Private mOzellikAd As String
    Private mOzellikTur As String

    Public Property OzellikAd() As String
        Get
            Return mOzellikAd
        End Get
        Set(ByVal Value As String)
            mOzellikAd = Value
        End Set
    End Property

    Public Property OzellikTur() As String
        Get
            Return mOzellikTur
        End Get
        Set(ByVal Value As String)
            mOzellikTur = Value
        End Set
    End Property

    Public Sub New(ByVal Ad As String, ByVal tur As String)
        OzellikAd = Ad
        OzellikTur = tur
    End Sub

End Class

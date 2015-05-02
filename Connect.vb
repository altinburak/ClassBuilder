Public Class Connect
    Private mServer As String
    Private mDatabaseAd As String
    Private mUserID As String
    Private mPassword As String
    Private mTur As String

    Public Property Server() As String
        Get
            Return mServer
        End Get
        Set(ByVal Value As String)
            mServer = Value
        End Set
    End Property

    Public Property DatabaseAd() As String
        Get
            Return mDatabaseAd
        End Get
        Set(ByVal Value As String)
            mDatabaseAd = Value
        End Set
    End Property

    Public Property UserID() As String
        Get
            Return mUserID
        End Get
        Set(ByVal Value As String)
            mUserID = Value
        End Set
    End Property

    Public Property Password() As String
        Get
            Return mPassword
        End Get
        Set(ByVal Value As String)
            mPassword = Value
        End Set
    End Property

    Public Property Tur() As String
        Get
            Return mTur
        End Get
        Set(ByVal Value As String)
            mTur = Value
        End Set
    End Property
End Class

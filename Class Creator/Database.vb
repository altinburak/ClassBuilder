Public Class Database
    Private mDatabaseName As String
    Private mTablolar As tabloCollection

    Public Property DatabaseName() As String
        Get
            Return mDatabaseName
        End Get
        Set(ByVal Value As String)
            mDatabaseName = Value
        End Set
    End Property

    Public Property Tablolar() As TabloCollection
        Get
            If mTablolar Is Nothing Then
                Return Tablo.GetirTablolar(DatabaseName)
            Else
                Return mTablolar
            End If
        End Get
        Set(ByVal Value As TabloCollection)
            mTablolar = Value
        End Set
    End Property

    Public Shared Function GetDatabases() As DatabaseCollection
        Dim dc As New DatabaseCollection
        Dim Con As New SqlClient.SqlConnection(Tools.ConStr)
        Dim Com As New SqlClient.SqlCommand("sp_databases", Con)
        Dim dr As SqlClient.SqlDataReader
        Con.Open()
        dr = Com.ExecuteReader
        While dr.Read
            Dim d As New Database
            d.DatabaseName = dr("DATABASE_NAME")
            d.Tablolar = Nothing
            dc.Add(d)
        End While
        Con.Close()
        Return dc
    End Function

    Public Overrides Function tostring() As String
        Return DatabaseName
    End Function
End Class

Public Class DatabaseCollection
    Inherits CollectionBase

    Public Sub Add(ByVal t As Database)
        MyBase.List.Add(t)
    End Sub
    Default Public Property Saglayici(ByVal index As Integer) As Database
        Get
            Return MyBase.InnerList(index)
        End Get
        Set(ByVal Value As Database)
            MyBase.InnerList(index) = Value
        End Set
    End Property
End Class



Public Class Tablo
    Private mTabloAdi As String
    Private mKolonlar As KolonCollection

    Public Property TabloAdi() As String
        Get
            Return mTabloAdi
        End Get
        Set(ByVal Value As String)
            mTabloAdi = Value
        End Set
    End Property

    Public Property Kolonlar(ByVal databaseAdi As String) As KolonCollection
        Get
            If mKolonlar Is Nothing Then
                Return Kolon.GetirKolonlar(TabloAdi, databaseAdi)
            Else
                Return mKolonlar
            End If
        End Get
        Set(ByVal Value As KolonCollection)
            mKolonlar = Value
        End Set
    End Property

    Public Shared Function GetirTablolar(ByVal DatabaseAdi As String) As TabloCollection
        Dim tc As New TabloCollection
        Dim Con As New SqlClient.SqlConnection("server=.;Database=" & DatabaseAdi & ";User ID=sa;Password=1;")
        Dim Com1 As New SqlClient.SqlCommand("select name from sysobjects where xtype='U' and not name='dtproperties'", Con)
        Dim dr As SqlClient.SqlDataReader

        Con.Open()
        dr = Com1.ExecuteReader
        While dr.Read
            Dim t As New Tablo
            t.TabloAdi = dr("name")
            t.Kolonlar(DatabaseAdi) = Nothing
            tc.Add(t)
        End While
        Con.Close()
        Return tc
    End Function

    Public Overrides Function tostring() As String
        Return TabloAdi
    End Function
End Class

Public Class TabloCollection
    Inherits CollectionBase

    Public Sub Add(ByVal t As Tablo)
        MyBase.List.Add(t)
    End Sub
    Default Public Property Saglayici(ByVal index As Integer) As Tablo
        Get
            Return MyBase.InnerList(index)
        End Get
        Set(ByVal Value As Tablo)
            MyBase.InnerList(index) = Value
        End Set
    End Property
End Class



Public Class Kolon
    Private mKolonAd As String
    Private mKolonTip As String
    Private mKolonUzunluk As Integer

    Public Property KolonAd() As String
        Get
            Return mKolonAd
        End Get
        Set(ByVal Value As String)
            mKolonAd = Value
        End Set
    End Property

    Public Property KolonTip() As String
        Get
            Return mKolonTip
        End Get
        Set(ByVal Value As String)
            mKolonTip = Value
        End Set
    End Property

    Public Property KolonUzunluk() As Integer
        Get
            Return mKolonUzunluk
        End Get
        Set(ByVal Value As Integer)
            mKolonUzunluk = Value
        End Set
    End Property

    Public Shared Function GetirKolonlar(ByVal tablo1 As String, ByVal DatabaseAdi As String) As KolonCollection
        Dim kcGec As New KolonCollection
        Dim Con As New SqlClient.SqlConnection("server=.;Database=" & DatabaseAdi & ";User ID=sa;Password=1;")
        Dim Com1 As New SqlClient.SqlCommand("select name,xtype from syscolumns where id=(select id from sysobjects where name=@name)", Con)
        Com1.Parameters.Add("@name", tablo1)

        Dim Com2 As New SqlClient.SqlCommand("select name,length from systypes where xtype=@type", Con)

        Dim dr As SqlClient.SqlDataReader

        Con.Open()
        dr = Com1.ExecuteReader
        While dr.Read
            Dim k As New Kolon
            k.KolonAd = dr("name")
            k.KolonTip = dr("xtype")
            k.KolonUzunluk = 0
            kcGec.Add(k)
        End While
        Con.Close()
        dr.Close()
        Dim kc As New KolonCollection
        Try

            Con.Open()
            For Each k As Kolon In kcGec
                Com2.Parameters.Add("@type", k.KolonTip)
                dr = Com2.ExecuteReader
                While dr.Read
                    k.KolonTip = dr("name")
                    k.KolonUzunluk = dr("length")
                    kc.Add(k)
                End While
                dr.Close()
                Com2.Parameters.Clear()
            Next
            Con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Return kc
    End Function

    Public Overrides Function tostring() As String
        Return KolonAd
    End Function
End Class

Public Class KolonCollection
    Inherits CollectionBase

    Public Sub Add(ByVal t As Kolon)
        MyBase.List.Add(t)
    End Sub
    Default Public Property Saglayici(ByVal index As Integer) As Kolon
        Get
            Return MyBase.InnerList(index)
        End Get
        Set(ByVal Value As Kolon)
            MyBase.InnerList(index) = Value
        End Set
    End Property
End Class

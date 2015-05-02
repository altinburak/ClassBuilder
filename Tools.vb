Public Class Tools

    Public Shared ConStr As String = "Data Source=GENL253;Initial Catalog=deneme;User ID=sa;Password=Tb111101"

    Public Shared PrimaryKey As Kolon = Nothing
    Public Shared TotoString As String = ""
    Public Shared ClassAdi As String = ""
    Public Shared SpDahilMi As Boolean = False
    Public Shared IsDeleted As Kolon = Nothing

    Public Shared Function Property_Olustur(ByVal Property_Ad As String, ByVal Property_Tur As String) As String

        Dim sonuc As String
        Property_Tur = Tools.DegiskenCevir(Property_Tur)
        sonuc = "Public property " & Property_Ad & "() as " & Property_Tur & vbCrLf
        sonuc &= vbTab & "Get" & vbCrLf
        sonuc &= vbTab & vbTab & "Return m" & Property_Ad & vbCrLf
        sonuc &= vbTab & "End Get" & vbCrLf
        sonuc &= vbTab & "Set (ByVal Value as " & Property_Tur & ")" & vbCrLf
        sonuc &= vbTab & vbTab & "m" & Property_Ad & " = Value" & vbCrLf
        sonuc &= vbTab & "End Set" & vbCrLf
        sonuc &= "End Property"
        Return sonuc
    End Function

    Public Shared Function Degisken_Olustur(ByVal Degisken_Ad As String, ByVal Degisken_Tur As String) As String
        Dim sonuc As String
        sonuc = "Private m" & Degisken_Ad & " As " & Tools.DegiskenCevir(Degisken_Tur)
        Return sonuc
    End Function

    Public Shared Function Ekle(ByVal TabloAdi As String, ByVal kolonlar As KolonCollection, ByVal Connection As String, ByVal spAdi As String) As String
        Dim sonuc As String
        For Each k As Kolon In kolonlar
            k.KolonTip = Tools.DegiskenCevir(k.KolonTip)
        Next

        'kolonlar.RemoveAt(0)

        sonuc = "Public Shared Sub Ekle("
        For Each k As Kolon In kolonlar
            If k.KolonAd <> Tools.IsDeleted.KolonAd Then
                sonuc &= "ByVal " & k.KolonAd & " as " & k.KolonTip & ","
            End If
        Next
        sonuc = sonuc.Substring(0, sonuc.Length - 1)
        sonuc &= ")" & vbCrLf
        If Connection.StartsWith("Tools.ConStr") Then
            sonuc &= vbTab & "Dim Con As New SqlConnection(" & Connection & ")" & vbCrLf
        Else
            sonuc &= vbTab & "Dim Con As New SqlConnection(" & Chr(34) & Connection & Chr(34) & ")" & vbCrLf
        End If

        If SpDahilMi Then
            sonuc &= vbTab & "Dim Com As new SqlCommand(" & Chr(34) & spAdi & Chr(34) & ",con)" & vbCrLf
            sonuc &= "Com.CommandType = CommandType.StoredProcedure" & vbCrLf
        Else
            sonuc &= vbTab & "Dim Com As new SqlCommand(" & Chr(34) & "Insert into " & TabloAdi & " ("
            For Each k As Kolon In kolonlar
                sonuc &= k.KolonAd & ","
            Next
            sonuc = sonuc.Substring(0, sonuc.Length - 1)
            sonuc &= ") VALUES ("
            For Each k As Kolon In kolonlar
                sonuc &= "@" & k.KolonAd & ","
            Next
            sonuc = sonuc.Substring(0, sonuc.Length - 1)
            sonuc &= ")" & Chr(34) & ",con)" & vbCrLf
        End If

        sonuc &= vbTab & "With Com.Parameters" & vbCrLf
        For Each k As Kolon In kolonlar
            If k.KolonAd <> Tools.IsDeleted.KolonAd Then
                sonuc &= vbTab & vbTab & ".AddWithValue(" & Chr(34) & "@" & k.KolonAd & Chr(34) & "," & k.KolonAd & ")" & vbCrLf
            End If
        Next
        sonuc &= vbTab & "End With" & vbCrLf

        sonuc &= vbTab & "Con.Open()" & vbCrLf
        sonuc &= vbTab & "Com.ExecuteNonQuery" & vbCrLf
        sonuc &= vbTab & "Con.Close()" & vbCrLf
        sonuc &= "End Sub" & vbCrLf

        Return sonuc
    End Function

    Public Shared Function Sil(ByVal TabloAdi As String, ByVal Connection As String, ByVal spAdi As String) As String
        Dim sonuc As String
        Tools.PrimaryKey.KolonTip = Tools.DegiskenCevir(Tools.PrimaryKey.KolonTip)
        sonuc = "Public Shared Sub DeleteRecord(ByVal " & Tools.PrimaryKey.KolonAd & " as " & Tools.PrimaryKey.KolonTip & ", ByVal kuladi As String)" & vbCrLf
        If Connection.StartsWith("Tools.ConStr") Then
            sonuc &= vbTab & "Dim Con As New SqlConnection(" & Connection & ")" & vbCrLf
        Else
            sonuc &= vbTab & "Dim Con As New SqlConnection(" & Chr(34) & Connection & Chr(34) & ")" & vbCrLf
        End If

        If SpDahilMi Then
            sonuc &= vbTab & "Dim Com As new SqlCommand(" & Chr(34) & spAdi & Chr(34) & ",con)" & vbCrLf
            sonuc &= "Com.CommandType = CommandType.StoredProcedure" & vbCrLf
        Else
            sonuc &= vbTab & "Dim Com As new SqlCommand(" & Chr(34) & "Delete From " & TabloAdi & " where " & Tools.PrimaryKey.KolonAd & "=@" & Tools.PrimaryKey.KolonAd & Chr(34) & ",con)" & vbCrLf
        End If

        sonuc &= vbTab & "Com.Parameters.AddWithValue(" & Chr(34) & "@" & Tools.PrimaryKey.KolonAd & Chr(34) & ", " & Tools.PrimaryKey.KolonAd & ")" & vbCrLf
        sonuc &= vbTab & "Com.Parameters.AddWithValue(" & Chr(34) & "@kul_adi" & Chr(34) & ",kuladi)" & vbCrLf
        sonuc &= vbTab & "Con.Open()" & vbCrLf
        sonuc &= vbTab & "Com.ExecuteNonQuery" & vbCrLf
        sonuc &= vbTab & "Con.Close()" & vbCrLf
        sonuc &= "End Sub" & vbCrLf

        Return sonuc
    End Function

    Public Shared Function Guncelle(ByVal kolonlar As KolonCollection, ByVal TabloAdi As String, ByVal Connection As String, ByVal spAdi As String) As String
        Dim sonuc As String
        sonuc = "Public Shared Sub Guncelle("
        'sonuc &= "ByVal " & PrimaryKey.KolonAd & " as " & PrimaryKey.KolonTip & ","
        For Each k As Kolon In kolonlar
            '   If PrimaryKey.KolonAd <> k.KolonAd Then
            sonuc &= "ByVal " & k.KolonAd & " as " & k.KolonTip & ","
            'End If
        Next
        sonuc = sonuc.Substring(0, sonuc.Length - 1)
        sonuc &= ")" & vbCrLf
        If Connection.StartsWith("Tools.ConStr") Then
            sonuc &= vbTab & "Dim Con As New SqlConnection(" & Connection & ")" & vbCrLf
        Else
            sonuc &= vbTab & "Dim Con As New SqlConnection(" & Chr(34) & Connection & Chr(34) & ")" & vbCrLf
        End If

        If SpDahilMi Then
            sonuc &= vbTab & "Dim Com As new SqlCommand(" & Chr(34) & spAdi & Chr(34) & ",con)" & vbCrLf
            sonuc &= "Com.CommandType = CommandType.StoredProcedure" & vbCrLf
        Else
            sonuc &= vbTab & "Dim Com As new SqlCommand(" & Chr(34) & "Update [" & TabloAdi & "] Set "

            For Each k As Kolon In kolonlar
                If PrimaryKey.KolonAd <> k.KolonAd Then
                    sonuc &= k.KolonAd & "=@" & k.KolonAd & ","
                End If
            Next
            sonuc = sonuc.Substring(0, sonuc.Length - 1)
            sonuc &= " where " & PrimaryKey.KolonAd & "=@" & PrimaryKey.KolonAd & Chr(34) & ",con)" & vbCrLf
        End If
        
        sonuc &= vbTab & "With Com.Parameters" & vbCrLf
        'sonuc &= vbTab & vbTab & ".AddWithValue(" & Chr(34) & "@" & PrimaryKey.KolonAd & Chr(34) & "," & PrimaryKey.KolonAd & ")" & vbCrLf

        For Each k As Kolon In kolonlar
            sonuc &= vbTab & vbTab & ".AddWithValue(" & Chr(34) & "@" & k.KolonAd & Chr(34) & "," & k.KolonAd & ")" & vbCrLf
        Next
        sonuc &= vbTab & "End With" & vbCrLf

        sonuc &= vbTab & "Con.Open()" & vbCrLf
        sonuc &= vbTab & "Com.ExecuteNonQuery" & vbCrLf
        sonuc &= vbTab & "Con.Close()" & vbCrLf
        sonuc &= "End Sub" & vbCrLf

        Return sonuc
    End Function

    Public Shared Function Get_ID(ByVal TabloAdi As String, ByVal Connection As String, ByVal kolonlar As KolonCollection, ByVal spAdi As String) As String
        Dim sonuc As String
        Dim degiskenAdi As String = ClassAdi.Substring(0, 2).ToLower
        sonuc = "Public Shared Function Get" & ClassAdi & "ByID(ByVal " & PrimaryKey.KolonAd & " as " & PrimaryKey.KolonTip & ") as " & ClassAdi & "_DAL" & vbCrLf
        sonuc &= vbTab & "Dim " & degiskenAdi & " as " & ClassAdi & "_DAL" & vbCrLf
        If Connection.StartsWith("Tools.ConStr") Then
            sonuc &= vbTab & "Dim Con As New SqlConnection(" & Connection & ")" & vbCrLf
        Else
            sonuc &= vbTab & "Dim Con As New SqlConnection(" & Chr(34) & Connection & Chr(34) & ")" & vbCrLf
        End If

        If SpDahilMi Then
            sonuc &= vbTab & "Dim Com As new SqlCommand(" & Chr(34) & spAdi & Chr(34) & ",con)" & vbCrLf
            sonuc &= "Com.CommandType = CommandType.StoredProcedure" & vbCrLf
        Else
            sonuc &= vbTab & "Dim Com As new SqlCommand(" & Chr(34) & "Select * from " & TabloAdi & " where " & PrimaryKey.KolonAd & "=@" & PrimaryKey.KolonAd & Chr(34) & ",con)" & vbCrLf
        End If

        sonuc &= vbTab & "Dim Dr as SqlDataReader" & vbCrLf
        sonuc &= vbTab & "Com.Parameters.AddWithValue(" & Chr(34) & "@" & PrimaryKey.KolonAd & Chr(34) & "," & PrimaryKey.KolonAd & ")" & vbCrLf
        sonuc &= vbTab & "Con.Open()" & vbCrLf
        sonuc &= vbTab & "Dr = Com.ExecuteReader" & vbCrLf
        sonuc &= vbTab & "While Dr.Read()" & vbCrLf
        sonuc &= vbTab & vbTab & degiskenAdi & "= New " & ClassAdi & "_DAL" & vbCrLf
        sonuc &= vbTab & vbTab & "With " & degiskenAdi & vbCrLf

        sonuc &= vbTab & vbTab & vbTab & "." & PrimaryKey.KolonAd & "=IIF(IsDBNull(dr(" & Chr(34) & PrimaryKey.KolonAd & Chr(34) & ")),0,dr(" & Chr(34) & PrimaryKey.KolonAd & Chr(34) & "))" & vbCrLf
        For Each kol As Kolon In kolonlar
            If kol.KolonTip.ToLower = "string" Then
                sonuc &= vbTab & vbTab & vbTab & "." & kol.KolonAd & "=IIF(IsDBNull(dr(" & Chr(34) & kol.KolonAd & Chr(34) & "))," & Chr(34) & Chr(34) & ",dr(" & Chr(34) & kol.KolonAd & Chr(34) & "))" & vbCrLf

            ElseIf kol.KolonTip.ToLower = "integer" Or kol.KolonTip.ToLower = "decimal" Then
                sonuc &= vbTab & vbTab & vbTab & "." & kol.KolonAd & "=IIF(IsDBNull(dr(" & Chr(34) & kol.KolonAd & Chr(34) & ")), 0,dr(" & Chr(34) & kol.KolonAd & Chr(34) & "))" & vbCrLf

            ElseIf kol.KolonTip.ToLower = "datetime" Then
                sonuc &= vbTab & vbTab & vbTab & "." & kol.KolonAd & "=IIF(IsDBNull(dr(" & Chr(34) & kol.KolonAd & Chr(34) & ")), #1/1/1900#,dr(" & Chr(34) & kol.KolonAd & Chr(34) & "))" & vbCrLf

            ElseIf kol.KolonTip.ToLower = "boolean" Then
                sonuc &= vbTab & vbTab & vbTab & "." & kol.KolonAd & "=IIF(IsDBNull(dr(" & Chr(34) & kol.KolonAd & Chr(34) & ")), false,dr(" & Chr(34) & kol.KolonAd & Chr(34) & "))" & vbCrLf
            End If
        Next
        sonuc &= vbTab & vbTab & "End With" & vbCrLf
        sonuc &= vbTab & "End While" & vbCrLf
        sonuc &= vbTab & "Con.Close()" & vbCrLf
        sonuc &= vbTab & "Return " & degiskenAdi & vbCrLf
        sonuc &= "End Function"
        Return sonuc
    End Function

    Public Shared Function Get_All(ByVal TabloAdi As String, ByVal Connection As String, ByVal kolonlar As KolonCollection, ByVal spAdi As String) As String
        Dim sonuc As String
        sonuc = "Public Shared Function GetAll" & ClassAdi & "() as " & ClassAdi & "_DAL" & "()" & vbCrLf
        sonuc &= vbTab & "Dim al as New ArrayList" & vbCrLf
        If Connection.StartsWith("Tools.ConStr") Then
            sonuc &= vbTab & "Dim Con As New SqlConnection(" & Connection & ")" & vbCrLf
        Else
            sonuc &= vbTab & "Dim Con As New SqlConnection(" & Chr(34) & Connection & Chr(34) & ")" & vbCrLf
        End If

        If SpDahilMi Then
            sonuc &= vbTab & "Dim Com As new SqlCommand(" & Chr(34) & spAdi & Chr(34) & ",con)" & vbCrLf
            sonuc &= "Com.CommandType = CommandType.StoredProcedure" & vbCrLf
        Else
            sonuc &= vbTab & "Dim Com As new SqlCommand(" & Chr(34) & "Select * from [" & TabloAdi & "]" & Chr(34) & ",con)" & vbCrLf
        End If
        sonuc &= vbTab & "Dim Dr as SqlDataReader" & vbCrLf
        sonuc &= vbTab & "Con.Open()" & vbCrLf
        sonuc &= vbTab & "Dr = Com.ExecuteReader" & vbCrLf
        sonuc &= vbTab & "While Dr.Read()" & vbCrLf
        sonuc &= vbTab & vbTab & "Dim " & ClassAdi.Substring(0, 1).ToLower & " as New " & ClassAdi & "_DAL" & vbCrLf
        sonuc &= vbTab & vbTab & "With " & ClassAdi.Substring(0, 1).ToLower & vbCrLf
        sonuc &= vbTab & vbTab & vbTab & ClassAdi.Substring(0, 1).ToLower & "." & PrimaryKey.KolonAd & "=IIF(IsDBNull(dr(" & Chr(34) & PrimaryKey.KolonAd & Chr(34) & ")),0,dr(" & Chr(34) & PrimaryKey.KolonAd & Chr(34) & "))" & vbCrLf
        For Each kol As Kolon In kolonlar
            If kol.KolonTip.ToLower = "string" Then
                sonuc &= vbTab & vbTab & vbTab & ClassAdi.Substring(0, 1).ToLower & "." & kol.KolonAd & "=IIF(IsDBNull(dr(" & Chr(34) & kol.KolonAd & Chr(34) & "))," & Chr(34) & Chr(34) & ",dr(" & Chr(34) & kol.KolonAd & Chr(34) & "))" & vbCrLf

            ElseIf kol.KolonTip.ToLower = "integer" Or kol.KolonTip.ToLower = "decimal" Then
                sonuc &= vbTab & vbTab & vbTab & ClassAdi.Substring(0, 1).ToLower & "." & kol.KolonAd & "=IIF(IsDBNull(dr(" & Chr(34) & kol.KolonAd & Chr(34) & ")), 0,dr(" & Chr(34) & kol.KolonAd & Chr(34) & "))" & vbCrLf

            ElseIf kol.KolonTip.ToLower = "datetime" Then
                sonuc &= vbTab & vbTab & vbTab & ClassAdi.Substring(0, 1).ToLower & "." & kol.KolonAd & "=IIF(IsDBNull(dr(" & Chr(34) & kol.KolonAd & Chr(34) & ")), #1/1/1900#,dr(" & Chr(34) & kol.KolonAd & Chr(34) & "))" & vbCrLf

            ElseIf kol.KolonTip.ToLower = "boolean" Then
                sonuc &= vbTab & vbTab & vbTab & ClassAdi.Substring(0, 1).ToLower & "." & kol.KolonAd & "=IIF(IsDBNull(dr(" & Chr(34) & kol.KolonAd & Chr(34) & ")), false,dr(" & Chr(34) & kol.KolonAd & Chr(34) & "))" & vbCrLf
            End If
        Next
        sonuc &= vbTab & vbTab & "End With" & vbCrLf
        sonuc &= vbTab & vbTab & "Al.Add(" & ClassAdi.Substring(0, 1).ToLower & ")" & vbCrLf
        sonuc &= vbTab & "End While" & vbCrLf
        sonuc &= vbTab & "Con.Close()" & vbCrLf
        sonuc &= vbTab & "Return al.ToArray(GetType(" & ClassAdi & "_DAL" & "))" & vbCrLf
        sonuc &= "End Function"

        Return sonuc
    End Function

    Public Shared Function Get_All_ds(ByVal TabloAdi As String, ByVal Connection As String, ByVal kolonlar As KolonCollection, ByVal spAdi As String) As String
        Dim sonuc As String
        sonuc = "Public Shared Function GetAllDs" & ClassAdi & "() as Dataset" & vbCrLf

        If Connection.StartsWith("Tools.ConStr") Then
            sonuc &= vbTab & "Dim Con As New SqlConnection(" & Connection & ")" & vbCrLf
        Else
            sonuc &= vbTab & "Dim Con As New SqlConnection(" & Chr(34) & Connection & Chr(34) & ")" & vbCrLf
        End If

        If SpDahilMi Then
            sonuc &= vbTab & "Dim Com As new SqlCommand(" & Chr(34) & spAdi & Chr(34) & ",con)" & vbCrLf
            sonuc &= "Com.CommandType = CommandType.StoredProcedure" & vbCrLf
        Else
            sonuc &= vbTab & "Dim Com As new SqlCommand(" & Chr(34) & "Select * from [" & TabloAdi & "]" & Chr(34) & ",con)" & vbCrLf
        End If
        sonuc &= vbTab & "Dim Da as New SqlDataAdapter" & vbCrLf
        sonuc &= vbTab & "Da.SelectCommand=Com" & vbCrLf
        sonuc &= vbTab & "Dim Ds as New DataSet" & vbCrLf
        sonuc &= vbTab & "Da.Fill(Ds)" & vbCrLf
        sonuc &= vbTab & "Return Ds" & vbCrLf
        
        sonuc &= "End Function"

        Return sonuc
    End Function


    Public Shared Function ClassOl(ByVal Kolonlar As KolonCollection, ByVal TabloAdi As String, ByVal ConnectionTuru As Connect) As String
        If PrimaryKey Is Nothing Then Return ""
        If IsDeleted Is Nothing Then Return ""
        If Tools.ClassAdi = "" Then Return ""

        Select Case ConnectionTuru.Tur
            Case "Trusted"
                Tools.ConStr = "server=" & ConnectionTuru.Server & ";Database=" & ConnectionTuru.DatabaseAd & ";Trusted_Connection=True;"
            Case "Tools"
                Tools.ConStr = "Tools.ConStr(" & Chr(34) & "eczMonF18_2013" & Chr(34) & ")"
            Case "User"
                Tools.ConStr = "server=" & ConnectionTuru.Server & ";Database=" & ConnectionTuru.DatabaseAd & ";User ID=" & ConnectionTuru.UserID & ";Password=" & ConnectionTuru.Password & ";"
        End Select

        Dim kc As New KolonCollection
        For Each kln As Kolon In Kolonlar
            kc.Add(kln)
        Next

        Dim sonuc As String = vbCrLf & "imports System.Data.SqlClient" & vbCrLf & vbCrLf

        sonuc &= "Public Class " & ClassAdi & "_DAL" & vbCrLf & vbCrLf

        'Deðiþkenleri oluþtur
        For Each k As Kolon In Kolonlar
            sonuc &= Tools.Degisken_Olustur(k.KolonAd, k.KolonTip) & vbCrLf
        Next
        'Deðiþkenler Oluþturuldu !!!
        sonuc &= vbCrLf

        sonuc &= "#Region " & Chr(34) & "Properties" & Chr(34) & vbCrLf

        'Property 'leri Oluþtur
        For Each k As Kolon In Kolonlar
            sonuc &= Tools.Property_Olustur(k.KolonAd, k.KolonTip) & vbCrLf & vbCrLf
        Next
        'Property'ler Oluþturuldu !!!

        sonuc &= vbCrLf & "#End Region " & vbCrLf & vbCrLf
        sonuc &= "#Region " & Chr(34) & "Methods" & Chr(34) & vbCrLf

        'Sil Methodunu Oluþtur
        sonuc &= Tools.Sil(TabloAdi, Tools.ConStr, "sp_" & Tools.ClassAdi & "Sil") & vbCrLf & vbCrLf
        'Sil Methodu Oluþturuldu !!!
        sonuc &= vbCrLf & vbCrLf
        'Ekle Methodunu Oluþtur
        sonuc &= Tools.Ekle(TabloAdi, kc, Tools.ConStr, "sp_" & Tools.ClassAdi & "Ekle")
        'Ekle Methodu Oluþturuldu
        sonuc &= vbCrLf & vbCrLf
        'Güncelle Methodunu Oluþtur
        sonuc &= Tools.Guncelle(kc, TabloAdi, Tools.ConStr, "sp_" & Tools.ClassAdi & "Guncelle")
        'Güncelle Methodu Oluþturuldu
        sonuc &= vbCrLf & vbCrLf

        sonuc &= Tools.Kaydet(kc, TabloAdi, Tools.ConStr, "sp_" & Tools.ClassAdi & "Kaydet")
        'Kaydet Methodu Oluþturuldu
        sonuc &= vbCrLf & vbCrLf

        'GetByID methodu oluþtur
        sonuc &= Tools.Get_ID(TabloAdi, Tools.ConStr, kc, "sp_" & Tools.ClassAdi & "GetByID")
        'GetByID methodu oluþturuldu
        sonuc &= vbCrLf & vbCrLf
        'GetAllByID methodu oluþtur
        sonuc &= Tools.Get_All(TabloAdi, Tools.ConStr, kc, "sp_" & Tools.ClassAdi & "GetAll")
        'GetAll methodu oluþturuldu
        sonuc &= vbCrLf & vbCrLf
        sonuc &= Tools.Get_All_ds(TabloAdi, Tools.ConStr, kc, "sp_" & Tools.ClassAdi & "GetAll")
        'GetAll methodu oluþturuldu
        sonuc &= vbCrLf & "#End Region " & vbCrLf

        If TotoString <> "" Then
            sonuc &= vbCrLf & vbTab & "Public Overrides Function toString() As String" & vbCrLf
            sonuc &= vbTab & vbTab & "Return m" & TotoString & vbCrLf
            sonuc &= vbTab & "End Function" & vbCrLf
        End If

        sonuc &= vbCrLf & "End Class" & vbCrLf & vbCrLf
        Return sonuc
    End Function


    Public Shared Function DegiskenCevir(ByVal DegiskenAdi As String) As String
        Select Case DegiskenAdi
            Case "int", "smallint", "bigint", "tinyint"
                DegiskenAdi = "integer"
            Case "char", "varchar", "nvarchar", "text", "ntext", "nchar"
                DegiskenAdi = "string"
            Case "money", "float"
                DegiskenAdi = "decimal"
            Case "bit"
                DegiskenAdi = "boolean"
        End Select
        Return DegiskenAdi
    End Function

    Public Shared Function CollectionOlustur() As String
        Dim sonuc As String
        sonuc = vbCrLf & "Public Class " & ClassAdi & "Collection" & vbCrLf
        sonuc &= vbTab & "Inherits Collectionbase" & vbCrLf & vbCrLf
        sonuc &= vbTab & "Public Sub Add(ByVal " & ClassAdi.Substring(0, 1).ToLower & " As " & ClassAdi & "_DAL)" & vbCrLf
        sonuc &= vbTab & vbTab & "MyBase.List.Add(" & ClassAdi.Substring(0, 1).ToLower & ")" & vbCrLf
        sonuc &= vbTab & "End Sub" & vbCrLf
        sonuc &= vbTab & "Default Public Property Saglayici(ByVal index As Integer) As " & ClassAdi & "_DAL" & vbCrLf
        sonuc &= vbTab & vbTab & "Get" & vbCrLf
        sonuc &= vbTab & vbTab & vbTab & "Return MyBase.InnerList(index)" & vbCrLf
        sonuc &= vbTab & vbTab & "End Get" & vbCrLf
        sonuc &= vbTab & vbTab & "Set(ByVal Value As " & ClassAdi & "_DAL)" & vbCrLf
        sonuc &= vbTab & vbTab & vbTab & "MyBase.InnerList(index) = Value" & vbCrLf
        sonuc &= vbTab & vbTab & "End Set" & vbCrLf
        sonuc &= vbTab & "End Property" & vbCrLf
        sonuc &= "End Class" & vbCrLf & vbCrLf
        Return sonuc
    End Function

    Public Shared Function Kaydet(ByVal kolonlar As KolonCollection, ByVal TabloAdi As String, ByVal Connection As String, ByVal spAdi As String) As String
        Dim sonuc As String
        sonuc = "Public Shared Sub Kaydet("
        'sonuc &= "ByVal " & PrimaryKey.KolonAd & " as " & PrimaryKey.KolonTip & ","
        For Each k As Kolon In kolonlar
            '   If PrimaryKey.KolonAd <> k.KolonAd Then
            If Tools.IsDeleted.KolonAd <> k.KolonAd Then
                sonuc &= "ByVal " & k.KolonAd & " as " & k.KolonTip & ","
            End If
            'End If
        Next
        'sonuc = sonuc.Substring(0, sonuc.Length - 1)
        sonuc &= "ByVal kuladi as String"
        sonuc &= ")" & vbCrLf
        If Connection.StartsWith("Tools.ConStr") Then
            sonuc &= vbTab & "Dim Con As New SqlConnection(" & Connection & ")" & vbCrLf
        Else
            sonuc &= vbTab & "Dim Con As New SqlConnection(" & Chr(34) & Connection & Chr(34) & ")" & vbCrLf
        End If

        If SpDahilMi Then
            sonuc &= vbTab & "Dim Com As new SqlCommand(" & Chr(34) & spAdi & Chr(34) & ",con)" & vbCrLf
            sonuc &= "Com.CommandType = CommandType.StoredProcedure" & vbCrLf
        Else
            sonuc &= vbTab & "Dim Com As new SqlCommand(" & Chr(34) & "Update [" & TabloAdi & "] Set "

            For Each k As Kolon In kolonlar
                If PrimaryKey.KolonAd <> k.KolonAd Then
                    sonuc &= k.KolonAd & "=@" & k.KolonAd & ","
                End If
            Next
            sonuc = sonuc.Substring(0, sonuc.Length - 1)
            sonuc &= " where " & PrimaryKey.KolonAd & "=@" & PrimaryKey.KolonAd & Chr(34) & ",con)" & vbCrLf
        End If

        sonuc &= vbTab & "With Com.Parameters" & vbCrLf
        'sonuc &= vbTab & vbTab & ".AddWithValue(" & Chr(34) & "@" & PrimaryKey.KolonAd & Chr(34) & "," & PrimaryKey.KolonAd & ")" & vbCrLf

        For Each k As Kolon In kolonlar
            sonuc &= vbTab & vbTab & ".AddWithValue(" & Chr(34) & "@" & k.KolonAd & Chr(34) & "," & k.KolonAd & ")" & vbCrLf
        Next
        sonuc &= vbTab & vbTab & ".AddWithValue(" & Chr(34) & "@kul_adi" & Chr(34) & ", kuladi)" & vbCrLf
        sonuc &= vbTab & "End With" & vbCrLf

        sonuc &= vbTab & "Con.Open()" & vbCrLf
        sonuc &= vbTab & "Com.ExecuteNonQuery" & vbCrLf
        sonuc &= vbTab & "Con.Close()" & vbCrLf
        sonuc &= "End Sub" & vbCrLf

        Return sonuc
    End Function


End Class












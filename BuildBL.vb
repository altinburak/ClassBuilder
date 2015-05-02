Public Class BuildBL
    Public Shared Function BLOlustur(ByVal Kolonlar As KolonCollection, ByVal TabloAdi As String) As String
        Dim sonuc As String = "imports f19_DAL." & TabloAdi & "_DAL" & vbCrLf & vbCrLf
        sonuc &= "Public Class " & TabloAdi & "_BL" & vbCrLf & vbCrLf

        sonuc &= BuildEkle(Kolonlar, TabloAdi)
        sonuc &= BuildSil(TabloAdi)
        sonuc &= BuildGuncelle(Kolonlar, TabloAdi)
        sonuc &= BuildKaydet(Kolonlar, TabloAdi)
        sonuc &= BuildGetById(TabloAdi)
        sonuc &= GetAllDizi(TabloAdi)
        sonuc &= GetAllDs(TabloAdi)
        'sonuc &= GetAllDt(TabloAdi)
        sonuc &= BuildGuncelleDs(TabloAdi)
        sonuc &= BuildGuncelleDt(TabloAdi)

        sonuc &= vbCrLf & "End Class"
        Return sonuc
    End Function

    Public Shared Function BuildEkle(ByVal Kolonlar As KolonCollection, ByVal TabloAdi As String) As String
        Dim sonuc As String = ""
        Dim EkleKolonlar As String = ""
        sonuc &= "Public Shared Sub Ekle("
        For Each kol As Kolon In Kolonlar
            If kol.KolonAd <> Tools.IsDeleted.KolonAd Then
                sonuc &= "ByVal " & kol.KolonAd & " As " & kol.KolonTip & ", "
                EkleKolonlar &= kol.KolonAd & ", "
            End If
        Next
        sonuc = sonuc.Substring(0, sonuc.Length - 2) & ")" & vbCrLf
        EkleKolonlar = EkleKolonlar.Substring(0, EkleKolonlar.Length - 2) & vbCrLf
        sonuc &= "f19_DAL." & TabloAdi & "_DAL.Ekle(" & EkleKolonlar & ")" & vbCrLf
        sonuc &= "End Sub" & vbCrLf

        Return sonuc
    End Function

    Public Shared Function BuildSil(ByVal TabloAdi As String) As String
        Dim sonuc As String = ""
        Dim EkleKolonlar As String = ""
        sonuc &= "Public Shared Sub Sil("
        sonuc &= "ByVal " & Tools.PrimaryKey.KolonAd & " As " & Tools.PrimaryKey.KolonTip & ", ByVal kuladi As String)" & vbCrLf
        sonuc &= "f19_DAL." & TabloAdi & "_DAL.DeleteRecord(" & Tools.PrimaryKey.KolonAd & ", kuladi)" & vbCrLf
        sonuc &= "End Sub" & vbCrLf

        Return sonuc
    End Function

    Public Shared Function BuildGuncelle(ByVal Kolonlar As KolonCollection, ByVal TabloAdi As String) As String
        Dim sonuc As String = ""
        Dim EkleKolonlar As String = ""
        sonuc &= "Public Shared Sub Guncelle("
        For Each kol As Kolon In Kolonlar
            sonuc &= "ByVal " & kol.KolonAd & " As " & kol.KolonTip & ", "
            'If kol.KolonAd <> Tools.IsDeleted.KolonAd Then
            EkleKolonlar &= kol.KolonAd & ", "
            'End If
        Next
        sonuc = sonuc.Substring(0, sonuc.Length - 2) & ")" & vbCrLf
        EkleKolonlar = EkleKolonlar.Substring(0, EkleKolonlar.Length - 2) & vbCrLf
        sonuc &= "f19_DAL." & TabloAdi & "_DAL.Guncelle(" & EkleKolonlar & ")" & vbCrLf
        sonuc &= "End Sub" & vbCrLf

        Return sonuc
    End Function

    Public Shared Function BuildGetById(ByVal TabloAdi As String) As String
        Dim sonuc As String = ""
        Dim EkleKolonlar As String = ""
        sonuc &= "Public Shared Function Get" & TabloAdi & "ByID("
        sonuc &= "ByVal " & Tools.PrimaryKey.KolonAd & " As " & Tools.PrimaryKey.KolonTip & ") As F19_DAL." & TabloAdi & "_DAL" & vbCrLf
        sonuc &= "Return f19_DAL." & TabloAdi & "_DAL.Get" & TabloAdi & "ByID(" & Tools.PrimaryKey.KolonAd & ")" & vbCrLf
        sonuc &= "End Function" & vbCrLf

        Return sonuc
    End Function

    Public Shared Function GetAllDizi(ByVal TabloAdi As String) As String
        Dim sonuc As String = ""
        Dim EkleKolonlar As String = ""
        sonuc &= "Public Shared Function GetAll" & TabloAdi & "()"
        sonuc &= " As F19_DAL." & TabloAdi & "_DAL()" & vbCrLf
        sonuc &= "Return F19_DAL." & TabloAdi & "_DAL.GetAll" & TabloAdi & "()" & vbCrLf
        sonuc &= "End Function" & vbCrLf

        Return sonuc
    End Function

    Public Shared Function GetAllDs(ByVal TabloAdi As String) As String
        Dim sonuc As String = ""
        Dim EkleKolonlar As String = ""
        sonuc &= "Public Shared Function GetAll" & TabloAdi & "Ds()" & " As Dataset" & vbCrLf
        sonuc &= "Return F19_DAL." & TabloAdi & "_DAL.GetAllDs" & TabloAdi & "()" & vbCrLf
        sonuc &= "End Function" & vbCrLf

        Return sonuc
    End Function

    'Public Shared Function GetAllDt(ByVal TabloAdi As String) As String
    '    Dim sonuc As String = ""
    '    Dim EkleKolonlar As String = ""
    '    sonuc &= "Public Shared Function GetAll" & TabloAdi & "Dt()" & " As DataTable" & vbCrLf
    '    sonuc &= "Return F19_DAL." & TabloAdi & "_DAL.GetAll" & TabloAdi & "Dt()" & vbCrLf
    '    sonuc &= "End Function" & vbCrLf

    '    Return sonuc
    'End Function

    Public Shared Function BuildGuncelleDs(ByVal TabloAdi As String) As String
        Dim sonuc As String = ""
        Dim EkleKolonlar As String = ""
        sonuc &= "Public Shared Sub Update" & TabloAdi & "Ds(ByVal Ds As Dataset)" & vbCrLf
        sonuc &= "Update" & TabloAdi & "Ds(Ds)" & vbCrLf
        sonuc &= "End Sub" & vbCrLf

        Return sonuc
    End Function

    Public Shared Function BuildGuncelleDt(ByVal TabloAdi As String) As String
        Dim sonuc As String = ""
        Dim EkleKolonlar As String = ""
        sonuc &= "Public Shared Sub Update" & TabloAdi & "Dt(ByVal Dt As DataTable)" & vbCrLf
        sonuc &= "Update" & TabloAdi & "Dt(Dt)" & vbCrLf
        sonuc &= "End Sub" & vbCrLf

        Return sonuc
    End Function

    Public Shared Function BuildKaydet(ByVal Kolonlar As KolonCollection, ByVal TabloAdi As String) As String
        Dim sonuc As String = ""
        Dim EkleKolonlar As String = ""
        sonuc &= "Public Shared Sub Kaydet("
        For Each kol As Kolon In Kolonlar
            If kol.KolonAd <> Tools.IsDeleted.KolonAd Then
                sonuc &= "ByVal " & kol.KolonAd & " As " & kol.KolonTip & ", "
                EkleKolonlar &= kol.KolonAd & ", "
            End If
        Next
        'sonuc = sonuc.Substring(0, sonuc.Length - 2) & ")" & vbCrLf
        sonuc &= "ByVal kuladi as String)" & vbCrLf
        'EkleKolonlar = EkleKolonlar.Substring(0, EkleKolonlar.Length - 2) & vbCrLf
        EkleKolonlar &= "kuladi"
        sonuc &= "f19_DAL." & TabloAdi & "_DAL.Kaydet(" & EkleKolonlar & ")" & vbCrLf
        sonuc &= "End Sub" & vbCrLf

        Return sonuc
    End Function
End Class

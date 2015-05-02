Public Class BuildSP

    Public Shared Function SpOlustur(ByVal Kolonlar As KolonCollection, ByVal TabloAdi As String, ByVal ConnectionTuru As Connect) As String
        Dim sonuc As String = ""
        sonuc = BuildHistTable(TabloAdi, Kolonlar)
        Dim spNames() As String = {"sp_" & TabloAdi & "Ekle", "sp_" & TabloAdi & "Sil", "sp_" & TabloAdi & "Guncelle", "sp_" & TabloAdi & "GetAll", "sp_" & TabloAdi & "GetByID", "sp_" & TabloAdi & "Kaydet"}
        For i As Integer = 0 To 5
            sonuc &= BuildExists(spNames(i))
        Next
        sonuc &= BuildEkle(TabloAdi, Kolonlar)
        sonuc &= BuildGuncelle(TabloAdi, Kolonlar)
        sonuc &= BuildSil(TabloAdi, Kolonlar)
        sonuc &= BuildHepsiniGetir(TabloAdi, Kolonlar)
        sonuc &= BuildGetir(TabloAdi)
        sonuc &= BuildKaydet(TabloAdi, Kolonlar)

        Return sonuc
    End Function
    Public Shared Function BuildExists(ByVal spName As String) As String
        Dim sonuc As String = ""
        sonuc &= "if exists (select * from dbo.sysobjects where name = '" & spName & "')" & vbCrLf
        sonuc &= "drop procedure " & spName
        sonuc &= vbCrLf & vbCrLf & "GO" & vbCrLf & vbCrLf

        Return sonuc
    End Function

    Public Shared Function BuildEkle(ByVal TabloAdi As String, ByVal kolonlar As KolonCollection) As String
        Dim sonuc As String = ""
        sonuc &= "CREATE PROCEDURE sp_" & Tools.ClassAdi & "Ekle" & vbCrLf
        Dim InsertKolonlar As String = ""
        Dim DegiskenKolonlar As String = ""
        For Each kol As Kolon In kolonlar
            If kol.KolonAd.EndsWith("is_deleted") = False Then
                If kol.KolonTip = "nvarchar" Then
                    sonuc &= "@" & kol.KolonAd & " " & kol.KolonTip & "(" & kol.KolonUzunluk & ")," & vbCrLf
                Else
                    sonuc &= "@" & kol.KolonAd & " " & kol.KolonTip & "," & vbCrLf
                End If
                InsertKolonlar &= kol.KolonAd & ", "
                DegiskenKolonlar &= "@" & kol.KolonAd & ", "
            End If
        Next
        sonuc = sonuc.Substring(0, sonuc.Length - 3) & vbCrLf
        InsertKolonlar = InsertKolonlar.Substring(0, InsertKolonlar.Length - 2)
        DegiskenKolonlar = DegiskenKolonlar.Substring(0, DegiskenKolonlar.Length - 2)

        sonuc &= "AS" & vbCrLf
        sonuc &= "INSERT " & TabloAdi & vbCrLf & "("
        sonuc &= InsertKolonlar
        sonuc &= ")" & vbCrLf & "Values" & vbCrLf & "("
        sonuc &= DegiskenKolonlar
        sonuc &= ")" & vbCrLf
        sonuc &= vbCrLf & vbCrLf & "GO" & vbCrLf & vbCrLf

        Return sonuc
    End Function

    Public Shared Function BuildGuncelle(ByVal TabloAdi As String, ByVal kolonlar As KolonCollection) As String
        Dim sonuc As String = ""
        sonuc &= "CREATE PROCEDURE sp_" & Tools.ClassAdi & "Guncelle" & vbCrLf
        Dim UpdateKolonlar As String = ""
        For Each kol As Kolon In kolonlar
            If kol.KolonAd.EndsWith("is_deleted") = False Then
                If kol.KolonAd <> Tools.PrimaryKey.KolonAd Then
                    UpdateKolonlar &= kol.KolonAd & " = @" & kol.KolonAd & ", "
                End If
                If kol.KolonTip = "nvarchar" Then
                    sonuc &= "@" & kol.KolonAd & " " & kol.KolonTip & "(" & kol.KolonUzunluk & ")," & vbCrLf
                Else
                    sonuc &= "@" & kol.KolonAd & " " & kol.KolonTip & "," & vbCrLf
                End If
            End If
        Next
        sonuc = sonuc.Substring(0, sonuc.Length - 3) & vbCrLf
        UpdateKolonlar = UpdateKolonlar.Substring(0, UpdateKolonlar.Length - 2)

        sonuc &= "AS" & vbCrLf
        sonuc &= "UPDATE " & TabloAdi & vbCrLf & " SET " & vbCrLf
        sonuc &= UpdateKolonlar & vbCrLf
        sonuc &= "WHERE " & Tools.PrimaryKey.KolonAd & " = @" & Tools.PrimaryKey.KolonAd
        sonuc &= vbCrLf & vbCrLf & "GO" & vbCrLf & vbCrLf

        Return sonuc
    End Function

    Public Shared Function BuildGetir(ByVal TabloAdi As String) As String
        Dim sonuc As String = ""
        sonuc &= "CREATE PROCEDURE sp_" & Tools.ClassAdi & "GetByID" & vbCrLf
        sonuc &= "(" & vbCrLf
        sonuc &= "@" & Tools.PrimaryKey.KolonAd & " " & Tools.PrimaryKey.KolonTip & vbCrLf
        sonuc &= ")" & vbCrLf
        sonuc &= "AS" & vbCrLf
        sonuc &= "SELECT * FROM " & TabloAdi & " WHERE " & Tools.PrimaryKey.KolonAd & "=@" & Tools.PrimaryKey.KolonAd
        sonuc &= " AND " & Tools.IsDeleted.KolonAd & " = 0 " & vbCrLf
        sonuc &= vbCrLf & vbCrLf & "GO" & vbCrLf & vbCrLf

        Return sonuc
    End Function

    'Public Shared Function BuildSil(ByVal TabloAdi As String) As String
    '    Dim sonuc As String = ""
    '    sonuc &= "CREATE PROCEDURE sp_" & Tools.ClassAdi & "Sil" & vbCrLf
    '    sonuc &= "(" & vbCrLf
    '    sonuc &= "@" & Tools.PrimaryKey.KolonAd & " " & Tools.PrimaryKey.KolonTip & vbCrLf
    '    sonuc &= ")" & vbCrLf
    '    sonuc &= "AS" & vbCrLf
    '    sonuc &= "DELETE FROM " & TabloAdi & " WHERE " & Tools.PrimaryKey.KolonAd & "=@" & Tools.PrimaryKey.KolonAd & vbCrLf
    '    sonuc &= vbCrLf & vbCrLf & "GO" & vbCrLf & vbCrLf

    '    Return sonuc
    'End Function

    Public Shared Function BuildSil(ByVal TabloAdi As String, ByVal kolonlar As KolonCollection) As String
        Dim sonuc As String = ""
        Dim InsertKolonlar As String = ""
        Dim InsertKolonlarHist As String = ""

        For Each kol As Kolon In kolonlar
            InsertKolonlar &= kol.KolonAd & ", "
            InsertKolonlarHist &= kol.KolonAd & ", "
        Next
        InsertKolonlarHist &= "username, create_date, trans_type"

        sonuc &= "CREATE PROCEDURE sp_" & Tools.ClassAdi & "Sil" & vbCrLf
        sonuc &= "(" & vbCrLf
        sonuc &= "@" & Tools.PrimaryKey.KolonAd & " " & Tools.PrimaryKey.KolonTip & "," & vbCrLf
        sonuc &= "@kul_adi nvarchar(50)" & vbCrLf
        sonuc &= ")" & vbCrLf
        sonuc &= "AS" & vbCrLf
        sonuc &= "BEGIN TRAN" & vbCrLf
        sonuc &= "INSERT INTO " & TabloAdi & "_hist (" & InsertKolonlarHist & ")" & vbCrLf
        sonuc &= "SELECT " & InsertKolonlar & "@kul_adi, GETDATE(), 'DEL' FROM " & TabloAdi & " WHERE " & Tools.PrimaryKey.KolonAd & "=@" & Tools.PrimaryKey.KolonAd & vbCrLf & vbCrLf

        sonuc &= "UPDATE " & TabloAdi & " SET " & Tools.IsDeleted.KolonAd & " = 1 WHERE " & Tools.PrimaryKey.KolonAd & "=@" & Tools.PrimaryKey.KolonAd & vbCrLf

        sonuc &= "IF @@Error > 0" & vbCrLf
        sonuc &= "ROLLBACK " & vbCrLf
        sonuc &= "ELSE " & vbCrLf
        sonuc &= "COMMIT TRAN " & vbCrLf
        sonuc &= vbCrLf & "GO" & vbCrLf & vbCrLf

        Return sonuc
    End Function

    Public Shared Function BuildHepsiniGetir(ByVal TabloAdi As String, ByVal kolonlar As KolonCollection) As String
        Dim sonuc As String = ""
        sonuc &= "CREATE PROCEDURE sp_" & Tools.ClassAdi & "GetAll" & vbCrLf
        sonuc &= "AS" & vbCrLf
        sonuc &= "SELECT * FROM " & TabloAdi & " WHERE " & Tools.IsDeleted.KolonAd & " = 0 " & vbCrLf
        sonuc &= vbCrLf & vbCrLf & "GO" & vbCrLf & vbCrLf

        Return sonuc
    End Function

    Public Shared Function BuildKaydet(ByVal TabloAdi As String, ByVal kolonlar As KolonCollection) As String
        Dim sonuc As String = ""
        sonuc &= "CREATE PROCEDURE sp_" & Tools.ClassAdi & "Kaydet" & vbCrLf
        Dim InsertKolonlar As String = ""
        Dim InsertKolonlarHist As String = ""
        Dim InsertKolonlarNoId As String = ""
        Dim DegiskenKolonlar As String = ""
        Dim DegiskenKolonlarNoId As String = ""
        Dim UpdateKolonlar As String = ""
        For Each kol As Kolon In kolonlar
            If kol.KolonAd.EndsWith("is_deleted") = False Then
                If kol.KolonTip = "nvarchar" Then
                    sonuc &= "@" & kol.KolonAd & " " & kol.KolonTip & "(" & kol.KolonUzunluk & ")," & vbCrLf
                Else
                    sonuc &= "@" & kol.KolonAd & " " & kol.KolonTip & "," & vbCrLf
                End If
                If kol.KolonAd <> Tools.PrimaryKey.KolonAd Then
                    InsertKolonlarNoId &= kol.KolonAd & ", "
                    DegiskenKolonlarNoId &= "@" & kol.KolonAd & ", "
                End If
                InsertKolonlar &= kol.KolonAd & ", "
                InsertKolonlarHist &= kol.KolonAd & ", "
                DegiskenKolonlar &= "@" & kol.KolonAd & ", "
                If kol.KolonAd <> Tools.PrimaryKey.KolonAd Then
                    UpdateKolonlar &= kol.KolonAd & " = @" & kol.KolonAd & ", "
                End If
            End If
        Next
        'sonuc = sonuc.Substring(0, sonuc.Length - 3) & vbCrLf
        sonuc &= "@kul_adi nvarchar(50)" & vbCrLf
        InsertKolonlar = InsertKolonlar.Substring(0, InsertKolonlar.Length - 2)
        InsertKolonlarHist &= "username, create_date, trans_type"
        InsertKolonlarNoId = InsertKolonlarNoId.Substring(0, InsertKolonlarNoId.Length - 2)
        DegiskenKolonlar = DegiskenKolonlar.Substring(0, DegiskenKolonlar.Length - 2)
        DegiskenKolonlarNoId = DegiskenKolonlarNoId.Substring(0, DegiskenKolonlarNoId.Length - 2)
        UpdateKolonlar = UpdateKolonlar.Substring(0, UpdateKolonlar.Length - 2)

        sonuc &= "AS" & vbCrLf
        sonuc &= "IF EXISTS(SELECT * FROM " & TabloAdi & " WHERE " & Tools.PrimaryKey.KolonAd & "=@" & Tools.PrimaryKey.KolonAd & ")" & vbCrLf
        sonuc &= "BEGIN" & vbCrLf
        sonuc &= "BEGIN TRAN" & vbCrLf
        sonuc &= "INSERT INTO " & TabloAdi & "_hist (" & InsertKolonlarHist & ")" & vbCrLf
        sonuc &= "SELECT " & InsertKolonlar & ", @kul_adi, GETDATE(), 'UPD' FROM " & TabloAdi & " WHERE " & Tools.PrimaryKey.KolonAd & "=@" & Tools.PrimaryKey.KolonAd & vbCrLf & vbCrLf

        sonuc &= "UPDATE " & TabloAdi & vbCrLf & " SET " & vbCrLf
        sonuc &= UpdateKolonlar & vbCrLf
        sonuc &= "WHERE " & Tools.PrimaryKey.KolonAd & " = @" & Tools.PrimaryKey.KolonAd & vbCrLf & vbCrLf

        sonuc &= "IF @@Error > 0" & vbCrLf
        sonuc &= "ROLLBACK " & vbCrLf
        sonuc &= "ELSE " & vbCrLf
        sonuc &= "COMMIT TRAN " & vbCrLf

        sonuc &= "END" & vbCrLf
        sonuc &= "ELSE" & vbCrLf
        sonuc &= "BEGIN" & vbCrLf
        sonuc &= "BEGIN TRAN " & vbCrLf

        sonuc &= "INSERT INTO " & TabloAdi & "_hist (" & InsertKolonlarHist & ")" & vbCrLf
        sonuc &= "SELECT " & InsertKolonlar & ", @kul_adi, GETDATE(), 'INS' FROM " & TabloAdi & " WHERE " & Tools.PrimaryKey.KolonAd & "=@" & Tools.PrimaryKey.KolonAd & vbCrLf & vbCrLf

        sonuc &= "INSERT " & TabloAdi & vbCrLf & "("
        sonuc &= InsertKolonlarNoId
        sonuc &= ")" & vbCrLf & "Values" & vbCrLf & "("
        sonuc &= DegiskenKolonlarNoId
        sonuc &= ")" & vbCrLf

        sonuc &= "IF @@Error > 0" & vbCrLf
        sonuc &= "ROLLBACK " & vbCrLf
        sonuc &= "ELSE " & vbCrLf
        sonuc &= "COMMIT TRAN " & vbCrLf

        sonuc &= "END" & vbCrLf & vbCrLf

        Return sonuc
    End Function

    Public Shared Function BuildHistTable(ByVal TabloAdi As String, ByVal kolonlar As KolonCollection) As String
        Dim sonuc As String = ""
        Dim HistTabloAdi As String = Tools.ClassAdi & "_hist"
        sonuc &= "CREATE TABLE [dbo].[" & HistTabloAdi & "](" & vbCrLf
        sonuc &= HistTabloAdi & "_id bigint IDENTITY(1,1) NOT NULL," & vbCrLf
        For Each kol As Kolon In kolonlar
            If kol.KolonTip = "nvarchar" Then
                sonuc &= kol.KolonAd & " " & kol.KolonTip & "(" & kol.KolonUzunluk & ") NULL," & vbCrLf
            Else
                sonuc &= kol.KolonAd & " " & kol.KolonTip & " NULL," & vbCrLf
            End If
        Next
        sonuc &= "username nvarchar(50) NULL," & vbCrLf
        sonuc &= "create_date datetime NULL," & vbCrLf
        sonuc &= "trans_type nvarchar(3) NULL," & vbCrLf
        sonuc &= "CONSTRAINT [PK_" & HistTabloAdi & "_1] PRIMARY KEY CLUSTERED" & vbCrLf
        sonuc &= "(" & vbCrLf
        sonuc &= "[" & HistTabloAdi & "_id] ASC" & vbCrLf
        sonuc &= ")WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]" & vbCrLf
        sonuc &= ") ON [PRIMARY] " & vbCrLf & vbCrLf
        sonuc &= "GO " & vbCrLf & vbCrLf

        Return sonuc
    End Function


End Class

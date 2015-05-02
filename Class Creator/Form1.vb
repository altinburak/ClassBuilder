Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents btnGetir As System.Windows.Forms.Button
    Friend WithEvents lbDatabaseler As System.Windows.Forms.ListBox
    Friend WithEvents lbTablolar As System.Windows.Forms.ListBox
    Friend WithEvents lbKolonlar As System.Windows.Forms.CheckedListBox
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnKaydet As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents txtSonuc As System.Windows.Forms.RichTextBox
    Friend WithEvents txtClassAd As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cbCollection As System.Windows.Forms.CheckBox
    Friend WithEvents cbTrusted As System.Windows.Forms.CheckBox
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    Friend WithEvents txtUserID As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents cbTools As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnGetir = New System.Windows.Forms.Button
        Me.lbDatabaseler = New System.Windows.Forms.ListBox
        Me.lbTablolar = New System.Windows.Forms.ListBox
        Me.lbKolonlar = New System.Windows.Forms.CheckedListBox
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.txtClassAd = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnKaydet = New System.Windows.Forms.Button
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.cbCollection = New System.Windows.Forms.CheckBox
        Me.txtSonuc = New System.Windows.Forms.RichTextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtServer = New System.Windows.Forms.TextBox
        Me.txtUserID = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cbTrusted = New System.Windows.Forms.CheckBox
        Me.cbTools = New System.Windows.Forms.CheckBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnGetir
        '
        Me.btnGetir.Location = New System.Drawing.Point(104, 240)
        Me.btnGetir.Name = "btnGetir"
        Me.btnGetir.Size = New System.Drawing.Size(48, 32)
        Me.btnGetir.TabIndex = 1
        Me.btnGetir.Text = "Getir"
        '
        'lbDatabaseler
        '
        Me.lbDatabaseler.Location = New System.Drawing.Point(224, 86)
        Me.lbDatabaseler.Name = "lbDatabaseler"
        Me.lbDatabaseler.Size = New System.Drawing.Size(152, 186)
        Me.lbDatabaseler.TabIndex = 2
        '
        'lbTablolar
        '
        Me.lbTablolar.Location = New System.Drawing.Point(396, 86)
        Me.lbTablolar.Name = "lbTablolar"
        Me.lbTablolar.Size = New System.Drawing.Size(152, 186)
        Me.lbTablolar.TabIndex = 2
        '
        'lbKolonlar
        '
        Me.lbKolonlar.ContextMenu = Me.ContextMenu1
        Me.lbKolonlar.Location = New System.Drawing.Point(568, 85)
        Me.lbKolonlar.Name = "lbKolonlar"
        Me.lbKolonlar.Size = New System.Drawing.Size(152, 180)
        Me.lbKolonlar.TabIndex = 3
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Primary Key"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "To String"
        '
        'txtClassAd
        '
        Me.txtClassAd.Location = New System.Drawing.Point(8, 80)
        Me.txtClassAd.Name = "txtClassAd"
        Me.txtClassAd.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtClassAd.Size = New System.Drawing.Size(208, 21)
        Me.txtClassAd.TabIndex = 4
        Me.txtClassAd.Text = ""
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label1.Location = New System.Drawing.Point(8, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(208, 14)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Class Adý"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnKaydet
        '
        Me.btnKaydet.Enabled = False
        Me.btnKaydet.Location = New System.Drawing.Point(160, 240)
        Me.btnKaydet.Name = "btnKaydet"
        Me.btnKaydet.Size = New System.Drawing.Size(56, 32)
        Me.btnKaydet.TabIndex = 1
        Me.btnKaydet.Text = "Kaydet"
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.Filter = "VB files|*.vb|All Files|*.*"
        '
        'cbCollection
        '
        Me.cbCollection.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.cbCollection.ForeColor = System.Drawing.Color.RoyalBlue
        Me.cbCollection.Location = New System.Drawing.Point(8, 112)
        Me.cbCollection.Name = "cbCollection"
        Me.cbCollection.Size = New System.Drawing.Size(136, 24)
        Me.cbCollection.TabIndex = 6
        Me.cbCollection.Text = "Collection da yaz"
        '
        'txtSonuc
        '
        Me.txtSonuc.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.txtSonuc.Location = New System.Drawing.Point(0, 284)
        Me.txtSonuc.Name = "txtSonuc"
        Me.txtSonuc.Size = New System.Drawing.Size(728, 224)
        Me.txtSonuc.TabIndex = 7
        Me.txtSonuc.Text = ""
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label3.Location = New System.Drawing.Point(16, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 23)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Server :"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(76, 24)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtServer.Size = New System.Drawing.Size(56, 21)
        Me.txtServer.TabIndex = 4
        Me.txtServer.Text = "."
        '
        'txtUserID
        '
        Me.txtUserID.Location = New System.Drawing.Point(204, 24)
        Me.txtUserID.Name = "txtUserID"
        Me.txtUserID.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtUserID.Size = New System.Drawing.Size(56, 21)
        Me.txtUserID.TabIndex = 4
        Me.txtUserID.Text = "sa"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label4.Location = New System.Drawing.Point(136, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(64, 23)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "UserID :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label5.Location = New System.Drawing.Point(264, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(80, 23)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Password :"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(348, 24)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtPassword.Size = New System.Drawing.Size(56, 21)
        Me.txtPassword.TabIndex = 4
        Me.txtPassword.Text = "1"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cbTrusted)
        Me.GroupBox1.Controls.Add(Me.txtPassword)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.txtServer)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtUserID)
        Me.GroupBox1.Controls.Add(Me.cbTools)
        Me.GroupBox1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.RoyalBlue
        Me.GroupBox1.Location = New System.Drawing.Point(8, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(712, 56)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Connection String"
        '
        'cbTrusted
        '
        Me.cbTrusted.Location = New System.Drawing.Point(416, 21)
        Me.cbTrusted.Name = "cbTrusted"
        Me.cbTrusted.Size = New System.Drawing.Size(152, 24)
        Me.cbTrusted.TabIndex = 6
        Me.cbTrusted.Text = "Trusted Connection"
        '
        'cbTools
        '
        Me.cbTools.Location = New System.Drawing.Point(584, 21)
        Me.cbTools.Name = "cbTools"
        Me.cbTools.Size = New System.Drawing.Size(120, 24)
        Me.cbTools.TabIndex = 6
        Me.cbTools.Text = "Tools kullan"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label2.Location = New System.Drawing.Point(224, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(136, 14)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Databaseler"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label6.Location = New System.Drawing.Point(400, 64)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(136, 14)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Tablolar"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label7.Location = New System.Drawing.Point(568, 64)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(136, 14)
        Me.Label7.TabIndex = 5
        Me.Label7.Text = "Kolonlar"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Form1
        '
        Me.AcceptButton = Me.btnGetir
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.BackColor = System.Drawing.Color.PapayaWhip
        Me.ClientSize = New System.Drawing.Size(728, 508)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txtSonuc)
        Me.Controls.Add(Me.cbCollection)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtClassAd)
        Me.Controls.Add(Me.lbKolonlar)
        Me.Controls.Add(Me.lbDatabaseler)
        Me.Controls.Add(Me.btnGetir)
        Me.Controls.Add(Me.lbTablolar)
        Me.Controls.Add(Me.btnKaydet)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label7)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Form1"
        Me.Text = "Class Oluþtur"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim baslangic As Integer = 0
    Dim bitis As Integer = 0

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        For Each d As Database In Database.GetDatabases
            Me.lbDatabaseler.Items.Add(d)
        Next
    End Sub

    Private Sub lbDatabaseler_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbDatabaseler.DoubleClick
        Dim d As Database = Me.lbDatabaseler.SelectedItem
        Me.lbTablolar.Items.Clear()
        For Each t As Tablo In d.Tablolar
            Me.lbTablolar.Items.Add(t)
        Next
        Me.lbKolonlar.Items.Clear()
        Me.btnKaydet.Enabled = False
    End Sub

    Private Sub lbTablolar_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbTablolar.DoubleClick
        Dim d As Database = Me.lbDatabaseler.SelectedItem
        Dim t As Tablo = Me.lbTablolar.SelectedItem

        Me.lbKolonlar.Items.Clear()
        For Each k As Kolon In t.Kolonlar(d.DatabaseName)
            Me.lbKolonlar.Items.Add(k)
        Next
        For i As Integer = 0 To Me.lbKolonlar.Items.Count - 1
            Me.lbKolonlar.SetItemChecked(i, True)
        Next
        Me.txtClassAd.Text = t.TabloAdi.Substring(0, t.TabloAdi.Length - 3)
    End Sub

    Private Sub btnGetir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetir.Click


        Dim d As Database = Me.lbDatabaseler.SelectedItem
        Dim t As Tablo = Me.lbTablolar.SelectedItem

        If d Is Nothing Or t Is Nothing Then
            MessageBox.Show("Database, Tablo ve kolon seçili olmalý !!!", "Hata")
            Exit Sub
        End If

        'Connection String'i oluþturmak için kullanýldý...
        Dim BaglantiCumlesi As New Connect
        If Me.cbTools.Checked Then
            BaglantiCumlesi.Tur = "Tools"
        ElseIf Me.cbTrusted.Checked Then
            BaglantiCumlesi.Tur = "Trusted"
            BaglantiCumlesi.Server = Me.txtServer.Text
            BaglantiCumlesi.DatabaseAd = d.DatabaseName
        ElseIf Me.txtUserID.Enabled = True And Me.txtPassword.Enabled = True Then
            BaglantiCumlesi.Tur = "User"
            BaglantiCumlesi.Server = Me.txtServer.Text
            BaglantiCumlesi.DatabaseAd = d.DatabaseName
            BaglantiCumlesi.UserID = Me.txtUserID.Text
            BaglantiCumlesi.Password = Me.txtPassword.Text
        End If


        Tools.ClassAdi = Me.txtClassAd.Text

        If Tools.ClassOl(Me.lbKolonlar.CheckedItems, t.TabloAdi, BaglantiCumlesi) = "" Then
            MessageBox.Show("Önce Primary Key Seçmelisin !!! veya Class Adý girmelisin")
            Exit Sub
        Else
            Me.txtSonuc.Text = Tools.ClassOl(Me.lbKolonlar.CheckedItems, t.TabloAdi, BaglantiCumlesi)
        End If

        If Me.cbCollection.Checked Then
            Me.txtSonuc.Text &= Tools.CollectionOlustur()
        End If

        Dim MaviyeBoyanacakKelimeler() As String = {"private", "public", "as", "class", "integer", "string", "datetime", "get", "set", "end", "shared", "sub", "function", "overrides", "property", "dim", "new", "while", "with", "add", "region", "default", "Inherits", "imports"}

        'Maviye boyamak için
        For i As Integer = 0 To MaviyeBoyanacakKelimeler.Length - 1
            baslangic = 0
            While baslangic < Me.txtSonuc.Text.Length
                baslangic = Me.txtSonuc.Find(MaviyeBoyanacakKelimeler(i), baslangic + 1, RichTextBoxFinds.WholeWord)
                If baslangic = -1 Then Exit While
                Me.txtSonuc.Select(baslangic, MaviyeBoyanacakKelimeler(i).Length)
                Me.txtSonuc.SelectionColor = Color.Blue
            End While
        Next

        'Kýrmýzýya boyamak için

        baslangic = 0
        bitis = 0
        While baslangic < Me.txtSonuc.Text.Length
            baslangic = Me.txtSonuc.Find(Chr(34), bitis + 1, RichTextBoxFinds.None)
            If baslangic = -1 Then Exit While
            bitis = Me.txtSonuc.Find(Chr(34), baslangic + 1, RichTextBoxFinds.None)
            Me.txtSonuc.Select(baslangic, bitis - baslangic + 1)
            Me.txtSonuc.SelectionColor = Color.Red
        End While

        Me.btnKaydet.Enabled = True
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        Dim k As Kolon = Me.lbKolonlar.SelectedItem
        Tools.PrimaryKey = k
    End Sub

    Private Sub btnKaydet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKaydet.Click
        If Me.SaveFileDialog1.ShowDialog = DialogResult.Cancel Then Exit Sub
        Dim yaz As New IO.StreamWriter(Me.SaveFileDialog1.FileName, False, System.Text.Encoding.GetEncoding(1254))
        yaz.Write(Me.txtSonuc.Text)
        yaz.Close()

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        Dim k As Kolon = Me.lbKolonlar.SelectedItem
        Tools.TotoString = k.KolonAd
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbTrusted.CheckedChanged
        If Me.cbTrusted.Checked = True Then

            Me.txtPassword.Enabled = False
            Me.txtUserID.Enabled = False
        Else

            Me.txtPassword.Enabled = True
            Me.txtUserID.Enabled = True
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbTools.CheckedChanged
        If Me.cbTools.Checked = True Then
            Me.txtServer.Enabled = False
            Me.txtPassword.Enabled = False
            Me.txtUserID.Enabled = False
            Me.cbTrusted.Enabled = False
        Else
            Me.txtServer.Enabled = True
            Me.txtPassword.Enabled = True
            Me.txtUserID.Enabled = True
            Me.cbTrusted.Enabled = True
        End If
    End Sub

    Private Sub lbDatabaseler_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbDatabaseler.SelectedIndexChanged

    End Sub
End Class

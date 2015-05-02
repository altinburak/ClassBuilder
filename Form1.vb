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
    Friend WithEvents txtSonuc_DAL As System.Windows.Forms.RichTextBox
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
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents txtSonuc_SP As System.Windows.Forms.RichTextBox
    Friend WithEvents txtSonuc_BL As System.Windows.Forms.RichTextBox
    Friend WithEvents cbSpKullan As System.Windows.Forms.CheckBox
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents Label7 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnGetir = New System.Windows.Forms.Button()
        Me.lbDatabaseler = New System.Windows.Forms.ListBox()
        Me.lbTablolar = New System.Windows.Forms.ListBox()
        Me.lbKolonlar = New System.Windows.Forms.CheckedListBox()
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu()
        Me.MenuItem3 = New System.Windows.Forms.MenuItem()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.txtClassAd = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnKaydet = New System.Windows.Forms.Button()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.cbCollection = New System.Windows.Forms.CheckBox()
        Me.txtSonuc_DAL = New System.Windows.Forms.RichTextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.txtUserID = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cbTrusted = New System.Windows.Forms.CheckBox()
        Me.cbTools = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.txtSonuc_BL = New System.Windows.Forms.RichTextBox()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.txtSonuc_SP = New System.Windows.Forms.RichTextBox()
        Me.cbSpKullan = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnGetir
        '
        Me.btnGetir.Location = New System.Drawing.Point(104, 255)
        Me.btnGetir.Name = "btnGetir"
        Me.btnGetir.Size = New System.Drawing.Size(48, 32)
        Me.btnGetir.TabIndex = 1
        Me.btnGetir.Text = "Getir"
        '
        'lbDatabaseler
        '
        Me.lbDatabaseler.Location = New System.Drawing.Point(224, 115)
        Me.lbDatabaseler.Name = "lbDatabaseler"
        Me.lbDatabaseler.Size = New System.Drawing.Size(152, 173)
        Me.lbDatabaseler.TabIndex = 2
        '
        'lbTablolar
        '
        Me.lbTablolar.Location = New System.Drawing.Point(383, 115)
        Me.lbTablolar.Name = "lbTablolar"
        Me.lbTablolar.Size = New System.Drawing.Size(152, 173)
        Me.lbTablolar.TabIndex = 2
        '
        'lbKolonlar
        '
        Me.lbKolonlar.ContextMenu = Me.ContextMenu1
        Me.lbKolonlar.Location = New System.Drawing.Point(542, 115)
        Me.lbKolonlar.Name = "lbKolonlar"
        Me.lbKolonlar.Size = New System.Drawing.Size(152, 180)
        Me.lbKolonlar.TabIndex = 3
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem3, Me.MenuItem1, Me.MenuItem2})
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 0
        Me.MenuItem3.Text = "Is Deleted"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 1
        Me.MenuItem1.Text = "Primary Key"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 2
        Me.MenuItem2.Text = "To String"
        '
        'txtClassAd
        '
        Me.txtClassAd.Location = New System.Drawing.Point(8, 115)
        Me.txtClassAd.Name = "txtClassAd"
        Me.txtClassAd.Size = New System.Drawing.Size(208, 21)
        Me.txtClassAd.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label1.Location = New System.Drawing.Point(8, 93)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(208, 14)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Class Adý"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnKaydet
        '
        Me.btnKaydet.Enabled = False
        Me.btnKaydet.Location = New System.Drawing.Point(160, 255)
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
        Me.cbCollection.Checked = True
        Me.cbCollection.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbCollection.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.cbCollection.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cbCollection.Location = New System.Drawing.Point(8, 141)
        Me.cbCollection.Name = "cbCollection"
        Me.cbCollection.Size = New System.Drawing.Size(136, 24)
        Me.cbCollection.TabIndex = 6
        Me.cbCollection.Text = "Collection da yaz"
        '
        'txtSonuc_DAL
        '
        Me.txtSonuc_DAL.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtSonuc_DAL.Location = New System.Drawing.Point(3, 3)
        Me.txtSonuc_DAL.Name = "txtSonuc_DAL"
        Me.txtSonuc_DAL.Size = New System.Drawing.Size(681, 234)
        Me.txtSonuc_DAL.TabIndex = 7
        Me.txtSonuc_DAL.Text = ""
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label3.Location = New System.Drawing.Point(16, 28)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 23)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Server :"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(86, 28)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(345, 21)
        Me.txtServer.TabIndex = 4
        '
        'txtUserID
        '
        Me.txtUserID.Location = New System.Drawing.Point(86, 52)
        Me.txtUserID.Name = "txtUserID"
        Me.txtUserID.Size = New System.Drawing.Size(139, 21)
        Me.txtUserID.TabIndex = 4
        Me.txtUserID.Text = "sa"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label4.Location = New System.Drawing.Point(16, 52)
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
        Me.Label5.Location = New System.Drawing.Point(231, 52)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(80, 23)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Password :"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(317, 52)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(114, 21)
        Me.txtPassword.TabIndex = 4
        Me.txtPassword.Text = "Tb111101"
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
        Me.GroupBox1.Location = New System.Drawing.Point(10, 1)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(686, 90)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Connection String"
        '
        'cbTrusted
        '
        Me.cbTrusted.Location = New System.Drawing.Point(483, 28)
        Me.cbTrusted.Name = "cbTrusted"
        Me.cbTrusted.Size = New System.Drawing.Size(152, 24)
        Me.cbTrusted.TabIndex = 6
        Me.cbTrusted.Text = "Trusted Connection"
        '
        'cbTools
        '
        Me.cbTools.Checked = True
        Me.cbTools.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbTools.Location = New System.Drawing.Point(483, 48)
        Me.cbTools.Name = "cbTools"
        Me.cbTools.Size = New System.Drawing.Size(120, 24)
        Me.cbTools.TabIndex = 6
        Me.cbTools.Text = "Tools kullan"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(162, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label2.Location = New System.Drawing.Point(224, 93)
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
        Me.Label6.Location = New System.Drawing.Point(382, 93)
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
        Me.Label7.Location = New System.Drawing.Point(546, 93)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(136, 14)
        Me.Label7.TabIndex = 5
        Me.Label7.Text = "Kolonlar"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.TabControl1.Location = New System.Drawing.Point(0, 293)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(695, 266)
        Me.TabControl1.TabIndex = 9
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.txtSonuc_DAL)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(687, 240)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "DAL"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.txtSonuc_BL)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(687, 240)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "BL"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'txtSonuc_BL
        '
        Me.txtSonuc_BL.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtSonuc_BL.Location = New System.Drawing.Point(3, 3)
        Me.txtSonuc_BL.Name = "txtSonuc_BL"
        Me.txtSonuc_BL.Size = New System.Drawing.Size(681, 234)
        Me.txtSonuc_BL.TabIndex = 0
        Me.txtSonuc_BL.Text = ""
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.txtSonuc_SP)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(687, 240)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "SP"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'txtSonuc_SP
        '
        Me.txtSonuc_SP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtSonuc_SP.Location = New System.Drawing.Point(0, 0)
        Me.txtSonuc_SP.Name = "txtSonuc_SP"
        Me.txtSonuc_SP.Size = New System.Drawing.Size(687, 240)
        Me.txtSonuc_SP.TabIndex = 0
        Me.txtSonuc_SP.Text = ""
        '
        'cbSpKullan
        '
        Me.cbSpKullan.AutoSize = True
        Me.cbSpKullan.Checked = True
        Me.cbSpKullan.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbSpKullan.Location = New System.Drawing.Point(8, 168)
        Me.cbSpKullan.Name = "cbSpKullan"
        Me.cbSpKullan.Size = New System.Drawing.Size(80, 17)
        Me.cbSpKullan.TabIndex = 10
        Me.cbSpKullan.Text = "SP Kullan"
        Me.cbSpKullan.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AcceptButton = Me.btnGetir
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.BackColor = System.Drawing.Color.LightBlue
        Me.ClientSize = New System.Drawing.Size(695, 559)
        Me.Controls.Add(Me.cbSpKullan)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.GroupBox1)
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
        Me.Text = "Sýnýf Oluþturucu"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Dim baslangic As Integer = 0
    Dim bitis As Integer = 0

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.lbDatabaseler.Items.Clear()
        If txtServer.Text <> "" Then
            Dim dc As DatabaseCollection = Database.GetDatabases
            For Each d As Database In dc
                Me.lbDatabaseler.Items.Add(d)
            Next
        End If
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
        Me.txtClassAd.Text = t.TabloAdi
    End Sub

    Private Sub btnGetir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetir.Click
        txtSonuc_DAL.Text = ""
        Dim d As Database = Me.lbDatabaseler.SelectedItem
        Dim t As Tablo = Me.lbTablolar.SelectedItem

        Tools.SpDahilMi = cbSpKullan.Checked

        Dim kc As New KolonCollection
        For Each kln As Kolon In lbKolonlar.CheckedItems
            kc.Add(kln)
        Next

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

        Me.txtSonuc_SP.Text = BuildSP.SpOlustur(kc, t.TabloAdi, BaglantiCumlesi)

        If Tools.ClassOl(kc, t.TabloAdi, BaglantiCumlesi) = "" Then
            MessageBox.Show("Önce Primary Key veya IsDeleted Seçmelisin !!! veya Class Adý girmelisin")
            Exit Sub
        Else
            Me.txtSonuc_DAL.Text = Tools.ClassOl(kc, t.TabloAdi, BaglantiCumlesi)
            Me.txtSonuc_BL.Text = BuildBL.BLOlustur(kc, t.TabloAdi)
        End If

        If Me.cbCollection.Checked Then
            Me.txtSonuc_DAL.Text &= Tools.CollectionOlustur()
        End If

        Boya(txtSonuc_DAL)
        Boya(txtSonuc_BL)
        Boya(txtSonuc_SP)

        Me.btnKaydet.Enabled = True
    End Sub

    Private Sub Boya(ByVal txt As RichTextBox)
        Dim MaviyeBoyanacakKelimeler() As String = {"private", "public", "as", "class", "integer", "string", "datetime", "get", "set", "end", "shared", "sub", "function", "overrides", "property", "dim", "new", "while", "with", "add", "region", "default", "Inherits", "imports", "Return"}

        'Maviye boyamak için
        For i As Integer = 0 To MaviyeBoyanacakKelimeler.Length - 1
            baslangic = 0
            While baslangic < txt.Text.Length
                baslangic = txt.Find(MaviyeBoyanacakKelimeler(i), baslangic + 1, RichTextBoxFinds.WholeWord)
                If baslangic = -1 Then Exit While
                txt.Select(baslangic, MaviyeBoyanacakKelimeler(i).Length)
                txt.SelectionColor = Color.Blue
            End While
        Next

        'Kýrmýzýya boyamak için
        baslangic = 0
        bitis = 0
        While baslangic < txt.Text.Length
            baslangic = txt.Find(Chr(34), bitis + 1, RichTextBoxFinds.None)
            If baslangic = -1 Then Exit While
            bitis = txt.Find(Chr(34), baslangic + 1, RichTextBoxFinds.None)
            txt.Select(baslangic, bitis - baslangic + 1)
            txt.SelectionColor = Color.Red
        End While
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        Dim k As Kolon = Me.lbKolonlar.SelectedItem
        Tools.PrimaryKey = k
    End Sub

    Private Sub btnKaydet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKaydet.Click
        Try

            'If Me.SaveFileDialog1.ShowDialog = DialogResult.Cancel Then Exit Sub
            Dim yaz1 As New IO.StreamWriter("C:\Users\burak.altin\Documents\Visual Studio 2012\Projects\f18_2.0\F19_DAL\" & txtClassAd.Text & "_DAL.vb", False, System.Text.Encoding.GetEncoding(1254))
            yaz1.Write(Me.txtSonuc_DAL.Text)
            yaz1.Close()

            Dim yaz2 As New IO.StreamWriter("C:\Users\burak.altin\Documents\Visual Studio 2012\Projects\F18_2.0\F19_BL\" & txtClassAd.Text & "_BL.vb", False, System.Text.Encoding.GetEncoding(1254))
            yaz2.Write(Me.txtSonuc_BL.Text)
            yaz2.Close()

            Dim yaz3 As New IO.StreamWriter("C:\Users\burak.altin\Desktop\F18 Faz2\SQL\" & txtClassAd.Text & "_SP.sql", False, System.Text.Encoding.GetEncoding(1254))
            yaz3.Write(Me.txtSonuc_SP.Text)
            yaz3.Close()

            MessageBox.Show("Dosyalar kaydedilmiþtir...")

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
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

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Dim k As Kolon = Me.lbKolonlar.SelectedItem
        Tools.IsDeleted = k
    End Sub

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    '    lbKolonlar.Items.Clear()
    '    lbTablolar.Items.Clear()
    '    txtClassAd.Text = ""
    '    lbDatabaseler.Items.Clear()
    '    Dim dc As DatabaseCollection = Database.GetDatabases
    '    For Each d As Database In dc
    '        Me.lbDatabaseler.Items.Add(d)
    '    Next
    'End Sub
End Class

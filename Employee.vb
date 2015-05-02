Imports System.Data.SqlClient

Public Class Employee

	Private mEmployeeID as Integer
	Private mLastName as String
	Private mFirstName as String
	Private mTitle as String
	Private mTitleOfCourtesy as String
	Private mBirthDate as Date
	Private mHireDate as Date
	Private mAddress as String
	Private mCity as String
	Private mRegion as String
	Private mPostalCode as String
	Private mCountry as String
	Private mHomePhone as String
	Private mExtension as String
	Private mPhoto as Byte()
	Private mNotes as String
	Private mReportsTo as Integer
	Private mPhotoPath as String

#Region "Propertyler"

	Public Property EmployeeID as Integer
		Get
			Return mEmployeeID
		End Get
		Set(ByVal Value As Integer)
			mEmployeeID=Value
		End Set
	End Property

	Public Property LastName as String
		Get
			Return mLastName
		End Get
		Set(ByVal Value As String)
			mLastName=Value
		End Set
	End Property

	Public Property FirstName as String
		Get
			Return mFirstName
		End Get
		Set(ByVal Value As String)
			mFirstName=Value
		End Set
	End Property

	Public Property Title as String
		Get
			Return mTitle
		End Get
		Set(ByVal Value As String)
			mTitle=Value
		End Set
	End Property

	Public Property TitleOfCourtesy as String
		Get
			Return mTitleOfCourtesy
		End Get
		Set(ByVal Value As String)
			mTitleOfCourtesy=Value
		End Set
	End Property

	Public Property BirthDate as Date
		Get
			Return mBirthDate
		End Get
		Set(ByVal Value As Date)
			mBirthDate=Value
		End Set
	End Property

	Public Property HireDate as Date
		Get
			Return mHireDate
		End Get
		Set(ByVal Value As Date)
			mHireDate=Value
		End Set
	End Property

	Public Property Address as String
		Get
			Return mAddress
		End Get
		Set(ByVal Value As String)
			mAddress=Value
		End Set
	End Property

	Public Property City as String
		Get
			Return mCity
		End Get
		Set(ByVal Value As String)
			mCity=Value
		End Set
	End Property

	Public Property Region as String
		Get
			Return mRegion
		End Get
		Set(ByVal Value As String)
			mRegion=Value
		End Set
	End Property

	Public Property PostalCode as String
		Get
			Return mPostalCode
		End Get
		Set(ByVal Value As String)
			mPostalCode=Value
		End Set
	End Property

	Public Property Country as String
		Get
			Return mCountry
		End Get
		Set(ByVal Value As String)
			mCountry=Value
		End Set
	End Property

	Public Property HomePhone as String
		Get
			Return mHomePhone
		End Get
		Set(ByVal Value As String)
			mHomePhone=Value
		End Set
	End Property

	Public Property Extension as String
		Get
			Return mExtension
		End Get
		Set(ByVal Value As String)
			mExtension=Value
		End Set
	End Property

	Public Property Photo as Byte()
		Get
			Return mPhoto
		End Get
		Set(ByVal Value As Byte())
			mPhoto=Value
		End Set
	End Property

	Public Property Notes as String
		Get
			Return mNotes
		End Get
		Set(ByVal Value As String)
			mNotes=Value
		End Set
	End Property

	Public Property ReportsTo as Integer
		Get
			Return mReportsTo
		End Get
		Set(ByVal Value As Integer)
			mReportsTo=Value
		End Set
	End Property

	Public Property PhotoPath as String
		Get
			Return mPhotoPath
		End Get
		Set(ByVal Value As String)
			mPhotoPath=Value
		End Set
	End Property


#End Region

#Region "Methodlar"

	Public Shared Sub Ekle(ByVal LastName as String,ByVal FirstName as String,ByVal Title as String,ByVal TitleOfCourtesy as String,ByVal BirthDate as Date,ByVal HireDate as Date,ByVal Address as String,ByVal City as String,ByVal Region as String,ByVal PostalCode as String,ByVal Country as String,ByVal HomePhone as String,ByVal Extension as String,ByVal Photo as Byte(),ByVal Notes as String,ByVal ReportsTo as Integer,ByVal PhotoPath as String)
		Dim Con As New SqlConnection(Tools.ConStr)
		Dim Com As new SqlCommand("INSERT Employees (LastName,FirstName,Title,TitleOfCourtesy,BirthDate,HireDate,Address,City,Region,PostalCode,Country,HomePhone,Extension,Photo,Notes,ReportsTo,PhotoPath) VALUES (@LastName,@FirstName,@Title,@TitleOfCourtesy,@BirthDate,@HireDate,@Address,@City,@Region,@PostalCode,@Country,@HomePhone,@Extension,@Photo,@Notes,@ReportsTo,@PhotoPath)",Con)
		With Com.Parameters
			.Add("@LastName",LastName)
			.Add("@FirstName",FirstName)
			.Add("@Title",Title)
			.Add("@TitleOfCourtesy",TitleOfCourtesy)
			.Add("@BirthDate",BirthDate)
			.Add("@HireDate",HireDate)
			.Add("@Address",Address)
			.Add("@City",City)
			.Add("@Region",Region)
			.Add("@PostalCode",PostalCode)
			.Add("@Country",Country)
			.Add("@HomePhone",HomePhone)
			.Add("@Extension",Extension)
			.Add("@Photo",Photo)
			.Add("@Notes",Notes)
			.Add("@ReportsTo",ReportsTo)
			.Add("@PhotoPath",PhotoPath)
		End With

		Con.Open()
		Com.ExecuteNonQuery()
		Con.Close()
	End Sub

	Public Shared Sub Sil(ByVal EmployeeID as Integer)
		Dim Con As New SqlConnection(Tools.ConStr)
		Dim Com As New SqlCommand("DELETE Employees WHERE EmployeeID=@EmployeeID")
		Com.Parameters.Add("@EmployeeID",EmployeeID)
		Con.Open()
		Com.ExecuteNonQuery()
		Con.Close()
	End Sub

	Public Shared Sub Guncelle(ByVal EmployeeID as Integer,ByVal LastName as String,ByVal FirstName as String,ByVal Title as String,ByVal TitleOfCourtesy as String,ByVal BirthDate as Date,ByVal HireDate as Date,ByVal Address as String,ByVal City as String,ByVal Region as String,ByVal PostalCode as String,ByVal Country as String,ByVal HomePhone as String,ByVal Extension as String,ByVal Photo as Byte(),ByVal Notes as String,ByVal ReportsTo as Integer,ByVal PhotoPath as String)
		Dim Con As New SqlConnection(Tools.ConStr)
		Dim Com As new SqlCommand("UPDATE [Employees]  SET LastName=@LastName,FirstName=@FirstName,Title=@Title,TitleOfCourtesy=@TitleOfCourtesy,BirthDate=@BirthDate,HireDate=@HireDate,Address=@Address,City=@City,Region=@Region,PostalCode=@PostalCode,Country=@Country,HomePhone=@HomePhone,Extension=@Extension,Photo=@Photo,Notes=@Notes,ReportsTo=@ReportsTo,PhotoPath=@PhotoPath WHERE EmployeeID=@EmployeeID",Con)
		With Com.Parameters
			.Add("@EmployeeID",EmployeeID)
			.Add("@LastName",LastName)
			.Add("@FirstName",FirstName)
			.Add("@Title",Title)
			.Add("@TitleOfCourtesy",TitleOfCourtesy)
			.Add("@BirthDate",BirthDate)
			.Add("@HireDate",HireDate)
			.Add("@Address",Address)
			.Add("@City",City)
			.Add("@Region",Region)
			.Add("@PostalCode",PostalCode)
			.Add("@Country",Country)
			.Add("@HomePhone",HomePhone)
			.Add("@Extension",Extension)
			.Add("@Photo",Photo)
			.Add("@Notes",Notes)
			.Add("@ReportsTo",ReportsTo)
			.Add("@PhotoPath",PhotoPath)
		End With

		Con.Open()
		Com.ExecuteNonQuery()
		Con.Close()
	End Sub

	Public Shared Function GetEmployeeByID(ByVal EmployeeID as Integer) As Employee
		Dim e As Employee
		Dim Con As New SqlConnection(Tools.ConStr)
		Dim Com As New SqlCommand("SELECT * FROM Employees WHERE EmployeeID=@EmployeeID",Con)
		Dim Dr as SqlDataReader
		Com.Parameters.Add("@EmployeeID",EmployeeID)
		Con.Open()
		Dr=Com.ExecuteReader
		While Dr.Read()
			e= New Employee
			With e
				e.EmployeeID=IIF(IsDBNull(dr("EmployeeID")), 0,dr("EmployeeID"))
				e.LastName=IIF(IsDBNull(dr("LastName")), "",dr("LastName"))
				e.FirstName=IIF(IsDBNull(dr("FirstName")), "",dr("FirstName"))
				e.Title=IIF(IsDBNull(dr("Title")), "",dr("Title"))
				e.TitleOfCourtesy=IIF(IsDBNull(dr("TitleOfCourtesy")), "",dr("TitleOfCourtesy"))
				e.BirthDate=IIF(IsDBNull(dr("BirthDate")), #1/1/2000#,dr("BirthDate"))
				e.HireDate=IIF(IsDBNull(dr("HireDate")), #1/1/2000#,dr("HireDate"))
				e.Address=IIF(IsDBNull(dr("Address")), "",dr("Address"))
				e.City=IIF(IsDBNull(dr("City")), "",dr("City"))
				e.Region=IIF(IsDBNull(dr("Region")), "",dr("Region"))
				e.PostalCode=IIF(IsDBNull(dr("PostalCode")), "",dr("PostalCode"))
				e.Country=IIF(IsDBNull(dr("Country")), "",dr("Country"))
				e.HomePhone=IIF(IsDBNull(dr("HomePhone")), "",dr("HomePhone"))
				e.Extension=IIF(IsDBNull(dr("Extension")), "",dr("Extension"))
                'e.Photo=IIF(IsDBNull(dr("Photo")), ,dr("Photo"))
				e.Notes=IIF(IsDBNull(dr("Notes")), "",dr("Notes"))
				e.ReportsTo=IIF(IsDBNull(dr("ReportsTo")), 0,dr("ReportsTo"))
				e.PhotoPath=IIF(IsDBNull(dr("PhotoPath")), "",dr("PhotoPath"))
			End With
		End While
		Con.Close()
		Return e
	End Function

	Public Shared Function GetAllEmployee() As Employee()
		Dim al As New ArrayList
		Dim Con As New SqlConnection(Tools.ConStr)
		Dim Com As New SqlCommand("SELECT * FROM [Employees] ",Con)
		Dim Dr as SqlDataReader
		Con.Open()
		Dr=Com.ExecuteReader
		While Dr.Read()
			Dim e As New Employee
			With e
				e.EmployeeID=IIF(IsDBNull(dr("EmployeeID")), 0,dr("EmployeeID"))
				e.LastName=IIF(IsDBNull(dr("LastName")), "",dr("LastName"))
				e.FirstName=IIF(IsDBNull(dr("FirstName")), "",dr("FirstName"))
				e.Title=IIF(IsDBNull(dr("Title")), "",dr("Title"))
				e.TitleOfCourtesy=IIF(IsDBNull(dr("TitleOfCourtesy")), "",dr("TitleOfCourtesy"))
				e.BirthDate=IIF(IsDBNull(dr("BirthDate")), #1/1/2000#,dr("BirthDate"))
				e.HireDate=IIF(IsDBNull(dr("HireDate")), #1/1/2000#,dr("HireDate"))
				e.Address=IIF(IsDBNull(dr("Address")), "",dr("Address"))
				e.City=IIF(IsDBNull(dr("City")), "",dr("City"))
				e.Region=IIF(IsDBNull(dr("Region")), "",dr("Region"))
				e.PostalCode=IIF(IsDBNull(dr("PostalCode")), "",dr("PostalCode"))
				e.Country=IIF(IsDBNull(dr("Country")), "",dr("Country"))
				e.HomePhone=IIF(IsDBNull(dr("HomePhone")), "",dr("HomePhone"))
				e.Extension=IIF(IsDBNull(dr("Extension")), "",dr("Extension"))
                'e.Photo=IIF(IsDBNull(dr("Photo")), ,dr("Photo"))
				e.Notes=IIF(IsDBNull(dr("Notes")), "",dr("Notes"))
				e.ReportsTo=IIF(IsDBNull(dr("ReportsTo")), 0,dr("ReportsTo"))
				e.PhotoPath=IIF(IsDBNull(dr("PhotoPath")), "",dr("PhotoPath"))
			End With
			Al.Add(e)
		End While
		Con.Close()
		Return al.ToArray(GetType(Employee))
	End Function

	Public Shared Function GetAllEmployeeDs() As DataSet
		Dim Con As New SqlConnection(Tools.ConStr)
		Dim Com As New SqlCommand("SELECT * FROM [Employees] ",Con)
		Dim Da as New SqlDataAdapter
		Da.SelectCommand=Com
		Dim Ds as New DataSet
		Da.Fill(Ds)
		Return Ds
	End Function

	Public Shared Function GetAllEmployeeDt() As DataTable
		Dim Con As New SqlConnection(Tools.ConStr)
		Dim Com As New SqlCommand("SELECT * FROM [Employees] ",Con)
		Dim Da as New SqlDataAdapter
		Da.SelectCommand=Com
		Dim Ds as New DataSet
		Da.Fill(Ds)
		Return Ds.Tables(0)
	End Function

	Public Shared Sub UpdateEmployeeDs(ByVal Ds As DataSet)
		Dim Con As New SqlConnection(Tools.ConStr)
		Dim Com As New SqlCommand("SELECT * FROM [Employees] ",Con)
		Dim Da As New SqlDataAdapter
		Da.SelectCommand = Com
		Dim Cb As New SqlCommandBuilder(Da)
		Da.Update(Ds)
	End Sub

	Public Shared Sub UpdateEmployeeDt(ByVal Dt As DataTable)
		Dim Con As New SqlConnection(Tools.ConStr)
		Dim Com As New SqlCommand("SELECT * FROM [Employees] ",Con)
		Dim Da As New SqlDataAdapter
		Da.SelectCommand = Com
		Dim Cb As New SqlCommandBuilder(Da)
		Da.Update(Dt)
	End Sub

#End Region

	Public Overrides Function toString() As String
		Return mFirstName
	End Function

End Class
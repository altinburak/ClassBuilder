Public Class OzellikCollection
    Inherits CollectionBase

    Public Sub Add(ByVal o As Ozellik)
        MyBase.List.Add(o)
    End Sub

    Default Public Property Ozellik(ByVal index As Integer) As Ozellik
        Get
            Return MyBase.InnerList(index)
        End Get
        Set(ByVal Value As Ozellik)
            MyBase.InnerList(index) = Value
        End Set
    End Property

End Class

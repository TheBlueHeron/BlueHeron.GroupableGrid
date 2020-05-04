
''' <summary>
''' 
''' </summary>
Friend Class DataSourceRow
	Inherits CollectionBase

#Region " Objects and variables "

	Private m_Manager As DataSourceManager
	Private m_BoundItem As Object

#End Region

#Region " Properties "

	ReadOnly Property BoundItem As Object
		Get
			Return m_BoundItem
		End Get
	End Property

	Default ReadOnly Property Item(index As Integer) As Object
		Get
			Return List(index)
		End Get
	End Property

#End Region

#Region " Public methods and functions "

	Public Function Add(value As Object) As Integer

		Return List.Add(value)

	End Function

#End Region

#Region " Construction "

	Public Sub New(manager As DataSourceManager, boundItem As Object)

		m_Manager = manager
		m_BoundItem = boundItem

	End Sub

#End Region

End Class
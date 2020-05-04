
''' <summary>
''' 
''' </summary>
Friend Class DataSourceRowComparer
	Implements IComparer

#Region " Objects and variables "

	Private m_BaseComparer As IComparer

#End Region

#Region " Public methods and functions "

	Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare

		Return m_BaseComparer.Compare(CType(x, DataSourceRow).BoundItem, CType(y, DataSourceRow).BoundItem)

	End Function

#End Region

#Region " Construction "

	Public Sub New(baseComparer As IComparer)

		m_BaseComparer = baseComparer

	End Sub

#End Region

End Class
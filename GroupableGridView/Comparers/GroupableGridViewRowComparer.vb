Imports System.ComponentModel

''' <summary>
''' 
''' </summary>
Public Class GroupableGridViewRowComparer
	Implements IComparer

#Region " Objects and variables "

	Private m_Direction As ListSortDirection
	Private m_ColumnIndex As Integer

#End Region

#Region " Public methods and functions "

	Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare

		Return String.Compare(CType(x, GroupableGridViewRow).Cells(m_ColumnIndex).Value.ToString, CType(y, GroupableGridViewRow).Cells(m_ColumnIndex).Value.ToString) * If(m_Direction = ListSortDirection.Ascending, 1, -1)

	End Function

#End Region

#Region " Construction "

	Public Sub New(columnIndex As Integer, sortDirection As ListSortDirection)

		m_ColumnIndex = columnIndex
		m_Direction = sortDirection

	End Sub

#End Region

End Class
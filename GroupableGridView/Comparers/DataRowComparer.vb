Imports System.ComponentModel

''' <summary>
''' 
''' </summary>
Public Class DataRowComparer
	Implements IComparer(Of DataRow)

#Region " Objects and variables "

	Private m_Direction As ListSortDirection
	Private m_ColumnIndex As Integer

#End Region

#Region " Public methods and functions "

	Public Function Compare(x As DataRow, y As DataRow) As Integer Implements IComparer(Of DataRow).Compare

		Return String.Compare(x(m_ColumnIndex).ToString, y(m_ColumnIndex).ToString) * If(m_Direction = ListSortDirection.Ascending, 1, -1)

	End Function

#End Region

#Region " Construction "

	Public Sub New(columnIndex As Integer, sortDirection As ListSortDirection)

		m_Direction = sortDirection
		m_ColumnIndex = columnIndex

	End Sub

#End Region

End Class
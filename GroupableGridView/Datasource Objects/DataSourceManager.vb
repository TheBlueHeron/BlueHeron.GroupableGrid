Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Reflection

''' <summary>
''' The DataDourceManager class is a wrapper class around different types of data sources.
''' In this case the DataSet, object list using reflection and the OutlookGridRow objects are supported by this class.
''' Basically the DataDourceManager works like a facade that provides access in a uniform way to the data source.
''' </summary>
Public Class DataSourceManager

#Region " Objects and variables "

	Private m_DataSource As Object
	Private m_DataMember As String

	Public Columns As ArrayList
	Public Rows As ArrayList
	Public ColumnTypes As List(Of Type)

#End Region

#Region " Properties "

	Public ReadOnly Property DataMember As String
		Get
			Return m_DataMember
		End Get
	End Property

	Public ReadOnly Property DataSource As Object
		Get
			Return m_DataSource
		End Get
	End Property

#End Region

#Region " Public methods and functions "

	Public Sub Sort(comparer As System.Collections.IComparer)

		Rows.Sort(New DataSourceRowComparer(comparer))

	End Sub

#End Region

#Region " Private methods and functions "

	Private Sub InitManager()

		Columns = New ArrayList()
		Rows = New ArrayList()
		ColumnTypes = New List(Of Type)

		If TypeOf (m_DataSource) Is IListSource Then
			InitDataSet()
		ElseIf TypeOf (m_DataSource) Is IList Then
			InitList()
		ElseIf TypeOf (m_DataSource) Is GroupableGridView Then
			InitGrid()
		End If

	End Sub

	Private Sub InitDataSet()
		Dim table As DataTable = CType(DataSource, DataSet).Tables(m_DataMember)

		For Each c As DataColumn In table.Columns
			Columns.Add(c.ColumnName)
			ColumnTypes.Add(c.GetType())
		Next
		For Each r As DataRow In table.Rows
			Dim row As New DataSourceRow(Me, r)

			For i As Integer = 0 To Columns.Count - 1
				row.Add(r(i))
			Next
			Rows.Add(row)
		Next

	End Sub

	Private Sub InitGrid()
		Dim grid As GroupableGridView = CType(DataSource, GroupableGridView)

		For Each c As DataGridViewColumn In grid.Columns
			Columns.Add(c.Name)
			ColumnTypes.Add(c.GetType())
		Next

		For Each r As GroupableGridViewRow In grid.Rows
			If Not r.IsGroupRow AndAlso Not r.IsNewRow Then
				Dim row As New DataSourceRow(Me, r)

				For i As Integer = 0 To Columns.Count - 1
					row.Add(r.Cells(i).Value)
				Next
				Rows.Add(row)
			End If
		Next

	End Sub

	Private Sub InitList()
		Dim list As IList = CType(DataSource, IList)
		Dim bf As BindingFlags = BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.GetProperty
		Dim props As PropertyInfo() = list(0).GetType().GetProperties()

		For Each pi As PropertyInfo In props
			Columns.Add(pi.Name)
			ColumnTypes.Add(pi.PropertyType)
		Next
		For Each obj As Object In list
			Dim row As New DataSourceRow(Me, obj)

			For Each pi As PropertyInfo In props
				Dim result As Object = obj.GetType().InvokeMember(pi.Name, bf, Nothing, obj, Nothing)
				row.Add(result)
			Next
			Rows.Add(row)
		Next

	End Sub

#End Region

#Region " Construction "

	Public Sub New(dataSource As Object, dataMember As String)

		m_DataSource = dataSource
		m_DataMember = dataMember
		InitManager()

	End Sub

#End Region

End Class
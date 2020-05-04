Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' 
''' </summary>
Public Class GroupableGridView
	Inherits DataGridView

#Region " Objects and variables "

	Private components As IContainer = Nothing

	Private m_DataSource As DataSourceManager
	Private m_GroupTemplate As IGroupableGridViewGroup
	Private m_IconCollapse As Image
	Private m_IconExpand As Image

#End Region

#Region " Properties "

	<Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
	Public Shadows ReadOnly Property RowTemplate As DataGridViewRow
		Get
			Return MyBase.RowTemplate
		End Get
	End Property

	<Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
	Public Property GroupTemplate As IGroupableGridViewGroup
		Get
			Return m_GroupTemplate
		End Get
		Set(value As IGroupableGridViewGroup)
			m_GroupTemplate = value
		End Set
	End Property

	<Category("Appearance")>
	Public Property CollapseIcon As Image
		Get
			Return m_IconCollapse
		End Get
		Set(value As Image)
			m_IconCollapse = value
		End Set
	End Property

	<Category("Appearance")>
	Public Property ExpandIcon As Image
		Get
			Return m_IconExpand
		End Get
		Set(value As Image)
			m_IconExpand = value
		End Set
	End Property

	Public Shadows ReadOnly Property DataSource As Object
		Get
			If m_DataSource Is Nothing Then
				Return Nothing
			End If
			If m_DataSource.DataSource.Equals(Me) Then
				Return Nothing
			End If
			Return m_DataSource.DataSource
		End Get
	End Property

#End Region

#Region " Public methods and functions "

	Public Sub CollapseAll()

		SetGroupCollapse(True)

	End Sub

	Public Sub ExpandAll()

		SetGroupCollapse(False)

	End Sub

	Public Sub ClearGroups()

		m_GroupTemplate.Column = Nothing
		FillGrid(Nothing)

	End Sub

	Public Sub BindData(dataSource As Object, dataMember As String, groupMember As String)

		Me.DataMember = dataMember
		If dataSource Is Nothing Then
			m_DataSource = Nothing
			Columns.Clear()
		Else
			m_DataSource = New DataSourceManager(dataSource, dataMember)
			SetupColumns(groupMember)
			FillGrid(Nothing)
		End If

	End Sub

	Public Overrides Sub Sort(comparer As System.Collections.IComparer)

		If m_DataSource Is Nothing Then
			m_DataSource = New DataSourceManager(Me, Nothing)
		End If
		m_DataSource.Sort(comparer)
		FillGrid(m_GroupTemplate)

	End Sub

	Public Overrides Sub Sort(dataGridViewColumn As DataGridViewColumn, direction As ListSortDirection)

		If m_DataSource Is Nothing Then
			m_DataSource = New DataSourceManager(Me, Nothing)
		End If

		m_DataSource.Sort(New GroupableGridViewRowComparer(dataGridViewColumn.Index, direction))
		FillGrid(GroupTemplate)

	End Sub

	Public Sub SetOutlookStyle()
		Dim dataGridViewCellStyle2 As New DataGridViewCellStyle()
		Dim dataGridViewCellStyle3 As New DataGridViewCellStyle()

		AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill

		dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
		dataGridViewCellStyle2.Font = New System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte(0)))
		dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
		dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
		dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
		DefaultCellStyle = dataGridViewCellStyle2
		AlternatingRowsDefaultCellStyle = dataGridViewCellStyle2

		dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
		dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Desktop
		dataGridViewCellStyle3.Font = New System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte(0)))
		dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
		dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
		dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
		RowHeadersDefaultCellStyle = dataGridViewCellStyle3

		GridColor = System.Drawing.SystemColors.Control
		BackgroundColor = System.Drawing.SystemColors.Window
		CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
		RowHeadersVisible = False
		ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
		Font = New System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte(0)))

		ClearGroups()

	End Sub

#End Region

#Region " Private methods and functions "

	''' <summary>
	''' The fill grid method fills the grid with the data from the DataSourceManager. It takes the grouping style into account, if it is set.
	''' </summary>
	Private Sub FillGrid(groupingStyle As IGroupableGridViewGroup)
		Dim list As ArrayList
		Dim row As GroupableGridViewRow

		Rows.Clear()

		If m_DataSource Is Nothing Then
			Exit Sub
		Else
			list = m_DataSource.Rows
		End If
		If list.Count <= 0 Then
			Exit Sub
		End If

		If groupingStyle Is Nothing Then
			For Each r As DataSourceRow In list
				row = CType(Me.RowTemplate.Clone(), GroupableGridViewRow)

				For Each val As Object In r
					row.Cells.Add(CreateCell(val))
				Next
				Rows.Add(row)
			Next
		Else
			Dim groupCur As IGroupableGridViewGroup = Nothing
			Dim result As Object = Nothing
			Dim counter As Integer = 0

			For Each r As DataSourceRow In list
				row = CType(Me.RowTemplate.Clone(), GroupableGridViewRow)
				result = r(groupingStyle.Column.Index)
				If Not groupCur Is Nothing AndAlso groupCur.CompareTo(result) = 0 Then ' item is part of the group
					row.Group = groupCur
					counter += 1
				Else ' item is not part of the group, so create new group
					If Not groupCur Is Nothing Then
						groupCur.ItemCount = counter
					End If
					groupCur = CType(groupingStyle.Clone(), IGroupableGridViewGroup)
					groupCur.Value = result
					row.Group = groupCur
					row.IsGroupRow = True
					row.Height = groupCur.Height
					row.CreateCells(Me, groupCur.Value)
					Rows.Add(row)

					row = CType(Me.RowTemplate.Clone(), GroupableGridViewRow)
					row.Group = groupCur
					counter = 1
				End If


				For Each val As Object In r
					row.Cells.Add(CreateCell(val))
				Next
				Rows.Add(row)
				groupCur.ItemCount = counter
			Next
		End If

	End Sub

	Private Sub InitializeComponent()

		components = New Container

	End Sub

	Protected Overrides Sub OnCellBeginEdit(e As DataGridViewCellCancelEventArgs)
		Dim row As GroupableGridViewRow = CType(MyBase.Rows(e.RowIndex), GroupableGridViewRow)

		If (row.IsGroupRow) Then
			e.Cancel = True
		Else
			MyBase.OnCellBeginEdit(e)
		End If

	End Sub

	Protected Overrides Sub OnCellDoubleClick(e As DataGridViewCellEventArgs)

		If e.RowIndex >= 0 Then
			Dim row As GroupableGridViewRow = CType(MyBase.Rows(e.RowIndex), GroupableGridViewRow)

			If row.IsGroupRow Then
				row.Group.IsCollapsed = Not row.Group.IsCollapsed

				' this is a workaround to make the grid re-calculate it's contents and background bounds so the background is updated correctly.
				' this will also invalidate the control, so it will redraw itself
				row.Visible = False
				row.Visible = True
				Exit Sub
			End If
		End If

		MyBase.OnCellClick(e)

	End Sub

	Protected Overrides Sub OnCellMouseDown(e As DataGridViewCellMouseEventArgs)

		If e.RowIndex < 0 Then
			Exit Sub
		End If

		Dim row As GroupableGridViewRow = CType(MyBase.Rows(e.RowIndex), GroupableGridViewRow)

		If row.IsGroupRow AndAlso row.IsIconHit(e) Then
			row.Group.IsCollapsed = Not row.Group.IsCollapsed
			row.Visible = False
			row.Visible = True
		Else
			MyBase.OnCellMouseDown(e)
		End If

	End Sub

	Private Sub SetGroupCollapse(collapsed As Boolean)

		If Rows.Count = 0 Then
			Exit Sub
		End If
		If m_GroupTemplate Is Nothing Then
			Exit Sub
		End If

		' set the default grouping style template collapsed property
		m_GroupTemplate.IsCollapsed = collapsed
		' loop through all rows to find the GroupRows
		For Each row As GroupableGridViewRow In Rows
			If row.IsGroupRow Then
				row.Group.IsCollapsed = collapsed
			End If
		Next

		Rows(0).Visible = Not Rows(0).Visible
		Rows(0).Visible = Not Rows(0).Visible

	End Sub

	Private Sub SetupColumns(groupMember As String)
		Dim list As ArrayList

		' clear all columns, this is a somewhat crude implementation refinement may be welcome.
		Columns.Clear()
		' start filling the grid
		If m_DataSource Is Nothing Then
			Exit Sub
		Else
			list = m_DataSource.Rows
		End If
		If list.Count <= 0 Then
			Exit Sub
		End If

		For i As Integer = 0 To m_DataSource.Columns.Count - 1
			Dim index As Integer
			Dim name As String = CStr(m_DataSource.Columns(i))
			Dim column As DataGridViewColumn = Columns(name)

			If column Is Nothing Then
				If name = groupMember Then
					index = Columns.Add(CreateColumn(name, m_DataSource.ColumnTypes(i), True))
					m_GroupTemplate.Column = Columns(index)
				Else
					index = Columns.Add(CreateColumn(name, m_DataSource.ColumnTypes(i), False))
				End If
			Else
				index = column.Index
			End If
			Columns(index).SortMode = DataGridViewColumnSortMode.Programmatic ' always programmatic!
		Next

	End Sub

	Private Function CreateColumn(name As String, typ As Type, isGroupingColumn As Boolean) As DataGridViewColumn

		If typ = GetType(Boolean) Then
			Return New DataGridViewCheckBoxColumn With {.HeaderText = name}
		ElseIf typ.BaseType = GetType([Enum]) Then
			If isGroupingColumn Then
				Return New DataGridViewTextBoxColumn With {.HeaderText = name}
			Else
				Return New DataGridViewComboBoxColumn With {.HeaderText = name, .DataSource = typ.ToDataSource(), .ValueMember = "Value", .DisplayMember = "Name"}
			End If
		Else
			Return New DataGridViewTextBoxColumn With {.HeaderText = name}
		End If


	End Function

	Private Function CreateCell(value As Object) As DataGridViewCell
		Dim type As Type = value.GetType

		If type = GetType(Boolean) Then
			Return New DataGridViewCheckBoxCell With {.Value = value}
		ElseIf type.BaseType = GetType([Enum]) Then
			Return New DataGridViewComboBoxCell With {.Value = value}
		Else
			Return New DataGridViewTextBoxCell With {.Value = value}
		End If


	End Function

#End Region

#Region " Construction and destruction "

	Public Sub New()

		InitializeComponent()
		MyBase.RowTemplate = New GroupableGridViewRow
		m_GroupTemplate = New GroupableGridViewDefaultGroup

	End Sub

	Protected Overrides Sub Dispose(disposing As Boolean)

		If disposing AndAlso Not components Is Nothing Then
			components.Dispose()
		End If
		MyBase.Dispose(disposing)

	End Sub

#End Region

	Private Sub GroupableGridView_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles Me.DataError

	End Sub
End Class
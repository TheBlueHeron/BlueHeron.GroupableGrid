Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' The OutlookGridRow has 2 main additional properties: the Group it belongs to and a IsRowGroup flag that indicates whether the OutlookGridRow object behaves like a regular row (with data) or should behave like a Group row.
''' </summary>
Public Class GroupableGridViewRow
	Inherits DataGridViewRow

#Region " Objects and variables "

	Private m_IsGroupRow As Boolean
	Private m_Group As IGroupableGridViewGroup

#End Region

#Region " Properties "

	<DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
	Public Property Group As IGroupableGridViewGroup
		Get
			Return m_Group
		End Get
		Set(value As IGroupableGridViewGroup)
			m_Group = value
		End Set
	End Property

	<DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
	Public Property IsGroupRow As Boolean
		Get
			Return m_IsGroupRow
		End Get
		Set(value As Boolean)
			m_IsGroupRow = value
		End Set
	End Property

#End Region

#Region " Private methods and functions "

	''' <summary>
	''' This function checks if the user hit the expand (+) or collapse (-) icon. If it was hit it will return true.
	''' </summary>
	''' <param name="e">Mouse click event arguments</param>
	''' <returns>Returns true if the icon was hit, false otherwise</returns>
	Friend Function IsIconHit(e As DataGridViewCellMouseEventArgs) As Boolean

		If e.ColumnIndex < 0 Then
			Return False
		End If

		Dim grid As GroupableGridView = CType(Me.DataGridView, GroupableGridView)
		Dim rowBounds As Rectangle = grid.GetRowDisplayRectangle(Me.Index, False)
		Dim x As Integer = e.X
		Dim c As DataGridViewColumn = grid.Columns(e.ColumnIndex)

		If m_IsGroupRow AndAlso (c.DisplayIndex = 0) AndAlso (x > rowBounds.Left + 4) AndAlso (x < rowBounds.Left + 16) AndAlso (e.Y > rowBounds.Height - 18) AndAlso (e.Y < rowBounds.Height - 7) Then
			Return True
		End If

		Return False

	End Function

	Protected Overrides Sub Paint(graphics As Graphics, clipBounds As Rectangle, rowBounds As Rectangle, rowIndex As Integer, rowState As DataGridViewElementStates, isFirstDisplayedRow As Boolean, isLastVisibleRow As Boolean)

		If m_IsGroupRow Then
			Dim grid As GroupableGridView = CType(Me.DataGridView, GroupableGridView)
			Dim rowHeadersWidth As Integer = If(grid.RowHeadersVisible, grid.RowHeadersWidth, 0)

			' this can be optimized
			Dim brush As New SolidBrush(grid.DefaultCellStyle.BackColor)
			Dim brush2 As New SolidBrush(Color.FromKnownColor(KnownColor.GradientActiveCaption))
			Dim gridWidth As Integer = grid.Columns.GetColumnsWidth(DataGridViewElementStates.Displayed)
			Dim rowBounds2 As Rectangle = grid.GetRowDisplayRectangle(Me.Index, True)

			' draw the background
			graphics.FillRectangle(brush, rowBounds.Left + rowHeadersWidth - grid.HorizontalScrollingOffset, rowBounds.Top, gridWidth, rowBounds.Height - 1)
			' draw text, using the current grid font
			graphics.DrawString(Group.Text, grid.Font, Brushes.Black, rowHeadersWidth - grid.HorizontalScrollingOffset + 23, rowBounds.Bottom - 18)
			'draw bottom line
			graphics.FillRectangle(brush2, rowBounds.Left + rowHeadersWidth - grid.HorizontalScrollingOffset, rowBounds.Bottom - 2, gridWidth - 1, 2)

			' draw right vertical bar
			If (grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleVertical OrElse grid.CellBorderStyle = DataGridViewCellBorderStyle.Single) Then
				graphics.FillRectangle(brush2, rowBounds.Left + rowHeadersWidth - grid.HorizontalScrollingOffset + gridWidth - 1, rowBounds.Top, 1, rowBounds.Height)
			End If
			If Group.IsCollapsed Then
				If Not grid.ExpandIcon Is Nothing Then
					graphics.DrawImage(grid.ExpandIcon, rowBounds.Left + rowHeadersWidth - grid.HorizontalScrollingOffset + 4, rowBounds.Bottom - 18, 11, 11)
				End If
			Else
				If Not grid.CollapseIcon Is Nothing Then
					graphics.DrawImage(grid.CollapseIcon, rowBounds.Left + rowHeadersWidth - grid.HorizontalScrollingOffset + 4, rowBounds.Bottom - 18, 11, 11)
				End If
			End If
			brush.Dispose()
			brush2.Dispose()
		End If

		MyBase.Paint(graphics, clipBounds, rowBounds, rowIndex, rowState, isFirstDisplayedRow, isLastVisibleRow)

	End Sub

	Protected Overrides Sub PaintCells(graphics As Graphics, clipBounds As Rectangle, rowBounds As Rectangle, rowIndex As Integer, rowState As DataGridViewElementStates, isFirstDisplayedRow As Boolean, isLastVisibleRow As Boolean, paintParts As DataGridViewPaintParts)

		If Not m_IsGroupRow Then
			Try
				MyBase.PaintCells(graphics, clipBounds, rowBounds, rowIndex, rowState, isFirstDisplayedRow, isLastVisibleRow, paintParts)
			Catch ex As Exception

			End Try
		End If

	End Sub

#End Region

#Region " Public methods and functions "

	Public Overrides Function GetState(rowIndex As Integer) As DataGridViewElementStates

		If (Not m_IsGroupRow) AndAlso (Not m_Group Is Nothing) AndAlso m_Group.IsCollapsed Then
			Return MyBase.GetState(rowIndex) And DataGridViewElementStates.Selected
		End If

		Return MyBase.GetState(rowIndex)

	End Function

#End Region

#Region " Construction "

	Public Sub New()

		Me.New(Nothing, False)

	End Sub

	Public Sub New(group As IGroupableGridViewGroup)

		Me.New(group, False)

	End Sub

	Public Sub New(group As IGroupableGridViewGroup, isGroupRow As Boolean)

		MyBase.New()
		m_Group = group
		m_IsGroupRow = isGroupRow

	End Sub

#End Region

End Class

Imports System.Windows.Forms

''' <summary>
''' IOutlookGridGroup specifies the interface of any implementation of a OutlookGridGroup class
''' Each implementation of the IOutlookGridGroup can override the behavior of the grouping mechanism
''' Notice also that ICloneable must be implemented. The OutlookGrid makes use of the Clone method of the Group
''' to create new Group clones. Related to this is the OutlookGrid.GroupTemplate property, which determines what
''' type of Group must be cloned.
''' </summary>
Public Interface IGroupableGridViewGroup
	Inherits ICloneable, IComparable

	''' <summary>
	''' The text to be displayed in the group row.
	''' </summary>
	Property Text As String

	''' <summary>
	''' Determines the value of the current group. This is used to compare the group value against each item's value.
	''' </summary>
	Property Value As Object

	''' <summary>
	''' Indicates whether the group is collapsed. If it is collapsed, its group items (rows) will not be displayed.
	''' </summary>
	Property IsCollapsed As Boolean

	''' <summary>
	''' Specifies which column is associated with this group.
	''' </summary>
	''' <value><see cref="DataGridViewColumn" /></value>
	''' <returns><see cref="DataGridViewColumn" /></returns>
	Property Column As DataGridViewColumn

	''' <summary>
	''' Specifies the number of items that are part of the current group. This value is automatically updated each time the grid is redrawn, e.g. after sorting the grid.
	''' </summary>
	Property ItemCount As Integer

	''' <summary>
	''' Specifies the default height of the group. Each group is cloned from the GroupStyle object.
	''' Setting the height of this object will also set the default height of each group.
	''' </summary>
	Property Height As Integer

End Interface
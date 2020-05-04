Imports System.Windows.Forms

''' <summary>
''' 
''' </summary>
Public Class GroupableGridViewDefaultGroup
	Implements IGroupableGridViewGroup

#Region " Objects and variables "

	Protected m_Val As Object
	Protected m_Text As String
	Protected m_Collapsed As Boolean
	Protected m_Column As DataGridViewColumn
	Protected m_ItemCount As Integer
	Protected m_Height As Integer

#End Region

#Region " Properties "

	Public Property Column As Windows.Forms.DataGridViewColumn Implements IGroupableGridViewGroup.Column
		Get
			Return m_Column
		End Get
		Set(value As Windows.Forms.DataGridViewColumn)
			m_Column = value
		End Set
	End Property

	Public Property Height As Integer Implements IGroupableGridViewGroup.Height
		Get
			Return m_Height
		End Get
		Set(value As Integer)
			m_Height = value
		End Set
	End Property

	Public Property IsCollapsed As Boolean Implements IGroupableGridViewGroup.IsCollapsed
		Get
			Return m_Collapsed
		End Get
		Set(value As Boolean)
			m_Collapsed = value
		End Set
	End Property

	Public Property ItemCount As Integer Implements IGroupableGridViewGroup.ItemCount
		Get
			Return m_ItemCount
		End Get
		Set(value As Integer)
			m_ItemCount = value
		End Set
	End Property

	Public Overridable Property Text As String Implements IGroupableGridViewGroup.Text
		Get
			If m_Column Is Nothing Then
				Return String.Format("Unbound group: {0} ({1})", m_Val, If(m_ItemCount = 1, "1 item", m_ItemCount & " items"))
			Else
				Return String.Format("{0}: {1} ({2})", m_Column.HeaderText, m_Val, If(m_ItemCount = 1, "1 item", m_ItemCount & " items"))
			End If
		End Get
		Set(value As String)
			m_Text = value
		End Set
	End Property

	Public Overridable Property Value As Object Implements IGroupableGridViewGroup.Value
		Get
			Return m_Val
		End Get
		Set(value As Object)
			m_Val = value
		End Set
	End Property

#End Region

#Region " Public methods and functions "

	Public Overridable Function Clone() As Object Implements ICloneable.Clone

		Return New GroupableGridViewDefaultGroup With {.m_Column = m_Column, .m_Val = m_Val, .m_Collapsed = m_Collapsed, .m_Text = m_Text, .m_Height = m_Height}

	End Function

	Public Overridable Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo

		Return String.Compare(m_Val.ToString, obj.ToString)

	End Function

#End Region

#Region " Construction "

	Public Sub New()

		m_Val = Nothing
		m_column = Nothing
		m_Height = 34 ' default height

	End Sub

#End Region

End Class

''' <summary>
''' 
''' </summary>
Public Class GroupableGridViewAlphabeticGroup
	Inherits GroupableGridViewDefaultGroup

#Region " Properties "

	Public Overrides Property Text As String
		Get
			Return String.Format("Alphabetic: {1} ({2})", m_Column.HeaderText, m_Val, If(m_ItemCount = 1, "1 item", m_ItemCount & " items"))
		End Get
		Set(value As String)
			m_Text = value
		End Set
	End Property

	Public Overrides Property Value As Object
		Get
			Return m_Val
		End Get
		Set(value As Object)
			m_Val = value.ToString.Substring(0, 1).ToUpper
		End Set
	End Property

#End Region

#Region " Public methods and functions "

	Public Overrides Function Clone() As Object

		Return New GroupableGridViewAlphabeticGroup With {.m_Column = m_Column, .m_Val = m_Val, .m_Collapsed = m_Collapsed, .m_Text = m_Text, .m_Height = m_Height}

	End Function

	Public Overrides Function CompareTo(obj As Object) As Integer

		Return String.Compare(m_Val.ToString, obj.ToString.Substring(0, 1).ToUpper)

	End Function

#End Region

#Region " Construction "

	Public Sub New()

		MyBase.New()

	End Sub

#End Region

End Class
Imports System.Runtime.CompilerServices

Module Extensions

	Public Structure EnumItem
		Public Name As String
		Public Value As Integer
	End Structure

	<Extension()>
	Public Function ToDataSource(enumeration As Type) As IEnumerable(Of EnumItem)
		Dim lst As New List(Of EnumItem)
		
		If enumeration.BaseType = GetType(System.Enum) Then
			For Each strName As String In System.Enum.GetNames(enumeration)
				lst.Add(New EnumItem With {.Name = strName, .Value = CInt(System.Enum.Parse(enumeration, strName))})
			Next
		End If

		Return lst

	End Function

End Module
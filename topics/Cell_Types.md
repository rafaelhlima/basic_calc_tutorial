# Dealing with Cell Types



```VBA
Sub PrintCellType
	'Gets the currently selected cell
	Dim oCell as Object
	oCell = ThisComponent.getCurrentSelection()
	'Enum with constants corresponding to all possible cell types
	Dim eTypes as Variant
	eTypes = com.sun.star.table.CellContentType
	'Check cell type
	Select Case oCell.getType()
		Case eTypes.EMPTY
			MsgBox "Empty Cell"
		Case eTypes.VALUE
			MsgBox "Cell with Numeric Value"
		Case eTypes.TEXT
			MsgBox "Cell with Text"
		Case eTypes.FORMULA
			MsgBox "Cell with Formula"
		Case Else
			MsgBox "Something Else"
	End Select
End Sub
```

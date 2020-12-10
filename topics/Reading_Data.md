# Reading values from Cells

Macros written for LibreOffice Calc often need to read data from spreadsheets to perform calculations or any other desired operation. In this topic I'll explain the most commonly used methods to (i) access cell objects and then (ii) read the data stored in these cells.

## Accessing Cell Objects

Before reading values in a cell or in multiple cells, we first need to get access to cell objects. If you read the [Hello World] example, you'll remember this part of the code.

```VBA
oSheet = ThisComponent.getCurrentController.getActiveSheet()
oCell = oSheet.getCellRangeByName("A1")
oCell.setString("Hello World!")
```

The first line assigns to *oSheet* the object that gives access to the Active Sheet of the current document. Then we use this object to get access to cell *A1* and store its corresponding object in *oCell*. Finally, in the third line we use the *oCell* object to call the method *setString("Hello World!")* that inputs the string into the cell.

This means that the key to reading/writing contents from/to cells is to first create an object that will grant us access to its properties and methods. In the following sections we'll discuss the many methods that can be used for this purpose.

## Getting Acess to Sheets

As discussed above, before getting acces to cells we first need to access a sheet object. As a first example, consider the following code:

```VBA
Sub ReadData
	Dim mySheet as Object
	Dim myCell as Object
	mySheet = ThisComponent.Sheets(0)
	myCell = mySheet.getCellRangeByName("A1")
	MsgBox myCell.getString
End Sub
```

This macro reads the string value in cell *A1* of the first sheet in the Calc file and shows a message box with this string. Note that here we are acessing the sheet using its index number, which always starts at zero.

```VBA
mySheet = ThisComponent.Sheets(0)
```

Hence, if we have three sheets in our file, then their indices will range from 0 to 2, as shown below.

[Sheet Indices](../images/Reading_Data_01.png)

Another approach is to access sheets by their names, as shown in the code below:

```VBA
Sub ReadData
	Dim mySheet as Object
	Dim myCell as Object
	If ThisComponent.Sheets.HasByName("Balance") Then
		mySheet = ThisComponent.Sheets.getByName("Balance")
	Else
		MsgBox "The sheet 'Balance' does not exist"
		Exit Sub
	End If
	myCell = mySheet.getCellRangeByName("A1")
	MsgBox myCell.getString
End Sub
```

In this example we are reading the contents of cell *A1* in a sheet named "Balance". Note that here the macro first tests if the file has a sheet with this name before proceeding. This is done with the *HasByName()* method of the sheet object.

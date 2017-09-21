' -----------------------------------------------
' MakeTable
' -----------------------------------------------
' Make table for DokuWiki
'
' No parameter: it is a Sub
'
' Usage:
' 1. Make a table  wherever in a spreadsheet.
' 2. Format cells as desired; it support merge cells and justifications (left, center, right)
' 3. Select the range to be encoded as DokuWiki table.
' 4. Run the macro.
' 5. The result will be written from the row 5 rows below of the selection.
' Important note: if there are existing values in the cells between the selection and the result rows it will not be deleted. So make sure that the cells are cleared beforehand.

Sub MakeTable ()

Dim mergedCells(0, 4) As Integer ' Array to hold data of merged cells. A row of the array has four elements: firstRow, firstCol, numCol, numRow.
							 ' For example, cells A1:C2 and D3:D4 are merged, the array shall hold (1, 1, 3, 2), (3, 4, 1, 2)

activeWindow = ThisComponent
activeSheet = activeWindow.CurrentController.ActiveSheet

' sourceCells is the range of cells to be encoded to DokuWiki table code.
sourceCells = activeWindow.getCurrentSelection()
sourceRowsCount = sourceCells.Rows.Count
sourceColsCount = sourceCells.Columns.Count

' Get the bottom of the sourceCells
sourceCellsBottom = sourceCells.getCellByPosition(0, sourceRowsCount - 1)
sourceBottomAddress = sourceCellsBottom.getCellAddress()
sourceBottomRow = sourceBottomAddress.Row

' The result will be 5 rows below of the bottom of the sourceCells
targetRowIndex = sourceBottomRow + 5

' It will iterate through rows and columns of sourceCells.
For row = 0 to sourceRowsCount - 1

	targetCell = activeSheet.getCellByPosition(0, targetRowIndex + row)
	
	If row = 0 Then
		divisor = "^"	' Divisor for header
	Else
		divisor = "|"	' Divisor for body
	End If
	
	' Put divisor at the first place of the row
	targetCell.SetString(divisor)
	
	For col = 0 to sourceColsCount - 1
		aCell = sourceCells.getCellByPosition(col, row)		'Note that getCellByPosition does not index the cell (row, col) but (col, row)
		
		If aCell.IsMerged Then
			mergedCellHoriSpan = MergedCellGetColSpanCount(aCell)
			mergedCellVertSpan = MergedCellGetRowSpanCount(aCell)
			
			maxMergedCells = UBound(mergedCells, 1)
			ReDim mergedCells(maxMergedCells + 1, 4)
			
			mergedCells(UBound(mergedCells, 1), 1) = row
			mergedCells(UBound(mergedCells, 1), 2) = col
			mergedCells(UBound(mergedCells, 1), 3) = mergedCellHoriSpan
			mergedCells(UBound(mergedCells, 1), 4) = mergedCellVertSpan
		End If
		
		hiddenRow = False
		hiddenCol = False
		
		' Check if the current cell is hidden by merge either in row or col.
		For r = 0 to UBound(mergedCells, 1)
			If col = mergedCells(r, 2) and row > mergedCells(r, 1) and row < mergedCells(r, 1) + mergedCells(r, 4) Then
				hiddenRow = True
			End If
			
			If row >= mergedCells(r, 1) and row < mergedCells(r, 1) + mergedCells(r, 4) and _
				col > mergedCells(r, 2) and col < mergedCells(r, 2) + mergedCells(r, 3) Then
				hiddenCol = True
			End If
		Next r
		
		If hiddenRow Then
			curString = ":::"
			
		Else If hiddenCol Then
			curString = ""
			
			Else
				curString = aCell.String
				
			End If
		End If

		If hiddenCol Then
			leftBuffer = ""
			rightBuffer = ""
		Else
			cellHoriJustify = aCell.getPropertyValue("HoriJustify")
			
			Select Case cellHoriJustify
				Case com.sun.star.table.CellHoriJustify.STANDARD, com.sun.star.table.CellHoriJustify.LEFT
					leftBuffer = " "
					rightBuffer = "  "
				case com.sun.star.table.CellHoriJustify.CENTER
					leftBuffer = "  "
					rightBuffer = "  "
				Case com.sun.star.table.CellHoriJustify.RIGHT
					leftBuffer = "  "
					rightBuffer = " "
			End Select
		End If
		
		curTargetString = targetCell.String
		targetCell.SetString(curTargetString & leftBuffer & curString & rightBuffer)

		curTargetString = targetCell.String
		targetCell.SetString(curTargetString & divisor)
	Next col
Next row

End Sub




'--------------------------------------------------------------------------------------------------'
' MergedCellGetRowSpanCount                                                                        '
'--------------------------------------------------------------------------------------------------'
' Returns a merged cell row span count.                                                            '
'                                                                                                  '
' Parameters:                                                                                      '
'                                                                                                  '
'   Cell As Variant                                                                                '
'     Reference to a range or cell (com.sun.star.table.XCellRange / com.sun.star.table.XCell) or a '
'     string name ("B5","R1C1", etc).                                                              '
'                                                                                                  '
'   Optional FailIfNotMerged As Boolean <Default = FALSE>                                          '
'     If set to TRUE, the function will return -1 if given Cell is not merged.                     '
'                                                                                                  '
' Examples:                                                                                        '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     rowspan = MergedCellGetRowSpanCount("B5")                                                    '
'     rowspan = MergedCellGetRowSpanCount("R1C1")                                                  '
'     rowspan = MergedCellGetRowSpanCount("$'Sheet.name.with.dots'.$G$9")                          '
'     rowspan = MergedCellGetRowSpanCount(ThisComponent.getCurrentSelection())                     '
' or                                                                                               '
'     cell = ThisComponent.Sheets.getByIndex(0).getCellByPosition(6,4)                             '
'     rowspan = MergedCellGetRowSpanCount(cell)                                                    '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     rowspan = MergedCellGetRowSpanCount("B5",TRUE)                                               '
'                                                                                                  '
' Will return -1 if B5 is not a merged cell.                                                       '
'--------------------------------------------------------------------------------------------------'
' Feedback & Issues:                                                                               '
'   https://github.com/aa6/libreoffice_calc_basic_extras/issues                                    '
'--------------------------------------------------------------------------------------------------'
Function MergedCellGetRowSpanCount(Cell As Variant, Optional FailIfNotMerged As Boolean)
    
    Dim args1(0) As New com.sun.star.beans.PropertyValue    
    Dim args2(1) As New com.sun.star.beans.PropertyValue    
    Dim args3(0) As New com.sun.star.beans.PropertyValue    
    Dim dispatcher As Object
    Dim target_cell As Object
    Dim previous_selection As Object

    previous_selection = ThisComponent.getCurrentSelection()
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    
    args1(0).Name = "ToPoint"
    If TypeName(Cell) = "String" Then
        args1(0).Value = Cell
    Else 
        args1(0).Value = IIf(TRUE,Cell,Cell).AbsoluteName ' `Object variable not set.` workaround. '
    End If
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:GoToCell", "", 0, args1())
    target_cell = ThisComponent.getCurrentSelection()
    
    If FailIfNotMerged = TRUE AND NOT target_cell.IsMerged Then
        MergedCellGetRowSpanCount = -1
    Else
        args2(0).Name = "By"
        args2(0).Value = 1
        args2(1).Name = "Sel"
        args2(1).Value = false
        dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:GoDown", "", 0, args2())
        MergedCellGetRowSpanCount = ThisComponent.getCurrentSelection().CellAddress.Row - target_cell.CellAddress.Row
    End If
        
    args3(0).Name = "ToPoint"
    args3(0).Value = previous_selection.AbsoluteName
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:GoToCell", "", 0, args3())
  
End Function

'--------------------------------------------------------------------------------------------------'
' MergedCellGetColSpanCount                                                                        '
'--------------------------------------------------------------------------------------------------'
' Returns a merged cell column span count.                                                         '
'                                                                                                  '
' Parameters:                                                                                      '
'                                                                                                  '
'   Cell As Variant                                                                                '
'     Reference to a range or cell (com.sun.star.table.XCellRange / com.sun.star.table.XCell) or a '
'     string name ("B5","R1C1", etc).                                                              '
'                                                                                                  '
'   Optional FailIfNotMerged As Boolean <Default = FALSE>                                          '
'     If set to TRUE, the function will return -1 if given Cell is not merged.                     '
'                                                                                                  '
' Examples:                                                                                        '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     colspan = MergedCellGetColSpanCount("B5")                                                    '
'     colspan = MergedCellGetColSpanCount("R1C1")                                                  '
'     colspan = MergedCellGetColSpanCount("$'Sheet.name.with.dots'.$G$9")                          '
'     colspan = MergedCellGetColSpanCount(ThisComponent.getCurrentSelection())                     '
' or                                                                                               '
'     cell = ThisComponent.Sheets.getByIndex(0).getCellByPosition(6,4)                             '
'     colspan = MergedCellGetColSpanCount(cell)                                                    '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     colspan = MergedCellGetColSpanCount("B5",TRUE)                                               '
'                                                                                                  '
' Will return -1 if B5 is not a merged cell.                                                       '
'--------------------------------------------------------------------------------------------------'
' Feedback & Issues:                                                                               '
'   https://github.com/aa6/libreoffice_calc_basic_extras/issues                                    '
'--------------------------------------------------------------------------------------------------'
Function MergedCellGetColSpanCount(Cell As Variant, Optional FailIfNotMerged As Boolean)
    
    Dim args1(0) As New com.sun.star.beans.PropertyValue    
    Dim args2(1) As New com.sun.star.beans.PropertyValue    
    Dim args3(0) As New com.sun.star.beans.PropertyValue    
    Dim dispatcher As Object
    Dim target_cell As Object
    Dim previous_selection As Object

    previous_selection = ThisComponent.getCurrentSelection()
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    
    args1(0).Name = "ToPoint"
    If TypeName(Cell) = "String" Then
        args1(0).Value = Cell
    Else 
        args1(0).Value = IIf(TRUE,Cell,Cell).AbsoluteName ' `Object variable not set.` workaround. '
    End If
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:GoToCell", "", 0, args1())
    target_cell = ThisComponent.getCurrentSelection()
    
    If FailIfNotMerged = TRUE AND NOT target_cell.IsMerged Then
        MergedCellGetColSpanCount = -1
    Else
        args2(0).Name = "By"
        args2(0).Value = 1
        args2(1).Name = "Sel"
        args2(1).Value = false
        dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:GoRight", "", 0, args2())
        MergedCellGetColSpanCount = ThisComponent.getCurrentSelection().CellAddress.Column - target_cell.CellAddress.Column
    End If
        
    args3(0).Name = "ToPoint"
    args3(0).Value = previous_selection.AbsoluteName
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:GoToCell", "", 0, args3())
  
End Function
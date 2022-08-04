Attribute VB_Name = "HelperFunctionsV2"
'
' LOCK ALL SHEETS
'
' Used to automatically lock all worksheets before sending out.
' UserInterfaceOnly=True allows VBA code to change the sheet.
'
Sub lockAllSheets()
  Dim mSheet As Worksheet
  For Each mSheet In ThisWorkbook.Sheets
    mSheet.Protect Password:="", UserInterfaceOnly:=True
  Next mSheet
End Sub


'
' UNLOCK ALL SHEETS
'
' Used to automatically unlock all sheets for editing.
'
Sub unlockAllSheets()
  Dim mSheet As Worksheet
  
  For Each mSheet In ThisWorkbook.Sheets
    mSheet.Unprotect
  Next mSheet
End Sub


'
' UNCHECK ALL RADIO BUTTONS
'
' Used to clear all radio buttons in all worksheets.
'
Sub uncheckAllRadioButtons()

  Application.ScreenUpdating = False
  On Error Resume Next
  
  Dim mSheet As Worksheet
  
  For Each mSheet In ActiveWorkbook.Sheets
    mSheet.Activate
    ActiveSheet.OptionButtons = False
  Next mSheet
  
  Application.ScreenUpdating = True

End Sub


'
' SHOW/HIDE RADIO GROUP BORDERS
'
' Each group of radio buttons is contained in a group box
' which is usually hidden for presentation purposes.
' Use this subroutine to show the boxes for editing/troubleshooting
' or hide the boxes prior to publishing/sending to a client.
'
Sub showHideRadioGroupBorders()

  Dim myGB As GroupBox

  For Each myGB In ActiveSheet.GroupBoxes
    myGB.Visible = False
  Next myGB
  
End Sub


'
' SET PRINTER ATTRIBUTES
'
' For consistency of print settings across forms, use this
' subroutine to automatically apply print settings rather
' than manually editing them for each workbook.
'
Sub setPrinterAttributes()

  Dim mSheet As Worksheet
  Application.PrintCommunication = False
  
  For Each mSheet In ActiveWorkbook.Worksheets
    mSheet.PageSetup.BlackAndWhite = True
    mSheet.PageSetup.FitToPagesWide = 1
    mSheet.PageSetup.FitToPagesTall = False
    mSheet.PageSetup.PaperSize = xlPaperA4
    mSheet.PageSetup.TopMargin = 54
    mSheet.PageSetup.BottomMargin = 54
    mSheet.PageSetup.LeftMargin = 18
    mSheet.PageSetup.RightMargin = 18
  Next mSheet
  
  Application.PrintCommunication = True
  
End Sub


'
' RANGE TO CELL
'
' Used to display table data capture in single cell.
' eg. Table with columns A, B, C containing values 1, 2, 3
' becomes {[A, B, C],[1, 2, 3]} which can be converted back
' later as required.
' This is used in the "Data" worksheet.
'
Public Function rangeToCell(mRange As Range) As String
  
  Dim mText As String
  
  If mRange.Columns.Count > 1 Or mRange.Rows.Count > 1 Then
    mText = "{["
    For mRow = 1 To mRange.Rows.Count
      For mCol = 1 To mRange.Columns.Count
        mText = mText + mRange(mRow, mCol).Text + ", "
      Next mCol
      mText = mText + "],["
    Next mRow

    mText = mText + "]}"
  Else
    mText = mRange.Value
  End If
  
  rangeToCell = mText

End Function


'
' CREATE DATA SHEET
'
' Select a blank worksheet then run this subroutine to
' create a list of all named ranges in this workbook along
' with their values and sheet/cell locations.
'
Sub CreateDataSheet()

  Application.ScreenUpdating = False
  
  ActiveSheet.Name = "Data"

  Cells.Select
  Selection.Font.Name = "Arial"
  Selection.Font.Size = 9

  Rows("1:1").Select
  Selection.Font.Bold = True
  
  Dim startCell As Range
  
  Dim cell As Range
  Dim targetCell As Range
  Dim mText As String
  
  Cells(1, 1).Select
  
  ActiveCell.Value = "Field"
  ActiveCell.ColumnWidth = 49.29
  
  ActiveCell.Offset(0, 1).Value = "Value"
  ActiveCell.Offset(0, 1).ColumnWidth = 99.29
  
  ActiveCell.Offset(0, 2).Value = "Sheet"
  ActiveCell.Offset(0, 2).ColumnWidth = 10.71
  
  ActiveCell.Offset(0, 3).Value = "Row"
  ActiveCell.Offset(0, 3).ColumnWidth = 10.71
  
  ActiveCell.Offset(0, 4).Value = "Column"
  ActiveCell.Offset(0, 4).ColumnWidth = 10.71
  
  ActiveCell.Offset(0, 5).Value = "Address"
  ActiveCell.Offset(0, 5).ColumnWidth = 10.71
  
  ActiveCell.Offset(1, 0).Select
  
  Dim mName As Name

  For Each mName In ThisWorkbook.Names
    If mName.Value <> "=#NAME?" And Right(mName.Name, 11) <> "!Print_Area" Then
    
      ActiveCell.Value = mName.Name
      ActiveCell.Offset(0, 1).Value = "=rangeToCell(" & Right(mName.Value, Len(mName.Value) - 1) & ")"
      ActiveCell.Offset(0, 2).Value = Trim(Replace(Replace(Replace(Replace(Left(mName.RefersTo, InStr(mName.RefersTo, "!")), "='", ""), "'!", ""), "=", ""), "!", ""))
      ActiveCell.Offset(0, 3).Value = mName.RefersToRange.Row
      ActiveCell.Offset(0, 4).Value = mName.RefersToRange.Column
      ActiveCell.Offset(0, 5).Value = mName.RefersToRange.Address
        
      ActiveCell.Offset(1, 0).Select
    End If
  Next mName
  
  Cells(1, 1).Select
  
  Application.ScreenUpdating = True
  
End Sub

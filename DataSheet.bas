Attribute VB_Name = "DataSheet"
'
' CREATE DATA SHEET
'
' REQUIRES: TableToText.bas for RangeToCellDelimited function
'
' Select a blank worksheet then run this subroutine to
' create a list of all named ranges in this workbook along
' with their values and sheet/cell locations.
'
Sub CreateDataSheet()

  Application.ScreenUpdating = False
  Application.DisplayAlerts = False
  
  If SheetExists("Data") Then
    Sheets("Data").Delete
  End If
  
  ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
  ActiveWorkbook.Sheets(ActiveWorkbook.Worksheets.Count).Name = "Data"
  ActiveWorkbook.Sheets("Data").Select


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
  
  ActiveCell.Value = "Form"
  ActiveCell.ColumnWidth = 15
  
  ActiveCell.Offset(0, 1).Value = "Field"
  ActiveCell.Offset(0, 1).ColumnWidth = 52
  
  ActiveCell.Offset(0, 2).Value = "Value"
  ActiveCell.Offset(0, 2).ColumnWidth = 100
  
  ActiveCell.Offset(0, 3).Value = "Sheet"
  ActiveCell.Offset(0, 3).ColumnWidth = 12
  
  ActiveCell.Offset(0, 4).Value = "Row"
  ActiveCell.Offset(0, 4).ColumnWidth = 12
  
  ActiveCell.Offset(0, 5).Value = "Column"
  ActiveCell.Offset(0, 5).ColumnWidth = 12
  
  ActiveCell.Offset(0, 6).Value = "Address"
  ActiveCell.Offset(0, 6).ColumnWidth = 12
  
  ActiveCell.Offset(1, 0).Select
  
  Dim mName As Name

  For Each mName In ThisWorkbook.Names
    If mName.Value <> "=#NAME?" And Right(mName.Name, 11) <> "!Print_Area" Then
    
      ActiveCell.Value = "=MID(CELL(""filename""),SEARCH(""["",CELL(""filename""))+1, SEARCH(""]"",CELL(""filename""))-SEARCH(""["",CELL(""filename""))-1)"
      ActiveCell.Offset(0, 1).Value = mName.Name
      ActiveCell.Offset(0, 2).Value = "=rangeToCellDelimited(" & Right(mName.Value, Len(mName.Value) - 1) & ")"
      ActiveCell.Offset(0, 3).Value = Trim(Replace(Replace(Replace(Replace(Left(mName.RefersTo, InStr(mName.RefersTo, "!")), "='", ""), "'!", ""), "=", ""), "!", ""))
      ActiveCell.Offset(0, 4).Value = mName.RefersToRange.Row
      ActiveCell.Offset(0, 5).Value = mName.RefersToRange.Column
      ActiveCell.Offset(0, 6).Value = mName.RefersToRange.Address
        
      ActiveCell.Offset(1, 0).Select
    End If
  Next mName
  
  Cells(1, 1).Select
  
  ActiveWindow.SelectedSheets.Visible = False
  Application.ScreenUpdating = True
  Application.DisplayAlerts = True
  ActiveWorkbook.Sheets(1).Select
End Sub


'
' SHEET EXISTS
'
Private Function SheetExists(SheetName As String, Optional wb As Excel.Workbook)
  Dim s As Excel.Worksheet
  
  If wb Is Nothing Then Set wb = ThisWorkbook
  
  On Error Resume Next
  Set s = wb.Sheets(SheetName)
  On Error GoTo 0
  
  SheetExists = Not s Is Nothing
End Function


'
' RANGE TO CELL DELIMITED
'
' Used to display table data capture in single cell.
' eg. Table with columns A, B, C containing values 1, 2, 3
' becomes {[A, B, C],[1, 2, 3]} which can be converted back
' later as required.
' This is used in the "Data" worksheet.
'
Public Function rangeToCellDelimited(mRange As Range) As String
  
  Dim mRow As Integer
  Dim mCol As Integer
  
  Dim mText As String
  Dim lastAddress As String
  lastAddress = ""
  
  If mRange.Columns.Count > 1 Or mRange.Rows.Count > 1 Then
    mText = "{"
    For mRow = 1 To mRange.Rows.Count
      mText = mText + "["
      For mCol = 1 To mRange.Columns.Count
        If mRange(mRow, mCol).MergeArea.Address <> lastAddress Then
          mText = mText + mRange(mRow, mCol).Text + ", "
          lastAddress = mRange(mRow, mCol).MergeArea.Address
        End If
      Next mCol
      mText = Left(mText, Len(mText) - 2) ' remove last ", "
      mText = mText + "]," + Chr(10)
    Next mRow

    mText = Left(mText, Len(mText) - 3)
    mText = mText + "]}"
  Else
    mText = mRange.Value
  End If
  
  rangeToCellDelimited = mText

End Function



'
' RANGE TO CELL TEXT
'
' Used to display table data capture in single cell.
' eg. Table with columns A, B, C containing values 1, 2, 3
' becomes {[A, B, C],[1, 2, 3]} which can be converted back
' later as required.
' This is used in the "Data" worksheet.
'
Public Function rangeToCellText(mRange As Range, Optional UnderlineHeadings As String = "", Optional CapitaliseHeadings As Boolean = False) As String
  
  Dim extraSpace As Integer
  extraSpace = 4
  
  Dim mRow As Integer
  Dim mCol As Integer
  
  Dim mText As String
  Dim lastAddress As String
  lastAddress = ""
  
  Dim colCnt As Integer
  Dim i As Integer
  
  Dim mContent As String
  
  If mRange.Columns.Count > 1 Or mRange.Rows.Count > 1 Then
  
    Dim maxColWidths() As Integer
    ReDim maxColWidths(1)
    
    For mRow = 1 To mRange.Rows.Count
      colCnt = 0
      For mCol = 1 To mRange.Columns.Count
        If mRange(mRow, mCol).MergeArea.Address <> lastAddress Then
          If colCnt > UBound(maxColWidths) Then ReDim Preserve maxColWidths(colCnt)
          If Len(mRange(mRow, mCol).Text) > maxColWidths(colCnt) Then maxColWidths(colCnt) = Len(mRange(mRow, mCol).Text)
          lastAddress = mRange(mRow, mCol).MergeArea.Address
          colCnt = colCnt + 1
        End If
      Next mCol
    Next mRow
    lastAddress = ""

  
    ' START generate mText
    For mRow = 1 To mRange.Rows.Count
      colCnt = 0
      For mCol = 1 To mRange.Columns.Count
      
        If mRange(mRow, mCol).MergeArea.Address <> lastAddress Then
          mContent = mRange(mRow, mCol).Text
          If CapitaliseHeadings Then mContent = UCase(mContent)
          mText = mText + Left(mContent + _
                  Space(maxColWidths(colCnt) + extraSpace), maxColWidths(colCnt) + extraSpace)
          lastAddress = mRange(mRow, mCol).MergeArea.Address
          colCnt = colCnt + 1
        End If
      Next mCol
      
      mText = mText + Chr(10)
      
      ' add ----- under table headings
      If UnderlineHeadings <> "" And mRow = 1 Then
        For i = LBound(maxColWidths) To UBound(maxColWidths)
          mText = mText + String(maxColWidths(i) + extraSpace - 2, UnderlineHeadings) + Space(2)
        Next i
        mText = mText + Chr(10)
      End If
      
    Next mRow
    ' END generate mText
    
  Else
    mText = mRange.Value
  End If
  
  rangeToCellText = mText
  
End Function


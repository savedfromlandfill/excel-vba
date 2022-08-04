Attribute VB_Name = "Utils1"
'
'
'
Sub ProperCaseAllSheetNames()
Dim sh As Worksheet

For Each sh In ActiveWorkbook.Worksheets
    sh.Name = StrConv(sh.Name, vbProperCase)
Next sh
End Sub

'
'
'
Sub CreateLinksToAllSheets()
Dim sh As Worksheet
Dim cell As Range
For Each sh In ActiveWorkbook.Worksheets
    If ActiveSheet.Name <> sh.Name Then
        ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        "'" & sh.Name & "'" & "!A1", TextToDisplay:=sh.Name
        ActiveCell.Offset(1, 0).Select
    End If
Next sh
End Sub

'
'
'
Sub SplitLinebreaksToSeperateCells()

  Application.ScreenUpdating = False
  
  Dim cell As Range
  Dim mText As String
  
  mText = ActiveCell.Value
  ActiveCell.Offset(1, 0).Select
  
  For i = 1 To Len(mText)

    mChar = Mid(mText, i, 1)
      
    If mChar = Chr(10) Then
      ActiveCell.Offset(1, 0).Select
    Else
      ActiveCell.Value = ActiveCell.Value & mChar
    End If
  
  Next i
  
  Application.ScreenUpdating = True

End Sub

'
'
'
Sub AddExtraLinebreakToCells()

  Application.ScreenUpdating = False
  
  Dim cell As Range
  Dim mText As String
  
  Do While ActiveCell.Value <> ""

    If Right(mText, 1) <> Chr(10) Then ActiveCell.Value = ActiveCell.Value & Chr(10)
    
    ActiveCell.Offset(1, 0).Select

  Loop
  
  Application.ScreenUpdating = True

End Sub

'
'
'
Public Function colWidth(mRange As Range) As Single

    Dim rngColumn As Range
    Dim colWidthCounter As Single
    
    colWidthCounter = 0
    
    For Each rngColumn In mRange.Columns
      colWidthCounter = colWidthCounter + rngColumn.ColumnWidth
    Next rngColumn
    
    colWidth = colWidthCounter
    
End Function

'
'
'
Public Function rowHeight(mRange As Range) As Single

    Dim rngRow As Range
    Dim rowHeightCounter As Single
    
    rowHeightCounter = 0
    
    For Each rngRow In mRange.Rows
      rowHeightCounter = rowHeightCounter + rngRow.rowHeight
    Next rngRow
    
    rowHeight = rowHeightCounter
    
End Function

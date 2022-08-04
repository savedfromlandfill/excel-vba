Attribute VB_Name = "TableToText"
Option Explicit

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


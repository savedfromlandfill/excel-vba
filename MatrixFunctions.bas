Attribute VB_Name = "MatrixFunctions"
'
'
'
Sub radioButton_Click()

Dim mShape As Shape
Dim mRange As Range
Dim mRow As Integer

If TypeName(Application.Caller) = "String" Then

  Set mShape = ActiveSheet.Shapes(Application.Caller)
  Set mRange = Range(mShape.TopLeftCell.Address, mShape.TopLeftCell.Address)
  
  mRow = Range(mRange.MergeArea.Cells(1, 1), mRange.MergeArea.Cells(1, 1)).Row
  
  ActiveSheet.Cells(mRow, 7) = mShape.TextFrame.Characters.Text

  If Worksheets("Matrix").Shapes("chkAutoscrollNextOutcome").OLEFormat.Object.Value > 0 Then
    ActiveWindow.ScrollRow = mRow + mRange.MergeArea.Rows.Count
  End If
  
End If

End Sub

Sub btnScrollUp_Click()
  ActiveWindow.SmallScroll Up:=1
End Sub

Sub btnScrollDown_Click()
  ActiveWindow.SmallScroll Down:=1
End Sub

Sub btnScrollTop_Click()
  ActiveWindow.ScrollRow = 3
End Sub

Sub btnClearAnswers_Click()
  If MsgBox("Clear all scores?", vbYesNo) = vbNo Then Exit Sub
  ActiveSheet.Unprotect
  ActiveSheet.OptionButtons = False
  
  Dim i As Integer
  For i = 4 To 32
    If ActiveSheet.Cells(i, 7).Interior.Color <> 14277081 Then
      ActiveSheet.Cells(i, 7).Value = 0
    End If
  Next i
  
  ActiveSheet.Protect Password:="", UserInterfaceOnly:=True
End Sub

'
' Used for the radio buttons on the Details Part 2 worksheet
'
Sub radioButtonInvestmentReady_Click()

  Dim mShape As Shape
  Dim mRange As Range
  Dim mRow As Integer

  If TypeName(Application.Caller) = "String" Then
    'ActiveSheet.Unprotect
    Set mShape = ActiveSheet.Shapes(Application.Caller)
    Set mRange = Range(mShape.TopLeftCell.Address, mShape.TopLeftCell.Address)
    mRow = Range(mRange.MergeArea.Cells(1, 1), mRange.MergeArea.Cells(1, 1)).Row
    
    Range(ActiveSheet.Cells(mRow, 4), ActiveSheet.Cells(mRow, 6)).Interior.Color = 14277081 ' light gray
    Range(mShape.TopLeftCell.Address).Interior.Color = 11854022 ' light green
    ActiveSheet.Cells(mShape.TopLeftCell.Row, 7).Value = mShape.TopLeftCell.Column - 2
    'ActiveSheet.Protect Password:="", UserInterfaceOnly:=True
  End If

End Sub

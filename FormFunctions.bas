Attribute VB_Name = "FormFunctions"
Sub lockAllSheets()
  Dim mSheet As Worksheet
  Dim selectedRange As Range
  Set selectedRange = ActiveCell
  For Each mSheet In ThisWorkbook.Sheets
    mSheet.EnableSelection = xlUnlockedCells
    mSheet.Protect Password:="", UserInterfaceOnly:=True, AllowFormattingRows:=True
  Next mSheet
  selectedRange.Select
End Sub

Sub ReferenceNumberInHeader()
  Dim mSheet As Worksheet
  For Each mSheet In ThisWorkbook.Sheets
    mSheet.PageSetup.RightHeader = ThisWorkbook.Names("form_reference_number").RefersToRange.Text
  Next mSheet
End Sub





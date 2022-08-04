Attribute VB_Name = "FindReplace"
Sub unlockAllSheets()
  Dim mSheet As Worksheet
  For Each mSheet In ThisWorkbook.Sheets
    mSheet.Unprotect
  Next mSheet
End Sub


Sub FindReplace_AllSheets()

  Application.ScreenUpdating = False
  
  Dim sheet As Worksheet
  Dim findList As Variant
  Dim rplcList As Variant
  Dim i As Integer

  findList = Array("search1", "search2", "search3", "search4", "search5")
  rplcList = Array("replace1", "replace2", "replace3", "replace4", "replace5")

  For i = LBound(findList) To UBound(findList)
    For Each sheet In ActiveWorkbook.Worksheets
      sheet.Cells.Replace What:=Trim(findList(i)), _
                   Replacement:=Trim(rplcList(i)), _
                        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, _
                  SearchFormat:=False, ReplaceFormat:=False
    Next sheet
  Next i

  Application.ScreenUpdating = True

End Sub

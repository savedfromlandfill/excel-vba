Attribute VB_Name = "Utils2"
'
'
'
Sub lockAllSheets()
  Dim mSheet As Worksheet
  For Each mSheet In ThisWorkbook.Sheets
    mSheet.Protect Password:="", UserInterfaceOnly:=True
  Next mSheet
End Sub

'
'
'
Sub unlockAllSheets()
  Dim mSheet As Worksheet
  
  For Each mSheet In ThisWorkbook.Sheets
    mSheet.Unprotect
  Next mSheet
End Sub


'
'
'
Sub UncheckAllRadioButtons()

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
'
'
Sub ShowHideRadioGroupBorders()

  Dim myGB As GroupBox

  For Each myGB In ActiveSheet.GroupBoxes
    myGB.Visible = False ' Not myGB.Visible
    'myGB.Height = 44
  Next myGB
  
End Sub

Sub FixRadioButtonHeights()

  Dim myRDB As OptionButton
  
  For Each myRDB In ActiveSheet.OptionButtons
    myRDB.Height = 16
  Next myRDB
  
End Sub

Sub FixCheckBoxHeights()

  Dim myChk As CheckBox
  
  For Each myChk In ActiveSheet.Checkboxes
    myChk.Height = 16
  Next myChk

End Sub


'
'
'
Sub SetPrintAttributes()

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


Attribute VB_Name = "RadioButtonsForMatrixV1"
Option Explicit


'
' select starting cell with text and run this sub-routine
' checkboxes will be added to each neighbouring column
Sub radioButtonsGenerator()

  Application.ScreenUpdating = False
  ActiveSheet.Unprotect
  
  ' clear existing controls
  Dim mShape As Shape
  For Each mShape In ActiveSheet.Shapes
    mShape.Delete
  Next mShape
  
  ' ADD COMPONENT: autoscroll option checkbox
  ActiveSheet.Checkboxes.Add(ActiveSheet.Cells(2, 1).Left, ActiveSheet.Cells(3, 1).Top - 16, 128, 16).Select
  Selection.Characters.Text = "Autoscroll to next outcome"
  Selection.Name = "chkAutoscrollNextOutcome"
  Selection.Value = xlOn
  Selection.PrintObject = False
  
  ' ADD COMPONENT: previous button
  ActiveSheet.Buttons.Add(ActiveSheet.Cells(2, 1).Left + 4, ActiveSheet.Cells(3, 1).Top - 32, 24.5, 16).Select
  Selection.Characters.Text = "3"
  Selection.Font.Name = "Webdings"
  Selection.Name = "btnPreviousRow"
  Selection.OnAction = "btnScrollUp_Click"
  Selection.PrintObject = False
        
  ' ADD COMPONENT: next button
  ActiveSheet.Buttons.Add(ActiveSheet.Cells(2, 1).Left + 32, ActiveSheet.Cells(3, 1).Top - 32, 24.5, 16).Select
  Selection.Characters.Text = "4"
  Selection.Font.Name = "Webdings"
  Selection.Name = "btnNextRow"
  Selection.OnAction = "btnScrollDown_Click"
  Selection.PrintObject = False
  
  ' ADD COMPONENT: top button
  ActiveSheet.Buttons.Add(ActiveSheet.Cells(2, 1).Left + 59.5, ActiveSheet.Cells(3, 1).Top - 32, 24.5, 16).Select
  Selection.Characters.Text = "5"
  Selection.Font.Name = "Webdings"
  Selection.Name = "btnScrollTop"
  Selection.OnAction = "btnScrollTop_Click"
  Selection.PrintObject = False
  
  ' ADD COMPONENT: clear button
  ActiveSheet.Buttons.Add(ActiveSheet.Cells(2, 1).Left + 87.5, ActiveSheet.Cells(3, 1).Top - 32, 24.5, 16).Select
  Selection.Characters.Text = "x"
  Selection.Font.Name = "Webdings"
  Selection.Name = "btnClearAnswers"
  Selection.OnAction = "btnClearAnswers_Click"
  Selection.PrintObject = False

  Dim mStartCell As Range ' used to return to selected cell after script completes
  Set mStartCell = ActiveCell
  ActiveSheet.Cells(1, 1).Select ' move to first row/column of sheet
  
  Dim mTargetCell As Range ' start target cell
  Dim mTargetRange As Range ' range of cells the width the group box is to fill
  
  Dim mFieldname As String

  ' control params
  Dim mLeft As Double
  Dim mTop As Double
  Dim mHeight As Double
  Dim mWidth As Double
  
  Dim mBoxHeight
  mBoxHeight = 32
  
  Dim mActiveSheet As Worksheet
  Set mActiveSheet = ActiveSheet

  ' MAIN LOOP
  Do While ActiveCell.Value <> ""
    If ActiveCell.Interior.ThemeColor = xlColorIndexNone Or _
       ActiveCell.Interior.Color = 16777215 Then ' 16777215 is white, if using colorindex=2 instead lightgray is also 2
    
      'ActiveCell.Rows.AutoFit
      'ActiveCell.RowHeight = ActiveCell.RowHeight + mBoxHeight + 4

      mFieldname = Replace(ActiveCell.Value, " ", "_")
      Set mTargetCell = ActiveCell.Offset(0, 1)
      Set mTargetRange = Range(mTargetCell, mTargetCell.Offset(mTargetCell.MergeArea.Rows.Count, 4))
    
    
      mLeft = mTargetCell.Left + 1
      mHeight = mBoxHeight
      mTop = mTargetCell.Offset(1, 0).Top - mBoxHeight - 1
      mWidth = mTargetRange.Width - 2
      
      
      mActiveSheet.GroupBoxes.Add(mLeft, mTop, mWidth, mHeight).Select
      Selection.Characters.Text = ""
      Selection.Visible = False
      
      Dim ctRadio As Integer
      
      Dim mRadioRange As Range
      
      ' Add the radio buttons
      
      ' OPTION 1: EVEN SPLIT
      Dim numOfButtons As Integer
      numOfButtons = 5
      If True = True Then
        For ctRadio = 1 To numOfButtons
          ActiveSheet.OptionButtons.Add(mLeft + 44 + (ctRadio - 1) * mWidth / numOfButtons, mTop + 1, mBoxHeight - 2, mBoxHeight - 2).Select
          Selection.Characters.Text = CStr(ctRadio)
          'Selection.LinkedCell = mTargetCell.Offset(0, 5).Address
          Selection.OnAction = "radioButton_Click"
          If Selection.TopLeftCell.Interior.Color = 15921906 Then Selection.Delete
        Next ctRadio
      End If
      
      ' OPTION 2: AS PER COLUMNS
      If True = False Then
        Set mRadioRange = Range(mTargetCell, mTargetCell.Offset(0, 1))
        For ctRadio = 1 To 4
          ActiveSheet.OptionButtons.Add(mRadioRange.Left + 3 + (ctRadio - 1) * mRadioRange.Width / 4, mTop + 1, mBoxHeight - 2, mBoxHeight - 2).Select
          Selection.Characters.Text = CStr(ctRadio)
          'Selection.LinkedCell = mTargetCell.Offset(0, 5).Address
          Selection.OnAction = "radioButton_Click"
        Next ctRadio
        
        Set mRadioRange = Range(mTargetCell.Offset(0, 2), mTargetCell.Offset(0, 2))
        For ctRadio = 5 To 7
          ActiveSheet.OptionButtons.Add(mRadioRange.Left + 3 + (ctRadio - 5) * mRadioRange.Width / 3, mTop + 1, mBoxHeight - 2, mBoxHeight - 2).Select
          Selection.Characters.Text = CStr(ctRadio)
          Selection.LinkedCell = mTargetCell.Offset(0, 5).Address
          Selection.OnAction = "radioButton_Click"
        Next ctRadio
        
        Set mRadioRange = Range(mTargetCell.Offset(0, 3), mTargetCell.Offset(0, 3))
        For ctRadio = 8 To 9
          ActiveSheet.OptionButtons.Add(mRadioRange.Left + 3 + (ctRadio - 8) * mRadioRange.Width / 2, mTop + 1, mBoxHeight - 2, mBoxHeight - 2).Select
          Selection.Characters.Text = CStr(ctRadio)
          Selection.LinkedCell = mTargetCell.Offset(0, 5).Address
          Selection.OnAction = "radioButton_Click"
        Next ctRadio
        
        Set mRadioRange = Range(mTargetCell.Offset(0, 4), mTargetCell.Offset(0, 4))
        ctRadio = 10
        ActiveSheet.OptionButtons.Add(mRadioRange.Left + 1, mTop + 1, mBoxHeight - 2, mBoxHeight - 2).Select
        Selection.Characters.Text = CStr(ctRadio)
        Selection.LinkedCell = mTargetCell.Offset(0, 5).Address
        Selection.OnAction = "radioButton_Click"
      End If
      
      ' END OF add the radio buttons
    
    
    End If ' checking cell color to skip heading rows
    
    ActiveCell.Offset(ActiveCell.Offset(0, 1).MergeArea.Rows.Count, 0).Select ' move to next row
    
  Loop
  ' END MAIN LOOP
  
  
  ' return selection to user selected cell
  mStartCell.Select
  
  Application.ScreenUpdating = True
  
End Sub


Sub deleteAllShapes()

  Dim mShape As Shape
  For Each mShape In ActiveSheet.Shapes
    mShape.Delete
  Next mShape
  
End Sub


'
' Used to add radio buttons to Assessment Matrix, Details Part 2
'
Sub simpleRadioButtonsGenerator()

  Application.ScreenUpdating = False
  ActiveSheet.Unprotect
  
  
  ' clear existing controls
  'Dim mShape As Shape
  'For Each mShape In ActiveSheet.Shapes
  '  mShape.Delete
  'Next mShape


  Dim mStartCell As Range ' used to return to selected cell after script completes
  Set mStartCell = ActiveCell
 
  Dim mTargetCell As Range ' start target cell
  Dim mTargetRange As Range ' range of cells the width the group box is to fill
  
  Dim mFieldname As String

  ' control params
  Dim mLeft As Double
  Dim mTop As Double
  Dim mHeight As Double
  Dim mWidth As Double
  
  Dim mBoxHeight
  mBoxHeight = 32
  
  Dim mActiveSheet As Worksheet
  Set mActiveSheet = ActiveSheet
  
  ' create the group box for the radio buttons
  mActiveSheet.GroupBoxes.Add(ActiveCell.Left + 1, ActiveCell.Offset(1, 0).Top - 30, _
    Range(ActiveCell.Offset(0, 0), ActiveCell.Offset(0, 9)).Width - 2, 30).Select
  Selection.Characters.Text = ""
  Selection.Visible = False
  
  Dim rdoCount As Integer
  rdoCount = 1

Do While ActiveCell.Value <> "" ' multirow radio generate
  ' MAIN LOOP
  Do While ActiveCell.Value <> ""
    
      mLeft = ActiveCell.Left + 45
      mHeight = 16
      mTop = ActiveCell.Offset(1, 0).Top - 22
      mWidth = 20
      
      
      Dim ctRadio As Integer
      
      Dim mRadioRange As Range
      
      ' Add the radio buttons
      
      ActiveSheet.OptionButtons.Add(mLeft, mTop, mWidth, mHeight).Select
      Selection.Characters.Text = ""
      Selection.OnAction = "radioButtonInvestmentReady_Click"
      Selection.Name = "chkReadiness" & Trim(Str(rdoCount))
      
      rdoCount = rdoCount + 1
      
      ' END OF add the radio buttons

    
    ActiveCell.Offset(0, 1).Select ' move to next col
    
  Loop
  ' END MAIN LOOP
  Cells(ActiveCell.Offset(1, 0).Row, mStartCell.Column).Select
Loop
  
  
  ' return selection to user selected cell
  mStartCell.Select
  
  Application.ScreenUpdating = True
  
End Sub


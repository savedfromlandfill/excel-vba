Attribute VB_Name = "CheckAndRadioGenerate"
Option Explicit



Sub Checkbox_RadioBox_Generate()
  
  Dim FieldCol As String
  Dim RadioCol As String
  Dim CheckCol As String
  Dim TotalRows As Integer
  TotalRows = 300
  
  ' SETTINGS EOI PART B
  'FieldCol = "N"
  'RadioCol = "K"
  'CheckCol = "C"
  ' END SETTINGS
  
  ' SETTINGS IR PART B & EOI PART A
  FieldCol = "P"
  RadioCol = "L"
  CheckCol = "L"
  ' END SETTINGS
  
   ' SETTINGS IR PART C and PART D and PART E
  'FieldCol = "P"
  'RadioCol = "L"
  'CheckCol = "C"
  ' END SETTINGS

  ' SETTINGS IR PART F
  'FieldCol = "L"
  'RadioCol = "D"
  'CheckCol = "D"
  ' END SETTINGS
  
  ' SETTINGS MATRIX Details Part 1
  FieldCol = "P"
  RadioCol = "L"
  CheckCol = "C"
  ' END SETTINGS
  

  Dim FieldColNum As Integer
  Dim RadioColNum As Integer
  Dim CheckColNum As Integer
  FieldColNum = Range(FieldCol & "1").Column
  RadioColNum = Range(RadioCol & "1").Column
  CheckColNum = Range(CheckCol & "1").Column
  
  Dim i As Integer
  Dim CellName1 As String
  Dim CellName2 As String
  
  Dim tmpRng As Range
  
  
  Application.ScreenUpdating = False
  ActiveSheet.Unprotect
  
  ' clear existing controls
  Dim mShape As Shape
  For Each mShape In ActiveSheet.Shapes
    mShape.Delete
  Next mShape
  
  ' main loop to create controls
  For i = 1 To TotalRows
    
    Cells(i, FieldColNum).Select
    
    CellName1 = ""
    CellName2 = ""
    On Error Resume Next
    CellName1 = ActiveCell.Name.Name
    CellName2 = ActiveCell.Offset(0, 1).Name.Name
    On Error GoTo 0
    

    If CellName1 <> "" Then
      ' named cell so we want to create a checkbox or radiobox
      
      
      If InStr(LCase(CellName1), "attach") Then
      
        ' attachment checkbox
        ActiveSheet.Checkboxes.Add(ActiveCell.Offset(0, CheckColNum - FieldColNum).Left + 28, _
                                   ActiveCell.Offset(0, CheckColNum - FieldColNum).Top - 2, _
                                   64, _
                                   16).Select
        Selection.Characters.Text = "Attached"
        Selection.LinkedCell = "$" + FieldCol + "$" + Trim(Str(i))
        
        If InStr(LCase(CellName2), "na") Then
        
          ' NA checkbox
          ActiveSheet.Checkboxes.Add(ActiveCell.Offset(0, CheckColNum - FieldColNum).Left + 64 + 28, _
                                     ActiveCell.Offset(0, CheckColNum - FieldColNum).Top - 2, _
                                     64, _
                                     16).Select
          Selection.Characters.Text = "N/A"
          Selection.LinkedCell = "$" + Chr(Asc(FieldCol) + 1) + "$" + Trim(Str(i))
          
        End If
        
      ElseIf InStr(CellName1, "_chk") Or InStr(CellName1, "_yes") And Not InStr(CellName1, "_yesno") Then
      
        ' other checkbox
        ActiveSheet.Checkboxes.Add(ActiveCell.Offset(0, CheckColNum - FieldColNum).Left + 28, _
                                   ActiveCell.Offset(0, CheckColNum - FieldColNum).Top - 2, _
                                   ActiveCell.Offset(0, CheckColNum - FieldColNum).Width, _
                                   16).Select
        Selection.Characters.Text = ""
        Selection.LinkedCell = "$" + FieldCol + "$" + Trim(Str(i))
                
      Else
        ' radiobox yes/no
      
        ' create the group box for the radio buttons
        ActiveSheet.GroupBoxes.Add(ActiveCell.Offset(0, RadioColNum - FieldColNum).Left, _
                                   ActiveCell.Offset(0, RadioColNum - FieldColNum).Top - 4, _
                                   72, _
                                   20).Select
        Selection.Characters.Text = ""
        Selection.Visible = False
        
        ' create yes option
        ActiveSheet.OptionButtons.Add(ActiveCell.Offset(0, RadioColNum - FieldColNum).Left + 1, _
                                      ActiveCell.Offset(0, RadioColNum - FieldColNum).Top - 2, _
                                      36, _
                                      16).Select
        Selection.Characters.Text = "Yes"
        Selection.LinkedCell = "$" + FieldCol + "$" + Trim(Str(i))
        
        ' create no option
        ActiveSheet.OptionButtons.Add(ActiveCell.Offset(0, RadioColNum - FieldColNum).Left + 36, _
                                      ActiveCell.Offset(0, RadioColNum - FieldColNum).Top - 2, _
                                      36, _
                                      16).Select
        Selection.Characters.Text = "No"
        Selection.LinkedCell = "$" + FieldCol + "$" + Trim(Str(i))
      
      End If ' attachment / radiobox
    
    
    ' checkbox linked to range in a different sheet
    ElseIf InStr(1, ActiveCell.Value, "!") Then
        
        ActiveSheet.Checkboxes.Add(ActiveCell.Offset(0, CheckColNum - FieldColNum).Left + 28, _
                                   ActiveCell.Offset(0, CheckColNum - FieldColNum).Top - 2, _
                                   64, _
                                   16).Select
        Selection.Characters.Text = "Attached"
        Selection.LinkedCell = "'" + ActiveCell.Text
        
        If InStr(1, ActiveCell.Offset(0, 1).Value, "!") Then
          ' 2nd checkbox
          ActiveSheet.Checkboxes.Add(ActiveCell.Offset(0, CheckColNum - FieldColNum).Left + 28 + 64, _
                                     ActiveCell.Offset(0, CheckColNum - FieldColNum).Top - 2, _
                                     64, _
                                     16).Select
          Selection.Characters.Text = "N/A"
          Selection.LinkedCell = "'" + ActiveCell.Offset(0, 1).Text
        End If
    
    End If ' namedrange / fieldname found
  
  Next i ' end main loop

  Application.ScreenUpdating = True
  
End Sub




Sub Get_Existing_Control_Details()

  Dim OptionBtn As OptionButton
  Dim Checkbox As Checkbox
  
  Dim UpdateCol As String
  
  Dim chkLinkedCell As String
  Dim chkLinkedSheet As String
  
  ' SETTINGS
  UpdateCol = "P"
  
  Dim UpdateColNum As Integer
  UpdateColNum = Range(UpdateCol & "1").Column


  Debug.Print "---START---"

  For Each OptionBtn In ActiveSheet.OptionButtons
    Debug.Print OptionBtn.Name + ": " + OptionBtn.LinkedCell
  Next OptionBtn

  For Each Checkbox In ActiveSheet.Checkboxes

    chkLinkedSheet = ""
    chkLinkedCell = ""
    If InStr(1, Checkbox.LinkedCell, "!", vbTextCompare) Then
      chkLinkedSheet = Replace(Split(Checkbox.LinkedCell, "!")(0), "'", "")
      chkLinkedCell = Split(Checkbox.LinkedCell, "!")(1)
    Else
      chkLinkedCell = Checkbox.LinkedCell
    End If
    
    If chkLinkedSheet = ActiveSheet.Name Then
      'Checkbox.LinkedCell = chkLinkedCell ' remove sheet reference if control is in same sheet as what it's referencing
    End If
    
    If chkLinkedSheet <> "" And chkLinkedSheet <> ActiveSheet.Name Then
      If ActiveSheet.Cells(Checkbox.TopLeftCell.Row, UpdateColNum).Value = "" Then
        ActiveSheet.Cells(Checkbox.TopLeftCell.Row, UpdateColNum).Value = Checkbox.LinkedCell
      Else
        ActiveSheet.Cells(Checkbox.TopLeftCell.Row, UpdateColNum + 1).Value = Checkbox.LinkedCell
      End If
    End If
    
    
    Debug.Print Checkbox.Name + ",  " + chkLinkedSheet + ",  LINK:" + chkLinkedCell + ", " + Checkbox.TopLeftCell.Address
    
    
    
  Next Checkbox
    
  Debug.Print "----END----"

End Sub



Sub Clear_Controls()

  Dim mShape As Shape
  For Each mShape In ActiveSheet.Shapes
    mShape.Delete
  Next mShape
  
End Sub

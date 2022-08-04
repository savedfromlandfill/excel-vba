Attribute VB_Name = "NameUtils"

' run to list all names in a blank worksheet
' delete all 'print area' names etc and other names you don't want to change
' names change be changed in another column using text transformation formulas, etc
' then run loopThroughNameChanges (below) to process name renaming operations
Sub listAllNames()

  Application.ScreenUpdating = False
  
  Dim startCell As Range
  
  Dim cell As Range
  Dim targetCell As Range
  Dim mText As String
  
  
  Set startCell = ActiveCell
  
  
  Dim mName As Name

  For Each mName In ThisWorkbook.Names
    ActiveCell.Value = mName.Name
      
    ActiveCell.Offset(1, 0).Select
  Next mName
  
  ' return selection to user selected cell
  startCell.Select
  
  Application.ScreenUpdating = True
  
End Sub

'
'
' After running list all names
Sub loopThroughNameChanges()

  Application.ScreenUpdating = False
  
  Dim startCell As Range
  
  Dim cell As Range
  Dim targetCell As Range
  Dim mText As String
  
  
  Set startCell = ActiveCell
  
  
  Do While ActiveCell.Value <> ""
    Debug.Print ActiveCell.Value & " rename to " & ActiveCell.Offset(0, 1).Value
    ThisWorkbook.Names(ActiveCell.Value).Name = ActiveCell.Offset(0, 1).Value
    ActiveCell.Offset(1, 0).Select
  Loop
  
  ' return selection to user selected cell
  startCell.Select
  
  Application.ScreenUpdating = True
  
End Sub

'
' CREATE DATA SHEET
'
' Select a blank worksheet then run this subroutine to
' create a list of all named ranges in this workbook along
' with their values and sheet/cell locations.
'
Sub Create_Data_Sheet()

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


'
' Used in cell to change CamelCase names to underscore_seperated
'
Public Function fixName(mName As String) As String
  
  Dim Counter As Integer
  Dim mChar As String
  Dim lastChar As String
  
  Dim newString As String

  For Counter = 1 To Len(mName)
    mChar = Mid(mName, Counter, 1)
    
    If mChar = UCase(mChar) Then
      If Counter > 1 And _
        lastChar <> UCase(lastChar) Then _
          newString = newString & "_" ' don't put underscore infront of first char
      newString = newString & LCase(mChar)
    Else
      newString = newString & mChar
    End If
    
    lastChar = mChar
    
  Next
  
  fixName = newString

End Function


'
' Create shortcut key for edit named range below.
'
Sub shortcut_editNamedRange()

  Application.OnKey "^+m", "editNamedRange" ' CTRL SHIFT M
  
End Sub


'
' EDIT NAMED RANGE
'
' Changing a name in Excel's Name Box adds another name to the
' named ranges collection. This subroutine prompts for a new
' name and replaces the existing name avoiding duplicate names
' for the same range.
'
Sub editNamedRange()

  On Error GoTo mError
  
  Dim result As String
  Dim currentName As String
  Dim textToAppend As String
  textToAppend = "" ' "_textbox" ' change this here as required
  
  currentName = ActiveCell.Name.Name ' name.name is the named range name if it exists
  result = InputBox("New Name:", "Change Name", currentName + textToAppend)
    
  If StrPtr(result) = 0 Then
    ' cancel
  ElseIf result = vbNullString Then
    ' empty string
  Else
    ThisWorkbook.Names(currentName).Name = result
  End If

Exit Sub

mError:   ' expect to land here when there's no name to edit.
          ' do nothing, exit subroutine.

End Sub


'
' CHANGE SCOPE ALL NAMES
'
Sub changeScopeAllNames()

  Dim mName As Name
  Dim mNameStr As String
  Dim mRefersTo As String
  
  Debug.Print "--- START changeScopeAllNames"

  For Each mName In ThisWorkbook.Names
    If InStr(1, mName.RefersTo, mName.Parent.Name, vbTextCompare) Then
      mNameStr = Right(mName.Name, Len(mName.Name) - InStr(1, mName.Name, "!"))        ' remove SheetName! from worksheet scoped name
      mRefersTo = mName.RefersTo
      If mNameStr <> "Print_Area" Then
        ' workSHEET scope
        'Debug.Print mName.Parent.Name + " | " + mName.Name + " | " + mName.RefersTo + " | " + mNameStr

        mName.Delete
        ThisWorkbook.Names.Add Name:=mNameStr, RefersTo:=mRefersTo
      End If
        
    ElseIf mName.Parent.Name = ThisWorkbook.Name Then
      ' workBOOK scope
      'Debug.Print "    **** " + ThisWorkbook.Name + " | " + mName.Name + " | " + mName.RefersTo
    End If
    
  Next mName
  Debug.Print "--- END changeScopeAllNames"
End Sub

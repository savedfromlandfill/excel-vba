Attribute VB_Name = "Checkboxes"
'
' select starting cell with text and run this sub-routine
' checkboxes will be added to each neighbouring column
Sub checkboxForEachRow()

  Application.ScreenUpdating = False
  
  Dim startCell As Range
  
  Dim cell As Range
  Dim targetCell As Range
  Dim mText As String
  Dim fieldNamePart1 As String
  Dim fieldNamePart2 As String
  Dim attachName As String
  Dim naName As String
  
  ' checkbox params
  Dim mLeft As Double
  Dim mTop As Double
  Dim mHeight As Double
  Dim mWidth As Double
  
  Set startCell = ActiveCell
  fieldNamePart1 = "inp" & Replace(ActiveSheet.Name, " ", "") & "chk"
  
  Debug.Print "Test"
  
  Do While ActiveCell.Value <> ""

    fieldName = Replace(ActiveCell.Value, " ", "_")
    Set targetCell = ActiveCell.Offset(0, 1)
    
    mLeft = targetCell.Left
    mTop = targetCell.Top
    mHeight = 20
    mWidth = 64

    ' add the checkboxes
    
    ActiveSheet.Checkboxes.Add(mLeft, mTop, mWidth, mHeight).Select
      With Selection
        .Caption = "Attached"
        .Value = xlOff
        .LinkedCell = "G" & targetCell.Row
        .Display3DShading = False
      End With
      
    ActiveSheet.Checkboxes.Add(mLeft + 70, mTop, mWidth, mHeight).Select
      With Selection
        .Caption = "N/A"
        .Value = xlOff
        .LinkedCell = "H" & targetCell.Row
        .Display3DShading = False
      End With
      
    ' name the cells that the checkboxes are linked to
      
    fieldNamePart2 = Replace(ActiveCell.Offset(0, 6).Value, Chr(10), "") ' range names in column I
    attachName = fieldNamePart1 + fieldNamePart2 + "Attached"
    naName = fieldNamePart1 + fieldNamePart2 + "NA"
    
    ActiveWorkbook.Names.Add Name:=attachName, RefersTo:=ActiveCell.Offset(0, 4)
    ActiveWorkbook.Names.Add Name:=naName, RefersTo:=ActiveCell.Offset(0, 5)
      
    ActiveCell.Offset(1, 0).Select
  Loop
  
  ' return selection to user selected cell
  startCell.Select
  
  Application.ScreenUpdating = True
  
End Sub

'
'
'
Sub deleteAllShapes()

  Dim mShape As Shape
  
  For Each mShape In ActiveSheet.Shapes
    mShape.Delete
  Next mShape
End Sub



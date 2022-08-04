Attribute VB_Name = "TestingTools"
Option Explicit

'
'  NUMBER  FORMATS
'  General      .NumberFormat  =  "General"
'  Number       .NumberFormat  =  "0.00"
'  Currency     .NumberFormat  =  "$#,##0.00"
'  Accounting   .NumberFormat  =  "_($*  #,##0.00_);_($*  (#,##0.00);_($*  ""-""??_);_(@_)"
'  Short  Date  .NumberFormat  =  "m/d/yyyy"
'  Long  Date   .NumberFormat  =  "[$-x-sysdate]dddd,  mmmm  dd,  yyyy"
'  Time         .NumberFormat  =  "[$-x-systime]h:mm:ss  AM/PM"
'  Percentage   .NumberFormat  =  "0.00%"
'  Fraction     .NumberFormat  =  "#  ?/?"
  'Scientific   .NumberFormat  =  "0.00E+00"
'  Text         .NumberFormat  =  "@"
'

Sub AllSheetsTestData()

    Dim mSheet   As Worksheet
    Application.ScreenUpdating = False
    
    For Each mSheet In ThisWorkbook.Worksheets
        mSheet.Activate
        Call AddTestDataToNamedRangesInSheet
    Next mSheet
    
    Application.ScreenUpdating = True

End Sub

'
' LOOP ALL EXCEL FILES IN FOLDER
'
' Prompts for a folder then opens all Excel files in that folder and
' populates any unlocked named ranges with 'lorem ipsum' etc data.
'
Sub LoopAllExcelFilesInFolder()

Dim wb As Workbook
Dim myPath As String
Dim myFile As String
Dim myExtension As String
Dim FldrPicker As FileDialog
Dim mSheet As Worksheet

'Optimize Macro Speed
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual

'Retrieve Target Folder Path From User
   Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With


'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.xls*"

'Target Path with Ending Extention
  myFile = Dir(myPath & myExtension)
  
  Dim fileNames() As String
  ReDim fileNames(0)
  
  Do While myFile <> ""
    fileNames(UBound(fileNames)) = myFile
    ReDim Preserve fileNames(UBound(fileNames) + 1)
    ' Get next file name
    myFile = Dir
  Loop
  ReDim Preserve fileNames(UBound(fileNames) - 1)

 Dim i As Integer
'Loop through each Excel file in folder
  For i = LBound(fileNames) To UBound(fileNames)
    
    myFile = fileNames(i)
    Debug.Print "Processing " + myFile
  
  
    'Set variable equal to opened workbook
    Set wb = Workbooks.Open(Filename:=myPath & myFile)

    'Ensure Workbook has opened before moving on to next line of code
      DoEvents
    

    For Each mSheet In wb.Worksheets
      mSheet.Activate
      Call AddTestDataToNamedRangesInSheet
    Next mSheet

    wb.Worksheets(1).Select

    'Save and Close Workbook
      wb.Close SaveChanges:=True
      
    'Ensure Workbook has closed before moving on to next line of code
      DoEvents

  Next i ' loop through fileNames array

'Message Box when tasks are completed
  MsgBox "Task Complete!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub



Sub AddTestDataToNamedRangesInSheet()

    Dim mName As Name
    Dim RefersToSheet As String
    Dim mCell As Range

    For Each mName In ActiveWorkbook.Names


        RefersToSheet = Trim(Split(mName.RefersTo, "!")(0))
        RefersToSheet = Replace(RefersToSheet, "'", "")
        RefersToSheet = Replace(RefersToSheet, "=", "")
        

        If RefersToSheet = ActiveSheet.Name And Not InStr(1, mName.Name, "Print_Area") > 0 Then

            writeToFile (ActiveWorkbook.Name + ", " + mName.Name + ", " + mName.RefersTo + ", " + RefersToSheet + ", " + strOrNull(mName.RefersToRange.NumberFormat))
         
            If mName.RefersToRange.Columns.Count > 1 Or mName.RefersToRange.Rows.Count > 1 Then
              For Each mCell In mName.RefersToRange
                Call ProcessCell(mCell, mName)
              Next mCell
            Else
                Call ProcessCell(mName.RefersToRange, mName)
            End If
            
        End If

    Next mName

End Sub


Function strOrNull(mstring As Variant) As String
    On Error GoTo Error
    strOrNull = mstring
Exit Function
Error:
strOrNull = ""
End Function


Sub ProcessCell(mCell As Range, mName As Name)
                If Not mCell.Locked Then
                
                    If mCell.Interior.Color = 14277081 Then
                        mCell.Value = RandomNum(1, 2)
                
                    ElseIf mCell.NumberFormat = "General" _
                            Or mCell.NumberFormat = "@" Then
                        '  TEXT  /  GENERAL

                        If Right(mName.Name, 8) = "_textbox" Then
                            mCell.Value = TestText(RandomNum(80, 300)) + "."
                        Else
                            mCell.Value = TestText(RandomNum(1, 5))
                        End If

                    ElseIf mCell.NumberFormat = "0" _
                            Or mCell.NumberFormat = "0.00" _
                            Or mCell.NumberFormat = "$#,##0.00" _
                            Or mCell.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)" _
                            Or mCell.NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-" _
                            Or mCell.NumberFormat = "_-$* #,##0_-;-$* #,##0_-;_-$* ""-""??_-;_-@_-" Then
                        '  NUMBER
                
                            If InStr(1, mName.Name, "postcode") > 0 Then
                                mCell.Value = RandomNum(4000, 4999)
                            ElseIf InStr(1, mName.Name, "abn") > 0 Then
                                mCell.Value = RandomNum(1, 32767)
                            Else
                                mCell.Value = RandomNum(1, 100)
                            End If
                    
                    ElseIf mCell.NumberFormat = "0.00%" _
                            Or mCell.NumberFormat = "0%" Then
                        '  PERCENTAGE
                
                        mCell.Value = RandomNum(1, 100) / 100

                    ElseIf mCell.NumberFormat = "m/d/yyyy" Then
                        '  DATE
                
                        mCell.Value = Now()

                    End If
                End If   '  cell  not  locked
End Sub

Function TestText(words As Integer) As String

    Dim i   As Integer
    Dim startWord   As Integer
    
    Dim AllText   As String
    AllText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer metus mauris, egestas vel varius nec, aliquam id leo. " _
            + "Vestibulum ultricies rhoncus quam fringilla consequat. Vestibulum hendrerit viverra ornare. Morbi suscipit aliquet justo, " _
            + "sed finibus lectus accumsan interdum. In facilisis nibh vitae urna fringilla mattis vitae sed dolor. Duis urna justo, " _
            + "sodales et tincidunt et, suscipit et turpis. Suspendisse eleifend massa at euismod luctus. " _
            + "Proin vestibulum consequat erat, quis malesuada turpis tempor non. Aliquam erat volutpat. Ut varius libero et lacus " _
            + "maximus scelerisque. Aenean non est auctor, rhoncus tellus vitae, bibendum elit. Vestibulum id nibh semper, congue ante id, " _
            + "dignissim tellus. Curabitur et volutpat urna. Aenean ut vestibulum tortor. Pellentesque velit elit, aliquam a mauris sed, " _
            + "tempus porttitor ligula. Duis id."

    Dim AllWords() As String
    AllWords = Split(AllText, " ")
    
    If words > UBound(AllWords) Then words = UBound(AllWords)
    
    startWord = RandomNum(0, UBound(AllWords) - words)
    
    For i = startWord To startWord + words - 1
        TestText = TestText + AllWords(i) + " "
    Next i
    
    TestText = Trim(TestText)
    TestText = UCase(Left(TestText, 1)) + Mid(TestText, 2)
    If Right(TestText, 1) = "," Or Right(TestText, 1) = "." Then TestText = Left(TestText, Len(TestText) - 1)

End Function


Function RandomNum(lowerbound, upperbound As Integer) As Integer

    RandomNum = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)

End Function

Sub writeToFile(output As String)
  Dim Path As String
  Dim File As Integer
  
  Path = ThisWorkbook.Path
  If Right(Path, 1) <> "\" Then Path = Path + "\"

  Path = Path + "form_testing_output.txt"

  ' LOG FILE: setup the log file, create if doesn't already exist
   File = FreeFile
  If Dir(Path) = "" Then
      ' file doesn't exist, create
      Open Path For Output As File
  Else
      ' file exists, append
      Open Path For Append As File
  End If
  
  Print #File, output
  Close File
  
End Sub

Attribute VB_Name = "DataCopy"
Option Explicit


Sub Run_Data_Copy()

  Dim MapFile As Integer
  Dim MapPath As String
  
  Dim LogFile As Integer
  Dim LogPath As String
  
  Dim Settings() As String
  ReDim Settings(1, 0)
  
  Dim sCol As Integer
  Dim sRow As Integer
  Dim settingLabel As String
  Dim settingValue As String
  Dim settingTemp As String
  
  Dim DirectoryPath As String
  Dim FromPath As String
  Dim ToPath As String
  Dim NewFromPath As String
  Dim NewToPath As String
  Dim FromToTemp As String
  
  Dim FromField As String
  Dim ToField As String
  Dim NewFromField As String
  Dim NewToField As String
  Dim FromValue As String
  Dim ToValue As String
  
  Dim FromNameCheck As Integer
  Dim ToNameCheck As Integer
  
  ' for table copy
  Dim mCol As Integer
  Dim mRow As Integer
  Dim lastAddress As String
  Dim colCnt As Integer
  
  DirectoryPath = ThisWorkbook.Path
  If Right(DirectoryPath, 1) <> "\" Then DirectoryPath = DirectoryPath + "\"
  
  Dim MapFileLine As String
  
  MapPath = ThisWorkbook.Path
  If Right(MapPath, 1) <> "\" Then MapPath = MapPath + "\"
  LogPath = MapPath
  
  MapPath = MapPath + "data_copy_map.txt"
  LogPath = LogPath + "data_copy_log.txt"

  ' MESSAGE BOX: Are you sure?
  Dim MsgStr As String
  MsgStr = "About to run Data Copy Tool!" + Chr(10) + Chr(10)
  MsgStr = MsgStr + "Map File Used: " + Chr(10) + MapPath + Chr(10) + Chr(10)
  MsgStr = MsgStr + "Log File Used: " + Chr(10) + LogPath + Chr(10) + Chr(10)
  MsgStr = MsgStr + "Are you sure you want to continue?"
  If MsgBox(MsgStr, vbYesNo) = vbNo Then Exit Sub

  Application.ScreenUpdating = False
  
  MapFile = FreeFile
  Open MapPath For Input As MapFile
  
  
  
  ' LOG FILE: setup the log file, create if doesn't already exist
  LogFile = FreeFile
  If Dir(LogPath) = "" Then
      ' file doesn't exist, create
      Open LogPath For Output As LogFile
  Else
      ' file exists, append
      Open LogPath For Append As LogFile
  End If
  
  Print #LogFile, "****** STARTING DATA COPY ******"
  
  
  
  ' MAP FILE: check the map file exists
  If Dir(MapPath) = "" Then
      ' map file not found, log and exit
      Print #LogFile, Now + "Error: Map file not found at " + MapPath
      MsgBox ("Map file not found.")
      Exit Sub
  End If
  
  Print #LogFile, Format(Now, "yyyy-mm-dd hh:MM:ss ") + "*** Reading map file: " + MapPath
  
  
  
  ' main processing
  While Not EOF(MapFile)
    Line Input #MapFile, MapFileLine
    
    MapFileLine = LCase(Trim(MapFileLine))

    ' get file definitions
    If MapFileLine = "" Or Left(MapFileLine, 1) = "-" Then
      ' ignore blank lines and
      ' ignore any line starting with dash (these are for comments/readability)
       
    ' DEFINITIONS
    ElseIf Left(MapFileLine, 7) = "define:" Then
      If Settings(0, UBound(Settings, 2)) <> "" Then
        ReDim Preserve Settings(1, UBound(Settings, 2) + 1)
      End If
          
      settingTemp = Right(MapFileLine, Len(MapFileLine) - 7)
      settingLabel = Split(settingTemp, "=")(0)
      settingValue = Split(settingTemp, "=")(1)
      
      Settings(0, UBound(Settings, 2)) = settingLabel
      Settings(1, UBound(Settings, 2)) = Replace(settingValue, """", "")
       
       
    ' FROM,TO
    ElseIf Left(MapFileLine, 7) = "fromto:" Then
      FromToTemp = Right(MapFileLine, Len(MapFileLine) - 7)
      Print #LogFile, Format(Now, "yyyy-mm-dd hh:MM:ss ") + "*** From,To: " + FromToTemp
      
      NewFromPath = GetArrayValue(Settings, Trim(Split(FromToTemp, ",")(0)))
      NewToPath = GetArrayValue(Settings, Trim(Split(FromToTemp, ",")(1)))
      
      If NewFromPath <> FromPath And IsWorkbookOpen(FromPath) Then Workbooks(FromPath).Close SaveChanges:=False
      If NewToPath <> ToPath And IsWorkbookOpen(ToPath) Then Workbooks(ToPath).Close SaveChanges:=True
      
      FromPath = NewFromPath
      Print #LogFile, Format(Now, "yyyy-mm-dd hh:MM:ss ") + "*** Opening Source: " + FromPath
      Workbooks.Open (DirectoryPath + FromPath)
      
      ToPath = NewToPath
      Print #LogFile, Format(Now, "yyyy-mm-dd hh:MM:ss ") + "*** Opening Target: " + ToPath
      Workbooks.Open (DirectoryPath + ToPath)
       
       
    ' MAPPING/LOADING
    Else
      ' must be a mapping line
      
      ' FromField
      ' ToField
      ' NewFromField
      ' NewToField
      ' FromValue
      ' ToValue
      
      ' get field names
      If InStr(1, MapFileLine, ",") Then
        FromField = Trim(Split(MapFileLine, ",")(0))
        ToField = Trim(Split(MapFileLine, ",")(1))
      Else
        ' a line with one field name means we're mapping on the same name in both workbooks
        FromField = Trim(MapFileLine)
        ToField = FromField
      End If
      
      ' check from name exists
      FromNameCheck = 0
      On Error Resume Next
        FromNameCheck = Len(Workbooks(FromPath).Names(FromField).Name)
      On Error GoTo 0
      If FromNameCheck <= 0 Then
        Print #LogFile, Format(Now, "yyyy-mm-dd hh:MM:ss ") + "!ERROR: From Field '" + FromField; "' not found in " + FromPath
      End If
      
      ' check to name exists
      ToNameCheck = 0
      On Error Resume Next
        ToNameCheck = Len(Workbooks(ToPath).Names(ToField).Name)
      On Error GoTo 0
      If ToNameCheck <= 0 Then
        Print #LogFile, Format(Now, "yyyy-mm-dd hh:MM:ss ") + "!ERROR: To Field '" + ToField; "' not found in " + ToPath
      End If
      
      If FromNameCheck > 0 And ToNameCheck > 0 Then
      ' COPY DATA - need to handle range with multiple rows/columns here, eg. a 'table'
      
        If Workbooks(FromPath).Names(FromField).RefersToRange.Columns.Count > 1 Or _
           Workbooks(FromPath).Names(FromField).RefersToRange.Rows.Count > 1 Then
          ' Table copy
      
          ' NOTE: The below column count check doesn't work because it doesn't take into
          '       account merged cells. The table copy code handles merged cells by checking
          '       for a change of MergeArea.Address
          '
          ' if ToField column count doesn't match then error: can't copy different number of columns
          'If Workbooks(ToPath).Names(ToField).RefersToRange.Columns.Count <> _
          '   Workbooks(FromPath).Names(FromField).RefersToRange.Columns.Count Then
          '   Print #LogFile, Format(Now, "yyyy-mm-dd hh:MM:ss ") + "!ERROR: Table Copy: To Field '" + ToField; "' and From Field '" + FromField + _
          '                                                         "' can't copy between tables with different number of columns."
          'Else ' column counts match, proceed with table copy
          
            ' if ToField row count less then FromField then insert rows
            If Workbooks(ToPath).Names(ToField).RefersToRange.Rows.Count < _
               Workbooks(FromPath).Names(FromField).RefersToRange.Rows.Count Then
               Call expand_range(Workbooks(ToPath).Names(ToField).RefersToRange, _
                 Workbooks(FromPath).Names(FromField).RefersToRange.Rows.Count - Workbooks(ToPath).Names(ToField).RefersToRange.Rows.Count, _
                 True, True)
            End If
            
            ' loop through rows and columns and copy to destination range
            lastAddress = ""

            For mRow = 1 To Workbooks(FromPath).Names(FromField).RefersToRange.Rows.Count
              colCnt = 0
              For mCol = 1 To Workbooks(FromPath).Names(FromField).RefersToRange.Columns.Count
      
                If Workbooks(FromPath).Names(FromField).RefersToRange(mRow, mCol).MergeArea.Address <> lastAddress Then
                  Workbooks(ToPath).Names(ToField).RefersToRange(mRow, mCol).Value = _
                    Workbooks(FromPath).Names(FromField).RefersToRange(mRow, mCol).Value

                  lastAddress = Workbooks(FromPath).Names(FromField).RefersToRange(mRow, mCol).MergeArea.Address
                  colCnt = colCnt + 1
                End If
              Next mCol
            Next mRow
          
          'End If ' column counts don't match
        Else
          ' Standard copy
          FromValue = Workbooks(FromPath).Names(FromField).RefersToRange.Value
          Workbooks(ToPath).Names(ToField).RefersToRange.Value = FromValue
                
        End If ' Columns or Rows Count > 1

        Print #LogFile, Format(Now, "yyyy-mm-dd hh:MM:ss ") + "'" + FromField + "' copied to '" + ToField + "'"

      End If ' FromNameCheck and ToNameCheck > 0

    End If
  Wend
  
  
  
  Print #LogFile, Format(Now, "yyyy-mm-dd hh:MM:ss ") + "*** Settings Used:"
  ' log array setup for testing/debugging
  For sCol = LBound(Settings, 2) To UBound(Settings, 2)
    For sRow = LBound(Settings, 1) To UBound(Settings, 1)
      Print #LogFile, Format(Now, "yyyy-mm-dd hh:MM:ss ") + "Setting " + Str(sRow) + "," + Str(sCol) + ": " + Settings(sRow, sCol)
    Next sRow
  Next sCol
  
  
  
  ' close workbooks
  If FromPath <> "" And IsWorkbookOpen(FromPath) Then Workbooks(FromPath).Close SaveChanges:=False
  If ToPath <> "" And IsWorkbookOpen(ToPath) Then Workbooks(ToPath).Close SaveChanges:=True
  
  ' close txt file
  Close MapFile
  
  ' close log file
  Print #LogFile, "****** FINISHED DATA COPY ******"
  Print #LogFile, ""
  Print #LogFile, ""
  Close LogFile
  
  
  
  Application.ScreenUpdating = True
  
  MsgBox ("Data copy complete." + Chr(10) + Chr(10) + "See log file for more detail:" + Chr(10) + LogPath)

End Sub






Function GetArrayValue(ByRef mArray() As String, mLabel As String) As String
  Dim mCol As Integer
  Dim mRow As Integer
  
  For mCol = LBound(mArray, 2) To UBound(mArray, 2)
    If mArray(0, mCol) = mLabel Then
      GetArrayValue = mArray(1, mCol)
      Exit Function
    End If
  Next mCol
  
  GetArrayValue = ""
End Function






Function IsWorkbookOpen(Name As String) As Boolean
  Dim mWorkbook As Workbook
  On Error Resume Next
  Set mWorkbook = Application.Workbooks.Item(Name)
  IsWorkbookOpen = (Not mWorkbook Is Nothing)
End Function




' https://stackoverflow.com/questions/2616355/how-to-insert-a-new-row-into-a-range-and-copy-formulas
' Appends one or more rows to a range.
' You can choose if you want to keep formulas and if you want to insert entire sheet rows.
Sub expand_range( _
                        target_range As Range, _
                        Optional num_rows As Integer = 1, _
                        Optional insert_entire_sheet_row As Boolean = False, _
                        Optional keep_formulas As Boolean = False _
                        )

    Application.ScreenUpdating = False
    On Error GoTo Cleanup

    Dim original_cell As Range: Set original_cell = ActiveCell
    Dim last_row As Range: Set last_row = target_range.Rows(target_range.Rows.Count)

    ' Insert new row(s) above the last row and copy contents from last row to the new one(s)
    IIf(insert_entire_sheet_row, last_row.Cells(1).EntireRow, last_row) _
        .Resize(num_rows).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
    last_row.Copy
    last_row.Offset(-num_rows).PasteSpecial
    last_row.ClearContents

    On Error Resume Next ' This will fail if there are no formulas and keep_formulas = True
        If keep_formulas Then
            With last_row.Offset(-num_rows).SpecialCells(xlCellTypeFormulas)
                .Copy
                .Offset(1).Resize(num_rows).PasteSpecial
            End With
        End If
    On Error GoTo Cleanup

Cleanup:
    On Error GoTo 0
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    original_cell.Select
    If Err Then Err.Raise Err.Number, , Err.Description
End Sub


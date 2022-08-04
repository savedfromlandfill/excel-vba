Attribute VB_Name = "PowerQueryVBA"

Sub test()

  Dim msheet As Worksheet
  Dim mlist As ListObject
  
  For Each msheet In ThisWorkbook.Sheets
     
     For Each mlist In msheet.ListObjects
        Debug.Print msheet.Name & ": " & mlist.Name & ", " & mlist.QueryTable.CommandText

     Next mlist
    
  
  Next msheet
End Sub


Sub test_connections()

  Dim cn As WorkbookConnection

  For Each cn In ThisWorkbook.Connections
    
    Debug.Print cn.Name
    Debug.Print cn.OLEDBConnection.Connection
    Debug.Print cn.OLEDBConnection.CommandText
    Debug.Print
  Next cn
  
End Sub

Sub test_queries()

  Debug.Print "START"
  Dim qry As WorkbookQuery
  
  For Each qry In ThisWorkbook.Queries
     Debug.Print "NAME: '" & qry.Name & "'"
     Debug.Print qry.Formula
     Debug.Print "************************************************"
  Next qry
  
  Debug.Print "END"
End Sub



Sub optimise_queries()

  Debug.Print "// START Query Output"
  Dim qry As WorkbookQuery
  
  Dim removeLoc, removeFieldsStart, removeFieldsEnd As Integer
  Dim txtLine, txtField As Variant
  Dim allRemovedFields As String
  Dim tempString As String
  
  Dim removeList1() As String
  ReDim removeList1(0)
  
  Dim i As Integer
  
  
  'For Each qry In ThisWorkbook.Queries
  Set qry = ThisWorkbook.Queries(1)
  
     Debug.Print "// QUERY: " & qry.Name
     For Each txtLine In Split(qry.Formula, vbNewLine)
        Debug.Print txtLine
        removeLoc = InStr(1, txtLine, "Table.RemoveColumns")
        If removeLoc > 0 Then
        ' IF REMOVE COMMAND FOUND
        
          removeFieldsStart = InStr(removeLoc, txtLine, "{")
          removeFieldsEnd = InStr(removeFieldsStart + 1, txtLine, "}")
          
          allRemovedFields = allRemovedFields & Mid(txtLine, removeFieldsStart + 1, removeFieldsEnd - removeFieldsStart - 1) + ", "
          
          For Each txtField In Split(Mid(txtLine, removeFieldsStart + 1, removeFieldsEnd - removeFieldsStart - 1), ",")
            If InStr(1, allRemovedFields, txtField) > 0 Then
              
              removeList1(UBound(removeList1)) = Trim(replace(txtField, """", ""))
              ReDim Preserve removeList1(UBound(removeList1) + 1)
              
            End If
          Next txtField
          
          
        ' END IF REMOVE COMMAND FOUND
        End If
     Next txtLine
     
     ReDim Preserve removeList1(UBound(removeList1) - 1)
     
     Debug.Print allRemovedFields
     
     tempString = ""
     For i = LBound(removeList1) To UBound(removeList1)
       tempString = tempString + """" + removeList1(i) + """, "
     Next i
     Debug.Print tempString
     tempString = ""
    
     ReDim removeList1(1)
     allRemovedFields = ""
     
  Debug.Print String(2, vbNewLine)

  'Next qry
  
  Debug.Print "// END Query Output"
End Sub


Sub refreshwithtime()
 
  Dim startTime, endTime As Date
  Dim i As Long
  
  Dim currBackgroundSetting As Boolean
  
  startTime = Now()
  Debug.Print "Start " & startTime
  
  currBackgroundSetting = ThisWorkbook.Connections("Query - Approved FIP").OLEDBConnection.BackgroundQuery

  ThisWorkbook.Connections("Query - Approved FIP").OLEDBConnection.BackgroundQuery = False
  ThisWorkbook.Connections("Query - Approved FIP").Refresh
  ThisWorkbook.Connections("Query - Approved FIP").OLEDBConnection.BackgroundQuery = currBackgroundSetting
  
  endTime = Now()
  Debug.Print "Finish " & endTime
  
  Debug.Print DateDiff("s", startTime, endTime) & " seconds to refresh"
  
End Sub







Sub search_replace_all_queries()
  Dim qry As WorkbookQuery
  Dim search1, replace1, search2, replace2, tmpFormula As String
  
  search1 = "test1"
  replace1 = "test2"
  
  search2 = "testreplace1"
  replace2 = "testreplace2"

  For Each qry In ThisWorkbook.Queries
     tmpFormula = qry.Formula
     tmpFormula = replace(tmpFormula, search1, replace1)
     tmpFormula = replace(tmpFormula, search2, replace2)
     qry.Formula = tmpFormula
     Debug.Print "Updated " + qry.Name
  Next qry

  Debug.Print "Completed search/replace in all queries."
End Sub

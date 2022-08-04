 Option Explicit
'---------------------------------------------------------------------------------------
' Procedure : RescopeNamedRangesToWorkbook
' Author    : JS20'07'11
' Date      : 11/18/2013
' Purpose   : Rescopes the parent of worksheet scoped named ranges to the active workbook
' for each named range with a scope equal to the active sheet in the active workbook.
'---------------------------------------------------------------------------------------

Public Sub RescopeNamedRangesToWorkbook()
Dim wb As Workbook
Dim ws As Worksheet
Dim objName As Name
Dim sWsName As String
Dim sWbName As String
Dim sRefersTo As String
Dim sObjName As String
Set wb = ActiveWorkbook
Set ws = ActiveSheet
sWsName = ws.Name
sWbName = wb.Name

'Loop through names in worksheet.
For Each objName In ws.Names
'Check name is visble.
    If objName.Visible = True Then
'Check name refers to a range on the active sheet.
        If InStr(1, objName.RefersTo, sWsName, vbTextCompare) Then
            sRefersTo = objName.RefersTo
            sObjName = objName.Name
'Check name is scoped to the worksheet.
            If objName.Parent.Name <> sWbName Then
'Delete the current name scoped to worksheet replacing with workbook scoped name.
                sObjName = Mid(sObjName, InStr(1, sObjName, "!") + 1, Len(sObjName))
                objName.Delete
                wb.Names.Add Name:=sObjName, RefersTo:=sRefersTo
            End If
        End If
    End If
Next objName
End Sub
'---------------------------------------------------------------------------------------
' Procedure : RescopeNamedRangesToWorksheet
' Author    : JS20'07'11
' Date      : 11/18/2013
' Purpose   : Rescopes each workbook scoped named range to the specific worksheet to
' which the range refers for each named range that refers to the active worksheet.
'---------------------------------------------------------------------------------------

Public Sub RescopeNamedRangesToWorksheet()
Dim wb As Workbook
Dim ws As Worksheet
Dim objName As Name
Dim sWsName As String
Dim sWbName As String
Dim sRefersTo As String
Dim sObjName As String
Set wb = ActiveWorkbook
Set ws = ActiveSheet
sWsName = ws.Name
sWbName = wb.Name

'Loop through names in worksheet.
For Each objName In wb.Names
'Check name is visble.
    If objName.Visible = True Then
'Check name refers to a range on the active sheet.
        If InStr(1, objName.RefersTo, sWsName, vbTextCompare) Then
            sRefersTo = objName.RefersTo
            sObjName = objName.Name
'Check name is scoped to the workbook.
            If objName.Parent.Name = sWbName Then
'Delete the current name scoped to workbook replacing with worksheet scoped name.
                objName.Delete
                ws.Names.Add Name:=sObjName, RefersTo:=sRefersTo
            End If
        End If
    End If
Next objName
End Sub
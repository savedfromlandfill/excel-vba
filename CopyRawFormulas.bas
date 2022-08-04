Attribute VB_Name = "CopyRawFormulas"
Sub copyRawFormulas()
  For Row = 1 To 50
    For Col = 1 To 20
       ActiveSheet.Cells(Row, Col).Formula = Workbooks("JRGF Proposals List.XLSX").Sheets("Summary").Cells(Row, Col).Formula
       ActiveSheet.Rows(Row).RowHeight = Workbooks("JRGF Proposals List.XLSX").Sheets("Summary").Rows(Row).RowHeight
       ActiveSheet.Columns(Col).ColumnWidth = Workbooks("JRGF Proposals List.XLSX").Sheets("Summary").Columns(Col).ColumnWidth
    Next Col
  Next Row
End Sub


Sub disableFormulas()
  For Row = 1 To 50
    For Col = 1 To 20
       ActiveSheet.Cells(Row, Col).Formula = "'" & ActiveSheet.Cells(Row, Col).Formula
    Next Col
  Next Row
End Sub

Sub enableFormulas()
  For Row = 1 To 50
    For Col = 1 To 20
       ActiveSheet.Cells(Row, Col).Formula = ActiveSheet.Cells(Row, Col).Value
    Next Col
  Next Row
End Sub

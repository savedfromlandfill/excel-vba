Attribute VB_Name = "FixPrintQualityAllSheets"
   Sub SetPrintQuality()

For Each xSheet In ActiveWorkbook.Sheets
           xSheet.PageSetup.PrintQuality = 600
       Next xSheet

End Sub




Sub Run_Macro()
    Dim topLeft
    Dim bottomLeft
    Dim macroWorksheet As Worksheet
    Set macroWorksheet = Worksheets("Macro")
    
    dataWorksheetName = macroWorksheet.Range("B3").Value
    topLeft = macroWorksheet.Range("C3").Value
    bottomLeft = macroWorksheet.Range("D3").Value
    
    Dim dataWorksheet As Worksheet
    Set dataWorksheet = Worksheets(dataWorksheetName)
    
    macroWorksheet.Range("J3:J" & Range("J3").End(xlDown).Row).Clear
    
    Application.DisplayAlerts = False
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name Like "Generated-*" Then
            ws.Delete
        End If
    Next
    Application.DisplayAlerts = True
   
    dataWorksheet.Range(topLeft, bottomLeft).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=macroWorksheet.Range("J3"), Unique:=True
    
    Dim uniqueCellValues As Range
    Set uniqueCellValues = macroWorksheet.Range("J4:J" & Range("J4").End(xlDown).Row)
    uniqueCellValues.Sort Key1:=uniqueCellValues.Cells(1), Order1:=xlAscending, Header:=xlNo
    uniqueCellValues.Select
    
    Dim newSheetName
    Dim filteredRange As Range

    For Each cell In uniqueCellValues.Cells
        newSheetName = "Generated-" & Left(cell.Value, 20)
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = newSheetName
        dataWorksheet.Range("A1").AutoFilter Field:=1, Criteria1:=cell.Value
        
        Set filteredRange = dataWorksheet.AutoFilter.Range
        filteredRange.Copy (Sheets(newSheetName).Range("A1"))
    Next
    
    If dataWorksheet.AutoFilterMode = True Then
        dataWorksheet.AutoFilterMode = False
    End If
    
End Sub

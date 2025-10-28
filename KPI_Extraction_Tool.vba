Sub ExtractByCellMapping()
    Dim folderPath As String, fileName As String
    Dim wbSource As Workbook, wsSource As Worksheet, wsTarget As Worksheet, wsMapping As Worksheet
    Dim targetRow As Long, i As Long
    Dim kpiCount As Long
    Dim cellAddress As String, kpiName As String
    Dim fullPath As String
    
    ' Set up mapping and target sheets
    Set wsMapping = ThisWorkbook.Sheets("KPI_Mapping")
    Set wsTarget = SetupTargetSheetFromMapping(wsMapping)
    
    ' Select folder
    folderPath = MacScript("return (choose folder with prompt ""Select data folder"") as string")
    If folderPath = "" Then Exit Sub
    If Right(folderPath, 1) <> ":" Then folderPath = folderPath & ":"
    
    targetRow = 2
    fileName = Dir(folderPath & "*.xlsx")
    
    ' Get KPI count
    kpiCount = wsMapping.Cells(wsMapping.Rows.Count, "A").End(xlUp).Row - 1
    
    Do While fileName <> ""
        Application.StatusBar = "Processing: " & fileName
        
        fullPath = folderPath & fileName
        fullPath = Replace(fullPath, ":", "/")
        fullPath = Replace(fullPath, "Macintosh HD", "")
        
        Set wbSource = Workbooks.Open(fullPath, ReadOnly:=True)
        If Not wbSource Is Nothing Then
            Set wsSource = wbSource.Sheets(1)
            
            ' First column: file name
            wsTarget.Cells(targetRow, 1).Value = fileName
            
            ' Extract each KPI based on mapping
            For i = 1 To kpiCount
                kpiName = wsMapping.Cells(i + 1, 1).Value
                cellAddress = wsMapping.Cells(i + 1, 2).Value
                
                If cellAddress <> "" Then
                    On Error Resume Next
                    wsTarget.Cells(targetRow, i + 1).Value = wsSource.Range(cellAddress).Value
                    If Err.Number <> 0 Then
                        wsTarget.Cells(targetRow, i + 1).Value = "Invalid Address"
                    End If
                    On Error GoTo 0
                End If
            Next i
            
            wbSource.Close False
            targetRow = targetRow + 1
        End If
        
        fileName = Dir
    Loop
    
    Application.StatusBar = False
    MsgBox "Data extraction completed! Processed " & (targetRow - 2) & " files."
End Sub

Function SetupTargetSheetFromMapping(wsMapping As Worksheet) As Worksheet
    Dim ws As Worksheet
    Dim i As Long, kpiCount As Long
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Extracted_Data").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Extracted_Data"
    
    ' Set up headers
    ws.Range("A1").Value = "Data Source"
    
    kpiCount = wsMapping.Cells(wsMapping.Rows.Count, "A").End(xlUp).Row - 1
    
    For i = 1 To kpiCount
        ws.Cells(1, i + 1).Value = wsMapping.Cells(i + 1, 1).Value
    Next i
    
    Set SetupTargetSheetFromMapping = ws
End Function

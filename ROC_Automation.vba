'==================================================================================
' Module Name: Financial Data Pipeline & Multi-Fund Loader Automation
' Author: Ting Huang
' Description: Automates Return of Capital (ROC) data extraction, multi-layer
'              exception filtering, string parsing, and automated GL file splitting.
'==================================================================================

Option Explicit

Sub FilteroutROC()
    Dim ws As Worksheet
    Dim macroWs As Worksheet
    Dim rng As Range
    
    Set ws = ThisWorkbook.Sheets("Daily_Check")
    Set macroWs = ThisWorkbook.Sheets("Macro")
    
    Application.ScreenUpdating = False
    
    ' Clear any existing filters
    If ws.FilterMode Then ws.ShowAllData
    
    ' Apply filter to column G (7) to show cells containing "ROC"
    ws.Range("A1").AutoFilter Field:=7, Criteria1:="*ROC*", Operator:=xlAnd
    
    ' Apply filter to column D (4) to exclude "Security as underlier" and empty cells
    ws.Range("A1").AutoFilter Field:=4, Criteria1:="<>Security as underlier", Operator:=xlAnd, Criteria2:="<>"
    
    ' Copy visible cells from column B
    Set rng = ws.Range("B:B").SpecialCells(xlCellTypeVisible)
    rng.Copy
    
    ' Paste to a temporary column in the Macro sheet
    macroWs.Range("Z1").PasteSpecial xlPasteValues
    
    ' Remove duplicates from the temporary column
    macroWs.Range("Z:Z").RemoveDuplicates Columns:=1, Header:=xlYes
    
    ' Copy unique values to column H
    macroWs.Range("Z:Z").Copy macroWs.Range("H1")
    
    ' Clear the temporary column
    macroWs.Range("Z:Z").Clear
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
    MsgBox "Filtered data in Daily_Check and copied unique values from column B to column H in Macro sheet.", vbInformation
End Sub


Sub CreateROCDATASheet()
    Dim wsActivePL As Worksheet
    Dim wsDailyCheck As Worksheet
    Dim wsROCDATA As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Set references to sheets
    Set wsActivePL = ThisWorkbook.Sheets("ActivePL")
    Set wsDailyCheck = ThisWorkbook.Sheets("Daily_Check")
    
    ' Check if ROCDATA sheet exists, if not create it
    On Error Resume Next
    Set wsROCDATA = ThisWorkbook.Sheets("ROCDATA")
    On Error GoTo 0
    
    If wsROCDATA Is Nothing Then
        Set wsROCDATA = ThisWorkbook.Sheets.Add
        wsROCDATA.Name = "ROCDATA"
    End If
    
    ' Preserve the first row (header) of ROCDATA
    If wsROCDATA.Cells(1, 1).Value = "" Then
        wsROCDATA.Rows(1).Value = Array("Column A", "Column B", "Column C", "Column D", "Column E", "Column F", "Column G", "Column H", "Column I", "Column J", "Column K", "Column L", "Column M")
    End If
    
    ' Clear existing data in ROCDATA, preserving the header
    wsROCDATA.Rows("2:" & wsROCDATA.Rows.Count).Clear
    
    ' Find last row in ActivePL
    lastRow = wsActivePL.Cells(wsActivePL.Rows.Count, "B").End(xlUp).Row
    
    ' Loop through rows in ActivePL and populate ROCDATA
    For i = 2 To lastRow
        wsROCDATA.Cells(i, 1).Value = wsActivePL.Cells(i, "B").Value
        wsROCDATA.Cells(i, 2).Value = wsActivePL.Cells(i, "C").Value
        wsROCDATA.Cells(i, 3).Value = wsActivePL.Cells(i, "D").Value
        wsROCDATA.Cells(i, 4).Value = wsActivePL.Cells(i, "G").Value
        wsROCDATA.Cells(i, 5).Value = wsActivePL.Cells(i, "H").Value
        wsROCDATA.Cells(i, 6).Value = wsActivePL.Cells(i, "I").Value
        wsROCDATA.Cells(i, 7).Value = wsActivePL.Cells(i, "J").Value
        wsROCDATA.Cells(i, 8).Value = wsActivePL.Cells(i, "L").Value
        wsROCDATA.Cells(i, 9).Value = wsActivePL.Cells(i, "AD").Value
        
        ' Dynamically parse text patterns for terms and payment dates using string formulas
        wsROCDATA.Cells(i, 10).Formula = _
            "=IFERROR(VALUE(MID(VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE), " & _
            "IF(ISNUMBER(FIND(""ROC TERMS:"",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE))), " & _
            "FIND(""ROC TERMS:"",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE))+11, " & _
            "FIND(""ROC TERMS"",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE))+10), " & _
            "FIND("" "",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE)," & _
            "IF(ISNUMBER(FIND(""ROC TERMS:"",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE))), " & _
            "FIND(""ROC TERMS:"",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE))+11, " & _
            "FIND(""ROC TERMS"",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE))+10)) - " & _
            "IF(ISNUMBER(FIND(""ROC TERMS:"",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE))), " & _
            "FIND(""ROC TERMS:"",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE))+11, " & _
            "FIND(""ROC TERMS"",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE))+10))),"""")"

        wsROCDATA.Cells(i, 13).Formula = _
            "=IFERROR(DATEVALUE(MID(VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE), " & _
            "IF(ISNUMBER(FIND(""PD:"",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE))), " & _
            "FIND(""PD:"",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE))+4, " & _
            "FIND(""PD"",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE))+3), " & _
            "IFERROR(FIND(""|"",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE)," & _
            "IF(ISNUMBER(FIND(""PD:"",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE))), " & _
            "FIND(""PD:"",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE))+4, " & _
            "FIND(""PD"",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE))+3)) - " & _
            "IF(ISNUMBER(FIND(""PD:"",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE))), " & _
            "FIND(""PD:"",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE))+4, " & _
            "FIND(""PD"",VLOOKUP(D" & i & ",Daily_Check!B:G,6,FALSE))+3), 10))),"""")"
        
        ' Calculate Column K (Quantity * Unit Rate)
        wsROCDATA.Cells(i, 11).Formula = "=H" & i & "*J" & i
        
        ' Record processing date
        wsROCDATA.Cells(i, 12).Value = Date
    Next i
    
    ' Format date columns
    wsROCDATA.Columns("L:M").NumberFormat = "mm/dd/yyyy"
    
    ' Restore settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "ROCDATA sheet has been created and populated successfully.", vbInformation
End Sub


Sub ExportROCDATAToAccrualWorkbook()
    Dim wsROCDATA As Worksheet
    Dim wsTempSheet As Worksheet
    Dim wsAccrualData As Worksheet
    Dim accrualWorkbook As Workbook
    Dim lastRowROCDATA As Long
    Dim i As Long, j As Long
    Dim rngToCopy As Range
    Dim targetFilePath As String
    
    ' Set reference to ROCDATA sheet
    On Error Resume Next
    Set wsROCDATA = ThisWorkbook.Sheets("ROCDATA")
    On Error GoTo 0
    
    If wsROCDATA Is Nothing Then
        MsgBox "ROCDATA sheet not found. Please run CreateROCDATASheet first.", vbExclamation
        Exit Sub
    End If
    
    ' Create a temporary staging sheet
    Set wsTempSheet = ThisWorkbook.Sheets.Add
    
    ' Find the last row in ROCDATA
    lastRowROCDATA = wsROCDATA.Cells(wsROCDATA.Rows.Count, "A").End(xlUp).Row
    
    ' Copy header to temporary sheet
    wsROCDATA.Rows(1).Copy wsTempSheet.Rows(1)
    
    ' Copy data leaving required alternating line gaps for target template
    For i = 2 To lastRowROCDATA
        For j = 1 To 13
            wsTempSheet.Cells(2 * i - 2, j).Value = wsROCDATA.Cells(i, j).Value
        Next j
    Next i
    
    Set rngToCopy = wsTempSheet.Range("A1:M" & (2 * lastRowROCDATA - 2))
    
    ' SANITIZED FILE PATH (Generic placeholder for GitHub portfolio)
    targetFilePath = "C:\Data\Client_Accruals\Running_Accrual_ROC_Master.xlsx"
    
    On Error Resume Next
    Set accrualWorkbook = Workbooks.Open(targetFilePath)
    On Error GoTo 0
    
    If accrualWorkbook Is Nothing Then
        MsgBox "Could not open target Accrual Workbook. Please verify the destination file path.", vbExclamation
        Application.DisplayAlerts = False
        wsTempSheet.Delete
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    Set wsAccrualData = accrualWorkbook.Sheets("Accrual Data")
    
    ' Paste transformed dataset
    rngToCopy.Copy
    wsAccrualData.Range("A1").PasteSpecial xlPasteValues
    
    Application.CutCopyMode = False
    
    ' Clean up temporary sheet
    Application.DisplayAlerts = False
    wsTempSheet.Delete
    Application.DisplayAlerts = True
    
    accrualWorkbook.Save
    accrualWorkbook.Activate
    
    MsgBox "ROCDATA has been exported to the destination Accrual workbook.", vbInformation
End Sub


Sub SplitGLDataIntoXLSXFiles()
    Dim wsGL As Worksheet
    Dim wsDatabaseInfo As Worksheet
    Dim databaseWorkbook As Workbook
    Dim dict As Object
    Dim wbDict As Object
    Dim fund As Variant
    Dim database As Variant
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim outputFolderPath As String
    Dim dbMappingFilePath As String
    Dim fileName As String
    Dim todayDate As String
    Dim wb As Workbook
    Dim rowCount As Long
    Dim key As Variant
    
    Set wsGL = ThisWorkbook.Sheets("GL")
    
    ' SANITIZED PATHS (Generic placeholders)
    dbMappingFilePath = "C:\Data\Client_Accruals\Fund_Mapping_Master.xlsx"
    outputFolderPath = "C:\Data\Client_Accruals\Loader\"
    
    On Error Resume Next
    Set databaseWorkbook = Workbooks.Open(dbMappingFilePath)
    On Error GoTo 0
    
    If databaseWorkbook Is Nothing Then
        MsgBox "Could not open the master fund database mapping file. Please check the path.", vbExclamation
        Exit Sub
    End If
    
    Set wsDatabaseInfo = databaseWorkbook.Sheets(1)
    
    ' Initialize dictionary for fast key-value lookup O(1)
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Map Funds to respective Database schemas
    For j = 1 To wsDatabaseInfo.Cells(1, Columns.Count).End(xlToLeft).Column
        database = wsDatabaseInfo.Cells(1, j).Value
        For i = 2 To wsDatabaseInfo.Cells(Rows.Count, j).End(xlUp).Row
            If Not IsEmpty(wsDatabaseInfo.Cells(i, j)) Then
                dict(Trim(wsDatabaseInfo.Cells(i, j).Value)) = database
            End If
        Next i
    Next j
    
    databaseWorkbook.Close SaveChanges:=False
    
    todayDate = Format(Date, "YYYYMMDD")
    
    lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row
    lastCol = wsGL.Cells(1, wsGL.Columns.Count).End(xlToLeft).Column
    
    Set wbDict = CreateObject("Scripting.Dictionary")
    
    Application.ScreenUpdating = False
    
    ' Split rows into dynamic destination workbooks by Database key
    For i = 2 To lastRow
        fund = Trim(wsGL.Cells(i, "E").Value)
        If dict.Exists(fund) Then
            database = dict(fund)
            If Not wbDict.Exists(database) Then
                Set wb = Workbooks.Add
                wbDict.Add database, wb
                ' Copy schema header
                wsGL.Range(wsGL.Cells(1, 1), wsGL.Cells(1, lastCol)).Copy wb.Sheets(1).Range("A1")
                rowCount = 2
            Else
                Set wb = wbDict(database)
                rowCount = wb.Sheets(1).Cells(wb.Sheets(1).Rows.Count, "A").End(xlUp).Row + 1
            End If
            wsGL.Range(wsGL.Cells(i, 1), wsGL.Cells(i, lastCol)).Copy wb.Sheets(1).Cells(rowCount, 1)
        End If
    Next i
    
    ' Save parameterized XML-based XLSX files for downstream ingestion
    For Each key In wbDict.Keys
        Set wb = wbDict(key)
        fileName = "ROC_LOADER_" & Replace(CStr(key), " ", "_") & "_" & todayDate & ".xlsx"
        wb.SaveAs outputFolderPath & fileName, xlOpenXMLWorkbook
        wb.Close SaveChanges:=False
    Next key
    
    ' Clean up memory
    Set dict = Nothing
    Set wbDict = Nothing
    
    Application.ScreenUpdating = True
    
    MsgBox "XLSX loader files have been generated for all fund databases.", vbInformation
End Sub

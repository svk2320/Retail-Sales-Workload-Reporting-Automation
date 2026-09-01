Option Explicit

Public silentMode As Boolean

Sub CleanData()
    Dim wsRaw As Worksheet, wsClean As Worksheet, wsLog As Worksheet
    Dim lastRow As Long, i As Long, outRow As Long
    Dim dictSeen As Object
    Dim rowKey As String
    Dim dupCount As Long, invalidCount As Long, totalIn As Long

    Set wsRaw = ThisWorkbook.Sheets("Raw_Data")
    Set wsClean = ThisWorkbook.Sheets("Clean_Data")
    Set wsLog = ThisWorkbook.Sheets("Cleaning_Log")
    Set dictSeen = CreateObject("Scripting.Dictionary")

    lastRow = wsRaw.Cells(wsRaw.Rows.count, "A").End(xlUp).Row
    totalIn = lastRow - 1 ' minus header

    ' Clear Clean_Data and rewrite header + Is_Valid column
    wsClean.Cells.Clear
    wsRaw.Range("A1:N1").Copy wsClean.Range("A1")
    wsClean.Range("O1").Value = "Is_Valid"

    outRow = 2
    dupCount = 0
    invalidCount = 0
    Dim rollupCount As Long
    rollupCount = 0

    For i = 2 To lastRow
        Dim monthYear As Variant, storeID As Variant, turnover As Variant
        Dim deptName As String, cityName As String

        monthYear = wsRaw.Cells(i, 1).Value
        storeID = wsRaw.Cells(i, 4).Value
        turnover = wsRaw.Cells(i, 11).Value

        ' Build a key to detect exact duplicate rows (all 14 cols)
        rowKey = ""
        Dim c As Integer
        For c = 1 To 14
            rowKey = rowKey & "|" & CStr(wsRaw.Cells(i, c).Value)
        Next c

        If dictSeen.Exists(rowKey) Then
            dupCount = dupCount + 1
            ' skip this row entirely -- duplicate
        ElseIf LCase(Trim(CStr(wsRaw.Cells(i, 7).Value))) = "all" Then
            rollupCount = rollupCount + 1
            ' skip this row entirely -- store-level rollup, not a real department
        Else
            dictSeen.Add rowKey, True

            ' Copy row across
            For c = 1 To 14
                wsClean.Cells(outRow, c).Value = wsRaw.Cells(i, c).Value
            Next c

            ' Standardize MonthYear to MM.YYYY format
            wsClean.Cells(outRow, 1).NumberFormat = "@"  ' force Text, stop Excel auto-converting to a date
            wsClean.Cells(outRow, 1).Value = StandardizeMonthYear(monthYear)

            ' Trim + normalize case on Dept. Name (col 7) and City (col 5)
            wsClean.Cells(outRow, 7).Value = ProperCaseTrim(CStr(wsRaw.Cells(i, 7).Value))
            wsClean.Cells(outRow, 5).Value = ProperCaseTrim(CStr(wsRaw.Cells(i, 5).Value))

            ' Flag invalid: blank StoreID or blank Turnover
            If (storeID = "" Or IsEmpty(storeID)) Or (turnover = "" Or IsEmpty(turnover)) Then
                wsClean.Cells(outRow, 15).Value = "INVALID"
                invalidCount = invalidCount + 1
            Else
                wsClean.Cells(outRow, 15).Value = "OK"
            End If

            outRow = outRow + 1
        End If
    Next i

    ' Write Cleaning_Log
    wsLog.Cells.Clear
    wsLog.Range("A1").Value = "CleanData Run Log"
    wsLog.Range("A3").Value = "Run Timestamp:"
    wsLog.Range("B3").Value = Now
    wsLog.Range("A4").Value = "Total rows in (Raw_Data):"
    wsLog.Range("B4").Value = totalIn
    wsLog.Range("A5").Value = "Duplicate rows removed:"
    wsLog.Range("B5").Value = dupCount
    wsLog.Range("A6").Value = "Rows flagged INVALID (blank StoreID/Turnover):"
    wsLog.Range("B6").Value = invalidCount
    wsLog.Range("A7").Value = "Rollup rows excluded (Dept. Name = 'All'):"
    wsLog.Range("B7").Value = rollupCount
    wsLog.Range("A8").Value = "Total rows out (Clean_Data):"
    wsLog.Range("B8").Value = outRow - 2
    
    Call RefreshLookups

    If Not silentMode Then
        MsgBox "CleanData complete." & vbCrLf & _
           "In: " & totalIn & " | Duplicates removed: " & dupCount & _
           " | Rollups excluded: " & rollupCount & _
           " | Flagged invalid: " & invalidCount & " | Out: " & (outRow - 2), vbInformation
    End If

End Sub

Function StandardizeMonthYear(val As Variant) As String
    Dim months As Variant
    months = Array("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")

    Dim v As String

    ' If Excel already silently converted this to a true Date type (not text),
    ' pull Month/Year directly. This only fires for actual Date variables,
    ' not for text that merely LOOKS like a date to VBA's IsDate.
    If VarType(val) = vbDate Then
        Dim d As Date
        d = CDate(val)
        StandardizeMonthYear = Format(Month(d), "00") & "." & Format(Year(d), "0000")
        Exit Function
    End If

    v = Trim(CStr(val))

    ' Already correct format MM.YYYY
    If v Like "##.####" Then
        StandardizeMonthYear = v
        Exit Function
    End If

    ' ISO format YYYY-MM
    If v Like "####-##" Then
        StandardizeMonthYear = Mid(v, 6, 2) & "." & Left(v, 4)
        Exit Function
    End If

    ' Mon-YY format e.g. "Apr-17" -- MUST be checked as TEXT PATTERN before
    ' any IsDate() check, because VBA's IsDate misreads "Apr-17" as
    ' "April 17, [current year]" instead of "April 2017"
    Dim parts() As String
    If InStr(v, "-") > 0 Then
        parts = Split(v, "-")
        If UBound(parts) = 1 Then
            If Len(parts(0)) = 3 And Len(parts(1)) = 2 Then
                Dim mIdx As Integer
                For mIdx = 0 To 11
                    If LCase(parts(0)) = months(mIdx) Then
                        StandardizeMonthYear = Format(mIdx + 1, "00") & ".20" & parts(1)
                        Exit Function
                    End If
                Next mIdx
            End If
        End If
    End If

    ' Last resort: if it's genuinely an unambiguous real date value, use it
    If IsDate(v) Then
        Dim d2 As Date
        d2 = CDate(v)
        StandardizeMonthYear = Format(Month(d2), "00") & "." & Format(Year(d2), "0000")
        Exit Function
    End If

    ' Fallback: couldn't parse, return as-is (flagged for manual review)
    StandardizeMonthYear = v
End Function

Function ProperCaseTrim(val As String) As String
    Dim v As String
    v = Trim(val)
    If v = "" Then
        ProperCaseTrim = v
        Exit Function
    End If
    ProperCaseTrim = UCase(Left(v, 1)) & LCase(Mid(v, 2))
End Function

Sub RefreshLookups()
    Dim wsClean As Worksheet, wsMaster As Worksheet
    Dim lastRow As Long, lastMaster As Long, i As Long
    Dim xlookupSupported As Boolean

    Set wsClean = ThisWorkbook.Sheets("Clean_Data")
    Set wsMaster = ThisWorkbook.Sheets("Store_Master")

    lastRow = wsClean.Cells(wsClean.Rows.count, "A").End(xlUp).Row
    lastMaster = wsMaster.Cells(wsMaster.Rows.count, "A").End(xlUp).Row

    ' Detect XLOOKUP support by testing it on a scratch cell
    xlookupSupported = True
    On Error Resume Next
    wsClean.Range("Z1").Formula = "=XLOOKUP(1,{1,2},{3,4})"
    If wsClean.Range("Z1").Value <> 3 Then xlookupSupported = False
    If Err.Number <> 0 Then xlookupSupported = False
    On Error GoTo 0
    wsClean.Range("Z1").ClearContents

    ' Header row for the 3 lookup columns
    wsClean.Range("P1").Value = "Store_Name_Lookup"
    wsClean.Range("Q1").Value = "Region_VLOOKUP"
    wsClean.Range("R1").Value = "Scheme_INDEXMATCH"

    If xlookupSupported Then
        wsClean.Range("P1").Value = "Store_Name_XLOOKUP"
    Else
        wsClean.Range("P1").Value = "Store_Name_INDEXMATCH"
    End If

    Dim masterRange As String
    masterRange = "Store_Master!$A$2:$D$" & lastMaster

    For i = 2 To lastRow
        Dim storeIDCell As String
        storeIDCell = "D" & i  ' StoreID is column D in Clean_Data

        If xlookupSupported Then
            wsClean.Range("P" & i).Formula = _
                "=XLOOKUP(" & storeIDCell & ",Store_Master!$A$2:$A$" & lastMaster & _
                ",Store_Master!$B$2:$B$" & lastMaster & ",""Not Found"")"
        Else
            wsClean.Range("P" & i).Formula = _
                "=IFERROR(INDEX(Store_Master!$B$2:$B$" & lastMaster & _
                ",MATCH(" & storeIDCell & ",Store_Master!$A$2:$A$" & lastMaster & ",0)),""Not Found"")"
        End If

        ' VLOOKUP: pull Region (column C of Store_Master, 3rd col in range)
        wsClean.Range("Q" & i).Formula = _
            "=IFERROR(VLOOKUP(" & storeIDCell & "," & masterRange & ",3,FALSE),""Not Found"")"

        ' INDEX-MATCH: pull Scheme (column D of Store_Master)
        wsClean.Range("R" & i).Formula = _
            "=IFERROR(INDEX(Store_Master!$D$2:$D$" & lastMaster & _
            ",MATCH(" & storeIDCell & ",Store_Master!$A$2:$A$" & lastMaster & ",0)),""Not Found"")"
    Next i

    If Not xlookupSupported And Not silentMode Then
        MsgBox "Note: XLOOKUP is not available in this Excel version." & vbCrLf & _
               "Store_Name_Lookup column used INDEX-MATCH instead.", vbInformation
    End If

End Sub

Sub ApplyFormatting()
    Dim wsClean As Worksheet, wsAlerts As Worksheet
    Dim lastRow As Long, lastCol As Long

    Set wsClean = ThisWorkbook.Sheets("Clean_Data")
    Set wsAlerts = ThisWorkbook.Sheets("Alerts")

    lastRow = wsClean.Cells(wsClean.Rows.count, "A").End(xlUp).Row
    lastCol = wsClean.Cells(1, wsClean.Columns.count).End(xlToLeft).Column

    ' Clear old Alerts content + formatting, then mirror Clean_Data (values only)
    wsAlerts.Cells.Clear
    wsClean.Range(wsClean.Cells(1, 1), wsClean.Cells(lastRow, lastCol)).Copy
    wsAlerts.Range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ' Remove any old conditional formatting before reapplying
    wsAlerts.Cells.FormatConditions.Delete

    Dim dataRange As Range
    Set dataRange = wsAlerts.Range(wsAlerts.Cells(2, 1), wsAlerts.Cells(lastRow, lastCol))

    ' Rule 1: Low turnover (< 500000) -- highlight the Turnover cell (col K) orange
    Dim turnoverRange As Range
    Set turnoverRange = wsAlerts.Range(wsAlerts.Cells(2, 11), wsAlerts.Cells(lastRow, 11))
    With turnoverRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="500000")
        .Interior.Color = RGB(255, 199, 132) ' light orange
    End With

    ' Rule 2: Blank/zero Customer count (col L) -- highlight yellow
    Dim customerRange As Range
    Set customerRange = wsAlerts.Range(wsAlerts.Cells(2, 12), wsAlerts.Cells(lastRow, 12))
    With customerRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
        .Interior.Color = RGB(255, 242, 150) ' light yellow
    End With

    ' Rule 3: INVALID rows (col O, Is_Valid) -- highlight whole row light red
    Dim wholeRowRange As Range
    Set wholeRowRange = wsAlerts.Range(wsAlerts.Cells(2, 1), wsAlerts.Cells(lastRow, lastCol))
    With wholeRowRange.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$O2=""INVALID""")
        .Interior.Color = RGB(255, 199, 199) ' light red
    End With

    If Not silentMode Then
        MsgBox "ApplyFormatting complete. Alerts sheet refreshed with " & (lastRow - 1) & " rows.", vbInformation
    End If
    
End Sub


Sub RefreshPivots()
    Dim pt As PivotTable
    Dim ws As Worksheet
    Dim count As Integer
    count = 0

    Set ws = ThisWorkbook.Sheets("Reports_Pivot")

    For Each pt In ws.PivotTables
        pt.RefreshTable
        count = count + 1
    Next pt

    If Not silentMode Then
        MsgBox "RefreshPivots complete. " & count & " PivotTable(s) refreshed.", vbInformation
    End If
    
End Sub

Sub GenerateSummary()
    Dim wsSum As Worksheet, wsPivot As Worksheet
    Dim r As Long

    Set wsSum = ThisWorkbook.Sheets("Summary_Export")
    Set wsPivot = ThisWorkbook.Sheets("Reports_Pivot")

    wsSum.Cells.Clear

    ' Title
    wsSum.Range("A1").Value = "Weekly Sales Summary"
    wsSum.Range("A1").Font.Size = 16
    wsSum.Range("A1").Font.Bold = True
    wsSum.Range("A2").Value = "Generated: " & Format(Now, "dd-mmm-yyyy hh:mm")

    ' KPI section (copy values from Reports_Pivot rows 24-27)
    wsSum.Range("A4").Value = "Key Figures"
    wsSum.Range("A4").Font.Bold = True
    wsPivot.Range("A25:B27").Copy
    wsSum.Range("A5").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ' Pivot 1: Turnover by Country (values only, from A3:B15 area)
    wsSum.Range("A9").Value = "Turnover by Country"
    wsSum.Range("A9").Font.Bold = True
    wsPivot.Range("A3:B15").Copy
    wsSum.Range("A10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ' Pivot 2: Top Departments by Turnover (from D3:E21 area)
    wsSum.Range("D9").Value = "Top Departments by Turnover"
    wsSum.Range("D9").Font.Bold = True
    wsPivot.Range("D3:E21").Copy
    wsSum.Range("D10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ' Pivot 3: Month-over-Month Trend (from G3:H8 area)
    wsSum.Range("G9").Value = "Month-over-Month Trend"
    wsSum.Range("G9").Font.Bold = True
    wsPivot.Range("G3:H8").Copy
    wsSum.Range("G10").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    wsSum.Columns("A:H").AutoFit

    If Not silentMode Then
        MsgBox "GenerateSummary complete. Summary_Export sheet refreshed.", vbInformation
    End If
    
End Sub

Sub ExportToPDF()
    Dim wsSum As Worksheet
    Dim filePath As String
    Dim fileName As String

    Set wsSum = ThisWorkbook.Sheets("Summary_Export")

    fileName = "Sales_Summary_" & Format(Now, "yyyy-mm-dd") & ".pdf"
    filePath = ThisWorkbook.Path & "\" & fileName

    wsSum.ExportAsFixedFormat Type:=xlTypePDF, fileName:=filePath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False

    If Not silentMode Then
        MsgBox "ExportToPDF complete." & vbCrLf & "Saved to: " & filePath, vbInformation
    End If
    
End Sub

Sub RunFullReport()
    Dim startTime As Double
    startTime = Timer

    silentMode = True

    Call CleanData
    Call RefreshPivots
    Call ApplyFormatting
    Call GenerateSummary
    Call ExportToPDF

    silentMode = False

    Dim elapsed As Double
    elapsed = Timer - startTime

    MsgBox "RunFullReport complete!" & vbCrLf & _
           "Total time: " & Format(elapsed, "0.0") & " seconds.", vbInformation
End Sub

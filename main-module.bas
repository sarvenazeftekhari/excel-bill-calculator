Attribute VB_Name = "Module1"
Option Explicit

' -----------------------------
' GenerateBillSplit
' -----------------------------

Sub GenerateBillSplit()
    Dim wsTen As Worksheet, wsBills As Worksheet
    Dim wsRes As Worksheet
    Dim tenants As Collection, bills As Collection
    Dim timestamp As String
    
    On Error GoTo ErrHandler
    Set wsTen = ThisWorkbook.Worksheets("Tenants")
    Set wsBills = ThisWorkbook.Worksheets("Bills")
    
    ' Read data
    Set tenants = ReadTenants(wsTen)
    Set bills = ReadBills(wsBills)
    
    If tenants.Count = 0 Then
        MsgBox "No tenants found on sheet 'Tenants'. Please enter tenants and try again.", vbExclamation
        Exit Sub
    End If
    If bills.Count = 0 Then
        MsgBox "No bills found on sheet 'Bills'. Please enter bills and try again.", vbExclamation
        Exit Sub
    End If
    
    ' Create result sheet with timestamp so we don't overwrite previous results
    timestamp = Format(Now, "yyyy-mm-dd_HHMMSS")
    Set wsRes = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsRes.Name = "Result_" & timestamp
    
    ' Calculate shares and write result table
    WriteResults wsRes, tenants, bills
    
    MsgBox "Result sheet created: " & wsRes.Name, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' -----------------------------
' Read tenants from Tenants sheet
' expects columns: A=Tenant, B=Status, C=Start date, D=End date
' returns Collection of Dictionary objects with keys: Name, StartDate (Date), EndDate (Variant: Date or Empty)
' -----------------------------
Function ReadTenants(ws As Worksheet) As Collection
    Dim col As New Collection
    Dim row As Long
    Dim nm As String, sDate As String, eDate As String
    Dim dict As Object
    
    row = 2 ' assume headers in row 1
    Do While Trim(CStr(ws.Cells(row, "A").Value)) <> ""
        nm = Trim(CStr(ws.Cells(row, "A").Value))
        sDate = Trim(CStr(ws.Cells(row, "C").Value))
        eDate = Trim(CStr(ws.Cells(row, "D").Value))
        
        Set dict = CreateObject("Scripting.Dictionary")
        dict("Name") = nm
        If sDate = "" Then
            ' if start date missing, use a very early date to avoid problems
            dict("StartDate") = DateSerial(1900, 1, 1)
        Else
            dict("StartDate") = SafeDate(sDate)
        End If
        
        If eDate = "" Then
            dict("EndDate") = Empty ' blank means still here
        Else
            dict("EndDate") = SafeDate(eDate)
        End If
        
        col.Add dict
        row = row + 1
    Loop
    
    Set ReadTenants = col
End Function

' -----------------------------
' Read bills from Bills sheet
' expects columns: A=Bill, B=Amount, C=Start date, D=End date
' returns Collection of Dictionary objects with keys: Name, Amount, StartDate, EndDate
' -----------------------------
Function ReadBills(ws As Worksheet) As Collection
    Dim col As New Collection
    Dim row As Long
    Dim nm As String, sAmt As Variant, sDate As String, eDate As String
    Dim dict As Object
    
    row = 2 ' assume headers in row 1
    Do While Trim(CStr(ws.Cells(row, "A").Value)) <> ""
        nm = Trim(CStr(ws.Cells(row, "A").Value))
        sAmt = ws.Cells(row, "B").Value
        sDate = Trim(CStr(ws.Cells(row, "C").Value))
        eDate = Trim(CStr(ws.Cells(row, "D").Value))
        
        If nm <> "" Then
            Set dict = CreateObject("Scripting.Dictionary")
            dict("Name") = nm
            dict("Amount") = CDbl(sAmt)
            dict("StartDate") = SafeDate(sDate)
            dict("EndDate") = SafeDate(eDate)
            col.Add dict
        End If
        row = row + 1
    Loop
    
    Set ReadBills = col
End Function

' -----------------------------
' Convert a value into a Date safely
' Accepts either Excel date or a string, if invalid it returns a sentinel large date
' -----------------------------
Function SafeDate(val As Variant) As Date
    If IsDate(val) Then
        SafeDate = CDate(val)
    ElseIf Trim(CStr(val)) = "" Then
        ' sentinel far future date for empty inputs when calling code expects a Date
        SafeDate = DateSerial(9999, 12, 31)
    Else
        On Error Resume Next
        SafeDate = CDate(val)
        If Err.Number <> 0 Then
            Err.Clear
            SafeDate = DateSerial(9999, 12, 31)
        End If
        On Error GoTo 0
    End If
End Function

' -----------------------------
' WriteResults
' Core logic to compute per-bill per-tenant shares and dump to result sheet
' -----------------------------
Sub WriteResults(ws As Worksheet, tenants As Collection, bills As Collection)
    Dim i As Long, j As Long
    Dim t As Object, b As Object
    Dim tenantNames() As String
    Dim tenantCount As Long
    Dim tenantOverallTotals() As Double
    Dim billCount As Long
    Dim participatingTenantFlags() As Boolean
    Dim sheetRow As Long, sheetCol As Long
    
    billCount = bills.Count
    tenantCount = tenants.Count
    ReDim tenantNames(1 To tenantCount)
    ReDim tenantOverallTotals(1 To tenantCount)
    ReDim participatingTenantFlags(1 To tenantCount)
    
    ' Use Variant array to store per-bill arrays (each element will hold an array of doubles)
    Dim perBillTenantShares() As Variant
    ReDim perBillTenantShares(1 To billCount)
    
    ' initialize
    For i = 1 To tenantCount
        tenantNames(i) = tenants(i)("Name")
        tenantOverallTotals(i) = 0#
        participatingTenantFlags(i) = False
    Next i
    
    ' For each bill compute tenant-days and shares
    For i = 1 To billCount
        Set b = bills(i)
        Dim billStart As Date, billEnd As Date, billAmt As Double
        billStart = b("StartDate")
        billEnd = b("EndDate")
        billAmt = b("Amount")
        
        Dim tenantDays() As Double
        ReDim tenantDays(1 To tenantCount)
        Dim totalTenantDays As Double
        totalTenantDays = 0#
        
        ' compute overlap days (inclusive)
        For j = 1 To tenantCount
            Set t = tenants(j)
            Dim tStart As Date, tEnd As Variant
            tStart = t("StartDate")
            tEnd = t("EndDate") ' may be Empty or a date
            
            Dim overlapStart As Date, overlapEnd As Date
            ' overlapStart = max(tStart, billStart)
            If tStart > billStart Then
                overlapStart = tStart
            Else
                overlapStart = billStart
            End If
            
            ' overlapEnd = min(tEnd (or billEnd if empty), billEnd)
            If IsEmpty(tEnd) Then
                overlapEnd = billEnd
            Else
                If CDate(tEnd) < billEnd Then
                    overlapEnd = CDate(tEnd)
                Else
                    overlapEnd = billEnd
                End If
            End If
            
            If overlapStart <= overlapEnd Then
                tenantDays(j) = DateDiff("d", overlapStart, overlapEnd) + 1 ' inclusive
            Else
                tenantDays(j) = 0
            End If
            
            totalTenantDays = totalTenantDays + tenantDays(j)
        Next j
        
        ' allocate shares
        Dim shares() As Double
        ReDim shares(1 To tenantCount)
        
        If totalTenantDays > 0 Then
            Dim sumAssigned As Double
            sumAssigned = 0#
            Dim lastIdx As Long
            lastIdx = 0
            ' find last participating tenant index to assign rounding remainder
            For j = tenantCount To 1 Step -1
                If tenantDays(j) > 0 Then
                    lastIdx = j
                    Exit For
                End If
            Next j
            
            For j = 1 To tenantCount
                If tenantDays(j) > 0 Then
                    If j <> lastIdx Then
                        shares(j) = Round((tenantDays(j) / totalTenantDays) * billAmt, 2)
                        sumAssigned = sumAssigned + shares(j)
                    Else
                        ' last participating tenant gets remainder so row sums exactly to billAmt
                        shares(j) = Round(billAmt - sumAssigned, 2)
                    End If
                    tenantOverallTotals(j) = tenantOverallTotals(j) + shares(j)
                    participatingTenantFlags(j) = participatingTenantFlags(j) Or (shares(j) <> 0)
                Else
                    shares(j) = 0#
                End If
            Next j
        Else
            ' no tenant present during bill, all zero
            For j = 1 To tenantCount
                shares(j) = 0#
            Next j
        End If
        
        perBillTenantShares(i) = shares ' store the per-bill array inside the variant array
    Next i
    
    ' Determine list of tenants to show (those with non-zero overall totals)
    Dim shownIdx() As Long
    Dim shownCount As Long
    ReDim shownIdx(1 To tenantCount)
    shownCount = 0
    For j = 1 To tenantCount
        If Abs(tenantOverallTotals(j)) > 0.00001 Then
            shownCount = shownCount + 1
            shownIdx(shownCount) = j
        End If
    Next j
    
    If shownCount = 0 Then
        MsgBox "No tenant owes anything for the provided bills (no overlaps).", vbInformation
        Exit Sub
    End If
    
    ' --- Write header row ---
    ws.Cells.Clear
    ws.Range("A1").Value = "" ' top-left
    For j = 1 To shownCount
        ws.Cells(1, 1 + j).Value = tenantNames(shownIdx(j))
    Next j
    ws.Cells(1, 1 + shownCount + 1).Value = "Total"
    
    ' --- Write each bill row ---
    For i = 1 To billCount
        ws.Cells(i + 1, 1).Value = bills(i)("Name")
        Dim rowTotal As Double
        rowTotal = 0#
        
        For j = 1 To shownCount
            Dim idx As Long
            idx = shownIdx(j)
            Dim val As Double
            val = perBillTenantShares(i)(idx) ' perBillTenantShares(i) is the shares() array
            ws.Cells(i + 1, 1 + j).Value = val
            rowTotal = rowTotal + val
        Next j
        ws.Cells(i + 1, 1 + shownCount + 1).Value = rowTotal ' should equal bill amount (mod rounding)
    Next i
    
    ' --- Totals row ---
    ws.Cells(billCount + 2, 1).Value = "Total"
    Dim grandTotal As Double
    grandTotal = 0#
    For j = 1 To shownCount
        Dim colSum As Double
        colSum = 0#
        For i = 1 To billCount
            colSum = colSum + perBillTenantShares(i)(shownIdx(j))
        Next i
        ws.Cells(billCount + 2, 1 + j).Value = colSum
        grandTotal = grandTotal + colSum
    Next j
    ws.Cells(billCount + 2, 1 + shownCount + 1).Value = grandTotal
    
    ' --- Formatting ---
    Dim lastCol As Long, lastRow As Long
    lastRow = billCount + 2
    lastCol = 1 + shownCount + 1
    
    ' Bold headers
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Font.Bold = True
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 1)).Font.Bold = True
    
    ' Format currency for the table area (exclude column A which is text)
    ws.Range(ws.Cells(2, 2), ws.Cells(lastRow, lastCol)).NumberFormat = "$#,##0.00"
    
    ' Adjust column widths
    ws.Columns.AutoFit
    
    ' Freeze top row
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

Sub UpdateClientTracker()
    Dim wsMaster As Worksheet, wsClient As Worksheet
    Dim wbMaster As Workbook
    Dim lastRowMaster As Long, lastRowClient As Long, newRow As Long
    Dim dict As Object
    Dim i As Long, j As Long
    Dim key As String
    Dim recordDate As Date
    Dim todayDate As Date, prevDate As Date
    
    ' Open Master Billing Tracker
    On Error Resume Next
    Set wbMaster = Workbooks.Open("YourMasterBillingTrackerFilePath")
    If wbMaster Is Nothing Then
        MsgBox "Error: Unable to open Master Billing Tracker.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Set worksheets
    Set wsMaster = wbMaster.Sheets("Sheet1") ' Master Tracker
    Set wsClient = ThisWorkbook.Sheets("Details") ' Client Tracker (Active Workbook)
    
    ' Get last row in both sheets
    lastRowMaster = wsMaster.Cells(wsMaster.Rows.Count, 1).End(xlUp).Row
    lastRowClient = wsClient.Cells(wsClient.Rows.Count, 1).End(xlUp).Row
    newRow = lastRowClient + 1 ' Start adding new records from next available row
    
    ' Set today's and previous day's date
    todayDate = Date
    prevDate = Date - 1
    
    ' Create dictionary to store Master Tracker data
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Loop through Master Tracker and store data
    For i = 2 To lastRowMaster ' Assuming row 1 is headers
        key = wsMaster.Cells(i, 15).Value & "|" & wsMaster.Cells(i, 16).Value ' X VISIT NO & Y VISIT NO
        recordDate = wsMaster.Cells(i, 19).Value ' Using Billed Date from Column S
        
        If Not dict.exists(key) Then
            dict.Add key, i ' Store row number
        End If
        
        ' Add new records for today and yesterday
        If recordDate = todayDate Or recordDate = prevDate Then
            wsClient.Cells(newRow, 1).Value = wsMaster.Cells(i, 15).Value ' X VISIT NO
            wsClient.Cells(newRow, 2).Value = wsMaster.Cells(i, 16).Value ' Y VISIT NO
            wsClient.Cells(newRow, 3).Value = wsMaster.Cells(i, 2).Value  ' Accession #
            wsClient.Cells(newRow, 4).Value = wsMaster.Cells(i, 3).Value  ' (F) Name
            wsClient.Cells(newRow, 5).Value = wsMaster.Cells(i, 4).Value  ' (L) Name
            wsClient.Cells(newRow, 6).Value = wsMaster.Cells(i, 5).Value  ' Full Name
            wsClient.Cells(newRow, 7).Value = wsMaster.Cells(i, 6).Value  ' DOB
            wsClient.Cells(newRow, 8).Value = wsMaster.Cells(i, 8).Value  ' DOS
            wsClient.Cells(newRow, 9).Value = wsMaster.Cells(i, 10).Value ' Facility
            wsClient.Cells(newRow, 10).Value = wsMaster.Cells(i, 11).Value ' Type
            wsClient.Cells(newRow, 11).Value = wsMaster.Cells(i, 12).Value ' Insurance Provider
            wsClient.Cells(newRow, 12).Value = wsMaster.Cells(i, 13).Value ' Insurance ID
            
            ' Update Billing Status
            Select Case wsMaster.Cells(i, 17).Value
                Case "COMPLETED", "CIP"
                    wsClient.Cells(newRow, 13).Value = "Entered to AMD"
                Case "REJECTED", "Escalated"
                    wsClient.Cells(newRow, 13).Value = "Not Entered to AMD"
                Case Else
                    wsClient.Cells(newRow, 13).Value = "Pending"
            End Select
            
            newRow = newRow + 1
        End If
    Next i
    
    ' Save and notify user
    wbMaster.Close False
    MsgBox "Client Tracker Updated with New Data!", vbInformation
End Sub

Sub UpdateBillingReconciliation()
    Dim wbMaster As Workbook, wbRecon As Workbook
    Dim wsMaster As Worksheet, wsRecon As Worksheet
    Dim lastRow As Long, reconRow As Long, i As Integer
    Dim latestDate As Date, targetDate As Date
    Dim totalClaims As Long, bloodCount As Long, stiCount As Long, utiCount As Long, gastroCount As Long
    Dim manualCount As Long, pendingCount As Long, duplicateCount As Long
    Dim escalationCount As Long, cipCount As Long, rejectedCount As Long
    Dim masterFilePath As String, reconFilePath As String
    
    ' Prompt for Master Billing Tracker file path
    masterFilePath = Application.GetOpenFilename("Excel Files (*.xlsm), *.xlsm", , "Select Master Billing Tracker")
    If masterFilePath = "False" Then Exit Sub
    
    ' Prompt for Billing Reconciliation file path
    reconFilePath = Application.GetOpenFilename("Excel Files (*.xlsm), *.xlsm", , "Select Billing Reconciliation File")
    If reconFilePath = "False" Then Exit Sub
    
    ' Open Master Billing Tracker
    Set wbMaster = Workbooks.Open(masterFilePath)
    Set wsMaster = wbMaster.Sheets(1) ' Use appropriate sheet name or index
    
    ' Open Billing Reconciliation file
    Set wbRecon = Workbooks.Open(reconFilePath)
    Set wsRecon = wbRecon.Sheets("Reconciliation Start to Date") ' Use appropriate sheet name
    
    ' Find the last row in Master Tracker
    lastRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row
    
    ' Find latest received date
    latestDate = Application.WorksheetFunction.Max(wsMaster.Range("A2:A" & lastRow))
    
    ' Loop through the last 10 days including latest
    For i = 0 To 9
        targetDate = latestDate - i
        
        ' Skip weekends (Saturday = 7, Sunday = 1)
        If Weekday(targetDate, vbMonday) > 5 Then
            ' If it's Saturday or Sunday, skip this iteration
            GoTo SkipDate
        End If
        
        ' Find the last used row in Billing Reconciliation to determine next row
        reconRow = wsRecon.Cells(wsRecon.Rows.Count, "B").End(xlUp).Row + 1
        If reconRow < 215 Then reconRow = 215
        
        ' Count total claims for the target date
        totalClaims = Application.WorksheetFunction.CountIf(wsMaster.Range("A2:A" & lastRow), targetDate)
        
        ' Count claims by type (Column K)
        bloodCount = Application.WorksheetFunction.CountIfs(wsMaster.Range("A2:A" & lastRow), targetDate, wsMaster.Range("K2:K" & lastRow), "Blood")
        stiCount = Application.WorksheetFunction.CountIfs(wsMaster.Range("A2:A" & lastRow), targetDate, wsMaster.Range("K2:K" & lastRow), "STI")
        utiCount = Application.WorksheetFunction.CountIfs(wsMaster.Range("A2:A" & lastRow), targetDate, wsMaster.Range("K2:K" & lastRow), "UTI")
        gastroCount = Application.WorksheetFunction.CountIfs(wsMaster.Range("A2:A" & lastRow), targetDate, wsMaster.Range("K2:K" & lastRow), "Gastro")
        
        ' Count claims by billing status (Column Q)
        manualCount = Application.WorksheetFunction.CountIfs(wsMaster.Range("A2:A" & lastRow), targetDate, wsMaster.Range("Q2:Q" & lastRow), "COMPLETED")
        pendingCount = Application.WorksheetFunction.CountIfs(wsMaster.Range("A2:A" & lastRow), targetDate, wsMaster.Range("Q2:Q" & lastRow), "")
        duplicateCount = Application.WorksheetFunction.CountIfs(wsMaster.Range("A2:A" & lastRow), targetDate, wsMaster.Range("Q2:Q" & lastRow), "Duplicate")
        escalationCount = Application.WorksheetFunction.CountIfs(wsMaster.Range("A2:A" & lastRow), targetDate, wsMaster.Range("Q2:Q" & lastRow), "Escalated")
        cipCount = Application.WorksheetFunction.CountIfs(wsMaster.Range("A2:A" & lastRow), targetDate, wsMaster.Range("Q2:Q" & lastRow), "CIP")
        rejectedCount = Application.WorksheetFunction.CountIfs(wsMaster.Range("A2:A" & lastRow), targetDate, wsMaster.Range("Q2:Q" & lastRow), "Rejected")
        
        ' Update the reconciliation sheet
        wsRecon.Cells(reconRow, 2).Value = targetDate
        wsRecon.Cells(reconRow, 3).Value = totalClaims
        wsRecon.Cells(reconRow, 4).Value = bloodCount
        wsRecon.Cells(reconRow, 5).Value = stiCount
        wsRecon.Cells(reconRow, 6).Value = utiCount
        wsRecon.Cells(reconRow, 7).Value = gastroCount
        wsRecon.Cells(reconRow, 8).Value = manualCount
        wsRecon.Cells(reconRow, 9).Value = pendingCount
        wsRecon.Cells(reconRow, 10).Value = duplicateCount
        wsRecon.Cells(reconRow, 11).Value = escalationCount
        wsRecon.Cells(reconRow, 12).Value = cipCount
        wsRecon.Cells(reconRow, 13).Value = rejectedCount
        
SkipDate:
    Next i
    
    ' Save workbooks but do not close
    wbRecon.Save
    wbMaster.Save
    
    ' Notify the user
    MsgBox "Billing Reconciliation updated successfully for the last 10 days starting from row " & reconRow & "!", vbInformation, "Success"
End Sub

Sub UpdateTracker()
    Dim wsLIS As Worksheet, wsTracker As Worksheet
    Dim lastRowLIS As Long, lastRowTracker As Long
    Dim rngLIS As Range, rngTracker As Range, cell As Range
    Dim missingClaims As Object
    
    ' Open LIS Report (User selects file)
    Dim lisFile As String
    lisFile = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Select LIS Report")
    If lisFile = "False" Then Exit Sub ' If the user cancels the file selection, exit the subroutine
    
    Workbooks.Open lisFile
    Set wsLIS = ActiveWorkbook.Sheets(1) ' Select the first sheet of the LIS Report
    
    ' Set Tracker Workbook and Sheet (Current workbook)
    Set wsTracker = ThisWorkbook.Sheets(1) ' Select the first sheet in the Tracker workbook
    
    ' Find Last Rows in both worksheets
    lastRowLIS = wsLIS.Cells(Rows.Count, 1).End(xlUp).Row ' Last row in LIS Report based on column A
    lastRowTracker = wsTracker.Cells(Rows.Count, 1).End(xlUp).Row ' Last row in Tracker based on column A
    
    ' Define Lookup Ranges (Accession numbers in LIS and Tracker)
    Set rngLIS = wsLIS.Range("A2:A" & lastRowLIS) ' Accession numbers in LIS Report (starting from row 2)
    Set rngTracker = wsTracker.Range("B2:B" & lastRowTracker) ' Accession numbers in Tracker (starting from row 2)
    
    ' Create Dictionary for Fast Lookup of existing claims in the Tracker
    Set missingClaims = CreateObject("Scripting.Dictionary")
    For Each cell In rngTracker
        missingClaims(cell.Value) = 1 ' Add each accession number from Tracker to the dictionary
    Next cell
    
    ' Find Missing Claims in LIS and Add to Tracker
    Dim newRow As Long
    newRow = lastRowTracker + 2 ' Maintain 2-row gap for new entries in Tracker
    For Each cell In rngLIS
        If Not missingClaims.exists(cell.Value) Then ' If the claim is not found in Tracker
            wsTracker.Cells(newRow, 1).Value = Date ' Today's Date as Received Date in Tracker
            wsTracker.Cells(newRow, 2).Value = cell.Value ' Accession No (from LIS Report)
            ' Map other columns from LIS Report to Tracker
            wsTracker.Cells(newRow, 3).Value = wsLIS.Cells(cell.Row, 4).Value ' Example: Column D in LIS to Column 3 in Tracker
            wsTracker.Cells(newRow, 4).Value = wsLIS.Cells(cell.Row, 5).Value ' Example: Column E in LIS to Column 4 in Tracker
            wsTracker.Cells(newRow, 5).Value = wsTracker.Cells(newRow, 4).Value & ", " & wsTracker.Cells(newRow, 3).Value ' Example: Full Name (Concatenate)
            wsTracker.Cells(newRow, 6).Value = wsLIS.Cells(cell.Row, 6).Value ' Example: Column F in LIS to Column 6 in Tracker
            ' Continue adding other necessary columns from LIS Report to Tracker...
            
            newRow = newRow + 1 ' Increment the row for the next missing claim
        End If
    Next cell
    
    ' Save and Close LIS Report after processing
    ActiveWorkbook.Close False
    
    ' Notify the User that the Tracker is updated
    MsgBox "Tracker Updated with Missing Claims!", vbInformation
End Sub

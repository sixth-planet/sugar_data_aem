Attribute VB_Name = "ConsolidateData"
'==============================================================================
' ConsolidateData Module
' Purpose: Consolidate rows by DATE column, handling duplicate values
' Author: Excel Data Cleaning Solution
' Version: 1.0
'==============================================================================

Option Explicit

' Main subroutine to consolidate data by DATE column
Sub ConsolidateDataByDate()
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dateCol As Long
    Dim i As Long, j As Long, k As Long
    Dim dict As Object
    Dim dateKey As Variant
    Dim dateValue As Date
    Dim outputRow As Long
    Dim colName As String
    Dim cellValue As String
    Dim existingValue As String
    Dim duplicateCount As Integer
    Dim newColName As String
    Dim headerRow As Long
    Dim progressCount As Long
    Dim totalRows As Long
    Dim foundHeader As Boolean
    
    ' Error handling
    On Error GoTo ErrorHandler
    
    ' Initialize
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Set source worksheet (active sheet)
    Set wsSource = ActiveSheet
    
    ' Find DATE column in header row (row 1)
    headerRow = 1
    dateCol = 0
    lastCol = wsSource.Cells(headerRow, wsSource.Columns.Count).End(xlToLeft).Column
    
    ' Limit to column Z (26)
    If lastCol > 26 Then lastCol = 26
    
    ' Find DATE column
    For i = 1 To lastCol
        If UCase(Trim(wsSource.Cells(headerRow, i).Value)) = "DATE" Then
            dateCol = i
            Exit For
        End If
    Next i
    
    ' Validate DATE column exists
    If dateCol = 0 Then
        MsgBox "Error: DATE column not found in the worksheet. Please ensure there is a column named 'DATE'.", vbCritical
        GoTo CleanUp
    End If
    
    ' Find last row with data
    lastRow = wsSource.Cells(wsSource.Rows.Count, dateCol).End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "Error: No data found to consolidate.", vbCritical
        GoTo CleanUp
    End If
    
    ' Display progress
    totalRows = lastRow - headerRow
    MsgBox "Starting consolidation of " & totalRows & " rows..." & vbCrLf & _
           "Columns A-" & Split(Cells(1, lastCol).Address, "$")(1) & " will be processed." & vbCrLf & _
           "This may take a few moments.", vbInformation
    
    ' Create or clear output worksheet
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Worksheets("Cleaned_Data")
    If wsOutput Is Nothing Then
        Set wsOutput = ThisWorkbook.Worksheets.Add
        wsOutput.Name = "Cleaned_Data"
    Else
        wsOutput.Cells.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' Copy headers to output sheet
    For i = 1 To lastCol
        wsOutput.Cells(1, i).Value = wsSource.Cells(headerRow, i).Value
    Next i
    
    ' Create dictionary to store consolidated data
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Process each row
    progressCount = 0
    For i = headerRow + 1 To lastRow
        ' Get date value
        On Error Resume Next
        dateValue = wsSource.Cells(i, dateCol).Value
        If Err.Number <> 0 Then
            Err.Clear
            ' Skip rows with invalid dates
            GoTo NextRow
        End If
        On Error GoTo ErrorHandler
        
        dateKey = CLng(dateValue)
        
        ' Initialize dictionary entry if not exists
        If Not dict.Exists(dateKey) Then
            dict.Add dateKey, CreateObject("Scripting.Dictionary")
            dict(dateKey)("_date") = dateValue
        End If
        
        ' Process each column
        For j = 1 To lastCol
            If j <> dateCol Then ' Skip DATE column itself
                cellValue = Trim(CStr(wsSource.Cells(i, j).Value))
                colName = Trim(CStr(wsSource.Cells(headerRow, j).Value))
                
                If cellValue <> "" Then
                    ' Check if this column already has a value for this date
                    If Not dict(dateKey).Exists(colName) Then
                        ' First value for this column/date combination
                        dict(dateKey)(colName) = cellValue
                    Else
                        ' Column already has a value - check if different
                        existingValue = dict(dateKey)(colName)
                        If existingValue <> cellValue Then
                            ' Different value - create additional column
                            duplicateCount = 1
                            newColName = colName & "%"
                            
                            ' Find next available column name
                            Do While dict(dateKey).Exists(newColName)
                                duplicateCount = duplicateCount + 1
                                newColName = colName & "%" & duplicateCount
                            Loop
                            
                            dict(dateKey)(newColName) = cellValue
                            
                            ' Add new column header if not exists
                            foundHeader = False
                            k = 1
                            Do While k <= wsOutput.Cells(1, wsOutput.Columns.Count).End(xlToLeft).Column + 1
                                If wsOutput.Cells(1, k).Value = newColName Then
                                    foundHeader = True
                                    Exit Do
                                End If
                                k = k + 1
                            Loop
                            
                            If Not foundHeader Then
                                lastCol = wsOutput.Cells(1, wsOutput.Columns.Count).End(xlToLeft).Column + 1
                                wsOutput.Cells(1, lastCol).Value = newColName
                            End If
                        End If
                    End If
                End If
            End If
        Next j
        
NextRow:
        ' Update progress every 100 rows
        progressCount = progressCount + 1
        If progressCount Mod 100 = 0 Then
            Application.StatusBar = "Processing row " & progressCount & " of " & totalRows & "..."
        End If
    Next i
    
    ' Write consolidated data to output sheet
    outputRow = 2
    Application.StatusBar = "Writing consolidated data..."
    
    For Each dateKey In dict.Keys
        ' Write date
        wsOutput.Cells(outputRow, dateCol).Value = dict(dateKey)("_date")
        
        ' Write all other values
        For i = 1 To wsOutput.Cells(1, wsOutput.Columns.Count).End(xlToLeft).Column
            colName = wsOutput.Cells(1, i).Value
            If colName <> "" And i <> dateCol Then
                If dict(dateKey).Exists(colName) Then
                    wsOutput.Cells(outputRow, i).Value = dict(dateKey)(colName)
                End If
            End If
        Next i
        
        outputRow = outputRow + 1
    Next dateKey
    
    ' Format output sheet
    With wsOutput
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(200, 200, 200)
        .Columns.AutoFit
    End With
    
    ' Activate output sheet
    wsOutput.Activate
    
    ' Display completion message
    MsgBox "Data consolidation completed successfully!" & vbCrLf & vbCrLf & _
           "Original rows: " & totalRows & vbCrLf & _
           "Consolidated rows: " & (outputRow - 2) & vbCrLf & _
           "Output columns: " & wsOutput.Cells(1, wsOutput.Columns.Count).End(xlToLeft).Column & vbCrLf & vbCrLf & _
           "Results are in the 'Cleaned_Data' sheet.", vbInformation
    
CleanUp:
    ' Clean up
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Set dict = Nothing
    Set wsSource = Nothing
    Set wsOutput = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Error number: " & Err.Number, vbCritical
    Resume CleanUp
End Sub

' Helper function to check if a value is a valid date
Function IsValidDate(value As Variant) As Boolean
    On Error Resume Next
    IsValidDate = IsDate(value)
    If Err.Number <> 0 Then
        IsValidDate = False
        Err.Clear
    End If
End Function

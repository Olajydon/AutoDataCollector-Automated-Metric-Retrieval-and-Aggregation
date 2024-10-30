Sub AutoDataCollector()
    ' Define key objects for workbook and sheet references
    Dim targetWorkbook As Workbook
    Dim dataSource As Worksheet
    Dim currentSheet As Worksheet
    Dim highlightRange As Range
    Dim cell As Range
    
    ' Define metrics array with specific metric names (generic names used here)
    Dim metrics As Variant
    Dim metric As Variant
    Dim metricCell As Range
    Dim valueToFill As Double
    Dim wbFound As Boolean
    Dim rowIndex As Long
    Dim found As Boolean
    Dim searchRange As Range
    Dim tempCell As Range
    Dim search_wordSum As Double
    
    ' Define generic metric names
    metrics = Array("metric_1", "metric_2", "metric_3", "metric_4", "metric_5", "metric_6", "metric_7", "metric_8", "metric_9")

    ' Flag to track if the data source workbook is found
    wbFound = False
    
    ' Search for an open workbook matching the data source file (e.g. genberic "data_source" used here) name
    For Each targetWorkbook In Workbooks
        If Trim(targetWorkbook.Name) Like "*data_source*.xlsm" Then
            wbFound = True
            Exit For
        End If
    Next targetWorkbook

    ' If the data source workbook isn't open, prompt the user to select it
    If Not wbFound Then
        MsgBox "data source workbook not found automatically. Please select the correct workbook.", vbExclamation
        On Error Resume Next
        Set targetWorkbook = Application.InputBox("Select the workbook containing the data source sheet", Type:=8).Parent
        On Error GoTo 0
        If targetWorkbook Is Nothing Then
            MsgBox "No workbook selected. Exiting.", vbExclamation
            Exit Sub
        End If
    End If

    ' Reference to the "data source" sheet in the selected workbook
    On Error Resume Next
    Set dataSource = targetWorkbook.Sheets("data_source")
    On Error GoTo 0

    ' Confirm that the data_source sheet exists
    If dataSource Is Nothing Then
        MsgBox "The 'data_source' sheet was not found in the selected workbook.", vbExclamation
        Exit Sub
    End If

    ' Reference the active sheet where the button was clicked to run the macro
    Set currentSheet = ActiveSheet

    ' Get the range of cells that the user selected to populate data
    On Error Resume Next
    Set highlightRange = Selection
    On Error GoTo 0

    ' Ensure that cells are selected for the macro to populate
    If highlightRange Is Nothing Or highlightRange.Cells.Count = 0 Then
        MsgBox "No cells selected. Please highlight the cells you want to populate.", vbExclamation
        Exit Sub
    End If

    ' Loop through each generic metric in the predefined array
    rowIndex = 0
    For Each cell In highlightRange
        ' Process each metric up to the count defined in metrics array
        If rowIndex < UBound(metrics) + 1 Then
            metric = Trim(metrics(rowIndex))
            valueToFill = 0
            
            ' Define the search range in the data_source sheet
            Set searchRange = dataSource.Range("A2:A") ' Adjust the range as necessary for the actual data
            
            ' Search for each metric in the data source and retrieve corresponding values
            found = False
            For Each metricCell In searchRange
                If UCase(Trim(metricCell.Value)) = UCase(metric) Then
                    ' Fetch the value next to the found metric
                    valueToFill = metricCell.Offset(0, 1).Value
                    found = True
                    Exit For
                End If
            Next metricCell

            ' If the metric is not found in the data source, set its value to zero
            If Not found Then
                valueToFill = 0
            End If

            ' Additional calculation for specific metrics (e.g., "metric_3" and "metric_7") based on "search_wordSum" values
            If metric = "metric_3" Or metric = "metric_7" Then
                search_wordSum = 0 ' Initialize the sum for "search_word" values
                For Each tempCell In searchRange
                    ' Check for entries containing both the metric and "search_word"
                    If (metric = "metric_3" And InStr(1, UCase(tempCell.Value), "metric_3", vbTextCompare) > 0 And InStr(1, UCase(tempCell.Value), "search_word", vbTextCompare) > 0) _
                    Or (metric = "metric_7" And InStr(1, UCase(tempCell.Value), "metric_7", vbTextCompare) > 0 And InStr(1, UCase(tempCell.Value), "search_word", vbTextCompare) > 0) Then
                        search_wordSum = search_wordSum + tempCell.Offset(0, 1).Value
                    End If
                Next tempCell
                ' Add the "search_word" sum to the original value for the metric
                valueToFill = valueToFill + search_wordSum
            End If

            ' Populate the selected cell with the calculated value for the metric
            cell.Value = valueToFill

            ' Increment to the next metric
            rowIndex = rowIndex + 1
        End If
    Next cell

    MsgBox "Data populated successfully!", vbInformation
End Sub

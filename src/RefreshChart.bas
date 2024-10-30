Sub RefreshChart()
    ' Define key objects and variables for sheet, chart, and data ranges
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim rngTarget As Range, rngYTD As Range, rngAdditionalYTD As Range, rngLegends As Range
    Dim rngLatestMonth As Range
    Dim lastMonthRow As Long, i As Integer
    Dim monthLabel As String
    Dim isMonthFilled As Boolean
    Dim isSpecialSheet As Boolean
    Dim chartTopLeft As Range

    ' Set the active sheet as the working sheet
    Set ws = ActiveSheet

    ' Identify if the current sheet is a "special" sheet based on name
    isSpecialSheet = (ws.Name = "SpecialSheet1" Or ws.Name = "SpecialSheet2")

    ' Clear any existing charts on the sheet to prevent duplication
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj

    ' Define common ranges for Target and Legends (these ranges are the same for all sheets)
    Set rngTarget = ws.Range("D6:L6")  'Set desired data range
    Set rngLegends = ws.Range("D4:L4")  'Set desired data range

    ' Define ranges for YTD and Additional YTD based on sheet type
    If isSpecialSheet Then
        ' Specific YTD ranges for "special" sheets
        If ws.Name = "SpecialSheet1" Then
            Set rngYTD = ws.Range("D22:L22")  'Set desired data range
        ElseIf ws.Name = "SpecialSheet2" Then
            Set rngYTD = ws.Range("D23:L23")  'Set desired data range
        End If
        Set rngAdditionalYTD = ws.Range("D24:L24") ' Additional YTD for both special sheets (or set desired data range)
        Set chartTopLeft = ws.Range("B27") ' Position chart at B27 for special sheets (or set desired data range)
    Else
        ' Default YTD range for other sheets
        Set rngYTD = ws.Range("D19:L19") 'Set desired data range
        Set chartTopLeft = ws.Range("B23") ' Position chart at B23 for other sheets (or set desired data range)
    End If

    ' Locate the latest filled month by scanning specific rows (adjust as needed for your data range)
    isMonthFilled = False
    For i = 18 To 7 Step -1
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, 4), ws.Cells(i, 12))) > 0 Then
            lastMonthRow = i
            isMonthFilled = True
            monthLabel = Format(ws.Cells(lastMonthRow, 3).Value, "MMM-YY") ' Format month label as MMM-YY
            Exit For
        End If
    Next i

    ' If a month is found, set the range for the latest month's data
    If isMonthFilled Then
        Set rngLatestMonth = ws.Range(ws.Cells(lastMonthRow, 4), ws.Cells(lastMonthRow, 12))
    Else
        MsgBox "No recent monthly data found to plot.", vbExclamation
        Exit Sub
    End If

    ' Create a new chart, positioning it based on the sheet type (B23 or B27)
    Set chartObj = ws.ChartObjects.Add(Left:=chartTopLeft.Left, Width:=ws.Range("B23:M23").Width, Top:=chartTopLeft.Top, Height:=ws.Range("B23:B60").Height)
    With chartObj.Chart
        .ChartType = xlColumnStacked

        ' Loop through each metric, adding a series for each with relevant data points
        For i = 1 To 9
            ' For "special" sheets, add a fourth data point (Additional YTD)
            If isSpecialSheet Then
                With .SeriesCollection.NewSeries
                    .Name = rngLegends.Cells(1, i).Value ' Metric name as series name
                    .Values = Array(rngTarget.Cells(1, i).Value, rngLatestMonth.Cells(1, i).Value, rngYTD.Cells(1, i).Value, rngAdditionalYTD.Cells(1, i).Value)
                    .xValues = Array("Target", monthLabel, "YTD", "Additional YTD") ' Category labels
                    ' Add data labels with smaller font in percentage format
                    .ApplyDataLabels
                    .DataLabels.ShowValue = True
                    .DataLabels.NumberFormat = "0.00%" ' Format to percentage with two decimal places
                    .DataLabels.Font.Size = 7.5
                End With
            Else
                ' For other sheets, add three data points: Target, Monthly, and YTD
                With .SeriesCollection.NewSeries
                    .Name = rngLegends.Cells(1, i).Value ' Metric name as series name
                    .Values = Array(rngTarget.Cells(1, i).Value, rngLatestMonth.Cells(1, i).Value, rngYTD.Cells(1, i).Value)
                    .xValues = Array("Target", monthLabel, "YTD") ' Category labels
                    ' Add data labels with smaller font in percentage format
                    .ApplyDataLabels
                    .DataLabels.ShowValue = True
                    .DataLabels.NumberFormat = "0.00%" ' Format to percentage with two decimal places
                    .DataLabels.Font.Size = 7.5
                End With
            End If
        Next i

        ' Set chart title, legend, and axis titles
        .HasTitle = True
        .ChartTitle.Text = ws.Name & " Performance Summary" ' Use sheet name in title for clarity
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom

        ' Configure axis titles
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Metrics"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Performance (%)"
    End With
End Sub

Attribute VB_Name = "Module1"
Option Explicit

Public Sub RefreshAndExport()
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.StatusBar = "Refreshing data connections..."
    
    ' Refresh all queries/pivots
    ThisWorkbook.RefreshAll
    WaitForRefresh

    ' Recalculate workbook
    Application.CalculateFull

    ' Ensure /exports/ exists
    Dim outDir As String
    outDir = ThisWorkbook.Path & "\exports\"
    On Error Resume Next
    MkDir outDir
    On Error GoTo 0
    
    ' Timestamped filenames
    Dim ts As String
    ts = Format(Now, "yyyymmdd_HHMM")

    ' Change if your main dashboard sheet has a different name
    Dim dashName As String
    dashName = "Dashboard"
    
    ' Export to PDF
    Worksheets(dashName).ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=outDir & "Engineering_KPI_Dashboard_" & ts & ".pdf", _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ' Export PNG snapshot (via temporary chart)
    Dim tmpChart As ChartObject
    Worksheets(dashName).Range("A1:K40").CopyPicture xlScreen, xlPicture
    Set tmpChart = Worksheets(dashName).ChartObjects.Add(Left:=10, Top:=10, Width:=800, Height:=600)
    tmpChart.Activate
    tmpChart.Chart.Paste
    tmpChart.Chart.Export Filename:=outDir & "Engineering_KPI_Dashboard_" & ts & ".png", FilterName:="PNG"
    tmpChart.Delete

    Application.StatusBar = "Done. Exported to /exports/"
    Application.ScreenUpdating = True
End Sub

Private Sub WaitForRefresh()
    Dim startTime As Single: startTime = Timer
    ' Wait up to ~60 seconds for refresh/recalc to complete
    Do While Application.CalculationState <> xlDone
        DoEvents
        If Timer - startTime > 60 Then Exit Do
    Loop
End Sub



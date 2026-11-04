Attribute VB_Name = "DashboardAppStyle"
Option Explicit

Private Const DASHBOARD_SHEET As String = "Dashboard"
Private Const FONT_NAME As String = "Aptos"

' Local numeric values for Office drawing enums. This avoids compile failures
' if the Office object-library constants are not exposed in the VBA project.
Private Const SHAPE_ROUND_RECT As Long = 5
Private Const OFFICE_FALSE As Long = 0
Private Const OFFICE_TRUE As Long = -1
Private Const OFFICE_ANCHOR_MIDDLE As Long = 3
Private Const OFFICE_ALIGN_CENTER As Long = 2

' Run this manually from the Macro dialog.
' Formats the existing Dashboard and adds worksheet-navigation buttons.
' Does not rewrite worksheet values, formulas, PivotTables, or chart data sources.
Public Sub ApplyDashboardAppStyle()
    Dim ws As Worksheet
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim navigationRow As Long
    Dim stage As String
    Dim failureNumber As Long
    Dim failureDescription As String

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents

    On Error GoTo HandleError
    stage = "locating the Dashboard worksheet"
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)

    stage = "switching to presentation mode"
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.Zoom = 85

    stage = "formatting the dashboard canvas and title"
    With ws
        .Cells.Font.Name = FONT_NAME
        .Range("A1:N46").Interior.Color = RGB(248, 250, 252)

        With .Range("B2:N3")
            .Interior.Color = RGB(15, 23, 42)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Font.Name = FONT_NAME
            .Font.Size = 20
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With

        With .Range("B4:N4")
            .Interior.Color = RGB(248, 250, 252)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Name = FONT_NAME
            .Font.Size = 10
            .Font.Color = RGB(100, 116, 139)
        End With

        .Rows(2).RowHeight = 28
        .Rows(3).RowHeight = 28
        .Rows(4).RowHeight = 20
    End With

    stage = "formatting the KPI cards"
    FormatKpiCard ws, "B6:D8", "B6:D6", "B7:D8", 20
    FormatKpiCard ws, "F6:H8", "F6:H6", "F7:H8", 15
    FormatKpiCard ws, "J6:L8", "J6:L6", "J7:L8", 20

    stage = "styling dashboard charts"
    StyleDashboardCharts ws

    stage = "adding worksheet-navigation buttons"
    navigationRow = GetNavigationStartRow(ws)

    AddNavigationButton ws, "navSourceData", "SOURCE DATA", "GoToAccountSales", _
        ws.Range("B" & navigationRow & ":D" & (navigationRow + 1))
    AddNavigationButton ws, "navAnnualSummary", "YEARLY SUMMARY", "GoToAnnualSummary", _
        ws.Range("E" & navigationRow & ":G" & (navigationRow + 1))
    AddNavigationButton ws, "navTopAccounts", "TOP 10 ACCOUNTS", "GoToTopAccounts", _
        ws.Range("H" & navigationRow & ":J" & (navigationRow + 1))
    AddNavigationButton ws, "navAccountTypes", "ACCOUNT TYPES", "GoToAccountTypes", _
        ws.Range("K" & navigationRow & ":M" & (navigationRow + 1))

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

HandleError:
    failureNumber = Err.Number
    failureDescription = Err.Description
    MsgBox "Dashboard styling stopped while " & stage & "." & vbCrLf & _
        "Error " & CStr(failureNumber) & ": " & failureDescription, _
        vbExclamation, "Dashboard App Style"
    Resume CleanExit
End Sub

Private Sub FormatKpiCard( _
    ByVal ws As Worksheet, _
    ByVal cardAddress As String, _
    ByVal labelAddress As String, _
    ByVal valueAddress As String, _
    ByVal valueFontSize As Double)

    Dim cardRange As Range
    Set cardRange = ws.Range(cardAddress)

    With cardRange
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Color = RGB(226, 232, 240)
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Color = RGB(226, 232, 240)
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Color = RGB(226, 232, 240)
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(226, 232, 240)
            .Weight = xlThin
        End With
    End With

    With ws.Range(labelAddress)
        .Font.Name = FONT_NAME
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(100, 116, 139)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    With ws.Range(valueAddress)
        .Font.Name = FONT_NAME
        .Font.Size = valueFontSize
        .Font.Bold = True
        .Font.Color = RGB(15, 23, 42)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub

Private Sub StyleDashboardCharts(ByVal ws As Worksheet)
    Dim chartObject As ChartObject
    Dim series As Series
    Dim pointIndex As Long
    Dim chartTitleText As String

    For Each chartObject In ws.ChartObjects
        On Error Resume Next

        With chartObject.Chart
            .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .ChartArea.Format.Line.ForeColor.RGB = RGB(226, 232, 240)
            .ChartArea.Format.Line.Weight = 0.75
            .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)

            chartTitleText = vbNullString
            If .HasTitle Then
                chartTitleText = .ChartTitle.Text
                With .ChartTitle.Format.TextFrame2.TextRange.Font
                    .Name = FONT_NAME
                    .Size = 12
                    .Bold = OFFICE_TRUE
                    .Fill.ForeColor.RGB = RGB(15, 23, 42)
                End With
            End If

            .Axes(xlCategory).TickLabels.Font.Name = FONT_NAME
            .Axes(xlCategory).TickLabels.Font.Size = 9
            .Axes(xlValue).TickLabels.Font.Name = FONT_NAME
            .Axes(xlValue).TickLabels.Font.Size = 9
        End With

        If chartObject.Chart.SeriesCollection.Count > 0 Then
            If InStr(1, chartTitleText, "Unit Sales by Year", vbTextCompare) > 0 Then
                With chartObject.Chart.SeriesCollection(1)
                    .Format.Line.ForeColor.RGB = RGB(15, 118, 110)
                    .Format.Line.Weight = 2.5
                    .MarkerStyle = xlMarkerStyleCircle
                    .MarkerSize = 6
                End With

            ElseIf InStr(1, chartTitleText, "Account Type", vbTextCompare) > 0 Then
                Set series = chartObject.Chart.SeriesCollection(1)
                For pointIndex = 1 To series.Points.Count
                    series.Points(pointIndex).Format.Fill.ForeColor.RGB = _
                        DashboardPaletteColor(pointIndex)
                Next pointIndex

            Else
                For Each series In chartObject.Chart.SeriesCollection
                    series.Format.Fill.ForeColor.RGB = RGB(15, 118, 110)
                    series.Format.Line.ForeColor.RGB = RGB(15, 118, 110)
                Next series
            End If
        End If

        On Error GoTo 0
    Next chartObject
End Sub

Private Function DashboardPaletteColor(ByVal colorIndex As Long) As Long
    Select Case ((colorIndex - 1) Mod 5) + 1
        Case 1: DashboardPaletteColor = RGB(15, 118, 110)
        Case 2: DashboardPaletteColor = RGB(8, 145, 178)
        Case 3: DashboardPaletteColor = RGB(51, 65, 85)
        Case 4: DashboardPaletteColor = RGB(100, 116, 139)
        Case Else: DashboardPaletteColor = RGB(148, 163, 184)
    End Select
End Function

Private Function GetNavigationStartRow(ByVal ws As Worksheet) As Long
    Dim chartObject As ChartObject
    Dim candidateRow As Long
    Dim lastChartRow As Long

    lastChartRow = 42
    For Each chartObject In ws.ChartObjects
        candidateRow = chartObject.BottomRightCell.Row + 2
        If candidateRow > lastChartRow Then lastChartRow = candidateRow
    Next chartObject

    If lastChartRow < 44 Then lastChartRow = 44
    GetNavigationStartRow = lastChartRow
End Function

Private Sub AddNavigationButton( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal buttonText As String, _
    ByVal macroName As String, _
    ByVal anchorRange As Range)

    Dim buttonShape As Shape

    Set buttonShape = Nothing
    On Error Resume Next
    Set buttonShape = ws.Shapes(shapeName)
    On Error GoTo 0

    If buttonShape Is Nothing Then
        Set buttonShape = ws.Shapes.AddShape( _
            SHAPE_ROUND_RECT, _
            anchorRange.Left, _
            anchorRange.Top, _
            anchorRange.Width, _
            anchorRange.Height)
    End If

    With buttonShape
        .Name = shapeName
        .Left = anchorRange.Left
        .Top = anchorRange.Top
        .Width = anchorRange.Width
        .Height = anchorRange.Height
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(15, 118, 110)
        .Line.Visible = OFFICE_FALSE
        .OnAction = "'" & ThisWorkbook.Name & "'!" & macroName

        With .TextFrame2
            .VerticalAnchor = OFFICE_ANCHOR_MIDDLE
            .MarginLeft = 6
            .MarginRight = 6
            .MarginTop = 2
            .MarginBottom = 2
            .TextRange.Text = buttonText
            .TextRange.ParagraphFormat.Alignment = OFFICE_ALIGN_CENTER
            With .TextRange.Font
                .Name = FONT_NAME
                .Size = 9
                .Bold = OFFICE_TRUE
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
            End With
        End With
    End With
End Sub

Private Sub ActivateDashboardWorksheet(ByVal targetSheet As String)
    On Error GoTo HandleError
    ThisWorkbook.Worksheets(targetSheet).Activate
    ThisWorkbook.Worksheets(targetSheet).Range("A1").Select
    Exit Sub

HandleError:
    MsgBox "The worksheet '" & targetSheet & "' could not be opened. Check its tab name.", _
        vbExclamation, "Dashboard Navigation"
End Sub

Public Sub GoToAccountSales()
    ActivateDashboardWorksheet "AccountSales"
End Sub

Public Sub GoToAnnualSummary()
    ActivateDashboardWorksheet "Annual_Sales_Summary"
End Sub

Public Sub GoToTopAccounts()
    ActivateDashboardWorksheet "PT_Top10_2021"
End Sub

Public Sub GoToAccountTypes()
    ActivateDashboardWorksheet "PT_Sales_by_Type"
End Sub

Public Sub ReturnToDashboard()
    ActivateDashboardWorksheet DASHBOARD_SHEET
End Sub

' Restore worksheet headings, gridlines, and a normal working zoom after presentation mode.
Public Sub RestoreDashboardEditView()
    On Error GoTo HandleError

    ThisWorkbook.Worksheets(DASHBOARD_SHEET).Activate
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.Zoom = 100
    ThisWorkbook.Worksheets(DASHBOARD_SHEET).Range("A1").Select
    Exit Sub

HandleError:
    MsgBox "The Dashboard edit view could not be restored: " & Err.Description, _
        vbExclamation, "Dashboard Edit View"
End Sub

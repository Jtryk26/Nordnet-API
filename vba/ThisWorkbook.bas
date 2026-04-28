Attribute VB_Name = "ThisWorkbook"
Option Explicit

Private Sub Workbook_Open()
    On Error GoTo SafeExit

    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.StatusBar = "Opdaterer Nordnet-data..."

    ' Opdater alle queries/dataforbindelser ved åbning.
    ThisWorkbook.RefreshAll

SafeExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

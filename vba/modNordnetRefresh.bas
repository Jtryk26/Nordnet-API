Attribute VB_Name = "modNordnetRefresh"
Option Explicit

Public Sub RefreshNordnetData()
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.StatusBar = "Manuel opdatering af Nordnet-data kører..."

    ThisWorkbook.RefreshAll

    MsgBox "Nordnet-data er opdateret.", vbInformation, "Refresh færdig"

CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Fejl under opdatering: " & Err.Description, vbExclamation, "Refresh fejl"
    Resume CleanExit
End Sub

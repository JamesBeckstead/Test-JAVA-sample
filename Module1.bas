Attribute VB_Name = "Module1"
Option Explicit

Sub AddTabData()
Dim TabLoop As Integer
Dim tgSht As Worksheet
Dim ShtName As String
Dim MainRng As Range
Dim CntRng As Range
Dim Cnt As Long
Dim Lrow As Long
Dim Lclm As Long
Dim TgFld As Long

    'this is a change
    MsgBox "Are you ready?! Don't blink!", Title:="HIRE MARCO PEREZ!"
    Application.ScreenUpdating = False
    'add necessary tabs
    For TabLoop = 1 To 30
        ShtName = "PALLET #" & TabLoop
        On Error Resume Next
        Set tgSht = Worksheets(ShtName)
        If Err.Number <> 0 Then
            Set tgSht = Worksheets.Add(after:=Sheets(Sheets.Count))
            tgSht.Name = ShtName
        End If
    Next TabLoop
    
    'capture main range
    MainSht.Activate
    Lrow = MainSht.Cells(Rows.Count, 1).End(xlUp).Row
    Lclm = MainSht.Cells(Lrow, 1).End(xlUp).End(xlToRight).Column
    Set MainRng = MainSht.Range(Cells(2, 1), Cells(Lrow, Lclm))
    Set CntRng = MainSht.Cells(2, 1).Offset(1, 0).Resize(Lrow - 1, Lclm)
    'filter each tab
    For TabLoop = 1 To 30
        ShtName = "PALLET #" & TabLoop
        TgFld = 6 + TabLoop
        MainRng.AutoFilter field:=6 + TabLoop, Criteria1:="<>"
        On Error Resume Next
        Cnt = CntRng.SpecialCells(xlCellTypeVisible).Rows.Count
        If Err.Number <> 0 Then GoTo NextTabLoop
        MainRng.Columns(5).SpecialCells(xlCellTypeVisible).Copy Destination:=Sheets(ShtName).Cells(13, 1)
        MainRng.Columns(TgFld).SpecialCells(xlCellTypeVisible).Copy Destination:=Sheets(ShtName).Cells(13, 2)
        
        
NextTabLoop:
        MainSht.ShowAllData
    Next TabLoop
    MainSht.Activate
    Application.ScreenUpdating = True
    MsgBox "All done! Just hire Marco!", Title:="YOU KNOW YOU WANT TO HIRE MARCO!"
End Sub

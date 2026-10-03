Option Explicit
'==============================================================
'  ThisWorkbook : ブックを開いたとき・シートを切り替えたときの処理
'==============================================================

Private Sub Workbook_Open()
    On Error Resume Next
    ConfigRefresh
    NormalizeNotes
    RefreshLaneCounts
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    If Sh.Name = shBoard.Name Then
        ConfigRefresh
        RefreshLaneCounts
    End If
End Sub

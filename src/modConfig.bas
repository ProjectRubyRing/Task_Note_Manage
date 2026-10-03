Option Explicit
'==============================================================
'  modConfig : 「設定」シートの読み取りと色・文字の補助関数
'==============================================================

Public Const APP_TITLE As String = "付箋タスクボード"
Public Const NOTE_MARK As String = "STKNOTE"
Public Const DEFAULT_FONT As String = "Meiryo UI"
Public Const MAX_STATUS As Long = 10
Public Const CONFIRM_NO As String = "いいえ"

' ステータス一覧のキャッシュ（各操作の最初に ConfigRefresh で読み直す）
Private mLoaded As Boolean
Private mCount As Long
Private mNames() As String
Private mColors() As Long

'--- 設定セル ------------------------------------------------------

Private Function CfgCell(ByVal nm As String) As Range
    On Error Resume Next
    Set CfgCell = ThisWorkbook.Names(nm).RefersToRange
End Function

Public Function CellText(ByVal c As Range) As String
    Dim v As Variant
    v = c.Value
    If IsError(v) Then Exit Function
    CellText = Trim$(CStr(v))
End Function

'--- ステータス ----------------------------------------------------

Public Sub ConfigRefresh()
    mLoaded = False
End Sub

Private Sub EnsureLoaded()
    Dim rTop As Range, n As Long, i As Long
    If mLoaded Then Exit Sub
    Set rTop = CfgCell("cfgStatusTop")
    If rTop Is Nothing Then
        Err.Raise vbObjectError + 513, APP_TITLE, _
            "「設定」シートのステータス表が見つかりません（名前の定義 cfgStatusTop）。"
    End If
    n = 0
    Do While n < MAX_STATUS
        If Len(CellText(rTop.Offset(n, 0))) = 0 Then Exit Do
        n = n + 1
    Loop
    mCount = n
    If n > 0 Then
        ReDim mNames(1 To n)
        ReDim mColors(1 To n)
        For i = 1 To n
            mNames(i) = CellText(rTop.Offset(i - 1, 0))
            mColors(i) = CLng(rTop.Offset(i - 1, 1).Interior.Color)
        Next
    End If
    mLoaded = True
End Sub

Public Function StatusCount() As Long
    EnsureLoaded
    StatusCount = mCount
End Function

Public Function StatusName(ByVal idx As Long) As String
    EnsureLoaded
    If idx >= 1 And idx <= mCount Then StatusName = mNames(idx)
End Function

Public Function StatusColor(ByVal idx As Long) As Long
    EnsureLoaded
    If idx >= 1 And idx <= mCount Then
        StatusColor = mColors(idx)
    Else
        StatusColor = RGB(255, 255, 255)
    End If
End Function

' 保存されているステータス名から番号を求める（見つからなければ 0）
Public Function StatusIndexOf(ByVal nm As String) As Long
    Dim i As Long
    EnsureLoaded
    For i = 1 To mCount
        If MetaSafe(mNames(i)) = nm Then
            StatusIndexOf = i
            Exit Function
        End If
    Next
End Function

Public Function CheckStatuses(ByRef msg As String) As Boolean
    Dim n As Long, i As Long, j As Long
    n = StatusCount()
    If n < 2 Then
        msg = "「設定」シートのステータスは2つ以上入力してください。"
        Exit Function
    End If
    For i = 1 To n - 1
        For j = i + 1 To n
            If StatusName(i) = StatusName(j) Then
                msg = "「設定」シートのステータス名「" & StatusName(i) & "」が重複しています。"
                Exit Function
            End If
        Next
    Next
    CheckStatuses = True
End Function

' 図形の代替テキストに保存できる形にする
Public Function MetaSafe(ByVal s As String) As String
    s = Replace(s, ";", "；")
    s = Replace(s, "=", "＝")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    MetaSafe = s
End Function

'--- その他の設定 --------------------------------------------------

Public Function CfgFont() As String
    Dim c As Range, s As String
    Set c = CfgCell("cfgFont")
    If Not c Is Nothing Then s = CellText(c)
    If Len(s) = 0 Then s = DEFAULT_FONT
    CfgFont = s
End Function

Private Function CfgNum(ByVal nm As String, ByVal def As Double, ByVal lo As Double, ByVal hi As Double) As Double
    Dim c As Range, v As Variant
    CfgNum = def
    Set c = CfgCell(nm)
    If c Is Nothing Then Exit Function
    v = c.Value
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    v = CDbl(v)
    If v < lo Then v = lo
    If v > hi Then v = hi
    CfgNum = v
End Function

Public Function NoteWidth() As Single
    NoteWidth = CfgNum("cfgWidth", 180, 120, 360)
End Function

Public Function SizeTitle() As Single
    SizeTitle = CfgNum("cfgSizeTitle", 10.5, 6, 24)
End Function

Public Function SizeBody() As Single
    SizeBody = CfgNum("cfgSizeBody", 9, 6, 20)
End Function

Public Function SizeMemo() As Single
    SizeMemo = CfgNum("cfgSizeMemo", 8, 6, 20)
End Function

Public Function ConfirmOnDone() As Boolean
    Dim c As Range
    ConfirmOnDone = True
    Set c = CfgCell("cfgConfirm")
    If Not c Is Nothing Then ConfirmOnDone = (CellText(c) <> CONFIRM_NO)
End Function

'--- 色 ------------------------------------------------------------

' 暗くする（f = 0～1）
Public Function ShadeColor(ByVal c As Long, ByVal f As Double) As Long
    Dim r As Long, g As Long, b As Long
    SplitRGB c, r, g, b
    ShadeColor = RGB(Clamp255(r * (1 - f)), Clamp255(g * (1 - f)), Clamp255(b * (1 - f)))
End Function

' 白に近づける（f = 0～1）
Public Function TintColor(ByVal c As Long, ByVal f As Double) As Long
    Dim r As Long, g As Long, b As Long
    SplitRGB c, r, g, b
    TintColor = RGB(Clamp255(r + (255 - r) * f), Clamp255(g + (255 - g) * f), Clamp255(b + (255 - b) * f))
End Function

' 背景色の上で読みやすい文字色（白 または 濃いグレー）
Public Function TextColorOn(ByVal bg As Long) As Long
    Dim r As Long, g As Long, b As Long
    SplitRGB bg, r, g, b
    If 0.299 * r + 0.587 * g + 0.114 * b < 150 Then
        TextColorOn = RGB(255, 255, 255)
    Else
        TextColorOn = RGB(51, 51, 51)
    End If
End Function

Private Sub SplitRGB(ByVal c As Long, ByRef r As Long, ByRef g As Long, ByRef b As Long)
    r = c And &HFF&
    g = (c \ &H100&) And &HFF&
    b = (c \ &H10000) And &HFF&
End Sub

Private Function Clamp255(ByVal v As Double) As Long
    If v < 0 Then v = 0
    If v > 255 Then v = 255
    Clamp255 = CLng(v)
End Function

'--- 文字 ----------------------------------------------------------

' おおよその文字幅（pt）。全角 = 文字サイズ、半角 = その約半分
Public Function EstTextWidth(ByVal s As String, ByVal sz As Single) As Single
    Dim i As Long, code As Long, w As Single
    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1)) And &HFFFF&
        If code < &H100& Or (code >= &HFF61& And code <= &HFF9F&) Then
            w = w + sz * 0.56
        Else
            w = w + sz
        End If
    Next
    EstTextWidth = w
End Function

' 先頭の改行と、末尾の空白・改行を取り除く
Public Function TrimLines(ByVal s As String) As String
    Do While Len(s) > 0
        If Left$(s, 1) <> vbLf Then Exit Do
        s = Mid$(s, 2)
    Loop
    Do While Len(s) > 0
        Select Case Right$(s, 1)
            Case " ", vbLf, vbTab, "　"
                s = Left$(s, Len(s) - 1)
            Case Else
                Exit Do
        End Select
    Loop
    TrimLines = s
End Function

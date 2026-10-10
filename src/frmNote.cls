Option Explicit
'==============================================================
'  frmNote : 付箋の入力画面（タイトル・内容・備考・開始日・終了日・進捗率・ステータス）
'            画面の部品は UserForm_Initialize で作る
'==============================================================

Public Confirmed As Boolean

Private WithEvents txtTitle As MSForms.TextBox
Private WithEvents txtBody As MSForms.TextBox
Private WithEvents txtMemo As MSForms.TextBox
Private WithEvents txtStart As MSForms.TextBox
Private WithEvents txtEnd As MSForms.TextBox
Private WithEvents btnStartToday As MSForms.CommandButton
Private WithEvents btnEndToday As MSForms.CommandButton
Private WithEvents cboProgress As MSForms.ComboBox
Private WithEvents cboStatus As MSForms.ComboBox
Private WithEvents btnOK As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton
Private lblInfo As MSForms.Label
Private lblSwatch As MSForms.Label
Private mColors As Variant
Private mStart As Date
Private mEnd As Date
Private mProgress As Long

Private Const PAD As Single = 12
Private Const INNER_W As Single = 380
Private Const LABEL_W As Single = 66        ' 開始日・進捗率・ステータスの見出しの幅
Private Const DAY_FORMAT As String = "yyyy/mm/dd"

Private Sub UserForm_Initialize()
    Dim y As Single, innerH As Single, dw As Single, dh As Single, lbl As MSForms.Label, i As Long

    Me.BackColor = RGB(250, 250, 250)

    Set lblInfo = NewLabel("", PAD, 8, INNER_W, 14)
    lblInfo.Font.Size = 8.5
    lblInfo.ForeColor = RGB(128, 128, 128)

    y = 26
    NewLabel "タイトル", PAD, y, INNER_W, 14
    Set txtTitle = Me.Controls.Add("Forms.TextBox.1", "c_title")
    Place txtTitle, PAD, y + 15, INNER_W, 24
    txtTitle.Font.Size = 11
    txtTitle.Font.Bold = True
    txtTitle.MaxLength = 200

    y = y + 48
    NewLabel "内容", PAD, y, INNER_W, 14
    Set txtBody = Me.Controls.Add("Forms.TextBox.1", "c_body")
    Place txtBody, PAD, y + 15, INNER_W, 120
    MultiLineBox txtBody

    y = y + 144
    NewLabel "備考（関係する実装など）", PAD, y, INNER_W, 14
    Set txtMemo = Me.Controls.Add("Forms.TextBox.1", "c_memo")
    Place txtMemo, PAD, y + 15, INNER_W, 72
    MultiLineBox txtMemo

    ' 開始日・終了日（［今日］で今日の日付を入れる）
    y = y + 98
    NewLabel "開始日", PAD, y + 4, LABEL_W, 14
    Set txtStart = Me.Controls.Add("Forms.TextBox.1", "c_start")
    Place txtStart, PAD + LABEL_W, y, 86, 22
    DayBox txtStart
    Set btnStartToday = TodayButton("c_start_today", PAD + LABEL_W + 90, y)
    NewLabel "終了日", PAD + 210, y + 4, 40, 14
    Set txtEnd = Me.Controls.Add("Forms.TextBox.1", "c_end")
    Place txtEnd, PAD + 250, y, 86, 22
    DayBox txtEnd
    Set btnEndToday = TodayButton("c_end_today", PAD + 340, y)
    Set lbl = NewLabel(DayExample() & " のように入力します（空欄でもかまいません）", PAD + LABEL_W, y + 25, 300, 14)
    lbl.Font.Size = 8
    lbl.ForeColor = RGB(140, 140, 140)

    ' 進捗率（一覧から選ぶか、数字を入力する）
    y = y + 44
    NewLabel "進捗率", PAD, y + 4, LABEL_W, 14
    Set cboProgress = Me.Controls.Add("Forms.ComboBox.1", "c_progress")
    Place cboProgress, PAD + LABEL_W, y, 64, 22
    cboProgress.Style = fmStyleDropDownCombo
    cboProgress.MatchEntry = fmMatchEntryNone
    cboProgress.IMEMode = fmIMEModeDisable
    cboProgress.MaxLength = 8
    cboProgress.ListRows = 11
    For i = 0 To 100 Step 10
        cboProgress.AddItem CStr(i)
    Next
    NewLabel "%", PAD + LABEL_W + 68, y + 4, 20, 14

    y = y + 34
    NewLabel "ステータス", PAD, y + 4, LABEL_W, 14
    Set cboStatus = Me.Controls.Add("Forms.ComboBox.1", "c_status")
    Place cboStatus, PAD + LABEL_W, y, 150, 22
    cboStatus.Style = fmStyleDropDownList
    cboStatus.ListRows = MAX_STATUS
    Set lblSwatch = NewLabel("", PAD + LABEL_W + 156, y, 44, 22)
    lblSwatch.BackStyle = fmBackStyleOpaque
    lblSwatch.BorderStyle = fmBorderStyleSingle
    lblSwatch.BorderColor = RGB(166, 166, 166)

    y = y + 36
    Set lbl = NewLabel("Ctrl + Enter で確定　／　Esc でキャンセル", PAD, y + 8, 200, 14)
    lbl.Font.Size = 8
    lbl.ForeColor = RGB(140, 140, 140)
    Set btnOK = Me.Controls.Add("Forms.CommandButton.1", "c_ok")
    Place btnOK, PAD + INNER_W - 180, y, 86, 28
    btnOK.Caption = "OK"
    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "c_cancel")
    Place btnCancel, PAD + INNER_W - 86, y, 86, 28
    btnCancel.Caption = "キャンセル"
    btnCancel.Cancel = True

    ' 中の広さが部品に合うように大きさを決める。
    ' 読み込んだ直後は Width と InsideWidth の関係が正しくないことがあるので、
    ' いったん大きさを設定してから枠（タイトルバーなど）の分を測り直す
    innerH = y + 28 + PAD
    Me.Width = INNER_W + PAD * 2 + 12
    Me.Height = innerH + 30
    dw = Me.Width - Me.InsideWidth
    dh = Me.Height - Me.InsideHeight
    Me.Width = INNER_W + PAD * 2 + dw
    Me.Height = innerH + dh
End Sub

Private Function NewLabel(ByVal caption As String, ByVal x As Single, ByVal y As Single, _
                          ByVal w As Single, ByVal h As Single) As MSForms.Label
    Dim l As MSForms.Label
    Set l = Me.Controls.Add("Forms.Label.1")
    Place l, x, y, w, h
    l.Caption = caption
    l.BackStyle = fmBackStyleTransparent
    Set NewLabel = l
End Function

Private Sub Place(ByVal c As Object, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single)
    c.Left = x
    c.Top = y
    c.Width = w
    c.Height = h
    c.Font.Name = CfgFont()
    c.Font.Size = 10
End Sub

Private Sub MultiLineBox(ByVal t As MSForms.TextBox)
    t.MultiLine = True
    t.WordWrap = True
    t.EnterKeyBehavior = True
    t.ScrollBars = fmScrollBarsVertical
End Sub

Private Sub DayBox(ByVal t As MSForms.TextBox)
    t.IMEMode = fmIMEModeDisable
    t.MaxLength = 20
End Sub

Private Function TodayButton(ByVal nm As String, ByVal x As Single, ByVal y As Single) As MSForms.CommandButton
    Dim b As MSForms.CommandButton
    Set b = Me.Controls.Add("Forms.CommandButton.1", nm)
    Place b, x, y, 36, 22
    b.Caption = "今日"
    b.Font.Size = 9
    b.TabStop = False
    b.TakeFocusOnClick = False
    Set TodayButton = b
End Function

'--- 呼び出し側から使うもの ----------------------------------------

Public Sub Prepare(ByVal formCaption As String, ByVal infoText As String, ByVal t As String, ByVal b As String, _
                   ByVal mm As String, ByVal stNames As Variant, ByVal stColors As Variant, ByVal sel As Long, _
                   ByVal sd As Date, ByVal ed As Date, ByVal pct As Long)
    Dim i As Long
    Me.Caption = formCaption
    lblInfo.Caption = infoText
    txtTitle.Text = t
    txtBody.Text = ToBoxText(b)
    txtMemo.Text = ToBoxText(mm)
    txtStart.Text = DayText(sd)
    txtEnd.Text = DayText(ed)
    cboProgress.Text = CStr(pct)
    mColors = stColors
    cboStatus.Clear
    For i = LBound(stNames) To UBound(stNames)
        cboStatus.AddItem stNames(i)
    Next
    If sel < 1 Or sel > cboStatus.ListCount Then sel = 1
    cboStatus.ListIndex = sel - 1
    UpdateSwatch
End Sub

Public Property Get TitleText() As String
    TitleText = Trim$(Replace(Replace(txtTitle.Text, vbCr, ""), vbLf, ""))
End Property

Public Property Get BodyText() As String
    BodyText = FromBoxText(txtBody.Text)
End Property

Public Property Get MemoText() As String
    MemoText = FromBoxText(txtMemo.Text)
End Property

Public Property Get StatusIndex() As Long
    StatusIndex = cboStatus.ListIndex + 1
End Property

' 開始日・終了日（0 = 空欄）、進捗率。［OK］で確かめた値
Public Property Get StartDay() As Date
    StartDay = mStart
End Property

Public Property Get EndDay() As Date
    EndDay = mEnd
End Property

Public Property Get ProgressPct() As Long
    ProgressPct = mProgress
End Property

'--- 内部 ----------------------------------------------------------

Private Function ToBoxText(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    ToBoxText = Replace(s, vbLf, vbCrLf)
End Function

Private Function FromBoxText(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    FromBoxText = TrimLines(s)
End Function

Private Sub UpdateSwatch()
    Dim i As Long
    i = cboStatus.ListIndex
    If i < 0 Or IsEmpty(mColors) Then Exit Sub
    lblSwatch.BackColor = mColors(LBound(mColors) + i)
End Sub

Private Function DayText(ByVal dt As Date) As String
    If dt <> 0 Then DayText = Format$(dt, DAY_FORMAT)
End Function

Private Function DayExample() As String
    DayExample = Format$(Date, DAY_FORMAT) & " や " & Format$(Date, "m/d")
End Function

Private Function IsDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Or Len(s) > 8 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next
    IsDigits = True
End Function

' 入力された日付を読む（空欄は 0）。読めなければ False
' 2026/10/12・2026-10-12・2026年10月12日・20261012、年を省いた 10/12・1012（今年）を受け付ける
Private Function ParseDay(ByVal s As String, ByRef dt As Date) As Boolean
    Dim p() As String, y As Long, m As Long, d As Long
    dt = 0
    s = Trim$(StrConv(s, vbNarrow))
    s = Replace(Replace(Replace(s, "年", "/"), "月", "/"), "日", "")
    s = Replace(Replace(s, "-", "/"), ".", "/")
    If Len(s) = 0 Then
        ParseDay = True
        Exit Function
    End If
    p = Split(s, "/")
    Select Case UBound(p)
    Case 0
        If Not IsDigits(s) Then Exit Function
        If Len(s) = 8 Then
            y = CLng(Left$(s, 4)): m = CLng(Mid$(s, 5, 2)): d = CLng(Right$(s, 2))
        ElseIf Len(s) = 4 Then
            y = Year(Date): m = CLng(Left$(s, 2)): d = CLng(Right$(s, 2))
        Else
            Exit Function
        End If
    Case 1
        If Not (IsDigits(p(0)) And IsDigits(p(1))) Then Exit Function
        y = Year(Date): m = CLng(p(0)): d = CLng(p(1))
    Case 2
        If Not (IsDigits(p(0)) And IsDigits(p(1)) And IsDigits(p(2))) Then Exit Function
        y = CLng(p(0)): m = CLng(p(1)): d = CLng(p(2))
        If Len(p(0)) <= 2 Then y = y + 2000
    Case Else
        Exit Function
    End Select
    If y < 1900 Or y > 9999 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    dt = DateSerial(y, m, d)
    If Month(dt) <> m Or Day(dt) <> d Then      ' 2/30 などはない日付
        dt = 0
        Exit Function
    End If
    ParseDay = True
End Function

' 入力された進捗率を読む（空欄は 0）。50・50%・５０ などを受け付ける
Private Function ParsePercent(ByVal s As String, ByRef pct As Long) As Boolean
    pct = 0
    s = Trim$(Replace(StrConv(s, vbNarrow), "%", ""))
    If Len(s) = 0 Then
        ParsePercent = True
        Exit Function
    End If
    If Not IsDigits(s) Then Exit Function
    If CLng(s) > 100 Then Exit Function
    pct = CLng(s)
    ParsePercent = True
End Function

' 入力の誤りを知らせて、その欄を選ぶ
Private Sub Reject(ByVal c As Object, ByVal msg As String)
    MsgBox msg, vbExclamation, APP_TITLE
    c.SetFocus
    c.SelStart = 0
    c.SelLength = Len(c.Text)
End Sub

Private Sub Accept()
    Dim sd As Date, ed As Date, pct As Long
    If Len(TitleText) = 0 Then
        MsgBox "タイトルを入力してください。", vbExclamation, APP_TITLE
        txtTitle.SetFocus
        Exit Sub
    End If
    If Not ParseDay(txtStart.Text, sd) Then
        Reject txtStart, "開始日は " & DayExample() & " のように入力してください。"
        Exit Sub
    End If
    If Not ParseDay(txtEnd.Text, ed) Then
        Reject txtEnd, "終了日は " & DayExample() & " のように入力してください。"
        Exit Sub
    End If
    If sd <> 0 And ed <> 0 And ed < sd Then
        Reject txtEnd, "終了日は、開始日と同じ日かそれより後の日付にしてください。"
        Exit Sub
    End If
    If Not ParsePercent(cboProgress.Text, pct) Then
        Reject cboProgress, "進捗率は 0 から 100 までの数字で入力してください。"
        Exit Sub
    End If
    mStart = sd
    mEnd = ed
    mProgress = pct
    Confirmed = True
    Me.Hide
End Sub

Private Sub CheckAcceptKey(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And (Shift And 2) = 2 Then
        KeyCode = 0
        Accept
    End If
End Sub

'--- イベント ------------------------------------------------------

Private Sub UserForm_Activate()
    On Error Resume Next
    txtTitle.SetFocus
    txtTitle.SelStart = 0
    txtTitle.SelLength = Len(txtTitle.Text)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Confirmed = False
        Me.Hide
    End If
End Sub

Private Sub btnOK_Click()
    Accept
End Sub

Private Sub btnCancel_Click()
    Confirmed = False
    Me.Hide
End Sub

Private Sub cboStatus_Change()
    UpdateSwatch
End Sub

Private Sub txtTitle_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CheckAcceptKey KeyCode, Shift
End Sub

Private Sub txtBody_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CheckAcceptKey KeyCode, Shift
End Sub

Private Sub txtMemo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CheckAcceptKey KeyCode, Shift
End Sub

Private Sub txtStart_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CheckAcceptKey KeyCode, Shift
End Sub

Private Sub txtEnd_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CheckAcceptKey KeyCode, Shift
End Sub

Private Sub cboProgress_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CheckAcceptKey KeyCode, Shift
End Sub

Private Sub btnStartToday_Click()
    txtStart.Text = Format$(Date, DAY_FORMAT)
End Sub

Private Sub btnEndToday_Click()
    txtEnd.Text = Format$(Date, DAY_FORMAT)
End Sub

Option Explicit
'==============================================================
'  frmNote : 付箋の入力画面（タイトル・内容・備考・ステータス）
'            画面の部品は UserForm_Initialize で作る
'==============================================================

Public Confirmed As Boolean

Private WithEvents txtTitle As MSForms.TextBox
Private WithEvents txtBody As MSForms.TextBox
Private WithEvents txtMemo As MSForms.TextBox
Private WithEvents cboStatus As MSForms.ComboBox
Private WithEvents btnOK As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton
Private lblInfo As MSForms.Label
Private lblSwatch As MSForms.Label
Private mColors As Variant

Private Const PAD As Single = 12
Private Const INNER_W As Single = 380

Private Sub UserForm_Initialize()
    Dim y As Single, innerH As Single, dw As Single, dh As Single, lbl As MSForms.Label

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

    y = y + 98
    NewLabel "ステータス", PAD, y + 4, 64, 14
    Set cboStatus = Me.Controls.Add("Forms.ComboBox.1", "c_status")
    Place cboStatus, PAD + 66, y, 150, 22
    cboStatus.Style = fmStyleDropDownList
    cboStatus.ListRows = MAX_STATUS
    Set lblSwatch = NewLabel("", PAD + 222, y, 44, 22)
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

'--- 呼び出し側から使うもの ----------------------------------------

Public Sub Prepare(ByVal formCaption As String, ByVal infoText As String, ByVal t As String, ByVal b As String, _
                   ByVal mm As String, ByVal stNames As Variant, ByVal stColors As Variant, ByVal sel As Long)
    Dim i As Long
    Me.Caption = formCaption
    lblInfo.Caption = infoText
    txtTitle.Text = t
    txtBody.Text = ToBoxText(b)
    txtMemo.Text = ToBoxText(mm)
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

Private Sub Accept()
    If Len(TitleText) = 0 Then
        MsgBox "タイトルを入力してください。", vbExclamation, APP_TITLE
        txtTitle.SetFocus
        Exit Sub
    End If
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

Option Explicit
'==============================================================
'  modMain : ボタン・付箋から呼び出されるマクロ
'==============================================================

Private Type PointApi
    X As Long
    Y As Long
End Type

#If VBA7 Then
Private Declare PtrSafe Function GetCursorPos Lib "user32" (ByRef lpPoint As PointApi) As Long
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
#Else
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As PointApi) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
#End If

Private Const VK_LBUTTON As Long = &H1
Private Const VK_RBUTTON As Long = &H2
Private Const VK_SHIFT As Long = &H10
Private Const VK_ESCAPE As Long = &H1B
Private Const SM_SWAPBUTTON As Long = 23
Private Const SM_CXDRAG As Long = 68
Private Const SM_CYDRAG As Long = 69
Private Const LOGPIXELSX As Long = 88

Private mBusy As Boolean

'==============================================================
'  共通処理
'==============================================================

Private Function BeginOp() As Boolean
    Dim msg As String
    If mBusy Then Exit Function
    ConfigRefresh
    If Not CheckStatuses(msg) Then
        MsgBox msg, vbExclamation, APP_TITLE
        Exit Function
    End If
    mBusy = True
    Application.ScreenUpdating = False
    ProtectBoards
    NormalizeNotes
    BeginOp = True
End Function

Private Sub EndOp()
    Application.ScreenUpdating = True
    mBusy = False
End Sub

Private Sub FailOp(ByVal what As String, ByVal num As Long, ByVal desc As String)
    EndOp
    MsgBox what & " の処理中にエラーが発生しました。" & vbCrLf & vbCrLf & "(" & num & ") " & desc, _
           vbExclamation, APP_TITLE
End Sub

Private Sub ShowInfo(ByVal msg As String)
    Application.ScreenUpdating = True
    MsgBox msg, vbInformation, APP_TITLE
End Sub

Private Function OneNote(ByVal g As Shape) As Collection
    Dim c As New Collection
    c.Add g
    Set OneNote = c
End Function

Private Function ShortTitle(ByVal s As String) As String
    s = Replace(s, vbLf, " ")
    If Len(s) > 30 Then s = Left$(s, 30) & "…"
    If Len(s) = 0 Then s = "（タイトルなし）"
    ShortTitle = s
End Function

' ステータスに対応する列（完了はボードに列がないので、その手前の列）
Private Function LaneOf(ByVal si As Long) As Long
    Dim n As Long
    n = StatusCount()
    If si >= n Then si = n - 1
    If si < 1 Then si = 1
    LaneOf = si
End Function

Private Sub Pause(ByVal sec As Single)
    Dim t As Single
    t = Timer
    Do While Timer - t < sec And Timer >= t
        DoEvents
    Loop
End Sub

' マクロから見た画面の DPI（GetCursorPos の座標と同じ基準）
Private Function ScreenDpi() As Long
#If VBA7 Then
    Dim dc As LongPtr
#Else
    Dim dc As Long
#End If
    dc = GetDC(0)
    ScreenDpi = GetDeviceCaps(dc, LOGPIXELSX)
    ReleaseDC 0, dc
    If ScreenDpi <= 0 Then ScreenDpi = 96
End Function

'==============================================================
'  タスクボード
'==============================================================

' ［＋ 新しい付箋］
Public Sub Board_NewNote()
    Dim d As NoteData, g As Shape
    If mBusy Then Exit Sub
    On Error GoTo EH
    If Not BeginOp() Then Exit Sub
    shBoard.Activate
    d.StatusIdx = 1
    If Not EditByForm(d, True) Then
        EndOp
        Exit Sub
    End If
    d.Id = NextNoteId()
    d.Created = Now
    d.Updated = d.Created
    Set g = PlaceInLane(shBoard, d, LaneOf(d.StatusIdx))
    If d.StatusIdx >= StatusCount() Then
        CompleteNotes OneNote(g), False
    Else
        g.Select
        ScrollIntoView g
    End If
    RefreshLaneCounts
    EndOp
    Exit Sub
EH:
    FailOp "新しい付箋", Err.Number, Err.Description
End Sub

' ［編集］
Public Sub Board_EditSelected()
    Dim notes As Collection, g As Shape, d As NoteData, x As Single, y As Single
    If mBusy Then Exit Sub
    On Error GoTo EH
    If Not BeginOp() Then Exit Sub
    Set notes = SelectedNotes(shBoard)
    If notes.Count = 0 Then
        EndOp
        ShowInfo "書き直す付箋をクリックして選んでから、［編集］を押してください。"
        Exit Sub
    End If
    Set g = notes(1)
    d = ReadNote(g)
    If Not EditByForm(d, False) Then
        EndOp
        Exit Sub
    End If
    d.Updated = Now
    x = g.Left
    y = g.Top
    g.Delete
    Set g = BuildNote(shBoard, d, x, y, False)
    If d.StatusIdx >= StatusCount() Then
        CompleteNotes OneNote(g), False
    Else
        g.Select
    End If
    RefreshLaneCounts
    EndOp
    Exit Sub
EH:
    FailOp "編集", Err.Number, Err.Description
End Sub

' ［戻す］
Public Sub Board_StepBack()
    StepSelected -1
End Sub

' ［進める］
Public Sub Board_StepForward()
    StepSelected 1
End Sub

Private Sub StepSelected(ByVal delta As Long)
    Dim notes As Collection
    If mBusy Then Exit Sub
    On Error GoTo EH
    If Not BeginOp() Then Exit Sub
    Set notes = SelectedNotes(shBoard)
    If notes.Count = 0 Then
        EndOp
        ShowInfo "ステータスを変える付箋をクリックして選んでから押してください。" & vbCrLf & _
                 "（Shift キーを押しながらクリックすると、複数の付箋を選べます）"
        Exit Sub
    End If
    ChangeStatus notes, delta, 0
    EndOp
    Exit Sub
EH:
    FailOp "ステータスの変更", Err.Number, Err.Description
End Sub

'==============================================================
'  付箋のクリック・ドラッグ
'==============================================================

' 付箋をクリックしたとき（付箋のどの部品をクリックしてもこのマクロが動く。
' シートを保護しているので、部品を直接動かしたり書き換えたりはできない）
'   ・ドラッグ … 付箋を動かす（選んでいる付箋はまとめて動かす）
'   ・右上のステータスボタン … 次のステータスへ進める
'   ・それ以外 … 付箋を選ぶ（Shift キーを押しながらだと、選択に加える・外す）
' このマクロはマウスのボタンを押したときに動くので、離すまでの動きでクリックかドラッグかを決める
Public Sub Note_Click()
    Dim part As Shape, g As Shape, ws As Worksheet, sel As Collection, moving As Collection
    Dim onTag As Boolean, withShift As Boolean, wasSelected As Boolean
    If mBusy Then Exit Sub
    On Error GoTo EH
    Set part = ClickedPart()   ' 名前の重なりに左右されないよう、先にクリックした図形を特定しておく
    If part Is Nothing Then Exit Sub
    Set g = NoteFromShape(part)
    If g Is Nothing Then Exit Sub
    ' g.Parent は使わない（ParentGroup で取ったグループの Parent は、シートではなく部品の図形になることがある）
    Set ws = g.TopLeftCell.Worksheet
    mBusy = True
    ProtectNotes ws
    withShift = (GetKeyState(VK_SHIFT) < 0)
    onTag = (RoleOf(part) = ROLE_TAG And ws.Name = shBoard.Name)
    Set sel = SelectedNotes(ws)
    wasSelected = HasNote(sel, g)

    ' ステータスボタンはクリックしても選択を変えない。それ以外は押したときに選ぶ
    If Not wasSelected And Not onTag Then
        g.Select Replace:=Not withShift
        Set sel = SelectedNotes(ws)
    End If
    If HasNote(sel, g) Then Set moving = sel Else Set moving = OneNote(g)

    If DragNotes(ws, moving) Then
        If Not HasNote(SelectedNotes(ws), g) Then g.Select Replace:=Not withShift
    ElseIf onTag Then
        mBusy = False
        If BeginOp() Then
            ChangeStatus OneNote(g), 1, 0
            EndOp
        End If
        Exit Sub
    ElseIf wasSelected Then
        If withShift Then UnselectNote ws, g Else g.Select
    End If
    mBusy = False
    Exit Sub
EH:
    FailOp "付箋の操作", Err.Number, Err.Description
End Sub

' クリックされた図形（付箋の部品）
Private Function ClickedPart() As Shape
    Dim caller As String, pt As PointApi, hit As Object, s As Shape
    If TypeName(Application.Caller) <> "String" Then Exit Function
    caller = Application.Caller
    ' マウスの位置にある図形から求める（同じ名前の付箋があっても正しく選べる）
    On Error Resume Next
    If GetCursorPos(pt) <> 0 Then
        Set hit = ActiveWindow.RangeFromPoint(pt.X, pt.Y)
        If Not hit Is Nothing Then
            If TypeName(hit) <> "Range" Then
                If hit.Name = caller Then Set s = hit.ShapeRange.Item(1)
            End If
        End If
    End If
    ' 見つからなければ名前から求める
    If s Is Nothing Then Set s = ActiveSheet.Shapes(caller)
    On Error GoTo 0
    Set ClickedPart = s
End Function

Private Function HasNote(ByVal notes As Collection, ByVal g As Shape) As Boolean
    Dim x As Shape
    For Each x In notes
        If x.ID = g.ID Then
            HasNote = True
            Exit Function
        End If
    Next
End Function

' 選んでいる付箋から g を外す（ほかに残らなければセルを選ぶ）
Private Sub UnselectNote(ByVal ws As Worksheet, ByVal g As Shape)
    Dim x As Shape, first As Boolean
    first = True
    For Each x In SelectedNotes(ws)
        If x.ID <> g.ID Then
            x.Select Replace:=first
            first = False
        End If
    Next
    If first Then g.TopLeftCell.Select
End Sub

' マウスのボタンを離すまで付箋を一緒に動かす。動かしたら True
' （見出し・ボタンの段より上へは出さない。Esc キーを押すと元の位置に戻してやめる）
Private Function DragNotes(ByVal ws As Worksheet, ByVal notes As Collection) As Boolean
    Dim btn As Long, p0 As PointApi, p As PointApi, kx As Double, ky As Double
    Dim x0() As Single, y0() As Single, w0() As Single, h0() As Single, minX As Single, minY As Single, topY As Single
    Dim dx As Single, dy As Single, i As Long, g As Shape, dragging As Boolean, canceled As Boolean
    btn = VK_LBUTTON
    If GetSystemMetrics(SM_SWAPBUTTON) <> 0 Then btn = VK_RBUTTON
    If GetCursorPos(p0) = 0 Then Exit Function
    ' 画面の 1 ピクセルが何 pt か（シートの拡大率・画面の拡大率を含む）。
    ' マクロの中の GetCursorPos は、拡大率の違う画面でもシステムの DPI で数えた座標を返すので、その DPI で換算する
    ' （PointsToScreenPixels はそのような画面で正しく変換されないことがあるので使わない）
    kx = 72# / (ScreenDpi() * ActiveWindow.Zoom / 100#)
    ky = kx
    ReDim x0(1 To notes.Count)
    ReDim y0(1 To notes.Count)
    ReDim w0(1 To notes.Count)
    ReDim h0(1 To notes.Count)
    For Each g In notes
        i = i + 1
        x0(i) = g.Left
        y0(i) = g.Top
        w0(i) = g.Width
        h0(i) = g.Height
        If i = 1 Or x0(i) < minX Then minX = x0(i)
        If i = 1 Or y0(i) < minY Then minY = y0(i)
    Next
    topY = ws.Rows(FIRST_ROW).Top

    Do While (GetAsyncKeyState(btn) And &H8000) <> 0
        GetCursorPos p
        If Not dragging Then
            If Abs(p.X - p0.X) > GetSystemMetrics(SM_CXDRAG) Or Abs(p.Y - p0.Y) > GetSystemMetrics(SM_CYDRAG) Then
                dragging = True
                For Each g In notes
                    g.ZOrder msoBringToFront
                Next
            End If
        End If
        If dragging Then
            canceled = ((GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0)
            If canceled Then
                dx = 0
                dy = 0
            Else
                dx = (p.X - p0.X) * kx
                dy = (p.Y - p0.Y) * ky
                If minX + dx < 0 Then dx = -minX
                If minY + dy < topY Then dy = topY - minY
            End If
            i = 0
            For Each g In notes
                i = i + 1
                g.Left = x0(i) + dx
                g.Top = y0(i) + dy
            Next
            If canceled Then Exit Do
        End If
        DoEvents
        Sleep 10
    Loop
    ' 拡大率の違う画面では、動かすたびに Excel が付箋の大きさを丸め直して少しずつ変わるので、元の大きさに戻す
    If dragging Then
        i = 0
        For Each g In notes
            i = i + 1
            If g.Width <> w0(i) Then g.Width = w0(i)
            If g.Height <> h0(i) Then g.Height = h0(i)
        Next
    End If
    DragNotes = dragging
End Function

' ステータスの見出し：選んでいる付箋をそのステータスにする
Public Sub Lane_Click()
    Dim idx As Long, notes As Collection, caller As String
    If mBusy Then Exit Sub
    On Error GoTo EH
    If TypeName(Application.Caller) <> "String" Then Exit Sub
    caller = Application.Caller
    idx = CLng(Val(Mid$(caller, Len(LANE_PREFIX) + 1)))
    If Not BeginOp() Then Exit Sub
    If idx < 1 Or idx > StatusCount() Then
        EndOp
        Exit Sub
    End If
    Set notes = SelectedNotes(shBoard)
    If notes.Count = 0 Then
        EndOp
        ShowInfo "付箋をクリックして選んでからこの見出しをクリックすると、" & vbCrLf & _
                 "その付箋を「" & StatusName(idx) & "」にできます。"
        Exit Sub
    End If
    ChangeStatus notes, 0, idx
    EndOp
    Exit Sub
EH:
    FailOp "ステータスの変更", Err.Number, Err.Description
End Sub

' ［整列］
Public Sub Board_Arrange()
    If mBusy Then Exit Sub
    On Error GoTo EH
    If Not BeginOp() Then Exit Sub
    ArrangeBoard True
    RefreshLaneCounts
    EndOp
    Exit Sub
EH:
    FailOp "整列", Err.Number, Err.Description
End Sub

' ［捨てる］
Public Sub Board_Discard()
    DiscardSelected shBoard
End Sub

'==============================================================
'  完了シート
'==============================================================

' ［ボードに戻す］
Public Sub Done_Restore()
    Dim notes As Collection, g As Shape, ng As Shape, d As NoteData, n As Long, back() As String, k As Long
    If mBusy Then Exit Sub
    On Error GoTo EH
    If Not BeginOp() Then Exit Sub
    Set notes = SelectedNotes(shDone)
    If notes.Count = 0 Then
        EndOp
        ShowInfo "ボードに戻す付箋をクリックして選んでから押してください。"
        Exit Sub
    End If
    n = StatusCount()
    ReDim back(1 To notes.Count)
    For Each g In notes
        d = ReadNote(g)
        d.StatusIdx = n - 1
        d.Done = 0
        d.Updated = Now
        g.Delete
        If d.HasPos Then
            ' 完了にする前に貼っていた場所へ。ふさがっていたら列の空いている場所へ
            Set ng = BuildNote(shBoard, d, d.PosLeft, d.PosTop, False)
            If HitsOther(shBoard, ng) Then
                ng.Left = LaneLeft(shBoard, n - 1)
                ng.Top = FreeTop(shBoard, ng.Left, ng.Width, ng.Height, ng)
            End If
        Else
            Set ng = PlaceInLane(shBoard, d, n - 1)
        End If
        k = k + 1
        back(k) = ng.Name
    Next
    ArrangeDoneSheet False
    RefreshLaneCounts
    shBoard.Activate
    shBoard.Shapes.Range(back).Select
    ScrollIntoView shBoard.Shapes(back(1))
    EndOp
    Exit Sub
EH:
    FailOp "ボードに戻す", Err.Number, Err.Description
End Sub

' ［整列］（完了シート）
Public Sub Done_Arrange()
    If mBusy Then Exit Sub
    On Error GoTo EH
    If Not BeginOp() Then Exit Sub
    ArrangeDoneSheet True
    EndOp
    Exit Sub
EH:
    FailOp "整列", Err.Number, Err.Description
End Sub

' ［捨てる］（完了シート）
Public Sub Done_Discard()
    DiscardSelected shDone
End Sub

Private Sub DiscardSelected(ByVal ws As Worksheet)
    Dim notes As Collection, g As Shape, d As NoteData, msg As String
    If mBusy Then Exit Sub
    On Error GoTo EH
    If Not BeginOp() Then Exit Sub
    Set notes = SelectedNotes(ws)
    If notes.Count = 0 Then
        EndOp
        ShowInfo "捨てる付箋をクリックして選んでから押してください。" & vbCrLf & _
                 "（Shift キーを押しながらクリックすると、複数の付箋を選べます）"
        Exit Sub
    End If
    If notes.Count = 1 Then
        d = ReadNote(notes(1))
        msg = "付箋「" & ShortTitle(d.Title) & "」を捨てます。"
    Else
        msg = notes.Count & " 枚の付箋を捨てます。"
    End If
    msg = msg & vbCrLf & "捨てた付箋は元に戻せません。よろしいですか？"
    Application.ScreenUpdating = True
    If MsgBox(msg, vbExclamation + vbOKCancel + vbDefaultButton2, APP_TITLE) <> vbOK Then
        EndOp
        Exit Sub
    End If
    Application.ScreenUpdating = False
    For Each g In notes
        g.Delete
    Next
    If ws.Name = shDone.Name Then ArrangeDoneSheet False
    RefreshLaneCounts
    EndOp
    Exit Sub
EH:
    FailOp "捨てる", Err.Number, Err.Description
End Sub

'==============================================================
'  設定シート
'==============================================================

' ［設定を反映］
Public Sub Config_Apply()
    If mBusy Then Exit Sub
    On Error GoTo EH
    If Not BeginOp() Then Exit Sub
    BuildToolbars
    RebuildNotes shBoard, False
    RebuildNotes shDone, True
    ArrangeDoneSheet False
    RefreshLaneCounts
    EndOp
    ShowInfo "設定を反映しました。"
    Exit Sub
EH:
    FailOp "設定の反映", Err.Number, Err.Description
End Sub

'==============================================================
'  ステータスの変更と完了
'==============================================================

' delta : +1 / -1 で前後へ。target > 0 ならそのステータスにする
Private Sub ChangeStatus(ByVal notes As Collection, ByVal delta As Long, ByVal target As Long)
    Dim n As Long, g As Shape, d As NoteData, ni As Long, toDone As Collection
    Set toDone = New Collection
    n = StatusCount()
    For Each g In notes
        d = ReadNote(g)
        If target > 0 Then ni = target Else ni = d.StatusIdx + delta
        If ni < 1 Then ni = 1
        If ni > n Then ni = n
        If ni = n Then
            toDone.Add g
        ElseIf ni <> d.StatusIdx Then
            d.StatusIdx = ni
            d.Updated = Now
            RestyleNote g, d, False
        End If
    Next
    If toDone.Count > 0 Then CompleteNotes toDone, True
    RefreshLaneCounts
End Sub

' 完了：その場でグレーアウトしてから「完了」シートへ移す
Private Sub CompleteNotes(ByVal notes As Collection, ByVal ask As Boolean)
    Dim n As Long, g As Shape, d As NoteData, msg As String
    n = StatusCount()
    If ask And ConfirmOnDone() Then
        If notes.Count = 1 Then
            d = ReadNote(notes(1))
            msg = "「" & ShortTitle(d.Title) & "」を「" & StatusName(n) & "」にします。"
        Else
            msg = notes.Count & " 枚の付箋を「" & StatusName(n) & "」にします。"
        End If
        msg = msg & vbCrLf & "付箋はグレーアウトして「" & shDone.Name & "」シートへ移動します。よろしいですか？"
        Application.ScreenUpdating = True
        If MsgBox(msg, vbQuestion + vbOKCancel, APP_TITLE) <> vbOK Then Exit Sub
        Application.ScreenUpdating = False
    End If

    ' 1) グレーアウト
    For Each g In notes
        d = ReadNote(g)
        d.StatusIdx = n
        d.Done = Now
        d.Updated = d.Done
        d.HasPos = True
        d.PosLeft = g.Left
        d.PosTop = g.Top
        RestyleNote g, d, True
    Next

    ' 2) グレーになったところを少し見せる
    Application.ScreenUpdating = True
    DoEvents
    Pause 0.6
    Application.ScreenUpdating = False

    ' 3) 完了シートへ移動
    For Each g In notes
        d = ReadNote(g)
        g.Delete
        BuildNote shDone, d, OriginX(shDone), OriginY(shDone), True
    Next
    ArrangeDoneSheet False
End Sub

'==============================================================
'  入力画面
'==============================================================

Private Function EditByForm(ByRef d As NoteData, ByVal isNew As Boolean) As Boolean
    Dim f As frmNote, n As Long, i As Long, stNames() As String, stColors() As Long
    Dim formCaption As String, infoText As String
    n = StatusCount()
    ReDim stNames(1 To n)
    ReDim stColors(1 To n)
    For i = 1 To n
        stNames(i) = StatusName(i)
        stColors(i) = StatusColor(i)
    Next
    If isNew Then
        formCaption = "新しい付箋"
        infoText = "起票日時は［OK］を押したときの日時が自動で入ります。"
    Else
        formCaption = "付箋の編集"
        infoText = "No." & Format$(d.Id, "0000") & "　　起票日時 " & Format$(d.Created, "yyyy/mm/dd hh:nn") & _
                   "　　更新 " & Format$(d.Updated, "yyyy/mm/dd hh:nn")
    End If
    Set f = New frmNote
    f.Prepare formCaption, infoText, d.Title, d.Body, d.Memo, stNames, stColors, d.StatusIdx, _
              d.StartDate, d.EndDate, d.Progress
    Application.ScreenUpdating = True
    f.Show
    Application.ScreenUpdating = False
    If f.Confirmed Then
        d.Title = f.TitleText
        d.Body = f.BodyText
        d.Memo = f.MemoText
        d.StatusIdx = f.StatusIndex
        d.StartDate = f.StartDay
        d.EndDate = f.EndDay
        d.Progress = f.ProgressPct
        EditByForm = True
    End If
    Unload f
End Function

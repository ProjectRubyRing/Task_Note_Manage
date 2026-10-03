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
#Else
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As PointApi) As Long
#End If

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
                 "（Ctrl キーを押しながらクリックすると、複数の付箋を選べます）"
        Exit Sub
    End If
    ChangeStatus notes, delta, 0
    EndOp
    Exit Sub
EH:
    FailOp "ステータスの変更", Err.Number, Err.Description
End Sub

' 付箋の右上のステータスボタン：次のステータスへ進める
Public Sub Note_TagClick()
    Dim g As Shape
    If mBusy Then Exit Sub
    On Error GoTo EH
    Set g = ClickedNote()   ' 名前の重なりに左右されないよう、先にクリックした付箋を特定しておく
    If Not BeginOp() Then Exit Sub
    If Not g Is Nothing Then
        If g.Parent.Name = shBoard.Name Then ChangeStatus OneNote(g), 1, 0
    End If
    EndOp
    Exit Sub
EH:
    FailOp "ステータスの変更", Err.Number, Err.Description
End Sub

' クリックされたステータスボタンの付箋
Private Function ClickedNote() As Shape
    Dim caller As String, pt As PointApi, hit As Object, g As Shape
    If TypeName(Application.Caller) <> "String" Then Exit Function
    caller = Application.Caller
    ' マウスの位置にある図形から求める（コピーで同じ名前の付箋があっても正しく選べる）
    On Error Resume Next
    If GetCursorPos(pt) <> 0 Then
        Set hit = ActiveWindow.RangeFromPoint(pt.X, pt.Y)
        If Not hit Is Nothing Then
            If TypeName(hit) <> "Range" Then
                If hit.Name = caller Then Set g = NoteFromShape(hit.ShapeRange.Item(1))
            End If
        End If
    End If
    On Error GoTo 0
    ' 見つからなければ名前から求める
    If g Is Nothing Then Set g = NoteByShapeName(ActiveSheet, caller)
    Set ClickedNote = g
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
                 "（Ctrl キーを押しながらクリックすると、複数の付箋を選べます）"
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
    Else
        formCaption = "付箋の編集"
        infoText = "No." & Format$(d.Id, "0000") & "　　起票 " & Format$(d.Created, "yyyy/m/d") & _
                   "　　更新 " & Format$(d.Updated, "yyyy/m/d")
    End If
    Set f = New frmNote
    f.Prepare formCaption, infoText, d.Title, d.Body, d.Memo, stNames, stColors, d.StatusIdx
    Application.ScreenUpdating = True
    f.Show
    Application.ScreenUpdating = False
    If f.Confirmed Then
        d.Title = f.TitleText
        d.Body = f.BodyText
        d.Memo = f.MemoText
        d.StatusIdx = f.StatusIndex
        EditByForm = True
    End If
    Unload f
End Function

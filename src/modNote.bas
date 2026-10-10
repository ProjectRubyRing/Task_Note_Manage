Option Explicit
'==============================================================
'  modNote : 付箋（図形のグループ）の作成・読み取り・色付け・配置
'
'  付箋 1 枚 = 次の図形をグループ化したもの（グループ名 "STK<番号>"）
'     base … 台紙（付箋の色・影）       hdr  … 見出し帯（番号・日付）
'     tag  … ステータスボタン（クリックで次へ）
'     ttl  … タイトル    body … 内容    memo … 備考    mlbl … 「備考」の文字
'     info … 起票日時・開始日・終了日    prog … 進捗率の文字
'     pbar … 進捗率のバーの枠            pfill … バーの塗り（進捗率が 0% のときは作らない）
'  付箋の情報（番号・ステータス・日付・進捗率）はグループの「代替テキスト」に保存する。
'  付箋を貼るシートは保護して部品を直接さわれないようにし、どの部品をクリックしても
'  Note_Click（modMain）が動くようにしている（選ぶ・動かす・ステータスを進める）。
'==============================================================

Public Type NoteData
    Id As Long
    Title As String
    Body As String
    Memo As String
    StatusIdx As Long
    Created As Date         ' 起票日時（新しい付箋を作ったときに自動で入る）
    Updated As Date
    Done As Date
    StartDate As Date       ' 開始日（0 = 未入力）
    EndDate As Date         ' 終了日（0 = 未入力）
    Progress As Long        ' 進捗率（0～100）
    HasPos As Boolean       ' 完了前にボードで貼られていた位置を持っているか
    PosLeft As Single
    PosTop As Single
    Sig As String           ' 作ったときの文字の要約（直接書き直されたかの判定に使う）
End Type

Public Const FIRST_ROW As Long = 4          ' 付箋を貼る領域の先頭行
Public Const GAP_X As Single = 14           ' 付箋どうしの横の間隔（pt）
Public Const GAP_Y As Single = 14           ' 付箋どうしの縦の間隔（pt）
Public Const DONE_COLS As Long = 5          ' 完了シートで横に並べる枚数
Public Const LANE_PREFIX As String = "LANE_"

Public Const ROLE_BASE As String = "base"
Public Const ROLE_HDR As String = "hdr"
Public Const ROLE_TAG As String = "tag"
Public Const ROLE_TITLE As String = "ttl"
Public Const ROLE_BODY As String = "body"
Public Const ROLE_MEMO As String = "memo"
Public Const ROLE_LABEL As String = "mlbl"
Public Const ROLE_INFO As String = "info"
Public Const ROLE_PROG As String = "prog"
Public Const ROLE_PBAR As String = "pbar"
Public Const ROLE_PFILL As String = "pfill"

Private Const HDR_H As Single = 20
Private Const TAG_H As Single = 15
Private Const TAG_SIZE As Single = 7.5
Private Const TAG_MIN_SIZE As Single = 5
Private Const HDR_SIZE As Single = 7.5
Private Const TITLE_MIN_H As Single = 22
Private Const TITLE_MIN_SIZE As Single = 6  ' タイトルを小さくするときの下限
Private Const TITLE_FIT_TOL As Single = 2   ' 欄に入ったとみなす高さの余裕（置いた位置による 1～2 ピクセルの丸めの差。1 行は 10pt 以上）
Private Const TEXT_SLACK As Single = 0.06   ' 折り返す文字の高さを測るときに狭める幅の割合（TextHeight）
Private Const BODY_MIN_H As Single = 60
Private Const MEMO_MIN_H As Single = 34
Private Const MEMO_LABEL_H As Single = 12
Private Const INFO_SIZE As Single = 7.5     ' 起票日時・開始日・終了日・進捗率の文字サイズ
Private Const INFO_TAB As Single = 36       ' 見出し（起票日時 など）のあとの値を揃える位置
Private Const PCT_W As Single = 26          ' 進捗率の「100%」の文字の分
Private Const PBAR_H As Single = 6
Private Const NOTE_MACRO As String = "Note_Click"
Private Const KIND_TEXTBOX As Long = 0

Private Const LBL_CREATED As String = "起票日時"
Private Const LBL_START As String = "開始日"
Private Const LBL_END As String = "終了日"
Private Const LBL_PROGRESS As String = "進捗率"
Private Const NO_VALUE As String = "―"

Private mSeq As Long
Private mAreaKey As String      ' タイトルの欄の高さを測ったときの設定
Private mAreaH As Single

'==============================================================
'  識別
'==============================================================

Public Function NoteName(ByVal id As Long) As String
    NoteName = "STK" & id
End Function

Public Function IsNote(ByVal shp As Shape) As Boolean
    On Error GoTo Fail
    If shp.Type = msoGroup Then
        IsNote = (Left$(shp.AlternativeText, Len(NOTE_MARK) + 1) = NOTE_MARK & ";")
    End If
    Exit Function
Fail:
    IsNote = False
End Function

' 子図形の役割（代替テキスト "role:xxx"、なければ名前の末尾）
Public Function RoleOf(ByVal shp As Shape) As String
    Dim a As String, p As Long
    On Error Resume Next
    a = shp.AlternativeText
    On Error GoTo 0
    If Left$(a, 5) = "role:" Then
        RoleOf = Mid$(a, 6)
    Else
        p = InStrRev(shp.Name, "_")
        If p > 0 Then RoleOf = Mid$(shp.Name, p + 1)
    End If
End Function

Public Function PartOf(ByVal grp As Shape, ByVal role As String) As Shape
    Dim i As Long
    For i = 1 To grp.GroupItems.Count
        If RoleOf(grp.GroupItems(i)) = role Then
            Set PartOf = grp.GroupItems(i)
            Exit Function
        End If
    Next
End Function

Public Function NotesOn(ByVal ws As Worksheet) As Collection
    Dim col As New Collection, shp As Shape
    For Each shp In ws.Shapes
        If IsNote(shp) Then col.Add shp
    Next
    Set NotesOn = col
End Function

' 付箋のグループ、またはその一部の図形から付箋のグループを求める
Public Function NoteFromShape(ByVal shp As Shape) As Shape
    Dim p As Shape
    If IsNote(shp) Then
        Set NoteFromShape = shp
        Exit Function
    End If
    On Error Resume Next
    Set p = shp.ParentGroup
    On Error GoTo 0
    If Not p Is Nothing Then
        If IsNote(p) Then Set NoteFromShape = p
    End If
End Function

' 選択中の付箋（複数可）
Public Function SelectedNotes(ByVal ws As Worksheet) As Collection
    Dim col As New Collection, sr As ShapeRange, i As Long, g As Shape, seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    If ActiveSheet.Name = ws.Name Then
        On Error Resume Next
        Set sr = Selection.ShapeRange
        On Error GoTo 0
        If Not sr Is Nothing Then
            For i = 1 To sr.Count
                Set g = NoteFromShape(sr.Item(i))
                If Not g Is Nothing Then
                    If Not seen.Exists(g.Name) Then
                        seen.Add g.Name, True
                        col.Add g
                    End If
                End If
            Next
        End If
    End If
    Set SelectedNotes = col
End Function

'==============================================================
'  付箋の情報（グループの代替テキスト）
'==============================================================

Private Function MetaGet(ByVal meta As String, ByVal key As String) As String
    Dim parts As Variant, i As Long
    parts = Split(meta, ";")
    For i = LBound(parts) To UBound(parts)
        If Left$(parts(i), Len(key) + 1) = key & "=" Then
            MetaGet = Mid$(parts(i), Len(key) + 2)
            Exit Function
        End If
    Next
End Function

Public Sub WriteMeta(ByVal grp As Shape, ByRef d As NoteData)
    Dim s As String
    s = NOTE_MARK & ";id=" & d.Id & ";st=" & MetaSafe(StatusName(ClampStatus(d.StatusIdx))) & _
        ";si=" & d.StatusIdx & ";cr=" & IsoText(d.Created) & ";up=" & IsoText(d.Updated) & _
        ";dn=" & IsoText(d.Done) & ";sd=" & IsoDay(d.StartDate) & ";ed=" & IsoDay(d.EndDate) & _
        ";pg=" & ClampPct(d.Progress) & ";tx=" & d.Sig
    If d.HasPos Then s = s & ";px=" & NumText(d.PosLeft) & ";py=" & NumText(d.PosTop)
    grp.AlternativeText = s
End Sub

Public Function ReadNote(ByVal grp As Shape) As NoteData
    Dim d As NoteData, m As String, px As String
    m = grp.AlternativeText
    d.Id = CLng(Val(MetaGet(m, "id")))
    d.StatusIdx = StatusFromMeta(m)
    d.Created = ParseIso(MetaGet(m, "cr"))
    d.Updated = ParseIso(MetaGet(m, "up"))
    d.Done = ParseIso(MetaGet(m, "dn"))
    d.StartDate = ParseIso(MetaGet(m, "sd"))
    d.EndDate = ParseIso(MetaGet(m, "ed"))
    d.Progress = ClampPct(Val(MetaGet(m, "pg")))
    px = MetaGet(m, "px")
    If Len(px) > 0 Then
        d.HasPos = True
        d.PosLeft = Val(px)
        d.PosTop = Val(MetaGet(m, "py"))
    End If
    d.Sig = MetaGet(m, "tx")
    d.Title = PartText(grp, ROLE_TITLE)
    d.Body = PartText(grp, ROLE_BODY)
    d.Memo = PartText(grp, ROLE_MEMO)
    ReadNote = d
End Function

' 文字の要約（文字数とハッシュ）
Private Function TextSig(ByVal s As String) As String
    Dim i As Long, h As Double
    For i = 1 To Len(s)
        h = h * 31 + (AscW(Mid$(s, i, 1)) And &HFFFF&)
        h = h - Int(h / 2147483647#) * 2147483647#
    Next
    TextSig = Len(s) & "-" & Format$(h, "0")
End Function

Private Function NoteSig(ByVal grp As Shape) As String
    NoteSig = TextSig(PartText(grp, ROLE_TITLE) & vbLf & PartText(grp, ROLE_BODY) & vbLf & PartText(grp, ROLE_MEMO))
End Function

' 文字を直接書き直した・大きさを変えたなどで、作り直しが必要か
Private Function NeedsRefit(ByVal grp As Shape) As Boolean
    If Abs(grp.Width - NoteWidth()) > 0.5 Then
        NeedsRefit = True
    ElseIf PartOf(grp, ROLE_INFO) Is Nothing Then
        NeedsRefit = True       ' 起票日時などの欄がない（前の版で作った）付箋
    Else
        NeedsRefit = (NoteSig(grp) <> MetaGet(grp.AlternativeText, "tx"))
    End If
End Function

' ステータス番号だけを読む（名前で探し、名前が変わっていたら保存時の番号を使う）
Public Function NoteStatus(ByVal grp As Shape) As Long
    NoteStatus = StatusFromMeta(grp.AlternativeText)
End Function

Private Function StatusFromMeta(ByVal m As String) As Long
    Dim si As Long
    si = StatusIndexOf(MetaGet(m, "st"))
    If si = 0 Then si = CLng(Val(MetaGet(m, "si")))
    StatusFromMeta = ClampStatus(si)
End Function

Public Function NoteId(ByVal grp As Shape) As Long
    NoteId = CLng(Val(MetaGet(grp.AlternativeText, "id")))
End Function

Public Function ClampStatus(ByVal si As Long) As Long
    Dim n As Long
    n = StatusCount()
    If si > n Then si = n
    If si < 1 Then si = 1
    ClampStatus = si
End Function

Public Function ClampPct(ByVal v As Double) As Long
    If v < 0 Then v = 0
    If v > 100 Then v = 100
    ClampPct = CLng(v)
End Function

Private Function IsoText(ByVal dt As Date) As String
    If dt = 0 Then Exit Function
    IsoText = Format$(dt, "yyyy-mm-dd hh:nn:ss")
End Function

Private Function IsoDay(ByVal dt As Date) As String
    If dt = 0 Then Exit Function
    IsoDay = Format$(dt, "yyyy-mm-dd")
End Function

Private Function ParseIso(ByVal s As String) As Date
    Dim dt As Date
    On Error GoTo Fail
    If Len(s) < 10 Then Exit Function
    dt = DateSerial(CInt(Val(Mid$(s, 1, 4))), CInt(Val(Mid$(s, 6, 2))), CInt(Val(Mid$(s, 9, 2))))
    If Len(s) >= 19 Then
        dt = dt + TimeSerial(CInt(Val(Mid$(s, 12, 2))), CInt(Val(Mid$(s, 15, 2))), CInt(Val(Mid$(s, 18, 2))))
    End If
    ParseIso = dt
    Exit Function
Fail:
    ParseIso = 0
End Function

Private Function NumText(ByVal v As Single) As String
    NumText = Trim$(Str$(Round(v, 2)))
End Function

'==============================================================
'  文字
'==============================================================

Private Function ToShapeText(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    ToShapeText = Replace(s, vbCr, vbLf)
End Function

Private Function PartText(ByVal grp As Shape, ByVal role As String) As String
    Dim s As Shape, t As String
    Set s = PartOf(grp, role)
    If s Is Nothing Then Exit Function
    t = s.TextFrame2.TextRange.Text
    t = Replace(t, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    t = Replace(t, Chr$(11), vbLf)
    PartText = TrimLines(t)
End Function

Private Sub SetText(ByVal shp As Shape, ByVal txt As String, ByVal sz As Single, ByVal isBold As Boolean)
    Dim fnt As String
    fnt = CfgFont()
    ' 付箋を動かすと Excel が大きさを画面の点に合わせ直して少し縮むことがある。
    ' 図形の文字は入りきらない行を隠す設定が既定なので、最後の行が消えないよう、はみ出しても表示する
    shp.TextFrame.VerticalOverflow = xlOartVerticalOverflowOverflow
    shp.TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
    With shp.TextFrame2
        .WordWrap = msoTrue
        .AutoSize = msoAutoSizeNone
        .VerticalAnchor = msoAnchorTop
        .MarginLeft = 6
        .MarginRight = 6
        .MarginTop = 3
        .MarginBottom = 3
        .TextRange.Text = ToShapeText(txt)
        .TextRange.ParagraphFormat.Alignment = msoAlignLeft
        With .TextRange.Font
            .Name = fnt
            .NameFarEast = fnt
            .Size = sz
            .Bold = IIf(isBold, msoTrue, msoFalse)
        End With
    End With
End Sub

' 文字だけを入れ替える（ステータスボタン・見出し帯）
Private Sub ReplaceText(ByVal shp As Shape, ByVal txt As String, ByVal sz As Single, ByVal isBold As Boolean)
    Dim fnt As String
    fnt = CfgFont()
    With shp.TextFrame2.TextRange
        .Text = txt
        .Font.Name = fnt
        .Font.NameFarEast = fnt
        .Font.Size = sz
        .Font.Bold = IIf(isBold, msoTrue, msoFalse)
    End With
End Sub

' 文字がちょうど収まる高さを測る。slack > 0 なら、文字を入れる幅をその割合だけ狭めて測る
' （表示の拡大率を下げると文字の幅が少し広がり、折り返しが 1 行増えることがあるので、その分の余裕を見る）
Private Function TextHeight(ByVal shp As Shape, Optional ByVal slack As Single = 0) As Single
    Dim t As Single, mr As Single
    t = shp.Top
    With shp.TextFrame2
        mr = .MarginRight
        If slack > 0 Then .MarginRight = mr + (shp.Width - .MarginLeft - mr) * slack
        .AutoSize = msoAutoSizeShapeToFitText
        TextHeight = shp.Height
        .AutoSize = msoAutoSizeNone
        .MarginRight = mr
    End With
    shp.Top = t
End Function

' 文字が収まる高さにする（最小 minH）
Private Sub FitHeight(ByVal shp As Shape, ByVal minH As Single, Optional ByVal slack As Single = 0)
    Dim t As Single, h As Single
    t = shp.Top
    h = TextHeight(shp, slack)
    If h < minH Then h = minH
    shp.Height = h
    shp.Top = t
End Sub

' タイトルの欄の高さ（タイトルの文字サイズで TitleLines 行分）。設定が同じあいだは測った値を使う
Private Function TitleAreaHeight(ByVal shp As Shape) As Single
    Dim key As String, s As String, i As Long
    key = CfgFont() & "|" & SizeTitle() & "|" & TitleLines() & "|" & Format$(shp.Width, "0.0")
    If key <> mAreaKey Then
        s = "あ"
        For i = 2 To TitleLines()
            s = s & vbLf & "あ"
        Next
        SetText shp, s, SizeTitle(), True
        mAreaH = TextHeight(shp)
        mAreaKey = key
    End If
    TitleAreaHeight = mAreaH
End Function

' タイトルを入れて高さを決める。全文を表示し、欄に入らないときは入るまで文字を小さくする
' （いちばん小さくしても入らないときは、欄を広げて全文を表示する）
Private Sub FitTitle(ByVal shp As Shape, ByVal txt As String)
    Dim areaH As Single, h As Single, lo As Long, hi As Long, md As Long
    areaH = TitleAreaHeight(shp)
    SetText shp, txt, SizeTitle(), True
    h = TextHeight(shp, TEXT_SLACK)
    If h > areaH + TITLE_FIT_TOL Then
        ' 0.5pt 刻みで、欄に入るいちばん大きい文字サイズを探す（lo, hi は pt の 2 倍。hi の大きさでは入らない）
        lo = CLng(TITLE_MIN_SIZE * 2)
        hi = -Int(-SizeTitle() * 2)
        Do While hi - lo > 1
            md = (lo + hi) \ 2
            shp.TextFrame2.TextRange.Font.Size = md / 2
            If TextHeight(shp, TEXT_SLACK) <= areaH + TITLE_FIT_TOL Then lo = md Else hi = md
        Loop
        shp.TextFrame2.TextRange.Font.Size = lo / 2
        h = TextHeight(shp, TEXT_SLACK)
    End If
    If h < TITLE_MIN_H Then h = TITLE_MIN_H
    shp.Height = h
    ' 幅を狭めて測った分、表示では 1 行少なく収まることがあるので、上下の中央に置く
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
End Sub

' 起票日時・開始日・終了日の欄の文字（開始日と終了日は、幅に入れば 1 行に並べる）
Private Function InfoText(ByRef d As NoteData, ByVal w As Single) As String
    Dim s As String, sd As String, ed As String, pair As String
    If d.Created > 0 Then s = Format$(d.Created, "yyyy/mm/dd hh:nn") Else s = NO_VALUE
    s = LBL_CREATED & vbTab & s
    sd = DayText(d.StartDate)
    ed = DayText(d.EndDate)
    pair = sd & "　" & LBL_END & " " & ed
    If INFO_TAB + EstTextWidth(pair, INFO_SIZE) * (1 + TEXT_SLACK) <= w - 12 Then
        InfoText = s & vbLf & LBL_START & vbTab & pair
    Else
        InfoText = s & vbLf & LBL_START & vbTab & sd & vbLf & LBL_END & vbTab & ed
    End If
End Function

Private Function ProgressText(ByVal pct As Long) As String
    ProgressText = LBL_PROGRESS & vbTab & vbTab & pct & "%"
End Function

Private Function DayText(ByVal dt As Date) As String
    If dt = 0 Then DayText = NO_VALUE Else DayText = Format$(dt, "yyyy/mm/dd")
End Function

' 起票日時・進捗率の欄の書式：見出しのあとの値を揃え、進捗率の % は右端に揃える
Private Sub StyleInfo(ByVal shp As Shape, ByVal w As Single, ByVal padTop As Single)
    With shp.TextFrame2
        .MarginTop = padTop
        .MarginBottom = 0
        .TextRange.ParagraphFormat.TabStops.Add msoTabStopLeft, INFO_TAB
        .TextRange.ParagraphFormat.TabStops.Add msoTabStopRight, w - .MarginLeft - .MarginRight
    End With
End Sub

' 見出し（起票日時・開始日・終了日・進捗率）を太字・指定の色にする
Private Sub MarkLabels(ByVal shp As Shape, ByVal c As Long)
    Dim t As String, lb As Variant, p As Long
    t = shp.TextFrame2.TextRange.Text
    For Each lb In Array(LBL_CREATED, LBL_START, LBL_END, LBL_PROGRESS)
        p = InStr(t, lb)
        If p > 0 Then
            With shp.TextFrame2.TextRange.Characters(p, Len(lb)).Font
                .Bold = msoTrue
                .Fill.ForeColor.RGB = c
            End With
        End If
    Next
End Sub

' 見出し帯の文字（番号・起票日。完了の付箋は完了日も）。withDates = False なら番号だけ
Private Function HeaderText(ByRef d As NoteData, ByVal isDone As Boolean, ByVal withDates As Boolean) As String
    Dim s As String
    s = "No." & Format$(d.Id, "0000")
    If withDates Then
        If d.Created > 0 Then s = s & "   " & Format$(d.Created, "m/d")
        If isDone And d.Done > 0 Then s = s & " → " & Format$(d.Done, "m/d")
    End If
    HeaderText = s
End Function

Private Function TagText(ByVal si As Long, ByVal isDone As Boolean) As String
    If isDone Then
        TagText = StatusName(StatusCount())
    Else
        TagText = StatusName(ClampStatus(si)) & " " & ChrW$(&H25B6)
    End If
End Function

' ステータスボタンの幅 tw と文字サイズ sz、見出し帯に日付を出すか withDates
'   ボードの付箋 … どのステータスでも同じ幅（いちばん長いステータス名に合わせる）
'   完了の付箋   … 「完了」の文字に合わせた幅
' 見出し帯の文字と重ならない幅までにする。日付まで出すとボタンの文字が入らないときは番号だけにし、
' それでも入らない文字は小さくする
Private Sub TagMetrics(ByRef d As NoteData, ByVal isDone As Boolean, ByRef tw As Single, ByRef sz As Single, _
                       ByRef withDates As Boolean)
    Dim i As Long, n As Long, w As Single, need As Single, hw As Single, mx As Single
    n = StatusCount()
    If isDone Then
        need = EstTextWidth(TagText(n, True), TAG_SIZE)
    Else
        For i = 1 To n - 1
            w = EstTextWidth(TagText(i, False), TAG_SIZE)
            If w > need Then need = w
        Next
    End If
    tw = need + 14
    If tw < 46 Then tw = 46
    hw = EstTextWidth(HeaderText(d, isDone, True), HDR_SIZE)
    If Not isDone Then
        w = EstTextWidth("No.0000   00/00", HDR_SIZE)     ' 日付の桁数で幅が変わらないように
        If w > hw Then hw = w
    End If
    mx = NoteWidth() - 4 - 6 - hw - 6
    withDates = (need + 6 <= mx)
    If Not withDates Then mx = NoteWidth() - 4 - 6 - EstTextWidth(HeaderText(d, isDone, False), HDR_SIZE) - 6
    If mx < 30 Then mx = 30
    If tw > mx Then tw = mx
    sz = TAG_SIZE
    need = EstTextWidth(TagText(d.StatusIdx, isDone), TAG_SIZE)
    If need > tw - 6 Then
        sz = Int(TAG_SIZE * (tw - 6) / need * 2) / 2
        If sz < TAG_MIN_SIZE Then sz = TAG_MIN_SIZE
    End If
End Sub

'==============================================================
'  付箋を作る
'==============================================================

Private Function AddPart(ByVal ws As Worksheet, ByVal kind As Long, ByVal tok As String, ByVal role As String, _
                         ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single) As Shape
    Dim s As Shape
    If kind = KIND_TEXTBOX Then
        Set s = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
    Else
        Set s = ws.Shapes.AddShape(kind, x, y, w, h)
    End If
    s.Name = tok & role
    s.AlternativeText = "role:" & role
    s.Fill.Visible = msoFalse
    s.Line.Visible = msoFalse
    s.Shadow.Visible = msoFalse
    Set AddPart = s
End Function

' 付箋を作って (x, y) に貼る。isDone = True なら完了（グレー）の付箋
Public Function BuildNote(ByVal ws As Worksheet, ByRef d As NoteData, ByVal x As Single, ByVal y As Single, _
                          ByVal isDone As Boolean) As Shape
    Dim w As Single, tw As Single, tsz As Single, withDates As Boolean, yy As Single, tok As String, nm As String
    Dim i As Long, bx As Single, by As Single, bw As Single, parts As String
    Dim sBase As Shape, sHdr As Shape, sTag As Shape, sTtl As Shape, sBody As Shape, sMemo As Shape, sLbl As Shape
    Dim sInfo As Shape, sProg As Shape, sBar As Shape, sFill As Shape
    Dim grp As Shape

    w = NoteWidth()
    mSeq = mSeq + 1
    tok = "tmp" & Format$(Timer * 100, "0") & "x" & mSeq & "_"

    ' 台紙
    Set sBase = AddPart(ws, msoShapeRectangle, tok, ROLE_BASE, x, y, w, 100)
    SetText sBase, "", SizeBody(), False

    ' 見出し帯（番号・日付）
    TagMetrics d, isDone, tw, tsz, withDates
    Set sHdr = AddPart(ws, msoShapeRectangle, tok, ROLE_HDR, x, y, w, HDR_H)
    SetText sHdr, HeaderText(d, isDone, withDates), HDR_SIZE, False
    With sHdr.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .MarginTop = 0
        .MarginBottom = 0
    End With

    ' ステータスボタン
    Set sTag = AddPart(ws, msoShapeRoundedRectangle, tok, ROLE_TAG, x + w - tw - 4, y + (HDR_H - TAG_H) / 2, tw, TAG_H)
    SetText sTag, TagText(d.StatusIdx, isDone), tsz, True
    With sTag.TextFrame2
        .WordWrap = msoFalse
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 2
        .MarginRight = 2
        .MarginTop = 0
        .MarginBottom = 0
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    sTag.Adjustments.Item(1) = 0.5

    ' タイトル（全文を表示する。欄に入らないときは文字を小さくする）
    yy = y + HDR_H + 2
    Set sTtl = AddPart(ws, KIND_TEXTBOX, tok, ROLE_TITLE, x, yy, w, TITLE_MIN_H)
    FitTitle sTtl, d.Title
    yy = yy + sTtl.Height

    ' 内容
    Set sBody = AddPart(ws, KIND_TEXTBOX, tok, ROLE_BODY, x, yy, w, BODY_MIN_H)
    SetText sBody, d.Body, SizeBody(), False
    FitHeight sBody, BODY_MIN_H, TEXT_SLACK
    yy = yy + sBody.Height + 2

    ' 備考（上に「備考」の小さな文字を重ねる）
    Set sMemo = AddPart(ws, KIND_TEXTBOX, tok, ROLE_MEMO, x, yy, w, MEMO_MIN_H)
    SetText sMemo, d.Memo, SizeMemo(), False
    sMemo.TextFrame2.MarginTop = MEMO_LABEL_H + 2
    FitHeight sMemo, MEMO_MIN_H, TEXT_SLACK
    Set sLbl = AddPart(ws, KIND_TEXTBOX, tok, ROLE_LABEL, x, yy, 60, MEMO_LABEL_H + 2)
    SetText sLbl, "備考", 7, True
    With sLbl.TextFrame2
        .MarginTop = 2
        .MarginBottom = 0
    End With
    yy = yy + sMemo.Height

    ' 起票日時・開始日・終了日
    Set sInfo = AddPart(ws, KIND_TEXTBOX, tok, ROLE_INFO, x, yy, w, 30)
    SetText sInfo, InfoText(d, w), INFO_SIZE, False
    StyleInfo sInfo, w, 4
    FitHeight sInfo, 0
    yy = yy + sInfo.Height

    ' 進捗率（文字の間にバーを置く）
    Set sProg = AddPart(ws, KIND_TEXTBOX, tok, ROLE_PROG, x, yy, w, 14)
    SetText sProg, ProgressText(ClampPct(d.Progress)), INFO_SIZE, False
    StyleInfo sProg, w, 0
    FitHeight sProg, 0
    bx = x + sProg.TextFrame2.MarginLeft + INFO_TAB
    bw = w - sProg.TextFrame2.MarginLeft - sProg.TextFrame2.MarginRight - INFO_TAB - PCT_W
    by = sProg.Top + (sProg.Height - PBAR_H) / 2
    Set sBar = AddPart(ws, msoShapeRectangle, tok, ROLE_PBAR, bx, by, bw, PBAR_H)
    parts = Join(Array(sBase.Name, sHdr.Name, sTag.Name, sTtl.Name, sBody.Name, sMemo.Name, sLbl.Name, _
                       sInfo.Name, sProg.Name, sBar.Name), vbTab)
    If ClampPct(d.Progress) > 0 Then
        Set sFill = AddPart(ws, msoShapeRectangle, tok, ROLE_PFILL, bx, by, bw * ClampPct(d.Progress) / 100, PBAR_H)
        parts = parts & vbTab & sFill.Name
    End If
    yy = yy + sProg.Height + 5

    sBase.Height = yy - y

    ' グループ化して名前と情報を付ける
    ' （Shapes.Range には String の配列を渡す。Variant の変数に入れた配列を渡すとエラー 1004 になる）
    Set grp = ws.Shapes.Range(Split(parts, vbTab)).Group
    nm = NoteName(d.Id)
    grp.Name = nm
    For i = 1 To grp.GroupItems.Count
        grp.GroupItems(i).Name = nm & "_" & RoleOf(grp.GroupItems(i))
    Next
    grp.Placement = xlMove
    grp.OnAction = NOTE_MACRO   ' グループに設定すると、すべての部品に設定される
    d.Sig = NoteSig(grp)
    WriteMeta grp, d
    PaintNote grp, d.StatusIdx, isDone
    Set BuildNote = grp
End Function

'==============================================================
'  色を付ける
'==============================================================

Public Sub PaintNote(ByVal grp As Shape, ByVal si As Long, ByVal isDone As Boolean)
    Dim i As Long, s As Shape, c As Long, txt As Long, hc As Long, tc As Long
    If isDone Then si = StatusCount()
    si = ClampStatus(si)
    c = StatusColor(si)
    hc = ShadeColor(c, 0.1)
    tc = ShadeColor(c, 0.5)
    If isDone Then txt = RGB(128, 128, 128) Else txt = TextColorOn(c)
    For i = 1 To grp.GroupItems.Count
        Set s = grp.GroupItems(i)
        Select Case RoleOf(s)
        Case ROLE_BASE
            FillWith s, c
            s.Line.Visible = msoTrue
            s.Line.ForeColor.RGB = ShadeColor(c, 0.2)
            s.Line.Weight = 0.75
            With s.Shadow
                .Visible = msoTrue
                .ForeColor.RGB = RGB(0, 0, 0)
                .Transparency = 0.72
                .OffsetX = 1.5
                .OffsetY = 2.5
                .Blur = 4
            End With
        Case ROLE_HDR
            FillWith s, hc
            If isDone Then ColorText s, txt Else ColorText s, TextColorOn(hc)
        Case ROLE_TAG
            FillWith s, tc
            ColorText s, TextColorOn(tc)
        Case ROLE_TITLE, ROLE_BODY
            ColorText s, txt
        Case ROLE_MEMO
            FillWith s, TintColor(c, 0.55)
            ColorText s, txt
        Case ROLE_LABEL
            ColorText s, ShadeColor(c, 0.55)
        Case ROLE_INFO, ROLE_PROG
            ColorText s, txt
            MarkLabels s, ShadeColor(c, 0.55)
        Case ROLE_PBAR
            FillWith s, TintColor(c, 0.65)
            s.Line.Visible = msoTrue
            s.Line.ForeColor.RGB = ShadeColor(c, 0.3)
            s.Line.Weight = 0.5
        Case ROLE_PFILL
            FillWith s, tc
        End Select
    Next
End Sub

Private Sub FillWith(ByVal s As Shape, ByVal c As Long)
    s.Fill.Visible = msoTrue
    s.Fill.Solid
    s.Fill.ForeColor.RGB = c
    s.Fill.Transparency = 0
End Sub

Private Sub ColorText(ByVal s As Shape, ByVal c As Long)
    s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = c
End Sub

' ステータスが変わったときの見た目の更新（作り直さずに色と文字だけ変える）
Public Sub RestyleNote(ByVal grp As Shape, ByRef d As NoteData, ByVal isDone As Boolean)
    Dim s As Shape, b As Shape, tw As Single, tsz As Single, withDates As Boolean
    TagMetrics d, isDone, tw, tsz, withDates
    Set s = PartOf(grp, ROLE_TAG)
    If Not s Is Nothing Then
        ReplaceText s, TagText(d.StatusIdx, isDone), tsz, True
        ' 完了にしたとき・ステータス名を変えたときは、幅を合わせて右端に置き直す
        If Abs(s.Width - tw) > 1 Then
            Set b = PartOf(grp, ROLE_BASE)
            s.Width = tw
            If Not b Is Nothing Then s.Left = b.Left + b.Width - tw - 4
        End If
    End If
    Set s = PartOf(grp, ROLE_HDR)
    If Not s Is Nothing Then ReplaceText s, HeaderText(d, isDone, withDates), HDR_SIZE, False
    WriteMeta grp, d
    PaintNote grp, d.StatusIdx, isDone
End Sub

' シート上の付箋を同じ位置に作り直す
'   onlyChanged = True  … 文字を直接書き直した付箋・大きさが変わった付箋だけ（はみ出しを直す）
'   onlyChanged = False … すべて（設定の変更を反映する）
Public Sub RebuildNotes(ByVal ws As Worksheet, ByVal isDone As Boolean, Optional ByVal onlyChanged As Boolean = False)
    Dim notes As Collection, g As Shape, d As NoteData, x As Single, y As Single, n As Long, doIt As Boolean
    n = StatusCount()
    Set notes = NotesOn(ws)
    For Each g In notes
        doIt = True
        If onlyChanged Then doIt = NeedsRefit(g)
        If doIt Then
            d = ReadNote(g)
            If Not isDone And d.StatusIdx >= n Then d.StatusIdx = n - 1
            x = g.Left
            y = g.Top
            g.Delete
            BuildNote ws, d, x, y, isDone
        End If
    Next
End Sub

'==============================================================
'  シートの保護（付箋の部品を直接さわれないようにする）
'==============================================================

' 図形だけを保護する（セルは保護しない）。マクロからは変更できるように UserInterfaceOnly で保護する。
' UserInterfaceOnly はブックを閉じると消えるので、開いたときと操作のたびにかけ直す
Public Sub ProtectNotes(ByVal ws As Worksheet)
    If ws.ProtectDrawingObjects And ws.ProtectionMode Then Exit Sub
    ws.Protect DrawingObjects:=True, Contents:=False, Scenarios:=False, UserInterfaceOnly:=True
End Sub

Public Sub ProtectBoards()
    ProtectNotes shBoard
    ProtectNotes shDone
End Sub

'==============================================================
'  位置
'==============================================================

Public Function OriginX(ByVal ws As Worksheet) As Single
    OriginX = ws.Columns(2).Left + 8
End Function

Public Function OriginY(ByVal ws As Worksheet) As Single
    OriginY = ws.Rows(FIRST_ROW).Top + 10
End Function

' 列（ステータスの見出し）の左端
Public Function LaneLeft(ByVal ws As Worksheet, ByVal lane As Long) As Single
    LaneLeft = OriginX(ws) + (lane - 1) * (NoteWidth() + GAP_X)
End Function

Private Function Overlaps(ByVal x1 As Single, ByVal y1 As Single, ByVal w1 As Single, ByVal h1 As Single, _
                          ByVal x2 As Single, ByVal y2 As Single, ByVal w2 As Single, ByVal h2 As Single) As Boolean
    Overlaps = (x1 < x2 + w2) And (x2 < x1 + w1) And (y1 < y2 + h2) And (y2 < y1 + h1)
End Function

' x の位置で、ほかの付箋と重ならない一番上の高さを探す
Public Function FreeTop(ByVal ws As Worksheet, ByVal x As Single, ByVal w As Single, ByVal h As Single, _
                        ByVal self As Shape) As Single
    Dim notes As Collection, g As Shape, y As Single, moved As Boolean, selfName As String
    Set notes = NotesOn(ws)
    If Not self Is Nothing Then selfName = self.Name
    y = OriginY(ws)
    Do
        moved = False
        For Each g In notes
            If g.Name <> selfName Then
                If Overlaps(x, y - GAP_Y + 1, w, h + GAP_Y * 2 - 2, g.Left, g.Top, g.Width, g.Height) Then
                    y = g.Top + g.Height + GAP_Y
                    moved = True
                End If
            End If
        Next
    Loop While moved
    FreeTop = y
End Function

Public Function HitsOther(ByVal ws As Worksheet, ByVal self As Shape) As Boolean
    Dim g As Shape
    For Each g In NotesOn(ws)
        If g.Name <> self.Name Then
            If Overlaps(self.Left, self.Top, self.Width, self.Height, g.Left, g.Top, g.Width, g.Height) Then
                HitsOther = True
                Exit Function
            End If
        End If
    Next
End Function

' 指定した列の空いている場所に付箋を貼る
Public Function PlaceInLane(ByVal ws As Worksheet, ByRef d As NoteData, ByVal lane As Long) As Shape
    Dim g As Shape, x As Single
    x = LaneLeft(ws, lane)
    Set g = BuildNote(ws, d, x, OriginY(ws), False)
    g.Top = FreeTop(ws, x, g.Width, g.Height, g)
    Set PlaceInLane = g
End Function

' タスクボード：ステータスごとに見出しの下へ並べる（列の中の上下の順番は今のまま）
'   refit = True なら、文字を直接書き直した付箋の大きさも整える
Public Sub ArrangeBoard(ByVal refit As Boolean)
    Dim ws As Worksheet, notes As Collection, g As Shape
    Dim n As Long, cnt As Long, i As Long, lane As Long, y As Single
    Dim nm() As String, ln() As Long, ky() As Double
    Set ws = shBoard
    If refit Then RebuildNotes ws, False, True
    Set notes = NotesOn(ws)
    cnt = notes.Count
    If cnt = 0 Then Exit Sub
    n = StatusCount()
    ReDim nm(1 To cnt)
    ReDim ln(1 To cnt)
    ReDim ky(1 To cnt)
    i = 0
    For Each g In notes
        i = i + 1
        nm(i) = g.Name
        ln(i) = NoteStatus(g)
        If ln(i) >= n Then ln(i) = n - 1
        ky(i) = CDbl(g.Top) * 100000# + g.Left
    Next
    SortByKey nm, ln, ky
    For lane = 1 To n - 1
        y = OriginY(ws)
        For i = 1 To cnt
            If ln(i) = lane Then
                Set g = ws.Shapes(nm(i))
                g.Left = LaneLeft(ws, lane)
                g.Top = y
                y = y + g.Height + GAP_Y
            End If
        Next
    Next
End Sub

' 完了シート：完了した日時の新しい順に並べる
'   refit = True なら、文字を直接書き直した付箋の大きさも整える
Public Sub ArrangeDoneSheet(ByVal refit As Boolean)
    Dim ws As Worksheet, notes As Collection, g As Shape, m As String
    Dim cnt As Long, i As Long, col As Long, y As Single, rowH As Single
    Dim nm() As String, ln() As Long, ky() As Double
    Set ws = shDone
    If refit Then RebuildNotes ws, True, True
    Set notes = NotesOn(ws)
    cnt = notes.Count
    If cnt = 0 Then Exit Sub
    ReDim nm(1 To cnt)
    ReDim ln(1 To cnt)
    ReDim ky(1 To cnt)
    i = 0
    For Each g In notes
        i = i + 1
        m = g.AlternativeText
        nm(i) = g.Name
        ky(i) = -(CDbl(ParseIso(MetaGet(m, "dn"))) * 86400# + Val(MetaGet(m, "id")) / 100000#)
    Next
    SortByKey nm, ln, ky
    y = OriginY(ws)
    For i = 1 To cnt
        Set g = ws.Shapes(nm(i))
        g.Left = LaneLeft(ws, col + 1)
        g.Top = y
        If g.Height > rowH Then rowH = g.Height
        col = col + 1
        If col >= DONE_COLS Then
            col = 0
            y = y + rowH + GAP_Y
            rowH = 0
        End If
    Next
End Sub

Private Sub SortByKey(ByRef nm() As String, ByRef ln() As Long, ByRef ky() As Double)
    Dim i As Long, j As Long, tn As String, tl As Long, tk As Double
    For i = LBound(ky) + 1 To UBound(ky)
        tn = nm(i)
        tl = ln(i)
        tk = ky(i)
        j = i - 1
        Do While j >= LBound(ky)
            If ky(j) <= tk Then Exit Do
            nm(j + 1) = nm(j)
            ln(j + 1) = ln(j)
            ky(j + 1) = ky(j)
            j = j - 1
        Loop
        nm(j + 1) = tn
        ln(j + 1) = tl
        ky(j + 1) = tk
    Next
End Sub

Public Sub ScrollIntoView(ByVal g As Shape)
    Dim r As Long, vr As Range
    On Error Resume Next
    r = g.TopLeftCell.Row
    With ActiveWindow
        Set vr = .Panes(.Panes.Count).VisibleRange
        If r < vr.Row Or r > vr.Row + vr.Rows.Count - 4 Then
            If r - 1 > FIRST_ROW Then .ScrollRow = r - 1 Else .ScrollRow = FIRST_ROW
        End If
    End With
End Sub

'==============================================================
'  番号
'==============================================================

Public Function NextNoteId() As Long
    Dim nmObj As Name, n As Long, mx As Long
    Set nmObj = NextIdName()
    n = CLng(Val(Mid$(nmObj.RefersTo, 2)))
    mx = MaxNoteId()
    If n <= mx Then n = mx + 1
    If n < 1 Then n = 1
    nmObj.RefersTo = "=" & (n + 1)
    NextNoteId = n
End Function

Private Function NextIdName() As Name
    Dim nmObj As Name
    On Error Resume Next
    Set nmObj = ThisWorkbook.Names("stkNextId")
    On Error GoTo 0
    If nmObj Is Nothing Then
        Set nmObj = ThisWorkbook.Names.Add(Name:="stkNextId", RefersTo:="=1", Visible:=False)
    End If
    Set NextIdName = nmObj
End Function

Private Function MaxNoteId() As Long
    Dim g As Shape, m As Long, id As Long
    For Each g In NotesOn(shBoard)
        id = NoteId(g)
        If id > m Then m = id
    Next
    For Each g In NotesOn(shDone)
        id = NoteId(g)
        If id > m Then m = id
    Next
    MaxNoteId = m
End Function

' コピー＆貼り付けで番号や名前が重なった付箋に、新しい番号を付け直す
Public Sub NormalizeNotes()
    Dim seen As Object, k As Long, ws As Worksheet, notes As Collection, g As Shape, id As Long, d As NoteData
    Set seen = CreateObject("Scripting.Dictionary")
    For k = 1 To 2
        If k = 1 Then Set ws = shBoard Else Set ws = shDone
        Set notes = NotesOn(ws)
        For Each g In notes
            id = NoteId(g)
            If id <= 0 Or seen.Exists(id) Then
                d = ReadNote(g)
                d.Id = NextNoteId()
                RenameNote g, d.Id
                RestyleNote g, d, (k = 2)
                id = d.Id
            ElseIf g.Name <> NoteName(id) Then
                RenameNote g, id
            End If
            seen(id) = True
        Next
    Next
End Sub

Private Sub RenameNote(ByVal grp As Shape, ByVal id As Long)
    Dim i As Long, nm As String
    nm = NoteName(id)
    For i = 1 To grp.GroupItems.Count
        grp.GroupItems(i).Name = nm & "_" & RoleOf(grp.GroupItems(i))
    Next
    grp.Name = nm
End Sub

'==============================================================
'  ステータスの見出し（枚数の表示）
'==============================================================

Public Function LaneCaption(ByVal i As Long, ByVal cnt As Long) As String
    LaneCaption = StatusName(i) & "（" & cnt & "）"
End Function

Public Sub RefreshLaneCounts()
    Dim n As Long, i As Long, cnt() As Long, g As Shape, s As Shape, txt As String
    n = StatusCount()
    If n < 1 Then Exit Sub
    ReDim cnt(1 To n)
    For Each g In NotesOn(shBoard)
        i = NoteStatus(g)
        cnt(i) = cnt(i) + 1
    Next
    cnt(n) = NotesOn(shDone).Count
    For i = 1 To n
        Set s = Nothing
        On Error Resume Next
        Set s = shBoard.Shapes(LANE_PREFIX & i)
        On Error GoTo 0
        If Not s Is Nothing Then
            txt = LaneCaption(i, cnt(i))
            If s.TextFrame2.TextRange.Text <> txt Then s.TextFrame2.TextRange.Text = txt
        End If
    Next
End Sub

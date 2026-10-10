Option Explicit
'==============================================================
'  modSetup : シートの書式・ボタン・サンプルの作成
'             SetupWorkbook はブックを作るとき（build.ps1）に一度だけ呼ばれる
'             BuildToolbars は［設定を反映］からも呼ばれる
'==============================================================

Private Const BTN_H As Single = 24
Private Const LAST_ROW As Long = 1000

Private mHelp As Collection

'==============================================================
'  初期設定（ビルド時）
'==============================================================

Public Function SetupWorkbook() As String
    On Error GoTo EH
    Application.ScreenUpdating = False
    ApplyFontScheme
    SetupConfigSheet
    ConfigRefresh
    SetupBoardSheet
    SetupDoneSheet
    SetupHelpSheet
    BuildToolbars
    CreateSamples
    RefreshLaneCounts
    ProtectBoards
    shConfig.Activate
    shConfig.Range("B5").Select
    shDone.Activate
    shDone.Range("A1").Select
    shBoard.Activate
    shBoard.Range("A1").Select
    Application.ScreenUpdating = True
    Exit Function
EH:
    SetupWorkbook = "ERR " & Err.Number & ": " & Err.Description
    Application.ScreenUpdating = True
End Function

' ブックのテーマのフォントを Meiryo UI にする（セル・新しい図形の既定のフォント）
Private Sub ApplyFontScheme()
    Dim p As String, f As Integer, fnt As String, faces As String
    fnt = DEFAULT_FONT
    faces = "<a:latin typeface=""" & fnt & """/><a:ea typeface=""" & fnt & """/><a:cs typeface=""""/>" & _
            "<a:font script=""Jpan"" typeface=""" & fnt & """/>"
    p = Environ$("TEMP") & "\stk_fontscheme_" & Format$(Timer * 100, "0") & ".xml"
    f = FreeFile
    Open p For Output As #f
    Print #f, "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
    Print #f, "<a:fontScheme xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" name=""" & fnt & """>"
    Print #f, "<a:majorFont>" & faces & "</a:majorFont>"
    Print #f, "<a:minorFont>" & faces & "</a:minorFont>"
    Print #f, "</a:fontScheme>"
    Close #f
    ThisWorkbook.Theme.ThemeFontScheme.Load p
    Kill p
    With ThisWorkbook.Styles("Normal").Font
        .Name = fnt
        .Size = 10
    End With
End Sub

'--- 設定シート ----------------------------------------------------

Private Sub SetupConfigSheet()
    Dim ws As Worksheet, i As Long, st As Variant, cl As Variant, ds As Variant, items As Variant
    Set ws = shConfig
    ws.Name = "設定"
    ws.Tab.Color = RGB(128, 128, 128)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 16
    ws.Columns("C").ColumnWidth = 9
    ws.Columns("D").ColumnWidth = 60
    ws.Columns("E").ColumnWidth = 3
    ws.Columns("F").ColumnWidth = 22
    ws.Columns("G").ColumnWidth = 14
    ws.Rows(1).RowHeight = 36
    With ws.Range("B1")
        .Value = "設定"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(64, 64, 64)
        .VerticalAlignment = xlCenter
    End With
    ws.Range("B2").Value = "ステータスの名前と付箋の色を変えられます。色は「色」の欄のセルの塗りつぶしの色で指定します（［ホーム］→［塗りつぶしの色］）。"
    ws.Range("B3").Value = "上から順にステータスが進み、いちばん下のステータスが「完了」です（付箋はグレーアウトして完了シートへ移動）。行を挿入すると10個まで増やせます。変更したら［設定を反映］を押してください。"
    ws.Range("B2:B3").Font.Color = RGB(89, 89, 89)
    ws.Rows("2:3").RowHeight = 18

    ' ステータス
    ws.Range("B4:D4").Value = Array("ステータス名", "色", "説明")
    st = Array("起票済み", "一般タスク対応中", "一般タスク対応済み", _
               "レビュー依頼受領済み", "レビュー対応中", "レビュー済み対応依頼中", "レビュー済み", _
               "連携済み", "完了")
    cl = Array(RGB(255, 241, 140), RGB(255, 196, 214), RGB(176, 222, 250), _
               RGB(216, 204, 248), RGB(255, 207, 153), RGB(250, 172, 160), RGB(168, 230, 212), _
               RGB(196, 230, 164), RGB(217, 217, 217))
    ds = Array("タスクを書き出した状態。新しい付箋はこの色で貼られます。", _
               "一般タスクの対応を進めている状態。", _
               "一般タスクの対応が終わった状態。", _
               "レビューの依頼を受け取った状態。", _
               "レビューを進めている状態。", _
               "レビューが終わり、指摘への対応を依頼している状態。", _
               "レビューが終わった状態。", _
               "関係する人へ連携した状態。", _
               "完了。付箋はグレーアウトして「完了」シートへ移動します。")
    For i = 0 To UBound(st)
        ws.Cells(5 + i, 2).Value = st(i)
        ws.Cells(5 + i, 3).Interior.Color = cl(i)
        ws.Cells(5 + i, 4).Value = ds(i)
    Next
    StyleHeader ws.Range("B4:D4")
    StyleTable ws.Range(ws.Cells(5, 2), ws.Cells(5 + UBound(st), 4))

    ' その他
    ws.Range("F4:G4").Value = Array("項目", "値")
    items = Array("フォント", "付箋の幅（pt）", "タイトルの文字サイズ", "タイトルの最大行数", _
                  "内容の文字サイズ", "備考の文字サイズ", "完了時に確認する")
    For i = 0 To UBound(items)
        ws.Cells(5 + i, 6).Value = items(i)
    Next
    ws.Range("G5").Value = DEFAULT_FONT
    ws.Range("G6").Value = 180
    ws.Range("G7").Value = 10.5
    ws.Range("G8").Value = 2
    ws.Range("G9").Value = 9
    ws.Range("G10").Value = 8
    ws.Range("G11").Value = "はい"
    With ws.Range("G11").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="はい," & CONFIRM_NO
    End With
    StyleHeader ws.Range("F4:G4")
    StyleTable ws.Range("F5:G11")
    ws.Range("G5:G11").HorizontalAlignment = xlLeft

    AddName "cfgStatusTop", ws.Range("B5")
    AddName "cfgFont", ws.Range("G5")
    AddName "cfgWidth", ws.Range("G6")
    AddName "cfgSizeTitle", ws.Range("G7")
    AddName "cfgTitleLines", ws.Range("G8")
    AddName "cfgSizeBody", ws.Range("G9")
    AddName "cfgSizeMemo", ws.Range("G10")
    AddName "cfgConfirm", ws.Range("G11")

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 100
End Sub

Private Sub AddName(ByVal nm As String, ByVal r As Range)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nm, RefersTo:="='" & r.Worksheet.Name & "'!" & r.Address
End Sub

Private Sub StyleHeader(ByVal r As Range)
    With r
        .Interior.Color = RGB(89, 89, 89)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .VerticalAlignment = xlCenter
        .RowHeight = 22
    End With
    BorderAll r
End Sub

Private Sub StyleTable(ByVal r As Range)
    r.VerticalAlignment = xlCenter
    r.RowHeight = 24
    BorderAll r
End Sub

Private Sub BorderAll(ByVal r As Range)
    Dim b As Variant
    For Each b In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
        With r.Borders(b)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(191, 191, 191)
        End With
    Next
End Sub

'--- タスクボード・完了シート（ノートの罫線風） ---------------------

Private Sub SetupBoardSheet()
    shBoard.Name = "タスクボード"
    shBoard.Tab.Color = RGB(68, 114, 196)
    SetupPaper shBoard, "付箋タスクボード", RGB(214, 228, 240), RGB(244, 186, 186)
End Sub

Private Sub SetupDoneSheet()
    shDone.Name = "完了"
    shDone.Tab.Color = RGB(166, 166, 166)
    SetupPaper shDone, "完了した付箋", RGB(226, 226, 226), RGB(210, 210, 210)
    With shDone.Range("B2")
        .Value = "完了した付箋はグレーアウトして、新しい順に並びます。付箋を選んで［ボードに戻す］［捨てる］ができます。"
        .Font.Size = 9
        .Font.Color = RGB(118, 118, 118)
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub SetupPaper(ByVal ws As Worksheet, ByVal title As String, ByVal lineColor As Long, ByVal marginColor As Long)
    ws.Columns(1).ColumnWidth = 1.5
    ws.Rows(1).RowHeight = 36
    ws.Rows(2).RowHeight = 28
    ws.Rows(3).RowHeight = 6
    ws.Rows("1:2").Interior.Color = RGB(242, 242, 242)
    With ws.Rows(3).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(191, 191, 191)
    End With
    With ws.Range("B1")
        .Value = title
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(64, 64, 64)
        .VerticalAlignment = xlCenter
    End With
    With ws.Rows(FIRST_ROW & ":" & LAST_ROW)
        .RowHeight = 18
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = lineColor
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = lineColor
        End With
    End With
    With ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(LAST_ROW, 1)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = marginColor
    End With
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .DisplayGridlines = False
        .DisplayHeadings = False
        .Zoom = 100
        .ScrollRow = 1
        .ScrollColumn = 1
        .SplitColumn = 0
        .SplitRow = FIRST_ROW - 1
        .FreezePanes = True
    End With
End Sub

'--- 使い方シート --------------------------------------------------

Private Sub HelpLine(ByVal s As String)
    mHelp.Add s
End Sub

Private Sub SetupHelpSheet()
    Dim ws As Worksheet, r As Long, s As Variant, nx As String, pv As String
    nx = ChrW$(&H25B6)
    pv = ChrW$(&H25C0)
    Set mHelp = New Collection
    HelpLine "# 付箋タスクボードの使い方"
    HelpLine ""
    HelpLine "■ 付箋を貼る"
    HelpLine "・「タスクボード」シートの［＋ 新しい付箋］を押して、タイトル・内容・備考（関係する実装など）を入力し、［OK］を押します。"
    HelpLine "・備考の下の欄で、開始日・終了日・進捗率も入力できます（空欄のままでもかまいません。［今日］で今日の日付が入ります）。"
    HelpLine "・黄色の付箋（起票済み）がボードに貼られます。付箋はドラッグして好きな場所へ動かせます。"
    HelpLine "・付箋の備考の下に、起票日時（付箋を作った日時。自動で入ります）・開始日・終了日・進捗率が表示されます。"
    HelpLine "・タイトルは付箋に全文が表示されます。タイトルの欄（「設定」シートの「タイトルの最大行数」）に入りきらないときは、文字が自動で小さくなります。"
    HelpLine ""
    HelpLine "■ ステータスを進める（付箋の色が変わります）"
    HelpLine "・付箋の右上のボタン（例：「起票済み " & nx & "」）をクリックすると、次のステータスに進んで付箋の色が変わります。"
    HelpLine "　　起票済み（黄）→ 一般タスク対応中（ピンク）→ 一般タスク対応済み（水色）→ レビュー依頼受領済み（紫）→ レビュー対応中（オレンジ）"
    HelpLine "　　→ レビュー済み対応依頼中（サーモン）→ レビュー済み（ミント）→ 連携済み（緑）→ 完了（グレー）　※初期設定の色"
    HelpLine "・付箋を選んでから、上の見出し（起票済み・一般タスク対応中 …）をクリックすると、そのステータスに直接変えられます。"
    HelpLine "　　使わないステータスを飛ばすときに使います（例：一般タスクの付箋は「一般タスク対応済み」から「連携済み」の見出しへ）。"
    HelpLine "・間違えたときは、付箋を選んで［" & pv & " 戻す］を押します。［進める " & nx & "］で選んだ付箋をまとめて進めることもできます。"
    HelpLine ""
    HelpLine "■ 完了"
    HelpLine "・「完了」になった付箋はグレーアウトして、自動で「完了」シートへ移動します（新しい順に並びます）。"
    HelpLine "・「完了」シートで付箋を選んで［" & pv & " ボードに戻す］を押すと、「連携済み」に戻ってボードの元の場所へ戻ります。"
    HelpLine "・いらなくなった付箋は、選んで［捨てる］を押すと削除できます。"
    HelpLine ""
    HelpLine "■ 書き直す"
    HelpLine "・付箋を選んで［編集］を押すと、入力画面で書き直せます（開始日・終了日・進捗率・ステータスも変えられます）。"
    HelpLine "・付箋の文字を直接書き換えたり、部品を動かしたりはできません（付箋がくずれないよう、シートを保護しています）。"
    HelpLine ""
    HelpLine "■ 整列"
    HelpLine "・［整列］を押すと、付箋がステータスごとに見出しの下へ並びます。見出しのかっこ内は付箋の枚数です。"
    HelpLine ""
    HelpLine "■ 付箋の選び方・動かし方"
    HelpLine "・付箋は1回クリックすると選べます。Shift キーを押しながらクリックすると、複数の付箋を選べます（もう一度で外れます）。"
    HelpLine "・複数の付箋を選んでからドラッグすると、まとめて動かせます。"
    HelpLine ""
    HelpLine "■ 設定"
    HelpLine "・「設定」シートで、ステータスの名前・順番・付箋の色・フォント（初期値：Meiryo UI）・文字の大きさなどを変えられます。"
    HelpLine "・変更したら、「設定」シートの［設定を反映］を押してください。"
    HelpLine "・ボタンや見出しを消してしまったときも、［設定を反映］を押すと作り直されます。"
    HelpLine ""
    HelpLine "■ ご注意"
    HelpLine "・このブックはマクロを使っています。開いたときに［コンテンツの有効化］を押してください。"
    HelpLine "・「タスクボード」「完了」シートは、付箋がくずれないように保護しています（解除しても次の操作で保護し直します）。"
    HelpLine "・保存するときは「Excel マクロ有効ブック（*.xlsm）」のまま保存してください。"

    Set ws = shHelp
    ws.Name = "使い方"
    ws.Tab.Color = RGB(112, 173, 71)
    ws.Columns(1).ColumnWidth = 2
    ws.Columns(2).ColumnWidth = 120
    r = 1
    For Each s In mHelp
        With ws.Cells(r, 2)
            If Left$(s, 2) = "# " Then
                .Value = Mid$(s, 3)
                .Font.Size = 14
                .Font.Bold = True
                .Font.Color = RGB(64, 64, 64)
                ws.Rows(r).RowHeight = 34
            ElseIf Left$(s, 1) = "■" Then
                .Value = s
                .Font.Size = 11
                .Font.Bold = True
                .Font.Color = RGB(47, 84, 150)
                ws.Rows(r).RowHeight = 24
            Else
                .Value = s
                ws.Rows(r).RowHeight = 18
            End If
            .VerticalAlignment = xlCenter
        End With
        r = r + 1
    Next
    Set mHelp = Nothing
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.Zoom = 100
    ws.Range("A1").Select
End Sub

'==============================================================
'  ボタン（図形）
'==============================================================

Public Sub BuildToolbars()
    Dim ws As Worksheet, x As Single, y As Single, n As Long, i As Long, w As Single
    Dim nx As String, pv As String
    nx = ChrW$(&H25B6)
    pv = ChrW$(&H25C0)

    ' タスクボード
    Set ws = shBoard
    DeleteShapes ws, "TB_"
    DeleteShapes ws, LANE_PREFIX
    y = ws.Rows(1).Top + (ws.Rows(1).Height - BTN_H) / 2
    x = ws.Columns(2).Left + 136
    x = AddButton(ws, "TB_New", "＋ 新しい付箋", "Board_NewNote", x, y, 112, True)
    x = AddButton(ws, "TB_Edit", "編集", "Board_EditSelected", x, y, 56, False)
    x = AddButton(ws, "TB_Back", pv & " 戻す", "Board_StepBack", x, y, 66, False)
    x = AddButton(ws, "TB_Fwd", "進める " & nx, "Board_StepForward", x, y, 74, False)
    x = AddButton(ws, "TB_Arrange", "整列", "Board_Arrange", x, y, 56, False)
    x = AddButton(ws, "TB_Discard", "捨てる", "Board_Discard", x, y, 62, False)
    AddHint ws, "TB_Hint", "付箋の右上のボタンでステータスが進みます（Shift+クリックで複数選択）。", x + 6, y
    n = StatusCount()
    y = ws.Rows(2).Top + 3
    For i = 1 To n
        If i < n Then w = NoteWidth() Else w = 120
        AddLane ws, i, LaneLeft(ws, i), y, w, 22
    Next

    ' 完了シート
    Set ws = shDone
    DeleteShapes ws, "TB_"
    y = ws.Rows(1).Top + (ws.Rows(1).Height - BTN_H) / 2
    x = ws.Columns(2).Left + 136
    x = AddButton(ws, "TB_Restore", pv & " ボードに戻す", "Done_Restore", x, y, 116, False)
    x = AddButton(ws, "TB_Arrange", "整列", "Done_Arrange", x, y, 56, False)
    x = AddButton(ws, "TB_Discard", "捨てる", "Done_Discard", x, y, 62, False)

    ' 設定シート
    Set ws = shConfig
    DeleteShapes ws, "TB_"
    y = ws.Rows(1).Top + (ws.Rows(1).Height - BTN_H) / 2
    AddButton ws, "TB_Apply", "設定を反映", "Config_Apply", ws.Columns(3).Left, y, 104, True
End Sub

Private Function AddButton(ByVal ws As Worksheet, ByVal nm As String, ByVal caption As String, ByVal macro As String, _
                           ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal primary As Boolean) As Single
    Dim s As Shape, fc As Long, lc As Long, tc As Long
    If primary Then
        fc = RGB(47, 117, 181)
        lc = RGB(47, 117, 181)
        tc = RGB(255, 255, 255)
    Else
        fc = RGB(255, 255, 255)
        lc = RGB(180, 180, 180)
        tc = RGB(64, 64, 64)
    End If
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, BTN_H)
    s.Name = nm
    s.Placement = xlFreeFloating
    s.Adjustments.Item(1) = 0.18
    s.Fill.Visible = msoTrue
    s.Fill.Solid
    s.Fill.ForeColor.RGB = fc
    s.Line.Visible = msoTrue
    s.Line.ForeColor.RGB = lc
    s.Line.Weight = 0.75
    s.Shadow.Visible = msoFalse
    StyleCaption s, caption, 9.5, True, tc
    s.OnAction = macro
    AddButton = x + w + 6
End Function

Private Sub AddLane(ByVal ws As Worksheet, ByVal i As Long, ByVal x As Single, ByVal y As Single, _
                    ByVal w As Single, ByVal h As Single)
    Dim s As Shape, c As Long, sz As Single, need As Single
    c = StatusColor(i)
    ' 長いステータス名が入りきらないときは文字を小さくする（枚数は 2 桁まで入るようにする）
    sz = 9.5
    need = EstTextWidth(LaneCaption(i, 99), sz)
    If need > w - 8 Then
        sz = Int(sz * (w - 8) / need * 2) / 2
        If sz < 6 Then sz = 6
    End If
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    s.Name = LANE_PREFIX & i
    s.Placement = xlFreeFloating
    s.Adjustments.Item(1) = 0.3
    s.Fill.Visible = msoTrue
    s.Fill.Solid
    s.Fill.ForeColor.RGB = c
    s.Line.Visible = msoTrue
    s.Line.ForeColor.RGB = ShadeColor(c, 0.25)
    s.Line.Weight = 0.75
    s.Shadow.Visible = msoFalse
    StyleCaption s, LaneCaption(i, 0), sz, True, TextColorOn(c)
    s.OnAction = "Lane_Click"
End Sub

Private Sub AddHint(ByVal ws As Worksheet, ByVal nm As String, ByVal txt As String, ByVal x As Single, ByVal y As Single)
    Dim s As Shape
    Set s = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, 360, BTN_H)
    s.Name = nm
    s.Placement = xlFreeFloating
    s.Fill.Visible = msoFalse
    s.Line.Visible = msoFalse
    StyleCaption s, txt, 9, False, RGB(118, 118, 118)
    s.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
End Sub

Private Sub StyleCaption(ByVal s As Shape, ByVal caption As String, ByVal sz As Single, ByVal isBold As Boolean, _
                         ByVal txtColor As Long)
    Dim fnt As String
    fnt = CfgFont()
    With s.TextFrame2
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeNone
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 2
        .MarginRight = 2
        .MarginTop = 0
        .MarginBottom = 0
        .TextRange.Text = caption
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        With .TextRange.Font
            .Name = fnt
            .NameFarEast = fnt
            .Size = sz
            .Bold = IIf(isBold, msoTrue, msoFalse)
            .Fill.ForeColor.RGB = txtColor
        End With
    End With
End Sub

Private Sub DeleteShapes(ByVal ws As Worksheet, ByVal prefix As String)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(prefix)) = prefix Then ws.Shapes(i).Delete
    Next
End Sub

'==============================================================
'  サンプルの付箋
'==============================================================

' ステータスの番号は SetupConfigSheet の並び（1 起票済み … 5 レビュー対応中 … 8 連携済み、9 完了）
Private Sub CreateSamples()
    Dim d As NoteData
    Sample 1, 6, "（サンプル）付箋の使い方", _
           "右上の「" & StatusName(1) & " " & ChrW$(&H25B6) & "」をクリックすると、ステータスが進んで付箋の色が変わります。", _
           "サンプルの付箋は、選んで［捨てる］で削除できます。", 0, 0, 0
    Sample 2, 4, "（サンプル）ログイン画面の修正", _
           "パスワードを間違えたときのメッセージを分かりやすくする。", _
           "LoginForm.ValidatePassword()", Date - 2, Date + 3, 40
    Sample 3, 5, "（サンプル）請求書PDFのレイアウト変更", _
           "ロゴの位置と明細の列幅を調整する。", _
           "ReportService.ExportInvoice", Date - 4, Date - 1, 100
    Sample 4, 1, "（サンプル）在庫引当バッチの性能改善とリトライ処理の見直しに関するプルリクエストのレビュー依頼", _
           "タイトルが長くて欄に入りきらないときは、全文が見えるように文字が自動で小さくなります。", _
           "batch/StockAllocator.java", Date + 1, Date + 2, 0
    Sample 5, 2, "（サンプル）受注APIのコードレビュー", _
           "受注登録APIの入力チェックと例外処理をレビューする。", _
           "OrderController.Create()", Date - 1, Date + 1, 60
    Sample 8, 3, "（サンプル）API仕様の変更を共有", _
           "取引先APIの変更点を関係チームへ連携した。", _
           "docs/api/v2.md", Date - 3, Date - 2, 100

    d.Id = NextNoteId()
    d.Title = "（サンプル）環境構築の手順書"
    d.Body = "新しいメンバー向けに、開発環境の作り方をまとめる。"
    d.Memo = "README.md"
    d.StatusIdx = StatusCount()
    d.Created = Now - 9
    d.Done = Now - 1
    d.Updated = d.Done
    d.StartDate = Date - 8
    d.EndDate = Date - 2
    d.Progress = 100
    BuildNote shDone, d, OriginX(shDone), OriginY(shDone), True
    ArrangeDoneSheet False
End Sub

Private Sub Sample(ByVal si As Long, ByVal daysAgo As Long, ByVal t As String, ByVal b As String, ByVal m As String, _
                   ByVal sd As Date, ByVal ed As Date, ByVal pct As Long)
    Dim d As NoteData
    d.Id = NextNoteId()
    d.Title = t
    d.Body = b
    d.Memo = m
    d.StatusIdx = si
    d.Created = Now - daysAgo
    d.Updated = Now
    d.StartDate = sd
    d.EndDate = ed
    d.Progress = pct
    PlaceInLane shBoard, d, si
End Sub

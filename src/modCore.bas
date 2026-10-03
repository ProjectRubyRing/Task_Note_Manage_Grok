Option Explicit
Option Private Module

' タスク付箋ノート
' ノートシートの4列（起票済み・対応中・対応済み・連携済み）に付箋を置き、
' ステータスを変えると色と列が変わる。完了にするとグレーにして完了シートの先頭へ移す。
' 各付箋の実体は、タイトル / タスクの内容 / 備考 / ステータスのセル塊。
' 識別子は付箋左の細い帯に隠し、見た目の文字は Meiryo UI。

Public SuppressPrompts As Boolean
Public LastNotice As String
Public LastCreatedId As String

Private InHandler As Boolean

Private Const FONT_NAME As String = "Meiryo UI"
Private Const SHEET_BOARD As String = "ノート"
Private Const SHEET_DONE As String = "完了"
Private Const SHEET_HELP As String = "使い方"

Private Const ST_NEW As String = "起票済み"
Private Const ST_DOING As String = "対応中"
Private Const ST_HANDLED As String = "対応済み"
Private Const ST_LINKED As String = "連携済み"
Private Const ST_DONE As String = "完了"

Private Const PLACE_TITLE As String = "タイトルを入力"
Private Const PLACE_BODY As String = "タスクの内容を入力"
Private Const PLACE_REMARK As String = "関係する実装や連絡先などを入力"

Private Const FIRST_ROW As Long = 6
Private Const SLOT_HEIGHT As Long = 9
Private Const MAX_SLOTS As Long = 25
Private Const BOARD_LANES As Long = 4
Private Const DONE_LANES As Long = 3
Private Const HEADER_ROW As Long = 4

Private Const OFF_GAP As Long = 0
Private Const OFF_CAP_TITLE As Long = 1
Private Const OFF_TITLE As Long = 2
Private Const OFF_CAP_BODY As Long = 3
Private Const OFF_BODY As Long = 4
Private Const OFF_CAP_REMARK As Long = 5
Private Const OFF_REMARK As Long = 6
Private Const OFF_CAP_STATUS As Long = 7
Private Const OFF_STATUS As Long = 8

Private Type StickyNote
    Id As String
    Title As String
    Body As String
    Remark As String
    Status As String
    CreatedAt As String
    CompletedAt As String
    PrevStatus As String
End Type

Public Function LayoutToken() As String
    LayoutToken = "FIRST=6;SLOT=9;GAP=0;CAP_TITLE=1;TITLE=2;CAP_BODY=3;BODY=4;CAP_REMARK=5;REMARK=6;CAP_STATUS=7;STATUS=8;STRIP0=2;BODY0=3;STEP=3;MAX=25;BLANES=4;DLANES=3;FONT=Meiryo UI"
End Function

Public Sub ResetRuntimeState()
    InHandler = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub

Public Sub BuildAll()
    Dim prevSilent As Boolean
    prevSilent = SuppressPrompts
    On Error GoTo Fail
    SuppressPrompts = True
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ThisWorkbook.Styles("Normal").Font.Name = FONT_NAME
    CreateAllSheets
    DeleteStraySheets
    SetupBoardChrome
    SetupDoneChrome
    SetupHelp
    ApplyBaseView SheetBoard(), "A1:M230", 100
    ApplyBaseView SheetDone(), "A1:J230", 100
    ApplyBaseView SheetHelp(), "A1:E40", 100
    PrepareAllSlotHeights SheetBoard()
    PrepareAllSlotHeights SheetDone()
    ProtectAll

    Application.EnableEvents = True
    SeedSamples
    Initialize

    SheetBoard().Activate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    SuppressPrompts = prevSilent
    Exit Sub
Fail:
    ResetRuntimeState
    Application.DisplayAlerts = True
    SuppressPrompts = prevSilent
    Err.Raise Err.Number, "BuildAll", Err.Description
End Sub

Public Sub Initialize()
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    UnprotectAll
    ThisWorkbook.Styles("Normal").Font.Name = FONT_NAME
    EnsureCounter
    ApplyColumnWidths SheetBoard(), BOARD_LANES
    ApplyColumnWidths SheetDone(), DONE_LANES
    ReformatAll SheetBoard(), BOARD_LANES
    ReformatAll SheetDone(), DONE_LANES
    RefreshRowBands SheetBoard(), BOARD_LANES
    RefreshRowBands SheetDone(), DONE_LANES
    UpdateBoardHeaders
    UpdateDoneTitle
    EnsureButtons
    Application.EnableEvents = True
    ReconcileBoard
    Application.EnableEvents = False
    UpdateBoardHeaders
    UpdateDoneTitle
    RefreshRowBands SheetBoard(), BOARD_LANES
    RefreshRowBands SheetDone(), DONE_LANES
    ProtectAll
    SheetBoard().ScrollArea = "A1:M230"
    SheetDone().ScrollArea = "A1:J230"
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
Fail:
    Dim initErr As Long
    Dim initDesc As String
    initErr = Err.Number
    initDesc = Err.Description
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    LastNotice = "Initialize: " & initDesc
    If Not SuppressPrompts Then
        MsgBox "初期化に失敗しました。" & vbCrLf & initDesc, vbExclamation, "タスク付箋ノート"
    End If
    Err.Raise initErr, "Initialize", initDesc
End Sub

Public Sub AddSticky()
    If InHandler Then Exit Sub
    InHandler = True
    On Error GoTo Fail
    LastNotice = ""
    EnsureWritable SheetBoard()
    If Len(CreateNote(PLACE_TITLE, PLACE_BODY, PLACE_REMARK)) = 0 Then
        InHandler = False
        Exit Sub
    End If
    InHandler = False
    Exit Sub
Fail:
    InHandler = False
    ResetRuntimeState
    LastNotice = "AddSticky: " & Err.Description
    Notify LastNotice
End Sub

Public Sub DeleteSticky()
    If InHandler Then Exit Sub
    InHandler = True
    On Error GoTo Fail
    LastNotice = ""
    EnsureWritable SheetBoard()
    If ActiveSheet Is Nothing Then
        Notify "ノート上の付箋を選んでから削除してください。"
        GoTo Done
    End If
    If ActiveSheet.Name <> SHEET_BOARD Then
        Notify "ノート上の付箋を選んでから削除してください。完了した付箋は、完了シートの「ノートに戻す」を使います。"
        GoTo Done
    End If
    Dim lane As Long
    Dim slot As Long
    If Not ActiveOccupiedNote(SheetBoard(), BOARD_LANES, lane, slot) Then
        Notify "削除する付箋を、ノート上でクリックしてからもう一度押してください。"
        GoTo Done
    End If
    If Not ConfirmDelete() Then GoTo Done
    Application.EnableEvents = False
    RemoveBoardSlot lane, slot
    Application.EnableEvents = True
Done:
    InHandler = False
    Exit Sub
Fail:
    InHandler = False
    ResetRuntimeState
    LastNotice = "DeleteSticky: " & Err.Description
    Notify LastNotice
End Sub

Public Sub RestoreSticky()
    If InHandler Then Exit Sub
    InHandler = True
    On Error GoTo Fail
    LastNotice = ""
    EnsureWritable SheetDone()
    EnsureWritable SheetBoard()
    If ActiveSheet Is Nothing Then
        Notify "完了シートで、戻す付箋を選んでください。"
        GoTo Done
    End If
    If ActiveSheet.Name <> SHEET_DONE Then
        Notify "完了シートで、戻す付箋を選んでください。"
        GoTo Done
    End If
    Dim lane As Long
    Dim slot As Long
    If Not ActiveOccupiedNote(SheetDone(), DONE_LANES, lane, slot) Then
        Notify "戻す付箋を、完了シート上でクリックしてからもう一度押してください。"
        GoTo Done
    End If
    Dim note As StickyNote
    note = ReadSlot(SheetDone(), lane, slot)
    If Len(note.Id) = 0 Then
        Notify "戻す付箋を選べませんでした。"
        GoTo Done
    End If
    RestoreNote note
Done:
    InHandler = False
    Exit Sub
Fail:
    InHandler = False
    ResetRuntimeState
    LastNotice = "RestoreSticky: " & Err.Description
    Notify LastNotice
End Sub

Public Sub DeleteAt(ByVal sheetName As String, ByVal lane As Long, ByVal slot As Long)
    InHandler = False
    On Error GoTo Fail
    Application.EnableEvents = False
    If sheetName = SHEET_BOARD Then
        If SlotOccupied(SheetBoard(), lane, slot) Then RemoveBoardSlot lane, slot
    ElseIf sheetName = SHEET_DONE Then
        Dim idx As Long
        idx = DoneIndexOf(lane, slot)
        If idx >= 0 Then RemoveDoneAt idx
    End If
    Application.EnableEvents = True
    Exit Sub
Fail:
    ResetRuntimeState
    LastNotice = "DeleteAt: " & Err.Description
    Err.Raise Err.Number, "DeleteAt", Err.Description
End Sub

Public Sub HandleSheetChange(ByVal Sh As Object, ByVal Target As Range)
    If InHandler Then Exit Sub
    If Sh Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub
    If TypeName(Sh) <> "Worksheet" Then Exit Sub

    Dim ws As Worksheet
    Set ws = Sh
    If ws.Name <> SHEET_BOARD Then Exit Sub

    Dim zone As Range
    Set zone = Intersect(Target, BoardInputZone(ws))
    If zone Is Nothing Then Exit Sub

    InHandler = True
    On Error GoTo Fail
    LastNotice = ""

    Dim c As Range
    Dim statusCell As Range
    For Each c In zone.Cells
        If IsStatusCell(ws, c) Then
            Set statusCell = c
        ElseIf IsTextCell(ws, c) Then
            ApplyTypingFont ws, c
        End If
    Next c

    If Not statusCell Is Nothing Then
        ProcessStatusChange ws, statusCell
    End If

    InHandler = False
    Exit Sub
Fail:
    InHandler = False
    ResetRuntimeState
    LastNotice = "HandleSheetChange: " & Err.Description
    Notify LastNotice
End Sub

Private Sub ProcessStatusChange(ByVal ws As Worksheet, ByVal statusCell As Range)
    Dim lane As Long
    Dim slot As Long
    If Not LocateCell(ws, statusCell.Row, statusCell.Column, lane, slot) Then Exit Sub
    If lane < 0 Or lane >= BOARD_LANES Then Exit Sub
    If Not SlotOccupied(ws, lane, slot) Then Exit Sub

    Dim newStatus As String
    newStatus = Trim$(CStr(statusCell.Value))
    Dim oldStatus As String
    oldStatus = StatusOfLane(lane)
    If Len(newStatus) = 0 Then newStatus = oldStatus

    Dim note As StickyNote
    note = ReadSlot(ws, lane, slot)
    If newStatus = oldStatus Then
        note.Status = oldStatus
        Application.EnableEvents = False
        WriteSlot ws, lane, slot, note
        Application.EnableEvents = True
        Exit Sub
    End If

    If Not IsKnownStatus(newStatus) Then
        Application.EnableEvents = False
        AssignText statusCell, oldStatus
        ApplyStatusValidation statusCell
        Application.EnableEvents = True
        Notify "ステータスは「起票済み」「対応中」「対応済み」「連携済み」「完了」から選んでください。"
        Exit Sub
    End If

    note.PrevStatus = oldStatus
    note.Status = newStatus

    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    If newStatus = ST_DONE Then
        note.CompletedAt = Format$(Now, "yyyy-mm-dd hh:nn")
        If Not InsertDoneAtTop(note) Then
            note.Status = oldStatus
            note.CompletedAt = ""
            note.PrevStatus = ""
            WriteSlot ws, lane, slot, note
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            Application.Cursor = xlDefault
            Notify "完了シートがいっぱいです。付箋は元の列に残しています。"
            Exit Sub
        End If
        ClearSlot ws, lane, slot
        CompactLane ws, lane
    Else
        Dim destLane As Long
        Dim destSlot As Long
        destLane = LaneOfStatus(newStatus)
        destSlot = FirstEmpty(ws, destLane)
        If destSlot < 0 Then
            note.Status = oldStatus
            note.PrevStatus = ""
            WriteSlot ws, lane, slot, note
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            Application.Cursor = xlDefault
            Notify "「" & newStatus & "」の列がいっぱいです。付箋は元の列に残しています。"
            Exit Sub
        End If
        note.CompletedAt = ""
        WriteSlot ws, destLane, destSlot, note
        If Not (destLane = lane And destSlot = slot) Then
            ClearSlot ws, lane, slot
        End If
        CompactLane ws, lane
    End If

    UpdateBoardHeaders
    UpdateDoneTitle
    RefreshRowBands ws, BOARD_LANES
    RefreshRowBands SheetDone(), DONE_LANES
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    RevealNote note.Id
End Sub

Private Sub RestoreNote(ByRef note As StickyNote)
    Dim back As String
    back = note.PrevStatus
    If LaneOfStatus(back) < 0 Then back = ST_NEW
    Dim destLane As Long
    Dim destSlot As Long
    destLane = LaneOfStatus(back)
    destSlot = FirstEmpty(SheetBoard(), destLane)
    If destSlot < 0 Then
        Notify "戻し先の「" & back & "」がいっぱいです。完了シートに残しています。"
        Exit Sub
    End If
    Dim idx As Long
    idx = DoneIndexOfId(note.Id)
    If idx < 0 Then
        Notify "完了シート上の位置を特定できませんでした。"
        Exit Sub
    End If

    Dim restored As StickyNote
    restored = note
    restored.Status = back
    restored.CompletedAt = ""
    restored.PrevStatus = ""

    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    WriteSlot SheetBoard(), destLane, destSlot, restored
    RemoveDoneAt idx
    UpdateBoardHeaders
    UpdateDoneTitle
    RefreshRowBands SheetBoard(), BOARD_LANES
    RefreshRowBands SheetDone(), DONE_LANES
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    RevealNote restored.Id
End Sub

Private Function CreateNote(ByVal title As String, ByVal body As String, ByVal remark As String) As String
    Dim slot As Long
    slot = FirstEmpty(SheetBoard(), 0)
    If slot < 0 Then
        Notify "起票済みの列がいっぱいです。先にほかのステータスへ進めるか、削除してください。"
        CreateNote = ""
        Exit Function
    End If

    Dim note As StickyNote
    note.Id = NextId()
    If Len(title) = 0 Then
        note.Title = PLACE_TITLE
    Else
        note.Title = title
    End If
    If Len(body) = 0 Then
        note.Body = PLACE_BODY
    Else
        note.Body = body
    End If
    If Len(remark) = 0 Then
        note.Remark = PLACE_REMARK
    Else
        note.Remark = remark
    End If
    note.Status = ST_NEW
    note.CreatedAt = Format$(Now, "yyyy-mm-dd hh:nn")
    note.CompletedAt = ""
    note.PrevStatus = ""

    Dim eventsWere As Boolean
    eventsWere = Application.EnableEvents
    Application.EnableEvents = False
    WriteSlot SheetBoard(), 0, slot, note
    UpdateBoardHeaders
    RefreshRowBands SheetBoard(), BOARD_LANES
    Application.EnableEvents = eventsWere

    LastCreatedId = note.Id
    CreateNote = note.Id
    If Not SuppressPrompts Then
        Application.Goto SheetBoard().Cells(SlotTop(slot) + OFF_TITLE, BodyCol(0)), True
    End If
End Function

Private Sub SeedSamples()
    If BoardNoteCount() > 0 Then
        Exit Sub
    End If
    If DoneNoteCount() > 0 Then
        Exit Sub
    End If

    CreateNote "週次レポートの自動化", _
        "毎週月曜の手作業集計をなくしたい。", _
        "実装候補: Power Query。集計ブック SalesWeekly.xlsx"

    Dim id As String
    id = CreateNote("ログイン失敗の調査", _
        "ステージング環境で 401 が断続的に返る。再現手順を整理する。", _
        "AuthFilter 周辺。チケット #441")
    SetStatusValue id, ST_DOING

    id = CreateNote("日付の入力チェック", _
        "未入力のとき保存できないようチェックを入れた。", _
        "FormValidator の日付判定")
    SetStatusValue id, ST_HANDLED

    id = CreateNote("APIエラーコードの共有", _
        "連携先へエラーコード一覧を送付し、確認待ち。", _
        "docs/error-codes.md を更新して共有済み")
    SetStatusValue id, ST_LINKED

    id = CreateNote("開発環境の初期構築", _
        "ローカルで起動する手順を整備した。", _
        "README のセットアップ手順")
    SetStatusValue id, ST_DONE
End Sub

Private Sub SetStatusValue(ByVal taskId As String, ByVal newStatus As String)
    Dim lane As Long
    Dim slot As Long
    If Not FindIdOnSheet(SheetBoard(), taskId, lane, slot) Then
        LastNotice = "付箋が見つかりません。"
        Exit Sub
    End If
    Dim cell As Range
    Set cell = SheetBoard().Cells(SlotTop(slot) + OFF_STATUS, BodyCol(lane))
    On Error Resume Next
    cell.Validation.ShowError = False
    On Error GoTo 0
    Application.EnableEvents = True
    cell.Value = newStatus
End Sub

Private Sub ReconcileBoard()
    Dim ws As Worksheet
    Set ws = SheetBoard()
    Dim lane As Long
    Dim s As Long
    Dim ids() As String
    Dim n As Long
    n = 0
    For lane = 0 To BOARD_LANES - 1
        For s = 0 To MAX_SLOTS - 1
            If SlotOccupied(ws, lane, s) Then
                Dim cellStatus As String
                cellStatus = Trim$(CStr(ws.Cells(SlotTop(s) + OFF_STATUS, BodyCol(lane)).Value))
                If cellStatus <> StatusOfLane(lane) And Len(cellStatus) > 0 Then
                    ReDim Preserve ids(0 To n)
                    ids(n) = CStr(ws.Cells(SlotTop(s) + OFF_TITLE, StripCol(lane)).Value)
                    n = n + 1
                End If
            End If
        Next s
    Next lane
    Dim i As Long
    For i = 0 To n - 1
        Dim foundLane As Long
        Dim foundSlot As Long
        If FindIdOnSheet(ws, ids(i), foundLane, foundSlot) Then
            ProcessStatusChange ws, ws.Cells(SlotTop(foundSlot) + OFF_STATUS, BodyCol(foundLane))
        End If
    Next i
End Sub

Private Sub RemoveBoardSlot(ByVal lane As Long, ByVal slot As Long)
    ClearSlot SheetBoard(), lane, slot
    CompactLane SheetBoard(), lane
    UpdateBoardHeaders
    RefreshRowBands SheetBoard(), BOARD_LANES
End Sub

Private Sub CompactLane(ByVal ws As Worksheet, ByVal lane As Long)
    Dim items() As StickyNote
    Dim n As Long
    Dim s As Long
    n = 0
    For s = 0 To MAX_SLOTS - 1
        If SlotOccupied(ws, lane, s) Then
            ReDim Preserve items(0 To n)
            items(n) = ReadSlot(ws, lane, s)
            n = n + 1
        End If
    Next s
    For s = 0 To MAX_SLOTS - 1
        If s < n Then
            WriteSlot ws, lane, s, items(s)
        Else
            ClearSlot ws, lane, s
        End If
    Next s
End Sub

Private Function InsertDoneAtTop(ByRef note As StickyNote) As Boolean
    Dim count As Long
    count = DoneNoteCount()
    If count >= DONE_LANES * MAX_SLOTS Then
        InsertDoneAtTop = False
        Exit Function
    End If
    Dim items() As StickyNote
    ReDim items(0 To count)
    items(0) = note
    Dim i As Long
    For i = 0 To count - 1
        items(i + 1) = ReadLinear(i)
    Next i
    For i = 0 To count
        WriteLinear i, items(i)
    Next i
    InsertDoneAtTop = True
End Function

Private Sub RemoveDoneAt(ByVal index As Long)
    Dim count As Long
    count = DoneNoteCount()
    If index < 0 Or index >= count Then Exit Sub
    Dim keep() As StickyNote
    Dim n As Long
    Dim i As Long
    n = 0
    For i = 0 To count - 1
        If i <> index Then
            ReDim Preserve keep(0 To n)
            keep(n) = ReadLinear(i)
            n = n + 1
        End If
    Next i
    For i = 0 To count - 1
        If i < n Then
            WriteLinear i, keep(i)
        Else
            Dim lane As Long
            Dim slot As Long
            LinearToPos i, lane, slot
            ClearSlot SheetDone(), lane, slot
        End If
    Next i
    UpdateDoneTitle
    RefreshRowBands SheetDone(), DONE_LANES
End Sub

Private Function ReadLinear(ByVal index As Long) As StickyNote
    Dim lane As Long
    Dim slot As Long
    LinearToPos index, lane, slot
    ReadLinear = ReadSlot(SheetDone(), lane, slot)
End Function

Private Sub WriteLinear(ByVal index As Long, ByRef note As StickyNote)
    Dim lane As Long
    Dim slot As Long
    LinearToPos index, lane, slot
    WriteSlot SheetDone(), lane, slot, note
End Sub

Private Sub LinearToPos(ByVal index As Long, ByRef lane As Long, ByRef slot As Long)
    lane = index \ MAX_SLOTS
    slot = index Mod MAX_SLOTS
End Sub

Private Function ReadSlot(ByVal ws As Worksheet, ByVal lane As Long, ByVal slot As Long) As StickyNote
    Dim top As Long
    Dim stripC As Long
    Dim bodyC As Long
    Dim note As StickyNote
    top = SlotTop(slot)
    stripC = StripCol(lane)
    bodyC = BodyCol(lane)
    note.Id = CellText(ws.Cells(top + OFF_TITLE, stripC))
    note.Title = CellText(ws.Cells(top + OFF_TITLE, bodyC))
    note.Body = CellText(ws.Cells(top + OFF_BODY, bodyC))
    note.Remark = CellText(ws.Cells(top + OFF_REMARK, bodyC))
    note.CreatedAt = CellText(ws.Cells(top + OFF_BODY, stripC))
    note.PrevStatus = CellText(ws.Cells(top + OFF_REMARK, stripC))
    note.CompletedAt = CellText(ws.Cells(top + OFF_STATUS, stripC))
    If ws.Name = SHEET_DONE Then
        note.Status = ST_DONE
    Else
        note.Status = Trim$(CellText(ws.Cells(top + OFF_STATUS, bodyC)))
    End If
    ReadSlot = note
End Function

Private Sub WriteSlot(ByVal ws As Worksheet, ByVal lane As Long, ByVal slot As Long, ByRef note As StickyNote)
    FormatSlot ws, lane, slot, note
End Sub

Private Sub FormatSlot(ByVal ws As Worksheet, ByVal lane As Long, ByVal slot As Long, ByRef note As StickyNote)
    Dim top As Long
    Dim stripC As Long
    Dim bodyC As Long
    top = SlotTop(slot)
    stripC = StripCol(lane)
    bodyC = BodyCol(lane)

    Dim stripColor As Long
    Dim titleColor As Long
    Dim bodyColor As Long
    Dim captionColor As Long
    Dim borderColor As Long
    Dim fontColor As Long
    Palette note.Status, stripColor, titleColor, bodyColor, captionColor, borderColor, fontColor

    ClearGap ws, top, stripC, bodyC

    Dim r As Long
    For r = OFF_CAP_TITLE To OFF_STATUS
        StyleMetaStrip ws.Cells(top + r, stripC), stripColor
    Next r

    StyleCaption ws.Cells(top + OFF_CAP_TITLE, bodyC), "タイトル", titleColor, captionColor
    StyleCaption ws.Cells(top + OFF_CAP_BODY, bodyC), "タスクの内容", bodyColor, captionColor
    StyleCaption ws.Cells(top + OFF_CAP_REMARK, bodyC), "備考（関係する実装など）", bodyColor, captionColor
    StyleCaption ws.Cells(top + OFF_CAP_STATUS, bodyC), "ステータス", titleColor, captionColor

    StyleField ws.Cells(top + OFF_TITLE, bodyC), note.Title, titleColor, fontColor, 14, True, xlCenter
    StyleField ws.Cells(top + OFF_BODY, bodyC), note.Body, bodyColor, fontColor, 11, False, xlTop
    StyleField ws.Cells(top + OFF_REMARK, bodyC), note.Remark, bodyColor, fontColor, 10, False, xlTop

    Dim statusText As String
    If note.Status = ST_DONE And Len(note.CompletedAt) > 0 Then
        statusText = ST_DONE & "  " & note.CompletedAt
    Else
        statusText = note.Status
    End If
    Dim statusCell As Range
    Set statusCell = ws.Cells(top + OFF_STATUS, bodyC)
    StyleField statusCell, statusText, titleColor, fontColor, 12, True, xlCenter
    statusCell.HorizontalAlignment = xlCenter

    If ws.Name <> SHEET_DONE Then
        TintIfPlaceholder ws.Cells(top + OFF_TITLE, bodyC)
        TintIfPlaceholder ws.Cells(top + OFF_BODY, bodyC)
        TintIfPlaceholder ws.Cells(top + OFF_REMARK, bodyC)
    End If

    PutMeta ws.Cells(top + OFF_TITLE, stripC), note.Id, stripColor
    PutMeta ws.Cells(top + OFF_BODY, stripC), note.CreatedAt, stripColor
    PutMeta ws.Cells(top + OFF_REMARK, stripC), note.PrevStatus, stripColor
    PutMeta ws.Cells(top + OFF_STATUS, stripC), note.CompletedAt, stripColor

    If ws.Name = SHEET_DONE Then
        LockBodyColumn ws, top, bodyC, True
        ClearValidation statusCell
    Else
        ws.Cells(top + OFF_CAP_TITLE, bodyC).Locked = True
        ws.Cells(top + OFF_CAP_BODY, bodyC).Locked = True
        ws.Cells(top + OFF_CAP_REMARK, bodyC).Locked = True
        ws.Cells(top + OFF_CAP_STATUS, bodyC).Locked = True
        ws.Cells(top + OFF_TITLE, bodyC).Locked = False
        ws.Cells(top + OFF_BODY, bodyC).Locked = False
        ws.Cells(top + OFF_REMARK, bodyC).Locked = False
        statusCell.Locked = False
        ApplyStatusValidation statusCell
    End If

    Dim block As Range
    Set block = ws.Range(ws.Cells(top + OFF_CAP_TITLE, stripC), ws.Cells(top + OFF_STATUS, bodyC))
    block.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, Color:=borderColor
    With ws.Range(ws.Cells(top + OFF_TITLE, stripC), ws.Cells(top + OFF_TITLE, bodyC)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = borderColor
    End With
    With ws.Range(ws.Cells(top + OFF_REMARK, stripC), ws.Cells(top + OFF_REMARK, bodyC)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = borderColor
    End With
End Sub

Private Sub ClearSlot(ByVal ws As Worksheet, ByVal lane As Long, ByVal slot As Long)
    Dim top As Long
    Dim stripC As Long
    Dim bodyC As Long
    Dim paper As Long
    top = SlotTop(slot)
    stripC = StripCol(lane)
    bodyC = BodyCol(lane)
    paper = PaperColor(ws)
    Dim r As Long
    For r = 0 To SLOT_HEIGHT - 1
        WipeCell ws.Cells(top + r, stripC), paper
        WipeCell ws.Cells(top + r, bodyC), paper
    Next r
End Sub

Private Sub WipeCell(ByVal cell As Range, ByVal paper As Long)
    On Error Resume Next
    cell.Validation.Delete
    On Error GoTo 0
    cell.Clear
    cell.Interior.Pattern = xlSolid
    cell.Interior.Color = paper
    cell.Font.Name = FONT_NAME
    cell.Font.Color = RGB(45, 40, 35)
    cell.Locked = True
    cell.WrapText = False
    cell.Borders.LineStyle = xlNone
End Sub

Private Sub ClearGap(ByVal ws As Worksheet, ByVal top As Long, ByVal stripC As Long, ByVal bodyC As Long)
    Dim paper As Long
    paper = PaperColor(ws)
    WipeCell ws.Cells(top + OFF_GAP, stripC), paper
    WipeCell ws.Cells(top + OFF_GAP, bodyC), paper
End Sub

Private Sub StyleMetaStrip(ByVal cell As Range, ByVal stripColor As Long)
    cell.Interior.Pattern = xlSolid
    cell.Interior.Color = stripColor
    cell.Font.Name = FONT_NAME
    cell.Font.Size = 1
    cell.Font.Color = stripColor
    cell.Locked = True
    cell.HorizontalAlignment = xlCenter
    cell.VerticalAlignment = xlCenter
End Sub

Private Sub PutMeta(ByVal cell As Range, ByVal text As String, ByVal stripColor As Long)
    cell.NumberFormat = "@"
    cell.Value = text
    cell.NumberFormat = ";;;"
    cell.Font.Color = stripColor
    cell.Font.Size = 1
    cell.Font.Name = FONT_NAME
    cell.Locked = True
End Sub

Private Sub StyleCaption(ByVal cell As Range, ByVal text As String, ByVal fill As Long, ByVal fontColor As Long)
    cell.NumberFormat = "@"
    cell.Value = text
    cell.Interior.Pattern = xlSolid
    cell.Interior.Color = fill
    cell.Font.Name = FONT_NAME
    cell.Font.Size = 8
    cell.Font.Bold = True
    cell.Font.Italic = False
    cell.Font.Color = fontColor
    cell.HorizontalAlignment = xlLeft
    cell.VerticalAlignment = xlCenter
    cell.WrapText = False
    cell.IndentLevel = 1
    cell.Locked = True
    cell.ShrinkToFit = False
End Sub

Private Sub StyleField(ByVal cell As Range, ByVal text As String, ByVal fill As Long, ByVal fontColor As Long, ByVal size As Long, ByVal isBold As Boolean, ByVal vAlign As Long)
    cell.NumberFormat = "@"
    cell.Value = text
    cell.Interior.Pattern = xlSolid
    cell.Interior.Color = fill
    cell.Font.Name = FONT_NAME
    cell.Font.Size = size
    cell.Font.Bold = isBold
    cell.Font.Italic = False
    cell.Font.Color = fontColor
    cell.HorizontalAlignment = xlLeft
    cell.VerticalAlignment = vAlign
    cell.WrapText = True
    cell.IndentLevel = 1
    cell.ShrinkToFit = False
End Sub

Private Sub TintIfPlaceholder(ByVal cell As Range)
    If IsPlaceholder(CellText(cell)) Then
        cell.Font.Color = RGB(145, 134, 110)
        cell.Font.Italic = False
    End If
End Sub

Private Sub ApplyTypingFont(ByVal ws As Worksheet, ByVal cell As Range)
    Dim offset As Long
    offset = (cell.Row - FIRST_ROW) Mod SLOT_HEIGHT
    cell.Font.Name = FONT_NAME
    cell.Font.Italic = False
    If IsPlaceholder(CellText(cell)) Then
        cell.Font.Color = RGB(145, 134, 110)
    Else
        cell.Font.Color = RGB(45, 40, 35)
    End If
    Select Case offset
        Case OFF_TITLE
            cell.Font.Size = 14
            cell.Font.Bold = True
        Case OFF_BODY
            cell.Font.Size = 11
            cell.Font.Bold = False
        Case OFF_REMARK
            cell.Font.Size = 10
            cell.Font.Bold = False
    End Select
    cell.WrapText = True
End Sub

Private Sub LockBodyColumn(ByVal ws As Worksheet, ByVal top As Long, ByVal bodyC As Long, ByVal locked As Boolean)
    Dim r As Long
    For r = OFF_CAP_TITLE To OFF_STATUS
        ws.Cells(top + r, bodyC).Locked = locked
    Next r
End Sub

Private Sub ApplyStatusValidation(ByVal cell As Range)
    Dim ws As Worksheet
    Set ws = cell.Worksheet
    Dim eventsWere As Boolean
    Dim prevAlerts As Boolean
    Dim restoreProtect As Boolean
    eventsWere = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    restoreProtect = ws.ProtectContents
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    If restoreProtect Then UnprotectSheet ws
    On Error Resume Next
    cell.Validation.Delete
    On Error GoTo 0
    Dim sep As String
    sep = Application.International(xlListSeparator)
    cell.Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Formula1:=ST_NEW & sep & ST_DOING & sep & ST_HANDLED & sep & ST_LINKED & sep & ST_DONE
    cell.Validation.IgnoreBlank = False
    cell.Validation.InCellDropdown = True
    cell.Validation.ShowError = True
    cell.Validation.ErrorTitle = "ステータス"
    cell.Validation.ErrorMessage = "起票済み、対応中、対応済み、連携済み、完了 から選んでください。"
    If restoreProtect Then ApplyProtection ws
    Application.DisplayAlerts = prevAlerts
    Application.EnableEvents = eventsWere
End Sub

Private Sub ClearValidation(ByVal cell As Range)
    On Error Resume Next
    cell.Validation.Delete
    On Error GoTo 0
End Sub

Private Sub Palette(ByVal status As String, ByRef stripColor As Long, ByRef titleColor As Long, ByRef bodyColor As Long, ByRef captionColor As Long, ByRef borderColor As Long, ByRef fontColor As Long)
    Select Case status
        Case ST_DOING
            stripColor = RGB(214, 106, 28)
            titleColor = RGB(255, 166, 86)
            bodyColor = RGB(255, 204, 150)
            captionColor = RGB(140, 70, 20)
            borderColor = RGB(196, 100, 30)
            fontColor = RGB(45, 40, 35)
        Case ST_HANDLED
            stripColor = RGB(56, 146, 72)
            titleColor = RGB(138, 206, 132)
            bodyColor = RGB(186, 228, 176)
            captionColor = RGB(36, 100, 48)
            borderColor = RGB(60, 140, 70)
            fontColor = RGB(45, 40, 35)
        Case ST_LINKED
            stripColor = RGB(36, 116, 196)
            titleColor = RGB(126, 184, 240)
            bodyColor = RGB(176, 214, 250)
            captionColor = RGB(24, 78, 140)
            borderColor = RGB(40, 110, 180)
            fontColor = RGB(45, 40, 35)
        Case ST_DONE
            stripColor = RGB(128, 128, 128)
            titleColor = RGB(186, 186, 186)
            bodyColor = RGB(210, 210, 210)
            captionColor = RGB(90, 90, 90)
            borderColor = RGB(120, 120, 120)
            fontColor = RGB(80, 80, 80)
        Case Else
            stripColor = RGB(214, 158, 18)
            titleColor = RGB(255, 214, 74)
            bodyColor = RGB(255, 236, 130)
            captionColor = RGB(128, 96, 20)
            borderColor = RGB(196, 146, 28)
            fontColor = RGB(45, 40, 35)
    End Select
End Sub

Private Function PaperColor(ByVal ws As Worksheet) As Long
    If Not ws Is Nothing Then
        If ws.Name = SHEET_DONE Then
            PaperColor = RGB(232, 232, 230)
            Exit Function
        End If
    End If
    PaperColor = RGB(243, 237, 224)
End Function

Private Sub AssignText(ByVal cell As Range, ByVal text As String)
    cell.NumberFormat = "@"
    cell.Value = text
End Sub

Private Function CellText(ByVal cell As Range) As String
    If IsEmpty(cell.Value) Then
        CellText = ""
    Else
        CellText = CStr(cell.Value)
    End If
End Function

Private Function IsPlaceholder(ByVal text As String) As Boolean
    IsPlaceholder = (text = PLACE_TITLE Or text = PLACE_BODY Or text = PLACE_REMARK)
End Function

Private Function IsKnownStatus(ByVal text As String) As Boolean
    IsKnownStatus = (text = ST_NEW Or text = ST_DOING Or text = ST_HANDLED Or text = ST_LINKED Or text = ST_DONE)
End Function

Private Function StatusOfLane(ByVal lane As Long) As String
    Select Case lane
        Case 0: StatusOfLane = ST_NEW
        Case 1: StatusOfLane = ST_DOING
        Case 2: StatusOfLane = ST_HANDLED
        Case 3: StatusOfLane = ST_LINKED
        Case Else: StatusOfLane = ""
    End Select
End Function

Private Function LaneOfStatus(ByVal status As String) As Long
    Select Case status
        Case ST_NEW: LaneOfStatus = 0
        Case ST_DOING: LaneOfStatus = 1
        Case ST_HANDLED: LaneOfStatus = 2
        Case ST_LINKED: LaneOfStatus = 3
        Case Else: LaneOfStatus = -1
    End Select
End Function

Private Function StripCol(ByVal lane As Long) As Long
    StripCol = 2 + lane * 3
End Function

Private Function BodyCol(ByVal lane As Long) As Long
    BodyCol = 3 + lane * 3
End Function

Private Function SlotTop(ByVal slot As Long) As Long
    SlotTop = FIRST_ROW + slot * SLOT_HEIGHT
End Function

Private Function SlotOccupied(ByVal ws As Worksheet, ByVal lane As Long, ByVal slot As Long) As Boolean
    SlotOccupied = Len(CellText(ws.Cells(SlotTop(slot) + OFF_TITLE, StripCol(lane)))) > 0
End Function

Private Function FirstEmpty(ByVal ws As Worksheet, ByVal lane As Long) As Long
    Dim s As Long
    For s = 0 To MAX_SLOTS - 1
        If Not SlotOccupied(ws, lane, s) Then
            FirstEmpty = s
            Exit Function
        End If
    Next s
    FirstEmpty = -1
End Function

Private Function LocateCell(ByVal ws As Worksheet, ByVal row As Long, ByVal col As Long, ByRef lane As Long, ByRef slot As Long) As Boolean
    lane = -1
    slot = -1
    If row < FIRST_ROW Then Exit Function
    Dim i As Long
    For i = 0 To 3
        If col = StripCol(i) Or col = BodyCol(i) Then
            lane = i
            Exit For
        End If
    Next i
    If lane < 0 Then Exit Function
    slot = (row - FIRST_ROW) \ SLOT_HEIGHT
    If slot < 0 Or slot >= MAX_SLOTS Then
        lane = -1
        slot = -1
        Exit Function
    End If
    LocateCell = True
End Function

Private Function IsStatusCell(ByVal ws As Worksheet, ByVal cell As Range) As Boolean
    Dim lane As Long
    Dim slot As Long
    If Not LocateCell(ws, cell.Row, cell.Column, lane, slot) Then Exit Function
    If lane >= BOARD_LANES Then Exit Function
    If cell.Column <> BodyCol(lane) Then Exit Function
    If ((cell.Row - FIRST_ROW) Mod SLOT_HEIGHT) <> OFF_STATUS Then Exit Function
    IsStatusCell = SlotOccupied(ws, lane, slot)
End Function

Private Function IsTextCell(ByVal ws As Worksheet, ByVal cell As Range) As Boolean
    Dim lane As Long
    Dim slot As Long
    If Not LocateCell(ws, cell.Row, cell.Column, lane, slot) Then Exit Function
    If lane >= BOARD_LANES Then Exit Function
    If cell.Column <> BodyCol(lane) Then Exit Function
    Dim offset As Long
    offset = (cell.Row - FIRST_ROW) Mod SLOT_HEIGHT
    If offset <> OFF_TITLE And offset <> OFF_BODY And offset <> OFF_REMARK Then Exit Function
    IsTextCell = SlotOccupied(ws, lane, slot)
End Function

Private Function BoardInputZone(ByVal ws As Worksheet) As Range
    Set BoardInputZone = ws.Range( _
        ws.Cells(FIRST_ROW, StripCol(0)), _
        ws.Cells(SlotTop(MAX_SLOTS) - 1, BodyCol(BOARD_LANES - 1)))
End Function

Private Function FindIdOnSheet(ByVal ws As Worksheet, ByVal taskId As String, ByRef laneOut As Long, ByRef slotOut As Long) As Boolean
    Dim maxLane As Long
    If ws.Name = SHEET_DONE Then
        maxLane = DONE_LANES - 1
    Else
        maxLane = BOARD_LANES - 1
    End If
    Dim lane As Long
    Dim s As Long
    For lane = 0 To maxLane
        For s = 0 To MAX_SLOTS - 1
            If CellText(ws.Cells(SlotTop(s) + OFF_TITLE, StripCol(lane))) = taskId Then
                laneOut = lane
                slotOut = s
                FindIdOnSheet = True
                Exit Function
            End If
        Next s
    Next lane
End Function

Private Function DoneIndexOfId(ByVal taskId As String) As Long
    Dim i As Long
    Dim lane As Long
    Dim slot As Long
    Dim count As Long
    count = DoneNoteCount()
    For i = 0 To count - 1
        LinearToPos i, lane, slot
        If CellText(SheetDone().Cells(SlotTop(slot) + OFF_TITLE, StripCol(lane))) = taskId Then
            DoneIndexOfId = i
            Exit Function
        End If
    Next i
    DoneIndexOfId = -1
End Function

Private Function DoneIndexOf(ByVal lane As Long, ByVal slot As Long) As Long
    Dim count As Long
    count = DoneNoteCount()
    Dim index As Long
    index = lane * MAX_SLOTS + slot
    If index >= 0 And index < count Then
        DoneIndexOf = index
    Else
        DoneIndexOf = -1
    End If
End Function

Private Function BoardNoteCount() As Long
    Dim lane As Long
    Dim n As Long
    For lane = 0 To BOARD_LANES - 1
        n = n + LaneCount(SheetBoard(), lane)
    Next lane
    BoardNoteCount = n
End Function

Private Function DoneNoteCount() As Long
    Dim i As Long
    Dim lane As Long
    Dim slot As Long
    For i = 0 To DONE_LANES * MAX_SLOTS - 1
        LinearToPos i, lane, slot
        If Not SlotOccupied(SheetDone(), lane, slot) Then
            DoneNoteCount = i
            Exit Function
        End If
    Next i
    DoneNoteCount = DONE_LANES * MAX_SLOTS
End Function

Private Function LaneCount(ByVal ws As Worksheet, ByVal lane As Long) As Long
    Dim s As Long
    Dim n As Long
    For s = 0 To MAX_SLOTS - 1
        If SlotOccupied(ws, lane, s) Then n = n + 1
    Next s
    LaneCount = n
End Function

Private Function ActiveOccupiedNote(ByVal ws As Worksheet, ByVal laneCount As Long, ByRef lane As Long, ByRef slot As Long) As Boolean
    If ActiveCell Is Nothing Then Exit Function
    If Not LocateCell(ws, ActiveCell.Row, ActiveCell.Column, lane, slot) Then Exit Function
    If lane < 0 Or lane >= laneCount Then Exit Function
    ActiveOccupiedNote = SlotOccupied(ws, lane, slot)
End Function

Private Sub RevealNote(ByVal taskId As String)
    If SuppressPrompts Then Exit Sub
    Dim ws As Worksheet
    Dim lane As Long
    Dim slot As Long
    Dim foundLane As Long
    Dim foundSlot As Long
    If FindIdOnSheet(SheetBoard(), taskId, foundLane, foundSlot) Then
        Set ws = SheetBoard()
        lane = foundLane
        slot = foundSlot
    ElseIf FindIdOnSheet(SheetDone(), taskId, foundLane, foundSlot) Then
        Set ws = SheetDone()
        lane = foundLane
        slot = foundSlot
    Else
        Exit Sub
    End If
    Application.Goto ws.Cells(SlotTop(slot) + OFF_TITLE, BodyCol(lane)), True
End Sub

Private Function ConfirmDelete() As Boolean
    If SuppressPrompts Then
        ConfirmDelete = True
        Exit Function
    End If
    ConfirmDelete = (MsgBox( _
        "選択中の付箋を削除します。" & vbCrLf & "この操作は元に戻せません。削除しますか？", _
        vbYesNo + vbQuestion + vbDefaultButton2, _
        "タスク付箋ノート") = vbYes)
End Function

Private Sub Notify(ByVal message As String)
    LastNotice = message
    If Not SuppressPrompts Then
        MsgBox message, vbExclamation, "タスク付箋ノート"
    End If
End Sub

Private Function NextId() As String
    Dim n As Long
    n = CurrentCounter()
    Dim maxExisting As Long
    maxExisting = MaxExistingIdNumber()
    If maxExisting + 1 > n Then n = maxExisting + 1
    If n < 1 Then n = 1
    NextId = "ST-" & Format$(n, "000")
    On Error Resume Next
    ThisWorkbook.Names("NextStickyId").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="NextStickyId", RefersTo:="=" & CStr(n + 1)
End Function

Private Function CurrentCounter() As Long
    On Error Resume Next
    CurrentCounter = CLng(Evaluate(ThisWorkbook.Names("NextStickyId").RefersTo))
    If Err.Number <> 0 Or CurrentCounter < 1 Then CurrentCounter = 1
    On Error GoTo 0
End Function

Private Sub EnsureCounter()
    If CurrentCounter() < 1 Then
        ThisWorkbook.Names.Add Name:="NextStickyId", RefersTo:="=1"
    End If
End Sub

Private Function MaxExistingIdNumber() As Long
    Dim maxN As Long
    maxN = 0
    maxN = MaxIdOnSheet(SheetBoard(), BOARD_LANES, maxN)
    maxN = MaxIdOnSheet(SheetDone(), DONE_LANES, maxN)
    MaxExistingIdNumber = maxN
End Function

Private Function MaxIdOnSheet(ByVal ws As Worksheet, ByVal laneCount As Long, ByVal seed As Long) As Long
    Dim lane As Long
    Dim s As Long
    Dim n As Long
    n = seed
    For lane = 0 To laneCount - 1
        For s = 0 To MAX_SLOTS - 1
            Dim idn As Long
            idn = IdNumber(CellText(ws.Cells(SlotTop(s) + OFF_TITLE, StripCol(lane))))
            If idn > n Then n = idn
        Next s
    Next lane
    MaxIdOnSheet = n
End Function

Private Function IdNumber(ByVal id As String) As Long
    If Len(id) > 3 Then
        If Left$(id, 3) = "ST-" Then
            If IsNumeric(Mid$(id, 4)) Then
                IdNumber = CLng(Mid$(id, 4))
                Exit Function
            End If
        End If
    End If
    IdNumber = 0
End Function

Private Sub UpdateBoardHeaders()
    Dim ws As Worksheet
    Set ws = SheetBoard()
    Dim lane As Long
    For lane = 0 To BOARD_LANES - 1
        Dim stripColor As Long
        Dim titleColor As Long
        Dim bodyColor As Long
        Dim captionColor As Long
        Dim borderColor As Long
        Dim fontColor As Long
        Palette StatusOfLane(lane), stripColor, titleColor, bodyColor, captionColor, borderColor, fontColor
        Dim header As Range
        Set header = ws.Range(ws.Cells(HEADER_ROW, StripCol(lane)), ws.Cells(HEADER_ROW, BodyCol(lane)))
        On Error Resume Next
        header.UnMerge
        On Error GoTo 0
        Dim prevAlerts As Boolean
        prevAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        header.Merge
        Application.DisplayAlerts = prevAlerts
        header.Cells(1, 1).Value = StatusOfLane(lane) & "（" & CStr(LaneCount(ws, lane)) & "）"
        With header
            .Interior.Pattern = xlSolid
            .Interior.Color = titleColor
            .Font.Name = FONT_NAME
            .Font.Size = 13
            .Font.Bold = True
            .Font.Color = RGB(50, 40, 30)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Locked = True
        End With
        header.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, Color:=borderColor
    Next lane
End Sub

Private Sub UpdateDoneTitle()
    Dim ws As Worksheet
    Set ws = SheetDone()
    ws.Range("A1").Value = "完了した付箋（" & CStr(DoneNoteCount()) & "）"
End Sub

Private Sub ReformatAll(ByVal ws As Worksheet, ByVal laneCount As Long)
    Dim lane As Long
    Dim s As Long
    For lane = 0 To laneCount - 1
        For s = 0 To MAX_SLOTS - 1
            If SlotOccupied(ws, lane, s) Then
                WriteSlot ws, lane, s, ReadSlot(ws, lane, s)
            Else
                ClearSlot ws, lane, s
            End If
        Next s
    Next lane
End Sub

Private Sub RefreshRowBands(ByVal ws As Worksheet, ByVal laneCount As Long)
    Dim maxUsed As Long
    maxUsed = -1
    Dim lane As Long
    Dim s As Long
    For lane = 0 To laneCount - 1
        For s = 0 To MAX_SLOTS - 1
            If SlotOccupied(ws, lane, s) Then
                If s > maxUsed Then maxUsed = s
            End If
        Next s
    Next lane
    For s = 0 To MAX_SLOTS - 1
        Dim band As Range
        Set band = ws.Rows(CStr(SlotTop(s)) & ":" & CStr(SlotTop(s) + SLOT_HEIGHT - 1))
        If s <= maxUsed Then
            band.Hidden = False
        Else
            band.Hidden = True
        End If
    Next s
End Sub

Private Sub PrepareAllSlotHeights(ByVal ws As Worksheet)
    Dim s As Long
    For s = 0 To MAX_SLOTS - 1
        ApplySlotHeights ws, s
    Next s
End Sub

Private Sub ApplySlotHeights(ByVal ws As Worksheet, ByVal slot As Long)
    Dim top As Long
    top = SlotTop(slot)
    ws.Rows(top + 0).RowHeight = 8
    ws.Rows(top + 1).RowHeight = 15
    ws.Rows(top + 2).RowHeight = 36
    ws.Rows(top + 3).RowHeight = 15
    ws.Rows(top + 4).RowHeight = 68
    ws.Rows(top + 5).RowHeight = 15
    ws.Rows(top + 6).RowHeight = 50
    ws.Rows(top + 7).RowHeight = 15
    ws.Rows(top + 8).RowHeight = 22
End Sub

Private Function SheetBoard() As Worksheet
    Set SheetBoard = ThisWorkbook.Worksheets(SHEET_BOARD)
End Function

Private Function SheetDone() As Worksheet
    Set SheetDone = ThisWorkbook.Worksheets(SHEET_DONE)
End Function

Private Function SheetHelp() As Worksheet
    Set SheetHelp = ThisWorkbook.Worksheets(SHEET_HELP)
End Function

Private Sub CreateAllSheets()
    EnsureSheet SHEET_BOARD
    EnsureSheet SHEET_DONE
    EnsureSheet SHEET_HELP
    SheetBoard().Move Before:=ThisWorkbook.Worksheets(1)
    SheetDone().Move After:=SheetBoard()
    SheetHelp().Move After:=SheetDone()
End Sub

Private Function EnsureSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = name
    End If
End Function

Private Sub DeleteStraySheets()
    Dim guard As Long
    Dim deleted As Boolean
    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Do
        deleted = False
        Dim ws As Worksheet
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name <> SHEET_BOARD And ws.Name <> SHEET_DONE And ws.Name <> SHEET_HELP Then
                If ThisWorkbook.Worksheets.Count > 1 Then
                    ws.Delete
                    deleted = True
                    Exit For
                End If
            End If
        Next ws
        guard = guard + 1
    Loop While deleted And guard < 8
    Application.DisplayAlerts = prevAlerts
End Sub

Private Sub SetupBoardChrome()
    Dim ws As Worksheet
    Set ws = SheetBoard()
    UnprotectSheet ws
    ws.Range("A1:R260").Clear
    ApplyColumnWidths ws, BOARD_LANES
    PaintNotebook ws, BOARD_LANES, RGB(92, 74, 58)
    ws.Rows(1).RowHeight = 38
    ws.Rows(2).RowHeight = 34
    ws.Rows(3).RowHeight = 32
    ws.Rows(4).RowHeight = 28
    ws.Rows(5).RowHeight = 8
    Dim title As Range
    Set title = ws.Range(ws.Cells(1, 1), ws.Cells(1, BodyCol(BOARD_LANES - 1)))
    title.Merge
    title.Cells(1, 1).Value = "タスク付箋ノート"
    With title
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(92, 74, 58)
        .Font.Name = FONT_NAME
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(255, 250, 240)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Locked = True
    End With
    Dim info As Range
    Set info = ws.Range(ws.Cells(2, 1), ws.Cells(2, BodyCol(BOARD_LANES - 1)))
    info.Merge
    info.Cells(1, 1).Value = "「付箋を貼る」で黄色い付箋を追加します。タイトル、タスクの内容、備考を書き、下のステータスを変えると色と列が変わります。「完了」でグレーになり、完了シートへ移動します。文字は Meiryo UI です。"
    With info
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(243, 237, 224)
        .Font.Name = FONT_NAME
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(70, 60, 50)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .IndentLevel = 1
        .Locked = True
    End With
    UpdateBoardHeaders
    ws.Tab.Color = RGB(214, 158, 18)
End Sub

Private Sub SetupDoneChrome()
    Dim ws As Worksheet
    Set ws = SheetDone()
    UnprotectSheet ws
    ws.Range("A1:R260").Clear
    ApplyColumnWidths ws, DONE_LANES
    PaintNotebook ws, DONE_LANES, RGB(96, 96, 96)
    ws.Rows(1).RowHeight = 38
    ws.Rows(2).RowHeight = 34
    ws.Rows(3).RowHeight = 32
    ws.Rows(4).RowHeight = 28
    ws.Rows(5).RowHeight = 8
    Dim title As Range
    Set title = ws.Range(ws.Cells(1, 1), ws.Cells(1, BodyCol(DONE_LANES - 1)))
    title.Merge
    title.Cells(1, 1).Value = "完了した付箋（0）"
    With title
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(96, 96, 96)
        .Font.Name = FONT_NAME
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Locked = True
    End With
    Dim info As Range
    Set info = ws.Range(ws.Cells(2, 1), ws.Cells(2, BodyCol(DONE_LANES - 1)))
    info.Merge
    info.Cells(1, 1).Value = "完了した付箋はここに集まります。左上ほど新しいものです。戻すときは付箋を選んで「ノートに戻す」を押します。完了前の色と列に戻ります。"
    With info
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(232, 232, 230)
        .Font.Name = FONT_NAME
        .Font.Size = 10
        .Font.Color = RGB(60, 60, 60)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .IndentLevel = 1
        .Locked = True
    End With
    Dim lane As Long
    For lane = 0 To DONE_LANES - 1
        Dim header As Range
        Set header = ws.Range(ws.Cells(HEADER_ROW, StripCol(lane)), ws.Cells(HEADER_ROW, BodyCol(lane)))
        header.Merge
        If lane = 0 Then
            header.Cells(1, 1).Value = "新しい順"
        Else
            header.Cells(1, 1).Value = "完了"
        End If
        With header
            .Interior.Pattern = xlSolid
            .Interior.Color = RGB(186, 186, 186)
            .Font.Name = FONT_NAME
            .Font.Size = 12
            .Font.Bold = True
            .Font.Color = RGB(50, 50, 50)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Locked = True
        End With
    Next lane
    ws.Tab.Color = RGB(128, 128, 128)
End Sub

Private Sub SetupHelp()
    Dim ws As Worksheet
    Set ws = SheetHelp()
    UnprotectSheet ws
    ws.Range("A1:E40").Clear
    ws.Columns(1).ColumnWidth = 3
    ws.Columns(2).ColumnWidth = 78
    ws.Columns(3).ColumnWidth = 8
    ws.Columns(4).ColumnWidth = 16
    ws.Columns(5).ColumnWidth = 46
    ws.Range("A1:E40").Interior.Pattern = xlSolid
    ws.Range("A1:E40").Interior.Color = RGB(252, 250, 246)
    ws.Range("A1:E40").Font.Name = FONT_NAME
    ws.Range("A1:E40").Font.Size = 12
    ws.Range("A1:E40").Font.Color = RGB(45, 40, 35)
    ws.Range("B1:E1").Merge
    With ws.Range("B1")
        .Value = "使い方"
        .Font.Name = FONT_NAME
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = RGB(92, 74, 58)
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 36

    PutHelpLine ws, 3, "1. このブックを開いたら「コンテンツの有効化」を押し、マクロを有効にしてください。色の移動はマクロで行います。"
    PutHelpLine ws, 4, "2. 「ノート」の「付箋を貼る」で、起票済み（黄色）の付箋が追加されます。"
    PutHelpLine ws, 5, "3. 各付箋に、タイトル、タスクの内容、備考（関係する実装、チケット、ファイル名など）を直接入力します。"
    PutHelpLine ws, 6, "4. 付箋の文字は Meiryo UI です。入力したあとも Meiryo UI のままです。"
    PutHelpLine ws, 7, "5. 付箋一番下のステータスを選ぶと、色が変わって対応する列へ自動で移動します。"
    PutHelpLine ws, 8, "6. 「完了」を選ぶと、付箋はグレーになり、「完了」シートの先頭へ自動で移動します。ノートから剥がして捨てるイメージです。"
    PutHelpLine ws, 9, "7. 誤って完了にした場合は、「完了」シートでその付箋を選んで「ノートに戻す」を押します。完了前のステータスと色に戻ります。"
    PutHelpLine ws, 10, "8. 不要な付箋は、「ノート」でその付箋を選んで「この付箋を削除」を押します。"
    PutHelpLine ws, 11, "9. マクロで移動した操作は、Ctrl+Z では戻せません。ステータスを選び直すか、「ノートに戻す」を使ってください。"

    ws.Range("B13").Value = "色の対応"
    ws.Range("B13").Font.Name = FONT_NAME
    ws.Range("B13").Font.Size = 16
    ws.Range("B13").Font.Bold = True
    ws.Rows(13).RowHeight = 26

    ws.Range("C15").Value = "色"
    ws.Range("D15").Value = "ステータス"
    ws.Range("E15").Value = "意味"
    ws.Range("C15:E15").Font.Name = FONT_NAME
    ws.Range("C15:E15").Font.Bold = True
    ws.Range("C15:E15").Font.Size = 11

    PutSwatch ws, 16, ST_NEW, "貼ったばかり。まだ手を付けていない"
    PutSwatch ws, 17, ST_DOING, "今取り組んでいる"
    PutSwatch ws, 18, ST_HANDLED, "自分の作業は終わった"
    PutSwatch ws, 19, ST_LINKED, "関係先へ連携した"
    PutSwatch ws, 20, ST_DONE, "グレーにしてノートから外し、完了シートへ置く"

    ws.Range("B22").Value = "列の順は、起票済み、対応中、対応済み、連携済みです。どのステータスにも直接変えられます。"
    ws.Range("B22").Font.Name = FONT_NAME
    ws.Range("B22").Font.Size = 12
    ws.Range("B22").WrapText = True
    ws.Rows(22).RowHeight = 36
    ws.Tab.Color = RGB(90, 110, 130)
    ws.Range("A1:E40").Locked = True
End Sub

Private Sub PutHelpLine(ByVal ws As Worksheet, ByVal row As Long, ByVal text As String)
    With ws.Cells(row, 2)
        .Value = text
        .Font.Name = FONT_NAME
        .Font.Size = 12
        .Font.Bold = False
        .Font.Color = RGB(45, 40, 35)
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(row).RowHeight = 34
End Sub

Private Sub PutSwatch(ByVal ws As Worksheet, ByVal row As Long, ByVal status As String, ByVal meaning As String)
    Dim stripColor As Long
    Dim titleColor As Long
    Dim bodyColor As Long
    Dim captionColor As Long
    Dim borderColor As Long
    Dim fontColor As Long
    Palette status, stripColor, titleColor, bodyColor, captionColor, borderColor, fontColor
    With ws.Cells(row, 3)
        .Interior.Pattern = xlSolid
        .Interior.Color = bodyColor
        .Borders.LineStyle = xlContinuous
        .Borders.Color = borderColor
        .Borders.Weight = xlMedium
    End With
    With ws.Cells(row, 4)
        .Value = status
        .Font.Name = FONT_NAME
        .Font.Size = 12
        .Font.Bold = True
        .Interior.Pattern = xlSolid
        .Interior.Color = titleColor
        .Font.Color = fontColor
    End With
    With ws.Cells(row, 5)
        .Value = meaning
        .Font.Name = FONT_NAME
        .Font.Size = 12
        .Font.Color = RGB(45, 40, 35)
    End With
    ws.Rows(row).RowHeight = 24
End Sub

Private Sub PaintNotebook(ByVal ws As Worksheet, ByVal laneCount As Long, ByVal cover As Long)
    Dim lastCol As Long
    lastCol = BodyCol(laneCount - 1)
    Dim area As Range
    Set area = ws.Range(ws.Cells(1, 1), ws.Cells(SlotTop(MAX_SLOTS) + 2, lastCol))
    area.Interior.Pattern = xlSolid
    area.Interior.Color = PaperColor(ws)
    area.Font.Name = FONT_NAME
    With ws.Range(ws.Cells(1, 1), ws.Cells(SlotTop(MAX_SLOTS) + 2, 1)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        If ws.Name = SHEET_DONE Then
            .Color = RGB(180, 180, 180)
        Else
            .Color = RGB(214, 150, 150)
        End If
    End With
    ws.Columns(1).ColumnWidth = 2.6
End Sub

Private Sub ApplyColumnWidths(ByVal ws As Worksheet, ByVal laneCount As Long)
    ws.Columns(1).ColumnWidth = 2.6
    Dim lane As Long
    For lane = 0 To laneCount - 1
        ws.Columns(StripCol(lane)).ColumnWidth = 1.8
        ws.Columns(BodyCol(lane)).ColumnWidth = 32
        If lane < laneCount - 1 Then
            Dim gapCol As Long
            gapCol = BodyCol(lane) + 1
            ws.Columns(gapCol).ColumnWidth = 2.3
            ws.Range(ws.Cells(1, gapCol), ws.Cells(SlotTop(MAX_SLOTS) + 2, gapCol)).Interior.Color = PaperColor(ws)
        End If
    Next lane
End Sub

Private Sub ApplyBaseView(ByVal ws As Worksheet, ByVal scroll As String, ByVal zoom As Long)
    ws.Activate
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
        .Zoom = zoom
        .FreezePanes = False
    End With
    ws.Range("A6").Select
    ActiveWindow.FreezePanes = True
    ws.ScrollArea = scroll
End Sub

Private Sub EnsureButtons()
    BuildButton SheetBoard(), "btnAdd", "付箋を貼る", "付箋を貼る", 2, RGB(255, 214, 74), RGB(180, 140, 30), RGB(50, 40, 20), 132
    BuildButton SheetBoard(), "btnDelete", "この付箋を削除", "この付箋を削除", 146, RGB(220, 214, 204), RGB(150, 140, 130), RGB(50, 40, 30), 156
    BuildButton SheetDone(), "btnRestore", "ノートに戻す", "完了付箋をノートに戻す", 2, RGB(186, 186, 186), RGB(110, 110, 110), RGB(40, 40, 40), 150
End Sub

Private Sub BuildButton(ByVal ws As Worksheet, ByVal shapeName As String, ByVal caption As String, ByVal macroName As String, ByVal leftOffset As Double, ByVal fillColor As Long, ByVal lineColor As Long, ByVal fontColor As Long, ByVal width As Double)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
    Dim anchor As Range
    Set anchor = ws.Cells(3, 2)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, anchor.Left + leftOffset, anchor.Top + 3, width, 26)
    shp.Name = shapeName
    shp.OnAction = macroName
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.ForeColor.RGB = lineColor
    shp.Line.Weight = 1
    shp.Placement = xlMove
    shp.Locked = True
    shp.TextFrame.Characters.Text = caption
    With shp.TextFrame.Characters.Font
        .Name = FONT_NAME
        .Size = 12
        .Bold = True
        .Color = fontColor
    End With
    shp.TextFrame.HorizontalAlignment = xlCenter
    shp.TextFrame.VerticalAlignment = xlCenter
    shp.TextFrame.MarginLeft = 4
    shp.TextFrame.MarginRight = 4
End Sub

Private Sub EnsureWritable(ByVal ws As Worksheet)
    ApplyProtection ws
End Sub

Private Sub ProtectAll()
    ApplyProtection SheetBoard()
    ApplyProtection SheetDone()
    ApplyProtection SheetHelp()
End Sub

Private Sub UnprotectAll()
    UnprotectSheet SheetBoard()
    UnprotectSheet SheetDone()
    UnprotectSheet SheetHelp()
End Sub

Private Sub UnprotectSheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=""
    On Error GoTo 0
End Sub

Private Sub ApplyProtection(ByVal ws As Worksheet)
    UnprotectSheet ws
    ws.Protect Password:="", _
        DrawingObjects:=True, _
        Contents:=True, _
        Scenarios:=True, _
        UserInterfaceOnly:=True, _
        AllowFormattingCells:=False, _
        AllowFormattingColumns:=False, _
        AllowFormattingRows:=True, _
        AllowInsertingHyperlinks:=False, _
        AllowSorting:=False, _
        AllowFiltering:=False
    ws.EnableSelection = xlNoRestrictions
End Sub

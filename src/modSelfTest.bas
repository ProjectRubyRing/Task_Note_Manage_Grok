Option Explicit

' ビルド時だけ注入する自己テスト。完成した xlsm には残さない。

Private Const TOKEN As String = "FIRST=6;SLOT=9;GAP=0;CAP_TITLE=1;TITLE=2;CAP_BODY=3;BODY=4;CAP_REMARK=5;REMARK=6;CAP_STATUS=7;STATUS=8;STRIP0=2;BODY0=3;STEP=3;MAX=25;BLANES=4;DLANES=3;FONT=Meiryo UI"
Private Const FONT_NAME As String = "Meiryo UI"
Private Const FIRST_ROW As Long = 6
Private Const SLOT_HEIGHT As Long = 9
Private Const OFF_TITLE As Long = 2
Private Const OFF_BODY As Long = 4
Private Const OFF_REMARK As Long = 6
Private Const OFF_STATUS As Long = 8
Private Const OFF_CAP_TITLE As Long = 1
Private Const OFF_CAP_BODY As Long = 3
Private Const OFF_CAP_REMARK As Long = 5
Private Const OFF_CAP_STATUS As Long = 7

Private report As String
Private fails As Long

Public Function RunSelfTest() As String
    On Error GoTo Blow
    report = ""
    fails = 0
    modCore.SuppressPrompts = True
    modCore.ResetRuntimeState
    modCore.SuppressPrompts = True
    modCore.LastNotice = ""

    Check "layout", modCore.LayoutToken() = TOKEN, modCore.LayoutToken()
    Check "font-normal", ThisWorkbook.Styles("Normal").Font.Name = FONT_NAME, ThisWorkbook.Styles("Normal").Font.Name

    Dim snap As String
    snap = Snapshot()

    CheckSample "週次レポートの自動化", "ノート", 0, "起票済み", "毎週月曜の手作業集計をなくしたい。", "実装候補: Power Query。集計ブック SalesWeekly.xlsx", RGB(255, 236, 130), RGB(255, 214, 74), RGB(214, 158, 18), RGB(45, 40, 35), False, ""
    CheckSample "ログイン失敗の調査", "ノート", 1, "対応中", "ステージング環境で 401 が断続的に返る。再現手順を整理する。", "AuthFilter 周辺。チケット #441", RGB(255, 204, 150), RGB(255, 166, 86), RGB(214, 106, 28), RGB(45, 40, 35), False, ""
    CheckSample "日付の入力チェック", "ノート", 2, "対応済み", "未入力のとき保存できないようチェックを入れた。", "FormValidator の日付判定", RGB(186, 228, 176), RGB(138, 206, 132), RGB(56, 146, 72), RGB(45, 40, 35), False, ""
    CheckSample "APIエラーコードの共有", "ノート", 3, "連携済み", "連携先へエラーコード一覧を送付し、確認待ち。", "docs/error-codes.md を更新して共有済み", RGB(176, 214, 250), RGB(126, 184, 240), RGB(36, 116, 196), RGB(45, 40, 35), False, ""
    CheckSample "開発環境の初期構築", "完了", 0, "完了", "ローカルで起動する手順を整備した。", "README のセットアップ手順", RGB(210, 210, 210), RGB(186, 186, 186), RGB(128, 128, 128), RGB(80, 80, 80), True, "起票済み"

    Check "captions", CaptionsOk(), "caption mismatch"
    Check "headers", HeaderText() = "起票済み（1）|対応中（1）|対応済み（1）|連携済み（1）", HeaderText()
    Check "done-title", InStr(1, CStr(SheetDone().Range("A1").Value), "完了した付箋（1）", vbBinaryCompare) = 1, CStr(SheetDone().Range("A1").Value)
    Check "slot1-hidden", SheetBoard().Rows(FIRST_ROW + SLOT_HEIGHT).Hidden = True, "row " & CStr(FIRST_ROW + SLOT_HEIGHT)
    Check "slot0-visible", SheetBoard().Rows(FIRST_ROW).Hidden = False, "row " & CStr(FIRST_ROW)
    Check "help-swatch-new", CLng(SheetHelp().Cells(16, 3).Interior.Color) = RGB(255, 236, 130), CStr(SheetHelp().Cells(16, 3).Interior.Color)
    Check "help-swatch-doing", CLng(SheetHelp().Cells(17, 3).Interior.Color) = RGB(255, 204, 150), CStr(SheetHelp().Cells(17, 3).Interior.Color)
    Check "help-swatch-handled", CLng(SheetHelp().Cells(18, 3).Interior.Color) = RGB(186, 228, 176), CStr(SheetHelp().Cells(18, 3).Interior.Color)
    Check "help-swatch-linked", CLng(SheetHelp().Cells(19, 3).Interior.Color) = RGB(176, 214, 250), CStr(SheetHelp().Cells(19, 3).Interior.Color)
    Check "help-swatch-done", CLng(SheetHelp().Cells(20, 3).Interior.Color) = RGB(210, 210, 210), CStr(SheetHelp().Cells(20, 3).Interior.Color)
    Check "help-font", SheetHelp().Cells(6, 2).Font.Name = FONT_NAME, SheetHelp().Cells(6, 2).Font.Name
    Check "protect-board", SheetBoard().ProtectContents = True, "board"
    Check "protect-done", SheetDone().ProtectContents = True, "done"
    Check "protect-help", SheetHelp().ProtectContents = True, "help"

    Dim shp As Shape
    Set shp = SheetBoard().Shapes("btnAdd")
    Check "button-action", InStr(1, shp.OnAction, "付箋を貼る", vbBinaryCompare) > 0, shp.OnAction
    Check "button-font", shp.TextFrame.Characters.Font.Name = FONT_NAME, shp.TextFrame.Characters.Font.Name
    Check "button-delete", InStr(1, SheetBoard().Shapes("btnDelete").OnAction, "この付箋を削除", vbBinaryCompare) > 0, SheetBoard().Shapes("btnDelete").OnAction
    Check "button-restore", InStr(1, SheetDone().Shapes("btnRestore").OnAction, "完了付箋をノートに戻す", vbBinaryCompare) > 0, SheetDone().Shapes("btnRestore").OnAction

    ' 選択していない削除は何も消さない
    SheetBoard().Activate
    Application.Goto SheetBoard().Range("A1"), True
    modCore.LastNotice = ""
    この付箋を削除
    Check "delete-needs-selection", InStr(1, modCore.LastNotice, "クリック", vbBinaryCompare) > 0, modCore.LastNotice
    Check "delete-keeps-sample", CountTitle("週次レポートの自動化") = 1, CStr(CountTitle("週次レポートの自動化"))

    ' ボタンで貼る
    modCore.LastNotice = ""
    付箋を貼る
    Dim idMove As String
    idMove = modCore.LastCreatedId
    Check "add-id", Len(idMove) > 0, idMove
    Check "add-notice", Len(modCore.LastNotice) = 0, modCore.LastNotice
    Dim lane As Long
    Dim slot As Long
    Dim ws As Worksheet
    Check "add-found", FindId(idMove, ws, lane, slot), idMove
    If fails = 0 Or Not ws Is Nothing Then
        Check "add-sheet", ws.Name = "ノート", SafeName(ws)
        Check "add-lane", lane = 0, CStr(lane)
        Check "add-slot", slot = 1, CStr(slot)
        Check "add-status", StatusText(ws, lane, slot) = "起票済み", StatusText(ws, lane, slot)
        Check "add-font", Field(ws, lane, slot, OFF_TITLE).Font.Name = FONT_NAME, Field(ws, lane, slot, OFF_TITLE).Font.Name
        Check "add-yellow", CLng(Field(ws, lane, slot, OFF_BODY).Interior.Color) = RGB(255, 236, 130), CStr(Field(ws, lane, slot, OFF_BODY).Interior.Color)
        Check "add-placeholder-color", CLng(Field(ws, lane, slot, OFF_TITLE).Font.Color) = RGB(145, 134, 110), CStr(Field(ws, lane, slot, OFF_TITLE).Font.Color)
        Check "add-placeholder-text", FieldText(ws, lane, slot, OFF_TITLE) = "タイトルを入力", FieldText(ws, lane, slot, OFF_TITLE)
        Check "add-unlocked", Field(ws, lane, slot, OFF_TITLE).Locked = False, "title locked"
        Check "add-caption-locked", Field(ws, lane, slot, OFF_CAP_TITLE).Locked = True, "caption"
        Check "add-validation", HasList(Field(ws, lane, slot, OFF_STATUS)), "no list"
        Check "add-bold", Field(ws, lane, slot, OFF_TITLE).Font.Bold = True, "bold"
        Check "add-wrap", Field(ws, lane, slot, OFF_BODY).WrapText = True, "wrap"
    End If

    ' 直接入力しても列は動かず、フォントは Meiryo UI のまま
    SetField idMove, OFF_TITLE, "【TEST】移動"
    SetField idMove, OFF_BODY, "移動テストの内容です。"
    SetField idMove, OFF_REMARK, "実装メモ: SampleService.cs"
    Check "type-stays", FindId(idMove, ws, lane, slot) And ws.Name = "ノート" And lane = 0, WhereText(idMove)
    Check "type-font", Field(ws, lane, slot, OFF_TITLE).Font.Name = FONT_NAME, Field(ws, lane, slot, OFF_TITLE).Font.Name
    Check "type-color", CLng(Field(ws, lane, slot, OFF_TITLE).Font.Color) = RGB(45, 40, 35), CStr(Field(ws, lane, slot, OFF_TITLE).Font.Color)
    Check "type-body", FieldText(ws, lane, slot, OFF_BODY) = "移動テストの内容です。", FieldText(ws, lane, slot, OFF_BODY)
    Check "type-remark", FieldText(ws, lane, slot, OFF_REMARK) = "実装メモ: SampleService.cs", FieldText(ws, lane, slot, OFF_REMARK)

    ' 不正なステータスは残さない
    Dim statusCell As Range
    Set statusCell = Field(ws, lane, slot, OFF_STATUS)
    statusCell.Validation.ShowError = False
    modCore.LastNotice = ""
    statusCell.Value = "未定"
    Check "invalid-reverted", FindId(idMove, ws, lane, slot) And ws.Name = "ノート" And lane = 0 And StatusText(ws, lane, slot) = "起票済み", WhereText(idMove) & " / " & StatusText(ws, lane, slot) & " / " & modCore.LastNotice

    MoveTo idMove, "対応中"
    Check "move-doing", PlaceIs(idMove, "ノート", 1, "対応中", RGB(255, 204, 150), RGB(45, 40, 35)), WhereText(idMove)
    Check "move-doing-text", TextIs(idMove, "【TEST】移動", "移動テストの内容です。", "実装メモ: SampleService.cs"), WhereText(idMove)
    Check "move-doing-once", CountTitle("【TEST】移動") = 1, CStr(CountTitle("【TEST】移動"))

    MoveTo idMove, "対応済み"
    Check "move-handled", PlaceIs(idMove, "ノート", 2, "対応済み", RGB(186, 228, 176), RGB(45, 40, 35)), WhereText(idMove)

    MoveTo idMove, "連携済み"
    Check "move-linked", PlaceIs(idMove, "ノート", 3, "連携済み", RGB(176, 214, 250), RGB(45, 40, 35)), WhereText(idMove)
    SetField idMove, OFF_REMARK, "実装メモ: SampleService.cs を更新"
    Check "edit-keeps-lane", PlaceIs(idMove, "ノート", 3, "連携済み", RGB(176, 214, 250), RGB(45, 40, 35)), WhereText(idMove)
    Check "edit-remark", TextIs(idMove, "【TEST】移動", "移動テストの内容です。", "実装メモ: SampleService.cs を更新"), "remark"

    MoveTo idMove, "完了"
    Check "done-sheet", PlaceIs(idMove, "完了", 0, "完了", RGB(210, 210, 210), RGB(80, 80, 80)), WhereText(idMove)
    Check "done-text", TextIs(idMove, "【TEST】移動", "移動テストの内容です。", "実装メモ: SampleService.cs を更新"), "done text"
    Check "done-prev", PrevOf(idMove) = "連携済み", PrevOf(idMove)
    Check "done-when", Len(CompletedOf(idMove)) > 0, CompletedOf(idMove)
    Check "done-display", Left$(StatusTextOf(idMove), 2) = "完了", StatusTextOf(idMove)
    Check "done-locked", LockedTitle(idMove) = True, "locked"
    Check "done-no-list", Not HasListOf(idMove), "list remains"
    Check "done-not-on-board", Not OnSheet(idMove, "ノート"), WhereText(idMove)
    Check "sample-pushed", FindTitle("開発環境の初期構築", ws, lane, slot) And ws.Name = "完了" And slot = 1, "sample slot " & CStr(slot)

    ' 新しい完了が上に来る
    付箋を貼る
    Dim idTop As String
    idTop = modCore.LastCreatedId
    SetField idTop, OFF_TITLE, "【TEST】後から完了"
    SetField idTop, OFF_BODY, "二枚目"
    SetField idTop, OFF_REMARK, "備考B"
    MoveTo idTop, "対応中"
    MoveTo idTop, "完了"
    Check "newest-top", PlaceIs(idTop, "完了", 0, "完了", RGB(210, 210, 210), RGB(80, 80, 80)) And SlotOf(idTop) = 0, WhereText(idTop)
    Check "older-below", SlotOf(idMove) = 1, WhereText(idMove)
    Check "top-prev", PrevOf(idTop) = "対応中", PrevOf(idTop)

    ' 公開マクロで戻す
    Dim backWs As Worksheet
    Dim backLane As Long
    Dim backSlot As Long
    FindId idTop, backWs, backLane, backSlot
    backWs.Activate
    Application.Goto Field(backWs, backLane, backSlot, OFF_TITLE), True
    modCore.LastNotice = ""
    完了付箋をノートに戻す
    Check "restore-lane", PlaceIs(idTop, "ノート", 1, "対応中", RGB(255, 204, 150), RGB(45, 40, 35)), WhereText(idTop) & " / " & modCore.LastNotice
    Check "restore-text", TextIs(idTop, "【TEST】後から完了", "二枚目", "備考B"), "restore text"
    Check "restore-gone-from-done", Not OnSheet(idTop, "完了"), WhereText(idTop)

    ' 公開マクロで削除
    FindId idTop, backWs, backLane, backSlot
    backWs.Activate
    Application.Goto Field(backWs, backLane, backSlot, OFF_TITLE), True
    modCore.LastNotice = ""
    この付箋を削除
    Check "deleted", CountTitle("【TEST】後から完了") = 0, CStr(CountTitle("【TEST】後から完了")) & " / " & modCore.LastNotice

    ' 列の詰め：上の付箋を動かすと下の付箋が繰り上がる
    付箋を貼る
    Dim idA As String
    idA = modCore.LastCreatedId
    SetField idA, OFF_TITLE, "【TEST】上"
    SetField idA, OFF_BODY, "A"
    SetField idA, OFF_REMARK, "A備考"
    付箋を貼る
    Dim idB As String
    idB = modCore.LastCreatedId
    SetField idB, OFF_TITLE, "【TEST】下"
    SetField idB, OFF_BODY, "B"
    SetField idB, OFF_REMARK, "B備考"
    Check "stack", SlotOf(idA) = 1 And SlotOf(idB) = 2 And LaneOf(idA) = 0 And LaneOf(idB) = 0, "A=" & CStr(SlotOf(idA)) & " B=" & CStr(SlotOf(idB))
    MoveTo idA, "対応中"
    Check "compact", LaneOf(idB) = 0 And SlotOf(idB) = 1, "B lane " & CStr(LaneOf(idB)) & " slot " & CStr(SlotOf(idB))
    Check "compact-text", TextIs(idB, "【TEST】下", "B", "B備考"), "B text"
    Check "sample-stays", LaneOfTitle("週次レポートの自動化") = 0 And SlotOfTitle("週次レポートの自動化") = 0, "sample moved"

    CleanupTests

    Check "no-test-left", CountPrefix("【TEST】") = 0, CStr(CountPrefix("【TEST】"))
    Check "snapshot", Snapshot() = snap, "before=" & snap & " after=" & Snapshot()
    Check "headers-restored", HeaderText() = "起票済み（1）|対応中（1）|対応済み（1）|連携済み（1）", HeaderText()
    Check "done-restored", InStr(1, CStr(SheetDone().Range("A1").Value), "完了した付箋（1）", vbBinaryCompare) = 1, CStr(SheetDone().Range("A1").Value)
    Check "hidden-restored", SheetBoard().Rows(FIRST_ROW + SLOT_HEIGHT).Hidden = True, "not hidden"

    If fails = 0 Then
        report = report & "ALL PASS" & vbCrLf
    Else
        report = report & "ALL FAIL " & CStr(fails) & vbCrLf
    End If
    FinishQuiet
    RunSelfTest = report
    Exit Function
Blow:
    Dim n As Long
    Dim d As String
    n = Err.Number
    d = Err.Description
    On Error Resume Next
    CleanupTests
    FinishQuiet
    RunSelfTest = report & "EXCEPTION " & CStr(n) & " " & d & vbCrLf & "ALL FAIL" & vbCrLf
End Function

Private Sub Check(ByVal name As String, ByVal ok As Boolean, ByVal detail As String)
    If ok Then
        report = report & "PASS " & name & vbCrLf
    Else
        fails = fails + 1
        report = report & "FAIL " & name & " | " & detail & vbCrLf
    End If
End Sub

Private Sub CheckSample(ByVal title As String, ByVal sheetName As String, ByVal lane As Long, ByVal status As String, ByVal body As String, ByVal remark As String, ByVal bodyColor As Long, ByVal titleColor As Long, ByVal stripColor As Long, ByVal fontColor As Long, ByVal archived As Boolean, ByVal prev As String)
    Dim ws As Worksheet
    Dim foundLane As Long
    Dim foundSlot As Long
    Dim n As Long
    n = FindTitle(title, ws, foundLane, foundSlot)
    If n <> 1 Then
        Check "sample-" & title, False, "count=" & CStr(n)
        Exit Sub
    End If
    Dim ok As Boolean
    ok = True
    Dim detail As String
    detail = ""
    If ws.Name <> sheetName Then
        ok = False
        detail = detail & "sheet=" & ws.Name & " "
    End If
    If foundLane <> lane Or foundSlot <> 0 Then
        ok = False
        detail = detail & "pos=" & CStr(foundLane) & "," & CStr(foundSlot) & " "
    End If
    If FieldText(ws, foundLane, foundSlot, OFF_BODY) <> body Then
        ok = False
        detail = detail & "body=" & FieldText(ws, foundLane, foundSlot, OFF_BODY) & " "
    End If
    If FieldText(ws, foundLane, foundSlot, OFF_REMARK) <> remark Then
        ok = False
        detail = detail & "remark=" & FieldText(ws, foundLane, foundSlot, OFF_REMARK) & " "
    End If
    If Field(ws, foundLane, foundSlot, OFF_TITLE).Font.Name <> FONT_NAME Then
        ok = False
        detail = detail & "font=" & Field(ws, foundLane, foundSlot, OFF_TITLE).Font.Name & " "
    End If
    If CLng(Field(ws, foundLane, foundSlot, OFF_BODY).Interior.Color) <> bodyColor Then
        ok = False
        detail = detail & "bodyColor=" & CStr(Field(ws, foundLane, foundSlot, OFF_BODY).Interior.Color) & " "
    End If
    If CLng(Field(ws, foundLane, foundSlot, OFF_TITLE).Interior.Color) <> titleColor Then
        ok = False
        detail = detail & "titleColor=" & CStr(Field(ws, foundLane, foundSlot, OFF_TITLE).Interior.Color) & " "
    End If
    If CLng(Field(ws, foundLane, foundSlot, OFF_TITLE).Interior.Color) <> titleColor Then
        ok = False
    End If
    If CLng(Strip(ws, foundLane, foundSlot).Interior.Color) <> stripColor Then
        ok = False
        detail = detail & "strip=" & CStr(Strip(ws, foundLane, foundSlot).Interior.Color) & " "
    End If
    If CLng(Field(ws, foundLane, foundSlot, OFF_TITLE).Font.Color) <> fontColor Then
        ok = False
        detail = detail & "fontColor=" & CStr(Field(ws, foundLane, foundSlot, OFF_TITLE).Font.Color) & " "
    End If
    If archived Then
        If Left$(StatusText(ws, foundLane, foundSlot), 2) <> "完了" Then
            ok = False
            detail = detail & "display=" & StatusText(ws, foundLane, foundSlot) & " "
        End If
        If Meta(ws, foundLane, foundSlot, OFF_REMARK) <> prev Then
            ok = False
            detail = detail & "prev=" & Meta(ws, foundLane, foundSlot, OFF_REMARK) & " "
        End If
        If Len(Meta(ws, foundLane, foundSlot, OFF_STATUS)) = 0 Then
            ok = False
            detail = detail & "completed-empty "
        End If
        If Field(ws, foundLane, foundSlot, OFF_TITLE).Locked <> True Then
            ok = False
            detail = detail & "not-locked "
        End If
    Else
        If StatusText(ws, foundLane, foundSlot) <> status Then
            ok = False
            detail = detail & "status=" & StatusText(ws, foundLane, foundSlot) & " "
        End If
        If Field(ws, foundLane, foundSlot, OFF_TITLE).Locked <> False Then
            ok = False
            detail = detail & "title-locked "
        End If
        If Not HasList(Field(ws, foundLane, foundSlot, OFF_STATUS)) Then
            ok = False
            detail = detail & "no-validation "
        End If
    End If
    Check "sample-" & Left$(title, 8), ok, detail
End Sub

Private Function CaptionsOk() As Boolean
    Dim ws As Worksheet
    Dim lane As Long
    Dim slot As Long
    If FindTitle("週次レポートの自動化", ws, lane, slot) <> 1 Then
        CaptionsOk = False
        Exit Function
    End If
    CaptionsOk = FieldText(ws, lane, slot, OFF_CAP_TITLE) = "タイトル" _
        And FieldText(ws, lane, slot, OFF_CAP_BODY) = "タスクの内容" _
        And FieldText(ws, lane, slot, OFF_CAP_REMARK) = "備考（関係する実装など）" _
        And FieldText(ws, lane, slot, OFF_CAP_STATUS) = "ステータス"
End Function

Private Function HeaderText() As String
    Dim ws As Worksheet
    Set ws = SheetBoard()
    HeaderText = CStr(ws.Cells(4, 2).Value) & "|" & CStr(ws.Cells(4, 5).Value) & "|" & CStr(ws.Cells(4, 8).Value) & "|" & CStr(ws.Cells(4, 11).Value)
End Function

Private Sub MoveTo(ByVal taskId As String, ByVal status As String)
    Dim ws As Worksheet
    Dim lane As Long
    Dim slot As Long
    If Not FindId(taskId, ws, lane, slot) Then Err.Raise 1001, "MoveTo", "missing " & taskId
    If ws.Name <> "ノート" Then Err.Raise 1002, "MoveTo", taskId & " is on " & ws.Name
    modCore.LastNotice = ""
    Field(ws, lane, slot, OFF_STATUS).Value = status
End Sub

Private Sub SetField(ByVal taskId As String, ByVal offset As Long, ByVal text As String)
    Dim ws As Worksheet
    Dim lane As Long
    Dim slot As Long
    If Not FindId(taskId, ws, lane, slot) Then Err.Raise 1003, "SetField", "missing " & taskId
    Field(ws, lane, slot, offset).Value = text
End Sub

Private Function PlaceIs(ByVal taskId As String, ByVal sheetName As String, ByVal lane As Long, ByVal status As String, ByVal bodyColor As Long, ByVal fontColor As Long) As Boolean
    Dim ws As Worksheet
    Dim foundLane As Long
    Dim foundSlot As Long
    If Not FindId(taskId, ws, foundLane, foundSlot) Then Exit Function
    If ws.Name <> sheetName Then Exit Function
    If foundLane <> lane Then Exit Function
    If CLng(Field(ws, foundLane, foundSlot, OFF_BODY).Interior.Color) <> bodyColor Then Exit Function
    If CLng(Field(ws, foundLane, foundSlot, OFF_TITLE).Font.Color) <> fontColor Then Exit Function
    If Field(ws, foundLane, foundSlot, OFF_TITLE).Font.Name <> FONT_NAME Then Exit Function
    If status = "完了" Then
        PlaceIs = (Left$(StatusText(ws, foundLane, foundSlot), 2) = "完了")
    Else
        PlaceIs = (StatusText(ws, foundLane, foundSlot) = status)
    End If
End Function

Private Function TextIs(ByVal taskId As String, ByVal title As String, ByVal body As String, ByVal remark As String) As Boolean
    Dim ws As Worksheet
    Dim lane As Long
    Dim slot As Long
    If Not FindId(taskId, ws, lane, slot) Then Exit Function
    TextIs = FieldText(ws, lane, slot, OFF_TITLE) = title _
        And FieldText(ws, lane, slot, OFF_BODY) = body _
        And FieldText(ws, lane, slot, OFF_REMARK) = remark
End Function

Private Function WhereText(ByVal taskId As String) As String
    Dim ws As Worksheet
    Dim lane As Long
    Dim slot As Long
    If Not FindId(taskId, ws, lane, slot) Then
        WhereText = "MISSING " & taskId & " notice=" & modCore.LastNotice
        Exit Function
    End If
    WhereText = ws.Name & " lane=" & CStr(lane) & " slot=" & CStr(slot) & " status=" & StatusText(ws, lane, slot) & " notice=" & modCore.LastNotice
End Function

Private Function PrevOf(ByVal taskId As String) As String
    Dim ws As Worksheet
    Dim lane As Long
    Dim slot As Long
    If FindId(taskId, ws, lane, slot) Then PrevOf = Meta(ws, lane, slot, OFF_REMARK)
End Function

Private Function CompletedOf(ByVal taskId As String) As String
    Dim ws As Worksheet
    Dim lane As Long
    Dim slot As Long
    If FindId(taskId, ws, lane, slot) Then CompletedOf = Meta(ws, lane, slot, OFF_STATUS)
End Function

Private Function StatusTextOf(ByVal taskId As String) As String
    Dim ws As Worksheet
    Dim lane As Long
    Dim slot As Long
    If FindId(taskId, ws, lane, slot) Then StatusTextOf = StatusText(ws, lane, slot)
End Function

Private Function LockedTitle(ByVal taskId As String) As Boolean
    Dim ws As Worksheet
    Dim lane As Long
    Dim slot As Long
    If FindId(taskId, ws, lane, slot) Then LockedTitle = (Field(ws, lane, slot, OFF_TITLE).Locked = True)
End Function

Private Function HasListOf(ByVal taskId As String) As Boolean
    Dim ws As Worksheet
    Dim lane As Long
    Dim slot As Long
    If FindId(taskId, ws, lane, slot) Then HasListOf = HasList(Field(ws, lane, slot, OFF_STATUS))
End Function

Private Function OnSheet(ByVal taskId As String, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    Dim lane As Long
    Dim slot As Long
    If Not FindId(taskId, ws, lane, slot) Then Exit Function
    OnSheet = (ws.Name = sheetName)
End Function

Private Function SlotOf(ByVal taskId As String) As Long
    Dim ws As Worksheet
    Dim lane As Long
    Dim slot As Long
    If FindId(taskId, ws, lane, slot) Then SlotOf = slot Else SlotOf = -1
End Function

Private Function LaneOf(ByVal taskId As String) As Long
    Dim ws As Worksheet
    Dim lane As Long
    Dim slot As Long
    If FindId(taskId, ws, lane, slot) Then LaneOf = lane Else LaneOf = -1
End Function

Private Function LaneOfTitle(ByVal title As String) As Long
    Dim ws As Worksheet
    Dim lane As Long
    Dim slot As Long
    If FindTitle(title, ws, lane, slot) = 1 Then LaneOfTitle = lane Else LaneOfTitle = -1
End Function

Private Function SlotOfTitle(ByVal title As String) As Long
    Dim ws As Worksheet
    Dim lane As Long
    Dim slot As Long
    If FindTitle(title, ws, lane, slot) = 1 Then SlotOfTitle = slot Else SlotOfTitle = -1
End Function

Private Function Snapshot() As String
    Snapshot = Sig("週次レポートの自動化") & vbLf _
        & Sig("ログイン失敗の調査") & vbLf _
        & Sig("日付の入力チェック") & vbLf _
        & Sig("APIエラーコードの共有") & vbLf _
        & Sig("開発環境の初期構築")
End Function

Private Function Sig(ByVal title As String) As String
    Dim ws As Worksheet
    Dim lane As Long
    Dim slot As Long
    Dim n As Long
    n = FindTitle(title, ws, lane, slot)
    If n <> 1 Then
        Sig = title & " count=" & CStr(n)
        Exit Function
    End If
    Sig = title & "|" & ws.Name & "|" & CStr(lane) & "|" & CStr(slot) & "|" _
        & FieldText(ws, lane, slot, OFF_TITLE) & "|" _
        & FieldText(ws, lane, slot, OFF_BODY) & "|" _
        & FieldText(ws, lane, slot, OFF_REMARK) & "|" _
        & StatusText(ws, lane, slot) & "|" _
        & Meta(ws, lane, slot, OFF_TITLE) & "|" _
        & Meta(ws, lane, slot, OFF_BODY) & "|" _
        & Meta(ws, lane, slot, OFF_REMARK) & "|" _
        & Meta(ws, lane, slot, OFF_STATUS) & "|" _
        & CStr(Field(ws, lane, slot, OFF_BODY).Interior.Color) & "|" _
        & Field(ws, lane, slot, OFF_TITLE).Font.Name
End Function

Private Sub CleanupTests()
    On Error Resume Next
    Dim guard As Long
    Do
        Dim ws As Worksheet
        Dim lane As Long
        Dim slot As Long
        Dim title As String
        title = FirstPrefix("【TEST】", ws, lane, slot)
        If Len(title) = 0 Then Exit Do
        modCore.DeleteAt ws.Name, lane, slot
        guard = guard + 1
    Loop While guard < 30
End Sub

Private Function FirstPrefix(ByVal prefix As String, ByRef wsOut As Worksheet, ByRef laneOut As Long, ByRef slotOut As Long) As String
    Dim names(1 To 2) As String
    names(1) = "ノート"
    names(2) = "完了"
    Dim i As Long
    For i = 1 To 2
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(names(i))
        Dim maxLane As Long
        If names(i) = "完了" Then maxLane = 2 Else maxLane = 3
        Dim lane As Long
        Dim slot As Long
        For lane = 0 To maxLane
            For slot = 0 To 24
                Dim text As String
                text = FieldText(ws, lane, slot, OFF_TITLE)
                If Left$(text, Len(prefix)) = prefix Then
                    Set wsOut = ws
                    laneOut = lane
                    slotOut = slot
                    FirstPrefix = text
                    Exit Function
                End If
            Next slot
        Next lane
    Next i
End Function

Private Function CountPrefix(ByVal prefix As String) As Long
    CountPrefix = CountMatch(prefix, True)
End Function

Private Function CountTitle(ByVal title As String) As Long
    CountTitle = CountMatch(title, False)
End Function

Private Function CountMatch(ByVal text As String, ByVal byPrefix As Boolean) As Long
    Dim n As Long
    Dim names(1 To 2) As String
    names(1) = "ノート"
    names(2) = "完了"
    Dim i As Long
    For i = 1 To 2
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(names(i))
        Dim maxLane As Long
        If names(i) = "完了" Then maxLane = 2 Else maxLane = 3
        Dim lane As Long
        Dim slot As Long
        For lane = 0 To maxLane
            For slot = 0 To 24
                Dim value As String
                value = FieldText(ws, lane, slot, OFF_TITLE)
                If byPrefix Then
                    If Left$(value, Len(text)) = text And Len(value) > 0 Then n = n + 1
                Else
                    If value = text Then n = n + 1
                End If
            Next slot
        Next lane
    Next i
    CountMatch = n
End Function

Private Function FindTitle(ByVal title As String, ByRef wsOut As Worksheet, ByRef laneOut As Long, ByRef slotOut As Long) As Long
    Dim n As Long
    Dim names(1 To 2) As String
    names(1) = "ノート"
    names(2) = "完了"
    Dim i As Long
    For i = 1 To 2
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(names(i))
        Dim maxLane As Long
        If names(i) = "完了" Then maxLane = 2 Else maxLane = 3
        Dim lane As Long
        Dim slot As Long
        For lane = 0 To maxLane
            For slot = 0 To 24
                If FieldText(ws, lane, slot, OFF_TITLE) = title Then
                    n = n + 1
                    Set wsOut = ws
                    laneOut = lane
                    slotOut = slot
                End If
            Next slot
        Next lane
    Next i
    FindTitle = n
End Function

Private Function FindId(ByVal taskId As String, ByRef wsOut As Worksheet, ByRef laneOut As Long, ByRef slotOut As Long) As Boolean
    Dim names(1 To 2) As String
    names(1) = "ノート"
    names(2) = "完了"
    Dim i As Long
    For i = 1 To 2
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(names(i))
        Dim maxLane As Long
        If names(i) = "完了" Then maxLane = 2 Else maxLane = 3
        Dim lane As Long
        Dim slot As Long
        For lane = 0 To maxLane
            For slot = 0 To 24
                If Meta(ws, lane, slot, OFF_TITLE) = taskId And Len(taskId) > 0 Then
                    Set wsOut = ws
                    laneOut = lane
                    slotOut = slot
                    FindId = True
                    Exit Function
                End If
            Next slot
        Next lane
    Next i
End Function

Private Function Field(ByVal ws As Worksheet, ByVal lane As Long, ByVal slot As Long, ByVal offset As Long) As Range
    Set Field = ws.Cells(FIRST_ROW + slot * SLOT_HEIGHT + offset, 3 + lane * 3)
End Function

Private Function Strip(ByVal ws As Worksheet, ByVal lane As Long, ByVal slot As Long) As Range
    Set Strip = ws.Cells(FIRST_ROW + slot * SLOT_HEIGHT + OFF_TITLE, 2 + lane * 3)
End Function

Private Function FieldText(ByVal ws As Worksheet, ByVal lane As Long, ByVal slot As Long, ByVal offset As Long) As String
    FieldText = CellText(Field(ws, lane, slot, offset))
End Function

Private Function StatusText(ByVal ws As Worksheet, ByVal lane As Long, ByVal slot As Long) As String
    StatusText = FieldText(ws, lane, slot, OFF_STATUS)
End Function

Private Function Meta(ByVal ws As Worksheet, ByVal lane As Long, ByVal slot As Long, ByVal offset As Long) As String
    Meta = CellText(ws.Cells(FIRST_ROW + slot * SLOT_HEIGHT + offset, 2 + lane * 3))
End Function

Private Function CellText(ByVal cell As Range) As String
    If IsEmpty(cell.Value) Then
        CellText = ""
    Else
        CellText = CStr(cell.Value)
    End If
End Function

Private Function HasList(ByVal cell As Range) As Boolean
    On Error Resume Next
    Dim t As Long
    t = 0
    t = cell.Validation.Type
    If Err.Number = 0 Then
        HasList = (t = 3)
    Else
        HasList = False
    End If
    Err.Clear
End Function

Private Function SheetBoard() As Worksheet
    Set SheetBoard = ThisWorkbook.Worksheets("ノート")
End Function

Private Function SheetDone() As Worksheet
    Set SheetDone = ThisWorkbook.Worksheets("完了")
End Function

Private Function SheetHelp() As Worksheet
    Set SheetHelp = ThisWorkbook.Worksheets("使い方")
End Function

Private Function SafeName(ByVal ws As Worksheet) As String
    If ws Is Nothing Then
        SafeName = "(none)"
    Else
        SafeName = ws.Name
    End If
End Function

Private Sub FinishQuiet()
    On Error Resume Next
    modCore.SuppressPrompts = False
    modCore.ResetRuntimeState
    SheetBoard().Activate
    Application.Goto SheetBoard().Range("A1"), True
End Sub

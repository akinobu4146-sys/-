# セル値貼り付けとシート名変更マクロ

## 取り込み方
1. Excel で Alt+F11（VBE を開く）
2. 「ファイル」→「ファイルのインポート」で `セル値貼り付けとシート名変更.bas` を選択（Shift_JIS 形式）
   - または下のコードを標準モジュールに貼り付け（1行目の `Attribute VB_Name` 行は除く）
3. Alt+F8 で `セル値貼り付けとシート名変更` を実行

## コード

```vba
Attribute VB_Name = "PasteAndRenameSheet"
Option Explicit

'==========================================================================
' セルの値を別のセルへ値貼り付けし、さらにシート名にもするマクロ
'
'  1) 第一：貼り付け先のセルをマウスで指定
'  2) 第二：値を取り出すセルをマウスで指定
'  3) 「貼り付け／シート名変更／両方／中止」から選択して実行
'  4) 「右側シートでも行う／左側シートでも行う／中止」を選択
'     左右を選ぶと、そのシートで第二のセルと同じ位置に
'     マウスポインタを移動し、1) から繰り返す
'==========================================================================

#If VBA7 Then
    Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
#Else
    Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
#End If

Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Private Const MODE_CANCEL As Long = 0
Private Const MODE_PASTE As Long = 1
Private Const MODE_RENAME As Long = 2
Private Const MODE_BOTH As Long = 3

Private Const DIR_CANCEL As Long = 0
Private Const DIR_RIGHT As Long = 1
Private Const DIR_LEFT As Long = 2

'--------------------------------------------------------------------------
' メイン（このマクロを実行してください）
'--------------------------------------------------------------------------
Public Sub セル値貼り付けとシート名変更()
    Dim destCell As Range       ' 第一：貼り付け先
    Dim srcCell As Range        ' 第二：値の取り出し元
    Dim ws As Worksheet         ' 処理対象シート
    Dim nextWs As Worksheet
    Dim destAddr As String      ' 左右シートで使う前回の位置
    Dim srcAddr As String
    Dim mode As Long
    Dim dirChoice As Long

    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "ワークシートを表示してから実行してください。", vbExclamation
        Exit Sub
    End If

    Do
        '--- 第一：貼り付け先セル -----------------------------------------
        Set destCell = AskCell( _
            "【第一】値を貼り付ける先のセルをマウスでクリックして指定し、OK を押してください。", _
            "第一：貼り付け先セルの指定", destAddr)
        If destCell Is Nothing Then Exit Sub
        Set ws = destCell.Worksheet

        '--- 第二：値の取り出し元セル -------------------------------------
        If Len(srcAddr) > 0 Then MoveMouseToCell ws.Range(srcAddr)
        Set srcCell = AskCell( _
            "【第二】貼り付けたい文字列（シート名にしたい文字列）が入っているセルを" & vbLf & _
            "マウスでクリックして指定し、OK を押してください。", _
            "第二：値のセルの指定", srcAddr)
        If srcCell Is Nothing Then Exit Sub

        '--- 実行内容の選択 -----------------------------------------------
        mode = AskMode(destCell, srcCell)
        If mode = MODE_CANCEL Then Exit Sub

        If Not Execute(mode, ws, destCell, srcCell) Then Exit Sub

        '--- 左右シートでも行うか -----------------------------------------
        dirChoice = AskDirection(ws)
        If dirChoice = DIR_CANCEL Then Exit Sub

        Set nextWs = NeighborSheet(ws, dirChoice)
        If nextWs Is Nothing Then
            MsgBox IIf(dirChoice = DIR_RIGHT, "右側", "左側") & _
                   "に表示されているワークシートがありません。終了します。", vbInformation
            Exit Sub
        End If

        ' 次のシートでは同じ位置を既定値として使う
        destAddr = destCell.Cells(1, 1).Address
        srcAddr = srcCell.Cells(1, 1).Address

        nextWs.Activate
        ' 直ちに第二で指定するはずのセル位置へマウスポインタを置く
        MoveMouseToCell nextWs.Range(srcAddr)
    Loop
End Sub

'--------------------------------------------------------------------------
' 選択内容を実行する。成功したら True
'--------------------------------------------------------------------------
Private Function Execute(ByVal mode As Long, ByVal ws As Worksheet, _
                         ByVal destCell As Range, ByVal srcCell As Range) As Boolean
    Dim v As Variant
    Dim newName As String
    Dim errMsg As String

    v = srcCell.Cells(1, 1).Value
    If IsError(v) Then
        MsgBox "第二で指定したセル（" & srcCell.Cells(1, 1).Address(False, False) & _
               "）はエラー値のため処理できません。", vbExclamation
        Exit Function
    End If

    ' 貼り付け（値のみ）
    If mode = MODE_PASTE Or mode = MODE_BOTH Then
        On Error GoTo PasteErr
        destCell.Cells(1, 1).Value = v
        On Error GoTo 0
    End If

    ' シート名変更
    If mode = MODE_RENAME Or mode = MODE_BOTH Then
        newName = CStr(v)
        errMsg = CheckSheetName(newName, ws)
        If Len(errMsg) > 0 Then
            MsgBox "シート名を「" & newName & "」に変更できません。" & vbLf & errMsg, vbExclamation
            Exit Function
        End If
        If ws.Name <> newName Then
            On Error GoTo RenameErr
            ws.Name = newName
            On Error GoTo 0
        End If
    End If

    Execute = True
    Exit Function

PasteErr:
    MsgBox "貼り付けに失敗しました。" & vbLf & Err.Description, vbExclamation
    Exit Function
RenameErr:
    MsgBox "シート名の変更に失敗しました。" & vbLf & Err.Description, vbExclamation
End Function

'--------------------------------------------------------------------------
' マウスでセルを指定させる。キャンセル時は Nothing
'--------------------------------------------------------------------------
Private Function AskCell(ByVal prompt As String, ByVal title As String, _
                         ByVal defAddr As String) As Range
    Dim r As Range
    On Error Resume Next
    If Len(defAddr) > 0 Then
        Set r = Application.InputBox(prompt, title, defAddr, Type:=8)
    Else
        Set r = Application.InputBox(prompt, title, Type:=8)
    End If
    On Error GoTo 0
    If Not r Is Nothing Then Set AskCell = r.Cells(1, 1)
End Function

'--------------------------------------------------------------------------
' 実行内容を選ばせる
'--------------------------------------------------------------------------
Private Function AskMode(ByVal destCell As Range, ByVal srcCell As Range) As Long
    Dim msg As String
    Dim ans As Variant

    msg = "第一（貼り付け先）： " & CellLabel(destCell) & vbLf & _
          "第二（値のセル）　： " & CellLabel(srcCell) & vbLf & _
          "第二の値　　　　　： " & ShortText(srcCell.Cells(1, 1).Text) & vbLf & vbLf & _
          "実行する内容を番号で入力してください。" & vbLf & _
          "  1 : 第一で指定したセルに第二で指定したセルの値を貼り付ける" & vbLf & _
          "  2 : 第二で指定したセルの値をシート名にする" & vbLf & _
          "  3 : 両方やる" & vbLf & _
          "  0 : 中止する"

    Do
        ans = Application.InputBox(msg, "実行内容の選択", 3, Type:=1)
        If VarType(ans) = vbBoolean Then AskMode = MODE_CANCEL: Exit Function
        Select Case ans
            Case MODE_CANCEL, MODE_PASTE, MODE_RENAME, MODE_BOTH
                AskMode = CLng(ans)
                Exit Function
        End Select
        MsgBox "0〜3 の番号を入力してください。", vbExclamation
    Loop
End Function

'--------------------------------------------------------------------------
' 左右どちらのシートで続けるか選ばせる
'--------------------------------------------------------------------------
Private Function AskDirection(ByVal ws As Worksheet) As Long
    Dim msg As String
    Dim ans As Variant
    Dim r As Worksheet, l As Worksheet

    Set r = NeighborSheet(ws, DIR_RIGHT)
    Set l = NeighborSheet(ws, DIR_LEFT)

    msg = "処理が終わりました。続けて行いますか？" & vbLf & vbLf & _
          "  1 : 右側シートでも行う（" & SheetLabel(r) & "）" & vbLf & _
          "  2 : 左側シートでも行う（" & SheetLabel(l) & "）" & vbLf & _
          "  0 : 中止する"

    Do
        ans = Application.InputBox(msg, "続けて行うシートの選択", 1, Type:=1)
        If VarType(ans) = vbBoolean Then AskDirection = DIR_CANCEL: Exit Function
        Select Case ans
            Case DIR_CANCEL, DIR_RIGHT, DIR_LEFT
                AskDirection = CLng(ans)
                Exit Function
        End Select
        MsgBox "0〜2 の番号を入力してください。", vbExclamation
    Loop
End Function

'--------------------------------------------------------------------------
' 左右の（表示されている）ワークシートを返す。無ければ Nothing
'--------------------------------------------------------------------------
Private Function NeighborSheet(ByVal ws As Worksheet, ByVal dirChoice As Long) As Worksheet
    Dim i As Long, stepVal As Long
    Dim sh As Object

    stepVal = IIf(dirChoice = DIR_RIGHT, 1, -1)
    i = ws.Index + stepVal
    Do While i >= 1 And i <= ws.Parent.Sheets.Count
        Set sh = ws.Parent.Sheets(i)
        If TypeName(sh) = "Worksheet" Then
            If sh.Visible = xlSheetVisible Then
                Set NeighborSheet = sh
                Exit Function
            End If
        End If
        i = i + stepVal
    Loop
End Function

'--------------------------------------------------------------------------
' シート名として使えるか確認。問題があれば理由を返す
'--------------------------------------------------------------------------
Private Function CheckSheetName(ByVal nm As String, ByVal ws As Worksheet) As String
    Dim bad As Variant
    Dim sh As Object

    If Len(nm) = 0 Then CheckSheetName = "値が空です。": Exit Function
    If Len(nm) > 31 Then CheckSheetName = "31文字を超えています（" & Len(nm) & "文字）。": Exit Function
    For Each bad In Array(":", "\", "/", "?", "*", "[", "]", vbCr, vbLf, vbTab)
        If InStr(nm, bad) > 0 Then
            CheckSheetName = "使えない文字 （ : \ / ? * [ ] 改行 など）が含まれています。"
            Exit Function
        End If
    Next
    If Left$(nm, 1) = "'" Or Right$(nm, 1) = "'" Then
        CheckSheetName = "先頭や末尾に ' は使えません。": Exit Function
    End If
    If StrComp(nm, "History", vbTextCompare) = 0 Then
        CheckSheetName = "「History」は予約語のため使えません。": Exit Function
    End If
    For Each sh In ws.Parent.Sheets
        If Not sh Is ws Then
            If StrComp(sh.Name, nm, vbTextCompare) = 0 Then
                CheckSheetName = "同じ名前のシートが既にあります。": Exit Function
            End If
        End If
    Next
End Function

'--------------------------------------------------------------------------
' 指定セルの中央にマウスポインタを移動する
'--------------------------------------------------------------------------
Private Sub MoveMouseToCell(ByVal rng As Range)
    Dim wnd As Window
    Dim pn As Pane, p As Pane
    Dim area As Range
    Dim cx As Double, cy As Double
    Dim x As Long, y As Long
    Dim dpiX As Long, dpiY As Long
    Dim hit As Object

    On Error GoTo Done
    Set area = rng.Cells(1, 1).MergeArea
    Set wnd = ActiveWindow
    If Not wnd.ActiveSheet Is rng.Worksheet Then Exit Sub

    ' セルが画面に見えるようにスクロール
    wnd.ScrollIntoView area.Left, area.Top, area.Width, area.Height
    DoEvents

    ' セルが表示されているペインを探す（ウィンドウ枠の固定対策）
    For Each p In wnd.Panes
        If Not Intersect(p.VisibleRange, area.Cells(1, 1)) Is Nothing Then
            Set pn = p
            Exit For
        End If
    Next
    If pn Is Nothing Then Set pn = wnd.ActivePane

    cx = area.Left + area.Width / 2
    cy = area.Top + area.Height / 2

    ' 方法1：ペインの座標変換
    x = pn.PointsToScreenPixelsX(cx)
    y = pn.PointsToScreenPixelsY(cy)

    ' 狙ったセルに当たっていなければ 方法2：倍率と DPI から計算
    Set hit = Nothing
    On Error Resume Next
    Set hit = wnd.RangeFromPoint(x, y)
    On Error GoTo Done
    If Not IsTargetCell(hit, area) Then
        GetDpi dpiX, dpiY
        x = wnd.PointsToScreenPixelsX(0) + _
            CLng((cx - pn.VisibleRange.Left) * wnd.Zoom / 100 * dpiX / 72)
        y = wnd.PointsToScreenPixelsY(0) + _
            CLng((cy - pn.VisibleRange.Top) * wnd.Zoom / 100 * dpiY / 72)
    End If

    SetCursorPos x, y
Done:
End Sub

Private Function IsTargetCell(ByVal hit As Object, ByVal area As Range) As Boolean
    If hit Is Nothing Then Exit Function
    If TypeName(hit) <> "Range" Then Exit Function
    If Not hit.Worksheet Is area.Worksheet Then Exit Function
    IsTargetCell = Not Intersect(hit, area) Is Nothing
End Function

Private Sub GetDpi(ByRef dpiX As Long, ByRef dpiY As Long)
#If VBA7 Then
    Dim hdc As LongPtr
#Else
    Dim hdc As Long
#End If
    hdc = GetDC(0)
    dpiX = GetDeviceCaps(hdc, LOGPIXELSX)
    dpiY = GetDeviceCaps(hdc, LOGPIXELSY)
    ReleaseDC 0, hdc
    If dpiX <= 0 Then dpiX = 96
    If dpiY <= 0 Then dpiY = 96
End Sub

'--------------------------------------------------------------------------
' 表示用
'--------------------------------------------------------------------------
Private Function CellLabel(ByVal r As Range) As String
    CellLabel = r.Worksheet.Name & " ! " & r.Cells(1, 1).Address(False, False)
End Function

Private Function SheetLabel(ByVal sh As Worksheet) As String
    If sh Is Nothing Then
        SheetLabel = "シートなし"
    Else
        SheetLabel = sh.Name
    End If
End Function

Private Function ShortText(ByVal s As String) As String
    If Len(s) = 0 Then
        ShortText = "（空）"
    ElseIf Len(s) > 40 Then
        ShortText = Left$(s, 40) & "…"
    Else
        ShortText = s
    End If
End Function
```

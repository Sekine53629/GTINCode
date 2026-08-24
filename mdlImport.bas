Option Explicit

REM ==========================================================
REM  G列のフルパスからヒート画像を一括配置
REM  データ入力後に実行。何度でも再実行可。
REM ----------------------------------------------------------
REM  画像の格納そのものは mdlPhoto の AssignHeatImage が担当し、
REM  ここは配置だけを行う。
REM ==========================================================
Private Const SHAPE_PREFIX As String = "HEAT_"
Private Const IMG_TOP_OFS  As Long = 2   ' 画像枠の上端オフセット
Private Const IMG_BTM_OFS  As Long = 15  ' 画像枠の下端オフセット
Private Const IMG_LFT_OFS  As Long = 3   ' 画像枠の左端オフセット
Private Const IMG_RGT_OFS  As Long = 4   ' 画像枠の右端オフセット
Private Const AUTO_SIZE    As Long = -1  ' AddPicture に元サイズを使わせる指定

Private Const MSG_PLACED  As String = "配置_ "
Private Const MSG_MISSING As String = "見つからないファイル_ "
Private Const MSG_UNIT    As String = " 件"

REM ==========================================================
REM  画像を配置して結果を通知する
REM ==========================================================
Public Sub ImportHeatImages()
    Dim okN As Long, ngN As Long, missing As String
    Dim t As String

    PlaceHeatImages okN, ngN, missing

    t = MSG_PLACED & okN & MSG_UNIT
    If ngN > 0 Then t = t & vbCrLf & MSG_MISSING & ngN & MSG_UNIT & missing
    MsgBox t, IIf(ngN > 0, vbExclamation, vbInformation)
End Sub

REM ==========================================================
REM  画像の配置本体
REM  通知を出さないので、テストからも直接呼べる。
REM  okN 配置できた件数 / ngN ファイルが無かった件数
REM  missing 見つからなかったパスの一覧
REM ==========================================================
Public Sub PlaceHeatImages(ByRef okN As Long, ByRef ngN As Long, ByRef missing As String)
    Dim wsD As Worksheet, wsC As Worksheet
    Dim i As Long, cc As Long, rr As Long, c0 As Long, r0 As Long
    Dim pth As String, sh As Shape, tgt As Range
    Dim scrSave As Boolean

    okN = 0: ngN = 0: missing = ""
    Set wsD = ThisWorkbook.Worksheets(SHT_DATA)
    Set wsC = ThisWorkbook.Worksheets(SHT_CARDS)
    scrSave = Application.ScreenUpdating
    Application.ScreenUpdating = False

    REM 既存の取込画像を削除（名前で判別）
    For i = wsC.Shapes.Count To 1 Step -1
        If Left$(wsC.Shapes(i).Name, Len(SHAPE_PREFIX)) = SHAPE_PREFIX Then wsC.Shapes(i).Delete
    Next i

    For i = 0 To N_ITEMS - 1
        pth = CStr(wsD.Cells(DATA_ROW1 + i, COL_FULLPATH).Value)   ' G列 = フルパス
        If Len(pth) > 0 Then
            If Len(Dir(pth)) > 0 Then
                cc = i Mod COLS_N
                rr = i \ COLS_N
                c0 = START_C + cc * CARD_W
                r0 = START_R + rr * (CARD_H + GAP_R)
                REM 画像枠の範囲
                Set tgt = wsC.Range(wsC.Cells(r0 + IMG_TOP_OFS, c0 + IMG_LFT_OFS), _
                                    wsC.Cells(r0 + IMG_BTM_OFS, c0 + CARD_W - IMG_RGT_OFS))
                Set sh = wsC.Shapes.AddPicture(pth, msoFalse, msoTrue, _
                             tgt.Left, tgt.Top, AUTO_SIZE, AUTO_SIZE)
                With sh
                    .Name = SHAPE_PREFIX & Format(i + 1, "00")
                    .LockAspectRatio = msoTrue
                    REM 枠に収まるよう縮小
                    If .Width / .Height > tgt.Width / tgt.Height Then
                        .Width = tgt.Width
                    Else
                        .Height = tgt.Height
                    End If
                    REM 中央寄せ
                    .Left = tgt.Left + (tgt.Width - .Width) / 2
                    .Top = tgt.Top + (tgt.Height - .Height) / 2
                    .Placement = xlMoveAndSize
                End With
                okN = okN + 1
            Else
                ngN = ngN + 1
                missing = missing & vbCrLf & "  No." & (i + 1) & "_ " & pth
            End If
        End If
    Next i

    Application.ScreenUpdating = scrSave
End Sub

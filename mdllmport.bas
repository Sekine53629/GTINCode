Option Explicit


REM ==========================================================
  REM おまけ: L列のパスからヒート画像を一括配置
  REM （データ入力後に実行。何度でも再実行可）
REM ==========================================================
Public Sub ImportHeatImages()
    Dim wsD As Worksheet, wsC As Worksheet
    Dim i As Long, cc As Long, rr As Long, c0 As Long, r0 As Long
    Dim pth As String, sh As Shape, tgt As Range
    Dim okN As Long, ngN As Long, msg As String

    Set wsD = ThisWorkbook.Worksheets("データ")
    Set wsC = ThisWorkbook.Worksheets("カード印刷")
    Application.ScreenUpdating = False

    REM 既存の取込画像を削除（名前で判別）
    For i = wsC.Shapes.Count To 1 Step -1
        If Left(wsC.Shapes(i).Name, 5) = "HEAT_" Then wsC.Shapes(i).Delete
    Next i

    For i = 0 To N_ITEMS - 1
        pth = CStr(wsD.Cells(5 + i, 12).Value)   ' L列 = フルパス
        If Len(pth) > 0 Then
            If Dir(pth) <> "" Then
                cc = i Mod COLS_N
                rr = i \ COLS_N
                c0 = START_C + cc * CARD_W
                r0 = START_R + rr * (CARD_H + GAP_R)
                REM 画像枠の範囲
                Set tgt = wsC.Range(wsC.Cells(r0 + 2, c0 + 3), _
                                    wsC.Cells(r0 + 15, c0 + CARD_W - 4))
                Set sh = wsC.Shapes.AddPicture(pth, msoFalse, msoTrue, _
                             tgt.Left, tgt.Top, -1, -1)
                With sh
                    .Name = "HEAT_" & Format(i + 1, "00")
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
                msg = msg & vbCrLf & "  No." & (i + 1) & ": " & pth
            End If
        End If
    Next i

    Application.ScreenUpdating = True
    Dim t As String
    t = "配置: " & okN & " 件"
    If ngN > 0 Then t = t & vbCrLf & "見つからないファイル: " & ngN & " 件" & msg
    MsgBox t, IIf(ngN > 0, vbExclamation, vbInformation)
End Sub
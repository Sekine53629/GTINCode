REM =========================================================
  REM 調剤在庫マスタ / ヒートGTIN カード印刷ブック  自動生成マクロ
  REM 白紙のブックで実行すると 4シートを一括生成します
REM ---------------------------------------------------------
  REM 使い方:
  REM  1. Alt+F11 で VBE を開く
  REM  2. 挿入 > 標準モジュール
  REM  3. 以下を全て貼り付け
  REM  4. F5 または Alt+F8 から BuildWorkbook を実行
REM =========================================================
Option Explicit

REM --- 定数（レイアウト定義）------------------------------------
Public Const N_ITEMS As Long = 20      ' 品目数
Public Const CARD_W  As Long = 129     ' カード幅（列数）= 45.5mm
Public Const CARD_H  As Long = 26      ' カード高さ（行数）
Public Const COLS_N  As Long = 4       ' 横に並べる枚数
Public Const GAP_R   As Long = 4       ' 段間の行数
Public Const START_C As Long = 2       ' 開始列 (B)
Public Const START_R As Long = 2       ' 開始行
REM バーコードの左オフセットは mdlBarcode の DBL_MODULES から実行時に求める

REM --- 定数（シート名）------------------------------------------
Public Const SHT_DATA    As String = "データ"
Public Const SHT_NAMING  As String = "命名規則"
Public Const SHT_ENCODER As String = "バーコード計算"
Public Const SHT_CARDS   As String = "カード印刷"

REM --- 定数（データシートの設定セル）----------------------------
Public Const CELL_IMG_DIR    As String = "J1"   ' 画像フォルダ
Public Const CELL_IMG_EXT    As String = "J2"   ' 拡張子
Public Const CELL_MASTER_URL As String = "J3"   ' 医薬品マスタの取込元URL
Private Const REF_IMG_DIR As String = "$J$1"    ' 数式から参照する絶対番地
Private Const REF_IMG_EXT As String = "$J$2"

Private Const DEFAULT_IMG_DIR As String = "C:\薬局\ヒート画像\"
Public  Const PATH_SEP        As String = "\"
Public  Const MAX_PATH_LEN    As Long = 259   ' Windows の従来パス長の上限
Private Const DEFAULT_IMG_EXT As String = ".jpg"

Private Const FMT_GENERAL   As String = "General"
Private Const FMT_TEXT      As String = "@"

Private Const MSG_REPAIRED  As String = "データシートの数式を貼り直しました。"
Private Const MSG_REPAIR_NG As String = "数式の貼り直しに失敗しました。"

REM --- 定数（データシートの列位置。1始まり）---------------------
Public Const COL_NO       As Long = 1
Public Const COL_DRUGNAME As Long = 2
Public Const COL_GTIN14   As Long = 3
Public Const COL_CHECK    As Long = 4
Public Const COL_CODE13   As Long = 5   ' 調剤包装単位コード13桁。マスタ照合キー
Public Const COL_FILENAME As Long = 6
Public Const COL_FULLPATH As Long = 7
Public Const DATA_ROW1    As Long = 5   ' 明細の開始行

REM --- 定数（A4 印刷フィット）--------------------------------
Private Const MM_PER_INCH   As Double = 25.4
Private Const PT_PER_MM     As Double = 72# / 25.4  ' 1mm あたりのポイント数
Private Const MODULE_PT     As Double = 1#          ' バーコード 1モジュール幅 = 列幅 1pt
Private Const NOMINAL_MODULE_MM As Double = 0.33    ' GS1 標準モジュール幅 倍率100%
Private Const MIN_MAGNIF    As Double = 0.8         ' GS1 が認める最小倍率 80%
Private Const MIN_ZOOM      As Long = 60            ' 縮小印刷の下限
Private Const ZOOM_TRIES    As Long = 12            ' 縮小率の追い込み回数
Private Const MSG_A4_ZOOM   As String = "A4 1ページに収めるため印刷倍率を下げました。"
Private Const MSG_A4_NG     As String = "A4 1ページに収まりませんでした。プリンタの紙サイズと余白設定を確認してください。"
Private Const MSG_MODULE    As String = "バーコード 1モジュール幅_ "
Private Const MSG_MAGNIF_NG As String = "GS1 の最小倍率 80% を下回っています。読み取り精度が落ちる可能性があります。"

Public Sub BuildWorkbook()
    Dim calcSave As XlCalculation
    Dim masterRows As Long
    calcSave = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    On Error GoTo Fail
    REM 医薬品マスタを先に作る。データシートの数式が tblDrugMaster を
    REM 参照するため、テーブルが存在しないと数式の書き込み自体が失敗する。
    RemoveOldSheets
    masterRows = BuildDrugMaster(DEFAULT_MASTER_URL)
    BuildData
    BuildNaming
    BuildEncoder
    BuildCards

    ArrangeSheets
    ThisWorkbook.Worksheets(SHT_DATA).Activate

    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "生成完了。" & vbCrLf & _
           "医薬品マスタ_ " & Format(masterRows, "#,##0") & " 件" & vbCrLf & _
           "データシートのC列にヒートのGTIN14桁を入力してください。薬剤名は自動で入ります。", vbInformation
    Exit Sub
Fail:
    Application.Calculation = calcSave
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "エラー " & Err.Number & ": " & Err.Description, vbCritical
End Sub

REM --- 既存シート削除 ------------------------------------------
Private Sub RemoveOldSheets()
    Dim nm As Variant, ws As Worksheet
    For Each nm In Array(SHT_DATA, SHT_NAMING, SHT_ENCODER, SHT_CARDS, MASTER_SHEET)
        On Error Resume Next
        Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(CStr(nm))
        On Error GoTo 0
        If Not ws Is Nothing Then
            If ThisWorkbook.Worksheets.Count = 1 Then ThisWorkbook.Worksheets.Add
            ws.Delete
        End If
    Next nm
End Sub

REM --- シート順を確定させる --------------------------------
REM  NewSheet の pos 指定だけだと、元の空シートや医薬品マスタの位置に
REM  左右されて順番が安定しないので、最後に一括で並べ直す。
Private Sub ArrangeSheets()
    Dim nm As Variant, i As Long
    i = 0
    For Each nm In Array(SHT_DATA, SHT_NAMING, SHT_ENCODER, SHT_CARDS, MASTER_SHEET)
        i = i + 1
        If i = 1 Then
            ThisWorkbook.Worksheets(CStr(nm)).Move Before:=ThisWorkbook.Worksheets(1)
        Else
            ThisWorkbook.Worksheets(CStr(nm)).Move After:=ThisWorkbook.Worksheets(i - 1)
        End If
    Next nm
End Sub

REM --- 新規シート取得（無ければ作る）---------------------------
Private Function NewSheet(ByVal nm As String, ByVal pos As Long) As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = nm
    If pos > 0 And pos <= ThisWorkbook.Worksheets.Count Then
        ws.Move After:=ThisWorkbook.Worksheets(pos)
    End If
    Set NewSheet = ws
End Function

REM ==========================================================
  REM 1. データシート
REM ==========================================================
Private Sub BuildData()
    Dim ws As Worksheet, i As Long
    Dim lastRow As Long
    Set ws = NewSheet(SHT_DATA, 0)
    ws.Move Before:=ThisWorkbook.Worksheets(1)
    lastRow = DATA_ROW1 + N_ITEMS - 1

    With ws.Range("A1")
        .Value = "調剤在庫マスタ（ヒートGTIN カード印刷用）"
        .Font.Size = 14: .Font.Bold = True
    End With
    With ws.Range("A2")
        .Value = "※ ヒートの (01) の後ろの14桁をC列に入力してください。薬剤名とファイル名は医薬品マスタから自動生成されます。"
        .Font.Size = 9: .Font.Color = RGB(102, 102, 102)
    End With

    REM --- 設定欄 ---
    ws.Range("I1").Value = "画像フォルダ"
    ws.Range("I2").Value = "拡張子"
    ws.Range("I3").Value = "医薬品マスタ URL"
    With ws.Range("I1:I3")
        .Font.Bold = True: .Font.Size = 10
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    ws.Range(CELL_IMG_DIR).Value = DEFAULT_IMG_DIR
    ws.Range(CELL_IMG_EXT).Value = DEFAULT_IMG_EXT
    ws.Range(CELL_MASTER_URL).Value = DEFAULT_MASTER_URL
    With ws.Range("J1:J3")
        .Font.Color = RGB(0, 0, 255): .Font.Name = "Consolas": .Font.Size = 9
        .Interior.Color = RGB(255, 249, 224)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 143, 0): .Borders.Weight = xlMedium
        .HorizontalAlignment = xlLeft
    End With
    With ws.Range("I4")
        .Value = "※ URLを変えたら RefreshDrugMaster を実行"
        .Font.Size = 9: .Font.Color = RGB(102, 102, 102)
    End With

    REM --- ヘッダ ---
    Dim hdr As Variant
    hdr = Array("No.", "薬剤名（自動取得）", "GTIN-14（ヒート入力）", "検算", _
                "13桁（マスタ照合用）", "ファイル名（自動生成）", "フルパス（自動生成）")
    For i = 0 To UBound(hdr)
        ws.Cells(4, i + 1).Value = hdr(i)
    Next i
    With ws.Range("A4:G4")
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    REM --- 連番のみ。薬剤名は医薬品マスタから引くので手入力しない ---
    For i = 0 To N_ITEMS - 1
        ws.Cells(DATA_ROW1 + i, COL_NO).Value = i + 1
    Next i

    WriteDataFormulas ws

    REM --- 書式: B列（自動取得）---
    With ws.Range("B5:B" & lastRow)
        .Interior.Color = RGB(242, 242, 242)
        .HorizontalAlignment = xlLeft
        .WrapText = False
    End With
    REM マスタ未収載を警告色
    With ws.Range("B5:B" & lastRow).FormatConditions
        .Delete
        With .Add(Type:=xlTextString, String:="該当なし", TextOperator:=xlContains)
            .Interior.Color = RGB(255, 199, 206)
            .Font.Color = RGB(156, 0, 6)
        End With
    End With

    REM --- 書式: C列（GTIN入力欄）---
    With ws.Range("C5:C" & lastRow)
        .NumberFormat = FMT_TEXT
        .Font.Color = RGB(0, 0, 255)
        .Font.Name = "Consolas"
        .HorizontalAlignment = xlCenter
    End With
    REM 未入力を黄色く（入力済みは白）
    With ws.Range("C5:C" & lastRow).FormatConditions
        .Delete
        .Add(xlCellValue, xlEqual, "=""""").Interior.Color = RGB(255, 242, 204)
    End With

    REM --- 書式: D列（検算）---
    With ws.Range("D5:D" & lastRow)
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With
    With ws.Range("D5:D" & lastRow).FormatConditions
        .Delete
        With .Add(Type:=xlTextString, String:="NG", TextOperator:=xlContains)
            .Interior.Color = RGB(255, 199, 206)
            .Font.Color = RGB(156, 0, 6)
        End With
        With .Add(Type:=xlTextString, String:="OK", TextOperator:=xlContains)
            .Interior.Color = RGB(198, 239, 206)
            .Font.Color = RGB(0, 97, 0)
        End With
    End With

    REM --- 書式: E列（13桁）---
    With ws.Range("E5:E" & lastRow)
        .NumberFormat = FMT_TEXT
        .HorizontalAlignment = xlCenter
        .Font.Name = "Consolas"
    End With

    REM --- 書式: F〜G列（自動生成・触らせない）---
    With ws.Range("F5:G" & lastRow)
        .Interior.Color = RGB(242, 242, 242)
        .Font.Color = RGB(0, 0, 0)
        .Font.Name = "Consolas"
        .Font.Size = 9
        .HorizontalAlignment = xlLeft
        .WrapText = False
    End With

    REM --- A列・罫線・行高 ---
    ws.Range("A5:A" & lastRow).HorizontalAlignment = xlCenter
    With ws.Range("A4:G" & lastRow)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(142, 169, 219)
        .Borders.Weight = xlThin
    End With
    ws.Rows("5:" & lastRow).RowHeight = 20

    REM --- 列幅 ---
    Dim cw As Variant, ci As Long
    cw = Array(40, 260, 150, 110, 130, 320, 400)
    For ci = 0 To UBound(cw)
        SetColWidthPoints ws.Columns(ci + 1), CDbl(cw(ci))
    Next ci
    SetColWidthPoints ws.Columns("H"), 20#
    SetColWidthPoints ws.Columns("I"), 110#
    SetColWidthPoints ws.Columns("J"), 430#

    REM --- 使い方 ---
    REM  見出し行は先頭の記号で判別して太字にする。行位置をベタ書きすると
    REM  文面を足すたびにずれるため。
    Dim u As Variant, ur As Long, wr As Long
    u = Array("■ 使い方", _
      "1. ヒート（PTPシート）の (01) の後ろの14桁を、C列にそのまま入力します。", _
      "2. D列の検算が OK になると、B列に薬剤名が医薬品マスタから自動で入ります。NG の間はバーコードも描画されません。", _
      "3. GS1データバー限定型のバーコードを描画します。ヒートに印字されているのと同じ symbology です。", _
      "   規格上、GTINの先頭桁が 0 か 1 のものだけ描画できます。調剤包装単位のコードは先頭0なので通常は問題ありません。", _
      "4. 対象行を選んで AssignHeatImage を実行し、ヒートの写真を選びます。F列の名前へ自動でリネームして格納されます。", _
      "5. 「カード印刷」シートで Ctrl+P。A4縦1ページに20枚印刷されます。", _
      "", _
      "■ 画像まわりのマクロ", _
      "AssignHeatImage         選択行に画像を割り当てる。複数行を選ぶと順番にダイアログが出る", _
      "AssignMissingHeatImages 画像が未配置の行だけを上から順に処理する", _
      "ImportHeatImages        G列のパスから画像をカードへ配置し直す", _
      "OpenImageFolder         画像フォルダをエクスプローラで開く", _
      "", _
      "■ 医薬品マスタの更新", _
      "配布ファイル名には年月日が入ります。新しい版が出たら J3 のURLを書き換えて RefreshDrugMaster を実行してください。")
    ur = 4 + N_ITEMS + 3
    For ci = 0 To UBound(u)
        ws.Cells(ur + ci, 1).Value = u(ci)
        If Left$(CStr(u(ci)), 1) = "■" Then
            ws.Cells(ur + ci, 1).Font.Bold = True
            ws.Cells(ur + ci, 1).Font.Size = 12
        End If
    Next ci

    wr = ur + UBound(u) + 2
    With ws.Cells(wr, 1)
        .Value = "【重要】C列はGTIN未入力です。必ず実物のヒートを見て入力してください。"
        .Font.Color = RGB(192, 0, 0): .Font.Bold = True: .Font.Size = 10
    End With
    With ws.Cells(wr + 1, 1)
        .Value = "   同じ薬剤名でもメーカー・規格・包装単位（PTP14錠/100錠等）ごとにGTINは異なります。"
        .Font.Color = RGB(192, 0, 0): .Font.Size = 9
    End With
    With ws.Cells(wr + 3, 1)
        .Value = "※ B列が「該当なし」の場合、そのGTINは医薬品マスタに収載されていません。マスタの版とヒートの表示を確認してください。"
        .Font.Color = RGB(102, 102, 102): .Font.Size = 9
    End With
    With ws.Cells(wr + 4, 1)
        .Value = "※ 画像は既定ではコピーされ、元ファイルは残ります。移動にしたい場合は mdlPhoto の MOVE_AFTER_STORE を True にしてください。"
        .Font.Color = RGB(102, 102, 102): .Font.Size = 9
    End With
    With ws.Cells(wr + 5, 1)
        .Value = "※ 本カードは棚札・目視確認補助用です。調剤監査は必ずヒート実物のGS1コードで行ってください。"
        .Font.Color = RGB(102, 102, 102): .Font.Size = 9
    End With

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A5").Select
    ActiveWindow.FreezePanes = True
End Sub

REM ==========================================================
REM  データシートの数式を書き込む
REM  BuildData からも RepairDataFormulas からも呼ぶ。
REM ----------------------------------------------------------
REM  画像フォルダ J1 の末尾に区切り文字が無い、拡張子 J2 の先頭に
REM  ドットが無い、という入力は普通に起こる。素直に連結すると
REM  「...\画像A0498....jpg」のような壊れたパスになり、格納は
REM  できているのに見つからない、という症状になるので数式側で吸収する。
REM ==========================================================
Private Sub WriteDataFormulas(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim fB As String, fD As String, fE As String, fF As String, fG As String

    lastRow = DATA_ROW1 + N_ITEMS - 1

    REM E列は 医薬品マスタの 調剤包装単位コード と同じ13桁なので
    REM これをキーに XLOOKUP で薬剤名を引く。
    fB = "=IF($E5="""","""",XLOOKUP($E5," & MASTER_TABLE & "[調剤包装単位コード]," & _
         MASTER_TABLE & "[薬剤名],""該当なし""))"
    fD = "=IF($C5="""","""",IF(LEN($C5)<>14,""NG ""&LEN($C5)&""桁""," & _
         "IF(VALUE(RIGHT($C5,1))<>MOD(10-MOD(SUMPRODUCT(MID($C5,SEQUENCE(13),1)*1," & _
         "{3;1;3;1;3;1;3;1;3;1;3;1;3}),10),10),""NG CD不一致"",""OK"")))"
    fE = "=IF($D5<>""OK"","""",RIGHT($C5,13))"
    fF = "=LET(nm,IF($E5="""","""",XLOOKUP($E5," & MASTER_TABLE & "[調剤包装単位コード]," & _
         MASTER_TABLE & "[ファイル名用名称],"""")),ex," & REF_IMG_EXT & "," & _
         "IF(nm="""","""",$C5&""_""&nm&IF(LEFT(ex,1)=""."",ex,"".""&ex)))"
    fG = "=LET(d," & REF_IMG_DIR & ",IF($F5="""",""""," & _
         "IF(RIGHT(d,1)=""" & PATH_SEP & """,d,d&""" & PATH_SEP & """)&$F5))"

    REM 表示形式が「文字列」のセルへ数式を書くと、数式そのものが文字列として
    REM 入ってしまう。貼り直しにも耐えるよう、書き込む前に標準へ戻す。
    ws.Range("B5:B" & lastRow).NumberFormat = FMT_GENERAL
    ws.Range("D5:G" & lastRow).NumberFormat = FMT_GENERAL

    ws.Range("B5").Formula = fB
    ws.Range("D5").Formula = fD
    ws.Range("E5").Formula = fE
    ws.Range("F5").Formula = fF
    ws.Range("G5").Formula = fG
    ws.Range("B5").AutoFill Destination:=ws.Range("B5:B" & lastRow)
    ws.Range("D5:G5").AutoFill Destination:=ws.Range("D5:G" & lastRow)

    REM E列は13桁を文字列として見せたいので書き込んだ後に戻す
    ws.Range("E5:E" & lastRow).NumberFormat = FMT_TEXT
End Sub

REM ==========================================================
REM  既存ブックの数式だけを貼り直す
REM  C列の入力値と設定欄は触らないので、ブックを作り直さずに
REM  数式の修正を取り込める。
REM ==========================================================
Public Sub RepairDataFormulas()
    On Error GoTo Fail
    WriteDataFormulas ThisWorkbook.Worksheets(SHT_DATA)
    Application.Calculate
    MsgBox MSG_REPAIRED, vbInformation
    Exit Sub
Fail:
    MsgBox MSG_REPAIR_NG & vbCrLf & Err.Number & "_ " & Err.Description, vbCritical
End Sub

REM ==========================================================
  REM 2. 命名規則シート
REM ==========================================================
Private Sub BuildNaming()
    Dim ws As Worksheet, i As Long
    Set ws = NewSheet(SHT_NAMING, 1)

    With ws.Range("A1")
        .Value = "ヒート画像 ファイル命名規則"
        .Font.Size = 16: .Font.Bold = True
    End With
    ws.Range("A3").Value = "■ 基本形"
    ws.Range("A3").Font.Bold = True: ws.Range("A3").Font.Size = 12
    With ws.Range("A4")
        .Value = "GTIN14_薬剤名.jpg"
        .Font.Name = "Consolas": .Font.Size = 12
        .Interior.Color = RGB(232, 240, 254)
    End With
    With ws.Range("A5")
        .Value = "薬剤名は医薬品マスタの「販売名 + 規格・製造承認時規格」です。データシートのF列に自動生成されます。"
        .Font.Size = 9: .Font.Color = RGB(102, 102, 102)
    End With

    ws.Range("A7").Value = "■ 実例"
    ws.Range("A7").Font.Bold = True: ws.Range("A7").Font.Size = 12
    Dim ex As Variant
    ex = Array("04987173574513_ポビドンヨードガーグル液７％「シオエ」 ７％１ｍＬ.jpg", _
               "04502072021816_ハリケインゲル歯科用２０％ １ｇ.jpg")
    For i = 0 To UBound(ex)
        With ws.Cells(8 + i, 1)
            .Value = ex(i)
            .Font.Name = "Consolas"
            .Interior.Color = RGB(242, 242, 242)
        End With
    Next i

    ws.Range("A11").Value = "■ 各要素のルール"
    ws.Range("A11").Font.Bold = True: ws.Range("A11").Font.Size = 12
    ws.Range("A12").Value = "要素": ws.Range("B12").Value = "内容": ws.Range("C12").Value = "ルール・注意点"
    With ws.Range("A12:C12")
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    Dim rl As Variant
    rl = Array( _
      Array("GTIN14", "ヒートの(01)の後ろ14桁", "先頭0を含めてそのまま。これが一意キーになる"), _
      Array("薬剤名", "販売名 + 規格", "医薬品マスタから自動取得。手で書き換えない"))
    For i = 0 To UBound(rl)
        ws.Cells(13 + i, 1).Value = rl(i)(0)
        ws.Cells(13 + i, 2).Value = rl(i)(1)
        ws.Cells(13 + i, 3).Value = rl(i)(2)
    Next i
    With ws.Range("A13:A14")
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
    End With
    With ws.Range("A12:C14")
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(142, 169, 219)
    End With

    ws.Range("A16").Value = "■ ファイル名に使えない文字"
    ws.Range("A16").Font.Bold = True: ws.Range("A16").Font.Size = 12
    With ws.Range("A17")
        .Value = "\  /  :  *  ?  ""  <  >  |"
        .Font.Name = "Consolas"
        .Interior.Color = RGB(255, 224, 224)
    End With
    With ws.Range("A18")
        .Value = "※ これらは Power Query 側で薬剤名から除去済みです。医薬品マスタの「ファイル名用名称」列を使っています。"
        .Font.Size = 9: .Font.Color = RGB(102, 102, 102)
    End With

    ws.Range("A20").Value = "■ 医薬品マスタの仕組み"
    ws.Range("A20").Font.Bold = True: ws.Range("A20").Font.Size = 12
    Dim pq As Variant
    pq = Array("・「医薬品マスタ」シートは Power Query が Web上のGS1コード一覧xlsxから自動生成します。", _
               "・キーは「調剤包装単位コード」13桁。ヒートの(01)から先頭0を除いた値と一致します。", _
               "・データシートのE列 13桁 をキーに XLOOKUP で薬剤名を引いています。", _
               "・取込元URLはデータシートの J3 セルです。配布ファイル名には年月日が入ります。", _
               "・URLを書き換えたら RefreshDrugMaster を実行するとマスタが作り直されます。", _
               "・同じURLのまま再取得するだけなら「データ」タブの「すべて更新」でも構いません。")
    For i = 0 To UBound(pq)
        ws.Cells(21 + i, 1).Value = pq(i)
    Next i

    ws.Range("A28").Value = "■ 運用手順"
    ws.Range("A28").Font.Bold = True: ws.Range("A28").Font.Size = 12
    Dim op As Variant
    op = Array("1. データシートのC列にヒートのGTIN14桁を入力する。", _
               "2. B列に薬剤名、F列にファイル名が自動で入る。", _
               "3. ヒートを撮影し、スマホ同期やダウンロードフォルダに置く。手でリネームする必要はない。", _
               "4. 対象行を選んで AssignHeatImage を実行し、写真を選ぶ。F列の名前へリネームして画像フォルダへ格納される。", _
               "5. 続けて画像を配置するか聞かれるので「はい」を選ぶ。あとから ImportHeatImages を実行しても同じ。")
    For i = 0 To UBound(op)
        ws.Cells(29 + i, 1).Value = op(i)
    Next i

    With ws.Range("A34")
        .Value = "■ 重要：パスを書いても自動では表示されません"
        .Font.Bold = True: .Font.Size = 12: .Font.Color = RGB(192, 0, 0)
    End With
    ws.Range("A35").Value = "Excelの仕様上、セルにパスを書くだけでは画像は出ません。画像はセル値ではなく図形オブジェクトのためです。"
    ws.Range("A36").Value = "IMAGE関数はWebのURL専用で、ローカルの C:\... は参照できません。"
    With ws.Range("A37")
        .Value = "⇒ ImportHeatImages マクロを実行すると、G列のパスから画像を一括配置します。"
        .Font.Bold = True
    End With

    SetColWidthPoints ws.Columns("A"), 380#
    SetColWidthPoints ws.Columns("B"), 180#
    SetColWidthPoints ws.Columns("C"), 380#
    ws.Rows(4).RowHeight = 24
    ThisWorkbook.Windows(1).Activate
    ws.Activate
    ActiveWindow.DisplayGridlines = False
End Sub

REM ==========================================================
  REM 3. バーコード計算シート（GS1データバー限定型）
REM ==========================================================
Private Sub BuildEncoder()
    Dim ws As Worksheet, i As Long, r As Long
    Dim lastRow As Long
    Set ws = NewSheet(SHT_ENCODER, 2)
    lastRow = DATA_ROW1 + N_ITEMS - 1

    With ws.Range("A1")
        .Value = "GS1データバー限定型 エンコード結果"
        .Font.Size = 14: .Font.Bold = True
    End With
    With ws.Range("A2")
        .Value = "※ ビット列は VBA の DataBarLimitedBits が生成します。1=バー 0=スペース。このシートは編集不要です。"
        .Font.Size = 9: .Font.Color = RGB(192, 0, 0)
    End With
    With ws.Range("A3")
        .Value = "※ ヒートに印字されているのと同じ symbology です。AI(01) + GTIN14桁 を符号化しています。"
        .Font.Size = 9: .Font.Color = RGB(102, 102, 102)
    End With

    ws.Range("A4").Value = "No."
    ws.Range("B4").Value = "GTIN-14"
    ws.Range("C4").Value = "モジュール列"
    ws.Range("D4").Value = "長さ検証"
    With ws.Range("A4:D4")
        .Font.Bold = True
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    For i = 0 To N_ITEMS - 1
        r = DATA_ROW1 + i
        ws.Cells(r, 1).Formula = "=" & SHT_DATA & "!A" & r
        ws.Cells(r, 2).Formula = "=IF(" & SHT_DATA & "!D" & r & "<>""OK"",""""," & SHT_DATA & "!C" & r & ")"
        ws.Cells(r, 3).Formula = "=IF($B" & r & "="""","""",DataBarLimitedBits($B" & r & "))"
        ws.Cells(r, 4).Formula = "=IF($B" & r & "="""","""",IF(LEN($C" & r & ")=" & DBL_MODULES & _
                                 ",""OK"",IF(LEN($C" & r & ")=0,""NG 先頭桁が0か1でない"",""NG ""&LEN($C" & r & "))))"
    Next i

    REM 数式を書いた後に文字列書式へ。先に設定すると数式が文字列として入る。
    ws.Range("B" & DATA_ROW1 & ":C" & lastRow).NumberFormat = FMT_TEXT
    ws.Range("B" & DATA_ROW1 & ":B" & lastRow).Font.Name = "Consolas"
    With ws.Range("C" & DATA_ROW1 & ":C" & lastRow)
        .Font.Name = "Consolas"
        .Font.Size = 7
    End With
    ws.Range("A" & DATA_ROW1 & ":A" & lastRow).HorizontalAlignment = xlCenter
    ws.Range("D" & DATA_ROW1 & ":D" & lastRow).HorizontalAlignment = xlCenter
    With ws.Range("A4:D" & lastRow)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(142, 169, 219)
    End With

    SetColWidthPoints ws.Columns("A"), 40#
    SetColWidthPoints ws.Columns("B"), 130#
    SetColWidthPoints ws.Columns("C"), 400#
    SetColWidthPoints ws.Columns("D"), 80#
End Sub

REM ==========================================================
  REM 4. カード印刷シート（4x5=20枚 / A4縦1ページ）
REM ==========================================================
Private Sub BuildCards()
    Dim ws As Worksheet
    Dim i As Long, cc As Long, rr As Long, c0 As Long, r0 As Long
    Dim bcL As Long, dRow As Long, lastR As Long, lastC As Long
    Set ws = NewSheet("カード印刷", 3)

    lastC = START_C + COLS_N * CARD_W - 1
    lastR = START_R + 5 * (CARD_H + GAP_R) - GAP_R - 1

    REM 全列を 1pt 幅に（モジュール = 0.353mm）
    REM ※ ColumnWidth は「文字数」単位なので、実測して 1pt に追い込む
    SetColWidthPoints ws.Range(ws.Columns(1), ws.Columns(lastC + 2)), 1#

    For i = 0 To N_ITEMS - 1
        cc = i Mod COLS_N
        rr = i \ COLS_N
        c0 = START_C + cc * CARD_W
        r0 = START_R + rr * (CARD_H + GAP_R)
        bcL = c0 + (CARD_W - DBL_MODULES) \ 2      ' カード内で左右中央に置く
        dRow = 5 + i

        REM --- カード外枠 ---
        With ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + CARD_H - 1, c0 + CARD_W - 1))
            .Interior.Color = RGB(255, 255, 255)
            .BorderAround xlContinuous, xlThin, , RGB(31, 56, 100)
        End With

        REM --- 薬剤名（2行・折り返し）---
        With ws.Range(ws.Cells(r0, c0 + 3), ws.Cells(r0 + 1, c0 + CARD_W - 4))
            .Merge
            .Formula = "=データ!B" & dRow
            .Font.Size = 7
            .Font.Bold = True
            .WrapText = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        REM --- ヒート画像枠（点線）---
        With ws.Range(ws.Cells(r0 + 2, c0 + 3), ws.Cells(r0 + 15, c0 + CARD_W - 4))
            .Merge
            .Value = "ヒート画像"
            .Font.Size = 6
            .Font.Color = RGB(170, 170, 170)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(250, 250, 250)
            .BorderAround xlDash, xlThin, , RGB(204, 204, 204)
        End With

        REM --- バーコード（GS1データバー限定型 74モジュール × 5行）---
        With ws.Range(ws.Cells(r0 + 17, bcL), ws.Cells(r0 + 21, bcL + DBL_MODULES - 1))
            .Formula = "=IFERROR(VALUE(MID(" & SHT_ENCODER & "!$C$" & (DATA_ROW1 + i) & "," & _
                       "COLUMN()-COLUMN($" & ColLetter(bcL) & "$" & (r0 + 17) & ")+1,1)),"""")"
            .NumberFormat = ";;;"
            .Interior.Color = RGB(255, 255, 255)
            With .FormatConditions
                .Delete
                .Add(xlCellValue, xlEqual, "=1").Interior.Color = RGB(0, 0, 0)
            End With
        End With

        REM --- 人が読む数字。GS1 なので AI を添える ---
        With ws.Range(ws.Cells(r0 + 23, c0 + 3), ws.Cells(r0 + 23, c0 + CARD_W - 4))
            .Merge
            .Style = "Normal"
            .Formula = "=IF(" & SHT_DATA & "!D" & dRow & "<>""OK"","""",""(01) ""&" & SHT_DATA & "!C" & dRow & ")"
            .NumberFormat = FMT_TEXT
            .Font.Name = "Consolas"
            .Font.Size = 6
            .Font.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next i

    REM --- 行高さ設定（実寸の要）---
    REM 行1 は印刷範囲外のスペーサ。既定の高さ 18.75pt のままだと
    REM ページレイアウト表示で下がはみ出すので、列A と同じ 1pt に揃える。
    ws.Rows(1).RowHeight = 1
    For i = 0 To 4
        r0 = START_R + i * (CARD_H + GAP_R)
        ws.Rows(r0 & ":" & (r0 + 1)).RowHeight = 10        ' 薬剤名
        ws.Rows((r0 + 2) & ":" & (r0 + 15)).RowHeight = 3.5 ' 画像枠
        ws.Rows(r0 + 16).RowHeight = 2                       ' 余白
        ws.Rows((r0 + 17) & ":" & (r0 + 21)).RowHeight = 13  ' バー 22.9mm
        ws.Rows(r0 + 22).RowHeight = 2                       ' 余白
        ws.Rows(r0 + 23).RowHeight = 8                       ' 人が読む数字
        ws.Rows((r0 + 24) & ":" & (r0 + 25)).RowHeight = 1   ' 下余白
        If i < 4 Then
            ws.Rows((r0 + CARD_H) & ":" & (r0 + CARD_H + GAP_R - 1)).RowHeight = 1
        End If
    Next i

    REM --- 印刷設定（scale 100 固定が実寸維持の要）---
    REM 改ページ位置を読むので、先にシートをアクティブにしておく。
    ThisWorkbook.Windows(1).Activate
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select

    REM 印刷範囲はカード実体 B2 から だけを囲む。
    REM A1 起点にするとスペーサの列A 1pt と行1 分が余分に入り、ページが割れる。
    FitPrintAreaToA4 ws, ws.Range(ws.Cells(START_R, START_C), ws.Cells(lastR, lastC))
End Sub

REM ==========================================================
REM  印刷範囲を A4 縦 1ページに収める
REM  用紙サイズから余白を逆算しただけでは収まらない。プリンタごとに
REM  印刷不可領域があり、実際の印刷可能幅は「用紙幅 - 余白」より狭いため。
REM  そこで実際の改ページ位置を見ながら、1ページに収まる最大の余白を選ぶ。
REM ==========================================================
Private Sub FitPrintAreaToA4(ByVal ws As Worksheet, ByVal target As Range)
    Dim cand As Variant, k As Long
    Dim marginX As Double, marginY As Double
    Dim okX As Boolean, okY As Boolean
    Dim scrSave As Boolean

    scrSave = Application.ScreenUpdating
    Application.ScreenUpdating = True    ' 改ページ位置の取得には再ページ割りが要る

    Application.PrintCommunication = False
    With ws.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = 100                  ' FitToPages は使わない。実寸維持の要
        .HeaderMargin = 0
        .FooterMargin = 0
        .CenterHorizontally = True
        .CenterVertically = True
        .PrintArea = target.Address
        .LeftMargin = 0
        .RightMargin = 0
        .TopMargin = 0
        .BottomMargin = 0
    End With
    Application.PrintCommunication = True

    REM 収まる最大の余白を探す。縦の改ページは左右余白、横の改ページは
    REM 上下余白だけで決まるので、左右と上下は独立に判定できる。
    cand = Array(15#, 12#, 10#, 8#, 6#, 4#, 2#, 0#)
    For k = 0 To UBound(cand)
        If Not okX Then
            ApplyMarginX ws, CDbl(cand(k))
            If ws.VPageBreaks.Count = 0 Then
                marginX = CDbl(cand(k))
                okX = True
            End If
        End If
        If Not okY Then
            ApplyMarginY ws, CDbl(cand(k))
            If ws.HPageBreaks.Count = 0 Then
                marginY = CDbl(cand(k))
                okY = True
            End If
        End If
        If okX And okY Then Exit For
    Next k

    ApplyMarginX ws, marginX
    ApplyMarginY ws, marginY

    REM 余白ゼロでも収まらない場合だけ印刷倍率を落とす
    If Not (okX And okY) Then ShrinkToOnePage ws, target

    Application.ScreenUpdating = scrSave
End Sub

REM --- 左右余白を mm 指定で設定 --------------------------------
Private Sub ApplyMarginX(ByVal ws As Worksheet, ByVal marginMm As Double)
    Dim ptVal As Double
    ptVal = Application.InchesToPoints(marginMm / MM_PER_INCH)
    ws.PageSetup.LeftMargin = ptVal
    ws.PageSetup.RightMargin = ptVal
End Sub

REM --- 上下余白を mm 指定で設定 --------------------------------
Private Sub ApplyMarginY(ByVal ws As Worksheet, ByVal marginMm As Double)
    Dim ptVal As Double
    ptVal = Application.InchesToPoints(marginMm / MM_PER_INCH)
    ws.PageSetup.TopMargin = ptVal
    ws.PageSetup.BottomMargin = ptVal
End Sub

REM --- 1ページ目に収まっている幅 pt ----------------------------
Private Function PageSpanWidth(ByVal ws As Worksheet, ByVal target As Range) As Double
    Dim c As Long
    If ws.VPageBreaks.Count = 0 Then
        PageSpanWidth = target.Width
        Exit Function
    End If
    c = ws.VPageBreaks(1).Location.Column - 1
    If c < target.Column Then
        PageSpanWidth = 0
    Else
        PageSpanWidth = ws.Range(ws.Cells(target.Row, target.Column), _
                                 ws.Cells(target.Row, c)).Width
    End If
End Function

REM --- 1ページ目に収まっている高さ pt --------------------------
Private Function PageSpanHeight(ByVal ws As Worksheet, ByVal target As Range) As Double
    Dim r As Long
    If ws.HPageBreaks.Count = 0 Then
        PageSpanHeight = target.Height
        Exit Function
    End If
    r = ws.HPageBreaks(1).Location.Row - 1
    If r < target.Row Then
        PageSpanHeight = 0
    Else
        PageSpanHeight = ws.Range(ws.Cells(target.Row, target.Column), _
                                  ws.Cells(r, target.Column)).Height
    End If
End Function

REM ==========================================================
REM  余白ゼロでも収まらない場合の退避処理
REM  実測した印刷可能領域から必要な縮小率を求め、1ページに収める。
REM  バーコードが縮むため、結果のモジュール幅と倍率を必ず通知する。
REM ==========================================================
Private Sub ShrinkToOnePage(ByVal ws As Worksheet, ByVal target As Range)
    Dim pw As Double, ph As Double, ratio As Double
    Dim z As Long, k As Long, fitted As Boolean
    Dim moduleMm As Double, magnif As Double
    Dim msg As String

    ws.PageSetup.Zoom = 100
    pw = PageSpanWidth(ws, target)
    ph = PageSpanHeight(ws, target)

    ratio = 1#
    If target.Width > 0 Then
        If pw / target.Width < ratio Then ratio = pw / target.Width
    End If
    If target.Height > 0 Then
        If ph / target.Height < ratio Then ratio = ph / target.Height
    End If

    z = Int(100# * ratio)
    For k = 1 To ZOOM_TRIES
        If z < MIN_ZOOM Then Exit For
        ws.PageSetup.Zoom = z
        If ws.VPageBreaks.Count = 0 And ws.HPageBreaks.Count = 0 Then
            fitted = True
            Exit For
        End If
        z = z - 1
    Next k

    moduleMm = MODULE_PT / PT_PER_MM * ws.PageSetup.Zoom / 100#
    magnif = moduleMm / NOMINAL_MODULE_MM
    If fitted Then
        msg = MSG_A4_ZOOM & vbCrLf & _
              "Zoom_ " & ws.PageSetup.Zoom & "%" & vbCrLf & _
              MSG_MODULE & Format(moduleMm, "0.000") & " mm _ " & Format(magnif * 100#, "0") & "%"
        If magnif < MIN_MAGNIF Then msg = msg & vbCrLf & MSG_MAGNIF_NG
        MsgBox msg, vbExclamation
    Else
        MsgBox MSG_A4_NG, vbCritical
    End If
End Sub

REM --- 列番号 → 列文字 -----------------------------------------
Private Function ColLetter(ByVal n As Long) As String
    Dim s As String
    Do While n > 0
        Dim m As Long
        m = (n - 1) Mod 26
        s = Chr(65 + m) & s
        n = (n - 1 - m) \ 26
    Loop
    ColLetter = s
End Function


REM ==========================================================
REM  列幅を「ポイント」で正確に設定するヘルパー
REM  ColumnWidth は文字数単位なので、実測して追い込む
REM ==========================================================
Private Sub SetColWidthPoints(ByVal cols As Range, ByVal ptTarget As Double)
    Dim i As Long, cw As Double, w As Double
    REM 初期推定（1文字 ≒ 7pt）
    cols.ColumnWidth = ptTarget / 7
    REM 実測して比例補正を2回反復
    For i = 1 To 2
        w = cols.Columns(1).Width
        If w <= 0 Then Exit For
        cw = cols.Columns(1).ColumnWidth * (ptTarget / w)
        If cw < 0 Then cw = 0
        If cw > 255 Then cw = 255
        cols.ColumnWidth = cw
    Next i
End Sub
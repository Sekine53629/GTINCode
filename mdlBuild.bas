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
Public Const BC_OFS  As Long = 17      ' バーコード左オフセット（余白6+QZ11）

REM --- 定数（A4 印刷フィット）--------------------------------
Private Const MM_PER_INCH   As Double = 25.4
Private Const PT_PER_MM     As Double = 72# / 25.4  ' 1mm あたりのポイント数
Private Const MODULE_PT     As Double = 1#          ' バーコード 1モジュール幅 = 列幅 1pt
Private Const JAN_MODULE_MM As Double = 0.33        ' JAN-13 標準モジュール幅 倍率100%
Private Const MIN_MAGNIF    As Double = 0.8         ' GS1 が認める最小倍率 80%
Private Const MIN_ZOOM      As Long = 60            ' 縮小印刷の下限
Private Const ZOOM_TRIES    As Long = 12            ' 縮小率の追い込み回数
Private Const MSG_A4_ZOOM   As String = "A4 1ページに収めるため印刷倍率を下げました。"
Private Const MSG_A4_NG     As String = "A4 1ページに収まりませんでした。プリンタの紙サイズと余白設定を確認してください。"
Private Const MSG_MODULE    As String = "バーコード 1モジュール幅_ "
Private Const MSG_MAGNIF_NG As String = "GS1 の最小倍率 80% を下回っています。読み取り精度が落ちる可能性があります。"

Public Sub BuildWorkbook()
    Dim calcSave As XlCalculation
    calcSave = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    On Error GoTo Fail
    RemoveOldSheets
    BuildData
    BuildNaming
    BuildEncoder
    BuildCards
    Sheets("データ").Activate

    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "生成完了。" & vbCrLf & "データシートのC列にヒートのGTIN14桁を入力してください。", vbInformation
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
    For Each nm In Array("データ", "命名規則", "バーコード計算", "カード印刷")
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
    Dim ws As Worksheet, i As Long, r As Long
    Set ws = NewSheet("データ", 0)
    ws.Move Before:=ThisWorkbook.Worksheets(1)

    With ws.Range("A1")
        .Value = "調剤在庫マスタ（ヒートGTIN カード印刷用）"
        .Font.Size = 14: .Font.Bold = True
    End With
    With ws.Range("A2")
        .Value = "※ ヒートの (01) の後ろの14桁をC列にそのまま入力。先頭0を除いた13桁でバーコード化します。"
        .Font.Size = 9: .Font.Color = RGB(102, 102, 102)
    End With

    REM 設定欄
    ws.Range("H1").Value = "画像フォルダ（ここを変えれば全行反映）"
    ws.Range("H1").Font.Bold = True: ws.Range("H1").Font.Size = 10
    With ws.Range("H2")
        .Value = "C:\薬局\ヒート画像\"
        .Font.Color = RGB(0, 0, 255): .Font.Name = "Consolas"
        .Interior.Color = RGB(255, 249, 224)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 143, 0): .Borders.Weight = xlMedium
    End With
    ws.Range("K1").Value = "拡張子"
    ws.Range("K1").Font.Bold = True: ws.Range("K1").Font.Size = 10
    With ws.Range("K2")
        .Value = ".jpg"
        .Font.Color = RGB(0, 0, 255): .Font.Name = "Consolas"
        .Interior.Color = RGB(255, 249, 224)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 143, 0): .Borders.Weight = xlMedium
    End With

    REM ヘッダ
    Dim hdr As Variant
    hdr = Array("No.", "薬剤名（規格）", "GTIN-14（ヒート入力）", "検算", _
                "JAN13（描画用）", "分類・備考", "成分名", "製剤種", _
                "容量規格", "屋号（メーカー）", "ファイル名（自動生成）", "フルパス（自動生成）")
    For i = 0 To UBound(hdr)
        ws.Cells(4, i + 1).Value = hdr(i)
    Next i
    With ws.Range("A4:L4")
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    REM 品目マスタ: 薬剤名 / 分類 / 成分名 / 製剤種 / 容量
    Dim m As Variant
    m = Array( _
      Array("レクサプロ錠 10mg", "抗うつ薬（SSRI）", "エスシタロプラム", "錠", "10mg"), _
      Array("ジェイゾロフト錠 25mg", "抗うつ薬（SSRI）", "セルトラリン", "錠", "25mg"), _
      Array("パキシル錠 10mg", "抗うつ薬（SSRI）", "パロキセチン", "錠", "10mg"), _
      Array("サインバルタカプセル 20mg", "抗うつ薬（SNRI）", "デュロキセチン", "カプセル", "20mg"), _
      Array("ベンラファキシン徐放カプセル 37.5mg", "抗うつ薬（SNRI）", "ベンラファキシン", "徐放カプセル", "37.5mg"), _
      Array("ミルタザピン錠 15mg", "抗うつ薬（NaSSA）", "ミルタザピン", "錠", "15mg"), _
      Array("トラゾドン錠 25mg", "抗うつ薬", "トラゾドン", "錠", "25mg"), _
      Array("アリピプラゾール錠 3mg", "抗精神病薬（非定型）", "アリピプラゾール", "錠", "3mg"), _
      Array("オランザピン錠 5mg", "抗精神病薬（非定型）", "オランザピン", "錠", "5mg"), _
      Array("リスペリドン錠 1mg", "抗精神病薬（非定型）", "リスペリドン", "錠", "1mg"), _
      Array("クエチアピン錠 25mg", "抗精神病薬（非定型）", "クエチアピン", "錠", "25mg"), _
      Array("ブロナンセリン錠 2mg", "抗精神病薬（非定型）", "ブロナンセリン", "錠", "2mg"), _
      Array("ハロペリドール錠 0.75mg", "抗精神病薬（定型）", "ハロペリドール", "錠", "0.75mg"), _
      Array("炭酸リチウム錠 200mg", "気分安定薬", "炭酸リチウム", "錠", "200mg"), _
      Array("バルプロ酸ナトリウム徐放錠 200mg", "気分安定薬", "バルプロ酸ナトリウム", "徐放錠", "200mg"), _
      Array("ラモトリギン錠 25mg", "気分安定薬（皮疹注意）", "ラモトリギン", "錠", "25mg"), _
      Array("エチゾラム錠 2mg", "睡眠薬（向精神薬）", "エチゾラム", "錠", "2mg"), _
      Array("スボレキサント錠 15mg", "睡眠薬（向精神薬）", "スボレキサント", "錠", "15mg"), _
      Array("レンボレキサント錠 5mg", "睡眠薬", "レンボレキサント", "錠", "5mg"), _
      Array("ロラゼパム錠 0.5mg", "抗不安薬（向精神薬）", "ロラゼパム", "錠", "0.5mg"))

    For i = 0 To N_ITEMS - 1
        r = 5 + i
        ws.Cells(r, 1).Value = i + 1          ' No.
        ws.Cells(r, 2).Value = m(i)(0)        ' 薬剤名
        ws.Cells(r, 6).Value = m(i)(1)        ' 分類
        ws.Cells(r, 7).Value = m(i)(2)        ' 成分名
        ws.Cells(r, 8).Value = m(i)(3)        ' 製剤種
        ws.Cells(r, 9).Value = m(i)(4)        ' 容量規格
        ws.Cells(r, 10).Value = "未定"        ' 屋号（要確認）
    Next i

    REM 数式（検算・13桁抽出・ファイル名・フルパス）
    Dim fD As String, fE As String, fK As String, fL As String
    fD = "=IF($C5="""","""",IF(LEN($C5)<>14,""✕ ""&LEN($C5)&""桁""," & _
         "IF(VALUE(RIGHT($C5,1))<>MOD(10-MOD(SUMPRODUCT(MID($C5,SEQUENCE(13),1)*1," & _
         "{3;1;3;1;3;1;3;1;3;1;3;1;3}),10),10),""✕ CD不一致"",""✓ OK"")))"
    fE = "=IF(LEFT($D5,1)<>""✓"","""",RIGHT($C5,13))"
    fK = "=IF($C5="""","""",$C5&""_""&$G5&""_""&$H5&""_""&" & _
         "SUBSTITUTE($I5,""."",""-"")&""_""&$J5&$K$2)"
    fL = "=IF($K5="""","""",$H$2&$K5)"
    ws.Range("D5").Formula = fD
    ws.Range("E5").Formula = fE
    ws.Range("K5").Formula = fK
    ws.Range("L5").Formula = fL
    ws.Range("D5:E5").AutoFill Destination:=ws.Range("D5:E" & (4 + N_ITEMS))
    ws.Range("K5:L5").AutoFill Destination:=ws.Range("K5:L" & (4 + N_ITEMS))

    REM 書式: C列（GTIN入力欄）
    With ws.Range("C5:C" & (4 + N_ITEMS))
        .NumberFormat = "@"
        .Font.Color = RGB(0, 0, 255)
        .Font.Name = "Consolas"
        .HorizontalAlignment = xlCenter
    End With
    REM 未入力を黄色く（入力済みは白）
    With ws.Range("C5:C" & (4 + N_ITEMS)).FormatConditions
        .Delete
        .Add(xlCellValue, xlEqual, "=""""").Interior.Color = RGB(255, 242, 204)
    End With

    REM 検算列
    With ws.Range("D5:D" & (4 + N_ITEMS))
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With
    With ws.Range("D5:D" & (4 + N_ITEMS)).FormatConditions
        .Delete
        With .Add(Type:=xlTextString, String:="✕", TextOperator:=xlContains)
            .Interior.Color = RGB(255, 199, 206)
            .Font.Color = RGB(156, 0, 6)
        End With
        With .Add(Type:=xlTextString, String:="✓", TextOperator:=xlContains)
            .Interior.Color = RGB(198, 239, 206)
            .Font.Color = RGB(0, 97, 0)
        End With
    End With

    REM E列（13桁）
    With ws.Range("E5:E" & (4 + N_ITEMS))
        .NumberFormat = "@"
        .HorizontalAlignment = xlCenter
        .Font.Name = "Consolas"
    End With

    REM G〜J列（手入力欄）
    With ws.Range("G5:J" & (4 + N_ITEMS))
        .Font.Color = RGB(0, 0, 255)
        .Interior.Color = RGB(255, 249, 224)
        .HorizontalAlignment = xlCenter
    End With
    REM J列「未定」を警告色
    With ws.Range("J5:J" & (4 + N_ITEMS)).FormatConditions
        .Delete
        With .Add(Type:=xlTextString, String:="未定", TextOperator:=xlContains)
            .Interior.Color = RGB(255, 242, 204)
            .Font.Color = RGB(191, 143, 0)
        End With
    End With

    REM K〜L列（自動生成・触らせない）
    With ws.Range("K5:L" & (4 + N_ITEMS))
        .Interior.Color = RGB(242, 242, 242)
        .Font.Color = RGB(0, 0, 0)
        .Font.Name = "Consolas"
        .Font.Size = 9
        .HorizontalAlignment = xlLeft
        .WrapText = False
    End With

    REM A列・罫線・行高
    ws.Range("A5:A" & (4 + N_ITEMS)).HorizontalAlignment = xlCenter
    With ws.Range("A4:L" & (4 + N_ITEMS))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(142, 169, 219)
        .Borders.Weight = xlThin
    End With
    ws.Rows("5:" & (4 + N_ITEMS)).RowHeight = 20

    REM 列幅
    Dim cw As Variant, ci As Long
    cw = Array(40, 200, 140, 100, 140, 160, 140, 80, 70, 100, 340, 440)
    For ci = 0 To UBound(cw)
        SetColWidthPoints ws.Columns(ci + 1), CDbl(cw(ci))
    Next ci

    REM 使い方
    Dim u As Variant, ur As Long
    u = Array("■ 使い方", _
      "1. ヒート（PTPシート）の (01) の後ろの14桁を、C列にそのまま入力します。", _
      "2. D列の検算が ✓ OK になれば正しいコードです。✕ の間はバーコードは描画されません。", _
      "3. 先頭の0を除いた13桁（E列）でJANバーコードを描画します。", _
      "4. ヒートの写真は「カード印刷」シートの点線枠に直接貼り付けてください。", _
      "5. 「カード印刷」シートで Ctrl+P。A4縦1ページに20枚印刷されます。")
    ur = 4 + N_ITEMS + 3
    For ci = 0 To UBound(u)
        ws.Cells(ur + ci, 1).Value = u(ci)
    Next ci
    ws.Cells(ur, 1).Font.Bold = True
    ws.Cells(ur, 1).Font.Size = 12

    With ws.Cells(ur + 8, 1)
        .Value = "⚠ 重要：C列はGTIN未入力です。必ず実物のヒートを見て入力してください。"
        .Font.Color = RGB(192, 0, 0): .Font.Bold = True: .Font.Size = 10
    End With
    With ws.Cells(ur + 9, 1)
        .Value = "   同じ薬剤名でもメーカー・規格・包装単位（PTP14錠/100錠等）ごとにGTINは異なります。"
        .Font.Color = RGB(192, 0, 0): .Font.Size = 9
    End With
    With ws.Cells(ur + 11, 1)
        .Value = "※ G〜J列の成分名・製剤種・容量は参考値です。実物と照合してください。"
        .Font.Color = RGB(102, 102, 102): .Font.Size = 9
    End With
    With ws.Cells(ur + 12, 1)
        .Value = "※ 本カードは棚札・目視確認補助用です。調剤監査は必ずヒート実物のGS1コードで行ってください。"
        .Font.Color = RGB(102, 102, 102): .Font.Size = 9
    End With

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A5").Select
    ActiveWindow.FreezePanes = True
End Sub

REM ==========================================================
  REM 2. 命名規則シート
REM ==========================================================
Private Sub BuildNaming()
    Dim ws As Worksheet, i As Long
    Set ws = NewSheet("命名規則", 1)

    With ws.Range("A1")
        .Value = "ヒート画像 ファイル命名規則"
        .Font.Size = 16: .Font.Bold = True
    End With
    ws.Range("A3").Value = "■ 基本形"
    ws.Range("A3").Font.Bold = True: ws.Range("A3").Font.Size = 12
    With ws.Range("A4")
        .Value = "GTIN14_成分名_製剤種_容量規格_屋号.jpg"
        .Font.Name = "Consolas": .Font.Size = 12
        .Interior.Color = RGB(232, 240, 254)
    End With

    ws.Range("A6").Value = "■ 実例"
    ws.Range("A6").Font.Bold = True: ws.Range("A6").Font.Size = 12
    Dim ex As Variant
    ex = Array("04987224712215_エスシタロプラム_錠_10mg_屋号.jpg", _
               "04987128131117_デュロキセチン_カプセル_20mg_屋号.jpg", _
               "04987041814116_ベンラファキシン_徐放カプセル_37-5mg_屋号.jpg")
    For i = 0 To UBound(ex)
        With ws.Cells(7 + i, 1)
            .Value = ex(i)
            .Font.Name = "Consolas"
            .Interior.Color = RGB(242, 242, 242)
        End With
    Next i
    With ws.Range("A10")
        .Value = "※ 実例のGTIN・屋号は書式を示すためのものです（実在確認していません）。実際の値はヒートでご確認ください。"
        .Font.Size = 9: .Font.Color = RGB(192, 0, 0)
    End With

    ws.Range("A12").Value = "■ 各要素のルール"
    ws.Range("A12").Font.Bold = True: ws.Range("A12").Font.Size = 12
    ws.Range("A13").Value = "要素": ws.Range("B13").Value = "内容": ws.Range("C13").Value = "ルール・注意点"
    With ws.Range("A13:C13")
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    Dim rl As Variant
    rl = Array( _
      Array("GTIN14", "ヒートの(01)の後ろ14桁", "先頭0を含めてそのまま。これが一意キーになる"), _
      Array("成分名", "一般名（カタカナ）", "商品名ではなく成分名。後発品との照合が楽になる"), _
      Array("製剤種", "錠 / カプセル / 徐放錠 等", "OD錠・徐放錠は別物として必ず区別する"), _
      Array("容量規格", "10mg / 37.5mg 等", "小数点は - （ハイフン）に置換。例: 37.5mg → 37-5mg"), _
      Array("屋号", "メーカー名・屋号", "後発品は屋号で別物になるため必須。未定なら「未定」"))
    For i = 0 To UBound(rl)
        ws.Cells(14 + i, 1).Value = rl(i)(0)
        ws.Cells(14 + i, 2).Value = rl(i)(1)
        ws.Cells(14 + i, 3).Value = rl(i)(2)
    Next i
    With ws.Range("A14:A18")
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
    End With
    With ws.Range("A13:C18")
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(142, 169, 219)
    End With

    ws.Range("A20").Value = "■ ファイル名に使えない文字"
    ws.Range("A20").Font.Bold = True: ws.Range("A20").Font.Size = 12
    With ws.Range("A21")
        .Value = "\  /  :  *  ?  ""  <  >  |"
        .Font.Name = "Consolas"
        .Interior.Color = RGB(255, 224, 224)
    End With
    With ws.Range("A22")
        .Value = "※ 容量の小数点「.」はファイル名では使えますが、拡張子と紛らわしいため - に置換しています。"
        .Font.Size = 9: .Font.Color = RGB(102, 102, 102)
    End With

    ws.Range("A24").Value = "■ 運用手順"
    ws.Range("A24").Font.Bold = True: ws.Range("A24").Font.Size = 12
    Dim op As Variant
    op = Array("1. ヒートを撮影し、画像フォルダに保存する。", _
               "2. 「データ」シートのC列にGTIN14、G〜J列に4要素を入力する。", _
               "3. K列に正しいファイル名が自動生成されるので、それをコピーして画像をリネームする。", _
               "4. L列のフルパスが実ファイルの場所と一致していれば取込可能になる。")
    For i = 0 To UBound(op)
        ws.Cells(25 + i, 1).Value = op(i)
    Next i

    With ws.Range("A30")
        .Value = "■ 重要：パスを書いても自動では表示されません"
        .Font.Bold = True: .Font.Size = 12: .Font.Color = RGB(192, 0, 0)
    End With
    ws.Range("A31").Value = "Excelの仕様上、セルにパスを書くだけでは画像は出ません（画像はセル値ではなく図形オブジェクトのため）。"
    ws.Range("A32").Value = "IMAGE関数はWebのURL専用で、ローカルのC:\... は参照できません。"
    With ws.Range("A33")
        .Value = "⇒ 下の ImportHeatImages マクロを実行すると、L列のパスから画像を一括配置します。"
        .Font.Bold = True
    End With

    SetColWidthPoints ws.Columns("A"), 300#
    SetColWidthPoints ws.Columns("B"), 200#
    SetColWidthPoints ws.Columns("C"), 420#
    ws.Rows(4).RowHeight = 24
    ThisWorkbook.Windows(1).Activate
    ws.Activate
    ActiveWindow.DisplayGridlines = False
End Sub

REM ==========================================================
  REM 3. バーコード計算シート（JAN-13 規格定義）
REM ==========================================================
Private Sub BuildEncoder()
    Dim ws As Worksheet, i As Long
    Set ws = NewSheet("バーコード計算", 2)

    With ws.Range("A1")
        .Value = "JAN-13 エンコード定義（編集不要）"
        .Font.Size = 14: .Font.Bold = True
    End With
    With ws.Range("A2")
        .Value = "※ このシートは規格定義のルックアップ表です。値を変更しないでください。"
        .Font.Size = 9: .Font.Color = RGB(192, 0, 0)
    End With

    REM パリティ表
    ws.Range("A4").Value = "■ 先頭桁別 左半分パリティパターン"
    ws.Range("A4").Font.Bold = True
    ws.Range("A5").Value = "桁"
    ws.Range("B5").Value = "パターン(2-7桁目)"
    Dim par As Variant
    par = Array("AAAAAA", "AABABB", "AABBAB", "AABBBA", "ABAABB", _
                "ABBAAB", "ABBBAA", "ABABAB", "ABABBA", "ABBABA")
    For i = 0 To 9
        ws.Cells(6 + i, 1).Value = i
        ws.Cells(6 + i, 2).Value = par(i)
    Next i

    REM エンコードテーブル
    ws.Range("D4").Value = "■ エンコードテーブル（1=黒 0=白）"
    ws.Range("D4").Font.Bold = True
    ws.Range("D5").Value = "数字"
    ws.Range("E5").Value = "A (左奇)"
    ws.Range("F5").Value = "B (左偶)"
    ws.Range("G5").Value = "C (右)"
    Dim ea As Variant, eb As Variant, ec As Variant
    ea = Array("0001101", "0011001", "0010011", "0111101", "0100011", _
               "0110001", "0101111", "0111011", "0110111", "0001011")
    eb = Array("0100111", "0110011", "0011011", "0100001", "0011101", _
               "0111001", "0000101", "0010001", "0001001", "0010111")
    ec = Array("1110010", "1100110", "1101100", "1000010", "1011100", _
               "1001110", "1010000", "1000100", "1001000", "1110100")
    For i = 0 To 9
        ws.Cells(6 + i, 4).Value = i
        ws.Cells(6 + i, 5).Value = "'" & ea(i)
        ws.Cells(6 + i, 6).Value = "'" & eb(i)
        ws.Cells(6 + i, 7).Value = "'" & ec(i)
    Next i
    ws.Range("E6:G15").NumberFormat = "@"
    ws.Range("E6:G15").Font.Name = "Consolas"

    REM 見出し行
    ws.Range("A18").Value = "■ 商品別 95モジュール展開"
    ws.Range("A18").Font.Bold = True
    ws.Range("A19").Value = "No."
    ws.Range("B19").Value = "'JAN13"
    ws.Range("C19").Value = "パリティ"
    ws.Range("D19").Value = "95ビット列"
    ws.Range("E19").Value = "長さ検証"
    ws.Range("A19:E19").Font.Bold = True
    ws.Range("A19:E19").Interior.Color = RGB(217, 217, 217)

    REM 展開数式
    Dim r As Long
    For i = 0 To N_ITEMS - 1
        r = 20 + i
        ws.Cells(r, 1).Formula = "=データ!A" & (5 + i)
        ws.Cells(r, 2).Formula = "=データ!E" & (5 + i)
        ws.Cells(r, 3).Formula = "=IF($B" & r & "="""",""""," & _
            "XLOOKUP(VALUE(LEFT($B" & r & ",1)),$A$6:$A$15,$B$6:$B$15))"
        ws.Cells(r, 4).Formula = "=IF($B" & r & "="""","""",""101""&" & _
            "TEXTJOIN("""",TRUE,MAP(SEQUENCE(6),LAMBDA(i,XLOOKUP(VALUE(MID($B" & r & ",i+1,1))," & _
            "$D$6:$D$15,IF(MID($C" & r & ",i,1)=""A"",$E$6:$E$15,$F$6:$F$15)))))" & _
            "&""01010""&TEXTJOIN("""",TRUE,MAP(SEQUENCE(6),LAMBDA(i," & _
            "XLOOKUP(VALUE(MID($B" & r & ",i+7,1)),$D$6:$D$15,$G$6:$G$15))))&""101"")"
        ws.Cells(r, 5).Formula = "=IF($D" & r & "="""",""""," & _
            "IF(LEN($D" & r & ")=95,""OK"",""NG ""&LEN($D" & r & ")))"
    Next i
    ws.Range("B20:B" & (19 + N_ITEMS)).NumberFormat = "@"
    ws.Range("D20:D" & (19 + N_ITEMS)).Font.Size = 7

    SetColWidthPoints ws.Columns("A"), 40#
    SetColWidthPoints ws.Columns("B"), 120#
    SetColWidthPoints ws.Columns("C"), 80#
    SetColWidthPoints ws.Columns("D"), 400#
    SetColWidthPoints ws.Columns("E"), 80#
End Sub

REM ==========================================================
  REM 4. カード印刷シート（4×5=20枚 / A4縦1ページ）
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
        bcL = c0 + BC_OFS
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

        REM --- バーコード（95列 × 5行）---
        With ws.Range(ws.Cells(r0 + 17, bcL), ws.Cells(r0 + 21, bcL + 94))
            .Formula = "=IFERROR(VALUE(MID(バーコード計算!$D$" & (20 + i) & "," & _
                       "COLUMN()-COLUMN($" & ColLetter(bcL) & "$" & (r0 + 17) & ")+1,1)),"""")"
            .NumberFormat = ";;;"
            .Interior.Color = RGB(255, 255, 255)
            With .FormatConditions
                .Delete
                .Add(xlCellValue, xlEqual, "=1").Interior.Color = RGB(0, 0, 0)
            End With
        End With

        REM --- JAN数字 ---
        With ws.Range(ws.Cells(r0 + 23, c0 + 3), ws.Cells(r0 + 23, c0 + CARD_W - 4))
            .Merge
            .Style = "Normal"
            .Formula = "=データ!E" & dRow
            .NumberFormat = "@"
            .Font.Name = "Consolas"
            .Font.Size = 6
            .Font.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next i

    REM --- 行高さ設定（実寸の要）---
    For i = 0 To 4
        r0 = START_R + i * (CARD_H + GAP_R)
        ws.Rows(r0 & ":" & (r0 + 1)).RowHeight = 10        ' 薬剤名
        ws.Rows((r0 + 2) & ":" & (r0 + 15)).RowHeight = 3.5 ' 画像枠
        ws.Rows(r0 + 16).RowHeight = 2                       ' 余白
        ws.Rows((r0 + 17) & ":" & (r0 + 21)).RowHeight = 13  ' バー 22.9mm
        ws.Rows(r0 + 22).RowHeight = 2                       ' 余白
        ws.Rows(r0 + 23).RowHeight = 8                       ' JAN数字
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
    magnif = moduleMm / JAN_MODULE_MM
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
REM =========================================================
REM  医薬品マスタ取込モジュール
REM  Web上に公開されている GS1コード一覧 xlsx を Power Query で
REM  取り込み、「医薬品マスタ」シートにテーブルとして展開する。
REM ---------------------------------------------------------
REM  取込元の列 調剤包装単位コード は13桁で、これがヒートに印字
REM  される GS1 の (01) から先頭0を除いた値そのもの。
REM  つまり データシートのE列 JAN13 と同じ値で突合できる。
REM =========================================================
Option Explicit

REM --- シート・テーブル・クエリ名 -------------------------------
Public Const MASTER_SHEET As String = "医薬品マスタ"
Public Const MASTER_TABLE As String = "tblDrugMaster"
Public Const MASTER_QUERY As String = "医薬品マスタ"

REM --- 既定の取込元URL -----------------------------------------
REM 配布ファイル名に年月日が入るため、新しい版が出たら
REM データシートの設定欄を書き換えて RefreshDrugMaster を実行する。
Public Const DEFAULT_MASTER_URL As String = _
    "https://kusuri-yakuzaishi.com/wp-content/uploads/2026/07/GS1-20260731.xlsx"

REM --- 取込元の列名。配布ファイルのヘッダと一致していること -----
Private Const SRC_CODE As String = "調剤包装単位コード"
Private Const SRC_NAME As String = "販売名"
Private Const SRC_SPEC As String = "規格・製造承認時規格"

REM --- 取込後の列構成 ------------------------------------------
Private Const OUT_CODE As String = "調剤包装単位コード"
Private Const OUT_GTIN As String = "GTIN14"
Private Const OUT_NAME As String = "薬剤名"
Private Const OUT_SAFE As String = "ファイル名用名称"

Private Const CODE_LEN As Long = 13     ' 調剤包装単位コードの桁数
Private Const QUOTE_CH As Long = 34     ' ダブルクォートの文字コード
Private Const M_QUOTE As String = "~"   ' Mコード組立て用の仮クォート

Private Const MSG_MASTER_ERR As String = _
    "医薬品マスタの取込に失敗しました。ネットワークとURLを確認し、RefreshDrugMaster を実行してください。"
Private Const MSG_MASTER_OK As String = "医薬品マスタを取り込みました。件数_ "
Private Const MSG_NO_URL As String = "医薬品マスタのURLが未設定です。データシートの設定欄を確認してください。"

REM ==========================================================
REM  データシートの設定欄のURLで医薬品マスタを取り込み直す
REM  マスタの新版が出たときはこれを実行する
REM ==========================================================
Public Sub RefreshDrugMaster()
    On Error GoTo Fail
    Dim srcUrl As String
    Dim n As Long

    srcUrl = Trim$(CStr(ThisWorkbook.Worksheets(SHT_DATA).Range(CELL_MASTER_URL).Value))
    If Len(srcUrl) = 0 Then
        MsgBox MSG_NO_URL, vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    n = BuildDrugMaster(srcUrl)
    Application.ScreenUpdating = True
    If n > 0 Then MsgBox MSG_MASTER_OK & Format(n, "#,##0") & " 件", vbInformation
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox MSG_MASTER_ERR & vbCrLf & Err.Number & "_ " & Err.Description, vbCritical
End Sub

REM ==========================================================
REM  医薬品マスタシートを作り直して取り込む
REM  戻り値は取り込んだ件数。失敗時は0を返し、後続の数式が壊れない
REM  ようヘッダだけのからテーブルを置いておく。
REM ==========================================================
Public Function BuildDrugMaster(ByVal srcUrl As String) As Long
    Dim ws As Worksheet
    Dim lo As ListObject

    DropMasterObjects
    Set ws = ThisWorkbook.Worksheets.Add( _
                 After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = MASTER_SHEET

    On Error GoTo Fallback
    ThisWorkbook.Queries.Add Name:=MASTER_QUERY, Formula:=MasterFormula(srcUrl)
    Set lo = ws.ListObjects.Add(SourceType:=xlSrcExternal, _
                                Source:=MasterConnString(), _
                                XlListObjectHasHeaders:=xlYes, _
                                Destination:=ws.Range("A1"))
    lo.Name = MASTER_TABLE
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = "SELECT * FROM [" & MASTER_QUERY & "]"
        .RefreshStyle = xlInsertDeleteCells
        .AdjustColumnWidth = False
        .BackgroundQuery = False
        .Refresh BackgroundQuery:=False
    End With
    On Error GoTo 0

    ws.Rows(1).Font.Bold = True
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    BuildDrugMaster = lo.ListRows.Count
    Exit Function

Fallback:
    BuildDrugMaster = 0
    BuildStubTable ws
    MsgBox MSG_MASTER_ERR & vbCrLf & Err.Number & "_ " & Err.Description, vbExclamation
End Function

REM --- 取込失敗時に置くヘッダだけのテーブル --------------------
REM  これが無いと データシートの XLOOKUP が #NAME? になり、
REM  数式の書き込み自体が失敗する。
Private Sub BuildStubTable(ByVal ws As Worksheet)
    Dim hdr As Variant, c As Long
    Dim lo As ListObject

    On Error Resume Next
    For c = ws.ListObjects.Count To 1 Step -1
        ws.ListObjects(c).Delete
    Next c
    ws.Cells.Clear
    On Error GoTo 0

    hdr = MasterHeaders()
    For c = 0 To UBound(hdr)
        ws.Cells(1, c + 1).Value = hdr(c)
    Next c
    Set lo = ws.ListObjects.Add(xlSrcRange, _
                 ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(hdr) + 1)), , xlYes)
    lo.Name = MASTER_TABLE
    ws.Rows(1).Font.Bold = True
End Sub

REM --- 取込後の列名一覧 ----------------------------------------
Private Function MasterHeaders() As Variant
    MasterHeaders = Array(OUT_CODE, OUT_GTIN, OUT_NAME, OUT_SAFE, _
                          "販売名", "規格・製造承認時規格", "剤形", "区分名", _
                          "データ登録企業名", "包装形態", "調剤包装単位名称")
End Function

REM --- Power Query 接続文字列 ----------------------------------
Private Function MasterConnString() As String
    MasterConnString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & _
                       MASTER_QUERY & ";Extended Properties=" & Chr(QUOTE_CH) & Chr(QUOTE_CH)
End Function

REM ==========================================================
REM  Mコードを組み立てる
REM  可読性のため " は ~ で書き、最後に一括置換する。
REM  規格・製造承認時規格 のように中黒を含む列名は M のブラケット
REM  記法では参照できず、[#"..."] と書く必要がある。
REM ==========================================================
Private Function MasterFormula(ByVal srcUrl As String) As String
    Dim m As String

    m = "let" & vbCrLf
    m = m & "    Src   = Excel.Workbook(Web.Contents(~" & srcUrl & "~), null, true)," & vbCrLf
    m = m & "    Raw   = Src{0}[Data]," & vbCrLf
    m = m & "    Head  = Table.PromoteHeaders(Raw, [PromoteAllScalars=true])," & vbCrLf
    m = m & "    Code  = Table.TransformColumns(Head, {{~" & SRC_CODE & "~, each Text.Trim(Text.From(_)), type text}})," & vbCrLf
    m = m & "    Valid = Table.SelectRows(Code, each [" & SRC_CODE & "] <> null and Text.Length([" & SRC_CODE & "]) = " & CODE_LEN & ")," & vbCrLf
    m = m & "    Uniq  = Table.Distinct(Valid, {~" & SRC_CODE & "~})," & vbCrLf
    m = m & "    Gtin  = Table.AddColumn(Uniq, ~" & OUT_GTIN & "~, each ~0~ & [" & SRC_CODE & "], type text)," & vbCrLf
    m = m & "    Nm    = Table.AddColumn(Gtin, ~" & OUT_NAME & "~, each Text.Trim(Text.Combine(List.RemoveNulls({[" & SRC_NAME & "], [#~" & SRC_SPEC & "~]}), ~ ~)), type text)," & vbCrLf
    m = m & "    Safe  = Table.AddColumn(Nm, ~" & OUT_SAFE & "~, each Text.Remove([" & OUT_NAME & "], {~\~, ~/~, ~:~, ~*~, ~?~, Character.FromNumber(" & QUOTE_CH & "), ~<~, ~>~, ~|~}), type text)," & vbCrLf
    m = m & "    Keep  = Table.SelectColumns(Safe, {" & QuotedHeaderList() & "})," & vbCrLf
    m = m & "    Srt   = Table.Sort(Keep, {{~" & SRC_CODE & "~, Order.Ascending}})" & vbCrLf
    m = m & "in" & vbCrLf
    m = m & "    Srt"

    MasterFormula = Replace(m, M_QUOTE, Chr(QUOTE_CH))
End Function

REM --- Mコード用の列名リスト ~列名~, ~列名~ ... ----------------
Private Function QuotedHeaderList() As String
    Dim hdr As Variant, i As Long, s As String
    hdr = MasterHeaders()
    For i = 0 To UBound(hdr)
        If i > 0 Then s = s & ", "
        s = s & M_QUOTE & hdr(i) & M_QUOTE
    Next i
    QuotedHeaderList = s
End Function

REM --- 既存のシート・接続・クエリを消す ------------------------
Private Sub DropMasterObjects()
    Dim i As Long
    Dim alertSave As Boolean

    alertSave = Application.DisplayAlerts
    Application.DisplayAlerts = False

    On Error Resume Next
    ThisWorkbook.Worksheets(MASTER_SHEET).Delete
    On Error GoTo 0

    For i = ThisWorkbook.Connections.Count To 1 Step -1
        If InStr(1, ThisWorkbook.Connections(i).Name, MASTER_QUERY, vbTextCompare) > 0 Then
            ThisWorkbook.Connections(i).Delete
        End If
    Next i
    For i = ThisWorkbook.Queries.Count To 1 Step -1
        If ThisWorkbook.Queries(i).Name = MASTER_QUERY Then ThisWorkbook.Queries(i).Delete
    Next i

    Application.DisplayAlerts = alertSave
End Sub

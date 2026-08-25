REM =========================================================
REM  ヒート画像 取込モジュール
REM  ダウンロードフォルダなどから画像を選び、データシートのF列に
REM  生成されたファイル名へリネームして画像フォルダへ格納する。
REM ---------------------------------------------------------
REM  AssignHeatImage         選択している行に画像を割り当てる
REM  AssignMissingHeatImages 画像が未配置の行を上から順に処理する
REM  OpenImageFolder         画像フォルダをエクスプローラで開く
REM =========================================================
Option Explicit

REM --- ダイアログ・フィルタ ------------------------------------
Private Const DLG_TITLE     As String = "ヒート画像を選択"
Private Const IMG_FILTER_NM As String = "画像ファイル"
Private Const IMG_FILTER_EX As String = "*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff;*.heic;*.webp"

REM --- ダウンロードフォルダの取得 ------------------------------
REM  既定以外の場所へ移動されている場合があるのでレジストリを優先する。
Private Const REG_DOWNLOADS As String = _
    "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\{374DE290-123F-4565-9164-39C4925E467B}"
Private Const ENV_USERPROFILE As String = "USERPROFILE"
Private Const DIR_DOWNLOADS   As String = "Downloads"
Private Const EXT_SEP         As String = "."

REM --- 動作切替 ------------------------------------------------
REM  True にすると格納後に元ファイルを削除する（コピーではなく移動になる）。
REM  既定は False。元を残しておけば取り違えてもやり直せる。
Private Const MOVE_AFTER_STORE As Boolean = False

REM --- メッセージ ----------------------------------------------
Private Const MSG_NOT_DATA  As String = "データシートを開いて、対象の行を選択してから実行してください。"
Private Const MSG_NO_ROW    As String = "対象の行が選択されていません。明細行を選択してください。"
Private Const MSG_NO_NAME   As String = "ファイル名が未生成の行です。C列のGTINを入力し、検算がOKになってから実行してください。"
Private Const MSG_OVERWRITE As String = "同名のファイルが既にあります。上書きしますか_"
Private Const MSG_EXT_DIFF  As String = "選んだ画像の拡張子が設定 _ 拡張子 と違います。設定側の拡張子で保存します。"
Private Const MSG_NONE_LEFT As String = "画像が未配置の行はありません。"
Private Const MSG_ERR       As String = "画像の格納に失敗しました。"
Private Const MSG_RESULT    As String = "格納_ "
Private Const MSG_SKIP      As String = " 件 / 見送り_ "
Private Const MSG_PLACE     As String = "カード印刷シートに画像を配置しますか_"
Private Const MSG_TOO_LONG  As String = "格納先のパスが長すぎます。画像フォルダを浅い場所へ移してください。文字数_ "
Private Const MSG_NO_DIR    As String = "画像フォルダが未設定です。データシートの設定欄を確認してください。"

REM ==========================================================
REM  選択している行に画像を割り当てる
REM  複数行を選択しておけば、行ごとに順番にピッカーが出る。
REM ==========================================================
Public Sub AssignHeatImage()
    On Error GoTo Fail
    Dim ws As Worksheet
    Dim targetRows As Collection

    Set ws = ThisWorkbook.Worksheets(SHT_DATA)
    If Not ActiveSheet Is ws Then
        MsgBox MSG_NOT_DATA, vbExclamation, DLG_TITLE
        Exit Sub
    End If

    Set targetRows = SelectedDataRows(ws)
    If targetRows.Count = 0 Then
        MsgBox MSG_NO_ROW, vbExclamation, DLG_TITLE
        Exit Sub
    End If

    ProcessRows ws, targetRows
    Exit Sub
Fail:
    MsgBox MSG_ERR & vbCrLf & Err.Number & "_ " & Err.Description, vbCritical, DLG_TITLE
End Sub

REM ==========================================================
REM  画像が未配置の行を上から順に処理する
REM ==========================================================
Public Sub AssignMissingHeatImages()
    On Error GoTo Fail
    Dim ws As Worksheet
    Dim targetRows As Collection
    Dim r As Long, lastRow As Long

    Set ws = ThisWorkbook.Worksheets(SHT_DATA)
    lastRow = DATA_ROW1 + N_ITEMS - 1
    Set targetRows = New Collection

    For r = DATA_ROW1 To lastRow
        If Len(CStr(ws.Cells(r, COL_FILENAME).Value)) > 0 Then
            If Len(Dir(CStr(ws.Cells(r, COL_FULLPATH).Value))) = 0 Then targetRows.Add r
        End If
    Next r

    If targetRows.Count = 0 Then
        MsgBox MSG_NONE_LEFT, vbInformation, DLG_TITLE
        Exit Sub
    End If

    ProcessRows ws, targetRows
    Exit Sub
Fail:
    MsgBox MSG_ERR & vbCrLf & Err.Number & "_ " & Err.Description, vbCritical, DLG_TITLE
End Sub

REM --- 行の集合を順に処理する ----------------------------------
Private Sub ProcessRows(ByVal ws As Worksheet, ByVal targetRows As Collection)
    Dim v As Variant, r As Long
    Dim srcPath As String, destPath As String, msg As String
    Dim okN As Long, skipN As Long
    Dim startDir As String
    Dim extWarned As Boolean
    Dim doStore As Boolean

    startDir = DownloadsFolder()

    For Each v In targetRows
        r = CLng(v)
        destPath = ImagePathFor(ws, r)
        If Len(destPath) = 0 Then
            skipN = skipN + 1
        Else
            srcPath = PickImageFile(startDir, r, CStr(ws.Cells(r, COL_DRUGNAME).Value))
            If Len(srcPath) = 0 Then Exit For   ' キャンセルされたら打ち切る

            REM 次の行は同じフォルダから開く
            startDir = ParentFolder(srcPath)

            doStore = True
            If Len(Dir(destPath)) > 0 Then
                If MsgBox(MSG_OVERWRITE & vbCrLf & destPath, _
                          vbYesNo + vbQuestion, DLG_TITLE) <> vbYes Then doStore = False
            End If
            If doStore And Not extWarned Then
                If LCase$(ExtensionOf(srcPath)) <> LCase$(ExtensionOf(destPath)) Then
                    MsgBox MSG_EXT_DIFF, vbExclamation, DLG_TITLE
                    extWarned = True
                End If
            End If

            If doStore Then
                If StoreHeatImage(ws, r, srcPath, msg) Then
                    okN = okN + 1
                Else
                    skipN = skipN + 1
                End If
            Else
                skipN = skipN + 1
            End If
        End If
    Next v

    MsgBox MSG_RESULT & okN & MSG_SKIP & skipN & " 件", vbInformation, DLG_TITLE
    If okN > 0 Then
        If MsgBox(MSG_PLACE, vbYesNo + vbQuestion, DLG_TITLE) = vbYes Then ImportHeatImages
    End If
End Sub

REM ==========================================================
REM  1行分の画像を格納する
REM  UIを出さないので、テストからも直接呼べる。
REM  戻り値 True で格納成功。msg に格納先または失敗理由を返す。
REM ==========================================================
Public Function StoreHeatImage(ByVal ws As Worksheet, ByVal targetRow As Long, _
                               ByVal srcPath As String, ByRef msg As String) As Boolean
    On Error GoTo Fail
    Dim destPath As String, destDir As String
    Dim fso As Object

    StoreHeatImage = False
    msg = ""

    If Len(NormalizeFolder(CStr(ws.Range(CELL_IMG_DIR).Value))) = 0 Then
        msg = MSG_NO_DIR
        Exit Function
    End If
    destPath = ImagePathFor(ws, targetRow)
    If Len(destPath) = 0 Then
        msg = MSG_NO_NAME
        Exit Function
    End If
    REM Dir も AddPicture も従来パス長までしか扱えない。
    REM ここで弾かないと、格納はできたのに配置で見つからない状態になる。
    If Len(destPath) > MAX_PATH_LEN Then
        msg = MSG_TOO_LONG & Len(destPath)
        Exit Function
    End If
    If Len(Dir(srcPath)) = 0 Then
        msg = srcPath
        Exit Function
    End If

    destDir = ParentFolder(destPath)
    EnsureFolder destDir

    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile srcPath, destPath, True
    If MOVE_AFTER_STORE Then
        If LCase$(srcPath) <> LCase$(destPath) Then fso.DeleteFile srcPath, True
    End If

    msg = destPath
    StoreHeatImage = True
    Exit Function
Fail:
    StoreHeatImage = False
    msg = Err.Number & "_ " & Err.Description
End Function

REM --- その行の画像の格納先フルパス。未生成なら空文字 ----------
Private Function ImagePathFor(ByVal ws As Worksheet, ByVal targetRow As Long) As String
    Dim fileName As String, destDir As String
    fileName = CStr(ws.Cells(targetRow, COL_FILENAME).Value)
    If Len(fileName) = 0 Then Exit Function
    destDir = NormalizeFolder(CStr(ws.Range(CELL_IMG_DIR).Value))
    If Len(destDir) = 0 Then Exit Function
    ImagePathFor = destDir & PATH_SEP & fileName
End Function

REM --- 選択範囲のうち明細行だけを拾う --------------------------
Private Function SelectedDataRows(ByVal ws As Worksheet) As Collection
    Dim c As Collection, area As Range, r As Long
    Dim lastRow As Long
    Set c = New Collection
    lastRow = DATA_ROW1 + N_ITEMS - 1

    If TypeName(Selection) <> "Range" Then
        Set SelectedDataRows = c
        Exit Function
    End If

    For Each area In Selection.Areas
        For r = area.Row To area.Row + area.Rows.Count - 1
            If r >= DATA_ROW1 And r <= lastRow Then
                If Not RowListed(c, r) Then c.Add r
            End If
        Next r
    Next area
    Set SelectedDataRows = c
End Function

Private Function RowListed(ByVal c As Collection, ByVal v As Long) As Boolean
    Dim x As Variant
    For Each x In c
        If CLng(x) = v Then
            RowListed = True
            Exit Function
        End If
    Next x
End Function

REM --- ファイルピッカー。キャンセルなら空文字を返す ------------
Private Function PickImageFile(ByVal startDir As String, ByVal targetRow As Long, _
                               ByVal drugName As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = DLG_TITLE & " - No." & (targetRow - DATA_ROW1 + 1) & " " & drugName
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add IMG_FILTER_NM, IMG_FILTER_EX
        If Len(startDir) > 0 Then .InitialFileName = startDir & PATH_SEP
        If .Show = -1 Then PickImageFile = .SelectedItems(1)
    End With
End Function

REM ==========================================================
REM  ダウンロードフォルダのパスを返す
REM  既定以外へ移動されている場合があるのでレジストリを優先し、
REM  取れなければ %USERPROFILE%\Downloads にフォールバックする。
REM ==========================================================
Public Function DownloadsFolder() As String
    Dim p As String

    On Error Resume Next
    p = CreateObject("WScript.Shell").RegRead(REG_DOWNLOADS)
    On Error GoTo 0

    If Len(p) = 0 Then p = Environ$(ENV_USERPROFILE) & PATH_SEP & DIR_DOWNLOADS
    p = NormalizeFolder(p)
    If Len(Dir(p, vbDirectory)) = 0 Then p = ""
    DownloadsFolder = p
End Function

REM --- 画像フォルダをエクスプローラで開く ----------------------
Public Sub OpenImageFolder()
    On Error GoTo Fail
    Dim p As String
    p = NormalizeFolder(CStr(ThisWorkbook.Worksheets(SHT_DATA).Range(CELL_IMG_DIR).Value))
    EnsureFolder p
    Shell "explorer.exe " & Chr(34) & p & Chr(34), vbNormalFocus
    Exit Sub
Fail:
    MsgBox MSG_ERR & vbCrLf & Err.Number & "_ " & Err.Description, vbCritical, DLG_TITLE
End Sub

REM --- 末尾の区切り文字を落とす --------------------------------
Private Function NormalizeFolder(ByVal p As String) As String
    p = Trim$(p)
    Do While Len(p) > 0
        If Right$(p, 1) <> PATH_SEP Then Exit Do
        p = Left$(p, Len(p) - 1)
    Loop
    NormalizeFolder = p
End Function

Private Function ParentFolder(ByVal filePath As String) As String
    ParentFolder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(filePath)
End Function

REM --- 拡張子。ドットは含めない --------------------------------
Private Function ExtensionOf(ByVal filePath As String) As String
    Dim i As Long
    i = InStrRev(filePath, EXT_SEP)
    If i > 0 Then ExtensionOf = Mid$(filePath, i + 1)
End Function

REM --- フォルダを親から順に作る --------------------------------
Private Sub EnsureFolder(ByVal folderPath As String)
    CreateFolderTree CreateObject("Scripting.FileSystemObject"), NormalizeFolder(folderPath)
End Sub

Private Sub CreateFolderTree(ByVal fso As Object, ByVal p As String)
    Dim parentPath As String
    If Len(p) = 0 Then Exit Sub
    If fso.FolderExists(p) Then Exit Sub
    parentPath = fso.GetParentFolderName(p)
    If Len(parentPath) > 0 Then CreateFolderTree fso, parentPath
    fso.CreateFolder p
End Sub

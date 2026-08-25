REM =========================================================
REM  GS1 データバー限定型 エンコーダ
REM  ISO/IEC 24724 準拠。医療用医薬品の調剤包装単位 ヒート に
REM  印字されているのと同じ symbology。
REM ---------------------------------------------------------
REM  DataBarLimitedBits はワークシート関数として使える。
REM    =DataBarLimitedBits("04987080908647")
REM  戻り値は 74文字のビット列。1=バー, 0=スペース。
REM ---------------------------------------------------------
REM  GTIN-14 の先頭13桁 チェックデジットを除く を値とみなし、
REM  2013571 で割った商と余りを2つのデータ文字に符号化する。
REM  先頭桁 インジケータ は 0 か 1 のみ扱える規格上の制約がある。
REM =========================================================
Option Explicit

Public Const DBL_MODULES As Long = 74      ' 記号本体のモジュール数
Public Const DBL_GTIN_LEN As Long = 14     ' 入力するGTINの桁数

Private Const DBL_VALUE_LEN As Long = 13   ' 符号化対象の桁数
Private Const DBL_DIVISOR   As Long = 2013571
Private Const DBL_ELEMENTS  As Long = 7    ' データ文字あたりの要素数
Private Const DBL_CHK_ELEM  As Long = 6    ' チェック文字あたりの要素数
Private Const DBL_CHK_MOD   As Long = 8    ' チェック文字のモジュール数
Private Const DBL_CHK_MAXW  As Long = 3    ' チェック文字の最大要素幅
Private Const DBL_CHK_DIV   As Long = 21
Private Const DBL_CHECKSUM_MOD As Long = 89
Private Const DBL_LEAD_SPACE   As Long = 1 ' 先頭のスペース。記号本体には含めない
Private Const DBL_TAIL_SPACE   As Long = 5 ' 末尾のスペース。記号本体には含めない

Private Const BIT_BAR   As String = "1"
Private Const BIT_SPACE As String = "0"

REM tab267 の1行 = 上限値, 下限値, 奇モジュール数, 偶モジュール数,
REM                奇最大幅, 偶最大幅, 奇組合せ数, 偶組合せ数
Private Const TAB_COLS As Long = 8
Private Const TAB_ROWS As Long = 7

Private m_tab267 As Variant
Private m_checkWeights As Variant
Private m_checkSeq As Variant
Private m_ready As Boolean

REM ==========================================================
REM  GTIN-14 を GS1データバー限定型 の74モジュール列に符号化する
REM  不正な入力は空文字を返す。数式から使うので Err は上げない。
REM ==========================================================
Public Function DataBarLimitedBits(ByVal gtin14 As String) As String
    On Error GoTo Fail
    Dim binval(0 To 12) As Long
    Dim d1 As Long, d2 As Long
    Dim i As Long, j As Long
    Dim row1 As Variant, row2 As Variant
    Dim d1wo As Variant, d1we As Variant, d2wo As Variant, d2we As Variant
    Dim d1w(0 To 13) As Long, d2w(0 To 13) As Long
    Dim widths(0 To 27) As Long
    Dim checkwidths(0 To 13) As Long
    Dim swidths As Variant, bwidths As Variant
    Dim checksum As Long, seq As Long
    Dim sbs() As Long
    Dim s As String

    DataBarLimitedBits = ""
    gtin14 = Trim$(gtin14)
    If Not IsValidGtin(gtin14) Then Exit Function

    InitTables

    REM --- 先頭13桁を 2013571 で割る。13桁は Long に収まらないので
    REM     BWIPP と同じ桁配列による筆算で商と余りを求める。
    For i = 0 To DBL_VALUE_LEN - 1
        binval(i) = CLng(Mid$(gtin14, i + 1, 1))
    Next i
    For i = 0 To DBL_VALUE_LEN - 2
        binval(i + 1) = binval(i + 1) + (binval(i) Mod DBL_DIVISOR) * 10
        binval(i) = binval(i) \ DBL_DIVISOR
    Next i
    d2 = binval(DBL_VALUE_LEN - 1) Mod DBL_DIVISOR
    binval(DBL_VALUE_LEN - 1) = binval(DBL_VALUE_LEN - 1) \ DBL_DIVISOR
    d1 = 0
    For j = 0 To DBL_VALUE_LEN - 1
        d1 = d1 * 10 + binval(j)
    Next j

    row1 = GroupOf(d1)
    row2 = GroupOf(d2)
    If IsEmpty(row1) Or IsEmpty(row2) Then Exit Function

    REM row = 上限, 下限, 奇モジュール数, 偶モジュール数, 奇最大幅, 偶最大幅, 奇組合せ, 偶組合せ
    d1wo = GetRssWidths((d1 - row1(1)) \ row1(7), row1(2), row1(4), DBL_ELEMENTS, False)
    d1we = GetRssWidths((d1 - row1(1)) Mod row1(7), row1(3), row1(5), DBL_ELEMENTS, True)
    d2wo = GetRssWidths((d2 - row2(1)) \ row2(7), row2(2), row2(4), DBL_ELEMENTS, False)
    d2we = GetRssWidths((d2 - row2(1)) Mod row2(7), row2(3), row2(5), DBL_ELEMENTS, True)

    For i = 0 To DBL_ELEMENTS - 1
        d1w(i * 2) = d1wo(i)
        d1w(i * 2 + 1) = d1we(i)
        d2w(i * 2) = d2wo(i)
        d2w(i * 2 + 1) = d2we(i)
    Next i

    For i = 0 To 13
        widths(i) = d1w(i)
        widths(i + 14) = d2w(i)
    Next i

    checksum = 0
    For i = 0 To 27
        checksum = checksum + widths(i) * m_checkWeights(i)
    Next i
    checksum = checksum Mod DBL_CHECKSUM_MOD
    seq = m_checkSeq(checksum)

    swidths = GetRssWidths(seq \ DBL_CHK_DIV, DBL_CHK_MOD, DBL_CHK_MAXW, DBL_CHK_ELEM, False)
    bwidths = GetRssWidths(seq Mod DBL_CHK_DIV, DBL_CHK_MOD, DBL_CHK_MAXW, DBL_CHK_ELEM, False)

    checkwidths(12) = 1
    checkwidths(13) = 1
    For i = 0 To DBL_CHK_ELEM - 1
        checkwidths(i * 2) = swidths(i)
        checkwidths(i * 2 + 1) = bwidths(i)
    Next i

    REM --- 要素列を組み立てる。偶数番目がバー、奇数番目がスペース ---
    ReDim sbs(0 To 48)
    sbs(0) = 0
    sbs(1) = DBL_LEAD_SPACE
    sbs(2) = 1
    For i = 0 To 13
        sbs(3 + i) = d1w(i)
        sbs(17 + i) = checkwidths(i)
        sbs(31 + i) = d2w(i)
    Next i
    sbs(45) = 1
    sbs(46) = 1
    sbs(47) = DBL_TAIL_SPACE
    sbs(48) = 0

    s = ""
    For i = 0 To UBound(sbs)
        If sbs(i) > 0 Then
            s = s & String$(sbs(i), IIf(i Mod 2 = 0, BIT_BAR, BIT_SPACE))
        End If
    Next i

    REM 先頭のスペースと末尾のスペースを落として記号本体だけを返す
    DataBarLimitedBits = Mid$(s, DBL_LEAD_SPACE + 1, DBL_MODULES)
    Exit Function
Fail:
    DataBarLimitedBits = ""
End Function

REM --- 入力チェック。14桁の数字で先頭が0か1であること ----------
Private Function IsValidGtin(ByVal gtin14 As String) As Boolean
    Dim i As Long, c As String
    If Len(gtin14) <> DBL_GTIN_LEN Then Exit Function
    For i = 1 To DBL_GTIN_LEN
        c = Mid$(gtin14, i, 1)
        If c < "0" Or c > "9" Then Exit Function
    Next i
    If Left$(gtin14, 1) <> "0" And Left$(gtin14, 1) <> "1" Then Exit Function
    IsValidGtin = True
End Function

REM --- 値が属するグループ行を返す ------------------------------
Private Function GroupOf(ByVal d As Long) As Variant
    Dim r As Long, i As Long
    Dim out(0 To TAB_COLS - 1) As Long
    For r = 0 To TAB_ROWS - 1
        If d <= m_tab267(r * TAB_COLS) Then
            For i = 0 To TAB_COLS - 1
                out(i) = m_tab267(r * TAB_COLS + i)
            Next i
            GroupOf = out
            Exit Function
        End If
    Next r
End Function

REM ==========================================================
REM  RSS の要素幅算出
REM  val 値 / nm モジュール総数 / mw 最大要素幅 / el 要素数
REM  oe True で偶数パリティ制約を課す
REM ==========================================================
Private Function GetRssWidths(ByVal val As Long, ByVal nm As Long, ByVal mw As Long, _
                              ByVal el As Long, ByVal oe As Boolean) As Variant
    Dim out() As Long
    Dim bar As Long, ew As Long, mask As Long
    Dim sval As Long, lval As Long, i As Long

    ReDim out(0 To el - 1)
    mask = 0
    For bar = 0 To el - 2
        ew = 1
        mask = mask Or CLng(2 ^ bar)
        Do
            sval = Ncr(nm - ew - 1, el - bar - 2)
            If oe And mask = 0 And (nm - ew - el * 2 + bar * 2) >= -2 Then
                sval = sval - Ncr(nm - ew - el + bar, el - bar - 2)
            End If
            If el - bar > 2 Then
                lval = 0
                For i = nm - ew - el + bar + 2 To mw + 1 Step -1
                    lval = lval + Ncr(nm - i - ew - 1, el - bar - 3)
                Next i
                sval = sval - lval * (el - bar - 1)
            Else
                If nm - ew > mw Then sval = sval - 1
            End If
            val = val - sval
            If val < 0 Then Exit Do
            ew = ew + 1
            mask = mask And Not CLng(2 ^ bar)
        Loop
        val = val + sval
        nm = nm - ew
        out(bar) = ew
    Next bar
    out(el - 1) = nm
    GetRssWidths = out
End Function

REM --- 組合せ nCr。途中で割ることで桁あふれを避ける -------------
Private Function Ncr(ByVal n As Long, ByVal r As Long) As Long
    Dim maxd As Long, mind As Long
    Dim v As Long, j As Long, i As Long

    If r > n - r Then
        maxd = r
        mind = n - r
    Else
        maxd = n - r
        mind = r
    End If
    v = 1
    j = 1
    For i = n To maxd + 1 Step -1
        v = v * i
        If j <= mind Then
            v = v \ j
            j = j + 1
        End If
    Next i
    Do While j <= mind
        v = v \ j
        j = j + 1
    Loop
    Ncr = v
End Function

REM ==========================================================
REM  規格の定数表を用意する
REM  VBA は配列定数を書けないので初回呼び出し時に組み立てる。
REM ==========================================================
Private Sub InitTables()
    If m_ready Then Exit Sub

    m_tab267 = Array( _
        183063, 0, 17, 9, 6, 3, 6538, 28, _
        820063, 183064, 13, 13, 5, 4, 875, 728, _
        1000775, 820064, 9, 17, 3, 6, 28, 6454, _
        1491020, 1000776, 15, 11, 5, 4, 2415, 203, _
        1979844, 1491021, 11, 15, 4, 5, 203, 2408, _
        1996938, 1979845, 19, 7, 8, 1, 17094, 1, _
        2013570, 1996939, 7, 19, 1, 8, 1, 16632)

    m_checkWeights = Array( _
        1, 3, 9, 27, 81, 65, 17, 51, 64, 14, 42, 37, 22, 66, _
        20, 60, 2, 6, 18, 54, 73, 41, 34, 13, 39, 28, 84, 74)

    m_checkSeq = BuildCheckSeq()
    m_ready = True
End Sub

REM --- チェック文字の並び 89通り -------------------------------
Private Function BuildCheckSeq() As Variant
    Dim seq(0 To DBL_CHECKSUM_MOD - 1) As Long
    Dim n As Long, i As Long

    n = 0
    For i = 0 To 43: seq(n) = i: n = n + 1: Next i
    seq(n) = 45: n = n + 1
    seq(n) = 52: n = n + 1
    seq(n) = 57: n = n + 1
    For i = 63 To 66: seq(n) = i: n = n + 1: Next i
    For i = 73 To 79: seq(n) = i: n = n + 1: Next i
    seq(n) = 82: n = n + 1
    For i = 126 To 130: seq(n) = i: n = n + 1: Next i
    seq(n) = 132: n = n + 1
    For i = 141 To 146: seq(n) = i: n = n + 1: Next i
    For i = 210 To 217: seq(n) = i: n = n + 1: Next i
    seq(n) = 220: n = n + 1
    For i = 316 To 320: seq(n) = i: n = n + 1: Next i
    seq(n) = 322: n = n + 1
    seq(n) = 323: n = n + 1
    seq(n) = 326: n = n + 1
    seq(n) = 337: n = n + 1

    BuildCheckSeq = seq
End Function

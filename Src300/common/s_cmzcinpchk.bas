Attribute VB_Name = "s_cmzcinpchk"
Option Explicit

''==========================================
'' 入力チェック関数群
''==========================================


'' 入力チェック関数の戻り値
Public Enum CHK_RESULT
    CHK_OK          '' 正常
    CHK_NG          '' 異常
    CHK_NULL        '' 未入力
End Enum

Public Enum CHK_TYPE
    CHK_NUMBER      '' 数値
    CHK_NUMSTR      '' 数字列
    CHK_STRING      '' 文字列
End Enum

Public Enum CHK_NUMTYPE
    NUMTYPE_ALL         ''+/0/- 全てOK
    NUMTYPE_PLUS        ''+ のみOK
    NUMTYPE_ZEROPLUS    ''+/0 のみOK
End Enum
    

'数値フォーマット(少数点以下は１桁以上で入っているところまで)
'整数のみの場合
Public Const FMT_U0 = "#,##0; ; "
Public Const FMT_M0 = "#,##0;-#,##0; "
'正の数のみ表示する場合
Public Const FMT_U1 = "0.0; ; "
Public Const FMT_U2 = "0.0#; ; "
Public Const FMT_U3 = "0.0##; ; "
Public Const FMT_U4 = "0.0###; ; "
Public Const FMT_U5 = "0.0####; ; "
Public Const FMT_U6 = "0.0#####; ; "
'負の数も表示する場合
Public Const FMT_M1 = "0.0;-0.0; "
Public Const FMT_M2 = "0.0#;-0.0#; "
Public Const FMT_M3 = "0.0##;-0.0##; "
Public Const FMT_M4 = "0.0###;-0.0###; "
Public Const FMT_M5 = "0.0####;-0.0####; "
Public Const FMT_M6 = "0.0#####;-0.0#####; "

'概要      :数値型入力のチェック
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :s             ,I  ,String    ,評価対象文字列
'          :upperLen      ,I  ,Integer   ,整数部桁数
'          :lowerLen      ,I  ,Integer   ,小数点以下桁数
'          :戻り値        ,O  ,CHK_RESULT,チェック結果
'説明      :整数・小数を含めた数値の桁数チェックを行い、正常ならばCHK_OKとする。
'履歴      :2001/06/20(wed) 長野  作成
Public Function ChkNumber(ByVal s$, ByVal upperLen%, ByVal lowerLen%, Optional numType As CHK_NUMTYPE = NUMTYPE_ALL) As CHK_RESULT
Dim Txt_Str     As String               '評価対象文字列
Dim Str_Num()   As String               '対象文字配列
Dim Num         As String
    

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcinpchk.bas -- Function ChkNumber"

    ChkNumber = CHK_NG                      'ステータス初期値セット(error)
    Txt_Str = s                             '引数sセット

    '*** Nullチェック ***
    If Trim$(Txt_Str) = vbNullString Then '引数sがNull値の場合
        ChkNumber = CHK_NULL                'Nullステータス
        GoTo proc_exit
    End If
    
    '*** 引数チェック ***
    If upperLen <= 0 Then                   '整数部桁数が０以下の場合
        GoTo proc_exit
    End If

    '*** 数値チェック ***
    If IsNumeric(Txt_Str) = False Then      '対象文字列が数値でない場合
        GoTo proc_exit
    End If
    If (Right$(Txt_Str, 1) = "+") Or (Right$(Txt_Str, 1) = "-") Then    '「10-」等もisnumeric()は通るので
        GoTo proc_exit
    End If

    '*** 桁数チェック ***
    Str_Num = Split(Txt_Str, ".")           '引数sを小数点区切りで配列にセット
    If Str_Num(0) = "" Then                 '整数部が無い場合
        Num = 0
    Else                                    '整数部がある場合
        Num = Abs(CDbl(Str_Num(0)))         '符号と｢，｣記号を除去
    End If
        '整数部チェック
    If Len(Num) > upperLen Then
        GoTo proc_exit
    End If
        '小数値のチェック
    If UBound(Str_Num) > 0 Then
        If Len(Str_Num(1)) > lowerLen Then
            GoTo proc_exit
        End If
    End If

    '*** 範囲チェック ***
    If numType = NUMTYPE_PLUS Then          '+ のみ可
        If val(Txt_Str) <= 0# Then
            GoTo proc_exit
        End If
    ElseIf numType = NUMTYPE_ZEROPLUS Then  '+/0 のみ可
        If val(Txt_Str) < 0# Then
            GoTo proc_exit
        End If
    End If

    ChkNumber = CHK_OK                      '正常ステータス


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function



'概要      :数字列入力のチェック
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :s             ,I  ,String    ,評価対象文字列
'          :sLen          ,I  ,Integer   ,有効文字数
'          :戻り値        ,O  ,CHK_RESULT,チェック結果
'説明      :有効文字数ちょうどの数字列のみを CHK_OK とする
'履歴      :2001/06/20 作成  野村
Public Function ChkNumStr(ByVal s$, ByVal sLen%) As CHK_RESULT

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcinpchk.bas -- Function ChkNumStr"

    If s = vbNullString Then
        ChkNumStr = CHK_NULL
    ElseIf s Like String(sLen, "#") Then
        ChkNumStr = CHK_OK
    Else
        ChkNumStr = CHK_NG
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :文字列入力のチェック
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型            ,説明
'          :s             ,I  ,String        ,評価対象文字列
'          :suLen         ,I  ,Integer       ,有効文字数(上限)
'          :slLen         ,I  ,Integer       ,有効文字数(下限)
'          :戻り値        ,O  ,CHK_RESULT,チェック結果
'説明      :指定文字数以内なら CHK_OK とする
'履歴      :2001/06/20 作成  野村
Public Function ChkString(ByVal s$, ByVal suLen%, ByVal slLen%) As CHK_RESULT
Dim chkS As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcinpchk.bas -- Function ChkString"

    chkS = StrConv(s, vbFromUnicode)
    If s = vbNullString Then
        ChkString = CHK_NULL
    ElseIf (LenB(chkS) >= slLen) And (LenB(chkS) <= suLen) Then
        ChkString = CHK_OK
    Else
        ChkString = CHK_NG
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function



'概要      :TextBoxの入力内容をチェックする
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :txt           ,I  ,TextBox   ,チェック対象のテキストボックス
'          :chkType       ,I  ,CHK_TYPE  ,入力チェックのタイプ
'          :upperLen      ,I  ,Integer   ,小数点より上の桁数（文字数）
'          :[lowerLen]    ,I  ,Integer   ,小数点より下の桁数
'          :[outFmt]      ,I  ,String    ,表示書式指定
'          :[nullOK]      ,I  ,Boolean   ,Null許可
'          :[numType]     ,I  ,CHK_NUMTYPE   ,数値の有効範囲
'          :戻り値        ,O  ,FUNCTION_RETURN, 入力OK/NG
'説明      :
'履歴      :2001/06/20 作成  野村
Public Function ChkTextBox(txt As TextBox, chkType As CHK_TYPE, upperLen%, Optional lowerLen% = 0, Optional outFmt$ = vbNullString, Optional nullOK As Boolean = False, Optional numType As CHK_NUMTYPE = NUMTYPE_ALL) As FUNCTION_RETURN
Dim chkTxt As String
Dim chkFmt As String
Dim chkResult As CHK_RESULT
Dim RET As FUNCTION_RETURN
    

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcinpchk.bas -- Function ChkTextBox"

    chkTxt = Trim$(txt.Text)
    
    ''chkType に従い、入力チェック関数を呼び出す
    Select Case chkType
      Case CHK_NUMBER
            chkResult = ChkNumber(chkTxt, upperLen, lowerLen, numType)
      Case CHK_NUMSTR
            chkResult = ChkNumStr(chkTxt, upperLen)
      Case CHK_STRING
            chkResult = ChkString(chkTxt, upperLen, lowerLen)
    End Select
    
    ''Nullの許可/不許可を踏まえてチェック結果を評価する
    If nullOK Then
        If chkResult = CHK_NG Then
            RET = FUNCTION_RETURN_FAILURE
        Else
            RET = FUNCTION_RETURN_SUCCESS
        End If
    Else
        If chkResult = CHK_OK Then
            RET = FUNCTION_RETURN_SUCCESS
        Else
            RET = FUNCTION_RETURN_FAILURE
        End If
    End If
    
    'チェック結果によって画面に反映する
    If RET = FUNCTION_RETURN_SUCCESS Then
        ''入力OKなら、テキストボックスの背景色を COLOR_OK に設定する
        txt.BackColor = COLOR_OK
        '書式指定がある場合、整形する
        If outFmt <> vbNullString Then
            txt.Text = Format$(chkTxt, outFmt)
        End If
    Else
        ''入力NGなら、テキストボックスの背景色を COLOR_NG に設定する
        txt.BackColor = COLOR_NG
        ''フォーカスをそのテキストボックスに移す
        txt.SetFocus
    End If
    
    ''チェック結果を返す
    ChkTextBox = RET

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :数値を指定桁までに切り捨てる
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :v             ,I  ,Double    ,元の値
'          :col           ,I  ,Integer   ,結果の小数点以下桁数
'          :戻り値        ,O  ,Double    ,結果
'説明      :負数の切り捨てについての定義は、ExcelのTrunc関数を参考にした
'          :0以下の桁数を指定されたときは、整数値に切り捨てる
'履歴      :2002/03/22 野村 作成
Function RoundDown(ByVal v As Double, ByVal col As Integer) As Double
Dim s As String

    If col <= 0 Then
        RoundDown = Fix(v)
    Else
        s = Format$(Abs(v), "0." & String(col + 1, "0"))
        s = Left$(s, Len(s) - 1)
        RoundDown = Sgn(v) * val(s)
    End If
End Function

'概要      :数値を指定桁までに切り上げる
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :v             ,I  ,Double    ,元の値
'          :col           ,I  ,Integer   ,結果の小数点以下桁数
'          :戻り値        ,O  ,Double    ,結果
'説明      :負数の切上げについての定義は、ExcelのRoundUp関数を参考にした
'          :0以下の桁数を指定されたときは、整数値に切り上げる
'履歴      :2002/03/22 野村 作成
Function RoundUp(ByVal v As Double, ByVal col As Integer) As Double
Dim d As Double

    If col < 0 Then col = 0
    d = Abs(RoundDown(v, col))
    If d < Abs(v) Then
        If col > 0 Then
            RoundUp = Sgn(v) * (d + val("0." & String(col - 1, "0") & "1"))
        Else
            RoundUp = Sgn(v) * (d + 1)
        End If
    Else
        RoundUp = Sgn(v) * d
    End If
End Function

Attribute VB_Name = "s_cmmc001a"
''
'' 抵抗偏析計算画面標準モジュール
''



'概要      :結晶情報を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型             ,説明
'          :tblTarget     ,I   ,typ_TBCME037  ,結晶情報テーブル
'          :strCryNum     ,I   ,String        ,結晶番号
'          :戻り値        ,O   ,FUNCTION_RETURN       ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :結晶情報より、ねらい品番を取得する
Public Function GetRpHinban(tblTarget As typ_TBCME037, strCryNum As String) As FUNCTION_RETURN
    Dim iRet        As Integer
    Dim tblGet()    As typ_TBCME037
    
    GetRpHinban = FUNCTION_RETURN_FAILURE
    
    '' 結晶情報を取得する
    iRet = DBDRV_GetTBCME037(tblGet, "where CRYNUM='" & strCryNum & "'")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function

    tblTarget = tblGet(1)

    GetRpHinban = FUNCTION_RETURN_SUCCESS
End Function


'概要      :引上げ終了実績を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :tblTarget     ,I   ,typ_TBCMH004 ,引上げ終了実績テーブル
'          :strCryNum     ,I   ,String       ,結晶番号
'          :戻り値        ,O   ,FUNCTION_RETURN       ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Public Function GetPlupEndRslt(tblTarget As typ_TBCMH004, strCryNum As String) As FUNCTION_RETURN
    Dim iRet        As Integer
    Dim tblGet()    As typ_TBCMH004
    Dim strCry9     As String * 1
    Dim strCryWork  As String
    
    GetPlupEndRslt = FUNCTION_RETURN_FAILURE

    '' 引上げ終了実績の取得
    iRet = DBDRV_GetTBCMH004(tblGet, "where CRYNUM='" & strCryNum & "'")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function
    tblTarget = tblGet(1)

    '' 結晶番号９桁目を取得（残量引き結晶の確認）
    strCry9 = Mid(strCryNum, 9, 1)
    If (strCry9 <> "") And (InStr(REST_WT_CRYCODE, strCry9) <> 0) Then
        If strCry9 <> "A" Then
            strCryWork = Left(strCryNum, 8) + "A" + Right(strCryNum, 3)
            iRet = DBDRV_GetTBCMH004(tblGet, "where CRYNUM='" & strCryWork & "'")
            If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
            If UBound(tblGet) = 0 Then Exit Function
            tblTarget.CHARGE = tblGet(1).CHARGE
        End If
    End If

    GetPlupEndRslt = FUNCTION_RETURN_SUCCESS
End Function


'概要      :抵抗実績を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :tblRs()       ,O   ,typ_TBCMJ002     ,抵抗実績テーブル配列(1～)
'          :strCryNum     ,I   ,String           ,結晶番号
'          :戻り値        ,O   ,FUNCTION_RETURN   ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Public Function GetResultsRs(tblRs() As typ_TBCMJ002, strCryNum As String) As FUNCTION_RETURN
    Dim iRet            As Integer

    GetResultsRs = FUNCTION_RETURN_FAILURE

    '' 結晶番号より抵抗実績を取得する（位置でソート）
    iRet = DBDRV_GetTBCMJ002(tblRs, _
             " A where CRYNUM='" & strCryNum & "'" & _
             " and TRANCNT=any(select max(TRANCNT) from TBCMJ002" & _
             " where CRYNUM='" & strCryNum & "' and POSITION=A.POSITION)" & _
             " order by POSITION")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblRs) = 0 Then Exit Function

    GetResultsRs = FUNCTION_RETURN_SUCCESS

End Function


'概要      :品番より製品仕様ＳＸＬデータ１を取得、そして、取得した製品仕様データを追加する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                   ,説明
'          :tblTarget     ,O   ,typ_TBCME018        ,製品仕様ＳＸＬデータ１テーブル
'          :tHinInf       ,I   ,tFullHinban         ,品番
'          :戻り値        ,O  ,Integer              ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Public Function GetSPSXLData1(tblTarget As typ_TBCME018, tHinInf As tFullHinban) As Integer
    Dim Index       As Long
    Dim iRet        As Integer
    Dim tblGet()    As typ_TBCME018
    
    GetSPSXLData1 = FUNCTION_RETURN_FAILURE

    '' 製品仕様ＳＸＬデータ１の取得
    iRet = DBDRV_GetTBCME018(tblGet, "where HINBAN='" & tHinInf.HINBAN & "' and MNOREVNO=" & tHinInf.mnorevno & " and FACTORY='" & tHinInf.factory & "' and OPECOND='" & tHinInf.opecond & "'")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function
    
    tblTarget = tblGet(1)
    
    GetSPSXLData1 = FUNCTION_RETURN_SUCCESS

End Function


'概要      :抵抗の値を表示用に文字列化する(指定の小数点以下桁数)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :rs            ,I  ,Double    ,抵抗値
'          :place         ,I  ,Integer   ,小数点以下桁数
'説明      :抵抗値表示桁数を統一するため。<0のときは空文字列を返す
'履歴      :2002/1/16 作成  野村 (2002/07 s_cmzc020a.basより移動)
'履歴      :2002/1/17 S.Sano
Public Function toRsStrByPlace(rs As Double, place As Integer) As String
Dim s$

    If rs < 0 Then
        s = vbNullString
'2002/01/17 S.Sano    ElseIf rs >= 99999.9 Then
'2002/01/17 S.Sano        s = "99999.9"
    Else
        s = Format$(rs, "0." & String(place, "0"))
        If Val(s) >= 100000 Then
            s = "99999." & String(place, "9")
        End If
    End If
    toRsStrByPlace = s
End Function

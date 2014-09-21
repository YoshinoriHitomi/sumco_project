Attribute VB_Name = "s_cmbc053"
'概要      :ＥＰＤ実績を登録する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :tblEPD        ,I  ,typ_TBCMJ001     ,ＥＰＤ実績テーブル
'          :戻り値        ,O  ,Integer          ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Public Function InsertTbl_EPD(tblEPD As typ_TBCMJ001) As Integer
    Dim iRet       As Integer
    Dim tblTarget  As typ_cmjc001i_Disp
    
    InsertTbl_EPD = FUNCTION_RETURN_FAILURE

    '' データ形式の変換
    ConvDate_F_cmjc001i_a tblEPD, tblTarget, True
    ''ＥＰＤ実績に登録
    iRet = DBDRV_Getcmjc001i_Exec(tblTarget, tblEPD.CRYNUM, tblEPD.TSTAFFID)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function

    InsertTbl_EPD = FUNCTION_RETURN_SUCCESS
End Function

'概要      :ＥＰＤ実績を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :tblEPD        ,O   ,typ_TBCMJ001      ,ＥＰＤ実績テーブル
'          :strCryNum     ,I   ,String           ,結晶番号
'          :iSmpNo        ,I   ,Long             ,サンプルNo.   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :iIngotPos     ,I   ,Integer          ,結晶内位置
'          :strSmpKbn     ,I   ,String           ,サンプル区分
'          :戻り値        ,O   ,Integer          ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Public Function GetResultsEPD(tblEPD As typ_TBCMJ001, strCryNum As String, iSmpNo As Long, iIngotpos As Integer, strSmpKbn As String) As Integer
    Dim iRet        As Integer
    Dim tblGetEPD() As typ_TBCMJ001

    GetResultsEPD = FUNCTION_RETURN_FAILURE

    '' ＥＰＤ実績の取得
    iRet = DBDRV_GetTBCMJ001(tblGetEPD, _
             "A where CRYNUM='" & strCryNum & "' and POSITION=" & iIngotpos & _
             " and TRANCNT=any(select max(TRANCNT) from TBCMJ001 where CRYNUM='" & strCryNum & "' and POSITION=" & iIngotpos & _
             " and SMPKBN=A.SMPKBN" & _
             ")", _
             " order by POSITION, SMPKBN" & IIf(strSmpKbn = "B", "", " desc"))
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGetEPD) = 0 Then Exit Function

    tblEPD = tblGetEPD(1)
    
    GetResultsEPD = FUNCTION_RETURN_SUCCESS

End Function
'Akizuki <<<<<XSDC1代替処理として作成中>>>>>
'概要      :結晶番号より品番管理テーブルを取得する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :tblHinban()   ,O   ,typ_TBCME041 ,品番管理テーブル
'          :strCryNum     ,I   ,String           ,結晶番号
'          :戻り値        ,O  ,Integer          ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Public Function GetHinban_X(tblHinban() As typ_TBCME041, strCryNum As String) As Integer
    Dim iRet        As Integer
    Dim tblGet()    As typ_TBCME041
    Dim Index       As Integer
    Dim tblPlup     As typ_TBCMH004

    GetHinban = FUNCTION_RETURN_FAILURE

    '' 品番管理テーブルを初期化
    RemoveAll_HinbanManage tblHinban

    '' 品番管理テーブルの取得
    iRet = DBDRV_GetTBCME041(tblGet, "where CRYNUM='" & strCryNum & "' ", "order by INGOTPOS")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then
        If Len(strCryNum) <> 12 Then Exit Function
        '' 品番管理テーブル０の場合、結晶引上失敗しているものとする
        '' 結晶は、Ｚ品番の結晶とする
        '' 引上長を取得
        If GetPlupEndRslt(tblPlup, strCryNum) <> FUNCTION_RETURN_SUCCESS Then Exit Function
        ReDim tblGet(1)
        With tblGet(1)
            .CRYNUM = strCryNum
            .INGOTPOS = 0
            .hinban = "Z"
            .REVNUM = 0
            .Factory = vbNullString
            .OpeCond = vbNullString
            .Length = tblPlup.LENGFREE
        End With
    End If

    For Index = 1 To UBound(tblGet)
        If Add_HinbanManage(tblHinban, tblGet(Index)) <> FUNCTION_RETURN_SUCCESS Then
            Exit Function
        End If
    Next Index

    If UBound(tblHinban) <= 0 Then
        Exit Function
    End If

    GetHinban = FUNCTION_RETURN_SUCCESS

End Function

'概要      :品番より結晶内側管理を取得する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                           ,説明
'          :tblData        ,O  ,typ_TBCME036               ,結晶内側管理テーブル
'          :tHinInf       ,I  ,tFullHinban                 ,品番
'          :戻り値        ,O  ,Integer                      ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Public Function GetSXLInsideSpecManager(tblData As typ_TBCME036, tHinInf As tFullHinban) As Integer
    Dim Index       As Long
    Dim iRet        As Integer
    Dim tblGet()    As typ_TBCME036
    
    GetSXLInsideSpecManager = FUNCTION_RETURN_FAILURE

    '' 結晶内側管理の取得
'    iRet = DBDRV_GetTBCME036(tblGet, _
'             "where HINBAN='" & tHinInf.hinban & "' and MNOREVNO=" & tHinInf.mnorevno & " and FACTORY='" & tHinInf.factory & "' and OPECOND='" & tHinInf.opecond & "'")
    '06/04/11 ooba
    iRet = DBDRV_GetTBCME036_cmbc028(tblGet, _
             "where HINBAN='" & tHinInf.hinban & "' and MNOREVNO=" & tHinInf.mnorevno & " and FACTORY='" & tHinInf.Factory & "' and OPECOND='" & tHinInf.OpeCond & "'")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function

    tblData = tblGet(1)
    
    GetSXLInsideSpecManager = FUNCTION_RETURN_SUCCESS

End Function
Public Function Add_EPDRslt(tblTarget() As typ_TBCMJ001, tblDat As typ_TBCMJ001, Optional Index As Long = -1) As Integer
    Dim tblIndex As Long
    
    Add_EPDRslt = FUNCTION_RETURN_FAILURE

    '' データの追加・更新チェック
    If Index > -1 Then
        '' データ更新の場合
        tblIndex = Index
        If Index > UBound(tblTarget) Then
            '' 更新データ位置インデックス範囲が無効の場合、エラー終了
            Exit Function
        End If
    Else
        '' データ追加の場合
        '' テーブルデータ格納領域拡張
        ReDim Preserve tblTarget(UBound(tblTarget) + 1)
        '' テーブルデータ数を取得
        tblIndex = UBound(tblTarget) - 1
    End If

    '' データ追加
    tblTarget(tblIndex) = tblDat

    Add_EPDRslt = FUNCTION_RETURN_SUCCESS
End Function

'概要      :データ変換を行う
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                ,説明
'          :tblLeft       ,IO   ,typ_TBCMJ001      ,テーブルデータ１
'          :tblRight      ,IO   ,typ_cmjc001i_Disp ,テーブルデータ２
'          :bFlg          ,I   ,Boolean           ,TRUE:引数１データ→引数２データへの変換  FALSE:引数１データ←引数２データへの変換
'説明      :
Public Sub ConvDate_F_cmjc001i_a(tblLeft As typ_TBCMJ001, tblRight As typ_cmjc001i_Disp, bFlg As Boolean)
    If bFlg = True Then
        With tblRight
            .POSITION = tblLeft.POSITION
            .SMPKBN = tblLeft.SMPKBN
            .TRANCOND = tblLeft.TRANCOND
            .SMPLNO = tblLeft.SMPLNO
            .SMPLUMU = tblLeft.SMPLUMU
            .KRPROCCD = tblLeft.KRPROCCD
            .PROCCODE = tblLeft.PROCCODE
            .hinban = tblLeft.hinban
            .REVNUM = tblLeft.REVNUM
            .Factory = tblLeft.Factory
            .OpeCond = tblLeft.OpeCond
            .GOUKI = tblLeft.GOUKI
            .MEASURE = tblLeft.MEASURE
        End With
    Else
        With tblLeft
            .POSITION = tblRight.POSITION
            .SMPKBN = tblRight.SMPKBN
            .TRANCOND = tblRight.TRANCOND
            .SMPLNO = tblRight.SMPLNO
            .SMPLUMU = tblRight.SMPLUMU
            .KRPROCCD = tblRight.KRPROCCD
            .PROCCODE = tblRight.PROCCODE
            .hinban = tblRight.hinban
            .REVNUM = tblRight.REVNUM
            .Factory = tblRight.Factory
            .OpeCond = tblRight.OpeCond
            .GOUKI = tblRight.GOUKI
            .MEASURE = tblRight.MEASURE
        End With
    End If

End Sub

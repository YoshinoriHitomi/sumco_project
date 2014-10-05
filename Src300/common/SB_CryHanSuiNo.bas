Attribute VB_Name = "SB_CryHanSuiNo"
Option Explicit

'検索データ構造体
Public Type typ_SeekData
    BLOCKID As String * 12      ' ブロックＩＤ
    TBKBN   As String * 1       ' T/B区分(T:Top, B:Bot)
    INPOS   As Integer          ' 結晶内位置
    HINB    As tFullHinban      ' 品番(構造体)
    IND     As String           ' 状態FLG(0:検査無, 1:通常, 2:反映, 3:推定)
End Type


'------------------------------------------------
' 結晶反映/推定チェック(実績なし)共通関数
'------------------------------------------------

'概要      :指定された評価項目№により、反映か推定かを判断し、結晶反映チェック(実績なし)、または、結晶推定チェック(実績なし)を呼び出す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :sBlockid      ,I  ,String       :ﾌﾞﾛｯｸID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :結晶番号
'          :tFullHin      ,I  ,tFullHinban  :品番(構造体)
'          :iSmplPos      ,I  ,Integer      :新サンプル位置(mm)
'          :iItemNo       ,I  ,Integer      :評価項目№ =  1 RS     ← 推定偏析計算
'                                                       =  2 Oi     ← ﾊﾟﾀｰﾝ1
'                                                       =  3 BMD1   ← ﾊﾟﾀｰﾝ1
'                                                       =  4 BMD2   ← ﾊﾟﾀｰﾝ1
'                                                       =  5 BMD3   ← ﾊﾟﾀｰﾝ1
'                                                       =  6 OSF1   ← ﾊﾟﾀｰﾝ1
'                                                       =  7 OSF2   ← ﾊﾟﾀｰﾝ1
'                                                       =  8 OSF3   ← ﾊﾟﾀｰﾝ1
'                                                       =  9 OSF4   ← ﾊﾟﾀｰﾝ1
'                                                       = 10 CS     ← ﾊﾟﾀｰﾝ2(上限値,下限値共0より大(0<)の場合,ﾊﾟﾀｰﾝ1)
'                                                       = 11 GD     ← ﾊﾟﾀｰﾝ1
'                                                       = 12 LT     ← ﾊﾟﾀｰﾝ3
'                                                       = 13 EPD    ← ﾊﾟﾀｰﾝ2
'          :tSeekData()   ,I  ,typ_SeekData :検索データ構造体配列
'          :iHanSuiKBN    ,O  ,Integer      :反映/推定区分(0:反映,1:推定)
'          :iGetDataNo1   ,O  ,Integer      :推定元の配列番号１
'          :iGetDataNo2   ,O  ,Integer      :推定元の配列番号２(反映時未使用)
'          :戻り値        ,O  ,Integer      :チェック結果 = 0 : 正常終了(反映/推定OK)
'                                                           1 : 正常終了(反映/推定NG)
'                                                          -1 : 入力引数値エラー
'                                                          -2 : 上記以外のエラー
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funChkSxlHanSuiNo(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                                  iItemNo As Integer, tSeekData() As typ_SeekData, iHanSuiKBN As Integer, _
                                  iGetDataNo1 As Integer, iGetDataNo2 As Integer) As Integer
    Dim retCode As Integer
    
    'パラメータチェック
    If (Len(sBlockId) <> 12) Then GoTo ChkSxlHanSuiNoParameterErr
    If (Len(sCryNum) <> 12) Then GoTo ChkSxlHanSuiNoParameterErr
    If UBound(tSeekData) = 0 Then GoTo ChkSxlHanSuiNoParameterErr
    
    '指定された評価項目№により、反映か推定かを判断し、結晶反映チェック、または、結晶推定チェックを呼び出す。
    Select Case iItemNo
    Case 1              'RS(比抵抗)
        retCode = funChkSxlSuiteiNo(sBlockId, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, tSeekData(), iGetDataNo1, iGetDataNo2)
        iHanSuiKBN = 1
'    Case 2              'Oi(酸素濃度)
'    Case 3              'BMD1
'    Case 4              'BMD2
'    Case 5              'BMD3
'    Case 6              'OSF1
'    Case 7              'OSF2
'    Case 8              'OSF3
'    Case 9              'OSF4
'    Case 10             'CS(炭素濃度)
'    Case 11             'GD
'    Case 12             'LT(ﾗｲﾌﾀｲﾑ)
'    Case 13             'EPD
    Case Else
        GoTo ChkSxlHanSuiNoParameterErr
    End Select
    
    '共通関数のチェック結果を当関数の結果として、呼び出し元へ返す。
    funChkSxlHanSuiNo = retCode
    Exit Function

ChkSxlHanSuiNoParameterErr:
    funChkSxlHanSuiNo = -1
    Exit Function

ChkSxlHanSuiNoSonotaErr:
    funChkSxlHanSuiNo = -2
End Function

'------------------------------------------------
' 結晶推定チェック(実績なし)
'------------------------------------------------

'概要      :指定された情報から、結晶推定チェック(実績なし)を行ない結果を返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :sBlockid      ,I  ,String       :ﾌﾞﾛｯｸID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :結晶番号
'          :tFullHin      ,I  ,tFullHinban  :品番(構造体)
'          :iSmplPos      ,I  ,Integer      :新サンプル位置(mm)
'          :iItemNo       ,I  ,Integer      :評価項目№ =  1 RS     ←対象
'                                                       =  2 Oi
'                                                       =  3 BMD1
'                                                       =  4 BMD2
'                                                       =  5 BMD3
'                                                       =  6 OSF1
'                                                       =  7 OSF2
'                                                       =  8 OSF3
'                                                       =  9 OSF4
'                                                       = 10 CS
'                                                       = 11 GD
'                                                       = 12 LT
'                                                       = 13 EPD
'          :tSeekData()   ,I  ,typ_SeekData :検索データ構造体配列
'          :iGetDataNo1   ,O  ,Integer      :推定元の配列番号１
'          :iGetDataNo2   ,O  ,Integer      :推定元の配列番号２
'          :戻り値        ,O  ,Integer      :チェック結果 = 0 : 正常終了(推定OK)
'                                                           1 : 正常終了(推定NG)
'                                                          -1 : 入力引数値エラー
'                                                          -2 : 上記以外のエラー
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funChkSxlSuiteiNo(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                                  iItemNo As Integer, tSeekData() As typ_SeekData, iGetDataNo1 As Integer, iGetDataNo2 As Integer) As Integer
'    Dim tSiyou          As type_DBDRV_scmzc_fcmkc001c_Siyou
    Dim wGetDataNoTop   As Integer
    Dim wGetDataNoBot   As Integer
    
    'パラメータチェック
    If (Len(sBlockId) <> 12) Then GoTo ChkSxlSuiteiNoParameterErr
    If (Len(sCryNum) <> 12) Then GoTo ChkSxlSuiteiNoParameterErr
    If UBound(tSeekData) = 0 Then GoTo ChkSxlSuiteiNoParameterErr
    
    '指定された評価項目№毎に必要な品番仕様値を取得する。（指定された評価項目№により、処理が分かれる。）
    Select Case iItemNo
    Case 1              'RS(比抵抗)
'        If funGet_TBCME018(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlSuiteiNoNG
    Case 2 To 13        'Oi(酸素濃度),BMD1,BMD2,BMD3,OSF1,OSF2,OSF3,OSF4,CS(炭素濃度),GD,LT(ﾗｲﾌﾀｲﾑ),EPD
        GoTo ChkSxlSuiteiNoNG
    Case Else
        GoTo ChkSxlSuiteiNoParameterErr
    End Select

    '結晶推定元の配列番号の取得（結晶推定元の配列番号が取得できたら推定ＯＫとする。）
    If funGetSuiteiNo(sBlockId, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, tSeekData(), _
                                            wGetDataNoTop, wGetDataNoBot) <> 0 Then GoTo ChkSxlSuiteiNoNG

    '取得配列番号と戻り値の設定
    iGetDataNo1 = wGetDataNoTop
    iGetDataNo2 = wGetDataNoBot
    
    funChkSxlSuiteiNo = 0
    Exit Function

ChkSxlSuiteiNoNG:
    funChkSxlSuiteiNo = 1
    Exit Function

ChkSxlSuiteiNoParameterErr:
    funChkSxlSuiteiNo = -1
    Exit Function

ChkSxlSuiteiNoSonotaErr:
    funChkSxlSuiteiNo = -2
End Function

'------------------------------------------------
' 結晶推定値取得(実績なし)
'------------------------------------------------

'概要      :指定された新ｻﾝﾌﾟﾙ位置情報から、結晶推定元位置１と結晶推定元位置２を検索データ構造体配列より検索し、それぞれの配列番号を返す。
'           推定元位置１と推定元位置２を検索する場合、新サンプル位置から最も近いTOPﾃﾞｰﾀと最も近いBOTﾃﾞｰﾀの位置が対象となる。
'           新ｻﾝﾌﾟﾙ位置の品番仕様とTOP／BOT位置の品番仕様が、それぞれ「3点測定」「5点測定」でのﾊﾟﾀｰﾝが考えられるが、このﾊﾟﾀｰﾝによって推定可否を判断する。
'           推定可否の判断として、XSDC1のSUIFLGC1の値(0:推定許可,1:推定禁止)も考慮する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :sBlockid      ,I  ,String       :ﾌﾞﾛｯｸID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :結晶番号
'          :tFullHin      ,I  ,tFullHinban  :品番(構造体)
'          :iSmplPos      ,I  ,Integer      :新サンプル位置(mm)
'          :iItemNo       ,I  ,Integer      :評価項目№ =  1 RS     ←対象
'                                                       =  2 Oi
'                                                       =  3 BMD1
'                                                       =  4 BMD2
'                                                       =  5 BMD3
'                                                       =  6 OSF1
'                                                       =  7 OSF2
'                                                       =  8 OSF3
'                                                       =  9 OSF4
'                                                       = 10 CS
'                                                       = 11 GD
'                                                       = 12 LT
'                                                       = 13 EPD
'          :tSeekData()   ,I  ,typ_SeekData :検索データ構造体配列
'          :iGetDataNo1   ,O  ,Integer      :推定元の配列番号１
'          :iGetDataNo2   ,O  ,Integer      :推定元の配列番号２
'          :戻り値        ,O  ,Integer      :取得結果 = 0 : 正常終了
'                                                       1 : 正常終了(該当サンプルなし)
'                                                      -1 : 異常終了
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funGetSuiteiNo(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                               iItemNo As Integer, tSeekData() As typ_SeekData, iGetDataNo1 As Integer, iGetDataNo2 As Integer) As Integer
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim wFind       As Boolean
    
    Dim getNewSpec  As String       '新ｻﾝﾌﾟﾙ位置比抵抗仕様値
    
    Dim getTopNo    As Integer      'TOP位置配列番号
    Dim getTopBlkID As String       'TOP位置ﾌﾞﾛｯｸID
    Dim getTopTB    As String       'TOP位置T/B区分
    Dim getTopHin   As tFullHinban  'TOP位置品番
    Dim getTopSpec  As String       'TOP位置比抵抗仕様値
    Dim getTopPtrn  As String       'TOP位置ﾊﾟﾀｰﾝｺｰﾄﾞ
    
    Dim getBotNo    As Integer      'BOT位置配列番号
    Dim getBotBlkID As String       'BOT位置ﾌﾞﾛｯｸID
    Dim getBotTB    As String       'BOT位置T/B区分
    Dim getBotHin   As tFullHinban  'BOT位置品番
    Dim getBotSpec  As String       'BOT位置比抵抗仕様値
    Dim getBotPtrn  As String       'BOT位置ﾊﾟﾀｰﾝｺｰﾄﾞ
    
    'パラメータチェック
    If (Len(sBlockId) <> 12) Then GoTo GetSuiteiNoParameterErr
    If (Len(sCryNum) <> 12) Then GoTo GetSuiteiNoParameterErr
    If UBound(tSeekData) = 0 Then GoTo GetSuiteiNoParameterErr
    
    '推定可否の判断 → XSDC1のSUIFLGC1の値(0:推定許可,1:推定禁止)
''    sql = "select SUIFLGC1 from XSDC1 where XTALC1 = '" & sCryNum & "'"       2003/10/14
    sql = "select SUIFLG from XSDC1 where XTALC1 = '" & sCryNum & "'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
''    If rs.EOF Or rs("SUIFLGC1") <> "0" Then           2003/10/14
    If rs.EOF Or rs("SUIFLG") <> "0" Then
        Set rs = Nothing
        GoTo GetSuiteiNoEmpty
    End If
    Set rs = Nothing
    
    '指定された新サンプル位置情報で検索データ構造体配列(入力)を検索し、推定元の配列番号を決定する。
    '≪推定元配列番号１(TOP位置)の取得≫
    wFind = False
'    For getTopNo = UBound(tSeekData) To 0 Step -1
'        If (tSeekData(getTopNo).TBKBN = "T") And (tSeekData(getTopNo).INPOS < iSmplPos) Then
    For getTopNo = 0 To UBound(tSeekData)
        If (tSeekData(getTopNo).TBKBN = "T") And (tSeekData(getTopNo).INPOS < iSmplPos) And (tSeekData(getTopNo).IND = "1") Then
            wFind = True
            Exit For
        End If
    Next getTopNo
'    getTopNo = 0                            '推定元TOPは、先頭位置とする
'    wFind = True
    
    If wFind = False Then GoTo GetSuiteiNoEmpty

    'TOP位置データの設定
    With tSeekData(getTopNo)
        getTopBlkID = .BLOCKID                  'TOP位置ﾌﾞﾛｯｸID
        getTopTB = .TBKBN                       'TOP位置T/B区分
        getTopHin.hinban = .HINB.hinban         'TOP位置品番
        getTopHin.mnorevno = .HINB.mnorevno     'TOP位置製品番号改訂番号
        getTopHin.factory = .HINB.factory       'TOP位置工場
        getTopHin.opecond = .HINB.opecond       'TOP位置操業条件
    End With

    '≪推定元配列番号２(BOT位置)の取得≫
    wFind = False
'    For getBotNo = 0 To UBound(tSeekData)
'        If (tSeekData(getBotNo).TBKBN = "B") And (tSeekData(getBotNo).INPOS > iSmplPos) Then
    For getBotNo = UBound(tSeekData) To 0 Step -1
        If (tSeekData(getBotNo).TBKBN = "B") And (tSeekData(getBotNo).INPOS > iSmplPos) And (tSeekData(getBotNo).IND = "1") Then
            wFind = True
            Exit For
        End If
    Next getBotNo
'    getBotNo = UBound(tSeekData)            '推定元BOTは、最終位置とする
'    wFind = True
    
    If wFind = False Then GoTo GetSuiteiNoEmpty

    'BOT位置データの設定
    With tSeekData(getBotNo)
        getBotBlkID = .BLOCKID                  'Bot位置ﾌﾞﾛｯｸID
        getBotTB = .TBKBN                       'Bot位置T/B区分
        getBotHin.hinban = .HINB.hinban         'Bot位置品番
        getBotHin.mnorevno = .HINB.mnorevno     'Bot位置製品番号改訂番号
        getBotHin.factory = .HINB.factory       'Bot位置工場
        getBotHin.opecond = .HINB.opecond       'Bot位置操業条件
    End With
    
    '各品番の比抵抗仕様値取得
    '≪指定された新サンプル位置≫
    getNewSpec = funGetSuiSpecRS(tFullHin)
    If getNewSpec = " " Then GoTo GetSuiteiNoEmpty
    
    '≪推定元サンプルＩＤ１(TOP位置)≫
    getTopSpec = funGetSuiSpecRS(getTopHin)
    If getTopSpec = " " Then GoTo GetSuiteiNoEmpty
    
    '≪推定元サンプルＩＤ２(BOT位置)≫
    getBotSpec = funGetSuiSpecRS(getBotHin)
    If getBotSpec = " " Then GoTo GetSuiteiNoEmpty

    'コードDB取得関数を呼び出し､コードテーブルから比抵抗推定パターンコードを取得する｡
    '≪推定元サンプルＩＤ１ ⇒ 新サンプル位置≫
    If funCodeDBGet("SB", "ST", getTopSpec, 1, getNewSpec, getTopPtrn) <> 0 Then GoTo GetSuiteiNoParameterErr
    If getTopPtrn <> "A" And getTopPtrn <> "B" Then GoTo GetSuiteiNoEmpty
    
    '≪推定元サンプルＩＤ２ ⇒ 新サンプル位置≫
    If funCodeDBGet("SB", "ST", getBotSpec, 1, getNewSpec, getBotPtrn) <> 0 Then GoTo GetSuiteiNoParameterErr
    If getBotPtrn <> "A" And getTopPtrn <> "B" Then GoTo GetSuiteiNoEmpty
    
    '呼び出し元への結果通知
    iGetDataNo1 = getTopNo          '推定元の配列番号１
    iGetDataNo2 = getBotNo          '推定元の配列番号２
    
    funGetSuiteiNo = 0
    Exit Function

GetSuiteiNoEmpty:
    funGetSuiteiNo = 1
    Exit Function

GetSuiteiNoParameterErr:
    funGetSuiteiNo = -1
End Function

'------------------------------------------------
' 結晶推定 比抵抗仕様値取得関数
'------------------------------------------------

'概要      :指定された品番から、TBCME018を検索し、比抵抗仕様値(品SX比抵抗測定位置_位)を取得し返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :tFullHin      ,I  ,tFullHinban  :品番(構造体)
'          :戻り値        ,O  ,Sting        :比抵抗仕様値(品SX比抵抗測定位置_位)
'                                            (取得できない場合は、空白を返す)
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Private Function funGetSuiSpecRS(tFullHin As tFullHinban) As String
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    
    '指定された品番からTBCME018のHSXRSPOI(品SX比抵抗測定位置_位)を検索する。
    sql = "select HSXRSPOI from TBCME018 "
    sql = sql & "where HINBAN = '" & tFullHin.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tFullHin.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tFullHin.factory & "' and "
    sql = sql & "      OPECOND = '" & tFullHin.opecond & "'"
    
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetSuiSpecRS = " "
        Set rs = Nothing
        Exit Function
    End If
    
    'TOP位置データの設定
    funGetSuiSpecRS = rs("HSXRSPOI")
    Set rs = Nothing

End Function

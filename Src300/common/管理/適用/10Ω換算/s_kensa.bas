Attribute VB_Name = "s_kensa"
''
'' 構造体、定数定義（結晶検査関連）
''

Option Explicit

'' コンボボックス選択文字列
Public Const SP_COMBO_UMU1 = "有" + vbTab + "無"                    '' （有／無）
Public Const SP_COMBO_UMU2 = "有" + vbTab + "無" + vbTab + "NDF"    '' （有／無／ＮＤＦ）
Public Const SP_COMBO_PN = "P" + vbTab + "N"                        '' （P／N）
'Add Start 2011/01/05 SMPK Miyata
Public Const SP_COMBO_PTN1 = "0:None" + vbTab + "1:Ring" + vbTab + "2:Disk" + vbTab + "3:DiskRing"
Public Const SP_COMBO_PTN2 = "0:None" + vbTab + "5:PB-band" + vbTab + "6:P-band" + vbTab + "7:B-band"
'Add End   2011/01/05 SMPK Miyata


'' 判定結果文字列
Public Const STR_JUDG_OK = "○"   '' 判定OK
Public Const STR_JUDG_NG = "×"   '' 判定NG

'' 酸素炭素検査方法コード定義
Public Const CODE_INSPECTWAY_FRIR = "CA"    '' FRIR
Public Const CODE_INSPECTWAY_FTIR = "CD"    '' FTIR
Public Const CODE_INSPECTWAY_SIMS = "CS"    '' SIMS
Public Const CODE_INSPECTWAY_GFA = "CG"     '' GFA

'' デフォルト値定義
Public Const DEF_PARAM_VALUE = -1                       '' デフォルト値
Public Const DEF_PARAM_DATE = "1901/01/01 01:00:01"     '' デフォルト日付

'' 検査入力可能検査指示コード
Public Const CODE_KENSA = "1234"


'' G品番デフォルト値
'' 抵抗
Public Const DEFCODE_G_POS_RS = "Q3A"   '' 測定位置＿方点位
Public Const DEFCODE_G_GUA_RS = "3S"    '' 保証方法＿対処
'' Ｃｓ
Public Const DEFCODE_G_POS_CS = "Q1Y"   '' 測定位置＿方点位
Public Const DEFCODE_G_GUA_CS = "1S"    '' 保証方法＿対処
'' Ｏｉ（測定位置＿方、＿位は、ねらい品番の仕様を使用する）
Public Const DEFCODE_G_POS_OI = " 3 "   '' 測定位置＿方点位
Public Const DEFCODE_G_GUA_OI = "3S"    '' 保証方法＿対処
'' ライフタイム
Public Const DEFCODE_G_POS_LT = "Q5J"   '' 測定位置＿方点位
Public Const DEFCODE_G_GUA_LT = "BS"    '' 保証方法＿対処

'' Z品番デフォルト値
'' 抵抗
Public Const DEFCODE_Z_POS_RS = "Q3A"   '' 測定位置＿方点位
Public Const DEFCODE_Z_GUA_RS = "3S"    '' 保証方法＿対処
'' Ｃｓ
Public Const DEFCODE_Z_POS_CS = "Q1Y"   '' 測定位置＿方点位
Public Const DEFCODE_Z_GUA_CS = "1S"    '' 保証方法＿対処


'' 熱処理法情報管理テーブル
Public Type typ_HeatInfo
    iHeatClass      As Integer  '' 熱処理分類（熱処理条件の番号に対応。例：BMD1→1）
    strHeatProc     As String   '' 熱処理方法（熱処理法コード）
End Type

'' 熱処理条件コンボボックス管理テーブル
Public Type typ_cmbTInfo
    iHeatClass      As Integer  '' 熱処理分類（熱処理条件の番号に対応。例：BMD1→1）
    strHeatProc     As String   '' 熱処理方法（熱処理法コード）
End Type


'' ユーザデータ管理テーブル定義
Public tbl_HeatInfo() As typ_HeatInfo   '' 熱処理法情報管理テーブル
Public tbl_cmbTInfo() As typ_cmbTInfo   '' 熱処理条件コンボボックス管理テーブル
''
'' ＤＢアクセスモジュール（結晶検査関連）
''

'' メッセージコード定義
Public Const MSG_NOTFOUND_STAFFID = "ESTAF" '' 担当者コードエラー
'Public Const MSG_NOTFOUND_CRYNUM = "ECRY0"  '' 結晶番号エラー
Public Const MSG_NOTFOUND_SMPLNO = "ENSMP"  '' サンプルNO.エラー
Public Const MSG_NOTFOUND_GOUKI_ = "ENGOK"  '' 号機エラー
Public Const MSG_INPUT_STAFFID = "EISTF"    '' 担当者コードを入力してください。
Public Const MSG_INPUT_CRYNUM = "EICRY"     '' 結晶番号を入力してください。
Public Const MSG_INPUT_SMPLNO = "EISMP"     '' サンプルNo.を入力してください。
Public Const MSG_INPUT_GOUKI = "EIGOK"      '' 号機を入力してください。
Public Const MSG_ERROR_PARAM = "EINPM"      '' 入力値が不正です。
Public Const MSG_JUDG_ERROR = "EJUDG"       '' 判定エラー
Public Const MSG_ENTRY_ERROR = "EETRY"      '' 登録エラー
Public Const MSG_SIGMACHECK_ERROR = "ESIGM" '' シグマチェックエラー
Public Const MSG_R2CHECK_ERROR = "ER2CK"    '' 相関係数チェックエラー
Public Const MSG_CALCULATE_ERROR = "ECALC"  '' 計算エラー
Public Const MSG_FTIRCHECK_ERROR = "EFTIR"  '' FTIR換算値チェックエラー
Public Const MSG_EFFECTTIME_ERROR = "EFTIM" '' FTIR相関式有効時間エラー
'Public Const MSG_GETERROR_DBDATA = "EGET"   '' DBデータ取得エラー
'Public Const MSG_DISPLAY_ERROR = "EDISP"    '' 表示エラー
Public Const MSG_ENTRY = "PPROK"            '' 登録メッセージ
Public Const MSG_KTKBN = "ESMPK"            '' 確定メッセージ
Public Const MSG_INSPECT_ERROR = "EINSP"    '' 検査方法未対応メッセージ

'' 結晶サンプル管理テーブル取得・更新モード
Public Const MODE_GETSMPL_FTIR = 1          '' FTIR(Oi,Cs)
Public Const MODE_GETSMPL_GFA = 2           '' GFA(Oi)
Public Const MODE_GETSMPL_RS = 3            '' 抵抗
Public Const MODE_GETSMPL_BMD = 4           '' BMD
Public Const MODE_GETSMPL_OSF = 5           '' OSF
Public Const MODE_GETSMPL_GD = 6            '' GD
Public Const MODE_GETSMPL_LT = 7            '' ライフタイム
Public Const MODE_GETSMPL_EPD = 8           '' EPD
Public Const MODE_GETSMPL_X = 9             '' X線
Public Const MODE_GETSMPL_CUDECO = 10       '' Cu-deco(C,CJ,CJLT,CJ2)   Add 2010/12/17 SMPK Miyata

'' 結晶検査種類
Public Enum chkKensaType
    CHK_OI         '' Oi
    CHK_CS         '' Cs
    CHK_RS         '' Rs
    CHK_B1         '' BMD1
    CHK_B2         '' BMD2
    CHK_B3         '' BMD3
    CHK_L1         '' OSF1
    CHK_L2         '' OSF2
    CHK_L3         '' OSF3
    CHK_L4         '' OSF4
    CHK_GD         '' GD
    CHK_LT         '' LT
    CHK_EP         '' EPD
    CHK_X          '' X線   2009/08 SUMCO Akizuki
    'Add Start 2010/12/17 SMPK Miyata
    CHK_C          '' C
    CHK_CJ         '' CJ
    CHK_CJLT       '' CJLT
    CHK_CJ2        '' CJ2
    'Add End   2010/12/17 SMPK Miyata
End Enum

'' データ管理テーブル定義（グローバル変数定義）
Public tbl_PrSpSXLData1() As typ_TBCME018   '' 製品仕様ＳＸＬデータ１
Public tbl_PrSpSXLData2() As typ_TBCME019   '' 製品仕様ＳＸＬデータ２
Public tbl_PrSpSXLData3() As typ_TBCME020   '' 製品仕様ＳＸＬデータ３
'*** UPDATE START Y.SIMIZU 2005/10/1 TBCME036ﾃｰﾌﾞﾙﾃﾞｰﾀ格納用
Public tbl_PrSpSXLData4() As typ_TBCME036   '' 製品仕様ＳＸＬデータ４
'*** UPDATE END Y.SIMIZU 2005/10/1 TBCME036ﾃｰﾌﾞﾙﾃﾞｰﾀ格納用
Public tbl_PrSpWFData1() As typ_TBCME021            '' 製品仕様ＷＦデータ１　05/03/01 ooba START ==>
Public tbl_PrSpWFData2() As typ_TBCME022            '' 製品仕様ＷＦデータ２
Public tbl_PrSpWFData6() As typ_TBCME026            '' 製品仕様ＷＦデータ６
Public tbl_PrSpWFData8() As typ_TBCME028            '' 製品仕様ＷＦデータ８  2005/06/15 ffc)tanabe

''Upd Start (TCS)T.Terauchi 2005/10/07
Public tbl_PrSpWFData36() As typ_TBCME036            '' 製品仕様ＷＦデータ
''Upd End   (TCS)T.Terauchi 2005/10/07

Public tbl_CrystalSampleManage_Cw() As typ_XSDCW    '' 新ｻﾝﾌﾟﾙ管理(SXL)
Public tbl_HinbanCW() As tFullHinban                '' 新ｻﾝﾌﾟﾙ管理(SXL)の品番
Public tbl_WFGDRslt() As typ_TBCMJ015               '' ＧＤ実績(WF)         05/03/01 ooba END ====>
Public tbl_WFSPVRslt() As typ_TBCMJ016              '' SPV実績(WF)          2005/06/16 ffc)tanabe
Public tbl_SXLInsideSpecManager() As typ_TBCME036   '' 結晶内側管理
Public tbl_PlupEndRslt() As typ_TBCMH004            '' 引上げ終了実績

Public tbl_GFADevInfo() As typ_TBCMB014             '' ＧＦＡ校正情報
Public tbl_HinbanManage() As typ_TBCME041           '' 品番管理
Public tbl_BlockManage() As typ_TBCME040            '' ブロック管理
Public tbl_CrystalSampleManage() As typ_XSDCS       '' 結晶サンプル管理
Public tbl_CrystalSampleManage2() As typ_XSDCS      '' 結晶サンプル管理     Add 2010/12/17 SMPK Miyata
Public tbl_EPDRslt() As typ_TBCMJ001                '' ＥＰＤ実績
Public tbl_CryRsRslt() As typ_TBCMJ002              '' 結晶抵抗実績
Public tbl_OiRslt() As typ_TBCMJ003                 '' Ｏｉ実績
Public tbl_CsRslt() As typ_TBCMJ004                 '' Ｃｓ実績
Public tbl_BMDRslt() As typ_TBCMJ008                '' ＢＭＤ実績
Public tbl_OSFRslt() As typ_TBCMJ005                '' ＯＳＦ実績
Public tbl_GDRslt() As typ_TBCMJ006                 '' ＧＤ実績
Public tbl_LifeTime() As typ_TBCMJ007               '' ライフタイム実績
Public tbl_XRslt() As typ_TBCMJ021                  '' X線測定実績          2009/08 SUMCO Akizuki
Public tbl_CuDecoRslt() As typ_TBCMJ023             '' Cu-deco実績          Add 2010/12/17 SMPK Miyata


'' 比抵抗計算位置計算情報
Public Type typ_CalcRsPosInf
    dChgWt      As Double   '' 仕込み重量(チャージ量)
    dTopWT      As Double   '' トップ重量
    dArea       As Double   '' 断面積
    dHenseki    As Double   '' 実行偏析値
    dSmpPos     As Double   '' サンプル位置
    dR0Ce       As Double   '' サンプル位置0mmにおける比抵抗中央値（存在する場合に設定）
    dRx         As Double   '' 対象サンプルの比抵抗中央値
End Type

'' Cs70%推定値計算情報
Public Type typ_Cs70PInf
    dChgWt      As Double   '' 仕込み重量(チャージ量)
    dTopWT      As Double   '' トップ重量
    dArea       As Double   '' 断面積
    dSmpPos     As Double   '' サンプル位置
    dCs         As Double   '' Cs濃度値
End Type

'概要      :コンボボックスの状態変更を行う
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :ctrlObj       ,I   ,Control   ,コントロールオブジェクト
'          :bFlag         ,I   ,Boolean   ,コントロール状態指示（True：有効　False：無効）
'          :[bClear]      ,I   ,Boolean   ,コンボボックス内容クリア指示（True：クリア　False：クリアしない）
'説明      :
Public Sub EnableComboBoxCtrl(ctrlObj As Control, bFlag As Boolean, Optional bClear As Boolean = False)

    If bFlag = True Then
        ctrlObj.Enabled = True
        ctrlObj.BackColor = vbWindowBackground
    Else
        ctrlObj.Enabled = False
        ctrlObj.BackColor = vbButtonFace
    End If

    If bClear Then
        ctrlObj.Clear
    End If

End Sub


'概要      :カンマ区切り文字列から任意の場所の文字列を切り取る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :strTarget     ,I   ,String    ,カンマ区切りの文字列
'          :iField        ,I   ,Integer   ,取得したい場所(1～)
'          :戻り値        ,O  ,String    ,取得文字列。取得できなかった場合vbNullStringを返す
'説明      :
Public Function GetStringField(strTarget As String, iField As Integer) As String
    Dim strWork     As String
    Dim strGet      As String
    Dim iPos        As Integer
    Dim Index       As Integer
    
    GetStringField = vbNullString

    If iField <= 0 Then Exit Function

    strWork = strTarget
    Index = 1
    Do
        iPos = InStr(strWork, ",")
        If iPos = 0 Then
            If Index < iField Then Exit Function
            strGet = strWork
            Exit Do
        End If
        strGet = Left(strWork, iPos - 1)
        strWork = Right(strWork, Len(strWork) - iPos)
        If Index = iField Then Exit Do
        Index = Index + 1
    Loop

    GetStringField = strGet

End Function

'概要      :小数点切り捨てを行う
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :param         ,I   ,Double    ,対象数値
'          :iPoint        ,I   ,Integer   ,小数点以下切り捨て数
'          :戻り値        ,O  ,Double    ,小数点以下切捨て結果
'説明      :
Public Function CutDecimalPointParam(ByVal param As Double, iPoint As Integer) As Double
    Dim Index      As Integer
    Dim strParam   As String
    Dim iStrLen    As Integer
    Dim iTen       As Integer
    Dim strWork    As String
    Dim bFlag      As Boolean

    bFlag = False
    strParam = Str(param)
    iStrLen = Len(strParam)
    For Index = 1 To iStrLen
        If Mid(strParam, Index, 1) = "." Then
            bFlag = True
            iTen = Index
            Exit For
        End If
    Next Index
    If bFlag <> True Then CutDecimalPointParam = param: Exit Function
    
    strWork = Mid(strParam, iTen + 1, iPoint)
    CutDecimalPointParam = val(Left(strParam, iTen) + strWork)

End Function

'概要      :小数点切り上げを行う
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :param         ,I   ,Double    ,対象数値
'          :iPoint        ,I   ,Integer   ,小数点以下切り上げ数
'          :戻り値        ,O  ,Double    ,小数点以下切り上げ結果
'説明      :
Public Function UpDecimalPointParam(ByVal param As Double, iPoint As Integer) As Double
    Dim dWork As Double
    
    dWork = param - CutDecimalPointParam(param, iPoint)
    If dWork > 0 Then
        UpDecimalPointParam = CutDecimalPointParam(param, iPoint) + (10 ^ (-iPoint))
    Else
        UpDecimalPointParam = param
    End If

End Function

'概要      :コンボボックスの状態変更を行う
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :ctrlObj       ,I   ,Control   ,コントロールオブジェクト
'          :戻り値        ,O  ,Integer   ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Public Function CheckIsInputText(ctrlObj As Control) As Integer
    Dim bDisable As Boolean
    
    '' 初期化
    bDisable = False
    
    CheckIsInputText = FUNCTION_RETURN_SUCCESS
    
    '' 表示項目である場合
    If ctrlObj.BackColor = COLOR_DISABLE And _
           ctrlObj.Locked = True And ctrlObj.TabStop = False Then
        bDisable = True
    End If
    
    '' 入力チェック
    If ctrlObj.Text <> "" Then  '' 入力されている場合
        If bDisable <> True Then
            CtrlEnabled ctrlObj, CTRL_ENABLE
        End If
        Exit Function
    Else                        '' 入力されていない場合
        If bDisable = True Then
            Exit Function
        End If
        CtrlEnabled ctrlObj, CTRL_WARNING
    End If
    
    CheckIsInputText = FUNCTION_RETURN_FAILURE
End Function


'概要      :指定サンプルNo.が確定されているか調べる
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型             ,説明
'          :tblCrySmp()   ,I  ,typ_XSDCS   ,結晶サンプル管理テーブル配列
'          :iSmpNo        ,I  ,Long        ,サンプルNo.     Integer→Long 6桁対応 2007/05/28 SETsw kubota
'          :戻り値        ,O  ,Boolean        ,True:確定している  False:確定していない
'説明      :
Public Function CheckKTKBN(tblCrySmp() As typ_XSDCS, iSmpNo As Long) As Boolean
    Dim Index As Integer
    
    CheckKTKBN = False
    For Index = 0 To UBound(tblCrySmp) - 1
        If (tblCrySmp(Index).REPSMPLIDCS = iSmpNo) And (tblCrySmp(Index).KTKBNCS = "1") Then
            CheckKTKBN = True
            Exit Function
        End If
    Next Index

End Function
'概要      :判定フラグ値より判定文字列を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名       ,IO ,型        ,説明
'          :bJudg        ,I   ,Boolean  ,判定フラグ値
'          :戻り値        ,O  ,String   ,判定文字列
'説明      :
Public Function GetJudgStr(bJudg As Boolean) As String

    If bJudg = True Then
        GetJudgStr = STR_JUDG_OK
    Else
        GetJudgStr = STR_JUDG_NG
    End If

End Function
'概要      :判定フラグ値より判定文字列を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名       ,IO ,型        ,説明
'          :strJudg      ,I   ,String  ,判定フラグ
'          :戻り値        ,O  ,String   ,判定文字列
'説明      :
Public Function GetResJudgStr(strJudg As String) As String

    Select Case strJudg
    Case "1"
        GetResJudgStr = STR_JUDG_OK
    Case "2"
        GetResJudgStr = STR_JUDG_NG
    Case Else
        GetResJudgStr = ""
    End Select

End Function
'概要      :判定フラグ値より結晶検査実績コードを作成する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :bJudg         ,I   ,Boolean   ,判定フラグ値
'          :戻り値        ,O  ,String    ,結晶検査実績コード
'説明      :
Public Function MakeCryResultCode(bJudg As Boolean) As String


''　実績FLGの「0:未検査,1:判定OK,2:判定NG」変更に伴い、再修正　2003/09/26 SystemBrain ==========================> START
''　検査指示変更　2003/09/10 Motegi ==========================> START
    If bJudg = True Then
        MakeCryResultCode = "1"
    Else
        MakeCryResultCode = "2"
    End If
     
'     MakeCryResultCode = "1"
''　検査指示変更　2003/09/10 Motegi ==========================> END

End Function
'概要      :測定点コードより測定点を求める
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :strMeasNum    ,I   ,String  ,測定点コード（1byte）
'          :[iDefNum]     ,I   ,Integer  ,デフォルト測定点（該当の測定点コードがない場合にこの値を返す）
'          :戻り値        ,O  ,Integer   ,測定点
'説明      :
Public Function GetMeasureNum(strMeasNum As String, Optional iDefNum As Integer = 0) As Integer
    Dim iNum    As Integer

    iNum = iDefNum
    If strMeasNum <> "" Then
        If Asc(strMeasNum) >= Asc("0") And Asc(strMeasNum) <= Asc("9") Then
            '' ０～９の場合
            iNum = val(strMeasNum)
        ElseIf Asc(strMeasNum) >= Asc("A") And Asc(strMeasNum) <= Asc("K") Then
            '' Ａ～Ｋの場合
            iNum = 10 + Asc(strMeasNum) - Asc("A")
        ElseIf strMeasNum = "X" Then
            '' Ｘの場合
            iNum = 20
        End If
    End If
    
    GetMeasureNum = iNum

End Function
'' 号機コンボボックス文字列よりコードを取得する
Public Function GetCmbCode(cmb As ComboBox) As String
Dim s As String
Dim POS As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmhc001j.frm -- Function GetCmbCode"

    s = cmb.Text
    POS = InStr(1, s, ":")
    If POS Then
        s = Left$(s, POS - 1)
    End If
    GetCmbCode = s

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function
'' 品番情報をクリアする
Public Sub ClearFullHinban(tHIN As tFullHinban)
    tHIN.hinban = ""
    tHIN.mnorevno = 0
    tHIN.factory = ""
    tHIN.opecond = ""
End Sub

'' 品番管理より品番情報をセットする
Public Sub SetFullHinban_TBCME041(tHIN As tFullHinban, tblHinban As typ_TBCME041)
    tHIN.hinban = tblHinban.hinban
    tHIN.mnorevno = tblHinban.REVNUM
    tHIN.factory = tblHinban.factory
    tHIN.opecond = tblHinban.opecond
End Sub

'概要      :サンプルＮｏ.より結晶番号を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :strCryNum     ,O   ,String    ,結晶番号
'          :lSmpNo        ,I   ,Long      ,サンプルＮｏ.
'          :lSmpMode      ,I   ,Long      ,サンプル管理テーブル取得モード
'          :戻り値        ,O  ,Integer    ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
'Public Function GetCryNum(strCryNum As String, iSmpNo As Long) As Integer
Public Function GetCryNum(strCryNum As String, lSmpNo As Long, Optional ByVal lSmpMode As Long = 0) As Integer
    Dim iRet        As Integer
    Dim tblGet()    As typ_XSDCS
    
    Dim lCnt        As Long
    Dim lXtalNoCnt  As Long

    GetCryNum = FUNCTION_RETURN_FAILURE

    '' 結晶番号の取得
    'iRet = DBDRV_GetTBCME043(tblGet, "where REPSMPLIDCS=" & CStr(lSmpNo) & " and KTKBNCS='0'")
    iRet = DBDRV_GetTBCME043(tblGet, "where REPSMPLIDCS=" & CStr(lSmpNo) & " and KTKBNCS='0'", "order by TDAYCS desc")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then
        'iRet = DBDRV_GetTBCME043(tblGet, "where REPSMPLIDCS=" & CStr(lSmpNo))
        iRet = DBDRV_GetTBCME043(tblGet, "where REPSMPLIDCS=" & CStr(lSmpNo), "order by TDAYCS desc")
        If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
        If UBound(tblGet) = 0 Then Exit Function
'>>>>> サンプルNo.チェックの追加 2011/06/27 SETsw kubota -----------------------------------
        '確定済みの場合、代表サンプルIDだけでなく、各サンプルのサンプルIDと合致するかをチェック
        lXtalNoCnt = 0
        For lCnt = 1 To UBound(tblGet)
            'サンプルIDが合致したらループから抜ける
            If ChkMeasSmpl(tblGet(lCnt), lSmpNo, lSmpMode) = True Then
                lXtalNoCnt = lCnt       '合致した行を返す
                Exit For
            End If
        Next lCnt
        If lXtalNoCnt = 0 Then
            '合致するデータが無ければエラー
            Exit Function
        End If
    Else
        '未確定のレコードがあれば、既存通り
        lXtalNoCnt = 1
'<<<<< サンプルNo.チェックの追加 2011/06/27 SETsw kubota -----------------------------------
    End If
    
    'strCryNum = tblGet(1).XTALCS
    strCryNum = tblGet(lXtalNoCnt).XTALCS
    
    GetCryNum = FUNCTION_RETURN_SUCCESS

End Function

'概要      :サンプルＮｏ.が各測定のサンプルIDカラムにあるかを判断
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :tXsdcs        ,I  ,typ_XSDCS ,XSDCSデータ
'          :lSmpNo        ,I  ,Long      ,サンプルＮｏ.
'          :lSmpMode      ,I  ,Long      ,サンプル管理テーブル取得モード
'          :戻り値        ,O  ,Boolean   ,サンプルIDが一致：True　一致しない：False
'説明      :2011/06/28追加 SETsw kubota
Public Function ChkMeasSmpl(ByRef tXsdcs As typ_XSDCS _
                          , ByVal lSmpNo As Long _
                          , ByVal lSmpMode As Long _
                          ) As Boolean
    
    Dim bSmpFlg     As Boolean
    bSmpFlg = False
    
    '各サンプルのサンプルIDと一致するかをチェック
    With tXsdcs
        Select Case lSmpMode
        Case MODE_GETSMPL_FTIR      '' FTIR(Oi,Cs)
            If lSmpNo = .CRYSMPLIDOICS _
            Or lSmpNo = .CRYSMPLIDCSCS Then
                bSmpFlg = True
            End If
        Case MODE_GETSMPL_GFA       '' GFA(Oi)
            If lSmpNo = .CRYSMPLIDOICS Then
                bSmpFlg = True
            End If
        Case MODE_GETSMPL_RS        '' 抵抗
            If lSmpNo = .CRYSMPLIDRSCS Then
                bSmpFlg = True
            End If
        Case MODE_GETSMPL_BMD       '' BMD
            If lSmpNo = .CRYSMPLIDB1CS _
            Or lSmpNo = .CRYSMPLIDB2CS _
            Or lSmpNo = .CRYSMPLIDB3CS Then
                bSmpFlg = True
            End If
        Case MODE_GETSMPL_OSF       '' OSF
            If lSmpNo = .CRYSMPLIDL1CS _
            Or lSmpNo = .CRYSMPLIDL2CS _
            Or lSmpNo = .CRYSMPLIDL3CS _
            Or lSmpNo = .CRYSMPLIDL4CS Then
                bSmpFlg = True
            End If
        Case MODE_GETSMPL_GD        '' GD
            If lSmpNo = .CRYSMPLIDGDCS Then
                bSmpFlg = True
            End If
        Case MODE_GETSMPL_LT        '' ライフタイム
            If lSmpNo = .CRYSMPLIDTCS Then
                bSmpFlg = True
            End If
        Case MODE_GETSMPL_EPD       '' EPD
            If lSmpNo = .CRYSMPLIDEPCS Then
                bSmpFlg = True
            End If
        Case MODE_GETSMPL_X         '' X線
            If lSmpNo = .CRYSMPLIDXCS Then
                bSmpFlg = True
            End If
        Case MODE_GETSMPL_CUDECO    '' Cu-deco(C,CJ,CJLT,CJ2)
            If lSmpNo = .CRYSMPLIDCCS _
            Or lSmpNo = .CRYSMPLIDCJCS _
            Or lSmpNo = .CRYSMPLIDCJLTCS _
            Or lSmpNo = .CRYSMPLIDCJ2CS Then
                bSmpFlg = True
            End If
        Case Else                   '' その他
            'その他はそのままOK
            bSmpFlg = True
        End Select
    End With

    ChkMeasSmpl = bSmpFlg

End Function

'概要      :社員コードより社員名を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :strName       ,O   ,String    ,社員名
'          :strID         ,I   ,String    ,社員コード
'          :戻り値        ,O  ,Integer    ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Public Function GetStaffNameStr(strName As String, strID As String) As Integer

    GetStaffNameStr = FUNCTION_RETURN_FAILURE

    '' 社員名の取得
        '2009/09 Akizuki TBCMB001参照から、KODA9参照へ変更
        'strName = GetStaffName(strID)
        strName = GetStaffName_KODA9(strID)
    
    If strName = vbNullString Then
        Exit Function
    End If

    GetStaffNameStr = FUNCTION_RETURN_SUCCESS

End Function
'概要      :結晶番号より品番管理テーブルを取得する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :tblHinban()   ,O   ,typ_TBCME041 ,品番管理テーブル
'          :strCryNum     ,I   ,String           ,結晶番号
'          :戻り値        ,O  ,Integer          ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Public Function GetHinban(tblHinban() As typ_TBCME041, strCryNum As String) As Integer
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
            .factory = vbNullString
            .opecond = vbNullString
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

'概要      :結晶番号より新ｻﾝﾌﾟﾙ管理(SXL)の品番を取得する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO  ,型               ,説明
'          :tblHinban()   ,O   ,tFullHinban      ,12桁品番
'          :strCryNum     ,I   ,String           ,結晶番号
'          :戻り値        ,O   ,Integer          ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
'履歴      :05/03/01 ooba
Public Function GetHinban_WF(tblHinban() As tFullHinban, strCryNum As String) As Integer

    Dim Index       As Integer
    Dim tblIndex    As Integer
    Dim iChk        As Integer
    Dim bChkFlg     As Boolean

    GetHinban_WF = FUNCTION_RETURN_FAILURE

    '' 12桁品番構造体を初期化
    ReDim tblHinban(0)

    '' 新ｻﾝﾌﾟﾙ管理(SXL)ﾃｰﾌﾞﾙが存在しなければ処理を抜ける
    If UBound(tbl_CrystalSampleManage_Cw) = 0 Then Exit Function
    
    For Index = 1 To UBound(tbl_CrystalSampleManage_Cw)
        If Trim(tbl_CrystalSampleManage_Cw(Index).HINBCW) <> "Z" And _
           Trim(tbl_CrystalSampleManage_Cw(Index).HINBCW) <> "G" And _
           Trim(tbl_CrystalSampleManage_Cw(Index).HINBCW) <> "" Then
            
            bChkFlg = False
            '' 既存品番との重複ﾁｪｯｸ
            If UBound(tblHinban) > 0 Then
                For iChk = 1 To UBound(tblHinban)
                    '' 品番が一致していればﾃﾞｰﾀをｾｯﾄしない
                    If tbl_CrystalSampleManage_Cw(Index).HINBCW = tblHinban(iChk).hinban And _
                       tbl_CrystalSampleManage_Cw(Index).REVNUMCW = tblHinban(iChk).mnorevno And _
                       tbl_CrystalSampleManage_Cw(Index).FACTORYCW = tblHinban(iChk).factory And _
                       tbl_CrystalSampleManage_Cw(Index).OPECW = tblHinban(iChk).opecond Then
                       
                        bChkFlg = True
                        Exit For
                    End If
                Next
            End If
            
            If bChkFlg = False Then
                '' 品番ﾃﾞｰﾀ格納領域拡張
                tblIndex = UBound(tblHinban) + 1
                ReDim Preserve tblHinban(tblIndex)
                '' 12桁品番ﾃﾞｰﾀの取得
                tblHinban(tblIndex).hinban = tbl_CrystalSampleManage_Cw(Index).HINBCW
                tblHinban(tblIndex).mnorevno = tbl_CrystalSampleManage_Cw(Index).REVNUMCW
                tblHinban(tblIndex).factory = tbl_CrystalSampleManage_Cw(Index).FACTORYCW
                tblHinban(tblIndex).opecond = tbl_CrystalSampleManage_Cw(Index).OPECW
            End If
        End If
    Next Index
    
    If UBound(tblHinban) = 0 Then
        If Len(strCryNum) <> 12 Then Exit Function
        
        ReDim tblHinban(1)
        tblHinban(1).hinban = "Z"
        tblHinban(1).mnorevno = 0
        tblHinban(1).factory = vbNullString
        tblHinban(1).opecond = vbNullString
    End If

    GetHinban_WF = FUNCTION_RETURN_SUCCESS

End Function

'概要      :品番より製品仕様ＳＸＬデータ１を取得、そして、取得した製品仕様データを追加する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                   ,説明
'          :tblSP()       ,O   ,typ_TBCME018        ,製品仕様ＳＸＬデータ１テーブル配列
'          :tHinInf       ,I   ,tFullHinban         ,品番
'          :ctrlFrm       ,I   ,Form                ,フォームID
'          :戻り値        ,O  ,Integer              ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Public Function GetSPSXLData1(tblSP() As typ_TBCME018, tHinInf As tFullHinban, ctrlFrm As Form) As Integer
    Dim Index       As Long
    Dim iRet        As Integer
    Dim tHinban(0)  As tFullHinban
    Dim tblGet()    As typ_TBCME018
    
    GetSPSXLData1 = FUNCTION_RETURN_FAILURE

    tHinban(0) = tHinInf
    
    ''Cng Start 2011/02/21 Y.Hitomi 処理順序を逆転
    '' すでに取得品番仕様を記憶している場合、取得データを追加しない
    For Index = 0 To UBound(tblSP) - 1
        If (tblSP(Index).hinban = tHinInf.hinban) And (tblSP(Index).mnorevno = tHinInf.mnorevno) And _
           (tblSP(Index).factory = tHinInf.factory) And (tblSP(Index).opecond = tHinInf.opecond) Then
            GetSPSXLData1 = FUNCTION_RETURN_SUCCESS
            Exit Function
        End If
    Next Index
    
    '' 製品仕様ＳＸＬデータ１の取得
    iRet = DBDRV_GetTBCME018(tblGet, ctrlFrm.Name, tHinban)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function
    ''Cng End 2011/02/21 Y.Hitomi
    
    '' 取得した製品仕様ＳＸＬデータの追加
    Add_PrSpSXLData1 tblSP, tblGet(1)
    
    GetSPSXLData1 = FUNCTION_RETURN_SUCCESS

End Function


'概要      :品番より製品仕様ＳＸＬデータ２を取得、そして、取得した製品仕様データを追加する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :tblSP()       ,O   ,typ_TBCME019     ,製品仕様ＳＸＬデータ２テーブル配列
'          :tHinInf       ,I   ,tFullHinban      ,品番
'          :ctrlFrm       ,I   ,Form             ,フォームID
'          :戻り値        ,O  ,Integer          ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Public Function GetSPSXLData2(tblSP() As typ_TBCME019, tHinInf As tFullHinban, ctrlFrm As Form) As Integer
    Dim Index       As Long
    Dim iRet        As Integer
    Dim tHinban(0)  As tFullHinban
    Dim tblGet()    As typ_TBCME019

    GetSPSXLData2 = FUNCTION_RETURN_FAILURE

    tHinban(0) = tHinInf

    ''Cng Start 2011/02/21 Y.Hitomi 処理順序を逆転
    '' すでに取得品番仕様を記憶している場合、取得データを追加しない
    For Index = 0 To UBound(tblSP) - 1
        If (tblSP(Index).hinban = tHinInf.hinban) And (tblSP(Index).mnorevno = tHinInf.mnorevno) And _
           (tblSP(Index).factory = tHinInf.factory) And (tblSP(Index).opecond = tHinInf.opecond) Then
            GetSPSXLData2 = FUNCTION_RETURN_SUCCESS
            Exit Function
        End If
    Next Index
    
        '' 製品仕様ＳＸＬデータ２の取得
    iRet = DBDRV_GetTBCME019(tblGet, ctrlFrm.Name, tHinban)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function
    ''Cng End 2011/02/21 Y.Hitomi
    
    '' 取得した製品仕様ＳＸＬデータの追加
    Add_PrSpSXLData2 tblSP, tblGet(1)
    
    GetSPSXLData2 = FUNCTION_RETURN_SUCCESS

End Function


'概要      :品番より製品仕様ＳＸＬデータ３を取得、そして、取得した製品仕様データを追加する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :tblSP()       ,O   ,typ_TBCME020 ,製品仕様ＳＸＬデータ３テーブル配列
'          :tHinInf       ,I   ,tFullHinban         ,品番
'          :ctrlFrm       ,I   ,Form             ,フォームID
'          :戻り値        ,O  ,Integer           ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Public Function GetSPSXLData3(tblSP() As typ_TBCME020, tHinInf As tFullHinban, ctrlFrm As Form) As Integer
    Dim Index       As Long
    Dim iRet        As Integer
    Dim tHinban(0)  As tFullHinban
    Dim tblGet()    As typ_TBCME020
    
    GetSPSXLData3 = FUNCTION_RETURN_FAILURE

    tHinban(0) = tHinInf

    ''Cng Start 2011/02/21 Y.Hitomi 処理順序を逆転
    '' すでに取得品番仕様を記憶している場合、取得データを追加しない
    For Index = 0 To UBound(tblSP) - 1
        If (tblSP(Index).hinban = tHinInf.hinban) And (tblSP(Index).mnorevno = tHinInf.mnorevno) And _
           (tblSP(Index).factory = tHinInf.factory) And (tblSP(Index).opecond = tHinInf.opecond) Then
            GetSPSXLData3 = FUNCTION_RETURN_SUCCESS
            Exit Function
        End If
    Next Index
    
    '' 製品仕様ＳＸＬデータ３の取得
    iRet = DBDRV_GetTBCME020(tblGet, ctrlFrm.Name, tHinban)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function
    ''Cng End 2011/02/21 Y.Hitomi
    
    '' 取得した製品仕様ＳＸＬデータの追加
    Add_PrSpSXLData3 tblSP, tblGet(1)
    
    GetSPSXLData3 = FUNCTION_RETURN_SUCCESS

End Function

'*** UPDATE ↓ Y.SIMIZU 2005/10/1
'概要      :品番より製品仕様ＳＸＬデータ４を取得、そして、取得した製品仕様データを追加する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :tblSP()       ,O   ,typ_TBCME036 ,製品仕様ＳＸＬデータ４テーブル配列
'          :tHinInf       ,I   ,tFullHinban         ,品番
'          :ctrlFrm       ,I   ,Form             ,フォームID
'          :戻り値        ,O  ,Integer           ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Public Function GetSPSXLData4(tblSP() As typ_TBCME036, tHinInf As tFullHinban, ctrlFrm As Form) As Integer
    Dim Index       As Long
    Dim iRet        As Integer
    Dim tHinban(0)  As tFullHinban
    Dim tblGet()    As typ_TBCME036
    
    GetSPSXLData4 = FUNCTION_RETURN_FAILURE

    tHinban(0) = tHinInf

    ''Cng Start 2011/02/21 Y.Hitomi 処理順序を逆転
    '' すでに取得品番仕様を記憶している場合、取得データを追加しない
    For Index = 0 To UBound(tblSP) - 1
        If (tblSP(Index).hinban = tHinInf.hinban) And (tblSP(Index).mnorevno = tHinInf.mnorevno) And _
           (tblSP(Index).factory = tHinInf.factory) And (tblSP(Index).opecond = tHinInf.opecond) Then
            GetSPSXLData4 = FUNCTION_RETURN_SUCCESS
            Exit Function
        End If
    Next Index
    
    '' 製品仕様ＳＸＬデータ４の取得
    iRet = DBDRV_GetTBCME036(tblGet, ctrlFrm.Name, tHinban)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function
    ''Cng End 2011/02/21 Y.Hitomi

    
    '' 取得した製品仕様ＳＸＬデータの追加
    Add_PrSpSXLData4 tblSP, tblGet(1)
    
    GetSPSXLData4 = FUNCTION_RETURN_SUCCESS

End Function
'*** UPDATE ↑ Y.SIMIZU 2005/10/1

'概要      :品番より製品仕様ＷＦデータ１を取得、そして、取得した製品仕様データを追加する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :tblSP()       ,O   ,typ_TBCME021   ,製品仕様ＷＦデータ１テーブル配列
'          :tHinInf       ,I   ,tFullHinban    ,12桁品番
'          :ctrlFrm       ,I   ,Form           ,フォームID
'          :戻り値        ,O  ,Integer          ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
'履歴      :05/03/01 ooba
Public Function GetSPWFData1(tblSP() As typ_TBCME021, tHinInf As tFullHinban, ctrlFrm As Form) As Integer

    Dim Index       As Integer
    Dim iRet        As Integer
    Dim tHinban(0)  As tFullHinban
    Dim tblGet()    As typ_TBCME021
    
    GetSPWFData1 = FUNCTION_RETURN_FAILURE

    tHinban(0) = tHinInf

    '' 製品仕様ＷＦデータ１の取得
    iRet = DBDRV_GetTBCME021(tblGet, ctrlFrm.Name, tHinban)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function

    '' 取得した製品仕様ＷＦデータの追加
    '' テーブルデータ格納領域拡張
    ReDim Preserve tblSP(UBound(tblSP) + 1)
    '' テーブルデータ数を取得
    Index = UBound(tblSP)

    '' データ追加
    tblSP(Index) = tblGet(1)
    
    GetSPWFData1 = FUNCTION_RETURN_SUCCESS
    
End Function

'概要      :品番より製品仕様ＷＦデータ２を取得、そして、取得した製品仕様データを追加する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :tblSP()       ,O   ,typ_TBCME022   ,製品仕様ＷＦデータ２テーブル配列
'          :tHinInf       ,I   ,tFullHinban    ,12桁品番
'          :ctrlFrm       ,I   ,Form           ,フォームID
'          :戻り値        ,O  ,Integer          ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
'履歴      :05/03/01 ooba
Public Function GetSPWFData2(tblSP() As typ_TBCME022, tHinInf As tFullHinban, ctrlFrm As Form) As Integer

    Dim Index       As Integer
    Dim iRet        As Integer
    Dim tHinban(0)  As tFullHinban
    Dim tblGet()    As typ_TBCME022
    
    GetSPWFData2 = FUNCTION_RETURN_FAILURE

    tHinban(0) = tHinInf

    '' 製品仕様ＷＦデータ２の取得
    iRet = DBDRV_GetTBCME022(tblGet, ctrlFrm.Name, tHinban)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function

    '' 取得した製品仕様ＷＦデータの追加
    '' テーブルデータ格納領域拡張
    ReDim Preserve tblSP(UBound(tblSP) + 1)
    '' テーブルデータ数を取得
    Index = UBound(tblSP)

    '' データ追加
    tblSP(Index) = tblGet(1)
    
    GetSPWFData2 = FUNCTION_RETURN_SUCCESS
    
End Function

'概要      :品番より製品仕様ＷＦデータ６を取得、そして、取得した製品仕様データを追加する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :tblSP()       ,O   ,typ_TBCME026   ,製品仕様ＷＦデータ６テーブル配列
'          :tHinInf       ,I   ,tFullHinban    ,12桁品番
'          :ctrlFrm       ,I   ,Form           ,フォームID
'          :戻り値        ,O  ,Integer          ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
'履歴      :05/03/01 ooba
Public Function GetSPWFData6(tblSP() As typ_TBCME026, tHinInf As tFullHinban, ctrlFrm As Form) As Integer

    Dim Index       As Integer
    Dim iRet        As Integer
    Dim tHinban(0)  As tFullHinban
    Dim tblGet()    As typ_TBCME026
    
    GetSPWFData6 = FUNCTION_RETURN_FAILURE

    tHinban(0) = tHinInf

    '' 製品仕様ＷＦデータ６の取得
    iRet = DBDRV_GetTBCME026(tblGet, ctrlFrm.Name, tHinban)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function

    '' 取得した製品仕様ＷＦデータの追加
    '' テーブルデータ格納領域拡張
    ReDim Preserve tblSP(UBound(tblSP) + 1)
    '' テーブルデータ数を取得
    Index = UBound(tblSP)

    '' データ追加
    tblSP(Index) = tblGet(1)
    
    GetSPWFData6 = FUNCTION_RETURN_SUCCESS
    
End Function

'概要      :品番より製品仕様ＷＦデータ８を取得、そして、取得した製品仕様データを追加する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :tblSP()       ,O   ,typ_TBCME028   ,製品仕様ＷＦデータ８テーブル配列
'          :tHinInf       ,I   ,tFullHinban    ,12桁品番
'          :ctrlFrm       ,I   ,Form           ,フォームID
'          :戻り値        ,O  ,Integer          ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
'履歴      :新規作成 05/06/16 ffc)tanabe
Public Function GetSPWFData8(tblSP() As typ_TBCME028, tHinInf As tFullHinban, ctrlFrm As Form) As Integer

    Dim Index       As Integer
    Dim iRet        As Integer
    Dim tHinban(0)  As tFullHinban
    Dim tblGet()    As typ_TBCME028
    
    GetSPWFData8 = FUNCTION_RETURN_FAILURE

    tHinban(0) = tHinInf

    '' 製品仕様ＷＦデータ８の取得
    iRet = DBDRV_GetTBCME028(tblGet, ctrlFrm.Name, tHinban)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function

    '' 取得した製品仕様ＷＦデータの追加
    '' テーブルデータ格納領域拡張
    ReDim Preserve tblSP(UBound(tblSP) + 1)
    '' テーブルデータ数を取得
    Index = UBound(tblSP)

    '' データ追加
    tblSP(Index) = tblGet(1)
    
    GetSPWFData8 = FUNCTION_RETURN_SUCCESS
    
End Function

'概要      :結晶サンプル検査指示"3"サンプルチェック
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                       ,説明
'          :tblCrySmp     ,I   ,typ_XSDCS               ,新サンプル管理（ブロック）テーブル
'          :kensa_typ     ,I   ,chkKensaType            ,検査指示項目
'          :[iChkOpt]     ,I   ,Integer                 ,サンプルチェック動作モード
'          :戻り値        ,O  ,Boolean                  ,サンプルチェック動作モードにより戻り値の意味が異なる
'説明      :
'           [iChkOpt = 1,2以外] ：  同結晶、同位置の他のサンプルに検査指示"3"が指示されているか確認
'                   戻り値：TRUE    指示あり
'                   　　　：FALSE   指示なし
'           [iChkOpt = 1] ：        同結晶、同位置の他のサンプルに検査指示"3"が指示されているか確認
'                                   続いて、第一引数の結晶サンプルのサンプル区分の種類を返す
'                   戻り値：TRUE    サンプル区分"B"である
'                   　　　：FALSE   サンプル区分"T"である
'           [iChkOpt = 2] ：        同結晶、同位置の他のサンプルに検査指示"3"が指示されているか確認
'                                   続いて、サンプル区分"B","T"のサンプルの確定区分を確認する
'                   戻り値：TRUE    サンプル区分"B","T"のサンプルのいずれかが確定されている
'                   　　　：FALSE   サンプル区分"B","T"のサンプルのいずれも確定されていない
'
Public Function ChkCommonKensa(tblCrySmp As typ_XSDCS, Kensa_Typ As chkKensaType, Optional iChkOpt = 1) As Boolean
    Dim iRet        As Integer
    Dim tblGet()    As typ_XSDCS
    Dim bFind       As Boolean
    Dim sqlWhere    As String
    Dim KeyItem     As String
    ''　結晶サンプル検査指示"3"チェック　2003/09/08 Motegi ===========================> START
'    Const keyComm = "3"
    '----------------------
    Const keyComm = "1"
    
    
    ChkCommonKensa = False

    '' 検査指示"3"のチェック
'    bFind = False
'    Select Case Kensa_Typ
'        Case CHK_OI         '' Oi
'            If tblCrySmp.CRYINDOICS = keyComm Then bFind = True: KeyItem = "CRYINDOICS"
'        Case CHK_CS         '' Cs
'            If tblCrySmp.CRYINDCSCS = keyComm Then bFind = True: KeyItem = "CRYINDCSCS"
'        Case CHK_RS         '' Rs
'            If tblCrySmp.CRYINDRSCS = keyComm Then bFind = True: KeyItem = "CRYINDRSCS"
'        Case CHK_B1         '' BMD1
'            If tblCrySmp.CRYINDB1CS = keyComm Then bFind = True: KeyItem = "CRYINDB1CS"
'        Case CHK_B2         '' BMD2
'            If tblCrySmp.CRYINDB2CS = keyComm Then bFind = True: KeyItem = "CRYINDB2CS"
'        Case CHK_B3         '' BMD3
'            If tblCrySmp.CRYINDB3CS = keyComm Then bFind = True: KeyItem = "CRYINDB3CS"
'        Case CHK_L1         '' OSF1
'            If tblCrySmp.CRYINDL1CS = keyComm Then bFind = True: KeyItem = "CRYINDL1CS"
'        Case CHK_L2         '' OSF2
'            If tblCrySmp.CRYINDL2CS = keyComm Then bFind = True: KeyItem = "CRYINDL2CS"
'        Case CHK_L3         '' OSF3
'            If tblCrySmp.CRYINDL3CS = keyComm Then bFind = True: KeyItem = "CRYINDL3CS"
'        Case CHK_L4         '' OSF4
'            If tblCrySmp.CRYINDL4CS = keyComm Then bFind = True: KeyItem = "CRYINDL4CS"
'        Case CHK_GD         '' GD
'            If tblCrySmp.CRYINDGDCS = keyComm Then bFind = True: KeyItem = "CRYINDGDCS"
'        Case CHK_LT         '' LT
'            If tblCrySmp.CRYINDTCS = keyComm Then bFind = True: KeyItem = "CRYINDTCS"
'        Case CHK_EP         '' EPD
'            If tblCrySmp.CRYINDEPCS = keyComm Then bFind = True: KeyItem = "CRYINDEPCS"
'    End Select
'    If bFind <> True Then Exit Function

    Select Case Kensa_Typ
        Case CHK_OI         '' Oi
            KeyItem = "CRYINDOICS"
        Case CHK_CS         '' Cs
            KeyItem = "CRYINDCSCS"
        Case CHK_RS         '' Rs
            KeyItem = "CRYINDRSCS"
        Case CHK_B1         '' BMD1
            KeyItem = "CRYINDB1CS"
        Case CHK_B2         '' BMD2
            KeyItem = "CRYINDB2CS"
        Case CHK_B3         '' BMD3
            KeyItem = "CRYINDB3CS"
        Case CHK_L1         '' OSF1
            KeyItem = "CRYINDL1CS"
        Case CHK_L2         '' OSF2
            KeyItem = "CRYINDL2CS"
        Case CHK_L3         '' OSF3
            KeyItem = "CRYINDL3CS"
        Case CHK_L4         '' OSF4
            KeyItem = "CRYINDL4CS"
        Case CHK_GD         '' GD
            KeyItem = "CRYINDGDCS"
        Case CHK_LT         '' LT
            KeyItem = "CRYINDTCS"
        Case CHK_EP         '' EPD
            KeyItem = "CRYINDEPCS"
    End Select

    ''　結晶サンプル検査指示"3"チェック　2003/09/08 Motegi ===========================> END

    '' SQL条件作成
    sqlWhere = "where XTALCS='" & tblCrySmp.XTALCS & "' "
    sqlWhere = sqlWhere + "and INPOSCS=" & tblCrySmp.INPOSCS & " "
    sqlWhere = sqlWhere + "and " & KeyItem & "='" & keyComm & "'"

    '' 結晶サンプル管理テーブルの取得
    iRet = DBDRV_GetTBCME043(tblGet, sqlWhere, "order by SMPKBNCS")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    
    If UBound(tblGet) <> 2 Then '' 2件以外の場合
        Exit Function
    End If

    Select Case iChkOpt '' 動作モードの選択
        Case 1
            ''　サンプル区分"T"の場合、区分"T"サンプルを表示させたいので戻り値Falseで処理終了
            If Trim(tblCrySmp.SMPKBNCS) = "T" Then Exit Function
        Case 2 '' 確定区分チェック
            If tblGet(1).KTKBNCS = "0" And tblGet(2).KTKBNCS = "0" Then Exit Function
    End Select

    ChkCommonKensa = True
End Function
'概要      :結晶番号に一致する結晶サンプル管理を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                       ,説明
'          :tblSmpl()     ,O   ,typ_XSDCS               ,新サンプル管理（ブロック）テーブル
'          :strCryNum     ,I   ,String                  ,結晶番号
'          :iMode         ,I   ,Integer                 ,結晶サンプル管理更新モード
'          :戻り値        ,O  ,Integer                  ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Public Function GetSmplManage(tblSmpl() As typ_XSDCS, strCryNum As String, iMode As Integer) As Integer
    
    Dim Index       As Long
    Dim iRet        As Integer
    Dim tblGet()    As typ_XSDCS
    Dim tblTgt()    As typ_XSDCS
    Dim bFind       As Boolean
    Dim HinbanMng() As typ_TBCME041
    Dim UpHin       As tFullHinban
    Dim downHin     As tFullHinban
    Dim sKensa      As String * 1
    Dim tHinInf     As tFullHinban
    
    GetSmplManage = FUNCTION_RETURN_FAILURE
    ReDim tblGet(0)
    ReDim tblTgt(0)

    '' 結晶サンプル管理テーブルを初期化
    RemoveAll_CrystalSampleManage tblSmpl

    '' 結晶サンプル管理テーブルの取得
    iRet = DBDRV_GetTBCME043(tblGet, "where XTALCS='" & strCryNum & "' ", "order by INPOSCS, SMPKBNCS")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function
    
    '' 対象、結晶サンプル管理テーブルの抜き出し
    For Index = 1 To UBound(tblGet)
        bFind = False
        '' 取得モードチェック
'新ｻﾝﾌﾟﾙ管理対応　2003/09/08 Motegi ========================================> 変更開始
'        Select Case iMode
'            Case MODE_GETSMPL_FTIR      '' FTIR(Oi,Cs)
'                If InStr(CODE_KENSA, tblGet(Index).CRYINDOICS) <> 0 Or _
'                   InStr(CODE_KENSA, tblGet(Index).CRYINDCSCS) <> 0 Then
'                    '' 共通検査項目チェック
'                    If ChkCommonKensa(tblGet(Index), CHK_OI) And _
'                       ChkCommonKensa(tblGet(Index), CHK_CS) Then
'                        bFind = False
'                    Else
'                        If ChkCommonKensa(tblGet(Index), CHK_OI, 2) Or _
'                           ChkCommonKensa(tblGet(Index), CHK_CS, 2) Then tblGet(Index).KTKBNCS = "1"
''既定の検査方法が設定されているため、仕様有無のチェックは不要
''                        If (GetReferHinban(tHinInf, tblGet(Index), CHK_OI) = FUNCTION_RETURN_SUCCESS) Or _
''                           (GetReferHinban(tHinInf, tblGet(Index), CHK_CS) = FUNCTION_RETURN_SUCCESS) Then
'                            bFind = True
''                        End If
'                    End If
'                End If
'            Case MODE_GETSMPL_GFA       '' GFA(Oi)
'                If InStr(CODE_KENSA, tblGet(Index).CRYINDOICS) <> 0 Then
'                    '' 共通検査項目チェック
'                    If ChkCommonKensa(tblGet(Index), CHK_OI) Then
'                        bFind = False
'                    Else
'                        If ChkCommonKensa(tblGet(Index), CHK_OI, 2) Then tblGet(Index).KTKBNCS = "1"
''                        If GetReferHinban(tHinInf, tblGet(Index), CHK_OI) = FUNCTION_RETURN_SUCCESS Then
'                            bFind = True
''                        End If
'                    End If
'                End If
'            Case MODE_GETSMPL_RS        '' 抵抗
'                If InStr(CODE_KENSA, tblGet(Index).CRYINDRSCS) <> 0 Then
'                    If ChkCommonKensa(tblGet(Index), CHK_RS) Then
'                        bFind = False
'                    Else
'                        If ChkCommonKensa(tblGet(Index), CHK_RS, 2) Then tblGet(Index).KTKBNCS = "1"
''                        If GetReferHinban(tHinInf, tblGet(Index), CHK_RS) = FUNCTION_RETURN_SUCCESS Then
'                            bFind = True
''                        End If
'                    End If
'            Case MODE_GETSMPL_BMD       '' BMD
'                If InStr(CODE_KENSA, tblGet(Index).CRYINDB1CS) <> 0 Or _
'                   InStr(CODE_KENSA, tblGet(Index).CRYINDB2CS) <> 0 Or _
'                   InStr(CODE_KENSA, tblGet(Index).CRYINDB3CS) <> 0 Then
'                    bFind = True
'                End If
'            Case MODE_GETSMPL_OSF       '' OSF
'                If InStr(CODE_KENSA, tblGet(Index).CRYINDL1CS) <> 0 Or _
'                   InStr(CODE_KENSA, tblGet(Index).CRYINDL2CS) <> 0 Or _
'                   InStr(CODE_KENSA, tblGet(Index).CRYINDL3CS) <> 0 Or _
'                   InStr(CODE_KENSA, tblGet(Index).CRYINDL4CS) <> 0 Then
'                    bFind = True
'                End If
'            Case MODE_GETSMPL_GD        '' GD
'                If InStr(CODE_KENSA, tblGet(Index).CRYINDGDCS) <> 0 Then
'                    If ChkCommonKensa(tblGet(Index), CHK_GD) Then
'                        bFind = False
'                    Else
'                        If ChkCommonKensa(tblGet(Index), CHK_GD, 2) Then tblGet(Index).KTKBNCS = "1"
'                        If GetReferHinban(tHinInf, tblGet(Index), CHK_GD) = FUNCTION_RETURN_SUCCESS Then
'                            bFind = True
'                        End If
'                    End If
'                End If
'            Case MODE_GETSMPL_LT        '' ライフタイム
'                If InStr(CODE_KENSA, tblGet(Index).CRYINDTCS) <> 0 Then
'                    If ChkCommonKensa(tblGet(Index), CHK_LT) Then
'                        bFind = False
'                    Else
'                        If ChkCommonKensa(tblGet(Index), CHK_LT, 2) Then tblGet(Index).KTKBNCS = "1"
'                        If GetReferHinban(tHinInf, tblGet(Index), CHK_LT) = FUNCTION_RETURN_SUCCESS Then
'                            bFind = True
'                        End If
'                    End If
'                End If
'            Case MODE_GETSMPL_EPD       '' EPD
'                If InStr(CODE_KENSA, tblGet(Index).CRYINDEPCS) <> 0 Then
'                    If ChkCommonKensa(tblGet(Index), CHK_EP) Then
'                        bFind = False
'                    Else
'                        If ChkCommonKensa(tblGet(Index), CHK_EP, 2) Then tblGet(Index).KTKBNCS = "1"
'                        If GetReferHinban(tHinInf, tblGet(Index), CHK_EP) = FUNCTION_RETURN_SUCCESS Then
'                            bFind = True
'                        End If
'                    End If
'                End If
'            Case Else                   '' その他
'                Exit Function
'        End Select
'-------------------------------------
        Select Case iMode
            Case MODE_GETSMPL_FTIR      '' FTIR(Oi,Cs)
                If InStr(CODE_KENSA, tblGet(Index).CRYINDOICS) = 1 Or _
                   InStr(CODE_KENSA, tblGet(Index).CRYINDCSCS) = 1 Then
                    bFind = True
                End If
            Case MODE_GETSMPL_GFA       '' GFA(Oi)
                If InStr(CODE_KENSA, tblGet(Index).CRYINDOICS) = 1 Then
                    bFind = True
                End If
            Case MODE_GETSMPL_RS        '' 抵抗
                If InStr(CODE_KENSA, tblGet(Index).CRYINDRSCS) = 1 Then
                    bFind = True
                End If
            Case MODE_GETSMPL_BMD       '' BMD
                If InStr(CODE_KENSA, tblGet(Index).CRYINDB1CS) = 1 Or _
                   InStr(CODE_KENSA, tblGet(Index).CRYINDB2CS) = 1 Or _
                   InStr(CODE_KENSA, tblGet(Index).CRYINDB3CS) = 1 Then
                    bFind = True
                End If
            Case MODE_GETSMPL_OSF       '' OSF
                If InStr(CODE_KENSA, tblGet(Index).CRYINDL1CS) = 1 Or _
                   InStr(CODE_KENSA, tblGet(Index).CRYINDL2CS) = 1 Or _
                   InStr(CODE_KENSA, tblGet(Index).CRYINDL3CS) = 1 Or _
                   InStr(CODE_KENSA, tblGet(Index).CRYINDL4CS) = 1 Then
                    bFind = True
                End If
            Case MODE_GETSMPL_GD        '' GD
                If InStr(CODE_KENSA, tblGet(Index).CRYINDGDCS) = 1 Then
                    bFind = True
                End If
            Case MODE_GETSMPL_LT        '' ライフタイム
                If InStr(CODE_KENSA, tblGet(Index).CRYINDTCS) = 1 Then
                    bFind = True
                End If
            Case MODE_GETSMPL_EPD       '' EPD
                If InStr(CODE_KENSA, tblGet(Index).CRYINDEPCS) = 1 Then
                    bFind = True
                End If
            
            '2009/08 SUMCO Akizuki　X線測定実績入力　作成に伴い追加
            '[1:検査指示あり]
            Case MODE_GETSMPL_X       '' X線
                If InStr(CODE_KENSA, tblGet(Index).CRYINDXCS) = 1 Then
                    bFind = True
                End If

            'Add Start 2010/12/17 SMPK Miyata
            Case MODE_GETSMPL_CUDECO    '' Cu-deco
                If InStr(CODE_KENSA, tblGet(Index).CRYINDCCS) = 1 Or _
                   InStr(CODE_KENSA, tblGet(Index).CRYINDCJCS) = 1 Or _
                   InStr(CODE_KENSA, tblGet(Index).CRYINDCJLTCS) = 1 Or _
                   InStr(CODE_KENSA, tblGet(Index).CRYINDCJ2CS) = 1 Then
                    bFind = True
                End If
            'Add End   2010/12/17 SMPK Miyata

            Case Else                   '' その他
                Exit Function
        End Select
'新ｻﾝﾌﾟﾙ管理対応　2003/09/08 Motegi ========================================> 変更終了
        
        '' 結晶検査指示がある場合、結晶サンプル管理テーブルに出力する
        If bFind = True Then
            If Add_CrystalSampleManage(tblSmpl, tblGet(Index)) <> FUNCTION_RETURN_SUCCESS Then
                Exit Function
            End If
        End If
    Next Index
    
    GetSmplManage = FUNCTION_RETURN_SUCCESS

End Function

'概要      :結晶サンプル管理を更新する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                       ,説明
'          :tblCrySmpMan  ,I   ,typ_XSDCS               ,新サンプル管理（ブロック）テーブル更新パラメータ
'          :strCryNum     ,I   ,String                  ,結晶番号
'          :iIngotPos     ,I   ,Integer                 ,結晶内位置
'          :strSmpKbn     ,I   ,String                  ,サンプル区分
'          :iSmpNo        ,I   ,Long                    ,サンプルNo.    Integer→Long 6桁対応 SETsw kubota
'          :iMode         ,I   ,Integer                 ,結晶サンプル管理更新モード
'          :[iOption]     ,I   ,Integer                 ,結晶サンプル管理更新モードオプション
'          :戻り値        ,O  ,Integer                  ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Public Function UpdateTbl_CrySmpManage(tblCrySmpMan As typ_XSDCS, strCryNum As String, iIngotpos As Integer, strSmpKbn As String, iSmpNo As Long, iMode As Integer, Optional iOption As Integer = 0) As Integer
    Dim iRet        As Integer
    Dim sqlWhere    As String
''　検査指示変更　2003/09/10 Motegi ==========================> START
'    Dim tblGet()    As typ_XSDCS
'    Dim strKeySmpKbn As String
'    Dim bUpdateFlag As Boolean
    Dim sqlUpdate   As String
''　検査指示変更　2003/09/10 Motegi ==========================> END
    
    UpdateTbl_CrySmpManage = FUNCTION_RETURN_FAILURE
''　検査指示変更　2003/09/10 Motegi ==========================> START
'    bUpdateFlag = False
    
'    If Trim(strSmpKbn) = "B" Then
'        strKeySmpKbn = "T"
'    Else
'        strKeySmpKbn = "B"
'    End If
    
'    ReDim tblGet(0)
'    '' 結晶サンプル管理テーブルの取得
'    sqlWhere = " where XTALCS='" & strCryNum & "' " & "and INPOSCS=" & iIngotPos & _
'               " and SMPKBNCS='" & strKeySmpKbn & "' "
'    iRet = DBDRV_GetTBCME043(tblGet, sqlWhere, "order by INPOSCS, SMPKBNCS")
'    If (iRet = FUNCTION_RETURN_SUCCESS) And (UBound(tblGet) > 0) Then
        '' サンプルに"3（共通）"の検査指示が設定されている場合、
        '' 検索された結晶サンプル管理テーブルの検査実績も更新する
'        Select Case iMode
'        Case MODE_GETSMPL_FTIR      '' FTIR(Oi,Cs)
'            If (tblCrySmpMan.CRYINDOICS = tblGet(1).CRYINDOICS) And (tblCrySmpMan.CRYINDOICS = "3") And (iOption = 1) Then
'                tblGet(1).CRYRESOICS = tblCrySmpMan.CRYRESOICS
'                bUpdateFlag = True
'            End If
'            If (tblCrySmpMan.CRYINDCSCS = tblGet(1).CRYINDCSCS) And (tblCrySmpMan.CRYINDCSCS = "3") And (iOption = 2) Then
'                tblGet(1).CRYRESCSCS = tblCrySmpMan.CRYRESCSCS
'                bUpdateFlag = True
'            End If
'        Case MODE_GETSMPL_GFA       '' GFA(Oi)
'            If (tblCrySmpMan.CRYINDOICS = tblGet(1).CRYINDOICS) And (tblCrySmpMan.CRYINDOICS = "3") Then
'                tblGet(1).CRYRESOICS = tblCrySmpMan.CRYRESOICS
'                bUpdateFlag = True
'            End If
'        Case MODE_GETSMPL_RS        '' 抵抗
'            If (tblCrySmpMan.CRYINDRSCS = tblGet(1).CRYINDRSCS) And (tblCrySmpMan.CRYINDRSCS = "3") Then
'                tblGet(1).CRYRESRS1CS = tblCrySmpMan.CRYRESRS1CS
'                bUpdateFlag = True
'            End If
'        Case MODE_GETSMPL_BMD       '' BMD
'            If (tblCrySmpMan.CRYINDB1CS = tblGet(1).CRYINDB1CS) And (tblCrySmpMan.CRYINDB1CS = "3") And (iOption = 1) Then
'                tblGet(1).CRYRESB1CS = tblCrySmpMan.CRYRESB1CS
'                bUpdateFlag = True
'            End If
'            If (tblCrySmpMan.CRYINDB2CS = tblGet(1).CRYINDB2CS) And (tblCrySmpMan.CRYINDB2CS = "3") And (iOption = 2) Then
'                tblGet(1).CRYRESB2CS = tblCrySmpMan.CRYRESB2CS
'                bUpdateFlag = True
'            End If
'            If (tblCrySmpMan.CRYINDB3CS = tblGet(1).CRYINDB3CS) And (tblCrySmpMan.CRYINDB3CS = "3") And (iOption = 3) Then
'                tblGet(1).CRYRESB3CS = tblCrySmpMan.CRYRESB3CS
'                bUpdateFlag = True
'            End If
'        Case MODE_GETSMPL_OSF       '' OSF
'            If (tblCrySmpMan.CRYINDL1CS = tblGet(1).CRYINDL1CS) And (tblCrySmpMan.CRYINDL1CS = "3") And (iOption = 1) Then
'                tblGet(1).CRYRESL1CS = tblCrySmpMan.CRYRESL1CS
'                bUpdateFlag = True
'            End If
'            If (tblCrySmpMan.CRYINDL2CS = tblGet(1).CRYINDL2CS) And (tblCrySmpMan.CRYINDL2CS = "3") And (iOption = 2) Then
'                tblGet(1).CRYRESL2CS = tblCrySmpMan.CRYRESL2CS
'                bUpdateFlag = True
'            End If
'            If (tblCrySmpMan.CRYINDL3CS = tblGet(1).CRYINDL3CS) And (tblCrySmpMan.CRYINDL3CS = "3") And (iOption = 3) Then
'                tblGet(1).CRYRESL3CS = tblCrySmpMan.CRYRESL3CS
'                bUpdateFlag = True
'            End If
'            If (tblCrySmpMan.CRYINDL4CS = tblGet(1).CRYINDL4CS) And (tblCrySmpMan.CRYINDL4CS = "3") And (iOption = 4) Then
'                tblGet(1).CRYRESL4CS = tblCrySmpMan.CRYRESL4CS
'                bUpdateFlag = True
'            End If
'        Case MODE_GETSMPL_GD        '' GD
'            If (tblCrySmpMan.CRYINDGDCS = tblGet(1).CRYINDGDCS) And (tblCrySmpMan.CRYINDGDCS = "3") Then
'                tblGet(1).CRYINDGDCS = tblCrySmpMan.CRYINDGDCS
'                bUpdateFlag = True
'            End If
'        Case MODE_GETSMPL_LT        '' ライフタイム
'            If (tblCrySmpMan.CRYINDTCS = tblGet(1).CRYINDTCS) And (tblCrySmpMan.CRYINDTCS = "3") Then
'                tblGet(1).CRYINDTCS = tblCrySmpMan.CRYINDTCS
'                bUpdateFlag = True
'            End If
'        Case MODE_GETSMPL_EPD       '' EPD
'            If (tblCrySmpMan.CRYINDEPCS = tblGet(1).CRYINDEPCS) And (tblCrySmpMan.CRYINDEPCS = "3") Then
'                tblGet(1).CRYINDEPCS = tblCrySmpMan.CRYINDEPCS
'                bUpdateFlag = True
'            End If
'        End Select
'------------------------------
    With tblCrySmpMan
'2009/08　SUMCO Akizuki サンプル管理更新処理に、変更情報の反映がなかったため、追加
'>>>>>
        sqlUpdate = "update XSDCS set "
        sqlUpdate = sqlUpdate & "KSTAFFCS = '" & .KSTAFFCS & "' ,"          '更新社員ID
        sqlUpdate = sqlUpdate & "KDAYCS = SYSDATE ,"                        '更新日付
        sqlUpdate = sqlUpdate & "SNDKDWHCS = '0' ,"                         '送信フラグ(DWH)
'<<<<<

        Select Case iMode
            Case MODE_GETSMPL_FTIR      '' FTIR(Oi,Cs)
                If iOption = 1 Then
                    sqlUpdate = sqlUpdate & "CRYRESOICS = '" & .CRYRESOICS & "' "           ' 結晶検査実績（Oi)
                    sqlWhere = "CRYSMPLIDOICS = " & iSmpNo
                ElseIf iOption = 2 Then
                    sqlUpdate = sqlUpdate & "CRYRESCSCS = '" & .CRYRESCSCS & "' "           ' 結晶検査実績（Cs)
                    sqlWhere = "CRYSMPLIDCSCS = " & iSmpNo
                End If
            
            Case MODE_GETSMPL_GFA       '' GFA(Oi)
                sqlUpdate = sqlUpdate & "CRYRESOICS = '" & .CRYRESOICS & "' "               ' 結晶検査実績（Oi)
                sqlWhere = "CRYSMPLIDOICS = " & iSmpNo
            
            Case MODE_GETSMPL_RS        '' 抵抗
                sqlUpdate = sqlUpdate & "CRYRESRS1CS = '" & .CRYRESRS1CS & "' "             ' 結晶検査実績（Rs)
                sqlWhere = "CRYSMPLIDRSCS = " & iSmpNo
            
            Case MODE_GETSMPL_BMD       '' BMD
                If iOption = 1 Then
                    sqlUpdate = sqlUpdate & "CRYRESB1CS = '" & .CRYRESB1CS & "' "           ' 結晶検査実績（BMD1)
                    sqlWhere = "CRYSMPLIDB1CS = " & iSmpNo
                ElseIf iOption = 2 Then
                    sqlUpdate = sqlUpdate & "CRYRESB2CS = '" & .CRYRESB2CS & "' "           ' 結晶検査実績（BMD2)
                    sqlWhere = "CRYSMPLIDB2CS = " & iSmpNo
                ElseIf iOption = 3 Then
                    sqlUpdate = sqlUpdate & "CRYRESB3CS = '" & .CRYRESB3CS & "' "           ' 結晶検査実績（BMD3)
                    sqlWhere = "CRYSMPLIDB3CS = " & iSmpNo
                End If
            
            Case MODE_GETSMPL_OSF       '' OSF
                If iOption = 1 Then
                    sqlUpdate = sqlUpdate & "CRYRESL1CS = '" & .CRYRESL1CS & "' "           ' 結晶検査実績（OSF1)
                    sqlWhere = "CRYSMPLIDL1CS = " & iSmpNo
                ElseIf iOption = 2 Then
                    sqlUpdate = sqlUpdate & "CRYRESL2CS = '" & .CRYRESL2CS & "' "           ' 結晶検査実績（OSF2)
                    sqlWhere = "CRYSMPLIDL2CS = " & iSmpNo
                ElseIf iOption = 3 Then
                    sqlUpdate = sqlUpdate & "CRYRESL3CS = '" & .CRYRESL3CS & "' "           ' 結晶検査実績（OSF3)
                    sqlWhere = "CRYSMPLIDL3CS = " & iSmpNo
                ElseIf iOption = 4 Then
                    sqlUpdate = sqlUpdate & "CRYRESL4CS = '" & .CRYRESL4CS & "' "           ' 結晶検査実績（OSF4)
                    sqlWhere = "CRYSMPLIDL4CS = " & iSmpNo
                End If
            
            Case MODE_GETSMPL_GD        '' GD
                sqlUpdate = sqlUpdate & "CRYRESGDCS = '" & .CRYRESGDCS & "' "               ' 結晶検査実績（GD)
                sqlWhere = "CRYSMPLIDGDCS = " & iSmpNo
            
            Case MODE_GETSMPL_LT        '' ライフタイム
                sqlUpdate = sqlUpdate & "CRYRESTCS = '" & .CRYRESTCS & "', "                 ' 結晶検査実績（LT)
                                        '' ライフタイム(10Ω換算)
                sqlUpdate = sqlUpdate & "CRYREST10CS = '" & .CRYREST10CS & "' "                 ' 結晶検査実績（LT)
                
                sqlWhere = "CRYSMPLIDTCS = " & iSmpNo
            
            Case MODE_GETSMPL_EPD       '' EPD
                sqlUpdate = sqlUpdate & "CRYRESEPCS = '" & .CRYRESEPCS & "' "               ' 結晶検査実績（EPD)
                sqlWhere = "CRYSMPLIDEPCS = " & iSmpNo
            
            '2009/08 Akizuki
            Case MODE_GETSMPL_X         '' X線
                sqlUpdate = sqlUpdate & "CRYRESXCS = '" & .CRYRESXCS & "' "                 ' 結晶検査実績（X線)
                sqlWhere = "CRYSMPLIDXCS = " & iSmpNo
        
            'Add Start 2011/01/07 SMPK Miyata
            Case MODE_GETSMPL_CUDECO    '' Cu-deco
                If iOption = 1 Then
                    sqlUpdate = sqlUpdate & "CRYRESCCS = '" & .CRYRESCCS & "' "             ' 結晶検査実績（C)
                    sqlWhere = "CRYSMPLIDCCS = " & iSmpNo
                
                ElseIf iOption = 2 Then
                    sqlUpdate = sqlUpdate & "CRYRESCJCS = '" & .CRYRESCJCS & "' "           ' 結晶検査実績（CJ)
                    sqlWhere = "CRYSMPLIDCJCS = " & iSmpNo

                ElseIf iOption = 3 Then
                    sqlUpdate = sqlUpdate & "CRYRESCJLTCS = '" & .CRYRESCJLTCS & "' "       ' 結晶検査実績（CJLT)
                    sqlWhere = "CRYSMPLIDCJLTCS = " & iSmpNo

                ElseIf iOption = 4 Then
                    sqlUpdate = sqlUpdate & "CRYRESCJ2CS = '" & .CRYRESCJ2CS & "' "         ' 結晶検査実績（CJ2)
                    sqlWhere = "CRYSMPLIDCJ2CS = " & iSmpNo
                End If
            'Add End   2011/01/07 SMPK Miyata
        
        End Select
        
    End With
                
''　検査指示変更　2003/09/10 Motegi ==========================> END

''　検査指示変更　2003/09/10 Motegi ==========================> 削除START
        '' 結晶サンプル管理テーブルの更新
'        If bUpdateFlag = True Then
'            '' 更新条件の作成
'            sqlWhere = " where XTALCS='" & strCryNum & "' " & " and INPOSCS=" & iIngotPos & _
'                       " and SMPKBNCS='" & strKeySmpKbn & "' " & " and REPSMPLIDCS=" & tblGet(1).REPSMPLIDCS
'            '' 結晶サンプル管理テーブルの更新
'            iRet = DBDRV_UpdateTBCME043(tblCrySmpMan, sqlWhere)
'            If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
'        End If
'    End If
''　検査指示変更　2003/09/10 Motegi ==========================> 削除END

    '' 更新条件の作成
''　検査指示変更　2003/09/10 Motegi ==========================> START
    sqlUpdate = sqlUpdate & " where XTALCS='" & strCryNum & "' " & " and " & sqlWhere
    
    '' 結晶サンプル管理テーブルの更新
    iRet = DBDRV_UpdateXSDCS(sqlUpdate)
''　検査指示変更　2003/09/10 Motegi ==========================> END
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function

    UpdateTbl_CrySmpManage = FUNCTION_RETURN_SUCCESS

End Function
'概要      :仕様があるかを調べる
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                    ,説明
'          :tHinInf       ,O  ,tFullHinban           ,調査品番
'          :kensa_Typ     ,I  ,chkKensaType          ,検査画面が対象とする検査
'          :戻り値        ,O  ,Boolean               ,True：仕様あり　False：仕様なし
'説明      :
Public Function ChkSpExistence(tHinInf As tFullHinban, Kensa_Typ As chkKensaType) As Boolean
    Dim chkSyo      As String * 1
    Dim chkPoint    As String * 1
    Dim iPoint      As Integer
    Dim idx1        As Integer
    Dim idx2        As Integer
    Dim idx3        As Integer
    Dim bFind1      As Boolean
    Dim bFind2      As Boolean
    Dim bFind3      As Boolean

    ChkSpExistence = False
    bFind1 = False
    bFind2 = False
    bFind3 = False

    '' 品番種類の特定
    If (Trim(tHinInf.hinban) = "") Or (Trim(tHinInf.hinban) = "Z") Then '' 空品番、Ｚ品番の場合
        '' 仕様の検査指示を調べる
        If (Kensa_Typ = CHK_RS) Or (Kensa_Typ = CHK_CS) Then
            ChkSpExistence = True
            Exit Function
        Else
            Exit Function
        End If
    ElseIf (Trim(tHinInf.hinban) = "G") Then  '' Ｇ品番の場合
        '' 仕様の検査指示を調べる
        If (Kensa_Typ = CHK_RS) Or (Kensa_Typ = CHK_CS) Or (Kensa_Typ = CHK_OI) Or (Kensa_Typ = CHK_LT) Then
            ChkSpExistence = True
            Exit Function
        Else
            Exit Function
        End If
    Else        '' その他、品番の場合
        '' 対象品番の製品仕様の探索
        For idx1 = 0 To UBound(tbl_PrSpSXLData1) - 1
            If (tbl_PrSpSXLData1(idx1).hinban = tHinInf.hinban) And (tbl_PrSpSXLData1(idx1).mnorevno = tHinInf.mnorevno) And _
               (tbl_PrSpSXLData1(idx1).factory = tHinInf.factory) And (tbl_PrSpSXLData1(idx1).opecond = tHinInf.opecond) Then
                bFind1 = True
                Exit For
            End If
        Next idx1
        For idx2 = 0 To UBound(tbl_PrSpSXLData2) - 1
            If (tbl_PrSpSXLData2(idx2).hinban = tHinInf.hinban) And (tbl_PrSpSXLData2(idx2).mnorevno = tHinInf.mnorevno) And _
               (tbl_PrSpSXLData2(idx2).factory = tHinInf.factory) And (tbl_PrSpSXLData2(idx2).opecond = tHinInf.opecond) Then
                bFind2 = True
                Exit For
            End If
        Next idx2
        For idx3 = 0 To UBound(tbl_PrSpSXLData3) - 1
            If (tbl_PrSpSXLData3(idx3).hinban = tHinInf.hinban) And (tbl_PrSpSXLData3(idx3).mnorevno = tHinInf.mnorevno) And _
               (tbl_PrSpSXLData3(idx3).factory = tHinInf.factory) And (tbl_PrSpSXLData3(idx3).opecond = tHinInf.opecond) Then
                bFind3 = True
                Exit For
            End If
        Next idx3
        
        '' 仕様の検査指示を調べる
        Select Case Kensa_Typ
            Case CHK_OI         '' Oi
                If bFind2 = False Then Exit Function
                With tbl_PrSpSXLData2(idx2)
                    chkSyo = .HSXONHWS
                    chkPoint = .HSXONSPT
                End With
            Case CHK_CS         '' Cs
                If bFind2 = False Then Exit Function
                With tbl_PrSpSXLData2(idx2)
                    chkSyo = .HSXCNHWS
                    chkPoint = .HSXCNSPT
                End With
            Case CHK_RS         '' Rs
                If bFind1 = False Then Exit Function
                With tbl_PrSpSXLData1(idx1)
                    chkSyo = .HSXRHWYS
                    chkPoint = .HSXRSPOT
                End With
            Case CHK_B1         '' BMD1
                If bFind3 = False Then Exit Function
                With tbl_PrSpSXLData3(idx3)
                    chkSyo = .HSXBM1HS
                    chkPoint = .HSXBM1ST
                End With
            Case CHK_B2         '' BMD2
                If bFind3 = False Then Exit Function
                With tbl_PrSpSXLData3(idx3)
                    chkSyo = .HSXBM2HS
                    chkPoint = .HSXBM2ST
                End With
            Case CHK_B3         '' BMD3
                If bFind3 = False Then Exit Function
                With tbl_PrSpSXLData3(idx3)
                    chkSyo = .HSXBM3HS
                    chkPoint = .HSXBM3ST
                End With
            Case CHK_L1         '' OSF1
                If bFind3 = False Then Exit Function
                With tbl_PrSpSXLData3(idx3)
                    chkSyo = .HSXOF1HS
                    chkPoint = .HSXOF1ST
                End With
            Case CHK_L2         '' OSF2
                If bFind3 = False Then Exit Function
                With tbl_PrSpSXLData3(idx3)
                    chkSyo = .HSXOF2HS
                    chkPoint = .HSXOF2ST
                End With
            Case CHK_L3         '' OSF3
                If bFind3 = False Then Exit Function
                With tbl_PrSpSXLData3(idx3)
                    chkSyo = .HSXOF3HS
                    chkPoint = .HSXOF3ST
                End With
            Case CHK_L4         '' OSF4
                If bFind3 = False Then Exit Function
                With tbl_PrSpSXLData3(idx3)
                    chkSyo = .HSXOF4HS
                    chkPoint = .HSXOF4ST
                End With
            Case CHK_GD         '' GD
                If bFind3 = False Then Exit Function
                With tbl_PrSpSXLData3(idx3)
                    If (.HSXDENHS = "H" Or .HSXDENHS = "S") Or _
                       (.HSXDVDHS = "H" Or .HSXDVDHS = "S") Or _
                       (.HSXLDLHS = "H" Or .HSXLDLHS = "S") Then
                        chkSyo = "S"
                        chkPoint = .HSXGDSPT
                    Else
                        Exit Function
                    End If
                End With
            Case CHK_LT         '' LT
                If bFind2 = False Then Exit Function
                With tbl_PrSpSXLData2(idx2)
                    chkSyo = .HSXLTHWS
                    chkPoint = .HSXLTSPT
                End With
            Case CHK_EP         '' EPD
                '' EPD は通常品番であれば仕様あり
                ChkSpExistence = True
                Exit Function
            Case Else           '' その他
                Exit Function
        End Select
    
        iPoint = GetMeasureNum(chkPoint)
        '' 保証方法＿処、測定点を調べる
        If (chkSyo = "H" Or chkSyo = "S") And (iPoint > 0) Then
            ChkSpExistence = True
        End If
    End If

End Function
'概要      :与えられた品番の測定点数を返す （抵抗とOIのみ対応）
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                    ,説明
'          :KensaTyp    ,I  ,chkKensaType          ,検査種類
'          :tHinInf       ,I  ,tfullhinban           ,上品番
'          :戻り値        ,O  ,integer       ,測定点数を返す
'説明      :
Private Function GetMark(Kensa_Typ As chkKensaType, tHinInf As tFullHinban) As Integer
    Dim chkSyo      As String * 1
    Dim chkPoint    As String * 1
    Dim iPoint      As Integer
    Dim idx        As Integer
    Dim bFind      As Boolean
    
    bFind = False
    
    GetMark = -1
    
    '' 品番種類の特定
    If (Trim(tHinInf.hinban) = "") Or (Trim(tHinInf.hinban) = "Z") Then '' 空品番、Ｚ品番の場合
        '' 仕様の検査指示を調べる
        If (Kensa_Typ = CHK_RS) Then
            GetMark = 3
            Exit Function
        ElseIf (Kensa_Typ = CHK_OI) Then
            GetMark = 1
            Exit Function
        Else
            GetMark = -1
            Exit Function
        End If
    ElseIf (Trim(tHinInf.hinban) = "G") Then  '' Ｇ品番の場合
        '' 仕様の検査指示を調べる
        If (Kensa_Typ = CHK_RS) Then
            GetMark = 3
            Exit Function
        ElseIf (Kensa_Typ = CHK_OI) Then
            GetMark = 3
            Exit Function
        Else
            GetMark = -1
            Exit Function
        End If
    Else        '' その他、品番の場合
        
        '' 仕様の検査指示を調べる
        Select Case Kensa_Typ
        Case CHK_OI         '' Oi
            For idx = 0 To UBound(tbl_PrSpSXLData2) - 1
                If (tbl_PrSpSXLData2(idx).hinban = tHinInf.hinban) And (tbl_PrSpSXLData2(idx).mnorevno = tHinInf.mnorevno) And _
                   (tbl_PrSpSXLData2(idx).factory = tHinInf.factory) And (tbl_PrSpSXLData2(idx).opecond = tHinInf.opecond) Then
                    bFind = True
                    Exit For
                End If
            Next idx
            
    
            If bFind = False Then Exit Function
            With tbl_PrSpSXLData2(idx)
                chkSyo = .HSXONHWS
                chkPoint = .HSXONSPT
            End With
        Case CHK_RS         '' Rs
            For idx = 0 To UBound(tbl_PrSpSXLData1) - 1
                If (tbl_PrSpSXLData1(idx).hinban = tHinInf.hinban) And (tbl_PrSpSXLData1(idx).mnorevno = tHinInf.mnorevno) And _
                   (tbl_PrSpSXLData1(idx).factory = tHinInf.factory) And (tbl_PrSpSXLData1(idx).opecond = tHinInf.opecond) Then
                    bFind = True
                    Exit For
                End If
            Next idx
        
            If bFind = False Then Exit Function
            With tbl_PrSpSXLData1(idx)
                chkSyo = .HSXRHWYS
                chkPoint = .HSXRSPOT
            End With
        End Select
    
        iPoint = GetMeasureNum(chkPoint)
        '' 保証方法＿処、測定点を調べる
'2002/02/14 S.Sano        If (chkSyo = "H") Then
            GetMark = iPoint
'2002/02/14 S.Sano        Else
'2002/02/14 S.Sano            GetMark = -1
'2002/02/14 S.Sano        End If
    End If
    
End Function
'概要      :測定点数が厳しい方の品番を返す
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                    ,説明
'          :KensaTyp    ,I  ,chkKensaType          ,検査種類
'          :UpHin       ,I  ,tFullHinban           ,上品番
'          :DwHin       ,I  ,tFullHinban           ,下品番
'          :SmpKbn      ,I  ,string                ,サンプル区分
'          :戻り値        ,O  ,tFullHinban       ,上下品番のどちらかを返す
'説明      :
Private Function GetManyMark(KensaTyp As chkKensaType, UpHin As tFullHinban, DwHin As tFullHinban, SMPKBN As String) As tFullHinban
    
    Dim UpMark As Integer, DwMark As Integer

    UpMark = GetMark(KensaTyp, UpHin)
    DwMark = GetMark(KensaTyp, DwHin)
    If UpMark = DwMark Then                ' 同じだったらサンプル区分の方向
        If Trim$(SMPKBN) = "T" Then
            GetManyMark = DwHin
        Else
            GetManyMark = UpHin
        End If
    ElseIf UpMark > DwMark Then
        GetManyMark = UpHin
    Else
        GetManyMark = DwHin
    End If
    
End Function
'概要      :参照品番を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                   ,説明
'          :tHinInf       ,O  ,tFullHinban          ,参照すべき品番
'          :tblCrySmp     ,I  ,typ_XSDCS            ,結晶サンプル管理テーブル
'          :kensa_Typ     ,I  ,chkKensaType         ,検査画面が対象とする検査
'          :戻り値        ,O  ,FUNCTION_RETURN      ,処理成功：参照品番がある　処理失敗：参照すべき品番はない
'説明      :
Public Function GetReferHinban(tHinInf As tFullHinban, tblCrySmp As typ_XSDCS, Kensa_Typ As chkKensaType) As FUNCTION_RETURN
    Dim UpHin As tFullHinban
    Dim downHin As tFullHinban
    Dim isBtm As Boolean
    Dim chkShiji As String * 1
    Dim LtHin   As tFullHinban
    Dim sLtspi  As String
''*** UPDATE ↓ Y.SIMIZU 2005/10/1 GDﾗｲﾝ数格納用
    Dim GDhin   As tFullHinban
''*** UPDATE ↑ Y.SIMIZU 2005/10/1 GDﾗｲﾝ数格納用
    
    GetReferHinban = FUNCTION_RETURN_FAILURE
    
    '' 初期化
    With tHinInf
        .hinban = vbNullString
        .mnorevno = 0
        .factory = vbNullString
        .opecond = vbNullString
    End With
    
    '' 結晶サンプル管理検査指示の取得
    Select Case Kensa_Typ
        Case CHK_OI         '' Oi
            chkShiji = tblCrySmp.CRYINDOICS
        Case CHK_CS         '' Cs
            chkShiji = tblCrySmp.CRYINDCSCS
        Case CHK_RS         '' Rs
            chkShiji = tblCrySmp.CRYINDRSCS
        Case CHK_B1         '' BMD1
            chkShiji = tblCrySmp.CRYINDB1CS
        Case CHK_B2         '' BMD2
            chkShiji = tblCrySmp.CRYINDB2CS
        Case CHK_B3         '' BMD3
            chkShiji = tblCrySmp.CRYINDB3CS
        Case CHK_L1         '' OSF1
            chkShiji = tblCrySmp.CRYINDL1CS
        Case CHK_L2         '' OSF2
            chkShiji = tblCrySmp.CRYINDL2CS
        Case CHK_L3         '' OSF3
            chkShiji = tblCrySmp.CRYINDL3CS
        Case CHK_L4         '' OSF4
            chkShiji = tblCrySmp.CRYINDL4CS
        Case CHK_GD         '' GD
            chkShiji = tblCrySmp.CRYINDGDCS
        Case CHK_LT         '' LT
            chkShiji = tblCrySmp.CRYINDTCS
        Case CHK_EP         '' EPD
            chkShiji = tblCrySmp.CRYINDEPCS
        Case CHK_X          '' X線              '2009/08 SUMCO Akizuki
            chkShiji = tblCrySmp.CRYINDXCS      '2009/08 SUMCO Akizuki
        'Add Start 2010/12/17 SMPK Miyata
        Case CHK_C         '' C
            chkShiji = tblCrySmp.CRYINDCCS
        Case CHK_CJ        '' CJ
            chkShiji = tblCrySmp.CRYINDCJCS
        Case CHK_CJLT      '' CJLT
            chkShiji = tblCrySmp.CRYINDCJLTCS
        Case CHK_CJ2       '' CJ2
            chkShiji = tblCrySmp.CRYINDCJ2CS
        'Add End   2010/12/17 SMPK Miyata
        Case Else           '' その他
            Exit Function
    End Select
    
    'ライフタイム実績入力の時は下端にサンプル位置を含むブロックの中で
    '最も厳しいLT測定位置をもつ品番を取得
    If Kensa_Typ = CHK_LT Then
        With tblCrySmp
            DBDRV_getLtHinbanInBlock .XTALCS, .INPOSCS, LtHin, sLtspi
            If LtHin.hinban <> "        " Then
                tHinInf = LtHin
                .HINBCS = LtHin.hinban
                .REVNUMCS = LtHin.mnorevno
                .FACTORYCS = LtHin.factory
                .OPECS = LtHin.opecond
                GetReferHinban = FUNCTION_RETURN_SUCCESS
            End If
        End With
        Exit Function
    End If
    
''*** UPDATE ↓ Y.SIMIZU 2005/10/1 最も厳しいGDﾗｲﾝ数品番を持つ品番を取得
    If Kensa_Typ = CHK_GD Then
        With tblCrySmp
            DBDRV_getGDHinbanInBlock tblCrySmp, GDhin
            If GDhin.hinban <> "        " Then
                tHinInf = GDhin
                .HINBCS = GDhin.hinban
                .REVNUMCS = GDhin.mnorevno
                .FACTORYCS = GDhin.factory
                .OPECS = GDhin.opecond
                GetReferHinban = FUNCTION_RETURN_SUCCESS
            End If
        End With
        Exit Function
    End If
''*** UPDATE ↑ Y.SIMIZU 2005/10/1 最も厳しいGDﾗｲﾝ数品番を持つ品番を取得

''検査方向指示変更　2003/09/10 Motegi ========================> START
    '' 上品番、下品番を求める
'    GetUpDownHinban UpHin, downHin, tblCrySmp, tbl_HinbanManage

    '' 結晶検査指示を調べる
'    If (chkShiji = "1") Or (chkShiji = "4" And tblCrySmp.SMPKBNCS = "T") Then       '' 検査方向↓の場合
'        If ChkSpExistence(downHin, Kensa_Typ) Then  '' 下品番に仕様がある場合
'            tHinInf = downHin
'        Else
'            Exit Function
'        End If
'    ElseIf (chkShiji = "2") Or (chkShiji = "4" And tblCrySmp.SMPKBNCS = "B") Then   '' 検査方向↑の場合
'        If ChkSpExistence(UpHin, Kensa_Typ) Then  '' 上品番に仕様がある場合
'            tHinInf = UpHin
'        Else
'            Exit Function
'        End If
'    ElseIf (chkShiji = "3") Then    '' 共通検査の場合
'
'        ' 抵抗、OI　に関しては測定点数が厳しい方を使用する
'        If (Kensa_Typ = CHK_RS) Or (Kensa_Typ = CHK_OI) Then
'            tHinInf = GetManyMark(Kensa_Typ, UpHin, downHin, tblCrySmp.SMPKBNCS)
'        Else
'            If (tblCrySmp.SMPKBNCS = "T") Then      '' サンプル区分"T"の場合
'                If ChkSpExistence(downHin, Kensa_Typ) Then  '' 下品番に仕様がある場合
'                    tHinInf = downHin
'                ElseIf ChkSpExistence(UpHin, Kensa_Typ) Then  '' 上品番に仕様がある場合
'                    tHinInf = UpHin
'                Else
'                    Exit Function
'                End If
'            ElseIf (tblCrySmp.SMPKBNCS = "B") Then  '' サンプル区分"B"の場合
'                If ChkSpExistence(UpHin, Kensa_Typ) Then  '' 上品番に仕様がある場合
'                    tHinInf = UpHin
'                ElseIf ChkSpExistence(downHin, Kensa_Typ) Then  '' 下品番に仕様がある場合
'                    tHinInf = downHin
'                Else
'                    Exit Function
'                End If
'            Else
'                Exit Function
'            End If
'        End If
'    Else    '' その他の検査指示
'        Exit Function
'    End If
'--------------------------------

    With tHinInf
        .hinban = tblCrySmp.HINBCS
        .mnorevno = tblCrySmp.REVNUMCS
        .factory = tblCrySmp.FACTORYCS
        .opecond = tblCrySmp.OPECS
    End With

''検査方向指示変更　2003/09/10 Motegi ========================> END

    GetReferHinban = FUNCTION_RETURN_SUCCESS

End Function

'結晶サンプル管理TBLから検査指示の値を得る
Private Function GetSijiFlg(iSmpNo%, strFldName$) As String
Dim sql$
Dim rs As OraDynaset

    sql = "select " & strFldName & " from XSDCS where REPSMPLIDCS=" & iSmpNo
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount > 0 Then
        GetSijiFlg = rs(strFldName)
    Else
        GetSijiFlg = vbNullString
    End If
    rs.Close
    Set rs = Nothing
End Function
Public Function Add_CryRsRslt(tblTarget() As typ_TBCMJ002, tblDat As typ_TBCMJ002, Optional Index As Long = -1) As Integer
    Dim tblIndex As Long
    
    Add_CryRsRslt = FUNCTION_RETURN_FAILURE

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

    Add_CryRsRslt = FUNCTION_RETURN_SUCCESS
End Function
'概要      :入力パラメータの合計値を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :dParam()      ,I   ,Double    ,パラメータ値配列
'          :戻り値        ,O  ,Double    ,合計値
'説明      :
Public Function GetSum(dParam() As Double) As Double
    Dim Index   As Integer
    Dim dWork   As Double

    On Error GoTo Err

    dWork = 0
    For Index = 0 To UBound(dParam)
        dWork = dWork + dParam(Index)
    Next Index

    GetSum = dWork
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    GetSum = 0
End Function


'概要      :入力パラメータの平均値を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :dParam()      ,I   ,Double    ,パラメータ値配列
'          :戻り値        ,O  ,Double    ,平均値
'説明      :
Public Function GetAve(dParam() As Double) As Double
    Dim dWork   As Double

    On Error GoTo Err

    GetAve = GetSum(dParam) / (UBound(dParam) + 1)

    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    GetAve = 0
End Function


'概要      :入力パラメータの最大値を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :dParam()      ,I   ,Double    ,パラメータ値配列
'          :戻り値        ,O  ,Double    ,最大値
'説明      :
Public Function GetMax(dParam() As Double) As Double
    Dim Index   As Integer
    Dim dMax    As Double

    On Error GoTo Err

    If UBound(dParam) = 0 Then GetMax = dParam(0): Exit Function
    
    dMax = dParam(0)
    For Index = 1 To UBound(dParam)
        If dMax < dParam(Index) Then dMax = dParam(Index)
    Next Index

    GetMax = dMax

    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    GetMax = 0
End Function

'概要      :入力パラメータの最小値を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :dParam()      ,I   ,Double    ,パラメータ値配列
'          :戻り値        ,O  ,Double    ,最小値
'説明      :
Public Function GetMin(dParam() As Double) As Double
    Dim Index   As Integer
    Dim dMin    As Double

    On Error GoTo Err

    If UBound(dParam) = 0 Then GetMin = dParam(0): Exit Function
    
    dMin = dParam(0)
    For Index = 1 To UBound(dParam)
        If dMin > dParam(Index) Then dMin = dParam(Index)
    Next Index

    GetMin = dMin

    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    GetMin = 0
End Function



'概要      :引数を母集団の標本であると見なして、母集団に対する標準偏差を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :dParam()      ,I   ,Double   ,パラメータ値配列
'          :戻り値        ,O  ,Double    ,標準偏差値
'説明      :
Public Function GetSTDEV(dParam() As Double) As Double
    Dim Index   As Integer
    Dim dNum    As Double
    Dim dCalc1  As Double
    Dim dCalc2  As Double

    On Error GoTo Err

    dNum = UBound(dParam) + 1

    dCalc1 = 0
    For Index = 0 To dNum - 1
        dCalc1 = dCalc1 + dParam(Index) ^ 2
    Next Index

    dCalc2 = GetSum(dParam) ^ 2

    GetSTDEV = ((dNum * dCalc1 - dCalc2) / (dNum * (dNum - 1))) ^ 0.5

    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    GetSTDEV = 0
End Function


'概要      :最小２乗法により、直線の傾きを計算する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :dY()          ,I  ,Double    ,既にわかっている y の値の系列の変数（y = mx + b）
'          :dX()          ,I  ,Double    ,既にわかっている x の値の系列の変数（y = mx + b）
'          :戻り値        ,O  ,Double    ,傾き
'説明      :
Public Function CalculateSlope(dY() As Double, dX() As Double) As Double
    Dim Index   As Integer
    Dim dNum    As Double
    Dim dCalc1  As Double
    Dim dCalc2  As Double
    Dim dParam  As Double

    On Error GoTo Err

    dNum = UBound(dY) + 1
    
    '' ∑X(i)Y(i) 計算
    dCalc1 = 0
    For Index = 0 To dNum - 1
        dCalc1 = dCalc1 + dX(Index) * dY(Index)
    Next Index
    '' ∑X(i)^2 計算
    dCalc2 = 0
    For Index = 0 To dNum - 1
        dCalc2 = dCalc2 + dX(Index) ^ 2
    Next Index

    dParam = ((GetSum(dX) ^ 2) - dNum * dCalc2)
    If dParam = 0 Then CalculateSlope = 0: Exit Function

    CalculateSlope = (GetSum(dX) * GetSum(dY) - dNum * dCalc1) / dParam

    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    CalculateSlope = 0
End Function


'概要      :最小２乗法により、直線のY切片を計算する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :dY()          ,I  ,Double    ,既にわかっている y の値の系列の変数（y = mx + b）
'          :dX()          ,I  ,Double    ,既にわかっている x の値の系列の変数（y = mx + b）
'          :戻り値        ,O  ,Double    ,Y切片
'説明      :
Public Function CalculateYFragment(dY() As Double, dX() As Double) As Double
    Dim Index   As Integer
    Dim dNum    As Double

    On Error GoTo Err

    dNum = UBound(dY) + 1

    CalculateYFragment = (GetSum(dY) - GetSum(dX) * CalculateSlope(dY, dX)) / dNum

    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    CalculateYFragment = 0
End Function

'概要      :相関係数を計算する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :dY()          ,I  ,Double    ,既にわかっている y の値の系列の変数（y = mx + b）
'          :dX()          ,I  ,Double    ,既にわかっている x の値の系列の変数（y = mx + b）
'          :戻り値        ,O  ,Double    ,Y切片
'説明      :
Public Function CalculateR2(dY() As Double, dX() As Double) As Double
    Dim Index   As Integer
    Dim dNum    As Double
    Dim dCalc1  As Double
    Dim dCalc2  As Double
    Dim dCalc3  As Double
    Dim dParam  As Double

    On Error GoTo Err

    dNum = UBound(dY) + 1

    dCalc1 = 0
    For Index = 0 To dNum - 1
        dCalc1 = dCalc1 + ((dX(Index) - GetAve(dX)) * (dY(Index) - GetAve(dY)))
    Next Index

    dCalc2 = 0
    For Index = 0 To dNum - 1
        dCalc2 = dCalc2 + (dX(Index) - GetAve(dX)) ^ 2
    Next Index

    dCalc3 = 0
    For Index = 0 To dNum - 1
        dCalc3 = dCalc3 + (dY(Index) - GetAve(dY)) ^ 2
    Next Index

    dParam = (dCalc2 / (dNum - 1)) ^ 0.5 * (dCalc3 / (dNum - 1)) ^ 0.5
    If dParam = 0 Then CalculateR2 = 0: Exit Function
    
    CalculateR2 = ((dCalc1 / (dNum - 1)) / dParam) ^ 2

    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    CalculateR2 = 0
End Function

Public Sub RemoveAll_PlupEndRslt(tblTarget() As typ_TBCMH004)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_PrSpSXLData1(tblTarget() As typ_TBCME018)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_PrSpSXLData2(tblTarget() As typ_TBCME019)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_PrSpSXLData3(tblTarget() As typ_TBCME020)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub

'*** UPDATE ↓ Y.SIMIZU 2005/10/1 TBCME036構造体のﾃﾞｰﾀを初期化
Public Sub RemoveAll_PrSpSXLData4(tblTarget() As typ_TBCME036)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub
'*** UPDATE ↑ Y.SIMIZU 2005/10/1 TBCME036構造体のﾃﾞｰﾀを初期化

Public Sub RemoveAll_SXLInsideSpecManager(tblTarget() As typ_TBCME036)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_GFADevInfo(tblTarget() As typ_TBCMB014)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_HinbanManage(tblTarget() As typ_TBCME041)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_BlockManage(tblTarget() As typ_TBCME040)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_CrystalSampleManage(tblTarget() As typ_XSDCS)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_EPDRslt(tblTarget() As typ_TBCMJ001)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_CryRsRslt(tblTarget() As typ_TBCMJ002)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_OiRslt(tblTarget() As typ_TBCMJ003)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_CsRslt(tblTarget() As typ_TBCMJ004)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_BMDrslt(tblTarget() As typ_TBCMJ008)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_OSFRslt(tblTarget() As typ_TBCMJ005)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_GDRslt(tblTarget() As typ_TBCMJ006)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_LifeTime(tblTarget() As typ_TBCMJ007)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub

'2009/08 SUMCO Akizuki X線測定実績作成に伴い追加
Public Sub RemoveAll_XRslt(tblTarget() As typ_TBCMJ021)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub

'追加 2005/06/17 ffc)tanabe
Public Sub RemoveAll_CrystalSampleManage_Cw(tblTarget() As typ_XSDCW)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub

'Add Start 2010/12/17 SMPK Miyata
Public Sub RemoveAll_CuDecoRslt(tblTarget() As typ_TBCMJ023)
    '' テーブルデータ全削除
    ReDim tblTarget(0)
End Sub
'Add End   2010/12/17 SMPK Miyata

Public Sub InitAllTable()

    '' すべてのテーブルのデータを初期化する
    RemoveAll_PrSpSXLData1 tbl_PrSpSXLData1
    RemoveAll_PrSpSXLData2 tbl_PrSpSXLData2
    RemoveAll_PrSpSXLData3 tbl_PrSpSXLData3
    RemoveAll_GFADevInfo tbl_GFADevInfo
    RemoveAll_HinbanManage tbl_HinbanManage
    RemoveAll_BlockManage tbl_BlockManage
    RemoveAll_CrystalSampleManage tbl_CrystalSampleManage
    RemoveAll_EPDRslt tbl_EPDRslt
    RemoveAll_CryRsRslt tbl_CryRsRslt
    RemoveAll_OiRslt tbl_OiRslt
    RemoveAll_CsRslt tbl_CsRslt
    RemoveAll_BMDrslt tbl_BMDRslt
    RemoveAll_OSFRslt tbl_OSFRslt
    RemoveAll_GDRslt tbl_GDRslt
    RemoveAll_LifeTime tbl_LifeTime
    RemoveAll_SXLInsideSpecManager tbl_SXLInsideSpecManager
    RemoveAll_PlupEndRslt tbl_PlupEndRslt
    RemoveAll_CrystalSampleManage_Cw tbl_CrystalSampleManage_Cw '追加 2005/06/17 ffc)tanabe
    RemoveAll_CuDecoRslt tbl_CuDecoRslt                         'Add 2010/12/17 SMPK Miyata
End Sub
'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCME041」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME041 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村
Public Function DBDRV_GetTBCME041(records() As typ_TBCME041, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, INGOTPOS, HINBAN, REVNUM, FACTORY, OPECOND, LENGTH, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME041"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME041 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .INGOTPOS = rs("INGOTPOS")       ' 結晶内開始位置
            .hinban = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
            .factory = rs("FACTORY")         ' 工場
            .opecond = rs("OPECOND")         ' 操業条件
            .Length = rs("LENGTH")           ' 長さ
            .REGDATE = rs("REGDATE")         ' 登録日付
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME041 = FUNCTION_RETURN_SUCCESS
End Function
Public Function Add_PrSpSXLData2(tblTarget() As typ_TBCME019, tblDat As typ_TBCME019, Optional Index As Long = -1) As Integer
    Dim tblIndex As Long
    
    Add_PrSpSXLData2 = FUNCTION_RETURN_FAILURE

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

    Add_PrSpSXLData2 = FUNCTION_RETURN_SUCCESS
End Function
Public Function Add_PrSpSXLData1(tblTarget() As typ_TBCME018, tblDat As typ_TBCME018, Optional Index As Long = -1) As Integer
    Dim tblIndex As Long
    
    Add_PrSpSXLData1 = FUNCTION_RETURN_FAILURE

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

    Add_PrSpSXLData1 = FUNCTION_RETURN_SUCCESS
End Function

Public Function Add_PrSpSXLData3(tblTarget() As typ_TBCME020, tblDat As typ_TBCME020, Optional Index As Long = -1) As Integer
    Dim tblIndex As Long
    
    Add_PrSpSXLData3 = FUNCTION_RETURN_FAILURE

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

    Add_PrSpSXLData3 = FUNCTION_RETURN_SUCCESS
End Function

'*** UPDATE ↓ Y.SIMIZU 2005/10/1 渡された品番仕様ﾃﾞｰﾀを構造体に追加
Public Function Add_PrSpSXLData4(tblTarget() As typ_TBCME036, tblDat As typ_TBCME036, Optional Index As Long = -1) As Integer
    Dim tblIndex As Long
    
    Add_PrSpSXLData4 = FUNCTION_RETURN_FAILURE

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

    Add_PrSpSXLData4 = FUNCTION_RETURN_SUCCESS
End Function
'*** UPDATE ↑ Y.SIMIZU 2005/10/1 渡された品番仕様ﾃﾞｰﾀを構造体に追加

'' サンプル管理の結晶内位置から見た上品番、下品番を取得する
Public Function GetUpDownHinban(tUpHin As tFullHinban, _
                                tDownHin As tFullHinban, _
                                tblCrySmp As typ_XSDCS, _
                                tblHinban() As typ_TBCME041) As FUNCTION_RETURN
    Dim Index       As Integer
    Dim iIngPos     As Integer
    Dim iPos2       As Integer
    Dim tblFHin()   As typ_TBCME041
    Dim iHin        As Integer

    GetUpDownHinban = FUNCTION_RETURN_SUCCESS
    
    ClearFullHinban tUpHin
    ClearFullHinban tDownHin
    
    iIngPos = tblCrySmp.INPOSCS
    iHin = 0
    ReDim tblFHin(iHin)
    With tblFHin(iHin)
        .hinban = vbNullString            ' 品番
        .factory = vbNullString           ' 工場
        .opecond = vbNullString           ' 操業条件
    End With
    
    '指定位置に接する品番をリストアップする（「指定位置を含む」ではない）
    For Index = 0 To UBound(tblHinban) - 1
        iPos2 = tblHinban(Index).INGOTPOS + tblHinban(Index).Length
        If (iIngPos >= tblHinban(Index).INGOTPOS) And (iIngPos <= iPos2) Then
            ReDim Preserve tblFHin(iHin)
            tblFHin(iHin) = tblHinban(Index)
            iHin = iHin + 1
        End If
    Next Index

    If UBound(tblFHin) = 0 Then
        SetFullHinban_TBCME041 tUpHin, tblFHin(0)
        SetFullHinban_TBCME041 tDownHin, tblFHin(0)
        Exit Function
    Else
        For Index = 0 To UBound(tblFHin)
            If iIngPos = tblFHin(Index).INGOTPOS Then
                SetFullHinban_TBCME041 tDownHin, tblFHin(Index)
            Else
                SetFullHinban_TBCME041 tUpHin, tblFHin(Index)
            End If
        Next
    End If

    GetUpDownHinban = FUNCTION_RETURN_FAILURE
End Function

'概要      :引上げ終了実績を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :tblTarget     ,I   ,typ_TBCMH004 ,引上げ終了実績テーブル
'          :strCryNum     ,I   ,String       ,結晶番号
'          :戻り値        ,O   ,Integer       ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Public Function GetPlupEndRslt(tblTarget As typ_TBCMH004, strCryNum As String) As Integer
    Dim iRet         As Integer
    Dim tblGet()    As typ_TBCMH004
    
    GetPlupEndRslt = FUNCTION_RETURN_FAILURE

    '' 引上げ終了実績の取得
    iRet = DBDRV_GetTBCMH004(tblGet, "where CRYNUM='" & strCryNum & "'")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function

    tblTarget = tblGet(1)


    GetPlupEndRslt = FUNCTION_RETURN_SUCCESS
End Function
Public Function Add_HinbanManage(tblTarget() As typ_TBCME041, tblDat As typ_TBCME041, Optional Index As Long = -1) As Integer
    Dim tblIndex As Long
    
    Add_HinbanManage = FUNCTION_RETURN_FAILURE

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

    Add_HinbanManage = FUNCTION_RETURN_SUCCESS
End Function
Public Function Add_CrystalSampleManage(tblTarget() As typ_XSDCS, tblDat As typ_XSDCS, Optional Index As Long = -1) As Integer
    Dim tblIndex As Long
    
    Add_CrystalSampleManage = FUNCTION_RETURN_FAILURE

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

    Add_CrystalSampleManage = FUNCTION_RETURN_SUCCESS
End Function

''Upd Start (TCS)T.Terauchi 2005/10/07  GDﾗｲﾝ数表示対応
'概要      :品番より製品仕様ＷＦデータを取得、そして、取得した製品仕様データを追加する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :tblSP()       ,O   ,typ_TBCME036   ,製品仕様ＷＦデータテーブル配列
'          :tHinInf       ,I   ,tFullHinban    ,12桁品番
'          :ctrlFrm       ,I   ,Form           ,フォームID
'          :戻り値        ,O  ,Integer          ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
'履歴      :05/10/07    (TCS)T.Terauchi
Public Function GetSPWFData36(tblSP() As typ_TBCME036, tHinInf As tFullHinban, ctrlFrm As Form) As Integer

    Dim Index       As Integer
    Dim iRet        As Integer
    Dim tHinban(0)  As tFullHinban
    Dim tblGet()    As typ_TBCME036
    
    GetSPWFData36 = FUNCTION_RETURN_FAILURE

    tHinban(0) = tHinInf

    '' 製品仕様ＷＦデータの取得
    iRet = DBDRV_GetTBCME036(tblGet, ctrlFrm.Name, tHinban)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function

    '' 取得した製品仕様ＷＦデータの追加
    '' テーブルデータ格納領域拡張
    ReDim Preserve tblSP(UBound(tblSP) + 1)
    '' テーブルデータ数を取得
    Index = UBound(tblSP)

    '' データ追加
    tblSP(Index) = tblGet(1)
    
    GetSPWFData36 = FUNCTION_RETURN_SUCCESS
    
End Function


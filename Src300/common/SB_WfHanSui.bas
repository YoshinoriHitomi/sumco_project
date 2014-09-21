Attribute VB_Name = "SB_WfHanSui"
Option Explicit

'-------------------------------------------------------------------------------
' 定数定義
'-------------------------------------------------------------------------------
'XSDCW
Private Const cWFSMPLID     As String = "WFSMPLID"      'XSDCWのサンプルＩＤ
Private Const cWFIND        As String = "WFIND"         'XSDCWの状態FLG
Private Const cWFRES        As String = "WFRES"         'XSDCWの実績FLG
Private Const cWFHS         As String = "WFHS"          'XSDCWの保証FLG '追加 05/01/28 ooba
Private Const cCW           As String = "CW"            'XSDCWの項目最終文字
Private Const cWF_RS        As String = "RS"            'XSDCWのRs
Private Const cWF_OI        As String = "OI"            'XSDCWのOi
Private Const cWF_B1        As String = "B1"            'XSDCWのBMD1
Private Const cWF_B2        As String = "B2"            'XSDCWのBMD2
Private Const cWF_B3        As String = "B3"            'XSDCWのBMD3
Private Const cWF_O1        As String = "L1"            'XSDCWのOSF1
Private Const cWF_O2        As String = "L2"            'XSDCWのOSF2
Private Const cWF_O3        As String = "L3"            'XSDCWのOSF3
Private Const cWF_O4        As String = "L4"            'XSDCWのOSF4
Private Const cWF_DS        As String = "DS"            'XSDCWのDS
Private Const cWF_DZ        As String = "DZ"            'XSDCWのDZ
Private Const cWF_SP        As String = "SP"            'XSDCWのSP
Private Const cWF_DO1       As String = "DO1"           'XSDCWのDO1
Private Const cWF_DO2       As String = "DO2"           'XSDCWのDO2
Private Const cWF_DO3       As String = "DO3"           'XSDCWのDO3
Private Const cWF_OT1       As String = "OT1"           'XSDCWのOT1
Private Const cWF_OT2       As String = "OT2"           'XSDCWのOT2
Private Const cWF_AOI       As String = "AOI"           'XSDCWのAOI
Private Const cWF_GD        As String = "GD"            'XSDCWのGD      '追加 05/01/28 ooba
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
Private Const cEPSMPLID     As String = "EPSMPLID"      'XSDCWのサンプルID(エピ項目)
Private Const cEPIND        As String = "EPIND"         'XSDCWの状態FLG(エピ項目)
Private Const cEPRES        As String = "EPRES"         'XSDCWの実績FLG(エピ項目)
Private Const cEP_B1        As String = "B1"            'XSDCWのBMD1E
Private Const cEP_B2        As String = "B2"            'XSDCWのBMD2E
Private Const cEP_B3        As String = "B3"            'XSDCWのBMD3E
Private Const cEP_O1        As String = "L1"            'XSDCWのOSF1E
Private Const cEP_O2        As String = "L2"            'XSDCWのOSF2E
Private Const cEP_O3        As String = "L3"            'XSDCWのOSF3E
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

'------------------------------------------------
' ＷＦ反映/推定チェック共通関数
'------------------------------------------------

'概要      :指定された評価項目№により、反映か推定かを判断し、ＷＦ反映チェック、または、ＷＦ推定チェックを呼び出す。（共通関数）
'           共通関数のチェック結果を当関数の結果として、呼び出し元へ返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :sSXLid        ,I  ,String       :SXL-ID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :結晶番号
'          :tFullHin      ,I  ,tFullHinban  :品番(構造体)
'          :iSmplPos      ,I  ,Integer      :新サンプル位置(mm)
'          :iItemNo       ,I  ,Integer      :評価項目№ =  1 RS     ← 対象外
'                                                       =  2 Oi     ← ﾊﾟﾀｰﾝ1
'                                                       =  3 BMD1   ← ﾊﾟﾀｰﾝ1
'                                                       =  4 BMD2   ← ﾊﾟﾀｰﾝ1
'                                                       =  5 BMD3   ← ﾊﾟﾀｰﾝ1
'                                                       =  6 OSF1   ← ﾊﾟﾀｰﾝ1
'                                                       =  7 OSF2   ← ﾊﾟﾀｰﾝ1
'                                                       =  8 OSF3   ← ﾊﾟﾀｰﾝ1
'                                                       =  9 OSF4   ← ﾊﾟﾀｰﾝ1
'                                                       = 10 DS     ← ﾊﾟﾀｰﾝ1
'                                                       = 11 DZ     ← ﾊﾟﾀｰﾝ1
'                                                       = 12 SP     ← ﾊﾟﾀｰﾝ2
'                                                       = 13 D1     ← ﾊﾟﾀｰﾝ1
'                                                       = 14 D2     ← ﾊﾟﾀｰﾝ1
'                                                       = 15 D3     ← ﾊﾟﾀｰﾝ1
'                                                       = 18 AO     ← ﾊﾟﾀｰﾝ1   '残存酸素追加　03/12/09 ooba
'                                                       = 19 GD     ← ﾊﾟﾀｰﾝ1   'GD追加　05/01/26 ooba
'                                                       = 20 BMD1E  ← ﾊﾟﾀｰﾝ1   '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 21 BMD2E  ← ﾊﾟﾀｰﾝ1   '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 22 BMD3E  ← ﾊﾟﾀｰﾝ1   '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 23 OSF1E  ← ﾊﾟﾀｰﾝ1   '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 24 OSF2E  ← ﾊﾟﾀｰﾝ1   '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 25 OSF3E  ← ﾊﾟﾀｰﾝ1   '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'          :iFromPos      ,I  ,Integer      :検索範囲From
'          :iToPos        ,I  ,Integer      :検索範囲To
'          :iHanSuiKBN    ,O  ,Integer      :反映/推定区分(0:反映,1:推定)
'          :sGetSmplID1   ,O  ,String       :元サンプルID1
'          :sGetSmplID2   ,O  ,String       :元サンプルID2 (反映時未使用)
'          :sGetHSflg1    ,O  ,String       :元サンプルの保証FLG    '追加　05/01/28 ooba
'          :戻り値        ,O  ,Integer      :チェック結果 = 0 : 正常終了(反映/推定OK)
'                                                           1 : 正常終了(反映/推定NG)
'                                                          -1 : 入力引数値エラー
'                                                          -2 : 上記以外のエラー
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funChkWfHanSui(sSXLID As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                                iItemNo As Integer, iFromPos As Integer, iToPos As Integer, iHanSuiKBN As Integer, _
                                sGetSmplID1 As String, sGetSmplID2 As String, Optional sGetHSflg1 As String = "") As Integer
    Dim retCode As Integer
    
    '元サンプルID初期化
    sGetSmplID1 = ""
    sGetSmplID2 = ""
    sGetHSflg1 = "0"     '05/02/18 ooba
    
    'パラメータチェック
    If (Len(sSXLID) <> 13) Then GoTo ChkWfHanSuiParameterErr
    If (Len(sCryNum) <> 12) Then GoTo ChkWfHanSuiParameterErr
    
    '指定された評価項目№により、反映か推定かを判断し、ＷＦ反映チェック、または、ＷＦ推定チェックを呼び出す。
    Select Case iItemNo
    Case 1          'RS(比抵抗)
        retCode = 1
        iHanSuiKBN = 1
'    Case 2 To 15
'    Case 2 To 18
    '残存酸素追加　03/12/09 ooba
    'GD追加　05/01/26 ooba
'    Case 2 To 19
    '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
    Case 2 To 25    'Oi(酸素濃度),BMD1,BMD2,BMD3,OSF1,OSF2,OSF3,OSF4,DS,DZ,SP,D1,D2,D3,--,--,AO,GD,BMD1E,BMD2E,BMD3E,OSF1E,OSF2E,OSF3E
'        retCode = funChkWfHanei(sSXLid, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, iFromPos, iToPos, sGetSmplID1)
        '保証FLG追加　05/01/28 ooba
        retCode = funChkWfHanei(sSXLID, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, iFromPos, iToPos, sGetSmplID1, sGetHSflg1)
        iHanSuiKBN = 0
    Case Else
        GoTo ChkWfHanSuiParameterErr
    End Select
    
    '共通関数のチェック結果を当関数の結果として、呼び出し元へ返す。
    funChkWfHanSui = retCode
    Exit Function

ChkWfHanSuiParameterErr:
    funChkWfHanSui = -1
    Exit Function

ChkWfHanSuiSonotaErr:
    funChkWfHanSui = -2
End Function

'------------------------------------------------
' ＷＦ反映チェック
'------------------------------------------------

'概要      :指定された情報から、ＷＦ反映チェックを行ない結果を返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :sSXLid        ,I  ,String       :SXL-ID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :結晶番号
'          :tFullHin      ,I  ,tFullHinban  :品番(構造体)
'          :iSmplPos      ,I  ,Integer      :新サンプル位置(mm)
'          :iItemNo       ,I  ,Integer      :評価項目№ =  1 RS     ← 対象外
'                                                       =  2 Oi     ← ﾊﾟﾀｰﾝ1
'                                                       =  3 BMD1   ← ﾊﾟﾀｰﾝ1
'                                                       =  4 BMD2   ← ﾊﾟﾀｰﾝ1
'                                                       =  5 BMD3   ← ﾊﾟﾀｰﾝ1
'                                                       =  6 OSF1   ← ﾊﾟﾀｰﾝ1
'                                                       =  7 OSF2   ← ﾊﾟﾀｰﾝ1
'                                                       =  8 OSF3   ← ﾊﾟﾀｰﾝ1
'                                                       =  9 OSF4   ← ﾊﾟﾀｰﾝ1
'                                                       = 10 DS     ← ﾊﾟﾀｰﾝ1
'                                                       = 11 DZ     ← ﾊﾟﾀｰﾝ1
'                                                       = 12 SP     ← ﾊﾟﾀｰﾝ2
'                                                       = 13 D1     ← ﾊﾟﾀｰﾝ1
'                                                       = 14 D2     ← ﾊﾟﾀｰﾝ1
'                                                       = 15 D3     ← ﾊﾟﾀｰﾝ1
'                                                       = 18 AO     ← ﾊﾟﾀｰﾝ1   '残存酸素追加　03/12/09 ooba
'                                                       = 19 GD     ← ﾊﾟﾀｰﾝ1   'GD追加　05/01/26 ooba
'                                                       = 20 BMD1E  ← ﾊﾟﾀｰﾝ1   '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 21 BMD2E  ← ﾊﾟﾀｰﾝ1   '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 22 BMD3E  ← ﾊﾟﾀｰﾝ1   '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 23 OSF1E  ← ﾊﾟﾀｰﾝ1   '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 24 OSF2E  ← ﾊﾟﾀｰﾝ1   '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 25 OSF3E  ← ﾊﾟﾀｰﾝ1   '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'          :iFromPos      ,I  ,Integer      :検索範囲From
'          :iToPos        ,I  ,Integer      :検索範囲To
'          :sGetSmplID    ,O  ,String       :反映元サンプルID
'          :sGetHSflg     ,O  ,String       :反映元サンプルの保証FLG    '追加　05/01/28 ooba
'          :戻り値        ,O  ,Integer      :チェック結果 = 0 : 正常終了(反映OK)
'                                                           1 : 正常終了(反映NG)
'                                                          -1 : 入力引数値エラー
'                                                          -2 : 上記以外のエラー
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funChkWfHanei(sSXLID As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                               iItemNo As Integer, iFromPos As Integer, iToPos As Integer, sGetSmplID As String, sGetHSflg As String) As Integer
    Dim wHPtrn          As Integer
    Dim tSiyou          As type_DBDRV_scmzc_fcmlc001c_Siyou
    Dim wGetSXLid       As String
    Dim wGetSmpKbn      As String
    Dim wGetSmplID      As String
    Dim wGetHSflg       As String       '05/01/28 ooba
    
    Dim tTBCMY013       As typ_TBCMY013
    Dim tGDjisseki      As typ_TBCMJ015 '05/01/31 ooba

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    Dim tTBCMY022       As typ_TBCMY022
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

''Upd start 2005/06/28 (TCS)t.terauchi  SPV9点対応
    Dim tSPVJisseki     As typ_TBCMJ016
    Dim sPos            As String
''Upd end   2005/06/28 (TCS)t.terauchi  SPV9点対応
    
    Dim retJudg         As Boolean
    Dim wIdFlg          As Integer
    Dim TmpData(2)      As String
    
    Dim dShiyo()        As Double       '2003/12/11 Null対応追加
    Dim sHosyo          As String       '2003/12/11 Null対応追加
    
    Dim tSiyou_Sxl      As type_DBDRV_scmzc_fcmkc001c_Siyou '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
    
    '初期化
    wGetSmplID = ""
    wGetHSflg = ""      '05/01/28 ooba
    
    'パラメータチェック
    If (Len(sSXLID) <> 13) Then GoTo ChkWfHaneiParameterErr
    If (Len(sCryNum) <> 12) Then GoTo ChkWfHaneiParameterErr
    
    '指定された評価項目№毎に必要な品番仕様値を取得し、ＷＦ反映値取得パターンを決定する。（指定された評価項目№により、処理が分かれる。）
    Select Case iItemNo
    Case 1              'RS(比抵抗)
        GoTo ChkWfHaneiNG
    Case 2              'Oi(酸素濃度)
        If funGet_TBCME025(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        wHPtrn = 1
        
        'Null対応処理追加 2003/12/11 SystenBrain ▽
        ReDim dShiyo(5)
        dShiyo(1) = tSiyou.HWFONMIN         ' 品ＷＦ酸素濃度下限
        dShiyo(2) = tSiyou.HWFONMAX         ' 品ＷＦ酸素濃度上限
        dShiyo(3) = tSiyou.HWFONMBP         ' 品ＷＦ酸素濃度面内分布
        dShiyo(4) = tSiyou.HWFONAMN         ' 品ＷＦ酸素濃度平均下限
        dShiyo(5) = tSiyou.HWFONAMX         ' 品ＷＦ酸素濃度平均上限
        If fncJissekiHantei_nl(tSiyou.HWFONHWS, dShiyo) = False Then GoTo ChkWfHaneiSonotaErr
        'Null対応処理追加 2003/12/11 SystenBrain △
        
    Case 3 To 9         'BMD1,BMD2,BMD3,OSF1,OSF2,OSF3,OSF4
        If funGet_TBCME029(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        wHPtrn = 1
        
        'Null対応処理追加 2003/12/11 SystenBrain ▽
        'BMD仕様NULLチェックを削除（判定でOKとする。）　          2003/12/19 tuku
''''        ReDim dShiyo(1)
''''        If iItemNo = 3 Then         'BMD1
''''            sHosyo = tSiyou.HWFBM1HS            ' 品ＷＦＢＭＤ１保証方法＿処
''''            dShiyo(1) = tSiyou.HWFBM1MBP        ' 品ＷＦＢＭＤ１面内分布
''''        ElseIf iItemNo = 4 Then     'BMD2
''''            sHosyo = tSiyou.HWFBM2HS            ' 品ＷＦＢＭＤ２保証方法＿処
''''            dShiyo(1) = tSiyou.HWFBM2MBP        ' 品ＷＦＢＭＤ２面内分布
''''        ElseIf iItemNo = 5 Then     'BMD3
''''            sHosyo = tSiyou.HWFBM3HS            ' 品ＷＦＢＭＤ３保証方法＿処
''''            dShiyo(1) = tSiyou.HWFBM3MBP        ' 品ＷＦＢＭＤ３面内分布
''''        ElseIf iItemNo = 6 Then     'OSF1
''''        ElseIf iItemNo = 7 Then     'OSF2
''''        ElseIf iItemNo = 8 Then     'OSF3
''''        ElseIf iItemNo = 9 Then     'OSF4
''''        End If
''''        If fncJissekiHantei_nl(sHosyo, dShiyo) = False Then GoTo ChkWfHaneiSonotaErr
        'Null対応処理追加 2003/12/11 SystenBrain △
        
    Case 10             'DSOD
        If funGet_TBCME026(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        wHPtrn = 1
        
        'Null対応処理追加 2003/12/11 SystenBrain ▽
        ReDim dShiyo(4)
        dShiyo(1) = tSiyou.HWFDSOMX         ' 品ＷＦＤＳＯＤ上限
        dShiyo(2) = tSiyou.HWFDSOMN         ' 品ＷＦＤＳＯＤ下限
        dShiyo(3) = tSiyou.HWFDSOAX         ' 品ＷＦＤＳＯＤ領域上限
        dShiyo(4) = tSiyou.HWFDSOAN         ' 品ＷＦＤＳＯＤ領域下限
        If fncJissekiHantei_nl(tSiyou.HWFDSOHS, dShiyo) = False Then GoTo ChkWfHaneiSonotaErr
        'Null対応処理追加 2003/12/11 SystenBrain △
        
    Case 11             'DZ幅
        If funGet_TBCME024(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        wHPtrn = 1
    Case 12             'SPVFE
        If funGet_TBCME028(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
'        wHPtrn = 2
        'SPVの反映ﾊﾟﾀｰﾝを1に変更　04/04/27 ooba
        wHPtrn = 1
    Case 13, 14, 15     'DOI1(酸素析出1),DOI2(酸素析出2),DOI3(酸素析出3)
        If funGet_TBCME025(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        wHPtrn = 1
    Case 18             'AO     '残存酸素追加　03/12/09 ooba
        If funGet_TBCME025(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        wHPtrn = 1
    Case 19             'GD     'GD追加　05/01/26 ooba
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
        If funGet_TBCME020(tFullHin, tSiyou_Sxl) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG

        If funGet_TBCME036(tFullHin, tSiyou_Sxl) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        
        tSiyou.HSXGDPTK = tSiyou_Sxl.HSXGDPTK
        tSiyou.HSXLDLRMN = tSiyou_Sxl.HSXLDLRMN
        tSiyou.HSXLDLRMX = tSiyou_Sxl.HSXLDLRMX
        tSiyou.HWFLDLRMN = tSiyou_Sxl.HWFLDLRMN
        tSiyou.HWFLDLRMX = tSiyou_Sxl.HWFLDLRMX
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
        
        If funGet_TBCME026(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        
    '*** UPDATE ↓ Y.SIMIZU 2005/10/12 GDﾗｲﾝ数取得
        If funGet_TBCME036_2(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
    '*** UPDATE ↑ Y.SIMIZU 2005/10/12 GDﾗｲﾝ数取得
        wHPtrn = 1
            
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    Case 20 To 25
        If funGet_TBCME050(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        wHPtrn = 1
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

    Case Else
        GoTo ChkWfHaneiParameterErr
    End Select

    'ＷＦ反映元サンプルＩＤの取得
    If wHPtrn = 1 Then              '結晶反映値取得パターン１
'        If funGetWfHanei1(sSXLid, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, iFromPos, iToPos,
'                                                                    wGetSXLid, wGetSmpKbn, wGetSmplID) <> 0 Then GoTo ChkWfHaneiNG
'        '保証FLG追加　05/01/28 ooba
'        If funGetWfHanei1(sSXLid, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, iFromPos, iToPos, _
'                                                                    wGetSXLid, wGetSmpKbn, wGetSmplID, wGetHSflg) <> 0 Then GoTo ChkWfHaneiNG

        If funGetWfHanei1(sSXLID, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, iFromPos, iToPos, _
                                                                    wGetSXLid, wGetSmpKbn, wGetSmplID, wGetHSflg) <> 0 Then
            '' 結晶GD反映対応　05/06/13 ooba START ======================================>
            'GD
            If iItemNo = 19 Then
                'TOPの場合
                If sTB = "T" And IsNumeric(CrySampleID.TsmplidGD) Then
                    '結晶ｻﾝﾌﾟﾙIDと保証FLG(1:結晶保証)をｾｯﾄ
                    wGetSmplID = CrySampleID.TsmplidGD
                    wGetHSflg = "1"
                'BOTの場合
                ElseIf sTB = "B" And IsNumeric(CrySampleID.BsmplidGD) Then
                    '結晶ｻﾝﾌﾟﾙIDと保証FLG(1:結晶保証)をｾｯﾄ
                    wGetSmplID = CrySampleID.BsmplidGD
                    wGetHSflg = "1"
                Else
                    GoTo ChkWfHaneiNG
                End If
            Else
                GoTo ChkWfHaneiNG
            End If
            '' 結晶GD反映対応　05/06/13 ooba END ========================================>
        End If
        
    ElseIf wHPtrn = 2 Then          '結晶反映値取得パターン２
        If funGetWfHanei2(sSXLID, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, iFromPos, iToPos, _
                                                                    wGetSXLid, wGetSmpKbn, wGetSmplID) <> 0 Then GoTo ChkWfHaneiNG
    End If
    
    '結晶反映元ｻﾝﾌﾟﾙIDから、結晶反映値（実績値）を取得する。（指定された評価項目№により、処理が分かれる。）
    Select Case iItemNo
'    Case 1              'RS(比抵抗)
'        GoTo ChkWfHaneiNG
    Case 2              'Oi(酸素濃度)
        'Oiの実績値を取得する
        If funGetTBCMY013(wGetSmplID, "OI", "OI", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
        'Oi総合判定を行なう
        If Not WfCrOiJudg(tSiyou, tTBCMY013, retJudg) Then GoTo ChkWfHaneiNG

    Case 3, 4, 5                'BMD1, BMD2, BMD3
        If iItemNo = 3 Then
            'BMD1の実績値を取得する
            If funGetTBCMY013(wGetSmplID, "BMD", "BMD1", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 1
        ElseIf iItemNo = 4 Then
            'BMD2の実績値を取得する
            If funGetTBCMY013(wGetSmplID, "BMD", "BMD2", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 2
        ElseIf iItemNo = 5 Then
            'BMD3の実績値を取得する
            If funGetTBCMY013(wGetSmplID, "BMD", "BMD3", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 3
        End If
        'BMDの総合判定を行なう
        If Not WfCrBmdJudg(tSiyou, tTBCMY013, retJudg, wIdFlg) Then GoTo ChkWfHaneiNG
        
    Case 6, 7, 8, 9             'OSF1, OSF2, OSF3, OSF4
        If iItemNo = 6 Then
            'OSF1の実績値を取得する
            If funGetTBCMY013(wGetSmplID, "OSF", "OSF1", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 1
        ElseIf iItemNo = 7 Then
            'OSF2の実績値を取得する
            If funGetTBCMY013(wGetSmplID, "OSF", "OSF2", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 2
        ElseIf iItemNo = 8 Then
            'OSF3の実績値を取得する
            If funGetTBCMY013(wGetSmplID, "OSF", "OSF3", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 3
        ElseIf iItemNo = 9 Then
            'OSF4の実績値を取得する
            If funGetTBCMY013(wGetSmplID, "OSF", "OSF4", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 4
        End If
        'OSFの総合判定を行なう
        If Not WfCrOsfJudg(tSiyou, tTBCMY013, retJudg, wIdFlg, TmpData) Then GoTo ChkWfHaneiNG
    
    Case 10             'DSOD
        'DSODの実績値を取得する
        If funGetTBCMY013(wGetSmplID, "DSOD", "DSOD", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
    
        'DSOD総合判定を行なう
        If Not WfCrDsodjudg(tSiyou, tTBCMY013, retJudg) Then GoTo ChkWfHaneiNG
    
    Case 11             'DZ幅
        'DZ幅の実績値を取得する
        If funGetTBCMY013(wGetSmplID, "DZ", "DZ", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
    
        'DZ幅総合判定を行なう
        If Not WfCrDzjudg(tSiyou, tTBCMY013, retJudg) Then GoTo ChkWfHaneiNG
    
    Case 12             'SPVFE
        
    ''Upd start 2005/06/28 (TCS)t.terauchi  SPV9点対応
'        'SPVFEの実績値を取得する
'        If funGetTBCMY013(wGetSmplID, "SPV", "SPV", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
'
'        'SPVFE総合判定を行なう
'        If Not WfCrSpvjudg(tSiyou, tTBCMY013, retJudg) Then GoTo ChkWfHaneiNG
        
'↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
        '' WF仕様(SPV)取得
        If funWfcGetDataEtc_SPV(tFullHin, _
                                tSiyou) <> FUNCTION_RETURN_SUCCESS Then GoTo ChkWfHaneiNG
'↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                
        'SPVFEの実績値を取得する
        If funGetSPVJisseki_J016(sCryNum, wGetSmplID, _
                        tSPVJisseki, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
        
        
        '実績ﾃﾞｰﾀなし
        If Trim(tSPVJisseki.SMPLNO) = "0" Then GoTo ChkWfHaneiNG
        
        If sTB = "T" Then
            sPos = "TOP"
        ElseIf sTB = "B" Then
            sPos = "BOT"
        Else
            GoTo ChkWfHaneiParameterErr
        End If
        
        'SPV(Fe濃度)総合判定を行なう
        If ((tSiyou.HWFSPVHS = "H") And CheckKHN(tSiyou.HWFSPVKN, 15, sPos)) Then
            If Not WfCrSpvJudg_New(tSiyou, tSPVJisseki, retJudg, 1, sPos) Then GoTo ChkWfHaneiNG
        Else
            retJudg = True
        End If
        
        If retJudg = True Then
            'SPV(拡散長)総合判定を行なう
            If ((tSiyou.HWFDLHWS = "H") And CheckKHN(tSiyou.HWFDLKHN, 16, sPos)) Then
                If Not WfCrSpvJudg_New(tSiyou, tSPVJisseki, retJudg, 2, sPos) Then GoTo ChkWfHaneiNG
            End If
        End If
    ''Upd end   2005/06/28 (TCS)t.terauchi  SPV9点対応

'↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
        If retJudg = True Then
            'SPV(Nr濃度)総合判定を行なう
            If ((tSiyou.HWFNRHS = "H") And CheckKHN(tSiyou.HWFNRKN, 19, sPos)) Then
                If Not WfCrSpvJudg_New(tSiyou, tSPVJisseki, retJudg, 3, sPos) Then GoTo ChkWfHaneiNG
            Else
                retJudg = True
            End If
        End If
'↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------


    
    Case 13, 14, 15             'DOI1(酸素析出1),DOI2(酸素析出2),DOI3(酸素析出3)
        If iItemNo = 13 Then
            'DOI1の実績値を取得する
            If funGetTBCMY013(wGetSmplID, "DOI", "DOI1", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 1
        ElseIf iItemNo = 14 Then
            'DOI2の実績値を取得する
            If funGetTBCMY013(wGetSmplID, "DOI", "DOI2", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 2
        ElseIf iItemNo = 15 Then
            'DOI3の実績値を取得する
            If funGetTBCMY013(wGetSmplID, "DOI", "DOI2", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 3
        End If
        'DOIの総合判定を行なう
        If Not WfCrDoiJudg(tSiyou, tTBCMY013, retJudg, wIdFlg) Then GoTo ChkWfHaneiNG
    
    ''残存酸素追加　03/12/09 ooba START ==================================================>
    Case 18             'AOi
        'AOiの実績値を取得する
        If funGetTBCMY013(wGetSmplID, "AOI", "AOI", tTBCMY013) <> 0 Then GoTo ChkWfHaneiNG
        
        'AOiの総合判定を行なう
        If Not WfCrAoiJudg(tSiyou, tTBCMY013, retJudg) Then GoTo ChkWfHaneiNG
    ''残存酸素追加　03/12/09 ooba END ====================================================>
    
    ''GD追加　05/01/31 ooba START ========================================================>
    Case 19             'GD
        If wGetHSflg = "1" Then
            'GDの結晶実績値を取得する
            If funGetGDJisseki_J006(sCryNum, wGetSmplID, tGDjisseki) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
            
            tSiyou.WFHSGDCW = "1"   '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
        Else
            'GDのWF実績値を取得する
            If funGetGDJisseki_J015(sCryNum, wGetSmplID, wGetHSflg, tGDjisseki) = FUNCTION_RETURN_FAILURE Then GoTo ChkWfHaneiNG
            
            tSiyou.WFHSGDCW = "0"   '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
        End If
        '実績ﾃﾞｰﾀなし
        If Trim(tGDjisseki.SMPLNO) = "" Then GoTo ChkWfHaneiNG
        
        'GDの総合判定を行なう
        If Not WfCrGdJudg(tSiyou, tGDjisseki, retJudg) Then GoTo ChkWfHaneiNG
    ''GD追加　05/01/31 ooba END ==========================================================>

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    Case 20 To 22
        If iItemNo = 20 Then
            'BMD1(EP)の実績値を取得する
            If funGetTBCMY022(wGetSmplID, "BMD", "BMD1", tTBCMY022) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 1
        ElseIf iItemNo = 21 Then
            'BMD2(EP)の実績値を取得する
            If funGetTBCMY022(wGetSmplID, "BMD", "BMD2", tTBCMY022) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 2
        ElseIf iItemNo = 22 Then
            'BMD3(EP)の実績値を取得する
            If funGetTBCMY022(wGetSmplID, "BMD", "BMD3", tTBCMY022) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 3
        End If
        'BMDの総合判定を行なう
        If Not EpBmdJudg(tSiyou, tTBCMY022, retJudg, wIdFlg) Then GoTo ChkWfHaneiNG
    Case 23 To 25
        If iItemNo = 23 Then
            'OSF1(EP)の実績値を取得する
            If funGetTBCMY022(wGetSmplID, "OSF", "OSF1", tTBCMY022) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 1
        ElseIf iItemNo = 24 Then
            'OSF2(EP)の実績値を取得する
            If funGetTBCMY022(wGetSmplID, "OSF", "OSF2", tTBCMY022) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 2
        ElseIf iItemNo = 25 Then
            'OSF3(EP)の実績値を取得する
            If funGetTBCMY022(wGetSmplID, "OSF", "OSF3", tTBCMY022) <> 0 Then GoTo ChkWfHaneiNG
            wIdFlg = 3
        End If
        'OSFの総合判定を行なう
        If Not EpOsfJudg(tSiyou, tTBCMY022, retJudg, wIdFlg, TmpData) Then GoTo ChkWfHaneiNG

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

'    Case Else
'        GoTo ChkWfHaneiParameterErr
    End Select
    
    '指定された評価項目№の総合判定がOKの場合、反映元サンプルIDを設定し、戻り値に'0'(正常終了(反映OK))を設定し、処理を終了する。
    '総合判定がNGの場合、戻り値に'1'(正常終了(反映NG))を設定し、処理を終了する。
    If retJudg = False Then GoTo ChkWfHaneiNG
        
    sGetSmplID = wGetSmplID
    sGetHSflg = wGetHSflg       '05/01/28 ooba
    funChkWfHanei = 0
    Exit Function

ChkWfHaneiNG:
    sGetSmplID = wGetSmplID
    sGetHSflg = wGetHSflg       '05/01/28 ooba
    funChkWfHanei = 1
    Exit Function

ChkWfHaneiParameterErr:
    funChkWfHanei = -1
    Exit Function

ChkWfHaneiSonotaErr:
    funChkWfHanei = -2
End Function

'------------------------------------------------
' ＷＦ反映値取得（パターン１）
'------------------------------------------------

'概要      :指定された新サンプル位置情報から、ＷＦ反映元サンプルＩＤを新サンプル管理(SXL)(XSDCW)より検索し、結果を返す。
'           反映しようとする新サンプル位置が、TOPの場合とBOTの場合で検索方法(方向)が異なる。
'           反映元サンプルＩＤを検索する場合、基本的には、新サンプル位置から見て、上下サンプルの中で近いほうのサンプルＩＤを抽出する。
'           検索する際の検索範囲は、指定された範囲内のみ有効とし、検索範囲内にみつからない場合、「該当ｻﾝﾌﾟﾙなし」とする。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :sSXLid        ,I  ,String       :SXL-ID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :結晶番号
'          :tFullHin      ,I  ,tFullHinban  :品番(構造体)
'          :iSmplPos      ,I  ,Integer      :新サンプル位置(mm)
'          :iItemNo       ,I  ,Integer      :評価項目№ =  1 RS
'                                                       =  2 Oi     ←対象
'                                                       =  3 BMD1   ←対象
'                                                       =  4 BMD2   ←対象
'                                                       =  5 BMD3   ←対象
'                                                       =  6 OSF1   ←対象
'                                                       =  7 OSF2   ←対象
'                                                       =  8 OSF3   ←対象
'                                                       =  9 OSF4   ←対象
'                                                       = 10 DS     ←対象
'                                                       = 11 DZ     ←対象
'                                                       = 12 SP
'                                                       = 13 D1     ←対象
'                                                       = 14 D2     ←対象
'                                                       = 15 D3     ←対象
'                                                       = 18 AO     ←対象  '追加　03/12/09 ooba
'                                                       = 19 GD     ←対象  '追加　05/01/26 ooba
'                                                       = 20 BMD1E  ←対象  '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 21 BMD2E  ←対象  '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 22 BMD3E  ←対象  '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 23 OSF1E  ←対象  '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 24 OSF2E  ←対象  '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 25 OSF3E  ←対象  '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'          :iFromPos      ,I  ,Integer      :検索範囲From
'          :iToPos        ,I  ,Integer      :検索範囲To
'          :sGetSXLid     ,O  ,String       :反映元SXL-ID
'          :sGetSmpKbn    ,O  ,String       :反映元サンプル区分
'          :sGetSmplID    ,O  ,String       :反映元サンプルＩＤ
'          :sGetHSflg     ,O  ,String       :反映元サンプルの保証FLG    '追加　05/01/28 ooba
'          :戻り値        ,O  ,Integer      :取得結果 = 0 : 正常終了
'                                                       1 : 正常終了(該当サンプルなし)
'                                                      -1 : 異常終了
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funGetWfHanei1(sSXLID As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                               iItemNo As Integer, iFromPos As Integer, iToPos As Integer, _
                               sGetSXLid As String, sGetSmpKbn As String, sGetSmplID As String, sGetHSflg As String) As Integer
    Dim kName       As String
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim ediSmpid    As String       'ｻﾝﾌﾟﾙID名称
    Dim ediInd      As String       '状態FLG名称
    Dim ediRes      As String       '実績FLG名称
    Dim ediHs       As String       '保証FLG名称    '05/01/28 ooba
    
    'パラメータチェック
    If (Len(sSXLID) <> 13) Then GoTo GetWfHanei1ParameterErr
    If (Len(sCryNum) <> 12) Then GoTo GetWfHanei1ParameterErr
    
    '指定された評価項目№から、検索対照評価項目名を決定する。
    kName = funGetWfKensaName(iItemNo)
    If kName = " " Then GoTo GetWfHanei1ParameterErr
        
    Select Case iItemNo
    Case 20 To 25
        'SQL文内で使用する名称に編集
        ediSmpid = cEPSMPLID & kName & cCW     'ｻﾝﾌﾟﾙID
        ediInd = cEPIND & kName & cCW          '状態FLG
        ediRes = cEPRES & kName & cCW          '実績FLG
    Case Else
        'SQL文内で使用する名称に編集
        ediSmpid = cWFSMPLID & kName & cCW     'ｻﾝﾌﾟﾙID
        ediInd = cWFIND & kName & cCW          '状態FLG
        ediRes = cWFRES & kName & cCW          '実績FLG
    End Select
    
    '保証ﾌﾗｸﾞ設定　05/01/28 ooba
    Select Case iItemNo
    Case 19     'GD
        ediHs = cWFHS & kName & cCW        '保証FLG
    Case Else
        ediHs = "'0'"
    End Select
    
    
    '指定された情報を元に、新ｻﾝﾌﾟﾙ管理(SXL)(XSDCW)を検索する。
'    sql = "select SXLIDCW, SMPKBNCW, " & ediSmpid & " as SMPLID from XSDCW "
    '保証ﾌﾗｸﾞ追加　05/01/28 ooba
    sql = "select "
    sql = sql & "SXLIDCW, "
    sql = sql & "SMPKBNCW, "
    sql = sql & ediSmpid & " as SMPLID, "
    sql = sql & ediHs & " as HSFLG "
    sql = sql & "from XSDCW "

    'TOP位置(T/B区分='T')の検索
    If sTB = "T" Then
        sql = sql & "where tbkbncw = '" & sTB & "' and "
        sql = sql & "      xtalcw = '" & sCryNum & "' and "
        sql = sql & "      inposcw <= " & iSmplPos & " and "
        sql = sql & "      (" & ediInd & " = '1' or " & ediInd & " = '2') and "
        sql = sql & "  " & ediRes & " <> '0' and "
        sql = sql & "      inposcw >= " & iFromPos & " and "
        sql = sql & "      inposcw <= " & iToPos & " "
        sql = sql & "order by inposcw desc"
    
    'BOT位置(T/B区分='B')の検索
    ElseIf sTB = "B" Then
        sql = sql & "where tbkbncw = '" & sTB & "' and "
        sql = sql & "      xtalcw = '" & sCryNum & "' and "
        sql = sql & "      inposcw >= " & iSmplPos & " and "
        sql = sql & "      (" & ediInd & " = '1' or " & ediInd & " = '2') and "
        sql = sql & "  " & ediRes & " <> '0' and "
        sql = sql & "      inposcw >= " & iFromPos & " and "
        sql = sql & "      inposcw <= " & iToPos & " "
        sql = sql & "order by inposcw asc"
    Else
        GoTo GetWfHanei1ParameterErr
    End If
    
    'SQL文の実行
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetWfHanei1 = 1
        Set rs = Nothing
        Exit Function
    End If
    
    '呼び出し元への結果通知
    sGetSXLid = rs("SXLIDCW")
    sGetSmpKbn = rs("SMPKBNCW")
    sGetSmplID = rs("SMPLID")
    sGetHSflg = rs("HSFLG")     '05/01/28 ooba
    Set rs = Nothing
    
    funGetWfHanei1 = 0
    Exit Function

GetWfHanei1ParameterErr:
    funGetWfHanei1 = -1
End Function

'------------------------------------------------
' ＷＦ反映値取得（パターン２）
'------------------------------------------------

'概要      :指定された新サンプル位置情報から、ＷＦ反映元サンプルＩＤを新サンプル管理(SXL)(XSDCW)より検索し、結果を返す。
'           反映元サンプルＩＤを検索する場合、基本的には、新サンプル位置から見て、下サンプルの中で近いほうのサンプルＩＤを抽出する。
'           検索する際の検索範囲は、指定された範囲内のみ有効とし、検索範囲内にみつからない場合、「該当ｻﾝﾌﾟﾙなし」とする。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :sSXLid        ,I  ,String       :SXL-ID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :結晶番号
'          :tFullHin      ,I  ,tFullHinban  :品番(構造体)
'          :iSmplPos      ,I  ,Integer      :新サンプル位置(mm)
'          :iItemNo       ,I  ,Integer      :評価項目№ =  1 RS
'                                                       =  2 Oi
'                                                       =  3 BMD1
'                                                       =  4 BMD2
'                                                       =  5 BMD3
'                                                       =  6 OSF1
'                                                       =  7 OSF2
'                                                       =  8 OSF3
'                                                       =  9 OSF4
'                                                       = 10 DS
'                                                       = 11 DZ
'                                                       = 12 SP     ←対象
'                                                       = 13 D1
'                                                       = 14 D2
'                                                       = 15 D3
'                                                       = 18 AO
'                                                       = 19 GD
'                                                       = 20 BMD1E  '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 21 BMD2E  '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 22 BMD3E  '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 23 OSF1E  '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 24 OSF2E  '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'                                                       = 25 OSF3E  '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
'          :iFromPos      ,I  ,Integer      :検索範囲From
'          :iToPos        ,I  ,Integer      :検索範囲To
'          :sGetSXLid     ,O  ,String       :反映元SXL-ID
'          :sGetSmpKbn    ,O  ,String       :反映元サンプル区分
'          :sGetSmplID    ,O  ,String       :反映元サンプルＩＤ
'          :戻り値        ,O  ,Integer      :取得結果 = 0 : 正常終了
'                                                       1 : 正常終了(該当サンプルなし)
'                                                      -1 : 異常終了
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funGetWfHanei2(sSXLID As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                               iItemNo As Integer, iFromPos As Integer, iToPos As Integer, _
                               sGetSXLid As String, sGetSmpKbn As String, sGetSmplID As String) As Integer
    Dim kName       As String
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim ediSmpid    As String       'ｻﾝﾌﾟﾙID名称
    Dim ediInd      As String       '状態FLG名称
    Dim ediRes      As String       '実績FLG名称
    
    'パラメータチェック
    If (Len(sSXLID) <> 13) Then GoTo GetWfHanei2ParameterErr
    If (Len(sCryNum) <> 12) Then GoTo GetWfHanei2ParameterErr
    
    '指定された評価項目№から、検索対照評価項目名を決定する。
    kName = funGetWfKensaName(iItemNo)
    If kName = " " Then GoTo GetWfHanei2ParameterErr
    

    Select Case iItemNo
    Case 20 To 25
        'SQL文内で使用する名称に編集
        ediSmpid = cEPSMPLID & kName & cCW     'ｻﾝﾌﾟﾙID
        ediInd = cEPIND & kName & cCW          '状態FLG
        ediRes = cEPRES & kName & cCW          '実績FLG
    Case Else
        'SQL文内で使用する名称に編集
        ediSmpid = cWFSMPLID & kName & cCW     'ｻﾝﾌﾟﾙID
        ediInd = cWFIND & kName & cCW          '状態FLG
        ediRes = cWFRES & kName & cCW          '実績FLG
    End Select
    
    '指定された情報を元に、新ｻﾝﾌﾟﾙ管理(SXL)(XSDCW)を検索する。
    sql = "select SXLIDCW, SMPKBNCW, " & ediSmpid & " as SMPLID from XSDCW "
    sql = sql & "where xtalcw = '" & sCryNum & "' and "
    'SPVは必ずBOT側から取得するように変更　04/04/23 ooba
    If kName = "SP" Then sql = sql & "TBKBNCW = 'B' and "
    sql = sql & "      inposcw > " & iSmplPos & " and "
    sql = sql & "      (" & ediInd & " = '1' or " & ediInd & " = '2') and "
    sql = sql & "  " & ediRes & " <> '0' and "
    sql = sql & "      inposcw >= " & iFromPos & " and "
    sql = sql & "      inposcw <= " & iToPos & " "
    sql = sql & "order by inposcw asc"
    
    'SQL文の実行
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetWfHanei2 = 1
        Set rs = Nothing
        Exit Function
    End If
    
    '呼び出し元への結果通知
    sGetSXLid = rs("SXLIDCW")
    sGetSmpKbn = rs("SMPKBNCW")
    sGetSmplID = rs("SMPLID")
    Set rs = Nothing
    
    funGetWfHanei2 = 0
    Exit Function
    
GetWfHanei2ParameterErr:
    funGetWfHanei2 = -1
End Function

'------------------------------------------------
' ＷＦ検査対象評価項目名取得
'------------------------------------------------

'概要      :評価項目№から、ＷＦ検査対象評価項目名を返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :iItemNo       ,I  ,Integer      :評価項目№ ⇒ Sxl =  1 RS
'                                                              =  2 Oi
'                                                              =  3 BMD1
'                                                              =  4 BMD2
'                                                              =  5 BMD3
'                                                              =  6 OSF1
'                                                              =  7 OSF2
'                                                              =  8 OSF3
'                                                              =  9 OSF4
'                                                              = 10 DS
'                                                              = 11 DZ
'                                                              = 12 SP
'                                                              = 13 DO1
'                                                              = 14 DO2
'                                                              = 15 DO3
'                                                              = 18 AO     '追加　03/12/09 ooba
'                                                              = 19 GD     '追加　05/01/28 ooba
'                                                              = 20 BMD1E  '2006/08/15 エピ先行評価追加対応 SMP)kondoh
'                                                              = 21 BMD2E  '2006/08/15 エピ先行評価追加対応 SMP)kondoh
'                                                              = 22 BMD3E  '2006/08/15 エピ先行評価追加対応 SMP)kondoh
'                                                              = 23 OSF1E  '2006/08/15 エピ先行評価追加対応 SMP)kondoh
'                                                              = 24 OSF2E  '2006/08/15 エピ先行評価追加対応 SMP)kondoh
'                                                              = 25 OSF3E  '2006/08/15 エピ先行評価追加対応 SMP)kondoh
'          :戻り値        ,O  ,Sting        :検査対象項目名(ﾊﾟﾗﾒｰﾀｴﾗｰ時は、空白を返す)
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funGetWfKensaName(iItemNo As Integer) As String
    
    'パラメータチェック
'    If iItemNo < 1 Or iItemNo > 15 Then GoTo GetWfKensaNameParameterErr
    ''残存酸素検査項目追加による変更　03/12/15 ooba
'    If iItemNo < 1 Or iItemNo > 18 Then GoTo GetWfKensaNameParameterErr
    'GD追加による変更　05/01/28 ooba
'    If iItemNo < 1 Or iItemNo > 19 Then GoTo GetWfKensaNameParameterErr
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    If iItemNo < 1 Or iItemNo > 25 Then GoTo GetWfKensaNameParameterErr
    
    'SXL
    Select Case iItemNo
    Case 1:     funGetWfKensaName = cWF_RS        'RS(比抵抗)
    Case 2:     funGetWfKensaName = cWF_OI        'Oi(酸素濃度)
    Case 3:     funGetWfKensaName = cWF_B1        'BMD1
    Case 4:     funGetWfKensaName = cWF_B2        'BMD2
    Case 5:     funGetWfKensaName = cWF_B3        'BMD3
    Case 6:     funGetWfKensaName = cWF_O1        'OSF1
    Case 7:     funGetWfKensaName = cWF_O2        'OSF2
    Case 8:     funGetWfKensaName = cWF_O3        'OSF3
    Case 9:     funGetWfKensaName = cWF_O4        'OSF4
    Case 10:    funGetWfKensaName = cWF_DS        'CS(炭素濃度)
    Case 11:    funGetWfKensaName = cWF_DZ        'GD
    Case 12:    funGetWfKensaName = cWF_SP        'LT(ﾗｲﾌﾀｲﾑ)
    Case 13:    funGetWfKensaName = cWF_DO1       'EPD
    Case 14:    funGetWfKensaName = cWF_DO2       'LT(ﾗｲﾌﾀｲﾑ)
    Case 15:    funGetWfKensaName = cWF_DO3       'EPD
    Case 18:    funGetWfKensaName = cWF_AOI       'AOi      '残存酸素追加　03/12/15 ooba
    Case 19:    funGetWfKensaName = cWF_GD        'GD       'GD追加　05/01/28 ooba
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    Case 20:    funGetWfKensaName = cEP_B1        'BMD1E
    Case 21:    funGetWfKensaName = cEP_B2        'BMD2E
    Case 22:    funGetWfKensaName = cEP_B3        'BMD3E
    Case 23:    funGetWfKensaName = cEP_O1        'OSF1E
    Case 24:    funGetWfKensaName = cEP_O2        'OSF2E
    Case 25:    funGetWfKensaName = cEP_O3        'OSF3E
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    End Select
    
    Exit Function

GetWfKensaNameParameterErr:
    funGetWfKensaName = " "
End Function

'------------------------------------------------
' 測定評価結果(WFの各種実績)取得関数
'------------------------------------------------

'概要      :サンプルＩＤから、TBCMY013を検索し、測定評価結果(WFの各種実績)を取得し返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                   :説明
'          :sSmplID       ,I  ,String               :サンプルＩＤ
'          :sItemName     ,I  ,String               :評価項目名称(RES,OI,BMD,OSF,DSOD,DZ,SPV,DOI)
'          :sItemDetail   ,I  ,String               :評価項目詳細名称(RES,OI,BMD1～BMD3,OSF1～OSF4,DSOD,DZ,SPV,DOI1～DOI3)
'          :tTBCMY013     ,O  ,typ_TBCMY013         :測定評価結果(構造体)
'          :戻り値        ,O  ,Integer              :取得結果 = 0 : 正常
'                                                              -1 : 異常
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funGetTBCMY013(sSmplID As String, sItemName As String, sItemDetail As String, tTBCMY013 As typ_TBCMY013) As Integer

    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    
    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "SB_WfHanSui.bas -- Function funGetTBCMY013"
    
    'サンプルＩＤ等からTBCMY013の測定評価結果(WFの各種実績)を検索する。
    sql = "select SAMPLEID, OSITEM, MAISU, SPEC, NETSU, ET, MES, DKAN, MESDATA1, MESDATA2, MESDATA3, MESDATA4, MESDATA5, "
    sql = sql & "MESDATA6, MESDATA7, MESDATA8, MESDATA9, MESDATA10, MESDATA11, MESDATA12, MESDATA13, MESDATA14, MESDATA15, "
    sql = sql & "TXID, REGDATE, SENDFLAG, SENDDATE "
    sql = sql & "from TBCMY013 "
    sql = sql & "where SAMPLEID = '" & sSmplID & "' and "
    sql = sql & "      OSITEM = '" & sItemName & "' and "
    sql = sql & "      SPEC = '" & sItemDetail & "'"
    
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetTBCMY013 = -1
        GoTo PROC_EXIT
    End If
    
     ''抽出結果を格納する
    With tTBCMY013
        .SAMPLEID = rs("SAMPLEID")       ' サンプルID
        .OSITEM = rs("OSITEM")           ' 評価項目
        .MAISU = rs("MAISU")             ' 評価枚数
        .Spec = rs("SPEC")               ' 規格値
        .NETSU = rs("NETSU")             ' 熱処理条件
        .ET = rs("ET")                   ' エッチング条件
        .MES = rs("MES")                 ' 計測方法
        .DKAN = rs("DKAN")               ' ＤＫアニール条件
        .MESDATA1 = rs("MESDATA1")       ' 測定データその１
        .MESDATA2 = rs("MESDATA2")       ' 測定データその２
        .MESDATA3 = rs("MESDATA3")       ' 測定データその３
        .MESDATA4 = rs("MESDATA4")       ' 測定データその４
        .MESDATA5 = rs("MESDATA5")       ' 測定データその５
        .MESDATA6 = rs("MESDATA6")       ' 測定データその６
        .MESDATA7 = rs("MESDATA7")       ' 測定データその７
        .MESDATA8 = rs("MESDATA8")       ' 測定データその８
        .MESDATA9 = rs("MESDATA9")       ' 測定データその９
        .MESDATA10 = rs("MESDATA10")     ' 測定データその１０
        .MESDATA11 = rs("MESDATA11")     ' 測定データその1１
        .MESDATA12 = rs("MESDATA12")     ' 測定データその1２
        .MESDATA13 = rs("MESDATA13")     ' 測定データその1３
        .MESDATA14 = rs("MESDATA14")     ' 測定データその1４
        .MESDATA15 = rs("MESDATA15")     ' 測定データその1５
        .TXID = rs("TXID")               ' トランザクションID
        .REGDATE = rs("REGDATE")         ' 登録日付
        .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
        .SENDDATE = rs("SENDDATE")       ' 送信日付
    End With
    
    funGetTBCMY013 = 0

PROC_EXIT:
    '終了
    Set rs = Nothing
    gErr.Pop
    Exit Function

PROC_ERR:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume PROC_EXIT
End Function

'------------------------------------------------
' EP先行評価結果取得関数
'------------------------------------------------

'概要      :サンプルＩＤ、評価項目からTBCMY022を検索し、EP先行評価結果を取得し返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                   :説明
'          :sSmplID       ,I  ,String               :サンプルＩＤ
'          :sItemName     ,I  ,String               :エピ評価項目名称(BMD,OSF)
'          :sItemDetail   ,I  ,String               :エピ評価項目詳細名称(BMD1～BMD3,OSF1～OSF3)
'          :tTBCMY022     ,O  ,typ_TBCMY022         :エピ測定評価結果(構造体)
'          :戻り値        ,O  ,Integer              :取得結果 = 0 : 正常
'                                                              -1 : 異常
'説明      :SB_WfHanSui.funGetTBCMY013を基に作成
'履歴      :新規作成 2006/08/15 エピ先行評価追加対応 SMP)kondoh

Public Function funGetTBCMY022(sSmplID As String, sItemName As String, sItemDetail As String, tTBCMY022 As typ_TBCMY022) As Integer

    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    
    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "SB_WfHanSui.bas -- Function funGetTBCMY022"
    
    'サンプルＩＤ等からTBCMY022のEP先行測定評価結果を検索する。
    sql = "select SAMPLEID, OSITEM, MAISU, SPEC, NETSU, ET, MES, DKAN, MESDATA1, MESDATA2, MESDATA3, MESDATA4, MESDATA5, "
    sql = sql & "MESDATA6, MESDATA7, MESDATA8, MESDATA9, MESDATA10, MESDATA11, MESDATA12, MESDATA13, MESDATA14, MESDATA15, "
    sql = sql & "TXID, REGDATE, SENDFLAG, SENDDATE "
    sql = sql & "from TBCMY022 "
    sql = sql & "where SAMPLEID = '" & sSmplID & "' and "
    sql = sql & "      OSITEM = '" & sItemName & "' and "
    sql = sql & "      SPEC = '" & sItemDetail & "'"
    
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetTBCMY022 = -1
        GoTo PROC_EXIT
    End If
    
     ''抽出結果を格納する
    With tTBCMY022
        .SAMPLEID = rs("SAMPLEID")       ' サンプルID
        .OSITEM = rs("OSITEM")           ' 評価項目
        .MAISU = rs("MAISU")             ' 評価枚数
        .Spec = rs("SPEC")               ' 規格値
        .NETSU = rs("NETSU")             ' 熱処理条件
        .ET = rs("ET")                   ' エッチング条件
        .MES = rs("MES")                 ' 計測方法
        .DKAN = rs("DKAN")               ' ＤＫアニール条件
        .MESDATA1 = rs("MESDATA1")       ' 測定データその１
        .MESDATA2 = rs("MESDATA2")       ' 測定データその２
        .MESDATA3 = rs("MESDATA3")       ' 測定データその３
        .MESDATA4 = rs("MESDATA4")       ' 測定データその４
        .MESDATA5 = rs("MESDATA5")       ' 測定データその５
        .MESDATA6 = rs("MESDATA6")       ' 測定データその６
        .MESDATA7 = rs("MESDATA7")       ' 測定データその７
        .MESDATA8 = rs("MESDATA8")       ' 測定データその８
        .MESDATA9 = rs("MESDATA9")       ' 測定データその９
        .MESDATA10 = rs("MESDATA10")     ' 測定データその１０
        .MESDATA11 = rs("MESDATA11")     ' 測定データその1１
        .MESDATA12 = rs("MESDATA12")     ' 測定データその1２
        .MESDATA13 = rs("MESDATA13")     ' 測定データその1３
        .MESDATA14 = rs("MESDATA14")     ' 測定データその1４
        .MESDATA15 = rs("MESDATA15")     ' 測定データその1５
        .TXID = rs("TXID")               ' トランザクションID
        .REGDATE = rs("REGDATE")         ' 登録日付
        .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
        .SENDDATE = rs("SENDDATE")       ' 送信日付
    End With
    
    funGetTBCMY022 = 0

PROC_EXIT:
    '終了
    Set rs = Nothing
    gErr.Pop
    Exit Function
    
PROC_ERR:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume PROC_EXIT
End Function



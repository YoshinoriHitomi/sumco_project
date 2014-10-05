Attribute VB_Name = "SB_CryJudg"
Option Explicit

' ＜判定フロー＞
' 仕様保証方法＿処 --+--なし --実績（該当位置）--あってもなくても判定OK
'　　　　　　　　　　|
'                   +--あり --実績（該当位置) --+--あり -- 判定チェック --+-- OK
'                                              |                        |
'                                              |                        +-- MG
'                                              |
'                                              +--なし --+-- 検査指示５・６以外の場合 --+--EPD、Cs、LTの場合下を探す --+-- なし -- NG
'                                                        |                            |                          　 |
'                                                        |                            +--EPD,Cs、LT以外 -- NG       +-- あり -- 判定チェック --+-- OK
'                                                        |                                                                                    |
'                                                        |                                                                                    +-- NG
'                                                        |
'                                                        +-- 検査指示５の場合 (Rs, Cs) なら推定なので全体から実績を探す --+
'                                                        |  　　　　　　　　　                                          |
'                                                        |                                                             |
'                                                        +-- 検査指示６の場合 TOPなら上へ、TAILなら下へ実績を探す       --+-- 判定チェック --+-- OK
'　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　                 　|
'　　　　　　　　　　　　　　　　　　　　　　　　　　　　　（検査指示は、指示を立てる側が正常に立てていると考えている）　　　　　                 +-- NG

'' デバック定義
''
'' 定数定義
''
Private Const MAXCNT    As Integer = 16         ' 最大件数
Public Const BlkTop     As Integer = 1          ' TOP側
Public Const BlkTail    As Integer = 2          ' TAIL側
Public Const MSYSCLASS  As String = "NM"        ' システム区分
Public Const KCLASS     As String = "01"        ' クラス
Public Const KCODE      As String = "1"         ' コード

'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : 実績パターン区分と製品仕様パターン区分の定義追加
Private Const CNST_JSK_PTN_None As String = "0"
Private Const CNST_JSK_PTN_Ring As String = "1"
Private Const CNST_JSK_PTN_Disk As String = "2"
Private Const CNST_JSK_PTN_DiskRing As String = "3"
Private Const CNST_JSK_PTN_PBband As String = "5"
Private Const CNST_JSK_PTN_Pband As String = "6"
Private Const CNST_JSK_PTN_Bband As String = "7"

Private Const CNST_SIYO_NO_Ring As String = "1"
Private Const CNST_SIYO_NO_Disk As String = "2"
Private Const CNST_SIYO_NO_Pattern As String = "3"
Private Const CNST_SIYO_Fumon As String = "4"
Private Const CNST_SIYO_NO_PBband As String = "5"
Private Const CNST_SIYO_NO_Pband As String = "6"
Private Const CNST_SIYO_NO_Bband As String = "7"
'Add End   2011/01/17 SMPK A.Nagamine

'各判定結果情報
Public Type typ_ALLRSLT
    pos     As Integer                  ' 結晶内開始位置
    NAIYO   As String                   ' 内容
    INFO1   As String                   ' 情報１
    INFO2   As String                   ' 情報２
    INFO3   As String                   ' 情報３
    INFO4   As String                   ' 情報４
    OKNG    As String                   ' 判定結果
    SMPLNO  As Long                     ' サンプルＮｏ  Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLID  As String                   ' サンプルID（WF_Judgで使用）
    BLOCKNG As Boolean                  ' GDエラーとなる品番を含むか判別
    hinban  As String                   ' 品番(12桁)
End Type
    
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
'既存の構造体に項目追加するとVBの制限に引っかかるので、別で管理する。
'各判定結果情報
Public Type typ_ALLRSLT_EX
    pos     As Integer                  ' 結晶内開始位置
    NAIYO   As String                   ' 内容
    INFO1   As String                   ' 情報１
    INFO2   As String                   ' 情報２
    INFO3   As String                   ' 情報３
    INFO4   As String                   ' 情報４
    INFO5   As String                   ' 情報５（AN温度）
'↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
'項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
    INFO6   As String                   ' 情報６（PUA値）
    INFO7   As String                   ' 情報７（PUA%値）
    INFO8   As String                   ' 情報８（STD値）
'↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
    OKNG    As String                   ' 判定結果
    SMPLNO  As Long                     ' サンプルＮｏ  Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLID  As String                   ' サンプルID（WF_Judgで使用）
    BLOCKNG As Boolean                  ' GDエラーとなる品番を含むか判別
    hinban  As String                   ' 品番(12桁)
End Type
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------

'全情報構造体
Type typ_AllTypesB
    intPFlg             As Integer                              ' 表示フラグ
    StrStaffId          As String                               ' スタッフID
    strStaffName        As String                               ' スタッフ名
    BLOCKID             As String * 12                          ' ブロックID
    Cut(2)              As Double                               ' 再カット位置
    COEF(2)             As Double                               ' 偏析係数
    CRCOEF              As Double                               ' 結晶偏析係数
    OKNG(2)             As Boolean                              ' 比抵抗判定
    Henseki             As Boolean                              ' 比抵抗実績有無(結晶全体TOP/TAIL)
    JudgRes(2)          As Boolean                              ' 比抵抗判定    2001/10/02 S.Sano
    JudgRrg(2)          As Boolean                              ' RRG判定       2001/10/02 S.Sano
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    JudgDkTmp(2)        As Boolean                              ' DK温度判定
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    typ_rsz()           As typ_TBCMJ002                         ' 結晶抵抗実績(結晶全体TOP/TAIL)
    typ_hage(2)         As typ_TBCMH004                         ' 引上げ終了実績
    typ_rslt(2, MAXCNT) As typ_ALLRSLT                          ' 各実績情報
    typ_zi              As type_DBDRV_scmzc_fcmkc001c_Zisseki   ' 実績をまとめた構造体
    typ_si()            As type_DBDRV_scmzc_fcmkc001c_Siyou     ' 仕様
    typ_cr()            As type_DBDRV_scmzc_fcmkc001c_CrySmp    ' 結晶サンプル管理取得用 (TOP,TAIL順で２レコード取得)
    blYONE              As Boolean                              ' 米沢フラグ（非購入単結晶　yaz）
    COEFflg             As Boolean                              ' ブロック偏析判定フラグ   2005/1/11追加
    Hinsyu              As String                               ' ブロック偏析判定(品種）  2005/1/11追加
    DOPEflg             As Boolean                              ' 追ﾄﾞｰﾌﾟ位置判定         2005/1/11追加
End Type

Public typ_b        As typ_AllTypesB        '全情報構造体
Public JudgSC_B(2)  As Judg_Spec_Cry        '仕様検査支持構造体
Public ciSmpGetFlg  As Integer              'ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ(0:ｻﾝﾌﾟﾙ指定なし, 1:ｻﾝﾌﾟﾙ指定あり)
Public ciKcnt       As Integer              '工程連番
'----2005/1/11
Type typ_Suitei
    COEF                As Double                               ' 偏析係数
    Henseki             As Boolean                              ' 比抵抗実績有無(結晶全体TOP/TAIL)
    SuiSpec             As type_DBDRV_scmzc_fcmkc001c_Siyou     ' 仕様
    SuiData(2)          As type_DBDRV_scmzc_fcmkc001c_CryR
    COEFflg             As Boolean                              ' ブロック偏析判定フラグ
    Hinsyu              As String                               ' ブロック偏析判定(品種）
    DOPEflg             As Boolean                              ' 追ﾄﾞｰﾌﾟ位置判定
    RsJudg(2)           As Boolean
End Type
'---TEST2004/10
Public SuiteiData() As typ_Suitei
''==複数品番判定対応　20060501 SMP桜井
'' 0:全検査項目で合否判定,1:Cs,LT,EPDで合否判定,2:Skip
Public giTpMultiFlg As Integer ''Topでの合否判定振り分け
Public giBtMultiFlg As Integer ''bottomでの合否判定振り分け
Private pJMEAS_Top() As Double
Private psKSTAFFID  As String
Private psHSXRSPOT  As String
Private psHSXRSPOI  As String
'C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
Public gsCOSF3Flg As String
'C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---

'--------------- 2008/07/25 INSERT START  By Systech ---------------
Private pbGDJudgeTbl(3) As Boolean          ' GD判定結果退避
'--------------- 2008/07/25 INSERT  END   By Systech ---------------

'------------------------------------------------
' 総合判定
'------------------------------------------------

'概要      :実績値の総合判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型             :説明
'          :sKeyID          ,I  ,String         :ﾌﾞﾛｯｸID、又は、結晶番号
'          :tNew_Hinban     ,I  ,tFullHinban    :振替候補品番
'          :bTotalJudg      ,O  ,Boolean        :トータル判定
'          :iErr_Code       ,O  ,Integer        :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String         :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :typ_B           ,O  ,typ_AllTypesB  :全情報構造体(構造体)
'          :iSmpGetFlg      ,I  ,Integer        :ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ(0:ｻﾝﾌﾟﾙ指定なし, 1:ｻﾝﾌﾟﾙ指定あり)
'          :iSamplID1       ,I  ,Long           :TOPｻﾝﾌﾟﾙID(省略可)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :iSamplID2       ,I  ,Long           :BOTｻﾝﾌﾟﾙID(省略可)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :iKcnt           ,I  ,Integer        :工程連番(省略可)
'          :戻り値          ,O  ,Integer        :取得の成否(0:正常終了, -1:異常終了)
'説明      :
'履歴      :2003/09/19 新規作成　SB

Public Function funCrySogoHantei(sKeyID As String, tNew_Hinban As tFullHinban, _
                bTotalJudg As Boolean, iErr_Code As Integer, sErr_Msg As String, typ_b As typ_AllTypesB, _
                iSmpGetFlg As Integer, Optional iSamplID1 As Long = 0, Optional iSamplID2 As Long = 0, _
                Optional iKcnt As Integer = 0) As Integer
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funCrySogoHantei = FUNCTION_RETURN_FAILURE
    TotalJudg = True
    
    'グローバル変数に設定
    ciSmpGetFlg = iSmpGetFlg
    ciKcnt = iKcnt
    
    'ブロックIDを設定
    sErr_Msg = "結晶総合判定(ﾌﾞﾛｯｸID設定)"
    typ_b.BLOCKID = sKeyID
    
    '画面情報設定
    sErr_Msg = "結晶総合判定(SetAllData)"
    If SetAllData(typ_b, tNew_Hinban, iErr_Code, sErr_Msg, iSmpGetFlg, iSamplID1, iSamplID2) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    
    '仕様検査指示取得
    sErr_Msg = "結晶総合判定(SpecJudgCheck)"
    Call SpecJudgCheck
    
    '2003/12/13 SystemBrain Null対応追加▽
    '仕様Nullチェック
    sErr_Msg = "仕様Nullﾁｪｯｸ"
    If funCryChkNull(typ_b.typ_si(BlkTop), sErr_Msg) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    '2003/12/13 SystemBrain Null対応追加△
    
    '実績データ判定(TOP)
    sErr_Msg = "結晶総合判定(判定(TOP))"
    
    '----TEST2004/10
    '画面出力用に実測抵抗値を退避しておく
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS1 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS1
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS2 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS2
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS3 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS3
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS4 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS4
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS5 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS5
        
    If Trim(typ_b.typ_zi.CRYRZ(BlkTop).KSTAFFID) <> KSTAFF_J002 Then
        '抵抗値を測定位置コードにより並べ替える
        
        If Set_Rs_Ichi(typ_b.typ_si(BlkTop).HSXRSPOT, typ_b.typ_si(BlkTop).HSXRSPOI, typ_b.typ_zi.CRYRZ(BlkTop).MEAS1, _
                        typ_b.typ_zi.CRYRZ(BlkTop).MEAS2, typ_b.typ_zi.CRYRZ(BlkTop).MEAS3, typ_b.typ_zi.CRYRZ(BlkTop).MEAS4, typ_b.typ_zi.CRYRZ(BlkTop).MEAS5) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    End If
    If CrAllJudg(typ_b, tNew_Hinban, BlkTop) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    '実績データ判定(TAIL)
    sErr_Msg = "結晶総合判定(判定(TAIL))"
    '画面出力用に実測抵抗値を退避しておく
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS1 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS1
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS2 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS2
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS3 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS3
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS4 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS4
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS5 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS5
        
    If Trim(typ_b.typ_zi.CRYRZ(BlkTail).KSTAFFID) <> KSTAFF_J002 Then
        '抵抗値を測定位置コードにより並べ替える
        If Set_Rs_Ichi(typ_b.typ_si(BlkTail).HSXRSPOT, typ_b.typ_si(BlkTail).HSXRSPOI, typ_b.typ_zi.CRYRZ(BlkTail).MEAS1, _
                        typ_b.typ_zi.CRYRZ(BlkTail).MEAS2, typ_b.typ_zi.CRYRZ(BlkTail).MEAS3, typ_b.typ_zi.CRYRZ(BlkTail).MEAS4, typ_b.typ_zi.CRYRZ(BlkTail).MEAS5) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    End If
    If CrAllJudg(typ_b, tNew_Hinban, BlkTail) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    
    bTotalJudg = TotalJudg
    
    funCrySogoHantei = FUNCTION_RETURN_SUCCESS
'------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funCrySogoHantei = -4
    iErr_Code = funCrySogoHantei
    GoTo Apl_Exit
    
End Function

'------------------------------------------------
' 総合判定(反映データの合否判定を行わない)
'------------------------------------------------

'概要      :実績値の総合判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型             :説明
'          :sKeyID          ,I  ,String         :ﾌﾞﾛｯｸID、又は、結晶番号
'          :Top_Hinban      ,I  ,tFullHinban    :TOP品番
'          :Tail_Hinban     ,I  ,tFullHinban    :TAIL品番
'          :bTotalJudg      ,O  ,Boolean        :トータル判定
'          :iErr_Code       ,O  ,Integer        :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String         :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :typ_B           ,O  ,typ_AllTypesB  :全情報構造体(構造体)
'          :iSmpGetFlg      ,I  ,Integer        :ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ(0:ｻﾝﾌﾟﾙ指定なし, 1:ｻﾝﾌﾟﾙ指定あり)
'          :iSamplID1       ,I  ,Long           :TOPｻﾝﾌﾟﾙID(省略可)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :iSamplID2       ,I  ,Long           :BOTｻﾝﾌﾟﾙID(省略可)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :iKcnt           ,I  ,Integer        :工程連番(省略可)
'          :戻り値          ,O  ,Integer        :取得の成否(0:正常終了, -1:異常終了)
'説明      :
'履歴      :2005/02/07 新規作成　追加 ffc)tanabe

Public Function funCrySogoHantei2(sKeyID As String, Top_Hinban As tFullHinban, Tail_Hinban As tFullHinban, _
                bTotalJudg As Boolean, iErr_Code As Integer, sErr_Msg As String, typ_b As typ_AllTypesB, _
                iSmpGetFlg As Integer, Optional iSamplID1 As Long = 0, Optional iSamplID2 As Long = 0, _
                Optional iKcnt As Integer = 0) As Integer
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funCrySogoHantei2 = FUNCTION_RETURN_FAILURE
    TotalJudg = True
    
    'グローバル変数に設定
    ciSmpGetFlg = iSmpGetFlg
    ciKcnt = iKcnt
    
    'ブロックIDを設定
    sErr_Msg = "結晶総合判定(ﾌﾞﾛｯｸID設定)"
    typ_b.BLOCKID = sKeyID
    
    '画面情報設定(TOP側)
    sErr_Msg = "結晶総合判定(SetAllData2)"
    If SetAllData2(typ_b, Top_Hinban, Tail_Hinban, iErr_Code, sErr_Msg, iSmpGetFlg, iSamplID1, iSamplID2) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    
    '仕様検査指示取得
    sErr_Msg = "結晶総合判定(SpecJudgCheck)"
    Call SpecJudgCheck
    
    '仕様Nullチェック
    sErr_Msg = "仕様Nullﾁｪｯｸ"
    If funCryChkNull(typ_b.typ_si(BlkTop), sErr_Msg) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    
    '実績データ判定(TOP)
    sErr_Msg = "結晶総合判定(判定(TOP))"
    
    '画面出力用に実測抵抗値を退避しておく
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS1 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS1
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS2 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS2
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS3 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS3
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS4 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS4
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS5 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS5
        
    If Trim(typ_b.typ_zi.CRYRZ(BlkTop).KSTAFFID) <> KSTAFF_J002 Then
        '抵抗値を測定位置コードにより並べ替える
        
        If Set_Rs_Ichi(typ_b.typ_si(BlkTop).HSXRSPOT, typ_b.typ_si(BlkTop).HSXRSPOI, typ_b.typ_zi.CRYRZ(BlkTop).MEAS1, _
                        typ_b.typ_zi.CRYRZ(BlkTop).MEAS2, typ_b.typ_zi.CRYRZ(BlkTop).MEAS3, typ_b.typ_zi.CRYRZ(BlkTop).MEAS4, typ_b.typ_zi.CRYRZ(BlkTop).MEAS5) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    End If
    If CrAllJudg(typ_b, Top_Hinban, BlkTop) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    '実績データ判定(TAIL)
    sErr_Msg = "結晶総合判定(判定(TAIL))"
    '画面出力用に実測抵抗値を退避しておく
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS1 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS1
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS2 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS2
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS3 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS3
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS4 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS4
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS5 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS5
        
    If Trim(typ_b.typ_zi.CRYRZ(BlkTail).KSTAFFID) <> KSTAFF_J002 Then
        '抵抗値を測定位置コードにより並べ替える
        If Set_Rs_Ichi(typ_b.typ_si(BlkTail).HSXRSPOT, typ_b.typ_si(BlkTail).HSXRSPOI, typ_b.typ_zi.CRYRZ(BlkTail).MEAS1, _
                        typ_b.typ_zi.CRYRZ(BlkTail).MEAS2, typ_b.typ_zi.CRYRZ(BlkTail).MEAS3, typ_b.typ_zi.CRYRZ(BlkTail).MEAS4, typ_b.typ_zi.CRYRZ(BlkTail).MEAS5) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    End If

    If CrAllJudg(typ_b, Tail_Hinban, BlkTail) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    
    bTotalJudg = TotalJudg
    
    funCrySogoHantei2 = FUNCTION_RETURN_SUCCESS

Apl_Exit:
    
    Exit Function
    
Apl_down:
    funCrySogoHantei2 = -4
    iErr_Code = funCrySogoHantei2
    GoTo Apl_Exit
    
End Function

'概要      :画面情報データ設定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typ_A         ,IO ,typ_AllTypes ,各情報構造体
'説明      :画面情報を情報構造体に設定する
'履歴      :

Public Function SetAllData(typ_b As typ_AllTypesB, tNew_Hinban As tFullHinban, iErr_Code As Integer, _
                           sErr_Msg As String, iSmpGetFlg As Integer, iSamplID1 As Long, iSamplID2 As Long) As FUNCTION_RETURN
    
    Dim typ_hi()    As typ_TBCMH004
    Dim typ_tan     As typ_TBCMG002
    Dim sErrMsg     As String
    Dim i           As Integer

    'ブロックID頭3桁で判断する
    'まだ、保留
    
    SetAllData = FUNCTION_RETURN_FAILURE ''2001/07/25 Sano修正

    '総合判定 各種データ取得
    sErr_Msg = "結晶総合判定(funCryGetDataEtc)"
    If funCryGetDataEtc(typ_b.BLOCKID, tNew_Hinban, _
                        typ_b.typ_si, _
                        typ_b.typ_cr, _
                        typ_b.typ_zi, _
                        sErrMsg, _
                        iSmpGetFlg, iSamplID1, iSamplID2) <> FUNCTION_RETURN_SUCCESS Then
        If sErrMsg = "0" Then sErr_Msg = "生死区分エラー"
        Exit Function
    End If
    
    typ_b.blYONE = True
    With typ_b
        ' 結晶検査指示（Rs)
        sErr_Msg = "結晶総合判定(RS-Top)"
        If InStr("123", .typ_cr(BlkTop).CRYINDRSCS) <> 0 And _
            .typ_zi.CRYRZ(BlkTop).SMPLUMU = "0" Then
            
            '引上げ終了実績取得
            ReDim typ_hi(0)
            sErr_Msg = "結晶総合判定(RS-Top引上げ終了実績取得)"
            If s_cmmc001db_Sql(.typ_cr(BlkTop).XTALCS, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                If UBound(typ_hi) = 0 Then
                   '引上げ終了実績取得失敗
                    Exit Function
                Else
                    .typ_hage(BlkTop) = typ_hi(1)
                End If
            End If
        End If
        
        ' 結晶検査指示（Rs)
        sErr_Msg = "結晶総合判定(RS-Bot)"
        If InStr("123", .typ_cr(BlkTail).CRYINDRSCS) <> 0 And _
            .typ_zi.CRYRZ(BlkTail).SMPLUMU = "0" Then
            
            '引上げ終了実績取得
            ReDim typ_hi(0)
            sErr_Msg = "結晶総合判定(RS-Bot引上げ終了実績取得)"
            If s_cmmc001db_Sql(.typ_cr(BlkTail).XTALCS, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                If UBound(typ_hi) = 0 Then
                   '引上げ終了実績取得失敗
                    Exit Function
                Else
                    .typ_hage(BlkTail) = typ_hi(1)
                End If
            End If
        End If
    End With
    
    '結晶全体TOP/TAIL抵抗実績値取得
    sErr_Msg = "結晶総合判定(結晶全体TOP/TAIL抵抗実績値取得)"
    If s_cmmc001db2_sql(typ_b.typ_si(1).CRYNUM, _
                        typ_b.typ_si(1).ADDDPPOS, _
                        typ_b.typ_si(1).FREELENG, _
                        typ_b.typ_cr(2).INPOSCS, _
                        typ_b.typ_rsz()) <> FUNCTION_RETURN_SUCCESS Then
       '抵抗実績値失敗
        typ_b.Henseki = False
    Else
        typ_b.Henseki = True
    End If
        
    SetAllData = FUNCTION_RETURN_SUCCESS
End Function

'概要      :画面情報データ設定(結晶総合判定：反映データの合否判定を行わない用関数)
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型             ,説明
'          :typ_B           ,IO ,typ_AllTypes   ,各情報構造体
'          :Top_Hinban      ,I  ,tFullHinban    ,TOP品番
'          :Tail_Hinban     ,I  ,tFullHinban    ,TAIL品番
'          :iErr_Code       ,O  ,Integer        ,ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String         ,ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :iSmpGetFlg      ,I  ,Integer        ,ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ(0:ｻﾝﾌﾟﾙ指定なし, 1:ｻﾝﾌﾟﾙ指定あり)
'          :iSamplID1       ,I  ,Long           ,TOPｻﾝﾌﾟﾙID(省略可)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :iSamplID2       ,I  ,Long           ,BOTｻﾝﾌﾟﾙID(省略可)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'説明      :画面情報を情報構造体に設定する
'履歴      :2005/02/08 ffc)tanabe

Public Function SetAllData2(typ_b As typ_AllTypesB, Top_Hinban As tFullHinban, Tail_Hinban As tFullHinban, iErr_Code As Integer, _
                           sErr_Msg As String, iSmpGetFlg As Integer, iSamplID1 As Long, iSamplID2 As Long) As FUNCTION_RETURN
    
    Dim typ_hi()    As typ_TBCMH004
    Dim typ_tan     As typ_TBCMG002
    Dim sErrMsg     As String
    Dim i           As Integer
    
    SetAllData2 = FUNCTION_RETURN_FAILURE

    '総合判定 各種データ取得
    sErr_Msg = "結晶総合判定(funCryGetDataEtc2)"
    If funCryGetDataEtc2(typ_b.BLOCKID, Top_Hinban, Tail_Hinban, _
                        typ_b.typ_si, _
                        typ_b.typ_cr, _
                        typ_b.typ_zi, _
                        sErrMsg, _
                        iSmpGetFlg, iSamplID1, iSamplID2) <> FUNCTION_RETURN_SUCCESS Then
    If sErrMsg = "0" Then sErr_Msg = "生死区分エラー"
        Exit Function
    End If
    
    typ_b.blYONE = True
    With typ_b
        ' 結晶検査指示（Rs)
        sErr_Msg = "結晶総合判定(RS-Top)"
        If InStr("123", .typ_cr(BlkTop).CRYINDRSCS) <> 0 And _
            .typ_zi.CRYRZ(BlkTop).SMPLUMU = "0" Then
            
            '引上げ終了実績取得
            ReDim typ_hi(0)
            sErr_Msg = "結晶総合判定(RS-Top引上げ終了実績取得)"
            If s_cmmc001db_Sql(.typ_cr(BlkTop).XTALCS, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                If UBound(typ_hi) = 0 Then
                   '引上げ終了実績取得失敗
                    Exit Function
                Else
                    .typ_hage(BlkTop) = typ_hi(1)
                End If
            End If
        End If
        
        ' 結晶検査指示（Rs)
        sErr_Msg = "結晶総合判定(RS-Bot)"
        If InStr("123", .typ_cr(BlkTail).CRYINDRSCS) <> 0 And _
            .typ_zi.CRYRZ(BlkTail).SMPLUMU = "0" Then
            
            '引上げ終了実績取得
            ReDim typ_hi(0)
            sErr_Msg = "結晶総合判定(RS-Bot引上げ終了実績取得)"
            If s_cmmc001db_Sql(.typ_cr(BlkTail).XTALCS, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                If UBound(typ_hi) = 0 Then
                   '引上げ終了実績取得失敗
                    Exit Function
                Else
                    .typ_hage(BlkTail) = typ_hi(1)
                End If
            End If
        End If
    End With
    
    '結晶全体TOP/TAIL抵抗実績値取得
    sErr_Msg = "結晶総合判定(結晶全体TOP/TAIL抵抗実績値取得)"
    If s_cmmc001db2_sql(typ_b.typ_si(1).CRYNUM, _
                        typ_b.typ_si(1).ADDDPPOS, _
                        typ_b.typ_si(1).FREELENG, _
                        typ_b.typ_cr(2).INPOSCS, _
                        typ_b.typ_rsz()) <> FUNCTION_RETURN_SUCCESS Then
       '抵抗実績値失敗
        typ_b.Henseki = False
    Else
        typ_b.Henseki = True
    End If
    
    SetAllData2 = FUNCTION_RETURN_SUCCESS

End Function

Public Sub SpecJudgCheck()
    Dim c0              As Integer
    Dim UDHinSpec(2)    As Judg_Spec_Cry
    Dim smpShared       As Boolean
    Dim KouteiKbn       As Integer              '工程区分　08/04/15 ooba
    Dim sSxlPos         As String               'SXL位置(TOP/BOT)　08/04/15 ooba
    
    '08/04/15 ooba START ======================================================>
    '工程により結晶判定有無を判断する。
    '①再抜試指示(CW760)以外の場合は結晶保証により判断。
    '              (結晶保証)=(X) →なし
    '                        =(H) →あり
    '②再抜試指示(CW760)の場合は結晶保証とWF保証の組合せにより判断。
    '       (結晶保証,WF保証)=(X,X) →なし
    '                        =(X,H) →なし
    '                        =(H,X) →あり
    '                        =(H,H) →なし
    '③WF工程(CC720)以降はCOSF3の判定を行なわない。
    
    '工程判断
    Select Case left(JudgKoutei, 4)
    '--結晶工程
    Case "CC10", "CC20", "CC30", "CC31", "CC40", "CC45", "CC46", "CC60", "CC61", "CC70", "CC72"
        KouteiKbn = 0
    '--WF工程(WF判定前)
    Case "CC73", "CW74", "CW75"
        KouteiKbn = 1
    '--WF工程(WF判定後)
    Case "CW76", "CW80"
        KouteiKbn = 2
    '--その他
    Case Else
        KouteiKbn = 0
    End Select
    
    For c0 = 1 To 2
        With typ_b.typ_si(c0)
            sSxlPos = IIf(c0 = SxlTop, "TOP", "BOT")
            '結晶工程
            If KouteiKbn = 0 Then
                JudgSC_B(c0).rs = (.HSXRHWYS = "H")
                JudgSC_B(c0).Oi = (.HSXONHWS = "H")
                JudgSC_B(c0).B1 = (.HSXBM1HS = "H")
                JudgSC_B(c0).B2 = (.HSXBM2HS = "H")
                JudgSC_B(c0).B3 = (.HSXBM3HS = "H")
                JudgSC_B(c0).L1 = (.HSXOF1HS = "H")
                JudgSC_B(c0).L2 = (.HSXOF2HS = "H")
                JudgSC_B(c0).L3 = (.HSXOF3HS = "H")
                JudgSC_B(c0).L4 = (.HSXOF4HS = "H")
                JudgSC_B(c0).COSF3 = (.COSF3FLAG = "H")
                JudgSC_B(c0).GD = (.HSXDENHS = "H") Or (.HSXLDLHS = "H") Or (.HSXDVDHS = "H")
                JudgSC_B(c0).Cs = (.HSXCNHWS = "H")
                JudgSC_B(c0).Lt = (.HSXLTHWS = "H")
                JudgSC_B(c0).EPD = True
                
              'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の項目追加
                JudgSC_B(c0).CuC = (.HSXCHS = "H")
                JudgSC_B(c0).CuCJ = (.HSXCJHS = "H")
                JudgSC_B(c0).CuCJLT = (.HSXCJLTHS = "H")
                JudgSC_B(c0).CuCJ2 = (.HSXCJ2HS = "H")
              'Add End   2011/01/17 SMPK A.Nagamine
            'WF工程(WF判定前)
            ElseIf KouteiKbn = 1 Then
                JudgSC_B(c0).rs = (.HSXRHWYS = "H")
                JudgSC_B(c0).Oi = (.HSXONHWS = "H")
                JudgSC_B(c0).B1 = (.HSXBM1HS = "H")
                JudgSC_B(c0).B2 = (.HSXBM2HS = "H")
                JudgSC_B(c0).B3 = (.HSXBM3HS = "H")
                JudgSC_B(c0).L1 = (.HSXOF1HS = "H")
                JudgSC_B(c0).L2 = (.HSXOF2HS = "H")
                JudgSC_B(c0).L3 = (.HSXOF3HS = "H")
                JudgSC_B(c0).L4 = (.HSXOF4HS = "H")
              'Add Start 2011/02/01 SMPK A.Nagamine     : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の項目追加
                'JudgSC_B(c0).COSF3 = False
                JudgSC_B(c0).COSF3 = (.COSF3FLAG = "H")
              'Add End   2011/02/01 SMPK A.Nagamine
                JudgSC_B(c0).GD = (.HSXDENHS = "H") Or (.HSXLDLHS = "H") Or (.HSXDVDHS = "H")
                JudgSC_B(c0).Cs = (.HSXCNHWS = "H")
                JudgSC_B(c0).Lt = (.HSXLTHWS = "H")
                JudgSC_B(c0).EPD = True
                
              'Add Start 2011/02/01 SMPK A.Nagamine     : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の項目追加
                JudgSC_B(c0).CuC = (.HSXCHS = "H")
                JudgSC_B(c0).CuCJ = (.HSXCJHS = "H")
                JudgSC_B(c0).CuCJLT = (.HSXCJLTHS = "H")
                JudgSC_B(c0).CuCJ2 = (.HSXCJ2HS = "H")
              'Add End   2011/02/01 SMPK A.Nagamine
            'WF工程(WF判定後)
            ElseIf KouteiKbn = 2 Then
                JudgSC_B(c0).rs = (.HSXRHWYS = "H") And _
                                    ((.HWFRHWYS <> "H") Or Not CheckKHN(.HWFRKHNN, 1, sSxlPos))
                JudgSC_B(c0).Oi = (.HSXONHWS = "H") And _
                                    ((.HWFONHWS <> "H") Or Not CheckKHN(.HWFONKHN, 2, sSxlPos))
                JudgSC_B(c0).B1 = (.HSXBM1HS = "H") And _
                                    ((.HWFBM1HS <> "H") Or Not CheckKHN(.HWFBM1KN, 7, sSxlPos))
                JudgSC_B(c0).B2 = (.HSXBM2HS = "H") And _
                                    ((.HWFBM2HS <> "H") Or Not CheckKHN(.HWFBM2KN, 8, sSxlPos))
                JudgSC_B(c0).B3 = (.HSXBM3HS = "H") And _
                                    ((.HWFBM3HS <> "H") Or Not CheckKHN(.HWFBM3KN, 9, sSxlPos))
                JudgSC_B(c0).L1 = (.HSXOF1HS = "H") And _
                                    ((.HWFOF1HS <> "H") Or Not CheckKHN(.HWFOF1KN, 3, sSxlPos))
                JudgSC_B(c0).L2 = (.HSXOF2HS = "H") And _
                                    ((.HWFOF2HS <> "H") Or Not CheckKHN(.HWFOF2KN, 4, sSxlPos))
                JudgSC_B(c0).L3 = (.HSXOF3HS = "H") And _
                                    ((.HWFOF3HS <> "H") Or Not CheckKHN(.HWFOF3KN, 5, sSxlPos))
                JudgSC_B(c0).L4 = (.HSXOF4HS = "H") And _
                                    ((.HWFOF4HS <> "H") Or Not CheckKHN(.HWFOF4KN, 6, sSxlPos))
              'Add Start 2011/02/01 SMPK A.Nagamine     : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の項目追加
                'JudgSC_B(c0).COSF3 = False
                JudgSC_B(c0).COSF3 = (.COSF3FLAG = "H")
              'Add End   2011/02/01 SMPK A.Nagamine
                JudgSC_B(c0).GD = ((.HSXDENHS = "H") Or (.HSXLDLHS = "H") Or (.HSXDVDHS = "H")) And _
                                    (((.HWFDENHS <> "H") And (.HWFLDLHS <> "H") And (.HWFDVDHS <> "H")) Or Not CheckKHN(.HWFGDKHN, 18, sSxlPos))
                JudgSC_B(c0).Cs = (.HSXCNHWS = "H")
                JudgSC_B(c0).Lt = (.HSXLTHWS = "H")
                JudgSC_B(c0).EPD = True
                
              'Add Start 2011/02/01 SMPK A.Nagamine     : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の項目追加
                JudgSC_B(c0).CuC = (.HSXCHS = "H")
                JudgSC_B(c0).CuCJ = (.HSXCJHS = "H")
                JudgSC_B(c0).CuCJLT = (.HSXCJLTHS = "H")
                JudgSC_B(c0).CuCJ2 = (.HSXCJ2HS = "H")
              'Add End   2011/02/01 SMPK A.Nagamine
            End If
        End With
    Next
    '08/04/15 ooba END ========================================================>
    
End Sub

'概要      :引上結晶判定(全)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               :説明
'          :typ_B         ,I  ,typ_AllTypesB    :各情報構造体
'          :tNew_Hinban   ,I  ,tFullHinban      :振替候補品番
'          :tt            ,I  ,Integer          :TopTail判定用
'説明      :検査指示に従い、実績判定を行う
'履歴      :
Public Function CrAllJudg(typ_b As typ_AllTypesB, tNew_Hinban As tFullHinban, tt As Integer) As FUNCTION_RETURN
    Dim IND         As String                   '検査指示
    Dim bJudg       As Boolean
    Dim i           As Integer
    Dim cnt         As Integer
    Dim typTmList() As typ_TBCMB005
    Dim minwk       As String, maxwk As String
    Dim vTemp       As Variant
    Dim RET         As FUNCTION_RETURN
    Dim Gd_si()     As type_DBDRV_scmzc_fcmkc001c_Siyou
    Dim jCs         As String                               'ブロック内品番のCs保証
    Dim jCsFromTo   As String                               'ブロック内品番のCs保証(FromTo)
    Dim hasSiji     As Boolean                              '検査指示あり
    Dim sHinban12   As String                               '品番(12桁)
    Dim bJudgXY     As Boolean                              'X線判定用フラグ追加 2009/10/22
    Dim bJudgX      As Boolean                              'X線判定用フラグ追加 2009/10/22
    Dim bJudgY      As Boolean                              'X線判定用フラグ追加 2009/10/22
    Dim Oi          As C_Oi       '2010/03/12
    
    CrAllJudg = FUNCTION_RETURN_FAILURE
    
    sHinban12 = tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond
    
    i = 0
       
    '結晶コードリスト取得
    If GetCodeList(MSYSCLASS, KCLASS, typTmList) <> FUNCTION_RETURN_SUCCESS Then
        '結晶コードリスト取得失敗
        Exit Function
    End If
    With typ_b
        '' 結晶検査指示(Rs)*****************************************************************
        '検査指示設定
        IND = IIf(tt = BlkTop, "123", "123")
        If JudgSC_B(tt).rs Then
            ' 指示が無い場合は、NGとして表示
            .OKNG(tt) = False
            If (InStr(IND, .typ_cr(tt).CRYINDRSCS) <> 0) Then
                If left(.typ_zi.CRYRZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                    ' サンプルが無い場合は、NGとして表示
                    If .typ_zi.CRYRZ(tt).SMPLUMU = "0" Then
                        '比抵抗判定
                        If Not CrResJudg(1, .typ_si(tt), .typ_zi.CRYRZ(tt), .OKNG(tt), tt) Then
                            '比抵抗判定失敗
                        End If
                    End If
                End If
            End If
            If .OKNG(tt) = False Then
                TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00100"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
        Else
            If (InStr(IND, .typ_cr(tt).CRYINDRSCS) <> 0) Then
                .OKNG(tt) = True
                If left(.typ_zi.CRYRZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                    ' サンプルが無い場合は、OKとして表示
                    If .typ_zi.CRYRZ(tt).SMPLUMU = "0" Then
                        '比抵抗判定
                        If Not CrResJudg(1, .typ_si(tt), .typ_zi.CRYRZ(tt), .OKNG(tt), tt) Then
                            '比抵抗判定失敗
                        End If
                    End If
                End If
            End If
        End If
        
        
        '検査指示設定
        IND = IIf(tt = BlkTop, "123", "123")
        '' 結晶検査指示(Oi)*****************************************************************
        If JudgSC_B(tt).Oi Then
            '画面表示内容設定
            .typ_rslt(tt, i).BLOCKNG = False
            .typ_rslt(tt, i).pos = -1                                       ' 結晶内開始位置
            .typ_rslt(tt, i).NAIYO = Search_CrCode("Oi", typTmList())       ' 内容
            .typ_rslt(tt, i).INFO1 = "仕様有"                               ' 情報１
            .typ_rslt(tt, i).INFO2 = "検査無"                               ' 情報２
            .typ_rslt(tt, i).INFO3 = ""                                     ' 情報３
            .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
            .typ_rslt(tt, i).SMPLNO = -1                                    ' サンプルＮｏ
            .typ_rslt(tt, i).OKNG = "NG"                                    ' 判定結果
            .typ_rslt(tt, i).hinban = sHinban12                             ' 品番(12桁)
            bJudg = False
            If (InStr(IND, .typ_cr(tt).CRYINDOICS) <> 0) Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = .typ_zi.OIZ(tt).POSITION                 ' 結晶内開始位置
                .typ_rslt(tt, i).NAIYO = Search_CrCode("Oi", typTmList())       ' 内容
                .typ_rslt(tt, i).SMPLNO = .typ_zi.OIZ(tt).SMPLNO                ' サンプルＮｏ
                .typ_rslt(tt, i).INFO2 = "検査有"                               ' 情報２
                .typ_rslt(tt, i).INFO3 = "実績無"                               ' 情報２
                If left(.typ_zi.OIZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                    '画面表示内容設定
                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ無"                          ' 情報３
                    If .typ_zi.OIZ(tt).SMPLUMU = "0" Then
                        'OI判定失敗
                        .typ_rslt(tt, i).INFO3 = "判定Err"                      ' 情報２
                        'OI判定
                        If CrOiJudg(.typ_si(tt), .typ_zi.OIZ(tt), bJudg) Then
                            Call GetOiMaxMin(.typ_zi.OIZ(tt), minwk, maxwk)
                            '画面表示内容設定
                            vTemp = CStr(.typ_zi.OIZ(tt).OIMEAS1)                       ' 情報１
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' 情報１
                            .typ_rslt(tt, i).INFO2 = DBData2DispData(maxwk, "0.00")     ' 情報２
                            .typ_rslt(tt, i).INFO3 = DBData2DispData(minwk, "0.00")     ' 情報３
                            vTemp = CStr(.typ_zi.OIZ(tt).ORGRES)                        ' 情報４
                            'ORGの小数桁数を6桁(7桁目四捨五入)に変更 2011/11/25 SETsw kubota
                            '.typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")     ' 情報４
                            .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.000000")     ' 情報４
                        Else
                            If .typ_zi.OIZ(tt).ORGRES = -999 Then               ' 2010/03/12 Kameda
                                ReDim Oi.Oi(4)
                                Oi.Oi(0) = .typ_zi.OIZ(tt).OIMEAS1
                                Oi.Oi(1) = .typ_zi.OIZ(tt).OIMEAS2
                                Oi.Oi(2) = .typ_zi.OIZ(tt).OIMEAS3
                                Oi.Oi(3) = .typ_zi.OIZ(tt).OIMEAS4
                                Oi.Oi(4) = .typ_zi.OIZ(tt).OIMEAS5
                                .typ_rslt(tt, i).INFO1 = "仕様" & .typ_si(tt).HSXONSPT & "点"   ' 情報１
                                .typ_rslt(tt, i).INFO2 = "検査" & GetTensu(Oi) & "点"                                ' 情報２
                                .typ_rslt(tt, i).INFO4 = "点数不足"     ' 情報４
                            End If
                        End If
                    End If
                End If
            End If
            
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                               ' 判定結果
            Else
                .typ_rslt(tt, i).OKNG = "NG"                               ' 判定結果
                TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00101"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
        Else
            If (InStr(IND, .typ_cr(tt).CRYINDOICS) <> 0) Then
                .typ_rslt(tt, i).BLOCKNG = False
                .typ_rslt(tt, i).OKNG = "OK"                                ' 判定結果
                '画面表示内容設定
                .typ_rslt(tt, i).pos = .typ_zi.OIZ(tt).POSITION             ' 結晶内開始位置
                .typ_rslt(tt, i).NAIYO = Search_CrCode("Oi", typTmList())   ' 内容
                .typ_rslt(tt, i).SMPLNO = .typ_zi.OIZ(tt).SMPLNO            ' サンプルＮｏ
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).OKNG = "N参"                                ' 判定結果
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).INFO1 = "仕様無"                           ' 情報１
                .typ_rslt(tt, i).INFO2 = "検査有"                           ' 情報２
                .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ無"                          ' 情報３
                .typ_rslt(tt, i).INFO4 = ""                                 ' 情報４
                .typ_rslt(tt, i).hinban = sHinban12                         ' 品番(12桁)
                If .typ_zi.OIZ(tt).SMPLUMU = "0" Then
                    'OI判定失敗
                    .typ_rslt(tt, i).INFO3 = "判定Err"                          ' 情報２
                    'OI判定
                    If CrOiJudg(.typ_si(tt), .typ_zi.OIZ(tt), bJudg) Then
                        Call GetOiMaxMin(.typ_zi.OIZ(tt), minwk, maxwk)
                        '画面表示内容設定
                        vTemp = CStr(.typ_zi.OIZ(tt).OIMEAS1)                       ' 情報１
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' 情報１
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(maxwk, "0.00")     ' 情報２
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(minwk, "0.00")     ' 情報３
                        vTemp = CStr(.typ_zi.OIZ(tt).ORGRES)                        ' 情報４
                        'ORGの小数桁数を6桁(7桁目四捨五入)に変更 2011/11/25 SETsw kubota
                        '.typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")     ' 情報４
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.000000")     ' 情報４
                    End If
                End If
                i = i + 1
            End If
        End If
        '' 結晶検査指示(B1)*****************************************************************
        BMDDataSet 1, tt, i, typTmList(), sHinban12
        '' 結晶検査指示(B2)*****************************************************************
        BMDDataSet 2, tt, i, typTmList(), sHinban12
        '' 結晶検査指示(B3)*****************************************************************
        BMDDataSet 3, tt, i, typTmList(), sHinban12
        '' 結晶検査指示(L1)*****************************************************************
        OSFDataSet 1, tt, i, typTmList(), sHinban12, .typ_si(tt).HSXOF1ARPTK    '' 引数に, .typ_si(tt).HSXOF1ARPTKを追加 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
        '' 結晶検査指示(L2)*****************************************************************
        OSFDataSet 2, tt, i, typTmList(), sHinban12, " "    '' 引数に, .typ_si(tt).HSXOF1ARPTKを追加 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
        ' 結晶検査指示(L3)*****************************************************************
        OSFDataSet 3, tt, i, typTmList(), sHinban12, " "    '' 引数に, .typ_si(tt).HSXOF1ARPTKを追加 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
        '' 結晶検査指示(L4)*****************************************************************
        OSFDataSet 4, tt, i, typTmList(), sHinban12, " "    '' 引数に, .typ_si(tt).HSXOF1ARPTKを追加 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
        '' 結晶検査指示(Cs)*****************************************************************
        If JudgSC_B(tt).Cs And (tt = BlkTail Or .typ_si(tt).HSXCNKHI = "6" Or .typ_si(tt).HSXCNKHI = "9") Then  'TOP/BOT保証対応 09/01/08 ooba
            '画面表示内容初期化
            .typ_rslt(tt, i).BLOCKNG = False
            .typ_rslt(tt, i).pos = -1                                   ' 結晶内開始位置
            .typ_rslt(tt, i).NAIYO = Search_CrCode("Cs", typTmList())   ' 内容
            .typ_rslt(tt, i).INFO1 = "仕様有"                           ' 情報１
            .typ_rslt(tt, i).INFO2 = "検査無"                           ' 情報２
            .typ_rslt(tt, i).INFO3 = ""                                 ' 情報３
            .typ_rslt(tt, i).INFO4 = ""                                 ' 情報４
            .typ_rslt(tt, i).SMPLNO = -1                                ' サンプルＮｏ
            .typ_rslt(tt, i).OKNG = "NG"                                ' 判定結果
            .typ_rslt(tt, i).hinban = sHinban12                         ' 品番(12桁)
            bJudg = False
            If (InStr(IND, .typ_cr(tt).CRYINDCSCS) <> 0) Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = .typ_zi.CSZ(tt).POSITION                 ' 結晶内開始位置
                .typ_rslt(tt, i).NAIYO = Search_CrCode("Cs", typTmList())       ' 内容
                .typ_rslt(tt, i).SMPLNO = .typ_zi.CSZ(tt).SMPLNO                ' サンプルＮｏ
                .typ_rslt(tt, i).INFO2 = "検査有"                               ' 情報２
                .typ_rslt(tt, i).INFO3 = "実績無"                               ' 情報２
                If left(.typ_zi.CSZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                    '画面表示内容設定
                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ無"                          ' 情報３
                    If .typ_zi.CSZ(tt).SMPLUMU = "0" Then
                        'Cs判定失敗
                        .typ_rslt(tt, i).INFO3 = "判定Err"                      ' 情報２
                        'CS判定取得
                        If CrCsjudg(.typ_si(tt), .typ_zi.CSZ(tt), bJudg) Then
                            '画面表示内容設定
                            vTemp = CStr(.typ_zi.CSZ(tt).CSMEAS)
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00") ' 情報１
                            .typ_rslt(tt, i).INFO2 = ""                             ' 情報２
                            .typ_rslt(tt, i).INFO3 = ""                             ' 情報３
                        End If
                    End If
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                               ' 判定結果
            Else
                .typ_rslt(tt, i).OKNG = "NG"                               ' 判定結果
                TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00111"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
        Else
            If (InStr(IND, .typ_cr(tt).CRYINDCSCS) <> 0) Then
                .typ_rslt(tt, i).BLOCKNG = False
                .typ_rslt(tt, i).OKNG = "OK"                                    ' 判定結果
                '画面表示内容設定
                .typ_rslt(tt, i).pos = .typ_zi.CSZ(tt).POSITION                 ' 結晶内開始位置
                .typ_rslt(tt, i).NAIYO = Search_CrCode("Cs", typTmList())       ' 内容
                .typ_rslt(tt, i).SMPLNO = .typ_zi.CSZ(tt).SMPLNO                ' サンプルＮｏ
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).OKNG = "N参"                                   ' 判定結果
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).INFO1 = "仕様無"                               ' 情報１
                .typ_rslt(tt, i).INFO2 = "検査有"                               ' 情報２
                .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ無"                              ' 情報３
                .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                .typ_rslt(tt, i).hinban = sHinban12                             ' 品番(12桁)
                If .typ_zi.CSZ(tt).SMPLUMU = "0" Then
                    'Cs判定失敗
                    .typ_rslt(tt, i).INFO3 = "判定Err"                      ' 情報２
                    'CS判定取得
                    If CrCsjudg(.typ_si(tt), .typ_zi.CSZ(tt), bJudg) Then
                        '画面表示内容設定
                        vTemp = CStr(.typ_zi.CSZ(tt).CSMEAS)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00") ' 情報１
                        .typ_rslt(tt, i).INFO2 = ""                             ' 情報２
                        .typ_rslt(tt, i).INFO3 = ""                             ' 情報３
                    End If
                End If
                i = i + 1
            End If
        End If
        '' 結晶検査指示(GD)*****************************************************************
        'ブロック内の全品番の仕様を取得
        .typ_rslt(tt, i).BLOCKNG = False
        If JudgSC_B(tt).GD Then
            '画面表示内容設定
            .typ_rslt(tt, i).pos = -1           ' 結晶内開始位置
            .typ_rslt(tt, i).NAIYO = Search_CrCode("GD", typTmList())       ' 内容
            .typ_rslt(tt, i).INFO1 = "仕様有"                               ' 情報１
            .typ_rslt(tt, i).INFO2 = "検査無"                               ' 情報２
            .typ_rslt(tt, i).INFO3 = ""                                     ' 情報３
            .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
            .typ_rslt(tt, i).SMPLNO = -1                                    ' サンプルＮｏ
            .typ_rslt(tt, i).hinban = sHinban12                             ' 品番(12桁)
            bJudg = False
            If (InStr(IND, .typ_cr(tt).CRYINDGDCS) <> 0) Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = .typ_zi.GDZ(tt).POSITION                 ' 結晶内開始位置
                .typ_rslt(tt, i).NAIYO = Search_CrCode("GD", typTmList())       ' 内容
                .typ_rslt(tt, i).SMPLNO = .typ_zi.GDZ(tt).SMPLNO                ' サンプルＮｏ
                .typ_rslt(tt, i).INFO2 = "検査有"                               ' 情報２
                .typ_rslt(tt, i).INFO3 = "実績無"                               ' 情報３
                .typ_rslt(tt, i).INFO4 = "実績無"                               ' 情報４    '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
                If left(.typ_zi.GDZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                    '画面表示内容設定
                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ無"                              ' 情報３
                    If .typ_zi.GDZ(tt).SMPLUMU = "0" Then
                        '画面表示内容設定
                        .typ_rslt(tt, i).INFO3 = "判定Err"                          ' 情報３
                        'GD判定取得
                        If CrGdjudg(.typ_si(tt), .typ_zi.GDZ(tt), bJudg) Then
                            '画面表示内容設定
                            vTemp = CStr(.typ_zi.GDZ(tt).MSRSDEN)                       ' 情報１
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' 情報１
                            vTemp = CStr(.typ_zi.GDZ(tt).MSRSLDL)                       ' 情報２
                            .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0")        ' 情報２
                            vTemp = CStr(.typ_zi.GDZ(tt).MSRSDVD2)                      ' 情報３
                            .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0")        ' 情報３
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
                            vTemp = CStr(.typ_zi.GDZ(tt).MSZEROMN)                      ' 情報４
                            .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0")        ' 情報４
                            .typ_rslt(tt, i).INFO4 = .typ_rslt(tt, i).INFO4 & " , "     ' 情報４
                            vTemp = CStr(.typ_zi.GDZ(tt).MSZEROMX)                      ' 情報４
                            .typ_rslt(tt, i).INFO4 = .typ_rslt(tt, i).INFO4 & _
                                                     DBData2DispData(vTemp, "0")        ' 情報４
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
                        End If
                    End If
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                               ' 判定結果
            Else
                .typ_rslt(tt, i).OKNG = "NG"                               ' 判定結果
                TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                If pbGDJudgeTbl(3) = False Then
                    gsTbcmy028ErrCode = "00114"
                ElseIf pbGDJudgeTbl(3) = False Then
                    gsTbcmy028ErrCode = "00113"
                Else
                    gsTbcmy028ErrCode = "00112"
                End If
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
        Else
            If (InStr(IND, .typ_cr(tt).CRYINDGDCS) <> 0) Then
                '画面表示内容設定
                .typ_rslt(tt, i).BLOCKNG = False
                .typ_rslt(tt, i).pos = .typ_zi.GDZ(tt).POSITION                 ' 結晶内開始位置
                .typ_rslt(tt, i).NAIYO = Search_CrCode("GD", typTmList())       ' 内容
                .typ_rslt(tt, i).INFO1 = "仕様無"                               ' 情報１
                .typ_rslt(tt, i).INFO2 = "検査有"                               ' 情報２
                .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ無"                              ' 情報３
                .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                .typ_rslt(tt, i).INFO4 = "実績無し"                              ' 情報４
                .typ_rslt(tt, i).SMPLNO = .typ_zi.GDZ(tt).SMPLNO                ' サンプルＮｏ
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).OKNG = "N参"                                    ' 判定結果
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).hinban = sHinban12                             ' 品番(12桁)
                If .typ_zi.GDZ(tt).SMPLUMU = "0" Then
                    'GD判定失敗
                    .typ_rslt(tt, i).INFO3 = "判定Err"                          ' 情報３
                    'GD判定取得
                    If CrGdjudg(.typ_si(tt), .typ_zi.GDZ(tt), bJudg) Then
                        '画面表示内容設定
                        vTemp = CStr(.typ_zi.GDZ(tt).MSRSDEN)                       ' 情報１
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' 情報１
                        vTemp = CStr(.typ_zi.GDZ(tt).MSRSLDL)                       ' 情報２
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0")        ' 情報２
                        vTemp = CStr(.typ_zi.GDZ(tt).MSRSDVD2)                      ' 情報３
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0")        ' 情報３
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
                        vTemp = CStr(.typ_zi.GDZ(tt).MSZEROMN)                      ' 情報４
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0")        ' 情報４
                        .typ_rslt(tt, i).INFO4 = .typ_rslt(tt, i).INFO4 & " , "     ' 情報４
                        vTemp = CStr(.typ_zi.GDZ(tt).MSZEROMX)                      ' 情報４
                        .typ_rslt(tt, i).INFO4 = .typ_rslt(tt, i).INFO4 & _
                                                 DBData2DispData(vTemp, "0")        ' 情報４
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
                    End If
                End If
                i = i + 1
            End If
        End If
        '' 結晶検査指示(T)*****************************************************************
Dim HIN As tFullHinban
Dim LTSPI As String

        If (InStr(IND, .typ_cr(tt).CRYINDTCS) <> 0) Then
            hasSiji = True
        Else
            hasSiji = False
        End If
        bJudg = True                                        '2004/01/15 SystemBrain
        If (JudgSC_B(tt).Lt) And (tt = BlkTail) Then        '2004/01/15 SystemBrain
            bJudg = False                                   '2004/01/15 SystemBrain
        Else                                                '2004/01/15 SystemBrain
            JudgSC_B(tt).Lt = False                         '2004/01/15 SystemBrain
        End If                                              '2004/01/15 SystemBrain
        
        'LTはBot端でブロック全域を判定することになったため、「Top端品番でLT指示があればBotで表示」は不要となった
        If (JudgSC_B(tt).Lt) Or (hasSiji And (tt = BlkTail)) Then '仕様あり or Bot端で検査あり
            .typ_rslt(tt, i).BLOCKNG = False
            
            '画面表示内容初期化
            .typ_rslt(tt, i).pos = .typ_zi.LTZ(tt).POSITION             ' 結晶内開始位置
            .typ_rslt(tt, i).SMPLNO = -1                                ' サンプルＮｏ
            .typ_rslt(tt, i).NAIYO = Search_CrCode("T", typTmList())    ' 内容
            If JudgSC_B(tt).Lt Then
                .typ_rslt(tt, i).INFO1 = "仕様有"                       ' 情報１
            Else
                .typ_rslt(tt, i).INFO1 = "仕様無"
                bJudg = True
            End If
            If hasSiji Then
                .typ_rslt(tt, i).INFO2 = "検査有"                       ' 情報２
            Else
                .typ_rslt(tt, i).INFO2 = "検査無"
            End If
            .typ_rslt(tt, i).INFO3 = "実績無"                           ' 情報３
            .typ_rslt(tt, i).INFO4 = ""                                 ' 情報４
            .typ_rslt(tt, i).hinban = sHinban12                         ' 品番(12桁)
            
            'ライフタイム
            bJudgX = True   '10Ω判定
            '判定と結果登録
            If .typ_zi.LTZ(tt).CRYNUM = .typ_si(1).CRYNUM Then
                .typ_rslt(tt, i).pos = .typ_zi.LTZ(tt).POSITION                 ' 結晶内開始位置
                .typ_rslt(tt, i).SMPLNO = .typ_zi.LTZ(tt).SMPLNO                ' サンプルＮｏ
                If (.typ_zi.LTZ(tt).SMPLUMU <> "0") Then
                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ無"                          ' 情報３
                Else
                    '2005/12/02 add SET高崎 LT計算関数call ->
                    'ライフタイム値を計算しなおす
                    Call Sub_LTReCalc(.typ_si(tt), .typ_zi.LTZ(tt))
                    '2005/12/02 add SET高崎 LT計算関数call <-

                    'LT判定取得
                    If CrLtjudg(.typ_si(tt), .typ_zi.LTZ(tt), bJudg) Then
''Add Start 2011/07/22 LT10Ω判定追加対応 T.Koi(SETsw)
                        If CrLt10judg(.typ_si(tt), .typ_zi.LTZ(tt), .typ_cr(tt), bJudgX) Then
''Add End   2011/07/22 LT10Ω判定追加対応 T.Koi(SETsw)
                            '画面表示内容設定
                            vTemp = CStr(.typ_zi.LTZ(tt).CALCMEAS)                  ' 情報１
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")
                            vTemp = CStr(.typ_zi.LTZ(tt).MEASPEAK)                  ' 情報２
                            .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0")
                            .typ_rslt(tt, i).INFO3 = ""                             ' 情報３
''Add Start 2011/07/22 LT10Ω判定追加対応 T.Koi(SETsw)
                            ' 情報４
                            If .typ_zi.LTZ(tt).CONVAL = (-1) Then
                                .typ_rslt(tt, i).INFO4 = "NULL"
                            Else
                                .typ_rslt(tt, i).INFO4 = CStr(.typ_zi.LTZ(tt).CONVAL)
                            End If
                        Else
                            .typ_rslt(tt, i).INFO3 = "LT10判定Err"                  ' 情報３
                        End If
''Add End   2011/07/22 LT10Ω判定追加対応 T.Koi(SETsw)
                    Else
                        .typ_rslt(tt, i).INFO3 = "判定Err"                      ' 情報３
                    End If
                End If
            Else    '実績なし
                If JudgSC_B(tt).Lt Then bJudg = False
            End If
            
''Add Start 2011/07/25 LT10Ω判定追加対応 T.Koi(SETsw)
            If bJudg = True Then
                If bJudgX = True Then
                    bJudg = True
                Else
                    bJudg = False
                End If
            End If
''Add End   2011/07/25 LT10Ω判定追加対応 T.Koi(SETsw)
            
            If (bJudg = False) Then
                .typ_rslt(tt, i).OKNG = "NG"
                TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00110"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
'====================== Debug Debug =====================================
            ElseIf .typ_si(tt).HSXLTHWS = "S" Then
                .typ_rslt(tt, i).OKNG = "N参"                            ' 判定結果
'====================== Debug Debug =====================================
            Else
                .typ_rslt(tt, i).OKNG = "OK"                            ' 判定結果
            End If
            i = i + 1
        End If
        '' 結晶検査指示(EPD)*****************************************************************
        If JudgSC_B(tt).EPD Then
            If tt = BlkTop Then
                .typ_rslt(tt, i).BLOCKNG = False
                If (InStr(IND, .typ_cr(tt).CRYINDEPCS) <> 0) Then
                    '画面表示内容設定
                    .typ_rslt(tt, i).pos = .typ_zi.EPDZ(tt).POSITION                ' 結晶内開始位置
                    .typ_rslt(tt, i).SMPLNO = .typ_zi.EPDZ(tt).SMPLNO               ' サンプルＮｏ
                    .typ_rslt(tt, i).NAIYO = Search_CrCode("EPD", typTmList())      ' 内容
                    .typ_rslt(tt, i).INFO1 = "仕様有"                               ' 情報１
                    .typ_rslt(tt, i).INFO2 = "検査有"                               ' 情報２
                    .typ_rslt(tt, i).INFO3 = "実績無"                               ' 情報３
                    .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                    .typ_rslt(tt, i).hinban = sHinban12                             ' 品番(12桁)
                    bJudg = False
                    If left(.typ_zi.EPDZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ無"                          ' 情報３
                        If .typ_zi.EPDZ(tt).SMPLUMU = "0" Then
                            'EPD判定失敗
                            .typ_rslt(tt, i).INFO3 = "判定Err"                      ' 情報３
                            'EPD判定取得
                            If CrEpdjudg(.typ_si(tt), .typ_zi.EPDZ(tt), bJudg) Then
                                '画面表示内容設定
                                vTemp = CStr(.typ_zi.EPDZ(tt).MEASURE)                  ' 情報１
                                .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")    ' 情報１
                                .typ_rslt(tt, i).INFO2 = ""                             ' 情報２
                                .typ_rslt(tt, i).INFO3 = ""                             ' 情報３
                            End If
                        End If
                    End If
                    If bJudg = True Then
                        .typ_rslt(tt, i).OKNG = "OK"                               ' 判定結果
                    Else
                        .typ_rslt(tt, i).OKNG = "NG"                               ' 判定結果
                        TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                        gsTbcmy028ErrCode = "00102"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
                    End If
                    i = i + 1
                End If
            Else
                '画面表示内容設定

'>>>>>  サンプル無し対応 2006/05/09変更 kubota
                .typ_rslt(tt, i).pos = .typ_zi.EPDZ(tt).POSITION    ' 結晶内開始位置
'<<<<<  サンプル無し対応 2006/05/09変更 kubota
                
                .typ_rslt(tt, i).NAIYO = Search_CrCode("EPD", typTmList())          ' 内容
                .typ_rslt(tt, i).INFO1 = "仕様有"                                   ' 情報１
                
'>>>>>  サンプル無し対応 2006/05/09変更 kubota
                .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
'<<<<<  サンプル無し対応 2006/05/09変更 kubota
                
                .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
                .typ_rslt(tt, i).INFO4 = ""                                         ' 情報４
                
'>>>>>  サンプル無し対応 2006/05/09変更 kubota
                .typ_rslt(tt, i).SMPLNO = .typ_zi.EPDZ(tt).SMPLNO   ' サンプルＮｏ
'<<<<<  サンプル無し対応 2006/05/09変更 kubota
                
                .typ_rslt(tt, i).hinban = sHinban12                                 ' 品番(12桁)
                bJudg = False
                If (InStr(IND, .typ_cr(tt).CRYINDEPCS) <> 0) Then
                    If left(.typ_zi.EPDZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ無"                              ' 情報３
                        If .typ_zi.EPDZ(tt).SMPLUMU = "0" Then
                            .typ_rslt(tt, i).pos = .typ_zi.EPDZ(tt).POSITION            ' 結晶内開始位置
                            .typ_rslt(tt, i).SMPLNO = .typ_zi.EPDZ(tt).SMPLNO           ' サンプルＮｏ
                            'EPD判定失敗
                            .typ_rslt(tt, i).INFO3 = "判定Err"                          ' 情報３
                            'EPD判定取得
                            If CrEpdjudg(.typ_si(tt), .typ_zi.EPDZ(tt), bJudg) Then
                                '画面表示内容設定
                                vTemp = CStr(.typ_zi.EPDZ(tt).MEASURE)                  ' 情報１
                                .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")    ' 情報１
                                .typ_rslt(tt, i).INFO2 = ""                             ' 情報２
                                .typ_rslt(tt, i).INFO3 = ""                             ' 情報３
                            End If
                        End If
                    End If
                Else
                    If left(.typ_zi.EPDZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(tt, i).pos = .typ_zi.EPDZ(tt).POSITION                ' 結晶内開始位置
                        .typ_rslt(tt, i).SMPLNO = .typ_zi.EPDZ(tt).SMPLNO               ' サンプルＮｏ
                        'EPD判定失敗
                        .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報３
                        'EPD判定取得
                        If CrEpdjudg(.typ_si(tt), .typ_zi.EPDZ(tt), bJudg) Then
                            '画面表示内容設定
                            vTemp = CStr(.typ_zi.EPDZ(tt).MEASURE)          ' 情報１
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' 情報１
                            .typ_rslt(tt, i).INFO2 = ""                                 ' 情報２
                            .typ_rslt(tt, i).INFO3 = ""                                 ' 情報３
                        End If
                    End If
                End If
                If bJudg = True Then
                    .typ_rslt(tt, i).OKNG = "OK"                               ' 判定結果
                Else
                    .typ_rslt(tt, i).OKNG = "NG"                               ' 判定結果
                    TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                    gsTbcmy028ErrCode = "00111"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
                End If
                i = i + 1
            End If
        Else
            If (InStr(IND, .typ_cr(tt).CRYINDEPCS) <> 0) Then
                '画面表示内容設定
                .typ_rslt(tt, i).BLOCKNG = False
                .typ_rslt(tt, i).pos = .typ_zi.EPDZ(tt).POSITION                    ' 結晶内開始位置
                .typ_rslt(tt, i).NAIYO = Search_CrCode("EPD", typTmList())          ' 内容
                .typ_rslt(tt, i).INFO1 = "仕様無"                                   ' 情報１
                .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
                .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ無"                                  ' 情報３
                .typ_rslt(tt, i).INFO4 = ""                                         ' 情報４
                .typ_rslt(tt, i).SMPLNO = .typ_zi.EPDZ(tt).SMPLNO                   ' サンプルＮｏ
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).OKNG = "N参"                                        ' 判定結果
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).hinban = sHinban12                                 ' 品番(12桁)
                If .typ_zi.EPDZ(tt).SMPLUMU = "0" Then
                    'EPD判定失敗
                    .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報３
                    'EPD判定取得
                    If CrEpdjudg(.typ_si(tt), .typ_zi.EPDZ(tt), bJudg) Then
                        '画面表示内容設定
                        vTemp = CStr(.typ_zi.EPDZ(tt).MEASURE)                      ' 情報１
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' 情報１
                        .typ_rslt(tt, i).INFO2 = ""                                 ' 情報２
                        .typ_rslt(tt, i).INFO3 = ""                                 ' 情報３
                    End If
                End If
                i = i + 1
            End If
        End If
        'SIRD判定データ設定   2010/02/04 add Kameda
        If tt = BlkTop Then
            If .typ_cr(tt).SIRDKBNY3 = "1" Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = .typ_zi.SIRD.POSITION                ' 結晶内開始位置
                '.typ_rslt(tt, i).SMPLNO = .typ_zi.SIRD.SMPLNO               ' サンプルＮｏ
                .typ_rslt(tt, i).NAIYO = Search_CrCode("SD", typTmList())       ' 内容
                .typ_rslt(tt, i).hinban = sHinban12                             ' 品番(12桁)
                bJudg = False
                'SIRD判定取得
                If CrSIRDjudg(.typ_si(tt), .typ_zi.SIRD, bJudg) Then
                    '画面表示内容設定
                    vTemp = CStr(.typ_zi.SIRD.SIRDCNT)                  ' 情報１
                    .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")    ' 情報１
                    .typ_rslt(tt, i).INFO2 = ""                             ' 情報２
                    .typ_rslt(tt, i).INFO3 = ""                             ' 情報３
                End If
                If bJudg = True Then
                    .typ_rslt(tt, i).OKNG = "OK"                               ' 判定結果
                Else
                    .typ_rslt(tt, i).OKNG = "NG"                               ' 判定結果
                    TotalJudg = False
                    ''gsTbcmy028ErrCode = ""
                End If
                '評価待ち参照時  2010/02/18 Kameda
                If .typ_zi.SIRD.NothingFlg = "1" Then
                    .typ_rslt(tt, i).INFO1 = ""                                ' 情報１
                    .typ_rslt(tt, i).OKNG = "評価待ち"                         ' 判定結果
                    'Add Start 2012/01/31 Y.Hitomi
                    TotalJudg = False
                    'Add End 2012/01/31 Y.Hitomi
                End If
                i = i + 1
            ElseIf .typ_cr(tt).SIRDKBNY3 = "2" Then       '2010/02/16 add Kameda
                '画面表示内容設定
                .typ_rslt(tt, i).pos = .typ_zi.SIRD.POSITION                 ' 結晶内開始位置
                '.typ_rslt(tt, i).SMPLNO = .typ_zi.SIRD.SMPLNO               ' サンプルＮｏ
                .typ_rslt(tt, i).NAIYO = Search_CrCode("SD", typTmList())       ' 内容
                .typ_rslt(tt, i).hinban = sHinban12                             ' 品番(12桁)
                'bJudg = False    表示のみ
                'SIRD表示
                '画面表示内容設定
                .typ_rslt(tt, i).INFO1 = "先行評価"                     ' 情報１
                .typ_rslt(tt, i).INFO2 = ""                             ' 情報２
                .typ_rslt(tt, i).INFO3 = ""                             ' 情報３
                .typ_rslt(tt, i).OKNG = "OK"                            ' 判定結果
                i = i + 1
            End If
        End If
        
        'X線判定データ設定   2009/08/12 add Kameda
        '合成角のみで判定 X,Yは警告を出す(背景赤）  2009/10/22 add Kameda
        If tt = BlkTail Then
            If .typ_cr(tt).CRYINDXC1 <> 0 Then
                'If CrXjudg(.typ_si(tt), .typ_zi.XZ, bJudg) Then     2009/10/22 Kameda
                If CrXjudg(.typ_si(tt), .typ_zi.XZ, bJudgXY, bJudgX, bJudgY) Then
                    If bJudgXY Then
                        '.typ_zi.XZ.JUDG = "OK"    2009/10/22
                        .typ_zi.XZ.JUDGXY = "OK"
                    Else
                        '.typ_zi.XZ.JUDG = "NG"    2009/10/22
                        .typ_zi.XZ.JUDGXY = "NG"
                        TotalJudg = False
                    End If
                    '警告を出すために項目追加     2009/10/22 Kameda
                    If bJudgX Then
                        .typ_zi.XZ.JUDGX = "OK"
                    Else
                        .typ_zi.XZ.JUDGX = "NG"
                    End If
                    If bJudgY Then
                        .typ_zi.XZ.JUDGY = "OK"
                    Else
                        .typ_zi.XZ.JUDGY = "NG"
                    End If
                End If
            Else
                '.typ_zi.XZ.JUDG = ""     2009/10/22
                .typ_zi.XZ.JUDGXY = ""
                .typ_zi.XZ.JUDGX = ""
                .typ_zi.XZ.JUDGY = ""
            End If
        End If
        
      'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の実績判定処理
        Call CuDecoDataSet_C(JudgSC_B, typ_b, tt, i, typTmList(), sHinban12)
        
        Call CuDecoDataSet_CJ(JudgSC_B, typ_b, tt, i, typTmList(), sHinban12)
        
        Call CuDecoDataSet_CJLT(JudgSC_B, typ_b, tt, i, typTmList(), sHinban12)
        
        Call CuDecoDataSet_CJ2(JudgSC_B, typ_b, tt, i, typTmList(), sHinban12)
        
      ''Add End   2011/01/17 SMPK A.Nagamine
        
    End With
    
    CrAllJudg = FUNCTION_RETURN_SUCCESS
End Function

'概要      :コード情報取得
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :strCode       ,I  ,String       ,検索コード
'          :CodeData      ,I  ,typ_TBCMB005 ,コードリスト構造体
'          :戻り値        ,O  ,String       ,該当コード文字列
'説明      :コード情報リストから該当コードの情報を取得する
'履歴      :
Private Function Search_CrCode(strCode As String, typ_CodeData() As typ_TBCMB005) As String
    Dim i As Integer
    
    'リストから該当コードの情報１を検索
    i = 1
    Do While typ_CodeData(i).INFO1 <> ""
        If strCode = Trim(typ_CodeData(i).CODE) Then
            Search_CrCode = typ_CodeData(i).INFO1
            Exit Function
        End If
        i = i + 1
    Loop
    Search_CrCode = ""
End Function

'概要      :OI実績測定値MIN/MAX値取得
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typ_oiz       ,I  ,type_DBDRV_scmzc_fcmkc001c_Oi ,各情報構造体
'          :min           ,O  ,String       ,MIN値
'          :max           ,O  ,String       ,MAX値
'説明      :OI実績測定値からMIN・MAX値を取得する
'履歴      :
Private Sub GetOiMaxMin(typ_oiz As type_DBDRV_scmzc_fcmkc001c_Oi, _
                            OiMin As String, OiMax As String)
    Dim wk(4) As Double

    With typ_oiz
        wk(0) = .OIMEAS1                ' Ｏｉ測定値１
        wk(1) = .OIMEAS2                ' Ｏｉ測定値２
        wk(2) = .OIMEAS3                ' Ｏｉ測定値３
        wk(3) = .OIMEAS4                ' Ｏｉ測定値４
        wk(4) = .OIMEAS5                ' Ｏｉ測定値５
    End With
    OiMin = JudgMin(wk())
    OiMax = JudgMax(wk())
End Sub

'概要      :抵抗判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :iCompFlg      ,I   ,Integer                             :偏析計算ﾌﾗｸﾞ(0:偏析計算なし, 1:偏析計算あり)
'          :typ_si        ,I   ,type_DBDRV_scmzc_fcmkc001c_Siyou    :仕様情報構造体
'          :typ_cryrz     ,I   ,type_DBDRV_scmzc_fcmkc001c_CryR     :RS実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(True:判定OK, False:判定NG)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :抵抗判定を行う
'履歴      :
Public Function CrResJudg(iCompFlg As Integer, _
                          typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_cryrz As type_DBDRV_scmzc_fcmkc001c_CryR, _
                          bJudg As Boolean, _
                          tt As Integer) As Boolean
    Dim ErrInfo     As ERROR_INFOMATION     'エラー情報構造体
    Dim rs          As C_RES                'RS判定構造体
    Dim cc          As type_Coefficient
    Dim rp          As type_ResPosCal
    Dim COEF        As Double
    Dim wgtCharge   As Long                 '偏析計算用パラメータ
    Dim wgtTop      As Double               '偏析計算用パラメータ
    Dim wgtTopCut   As Double               '偏析計算用パラメータ
    Dim DM          As Double               '偏析計算用パラメータ
    Dim test As String
    Dim cf As C_COEF
    Dim sMcno2 As String
    Dim sMcno1 As String
    
    bJudg = True
    
    '抵抗判定引数設定
    rs.GuaranteeRes.cMeth = typ_si.HSXRSPOH     '測定位置_方
    rs.GuaranteeRes.cCount = typ_si.HSXRSPOT    '測定位置_点
    rs.GuaranteeRes.cPos = typ_si.HSXRSPOI      '測定位置_位(OSFの場合 領)
    rs.GuaranteeRes.cObj = typ_si.HSXRHWYT      '保証方法_対
    rs.GuaranteeRes.cJudg = typ_si.HSXRHWYS     '保証方法_処
    rs.GuaranteeRes.cBunp = typ_si.HSXRMCAL     '分布計算
    rs.SpecResMin = typ_si.HSXRMIN              ' 品ＳＸ比抵抗下限
    rs.SpecResMax = typ_si.HSXRMAX              ' 品ＳＸ比抵抗上限
    rs.SpecResAveMin = typ_si.HSXRAMIN          ' 品ＳＸ比抵抗平均下限
    rs.SpecResAveMax = typ_si.HSXRAMAX          ' 品ＳＸ比抵抗平均上限
    rs.SpecRrg = typ_si.HSXRMBNP                ' 品ＳＸ比抵抗面内分布
    rs.Res(0) = typ_cryrz.MEAS1                 ' 測定値１
    rs.Res(1) = typ_cryrz.MEAS2                 ' 測定値２
    rs.Res(2) = typ_cryrz.MEAS3                 ' 測定値３
    rs.Res(3) = typ_cryrz.MEAS4                 ' 測定値４
    rs.Res(4) = typ_cryrz.MEAS5                 ' 測定値５
    rs.RRG = typ_cryrz.RRG                      ' ＲＲＧ
'--------------- 2008/08/25 INSERT START  By Systech --------------
    rs.DkTmpSiyo = typ_si.HSXDKTMP
    rs.DkTmpJsk = typ_cryrz.HSXDKTMP
'--------------- 2008/08/25 INSERT  END   By Systech --------------
    '抵抗判定
    If CrystalRESJudg(rs, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrResJudg = False
        typ_cryrz.RRG = rs.RRG '再計算結果をとりあえず表示しない
        If iCompFlg = 1 Then
            typ_b.JudgRes(tt) = rs.JudgRes1 '2001/10/02 S.Sano
            typ_b.JudgRrg(tt) = rs.JudgRrg '2001/10/02 S.Sano
'--------------- 2008/08/25 INSERT START  By Systech --------------
            typ_b.JudgDkTmp(tt) = rs.JudgDkTmp
'--------------- 2008/08/25 INSERT  END   By Systech --------------
        End If
        Exit Function
    End If
    
    typ_cryrz.RRG = rs.RRG '2001/10/02 S.Sano 再計算結果をとりあえず表示しない
    If (iCompFlg = 1) And (ciSmpGetFlg = 0) Then
        typ_b.JudgRes(tt) = rs.JudgRes1 '2001/10/02 S.Sano
        typ_b.JudgRrg(tt) = rs.JudgRrg '2001/10/02 S.Sano
'--------------- 2008/08/25 INSERT START  By Systech --------------
        typ_b.JudgDkTmp(tt) = rs.JudgDkTmp
'--------------- 2008/08/25 INSERT  END   By Systech --------------
    
        '偏析係数計算 マルチ引上対応 参照関数変更 2008/04/23 SETsw Nakada
        If GetCoeffParams_new(typ_cryrz.CRYNUM, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then
            Debug.Print "偏析計算用パラメータの取得に失敗した"
        End If
        With typ_b
            cc.DUNMENSEKI = AreaOfCircle(DM)
            cc.TOPSMPLPOS = .typ_zi.CRYRZ(1).POSITION
            cc.BOTSMPLPOS = .typ_zi.CRYRZ(2).POSITION
            cc.CHARGEWEIGHT = wgtCharge
            cc.TOPWEIGHT = wgtTop + wgtTopCut
            cc.TOPRES = .typ_zi.CRYRZ(1).MEAS1
            cc.BOTRES = .typ_zi.CRYRZ(2).MEAS1
            .COEF(tt) = CoefficientCalculation(cc)
            If .Henseki = True Then
                '結晶偏析係数計算
                cc.DUNMENSEKI = AreaOfCircle(DM)
                cc.TOPSMPLPOS = .typ_rsz(1).POSITION
                cc.BOTSMPLPOS = .typ_rsz(2).POSITION
                cc.CHARGEWEIGHT = wgtCharge
                cc.TOPWEIGHT = wgtTop + wgtTopCut
                cc.TOPRES = .typ_rsz(1).MEAS1
                cc.BOTRES = .typ_rsz(2).MEAS1
                .CRCOEF = CoefficientCalculation(cc)
            End If
            '2005/01/11 ブロック偏析判定処理追加 -------
            sMcno1 = Mid(Trim(.typ_si(tt).PRODCOND), 2, 1)
            sMcno2 = Mid(Trim(.typ_si(tt).PRODCOND), 1, 1)
            cf.JudgCOEF = True
            Select Case sMcno1
                Case "H", "I", "J", "K"
                    cf.NP = "n"
                Case "A", "B", "C"
                    Select Case sMcno2
                        Case "A", "B"
                            cf.NP = "p+"
                        Case "1", "2", "3", "4", "5", "6", "7", "C", "E"
                            cf.NP = "p-"
                        Case Else
                            cf.JudgCOEF = False
                    End Select
                Case Else
                    cf.JudgCOEF = False
            End Select
            If cf.JudgCOEF Then
                cf.COEF = .COEF(tt)
                If CrystalCOEFJudg(cf, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
                    cf.JudgCOEF = False
                End If
            End If
            'エラー表示用にフラグをセットする
            If cf.JudgCOEF Then
                .COEFflg = True
            Else
                .COEFflg = False
            End If
            .Hinsyu = cf.NP
            '追加ﾄﾞｰﾌﾟ位置のチェック
            .DOPEflg = True
            If .typ_si(tt).ADDDPPOS <> 0 Then
                If .typ_si(tt).INGOTPOS <= .typ_si(tt).ADDDPPOS And _
                   .typ_si(tt).INGOTPOS + .typ_si(tt).Length >= .typ_si(tt).ADDDPPOS Then
                   .DOPEflg = False
                End If
            End If
            '2005/01/11 --------------------------------
        
        End With
    End If
    
    If Not rs.JudgRes Then '2001/10/02 S.Sano
        If (iCompFlg = 1) And (ciSmpGetFlg = 0) Then
            With typ_b
                '偏析計算から再カット位置を計算
                rp.COEFFICIENT = .COEF(tt)
                rp.DUNMENSEKI = AreaOfCircle(DM)
                rp.CHARGEWEIGHT = wgtCharge
                rp.TOPWEIGHT = wgtTop + wgtTopCut
                rp.TOPSMPLPOS = .typ_zi.CRYRZ(1).POSITION
                rp.TOPRES = .typ_zi.CRYRZ(1).MEAS1
                rp.target = IIf(tt = BlkTop, .typ_si(tt).HSXRMAX, .typ_si(tt).HSXRMIN)
                .Cut(tt) = PosCalculation(rp)
            End With
        End If
        bJudg = False
    End If
    CrResJudg = True
    
End Function

'概要      :OI判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :仕様情報構造体
'          :typ_oiz       ,I  ,type_DBDRV_scmzc_fcmkc001c_Oi        :OI実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(True:判定OK, False:判定NG)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :OI判定を行う
'履歴      :
Public Function CrOiJudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                         typ_oiz As type_DBDRV_scmzc_fcmkc001c_Oi, _
                         bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim Oi      As C_Oi                     'Oi判定構造体
    
    ReDim Oi.Oi(4) As Double
    
    bJudg = True
        
    'OI判定引数設定
    Oi.GuaranteeOi.cMeth = typ_si.HSXONSPH      '測定位置_方
    Oi.GuaranteeOi.cCount = typ_si.HSXONSPT     '測定位置_点
    Oi.GuaranteeOi.cPos = typ_si.HSXONSPI       '測定位置_位(OSFの場合 領)
    Oi.GuaranteeOi.cObj = typ_si.HSXONHWT       '保証方法_対
    Oi.GuaranteeOi.cJudg = typ_si.HSXONHWS      '保証方法_処
    Oi.GuaranteeOi.cBunp = typ_si.HSXONMCL      '分布計算
    Oi.SpecOiMin = typ_si.HSXONMIN              '品SX酸素濃度下限
    Oi.SpecOiMax = typ_si.HSXONMAX              '品SX酸素濃度上限
    Oi.SpecORG = typ_si.HSXONMBP                '品SX酸素濃度面内分布
    Oi.SpecOiAveMin = typ_si.HSXONAMN           '品SX酸素濃度平均下限
    Oi.SpecOiAveMax = typ_si.HSXONAMX           '品SX酸素濃度平均上限
    
    Oi.Oi(0) = typ_oiz.OIMEAS1             'Oi測定値
    Oi.Oi(1) = typ_oiz.OIMEAS2             'Oi測定値
    Oi.Oi(2) = typ_oiz.OIMEAS3             'Oi測定値
    Oi.Oi(3) = typ_oiz.OIMEAS4             'Oi測定値
    Oi.Oi(4) = typ_oiz.OIMEAS5             'Oi測定値
    Oi.ORG = typ_oiz.ORGRES                'ORG計算値
    '2010/05/10 参考仕様対応 Y.Hitomi
    If Oi.GuaranteeOi.cCount >= "1" Then
        '測定点数のチェック   2010/03/12 Kameda
        If Oi.Oi(CInt(Oi.GuaranteeOi.cCount) - 1) = -1 Then
            typ_oiz.ORGRES = -999   '測定点数不足
            CrOiJudg = False
            bJudg = False
            Exit Function
        End If
    End If
    
    'OI判定
    If CrystalOiJudg(Oi, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        CrOiJudg = False
        bJudg = False
        Exit Function
    End If
    
    'ORGの再計算の値を表示する
    typ_oiz.ORGRES = Oi.ORG                   'ORG計算値
    
    If Oi.JudgOi <> True Or Oi.JudgOrg <> True Then
        bJudg = False
    End If
    
    CrOiJudg = True
End Function

'概要      :BMD判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :仕様情報構造体
'          :typ_bmdz      ,I  ,type_DBDRV_scmzc_fcmkc001c_BMD       :BMD実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(True:判定OK, False:判定NG)
'          :bmflg         ,I  ,Integer                              :BMDﾌﾗｸﾞ(1:BMD1, 2:BMD2, 3:BMD3)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :BMD判定を行う
'履歴      :
Public Function CrBmdJudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_bmdz As type_DBDRV_scmzc_fcmkc001c_BMD, _
                          bJudg As Boolean, _
                          bmflg As Integer) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim bm      As C_BMD                    'BMD構造体
    Dim w_Bunpu As Double

    bJudg = True

    'BMD判定引数設定
    Select Case bmflg
    Case 1
        bm.GuaranteeBmd.cMeth = typ_si.HSXBM1SH   '測定位置_方
        bm.GuaranteeBmd.cCount = typ_si.HSXBM1ST  '測定位置_点
        bm.GuaranteeBmd.cPos = typ_si.HSXBM1SR    '測定位置_位(OSFの場合 領)
        bm.GuaranteeBmd.cObj = typ_si.HSXBM1HT    '保証方法_対
        bm.GuaranteeBmd.cJudg = typ_si.HSXBM1HS   '保証方法_処
        bm.SpecBmdAveMin = typ_si.HSXBM1AN        '品SXBMD平均下限
        bm.SpecBmdAveMax = typ_si.HSXBM1AX        '品SXBMD平均上限
    Case 2
        bm.GuaranteeBmd.cMeth = typ_si.HSXBM2SH   '測定位置_方
        bm.GuaranteeBmd.cCount = typ_si.HSXBM2ST  '測定位置_点
        bm.GuaranteeBmd.cPos = typ_si.HSXBM2SR    '測定位置_位(OSFの場合 領)
        bm.GuaranteeBmd.cObj = typ_si.HSXBM2HT    '保証方法_対
        bm.GuaranteeBmd.cJudg = typ_si.HSXBM2HS   '保証方法_処
        bm.SpecBmdAveMin = typ_si.HSXBM2AN        '品SXBMD平均下限
        bm.SpecBmdAveMax = typ_si.HSXBM2AX        '品SXBMD平均上限
    Case 3
        bm.GuaranteeBmd.cMeth = typ_si.HSXBM3SH   '測定位置_方
        bm.GuaranteeBmd.cCount = typ_si.HSXBM3ST  '測定位置_点
        bm.GuaranteeBmd.cPos = typ_si.HSXBM3SR    '測定位置_位(OSFの場合 領)
        bm.GuaranteeBmd.cObj = typ_si.HSXBM3HT    '保証方法_対
        bm.GuaranteeBmd.cJudg = typ_si.HSXBM3HS   '保証方法_処
        bm.SpecBmdAveMin = typ_si.HSXBM3AN        '品SXBMD平均下限
        bm.SpecBmdAveMax = typ_si.HSXBM3AX        '品SXBMD平均上限
    End Select
    
    bm.BMD(0) = typ_bmdz.MEAS1                      'BMD測定値
    bm.BMD(1) = typ_bmdz.MEAS2                      'BMD測定値
    bm.BMD(2) = typ_bmdz.MEAS3                      'BMD測定値
    bm.BMD(3) = typ_bmdz.MEAS4                      'BMD測定値
    bm.BMD(4) = typ_bmdz.MEAS5                      'BMD測定値
    bm.Min = typ_bmdz.MEASMIN                       '最小値
    bm.max = typ_bmdz.MEASMAX                       '最大値
    bm.AVE = typ_bmdz.MEASAVE                       '平均値
    
    w_Bunpu = typ_bmdz.BMDMNBUNP

    If typ_si.HSXBM1HS = "H" And typ_si.HSXBM1HT <> "" Then
       If bmflg = "1" Then
          If typ_si.HSXBMD1MBP < w_Bunpu And typ_si.HSXBMD1MBP <> -1 Then
             bJudg = False
          End If
       ElseIf bmflg = "2" Then
          If typ_si.HSXBMD2MBP < w_Bunpu And typ_si.HSXBMD2MBP <> -1 Then
             bJudg = False
          End If
       ElseIf bmflg = "3" Then
          If typ_si.HSXBMD3MBP < w_Bunpu And typ_si.HSXBMD3MBP <> -1 Then
             bJudg = False
          End If
       End If
    End If
    
    'BMD判定
    If CrystalBMDJudg(bm, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrBmdJudg = False
        Exit Function
    End If
    If bm.JudgBmd <> True Then
        bJudg = False
    End If
    
    CrBmdJudg = True

End Function

'概要      :OSF判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :仕様情報構造体
'          :typ_osfz      ,I  ,type_DBDRV_scmzc_fcmkc001c_OSF       :OSF実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(True:判定OK, False:判定NG)
'          :osfflg        ,I  ,Integer                              :OSFﾌﾗｸﾞ(1:OSF1, 2:OSF2, 3:OSF3, 4:OSF4)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :OSF判定を行う
'履歴      :
Public Function CrOsfJudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_Osfz As type_DBDRV_scmzc_fcmkc001c_OSF, _
                          bJudg As Boolean, _
                          osfflg As Integer) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim os      As C_OSF                    'OSF構造体
    Dim w_RD    As String
    Dim j       As Integer

    bJudg = True

    'OSF判定引数設定
    Select Case osfflg
    Case 1
        os.GuaranteeOsf.cMeth = typ_si.HSXOF1SH   '測定位置_方
        os.GuaranteeOsf.cCount = typ_si.HSXOF1ST  '測定位置_点
        os.GuaranteeOsf.cPos = typ_si.HSXOF1SR    '測定位置_位(OSFの場合 領)
        os.GuaranteeOsf.cObj = typ_si.HSXOF1HT    '保証方法_対
        os.GuaranteeOsf.cJudg = typ_si.HSXOF1HS   '保証方法_処
        os.SpecOsfAveMax = typ_si.HSXOF1AX        '品SXOSF平均上限
        os.SpecOsfMax = typ_si.HSXOF1MX           '品SX上限
    Case 2
        os.GuaranteeOsf.cMeth = typ_si.HSXOF2SH   '測定位置_方
        os.GuaranteeOsf.cCount = typ_si.HSXOF2ST  '測定位置_点
        os.GuaranteeOsf.cPos = typ_si.HSXOF2SR    '測定位置_位(OSFの場合 領)
        os.GuaranteeOsf.cObj = typ_si.HSXOF2HT    '保証方法_対
        os.GuaranteeOsf.cJudg = typ_si.HSXOF2HS   '保証方法_処
        os.SpecOsfAveMax = typ_si.HSXOF2AX        '品SXOSF平均上限
        os.SpecOsfMax = typ_si.HSXOF2MX           '品SX上限
    Case 3
        os.GuaranteeOsf.cMeth = typ_si.HSXOF3SH   '測定位置_方
        os.GuaranteeOsf.cCount = typ_si.HSXOF3ST  '測定位置_点
        os.GuaranteeOsf.cPos = typ_si.HSXOF3SR    '測定位置_位(OSFの場合 領)
        os.GuaranteeOsf.cObj = typ_si.HSXOF3HT    '保証方法_対
        os.GuaranteeOsf.cJudg = typ_si.HSXOF3HS   '保証方法_処
        os.SpecOsfAveMax = typ_si.HSXOF3AX        '品SXOSF平均上限
        os.SpecOsfMax = typ_si.HSXOF3MX           '品SX上限
    Case 4
        os.GuaranteeOsf.cMeth = typ_si.HSXOF4SH   '測定位置_方
        os.GuaranteeOsf.cCount = typ_si.HSXOF4ST  '測定位置_点
        os.GuaranteeOsf.cPos = typ_si.HSXOF4SR    '測定位置_位(OSFの場合 領)
        os.GuaranteeOsf.cObj = typ_si.HSXOF4HT    '保証方法_対
        os.GuaranteeOsf.cJudg = typ_si.HSXOF4HS   '保証方法_処
        os.SpecOsfAveMax = typ_si.HSXOF4AX        '品SXOSF平均上限
        os.SpecOsfMax = typ_si.HSXOF4MX           '品SX上限
    End Select

    os.OSF(0) = typ_Osfz.MEAS1        'OSF測定値
    os.OSF(1) = typ_Osfz.MEAS2        'OSF測定値
    os.OSF(2) = typ_Osfz.MEAS3        'OSF測定値
    os.OSF(3) = typ_Osfz.MEAS4        'OSF測定値
    os.OSF(4) = typ_Osfz.MEAS5        'OSF測定値
    os.OSF(5) = typ_Osfz.MEAS6        'OSF測定値
    os.OSF(6) = typ_Osfz.MEAS7        'OSF測定値
    os.OSF(7) = typ_Osfz.MEAS8        'OSF測定値
    os.OSF(8) = typ_Osfz.MEAS9        'OSF測定値
    os.OSF(9) = typ_Osfz.MEAS10       'OSF測定値
    os.OSF(10) = typ_Osfz.MEAS11      'OSF測定値
    os.OSF(11) = typ_Osfz.MEAS12      'OSF測定値
    os.OSF(12) = typ_Osfz.MEAS13      'OSF測定値
    os.OSF(13) = typ_Osfz.MEAS14      'OSF測定値
    os.OSF(14) = typ_Osfz.MEAS15      'OSF測定値
    os.OSF(15) = typ_Osfz.MEAS16      'OSF測定値
    os.OSF(16) = typ_Osfz.MEAS17      'OSF測定値
    os.OSF(17) = typ_Osfz.MEAS18      'OSF測定値
    os.OSF(18) = typ_Osfz.MEAS19      'OSF測定値
    os.OSF(19) = typ_Osfz.MEAS20      'OSF測定値
    os.max = typ_Osfz.CALCMAX         '最大値
    os.AVE = typ_Osfz.CALCAVE         '平均値
    
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    os.ARPTK = typ_si.HSXOF1ARPTK       '品SXOSF1(ArAN)パタン区分
    os.ARMIN = typ_si.HSXOFARMIN        '品SXOSF(ArAN)下限
    os.ARMAX = typ_si.HSXOFARMAX        '品SXOSF(ArAN)上限
    os.ARMHMX = typ_si.HSXOFARMHMX      '品SXOSF(ArAN)面内比上限
    os.CALCMH = typ_Osfz.CALCMH         '面内比(MAX/MIN)
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End

    w_RD = typ_Osfz.OSFRD1 + typ_Osfz.OSFRD2 + typ_Osfz.OSFRD3

    If os.GuaranteeOsf.cJudg = "H" And os.GuaranteeOsf.cObj <> "" Then
       If osfflg = 1 Then
           Select Case typ_si.HSXOSF1PTK
               Case "1"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "R" Then bJudg = False
                   Next
               Case "2"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "D" Then bJudg = False
                   Next
               Case "3"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "R" Or Mid(w_RD, j, 1) = "D" Then bJudg = False
                   Next
            End Select
       ElseIf osfflg = 2 Then
           Select Case typ_si.HSXOSF2PTK
               Case "1"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "R" Then bJudg = False
                   Next
               Case "2"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "D" Then bJudg = False
                   Next
               Case "3"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "R" Or Mid(w_RD, j, 1) = "D" Then bJudg = False
                   Next
            End Select
       ElseIf osfflg = 3 Then
           Select Case typ_si.HSXOSF3PTK
               Case "1"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "R" Then bJudg = False
                   Next
               Case "2"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "D" Then bJudg = False
                   Next
               Case "3"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "R" Or Mid(w_RD, j, 1) = "D" Then bJudg = False
                   Next
            End Select
       ElseIf osfflg = 4 Then
           Select Case typ_si.HSXOSF4PTK
               Case "1"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "R" Then bJudg = False
                   Next
               Case "2"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "D" Then bJudg = False
                   Next
               Case "3"
                   For j = 1 To 3
                       If Mid(w_RD, j, 1) = "R" Or Mid(w_RD, j, 1) = "D" Then bJudg = False
                   Next
            End Select
       End If
    End If
    
    'OSF判定
    If CrystalOSFJudg(os, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrOsfJudg = False
        Exit Function
    End If
    
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    If osfflg = 1 Then
        'OSF判定
        If CrystalOSFJudg_02(os, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
            bJudg = False
            CrOsfJudg = False
            Exit Function
        End If
        
        os.JudgOsf = os.JudgOsf And os.JudgOsfPtn
        
        If os.ARPTK = "1" Or os.ARPTK = "2" Then
            If os.JudgOsfPtn = True Then
                typ_Osfz.PTNJUDGRES = "1"
            Else
                typ_Osfz.PTNJUDGRES = "9"
            End If
        Else
            typ_Osfz.PTNJUDGRES = " "
        End If
    End If
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    
    If os.JudgOsf <> True Then
        bJudg = False
    End If
    
    CrOsfJudg = True

End Function

'概要      :CS判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :仕様情報構造体
'          :typ_csz       ,I  ,type_DBDRV_scmzc_fcmkc001c_CS        :CS実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(True:判定OK, False:判定NG)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :CS判定を行う
'履歴      :
Public Function CrCsjudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                         typ_csz As type_DBDRV_scmzc_fcmkc001c_CS, _
                         bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim Cs      As C_Cs                     'CS構造体
    
    bJudg = True
        
    'CS判定引数設定
    Cs.GuaranteeCs.cMeth = typ_si.HSXCNSPH   '測定位置_方
    Cs.GuaranteeCs.cCount = typ_si.HSXCNSPT  '測定位置_点
    Cs.GuaranteeCs.cPos = typ_si.HSXCNSPI    '測定位置_位(OSFの場合 領)
    Cs.GuaranteeCs.cObj = typ_si.HSXCNHWT    '保証方法_対
    Cs.GuaranteeCs.cJudg = typ_si.HSXCNHWS   '保証方法_処
    Cs.SpecCsMin = typ_si.HSXCNMIN           '品SX炭素濃度下限
    Cs.SpecCsMax = typ_si.HSXCNMAX           '品SX炭素濃度上限
    Cs.SpecCsKHI = typ_si.HSXCNKHI           '検査頻度_位 09/01/08 ooba
    Cs.Cs = typ_csz.CSMEAS                   'Cs測定値
    
    'CS判定
    If CrystalCsJudg(Cs, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrCsjudg = False
        Exit Function
    End If
    
    If Cs.JudgCs <> True Then
        bJudg = False
    End If
    
    CrCsjudg = True

End Function

'概要      :GD判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :仕様情報構造体
'          :typ_gdz       ,I  ,type_DBDRV_scmzc_fcmkc001c_GD        :GD実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(True:判定OK, False:判定NG)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :GD判定を行う
'履歴      :
Public Function CrGdjudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                         typ_gdz As type_DBDRV_scmzc_fcmkc001c_GD, _
                         bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim GD      As C_GD                     'GD構造体
    
    bJudg = True
        
    'GD判定引数設定
    GD.GuaranteeDen.cMeth = ""                  '測定位置_方
    GD.GuaranteeDen.cCount = ""                 '測定位置_点
    GD.GuaranteeDen.cPos = ""                   '測定位置_位(OSFの場合 領)
    GD.GuaranteeDen.cObj = typ_si.HSXDENHT      '保証方法_対
    GD.GuaranteeDen.cJudg = typ_si.HSXDENHS     '保証方法_処
    
    GD.GuaranteeLdl.cMeth = ""                  '測定位置_方
    GD.GuaranteeLdl.cCount = ""                 '測定位置_点
    GD.GuaranteeLdl.cPos = ""                   '測定位置_位(OSFの場合 領)
    GD.GuaranteeLdl.cObj = typ_si.HSXLDLHT      '保証方法_対
    GD.GuaranteeLdl.cJudg = typ_si.HSXLDLHS     '保証方法_処
    
    GD.GuaranteeDvd2.cMeth = ""                 '測定位置_方
    GD.GuaranteeDvd2.cCount = ""                '測定位置_点
    GD.GuaranteeDvd2.cPos = ""                  '測定位置_位(OSFの場合 領)
    GD.GuaranteeDvd2.cObj = typ_si.HSXDVDHT     '保証方法_対
    GD.GuaranteeDvd2.cJudg = typ_si.HSXDVDHS    '保証方法_処
    
    GD.JudgFlagDen = typ_si.HSXDENKU            '品SXDen検査有無
    GD.JudgFlagLdl = typ_si.HSXLDLKU            '品SXL/DL検査有無
    GD.JudgFlagDvd2 = typ_si.HSXDVDKU           '品SXDVD2検査有無
    
    GD.SpecDenMin = typ_si.HSXDENMN             '品SXDen下限
    GD.SpecDenMax = typ_si.HSXDENMX             '品SXDen上限
    GD.SpecLdlMin = typ_si.HSXLDLMN             '品SXLdl下限
    GD.SpecLdlMax = typ_si.HSXLDLMX             '品SXLdl上限
    GD.SpecDvd2Min = typ_si.HSXDVDMN            '品SXDvd2下限
    GD.SpecDvd2Max = typ_si.HSXDVDMX            '品SXDvd2上限
'*** UPDATE ↓ Y.SIMIZU 2005/10/13 品SXGDﾗｲﾝ数追加
    GD.SpecGdLine = typ_si.HSXGDLINE            '品SXGDﾗｲﾝ数
'*** UPDATE ↑ Y.SIMIZU 2005/10/13 品SXGDﾗｲﾝ数追加
    
    GD.Den = typ_gdz.MSRSDEN                    'Den計算値
    GD.Ldl = typ_gdz.MSRSLDL                    'L/DL計算値
    GD.Dvd2 = typ_gdz.MSRSDVD2                  'Dvd2計算値
    
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    GD.ZeroLdlMin = typ_si.HSXLDLRMN            '品SXL/DL連続0下限
    GD.ZeroLdlMax = typ_si.HSXLDLRMX            '品SXL/DL連続0上限
    GD.LdlMin = typ_gdz.MSZEROMN                'L/DL0連続数最小値
    GD.LdlMax = typ_gdz.MSZEROMX                'L/DL0連続数最大値
    GD.GDPTK = typ_si.HSXGDPTK                  '品ＳＸＧＤパタン区分
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    
'*** UPDATE ↓ Y.SIMIZU 2005/10/13 ﾗｲﾝ数対応
    'GDﾗｲﾝ数が3又は4.5又は5でない場合は判定ｴﾗｰ
    If GD.SpecGdLine <> 3 And GD.SpecGdLine <> 4.5 And GD.SpecGdLine <> 5 Then
        bJudg = False
        CrGdjudg = False
        Exit Function
    End If
    
    'GDﾗｲﾝ数分の実績があるかをﾁｪｯｸする
    If ChkGD_Data(typ_gdz, GD) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrGdjudg = False
        Exit Function
    End If
'*** UPDATE ↑ Y.SIMIZU 2005/10/13 ﾗｲﾝ数対応

    'GD判定
    If CrystalGDJudg(GD, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrGdjudg = False
        Exit Function
    End If
    
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    GD.JudgLdl = GD.JudgLdl And GD.JudgLdlPtn

    If GD.GDPTK = "1" Or GD.GDPTK = "2" Then
        If GD.JudgLdlPtn = True Then
            typ_gdz.PTNJUDGRES = "1"
        Else
            typ_gdz.PTNJUDGRES = "9"
        End If
    Else
        typ_gdz.PTNJUDGRES = " "
    End If
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    
    If GD.JudgDen <> True Or GD.JudgLdl <> True Or GD.JudgDvd2 <> True Then
        bJudg = False
    End If
    
'--------------- 2008/07/25 INSERT START  By Systech ---------------
    pbGDJudgeTbl(1) = GD.JudgDen
    pbGDJudgeTbl(2) = GD.JudgDvd2
    pbGDJudgeTbl(3) = GD.JudgLdl
'--------------- 2008/07/25 INSERT  END   By Systech ---------------

    CrGdjudg = True

End Function

'概要      :仕様のGDﾗｲﾝ数分測定値が存在するかをﾁｪｯｸする
'ﾊﾟﾗﾒｰﾀ    :変数名      ,IO ,型                             :説明
'          :tGDdata    ,I   ,type_DBDRV_scmzc_fcmkc001c_GD  :GD実績構造体
'          :GD         ,O   ,C_GD                           :GD仕様構造体
'          :戻り値      ,O  ,FUNCTION_RETURN                :結果 = FUNCTION_RETURN_SUCCESS : OK
'                                                           FUNCTION_RETURN_FAILURE : NG
'説明      :
'履歴      :05/10/13 Y.SIMIZU
Private Function ChkGD_Data(tGDdata As type_DBDRV_scmzc_fcmkc001c_GD, GD As C_GD) As FUNCTION_RETURN
    Dim iCnt            As Integer
    Dim iPoint          As Integer
    Dim iLine           As Integer
    Dim iTden(5, 15)    As Integer
    Dim iTldl(5, 15)    As Integer
    Dim iTdvd2(5)       As Integer

    'GD測定値ｾｯﾄ
    With tGDdata
        iTden(1, 1) = .MS01DEN1         '測定値01 Den1
        iTden(2, 1) = .MS01DEN2         '測定値01 Den2
        iTden(3, 1) = .MS01DEN3         '測定値01 Den3
        iTden(4, 1) = .MS01DEN4         '測定値01 Den4
        iTden(5, 1) = .MS01DEN5         '測定値01 Den5
        iTden(1, 2) = .MS02DEN1         '測定値02 Den1
        iTden(2, 2) = .MS02DEN2         '測定値02 Den2
        iTden(3, 2) = .MS02DEN3         '測定値02 Den3
        iTden(4, 2) = .MS02DEN4         '測定値02 Den4
        iTden(5, 2) = .MS02DEN5         '測定値02 Den5
        iTden(1, 3) = .MS03DEN1         '測定値03 Den1
        iTden(2, 3) = .MS03DEN2         '測定値03 Den2
        iTden(3, 3) = .MS03DEN3         '測定値03 Den3
        iTden(4, 3) = .MS03DEN4         '測定値03 Den4
        iTden(5, 3) = .MS03DEN5         '測定値03 Den5
        iTden(1, 4) = .MS04DEN1         '測定値04 Den1
        iTden(2, 4) = .MS04DEN2         '測定値04 Den2
        iTden(3, 4) = .MS04DEN3         '測定値04 Den3
        iTden(4, 4) = .MS04DEN4         '測定値04 Den4
        iTden(5, 4) = .MS04DEN5         '測定値04 Den5
        iTden(1, 5) = .MS05DEN1         '測定値05 Den1
        iTden(2, 5) = .MS05DEN2         '測定値05 Den2
        iTden(3, 5) = .MS05DEN3         '測定値05 Den3
        iTden(4, 5) = .MS05DEN4         '測定値05 Den4
        iTden(5, 5) = .MS05DEN5         '測定値05 Den5
        iTden(1, 6) = .MS06DEN1         '測定値06 Den1
        iTden(2, 6) = .MS06DEN2         '測定値06 Den2
        iTden(3, 6) = .MS06DEN3         '測定値06 Den3
        iTden(4, 6) = .MS06DEN4         '測定値06 Den4
        iTden(5, 6) = .MS06DEN5         '測定値06 Den5
        iTden(1, 7) = .MS07DEN1         '測定値07 Den1
        iTden(2, 7) = .MS07DEN2         '測定値07 Den2
        iTden(3, 7) = .MS07DEN3         '測定値07 Den3
        iTden(4, 7) = .MS07DEN4         '測定値07 Den4
        iTden(5, 7) = .MS07DEN5         '測定値07 Den5
        iTden(1, 8) = .MS08DEN1         '測定値08 Den1
        iTden(2, 8) = .MS08DEN2         '測定値08 Den2
        iTden(3, 8) = .MS08DEN3         '測定値08 Den3
        iTden(4, 8) = .MS08DEN4         '測定値08 Den4
        iTden(5, 8) = .MS08DEN5         '測定値08 Den5
        iTden(1, 9) = .MS09DEN1         '測定値09 Den1
        iTden(2, 9) = .MS09DEN2         '測定値09 Den2
        iTden(3, 9) = .MS09DEN3         '測定値09 Den3
        iTden(4, 9) = .MS09DEN4         '測定値09 Den4
        iTden(5, 9) = .MS09DEN5         '測定値09 Den5
        iTden(1, 10) = .MS10DEN1        '測定値10 Den1
        iTden(2, 10) = .MS10DEN2        '測定値10 Den2
        iTden(3, 10) = .MS10DEN3        '測定値10 Den3
        iTden(4, 10) = .MS10DEN4        '測定値10 Den4
        iTden(5, 10) = .MS10DEN5        '測定値10 Den5
        iTden(1, 11) = .MS11DEN1        '測定値11 Den1
        iTden(2, 11) = .MS11DEN2        '測定値11 Den2
        iTden(3, 11) = .MS11DEN3        '測定値11 Den3
        iTden(4, 11) = .MS11DEN4        '測定値11 Den4
        iTden(5, 11) = .MS11DEN5        '測定値11 Den5
        iTden(1, 12) = .MS12DEN1        '測定値12 Den1
        iTden(2, 12) = .MS12DEN2        '測定値12 Den2
        iTden(3, 12) = .MS12DEN3        '測定値12 Den3
        iTden(4, 12) = .MS12DEN4        '測定値12 Den4
        iTden(5, 12) = .MS12DEN5        '測定値12 Den5
        iTden(1, 13) = .MS13DEN1        '測定値13 Den1
        iTden(2, 13) = .MS13DEN2        '測定値13 Den2
        iTden(3, 13) = .MS13DEN3        '測定値13 Den3
        iTden(4, 13) = .MS13DEN4        '測定値13 Den4
        iTden(5, 13) = .MS13DEN5        '測定値13 Den5
        iTden(1, 14) = .MS14DEN1        '測定値14 Den1
        iTden(2, 14) = .MS14DEN2        '測定値14 Den2
        iTden(3, 14) = .MS14DEN3        '測定値14 Den3
        iTden(4, 14) = .MS14DEN4        '測定値14 Den4
        iTden(5, 14) = .MS14DEN5        '測定値14 Den5
        iTden(1, 15) = .MS15DEN1        '測定値15 Den1
        iTden(2, 15) = .MS15DEN2        '測定値15 Den2
        iTden(3, 15) = .MS15DEN3        '測定値15 Den3
        iTden(4, 15) = .MS15DEN4        '測定値15 Den4
        iTden(5, 15) = .MS15DEN5        '測定値15 Den5
        
        iTldl(1, 1) = .MS01LDL1         '測定値01 L/DL1
        iTldl(2, 1) = .MS01LDL2         '測定値01 L/DL2
        iTldl(3, 1) = .MS01LDL3         '測定値01 L/DL3
        iTldl(4, 1) = .MS01LDL4         '測定値01 L/DL4
        iTldl(5, 1) = .MS01LDL5         '測定値01 L/DL5
        iTldl(1, 2) = .MS02LDL1         '測定値02 L/DL1
        iTldl(2, 2) = .MS02LDL2         '測定値02 L/DL2
        iTldl(3, 2) = .MS02LDL3         '測定値02 L/DL3
        iTldl(4, 2) = .MS02LDL4         '測定値02 L/DL4
        iTldl(5, 2) = .MS02LDL5         '測定値02 L/DL5
        iTldl(1, 3) = .MS03LDL1         '測定値03 L/DL1
        iTldl(2, 3) = .MS03LDL2         '測定値03 L/DL2
        iTldl(3, 3) = .MS03LDL3         '測定値03 L/DL3
        iTldl(4, 3) = .MS03LDL4         '測定値03 L/DL4
        iTldl(5, 3) = .MS03LDL5         '測定値03 L/DL5
        iTldl(1, 4) = .MS04LDL1         '測定値04 L/DL1
        iTldl(2, 4) = .MS04LDL2         '測定値04 L/DL2
        iTldl(3, 4) = .MS04LDL3         '測定値04 L/DL3
        iTldl(4, 4) = .MS04LDL4         '測定値04 L/DL4
        iTldl(5, 4) = .MS04LDL5         '測定値04 L/DL5
        iTldl(1, 5) = .MS05LDL1         '測定値05 L/DL1
        iTldl(2, 5) = .MS05LDL2         '測定値05 L/DL2
        iTldl(3, 5) = .MS05LDL3         '測定値05 L/DL3
        iTldl(4, 5) = .MS05LDL4         '測定値05 L/DL4
        iTldl(5, 5) = .MS05LDL5         '測定値05 L/DL5
        iTldl(1, 6) = .MS06LDL1         '測定値06 L/DL1
        iTldl(2, 6) = .MS06LDL2         '測定値06 L/DL2
        iTldl(3, 6) = .MS06LDL3         '測定値06 L/DL3
        iTldl(4, 6) = .MS06LDL4         '測定値06 L/DL4
        iTldl(5, 6) = .MS06LDL5         '測定値06 L/DL5
        iTldl(1, 7) = .MS07LDL1         '測定値07 L/DL1
        iTldl(2, 7) = .MS07LDL2         '測定値07 L/DL2
        iTldl(3, 7) = .MS07LDL3         '測定値07 L/DL3
        iTldl(4, 7) = .MS07LDL4         '測定値07 L/DL4
        iTldl(5, 7) = .MS07LDL5         '測定値07 L/DL5
        iTldl(1, 8) = .MS08LDL1         '測定値08 L/DL1
        iTldl(2, 8) = .MS08LDL2         '測定値08 L/DL2
        iTldl(3, 8) = .MS08LDL3         '測定値08 L/DL3
        iTldl(4, 8) = .MS08LDL4         '測定値08 L/DL4
        iTldl(5, 8) = .MS08LDL5         '測定値08 L/DL5
        iTldl(1, 9) = .MS09LDL1         '測定値09 L/DL1
        iTldl(2, 9) = .MS09LDL2         '測定値09 L/DL2
        iTldl(3, 9) = .MS09LDL3         '測定値09 L/DL3
        iTldl(4, 9) = .MS09LDL4         '測定値09 L/DL4
        iTldl(5, 9) = .MS09LDL5         '測定値09 L/DL5
        iTldl(1, 10) = .MS10LDL1        '測定値10 L/DL1
        iTldl(2, 10) = .MS10LDL2        '測定値10 L/DL2
        iTldl(3, 10) = .MS10LDL3        '測定値10 L/DL3
        iTldl(4, 10) = .MS10LDL4        '測定値10 L/DL4
        iTldl(5, 10) = .MS10LDL5        '測定値10 L/DL5
        iTldl(1, 11) = .MS11LDL1        '測定値11 L/DL1
        iTldl(2, 11) = .MS11LDL2        '測定値11 L/DL2
        iTldl(3, 11) = .MS11LDL3        '測定値11 L/DL3
        iTldl(4, 11) = .MS11LDL4        '測定値11 L/DL4
        iTldl(5, 11) = .MS11LDL5        '測定値11 L/DL5
        iTldl(1, 12) = .MS12LDL1        '測定値12 L/DL1
        iTldl(2, 12) = .MS12LDL2        '測定値12 L/DL2
        iTldl(3, 12) = .MS12LDL3        '測定値12 L/DL3
        iTldl(4, 12) = .MS12LDL4        '測定値12 L/DL4
        iTldl(5, 12) = .MS12LDL5        '測定値12 L/DL5
        iTldl(1, 13) = .MS13LDL1        '測定値13 L/DL1
        iTldl(2, 13) = .MS13LDL2        '測定値13 L/DL2
        iTldl(3, 13) = .MS13LDL3        '測定値13 L/DL3
        iTldl(4, 13) = .MS13LDL4        '測定値13 L/DL4
        iTldl(5, 13) = .MS13LDL5        '測定値13 L/DL5
        iTldl(1, 14) = .MS14LDL1        '測定値14 L/DL1
        iTldl(2, 14) = .MS14LDL2        '測定値14 L/DL2
        iTldl(3, 14) = .MS14LDL3        '測定値14 L/DL3
        iTldl(4, 14) = .MS14LDL4        '測定値14 L/DL4
        iTldl(5, 14) = .MS14LDL5        '測定値14 L/DL5
        iTldl(1, 15) = .MS15LDL1        '測定値15 L/DL1
        iTldl(2, 15) = .MS15LDL2        '測定値15 L/DL2
        iTldl(3, 15) = .MS15LDL3        '測定値15 L/DL3
        iTldl(4, 15) = .MS15LDL4        '測定値15 L/DL4
        iTldl(5, 15) = .MS15LDL5        '測定値15 L/DL5
        
        iTdvd2(1) = .MS01DVD2           '測定値01 DVD2
        iTdvd2(2) = .MS02DVD2           '測定値02 DVD2
        iTdvd2(3) = .MS03DVD2           '測定値03 DVD2
        iTdvd2(4) = .MS04DVD2           '測定値04 DVD2
        iTdvd2(5) = .MS05DVD2           '測定値05 DVD2
    End With
    
    'Denの仕様が検査有り,保証有りの場合
    If (GD.JudgFlagDen = "1" And GD.GuaranteeDen.cJudg = JudgCodeC01) Or _
       (GD.JudgFlagDvd2 = "1" And GD.GuaranteeDvd2.cJudg = JudgCodeC01 And GD.SpecDvd2Min = 0 And GD.SpecDvd2Max = 0) Then
    
        'Denの測定値がﾗｲﾝ数分あるかをﾁｪｯｸ
        For iPoint = 1 To 15
            '測定点7まで
            If iPoint <= 7 Then
                '仕様が3ﾗｲﾝの場合
                If GD.SpecGdLine = 3 Then
                    iLine = 3
                '仕様が4.5ﾗｲﾝ又は5ﾗｲﾝの場合
                ElseIf GD.SpecGdLine = 4.5 Or GD.SpecGdLine = 5 Then
                    iLine = 5
                End If
            '測定点8から
            Else
                '仕様が3ﾗｲﾝの場合
                If GD.SpecGdLine = 3 Then
                    iLine = 3
                '仕様が4.5ﾗｲﾝの場合
                ElseIf GD.SpecGdLine = 4.5 Then
                    iLine = 4
                '仕様が5ﾗｲﾝの場合
                Else
                    iLine = 5
                End If
            End If
            
            For iCnt = 1 To iLine
                'DENの測定値がない場合
                If iTden(iCnt, iPoint) = -1 Then
                    ChkGD_Data = FUNCTION_RETURN_FAILURE
                    '判定ｴﾗｰ(処理を抜ける)
                    Exit Function
                End If
            Next iCnt
        Next iPoint
    End If
    
    'LDLの仕様が検査有り,保証有りの場合
    If GD.JudgFlagLdl = "1" And GD.GuaranteeLdl.cJudg = JudgCodeC01 Then
    
        'L/DLの測定値がﾗｲﾝ数分あるかをﾁｪｯｸ
        For iPoint = 1 To 15
            '測定点7まで
            If iPoint <= 7 Then
                '仕様が3ﾗｲﾝの場合
                If GD.SpecGdLine = 3 Then
                    iLine = 3
                '仕様が4.5ﾗｲﾝの場合
                ElseIf GD.SpecGdLine = 4.5 Or GD.SpecGdLine = 5 Then
                    iLine = 5
                End If
            '測定点8から
            Else
                '仕様が3ﾗｲﾝの場合
                If GD.SpecGdLine = 3 Then
                    iLine = 3
                '仕様が4.5ﾗｲﾝの場合
                ElseIf GD.SpecGdLine = 4.5 Then
                    iLine = 4
                '仕様が5ﾗｲﾝの場合
                Else
                    iLine = 5
                End If
            End If
            
            For iCnt = 1 To iLine
                'L/DLの測定値がない場合
                If iTldl(iCnt, iPoint) = -1 Then
                    ChkGD_Data = FUNCTION_RETURN_FAILURE
                    '判定ｴﾗｰ(処理を抜ける)
                    Exit Function
                End If
            Next iCnt
        Next iPoint
    End If
    
    ChkGD_Data = FUNCTION_RETURN_SUCCESS
    
End Function

'概要      :LifeTime判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :仕様情報構造体
'          :typ_ltz       ,I  ,type_DBDRV_scmzc_fcmkc001c_LT        :LT実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(True:判定OK, False:判定NG)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :LifeTime判定を行う
'履歴      :
Public Function CrLtjudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                         typ_ltz As type_DBDRV_scmzc_fcmkc001c_LT, _
                         bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim Lt      As C_LT                     'LifeTime構造体
    
    bJudg = True
        
    'LT判定引数設定
    Lt.GuaranteeLt.cMeth = typ_si.HSXLTSPH         '測定位置_方
    Lt.GuaranteeLt.cCount = typ_si.HSXLTSPT        '測定位置_点
    Lt.GuaranteeLt.cPos = typ_si.HSXLTSPI          '測定位置_位(OSFの場合 領)
    Lt.GuaranteeLt.cObj = typ_si.HSXLTHWT          '保証方法_対
    Lt.GuaranteeLt.cJudg = typ_si.HSXLTHWS         '保証方法_処
    
    Lt.SpecLtMin = typ_si.HSXLTMIN                 '品SXLタイム下限
    Lt.SpecLtMax = typ_si.HSXLTMAX                 '品SXLタイム上限

    Lt.Lt = typ_ltz.CALCMEAS                       'ライフタイム計算値
    
    'LT判定
    If CrystalLTJudg(Lt, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrLtjudg = False
        Exit Function
    End If
    
    If Lt.JudgLt <> True Then
        bJudg = False
    End If
    
    CrLtjudg = True

End Function

'概要      :LT10判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :仕様情報構造体
'          :typ_ltz       ,I  ,type_DBDRV_scmzc_fcmkc001c_LT        :LT実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(True:判定OK, False:判定NG)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :LT10判定を行う
'履歴      :2011/07/22 T.Koi(SETsw)
Public Function CrLt10judg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                         typ_ltz As type_DBDRV_scmzc_fcmkc001c_LT, _
                         typ_cr As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                         bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim Lt      As C_LT                     'LifeTime構造体
    Dim CRYREST10CS As String
    
    bJudg = True
    
    'LT判定引数設定
    Lt.GuaranteeLt.cMeth = typ_si.HSXLTSPH         '測定位置_方
    Lt.GuaranteeLt.cCount = typ_si.HSXLTSPT        '測定位置_点
    Lt.GuaranteeLt.cPos = typ_si.HSXLTSPI          '測定位置_位(OSFの場合 領)
    Lt.GuaranteeLt.cObj = typ_si.HSXLTHWT          '保証方法_対
    Lt.GuaranteeLt.cJudg = typ_si.HSXLTHWS         '保証方法_処
    
    'LT10実績フラグ
    CRYREST10CS = typ_cr.CRYREST10CS               'LT10実績フラグ
    If CRYREST10CS = "9" Then                      '対象外
        CrLt10judg = True
        Exit Function
    End If
    
    Lt.SpecLt10Min = typ_si.HSXLT10MIN               '品SXLタイム下限

    Lt.Lt10 = typ_ltz.CONVAL                         'LT10Ω計算値
    
    'LT10判定
    If CrystalLT10Judg(Lt, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrLt10judg = False
        Exit Function
    End If
    
    If Lt.JudgLt10 <> True Then
        bJudg = False
    End If
    
    CrLt10judg = True

End Function

'概要      :EPD判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :仕様情報構造体
'          :typ_epdz      ,I  ,type_DBDRV_scmzc_fcmkc001c_EPD       :EPD実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(True:判定OK, False:判定NG)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :EPD判定を行う
'履歴      :
Public Function CrEpdjudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_epdz As type_DBDRV_scmzc_fcmkc001c_EPD, _
                          bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim ep      As C_EPD                    'EPD構造体
    
    bJudg = True
        
    'EPD判定引数設定
    ep.SpecEpdMax = typ_si.EPDUP            '結晶内側管理､EPD上限
    ep.EPD = typ_epdz.MEASURE               'EPD測定値
        
    'EPD判定
    If CrystalEPDJudg(ep, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrEpdjudg = False
        Exit Function
    End If
        
    If ep.JudgEpd <> True Then
        bJudg = False
    End If
    
    CrEpdjudg = True

End Function
'概要      :X線判定 2009/08/12 Kameda
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :仕様情報構造体
'          :typ_xz        ,I  ,type_DBDRV_scmzc_fcmkc001c_X         :X線実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(True:判定OK, False:判定NG)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :X線判定を行う
'履歴      :2009/10/22 判定は合成角のみで行う(X,Yが外れている時は赤表示)
Public Function CrXjudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_xz As type_DBDRV_scmzc_fcmkc001c_X, _
                          bJudgXY As Boolean, bJudgX As Boolean, bJudgY As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim x      As C_XY                    'X判定構造体
    
    'bJudg = True       2009/10/22 Kameda
    bJudgXY = True
    bJudgX = True
    bJudgY = True
        
    'X線判定引数設定
    '合成
    x.SpecXY_Min = typ_si.HSXCSMIN
    x.SpecXY_Max = typ_si.HSXCSMAX
    x.Spec_XY = typ_xz.XXY
    
    
    '縦
    x.SpecY_Min = typ_si.HSXCTMIN
    x.SpecY_Max = typ_si.HSXCTMAX
    x.Spec_Y = typ_xz.XY
    
    '
    x.SpecX_Min = typ_si.HSXCYMIN
    x.SpecX_Max = typ_si.HSXCYMAX
    x.Spec_X = typ_xz.XX
    
    If CrystalXYJudg(x, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        'bJudg = False      2009/10/22
        bJudgXY = False
        CrXjudg = False
        Exit Function
    End If
        
    If x.JudgResult_XY <> True Then
        'bJudg = False      2009/10/22
        bJudgXY = False
    End If
    If x.JudgResult_Y <> True Then
        'bJudg = False      2009/10/22
        bJudgY = False
    End If
    If x.JudgResult_X <> True Then
        'bJudg = False      2009/10/22
        bJudgX = False
    End If
    
    CrXjudg = True

End Function

'概要      :SIRD判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :仕様情報構造体
'          :typ_epdz      ,I  ,type_DBDRV_scmzc_fcmkc001c_SIRD      :SIRD実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(True:判定OK, False:判定NG)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :SIRD判定を行う
'履歴      :2010/02/04 Kameda
Public Function CrSIRDjudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_sird As type_DBDRV_scmzc_fcmkc001c_SIRD, _
                          bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim SIRD      As C_SIRD                    'SIRD構造体
    
    bJudg = True
        
    'SIRD判定引数設定
    SIRD.SpecSirdMax = typ_si.HWFSIRDMX     '仕様面内個数上限
    SIRD.SIRDCNT = typ_sird.SIRDCNT         'SIRD測定値
        
    'SIRD判定
    If CrystalSIRDJudg(SIRD, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrSIRDjudg = False
        Exit Function
    End If
        
    If SIRD.JudgSird <> True Then
        bJudg = False
    End If
    
    CrSIRDjudg = True

End Function

'Add Start 2011/01/31 SMPK Miyata
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :仕様情報構造体
'          :typ_cz        ,I  ,type_DBDRV_scmzc_fcmkc001c_C         :C実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(True:判定OK, False:判定NG)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :C判定を行う
'履歴      :
Public Function CrCjudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_cz As type_DBDRV_scmzc_fcmkc001c_C, _
                          bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim C       As C_C                      'C判定構造体
    
    bJudg = True
        
    'C判定引数設定
    C.GuaranteeC.cObj = typ_si.HSXCHT       '品ＳＸＣ保証方法＿対
    C.GuaranteeC.cJudg = typ_si.HSXCHS      '品ＳＸＣ保証方法＿処

    C.HSXCPK = typ_si.HSXCPK                ''品ＳＸＣパターン区分
    C.HSXCSZ = typ_si.HSXCSZ                ''品ＳＸＣ測定条件
    C.CPTNJSK = typ_cz.CPTNJSK              ''C パターン実績
    C.CDISKJSK = typ_cz.CDISKJSK            ''C Disk半径実績
    C.CRINGNKJSK = typ_cz.CRINGNKJSK        ''C Ring内径実績
    C.CRINGGKJSK = typ_cz.CRINGGKJSK        ''C Ring外径実績

    If CrystalCJudg(C, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrCjudg = False
        Exit Function
    End If
        
    If C.JudgC <> True Then
        bJudg = False
    End If
    
    CrCjudg = True

End Function

'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :仕様情報構造体
'          :typ_cjz       ,I  ,type_DBDRV_scmzc_fcmkc001c_CJ        :CJ実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(True:判定OK, False:判定NG)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :CJ判定を行う
'履歴      :
Public Function CrCJjudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_cjz As type_DBDRV_scmzc_fcmkc001c_CJ, _
                          bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim CJ      As C_CJ                     'CJ判定構造体
    
    bJudg = True
        
    'CJ判定引数設定
    CJ.GuaranteeCJ.cObj = typ_si.HSXCJHT        '品ＳＸＣＪ保証方法＿対
    CJ.GuaranteeCJ.cJudg = typ_si.HSXCJHS       '品ＳＸＣＪ保証方法＿処

    CJ.HSXCJPK = typ_si.HSXCJPK                 ''品ＳＸＣＪパターン区分
    CJ.HSXCJNS = typ_si.HSXCJNS                 ''品ＳＸＣＪ熱処理法

    CJ.CJPTNJSK = typ_cjz.CJPTNJSK              ''CJ パターン実績
    CJ.CJDISKJSK = typ_cjz.CJDISKJSK            ''CJ Disk半径実績
    CJ.CJRINGNKJSK = typ_cjz.CJRINGNKJSK        ''CJ Ring内径実績
    CJ.CJRINGGKJSK = typ_cjz.CJRINGGKJSK        ''CJ Ring外径実績
    CJ.CJBANDNKJSK = typ_cjz.CJBANDNKJSK        ''CJ Band内径実績
    CJ.CJBANDGKJSK = typ_cjz.CJBANDGKJSK        ''CJ Band外径実績
    CJ.CJRINGCALC = typ_cjz.CJRINGCALC          ''CJ Ring幅計算
    CJ.CJPICALC = typ_cjz.CJPICALC              ''CJ Pi幅計算
    CJ.CJHANTEI = typ_cjz.CJHANTEI              ''CJ 判定結果
    CJ.CJDMAXPIC5 = typ_cjz.CJDMAXPIC5          ''CJ Diskのみパターン Pi幅上限値
    CJ.CJRMAXPIC5 = typ_cjz.CJRMAXPIC5          ''CJ Ringのみパターン Pi幅上限値
    CJ.CJDRMAXPIC5 = typ_cjz.CJDRMAXPIC5        ''CJ DiskRingパターン Pi幅上限値
    CJ.CJALLMAXDIC5 = typ_cjz.CJALLMAXDIC5      ''CJ 共通Disk半径上限値
    CJ.CJALLMINRINC5 = typ_cjz.CJALLMINRINC5    ''CJ 共通Ring内径下限値
    CJ.CJALLMAXRIGC5 = typ_cjz.CJALLMAXRIGC5    ''CJ 共通Ring外径上限値

    If CrystalCJJudg(CJ, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrCJjudg = False
        Exit Function
    End If
        
    If CJ.JudgCJ <> True Then
        bJudg = False
    End If
    
    CrCJjudg = True

End Function

'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :仕様情報構造体
'          :typ_cjltz     ,I  ,type_DBDRV_scmzc_fcmkc001c_CJLT      :CJLT実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(True:判定OK, False:判定NG)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :CJLT判定を行う
'履歴      :
Public Function CrCJLTjudg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_cjltz As type_DBDRV_scmzc_fcmkc001c_CJLT, _
                          bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim CJLT    As C_CJLT                   'CJ判定構造体
    
    bJudg = True
        
    'CJLT判定引数設定
    CJLT.GuaranteeCJLT.cObj = typ_si.HSXCJLTHT  '品ＳＸＣＪＬＴ保証方法＿対
    CJLT.GuaranteeCJLT.cJudg = typ_si.HSXCJLTHS '品ＳＸＣＪＬＴ保証方法＿処

    CJLT.HSXCJLTPK = typ_si.HSXCJLTPK           '品ＳＸＣＪＬＴパターン区分
    CJLT.HSXCJLTNS = typ_si.HSXCJLTNS           '品ＳＸＣＪＬＴ熱処理法
    CJLT.CJLTPTNJSK = typ_cjltz.CJLTPTNJSK      ''CJ(LT) パターン実績
    CJLT.CJLTDISKJSK = typ_cjltz.CJLTDISKJSK    ''CJ(LT) Disk半径実績
    CJLT.CJLTRINGNKJSK = typ_cjltz.CJLTRINGNKJSK ''CJ(LT) Ring内径実績
    CJLT.CJLTRINGGKJSK = typ_cjltz.CJLTRINGGKJSK ''CJ(LT) Ring外径実績
    CJLT.CJLTBANDNKJSK = typ_cjltz.CJLTBANDNKJSK ''CJ(LT) Band内径実績
    CJLT.CJLTBANDGKJSK = typ_cjltz.CJLTBANDGKJSK ''CJ(LT) Band外径実績
    CJLT.CJLTRINGCALC = typ_cjltz.CJLTRINGCALC  ''CJ(LT) Ring幅計算
    CJLT.CJLTPICALC = typ_cjltz.CJLTPICALC      ''CJ(LT) Pi幅計算
    CJLT.CJLTBANDCALC = typ_cjltz.CJLTBANDCALC  ''CJ(LT) Band幅計算
    CJLT.HSXCJLTBND = typ_cjltz.HSXCJLTBND      ''CJ(LT) Band幅上限値

    If CrystalCJLTJudg(CJLT, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrCJLTjudg = False
        Exit Function
    End If
        
    If CJLT.JudgCJLT <> True Then
        bJudg = False
    End If
    
    CrCJLTjudg = True

End Function

'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :仕様情報構造体
'          :typ_cj2z      ,I  ,type_DBDRV_scmzc_fcmkc001c_CJ2       :CJ2実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(True:判定OK, False:判定NG)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :CJLT判定を行う
'履歴      :
Public Function CrCJ2judg(typ_si As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_cj2z As type_DBDRV_scmzc_fcmkc001c_CJ2, _
                          bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim CJ2     As C_CJ2                    'CJ判定構造体
    
    bJudg = True
        
    'CJ2判定引数設定
    CJ2.GuaranteeCJ2.cObj = typ_si.HSXCJ2HT     '品ＳＸＣＪ２保証方法＿対
    CJ2.GuaranteeCJ2.cJudg = typ_si.HSXCJ2HS    '品ＳＸＣＪ２保証方法＿処

    CJ2.HSXCJ2PK = typ_si.HSXCJ2PK              '品ＳＸＣＪＬＴパターン区分
    CJ2.HSXCJ2NS = typ_si.HSXCJ2NS              '品ＳＸＣＪＬＴ熱処理法

    CJ2.CJ2PTNJSK = typ_cj2z.CJ2PTNJSK          ''CJ2 パターン実績
    CJ2.CJ2DISKJSK = typ_cj2z.CJ2DISKJSK        ''CJ2 Disk半径実績
    CJ2.CJ2RINGNKJSK = typ_cj2z.CJ2RINGNKJSK    ''CJ2 Ring内径実績
    CJ2.CJ2RINGGKJSK = typ_cj2z.CJ2RINGGKJSK    ''CJ2 Ring外径実績
    CJ2.CJ2PICALC = typ_cj2z.CJ2PICALC          ''CJ2 Pi幅計算
    CJ2.CJ2HANTEI = typ_cj2z.CJ2HANTEI          ''CJ2 判定結果
    CJ2.CJ2DMAXPIC5 = typ_cj2z.CJ2DMAXPIC5      ''CJ2 Diskのみパターン Pi幅上限値
    CJ2.CJ2RMAXPIC5 = typ_cj2z.CJ2RMAXPIC5      ''CJ2 Ringのみパターン Pi幅上限値
    CJ2.CJ2RMINRINC5 = typ_cj2z.CJ2RMINRINC5    ''CJ2 Ringのみパターン Ring内径下限値
    CJ2.CJ2RMAXRIGC5 = typ_cj2z.CJ2RMAXRIGC5    ''CJ2 Ringのみパターン Ring外径上限値
    CJ2.CJ2DRMAXPIC5 = typ_cj2z.CJ2DRMAXPIC5    ''CJ2 DiskRingパターン Pi幅上限値
    CJ2.CJ2DRMINRINC5 = typ_cj2z.CJ2DRMINRINC5  ''CJ2 DiskRingパターン Ring内径下限値
    CJ2.CJ2DRMAXRIGC5 = typ_cj2z.CJ2DRMAXRIGC5  ''CJ2 DiskRingパターン Ring外径上限値

    If CrystalCJ2Judg(CJ2, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        CrCJ2judg = False
        Exit Function
    End If
        
    If CJ2.JudgCJ2 <> True Then
        bJudg = False
    End If
    
    CrCJ2judg = True

End Function

'Add End   2011/01/31 SMPK Miyata

Private Function NtoS(strWk As String) As String
    If Mid(strWk, 1, 1) = Chr(0) Then
        NtoS = " "
        Exit Function
    End If
    NtoS = strWk
End Function

Private Sub BMDDataSet(BmdNo As Integer, UpDo As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String)
    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4                       '検査指示
    Dim typ_bmdz        As type_DBDRV_scmzc_fcmkc001c_BMD
    Dim bJudg           As Boolean
    Dim vTemp           As Variant
    
    '検査指示設定
    IND = IIf(UpDo = BlkTop, "123", "123")
    
    With typ_b
        Select Case BmdNo
        Case 1
            JudgSpecCode = JudgSC_B(UpDo).B1
            SCC = "B1"
            shiji = (InStr(IND, .typ_cr(UpDo).CRYINDB1CS) <> 0)
            typ_bmdz = .typ_zi.BMD1Z(UpDo)
        Case 2
            JudgSpecCode = JudgSC_B(UpDo).B2
            SCC = "B2"
            shiji = (InStr(IND, .typ_cr(UpDo).CRYINDB2CS) <> 0)
            typ_bmdz = .typ_zi.BMD2Z(UpDo)
        Case 3
            JudgSpecCode = JudgSC_B(UpDo).B3
            SCC = "B3"
            shiji = (InStr(IND, .typ_cr(UpDo).CRYINDB3CS) <> 0)
            typ_bmdz = .typ_zi.BMD3Z(UpDo)
        End Select
        
        If JudgSpecCode Then
            '画面表示内容設定
            .typ_rslt(UpDo, DispLineCount).pos = -1                                     ' 結晶内開始位置
            .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' 内容
            .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                             ' 情報１
            .typ_rslt(UpDo, DispLineCount).INFO2 = "検査無"                             ' 情報２
            .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                   ' 情報３
            .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
            .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                                  ' サンプルＮｏ
            .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' 品番(12桁)
            bJudg = False
            If shiji Then
                '画面表示内容設定
                .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                         ' 情報１
                .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                         ' 情報２
                .typ_rslt(UpDo, DispLineCount).INFO3 = "実績無"                         ' 情報３
                If left(typ_bmdz.CRYNUM, 9) = left(.BLOCKID, 9) Then
                    .typ_rslt(UpDo, DispLineCount).pos = typ_bmdz.POSITION              ' 結晶内開始位置
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = typ_bmdz.SMPLNO             ' サンプルＮｏ
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ無"                    ' 情報３
                    If typ_bmdz.SMPLUMU = "0" Then
                        'BMD1判定失敗
                        .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                                ' 情報２
                        'BMD1判定
                        If CrBmdJudg(.typ_si(UpDo), typ_bmdz, bJudg, BmdNo) Then
                            vTemp = CStr(typ_bmdz.MEASAVE)                                              ' 情報１
                            .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.0")        ' 情報１
                            vTemp = CStr(typ_bmdz.MEASMAX)                                              ' 情報２
                            .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")        ' 情報２
                            .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                   ' 情報３
' OSF，BMD項目追加対応  2002.04.02 yakimura
                            vTemp = CStr(typ_bmdz.BMDMNBUNP)
                            .typ_rslt(UpDo, DispLineCount).INFO4 = DBData2DispData(vTemp, "0.0")        ' 情報4
' OSF，BMD項目追加対応  2002.04.02 yakimura
                        End If
                    End If
                End If
            End If
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                               ' 判定結果
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                               ' 判定結果
                TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                Select Case BmdNo
                Case 1
                    gsTbcmy028ErrCode = "00107"
                Case 2
                    gsTbcmy028ErrCode = "00108"
                Case 3
                    gsTbcmy028ErrCode = "00109"
                End Select
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            DispLineCount = DispLineCount + 1
        Else
            If shiji Then
                '画面表示内容設定
                .typ_rslt(UpDo, DispLineCount).pos = typ_bmdz.POSITION                      ' 結晶内開始位置
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' 内容
                .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様無"                             ' 情報１
                .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                             ' 情報２
                .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ無"                            ' 情報３
' OSF，BMD項目追加対応  2002.04.02 yakimura
                .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報4
' OSF，BMD項目追加対応  2002.04.02 yakimura
                .typ_rslt(UpDo, DispLineCount).SMPLNO = typ_bmdz.SMPLNO                     ' サンプルＮｏ
'====================== Debug Debug =====================================
                .typ_rslt(UpDo, DispLineCount).OKNG = "N参"                                  ' 判定結果
'====================== Debug Debug =====================================
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' 品番(12桁)
                If typ_bmdz.SMPLUMU = "0" Then
                    'BMD1判定失敗
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                                ' 情報２
                    'BMD1判定
                    If CrBmdJudg(.typ_si(UpDo), typ_bmdz, bJudg, BmdNo) Then
                        vTemp = CStr(typ_bmdz.MEASAVE)                                              ' 情報１
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.0")        ' 情報１
                        vTemp = CStr(typ_bmdz.MEASMAX)                                              ' 情報２
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")        ' 情報２
                        .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                   ' 情報２
' OSF，BMD項目追加対応  2002.04.02 yakimura
                        vTemp = CStr(typ_bmdz.BMDMNBUNP)
                        .typ_rslt(UpDo, DispLineCount).INFO4 = DBData2DispData(vTemp, "0.0")        ' 情報4
' OSF，BMD項目追加対応  2002.04.02 yakimura
                    End If
                End If
                DispLineCount = DispLineCount + 1
            End If
        End If
    End With
End Sub

'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 UPD By Systech Start
'' 品SXOSF1(ArAN)パタン区分を引数に追加
Private Sub OSFDataSet(OsfNo As Integer, UpDo As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String, sAranPtn As String)
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 UPD By Systech End
    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4                  '検査指示
    Dim typ_Osfz        As type_DBDRV_scmzc_fcmkc001c_OSF
    Dim bJudg           As Boolean
    Dim vTemp           As Variant
    Dim w_1             As String
    Dim w_2             As String
    Dim w_3             As String
    
'C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
    Dim strXTALC1       As String
    Dim strJDGEIDC      As String
    Dim strSYNFLG       As String
    Dim strYMKFLG       As String
    Dim strOSFRD1       As String
    Dim strOSFRD2       As String
    Dim lngOSFWID1      As Long
    Dim lngOSFWID2      As Long
    Dim lngSMPPOS       As Long
    Dim strRMAXC5       As String
    Dim strDMAXC5       As String
    Dim strDRRMAXC5     As String
    Dim strDRDMAXC5     As String
    Dim lngRMAXC5       As Long
    Dim lngDMAXC5       As Long
    Dim lngDRRMAXC5     As Long
    Dim lngDRDMAXC5     As Long
    Dim ErrFlg          As Boolean
    Dim SYNErrFlg       As Boolean
    Dim DBErrFlg        As Boolean
    Dim YFlg            As Boolean
       
'C－OSF3判定機能追加 2007/04/23 M.Kaga END ---

    'OSF3判定用ﾌﾗｸﾞ初期化
    gsCOSF3Flg = ""

    '検査指示設定
    IND = IIf(UpDo = BlkTop, "123", "123")
        
    With typ_b
        Select Case OsfNo
        Case 1
            JudgSpecCode = JudgSC_B(UpDo).L1
            SCC = "L1"
            shiji = (InStr(IND, .typ_cr(UpDo).CRYINDL1CS) <> 0)
            typ_Osfz = .typ_zi.OSF1Z(UpDo)
        Case 2
            JudgSpecCode = JudgSC_B(UpDo).L2
            SCC = "L2"
            shiji = (InStr(IND, .typ_cr(UpDo).CRYINDL2CS) <> 0)
            typ_Osfz = .typ_zi.OSF2Z(UpDo)
        Case 3
            JudgSpecCode = JudgSC_B(UpDo).L3
            SCC = "L3"
            shiji = (InStr(IND, .typ_cr(UpDo).CRYINDL3CS) <> 0)
            typ_Osfz = .typ_zi.OSF3Z(UpDo)
        Case 4
'C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
            JudgSpecCode = JudgSC_B(UpDo).COSF3
            SCC = "COSF3"
'C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---
            'SCC = "L4"
            shiji = (InStr(IND, .typ_cr(UpDo).CRYINDL4CS) <> 0)
            typ_Osfz = .typ_zi.OSF4Z(UpDo)
        End Select
                   
        '保証ﾌﾗｸﾞ="H"の場合
        If JudgSpecCode Then
        
'C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
            '熱処理区分:OSF4の場合
            If OsfNo = 4 Then
                'ﾌﾞﾛｯｸID
                strXTALC1 = Trim(typ_b.BLOCKID)
                '結晶番号
                strXTALC1 = left(strXTALC1, 9) & "000"

                'OSF実績入力判定処理
                '判定ﾊﾟﾀｰﾝ&実績値退避
                
                If Trim(typ_Osfz.OSFRD1) = "R" Or Trim(typ_Osfz.OSFRD1) = "D" Then
                    strOSFRD1 = Trim(typ_Osfz.OSFRD1)
                Else
                    strOSFRD1 = "-"
                End If
                
                If Trim(typ_Osfz.OSFRD2) = "D" Then
                    strOSFRD2 = Trim(typ_Osfz.OSFRD2)
                Else
                    strOSFRD2 = "-"
                End If
                
                If IsNull(typ_Osfz.OSFWID1) = True Then
                   lngOSFWID1 = -1
                ElseIf IsNumeric(typ_Osfz.OSFWID1) = False Then
                   lngOSFWID1 = -1
                Else
                   lngOSFWID1 = Trim(typ_Osfz.OSFWID1)
                End If
                
                If IsNull(typ_Osfz.OSFWID2) = True Then
                   lngOSFWID2 = -1
                ElseIf IsNumeric(typ_Osfz.OSFWID2) = False Then
                   lngOSFWID2 = -1
                Else
                   lngOSFWID2 = Trim(typ_Osfz.OSFWID2)
                End If
                
                '-1以外の数値考慮
                If lngOSFWID1 < 0 Then
                   lngOSFWID1 = -1
                End If
                If lngOSFWID2 < 0 Then
                   lngOSFWID2 = -1
                End If

               'ｻﾝﾌﾟﾙ位置
               lngSMPPOS = Trim(typ_Osfz.POSITION)

               '判定ﾌﾗｸﾞ初期化
               ErrFlg = True
               YFlg = False
                               
               'ﾊﾟﾀｰﾝ区分、実績値がNULLの場合
               If strOSFRD1 = "-" And strOSFRD2 <> "-" Then
                   ErrFlg = False
               ElseIf strOSFRD1 <> "-" And lngOSFWID1 = -1 Then
                    ErrFlg = False
               ElseIf strOSFRD2 <> "-" And lngOSFWID2 = -1 Then
                    ErrFlg = False
               ElseIf strOSFRD2 = "-" And lngOSFWID2 > 0 Then
                   ErrFlg = False
               End If
               
                '結晶番号をｷｰとしてXSDC1よりC－OSF3判定IDを獲得する
               If GetCOSF3ID(strJDGEIDC, strXTALC1) <> FUNCTION_RETURN_SUCCESS Then
                    ErrFlg = False
               Else
                   If Trim(strJDGEIDC) = "" Then
                        gsCOSF3Flg = "1"
                        ErrFlg = False
                   '判定ID=「9」の場合は判定なし(判定OK)　07/08/01 M.Kaga
                   ElseIf Trim(strJDGEIDC) = "9" Then
                        YFlg = True
                        bJudg = True
                   Else
                       '獲得したC-OSF3判定IDでXODC5_OSF30より承認ﾌﾗｸﾞの獲得
                       If GetSYNFLAGC5(strSYNFLG, strYMKFLG, strJDGEIDC) <> FUNCTION_RETURN_SUCCESS Then
                          ErrFlg = False
                       Else
                          '承認ﾌﾗｸﾞ:0　未承認の場合
                          If Trim(strSYNFLG) = "0" Or Trim(strSYNFLG) = "" Or IsNull(strSYNFLG) Then
                             gsCOSF3Flg = "2"
                             ErrFlg = False
                          End If
                       End If
                   End If
               End If
               
               If ErrFlg = False Then
                   '画面表示内容設定
                   .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' 結晶内開始位置
                   .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' 内容
                   .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                         ' 情報１
                   .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                         ' 情報２
                   .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                      ' 情報３
                   .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' サンプルＮｏ
                   .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' 品番(12桁)
                   If left(typ_Osfz.CRYNUM, 9) = left(.BLOCKID, 9) Then
                       .typ_rslt(UpDo, DispLineCount).pos = typ_Osfz.POSITION              ' 結晶内開始位置
                       .typ_rslt(UpDo, DispLineCount).SMPLNO = typ_Osfz.SMPLNO             ' サンプルＮｏ
                   End If
                   bJudg = False
               Else
                   If YFlg = False Then
                       'ﾊﾟﾀｰﾝ区分により処理分岐
                       'Rのみの場合
                       If strOSFRD1 = "R" And strOSFRD2 = "-" Then
                           'Rのみ上限値の獲得を行う
                           If GetCOSF3PTN(strJDGEIDC, lngSMPPOS, strOSFRD1, strOSFRD2, strRMAXC5, strDMAXC5, strDRRMAXC5, strDRDMAXC5) <> FUNCTION_RETURN_SUCCESS Then
                               ErrFlg = False
                           End If
                           'ﾚｺｰﾄﾞ無：VBエラー(後で考える)
                           If Trim(strRMAXC5) = "" Then
                               ErrFlg = False
                           Else
                               lngRMAXC5 = Trim(strRMAXC5)
                               '実績値の判定
    
                               If lngOSFWID1 <= lngRMAXC5 Then
                                   '判定OK
                                   bJudg = True
                               ElseIf lngOSFWID1 > lngRMAXC5 Then
                                   '判定NG
                                   bJudg = False
                               End If
                           End If
                       'Dのみの場合
                       ElseIf strOSFRD1 = "D" Then
                           'Dのみ上限値の獲得を行う
                           If GetCOSF3PTN(strJDGEIDC, lngSMPPOS, strOSFRD1, strOSFRD2, strRMAXC5, strDMAXC5, strDRRMAXC5, strDRDMAXC5) <> FUNCTION_RETURN_SUCCESS Then
                               ErrFlg = False
                           End If
    
                           'ﾚｺｰﾄﾞ無又はﾏｽﾀの実績値がNULL：VBエラー(後で考える)
                           If Trim(strDMAXC5) = "" Then
                               ErrFlg = False
                           Else
                               lngDMAXC5 = Trim(strDMAXC5)
                               '実績値の判定
                               If lngOSFWID1 <= lngDMAXC5 Then
                                   '判定OK
                                   bJudg = True
                               ElseIf lngOSFWID1 > lngDMAXC5 Then
                                   '判定NG
                                   bJudg = False
                               End If
                           End If
                       'R&Dの場合
                       ElseIf strOSFRD1 = "R" And strOSFRD2 = "D" Then
                           'D共存上限値並びR共存上限値の獲得を行う
                           If GetCOSF3PTN(strJDGEIDC, lngSMPPOS, strOSFRD1, strOSFRD2, strRMAXC5, strDMAXC5, strDRRMAXC5, strDRDMAXC5) <> FUNCTION_RETURN_SUCCESS Then
                               ErrFlg = False
                           End If
    
                           'ﾚｺｰﾄﾞ無又はﾏｽﾀの実績値がNULL：VBエラー(後で考える)
                           If Trim(strDRRMAXC5) = "" Or Trim(strDRDMAXC5) = "" Then
                               ErrFlg = False
                           Else
                               lngDRRMAXC5 = Trim(strDRRMAXC5)
                               lngDRDMAXC5 = Trim(strDRDMAXC5)
                               '実績値の判定
                               If lngOSFWID1 <= lngDRRMAXC5 And lngOSFWID2 <= lngDRDMAXC5 Then
                                   '判定OK
                                   bJudg = True
                               ElseIf lngOSFWID1 > lngDRRMAXC5 Or lngOSFWID2 > lngDRDMAXC5 Then
                                   '判定NG
                                   bJudg = False
                               End If
                           End If
                       Else
                           '実績値無し、判定ﾊﾟﾀｰﾝ無しの場合判定OK
                           bJudg = True
                       End If
                   End If
                   If ErrFlg = False Then
                       '画面表示内容設定
                       .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' 結晶内開始位置
                       .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' 内容
                       .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                         ' 情報１
                       .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                         ' 情報２
                       .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                        ' 情報３
                       .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' サンプルＮｏ
                       .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' 品番(12桁)
                       If left(typ_Osfz.CRYNUM, 9) = left(.BLOCKID, 9) Then
                           .typ_rslt(UpDo, DispLineCount).pos = typ_Osfz.POSITION              ' 結晶内開始位置
                           .typ_rslt(UpDo, DispLineCount).SMPLNO = typ_Osfz.SMPLNO             ' サンプルＮｏ
                       End If
                       bJudg = False

                   Else
                       '画面表示内容設定
                       .typ_rslt(UpDo, DispLineCount).pos = -1                                     ' 結晶内開始位置
                       .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' 内容
                       .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                             ' 情報１
                       .typ_rslt(UpDo, DispLineCount).INFO2 = "検査無"                             ' 情報２
                       .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                   ' 情報３
                       .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                       .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                                  ' サンプルＮｏ
                       .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' 品番(12桁)
                       If shiji Then
                           '画面表示内容設定
                           .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                         ' 情報１
                           .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                         ' 情報２
                           .typ_rslt(UpDo, DispLineCount).INFO3 = "実績無"                         ' 情報３
                           If left(typ_Osfz.CRYNUM, 9) = left(.BLOCKID, 9) Then
                               .typ_rslt(UpDo, DispLineCount).pos = typ_Osfz.POSITION              ' 結晶内開始位置
                               .typ_rslt(UpDo, DispLineCount).SMPLNO = typ_Osfz.SMPLNO             ' サンプルＮｏ
                               .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ無"                    ' 情報３
                               If typ_Osfz.SMPLUMU = "0" Then
                                   .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                ' 情報３
                                   '画面表示内容設定
                                    .typ_rslt(UpDo, DispLineCount).INFO1 = strOSFRD1               ' 情報１
                                    .typ_rslt(UpDo, DispLineCount).INFO2 = lngOSFWID1              ' 情報２
                                    .typ_rslt(UpDo, DispLineCount).INFO3 = strOSFRD2               ' 情報３
                                    .typ_rslt(UpDo, DispLineCount).INFO4 = lngOSFWID2              ' 情報４
                               End If
                           End If
                       End If
                   End If
               End If
               If bJudg = True Then
                   .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                               ' 判定結果
               Else
                   .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                               ' 判定結果
                   TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                    Select Case OsfNo
                    Case 1
                        gsTbcmy028ErrCode = "00103"
                    Case 2
                        gsTbcmy028ErrCode = "00104"
                    Case 3
                        gsTbcmy028ErrCode = "00105"
                    Case 4
                        gsTbcmy028ErrCode = "00106"
                    End Select
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
               End If
               DispLineCount = DispLineCount + 1
           Else
'C－OSF3判定機能追加 2007/04/23 M.Kaga END ---

                '画面表示内容設定
                .typ_rslt(UpDo, DispLineCount).pos = -1                                     ' 結晶内開始位置
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' 内容
                .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                             ' 情報１
                .typ_rslt(UpDo, DispLineCount).INFO2 = "検査無"                             ' 情報２
                .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                   ' 情報３
                .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                                  ' サンプルＮｏ
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' 品番(12桁)
                bJudg = False
                If shiji Then
                    '画面表示内容設定
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                         ' 情報１
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                         ' 情報２
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "実績無"                         ' 情報３
                    If left(typ_Osfz.CRYNUM, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(UpDo, DispLineCount).pos = typ_Osfz.POSITION              ' 結晶内開始位置
                        .typ_rslt(UpDo, DispLineCount).SMPLNO = typ_Osfz.SMPLNO             ' サンプルＮｏ
                        .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ無"                    ' 情報３
                        If typ_Osfz.SMPLUMU = "0" Then
                            'OSF判定失敗
                            .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                                        ' 情報３
                            'OSF判定取得
                            If CrOsfJudg(.typ_si(UpDo), typ_Osfz, bJudg, OsfNo) Then
                                '画面表示内容設定
                                vTemp = CStr(typ_Osfz.CALCAVE)                                                      ' 情報１
                                .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")               ' 情報１
                                vTemp = CStr(typ_Osfz.CALCMAX)                                                      ' 情報２
                                .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")                ' 情報２
                                vTemp = CStr(typ_Osfz.MEAS6)                                                        ' 情報３
                                .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")               ' 情報３
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
                                If sAranPtn = "1" Or sAranPtn = "2" Then
                                    vTemp = CStr(typ_Osfz.CALCMH)                                                            ' 情報３
                                    .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.0")                   ' 情報３
                                End If
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End

    ' OSF，BMD項目追加対応  2002.04.02 yakimura
                                w_1 = IIf(typ_Osfz.OSFRD1 = Null Or typ_Osfz.OSFRD1 = " ", "－", typ_Osfz.OSFRD1)
                                w_2 = IIf(typ_Osfz.OSFRD2 = Null Or typ_Osfz.OSFRD2 = " ", "－", typ_Osfz.OSFRD2)
                                w_3 = IIf(typ_Osfz.OSFRD3 = Null Or typ_Osfz.OSFRD3 = " ", "－", typ_Osfz.OSFRD3)
                                .typ_rslt(UpDo, DispLineCount).INFO4 = w_1 & w_2 & w_3                              ' 情報4
                                
    ' OSF，BMD項目追加対応  2002.04.02 yakimura
                            End If
                        End If
                    End If
                End If
                If bJudg = True Then
                    .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                               ' 判定結果
                Else
                    .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                               ' 判定結果
                    TotalJudg = False
                End If
                DispLineCount = DispLineCount + 1
            End If
            
        '保証ﾌﾗｸﾞ=S又はNULLの場合
        Else
            If shiji Then
                If OsfNo = 4 Then
            
                    '画面表示内容設定
                    .typ_rslt(UpDo, DispLineCount).pos = typ_Osfz.POSITION                      ' 結晶内開始位置
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' 内容
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様無"                             ' 情報１
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                             ' 情報２
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ無"                            ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = typ_Osfz.SMPLNO                     ' サンプルＮｏ
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).OKNG = "N参"                                  ' 判定結果
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' 品番(12桁)
                    If typ_Osfz.SMPLUMU = "0" Then
                        '画面表示内容設定
                        If IsNull(typ_Osfz.OSFRD1) = True Then
                            .typ_rslt(UpDo, DispLineCount).INFO1 = ""
                        Else
                            .typ_rslt(UpDo, DispLineCount).INFO1 = typ_Osfz.OSFRD1
                        End If
                        If IsNull(typ_Osfz.OSFWID1) = True Then
                            .typ_rslt(UpDo, DispLineCount).INFO2 = ""
                        Else
                            .typ_rslt(UpDo, DispLineCount).INFO2 = typ_Osfz.OSFWID1
                        End If
                        If IsNull(typ_Osfz.OSFRD2) = True Then
                            .typ_rslt(UpDo, DispLineCount).INFO3 = ""
                        Else
                            .typ_rslt(UpDo, DispLineCount).INFO3 = typ_Osfz.OSFRD2
                        End If
                        If IsNull(typ_Osfz.OSFWID2) = True Then
                            .typ_rslt(UpDo, DispLineCount).INFO4 = ""
                        Else
                            .typ_rslt(UpDo, DispLineCount).INFO4 = typ_Osfz.OSFWID2
                        End If
                       
                    End If
                Else
            
                    '画面表示内容設定
                    .typ_rslt(UpDo, DispLineCount).pos = typ_Osfz.POSITION                      ' 結晶内開始位置
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' 内容
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様無"                             ' 情報１
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                             ' 情報２
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ無"                            ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = typ_Osfz.SMPLNO                     ' サンプルＮｏ
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).OKNG = "N参"                                  ' 判定結果
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' 品番(12桁)
                    If typ_Osfz.SMPLUMU = "0" Then
                        'OSF判定失敗
                        .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                                            ' 情報２
                        'OSF判定
                        If CrOsfJudg(.typ_si(UpDo), typ_Osfz, bJudg, OsfNo) Then
                            vTemp = CStr(typ_Osfz.CALCAVE)                                                          ' 情報１
                            .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")                   ' 情報１
                            vTemp = CStr(typ_Osfz.CALCMAX)                                                          ' 情報２
                            .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")                    ' 情報２
                            vTemp = CStr(typ_Osfz.MEAS6)                                                            ' 情報３
                            .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")                   ' 情報３
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
                            If sAranPtn = "1" Or sAranPtn = "2" Then
                                vTemp = CStr(typ_Osfz.CALCMH)                                                            ' 情報３
                                .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.0")                   ' 情報３
                            End If
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End

    ' OSF，BMD項目追加対応  2002.04.02 yakimura
                            w_1 = IIf(typ_Osfz.OSFRD1 = Null Or typ_Osfz.OSFRD1 = " ", "－", typ_Osfz.OSFRD1)
                            w_2 = IIf(typ_Osfz.OSFRD2 = Null Or typ_Osfz.OSFRD2 = " ", "－", typ_Osfz.OSFRD2)
                            w_3 = IIf(typ_Osfz.OSFRD3 = Null Or typ_Osfz.OSFRD3 = " ", "－", typ_Osfz.OSFRD3)
                            .typ_rslt(UpDo, DispLineCount).INFO4 = w_1 & w_2 & w_3                                  ' 情報4
    ' OSF，BMD項目追加対応  2002.04.02 yakimura
                        End If
                    End If
                End If
                DispLineCount = DispLineCount + 1
            End If
        End If
    End With
End Sub

Private Function AllHinGdjudg(Gd_si() As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                          typ_gdz As type_DBDRV_scmzc_fcmkc001c_GD, _
                          cnt As Integer) As Boolean
Dim RET     As Boolean
Dim Judg    As Boolean
Dim i       As Integer

    AllHinGdjudg = False
    
    For i = 1 To cnt
        RET = CrGdjudg(Gd_si(i), typ_gdz, Judg)
        If Judg = False Then
            AllHinGdjudg = True
        End If
    Next

End Function

Public Function DBData2DispData(data As Variant, Optional Formatstr As String) As Variant
    If data = -1 Then
        DBData2DispData = ""
    Else
        If Formatstr = "" Then
            DBData2DispData = data
        Else
            DBData2DispData = Format(data, Formatstr)
        End If
    End If
End Function

'------------------------------------------------
' 仕様Nullチェック(結晶)
'------------------------------------------------

'概要      :結晶総合判定の各検査項目の保証方法が'H'または'S'の場合、仕様値がNull(-1)かどうかを判断する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :tSiyou        ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :品番、仕様、結晶内側取得用
'          :sErrMsg       ,IO ,String                               :ｴﾗｰﾒｯｾｰｼﾞ
'          :戻り値        ,O  ,FUNCTION_RETURN                      :結果 = FUNCTION_RETURN_SUCCESS : OK
'                                                                           FUNCTION_RETURN_FAILURE : NG
'説明      :
'履歴      :2003/12/13 新規作成　システムブレイン

Private Function funCryChkNull(tSiyou As type_DBDRV_scmzc_fcmkc001c_Siyou, sErrMsg As String) As FUNCTION_RETURN
    Dim dShiyo()    As Double
    Dim sHosyo      As String
    Dim cnt         As Integer
    
    '初期化
    funCryChkNull = FUNCTION_RETURN_SUCCESS
    
    '--------------- RS(比抵抗) ---------------
    ReDim dShiyo(5)
    dShiyo(1) = tSiyou.HSXRMIN          ' 品ＳＸ比抵抗下限
    dShiyo(2) = tSiyou.HSXRMAX          ' 品ＳＸ比抵抗上限
    dShiyo(3) = tSiyou.HSXRAMIN         ' 品ＳＸ比抵抗平均下限
    dShiyo(4) = tSiyou.HSXRAMAX         ' 品ＳＸ比抵抗平均上限
    dShiyo(5) = tSiyou.HSXRMBNP         ' 品ＳＸ比抵抗面内分布
    If fncJissekiHantei_nl(tSiyou.HSXRHWYS, dShiyo) = False Then
        sErrMsg = sErrMsg & "(RS)"
        funCryChkNull = FUNCTION_RETURN_FAILURE
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00100"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        Exit Function
    End If
    
    '--------------- Oi(酸素濃度) ---------------
    ReDim dShiyo(5)
    dShiyo(1) = tSiyou.HSXONMIN         ' 品ＳＸ酸素濃度下限
    dShiyo(2) = tSiyou.HSXONMAX         ' 品ＳＸ酸素濃度上限
    dShiyo(3) = tSiyou.HSXONAMN         ' 品ＳＸ酸素濃度平均下限
    dShiyo(4) = tSiyou.HSXONAMX         ' 品ＳＸ酸素濃度平均上限
    dShiyo(5) = tSiyou.HSXONMBP         ' 品ＳＸ酸素濃度面内分布
    If fncJissekiHantei_nl(tSiyou.HSXONHWS, dShiyo) = False Then
        sErrMsg = sErrMsg & "(Oi)"
        funCryChkNull = FUNCTION_RETURN_FAILURE
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00101"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        Exit Function
    End If
        
    '--------------- CS(炭素濃度) ---------------
    ReDim dShiyo(2)
    dShiyo(1) = tSiyou.HSXCNMIN         ' 品ＳＸ炭素濃度下限
    dShiyo(2) = tSiyou.HSXCNMAX         ' 品ＳＸ炭素濃度上限
    If fncJissekiHantei_nl(tSiyou.HSXCNHWS, dShiyo) = False Then
        sErrMsg = sErrMsg & "(CS)"
        funCryChkNull = FUNCTION_RETURN_FAILURE
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00111"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        Exit Function
    End If
    
    '--------------- LT(ﾗｲﾌﾀｲﾑ) ---------------
    ReDim dShiyo(1)
'   ReDim dShiyo(2)

    dShiyo(1) = tSiyou.HSXLTMIN         ' 品ＳＸＬタイム下限
'   dShiyo(2) = tSiyou.HSXLTMAX         ' 品ＳＸＬタイム上限
    If fncJissekiHantei_nl(tSiyou.HSXLTHWS, dShiyo) = False Then
        sErrMsg = sErrMsg & "(LT)"
        funCryChkNull = FUNCTION_RETURN_FAILURE
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00110"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        Exit Function
    End If
    
'''Add Start 2011/07/25 LT10Ω判定追加対応 T.Koi(SETsw)
'    dShiyo(1) = tSiyou.HSXLT10MIN         ' 品ＳＸＬLT10下限
'    If fncJissekiHantei_nl(tSiyou.HSXLTHWS, dShiyo) = False Then
'        sErrMsg = sErrMsg & "(LT10)"
'        funCryChkNull = FUNCTION_RETURN_FAILURE
'        gsTbcmy028ErrCode = "00110"
'        Exit Function
'    End If
'
'''Add End   2011/07/25 LT10Ω判定追加対応 T.Koi(SETsw)
    
    '--------------- EPD ---------------
    If tSiyou.EPDUP = -1 Then           ' EPD上限
        sErrMsg = sErrMsg & "(EPD)"
        funCryChkNull = FUNCTION_RETURN_FAILURE
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00102"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        Exit Function
    End If

End Function
'概要      :ブロック偏析、追ﾄﾞｰﾌﾟ位置判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :TopPos        ,I   ,トップｻﾝﾌﾟﾙ位置
'          :BotPos        ,I   ,ボトムｻﾝﾌﾟﾙ位置
'          :TopMeas       ,I   ,トップ中心実測値
'          :BotMeas       ,I   ,ボトム中心実測値
'          :JDCryNum      ,I   ,結晶番号
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :ブロック偏析範囲および追ﾄﾞｰﾌﾟ位置を含むかチェックを行う（推定時限定）
'履歴      :2005/1/11
Public Function HenDopeJudg(TOPPOS As Integer, BOTPOS As Integer, TopMeas As Double, BotMeas As Double, _
                          JDCryNum As String, JDHinb As tFullHinban) As Boolean
    Dim COEF        As Double
    Dim wgtCharge   As Long                 '偏析計算用パラメータ
    Dim wgtTop      As Double               '偏析計算用パラメータ
    Dim wgtTopCut   As Double               '偏析計算用パラメータ
    Dim DM          As Double               '偏析計算用パラメータ
    Dim cf As C_COEF
    Dim sMcno2 As String
    Dim sMcno1 As String
    Dim sMcno  As String
    Dim cc          As type_Coefficient
    Dim ErrInfo     As ERROR_INFOMATION     'エラー情報構造体
    Dim i As Integer
    
    HenDopeJudg = True
    
    '偏析係数計算 マルチ引上対応 参照関数変更 2008/04/23 SETsw Nakada
    If GetCoeffParams_new(JDCryNum, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then
        Debug.Print "偏析計算用パラメータの取得に失敗した"
    End If
    
    cc.DUNMENSEKI = AreaOfCircle(DM)
    cc.TOPSMPLPOS = TOPPOS
    cc.BOTSMPLPOS = BOTPOS
    cc.CHARGEWEIGHT = wgtCharge
    cc.TOPWEIGHT = wgtTop + wgtTopCut
    cc.TOPRES = TopMeas
    cc.BOTRES = BotMeas
    COEF = CoefficientCalculation(cc)
    'ブロック偏析判定処理 -------
    '品番より製作条件ナンバーを求める
    sMcno = Trim(GetMcno(JDHinb))
    i = UBound(SuiteiData)
    SuiteiData(i).SuiSpec.PRODCOND = sMcno
    sMcno1 = Mid(sMcno, 2, 1)
    sMcno2 = Mid(sMcno, 1, 1)
    cf.JudgCOEF = True
    Select Case sMcno1
        Case "H", "I", "J", "K"
            cf.NP = "n"
        Case "A", "B", "C"
            Select Case sMcno2
                Case "A", "B"
                    cf.NP = "p+"
                Case "1", "2", "3", "4", "5", "6", "7", "C", "E"
                    cf.NP = "p-"
                Case Else
                    cf.JudgCOEF = False
            End Select
        Case Else
            cf.JudgCOEF = False
    End Select
    If cf.JudgCOEF Then
        cf.COEF = COEF
        If CrystalCOEFJudg(cf, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
            cf.JudgCOEF = False
        End If
    End If
    With SuiteiData(i)
        'エラー表示用にフラグをセットする,推定不可を返す
        If cf.JudgCOEF Then
            .COEFflg = True
        Else
            .COEFflg = False
            HenDopeJudg = False
        End If
        .Hinsyu = cf.NP
        .COEF = cf.COEF
        '追加ﾄﾞｰﾌﾟ位置のチェック
        .DOPEflg = True
        If typ_b.typ_si(1).ADDDPPOS <> 0 Then
            If TOPPOS <= typ_b.typ_si(1).ADDDPPOS And BOTPOS >= typ_b.typ_si(1).ADDDPPOS Then
               .DOPEflg = False
                HenDopeJudg = False
            End If
        End If
    End With
        
End Function

'概要      :製作条件№を取得する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:jdHinban    ,I  ,String           ,
'      　　:戻り値       ,O  ,Long           　,処理回数
'作成      :2005/1/11
Public Function GetMcno(jd_Hinban As tFullHinban) As String

    Dim sSql As String
    Dim rs As OraDynaset
    
    
    sSql = "SELECT MCNO FROM TBCME036"
    sSql = sSql & " WHERE HINBAN = '" & jd_Hinban.hinban & "' "
    sSql = sSql & " and MNOREVNO = '" & jd_Hinban.mnorevno & "'"
    sSql = sSql & " and FACTORY  = '" & jd_Hinban.factory & "'"
    sSql = sSql & " and OPECOND  = '" & jd_Hinban.opecond & "'"
    
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    
    If rs.RecordCount = 0 Then
        GetMcno = ""
    Else
        GetMcno = rs("MCNO")
    End If
    
End Function

'
'ライフタイムを再計算する
'概要      :LT計算関数を呼び出し値を返す
'ﾊﾟﾗﾒｰﾀ　　:変数名   ,IO ,型                                ,説明
'      　　:Siyou    ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou  ,測定位置を取得
'      　　:jisseki  ,IO ,type_DBDRV_scmzc_fcmkc001c_LT     ,測定値1～10を取得し、計算結果を返す
'      　　:戻り値   なし
'作成      :2005/12/02 SETsw　高崎　伸行
'          :
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Sub_LTReCalc(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                jisseki As type_DBDRV_scmzc_fcmkc001c_LT)

    Dim MEAS(9) As Integer      '測定値格納配列
    Dim iOldFlg As Integer      '旧データ判定フラグ
    Dim iRet As Integer         '戻り値
    Dim iResult As Integer      '計算結果
    
    MEAS(0) = jisseki.MEAS1
    MEAS(1) = jisseki.MEAS2
    MEAS(2) = jisseki.MEAS3
    MEAS(3) = jisseki.MEAS4
    MEAS(4) = jisseki.MEAS5
    MEAS(5) = jisseki.MEAS6
    MEAS(6) = jisseki.MEAS7
    MEAS(7) = jisseki.MEAS8
    MEAS(8) = jisseki.MEAS9
    MEAS(9) = jisseki.MEAS10
    
    '旧データ判定フラグ確認
    If jisseki.LTSPIFLG <> "" Then
        iOldFlg = 0
    Else
        iOldFlg = 1
    End If
    
    'ライフタイム計算値
    iRet = KNS_CalculateMeasResult_LT(iResult, MEAS, siyou.HSXLTSPI, iOldFlg)
    If iRet <> FUNC_RET_LT_SUCCESS Then
        jisseki.CALCMEAS = -1
    Else
        jisseki.CALCMEAS = iResult
    End If
End Sub

'------------------------------------------------
' 複数品番判定対応
'------------------------------------------------

'概要      :実績値の総合判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型             :説明
'          :sKeyID          ,I  ,String         :ﾌﾞﾛｯｸID、又は、結晶番号
'          :tNew_Hinban     ,I  ,tFullHinban    :振替候補品番
'          :bTotalJudg      ,O  ,Boolean        :トータル判定
'          :iErr_Code       ,O  ,Integer        :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String         :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :typ_B           ,O  ,typ_AllTypesB  :全情報構造体(構造体)
'          :iSmpGetFlg      ,I  ,Integer        :ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ(0:ｻﾝﾌﾟﾙ指定なし, 1:ｻﾝﾌﾟﾙ指定あり)
'          :iSamplID1       ,I  ,Long           :TOPｻﾝﾌﾟﾙID(省略可)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :iSamplID2       ,I  ,Long           :BOTｻﾝﾌﾟﾙID(省略可)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :iKcnt           ,I  ,Integer        :工程連番(省略可)
'          :戻り値          ,O  ,Integer        :取得の成否(0:正常終了, -1:異常終了)
'説明      :
'履歴      :複数品番判定対応 20060501 SMP桜井 funCrySogoHanteiより改変
''memo:            tOld_Hinban = TOP tNew_Hinban=Tail
Public Function funCrySogoHantei_CC600Multi(sKeyID As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                bTotalJudg As Boolean, iErr_Code As Integer, sErr_Msg As String, typ_b As typ_AllTypesB, _
                iSmpGetFlg As Integer, Optional iSamplID1 As Long = 0, Optional iSamplID2 As Long = 0, _
                Optional iKcnt As Integer = 0) As Integer
    
    On Error GoTo Apl_down
    Dim liCnt As Integer
    
    '戻り値初期化
    funCrySogoHantei_CC600Multi = FUNCTION_RETURN_FAILURE
    TotalJudg = True
    
    'グローバル変数に設定
    ciSmpGetFlg = iSmpGetFlg
    ciKcnt = iKcnt
    
    'ブロックIDを設定
    sErr_Msg = "結晶総合判定(ﾌﾞﾛｯｸID設定)"
    typ_b.BLOCKID = sKeyID
    
    '画面情報設定
    sErr_Msg = "結晶総合判定(SetAllData)"
    
    If SetAllData2(typ_b, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg, iSmpGetFlg, iSamplID1, iSamplID2) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
  
  
    '仕様検査指示取得
    sErr_Msg = "結晶総合判定(SpecJudgCheck)"
    Call SpecJudgCheck
    
    '2003/12/13 SystemBrain Null対応追加▽
    '仕様Nullチェック
    sErr_Msg = "仕様Nullﾁｪｯｸ"
    If funCryChkNull(typ_b.typ_si(BlkTop), sErr_Msg) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    '2003/12/13 SystemBrain Null対応追加△
    
    '実績データ判定(TOP)
    sErr_Msg = "結晶総合判定(判定(TOP))"
    
    
    
    '----TEST2004/10
    '画面出力用に実測抵抗値を退避しておく
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS1 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS1
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS2 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS2
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS3 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS3
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS4 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS4
    typ_b.typ_zi.CRYRZ(BlkTop).JMEAS5 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS5
        
    If giTpMultiFlg = 0 Then ''<<複数品番判定対応改変部分　一番Topのときさせる
        '--Topのときのtyp_B.typ_zi.CRYRZ(BlkTop).JMEASxxを保持
        Erase pJMEAS_Top
        ReDim pJMEAS_Top(5)
        pJMEAS_Top(1) = typ_b.typ_zi.CRYRZ(BlkTop).MEAS1
        pJMEAS_Top(2) = typ_b.typ_zi.CRYRZ(BlkTop).MEAS2
        pJMEAS_Top(3) = typ_b.typ_zi.CRYRZ(BlkTop).MEAS3
        pJMEAS_Top(4) = typ_b.typ_zi.CRYRZ(BlkTop).MEAS4
        pJMEAS_Top(5) = typ_b.typ_zi.CRYRZ(BlkTop).MEAS5
        psKSTAFFID = typ_b.typ_zi.CRYRZ(BlkTop).KSTAFFID
        psHSXRSPOT = typ_b.typ_si(BlkTop).HSXRSPOT
        psHSXRSPOI = typ_b.typ_si(BlkTop).HSXRSPOI

        If Trim(typ_b.typ_zi.CRYRZ(BlkTop).KSTAFFID) <> KSTAFF_J002 Then
            '抵抗値を測定位置コードにより並べ替える
            ''--<<<<一番Top
            If Set_Rs_Ichi(typ_b.typ_si(BlkTop).HSXRSPOT, typ_b.typ_si(BlkTop).HSXRSPOI, typ_b.typ_zi.CRYRZ(BlkTop).MEAS1, _
                            typ_b.typ_zi.CRYRZ(BlkTop).MEAS2, typ_b.typ_zi.CRYRZ(BlkTop).MEAS3, typ_b.typ_zi.CRYRZ(BlkTop).MEAS4, typ_b.typ_zi.CRYRZ(BlkTop).MEAS5) = FUNCTION_RETURN_FAILURE Then
                GoTo Apl_down
            End If
        End If
    End If
    

    If giBtMultiFlg = 0 And giTpMultiFlg = 1 Then ''<<<<改変部分　一番したのBottomのときさせる
        typ_b.typ_zi.CRYRZ(BlkTop).MEAS1 = pJMEAS_Top(1)
        typ_b.typ_zi.CRYRZ(BlkTop).MEAS2 = pJMEAS_Top(2)
        typ_b.typ_zi.CRYRZ(BlkTop).MEAS3 = pJMEAS_Top(3)
        typ_b.typ_zi.CRYRZ(BlkTop).MEAS4 = pJMEAS_Top(4)
        typ_b.typ_zi.CRYRZ(BlkTop).MEAS5 = pJMEAS_Top(5)

        typ_b.typ_zi.CRYRZ(BlkTop).JMEAS1 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS1
        typ_b.typ_zi.CRYRZ(BlkTop).JMEAS2 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS2
        typ_b.typ_zi.CRYRZ(BlkTop).JMEAS3 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS3
        typ_b.typ_zi.CRYRZ(BlkTop).JMEAS4 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS4
        typ_b.typ_zi.CRYRZ(BlkTop).JMEAS5 = typ_b.typ_zi.CRYRZ(BlkTop).MEAS5

        If Trim(psKSTAFFID) <> KSTAFF_J002 Then
            '抵抗値を測定位置コードにより並べ替える
            ''--<<<<一番Top
            If Set_Rs_Ichi(psHSXRSPOT, psHSXRSPOI, typ_b.typ_zi.CRYRZ(BlkTop).MEAS1, _
                            typ_b.typ_zi.CRYRZ(BlkTop).MEAS2, typ_b.typ_zi.CRYRZ(BlkTop).MEAS3, typ_b.typ_zi.CRYRZ(BlkTop).MEAS4, typ_b.typ_zi.CRYRZ(BlkTop).MEAS5) = FUNCTION_RETURN_FAILURE Then
                GoTo Apl_down
            End If
        End If
    End If
    
    ''-Topの判定>>>>>>>>>>>>>>>>>改変部分
    If giTpMultiFlg = 0 Then '全検査項目
        If CrAllJudg(typ_b, tNew_Hinban, BlkTop) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    ElseIf giTpMultiFlg = 1 Then ''Cs,LT,EPDで合否判定
        If CrAllJudgCC600Multi(typ_b, tOld_Hinban, BlkTop) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    End If
    ''<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    '実績データ判定(TAIL)
    sErr_Msg = "結晶総合判定(判定(TAIL))"
    
    '画面出力用に実測抵抗値を退避しておく
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS1 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS1
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS2 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS2
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS3 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS3
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS4 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS4
    typ_b.typ_zi.CRYRZ(BlkTail).JMEAS5 = typ_b.typ_zi.CRYRZ(BlkTail).MEAS5
        
    If giBtMultiFlg = 0 Then ''<<<<改変部分　一番したのBottomのときさせる
        If Trim(typ_b.typ_zi.CRYRZ(BlkTail).KSTAFFID) <> KSTAFF_J002 Then
            '抵抗値を測定位置コードにより並べ替える
            ''--<<<<一番Bottom
            If Set_Rs_Ichi(typ_b.typ_si(BlkTail).HSXRSPOT, typ_b.typ_si(BlkTail).HSXRSPOI, typ_b.typ_zi.CRYRZ(BlkTail).MEAS1, _
                            typ_b.typ_zi.CRYRZ(BlkTail).MEAS2, typ_b.typ_zi.CRYRZ(BlkTail).MEAS3, typ_b.typ_zi.CRYRZ(BlkTail).MEAS4, typ_b.typ_zi.CRYRZ(BlkTail).MEAS5) = FUNCTION_RETURN_FAILURE Then
                GoTo Apl_down
            End If
        End If
    End If
    ''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>改変部分
    '--Bottomの判定
    If giBtMultiFlg = 0 Then ''全検査項目
        If CrAllJudg(typ_b, tNew_Hinban, BlkTail) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    ElseIf giBtMultiFlg = 1 Then ''Cs,LT,EPDで合否判定
        If CrAllJudgCC600Multi(typ_b, tNew_Hinban, BlkTail) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    End If
    ''<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    bTotalJudg = TotalJudg
    
    funCrySogoHantei_CC600Multi = FUNCTION_RETURN_SUCCESS
'------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funCrySogoHantei_CC600Multi = -4
    iErr_Code = funCrySogoHantei_CC600Multi
    GoTo Apl_Exit
    
End Function

'概要      :引上結晶判定  CC600マルチブロック対応
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               :説明
'          :typ_B         ,I  ,typ_AllTypesB    :各情報構造体
'          :tNew_Hinban   ,I  ,tFullHinban      :振替候補品番
'          :tt            ,I  ,Integer          :TopTail判定用
'説明      :検査指示に従い、Cs,EPD,LTの実績判定を行う
'履歴      : 複数品番判定対応　20060501 SMP桜井  CrAllJudgの改変
'''
Public Function CrAllJudgCC600Multi(typ_b As typ_AllTypesB, tNew_Hinban As tFullHinban, tt As Integer) As FUNCTION_RETURN
    Dim IND         As String                   '検査指示
    Dim bJudg       As Boolean
    Dim i           As Integer
    Dim cnt         As Integer
    Dim typTmList() As typ_TBCMB005
    Dim minwk       As String, maxwk As String
    Dim vTemp       As Variant
    Dim RET         As FUNCTION_RETURN
    Dim Gd_si()     As type_DBDRV_scmzc_fcmkc001c_Siyou
    Dim jCs         As String                               'ブロック内品番のCs保証
    Dim jCsFromTo   As String                               'ブロック内品番のCs保証(FromTo)
    Dim hasSiji     As Boolean                              '検査指示あり
    Dim sHinban12   As String                               '品番(12桁)
    Dim bJudgXY     As Boolean                              'X線判定用フラグ追加 2009/10/22
    Dim bJudgX      As Boolean                              'X線判定用フラグ追加 2009/10/22
    Dim bJudgY      As Boolean                              'X線判定用フラグ追加 2009/10/22
    Dim Oi          As C_Oi       '2010/03/12
    
    CrAllJudgCC600Multi = FUNCTION_RETURN_FAILURE
    
    sHinban12 = tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond
    
    i = 0
       
    '結晶コードリスト取得
    If GetCodeList(MSYSCLASS, KCLASS, typTmList) <> FUNCTION_RETURN_SUCCESS Then
        '結晶コードリスト取得失敗
        Exit Function
    End If
    With typ_b
'>>>>> Oiの追加 2011/02/09 SETsw kubota -------------------------
        '検査指示設定
        IND = IIf(tt = BlkTop, "123", "123")
        '' 結晶検査指示(Oi)*****************************************************************
        If JudgSC_B(tt).Oi Then
            '画面表示内容設定
            .typ_rslt(tt, i).BLOCKNG = False
            .typ_rslt(tt, i).pos = -1                                       ' 結晶内開始位置
            .typ_rslt(tt, i).NAIYO = Search_CrCode("Oi", typTmList())       ' 内容
            .typ_rslt(tt, i).INFO1 = "仕様有"                               ' 情報１
            .typ_rslt(tt, i).INFO2 = "検査無"                               ' 情報２
            .typ_rslt(tt, i).INFO3 = ""                                     ' 情報３
            .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
            .typ_rslt(tt, i).SMPLNO = -1                                    ' サンプルＮｏ
            .typ_rslt(tt, i).OKNG = "NG"                                    ' 判定結果
            .typ_rslt(tt, i).hinban = sHinban12                             ' 品番(12桁)
            bJudg = False
            If (InStr(IND, .typ_cr(tt).CRYINDOICS) <> 0) Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = .typ_zi.OIZ(tt).POSITION                 ' 結晶内開始位置
                .typ_rslt(tt, i).NAIYO = Search_CrCode("Oi", typTmList())       ' 内容
                .typ_rslt(tt, i).SMPLNO = .typ_zi.OIZ(tt).SMPLNO                ' サンプルＮｏ
                .typ_rslt(tt, i).INFO2 = "検査有"                               ' 情報２
                .typ_rslt(tt, i).INFO3 = "実績無"                               ' 情報２
                If left(.typ_zi.OIZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                    '画面表示内容設定
                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ無"                          ' 情報３
                    If .typ_zi.OIZ(tt).SMPLUMU = "0" Then
                        'OI判定失敗
                        .typ_rslt(tt, i).INFO3 = "判定Err"                      ' 情報２
                        'OI判定
                        If CrOiJudg(.typ_si(tt), .typ_zi.OIZ(tt), bJudg) Then
                            Call GetOiMaxMin(.typ_zi.OIZ(tt), minwk, maxwk)
                            '画面表示内容設定
                            vTemp = CStr(.typ_zi.OIZ(tt).OIMEAS1)                       ' 情報１
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' 情報１
                            .typ_rslt(tt, i).INFO2 = DBData2DispData(maxwk, "0.00")     ' 情報２
                            .typ_rslt(tt, i).INFO3 = DBData2DispData(minwk, "0.00")     ' 情報３
                            vTemp = CStr(.typ_zi.OIZ(tt).ORGRES)                        ' 情報４
                            'ORGの小数桁数を6桁(7桁目四捨五入)に変更 2011/11/25 SETsw kubota
                            '.typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")     ' 情報４
                            .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.000000")     ' 情報４
                        Else
                            If .typ_zi.OIZ(tt).ORGRES = -999 Then               ' 2010/03/12 Kameda
                                ReDim Oi.Oi(4)
                                Oi.Oi(0) = .typ_zi.OIZ(tt).OIMEAS1
                                Oi.Oi(1) = .typ_zi.OIZ(tt).OIMEAS2
                                Oi.Oi(2) = .typ_zi.OIZ(tt).OIMEAS3
                                Oi.Oi(3) = .typ_zi.OIZ(tt).OIMEAS4
                                Oi.Oi(4) = .typ_zi.OIZ(tt).OIMEAS5
                                .typ_rslt(tt, i).INFO1 = "仕様" & .typ_si(tt).HSXONSPT & "点"   ' 情報１
                                .typ_rslt(tt, i).INFO2 = "検査" & GetTensu(Oi) & "点"                                ' 情報２
                                .typ_rslt(tt, i).INFO4 = "点数不足"     ' 情報４
                            End If
                        End If
                    End If
                End If
            End If
            
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                               ' 判定結果
            Else
                .typ_rslt(tt, i).OKNG = "NG"                               ' 判定結果
                TotalJudg = False
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00101"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
        Else
            If (InStr(IND, .typ_cr(tt).CRYINDOICS) <> 0) Then
                .typ_rslt(tt, i).BLOCKNG = False
                .typ_rslt(tt, i).OKNG = "OK"                                ' 判定結果
                '画面表示内容設定
                .typ_rslt(tt, i).pos = .typ_zi.OIZ(tt).POSITION             ' 結晶内開始位置
                .typ_rslt(tt, i).NAIYO = Search_CrCode("Oi", typTmList())   ' 内容
                .typ_rslt(tt, i).SMPLNO = .typ_zi.OIZ(tt).SMPLNO            ' サンプルＮｏ
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).OKNG = "N参"                                ' 判定結果
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).INFO1 = "仕様無"                           ' 情報１
                .typ_rslt(tt, i).INFO2 = "検査有"                           ' 情報２
                .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ無"                          ' 情報３
                .typ_rslt(tt, i).INFO4 = ""                                 ' 情報４
                .typ_rslt(tt, i).hinban = sHinban12                         ' 品番(12桁)
                If .typ_zi.OIZ(tt).SMPLUMU = "0" Then
                    'OI判定失敗
                    .typ_rslt(tt, i).INFO3 = "判定Err"                          ' 情報２
                    'OI判定
                    If CrOiJudg(.typ_si(tt), .typ_zi.OIZ(tt), bJudg) Then
                        Call GetOiMaxMin(.typ_zi.OIZ(tt), minwk, maxwk)
                        '画面表示内容設定
                        vTemp = CStr(.typ_zi.OIZ(tt).OIMEAS1)                       ' 情報１
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' 情報１
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(maxwk, "0.00")     ' 情報２
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(minwk, "0.00")     ' 情報３
                        vTemp = CStr(.typ_zi.OIZ(tt).ORGRES)                        ' 情報４
                        'ORGの小数桁数を6桁(7桁目四捨五入)に変更 2011/11/25 SETsw kubota
                        '.typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")     ' 情報４
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.000000")     ' 情報４
                    End If
                End If
                i = i + 1
            End If
        End If
'<<<<< Oiの追加 2011/02/09 SETsw kubota -------------------------
        '' 結晶検査指示(Cs)*****************************************************************
        '検査指示設定
        IND = IIf(tt = BlkTop, "123", "123")
        If JudgSC_B(tt).Cs And (tt = BlkTail Or .typ_si(tt).HSXCNKHI = "6" Or .typ_si(tt).HSXCNKHI = "9") Then  'TOP/BOT保証対応 09/01/08 ooba
            '画面表示内容初期化
            .typ_rslt(tt, i).BLOCKNG = False
            .typ_rslt(tt, i).pos = -1                                   ' 結晶内開始位置
            .typ_rslt(tt, i).NAIYO = Search_CrCode("Cs", typTmList())   ' 内容
            .typ_rslt(tt, i).INFO1 = "仕様有"                           ' 情報１
            .typ_rslt(tt, i).INFO2 = "検査無"                           ' 情報２
            .typ_rslt(tt, i).INFO3 = ""                                 ' 情報３
            .typ_rslt(tt, i).INFO4 = ""                                 ' 情報４
            .typ_rslt(tt, i).SMPLNO = -1                                ' サンプルＮｏ
            .typ_rslt(tt, i).OKNG = "NG"                                ' 判定結果
            .typ_rslt(tt, i).hinban = sHinban12                         ' 品番(12桁)
            bJudg = False
            If (InStr(IND, .typ_cr(tt).CRYINDCSCS) <> 0) Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = .typ_zi.CSZ(tt).POSITION                 ' 結晶内開始位置
                .typ_rslt(tt, i).NAIYO = Search_CrCode("Cs", typTmList())       ' 内容
                .typ_rslt(tt, i).SMPLNO = .typ_zi.CSZ(tt).SMPLNO                ' サンプルＮｏ
                .typ_rslt(tt, i).INFO2 = "検査有"                               ' 情報２
                .typ_rslt(tt, i).INFO3 = "実績無"                               ' 情報２
                If left(.typ_zi.CSZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                    '画面表示内容設定
                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ無"                          ' 情報３
                    If .typ_zi.CSZ(tt).SMPLUMU = "0" Then
                        'Cs判定失敗
                        .typ_rslt(tt, i).INFO3 = "判定Err"                      ' 情報２
                        'CS判定取得
                        If CrCsjudg(.typ_si(tt), .typ_zi.CSZ(tt), bJudg) Then
                            '画面表示内容設定
                            vTemp = CStr(.typ_zi.CSZ(tt).CSMEAS)
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00") ' 情報１
                            .typ_rslt(tt, i).INFO2 = ""                             ' 情報２
                            .typ_rslt(tt, i).INFO3 = ""                             ' 情報３
                        End If
                    End If
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                               ' 判定結果
            Else
                .typ_rslt(tt, i).OKNG = "NG"                               ' 判定結果
                TotalJudg = False
            End If
            i = i + 1
        Else
            If (InStr(IND, .typ_cr(tt).CRYINDCSCS) <> 0) Then
                .typ_rslt(tt, i).BLOCKNG = False
                .typ_rslt(tt, i).OKNG = "OK"                                    ' 判定結果
                '画面表示内容設定
                .typ_rslt(tt, i).pos = .typ_zi.CSZ(tt).POSITION                 ' 結晶内開始位置
                .typ_rslt(tt, i).NAIYO = Search_CrCode("Cs", typTmList())       ' 内容
                .typ_rslt(tt, i).SMPLNO = .typ_zi.CSZ(tt).SMPLNO                ' サンプルＮｏ
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).OKNG = "N参"                                   ' 判定結果
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).INFO1 = "仕様無"                               ' 情報１
                .typ_rslt(tt, i).INFO2 = "検査有"                               ' 情報２
                .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ無"                              ' 情報３
                .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                .typ_rslt(tt, i).hinban = sHinban12                             ' 品番(12桁)
                If .typ_zi.CSZ(tt).SMPLUMU = "0" Then
                    'Cs判定失敗
                    .typ_rslt(tt, i).INFO3 = "判定Err"                      ' 情報２
                    'CS判定取得
                    If CrCsjudg(.typ_si(tt), .typ_zi.CSZ(tt), bJudg) Then
                        '画面表示内容設定
                        vTemp = CStr(.typ_zi.CSZ(tt).CSMEAS)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00") ' 情報１
                        .typ_rslt(tt, i).INFO2 = ""                             ' 情報２
                        .typ_rslt(tt, i).INFO3 = ""                             ' 情報３
                    End If
                End If
                i = i + 1
            End If
        End If
        
        '' 結晶検査指示(T)*****************************************************************
Dim HIN As tFullHinban
Dim LTSPI As String

        If (InStr(IND, .typ_cr(tt).CRYINDTCS) <> 0) Then
            hasSiji = True
        Else
            hasSiji = False
        End If
        bJudg = True                                        '2004/01/15 SystemBrain
        If (JudgSC_B(tt).Lt) And (tt = BlkTail) Then        '2004/01/15 SystemBrain
            bJudg = False                                   '2004/01/15 SystemBrain
        Else                                                '2004/01/15 SystemBrain
            JudgSC_B(tt).Lt = False                         '2004/01/15 SystemBrain
        End If                                              '2004/01/15 SystemBrain
        
        'LTはBot端でブロック全域を判定することになったため、「Top端品番でLT指示があればBotで表示」は不要となった
        If (JudgSC_B(tt).Lt) Or (hasSiji And (tt = BlkTail)) Then '仕様あり or Bot端で検査あり
            .typ_rslt(tt, i).BLOCKNG = False
            
            '画面表示内容初期化
            .typ_rslt(tt, i).pos = .typ_zi.LTZ(tt).POSITION             ' 結晶内開始位置
            .typ_rslt(tt, i).SMPLNO = -1                                ' サンプルＮｏ
            .typ_rslt(tt, i).NAIYO = Search_CrCode("T", typTmList())    ' 内容
            If JudgSC_B(tt).Lt Then
                .typ_rslt(tt, i).INFO1 = "仕様有"                       ' 情報１
            Else
                .typ_rslt(tt, i).INFO1 = "仕様無"
                bJudg = True
            End If
            If hasSiji Then
                .typ_rslt(tt, i).INFO2 = "検査有"                       ' 情報２
            Else
                .typ_rslt(tt, i).INFO2 = "検査無"
            End If
            .typ_rslt(tt, i).INFO3 = "実績無"                           ' 情報３
            .typ_rslt(tt, i).INFO4 = ""                                 ' 情報４
            .typ_rslt(tt, i).hinban = sHinban12                         ' 品番(12桁)
            
            'ライフタイム
            bJudgX = True   '10Ω判定
            '判定と結果登録
            If .typ_zi.LTZ(tt).CRYNUM = .typ_si(1).CRYNUM Then
                .typ_rslt(tt, i).pos = .typ_zi.LTZ(tt).POSITION                 ' 結晶内開始位置
                .typ_rslt(tt, i).SMPLNO = .typ_zi.LTZ(tt).SMPLNO                ' サンプルＮｏ
                If (.typ_zi.LTZ(tt).SMPLUMU <> "0") Then
                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ無"                          ' 情報３
                Else
                    '2005/12/02 add SET高崎 LT計算関数call ->
                    'ライフタイム値を計算しなおす
                    Call Sub_LTReCalc(.typ_si(tt), .typ_zi.LTZ(tt))
                    '2005/12/02 add SET高崎 LT計算関数call <-
                    
                    'LT判定取得
                    If CrLtjudg(.typ_si(tt), .typ_zi.LTZ(tt), bJudg) Then
''Add Start 2011/07/22 LT10Ω判定追加対応 T.Koi(SETsw)
                        If CrLt10judg(.typ_si(tt), .typ_zi.LTZ(tt), .typ_cr(tt), bJudgX) Then
''Add End   2011/07/22 LT10Ω判定追加対応 T.Koi(SETsw)
                            '画面表示内容設定
                            vTemp = CStr(.typ_zi.LTZ(tt).CALCMEAS)                  ' 情報１
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")
                            vTemp = CStr(.typ_zi.LTZ(tt).MEASPEAK)                  ' 情報２
                            .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0")
                            .typ_rslt(tt, i).INFO3 = ""                             ' 情報３
''Add Start 2011/07/22 LT10Ω判定追加対応 T.Koi(SETsw)
                            ' 情報４
                            If .typ_zi.LTZ(tt).CONVAL = (-1) Then
                                .typ_rslt(tt, i).INFO4 = "NULL"
                            Else
                                .typ_rslt(tt, i).INFO4 = CStr(.typ_zi.LTZ(tt).CONVAL)
                            End If
                        Else
                            .typ_rslt(tt, i).INFO3 = "LT10判定Err"                  ' 情報３
                        End If
''Add End   2011/07/22 LT10Ω判定追加対応 T.Koi(SETsw)
                    Else
                        .typ_rslt(tt, i).INFO3 = "判定Err"                      ' 情報３
                    End If
                End If
            Else    '実績なし
                If JudgSC_B(tt).Lt Then bJudg = False
            End If
            
''Add Start 2011/07/25 LT10Ω判定追加対応 T.Koi(SETsw)
            If bJudg = True Then
                If bJudgX = True Then
                    bJudg = True
                Else
                    bJudg = False
                End If
            End If
''Add End   2011/07/25 LT10Ω判定追加対応 T.Koi(SETsw)
            
            If tt = BlkTail Then ''<<Tailのときのみ判定させる
            If (bJudg = False) Then
                .typ_rslt(tt, i).OKNG = "NG"
                TotalJudg = False
'====================== Debug Debug =====================================
            ElseIf .typ_si(tt).HSXLTHWS = "S" Then
                .typ_rslt(tt, i).OKNG = "N参"                            ' 判定結果
'====================== Debug Debug =====================================
            Else
                .typ_rslt(tt, i).OKNG = "OK"                            ' 判定結果
            End If
            End If
            i = i + 1
        End If
        '' 結晶検査指示(EPD)*****************************************************************
        If JudgSC_B(tt).EPD Then
            If tt = BlkTop Then
                .typ_rslt(tt, i).BLOCKNG = False
                If (InStr(IND, .typ_cr(tt).CRYINDEPCS) <> 0) Then
                    '画面表示内容設定
                    .typ_rslt(tt, i).pos = .typ_zi.EPDZ(tt).POSITION                ' 結晶内開始位置
                    .typ_rslt(tt, i).SMPLNO = .typ_zi.EPDZ(tt).SMPLNO               ' サンプルＮｏ
                    .typ_rslt(tt, i).NAIYO = Search_CrCode("EPD", typTmList())      ' 内容
                    .typ_rslt(tt, i).INFO1 = "仕様有"                               ' 情報１
                    .typ_rslt(tt, i).INFO2 = "検査有"                               ' 情報２
                    .typ_rslt(tt, i).INFO3 = "実績無"                               ' 情報３
                    .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                    .typ_rslt(tt, i).hinban = sHinban12                             ' 品番(12桁)
                    bJudg = False
                    If left(.typ_zi.EPDZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ無"                          ' 情報３
                        If .typ_zi.EPDZ(tt).SMPLUMU = "0" Then
                            'EPD判定失敗
                            .typ_rslt(tt, i).INFO3 = "判定Err"                      ' 情報３
                            'EPD判定取得
                            If CrEpdjudg(.typ_si(tt), .typ_zi.EPDZ(tt), bJudg) Then
                                '画面表示内容設定
                                vTemp = CStr(.typ_zi.EPDZ(tt).MEASURE)                  ' 情報１
                                .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")    ' 情報１
                                .typ_rslt(tt, i).INFO2 = ""                             ' 情報２
                                .typ_rslt(tt, i).INFO3 = ""                             ' 情報３
                            End If
                        End If
                    End If
                    .typ_rslt(tt, i).OKNG = "OK"                               ' 判定結果帳尻あわせ
                    i = i + 1
                End If
            Else
                '画面表示内容設定
                .typ_rslt(tt, i).pos = -1           ' 結晶内開始位置
                .typ_rslt(tt, i).NAIYO = Search_CrCode("EPD", typTmList())          ' 内容
                .typ_rslt(tt, i).INFO1 = "仕様有"                                   ' 情報１
                .typ_rslt(tt, i).INFO2 = "検査無"                                   ' 情報２
                .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
                .typ_rslt(tt, i).INFO4 = ""                                         ' 情報４
                .typ_rslt(tt, i).SMPLNO = -1                                        ' サンプルＮｏ
                .typ_rslt(tt, i).hinban = sHinban12                                 ' 品番(12桁)
                bJudg = False
                If (InStr(IND, .typ_cr(tt).CRYINDEPCS) <> 0) Then
                    If left(.typ_zi.EPDZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ無"                              ' 情報３
                        If .typ_zi.EPDZ(tt).SMPLUMU = "0" Then
                            .typ_rslt(tt, i).pos = .typ_zi.EPDZ(tt).POSITION            ' 結晶内開始位置
                            .typ_rslt(tt, i).SMPLNO = .typ_zi.EPDZ(tt).SMPLNO           ' サンプルＮｏ
                            'EPD判定失敗
                            .typ_rslt(tt, i).INFO3 = "判定Err"                          ' 情報３
                            'EPD判定取得
                            If CrEpdjudg(.typ_si(tt), .typ_zi.EPDZ(tt), bJudg) Then
                                '画面表示内容設定
                                vTemp = CStr(.typ_zi.EPDZ(tt).MEASURE)                  ' 情報１
                                .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")    ' 情報１
                                .typ_rslt(tt, i).INFO2 = ""                             ' 情報２
                                .typ_rslt(tt, i).INFO3 = ""                             ' 情報３
                            End If
                        End If
                    End If
                Else
                    If left(.typ_zi.EPDZ(tt).CRYNUM, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(tt, i).pos = .typ_zi.EPDZ(tt).POSITION                ' 結晶内開始位置
                        .typ_rslt(tt, i).SMPLNO = .typ_zi.EPDZ(tt).SMPLNO               ' サンプルＮｏ
                        'EPD判定失敗
                        .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報３
                        'EPD判定取得
                        If CrEpdjudg(.typ_si(tt), .typ_zi.EPDZ(tt), bJudg) Then
                            '画面表示内容設定
                            vTemp = CStr(.typ_zi.EPDZ(tt).MEASURE)          ' 情報１
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' 情報１
                            .typ_rslt(tt, i).INFO2 = ""                                 ' 情報２
                            .typ_rslt(tt, i).INFO3 = ""                                 ' 情報３
                        End If
                    End If
                End If
                If bJudg = True Then
                    .typ_rslt(tt, i).OKNG = "OK"                               ' 判定結果
                Else
                    .typ_rslt(tt, i).OKNG = "NG"                               ' 判定結果
                    TotalJudg = False
                End If
                i = i + 1
            End If
        Else
            If (InStr(IND, .typ_cr(tt).CRYINDEPCS) <> 0) Then
                '画面表示内容設定
                .typ_rslt(tt, i).BLOCKNG = False
                .typ_rslt(tt, i).pos = .typ_zi.EPDZ(tt).POSITION                    ' 結晶内開始位置
                .typ_rslt(tt, i).NAIYO = Search_CrCode("EPD", typTmList())          ' 内容
                .typ_rslt(tt, i).INFO1 = "仕様無"                                   ' 情報１
                .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
                .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ無"                                  ' 情報３
                .typ_rslt(tt, i).INFO4 = ""                                         ' 情報４
                .typ_rslt(tt, i).SMPLNO = .typ_zi.EPDZ(tt).SMPLNO                   ' サンプルＮｏ
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).OKNG = "N参"                                        ' 判定結果
'====================== Debug Debug =====================================
                .typ_rslt(tt, i).hinban = sHinban12                                 ' 品番(12桁)
                If .typ_zi.EPDZ(tt).SMPLUMU = "0" Then
                    'EPD判定失敗
                    .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報３
                    'EPD判定取得
                    If CrEpdjudg(.typ_si(tt), .typ_zi.EPDZ(tt), bJudg) Then
                        '画面表示内容設定
                        vTemp = CStr(.typ_zi.EPDZ(tt).MEASURE)                      ' 情報１
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' 情報１
                        .typ_rslt(tt, i).INFO2 = ""                                 ' 情報２
                        .typ_rslt(tt, i).INFO3 = ""                                 ' 情報３
                    End If
                End If
                i = i + 1
            End If
        End If
        'SIRD判定データ設定   2010/02/04 add Kameda
        If tt = BlkTop Then
            .typ_rslt(tt, i).BLOCKNG = False
            If .typ_cr(tt).SIRDKBNY3 = "1" Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = .typ_zi.SIRD.POSITION                ' 結晶内開始位置
                '.typ_rslt(tt, i).SMPLNO = .typ_zi.SIRD.SMPLNO               ' サンプルＮｏ
                .typ_rslt(tt, i).NAIYO = Search_CrCode("SD", typTmList())       ' 内容
                .typ_rslt(tt, i).hinban = sHinban12                             ' 品番(12桁)
                bJudg = False
                'SIRD判定取得
                If CrSIRDjudg(.typ_si(tt), .typ_zi.SIRD, bJudg) Then
                    '画面表示内容設定
                    vTemp = CStr(.typ_zi.SIRD.SIRDCNT)                  ' 情報１
                    .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")    ' 情報１
                    .typ_rslt(tt, i).INFO2 = ""                             ' 情報２
                    .typ_rslt(tt, i).INFO3 = ""                             ' 情報３
                End If
                If bJudg = True Then
                    .typ_rslt(tt, i).OKNG = "OK"                               ' 判定結果
                Else
                    .typ_rslt(tt, i).OKNG = "NG"                               ' 判定結果
                    TotalJudg = False
                    ''gsTbcmy028ErrCode = ""
                End If
                '評価待ち参照時  2010/02/18 Kameda
                If .typ_zi.SIRD.NothingFlg = "1" Then
                    .typ_rslt(tt, i).INFO1 = ""                                ' 情報１
                    .typ_rslt(tt, i).OKNG = "評価待ち"                         ' 判定結果
                End If
                i = i + 1
            ElseIf .typ_cr(tt).SIRDKBNY3 = "2" Then       '2010/02/16 add Kameda
                '画面表示内容設定
                .typ_rslt(tt, i).pos = .typ_zi.SIRD.POSITION                 ' 結晶内開始位置
                '.typ_rslt(tt, i).SMPLNO = .typ_zi.SIRD.SMPLNO               ' サンプルＮｏ
                .typ_rslt(tt, i).NAIYO = Search_CrCode("SD", typTmList())       ' 内容
                .typ_rslt(tt, i).hinban = sHinban12                             ' 品番(12桁)
                'bJudg = False    表示のみ
                'SIRD表示
                '画面表示内容設定
                .typ_rslt(tt, i).INFO1 = "先行評価"                     ' 情報１
                .typ_rslt(tt, i).INFO2 = ""                             ' 情報２
                .typ_rslt(tt, i).INFO3 = ""                             ' 情報３
                .typ_rslt(tt, i).OKNG = "OK"                            ' 判定結果
                i = i + 1
            End If
        End If
        
        'X線判定データ設定   2009/08/12 add Kameda
        '合成角のみで判定 X,Yは警告を出す(背景赤）  2009/10/22 add Kameda
        If tt = BlkTail Then
            If .typ_cr(tt).CRYINDXC1 <> 0 Then
                'If CrXjudg(.typ_si(tt), .typ_zi.XZ, bJudg) Then     2009/10/22 Kameda
                If CrXjudg(.typ_si(tt), .typ_zi.XZ, bJudgXY, bJudgX, bJudgY) Then
                    If bJudgXY Then
                        '.typ_zi.XZ.JUDG = "OK"    2009/10/22
                        .typ_zi.XZ.JUDGXY = "OK"
                    Else
                        '.typ_zi.XZ.JUDG = "NG"    2009/10/22
                        .typ_zi.XZ.JUDGXY = "NG"
                        TotalJudg = False
                    End If
                    '警告を出すために項目追加     2009/10/22 Kameda
                    If bJudgX Then
                        .typ_zi.XZ.JUDGX = "OK"
                    Else
                        .typ_zi.XZ.JUDGX = "NG"
                    End If
                    If bJudgY Then
                        .typ_zi.XZ.JUDGY = "OK"
                    Else
                        .typ_zi.XZ.JUDGY = "NG"
                    End If
                End If
            Else
                '.typ_zi.XZ.JUDG = ""     2009/10/22
                .typ_zi.XZ.JUDGXY = ""
                .typ_zi.XZ.JUDGX = ""
                .typ_zi.XZ.JUDGY = ""
            End If
        End If
        
      'Add Start 2011/02/01 SMPK A.Nagamine     : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の実績判定処理
        Call CuDecoDataSet_C(JudgSC_B, typ_b, tt, i, typTmList(), sHinban12)
        
        Call CuDecoDataSet_CJ(JudgSC_B, typ_b, tt, i, typTmList(), sHinban12)
        
        Call CuDecoDataSet_CJLT(JudgSC_B, typ_b, tt, i, typTmList(), sHinban12)
        
        Call CuDecoDataSet_CJ2(JudgSC_B, typ_b, tt, i, typTmList(), sHinban12)
        
      ''Add End   2011/02/01 SMPK A.Nagamine
        
    End With
        
    CrAllJudgCC600Multi = FUNCTION_RETURN_SUCCESS
End Function
'概要      :測定点数を取得する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:OI           ,I  ,           ,
'      　　:戻り値       ,O  ,Long           　,測定点数
'作成      :2010/3/12 Kameda
Public Function GetTensu(Oi As C_Oi) As Integer

    Dim i As Integer
        GetTensu = 5
        For i = 0 To 4
            If Oi.Oi(i) = -1 Then
                GetTensu = i
                Exit Function
            End If
        Next
End Function

'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : Cu-deco(C)の実績判定処理
Public Function CuDecoDataSet_C(pJudgSC_B() As Judg_Spec_Cry, ptyp_b As typ_AllTypesB, UpDo As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String, Optional pblnFlag As Boolean = False) As Boolean
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 UPD By Systech End
    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4                  '検査指示
    Dim bJudg           As Boolean
    Dim strXTALC1       As String
    Dim strJDGEIDC      As String
    Dim strSYNFLG       As String
    Dim strYMKFLG       As String
    Dim lngSMPPOS       As Long
    Dim iErrFlg         As Integer
    Dim YFlg            As Boolean
    Dim intRet          As Integer
    Dim sResult         As String
    Dim strSampUmu      As String
    Dim strInfo(4)      As String
    Dim intSiyou        As Integer
    Dim intJisseki      As Integer
    Dim StrCryNum       As String
    Dim lngSampNo       As Long
    Dim strPtnJsk       As String
    Dim str028ErrCode   As String
    Dim blnRet          As Boolean
    Dim typ_Ret_CuDeco As typ_SB_com_xodb5_osf31_Cudeco
    Dim intNG_Num       As Integer
    
    blnRet = False
    intNG_Num = 0
    
    '検査指示設定
    IND = IIf(UpDo = BlkTop, "123", "123")
    
    With ptyp_b
        JudgSpecCode = pJudgSC_B(UpDo).CuC
        SCC = "C"
        shiji = (InStr(IND, .typ_cr(UpDo).CRYINDCCS) <> 0)
        strSampUmu = .typ_zi.CuC(UpDo).SMPLUMUC                 ' C.サンプル有無 TBCMJ023.SMPLUMUC "0"=サンプル有り,"1"=サンプル無し
        strInfo(0) = CStr(.typ_zi.CuC(UpDo).CDISKJSK)           ' 情報１ C.Disk半径
        strInfo(1) = CStr(.typ_zi.CuC(UpDo).CRINGNKJSK)         ' 情報２ C.Ring内径
        strInfo(2) = CStr(.typ_zi.CuC(UpDo).CRINGGKJSK)         ' 情報３ C.Ring外径
        strInfo(3) = .typ_zi.CuC(UpDo).CPTNJSK                  ' 情報４ C.パターン実績
        lngSMPPOS = .typ_zi.CuC(UpDo).POSITION                  ' C.ｻﾝﾌﾟﾙ位置
        StrCryNum = .typ_zi.CuC(UpDo).CRYNUM                    ' C.結晶番号
        lngSampNo = .typ_zi.CuC(UpDo).SMPLNO                    ' C.サンプル№(代表サンプルID)
        strPtnJsk = .typ_zi.CuC(UpDo).CPTNJSK                   ' C.パターン実績
        str028ErrCode = "00155"
        
        '保証ﾌﾗｸﾞ="H"の場合か第７パラメータOptional がTrue指定されたとき
        If (JudgSpecCode) Or (pblnFlag) Then
            
            bJudg = False
            
            '判定ﾌﾗｸﾞ初期化
            iErrFlg = 0
            YFlg = False
            
            'ﾌﾞﾛｯｸID
            strXTALC1 = Trim(ptyp_b.BLOCKID)
            '結晶番号
            strXTALC1 = left(strXTALC1, 9) & "000"
            
            '結晶番号をｷｰとしてXSDC1よりC－OSF3判定IDを獲得する
            If GetCOSF3ID(strJDGEIDC, strXTALC1) <> FUNCTION_RETURN_SUCCESS Then
                iErrFlg = 11
            Else
                If Trim(strJDGEIDC) = "" Then
                    iErrFlg = 12
'                '判定ID=「9」の場合は判定なし(判定OK)　07/08/01 M.Kaga
'                ElseIf Trim(strJDGEIDC) = "9" Then
'                    YFlg = True
'                    bJudg = True
                Else
                    '獲得したC-OSF3判定IDでXODC5_OSF30より承認ﾌﾗｸﾞの獲得
                    If GetSYNFLAGC5(strSYNFLG, strYMKFLG, strJDGEIDC) <> FUNCTION_RETURN_SUCCESS Then
                        iErrFlg = 13
                    Else
                        '承認ﾌﾗｸﾞ:0　未承認の場合
                        If Trim(strSYNFLG) = "0" Or Trim(strSYNFLG) = "" Or IsNull(strSYNFLG) Then
                            iErrFlg = 14
                        End If
                    End If
                End If
            End If
            
            If iErrFlg > 0 Then
                '画面表示内容設定
                .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' 結晶内開始位置
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' 内容
                .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                         ' 情報１
                .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                         ' 情報２
                .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                        ' 情報３
                .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' サンプルＮｏ
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' 品番(12桁)
                If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                    .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS           ' 結晶内開始位置
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo          ' サンプルＮｏ
                End If
                bJudg = False
            Else
                If YFlg = False Then
                    
                    If GetOsf31_CuDeco(strJDGEIDC, lngSMPPOS, typ_Ret_CuDeco) <> FUNCTION_RETURN_SUCCESS Then
                        iErrFlg = 21
                    Else
                        intSiyou = -1
                        intJisseki = -1
                        
                        If (IsNumeric(.typ_si(UpDo).HSXCPK)) Then
                            intSiyou = CInt(.typ_si(UpDo).HSXCPK)
                        End If
                        
                        If (IsNumeric(strPtnJsk)) Then
                            intJisseki = CInt(strPtnJsk) + 1
                        End If
                        
                        If (intSiyou >= 1) And (intSiyou <= 4) And (intJisseki >= 1) And (intJisseki <= 4) Then
                        
                           'intRet = funCodeDBGet(SYSCLASS, CLASS, 製品仕様パターン区分, 1, 実績パターン区分, 戻り値(tbcmb005.info1))
                            intRet = funCodeDBGet("SB", "S1", CStr(intSiyou), 1, CStr(intJisseki), sResult)
                            If (intRet = 0) And (sResult <> vbNullString) And (Len(sResult) >= 1) Then
                                If sResult = "1" Then
                                    bJudg = True
                                Else
                                    bJudg = False
                                    intNG_Num = 51
                                End If
                            Else
                                'sErr_Msg = sAdd_Msg & sErr_Msg & "→仕様:" & CStr(intSiyou) & ", 実績:" & CStr(intJisseki)
                                bJudg = False
                                intNG_Num = 52
'                                GoTo CodeDBGet_Error
                            End If
                            
                        End If
                    End If
                End If
                
                If iErrFlg > 0 Then
                    '画面表示内容設定
                    .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' 結晶内開始位置
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' 内容
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                         ' 情報１
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                         ' 情報２
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                        ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = CInt(iErrFlg)                    ' 情報４
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' サンプルＮｏ
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' 品番(12桁)
                    If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS              ' 結晶内開始位置
                        .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo             ' サンプルＮｏ
                    End If
                    bJudg = False
                    
                Else
                    '画面表示内容設定
                    .typ_rslt(UpDo, DispLineCount).pos = -1                                     ' 結晶内開始位置
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' 内容
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                             ' 情報１
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "検査無"                             ' 情報２
                    .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                   ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                                  ' サンプルＮｏ
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' 品番(12桁)
                    
                    If shiji Then
                        '画面表示内容設定
                        .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                         ' 情報１
                        .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                         ' 情報２
                        .typ_rslt(UpDo, DispLineCount).INFO3 = "実績無"                         ' 情報３
                        
                        If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                            .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS              ' 結晶内開始位置
                            .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo             ' サンプルＮｏ
                            .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ無"                    ' 情報３
                            
                            If strSampUmu = "0" Then
                                .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                ' 情報３
                                '画面表示内容設定
                                .typ_rslt(UpDo, DispLineCount).INFO1 = strInfo(0)
                                .typ_rslt(UpDo, DispLineCount).INFO2 = strInfo(1)
                                .typ_rslt(UpDo, DispLineCount).INFO3 = strInfo(2)
                                .typ_rslt(UpDo, DispLineCount).INFO4 = strInfo(3)
                            End If
                        End If
                    End If
                End If
            End If
            
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                               ' 判定結果
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                               ' 判定結果
                TotalJudg = False
                gsTbcmy028ErrCode = str028ErrCode
            End If
            DispLineCount = DispLineCount + 1
            
        Else
        ' Add Start 2011/02/15 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : 「保証方法_処」参考時の処理追加
            If shiji Then
                    '画面表示内容設定
                    .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS                              ' 結晶内開始位置
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' 内容
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様無"                             ' 情報１
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                             ' 情報２
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ無"                            ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo                           ' サンプルＮｏ
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).OKNG = "N参"                                 ' 判定結果
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' 品番(12桁)
                    
                    If strSampUmu = "0" Then
                        '画面表示内容設定
                        .typ_rslt(UpDo, DispLineCount).INFO1 = strInfo(0)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = strInfo(1)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = strInfo(2)
                        .typ_rslt(UpDo, DispLineCount).INFO4 = strInfo(3)
                    End If
                DispLineCount = DispLineCount + 1
            End If
        ' Add End   2011/02/15 SMPK A.Nagamine
        End If      '/*  End of If (JudgSpecCode) Or (pblnFlag) Then */
    End With
    
    CuDecoDataSet_C = blnRet
    
End Function
''Add End   2011/01/17 SMPK A.Nagamine


'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : Cu-deco(CJ)の実績判定処理
Public Function CuDecoDataSet_CJ(pJudgSC_B() As Judg_Spec_Cry, ptyp_b As typ_AllTypesB, UpDo As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String, Optional pblnFlag As Boolean = False) As Boolean
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 UPD By Systech End
    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4                  '検査指示
    Dim bJudg           As Boolean
    Dim strXTALC1       As String
    Dim strJDGEIDC      As String
    Dim strSYNFLG       As String
    Dim strYMKFLG       As String
    Dim lngSMPPOS       As Long
    Dim iErrFlg         As Integer
    Dim YFlg            As Boolean
    Dim intRet          As Integer
    Dim sResult         As String
    Dim strSampUmu      As String
    Dim strInfo(4)      As String
    Dim intSiyou        As Integer
    Dim intJisseki      As Integer
    Dim StrCryNum       As String
    Dim lngSampNo       As Long
    Dim strPtnJsk       As String
    Dim str028ErrCode   As String
    Dim blnRet          As Boolean
    Dim typ_Ret_CuDeco  As typ_SB_com_xodb5_osf31_Cudeco
    Dim intNG_Num       As Integer
    Dim intMax          As Integer
    
    blnRet = False
    intNG_Num = 0
    
    '検査指示設定
    IND = IIf(UpDo = BlkTop, "123", "123")
    
    With ptyp_b
        ' CJ 実績
        JudgSpecCode = pJudgSC_B(UpDo).CuCJ
        SCC = "CJ"
        shiji = (InStr(IND, .typ_cr(UpDo).CRYINDCJCS) <> 0)
        strSampUmu = .typ_zi.CuCJ(UpDo).SMPLUMUCJ               ' CJ.サンプル有無 TBCMJ023.SMPLUMUCJ "0"=サンプル有り,"1"=サンプル無し
        strInfo(0) = CStr(.typ_zi.CuCJ(UpDo).CJDISKJSK)         ' 情報１ CJ.Disk半径
        strInfo(1) = CStr(.typ_zi.CuCJ(UpDo).CJRINGNKJSK)       ' 情報２ CJ.Ring内径
        strInfo(2) = CStr(.typ_zi.CuCJ(UpDo).CJRINGGKJSK)       ' 情報３ CJ.Ring外径
        strInfo(3) = .typ_zi.CuCJ(UpDo).CJPTNJSK                ' 情報４ CJ.パターン実績
        lngSMPPOS = .typ_zi.CuCJ(UpDo).POSITION                 ' CJ.ｻﾝﾌﾟﾙ位置
        StrCryNum = .typ_zi.CuCJ(UpDo).CRYNUM                   ' CJ.結晶番号
        lngSampNo = .typ_zi.CuCJ(UpDo).SMPLNO                   ' CJ.サンプル№(代表サンプルID)
        strPtnJsk = .typ_zi.CuCJ(UpDo).CJPTNJSK                 ' CJ.パターン実績
        str028ErrCode = "00156"
        
        '保証ﾌﾗｸﾞ="H"の場合か第７パラメータOptional がTrue指定されたとき
        If (JudgSpecCode) Or (pblnFlag) Then
            
            bJudg = False
            
            '判定ﾌﾗｸﾞ初期化
            iErrFlg = 0
            YFlg = False
            
            'ﾌﾞﾛｯｸID
            strXTALC1 = Trim(ptyp_b.BLOCKID)
            '結晶番号
            strXTALC1 = left(strXTALC1, 9) & "000"
            
            '結晶番号をｷｰとしてXSDC1よりC－OSF3判定IDを獲得する
            If GetCOSF3ID(strJDGEIDC, strXTALC1) <> FUNCTION_RETURN_SUCCESS Then
                iErrFlg = 11
            Else
                If Trim(strJDGEIDC) = "" Then
                    iErrFlg = 12
'                '判定ID=「9」の場合は判定なし(判定OK)　07/08/01 M.Kaga
'                ElseIf Trim(strJDGEIDC) = "9" Then
'                    YFlg = True
'                    bJudg = True
                Else
                    '獲得したC-OSF3判定IDでXODC5_OSF30より承認ﾌﾗｸﾞの獲得
                    If GetSYNFLAGC5(strSYNFLG, strYMKFLG, strJDGEIDC) <> FUNCTION_RETURN_SUCCESS Then
                        iErrFlg = 13
                    Else
                        '承認ﾌﾗｸﾞ:0　未承認の場合
                        If Trim(strSYNFLG) = "0" Or Trim(strSYNFLG) = "" Or IsNull(strSYNFLG) Then
                            iErrFlg = 14
                        End If
                    End If
                End If
            End If
            
            If iErrFlg > 0 Then
                '画面表示内容設定
                .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' 結晶内開始位置
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' 内容
                .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                         ' 情報１
                .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                         ' 情報２
                .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                        ' 情報３
                .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' サンプルＮｏ
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' 品番(12桁)
                If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                    .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS           ' 結晶内開始位置
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo          ' サンプルＮｏ
                End If
                bJudg = False
            Else
                If YFlg = False Then
                    
                    If GetOsf31_CuDeco(strJDGEIDC, lngSMPPOS, typ_Ret_CuDeco) <> FUNCTION_RETURN_SUCCESS Then
                        iErrFlg = 21
                    Else
                        intSiyou = -1
                        intJisseki = -1
                        
                        If (IsNumeric(.typ_si(UpDo).HSXCJPK)) Then
                            intSiyou = CInt(.typ_si(UpDo).HSXCJPK)
                        End If
                        
                        If (IsNumeric(strPtnJsk)) Then
                            intJisseki = CInt(strPtnJsk) + 1
                        End If
                        
                        If (intSiyou >= 1) And (intSiyou <= 4) And (intJisseki >= 1) And (intJisseki <= 4) Then
                        
                           'intRet = funCodeDBGet(SYSCLASS, CLASS, 製品仕様パターン区分, 1, 実績パターン区分, 戻り値(tbcmb005.info1))
                            intRet = funCodeDBGet("SB", "S1", CStr(intSiyou), 1, CStr(intJisseki), sResult)
                            If (intRet = 0) And (sResult <> vbNullString) And (Len(sResult) >= 1) Then
                                If sResult = "1" Then
                                    bJudg = True
                                Else
                                    bJudg = False
                                    intNG_Num = 51
                                End If
                            Else
                                'sErr_Msg = sAdd_Msg & sErr_Msg & "→仕様:" & CStr(intSiyou) & ", 実績:" & CStr(intJisseki)
                                bJudg = False
                                intNG_Num = 52
'                                GoTo CodeDBGet_Error
                            End If
                            
                        End If
                        
                        ' CJ Ring内径・外径の判定
                        If bJudg Then
                            If (strPtnJsk = CNST_JSK_PTN_Ring) Or (strPtnJsk = CNST_JSK_PTN_DiskRing) Then
                                If (typ_Ret_CuDeco.CJALLMINRINC5 = -1) Or (typ_Ret_CuDeco.CJALLMINRINC5 > 150) Then
                                    bJudg = False
                                    intNG_Num = 61
                                ElseIf (.typ_zi.CuCJ(UpDo).CJRINGNKJSK = -1) Or (.typ_zi.CuCJ(UpDo).CJRINGNKJSK > 150) Then
                                    bJudg = False
                                    intNG_Num = 62
                                ElseIf (typ_Ret_CuDeco.CJALLMAXRIGC5 = -1) Or (typ_Ret_CuDeco.CJALLMAXRIGC5 > 150) Then
                                    bJudg = False
                                    intNG_Num = 63
                                ElseIf (.typ_zi.CuCJ(UpDo).CJRINGGKJSK = -1) Or (.typ_zi.CuCJ(UpDo).CJRINGGKJSK > 150) Then
                                    bJudg = False
                                    intNG_Num = 64
                                ElseIf (typ_Ret_CuDeco.CJALLMINRINC5 > .typ_zi.CuCJ(UpDo).CJRINGNKJSK) Then
                                    bJudg = False
                                    intNG_Num = 65
                                ElseIf (typ_Ret_CuDeco.CJALLMAXRIGC5 < .typ_zi.CuCJ(UpDo).CJRINGGKJSK) Then
                                    bJudg = False
                                    intNG_Num = 66
                                End If
                            End If
                        End If
                        
                        ' CJ Disk半径の判定
                        If bJudg Then
                            If (strPtnJsk = CNST_JSK_PTN_Disk) Or (strPtnJsk = CNST_JSK_PTN_DiskRing) Then
                                If (typ_Ret_CuDeco.CJALLMAXDIC5 = -1) Or (typ_Ret_CuDeco.CJALLMAXDIC5 > 150) Then
                                    bJudg = False
                                    intNG_Num = 61
                                ElseIf (.typ_zi.CuCJ(UpDo).CJDISKJSK = -1) Or (.typ_zi.CuCJ(UpDo).CJDISKJSK > 150) Then
                                    bJudg = False
                                    intNG_Num = 62
                                ElseIf (typ_Ret_CuDeco.CJALLMAXDIC5 < .typ_zi.CuCJ(UpDo).CJDISKJSK) Then
                                    bJudg = False
                                    intNG_Num = 71
                                End If
                            End If
                        End If
                        
                        'CJ 計算Pi幅の判定(上限値チェック)
                        If bJudg Then
                            If (strPtnJsk = CNST_JSK_PTN_Disk) Or (strPtnJsk = CNST_JSK_PTN_Ring) Or (strPtnJsk = CNST_JSK_PTN_DiskRing) Then
                                If (strPtnJsk = CNST_JSK_PTN_Disk) Then
                                    intMax = typ_Ret_CuDeco.CJDMAXPIC5
                                ElseIf (strPtnJsk = CNST_JSK_PTN_Ring) Then
                                    intMax = typ_Ret_CuDeco.CJRMAXPIC5
                                Else
                                    intMax = typ_Ret_CuDeco.CJDRMAXPIC5
                                End If
                                
                                If (intMax = -1) Or (intMax > 150) Then
                                    bJudg = False
                                    intNG_Num = 81
                                ElseIf (.typ_zi.CuCJ(UpDo).CJPICALC = -1) Or (.typ_zi.CuCJ(UpDo).CJPICALC > 150) Then
                                    bJudg = False
                                    intNG_Num = 82
                                ElseIf (intMax < .typ_zi.CuCJ(UpDo).CJPICALC) Then
                                    bJudg = False
                                    intNG_Num = 83
                                End If
                            End If
                        End If
                        
                    End If
                End If
                
                
                If iErrFlg > 0 Then
                    '画面表示内容設定
                    .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' 結晶内開始位置
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' 内容
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                         ' 情報１
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                         ' 情報２
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                        ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = CInt(iErrFlg)                    ' 情報４
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' サンプルＮｏ
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' 品番(12桁)
                    If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS              ' 結晶内開始位置
                        .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo             ' サンプルＮｏ
                    End If
                    bJudg = False
                    
                Else
                    '画面表示内容設定
                    .typ_rslt(UpDo, DispLineCount).pos = -1                                     ' 結晶内開始位置
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' 内容
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                             ' 情報１
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "検査無"                             ' 情報２
                    .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                   ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                                  ' サンプルＮｏ
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' 品番(12桁)
                    
                    If shiji Then
                        '画面表示内容設定
                        .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                         ' 情報１
                        .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                         ' 情報２
                        .typ_rslt(UpDo, DispLineCount).INFO3 = "実績無"                         ' 情報３
                        
                        If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                            .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS              ' 結晶内開始位置
                            .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo             ' サンプルＮｏ
                            .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ無"                    ' 情報３
                            
                            If strSampUmu = "0" Then
                                .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                ' 情報３
                                '画面表示内容設定
                                .typ_rslt(UpDo, DispLineCount).INFO1 = strInfo(0)
                                .typ_rslt(UpDo, DispLineCount).INFO2 = strInfo(1)
                                .typ_rslt(UpDo, DispLineCount).INFO3 = strInfo(2)
                                .typ_rslt(UpDo, DispLineCount).INFO4 = strInfo(3)
                            End If
                        End If
                    End If
                End If
            End If
            
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                               ' 判定結果
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                               ' 判定結果
                TotalJudg = False
                gsTbcmy028ErrCode = str028ErrCode
            End If
            DispLineCount = DispLineCount + 1
            
        Else
        ' Add Start 2011/02/15 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : 「保証方法_処」参考時の処理追加
            If shiji Then
                    '画面表示内容設定
                    .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS                              ' 結晶内開始位置
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' 内容
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様無"                             ' 情報１
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                             ' 情報２
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ無"                            ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo                           ' サンプルＮｏ
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).OKNG = "N参"                                 ' 判定結果
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' 品番(12桁)
                    
                    If strSampUmu = "0" Then
                        '画面表示内容設定
                        .typ_rslt(UpDo, DispLineCount).INFO1 = strInfo(0)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = strInfo(1)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = strInfo(2)
                        .typ_rslt(UpDo, DispLineCount).INFO4 = strInfo(3)
                    End If
                DispLineCount = DispLineCount + 1
            End If
        ' Add End   2011/02/15 SMPK A.Nagamine
        End If      '/*  End of If (JudgSpecCode) Or (pblnFlag) Then */
    End With
    
    CuDecoDataSet_CJ = blnRet
    
End Function
''Add End   2011/01/17 SMPK A.Nagamine

'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : Cu-deco(CJ(LT))の実績判定処理
Public Function CuDecoDataSet_CJLT(pJudgSC_B() As Judg_Spec_Cry, ptyp_b As typ_AllTypesB, UpDo As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String, Optional pblnFlag As Boolean = False) As Boolean
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 UPD By Systech End
    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4                  '検査指示
    Dim bJudg           As Boolean
    Dim strXTALC1       As String
    Dim strJDGEIDC      As String
    Dim strSYNFLG       As String
    Dim strYMKFLG       As String
    Dim lngSMPPOS       As Long
    Dim iErrFlg         As Integer
    Dim YFlg            As Boolean
    Dim intRet          As Integer
    Dim sResult         As String
    Dim strSampUmu      As String
    Dim strInfo(4)      As String
    Dim intSiyou        As Integer
    Dim intJisseki      As Integer
    Dim StrCryNum       As String
    Dim lngSampNo       As Long
    Dim strPtnJsk       As String
    Dim str028ErrCode   As String
    Dim blnRet          As Boolean
    Dim typ_Ret_CuDeco  As typ_SB_com_xodb5_osf31_Cudeco
    Dim intNG_Num       As Integer
    
    blnRet = False
    intNG_Num = 0
    
    '検査指示設定
    IND = IIf(UpDo = BlkTop, "123", "123")
    
    With ptyp_b
        ' CJ(LT) 実績
        JudgSpecCode = pJudgSC_B(UpDo).CuCJLT
        SCC = "CJLT"
        shiji = (InStr(IND, .typ_cr(UpDo).CRYINDCJLTCS) <> 0)
        strSampUmu = .typ_zi.CuCJLT(UpDo).SMPLUMUCJLT           ' CJ(LT).サンプル有無 TBCMJ023.SMPLUMUCJLT "0"=サンプル有り,"1"=サンプル無し
        strInfo(0) = CStr(.typ_zi.CuCJLT(UpDo).CJLTPICALC)      ' 情報１ CJ(LT).Pi幅計算
        strInfo(1) = CStr(.typ_zi.CuCJLT(UpDo).CJLTBANDNKJSK)   ' 情報２ CJ(LT).Band内径実績
        strInfo(2) = CStr(.typ_zi.CuCJLT(UpDo).CJLTBANDGKJSK)   ' 情報３ CJ(LT).Band外径実績
        strInfo(3) = .typ_zi.CuCJLT(UpDo).CJLTPTNJSK            ' 情報４ CJ(LT).パターン実績
        lngSMPPOS = .typ_zi.CuCJLT(UpDo).POSITION               ' CJ(LT).ｻﾝﾌﾟﾙ位置
        StrCryNum = .typ_zi.CuCJLT(UpDo).CRYNUM                 ' CJ(LT).結晶番号
        lngSampNo = .typ_zi.CuCJLT(UpDo).SMPLNO                 ' CJ(LT).サンプル№(代表サンプルID)
        strPtnJsk = .typ_zi.CuCJLT(UpDo).CJLTPTNJSK             ' CJ(LT).パターン実績
        str028ErrCode = "00157"
        
        '保証ﾌﾗｸﾞ="H"の場合か第７パラメータOptional がTrue指定されたとき
        If (JudgSpecCode) Or (pblnFlag) Then
            
            bJudg = False
            
            '判定ﾌﾗｸﾞ初期化
            iErrFlg = 0
            YFlg = False
            
            'ﾌﾞﾛｯｸID
            strXTALC1 = Trim(ptyp_b.BLOCKID)
            '結晶番号
            strXTALC1 = left(strXTALC1, 9) & "000"
            
            '結晶番号をｷｰとしてXSDC1よりC－OSF3判定IDを獲得する
            If GetCOSF3ID(strJDGEIDC, strXTALC1) <> FUNCTION_RETURN_SUCCESS Then
                iErrFlg = 11
            Else
                If Trim(strJDGEIDC) = "" Then
                    iErrFlg = 12
'                '判定ID=「9」の場合は判定なし(判定OK)　07/08/01 M.Kaga
'                ElseIf Trim(strJDGEIDC) = "9" Then
'                    YFlg = True
'                    bJudg = True
                Else
                    '獲得したC-OSF3判定IDでXODC5_OSF30より承認ﾌﾗｸﾞの獲得
                    If GetSYNFLAGC5(strSYNFLG, strYMKFLG, strJDGEIDC) <> FUNCTION_RETURN_SUCCESS Then
                        iErrFlg = 13
                    Else
                        '承認ﾌﾗｸﾞ:0　未承認の場合
                        If Trim(strSYNFLG) = "0" Or Trim(strSYNFLG) = "" Or IsNull(strSYNFLG) Then
                            iErrFlg = 14
                        End If
                    End If
                End If
            End If
            
            If iErrFlg > 0 Then
                '画面表示内容設定
                .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' 結晶内開始位置
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' 内容
                .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                         ' 情報１
                .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                         ' 情報２
                .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                        ' 情報３
                .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' サンプルＮｏ
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' 品番(12桁)
                If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                    .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS           ' 結晶内開始位置
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo          ' サンプルＮｏ
                End If
                bJudg = False
            Else
                If YFlg = False Then
                    
                    If GetOsf31_CuDeco(strJDGEIDC, lngSMPPOS, typ_Ret_CuDeco) <> FUNCTION_RETURN_SUCCESS Then
                        iErrFlg = 21
                    Else
                        intSiyou = -1
                        intJisseki = -1
                        
                        If (IsNumeric(.typ_si(UpDo).HSXCJLTPK)) Then
                            intSiyou = CInt(.typ_si(UpDo).HSXCJLTPK)
                        End If
                        
                        If (IsNumeric(strPtnJsk)) Then
                            intJisseki = CInt(strPtnJsk) + 1
                        End If
                        
                        If (intSiyou >= 1) And (intSiyou <= 7) And (intJisseki >= 1) And (intJisseki <= 8) Then
                        
                           'intRet = funCodeDBGet(SYSCLASS, CLASS, 製品仕様パターン区分, 1, 実績パターン区分, 戻り値(tbcmb005.info1))
                            intRet = funCodeDBGet("SB", "S2", CStr(intSiyou), 1, CStr(intJisseki), sResult)
                            If (intRet = 0) And (sResult <> vbNullString) And (Len(sResult) >= 1) Then
                                If sResult = "1" Then
                                    bJudg = True
                                Else
                                    bJudg = False
                                    intNG_Num = 51
                                End If
                            Else
                                'sErr_Msg = sAdd_Msg & sErr_Msg & "→仕様:" & CStr(intSiyou) & ", 実績:" & CStr(intJisseki)
                                bJudg = False
                                intNG_Num = 52
'                                GoTo CodeDBGet_Error
                            End If
                            
                        End If
                        
                        ' CJ(LT) 計算Band幅の判定
                        If bJudg Then
                            'Cng Start 2012/06/05 Y.Hitomi
                                If (.typ_si(UpDo).HSXCJLTBND = -1) Or (.typ_si(UpDo).HSXCJLTBND > 150) Then
'                            If (strPtnJsk = CNST_JSK_PTN_PBband) Or (strPtnJsk = CNST_JSK_PTN_Pband) Or (strPtnJsk = CNST_JSK_PTN_Bband) Then
'                                If (.typ_si(UpDo).HSXCJLTBND = -1) Or (.typ_si(UpDo).HSXCJLTBND > 150) Then
                            'Cng Start 2012/06/05 Y.Hitomi
                                    bJudg = False
                                    intNG_Num = 61
                                ElseIf (.typ_zi.CuCJLT(UpDo).CJLTBANDGKJSK = -1) Or (.typ_zi.CuCJLT(UpDo).CJLTBANDGKJSK > 150) Then
                                    bJudg = False
                                    intNG_Num = 62
                                ElseIf (.typ_zi.CuCJLT(UpDo).CJLTBANDNKJSK = -1) Or (.typ_zi.CuCJLT(UpDo).CJLTBANDNKJSK > 150) Then
                                    bJudg = False
                                    intNG_Num = 63
                                ElseIf (.typ_zi.CuCJLT(UpDo).CJLTBANDGKJSK < .typ_zi.CuCJLT(UpDo).CJLTBANDNKJSK) Then
                                    bJudg = False
                                    intNG_Num = 64
                                ElseIf (.typ_si(UpDo).HSXCJLTBND < (.typ_zi.CuCJLT(UpDo).CJLTBANDGKJSK - .typ_zi.CuCJLT(UpDo).CJLTBANDNKJSK)) Then
                                    bJudg = False
                                    intNG_Num = 65
                                End If
'                            End If
                        End If
                    End If
                End If
                
                If iErrFlg > 0 Then
                    '画面表示内容設定
                    .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' 結晶内開始位置
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' 内容
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                         ' 情報１
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                         ' 情報２
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                        ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = CInt(iErrFlg)                    ' 情報４
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' サンプルＮｏ
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' 品番(12桁)
                    If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS              ' 結晶内開始位置
                        .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo             ' サンプルＮｏ
                    End If
                    bJudg = False
                    
                Else
                    '画面表示内容設定
                    .typ_rslt(UpDo, DispLineCount).pos = -1                                     ' 結晶内開始位置
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' 内容
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                             ' 情報１
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "検査無"                             ' 情報２
                    .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                   ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                                  ' サンプルＮｏ
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' 品番(12桁)
                    
                    If shiji Then
                        '画面表示内容設定
                        .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                         ' 情報１
                        .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                         ' 情報２
                        .typ_rslt(UpDo, DispLineCount).INFO3 = "実績無"                         ' 情報３
                        
                        If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                            .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS              ' 結晶内開始位置
                            .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo             ' サンプルＮｏ
                            .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ無"                    ' 情報３
                            
                            If strSampUmu = "0" Then
                                .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                ' 情報３
                                '画面表示内容設定
                                .typ_rslt(UpDo, DispLineCount).INFO1 = strInfo(0)
                                .typ_rslt(UpDo, DispLineCount).INFO2 = strInfo(1)
                                .typ_rslt(UpDo, DispLineCount).INFO3 = strInfo(2)
                                .typ_rslt(UpDo, DispLineCount).INFO4 = strInfo(3)
                            End If
                        End If
                    End If
                End If
            End If
            
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                               ' 判定結果
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                               ' 判定結果
                TotalJudg = False
                gsTbcmy028ErrCode = str028ErrCode
            End If
            DispLineCount = DispLineCount + 1
            
        Else
        ' Add Start 2011/02/15 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : 「保証方法_処」参考時の処理追加
            If shiji Then
                    '画面表示内容設定
                    .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS                              ' 結晶内開始位置
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' 内容
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様無"                             ' 情報１
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                             ' 情報２
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ無"                            ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo                           ' サンプルＮｏ
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).OKNG = "N参"                                 ' 判定結果
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' 品番(12桁)
                    
                    If strSampUmu = "0" Then
                        '画面表示内容設定
                        .typ_rslt(UpDo, DispLineCount).INFO1 = strInfo(0)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = strInfo(1)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = strInfo(2)
                        .typ_rslt(UpDo, DispLineCount).INFO4 = strInfo(3)
                    End If
                DispLineCount = DispLineCount + 1
            End If
        ' Add End   2011/02/15 SMPK A.Nagamine
        End If      '/*  End of If (JudgSpecCode) Or (pblnFlag) Then */
    End With
    
    CuDecoDataSet_CJLT = blnRet
    
End Function
''Add End   2011/01/17 SMPK A.Nagamine

'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : Cu-deco(CJ2)の実績判定処理
Public Function CuDecoDataSet_CJ2(pJudgSC_B() As Judg_Spec_Cry, ptyp_b As typ_AllTypesB, UpDo As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String, Optional pblnFlag As Boolean = False) As Boolean
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 UPD By Systech End
    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4                  '検査指示
    Dim bJudg           As Boolean
    Dim strXTALC1       As String
    Dim strJDGEIDC      As String
    Dim strSYNFLG       As String
    Dim strYMKFLG       As String
    Dim lngSMPPOS       As Long
    Dim iErrFlg         As Integer
    Dim YFlg            As Boolean
    Dim intRet          As Integer
    Dim sResult         As String
    Dim strSampUmu      As String
    Dim strInfo(4)      As String
    Dim intSiyou        As Integer
    Dim intJisseki      As Integer
    Dim StrCryNum       As String
    Dim lngSampNo       As Long
    Dim strPtnJsk       As String
    Dim str028ErrCode   As String
    Dim blnRet          As Boolean
    Dim typ_Ret_CuDeco  As typ_SB_com_xodb5_osf31_Cudeco
    Dim intNG_Num       As Integer
    Dim intMin          As Integer
    
    blnRet = False
    intNG_Num = 0
    
    '検査指示設定
    IND = IIf(UpDo = BlkTop, "123", "123")
    
    With ptyp_b
        ' CJ2 実績
        JudgSpecCode = pJudgSC_B(UpDo).CuCJ2
        SCC = "CJ2"
        shiji = (InStr(IND, .typ_cr(UpDo).CRYINDCJ2CS) <> 0)
        strSampUmu = .typ_zi.CuCJ2(UpDo).SMPLUMUCJ2             ' CJ2.サンプル有無 TBCMJ023.SMPLUMUCJ2 "0"=サンプル有り,"1"=サンプル無し
        strInfo(0) = CStr(.typ_zi.CuCJ2(UpDo).CJ2DISKJSK)         ' 情報１ CJ2.Disk半径
        strInfo(1) = CStr(.typ_zi.CuCJ2(UpDo).CJ2RINGNKJSK)       ' 情報２ CJ2.Ring内径
        strInfo(2) = CStr(.typ_zi.CuCJ2(UpDo).CJ2RINGGKJSK)       ' 情報３ CJ2.Ring外径
        strInfo(3) = .typ_zi.CuCJ2(UpDo).CJ2PTNJSK              ' 情報４ CJ2.パターン実績
        lngSMPPOS = .typ_zi.CuCJ2(UpDo).POSITION                ' CJ2.ｻﾝﾌﾟﾙ位置
        StrCryNum = .typ_zi.CuCJ2(UpDo).CRYNUM                  ' CJ2.結晶番号
        lngSampNo = .typ_zi.CuCJ2(UpDo).SMPLNO                  ' CJ2.サンプル№(代表サンプルID)
        strPtnJsk = .typ_zi.CuCJ2(UpDo).CJ2PTNJSK               ' CJ2.パターン実績
        str028ErrCode = "00158"
        
        '保証ﾌﾗｸﾞ="H"の場合か第７パラメータOptional がTrue指定されたとき
        If (JudgSpecCode) Or (pblnFlag) Then
            
            bJudg = False
            
            '判定ﾌﾗｸﾞ初期化
            iErrFlg = 0
            YFlg = False
            
            'ﾌﾞﾛｯｸID
            strXTALC1 = Trim(ptyp_b.BLOCKID)
            '結晶番号
            strXTALC1 = left(strXTALC1, 9) & "000"
            
            '結晶番号をｷｰとしてXSDC1よりC－OSF3判定IDを獲得する
            If GetCOSF3ID(strJDGEIDC, strXTALC1) <> FUNCTION_RETURN_SUCCESS Then
                iErrFlg = 11
            Else
                If Trim(strJDGEIDC) = "" Then
                    iErrFlg = 12
'                '判定ID=「9」の場合は判定なし(判定OK)　07/08/01 M.Kaga
'                ElseIf Trim(strJDGEIDC) = "9" Then
'                    YFlg = True
'                    bJudg = True
                Else
                    '獲得したC-OSF3判定IDでXODC5_OSF30より承認ﾌﾗｸﾞの獲得
                    If GetSYNFLAGC5(strSYNFLG, strYMKFLG, strJDGEIDC) <> FUNCTION_RETURN_SUCCESS Then
                        iErrFlg = 13
                    Else
                        '承認ﾌﾗｸﾞ:0　未承認の場合
                        If Trim(strSYNFLG) = "0" Or Trim(strSYNFLG) = "" Or IsNull(strSYNFLG) Then
                            iErrFlg = 14
                        End If
                    End If
                End If
            End If
            
            If iErrFlg > 0 Then
                '画面表示内容設定
                .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' 結晶内開始位置
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' 内容
                .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                         ' 情報１
                .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                         ' 情報２
                .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                        ' 情報３
                .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' サンプルＮｏ
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' 品番(12桁)
                If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                    .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS           ' 結晶内開始位置
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo          ' サンプルＮｏ
                End If
                bJudg = False
            Else
                If YFlg = False Then
                    
                    If GetOsf31_CuDeco(strJDGEIDC, lngSMPPOS, typ_Ret_CuDeco) <> FUNCTION_RETURN_SUCCESS Then
                        iErrFlg = 21
                    Else
                        intSiyou = -1
                        intJisseki = -1
                        
                        If (IsNumeric(.typ_si(UpDo).HSXCJ2PK)) Then
                            intSiyou = CInt(.typ_si(UpDo).HSXCJ2PK)
                        End If
                        
                        If (IsNumeric(strPtnJsk)) Then
                            intJisseki = CInt(strPtnJsk) + 1
                        End If
                        
                        If (intSiyou >= 1) And (intSiyou <= 4) And (intJisseki >= 1) And (intJisseki <= 4) Then
                        
                           'intRet = funCodeDBGet(SYSCLASS, CLASS, 製品仕様パターン区分, 1, 実績パターン区分, 戻り値(tbcmb005.info1))
                            intRet = funCodeDBGet("SB", "S1", CStr(intSiyou), 1, CStr(intJisseki), sResult)
                            If (intRet = 0) And (sResult <> vbNullString) And (Len(sResult) >= 1) Then
                                If sResult = "1" Then
                                    bJudg = True
                                Else
                                    bJudg = False
                                    intNG_Num = 51
                                End If
                            Else
                                'sErr_Msg = sAdd_Msg & sErr_Msg & "→仕様:" & CStr(intSiyou) & ", 実績:" & CStr(intJisseki)
                                bJudg = False
                                intNG_Num = 52
'                                GoTo CodeDBGet_Error
                            End If
                            
                        End If
                        
                        ' CJ2 Ring内径・外径の判定
                        If bJudg Then
                            If (strPtnJsk = CNST_JSK_PTN_Ring) Then
                                If (typ_Ret_CuDeco.CJ2RMINRINC5 = -1) Or (typ_Ret_CuDeco.CJ2RMINRINC5 > 150) Then
                                    bJudg = False
                                    intNG_Num = 61
                                ElseIf (.typ_zi.CuCJ2(UpDo).CJ2RINGNKJSK = -1) Or (.typ_zi.CuCJ2(UpDo).CJ2RINGNKJSK > 150) Then
                                    bJudg = False
                                    intNG_Num = 62
                                ElseIf (typ_Ret_CuDeco.CJ2RMAXRIGC5 = -1) Or (typ_Ret_CuDeco.CJ2RMAXRIGC5 > 150) Then
                                    bJudg = False
                                    intNG_Num = 63
                                ElseIf (.typ_zi.CuCJ2(UpDo).CJ2RINGGKJSK = -1) Or (.typ_zi.CuCJ2(UpDo).CJ2RINGGKJSK > 150) Then
                                    bJudg = False
                                    intNG_Num = 64
                                ElseIf (typ_Ret_CuDeco.CJ2RMINRINC5 > .typ_zi.CuCJ2(UpDo).CJ2RINGNKJSK) Then
                                    bJudg = False
                                    intNG_Num = 65
                                ElseIf (typ_Ret_CuDeco.CJ2RMAXRIGC5 < .typ_zi.CuCJ2(UpDo).CJ2RINGGKJSK) Then
                                    bJudg = False
                                    intNG_Num = 66
                                End If
                            ElseIf (strPtnJsk = CNST_JSK_PTN_DiskRing) Then
                                If (typ_Ret_CuDeco.CJ2DRMINRINC5 = -1) Or (typ_Ret_CuDeco.CJ2DRMINRINC5 > 150) Then
                                    bJudg = False
                                    intNG_Num = 71
                                ElseIf (.typ_zi.CuCJ2(UpDo).CJ2RINGNKJSK = -1) Or (.typ_zi.CuCJ2(UpDo).CJ2RINGNKJSK > 150) Then
                                    bJudg = False
                                    intNG_Num = 72
                                ElseIf (typ_Ret_CuDeco.CJ2DRMAXRIGC5 = -1) Or (typ_Ret_CuDeco.CJ2DRMAXRIGC5 > 150) Then
                                    bJudg = False
                                    intNG_Num = 73
                                ElseIf (.typ_zi.CuCJ2(UpDo).CJ2RINGGKJSK = -1) Or (.typ_zi.CuCJ2(UpDo).CJ2RINGGKJSK > 150) Then
                                    bJudg = False
                                    intNG_Num = 74
                                ElseIf (typ_Ret_CuDeco.CJ2DRMINRINC5 > .typ_zi.CuCJ2(UpDo).CJ2RINGNKJSK) Then
                                    bJudg = False
                                    intNG_Num = 75
                                ElseIf (typ_Ret_CuDeco.CJ2DRMAXRIGC5 < .typ_zi.CuCJ2(UpDo).CJ2RINGGKJSK) Then
                                    bJudg = False
                                    intNG_Num = 76
                                End If
                            End If
                        End If
                        
                        'CJ2 計算Pi幅の判定(下限値チェック)
                        If bJudg Then
                            If (strPtnJsk = CNST_JSK_PTN_Disk) Or (strPtnJsk = CNST_JSK_PTN_Ring) Or (strPtnJsk = CNST_JSK_PTN_DiskRing) Then
                                If (strPtnJsk = CNST_JSK_PTN_Disk) Then
                                    intMin = typ_Ret_CuDeco.CJ2DMAXPIC5
                                ElseIf (strPtnJsk = CNST_JSK_PTN_Ring) Then
                                    intMin = typ_Ret_CuDeco.CJ2RMAXPIC5
                                Else
                                    intMin = typ_Ret_CuDeco.CJ2DRMAXPIC5
                                End If
                                
                                If (intMin = -1) Or (intMin > 150) Then
                                    bJudg = False
                                    intNG_Num = 81
                                ElseIf (.typ_zi.CuCJ2(UpDo).CJ2PICALC = -1) Or (.typ_zi.CuCJ2(UpDo).CJ2PICALC > 150) Then
                                    bJudg = False
                                    intNG_Num = 82
                                ElseIf (intMin > .typ_zi.CuCJ2(UpDo).CJ2PICALC) Then
                                    bJudg = False
                                    intNG_Num = 83
                                End If
                            End If
                        End If
                        
                    End If
                End If
                
                If iErrFlg > 0 Then
                    '画面表示内容設定
                    .typ_rslt(UpDo, DispLineCount).pos = -1                                 ' 結晶内開始位置
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())  ' 内容
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                         ' 情報１
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                         ' 情報２
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                        ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = CInt(iErrFlg)                    ' 情報４
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                              ' サンプルＮｏ
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                       ' 品番(12桁)
                    If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                        .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS              ' 結晶内開始位置
                        .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo             ' サンプルＮｏ
                    End If
                    bJudg = False
                    
                Else
                    '画面表示内容設定
                    .typ_rslt(UpDo, DispLineCount).pos = -1                                     ' 結晶内開始位置
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' 内容
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                             ' 情報１
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "検査無"                             ' 情報２
                    .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                   ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = -1                                  ' サンプルＮｏ
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' 品番(12桁)
                    
                    If shiji Then
                        '画面表示内容設定
                        .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                         ' 情報１
                        .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                         ' 情報２
                        .typ_rslt(UpDo, DispLineCount).INFO3 = "実績無"                         ' 情報３
                        
                        If left(StrCryNum, 9) = left(.BLOCKID, 9) Then
                            .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS              ' 結晶内開始位置
                            .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo             ' サンプルＮｏ
                            .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ無"                    ' 情報３
                            
                            If strSampUmu = "0" Then
                                .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                ' 情報３
                                '画面表示内容設定
                                .typ_rslt(UpDo, DispLineCount).INFO1 = strInfo(0)
                                .typ_rslt(UpDo, DispLineCount).INFO2 = strInfo(1)
                                .typ_rslt(UpDo, DispLineCount).INFO3 = strInfo(2)
                                .typ_rslt(UpDo, DispLineCount).INFO4 = strInfo(3)
                            End If
                        End If
                    End If
                End If
            End If
            
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                               ' 判定結果
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                               ' 判定結果
                TotalJudg = False
                gsTbcmy028ErrCode = str028ErrCode
            End If
            DispLineCount = DispLineCount + 1
            
        Else
        ' Add Start 2011/02/15 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : 「保証方法_処」参考時の処理追加
            If shiji Then
                    '画面表示内容設定
                    .typ_rslt(UpDo, DispLineCount).pos = lngSMPPOS                              ' 結晶内開始位置
                    .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())      ' 内容
                    .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様無"                             ' 情報１
                    .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                             ' 情報２
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ無"                            ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                    .typ_rslt(UpDo, DispLineCount).SMPLNO = lngSampNo                           ' サンプルＮｏ
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).OKNG = "N参"                                 ' 判定結果
    '====================== Debug Debug =====================================
                    .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                           ' 品番(12桁)
                    
                    If strSampUmu = "0" Then
                        '画面表示内容設定
                        .typ_rslt(UpDo, DispLineCount).INFO1 = strInfo(0)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = strInfo(1)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = strInfo(2)
                        .typ_rslt(UpDo, DispLineCount).INFO4 = strInfo(3)
                    End If
                DispLineCount = DispLineCount + 1
            End If
        ' Add End   2011/02/15 SMPK A.Nagamine
        End If      '/*  End of If (JudgSpecCode) Or (pblnFlag) Then */
    End With
    
    CuDecoDataSet_CJ2 = blnRet
    
End Function
''Add End   2011/01/17 SMPK A.Nagamine


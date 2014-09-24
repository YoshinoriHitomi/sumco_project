Attribute VB_Name = "SB_WfJudg"
Option Explicit

'''''Public typ_Param001b As DBDRV_scmzc_fcmlc001b_SXL
'''''Public Const MAXREC As Integer = 256
'''''Private intChkPos As Integer                                    ' チェック位置

''Public MaxLine As Integer
Public SelectSxlID As String
'''''Public typ_ww() As DBDRV_scmzc_fcmlc001b_SXL   '待ち一覧情報
'''''Public WFJudgExecOkFlag() As Boolean    'WF総合判定実行可能フラグ

'判定フローはこんな感じですか？
' ＜判定フロー＞
' 仕様保証方法＿処 --+--なし --実績（該当位置）--あってもなくても判定OK
'　　　　　　　　　　|
'                   +--あり --実績（該当位置) --+--あり -- 判定チェック --+-- OK
'                                              |                        |
'                                              |                        +-- MG
'                                              |
'                                              +--なし --+-- 検査指示５・６以外の場合 ---NG
'　　　　　　　　　　　　　　　　　　　　　　　　（検査指示は、指示を立てる側が正常に立てていると考えている）


''
'' 定数定義
''
'''''Public lStfMst As Long
'''''Public intEnCmd As Integer
Private Const MAXCNT As Integer = 18                             ' 最大件数
Private Const SXL_MAXSMP As Integer = 1 + 1 + 10                ' SXL内の最大サンプル件数　'Add 2011/03/07 SMPK Miyata
                                                                '  - Top:MAX1件、Bot:MAX1件、中間抜試:MAX10件
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
Private Const MAXCNT_EP As Integer = 6                             ' 最大件数
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
Public Const SxlTop As Integer = 1                                 ' TOP側
Public Const SxlTail As Integer = 2                                ' TAIL側
Public Const SxlMidl As Integer = 3                                ' MIDLE側    'Add 2011/03/07 SMPK Miyata
'''''Public Const KSYSCLASS As String = "GP"                         ' システム区分
Public Const MSYSCLASS As String = "NM"                         ' システム区分
Public Const KCLASS As String = "01"                            ' クラス
Public Const KCODE As String = "1"                              ' コード

'''''Private Const cnEnableColor As Long = &H80FF80                  ' 有効カラー
'''''Private Const cnEnableColor2 As Long = vbWindowBackground       ' 有効カラー
'''''Private Const cnDisenableColor As Long = &H80FF80               ' 無効カラー
'''''Private Const cnDisenableGrayColor As Long = vbButtonFace       ' 無効カラー（灰色）
'''''Private Const cnWarningColor As Long = &H8080FF                 ' 警告カラー

Public Const WFRES As Integer = 0
Public Const WFOI As Integer = 1
Public Const WFBMD1 As Integer = 2
Public Const WFBMD2 As Integer = 3
Public Const WFBMD3 As Integer = 4
Public Const WFOSF1 As Integer = 5
Public Const WFOSF2 As Integer = 6
Public Const WFOSF3 As Integer = 7
Public Const WFOSF4 As Integer = 8
Public Const WFDS As Integer = 9
Public Const WFDZ As Integer = 10
Public Const WFSP As Integer = 11
Public Const WFDOI1 As Integer = 12
Public Const WFDOI2 As Integer = 13
Public Const WFDOI3 As Integer = 14
Public Const WFOT1 As Integer = 15
Public Const WFOT2 As Integer = 16
Public Const WFAOI As Integer = 17
Public Const WFGD As Integer = 18           '05/02/07 ooba
'''--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
Public Const EPBMD1 As Integer = 0
Public Const EPBMD2 As Integer = 1
Public Const EPBMD3 As Integer = 2
Public Const EPOSF1 As Integer = 3
Public Const EPOSF2 As Integer = 4
Public Const EPOSF3 As Integer = 5
Public Const EPOT2 As Integer = 6
'''--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
'Add 2010/01/07 SIRD対応 Y.Hitomi
Public Const WFSIRD As Integer = 8

Public Const OSWFRES As String = "RES"
Public Const OSWFOI As String = "OI"
Public Const OSWFBMD1 As String = "BMD1"
Public Const OSWFBMD2 As String = "BMD2"
Public Const OSWFBMD3 As String = "BMD3"
Public Const OSWFOSF1 As String = "OSF1"
Public Const OSWFOSF2 As String = "OSF2"
Public Const OSWFOSF3 As String = "OSF3"
Public Const OSWFOSF4 As String = "OSF4"
Public Const OSWFDS As String = "DSOD"
Public Const OSWFDZ As String = "DZ"
Public Const OSWFSP As String = "SPV"
Public Const OSWFDOI1 As String = "DOI1"
Public Const OSWFDOI2 As String = "DOI2"
Public Const OSWFDOI3 As String = "DOI3"
Public Const OSWFOT1 As String = "OT1"
Public Const OSWFOT2 As String = "OT2"
Public Const OSWFAOI As String = "AOI"      ''残存酸素追加　03/12/15 ooba
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
Public Const OSEPBMD1 As String = "BMD1"
Public Const OSEPBMD2 As String = "BMD2"
Public Const OSEPBMD3 As String = "BMD3"
Public Const OSEPOSF1 As String = "OSF1"
Public Const OSEPOSF2 As String = "OSF2"
Public Const OSEPOSF3 As String = "OSF3"
Public Const OSEPOT2 As String = "OTHER2"
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
'Add 2010/01/07 SIRD対応 Y.Hitomi
Public Const OSWFSIRD As String = "SIRD"


'''''' コードマスター
'''''Public Type typ_CodeMaster
'''''    SYSCLASS As String * 2          ' システム区分
'''''    Class As String * 2             ' 区分
'''''    CODE As String * 5              ' コード
'''''    INFO1 As String                 ' 情報１
'''''    INFO2 As String                 ' 情報２
'''''    INFO3 As String                 ' 情報３
'''''    INFO4 As String                 ' 情報４
'''''    INFO5 As String                 ' 情報５
'''''    INFO6 As String                 ' 情報６
'''''    INFO7 As String                 ' 情報７
'''''    INFO8 As String                 ' 情報８
'''''    INFO9 As String                 ' 情報９
'''''    NOTE As String                  ' 備考
'''''    TSTAFFID As String * 8          ' 登録社員ID
'''''    REGDATE As Date                 ' 登録日付
'''''    KSTAFFID As String * 8          ' 更新社員ID
'''''    UPDDATE As Date                 ' 更新日付
'''''End Type

''''''各実績情報
'''''Public Type typ_ALLRSLT
'''''    pos As Integer                    ' 結晶内開始位置
'''''    NAIYO As String                   ' 内容
'''''    INFO1 As String                   ' 情報１
'''''    INFO2 As String                   ' 情報２
'''''    INFO3 As String                   ' 情報３
'''''    INFO4 As String                   ' 情報４
'''''    OKNG  As String                   ' 判定結果
'''''    SMPLID As String                  ' サンプルＮｏ
'''''End Type

'Add Start 2011/03/10 SMPK Miyata
Public Type typ_TBCMY013_arry
    typ_y013midl()      As typ_TBCMY013
End Type

Public Type typ_TBCMY022_arry
    typ_y022midl()      As typ_TBCMY022
End Type
'Add End   2011/03/10 SMPK Miyata

'全情報構造体
Public Type typ_AllTypesC
    StrStaffId          As String                               ' スタッフID
    strStaffName        As String                               ' スタッフ名
'Chg Start 2011/03/09 SMPK Miyata
'    dblScut(2)          As Double                               ' 再カット位置
'    bOKNG(2)            As Boolean                              ' 比抵抗判定
'    COEF(2)             As Double                               ' 偏析係数
'    JudgRes(2)          As Boolean                              ' 比抵抗判定    2002/01/15 S.Sano
'    JudgRrg(2)          As Boolean                              ' RRG判定       2002/01/15 S.Sano
    dblScut()           As Double                               ' 再カット位置
    bOKNG()             As Boolean                              ' 比抵抗判定
    COEF()              As Double                               ' 偏析係数
    JudgRes()           As Boolean                              ' 比抵抗判定
    JudgRrg()           As Boolean                              ' RRG判定
'Chg End   2011/03/09 SMPK Miyata

'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    'AN温度判定用
    JudgAntnp(12)        As Boolean                              ' AN温度判定  'Cng 2011/08/12 Y.Hitomi
'    JudgAntnp(2)        As Boolean                              ' AN温度判定
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    JudgDkTmp(12)        As Boolean                              ' DK温度判定  'Cng 2011/08/12 Y.Hitomi
    DkTmpJsk(12)         As String                               ' DK温度(実績)'Cng 2011/08/12 Y.Hitomi
'    JudgDkTmp(2)        As Boolean                              ' DK温度判定
'    DkTmpJsk(2)         As String                               ' DK温度(実績)
    DkTmpSiyo           As String                               ' DK温度(仕様)
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    typ_Param           As DBDRV_scmzc_fcmlc001b_SXL            ' SXL管理（待ち一覧から）
    typ_si              As type_DBDRV_scmzc_fcmlc001c_Siyou     ' 製品仕様
    typ_y013top()       As typ_TBCMY013                         ' 測定結果(TOP)
    typ_y013tail()      As typ_TBCMY013                         ' 測定結果(TAIL)
    typ_y013midl_ary()  As typ_TBCMY013_arry                    ' 測定結果(MIDLE)   'Add 2011/03/07 SMPK Miyata
'Chg Start 2011/03/09 SMPK Miyata
'* VBの64k制限によりtyp_AllTypesC内のサイズを大きくできないので、
'* typ_y013を静的型→動的型に変更する
'    typ_y013(2, MAXCNT) As typ_TBCMY013                         ' 測定結果
'    typ_hage(2)         As typ_TBCMH004                         ' 引上げ終了実績
'    typ_rslt(2, MAXCNT) As typ_ALLRSLT                          ' 各実績情報
    typ_y013()          As typ_TBCMY013                         ' 測定結果
    typ_hage()          As typ_TBCMH004                         ' 引上げ終了実績
    typ_rslt()          As typ_ALLRSLT                          ' 各実績情報
'Chg End   2011/03/09 SMPK Miyata
    sMidErrMsg          As String                               ' 中間抜試チェックエラーメッセージ  'Add 2011/05/10 SMPK Miyata
End Type

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
'VBの64k制限によりtyp_AllTypesC内に定義できないため、別構造体として作成する
'全情報構造体(エピ分)
Public Type typ_AllTypesC_EP
    typ_y022top()       As typ_TBCMY022                         ' エピ先行測定結果(TOP)
    typ_y022tail()      As typ_TBCMY022                         ' エピ先行測定結果(TAIL)
    typ_y022midl_ary()  As typ_TBCMY022_arry                    ' エピ先行測定結果(MIDLE)   'Add 2011/03/10 SMPK Miyata
'Chg Start 2011/03/10 SMPK Miyata
'    typ_y022(2, MAXCNT_EP)      As typ_TBCMY022                 ' エピ先行測定結果
'    typ_rslt(2, MAXCNT_EP)      As typ_ALLRSLT_EX               ' 各実績情報
    typ_y022(SXL_MAXSMP, MAXCNT_EP)     As typ_TBCMY022          ' エピ先行測定結果
    typ_rslt(SXL_MAXSMP, MAXCNT_EP)     As typ_ALLRSLT_EX        ' 各実績情報
'Chg End   2011/03/10 SMPK Miyata
End Type
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

'仕様検査支持構造体
Type Judg_Spec_Wf
    rs      As Boolean
    Oi      As Boolean
    B1      As Boolean
    B2      As Boolean
    B3      As Boolean
    L1      As Boolean
    L2      As Boolean
    L3      As Boolean
    L4      As Boolean
    Dsod    As Boolean
    sp      As Boolean
    DZ      As Boolean
    Doi1    As Boolean
    Doi2    As Boolean
    Doi3    As Boolean
    OT1     As Boolean
    OT2     As Boolean
    AOI     As Boolean
    GD      As Boolean      'GD追加　05/01/27 ooba
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    B1E     As Boolean
    B2E     As Boolean
    B3E     As Boolean
    L1E     As Boolean
    L2E     As Boolean
    L3E     As Boolean
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
'Add 2010/01/07 SIRD対応 Y.Hitomi
    SIRD    As Boolean
End Type

Public JudgSW       As Judg_Spec_Wf             '仕様検査支持構造体
Public typ_CType    As typ_AllTypesC            '全情報構造体
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
Public typ_CType_EP As typ_AllTypesC_EP         '全情報構造体(エピ)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
Public TotalJudg    As Boolean                  'トータル判定
Public MidlJudg     As Boolean                  '中間抜試判定   Add 2011/03/09 SMPK Miyata

'Chg Start 2011/03/07 SMPK Miyata
'Public typ_J015_WFGDJudg(2) As typ_TBCMJ015     'WF総合判定用GD実績　05/02/04 ooba
Public typ_J015_WFGDJudg() As typ_TBCMJ015     'WF総合判定用GD実績
'Chg End   2011/03/07 SMPK Miyata
Public typ_J015_WFGDUpd() As typ_TBCMJ015       'TBCMJ015-UPDATE用GD実績　05/02/07 ooba
Public iCntJ015upd As Integer                   'TBCMJ015-UPDATEﾚｺｰﾄﾞ数　05/02/07 ooba

'Chg Start 2011/03/07 SMPK Miyata
'''Upd start 2005/06/21 (TCS)T.Terauchi  SPV9点対応
'Public typ_J016_WFSPVJudg(2) As typ_TBCMJ016     'WF総合判定用SPV実績
'''Upd end   2005/06/21 (TCS)T.Terauchi  SPV9点対応
Public typ_J016_WFSPVJudg() As typ_TBCMJ016     'WF総合判定用SPV実績
'Chg End   2011/03/07 SMPK Miyata

'Chg Start 2011/03/07 SMPK Miyata
'''↓Add 2010/01/12 SIRD対応 Y.Hitomi
'Public typ_J022_WFSDJudg(2) As typ_TBCMJ022     'WF総合判定用SIRD実績
'''↑Add 2010/01/12 SIRD対応 Y.Hitomi
Public typ_J022_WFSDJudg() As typ_TBCMJ022     'WF総合判定用SIRD実績
'Chg End   2011/03/07 SMPK Miyata

'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
'既存の構造体に項目追加するとVBの制限に引っかかるので、別で管理する。
'各判定結果情報
'Chg Start 2011/03/09 SMPK Miyata
'Public typ_rslt_ex(2, MAXCNT) As typ_ALLRSLT_EX                          ' 各実績情報
Public typ_rslt_ex(SXL_MAXSMP, MAXCNT) As typ_ALLRSLT_EX                  ' 各実績情報
'Chg End   2011/03/09 SMPK Miyata
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------


'''''Public HErrMsg As String
'''''Public typ_rt(2) As typ_TBCMW009            '総合判定測定値 2001/09/14 S.Sano
'''''Public bPPlus As Boolean                    'P+Flag 2001/12/18 S.Sano
'''''Public bNPlus As Boolean                    'N+Flag 2002/01/08 S.Sano
'Chg Start 2011/03/09 SMPK Miyata
'Public JiltusekiUmu(2, MAXCNT) As Boolean       '実績有無情報 2001/12/19 S.Sano
Public JiltusekiUmu(SXL_MAXSMP, MAXCNT) As Boolean  '実績有無情報
'Chg End   2011/03/09 SMPK Miyata
'''''Public MeasFlag(2) As Judg_Spec_Wf         '仕様検査支持構造体
''''Public Tokusai As String                    ' 特採フラグ    'del 2003/05/28 hitec)matsumoto 宣言が２つあった

'Chg Start 2011/03/09 SMPK Miyata
'Public TmpOsfData(1, 2, MAXCNT) As String  'OSF平均/最大値　2003/05/20 ooba
'Public TmpOsfMBNP(2, 2, MAXCNT) As String * 1  'OSF面内分布　2003/05/21 ooba
Public TmpOsfData(1, SXL_MAXSMP, MAXCNT) As String      'OSF平均/最大値
Public TmpOsfMBNP(2, SXL_MAXSMP, MAXCNT) As String * 1  'OSF面内分布
'Chg End   2011/03/09 SMPK Miyata
Public wiSmpGetFlg  As Integer              'ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ(0:ｻﾝﾌﾟﾙ指定なし, 1:ｻﾝﾌﾟﾙ指定あり)
Public wiKcnt       As Integer              '工程連番

'--------------- 2008/07/25 INSERT START  By Systech ---------------
Private pbGDJudgeTbl(3) As Boolean          ' GD判定結果退避
'--------------- 2008/07/25 INSERT  END   By Systech ---------------

''Public STAFFIDBUFF  As String
''
''''エラーメッセージ
''Public Const ESTAF = "ESTAF" ''担当者コードが無効です｡
''Public Const EIE00 = "EIE00" ''全てのデータ入力が完了していません｡
''Public Const EIE01 = "EBLK1" ''ブロックIDの桁数が間違っています｡
''Public Const EIM00 = "EIM00" ''購入単結晶受入実績　問い合わせ中。
''Public Const EGET = "EGET" ''DBからの読込に失敗しました。
''Public Const EAPLY = "EAPLY" ''DBへの書込に失敗しました。
''Public Const EMAT1 = "EMAT1" '' 原料番号の桁数が間違っています｡
''Public Const EMAT2 = "EMAT2" '' 指定した原料番号は未登録です。
''Public Const KIE00 = "EBLK0" ''入力されたブロックIDは､存在しません｡
''Public Const KDE01 = "KDE01" ''購入単結晶は、イメージ表示できません。
''Public Const PWAIT = "PWAIT" ''少々お待ち下さい
''Public Const KC001 = "EKC01" ''クリスタルカタログ処理が失敗しました！
''Public Const TJE01 = "PJE01" ''総合判定NGです。
''Public Const ESXL0 = "ESXL0" ''入力されたSXLIDは、存在しません。"
''Public Const ESXL1 = "ESXL1" ''SXLIDの桁数が間違っています。"
''Public Const EHIN1 = "EHIN1" ''品番の桁数が間違っています。"
''Public Const EHIN0 = "EHIN0" ''指定の品番は未登録です。"



'''''' SXLの対象となるﾌﾞﾛｯｸ保存用構造体
'''''Public Type typ_IntoBlock
'''''    SORTID As String
'''''    FULLID As String
'''''End Type

'''''' ブロック情報
'''''Public Type typ_BlkInf
'''''    BLOCKID As String * 12      ' ブロックID
'''''    LENGTH As Integer           ' 長さ
'''''    REALLEN As Integer          ' 実長さ
'''''    KRPROCCD As String * 5      ' 現在管理工程
'''''    NOWPROC As String * 5       ' 現在工程
'''''    LPKRPROCCD As String * 5    ' 最終通過管理工程
'''''    LASTPASS As String * 5      ' 最終通過工程
'''''    RSTATCLS As String * 1      ' 流動状態区分
'''''    SEED As String * 4          ' シード
'''''    COF As type_Coefficient     ' 偏析係数計算
'''''    SAMPFLAG As Boolean         ' サンプル取得フラグ
'''''End Type

''''''カット位置用構造体
'''''Public Type typ_CMKC001C
'''''    CRYNUM As String * 12       ' 結晶番号
'''''    IngotPos As Integer         ' 結晶内開始位置
'''''    LENGTH As Integer           ' 長さ
'''''End Type

'''''' ブロック情報
'''''Public Type typ_BlkInf3
'''''    BLOCKID As String * 12      ' ブロックID
'''''    LENGTH As Integer           ' 長さ
'''''    REALLEN As Integer          ' 実長さ
'''''    NOWPROC As String * 5       ' 現在工程
'''''    DELFLG As String * 1        ' 削除区分
'''''    COF As type_Coefficient     ' 偏析係数計算
'''''End Type

'''''Public tblHinMng() As typ_TBCME041                      ' 品番管理
'''''Public tblWafSmp() As typ_TBCME044                      ' ＷＦサンプル管理
'''''Public tblBlkInf() As typ_BlkInf                        ' ブロック情報テーブル
'''''Public tblTotal As typ_AllTypesC                        ' 前画面からの情報保持構造体
'''''Public tblWfSxlMng() As typ_TBCME042                    ' SXL管理構造体
'''''Public tblWfSxlMngS() As typ_TBCME042                   ' 測定評価指示用SXL管理構造体
'''''Public tblWfSample() As typ_WfSampleGr                  ' WFサンプル管理
'''''Public SxlIntoBlock() As typ_IntoBlock                  ' SXLの対象となるﾌﾞﾛｯｸ構造体
'''''Public tblPrcList() As typ_TBCMB005                     ' 区分用コードマスター構造体
'''''Public tblHinbanRs() As type_DBDRV_scmzc_fcmlc001d_In   ' 品番情報保持構造体
'''''Public tblsiyou() As type_DBDRV_scmzc_fcmlc001d_WfSiyou ' 仕様情報構造体(表示用)
'''''Public tblsmp() As type_DBDRV_scmzc_fcmlc001d_WfSmp     ' サンプル情報構造体(表示用)
'''''Public tblWfHantei As typ_TBCMW005                      ' WF総合判定実績
'''''Public tblHuriHai() As typ_TBCMW006                     ' 振替廃棄実績
'''''Public tblSokuSizi() As typ_TBCMY003                    ' 測定評価方法指示構造体
'''''Public tblSxlKSiji() As typ_TBCMY007                    ' Ｓｘｌ確定指示
'''''Public NoTestHinList() As tFullHinban  ' 抜試の発生しない品番

'''''' 抜試指示
'''''Public Type typ_WafInd
'''''    BLOCKID As String * 12      ' ブロックID
'''''    BlockPos As Integer         ' ブロックＰ
'''''''''    BkSampleId  As Variant      ' add 2003/03/28 hitec)matsumoto 元サンプルIDを取得
'''''    SAMPLEID    As Variant      ' add 2003/03/28 hitec)matsumoto サンプルIDを取得
'''''    SAMPLEID2   As Variant      ' add 2003/03/28 hitec)matsumoto サンプルID2を取得
'''''    IngotPos As Integer         ' 結晶Ｐ
'''''    BkIngotPos  As Integer
'''''    LENGTH As Integer           ' 長さ
'''''    HINUP As tFullHinban        ' 上品番
'''''    HINDN As tFullHinban        ' 下品番
'''''    SMP As typ_WFSample         ' 検査項目
'''''    HinFlg As Boolean           ' 品番区切りフラグ
'''''    SMPFLG As Boolean           ' WFサンプル区切りフラグ
'''''    ERRDNFLG As Boolean         ' 下品番エラーフラグ
'''''    SMPLKBN1 As String * 1      ' サンプル区分１
'''''    SMPLKBN2 As String * 1      ' サンプル区分２
'''''End Type
'''''Public tblWafInd() As typ_WafInd        ' 抜試指示テーブル

'''''' 欠落ウェハー
'''''Public Type typ_LackMap
'''''    BLOCKID As String * 12      ' ブロックID
'''''    LACKPOSS As Double          ' 欠落位置(From)
'''''    LACKPOSE As Double          ' 欠落位置(To)
'''''    REJCAT As String * 1        ' 欠落理由
'''''    LACKCNTS As Integer         ' 欠落枚目(From)
'''''    LACKCNTE As Integer         ' 欠落枚目(To)
'''''End Type
'''''Public tblLackMap() As typ_LackMap      ' 欠落ウェハーテーブル


'''''' SXLサンプル情報
'''''Public Type typ_SxlSmp
'''''    strCRYNUM As String * 12          ' 結晶番号
'''''    intINGOTPOS As Integer            ' 結晶内開始位置
'''''    intLength As Integer              ' 長さ
'''''    strSXLID As String * 13           ' SXLID
'''''    strHINBAN As String * 12          ' 品番
'''''    strSMPLID As String * 16          ' サンプルID
'''''    intCount As Integer               ' 枚数
'''''    strSMPLUMU As String * 1          ' サンプル有無区分
'''''    datREGDATE As Date                ' 登録日付
'''''    datUPDDATE As Date                ' 更新日付
'''''    strWFINDRS As String * 1          ' WF検査指示（Rs)
'''''    strWFINDOI As String * 1          ' WF検査指示（Oi)
'''''    strWFINDB1 As String * 1          ' WF検査指示（B1)
'''''    strWFINDB2 As String * 1          ' WF検査指示（B2）
'''''    strWFINDB3 As String * 1          ' WF検査指示（B3)
'''''    strWFINDL1 As String * 1          ' WF検査指示（L1)
'''''    strWFINDL2 As String * 1          ' WF検査指示（L2)
'''''    strWFINDL3 As String * 1          ' WF検査指示（L3)
'''''    strWFINDL4 As String * 1          ' WF検査指示（L4)
'''''    strWFINDDS As String * 1          ' WF検査指示（DS)
'''''    strWFINDDZ As String * 1          ' WF検査指示（DZ)
'''''    strWFINDSP As String * 1          ' WF検査指示（SP)
'''''    strWFINDDO1 As String * 1         ' WF検査指示（DO1)
'''''    strWFINDDO2 As String * 1         ' WF検査指示（DO2)
'''''    strWFINDDO3 As String * 1         ' WF検査指示（DO3)
'''''    strWFRESRS As String * 1          ' WF検査実績（Rs)
'''''    strWFRESOI As String * 1          ' WF検査実績（Oi)
'''''    strWFRESB1 As String * 1          ' WF検査実績（B1)
'''''    strWFRESB2 As String * 1          ' WF検査実績（B2）
'''''    strWFRESB3 As String * 1          ' WF検査実績（B3)
'''''    strWFRESL1 As String * 1          ' WF検査実績（L1)
'''''    strWFRESL2 As String * 1          ' WF検査実績（L2)
'''''    strWFRESL3 As String * 1          ' WF検査実績（L3)
'''''    strWFRESL4 As String * 1          ' WF検査実績（L4)
'''''    strWFRESDS As String * 1          ' WF検査実績（DS)
'''''    strWFRESDZ As String * 1          ' WF検査実績（DZ)
'''''    strWFRESSP As String * 1          ' WF検査実績（SP)
'''''    strWFRESDO1 As String * 1         ' WF検査実績（DO1)
'''''    strWFRESDO2 As String * 1         ' WF検査実績（DO2)
'''''    strWFRESDO3 As String * 1         ' WF検査実績（DO3)
'''''End Type
'''''============================================================================================================================
'''''
''''''概要      :パラメータ設定
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
''''''説明      :前画面からの引数を設定する
''''''履歴      :
'''''Public Sub S_SetParamData()
'''''    typ_CType.typ_Param = typ_Param001b
'''''End Sub
'''''============================================================================================================================

'----------------------------------------------------------------------
'引数
'----------------------------------------------------------------------
'品種
'SelectBlkID
'tt(top,tail)
'全情報構造体
'仕様検査支持構造体
'仕様検査支持構造体
'トータル判定
'戻り値（配列）

'------------------------------------------------
' 総合判定
'------------------------------------------------

'概要      :実績値の総合判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型             :説明
'          :sKeyID          ,I  ,String         :SXL-ID
'          :tNew_Hinban     ,I  ,String         :振替候補品番
'          :bTotalJudg      ,O  ,Boolean        :トータル判定
'          :iErr_Code       ,O  ,Integer        :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String         :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :typ_CType       ,O  ,typ_AllTypesC  :全情報構造体(構造体)
'          :iSmpGetFlg      ,I  ,Integer        :ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ(0:ｻﾝﾌﾟﾙ指定なし, 1:ｻﾝﾌﾟﾙ指定あり)
'          :sSamplID1       ,I  ,String         :TOPｻﾝﾌﾟﾙID(省略可)
'          :sSamplID2       ,I  ,String         :BOTｻﾝﾌﾟﾙID(省略可)
'          :iKcnt           ,I  ,Integer        :工程連番(省略可)
'          :戻り値          ,O  ,Integer        :取得の成否(0:正常終了, -1:異常終了)
'説明      :
'履歴      :2003/09/19 新規作成　SB

Public Function funWfcSogoHantei(sKeyID As String, tNew_Hinban As tFullHinban, _
                bTotalJudg As Boolean, iErr_Code As Integer, sErr_Msg As String, typ_CType As typ_AllTypesC, _
                iSmpGetFlg As Integer, Optional sSamplID1 As String = vbNullString, Optional sSamplID2 As String = vbNullString, _
                Optional iKcnt As Integer = 0) As Integer
    
    Dim i       As Integer      'Add 2011/03/09 SMPK Miyata
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funWfcSogoHantei = FUNCTION_RETURN_FAILURE
    
    'グローバル変数に設定
    wiSmpGetFlg = iSmpGetFlg
    wiKcnt = iKcnt
    
    '初期設定
    sErr_Msg = "WFC総合判定(SetInitData)"
    If SetInitData(sKeyID, tNew_Hinban, typ_CType.typ_Param, iSmpGetFlg, sSamplID1, sSamplID2) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    
    '画面情報設定
    sErr_Msg = "WFC総合判定(SetAllData)"
    If SetAllData(typ_CType, tNew_Hinban, iSmpGetFlg, iErr_Code, sErr_Msg) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
        
    TotalJudg = True
    MidlJudg = True             '中間抜試判定   Add 2011/03/09 SMPK Miyata
    typ_CType.sMidErrMsg = ""   '中間抜試チェックエラーメッセージ   Add 2011/05/10 SMPK Miyata

'    funWfcSogoHantei = FUNCTION_RETURN_FAILURE
'
    '仕様検査指示取得
    sErr_Msg = "WFC総合判定(SpecJudgCheck)"
    SpecJudgCheck
    
    '2003/12/13 SystemBrain Null対応追加▽
    '仕様Nullチェック
    sErr_Msg = "仕様Nullﾁｪｯｸ"
    If funWfChkNull(typ_CType.typ_si, sErr_Msg) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    '2003/12/13 SystemBrain Null対応追加△
    
    '実績データ判定(TOP)
    sErr_Msg = "WFC総合判定(判定(TOP))"
    If WfAllJudg(typ_CType, tNew_Hinban, SxlTop) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    
    '実績データ判定(TAIL)
    sErr_Msg = "WFC総合判定(判定(TAIL))"
    If WfAllJudg(typ_CType, tNew_Hinban, SxlTail) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If

'Add Start 2011/03/09 SMPK Miyata
    '実績データ判定(MIDLE)
    sErr_Msg = "WFC総合判定(判定(MIDLE))"
    For i = SxlMidl To UBound(typ_CType.typ_Param.WFSMP)
        If WfAllJudg(typ_CType, tNew_Hinban, i) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    Next i

    'Add Start 2011/07/19 Y.Hitomi
        '中間抜試枚数のチェック
        Dim iMidCnt         As Integer       '中間抜試の枚数
        Dim iSmpMai()       As Integer       '抜試枚数保存配列
'Cng Start 2011/08/10 Y.Hitomi
            ' 中間抜試品の場合
    If typ_CType.typ_si.MSMPFLG = "1" Or typ_CType.typ_si.MSMPFLG = "2" Or typ_CType.typ_si.MSMPFLG = "3" Then
'        If typ_CType.typ_si.MSMPFLG = "1" Then
'Cng End   2011/08/10 Y.Hitomi
            ReDim iSxlPos(0)
            '抜試枚数の取得
            If fncGetSmpMai(sKeyID, iSmpMai) = FUNCTION_RETURN_FAILURE Then
                MidlJudg = False
            Else
                
                For i = 0 To UBound(iSmpMai)
                    '最終位置の枚数チェック
                    If i = UBound(iSmpMai) Then
                    '枚数のチェック
'Cng Start 2011/10/25 Y.Hitomi
                        If iSmpMai(i) <= typ_CType.typ_si.MSMPTANIMAI Then
'                        If iSmpMai(i) >= typ_CType.typ_si.MSMPCONSTMAI And _
'                            iSmpMai(i) <= typ_CType.typ_si.MSMPTANIMAI Then
'Cng End 2011/10/25 Y.Hitomi
                        Else
                            If iSmpMai(i) > typ_CType.typ_si.MSMPTANIMAI And _
                                iSmpMai(i) < typ_CType.typ_si.MSMPCONSTMAI + typ_CType.typ_si.MSMPTANIMAI Then
                            Else
                                typ_CType.sMidErrMsg = "中間抜試枚数が不足しています。実績枚数(" & CStr(iSmpMai(i)) & ")"
                                MidlJudg = False
                            End If
                        End If
                    Else
                    '枚数のチェック
'Cng Start 2011/08/25 Y.Hitomi
'                        If iSmpMai(i) >= typ_CType.typ_si.MSMPCONSTMAI And _
'                            iSmpMai(i) <= typ_CType.typ_si.MSMPTANIMAI Then
                    If iSmpMai(i) <= typ_CType.typ_si.MSMPTANIMAI Then
'Cng End   2011/08/25 Y.Hitomi
                        Else
                            typ_CType.sMidErrMsg = "中間抜試枚数が不足しています。実績枚数(" & CStr(iSmpMai(i)) & ")"
                            MidlJudg = False
                            Exit For
                        End If
                    End If
                Next i
                
            End If
        End If
        
        Dim iMinMidCnt      As Integer       '中間抜試の必要数
        Dim iRstMidCnt      As Integer       '中間抜試の件数
            
'Cng Start 2011/08/10 Y.Hitomi
            ' 中間抜試品の場合
    If typ_CType.typ_si.MSMPFLG = "1" Or typ_CType.typ_si.MSMPFLG = "2" Or typ_CType.typ_si.MSMPFLG = "3" Then
'        If typ_CType.typ_si.MSMPFLG = "1" Then
'Cng End   2011/08/10 Y.Hitomi
                '中間抜試の必要数 = (SXLのWF枚数 - 中間抜試許容値(枚数)) / 中間抜試単位(枚数)
                iMinMidCnt = Fix((typ_CType.typ_Param.COUNT - typ_CType.typ_si.MSMPCONSTMAI) / typ_CType.typ_si.MSMPTANIMAI)
                'マイナスの場合、０とする
                If iMinMidCnt < 0 Then iMinMidCnt = 0
                
                '中間抜試の件数
                iRstMidCnt = (UBound(typ_CType.typ_Param.WFSMP) - SxlMidl) + 1
                If iRstMidCnt < iMinMidCnt Then
                    typ_CType.sMidErrMsg = "中間抜試実績がありません。　仕様(" & iMinMidCnt & ") 実績(" & iRstMidCnt & ")"
                    MidlJudg = False
                End If
            End If
            
'Add Start 2011/11/28 Y.Hitomi
        '中間抜試=保証の場合のみ、判定NGとする。
    If typ_CType.typ_si.MSMPFLG = "2" Or typ_CType.typ_si.MSMPFLG = "0" Or typ_CType.typ_si.MSMPFLG = " " Then
                MidlJudg = True
        End If
'Add End   2011/11/28 Y.Hitomi


'Chg Start 2011/03/09 SMPK Miyata
'    bTotalJudg = TotalJudg
    bTotalJudg = TotalJudg And MidlJudg
'Chg End   2011/03/09 SMPK Miyata

    funWfcSogoHantei = FUNCTION_RETURN_SUCCESS
    
'------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funWfcSogoHantei = -4
    iErr_Code = funWfcSogoHantei
    GoTo Apl_Exit
    
End Function

'概要      :画面情報データ設定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typ_CType     ,I  ,typ_AllTypesC ,各情報構造体
'説明      :画面情報を情報構造体に設定する
'履歴      :
Private Function SetAllData(typ_CType As typ_AllTypesC, tNew_Hinban As tFullHinban, iSmpGetFlg As Integer, _
                                                                        iErr_Code As Integer, sErr_Msg As String) As FUNCTION_RETURN
    
    Dim typ_in      As type_DBDRV_scmzc_fcmlc001c_In     ' DBアクセス入力用
    Dim fret(2)     As FUNCTION_RETURN
    Dim RET         As FUNCTION_RETURN
    Dim typ_hi()    As typ_TBCMH004
    Dim sErrMsg     As String
    Dim FuncAns     As FUNCTION_RETURN ''2001/12/18 S.Sano
    Dim records()   As typ_TBCMH001
'Add Start 2011/03/07 SMPK Miyata
    Dim i           As Integer      'カウンタ
    Dim iMidNo      As Integer      '中間抜試No
'Add End   2011/03/07 SMPK Miyata

    SetAllData = FUNCTION_RETURN_FAILURE
    
    'TOP側
    sErr_Msg = "WFC総合判定(TOP 初期ﾃﾞｰﾀ設定)"
    typ_in.HIN.hinban = typ_CType.typ_Param.hinban
    typ_in.HIN.factory = typ_CType.typ_Param.factory
    typ_in.HIN.mnorevno = typ_CType.typ_Param.REVNUM
    typ_in.HIN.opecond = typ_CType.typ_Param.opecond
    typ_in.SAMPLEID = typ_CType.typ_Param.WFSMP(SxlTop).REPSMPLIDCW
    typ_in.SXLID = typ_CType.typ_Param.SXLID
    typ_in.WFSMP = typ_CType.typ_Param.WFSMP(SxlTop)
    
    With typ_CType
        ReDim .typ_y013top(0)
        '評価測定結果取得
        sErr_Msg = "WFC総合判定(TOP funWfcGetDataEtc)"

'↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
        '' WF仕様(SPV)取得
        If funWfcGetDataEtc_SPV(tNew_Hinban, _
                                .typ_si, _
                                sErrMsg) <> FUNCTION_RETURN_SUCCESS Then
            'WF仕様(SPV)取得失敗
            SetAllData = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
'↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------

        'パラメータSxlTop追加　Add 2011/03/07 SMPK Miyata
        fret(SxlTop) = funWfcGetDataEtc(typ_in, SxlTop, tNew_Hinban, iSmpGetFlg, _
                                        .typ_si, _
                                        .typ_y013top(), _
                                        sErrMsg)
        If fret(SxlTop) = FUNCTION_RETURN_SUCCESS Then
            ' 評価測定結果整列
            sErr_Msg = "WFC総合判定(TOP 評価測定結果整列)"
            If SetMERInd(typ_CType, .typ_y013top(), SxlTop) <> True Then
                '評価測定結果整列失敗
                Exit Function
            End If
'''''            ' WF検査指示（Rs)
'''''            If InStr("1345", .typ_Param.WFSMP(SxlTop).WFINDRSCW) <> 0 _
'''''            And .typ_Param.WFSMP(SxlTop).WFINDRSCW = "1" Then
'''''            End If
'''''            ' WF検査指示（Oi)
'''''            If InStr("1345", .typ_Param.WFSMP(SxlTop).WFINDOICW) <> 0 _
'''''            And .typ_Param.WFSMP(SxlTop).WFINDOICW = "1" Then
'''''            End If
            '引上げ終了実績取得
            ReDim typ_hi(0)
'頭8を購入単結晶扱いしない 2007/10/10 SETsw kubota
'            If Mid(.typ_Param.CRYNUM, 1, 1) <> "8" Then
                sErr_Msg = "WFC総合判定(TOP 引上げ終了実績取得)"
                If s_cmmc001db_Sql(.typ_Param.CRYNUM, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                    '引上げ終了実績取得失敗
                    Exit Function
                Else
                    If UBound(typ_hi) <> 0 Then
                        .typ_hage(SxlTop) = typ_hi(1)
                    Else
                        '引上げ終了実績取得失敗
                        SetAllData = FUNCTION_RETURN_FAILURE
                        Exit Function
                    End If
                End If
'            End If
        Else
            SetAllData = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
    
            
    
        'TAIL側
        sErr_Msg = "WFC総合判定(TAIL 初期ﾃﾞｰﾀ設定)"
        typ_in.SAMPLEID = .typ_Param.WFSMP(SxlTail).REPSMPLIDCW
        typ_in.WFSMP = typ_CType.typ_Param.WFSMP(SxlTail)
    
        '評価測定結果取得
        ReDim .typ_y013tail(0)
        sErr_Msg = "WFC総合判定(TAIL funWfcGetDataEtc)"
        'パラメータSxlTail追加　Add 2011/03/07 SMPK Miyata
        fret(SxlTail) = funWfcGetDataEtc(typ_in, SxlTail, tNew_Hinban, iSmpGetFlg, _
                                         .typ_si, _
                                         .typ_y013tail(), _
                                         sErrMsg)
        If fret(SxlTail) = FUNCTION_RETURN_SUCCESS Then
            ' 評価測定結果整列
            sErr_Msg = "WFC総合判定(TAIL 評価測定結果整列)"
            If SetMERInd(typ_CType, .typ_y013tail(), SxlTail) <> True Then
                '評価測定結果整列失敗
                Exit Function
            End If
'''''            ' WF検査指示（Rs)
'''''            If InStr("2345", .typ_Param.WFSMP(SxlTail).WFINDRSCW) <> 0 _
'''''            And .typ_Param.WFSMP(SxlTail).WFINDRSCW = "1" Then
'''''            End If
'''''            ' WF検査指示（Oi)
'''''            If InStr("2345", .typ_Param.WFSMP(SxlTail).WFINDOICW) <> 0 _
'''''            And .typ_Param.WFSMP(SxlTail).WFINDOICW = "1" Then
'''''            End If
            '引上げ終了実績取得
            ReDim typ_hi(0)
'頭8を購入単結晶扱いしない 2007/10/10 SETsw kubota
'            If Mid(.typ_Param.CRYNUM, 1, 1) <> "8" Then
                sErr_Msg = "WFC総合判定(TAIL 引上げ終了実績取得)"
                If s_cmmc001db_Sql(.typ_Param.CRYNUM, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                    '引上げ終了実績取得失敗
                    Exit Function
                Else
                    If UBound(typ_hi) <> 0 Then
                        .typ_hage(SxlTail) = typ_hi(1)
                    Else
                        '引上げ終了実績取得失敗
                        SetAllData = FUNCTION_RETURN_FAILURE
                        Exit Function
                    End If
                End If
'            End If
        Else
            SetAllData = FUNCTION_RETURN_FAILURE
            Exit Function
        End If

'Add Start 2011/03/07 SMPK Miyata
        For i = SxlMidl To UBound(.typ_Param.WFSMP)
            iMidNo = i - SxlMidl + 1
            
            If iMidNo > SXL_MAXSMP Then
                ' 中間抜試最大件数オーバー
                Exit Function
            End If

            'MIDLE側
            sErr_Msg = "WFC総合判定(MIDLE_" & iMidNo & " 初期ﾃﾞｰﾀ設定)"
            typ_in.SAMPLEID = .typ_Param.WFSMP(i).REPSMPLIDCW
            typ_in.WFSMP = typ_CType.typ_Param.WFSMP(i)
        
            '評価測定結果取得
            ReDim Preserve .typ_y013midl_ary(iMidNo)
            sErr_Msg = "WFC総合判定(MIDLE_" & iMidNo & " funWfcGetDataEtc)"
            RET = funWfcGetDataEtc(typ_in, i, tNew_Hinban, iSmpGetFlg, _
                                    .typ_si, _
                                    .typ_y013midl_ary(iMidNo).typ_y013midl, _
                                    sErrMsg)
            If RET = FUNCTION_RETURN_SUCCESS Then

                ' 評価測定結果整列
                sErr_Msg = "WFC総合判定(MIDLE_" & iMidNo & " 評価測定結果整列)"
                If SetMERInd(typ_CType, .typ_y013midl_ary(iMidNo).typ_y013midl, i) <> True Then
                    '評価測定結果整列失敗
                    Exit Function
                End If
                
                '引上げ終了実績取得
                ReDim typ_hi(0)
                sErr_Msg = "WFC総合判定(MIDLE_" & iMidNo & " 引上げ終了実績取得)"
                If s_cmmc001db_Sql(.typ_Param.CRYNUM, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                    '引上げ終了実績取得失敗
                    Exit Function
                Else
                    If UBound(typ_hi) <> 0 Then
                        .typ_hage(i) = typ_hi(1)
                    Else
                        '引上げ終了実績取得失敗
                        SetAllData = FUNCTION_RETURN_FAILURE
                        Exit Function
                    End If
                End If
            Else
                SetAllData = FUNCTION_RETURN_FAILURE
                Exit Function
            End If
        Next i
'Add End   2011/03/07 SMPK Miyata
    End With
    
''2001/12/18 S.Sano Start
    '' Ｐ＋結晶の判断
    sErr_Msg = "WFC総合判定(P+結晶の判断)"
    '2004.09.09 Y.K 紐付け変更（特に取得データは使用していないみたいだが修正した）
'    If DBDRV_GetTBCMH001(records(), "where UPINDNO = '" & Left(SelectSxlID, 7) & "00" & "'") = FUNCTION_RETURN_SUCCESS Then
    If DBDRV_GetTBCMH001(records(), "where UPINDNO = '" & left(SelectSxlID, 7) & "0" & Mid(SelectSxlID, 9, 1) & "'") = FUNCTION_RETURN_SUCCESS Then

'ＷＦサンプル処理変更 2003.05.20 yakimura
'        If Left(SelectSxlID, 1) <> "8" Then
'            bPPlus = ((records(1).AMRESIST <= CDbl(GetCodeField("LG", "02", "P+", "INFO1"))) And (typ_CType.typ_si.HWFTYPE = "P"))
'            bNPlus = ((records(1).AMRESIST <= CDbl(GetCodeField("LG", "02", "N+", "INFO1"))) And (typ_CType.typ_si.HWFTYPE = "N"))
'        Else
'            bPPlus = False
'            bNPlus = False
'        End If
'ＷＦサンプル処理変更 2003.05.20 yakimura
    
    Else
        SetAllData = FUNCTION_RETURN_FAILURE
        Exit Function
    End If
''2001/12/18 S.Sano End
    
    SetAllData = FUNCTION_RETURN_SUCCESS
End Function

'概要      :測定評価結果設定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typ_a         ,IO ,typ_AllTypesC ,各情報構造体
'          :typ_y013()    ,I  ,typ_TBCMY013 ,測定評価結果情報構造体
'          :tt            ,I  ,Integer      ,TOP・TAIL
'          :戻り値        ,O  ,Integer      ,True:正常終了　False:異常終了
'説明      :測定評価結果配列にDB検索したレコードを整列する
'履歴      :
Private Function SetMERInd(typ_CType As typ_AllTypesC, _
                          typ_y013() As typ_TBCMY013, _
                          tt As Integer) As Boolean
    Dim i As Integer
    
    With typ_CType
        For i = 1 To UBound(typ_y013)
            Select Case Trim(typ_y013(i).Spec)
            Case OSWFRES ' RES
                .typ_y013(tt, WFRES) = typ_y013(i)
            Case OSWFOI ' OI
                .typ_y013(tt, WFOI) = typ_y013(i)
            Case OSWFBMD1 ' BMD1
                .typ_y013(tt, WFBMD1) = typ_y013(i)
            Case OSWFBMD2 ' BMD2
                .typ_y013(tt, WFBMD2) = typ_y013(i)
            Case OSWFBMD3 ' BMD3
                .typ_y013(tt, WFBMD3) = typ_y013(i)
            Case OSWFOSF1 ' OSF1
                .typ_y013(tt, WFOSF1) = typ_y013(i)
            Case OSWFOSF2 ' OSF2
                .typ_y013(tt, WFOSF2) = typ_y013(i)
            Case OSWFOSF3 ' OSF3
                .typ_y013(tt, WFOSF3) = typ_y013(i)
'            Del 2010/01/07 SIRD対応 Y.Hitomi
'            Case OSWFOSF4 ' OSF4
'                .typ_y013(tt, WFOSF4) = typ_y013(i)
            Case OSWFDS ' DSOD
                .typ_y013(tt, WFDS) = typ_y013(i)
            Case OSWFDZ ' DZ
                .typ_y013(tt, WFDZ) = typ_y013(i)
            
        ''Upd start 2005/06/21 (TCS)T.Terauchi  SPV9点対応  SPVはSPV実績(TBCMJ016)より取得する為、コメント
'            Case OSWFSP ' SPV
'                .typ_y013(tt, WFSP) = typ_y013(i)
        ''Upd end   2005/06/21 (TCS)T.Terauchi  SPV9点対応  SPVはSPV実績(TBCMJ016)より取得する為、コメント
            
            Case OSWFDOI1 ' DOI1
                .typ_y013(tt, WFDOI1) = typ_y013(i)
            Case OSWFDOI2 ' DOI2
                .typ_y013(tt, WFDOI2) = typ_y013(i)
            Case OSWFDOI3 ' DOI3
                .typ_y013(tt, WFDOI3) = typ_y013(i)
            Case OSWFOT1 ' OT1
                .typ_y013(tt, WFOT1) = typ_y013(i)
            Case OSWFOT2 ' OT2
                .typ_y013(tt, WFOT2) = typ_y013(i)
            ''残存酸素追加　03/12/15 ooba
            Case OSWFAOI ' AOI
                .typ_y013(tt, WFAOI) = typ_y013(i)
            
            'Add 2010/01/07 SIRD対応 Y.Hitomi
            Case OSWFSIRD ' SIRD
                .typ_y013(tt, WFSIRD) = typ_y013(i)

            End Select
        Next
    End With
    SetMERInd = True
End Function

Private Sub SpecJudgCheck()
    Dim IND As String * 4               '検査指示
    Dim c0  As Integer
    
    With typ_CType
'test Git 2014/09/24   

'ＷＦサンプル処理変更 2003.05.20 yakimura
'        JudgSW.rs = (.typ_si.HWFRHWYS = "X")
'        JudgSW.Oi = (.typ_si.HWFONHWS = "X")
'        JudgSW.B1 = (.typ_si.HWFBM1HS = "X")
'        JudgSW.B2 = (.typ_si.HWFBM2HS = "X")
'        JudgSW.B3 = (.typ_si.HWFBM3HS = "X")
'        JudgSW.L1 = (.typ_si.HWFOF1HS = "X")
'        JudgSW.L2 = (.typ_si.HWFOF2HS = "X")
'        JudgSW.L3 = (.typ_si.HWFOF3HS = "X")
'        JudgSW.L4 = (.typ_si.HWFOF4HS = "X")
'        JudgSW.Dsod = (.typ_si.HWFDSOHS = "X")
'        JudgSW.Dz = (.typ_si.HWFMKHWS = "X")
'        JudgSW.Doi1 = (.typ_si.HWFOS1HS = "X")
'        JudgSW.Doi2 = (.typ_si.HWFOS2HS = "X")
'        JudgSW.Doi3 = (.typ_si.HWFOS3HS = "X")
'        JudgSW.sp = (.typ_si.HWFSPVHS = "X") Or (.typ_si.HWFDLHWS = "X")
        
        JudgSW.rs = (.typ_si.HWFRHWYS = "H")
        JudgSW.Oi = (.typ_si.HWFONHWS = "H")
        JudgSW.B1 = (.typ_si.HWFBM1HS = "H")
        JudgSW.B2 = (.typ_si.HWFBM2HS = "H")
        JudgSW.B3 = (.typ_si.HWFBM3HS = "H")
        JudgSW.L1 = (.typ_si.HWFOF1HS = "H")
        JudgSW.L2 = (.typ_si.HWFOF2HS = "H")
        JudgSW.L3 = (.typ_si.HWFOF3HS = "H")
        JudgSW.L4 = (.typ_si.HWFOF4HS = "H")
        JudgSW.Dsod = (.typ_si.HWFDSOHS = "H")
        JudgSW.DZ = (.typ_si.HWFMKHWS = "H")
        JudgSW.Doi1 = (.typ_si.HWFOS1HS = "H")
        JudgSW.Doi2 = (.typ_si.HWFOS2HS = "H")
        JudgSW.Doi3 = (.typ_si.HWFOS3HS = "H")
        JudgSW.sp = (.typ_si.HWFSPVHS = "H") Or (.typ_si.HWFDLHWS = "H")
        JudgSW.AOI = (.typ_si.HWFZOHWS = "H")           '残存酸素追加　03/12/09 ooba
        'GD追加　05/01/27 ooba
        JudgSW.GD = (.typ_si.HWFDENHS = "H") Or (.typ_si.HWFLDLHS = "H") Or _
                    (.typ_si.HWFDVDHS = "H")
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
        JudgSW.B1E = (.typ_si.HEPBM1HS = "H")
        JudgSW.B2E = (.typ_si.HEPBM2HS = "H")
        JudgSW.B3E = (.typ_si.HEPBM3HS = "H")
        JudgSW.L1E = (.typ_si.HEPOF1HS = "H")
        JudgSW.L2E = (.typ_si.HEPOF2HS = "H")
        JudgSW.L3E = (.typ_si.HEPOF3HS = "H")
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
'ＷＦサンプル処理変更 2003.05.20 yakimura
        
'''''        For c0 = 1 To 2
'''''            IND = IIf(c0 = SxlTop, "1346", "2346")
'''''            MeasFlag(c0).B1 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDB1CW) <> 0)
'''''            MeasFlag(c0).B2 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDB2CW) <> 0)
'''''            MeasFlag(c0).B3 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDB3CW) <> 0)
'''''            MeasFlag(c0).Doi1 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDDO1CW) <> 0)
'''''            MeasFlag(c0).Doi2 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDDO2CW) <> 0)
'''''            MeasFlag(c0).Doi3 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDDO3CW) <> 0)
'''''            MeasFlag(c0).Dsod = (InStr(IND, .typ_Param.WFSMP(c0).WFINDDSCW) <> 0)
'''''            MeasFlag(c0).Dz = (InStr(IND, .typ_Param.WFSMP(c0).WFINDDZCW) <> 0)
'''''            MeasFlag(c0).L1 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDL1CW) <> 0)
'''''            MeasFlag(c0).L2 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDL2CW) <> 0)
'''''            MeasFlag(c0).L3 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDL3CW) <> 0)
'''''            MeasFlag(c0).L4 = (InStr(IND, .typ_Param.WFSMP(c0).WFINDL4CW) <> 0)
'''''            MeasFlag(c0).Oi = (InStr(IND, .typ_Param.WFSMP(c0).WFINDOICW) <> 0)
'''''            MeasFlag(c0).rs = (InStr(IND, .typ_Param.WFSMP(c0).WFINDRSCW) <> 0)
'''''            MeasFlag(c0).sp = (InStr(IND, .typ_Param.WFSMP(c0).WFINDSPCW) <> 0)
'''''        Next
        
    End With
End Sub

'概要      :結晶判定(全)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :typ_CType     ,I  ,typ_AllTypesC    ,各情報構造体
'          :tNew_Hinban   ,I  ,tFullHinban      :振替候補品番
'          :tt            ,I  ,Integer          ,TopTail判定用
'説明      :検査指示に従い、実績判定を行う
'履歴      :
Public Function WfAllJudg(typ_CType As typ_AllTypesC, tNew_Hinban As tFullHinban, tt As Integer) As FUNCTION_RETURN
    
    Dim IND         As String * 4                  '検査指示
    Dim bJudg       As Boolean
    Dim i           As Integer
    Dim typTmList() As typ_TBCMB005
'Chg Start 2011/03/09 SMPK Miyata
'    Dim INGOTPOS(2) As Integer
    Dim INGOTPOS(SXL_MAXSMP) As Integer
'Chg End   2011/03/09 SMPK Miyata
    Dim vTemp       As Variant
    Dim sHinban12   As String                               '品番(12桁)
    Dim sSxlPos     As String       'SXL位置(TOP/BOT)　04/04/12 ooba

    i = 0
    WfAllJudg = FUNCTION_RETURN_FAILURE
    
    sHinban12 = tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond
    
    If tt = SxlTop Then
        INGOTPOS(tt) = typ_CType.typ_Param.INGOTPOS
'Chg Start 2011/03/09 SMPK Miyata
'    Else
'        INGOTPOS(tt) = typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH
    ElseIf tt = SxlTail Then
        INGOTPOS(tt) = typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH
    Else
        INGOTPOS(tt) = typ_CType.typ_Param.WFSMP(tt).INPOSCW
'Chg End   2011/03/09 SMPK Miyata
    End If
    
    '検査指示設定
    If tt = SxlTop Then
        IND = "123"
    Else
        IND = "123"
    End If

'Chg Start 2011/03/09 SMPK Miyata
'    sSxlPos = IIf(tt = SxlTop, "TOP", "BOT")        '04/04/12 ooba
    Select Case tt
        Case SxlTop:    sSxlPos = "TOP"
        Case SxlTail:   sSxlPos = "BOT"
        Case Else:      sSxlPos = "MID"
    End Select
'Chg End   2011/03/09 SMPK Miyata
    
    '結晶コードリスト取得
    If GetCodeList(MSYSCLASS, KCLASS, typTmList()) <> FUNCTION_RETURN_SUCCESS Then
        '結晶コードリスト取得失敗
        Exit Function
    End If
    
    With typ_CType
        '' WF検査指示（Rs)*****************************************************************
'        If JudgSW.rs Then
        '保証方法ﾁｪｯｸ追加　04/04/12 ooba
'Cng Start 2011/08/10 Y.Hitomi
'Chg Start 2011/03/10 SMPK Miyata
'        If JudgSW.rs And CheckKHN(.typ_si.HWFRKHNN, 1, sSxlPos) Then
        '通常抜試：保証方法=保証 かつ 検査頻度＿抜チェック有りの場合、仕様有とする
        '中間抜試：保証方法=保証 かつ 中間抜試指示有りかつ　中間抜試フラグが保証の場合、仕様有とする
        If (JudgSW.rs And CheckKHN(.typ_si.HWFRKHNN, 1, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
           (JudgSW.rs And .typ_si.MSMPFLGWFR = "1" And (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3") And _
           (tt >= SxlMidl)) Then
'        If (JudgSW.rs And CheckKHN(.typ_si.HWFRKHNN, 1, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
'           (JudgSW.rs And .typ_si.MSMPFLGWFR = "1" And (tt >= SxlMidl)) Then
'Chg End   2011/03/10 SMPK Miyata
'Cng End  2011/08/10 Y.Hitomi
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDRSCW) <> 0 Then
'                If .typ_Param.WFSMP(tt).WFRESRS1CW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESRS1CW = "1") And (Trim(.typ_y013(tt, WFRES).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
                    '比抵抗判定
                    If WfCrResJudg(typ_CType, .typ_si, .bOKNG(tt), .dblScut(tt), tt) Then
                        JiltusekiUmu(tt, WFRES) = True '2001/12/19 S.Sano
                    End If
                Else
                    ' サンプルが無い場合は、NGとして表示
                    .bOKNG(tt) = False
                End If
                If .bOKNG(tt) = False Then
'Chg Start 2011/03/09 SMPK Miyata
'                    TotalJudg = False
                    If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
                    
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                    gsTbcmy028ErrCode = "00130"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
                End If
            Else 'If .typ_Param.WFSMP(tt).WFRESRS = "2" Then
                ' 指示が無い場合は、NGとして表示
                .bOKNG(tt) = False
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata

'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00130"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
                        
        Else
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDRSCW) <> 0 Then
'                If .typ_Param.WFSMP(tt).WFRESRS1CW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESRS1CW = "1") And (Trim(.typ_y013(tt, WFRES).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
                    '比抵抗判定
                    If WfCrResJudg(typ_CType, .typ_si, .bOKNG(tt), .dblScut(tt), tt) Then
                        JiltusekiUmu(tt, WFRES) = True '2001/12/19 S.Sano
                    End If
                Else
                    ' サンプルが無い場合は、OKとして表示
                    .bOKNG(tt) = True
                End If
                'Add Start 2011/11/28 Y.Hitomi 中間抜試(保証）の場合は、参考表示
                If sSxlPos = "MID" And JudgSW.rs And .bOKNG(tt) = False And _
                   (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3") Then
'                If sSxlPos = "MID" And JudgSW.rs And .bOKNG(tt) = False Then
                    MidlJudg = False
                End If
                'Add End   2011/11/28 Y.Hitomi
            Else
                .bOKNG(tt) = True
            End If
            
        End If

        '' WF検査指示（Oi)*****************************************************************
'        If JudgSW.OI Then
        '保証方法ﾁｪｯｸ追加　04/04/12 ooba
'Chg Start 2011/08/10 Y.Hitomi
'Chg Start 2011/03/10 SMPK Miyata
'        If JudgSW.Oi And CheckKHN(.typ_si.HWFONKHN, 2, sSxlPos) Then
        '通常抜試：保証方法=保証 かつ 検査頻度＿抜チェック有りの場合、仕様有とする
        '中間抜試：保証方法=保証 かつ 中間抜試(保証)の場合、仕様有とする
        If (JudgSW.Oi And CheckKHN(.typ_si.HWFONKHN, 2, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
           (JudgSW.Oi And .typ_si.MSMPFLGWFO = "1" And (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3") And _
           (tt >= SxlMidl)) Then
'        If (JudgSW.Oi And CheckKHN(.typ_si.HWFONKHN, 2, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
'           (JudgSW.Oi And .typ_si.MSMPFLGWFO = "1" And (tt >= SxlMidl)) Then
'Chg End   2011/03/10 SMPK Miyata
'Chg End   2011/08/10 Y.Hitomi
            '画面表示内容設定
            .typ_rslt(tt, i).pos = -1                                               ' 結晶内開始位置
            .typ_rslt(tt, i).NAIYO = Search_CrCode("Oi", typTmList())               ' 内容
            .typ_rslt(tt, i).INFO1 = "仕様有"                                       ' 情報１
            .typ_rslt(tt, i).INFO2 = "検査無"                                       ' 情報２
            .typ_rslt(tt, i).INFO3 = ""                                             ' 情報３
            .typ_rslt(tt, i).INFO4 = ""                                             ' 情報４
        '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        '2.1.3 AN温度 実績反映チェック追加
            '5番目の情報：AN温度を追加
            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
            typ_rslt_ex(tt, i).INFO5 = ""                                             ' 情報５
        '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
        '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
            typ_rslt_ex(tt, i).INFO6 = ""                                           ' 情報6
            typ_rslt_ex(tt, i).INFO7 = ""                                           ' 情報7
            typ_rslt_ex(tt, i).INFO8 = ""                                           ' 情報8
        '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
            .typ_rslt(tt, i).SMPLID = -1                                            ' サンプルＮｏ
            .typ_rslt(tt, i).OKNG = "NG"                                            ' 判定結果
            .typ_rslt(tt, i).hinban = sHinban12                                     ' 品番(12桁)
            bJudg = False
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDOICW) <> 0 Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' 結晶内開始位置
                .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
                .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFOI).SAMPLEID              ' サンプルＮｏ
'                If .typ_Param.WFSMP(tt).WFRESOICW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESOICW = "1") And (Trim(.typ_y013(tt, WFOI).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
                    'OI判定失敗
                    .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報２
                    'OI判定
'Cng Start 2011/08/01 Y.Hitomi
                    If WfCrOiJudg(.typ_si, .typ_y013(tt, WFOI), bJudg, sSxlPos) Then
'                    If WfCrOiJudg(.typ_si, .typ_y013(tt, WFOI), bJudg) Then
'Cng Start 2011/08/01 Y.Hitomi
                        '画面表示内容設定
                        vTemp = CVar(.typ_y013(tt, WFOI).MESDATA1)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' 情報1
                        vTemp = CVar(.typ_y013(tt, WFOI).MESDATA13)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.00")     ' 情報2
                        vTemp = CVar(.typ_y013(tt, WFOI).MESDATA12)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.00")     ' 情報3
                        vTemp = CVar(.typ_y013(tt, WFOI).MESDATA11)
                        'ORGの小数桁数を6桁(7桁目四捨五入)に変更 2011/11/25 SETsw kubota
                        '.typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")     ' 情報4
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.000000")     ' 情報4
                        JiltusekiUmu(tt, WFOI) = True '2001/12/19 S.Sano
                    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    '2.1.3 AN温度 実績反映チェック追加
                        '5番目の情報：AN温度を追加
                        vTemp = CVar(.typ_y013(tt, WFOI).DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESOICW = "2" Then
                    '画面表示内容設定
                    bJudg = False
                    .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                       ' 結晶内開始位置
                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ異常"                            ' 情報３
                    .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                                        ' 判定結果
            Else
                .typ_rslt(tt, i).OKNG = "NG"                                        ' 判定結果
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata

'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00131"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
            
        Else
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDOICW) <> 0 Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' 結晶内開始位置
                .typ_rslt(tt, i).NAIYO = Search_CrCode("Oi", typTmList())           ' 内容
                .typ_rslt(tt, i).INFO1 = "仕様無"                                   ' 情報１
                .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
                .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
                .typ_rslt(tt, i).INFO4 = ""                                         ' 情報４
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '2.1.3 AN温度 実績反映チェック追加
                '5番目の情報：AN温度を追加
                typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                typ_rslt_ex(tt, i).INFO5 = ""        ' 情報5
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
            '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
                typ_rslt_ex(tt, i).INFO6 = ""                                       ' 情報6
                typ_rslt_ex(tt, i).INFO7 = ""                                       ' 情報7
                typ_rslt_ex(tt, i).INFO8 = ""                                       ' 情報8
            '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFOI).SAMPLEID              ' サンプルＮｏ
                .typ_rslt(tt, i).OKNG = "OK"                                        ' 判定結果
                .typ_rslt(tt, i).hinban = sHinban12                                 ' 品番(12桁)
                bJudg = False
'                If .typ_Param.WFSMP(tt).WFRESOICW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESOICW = "1") And (Trim(.typ_y013(tt, WFOI).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
                    'OI判定失敗
                    .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報２
                    'OI判定
'Cng Start 2011/08/01 Y.Hitomi
                    If WfCrOiJudg(.typ_si, .typ_y013(tt, WFOI), bJudg, sSxlPos) Then
'                    If WfCrOiJudg(.typ_si, .typ_y013(tt, WFOI), bJudg) Then
'Cng Start 2011/08/01 Y.Hitomi
                        '画面表示内容設定
                        .typ_rslt(tt, i).NAIYO = Search_CrCode("Oi", typTmList())   ' 内容
                        vTemp = CVar(.typ_y013(tt, WFOI).MESDATA1)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' 情報1
                        vTemp = CVar(.typ_y013(tt, WFOI).MESDATA13)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.00")     ' 情報2
                        vTemp = CVar(.typ_y013(tt, WFOI).MESDATA12)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.00")     ' 情報3
                        vTemp = CVar(.typ_y013(tt, WFOI).MESDATA11)
                        'ORGの小数桁数を6桁(7桁目四捨五入)に変更 2011/11/25 SETsw kubota
                        '.typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")     ' 情報4
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.000000")     ' 情報4
                        JiltusekiUmu(tt, WFOI) = True '2001/12/19 S.Sano
                    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    '2.1.3 AN温度 実績反映チェック追加
                        '5番目の情報：AN温度を追加
                        vTemp = CVar(.typ_y013(tt, WFOI).DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESOICW = "2" Then
                    '画面表示内容設定
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ異常"                            ' 情報３
                    .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                End If
                'Add Start 2011/11/28 Y.Hitomi 中間抜試の場合は、参考表示する
                If sSxlPos = "MID" And JudgSW.Oi And bJudg = False Then
                    .typ_rslt(tt, i).OKNG = "参考"                                  ' 判定結果
                    MidlJudg = False
                End If
                'Add End   2011/11/28 Y.Hitomi
                i = i + 1
            End If
        End If

        '' 結晶検査指示(B1)*****************************************************************
        BMDDataSet 1, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' 結晶検査指示(B2)*****************************************************************
        BMDDataSet 2, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' 結晶検査指示(B3)*****************************************************************
        BMDDataSet 3, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' 結晶検査指示(L1)*****************************************************************
        OSFDataSet 1, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' 結晶検査指示(L2)*****************************************************************
        OSFDataSet 2, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' 結晶検査指示(L3)*****************************************************************
        OSFDataSet 3, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        
        '' 結晶検査指示(L4)*****************************************************************
    'Del 2010/01/07 SIRD対応 Y.Hitomi
'        OSFDataSet 4, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        
        '' WF検査指示（Dsod)*****************************************************************
'        If JudgSW.Dsod Then
        '保証方法ﾁｪｯｸ追加　04/04/12 ooba

'Chg Start 2011/08/10 Y.Hitomi
'Chg Start 2011/03/10 SMPK Miyata
'        If JudgSW.Dsod And CheckKHN(.typ_si.HWFDSOKN, 13, sSxlPos) Then
        '通常抜試：保証方法=保証 かつ 検査頻度＿抜チェック有りの場合、仕様有とする
        '中間抜試：保証方法=保証 かつ 中間抜試=保証の場合、仕様有とする
        If (JudgSW.Dsod And CheckKHN(.typ_si.HWFDSOKN, 13, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
           (JudgSW.Dsod And .typ_si.MSMPFLGWFDS = "1" And (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3") And _
            (tt >= SxlMidl)) Then
'        If (JudgSW.Dsod And CheckKHN(.typ_si.HWFDSOKN, 13, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
'           (JudgSW.Dsod And .typ_si.MSMPFLGWFDS = "1" And (tt >= SxlMidl)) Then
'Chg End   2011/03/10 SMPK Miyata
'Chg End   2011/08/10 Y.Hitomi

            '画面表示内容設定
            .typ_rslt(tt, i).pos = -1              ' 結晶内開始位置
            .typ_rslt(tt, i).NAIYO = Search_CrCode("DS", typTmList())               ' 内容
            .typ_rslt(tt, i).INFO1 = "仕様有"                                       ' 情報１
            .typ_rslt(tt, i).INFO2 = "検査無"                                       ' 情報２
            .typ_rslt(tt, i).INFO3 = ""                                             ' 情報３
            .typ_rslt(tt, i).INFO4 = ""                                             ' 情報４
        '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        '2.1.3 AN温度 実績反映チェック追加
            '5番目の情報：AN温度を追加
            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
            typ_rslt_ex(tt, i).INFO5 = ""                                           ' 情報5
        '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
        '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
            typ_rslt_ex(tt, i).INFO6 = ""                                           ' 情報6
            typ_rslt_ex(tt, i).INFO7 = ""                                           ' 情報7
            typ_rslt_ex(tt, i).INFO8 = ""                                           ' 情報8
        '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
            .typ_rslt(tt, i).SMPLID = -1                                            ' サンプルＮｏ
            .typ_rslt(tt, i).OKNG = "NG"                                            ' 判定結果
            .typ_rslt(tt, i).hinban = sHinban12                                     ' 品番(12桁)
            bJudg = False
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDDSCW) <> 0 Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' 結晶内開始位置
                .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
                .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFDS).SAMPLEID              ' サンプルＮｏ
                'DS判定取得
'                If .typ_Param.WFSMP(tt).WFRESDSCW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESDSCW = "1") And (Trim(.typ_y013(tt, WFDS).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
                    'DS判定失敗
                    .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報２
                    'DS判定取得
                    If WfCrDsodjudg(.typ_si, .typ_y013(tt, WFDS), bJudg) Then
                        '画面表示内容設定
                        vTemp = CVar(.typ_y013(tt, WFDS).MESDATA1)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' 情報１
                        .typ_rslt(tt, i).INFO2 = ""                                 ' 情報２
                        .typ_rslt(tt, i).INFO3 = ""                                 ' 情報３
'                        .typ_rslt(tt, i).INFO4 = ""                                 ' 情報４
                        'DSODﾊﾟﾀｰﾝ表示追加　04/07/28 ooba START ==========================================================================>
                        vTemp = CVar(IIf(Trim(.typ_y013(tt, WFDS).MESDATA4) = "", "-", Trim(.typ_y013(tt, WFDS).MESDATA4)) _
                                        & "  " & IIf(Trim(.typ_y013(tt, WFDS).MESDATA7) = "", "-", Trim(.typ_y013(tt, WFDS).MESDATA7))) & "   "
                        .typ_rslt(tt, i).INFO4 = vTemp                              ' 情報４
                        'DSODﾊﾟﾀｰﾝ表示追加　04/07/28 ooba END ============================================================================>
                        JiltusekiUmu(tt, WFDS) = True '2001/12/19 S.Sano
                    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    '2.1.3 AN温度 実績反映チェック追加
                        '5番目の情報：AN温度を追加
                        vTemp = CVar(.typ_y013(tt, WFDS).DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESDSCW = "2" Then
                    '画面表示内容設定
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ異常"                            ' 情報３
                    .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                                        ' 判定結果
            Else
                .typ_rslt(tt, i).OKNG = "NG"                                        ' 判定結果
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata

'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00143"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
            
        Else
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDDSCW) <> 0 Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' 結晶内開始位置
                .typ_rslt(tt, i).NAIYO = Search_CrCode("DS", typTmList())           ' 内容
                .typ_rslt(tt, i).INFO1 = "仕様無"                                   ' 情報１
                .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
                .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
                .typ_rslt(tt, i).INFO4 = ""                                         ' 情報４
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '2.1.3 AN温度 実績反映チェック追加
                '5番目の情報：AN温度を追加
                typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                typ_rslt_ex(tt, i).INFO5 = ""                                       ' 情報5
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
            '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
                typ_rslt_ex(tt, i).INFO6 = ""                                       ' 情報6
                typ_rslt_ex(tt, i).INFO7 = ""                                       ' 情報7
                typ_rslt_ex(tt, i).INFO8 = ""                                       ' 情報8
            '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFDS).SAMPLEID              ' サンプルＮｏ
                .typ_rslt(tt, i).OKNG = "OK"                                        ' 判定結果
                .typ_rslt(tt, i).hinban = sHinban12                                 ' 品番(12桁)
                bJudg = False
                'DS判定取得
'                If .typ_Param.WFSMP(tt).WFRESDSCW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESDSCW = "1") And (Trim(.typ_y013(tt, WFDS).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
                    'DS判定失敗
                    .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報２
                    If WfCrDsodjudg(.typ_si, .typ_y013(tt, WFDS), bJudg) Then
                        '画面表示内容設定
                        vTemp = CVar(.typ_y013(tt, WFDS).MESDATA1)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' 情報１
                        .typ_rslt(tt, i).INFO2 = ""                                 ' 情報２
                        .typ_rslt(tt, i).INFO3 = ""                                 ' 情報３
                        .typ_rslt(tt, i).INFO4 = ""                                 ' 情報４
                        JiltusekiUmu(tt, WFDS) = True '2001/12/19 S.Sano
                    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    '2.1.3 AN温度 実績反映チェック追加
                        '5番目の情報：AN温度を追加
                        vTemp = CVar(.typ_y013(tt, WFDS).DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESDSCW = "2" Then
                    '画面表示内容設定
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = ""                                     ' 情報３
                    .typ_rslt(tt, i).INFO4 = "ｻﾝﾌﾟﾙ異常"                            ' 情報４
                End If
                'Add Start 2011/11/28 Y.Hitomi 中間抜試の場合は、参考表示する
                If sSxlPos = "MID" And JudgSW.Dsod And bJudg = False Then
                    .typ_rslt(tt, i).OKNG = "参考"
                    MidlJudg = False
                End If
                'Add End   2011/11/28 Y.Hitomi
                i = i + 1
            End If
        End If
        
        
        '' WF検査指示（DZ)*****************************************************************
'        If JudgSW.DZ Then
        '保証方法ﾁｪｯｸ追加　04/04/12 ooba
'Chg Start 2011/08/10 Y.Hitomi
'Chg Start 2011/03/10 SMPK Miyata
'        If JudgSW.DZ And CheckKHN(.typ_si.HWFMKKHN, 14, sSxlPos) Then
        '通常抜試：保証方法=保証 かつ 検査頻度＿抜チェック有りの場合、仕様有とする
        '中間抜試：保証方法=保証 かつ 中間抜試=保証の場合、仕様有とする
        If (JudgSW.DZ And CheckKHN(.typ_si.HWFMKKHN, 14, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
           (JudgSW.DZ And .typ_si.MSMPFLGWFDZ = "1" And (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3") And _
           (tt >= SxlMidl)) Then
'        If (JudgSW.DZ And CheckKHN(.typ_si.HWFMKKHN, 14, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
'           (JudgSW.DZ And .typ_si.MSMPFLGWFDZ = "1" And (tt >= SxlMidl)) Then
'Chg End   2011/03/10 SMPK Miyata
'Chg End   2011/08/10 Y.Hitomi

            '画面表示内容設定
            .typ_rslt(tt, i).pos = -1                                               ' 結晶内開始位置
            .typ_rslt(tt, i).NAIYO = Search_CrCode("DZ", typTmList())               ' 内容
            .typ_rslt(tt, i).INFO1 = "仕様有"                                       ' 情報１
            .typ_rslt(tt, i).INFO2 = "検査無"                                       ' 情報２
            .typ_rslt(tt, i).INFO3 = ""                                             ' 情報３
            .typ_rslt(tt, i).INFO4 = ""                                             ' 情報４
        '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        '2.1.3 AN温度 実績反映チェック追加
            '5番目の情報：AN温度を追加
            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
            typ_rslt_ex(tt, i).INFO5 = ""                                           ' 情報5
'↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
'項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
            typ_rslt_ex(tt, i).INFO6 = ""                                           ' 情報6
            typ_rslt_ex(tt, i).INFO7 = ""                                           ' 情報7
            typ_rslt_ex(tt, i).INFO8 = ""                                           ' 情報8
'↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
        '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            .typ_rslt(tt, i).SMPLID = -1                                            ' サンプルＮｏ
            .typ_rslt(tt, i).OKNG = "NG"                                            ' 判定結果
            .typ_rslt(tt, i).hinban = sHinban12                                     ' 品番(12桁)
            bJudg = False
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDDZCW) <> 0 Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' 結晶内開始位置
                .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
                .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFDZ).SAMPLEID              ' サンプルＮｏ
                'DZ判定取得
'                If .typ_Param.WFSMP(tt).WFRESDZCW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESDZCW = "1") And (Trim(.typ_y013(tt, WFDZ).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
                    'DZ判定取得
                    If WfCrDzjudg(.typ_si, .typ_y013(tt, WFDZ), bJudg) Then
                        'DZ判定失敗
                        .typ_rslt(tt, i).INFO3 = "判定Err"                          ' 情報２
                        '画面表示内容設定
                        vTemp = CVar(.typ_y013(tt, WFDZ).MESDATA5)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.0")      ' 情報1
                        vTemp = CVar(.typ_y013(tt, WFDZ).MESDATA6)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.0")      ' 情報2
                        vTemp = CVar(.typ_y013(tt, WFDZ).MESDATA7)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.0")      ' 情報3
                        .typ_rslt(tt, i).INFO4 = ""                                 ' 情報４
                        JiltusekiUmu(tt, WFDZ) = True '2001/12/19 S.Sano
                    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    '2.1.3 AN温度 実績反映チェック追加
                        '5番目の情報：AN温度を追加
                        vTemp = CVar(.typ_y013(tt, WFDZ).DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESDZCW = "2" Then
                    '画面表示内容設定
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ異常"                            ' 情報３
                    .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                                        ' 判定結果
            Else
                .typ_rslt(tt, i).OKNG = "NG"                                        ' 判定結果
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00144"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
            
        Else
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDDZCW) <> 0 Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' 結晶内開始位置
                .typ_rslt(tt, i).NAIYO = Search_CrCode("DZ", typTmList())           ' 内容
                .typ_rslt(tt, i).INFO1 = "仕様無"                                   ' 情報１
                .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
                .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
                .typ_rslt(tt, i).INFO4 = ""                                         ' 情報４
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '2.1.3 AN温度 実績反映チェック追加
                '5番目の情報：AN温度を追加
                typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                typ_rslt_ex(tt, i).INFO5 = ""                                       ' 情報5
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
            '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
                typ_rslt_ex(tt, i).INFO6 = ""                                       ' 情報6
                typ_rslt_ex(tt, i).INFO7 = ""                                       ' 情報7
                typ_rslt_ex(tt, i).INFO8 = ""                                       ' 情報8
            '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFDZ).SAMPLEID              ' サンプルＮｏ
                .typ_rslt(tt, i).OKNG = "OK"                                        ' 判定結果
                .typ_rslt(tt, i).hinban = sHinban12                                 ' 品番(12桁)
                bJudg = False
                'DZ判定取得
'                If .typ_Param.WFSMP(tt).WFRESDZCW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESDZCW = "1") And (Trim(.typ_y013(tt, WFDZ).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
                    'DZ判定失敗
                    .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報２
                    'DZ判定取得
                    If WfCrDzjudg(.typ_si, .typ_y013(tt, WFDZ), bJudg) Then
                        '画面表示内容設定
                        vTemp = CVar(.typ_y013(tt, WFDZ).MESDATA5)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.0")      ' 情報1
                        vTemp = CVar(.typ_y013(tt, WFDZ).MESDATA6)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.0")      ' 情報2
                        vTemp = CVar(.typ_y013(tt, WFDZ).MESDATA7)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.0")      ' 情報3
                        .typ_rslt(tt, i).INFO4 = ""                                 ' 情報４
                        JiltusekiUmu(tt, WFDZ) = True '2001/12/19 S.Sano
                    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    '2.1.3 AN温度 実績反映チェック追加
                        '5番目の情報：AN温度を追加
                        vTemp = CVar(.typ_y013(tt, WFDZ).DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESDZCW = "2" Then
                    '画面表示内容設定
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ異常"                            ' 情報３
                    .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                End If
                'Add Start 2011/11/28 Y.Hitomi 中間抜試の場合は、参考表示する
                If sSxlPos = "MID" And JudgSW.DZ And bJudg = False Then
                    .typ_rslt(tt, i).OKNG = "参考"
                    MidlJudg = False
                End If
                'Add End   2011/11/28 Y.Hitomi
                i = i + 1
            End If
        End If
        
                
    ''Upd start 2005/06/21 (TCS)t.terauchi      SPV9点対応  Fe濃度・拡散長に分けて表示

'        '' WF検査指示（SP)*****************************************************************
'        '保証方法ﾁｪｯｸ追加　04/04/12 ooba
'        JudgSW.sp = ((.typ_si.HWFSPVHS = "H") And CheckKHN(.typ_si.HWFSPVKN, 15, sSxlPos)) _
'                    Or ((.typ_si.HWFDLHWS = "H") And CheckKHN(.typ_si.HWFDLKHN, 16, sSxlPos))
'        If JudgSW.sp Then
'
'            '画面表示内容設定
'            .typ_rslt(tt, i).pos = -1                                               ' 結晶内開始位置
'            .typ_rslt(tt, i).NAIYO = Search_CrCode("SP", typTmList())               ' 内容
'            .typ_rslt(tt, i).INFO1 = "仕様有"                                       ' 情報１
'            .typ_rslt(tt, i).INFO2 = "検査無"                                       ' 情報２
'            .typ_rslt(tt, i).INFO3 = ""                                             ' 情報３
'            .typ_rslt(tt, i).INFO4 = ""                                             ' 情報４
'            .typ_rslt(tt, i).SMPLID = -1                                            ' サンプルＮｏ
'            .typ_rslt(tt, i).OKNG = "NG"                                            ' 判定結果
'            .typ_rslt(tt, i).hinban = sHinban12                                     ' 品番(12桁)
'            bJudg = False
'            If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 Then
'                '画面表示内容設定
'                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' 結晶内開始位置
'                .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
'                .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
'                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFSP).SAMPLEID              ' サンプルＮｏ
'                'SP判定取得
''                If .typ_Param.WFSMP(tt).WFRESSPCW = "1" Then
'                If (.typ_Param.WFSMP(tt).WFRESSPCW = "1") And (Trim(.typ_y013(tt, WFSP).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
'                    'LT判定失敗
'                    .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報２
'                    'SP判定取得
'                    If WfCrSpvjudg(.typ_si, .typ_y013(tt, WFSP), bJudg) Then
'                        '画面表示内容設定
'                        vTemp = CVar(.typ_y013(tt, WFSP).MESDATA5)
'                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' 情報１
'                        vTemp = CVar(.typ_y013(tt, WFSP).MESDATA4)
'                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.00")     ' 情報２
'                        vTemp = CVar(.typ_y013(tt, WFSP).MESDATA3)
'                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0")        ' 情報３
'                        vTemp = CVar(.typ_y013(tt, WFSP).MESDATA2)
'                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0")        ' 情報４
'                        JiltusekiUmu(tt, WFSP) = True '2001/12/19 S.Sano
'                    End If
'                ElseIf .typ_Param.WFSMP(tt).WFRESSPCW = "2" Then
'                    '画面表示内容設定
'                    bJudg = False
'                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ異常"                            ' 情報３
'                    .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
'                End If
'            End If
'            If bJudg = True Then
'                .typ_rslt(tt, i).OKNG = "OK"                                        ' 判定結果
'            Else
'                .typ_rslt(tt, i).OKNG = "NG"                                        ' 判定結果
'                TotalJudg = False
'            End If
'            i = i + 1
'
'        Else
'            If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 And .typ_Param.WFSMP(tt).WFRESSPCW = "1" Then
'
'                '画面表示内容設定
'                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' 結晶内開始位置
'                .typ_rslt(tt, i).NAIYO = Search_CrCode("SP", typTmList())           ' 内容
'                .typ_rslt(tt, i).INFO1 = "仕様無"                                   ' 情報１
'                .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
'                .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
'                .typ_rslt(tt, i).INFO4 = ""                                         ' 情報４
'                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFSP).SAMPLEID              ' サンプルＮｏ
'                .typ_rslt(tt, i).OKNG = "OK"                                        ' 判定結果
'                .typ_rslt(tt, i).hinban = sHinban12                                 ' 品番(12桁)
'                'SP判定取得
''                If .typ_Param.WFSMP(tt).WFRESSPCW = "1" Then
'                If (.typ_Param.WFSMP(tt).WFRESSPCW = "1") And (Trim(.typ_y013(tt, WFSP).SAMPLEID) <> "0") Then      '2003/12/19 SystemBrain
'                    'LT判定失敗
'                    .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報２
'                    'SP判定取得
'                    If WfCrSpvjudg(.typ_si, .typ_y013(tt, WFSP), bJudg) Then
'                        '画面表示内容設定
'                        vTemp = CVar(.typ_y013(tt, WFSP).MESDATA5)
'                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' 情報１
'                        vTemp = CVar(.typ_y013(tt, WFSP).MESDATA4)
'                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.00")     ' 情報２
'                        vTemp = CVar(.typ_y013(tt, WFSP).MESDATA3)
'                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0")        ' 情報３
'                        vTemp = CVar(.typ_y013(tt, WFSP).MESDATA2)
'                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0")        ' 情報４
'                        JiltusekiUmu(tt, WFSP) = True '2001/12/19 S.Sano
'                    End If
'                ElseIf .typ_Param.WFSMP(tt).WFRESSPCW = "2" Then
'                    '画面表示内容設定
'                    bJudg = False
'                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ異常"                            ' 情報３
'                    .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
'                End If
'                i = i + 1
'            End If
'        End If

        ''Fe濃度***************
'Chg Start 2011/03/10 SMPK Miyata
'        JudgSW.sp = (.typ_si.HWFSPVHS = "H") And CheckKHN(.typ_si.HWFSPVKN, 15, sSxlPos)
        JudgSW.sp = (.typ_si.HWFSPVHS = "H") And CheckKHN(.typ_si.HWFSPVKN, 15, sSxlPos) _
                    And (tt = SxlTop Or tt = SxlTail)
'Chg End   2011/03/10 SMPK Miyata
        
        If JudgSW.sp Then

            '画面表示内容設定
            .typ_rslt(tt, i).pos = -1                                               ' 結晶内開始位置
            .typ_rslt(tt, i).NAIYO = Search_CrCode("SPFE", typTmList())             ' 内容
            .typ_rslt(tt, i).INFO1 = "仕様有"                                       ' 情報１
            .typ_rslt(tt, i).INFO2 = "検査無"                                       ' 情報２
            .typ_rslt(tt, i).INFO3 = ""                                             ' 情報３
            .typ_rslt(tt, i).INFO4 = ""                                             ' 情報４
        '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        '2.1.3 AN温度 実績反映チェック追加
            '5番目の情報：AN温度を追加
            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
            typ_rslt_ex(tt, i).INFO5 = ""                                           ' 情報5
        '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
        '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
            typ_rslt_ex(tt, i).INFO6 = ""                                           ' 情報6
            typ_rslt_ex(tt, i).INFO7 = ""                                           ' 情報7
            typ_rslt_ex(tt, i).INFO8 = ""                                           ' 情報8
        '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
            .typ_rslt(tt, i).SMPLID = -1                                            ' サンプルＮｏ
            .typ_rslt(tt, i).OKNG = "NG"                                            ' 判定結果
            .typ_rslt(tt, i).hinban = sHinban12                                     ' 品番(12桁)
            bJudg = False
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' 結晶内開始位置
                .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
                .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
                .typ_rslt(tt, i).SMPLID = Trim(typ_J016_WFSPVJudg(tt).SMPLNO)       ' サンプルＮｏ
                'SP判定取得
                If (.typ_Param.WFSMP(tt).WFRESSPCW = "1") And (Trim(typ_J016_WFSPVJudg(tt).SMPLNO) <> "0") Then
                    'SP判定失敗
                    .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報２
                    'SP判定取得
                    If WfCrSpvJudg_New(.typ_si, typ_J016_WFSPVJudg(tt), bJudg, 1, sSxlPos) Then
                        '画面表示内容設定
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).MAX_FE)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' 情報１
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).MIN_FE)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.00")     ' 情報２
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).AVE_FE)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.00")     ' 情報３
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).CENTER_FE)
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")     ' 情報４
                    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    '2.1.3 AN温度 実績反映チェック追加
                        '5番目の情報：AN温度を追加
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                    '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).PUA_FE)
                        typ_rslt_ex(tt, i).INFO6 = DBData2DispData(vTemp, "0.00")     ' 情報6
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).PUAP_FE)
                        typ_rslt_ex(tt, i).INFO7 = DBData2DispData(vTemp, "0.000")    ' 情報7
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).STD_FE)
                        typ_rslt_ex(tt, i).INFO8 = DBData2DispData(vTemp, "0.000")    ' 情報8
                    '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESSPCW = "2" Then
                    '画面表示内容設定
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ異常"                            ' 情報３
                    .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                                        ' 判定結果
            Else
                .typ_rslt(tt, i).OKNG = "NG"                                        ' 判定結果
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00145"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
        Else
            'If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 And .typ_Param.WFSMP(tt).WFRESSPCW = "1" Then
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 Then
                
                ''実績がある時のみ表示する
                '↓変更 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
'                If typ_J016_WFSPVJudg(tt).MAX_FE <> -1 _
'                    Or typ_J016_WFSPVJudg(tt).MIN_FE <> -1 _
'                    Or typ_J016_WFSPVJudg(tt).AVE_FE <> -1 _
'                    Or typ_J016_WFSPVJudg(tt).CENTER_FE <> -1 Then
                If typ_J016_WFSPVJudg(tt).MAX_FE <> -1 _
                    Or typ_J016_WFSPVJudg(tt).MIN_FE <> -1 _
                    Or typ_J016_WFSPVJudg(tt).AVE_FE <> -1 _
                    Or typ_J016_WFSPVJudg(tt).CENTER_FE <> -1 _
                    Or typ_J016_WFSPVJudg(tt).PUA_FE <> -1 _
                    Or typ_J016_WFSPVJudg(tt).PUAP_FE <> -1 _
                    Or typ_J016_WFSPVJudg(tt).STD_FE <> -1 Then
                '↑変更 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                    
                    '画面表示内容設定
                    .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' 結晶内開始位置
                    .typ_rslt(tt, i).NAIYO = Search_CrCode("SPFE", typTmList())           ' 内容
                    .typ_rslt(tt, i).INFO1 = "仕様無"                                   ' 情報１
                    .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
                    .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
                    .typ_rslt(tt, i).INFO4 = ""                                         ' 情報４
                '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                '2.1.3 AN温度 実績反映チェック追加
                    '5番目の情報：AN温度を追加
                    typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                    typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                    typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                    typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                    typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                    typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                    typ_rslt_ex(tt, i).INFO5 = ""                                       ' 情報5
                '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
                    typ_rslt_ex(tt, i).INFO6 = ""                                       ' 情報6
                    typ_rslt_ex(tt, i).INFO7 = ""                                       ' 情報7
                    typ_rslt_ex(tt, i).INFO8 = ""                                       ' 情報8
                '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                    .typ_rslt(tt, i).SMPLID = Trim(typ_J016_WFSPVJudg(tt).SMPLNO)       ' サンプルＮｏ
                    .typ_rslt(tt, i).OKNG = "OK"                                        ' 判定結果
                    .typ_rslt(tt, i).hinban = sHinban12                                 ' 品番(12桁)
                    'SP判定取得
                    If (.typ_Param.WFSMP(tt).WFRESSPCW = "1") And (Trim(typ_J016_WFSPVJudg(tt).SMPLNO) <> "0") Then      '2003/12/19 SystemBrain
                        'SP判定失敗
                        .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報２
                        'SP判定取得
                        If WfCrSpvJudg_New(.typ_si, typ_J016_WFSPVJudg(tt), bJudg, 1, sSxlPos) Then
                            '画面表示内容設定
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).MAX_FE)
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")     ' 情報１
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).MIN_FE)
                            .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.00")     ' 情報２
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).AVE_FE)
                            .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.00")     ' 情報３
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).CENTER_FE)
                            .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")     ' 情報４
                        '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                        '2.1.3 AN温度 実績反映チェック追加
                            '5番目の情報：AN温度を追加
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).DKAN)
                            '3～6桁目がAN温度
                            vTemp = Mid(vTemp, 3, 4)
                            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                            typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                        '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                        '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                        '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).PUA_FE)
                            typ_rslt_ex(tt, i).INFO6 = DBData2DispData(vTemp, "0.00")     ' 情報6
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).PUAP_FE)
                            typ_rslt_ex(tt, i).INFO7 = DBData2DispData(vTemp, "0.000")    ' 情報7
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).STD_FE)
                            typ_rslt_ex(tt, i).INFO8 = DBData2DispData(vTemp, "0.000")    ' 情報8
                        '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                        End If
                    ElseIf .typ_Param.WFSMP(tt).WFRESSPCW = "2" Then
                        '画面表示内容設定
                        bJudg = False
                        .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ異常"                            ' 情報３
                        .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                    End If
                    i = i + 1
                End If
            End If
        End If

    ''拡散長***************
'Chg Start 2011/03/10 SMPK Miyata
'        JudgSW.sp = (.typ_si.HWFDLHWS = "H") And CheckKHN(.typ_si.HWFDLKHN, 16, sSxlPos)
        JudgSW.sp = (.typ_si.HWFDLHWS = "H") And CheckKHN(.typ_si.HWFDLKHN, 16, sSxlPos) _
                    And (tt = SxlTop Or tt = SxlTail)
'Chg End   2011/03/10 SMPK Miyata

        If JudgSW.sp Then
            '画面表示内容設定
            .typ_rslt(tt, i).pos = -1                                               ' 結晶内開始位置
            .typ_rslt(tt, i).NAIYO = Search_CrCode("SPKL", typTmList())             ' 内容
            .typ_rslt(tt, i).INFO1 = "仕様有"                                       ' 情報１
            .typ_rslt(tt, i).INFO2 = "検査無"                                       ' 情報２
            .typ_rslt(tt, i).INFO3 = ""                                             ' 情報３
            .typ_rslt(tt, i).INFO4 = ""                                             ' 情報４
        '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        '2.1.3 AN温度 実績反映チェック追加
            '5番目の情報：AN温度を追加
            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
            typ_rslt_ex(tt, i).INFO5 = ""                                           ' 情報5
        '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
        '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
            typ_rslt_ex(tt, i).INFO6 = ""                                           ' 情報6
            typ_rslt_ex(tt, i).INFO7 = ""                                           ' 情報7
            typ_rslt_ex(tt, i).INFO8 = ""                                           ' 情報8
        '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
            .typ_rslt(tt, i).SMPLID = -1                                            ' サンプルＮｏ
            .typ_rslt(tt, i).OKNG = "NG"                                            ' 判定結果
            .typ_rslt(tt, i).hinban = sHinban12                                     ' 品番(12桁)
            bJudg = False
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' 結晶内開始位置
                .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
                .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
                .typ_rslt(tt, i).SMPLID = Trim(typ_J016_WFSPVJudg(tt).SMPLNO)       ' サンプルＮｏ
                
                If (.typ_Param.WFSMP(tt).WFRESSPCW = "1") And (Trim(typ_J016_WFSPVJudg(tt).SMPLNO) <> "0") Then      '2003/12/19 SystemBrain
                    'SP判定失敗
                    .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報２
                    'SP判定取得
                    If WfCrSpvJudg_New(.typ_si, typ_J016_WFSPVJudg(tt), bJudg, 2, sSxlPos) Then
                        '画面表示内容設定
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).MAX_DIFF)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.0")      ' 情報１
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).MIN_DIFF)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.0")      ' 情報２
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).AVE_DIFF)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.0")      ' 情報３
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).CENTER_DIFF)
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.0")      ' 情報４
                    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    '2.1.3 AN温度 実績反映チェック追加
                        '5番目の情報：AN温度を追加
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                    '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).PUA_DIFF)
                        typ_rslt_ex(tt, i).INFO6 = DBData2DispData(vTemp, "0.00")     ' 情報6
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).PUAP_DIFF)
                        typ_rslt_ex(tt, i).INFO7 = DBData2DispData(vTemp, "0.000")    ' 情報7
''                        vTemp = CVar(typ_J016_WFSPVJudg(tt).SPV_Fe_STD)
''                        typ_rslt_ex(tt, i).INFO8 = DBData2DispData(vTemp, "0.000")    ' 情報8
                    '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESSPCW = "2" Then
                    '画面表示内容設定
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ異常"                            ' 情報３
                    .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                                        ' 判定結果
            Else
                .typ_rslt(tt, i).OKNG = "NG"                                        ' 判定結果
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00145"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
        
        Else
            'If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 And .typ_Param.WFSMP(tt).WFRESSPCW = "1" Then
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 Then
                
                ''実績がある時のみ、表示する
                '↓変更 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
'                If typ_J016_WFSPVJudg(tt).MAX_DIFF <> -1 _
'                    Or typ_J016_WFSPVJudg(tt).MIN_DIFF <> -1 _
'                    Or typ_J016_WFSPVJudg(tt).AVE_DIFF <> -1 _
'                    Or typ_J016_WFSPVJudg(tt).CENTER_DIFF <> -1 Then
                If typ_J016_WFSPVJudg(tt).MAX_DIFF <> -1 _
                    Or typ_J016_WFSPVJudg(tt).MIN_DIFF <> -1 _
                    Or typ_J016_WFSPVJudg(tt).AVE_DIFF <> -1 _
                    Or typ_J016_WFSPVJudg(tt).CENTER_DIFF <> -1 _
                    Or typ_J016_WFSPVJudg(tt).PUA_DIFF <> -1 _
                    Or typ_J016_WFSPVJudg(tt).PUAP_DIFF <> -1 Then
                '↑変更 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                
                    '画面表示内容設定
                    .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' 結晶内開始位置
                    .typ_rslt(tt, i).NAIYO = Search_CrCode("SPKL", typTmList())         ' 内容
                    .typ_rslt(tt, i).INFO1 = "仕様無"                                   ' 情報１
                    .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
                    .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
                    .typ_rslt(tt, i).INFO4 = ""                                         ' 情報４
                '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                '2.1.3 AN温度 実績反映チェック追加
                    '5番目の情報：AN温度を追加
                    typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                    typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                    typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                    typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                    typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                    typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                    typ_rslt_ex(tt, i).INFO5 = ""                                       ' 情報5
                '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
                    typ_rslt_ex(tt, i).INFO6 = ""                                       ' 情報6
                    typ_rslt_ex(tt, i).INFO7 = ""                                       ' 情報7
                    typ_rslt_ex(tt, i).INFO8 = ""                                       ' 情報8
                '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                    .typ_rslt(tt, i).SMPLID = Trim(typ_J016_WFSPVJudg(tt).SMPLNO)       ' サンプルＮｏ
                    .typ_rslt(tt, i).OKNG = "OK"                                        ' 判定結果
                    .typ_rslt(tt, i).hinban = sHinban12                                 ' 品番(12桁)
                    'SP判定取得
                    If (.typ_Param.WFSMP(tt).WFRESSPCW = "1") And (Trim(typ_J016_WFSPVJudg(tt).SMPLNO) <> "0") Then      '2003/12/19 SystemBrain
                        'SP判定失敗
                        .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報２
                        'SP判定取得
                        If WfCrSpvJudg_New(.typ_si, typ_J016_WFSPVJudg(tt), bJudg, 2, sSxlPos) Then
                            '画面表示内容設定
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).MAX_DIFF)
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.0")      ' 情報１
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).MIN_DIFF)
                            .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.0")      ' 情報２
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).AVE_DIFF)
                            .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.0")      ' 情報３
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).CENTER_DIFF)
                            .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.0")      ' 情報４
                        '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                        '2.1.3 AN温度 実績反映チェック追加
                            '5番目の情報：AN温度を追加
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).DKAN)
                            '3～6桁目がAN温度
                            vTemp = Mid(vTemp, 3, 4)
                            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                            typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                        '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                        '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                        '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).PUA_DIFF)
                            typ_rslt_ex(tt, i).INFO6 = DBData2DispData(vTemp, "0.00")     ' 情報6
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).PUAP_DIFF)
                            typ_rslt_ex(tt, i).INFO7 = DBData2DispData(vTemp, "0.000")    ' 情報7
''                            vTemp = CVar(typ_J016_WFSPVJudg(tt).SPV_Fe_STD)
''                            typ_rslt_ex(tt, i).INFO8 = DBData2DispData(vTemp, "0.000")    ' 情報8
                        '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                        End If
                    ElseIf .typ_Param.WFSMP(tt).WFRESSPCW = "2" Then
                        '画面表示内容設定
                        bJudg = False
                        .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ異常"                            ' 情報３
                        .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                    End If
                    i = i + 1
                End If
            End If
        End If
        
    ''Upd end  2005/06/21 (TCS)t.terauchi      SPV9点対応   Fe濃度・拡散長に分けて表示


'↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
'Nr濃度(OtherRecords)の実績表示行の追加による変更

        ''Nr濃度***************
'Chg Start 2011/03/10 SMPK Miyata
'        JudgSW.sp = (.typ_si.HWFNRHS = "H") And CheckKHN(.typ_si.HWFNRKN, 19, sSxlPos)
        JudgSW.sp = (.typ_si.HWFNRHS = "H") And CheckKHN(.typ_si.HWFNRKN, 19, sSxlPos) _
                    And (tt = SxlTop Or tt = SxlTail)
'Chg End   2011/03/10 SMPK Miyata

        If JudgSW.sp Then
            '画面表示内容設定
            .typ_rslt(tt, i).pos = -1                                               ' 結晶内開始位置
            .typ_rslt(tt, i).NAIYO = Search_CrCode("SPNR", typTmList())             ' 内容
            .typ_rslt(tt, i).INFO1 = "仕様有"                                       ' 情報１
            .typ_rslt(tt, i).INFO2 = "検査無"                                       ' 情報２
            .typ_rslt(tt, i).INFO3 = ""                                             ' 情報３
            .typ_rslt(tt, i).INFO4 = ""                                             ' 情報４
        '2.1.3 AN温度 実績反映チェック追加
            '5番目の情報：AN温度を追加
            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
            typ_rslt_ex(tt, i).INFO5 = ""                                           ' 情報5
            typ_rslt_ex(tt, i).INFO6 = ""                                           ' 情報6
            typ_rslt_ex(tt, i).INFO7 = ""                                           ' 情報7
            typ_rslt_ex(tt, i).INFO8 = ""                                           ' 情報8
            .typ_rslt(tt, i).SMPLID = -1                                            ' サンプルＮｏ
            .typ_rslt(tt, i).OKNG = "NG"                                            ' 判定結果
            .typ_rslt(tt, i).hinban = sHinban12                                     ' 品番(12桁)
            bJudg = False
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' 結晶内開始位置
                .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
                .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
                .typ_rslt(tt, i).SMPLID = Trim(typ_J016_WFSPVJudg(tt).SMPLNO)       ' サンプルＮｏ
                
                If (.typ_Param.WFSMP(tt).WFRESSPCW = "1") And (Trim(typ_J016_WFSPVJudg(tt).SMPLNO) <> "0") Then      '2003/12/19 SystemBrain
                    'SP判定失敗
                    .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報２
                    'SP判定取得
                    If WfCrSpvJudg_New(.typ_si, typ_J016_WFSPVJudg(tt), bJudg, 3, sSxlPos) Then
                        '画面表示内容設定
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).MAX_NR)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")      ' 情報１
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).MIN_NR)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.00")      ' 情報２
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).AVE_NR)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.00")      ' 情報３
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).CENTER_NR)
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")      ' 情報４
                    '2.1.3 AN温度 実績反映チェック
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).PUA_NR)
                        typ_rslt_ex(tt, i).INFO6 = DBData2DispData(vTemp, "0.00")     ' 情報6
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).PUAP_NR)
                        typ_rslt_ex(tt, i).INFO7 = DBData2DispData(vTemp, "0.000")    ' 情報7
                        vTemp = CVar(typ_J016_WFSPVJudg(tt).STD_NR)
                        typ_rslt_ex(tt, i).INFO8 = DBData2DispData(vTemp, "0.000")    ' 情報8
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESSPCW = "2" Then
                    '画面表示内容設定
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ異常"                            ' 情報３
                    .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                                        ' 判定結果
            Else
                .typ_rslt(tt, i).OKNG = "NG"                                        ' 判定結果
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00145"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
        
        Else
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDSPCW) <> 0 Then
                
                ''実績がある時のみ、表示する
                If typ_J016_WFSPVJudg(tt).MAX_NR <> -1 _
                    Or typ_J016_WFSPVJudg(tt).MIN_NR <> -1 _
                    Or typ_J016_WFSPVJudg(tt).AVE_NR <> -1 _
                    Or typ_J016_WFSPVJudg(tt).CENTER_NR <> -1 _
                    Or typ_J016_WFSPVJudg(tt).PUA_NR <> -1 _
                    Or typ_J016_WFSPVJudg(tt).PUAP_NR <> -1 _
                    Or typ_J016_WFSPVJudg(tt).STD_NR <> -1 Then
                
                    '画面表示内容設定
                    .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' 結晶内開始位置
                    .typ_rslt(tt, i).NAIYO = Search_CrCode("SPNR", typTmList())         ' 内容
                    .typ_rslt(tt, i).INFO1 = "仕様無"                                   ' 情報１
                    .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
                    .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
                    .typ_rslt(tt, i).INFO4 = ""                                         ' 情報４
                    '5番目の情報：AN温度
                    typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                    typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                    typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                    typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                    typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                    typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                    typ_rslt_ex(tt, i).INFO5 = ""                                       ' 情報5
                    typ_rslt_ex(tt, i).INFO6 = ""                                       ' 情報6
                    typ_rslt_ex(tt, i).INFO7 = ""                                       ' 情報7
                    typ_rslt_ex(tt, i).INFO8 = ""                                       ' 情報8
                    .typ_rslt(tt, i).SMPLID = Trim(typ_J016_WFSPVJudg(tt).SMPLNO)       ' サンプルＮｏ
                    .typ_rslt(tt, i).OKNG = "OK"                                        ' 判定結果
                    .typ_rslt(tt, i).hinban = sHinban12                                 ' 品番(12桁)
                    'SP判定取得
                    If (.typ_Param.WFSMP(tt).WFRESSPCW = "1") And (Trim(typ_J016_WFSPVJudg(tt).SMPLNO) <> "0") Then
                        'SP判定失敗
                        .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報２
                        'SP判定取得
                        If WfCrSpvJudg_New(.typ_si, typ_J016_WFSPVJudg(tt), bJudg, 3, sSxlPos) Then
                            '画面表示内容設定
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).MAX_NR)
                            .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.00")      ' 情報１
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).MIN_NR)
                            .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.00")      ' 情報２
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).AVE_NR)
                            .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.00")      ' 情報３
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).CENTER_NR)
                            .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0.00")      ' 情報４
                            '5番目の情報：AN温度
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).DKAN)
                            '3～6桁目がAN温度
                            vTemp = Mid(vTemp, 3, 4)
                            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                            typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).PUA_NR)
                            typ_rslt_ex(tt, i).INFO6 = DBData2DispData(vTemp, "0.00")     ' 情報6
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).PUAP_NR)
                            typ_rslt_ex(tt, i).INFO7 = DBData2DispData(vTemp, "0.000")    ' 情報7
                            vTemp = CVar(typ_J016_WFSPVJudg(tt).STD_NR)
                            typ_rslt_ex(tt, i).INFO8 = DBData2DispData(vTemp, "0.000")    ' 情報8
                        End If
                    ElseIf .typ_Param.WFSMP(tt).WFRESSPCW = "2" Then
                        '画面表示内容設定
                        bJudg = False
                        .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ異常"                            ' 情報３
                        .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                    End If
                    i = i + 1
                End If
            End If
        End If
        
'↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------


        '' 結晶検査指示(DOI1)*****************************************************************
        DOIDataSet 1, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' 結晶検査指示(DOI2)*****************************************************************
        DOIDataSet 2, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' 結晶検査指示(DOI3)*****************************************************************
        DOIDataSet 3, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        
        
        ''残存酸素実績判定/表示処理追加　03/12/09 ooba START ====================================>

        '' 結晶検査指示(AOI)******************************************************************
'        If JudgSW.AOI Then
        '保証方法ﾁｪｯｸ追加　04/04/12 ooba
'Chg Start 2011/08/10 Y.Hitomi
'Chg Start 2011/03/10 SMPK Miyata
'        If JudgSW.AOI And CheckKHN(.typ_si.HWFZOKHN, 17, sSxlPos) Then
        '通常抜試：保証方法=保証 かつ 検査頻度＿抜チェック有りの場合、仕様有とする
        '中間抜試：保証方法=保証 かつ 中間抜試=保証の場合、仕様有とする
        If (JudgSW.AOI And CheckKHN(.typ_si.HWFZOKHN, 17, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
           (JudgSW.AOI And .typ_si.MSMPFLGWFAOI = "1" And (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3") And _
           (tt >= SxlMidl)) Then
'Chg End   2011/03/10 SMPK Miyata
'Chg End   2011/08/10 Y.Hitomi
            
            '画面表示内容設定
            .typ_rslt(tt, i).pos = -1                                               ' 結晶内開始位置
            .typ_rslt(tt, i).NAIYO = Search_CrCode("AO", typTmList())               ' 内容
            .typ_rslt(tt, i).INFO1 = "仕様有"                                       ' 情報１
            .typ_rslt(tt, i).INFO2 = "検査無"                                       ' 情報２
            .typ_rslt(tt, i).INFO3 = ""                                             ' 情報３
            .typ_rslt(tt, i).INFO4 = ""                                             ' 情報４
        '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        '2.1.3 AN温度 実績反映チェック追加
            '5番目の情報：AN温度を追加
            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
            typ_rslt_ex(tt, i).INFO5 = ""                                           ' 情報5
        '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
        '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
            typ_rslt_ex(tt, i).INFO6 = ""                                           ' 情報6
            typ_rslt_ex(tt, i).INFO7 = ""                                           ' 情報7
            typ_rslt_ex(tt, i).INFO8 = ""                                           ' 情報8
        '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
            .typ_rslt(tt, i).SMPLID = -1                                            ' サンプルＮｏ
            .typ_rslt(tt, i).OKNG = "NG"                                            ' 判定結果
            .typ_rslt(tt, i).hinban = sHinban12                                     ' 品番(12桁)
            bJudg = False
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDAOICW) <> 0 Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' 結晶内開始位置
                .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
                .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFAOI).SAMPLEID             ' サンプルＮｏ
'                If .typ_Param.WFSMP(tt).WFRESAOICW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESAOICW = "1") And (Trim(.typ_y013(tt, WFAOI).SAMPLEID) <> "0") Then
                    'AOI判定失敗
                    .typ_rslt(tt, i).INFO3 = "判定Err"                             ' 情報２
                    'AOI判定
                    If WfCrAoiJudg(.typ_si, .typ_y013(tt, WFAOI), bJudg) Then
                        '画面表示内容設定
                        vTemp = CVar(.typ_y013(tt, WFAOI).MESDATA4)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.0")     ' 情報1
                        vTemp = CVar(.typ_y013(tt, WFAOI).MESDATA5)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.0")     ' 情報2
                        vTemp = CVar(.typ_y013(tt, WFAOI).MESDATA6)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.0")     ' 情報3
                        .typ_rslt(tt, i).INFO4 = ""                                ' 情報4
                        JiltusekiUmu(tt, WFAOI) = True '2001/12/19 S.Sano
                    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    '2.1.3 AN温度 実績反映チェック追加
                        '5番目の情報：AN温度を追加
                        vTemp = CVar(.typ_y013(tt, WFAOI).DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESAOICW = "2" Then
                    '画面表示内容設定
                    bJudg = False
                    .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                       ' 結晶内開始位置
                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ異常"                            ' 情報３
                    .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                                        ' 判定結果
            Else
                .typ_rslt(tt, i).OKNG = "NG"                                        ' 判定結果
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                gsTbcmy028ErrCode = "00142"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1

        Else
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDAOICW) <> 0 Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' 結晶内開始位置
                .typ_rslt(tt, i).NAIYO = Search_CrCode("AO", typTmList())           ' 内容
                .typ_rslt(tt, i).INFO1 = "仕様無"                                   ' 情報１
                .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
                .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
                .typ_rslt(tt, i).INFO4 = ""                                         ' 情報４
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '2.1.3 AN温度 実績反映チェック追加
                '5番目の情報：AN温度を追加
                typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                typ_rslt_ex(tt, i).INFO5 = ""                                       ' 情報5
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
            '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
                typ_rslt_ex(tt, i).INFO6 = ""                                       ' 情報6
                typ_rslt_ex(tt, i).INFO7 = ""                                       ' 情報7
                typ_rslt_ex(tt, i).INFO8 = ""                                       ' 情報8
            '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                .typ_rslt(tt, i).SMPLID = .typ_y013(tt, WFAOI).SAMPLEID             ' サンプルＮｏ
                .typ_rslt(tt, i).OKNG = "OK"                                        ' 判定結果
                .typ_rslt(tt, i).hinban = sHinban12                                 ' 品番(12桁)
                bJudg = False
'                If .typ_Param.WFSMP(tt).WFRESAOICW = "1" Then
                If (.typ_Param.WFSMP(tt).WFRESAOICW = "1") And (Trim(.typ_y013(tt, WFAOI).SAMPLEID) <> "0") Then
                    'AOI判定失敗
                    .typ_rslt(tt, i).INFO3 = "判定Err"                             ' 情報２
                    'AOI判定
                    If WfCrAoiJudg(.typ_si, .typ_y013(tt, WFAOI), bJudg) Then
                        '画面表示内容設定
                        .typ_rslt(tt, i).NAIYO = Search_CrCode("AO", typTmList())  ' 内容
                        vTemp = CVar(.typ_y013(tt, WFAOI).MESDATA4)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0.0")     ' 情報1
                        vTemp = CVar(.typ_y013(tt, WFAOI).MESDATA5)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0.0")     ' 情報2
                        vTemp = CVar(.typ_y013(tt, WFAOI).MESDATA6)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0.0")     ' 情報3
                        .typ_rslt(tt, i).INFO4 = ""                                ' 情報4
                        JiltusekiUmu(tt, WFAOI) = True '2001/12/19 S.Sano
                    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    '2.1.3 AN温度 実績反映チェック追加
                        '5番目の情報：AN温度を追加
                        vTemp = CVar(.typ_y013(tt, WFAOI).DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESAOICW = "2" Then
                    '画面表示内容設定
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ異常"                            ' 情報３
                    .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                End If
                'Add Start 2011/11/28 Y.Hitomi 中間抜試の場合は、参考表示する
                If sSxlPos = "MID" And JudgSW.AOI And bJudg = False Then
                    .typ_rslt(tt, i).OKNG = "参考"
                    MidlJudg = False
                End If
                'Add End   2011/11/28 Y.Hitomi
                i = i + 1
            End If
        End If
        ''残存酸素実績判定/表示処理追加　03/12/09 ooba END ======================================>
        
        ''GD実績判定/表示処理追加　05/02/04 ooba START =========================================>
'Cng Start 2011/11/28 Y.Hitomi
'Chg Start 2011/03/10 SMPK Miyata
'        If JudgSW.GD And CheckKHN(.typ_si.HWFGDKHN, 18, sSxlPos) Then
        '通常抜試：保証方法=保証 かつ 検査頻度＿抜チェック有りの場合、仕様有とする
        '中間抜試：保証方法=保証 かつ 中間抜試フラグが保証の場合、仕様有とする
        If (JudgSW.GD And CheckKHN(.typ_si.HWFGDKHN, 18, sSxlPos) And (tt = SxlTop Or tt = SxlTail)) Or _
           (JudgSW.GD And .typ_si.MSMPFLGWFGD = "1" And (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3") And _
           (tt >= SxlMidl)) Then
'Chg End   2011/03/10 SMPK Miyata
'Cng End   2011/11/28 Y.Hitomi
            '画面表示内容設定
            .typ_rslt(tt, i).pos = -1                                               ' 結晶内開始位置
            .typ_rslt(tt, i).NAIYO = Search_CrCode("GD", typTmList())               ' 内容
            .typ_rslt(tt, i).INFO1 = "仕様有"                                       ' 情報１
            .typ_rslt(tt, i).INFO2 = "検査無"                                       ' 情報２
            .typ_rslt(tt, i).INFO3 = ""                                             ' 情報３
            .typ_rslt(tt, i).INFO4 = ""                                             ' 情報４
        '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        '2.1.3 AN温度 実績反映チェック追加
            '5番目の情報：AN温度を追加
            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
            typ_rslt_ex(tt, i).INFO5 = ""                                           ' 情報5
        '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
        '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
            typ_rslt_ex(tt, i).INFO6 = ""                                           ' 情報6
            typ_rslt_ex(tt, i).INFO7 = ""                                           ' 情報7
            typ_rslt_ex(tt, i).INFO8 = ""                                           ' 情報8
        '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
            .typ_rslt(tt, i).SMPLID = -1                                            ' サンプルＮｏ
            .typ_rslt(tt, i).OKNG = "NG"                                            ' 判定結果
            .typ_rslt(tt, i).hinban = sHinban12                                     ' 品番(12桁)
            bJudg = False
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDGDCW) <> 0 Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' 結晶内開始位置
                .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
                .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
                If .typ_Param.WFSMP(tt).WFHSGDCW = "1" Then
                    '結晶実績
                    .typ_rslt(tt, i).SMPLID = Format(Trim(typ_J015_WFGDJudg(tt).SMPLNO), "0000") & "       【結晶】"   ' サンプルＮｏ
                Else
                    'WF実績
                    .typ_rslt(tt, i).SMPLID = typ_J015_WFGDJudg(tt).SMPLNO          ' サンプルＮｏ
                End If
                If (.typ_Param.WFSMP(tt).WFRESGDCW = "1") And (Trim(typ_J015_WFGDJudg(tt).SMPLNO) <> "") Then
                    'GD判定失敗
                    .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報２
                    
                    'ＷＦ実績/結晶実績識別
                    .typ_si.WFHSGDCW = .typ_Param.WFSMP(tt).WFHSGDCW    '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
                    
                    'GD判定
                    If WfCrGdJudg(.typ_si, typ_J015_WFGDJudg(tt), bJudg) Then
                        '画面表示内容設定
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSRSDEN)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' 情報１
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSRSLDL)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0")        ' 情報２
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSRSDVD2)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0")        ' 情報３
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 UPD By Systech Start
''                        .typ_rslt(tt, i).INFO4 = ""                                 ' 情報４
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSZEROMN)
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0")        ' 情報４
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSZEROMX)
                        .typ_rslt(tt, i).INFO4 = .typ_rslt(tt, i).INFO4 & " , " & DBData2DispData(vTemp, "0")        ' 情報４
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 UPD By Systech End
                        JiltusekiUmu(tt, WFGD) = True
                    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    '2.1.3 AN温度 実績反映チェック追加
                        '5番目の情報：AN温度を追加
                        vTemp = CVar(typ_J015_WFGDJudg(tt).DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESGDCW = "2" Then
                    '画面表示内容設定
                    bJudg = False
                    .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                       ' 結晶内開始位置
                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ異常"                            ' 情報３
                    .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                End If
            End If
            If bJudg = True Then
                .typ_rslt(tt, i).OKNG = "OK"                                        ' 判定結果
            Else
                .typ_rslt(tt, i).OKNG = "NG"                                        ' 判定結果
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                If pbGDJudgeTbl(3) = False Then
                    gsTbcmy028ErrCode = "00148"
                ElseIf pbGDJudgeTbl(3) = False Then
                    gsTbcmy028ErrCode = "00147"
                Else
                    gsTbcmy028ErrCode = "00146"
                End If
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            i = i + 1
            
        Else
            If InStr(IND, .typ_Param.WFSMP(tt).WFINDGDCW) <> 0 Then
                '画面表示内容設定
                .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' 結晶内開始位置
                .typ_rslt(tt, i).NAIYO = Search_CrCode("GD", typTmList())           ' 内容
                .typ_rslt(tt, i).INFO1 = "仕様無"                                   ' 情報１
                .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
                .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
                .typ_rslt(tt, i).INFO4 = ""                                         ' 情報４
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '2.1.3 AN温度 実績反映チェック追加
                '5番目の情報：AN温度を追加
                typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                typ_rslt_ex(tt, i).INFO5 = ""                                       ' 情報5
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
            '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
                typ_rslt_ex(tt, i).INFO6 = ""                                       ' 情報6
                typ_rslt_ex(tt, i).INFO7 = ""                                       ' 情報7
                typ_rslt_ex(tt, i).INFO8 = ""                                       ' 情報8
            '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                If .typ_Param.WFSMP(tt).WFHSGDCW = "1" Then
                    '結晶実績
                    .typ_rslt(tt, i).SMPLID = Format(Trim(typ_J015_WFGDJudg(tt).SMPLNO), "0000") & "       【結晶】"   ' サンプルＮｏ
                Else
                    'WF実績
                    .typ_rslt(tt, i).SMPLID = typ_J015_WFGDJudg(tt).SMPLNO          ' サンプルＮｏ
                End If
                .typ_rslt(tt, i).OKNG = "OK"                                        ' 判定結果
                .typ_rslt(tt, i).hinban = sHinban12                                 ' 品番(12桁)
                bJudg = False
                If (.typ_Param.WFSMP(tt).WFRESGDCW = "1") And (Trim(typ_J015_WFGDJudg(tt).SMPLNO) <> "") Then
                    'GD判定失敗
                    .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報２
                    
                    'ＷＦ実績/結晶実績識別
                    .typ_si.WFHSGDCW = .typ_Param.WFSMP(tt).WFHSGDCW    '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
                    
                    'GD判定
                    If WfCrGdJudg(.typ_si, typ_J015_WFGDJudg(tt), bJudg) Then
                        '画面表示内容設定
                        .typ_rslt(tt, i).NAIYO = Search_CrCode("GD", typTmList())   ' 内容
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSRSDEN)
                        .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' 情報１
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSRSLDL)
                        .typ_rslt(tt, i).INFO2 = DBData2DispData(vTemp, "0")        ' 情報２
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSRSDVD2)
                        .typ_rslt(tt, i).INFO3 = DBData2DispData(vTemp, "0")        ' 情報３
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 UPD By Systech Start
''                        .typ_rslt(tt, i).INFO4 = ""                                 ' 情報４
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSZEROMN)
                        .typ_rslt(tt, i).INFO4 = DBData2DispData(vTemp, "0")        ' 情報４
                        vTemp = CVar(typ_J015_WFGDJudg(tt).MSZEROMX)
                        .typ_rslt(tt, i).INFO4 = .typ_rslt(tt, i).INFO4 & " , " & DBData2DispData(vTemp, "0")        ' 情報４
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 UPD By Systech End
                        JiltusekiUmu(tt, WFGD) = True
                    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    '2.1.3 AN温度 実績反映チェック追加
                        '5番目の情報：AN温度を追加
                        vTemp = CVar(typ_J015_WFGDJudg(tt).DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                        typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                        typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                        typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                        typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                        typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                        typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    End If
                ElseIf .typ_Param.WFSMP(tt).WFRESGDCW = "2" Then
                    '画面表示内容設定
                    bJudg = False
                    .typ_rslt(tt, i).INFO3 = "ｻﾝﾌﾟﾙ異常"                            ' 情報３
                    .typ_rslt(tt, i).INFO4 = ""                                     ' 情報４
                End If
                'Add Start 2011/11/28 Y.Hitomi 中間抜試の場合は、参考表示する
                If sSxlPos = "MID" And JudgSW.GD And bJudg = False Then
                    .typ_rslt(tt, i).OKNG = "参考"
                    MidlJudg = False
                End If
                'Add End   2011/11/28 Y.Hitomi
                i = i + 1
            End If
        End If
        ''GD実績判定/表示処理追加　05/02/04 ooba END ===========================================>

'Chg Start 2011/03/10 SMPK Miyata
'        JudgSW.sp = (.typ_si.HWFNRHS = "H") And CheckKHN(.typ_si.HWFNRKN, 19, sSxlPos)
        JudgSW.sp = (.typ_si.HWFNRHS = "H") And CheckKHN(.typ_si.HWFNRKN, 19, sSxlPos) _
                    And (tt = SxlTop Or tt = SxlTail)

'Chg End   2011/03/10 SMPK Miyata

        ''↓Add 2010/01/07 SIRD対応 Y.Hitomi
'Chg Start 2011/03/10 SMPK Miyata
'        If InStr(IND, .typ_Param.WFSMP(tt).WFINDL4CW) <> 0 Then
        If InStr(IND, .typ_Param.WFSMP(tt).WFINDL4CW) <> 0 And (tt = SxlTop Or tt = SxlTail) Then
'Chg End   2011/03/10 SMPK Miyata
            '画面表示内容設定
            .typ_rslt(tt, i).pos = CStr(INGOTPOS(tt))                           ' 結晶内開始位置
            .typ_rslt(tt, i).NAIYO = Search_CrCode("SD", typTmList())           ' 内容
            .typ_rslt(tt, i).INFO1 = "仕様無"                                   ' 情報１
            .typ_rslt(tt, i).INFO2 = "検査有"                                   ' 情報２
            .typ_rslt(tt, i).INFO3 = "実績無"                                   ' 情報３
            .typ_rslt(tt, i).INFO4 = ""                                         ' 情報４
            
            typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
            typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
            typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
            typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
            typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
            typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
            typ_rslt_ex(tt, i).INFO5 = ""                                       ' 情報5
            typ_rslt_ex(tt, i).INFO6 = ""                                       ' 情報6
            typ_rslt_ex(tt, i).INFO7 = ""                                       ' 情報7
            typ_rslt_ex(tt, i).INFO8 = ""                                       ' 情報8
            .typ_rslt(tt, i).SMPLID = typ_J022_WFSDJudg(tt).SMPLNO             ' サンプルＮｏ
            .typ_rslt(tt, i).OKNG = "OK"                                        ' 判定結果
            .typ_rslt(tt, i).hinban = sHinban12                                 ' 品番(12桁)
            bJudg = False
            'SIRD判定取得
            If (.typ_Param.WFSMP(tt).WFRESL4CW = "1") And (Trim(typ_J022_WFSDJudg(tt).SMPLNO) <> "") Then
                'SIRD判定失敗
                .typ_rslt(tt, i).INFO3 = "判定Err"                              ' 情報２
                If WfCrSdjudg(.typ_si, typ_J022_WFSDJudg(tt), bJudg) Then
                    
                    '画面表示内容設定
                    vTemp = CVar(typ_J022_WFSDJudg(tt).SIRDCNT)
                    .typ_rslt(tt, i).INFO1 = DBData2DispData(vTemp, "0")        ' 情報１
                    .typ_rslt(tt, i).INFO2 = ""                                 ' 情報２
                    .typ_rslt(tt, i).INFO3 = ""                                 ' 情報３
                    .typ_rslt(tt, i).INFO4 = ""                                 ' 情報４
                    
                    JiltusekiUmu(tt, WFSIRD) = True
                    
                    vTemp = CVar(typ_J022_WFSDJudg(tt).DKAN)
                    vTemp = Mid(vTemp, 3, 4)
                    typ_rslt_ex(tt, i).pos = .typ_rslt(tt, i).pos
                    typ_rslt_ex(tt, i).NAIYO = .typ_rslt(tt, i).NAIYO
                    typ_rslt_ex(tt, i).INFO1 = .typ_rslt(tt, i).INFO1
                    typ_rslt_ex(tt, i).INFO2 = .typ_rslt(tt, i).INFO2
                    typ_rslt_ex(tt, i).INFO3 = .typ_rslt(tt, i).INFO3
                    typ_rslt_ex(tt, i).INFO4 = .typ_rslt(tt, i).INFO4
                    typ_rslt_ex(tt, i).INFO5 = DBData2DispData(vTemp, "0")      ' 情報5
                    
                End If
            ElseIf .typ_Param.WFSMP(tt).WFRESL4CW = "2" Then
                '画面表示内容設定
                bJudg = False
                .typ_rslt(tt, i).INFO3 = ""                                     ' 情報３
                .typ_rslt(tt, i).INFO4 = "ｻﾝﾌﾟﾙ異常"                            ' 情報４
            End If
            
''Del Start 2012/03/14 Y.Hitomi
 ''Add Start 2011/03/10 SMPK Miyata
 '            '保証方法が保証か？
 '            If JudgSW.L4 = True Then
 ''Add End   2011/03/10 SMPK Miyata
''Del End 2012/03/14 Y.Hitomi
                If bJudg = True Then
                    .typ_rslt(tt, i).OKNG = "OK"                                ' 判定結果
                Else
                    .typ_rslt(tt, i).OKNG = "NG"                                ' 判定結果
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                    If tt = SxlTop Or tt = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
                End If

'Del Start 2012/03/13 Y.Hitomi
'            End If          'Add 2011/03/10 SMPK Miyata
'Del End   2012/03/13 Y.Hitomi
            
            i = i + 1
        End If
    ''↑Add 2010/01/07 SIRD対応 Y.Hitomi
        
    End With
    WfAllJudg = FUNCTION_RETURN_SUCCESS
End Function
'''''============================================================================================================================
'''''
''''''概要      :製品シート表示
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
''''''          :typ_a         ,I  ,typ_AllTypesC ,各情報構造体
''''''説明      :シートに値を設定する
''''''履歴      :
'''''Public Sub PutSeihinTop()
'''''    Dim i As Integer, j As Integer      ' ﾙｰﾌﾟ ｶｳﾝﾀ
'''''
'''''    With f_cmbc039_2
'''''        For i = 1 To 4
'''''            .spdHinbanTop.col = i
'''''            .spdHinbanTop.row = 1
'''''            Select Case i
'''''            Case 1
'''''                '品番
'''''                .spdHinbanTop.Value = typ_CType.typ_Param.hinban
'''''            Case 2
'''''                'タイプ
'''''                .spdHinbanTop.Value = typ_CType.typ_si.HWFTYPE
'''''            Case 3
'''''                '方位
'''''                .spdHinbanTop.Value = typ_CType.typ_si.HWFCDIR
'''''            Case 4
'''''                '結晶ドープ
'''''                .spdHinbanTop.Value = typ_CType.typ_si.HWFCDOP
'''''            End Select
'''''        Next i
'''''    End With
'''''
'''''End Sub
'''''============================================================================================================================
'''''
''''''概要      :製品シート表示
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
''''''          :typ_a         ,I  ,typ_AllTypesC ,各情報構造体
''''''説明      :シートに値を設定する
''''''履歴      :
'''''Public Sub PutSeihinCenter()
'''''    Dim i As Integer, j As Integer      ' ﾙｰﾌﾟ ｶｳﾝﾀ
'''''
'''''    'CENTER側
'''''    With f_cmbc039_2
'''''        For i = 1 To 9
'''''            .spdHinbanCen.col = i
'''''            .spdHinbanCen.row = 1
'''''            Select Case i
'''''            Case 1
'''''                '比抵抗
''''''2001/12/26 S.Sano                .spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFRMIN, "0.0000") & " - " & DBData2DispData(typ_CType.typ_si.HWFRMAX, "0.0000")
'''''                .spdHinbanCen.Value = toRsStr(typ_CType.typ_si.HWFRMIN) & " - " & toRsStr(typ_CType.typ_si.HWFRMAX) '2001/12/26 S.Sano
'''''            Case 2
'''''                'Oi
'''''                .spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFONMIN, "0.00") & " - " & DBData2DispData(typ_CType.typ_si.HWFONMAX, "0.00")
'''''            Case 3
'''''                'BMD1
'''''                '.spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFBM1AN, "0") & " - " & DBData2DispData(typ_CType.typ_si.HWFBM1AX, "0")
'''''                'べき乗数変更　2003/05/19 osawa
'''''                .spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFBM1AN, "0.0") & " - " & DBData2DispData(typ_CType.typ_si.HWFBM1AX, "0.0")
'''''            Case 4
'''''                'BMD2
'''''                '.spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFBM2AN, "0") & " - " & DBData2DispData(typ_CType.typ_si.HWFBM2AX, "0")
'''''                .spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFBM2AN, "0.0") & " - " & DBData2DispData(typ_CType.typ_si.HWFBM2AX, "0.0")
'''''            Case 5
'''''                'BMD3
'''''                '.spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFBM3AN, "0") & " - " & DBData2DispData(typ_CType.typ_si.HWFBM3AX, "0")
'''''                .spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFBM3AN, "0.0") & " - " & DBData2DispData(typ_CType.typ_si.HWFBM3AX, "0.0")
'''''            Case 6
'''''                'OSF1
'''''                .spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFOF1AX, "0.00") & " , " & DBData2DispData(typ_CType.typ_si.HWFOF1MX, "0.0")
'''''            Case 7
'''''                'OSF2
'''''                .spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFOF2AX, "0.00") & " , " & DBData2DispData(typ_CType.typ_si.HWFOF2MX, "0.0")
'''''            Case 8
'''''                'OSF3
'''''                .spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFOF3AX, "0.00") & " , " & DBData2DispData(typ_CType.typ_si.HWFOF3MX, "0.0")
'''''            Case 9
'''''                'OSF4
'''''                .spdHinbanCen.Value = DBData2DispData(typ_CType.typ_si.HWFOF4AX, "0.00") & " , " & DBData2DispData(typ_CType.typ_si.HWFOF4MX, "0.0")
'''''            End Select
'''''        Next i
'''''    End With
'''''
'''''End Sub
'''''============================================================================================================================
'''''
''''''概要      :製品シート表示
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
''''''          :typ_a         ,I  ,typ_AllTypesC ,各情報構造体
''''''説明      :シートに値を設定する
''''''履歴      :
'''''Public Sub PutSeihinTail()
'''''    Dim i As Integer, j As Integer      ' ﾙｰﾌﾟ ｶｳﾝﾀ
'''''
'''''    'TAIL側
'''''    With f_cmbc039_2
'''''        For i = 1 To 6
'''''            .spdHinbanTail.col = i
'''''            .spdHinbanTail.row = 1
'''''            Select Case i
'''''            Case 1
'''''                'DS
'''''                .spdHinbanTail.Value = DBData2DispData(typ_CType.typ_si.HWFDSOMN, "0") & " - " & _
'''''                                      DBData2DispData(typ_CType.typ_si.HWFDSOMX, "0")
'''''            Case 2
'''''                'DZ
'''''                .spdHinbanTail.Value = DBData2DispData(typ_CType.typ_si.HWFMKMIN, "0.0") & " - " & _
'''''                                      DBData2DispData(typ_CType.typ_si.HWFMKMAX, "0.0")
'''''            Case 3
'''''                'SP
'''''                .spdHinbanTail.Value = DBData2DispData(typ_CType.typ_si.HWFSPVMX, "0.00") & " , " & _
'''''                                      DBData2DispData(typ_CType.typ_si.HWFDLMIN, "0") & " - " & _
'''''                                      DBData2DispData(typ_CType.typ_si.HWFDLMAX, "0")
'''''            Case 4
'''''                'D1
'''''                .spdHinbanTail.Value = DBData2DispData(typ_CType.typ_si.HWFOS1MN, "0.0") & " - " & _
'''''                                      DBData2DispData(typ_CType.typ_si.HWFOS1MX, "0.0")
'''''            Case 5
'''''                'D2
'''''                .spdHinbanTail.Value = DBData2DispData(typ_CType.typ_si.HWFOS2MN, "0.0") & " - " & _
'''''                                      DBData2DispData(typ_CType.typ_si.HWFOS2MX, "0.0")
'''''            Case 6
'''''                'D3
'''''                .spdHinbanTail.Value = DBData2DispData(typ_CType.typ_si.HWFOS3MN, "0.0") & " - " & _
'''''                                      DBData2DispData(typ_CType.typ_si.HWFOS3MX, "0.0")
'''''            End Select
'''''        Next i
'''''    End With
'''''
'''''End Sub
'''''============================================================================================================================
'''''
''''''概要      :比抵抗値表示
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
''''''          :typ_a         ,I  ,typ_AllTypesC ,各情報構造体
''''''説明      :比抵抗領域に値を表示する
''''''履歴      :
'''''Public Sub PutRs()
'''''
'''''    '比抵抗値表示(TOP側)
'''''    PutRsTop
'''''
'''''    '比抵抗値表示(TAIL側)
'''''    PutRsTail
'''''
'''''End Sub
'''''============================================================================================================================
'''''
''''''概要      :比抵抗値表示(TOP)
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
''''''          :typ_a         ,I  ,typ_AllTypesC ,各情報構造体
''''''          :bJudg         ,I  ,Boolean      ,判定結果
''''''          :dblScut       ,I  ,Double       ,再カット位置
''''''          :dblCoef       ,I  ,Double       ,実行偏析
''''''説明      :比抵抗領域に値を表示する
''''''履歴      :
'''''Public Sub PutRsTop()
'''''    Dim bJudg As Boolean
'''''    Dim dblScut As Double
'''''    Dim dblCoef As Double
'''''
'''''''2001/12/18 S.Sano    bJudg = typ_CType.bOKNG(SxlTop)
''''''2002/03/04 S.Sano    bJudg = (typ_CType.bOKNG(SxlTop) Or bPPlus Or bNPlus) ''2001/12/18 S.Sano
'''''    dblScut = typ_CType.dblScut(SxlTop)
'''''    dblCoef = typ_CType.COEF(SxlTop)
'''''
'''''    With f_cmbc039_2
'''''        '' WF検査指示（Rs)*****************************************************************
'''''        If JudgSW.rs Then
'''''            If InStr("1345", typ_CType.typ_Param.WFSMP(SxlTop).WFINDRS) <> 0 Then
'''''
'''''                If typ_CType.typ_Param.WFSMP(SxlTop).WFRESRS = "1" Then
'''''                    .txtSXLTop.Text = DBData2DispData(typ_CType.typ_Param.IngotPos, "0")            '位置
'''''                    .txtRRGTop.Text = DBData2DispData(typ_CType.typ_y013(SxlTop, WFRES).MESDATA6, "0.00")  'RRG
'''''
''''''2002/03/04 S.Sano                    If typ_CType.JudgRrg(1) = False Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''                    If Not (typ_CType.JudgRrg(SxlTop) Or bPPlus Or bNPlus) Then '2002/03/04 S.Sano
'''''                    If Not (typ_CType.JudgRrg(SxlTop)) Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''                        CtrlEnabled .txtRRGTop, CTRL_DISABLE_WARNING, False  'RRG
'''''                    End If
'''''                    If dblCoef = -1 Or dblCoef = -9999 Or Mid(typ_CType.typ_Param.CRYNUM, 1, 1) = "8" Then
'''''                        .txtJHAll.Text = ""         '実行偏析ブロック
'''''                    Else
'''''                        .txtJHAll.Text = DBData2DispData(dblCoef, "0.000")         '実行偏析ブロック
'''''                    End If
'''''
'''''                    If Mid(typ_CType.typ_Param.CRYNUM, 1, 1) <> "8" Then
'''''                        '再カット位置
''''''2002/03/04 S.Sano                        If typ_CType.JudgRes(1) Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''                        If typ_CType.JudgRes(SxlTop) Or bPPlus Or bNPlus Then '2002/03/04 S.Sano
'''''                        If typ_CType.JudgRes(SxlTop) Then '2002/03/04 S.Sano
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''                            .txtCutPosTop.Text = "OK"
'''''                        Else
''''''2001/08/30 S.Sano                            If dblScut < 0 Then
''''''2001/08/30 S.Sano                                .txtCutPosTop.Text = "0"
''''''2001/08/30 S.Sano                            ElseIf dblScut <= typ_CType.typ_Param.INGOTPOS Then
''''''2001/08/30 S.Sano                                .txtCutPosTop.Text = "0"
''''''2001/08/30 S.Sano                            ElseIf dblScut >= typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.INGOTPOS Then
''''''2001/08/30 S.Sano                                .txtCutPosTop.Text = typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.INGOTPOS
''''''2001/08/30 S.Sano                            Else
''''''2001/08/30 S.Sano                                .txtCutPosTop.Text = DBData2DispData(dblScut, "0")
''''''2001/08/30 S.Sano                            End If
''''''2001/08/30 S.Sano Start
'''''                            Select Case dblScut
'''''                            Case -9999
'''''                                .txtCutPosTop.Text = ""
'''''                            Case Is <= typ_CType.typ_Param.IngotPos
'''''                                .txtCutPosTop.Text = typ_CType.typ_Param.IngotPos
'''''                            Case Is >= typ_CType.typ_Param.IngotPos + typ_CType.typ_Param.LENGTH
'''''                                .txtCutPosTop.Text = typ_CType.typ_Param.IngotPos + typ_CType.typ_Param.LENGTH
'''''                            Case Else
'''''                                .txtCutPosTail.Text = DBData2DispData(dblScut, "0")
'''''                            End Select
''''''2001/08/30 S.Sano End
'''''                            CtrlEnabled .txtCutPosTop, CTRL_DISABLE_WARNING, False  'TOP再カット
'''''                            intEnCmd = 1
'''''                        End If
'''''                    Else
''''''2001/08/30 S.Sano                        .txtCutPosTop.Text = "OK"
''''''2001/08/30 S.Sano Start
''''''2002/03/04 S.Sano                        If typ_CType.JudgRes(1) Then
'''''
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''                        If typ_CType.JudgRes(SxlTop) Or bPPlus Or bNPlus Then '2002/03/04 S.Sano
'''''                        If typ_CType.JudgRes(SxlTop) Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''                            .txtCutPosTop.Text = "OK"
'''''                        Else
'''''                            Select Case dblScut
'''''                            Case -9999
'''''                                .txtCutPosTop.Text = ""
'''''                            Case Is <= typ_CType.typ_Param.IngotPos
'''''                                .txtCutPosTop.Text = typ_CType.typ_Param.IngotPos
'''''                            Case Is >= typ_CType.typ_Param.IngotPos + typ_CType.typ_Param.LENGTH
'''''                                .txtCutPosTop.Text = typ_CType.typ_Param.IngotPos + typ_CType.typ_Param.LENGTH
'''''                            End Select
'''''                            CtrlEnabled .txtCutPosTop, CTRL_DISABLE_WARNING, False  'TOP再カット
'''''                            intEnCmd = 1
'''''                        End If
''''''2001/08/30 S.Sano End
'''''                    End If
'''''
'''''                    '比抵抗
'''''                    If UBound(typ_CType.typ_y013top) > 0 Then
'''''                        With .spdMeasTop
'''''                            .SetFloat 1, 1, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA1)
'''''                            .SetFloat 1, 2, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA2)
'''''                            .SetFloat 1, 3, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA3)
'''''                            .SetFloat 1, 4, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA4)
'''''                            .SetFloat 1, 5, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA5)
'''''                        End With
'''''                        RsSpreadSet .spdMeasTop, 1 '2002/01/25 S.Sano
'''''                    Else
'''''                        With .spdMeasTop
'''''                            .col = 1
'''''                            .row = 1:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "仕様有"
'''''                            .row = 2:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "検査有"
'''''                            .row = 3:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "実績無"
'''''                        End With
'''''                    End If
'''''                End If
'''''            Else
'''''                .txtSXLTop.Text = ""            '位置
'''''                .txtRRGTop.Text = ""            'RRG
'''''                .txtJHAll.Text = ""
'''''                '再カット位置
'''''
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''                If bPPlus Or bNPlus Then
''''''                    .txtCutPosTop.Text = "OK"
''''''                Else
'''''                    .txtCutPosTop.Text = "NG"
'''''                    CtrlEnabled .txtCutPosTop, CTRL_DISABLE_WARNING, False  'TOP再カット
''''''                End If
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''                '比抵抗
'''''                With .spdMeasTop
'''''                    .col = 1
'''''                    .row = 1:
'''''                            .CellType = CellTypeStaticText
'''''                            .Value = "仕様有"
'''''                    .row = 2:
'''''                            .CellType = CellTypeStaticText
'''''                            .Value = "検査無"
'''''                    .row = 3: .Value = ""
'''''                    .row = 4: .Value = ""
'''''                    .row = 5: .Value = ""
'''''                End With
'''''            End If
'''''        Else
'''''            If InStr("1345", typ_CType.typ_Param.WFSMP(SxlTop).WFINDRS) <> 0 Then
'''''
'''''                If typ_CType.typ_Param.WFSMP(SxlTop).WFRESRS = "1" Then
'''''                    .txtSXLTop.Text = DBData2DispData(typ_CType.typ_Param.IngotPos, "0")            '位置
'''''                    .txtRRGTop.Text = DBData2DispData(typ_CType.typ_y013(SxlTop, WFRES).MESDATA6, "0.00")  'RRG
'''''                    If dblCoef = -1 Or dblCoef = -9999 Then
'''''                        .txtJHAll.Text = ""         '実行偏析ブロック
'''''                    Else
'''''                        .txtJHAll.Text = DBData2DispData(dblCoef, "0.000")         '実行偏析ブロック
'''''                    End If
'''''
'''''                    '再カット位置
'''''                    .txtCutPosTop.Text = "OK"
'''''
'''''                    '比抵抗
'''''                    If UBound(typ_CType.typ_y013top) > 0 Then
'''''                        With .spdMeasTop
'''''                            .SetFloat 1, 1, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA1)
'''''                            .SetFloat 1, 2, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA2)
'''''                            .SetFloat 1, 3, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA3)
'''''                            .SetFloat 1, 4, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA4)
'''''                            .SetFloat 1, 5, CDbl(typ_CType.typ_y013(SxlTop, WFRES).MESDATA5)
'''''                        End With
'''''                        RsSpreadSet .spdMeasTop, 1 '2002/01/25 S.Sano
'''''                    Else
'''''                        With .spdMeasTop
'''''                            .col = 1
'''''                            .row = 1:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "仕様無"
'''''                            .row = 2:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "検査有"
'''''                            .row = 3:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "実績無"
'''''                        End With
'''''                    End If
'''''                Else
'''''                    .txtSXLTop.Text = DBData2DispData(typ_CType.typ_Param.IngotPos, "0")            '位置
'''''                    .txtRRGTop.Text = DBData2DispData(typ_CType.typ_y013(SxlTop, WFRES).MESDATA6, "0.00")  'RRG
'''''                    .txtJHAll.Text = ""
'''''                    '再カット位置
'''''                    .txtCutPosTop.Text = "OK"
'''''                    '比抵抗
'''''                    With .spdMeasTop
'''''                        .col = 1
'''''                        .row = 1:
'''''                                .CellType = CellTypeStaticText
'''''                                .Value = "仕様無"
'''''                        .row = 2:
'''''                                .CellType = CellTypeStaticText
'''''                                .Value = "検査有"
'''''                        .row = 3:
'''''                                .CellType = CellTypeStaticText
'''''                                .Value = "実績無"
'''''                        .row = 4: .Value = ""
'''''                        .row = 5: .Value = ""
'''''                    End With
'''''                End If
'''''            End If
'''''        End If
'''''    End With
'''''
'''''End Sub
'''''============================================================================================================================
'''''
''''''概要      :比抵抗値表示(TAIL)
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
''''''          :typ_a         ,I  ,typ_AllTypesC ,各情報構造体
''''''          :bJudg         ,I  ,Boolean      ,判定結果
''''''          :dblScut       ,I  ,Double       ,再カット位置
''''''説明      :比抵抗領域に値を表示する
''''''履歴      :
'''''Public Sub PutRsTail()
'''''    Dim bJudg As Boolean
'''''    Dim dblScut As Double
'''''
'''''''2001/12/18 S.Sano    bJudg = typ_CType.bOKNG(SxlTail)
''''''2002/03/04 S.Sano    bJudg = (typ_CType.bOKNG(SxlTail) Or bPPlus Or bNPlus) ''2001/12/18 S.Sano
'''''    dblScut = typ_CType.dblScut(SxlTail)
'''''
'''''
'''''    With f_cmbc039_2
'''''        '' WF検査指示（Rs)*****************************************************************
'''''        If JudgSW.rs Then
'''''            If InStr("2345", typ_CType.typ_Param.WFSMP(SxlTail).WFINDRS) <> 0 Then
'''''
'''''                If typ_CType.typ_Param.WFSMP(SxlTail).WFRESRS = "1" Then
'''''                    .txtSXLTail.Text = DBData2DispData(typ_CType.typ_Param.IngotPos + typ_CType.typ_Param.LENGTH, "0")           '位置
'''''                    .txtRRGTail.Text = DBData2DispData(typ_CType.typ_y013(SxlTail, WFRES).MESDATA6, "0.00")  'RRG
''''''2002/03/04 S.Sano                    If typ_CType.JudgRrg(1) = False Then
'''''
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''                    If Not (typ_CType.JudgRrg(SxlTail) Or bPPlus Or bNPlus) Then
'''''                    If Not (typ_CType.JudgRrg(SxlTail)) Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''                        CtrlEnabled .txtRRGTail, CTRL_DISABLE_WARNING, False  'RRG
'''''                    End If
'''''
'''''                    '再カット位置
'''''                    If Mid(typ_CType.typ_Param.CRYNUM, 1, 1) <> "8" Then
''''''2002/03/04 S.Sano                        If typ_CType.JudgRes(2) = True Then
'''''
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''                        If typ_CType.JudgRes(SxlTail) Or bPPlus Or bNPlus Then '2002/03/04 S.Sano
'''''                        If typ_CType.JudgRes(SxlTail) Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''                            .txtCutPosTail.Text = "OK"
'''''                        Else
''''''2001/08/30 S.Sano                            If dblScut < 0 Then
''''''2001/08/30 S.Sano                                .txtCutPostail.Text = "0"
''''''2001/08/30 S.Sano                            ElseIf dblScut <= typ_CType.typ_Param.INGOTPOS Then
''''''2001/08/30 S.Sano                                .txtCutPostail.Text = "0"
''''''2001/08/30 S.Sano                            ElseIf dblScut >= typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.INGOTPOS Then
''''''2001/08/30 S.Sano                                .txtCutPostail.Text = typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.INGOTPOS
''''''2001/08/30 S.Sano                            Else
''''''2001/08/30 S.Sano                                .txtCutPostail.Text = DBData2DispData(dblScut, "0")
''''''2001/08/30 S.Sano                            End If
''''''2001/08/30 S.Sano Start
'''''                            Select Case dblScut
'''''                            Case -9999
'''''                                .txtCutPosTail.Text = ""
'''''                            Case Is <= typ_CType.typ_Param.IngotPos
'''''                                .txtCutPosTail.Text = typ_CType.typ_Param.IngotPos
'''''                            Case Is >= typ_CType.typ_Param.IngotPos + typ_CType.typ_Param.LENGTH
'''''                                .txtCutPosTail.Text = typ_CType.typ_Param.IngotPos + typ_CType.typ_Param.LENGTH
'''''                            Case Else
'''''                                .txtCutPosTail.Text = DBData2DispData(dblScut, "0")
'''''                            End Select
''''''2001/08/30 S.Sano End
'''''                            CtrlEnabled .txtCutPosTail, CTRL_DISABLE_WARNING, False  'tail再カット
'''''                            intEnCmd = 1
'''''                        End If
'''''                    Else
''''''2001/08/30 S.Sano                        .txtCutPostail.Text = "OK"
''''''2001/08/30 S.Sano Start
''''''2002/03/04 S.Sano                        If typ_CType.JudgRes(2) = True Then
'''''
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''                        If typ_CType.JudgRes(SxlTail) Or bPPlus Or bNPlus Then '2002/03/04 S.Sano
'''''                        If typ_CType.JudgRes(SxlTail) Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''                            .txtCutPosTail.Text = "OK"
'''''                        Else
'''''                            Select Case dblScut
'''''                            Case -9999
'''''                                .txtCutPosTail.Text = ""
'''''                            Case Is <= typ_CType.typ_Param.IngotPos
'''''                                .txtCutPosTail.Text = typ_CType.typ_Param.IngotPos
'''''                            Case Is >= typ_CType.typ_Param.IngotPos + typ_CType.typ_Param.LENGTH
'''''                                .txtCutPosTail.Text = typ_CType.typ_Param.IngotPos + typ_CType.typ_Param.LENGTH
'''''                            End Select
'''''                            CtrlEnabled .txtCutPosTail, CTRL_DISABLE_WARNING, False  'tail再カット
'''''                            intEnCmd = 1
'''''                        End If
''''''2001/08/30 S.Sano End
'''''                    End If
'''''
'''''                    '比抵抗
'''''                    If UBound(typ_CType.typ_y013tail) > 0 Then
'''''                        With .spdMeasTail
'''''                            .SetFloat 1, 1, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA1)
'''''                            .SetFloat 1, 2, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA2)
'''''                            .SetFloat 1, 3, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA3)
'''''                            .SetFloat 1, 4, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA4)
'''''                            .SetFloat 1, 5, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA5)
'''''                        End With
'''''                        RsSpreadSet .spdMeasTail, 1 '2002/01/25 S.Sano
'''''                    Else
'''''                        With .spdMeasTail
'''''                            .col = 1
'''''                            .row = 1:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "仕様有"
'''''                            .row = 2:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "検査有"
'''''                            .row = 3:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "実績無"
'''''                        End With
'''''                    End If
'''''                End If
'''''            Else
'''''                .txtSXLTail.Text = ""            '位置
'''''                .txtRRGTail.Text = ""            'RRG
'''''                .txtJHAll.Text = ""
'''''                '再カット位置
'''''
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''                If bPPlus Or bNPlus Then
''''''                    .txtCutPosTail.Text = "OK"
''''''                Else
'''''                    .txtCutPosTail.Text = "NG"
'''''                    CtrlEnabled .txtCutPosTail, CTRL_DISABLE_WARNING, False  'Tail再カット
''''''                End If
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''                '比抵抗
'''''                With .spdMeasTail
'''''                    .col = 1
'''''                    .row = 1:
'''''                            .CellType = CellTypeStaticText
'''''                            .Value = "仕様有"
'''''                    .row = 2:
'''''                            .CellType = CellTypeStaticText
'''''                            .Value = "検査無"
'''''                    .row = 3: .Value = ""
'''''                    .row = 4: .Value = ""
'''''                    .row = 5: .Value = ""
'''''                End With
'''''            End If
'''''        Else
'''''            If InStr("2345", typ_CType.typ_Param.WFSMP(SxlTail).WFINDRS) <> 0 Then
'''''
'''''                If typ_CType.typ_Param.WFSMP(SxlTail).WFRESRS = "1" Then
'''''                    .txtSXLTail.Text = DBData2DispData(typ_CType.typ_Param.IngotPos, "0")            '位置
'''''                    .txtRRGTail.Text = DBData2DispData(typ_CType.typ_y013(SxlTail, WFRES).MESDATA6, "0.00")  'RRG
'''''
'''''                    '再カット位置
'''''                    .txtCutPosTail.Text = "OK"
'''''
'''''                    '比抵抗
'''''                    If UBound(typ_CType.typ_y013tail) > 0 Then
'''''                        With .spdMeasTail
'''''                            .SetFloat 1, 1, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA1)
'''''                            .SetFloat 1, 2, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA2)
'''''                            .SetFloat 1, 3, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA3)
'''''                            .SetFloat 1, 4, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA4)
'''''                            .SetFloat 1, 5, CDbl(typ_CType.typ_y013(SxlTail, WFRES).MESDATA5)
'''''                        End With
'''''                        RsSpreadSet .spdMeasTail, 1 '2002/01/25 S.Sano
'''''                    Else
'''''                        With .spdMeasTail
'''''                            .col = 1
'''''                            .row = 1:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "仕様無"
'''''                            .row = 2:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "検査有"
'''''                            .row = 3:
'''''                                    .CellType = CellTypeStaticText
'''''                                    .Value = "実績無"
'''''                        End With
'''''                    End If
'''''                Else
'''''                    .txtSXLTail.Text = DBData2DispData(typ_CType.typ_Param.IngotPos, "0")            '位置
'''''                    .txtRRGTail.Text = DBData2DispData(typ_CType.typ_y013(SxlTail, WFRES).MESDATA6, "0.00")  'RRG
'''''                    .txtJHAll.Text = ""
'''''                    '再カット位置
'''''                    .txtCutPosTop.Text = "OK"
'''''                    '比抵抗
'''''                    With .spdMeasTop
'''''                        .col = 1
'''''                        .row = 1:
'''''                                .CellType = CellTypeStaticText
'''''                                .Value = "仕様無"
'''''                        .row = 2:
'''''                                .CellType = CellTypeStaticText
'''''                                .Value = "検査有"
'''''                        .row = 3:
'''''                                .CellType = CellTypeStaticText
'''''                                .Value = "実績無"
'''''                        .row = 4: .Value = ""
'''''                        .row = 5: .Value = ""
'''''                    End With
'''''                End If
'''''            End If
'''''        End If
'''''    End With
'''''
'''''End Sub
'''''============================================================================================================================
'''''
''''''概要      :実績値表示(TOP)
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
''''''          :typ_rslt()    ,I  ,typ_ALLRSLT  ,実績情報構造体
''''''          :tt            ,I  ,Integer      ,TopTail判定用
''''''説明      :実績領域に値を表示する
''''''履歴      :
'''''Public Sub PutRslt(typ_rslt() As typ_ALLRSLT, tt As Integer)
'''''    Dim i, j As Integer
'''''    Dim va As vaSpread
'''''    Dim spMaxLine As Long
'''''
'''''    ''最大行数取得
'''''    spMaxLine = 0
'''''    Do While typ_rslt(tt, spMaxLine).OKNG <> ""
'''''        spMaxLine = spMaxLine + 1
'''''    Loop
'''''
'''''    If tt = SxlTop Then
'''''        Set va = f_cmbc039_2.spdKensaTop
'''''    Else
'''''        Set va = f_cmbc039_2.spdKensaTail
'''''    End If
'''''
'''''    SpCtrlInit va, spMaxLine
'''''    SpCtrlBlockEnabled va, 1, 1, spMaxLine, 5, CTRL_DISABLE
'''''
'''''
'''''    i = 1
'''''    Do While typ_rslt(tt, i - 1).OKNG <> ""
'''''        With typ_rslt(tt, i - 1)
'''''            va.row = i
'''''            For j = 1 To 8
'''''                va.col = j
'''''                Select Case j
'''''                Case 1
'''''                    '位置
'''''                    va.Value = DBData2DispData(CVar(.pos), "0")
'''''                Case 2
'''''                    '内容
'''''                    va.Value = .NAIYO
'''''                Case 3
'''''                    '情報１
'''''                    va.Value = .INFO1
'''''                Case 4
'''''                    '情報２
'''''                    va.Value = .INFO2
'''''                Case 5
'''''                    '情報３
'''''                    va.Value = .INFO3
'''''                Case 6
'''''                    '情報４
'''''                    va.Value = .INFO4
'''''                Case 7
'''''                    '判定
'''''''2001/12/18 S.Sano                    If .OKNG = "NG" Then
'''''''2001/12/18 S.Sano                        SpCtrlEnabled va, va.Col, va.row, CTRL_DISABLE_WARNING
'''''''2001/12/18 S.Sano                        'va.BackColor = &H8080FF
'''''''2001/12/18 S.Sano                        intEnCmd = 1
'''''''2001/12/18 S.Sano                    End If
'''''''2001/12/18 S.Sano                    va.Value = .OKNG
'''''''2001/12/18 S.Sano Start
'''''
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''                    If bPPlus Or bNPlus Then
''''''                        va.Value = "OK"
''''''                    Else
'''''                        If .OKNG = "NG" Then
'''''                            SpCtrlEnabled va, va.col, va.row, CTRL_DISABLE_WARNING
'''''                            'va.BackColor = &H8080FF
'''''                            intEnCmd = 1
'''''                        End If
'''''                        va.Value = .OKNG
''''''                    End If
'''''''2001/12/18 S.Sano End
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''                Case 8
'''''                    '位置
'''''                    va.Value = CStr(DBData2DispData(.SMPLID, "0"))
'''''                End Select
'''''            Next j
'''''        End With
'''''        i = i + 1
'''''    Loop
'''''
'''''    'ソート処理
'''''    If i <> 1 Then
'''''        With va
'''''            .MaxRows = i - 1                      '　品番（行数）
'''''            .row = 1                            ' セルブロックを設定
'''''            .col = 1
'''''            .row2 = i - 1
'''''            .col2 = 8
'''''            .SortBy = SS_SORT_BY_ROW
'''''            .SortKey(1) = 7                    ' 第１ソートキーを設定
'''''            ' 昇順に並べ替え
'''''            .SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
'''''            .Action = SS_ACTION_SORT
'''''        End With
'''''    End If
'''''
'''''End Sub

'概要      :抵抗判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :typ_CType     ,I  ,typ_AllTypesC                        :全部構造体
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :仕様情報構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(0:判定OK, 1:判定NG)
'          :dblScut       ,O  ,Double                               :再カット位置
'          :tt            ,I  ,Integer                              :Top,Tail判定用
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :抵抗判定を行い、判定がNGだった場合は再カット位置を返す
'履歴      :
Public Function WfCrResJudg(typ_CType As typ_AllTypesC, _
                            typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            bJudg As Boolean, _
                            dblScut As Double, _
                            tt As Integer) As Boolean

    Dim ErrInfo     As ERROR_INFOMATION
    Dim rs          As W_RES
    Dim cc          As type_Coefficient
    Dim rp          As type_ResPosCal
    Dim COEF        As Double
    Dim wgtCharge   As Long         '偏析計算用パラメータ
    Dim wgtTop      As Double       '偏析計算用パラメータ
    Dim wgtTopCut   As Double       '偏析計算用パラメータ
    Dim DM          As Double       '偏析計算用パラメータ
    
    '抵抗判定引数設定
    rs.GuaranteeRes.cMeth = typ_si.HWFRSPOH         ' 品ＷＦ比抵抗測定位置＿方
    rs.GuaranteeRes.cCount = typ_si.HWFRSPOT        ' 品ＷＦ比抵抗測定位置＿点
    rs.GuaranteeRes.cPos = typ_si.HWFRSPOI          ' 品ＷＦ比抵抗測定位置＿位
    rs.GuaranteeRes.cObj = typ_si.HWFRHWYT          ' 品ＷＦ比抵抗保証方法＿対
    rs.GuaranteeRes.cJudg = typ_si.HWFRHWYS         ' 品ＷＦ比抵抗保証方法＿処
    rs.GuaranteeCal = typ_si.HWFRMCAL               ' 品ＷＦ比抵抗面内計算 2001/11/08 S.Sano
    rs.SpecResMin = typ_si.HWFRMIN                  ' 品ＷＦ比抵抗下限
    rs.SpecResMax = typ_si.HWFRMAX                  ' 品ＷＦ比抵抗上限
    rs.SpecResAveMin = typ_si.HWFRAMIN              ' 品ＷＦ比抵抗平均下限
    rs.SpecResAveMax = typ_si.HWFRAMAX              ' 品ＷＦ比抵抗平均上限
    rs.SpecRrg = typ_si.HWFRMBNP                    ' 品ＷＦ比抵抗面内分布
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
    'チェック用AN温度を追加
    rs.Antnp = typ_si.HWFANTNP                      ' 品ＷＦＡＮ温度
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'--------------- 2008/08/25 INSERT START  By Systech --------------
    rs.DkTmpSiyo = typ_CType.DkTmpSiyo
    
'Cng Start 2011/09/26 Y.Hitomi
    If tt = SxlTop Or tt = SxlTail Then
        rs.DkTmpJsk = typ_CType.DkTmpJsk(tt)
    Else
        typ_CType.DkTmpJsk(tt) = GetWfDKTmpCode(False, typ_CType.typ_Param.WFSMP(1))
        rs.DkTmpJsk = typ_CType.DkTmpJsk(tt)
    End If
''Chg Start 2011/03/09 SMPK Miyata
'    If tt = SxlTop Or tt = SxlTail Then
'    rs.DkTmpJsk = typ_CType.DkTmpJsk(tt)
'    Else
'        '中間抜試のDK温度判定はDK温度実績なしで判定OKにする
'        rs.DkTmpJsk = ""
'    End If
''Chg End   2011/03/09 SMPK Miyata
'Cng End   2011/09/26 Y.Hitomi
'--------------- 2008/08/25 INSERT  END   By Systech --------------
    With typ_CType
        rs.Res(0) = NtoZ2(.typ_y013(tt, WFRES).MESDATA1)   ' 測定値１
        rs.Res(1) = NtoZ2(.typ_y013(tt, WFRES).MESDATA2)   ' 測定値２
        rs.Res(2) = NtoZ2(.typ_y013(tt, WFRES).MESDATA3)   ' 測定値３
        rs.Res(3) = NtoZ2(.typ_y013(tt, WFRES).MESDATA4)   ' 測定値４
        rs.Res(4) = NtoZ2(.typ_y013(tt, WFRES).MESDATA5)   ' 測定値５
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    '2.1.3 AN温度 実績反映チェック追加
        'チェック用AN温度を追加
        rs.ResAntnp = NtoZ2(Mid(.typ_y013(tt, WFRES).DKAN, 3, 4)) ' 測定値６
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    End With
    
    '抵抗判定
    If WfRESJudg(rs, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrResJudg = False
        typ_CType.JudgRes(tt) = rs.JudgRes '2002/01/25 S.Sano
        typ_CType.JudgRrg(tt) = rs.JudgRrg '2002/01/25 S.Sano
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    '2.1.3 AN温度 実績反映チェック追加
        'チェック用AN温度を追加
'Chg Start 2011/08/12 Y.Hitomi
        typ_CType.JudgAntnp(tt) = rs.JudgAntnp
'        If tt = SxlTop Or tt = SxlTail Then typ_CType.JudgAntnp(tt) = rs.JudgAntnp
'Chg End   2011/08/12 Y.Hitomi

    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'--------------- 2008/08/25 INSERT START  By Systech --------------
'Chg Start 2011/03/25 Y.Hitomi
        typ_CType.JudgDkTmp(tt) = rs.JudgDkTmp
'        If tt = SxlTop Or tt = SxlTail Then typ_CType.JudgDkTmp(tt) = rs.JudgDkTmp
'Chg End   2011/03/25 Y.Hitomi

'--------------- 2008/08/25 INSERT  END   By Systech --------------
        Exit Function
    End If
    typ_CType.JudgRes(tt) = rs.JudgRes '2002/01/25 S.Sano
    typ_CType.JudgRrg(tt) = rs.JudgRrg '2002/01/25 S.Sano
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
    'チェック用AN温度を追加
'Chg Start 2011/08/12 Y.Hitomi
    typ_CType.JudgAntnp(tt) = rs.JudgAntnp
'    If tt = SxlTop Or tt = SxlTail Then typ_CType.JudgAntnp(tt) = rs.JudgAntnp
'Chg End   2011/08/12 Y.Hitomi

'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'--------------- 2008/08/25 INSERT START  By Systech --------------
'Chg Start 2011/03/25 Y.Hitomi
    typ_CType.JudgDkTmp(tt) = rs.JudgDkTmp
'    If tt = SxlTop Or tt = SxlTail Then typ_CType.JudgDkTmp(tt) = rs.JudgDkTmp
'Chg End   2011/03/25 Y.Hitomi

'--------------- 2008/08/25 INSERT  END   By Systech --------------

'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
'    bJudg = (rs.JudgRes And rs.JudgRrg)  '2002/01/25 S.Sano
'--------------- 2008/08/25 UPDATE START  By Systech --------------
'    bJudg = (rs.JudgRes And rs.JudgRrg And rs.JudgAntnp)  '2002/01/25 S.Sano
    bJudg = (rs.JudgRes And rs.JudgRrg And rs.JudgAntnp And rs.JudgDkTmp)
'--------------- 2008/08/25 UPDATE  END   By Systech --------------
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    
    typ_CType.typ_y013(tt, WFRES).MESDATA6 = rs.RRG '2002/01/25 S.Sano
    If wiSmpGetFlg = 0 Then
        With typ_CType
            '実行偏析用パラメータ取得 マルチ引上対応 参照関数変更 2008/04/23 SETsw Nakada
            If GetCoeffParams_new(.typ_hage(tt).CRYNUM, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then
'            If GetCoeffParams(.typ_hage(tt).CRYNUM, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then
                Debug.Print "偏析計算用パラメータの取得に失敗した"
            End If
            
            '偏析係数計算
            cc.DUNMENSEKI = AreaOfCircle(DM)
            cc.TOPSMPLPOS = .typ_Param.INGOTPOS
            cc.BOTSMPLPOS = .typ_Param.INGOTPOS + .typ_Param.LENGTH
            cc.CHARGEWEIGHT = wgtCharge
            cc.TOPWEIGHT = wgtTop + wgtTopCut
            cc.TOPRES = .typ_y013(SxlTop, WFRES).MESDATA5
            cc.BOTRES = .typ_y013(SxlTail, WFRES).MESDATA5
            .COEF(tt) = CoefficientCalculation(cc)
    
            If rs.JudgRes <> True Then
                '偏析計算から再カット位置を計算
                rp.COEFFICIENT = .COEF(tt)
                rp.DUNMENSEKI = AreaOfCircle(DM)
                rp.CHARGEWEIGHT = wgtCharge
                rp.TOPWEIGHT = wgtTop + wgtTopCut
                rp.TOPSMPLPOS = IIf(tt = SxlTop, .typ_Param.INGOTPOS, .typ_Param.INGOTPOS + .typ_Param.LENGTH)
                rp.TOPRES = .typ_y013(tt, WFRES).MESDATA5
                rp.target = IIf(tt = SxlTop, .typ_si.HWFRMAX, .typ_si.HWFRMIN)
                dblScut = PosCalculation(rp)
            Else                                                                            '2002/01/25 S.Sano
                If tt = 1 Then                                                              '2002/01/25 S.Sano
                    dblScut = typ_CType.typ_Param.INGOTPOS                                  '2002/01/25 S.Sano
                Else                                                                        '2002/01/25 S.Sano
                    dblScut = typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH     '2002/01/25 S.Sano
                End If                                                                      '2002/01/25 S.Sano
            End If
        End With
    End If
    
    WfCrResJudg = True
    
End Function

'概要      :OI判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :仕様情報構造体
'          :typ_y013      ,I  ,typ_TBCMY013                         :OI実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(0:判定OK, 1:判定NG)
'          :sSxlPos       ,I  ,String                               :サンプル種別("MID"＝中間抜試)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :OI判定を行う
'履歴      :
'Cng Start 2011/08/01 Y.Hitomi
Public Function WfCrOiJudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                           typ_y013 As typ_TBCMY013, _
                           bJudg As Boolean, _
                           Optional sSxlPos As String) As Boolean
'Public Function WfCrOiJudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
'                           typ_y013 As typ_TBCMY013, _
'                           bJudg As Boolean) As Boolean
'Cng End 2011/08/01 Y.Hitomi
    
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim WOi     As W_OI                     'OI構造体
        
    bJudg = True
        
    'OI判定引数設定
    WOi.GuaranteeOi.cMeth = typ_si.HWFONSPH   '品ＷＦ酸素濃度測定位置＿方
    WOi.GuaranteeOi.cCount = typ_si.HWFONSPT  '品ＷＦ酸素濃度測定位置＿点
    WOi.GuaranteeOi.cPos = typ_si.HWFONSPI    '品ＷＦ酸素濃度測定位置＿位
    WOi.GuaranteeOi.cObj = typ_si.HWFONHWT    '品ＷＦ酸素濃度保証方法＿対
    WOi.GuaranteeOi.cJudg = typ_si.HWFONHWS   '品ＷＦ酸素濃度保証方法＿処
    WOi.GuaranteeCal = typ_si.HWFONMCL        '品ＷＦ酸素濃度面内計算 2001/11/08 S.Sano
    WOi.SpecOiMin = typ_si.HWFONMIN           '品WF酸素濃度下限
    WOi.SpecOiMax = typ_si.HWFONMAX           '品WF酸素濃度上限
    WOi.SpecORG = typ_si.HWFONMBP             '品WF酸素濃度面内分布
    WOi.SpecOiAveMin = typ_si.HWFONAMN        '品WF酸素濃度平均下限
    WOi.SpecOiAveMax = typ_si.HWFONAMX        '品WF酸素濃度平均上限
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
    WOi.Antnp = typ_si.HWFANTNP               '品ＷＦＡＮ温度
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    
    WOi.Oi(0) = NtoZ2(typ_y013.MESDATA1)               'Oi測定値
    WOi.Oi(1) = NtoZ2(typ_y013.MESDATA2)               'Oi測定値
    WOi.Oi(2) = NtoZ2(typ_y013.MESDATA3)               'Oi測定値
    WOi.Oi(3) = NtoZ2(typ_y013.MESDATA4)               'Oi測定値
    WOi.Oi(4) = NtoZ2(typ_y013.MESDATA5)               'Oi測定値
    WOi.Oi(5) = NtoZ2(typ_y013.MESDATA6)               'Oi測定値
    WOi.Oi(6) = NtoZ2(typ_y013.MESDATA7)               'Oi測定値
    WOi.Oi(7) = NtoZ2(typ_y013.MESDATA8)               'Oi測定値
    WOi.Oi(8) = NtoZ2(typ_y013.MESDATA9)               'Oi測定値
    WOi.Oi(9) = NtoZ2(typ_y013.MESDATA10)              'Oi測定値
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
    'チェック用AN温度を追加
    WOi.OiAntnp = NtoZ2(Mid(typ_y013.DKAN, 3, 4))      'Oi測定値
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    
    'OI判定
    If WfOiJudg(WOi, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrOiJudg = False
        Exit Function
    End If
    
    typ_y013.MESDATA11 = WOi.ORG '2002/01/25 S.Sano
    typ_y013.MESDATA12 = WOi.OiMin '2002/01/25 S.Sano
    typ_y013.MESDATA13 = WOi.OiMax '2002/01/25 S.Sano
    
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
'Cng Start 2011/08/12 Y.Hitomi
'    If sSxlPos = "MID" Then
'        bJudg = (WOi.JudgOi And WOi.JudgOrg)
'    Else
'        bJudg = (WOi.JudgOi And WOi.JudgOrg And WOi.JudgAntnp)
'    End If
'    bJudg = (WOi.JudgOi And WOi.JudgOrg) '2002/01/25 S.Sano
    bJudg = (WOi.JudgOi And WOi.JudgOrg And WOi.JudgAntnp) '2002/01/25 S.Sano
'Cng Start 2011/08/12 Y.Hitomi
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    
    WfCrOiJudg = True
End Function

'概要      :BMD判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :仕様情報構造体
'          :typ_y013      ,I  ,typ_TBCMY013                         :BMD実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(0:判定OK, 1:判定NG)
'          :bmflg         ,I  ,Integer                              :BMDﾌﾗｸﾞ(1:BMD1, 2:BMD2, 3:BMD3)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :BMD判定を行う
'履歴      :
Public Function WfCrBmdJudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            typ_y013 As typ_TBCMY013, _
                            bJudg As Boolean, _
                            bmflg As Integer) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim bm      As W_BMD                    'BMD構造体
    Dim c0      As Integer
    
    'Const keisu As Double = 10
    '' 2006/09/25 SMP)kondoh Del -s-
''    Const keisu As Double = 1        'BMDべき乗数変更対応　2003/05/19 osawa
    '' 2006/09/25 SMP)kondoh Del -e-
    '' 2006/09/25 SMP)kondoh Add -s-
    Dim keisu As Double
    Const keisu1 As Double = 10000
    Const keisu2 As Double = 10000
    Const keisu3 As Double = 10000
    Const keisu4 As Double = 10000
    Const keisu5 As Double = 10000
    Const keisu6 As Double = 333000
    Const keisu7 As Double = 10000
    Const keisu8 As Double = 10000
    Const keisu9 As Double = 10000 'Add 2012/07/20 Y.Hitomi

    '' 2006/09/25 SMP)kondoh Add -e-

    bJudg = True

    'BMD判定引数設定
    Select Case bmflg
    Case 1
        bm.GuaranteeBmd.cMeth = typ_si.HWFBM1SH   '品ＷＦＢＭＤ１測定位置＿方
        bm.GuaranteeBmd.cCount = typ_si.HWFBM1ST  '品ＷＦＢＭＤ１測定位置＿点
        bm.GuaranteeBmd.cPos = typ_si.HWFBM1SR    '品ＷＦＢＭＤ１測定位置＿領
        bm.GuaranteeBmd.cObj = typ_si.HWFBM1HT    '品ＷＦＢＭＤ１保証方法＿対
        bm.GuaranteeBmd.cJudg = typ_si.HWFBM1HS   '品ＷＦＢＭＤ１保証方法＿処
        bm.SpecBmdAveMin = typ_si.HWFBM1AN        '品ＷＦＢＭＤ１平均下限
        bm.SpecBmdAveMax = typ_si.HWFBM1AX        '品ＷＦＢＭＤ１平均上限
        bm.SpecBmdMBP = typ_si.HWFBM1MBP          '品ＷＦＢＭＤ１面内分布　2003/05/20 ooba
        bm.SpecBmdMCL = NtoS(typ_si.HWFBM1MCL)    '品ＷＦＢＭＤ１面内計算　2003/05/20 ooba
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        bm.Antnp = typ_si.HWFANTNP                '品ＷＦＡＮ温度
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    Case 2
        bm.GuaranteeBmd.cMeth = typ_si.HWFBM2SH   '品ＷＦＢＭＤ２測定位置＿方
        bm.GuaranteeBmd.cCount = typ_si.HWFBM2ST  '品ＷＦＢＭＤ２測定位置＿点
        bm.GuaranteeBmd.cPos = typ_si.HWFBM2SR    '品ＷＦＢＭＤ２測定位置＿領
        bm.GuaranteeBmd.cObj = typ_si.HWFBM2HT    '品ＷＦＢＭＤ２保証方法＿対
        bm.GuaranteeBmd.cJudg = typ_si.HWFBM2HS   '品ＷＦＢＭＤ２保証方法＿処
        bm.SpecBmdAveMin = typ_si.HWFBM2AN        '品ＷＦＢＭＤ２平均下限
        bm.SpecBmdAveMax = typ_si.HWFBM2AX        '品ＷＦＢＭＤ２平均上限
        bm.SpecBmdMBP = typ_si.HWFBM2MBP          '品ＷＦＢＭＤ２面内分布　2003/05/20 ooba
        bm.SpecBmdMCL = NtoS(typ_si.HWFBM2MCL)    '品ＷＦＢＭＤ２面内計算　2003/05/20 ooba
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        bm.Antnp = typ_si.HWFANTNP                '品ＷＦＡＮ温度
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    Case 3
        bm.GuaranteeBmd.cMeth = typ_si.HWFBM3SH   '品ＷＦＢＭＤ３測定位置＿方
        bm.GuaranteeBmd.cCount = typ_si.HWFBM3ST  '品ＷＦＢＭＤ３測定位置＿点
        bm.GuaranteeBmd.cPos = typ_si.HWFBM3SR    '品ＷＦＢＭＤ３測定位置＿領
        bm.GuaranteeBmd.cObj = typ_si.HWFBM3HT    '品ＷＦＢＭＤ３保証方法＿対
        bm.GuaranteeBmd.cJudg = typ_si.HWFBM3HS   '品ＷＦＢＭＤ３保証方法＿処
        bm.SpecBmdAveMin = typ_si.HWFBM3AN        '品ＷＦＢＭＤ３平均下限
        bm.SpecBmdAveMax = typ_si.HWFBM3AX        '品ＷＦＢＭＤ３平均上限
        bm.SpecBmdMBP = typ_si.HWFBM3MBP          '品ＷＦＢＭＤ３面内分布　2003/05/20 ooba
        bm.SpecBmdMCL = NtoS(typ_si.HWFBM3MCL)    '品ＷＦＢＭＤ３面内計算　2003/05/20 ooba
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        bm.Antnp = typ_si.HWFANTNP                '品ＷＦＡＮ温度
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    End Select

    '' 2006/09/25 SMP)kondoh Add -s-
    If bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "4" And bm.GuaranteeBmd.cPos = "H" Then
        keisu = keisu1
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "3" And bm.GuaranteeBmd.cPos = "H" Then
        keisu = keisu2
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "4" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu3
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "3" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu4
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "5" And bm.GuaranteeBmd.cPos = "A" Then
        keisu = keisu5
    ElseIf bm.GuaranteeBmd.cMeth = "G" And bm.GuaranteeBmd.cCount = "3" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu6
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "5" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu7
    ElseIf bm.GuaranteeBmd.cMeth = "8" And bm.GuaranteeBmd.cCount = "4" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu8
    'Add Start 2012/07/20 Y.Hitomi
    ElseIf bm.GuaranteeBmd.cMeth = "P" Then
        keisu = keisu9
    'Add End 2012/07/20 Y.Hitomi
    
    Else
        bJudg = False
        WfCrBmdJudg = False
        Exit Function
    End If
    '' 2006/09/25 SMP)kondoh Add -e-

    With bm
        .BMD(0) = NtoZ2(typ_y013.MESDATA1)                   'BMD測定値
        .BMD(1) = NtoZ2(typ_y013.MESDATA2)                   'BMD測定値
        .BMD(2) = NtoZ2(typ_y013.MESDATA3)                   'BMD測定値
        .BMD(3) = NtoZ2(typ_y013.MESDATA4)                   'BMD測定値
        .BMD(4) = NtoZ2(typ_y013.MESDATA5)                   'BMD測定値　2003/05/20 ooba
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        'チェック用AN温度を追加
        .BmdAntnp = NtoZ2(Mid(typ_y013.DKAN, 3, 4))
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------

        For c0 = 0 To 4                                      ' 2003/05/20 ooba
    '' 2006/09/25 SMP)kondoh Cng -s-
''            .BMD(c0) = IIf(.BMD(c0) <> -1, .BMD(c0) * keisu, -1)
            .BMD(c0) = IIf(.BMD(c0) <> -1, .BMD(c0) * CDbl(keisu / 10000), -1)
    '' 2006/09/25 SMP)kondoh Cng -e-
        Next
    End With
    
    'BMD判定
    If WfBMDJudg(bm, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrBmdJudg = False
        Exit Function
    End If
    
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
'    If bm.JudgBmd <> True Then
    If bm.JudgBmd <> True Or bm.JudgAntnp <> True Then
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        bJudg = False
    End If
    
    typ_y013.MESDATA6 = bm.JudgDataAve            '　▼2003/05/20 ooba
    typ_y013.MESDATA7 = bm.JudgDataMax
    typ_y013.MESDATA8 = bm.JudgDataMin
    typ_y013.MESDATA9 = bm.JudgDataMBP            '　▲2003/05/20 ooba
     
    WfCrBmdJudg = True
End Function

'概要      :OSF判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :仕様情報構造体
'          :typ_y013      ,I  ,typ_TBCMY013                         :OSF実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(0:判定OK, 1:判定NG)
'          :osfflg        ,I  ,Integer                              :OSFﾌﾗｸﾞ(1:OSF1, 2:OSF2, 3:OSF3, 4:OSF4)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :OSF判定を行う
'履歴      :
Public Function WfCrOsfJudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            typ_y013 As typ_TBCMY013, _
                            bJudg As Boolean, _
                            osfflg As Integer, _
                            TmpData() As String) As Boolean

    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim os      As W_OSF                    'OSF構造体
    Dim keisu   As Double
    Dim c0      As Integer
    
    Const keisu1 As Double = 1.8248175
    Const keisu2 As Double = 1.8518519
    Const keisu3 As Double = 1.9230769
    Const keisu4 As Double = 3.649635
    Const keisu5 As Double = 3.7037037
    Const keisu6 As Double = 3.8461538
    '' 2006/09/25 SMP)kondoh Add -s-
    Const keisu7 As Double = 7.6923077
    '' 2006/09/25 SMP)kondoh Add -e-
        
    bJudg = True

    'OSF判定引数設定
    Select Case osfflg
    Case 1
        os.GuaranteeOsf.cMeth = typ_si.HWFOF1SH  '品ＷＦＯＳＦ１測定位置＿方
        os.GuaranteeOsf.cCount = typ_si.HWFOF1ST '品ＷＦＯＳＦ１測定位置＿点
        os.GuaranteeOsf.cPos = typ_si.HWFOF1SR   '品ＷＦＯＳＦ１測定位置＿領
        os.GuaranteeOsf.cObj = typ_si.HWFOF1HT   '品ＷＦＢＭＤ１保証方法＿対
        os.GuaranteeOsf.cJudg = typ_si.HWFOF1HS  '品ＷＦＢＭＤ１保証方法＿処
        os.SpecOsfAveMax = typ_si.HWFOF1AX       '品ＷＦＯＳＦ１平均上限
        os.SpecOsfMax = typ_si.HWFOF1MX          '品ＷＦＯＳＦ１上限
        os.JudgDataPTK = NtoS(typ_si.HWFOSF1PTK) '品ＷＦＯＳＦ１パタン区分　2003/05/17 ooba
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        os.Antnp = typ_si.HWFANTNP                '品ＷＦＡＮ温度
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    Case 2
        os.GuaranteeOsf.cMeth = typ_si.HWFOF2SH  '品ＷＦＯＳＦ２測定位置＿方
        os.GuaranteeOsf.cCount = typ_si.HWFOF2ST '品ＷＦＯＳＦ２測定位置＿点
        os.GuaranteeOsf.cPos = typ_si.HWFOF2SR   '品ＷＦＯＳＦ２測定位置＿領
        os.GuaranteeOsf.cObj = typ_si.HWFOF2HT   '品ＷＦＢＭＤ２保証方法＿対
        os.GuaranteeOsf.cJudg = typ_si.HWFOF2HS  '品ＷＦＢＭＤ２保証方法＿処
        os.SpecOsfAveMax = typ_si.HWFOF2AX       '品ＷＦＯＳＦ２平均上限
        os.SpecOsfMax = typ_si.HWFOF2MX          '品ＷＦＯＳＦ２上限
        os.JudgDataPTK = NtoS(typ_si.HWFOSF2PTK) '品ＷＦＯＳＦ２パタン区分　2003/05/17 ooba
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        os.Antnp = typ_si.HWFANTNP                '品ＷＦＡＮ温度
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    Case 3
        os.GuaranteeOsf.cMeth = typ_si.HWFOF3SH  '品ＷＦＯＳＦ３測定位置＿方
        os.GuaranteeOsf.cCount = typ_si.HWFOF3ST '品ＷＦＯＳＦ３測定位置＿点
        os.GuaranteeOsf.cPos = typ_si.HWFOF3SR   '品ＷＦＯＳＦ３測定位置＿領
        os.GuaranteeOsf.cObj = typ_si.HWFOF3HT   '品ＷＦＢＭＤ３保証方法＿対
        os.GuaranteeOsf.cJudg = typ_si.HWFOF3HS  '品ＷＦＢＭＤ３保証方法＿処
        os.SpecOsfAveMax = typ_si.HWFOF3AX       '品ＷＦＯＳＦ３平均上限
        os.SpecOsfMax = typ_si.HWFOF3MX          '品ＷＦＯＳＦ３上限
        os.JudgDataPTK = NtoS(typ_si.HWFOSF3PTK) '品ＷＦＯＳＦ３パタン区分　2003/05/17 ooba
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        os.Antnp = typ_si.HWFANTNP                '品ＷＦＡＮ温度
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    Case 4
        os.GuaranteeOsf.cMeth = typ_si.HWFOF4SH  '品ＷＦＯＳＦ４測定位置＿方
        os.GuaranteeOsf.cCount = typ_si.HWFOF4ST '品ＷＦＯＳＦ４測定位置＿点
        os.GuaranteeOsf.cPos = typ_si.HWFOF4SR   '品ＷＦＯＳＦ４測定位置＿領
        os.GuaranteeOsf.cObj = typ_si.HWFOF4HT   '品ＷＦＢＭＤ４保証方法＿対
        os.GuaranteeOsf.cJudg = typ_si.HWFOF4HS  '品ＷＦＢＭＤ４保証方法＿処
        os.SpecOsfAveMax = typ_si.HWFOF4AX       '品ＷＦＯＳＦ４平均上限
        os.SpecOsfMax = typ_si.HWFOF4MX          '品ＷＦＯＳＦ４上限
        os.JudgDataPTK = NtoS(typ_si.HWFOSF4PTK) '品ＷＦＯＳＦ４パタン区分　2003/05/17 ooba
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        os.Antnp = typ_si.HWFANTNP                '品ＷＦＡＮ温度
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    End Select
    
    If os.GuaranteeOsf.cMeth = "5" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "3" Then
        keisu = keisu1
    ElseIf os.GuaranteeOsf.cMeth = "5" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "5" Then
        keisu = keisu2
    ElseIf os.GuaranteeOsf.cMeth = "5" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "A" Then
        keisu = keisu3
    ElseIf os.GuaranteeOsf.cMeth = "6" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "3" Then
        keisu = keisu4
    ElseIf os.GuaranteeOsf.cMeth = "6" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "5" Then
        keisu = keisu5
    ElseIf os.GuaranteeOsf.cMeth = "6" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "A" Then
        keisu = keisu6
    '' 2006/09/25 SMP)kondoh Add -s-
    ElseIf os.GuaranteeOsf.cMeth = "E" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "A" Then
        keisu = keisu7
    '' 2006/09/25 SMP)kondoh Add -e-
    Else
        bJudg = False
        WfCrOsfJudg = False
        Exit Function
    End If

    With os
        .OSF(0) = NtoZ2(typ_y013.MESDATA1)                   'OSF測定値
        .OSF(1) = NtoZ2(typ_y013.MESDATA2)                   'OSF測定値
        .OSF(2) = NtoZ2(typ_y013.MESDATA3)                   'OSF測定値
        .OSF(3) = NtoZ2(typ_y013.MESDATA4)                   'OSF測定値
        .OSF(4) = NtoZ2(typ_y013.MESDATA5)                   'OSF測定値
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        'チェック用AN温度を追加
        .OsfAntnp = NtoZ2(Mid(typ_y013.DKAN, 3, 4))
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        For c0 = 0 To 4
            .OSF(c0) = IIf(.OSF(c0) <> -1, .OSF(c0) * keisu, -1)
        Next
        typ_y013.MESDATA6 = typ_y013.MESDATA6 * 100
        .OSFp(0) = Trim(typ_y013.MESDATA9)                   'OSFパターン実績(大)　▼2003/05/17 ooba
        .OSFp(1) = Trim(typ_y013.MESDATA12)                  'OSFパターン実績(中)
        .OSFp(2) = Trim(typ_y013.MESDATA15)                  'OSFパターン実績(小)　▲2003/05/17 ooba
    End With
    typ_y013.MESDATA6 = typ_y013.MESDATA6 * 100
    
    'OSF判定
    If WfOSFJudg(os, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrOsfJudg = False
        Exit Function
    End If
    
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
'    If os.JudgOsf <> True Then
    If os.JudgOsf <> True Or os.JudgAntnp <> True Then
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        bJudg = False
    End If
    
    TmpData(0) = os.JudgDataAve                                  ' 2003/05/20 ooba
    TmpData(1) = os.JudgDataMax                                  ' 2003/05/20 ooba
'    typ_y013.MESDATA7 = os.JudgDataAve
'    typ_y013.MESDATA8 = os.JudgDataMax
     WfCrOsfJudg = True
End Function

'概要      :DSOD判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :仕様情報構造体
'          :typ_y013      ,I  ,typ_TBCMY013                         :DSOD実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(0:判定OK, 1:判定NG)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :CS判定を行う
'履歴      :
Public Function WfCrDsodjudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                             typ_y013 As typ_TBCMY013, _
                             bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim Dsod    As W_DSOD                   'W_DSOD構造体
    
    bJudg = True
        
    'DSOD判定引数設定
    Dsod.GuaranteeDsod.cMeth = ""  '
    Dsod.GuaranteeDsod.cCount = "" '
    Dsod.GuaranteeDsod.cPos = ""  '
    Dsod.GuaranteeDsod.cObj = typ_si.HWFDSOHT    '品ＷＦＤＳＯＤ保証方法＿対
    Dsod.GuaranteeDsod.cJudg = typ_si.HWFDSOHS   '品ＷＦＤＳＯＤ保証方法＿処
    Dsod.SpecDsodMin = typ_si.HWFDSOMN           '品ＷＦＤＳＯＤ下限
    Dsod.SpecDsodMax = typ_si.HWFDSOMX           '品ＷＦＤＳＯＤ上限
    Dsod.JudgDataPTK = NtoS(typ_si.HWFDSOPTK)    '品ＷＦＤＳＯＤパタン区分　04/07/28 ooba
    
    Dsod.Dsod = NtoZ2(typ_y013.MESDATA1)         'DSOD測定値
    Dsod.Dsodp(0) = Trim(typ_y013.MESDATA4)      'DSODパタン実績1　04/07/28 ooba
    Dsod.Dsodp(1) = Trim(typ_y013.MESDATA7)      'DSODパタン実績2　04/07/28 ooba
    
    Dsod.Antnp = typ_si.HWFANTNP                        '品WFAN温度(仕様)　06/12/22 ooba
    Dsod.DsodAntnp = NtoZ2(Mid(typ_y013.DKAN, 3, 4))    '品WFAN温度(実績)　06/12/22 ooba
    
    'DSOD判定
    If WfDSODJudg(Dsod, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrDsodjudg = False
        Exit Function
    End If
    
'    If Dsod.JudgDsod <> True Then
    If Dsod.JudgDsod <> True Or Dsod.JudgAntnp <> True Then  'AN温度判定結果追加　06/12/22 ooba
        bJudg = False
    End If
    
    WfCrDsodjudg = True
End Function

'概要      :DZ判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :仕様情報構造体
'          :typ_y013      ,I  ,typ_TBCMY013                         :DZ実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(0:判定OK, 1:判定NG)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :DZ判定を行う
'履歴      :
Public Function WfCrDzjudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                           typ_y013 As typ_TBCMY013, _
                           bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim DZ      As W_DZ                     'DZ構造体
    
    bJudg = True
        
    'DZ判定引数設定
    DZ.GuaranteeDz.cMeth = typ_si.HWFMKSPH   '品ＷＦ無欠陥層測定位置＿方
    DZ.GuaranteeDz.cCount = typ_si.HWFMKSPT  '品ＷＦ無欠陥層測定位置＿点
    DZ.GuaranteeDz.cPos = typ_si.HWFMKSPR    '品ＷＦ無欠陥層測定位置＿領
    DZ.GuaranteeDz.cObj = typ_si.HWFMKHWT    '品ＷＦ無欠陥層保証方法＿対
    DZ.GuaranteeDz.cJudg = typ_si.HWFMKHWS   '品ＷＦ無欠陥層保証方法＿処
    DZ.SpecDzMin = typ_si.HWFMKMIN           '品ＷＦ無欠陥層下限
    DZ.SpecDzMax = typ_si.HWFMKMAX           '品ＷＦ無欠陥層上限
    
    DZ.DZ(0) = NtoZ2(typ_y013.MESDATA1)               'DZ測定値
    DZ.DZ(1) = NtoZ2(typ_y013.MESDATA2)               'DZ測定値
    DZ.DZ(2) = NtoZ2(typ_y013.MESDATA3)               'DZ測定値
    DZ.DZ(3) = NtoZ2(typ_y013.MESDATA4)               'DZ測定値
    
    'DZ判定
    If WfDZJudg(DZ, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrDzjudg = False
        Exit Function
    End If
    
    If DZ.JudgDz <> True Then
        bJudg = False
    End If
        
    typ_y013.MESDATA5 = DZ.JudgDataAve
    typ_y013.MESDATA6 = DZ.JudgDataMax
    typ_y013.MESDATA7 = DZ.JudgDataMin
    
    WfCrDzjudg = True
End Function

'概要      :SPVFE判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :仕様情報構造体
'          :typ_y013      ,I  ,typ_TBCMY013                         :SPVFE実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(0:判定OK, 1:判定NG)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :SPV判定を行う
'履歴      :
Public Function WfCrSpvjudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            typ_y013 As typ_TBCMY013, _
                            bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim sp      As W_SPV                    'SPV構造体
    
    bJudg = True
        
    'SPV判定引数設定
    sp.GuaranteeSpv.cMeth = typ_si.HWFDLSPH    '品ＷＦ拡散長測定位置＿方
    sp.GuaranteeSpv.cCount = typ_si.HWFDLSPT   '品ＷＦ拡散長測定位置＿点
    sp.GuaranteeSpv.cPos = typ_si.HWFDLSPI     '品ＷＦ拡散長測定位置＿位
    sp.GuaranteeSpv.cObj = typ_si.HWFDLHWT     '品ＷＦ拡散長保証方法＿対
    sp.GuaranteeSpv.cJudg = typ_si.HWFDLHWS    '品ＷＦ拡散長保証方法＿処
    
    sp.GuaranteeSpvFe.cMeth = typ_si.HWFSPVSH  '品ＷＦＳＰＶＦＥ測定位置＿方
    sp.GuaranteeSpvFe.cCount = typ_si.HWFSPVST '品ＷＦＳＰＶＦＥ測定位置＿点
    sp.GuaranteeSpvFe.cPos = typ_si.HWFSPVSI   '品ＷＦＳＰＶＦＥ測定位置＿位
    sp.GuaranteeSpvFe.cObj = typ_si.HWFSPVHT   '品ＷＦＳＰＶＦＥ保証方法＿対
    sp.GuaranteeSpvFe.cJudg = typ_si.HWFSPVHS  '品ＷＦＳＰＶＦＥ保証方法＿処
    
    sp.SpecSpvMin = typ_si.HWFDLMIN            '品WF拡散長下限
    sp.SpecSpvMax = typ_si.HWFDLMAX            '品WF拡散長上限
    sp.SpecSpvFeMax = typ_si.HWFSPVMX          '品WFFe濃度上限
    '----TEST2004/10
    sp.SpecSpvAvMax = typ_si.HWFSPVAM
    
    sp.Spv(0) = NtoZ2(typ_y013.MESDATA1)                'SPV測定値
    sp.Spv(1) = NtoZ2(typ_y013.MESDATA2)                'SPV測定値
    sp.Spv(2) = NtoZ2(typ_y013.MESDATA3)                'SPV測定値
    sp.Spv(3) = NtoZ2(typ_y013.MESDATA4)                'SPV測定値
    sp.Spv(4) = NtoZ2(typ_y013.MESDATA5)                'SPV測定値
    
    'SPV判定
    If WfSPVJudg(sp, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrSpvjudg = False
        Exit Function
    End If
    
    If sp.JudgSpv <> True Then
        bJudg = False
    End If
    
    WfCrSpvjudg = True
End Function

'概要      :DOI判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :仕様情報構造体
'          :typ_y013      ,I  ,typ_TBCMY013                         :DOI実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(0:判定OK, 1:判定NG)
'          :doiflg        ,I  ,Integer                              :DOIﾌﾗｸﾞ(1:DOI1, 2:DOI2, 3:DOI3)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :DOI判定を行う
'履歴      :
Public Function WfCrDoiJudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            typ_y013 As typ_TBCMY013, _
                            bJudg As Boolean, _
                            doiflg As Integer) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim WFDOI   As W_DOI                    'DOI構造体

    bJudg = True

    'DOI判定引数設定
    Select Case doiflg
    Case 1
        WFDOI.GuaranteeDoi.cMeth = typ_si.HWFOS1SH    '品ＷＦ酸素析出１測定位置＿方
        WFDOI.GuaranteeDoi.cCount = typ_si.HWFOS1ST   '品ＷＦ酸素析出１測定位置＿点
        WFDOI.GuaranteeDoi.cPos = typ_si.HWFOS1SI     '品ＷＦ酸素析出１測定位置＿位
        WFDOI.GuaranteeDoi.cObj = typ_si.HWFOS1HT     '品ＷＦ酸素析出１保証方法＿対
        WFDOI.GuaranteeDoi.cJudg = typ_si.HWFOS1HS    '品ＷＦ酸素析出１保証方法＿処
        WFDOI.SpecDoiMin = typ_si.HWFOS1MN            '品ＷＦ酸素析出１下限
        WFDOI.SpecDoiMax = typ_si.HWFOS1MX            '品ＷＦ酸素析出１上限
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        WFDOI.Antnp = typ_si.HWFANTNP                 '品ＷＦＡＮ温度
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    Case 2
        WFDOI.GuaranteeDoi.cMeth = typ_si.HWFOS2SH    '品ＷＦ酸素析出１測定位置＿方
        WFDOI.GuaranteeDoi.cCount = typ_si.HWFOS2ST   '品ＷＦ酸素析出１測定位置＿点
        WFDOI.GuaranteeDoi.cPos = typ_si.HWFOS2SI     '品ＷＦ酸素析出１測定位置＿位
        WFDOI.GuaranteeDoi.cObj = typ_si.HWFOS2HT     '品ＷＦ酸素析出１保証方法＿対
        WFDOI.GuaranteeDoi.cJudg = typ_si.HWFOS2HS    '品ＷＦ酸素析出１保証方法＿処
        WFDOI.SpecDoiMin = typ_si.HWFOS2MN            '品ＷＦ酸素析出１下限
        WFDOI.SpecDoiMax = typ_si.HWFOS2MX            '品ＷＦ酸素析出１上限
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        WFDOI.Antnp = typ_si.HWFANTNP                 '品ＷＦＡＮ温度
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    Case 3
        WFDOI.GuaranteeDoi.cMeth = typ_si.HWFOS3SH    '品ＷＦ酸素析出１測定位置＿方
        WFDOI.GuaranteeDoi.cCount = typ_si.HWFOS3ST   '品ＷＦ酸素析出１測定位置＿点
        WFDOI.GuaranteeDoi.cPos = typ_si.HWFOS3SI     '品ＷＦ酸素析出１測定位置＿位
        WFDOI.GuaranteeDoi.cObj = typ_si.HWFOS3HT     '品ＷＦ酸素析出１保証方法＿対
        WFDOI.GuaranteeDoi.cJudg = typ_si.HWFOS3HS    '品ＷＦ酸素析出１保証方法＿処
        WFDOI.SpecDoiMin = typ_si.HWFOS3MN            '品ＷＦ酸素析出１下限
        WFDOI.SpecDoiMax = typ_si.HWFOS3MX            '品ＷＦ酸素析出１上限
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        WFDOI.Antnp = typ_si.HWFANTNP                 '品ＷＦＡＮ温度
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    End Select
    
    WFDOI.Doi(0) = NtoZ2(typ_y013.MESDATA1)                    'DOI測定値
    WFDOI.Doi(1) = NtoZ2(typ_y013.MESDATA2)                    'DOI測定値
    WFDOI.Doi(2) = NtoZ2(typ_y013.MESDATA3)                    'DOI測定値
    WFDOI.Doi(3) = NtoZ2(typ_y013.MESDATA4)                    'DOI測定値
    WFDOI.Doi(4) = NtoZ2(typ_y013.MESDATA5)                    'DOI測定値
    WFDOI.Doi(5) = NtoZ2(typ_y013.MESDATA6)                    'DOI測定値　-*-*-　20010912 add
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    'チェック用AN温度を追加
    WFDOI.DoiAntnp = NtoZ2(Mid(typ_y013.DKAN, 3, 4))
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    
    'DOI判定
    If WfDOiJudg(WFDOI, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrDoiJudg = False
        Exit Function
    End If
    
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
'    If WFDOI.JudgDoi <> True Then
    If WFDOI.JudgDoi <> True Or WFDOI.JudgAntnp <> True Then
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        bJudg = False
    End If
        
    WfCrDoiJudg = True
End Function
'''''============================================================================================================================
'''''
''''''概要      :実績情報データ設定
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
''''''          :typ_a         ,I  ,typ_AllTypesC ,各情報構造体
''''''説明      :評価情報を情報構造体に設定する
''''''履歴      :
'''''Public Function JudgAllRsltData() As FUNCTION_RETURN
'''''
'''''    TotalJudg = True
'''''
''''''''''    typ_rtInit '2001/09/14 S.Sano
'''''
'''''    JudgAllRsltData = FUNCTION_RETURN_FAILURE
'''''
'''''    '仕様検査支持取得
'''''    SpecJudgCheck
'''''
'''''
''''''''''    WFCJudgDialog.WFCErrorMessage SelectSxlID & " ******************"
'''''    '実績データ判定(TOP)
'''''    If WfAllJudg(SxlTop) = FUNCTION_RETURN_FAILURE Then
'''''        Exit Function
'''''    End If
'''''    '実績データ判定(TAIL)
'''''    If WfAllJudg(SxlTail) = FUNCTION_RETURN_FAILURE Then
'''''        Exit Function
'''''    End If
'''''
'''''    JudgAllRsltData = FUNCTION_RETURN_SUCCESS
'''''End Function

'概要      :AOI判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :仕様情報構造体
'          :typ_y013      ,I  ,typ_TBCMY013                         :AOI実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(0:判定OK, 1:判定NG)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :AOI判定を行う
'履歴      :03/12/09 ooba
Public Function WfCrAoiJudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            typ_y013 As typ_TBCMY013, _
                            bJudg As Boolean) As Boolean
                            
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim WFAOI   As W_AOI                    'AOI構造体

    bJudg = True

    'AOI判定引数設定
    WFAOI.GuaranteeAoi.cMeth = typ_si.HWFZOSPH    '品ＷＦ残存酸素測定位置＿方
    WFAOI.GuaranteeAoi.cCount = typ_si.HWFZOSPT   '品ＷＦ残存酸素測定位置＿点
    WFAOI.GuaranteeAoi.cPos = typ_si.HWFZOSPI     '品ＷＦ残存酸素測定位置＿位
    WFAOI.GuaranteeAoi.cObj = typ_si.HWFZOHWT     '品ＷＦ残存酸素保証方法＿対
    WFAOI.GuaranteeAoi.cJudg = typ_si.HWFZOHWS    '品ＷＦ残存酸素保証方法＿処
    WFAOI.SpecAoiMin = typ_si.HWFZOMIN            '品ＷＦ残存酸素下限
    WFAOI.SpecAoiMax = typ_si.HWFZOMAX            '品ＷＦ残存酸素上限
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    WFAOI.Antnp = typ_si.HWFANTNP                 '品ＷＦＡＮ温度
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    
    WFAOI.AOI(0) = NtoZ2(typ_y013.MESDATA4)       'AOI測定値
    WFAOI.AOI(1) = NtoZ2(typ_y013.MESDATA5)       'AOI測定値
    WFAOI.AOI(2) = NtoZ2(typ_y013.MESDATA6)       'AOI測定値
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    'チェック用AN温度を追加
    WFAOI.AoiAntnp = NtoZ2(Mid(typ_y013.DKAN, 3, 4))
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    
    'AOI判定
    If WfAOiJudg(WFAOI, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrAoiJudg = False
        Exit Function
    End If

'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.3 AN温度 実績反映チェック追加
'    If WFAOI.JudgAoi <> True Then
    If WFAOI.JudgAoi <> True Or WFAOI.JudgAntnp <> True Then
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        bJudg = False
    End If
        
    WfCrAoiJudg = True
End Function

'概要      :SIRD(SD)判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :仕様情報構造体
'          :typ_y013      ,I  ,typ_TBCMJ022                         :SD実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(0:判定OK, 1:判定NG)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :SD判定を行う
'履歴      :2010/01/07 SIRD対応 Y.Hitomi
Public Function WfCrSdjudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                           typ_j022 As typ_TBCMJ022, _
                           bJudg As Boolean) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim SD      As W_SD                     'SD構造体
    
    bJudg = True
        
    'SD判定引数設定
    SD.GuaranteeSd.cObj = typ_si.HWFSIRDHT   '軸状転位保証方法＿対
    SD.GuaranteeSd.cJudg = typ_si.HWFSIRDHS  '軸状転位保証方法＿処
    SD.SpecSdMax = typ_si.HWFSIRDMX          '軸状転位上限
    
    SD.SdMeasData = val((typ_j022.SIRDCNT))  'SD測定値
    
    'DZ判定
    If WfSDJudg(SD, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrSdjudg = False
        Exit Function
    End If
    
    If SD.JudgSD = False Then
        bJudg = False
    End If
            
    WfCrSdjudg = True
End Function
'概要      :GD判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :仕様情報構造体
'          :typ_j015      ,I  ,typ_TBCMJ015                         :GD実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(0:判定OK, 1:判定NG)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :GD判定を行う
'履歴      :05/01/31 ooba
Public Function WfCrGdJudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            typ_j015 As typ_TBCMJ015, _
                            bJudg As Boolean) As Boolean

    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim WFGD    As W_GD                     'GD構造体
    Dim iCnt    As Integer
    Dim bUpdFlg As Boolean                  'TBCMJ015-UPDATE有無ﾌﾗｸﾞ
    Dim bDenData    As Boolean              'Den実績ﾃﾞｰﾀ存在ﾌﾗｸﾞ　05/10/25 ooba
    Dim bLdlData    As Boolean              'L/DL実績ﾃﾞｰﾀ存在ﾌﾗｸﾞ　05/10/25 ooba
    Dim bDvd2Data   As Boolean              'DVD2実績ﾃﾞｰﾀ存在ﾌﾗｸﾞ　05/10/25 ooba
    Dim SYORIKBN As String                  '初期登録かのフラグ　　10/04/01 hama


'   Dim WFGD2       As W_GD                 'GD構造体   '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
    SYORIKBN = ""
    bJudg = True
    
    'GD判定引数設定
    WFGD.GuaranteeDen.cMeth = ""                '測定位置_方
    WFGD.GuaranteeDen.cCount = ""               '測定位置_点
    WFGD.GuaranteeDen.cPos = ""                 '測定位置_位
    WFGD.GuaranteeDen.cObj = typ_si.HWFDENHT    '保証方法_対
    WFGD.GuaranteeDen.cJudg = typ_si.HWFDENHS   '保証方法_処
    
    WFGD.GuaranteeLdl.cMeth = ""                '測定位置_方
    WFGD.GuaranteeLdl.cCount = ""               '測定位置_点
    WFGD.GuaranteeLdl.cPos = ""                 '測定位置_位
    WFGD.GuaranteeLdl.cObj = typ_si.HWFLDLHT    '保証方法_対
    WFGD.GuaranteeLdl.cJudg = typ_si.HWFLDLHS   '保証方法_処
    
    WFGD.GuaranteeDvd2.cMeth = ""               '測定位置_方
    WFGD.GuaranteeDvd2.cCount = ""              '測定位置_点
    WFGD.GuaranteeDvd2.cPos = ""                '測定位置_位
    WFGD.GuaranteeDvd2.cObj = typ_si.HWFDVDHT   '保証方法_対
    WFGD.GuaranteeDvd2.cJudg = typ_si.HWFDVDHS  '保証方法_処
    
    WFGD.JudgFlagDen = typ_si.HWFDENKU          '品WFDen検査有無
    WFGD.JudgFlagLdl = typ_si.HWFLDLKU          '品WFL/DL検査有無
    WFGD.JudgFlagDvd2 = typ_si.HWFDVDKU         '品WFDVD2検査有無
    
    WFGD.SpecDenMin = typ_si.HWFDENMN           '品WFDen下限
    WFGD.SpecDenMax = typ_si.HWFDENMX           '品WFDen上限
    WFGD.SpecLdlMin = typ_si.HWFLDLMN           '品WFL/DL下限
    WFGD.SpecLdlMax = typ_si.HWFLDLMX           '品WFL/DL上限
    WFGD.SpecDvd2Min = typ_si.HWFDVDMNN         '品WFDVD2下限
    WFGD.SpecDvd2Max = typ_si.HWFDVDMXN         '品WFDVD2上限
    
'*** UPDATE ↓ Y.SIMIZU 2005/10/7 品WFGDﾗｲﾝ数追加
    WFGD.SpecGdLine = typ_si.HWFGDLINE          '品WFGDﾗｲﾝ数
'*** UPDATE ↑ Y.SIMIZU 2005/10/7 品WFGDﾗｲﾝ数追加

    WFGD.Antnp = typ_si.HWFANTNP                        '品WFAN温度(仕様)　06/12/22 ooba
    WFGD.GdAntnp = NtoZ2(Mid(typ_j015.DKAN, 3, 4))      '品WFAN温度(実績)　06/12/22 ooba
    
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
''    If typ_si.WFHSGDCW = "1" Then
''        ' 結晶
''        WFGD.GDPTK = typ_si.HSXGDPTK
''        WFGD.ZeroLdlMin = typ_si.HSXLDLRMN
''        WFGD.ZeroLdlMax = typ_si.HSXLDLRMX
''    Else
        ' WF
        WFGD.GDPTK = typ_si.HWFGDPTK
        WFGD.ZeroLdlMin = typ_si.HWFLDLRMN
        WFGD.ZeroLdlMax = typ_si.HWFLDLRMX
''    End If
    ' 構造体はGD実績(WF)だが、結晶の実績も設定される場合がある
    WFGD.LdlMin = typ_j015.MSZEROMN
    WFGD.LdlMax = typ_j015.MSZEROMX
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    
'*** UPDATE ↓ Y.SIMIZU 2005/10/12 GDﾗｲﾝ数が3又は4.5又は5でない場合は判定ｴﾗｰ
    If WFGD.SpecGdLine <> 3 And WFGD.SpecGdLine <> 4.5 And WFGD.SpecGdLine <> 5 Then
        bJudg = False
        WfCrGdJudg = False
        Exit Function
    End If
'*** UPDATE ↑ Y.SIMIZU 2005/10/12 GDﾗｲﾝ数が3又は4.5又は5でない場合は判定ｴﾗｰ

'*** UPDATE ↓ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数分の実績があるかをﾁｪｯｸする
'    If ChkGD_Data(typ_j015, WFGD) <> FUNCTION_RETURN_SUCCESS Then
    If ChkGD_Data(typ_j015, WFGD, bDenData, bLdlData, bDvd2Data) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrGdJudg = False
        Exit Function
    End If
'*** UPDATE ↑ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数分の実績があるかをﾁｪｯｸする
    
    '測定結果ﾃﾞｰﾀが存在しない場合は計算で求める
    If typ_j015.MSRSDEN = -1 Or typ_j015.MSRSLDL = -1 Or typ_j015.MSRSDVD2 = -1 Then
    '*** UPDATE ↓ Y.SIMIZU 2005/10/7 仕様のGDﾗｲﾝ数,DVD2上下限を渡すように変更
'        If Calculate_GD(typ_j015) <> FUNCTION_RETURN_SUCCESS Then
        If Calculate_GD(typ_j015, WFGD.SpecGdLine, WFGD.SpecDvd2Min, WFGD.SpecDvd2Max) <> FUNCTION_RETURN_SUCCESS Then
    '*** UPDATE ↑ Y.SIMIZU 2005/10/7 仕様のGDﾗｲﾝ数,DVD2上下限を渡すように変更
            bJudg = False
            WfCrGdJudg = False
            Exit Function
        End If
        
        If Not bDenData Then typ_j015.MSRSDEN = -1      '05/10/25 ooba
        If Not bLdlData Then typ_j015.MSRSLDL = -1      '05/10/25 ooba
        If Not bDvd2Data Then typ_j015.MSRSDVD2 = -1    '05/10/25 ooba
        
        'L/DL桁数ﾁｪｯｸ追加　05/10/26 ooba START ==================================>
        If Len(CStr(typ_j015.MSRSLDL)) > 3 Then
            bJudg = False
            WfCrGdJudg = False
            Exit Function
        End If
        'L/DL桁数ﾁｪｯｸ追加　05/10/26 ooba END ====================================>
        SYORIKBN = "1"
    End If
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
'' 2010/04/01 解析⇒SXL検査票にパターン判定を付ける為にパターン結果を残したかった
''            その為にこの場所に下記のロジックを追加したのだと思われる。
''            MSRSDEN/MSRSLDL/MSRSDVD2が保存データがあるかのチェックをしたのだが
''　　　　　　測定してない場合は、いつまでたってもこの処理を通過することになる。
''　　　　　　また、その初期だと思ってこの処理を入れているのだがWFGD2の構造体を
''　　　　　　作ったのが意味が不明である。
''　　　　　　パターン登録用とするのかそれともなんの目的なのかが不明確である。
''　　　　　　また、WFGdJudgがGDパターンの分岐により使うかの判断をすること自体が
''　　　　　　許されるものではない。
''
''       If WFGD.GDPTK = "1" Or WFGD.GDPTK = "2" Then
''          WFGD2 = WFGD
''
''           '再測定結果反映
'            WFGD2.Den = typ_j015.MSRSDEN
'            WFGD2.Dvd2 = typ_j015.MSRSDVD2
'            WFGD2.Ldl = typ_j015.MSRSLDL
''           WFGD2.LdlMin = typ_j015.MSZEROMN
''           WFGD2.LdlMax = typ_j015.MSZEROMX
''
''           'GD判定
''           If WfGdJudg(WFGD2, typ_j015.HSFLG, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
'                bJudg = False
'                WfCrGdJudg = False
'                Exit Function
''            End If
''            ''L/DLの判定結果をGD実績(WF)に反映
'            If WFGD2.JudgLdlPtn = True Then
''                typ_j015.PTNJUDGRES = "1"
''            Else
''                typ_j015.PTNJUDGRES = "9"
''            End If
''        Else
''            typ_j015.PTNJUDGRES = " "
''        End If
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
''
''        'TBCMJ015-UPDATE用GD実績ﾃﾞｰﾀｾｯﾄ
''        bUpdFlg = True
''        If iCntJ015upd > 0 Then
''            For iCnt = 1 To iCntJ015upd
''                '既にﾃﾞｰﾀが存在する場合
''                If typ_j015.SMPLNO = typ_J015_WFGDUpd(iCnt).SMPLNO Then
''                    bUpdFlg = False
''                   Exit For
''               End If
''           Next
''        End If
''        If bUpdFlg Then
''            iCntJ015upd = iCntJ015upd + 1
''            ReDim Preserve typ_J015_WFGDUpd(iCntJ015upd)
''            typ_J015_WFGDUpd(iCntJ015upd) = typ_j015
''        End If
'' 2010/3/31 GDのパターン判定とGD判定について不一致が発生
'' この場所でパターンをクリアすることが意味がわからない。
'' 開発当時はなかったとの事なので取り合えずコメントにします。
'' ロジック的に問題があれば濱まで連絡をください。
''     WFGD.GDPTK = " "
''---------------------------------------------------------
    
''    End If
    
    WFGD.Den = typ_j015.MSRSDEN                  'Den計算値
    WFGD.Ldl = typ_j015.MSRSLDL                  'L/DL計算値
    WFGD.Dvd2 = typ_j015.MSRSDVD2                'DVD2計算値
    WFGD.LdlMin = typ_j015.MSZEROMN
    WFGD.LdlMax = typ_j015.MSZEROMX              'GD判定


'    If WfGdJudg(WFGD, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
    '保証ﾌﾗｸﾞ追加　06/12/22 ooba
    If WfGdJudg(WFGD, typ_j015.HSFLG, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        WfCrGdJudg = False
        
        If SYORIKBN = "1" Then
            ''L/DLの判定結果をGD実績(WF)に反映
                   typ_j015.PTNJUDGRES = " "
              If WFGD.JudgLdlPtn = True Then
                   typ_j015.PTNJUDGRES = "1"
              Else
                   typ_j015.PTNJUDGRES = "9"
              End If

'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End

            ''TBCMJ015-UPDATE用GD実績ﾃﾞｰﾀｾｯﾄ
                  bUpdFlg = True
              If iCntJ015upd > 0 Then
                For iCnt = 1 To iCntJ015upd
                  '既にﾃﾞｰﾀが存在する場合
                  If typ_j015.SMPLNO = typ_J015_WFGDUpd(iCnt).SMPLNO Then
                      bUpdFlg = False
                      Exit For
                  End If
                Next
              End If
              If bUpdFlg Then
                 iCntJ015upd = iCntJ015upd + 1
                 ReDim Preserve typ_J015_WFGDUpd(iCntJ015upd)
                 typ_J015_WFGDUpd(iCntJ015upd) = typ_j015
              End If
        End If
        Exit Function
    Else
        If SYORIKBN = "1" Then
            ''L/DLの判定結果をGD実績(WF)に反映
                   typ_j015.PTNJUDGRES = " "
              If WFGD.JudgLdlPtn = True Then
                   typ_j015.PTNJUDGRES = "1"
              Else
                   typ_j015.PTNJUDGRES = "9"
              End If

'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End

            ''TBCMJ015-UPDATE用GD実績ﾃﾞｰﾀｾｯﾄ
                  bUpdFlg = True
              If iCntJ015upd > 0 Then
                For iCnt = 1 To iCntJ015upd
                  '既にﾃﾞｰﾀが存在する場合
                  If typ_j015.SMPLNO = typ_J015_WFGDUpd(iCnt).SMPLNO Then
                      bUpdFlg = False
                      Exit For
                  End If
                Next
              End If
              If bUpdFlg Then
                 iCntJ015upd = iCntJ015upd + 1
                 ReDim Preserve typ_J015_WFGDUpd(iCntJ015upd)
                 typ_J015_WFGDUpd(iCntJ015upd) = typ_j015
              End If
        End If
      End If
'    If WFGD.JudgDen <> True Or WFGD.JudgLdl <> True Or WFGD.JudgDvd2 <> True Then
    'AN温度判定結果追加　06/12/22 ooba
    If WFGD.JudgDen <> True Or WFGD.JudgLdl <> True Or WFGD.JudgDvd2 <> True _
                                                            Or WFGD.JudgAntnp <> True Then
        bJudg = False
    End If
    
'--------------- 2008/07/25 INSERT START  By Systech ---------------
    pbGDJudgeTbl(1) = WFGD.JudgDen
    pbGDJudgeTbl(2) = WFGD.JudgDvd2
    pbGDJudgeTbl(3) = WFGD.JudgLdl
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
    
    WfCrGdJudg = True
    
End Function

'概要      :GD計算
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                        :説明
'          :tJ015         ,I  ,typ_TBCMJ015              :GD実績構造体
'          :戻り値        ,O  ,FUNCTION_RETURN           :結果 = FUNCTION_RETURN_SUCCESS : OK
'                                                               FUNCTION_RETURN_FAILURE : NG
'説明      :GD測定結果を計算で求める。
'履歴      :05/01/31 ooba
'*** UPDATE ↓ Y.SIMIZU 2005/10/7 仕様のGDﾗｲﾝ数,DVD2上下限を引数にする
'Public Function Calculate_GD(tJ015 As typ_TBCMJ015) As FUNCTION_RETURN
Public Function Calculate_GD(tJ015 As typ_TBCMJ015, ByVal iNum As Single, ByVal dSpecDvd2Min As Double, ByVal dSpecDvd2Max As Double) As FUNCTION_RETURN
'*** UPDATE ↑ Y.SIMIZU 2005/10/7 仕様のGDﾗｲﾝ数,DVD2上下限を引数にする

    Dim iCntX As Integer            'ｶｳﾝﾀX(5)
    Dim iCntY As Integer            'ｶｳﾝﾀY(15)
    Dim iPoint As Integer           '測定点数
    Dim iNoZero As Integer          'Denの平均値が測定点から見てｾﾞﾛでなくなった点までの個数
    Dim dSum As Double
    Dim dAveDen(15) As Double       '各測定点の平均値(Den)
    Dim dAveLDL(15) As Double       '各測定点の平均値(L/DL)
    
'*** UPDATE ↓ Y.SIMIZU 2005/10/7 引数として,仕様のGDﾗｲﾝ数を取得する
'    Dim iNum As Integer             '測定値数(3or5)
'*** UPDATE ↑ Y.SIMIZU 2005/10/7 引数として,仕様のGDﾗｲﾝ数を取得する
    Dim iDen(5, 15) As Integer      '測定値Den
    Dim iLDL(5, 15) As Integer      '測定値L/DL
    Dim iDVD2(5) As Integer         '測定値DVD2
    Dim bDVD2flg As Boolean         '測定値DVD2存在ﾌﾗｸﾞ(True:有、False:無)
'*** UPDATE ↓ Y.SIMIZU 2005/10/7
    Dim iNum2 As Integer
'*** UPDATE ↑ Y.SIMIZU 2005/10/7
    
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    Dim iZeroCnt        As Integer          ' ZEROカウンタ
    Dim dLDLSum(15)     As Double           ' L/DL合計
    Dim dLDLZero(15)    As Double           ' L/DL連続0
    Dim iLDLZeroCnt     As Integer          '
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End


'*** UPDATE ↓ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数,DVD2ﾌﾗｸﾞは仕様のGDﾗｲﾝ数を使用する
'    '変数にGD測定値、測定値数をｾｯﾄ
'    If CalcGD_DataSet(tJ015, iNum, bDVD2flg, iDen, iLDL, iDVD2) <> FUNCTION_RETURN_SUCCESS Then
    If CalcGD_DataSet(tJ015, iDen, iLDL, iDVD2) <> FUNCTION_RETURN_SUCCESS Then
'*** UPDATE ↑ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数,DVD2ﾌﾗｸﾞは仕様のGDﾗｲﾝ数を使用する
        Calculate_GD = FUNCTION_RETURN_FAILURE
        Exit Function
    End If
    
    '初期化
'*** UPDATE ↓ Y.SIMIZU 2005/10/7 ﾗｲﾝ数分測定値があるかを調べてDVD2の計算方法を変える
    '初期化
    bDVD2flg = True

    '3ﾗｲﾝの場合
    If iNum = 3 Then
        iNum2 = 3
    '4.5ﾗｲﾝ又は5ﾗｲﾝの場合
    Else
        iNum2 = 5
    End If
    
    For iCntX = 1 To iNum2
        '測定値が足りない場合は計算によってDVD2をだす
        If iDVD2(iCntX) = -1 Then
            bDVD2flg = False
        End If
    Next
'*** UPDATE ↑ Y.SIMIZU 2005/10/7 ﾗｲﾝ数分測定値があるかを調べてDVD2の計算方法を変える
    
'--DVD2値の計算

    '測定値DVD2が存在する場合
    If bDVD2flg = True Then
    '*** UPDATE ↓ Y.SIMIZU 2005/10/7 4.5ﾗｲﾝ対応
'        'DVD2集計
'        For iCntX = 1 To iNum
'            If iDVD2(iCntX) = -1 Then
'                GoTo MEAS_TEN1
'            End If
'            dSum = dSum + iDVD2(iCntX)
'        Next

        '3ﾗｲﾝの場合
        If iNum = 3 Then
            iNum2 = 3
        '4.5ﾗｲﾝの場合
        Else
            iNum2 = 5
        End If
        
        'DVD2集計
        For iCntX = 1 To iNum2
            If iDVD2(iCntX) = -1 Then
                GoTo MEAS_TEN1
            End If
            dSum = dSum + iDVD2(iCntX)
        Next
    '*** UPDATE ↑ Y.SIMIZU 2005/10/7 4.5ﾗｲﾝ対応
MEAS_TEN1:
        'DVD2計算(平均)
        '小数点第2位で四捨五入、第1位で切り捨て
        tJ015.MSRSDVD2 = Int(Round(dSum / (iCntX - 1), 1))
    End If
    
'*** UPDATE ↓ Y.SIMIZU 2005/10/7 4.5ﾗｲﾝ対応
'    '測定値DenよりDVD2を求める
'    '各測定点の平均を求める
'    For iCntY = 1 To 15
'        dSum = 0
'        For iCntX = 1 To iNum
'            If iDen(iCntX, iCntY) = -1 Then
'                GoTo MEAS_TEN2
'            End If
'            dSum = dSum + iDen(iCntX, iCntY)
'        Next
'        dAveDen(iCntY) = dSum / (iCntX - 1)
'    Next

    '測定値DenよりDVD2を求める
    '各測定点の平均を求める
    For iCntY = 1 To 15
        dSum = 0
        
        '測定点7まで
        If iCntY <= 7 Then
            '3ﾗｲﾝの場合
            If iNum = 3 Then
                iNum2 = 3
            '4.5ﾗｲﾝ又は5ﾗｲﾝの場合
            Else
                iNum2 = 5
            End If
        '測定点8から
        Else
            '3ﾗｲﾝの場合
            If iNum = 3 Then
                iNum2 = 3
            '4.5ﾗｲﾝの場合
            ElseIf iNum = 4.5 Then
                iNum2 = 4
            '5ﾗｲﾝの場合
            Else
                iNum2 = 5
            End If
        End If
        
        For iCntX = 1 To iNum2
            If iDen(iCntX, iCntY) = -1 Then
                GoTo MEAS_TEN2
            End If
            dSum = dSum + iDen(iCntX, iCntY)
        Next
        dAveDen(iCntY) = dSum / (iCntX - 1)
    Next
'*** UPDATE ↑ Y.SIMIZU 2005/10/7 4.5ﾗｲﾝ対応
MEAS_TEN2:
    iPoint = iCntY - 1

    '測定点から見て0でなくなった点までの個数を取得(DVD2範囲を取得)
    For iCntY = iPoint To 1 Step -1
        If dAveDen(iCntY) <> 0 Then
            Exit For
        End If
    Next
    iNoZero = iCntY
    
    '測定値DVD2が存在しない場合
    If bDVD2flg = False Then
        'DVD2計算
        tJ015.MSRSDVD2 = Round(iNoZero * 2 * 10, 0)
    End If

'--Den値の計算
    
    '∑AVEを求める
    dSum = 0
    For iCntY = 1 To iPoint
        dSum = dSum + dAveDen(iCntY)
    Next
    
    'Den計算
    If tJ015.MSRSDVD2 = 0 Then
        tJ015.MSRSDEN = 0
    Else
        tJ015.MSRSDEN = RoundUp((dSum * 10) / (tJ015.MSRSDVD2 / 20), 0)
    End If

'--L/DL値の計算

    If iNoZero = iPoint Then
        tJ015.MSRSLDL = 0
    Else
    
        'L/DL各測定点の平均を求める
        For iCntY = iNoZero + 1 To iPoint
        '*** UPDATE ↓ Y.SIMIZU 2005/10/7 4.5ﾗｲﾝ対応
'            dSum = 0
'            For iCntX = 0 To iNum
'                dSum = dSum + iLDL(iCntX, iCntY)
'            Next
'            dAveLDL(iCntY) = dSum / (iCntX - 1)

            '測定点7まで
            If iCntY <= 7 Then
                '3ﾗｲﾝの場合
                If iNum = 3 Then
                    iNum2 = 3
                '4.5ﾗｲﾝ又は5ﾗｲﾝの場合
                Else
                    iNum2 = 5
                End If
            '測定点8から
            Else
                '3ﾗｲﾝの場合
                If iNum = 3 Then
                    iNum2 = 3
                '4.5ﾗｲﾝの場合
                ElseIf iNum = 4.5 Then
                    iNum2 = 4
                '5ﾗｲﾝの場合
                Else
                    iNum2 = 5
                End If
            End If

            dSum = 0
            For iCntX = 0 To iNum2
                dSum = dSum + iLDL(iCntX, iCntY)
            Next
            dAveLDL(iCntY) = dSum / (iCntX - 1)
        '*** UPDATE ↑ Y.SIMIZU 2005/10/7 4.5ﾗｲﾝ対応
            
        Next
    
        '∑AVEを求める
        dSum = 0
        For iCntY = iNoZero + 1 To iPoint
            dSum = dSum + dAveLDL(iCntY)
        Next
        
        'L/DL計算
        tJ015.MSRSLDL = RoundUp((dSum / (iPoint - iNoZero)) * 10, 0)
    End If

'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
        tJ015.MSZEROMN = 0     ' 連続0MIN
        tJ015.MSZEROMX = 0     ' 連続0MAX
        Erase dLDLSum
        Erase dLDLZero
        iZeroCnt = 0    ' ZEROカウンタ
        iLDLZeroCnt = 0
    
        '' 測定点別のライン合計を求める
        For iCntY = 1 To 15
            For iCntX = 1 To 3   '一覧横
                dLDLSum(iCntY) = dLDLSum(iCntY) + iLDL(iCntX, iCntY)
            Next iCntX
            
            If dLDLSum(iCntY) = 0# Then
                iZeroCnt = iZeroCnt + 1
            End If
            If (dLDLSum(iCntY) <> 0# Or iCntY = 15) _
               And iZeroCnt > 0 Then
               iLDLZeroCnt = iLDLZeroCnt + 1
               dLDLZero(iLDLZeroCnt) = iZeroCnt
               iZeroCnt = 0
            End If
        Next iCntY
        
        For iCntY = 1 To iLDLZeroCnt
            If dLDLZero(iCntY) > tJ015.MSZEROMX Or iCntY = 1 Then
                tJ015.MSZEROMX = dLDLZero(iCntY)
            End If
            If dLDLZero(iCntY) < tJ015.MSZEROMN Or iCntY = 1 Then
                tJ015.MSZEROMN = dLDLZero(iCntY)
            End If
        Next iCntY
        'Centerが0以外の場合、最小値を0にする
        If dLDLSum(1) <> 0 Then
            tJ015.MSZEROMN = 0
        Else
            tJ015.MSZEROMN = dLDLZero(1)
        End If
        
'        ' GDライン数に関係なく、1～3ラインで判定する
'        iZeroCnt = 0    ' ZEROカウンタ
'        For iCntY = 1 To 15 '一覧縦
'            '入力範囲の場合のみ取得
'            If dLDLSum(iCntY) = 0# Then
'                iZeroCnt = iZeroCnt + 1
'            End If
'            If (dLDLSum(iCntY) <> 0# Or iCntY = 15) _
'               And iZeroCnt > 0 Then
'                If iZeroCnt = 1 Then iZeroCnt = 0   ' 0が1個の場合、連続0とする
'
'                If iZeroCnt > tJ015.MSZEROMX Then
'                    tJ015.MSZEROMX = iZeroCnt
'                End If
'
'                If tJ015.MSZEROMN = -1 Then
'                    tJ015.MSZEROMN = tJ015.MSZEROMX
'                End If
'
'                If iZeroCnt < tJ015.MSZEROMN Then
'                    tJ015.MSZEROMN = iZeroCnt
'                End If
'
'                iZeroCnt = 0
'            End If
'        Next iCntY
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End

End Function

'概要      :GD測定値、測定値数ｾｯﾄ
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                        :説明
'          :tJ015         ,I  ,typ_TBCMJ015              :GD実績構造体
'          :iTnum         ,O  ,Integer                   :測定値数(3or5)
'          :bTflg         ,O  ,Boolean                   :測定値DVD2存在ﾌﾗｸﾞ(True:有、False:無)
'          :iTden()       ,O  ,Integer                   :測定値Den(5,15)
'          :iTldl()       ,O  ,Integer                   :測定値L/DL(5,15)
'          :iTdvd2()      ,O  ,Integer                   :測定値DVD2(5)
'          :戻り値        ,O  ,FUNCTION_RETURN           :結果 = FUNCTION_RETURN_SUCCESS : OK
'                                                               FUNCTION_RETURN_FAILURE : NG
'説明      :
'履歴      :05/01/31 ooba
'*** UPDATE ↓ Y.SIMIZU 2005/10/7 仕様のGDﾗｲﾝ数を使用する
'Private Function CalcGD_DataSet(tGDdata As typ_TBCMJ015, iTnum As Integer, bTflg As Boolean, _
'                                    iTden() As Integer, iTldl() As Integer, iTdvd2() As Integer) _
'                                                                                As FUNCTION_RETURN
'    '初期値として測定値数=3をｾｯﾄ
'    iTnum = 3
Private Function CalcGD_DataSet(tGDdata As typ_TBCMJ015, _
                                    iTden() As Integer, iTldl() As Integer, iTdvd2() As Integer) _
                                                                                As FUNCTION_RETURN
'*** UPDATE ↑ Y.SIMIZU 2005/10/7 仕様のGDﾗｲﾝ数を使用する

'*** UPDATE ↓ Y.SIMIZU 2005/10/7 DVD2を計算するかのﾌﾗｸﾞはGDの仕様から立てるように変更
'    bTflg = False
'*** UPDATE ↑ Y.SIMIZU 2005/10/7 DVD2を計算するかのﾌﾗｸﾞはGDの仕様から立てるように変更

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
    
'*** UPDATE ↓ Y.SIMIZU 2005/10/7 ﾗｲﾝ数分の測定値ﾁｪｯｸはChkGD_Dataで行う
'    '測定値DVD2存在ﾁｪｯｸ
'    For iCnt = 1 To 5
'        If iTdvd2(iCnt) <> -1 Then
'            bTflg = True
'            Exit For
'        End If
'    Next
    
'    '測定値数ﾁｪｯｸ
'    For iCnt = 1 To 15
'        '測定値数=5
'        If iTden(5, iCnt) <> -1 Or iTldl(5, iCnt) <> -1 Then
'            iTnum = 5
'            Exit For
'        End If
'    Next
'*** UPDATE ↑ Y.SIMIZU 2005/10/7 ﾗｲﾝ数分の測定値ﾁｪｯｸはChkGD_Dataで行う
    
    CalcGD_DataSet = FUNCTION_RETURN_SUCCESS
    
End Function

'概要      :仕様のGDﾗｲﾝ数分測定値が存在するかをﾁｪｯｸする
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型                 :説明
'          :tGDdata         ,I  ,typ_TBCMJ015      :GD実績構造体
'          :WFGD            ,O  ,W_GD              :GD仕様構造体
'          :bDenChk         ,O  ,Boolean           :Den実績ﾃﾞｰﾀ存在ﾌﾗｸﾞ　05/10/25 ooba
'          :bLdlChk         ,O  ,Boolean           :L/DL実績ﾃﾞｰﾀ存在ﾌﾗｸﾞ　05/10/25 ooba
'          :bDvd2Chk        ,O  ,Boolean           :DVD2実績ﾃﾞｰﾀ存在ﾌﾗｸﾞ　05/10/25 ooba
'          :戻り値          ,O  ,FUNCTION_RETURN    :結果 = FUNCTION_RETURN_SUCCESS : OK
'                                                               FUNCTION_RETURN_FAILURE : NG
'説明      :
'履歴      :05/10/07 Y.SIMIZU
Private Function ChkGD_Data(tGDdata As typ_TBCMJ015, WFGD As W_GD, _
                            bDenChk As Boolean, bLdlChk As Boolean, bDvd2Chk As Boolean) _
                            As FUNCTION_RETURN
    Dim iCnt        As Integer
    Dim iPoint      As Integer
    Dim iLine       As Integer
    Dim iTden(5, 15)     As Integer
    Dim iTldl(5, 15)     As Integer
    Dim iTdvd2(5)    As Integer
    
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
    
    
    'GD実績ﾃﾞｰﾀ存在ﾁｪｯｸ　05/10/25 ooba START ===================================>
    bDenChk = False
    bLdlChk = False
    bDvd2Chk = False
    
    For iCnt = 1 To 5
        If iTdvd2(iCnt) <> -1 Then bDvd2Chk = True
        For iPoint = 1 To 15
            If iTden(iCnt, iPoint) <> -1 Then bDenChk = True
            If iTldl(iCnt, iPoint) <> -1 Then bLdlChk = True
        Next iPoint
    Next iCnt
    If bDenChk Then bDvd2Chk = True
    'GD実績ﾃﾞｰﾀ存在ﾁｪｯｸ　05/10/25 ooba END =====================================>
    
    
    'Denの仕様が検査有り,保証有りの場合
    If WFGD.JudgFlagDen = "1" And WFGD.GuaranteeDen.cJudg = JudgCodeC01 Then
        'Denの測定値がﾗｲﾝ数分あるかをﾁｪｯｸ
        For iPoint = 1 To 15
            '測定点7まで
            If iPoint <= 7 Then
                '仕様が3ﾗｲﾝの場合
                If WFGD.SpecGdLine = 3 Then
                    iLine = 3
                '仕様が4.5ﾗｲﾝ又は5ﾗｲﾝの場合
                Else
                    iLine = 5
                End If
            '測定点8から
            Else
                '仕様が3ﾗｲﾝの場合
                If WFGD.SpecGdLine = 3 Then
                    iLine = 3
                '仕様が4.5ﾗｲﾝの場合
                ElseIf WFGD.SpecGdLine = 4.5 Then
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
    If WFGD.JudgFlagLdl = "1" And WFGD.GuaranteeLdl.cJudg = JudgCodeC01 Then
        'L/DLの測定値がﾗｲﾝ数分あるかをﾁｪｯｸ
        For iPoint = 1 To 15
            '測定点7まで
            If iPoint <= 7 Then
                '仕様が3ﾗｲﾝの場合
                If WFGD.SpecGdLine = 3 Then
                    iLine = 3
                '仕様が4.5ﾗｲﾝ又は5ﾗｲﾝの場合
                Else
                    iLine = 5
                End If
            '測定点8から
            Else
                '仕様が3ﾗｲﾝの場合
                If WFGD.SpecGdLine = 3 Then
                    iLine = 3
                '仕様が4.5ﾗｲﾝの場合
                ElseIf WFGD.SpecGdLine = 4.5 Then
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

'概要      :コード情報取得
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :code          ,   ,Variant      ,
'          :CodeData      ,   ,typ_CodeMaster ,
'          :戻り値        ,O  ,String       ,
'説明      :コード情報リストから該当コードの情報を取得する
'履歴      :
Private Function Search_CrCode(strCode As String, typ_CodeData() As typ_TBCMB005) As String
    Dim i As Integer
    
    'リストから該当コードの情報１を検索
    For i = 1 To UBound(typ_CodeData)
        If strCode = Trim(typ_CodeData(i).CODE) Then
            Search_CrCode = typ_CodeData(i).INFO1
            Exit Function
        End If
    Next
    Search_CrCode = ""
End Function

Public Function NtoS(strWk As String) As String
    If Mid(strWk, 1, 1) = Chr(0) Then
        NtoS = " "
        Exit Function
    End If
    NtoS = strWk
End Function

Public Function NtoZ2(strWk As String) As Double
    If Trim(strWk) = "" Then
        NtoZ2 = -1
        Exit Function
    End If
    NtoZ2 = CDbl(strWk)
End Function

Public Sub BMDDataSet(BmdNo As Integer, UpDo As Integer, INGOTPOS As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String)

    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4               '検査指示
    Dim typ_y013z       As typ_TBCMY013
    Dim bJudg           As Boolean
    Dim vTemp           As Variant
    Dim SijiUmu         As String
    Dim WFBmSokuP       As String                   ' 測定位置＿点
    Dim WFBMD           As Integer                  '2001/12/19 S.Sano
    Dim sSxlPos         As String                   'SXL位置(TOP/BOT)　04/04/12 ooba

    '検査指示設定
    IND = IIf(UpDo = SxlTop, "123", "123")
    
'Chg Start 2011/03/09 SMPK Miyata
'    sSxlPos = IIf(UpDo = SxlTop, "TOP", "BOT")      '04/04/12 ooba
    Select Case UpDo
        Case SxlTop:    sSxlPos = "TOP"
        Case SxlTail:   sSxlPos = "BOT"
        Case Else:      sSxlPos = "MID"
    End Select
'Chg End   2011/03/09 SMPK Miyata

    With typ_CType

        Select Case BmdNo
        Case 1
            WFBMD = WFBMD1 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.B1
            '保証方法ﾁｪｯｸ追加　04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.B1 And CheckKHN(.typ_si.HWFBM1KN, 7, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.B1 And CheckKHN(.typ_si.HWFBM1KN, 7, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '保証方法=保証 かつ 中間抜試（保証）の場合、仕様有とする
                JudgSpecCode = (JudgSW.B1 And .typ_si.MSMPFLGWFBM = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.B1 And .typ_si.MSMPFLGWFBM = "1")
'Chg End　 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "B1"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDB1CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESB1CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFBMD1)
            WFBmSokuP = .typ_si.HWFBM1ST
        Case 2
            WFBMD = WFBMD2 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.B2
            '保証方法ﾁｪｯｸ追加　04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.B2 And CheckKHN(.typ_si.HWFBM2KN, 8, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.B2 And CheckKHN(.typ_si.HWFBM2KN, 8, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '保証方法=保証 かつ 中間抜試（保証）の場合、仕様有とする
                JudgSpecCode = (JudgSW.B2 And .typ_si.MSMPFLGWFBM = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.B2 And .typ_si.MSMPFLGWFBM = "1")
'Chg End　 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "B2"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDB2CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESB2CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFBMD2)
            WFBmSokuP = .typ_si.HWFBM2ST
        Case 3
            WFBMD = WFBMD3 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.B3
            '保証方法ﾁｪｯｸ追加　04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.B3 And CheckKHN(.typ_si.HWFBM3KN, 9, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.B3 And CheckKHN(.typ_si.HWFBM3KN, 9, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '保証方法=保証 かつ 中間抜試（保証）の場合、仕様有とする
                JudgSpecCode = (JudgSW.B3 And .typ_si.MSMPFLGWFBM = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.B3 And .typ_si.MSMPFLGWFBM = "1")
'Chg End　 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "B3"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDB3CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESB3CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFBMD3)
            WFBmSokuP = .typ_si.HWFBM3ST
        End Select
            typ_y013z = .typ_y013(UpDo, WFBMD) '2001/12/19 S.Sano
    
    
        '' WF検査指示（B1)*****************************************************************
        If JudgSpecCode Then
            '画面表示内容設定
            .typ_rslt(UpDo, DispLineCount).pos = -1                                             ' 結晶内開始位置
            .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())              ' 内容
            .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                                     ' 情報１
            .typ_rslt(UpDo, DispLineCount).INFO2 = "検査無"                                     ' 情報２
            .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                           ' 情報３
            .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                           ' 情報４
        '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '5番目の情報：AN温度を追加
            typ_rslt_ex(UpDo, DispLineCount).pos = .typ_rslt(UpDo, DispLineCount).pos
            typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
            typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
            typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
            typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
            typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
            typ_rslt_ex(UpDo, DispLineCount).INFO5 = ""                                         ' 情報5
        '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
        '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
            typ_rslt_ex(UpDo, DispLineCount).INFO6 = ""                                         ' 情報6
            typ_rslt_ex(UpDo, DispLineCount).INFO7 = ""                                         ' 情報7
            typ_rslt_ex(UpDo, DispLineCount).INFO8 = ""                                         ' 情報8
        '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------

            .typ_rslt(UpDo, DispLineCount).SMPLID = -1                                          ' サンプルＮｏ
            .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                          ' 判定結果
            .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                                   ' 品番(12桁)
            bJudg = False
            If shiji <> 0 Then
                '画面表示内容設定
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' 結晶内開始位置
                .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                                 ' 情報２
                .typ_rslt(UpDo, DispLineCount).INFO3 = "実績無"                                 ' 情報３
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y013z.SAMPLEID                      ' サンプルＮｏ
                    
                'BMD1判定
'                If SijiUmu = "1" Then
                If (SijiUmu = "1") And (Trim(typ_y013z.SAMPLEID) <> "0") Then           '2003/12/19 SystemBrain
                    'BMD1判定失敗
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                            ' 情報２
                    'BMD1判定
                    If WfCrBmdJudg(.typ_si, typ_y013z, bJudg, BmdNo) Then
'                        '画面表示内容設定
'                        vTemp = CVar(typ_y013z.MESDATA5)
'                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' 情報1
'                        vTemp = CVar(typ_y013z.MESDATA6)
'                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.00")   ' 情報2
'                        vTemp = CVar(typ_y013z.MESDATA8)
'                        .typ_rslt(UpDo, DispLineCount).INFO3 = ""                               ' 情報３
'                        .typ_rslt(UpDo, DispLineCount).INFO4 = ""                               ' 情報４

                        '画面表示内容設定　　2003/05/20 ooba
                        vTemp = CVar(typ_y013z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' 情報1
                        vTemp = CVar(typ_y013z.MESDATA7)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.00")   ' 情報2
                        vTemp = CVar(typ_y013z.MESDATA8)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' 情報３
                        vTemp = CVar(typ_y013z.MESDATA9)
                        .typ_rslt(UpDo, DispLineCount).INFO4 = DBData2DispData(vTemp, "0.0")    ' 情報４
                        JiltusekiUmu(UpDo, WFBMD) = True '2001/12/19 S.Sano
                    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                        '5番目の情報：AN温度を追加
                        vTemp = CVar(typ_y013z.DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(UpDo, DispLineCount).pos = .typ_rslt(UpDo, DispLineCount).pos
                        typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
                        typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
                        typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
                        typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
                        typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
                        typ_rslt_ex(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    End If
                ElseIf SijiUmu = "2" Then
                    bJudg = False
                    '画面表示内容設定
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ異常"                          ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                End If
                
            End If
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' 判定結果
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                      ' 判定結果
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If UpDo = SxlTop Or UpDo = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                Select Case BmdNo
                Case 1
                    gsTbcmy028ErrCode = "00136"
                Case 2
                    gsTbcmy028ErrCode = "00137"
                Case 3
                    gsTbcmy028ErrCode = "00138"
                End Select
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            DispLineCount = DispLineCount + 1
        Else
            If shiji <> 0 Then
                '画面表示内容設定
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' 結晶内開始位置
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())          ' 内容
                .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様無"                                 ' 情報１
                .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                                 ' 情報２
                .typ_rslt(UpDo, DispLineCount).INFO3 = "実績無"                                 ' 情報３
                .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                       ' 情報４
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                typ_rslt_ex(UpDo, DispLineCount).pos = .typ_rslt(UpDo, DispLineCount).pos
                typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
                typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
                typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
                typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
                typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
                typ_rslt_ex(UpDo, DispLineCount).INFO5 = ""                                     ' 情報5
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
            '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
                typ_rslt_ex(UpDo, DispLineCount).INFO6 = ""                                     ' 情報6
                typ_rslt_ex(UpDo, DispLineCount).INFO7 = ""                                     ' 情報7
                typ_rslt_ex(UpDo, DispLineCount).INFO8 = ""                                     ' 情報8
            '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y013z.SAMPLEID                      ' サンプルＮｏ
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' 判定結果
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                               ' 品番(12桁)
'                If SijiUmu = "1" Then
                If (SijiUmu = "1") And (Trim(typ_y013z.SAMPLEID) <> "0") Then           '2003/12/19 SystemBrain
                    'BMD1判定失敗
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                            ' 情報２
                    'BMD1判定
                    If WfCrBmdJudg(.typ_si, typ_y013z, bJudg, BmdNo) Then
'                        '画面表示内容設定
'                        vTemp = CVar(typ_y013z.MESDATA5)
'                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' 情報1
'                        vTemp = CVar(typ_y013z.MESDATA6)
'                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.00")   ' 情報2
'                        vTemp = CVar(typ_y013z.MESDATA8)
'                        .typ_rslt(UpDo, DispLineCount).INFO3 = ""                               ' 情報３
'                        .typ_rslt(UpDo, DispLineCount).INFO4 = ""                               ' 情報４

                        '画面表示内容設定　　2003/05/20 ooba
                        vTemp = CVar(typ_y013z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' 情報1
                        vTemp = CVar(typ_y013z.MESDATA7)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.00")   ' 情報2
                        vTemp = CVar(typ_y013z.MESDATA8)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' 情報３
                        vTemp = CVar(typ_y013z.MESDATA9)
                        .typ_rslt(UpDo, DispLineCount).INFO4 = DBData2DispData(vTemp, "0.0")    ' 情報４
                        JiltusekiUmu(UpDo, WFBMD) = True '2001/12/19 S.Sano
                    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                        vTemp = CVar(typ_y013z.DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(UpDo, DispLineCount).pos = .typ_rslt(UpDo, DispLineCount).pos
                        typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
                        typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
                        typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
                        typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
                        typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
                        typ_rslt_ex(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    End If
                ElseIf SijiUmu = "2" Then
                    bJudg = False
                    '画面表示内容設定
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ異常"                          ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                End If
                
                'Add Start 2011/11/28 Y.Hitomi 中間抜試の場合は、参考表示する
                If sSxlPos = "MID" And bJudg = False Then
                    If BmdNo = 1 And JudgSW.B1 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "参考"                                      ' 判定結果
                        MidlJudg = False
                    ElseIf BmdNo = 2 And JudgSW.B2 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "参考"                                      ' 判定結果
                        MidlJudg = False
                    ElseIf BmdNo = 3 And JudgSW.B3 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "参考"                                      ' 判定結果
                        MidlJudg = False
                    End If
                End If
                'Add End  2011/11/28 Y.Hitomi
    
                DispLineCount = DispLineCount + 1
            End If
        End If
    
        Select Case BmdNo
        Case 1
            .typ_y013(UpDo, WFBMD1) = typ_y013z
        Case 2
            .typ_y013(UpDo, WFBMD2) = typ_y013z
        Case 3
            .typ_y013(UpDo, WFBMD3) = typ_y013z
        End Select
    
    End With
    
End Sub

Public Sub OSFDataSet(OsfNo As Integer, UpDo As Integer, INGOTPOS As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String)

    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4                   '検査指示
    Dim typ_y013z       As typ_TBCMY013
    Dim AveMax(1)       As String                       '平均/最大判定値　2003/05/20 ooba
    Dim bJudg           As Boolean
    Dim vTemp           As Variant
    Dim SijiUmu         As String
    Dim WFBmSokuP       As String                       ' 測定位置＿点
    Dim WFBmSokuHou     As String                       ' 品WFOSF1測定位置_方
    Dim WFBmSokuRyou    As String                       ' 品WFOSF1測定位置_領
    Dim WFOSF           As Integer                      '2001/12/19 S.Sano
    Dim sSxlPos         As String                       'SXL位置(TOP/BOT)　04/04/12 ooba
    
    '検査指示設定
    IND = IIf(UpDo = SxlTop, "123", "123")
        
'Chg Start 2011/03/09 SMPK Miyata
'    sSxlPos = IIf(UpDo = SxlTop, "TOP", "BOT")      '04/04/12 ooba
    Select Case UpDo
        Case SxlTop:    sSxlPos = "TOP"
        Case SxlTail:   sSxlPos = "BOT"
        Case Else:      sSxlPos = "MID"
    End Select
'Chg End   2011/03/09 SMPK Miyata
    With typ_CType
        Select Case OsfNo
        Case 1
            WFOSF = WFOSF1 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.L1
            '保証方法ﾁｪｯｸ追加　04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.L1 And CheckKHN(.typ_si.HWFOF1KN, 3, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.L1 And CheckKHN(.typ_si.HWFOF1KN, 3, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '保証方法=保証 かつ 中間抜試（保証）の場合、仕様有とする
                JudgSpecCode = (JudgSW.L1 And .typ_si.MSMPFLGWFOF = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.L1 And .typ_si.MSMPFLGWFOF = "1")
'Chg End　 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata

            SCC = "L1"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDL1CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESL1CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFOSF1)
            WFBmSokuHou = .typ_si.HWFOF1SH
            WFBmSokuP = .typ_si.HWFOF1ST
            WFBmSokuRyou = .typ_si.HWFOF1SR
        Case 2
            WFOSF = WFOSF2 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.L2
            '保証方法ﾁｪｯｸ追加　04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.L2 And CheckKHN(.typ_si.HWFOF2KN, 4, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.L2 And CheckKHN(.typ_si.HWFOF2KN, 4, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '保証方法=保証 かつ 中間抜試（保証）の場合、仕様有とする
                JudgSpecCode = (JudgSW.L2 And .typ_si.MSMPFLGWFOF = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.L2 And .typ_si.MSMPFLGWFOF = "1")
'Chg End　 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "L2"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDL2CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESL2CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFOSF2)
            WFBmSokuHou = .typ_si.HWFOF2SH
            WFBmSokuP = .typ_si.HWFOF2ST
            WFBmSokuRyou = .typ_si.HWFOF2SR
        Case 3
            WFOSF = WFOSF3 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.L3
            '保証方法ﾁｪｯｸ追加　04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.L3 And CheckKHN(.typ_si.HWFOF3KN, 5, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.L3 And CheckKHN(.typ_si.HWFOF3KN, 5, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '保証方法=保証 かつ 中間抜試（保証）の場合、仕様有とする
                JudgSpecCode = (JudgSW.L3 And .typ_si.MSMPFLGWFOF = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.L3 And .typ_si.MSMPFLGWFOF = "1")
'Chg End　 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "L3"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDL3CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESL3CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFOSF3)
            WFBmSokuHou = .typ_si.HWFOF3SH
            WFBmSokuP = .typ_si.HWFOF3ST
            WFBmSokuRyou = .typ_si.HWFOF3SR
        Case 4
            WFOSF = WFOSF4 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.L4
            '保証方法ﾁｪｯｸ追加　04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.L4 And CheckKHN(.typ_si.HWFOF4KN, 6, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.L4 And CheckKHN(.typ_si.HWFOF4KN, 6, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '保証方法=保証 かつ 中間抜試（保証）の場合、仕様有とする
                JudgSpecCode = (JudgSW.L4 And .typ_si.MSMPFLGWFOF = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.L4 And .typ_si.MSMPFLGWFOF = "1")
'Chg End　 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "L4"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDL4CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESL4CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFOSF4)
            WFBmSokuHou = .typ_si.HWFOF4SH
            WFBmSokuP = .typ_si.HWFOF4ST
            WFBmSokuRyou = .typ_si.HWFOF4SR
        End Select
        typ_y013z = .typ_y013(UpDo, WFOSF)
        
        
        '' WF検査指示（L1)*****************************************************************
        If JudgSpecCode Then
            '画面表示内容設定
            .typ_rslt(UpDo, DispLineCount).pos = -1                                             ' 結晶内開始位置
            .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())              ' 内容
            .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                                     ' 情報１
            .typ_rslt(UpDo, DispLineCount).INFO2 = "検査無"                                     ' 情報２
            .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                           ' 情報３
            .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                           ' 情報４
        '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            typ_rslt_ex(UpDo, DispLineCount).pos = .typ_rslt(UpDo, DispLineCount).pos
            typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
            typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
            typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
            typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
            typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
            typ_rslt_ex(UpDo, DispLineCount).INFO5 = ""                                         ' 情報5
        '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
        '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
            typ_rslt_ex(UpDo, DispLineCount).INFO6 = ""                                         ' 情報6
            typ_rslt_ex(UpDo, DispLineCount).INFO7 = ""                                         ' 情報7
            typ_rslt_ex(UpDo, DispLineCount).INFO8 = ""                                         ' 情報8
        '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
            .typ_rslt(UpDo, DispLineCount).SMPLID = -1                                          ' サンプルＮｏ
            .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                          ' 判定結果
            .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                                   ' 品番(12桁)
            bJudg = False
            If shiji <> 0 Then
                '画面表示内容設定
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' 結晶内開始位置
                .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                                 ' 情報２
                .typ_rslt(UpDo, DispLineCount).INFO3 = "実績無"                                 ' 情報３
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y013z.SAMPLEID                      ' サンプルＮｏ
                'OSF判定取得
'                If SijiUmu = "1" Then
                If (SijiUmu = "1") And (Trim(typ_y013z.SAMPLEID) <> "0") Then           '2003/12/19 SystemBrain
                    'OSF判定失敗
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                            ' 情報２
                    'OSF1判定取得
                    If WfCrOsfJudg(.typ_si, typ_y013z, bJudg, OsfNo, AveMax()) Then             ' AveMax追加　2003/05/20 ooba
'                        '画面表示内容設定
'                        vTemp = CVar(typ_y013z.MESDATA7)
'                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' 情報１
'                        vTemp = CVar(typ_y013z.MESDATA8)
'                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")    ' 情報２
'                        vTemp = CVar(typ_y013z.MESDATA6)
'                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' 情報３
'                        .typ_rslt(UpDo, DispLineCount).INFO4 = ""                               ' 情報４
                        
                        '画面表示内容設定　　2003/05/21 ooba
                        vTemp = CVar(AveMax(0))
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' 情報１
                        vTemp = CVar(AveMax(1))
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")    ' 情報２
                        vTemp = CVar(typ_y013z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' 情報３
                        vTemp = CVar(IIf(Trim(typ_y013z.MESDATA9) = "", "-", Trim(typ_y013z.MESDATA9)) _
                                        & "  " & IIf(Trim(typ_y013z.MESDATA12) = "", "-", Trim(typ_y013z.MESDATA12)) _
                                        & "  " & IIf(Trim(typ_y013z.MESDATA15) = "", "-", Trim(typ_y013z.MESDATA15)))
                        .typ_rslt(UpDo, DispLineCount).INFO4 = vTemp                            ' 情報４
                        
                        JiltusekiUmu(UpDo, WFOSF) = True '2001/12/19 S.Sano
                    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                        vTemp = CVar(typ_y013z.DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(UpDo, DispLineCount).pos = .typ_rslt(UpDo, DispLineCount).pos
                        typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
                        typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
                        typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
                        typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
                        typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
                        typ_rslt_ex(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    End If
                ElseIf SijiUmu = "2" Then
                    '画面表示内容設定
                    bJudg = False
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ異常"                          ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                End If
            End If
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' 判定結果
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                      ' 判定結果
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If UpDo = SxlTop Or UpDo = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                Select Case OsfNo
                Case 1
                    gsTbcmy028ErrCode = "00132"
                Case 2
                    gsTbcmy028ErrCode = "00133"
                Case 3
                    gsTbcmy028ErrCode = "00134"
                Case 4
                    gsTbcmy028ErrCode = "00135"
                End Select
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            DispLineCount = DispLineCount + 1
        Else
            If shiji <> 0 Then
                '画面表示内容設定
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' 結晶内開始位置
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())          ' 内容
                .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様無"                                 ' 情報１
                .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                                 ' 情報２
                .typ_rslt(UpDo, DispLineCount).INFO3 = "実績無"                                 ' 情報３
                .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                       ' 情報４
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                typ_rslt_ex(UpDo, DispLineCount).pos = .typ_rslt(UpDo, DispLineCount).pos
                typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
                typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
                typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
                typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
                typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
                typ_rslt_ex(UpDo, DispLineCount).INFO5 = ""                                     ' 情報5
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
            '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
                typ_rslt_ex(UpDo, DispLineCount).INFO6 = ""                                     ' 情報6
                typ_rslt_ex(UpDo, DispLineCount).INFO7 = ""                                     ' 情報7
                typ_rslt_ex(UpDo, DispLineCount).INFO8 = ""                                     ' 情報8
            '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y013z.SAMPLEID                      ' サンプルＮｏ
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' 判定結果
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                               ' 品番(12桁)
                'OSF判定取得
'                If SijiUmu = "1" Then
                If (SijiUmu = "1") And (Trim(typ_y013z.SAMPLEID) <> "0") Then           '2003/12/19 SystemBrain
                    'OSF判定失敗
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                            ' 情報２
                    'OSF判定取得
                    If WfCrOsfJudg(.typ_si, typ_y013z, bJudg, OsfNo, AveMax()) Then             ' AveMax追加　2003/05/20 ooba
'                        '画面表示内容設定
'                        vTemp = CVar(typ_y013z.MESDATA7)
'                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' 情報１
'                        vTemp = CVar(typ_y013z.MESDATA8)
'                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")    ' 情報２
'                        vTemp = CVar(typ_y013z.MESDATA6)
'                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' 情報３
'                        .typ_rslt(UpDo, DispLineCount).INFO4 = ""                               ' 情報４
                        
                        '画面表示内容設定　　2003/05/21 ooba
                        vTemp = CVar(AveMax(0))
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' 情報１
                        vTemp = CVar(AveMax(1))
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")    ' 情報２
                        vTemp = CVar(typ_y013z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' 情報３
                        vTemp = CVar(IIf(Trim(typ_y013z.MESDATA9) = "", "-", Trim(typ_y013z.MESDATA9)) _
                                        & "  " & IIf(Trim(typ_y013z.MESDATA12) = "", "-", Trim(typ_y013z.MESDATA12)) _
                                        & "  " & IIf(Trim(typ_y013z.MESDATA15) = "", "-", Trim(typ_y013z.MESDATA15)))
                        .typ_rslt(UpDo, DispLineCount).INFO4 = vTemp                            ' 情報４
                         JiltusekiUmu(UpDo, WFOSF) = True '2001/12/19 S.Sano
                    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                        vTemp = CVar(typ_y013z.DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(UpDo, DispLineCount).pos = .typ_rslt(UpDo, DispLineCount).pos
                        typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
                        typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
                        typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
                        typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
                        typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
                        typ_rslt_ex(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    End If
                ElseIf SijiUmu = "2" Then
                    '画面表示内容設定
                    bJudg = False
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ異常"                          ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                End If
                'Add Start 2011/11/28 Y.Hitomi 中間抜試の場合は、参考表示する
                If sSxlPos = "MID" And bJudg = False Then
                    If OsfNo = 1 And JudgSW.L1 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "参考"                                      ' 判定結果
                        MidlJudg = False
                    ElseIf OsfNo = 2 And JudgSW.L2 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "参考"                                      ' 判定結果
                        MidlJudg = False
                    ElseIf OsfNo = 3 And JudgSW.L3 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "参考"                                      ' 判定結果
                        MidlJudg = False
                    ElseIf OsfNo = 4 And JudgSW.L4 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "参考"                                      ' 判定結果
                        MidlJudg = False
                    End If
                End If
                'Add End  2011/11/28 Y.Hitomi
                
                DispLineCount = DispLineCount + 1
            End If
        End If
    
        Select Case OsfNo
        Case 1
            .typ_y013(UpDo, WFOSF1) = typ_y013z
            TmpOsfData(0, UpDo, WFOSF1) = AveMax(0)                                                 '　▼2003/05/20 ooba
            TmpOsfData(1, UpDo, WFOSF1) = AveMax(1)
            TmpOsfMBNP(0, UpDo, WFOSF1) = IIf(typ_y013z.MESDATA9 = "-", " ", typ_y013z.MESDATA9)
            TmpOsfMBNP(1, UpDo, WFOSF1) = IIf(typ_y013z.MESDATA12 = "-", " ", typ_y013z.MESDATA12)
            TmpOsfMBNP(2, UpDo, WFOSF1) = IIf(typ_y013z.MESDATA15 = "-", " ", typ_y013z.MESDATA15)  '　▲2003/05/21 ooba
        Case 2
            .typ_y013(UpDo, WFOSF2) = typ_y013z
            TmpOsfData(0, UpDo, WFOSF2) = AveMax(0)                                                 '　▼2003/05/20 ooba
            TmpOsfData(1, UpDo, WFOSF2) = AveMax(1)
            TmpOsfMBNP(0, UpDo, WFOSF2) = IIf(typ_y013z.MESDATA9 = "-", " ", typ_y013z.MESDATA9)
            TmpOsfMBNP(1, UpDo, WFOSF2) = IIf(typ_y013z.MESDATA12 = "-", " ", typ_y013z.MESDATA12)
            TmpOsfMBNP(2, UpDo, WFOSF2) = IIf(typ_y013z.MESDATA15 = "-", " ", typ_y013z.MESDATA15)  '　▲2003/05/21 ooba
        Case 3
            .typ_y013(UpDo, WFOSF3) = typ_y013z
            TmpOsfData(0, UpDo, WFOSF3) = AveMax(0)                                                 '　▼2003/05/20 ooba
            TmpOsfData(1, UpDo, WFOSF3) = AveMax(1)
            TmpOsfMBNP(0, UpDo, WFOSF3) = IIf(typ_y013z.MESDATA9 = "-", " ", typ_y013z.MESDATA9)
            TmpOsfMBNP(1, UpDo, WFOSF3) = IIf(typ_y013z.MESDATA12 = "-", " ", typ_y013z.MESDATA12)
            TmpOsfMBNP(2, UpDo, WFOSF3) = IIf(typ_y013z.MESDATA15 = "-", " ", typ_y013z.MESDATA15)  '　▲2003/05/21 ooba
        Case 4
            .typ_y013(UpDo, WFOSF4) = typ_y013z
            TmpOsfData(0, UpDo, WFOSF4) = AveMax(0)                                                 '　▼2003/05/20 ooba
            TmpOsfData(1, UpDo, WFOSF4) = AveMax(1)
            TmpOsfMBNP(0, UpDo, WFOSF4) = IIf(typ_y013z.MESDATA9 = "-", " ", typ_y013z.MESDATA9)
            TmpOsfMBNP(1, UpDo, WFOSF4) = IIf(typ_y013z.MESDATA12 = "-", " ", typ_y013z.MESDATA12)
            TmpOsfMBNP(2, UpDo, WFOSF4) = IIf(typ_y013z.MESDATA15 = "-", " ", typ_y013z.MESDATA15)  '　▲2003/05/21 ooba
        End Select
    
    End With
End Sub

Public Sub DOIDataSet(DoiNo As Integer, UpDo As Integer, INGOTPOS As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String)

    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4               '検査指示
    Dim typ_y013z       As typ_TBCMY013
    Dim bJudg           As Boolean
    Dim vTemp           As Variant
    Dim SijiUmu         As String
    Dim WFBmSokuP       As String                   ' 測定位置＿点
    Dim WFDOI           As Integer                  '2001/12/19 S.Sano
    Dim sSxlPos         As String                   'SXL位置(TOP/BOT)　04/04/12 ooba
    
    '検査指示設定
'    IND = IIf(UpDo = SxlTop, "12346", "123")
    IND = IIf(UpDo = SxlTop, "123", "123")
        
'Chg Start 2011/03/09 SMPK Miyata
'    sSxlPos = IIf(UpDo = SxlTop, "TOP", "BOT")      '04/04/12 ooba
    Select Case UpDo
        Case SxlTop:    sSxlPos = "TOP"
        Case SxlTail:   sSxlPos = "BOT"
        Case Else:      sSxlPos = "MID"
    End Select
'Chg End   2011/03/09 SMPK Miyata
    
    With typ_CType
        Select Case DoiNo
        Case 1
            WFDOI = WFDOI1 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.Doi1
            '保証方法ﾁｪｯｸ追加　04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.Doi1 And CheckKHN(.typ_si.HWFOS1KN, 10, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.Doi1 And CheckKHN(.typ_si.HWFOS1KN, 10, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '保証方法=保証 かつ 中間抜試（保証）の場合、仕様有とする
                JudgSpecCode = (JudgSW.Doi1 And .typ_si.MSMPFLGWFDOI = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.Doi1 And .typ_si.MSMPFLGWFDOI = "1")
'Chg End　 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "DO1"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDDO1CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESDO1CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFDOI1)
        Case 2
            WFDOI = WFDOI2 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.Doi2
            '保証方法ﾁｪｯｸ追加　04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.Doi2 And CheckKHN(.typ_si.HWFOS2KN, 11, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.Doi2 And CheckKHN(.typ_si.HWFOS2KN, 11, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '保証方法=保証 かつ 中間抜試（保証）の場合、仕様有とする
                JudgSpecCode = (JudgSW.Doi2 And .typ_si.MSMPFLGWFDOI = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.Doi2 And .typ_si.MSMPFLGWFDOI = "1")
'Chg End　 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "DO2"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDDO2CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESDO2CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFDOI2)
        Case 3
            WFDOI = WFDOI3 '2001/12/19 S.Sano
'            JudgSpecCode = JudgSW.Doi3
            '保証方法ﾁｪｯｸ追加　04/04/12 ooba
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.Doi3 And CheckKHN(.typ_si.HWFOS3KN, 12, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.Doi3 And CheckKHN(.typ_si.HWFOS3KN, 12, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '保証方法=保証 かつ 中間抜試（保証）の場合、仕様有とする
                JudgSpecCode = (JudgSW.Doi3 And .typ_si.MSMPFLGWFDOI = "1" And _
                                (.typ_si.MSMPFLG = "1" Or .typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.Doi3 And .typ_si.MSMPFLGWFDOI = "1")
'Chg End　 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "DO3"
            shiji = (InStr(IND, .typ_Param.WFSMP(UpDo).WFINDDO3CW) <> 0)
            SijiUmu = .typ_Param.WFSMP(UpDo).WFRESDO3CW
'2001/12/19 S.Sano            typ_y013z = .typ_y013(UpDo, WFDOI3)
        End Select
        typ_y013z = .typ_y013(UpDo, WFDOI)

        
        '' WF検査指示（DOI)*****************************************************************
        If JudgSpecCode Then
            '画面表示内容設定
            .typ_rslt(UpDo, DispLineCount).pos = -1                                             ' 結晶内開始位置
            .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())              ' 内容
            .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                                     ' 情報１
            .typ_rslt(UpDo, DispLineCount).INFO2 = "検査無"                                     ' 情報２
            .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                           ' 情報３
            .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                           ' 情報４
        '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
            typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
            typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
            typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
            typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
            typ_rslt_ex(UpDo, DispLineCount).INFO5 = ""                                         ' 情報5
        '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
        '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
            typ_rslt_ex(UpDo, DispLineCount).INFO6 = ""                                         ' 情報6
            typ_rslt_ex(UpDo, DispLineCount).INFO7 = ""                                         ' 情報7
            typ_rslt_ex(UpDo, DispLineCount).INFO8 = ""                                         ' 情報8
        '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
            .typ_rslt(UpDo, DispLineCount).SMPLID = -1                                          ' サンプルＮｏ
            .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                          ' 判定結果
            .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                                   ' 品番(12桁)
            bJudg = False
            If shiji <> 0 Then
                '画面表示内容設定
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' 結晶内開始位置
                .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                                 ' 情報２
                .typ_rslt(UpDo, DispLineCount).INFO3 = "実績無"                                 ' 情報３
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y013z.SAMPLEID                      ' サンプルＮｏ
                'DOI判定取得
'                If SijiUmu = "1" Then
                If (SijiUmu = "1") And (Trim(typ_y013z.SAMPLEID) <> "0") Then           '2003/12/19 SystemBrain
                    'DOI判定失敗
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                            ' 情報２
                    'DOI判定取得
                    If WfCrDoiJudg(.typ_si, typ_y013z, bJudg, DoiNo) Then
                        '画面表示内容設定
                        vTemp = CVar(typ_y013z.MESDATA1 - typ_y013z.MESDATA4)
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.0")    ' 情報1
                        vTemp = CVar(typ_y013z.MESDATA2 - typ_y013z.MESDATA5)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")    ' 情報2
                        vTemp = CVar(typ_y013z.MESDATA3 - typ_y013z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.0")    ' 情報3
                        .typ_rslt(UpDo, DispLineCount).INFO4 = ""                               ' 情報４
                        JiltusekiUmu(UpDo, WFDOI) = True '2001/12/19 S.Sano
                    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                        vTemp = CVar(typ_y013z.DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
                        typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
                        typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
                        typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
                        typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
                        typ_rslt_ex(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    End If
                ElseIf SijiUmu = "2" Then
                    '画面表示内容設定
                    bJudg = False
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ異常"                          ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                End If
            End If
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' 判定結果
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                      ' 判定結果
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If UpDo = SxlTop Or UpDo = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                Select Case DoiNo
                Case 1
                    gsTbcmy028ErrCode = "00139"
                Case 2
                    gsTbcmy028ErrCode = "00140"
                Case 3
                    gsTbcmy028ErrCode = "00141"
                End Select
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            DispLineCount = DispLineCount + 1
        Else
            If shiji <> 0 Then
                '画面表示内容設定
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' 結晶内開始位置
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())          ' 内容
                .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様無"                                 ' 情報１
                .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                                 ' 情報２
                .typ_rslt(UpDo, DispLineCount).INFO3 = "実績無"                                 ' 情報３
                .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                       ' 情報４
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
                typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
                typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
                typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
                typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
                typ_rslt_ex(UpDo, DispLineCount).INFO5 = ""                                     ' 情報5
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
            '項目6(PUA値)、7(PUA%値)、8(STD値)追加による変更
                typ_rslt_ex(UpDo, DispLineCount).INFO6 = ""                                     ' 情報6
                typ_rslt_ex(UpDo, DispLineCount).INFO7 = ""                                     ' 情報7
                typ_rslt_ex(UpDo, DispLineCount).INFO8 = ""                                     ' 情報8
            '↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y013z.SAMPLEID                      ' サンプルＮｏ
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' 判定結果
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                               ' 品番(12桁)
                'DOI判定取得
'                If SijiUmu = "1" Then
                If (SijiUmu = "1") And (Trim(typ_y013z.SAMPLEID) <> "0") Then           '2003/12/19 SystemBrain
                    'DOI判定失敗
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                            ' 情報２
                    'DOI判定取得
                    If WfCrDoiJudg(.typ_si, typ_y013z, bJudg, DoiNo) Then
                        '画面表示内容設定
                        vTemp = CVar(typ_y013z.MESDATA1 - typ_y013z.MESDATA4)
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.0")    ' 情報1
                        vTemp = CVar(typ_y013z.MESDATA2 - typ_y013z.MESDATA5)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")    ' 情報2
                        vTemp = CVar(typ_y013z.MESDATA3 - typ_y013z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.0")    ' 情報3
                        .typ_rslt(UpDo, DispLineCount).INFO4 = ""                               ' 情報４
                        JiltusekiUmu(UpDo, WFDOI) = True '2001/12/19 S.Sano
                    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                        vTemp = CVar(typ_y013z.DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        typ_rslt_ex(UpDo, DispLineCount).NAIYO = .typ_rslt(UpDo, DispLineCount).NAIYO
                        typ_rslt_ex(UpDo, DispLineCount).INFO1 = .typ_rslt(UpDo, DispLineCount).INFO1
                        typ_rslt_ex(UpDo, DispLineCount).INFO2 = .typ_rslt(UpDo, DispLineCount).INFO2
                        typ_rslt_ex(UpDo, DispLineCount).INFO3 = .typ_rslt(UpDo, DispLineCount).INFO3
                        typ_rslt_ex(UpDo, DispLineCount).INFO4 = .typ_rslt(UpDo, DispLineCount).INFO4
                        typ_rslt_ex(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")        ' 情報5
                    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                    End If
                ElseIf SijiUmu = "2" Then
                    '画面表示内容設定
                    bJudg = False
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ異常"                          ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                End If
                'Add Start 2011/11/28 Y.Hitomi 中間抜試の場合は、参考表示する
                If sSxlPos = "MID" And bJudg = False Then
                    If DoiNo = 1 And JudgSW.Doi1 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "参考"                                      ' 判定結果
                        MidlJudg = False
                    ElseIf DoiNo = 2 And JudgSW.Doi2 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "参考"                                      ' 判定結果
                        MidlJudg = False
                    ElseIf DoiNo = 3 And JudgSW.Doi3 Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "参考"                                      ' 判定結果
                        MidlJudg = False
                    End If
                End If
                'Add End  2011/11/28 Y.Hitomi
                DispLineCount = DispLineCount + 1
            End If
        End If
    End With
End Sub
'''''============================================================================================================================
'''''
''''''概要      :サンプルＩＤの取得
''''''ﾊﾟﾗﾒｰﾀ　　:変数名　　　　,IO ,型       ,説明
''''''　　      :iWafPos      ,I  ,Integer　,抜試指示テーブル位置
''''''　　      :sSampID1     ,I  ,String 　,サンプルＩＤ１
''''''　　      :sSampID2     ,I  ,String 　,サンプルＩＤ２
''''''　　      :戻り値　　　　,O  ,Boolean　,選択の有無
''''''説明      :抜試指示サンプルのサンプルＩＤを取得する
''''''履歴      :2001/07/11　 作成
''''''           2003/04/05   hitec)matsumoto bKyotuFlg追加
'''''Public Function GetSampleID(iWafPos As Integer, sSampID1 As String, sSampID2 As String, _
'''''                                                     Optional iKubun As Integer) As Boolean
'''''
'''''    Dim bBot As Boolean
'''''    Dim bTop As Boolean
'''''    Dim bBlk As Boolean
'''''    Dim TargetBlkPos As Integer
'''''    Dim p As Integer
'''''    Dim m As Integer
'''''    Dim i As Integer
'''''
'''''    Dim iHinbanRow  As Integer
'''''    Dim vUpHinban   As Variant
'''''
'''''
'''''    bBot = False
'''''    bTop = False
'''''    bBlk = False
'''''    p = iWafPos
'''''    With tblWafInd(iWafPos)
'''''        m = UBound(tblBlkInf)
''''''        For i = 1 To m
''''''            If .IngotPos = tblBlkInf(i).COF.TOPSMPLPOS Or _
''''''               i = m And .IngotPos = tblBlkInf(i).COF.BOTSMPLPOS Then
''''''                bBlk = True
''''''                Exit For
''''''            End If
''''''        Next i
'''''        For i = 1 To UBound(tblBlkInf)
'''''            If tblWafInd(p).BLOCKID = tblBlkInf(i).BLOCKID Then
'''''                TargetBlkPos = i
'''''                Exit For
'''''            End If
'''''        Next
'''''
'''''        bBot = False
'''''        bTop = False
'''''        Call GetSampleBT(.SMP.CRYINDRS, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDOI, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDB1, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDB2, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDB3, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDL1, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDL2, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDL3, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDL4, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDDS, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDDZ, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDSP, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDD1, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDD2, bTop, bBot)
'''''        Call GetSampleBT(.SMP.CRYINDD3, bTop, bBot)
''''''=========================================2003/04/16 okazaki
''''''上下品番がZ
'''''        If iWafPos >= 1 Then
'''''            If Trim(tblWafInd(iWafPos).HINDN.hinban) = "Z" Or _
'''''               Trim(tblWafInd(iWafPos).HINUP.hinban) = "Z" Or _
'''''               iKubun = 3 Then      'ブロックが変わる初期表示行
'''''                    bTop = True
'''''                    bBot = True
'''''                    bBlk = False
'''''                    'チェックボックス追加によりサンプル切替の判定を追加 (iWafPos - 1を使用するため区分３のときのみ判定する) 2003/06/01 okazaki
'''''                    If Trim(tblWafInd(iWafPos).HINDN.hinban) <> "Z" And tblWafInd(iWafPos).HINDN.hinban = tblWafInd(iWafPos - 1).HINDN.hinban Then
'''''                        bTop = False
'''''                        bBot = False
'''''                        bBlk = False
'''''                    End If
'''''                    '2003/06/01 end
'''''            End If
'''''        End If
''''''=========================================2003/04/16 end
'''''        '' 上方向／下方向サンプル（別）
'''''        If bTop = True And bBot = True Then
'''''            If bBlk = True Then
'''''                If .BlockPos = 0 Then
'''''                    If tblBlkInf(TargetBlkPos - 1).NOWPROC = PROCD_WFC_SOUGOUHANTEI Then
'''''                        sSampID1 = Right(tblBlkInf(TargetBlkPos - 1).BLOCKID, 3) & "-" & GetWafPos(tblBlkInf(TargetBlkPos - 1).LENGTH) & "B"
'''''                        sSampID2 = Right(.BLOCKID, 3) & "-000T"
'''''                    Else
'''''                        sSampID1 = Right(.BLOCKID, 3) & "-000T"
'''''                        sSampID2 = ""
'''''                    End If
'''''                Else
'''''                    sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "B"
'''''                    sSampID2 = Right(tblBlkInf(TargetBlkPos + 1).BLOCKID, 3) & "-000T"
'''''                End If
'''''            Else
'''''                sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "U"
'''''                sSampID2 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "D"
'''''            End If
'''''            GetSampleID = False
'''''        '' 下方向サンプル
'''''        ElseIf bTop = True And bBot = False Then
'''''            If bBlk = True Then
'''''                sSampID1 = Right(.BLOCKID, 3) & "-000T"
'''''            Else
'''''                sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "D"
'''''            End If
'''''            sSampID2 = ""
'''''            GetSampleID = False
'''''        '' 上方向サンプル
'''''        ElseIf bTop = False And bBot = True Then
'''''            If bBlk = True Then
'''''                If .BlockPos = 0 Then
'''''                    sSampID1 = Right(tblBlkInf(TargetBlkPos - 1).BLOCKID, 3) & "-" & GetWafPos(tblBlkInf(TargetBlkPos - 1).LENGTH) & "B"
'''''                Else
'''''                    sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "B"
'''''                End If
'''''            Else
'''''                sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "U"
'''''            End If
'''''            sSampID2 = ""
'''''            GetSampleID = False
'''''        '' 上方向／下方向サンプル（共通）
'''''        ElseIf bTop = False And bBot = False Then
'''''            If bBlk = True Then
'''''                If .BlockPos = 0 Then
'''''                    If tblBlkInf(TargetBlkPos).NOWPROC = PROCD_WFC_SOUGOUHANTEI Then
'''''                        sSampID1 = Right(tblBlkInf(TargetBlkPos).BLOCKID, 3) & "-" & GetWafPos(tblBlkInf(TargetBlkPos - 1).LENGTH) & "B"
'''''                        sSampID2 = Right(.BLOCKID, 3) & "-000T"
'''''                    Else
'''''                        sSampID1 = Right(.BLOCKID, 3) & "-000T"
'''''                        sSampID2 = ""
'''''                        GetSampleID = False
'''''                        Exit Function
'''''                    End If
'''''                Else
'''''                    If tblBlkInf(TargetBlkPos + 1).NOWPROC = PROCD_WFC_SOUGOUHANTEI Then
'''''                        sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "B"
'''''                        sSampID2 = Right(tblBlkInf(TargetBlkPos + 1).BLOCKID, 3) & "-000T"
'''''                    Else
'''''                        sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "B"
'''''                        sSampID2 = ""
'''''                        GetSampleID = False
'''''                        Exit Function
'''''                    End If
'''''                End If
'''''            Else
'''''                sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "U"
'''''                sSampID2 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "D"
'''''            End If
'''''            GetSampleID = True
'''''        End If
'''''    End With
'''''
'''''End Function
'''''============================================================================================================================
'''''
''''''概要      :サンプルのトップ側／ボトム側区分の取得
''''''ﾊﾟﾗﾒｰﾀ　　:変数名　　　　,IO ,型       ,説明
''''''　　      :sSample      ,I  ,String 　,サンプル
''''''　　      :bTop         ,O  ,Boolean　,トップ側区分の有無
''''''　　      :bBot         ,O  ,Boolean　,ボトム側区分の有無
''''''説明      :サンプル区分の有無を返す
''''''履歴      :2001/07/11　 作成
'''''Public Sub GetSampleBT(ByVal sSample As String, bTop As Boolean, bBot As Boolean)
'''''
'''''    Select Case sSample
'''''    Case "1"
'''''        bTop = True
'''''    Case "2"
'''''        bBot = True
'''''    Case "4"
'''''        bTop = True
'''''        bBot = True
'''''    End Select
'''''
'''''End Sub
'''''============================================================================================================================
'''''
''''''概要      :欠落ウェハーテーブルの作成
''''''ﾊﾟﾗﾒｰﾀ　　:変数名　　　　  ,IO ,型              ,説明
''''''　　      :tblBlkInf      ,I  ,typ_BlkInf3   　,ブロック管理構造体
''''''　　      :tmpLackWaf     ,I  ,typ_LackWaf　   ,欠落情報
''''''　　      :BlkInfPos      ,I  ,Integer       　,結晶内全体のブロック数に対する対象ブロックの開始位置
''''''　　      :BlkCnt         ,I  ,Integer　       ,対象ブロック数
''''''　　      :RftblLackMap ,O  ,typ_LackMap     　,ウェハーテーブル構造体
''''''説明      :
''''''履歴      :
'''''Public Function LackMapMake(tblBlkInf() As typ_BlkInf3, tmpLackWaf() As typ_LackWaf, BlkInfPos As Integer, BlkCnt As Integer) As FUNCTION_RETURN
'''''
'''''    Dim bFlag As Boolean
'''''    Dim p As Integer
'''''    Dim m As Integer
'''''    Dim n As Integer
'''''    Dim i As Integer
'''''    Dim j As Integer
'''''    Dim k As Integer
'''''
'''''    '' 欠落ウェハーテーブルの作成
'''''    k = 0
'''''    m = BlkCnt + BlkInfPos - 1
'''''    n = UBound(tmpLackWaf)
'''''    ReDim tblLackMap(n)
'''''
'''''    '' ブロックの始まりから
'''''    For i = BlkInfPos To m
'''''        DoEvents
'''''        For j = 1 To n
'''''            DoEvents
'''''            If tblBlkInf(i).BLOCKID = tmpLackWaf(j).BLOCKID Then
'''''                If bFlag = False Then
'''''                    k = k + 1
'''''                    tblLackMap(k).BLOCKID = tmpLackWaf(j).BLOCKID
'''''                    p = tmpLackWaf(j).WAFERNO
'''''                    If p = -1 Then
'''''                        tblLackMap(k).LACKPOSS = 0
'''''                        tblLackMap(k).LACKCNTS = -1
'''''                        tblLackMap(k).LACKPOSE = tblBlkInf(i).REALLEN
'''''                        tblLackMap(k).LACKCNTE = -1
'''''                        Exit For
'''''                    End If
'''''                    tblLackMap(k).LACKPOSS = tmpLackWaf(j).TOP_POS
'''''                    tblLackMap(k).LACKCNTS = tmpLackWaf(j).WAFERNO
'''''                    bFlag = True
'''''                Else
'''''                    If tmpLackWaf(j).WAFERNO = p + 1 Then
'''''                        p = p + 1
'''''                        If bFlag = True And j = n Then
'''''                            tblLackMap(k).LACKPOSE = tmpLackWaf(j).TAIL_POS
'''''                            tblLackMap(k).LACKCNTE = tmpLackWaf(j).WAFERNO
'''''                        End If
'''''                    Else
'''''                        tblLackMap(k).LACKPOSE = tmpLackWaf(j - 1).TAIL_POS
'''''                        tblLackMap(k).LACKCNTE = tmpLackWaf(j - 1).WAFERNO
'''''                        k = k + 1
'''''                        tblLackMap(k).BLOCKID = tmpLackWaf(j).BLOCKID
'''''                        tblLackMap(k).LACKPOSS = tmpLackWaf(j).TOP_POS
'''''                        tblLackMap(k).LACKCNTS = tmpLackWaf(j).WAFERNO
'''''                        p = tmpLackWaf(j).WAFERNO
'''''                    End If
'''''                End If
'''''            Else
'''''                If bFlag = True Then
'''''                    tblLackMap(k).LACKPOSE = tmpLackWaf(j - 1).TAIL_POS
'''''                    tblLackMap(k).LACKCNTE = tmpLackWaf(j - 1).WAFERNO
'''''                    bFlag = False
'''''                    Exit For
'''''                End If
'''''            End If
'''''        Next j
'''''    Next i
'''''    ReDim Preserve tblLackMap(k)
'''''
'''''    For i = 1 To k
'''''        With tblLackMap(i)
'''''            If .LACKPOSS > 0 And .LACKPOSE = 0 Then
'''''                .LACKPOSE = .LACKPOSS
'''''            End If
'''''            If .LACKCNTS > 0 And .LACKCNTE = 0 Then
'''''                .LACKCNTE = .LACKCNTS
'''''            End If
'''''        End With
'''''    Next
'''''
'''''End Function
'''''============================================================================================================================
'''''
'''''Public Function NoTestCheck(lblMsg As Label) As FUNCTION_RETURN
'''''    Dim c0 As Long
'''''
'''''
'''''    Dim HIN(1) As tFullHinban
'''''    Dim Inf(1) As NoTest_Info
'''''
'''''    NoTestCheck = FUNCTION_RETURN_FAILURE
'''''    For c0 = 1 To 2
'''''        '元品番セット
'''''        HIN(0).factory = tblTotal.typ_Param.factory
'''''        HIN(0).hinban = tblTotal.typ_Param.hinban
'''''        HIN(0).mnorevno = tblTotal.typ_Param.REVNUM
'''''        HIN(0).opecond = tblTotal.typ_Param.opecond
'''''        '振替先品番セット
'''''        If c0 = 1 Then
'''''            HIN(1) = tblWafInd(1).HINDN
'''''        Else
'''''            HIN(1) = tblWafInd(UBound(tblWafInd())).HINUP
'''''        End If
'''''        If Trim(HIN(1).hinban) = "Z" Then
'''''            Exit For
'''''        End If
'''''        If DBDRV_GetNoTestHinInfo(HIN(), Inf()) = FUNCTION_RETURN_FAILURE Then
'''''            Exit Function
'''''        End If
'''''
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESRS <> "1") Then
'''''        '実績無
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).Res.HWFRHWYS = "X") Or (Inf(1).Res.HWFRHWYS = "S") Then
'''''            If (Inf(1).Res.HWFRHWYS = "H") Or (Inf(1).Res.HWFRHWYS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " RES実績無")  '03/06/06 後藤
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESOI <> "1") Then
'''''        '実績無
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).Oi.HWFONHWS = "X") Or (Inf(1).Oi.HWFONHWS = "S") Then
'''''            If (Inf(1).Oi.HWFONHWS = "H") Or (Inf(1).Oi.HWFONHWS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " OI実績無")  '03/06/06 後藤
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESB1 = "1") Then
'''''        '実績有り
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).BMD(0).HWFBMxHS = "X") Or (Inf(1).BMD(0).HWFBMxHS = "S") Then
'''''            If (Inf(1).BMD(0).HWFBMxHS = "H") Or (Inf(1).BMD(0).HWFBMxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                If Inf(0).BMD(0).HWFBMxET <> Inf(1).BMD(0).HWFBMxET Or _
'''''                   Inf(0).BMD(0).HWFBMxNS <> Inf(1).BMD(0).HWFBMxNS Or _
'''''                   Inf(0).BMD(0).HWFBMxSH <> Inf(1).BMD(0).HWFBMxSH Or _
'''''                   Inf(0).BMD(0).HWFBMxSR <> Inf(1).BMD(0).HWFBMxSR Or _
'''''                   Inf(0).BMD(0).HWFBMxST <> Inf(1).BMD(0).HWFBMxST Or _
'''''                   Inf(0).BMD(0).HWFBMxSZ <> Inf(1).BMD(0).HWFBMxSZ Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " BMD1")  '03/06/06
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
'''''        '実績無
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).BMD(0).HWFBMxHS = "X") Or (Inf(1).BMD(0).HWFBMxHS = "S") Then
'''''            If (Inf(1).BMD(0).HWFBMxHS = "H") Or (Inf(1).BMD(0).HWFBMxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " BMD1実績無")  '03/06/06 後藤
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESB2 = "1") Then
'''''        '実績有り
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).BMD(1).HWFBMxHS = "X") Or (Inf(1).BMD(1).HWFBMxHS = "S") Then
'''''            If (Inf(1).BMD(1).HWFBMxHS = "H") Or (Inf(1).BMD(1).HWFBMxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                If Inf(0).BMD(1).HWFBMxET <> Inf(1).BMD(1).HWFBMxET Or _
'''''                   Inf(0).BMD(1).HWFBMxNS <> Inf(1).BMD(1).HWFBMxNS Or _
'''''                   Inf(0).BMD(1).HWFBMxSH <> Inf(1).BMD(1).HWFBMxSH Or _
'''''                   Inf(0).BMD(1).HWFBMxSR <> Inf(1).BMD(1).HWFBMxSR Or _
'''''                   Inf(0).BMD(1).HWFBMxST <> Inf(1).BMD(1).HWFBMxST Or _
'''''                   Inf(0).BMD(1).HWFBMxSZ <> Inf(1).BMD(1).HWFBMxSZ Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " BMD2")  '03/06/06 後藤
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).BMD(1).HWFBMxHS = "X") Or (Inf(1).BMD(1).HWFBMxHS = "S") Then
'''''            If (Inf(1).BMD(1).HWFBMxHS = "H") Or (Inf(1).BMD(1).HWFBMxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " BMD2実績無")  '03/06/06
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESB3 = "1") Then
'''''        '実績有り
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).BMD(2).HWFBMxHS = "X") Or (Inf(1).BMD(2).HWFBMxHS = "S") Then
'''''            If (Inf(1).BMD(2).HWFBMxHS = "H") Or (Inf(1).BMD(2).HWFBMxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                If Inf(0).BMD(2).HWFBMxET <> Inf(1).BMD(2).HWFBMxET Or _
'''''                   Inf(0).BMD(2).HWFBMxNS <> Inf(1).BMD(2).HWFBMxNS Or _
'''''                   Inf(0).BMD(2).HWFBMxSH <> Inf(1).BMD(2).HWFBMxSH Or _
'''''                   Inf(0).BMD(2).HWFBMxSR <> Inf(1).BMD(2).HWFBMxSR Or _
'''''                   Inf(0).BMD(2).HWFBMxST <> Inf(1).BMD(2).HWFBMxST Or _
'''''                   Inf(0).BMD(2).HWFBMxSZ <> Inf(1).BMD(2).HWFBMxSZ Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " BMD3") '03/06/06
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).BMD(2).HWFBMxHS = "X") Or (Inf(1).BMD(2).HWFBMxHS = "S") Then
'''''            If (Inf(1).BMD(2).HWFBMxHS = "H") Or (Inf(1).BMD(2).HWFBMxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " BMD3実績無")  '03/06/06 後藤
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESL1 = "1") Then
'''''        '実績有り
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).OSF(0).HWFOFxHS = "X") Or (Inf(1).OSF(0).HWFOFxHS = "S") Then
'''''            If (Inf(1).OSF(0).HWFOFxHS = "H") Or (Inf(1).OSF(0).HWFOFxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                If Inf(0).OSF(0).HWFOFxET <> Inf(1).OSF(0).HWFOFxET Or _
'''''                   Inf(0).OSF(0).HWFOFxNS <> Inf(1).OSF(0).HWFOFxNS Or _
'''''                   Inf(0).OSF(0).HWFOFxSH <> Inf(1).OSF(0).HWFOFxSH Or _
'''''                   Inf(0).OSF(0).HWFOFxSR <> Inf(1).OSF(0).HWFOFxSR Or _
'''''                   Inf(0).OSF(0).HWFOFxST <> Inf(1).OSF(0).HWFOFxST Or _
'''''                   Inf(0).OSF(0).HWFOFxSZ <> Inf(1).OSF(0).HWFOFxSZ Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " OSF1")  '03/06/06 後藤
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).OSF(0).HWFOFxHS = "X") Or (Inf(1).OSF(0).HWFOFxHS = "S") Then
'''''            If (Inf(1).OSF(0).HWFOFxHS = "H") Or (Inf(1).OSF(0).HWFOFxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " OSF1実績無") '03/06/06 後藤
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESL2 = "1") Then
'''''        '実績有り
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).OSF(1).HWFOFxHS = "X") Or (Inf(1).OSF(1).HWFOFxHS = "S") Then
'''''            If (Inf(1).OSF(1).HWFOFxHS = "H") Or (Inf(1).OSF(1).HWFOFxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                If Inf(0).OSF(1).HWFOFxET <> Inf(1).OSF(1).HWFOFxET Or _
'''''                   Inf(0).OSF(1).HWFOFxNS <> Inf(1).OSF(1).HWFOFxNS Or _
'''''                   Inf(0).OSF(1).HWFOFxSH <> Inf(1).OSF(1).HWFOFxSH Or _
'''''                   Inf(0).OSF(1).HWFOFxSR <> Inf(1).OSF(1).HWFOFxSR Or _
'''''                   Inf(0).OSF(1).HWFOFxST <> Inf(1).OSF(1).HWFOFxST Or _
'''''                   Inf(0).OSF(1).HWFOFxSZ <> Inf(1).OSF(1).HWFOFxSZ Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " OSF2") '03/06/06 後藤
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).OSF(1).HWFOFxHS = "X") Or (Inf(1).OSF(1).HWFOFxHS = "S") Then
'''''            If (Inf(1).OSF(1).HWFOFxHS = "H") Or (Inf(1).OSF(1).HWFOFxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " OSF2実績無")  '03/06/06 後藤
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESL3 = "1") Then
'''''        '実績有り
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).OSF(2).HWFOFxHS = "X") Or (Inf(1).OSF(2).HWFOFxHS = "S") Then
'''''            If (Inf(1).OSF(2).HWFOFxHS = "H") Or (Inf(1).OSF(2).HWFOFxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                If Inf(0).OSF(2).HWFOFxET <> Inf(1).OSF(2).HWFOFxET Or _
'''''                   Inf(0).OSF(2).HWFOFxNS <> Inf(1).OSF(2).HWFOFxNS Or _
'''''                   Inf(0).OSF(2).HWFOFxSH <> Inf(1).OSF(2).HWFOFxSH Or _
'''''                   Inf(0).OSF(2).HWFOFxSR <> Inf(1).OSF(2).HWFOFxSR Or _
'''''                   Inf(0).OSF(2).HWFOFxST <> Inf(1).OSF(2).HWFOFxST Or _
'''''                   Inf(0).OSF(2).HWFOFxSZ <> Inf(1).OSF(2).HWFOFxSZ Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " OSF3")  '03/06/06 後藤
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).OSF(2).HWFOFxHS = "X") Or (Inf(1).OSF(2).HWFOFxHS = "S") Then
'''''            If (Inf(1).OSF(2).HWFOFxHS = "H") Or (Inf(1).OSF(2).HWFOFxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " OSF3実績無")  '03/06/06 後藤
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESL4 = "1") Then
'''''        '実績有り
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).OSF(3).HWFOFxHS = "X") Or (Inf(1).OSF(3).HWFOFxHS = "S") Then
'''''            If (Inf(1).OSF(3).HWFOFxHS = "H") Or (Inf(1).OSF(3).HWFOFxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                If Inf(0).OSF(3).HWFOFxET <> Inf(1).OSF(3).HWFOFxET Or _
'''''                   Inf(0).OSF(3).HWFOFxNS <> Inf(1).OSF(3).HWFOFxNS Or _
'''''                   Inf(0).OSF(3).HWFOFxSH <> Inf(1).OSF(3).HWFOFxSH Or _
'''''                   Inf(0).OSF(3).HWFOFxSR <> Inf(1).OSF(3).HWFOFxSR Or _
'''''                   Inf(0).OSF(3).HWFOFxST <> Inf(1).OSF(3).HWFOFxST Or _
'''''                   Inf(0).OSF(3).HWFOFxSZ <> Inf(1).OSF(3).HWFOFxSZ Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " OSF4")   '03/06/06 後藤
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).OSF(3).HWFOFxHS = "X") Or (Inf(1).OSF(3).HWFOFxHS = "S") Then
'''''            If (Inf(1).OSF(3).HWFOFxHS = "H") Or (Inf(1).OSF(3).HWFOFxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " OSF4実績無")  '03/06/06 後藤
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESDS = "1") Then
'''''        '実績有り
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).Dsod.HWFDSOHS = "X") Or (Inf(1).Dsod.HWFDSOHS = "S") Then
'''''            If (Inf(1).Dsod.HWFDSOHS = "H") Or (Inf(1).Dsod.HWFDSOHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                If Inf(0).Dsod.HWFDSOKE <> Inf(1).Dsod.HWFDSOKE Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " DSOD")  '03/06/06 後藤
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).Dsod.HWFDSOHS = "X") Or (Inf(1).Dsod.HWFDSOHS = "S") Then
'''''            If (Inf(1).Dsod.HWFDSOHS = "H") Or (Inf(1).Dsod.HWFDSOHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " DSOD実績無")   '03/06/06 後藤
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESDZ = "1") Then
'''''        '実績有り
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).Dz.HWFMKHWS = "X") Or (Inf(1).Dz.HWFMKHWS = "S") Then
'''''            If (Inf(1).Dz.HWFMKHWS = "H") Or (Inf(1).Dz.HWFMKHWS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                If Inf(0).Dz.HWFMKSPH <> Inf(1).Dz.HWFMKSPH Or _
'''''                   Inf(0).Dz.HWFMKSPR <> Inf(1).Dz.HWFMKSPR Or _
'''''                   Inf(0).Dz.HWFMKSPT <> Inf(1).Dz.HWFMKSPT Or _
'''''                   Inf(0).Dz.HWFMKSZY <> Inf(1).Dz.HWFMKSZY Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " DZ")   '03/06/06 後藤
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).Dz.HWFMKHWS = "X") Or (Inf(1).Dz.HWFMKHWS = "S") Then
'''''            If (Inf(1).Dz.HWFMKHWS = "H") Or (Inf(1).Dz.HWFMKHWS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " DZ実績無")  '03/06/06 後藤
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESSP = "1") Then
'''''        '実績有り
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).SpvFe.HWFSPVHS = "X") Or (Inf(1).SpvFe.HWFSPVHS = "S") Then
'''''            If (Inf(1).SpvFe.HWFSPVHS = "H") Or (Inf(1).SpvFe.HWFSPVHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                If Inf(0).SpvFe.HWFSPVSH <> Inf(1).SpvFe.HWFSPVSH Or _
'''''                   Inf(0).SpvFe.HWFSPVSI <> Inf(1).SpvFe.HWFSPVSI Or _
'''''                   Inf(0).SpvFe.HWFSPVST <> Inf(1).SpvFe.HWFSPVST Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " SPVFE")  '03/06/06 後藤
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).SpvFe.HWFSPVHS = "X") Or (Inf(1).SpvFe.HWFSPVHS = "S") Then
'''''            If (Inf(1).SpvFe.HWFSPVHS = "H") Or (Inf(1).SpvFe.HWFSPVHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''            '検査有り
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " SPVFE実績無")  '03/06/06 後藤
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESSP = "1") Then
'''''        '実績有り
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).Spv.HWFDLHWS = "X") Or (Inf(1).Spv.HWFDLHWS = "S") Then
'''''            If (Inf(1).Spv.HWFDLHWS = "H") Or (Inf(1).Spv.HWFDLHWS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                If Inf(0).Spv.HWFDLSPH <> Inf(1).Spv.HWFDLSPH Or _
'''''                   Inf(0).Spv.HWFDLSPI <> Inf(1).Spv.HWFDLSPI Or _
'''''                   Inf(0).Spv.HWFDLSPT <> Inf(1).Spv.HWFDLSPT Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " SPV拡散長")  '03/06/06 後藤
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).Spv.HWFDLHWS = "X") Or (Inf(1).Spv.HWFDLHWS = "S") Then
'''''            If (Inf(1).Spv.HWFDLHWS = "H") Or (Inf(1).Spv.HWFDLHWS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''            '検査有り
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " SPV拡散長実績無")  '03/06/06 後藤
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESDO1 = "1") Then
'''''        '実績有り
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).Doi(0).HWFOSxHS = "X") Or (Inf(1).Doi(0).HWFOSxHS = "S") Then
'''''            If (Inf(1).Doi(0).HWFOSxHS = "H") Or (Inf(1).Doi(0).HWFOSxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                If Inf(0).Doi(0).HWFOSxNS <> Inf(1).Doi(0).HWFOSxNS Or _
'''''                   Inf(0).Doi(0).HWFOSxSH <> Inf(1).Doi(0).HWFOSxSH Or _
'''''                   Inf(0).Doi(0).HWFOSxSI <> Inf(1).Doi(0).HWFOSxSI Or _
'''''                   Inf(0).Doi(0).HWFOSxST <> Inf(1).Doi(0).HWFOSxST Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " ⊿Oi1")  '03/06/06 後藤
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).Doi(0).HWFOSxHS = "X") Or (Inf(1).Doi(0).HWFOSxHS = "S") Then
'''''            If (Inf(1).Doi(0).HWFOSxHS = "H") Or (Inf(1).Doi(0).HWFOSxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " ⊿Oi1実績無")  '03/06/06 後藤
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESDO2 = "1") Then
'''''        '実績有り
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).Doi(1).HWFOSxHS = "X") Or (Inf(1).Doi(1).HWFOSxHS = "S") Then
'''''            If (Inf(1).Doi(1).HWFOSxHS = "H") Or (Inf(1).Doi(1).HWFOSxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                If Inf(0).Doi(0).HWFOSxNS <> Inf(1).Doi(0).HWFOSxNS Or _
'''''                   Inf(0).Doi(0).HWFOSxSH <> Inf(1).Doi(0).HWFOSxSH Or _
'''''                   Inf(0).Doi(0).HWFOSxSI <> Inf(1).Doi(0).HWFOSxSI Or _
'''''                   Inf(0).Doi(0).HWFOSxST <> Inf(1).Doi(0).HWFOSxST Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " ⊿Oi2")  '03/06/06 後藤
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).Doi(1).HWFOSxHS = "X") Or (Inf(1).Doi(1).HWFOSxHS = "S") Then
'''''            If (Inf(1).Doi(1).HWFOSxHS = "H") Or (Inf(1).Doi(1).HWFOSxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " ⊿Oi2実績無")  '03/06/06 後藤
'''''                Exit Function
'''''            End If
'''''        End If
'''''        If (tblTotal.typ_Param.WFSMP(c0).WFRESDO3 = "1") Then
'''''        '実績有り
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).Doi(2).HWFOSxHS = "X") Or (Inf(1).Doi(2).HWFOSxHS = "S") Then
'''''            If (Inf(1).Doi(2).HWFOSxHS = "H") Or (Inf(1).Doi(2).HWFOSxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                If Inf(0).Doi(0).HWFOSxNS <> Inf(1).Doi(0).HWFOSxNS Or _
'''''                   Inf(0).Doi(0).HWFOSxSH <> Inf(1).Doi(0).HWFOSxSH Or _
'''''                   Inf(0).Doi(0).HWFOSxSI <> Inf(1).Doi(0).HWFOSxSI Or _
'''''                   Inf(0).Doi(0).HWFOSxST <> Inf(1).Doi(0).HWFOSxST Then
'''''                    lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " ⊿Oi3")  '03/06/06 後藤
'''''                    Exit Function
'''''                End If
'''''            End If
'''''        Else
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
''''''            If (Inf(1).Doi(2).HWFOSxHS = "X") Or (Inf(1).Doi(2).HWFOSxHS = "S") Then
'''''            If (Inf(1).Doi(2).HWFOSxHS = "H") Or (Inf(1).Doi(2).HWFOSxHS = "S") Then
''''''ＷＦサンプル処理変更 2003.05.20 yakimura
'''''
'''''            '検査有り
'''''                lblMsg.Caption = GetMsgStr("EHINC", Trim(HIN(1).hinban) & " " & " ⊿Oi3実績無")  '03/06/06 後藤
'''''                Exit Function
'''''            End If
'''''        End If
'''''    Next
'''''
'''''    NoTestCheck = FUNCTION_RETURN_SUCCESS
'''''
'''''End Function

''概要      :指定の結晶番号に含まれる品番の一覧を得る
''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型          ,説明
''          :cryno         ,I  ,String      ,結晶番号
''          :hinban()      ,O  ,tFullHinban ,品番リスト
''          :戻り値        ,O  ,FUNCTION_RETURN,抽出の成否
''説明      :
''履歴      :2001/06/27 作成  長野 (2002/07 s_cmzc010a.basより移動)
'Public Function GetXlHinban(cryno$, HINBAN() As tFullHinban) As FUNCTION_RETURN
'Dim rs      As OraDynaset               '抽出RecordDynaset
'Dim rsCnt   As Integer                  'ﾚｺｰﾄﾞｶｳﾝﾄ
'Dim sql     As String                   'SQL文
'Dim i       As Integer                  'ﾙｰﾌﾟｶｳﾝﾄ
'
'    'エラーハンドラの設定
'    On Error GoTo proc_err
''(2002/07)    gErr.Push "s_cmzc010a.bas -- Function GetXlHinban"
'    gErr.Push "-- Function GetXlHinban"
'
'    'SQL文の作成
'    sql = "Select CRYNUM, HINBAN, REVNUM, FACTORY, OPECOND from TBCME041 "
'    sql = sql & "Where(CRYNUM = '" & cryno & "')"
'
'    'データの抽出
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'    '''抽出レコードが存在しない場合
'    If rs.EOF Then
'        ReDim HINBAN(0)                     '配列の初期化
'        GetXlHinban = FUNCTION_RETURN_FAILURE   'ｴﾗｰｽﾃｰﾀｽ
'        GoTo proc_exit
'    End If
'
'    rsCnt = rs.RecordCount                  'ﾚｺｰﾄﾞ数のｶｳﾝﾄを取る
'    ReDim HINBAN(rsCnt - 1)                 '配列の再定義
'
'    '配列に値をセット
'    rs.MoveFirst                            '先頭ﾚｺｰﾄﾞに移動
'    For i = 0 To rsCnt - 1                  'ﾚｺｰﾄﾞ数分ﾙｰﾌﾟ
'        DoEvents
'        With HINBAN(i)
'            .HINBAN = rs!HINBAN             '品番
'            .mnorevno = rs!REVNUM           '製品番号改訂番号
'            .FACTORY = rs!FACTORY           '工場
'            .OPECOND = rs!OPECOND           '操業条件
'        End With
'        rs.MoveNext                         '次ﾚｺｰﾄﾞに移動
'    Next
'
'    GetXlHinban = FUNCTION_RETURN_SUCCESS   '正常ｽﾃｰﾀｽ
'
'
'proc_exit:
'    '終了
'    gErr.Pop
'    Exit Function
'
'proc_err:
'    'エラーハンドラ
'    gErr.HandleError
'    Resume proc_exit
'End Function
'''''============================================================================================================================
'''''
'''''Public Function DBData2DispData(data As Variant, Optional Formatstr As String) As Variant
'''''    If data = -1 Then
'''''        DBData2DispData = ""
'''''    Else
'''''        If Formatstr = "" Then
'''''            DBData2DispData = data
'''''        Else
'''''            DBData2DispData = Format(data, Formatstr)
'''''        End If
'''''    End If
'''''End Function
'''''============================================================================================================================
'''''
''''''概要      :SXL IDの取得
''''''ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型       ,説明
''''''　　      :sBlockID 　　　,I  ,String 　,ブロックID
''''''　　      :iIngotPos　　　,I  ,Integer　,結晶内開始位置
''''''　　      :戻り値         ,O  ,String 　,SXL ID
''''''説明      :SXL IDを返す
''''''履歴      :2001/07/11　大塚 作成
'''''Public Function GetSXLID(sBlockID As String, iIngotpos As Integer) As String
'''''
'''''    GetSXLID = Left(sBlockID, 10) & GetWafPos(iIngotpos)
'''''
'''''End Function
'''''============================================================================================================================
'''''
''''''概要      :抜試位置文字列の取得
''''''ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型       ,説明
''''''　　      :iIngotPos　　　,I  ,Integer　,結晶内開始位置
''''''　　      :戻り値         ,O  ,String 　,抜試位置文字列
''''''説明      :抜試位置文字列を返す
''''''履歴      :2001/07/11　大塚 作成
'''''Public Function GetWafPos(iIngotpos As Integer) As String
'''''
'''''    Dim i As Integer
'''''    Dim j As Integer
'''''
'''''    If iIngotpos >= 1000 Then
'''''        i = Int(iIngotpos / 100)
'''''        j = iIngotpos Mod 100
'''''        GetWafPos = Chr$(i - 10 + Asc("A")) & Format(j, "00")
'''''    Else
'''''        GetWafPos = Format(iIngotpos, "000")
'''''    End If
'''''
'''''End Function
'''''============================================================================================================================
'''''
''''''概要      :測定評価方法指示テーブルの作成
''''''ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
''''''　　      :pSXLMng　　　,I  ,typ_TBCME042   　,SXL管理
''''''　　      :pWafSmp　　　,I  ,typ_TBCME044   　,WFサンプル管理
''''''　　      :pMesInd　　　,O  ,typ_TBCMY003   　,測定評価方法指示
''''''　　      :戻り値       ,O  ,FUNCTION_RETURN　,読み込みの成否
''''''説明      :測定評価方法指示テーブルを作成する
''''''履歴      :2001/07/23　大塚 作成
'''''Public Function MakeMesIndTbl(pSXLMng() As typ_TBCME042, pWafSmp() As typ_TBCME044, pMesInd() As typ_TBCMY003) As FUNCTION_RETURN
'''''
'''''    Dim tmpSpWFSamp() As typ_SpWFSamp
'''''    Dim sHin As String
'''''    Dim sDKAN As String
'''''    Dim m As Integer
'''''    Dim n As Integer
'''''    Dim i As Integer
'''''    Dim j As Integer
'''''    Dim k As Integer
'''''
'''''    '' 測定評価方法指示用の製品仕様を取得
'''''    j = 0
'''''    m = UBound(pSXLMng)
'''''    ReDim tmpSpWFSamp(m)
'''''    For i = 1 To m
''''''==================================== 2003/04/17 okazaki
'''''            sHin = RTrim$(pSXLMng(i).hinban)
'''''        If (sHin <> "" And sHin <> "G" And sHin <> "Z") Then
''''''=======================================================
'''''            j = j + 1
'''''            tmpSpWFSamp(j).HIN.hinban = pSXLMng(i).hinban
'''''            tmpSpWFSamp(j).HIN.mnorevno = pSXLMng(i).REVNUM
'''''            tmpSpWFSamp(j).HIN.factory = pSXLMng(i).factory
'''''            tmpSpWFSamp(j).HIN.opecond = pSXLMng(i).opecond
'''''            If scmzc_getWF(tmpSpWFSamp(j)) = FUNCTION_RETURN_FAILURE Then
'''''                MakeMesIndTbl = FUNCTION_RETURN_FAILURE
'''''                Exit Function
'''''            End If
'''''        End If
'''''    Next i
'''''    ReDim Preserve tmpSpWFSamp(j)
'''''
'''''    '' 測定評価方法指示テーブルの作成
'''''    k = 0
'''''    m = UBound(pWafSmp)
'''''    n = UBound(tmpSpWFSamp)
'''''
'''''    ReDim pMesInd(m * 17)   ''### Add.03/05/20 後藤 ###
''''''''    ReDim pMesInd(m * 15)
'''''    For i = 1 To m
'''''        For j = 1 To n
'''''            If tmpSpWFSamp(j).HIN.hinban = pWafSmp(i).hinban Then
'''''                Exit For
'''''            End If
'''''        Next j
'''''        If j <= n Then
'''''            With tmpSpWFSamp(j)
'''''                sDKAN = IIf(.HWFIGKBN = "3", "R ", "V ") & Format(.HWFANTNP, "@@@@") & Format(.HWFANTIM, "@@@@")
'''''                If pWafSmp(i).WFINDRS <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "RES"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "RES"
'''''                    pMesInd(k).NETSU = ""
'''''                    pMesInd(k).ET = ""
'''''                    pMesInd(k).MES = .HWFRSPOH & .HWFRSPOT & .HWFRSPOI
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDOI <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "OI"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "OI"
'''''                    pMesInd(k).NETSU = ""
'''''                    pMesInd(k).ET = ""
'''''                    pMesInd(k).MES = .HWFONSPH & .HWFONSPT & .HWFONSPI
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDB1 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "BMD"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "BMD1"
'''''                    pMesInd(k).NETSU = .HWFBM1NS
'''''                    pMesInd(k).ET = .HWFBM1SZ & Format(.HWFBM1ET, "00")
'''''                    pMesInd(k).MES = .HWFBM1SH & .HWFBM1ST & .HWFBM1SR
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDB2 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "BMD"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "BMD2"
'''''                    pMesInd(k).NETSU = .HWFBM2NS
'''''                    pMesInd(k).ET = .HWFBM2SZ & Format(.HWFBM2ET, "00")
'''''                    pMesInd(k).MES = .HWFBM2SH & .HWFBM2ST & .HWFBM2SR
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDB3 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "BMD"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "BMD3"
'''''                    pMesInd(k).NETSU = .HWFBM3NS
'''''                    pMesInd(k).ET = .HWFBM3SZ & Format(.HWFBM3ET, "00")
'''''                    pMesInd(k).MES = .HWFBM3SH & .HWFBM3ST & .HWFBM3SR
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDL1 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "OSF"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "OSF1"
'''''                    pMesInd(k).NETSU = .HWFOF1NS
'''''                    pMesInd(k).ET = .HWFOF1SZ & Format(.HWFOF1ET, "00")
'''''                    pMesInd(k).MES = .HWFOF1SH & .HWFOF1ST & .HWFOF1SR
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDL2 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "OSF"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "OSF2"
'''''                    pMesInd(k).NETSU = .HWFOF2NS
'''''                    pMesInd(k).ET = .HWFOF2SZ & Format(.HWFOF2ET, "00")
'''''                    pMesInd(k).MES = .HWFOF2SH & .HWFOF2ST & .HWFOF2SR
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDL3 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "OSF"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "OSF3"
'''''                    pMesInd(k).NETSU = .HWFOF3NS
'''''                    pMesInd(k).ET = .HWFOF3SZ & Format(.HWFOF3ET, "00")
'''''                    pMesInd(k).MES = .HWFOF3SH & .HWFOF3ST & .HWFOF3SR
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDL4 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "OSF"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "OSF4"
'''''                    pMesInd(k).NETSU = .HWFOF4NS
'''''                    pMesInd(k).ET = .HWFOF4SZ & Format(.HWFOF4ET, "00")
'''''                    pMesInd(k).MES = .HWFOF4SH & .HWFOF4ST & .HWFOF4SR
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDDS <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "DSOD"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "DSOD"
'''''                    pMesInd(k).NETSU = "G0"
'''''                    pMesInd(k).ET = ""
'''''                    pMesInd(k).MES = ""
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDDZ <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "DZ"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "DZ"
'''''                    pMesInd(k).NETSU = .HWFMKNSW
'''''                    pMesInd(k).ET = .HWFMKSZY & Format(.HWFMKCET, "00")
'''''                    pMesInd(k).MES = .HWFMKSPH & .HWFMKSPT & .HWFMKSPR
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDSP <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "SPV"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "SPV"
'''''                    pMesInd(k).NETSU = ""
'''''                    pMesInd(k).ET = ""
'''''                    pMesInd(k).MES = .HWFSPVSH & .HWFSPVST & .HWFSPVSI
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDDO1 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "DOI"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "DOI1"
'''''                    pMesInd(k).NETSU = .HWFOS1NS
'''''                    pMesInd(k).ET = ""
'''''                    pMesInd(k).MES = .HWFOS1SH & .HWFOS1ST & .HWFOS1SI
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDDO2 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "DOI"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "DOI2"
'''''                    pMesInd(k).NETSU = .HWFOS2NS
'''''                    pMesInd(k).ET = ""
'''''                    pMesInd(k).MES = .HWFOS2SH & .HWFOS2ST & .HWFOS2SI
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                If pWafSmp(i).WFINDDO3 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''                    pMesInd(k).OSITEM = "DOI"
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "DOI3"
'''''                    pMesInd(k).NETSU = .HWFOS3NS
'''''                    pMesInd(k).ET = ""
'''''                    pMesInd(k).MES = .HWFOS3SH & .HWFOS3ST & .HWFOS3SI
'''''                    pMesInd(k).DKAN = sDKAN
'''''                End If
'''''                '################################## Add,03/05/20 後藤 ##########
'''''                If pWafSmp(i).WFINDOT1 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''''''                    pMesInd(k).OSITEM = "OTH"
'''''                    pMesInd(k).OSITEM = "OTH1"  'upd 2003/06/10 hitec)matsumoto
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "OTHER1"
''''''''                    pMesInd(k).NETSU = .HWFOS3NS
''''''''                    pMesInd(k).ET = ""
''''''''                    pMesInd(k).MES = .HWFOS3SH & .HWFOS3ST & .HWFOS3SI
'''''                    pMesInd(k).NETSU = vbNullString
'''''                    pMesInd(k).ET = vbNullString
'''''                    pMesInd(k).MES = vbNullString
'''''''''                    pMesInd(k).DKAN = vbNullString  '03/05/22
'''''                    pMesInd(k).DKAN = sDKAN 'upd 2003/09/10 hitec)matsumoto
'''''                End If
'''''                If pWafSmp(i).WFINDOT2 <> "0" Then
'''''                    k = k + 1
'''''                    pMesInd(k).SAMPLEID = pWafSmp(i).SMPLID
'''''''''                    pMesInd(k).OSITEM = "OTH"
'''''                    pMesInd(k).OSITEM = "OTH2"  'upd 2003/06/10 hitec)matsumoto
'''''                    pMesInd(k).SAMPLEKB = "A"
'''''                    pMesInd(k).Spec = "OTHER2"
''''''''                    pMesInd(k).NETSU = .HWFOS3NS
''''''''                    pMesInd(k).ET = ""
''''''''                    pMesInd(k).MES = .HWFOS3SH & .HWFOS3ST & .HWFOS3SI
'''''                    pMesInd(k).NETSU = vbNullString
'''''                    pMesInd(k).ET = vbNullString
'''''                    pMesInd(k).MES = vbNullString
'''''''''                    pMesInd(k).DKAN = vbNullString  '03/05/22
'''''                    pMesInd(k).DKAN = sDKAN 'upd 2003/09/10 hitec)matsumoto
'''''                End If
'''''                '################################## End,03/05/20 後藤 ##########
'''''            End With
'''''        End If
'''''    Next i
'''''    ReDim Preserve pMesInd(k)
'''''
'''''    MakeMesIndTbl = FUNCTION_RETURN_SUCCESS
'''''
'''''End Function
'''''============================================================================================================================
'''''
''''''概要      :Ｕ／Ｄを上下に分割
''''''説明      :Ｕ／Ｄサンプルを上下に分割する
''''''履歴      :2001/10/05　大塚 作成
'''''Public Sub SeparateUD()
'''''    'Step3.2にて、機能廃止
'''''End Sub

'概要      :変数初期化
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型              ,説明
'          :typ_A         ,IO ,typ_AllTypes ,各情報構造体
'説明      :
'履歴      :
Public Sub InitHensu2(typ_C As typ_AllTypesC)
    Dim i As Integer, j As Integer
    
'Chg Start 2011/03/10 SMPK Miyata
'    For i = 1 To 2
    For i = 1 To SXL_MAXSMP
'Chg End   2011/03/10 SMPK Miyata
        For j = 0 To MAXCNT
            With typ_C
                .typ_y013(i, j).SAMPLEID = "0"         ' サンプルID
                .typ_y013(i, j).MESDATA1 = "0"         ' 測定データその１
                .typ_y013(i, j).MESDATA2 = "0"         ' 測定データその２
                .typ_y013(i, j).MESDATA3 = "0"         ' 測定データその３
                .typ_y013(i, j).MESDATA4 = "0"         ' 測定データその４
                .typ_y013(i, j).MESDATA5 = "0"         ' 測定データその５
                .typ_y013(i, j).MESDATA6 = "0"         ' 測定データその６
                .typ_y013(i, j).MESDATA7 = "0"         ' 測定データその７
                .typ_y013(i, j).MESDATA8 = "0"         ' 測定データその８
                .typ_y013(i, j).MESDATA9 = "0"         ' 測定データその９
                .typ_y013(i, j).MESDATA10 = "0"        ' 測定データその１０
                .typ_y013(i, j).MESDATA11 = "0"        ' 測定データその１１
                .typ_y013(i, j).MESDATA12 = "0"        ' 測定データその１２
                .typ_y013(i, j).MESDATA13 = "0"        ' 測定データその１３
                .typ_y013(i, j).MESDATA14 = "0"        ' 測定データその１４
                .typ_y013(i, j).MESDATA15 = "0"        ' 測定データその１５
            End With
        Next
    Next
End Sub

'概要      :変数初期化
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型              ,説明
'          :typ_A_EP      ,IO ,typ_AllTypes_EP ,各情報構造体
'説明      :
'履歴      :2006/08/15 新規作成 エピ先行評価追加対応 SMP)kondoh
Public Sub InitHensu2_EP(typ_C_EP As typ_AllTypesC_EP)
    Dim i As Integer, j As Integer
    
    For i = 1 To 2
        For j = 0 To MAXCNT_EP
            With typ_C_EP
                .typ_y022(i, j).SAMPLEID = "0"          ' サンプルID
                .typ_y022(i, j).MESDATA1 = "0"         ' 測定データその１
                .typ_y022(i, j).MESDATA2 = "0"         ' 測定データその２
                .typ_y022(i, j).MESDATA3 = "0"         ' 測定データその３
                .typ_y022(i, j).MESDATA4 = "0"         ' 測定データその４
                .typ_y022(i, j).MESDATA5 = "0"         ' 測定データその５
                .typ_y022(i, j).MESDATA6 = "0"         ' 測定データその６
                .typ_y022(i, j).MESDATA7 = "0"         ' 測定データその７
                .typ_y022(i, j).MESDATA8 = "0"         ' 測定データその８
                .typ_y022(i, j).MESDATA9 = "0"         ' 測定データその９
                .typ_y022(i, j).MESDATA10 = "0"        ' 測定データその１０
                .typ_y022(i, j).MESDATA11 = "0"        ' 測定データその１１
                .typ_y022(i, j).MESDATA12 = "0"        ' 測定データその１２
                .typ_y022(i, j).MESDATA13 = "0"        ' 測定データその１３
                .typ_y022(i, j).MESDATA14 = "0"        ' 測定データその１４
                .typ_y022(i, j).MESDATA15 = "0"        ' 測定データその１５
            End With
        Next
    Next
End Sub

'------------------------------------------------
' 仕様Nullチェック(WFC)
'------------------------------------------------

'概要      :WFC総合判定の各検査項目の保証方法が'H'または'S'の場合、仕様値がNull(-1)かどうかを判断する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :tSiyou        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :品番、仕様、結晶内側取得用
'          :sErrMsg       ,IO ,String                               :ｴﾗｰﾒｯｾｰｼﾞ
'          :戻り値        ,O  ,FUNCTION_RETURN                      :結果 = FUNCTION_RETURN_SUCCESS : OK
'                                                                           FUNCTION_RETURN_FAILURE : NG
'説明      :
'履歴      :2003/12/13 新規作成　システムブレイン

Private Function funWfChkNull(tSiyou As type_DBDRV_scmzc_fcmlc001c_Siyou, sErrMsg As String) As FUNCTION_RETURN
    Dim dShiyo()    As Double
    Dim sHosyo      As String
    Dim cnt         As Integer
    
    '初期化
    funWfChkNull = FUNCTION_RETURN_SUCCESS
    
    '--------------- RS(比抵抗) ---------------
    ReDim dShiyo(5)
    dShiyo(1) = tSiyou.HWFRMIN          ' 品ＷＦ比抵抗下限
    dShiyo(2) = tSiyou.HWFRMAX          ' 品ＷＦ比抵抗上限
    dShiyo(3) = tSiyou.HWFRAMIN         ' 品ＷＦ比抵抗平均下限
    dShiyo(4) = tSiyou.HWFRAMAX         ' 品ＷＦ比抵抗平均上限
    dShiyo(5) = tSiyou.HWFRMBNP         ' 品ＷＦ比抵抗面内分布
    If fncJissekiHantei_nl(tSiyou.HWFRHWYS, dShiyo) = False Then
        sErrMsg = sErrMsg & "(RS)"
        funWfChkNull = FUNCTION_RETURN_FAILURE
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00130"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        Exit Function
    End If
    
    '--------------- Oi(酸素濃度) ---------------
    ReDim dShiyo(5)
    dShiyo(1) = tSiyou.HWFONMIN         ' 品ＷＦ酸素濃度下限
    dShiyo(2) = tSiyou.HWFONMAX         ' 品ＷＦ酸素濃度上限
    dShiyo(3) = tSiyou.HWFONMBP         ' 品ＷＦ酸素濃度面内分布
    dShiyo(4) = tSiyou.HWFONAMN         ' 品ＷＦ酸素濃度平均下限
    dShiyo(5) = tSiyou.HWFONAMX         ' 品ＷＦ酸素濃度平均上限
    If fncJissekiHantei_nl(tSiyou.HWFONHWS, dShiyo) = False Then
        sErrMsg = sErrMsg & "(Oi)"
        funWfChkNull = FUNCTION_RETURN_FAILURE
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00131"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        Exit Function
    End If
    
    '--------------- BMD1,BMD2,BMD3 ---------------
    'BMDの使用NULLチェックを削除（判定でOKとする。）　          2003/12/19 tuku
''''    For cnt = 1 To 3
''''        ReDim dShiyo(1)
''''        If cnt = 1 Then         'BMD1
''''            sHosyo = tSiyou.HWFBM1HS            ' 品ＷＦＢＭＤ１保証方法＿処
''''            dShiyo(1) = tSiyou.HWFBM1MBP        ' 品ＷＦＢＭＤ１面内分布
''''        ElseIf cnt = 2 Then     'BMD2
''''            sHosyo = tSiyou.HWFBM2HS            ' 品ＷＦＢＭＤ２保証方法＿処
''''            dShiyo(1) = tSiyou.HWFBM2MBP        ' 品ＷＦＢＭＤ２面内分布
''''        ElseIf cnt = 3 Then     'BMD3
''''            sHosyo = tSiyou.HWFBM3HS            ' 品ＷＦＢＭＤ３保証方法＿処
''''            dShiyo(1) = tSiyou.HWFBM3MBP        ' 品ＷＦＢＭＤ３面内分布
''''        End If
''''        If fncJissekiHantei_nl(sHosyo, dShiyo) = False Then
''''            sErrMsg = sErrMsg & "(BMD" & cnt & ")"
''''            funWfChkNull = FUNCTION_RETURN_FAILURE
''''            Exit Function
''''        End If
''''    Next cnt
    
    '--------------- OSF1,OSF2,OSF3,OSF4 ---------------
    'チェックなし
    
    '--------------- DSOD ---------------
    ReDim dShiyo(4)
    dShiyo(1) = tSiyou.HWFDSOMX         ' 品ＷＦＤＳＯＤ上限
    dShiyo(2) = tSiyou.HWFDSOMN         ' 品ＷＦＤＳＯＤ下限
    dShiyo(3) = tSiyou.HWFDSOAX         ' 品ＷＦＤＳＯＤ領域上限
    dShiyo(4) = tSiyou.HWFDSOAN         ' 品ＷＦＤＳＯＤ領域下限
    If fncJissekiHantei_nl(tSiyou.HWFDSOHS, dShiyo) = False Then
        sErrMsg = sErrMsg & "(DSOD)"
        funWfChkNull = FUNCTION_RETURN_FAILURE
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00143"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        Exit Function
    End If
    
    '--------------- DZ幅 ---------------
    'チェックなし
        
    '--------------- SPVFE ---------------
    'チェックなし
        
    '--------------- DOI1(酸素析出1),DOI2(酸素析出2),DOI3(酸素析出3) ---------------
    'チェックなし
        
    '--------------- AOI(残存酸素) ---------------
    'チェックなし

    '--------------- GD ---------------         'GD追加　05/01/27 ooba
    'チェックなし
'    ReDim dShiyo(2)     'Den
'    dShiyo(1) = tSiyou.HWFDENMX         ' 品ＷＦＤｅｎ上限
'    dShiyo(2) = tSiyou.HWFDENMN         ' 品ＷＦＤｅｎ下限
'    If fncJissekiHantei_nl(tSiyou.HWFDENHS, dShiyo) = False Then
'        sErrMsg = sErrMsg & "(GD_Den)"
'        funWfChkNull = FUNCTION_RETURN_FAILURE
'        Exit Function
'    End If
'
'    ReDim dShiyo(2)     'DVD2
'    dShiyo(1) = tSiyou.HWFDVDMXN        ' 品ＷＦＤＶＤ２上限
'    dShiyo(2) = tSiyou.HWFDVDMNN        ' 品ＷＦＤＶＤ２下限
'    If fncJissekiHantei_nl(tSiyou.HWFDVDHS, dShiyo) = False Then
'        sErrMsg = sErrMsg & "(GD_DVD2)"
'        funWfChkNull = FUNCTION_RETURN_FAILURE
'        Exit Function
'    End If
'
'    ReDim dShiyo(2)     'L/DL
'    dShiyo(1) = tSiyou.HWFLDLMX         ' 品ＷＦＬ／ＤＬ上限
'    dShiyo(2) = tSiyou.HWFLDLMN         ' 品ＷＦＬ／ＤＬ下限
'    If fncJissekiHantei_nl(tSiyou.HWFLDLHS, dShiyo) = False Then
'        sErrMsg = sErrMsg & "(GD_LDL)"
'        funWfChkNull = FUNCTION_RETURN_FAILURE
'        Exit Function
'    End If
    
End Function

''Upd start 2005/06/22 (TCS)t.terauchi  SPV9点対応
'概要      :SPV判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :仕様情報構造体
'          :typ_j016      ,I  ,typ_TBCMJ016                         :SPV実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(0:判定OK, 1:判定NG)
'          :iKubun        ,I  ,Integer                              :区分(1:Fe濃度, 2:拡散長, 3:Nr濃度)
'          :sSxlPos       ,I  ,String                               :SXL位置(TOP/BOT)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :SPV判定を行う
'履歴      :2005/06/22 新規作成　(TCS)t.terauchi
Public Function WfCrSpvJudg_New(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            typ_j016 As typ_TBCMJ016, _
                            bJudg As Boolean, iKubun As Integer, sSxlPos As String) As Boolean
    Dim ErrInfo         As ERROR_INFOMATION         'エラー情報構造体
    Dim sp              As W_SPV                    'SPV構造体
    Dim sSokutei_Fe     As String                   'Fe濃度　測定方法
    Dim sSokutei_Diff   As String                   '拡散長　測定方法
'↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
    Dim sSokutei_Nr     As String                   'Nr濃度　測定方法
'↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------

    bJudg = True

    '測定方法取得
    sSokutei_Fe = typ_si.HWFSPVSH & typ_si.HWFSPVST & typ_si.HWFSPVSI
    sSokutei_Diff = typ_si.HWFDLSPH & typ_si.HWFDLSPT & typ_si.HWFDLSPI
'↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
    sSokutei_Nr = typ_si.HWFNRSH & typ_si.HWFNRST & typ_si.HWFNRSI
'↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------

    'Fe濃度と拡散長の測定方法が違う場合
    If sSokutei_Fe <> sSokutei_Diff Then
        
        'Fe濃度・拡散長共に、仕様有りの場合判定Errとする
        If ((typ_si.HWFDLHWS = "H") And CheckKHN(typ_si.HWFDLKHN, 16, sSxlPos)) _
            And ((typ_si.HWFSPVHS = "H") And CheckKHN(typ_si.HWFSPVKN, 15, sSxlPos)) Then
        
            bJudg = False
            WfCrSpvJudg_New = False
            Exit Function
        End If
    End If

    'SPV判定引数設定
    sp.GuaranteeSpvFe.cMeth = typ_si.HWFSPVSH   '品WFSPVFE測定位置＿方
    sp.GuaranteeSpvFe.cCount = typ_si.HWFSPVST  '品WFSPVFE測定位置＿点
    sp.GuaranteeSpvFe.cPos = typ_si.HWFSPVSI    '品WFSPVFE測定位置＿位
    sp.GuaranteeSpvFe.cObj = typ_si.HWFSPVHT    '品WFSPVFE保証方法＿対
    sp.GuaranteeSpvFe.cJudg = typ_si.HWFSPVHS   '品WFSPVFE保証方法＿処
    sp.GuaranteeSpv.cMeth = typ_si.HWFDLSPH     '品WF拡散長測定位置＿方
    sp.GuaranteeSpv.cCount = typ_si.HWFDLSPT    '品WF拡散長測定位置＿点
    sp.GuaranteeSpv.cPos = typ_si.HWFDLSPI      '品WF拡散長測定位置＿位
    sp.GuaranteeSpv.cObj = typ_si.HWFDLHWT      '品WF拡散長保証方法＿対
    sp.GuaranteeSpv.cJudg = typ_si.HWFDLHWS     '品WF拡散長保証方法＿処
    
    
'↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
    '判定項目にNr濃度を追加
    sp.GuaranteeSpvNr.cMeth = typ_si.HWFNRSH    '品WFSPVNR測定位置＿方
    sp.GuaranteeSpvNr.cCount = typ_si.HWFNRST   '品WFSPVNR測定位置＿点
    sp.GuaranteeSpvNr.cPos = typ_si.HWFNRSI     '品WFSPVNR測定位置＿位
    sp.GuaranteeSpvNr.cObj = typ_si.HWFNRHT     '品WFSPVNR保証方法＿対
    sp.GuaranteeSpvNr.cJudg = typ_si.HWFNRHS    '品WFSPVNR保証方法＿処
    sp.SpecSpvNrMax = typ_si.HWFNRMX            '品WFNR濃度上限
    sp.SpecSpvNrAvMax = typ_si.HWFNRAM          '品WFNR平均上限
'↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------

    sp.SpecSpvFeMax = typ_si.HWFSPVMX           '品WFFe濃度上限
    sp.SpecSpvAvMax = typ_si.HWFSPVAM           '品WF平均上限
    sp.SpecSpvMin = typ_si.HWFDLMIN             '品WF拡散長下限
    sp.SpecSpvMax = typ_si.HWFDLMAX             '品WF拡散長上限

    'Fe濃度判定
    If iKubun = "1" Then
        sp.Spv(0) = typ_j016.MAX_FE                                 'Fe濃度－MAX
        sp.Spv(1) = typ_j016.MIN_FE                                 'Fe濃度－MIN
        sp.Spv(2) = Format(typ_j016.AVE_FE, "0.00")                 'Fe濃度－AVE
        sp.Spv(3) = typ_j016.CENTER_FE                              'Fe濃度－センター
    
        If sSokutei_Fe = "AMX" Then
            'SPV(Fe濃度 MAP測定)判定
            If WfSPV_Fe_AMXJudg(sp, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
                bJudg = False
                WfCrSpvJudg_New = False
                Exit Function
            End If
'↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
            If sp.GuaranteeSpvFe.cJudg = JudgCodeW01 Then ''SPVFE濃度　判定有り
                ' SPV_Fe PUA値がFe濃度PUA限以下
                If sp.JudgSpv = JUDG_OK Then
                    sp.JudgSpv = RangeDecision_nl(typ_j016.PUA_FE, -1, typ_si.HWFSPVPUG)
                End If
                ' SPV_Fe PUA%値がFe濃度PUA率
                If sp.JudgSpv = JUDG_OK Then
                    sp.JudgSpv = RangeDecision_nl(typ_j016.PUAP_FE, -1, typ_si.HWFSPVPUR)
                End If
                ' SPV_Fe STD値がFe濃度標準偏差以下
                If sp.JudgSpv = JUDG_OK Then
                    sp.JudgSpv = RangeDecision_nl(typ_j016.STD_FE, -1, typ_si.HWFSPVSTD)
                End If
            End If
'↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
        ElseIf sSokutei_Fe = "V9T" Then
            'SPV(Fe濃度 9点測定)判定
            If WfSPV_Fe_V9TJudg(sp, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
                bJudg = False
                WfCrSpvJudg_New = False
                Exit Function
            End If
        Else
            bJudg = False
            WfCrSpvJudg_New = False
            Exit Function
        End If
    
    '拡散長判定
    ElseIf iKubun = "2" Then
        sp.Spv(0) = typ_j016.MAX_DIFF                               '拡散長－MAX
        sp.Spv(1) = typ_j016.MIN_DIFF                               '拡散長－MIN
        sp.Spv(2) = Format(typ_j016.AVE_DIFF, "0.0")                '拡散長－AVE
        sp.Spv(3) = typ_j016.CENTER_DIFF                            '拡散長－センター
    
        If sSokutei_Diff = "AMX" Then
            'SPV(拡散長 MAP測定)判定
            If WfSPV_DIFF_AMXJudg(sp, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
                bJudg = False
                WfCrSpvJudg_New = False
                Exit Function
            End If
'↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
            If sp.GuaranteeSpv.cJudg = JudgCodeW01 Then ''SPV拡散長　判定有り
                ' SPV_拡散長PUA値が拡散長PUA限以上
                If sp.JudgSpv = JUDG_OK Then
                    sp.JudgSpv = RangeDecision_nl(typ_j016.PUA_DIFF, typ_si.HWFDLPUG, -1)
                End If
                ' SPV_拡散長PUA%値が拡散長PUA率以上
                If sp.JudgSpv = JUDG_OK Then
                    sp.JudgSpv = RangeDecision_nl(typ_j016.PUAP_DIFF, typ_si.HWFDLPUR, -1)
                End If
            End If
'↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
        ElseIf sSokutei_Diff = "V9T" Then
            'SPV(拡散長 9点測定)判定
            If WfSPV_DIFF_V9TJudg(sp, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
                bJudg = False
                WfCrSpvJudg_New = False
                Exit Function
            End If
        Else
            bJudg = False
            WfCrSpvJudg_New = False
            Exit Function
        End If
'↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
    'Nr判定
    ElseIf iKubun = "3" Then

        sp.Spv(0) = typ_j016.SPV_Nr_MAX                             'Nr濃度－MAX
        sp.Spv(2) = Format(typ_j016.SPV_Nr_AVE, "0.00")             'Nr濃度－AVE
    
        If sSokutei_Nr = "AMX" Then
            'SPV(Fe濃度 MAP測定)判定
            If WfSPV_Nr_AMXJudg(sp, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
                bJudg = False
                WfCrSpvJudg_New = False
                Exit Function
            End If
            If sp.GuaranteeSpvNr.cJudg = JudgCodeW01 Then ''SPVNR濃度　判定有り
                ' SPV_Nr PUA値がNr濃度PUA限以下
                If sp.JudgSpv = JUDG_OK Then
                    sp.JudgSpv = RangeDecision_nl(typ_j016.PUA_NR, -1, typ_si.HWFNRPUG)
                End If
                ' SPV_Nr PUA%値がNr濃度PUA率
                If sp.JudgSpv = JUDG_OK Then
                    sp.JudgSpv = RangeDecision_nl(typ_j016.PUAP_NR, -1, typ_si.HWFNRPUR)
                End If
                ' SPV_Nr STD値がNr濃度標準偏差以下
                If sp.JudgSpv = JUDG_OK Then
                    sp.JudgSpv = RangeDecision_nl(typ_j016.STD_NR, -1, typ_si.HWFNRSTD)
                End If
            End If
        ElseIf sSokutei_Fe = "V9T" Then
            'SPV(Fe濃度 9点測定)判定
            If WfSPV_Nr_V9TJudg(sp, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
                bJudg = False
                WfCrSpvJudg_New = False
                Exit Function
            End If
        Else
            bJudg = False
            WfCrSpvJudg_New = False
            Exit Function
        End If
'↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
    Else
        bJudg = False
        WfCrSpvJudg_New = False
        Exit Function
    End If

    If sp.JudgSpv <> True Then
        bJudg = False
    End If

    WfCrSpvJudg_New = True

End Function
'Upd end   2005/06/21 (TCS)t.terauchi  SPV9点対応

'概要      :Warp判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                       :説明
'          :dWarpMax      ,I  ,Double                   :Warp上限
'          :dMeas         ,I  ,Double                   :測定値
'          :戻り値        ,O  ,Boolean                  :True→判定OK,False→判定NG
'説明      :
'履歴      :05/12/16 ooba
Public Function WfWarpJudg(dWarpMax As Double, dMeas As Double) As Boolean

    WfWarpJudg = True
    
    If dMeas = -1 Then Exit Function
    
    '仕様値が0orNULLの場合は判定OKとする
    If dWarpMax = 0 Then dWarpMax = -1
    
    'Warp判定(測定値≦上限値なら判定OK)
    WfWarpJudg = RangeDecision_nl(dMeas, -1, dWarpMax)
        
End Function

'概要      :合成角度判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                       :説明
'          :dKakuMin      ,I  ,Double                   :結晶面傾下限
'          :dKakuMax      ,I  ,Double                   :結晶面傾上限
'          :dMeas         ,I  ,Double                   :測定値
'          :戻り値        ,O  ,Boolean                  :True→判定OK,False→判定NG
'説明      :
'履歴      :05/12/16 ooba
Public Function WfKakuJudg(dKakuMin As Double, dKakuMax As Double, dMeas As Double) As Boolean

    WfKakuJudg = True
    
    If dMeas = -1 Then Exit Function
    
    '仕様値が0orNULLの場合は判定OKとする
    If dKakuMin = 0 Then dKakuMin = -1
    If dKakuMax = 0 Then dKakuMax = -1
    
    '合成角度判定(下限値≦測定値≦上限値なら判定OK)
    WfKakuJudg = RangeDecision_nl(dMeas, dKakuMin, dKakuMax)
        
End Function

'概要      :WF仕様構造体クリア関数
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                       :説明
'説明      :既存のプロシージャに記述するとVBの制限に引っかかるので、別プロシージャで作成する。
'履歴      :06/06/12 新規作成
Public Sub Crear_type_Siyou_Spv()
    'typ_Ctypeを初期化
    Dim clear_typeC(0) As typ_AllTypesC
    typ_CType = clear_typeC(0)
'Add Start 2011/03/09 SMPK Miyata
    ReDim typ_CType.dblScut(SXL_MAXSMP)             ' 再カット位置
    ReDim typ_CType.bOKNG(SXL_MAXSMP)               ' 比抵抗判定
    ReDim typ_CType.COEF(SXL_MAXSMP)                ' 偏析係数
    ReDim typ_CType.JudgRes(SXL_MAXSMP)             ' 比抵抗判定
    ReDim typ_CType.JudgRrg(SXL_MAXSMP)             ' RRG判定
    ReDim typ_CType.typ_y013(SXL_MAXSMP, MAXCNT)    ' 測定結果
    ReDim typ_CType.typ_hage(SXL_MAXSMP)            ' 引上げ終了実績
    ReDim typ_CType.typ_rslt(SXL_MAXSMP, MAXCNT)    ' 各実績情報
'Add End   2011/03/09 SMPK Miyata

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    'typ_Ctype_EPを初期化
    Dim clear_typeC_EP(0) As typ_AllTypesC_EP
    typ_CType_EP = clear_typeC_EP(0)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
'    Dim clear_typ_si_Spv As type_DBDRV_scmzc_fcmlc001c_Siyou_Spv
'    typ_si_Spv = clear_typ_si_Spv
End Sub

'概要      :BMD(EP)判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :仕様情報構造体
'          :typ_y022      ,I  ,typ_TBCMY022                         :BMD実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(0:判定OK, 1:判定NG)
'          :bmflg         ,I  ,Integer                              :BMDﾌﾗｸﾞ(1:BMD1, 2:BMD2, 3:BMD3)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :BMD判定を行う(関数WfCrBmdJudgを基に作成)
'履歴      :新規作成 2006/08/15 エピ先行評価追加対応 SMP)kondoh
Public Function EpBmdJudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            typ_y022 As typ_TBCMY022, _
                            bJudg As Boolean, _
                            bmflg As Integer) As Boolean
    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim bm      As W_BMD                    'BMD構造体
    Dim c0      As Integer

    Dim keisu As Double
    Const keisu1 As Double = 10000
    Const keisu2 As Double = 10000
    Const keisu3 As Double = 10000
    Const keisu4 As Double = 10000
    Const keisu5 As Double = 10000
    Const keisu6 As Double = 333000
    Const keisu7 As Double = 10000
    Const keisu8 As Double = 10000
    Const keisu9 As Double = 10000 'Add 2012/07/20 Y.Hitomi
    
    bJudg = True

    'BMD(EP)判定引数設定
    Select Case bmflg
    Case 1
        bm.GuaranteeBmd.cMeth = typ_si.HEPBM1SH   '品EPBMD1測定位置_方
        bm.GuaranteeBmd.cCount = typ_si.HEPBM1ST  '品EPBMD1測定位置_点
        bm.GuaranteeBmd.cPos = typ_si.HEPBM1SR    '品EPBMD1測定位置_領
        bm.GuaranteeBmd.cObj = typ_si.HEPBM1HT    '品EPBMD1保証方法_対
        bm.GuaranteeBmd.cJudg = typ_si.HEPBM1HS   '品EPBMD1保証方法_処
        bm.SpecBmdAveMin = typ_si.HEPBM1AN        '品EPBMD1平均下限
        bm.SpecBmdAveMax = typ_si.HEPBM1AX        '品EPBMD1平均上限
        bm.SpecBmdMBP = typ_si.HEPBM1MBP          '品EPBMD1面内分布
        bm.SpecBmdMCL = NtoS(typ_si.HEPBM1MCL)    '品EPBMD1面内計算
        bm.Antnp = typ_si.HEPANTNP                '品EPAN温度
    Case 2
        bm.GuaranteeBmd.cMeth = typ_si.HEPBM2SH   '品EPBMD2測定位置_方
        bm.GuaranteeBmd.cCount = typ_si.HEPBM2ST  '品EPBMD2測定位置_点
        bm.GuaranteeBmd.cPos = typ_si.HEPBM2SR    '品EPBMD2測定位置_領
        bm.GuaranteeBmd.cObj = typ_si.HEPBM2HT    '品EPBMD2保証方法_対
        bm.GuaranteeBmd.cJudg = typ_si.HEPBM2HS   '品EPBMD2保証方法_処
        bm.SpecBmdAveMin = typ_si.HEPBM2AN        '品EPBMD2平均下限
        bm.SpecBmdAveMax = typ_si.HEPBM2AX        '品EPBMD2平均上限
        bm.SpecBmdMBP = typ_si.HEPBM2MBP          '品EPBMD2面内分布
        bm.SpecBmdMCL = NtoS(typ_si.HEPBM2MCL)    '品EPBMD2面内計算
        bm.Antnp = typ_si.HEPANTNP                '品EPAN温度
    Case 3
        bm.GuaranteeBmd.cMeth = typ_si.HEPBM3SH   '品EPBMD3測定位置_方
        bm.GuaranteeBmd.cCount = typ_si.HEPBM3ST  '品EPBMD3測定位置_点
        bm.GuaranteeBmd.cPos = typ_si.HEPBM3SR    '品EPBMD3測定位置_領
        bm.GuaranteeBmd.cObj = typ_si.HEPBM3HT    '品EPBMD3保証方法_対
        bm.GuaranteeBmd.cJudg = typ_si.HEPBM3HS   '品EPBMD3保証方法_処
        bm.SpecBmdAveMin = typ_si.HEPBM3AN        '品EPBMD3平均下限
        bm.SpecBmdAveMax = typ_si.HEPBM3AX        '品EPBMD3平均上限
        bm.SpecBmdGsAveMin = typ_si.HEPBM3GSAN    '品EPBMD3平均下限(外周)　09/05/07 ooba
        bm.SpecBmdGsAveMax = typ_si.HEPBM3GSAX    '品EPBMD3平均上限(外周)　09/05/07 ooba
        bm.SpecBmdMBP = typ_si.HEPBM3MBP          '品EPBMD3面内分布
        bm.SpecBmdMCL = NtoS(typ_si.HEPBM3MCL)    '品EPBMD3面内計算
        bm.Antnp = typ_si.HEPANTNP                '品EPAN温度
    End Select
    
    If bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "4" And bm.GuaranteeBmd.cPos = "H" Then
        keisu = keisu1
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "3" And bm.GuaranteeBmd.cPos = "H" Then
        keisu = keisu2
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "4" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu3
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "3" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu4
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "5" And bm.GuaranteeBmd.cPos = "A" Then
        keisu = keisu5
    ElseIf bm.GuaranteeBmd.cMeth = "G" And bm.GuaranteeBmd.cCount = "3" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu6
    ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "5" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu7
    ElseIf bm.GuaranteeBmd.cMeth = "8" And bm.GuaranteeBmd.cCount = "4" And bm.GuaranteeBmd.cPos = "8" Then
        keisu = keisu8
    'Add Start 2012/07/20 Y.Hitomi
    ElseIf bm.GuaranteeBmd.cMeth = "P" Then
        keisu = keisu9
    'Add End 2012/07/20 Y.Hitomi
    Else
        bJudg = False
        EpBmdJudg = False
        Exit Function
    End If
    
    With bm
        .BMD(0) = NtoZ2(typ_y022.MESDATA1)                   'BMD測定値
        .BMD(1) = NtoZ2(typ_y022.MESDATA2)                   'BMD測定値
        .BMD(2) = NtoZ2(typ_y022.MESDATA3)                   'BMD測定値
        .BMD(3) = NtoZ2(typ_y022.MESDATA4)                   'BMD測定値
        .BMD(4) = NtoZ2(typ_y022.MESDATA5)                   'BMD測定値
        .BmdAntnp = NtoZ2(Mid(typ_y022.DKAN, 3, 4))
        For c0 = 0 To 4
            .BMD(c0) = IIf(.BMD(c0) <> -1, .BMD(c0) * CDbl(keisu / 10000), -1)
        Next
    End With

    'BMD判定
'    If WfBMDJudg(bm, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
    If WfBMDJudg(bm, ErrInfo, bmflg) <> FUNCTION_RETURN_SUCCESS Then    'BMDno追加　09/05/07 ooba
        bJudg = False
        EpBmdJudg = False
        Exit Function
    End If
    
    If bm.JudgBmd <> True Or bm.JudgAntnp <> True Then
        bJudg = False
    End If
    
    typ_y022.MESDATA6 = bm.JudgDataAve
    typ_y022.MESDATA7 = bm.JudgDataMax
    typ_y022.MESDATA8 = bm.JudgDataMin
    typ_y022.MESDATA9 = bm.JudgDataMBP
    
'' 2008/10/20 BMD評価,外周1点保証機能追加 ADD By Systech Start
    If bm.GuaranteeBmd.cObj = ObjCode18 Then
        If bm.BMD(0) <> -1 Then
            typ_y022.MESDATA9 = bm.BMD(0)
        End If
    End If
'' 2008/10/20 BMD評価,外周1点保証機能追加 ADD By Systech End
     
    EpBmdJudg = True
End Function

'概要      :OSF(EP)判定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :typ_si        ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :仕様情報構造体
'          :typ_y022      ,I  ,typ_TBCMY022                         :OSF実績構造体
'          :bJudg         ,O  ,Boolean                              :判定結果(0:判定OK, 1:判定NG)
'          :osfflg        ,I  ,Integer                              :OSFﾌﾗｸﾞ(1:OSF1, 2:OSF2, 3:OSF3)
'          :戻り値        ,O  ,Boolean                              :True:正常終了, False:異常終了
'説明      :OSF判定を行う(関数WfCrOsfJudgを基に作成)
'履歴      :新規作成 2006/08/15 エピ先行評価追加対応 SMP)kondoh
Public Function EpOsfJudg(typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                            typ_y022 As typ_TBCMY022, _
                            bJudg As Boolean, _
                            osfflg As Integer, _
                            TmpData() As String) As Boolean

    Dim ErrInfo As ERROR_INFOMATION         'エラー情報構造体
    Dim os      As W_OSF                    'OSF構造体
    Dim keisu   As Double
    Dim c0      As Integer
    
    Const keisu1 As Double = 1.8248175
    Const keisu2 As Double = 1.8518519
    Const keisu3 As Double = 1.9230769
    Const keisu4 As Double = 3.649635
    Const keisu5 As Double = 3.7037037
    Const keisu6 As Double = 3.8461538
    Const keisu7 As Double = 7.6923077
        
    bJudg = True

    'OSF(EP)判定引数設定
    Select Case osfflg
    Case 1
        os.GuaranteeOsf.cMeth = typ_si.HEPOF1SH     '品EPOSF1測定位置_方
        os.GuaranteeOsf.cCount = typ_si.HEPOF1ST    '品EPOSF1測定位置_点
        os.GuaranteeOsf.cPos = typ_si.HEPOF1SR      '品EPOSF1測定位置_領
        os.GuaranteeOsf.cObj = typ_si.HEPOF1HT      '品EPBMD1保証方法_対
        os.GuaranteeOsf.cJudg = typ_si.HEPOF1HS     '品EPBMD1保証方法_処
        os.SpecOsfAveMax = typ_si.HEPOF1AX          '品EPOSF1平均上限
        os.SpecOsfMax = typ_si.HEPOF1MX             '品EPOSF1上限
        os.JudgDataPTK = NtoS(typ_si.HEPOSF1PTK)    '品EPOSF1ﾊﾟﾀﾝ区分
        os.Antnp = typ_si.HEPANTNP                  '品EPAN温度
    Case 2
        os.GuaranteeOsf.cMeth = typ_si.HEPOF2SH     '品EPOSF2測定位置_方
        os.GuaranteeOsf.cCount = typ_si.HEPOF2ST    '品EPOSF2測定位置_点
        os.GuaranteeOsf.cPos = typ_si.HEPOF2SR      '品EPOSF2測定位置_領
        os.GuaranteeOsf.cObj = typ_si.HEPOF2HT      '品EPBMD2保証方法_対
        os.GuaranteeOsf.cJudg = typ_si.HEPOF2HS     '品EPBMD2保証方法_処
        os.SpecOsfAveMax = typ_si.HEPOF2AX          '品EPOSF2平均上限
        os.SpecOsfMax = typ_si.HEPOF2MX             '品EPOSF2上限
        os.JudgDataPTK = NtoS(typ_si.HEPOSF2PTK)    '品EPOSF2ﾊﾟﾀﾝ区分
        os.Antnp = typ_si.HEPANTNP                  '品EPAN温度
    Case 3
        os.GuaranteeOsf.cMeth = typ_si.HEPOF3SH     '品EPOSF3測定位置_方
        os.GuaranteeOsf.cCount = typ_si.HEPOF3ST    '品EPOSF3測定位置_点
        os.GuaranteeOsf.cPos = typ_si.HEPOF3SR      '品EPOSF3測定位置_領
        os.GuaranteeOsf.cObj = typ_si.HEPOF3HT      '品EPBMD3保証方法_対
        os.GuaranteeOsf.cJudg = typ_si.HEPOF3HS     '品EPBMD3保証方法_処
        os.SpecOsfAveMax = typ_si.HEPOF3AX          '品EPOSF3平均上限
        os.SpecOsfMax = typ_si.HEPOF3MX             '品EPOSF3上限
        os.JudgDataPTK = NtoS(typ_si.HEPOSF3PTK)    '品EPOSF3ﾊﾟﾀﾝ区分
        os.Antnp = typ_si.HEPANTNP                  '品EPAN温度
    End Select
    
    If os.GuaranteeOsf.cMeth = "5" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "3" Then
        keisu = keisu1
    ElseIf os.GuaranteeOsf.cMeth = "5" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "5" Then
        keisu = keisu2
    ElseIf os.GuaranteeOsf.cMeth = "5" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "A" Then
        keisu = keisu3
    ElseIf os.GuaranteeOsf.cMeth = "6" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "3" Then
        keisu = keisu4
    ElseIf os.GuaranteeOsf.cMeth = "6" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "5" Then
        keisu = keisu5
    ElseIf os.GuaranteeOsf.cMeth = "6" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "A" Then
        keisu = keisu6
    ElseIf os.GuaranteeOsf.cMeth = "E" And os.GuaranteeOsf.cCount = "5" And os.GuaranteeOsf.cPos = "A" Then
        keisu = keisu7
    Else
        bJudg = False
        EpOsfJudg = False
        Exit Function
    End If

    With os
        .OSF(0) = NtoZ2(typ_y022.MESDATA1)                   'OSF測定値
        .OSF(1) = NtoZ2(typ_y022.MESDATA2)                   'OSF測定値
        .OSF(2) = NtoZ2(typ_y022.MESDATA3)                   'OSF測定値
        .OSF(3) = NtoZ2(typ_y022.MESDATA4)                   'OSF測定値
        .OSF(4) = NtoZ2(typ_y022.MESDATA5)                   'OSF測定値
        .OsfAntnp = NtoZ2(Mid(typ_y022.DKAN, 3, 4))
        For c0 = 0 To 4
            .OSF(c0) = IIf(.OSF(c0) <> -1, .OSF(c0) * keisu, -1)
        Next
        typ_y022.MESDATA6 = typ_y022.MESDATA6 * 100
        .OSFp(0) = Trim(typ_y022.MESDATA9)                   'OSFパターン実績(大)
        .OSFp(1) = Trim(typ_y022.MESDATA12)                  'OSFパターン実績(中)
        .OSFp(2) = Trim(typ_y022.MESDATA15)                  'OSFパターン実績(小)
    End With
    typ_y022.MESDATA6 = typ_y022.MESDATA6 * 100
    
    'OSF判定
    If WfOSFJudg(os, ErrInfo) <> FUNCTION_RETURN_SUCCESS Then
        bJudg = False
        EpOsfJudg = False
        Exit Function
    End If
    
    If os.JudgOsf <> True Or os.JudgAntnp <> True Then
        bJudg = False
    End If
    
    TmpData(0) = os.JudgDataAve
    TmpData(1) = os.JudgDataMax
     EpOsfJudg = True
End Function

'概要      :エピ実績値の総合判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型             :説明
'          :sKeyID          ,I  ,String         :SXL-ID
'          :tNew_Hinban     ,I  ,String         :振替候補品番
'          :bTotalJudg      ,O  ,Boolean        :トータル判定
'          :iErr_Code       ,O  ,Integer        :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String         :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :typ_CType       ,O  ,typ_AllTypesC  :全情報構造体(構造体)
'          :typ_CType_EP    ,O  ,typ_AllTypesC_EP  :全情報構造体(構造体)(エピ用)
'          :iSmpGetFlg      ,I  ,Integer        :ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ(0:ｻﾝﾌﾟﾙ指定なし, 1:ｻﾝﾌﾟﾙ指定あり)
'          :sSamplID1       ,I  ,String         :TOPｻﾝﾌﾟﾙID(省略可)
'          :sSamplID2       ,I  ,String         :BOTｻﾝﾌﾟﾙID(省略可)
'          :iKcnt           ,I  ,Integer        :工程連番(省略可)
'          :戻り値          ,O  ,Integer        :取得の成否(0:正常終了, -1:異常終了)
'説明      :
'履歴      :2006/08/15 Add エピ先行評価追加対応 SMP)kondoh

Public Function funWfcSogoHantei_EP(sKeyID As String, tNew_Hinban As tFullHinban, _
                bTotalJudg As Boolean, iErr_Code As Integer, sErr_Msg As String, typ_CType As typ_AllTypesC, typ_CType_EP As typ_AllTypesC_EP, _
                iSmpGetFlg As Integer, Optional sSamplID1 As String = vbNullString, Optional sSamplID2 As String = vbNullString, _
                Optional iKcnt As Integer = 0) As Integer

    Dim i       As Integer      'Add 2011/03/09 SMPK Miyata

    On Error GoTo Apl_down
    
    '戻り値初期化
    funWfcSogoHantei_EP = FUNCTION_RETURN_FAILURE
    
    'グローバル変数に設定
    wiSmpGetFlg = iSmpGetFlg
    wiKcnt = iKcnt
    
    '初期設定
    sErr_Msg = "WFC総合判定(SetInitData)"
    If SetInitData(sKeyID, tNew_Hinban, typ_CType.typ_Param, iSmpGetFlg, sSamplID1, sSamplID2) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
    
    '画面情報設定
    sErr_Msg = "WFC総合判定(エピ)(SetAllData_EP)"
    If SetAllData_EP(typ_CType, typ_CType_EP, tNew_Hinban, iSmpGetFlg, iErr_Code, sErr_Msg) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If
        
    TotalJudg = True
    MidlJudg = True             '中間抜試判定   Add 2011/03/09 SMPK Miyata

    '仕様検査支持取得
    sErr_Msg = "WFC総合判定(エピ)(SpecJudgCheck)"
    SpecJudgCheck

'''    '2003/12/13 SystemBrain Null対応追加▽
'''    '仕様Nullチェック
'''    sErr_Msg = "仕様Nullﾁｪｯｸ(エピ)"
'''    If funWfChkNull(typ_CType.typ_si, sErr_Msg) = FUNCTION_RETURN_FAILURE Then
'''        GoTo Apl_down
'''    End If
'''    '2003/12/13 SystemBrain Null対応追加△

    '実績データ判定(TOP)
    sErr_Msg = "WFC総合判定(エピ)(判定(TOP))"
    If EPJudge(typ_CType, typ_CType_EP, tNew_Hinban, SxlTop) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If

    '実績データ判定(TAIL)
    sErr_Msg = "WFC総合判定(エピ)(判定(TAIL))"
    If EPJudge(typ_CType, typ_CType_EP, tNew_Hinban, SxlTail) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If

'Add Start 2011/03/10 SMPK Miyata
    '実績データ判定(MIDLE)
    sErr_Msg = "WFC総合判定(エピ)(判定(MIDLE))"
    For i = SxlMidl To UBound(typ_CType.typ_Param.WFSMP)
        If EPJudge(typ_CType, typ_CType_EP, tNew_Hinban, i) = FUNCTION_RETURN_FAILURE Then
            GoTo Apl_down
        End If
    Next i
'Add End   2011/03/10 SMPK Miyata

'Chg Start 2011/03/09 SMPK Miyata
'    bTotalJudg = TotalJudg
    bTotalJudg = TotalJudg And MidlJudg
'Chg End   2011/03/09 SMPK Miyata

    funWfcSogoHantei_EP = FUNCTION_RETURN_SUCCESS
    
'------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funWfcSogoHantei_EP = -4
    iErr_Code = funWfcSogoHantei_EP
    GoTo Apl_Exit
    
End Function

'概要      :画面情報データ設定(エピ)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typ_CType     ,I  ,typ_AllTypesC ,各情報構造体
'          :typ_CType_EP  ,I  ,typ_AllTypesC_EP ,各情報構造体
'説明      :画面情報を情報構造体に設定する
'履歴      :
Private Function SetAllData_EP(typ_CType As typ_AllTypesC, typ_CType_EP As typ_AllTypesC_EP, tNew_Hinban As tFullHinban, iSmpGetFlg As Integer, _
                                                    iErr_Code As Integer, sErr_Msg As String) As FUNCTION_RETURN
    
    Dim typ_in      As type_DBDRV_scmzc_fcmlc001c_In     ' DBアクセス入力用
    Dim fret(2)     As FUNCTION_RETURN
    Dim typ_hi()    As typ_TBCMH004
    Dim sErrMsg     As String
    Dim FuncAns     As FUNCTION_RETURN
    Dim records()   As typ_TBCMH001
'Add Start 2011/03/07 SMPK Miyata
    Dim i           As Integer      'カウンタ
    Dim iMidNo      As Integer      '中間抜試No
'Add End   2011/03/07 SMPK Miyata

    SetAllData_EP = FUNCTION_RETURN_FAILURE
    
    typ_in.HIN.hinban = typ_CType.typ_Param.hinban
    typ_in.HIN.factory = typ_CType.typ_Param.factory
    typ_in.HIN.mnorevno = typ_CType.typ_Param.REVNUM
    typ_in.HIN.opecond = typ_CType.typ_Param.opecond
    typ_in.SXLID = typ_CType.typ_Param.SXLID

    With typ_CType_EP
        
        'TOP側
        sErr_Msg = "WFC総合判定(TOP 初期ﾃﾞｰﾀ設定)(エピ)"
        typ_in.SAMPLEID = typ_CType.typ_Param.WFSMP(SxlTop).REPSMPLIDCW
        typ_in.WFSMP = typ_CType.typ_Param.WFSMP(SxlTop)
    
        '評価測定結果取得
        ReDim .typ_y022top(0)
        
        sErr_Msg = "WFC総合判定(TOP funGet_TBCME050)"
        '' エピ仕様を取得
        If funGet_TBCME050(tNew_Hinban, typ_CType.typ_si, sErrMsg) <> FUNCTION_RETURN_SUCCESS Then
            'エピ仕様取得失敗
            SetAllData_EP = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
        
        sErr_Msg = "WFC総合判定(TOP funGetTBCMY022_All)"
        '' エピ測定評価結果(実績値)を取得(0件でもエラーではない)
        If funGetTBCMY022_All(typ_in, .typ_y022top()) = FUNCTION_RETURN_FAILURE Then
            sErrMsg = GetMsgStr("EGET2", "Y022")
            SetAllData_EP = FUNCTION_RETURN_FAILURE
            Exit Function
        End If

        ' 評価測定結果整列
        sErr_Msg = "WFC総合判定(エピ)(TOP 評価測定結果整列)"
        If SetMERInd_EP(typ_CType_EP, .typ_y022top(), SxlTop) <> True Then
            '評価測定結果整列失敗
            Exit Function
        End If
        '引上げ終了実績取得
        ReDim typ_hi(0)
'頭8を購入単結晶扱いしない 2007/10/10 SETsw kubota
'        If Mid(typ_CType.typ_Param.CRYNUM, 1, 1) <> "8" Then
            sErr_Msg = "WFC総合判定(エピ)(TOP 引上げ終了実績取得)"
            If s_cmmc001db_Sql(typ_CType.typ_Param.CRYNUM, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                '引上げ終了実績取得失敗
                Exit Function
            Else
                If UBound(typ_hi) <> 0 Then
                    typ_CType.typ_hage(SxlTop) = typ_hi(1)
                Else
                    '引上げ終了実績取得失敗
                    SetAllData_EP = FUNCTION_RETURN_FAILURE
                    Exit Function
                End If
            End If
'        End If
    
        'TAIL側
        sErr_Msg = "WFC総合判定(TAIL 初期ﾃﾞｰﾀ設定)(エピ)"
        typ_in.SAMPLEID = typ_CType.typ_Param.WFSMP(SxlTail).REPSMPLIDCW
        typ_in.WFSMP = typ_CType.typ_Param.WFSMP(SxlTail)
    
        ReDim .typ_y022tail(0)
        
        '' エピ測定評価結果(実績値)を取得(0件でもエラーではない)
        If funGetTBCMY022_All(typ_in, .typ_y022tail()) = FUNCTION_RETURN_FAILURE Then
            sErrMsg = GetMsgStr("EGET2", "Y022")
            SetAllData_EP = FUNCTION_RETURN_FAILURE
            Exit Function
        End If

        ' 評価測定結果整列
        sErr_Msg = "WFC総合判定(エピ)(TAIL 評価測定結果整列)"
        If SetMERInd_EP(typ_CType_EP, .typ_y022tail(), SxlTail) <> True Then
            '評価測定結果整列失敗
            Exit Function
        End If

        '引上げ終了実績取得
        ReDim typ_hi(0)
        
'頭8を購入単結晶扱いしない 2007/10/10 SETsw kubota
'        If Mid(typ_CType.typ_Param.CRYNUM, 1, 1) <> "8" Then
            sErr_Msg = "WFC総合判定(エピ)(TAIL 引上げ終了実績取得)"
            If s_cmmc001db_Sql(typ_CType.typ_Param.CRYNUM, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                '引上げ終了実績取得失敗
                Exit Function
            Else
                If UBound(typ_hi) <> 0 Then
                    typ_CType.typ_hage(SxlTail) = typ_hi(1)
                Else
                    '引上げ終了実績取得失敗
                    SetAllData_EP = FUNCTION_RETURN_FAILURE
                    Exit Function
                End If
            End If
'        End If

'Add Start 2011/03/10 SMPK Miyata
        For i = SxlMidl To UBound(typ_CType.typ_Param.WFSMP)

            iMidNo = i - SxlMidl + 1
            
            If iMidNo > SXL_MAXSMP Then
                ' 中間抜試最大件数オーバー
                Exit Function
            End If

            'MIDLE側
            sErr_Msg = "WFC総合判定(MIDLE_" & iMidNo & " 初期ﾃﾞｰﾀ設定)(エピ)"
            typ_in.SAMPLEID = typ_CType.typ_Param.WFSMP(i).REPSMPLIDCW
            typ_in.WFSMP = typ_CType.typ_Param.WFSMP(i)
    
            '評価測定結果取得
            ReDim Preserve .typ_y022midl_ary(iMidNo)
            
            sErr_Msg = "WFC総合判定(MIDLE_" & iMidNo & " funGet_TBCME050)"
            '' エピ仕様を取得
            If funGet_TBCME050(tNew_Hinban, typ_CType.typ_si, sErrMsg) <> FUNCTION_RETURN_SUCCESS Then
                'エピ仕様取得失敗
                SetAllData_EP = FUNCTION_RETURN_FAILURE
                Exit Function
            End If
        
            sErr_Msg = "WFC総合判定(MIDLE_" & iMidNo & " funGetTBCMY022_All)"

            '' エピ測定評価結果(実績値)を取得(0件でもエラーではない)
            If funGetTBCMY022_All(typ_in, .typ_y022midl_ary(iMidNo).typ_y022midl) = FUNCTION_RETURN_FAILURE Then
                sErrMsg = GetMsgStr("EGET2", "Y022")
                SetAllData_EP = FUNCTION_RETURN_FAILURE
                Exit Function
            End If

            ' 評価測定結果整列
            sErr_Msg = "WFC総合判定(エピ)(MIDLE_" & iMidNo & " 評価測定結果整列)"
            If SetMERInd_EP(typ_CType_EP, .typ_y022midl_ary(iMidNo).typ_y022midl, i) <> True Then
                '評価測定結果整列失敗
                Exit Function
            End If

            '引上げ終了実績取得
            ReDim typ_hi(0)
            
            sErr_Msg = "WFC総合判定(エピ)(MIDLE_" & iMidNo & " 引上げ終了実績取得)"
            If s_cmmc001db_Sql(typ_CType.typ_Param.CRYNUM, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                '引上げ終了実績取得失敗
                Exit Function
            Else
                If UBound(typ_hi) <> 0 Then
                    typ_CType.typ_hage(i) = typ_hi(1)
                Else
                    '引上げ終了実績取得失敗
                    SetAllData_EP = FUNCTION_RETURN_FAILURE
                    Exit Function
                End If
            End If
        Next i
'Add End   2011/03/10 SMPK Miyata
    End With
    
    '' Ｐ＋結晶の判断
    sErr_Msg = "WFC総合判定(P+結晶の判断)"
    If DBDRV_GetTBCMH001(records(), "where UPINDNO = '" & left(SelectSxlID, 7) & "0" & Mid(SelectSxlID, 9, 1) & "'") = FUNCTION_RETURN_SUCCESS Then
    Else
        SetAllData_EP = FUNCTION_RETURN_FAILURE
        Exit Function
    End If
    
    SetAllData_EP = FUNCTION_RETURN_SUCCESS
End Function

'概要      :測定評価結果のソート
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typ_a         ,IO ,typ_AllTypesC ,各情報構造体
'          :typ_y022()    ,I  ,typ_TBCMY022 ,測定評価結果情報構造体
'          :tt            ,I  ,Integer      ,TOP・TAIL
'          :戻り値        ,O  ,Integer      ,True:正常終了　False:異常終了
'説明      :測定評価結果配列にDB検索したレコードを整列する
'履歴      :SB_WfJudg.SetMERIndを基に作成
Private Function SetMERInd_EP(typ_CType_EP As typ_AllTypesC_EP, _
                          typ_y022() As typ_TBCMY022, _
                          tt As Integer) As Boolean
    Dim i As Integer
    
    With typ_CType_EP
        For i = 1 To UBound(typ_y022)
            Select Case Trim(typ_y022(i).Spec)
            Case OSEPBMD1 ' BMD1
                .typ_y022(tt, EPBMD1) = typ_y022(i)
            Case OSEPBMD2 ' BMD2
                .typ_y022(tt, EPBMD2) = typ_y022(i)
            Case OSEPBMD3 ' BMD3
                .typ_y022(tt, EPBMD3) = typ_y022(i)
            Case OSEPOSF1 ' OSF1
                .typ_y022(tt, EPOSF1) = typ_y022(i)
            Case OSEPOSF2 ' OSF2
                .typ_y022(tt, EPOSF2) = typ_y022(i)
            Case OSEPOSF3 ' OSF3
                .typ_y022(tt, EPOSF3) = typ_y022(i)
            Case OSEPOT2 ' OT2
                .typ_y022(tt, EPOT2) = typ_y022(i)
            End Select
        Next
    End With
    SetMERInd_EP = True
End Function

'概要      :結晶判定(エピ)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :typ_CType     ,I  ,typ_AllTypesC    ,各情報構造体
'          :typ_CType_EP  ,I  ,typ_AllTypesC_EP ,各情報構造体(エピ)
'          :tNew_Hinban   ,I  ,tFullHinban      :振替候補品番
'          :tt            ,I  ,Integer          ,TopTail判定用
'説明      :検査指示に従い、エピ実績判定を行う
'履歴      :
Public Function EPJudge(typ_CType As typ_AllTypesC, typ_CType_EP As typ_AllTypesC_EP, tNew_Hinban As tFullHinban, tt As Integer) As FUNCTION_RETURN
    
    Dim IND         As String * 4                  '検査指示
    Dim bJudg       As Boolean
    Dim i           As Integer
    Dim typTmList() As typ_TBCMB005
'Chg Start 2011/03/10 SMPK Miyata
'    Dim INGOTPOS(2) As Integer
    Dim INGOTPOS(SXL_MAXSMP) As Integer
'Chg End   2011/03/10 SMPK Miyata
    Dim vTemp       As Variant
    Dim sHinban12   As String                               '品番(12桁)
    Dim sSxlPos     As String       'SXL位置(TOP/BOT)　04/04/12 ooba

    i = 0
    EPJudge = FUNCTION_RETURN_FAILURE
    
    sHinban12 = tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond
    
    If tt = SxlTop Then
        INGOTPOS(tt) = typ_CType.typ_Param.INGOTPOS
'Chg Start 2011/03/10 SMPK Miyata
'    Else
'        INGOTPOS(tt) = typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH
    ElseIf tt = SxlTail Then
        INGOTPOS(tt) = typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH
    Else
        INGOTPOS(tt) = typ_CType.typ_Param.WFSMP(tt).INPOSCW
'Chg End   2011/03/10 SMPK Miyata
    End If
    
    '検査指示設定
    If tt = SxlTop Then
        IND = "123"
    Else
        IND = "123"
    End If
    
'Chg Start 2011/03/10 SMPK Miyata
'    sSxlPos = IIf(tt = SxlTop, "TOP", "BOT")        '04/04/12 ooba
    Select Case tt
        Case SxlTop:    sSxlPos = "TOP"
        Case SxlTail:   sSxlPos = "BOT"
        Case Else:      sSxlPos = "MID"
    End Select
'Chg End   2011/03/10 SMPK Miyata

    '結晶コードリスト取得
    If GetCodeList(MSYSCLASS, KCLASS, typTmList()) <> FUNCTION_RETURN_SUCCESS Then
        '結晶コードリスト取得失敗
        Exit Function
    End If
    
        '' 結晶検査指示(B1)*****************************************************************
        BMDDataSet_EP 1, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' 結晶検査指示(B2)*****************************************************************
        BMDDataSet_EP 2, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' 結晶検査指示(B3)*****************************************************************
        BMDDataSet_EP 3, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' 結晶検査指示(L1)*****************************************************************
        OSFDataSet_EP 1, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' 結晶検査指示(L2)*****************************************************************
        OSFDataSet_EP 2, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        '' 結晶検査指示(L3)*****************************************************************
        OSFDataSet_EP 3, tt, CStr(INGOTPOS(tt)), i, typTmList(), sHinban12
        
    EPJudge = FUNCTION_RETURN_SUCCESS
End Function

Public Sub BMDDataSet_EP(BmdNo As Integer, UpDo As Integer, INGOTPOS As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String)

    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4               '検査指示
    Dim typ_y022z       As typ_TBCMY022
    Dim bJudg           As Boolean
    Dim vTemp           As Variant
    Dim SijiUmu         As String
    Dim EPBmSokuP       As String                   ' 測定位置＿点
    Dim EPBMD           As Integer
    Dim sSxlPos         As String                   'SXL位置(TOP/BOT)

    '検査指示設定
    IND = IIf(UpDo = SxlTop, "123", "123")
    
'Chg Start 2011/03/10 SMPK Miyata
'    sSxlPos = IIf(UpDo = SxlTop, "TOP", "BOT")
    Select Case UpDo
        Case SxlTop:    sSxlPos = "TOP"
        Case SxlTail:   sSxlPos = "BOT"
        Case Else:      sSxlPos = "MID"
    End Select
'Chg End   2011/03/10 SMPK Miyata

    With typ_CType_EP
        
        Select Case BmdNo
        Case 1
            EPBMD = EPBMD1
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.B1E And CheckKHN_EP(typ_CType.typ_si.HEPBM1KN, 1, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.B1E And CheckKHN_EP(typ_CType.typ_si.HEPBM1KN, 1, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '保証方法=保証 かつ 中間抜試（保証）の場合、仕様有とする
                JudgSpecCode = (JudgSW.B1E And typ_CType.typ_si.MSMPFLGEPBM = "1" And _
                                (typ_CType.typ_si.MSMPFLG = "1" Or typ_CType.typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.B1E And typ_CType.typ_si.MSMPFLGEPBM = "1")
'Chg End　 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "B1E"
            shiji = (InStr(IND, typ_CType.typ_Param.WFSMP(UpDo).EPINDB1CW) <> 0)
            SijiUmu = typ_CType.typ_Param.WFSMP(UpDo).EPRESB1CW
            EPBmSokuP = typ_CType.typ_si.HEPBM1ST
        Case 2
            EPBMD = EPBMD2
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.B2E And CheckKHN_EP(typ_CType.typ_si.HEPBM2KN, 2, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.B2E And CheckKHN_EP(typ_CType.typ_si.HEPBM2KN, 2, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '保証方法=保証 かつ 中間抜試（保証）の場合、仕様有とする
                JudgSpecCode = (JudgSW.B2E And typ_CType.typ_si.MSMPFLGEPBM = "1" And _
                                (typ_CType.typ_si.MSMPFLG = "1" Or typ_CType.typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.B2E And typ_CType.typ_si.MSMPFLGEPBM = "1")
'Chg End　 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "B2E"
            shiji = (InStr(IND, typ_CType.typ_Param.WFSMP(UpDo).EPINDB2CW) <> 0)
            SijiUmu = typ_CType.typ_Param.WFSMP(UpDo).EPRESB2CW
            EPBmSokuP = typ_CType.typ_si.HEPBM2ST
        Case 3
            EPBMD = EPBMD3
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.B3E And CheckKHN_EP(typ_CType.typ_si.HEPBM3KN, 3, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.B3E And CheckKHN_EP(typ_CType.typ_si.HEPBM3KN, 3, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '保証方法=保証 かつ 中間抜試（保証）の場合、仕様有とする
                JudgSpecCode = (JudgSW.B3E And typ_CType.typ_si.MSMPFLGEPBM = "1" And _
                                (typ_CType.typ_si.MSMPFLG = "1" Or typ_CType.typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.B3E And typ_CType.typ_si.MSMPFLGEPBM = "1")
'Chg End　 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "B3E"
            shiji = (InStr(IND, typ_CType.typ_Param.WFSMP(UpDo).EPINDB3CW) <> 0)
            SijiUmu = typ_CType.typ_Param.WFSMP(UpDo).EPRESB3CW
            EPBmSokuP = typ_CType.typ_si.HEPBM3ST
        End Select
            typ_y022z = .typ_y022(UpDo, EPBMD)
    
    
        '' EP検査指示（BMDE)*****************************************************************
        If JudgSpecCode Then
            '画面表示内容初期設定
            .typ_rslt(UpDo, DispLineCount).pos = -1                                             ' 結晶内開始位置
            .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())              ' 内容
            .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                                     ' 情報1
            .typ_rslt(UpDo, DispLineCount).INFO2 = "検査無"                                     ' 情報2
            .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                           ' 情報3
            .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                           ' 情報4
            .typ_rslt(UpDo, DispLineCount).INFO5 = ""                                           ' 情報5
            .typ_rslt(UpDo, DispLineCount).INFO6 = ""                                           ' 情報6
            .typ_rslt(UpDo, DispLineCount).INFO7 = ""                                           ' 情報7
            .typ_rslt(UpDo, DispLineCount).INFO8 = ""                                           ' 情報8
            .typ_rslt(UpDo, DispLineCount).SMPLID = -1                                          ' サンプルＮｏ
            .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                          ' 判定結果
            .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                                   ' 品番(12桁)
            bJudg = False
            If shiji <> 0 Then
                '画面表示内容設定
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' 結晶内開始位置
                .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                                 ' 情報2
                .typ_rslt(UpDo, DispLineCount).INFO3 = "実績無"                                 ' 情報3
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y022z.SAMPLEID                      ' サンプルＮｏ
                    
                'BMDE判定
'                If SijiUmu = "1" Then
                If (SijiUmu = "1") And (Trim(typ_y022z.SAMPLEID) <> "0") Then           '2003/12/19 SystemBrain
                    'BMDE判定失敗
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                            ' 情報２
                    'BMDE判定
                    If EpBmdJudg(typ_CType.typ_si, typ_y022z, bJudg, BmdNo) Then
                        vTemp = CVar(typ_y022z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' 情報1
                        vTemp = CVar(typ_y022z.MESDATA7)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.00")   ' 情報2
                        vTemp = CVar(typ_y022z.MESDATA8)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' 情報３
                        vTemp = CVar(typ_y022z.MESDATA9)
                        .typ_rslt(UpDo, DispLineCount).INFO4 = DBData2DispData(vTemp, "0.00")   ' 情報４
                        JiltusekiUmu(UpDo, EPBMD) = True
                        '5番目の情報：AN温度
                        vTemp = CVar(typ_y022z.DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        .typ_rslt(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")      ' 情報5
                    End If
                ElseIf SijiUmu = "2" Then
                    bJudg = False
                    '画面表示内容設定
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ異常"                          ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                End If
                
            End If
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' 判定結果
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                      ' 判定結果
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If UpDo = SxlTop Or UpDo = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                Select Case BmdNo
                Case 1
                    gsTbcmy028ErrCode = "00152"
                Case 2
                    gsTbcmy028ErrCode = "00153"
                Case 3
                    gsTbcmy028ErrCode = "00154"
                End Select
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            DispLineCount = DispLineCount + 1
        Else
            If shiji <> 0 Then
                '画面表示内容設定
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' 結晶内開始位置
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())          ' 内容
                .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様無"                                 ' 情報1
                .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                                 ' 情報2
                .typ_rslt(UpDo, DispLineCount).INFO3 = "実績無"                                 ' 情報3
                .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                       ' 情報4
                .typ_rslt(UpDo, DispLineCount).INFO5 = ""                                       ' 情報5
                .typ_rslt(UpDo, DispLineCount).INFO6 = ""                                       ' 情報6
                .typ_rslt(UpDo, DispLineCount).INFO7 = ""                                       ' 情報7
                .typ_rslt(UpDo, DispLineCount).INFO8 = ""                                       ' 情報8
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y022z.SAMPLEID                      ' サンプルＮｏ
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' 判定結果
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                               ' 品番(12桁)
                If (SijiUmu = "1") And (Trim(typ_y022z.SAMPLEID) <> "0") Then
                    'BMDE判定失敗
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                            ' 情報2
                    'BMD判定
                    If EpBmdJudg(typ_CType.typ_si, typ_y022z, bJudg, BmdNo) Then
                        '画面表示内容設定　　2003/05/20 ooba
                        vTemp = CVar(typ_y022z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' 情報1
                        vTemp = CVar(typ_y022z.MESDATA7)
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.00")   ' 情報2
                        vTemp = CVar(typ_y022z.MESDATA8)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' 情報3
                        vTemp = CVar(typ_y022z.MESDATA9)
                        .typ_rslt(UpDo, DispLineCount).INFO4 = DBData2DispData(vTemp, "0.00")   ' 情報4
                        JiltusekiUmu(UpDo, EPBMD) = True
                        vTemp = CVar(typ_y022z.DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        .typ_rslt(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")      ' 情報5
                    End If
                ElseIf SijiUmu = "2" Then
                    bJudg = False
                    '画面表示内容設定
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ異常"                          ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                End If
                'Add Start 2011/11/28 Y.Hitomi 中間抜試の場合は、参考表示する
                If sSxlPos = "MID" And bJudg = False Then
                    If BmdNo = 1 And JudgSW.B1E Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "参考"                                      ' 判定結果
                        MidlJudg = False
                    ElseIf BmdNo = 2 And JudgSW.B2E Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "参考"                                      ' 判定結果
                        MidlJudg = False
                    ElseIf BmdNo = 3 And JudgSW.B3E Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "参考"                                      ' 判定結果
                        MidlJudg = False
                    End If
                End If
                'Add End  2011/11/28 Y.Hitomi
                DispLineCount = DispLineCount + 1
            End If
        End If
    
        Select Case BmdNo
        Case 1
            .typ_y022(UpDo, EPBMD1) = typ_y022z
        Case 2
            .typ_y022(UpDo, EPBMD2) = typ_y022z
        Case 3
            .typ_y022(UpDo, EPBMD3) = typ_y022z
        End Select
    
    End With
    
End Sub

Public Sub OSFDataSet_EP(OsfNo As Integer, UpDo As Integer, INGOTPOS As Integer, DispLineCount As Integer, typTmList() As typ_TBCMB005, sHinban12 As String)

    Dim JudgSpecCode    As Boolean
    Dim SCC             As String
    Dim shiji           As Boolean
    Dim IND             As String * 4                   '検査指示
''    Dim typ_y013z       As typ_TBCMY013
    Dim typ_y022z       As typ_TBCMY022
    Dim AveMax(1)       As String                       '平均/最大判定値
    Dim bJudg           As Boolean
    Dim vTemp           As Variant
    Dim SijiUmu         As String
    Dim EPBmSokuP       As String                       ' 測定位置＿点
    Dim EPBmSokuHou     As String                       ' 品WFOSF1測定位置_方
    Dim EPBmSokuRyou    As String                       ' 品WFOSF1測定位置_領
    Dim EPOSF           As Integer
    Dim sSxlPos         As String                       'SXL位置(TOP/BOT)

    '検査指示設定
    IND = IIf(UpDo = SxlTop, "123", "123")

'Chg Start 2011/03/10 SMPK Miyata
'    sSxlPos = IIf(UpDo = SxlTop, "TOP", "BOT")
    Select Case UpDo
        Case SxlTop:    sSxlPos = "TOP"
        Case SxlTail:   sSxlPos = "BOT"
        Case Else:      sSxlPos = "MID"
    End Select
'Chg End   2011/03/10 SMPK Miyata

    With typ_CType_EP
        Select Case OsfNo
        Case 1
            EPOSF = EPOSF1
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.L1E And CheckKHN_EP(typ_CType.typ_si.HEPOF1KN, 3, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.L1E And CheckKHN_EP(typ_CType.typ_si.HEPOF1KN, 3, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '保証方法=保証 かつ 中間抜試（保証）の場合、仕様有とする
                JudgSpecCode = (JudgSW.L1E And typ_CType.typ_si.MSMPFLGEPOF = "1" And _
                                (typ_CType.typ_si.MSMPFLG = "1" Or typ_CType.typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.L1E And .typ_si.MSMPFLGEPBM = "1")
'Chg End　 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "L1E"
            shiji = (InStr(IND, typ_CType.typ_Param.WFSMP(UpDo).EPINDL1CW) <> 0)
            SijiUmu = typ_CType.typ_Param.WFSMP(UpDo).EPRESL1CW
            EPBmSokuHou = typ_CType.typ_si.HEPOF1SH
            EPBmSokuP = typ_CType.typ_si.HEPOF1ST
            EPBmSokuRyou = typ_CType.typ_si.HEPOF1SR
        Case 2
            EPOSF = EPOSF2
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.L2E And CheckKHN_EP(typ_CType.typ_si.HEPOF2KN, 4, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.L2E And CheckKHN_EP(typ_CType.typ_si.HEPOF2KN, 4, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '保証方法=保証 かつ 中間抜試（保証）の場合、仕様有とする
                JudgSpecCode = (JudgSW.L2E And typ_CType.typ_si.MSMPFLGEPOF = "1" And _
                                (typ_CType.typ_si.MSMPFLG = "1" Or typ_CType.typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.L2E And .typ_si.MSMPFLGEPBM = "1")
'Chg End　 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "L2E"
            shiji = (InStr(IND, typ_CType.typ_Param.WFSMP(UpDo).EPINDL2CW) <> 0)
            SijiUmu = typ_CType.typ_Param.WFSMP(UpDo).EPRESL2CW
            EPBmSokuHou = typ_CType.typ_si.HEPOF2SH
            EPBmSokuP = typ_CType.typ_si.HEPOF2ST
            EPBmSokuRyou = typ_CType.typ_si.HEPOF2SR
        Case 3
            EPOSF = EPOSF3
'Chg Start 2011/03/10 SMPK Miyata
'            JudgSpecCode = JudgSW.L3E And CheckKHN_EP(typ_CType.typ_si.HEPOF3KN, 5, sSxlPos)
            If UpDo = SxlTop Or UpDo = SxlTail Then
                JudgSpecCode = JudgSW.L3E And CheckKHN_EP(typ_CType.typ_si.HEPOF3KN, 5, sSxlPos)
            Else
'Chg Start 2011/08/10 Y.Hitomi
                '保証方法=保証 かつ 中間抜試（保証）の場合、仕様有とする
                JudgSpecCode = (JudgSW.L3E And typ_CType.typ_si.MSMPFLGEPOF = "1" And _
                                (typ_CType.typ_si.MSMPFLG = "1" Or typ_CType.typ_si.MSMPFLG = "3"))
'                JudgSpecCode = (JudgSW.L3E And .typ_si.MSMPFLGEPBM = "1")
'Chg End　 2011/08/10 Y.Hitomi
            End If
'Chg End   2011/03/10 SMPK Miyata
            SCC = "L3E"
            shiji = (InStr(IND, typ_CType.typ_Param.WFSMP(UpDo).EPINDL3CW) <> 0)
            SijiUmu = typ_CType.typ_Param.WFSMP(UpDo).EPRESL3CW
            EPBmSokuHou = typ_CType.typ_si.HEPOF3SH
            EPBmSokuP = typ_CType.typ_si.HEPOF3ST
            EPBmSokuRyou = typ_CType.typ_si.HEPOF3SR
        End Select
        typ_y022z = .typ_y022(UpDo, EPOSF)


        '' WF検査指示（OSFE)*****************************************************************
        If JudgSpecCode Then
            '画面表示内容設定
            .typ_rslt(UpDo, DispLineCount).pos = -1                                             ' 結晶内開始位置
            .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())              ' 内容
            .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様有"                                     ' 情報1
            .typ_rslt(UpDo, DispLineCount).INFO2 = "検査無"                                     ' 情報2
            .typ_rslt(UpDo, DispLineCount).INFO3 = ""                                           ' 情報3
            .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                           ' 情報4
            .typ_rslt(UpDo, DispLineCount).INFO5 = ""                                           ' 情報5
            .typ_rslt(UpDo, DispLineCount).INFO6 = ""                                           ' 情報6
            .typ_rslt(UpDo, DispLineCount).INFO7 = ""                                           ' 情報7
            .typ_rslt(UpDo, DispLineCount).INFO8 = ""                                           ' 情報8
            .typ_rslt(UpDo, DispLineCount).SMPLID = -1                                          ' サンプルＮｏ
            .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                          ' 判定結果
            .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                                   ' 品番(12桁)
            bJudg = False
            If shiji <> 0 Then
                '画面表示内容設定
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' 結晶内開始位置
                .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                                 ' 情報2
                .typ_rslt(UpDo, DispLineCount).INFO3 = "実績無"                                 ' 情報3
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y022z.SAMPLEID                      ' サンプルＮｏ
                'OSF判定取得
                If (SijiUmu = "1") And (Trim(typ_y022z.SAMPLEID) <> "0") Then
                    'OSF判定失敗
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                            ' 情報2
                    'OSF判定取得
                    If EpOsfJudg(typ_CType.typ_si, typ_y022z, bJudg, OsfNo, AveMax()) Then             ' AveMax
                        '画面表示内容設定
                        vTemp = CVar(AveMax(0))
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' 情報1
                        vTemp = CVar(AveMax(1))
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")    ' 情報2
                        vTemp = CVar(typ_y022z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' 情報3
                        vTemp = CVar(IIf(Trim(typ_y022z.MESDATA9) = "", "-", Trim(typ_y022z.MESDATA9)) _
                                        & "  " & IIf(Trim(typ_y022z.MESDATA12) = "", "-", Trim(typ_y022z.MESDATA12)) _
                                        & "  " & IIf(Trim(typ_y022z.MESDATA15) = "", "-", Trim(typ_y022z.MESDATA15)))
                        .typ_rslt(UpDo, DispLineCount).INFO4 = vTemp                            ' 情報4
                        JiltusekiUmu(UpDo, EPOSF) = True
                        vTemp = CVar(typ_y022z.DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        .typ_rslt(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")      ' 情報5
                    End If
                ElseIf SijiUmu = "2" Then
                    '画面表示内容設定
                    bJudg = False
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ異常"                          ' 情報３
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報４
                End If
            End If
            If bJudg = True Then
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' 判定結果
            Else
                .typ_rslt(UpDo, DispLineCount).OKNG = "NG"                                      ' 判定結果
'Chg Start 2011/03/09 SMPK Miyata
'                TotalJudg = False
                If UpDo = SxlTop Or UpDo = SxlTail Then TotalJudg = False Else MidlJudg = False
'Chg End   2011/03/09 SMPK Miyata
'--------------- 2008/07/25 INSERT START  By Systech ---------------
                Select Case OsfNo
                Case 1
                    gsTbcmy028ErrCode = "00149"
                Case 2
                    gsTbcmy028ErrCode = "00150"
                Case 3
                    gsTbcmy028ErrCode = "00151"
                End Select
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            End If
            DispLineCount = DispLineCount + 1
        Else
            If shiji <> 0 Then
                '画面表示内容設定
                .typ_rslt(UpDo, DispLineCount).pos = CStr(INGOTPOS)                             ' 結晶内開始位置
                .typ_rslt(UpDo, DispLineCount).NAIYO = Search_CrCode(SCC, typTmList())          ' 内容
                .typ_rslt(UpDo, DispLineCount).INFO1 = "仕様無"                                 ' 情報1
                .typ_rslt(UpDo, DispLineCount).INFO2 = "検査有"                                 ' 情報2
                .typ_rslt(UpDo, DispLineCount).INFO3 = "実績無"                                 ' 情報3
                .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                       ' 情報4
                .typ_rslt(UpDo, DispLineCount).INFO5 = ""                                       ' 情報5
                .typ_rslt(UpDo, DispLineCount).INFO6 = ""                                       ' 情報6
                .typ_rslt(UpDo, DispLineCount).INFO7 = ""                                       ' 情報7
                .typ_rslt(UpDo, DispLineCount).INFO8 = ""                                       ' 情報8
                .typ_rslt(UpDo, DispLineCount).SMPLID = typ_y022z.SAMPLEID                      ' サンプルＮｏ
                .typ_rslt(UpDo, DispLineCount).OKNG = "OK"                                      ' 判定結果
                .typ_rslt(UpDo, DispLineCount).hinban = sHinban12                               ' 品番(12桁)
                'OSF判定取得
                If (SijiUmu = "1") And (Trim(typ_y022z.SAMPLEID) <> "0") Then
                    'OSF判定失敗
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "判定Err"                            ' 情報2
                    'OSF判定取得
                    If EpOsfJudg(typ_CType.typ_si, typ_y022z, bJudg, OsfNo, AveMax()) Then             ' AveMax
                        '画面表示内容設定
                        vTemp = CVar(AveMax(0))
                        .typ_rslt(UpDo, DispLineCount).INFO1 = DBData2DispData(vTemp, "0.00")   ' 情報1
                        vTemp = CVar(AveMax(1))
                        .typ_rslt(UpDo, DispLineCount).INFO2 = DBData2DispData(vTemp, "0.0")    ' 情報2
                        vTemp = CVar(typ_y022z.MESDATA6)
                        .typ_rslt(UpDo, DispLineCount).INFO3 = DBData2DispData(vTemp, "0.00")   ' 情報3
                        vTemp = CVar(IIf(Trim(typ_y022z.MESDATA9) = "", "-", Trim(typ_y022z.MESDATA9)) _
                                        & "  " & IIf(Trim(typ_y022z.MESDATA12) = "", "-", Trim(typ_y022z.MESDATA12)) _
                                        & "  " & IIf(Trim(typ_y022z.MESDATA15) = "", "-", Trim(typ_y022z.MESDATA15)))
                        .typ_rslt(UpDo, DispLineCount).INFO4 = vTemp                            ' 情報4
                         JiltusekiUmu(UpDo, EPOSF) = True
                        vTemp = CVar(typ_y022z.DKAN)
                        '3～6桁目がAN温度
                        vTemp = Mid(vTemp, 3, 4)
                        .typ_rslt(UpDo, DispLineCount).INFO5 = DBData2DispData(vTemp, "0")      ' 情報5
                    End If
                ElseIf SijiUmu = "2" Then
                    '画面表示内容設定
                    bJudg = False
                    .typ_rslt(UpDo, DispLineCount).INFO3 = "ｻﾝﾌﾟﾙ異常"                          ' 情報3
                    .typ_rslt(UpDo, DispLineCount).INFO4 = ""                                   ' 情報4
                End If
                'Add Start 2011/11/28 Y.Hitomi 中間抜試の場合は、参考表示する
                If sSxlPos = "MID" And bJudg = False Then
                    If OsfNo = 1 And JudgSW.L1E Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "参考"                                      ' 判定結果
                        MidlJudg = False
                    ElseIf OsfNo = 2 And JudgSW.L2E Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "参考"                                      ' 判定結果
                        MidlJudg = False
                    ElseIf OsfNo = 3 And JudgSW.L3E Then
                        .typ_rslt(UpDo, DispLineCount).OKNG = "参考"                                      ' 判定結果
                        MidlJudg = False
                    End If
                End If
                'Add End  2011/11/28 Y.Hitomi
                DispLineCount = DispLineCount + 1
            End If
        End If

        Select Case OsfNo
        Case 1
            .typ_y022(UpDo, EPOSF1) = typ_y022z
            TmpOsfData(0, UpDo, EPOSF1) = AveMax(0)
            TmpOsfData(1, UpDo, EPOSF1) = AveMax(1)
            TmpOsfMBNP(0, UpDo, EPOSF1) = IIf(typ_y022z.MESDATA9 = "-", " ", typ_y022z.MESDATA9)
            TmpOsfMBNP(1, UpDo, EPOSF1) = IIf(typ_y022z.MESDATA12 = "-", " ", typ_y022z.MESDATA12)
            TmpOsfMBNP(2, UpDo, EPOSF1) = IIf(typ_y022z.MESDATA15 = "-", " ", typ_y022z.MESDATA15)
        Case 2
            .typ_y022(UpDo, WFOSF2) = typ_y022z
            TmpOsfData(0, UpDo, EPOSF2) = AveMax(0)
            TmpOsfData(1, UpDo, EPOSF2) = AveMax(1)
            TmpOsfMBNP(0, UpDo, EPOSF2) = IIf(typ_y022z.MESDATA9 = "-", " ", typ_y022z.MESDATA9)
            TmpOsfMBNP(1, UpDo, EPOSF2) = IIf(typ_y022z.MESDATA12 = "-", " ", typ_y022z.MESDATA12)
            TmpOsfMBNP(2, UpDo, EPOSF2) = IIf(typ_y022z.MESDATA15 = "-", " ", typ_y022z.MESDATA15)
        Case 3
            .typ_y022(UpDo, EPOSF3) = typ_y022z
            TmpOsfData(0, UpDo, EPOSF3) = AveMax(0)
            TmpOsfData(1, UpDo, EPOSF3) = AveMax(1)
            TmpOsfMBNP(0, UpDo, EPOSF3) = IIf(typ_y022z.MESDATA9 = "-", " ", typ_y022z.MESDATA9)
            TmpOsfMBNP(1, UpDo, EPOSF3) = IIf(typ_y022z.MESDATA12 = "-", " ", typ_y022z.MESDATA12)
            TmpOsfMBNP(2, UpDo, EPOSF3) = IIf(typ_y022z.MESDATA15 = "-", " ", typ_y022z.MESDATA15)
        End Select

    End With
End Sub

'------------------------------------------------
' エピ測定評価結果取得
'------------------------------------------------

'概要      :サンプルＩＤからTBCMY022を検索し、EP先行評価結果を取得し返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typIn         ,I  ,type_DBDRV_scmzc_fcmlc001c_In         ,入力用
'          :records()     ,O  ,typ_TBCMY022 ,抽出レコード
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :SB_WfJudg_SQL.funGetTBCMY013を基に作成
'履歴      :新規作成 2006/08/15 エピ先行評価追加対応 SMP)kondoh
Public Function funGetTBCMY022_All(typIn As type_DBDRV_scmzc_fcmlc001c_In, records() As typ_TBCMY022) As FUNCTION_RETURN
    
    Dim sql     As String       'SQL全体
    Dim rs      As OraDynaset   'RecordSet
    Dim recCnt  As Long         'レコード数
    Dim i       As Long

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_WfJudg_SQL.bas -- Function funGetTBCMY022_All"

    ''SQLを組み立てる
    sql = "select SAMPLEID, OSITEM, MAISU, SPEC, NETSU, ET, MES, DKAN, MESDATA1, MESDATA2, MESDATA3, MESDATA4, MESDATA5, "
    sql = sql & "MESDATA6, MESDATA7, MESDATA8, MESDATA9, MESDATA10, MESDATA11, MESDATA12, MESDATA13, MESDATA14, MESDATA15, "
    sql = sql & "TXID, REGDATE, SENDFLAG, SENDDATE "
    sql = sql & "from TBCMY022 "
    sql = sql & "where ('" & typIn.WFSMP.EPINDB1CW & "' > '0' and SAMPLEID = '" & typIn.WFSMP.EPSMPLIDB1CW & "' and SPEC = '" & OSEPBMD1 & "') or "
    sql = sql & "      ('" & typIn.WFSMP.EPINDB2CW & "' > '0' and SAMPLEID = '" & typIn.WFSMP.EPSMPLIDB2CW & "' and SPEC = '" & OSEPBMD2 & "') or "
    sql = sql & "      ('" & typIn.WFSMP.EPINDB3CW & "' > '0' and SAMPLEID = '" & typIn.WFSMP.EPSMPLIDB3CW & "' and SPEC = '" & OSEPBMD3 & "') or "
    sql = sql & "      ('" & typIn.WFSMP.EPINDL1CW & "' > '0' and SAMPLEID = '" & typIn.WFSMP.EPSMPLIDL1CW & "' and SPEC = '" & OSEPOSF1 & "') or "
    sql = sql & "      ('" & typIn.WFSMP.EPINDL2CW & "' > '0' and SAMPLEID = '" & typIn.WFSMP.EPSMPLIDL2CW & "' and SPEC = '" & OSEPOSF2 & "') or "
    sql = sql & "      ('" & typIn.WFSMP.EPINDL3CW & "' > '0' and SAMPLEID = '" & typIn.WFSMP.EPSMPLIDL3CW & "' and SPEC = '" & OSEPOSF3 & "') or "
    sql = sql & "      ('" & typIn.WFSMP.WFINDOT2CW & "' > '0' and SAMPLEID = '" & typIn.WFSMP.WFSMPLIDOT2CW & "' and SPEC = '" & OSEPOT2 & "') "
    
    Debug.Print sql
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        Set rs = Nothing
        ReDim records(0)
        funGetTBCMY022_All = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
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
        rs.MoveNext
    Next
    Set rs = Nothing

    funGetTBCMY022_All = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    funGetTBCMY022_All = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'Add Start 2011/04/25 SMPK Miyata
'------------------------------------------------
' 中間抜試実績判定
'------------------------------------------------

'概要      :中間抜試の実績値判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型             :説明
'          :sKeyID          ,I  ,String         :SXL-ID
'          :tNew_Hinban     ,I  ,String         :振替候補品番
'          :bTotalJudg      ,O  ,Boolean        :トータル判定
'          :iErr_Code       ,O  ,Integer        :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String         :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :typ_CType       ,O  ,typ_AllTypesC  :全情報構造体(構造体)
'          :iSmpGetFlg      ,I  ,Integer        :ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ(0:ｻﾝﾌﾟﾙ指定なし, 1:ｻﾝﾌﾟﾙ指定あり)
'          :sSamplID1       ,I  ,String         :TOPｻﾝﾌﾟﾙID(省略可)
'          :sSamplID2       ,I  ,String         :BOTｻﾝﾌﾟﾙID(省略可)
'          :iKcnt           ,I  ,Integer        :工程連番(省略可)
'          :戻り値          ,O  ,Integer        :取得の成否(0:正常終了, -1:異常終了)
'説明      :
'履歴      :

Public Function funWfcMidleHantei(sKeyID As String, tNew_Hinban As tFullHinban, _
                bTotalJudg As Boolean, iErr_Code As Integer, sErr_Msg As String, typ_CType As typ_AllTypesC, _
                iSmpGetFlg As Integer, Optional sSamplID1 As String = vbNullString, Optional sSamplID2 As String = vbNullString, _
                Optional iKcnt As Integer = 0) As Integer
    
    Dim i       As Integer      'Add 2011/03/09 SMPK Miyata
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funWfcMidleHantei = FUNCTION_RETURN_FAILURE
    
    'グローバル変数に設定
    wiSmpGetFlg = iSmpGetFlg
    wiKcnt = iKcnt

    tNew_Hinban = tMapHinG.HIN

    '初期設定
    sErr_Msg = "WFC中間抜試実績判定(SetInitData)"
    If SetInitData(sKeyID, tNew_Hinban, typ_CType.typ_Param, iSmpGetFlg, sSamplID1, sSamplID2) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If

    '画面情報設定
    sErr_Msg = "WFC中間抜試実績判定(SetAllData_Mid)"
    If SetAllData_Mid(typ_CType, tNew_Hinban, iSmpGetFlg, iErr_Code, sErr_Msg) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If

    TotalJudg = True
    MidlJudg = True             '中間抜試判定

    '仕様検査指示取得
    sErr_Msg = "WFC中間抜試実績判定(SpecJudgCheck)"
    SpecJudgCheck

    '仕様Nullチェック
    sErr_Msg = "仕様Nullﾁｪｯｸ"
    If funWfChkNull(typ_CType.typ_si, sErr_Msg) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If

    '実績データ判定(MIDLE)
    sErr_Msg = "WFC中間抜試実績判定(WfAllJudg(MIDLE))"
    
    For i = SxlMidl To UBound(typ_CType.typ_Param.WFSMP)
    
        '該当ﾌﾞﾛｯｸの中間抜試か？
        If typ_CType.typ_Param.WFSMP(i).INPOSCW >= tMapHinG.INPOSCS_S And _
           typ_CType.typ_Param.WFSMP(i).INPOSCW < tMapHinG.INPOSCS_E Then
        
            'WF判定 (全)
            If WfAllJudg(typ_CType, tNew_Hinban, i) = FUNCTION_RETURN_FAILURE Then
                GoTo Apl_down
            End If
        End If

    Next i

    '画面情報設定
    sErr_Msg = "WFC中間抜試実績判定(エピ)(SetAllData_Mid_EP)"
    If SetAllData_Mid_EP(typ_CType, typ_CType_EP, tNew_Hinban, iSmpGetFlg, iErr_Code, sErr_Msg) = FUNCTION_RETURN_FAILURE Then
        GoTo Apl_down
    End If

    '仕様検査指示取得
    sErr_Msg = "WFC中間抜試実績判定(エピ)(SpecJudgCheck)"
    SpecJudgCheck

    '実績データ判定(MIDLE)
    sErr_Msg = "WFC中間抜試実績判定(エピ)(EPJudge(MIDLE))"
    
    For i = SxlMidl To UBound(typ_CType.typ_Param.WFSMP)
        '該当ﾌﾞﾛｯｸの中間抜試か？
        If typ_CType.typ_Param.WFSMP(i).INPOSCW >= tMapHinG.INPOSCS_S And _
           typ_CType.typ_Param.WFSMP(i).INPOSCW < tMapHinG.INPOSCS_E Then
            '結晶判定(エピ)
            If EPJudge(typ_CType, typ_CType_EP, tNew_Hinban, i) = FUNCTION_RETURN_FAILURE Then
                GoTo Apl_down
            End If
        End If
    Next i

    Dim iMinMidCnt      As Integer       '中間抜試の必要数
    Dim iRstMidCnt      As Integer       '中間抜試の件数
    
    ' 中間抜試品か？
    If typ_CType.typ_si.MSMPFLG = "1" Then
        '中間抜試の必要数 = (SXLのWF枚数 - 中間抜試許容値(枚数)) / 中間抜試単位(枚数)
        iMinMidCnt = Fix((typ_CType.typ_Param.COUNT - typ_CType.typ_si.MSMPCONSTMAI) / typ_CType.typ_si.MSMPTANIMAI)
        'マイナスの場合、０とする
        If iMinMidCnt < 0 Then iMinMidCnt = 0
        
        '中間抜試の件数
        iRstMidCnt = (UBound(typ_CType.typ_Param.WFSMP) - SxlMidl) + 1
        If iRstMidCnt < iMinMidCnt Then
            typ_CType.sMidErrMsg = "中間抜試実績がありません。　仕様(" & iMinMidCnt & ") 実績(" & iRstMidCnt & ")"
            MidlJudg = False
        End If
        
    End If

    bTotalJudg = TotalJudg And MidlJudg

    funWfcMidleHantei = FUNCTION_RETURN_SUCCESS
    
'------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funWfcMidleHantei = -4
    iErr_Code = funWfcMidleHantei
    GoTo Apl_Exit
    
End Function


'概要      :画面情報データ設定
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typ_CType     ,I  ,typ_AllTypesC ,各情報構造体
'説明      :画面情報を情報構造体に設定する
'履歴      :
Private Function SetAllData_Mid(typ_CType As typ_AllTypesC, tNew_Hinban As tFullHinban, iSmpGetFlg As Integer, _
                                iErr_Code As Integer, sErr_Msg As String) As FUNCTION_RETURN
    
    Dim typ_in      As type_DBDRV_scmzc_fcmlc001c_In     ' DBアクセス入力用
    Dim fret(2)     As FUNCTION_RETURN
    Dim RET         As FUNCTION_RETURN
    Dim typ_hi()    As typ_TBCMH004
    Dim sErrMsg     As String
    Dim FuncAns     As FUNCTION_RETURN ''2001/12/18 S.Sano
    Dim records()   As typ_TBCMH001
    Dim i           As Integer      'カウンタ
    Dim iMidNo      As Integer      '中間抜試No

    SetAllData_Mid = FUNCTION_RETURN_FAILURE
    
    typ_in.HIN = tNew_Hinban
    typ_in.SXLID = typ_CType.typ_Param.SXLID

    With typ_CType
        
        For i = SxlMidl To UBound(.typ_Param.WFSMP)
            iMidNo = i - SxlMidl + 1

            If iMidNo > SXL_MAXSMP Then
                ' 中間抜試最大件数オーバー
                Exit Function
            End If

            'MIDLE側
            sErr_Msg = "WFC中間抜試実績判定(MIDLE_" & iMidNo & " 初期ﾃﾞｰﾀ設定)"
            typ_in.SAMPLEID = .typ_Param.WFSMP(i).REPSMPLIDCW
            typ_in.WFSMP = typ_CType.typ_Param.WFSMP(i)

            '評価測定結果取得
            ReDim Preserve .typ_y013midl_ary(iMidNo)
            sErr_Msg = "WFC中間抜試実績判定(MIDLE_" & iMidNo & " funWfcGetDataEtc)"
            RET = funWfcGetDataEtc(typ_in, i, tNew_Hinban, iSmpGetFlg, _
                                    .typ_si, _
                                    .typ_y013midl_ary(iMidNo).typ_y013midl, _
                                    sErrMsg)
            If RET = FUNCTION_RETURN_SUCCESS Then

                ' 評価測定結果整列
                sErr_Msg = "WFC中間抜試実績判定(MIDLE_" & iMidNo & " 評価測定結果整列)"
                If SetMERInd(typ_CType, .typ_y013midl_ary(iMidNo).typ_y013midl, i) <> True Then
                    '評価測定結果整列失敗
                    Exit Function
                End If

                '引上げ終了実績取得
                ReDim typ_hi(0)
                sErr_Msg = "WFC中間抜試実績判定(MIDLE_" & iMidNo & " 引上げ終了実績取得)"
                If s_cmmc001db_Sql(.typ_Param.CRYNUM, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                    '引上げ終了実績取得失敗
                    Exit Function
                Else
                    If UBound(typ_hi) <> 0 Then
                        .typ_hage(i) = typ_hi(1)
                    Else
                        '引上げ終了実績取得失敗
                        SetAllData_Mid = FUNCTION_RETURN_FAILURE
                        Exit Function
                    End If
                End If
            Else
                SetAllData_Mid = FUNCTION_RETURN_FAILURE
                Exit Function
            End If
        Next i
    End With
    
    '' Ｐ＋結晶の判断
    sErr_Msg = "WFC中間抜試実績判定(P+結晶の判断)"
    If DBDRV_GetTBCMH001(records(), "where UPINDNO = '" & left(SelectSxlID, 7) & "0" & Mid(SelectSxlID, 9, 1) & "'") = FUNCTION_RETURN_SUCCESS Then

    Else
        SetAllData_Mid = FUNCTION_RETURN_FAILURE
        Exit Function
    End If
    
    SetAllData_Mid = FUNCTION_RETURN_SUCCESS
End Function

'概要      :画面情報データ設定(エピ)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :typ_CType     ,I  ,typ_AllTypesC ,各情報構造体
'          :typ_CType_EP  ,I  ,typ_AllTypesC_EP ,各情報構造体
'説明      :画面情報を情報構造体に設定する
'履歴      :
Private Function SetAllData_Mid_EP(typ_CType As typ_AllTypesC, typ_CType_EP As typ_AllTypesC_EP, tNew_Hinban As tFullHinban, iSmpGetFlg As Integer, _
                                                    iErr_Code As Integer, sErr_Msg As String) As FUNCTION_RETURN
    
    Dim typ_in      As type_DBDRV_scmzc_fcmlc001c_In     ' DBアクセス入力用
    Dim fret(2)     As FUNCTION_RETURN
    Dim typ_hi()    As typ_TBCMH004
    Dim sErrMsg     As String
    Dim FuncAns     As FUNCTION_RETURN
    Dim records()   As typ_TBCMH001
    Dim i           As Integer      'カウンタ
    Dim iMidNo      As Integer      '中間抜試No

    SetAllData_Mid_EP = FUNCTION_RETURN_FAILURE
    
    typ_in.HIN = tNew_Hinban
    typ_in.SXLID = typ_CType.typ_Param.SXLID

    With typ_CType_EP
        
        For i = SxlMidl To UBound(typ_CType.typ_Param.WFSMP)

            iMidNo = i - SxlMidl + 1
            
            If iMidNo > SXL_MAXSMP Then
                ' 中間抜試最大件数オーバー
                Exit Function
            End If

            'MIDLE側
            sErr_Msg = "WFC中間抜試実績判定(エピ)(MIDLE_" & iMidNo & " 初期ﾃﾞｰﾀ設定)(エピ)"
            typ_in.SAMPLEID = typ_CType.typ_Param.WFSMP(i).REPSMPLIDCW
            typ_in.WFSMP = typ_CType.typ_Param.WFSMP(i)
    
            '評価測定結果取得
            ReDim Preserve .typ_y022midl_ary(iMidNo)
            
            sErr_Msg = "WFC総合判定(MIDLE_" & iMidNo & " funGet_TBCME050)"
            '' エピ仕様を取得
            If funGet_TBCME050(tNew_Hinban, typ_CType.typ_si, sErrMsg) <> FUNCTION_RETURN_SUCCESS Then
                'エピ仕様取得失敗
                SetAllData_Mid_EP = FUNCTION_RETURN_FAILURE
                Exit Function
            End If
        
            sErr_Msg = "WFC中間抜試実績判定(エピ)(MIDLE_" & iMidNo & " funGetTBCMY022_All)"

            '' エピ測定評価結果(実績値)を取得(0件でもエラーではない)
            If funGetTBCMY022_All(typ_in, .typ_y022midl_ary(iMidNo).typ_y022midl) = FUNCTION_RETURN_FAILURE Then
                sErrMsg = GetMsgStr("EGET2", "Y022")
                SetAllData_Mid_EP = FUNCTION_RETURN_FAILURE
                Exit Function
            End If

            ' 評価測定結果整列
            sErr_Msg = "WFC中間抜試実績判定(エピ)(MIDLE_" & iMidNo & " 評価測定結果整列)"
            If SetMERInd_EP(typ_CType_EP, .typ_y022midl_ary(iMidNo).typ_y022midl, i) <> True Then
                '評価測定結果整列失敗
                Exit Function
            End If

            '引上げ終了実績取得
            ReDim typ_hi(0)
            
            sErr_Msg = "WFC中間抜試実績判定(エピ)(MIDLE_" & iMidNo & " 引上げ終了実績取得)"
            If s_cmmc001db_Sql(typ_CType.typ_Param.CRYNUM, typ_hi()) <> FUNCTION_RETURN_SUCCESS Then
                '引上げ終了実績取得失敗
                Exit Function
            Else
                If UBound(typ_hi) <> 0 Then
                    typ_CType.typ_hage(i) = typ_hi(1)
                Else
                    '引上げ終了実績取得失敗
                    SetAllData_Mid_EP = FUNCTION_RETURN_FAILURE
                    Exit Function
                End If
            End If
        Next i
    End With
    
    '' Ｐ＋結晶の判断
    sErr_Msg = "WFC中間抜試実績判定(エピ)(P+結晶の判断)"
    If DBDRV_GetTBCMH001(records(), "where UPINDNO = '" & left(SelectSxlID, 7) & "0" & Mid(SelectSxlID, 9, 1) & "'") = FUNCTION_RETURN_SUCCESS Then
    Else
        SetAllData_Mid_EP = FUNCTION_RETURN_FAILURE
        Exit Function
    End If
    
    SetAllData_Mid_EP = FUNCTION_RETURN_SUCCESS
End Function

'Add End   2011/04/25 SMPK Miyata

'概要      :サンプル間の枚数を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型          ,説明
'          :sSxlid        ,I  ,String      ,SXLID
'          :iMaisu()      ,O  ,Integer  　 ,サンプル間枚数
'          :戻り値        ,O  ,FUNCTION_RETURN,抽出の成否
'説明      :
'履歴      :2011/07/19 作成  Marushita
Public Function fncGetSmpMai(sSXLID As String, ByRef iMaisu() As Integer) As FUNCTION_RETURN
Dim rs      As OraDynaset               '抽出RecordDynaset
Dim rsCnt   As Integer                  'ﾚｺｰﾄﾞｶｳﾝﾄ
Dim sql     As String                   'SQL文
Dim i       As Integer                  'ﾙｰﾌﾟｶｳﾝﾄ
Dim iMCnt   As Integer                  'データ枚数
Dim iKcnt   As Integer                  '枚数件数
Dim iSflg   As Integer                  'サンプルフラグ
Dim sSmpId  As String                   'サンプルID

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "-- Function fncGetSmpMai"

    'SQL文の作成
    sql = "Select NVL(WFSTA,' ') AS WFSTA,  NVL(MSMPLEID,' ') AS MSMPLEID FROM TBCMY011 "
    sql = sql & "Where MSXLID = '" & sSXLID & "' "
    sql = sql & "AND   EXISTFLG = 'Y' "
    sql = sql & "ORDER BY LOTID, BLOCKSEQ "

    'データの抽出
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    '''抽出レコードが存在しない場合
    If rs.EOF Then
        ReDim iMaisu(0)                     '配列の初期化
        fncGetSmpMai = FUNCTION_RETURN_FAILURE   'ｴﾗｰｽﾃｰﾀｽ
        GoTo proc_exit
    End If

    iKcnt = 0
    ReDim iMaisu(iKcnt)
    iMCnt = 0
    iSflg = 0
    sSmpId = ""
    rsCnt = rs.RecordCount                  'ﾚｺｰﾄﾞ数のｶｳﾝﾄを取る

    '配列に値をセット
    rs.MoveFirst                            '先頭ﾚｺｰﾄﾞに移動
    For i = 0 To rsCnt - 1                  'ﾚｺｰﾄﾞ数分ﾙｰﾌﾟ
        DoEvents
        '欠落はSKIP
        If Trim(CStr(rs!MSMPLEID)) = "" And CStr(rs!WFSTA) = "4" Then
        Else
            '抜試単位の判断
            If Trim(CStr(rs!MSMPLEID)) <> "" And sSmpId <> Trim(CStr(rs!MSMPLEID)) Then
                '先頭データの判断
                If sSmpId = "" Then
                    sSmpId = Trim(CStr(rs!MSMPLEID))
                    iMCnt = 1
                Else
                    '同一サンプルIDのチェック(サンプルが変わった時の同一サンプルID判断用処理)
                    If iSflg = 0 Then
                        iMCnt = iMCnt + 1
                        sSmpId = Trim(CStr(rs!MSMPLEID))
                        iSflg = 1
                    Else
                        iMaisu(iKcnt) = iMCnt
                        iMCnt = 1
                        iKcnt = iKcnt + 1
                        ReDim Preserve iMaisu(iKcnt)
                        iSflg = 0
                    End If
                End If
            Else
                '同じサンプルIDの判断
                If Trim(CStr(rs!MSMPLEID)) <> "" And sSmpId = Trim(CStr(rs!MSMPLEID)) Then
                    iMCnt = iMCnt + 1
                Else
                    'サンプルIDなしの判断
                    If Trim(CStr(rs!MSMPLEID)) = "" And CStr(rs!WFSTA) = "0" Then
                        '同一サンプルIDのチェック(同一サンプルがなくなったら件数をセット)
                        If iSflg = 1 Then
                            iMaisu(iKcnt) = iMCnt
                            iMCnt = 1
                            iKcnt = iKcnt + 1
                            ReDim Preserve iMaisu(iKcnt)
                            iSflg = 0
                        Else
                            iMCnt = iMCnt + 1
                        End If
                    End If
                End If
            End If
        End If
        rs.MoveNext                         '次ﾚｺｰﾄﾞに移動
    Next
    If iMCnt > 0 Then
        iMaisu(iKcnt) = iMCnt
    End If

    fncGetSmpMai = FUNCTION_RETURN_SUCCESS   '正常ｽﾃｰﾀｽ


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


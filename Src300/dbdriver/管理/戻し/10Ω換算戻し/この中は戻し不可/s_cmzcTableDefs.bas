Attribute VB_Name = "s_cmzcTableDefs"
Option Explicit
'7/30

Public STAFFIDBUFF  As String
Public spread_Col As Long
Public spread_Row As Long
Public MaxLine As Long

''エラーメッセージ
Public Const ESTAF = "ESTAF" ''担当者コードが無効です｡
Public Const EIE00 = "EIE00" ''全てのデータ入力が完了していません｡
Public Const EIE01 = "EBLK1" ''ブロックIDの桁数が間違っています｡
Public Const EIM00 = "EIM00" ''購入単結晶受入実績　問い合わせ中。
Public Const EGET = "EGET" ''DBからの読込に失敗しました。
Public Const EAPLY = "EAPLY" ''DBへの書込に失敗しました。
Public Const EMAT1 = "EMAT1" '' 原料番号の桁数が間違っています｡
Public Const EMAT2 = "EMAT2" '' 指定した原料番号は未登録です。
Public Const KIE00 = "EBLK0" ''入力されたブロックIDは､存在しません｡
'Public Const KDE01 = "KDE01" ''購入単結晶は、イメージ表示できません。  ??????
Public Const KDE01 = "EKDE1" ''購入単結晶は、イメージ表示できません。
Public Const PWAIT = "PWAIT" ''少々お待ち下さい
Public Const KC001 = "EKC01" ''クリスタルカタログ処理が失敗しました！
Public Const TJE01 = "PJE01" ''総合判定NGです。
Public Const ESXL0 = "ESXL0" ''入力されたSXLIDは、存在しません。"
Public Const ESXL1 = "ESXL1" ''SXLIDの桁数が間違っています。"
Public Const PWCC0 = "PWCC0" ''クリスタルカタログ 検索中。
Public Const E0001 = "E0001" ''ブロックが選択されていません。
Public Const EGB01 = "EGB01" ''受入重量が多すぎます。
Public Const EGB02 = "EGB02" ''受入重量が正しくありません。
Public Const EINPM = "EINPM" ''入力値が不正です｡
Public Const PIN16 = "PLBL1" ''ラベルを再発行します。よろしいですか？
Public Const POK06 = "PLBL2" ''ラベルを再発行しました。
Public Const PLBL3 = "PLBL3" ''ラベル印刷中｡
Public Const ELB00 = "ELBL0" ''ラベル印刷エラー｡
Public Const EHIN1 = "EHIN1" ''品番の桁数が間違っています。"
Public Const EHIN0 = "EHIN0" ''指定の品番は未登録です。"
Public Const EBLK5 = "EBLK5" ''ブロックIDが、正しくありません。

Public Const E0002 = "E0002" ''選択したブロックは連続していません｡
Public Const WGD01 = "WGD01" ''GDエラーとなる品番があります
Public Const EKDE2 = "EKDE2" ''購入単結晶は､情報変更できません｡
Public Const EGET2 = "EGET2" ''DBからの読込に失敗しました(%s)。


Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Public EndFlag As Boolean
'                                     2001/08/24
'================================================
' ユーザ定義型の宣言
' 定義内容: 060200_全テーブル
'================================================


' SXL検査書
Public Type typ_TBCMX001
    SXLID As String * 13            ' SXLID
    FROMTOKBN As String * 1         ' FROMTO区分
    SAMPLE_FROM As String * 16      ' サンプルID (From)
    SAMPLE_TO As String * 16        ' サンプルID (To)
    BLOCKID As String * 12          ' ブロックID
    CRYNUM As String * 12           ' 結晶番号
    SXLDECDATE As Date              ' SXL-ID確定日付
    PLUPDATE As Date                ' 引上日付
    INGOTPOS As Integer             ' 結晶内開始位置
    hinban As String * 12           ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    PRODCOND As String * 10         ' 製作条件
    PGID As String * 8              ' ＰＧ－ＩＤ
    UPLENGTH As Integer             ' 引上げ長さ
    SXLPOS As Integer               ' SXL位置
    SXLLENGTH As Integer            ' SXL-ID確定長さ
    SXLWAFERCNT As Integer          ' SXL-ID確定時の枚数
    FREELENG As Integer             ' フリー長

    DIAMETER As Integer             ' 直径
    CHARGE As Long                  ' チャージ量
    SEED As String * 4              ' シード
    SAMPID As String * 16           ' サンプルID
    SXL_RS_SMPPOS As Integer        ' SXLRSｻﾝﾌﾟﾙ測定位置（SXL測定情報）
    SXLRS_MEAS1 As Double           ' SXLRS_測定値１
    SXLRS_MEAS2 As Double           ' SXLRS_測定値２
    SXLRS_MEAS3 As Double           ' SXLRS_測定値３
    SXLRS_MEAS4 As Double           ' SXLRS_測定値４
    SXLRS_MEAS5 As Double           ' SXLRS_測定値５
    SXLRS_EFEHS As Double           ' SXLRS_実効偏析
    SXLRS_RRG As Double             ' SXLRS_ＲＲＧ
    SXL_OI_SMPPOS As Integer        ' SXLOIｻﾝﾌﾟﾙ測定位置（SXL測定情報）
    SXLOI_OIMEAS1 As Double         ' SXLOI_Ｏｉ測定値１
    SXLOI_OIMEAS2 As Double         ' SXLOI_Ｏｉ測定値２
    SXLOI_OIMEAS3 As Double         ' SXLOI_Ｏｉ測定値３
    SXLOI_OIMEAS4 As Double         ' SXLOI_Ｏｉ測定値４
    SXLOI_OIMEAS5 As Double         ' SXLOI_Ｏｉ測定値５
    SXLOI_ORGRES As Double          ' SXLOI_ＯＲＧ結果
    SXLOI_INSPECTWAY As String * 2  ' SXLOI_検査方法
    SXL_CS_SMPPOS As Integer        ' SXLCSｻﾝﾌﾟﾙ測定位置（SXL測定情報）
    SXLCS_CSMEAS As Double          ' SXLCS_Cs実測値
    SXLCS_70PPRE As Double          ' SXLCS_７０％推定値
    SXLOSF_SMPPOS As Integer        ' OSFｻﾝﾌﾟﾙ測定位置（SXL位置情報）
    SXLOSF1_KKSP As String * 3      ' OSF1結晶欠陥測定位置
    SXLOSF1_NETU As String * 2      ' OSF1熱処理法
    SXLOSF1_KKSET As String * 3     ' OSF1結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLOSF1_CALCMAX As Double       ' OSF1SXL計算結果 Max_1
    SXLOSF1_CALCAVE As Double       ' OSF1SXL計算結果 Ave_1
    SXLOSF2_KKSP As String * 3      ' OSF２結晶欠陥測定位置
    SXLOSF2_NETU As String * 2      ' OSF２熱処理法
    SXLOSF2_KKSET As String * 3     ' OSF２結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLOSF2_CALCMAX As Double       ' OSF２SXL計算結果 Max_2
    SXLOSF2_CALCAVE As Double       ' OSF２SXL計算結果 Ave_2
    SXLOSF3_KKSP As String * 3      ' OSF３結晶欠陥測定位置
    SXLOSF3_NETU As String * 2      ' OSF３熱処理法
    SXLOSF3_KKSET As String * 3     ' OSF３結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLOSF3_CALCMAX As Double       ' OSF３SXL計算結果 Max_3
    SXLOSF3_CALCAVE As Double       ' OSF３SXL計算結果 Ave_3
    SXLOSF4_KKSP As String * 3      ' OSF４結晶欠陥測定位置
    SXLOSF4_NETU As String * 2      ' OSF４熱処理法
    SXLOSF4_KKSET As String * 3     ' OSF４結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLOSF4_CALCMAX As Double       ' OSF４SXL計算結果 Max_4
    SXLOSF4_CALCAVE As Double       ' OSF４SXL計算結果 Ave_4
    SXLBMD_SMPPOS As Integer        ' BMDｻﾝﾌﾟﾙ測定位置（SXL位置情報）
    SXLBMD1_KKSP As String * 3      ' BMD1結晶欠陥測定位置
    SXLBMD1_NETU As String * 2      ' BMD1熱処理法
    SXLBMD1_KKSET As String * 3     ' BMD1結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLBMD1_CALCMAX As Double       ' BMD1SXL計算結果 Max
    SXLBMD1_CALCAVE As Double       ' BMD1SXL計算結果 Ave
    SXLBMD2_KKSP As String * 3      ' BMD２結晶欠陥測定位置
    SXLBMD2_NETU As String * 2      ' BMD２熱処理法
    SXLBMD2_KKSET As String * 3     ' BMD２結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLBMD2_CALCMAX As Double       ' BMD２SXL計算結果 Max
    SXLBMD2_CALCAVE As Double       ' BMD２SXL計算結果 Ave
    SXLBMD3_KKSP As String * 3      ' BMD３結晶欠陥測定位置
    SXLBMD3_NETU As String * 2      ' BMD３熱処理法
    SXLBMD3_KKSET As String * 3     ' BMD３結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLBMD3_CALCMAX As Double       ' BMD３SXL計算結果 Max
    SXLBMD3_CALCAVE As Double       ' BMD３SXL計算結果 Ave
    SXLGD_SMPPOS As Integer         ' GDｻﾝﾌﾟﾙ測定位置（SXL位置情報）
    SXLGD_MSRSDEN As Integer        ' SXLGD_測定結果 Den
    SXLGD_MSRSLDL As Integer        ' SXLGD_測定結果 L/DL
    SXLGD_MSRSDVD2 As Integer       ' SXLGD_測定結果 DVD2
    SXLLT_SMPPOS As Integer         ' LTｻﾝﾌﾟﾙ測定位置（SXL位置情報）
    SXLLT_MEASPEAK As Integer       ' SXLLT_測定値 ピーク値
    SXLLT_CALCMEAS As Integer       ' SXLLT_計算結果
    WFOI_SMPPOS As Integer          ' WFOIｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFOI_NETSU As String * 2        ' WFOI_熱処理条件
    WFOI_ET As String * 3           ' WFOI_エッチング条件
    WFOI_MES As String * 3          ' WFOI_計測方法
    WFOI_MESDATA1 As Double         ' WFOI_測定データその１
    WFOI_MESDATA2 As Double         ' WFOI_測定データその２
    WFOI_MESDATA3 As Double         ' WFOI_測定データその３
    WFOI_MESDATA4 As Double         ' WFOI_測定データその４
    WFOI_MESDATA5 As Double         ' WFOI_測定データその５
    WFOI_MESDATA6 As Double         ' WFOI_測定データその６
    WFOI_MESDATA7 As Double         ' WFOI_測定データその７
    WFOI_MESDATA8 As Double         ' WFOI_測定データその８
    WFOI_MESDATA9 As Double         ' WFOI_測定データその９
    WFOI_MESDATA10 As Double        ' WFOI_測定データその１０
    WFOI_ORG As Double              ' WFOI_ORG計算結果
    WFRS_SMPPOS As Integer          ' WFRSｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFRS_NETSU As String * 2        ' WFRS_熱処理条件
    WFRS_ET As String * 3           ' WFRS_エッチング条件
    WFRS_MES As String * 3          ' WFRS_計測方法
    WFRS_MESDATA1 As Double         ' WFRS_測定データその１
    WFRS_MESDATA2 As Double         ' WFRS_測定データその２
    WFRS_MESDATA3 As Double         ' WFRS_測定データその３
    WFRS_MESDATA4 As Double         ' WFRS_測定データその４
    WFRS_MESDATA5 As Double         ' WFRS_測定データその５
    WFRS_RRG As Double              ' WFRS_RRG計算結果
    WFDOI_SMPPOS As Integer         ' WFDOIｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）　number(4)
    WFDOI_NETU_1 As String * 2      ' WFDOI_熱処理条件_1
    WFDOI_MES_1 As String * 3       ' WFDOI_計測方法_1
    WFDOI_MESDATA1_1 As Double      ' WFDOI_(ｲﾆｼｬﾙOi-AfterOi）１_1
    WFDOI_MESDATA2_1 As Double      ' WFDOI_(ｲﾆｼｬﾙOi-AfterOi)２_1
    WFDOI_MESDATA3_1 As Double      ' WFDOI_(ｲﾆｼｬﾙOi-AfterOi)３_1
    WFDOI_NETU_2 As String * 2      ' WFDOI_熱処理条件_２
    WFDOI_MES_2 As String * 3       ' WFDOI_計測方法_２
    WFDOI_MESDATA1_2 As Double      ' WFDOI_(ｲﾆｼｬﾙOi-AfterOi）１_２
    WFDOI_MESDATA2_2 As Double      ' WFDOI_(ｲﾆｼｬﾙOi-AfterOi)２_２
    WFDOI_MESDATA3_2 As Double      ' WFDOI_(ｲﾆｼｬﾙOi-AfterOi)３_２
    WFDOI_NETU_3 As String * 2      ' WFDOI_熱処理条件_３
    WFDOI_MES_3 As String * 3       ' WFDOI_計測方法_３
    WFDOI_MESDATA1_3 As Double      ' WFDOI_(ｲﾆｼｬﾙOi-AfterOi）１_３
    WFDOI_MESDATA2_3 As Double      ' WFDOI_(ｲﾆｼｬﾙOi-AfterOi)２_３
    WFDOI_MESDATA3_3 As Double      ' WFDOI_(ｲﾆｼｬﾙOi-AfterOi)３_３
    WFOSF1_SMPPOS As Integer        ' WFOSF1ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFOSF1_NETSU As String * 2      ' WFOSF1_熱処理条件
    WFOSF1_ET As String * 3         ' WFOSF1_エッチング条件
    WFOSF1_MES As String * 3        ' WFOSF1_計測方法
    WFOSF1_MAX As Double            ' WFOSF1_判定時のMAX値_1
    WFOSF1_AVE As Double            ' WFOSF1_判定時のAVE値_1
    WFOSF2_SMPPOS As Integer        ' WFOSF２ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）　number(4)
    WFOSF2_NETSU As String * 2      ' WFOSF2_熱処理条件_２
    WFOSF2_ET As String * 3         ' WFOSF2_エッチング条件_２
    WFOSF2_MES As String * 3        ' WFOSF2_計測方法_２
    WFOSF2_MAX As Double            ' WFOSF2_判定時のMAX値_２
    WFOSF2_AVE As Double            ' WFOSF2_判定時のAVE値_２
    WFOSF3_SMPPOS As Integer        ' WFOSF３ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFOSF3_NETSU As String * 2      ' WFOSF3_熱処理条件_３
    WFOSF3_ET As String * 3         ' WFOSF3_エッチング条件_３
    WFOSF3_MES As String * 3        ' WFOSF3_計測方法_３
    WFOSF3_MAX As Double            ' WFOSF3_判定時のMAX値_３
    WFOSF3_AVE As Double            ' WFOSF3_判定時のAVE値_３
    WFOSF4_SMPPOS As Integer        ' WFOSF４ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFOSF4_NETSU As String * 2      ' WFOSF4_熱処理条件_４
    WFOSF4_ET As String * 3         ' WFOSF4_エッチング条件_４
    WFOSF4_MES As String * 3        ' WFOSF4_計測方法_４
    WFOSF4_MAX As Double            ' WFOSF4_判定時のMAX値_４
    WFOSF4_AVE As Double            ' WFOSF4_判定時のAVE値_４
    WFBMD1_SMPPOS As Integer        ' WFBMD1ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFBMD1_NETSU As String * 2      ' WFBMD1_熱処理条件_1
    WFBMD1_ET As String * 3         ' WFBMD1_エッチング条件_1
    WFBMD1_MES As String * 3        ' WFBMD1_計測方法_1
    WFBMD1_MAX As Double            ' WFBMD1_判定時のMAX値_1
    WFBMD1_AVE As Double            ' WFBMD1_判定時のAVE値_1
    WFBMD2_SMPPOS As Integer        ' WFBMD２ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFBMD2_NETSU As String * 2      ' WFBMD2_熱処理条件_２
    WFBMD2_ET As String * 3         ' WFBMD2_エッチング条件_２
    WFBMD2_MES As String * 3        ' WFBMD2_計測方法_２
    WFBMD2_MAX As Double            ' WFBMD2_判定時のMAX値_２
    WFBMD2_AVE As Double            ' WFBMD2_判定時のAVE値_２
    WFBMD3_SMPPOS As Integer        ' WFBMD３ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFBMD3_NETSU As String * 2      ' WFBMD3_熱処理条件_３
    WFBMD3_ET As String * 3         ' WFBMD3_エッチング条件_３
    WFBMD3_MES As String * 3        ' WFBMD3_計測方法_３
    WFBMD3_MAX As Double            ' WFBMD3_判定時のMAX値_３
    WFBMD3_AVE As Double            ' WFBMD3_判定時のAVE値_３
    WFDSOD_SMPPOS As Integer        ' WFDSODｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFDSOD_NETSU As String * 2      ' WFDSOD_熱処理条件
    WFDSOD_ET As String * 3         ' WFDSOD_エッチング条件
    WFDSOD_MES As String * 3        ' WFDSOD_計測方法
    WFDSOD_TOTAL As Integer         ' WFDSOD_判定時のTOTAL値
    WFSPV_SMPPOS As Integer         ' WFSPVｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFSPV_NETSU As String * 2       ' WFSVP_熱処理条件
    WFSPV_ET As String * 3          ' WFSPV_エッチング条件
    WFSPV_MES As String * 3         ' WFSPV_計測方法
    WFSPV_KST_MAX As Double         ' WFSPV_拡散長判定時のＭＡＸ値
    WFSPV_KST_AVE As Double         ' WFSPV_拡散長判定時のAVE値
    WFSPV_KST_MIN As Double         ' WFSPV_拡散長判定時のMIN値
    WFSPV_FE_MAX As Double          ' WFSPV_Fe濃度判定時のMAX値
    WFSPV_FE_AVE As Double          ' WFSPV_Fe濃度判定時のAVE値
    WFSPV_FE_MIN As Double          ' WFSPV_Fe濃度判定時のMIN値
    WFDZ_SMPPOS As Integer          ' WFDZｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFDZ_NETSU As String * 2        ' WFDZ_熱処理条件
    WFDZ_ET As String * 3           ' WFDZ_エッチング条件
    WFDZ_MES As String * 3          ' WFDZ_計測方法
    WFDZ_MAX As Double              ' WFDZ_判定時のMAX値_
    WFDZ_AVE As Double              ' WFDZ_判定時のAVE値
    REGDATE As Date                 ' 登録日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
    
    SXLOSF1_PTNJUDGRES  As String * 1   ' OSF1パターン判定結果  '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
End Type


' SXL測定点デ－タ
Public Type typ_TBCMX002
    SXLID As String * 13            ' SXLID
    FROMTOKBN As String * 1         ' FROMTO区分
    SAMPLE_FROM As String * 16      ' サンプルID (From)
    SAMPLE_TO As String * 16        ' サンプルID (To)
    BLOCKID As String * 12          ' ブロックID
    CRYNUM As String * 12           ' 結晶番号
    SXLDECDATE As Date              ' SXL-ID確定日付
    PLUPDATE As Date                ' 引上日付
    INGOTPOS As Integer             ' 結晶内開始位置
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    UPLENGTH As Integer             ' 引上げ長さ
    SXLPOS As Integer               ' SXL位置
    SXLLENGTH As Integer            ' SXL-ID確定長さ
    SXLWAFERCNT As Integer          ' SXL-ID確定時の枚数
    FREELENG As Integer             ' フリー長
    SAMPID_1 As String * 16         ' サンプルID 1
    SXLOSF1_SMPPOS As Integer       ' SXLOSFｻﾝﾌﾟﾙ測定位置（SXL位置情報）
    SXLOSF1_KKSP As String * 3      ' SXLOSF1結晶欠陥測定位置
    SXLOSF1_NETU As String * 2      ' SXLOSF1熱処理法
    SXLOSF1_KKSET As String * 3     ' SXLOSF1結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLOSF1_MEAS1 As Integer        ' SXLOSF1測定点１
    SXLOSF1_MEAS2 As Integer        ' SXLOSF1測定点2
    SXLOSF1_MEAS3 As Integer        ' SXLOSF1測定点3
    SXLOSF1_MEAS4 As Integer        ' SXLOSF1測定点4
    SXLOSF1_MEAS5 As Integer        ' SXLOSF1測定点5
    SXLOSF1_MEAS6 As Integer        ' SXLOSF1測定点6
    SXLOSF1_MEAS7 As Integer        ' SXLOSF1測定点7
    SXLOSF1_MEAS8 As Integer        ' SXLOSF1測定点8
    SXLOSF1_MEAS9 As Integer        ' SXLOSF1測定点9
    SXLOSF1_MEAS10 As Integer       ' SXLOSF1測定点10
    SXLOSF1_MEAS11 As Integer       ' SXLOSF1測定点11
    SXLOSF1_MEAS12 As Integer       ' SXLOSF1測定点12
    SXLOSF1_MEAS13 As Integer       ' SXLOSF1測定点13
    SXLOSF1_MEAS14 As Integer       ' SXLOSF1測定点14
    SXLOSF1_MEAS15 As Integer       ' SXLOSF1測定点15
    SXLOSF1_MEAS16 As Integer       ' SXLOSF1測定点16
    SXLOSF1_MEAS17 As Integer       ' SXLOSF1測定点17
    SXLOSF1_MEAS18 As Integer       ' SXLOSF1測定点18
    SXLOSF1_MEAS19 As Integer       ' SXLOSF1測定点19
    SXLOSF1_MEAS20 As Integer       ' SXLOSF1測定点20
    SXLOSF2_KKSP As String * 3      ' SXLOSF２結晶欠陥測定位置
    SXLOSF2_NETU As String * 2      ' SXLOSF２熱処理法
    SXLOSF2_KKSET As String * 3     ' SXLOSF２結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLOSF2_MEAS1 As Integer        ' SXLOSF2測定点１
    SXLOSF2_MEAS2 As Integer        ' SXLOSF2測定点2
    SXLOSF2_MEAS3 As Integer        ' SXLOSF2測定点3
    SXLOSF2_MEAS4 As Integer        ' SXLOSF2測定点4
    SXLOSF2_MEAS5 As Integer        ' SXLOSF2測定点5
    SXLOSF2_MEAS6 As Integer        ' SXLOSF2測定点6
    SXLOSF2_MEAS7 As Integer        ' SXLOSF2測定点7
    SXLOSF2_MEAS8 As Integer        ' SXLOSF2測定点8
    SXLOSF2_MEAS9 As Integer        ' SXLOSF2測定点9
    SXLOSF2_MEAS10 As Integer       ' SXLOSF2測定点10
    SXLOSF2_MEAS11 As Integer       ' SXLOSF2測定点11
    SXLOSF2_MEAS12 As Integer       ' SXLOSF2測定点12
    SXLOSF2_MEAS13 As Integer       ' SXLOSF2測定点13
    SXLOSF2_MEAS14 As Integer       ' SXLOSF2測定点14
    SXLOSF2_MEAS15 As Integer       ' SXLOSF2測定点15
    SXLOSF2_MEAS16 As Integer       ' SXLOSF2測定点16
    SXLOSF2_MEAS17 As Integer       ' SXLOSF2測定点17
    SXLOSF2_MEAS18 As Integer       ' SXLOSF2測定点18
    SXLOSF2_MEAS19 As Integer       ' SXLOSF2測定点19
    SXLOSF2_MEAS20 As Integer       ' SXLOSF2測定点20
    SXLOSF3_KKSP As String * 3      ' SXLOSF３結晶欠陥測定位置
    SXLOSF3_NETU As String * 2      ' SXLOSF３熱処理法
    SXLOSF3_KKSET As String * 3     ' SXLOSF３結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLOSF3_MEAS1 As Integer        ' SXLOSF3測定点１
    SXLOSF3_MEAS2 As Integer        ' SXLOSF3測定点2
    SXLOSF3_MEAS3 As Integer        ' SXLOSF3測定点3
    SXLOSF3_MEAS4 As Integer        ' SXLOSF3測定点4
    SXLOSF3_MEAS5 As Integer        ' SXLOSF3測定点5
    SXLOSF3_MEAS6 As Integer        ' SXLOSF3測定点6
    SXLOSF3_MEAS7 As Integer        ' SXLOSF3測定点7
    SXLOSF3_MEAS8 As Integer        ' SXLOSF3測定点8
    SXLOSF3_MEAS9 As Integer        ' SXLOSF3測定点9
    SXLOSF3_MEAS10 As Integer       ' SXLOSF3測定点10
    SXLOSF3_MEAS11 As Integer       ' SXLOSF3測定点11
    SXLOSF3_MEAS12 As Integer       ' SXLOSF3測定点12
    SXLOSF3_MEAS13 As Integer       ' SXLOSF3測定点13
    SXLOSF3_MEAS14 As Integer       ' SXLOSF3測定点14
    SXLOSF3_MEAS15 As Integer       ' SXLOSF3測定点15
    SXLOSF3_MEAS16 As Integer       ' SXLOSF3測定点16
    SXLOSF3_MEAS17 As Integer       ' SXLOSF3測定点17
    SXLOSF3_MEAS18 As Integer       ' SXLOSF3測定点18
    SXLOSF3_MEAS19 As Integer       ' SXLOSF3測定点19
    SXLOSF3_MEAS20 As Integer       ' SXLOSF3測定点20
    SXLOSF4_KKSP As String * 3      ' SXLOSF４結晶欠陥測定位置
    SXLOSF4_NETU As String * 2      ' SXLOSF４熱処理法
    SXLOSF4_KKSET As String * 3     ' SXLOSF４結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLOSF4_MEAS1 As Integer        ' SXLOSF4測定点１
    SXLOSF4_MEAS2 As Integer        ' SXLOSF4測定点2
    SXLOSF4_MEAS3 As Integer        ' SXLOSF4測定点3
    SXLOSF4_MEAS4 As Integer        ' SXLOSF4測定点4
    SXLOSF4_MEAS5 As Integer        ' SXLOSF4測定点5
    SXLOSF4_MEAS6 As Integer        ' SXLOSF4測定点6
    SXLOSF4_MEAS7 As Integer        ' SXLOSF4測定点7
    SXLOSF4_MEAS8 As Integer        ' SXLOSF4測定点8
    SXLOSF4_MEAS9 As Integer        ' SXLOSF4測定点9
    SXLOSF4_MEAS10 As Integer       ' SXLOSF4測定点10
    SXLOSF4_MEAS11 As Integer       ' SXLOSF4測定点11
    SXLOSF4_MEAS12 As Integer       ' SXLOSF4測定点12
    SXLOSF4_MEAS13 As Integer       ' SXLOSF4測定点13
    SXLOSF4_MEAS14 As Integer       ' SXLOSF4測定点14
    SXLOSF4_MEAS15 As Integer       ' SXLOSF4測定点15
    SXLOSF4_MEAS16 As Integer       ' SXLOSF4測定点16
    SXLOSF4_MEAS17 As Integer       ' SXLOSF4測定点17
    SXLOSF4_MEAS18 As Integer       ' SXLOSF4測定点18
    SXLOSF4_MEAS19 As Integer       ' SXLOSF4測定点19
    SXLOSF4_MEAS20 As Integer       ' SXLOSF4測定点20
    SXLBMD_SMPPOS As Integer        ' SXLBMDｻﾝﾌﾟﾙ測定位置（SXL位置情報）
    SXLBMD1_KKSP As String * 3      ' SXLBMD1結晶欠陥測定位置
    SXLBMD1_NETU As String * 2      ' SXLBMD1熱処理法
    SXLBMD1_KKSET As String * 3     ' SXLBMD1結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLBMD1_MEAS1 As Integer        ' SXLBMD1測定点１
    SXLBMD1_MEAS2 As Integer        ' SXLBMD1測定点2
    SXLBMD1_MEAS3 As Integer        ' SXLBMD1測定点3
    SXLBMD1_MEAS4 As Integer        ' SXLBMD1測定点4
    SXLBMD1_MEAS5 As Integer        ' SXLBMD1測定点5
    SXLBMD2_KKSP As String * 3      ' SXLBMD２結晶欠陥測定位置
    SXLBMD2_NETU As String * 2      ' SXLBMD２熱処理法
    SXLBMD2_KKSET As String * 3     ' SXLBMD２結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLBMD2_MEAS1 As Integer        ' SXLBMD2測定点１
    SXLBMD2_MEAS2 As Integer        ' SXLBMD2測定点2
    SXLBMD2_MEAS3 As Integer        ' SXLBMD2測定点3
    SXLBMD2_MEAS4 As Integer        ' SXLBMD2測定点4
    SXLBMD2_MEAS5 As Integer        ' SXLBMD2測定点5
    SXLBMD3_KKSP As String * 3      ' SXLBMD３結晶欠陥測定位置
    SXLBMD3_NETU As String * 2      ' SXLBMD３熱処理法
    SXLBMD3_KKSET As String * 3     ' SXLBMD３結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLBMD3_MEAS1 As Integer        ' SXLBMD3測定点１
    SXLBMD3_MEAS2 As Integer        ' SXLBMD3測定点2
    SXLBMD3_MEAS3 As Integer        ' SXLBMD3測定点3
    SXLBMD3_MEAS4 As Integer        ' SXLBMD3測定点4
    SXLBMD3_MEAS5 As Integer        ' SXLBMD3測定点5
    SXLGD_SMPPOS As Integer         ' SXLGDｻﾝﾌﾟﾙ測定位置（SXL位置情報）
    SXLGD_MS01LDL1 As Integer       ' SXLGD_測定値01 L/DL1
    SXLGD_MS01LDL2 As Integer       ' SXLGD_測定値01 L/DL2
    SXLGD_MS01LDL3 As Integer       ' SXLGD_測定値01 L/DL3
    SXLGD_MS01LDL4 As Integer       ' SXLGD_測定値01 L/DL4
    SXLGD_MS01LDL5 As Integer       ' SXLGD_測定値01 L/DL5
    SXLGD_MS01DEN1 As Integer       ' SXLGD_測定値01 Den1
    SXLGD_MS01DEN2 As Integer       ' SXLGD_測定値01 Den2
    SXLGD_MS01DEN3 As Integer       ' SXLGD_測定値01 Den3
    SXLGD_MS01DEN4 As Integer       ' SXLGD_測定値01 Den4
    SXLGD_MS01DEN5 As Integer       ' SXLGD_測定値01 Den5
    SXLGD_MS02LDL1 As Integer       ' SXLGD_測定値02 L/DL1
    SXLGD_MS02LDL2 As Integer       ' SXLGD_測定値02 L/DL2
    SXLGD_MS02LDL3 As Integer       ' SXLGD_測定値02 L/DL3
    SXLGD_MS02LDL4 As Integer       ' SXLGD_測定値02 L/DL4
    SXLGD_MS02LDL5 As Integer       ' SXLGD_測定値02 L/DL5
    SXLGD_MS02DEN1 As Integer       ' SXLGD_測定値02 Den1
    SXLGD_MS02DEN2 As Integer       ' SXLGD_測定値02 Den2
    SXLGD_MS02DEN3 As Integer       ' SXLGD_測定値02 Den3
    SXLGD_MS02DEN4 As Integer       ' SXLGD_測定値02 Den4
    SXLGD_MS02DEN5 As Integer       ' SXLGD_測定値02 Den5
    SXLGD_MS03LDL1 As Integer       ' SXLGD_測定値03 L/DL1
    SXLGD_MS03LDL2 As Integer       ' SXLGD_測定値03 L/DL2
    SXLGD_MS03LDL3 As Integer       ' SXLGD_測定値03 L/DL3
    SXLGD_MS03LDL4 As Integer       ' SXLGD_測定値03 L/DL4
    SXLGD_MS03LDL5 As Integer       ' SXLGD_測定値03 L/DL5
    SXLGD_MS03DEN1 As Integer       ' SXLGD_測定値03 Den1
    SXLGD_MS03DEN2 As Integer       ' SXLGD_測定値03 Den2
    SXLGD_MS03DEN3 As Integer       ' SXLGD_測定値03 Den3
    SXLGD_MS03DEN4 As Integer       ' SXLGD_測定値03 Den4
    SXLGD_MS03DEN5 As Integer       ' SXLGD_測定値03 Den5
    SXLGD_MS04LDL1 As Integer       ' SXLGD_測定値04 L/DL1
    SXLGD_MS04LDL2 As Integer       ' SXLGD_測定値04 L/DL2
    SXLGD_MS04LDL3 As Integer       ' SXLGD_測定値04 L/DL3
    SXLGD_MS04LDL4 As Integer       ' SXLGD_測定値04 L/DL4
    SXLGD_MS04LDL5 As Integer       ' SXLGD_測定値04 L/DL5
    SXLGD_MS04DEN1 As Integer       ' SXLGD_測定値04 Den1
    SXLGD_MS04DEN2 As Integer       ' SXLGD_測定値04 Den2
    SXLGD_MS04DEN3 As Integer       ' SXLGD_測定値04 Den3
    SXLGD_MS04DEN4 As Integer       ' SXLGD_測定値04 Den4
    SXLGD_MS04DEN5 As Integer       ' SXLGD_測定値04 Den5
    SXLGD_MS05LDL1 As Integer       ' SXLGD_測定値05 L/DL1
    SXLGD_MS05LDL2 As Integer       ' SXLGD_測定値05 L/DL2
    SXLGD_MS05LDL3 As Integer       ' SXLGD_測定値05 L/DL3
    SXLGD_MS05LDL4 As Integer       ' SXLGD_測定値05 L/DL4
    SXLGD_MS05LDL5 As Integer       ' SXLGD_測定値05 L/DL5
    SXLGD_MS05DEN1 As Integer       ' SXLGD_測定値05 Den1
    SXLGD_MS05DEN2 As Integer       ' SXLGD_測定値05 Den2
    SXLGD_MS05DEN3 As Integer       ' SXLGD_測定値05 Den3
    SXLGD_MS05DEN4 As Integer       ' SXLGD_測定値05 Den4
    SXLGD_MS05DEN5 As Integer       ' SXLGD_測定値05 Den5
    SXLGD_MS06LDL1 As Integer       ' SXLGD_測定値06 L/DL1
    SXLGD_MS06LDL2 As Integer       ' SXLGD_測定値06 L/DL2
    SXLGD_MS06LDL3 As Integer       ' SXLGD_測定値06 L/DL3
    SXLGD_MS06LDL4 As Integer       ' SXLGD_測定値06 L/DL4
    SXLGD_MS06LDL5 As Integer       ' SXLGD_測定値06 L/DL5
    SXLGD_MS06DEN1 As Integer       ' SXLGD_測定値06 Den1
    SXLGD_MS06DEN2 As Integer       ' SXLGD_測定値06 Den2
    SXLGD_MS06DEN3 As Integer       ' SXLGD_測定値06 Den3
    SXLGD_MS06DEN4 As Integer       ' SXLGD_測定値06 Den4
    SXLGD_MS06DEN5 As Integer       ' SXLGD_測定値06 Den5
    SXLGD_MS07LDL1 As Integer       ' SXLGD_測定値07 L/DL1
    SXLGD_MS07LDL2 As Integer       ' SXLGD_測定値07 L/DL2
    SXLGD_MS07LDL3 As Integer       ' SXLGD_測定値07 L/DL3
    SXLGD_MS07LDL4 As Integer       ' SXLGD_測定値07 L/DL4
    SXLGD_MS07LDL5 As Integer       ' SXLGD_測定値07 L/DL5
    SXLGD_MS07DEN1 As Integer       ' SXLGD_測定値07 Den1
    SXLGD_MS07DEN2 As Integer       ' SXLGD_測定値07 Den2
    SXLGD_MS07DEN3 As Integer       ' SXLGD_測定値07 Den3
    SXLGD_MS07DEN4 As Integer       ' SXLGD_測定値07 Den4
    SXLGD_MS07DEN5 As Integer       ' SXLGD_測定値07 Den5
    SXLGD_MS08LDL1 As Integer       ' SXLGD_測定値08 L/DL1
    SXLGD_MS08LDL2 As Integer       ' SXLGD_測定値08 L/DL2
    SXLGD_MS08LDL3 As Integer       ' SXLGD_測定値08 L/DL3
    SXLGD_MS08LDL4 As Integer       ' SXLGD_測定値08 L/DL4
    SXLGD_MS08LDL5 As Integer       ' SXLGD_測定値08 L/DL5
    SXLGD_MS08DEN1 As Integer       ' SXLGD_測定値08 Den1
    SXLGD_MS08DEN2 As Integer       ' SXLGD_測定値08 Den2
    SXLGD_MS08DEN3 As Integer       ' SXLGD_測定値08 Den3
    SXLGD_MS08DEN4 As Integer       ' SXLGD_測定値08 Den4
    SXLGD_MS08DEN5 As Integer       ' SXLGD_測定値08 Den5
    SXLGD_MS09LDL1 As Integer       ' SXLGD_測定値09 L/DL1
    SXLGD_MS09LDL2 As Integer       ' SXLGD_測定値09 L/DL2
    SXLGD_MS09LDL3 As Integer       ' SXLGD_測定値09 L/DL3
    SXLGD_MS09LDL4 As Integer       ' SXLGD_測定値09 L/DL4
    SXLGD_MS09LDL5 As Integer       ' SXLGD_測定値09 L/DL5
    SXLGD_MS09DEN1 As Integer       ' SXLGD_測定値09 Den1
    SXLGD_MS09DEN2 As Integer       ' SXLGD_測定値09 Den2
    SXLGD_MS09DEN3 As Integer       ' SXLGD_測定値09 Den3
    SXLGD_MS09DEN4 As Integer       ' SXLGD_測定値09 Den4
    SXLGD_MS09DEN5 As Integer       ' SXLGD_測定値09 Den5
    SXLGD_MS10LDL1 As Integer       ' SXLGD_測定値10 L/DL1
    SXLGD_MS10LDL2 As Integer       ' SXLGD_測定値10 L/DL2
    SXLGD_MS10LDL3 As Integer       ' SXLGD_測定値10 L/DL3
    SXLGD_MS10LDL4 As Integer       ' SXLGD_測定値10 L/DL4
    SXLGD_MS10LDL5 As Integer       ' SXLGD_測定値10 L/DL5
    SXLGD_MS10DEN1 As Integer       ' SXLGD_測定値10 Den1
    SXLGD_MS10DEN2 As Integer       ' SXLGD_測定値10 Den2
    SXLGD_MS10DEN3 As Integer       ' SXLGD_測定値10 Den3
    SXLGD_MS10DEN4 As Integer       ' SXLGD_測定値10 Den4
    SXLGD_MS10DEN5 As Integer       ' SXLGD_測定値10 Den5
    SXLGD_MS11LDL1 As Integer       ' SXLGD_測定値11 L/DL1
    SXLGD_MS11LDL2 As Integer       ' SXLGD_測定値11 L/DL2
    SXLGD_MS11LDL3 As Integer       ' SXLGD_測定値11 L/DL3
    SXLGD_MS11LDL4 As Integer       ' SXLGD_測定値11 L/DL4
    SXLGD_MS11LDL5 As Integer       ' SXLGD_測定値11 L/DL5
    SXLGD_MS11DEN1 As Integer       ' SXLGD_測定値11 Den1
    SXLGD_MS11DEN2 As Integer       ' SXLGD_測定値11 Den2
    SXLGD_MS11DEN3 As Integer       ' SXLGD_測定値11 Den3
    SXLGD_MS11DEN4 As Integer       ' SXLGD_測定値11 Den4
    SXLGD_MS11DEN5 As Integer       ' SXLGD_測定値11 Den5
    SXLGD_MS12LDL1 As Integer       ' SXLGD_測定値12 L/DL1
    SXLGD_MS12LDL2 As Integer       ' SXLGD_測定値12 L/DL2
    SXLGD_MS12LDL3 As Integer       ' SXLGD_測定値12 L/DL3
    SXLGD_MS12LDL4 As Integer       ' SXLGD_測定値12 L/DL4
    SXLGD_MS12LDL5 As Integer       ' SXLGD_測定値12 L/DL5
    SXLGD_MS12DEN1 As Integer       ' SXLGD_測定値12 Den1
    SXLGD_MS12DEN2 As Integer       ' SXLGD_測定値12 Den2
    SXLGD_MS12DEN3 As Integer       ' SXLGD_測定値12 Den3
    SXLGD_MS12DEN4 As Integer       ' SXLGD_測定値12 Den4
    SXLGD_MS12DEN5 As Integer       ' SXLGD_測定値12 Den5
    SXLGD_MS13LDL1 As Integer       ' SXLGD_測定値13 L/DL1
    SXLGD_MS13LDL2 As Integer       ' SXLGD_測定値13 L/DL2
    SXLGD_MS13LDL3 As Integer       ' SXLGD_測定値13 L/DL3
    SXLGD_MS13LDL4 As Integer       ' SXLGD_測定値13 L/DL4
    SXLGD_MS13LDL5 As Integer       ' SXLGD_測定値13 L/DL5
    SXLGD_MS13DEN1 As Integer       ' SXLGD_測定値13 Den1
    SXLGD_MS13DEN2 As Integer       ' SXLGD_測定値13 Den2
    SXLGD_MS13DEN3 As Integer       ' SXLGD_測定値13 Den3
    SXLGD_MS13DEN4 As Integer       ' SXLGD_測定値13 Den4
    SXLGD_MS13DEN5 As Integer       ' SXLGD_測定値13 Den5
    SXLGD_MS14LDL1 As Integer       ' SXLGD_測定値14 L/DL1
    SXLGD_MS14LDL2 As Integer       ' SXLGD_測定値14 L/DL2
    SXLGD_MS14LDL3 As Integer       ' SXLGD_測定値14 L/DL3
    SXLGD_MS14LDL4 As Integer       ' SXLGD_測定値14 L/DL4
    SXLGD_MS14LDL5 As Integer       ' SXLGD_測定値14 L/DL5
    SXLGD_MS14DEN1 As Integer       ' SXLGD_測定値14 Den1
    SXLGD_MS14DEN2 As Integer       ' SXLGD_測定値14 Den2
    SXLGD_MS14DEN3 As Integer       ' SXLGD_測定値14 Den3
    SXLGD_MS14DEN4 As Integer       ' SXLGD_測定値14 Den4
    SXLGD_MS14DEN5 As Integer       ' SXLGD_測定値14 Den5
    SXLGD_MS15LDL1 As Integer       ' SXLGD_測定値15 L/DL1
    SXLGD_MS15LDL2 As Integer       ' SXLGD_測定値15 L/DL2
    SXLGD_MS15LDL3 As Integer       ' SXLGD_測定値15 L/DL3
    SXLGD_MS15LDL4 As Integer       ' SXLGD_測定値15 L/DL4
    SXLGD_MS15LDL5 As Integer       ' SXLGD_測定値15 L/DL5
    SXLGD_MS15DEN1 As Integer       ' SXLGD_測定値15 Den1
    SXLGD_MS15DEN2 As Integer       ' SXLGD_測定値15 Den2
    SXLGD_MS15DEN3 As Integer       ' SXLGD_測定値15 Den3
    SXLGD_MS15DEN4 As Integer       ' SXLGD_測定値15 Den4
    SXLGD_MS15DEN5 As Integer       ' SXLGD_測定値15 Den5
    SXLT_SMPPOS As Integer          ' SXLLTｻﾝﾌﾟﾙ測定位置（SXL位置情報）
    SXLLT_MEASPEAK As Integer       ' SXLLT_測定値 ピーク値
    SXLLT_MEAS1 As Integer          ' SXLLT_測定値1
    SXLLT_MEAS2 As Integer          ' SXLLT_測定値2
    SXLLT_MEAS3 As Integer          ' SXLLT_測定値3
    SXLLT_MEAS4 As Integer          ' SXLLT_測定値4
    SXLLT_MEAS5 As Integer          ' SXLLT_測定値5
    WFDOI_SMPPOS As Integer         ' WFDOIｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）　number(4)
    WFDOI1_NETSU As String * 2      ' WFDOI-1_熱処理条件
    WFDOI1_MES As String * 3        ' WFDOI-1_計測方法
    WFDOI1_MESDATA1 As String * 10  ' WFDOI-1_測定値１
    WFDOI1_MESDATA2 As String * 10  ' WFDOI-1_測定値2
    WFDOI1_MESDATA3 As String * 10  ' WFDOI-1_測定値3
    WFDOI1_MESDATA4 As String * 10  ' WFDOI-1_測定値4
    WFDOI1_MESDATA5 As String * 10  ' WFDOI-1_測定値5
    WFDOI1_MESDATA6 As String * 10  ' WFDOI-1_測定値6
    WFDOI1_MESDATA7 As String * 10  ' WFDOI-1_測定値7
    WFDOI1_MESDATA8 As String * 10  ' WFDOI-1_測定値8
    WFDOI1_MESDATA9 As String * 10  ' WFDOI-1_測定値9
    WFDOI1_MESDATA10 As String * 10 ' WFDOI-1_測定値10
    WFDOI1_MESDATA11 As String * 10 ' WFDOI-1_測定値11
    WFDOI1_MESDATA12 As String * 10 ' WFDOI-1_測定値12
    WFDOI1_MESDATA13 As String * 10 ' WFDOI-1_測定値13
    WFDOI1_MESDATA14 As String * 10 ' WFDOI-1_測定値14
    WFDOI1_MESDATA15 As String * 10 ' WFDOI-1_測定値15
    WFDOI2_NETSU As String * 2      ' WFDOI-2_熱処理条件
    WFDOI2_MES As String * 3        ' WFDOI-2_計測方法
    WFDOI2_MESDATA1 As String * 10  ' WFDOI-2_測定値１
    WFDOI2_MESDATA2 As String * 10  ' WFDOI-2_測定値2
    WFDOI2_MESDATA3 As String * 10  ' WFDOI-2_測定値3
    WFDOI2_MESDATA4 As String * 10  ' WFDOI-2_測定値4
    WFDOI2_MESDATA5 As String * 10  ' WFDOI-2_測定値5
    WFDOI2_MESDATA6 As String * 10  ' WFDOI-2_測定値6
    WFDOI2_MESDATA7 As String * 10  ' WFDOI-2_測定値7
    WFDOI2_MESDATA8 As String * 10  ' WFDOI-2_測定値8
    WFDOI2_MESDATA9 As String * 10  ' WFDOI-2_測定値9
    WFDOI2_MESDATA10 As String * 10 ' WFDOI-2_測定値10
    WFDOI2_MESDATA11 As String * 10 ' WFDOI-2_測定値11
    WFDOI2_MESDATA12 As String * 10 ' WFDOI-2_測定値12
    WFDOI2_MESDATA13 As String * 10 ' WFDOI-2_測定値13
    WFDOI2_MESDATA14 As String * 10 ' WFDOI-2_測定値14
    WFDOI2_MESDATA15 As String * 10 ' WFDOI-2_測定値15
    WFDOI3_NETSU As String * 2      ' WFDOI-3_熱処理条件
    WFDOI3_MES As String * 3        ' WFDOI-3_計測方法
    WFDOI3_MESDATA1 As String * 10  ' WFDOI-3_測定値１
    WFDOI3_MESDATA2 As String * 10  ' WFDOI-3_測定値2
    WFDOI3_MESDATA3 As String * 10  ' WFDOI-3_測定値3
    WFDOI3_MESDATA4 As String * 10  ' WFDOI-3_測定値4
    WFDOI3_MESDATA5 As String * 10  ' WFDOI-3_測定値5
    WFDOI3_MESDATA6 As String * 10  ' WFDOI-3_測定値6
    WFDOI3_MESDATA7 As String * 10  ' WFDOI-3_測定値7
    WFDOI3_MESDATA8 As String * 10  ' WFDOI-3_測定値8
    WFDOI3_MESDATA9 As String * 10  ' WFDOI-3_測定値9
    WFDOI3_MESDATA10 As String * 10 ' WFDOI-3_測定値10
    WFDOI3_MESDATA11 As String * 10 ' WFDOI-3_測定値11
    WFDOI3_MESDATA12 As String * 10 ' WFDOI-3_測定値12
    WFDOI3_MESDATA13 As String * 10 ' WFDOI-3_測定値13
    WFDOI3_MESDATA14 As String * 10 ' WFDOI-3_測定値14
    WFDOI3_MESDATA15 As String * 10 ' WFDOI-3_測定値15
    WFOSF1_SMPPOS As Integer        ' WFOSF1ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFOSF1_NETSU As String * 2      ' WFOSF1_熱処理条件
    WFOSF1_ET As String * 3         ' WFOSF1_エッチング条件
    WFOSF1_MES As String * 3        ' WFOSF1_計測方法
    WFOSF1_DKAN As String * 10      ' WFOSF1_ＤＫアニール条件
    WFOSF1_MESDATA1 As String * 10  ' WFOSF1測定点１
    WFOSF1_MESDATA2 As String * 10  ' WFOSF1測定点2
    WFOSF1_MESDATA3 As String * 10  ' WFOSF1測定点3
    WFOSF1_MESDATA4 As String * 10  ' WFOSF1測定点4
    WFOSF1_MESDATA5 As String * 10  ' WFOSF1測定点5
    WFOSF1_MESDATA6 As String * 10  ' WFOSF1測定点6
    WFOSF1_MESDATA7 As String * 10  ' WFOSF1測定点7
    WFOSF1_MESDATA8 As String * 10  ' WFOSF1測定点8
    WFOSF1_MESDATA9 As String * 10  ' WFOSF1測定点9
    WFOSF1_MESDATA10 As String * 10 ' WFOSF1測定点10
    WFOSF1_MESDATA11 As String * 10 ' WFOSF1測定点11
    WFOSF1_MESDATA12 As String * 10 ' WFOSF1測定点12
    WFOSF1_MESDATA13 As String * 10 ' WFOSF1測定点13
    WFOSF1_MESDATA14 As String * 10 ' WFOSF1測定点14
    WFOSF1_MESDATA15 As String * 10 ' WFOSF1測定点15
    WFOSF2_SMPPOS As Integer        ' WFOSF２ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）　number(4)
    WFOSF2_NETSU As String * 2      ' WFOSF2_熱処理条件
    WFOSF2_ET As String * 3         ' WFOSF2_エッチング条件
    WFOSF2_MES As String * 3        ' WFOSF2_計測方法
    WFOSF2_DKAN As String * 10      ' WFOSF2_ＤＫアニール条件
    WFOSF2_MESDATA1 As String * 10  ' WFOSF2測定点１
    WFOSF2_MESDATA2 As String * 10  ' WFOSF2測定点2
    WFOSF2_MESDATA3 As String * 10  ' WFOSF2測定点3
    WFOSF2_MESDATA4 As String * 10  ' WFOSF2測定点4
    WFOSF2_MESDATA5 As String * 10  ' WFOSF2測定点5
    WFOSF2_MESDATA6 As String * 10  ' WFOSF2測定点6
    WFOSF2_MESDATA7 As String * 10  ' WFOSF2測定点7
    WFOSF2_MESDATA8 As String * 10  ' WFOSF2測定点8
    WFOSF2_MESDATA9 As String * 10  ' WFOSF2測定点9
    WFOSF2_MESDATA10 As String * 10 ' WFOSF2測定点10
    WFOSF2_MESDATA11 As String * 10 ' WFOSF2測定点11
    WFOSF2_MESDATA12 As String * 10 ' WFOSF2測定点12
    WFOSF2_MESDATA13 As String * 10 ' WFOSF2測定点13
    WFOSF2_MESDATA14 As String * 10 ' WFOSF2測定点14
    WFOSF2_MESDATA15 As String * 10 ' WFOSF2測定点15
    WFOSF3_SMPPOS As Integer        ' WFOSF３ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFOSF3_NETSU As String * 2      ' WFOSF3_熱処理条件
    WFOSF3_ET As String * 3         ' WFOSF3_エッチング条件
    WFOSF3_MES As String * 3        ' WFOSF3_計測方法
    WFOSF3_DKAN As String * 10      ' WFOSF3_ＤＫアニール条件
    WFOSF3_MESDATA1 As String * 10  ' WFOSF3測定点１
    WFOSF3_MESDATA2 As String * 10  ' WFOSF3測定点2
    WFOSF3_MESDATA3 As String * 10  ' WFOSF3測定点3
    WFOSF3_MESDATA4 As String * 10  ' WFOSF3測定点4
    WFOSF3_MESDATA5 As String * 10  ' WFOSF3測定点5
    WFOSF3_MESDATA6 As String * 10  ' WFOSF3測定点6
    WFOSF3_MESDATA7 As String * 10  ' WFOSF3測定点7
    WFOSF3_MESDATA8 As String * 10  ' WFOSF3測定点8
    WFOSF3_MESDATA9 As String * 10  ' WFOSF3測定点9
    WFOSF3_MESDATA10 As String * 10 ' WFOSF3測定点10
    WFOSF3_MESDATA11 As String * 10 ' WFOSF3測定点11
    WFOSF3_MESDATA12 As String * 10 ' WFOSF3測定点12
    WFOSF3_MESDATA13 As String * 10 ' WFOSF3測定点13
    WFOSF3_MESDATA14 As String * 10 ' WFOSF3測定点14
    WFOSF3_MESDATA15 As String * 10 ' WFOSF3測定点15
    WFOSF4_SMPPOS As Integer        ' WFOSF４ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFOSF4_NETSU As String * 2      ' WFOSF4_熱処理条件
    WFOSF4_ET As String * 3         ' WFOSF4_エッチング条件
    WFOSF4_MES As String * 3        ' WFOSF4_計測方法
    WFOSF4_DKAN As String * 10      ' WFOSF4_ＤＫアニール条件
    WFOSF4_MESDATA1 As String * 10  ' WFOSF4測定点１
    WFOSF4_MESDATA2 As String * 10  ' WFOSF4測定点2
    WFOSF4_MESDATA3 As String * 10  ' WFOSF4測定点3
    WFOSF4_MESDATA4 As String * 10  ' WFOSF4測定点4
    WFOSF4_MESDATA5 As String * 10  ' WFOSF4測定点5
    WFOSF4_MESDATA6 As String * 10  ' WFOSF4測定点6
    WFOSF4_MESDATA7 As String * 10  ' WFOSF4測定点7
    WFOSF4_MESDATA8 As String * 10  ' WFOSF4測定点8
    WFOSF4_MESDATA9 As String * 10  ' WFOSF4測定点9
    WFOSF4_MESDATA10 As String * 10 ' WFOSF4測定点10
    WFOSF4_MESDATA11 As String * 10 ' WFOSF4測定点11
    WFOSF4_MESDATA12 As String * 10 ' WFOSF4測定点12
    WFOSF4_MESDATA13 As String * 10 ' WFOSF4測定点13
    WFOSF4_MESDATA14 As String * 10 ' WFOSF4測定点14
    WFOSF4_MESDATA15 As String * 10 ' WFOSF4測定点15
    WFBMD1_SMPPOS As Integer        ' WFBMD1ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFBMD1_NETSU As String * 2      ' WFBMD1_熱処理条件
    WFBMD1_ET As String * 3         ' WFBMD1_エッチング条件
    WFBMD1_MES As String * 3        ' WFBMD1_計測方法
    WFBMD1_DKAN As String * 10      ' WFBMD1_ＤＫアニール条件
    WFBMD1_MESDATA1 As String * 10  ' WFBMD1測定点１
    WFBMD1_MESDATA2 As String * 10  ' WFBMD1測定点2
    WFBMD1_MESDATA3 As String * 10  ' WFBMD1測定点3
    WFBMD1_MESDATA4 As String * 10  ' WFBMD1測定点4
    WFBMD1_MESDATA5 As String * 10  ' WFBMD1測定点5
    WFBMD1_MESDATA6 As String * 10  ' WFBMD1測定点6
    WFBMD1_MESDATA7 As String * 10  ' WFBMD1測定点7
    WFBMD1_MESDATA8 As String * 10  ' WFBMD1測定点8
    WFBMD1_MESDATA9 As String * 10  ' WFBMD1測定点9
    WFBMD1_MESDATA10 As String * 10 ' WFBMD1測定点10
    WFBMD1_MESDATA11 As String * 10 ' WFBMD1測定点11
    WFBMD1_MESDATA12 As String * 10 ' WFBMD1測定点12
    WFBMD1_MESDATA13 As String * 10 ' WFBMD1測定点13
    WFBMD1_MESDATA14 As String * 10 ' WFBMD1測定点14
    WFBMD1_MESDATA15 As String * 10 ' WFBMD1測定点15
    WFBMD2_SMPPOS As Integer        ' WFBMD２ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFBMD2_NETSU As String * 2      ' WFBMD2_熱処理条件
    WFBMD2_ET As String * 3         ' WFBMD2_エッチング条件
    WFBMD2_MES As String * 3        ' WFBMD2_計測方法
    WFBMD2_DKAN As String * 10      ' WFBMD2_ＤＫアニール条件
    WFBMD2_MESDATA1 As String * 10  ' WFBMD2測定点１
    WFBMD2_MESDATA2 As String * 10  ' WFBMD2測定点2
    WFBMD2_MESDATA3 As String * 10  ' WFBMD2測定点3
    WFBMD2_MESDATA4 As String * 10  ' WFBMD2測定点4
    WFBMD2_MESDATA5 As String * 10  ' WFBMD2測定点5
    WFBMD2_MESDATA6 As String * 10  ' WFBMD2測定点6
    WFBMD2_MESDATA7 As String * 10  ' WFBMD2測定点7
    WFBMD2_MESDATA8 As String * 10  ' WFBMD2測定点8
    WFBMD2_MESDATA9 As String * 10  ' WFBMD2測定点9
    WFBMD2_MESDATA10 As String * 10 ' WFBMD2測定点10
    WFBMD2_MESDATA11 As String * 10 ' WFBMD2測定点11
    WFBMD2_MESDATA12 As String * 10 ' WFBMD2測定点12
    WFBMD2_MESDATA13 As String * 10 ' WFBMD2測定点13
    WFBMD2_MESDATA14 As String * 10 ' WFBMD2測定点14
    WFBMD2_MESDATA15 As String * 10 ' WFBMD2測定点15
    WFBMD3_SMPPOS As Integer        ' WFBMD３ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFBMD3_NETSU As String * 2      ' WFBMD3_熱処理条件
    WFBMD3_ET As String * 3         ' WFBMD3_エッチング条件
    WFBMD3_MES As String * 3        ' WFBMD3_計測方法
    WFBMD3_DKAN As String * 10      ' WFBMD3_ＤＫアニール条件
    WFBMD3_MESDATA1 As String * 10  ' WFBMD3測定点１
    WFBMD3_MESDATA2 As String * 10  ' WFBMD3測定点2
    WFBMD3_MESDATA3 As String * 10  ' WFBMD3測定点3
    WFBMD3_MESDATA4 As String * 10  ' WFBMD3測定点4
    WFBMD3_MESDATA5 As String * 10  ' WFBMD3測定点5
    WFBMD3_MESDATA6 As String * 10  ' WFBMD3測定点6
    WFBMD3_MESDATA7 As String * 10  ' WFBMD3測定点7
    WFBMD3_MESDATA8 As String * 10  ' WFBMD3測定点8
    WFBMD3_MESDATA9 As String * 10  ' WFBMD3測定点9
    WFBMD3_MESDATA10 As String * 10 ' WFBMD3測定点10
    WFBMD3_MESDATA11 As String * 10 ' WFBMD3測定点11
    WFBMD3_MESDATA12 As String * 10 ' WFBMD3測定点12
    WFBMD3_MESDATA13 As String * 10 ' WFBMD3測定点13
    WFBMD3_MESDATA14 As String * 10 ' WFBMD3測定点14
    WFBMD3_MESDATA15 As String * 10 ' WFBMD3測定点15
    WFDSOD_SMPPOS As Integer        ' WFDSODｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFDSOD_NETSU As String * 2      ' WFDSOD_熱処理条件
    WFDSOD_ET As String * 3         ' WFDSOD_エッチング条件
    WFDSOD_MES As String * 3        ' WFDSOD_計測方法
    WFDSOD_DKAN As String * 10      ' WFDSOD_ＤＫアニール条件
    WFDSOD_MESDATA1 As String * 10  ' WFDSOD測定点１
    WFDSOD_MESDATA2 As String * 10  ' WFDSOD測定点2
    WFDSOD_MESDATA3 As String * 10  ' WFDSOD測定点3
    WFDSOD_MESDATA4 As String * 10  ' WFDSOD測定点4
    WFDSOD_MESDATA5 As String * 10  ' WFDSOD測定点5
    WFDSOD_MESDATA6 As String * 10  ' WFDSOD測定点6
    WFDSOD_MESDATA7 As String * 10  ' WFDSOD測定点7
    WFDSOD_MESDATA8 As String * 10  ' WFDSOD測定点8
    WFDSOD_MESDATA9 As String * 10  ' WFDSOD測定点9
    WFDSOD_MESDATA10 As String * 10 ' WFDSOD測定点10
    WFDSOD_MESDATA11 As String * 10 ' WFDSOD測定点11
    WFDSOD_MESDATA12 As String * 10 ' WFDSOD測定点12
    WFDSOD_MESDATA13 As String * 10 ' WFDSOD測定点13
    WFDSOD_MESDATA14 As String * 10 ' WFDSOD測定点14
    WFDSOD_MESDATA15 As String * 10 ' WFDSOD測定点15
    WFSPV_SMPPOS As Integer         ' WFSPVｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFSPV_NETSU As String * 2       ' WFSVP_熱処理条件
    WFSPV_ET As String * 3          ' WFSPV_エッチング条件
    WFSPV_MES As String * 3         ' WFSPV_計測方法
    WFSPV_DKAN As String * 10       ' WFSPV_ＤＫアニール条件
    WFSPV_MESDATA1 As String * 10   ' WFSPV測定点１
    WFSPV_MESDATA2 As String * 10   ' WFSPV測定点2
    WFSPV_MESDATA3 As String * 10   ' WFSPV測定点3
    WFSPV_MESDATA4 As String * 10   ' WFSPV測定点4
    WFSPV_MESDATA5 As String * 10   ' WFSPV測定点5
    WFSPV_MESDATA6 As String * 10   ' WFSPV測定点6
    WFSPV_MESDATA7 As String * 10   ' WFSPV測定点7
    WFSPV_MESDATA8 As String * 10   ' WFSPV測定点8
    WFSPV_MESDATA9 As String * 10   ' WFSPV測定点9
    WFSPV_MESDATA10 As String * 10  ' WFSPV測定点10
    WFSPV_MESDATA11 As String * 10  ' WFSPV測定点11
    WFSPV_MESDATA12 As String * 10  ' WFSPV測定点12
    WFSPV_MESDATA13 As String * 10  ' WFSPV測定点13
    WFSPV_MESDATA14 As String * 10  ' WFSPV測定点14
    WFSPV_MESDATA15 As String * 10  ' WFSPV測定点15
    WFDZ_SMPPOS As Integer          ' WFDZｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFDZ_NETSU As String * 2        ' WFDZ_熱処理条件
    WFDZ_ET As String * 3           ' WFDZ_エッチング条件
    WFDZ_MES As String * 3          ' WFDZ_計測方法
    WFDZ_DKAN As String * 10        ' WFDZ_ＤＫアニール条件
    WFDZ_MESDATA1 As String * 10    ' WFDZ測定点１
    WFDZ_MESDATA2 As String * 10    ' WFDZ測定点2
    WFDZ_MESDATA3 As String * 10    ' WFDZ測定点3
    WFDZ_MESDATA4 As String * 10    ' WFDZ測定点4
    WFDZ_MESDATA5 As String * 10    ' WFDZ測定点5
    WFDZ_MESDATA6 As String * 10    ' WFDZ測定点6
    WFDZ_MESDATA7 As String * 10    ' WFDZ測定点7
    WFDZ_MESDATA8 As String * 10    ' WFDZ測定点8
    WFDZ_MESDATA9 As String * 10    ' WFDZ測定点9
    WFDZ_MESDATA10 As String * 10   ' WFDZ測定点10
    WFDZ_MESDATA11 As String * 10   ' WFDZ測定点11
    WFDZ_MESDATA12 As String * 10   ' WFDZ測定点12
    WFDZ_MESDATA13 As String * 10   ' WFDZ測定点13
    WFDZ_MESDATA14 As String * 10   ' WFDZ測定点14
    WFDZ_MESDATA15 As String * 10   ' WFDZ測定点15
    REGDATE As Date                 ' 登録日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 顧客仕様WFﾃﾞｰﾀ１
Public Type typ_TBCME008
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KMGSRRNO As String * 9          ' 購管理仕様登録依頼番号
    CONFLAG As String * 1           ' 確認フラグ
    REINFLAG As String * 1          ' 再付与フラグ
    KPRFACES As String * 2          ' 購製品表面仕上げ
    KPRBACKS As String * 2          ' 購製品裏仕上げ
    KPRBACK2 As String * 2          ' 購製品裏仕上げ２
    KPRBDSWY As String * 2          ' 購製品ＢＤ処理方法
    KPRFKBWK As String * 1          ' 購製品表面区分方法＿区
    KPRFKBWS As String * 1          ' 購製品表面区分方法＿指
    KPRTYPE As String * 1           ' 購製品タイプ
    KPRTYPKB As String * 1          ' 購製品タイプ検査区分
    KPRTYPKW As String * 1          ' 購製品タイプ検査方法
    KPRDOP As String * 1            ' 購製品ドーパント
    KPRRMIN As Double               ' 購製品比抵抗下限
    KPRRMAX As Double               ' 購製品比抵抗上限
    KPRRUNIT As String * 1          ' 購製品比抵抗単位
    KPRRSPOH As String * 1          ' 購製品比抵抗測定位置＿方
    KPRRSPOT As String * 1          ' 購製品比抵抗測定位置＿点
    KPRRSPOI As String * 1          ' 購製品比抵抗測定位置＿位
    KPRRHWYT As String * 1          ' 購製品比抵抗保証方法＿対
    KPRRHWYS As String * 1          ' 購製品比抵抗保証方法＿処
    KPRRKKBN As String * 1          ' 購製品比抵抗検査区分
    KPRRKWAY As String * 2          ' 購製品比抵抗検査方法
    KPRRKHNM As String * 1          ' 購製品比抵抗検査頻度＿枚
    KPRRKHNN As String * 1          ' 購製品比抵抗検査頻度＿抜
    KPRRKHNH As String * 1          ' 購製品比抵抗検査頻度＿保
    KPRRKHNU As String * 1          ' 購製品比抵抗検査頻度＿ウ
    KPRRSDEV As Double              ' 購製品比抵抗標準偏差
    KPRRAMIN As Double              ' 購製品比抵抗平均下限
    KPRRAMAX As Double              ' 購製品比抵抗平均上限
    KPRRMBNP As Double              ' 購製品比抵抗面内分布
    KPRRMCAL As String * 1          ' 購製品比抵抗面内計算
    KPRRMBP2 As Double              ' 購製品比抵抗面内分布２
    KPRRMCL2 As String * 1          ' 購製品比抵抗面内計算２
    KPRRKBSH As String * 1          ' 購製品比抵抗振区分測定位置＿方
    KPRRKBST As String * 1          ' 購製品比抵抗振区分測定位置＿点
    KPRRKBSI As String * 1          ' 購製品比抵抗振区分測定位置＿位
    KPRRKBHT As String * 1          ' 購製品比抵抗振区分保証方法＿対
    KPRRKBHS As String * 1          ' 購製品比抵抗振区分保証方法＿処
    KPRSTMAX As Double              ' 購製品ストリエ上限
    KPRSTSPH As String * 1          ' 購製品ストリエ測定位置＿方
    KPRSTSPT As String * 1          ' 購製品ストリエ測定位置＿点
    KPRSTSPI As String * 1          ' 購製品ストリエ測定位置＿位
    KPRSTHWT As String * 1          ' 購製品ストリエ保証方法＿対
    KPRSTHWS As String * 1          ' 購製品ストリエ保証方法＿処
    KPRSTKBN As String * 1          ' 購製品ストリエ検査区分
    KPRSTKWY As String * 2          ' 購製品ストリエ検査方法
    KPRSTKHM As String * 1          ' 購製品ストリエ検査頻度＿枚
    KPRSTKHN As String * 1          ' 購製品ストリエ検査頻度＿抜
    KPRSTKHH As String * 1          ' 購製品ストリエ検査頻度＿保
    KPRSTKHU As String * 1          ' 購製品ストリエ検査頻度＿ウ
    KPRRHCAL As String * 2          ' 購製品比抵抗補正計算
    KPRRMINH As Double              ' 購製品比抵抗下限補正
    KPRRMAXH As Double              ' 購製品比抵抗上限補正
    KPRACEN As Double               ' 購製品厚中心
    KPRAMIN As Double               ' 購製品厚下限
    KPRAMAX As Double               ' 購製品厚上限
    KPRAUNIT As String * 1          ' 購製品厚単位
    KPRASPOH As String * 1          ' 購製品厚測定位置＿方
    KPRASPOT As String * 1          ' 購製品厚測定位置＿点
    KPRASPOI As String * 1          ' 購製品厚測定位置＿位
    KPRAHWYT As String * 1          ' 購製品厚保証方法＿対
    KPRAHWYS As String * 1          ' 購製品厚保証方法＿処
    KPRAKKBN As String * 1          ' 購製品厚検査区分
    KPRAKWAY As String * 1          ' 購製品厚検査方法
    KPRAKHNM As String * 1          ' 購製品厚検査頻度＿枚
    KPRAKHNN As String * 1          ' 購製品厚検査頻度＿抜
    KPRAKHNH As String * 1          ' 購製品厚検査頻度＿保
    KPRAKHNU As String * 1          ' 購製品厚検査頻度＿ウ
    KPRASDEV As Double              ' 購製品厚標準偏差
    KPRAAMIN As Double              ' 購製品厚平均下限
    KPRAAMAX As Double              ' 購製品厚平均上限
    KPRAMBNP As Double              ' 購製品厚面内分布
    KPRAMCAL As String * 1          ' 購製品厚面内計算
    KPRALTBP As Double              ' 購製品厚ＬＴ分布
    KPRALTCL As String * 1          ' 購製品厚ＬＴ計算
    KPRALTRA As Double              ' 購製品厚ＬＴ範囲
    KPRAMRAN As Double              ' 購製品厚面内範囲
    KPRAKBSH As String * 1          ' 購製品厚振区分測定位置＿方
    KPRAKBST As String * 1          ' 購製品厚振区分測定位置＿点
    KPRAKBSI As String * 1          ' 購製品厚振区分測定位置＿位
    KPRAKBHT As String * 1          ' 購製品厚振区分保証方法＿対
    KPRAKBHS As String * 1          ' 購製品厚振区分保証方法＿処
    KPRWFORM As String * 1          ' 購製品ウェーハ形状
    KPRD1CEN As Double              ' 購製品直径１中心
    KPRD1MIN As Double              ' 購製品直径１下限
    KPRD1MAX As Double              ' 購製品直径１上限
    KPRD1KBN As String * 1          ' 購製品直径１検査区分
    KPRD2CEN As Double              ' 購製品直径２中心
    KPRD2MIN As Double              ' 購製品直径２下限
    KPRD2MAX As Double              ' 購製品直径２上限
    KPRD2KBN As String * 1          ' 購製品直径２検査区分
    KPRDUNIT As String * 1          ' 購製品直径単位
    KPRDKHNM As String * 1          ' 購製品直径検査頻度＿枚
    KPRDKHNN As String * 1          ' 購製品直径検査頻度＿抜
    KPRDKHNH As String * 1          ' 購製品直径検査頻度＿保
    KPRDKHNU As String * 1          ' 購製品直径検査頻度＿ウ
    KPRLPMNP As Integer             ' 購製品ＬＰ厚最小加工代
    KPRSGMNP As Integer             ' 購製品ＳＧ厚最小加工代
    KPRETMNP As Integer             ' 購製品ＥＴ厚最小加工代
    KPRMPMNP As Integer             ' 購製品ＭＰ厚最小加工代
    KPRLPKS1 As String * 1          ' 購製品ＬＰ研磨材種１
    KPRLPKS2 As String * 1          ' 購製品ＬＰ研磨材種２
    KPRLPKZ1 As String * 1          ' 購製品ＬＰ研磨材粒度種１
    KPRLPKZ2 As String * 1          ' 購製品ＬＰ研磨材粒度種２
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 顧客仕様WFﾃﾞｰﾀ２
Public Type typ_TBCME009
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KMGSRRNO As String * 9          ' 購管理仕様登録依頼番号
    KPRCDIR As String * 1           ' 購製品結晶面方位
    KPRCSCEN As Double              ' 購製品結晶面傾中心
    KPRCSMIN As Double              ' 購製品結晶面傾下限
    KPRCSMAX As Double              ' 購製品結晶面傾上限
    KPRCSDIS As String * 1          ' 購製品結晶面傾方位指定
    KPRCSDIR As String * 2          ' 購製品結晶面傾方位
    KPRCKKBN As String * 1          ' 購製品結晶面検査区分
    KPRCKWAY As String * 2          ' 購製品結晶面検査方法
    KPRCKHNM As String * 1          ' 購製品結晶面検査頻度＿枚
    KPRCKHNN As String * 1          ' 購製品結晶面検査頻度＿抜
    KPRCKHNH As String * 1          ' 購製品結晶面検査頻度＿保
    KPRCKHNU As String * 1          ' 購製品結晶面検査頻度＿ウ
    KPRCTDIR As String * 2          ' 購製品結晶面傾縦方位
    KPRCTCEN As Double              ' 購製品結晶面傾縦中心
    KPRCTMIN As Double              ' 購製品結晶面傾縦下限
    KPRCTMAX As Double              ' 購製品結晶面傾縦上限
    KPRCYDIR As String * 2          ' 購製品結晶面傾横方位
    KPRCYCEN As Double              ' 購製品結晶面傾横中心
    KPRCYMIN As Double              ' 購製品結晶面傾横下限
    KPRCYMAX As Double              ' 購製品結晶面傾横上限
    KPRCSDSC As Double              ' 購製品結晶面傾方位傾中心
    KPRCSDSN As Double              ' 購製品結晶面傾方位傾下限
    KPRCSDSX As Double              ' 購製品結晶面傾方位傾上限
    KPROFPKM As String * 1          ' 購製品ＯＦ位置検査頻度＿枚
    KPROFPKN As String * 1          ' 購製品ＯＦ位置検査頻度＿抜
    KPROFPKH As String * 1          ' 購製品ＯＦ位置検査頻度＿保
    KPROFPKU As String * 1          ' 購製品ＯＦ位置検査頻度＿ウ
    KPROFLKM As String * 1          ' 購製品ＯＦ長検査頻度＿枚
    KPROFLKN As String * 1          ' 購製品ＯＦ長検査頻度＿抜
    KPROFLKH As String * 1          ' 購製品ＯＦ長検査頻度＿保
    KPROFLKU As String * 1          ' 購製品ＯＦ長検査頻度＿ウ
    KPROF1PD As String * 2          ' 購製品ＯＦ１位置方位
    KPROF1PN As Double              ' 購製品ＯＦ１位置下限
    KPROF1PX As Double              ' 購製品ＯＦ１位置上限
    KPROF1PK As String * 1          ' 購製品ＯＦ１位置検査区分
    KPROF1PW As String * 2          ' 購製品ＯＦ１位置検査方法
    KPROF1LC As Double              ' 購製品ＯＦ１長中心
    KPROF1LN As Double              ' 購製品ＯＦ１長下限
    KPROF1LX As Double              ' 購製品ＯＦ１長上限
    KPROF1LK As String * 1          ' 購製品ＯＦ１長検査区分
    KPROF1RF As String * 1          ' 購製品ＯＦ１両端Ｒ形状
    KPROFRRC As Double              ' 購製品ＯＦ両端Ｒ右中心
    KPROFRRN As Double              ' 購製品ＯＦ両端Ｒ右下限
    KPROFRRX As Double              ' 購製品ＯＦ両端Ｒ右上限
    KPROFRLC As Double              ' 購製品ＯＦ両端Ｒ左中心
    KPROFRLN As Double              ' 購製品ＯＦ両端Ｒ左下限
    KPROFRLX As Double              ' 購製品ＯＦ両端Ｒ左上限
    KPROFRKB As String * 1          ' 購製品ＯＦ両端Ｒ検査区分
    KPROF1DC As Double              ' 購製品ＯＦ１直径中心
    KPROF1DN As Double              ' 購製品ＯＦ１直径下限
    KPROF1DX As Double              ' 購製品ＯＦ１直径上限
    KPROF1DK As String * 1          ' 購製品ＯＦ１直径検査区分
    KPRDFORM As String * 1          ' 購製品溝形状
    KPRDFKBN As String * 1          ' 購製品溝形状検査区分
    KPRDFKHM As String * 1          ' 購製品溝形状検査頻度＿枚
    KPRDFKHN As String * 1          ' 購製品溝形状検査頻度＿抜
    KPRDFKHH As String * 1          ' 購製品溝形状検査頻度＿保
    KPRDFKHU As String * 1          ' 購製品溝形状検査頻度＿ウ
    KPRDPDRC As String * 1          ' 購製品溝位置方向
    KPRDPACN As Integer             ' 購製品溝位置角度中心
    KPRDPAMN As Integer             ' 購製品溝位置角度下限
    KPRDPAMX As Integer             ' 購製品溝位置角度上限
    KPRDPDIR As String * 2          ' 購製品溝位置方位
    KPRDPMIN As Double              ' 購製品溝位置下限
    KPRDPMAX As Double              ' 購製品溝位置上限
    KPRDPKBN As String * 1          ' 購製品溝位置検査区分
    KPRDPKWY As String * 2          ' 購製品溝位置検査方法
    KPRDPKHM As String * 1          ' 購製品溝位置検査頻度＿枚
    KPRDPKHB As String * 1          ' 購製品溝位置検査頻度＿抜
    KPRDPKHH As String * 1          ' 購製品溝位置検査頻度＿保
    KPRDPKHU As String * 1          ' 購製品溝位置検査頻度＿ウ
    KPRDACEN As Double              ' 購製品溝角度中心
    KPRDAMIN As Double              ' 購製品溝角度下限
    KPRDAMAX As Double              ' 購製品溝角度上限
    KPRDAKBN As String * 1          ' 購製品溝角度検査区分
    KPRDWCEN As Double              ' 購製品溝巾中心
    KPRDWMIN As Double              ' 購製品溝巾下限
    KPRDWMAX As Double              ' 購製品溝巾上限
    KPRDWKBN As String * 1          ' 購製品溝巾検査区分
    KPRDDCEN As Double              ' 購製品溝深中心
    KPRDDMIN As Double              ' 購製品溝深下限
    KPRDDMAX As Double              ' 購製品溝深上限
    KPRDDKBN As String * 1          ' 購製品溝深検査区分
    KPRDBRCN As Double              ' 購製品溝底Ｒ中心
    KPRDBRMN As Double              ' 購製品溝底Ｒ下限
    KPRDBRMX As Double              ' 購製品溝底Ｒ上限
    KPRDBRKB As String * 1          ' 購製品溝底Ｒ検査区分
    KPRDRRCN As Double              ' 購製品溝右Ｒ中心
    KPRDRRMN As Double              ' 購製品溝右Ｒ下限
    KPRDRRMX As Double              ' 購製品溝右Ｒ上限
    KPRDLRCN As Double              ' 購製品溝左Ｒ中心
    KPRDLRMN As Double              ' 購製品溝左Ｒ下限
    KPRDLRMX As Double              ' 購製品溝左Ｒ上限
    KPRDRRKB As String * 1          ' 購製品溝両端Ｒ検査区分
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 顧客仕様WFﾃﾞｰﾀ３
Public Type typ_TBCME010
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KMGSRRNO As String * 9          ' 購管理仕様登録依頼番号
    KPRMFORM As String * 1          ' 購製品面取形状
    KPRMM As String * 1             ' 購製品面取面粗
    KPRMFKBN As String * 1          ' 購製品面取形状検査区分
    KPRMFKHM As String * 1          ' 購製品面取形状検査頻度＿枚
    KPRMFKHN As String * 1          ' 購製品面取形状検査頻度＿抜
    KPRMFKHH As String * 1          ' 購製品面取形状検査頻度＿保
    KPRMFKHU As String * 1          ' 購製品面取形状検査頻度＿ウ
    KPRMMKBN As String * 1          ' 購製品面取面粗検査区分
    KPRMACEN As Double              ' 購製品面取角度中心
    KPRMAMIN As Double              ' 購製品面取角度下限
    KPRMAMAX As Double              ' 購製品面取角度上限
    KPRMAKBN As String * 1          ' 購製品面取角度検査区分
    KPRMWFCN As Integer             ' 購製品面取巾表中心
    KPRMWFMN As Integer             ' 購製品面取巾表下限
    KPRMWFMX As Integer             ' 購製品面取巾表上限
    KPRMWBCN As Integer             ' 購製品面取巾裏中心
    KPRMWBMN As Integer             ' 購製品面取巾裏下限
    KPRMWBMX As Integer             ' 購製品面取巾裏上限
    KPRMHKBN As String * 1          ' 購製品面取高検査区分
    KPRMHCEN As Integer             ' 購製品面取高中心
    KPRMHMIN As Integer             ' 購製品面取高下限
    KPRMHMAX As Integer             ' 購製品面取高上限
    KPRMWKBN As String * 1          ' 購製品面取巾検査区分
    KPRMPWCN As Integer             ' 購製品面取先端巾中心
    KPRMPWMN As Integer             ' 購製品面取先端巾下限
    KPRMPWMX As Integer             ' 購製品面取先端巾上限
    KPRMPWKB As String * 1          ' 購製品面取先端巾検査区分
    KPRMPRCN As Double              ' 購製品面取先端Ｒ中心
    KPRMPRMN As Double              ' 購製品面取先端Ｒ下限
    KPRMPRMX As Double              ' 購製品面取先端Ｒ上限
    KPRMPRKB As String * 1          ' 購製品面取先端Ｒ検査区分
    KPRDMFRM As String * 1          ' 購製品溝面取形状
    KPRDMM As String * 1            ' 購製品溝面取面粗
    KPRDMPRC As Double              ' 購製品溝面取先端Ｒ中心
    KPRDMACN As Double              ' 購製品溝面取角度中心
    KPRIDSTA As String * 2          ' 購製品ＩＤ規格
    KPRIDWAY As String * 1          ' 購製品ＩＤ方法
    KPRIDPRI As String * 1          ' 購製品ＩＤ印字種類
    KPRIDKND As String * 1          ' 購製品ＩＤ種類
    KPRIDDIR As String * 1          ' 購製品ＩＤ方向
    KPRIDFAC As String * 1          ' 購製品ＩＤ面
    KPRCSIZE As String * 1          ' 購製品文字サイズ
    KPRIDPBS As String * 1          ' 購製品ＩＤ位置測定基準
    KPRIDFIG As Integer             ' 購製品ＩＤ桁数
    KPRIDCON As String              ' 購製品ＩＤ内容
    KPRIDZAR As Double              ' 購製品ＩＤ除外領域
    KPRIDPAP As String * 1          ' 購製品ＩＤ印字連番指定
    KPRIDDCN As Integer             ' 購製品ＩＤドット深中心
    KPRIDDMX As Integer             ' 購製品ＩＤドット深上限
    KPRIDDMN As Integer             ' 購製品ＩＤドット深下限
    KPRIDSCN As Integer             ' 購製品ＩＤドットＳ中心
    KPRIDSMX As Integer             ' 購製品ＩＤドットＳ上限
    KPRIDSMN As Integer             ' 購製品ＩＤドットＳ下限
    KPRBDPRS As Double              ' 購製品ＢＤ圧力
    KPRBDTIM As Integer             ' 購製品ＢＤ回数
    KPRETWAY As String * 2          ' 購製品ＥＴ方法
    KPRMPFIN As String * 1          ' 購製品ＭＰ仕上げ
    KPRLWASW As String * 1          ' 購製品最終洗浄方法
    KPRCDOP As String * 1           ' 購製品結晶ドープ
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 顧客仕様WFﾃﾞｰﾀ４
Public Type typ_TBCME011
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KMGSRRNO As String * 9          ' 購管理仕様登録依頼番号
    KPRM1S As String * 1            ' 購製品膜１種
    KPRM1H As String * 1            ' 購製品膜１付面
    KPRM2S As String * 1            ' 購製品膜２種
    KPRM2H As String * 1            ' 購製品膜２付面
    KPRNJSUM As String * 1          ' 購製品ノジュール処理有無
    KPRNJSMX As Double              ' 購製品ノジュール処理巾上限
    KPRNJSMN As Double              ' 購製品ノジュール処理巾下限
    KPROXCEN As Long                ' 購製品酸化膜厚中心
    KPROXMIN As Long                ' 購製品酸化膜厚下限
    KPROXMAX As Long                ' 購製品酸化膜厚上限
    KPROXUNT As String * 1          ' 購製品酸化膜厚単位
    KPROXSPH As String * 1          ' 購製品酸化膜厚測定位置＿方
    KPROXSPT As String * 1          ' 購製品酸化膜厚測定位置＿点
    KPROXSPI As String * 1          ' 購製品酸化膜厚測定位置＿位
    KPROXHWT As String * 1          ' 購製品酸化膜厚保証方法＿対
    KPROXHWS As String * 1          ' 購製品酸化膜厚保証方法＿処
    KPROXHWY As String * 2          ' 購製品酸化膜厚検査方法
    KPROXNPO As String * 1          ' 購製品酸化膜厚抜取位置
    KPROXKHM As String * 1          ' 購製品酸化膜厚検査頻度＿枚
    KPROXKHN As String * 1          ' 購製品酸化膜厚検査頻度＿抜
    KPROXKHH As String * 1          ' 購製品酸化膜厚検査頻度＿保
    KPROXKHU As String * 1          ' 購製品酸化膜厚検査頻度＿ウ
    KPROXZAR As Integer             ' 購製品酸化膜除外領域
    KPROXMBP As Double              ' 購製品酸化膜厚面内分布
    KPROXMCL As String * 1          ' 購製品酸化膜厚面内計算
    KPROXMRA As Integer             ' 購製品酸化膜厚面内範囲
    KPROXLTB As Double              ' 購製品酸化膜厚ＬＴ分布
    KPROXLTC As String * 1          ' 購製品酸化膜厚ＬＴ計算
    KPROXLTR As Integer             ' 購製品酸化膜厚ＬＴ範囲
    KPRPSCEN As Double              ' 購製品ポリシリ厚中心
    KPRPSMIN As Double              ' 購製品ポリシリ厚下限
    KPRPSMAX As Double              ' 購製品ポリシリ厚上限
    KPRPSUNT As String * 1          ' 購製品ポリシリ膜厚単位
    KPRPSSPH As String * 1          ' 購製品ポリシリ厚測定位置＿方
    KPRPSSPT As String * 1          ' 購製品ポリシリ厚測定位置＿点
    KPRPSSPI As String * 1          ' 購製品ポリシリ厚測定位置＿位
    KPRPSHWT As String * 1          ' 購製品ポリシリ厚保証方法＿対
    KPRPSHWS As String * 1          ' 購製品ポリシリ厚保証方法＿処
    KPRPSKWY As String * 2          ' 購製品ポリシリ厚検査方法
    KPRPSNPS As String * 1          ' 購製品ポリシリ厚抜取位置
    KPRPSKHM As String * 1          ' 購製品ポリシリ厚検査頻度＿枚
    KPRPSKHN As String * 1          ' 購製品ポリシリ厚検査頻度＿抜
    KPRPSKHH As String * 1          ' 購製品ポリシリ厚検査頻度＿保
    KPRPSKHU As String * 1          ' 購製品ポリシリ厚検査頻度＿ウ
    KPRPSMBP As Double              ' 購製品ポリシリ厚面内分布
    KPRPSMCL As String * 1          ' 購製品ポリシリ厚面内計算
    KPRPSMRA As Double              ' 購製品ポリシリ厚面内範囲
    KPRNOXCN As Long                ' 購製品窒化膜厚中心
    KPRNOXMN As Long                ' 購製品窒化膜厚下限
    KPRNOXMX As Long                ' 購製品窒化膜厚上限
    KPRNOXUN As String * 1          ' 購製品窒化膜厚単位
    KPRNOXSH As String * 1          ' 購製品窒化膜厚測定位置＿方
    KPRNOXST As String * 1          ' 購製品窒化膜厚測定位置＿点
    KPRNOXSI As String * 1          ' 購製品窒化膜厚測定位置＿位
    KPRNOXHT As String * 1          ' 購製品窒化膜厚保証方法＿対
    KPRNOXHS As String * 1          ' 購製品窒化膜厚保証方法＿処
    KPRNOXHW As String * 2          ' 購製品窒化膜厚検査方法
    KPRNOXNP As String * 1          ' 購製品窒化膜厚抜取位置
    KPRNOXKM As String * 1          ' 購製品窒化膜厚検査頻度＿枚
    KPRNOXKN As String * 1          ' 購製品窒化膜厚検査頻度＿抜
    KPRNOXKH As String * 1          ' 購製品窒化膜厚検査頻度＿保
    KPRNOXKU As String * 1          ' 購製品窒化膜厚検査頻度＿ウ
    KPRNOXMB As Double              ' 購製品窒化膜厚面内分布
    KPRNOXMC As String * 1          ' 購製品窒化膜厚面内計算
    KPRNOXMR As Integer             ' 購製品窒化膜厚面内範囲
    KPRMKMIN As Double              ' 購製品無欠陥層下限
    KPRMKMAX As Double              ' 購製品無欠陥層上限
    KPRMKSPH As String * 1          ' 購製品無欠陥層測定位置＿方
    KPRMKSPT As String * 1          ' 購製品無欠陥層測定位置＿点
    KPRMKSPR As String * 1          ' 購製品無欠陥層測定位置＿領
    KPRMKHWT As String * 1          ' 購製品無欠陥層保証方法＿対
    KPRMKHWS As String * 1          ' 購製品無欠陥層保証方法＿処
    KPRMKSZY As String * 1          ' 購製品無欠陥層測定条件
    KPRMKKHM As String * 1          ' 購製品無欠陥層検査頻度＿枚
    KPRMKKHN As String * 1          ' 購製品無欠陥層検査頻度＿抜
    KPRMKKHH As String * 1          ' 購製品無欠陥層検査頻度＿保
    KPRMKKHU As String * 1          ' 購製品無欠陥層検査頻度＿ウ
    KPRMKNSW As String * 2          ' 購製品無欠陥層熱処理法
    KPRMKFGS As String * 1          ' 購製品無欠陥層雰囲気ガス
    KPRMKCET As Integer             ' 購製品無欠陥層選択ＥＴ代
    KPRDZSWY As String * 1          ' 購製品ＤＺ処理方法
    KPRD1STO As Integer             ' 購製品ＤＺ１ＳＴ温度
    KPRD1STT As Integer             ' 購製品ＤＺ１ＳＴ時間
    KPRD1STG As String * 1          ' 購製品ＤＺ１ＳＴガス条件
    KPRD2NDO As Integer             ' 購製品ＤＺ２ＮＤ温度
    KPRD2NDC As Integer             ' 購製品ＤＺ２ＮＤ温度定常
    KPRD2NDT As Integer             ' 購製品ＤＺ２ＮＤ時間
    KPRD3RDO As Integer             ' 購製品ＤＺ３ＲＤ温度
    KPRD3RDT As Integer             ' 購製品ＤＺ３ＲＤ時間
    KPRDZMPS As String * 1          ' 購製品ＤＺＭＰ処理区分
    KPRH2ANO As Integer             ' 購製品Ｈ２ＡＮ温度
    KPRH2ANT As Integer             ' 購製品Ｈ２ＡＮ時間
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 顧客仕様WFﾃﾞｰﾀ５
Public Type typ_TBCME012
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KMGSRRNO As String * 9          ' 購管理仕様登録依頼番号
    KPRTMMAX As Long                ' 購製品転位密度上限
    KPRTMSPH As String * 1          ' 購製品転位密度測定位置＿方
    KPRTMSPT As String * 1          ' 購製品転位密度測定位置＿点
    KPRTMSPR As String * 1          ' 購製品転位密度測定位置＿領
    KPRTMKBN As String * 1          ' 購製品転位密度検査区分
    KPRTMKHM As String * 1          ' 購製品転位密度検査頻度＿枚
    KPRTMKHN As String * 1          ' 購製品転位密度検査頻度＿抜
    KPRTMKHH As String * 1          ' 購製品転位密度検査頻度＿保
    KPRTMKHU As String * 1          ' 購製品転位密度検査頻度＿ウ
    KPRLTMIN As Integer             ' 購製品Ｌタイム下限
    KPRLTMAX As Integer             ' 購製品Ｌタイム上限
    KPRLTUNT As String * 1          ' 購製品Ｌタイム単位
    KPRLTSPH As String * 1          ' 購製品Ｌタイム測定位置＿方
    KPRLTSPT As String * 1          ' 購製品Ｌタイム測定位置＿点
    KPRLTSPI As String * 1          ' 購製品Ｌタイム測定位置＿位
    KPRLTHWT As String * 1          ' 購製品Ｌタイム保証方法＿対
    KPRLTHWS As String * 1          ' 購製品Ｌタイム保証方法＿処
    KPRLTNSW As String * 2          ' 購製品Ｌタイム熱処理法
    KPRLTKBN As String * 1          ' 購製品Ｌタイム検査区分
    KPRLTKWY As String * 2          ' 購製品Ｌタイム検査方法
    KPRLTKHM As String * 1          ' 購製品Ｌタイム検査頻度＿枚
    KPRLTKHN As String * 1          ' 購製品Ｌタイム検査頻度＿抜
    KPRLTKHH As String * 1          ' 購製品Ｌタイム検査頻度＿保
    KPRLTKHU As String * 1          ' 購製品Ｌタイム検査頻度＿ウ
    KPRLTMBP As Double              ' 購製品Ｌタイム面内分布
    KPRLTMCL As String * 1          ' 購製品Ｌタイム面内計算
    KPRCNMIN As Double              ' 購製品炭素濃度下限
    KPRCNMAX As Double              ' 購製品炭素濃度上限
    KPRCNUNT As String * 1          ' 購製品炭素濃度単位
    KPRCNIND As String * 2          ' 購製品炭素濃度指数
    KPRCNSPH As String * 1          ' 購製品炭素濃度測定位置＿方
    KPRCNSPT As String * 1          ' 購製品炭素濃度測定位置＿点
    KPRCNSPI As String * 1          ' 購製品炭素濃度測定位置＿位
    KPRCNHWT As String * 1          ' 購製品炭素濃度保証方法＿対
    KPRCNHWS As String * 1          ' 購製品炭素濃度保証方法＿処
    KPRCNKBN As String * 1          ' 購製品炭素濃度検査区分
    KPRCNKWY As String * 2          ' 購製品炭素濃度検査方法
    KPRONMIN As Double              ' 購製品酸素濃度下限
    KPRONMAX As Double              ' 購製品酸素濃度上限
    KPRONUNT As String * 1          ' 購製品酸素濃度単位
    KPRONIND As String * 2          ' 購製品酸素濃度指数
    KPRONSPH As String * 1          ' 購製品酸素濃度測定位置＿方
    KPRONSPT As String * 1          ' 購製品酸素濃度測定位置＿点
    KPRONSPI As String * 1          ' 購製品酸素濃度測定位置＿位
    KPRONHWT As String * 1          ' 購製品酸素濃度保証方法＿対
    KPRONHWS As String * 1          ' 購製品酸素濃度保証方法＿処
    KPRONKBN As String * 1          ' 購製品酸素濃度検査区分
    KPRONKWY As String * 2          ' 購製品酸素濃度検査方法
    KPRONKHM As String * 1          ' 購製品酸素濃度検査頻度＿枚
    KPRONKHN As String * 1          ' 購製品酸素濃度検査頻度＿抜
    KPRONKHH As String * 1          ' 購製品酸素濃度検査頻度＿保
    KPRONKHU As String * 1          ' 購製品酸素濃度検査頻度＿ウ
    KPRONMBP As Double              ' 購製品酸素濃度面内分布
    KPRONMCL As String * 1          ' 購製品酸素濃度面内計算
    KPRONLTB As Double              ' 購製品酸素濃度ＬＴ分布
    KPRONLTC As String * 1          ' 購製品酸素濃度ＬＴ計算
    KPRONSDV As Double              ' 購製品酸素濃度標準偏差
    KPRONAMN As Double              ' 購製品酸素濃度平均下限
    KPRONAMX As Double              ' 購製品酸素濃度平均上限
    KPRONAST As String * 1          ' 購製品酸素濃度ＡＳＴＭ新旧
    KPRONHCL As String * 2          ' 購製品酸素濃度補正計算
    KPRONMNH As Double              ' 購製品酸素濃度下限補正
    KPRONMXH As Double              ' 購製品酸素濃度上限補正
    KPROKBSH As String * 1          ' 購製品酸素振区分測定位置＿方
    KPROKBST As String * 1          ' 購製品酸素振区分測定位置＿点
    KPROKBSI As String * 1          ' 購製品酸素振区分測定位置＿位
    KPROKBHT As String * 1          ' 購製品酸素振区分保証方法＿対
    KPROKBHS As String * 1          ' 購製品酸素振区分保証方法＿処
    KPROS1MN As Double              ' 購製品酸素析出１下限
    KPROS1MX As Double              ' 購製品酸素析出１上限
    KPROS1NS As String * 2          ' 購製品酸素析出１熱処理法
    KPROS1SH As String * 1          ' 購製品酸素析出１測定位置＿方
    KPROS1ST As String * 1          ' 購製品酸素析出１測定位置＿点
    KPROS1SI As String * 1          ' 購製品酸素析出１測定位置＿位
    KPROS1HT As String * 1          ' 購製品酸素析出１保証方法＿対
    KPROS1HS As String * 1          ' 購製品酸素析出１保証方法＿処
    KPROS1HM As String * 1          ' 購製品酸素析出１検査頻度＿枚
    KPROS1KN As String * 1          ' 購製品酸素析出１検査頻度＿抜
    KPROS1KH As String * 1          ' 購製品酸素析出１検査頻度＿保
    KPROS1KU As String * 1          ' 購製品酸素析出１検査頻度＿ウ
    KPROS2MN As Double              ' 購製品酸素析出２下限
    KPROS2MX As Double              ' 購製品酸素析出２上限
    KPROS2NS As String * 2          ' 購製品酸素析出２熱処理法
    KPROS2SH As String * 1          ' 購製品酸素析出２測定位置＿方
    KPROS2ST As String * 1          ' 購製品酸素析出２測定位置＿点
    KPROS2SI As String * 1          ' 購製品酸素析出２測定位置＿位
    KPROS2HT As String * 1          ' 購製品酸素析出２保証方法＿対
    KPROS2HS As String * 1          ' 購製品酸素析出２保証方法＿処
    KPROS2KM As String * 1          ' 購製品酸素析出２検査頻度＿枚
    KPROS2KN As String * 1          ' 購製品酸素析出２検査頻度＿抜
    KPROS2KH As String * 1          ' 購製品酸素析出２検査頻度＿保
    KPROS2KU As String * 1          ' 購製品酸素析出２検査頻度＿ウ
    KPROS3MN As Double              ' 購製品酸素析出３下限
    KPROS3MX As Double              ' 購製品酸素析出３上限
    KPROS3NS As String * 2          ' 購製品酸素析出３熱処理法
    KPROS3SH As String * 1          ' 購製品酸素析出３測定位置＿方
    KPROS3ST As String * 1          ' 購製品酸素析出３測定位置＿点
    KPROS3SI As String * 1          ' 購製品酸素析出３測定位置＿位
    KPROS3HT As String * 1          ' 購製品酸素析出３保証方法＿対
    KPROS3HS As String * 1          ' 購製品酸素析出３保証方法＿処
    KPROS3KM As String * 1          ' 購製品酸素析出３検査頻度＿枚
    KPROS3KN As String * 1          ' 購製品酸素析出３検査頻度＿抜
    KPROS3KH As String * 1          ' 購製品酸素析出３検査頻度＿保
    KPROS3KU As String * 1          ' 購製品酸素析出３検査頻度＿ウ
    KPRANTNP As Integer             ' 購製品ＡＮ温度
    KPRANTIM As Integer             ' 購製品ＡＮ時間
    KPRANTMN As Integer             ' 購製品ＡＮ時間下限
    KPRANTMX As Integer             ' 購製品ＡＮ時間上限
    KPRZOMIN As Double              ' 購製品残存酸素下限
    KPRZOMAX As Double              ' 購製品残存酸素上限
    KPRZOSPH As String * 1          ' 購製品残存酸素測定位置＿方
    KPRZOSPT As String * 1          ' 購製品残存酸素測定位置＿点
    KPRZOSPI As String * 1          ' 購製品残存酸素測定位置＿位
    KPRZOHWT As String * 1          ' 購製品残存酸素保証方法＿対
    KPRZOHWS As String * 1          ' 購製品残存酸素保証方法＿処
    KPRZONSW As String * 2          ' 購製品残存酸素熱処理法
    KPRZOKWY As String * 2          ' 購製品残存酸素検査方法
    KPRZOKHM As String * 1          ' 購製品残存酸素検査頻度＿枚
    KPRZOKHN As String * 1          ' 購製品残存酸素検査頻度＿抜
    KPRZOKHH As String * 1          ' 購製品残存酸素検査頻度＿保
    KPRZOKHU As String * 1          ' 購製品残存酸素検査頻度＿ウ
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 顧客仕様WFﾃﾞｰﾀ６
Public Type typ_TBCME013
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KMGSRRNO As String * 9          ' 購管理仕様登録依頼番号
    KPRBDOMN As Integer             ' 購製品ＢＤＯＳＦ下限
    KPRBDOMX As Integer             ' 購製品ＢＤＯＳＦ上限
    KPRBDOSH As String * 1          ' 購製品ＢＤＯＳＦ測定位置＿方
    KPRBDOST As String * 1          ' 購製品ＢＤＯＳＦ測定位置＿点
    KPRBDOSR As String * 1          ' 購製品ＢＤＯＳＦ測定位置＿領
    KPRBDOHT As String * 1          ' 購製品ＢＤＯＳＦ保証方法＿対
    KPRBDOHS As String * 1          ' 購製品ＢＤＯＳＦ保証方法＿処
    KPRBDOSZ As String * 1          ' 購製品ＢＤＯＳＦ測定条件
    KPRBDONS As String * 2          ' 購製品ＢＤＯＳＦ熱処理法
    KPRBDOKM As String * 1          ' 購製品ＢＤＯＳＦ検査頻度＿枚
    KPRBDOKN As String * 1          ' 購製品ＢＤＯＳＦ検査頻度＿抜
    KPRBDOKH As String * 1          ' 購製品ＢＤＯＳＦ検査頻度＿保
    KPRBDOKU As String * 1          ' 購製品ＢＤＯＳＦ検査頻度＿ウ
    KPRBDOET As Integer             ' 購製品ＢＤＯＳＦ選択ＥＴ代
    KPRBDSMN As Integer             ' 購製品ＢＤＳＴ跡下限
    KPRBDSMX As Integer             ' 購製品ＢＤＳＴ跡上限
    KPRBDSSH As String * 1          ' 購製品ＢＤＳＴ跡測定位置＿方
    KPRBDSST As String * 1          ' 購製品ＢＤＳＴ跡測定位置＿点
    KPRBDSSR As String * 1          ' 購製品ＢＤＳＴ跡測定位置＿領
    KPRBDSHT As String * 1          ' 購製品ＢＤＳＴ跡保証方法＿対
    KPRBDSHS As String * 1          ' 購製品ＢＤＳＴ跡保証方法＿処
    KPRBDSSZ As String * 1          ' 購製品ＢＤＳＴ跡測定条件
    KPRBDSKM As String * 1          ' 購製品ＢＤＳＴ跡検査頻度＿枚
    KPRBDSKN As String * 1          ' 購製品ＢＤＳＴ跡検査頻度＿抜
    KPRBDSKH As String * 1          ' 購製品ＢＤＳＴ跡検査頻度＿保
    KPRBDSKU As String * 1          ' 購製品ＢＤＳＴ跡検査頻度＿ウ
    KPRBDSET As Integer             ' 購製品ＢＤＳＴ跡選択ＥＴ代
    KPRRNFMX As Double              ' 購製品ラフネス表上限
    KPRRNFKB As String * 1          ' 購製品ラフネス表検査区分
    KPRRNFKW As String * 2          ' 購製品ラフネス表検査方法
    KPRRNFZA As Integer             ' 購製品ラフネス表除外領域
    KPRRNBMX As Double              ' 購製品ラフネス裏上限
    KPRRNBKB As String * 1          ' 購製品ラフネス裏検査区分
    KPRRNBKW As String * 2          ' 購製品ラフネス裏検査方法
    KPRRNBZA As Integer             ' 購製品ラフネス裏除外領域
    KPRDENKU As String * 1          ' 購製品Ｄｅｎ検査有無
    KPRDENMX As Integer             ' 購製品Ｄｅｎ上限
    KPRDENMN As Integer             ' 購製品Ｄｅｎ下限
    KPRDENHT As String * 1          ' 購製品Ｄｅｎ保証方法＿対
    KPRDENHS As String * 1          ' 購製品Ｄｅｎ保証方法＿処
    KPRLDLKU As String * 1          ' 購製品Ｌ／ＤＬ検査有無
    KPRLDLMX As Integer             ' 購製品Ｌ／ＤＬ上限
    KPRLDLMN As Integer             ' 購製品Ｌ／ＤＬ下限
    KPRLDLHT As String * 1          ' 購製品Ｌ／ＤＬ保証方法＿対
    KPRLDLHS As String * 1          ' 購製品Ｌ／ＤＬ保証方法＿処
    KPRDVDKU As String * 1          ' 購製品ＤＶＤ２検査有無
    KPRDVDMX As Integer             ' 購製品ＤＶＤ２上限
    KPRDVDMN As Integer             ' 購製品ＤＶＤ２下限
    KPRDVDHT As String * 1          ' 購製品ＤＶＤ２保証方法＿対
    KPRDVDHS As String * 1          ' 購製品ＤＶＤ２保証方法＿処
    KPRGDSPH As String * 1          ' 購製品ＧＤ測定位置＿方
    KPRGDSPT As String * 1          ' 購製品ＧＤ測定位置＿点
    KPRGDSPR As String * 1          ' 購製品ＧＤ測定位置＿領
    KPRGDSZY As String * 1          ' 購製品ＧＤ測定条件
    KPRGDZAR As Integer             ' 購製品ＧＤ除外領域
    KPRGDKHM As String * 1          ' 購製品ＧＤ検査頻度＿枚
    KPRGDKHN As String * 1          ' 購製品ＧＤ検査頻度＿抜
    KPRGDKHH As String * 1          ' 購製品ＧＤ検査頻度＿保
    KPRGDKHU As String * 1          ' 購製品ＧＤ検査頻度＿ウ
    KPRDSOKE As String * 1          ' 購製品ＤＳＯＤ検査
    KPRDSOMX As Long                ' 購製品ＤＳＯＤ上限
    KPRDSOMN As Long                ' 購製品ＤＳＯＤ下限
    KPRDSOAX As Integer             ' 購製品ＤＳＯＤ領域上限
    KPRDSOAN As Integer             ' 購製品ＤＳＯＤ領域下限
    KPRDSOHT As String * 1          ' 購製品ＤＳＯＤ保証方法＿対
    KPRDSOHS As String * 1          ' 購製品ＤＳＯＤ保証方法＿処
    KPRDSOKM As String * 1          ' 購製品ＤＳＯＤ検査頻度＿枚
    KPRDSOKN As String * 1          ' 購製品ＤＳＯＤ検査頻度＿抜
    KPRDSOKH As String * 1          ' 購製品ＤＳＯＤ検査頻度＿保
    KPRDSOKU As String * 1          ' 購製品ＤＳＯＤ検査頻度＿ウ
    KPRNTPUM As String * 1          ' 購製品平坦ナノトポ有無
    KPRNTPK1 As Double              ' 購製品平坦ナノトポ規格１
    KPRNTPP1 As Double              ' 購製品平坦ナノトポＰＵＡ１
    KPRNTPS1 As Double              ' 購製品平坦ナノトポサイト１
    KPRNTPK2 As Double              ' 購製品平坦ナノトポ規格２
    KPRNTPP2 As Double              ' 購製品平坦ナノトポＰＵＡ２
    KPRNTPS2 As Double              ' 購製品平坦ナノトポサイト２
    KPRNTPK3 As Double              ' 購製品平坦ナノトポ規格３
    KPRNTPP3 As Double              ' 購製品平坦ナノトポＰＵＡ３
    KPRNTPS3 As Double              ' 購製品平坦ナノトポサイト３
    KPRNTPZA As Integer             ' 購製品平坦ナノトポ除外領域
    KPRNTPHT As String * 1          ' 購製品平坦ナノトポ保証方法＿対
    KPRNTPHS As String * 1          ' 購製品平坦ナノトポ保証方法＿処
    KPRNTPKM As String * 1          ' 購製品平坦ナノトポ検査頻度＿枚
    KPRNTPKN As String * 1          ' 購製品平坦ナノトポ検査頻度＿抜
    KPRNTPKH As String * 1          ' 購製品平坦ナノトポ検査頻度＿保
    KPRNTPKU As String * 1          ' 購製品平坦ナノトポ検査頻度＿ウ
    KPRCRSSK As String * 1          ' 購製品平坦クロスＳＳ検査
    KPRMDCEN As Double              ' 購製品平坦面ダレ高低差中心
    KPRMDMAX As Double              ' 購製品平坦面ダレ高低差上限
    KPRMDMIN As Double              ' 購製品平坦面ダレ高低差下限
    KPRMDSPH As String * 3          ' 購製品平坦面ダレ測定位置＿方
    KPRMDSPT As String * 3          ' 購製品平坦面ダレ測定位置＿点
    KPRMDSPI As String * 3          ' 購製品平坦面ダレ測定位置＿位
    KPRMDHWT As String * 2          ' 購製品平坦面ダレ保証方法＿対
    KPRMDHWS As String * 2          ' 購製品平坦面ダレ保証方法＿処
    KPRMDKHM As String * 4          ' 購製品平坦面ダレ検査頻度＿枚
    KPRMDKHN As String * 4          ' 購製品平坦面ダレ検査頻度＿抜
    KPRMDKHH As String * 4          ' 購製品平坦面ダレ検査頻度＿保
    KPRMDKHU As String * 4          ' 購製品平坦面ダレ検査頻度＿ウ
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 顧客仕様WFﾃﾞｰﾀ７
Public Type typ_TBCME014
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KMGSRRNO As String * 9          ' 購管理仕様登録依頼番号
    KPRSMIN As Double               ' 購製品反り下限
    KPRSMAX As Double               ' 購製品反り上限
    KPRSHWYT As String * 1          ' 購製品反り保証方法＿対
    KPRSHWYS As String * 1          ' 購製品反り保証方法＿処
    KPRSKKBN As String * 1          ' 購製品反り検査区分
    KPRSKWAY As String * 2          ' 購製品反り検査方法
    KPRSSZYO As String * 1          ' 購製品反り測定条件
    KPRSZARA As Integer             ' 購製品反り除外領域
    KPRSSDEV As Double              ' 購製品反り標準偏差
    KPRSAMIN As Double              ' 購製品反り平均下限
    KPRSAMAX As Double              ' 購製品反り平均上限
    KPRSSREC As String * 1          ' 購製品反り測定器
    KPRSBO1 As Double               ' 購製品反り境界１
    KPRSBO1B As Integer             ' 購製品反り境界１下
    KPRSBO2 As Double               ' 購製品反り境界２
    KPRSBO2B As Integer             ' 購製品反り境界２下
    KPRSBO3 As Double               ' 購製品反り境界３
    KPRSBO3B As Integer             ' 購製品反り境界３下
    KPRWARMX As Double              ' 購製品ＷＡＲＰ上限
    KPRWARSZ As String * 1          ' 購製品ＷＡＲＰ測定条件
    KPRWARHT As String * 1          ' 購製品ＷＡＲＰ保証方法＿対
    KPRWARHS As String * 1          ' 購製品ＷＡＲＰ保証方法＿処
    KPRWARKB As String * 1          ' 購製品ＷＡＲＰ検査区分
    KPRWARKW As String * 2          ' 購製品ＷＡＲＰ検査方法
    KPRWARZA As Integer             ' 購製品ＷＡＲＰ除外領域
    KPRWARSR As String * 1          ' 購製品ＷＡＲＰ測定器
    KPRWAB1 As Double               ' 購製品ＷＡＲＰ境界１
    KPRWAB1B As Integer             ' 購製品ＷＡＲＰ境界１下
    KPRWAB2 As Double               ' 購製品ＷＡＲＰ境界２
    KPRWAB2B As Integer             ' 購製品ＷＡＲＰ境界２下
    KPRWAB3 As Double               ' 購製品ＷＡＲＰ境界３
    KPRWAB3B As Integer             ' 購製品ＷＡＲＰ境界３下
    KPRFKKBN As String * 1          ' 購製品平坦検査区分
    KPRFSZYO As String * 1          ' 購製品平坦測定条件
    KPRFSREC As String * 1          ' 購製品平坦測定器
    KPRGBMAX As Double              ' 購製品平坦ＧＢ上限
    KPRGBPUG As Double              ' 購製品平坦ＧＢＰＵＡ限
    KPRGBPUR As Integer             ' 購製品平坦ＧＢＰＵＡ率
    KPRGBHWT As String * 1          ' 購製品平坦ＧＢ保証方法＿対
    KPRGBHWS As String * 1          ' 購製品平坦ＧＢ保証方法＿処
    KPRGBKW As String * 4           ' 購製品平坦ＧＢ検査方法
    KPRGBKWO As String * 4          ' 購製品平坦ＧＢ検査方法旧
    KPRGBZAR As Integer             ' 購製品平坦ＧＢ除外領域
    KPRGBB1 As Double               ' 購製品平坦ＧＢ境界１
    KPRGBB1B As Integer             ' 購製品平坦ＧＢ境界１下
    KPRGBB2 As Double               ' 購製品平坦ＧＢ境界２
    KPRGBB2B As Integer             ' 購製品平坦ＧＢ境界２下
    KPRGBB3 As Double               ' 購製品平坦ＧＢ境界３
    KPRGBB3B As Integer             ' 購製品平坦ＧＢ境界３下
    KPRGFDMX As Double              ' 購製品平坦ＧＦＤ上限
    KPRGFDPG As Double              ' 購製品平坦ＧＦＤＰＵＡ限
    KPRGFDPR As Integer             ' 購製品平坦ＧＦＤＰＵＡ率
    KPRGFDHT As String * 1          ' 購製品平坦ＧＦＤ保証方法＿対
    KPRGFDHS As String * 1          ' 購製品平坦ＧＦＤ保証方法＿処
    KPRGFDBM As String * 1          ' 購製品平坦ＧＦＤ基準面
    KPRGFDKW As String * 4          ' 購製品平坦ＧＦＤ検査方法
    KPRGFDKO As String * 4          ' 購製品平坦ＧＦＤ検査方法旧
    KPRGFDZA As Integer             ' 購製品平坦ＧＦＤ除外領域
    KPRGDB1 As Double               ' 購製品平坦ＧＦＤ境界１
    KPRGDB1B As Integer             ' 購製品平坦ＧＦＤ境界１下
    KPRGDB2 As Double               ' 購製品平坦ＧＦＤ境界２
    KPRGDB2B As Integer             ' 購製品平坦ＧＦＤ境界２下
    KPRGDB3 As Double               ' 購製品平坦ＧＦＤ境界３
    KPRGDB3B As Integer             ' 購製品平坦ＧＦＤ境界３下
    KPRGFRMX As Double              ' 購製品平坦ＧＦＲ上限
    KPRGFRPG As Double              ' 購製品平坦ＧＦＲＰＵＡ限
    KPRGFRPR As Integer             ' 購製品平坦ＧＦＲＰＵＡ率
    KPRGFRHT As String * 1          ' 購製品平坦ＧＦＲ保証方法＿対
    KPRGFRHS As String * 1          ' 購製品平坦ＧＦＲ保証方法＿処
    KPRGFRBM As String * 1          ' 購製品平坦ＧＦＲ基準面
    KPRGFRKW As String * 4          ' 購製品平坦ＧＦＲ検査方法
    KPRGFRKO As String * 4          ' 購製品平坦ＧＦＲ検査方法旧
    KPRGFRZA As Integer             ' 購製品平坦ＧＦＲ除外領域
    KPRGRB1 As Double               ' 購製品平坦ＧＦＲ境界１
    KPRGRB1B As Integer             ' 購製品平坦ＧＦＲ境界１下
    KPRGRB2 As Double               ' 購製品平坦ＧＦＲ境界２
    KPRGRB2B As Integer             ' 購製品平坦ＧＦＲ境界２下
    KPRGRB3 As Double               ' 購製品平坦ＧＦＲ境界３
    KPRGRB3B As Integer             ' 購製品平坦ＧＦＲ境界３下
    KPRSBMAX As Double              ' 購製品平坦ＳＢ上限
    KPRSBPUG As Double              ' 購製品平坦ＳＢＰＵＡ限
    KPRSBPUR As Integer             ' 購製品平坦ＳＢＰＵＡ率
    KPRSBSZX As Double              ' 購製品平坦ＳＢサイズＸ
    KPRSBSZY As Double              ' 購製品平坦ＳＢサイズＹ
    KPRSBHWT As String * 1          ' 購製品平坦ＳＢ保証方法＿対
    KPRSBHWS As String * 1          ' 購製品平坦ＳＢ保証方法＿処
    KPRSBBM As String * 1           ' 購製品平坦ＳＢ基準面
    KPRSBKW As String * 4           ' 購製品平坦ＳＢ検査方法
    KPRSBKWO As String * 4          ' 購製品平坦ＳＢ検査方法旧
    KPRSBZAR As Integer             ' 購製品平坦ＳＢ除外領域
    KPRSBB1 As Double               ' 購製品平坦ＳＢ境界１
    KPRSBB1B As Integer             ' 購製品平坦ＳＢ境界１下
    KPRSBB2 As Double               ' 購製品平坦ＳＢ境界２
    KPRSBB2B As Integer             ' 購製品平坦ＳＢ境界２下
    KPRSBB3 As Double               ' 購製品平坦ＳＢ境界３
    KPRSBB3B As Integer             ' 購製品平坦ＳＢ境界３下
    KPRSFMAX As Double              ' 購製品平坦ＳＦ上限
    KPRSFPUG As Double              ' 購製品平坦ＳＦＰＵＡ限
    KPRSFPUR As Integer             ' 購製品平坦ＳＦＰＵＡ率
    KPRSFSZX As Double              ' 購製品平坦ＳＦサイズＸ
    KPRSFSZY As Double              ' 購製品平坦ＳＦサイズＹ
    KPRSFHWT As String * 1          ' 購製品平坦ＳＦ保証方法＿対
    KPRSFHWS As String * 1          ' 購製品平坦ＳＦ保証方法＿処
    KPRSFBM As String * 1           ' 購製品平坦ＳＦ基準面
    KPRSFKW As String * 4           ' 購製品平坦ＳＦ検査方法
    KPRSFKWO As String * 4          ' 購製品平坦ＳＦ検査方法旧
    KPRSFZAR As Integer             ' 購製品平坦ＳＦ除外領域
    KPRSFB1 As Double               ' 購製品平坦ＳＦ境界１
    KPRSFB1B As Integer             ' 購製品平坦ＳＦ境界１下
    KPRSFB2 As Double               ' 購製品平坦ＳＦ境界２
    KPRSFB2B As Integer             ' 購製品平坦ＳＦ境界２下
    KPRSFB3 As Double               ' 購製品平坦ＳＦ境界３
    KPRSFB3B As Integer             ' 購製品平坦ＳＦ境界３下
    KPRFSXOF As Double              ' 購製品平坦サイトＸＯＦ
    KPRFSYOF As Double              ' 購製品平坦サイトＹＯＦ
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 顧客仕様WFﾃﾞｰﾀ８
Public Type typ_TBCME015
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KMGSRRNO As String * 9          ' 購管理仕様登録依頼番号
    KPRMK1SI As Double              ' 購製品面検欠陥１サイズ
    KPRMK1MX As Integer             ' 購製品面検欠陥１上限
    KPRMK1SZ As String * 1          ' 購製品面検欠陥１測定条件
    KPRMK1ZA As Integer             ' 購製品面検欠陥１除外領域
    KPRMK1HT As String * 1          ' 購製品面検欠陥１保証方法＿対
    KPRMK1HS As String * 1          ' 購製品面検欠陥１保証方法＿処
    KPRM1B1 As Integer              ' 購製品面検欠陥１境界１
    KPRM1B1B As Integer             ' 購製品面検欠陥１境界１下
    KPRM1B2 As Integer              ' 購製品面検欠陥１境界２
    KPRM1B2B As Integer             ' 購製品面検欠陥１境界２下
    KPRM1B3 As Integer              ' 購製品面検欠陥１境界３
    KPRM1B3B As Integer             ' 購製品面検欠陥１境界３下
    KPRMK2SI As Double              ' 購製品面検欠陥２サイズ
    KPRMK2MX As Integer             ' 購製品面検欠陥２上限
    KPRMK2HT As String * 1          ' 購製品面検欠陥２保証方法＿対
    KPRMK2HS As String * 1          ' 購製品面検欠陥２保証方法＿処
    KPRM2B1 As Integer              ' 購製品面検欠陥２境界１
    KPRM2B1B As Integer             ' 購製品面検欠陥２境界１下
    KPRM2B2 As Integer              ' 購製品面検欠陥２境界２
    KPRM2B2B As Integer             ' 購製品面検欠陥２境界２下
    KPRM2B3 As Integer              ' 購製品面検欠陥２境界３
    KPRM2B3B As Integer             ' 購製品面検欠陥２境界３下
    KPRMK3SI As Double              ' 購製品面検欠陥３サイズ
    KPRMK3MX As Integer             ' 購製品面検欠陥３上限
    KPRMK3HT As String * 1          ' 購製品面検欠陥３保証方法＿対
    KPRMK3HS As String * 1          ' 購製品面検欠陥３保証方法＿処
    KPRM3B1 As Integer              ' 購製品面検欠陥３境界１
    KPRM3B1B As Integer             ' 購製品面検欠陥３境界１下
    KPRM3B2 As Integer              ' 購製品面検欠陥３境界２
    KPRM3B2B As Integer             ' 購製品面検欠陥３境界２下
    KPRM3B3 As Integer              ' 購製品面検欠陥３境界３
    KPRM3B3B As Integer             ' 購製品面検欠陥３境界３下
    KPRMK4SI As Double              ' 購製品面検欠陥４サイズ
    KPRMK4MX As Integer             ' 購製品面検欠陥４上限
    KPRMK4HT As String * 1          ' 購製品面検欠陥４保証方法＿対
    KPRMK4HS As String * 1          ' 購製品面検欠陥４保証方法＿処
    KPRM4B1 As Integer              ' 購製品面検欠陥４境界１
    KPRM4B1B As Integer             ' 購製品面検欠陥４境界１下
    KPRM4B2 As Integer              ' 購製品面検欠陥４境界２
    KPRM4B2B As Integer             ' 購製品面検欠陥４境界２下
    KPRM4B3 As Integer              ' 購製品面検欠陥４境界３
    KPRM4B3B As Integer             ' 購製品面検欠陥４境界３下
    KPRMB1SI As Double              ' 購製品面検欠陥裏１サイズ
    KPRMB1MX As Integer             ' 購製品面検欠陥裏１上限
    KPRMB1SZ As String * 1          ' 購製品面検欠陥裏１測定条件
    KPRMB1ZA As Integer             ' 購製品面検欠陥裏１除外領域
    KPRMB1HT As String * 1          ' 購製品面検欠陥裏１保証方法＿対
    KPRMB1HS As String * 1          ' 購製品面検欠陥裏１保証方法＿処
    KPRMB2SI As Double              ' 購製品面検欠陥裏２サイズ
    KPRMB2MX As Integer             ' 購製品面検欠陥裏２上限
    KPRMB2SZ As String * 1          ' 購製品面検欠陥裏２測定条件
    KPRMB2ZA As Integer             ' 購製品面検欠陥裏２除外領域
    KPRMB2HT As String * 1          ' 購製品面検欠陥裏２保証方法＿対
    KPRMB2HS As String * 1          ' 購製品面検欠陥裏２保証方法＿処
    KPRMKSRE As String * 1          ' 購製品面検欠陥測定器
    KPRMPIPT As String * 1          ' 購製品面検欠陥ＰＩＰ検査
    KPRMPIPK As Integer             ' 購製品面検欠陥ＰＩＰ個数
    KPRMPISH As String * 1          ' 購製品面検ＰＩＰ測定位置＿方
    KPRMPIST As String * 1          ' 購製品面検ＰＩＰ測定位置＿点
    KPRMPISI As String * 1          ' 購製品面検ＰＩＰ測定位置＿位
    KPRMPIKM As String * 1          ' 購製品面検ＰＩＰ検査頻度＿枚
    KPRMPIKN As String * 1          ' 購製品面検ＰＩＰ検査頻度＿抜
    KPRMPIKH As String * 1          ' 購製品面検ＰＩＰ検査頻度＿保
    KPRMPIKU As String * 1          ' 購製品面検ＰＩＰ検査頻度＿ウ
    KPRMNIND As String * 2          ' 購製品金属濃度指数
    KPRMNMAX As Double              ' 購製品金属濃度上限
    KPRMNALX As Double              ' 購製品金属濃度ＡＬ上限
    KPRMNCAX As Double              ' 購製品金属濃度ＣＡ上限
    KPRMNCRX As Double              ' 購製品金属濃度ＣＲ上限
    KPRMNCUX As Double              ' 購製品金属濃度ＣＵ上限
    KPRMNFEX As Double              ' 購製品金属濃度ＦＥ上限
    KPRMNKMX As Double              ' 購製品金属濃度Ｋ上限
    KPRMNMGX As Double              ' 購製品金属濃度ＭＧ上限
    KPRMNNAX As Double              ' 購製品金属濃度ＮＡ上限
    KPRMNNIX As Double              ' 購製品金属濃度ＮＩ上限
    KPRMNZNX As Double              ' 購製品金属濃度ＺＮ上限
    KPRMNKWY As String * 2          ' 購製品金属濃度検査方法
    KPRMNZAR As Integer             ' 購製品金属濃度除外領域
    KPRMNKHM As String * 1          ' 購製品金属濃度検査頻度＿枚
    KPRMNKHN As String * 1          ' 購製品金属濃度検査頻度＿抜
    KPRMNKHH As String * 1          ' 購製品金属濃度検査頻度＿保
    KPRMNKHU As String * 1          ' 購製品金属濃度検査頻度＿ウ
    KPRSPVMX As Double              ' 購製品ＳＰＶＦＥ上限
    KPRSPVKM As String * 1          ' 購製品ＳＰＶＦＥ検査頻度＿枚
    KPRSPVKN As String * 1          ' 購製品ＳＰＶＦＥ検査頻度＿抜
    KPRSPVKH As String * 1          ' 購製品ＳＰＶＦＥ検査頻度＿保
    KPRSPVKU As String * 1          ' 購製品ＳＰＶＦＥ検査頻度＿ウ
    KPRSPVIN As String * 2          ' 購製品ＳＰＶＦＥ指数
    KPRDLMIN As Integer             ' 購製品拡散長下限
    KPRDLMAX As Integer             ' 購製品拡散長上限
    KPRDLSPH As String * 1          ' 購製品拡散長測定位置＿方
    KPRDLSPT As String * 1          ' 購製品拡散長測定位置＿点
    KPRDLSPI As String * 1          ' 購製品拡散長測定位置＿位
    KPRDLHWT As String * 1          ' 購製品拡散長保証方法＿対
    KPRDLHWS As String * 1          ' 購製品拡散長保証方法＿処
    KPRDLKHM As String * 1          ' 購製品拡散長検査頻度＿枚
    KPRDLKHN As String * 1          ' 購製品拡散長検査頻度＿抜
    KPRDLKHH As String * 1          ' 購製品拡散長検査頻度＿保
    KPRDLKHU As String * 1          ' 購製品拡散長検査頻度＿ウ
    KPROTMIN As Double              ' 購製品酸化膜耐圧下限
    KPROTSPH As String * 1          ' 購製品酸化膜耐圧測定位置＿方
    KPROTSPT As String * 1          ' 購製品酸化膜耐圧測定位置＿点
    KPROTSPI As String * 1          ' 購製品酸化膜耐圧測定位置＿位
    KPROTKWY As String * 2          ' 購製品酸化膜耐圧検査方法
    KPROTZAR As Integer             ' 購製品酸化膜耐圧除外領域
    KPROTKHM As String * 1          ' 購製品酸化膜耐圧検査頻度＿枚
    KPROTKHN As String * 1          ' 購製品酸化膜耐圧検査頻度＿抜
    KPROTKHH As String * 1          ' 購製品酸化膜耐圧検査頻度＿保
    KPROTKHU As String * 1          ' 購製品酸化膜耐圧検査頻度＿ウ
    KPROTMX1 As Double              ' 購製品酸化膜耐圧上限１
    KPROTMX2 As Double              ' 購製品酸化膜耐圧上限２
    KPROTKW1 As String * 2          ' 購製品酸化膜耐圧検査方法１
    KPROTKW2 As String * 2          ' 購製品酸化膜耐圧検査方法２
    KPROTHWT As String * 1          ' 購製品酸化膜耐圧保証方法＿対
    KPROTHWS As String * 1          ' 購製品酸化膜耐圧保証方法＿処
    KPRLTDCX As Double              ' 購製品ＬＴＤ濃度ＣＵ上限
    KPRLTDIN As String * 2          ' 購製品ＬＴＤ濃度指数
    KPRLTDKW As String * 2          ' 購製品ＬＴＤ濃度検査方法
    KPRLTDSH As String * 1          ' 購製品ＬＴＤ濃度測定位置＿方
    KPRLTDST As String * 1          ' 購製品ＬＴＤ濃度測定位置＿点
    KPRLTDSI As String * 1          ' 購製品ＬＴＤ濃度測定位置＿位
    KPRLTDHT As String * 1          ' 購製品ＬＴＤ濃度保証方法＿対
    KPRLTDHS As String * 1          ' 購製品ＬＴＤ濃度保証方法＿処
    KPRLTDKM As String * 1          ' 購製品ＬＴＤ濃度検査頻度＿枚
    KPRLTDKN As String * 1          ' 購製品ＬＴＤ濃度検査頻度＿抜
    KPRLTDKH As String * 1          ' 購製品ＬＴＤ濃度検査頻度＿保
    KPRLTDKU As String * 1          ' 購製品ＬＴＤ濃度検査頻度＿ウ
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 顧客仕様WFﾃﾞｰﾀ９
Public Type typ_TBCME016
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KMGSRRNO As String * 9          ' 購管理仕様登録依頼番号
    KPROS1AX As Double              ' 購製品ＯＳＦ１平均上限
    KPROS1MX As Double              ' 購製品ＯＳＦ１上限
    KPROS1O1 As Integer             ' 購製品ＯＳＦ１処理温度１
    KPROS1T1 As Integer             ' 購製品ＯＳＦ１処理時間１
    KPROS1GS As String * 1          ' 購製品ＯＳＦ１雰囲気ガス
    KPROS1ET As Integer             ' 購製品ＯＳＦ１選択ＥＴ代
    KPROS1NS As String * 2          ' 購製品ＯＳＦ１熱処理法
    KPROS1SZ As String * 1          ' 購製品ＯＳＦ１測定条件
    KPROS1SH As String * 1          ' 購製品ＯＳＦ１測定位置＿方
    KPROS1ST As String * 1          ' 購製品ＯＳＦ１測定位置＿点
    KPROS1SR As String * 1          ' 購製品ＯＳＦ１測定位置＿領
    KPROS1HT As String * 1          ' 購製品ＯＳＦ１保証方法＿対
    KPROS1HS As String * 1          ' 購製品ＯＳＦ１保証方法＿処
    KPROS1KB As String * 1          ' 購製品ＯＳＦ１検査区分
    KPROS1KM As String * 1          ' 購製品ＯＳＦ１検査頻度＿枚
    KPROS1KN As String * 1          ' 購製品ＯＳＦ１検査頻度＿抜
    KPROS1KH As String * 1          ' 購製品ＯＳＦ１検査頻度＿保
    KPROS1KU As String * 1          ' 購製品ＯＳＦ１検査頻度＿ウ
    KPROS2AX As Double              ' 購製品ＯＳＦ２平均上限
    KPROS2MX As Double              ' 購製品ＯＳＦ２上限
    KPROS2O1 As Integer             ' 購製品ＯＳＦ２処理温度１
    KPROS2T1 As Integer             ' 購製品ＯＳＦ２処理時間１
    KPROS2GS As String * 1          ' 購製品ＯＳＦ２雰囲気ガス
    KPROS2ET As Integer             ' 購製品ＯＳＦ２選択ＥＴ代
    KPROS2NS As String * 2          ' 購製品ＯＳＦ２熱処理法
    KPROS2SZ As String * 1          ' 購製品ＯＳＦ２測定条件
    KPROS2SH As String * 1          ' 購製品ＯＳＦ２測定位置＿方
    KPROS2ST As String * 1          ' 購製品ＯＳＦ２測定位置＿点
    KPROS2SR As String * 1          ' 購製品ＯＳＦ２測定位置＿領
    KPROS2HT As String * 1          ' 購製品ＯＳＦ２保証方法＿対
    KPROS2HS As String * 1          ' 購製品ＯＳＦ２保証方法＿処
    KPROS2KB As String * 1          ' 購製品ＯＳＦ２検査区分
    KPROS2KM As String * 1          ' 購製品ＯＳＦ２検査頻度＿枚
    KPROS2KN As String * 1          ' 購製品ＯＳＦ２検査頻度＿抜
    KPROS2KH As String * 1          ' 購製品ＯＳＦ２検査頻度＿保
    KPROS2KU As String * 1          ' 購製品ＯＳＦ２検査頻度＿ウ
    KPROS3AX As Double              ' 購製品ＯＳＦ３平均上限
    KPROS3MX As Double              ' 購製品ＯＳＦ３上限
    KPROS3O1 As Integer             ' 購製品ＯＳＦ３処理温度１
    KPROS3T1 As Integer             ' 購製品ＯＳＦ３処理時間１
    KPROS3GS As String * 1          ' 購製品ＯＳＦ３雰囲気ガス
    KPROS3ET As Integer             ' 購製品ＯＳＦ３選択ＥＴ代
    KPROS3NS As String * 2          ' 購製品ＯＳＦ３熱処理法
    KPROS3SZ As String * 1          ' 購製品ＯＳＦ３測定条件
    KPROS3SH As String * 1          ' 購製品ＯＳＦ３測定位置＿方
    KPROS3ST As String * 1          ' 購製品ＯＳＦ３測定位置＿点
    KPROS3SR As String * 1          ' 購製品ＯＳＦ３測定位置＿領
    KPROS3HT As String * 1          ' 購製品ＯＳＦ３保証方法＿対
    KPROS3HS As String * 1          ' 購製品ＯＳＦ３保証方法＿処
    KPROS3KB As String * 1          ' 購製品ＯＳＦ３検査区分
    KPROS3KM As String * 1          ' 購製品ＯＳＦ３検査頻度＿枚
    KPROS3KN As String * 1          ' 購製品ＯＳＦ３検査頻度＿抜
    KPROS3KH As String * 1          ' 購製品ＯＳＦ３検査頻度＿保
    KPROS3KU As String * 1          ' 購製品ＯＳＦ３検査頻度＿ウ
    KPROS4AX As Double              ' 購製品ＯＳＦ４平均上限
    KPROS4MX As Double              ' 購製品ＯＳＦ４上限
    KPROS4O1 As Integer             ' 購製品ＯＳＦ４処理温度１
    KPROS4T1 As Integer             ' 購製品ＯＳＦ４処理時間１
    KPROS4GS As String * 1          ' 購製品ＯＳＦ４雰囲気ガス
    KPROS4ET As Integer             ' 購製品ＯＳＦ４選択ＥＴ代
    KPROS4NS As String * 2          ' 購製品ＯＳＦ４熱処理法
    KPROS4SZ As String * 1          ' 購製品ＯＳＦ４測定条件
    KPROS4SH As String * 1          ' 購製品ＯＳＦ４測定位置＿方
    KPROS4ST As String * 1          ' 購製品ＯＳＦ４測定位置＿点
    KPROS4SR As String * 1          ' 購製品ＯＳＦ４測定位置＿領
    KPROS4HT As String * 1          ' 購製品ＯＳＦ４保証方法＿対
    KPROS4HS As String * 1          ' 購製品ＯＳＦ４保証方法＿処
    KPROS4KB As String * 1          ' 購製品ＯＳＦ４検査区分
    KPROS4KM As String * 1          ' 購製品ＯＳＦ４検査頻度＿枚
    KPROS4KN As String * 1          ' 購製品ＯＳＦ４検査頻度＿抜
    KPROS4KH As String * 1          ' 購製品ＯＳＦ４検査頻度＿保
    KPROS4KU As String * 1          ' 購製品ＯＳＦ４検査頻度＿ウ
    KPRBM1AN As Double              ' 購製品ＢＭＤ１平均下限
    KPRBM1AX As Double              ' 購製品ＢＭＤ１平均上限
    KPRBM1GS As String * 1          ' 購製品ＢＭＤ１雰囲気ガス
    KPRBM1ET As Integer             ' 購製品ＢＭＤ１選択ＥＴ代
    KPRBM1NS As String * 2          ' 購製品ＢＭＤ１熱処理法
    KPRBM1SZ As String * 1          ' 購製品ＢＭＤ１測定条件
    KPRBM1SH As String * 1          ' 購製品ＢＭＤ１測定位置＿方
    KPRBM1ST As String * 1          ' 購製品ＢＭＤ１測定位置＿点
    KPRBM1SR As String * 1          ' 購製品ＢＭＤ１測定位置＿領
    KPRBM1HT As String * 1          ' 購製品ＢＭＤ１保証方法＿対
    KPRBM1HS As String * 1          ' 購製品ＢＭＤ１保証方法＿処
    KPRBM1KB As String * 1          ' 購製品ＢＭＤ１検査区分
    KPRBM1KM As String * 1          ' 購製品ＢＭＤ１検査頻度＿枚
    KPRBM1KN As String * 1          ' 購製品ＢＭＤ１検査頻度＿抜
    KPRBM1KH As String * 1          ' 購製品ＢＭＤ１検査頻度＿保
    KPRBM1KU As String * 1          ' 購製品ＢＭＤ１検査頻度＿ウ
    KPRBM2AN As Double              ' 購製品ＢＭＤ２平均下限
    KPRBM2AX As Double              ' 購製品ＢＭＤ２平均上限
    KPRBM2GS As String * 1          ' 購製品ＢＭＤ２雰囲気ガス
    KPRBM2ET As Integer             ' 購製品ＢＭＤ２選択ＥＴ代
    KPRBM2NS As String * 2          ' 購製品ＢＭＤ２熱処理法
    KPRBM2SZ As String * 1          ' 購製品ＢＭＤ２測定条件
    KPRBM2SH As String * 1          ' 購製品ＢＭＤ２測定位置＿方
    KPRBM2ST As String * 1          ' 購製品ＢＭＤ２測定位置＿点
    KPRBM2SR As String * 1          ' 購製品ＢＭＤ２測定位置＿領
    KPRBM2HT As String * 1          ' 購製品ＢＭＤ２保証方法＿対
    KPRBM2HS As String * 1          ' 購製品ＢＭＤ２保証方法＿処
    KPRBM2KB As String * 1          ' 購製品ＢＭＤ２検査区分
    KPRBM2KM As String * 1          ' 購製品ＢＭＤ２検査頻度＿枚
    KPRBM2KN As String * 1          ' 購製品ＢＭＤ２検査頻度＿抜
    KPRBM2KH As String * 1          ' 購製品ＢＭＤ２検査頻度＿保
    KPRBM2KU As String * 1          ' 購製品ＢＭＤ２検査頻度＿ウ
    KPRBM3AN As Double              ' 購製品ＢＭＤ３平均下限
    KPRBM3AX As Double              ' 購製品ＢＭＤ３平均上限
    KPRBM3GS As String * 1          ' 購製品ＢＭＤ３雰囲気ガス
    KPRBM3ET As Integer             ' 購製品ＢＭＤ３選択ＥＴ代
    KPRBM3NS As String * 2          ' 購製品ＢＭＤ３熱処理法
    KPRBM3SZ As String * 1          ' 購製品ＢＭＤ３測定条件
    KPRBM3SH As String * 1          ' 購製品ＢＭＤ３測定位置＿方
    KPRBM3ST As String * 1          ' 購製品ＢＭＤ３測定位置＿点
    KPRBM3SR As String * 1          ' 購製品ＢＭＤ３測定位置＿領
    KPRBM3HT As String * 1          ' 購製品ＢＭＤ３保証方法＿対
    KPRBM3HS As String * 1          ' 購製品ＢＭＤ３保証方法＿処
    KPRBM3KB As String * 1          ' 購製品ＢＭＤ３検査区分
    KPRBM3KM As String * 1          ' 購製品ＢＭＤ３検査頻度＿枚
    KPRBM3KN As String * 1          ' 購製品ＢＭＤ３検査頻度＿抜
    KPRBM3KH As String * 1          ' 購製品ＢＭＤ３検査頻度＿保
    KPRBM3KU As String * 1          ' 購製品ＢＭＤ３検査頻度＿ウ
    KPRBMDVO As String * 1          ' 購製品ＢＭＤ体積換算有無
    KPROSPAX As Integer             ' 購製品ＯＳＰ平均上限
    KPROSPMX As Integer             ' 購製品ＯＳＰ上限
    KPROSPSH As String * 1          ' 購製品ＯＳＰ測定位置＿方
    KPROSPST As String * 1          ' 購製品ＯＳＰ測定位置＿点
    KPROSPSR As String * 1          ' 購製品ＯＳＰ測定位置＿領
    KPROSPHT As String * 1          ' 購製品ＯＳＰ保証方法＿対
    KPROSPHS As String * 1          ' 購製品ＯＳＰ保証方法＿処
    KPROSPNS As String * 2          ' 購製品ＯＳＰ熱処理法
    KPROSPSZ As String * 1          ' 購製品ＯＳＰ測定条件
    KPROSPKM As String * 1          ' 購製品ＯＳＰ検査頻度＿枚
    KPROSPKN As String * 1          ' 購製品ＯＳＰ検査頻度＿抜
    KPROSPKH As String * 1          ' 購製品ＯＳＰ検査頻度＿保
    KPROSPKU As String * 1          ' 購製品ＯＳＰ検査頻度＿ウ
    KPROSPET As Integer             ' 購製品ＯＳＰ選択ＥＴ代
    KPRTSPHM As String * 1          ' 購製品トレスサンプル頻度＿枚
    KPRTSPHN As String * 1          ' 購製品トレスサンプル頻度＿抜
    KPRTSPHH As String * 1          ' 購製品トレスサンプル頻度＿保
    KPRTSPHU As String * 1          ' 購製品トレスサンプル頻度＿ウ
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 製品仕様管理
Public Type typ_TBCME017
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    SHNAME As String * 11           ' 社内品名
    HMGSTRRNO As String * 9         ' 品管理仕様登録依頼番号
    HMGSTFNO As String * 8          ' 品管理社員Ｎｏ
    HMGSXSNO As String * 6          ' 品管理ＳＸ製品番号
    HMGSXSNE As Integer             ' 品管理ＳＸ製品番号枝番
    HMGWFSNO As String * 6          ' 品管理ＷＦ製品番号
    HMGWFSNE As Integer             ' 品管理ＷＦ製品番号枝番
    HMGEPSNO As String * 6          ' 品管理ＥＰ製品番号
    HMGEPSNE As Integer             ' 品管理ＥＰ製品番号枝番
    SPECMWAY As String * 1          ' 仕様作成方法
    UNIFLAG As String * 1           ' 統合フラグ
    CONFLAG As String * 1           ' 確認フラグ
    REINFLAG As String * 1          ' 再付与フラグ
    HMGRDIAM As Integer             ' 品管理代表直径
    HMGWKBN As String * 2           ' 品管理製法区分
    HMGPMKBN As String * 1          ' 品管理設計管理区分
    HSXSKBN As String * 1           ' 品ＳＸ製品区分
    HSXNCKBN As String * 1          ' 品ＳＸノッチ区分
    HWFSKBN As String * 1           ' 品ＷＦ製品区分
    HWFRKBNK As String * 1          ' 品ＷＦ比抵抗区分種別
    HWFAKBUM As String * 1          ' 品ＷＦ厚み区分有無
    HWFOXKBN As String * 1          ' 品ＷＦ酸素区分
    HWFIGKBN As String * 1          ' 品ＷＦＩＧ区分
    HWFNCKBN As String * 1          ' 品ＷＦノッチ区分
    HWFCMPKU As String * 1          ' 品ＷＦＣＭＰ加工有無
    HWFSZKBN As String * 1          ' 品ＷＦ支給材料区分
    HWFSZMUM As String * 1          ' 品ＷＦ支給材料面取有無
    HWFKZKBN As String * 1          ' 品ＷＦ購入材料区分
    HWFHGRAD As String * 7          ' 品ＷＦ品質グレード
    HEPSKBN As String * 1           ' 品ＥＰ製品区分
    HEPRKBNU As String * 1          ' 品ＥＰ比抵抗区分有無
    HEPAKBUM As String * 1          ' 品ＥＰ厚み区分有無
    HEPSZKBN As String * 1          ' 品ＥＰ支給材料区分
    HEPKZKBN As String * 1          ' 品ＥＰ購入材料区分
    HMGTRKSI As String * 1          ' 品管理ＴＲＫ＃指定
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 製品仕様SXLﾃﾞｰﾀ１
Public Type typ_TBCME018
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    HMGSTRRNO As String * 9         ' 品管理仕様登録依頼番号
    HMGSTFNO As String * 8          ' 品管理社員Ｎｏ
    HMGSXSNO As String * 6          ' 品管理ＳＸ製品番号
    HMGSXSNE As Integer             ' 品管理ＳＸ製品番号枝番
    CONFLAG As String * 1           ' 確認フラグ
    REINFLAG As String * 1          ' 再付与フラグ
    HSXTRWKB As String * 1          ' 品ＳＸ統合可否区分
    HSXTYPE As String * 1           ' 品ＳＸタイプ
    KSXTYPKW As String * 1          ' 品ＳＸタイプ検査方法
    HSXDOP As String * 1            ' 品ＳＸドーパント
    HSXRMIN As Double               ' 品ＳＸ比抵抗下限
    HSXRMAX As Double               ' 品ＳＸ比抵抗上限
    HSXRSPOH As String * 1          ' 品ＳＸ比抵抗測定位置＿方
    HSXRSPOT As String * 1          ' 品ＳＸ比抵抗測定位置＿点
    HSXRSPOI As String * 1          ' 品ＳＸ比抵抗測定位置＿位
    HSXRHWYT As String * 1          ' 品ＳＸ比抵抗保証方法＿対
    HSXRHWYS As String * 1          ' 品ＳＸ比抵抗保証方法＿処
    HSXRKWAY As String * 2          ' 品ＳＸ比抵抗検査方法
    HSXRKHNM As String * 1          ' 品ＳＸ比抵抗検査頻度＿枚
    HSXRKHNI As String * 1          ' 品ＳＸ比抵抗検査頻度＿位
    HSXRKHNH As String * 1          ' 品ＳＸ比抵抗検査頻度＿保
    HSXRKHNS As String * 1          ' 品ＳＸ比抵抗検査頻度＿試
    HSXRMCAL As String * 1          ' 品ＳＸ比抵抗面内計算
    HSXRMBNP As Double              ' 品ＳＸ比抵抗面内分布
    HSXRMCL2 As String * 1          ' 品ＳＸ比抵抗面内計算２
    HSXRMBP2 As Double              ' 品ＳＸ比抵抗面内分布２
    HSXRSDEV As Double              ' 品ＳＸ比抵抗標準偏差
    HSXRAMIN As Double              ' 品ＳＸ比抵抗平均下限
    HSXRAMAX As Double              ' 品ＳＸ比抵抗平均上限
    HSXFORM As String * 1           ' 品ＳＸ形状
    HSXD1CEN As Double              ' 品ＳＸ直径１中心
    HSXD1MIN As Double              ' 品ＳＸ直径１下限
    HSXD1MAX As Double              ' 品ＳＸ直径１上限
    HSXD2CEN As Double              ' 品ＳＸ直径２中心
    HSXD2MIN As Double              ' 品ＳＸ直径２下限
    HSXD2MAX As Double              ' 品ＳＸ直径２上限
    HSXCDIR As String * 1           ' 品ＳＸ結晶面方位
    HSXCSCEN As Double              ' 品ＳＸ結晶面傾中心
    HSXCSMIN As Double              ' 品ＳＸ結晶面傾下限
    HSXCSMAX As Double              ' 品ＳＸ結晶面傾上限
    HSXCKWAY As String * 2          ' 品ＳＸ結晶面検査方法
    HSXCKHNM As String * 1          ' 品ＳＸ結晶面検査頻度＿枚
    HSXCKHNI As String * 1          ' 品ＳＸ結晶面検査頻度＿位
    HSXCKHNH As String * 1          ' 品ＳＸ結晶面検査頻度＿保
    HSXCKHNS As String * 1          ' 品ＳＸ結晶面検査頻度＿試
    HSXCSDIR As String * 2          ' 品ＳＸ結晶面傾方位
    HSXCSDIS As String * 1          ' 品ＳＸ結晶面傾方位指定
    HSXCTDIR As String * 2          ' 品ＳＸ結晶面傾縦方位
    HSXCTCEN As Double              ' 品ＳＸ結晶面傾縦中心
    HSXCTMIN As Double              ' 品ＳＸ結晶面傾縦下限
    HSXCTMAX As Double              ' 品ＳＸ結晶面傾縦上限
    HSXCYDIR As String * 2          ' 品ＳＸ結晶面傾横方位
    HSXCYCEN As Double              ' 品ＳＸ結晶面傾横中心
    HSXCYMIN As Double              ' 品ＳＸ結晶面傾横下限
    HSXCYMAX As Double              ' 品ＳＸ結晶面傾横上限
    HSXOF1PD As String * 2          ' 品ＳＸＯＦ１位置方位
    HSXOF1PN As Double              ' 品ＳＸＯＦ１位置下限
    HSXOF1PX As Double              ' 品ＳＸＯＦ１位置上限
    HSXOF1PW As String * 2          ' 品ＳＸＯＦ１位置検査方法
    HSXOF1LC As Double              ' 品ＳＸＯＦ１長中心
    HSXOF1LN As Double              ' 品ＳＸＯＦ１長下限
    HSXOF1LX As Double              ' 品ＳＸＯＦ１長上限
    HSXOF1DC As Double              ' 品ＳＸＯＦ１直径中心
    HSXOF1DN As Double              ' 品ＳＸＯＦ１直径下限
    HSXOF1DX As Double              ' 品ＳＸＯＦ１直径上限
    HSXDFORM As String * 1          ' 品ＳＸ溝形状
    HSXDPDRC As String * 1          ' 品ＳＸ溝位置方向
    HSXDPACN As Integer             ' 品ＳＸ溝位置角度中心
    HSXDPAMN As Integer             ' 品ＳＸ溝位置角度下限
    HSXDPAMX As Integer             ' 品ＳＸ溝位置角度上限
    HSXDPKWY As String * 2          ' 品ＳＸ溝位置検査方法
    HSXDPDIR As String * 2          ' 品ＳＸ溝位置方位
    HSXDPMIN As Double              ' 品ＳＸ溝位置下限
    HSXDPMAX As Double              ' 品ＳＸ溝位置上限
    HSXDWCEN As Double              ' 品ＳＸ溝巾中心
    HSXDWMIN As Double              ' 品ＳＸ溝巾下限
    HSXDWMAX As Double              ' 品ＳＸ溝巾上限
    HSXDDCEN As Double              ' 品ＳＸ溝深中心
    HSXDDMIN As Double              ' 品ＳＸ溝深下限
    HSXDDMAX As Double              ' 品ＳＸ溝深上限
    HSXDACEN As Double              ' 品ＳＸ溝角度中心
    HSXDAMIN As Double              ' 品ＳＸ溝角度下限
    HSXDAMAX As Double              ' 品ＳＸ溝角度上限
    MCNO As String * 10             ' 結晶操業内製作条件
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 製品仕様SXLﾃﾞｰﾀ２
Public Type typ_TBCME019
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    HMGSTRRNO As String * 9         ' 品管理仕様登録依頼番号
    HMGSTFNO As String * 8          ' 品管理社員Ｎｏ
    HMGSXSNO As String * 6          ' 品管理ＳＸ製品番号
    HMGSXSNE As Integer             ' 品管理ＳＸ製品番号枝番
    HSXTMMAX As Double              ' 品ＳＸ転位密度上限     項目追加，修正対応 2003.05.20 yakimura
    HSXTMSPH As String * 1          ' 品ＳＸ転位密度測定位置＿方
    HSXTMSPT As String * 1          ' 品ＳＸ転位密度測定位置＿点
    HSXTMSPR As String * 1          ' 品ＳＸ転位密度測定位置＿領
    HSXTMKHM As String * 1          ' 品ＳＸ転位密度検査頻度＿枚
    HSXTMKHI As String * 1          ' 品ＳＸ転位密度検査頻度＿位
    HSXTMKHH As String * 1          ' 品ＳＸ転位密度検査頻度＿保
    HSXTMKHS As String * 1          ' 品ＳＸ転位密度検査頻度＿試
    HSXLTMIN As Integer             ' 品ＳＸＬタイム下限
    HSXLTMAX As Integer             ' 品ＳＸＬタイム上限
    HSXLTSPH As String * 1          ' 品ＳＸＬタイム測定位置＿方
    HSXLTSPT As String * 1          ' 品ＳＸＬタイム測定位置＿点
    HSXLTSPI As String * 1          ' 品ＳＸＬタイム測定位置＿位
    HSXLTHWT As String * 1          ' 品ＳＸＬタイム保証方法＿対
    HSXLTHWS As String * 1          ' 品ＳＸＬタイム保証方法＿処
    HSXLTKWY As String * 2          ' 品ＳＸＬタイム検査方法
    HSXLTNSW As String * 2          ' 品ＳＸＬタイム熱処理法
    HSXLTKHM As String * 1          ' 品ＳＸＬタイム検査頻度＿枚
    HSXLTKHI As String * 1          ' 品ＳＸＬタイム検査頻度＿位
    HSXLTKHH As String * 1          ' 品ＳＸＬタイム検査頻度＿保
    HSXLTKHS As String * 1          ' 品ＳＸＬタイム検査頻度＿試
    HSXLTMBP As Double              ' 品ＳＸＬタイム面内分布
    HSXLTMCL As String * 1          ' 品ＳＸＬタイム面内計算
    HSXCNMIN As Double              ' 品ＳＸ炭素濃度下限
    HSXCNMAX As Double              ' 品ＳＸ炭素濃度上限
    HSXCNSPH As String * 1          ' 品ＳＸ炭素濃度測定位置＿方
    HSXCNSPT As String * 1          ' 品ＳＸ炭素濃度測定位置＿点
    HSXCNSPI As String * 1          ' 品ＳＸ炭素濃度測定位置＿位
    HSXCNHWT As String * 1          ' 品ＳＸ炭素濃度保証方法＿対
    HSXCNHWS As String * 1          ' 品ＳＸ炭素濃度保証方法＿処
    HSXCNKWY As String * 2          ' 品ＳＸ炭素濃度検査方法
    HSXCNKHM As String * 1          ' 品ＳＸ炭素濃度検査頻度＿枚
    HSXCNKHI As String * 1          ' 品ＳＸ炭素濃度検査頻度＿位
    HSXCNKHH As String * 1          ' 品ＳＸ炭素濃度検査頻度＿保
    HSXCNKHS As String * 1          ' 品ＳＸ炭素濃度検査頻度＿試
    HSXONMIN As Double              ' 品ＳＸ酸素濃度下限
    HSXONMAX As Double              ' 品ＳＸ酸素濃度上限
    HSXONSPH As String * 1          ' 品ＳＸ酸素濃度測定位置＿方
    HSXONSPT As String * 1          ' 品ＳＸ酸素濃度測定位置＿点
    HSXONSPI As String * 1          ' 品ＳＸ酸素濃度測定位置＿位
    HSXONHWT As String * 1          ' 品ＳＸ酸素濃度保証方法＿対
    HSXONHWS As String * 1          ' 品ＳＸ酸素濃度保証方法＿処
    HSXONKWY As String * 2          ' 品ＳＸ酸素濃度検査方法
    HSXONKHM As String * 1          ' 品ＳＸ酸素濃度検査頻度＿枚
    HSXONKHI As String * 1          ' 品ＳＸ酸素濃度検査頻度＿位
    HSXONKHH As String * 1          ' 品ＳＸ酸素濃度検査頻度＿保
    HSXONKHS As String * 1          ' 品ＳＸ酸素濃度検査頻度＿試
    HSXONMBP As Double              ' 品ＳＸ酸素濃度面内分布
    HSXONMCL As String * 1          ' 品ＳＸ酸素濃度面内計算
    HSXONLTB As Double              ' 品ＳＸ酸素濃度ＬＴ分布
    HSXONLTC As String * 1          ' 品ＳＸ酸素濃度ＬＴ計算
    HSXONSDV As Double              ' 品ＳＸ酸素濃度標準偏差
    HSXONAMN As Double              ' 品ＳＸ酸素濃度平均下限
    HSXONAMX As Double              ' 品ＳＸ酸素濃度平均上限
    HSXOS1MN As Double              ' 品ＳＸ酸素析出１下限
    HSXOS1MX As Double              ' 品ＳＸ酸素析出１上限
    HSXOS1NS As String * 2          ' 品ＳＸ酸素析出１熱処理法
    HSXOS1SH As String * 1          ' 品ＳＸ酸素析出１測定位置＿方
    HSXOS1ST As String * 1          ' 品ＳＸ酸素析出１測定位置＿点
    HSXOS1SI As String * 1          ' 品ＳＸ酸素析出１測定位置＿位
    HSXOS1HT As String * 1          ' 品ＳＸ酸素析出１保証方法＿対
    HSXOS1HS As String * 1          ' 品ＳＸ酸素析出１保証方法＿処
    HSXOS1HM As String * 1          ' 品ＳＸ酸素析出１検査頻度＿枚
    HSXOS1KI As String * 1          ' 品ＳＸ酸素析出１検査頻度＿位
    HSXOS1KH As String * 1          ' 品ＳＸ酸素析出１検査頻度＿保
    HSXOS1KS As String * 1          ' 品ＳＸ酸素析出１検査頻度＿試
    HSXOS2MN As Double              ' 品ＳＸ酸素析出２下限
    HSXOS2MX As Double              ' 品ＳＸ酸素析出２上限
    HSXOS2NS As String * 2          ' 品ＳＸ酸素析出２熱処理法
    HSXOS2SH As String * 1          ' 品ＳＸ酸素析出２測定位置＿方
    HSXOS2ST As String * 1          ' 品ＳＸ酸素析出２測定位置＿点
    HSXOS2SI As String * 1          ' 品ＳＸ酸素析出２測定位置＿位
    HSXOS2HT As String * 1          ' 品ＳＸ酸素析出２保証方法＿対
    HSXOS2HS As String * 1          ' 品ＳＸ酸素析出２保証方法＿処
    HSXOS2KM As String * 1          ' 品ＳＸ酸素析出２検査頻度＿枚
    HSXOS2KN As String * 1          ' 品ＳＸ酸素析出２検査頻度＿位
    HSXOS2KH As String * 1          ' 品ＳＸ酸素析出２検査頻度＿保
    HSXOS2KU As String * 1          ' 品ＳＸ酸素析出２検査頻度＿試
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
' 追加 2003/09.11 SystemBrain Start
    HSXTMMAXN As Double             ' 品ＳＸ転位密度上限
' 追加 2003/09.11 SystemBrain End
End Type


' 製品仕様SXLﾃﾞｰﾀ３
Public Type typ_TBCME020
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    HMGSTRRNO As String * 9         ' 品管理仕様登録依頼番号
    HMGSTFNO As String * 8          ' 品管理社員Ｎｏ
    HMGSXSNO As String * 6          ' 品管理ＳＸ製品番号
    HMGSXSNE As Integer             ' 品管理ＳＸ製品番号枝番
    HSXDENKU As String * 1          ' 品ＳＸＤｅｎ検査有無
    HSXDENMX As Integer             ' 品ＳＸＤｅｎ上限
    HSXDENMN As Integer             ' 品ＳＸＤｅｎ下限
    HSXDENHT As String * 1          ' 品ＳＸＤｅｎ保証方法＿対
    HSXDENHS As String * 1          ' 品ＳＸＤｅｎ保証方法＿処
    HSXDVDKU As String * 1          ' 品ＳＸＤＶＤ２検査有無
    HSXDVDMX As Integer             ' 品ＳＸＤＶＤ２上限
    HSXDVDMN As Integer             ' 品ＳＸＤＶＤ２下限
    HSXDVDHT As String * 1          ' 品ＳＸＤＶＤ２保証方法＿対
    HSXDVDHS As String * 1          ' 品ＳＸＤＶＤ２保証方法＿処
    HSXLDLKU As String * 1          ' 品ＳＸＬ／ＤＬ検査有無
    HSXLDLMX As Integer             ' 品ＳＸＬ／ＤＬ上限
    HSXLDLMN As Integer             ' 品ＳＸＬ／ＤＬ下限
    HSXLDLHT As String * 1          ' 品ＳＸＬ／ＤＬ保証方法＿対
    HSXLDLHS As String * 1          ' 品ＳＸＬ／ＤＬ保証方法＿処
    HSXGDSZY As String * 1          ' 品ＳＸＧＤ測定条件
    HSXGDSPH As String * 1          ' 品ＳＸＧＤ測定位置＿方
    HSXGDSPT As String * 1          ' 品ＳＸＧＤ測定位置＿点
    HSXGDSPR As String * 1          ' 品ＳＸＧＤ測定位置＿領
    HSXGDZAR As Integer             ' 品ＳＸＧＤ除外領域
    HSXGDKHM As String * 1          ' 品ＳＸＧＤ検査頻度＿枚
    HSXGDKHI As String * 1          ' 品ＳＸＧＤ検査頻度＿位
    HSXGDKHH As String * 1          ' 品ＳＸＧＤ検査頻度＿保
    HSXGDKHS As String * 1          ' 品ＳＸＧＤ検査頻度＿試
    HSXDSOKE As String * 1          ' 品ＳＸＤＳＯＤ検査
    HSXDSOMX As Long                ' 品ＳＸＤＳＯＤ上限
    HSXDSOMN As Long                ' 品ＳＸＤＳＯＤ下限
    HSXDSOAX As Integer             ' 品ＳＸＤＳＯＤ領域上限
    HSXDSOAN As Integer             ' 品ＳＸＤＳＯＤ領域下限
    HSXDSOHT As String * 1          ' 品ＳＸＤＳＯＤ保証方法＿対
    HSXDSOHS As String * 1          ' 品ＳＸＤＳＯＤ保証方法＿処
    HSXDSOKM As String * 1          ' 品ＳＸＤＳＯＤ検査頻度＿枚
    HSXDSOKI As String * 1          ' 品ＳＸＤＳＯＤ検査頻度＿位
    HSXDSOKH As String * 1          ' 品ＳＸＤＳＯＤ検査頻度＿保
    HSXDSOKS As String * 1          ' 品ＳＸＤＳＯＤ検査頻度＿試
    HSXLIFTW As String * 2          ' 品ＳＸ引上方法
    HSXSDSLP As String * 1          ' 品ＳＸシード傾
    HSXGKKNO As String * 6          ' 品ＳＸ外観規格Ｎｏ
    HSXCDOP As String * 1           ' 品ＳＸ結晶ドープ
    HSXCDOPN As Double              ' 品ＳＸ結晶ドープ濃度
    HSXCDPNI As String * 2          ' 品ＳＸ結晶ドープ濃度指数
    HSXGSFIN As String * 1          ' 品ＳＸ外周仕上げ
    HSXCLMIN As Integer             ' 品ＳＸ結晶長下限
    HSXCLMAX As Integer             ' 品ＳＸ結晶長上限
    HSXCLPMN As Integer             ' 品ＳＸ結晶長許容下限
    HSXCLPR As Double               ' 品ＳＸ結晶長許容比率
    HSXWFWAR As String * 1          ' 品ＳＸＷＦＷａｒｐランク
    HSXOF1AX As Double              ' 品ＳＸＯＳＦ１平均上限
    HSXOF1MX As Double              ' 品ＳＸＯＳＦ１上限
    HSXOF1SH As String * 1          ' 品ＳＸＯＳＦ１測定位置＿方
    HSXOF1ST As String * 1          ' 品ＳＸＯＳＦ１測定位置＿点
    HSXOF1SR As String * 1          ' 品ＳＸＯＳＦ１測定位置＿領
    HSXOF1HT As String * 1          ' 品ＳＸＯＳＦ１保証方法＿対
    HSXOF1HS As String * 1          ' 品ＳＸＯＳＦ１保証方法＿処
    HSXOF1SZ As String * 1          ' 品ＳＸＯＳＦ１測定条件
    HSXOF1KM As String * 1          ' 品ＳＸＯＳＦ１検査頻度＿枚
    HSXOF1KI As String * 1          ' 品ＳＸＯＳＦ１検査頻度＿位
    HSXOF1KH As String * 1          ' 品ＳＸＯＳＦ１検査頻度＿保
    HSXOF1KS As String * 1          ' 品ＳＸＯＳＦ１検査頻度＿試
    HSXOF1NS As String * 2          ' 品ＳＸＯＳＦ１熱処理法
    HSXOF1ET As Integer             ' 品ＳＸＯＳＦ１選択ＥＴ代
    HSXOF2AX As Double              ' 品ＳＸＯＳＦ２平均上限
    HSXOF2MX As Double              ' 品ＳＸＯＳＦ２上限
    HSXOF2SH As String * 1          ' 品ＳＸＯＳＦ２測定位置＿方
    HSXOF2ST As String * 1          ' 品ＳＸＯＳＦ２測定位置＿点
    HSXOF2SR As String * 1          ' 品ＳＸＯＳＦ２測定位置＿領
    HSXOF2HT As String * 1          ' 品ＳＸＯＳＦ２保証方法＿対
    HSXOF2HS As String * 1          ' 品ＳＸＯＳＦ２保証方法＿処
    HSXOF2SZ As String * 1          ' 品ＳＸＯＳＦ２測定条件
    HSXOF2KM As String * 1          ' 品ＳＸＯＳＦ２検査頻度＿枚
    HSXOF2KI As String * 1          ' 品ＳＸＯＳＦ２検査頻度＿位
    HSXOF2KH As String * 1          ' 品ＳＸＯＳＦ２検査頻度＿保
    HSXOF2KS As String * 1          ' 品ＳＸＯＳＦ２検査頻度＿試
    HSXOF2NS As String * 2          ' 品ＳＸＯＳＦ２熱処理法
    HSXOF2ET As Integer             ' 品ＳＸＯＳＦ２選択ＥＴ代
    HSXOF3AX As Double              ' 品ＳＸＯＳＦ３平均上限
    HSXOF3MX As Double              ' 品ＳＸＯＳＦ３上限
    HSXOF3SH As String * 1          ' 品ＳＸＯＳＦ３測定位置＿方
    HSXOF3ST As String * 1          ' 品ＳＸＯＳＦ３測定位置＿点
    HSXOF3SR As String * 1          ' 品ＳＸＯＳＦ３測定位置＿領
    HSXOF3HT As String * 1          ' 品ＳＸＯＳＦ３保証方法＿対
    HSXOF3HS As String * 1          ' 品ＳＸＯＳＦ３保証方法＿処
    HSXOF3SZ As String * 1          ' 品ＳＸＯＳＦ３測定条件
    HSXOF3KM As String * 1          ' 品ＳＸＯＳＦ３検査頻度＿枚
    HSXOF3KI As String * 1          ' 品ＳＸＯＳＦ３検査頻度＿位
    HSXOF3KH As String * 1          ' 品ＳＸＯＳＦ３検査頻度＿保
    HSXOF3KS As String * 1          ' 品ＳＸＯＳＦ３検査頻度＿試
    HSXOF3NS As String * 2          ' 品ＳＸＯＳＦ３熱処理法
    HSXOF3ET As Integer             ' 品ＳＸＯＳＦ３選択ＥＴ代
    HSXOF4AX As Double              ' 品ＳＸＯＳＦ４平均上限
    HSXOF4MX As Double              ' 品ＳＸＯＳＦ４上限
    HSXOF4SH As String * 1          ' 品ＳＸＯＳＦ４測定位置＿方
    HSXOF4ST As String * 1          ' 品ＳＸＯＳＦ４測定位置＿点
    HSXOF4SR As String * 1          ' 品ＳＸＯＳＦ４測定位置＿領
    HSXOF4HT As String * 1          ' 品ＳＸＯＳＦ４保証方法＿対
    HSXOF4HS As String * 1          ' 品ＳＸＯＳＦ４保証方法＿処
    HSXOF4SZ As String * 1          ' 品ＳＸＯＳＦ４測定条件
    HSXOF4KM As String * 1          ' 品ＳＸＯＳＦ４検査頻度＿枚
    HSXOF4KI As String * 1          ' 品ＳＸＯＳＦ４検査頻度＿位
    HSXOF4KH As String * 1          ' 品ＳＸＯＳＦ４検査頻度＿保
    HSXOF4KS As String * 1          ' 品ＳＸＯＳＦ４検査頻度＿試
    HSXOF4NS As String * 2          ' 品ＳＸＯＳＦ４熱処理法
    HSXOF4ET As Integer             ' 品ＳＸＯＳＦ４選択ＥＴ代
    HSXBM1AN As Double              ' 品ＳＸＢＭＤ１平均下限
    HSXBM1AX As Double              ' 品ＳＸＢＭＤ１平均上限
    HSXBM1SH As String * 1          ' 品ＳＸＢＭＤ１測定位置＿方
    HSXBM1ST As String * 1          ' 品ＳＸＢＭＤ１測定位置＿点
    HSXBM1SR As String * 1          ' 品ＳＸＢＭＤ１測定位置＿領
    HSXBM1HT As String * 1          ' 品ＳＸＢＭＤ１保証方法＿対
    HSXBM1HS As String * 1          ' 品ＳＸＢＭＤ１保証方法＿処
    HSXBM1SZ As String * 1          ' 品ＳＸＢＭＤ１測定条件
    HSXBM1KM As String * 1          ' 品ＳＸＢＭＤ１検査頻度＿枚
    HSXBM1KI As String * 1          ' 品ＳＸＢＭＤ１検査頻度＿位
    HSXBM1KH As String * 1          ' 品ＳＸＢＭＤ１検査頻度＿保
    HSXBM1KS As String * 1          ' 品ＳＸＢＭＤ１検査頻度＿試
    HSXBM1NS As String * 2          ' 品ＳＸＢＭＤ１熱処理法
    HSXBM1ET As Integer             ' 品ＳＸＢＭＤ１選択ＥＴ代
    HSXBM2AN As Double              ' 品ＳＸＢＭＤ２平均下限
    HSXBM2AX As Double              ' 品ＳＸＢＭＤ２平均上限
    HSXBM2SH As String * 1          ' 品ＳＸＢＭＤ２測定位置＿方
    HSXBM2ST As String * 1          ' 品ＳＸＢＭＤ２測定位置＿点
    HSXBM2SR As String * 1          ' 品ＳＸＢＭＤ２測定位置＿領
    HSXBM2HT As String * 1          ' 品ＳＸＢＭＤ２保証方法＿対
    HSXBM2HS As String * 1          ' 品ＳＸＢＭＤ２保証方法＿処
    HSXBM2SZ As String * 1          ' 品ＳＸＢＭＤ２測定条件
    HSXBM2KM As String * 1          ' 品ＳＸＢＭＤ２検査頻度＿枚
    HSXBM2KI As String * 1          ' 品ＳＸＢＭＤ２検査頻度＿位
    HSXBM2KH As String * 1          ' 品ＳＸＢＭＤ２検査頻度＿保
    HSXBM2KS As String * 1          ' 品ＳＸＢＭＤ２検査頻度＿試
    HSXBM2NS As String * 2          ' 品ＳＸＢＭＤ２熱処理法
    HSXBM2ET As Integer             ' 品ＳＸＢＭＤ２選択ＥＴ代
    HSXBM3AN As Double              ' 品ＳＸＢＭＤ３平均下限
    HSXBM3AX As Double              ' 品ＳＸＢＭＤ３平均上限
    HSXBM3SH As String * 1          ' 品ＳＸＢＭＤ３測定位置＿方
    HSXBM3ST As String * 1          ' 品ＳＸＢＭＤ３測定位置＿点
    HSXBM3SR As String * 1          ' 品ＳＸＢＭＤ３測定位置＿領
    HSXBM3HT As String * 1          ' 品ＳＸＢＭＤ３保証方法＿対
    HSXBM3HS As String * 1          ' 品ＳＸＢＭＤ３保証方法＿処
    HSXBM3SZ As String * 1          ' 品ＳＸＢＭＤ３測定条件
    HSXBM3KM As String * 1          ' 品ＳＸＢＭＤ３検査頻度＿枚
    HSXBM3KI As String * 1          ' 品ＳＸＢＭＤ３検査頻度＿位
    HSXBM3KH As String * 1          ' 品ＳＸＢＭＤ３検査頻度＿保
    HSXBM3KS As String * 1          ' 品ＳＸＢＭＤ３検査頻度＿試
    HSXBM3NS As String * 2          ' 品ＳＸＢＭＤ３熱処理法
    HSXBM3ET As Integer             ' 品ＳＸＢＭＤ３選択ＥＴ代
    HSXNOTE As String               ' 品ＳＸ特記
    HSXRS1N As String               ' 品ＳＸ予備１＿内
    HSXRS1Y As String               ' 品ＳＸ予備１＿用
    HSXRS2N As String               ' 品ＳＸ予備２＿内
    HSXRS2Y As String               ' 品ＳＸ予備２＿用
    HSXRS3N As String               ' 品ＳＸ予備３＿内
    HSXRS3Y As String               ' 品ＳＸ予備３＿用
    HSXRS4N As String               ' 品ＳＸ予備４＿内
    HSXRS4Y As String               ' 品ＳＸ予備４＿用
    HSXRS5N As String               ' 品ＳＸ予備５＿内
    HSXRS5Y As String               ' 品ＳＸ予備５＿用
    HSXRS6N As String               ' 品ＳＸ予備６＿内
    HSXRS6Y As String               ' 品ＳＸ予備６＿用
    HSXRS7N As String               ' 品ＳＸ予備７＿内
    HSXRS7Y As String               ' 品ＳＸ予備７＿用
    HSXRS8N As String               ' 品ＳＸ予備８＿内
    HSXRS8Y As String               ' 品ＳＸ予備８＿用
    HSXRS9N As String               ' 品ＳＸ予備９＿内
    HSXRS9Y As String               ' 品ＳＸ予備９＿用
    HSXRS10N As String              ' 品ＳＸ予備１０＿内
    HSXRS10Y As String              ' 品ＳＸ予備１０＿用
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
' 追加 2003/09.11 SystemBrain Start
    HSXDVDMXN As Integer            ' 品ＳＸＤＶＤ２上限
    HSXDVDMNN As Integer            ' 品ＳＸＤＶＤ２下限
    HSXDSONS As String * 2          ' 品ＳＸＤＳＯＤ熱処理法
    HSXCDOPMX As Double             ' 品ＳＸ結晶ドープ濃度下限
    HSXCDOPMN As Double             ' 品ＳＸ結晶ドープ濃度上限
' 追加 2003/09.11 SystemBrain End
' OSF，BMD項目追加対応  2002.04.02 yakimura
    HSXOSF1PTK As String * 1        ' 品ＳＸＯＳＦ１パタン区分
    HSXOSF2PTK As String * 1        ' 品ＳＸＯＳＦ２パタン区分
    HSXOSF3PTK As String * 1        ' 品ＳＸＯＳＦ３パタン区分
    HSXOSF4PTK As String * 1        ' 品ＳＸＯＳＦ４パタン区分
    HSXBMD1MBP As Double            ' 品ＳＸＢＭＤ１面内分布
    HSXBMD2MBP As Double            ' 品ＳＸＢＭＤ２面内分布
    HSXBMD3MBP As Double            ' 品ＳＸＢＭＤ３面内分布
' OSF，BMD項目追加対応  2002.04.02 yakimura
' 追加 2003/09.11 SystemBrain Start
    HSXBMD1MCL As String * 2        ' 品SXBMD1面内計算
    HSXBMD2MCL As String * 2        ' 品SXBMD2面内計算
    HSXBMD3MCL As String * 2        ' 品SXBMD3面内計算
' 追加 2003/09.11 SystemBrain End

'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    HSXGDPTK As String * 1          ' 品ＳＸＧＤパタン区分
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End

    'Add Start 2011/01/26 SMPK Miyata
    HSXCPK      As String * 1       '品ＳＸＣパターン区分
    HSXCSZ      As String * 1       '品ＳＸＣ測定条件
    HSXCHT      As String * 1       '品ＳＸＣ保証方法＿対
    HSXCHS      As String * 1       '品ＳＸＣ保証方法＿処
    HSXCJPK     As String * 1       '品ＳＸＣＪパターン区分
    HSXCJNS     As String * 2       '品ＳＸＣＪ熱処理法
    HSXCJHT     As String * 1       '品ＳＸＣＪ保証方法＿対
    HSXCJHS     As String * 1       '品ＳＸＣＪ保証方法＿処
    HSXCJLTPK   As String * 1       '品ＳＸＣＪＬＴパターン区分
    HSXCJLTNS   As String * 2       '品ＳＸＣＪＬＴ熱処理法
    HSXCJLTHT   As String * 1       '品ＳＸＣＪＬＴ保証方法＿対
    HSXCJLTHS   As String * 1       '品ＳＸＣＪＬＴ保証方法＿処
    HSXCJ2PK    As String * 1       '品ＳＸＣＪ２パターン区分
    HSXCJ2NS    As String * 2       '品ＳＸＣＪ２熱処理法
    HSXCJ2HT    As String * 1       '品ＳＸＣＪ２保証方法＿対
    HSXCJ2HS    As String * 1       '品ＳＸＣＪ２保証方法＿処
    'Add End   2011/01/26 SMPK Miyata
    
    'Add Start 2011/02/17 Y.Hitomi
    HSXCOSF3NS As String * 2        '品ＳＸＣＯＳＦ３熱処理法
    'Add End   2011/02/17 Y.Hitomi

End Type


' 製品仕様WFﾃﾞｰﾀ１
Public Type typ_TBCME021
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    HMGSTRRNO As String * 9         ' 品管理仕様登録依頼番号
    HMGSTFNO As String * 8          ' 品管理社員Ｎｏ
    HMGWFSNO As String * 6          ' 品管理ＷＦ製品番号
    HMGWFSNE As Integer             ' 品管理ＷＦ製品番号枝番
    CONFLAG As String * 1           ' 確認フラグ
    REINFLAG As String * 1          ' 再付与フラグ
    HWFTRWKB As String * 1          ' 品ＷＦ統合可否区分
    HWFFACES As String * 2          ' 品ＷＦ表面仕上げ
    HWFBACKS As String * 2          ' 品ＷＦ裏仕上げ
    HWFBDSWY As String * 2          ' 品ＷＦＢＤ処理方法
'    HSXTYPE As String * 1           ' 品ＷＦタイプ
    HWFTYPE As String * 1           ' 品ＷＦタイプ　05/03/01 ooba
    HWFTYPKW As String * 1          ' 品ＷＦタイプ検査方法
    HWFDOP As String * 1            ' 品ＷＦドーパント
    HWFFKBWK As String * 1          ' 品ＷＦ表面区分方法＿区
    HWFFKBWS As String * 1          ' 品ＷＦ表面区分方法＿指
    HWFRMIN As Double               ' 品ＷＦ比抵抗下限
    HWFRMAX As Double               ' 品ＷＦ比抵抗上限
    HWFRSPOH As String * 1          ' 品ＷＦ比抵抗測定位置＿方
    HWFRSPOT As String * 1          ' 品ＷＦ比抵抗測定位置＿点
    HWFRSPOI As String * 1          ' 品ＷＦ比抵抗測定位置＿位
    HWFRHWYT As String * 1          ' 品ＷＦ比抵抗保証方法＿対
    HWFRHWYS As String * 1          ' 品ＷＦ比抵抗保証方法＿処
    HWFRKWAY As String * 2          ' 品ＷＦ比抵抗検査方法
    HWFRKHNM As String * 1          ' 品ＷＦ比抵抗検査頻度＿枚
    HWFRKHNN As String * 1          ' 品ＷＦ比抵抗検査頻度＿抜
    HWFRKHNH As String * 1          ' 品ＷＦ比抵抗検査頻度＿保
    HWFRKHNU As String * 1          ' 品ＷＦ比抵抗検査頻度＿ウ
    HWFRSDEV As Double              ' 品ＷＦ比抵抗標準偏差
    HWFRAMIN As Double              ' 品ＷＦ比抵抗平均下限
    HWFRAMAX As Double              ' 品ＷＦ比抵抗平均上限
    HWFRMBNP As Double              ' 品ＷＦ比抵抗面内分布
    HWFRMCAL As String * 1          ' 品ＷＦ比抵抗面内計算
    HWFRMBP2 As Double              ' 品ＷＦ比抵抗面内分布２
    HWFRMCL2 As String * 1          ' 品ＷＦ比抵抗面内計算２
    HWFRKBSH As String * 1          ' 品ＷＦ比抵抗振区分測定位置＿方
    HWFRKBST As String * 1          ' 品ＷＦ比抵抗振区分測定位置＿点
    HWFRKBSI As String * 1          ' 品ＷＦ比抵抗振区分測定位置＿位
    HWFRKBHT As String * 1          ' 品ＷＦ比抵抗振区分保証方法＿対
    HWFRKBHS As String * 1          ' 品ＷＦ比抵抗振区分保証方法＿処
    HWFSTMAX As Double              ' 品ＷＦストリエ上限
    HWFSTSPH As String * 1          ' 品ＷＦストリエ測定位置＿方
    HWFSTSPT As String * 1          ' 品ＷＦストリエ測定位置＿点
    HWFSTSPI As String * 1          ' 品ＷＦストリエ測定位置＿位
    HWFSTHWT As String * 1          ' 品ＷＦストリエ保証方法＿対
    HWFSTHWS As String * 1          ' 品ＷＦストリエ保証方法＿処
    HWFSTKWY As String * 2          ' 品ＷＦストリエ検査方法
    HWFSTKHM As String * 1          ' 品ＷＦストリエ検査頻度＿枚
    HWFSTKHN As String * 1          ' 品ＷＦストリエ検査頻度＿抜
    HWFSTKHH As String * 1          ' 品ＷＦストリエ検査頻度＿保
    HWFSTKHU As String * 1          ' 品ＷＦストリエ検査頻度＿ウ
    HWFACEN As Double               ' 品ＷＦ厚中心
    HWFAMIN As Double               ' 品ＷＦ厚下限
    HWFAMAX As Double               ' 品ＷＦ厚上限
    HWFASPOH As String * 1          ' 品ＷＦ厚測定位置＿方
    HWFASPOT As String * 1          ' 品ＷＦ厚測定位置＿点
    HWFASPOI As String * 1          ' 品ＷＦ厚測定位置＿位
    HWFAHWYT As String * 1          ' 品ＷＦ厚保証方法＿対
    HWFAHWYS As String * 1          ' 品ＷＦ厚保証方法＿処
    HWFAKWAY As String * 1          ' 品ＷＦ厚検査方法
    HWFAKHNM As String * 1          ' 品ＷＦ厚検査頻度＿枚
    HWFAKHNN As String * 1          ' 品ＷＦ厚検査頻度＿抜
    HWFAKHNH As String * 1          ' 品ＷＦ厚検査頻度＿保
    HWFAKHNU As String * 1          ' 品ＷＦ厚検査頻度＿ウ
    HWFASDEV As Double              ' 品ＷＦ厚標準偏差
    HWFAAMIN As Double              ' 品ＷＦ厚平均下限
    HWFAAMAX As Double              ' 品ＷＦ厚平均上限
    HWFAMBNP As Double              ' 品ＷＦ厚面内分布
    HWFAMCAL As String * 1          ' 品ＷＦ厚面内計算
    HWFALTBP As Double              ' 品ＷＦ厚ＬＴ分布
    HWFALTCL As String * 1          ' 品ＷＦ厚ＬＴ計算
    HWFALTRA As Double              ' 品ＷＦ厚ＬＴ範囲
    HWFAMRAN As Double              ' 品ＷＦ厚面内範囲
    HWFDIVS As Integer              ' 品ＷＦ分割数
    HWFAKBSH As String * 1          ' 品ＷＦ厚振区分測定位置＿方
    HWFAKBST As String * 1          ' 品ＷＦ厚振区分測定位置＿点
    HWFAKBSI As String * 1          ' 品ＷＦ厚振区分測定位置＿位
    HWFAKBHT As String * 1          ' 品ＷＦ厚振区分保証方法＿対
    HWFAKBHS As String * 1          ' 品ＷＦ厚振区分保証方法＿処
    HWFWFORM As String * 1          ' 品ＷＦウェーハ形状
    HWFD1CEN As Double              ' 品ＷＦ直径１中心
    HWFD1MIN As Double              ' 品ＷＦ直径１下限
    HWFD1MAX As Double              ' 品ＷＦ直径１上限
    HWFD2CEN As Double              ' 品ＷＦ直径２中心
    HWFD2MIN As Double              ' 品ＷＦ直径２下限
    HWFD2MAX As Double              ' 品ＷＦ直径２上限
    HWFDKHNM As String * 1          ' 品ＷＦ直径検査頻度＿枚
    HWFDKHNN As String * 1          ' 品ＷＦ直径検査頻度＿抜
    HWFDKHNH As String * 1          ' 品ＷＦ直径検査頻度＿保
    HWFDKHNU As String * 1          ' 品ＷＦ直径検査頻度＿ウ
    HWFLPMNP As Integer             ' 品ＷＦＬＰ厚最小加工代
    HWFSGMNP As Integer             ' 品ＷＦＳＧ厚最小加工代
    HWFETMNP As Integer             ' 品ＷＦＥＴ厚最小加工代
    HWFMPMNP As Integer             ' 品ＷＦＭＰ厚最小加工代
    HWFLPKS1 As String * 1          ' 品ＷＦＬＰ研磨材種１
    HWFLPKS2 As String * 1          ' 品ＷＦＬＰ研磨材種２
    HWFLPKZ1 As String * 1          ' 品ＷＦＬＰ研磨材粒度種１
    HWFLPKZ2 As String * 1          ' 品ＷＦＬＰ研磨材粒度種２
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 製品仕様WFﾃﾞｰﾀ２
Public Type typ_TBCME022
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    HMGSTRRNO As String * 9         ' 品管理仕様登録依頼番号
    HMGSTFNO As String * 8          ' 品管理社員Ｎｏ
    HMGWFSNO As String * 6          ' 品管理ＷＦ製品番号
    HMGWFSNE As Integer             ' 品管理ＷＦ製品番号枝番
    HWFCDIR As String * 1           ' 品ＷＦ結晶面方位
    HWFCSCEN As Double              ' 品ＷＦ結晶面傾中心
    HWFCSMIN As Double              ' 品ＷＦ結晶面傾下限
    HWFCSMAX As Double              ' 品ＷＦ結晶面傾上限
    HWFCSDIS As String * 1          ' 品ＷＦ結晶面傾方位指定
    HWFCSDIR As String * 2          ' 品ＷＦ結晶面傾方位
    HWFCKWAY As String * 2          ' 品ＷＦ結晶面検査方法
    HWFCKHNM As String * 1          ' 品ＷＦ結晶面検査頻度＿枚
    HWFCKHNN As String * 1          ' 品ＷＦ結晶面検査頻度＿抜
    HWFCKHNH As String * 1          ' 品ＷＦ結晶面検査頻度＿保
    HWFCKHNU As String * 1          ' 品ＷＦ結晶面検査頻度＿ウ
    HWFCTDIR As String * 2          ' 品ＷＦ結晶面傾縦方位
    HWFCTCEN As Double              ' 品ＷＦ結晶面傾縦中心
    HWFCTMIN As Double              ' 品ＷＦ結晶面傾縦下限
    HWFCTMAX As Double              ' 品ＷＦ結晶面傾縦上限
    HWFCYDIR As String * 2          ' 品ＷＦ結晶面傾横方位
    HWFCYCEN As Double              ' 品ＷＦ結晶面傾横中心
    HWFCYMIN As Double              ' 品ＷＦ結晶面傾横下限
    HWFCYMAX As Double              ' 品ＷＦ結晶面傾横上限
    HWFKPTNN As String * 3          ' 品ＷＦ光像パタン名
    HWFOFPKM As String * 1          ' 品ＷＦＯＦ位置検査頻度＿枚
    HWFOFPKN As String * 1          ' 品ＷＦＯＦ位置検査頻度＿抜
    HWFOFPKH As String * 1          ' 品ＷＦＯＦ位置検査頻度＿保
    HWFOFPKU As String * 1          ' 品ＷＦＯＦ位置検査頻度＿ウ
    HWFOFLKM As String * 1          ' 品ＷＦＯＦ長検査頻度＿枚
    HWFOFLKN As String * 1          ' 品ＷＦＯＦ長検査頻度＿抜
    HWFOFLKH As String * 1          ' 品ＷＦＯＦ長検査頻度＿保
    HWFOFLKU As String * 1          ' 品ＷＦＯＦ長検査頻度＿ウ
    HWFOF1PD As String * 2          ' 品ＷＦＯＦ１位置方位
    HWFOF1PN As Double              ' 品ＷＦＯＦ１位置下限
    HWFOF1PX As Double              ' 品ＷＦＯＦ１位置上限
    HWFOF1PW As String * 2          ' 品ＷＦＯＦ１位置検査方法
    HWFOF1LC As Double              ' 品ＷＦＯＦ１長中心
    HWFOF1LN As Double              ' 品ＷＦＯＦ１長下限
    HWFOF1LX As Double              ' 品ＷＦＯＦ１長上限
    HWFOF1RF As String * 1          ' 品ＷＦＯＦ１両端Ｒ形状
    HWFOFRRC As Double              ' 品ＷＦＯＦ両端Ｒ右中心
    HWFOFRRN As Double              ' 品ＷＦＯＦ両端Ｒ右下限
    HWFOFRRX As Double              ' 品ＷＦＯＦ両端Ｒ右上限
    HWFOFRLC As Double              ' 品ＷＦＯＦ両端Ｒ左中心
    HWFOFRLN As Double              ' 品ＷＦＯＦ両端Ｒ左下限
    HWFOFRLX As Double              ' 品ＷＦＯＦ両端Ｒ左上限
    HWFOF1DC As Double              ' 品ＷＦＯＦ１直径中心
    HWFOF1DN As Double              ' 品ＷＦＯＦ１直径下限
    HWFOF1DX As Double              ' 品ＷＦＯＦ１直径上限
    HWFZFORM As String * 1          ' 品ＷＦ材料形状
    HWFD3CEN As Double              ' 品ＷＦ直径３中心
    HWFD3MIN As Double              ' 品ＷＦ直径３下限
    HWFD3MAX As Double              ' 品ＷＦ直径３上限
    HWFDFKJ As String * 1           ' 品ＷＦ溝形状
    HWFDFKHM As String * 1          ' 品ＷＦ溝形状検査頻度＿枚
    HWFDFKHN As String * 1          ' 品ＷＦ溝形状検査頻度＿抜
    HWFDFKHH As String * 1          ' 品ＷＦ溝形状検査頻度＿保
    HWFDFKHU As String * 1          ' 品ＷＦ溝形状検査頻度＿ウ
    HWFDPDRC As String * 1          ' 品ＷＦ溝位置方向
    HWFDPACN As Integer             ' 品ＷＦ溝位置角度中心
    HWFDPAMN As Integer             ' 品ＷＦ溝位置角度下限
    HWFDPAMX As Integer             ' 品ＷＦ溝位置角度上限
    HWFDPDIR As String * 2          ' 品ＷＦ溝位置方位
    HWFDPMIN As Double              ' 品ＷＦ溝位置下限
    HWFDPMAX As Double              ' 品ＷＦ溝位置上限
    HWFDPKWY As String * 2          ' 品ＷＦ溝位置検査方法
    HWFDPKHM As String * 1          ' 品ＷＦ溝位置検査頻度＿枚
    HWFDPKHB As String * 1          ' 品ＷＦ溝位置検査頻度＿抜
    HWFDPKHH As String * 1          ' 品ＷＦ溝位置検査頻度＿保
    HWFDPKHU As String * 1          ' 品ＷＦ溝位置検査頻度＿ウ
    HWFDACEN As Double              ' 品ＷＦ溝角度中心
    HWFDAMIN As Double              ' 品ＷＦ溝角度下限
    HWFDAMAX As Double              ' 品ＷＦ溝角度上限
    HWFDWCEN As Double              ' 品ＷＦ溝巾中心
    HWFDWMIN As Double              ' 品ＷＦ溝巾下限
    HWFDWMAX As Double              ' 品ＷＦ溝巾上限
    HWFDDCEN As Double              ' 品ＷＦ溝深中心
    HWFDDMIN As Double              ' 品ＷＦ溝深下限
    HWFDDMAX As Double              ' 品ＷＦ溝深上限
    HWFDBRCN As Double              ' 品ＷＦ溝底Ｒ中心
    HWFDBRMN As Double              ' 品ＷＦ溝底Ｒ下限
    HWFDBRMX As Double              ' 品ＷＦ溝底Ｒ上限
    HWFDRRCN As Double              ' 品ＷＦ溝右Ｒ中心
    HWFDRRMN As Double              ' 品ＷＦ溝右Ｒ下限
    HWFDRRMX As Double              ' 品ＷＦ溝右Ｒ上限
    HWFDLRCN As Double              ' 品ＷＦ溝左Ｒ中心
    HWFDLRMN As Double              ' 品ＷＦ溝左Ｒ下限
    HWFDLRMX As Double              ' 品ＷＦ溝左Ｒ上限
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 製品仕様WFﾃﾞｰﾀ３
Public Type typ_TBCME023
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    HMGSTRRNO As String * 9         ' 品管理仕様登録依頼番号
    HMGSTFNO As String * 8          ' 品管理社員Ｎｏ
    HMGWFSNO As String * 6          ' 品管理ＷＦ製品番号
    HMGWFSNE As Integer             ' 品管理ＷＦ製品番号枝番
    HWFMFORM As String * 1          ' 品ＷＦ面取形状
    KWFMM As String * 1             ' 品ＷＦ面取面粗
    HWFMFKHM As String * 1          ' 品ＷＦ面取形状検査頻度＿枚
    HWFMFKHN As String * 1          ' 品ＷＦ面取形状検査頻度＿抜
    HWFMFKHH As String * 1          ' 品ＷＦ面取形状検査頻度＿保
    HWFMFKHU As String * 1          ' 品ＷＦ面取形状検査頻度＿ウ
    HWFMACEN As Double              ' 品ＷＦ面取角度中心
    HWFMAMIN As Double              ' 品ＷＦ面取角度下限
    HWFMAMAX As Double              ' 品ＷＦ面取角度上限
    HWFMWFCN As Integer             ' 品ＷＦ面取巾表中心
    HWFMWFMN As Integer             ' 品ＷＦ面取巾表下限
    HWFMWFMX As Integer             ' 品ＷＦ面取巾表上限
    HWFMWBCN As Integer             ' 品ＷＦ面取巾裏中心
    HWFMWBMN As Integer             ' 品ＷＦ面取巾裏下限
    HWFMWBMX As Integer             ' 品ＷＦ面取巾裏上限
    HWFMHCEN As Integer             ' 品ＷＦ面取高中心
    HWFMHMIN As Integer             ' 品ＷＦ面取高下限
    HWFMHMAX As Integer             ' 品ＷＦ面取高上限
    HWFMPWCN As Integer             ' 品ＷＦ面取先端巾中心
    HWFMPWMN As Integer             ' 品ＷＦ面取先端巾下限
    HWFMPWMX As Integer             ' 品ＷＦ面取先端巾上限
    HWFMPRCN As Double              ' 品ＷＦ面取先端Ｒ中心
    HWFMPRMN As Double              ' 品ＷＦ面取先端Ｒ下限
    HWFMPRMX As Double              ' 品ＷＦ面取先端Ｒ上限
    HWFMBACEN As Double              ' 品ＷＦ面取裏角度中心　6/22 Yam
    HWFMBAMIN As Double              ' 品ＷＦ面取裏角度下限
    HWFMBAMAX As Double              ' 品ＷＦ面取裏角度上限
    HWFDMFRM As String * 1          ' 品ＷＦ溝面取形状
    HWFDMM As String * 1            ' 品ＷＦ溝面取面粗
    HWFDMACN As Double              ' 品ＷＦ溝面取角度中心
    HWFDMPRC As Double              ' 品ＷＦ溝面取先端Ｒ中心
    HWFIDKBU As String * 1          ' 品ＷＦＩＤ区分有無
    HWFIDWAY As String * 1          ' 品ＷＦＩＤ方法
    HWFIDPRI As String * 1          ' 品ＷＦＩＤ印字種類
    HWFIDKND As String * 1          ' 品ＷＦＩＤ種類
    HWFIDDIR As String * 1          ' 品ＷＦＩＤ方向
    HWFIDFAC As String * 1          ' 品ＷＦＩＤ面
    HWFCSIZE As String * 1          ' 品ＷＦ文字サイズ
    HWFIDPBS As String * 1          ' 品ＷＦＩＤ位置測定基準
    HWFIDFIG As Integer             ' 品ＷＦＩＤ桁数
    HWFIDCON As String              ' 品ＷＦＩＤ内容
    HWFIDZAR As Double              ' 品ＷＦＩＤ除外領域
    HWFIDPAP As String * 1          ' 品ＷＦＩＤ印字連番指定
    HWFIDDCN As Integer             ' 品ＷＦＩＤドット深中心
    HWFIDDMX As Integer             ' 品ＷＦＩＤドット深上限
    HWFIDDMN As Integer             ' 品ＷＦＩＤドット深下限
    HWFIDSCN As Integer             ' 品ＷＦＩＤドットＳ中心
    HWFIDSMX As Integer             ' 品ＷＦＩＤドットＳ上限
    HWFIDSMN As Integer             ' 品ＷＦＩＤドットＳ下限
    HWFIDBCZ As String * 3          ' 品ＷＦＩＤＢＣ詳細図面
    HWFIDZNO As Long                ' 品ＷＦＩＤ図番号
    HWFBDPRS As Double              ' 品ＷＦＢＤ圧力
    HWFBDTIM As Integer             ' 品ＷＦＢＤ回数
    HWFETWAY As String * 2          ' 品ＷＦＥＴ方法
    HWFMPFIN As String * 1          ' 品ＷＦＭＰ仕上げ
    HWFLWASW As String * 1          ' 品ＷＦ最終洗浄方法
    HWFCDOP As String * 1           ' 品ＷＦ結晶ドープ
    HWFCDOPN As Double              ' 品ＷＦ結晶ドープ濃度
    HWFCDPNI As String * 2          ' 品ＷＦ結晶ドープ濃度指数
    HWFCMPUL As String * 1          ' 品ＷＦＣＭＰウネリレベル
    HWFTPROC As String * 1          ' 品ＷＦ耐圧加工
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 製品仕様WFﾃﾞｰﾀ４
Public Type typ_TBCME024
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    HMGSTRRNO As String * 9         ' 品管理仕様登録依頼番号
    HMGSTFNO As String * 8          ' 品管理社員Ｎｏ
    HMGWFSNO As String * 6          ' 品管理ＷＦ製品番号
    HMGWFSNE As Integer             ' 品管理ＷＦ製品番号枝番
    HWFM1S As String * 1            ' 品ＷＦ膜１種
    HWFM1H As String * 1            ' 品ＷＦ膜１付面
    HWFM2S As String * 1            ' 品ＷＦ膜２種
    HWFM2H As String * 1            ' 品ＷＦ膜２付面
    HWFNJSUM As String * 1          ' 品ＷＦノジュール処理有無
    HWFNJSMX As Double              ' 品ＷＦノジュール処理巾上限
    HWFNJSMN As Double              ' 品ＷＦノジュール処理巾下限
    HWFOXCEN As Long                ' 品ＷＦ酸化膜厚中心
    HWFOXMIN As Long                ' 品ＷＦ酸化膜厚下限
    HWFOXMAX As Long                ' 品ＷＦ酸化膜厚上限
    HWFOXSPH As String * 1          ' 品ＷＦ酸化膜厚測定位置＿方
    HWFOXSPT As String * 1          ' 品ＷＦ酸化膜厚測定位置＿点
    HWFOXSPI As String * 1          ' 品ＷＦ酸化膜厚測定位置＿位
    HWFOXHWT As String * 1          ' 品ＷＦ酸化膜厚保証方法＿対
    HWFOXHWS As String * 1          ' 品ＷＦ酸化膜厚保証方法＿処
    HWFOXHWY As String * 2          ' 品ＷＦ酸化膜厚検査方法
    HWFOXNPO As String * 1          ' 品ＷＦ酸化膜厚抜取位置
    HWFOXKHM As String * 1          ' 品ＷＦ酸化膜厚検査頻度＿枚
    HWFOXKHN As String * 1          ' 品ＷＦ酸化膜厚検査頻度＿抜
    HWFOXKHH As String * 1          ' 品ＷＦ酸化膜厚検査頻度＿保
    HWFOXKHU As String * 1          ' 品ＷＦ酸化膜厚検査頻度＿ウ
    HWFOXZAR As Integer             ' 品ＷＦ酸化膜除外領域
    HWFOXMBP As Double              ' 品ＷＦ酸化膜厚面内分布
    HWFOXMCL As String * 1          ' 品ＷＦ酸化膜厚面内計算
    HWFOXMRA As Integer             ' 品ＷＦ酸化膜厚面内範囲
    HWFOXLTB As Double              ' 品ＷＦ酸化膜厚ＬＴ分布
    HWFOXLTC As String * 1          ' 品ＷＦ酸化膜厚ＬＴ計算
    HWFOXLTR As Integer             ' 品ＷＦ酸化膜厚ＬＴ範囲
    HWFPSCEN As Long                ' 品ＷＦポリシリ厚中心
    HWFPSMIN As Long                ' 品ＷＦポリシリ厚下限
    HWFPSMAX As Long                ' 品ＷＦポリシリ厚上限
    HWFPSSPH As String * 1          ' 品ＷＦポリシリ厚測定位置＿方
    HWFPSSPT As String * 1          ' 品ＷＦポリシリ厚測定位置＿点
    HWFPSSPI As String * 1          ' 品ＷＦポリシリ厚測定位置＿位
    HWFPSHWT As String * 1          ' 品ＷＦポリシリ厚保証方法＿対
    HWFPSHWS As String * 1          ' 品ＷＦポリシリ厚保証方法＿処
    HWFPSKWY As String * 2          ' 品ＷＦポリシリ厚検査方法
    HWFPSNPS As String * 1          ' 品ＷＦポリシリ厚抜取位置
    HWFPSKHM As String * 1          ' 品ＷＦポリシリ厚検査頻度＿枚
    HWFPSKHN As String * 1          ' 品ＷＦポリシリ厚検査頻度＿抜
    HWFPSKHH As String * 1          ' 品ＷＦポリシリ厚検査頻度＿保
    HWFPSKHU As String * 1          ' 品ＷＦポリシリ厚検査頻度＿ウ
    HWFPSMBP As Double              ' 品ＷＦポリシリ厚面内分布
    HWFPSMCL As String * 1          ' 品ＷＦポリシリ厚面内計算
    HWFPSMRA As Integer             ' 品ＷＦポリシリ厚面内範囲
    HWFNOXCN As Long                ' 品ＷＦ窒化膜厚中心
    HWFNOXMN As Long                ' 品ＷＦ窒化膜厚下限
    HWFNOXMX As Long                ' 品ＷＦ窒化膜厚上限
    HWFNOXSH As String * 1          ' 品ＷＦ窒化膜厚測定位置＿方
    HWFNOXST As String * 1          ' 品ＷＦ窒化膜厚測定位置＿点
    HWFNOXSI As String * 1          ' 品ＷＦ窒化膜厚測定位置＿位
    HWFNOXHT As String * 1          ' 品ＷＦ窒化膜厚保証方法＿対
    HWFNOXHS As String * 1          ' 品ＷＦ窒化膜厚保証方法＿処
    HWFNOXHW As String * 2          ' 品ＷＦ窒化膜厚検査方法
    HWFNOXNP As String * 1          ' 品ＷＦ窒化膜厚抜取位置
    HWFNOXKM As String * 1          ' 品ＷＦ窒化膜厚検査頻度＿枚
    HWFNOXKN As String * 1          ' 品ＷＦ窒化膜厚検査頻度＿抜
    HWFNOXKH As String * 1          ' 品ＷＦ窒化膜厚検査頻度＿保
    HWFNOXKU As String * 1          ' 品ＷＦ窒化膜厚検査頻度＿ウ
    HWFNOXMB As Double              ' 品ＷＦ窒化膜厚面内分布
    HWFNOXMC As String * 1          ' 品ＷＦ窒化膜厚面内計算
    HWFNOXMR As Integer             ' 品ＷＦ窒化膜厚面内範囲
    HWFMKMIN As Double              ' 品ＷＦ無欠陥層下限
    HWFMKMAX As Double              ' 品ＷＦ無欠陥層上限
    HWFMKSPH As String * 1          ' 品ＷＦ無欠陥層測定位置＿方
    HWFMKSPT As String * 1          ' 品ＷＦ無欠陥層測定位置＿点
    HWFMKSPR As String * 1          ' 品ＷＦ無欠陥層測定位置＿領
    HWFMKHWT As String * 1          ' 品ＷＦ無欠陥層保証方法＿対
    HWFMKHWS As String * 1          ' 品ＷＦ無欠陥層保証方法＿処
    HWFMKSZY As String * 1          ' 品ＷＦ無欠陥層測定条件
    HWFMKKHM As String * 1          ' 品ＷＦ無欠陥層検査頻度＿枚
    HWFMKKHN As String * 1          ' 品ＷＦ無欠陥層検査頻度＿抜
    HWFMKKHH As String * 1          ' 品ＷＦ無欠陥層検査頻度＿保
    HWFMKKHU As String * 1          ' 品ＷＦ無欠陥層検査頻度＿ウ
    HWFMKNSW As String * 2          ' 品ＷＦ無欠陥層熱処理法
    HWFMKCET As Integer             ' 品ＷＦ無欠陥層選択ＥＴ代
    HWFDZSWY As String * 1          ' 品ＷＦＤＺ処理方法
    HWFD1STO As Integer             ' 品ＷＦＤＺ１ＳＴ温度
    HWFD1STT As Integer             ' 品ＷＦＤＺ１ＳＴ時間
    HWFD1STG As String * 1          ' 品ＷＦＤＺ１ＳＴガス条件
    HWFD2NDO As Integer             ' 品ＷＦＤＺ２ＮＤ温度
    HWFD2NDC As Integer             ' 品ＷＦＤＺ２ＮＤ温度定常
    HWFD2NDT As Integer             ' 品ＷＦＤＺ２ＮＤ時間
    HWFD3RDO As Integer             ' 品ＷＦＤＺ３ＲＤ温度
    HWFD3RDT As Integer             ' 品ＷＦＤＺ３ＲＤ時間
    HWFDZMPS As String * 1          ' 品ＷＦＤＺＭＰ処理区分
    HWFH2ANO As Integer             ' 品ＷＦＨ２ＡＮ温度
    HWFH2ANT As Integer             ' 品ＷＦＨ２ＡＮ時間
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
' 追加 2003/09.11 SystemBrain Start
    HWFANGZY As String * 1          ' 品ＷＦ高温ＡＮガス条件
' 追加 2003/09.11 SystemBrain End
End Type


' 製品仕様WFﾃﾞｰﾀ５
Public Type typ_TBCME025
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    HMGSTRRNO As String * 9         ' 品管理仕様登録依頼番号
    HMGSTFNO As String * 8          ' 品管理社員Ｎｏ
    HMGWFSNO As String * 6          ' 品管理ＷＦ製品番号
    HMGWFSNE As Integer             ' 品管理ＷＦ製品番号枝番
    HWFTMMAX As Long                ' 品ＷＦ転位密度上限
    HWFTMSPH As String * 1          ' 品ＷＦ転位密度測定位置＿方
    HWFTMSPT As String * 1          ' 品ＷＦ転位密度測定位置＿点
    HWFTMSPR As String * 1          ' 品ＷＦ転位密度測定位置＿領
    HWFTMKHM As String * 1          ' 品ＷＦ転位密度検査頻度＿枚
    HWFTMKHN As String * 1          ' 品ＷＦ転位密度検査頻度＿抜
    HWFTMKHH As String * 1          ' 品ＷＦ転位密度検査頻度＿保
    HWFTMKHU As String * 1          ' 品ＷＦ転位密度検査頻度＿ウ
    HWFLTMIN As Integer             ' 品ＷＦＬタイム下限
    HWFLTMAX As Integer             ' 品ＷＦＬタイム上限
    HWFLTSPH As String * 1          ' 品ＷＦＬタイム測定位置＿方
    HWFLTSPT As String * 1          ' 品ＷＦＬタイム測定位置＿点
    HWFLTSPI As String * 1          ' 品ＷＦＬタイム測定位置＿位
    HWFLTHWT As String * 1          ' 品ＷＦＬタイム保証方法＿対
    HWFLTHWS As String * 1          ' 品ＷＦＬタイム保証方法＿処
    HWFLTNSW As String * 2          ' 品ＷＦＬタイム熱処理法
    HWFLTKWY As String * 2          ' 品ＷＦＬタイム検査方法
    HWFLTKHM As String * 1          ' 品ＷＦＬタイム検査頻度＿枚
    HWFLTKHN As String * 1          ' 品ＷＦＬタイム検査頻度＿抜
    HWFLTKHH As String * 1          ' 品ＷＦＬタイム検査頻度＿保
    HWFLTKHU As String * 1          ' 品ＷＦＬタイム検査頻度＿ウ
    HWFLTMBP As Double              ' 品ＷＦＬタイム面内分布
    HWFLTMCL As String * 1          ' 品ＷＦＬタイム面内計算
    HWFCNMIN As Double              ' 品ＷＦ炭素濃度下限
    HWFCNMAX As Double              ' 品ＷＦ炭素濃度上限
    HWFCNSPH As String * 1          ' 品ＷＦ炭素濃度測定位置＿方
    HWFCNSPT As String * 1          ' 品ＷＦ炭素濃度測定位置＿点
    HWFCNSPI As String * 1          ' 品ＷＦ炭素濃度測定位置＿位
    HWFCNHWT As String * 1          ' 品ＷＦ炭素濃度保証方法＿対
    HWFCNHWS As String * 1          ' 品ＷＦ炭素濃度保証方法＿処
    HWFCNKWY As String * 2          ' 品ＷＦ炭素濃度検査方法
    HWFCNKHM As String * 1          ' 品ＷＦ炭素濃度検査頻度＿枚
    HWFCNKHN As String * 1          ' 品ＷＦ炭素濃度検査頻度＿抜
    HWFCNKHH As String * 1          ' 品ＷＦ炭素濃度検査頻度＿保
    HWFCNKHU As String * 1          ' 品ＷＦ炭素濃度検査頻度＿ウ
    HWFONMIN As Double              ' 品ＷＦ酸素濃度下限
    HWFONMAX As Double              ' 品ＷＦ酸素濃度上限
    HWFONSPH As String * 1          ' 品ＷＦ酸素濃度測定位置＿方
    HWFONSPT As String * 1          ' 品ＷＦ酸素濃度測定位置＿点
    HWFONSPI As String * 1          ' 品ＷＦ酸素濃度測定位置＿位
    HWFONHWT As String * 1          ' 品ＷＦ酸素濃度保証方法＿対
    HWFONHWS As String * 1          ' 品ＷＦ酸素濃度保証方法＿処
    HWFONKWY As String * 2          ' 品ＷＦ酸素濃度検査方法
    HWFONKHM As String * 1          ' 品ＷＦ酸素濃度検査頻度＿枚
    HWFONKHN As String * 1          ' 品ＷＦ酸素濃度検査頻度＿抜
    HWFONKHH As String * 1          ' 品ＷＦ酸素濃度検査頻度＿保
    HWFONKHU As String * 1          ' 品ＷＦ酸素濃度検査頻度＿ウ
    HWFONMBP As Double              ' 品ＷＦ酸素濃度面内分布
    HWFONMCL As String * 1          ' 品ＷＦ酸素濃度面内計算
    HWFONLTB As Double              ' 品ＷＦ酸素濃度ＬＴ分布
    HWFONLTC As String * 1          ' 品ＷＦ酸素濃度ＬＴ計算
    HWFONSDV As Double              ' 品ＷＦ酸素濃度標準偏差
    HWFONAMN As Double              ' 品ＷＦ酸素濃度平均下限
    HWFONAMX As Double              ' 品ＷＦ酸素濃度平均上限
    HWFOKBSH As String * 1          ' 品ＷＦ酸素振区分測定位置＿方
    HWFOKBST As String * 1          ' 品ＷＦ酸素振区分測定位置＿点
    HWFOKBSI As String * 1          ' 品ＷＦ酸素振区分測定位置＿位
    HWFOKBHT As String * 1          ' 品ＷＦ酸素振区分保証方法＿対
    HWFOKBHS As String * 1          ' 品ＷＦ酸素振区分保証方法＿処
    HWFOS1MN As Double              ' 品ＷＦ酸素析出１下限
    HWFOS1MX As Double              ' 品ＷＦ酸素析出１上限
    HWFOS1NS As String * 2          ' 品ＷＦ酸素析出１熱処理法
    HWFOS1SH As String * 1          ' 品ＷＦ酸素析出１測定位置＿方
    HWFOS1ST As String * 1          ' 品ＷＦ酸素析出１測定位置＿点
    HWFOS1SI As String * 1          ' 品ＷＦ酸素析出１測定位置＿位
    HWFOS1HT As String * 1          ' 品ＷＦ酸素析出１保証方法＿対
    HWFOS1HS As String * 1          ' 品ＷＦ酸素析出１保証方法＿処
    HWFOS1HM As String * 1          ' 品ＷＦ酸素析出１検査頻度＿枚
    HWFOS1KN As String * 1          ' 品ＷＦ酸素析出１検査頻度＿抜
    HWFOS1KH As String * 1          ' 品ＷＦ酸素析出１検査頻度＿保
    HWFOS1KU As String * 1          ' 品ＷＦ酸素析出１検査頻度＿ウ
    HWFOS2MN As Double              ' 品ＷＦ酸素析出２下限
    HWFOS2MX As Double              ' 品ＷＦ酸素析出２上限
    HWFOS2NS As String * 2          ' 品ＷＦ酸素析出２熱処理法
    HWFOS2SH As String * 1          ' 品ＷＦ酸素析出２測定位置＿方
    HWFOS2ST As String * 1          ' 品ＷＦ酸素析出２測定位置＿点
    HWFOS2SI As String * 1          ' 品ＷＦ酸素析出２測定位置＿位
    HWFOS2HT As String * 1          ' 品ＷＦ酸素析出２保証方法＿対
    HWFOS2HS As String * 1          ' 品ＷＦ酸素析出２保証方法＿処
    HWFOS2KM As String * 1          ' 品ＷＦ酸素析出２検査頻度＿枚
    HWFOS2KN As String * 1          ' 品ＷＦ酸素析出２検査頻度＿抜
    HWFOS2KH As String * 1          ' 品ＷＦ酸素析出２検査頻度＿保
    HWFOS2KU As String * 1          ' 品ＷＦ酸素析出２検査頻度＿ウ
    HWFOS3MN As Double              ' 品ＷＦ酸素析出３下限
    HWFOS3MX As Double              ' 品ＷＦ酸素析出３上限
    HWFOS3NS As String * 2          ' 品ＷＦ酸素析出３熱処理法
    HWFOS3SH As String * 1          ' 品ＷＦ酸素析出３測定位置＿方
    HWFOS3ST As String * 1          ' 品ＷＦ酸素析出３測定位置＿点
    HWFOS3SI As String * 1          ' 品ＷＦ酸素析出３測定位置＿位
    HWFOS3HT As String * 1          ' 品ＷＦ酸素析出３保証方法＿対
    HWFOS3HS As String * 1          ' 品ＷＦ酸素析出３保証方法＿処
    HWFOS3KM As String * 1          ' 品ＷＦ酸素析出３検査頻度＿枚
    HWFOS3KN As String * 1          ' 品ＷＦ酸素析出３検査頻度＿抜
    HWFOS3KH As String * 1          ' 品ＷＦ酸素析出３検査頻度＿保
    HWFOS3KU As String * 1          ' 品ＷＦ酸素析出３検査頻度＿ウ
    HWFANTNP As Integer             ' 品ＷＦＡＮ温度
    HWFANTIM As Integer             ' 品ＷＦＡＮ時間
    HWFANTMN As Integer             ' 品ＷＦＡＮ時間下限
    HWFANTMX As Integer             ' 品ＷＦＡＮ時間上限
    HWFZOMIN As Double              ' 品ＷＦ残存酸素下限
    HWFZOMAX As Double              ' 品ＷＦ残存酸素上限
    HWFZOSPH As String * 1          ' 品ＷＦ残存酸素測定位置＿方
    HWFZOSPT As String * 1          ' 品ＷＦ残存酸素測定位置＿点
    HWFZOSPI As String * 1          ' 品ＷＦ残存酸素測定位置＿位
    HWFZOHWT As String * 1          ' 品ＷＦ残存酸素保証方法＿対
    HWFZOHWS As String * 1          ' 品ＷＦ残存酸素保証方法＿処
    HWFZONSW As String * 2          ' 品ＷＦ残存酸素熱処理法
    HWFZOKWY As String * 2          ' 品ＷＦ残存酸素検査方法
    HWFZOKHM As String * 1          ' 品ＷＦ残存酸素検査頻度＿枚
    HWFZOKHN As String * 1          ' 品ＷＦ残存酸素検査頻度＿抜
    HWFZOKHH As String * 1          ' 品ＷＦ残存酸素検査頻度＿保
    HWFZOKHU As String * 1          ' 品ＷＦ残存酸素検査頻度＿ウ
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
' 追加 2003/09.11 SystemBrain Start
    HWFTMMAXN As Double             ' 品ＷＦ転位密度上限
    HWFANTTAN As String * 1         ' 品ＷＦＡＮ時間単位
' 追加 2003/09.11 SystemBrain End
End Type


' 製品仕様WFﾃﾞｰﾀ６
Public Type typ_TBCME026
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    HMGSTRRNO As String * 9         ' 品管理仕様登録依頼番号
    HMGSTFNO As String * 8          ' 品管理社員Ｎｏ
    HMGWFSNO As String * 6          ' 品管理ＷＦ製品番号
    HMGWFSNE As Integer             ' 品管理ＷＦ製品番号枝番
    HWFBDOMN As Integer             ' 品ＷＦＢＤＯＳＦ下限
    HWFBDOMX As Integer             ' 品ＷＦＢＤＯＳＦ上限
    HWFBDOSH As String * 1          ' 品ＷＦＢＤＯＳＦ測定位置＿方
    HWFBDOST As String * 1          ' 品ＷＦＢＤＯＳＦ測定位置＿点
    HWFBDOSR As String * 1          ' 品ＷＦＢＤＯＳＦ測定位置＿領
    HWFBDOHT As String * 1          ' 品ＷＦＢＤＯＳＦ保証方法＿対
    HWFBDOHS As String * 1          ' 品ＷＦＢＤＯＳＦ保証方法＿処
    HWFBDOSZ As String * 1          ' 品ＷＦＢＤＯＳＦ測定条件
    HWFBDONS As String * 2          ' 品ＷＦＢＤＯＳＦ熱処理法
    HWFBDOKM As String * 1          ' 品ＷＦＢＤＯＳＦ検査頻度＿枚
    HWFBDOKN As String * 1          ' 品ＷＦＢＤＯＳＦ検査頻度＿抜
    HWFBDOKH As String * 1          ' 品ＷＦＢＤＯＳＦ検査頻度＿保
    HWFBDOKU As String * 1          ' 品ＷＦＢＤＯＳＦ検査頻度＿ウ
    HWFBDOET As Integer             ' 品ＷＦＢＤＯＳＦ選択ＥＴ代
    HWFBDSMN As Integer             ' 品ＷＦＢＤＳＴ跡下限
    HWFBDSMX As Integer             ' 品ＷＦＢＤＳＴ跡上限
    HWFBDSSH As String * 1          ' 品ＷＦＢＤＳＴ跡測定位置＿方
    HWFBDSST As String * 1          ' 品ＷＦＢＤＳＴ跡測定位置＿点
    HWFBDSSR As String * 1          ' 品ＷＦＢＤＳＴ跡測定位置＿領
    HWFBDSHT As String * 1          ' 品ＷＦＢＤＳＴ跡保証方法＿対
    HWFBDSHS As String * 1          ' 品ＷＦＢＤＳＴ跡保証方法＿処
    HWFBDSSZ As String * 1          ' 品ＷＦＢＤＳＴ跡測定条件
    HWFBDSNS As String * 2          ' 品ＷＦＢＤＳＴ跡熱処理法
    HWFBDSKM As String * 1          ' 品ＷＦＢＤＳＴ跡検査頻度＿枚
    HWFBDSKN As String * 1          ' 品ＷＦＢＤＳＴ跡検査頻度＿抜
    HWFBDSKH As String * 1          ' 品ＷＦＢＤＳＴ跡検査頻度＿保
    HWFBDSKU As String * 1          ' 品ＷＦＢＤＳＴ跡検査頻度＿ウ
    HWFBDSET As Integer             ' 品ＷＦＢＤＳＴ跡選択ＥＴ代
    HWFRNFMX As Double              ' 品ＷＦラフネス表上限
    HWFRNFSH As String * 1          ' 品ＷＦラフネス表測定位置＿方
    HWFRNFST As String * 1          ' 品ＷＦラフネス表測定位置＿点
    HWFRNFSI As String * 1          ' 品ＷＦラフネス表測定位置＿位
    HWFRNFKW As String * 2          ' 品ＷＦラフネス表検査方法
    HWFRNFZA As Integer             ' 品ＷＦラフネス表除外領域
    HWFRNBMX As Double              ' 品ＷＦラフネス裏上限
    HWFRNBSH As String * 1          ' 品ＷＦラフネス裏測定位置＿方
    HWFRNBST As String * 1          ' 品ＷＦラフネス裏測定位置＿点
    HWFRNBSI As String * 1          ' 品ＷＦラフネス裏測定位置＿位
    HWFRNBKW As String * 2          ' 品ＷＦラフネス裏検査方法
    HWFRNBZA As Integer             ' 品ＷＦラフネス裏除外領域
    HWFDENKU As String * 1          ' 品ＷＦＤｅｎ検査有無
    HWFDENMX As Integer             ' 品ＷＦＤｅｎ上限
    HWFDENMN As Integer             ' 品ＷＦＤｅｎ下限
    HWFDENHT As String * 1          ' 品ＷＦＤｅｎ保証方法＿対
    HWFDENHS As String * 1          ' 品ＷＦＤｅｎ保証方法＿処
    HWFDVDKU As String * 1          ' 品ＷＦＤＶＤ２検査有無
    HWFDVDMX As Integer             ' 品ＷＦＤＶＤ２上限
    HWFDVDMN As Integer             ' 品ＷＦＤＶＤ２下限
    HWFDVDHT As String * 1          ' 品ＷＦＤＶＤ２保証方法＿対
    HWFDVDHS As String * 1          ' 品ＷＦＤＶＤ２保証方法＿処
    HWFLDLKU As String * 1          ' 品ＷＦＬ／ＤＬ検査有無
    HWFLDLMX As Integer             ' 品ＷＦＬ／ＤＬ上限
    HWFLDLMN As Integer             ' 品ＷＦＬ／ＤＬ下限
    HWFLDLHT As String * 1          ' 品ＷＦＬ／ＤＬ保証方法＿対
    HWFLDLHS As String * 1          ' 品ＷＦＬ／ＤＬ保証方法＿処
    HWFGDSPH As String * 1          ' 品ＷＦＧＤ測定位置＿方
    HWFGDSPT As String * 1          ' 品ＷＦＧＤ測定位置＿点
    HWFGDSPR As String * 1          ' 品ＷＦＧＤ測定位置＿領
    HWFGDSZY As String * 1          ' 品ＷＦＧＤ測定条件
    HWFGDZAR As Integer             ' 品ＷＦＧＤ除外領域
    HWFGDKHM As String * 1          ' 品ＷＦＧＤ検査頻度＿枚
    HWFGDKHN As String * 1          ' 品ＷＦＧＤ検査頻度＿抜
    HWFGDKHH As String * 1          ' 品ＷＦＧＤ検査頻度＿保
    HWFGDKHU As String * 1          ' 品ＷＦＧＤ検査頻度＿ウ
    HWFDSOKE As String * 1          ' 品ＷＦＤＳＯＤ検査
    HWFDSOMX As Long                ' 品ＷＦＤＳＯＤ上限
    HWFDSOMN As Long                ' 品ＷＦＤＳＯＤ下限
    HWFDSOAX As Integer             ' 品ＷＦＤＳＯＤ領域上限
    HWFDSOAN As Integer             ' 品ＷＦＤＳＯＤ領域下限
    HWFDSOHT As String * 1          ' 品ＷＦＤＳＯＤ保証方法＿対
    HWFDSOHS As String * 1          ' 品ＷＦＤＳＯＤ保証方法＿処
    HWFDSOKM As String * 1          ' 品ＷＦＤＳＯＤ検査頻度＿枚
    HWFDSOKN As String * 1          ' 品ＷＦＤＳＯＤ検査頻度＿抜
    HWFDSOKH As String * 1          ' 品ＷＦＤＳＯＤ検査頻度＿保
    HWFDSOKU As String * 1          ' 品ＷＦＤＳＯＤ検査頻度＿ウ
    HWFNTPUM As String * 1          ' 品ＷＦ平坦ナノトポ有無
    HWFNTPK1 As Double              ' 品ＷＦ平坦ナノトポ規格１
    HWFNTPP1 As Double              ' 品ＷＦ平坦ナノトポＰＵＡ１
    HWFNTPS1 As Double              ' 品ＷＦ平坦ナノトポサイト１
    HWFNTPK2 As Double              ' 品ＷＦ平坦ナノトポ規格２
    HWFNTPP2 As Double              ' 品ＷＦ平坦ナノトポＰＵＡ２
    HWFNTPS2 As Double              ' 品ＷＦ平坦ナノトポサイト２
    HWFNTPK3 As Double              ' 品ＷＦ平坦ナノトポ規格３
    HWFNTPP3 As Double              ' 品ＷＦ平坦ナノトポＰＵＡ３
    HWFNTPS3 As Double              ' 品ＷＦ平坦ナノトポサイト３
    HWFNTPZA As Integer             ' 品ＷＦ平坦ナノトポ除外領域
    HWFNTPHT As String * 1          ' 品ＷＦ平坦ナノトポ保証方法＿対
    HWFNTPHS As String * 1          ' 品ＷＦ平坦ナノトポ保証方法＿処
    HWFNTPKM As String * 1          ' 品ＷＦ平坦ナノトポ検査頻度＿枚
    HWFNTPKN As String * 1          ' 品ＷＦ平坦ナノトポ検査頻度＿抜
    HWFNTPKH As String * 1          ' 品ＷＦ平坦ナノトポ検査頻度＿保
    HWFNTPKU As String * 1          ' 品ＷＦ平坦ナノトポ検査頻度＿ウ
    HWFCRSSK As String * 1          ' 品ＷＦ平坦クロスＳＳ検査
    HWFMDCEN As Double              ' 品ＷＦ平坦面ダレ高低差中心
    HWFMDMAX As Double              ' 品ＷＦ平坦面ダレ高低差上限
    HWFMDMIN As Double              ' 品ＷＦ平坦面ダレ高低差下限
    HWFMDSPH As String * 1          ' 品ＷＦ平坦面ダレ測定位置＿方
    HWFMDSPT As String * 1          ' 品ＷＦ平坦面ダレ測定位置＿点
    HWFMDSPI As String * 1          ' 品ＷＦ平坦面ダレ測定位置＿位
    HWFMDHWT As String * 1          ' 品ＷＦ平坦面ダレ保証方法＿対
    HWFMDHWS As String * 1          ' 品ＷＦ平坦面ダレ保証方法＿処
    HWFMDKHM As String * 1          ' 品ＷＦ平坦面ダレ検査頻度＿枚
    HWFMDKHN As String * 1          ' 品ＷＦ平坦面ダレ検査頻度＿抜
    HWFMDKHH As String * 1          ' 品ＷＦ平坦面ダレ検査頻度＿保
    HWFMDKHU As String * 1          ' 品ＷＦ平坦面ダレ検査頻度＿ウ
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
' 追加 2003/09.11 SystemBrain Start
    HWFDVDMXN As Integer            ' 品WFDVD2上限
    HWFDVDMNN As Integer            ' 品WFDVD2下限
    HWFDSONWY As String * 2         ' 品WFDSOD熱処理法
    HWFMSUMX As Integer             ' 品WFMスクラッチ上限
    HWFMSUZY As String * 1          ' 品WFMスクラッチ測定条件
    HWFMSUKW As String * 1          ' 品WFMスクラッチ検査方法
    HWFMSUSZ As Double              ' 品WFMスクラッチサイズ
    KSTAFFID As String * 8          ' 更新社員ID
    sStaffID As String * 8          ' 承認社員ID
    SYNFLAG As String * 1           ' 承認フラグ
    SYNDATE As Date                 ' 承認日付
    HWFNP1AR As Double              ' 品WFナノトポ1エリア
    HWFNP1MAX As Double             ' 品WFナノトポ1上限
    HWFNP2AR As Double              ' 品WFナノトポ2エリア
    HWFNP2MAX As Double             ' 品WFナノトポ2上限
' 追加 2003/09.11 SystemBrain End
End Type


' 製品仕様WFﾃﾞｰﾀ７
Public Type typ_TBCME027
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    HMGSTRRNO As String * 9         ' 品管理仕様登録依頼番号
    HMGSTFNO As String * 8          ' 品管理社員Ｎｏ
    HMGWFSNO As String * 6          ' 品管理ＷＦ製品番号
    HMGWFSNE As Integer             ' 品管理ＷＦ製品番号枝番
    HWFSMIN As Double               ' 品ＷＦ反り下限
    HWFSMAX As Double               ' 品ＷＦ反り上限
    HWFSHWYT As String * 1          ' 品ＷＦ反り保証方法＿対
    HWFSHWYS As String * 1          ' 品ＷＦ反り保証方法＿処
    HWFSKWAY As String * 2          ' 品ＷＦ反り検査方法
    HWFSKHM As String * 1           ' 品ＷＦ反り検査頻度＿枚
    HWFSKHN As String * 1           ' 品ＷＦ反り検査頻度＿抜
    HWFSKHH As String * 1           ' 品ＷＦ反り検査頻度＿保
    HWFSKHU As String * 1           ' 品ＷＦ反り検査頻度＿ウ
    HWFSSZYO As String * 1          ' 品ＷＦ反り測定条件
    'HWFSZARA As Integer             ' 品ＷＦ反り除外領域
    HWFSZARAN As Double             ' 品ＷＦ反り除外領域 6/22 Yam
    HWFSSDEV As Double              ' 品ＷＦ反り標準偏差
    HWFSAMIN As Double              ' 品ＷＦ反り平均下限
    HWFSAMAX As Double              ' 品ＷＦ反り平均上限
    HWFSSREC As String * 1          ' 品ＷＦ反り測定器
    HWFSBO1 As Double               ' 品ＷＦ反り境界１
    HWFSBO1B As Integer             ' 品ＷＦ反り境界１下
    HWFSBO2 As Double               ' 品ＷＦ反り境界２
    HWFSBO2B As Integer             ' 品ＷＦ反り境界２下
    HWFSBO3 As Double               ' 品ＷＦ反り境界３
    HWFSBO3B As Integer             ' 品ＷＦ反り境界３下
    HWFWARMX As Double              ' 品ＷＦＷＡＲＰ上限
    HWFWARSZ As String * 1          ' 品ＷＦＷＡＲＰ測定条件
    HWFWARHT As String * 1          ' 品ＷＦＷＡＲＰ保証方法＿対
    HWFWARHS As String * 1          ' 品ＷＦＷＡＲＰ保証方法＿処
    HWFWARKW As String * 2          ' 品ＷＦＷＡＲＰ検査方法
    'HWFWARZA As Integer             ' 品ＷＦＷＡＲＰ除外領域
    HWFWARZAN As Double             ' 品ＷＦＷＡＲＰ除外領域 6/22 Yam
    HWFWARKM As String * 1          ' 品ＷＦＷＡＲＰ検査頻度＿枚
    HWFWARKN As String * 1          ' 品ＷＦＷＡＲＰ検査頻度＿抜
    HWFWARKH As String * 1          ' 品ＷＦＷＡＲＰ検査頻度＿保
    HWFWARKU As String * 1          ' 品ＷＦＷＡＲＰ検査頻度＿ウ
    HWFWARSR As String * 1          ' 品ＷＦＷＡＲＰ測定器
    HWFWAB1 As Double               ' 品ＷＦＷＡＲＰ境界１
    HWFWAB1B As Integer             ' 品ＷＦＷＡＲＰ境界１下
    HWFWAB2 As Double               ' 品ＷＦＷＡＲＰ境界２
    HWFWAB2B As Integer             ' 品ＷＦＷＡＲＰ境界２下
    HWFWAB3 As Double               ' 品ＷＦＷＡＲＰ境界３
    HWFWAB3B As Integer             ' 品ＷＦＷＡＲＰ境界３下
    HWFWARPR As String * 1          ' 品ＷＦＷａｒｐランク
    HWFFSZYO As String * 1          ' 品ＷＦ平坦測定条件
    HWFFSREC As String * 1          ' 品ＷＦ平坦測定器
    HWFGBMAX As Double              ' 品ＷＦ平坦ＧＢ上限
    HWFGBPUG As Double              ' 品ＷＦ平坦ＧＢＰＵＡ限
    HWFGBPUR As Integer             ' 品ＷＦ平坦ＧＢＰＵＡ率
    HWFGBHWT As String * 1          ' 品ＷＦ平坦ＧＢ保証方法＿対
    HWFGBHWS As String * 1          ' 品ＷＦ平坦ＧＢ保証方法＿処
    HWFGBKW As String * 4           ' 品ＷＦ平坦ＧＢ検査方法
    'HWFGBZAR As Integer             ' 品ＷＦ平坦ＧＢ除外領域
    HWFGBZARN As Double             ' 品ＷＦ平坦ＧＢ除外領域
    HWFGBKHM As String * 1          ' 品ＷＦ平坦ＧＢ検査頻度＿枚
    HWFGBKHN As String * 1          ' 品ＷＦ平坦ＧＢ検査頻度＿抜
    HWFGBKHH As String * 1          ' 品ＷＦ平坦ＧＢ検査頻度＿保
    HWFGBKHU As String * 1          ' 品ＷＦ平坦ＧＢ検査頻度＿ウ
    HWFGBB1 As Double               ' 品ＷＦ平坦ＧＢ境界１
    HWFGBB1B As Integer             ' 品ＷＦ平坦ＧＢ境界１下
    HWFGBB2 As Double               ' 品ＷＦ平坦ＧＢ境界２
    HWFGBB2B As Integer             ' 品ＷＦ平坦ＧＢ境界２下
    HWFGBB3 As Double               ' 品ＷＦ平坦ＧＢ境界３
    HWFGBB3B As Integer             ' 品ＷＦ平坦ＧＢ境界３下
    HWFGFDMX As Double              ' 品ＷＦ平坦ＧＦＤ上限
    HWFGFDPG As Double              ' 品ＷＦ平坦ＧＦＤＰＵＡ限
    HWFGFDPR As Integer             ' 品ＷＦ平坦ＧＦＤＰＵＡ率
    HWFGFDHT As String * 1          ' 品ＷＦ平坦ＧＦＤ保証方法＿対
    HWFGFDHS As String * 1          ' 品ＷＦ平坦ＧＦＤ保証方法＿処
    HWFGFDKW As String * 4          ' 品ＷＦ平坦ＧＦＤ検査方法
    'HWFGFDZA As Integer             ' 品ＷＦ平坦ＧＦＤ除外領域
    HWFGFDZAN As Double             ' 品ＷＦ平坦ＧＦＤ除外領域
    HWFGFDKM As String * 1          ' 品ＷＦ平坦ＧＦＤ検査頻度＿枚
    HWFGFDKN As String * 1          ' 品ＷＦ平坦ＧＦＤ検査頻度＿抜
    HWFGFDKH As String * 1          ' 品ＷＦ平坦ＧＦＤ検査頻度＿保
    HWFGFDKU As String * 1          ' 品ＷＦ平坦ＧＦＤ検査頻度＿ウ
    HWFGDB1 As Double               ' 品ＷＦ平坦ＧＦＤ境界１
    HWFGDB1B As Integer             ' 品ＷＦ平坦ＧＦＤ境界１下
    HWFGDB2 As Double               ' 品ＷＦ平坦ＧＦＤ境界２
    HWFGDB2B As Integer             ' 品ＷＦ平坦ＧＦＤ境界２下
    HWFGDB3 As Double               ' 品ＷＦ平坦ＧＦＤ境界３
    HWFGDB3B As Integer             ' 品ＷＦ平坦ＧＦＤ境界３下
    HWFGFRMX As Double              ' 品ＷＦ平坦ＧＦＲ上限
    HWFGFRPG As Double              ' 品ＷＦ平坦ＧＦＲＰＵＡ限
    HWFGFRPR As Integer             ' 品ＷＦ平坦ＧＦＲＰＵＡ率
    HWFGFRHT As String * 1          ' 品ＷＦ平坦ＧＦＲ保証方法＿対
    HWFGFRHS As String * 1          ' 品ＷＦ平坦ＧＦＲ保証方法＿処
    HWFGFRKW As String * 4          ' 品ＷＦ平坦ＧＦＲ検査方法
    'HWFGFRZA As Integer             ' 品ＷＦ平坦ＧＦＲ除外領域
    HWFGFRZAN As Double             ' 品ＷＦ平坦ＧＦＲ除外領域
    HWFGFRKM As String * 1          ' 品ＷＦ平坦ＧＦＲ検査頻度＿枚
    HWFGFRKN As String * 1          ' 品ＷＦ平坦ＧＦＲ検査頻度＿抜
    HWFGFRKH As String * 1          ' 品ＷＦ平坦ＧＦＲ検査頻度＿保
    HWFGFRKU As String * 1          ' 品ＷＦ平坦ＧＦＲ検査頻度＿ウ
    HWFGRB1 As Double               ' 品ＷＦ平坦ＧＦＲ境界１
    HWFGRB1B As Integer             ' 品ＷＦ平坦ＧＦＲ境界１下
    HWFGRB2 As Double               ' 品ＷＦ平坦ＧＦＲ境界２
    HWFGRB2B As Integer             ' 品ＷＦ平坦ＧＦＲ境界２下
    HWFGRB3 As Double               ' 品ＷＦ平坦ＧＦＲ境界３
    HWFGRB3B As Integer             ' 品ＷＦ平坦ＧＦＲ境界３下
    HWFSBMAX As Double              ' 品ＷＦ平坦ＳＢ上限
    HWFSBPUG As Double              ' 品ＷＦ平坦ＳＢＰＵＡ限
    HWFSBPUR As Integer             ' 品ＷＦ平坦ＳＢＰＵＡ率
    HWFSBSZX As Double              ' 品ＷＦ平坦ＳＢサイズＸ
    HWFSBSZY As Double              ' 品ＷＦ平坦ＳＢサイズＹ
    HWFSBHWT As String * 1          ' 品ＷＦ平坦ＳＢ保証方法＿対
    HWFSBHWS As String * 1          ' 品ＷＦ平坦ＳＢ保証方法＿処
    HWFSBKW As String * 4           ' 品ＷＦ平坦ＳＢ検査方法
    'HWFSBZAR As Integer             ' 品ＷＦ平坦ＳＢ除外領域
    HWFSBZARN As Double             ' 品ＷＦ平坦ＳＢ除外領域
    HWFSBKHM As String * 1          ' 品ＷＦ平坦ＳＢ検査頻度＿枚
    HWFSBKHN As String * 1          ' 品ＷＦ平坦ＳＢ検査頻度＿抜
    HWFSBKHH As String * 1          ' 品ＷＦ平坦ＳＢ検査頻度＿保
    HWFSBKHU As String * 1          ' 品ＷＦ平坦ＳＢ検査頻度＿ウ
    HWFSBB1 As Double               ' 品ＷＦ平坦ＳＢ境界１
    HWFSBB1B As Integer             ' 品ＷＦ平坦ＳＢ境界１下
    HWFSBB2 As Double               ' 品ＷＦ平坦ＳＢ境界２
    HWFSBB2B As Integer             ' 品ＷＦ平坦ＳＢ境界２下
    HWFSBB3 As Double               ' 品ＷＦ平坦ＳＢ境界３
    HWFSBB3B As Integer             ' 品ＷＦ平坦ＳＢ境界３下
    HWFSFMAX As Double              ' 品ＷＦ平坦ＳＦ上限
    HWFSFPUG As Double              ' 品ＷＦ平坦ＳＦＰＵＡ限
    HWFSFPUR As Integer             ' 品ＷＦ平坦ＳＦＰＵＡ率
    HWFSFSZX As Double              ' 品ＷＦ平坦ＳＦサイズＸ
    HWFSFSZY As Double              ' 品ＷＦ平坦ＳＦサイズＹ
    HWFSFHWT As String * 1          ' 品ＷＦ平坦ＳＦ保証方法＿対
    HWFSFHWS As String * 1          ' 品ＷＦ平坦ＳＦ保証方法＿処
    HWFSFKW As String * 4           ' 品ＷＦ平坦ＳＦ検査方法
    'HWFSFZAR As Integer             ' 品ＷＦ平坦ＳＦ除外領域
    HWFSFZARN As Double             ' 品ＷＦ平坦ＳＦ除外領域
    HWFSFKHM As String * 1          ' 品ＷＦ平坦ＳＦ検査頻度＿枚
    HWFSFKHN As String * 1          ' 品ＷＦ平坦ＳＦ検査頻度＿抜
    HWFSFKHH As String * 1          ' 品ＷＦ平坦ＳＦ検査頻度＿保
    HWFSFKHU As String * 1          ' 品ＷＦ平坦ＳＦ検査頻度＿ウ
    HWFSFB1 As Double               ' 品ＷＦ平坦ＳＦ境界１
    HWFSFB1B As Integer             ' 品ＷＦ平坦ＳＦ境界１下
    HWFSFB2 As Double               ' 品ＷＦ平坦ＳＦ境界２
    HWFSFB2B As Integer             ' 品ＷＦ平坦ＳＦ境界２下
    HWFSFB3 As Double               ' 品ＷＦ平坦ＳＦ境界３
    HWFSFB3B As Integer             ' 品ＷＦ平坦ＳＦ境界３下
    HWFFSXOF As Double              ' 品ＷＦ平坦サイトＸＯＦ
    HWFFSYOF As Double              ' 品ＷＦ平坦サイトＹＯＦ
    HWFFPSUM As String * 1          ' 品ＷＦ平坦Ｐサイト有無
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
' 追加 2003/09.11 SystemBrain Start
    HWFSBPUAGN As Double            ' 品WF平坦SBPUA限
    HWFSBMAXN As Double             ' 品WF平坦SB上限
    HWFSBB1N As Double              ' 品WF平坦SB境界1
    HWFSBB2N As Double              ' 品WF平坦SB境界2
    HWFSBB3N As Double              ' 品WF平坦SB境界3
    HWFSFPUAGN As Double            ' 品WF平坦SFPUA限
    HWFSFMAXN As Double             ' 品WF平坦SF上限
    HWFSFB1N As Double              ' 品WF平坦SF境界1
    HWFSFB2N As Double              ' 品WF平坦SF境界2
    HWFSFB3N As Double              ' 品WF平坦SF境界3
' 追加 2003/09.11 SystemBrain End
End Type


' 製品仕様WFﾃﾞｰﾀ８
Public Type typ_TBCME028
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    HMGSTRRNO As String * 9         ' 品管理仕様登録依頼番号
    HMGSTFNO As String * 8          ' 品管理社員Ｎｏ
    HMGWFSNO As String * 6          ' 品管理ＷＦ製品番号
    HMGWFSNE As Integer             ' 品管理ＷＦ製品番号枝番
    HWFMK1SI As Double              ' 品ＷＦ面検欠陥１サイズ
    HWFMK1MX As Integer             ' 品ＷＦ面検欠陥１上限
    HWFMK1SZ As String * 1          ' 品ＷＦ面検欠陥１測定条件
    HWFMK1ZA As Integer             ' 品ＷＦ面検欠陥１除外領域
    HWFMK1HT As String * 1          ' 品ＷＦ面検欠陥１保証方法＿対
    HWFMK1HS As String * 1          ' 品ＷＦ面検欠陥１保証方法＿処
    HWFMK1KM As String * 1          ' 品ＷＦ面検欠陥１検査頻度＿枚
    HWFMK1KN As String * 1          ' 品ＷＦ面検欠陥１検査頻度＿抜
    HWFMK1KH As String * 1          ' 品ＷＦ面検欠陥１検査頻度＿保
    HWFMK1KU As String * 1          ' 品ＷＦ面検欠陥１検査頻度＿ウ
    HWFM1B1 As Integer              ' 品ＷＦ面検欠陥１境界１
    HWFM1B1B As Integer             ' 品ＷＦ面検欠陥１境界１下
    HWFM1B2 As Integer              ' 品ＷＦ面検欠陥１境界２
    HWFM1B2B As Integer             ' 品ＷＦ面検欠陥１境界２下
    HWFM1B3 As Integer              ' 品ＷＦ面検欠陥１境界３
    HWFM1B3B As Integer             ' 品ＷＦ面検欠陥１境界３下
    HWFMK2SI As Double              ' 品ＷＦ面検欠陥２サイズ
    HWFMK2MX As Integer             ' 品ＷＦ面検欠陥２上限
    HWFMK2HT As String * 1          ' 品ＷＦ面検欠陥２保証方法＿対
    HWFMK2HS As String * 1          ' 品ＷＦ面検欠陥２保証方法＿処
    HWFMK2KM As String * 1          ' 品ＷＦ面検欠陥２検査頻度＿枚
    HWFMK2KN As String * 1          ' 品ＷＦ面検欠陥２検査頻度＿抜
    HWFMK2KH As String * 1          ' 品ＷＦ面検欠陥２検査頻度＿保
    HWFMK2KU As String * 1          ' 品ＷＦ面検欠陥２検査頻度＿ウ
    HWFM2B1 As Integer              ' 品ＷＦ面検欠陥２境界１
    HWFM2B1B As Integer             ' 品ＷＦ面検欠陥２境界１下
    HWFM2B2 As Integer              ' 品ＷＦ面検欠陥２境界２
    HWFM2B2B As Integer             ' 品ＷＦ面検欠陥２境界２下
    HWFM2B3 As Integer              ' 品ＷＦ面検欠陥２境界３
    HWFM2B3B As Integer             ' 品ＷＦ面検欠陥２境界３下
    HWFMK3SI As Double              ' 品ＷＦ面検欠陥３サイズ
    HWFMK3MX As Integer             ' 品ＷＦ面検欠陥３上限
    HWFMK3HT As String * 1          ' 品ＷＦ面検欠陥３保証方法＿対
    HWFMK3HS As String * 1          ' 品ＷＦ面検欠陥３保証方法＿処
    HWFMK3KM As String * 1          ' 品ＷＦ面検欠陥３検査頻度＿枚
    HWFMK3KN As String * 1          ' 品ＷＦ面検欠陥３検査頻度＿抜
    HWFMK3KH As String * 1          ' 品ＷＦ面検欠陥３検査頻度＿保
    HWFMK3KU As String * 1          ' 品ＷＦ面検欠陥３検査頻度＿ウ
    HWFM3B1 As Integer              ' 品ＷＦ面検欠陥３境界１
    HWFM3B1B As Integer             ' 品ＷＦ面検欠陥３境界１下
    HWFM3B2 As Integer              ' 品ＷＦ面検欠陥３境界２
    HWFM3B2B As Integer             ' 品ＷＦ面検欠陥３境界２下
    HWFM3B3 As Integer              ' 品ＷＦ面検欠陥３境界３
    HWFM3B3B As Integer             ' 品ＷＦ面検欠陥３境界３下
    HWFMK4SI As Double              ' 品ＷＦ面検欠陥４サイズ
    HWFMK4MX As Integer             ' 品ＷＦ面検欠陥４上限
    HWFMK4HT As String * 1          ' 品ＷＦ面検欠陥４保証方法＿対
    HWFMK4HS As String * 1          ' 品ＷＦ面検欠陥４保証方法＿処
    HWFMK4KM As String * 1          ' 品ＷＦ面検欠陥４検査頻度＿枚
    HWFMK4KN As String * 1          ' 品ＷＦ面検欠陥４検査頻度＿抜
    HWFMK4KH As String * 1          ' 品ＷＦ面検欠陥４検査頻度＿保
    HWFMK4KU As String * 1          ' 品ＷＦ面検欠陥４検査頻度＿ウ
    HWFM4B1 As Integer              ' 品ＷＦ面検欠陥４境界１
    HWFM4B1B As Integer             ' 品ＷＦ面検欠陥４境界１下
    HWFM4B2 As Integer              ' 品ＷＦ面検欠陥４境界２
    HWFM4B2B As Integer             ' 品ＷＦ面検欠陥４境界２下
    HWFM4B3 As Integer              ' 品ＷＦ面検欠陥４境界３
    HWFM4B3B As Integer             ' 品ＷＦ面検欠陥４境界３下
    HWFMB1SI As Double              ' 品ＷＦ面検欠陥裏１サイズ
    HWFMB1MX As Integer             ' 品ＷＦ面検欠陥裏１上限
    HWFMB1SZ As String * 1          ' 品ＷＦ面検欠陥裏１測定条件
    HWFMB1ZA As Integer             ' 品ＷＦ面検欠陥裏１除外領域
    HWFMB1HT As String * 1          ' 品ＷＦ面検欠陥裏１保証方法＿対
    HWFMB1HS As String * 1          ' 品ＷＦ面検欠陥裏１保証方法＿処
    HWFMB1KM As String * 1          ' 品ＷＦ面検欠陥裏１検査頻度＿枚
    HWFMB1KN As String * 1          ' 品ＷＦ面検欠陥裏１検査頻度＿抜
    HWFMB1KH As String * 1          ' 品ＷＦ面検欠陥裏１検査頻度＿保
    HWFMB1KU As String * 1          ' 品ＷＦ面検欠陥裏１検査頻度＿ウ
    HWFMB2SI As Double              ' 品ＷＦ面検欠陥裏２サイズ
    HWFMB2MX As Integer             ' 品ＷＦ面検欠陥裏２上限
    HWFMB2SZ As String * 1          ' 品ＷＦ面検欠陥裏２測定条件
    HWFMB2ZA As Integer             ' 品ＷＦ面検欠陥裏２除外領域
    HWFMB2HT As String * 1          ' 品ＷＦ面検欠陥裏２保証方法＿対
    HWFMB2HS As String * 1          ' 品ＷＦ面検欠陥裏２保証方法＿処
    HWFMB2KM As String * 1          ' 品ＷＦ面検欠陥裏２検査頻度＿枚
    HWFMB2KN As String * 1          ' 品ＷＦ面検欠陥裏２検査頻度＿抜
    HWFMB2KH As String * 1          ' 品ＷＦ面検欠陥裏２検査頻度＿保
    HWFMB2KU As String * 1          ' 品ＷＦ面検欠陥裏２検査頻度＿ウ
    HWFMKSRE As String * 1          ' 品ＷＦ面検欠陥測定器
    HWFMKKW As String * 1           ' 品ＷＦ面検欠陥検査方法
    HWFMPIPT As String * 1          ' 品ＷＦ面検欠陥ＰＩＰ検査
    HWFMPIPK As Integer             ' 品ＷＦ面検欠陥ＰＩＰ個数
    HWFMPISH As String * 1          ' 品ＷＦ面検ＰＩＰ測定位置＿方
    HWFMPIST As String * 1          ' 品ＷＦ面検ＰＩＰ測定位置＿点
    HWFMPISI As String * 1          ' 品ＷＦ面検ＰＩＰ測定位置＿位
    HWFMPIKM As String * 1          ' 品ＷＦ面検ＰＩＰ検査頻度＿枚
    HWFMPIKN As String * 1          ' 品ＷＦ面検ＰＩＰ検査頻度＿抜
    HWFMPIKH As String * 1          ' 品ＷＦ面検ＰＩＰ検査頻度＿保
    HWFMPIKU As String * 1          ' 品ＷＦ面検ＰＩＰ検査頻度＿ウ
    HWFMNMAX As Double              ' 品ＷＦ金属濃度上限
    HWFMNALX As Double              ' 品ＷＦ金属濃度ＡＬ上限
    HWFMNCAX As Double              ' 品ＷＦ金属濃度ＣＡ上限
    HWFMNCRX As Double              ' 品ＷＦ金属濃度ＣＲ上限
    HWFMNCUX As Double              ' 品ＷＦ金属濃度ＣＵ上限
    HWFMNFEX As Double              ' 品ＷＦ金属濃度ＦＥ上限
    HWFMNKMX As Double              ' 品ＷＦ金属濃度Ｋ上限
    HWFMNMGX As Double              ' 品ＷＦ金属濃度ＭＧ上限
    HWFMNNAX As Double              ' 品ＷＦ金属濃度ＮＡ上限
    HWFMNNIX As Double              ' 品ＷＦ金属濃度ＮＩ上限
    HWFMNZNX As Double              ' 品ＷＦ金属濃度ＺＮ上限
    HWFMNKWY As String * 2          ' 品ＷＦ金属濃度検査方法
    HWFMNSPH As String * 1          ' 品ＷＦ金属濃度測定位置＿方
    HWFMNSPT As String * 1          ' 品ＷＦ金属濃度測定位置＿点
    HWFMNSPI As String * 1          ' 品ＷＦ金属濃度測定位置＿位
    HWFMNHWT As String * 1          ' 品ＷＦ金属濃度保証方法＿対
    HWFMNHWS As String * 1          ' 品ＷＦ金属濃度保証方法＿処
    HWFMNKHM As String * 1          ' 品ＷＦ金属濃度検査頻度＿枚
    HWFMNKHN As String * 1          ' 品ＷＦ金属濃度検査頻度＿抜
    HWFMNKHH As String * 1          ' 品ＷＦ金属濃度検査頻度＿保
    HWFMNKHU As String * 1          ' 品ＷＦ金属濃度検査頻度＿ウ
    HWFSPVMX As Double              ' 品ＷＦＳＰＶＦＥ上限
'    HWFSPVMXN As Double              ' 品ＷＦＳＰＶＦＥ上限  6/22 Yam
    HWFSPVKM As String * 1          ' 品ＷＦＳＰＶＦＥ検査頻度＿枚
    HWFSPVKN As String * 1          ' 品ＷＦＳＰＶＦＥ検査頻度＿抜
    HWFSPVKH As String * 1          ' 品ＷＦＳＰＶＦＥ検査頻度＿保
    HWFSPVKU As String * 1          ' 品ＷＦＳＰＶＦＥ検査頻度＿ウ
    HWFSPVSH As String * 1          ' 品ＷＦＳＰＶＦＥ測定位置＿方
    HWFSPVST As String * 1          ' 品ＷＦＳＰＶＦＥ測定位置＿点
    HWFSPVSI As String * 1          ' 品ＷＦＳＰＶＦＥ測定位置＿位
    HWFSPVHT As String * 1          ' 品ＷＦＳＰＶＦＥ保証方法＿対
    HWFSPVHS As String * 1          ' 品ＷＦＳＰＶＦＥ保証方法＿処
    HWFDLMIN As Integer             ' 品ＷＦ拡散長下限
    HWFDLMAX As Integer             ' 品ＷＦ拡散長上限
    HWFDLKHM As String * 1          ' 品ＷＦ拡散長検査頻度＿枚
    HWFDLKHN As String * 1          ' 品ＷＦ拡散長検査頻度＿抜
    HWFDLKHH As String * 1          ' 品ＷＦ拡散長検査頻度＿保
    HWFDLKHU As String * 1          ' 品ＷＦ拡散長検査頻度＿ウ
    HWFDLSPH As String * 1          ' 品ＷＦ拡散長測定位置＿方
    HWFDLSPT As String * 1          ' 品ＷＦ拡散長測定位置＿点
    HWFDLSPI As String * 1          ' 品ＷＦ拡散長測定位置＿位
    HWFDLHWT As String * 1          ' 品ＷＦ拡散長保証方法＿対
    HWFDLHWS As String * 1          ' 品ＷＦ拡散長保証方法＿処
    HWFGKNO1 As String * 6          ' 品ＷＦ外観規格Ｎｏ１
    HWFGKNO2 As String * 6          ' 品ＷＦ外観規格Ｎｏ２
    HWFOTMIN As Double              ' 品ＷＦ酸化膜耐圧下限
    HWFOTMX1 As Double              ' 品ＷＦ酸化膜耐圧上限１
    HWFOTMX2 As Double              ' 品ＷＦ酸化膜耐圧上限２
    HWFOTSPH As String * 1          ' 品ＷＦ酸化膜耐圧測定位置＿方
    HWFOTSPT As String * 1          ' 品ＷＦ酸化膜耐圧測定位置＿点
    HWFOTSPI As String * 1          ' 品ＷＦ酸化膜耐圧測定位置＿位
    HWFOTHWT As String * 1          ' 品ＷＦ酸化膜耐圧保証方法＿対
    HWFOTHWS As String * 1          ' 品ＷＦ酸化膜耐圧保証方法＿処
    HWFOTKWY As String * 2          ' 品ＷＦ酸化膜耐圧検査方法
    HWFOTKW1 As String * 2          ' 品ＷＦ酸化膜耐圧検査方法１
    HWFOTKW2 As String * 2          ' 品ＷＦ酸化膜耐圧検査方法２
    HWFOTKHM As String * 1          ' 品ＷＦ酸化膜耐圧検査頻度＿枚
    HWFOTKHN As String * 1          ' 品ＷＦ酸化膜耐圧検査頻度＿抜
    HWFOTKHH As String * 1          ' 品ＷＦ酸化膜耐圧検査頻度＿保
    HWFOTKHU As String * 1          ' 品ＷＦ酸化膜耐圧検査頻度＿ウ
    HWFTSPHM As String * 1          ' 品ＷＦトレスサンプル頻度＿枚
    HWFTSPHN As String * 1          ' 品ＷＦトレスサンプル頻度＿抜
    HWFTSPHH As String * 1          ' 品ＷＦトレスサンプル頻度＿保
    HWFTSPHU As String * 1          ' 品ＷＦトレスサンプル頻度＿ウ
    HWFLTDCX As Double              ' 品ＷＦＬＴＤ濃度ＣＵ上限
    HWFLTDIN As String * 2          ' 品ＷＦＬＴＤ濃度指数
    HWFLTDKW As String * 2          ' 品ＷＦＬＴＤ濃度検査方法
    HWFLTDSH As String * 1          ' 品ＷＦＬＴＤ濃度測定位置＿方
    HWFLTDST As String * 1          ' 品ＷＦＬＴＤ濃度測定位置＿点
    HWFLTDSI As String * 1          ' 品ＷＦＬＴＤ濃度測定位置＿位
    HWFLTDHT As String * 1          ' 品ＷＦＬＴＤ濃度保証方法＿対
    HWFLTDHS As String * 1          ' 品ＷＦＬＴＤ濃度保証方法＿処
    HWFLTDKM As String * 1          ' 品ＷＦＬＴＤ濃度検査頻度＿枚
    HWFLTDKN As String * 1          ' 品ＷＦＬＴＤ濃度検査頻度＿抜
    HWFLTDKH As String * 1          ' 品ＷＦＬＴＤ濃度検査頻度＿保
    HWFLTDKU As String * 1          ' 品ＷＦＬＴＤ濃度検査頻度＿ウ
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
' 追加 2003/09.11 SystemBrain Start
    HWFSPVAM As Double              ' 品WFSPVFE平均
'    HWFSPVAMN As Double              ' 品WFSPVFE平均 6/22
    HWFMK1MC As String * 1          ' 品WF面検欠陥1面指定
    HWFMK2MC As String * 1          ' 品WF面検欠陥2面指定
    HWFMK3MC As String * 1          ' 品WF面検欠陥3面指定
    HWFMK4MC As String * 1          ' 品WF面検欠陥4面指定
    HWFMK5MC As String * 1          ' 品WF面検欠陥5面指定
    HWFMK6MC As String * 1          ' 品WF面検欠陥6面指定
    HWFMK2SZ As String * 1          ' 品WF面検欠陥2測定条件
    HWFMK3SZ As String * 1          ' 品WF面検欠陥3測定条件
    HWFMK4SZ As String * 1          ' 品WF面検欠陥4測定条件
    HWFMK2ZAR As Integer            ' 品WF面検欠陥2除外領域
    HWFMK3ZAR As Integer            ' 品WF面検欠陥3除外領域
    HWFMK4ZAR As Integer            ' 品WF面検欠陥4除外領域
    HWFMK5B1 As Integer             ' 品WF面検欠陥5境界1
    HWFMK5B1B As Integer            ' 品WF面検欠陥5境界1下
    HWFMK5B2 As Integer             ' 品WF面検欠陥5境界2
    HWFMK5B2B As Integer            ' 品WF面検欠陥5境界2下
    HWFMK5B3 As Integer             ' 品WF面検欠陥5境界3
    HWFMK5B3B As Integer            ' 品WF面検欠陥5境界3下
    HWFMK6B1 As Integer             ' 品WF面検欠陥6境界1
    HWFMK6B1B As Integer            ' 品WF面検欠陥6境界1下
    HWFMK6B2 As Integer             ' 品WF面検欠陥6境界2
    HWFMK6B2B As Integer            ' 品WF面検欠陥6境界2下
    HWFMK6B3 As Integer             ' 品WF面検欠陥6境界3
    HWFMK6B3B As Integer            ' 品WF面検欠陥6境界3下
' 追加 2003/09.11 SystemBrain End
' 追加 2005/06/16 ffc)tanabe start
    HWFMK7MC    As String * 1       '品ＷＦ面検欠陥７面指定
    HWFMK7SI    As Double           '品ＷＦ面検欠陥７サイズ
    HWFMK7MX    As Integer          '品ＷＦ面検欠陥７上限
    HWFMK7SZ    As String * 1       '品ＷＦ面検欠陥７測定条件
    HWFMK7ZA    As Integer          '品ＷＦ面検欠陥７除外領域
    HWFMK7HT    As String * 1       '品ＷＦ面検欠陥７保証方法＿対
    HWFMK7HS    As String * 1       '品ＷＦ面検欠陥７保証方法＿処
    HWFMK8MC    As String * 1       '品ＷＦ面検欠陥８面指定
    HWFMK8SI    As Double           '品ＷＦ面検欠陥８サイズ
    HWFMK8MX    As Integer          '品ＷＦ面検欠陥８上限
    HWFMK8SZ    As String * 1       '品ＷＦ面検欠陥８測定条件
    HWFMK8ZA    As Integer          '品ＷＦ面検欠陥８除外領域
    HWFMK8HT    As String * 1       '品ＷＦ面検欠陥８保証方法＿対
    HWFMK8HS    As String * 1       '品ＷＦ面検欠陥８保証方法＿処
    HWFMK9MC    As String * 1       '品ＷＦ面検欠陥９面指定
    HWFMK9SI    As Double           '品ＷＦ面検欠陥９サイズ
    HWFMK9MX    As Integer          '品ＷＦ面検欠陥９上限
    HWFMK9SZ    As String * 1       '品ＷＦ面検欠陥９測定条件
    HWFMK9ZA    As Integer          '品ＷＦ面検欠陥９除外領域
    HWFMK9HT    As String * 1       '品ＷＦ面検欠陥９保証方法＿対
    HWFMK9HS    As String * 1       '品ＷＦ面検欠陥９保証方法＿処
    HWFMK10MC   As String * 1       '品ＷＦ面検欠陥１０面指定
    HWFMK10SI   As Double           '品ＷＦ面検欠陥１０サイズ
    HWFMK10MX   As Integer          '品ＷＦ面検欠陥１０上限
    HWFMK10SZ   As String * 1       '品ＷＦ面検欠陥１０測定条件
    HWFMK10ZA   As Integer          '品ＷＦ面検欠陥１０除外領域
    HWFMK10HT   As String * 1       '品ＷＦ面検欠陥１０保証方法＿対
    HWFMK10HS   As String * 1       '品ＷＦ面検欠陥１０保証方法＿処
    HWFMK11MC   As String * 1       '品ＷＦ面検欠陥１１面指定
    HWFMK11SI   As Double           '品ＷＦ面検欠陥１１サイズ
    HWFMK11MX   As Integer          '品ＷＦ面検欠陥１１上限
    HWFMK11SZ   As String * 1       '品ＷＦ面検欠陥１１測定条件
    HWFMK11ZA   As Integer          '品ＷＦ面検欠陥１１除外領域
    HWFMK11HT   As String * 1       '品ＷＦ面検欠陥１１保証方法＿対
    HWFMK11HS   As String * 1       '品ＷＦ面検欠陥１１保証方法＿処
    HWFMK12MC   As String * 1       '品ＷＦ面検欠陥１２面指定
    HWFMK12SI   As Double           '品ＷＦ面検欠陥１２サイズ
    HWFMK12MX   As Integer          '品ＷＦ面検欠陥１２上限
    HWFMK12SZ   As String * 1       '品ＷＦ面検欠陥１２測定条件
    HWFMK12ZA   As Integer          '品ＷＦ面検欠陥１２除外領域
    HWFMK12HT   As String * 1       '品ＷＦ面検欠陥１２保証方法＿対
    HWFMK12HS   As String * 1       '品ＷＦ面検欠陥１２保証方法＿処
    HWFMK13MC   As String * 1       '品ＷＦ面検欠陥１３面指定
    HWFMK13SI   As Double           '品ＷＦ面検欠陥１３サイズ
    HWFMK13MX   As Integer          '品ＷＦ面検欠陥１３上限
    HWFMK13SZ   As String * 1       '品ＷＦ面検欠陥１３測定条件
    HWFMK13ZA   As Integer          '品ＷＦ面検欠陥１３除外領域
    HWFMK13HT   As String * 1       '品ＷＦ面検欠陥１３保証方法＿対
    HWFMK13HS   As String * 1       '品ＷＦ面検欠陥１３保証方法＿処
    HWFMK14MC   As String * 1       '品ＷＦ面検欠陥１４面指定
    HWFMK14SI   As Double           '品ＷＦ面検欠陥１４サイズ
    HWFMK14MX   As Integer          '品ＷＦ面検欠陥１４上限
    HWFMK14SZ   As String * 1       '品ＷＦ面検欠陥１４測定条件
    HWFMK14ZA   As Integer          '品ＷＦ面検欠陥１４除外領域
    HWFMK14HT   As String * 1       '品ＷＦ面検欠陥１４保証方法＿対
    HWFMK14HS   As String * 1       '品ＷＦ面検欠陥１４保証方法＿処
    HWFMK15MC   As String * 1       '品ＷＦ面検欠陥１５面指定
    HWFMK15SI   As Double           '品ＷＦ面検欠陥１５サイズ
    HWFMK15MX   As Integer          '品ＷＦ面検欠陥１５上限
    HWFMK15SZ   As String * 1       '品ＷＦ面検欠陥１５測定条件
    HWFMK15ZA   As Integer          '品ＷＦ面検欠陥１５除外領域
    HWFMK15HT   As String * 1       '品ＷＦ面検欠陥１５保証方法＿対
    HWFMK15HS   As String * 1       '品ＷＦ面検欠陥１５保証方法＿処
    HWFSPVMXN   As Double           '品ＷＦＳＰＶＦＥ上限
    HWFSPVAMN   As Double           '品ＷＦＳＰＶＦＥ平均
' 追加 2005/06/16 ffc)tanabe end
End Type


' 製品仕様WFﾃﾞｰﾀ９
Public Type typ_TBCME029
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    HMGSTRRNO As String * 9         ' 品管理仕様登録依頼番号
    HMGSTFNO As String * 8          ' 品管理社員Ｎｏ
    HMGWFSNO As String * 6          ' 品管理ＷＦ製品番号
    HMGWFSNE As Integer             ' 品管理ＷＦ製品番号枝番
    HWFOF1AX As Double              ' 品ＷＦＯＳＦ１平均上限
    HWFOF1MX As Double              ' 品ＷＦＯＳＦ１上限
    HWFOF1ET As Integer             ' 品ＷＦＯＳＦ１選択ＥＴ代
    HWFOF1NS As String * 2          ' 品ＷＦＯＳＦ１熱処理法
    HWFOF1SZ As String * 1          ' 品ＷＦＯＳＦ１測定条件
    HWFOF1SH As String * 1          ' 品ＷＦＯＳＦ１測定位置＿方
    HWFOF1ST As String * 1          ' 品ＷＦＯＳＦ１測定位置＿点
    HWFOF1SR As String * 1          ' 品ＷＦＯＳＦ１測定位置＿領
    HWFOF1HT As String * 1          ' 品ＷＦＯＳＦ１保証方法＿対
    HWFOF1HS As String * 1          ' 品ＷＦＯＳＦ１保証方法＿処
    HWFOF1KM As String * 1          ' 品ＷＦＯＳＦ１検査頻度＿枚
    HWFOF1KN As String * 1          ' 品ＷＦＯＳＦ１検査頻度＿抜
    HWFOF1KH As String * 1          ' 品ＷＦＯＳＦ１検査頻度＿保
    HWFOF1KU As String * 1          ' 品ＷＦＯＳＦ１検査頻度＿ウ
    HWFOF2AX As Double              ' 品ＷＦＯＳＦ２平均上限
    HWFOF2MX As Double              ' 品ＷＦＯＳＦ２上限
    HWFOF2ET As Integer             ' 品ＷＦＯＳＦ２選択ＥＴ代
    HWFOF2NS As String * 2          ' 品ＷＦＯＳＦ２熱処理法
    HWFOF2SZ As String * 1          ' 品ＷＦＯＳＦ２測定条件
    HWFOF2SH As String * 1          ' 品ＷＦＯＳＦ２測定位置＿方
    HWFOF2ST As String * 1          ' 品ＷＦＯＳＦ２測定位置＿点
    HWFOF2SR As String * 1          ' 品ＷＦＯＳＦ２測定位置＿領
    HWFOF2HT As String * 1          ' 品ＷＦＯＳＦ２保証方法＿対
    HWFOF2HS As String * 1          ' 品ＷＦＯＳＦ２保証方法＿処
    HWFOF2KM As String * 1          ' 品ＷＦＯＳＦ２検査頻度＿枚
    HWFOF2KN As String * 1          ' 品ＷＦＯＳＦ２検査頻度＿抜
    HWFOF2KH As String * 1          ' 品ＷＦＯＳＦ２検査頻度＿保
    HWFOF2KU As String * 1          ' 品ＷＦＯＳＦ２検査頻度＿ウ
    HWFOF3AX As Double              ' 品ＷＦＯＳＦ３平均上限
    HWFOF3MX As Double              ' 品ＷＦＯＳＦ３上限
    HWFOF3ET As Integer             ' 品ＷＦＯＳＦ３選択ＥＴ代
    HWFOF3NS As String * 2          ' 品ＷＦＯＳＦ３熱処理法
    HWFOF3SZ As String * 1          ' 品ＷＦＯＳＦ３測定条件
    HWFOF3SH As String * 1          ' 品ＷＦＯＳＦ３測定位置＿方
    HWFOF3ST As String * 1          ' 品ＷＦＯＳＦ３測定位置＿点
    HWFOF3SR As String * 1          ' 品ＷＦＯＳＦ３測定位置＿領
    HWFOF3HT As String * 1          ' 品ＷＦＯＳＦ３保証方法＿対
    HWFOF3HS As String * 1          ' 品ＷＦＯＳＦ３保証方法＿処
    HWFOF3KM As String * 1          ' 品ＷＦＯＳＦ３検査頻度＿枚
    HWFOF3KN As String * 1          ' 品ＷＦＯＳＦ３検査頻度＿抜
    HWFOF3KH As String * 1          ' 品ＷＦＯＳＦ３検査頻度＿保
    HWFOF3KU As String * 1          ' 品ＷＦＯＳＦ３検査頻度＿ウ
    HWFOF4AX As Double              ' 品ＷＦＯＳＦ４平均上限
    HWFOF4MX As Double              ' 品ＷＦＯＳＦ４上限
    HWFOF4ET As Integer             ' 品ＷＦＯＳＦ４選択ＥＴ代
    HWFOF4NS As String * 2          ' 品ＷＦＯＳＦ４熱処理法
    HWFOF4SZ As String * 1          ' 品ＷＦＯＳＦ４測定条件
    HWFOF4SH As String * 1          ' 品ＷＦＯＳＦ４測定位置＿方
    HWFOF4ST As String * 1          ' 品ＷＦＯＳＦ４測定位置＿点
    HWFOF4SR As String * 1          ' 品ＷＦＯＳＦ４測定位置＿領
    HWFOF4HT As String * 1          ' 品ＷＦＯＳＦ４保証方法＿対
    HWFOF4HS As String * 1          ' 品ＷＦＯＳＦ４保証方法＿処
    HWFOF4KM As String * 1          ' 品ＷＦＯＳＦ４検査頻度＿枚
    HWFOF4KN As String * 1          ' 品ＷＦＯＳＦ４検査頻度＿抜
    HWFOF4KH As String * 1          ' 品ＷＦＯＳＦ４検査頻度＿保
    HWFOF4KU As String * 1          ' 品ＷＦＯＳＦ４検査頻度＿ウ
    HWFBM1AN As Double              ' 品ＷＦＢＭＤ１平均下限
    HWFBM1AX As Double              ' 品ＷＦＢＭＤ１平均上限
    HWFBM1ET As Integer             ' 品ＷＦＢＭＤ１選択ＥＴ代
    HWFBM1NS As String * 2          ' 品ＷＦＢＭＤ１熱処理法
    HWFBM1SZ As String * 1          ' 品ＷＦＢＭＤ１測定条件
    HWFBM1SH As String * 1          ' 品ＷＦＢＭＤ１測定位置＿方
    HWFBM1ST As String * 1          ' 品ＷＦＢＭＤ１測定位置＿点
    HWFBM1SR As String * 1          ' 品ＷＦＢＭＤ１測定位置＿領
    HWFBM1HT As String * 1          ' 品ＷＦＢＭＤ１保証方法＿対
    HWFBM1HS As String * 1          ' 品ＷＦＢＭＤ１保証方法＿処
    HWFBM1KM As String * 1          ' 品ＷＦＢＭＤ１検査頻度＿枚
    HWFBM1KN As String * 1          ' 品ＷＦＢＭＤ１検査頻度＿抜
    HWFBM1KH As String * 1          ' 品ＷＦＢＭＤ１検査頻度＿保
    HWFBM1KU As String * 1          ' 品ＷＦＢＭＤ１検査頻度＿ウ
    HWFBM2AN As Double              ' 品ＷＦＢＭＤ２平均下限
    HWFBM2AX As Double              ' 品ＷＦＢＭＤ２平均上限
    HWFBM2ET As Integer             ' 品ＷＦＢＭＤ２選択ＥＴ代
    HWFBM2NS As String * 2          ' 品ＷＦＢＭＤ２熱処理法
    HWFBM2SZ As String * 1          ' 品ＷＦＢＭＤ２測定条件
    HWFBM2SH As String * 1          ' 品ＷＦＢＭＤ２測定位置＿方
    HWFBM2ST As String * 1          ' 品ＷＦＢＭＤ２測定位置＿点
    HWFBM2SR As String * 1          ' 品ＷＦＢＭＤ２測定位置＿領
    HWFBM2HT As String * 1          ' 品ＷＦＢＭＤ２保証方法＿対
    HWFBM2HS As String * 1          ' 品ＷＦＢＭＤ２保証方法＿処
    HWFBM2KM As String * 1          ' 品ＷＦＢＭＤ２検査頻度＿枚
    HWFBM2KN As String * 1          ' 品ＷＦＢＭＤ２検査頻度＿抜
    HWFBM2KH As String * 1          ' 品ＷＦＢＭＤ２検査頻度＿保
    HWFBM2KU As String * 1          ' 品ＷＦＢＭＤ２検査頻度＿ウ
    HWFBM3AN As Double              ' 品ＷＦＢＭＤ３平均下限
    HWFBM3AX As Double              ' 品ＷＦＢＭＤ３平均上限
    HWFBM3ET As Integer             ' 品ＷＦＢＭＤ３選択ＥＴ代
    HWFBM3NS As String * 2          ' 品ＷＦＢＭＤ３熱処理法
    HWFBM3SZ As String * 1          ' 品ＷＦＢＭＤ３測定条件
    HWFBM3SH As String * 1          ' 品ＷＦＢＭＤ３測定位置＿方
    HWFBM3ST As String * 1          ' 品ＷＦＢＭＤ３測定位置＿点
    HWFBM3SR As String * 1          ' 品ＷＦＢＭＤ３測定位置＿領
    HWFBM3HT As String * 1          ' 品ＷＦＢＭＤ３保証方法＿対
    HWFBM3HS As String * 1          ' 品ＷＦＢＭＤ３保証方法＿処
    HWFBM3KM As String * 1          ' 品ＷＦＢＭＤ３検査頻度＿枚
    HWFBM3KN As String * 1          ' 品ＷＦＢＭＤ３検査頻度＿抜
    HWFBM3KH As String * 1          ' 品ＷＦＢＭＤ３検査頻度＿保
    HWFBM3KU As String * 1          ' 品ＷＦＢＭＤ３検査頻度＿ウ
    HWFOSPAX As Integer             ' 品ＷＦＯＳＰ平均上限
    HWFOSPMX As Integer             ' 品ＷＦＯＳＰ上限
    HWFOSPSH As String * 1          ' 品ＷＦＯＳＰ測定位置＿方
    HWFOSPST As String * 1          ' 品ＷＦＯＳＰ測定位置＿点
    HWFOSPSR As String * 1          ' 品ＷＦＯＳＰ測定位置＿領
    HWFOSPHT As String * 1          ' 品ＷＦＯＳＰ保証方法＿対
    HWFOSPHS As String * 1          ' 品ＷＦＯＳＰ保証方法＿処
    HWFOSPNS As String * 2          ' 品ＷＦＯＳＰ熱処理法
    HWFOSPSZ As String * 1          ' 品ＷＦＯＳＰ測定条件
    HWFOSPKM As String * 1          ' 品ＷＦＯＳＰ検査頻度＿枚
    HWFOSPKN As String * 1          ' 品ＷＦＯＳＰ検査頻度＿抜
    HWFOSPKH As String * 1          ' 品ＷＦＯＳＰ検査頻度＿保
    HWFOSPKU As String * 1          ' 品ＷＦＯＳＰ検査頻度＿ウ
    HWFOSPET As Integer             ' 品ＷＦＯＳＰ選択ＥＴ代
    HWFNOTE As String               ' 品ＷＦ特記
    HWFRS1N As String               ' 品ＷＦ予備１＿内
    HWFRS1Y As String               ' 品ＷＦ予備１＿用
    HWFRS2N As String               ' 品ＷＦ予備２＿内
    HWFRS2Y As String               ' 品ＷＦ予備２＿用
    HWFRS3N As String               ' 品ＷＦ予備３＿内
    HWFRS3Y As String               ' 品ＷＦ予備３＿用
    HWFRS4N As String               ' 品ＷＦ予備４＿内
    HWFRS4Y As String               ' 品ＷＦ予備４＿用
    HWFRS5N As String               ' 品ＷＦ予備５＿内
    HWFRS5Y As String               ' 品ＷＦ予備５＿用
    HWFRS6N As String               ' 品ＷＦ予備６＿内
    HWFRS6Y As String               ' 品ＷＦ予備６＿用
    HWFRS7N As String               ' 品ＷＦ予備７＿内
    HWFRS7Y As String               ' 品ＷＦ予備７＿用
    HWFRS8N As String               ' 品ＷＦ予備８＿内
    HWFRS8Y As String               ' 品ＷＦ予備８＿用
    HWFRS9N As String               ' 品ＷＦ予備９＿内
    HWFRS9Y As String               ' 品ＷＦ予備９＿用
    HWFRS10N As String              ' 品ＷＦ予備１０＿内
    HWFRS10Y As String              ' 品ＷＦ予備１０＿用
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
' 追加 2003/09.11 SystemBrain Start
    HWFOSF1PTK As String * 1        ' 品WFOSF1パタン区分
    HWFOSF2PTK As String * 1        ' 品WFOSF2パタン区分
    HWFOSF3PTK As String * 1        ' 品WFOSF3パタン区分
    HWFOSF4PTK As String * 1        ' 品WFOSF4パタン区分
    HWFBM1MBP As Double             ' 品WFBMD1面内分布
    HWFBM2MBP As Double             ' 品WFBMD2面内分布
    HWFBM3MBP As Double             ' 品WFBMD3面内分布
    HWFBM1MCL As String * 2         ' 品WFBMD1面内計算
    HWFBM2MCL As String * 2         ' 品WFBMD2面内計算
    HWFBM3MCL As String * 2         ' 品WFBMD3面内計算
' 追加 2003/09.11 SystemBrain End
End Type


' SXL製作条件ﾃﾞｰﾀ
Public Type typ_TBCME030
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    SSXLIFTW As String * 2          ' 製ＳＸ引上方法
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' ＳＸＬ製作条件付与取消
Public Type typ_TBCME031
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    STAFFNO As String * 8           ' 社員№
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 使用開始
Public Type typ_TBCME032
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    STAFFNO As String * 8           ' 社員№
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 汎用ｺｰﾄﾞﾏｽﾀ
Public Type typ_TBCME033
    codeNo As String * 12           ' コードＮＯ
    CODE As String * 5              ' コード
    codeCont As String              ' コード内容
    INDORDER As Long                ' 表示順
    codename As String              ' コード名称
    KUBUN As String                 ' 区分
    READTIME As Double              ' リードタイム
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 比抵抗補正計算
Public Type typ_TBCME034
    RESIHCAL As String * 2          ' 比抵抗補正計算
    RESIHINA As Double              ' 比抵抗補正係数Ａ
    RESIHINB As Double              ' 比抵抗補正係数Ｂ
    CSGROUP As String * 3           ' 顧客グループ
    CSCODE As String * 8            ' 顧客コード
    CSNAME As String                ' 顧客名
    NOTE As String                  ' 特記
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 酸素濃度補正計算
Public Type typ_TBCME035
    OXYNHCAL As String * 2          ' 酸素濃度補正計算
    OXYNHINA As Double              ' 酸素濃度補正係数Ａ
    OXYNHINB As Double              ' 酸素濃度補正係数Ｂ
    CSGROUP As String * 3           ' 顧客グループ
    CSCODE As String * 8            ' 顧客コード
    CSNAME As String                ' 顧客名
    NOTE As String                  ' 特記
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 結晶内側管理
Public Type typ_TBCME036
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    EPDSETCH As String * 1          ' EPD　選択エッチ
    EPDUP As Integer                ' EPD　上限
    CUTUNIT As Integer              ' カット単位
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
' 払出規制項目追加対応 yakimura 2002.12.01 start
    TOPREG As Integer               ' TOP規制
    TAILREG As Double               ' TAIL規制
    BTMSPRT As Integer              ' ボトム析出規制
' 払出規制項目追加対応 yakimura 2002.12.01 end
' 追加 2003/09.11 SystemBrain Start
    OTHER1 As String * 1            '
    OTHER2 As String * 1            '
    OTHERTIME As Date               '
    DCHYUUBU As String * 1          ' ドローチューブ
    KSTAFFID As String * 8          ' 更新社員ID
    sStaffID As String * 8          ' 承認社員ID
    SYNFLAG As String * 1           ' 承認フラグ
    SYNDATE As Date                 ' 承認日付
    SNOTE As String * 255           ' 製品仕様特記
    JNOTE As String * 255           ' 製作条件特記
    BLOCKHFLAG As String * 1        ' ブロック単位保証品番フラグ
' 追加 2003/09.11 SystemBrain End
' WFカット単位機能追加 2005/04/12 ffc)tanabe start
    WFCUTUNIT As String * 4         'WFカット単位
' WFカット単位機能追加 2005/04/12 ffc)tanabe end
'*** UPDATE ↓ Y.SIMIZU 2005/10/1 GDﾗｲﾝ数
    HSXGDLINE   As Single           '品SXGDﾗｲﾝ数
    HWFGDLINE   As Single           '品WFGDﾗｲﾝ数
'*** UPDATE ↑ Y.SIMIZU 2005/10/1 GDﾗｲﾝ数

'--------------- 2008/08/25 INSERT START  By Systech ---------------
    HSXDKTMP    As String * 1       '品SXDK温度
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    HSXLDLRMN   As Integer          ' 品SXL/DL連続0下限
    HSXLDLRMX   As Integer          ' 品SXL/DL連続0上限
    HWFLDLRMN   As Integer          ' 品WFL/DL連続0下限
    HWFLDLRMX   As Integer          ' 品WFL/DL連続0上限
    HSXOF1ARPTK As String * 1       ' 品SXOSF1(ArAN)パタン区分
    HSXOFARMIN  As Double           ' 品SXOSF(ArAN)下限
    HSXOFARMAX  As Double           ' 品SXOSF(ArAN)上限
    HSXOFARMHMX As Double           ' 品SXOSF(ArAN)面内比上限
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    'Add Start 2011/01/27 SMPK Miyata
    HSXCJLTBND  As Integer          ' 品SXL/CJLTバンド幅
    'Add End   2011/01/27 SMPK Miyata

End Type


' ブロック新規情報（抜試指示付）
Public Type typ_TBCMY001
    BLOCKID As String * 12          ' ブロックID
    BLOCKLEN As String * 3          ' ブロックの長さ
    MAINHINBAN As String * 10       ' 代表品番
    PNTYPE As String * 1            ' タイプ
    ROUP As String * 8              ' 比抵抗上限値
    ROLOW As String * 8             ' 比抵抗下限値
    OIUP As String * 5              ' 酸素濃度上限値
    OILOW As String * 5             ' 酸素濃度下限値
    TANMEN As String * 3            ' 端面角度
    WARPRANK As String * 1          ' ワープランク
    CRYSTALMEN As String * 3        ' 結晶面
    SLPCEN As String * 4            ' 傾中心
    SLPLOW As String * 5            ' 傾下限
    SLPUP As String * 5             ' 傾上限
    INSPMETH As String * 2          ' 検査方法
    INSPFREQ As String * 4          ' 検査頻度
    SLPDRC As String * 2            ' 傾方位
    SLPDRCAPP As String * 1         ' 傾方位指定
    SLPHEIDRC As String * 2         ' 傾縦方位
    SLPHEICEN As String * 5         ' 傾縦中心
    SLPHEILOW As String * 5         ' 傾縦下限
    SLPHEIUP As String * 5          ' 傾縦上限
    SLPWIDDRC As String * 2         ' 傾横方位
    SLPWIDCEN As String * 5         ' 傾横中心
    SLPWIDLOW As String * 5         ' 傾横下限
    SLPWIDUP As String * 5          ' 傾横上限
    SEED As String * 1              ' 引上時使用したシ－ド傾き
    TXID As String * 6              ' トランザクションID
    sBlockId As String * 12         ' 先頭ブロックID
    BLOCKORDER As Integer           ' ブロック順序
    REGDATE As Date                 ' 登録日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
'2007/07/17 UPDATE_STR マルチブロック対応 SHINDOH
    HINCNT As Integer               ' 構成品番数
    MULUTIHINBAN1 As String * 10    ' 構成品番その１品番
    TOPICHI1 As Integer             ' 構成品番その１Top位置(mm)
    TAILICHI1 As Integer            ' 構成品番その１Tail位置(mm)
    HINBANLEN1 As Integer           ' 構成品番その１長さ(mm)
    MULUTIHINBAN2 As String * 10    ' 構成品番その２品番
    TOPICHI2 As Integer             ' 構成品番その２Top位置(mm)
    TAILICHI2 As Integer            ' 構成品番その２Tail位置(mm)
    HINBANLEN2 As Integer           ' 構成品番その２長さ(mm)
    MULUTIHINBAN3 As String * 10    ' 構成品番その３品番
    TOPICHI3 As Integer             ' 構成品番その３Top位置(mm)
    TAILICHI3 As Integer            ' 構成品番その３Tail位置(mm)
    HINBANLEN3 As Integer           ' 構成品番その３長さ(mm)
    MULUTIHINBAN4 As String * 10    ' 構成品番その４品番
    TOPICHI4 As Integer             ' 構成品番その４Top位置(mm)
    TAILICHI4 As Integer            ' 構成品番その４Tail位置(mm)
    HINBANLEN4 As Integer           ' 構成品番その４長さ(mm)
    MULUTIHINBAN5 As String * 10    ' 構成品番その５品番
    TOPICHI5 As Integer             ' 構成品番その５Top位置(mm)
    TAILICHI5 As Integer            ' 構成品番その５Tail位置(mm)
    HINBANLEN5 As Integer           ' 構成品番その５長さ(mm)
'2007/07/17 UPDATE_END マルチブロック対応 SHINDOH
End Type


' ブロック新規情報返答
Public Type typ_TBCMY002
    BLOCKID As String * 12          ' ブロックID
    RET As String * 6               ' リターンコード
    TXID As String * 6              ' トランザクションID
    TXIDRET As String * 6           ' トランザクションID リターンコード
    BLKIDRET As String * 6          ' ブロックIDのリターンコード
    REGDATE As Date                 ' 登録日付
    CHECKFLG As String * 1          ' チェックフラグ
End Type


' 測定評価方法指示
Public Type typ_TBCMY003
    SAMPLEID As String * 16         ' サンプルID
    OSITEM As String * 4            ' 評価項目
    TRANCNT As Integer              ' 処理回数
    SAMPLEKB As String * 1          ' サンプル区分
    MAISU As String * 1             ' 評価枚数
    Spec As String * 10             ' 規格値
    NETSU As String * 2             ' 熱処理条件
    ET As String * 3                ' エッチング条件
    MES As String * 3               ' 計測方法
    DKAN As String * 10             ' ＤＫアニール条件
    TXID As String * 6              ' トランザクションID
    REGDATE As Date                 ' 登録日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
    FEPUA       As String * 10      ' SPV_Fe_PUA値 (number(5,2))    06/06/08 ooba START ======>
    FEPUAPCT    As String * 10      ' SPV_Fe_PUA％値 (number(6,3))
    FESTD       As String * 10      ' SPV_Fe_STD (number(6,3))
    DIFFPUA     As String * 10      ' SPV_拡散長_PUA値 (number(5,1))
    DIFFPUAPCT  As String * 10      ' SPV_拡散長_PUA％値 (number(6,3))
    NRPUA       As String * 10      ' SPV_NR_PUA値 (number(5,2))
    NRPUAPCT    As String * 10      ' SPV_NR_PUA%値 (number(6,3))
    NRSTD       As String * 10      ' SPV_NR_STD (number(6,3))      06/06/08 ooba END ========>
    MUKESAKI As String              ' 07/09/05 SPK Tsutsumi Add
End Type

' エピ測定評価方法指示  2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
Public Type typ_TBCMY020
    SAMPLEID As String * 16         ' サンプルID
    OSITEM As String * 4            ' 評価項目
    TRANCNT As Integer              ' 処理回数
    SAMPLEKB As String * 1          ' サンプル区分
    MAISU As String * 1             ' 評価枚数
    Spec As String * 10             ' 規格値
    NETSU As String * 2             ' 熱処理条件
    ET As String * 3                ' エッチング条件
    MES As String * 3               ' 計測方法
    DKAN As String * 10             ' ＤＫアニール条件
    TXID As String * 6              ' トランザクションID
    REGDATE As Date                 ' 登録日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
    MUKESAKI As String              ' 07/09/05 SPK Tsutsumi Add
End Type


' 測定評価方法指示返答
Public Type typ_TBCMY004
    SAMPLEID As String * 16         ' サンプルID
    TRANCNT As Integer              ' 処理回数
    TXID As String * 6              ' トランザクションID
    RET As String * 6               ' リターンコード
    REGDATE As Date                 ' 登録日付
    CHECKFLG As String * 1          ' チェックフラグ
End Type


' ブロック変更情報
Public Type typ_TBCMY005
    BLOCKID As String * 12          ' ブロックID
    TRANCNT As Integer              ' 処理回数
    DELFLG As String * 1            ' 削除指示
    BLOCKLEN As String * 3          ' ブロックの長さ
    MAINHINBAN As String * 10       ' 代表品番
    PNTYPE As String * 1            ' タイプ
    ROUP As String * 8              ' 比抵抗上限値
    ROLOW As String * 8             ' 比抵抗下限値
    OIUP As String * 5              ' 酸素濃度上限値
    OILOW As String * 5             ' 酸素濃度下限値
    TANMEN As String * 3            ' 端面角度
    WARPRANK As String * 1          ' ワープランク
    CRYSTALMEN As String * 3        ' 結晶面
    SLPCEN As String * 4            ' 傾中心
    SLPLOW As String * 5            ' 傾下限
    SLPUP As String * 5             ' 傾上限
    INSPMETH As String * 2          ' 検査方法
    INSPFREQ As String * 4          ' 検査頻度
    SLPDRC As String * 2            ' 傾方位
    SLPDRCAPP As String * 1         ' 傾方位指定
    SLPHEIDRC As String * 2         ' 傾縦方位
    SLPHEICEN As String * 5         ' 傾縦中心
    SLPHEILOW As String * 5         ' 傾縦下限
    SLPHEIUP As String * 5          ' 傾縦上限
    SLPWIDDRC As String * 2         ' 傾横方位
    SLPWIDCEN As String * 5         ' 傾横中心
    SLPWIDLOW As String * 5         ' 傾横下限
    SLPWIDUP As String * 5          ' 傾横上限
    SEED As String * 1              ' 引上時使用したシ－ド傾き
    TXID As String * 6              ' トランザクションID
    REGDATE As Date                 ' 登録日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' ブロック変更情報返答
Public Type typ_TBCMY006
    BLOCKID As String * 12          ' ブロックID
    TRANCNT As Integer              ' 処理回数
    RET As String * 6               ' リターンコード
    TXID As String * 6              ' トランザクションID
    TXIDRET As String * 6           ' トランザクションID リターンコード
    BLKIDRET As String * 6          ' ブロックIDのリターンコード
    REGDATE As Date                 ' 登録日付
    CHECKFLG As String * 1          ' チェックフラグ
End Type


' Ｓｘｌ確定指示
Public Type typ_TBCMY007
    SXL_ID As String * 13           ' SXL-ID
    SAMPLE_FROM As String * 16      ' サンプルID (From)
    SAMPLE_TO As String * 16        ' サンプルID (To)
    BLOCKID As String * 12          ' ブロックＩＤ
    hinban As String * 10           ' 確定品番
    KUBUN As String * 2             ' 区分コード
    TXID As String * 6              ' トランザクションID
    REGDATE As Date                 ' 登録日付
    SUMMITSENDFLAG As String * 1    ' SUMMIT送信フラグ
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
    MESDATA1TOP As String * 10      ' 測定値１(Top)  center         '04/02/12 ooba START =======>
    MESDATA2TOP As String * 10      ' 測定値２(Top)  R/2
    MESDATA3TOP As String * 10      ' 測定値３(Top)  Inside 10mm
    MESDATA4TOP As String * 10      ' 測定値４(Top)  Inside   6mm
    MESDATA5TOP As String * 10      ' 測定値５(Top)  Inside   3mm
    MESDATA1BOT As String * 10      ' 測定値１(Tail)  center
    MESDATA2BOT As String * 10      ' 測定値２(Tail)  R/2
    MESDATA3BOT As String * 10      ' 測定値３(Tail)  Inside 10mm
    MESDATA4BOT As String * 10      ' 測定値４(Tail)  Inside   6mm
    MESDATA5BOT As String * 10      ' 測定値５(Tail)  Inside   3mm  '04/02/12 ooba END =========>
End Type


' Ｓｘｌ確定指示返答
Public Type typ_TBCMY008
    SXL_ID As String * 13           ' SXL-ID
    TXID As String * 6              ' トランザクションID
    RET As String * 6               ' リターンコード
    REGDATE As Date                 ' 登録日付
    CHECKFLG As String * 1          ' チェックフラグ
End Type


' 受入実績
Public Type typ_TBCMY009
    LOTID As String * 12            ' ブロックID
    STRDTM As Date                  ' 受入日時
    STRUSER_ID As String * 10       ' 受入者
    TXID As String * 6              ' トランザクションID
    REGDATE As Date                 ' 登録日付
    SUMMITSENDFLAG As String * 1    ' SUMMIT送信フラグ
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 作業開始・終了
Public Type typ_TBCMY010
    LOTID As String * 12            ' ブロックID
    TRANCNT As Integer              ' 処理回数
    ROUTE_ID As String * 10         ' ル－トＩＤ
    ROUTE_VER As String * 3         ' ル－トIDバージョン
    OPE_ID As String * 6            ' 工程ID
    EQPID As String * 8             ' 装置ID
    STRDTM As Date                  ' 作業開始日時
    STRUSER_ID As String * 10       ' 作業開始者
    CMPDTM As Date                  ' 作業終了日時
    CMPUSER_ID As String * 10       ' 作業終了者
    CURRWPCS As Integer             ' ウェハー枚数
    TXID As String * 6              ' トランザクションID
    REGDATE As Date                 ' 登録日付
    SUMMITSENDFLAG As String * 1    ' SUMMIT送信フラグ
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' ウェハ－センタ－入庫情報
Public Type typ_TBCMY011
    LOTID As String * 12            ' ブロックID
    BLOCKSEQ As Integer             ' ブロック内連番
    INDTM As Date                   ' ウェハーセンター入庫日時
    BASKETID As String * 6          ' バスケットID
    SLOTNO As Integer               ' スロットNo
    CURRWPCS As Integer             ' ウェハー枚数
    EXISTFLG As String * 1          ' 存在フラグ
    TOP_POS As Integer              ' ブロックのTopからの 位置
    REJCAT As String * 1            ' 欠落理由
    TXID As String * 6              ' トランザクションID
    REGDATE As Date                 ' 登録日付
    SUMMITSENDFLAG As String * 1    ' SUMMIT送信フラグ
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 欠落情報
Public Type typ_TBCMY012
    LOTID As String * 12            ' ブロックID
    BLOCKSEQ As Integer             ' ブロック内連番
    REJPCS As Integer               ' 不良枚数
    TOP_POS As Integer              ' ブロックのTopからの 位置
    REJCAT As String * 1            ' 欠落理由
    REJDTTM As Date                 ' 欠落日
    REJPROC As String * 12          ' 欠落発見工程
    ALLSCRAP As String * 1          ' 全数スクラップ
    LENFROM As Integer              ' 長さ　FROM
    LENTO As Integer                ' 長さ　TO
    TXID As String * 6              ' トランザクションID
    CHKFLG As String * 1            ' チェックフラグ
    REGDATE As Date                 ' 登録日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type

' 測定評価結果
Public Type typ_TBCMY013
    SAMPLEID As String * 16         ' サンプルID
    OSITEM As String * 4            ' 評価項目
    MAISU As Integer                ' 評価枚数
    Spec As String * 10             ' 規格値
    NETSU As String * 2             ' 熱処理条件
    ET As String * 3                ' エッチング条件
    MES As String * 3               ' 計測方法
    DKAN As String * 10             ' ＤＫアニール条件
    MESDATA1 As String * 10         ' 測定データその１
    MESDATA2 As String * 10         ' 測定データその２
    MESDATA3 As String * 10         ' 測定データその３
    MESDATA4 As String * 10         ' 測定データその４
    MESDATA5 As String * 10         ' 測定データその５
    MESDATA6 As String * 10         ' 測定データその６
    MESDATA7 As String * 10         ' 測定データその７
    MESDATA8 As String * 10         ' 測定データその８
    MESDATA9 As String * 10         ' 測定データその９
    MESDATA10 As String * 10        ' 測定データその１０
    MESDATA11 As String * 10        ' 測定データその1１
    MESDATA12 As String * 10        ' 測定データその1２
    MESDATA13 As String * 10        ' 測定データその1３
    MESDATA14 As String * 10        ' 測定データその1４
    MESDATA15 As String * 10        ' 測定データその1５
    TXID As String * 6              ' トランザクションID
    REGDATE As Date                 ' 登録日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type

' エピ測定評価方法指示  2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
' エピ先行測定評価結果
Public Type typ_TBCMY022
    SAMPLEID As String * 16         ' サンプルID
    OSITEM As String * 4            ' 評価項目
    MAISU As Integer                ' 評価枚数
    Spec As String * 10             ' 規格値
    NETSU As String * 2             ' 熱処理条件
    ET As String * 3                ' エッチング条件
    MES As String * 3               ' 計測方法
    DKAN As String * 10             ' ＤＫアニール条件
    MESDATA1 As String * 10         ' 測定データその１
    MESDATA2 As String * 10         ' 測定データその２
    MESDATA3 As String * 10         ' 測定データその３
    MESDATA4 As String * 10         ' 測定データその４
    MESDATA5 As String * 10         ' 測定データその５
    MESDATA6 As String * 10         ' 測定データその６
    MESDATA7 As String * 10         ' 測定データその７
    MESDATA8 As String * 10         ' 測定データその８
    MESDATA9 As String * 10         ' 測定データその９
    MESDATA10 As String * 10        ' 測定データその１０
    MESDATA11 As String * 10        ' 測定データその1１
    MESDATA12 As String * 10        ' 測定データその1２
    MESDATA13 As String * 10        ' 測定データその1３
    MESDATA14 As String * 10        ' 測定データその1４
    MESDATA15 As String * 10        ' 測定データその1５
    TXID As String * 6              ' トランザクションID
    REGDATE As Date                 ' 登録日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type

' ブロック確定情報
Public Type typ_TBCMY014
    LOTID As String * 12            ' ブロックID
    BLOCKSEQ As Integer             ' ブロック内連番
    CURRWPCS As Integer             ' ウェハー枚数
    EXISTFLG As String * 1          ' 存在フラグ
    SXL_ID As String * 13           ' シングルID
    TOP_POS As String * 3           ' ブロックのTopからの 位置
    REJCAT As String * 1            ' 欠落理由
    TXID As String * 6              ' トランザクションID
    REGDATE As Date                 ' 登録日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' シングルマップ情報
Public Type typ_TBCMY015
    SXL_ID As String * 13           ' シングルID
    SXLSEQ As Integer               ' シングル内連番
    SXLWPCS As Integer              ' ウェハー枚数
    BLOCKID As String * 12          ' ブロックID
    BLOCKSEQ As Integer             ' ブロック内連番
    EXISTFLG As String * 1          ' 存在フラグ
    REJCAT As String * 1            ' 欠落理由
    REGDATE As Date                 ' 登録日付
    SUMMITSENDFLAG As String * 1    ' SUMMIT送信フラグ
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 抜試異常返答
Public Type typ_TBCMY016
    SAMPLEID As String * 16         ' SAMPLEID
    RPTDATE As Date                 ' 報告日時
    RET As String * 6               ' 抜試結果
    TXID As String * 6              ' トランザクションID
    REGDATE As Date                 ' 登録日付
End Type

' 関連ブロック紐付紐切　07/08/06 ooba
Public Type typ_TBCMY023
    CRYNUM As String * 12           '結晶番号
    TRANCNT As Integer              '処理回数
    BLOCKID As String * 12          'ﾌﾞﾛｯｸID
    PROCCAT As String * 1           '処理区分
    TXID As String * 6              'ﾄﾗﾝｻﾞｸｼｮﾝID
End Type

'--------------- 2008/07/25 INSERT START  By Systech ---------------
' ブロック品番振替要求
Public Type typ_TBCMY027
    CRYNUM      As String * 12      '結晶番号
    TRANCNT     As String * 2       '処理回数
    BLOCKID     As String * 12      'ブロックID
    TXID        As String * 6       'トランザクションID
    REQDATE     As Date             '振替要求日時
    USER_ID     As String * 10      '作業者コード
    FRMAINHIN   As String * 10      '振替元代表品番
    TOMAINHIN   As String * 10      '振替先代表品番
    HINCNT      As Integer          '構成品番数
    FRHIN1      As String * 10      '振替元構成品番１
    TOHIN1      As String * 10      '振替先構成品番１
    FRHIN2      As String * 10      '振替元構成品番２
    TOHIN2      As String * 10      '振替先構成品番２
    FRHIN3      As String * 10      '振替元構成品番３
    TOHIN3      As String * 10      '振替先構成品番３
    FRHIN4      As String * 10      '振替元構成品番４
    TOHIN4      As String * 10      '振替先構成品番４
    FRHIN5      As String * 10      '振替元構成品番５
    TOHIN5      As String * 10      '振替先構成品番５
    REGDATE     As Date             '登録日時
    CHECKFLG    As String * 1       'チェックフラグ
    SNDKDWH     As String * 1       'DWH送信フラグ
    SDAYDWH     As Date             'DWH送信日付
    SNDKSPC     As String * 1       'SPC送信フラグ
    SDAYSPC     As Date             'SPC送信日付
    PLANTCAT    As String * 2       '事業所区分
End Type

' ブロック品番振替要求応答
Public Type typ_TBCMY028
    CRYNUM      As String * 12      '結晶番号
    TRANCNT     As String * 2       '処理回数
    BLOCKID     As String * 12      'ブロックID
    TXID        As String * 6       'トランザクションID
    ALLJUDGRES  As String * 1       '総合判定結果
    JUDGDATE    As Date             '判定日時
    HINCNT      As Integer          '構成品番数
    FRHIN1      As String * 10      '振替元構成品番１
    TOHIN1      As String * 10      '振替先構成品番１
    JudgRes1    As String * 1       '判定結果１
    ERRCODE1    As String * 5       'エラーコード１
    FRHIN2      As String * 10      '振替元構成品番２
    TOHIN2      As String * 10      '振替先構成品番２
    JUDGRES2    As String * 1       '判定結果２
    ERRCODE2    As String * 5       'エラーコード２
    FRHIN3      As String * 10      '振替元構成品番３
    TOHIN3      As String * 10      '振替先構成品番３
    JUDGRES3    As String * 1       '判定結果３
    ERRCODE3    As String * 5       'エラーコード３
    FRHIN4      As String * 10      '振替元構成品番４
    TOHIN4      As String * 10      '振替先構成品番４
    JUDGRES4    As String * 1       '判定結果４
    ERRCODE4    As String * 5       'エラーコード４
    FRHIN5      As String * 10      '振替元構成品番５
    TOHIN5      As String * 10      '振替先構成品番５
    JUDGRES5    As String * 1       '判定結果５
    ERRCODE5    As String * 5       'エラーコード５
    REGDATE     As Date             '登録日付
    SENDFLAG    As String * 1       '送信フラグ
    SENDDATE    As Date             '送信日付
    SNDKDWH     As String * 1       'DWH送信フラグ
    SDAYDWH     As Date             'DWH送信日付
    SNDKSPC     As String * 1       'SPC送信フラグ
    SDAYSPC     As Date             'SPC送信日付
    PLANTCAT    As String * 2       '事業所区分
End Type
'--------------- 2008/07/25 INSERT  END   By Systech ---------------

' 社員マスター
Public Type typ_TBCMB001
    StaffID As String * 8           ' 社員ID
    PASSWD As String * 8            ' パスワード
    JFMLNAME As String              ' 日本語名（氏）
    JFSTNAME As String              ' 日本語名（名）
    RFMLNAME As String              ' ローマ字名（氏）
    RFSTNAME As String              ' ローマ字名（名）
    EXECODE As String * 4           ' 実行権限コード
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
End Type


' 工程コードマスター
Public Type typ_TBCMB002
    KRPROCID As String * 5          ' 管理工程ID
    PROCCODE As String * 5          ' 工程コード
    JPNNAME As String               ' 日本語名
    PROCSEQ As Integer              ' 工程内順序
    NOTE As String                  ' 備考
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
End Type


' テーブル情報マスター
Public Type typ_TBCMB018
    TABLENAME As String             ' テーブル名
    COLUM As String                 ' カラム名
    NO As Integer                   ' カラム順
    TYPE As String * 16             ' 型
    PKEY As String * 1              ' 主キー
    BASETYPE As String * 16         ' 基本型
    SIZE1 As Integer                ' 型サイズ１
    SIZE2 As Integer                ' 型サイズ２
    MQBYTE As Long                  ' ＭＱバイト長
    JTABLE As String                ' 日本語テーブル名
    JCOLUM As String                ' 日本語カラム名
    REF1 As String                  ' 備考１
    REF2 As String                  ' 備考２
    TBLKBN As String * 1            ' テーブル種別
    REGDATE As Date                 ' 登録日付
End Type


' メッセージマスター
Public Type typ_TBCMB003
    MsgID As String * 5             ' メッセージID
    FORMINFO As String              ' Format情報
    USEPRCID As String * 5          ' 利用処理ID
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
End Type


' 権限マスター
Public Type typ_TBCMB004
    AUTHCODE As String * 4          ' 実行権限コード
    TRANID As String * 5            ' 処理ID
    PWCHECK As String * 1           ' パスワードチェック
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
End Type


' コードマスター
Public Type typ_TBCMB005
    SYSCLASS As String * 2          ' システム区分
    Class As String * 2             ' 区分
    CODE As String * 5              ' コード
    INFO1 As String                 ' 情報１
    INFO2 As String                 ' 情報２
    INFO3 As String                 ' 情報３
    INFO4 As String                 ' 情報４
    INFO5 As String                 ' 情報５
    INFO6 As String                 ' 情報６
    INFO7 As String                 ' 情報７
    INFO8 As String                 ' 情報８
    INFO9 As String                 ' 情報９
    NOTE As String                  ' 備考
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
End Type


' 実行管理マスター
Public Type typ_TBCMB006
    ProcID As String * 12           ' 処理ID
    EXENAME As String               ' 実行ファイル名
    BIKOU As String                 ' 備考
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
End Type


' ρ区分マスター
Public Type typ_TBCMB007
    RCLSCODE As String * 3          ' ρ区分コード
    TYPE As String * 1              ' タイプ
    MINRESIST As Double             ' MIN　抵抗値
    MINMOVAL As Double              ' MIN　MO値
    MINFVAL As Double               ' MIN　F値
    MAXRESIST As Double             ' MAX　抵抗値
    MAXMOVAL As Double              ' MAX　MO値
    MAXFVAL As Double               ' MAX　F値
    REPRESIST As Double             ' 代表　抵抗値
    REPMOVAL As Double              ' 代表　MO値
    REPFVAL As Double               ' 代表　F値
    IonDensity As Double            ' 代表イオン濃度
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
End Type


' 精製原料濃度マスター
Public Type typ_TBCMB008
    MltType As String * 3           ' タイプ
    MINRESIST As Double             ' MIN　抵抗値
    MAXRESIST As Double             ' MAX　抵抗値
    IonDensity As Double            ' 代表イオン濃度
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
End Type


' ドーパント濃度マスター
Public Type typ_TBCMB009
    DopKind As String * 4           ' ドーパント種類
    IonDensity As Double            ' イオン濃度
    CoreCoeff As Integer            ' 補正係数
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
End Type


' ドーパント計算係数マスター
Public Type typ_TBCMB010
    TYPE As String * 1              ' タイプ
    ResFrom As Double               ' 抵抗From
    ResTo As Double                 ' 抵抗To
    FixNumA As Double               ' 定数A
    FizNumB As Double               ' 定数B
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
End Type


' PG-ID管理
Public Type typ_TBCMB011
    PGID As String * 10             ' PG-ID
    HZPART As String * 4            ' HZパーツ
    HZPTRN As String * 2            ' HZパターン
    SPACER As String * 5            ' スペーサ
    UPRING As String * 5            ' アッパーリング
    CHARGE As Long                  ' チャージ量
    RTBPOS As Integer               ' ルツボ位置
    RTBSIZE As String * 2           ' ルツボサイズ
    GAP As Integer                  ' ギャップ
    UPDM As Integer                 ' 引上直径
    UPLENGTH As Integer             ' 引上長（全長）
    UPRC As Integer                 ' 引上（RC）
    RFRNEED As String * 1           ' リフラクタ要否
    UPSPIN As String * 10           ' 上軸回転数
    DOWNSPIN As String * 10         ' 下軸回転数
    ROPRESS As String * 8           ' 炉内圧
    ARUGON As String * 7            ' アルゴン量
    AIMOIMIN As Double              ' ねらいOi（MIN)
    AIMOIMAX As Double              ' ねらいOi（MAX)
    HCCLASS As String * 7           ' HC種類
    HC As String                    ' HC
    AVEUPSPD As Double              ' 平均引上速度
    UPCNTL As String * 1            ' 引上制御
    BTMSHAPE As String * 1          ' ボトム形状
    MAGSTR As Double                ' 磁場強度
    MAGPOS As Long                  ' 磁場位置
    CONDGRT As String * 10          ' 条件保証登録
    MODEL As String * 4             ' 機種
    UPMETHOD As String * 4          ' 引上方法
    UPCLASS As String * 2           ' 引上区分
    UPNUM As String * 1             ' 引上本数
    OPETIME As Long                 ' 運転時間
    WTRCOOL As String * 1           ' 水冷管要否
    PGID2 As String * 10            ' PG-ID（一本引）
    RCPT1 As String * 3             ' 対応レシピNo（T1)
    RCPT2 As String * 3             ' 対応レシピNo（T2)
    RCPT3 As String * 3             ' 対応レシピNo（T3)
    RCPT4 As String * 3             ' 対応レシピNo（T4)
    RCPT5 As String * 3             ' 対応レシピNo（T5)
    RCPT6 As String * 3             ' 対応レシピNo（T6)
    CNTL1 As String * 1             ' 制限項目（1）
    CNTL2 As String * 1             ' 制限項目（2）
    CNTL3 As String * 1             ' 制限項目（3）
    CNTL4 As String * 1             ' 制限項目（4）
    CNTL5 As String * 1             ' 制限項目（5）
    CNTL6 As String * 1             ' 制限項目（6）
    CNTL7 As String * 1             ' 制限項目（7）
    CNTL8 As String * 1             ' 制限項目（8）
    CNTL9 As String * 1             ' 制限項目（9）
    CNTL10 As String * 1            ' 制限項目（10）
    CNTL11 As String * 1            ' 制限項目（11）
    CNTL12 As String * 1            ' 制限項目（12）
    CNTL13 As String * 1            ' 制限項目（13）
    CNTL14 As String * 1            ' 制限項目（14）
    CNTL15 As String * 1            ' 制限項目（15）
    DRDOP  As String                ' ドープ     4/30
    DRAR3  As String                ' アルゴン№３流量
    RUNCOND1 As String              ' 運転条件１
    RUNCOND2 As String              ' 運転条件２
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 制作条件
Public Type typ_TBCMB012
    MKCONDNO As String * 12         ' 制作条件No.
    MODEL As String * 1             ' 機種
    RTBSIZE As String * 1           ' ルツボサイズ
    CHARGE As String * 1            ' チャージ量
    HZTYPE As String * 1            ' HZタイプ
    UPSPDTYP As String * 1          ' 引上げ速度タイプ
    MAGTYPE As String * 1           ' 磁場タイプ
    USECLS As String * 1            ' 使用区分
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日時
End Type


' 制作条件PG-ID対応
Public Type typ_TBCMB013
    MKCONDNO As String * 12         ' 制作条件No.
    PGIDNO As String * 10           ' PG-IDNo
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' GFA校正情報
Public Type typ_TBCMB014
    GOUKI As String * 3             ' 号機
    INPDATE As Date                 ' 日付
    FTIRFZI As Double               ' FTIR（FZ)
    FTIRCZH As Double               ' FTIR（CZ高）
    FTIRCZC As Double               ' FTIR（CZ中）
    MS1FZ As Double                 ' 測定サンプル1（FZ)
    MS1CZ1 As Double                ' 測定サンプル1（CZ-1)
    MS1CZ2 As Double                ' 測定サンプル1（CZ-2)
    MS2FZ As Double                 ' 測定サンプル2（FZ)
    MS2CZ1 As Double                ' 測定サンプル2（CZ-1)
    MS2CZ2 As Double                ' 測定サンプル2（CZ-2)
    MS3FZ As Double                 ' 測定サンプル3（FZ)
    MS3CZ1 As Double                ' 測定サンプル3（CZ-1)
    MS3CZ2 As Double                ' 測定サンプル3（CZ-2)
    MS4FZ As Double                 ' 測定サンプル4（FZ)
    MS4CZ1 As Double                ' 測定サンプル4（CZ-1)
    MS4CZ2 As Double                ' 測定サンプル4（CZ-2)
    MS5FZ As Double                 ' 測定サンプル5（FZ)
    MS5CZ1 As Double                ' 測定サンプル5（CZ-1)
    MS5CZ2 As Double                ' 測定サンプル5（CZ-2)
    MSAVEFZ As Double               ' 測定平均（FZ）
    MSAVECZ1 As Double              ' 測定平均（CZ-1）
    MSAVECZ2 As Double              ' 測定平均（CZ-2）
    MSSGFZ As Double                ' 測定σ（FZ）
    MSSGCZ1 As Double               ' 測定σ（CZ-1）
    MSSGCZ2 As Double               ' 測定σ（CZ-2）
    MSPSGFZ As Double               ' 測定AVE+σ（FZ）
    MSPSGCZ1 As Double              ' 測定AVE+σ（CZ-1）
    MSPSGCZ2 As Double              ' 測定AVE+σ（CZ-2）
    MSNSGFZ As Double               ' 測定AVE-σ（FZ）
    MSNSGCZ1 As Double              ' 測定AVE-σ（CZ-1）
    MSNSGCZ2 As Double              ' 測定AVE-σ（CZ-2）
    MINFZ As Double                 ' MIN（FZ）
    MINCZ1 As Double                ' MIN（CZ-1）
    MINCZ2 As Double                ' MIN（CZ-2）
    MAXFZ As Double                 ' MAX（FZ）
    MAXCZ1 As Double                ' MAX（CZ-1）
    MAXCZ2 As Double                ' MAX（CZ-2）
    SGCK1FZ As Double               ' σckサンプル1（FZ)
    SGCK1CZ1 As Double              ' σckサンプル1（CZ-1)
    SGCK1CZ2 As Double              ' σckサンプル1（CZ-2)
    SGCK2FZ As Double               ' σckサンプル2（FZ)
    SGCK2CZ1 As Double              ' σckサンプル2（CZ-1)
    SGCK2CZ2 As Double              ' σckサンプル2（CZ-2)
    SGCK3FZ As Double               ' σckサンプル3（FZ)
    SGCK3CZ1 As Double              ' σckサンプル3（CZ-1)
    SGCK3CZ2 As Double              ' σckサンプル3（CZ-2)
    SGCK4FZ As Double               ' σckサンプル4（FZ)
    SGCK4CZ1 As Double              ' σckサンプル4（CZ-1)
    SGCK4CZ2 As Double              ' σckサンプル4（CZ-2)
    SGCK5FZ As Double               ' σckサンプル5（FZ)
    SGCK5CZ1 As Double              ' σckサンプル5（CZ-1)
    SGCK5CZ2 As Double              ' σckサンプル5（CZ-2)
    SGCKDFZ As Double               ' σckデータ数（FZ）
    SGCKDCZ1 As Double              ' σckデータ数（CZ-1）
    SGCKDCZ2 As Double              ' σckデータ数（CZ-2）
    SGCKAFZ As Double               ' σck平均（FZ）
    SGCKAACZ1 As Double             ' σck平均（CZ-1）
    SGCKACZ2 As Double              ' σck平均（CZ-2）
    SGNFZ As Double                 ' σckσ（FZ）
    SGNCZ1 As Double                ' σckσ CZ-1）
    SGNCZ2 As Double                ' σckσ（CZ-2）
    FTIRFZ As Double                ' FTIR換算（FZ）
    FTIRCZ1 As Double               ' FTIR換算（CZ-1）
    FTIRCZ2 As Double               ' FTIR換算（CZ-2）
    EFFECTTM As Integer             ' 有効時間
    YCOEF As Double                 ' ＦＴＩＲ換算式（Ｙ切片）
    XCOEF As Double                 ' ＦＴＩＲ換算式（Ｘ係数）
    RSQUARE As Double               ' Ｒ２乗
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
'2006/05/22追加
    SGCKST      As Double           ' σ判定基準
    SGCKFZ      As String * 1       ' σ判定(FZ)
    SGCKCZ1     As String * 1       ' σ判定(CZ-1)
    SGCKCZ2     As String * 1       ' σ判定(CZ-2)
    FTIRCKST    As Double           ' FTIR換算判定基準
    FTIRCKFZ    As String * 1       ' FTIR換算判定(FZ)
    FTIRCKCZ1   As String * 1       ' FTIR換算判定(CZ-1)
    FTIRCKCZ2   As String * 1       ' FTIR換算判定(CZ-2)
'2010/02/09追加 SETsw kubota
    MS6FZ As Double                 ' 測定サンプル6（FZ)
    MS6CZ1 As Double                ' 測定サンプル6（CZ-1)
    MS6CZ2 As Double                ' 測定サンプル6（CZ-2)
    SGCK6FZ As Double               ' σckサンプル6（FZ)
    SGCK6CZ1 As Double              ' σckサンプル6（CZ-1)
    SGCK6CZ2 As Double              ' σckサンプル6（CZ-2)
    CVFZ As Double                  ' CV(%)（FZ）
    CVCZ1 As Double                 ' CV(%)（CZ-1）
    CVCZ2 As Double                 ' CV(%)（CZ-2）
End Type


' 連番管理
Public Type typ_TBCMB015
    CNTMNGCD As String * 4          ' 連番種別管理コード
    CNTNUMCD As String * 4          ' 連番種別コード
    CONTNUM As Long                 ' 連番
    MAXFIG As Integer               ' 最大桁数
    NUMUNIT As String * 1           ' 連番単位区分
    NUMNAME As String               ' 連番名
    CLRDATE As Date                 ' クリア日付
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
End Type


' バージョン管理
Public Type typ_TBCMB016
    MACHINE As String * 8           ' マシン名
    EXENAME As String               ' 実行ファイル名
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
End Type


' ラベルプリンタ要求
Public Type typ_TBCMB017
    QUEDATE As Date                 ' キュー日付
    PRINTKIND As String * 4         ' 印刷種類
    ENDFLG As String * 1            ' 完了区分
    STATUS As String * 4            ' 終了ステータス
    PrintInfo1 As String            ' 印刷情報１
    PrintInfo2 As String            ' 印刷情報２
    PrintInfo3 As String            ' 印刷情報３
    PrintInfo4 As String            ' 印刷情報４
    PrintInfo5 As String            ' 印刷情報５
    PrintInfo6 As String            ' 印刷情報６
    PrintInfo7 As String            ' 印刷情報７
    PrintInfo8 As String            ' 印刷情報８
    PrintInfo9 As String            ' 印刷情報９
    PrintInfo10 As String           ' 印刷情報１０
    StaffID As String * 8           ' 要求担当者ID
    MACHINE As String * 8           ' 要求マシン名
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
End Type


'Add Start 2011/03/30 SMPK H.Ohkubo
' FRS校正情報
Public Type typ_TBCMB019
    GOUKI       As String * 3       ' 号機
    INPDATE     As Date             ' 日付
    FTIROIL     As Double           ' FTIR（Oi低)
    FTIROIM     As Double           ' FTIR（Oi中）
    FTIROIH     As Double           ' FTIR（Oi高）
    MS1OIL      As Double           ' 測定サンプル1（Oi低)
    MS1OIM      As Double           ' 測定サンプル1（Oi中)
    MS1OIH      As Double           ' 測定サンプル1（Oi高)
    MS2OIL      As Double           ' 測定サンプル2（Oi低)
    MS2OIM      As Double           ' 測定サンプル2（Oi中)
    MS2OIH      As Double           ' 測定サンプル2（Oi高)
    MS3OIL      As Double           ' 測定サンプル3（Oi低)
    MS3OIM      As Double           ' 測定サンプル3（Oi中)
    MS3OIH      As Double           ' 測定サンプル3（Oi高)
    MS4OIL      As Double           ' 測定サンプル4（Oi低)
    MS4OIM      As Double           ' 測定サンプル4（Oi中)
    MS4OIH      As Double           ' 測定サンプル4（Oi高)
    MS5OIL      As Double           ' 測定サンプル5（Oi低)
    MS5OIM      As Double           ' 測定サンプル5（Oi中)
    MS5OIH      As Double           ' 測定サンプル5（Oi高)
    MSAVEOIL    As Double           ' 測定平均（Oi低）
    MSAVEOIM    As Double           ' 測定平均（Oi中）
    MSAVEOIH    As Double           ' 測定平均（Oi高）
    MSSGOIL     As Double           ' 測定σ（Oi低）
    MSSGOIM     As Double           ' 測定σ（Oi中）
    MSSGOIH     As Double           ' 測定σ（Oi高）
    MSPSGOIL    As Double           ' 測定AVE+σ（Oi低）
    MSPSGOIM    As Double           ' 測定AVE+σ（Oi中）
    MSPSGOIH    As Double           ' 測定AVE+σ（Oi高）
    MSNSGOIL    As Double           ' 測定AVE-σ（Oi低）
    MSNSGOIM    As Double           ' 測定AVE-σ（Oi中）
    MSNSGOIH    As Double           ' 測定AVE-σ（Oi高）
    MINOIL      As Double           ' MIN（Oi低）
    MINOIM      As Double           ' MIN（Oi中）
    MINOIH      As Double           ' MIN（Oi高）
    MAXOIL      As Double           ' MAX（Oi低）
    MAXOIM      As Double           ' MAX（Oi中）
    MAXOIH      As Double           ' MAX（Oi高）
    SGCK1OIL    As Double           ' σckサンプル1（Oi低)
    SGCK1OIM    As Double           ' σckサンプル1（Oi中)
    SGCK1OIH    As Double           ' σckサンプル1（Oi高)
    SGCK2OIL    As Double           ' σckサンプル2（Oi低)
    SGCK2OIM    As Double           ' σckサンプル2（Oi中)
    SGCK2OIH    As Double           ' σckサンプル2（Oi高)
    SGCK3OIL    As Double           ' σckサンプル3（Oi低)
    SGCK3OIM    As Double           ' σckサンプル3（Oi中)
    SGCK3OIH    As Double           ' σckサンプル3（Oi高)
    SGCK4OIL    As Double           ' σckサンプル4（Oi低)
    SGCK4OIM    As Double           ' σckサンプル4（Oi中)
    SGCK4OIH    As Double           ' σckサンプル4（Oi高)
    SGCK5OIL    As Double           ' σckサンプル5（Oi低)
    SGCK5OIM    As Double           ' σckサンプル5（Oi中)
    SGCK5OIH    As Double           ' σckサンプル5（Oi高)
    SGCKDOIL    As Double           ' σckデータ数（Oi低）
    SGCKDOIM    As Double           ' σckデータ数（Oi中）
    SGCKDOIH    As Double           ' σckデータ数（Oi高）
    SGCKAOIL    As Double           ' σck平均（Oi低）
    SGCKAAOIM   As Double           ' σck平均（Oi中）
    SGCKAOIH    As Double           ' σck平均（Oi高）
    SGNOIL      As Double           ' σckσ（Oi低）
    SGNOIM      As Double           ' σckσ（Oi中）
    SGNOIH      As Double           ' σckσ（Oi高）
    FTIRKOIL    As Double           ' FTIR換算（Oi低）
    FTIRKOIM    As Double           ' FTIR換算（Oi中）
    FTIRKOIH    As Double           ' FTIR換算（Oi高）
    EFFECTTM    As Integer          ' 有効時間
    YCOEF       As Double           ' ＦＴＩＲ換算式（Ｙ切片）
    XCOEF       As Double           ' ＦＴＩＲ換算式（Ｘ係数）
    RSQUARE     As Double           ' Ｒ２乗
    SGCKST      As Double           ' σ判定基準
    SGCKOIL     As String * 1       ' σ判定(Oi低)
    SGCKOIM     As String * 1       ' σ判定(Oi中)
    SGCKOIH     As String * 1       ' σ判定(Oi高)
    FTIRCKST    As Double           ' FTIR換算判定基準
    FTIRCKOIL   As String * 1       ' FTIR換算判定(Oi低)
    FTIRCKOIM   As String * 1       ' FTIR換算判定(Oi中)
    FTIRCKOIH   As String * 1       ' FTIR換算判定(Oi高)
    MS6OIL      As Double           ' 測定サンプル6（Oi低)
    MS6OIM      As Double           ' 測定サンプル6（Oi中)
    MS6OIH      As Double           ' 測定サンプル6（Oi高)
    SGCK6OIL    As Double           ' σckサンプル6（Oi低)
    SGCK6OIM    As Double           ' σckサンプル6（Oi中)
    SGCK6OIH    As Double           ' σckサンプル6（Oi高)
    CVOIL       As Double           ' CV(%)（Oi低）
    CVOIM       As Double           ' CV(%)（Oi中）
    CVOIH       As Double           ' CV(%)（Oi高）
    TSTAFFID    As String * 8       ' 登録社員ID
    REGDATE     As Date             ' 登録日付
    KSTAFFID    As String * 8       ' 更新社員ID
    UPDDATE     As Date             ' 更新日付
    SENDFLAG    As String * 1       ' 送信フラグ
    SENDDATE    As Date             ' 送信日付
End Type
'Add End 2011/03/30 SMPK H.Ohkubo

' 結晶情報
Public Type typ_TBCME037
    CRYNUM As String * 12           ' 結晶番号
    DELCLS As String * 1            ' 削除区分
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCD As String * 5            ' 工程コード
    LPKRPROCCD As String * 5        ' 最終通過管理工程
    LASTPASS As String * 5          ' 最終通過工程
    RPHINBAN As String * 8          ' ねらい品番
    RPREVNUM As Integer             ' ねらい品番製品番号改訂番号
    RPFACT As String * 1            ' ねらい品番工場
    RPOPCOND As String * 1          ' ねらい品番操業条件
    PRODCOND As String * 12         ' 製作条件
    PGID As String * 8              ' ＰＧ－ＩＤ
    UPLENGTH As Integer             ' 引上げ長さ
    TOPLENG As Integer              ' ＴＯＰ長さ
    BODYLENG As Integer             ' 直胴長さ
    BOTLENG As Integer              ' ＢＯＴ長さ
    FREELENG As Integer             ' フリー長
    DIAMETER As Integer             ' 直径
    CHARGE As Long                  ' チャージ量
    SEED As String * 4              ' シード
    ADDDPCLS As String * 4          ' 追加ドープ種類
    ADDDPPOS As Integer             ' 追加ドープ位置
    ADDDPVAL As Double              ' 追加ドープ量
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' ブロック設計
Public Type typ_TBCME038
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' 結晶内開始位置
    Length As Integer               ' 長さ
    USECLASS As String * 1          ' 使用区分
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 品番設計
Public Type typ_TBCME039
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' 結晶内開始位置
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 改訂番号
    FACT As String * 1              ' 工場
    OPCOND As String * 1            ' 操業条件
    Length As Integer               ' 長さ
    USECLASS As String * 1          ' 使用区分
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' ブロック管理
Public Type typ_TBCME040
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' 結晶内開始位置
    Length As Integer               ' 長さ
    REALLEN As Integer              ' 実長さ
    BLOCKID As String * 12          ' ブロックID
    KRPROCCD As String * 5          ' 現在管理工程
    NOWPROC As String * 5           ' 現在工程
    LPKRPROCCD As String * 5        ' 最終通過管理工程
    LASTPASS As String * 5          ' 最終通過工程
    DELCLS As String * 1            ' 削除区分
    LSTATCLS As String * 1          ' 最終状態区分
    RSTATCLS As String * 1          ' 流動状態区分
    HOLDCLS As String * 1           ' ホールド区分
    BDCAUS As String * 3            ' 不良理由
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SUMMITSENDFLAG As String * 1    ' SUMMIT送信フラグ
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
    PASSFLAG As String * 1          ' 通過フラグ　　'7/5　hama
End Type


' 品番管理
Public Type typ_TBCME041
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' 結晶内開始位置
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    Length As Integer               ' 長さ
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' SXL管理
Public Type typ_TBCME042
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' 結晶内開始位置
    Length As Integer               ' 長さ
    SXLID As String * 13            ' SXLID
    KRPROCCD As String * 5          ' 管理工程
    NOWPROC As String * 5           ' 現在工程
    LPKRPROCCD As String * 5        ' 最終通過管理工程
    LASTPASS As String * 5          ' 最終通過工程
    DELCLS As String * 1            ' 削除区分
    LSTATCLS As String * 1          ' 最終状態区分
    HOLDCLS As String * 1           ' ホールド区分
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    BDCAUS As String * 3            ' 不良理由
    COUNT As Integer                ' 枚数
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SUMMITSENDFLAG As String * 1    ' SUMMIT送信フラグ
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
    PASSFLAG As String * 1          ' 通過フラグ　　'4/5　Yam
End Type

' 結晶サンプル管理
Public Type typ_TBCME043
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' 結晶内位置
    SMPKBN As String * 1            ' サンプル区分
    SMPLNO As Long                  ' サンプルNo    Integer→Long   サンプル№6桁対応 2007/05/28 SETsw kubota
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KTKBN As String * 1             ' 確定区分
    CRYINDRS As String * 1          ' 結晶検査指示（Rs)
    CRYINDOI As String * 1          ' 結晶検査指示（Oi)
    CRYINDB1 As String * 1          ' 結晶検査指示（B1)
    CRYINDB2 As String * 1          ' 結晶検査指示（B2）
    CRYINDB3 As String * 1          ' 結晶検査指示（B3)
    CRYINDL1 As String * 1          ' 結晶検査指示（L1)
    CRYINDL2 As String * 1          ' 結晶検査指示（L2)
    CRYINDL3 As String * 1          ' 結晶検査指示（L3)
    CRYINDL4 As String * 1          ' 結晶検査指示（L4)
    CRYINDCS As String * 1          ' 結晶検査指示（Cs)
    CRYINDGD As String * 1          ' 結晶検査指示（GD)
    CRYINDT As String * 1           ' 結晶検査指示（T)
    CRYINDEP As String * 1          ' 結晶検査指示（EPD)
    CRYRESRS As String * 1          ' 結晶検査実績（Rs)
    CRYRESOI As String * 1          ' 結晶検査実績（Oi)
    CRYRESB1 As String * 1          ' 結晶検査実績（B1)
    CRYRESB2 As String * 1          ' 結晶検査実績（B2）
    CRYRESB3 As String * 1          ' 結晶検査実績（B3)
    CRYRESL1 As String * 1          ' 結晶検査実績（L1)
    CRYRESL2 As String * 1          ' 結晶検査実績（L2)
    CRYRESL3 As String * 1          ' 結晶検査実績（L3)
    CRYRESL4 As String * 1          ' 結晶検査実績（L4)
    CRYRESCS As String * 1          ' 結晶検査実績（Cs)
    CRYRESGD As String * 1          ' 結晶検査実績（GD)
    CRYREST As String * 1           ' 結晶検査実績（T)
    CRYRESEP As String * 1          ' 結晶検査実績（EPD)
    SMPLNUM As Integer              ' サンプル枚数
    SMPLPAT As String * 1           ' サンプルパターン
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
    '2003/09/29 追加　KURO
    SMPLNOOI As Long                ' サンプルNo(OI)    Integer→Long   サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLNOCS As Long                ' サンプルNo(CS)    Integer→Long   サンプル№6桁対応 2007/05/28 SETsw kubota
    XTALCS   As String * 12         ' 結晶番号
    '' 2007/11/26 y.hosokawa Update 15桁対応
    BLOCKCS  As String * 15         ' ブロックID
    'BLOCKCS  As String * 12         ' ブロックID
    CRYSMPLIDRS1CS As Long          '                   Integer→Long   サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYSMPLIDRS2CS As Long          '                   Integer→Long   サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYRESRS1CS    As String * 1
    CRYRESRS2CS    As String * 1
    CRYSMPLIDB1CS  As Long          '                   Integer→Long   サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYSMPLIDB2CS  As Long          '                   Integer→Long   サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYSMPLIDB3CS  As Long          '                   Integer→Long   サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYSMPLIDL1CS  As Long          '                   Integer→Long   サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYSMPLIDL2CS  As Long          '                   Integer→Long   サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYSMPLIDL3CS  As Long          '                   Integer→Long   サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYSMPLIDL4CS  As Long          '                   Integer→Long   サンプル№6桁対応 2007/05/28 SETsw kubota
    QCKBNCS As String * 1           ' 管理区分   2009/11/05 SETsw kubota
End Type

'' 新サンプル管理(ﾌﾞﾛｯｸ)
Public Type typ_XSDCS
    CRYNUMCS As String * 12         'ブロックID
    SMPKBNCS As String * 1          'サンプル区分
    TBKBNCS As String * 1           'T/B区分
    REPSMPLIDCS As Long             '代表サンプルID     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    XTALCS As String * 12           '結晶番号
    INPOSCS As Integer              '結晶内位置
    HINBCS As String * 8            '品番
    REVNUMCS As Integer             '製品番号改訂番号
    FACTORYCS As String * 1         '工場
    OPECS As String * 1             '操業条件
    KTKBNCS As String * 1           '確定区分
    BLKKTFLAGCS As String * 1       'ブロック確定フラグ
    CRYSMPLIDRSCS As Long           'サンプルID(Rs)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYSMPLIDRS1CS As Long          '推定サンプルID1(Rs)    Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYSMPLIDRS2CS As Long          '推定サンプルID2(Rs)    Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYINDRSCS As String * 1        '状態FLG(Rs)
    CRYRESRS1CS As String * 1       '実績FLG1(Rs)
    CRYRESRS2CS As String * 1       '実績FLG2(Rs)
    CRYSMPLIDOICS As Long           'サンプルID(Oi)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYINDOICS As String * 1        '状態FLG(Oi)
    CRYRESOICS As String * 1        '実績FLG(Oi)
    CRYSMPLIDB1CS As Long           'サンプルID(B1)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYINDB1CS As String * 1        '状態FLG(B1)
    CRYRESB1CS As String * 1        '実績FLG(B1)
    CRYSMPLIDB2CS As Long           'サンプルID(B2)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYINDB2CS As String * 1        '状態FLG(B2)
    CRYRESB2CS As String * 1        '実績FLG(B2)
    CRYSMPLIDB3CS As Long           'サンプルID(B3)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYINDB3CS As String * 1        '状態FLG(B3)
    CRYRESB3CS As String * 1        '実績FLG(B3)
    CRYSMPLIDL1CS As Long           'サンプルID(L1)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYINDL1CS As String * 1        '状態FLG(L1)
    CRYRESL1CS As String * 1        '実績FLG(L1)
    CRYSMPLIDL2CS As Long           'サンプルID(L2)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYINDL2CS As String * 1        '状態FLG(L2)
    CRYRESL2CS As String * 1        '実績FLG(L2)
    CRYSMPLIDL3CS As Long           'サンプルID(L3)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYINDL3CS As String * 1        '状態FLG(L3)
    CRYRESL3CS As String * 1        '実績FLG(L3)
    CRYSMPLIDL4CS As Long           'サンプルID(L4)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYINDL4CS As String * 1        '状態FLG(L4)
    CRYRESL4CS As String * 1        '実績FLG(L4)
    CRYSMPLIDCSCS As Long           'サンプルID(Cs)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYINDCSCS As String * 1        '状態FLG(Cs)
    CRYRESCSCS As String * 1        '実績FLG(Cs)
    CRYSMPLIDGDCS As Long           'サンプルID(GD)     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYINDGDCS As String * 1        '状態FLG(GD)
    CRYRESGDCS As String * 1        '実績FLG(GD)
    CRYSMPLIDTCS As Long            'サンプルID(T)      Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYINDTCS As String * 1         '状態FLG(T)
    CRYRESTCS As String * 1         '実績FLG(T)
    CRYSMPLIDEPCS As Long           'サンプルID(EPD)    Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    CRYINDEPCS As String * 1        '状態FLG(EPD)
    CRYRESEPCS As String * 1        '実績FLG(EPD)
    SMPLNUMCS As Integer            'サンプル枚数
    SMPLPATCS As String * 1         'サンプルパターン
    LIVKCS As String * 1            '生死区分
    TSTAFFCS As String * 8          '登録社員ID
    TDAYCS As Date                  '登録日付
    KSTAFFCS As String * 8          '更新社員ID
    KDAYCS As Date                  '更新日付
    SNDKCS As String * 1            '送信フラグ
    SNDDAYCS As Date                '送信日付
    RPCRYNUMCS As String * 12       '親ブロックID　05/10/17 ooba
    CUTFLGCS As String * 1          '切断フラグ　05/10/17 ooba
    
    '2009/08 SUMCO Akizuki X線測定実績　項目追加
    CRYSMPLIDXCS As Long            'サンプルID(X線)
    CRYINDXCS As String * 1         '状態FLG(X線)
    CRYRESXCS As String * 1         '実績FLG(X線)
    
    QCKBNCS As String * 1           '管理区分   2009/11/05 SETsw kubota

    'Add Start 2010/12/13 SMPK Miyata
    CRYSMPLIDCCS    As Long         'サンプルID(C)
    CRYINDCCS       As String * 1   '状態FLG(C)
    CRYRESCCS       As String * 1   '実績FLG(C)
    CRYSMPLIDCJCS   As Long         'サンプルID(CJ)
    CRYINDCJCS      As String * 1   '状態FLG(CJ)
    CRYRESCJCS      As String * 1   '実績FLG(CJ)
    CRYSMPLIDCJLTCS As Long         'サンプルID(CJLT)
    CRYINDCJLTCS    As String * 1   '状態FLG(CJLT)
    CRYRESCJLTCS    As String * 1   '実績FLG(CJLT)
    CRYSMPLIDCJ2CS  As Long         'サンプルID(CJ2)
    CRYINDCJ2CS     As String * 1   '状態FLG(CJ2)
    CRYRESCJ2CS     As String * 1   '実績FLG(CJ2)
    'Add End   2010/12/13 SMPK Miyata

End Type

'2003/09/1 ｺﾒﾝﾄｱｳﾄ SystemBrain
' WFサンプル管理
'Public Type typ_TBCME044
'    CRYNUM As String * 12           ' 結晶番号
'    IngotPos As Integer             ' 結晶内位置
'    SMPKBN As String * 1            ' サンプル区分
'    SMPLID As String * 16           ' サンプルID
'    BKSMPLID As String * 16         ' 変更前サンプルID  'add 2003/05/06 hitec)matsumoto
'    hinban As String * 8            ' 品番
'    REVNUM As Integer               ' 製品番号改訂番号
'    factory As String * 1           ' 工場
'    opecond As String * 1           ' 操業条件
'    KTKBN As String * 1             ' 確定区分
'    WFINDRS As String * 1           ' WF検査指示（Rs)
'    WFINDOI As String * 1           ' WF検査指示（Oi)
'    WFINDB1 As String * 1           ' WF検査指示（B1)
'    WFINDB2 As String * 1           ' WF検査指示（B2）
'    WFINDB3 As String * 1           ' WF検査指示（B3)
'    WFINDL1 As String * 1           ' WF検査指示（L1)
'    WFINDL2 As String * 1           ' WF検査指示（L2)
'    WFINDL3 As String * 1           ' WF検査指示（L3)
'    WFINDL4 As String * 1           ' WF検査指示（L4)
'    WFINDDS As String * 1           ' WF検査指示（DS)
'    WFINDDZ As String * 1           ' WF検査指示（DZ)
'    WFINDSP As String * 1           ' WF検査指示（SP)
'    WFINDDO1 As String * 1          ' WF検査指示（DO1)
'    WFINDDO2 As String * 1          ' WF検査指示（DO2)
'    WFINDDO3 As String * 1          ' WF検査指示（DO3)
'    WFINDOT1 As String * 1          ' WF検査指示（OT1)  'Add.03/05/20
'    WFINDOT2 As String * 1          ' WF検査指示（OT2)  'Add.03/05/20
'    WFRESRS As String * 1           ' WF検査実績（Rs)
'    WFRESOI As String * 1           ' WF検査実績（Oi)
'    WFRESB1 As String * 1           ' WF検査実績（B1)
'    WFRESB2 As String * 1           ' WF検査実績（B2）
'    WFRESB3 As String * 1           ' WF検査実績（B3)
'    WFRESL1 As String * 1           ' WF検査実績（L1)
'    WFRESL2 As String * 1           ' WF検査実績（L2)
'    WFRESL3 As String * 1           ' WF検査実績（L3)
'    WFRESL4 As String * 1           ' WF検査実績（L4)
'    WFRESDS As String * 1           ' WF検査実績（DS)
'    WFRESDZ As String * 1           ' WF検査実績（DZ)
'    WFRESSP As String * 1           ' WF検査実績（SP)
'    WFRESDO1 As String * 1          ' WF検査実績（DO1)
'    WFRESDO2 As String * 1          ' WF検査実績（DO2)
'    WFRESDO3 As String * 1          ' WF検査実績（DO3)
'    WFRESOT1 As String * 1          ' WF検査実績（OT1)  'Add.03/05/20
'    WFRESOT2 As String * 1          ' WF検査実績（OT2)  'Add.03/05/20
'    REGDATE As Date                 ' 登録日付
'    UPDDATE As Date                 ' 更新日付
'    SENDFLAG As String * 1          ' 送信フラグ
'    SENDDATE As Date                ' 送信日付
'    BkIngotPos  As Integer          ' add 2003/03/28 hitec)matsumoto
'End Type

''新サンプル管理(SXL)
Public Type typ_XSDCW
    SXLIDCW As String * 13          'SXLID
    SMPKBNCW As String * 1          'サンプル区分
    TBKBNCW As String * 1           'T/B区分
    REPSMPLIDCW As String * 16      '代表サンプルID
    XTALCW As String * 12           '結晶番号
    INPOSCW As Integer              '結晶内位置
    HINBCW As String * 8            '品番
    REVNUMCW As Integer             '製品番号改訂番号
    FACTORYCW As String * 1         '工場
    OPECW As String * 1             '操業条件
    KTKBNCW As String * 1           '確定区分
    SMCRYNUMCW As String * 12       'サンプルブロックID
    WFSMPLIDRSCW As String * 16     'サンプルID(Rs)
    WFSMPLIDRS1CW As String * 16    '推定サンプルID1(Rs)
    WFSMPLIDRS2CW As String * 16    '推定サンプルID2(Rs)
    WFINDRSCW As String * 1         '状態FLG(Rs)
    WFRESRS1CW As String * 1        '実績FLG1(Rs)
    WFRESRS2CW As String * 1        '実績FLG2(Rs)
    WFSMPLIDOICW As String * 16     'サンプルID(Oi)
    WFINDOICW As String * 1         '状態FLG(Oi)
    WFRESOICW As String * 1         '実績FLG(Oi)
    WFSMPLIDB1CW As String * 16     'サンプルID(B1)
    WFINDB1CW As String * 1         '状態FLG(B1)
    WFRESB1CW As String * 1         '実績FLG(B1)
    WFSMPLIDB2CW As String * 16     'サンプルID(B2)
    WFINDB2CW As String * 1         '状態FLG(B2)
    WFRESB2CW As String * 1         '実績FLG(B2)
    WFSMPLIDB3CW As String * 16     'サンプルID(B3)
    WFINDB3CW As String * 1         '状態FLG(B3)
    WFRESB3CW As String * 1         '実績FLG(B3)
    WFSMPLIDL1CW As String * 16     'サンプルID(L1)
    WFINDL1CW As String * 1         '状態FLG(L1)
    WFRESL1CW As String * 1         '実績FLG(L1)
    WFSMPLIDL2CW As String * 16     'サンプルID(L2)
    WFINDL2CW As String * 1         '状態FLG(L2)
    WFRESL2CW As String * 1         '実績FLG(L2)
    WFSMPLIDL3CW As String * 16     'サンプルID(L3)
    WFINDL3CW As String * 1         '状態FLG(L3)
    WFRESL3CW As String * 1         '実績FLG(L3)
    WFSMPLIDL4CW As String * 16     'サンプルID(L4)
    WFINDL4CW As String * 1         '状態FLG(L4)
    WFRESL4CW As String * 1         '実績FLG(L4)
    WFSMPLIDDSCW As String * 16     'サンプルID(DS)
    WFINDDSCW As String * 1         '状態FLG(DS)
    WFRESDSCW As String * 1         '実績FLG(DS)
    WFSMPLIDDZCW As String * 16     'サンプルID(DZ)
    WFINDDZCW As String * 1         '状態FLG(DZ)
    WFRESDZCW As String * 1         '実績FLG(DZ)
    WFSMPLIDSPCW As String * 16     'サンプルID(SP)
    WFINDSPCW As String * 1         '状態FLG(SP)
    WFRESSPCW As String * 1         '実績FLG(SP)
    WFSMPLIDDO1CW As String * 16    'サンプルID(DO1)
    WFINDDO1CW As String * 1        '状態FLG(DO1)
    WFRESDO1CW As String * 1        '実績FLG(DO1)
    WFSMPLIDDO2CW As String * 16    'サンプルID(DO2)
    WFINDDO2CW As String * 1        '状態FLG(DO2)
    WFRESDO2CW As String * 1        '実績FLG(DO2)
    WFSMPLIDDO3CW As String * 16    'サンプルID(DO3)
    WFINDDO3CW As String * 1        '状態FLG(DO3)
    WFRESDO3CW As String * 1        '実績FLG(DO3)
    WFSMPLIDOT1CW As String * 16    'サンプルID(OT1)
    WFINDOT1CW As String * 1        '状態FLG(OT1)
    WFRESOT1CW As String * 1        '実績FLG(OT1)
    WFSMPLIDOT2CW As String * 16    'サンプルID(OT2)
    WFINDOT2CW As String * 1        '状態FLG(OT2)
    WFRESOT2CW As String * 1        '実績FLG(OT2)
    WFSMPLIDAOICW As String * 16    'サンプルID(AOi)
    WFINDAOICW As String * 1        '状態FLG(AOi)
    WFRESAOICW As String * 1        '実績FLG(AOi)
    SMPLNUMCW As Integer            'サンプル枚数
    SMPLPATCW As String * 1         'サンプルパターン
    LIVKCW As String * 1            '生死区分　'追加 2003/10/04
    TSTAFFCW As String * 8          '登録社員ID
    TDAYCW As Date                  '登録日付
    KSTAFFCW As String * 8          '更新社員ID
    KDAYCW As Date                  '更新日付
    SNDKCW As String * 1            '送信フラグ
    SNDDAYCW As Date                '送信日付
    WFSMPLIDGDCW As String * 16     'サンプルID(GD)     '05/01/17 ooba START =======>
    WFINDGDCW As String * 1         '状態FLG(GD)
    WFRESGDCW As String * 1         '実績FLG(GD)
    WFHSGDCW As String * 1          '保証FLG(GD)        '05/01/17 ooba END =========>
'    BkIngotPos  As Integer          ' add 2003/03/28 hitec)matsumoto
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    EPSMPLIDB1CW As String * 16     'サンプルID(BMD1)
    EPINDB1CW As String * 1         '状態FLG(BMD1)
    EPRESB1CW As String * 1         '実績FLG(BMD1)
    EPSMPLIDB2CW As String * 16     'サンプルID(BMD2)
    EPINDB2CW As String * 1         '状態FLG(BMD2)
    EPRESB2CW As String * 1         '実績FLG(BMD2)
    EPSMPLIDB3CW As String * 16     'サンプルID(BMD3)
    EPINDB3CW As String * 1         '状態FLG(BMD3)
    EPRESB3CW As String * 1         '実績FLG(BMD3)
    EPSMPLIDL1CW As String * 16     'サンプルID(OSF1)
    EPINDL1CW As String * 1         '状態FLG(OSF1)
    EPRESL1CW As String * 1         '実績FLG(OSF1)
    EPSMPLIDL2CW As String * 16     'サンプルID(OSF2)
    EPINDL2CW As String * 1         '状態FLG(OSF2)
    EPRESL2CW As String * 1         '実績FLG(OSF2)
    EPSMPLIDL3CW As String * 16     'サンプルID(OSF3)
    EPINDL3CW As String * 1         '状態FLG(OSF3)
    EPRESL3CW As String * 1         '実績FLG(OSF3)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
End Type


' 切断指示
Public Type typ_TBCME045
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' 結晶内開始位置
    TRANCNT As Integer              ' 処理回数
    Length As Integer               ' 長さ
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    StaffID As String * 8           ' 社員ID
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 品番製品番号改訂番号
    FACTORY As String * 1           ' 品番工場
    OPECOND As String * 1           ' 品番操業条件
    BDCAUS As String * 3            ' 区分コード
    STATCLS As String * 1           ' 状態区分
    BLOCKID As String * 12          ' ブロックID
    CRYINDRS As String * 1          ' 状態FLG（Rs)
    CRYINDOI As String * 1          ' 状態FLG（Oi)
    CRYINDB1 As String * 1          ' 状態FLG（B1)
    CRYINDB2 As String * 1          ' 状態FLG（B2）
    CRYINDB3 As String * 1          ' 状態FLG（B3)
    CRYINDL1 As String * 1          ' 状態FLG（L1)
    CRYINDL2 As String * 1          ' 状態FLG（L2)
    CRYINDL3 As String * 1          ' 状態FLG（L3)
    CRYINDL4 As String * 1          ' 状態FLG（L4)
    CRYINDCS As String * 1          ' 状態FLG（Cs)
    CRYINDGD As String * 1          ' 状態FLG（GD)
    CRYINDT As String * 1           ' 状態FLG（T)
    CRYINDEP As String * 1          ' 状態FLG（EPD)
    PRIORITY As String * 1          ' 優先度
    PALTNUM As String * 4           ' パレット番号
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 多結晶受入実績
Public Type typ_TBCMG001
    MTRLNUM As String * 10          ' 原料番号
    JDATE As Date                   ' 日付
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    MTRLTYPE As String * 3          ' 原料種類
    MAKERNO As String * 6           ' メーカ管理No
    RVWEIGHT As Long                ' 受入購入重量
    CRYCOMMENT As String            ' コメント
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ＩＤ
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 購入単結晶受入実績
Public Type typ_TBCMG002
    CRYNUM As String * 12           ' 結晶番号
    TRANCNT As Integer              ' 処理回数
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    REPCCL As String * 1            ' 受入取消区分
    RBATCHNO As String * 10         ' 炉バッチＮｏ
    DMTOP1 As Integer               ' 直径ＴＯＰ１
    DMTOP2 As Integer               ' 直径ＴＯＰ２
    DMTAIL1 As Integer              ' 直径ＴＡＩＬ１
    DMTAIL2 As Integer              ' 直径ＴＡＩＬ２
    NCHDPTH1 As Integer             ' ノッチ深さ１
    NCHDPTH2 As Integer             ' ノッチ深さ２
    UPLENGTH As Integer             ' 引上げ長
    SXLPOS As Integer               ' ＳＸＬ位置
    BlkLen As Integer               ' ブロック長さ
    BLKWGHT As Long                 ' ブロック重量
    CMPTOP1 As Double               ' 比抵抗TOP　１
    CMPTOP2 As Double               ' 比抵抗TOP　２
    CMPTOP3 As Double               ' 比抵抗TOP　３
    CMPTOP4 As Double               ' 比抵抗TOP　４
    CMPTOP5 As Double               ' 比抵抗TOP　５
    CMPTOPR As Double               ' 比抵抗TOP　RRG
    CMPTAIL1 As Double              ' 比抵抗TAIL　１
    CMPTAIL2 As Double              ' 比抵抗TAIL　２
    CMPTAIL3 As Double              ' 比抵抗TAIL　３
    CMPTAIL4 As Double              ' 比抵抗TAIL　４
    CMPTAIL5 As Double              ' 比抵抗TAIL　５
    CMPTAILR As Double              ' 比抵抗TAIL　RRG
    OITOP1 As Double                ' Oi　TOP　１
    OITOP2 As Double                ' Oi　TOP　２
    OITOP3 As Double                ' Oi　TOP　３
    OITOP4 As Double                ' Oi　TOP　４
    OITOP5 As Double                ' Oi　TOP　５
    OITOPR As Double                ' Oi　TOP　ROG
    OITAIL1 As Double               ' Oi　TAIL　１
    OITAIL2 As Double               ' Oi　TAIL　２
    OITAIL3 As Double               ' Oi　TAIL　３
    OITAIL4 As Double               ' Oi　TAIL　４
    OITAIL5 As Double               ' Oi　TAIL　５
    OITAILR As Double               ' Oi　TAIL　ROG
    CSTOP As Double                 ' Cs　TOP
    CSTAIL As Double                ' Cs　TAIL
    LD1TOPMX As Double              ' LD-1　TOP　MAX
    LD1TOPAV As Double              ' LD-1　TOP　AVE
    LD1TAILM As Double              ' LD-1　TAIL　MAX
    LD1TAILA As Double              ' LD-1　TAIL　AVE
    LD2TOPMM As Double              ' LD-2　TOP　MAX
    LD2TOPAV As Double              ' LD-2　TOP　AVE
    LD2TAILM As Double              ' LD-2　TAIL　MAX
    LD2TAILA As Double              ' LD-2　TAIL　AVE
    BMDTOPMX As Double              ' BMD　TOP　MAX
    BMDTOPAV As Double              ' BMD　TOP　AVE
    BMDTAILM As Double              ' BMD　TAIL　MAX
    BMDTAILA As Double              ' BMD　TAIL　AVE
    GD1TOP As Integer               ' GD1 TOP
    GD1TAIL As Integer              ' GD1 TAIL
    GD2TOP As Integer               ' GD2 TOP
    GD2TAIL As Integer              ' GD2 TAIL
    DIA1TOP As Integer              ' DIA1 TOP
    DIA1TAIL As Integer             ' DIA1 TAIL
    DIA2TOP As Integer              ' DIA2 TOP
    DIA2TAIL As Integer             ' DIA2 TAIL
    LTFTOP As Integer               ' LIFETIME from TOP
    LTFTAIL As Integer              ' LIFETIME from TAIL
    EPD As Integer                  ' EPD
    HCNO As String * 10             ' 発注No
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' リメルト受入切断実績
Public Type typ_TBCMG003
    CRYNUM As String * 12           ' 結晶番号
    ROCLASS As String * 3           ' ρ区分
    HRCLASS As String * 1           ' 廃棄・ロス区分
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    Weight As Long                  ' 重量
    RMSHAPE As String * 1           ' リメルト形状
    RMMTRLNUM As String * 10        ' リメルト原料番号
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' リメルト洗浄払出実績
Public Type typ_TBCMG004
    MTRLNUM As String * 10          ' 原料番号
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    DRWEIGHT As Long                ' 乾燥後重量
    LSWEIGHT As Long                ' ロス重量
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SUMITSENDFLAG As String * 1     ' SUMIT送信フラグ
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 原料在庫管理
Public Type typ_TBCMG005
    MTRLNUM As String * 10          ' 原料番号
    USABLCLS As String * 1          ' 使用可能区分
    Weight As Long                  ' 重量
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
End Type


' 原料在庫実績
Public Type typ_TBCMG006
    MTRLNUM As String * 10          ' 原料番号
    TRANCNT As Long                 ' 処理回数
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    Class As String * 1             ' 区分
    INWEIGHT As Long                ' 入力重量
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' クリスタルカタログ受入実績
Public Type typ_TBCMG007
    CRYNUM As String * 12           ' 結晶番号
    TRANCNT As Integer              ' 処理回数
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    BDCODE As String * 3            ' 不良理由コード
    PALTNUM As String * 4           ' パレット番号
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 格上げ実績
Public Type typ_TBCMG008
    CRYNUM As String * 12           ' 結晶番号（格上げ）
    TRANCNT As Integer              ' 処理回数
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    NHINBAN As String * 8           ' 新品番
    NMNOREVNO As Integer            ' 新製品番号改訂番号
    NFACTORY As String * 1          ' 新工場
    NOPECOND As String * 1          ' 新操業条件
    OHINBAN As String * 8           ' 旧品番
    OMNOREVNO As Integer            ' 旧製品番号改訂番号
    OFACTORY As String * 1          ' 旧工場
    OOPECOND As String * 1          ' 旧操業条件
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 引上げ指示実績
Public Type typ_TBCMH001
    UPINDNO As String * 9           ' 引上げ指示No.
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    MODEL As String * 4             ' 機種
    GOUKI As String * 3             ' 号機
    PGID As String * 10             ' PG-ID
    CPORGIND As String * 12         ' 複写元指示No
    hinban As String * 8            ' 品番
    NMNOREVNO As Integer            ' 製品番号改訂番号
    NFACTORY As String * 1          ' 工場
    NOPECOND As String * 1          ' 操業条件
    NUMNOTE1 As String              ' 品番備考１
    NUMNOTE2 As String              ' 品番備考２
    SEED As String * 4              ' シード
    SEKIERTB As String * 7          ' 石英ルツボ
    DPNTCLS As String * 7           ' ドーパント種類
    DOPANT As Double                ' ドーパント量
    AMRESIST As Double              ' ねらい抵抗
    CRYDOPCL As String * 7          ' 結晶ドープ種類
    CRYDOPVL As Double              ' 結晶ドープ量
    UPBTCHNM As Integer             ' 引上げバッチ数
    ADDDOPCL As String * 7          ' 追加ドーパント種類
    ADDDOPVL As Double              ' 追加ドーパント量
    ADDDOPPT As Integer             ' 追加ドーパント位置
    BCNT1COD As String * 3          ' バッチ備考1（コード）
    BCNT1CMT As String              ' バッチ備考1（ｺﾒﾝﾄ）
    BCNT2COD As String * 3          ' バッチ備考2（コード）
    BCNT2CMT As String              ' バッチ備考2（ｺﾒﾝﾄ）
    MTCLS1 As String * 3            ' 原料種類1
    MTWGHT1 As Long                 ' 原料重量1
    ESWGHT1 As Long                 ' 推定残重量1
    MTCLS2 As String * 3            ' 原料種類2
    MTWGHT2 As Long                 ' 原料重量2
    ESWGHT2 As Long                 ' 推定残重量2
    MTCLS3 As String * 3            ' 原料種類3
    MTWGHT3 As Long                 ' 原料重量3
    ESWGHT3 As Long                 ' 推定残重量3
    MTCLS4 As String * 3            ' 原料種類4
    MTWGHT4 As Long                 ' 原料重量4
    ESWGHT4 As Long                 ' 推定残重量4
    MTCLS5 As String * 3            ' 原料種類5
    MTWGHT5 As Long                 ' 原料重量5
    ESWGHT5 As Long                 ' 推定残重量5
    MTCLS6 As String * 3            ' 原料種類6
    MTWGHT6 As Long                 ' 原料重量6
    ESWGHT6 As Long                 ' 推定残重量6
    MTCLS7 As String * 3            ' 原料種類7
    MTWGHT7 As Long                 ' 原料重量7
    ESWGHT7 As Long                 ' 推定残重量7
    MTCLS8 As String * 3            ' 原料種類8
    MTWGHT8 As Long                 ' 原料重量8
    ESWGHT8 As Long                 ' 推定残重量8
    MTCLS9 As String * 3            ' 原料種類9
    MTWGHT9 As Long                 ' 原料重量9
    ESWGHT9 As Long                 ' 推定残重量9
    MTCLS10 As String * 3           ' 原料種類10
    MTWGHT10 As Long                ' 原料重量10
    ESWGHT10 As Long                ' 推定残重量10
    MTCLS11 As String * 3           ' 原料種類11
    MTWGHT11 As Long                ' 原料重量11
    ESWGHT11 As Long                ' 推定残重量11
    MTCLS12 As String * 3           ' 原料種類12
    MTWGHT12 As Long                ' 原料重量12
    ESWGHT12 As Long                ' 推定残重量12
    MTCLS13 As String * 3           ' 原料種類13
    MTWGHT13 As Long                ' 原料重量13
    ESWGHT13 As Long                ' 推定残重量13
    MTCLS14 As String * 3           ' 原料種類14
    MTWGHT14 As Long                ' 原料重量14
    ESWGHT14 As Long                ' 推定残重量14
    MTCLS15 As String * 3           ' 原料種類15
    MTWGHT15 As Long                ' 原料重量15
    ESWGHT15 As Long                ' 推定残重量15
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 引上げ投入実績
Public Type typ_TBCMH002
    UPINDNO As String * 9           ' 引上げ指示No
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    TRANCLS As String * 1           ' 処理区分
    PGID As String * 8              ' PG-ID
    TYPE As String * 2              ' タイプ
    SEKIERTB As String * 10         ' 石英ルツボ
    DPNTCLS As String * 7           ' ドーパント種類
    DOPANT As Double                ' ドーパント量
    CRYDOP As String * 1            ' 結晶ドープ
    CRYDOPVL As Double              ' 結晶ドープ量
    SEED As String * 4              ' シード
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SUMMITSENDFLAG As String * 1    ' SUMMIT送信フラグ
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 引上げ投入原料実績
Public Type typ_TBCMH003
    UPINDNO As String * 9           ' 引上げ指示No
    MTRLNUM As String * 10          ' 原料番号
    Weight As Long                  ' 重量
    ESWEIGHT As Long                ' 推定残重量
End Type


' 引上げ終了実績
Public Type typ_TBCMH004
    CRYNUM As String * 12           ' 結晶番号
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    LENGTOP As Integer              ' 長さ（TOP）
    LENGTKDO As Integer             ' 長さ（直胴）
    LENGTAIL As Integer             ' 長さ（TAIL）
    LENGFREE As Integer             ' フリー長さ
    DM1 As Double                   ' 直胴直径１
    DM2 As Double                   ' 直胴直径２
    DM3 As Double                   ' 直胴直径３
    WGHTTOP As Long                 ' 重量（TOP）
    WGHTTKDO As Long                ' 重量（直胴）
    WGHTTAIL As Long                ' 重量（TAIL)
    WGHTFREE As Long                ' 重量（フリー長さ）
    WGTOPCUT As Long                ' トップカット重量
    UPWEIGHT As Long                ' 引上げ重量
    CHARGE As Long                  ' チャージ量
    SEED As String * 4              ' シード
    STATCLS As String * 3           ' BOT状況区分
    JDGECODE As String * 3          ' 判定コード
    PWTIME As Double                ' パワー時間
    ADDDPPOS As Integer             ' 追加ドープ位置
    ADDDPCLS As String * 7          ' 追加ドーパント種類
    ADDDPVAL As Double              ' 追加ドープ量
    ADDDPNAM As String              ' 追加ドープ名
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SUMMITSENDFLAG As String * 1    ' SUMMIT送信フラグ
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
    PULENTKC1 As Long               ' 引上直胴長さ  2010/09/06 add Kameda
    PUWGHTTKC1 As Long              ' 引上直胴重量  2010/09/06 add Kameda
End Type


' 引上げ終了残重量実績
Public Type typ_TBCMH005
    RSCRYNUM As String * 12         ' 残重量結晶番号
    CRYNUM As String * 9            ' 元結晶番号
    RSWEIGHT As Long                ' 残重量
End Type


' 引上機種号機情報管理
Public Type typ_TBCMH006
    MODEL As String * 4             ' 機種
    GOUKI As String * 3             ' 号機
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    PROCDATE As Date                ' 日付
    CRYNUM As String * 12           ' 結晶番号
    hinban As String * 8            ' 品番
    PGID As String * 10             ' PG-ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
End Type


' 加工払出実績
Public Type typ_TBCMI001
    CRYNUM As String * 12           ' 結晶番号
    TRANCNT As Integer              ' 処理回数
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    UPLENGTH As Integer             ' 引上げ長さ
    FREELENG As Integer             ' フリー長
    UPWEIGHT As Long                ' 引上げ重量
    SEED As String * 4              ' シード
    PRCMCN As String * 1            ' 研削機
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    SUMMITSENDFLAG As String * 1    ' SUMMIT送信フラグ
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
    TOPSFTQTY As String * 4         ' トップシフト量  add 06/11/13 SET/Miyazaki
    BOTSFTQTY As String * 4         ' ボトムシフト量  add 06/11/13 SET/Miyazaki
End Type


' 研削加工実績
Public Type typ_TBCMI002
    CRYNUM As String * 12           ' 結晶番号
    TRANCNT As Integer              ' 処理回数
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    DMTOP1 As Double                ' 直径TOP１
    DMTOP2 As Double                ' 直径TOP２
    DMTAIL1 As Double               ' 直径TAIL１
    DMTAIL2 As Double               ' 直径TAIL２
    NCHPOS As String * 2            ' ノッチ位置
    NCHDPTH As Double               ' ノッチ深さ
    NCHWIDTH As Double              ' ノッチ幅
    BDLNTOP As Integer              ' 不良長さ（TOP）
    BDCDTOP As String * 3           ' 不良判定コード（TOP）
    BDLNTAIL As Integer             ' 不良長さ（TAIL)
    BDCDTAIL As String * 3          ' 不良判定コード（TAIL）
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
    INGOTPOS As Integer             ' インゴット内位置
    Length As Integer               ' 長さ
    GOUKI As String * 5             ' 号機             2003/06/12 osawa
    NCHWTAIL As Double              ' ノッチ深さ(TAIL) 2004/05/25
    BLOCKID As String * 12          ' ﾌﾞﾛｯｸID          2006/02/01 tuku
    CYGRTIM As String * 6           ' 研削時間         2006/11/09 SETsw J.W
    NCHLENGTH As String * 4         ' ノッチ長さ       2006/11/09 SETsw J.W
    SOPROCTIM As String * 6         ' 加工時間(粗研)   2006/11/20 SETsw Y.M
    SEIPROCTIM As String * 6        ' 加工時間(精研)   2006/11/20 SETsw Y.M
    NCHANGLE As Double              ' ノッチ角度　     2009/09    SUMCO Akizuki
End Type


' 切断実績
Public Type typ_TBCMI003
    CRYNUM As String * 12           ' 結晶番号
    TRANCNT As Integer              ' 処理回数
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
    GOUKI As String                 ' 号機　　3/13 Yam
    BLOCKID As String * 12          ' ﾌﾞﾛｯｸID  2006/02/01 tuku
    cutNum As String * 2            ' カット数
    CUTNUM2 As String * 4           ' カット数2 2006/11/02 SETsw
    PROCTIM As String * 6           ' 加工時間  2006/11/02 SETsw
End Type


' EPD実績
Public Type typ_TBCMJ001
    CRYNUM As String * 12           ' 結晶番号
    POSITION As Integer             ' 位置
    SMPKBN As String * 1            ' サンプル区分
    TRANCOND As String * 1          ' 処理条件
    TRANCNT As Integer              ' 処理回数
    SMPLNO As Long                  ' サンプルＮｏ      Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' サンプル有無
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    GOUKI As String * 3             ' 号機
    MEASURE As Integer              ' 測定値
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 結晶抵抗実績
Public Type typ_TBCMJ002
    CRYNUM As String * 12           ' 結晶番号
    POSITION As Integer             ' 位置
    SMPKBN As String * 1            ' サンプル区分
    TRANCOND As String * 1          ' 処理条件
    TRANCNT As Integer              ' 処理回数
    SMPLNO As Long                  ' サンプルＮｏ      Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' サンプル有無
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    GOUKI As String * 3             ' 号機
    TYPE As String * 1              ' タイプ
    MEAS1 As Double                 ' 測定値１
    MEAS2 As Double                 ' 測定値２
    MEAS3 As Double                 ' 測定値３
    MEAS4 As Double                 ' 測定値４
    MEAS5 As Double                 ' 測定値５
    JMEAS1 As Double                ' 測定値１
    JMEAS2 As Double                ' 測定値２
    JMEAS3 As Double                ' 測定値３
    JMEAS4 As Double                ' 測定値４
    JMEAS5 As Double                ' 測定値５
    EFEHS As Double                 ' 実効偏析
    RRG As Double                   ' ＲＲＧ
    JudgData As Double              ' 検索対象値
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
    SUIFLG  As String               ' 推定FLG
    LTDATA As String                ' 同部位のLT実測値
    KANSANCHI As String             ' 10Ω換算値
End Type


' Ｏｉ実績
Public Type typ_TBCMJ003
    CRYNUM As String * 12           ' 結晶番号
    POSITION As Integer             ' 位置
    SMPKBN As String * 1            ' サンプル区分
    TRANCOND As String * 1          ' 処理条件
    TRANCNT As Integer              ' 処理回数
    SMPLNO As Long                  ' サンプルＮｏ      Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' サンプル有無
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    GOUKI As String * 3             ' 号機
    OIMEAS1 As Double               ' Ｏｉ測定値１
    OIMEAS2 As Double               ' Ｏｉ測定値２
    OIMEAS3 As Double               ' Ｏｉ測定値３
    OIMEAS4 As Double               ' Ｏｉ測定値４
    OIMEAS5 As Double               ' Ｏｉ測定値５
    ORGRES As Double                ' ＯＲＧ結果
    SETDTM As Date                  ' 設定日時
    EFFECTTM As Integer             ' 有効時間
    FTIRMETH As String              ' ＦＴＩＲ相関式
    YCOEF As Double                 ' ＦＴＩＲ換算式（Ｙ切片）
    XCOEF As Double                 ' ＦＴＩＲ換算式（Ｘ係数）
    AVE As Double                   ' ＡＶＥ
    SIGMA As Double                 ' σ（シグマ）
    FTIRCONV As Double              ' ＦＴＩＲ換算
    INSPECTWAY As String * 2        ' 検査方法
    JudgData As Double              ' 検索対象値
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' Cs実績
Public Type typ_TBCMJ004
    CRYNUM As String * 12           ' 結晶番号
    POSITION As Integer             ' 位置
    SMPKBN As String * 1            ' サンプル区分
    TRANCOND As String * 1          ' 処理条件
    TRANCNT As Integer              ' 処理回数
    SMPLNO As Long                  ' サンプルＮｏ      Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' サンプル有無
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    GOUKI As String * 3             ' 号機
    CSMEAS As Double                ' Cs実測値
    PRE70P As Double                ' ７０％推定値
    INSPECTWAY As String * 2        ' 検査方法
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' ＯＳＦ実績
Public Type typ_TBCMJ005
    CRYNUM As String * 12           ' 結晶番号
    POSITION As Integer             ' 位置
    SMPKBN As String * 1            ' サンプル区分
    TRANCOND As String * 1          ' 処理条件
    TRANCNT As Integer              ' 処理回数
    SMPLNO As Long                  ' サンプルＮｏ      Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' サンプル有無
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    GOUKI As String * 3             ' 号機
    MEASMETH As String * 1          ' 測定方法
    MEASSPOT As Integer             ' 測定点
    MAG As String * 4               ' 倍率
    HTPRC As String * 2             ' 熱処理方法
    KKSP As String * 3              ' 結晶欠陥測定位置
    KKSET As String * 3             ' 結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    CALCMAX As Double               ' 計算結果 Max
    CALCAVE As Double               ' 計算結果 Ave
    MEAS1 As Integer                ' 測定値１
    MEAS2 As Integer                ' 測定値２
    MEAS3 As Integer                ' 測定値３
    MEAS4 As Integer                ' 測定値４
    MEAS5 As Integer                ' 測定値５
    MEAS6 As Integer                ' 測定値６
    MEAS7 As Integer                ' 測定値７
    MEAS8 As Integer                ' 測定値８
    MEAS9 As Integer                ' 測定値９
    MEAS10 As Integer               ' 測定値１０
    MEAS11 As Integer               ' 測定値１１
    MEAS12 As Integer               ' 測定値１２
    MEAS13 As Integer               ' 測定値１３
    MEAS14 As Integer               ' 測定値１４
    MEAS15 As Integer               ' 測定値１５
    MEAS16 As Integer               ' 測定値１６
    MEAS17 As Integer               ' 測定値１７
    MEAS18 As Integer               ' 測定値１８
    MEAS19 As Integer               ' 測定値１９
    MEAS20 As Integer               ' 測定値２０
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
'OSF，BMD項目追加対応  2002.04.02 yakimura
    OSFPOS1 As Double               ' ﾊﾟﾀｰﾝ区分１位置
    OSFWID1 As Double               ' ﾊﾟﾀｰﾝ区分１幅
    OSFRD1  As String               ' ﾊﾟﾀｰﾝ区分１R/D
    OSFPOS2 As Double               ' ﾊﾟﾀｰﾝ区分２位置
    OSFWID2 As Double               ' ﾊﾟﾀｰﾝ区分２幅
    OSFRD2  As String               ' ﾊﾟﾀｰﾝ区分２R/D
    OSFPOS3 As Double               ' ﾊﾟﾀｰﾝ区分３位置
    OSFWID3 As Double               ' ﾊﾟﾀｰﾝ区分３幅
    OSFRD3  As String               ' ﾊﾟﾀｰﾝ区分３R/D
'OSF，BMD項目追加対応  2002.04.02 yakimura
'スポットMAX追加 K.Goto 2006/03/31
    SPOTMAX As Long              ' スポットMAX
    
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    CALCMH  As Double               ' 面内比(MAX/MIN)
    PTNJUDGRES  As String * 1       ' パターン判定結果
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
End Type


' ＧＤ実績
Public Type typ_TBCMJ006
    CRYNUM As String * 12           ' 結晶番号
    POSITION As Integer             ' 位置
    SMPKBN As String * 1            ' サンプル区分
    TRANCOND As String * 1          ' 処理条件
    TRANCNT As Integer              ' 処理回数
    SMPLNO As Long                  ' サンプルＮｏ      Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' サンプル有無
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    GOUKI As String * 3             ' 号機
    MSRSDEN As Integer              ' 測定結果 Den
    MSRSLDL As Integer              ' 測定結果 L/DL
    MSRSDVD2 As Integer             ' 測定結果 DVD2
    MS01LDL1 As Integer             ' 測定値01 L/DL1
    MS01LDL2 As Integer             ' 測定値01 L/DL2
    MS01LDL3 As Integer             ' 測定値01 L/DL3
    MS01LDL4 As Integer             ' 測定値01 L/DL4
    MS01LDL5 As Integer             ' 測定値01 L/DL5
    MS01DEN1 As Integer             ' 測定値01 Den1
    MS01DEN2 As Integer             ' 測定値01 Den2
    MS01DEN3 As Integer             ' 測定値01 Den3
    MS01DEN4 As Integer             ' 測定値01 Den4
    MS01DEN5 As Integer             ' 測定値01 Den5
    MS02LDL1 As Integer             ' 測定値02 L/DL1
    MS02LDL2 As Integer             ' 測定値02 L/DL2
    MS02LDL3 As Integer             ' 測定値02 L/DL3
    MS02LDL4 As Integer             ' 測定値02 L/DL4
    MS02LDL5 As Integer             ' 測定値02 L/DL5
    MS02DEN1 As Integer             ' 測定値02 Den1
    MS02DEN2 As Integer             ' 測定値02 Den2
    MS02DEN3 As Integer             ' 測定値02 Den3
    MS02DEN4 As Integer             ' 測定値02 Den4
    MS02DEN5 As Integer             ' 測定値02 Den5
    MS03LDL1 As Integer             ' 測定値03 L/DL1
    MS03LDL2 As Integer             ' 測定値03 L/DL2
    MS03LDL3 As Integer             ' 測定値03 L/DL3
    MS03LDL4 As Integer             ' 測定値03 L/DL4
    MS03LDL5 As Integer             ' 測定値03 L/DL5
    MS03DEN1 As Integer             ' 測定値03 Den1
    MS03DEN2 As Integer             ' 測定値03 Den2
    MS03DEN3 As Integer             ' 測定値03 Den3
    MS03DEN4 As Integer             ' 測定値03 Den4
    MS03DEN5 As Integer             ' 測定値03 Den5
    MS04LDL1 As Integer             ' 測定値04 L/DL1
    MS04LDL2 As Integer             ' 測定値04 L/DL2
    MS04LDL3 As Integer             ' 測定値04 L/DL3
    MS04LDL4 As Integer             ' 測定値04 L/DL4
    MS04LDL5 As Integer             ' 測定値04 L/DL5
    MS04DEN1 As Integer             ' 測定値04 Den1
    MS04DEN2 As Integer             ' 測定値04 Den2
    MS04DEN3 As Integer             ' 測定値04 Den3
    MS04DEN4 As Integer             ' 測定値04 Den4
    MS04DEN5 As Integer             ' 測定値04 Den5
    MS05LDL1 As Integer             ' 測定値05 L/DL1
    MS05LDL2 As Integer             ' 測定値05 L/DL2
    MS05LDL3 As Integer             ' 測定値05 L/DL3
    MS05LDL4 As Integer             ' 測定値05 L/DL4
    MS05LDL5 As Integer             ' 測定値05 L/DL5
    MS05DEN1 As Integer             ' 測定値05 Den1
    MS05DEN2 As Integer             ' 測定値05 Den2
    MS05DEN3 As Integer             ' 測定値05 Den3
    MS05DEN4 As Integer             ' 測定値05 Den4
    MS05DEN5 As Integer             ' 測定値05 Den5
    MS06LDL1 As Integer             ' 測定値06 L/DL1
    MS06LDL2 As Integer             ' 測定値06 L/DL2
    MS06LDL3 As Integer             ' 測定値06 L/DL3
    MS06LDL4 As Integer             ' 測定値06 L/DL4
    MS06LDL5 As Integer             ' 測定値06 L/DL5
    MS06DEN1 As Integer             ' 測定値06 Den1
    MS06DEN2 As Integer             ' 測定値06 Den2
    MS06DEN3 As Integer             ' 測定値06 Den3
    MS06DEN4 As Integer             ' 測定値06 Den4
    MS06DEN5 As Integer             ' 測定値06 Den5
    MS07LDL1 As Integer             ' 測定値07 L/DL1
    MS07LDL2 As Integer             ' 測定値07 L/DL2
    MS07LDL3 As Integer             ' 測定値07 L/DL3
    MS07LDL4 As Integer             ' 測定値07 L/DL4
    MS07LDL5 As Integer             ' 測定値07 L/DL5
    MS07DEN1 As Integer             ' 測定値07 Den1
    MS07DEN2 As Integer             ' 測定値07 Den2
    MS07DEN3 As Integer             ' 測定値07 Den3
    MS07DEN4 As Integer             ' 測定値07 Den4
    MS07DEN5 As Integer             ' 測定値07 Den5
    MS08LDL1 As Integer             ' 測定値08 L/DL1
    MS08LDL2 As Integer             ' 測定値08 L/DL2
    MS08LDL3 As Integer             ' 測定値08 L/DL3
    MS08LDL4 As Integer             ' 測定値08 L/DL4
    MS08LDL5 As Integer             ' 測定値08 L/DL5
    MS08DEN1 As Integer             ' 測定値08 Den1
    MS08DEN2 As Integer             ' 測定値08 Den2
    MS08DEN3 As Integer             ' 測定値08 Den3
    MS08DEN4 As Integer             ' 測定値08 Den4
    MS08DEN5 As Integer             ' 測定値08 Den5
    MS09LDL1 As Integer             ' 測定値09 L/DL1
    MS09LDL2 As Integer             ' 測定値09 L/DL2
    MS09LDL3 As Integer             ' 測定値09 L/DL3
    MS09LDL4 As Integer             ' 測定値09 L/DL4
    MS09LDL5 As Integer             ' 測定値09 L/DL5
    MS09DEN1 As Integer             ' 測定値09 Den1
    MS09DEN2 As Integer             ' 測定値09 Den2
    MS09DEN3 As Integer             ' 測定値09 Den3
    MS09DEN4 As Integer             ' 測定値09 Den4
    MS09DEN5 As Integer             ' 測定値09 Den5
    MS10LDL1 As Integer             ' 測定値10 L/DL1
    MS10LDL2 As Integer             ' 測定値10 L/DL2
    MS10LDL3 As Integer             ' 測定値10 L/DL3
    MS10LDL4 As Integer             ' 測定値10 L/DL4
    MS10LDL5 As Integer             ' 測定値10 L/DL5
    MS10DEN1 As Integer             ' 測定値10 Den1
    MS10DEN2 As Integer             ' 測定値10 Den2
    MS10DEN3 As Integer             ' 測定値10 Den3
    MS10DEN4 As Integer             ' 測定値10 Den4
    MS10DEN5 As Integer             ' 測定値10 Den5
    MS11LDL1 As Integer             ' 測定値11 L/DL1
    MS11LDL2 As Integer             ' 測定値11 L/DL2
    MS11LDL3 As Integer             ' 測定値11 L/DL3
    MS11LDL4 As Integer             ' 測定値11 L/DL4
    MS11LDL5 As Integer             ' 測定値11 L/DL5
    MS11DEN1 As Integer             ' 測定値11 Den1
    MS11DEN2 As Integer             ' 測定値11 Den2
    MS11DEN3 As Integer             ' 測定値11 Den3
    MS11DEN4 As Integer             ' 測定値11 Den4
    MS11DEN5 As Integer             ' 測定値11 Den5
    MS12LDL1 As Integer             ' 測定値12 L/DL1
    MS12LDL2 As Integer             ' 測定値12 L/DL2
    MS12LDL3 As Integer             ' 測定値12 L/DL3
    MS12LDL4 As Integer             ' 測定値12 L/DL4
    MS12LDL5 As Integer             ' 測定値12 L/DL5
    MS12DEN1 As Integer             ' 測定値12 Den1
    MS12DEN2 As Integer             ' 測定値12 Den2
    MS12DEN3 As Integer             ' 測定値12 Den3
    MS12DEN4 As Integer             ' 測定値12 Den4
    MS12DEN5 As Integer             ' 測定値12 Den5
    MS13LDL1 As Integer             ' 測定値13 L/DL1
    MS13LDL2 As Integer             ' 測定値13 L/DL2
    MS13LDL3 As Integer             ' 測定値13 L/DL3
    MS13LDL4 As Integer             ' 測定値13 L/DL4
    MS13LDL5 As Integer             ' 測定値13 L/DL5
    MS13DEN1 As Integer             ' 測定値13 Den1
    MS13DEN2 As Integer             ' 測定値13 Den2
    MS13DEN3 As Integer             ' 測定値13 Den3
    MS13DEN4 As Integer             ' 測定値13 Den4
    MS13DEN5 As Integer             ' 測定値13 Den5
    MS14LDL1 As Integer             ' 測定値14 L/DL1
    MS14LDL2 As Integer             ' 測定値14 L/DL2
    MS14LDL3 As Integer             ' 測定値14 L/DL3
    MS14LDL4 As Integer             ' 測定値14 L/DL4
    MS14LDL5 As Integer             ' 測定値14 L/DL5
    MS14DEN1 As Integer             ' 測定値14 Den1
    MS14DEN2 As Integer             ' 測定値14 Den2
    MS14DEN3 As Integer             ' 測定値14 Den3
    MS14DEN4 As Integer             ' 測定値14 Den4
    MS14DEN5 As Integer             ' 測定値14 Den5
    MS15LDL1 As Integer             ' 測定値15 L/DL1
    MS15LDL2 As Integer             ' 測定値15 L/DL2
    MS15LDL3 As Integer             ' 測定値15 L/DL3
    MS15LDL4 As Integer             ' 測定値15 L/DL4
    MS15LDL5 As Integer             ' 測定値15 L/DL5
    MS15DEN1 As Integer             ' 測定値15 Den1
    MS15DEN2 As Integer             ' 測定値15 Den2
    MS15DEN3 As Integer             ' 測定値15 Den3
    MS15DEN4 As Integer             ' 測定値15 Den4
    MS15DEN5 As Integer             ' 測定値15 Den5
    MS01DVD2 As Integer             ' 測定値01 DVD   2002/7/02 tuku
    MS02DVD2 As Integer             ' 測定値02 DVD
    MS03DVD2 As Integer             ' 測定値03 DVD
    MS04DVD2 As Integer             ' 測定値04 DVD
    MS05DVD2 As Integer             ' 測定値05 DVD
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
    
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    MSZEROMN    As Integer          ' L/DL0連続数最小値
    MSZEROMX    As Integer          ' L/DL0連続数最大値
    PTNJUDGRES  As String * 1       ' パターン判定結果
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
End Type


' ライフタイム
Public Type typ_TBCMJ007
    CRYNUM As String * 12           ' 結晶番号
    POSITION As Integer             ' 位置
    SMPKBN As String * 1            ' サンプル区分
    TRANCOND As String * 1          ' 処理条件
    TRANCNT As Integer              ' 処理回数
    SMPLNO As Long                  ' サンプルＮｏ      Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' サンプル有無
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    GOUKI As String * 3             ' 号機
    MEAS1 As Integer                ' 測定値１
    MEAS2 As Integer                ' 測定値２
    MEAS3 As Integer                ' 測定値３
    MEAS4 As Integer                ' 測定値４
    MEAS5 As Integer                ' 測定値５
    MEASPEAK As Integer             ' 測定値 ピーク値
    CALCMEAS As Integer             ' 計算結果
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
''add 2005/11/11 高崎->
''->スプレッドに空文字列を表示するため
    MEAS6 As Integer                ' 測定値６
    MEAS7 As Integer                ' 測定値７
    MEAS8 As Integer                ' 測定値８
    MEAS9 As Integer                ' 測定値９
    MEAS10 As Integer               ' 測定値１０
    MEASFILE As String              ' 測定データファイル名
    RESVAL As String                ' 実測抵抗
    INCVAL As String                ' 傾き
    CUTVAL As String                ' 切片
    SETVAL As String                ' 設定値
    CONVAL As String                ' 10Ω換算値
    MEAS1DAT1 As String             ' 測定値１　生データ１
    MEAS1DAT2 As String             ' 測定値１　生データ２
    MEAS1DAT3 As String             ' 測定値１　生データ３
    MEAS1DAT4 As String             ' 測定値１　生データ４
    MEAS1DAT5 As String             ' 測定値１　生データ５
    MEAS2DAT1 As String             ' 測定値２　生データ１
    MEAS2DAT2 As String             ' 測定値２　生データ２
    MEAS2DAT3 As String             ' 測定値２　生データ３
    MEAS2DAT4 As String             ' 測定値２　生データ４
    MEAS2DAT5 As String             ' 測定値２　生データ５
    MEAS3DAT1 As String             ' 測定値３　生データ１
    MEAS3DAT2 As String             ' 測定値３　生データ２
    MEAS3DAT3 As String             ' 測定値３　生データ３
    MEAS3DAT4 As String             ' 測定値３　生データ４
    MEAS3DAT5 As String             ' 測定値３　生データ５
    MEAS4DAT1 As String             ' 測定値４　生データ１
    MEAS4DAT2 As String             ' 測定値４　生データ２
    MEAS4DAT3 As String             ' 測定値４　生データ３
    MEAS4DAT4 As String             ' 測定値４　生データ４
    MEAS4DAT5 As String             ' 測定値４　生データ５
    MEAS5DAT1 As String             ' 測定値５　生データ１
    MEAS5DAT2 As String             ' 測定値５　生データ２
    MEAS5DAT3 As String             ' 測定値５　生データ３
    MEAS5DAT4 As String             ' 測定値５　生データ４
    MEAS5DAT5 As String             ' 測定値５　生データ５
    MEAS6DAT1 As String             ' 測定値６　生データ１
    MEAS6DAT2 As String             ' 測定値６　生データ２
    MEAS6DAT3 As String             ' 測定値６　生データ３
    MEAS6DAT4 As String             ' 測定値６　生データ４
    MEAS6DAT5 As String             ' 測定値６　生データ５
    MEAS7DAT1 As String             ' 測定値７　生データ１
    MEAS7DAT2 As String             ' 測定値７　生データ２
    MEAS7DAT3 As String             ' 測定値７　生データ３
    MEAS7DAT4 As String             ' 測定値７　生データ４
    MEAS7DAT5 As String             ' 測定値７　生データ５
    MEAS8DAT1 As String             ' 測定値８　生データ１
    MEAS8DAT2 As String             ' 測定値８　生データ２
    MEAS8DAT3 As String             ' 測定値８　生データ３
    MEAS8DAT4 As String             ' 測定値８　生データ４
    MEAS8DAT5 As String             ' 測定値８　生データ５
    MEAS9DAT1 As String             ' 測定値９　生データ１
    MEAS9DAT2 As String             ' 測定値９　生データ２
    MEAS9DAT3 As String             ' 測定値９　生データ３
    MEAS9DAT4 As String             ' 測定値９　生データ４
    MEAS9DAT5 As String             ' 測定値９　生データ５
    MEAS10DAT1 As String            ' 測定値１０　生データ１
    MEAS10DAT2 As String            ' 測定値１０　生データ２
    MEAS10DAT3 As String            ' 測定値１０　生データ３
    MEAS10DAT4 As String            ' 測定値１０　生データ４
    MEAS10DAT5 As String            ' 測定値１０　生データ５
    LTSPIFLG As String              ' 測定位置判定フラグ
''add 2005/11/11 高崎->
End Type


' ＢＭＤ実績
Public Type typ_TBCMJ008
    CRYNUM As String * 12           ' 結晶番号
    POSITION As Integer             ' 位置
    SMPKBN As String * 1            ' サンプル区分
    TRANCOND As String * 1          ' 処理条件
    TRANCNT As Integer              ' 処理回数
    SMPLNO As Long                  ' サンプルＮｏ      Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' サンプル有無
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    GOUKI As String * 3             ' 号機
    MEASMETH As String * 1          ' 測定方法
    MEASSPOT As Integer             ' 測定点
    MAG As String * 4               ' 倍率
    HTPRC As String * 2             ' 熱処理方法
    KKSP As String * 3              ' 結晶欠陥測定位置
    KKSET As String * 3             ' 結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    MEAS1 As Integer                ' 測定値１
    MEAS2 As Integer                ' 測定値２
    MEAS3 As Integer                ' 測定値３
    MEAS4 As Integer                ' 測定値４
    MEAS5 As Integer                ' 測定値５
    MEASMIN As Double               ' MIN
    MEASMAX As Double               ' MAX
    MEASAVE As Double               ' AVE
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
'OSF，BMD項目追加対応  2002.04.02 yakimura
    BMDMNBUNP As Double             ' ＢＭＤ面内分布
'OSF，BMD項目追加対応  2002.04.02 yakimura
End Type


' 総合判定実績
Public Type typ_TBCMJ009
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' インゴット内位置
    TRANCNT As Integer              ' 処理回数
    Length As Integer               ' 長さ
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    CODE As String * 1              ' 区分コード
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 結晶最終検査
Public Type typ_TBCMJ010
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' インゴット内位置
    TRANCNT As Integer              ' 処理回数
    Length As Integer               ' 長さ
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    PAYCLASS As String * 1          ' 払い出し区分
    OUTLENGTH As Integer            ' 出荷長さ
    PART1 As Integer                ' 部位１
    P1BDLEN As Integer              ' 部位１不良長さ
    P1BDCAUS As String * 3          ' 部位１不良理由
    PART2 As Integer                ' 部位２
    P2BDLEN As Integer              ' 部位２不良長さ
    P2BDCAUS As String * 3          ' 部位２不良理由
    PART3 As Integer                ' 部位３
    P3BDLEN As Integer              ' 部位３不良長さ
    P3BDCAUS As String * 3          ' 部位３不良理由
    PART4 As Integer                ' 部位４
    P4BDLEN As Integer              ' 部位４不良長さ
    P4BDCAUS As String * 3          ' 部位４不良理由
    PART5 As Integer                ' 部位５
    P5BDLEN As Integer              ' 部位５不良長さ
    P5BDCAUS As String * 3          ' 部位５不良理由
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SUMMITSENDFLAG As String * 1    ' SUMMIT送信フラグ
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' ＷＦ払出実績
Public Type typ_TBCMJ011
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' インゴット内位置
    Length As Integer               ' 長さ
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    BLOCKID As String * 12          ' ブロックID
    sBlockId As String * 12         ' 先頭ブロックID
    BLOCKORDER As Integer           ' ブロック順序
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SUMMITSENDFLAG As String * 1    ' SUMMIT送信フラグ
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' ホールド（解除）実績
Public Type typ_TBCMJ012
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' インゴット内位置
    TRANCNT As Integer              ' 処理回数
    Length As Integer               ' 長さ
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    HLDTRCLS As String * 1          ' ホールド処理区分
    HLDCAUSE As String * 3          ' ホールド理由
    HLDCMNT As String               ' ホールドコメント
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
    HOLDKT As String * 5            ' ﾎｰﾙﾄﾞ工程  2005/07
End Type


' 転用実績
Public Type typ_TBCMJ013
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' インゴット内位置
    TRANCNT As Integer              ' 処理回数
    Length As Integer               ' 長さ
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    DUNWNUM As String * 12          ' 転用先品番
    DUNWREV As Integer              ' 転用先品番 製品番号改訂番号
    DUNWFACT As String * 1          ' 転用先品番 工場
    DUNWOPCD As String * 1          ' 転用先品番 操業条件
    DUOGNUM As String * 12          ' 転用元品番
    DUOGREV As Integer              ' 転用元品番 製品番号改訂番号
    DUOGFACT As String * 1          ' 転用元品番 工場
    DUOGOPCD As String * 1          ' 転用元品番 操業条件
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
End Type


' 結晶総合判定測定値
Public Type typ_TBCMJ014
    CRYNUM As String * 12           ' 結晶番号
    POSITION As Integer             ' 位置
    SMPKBN As String * 1            ' サンプル区分
    Length As Integer               ' 長さ
    UBLOCKID As String * 12         ' UブロックID
    DBLOCKID As String * 12         ' DブロックID
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    PRODCOND As String * 10         ' 製作条件
    PGID As String * 8              ' ＰＧ－ＩＤ
    UPLENGTH As Integer             ' 引上げ長さ
    PLUPDATE As Date                ' 引上日付
    FREELENG As Integer             ' フリー長
    DIAMETER As Integer             ' 直径
    CHARGE As Long                  ' チャージ量
    SEED As String * 4              ' シード
    SXL_RS_SMPPOS As Integer        ' SXLRSｻﾝﾌﾟﾙ測定位置（SXL測定情報）
    SXLRS_MEAS1 As Double           ' SXLRS_測定値１
    SXLRS_MEAS2 As Double           ' SXLRS_測定値２
    SXLRS_MEAS3 As Double           ' SXLRS_測定値３
    SXLRS_MEAS4 As Double           ' SXLRS_測定値４
    SXLRS_MEAS5 As Double           ' SXLRS_測定値５
    SXLRS_EFEHS As Double           ' SXLRS_実効偏析
    SXLRS_RRG As Double             ' SXLRS_ＲＲＧ
    SXL_OI_SMPPOS As Integer        ' SXLOIｻﾝﾌﾟﾙ測定位置（SXL測定情報）
    SXLOI_OIMEAS1 As Double         ' SXLOI_Ｏｉ測定値１
    SXLOI_OIMEAS2 As Double         ' SXLOI_Ｏｉ測定値２
    SXLOI_OIMEAS3 As Double         ' SXLOI_Ｏｉ測定値３
    SXLOI_OIMEAS4 As Double         ' SXLOI_Ｏｉ測定値４
    SXLOI_OIMEAS5 As Double         ' SXLOI_Ｏｉ測定値５
    SXLOI_ORGRES As Double          ' SXLOI_ＯＲＧ結果
    SXLOI_INSPECTWAY As String * 2  ' SXLOI_検査方法
    SXL_CS_SMPPOS As Integer        ' SXLCSｻﾝﾌﾟﾙ測定位置（SXL測定情報）
    SXLCS_CSMEAS As Double          ' SXLCS_Cs実測値
    SXLCS_70PPRE As Double          ' SXLCS_７０％推定値
    SXLOSF1_SMPPOS As Integer       ' SXLOSFｻﾝﾌﾟﾙ測定位置（SXL位置情報）
    SXLOSF1_KKSP As String * 3      ' SXLOSF1結晶欠陥測定位置
    SXLOSF1_NETU As String * 2      ' SXLOSF1熱処理法
    SXLOSF1_KKSET As String * 3     ' SXLOSF1結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLOSF1_MEAS1 As Integer        ' SXLOSF1測定点１
    SXLOSF1_MEAS2 As Integer        ' SXLOSF1測定点2
    SXLOSF1_MEAS3 As Integer        ' SXLOSF1測定点3
    SXLOSF1_MEAS4 As Integer        ' SXLOSF1測定点4
    SXLOSF1_MEAS5 As Integer        ' SXLOSF1測定点5
    SXLOSF1_MEAS6 As Integer        ' SXLOSF1測定点6
    SXLOSF1_MEAS7 As Integer        ' SXLOSF1測定点7
    SXLOSF1_MEAS8 As Integer        ' SXLOSF1測定点8
    SXLOSF1_MEAS9 As Integer        ' SXLOSF1測定点9
    SXLOSF1_MEAS10 As Integer       ' SXLOSF1測定点10
    SXLOSF1_MEAS11 As Integer       ' SXLOSF1測定点11
    SXLOSF1_MEAS12 As Integer       ' SXLOSF1測定点12
    SXLOSF1_MEAS13 As Integer       ' SXLOSF1測定点13
    SXLOSF1_MEAS14 As Integer       ' SXLOSF1測定点14
    SXLOSF1_MEAS15 As Integer       ' SXLOSF1測定点15
    SXLOSF1_MEAS16 As Integer       ' SXLOSF1測定点16
    SXLOSF1_MEAS17 As Integer       ' SXLOSF1測定点17
    SXLOSF1_MEAS18 As Integer       ' SXLOSF1測定点18
    SXLOSF1_MEAS19 As Integer       ' SXLOSF1測定点19
    SXLOSF1_MEAS20 As Integer       ' SXLOSF1測定点20
    SXLOSF1_CALCMAX As Double       ' OSF1SXL計算結果 Max_1
    SXLOSF1_CALCAVE As Double       ' OSF1SXL計算結果 Ave_1
    SXLOSF2_KKSP As String * 3      ' SXLOSF２結晶欠陥測定位置
    SXLOSF2_NETU As String * 2      ' SXLOSF２熱処理法
    SXLOSF2_KKSET As String * 3     ' SXLOSF２結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLOSF2_MEAS1 As Integer        ' SXLOSF2測定点１
    SXLOSF2_MEAS2 As Integer        ' SXLOSF2測定点2
    SXLOSF2_MEAS3 As Integer        ' SXLOSF2測定点3
    SXLOSF2_MEAS4 As Integer        ' SXLOSF2測定点4
    SXLOSF2_MEAS5 As Integer        ' SXLOSF2測定点5
    SXLOSF2_MEAS6 As Integer        ' SXLOSF2測定点6
    SXLOSF2_MEAS7 As Integer        ' SXLOSF2測定点7
    SXLOSF2_MEAS8 As Integer        ' SXLOSF2測定点8
    SXLOSF2_MEAS9 As Integer        ' SXLOSF2測定点9
    SXLOSF2_MEAS10 As Integer       ' SXLOSF2測定点10
    SXLOSF2_MEAS11 As Integer       ' SXLOSF2測定点11
    SXLOSF2_MEAS12 As Integer       ' SXLOSF2測定点12
    SXLOSF2_MEAS13 As Integer       ' SXLOSF2測定点13
    SXLOSF2_MEAS14 As Integer       ' SXLOSF2測定点14
    SXLOSF2_MEAS15 As Integer       ' SXLOSF2測定点15
    SXLOSF2_MEAS16 As Integer       ' SXLOSF2測定点16
    SXLOSF2_MEAS17 As Integer       ' SXLOSF2測定点17
    SXLOSF2_MEAS18 As Integer       ' SXLOSF2測定点18
    SXLOSF2_MEAS19 As Integer       ' SXLOSF2測定点19
    SXLOSF2_MEAS20 As Integer       ' SXLOSF2測定点20
    SXLOSF2_CALCMAX As Double       ' OSF２SXL計算結果 Max_2
    SXLOSF2_CALCAVE As Double       ' OSF２SXL計算結果 Ave_2
    SXLOSF3_KKSP As String * 3      ' SXLOSF３結晶欠陥測定位置
    SXLOSF3_NETU As String * 2      ' SXLOSF３熱処理法
    SXLOSF3_KKSET As String * 3     ' SXLOSF３結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLOSF3_MEAS1 As Integer        ' SXLOSF3測定点１
    SXLOSF3_MEAS2 As Integer        ' SXLOSF3測定点2
    SXLOSF3_MEAS3 As Integer        ' SXLOSF3測定点3
    SXLOSF3_MEAS4 As Integer        ' SXLOSF3測定点4
    SXLOSF3_MEAS5 As Integer        ' SXLOSF3測定点5
    SXLOSF3_MEAS6 As Integer        ' SXLOSF3測定点6
    SXLOSF3_MEAS7 As Integer        ' SXLOSF3測定点7
    SXLOSF3_MEAS8 As Integer        ' SXLOSF3測定点8
    SXLOSF3_MEAS9 As Integer        ' SXLOSF3測定点9
    SXLOSF3_MEAS10 As Integer       ' SXLOSF3測定点10
    SXLOSF3_MEAS11 As Integer       ' SXLOSF3測定点11
    SXLOSF3_MEAS12 As Integer       ' SXLOSF3測定点12
    SXLOSF3_MEAS13 As Integer       ' SXLOSF3測定点13
    SXLOSF3_MEAS14 As Integer       ' SXLOSF3測定点14
    SXLOSF3_MEAS15 As Integer       ' SXLOSF3測定点15
    SXLOSF3_MEAS16 As Integer       ' SXLOSF3測定点16
    SXLOSF3_MEAS17 As Integer       ' SXLOSF3測定点17
    SXLOSF3_MEAS18 As Integer       ' SXLOSF3測定点18
    SXLOSF3_MEAS19 As Integer       ' SXLOSF3測定点19
    SXLOSF3_MEAS20 As Integer       ' SXLOSF3測定点20
    SXLOSF3_CALCMAX As Double       ' OSF３SXL計算結果 Max_3
    SXLOSF3_CALCAVE As Double       ' OSF３SXL計算結果 Ave_3
    SXLOSF4_KKSP As String * 3      ' SXLOSF４結晶欠陥測定位置
    SXLOSF4_NETU As String * 2      ' SXLOSF４熱処理法
    SXLOSF4_KKSET As String * 3     ' SXLOSF４結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLOSF4_MEAS1 As Integer        ' SXLOSF4測定点１
    SXLOSF4_MEAS2 As Integer        ' SXLOSF4測定点2
    SXLOSF4_MEAS3 As Integer        ' SXLOSF4測定点3
    SXLOSF4_MEAS4 As Integer        ' SXLOSF4測定点4
    SXLOSF4_MEAS5 As Integer        ' SXLOSF4測定点5
    SXLOSF4_MEAS6 As Integer        ' SXLOSF4測定点6
    SXLOSF4_MEAS7 As Integer        ' SXLOSF4測定点7
    SXLOSF4_MEAS8 As Integer        ' SXLOSF4測定点8
    SXLOSF4_MEAS9 As Integer        ' SXLOSF4測定点9
    SXLOSF4_MEAS10 As Integer       ' SXLOSF4測定点10
    SXLOSF4_MEAS11 As Integer       ' SXLOSF4測定点11
    SXLOSF4_MEAS12 As Integer       ' SXLOSF4測定点12
    SXLOSF4_MEAS13 As Integer       ' SXLOSF4測定点13
    SXLOSF4_MEAS14 As Integer       ' SXLOSF4測定点14
    SXLOSF4_MEAS15 As Integer       ' SXLOSF4測定点15
    SXLOSF4_MEAS16 As Integer       ' SXLOSF4測定点16
    SXLOSF4_MEAS17 As Integer       ' SXLOSF4測定点17
    SXLOSF4_MEAS18 As Integer       ' SXLOSF4測定点18
    SXLOSF4_MEAS19 As Integer       ' SXLOSF4測定点19
    SXLOSF4_MEAS20 As Integer       ' SXLOSF4測定点20
    SXLOSF4_CALCMAX As Double       ' OSF４SXL計算結果 Max_4
    SXLOSF4_CALCAVE As Double       ' OSF４SXL計算結果 Ave_4
    SXLBMD_SMPPOS As Integer        ' SXLBMDｻﾝﾌﾟﾙ測定位置（SXL位置情報）
    SXLBMD1_KKSP As String * 3      ' SXLBMD1結晶欠陥測定位置
    SXLBMD1_NETU As String * 2      ' SXLBMD1熱処理法
    SXLBMD1_KKSET As String * 3     ' SXLBMD1結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLBMD1_MEAS1 As Integer        ' SXLBMD1測定点１
    SXLBMD1_MEAS2 As Integer        ' SXLBMD1測定点2
    SXLBMD1_MEAS3 As Integer        ' SXLBMD1測定点3
    SXLBMD1_MEAS4 As Integer        ' SXLBMD1測定点4
    SXLBMD1_MEAS5 As Integer        ' SXLBMD1測定点5
    SXLBMD1_CALCMAX As Double       ' BMD1SXL計算結果 Max
    SXLBMD1_CALCAVE As Double       ' BMD1SXL計算結果 Ave
    SXLBMD2_KKSP As String * 3      ' SXLBMD２結晶欠陥測定位置
    SXLBMD2_NETU As String * 2      ' SXLBMD２熱処理法
    SXLBMD2_KKSET As String * 3     ' SXLBMD２結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLBMD2_MEAS1 As Integer        ' SXLBMD2測定点１
    SXLBMD2_MEAS2 As Integer        ' SXLBMD2測定点2
    SXLBMD2_MEAS3 As Integer        ' SXLBMD2測定点3
    SXLBMD2_MEAS4 As Integer        ' SXLBMD2測定点4
    SXLBMD2_MEAS5 As Integer        ' SXLBMD2測定点5
    SXLBMD2_CALCMAX As Double       ' BMD２SXL計算結果 Max
    SXLBMD2_CALCAVE As Double       ' BMD２SXL計算結果 Ave
    SXLBMD3_KKSP As String * 3      ' SXLBMD３結晶欠陥測定位置
    SXLBMD3_NETU As String * 2      ' SXLBMD３熱処理法
    SXLBMD3_KKSET As String * 3     ' SXLBMD３結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    SXLBMD3_MEAS1 As Integer        ' SXLBMD3測定点１
    SXLBMD3_MEAS2 As Integer        ' SXLBMD3測定点2
    SXLBMD3_MEAS3 As Integer        ' SXLBMD3測定点3
    SXLBMD3_MEAS4 As Integer        ' SXLBMD3測定点4
    SXLBMD3_MEAS5 As Integer        ' SXLBMD3測定点5
    SXLBMD3_CALCMAX As Double       ' BMD３SXL計算結果 Max
    SXLBMD3_CALCAVE As Double       ' BMD３SXL計算結果 Ave
    SXLGD_SMPPOS As Integer         ' SXLGDｻﾝﾌﾟﾙ測定位置（SXL位置情報）
    SXLGD_MS01LDL1 As Integer       ' SXLGD_測定値01 L/DL1
    SXLGD_MS01LDL2 As Integer       ' SXLGD_測定値01 L/DL2
    SXLGD_MS01LDL3 As Integer       ' SXLGD_測定値01 L/DL3
    SXLGD_MS01LDL4 As Integer       ' SXLGD_測定値01 L/DL4
    SXLGD_MS01LDL5 As Integer       ' SXLGD_測定値01 L/DL5
    SXLGD_MS01DEN1 As Integer       ' SXLGD_測定値01 Den1
    SXLGD_MS01DEN2 As Integer       ' SXLGD_測定値01 Den2
    SXLGD_MS01DEN3 As Integer       ' SXLGD_測定値01 Den3
    SXLGD_MS01DEN4 As Integer       ' SXLGD_測定値01 Den4
    SXLGD_MS01DEN5 As Integer       ' SXLGD_測定値01 Den5
    SXLGD_MS02LDL1 As Integer       ' SXLGD_測定値02 L/DL1
    SXLGD_MS02LDL2 As Integer       ' SXLGD_測定値02 L/DL2
    SXLGD_MS02LDL3 As Integer       ' SXLGD_測定値02 L/DL3
    SXLGD_MS02LDL4 As Integer       ' SXLGD_測定値02 L/DL4
    SXLGD_MS02LDL5 As Integer       ' SXLGD_測定値02 L/DL5
    SXLGD_MS02DEN1 As Integer       ' SXLGD_測定値02 Den1
    SXLGD_MS02DEN2 As Integer       ' SXLGD_測定値02 Den2
    SXLGD_MS02DEN3 As Integer       ' SXLGD_測定値02 Den3
    SXLGD_MS02DEN4 As Integer       ' SXLGD_測定値02 Den4
    SXLGD_MS02DEN5 As Integer       ' SXLGD_測定値02 Den5
    SXLGD_MS03LDL1 As Integer       ' SXLGD_測定値03 L/DL1
    SXLGD_MS03LDL2 As Integer       ' SXLGD_測定値03 L/DL2
    SXLGD_MS03LDL3 As Integer       ' SXLGD_測定値03 L/DL3
    SXLGD_MS03LDL4 As Integer       ' SXLGD_測定値03 L/DL4
    SXLGD_MS03LDL5 As Integer       ' SXLGD_測定値03 L/DL5
    SXLGD_MS03DEN1 As Integer       ' SXLGD_測定値03 Den1
    SXLGD_MS03DEN2 As Integer       ' SXLGD_測定値03 Den2
    SXLGD_MS03DEN3 As Integer       ' SXLGD_測定値03 Den3
    SXLGD_MS03DEN4 As Integer       ' SXLGD_測定値03 Den4
    SXLGD_MS03DEN5 As Integer       ' SXLGD_測定値03 Den5
    SXLGD_MS04LDL1 As Integer       ' SXLGD_測定値04 L/DL1
    SXLGD_MS04LDL2 As Integer       ' SXLGD_測定値04 L/DL2
    SXLGD_MS04LDL3 As Integer       ' SXLGD_測定値04 L/DL3
    SXLGD_MS04LDL4 As Integer       ' SXLGD_測定値04 L/DL4
    SXLGD_MS04LDL5 As Integer       ' SXLGD_測定値04 L/DL5
    SXLGD_MS04DEN1 As Integer       ' SXLGD_測定値04 Den1
    SXLGD_MS04DEN2 As Integer       ' SXLGD_測定値04 Den2
    SXLGD_MS04DEN3 As Integer       ' SXLGD_測定値04 Den3
    SXLGD_MS04DEN4 As Integer       ' SXLGD_測定値04 Den4
    SXLGD_MS04DEN5 As Integer       ' SXLGD_測定値04 Den5
    SXLGD_MS05LDL1 As Integer       ' SXLGD_測定値05 L/DL1
    SXLGD_MS05LDL2 As Integer       ' SXLGD_測定値05 L/DL2
    SXLGD_MS05LDL3 As Integer       ' SXLGD_測定値05 L/DL3
    SXLGD_MS05LDL4 As Integer       ' SXLGD_測定値05 L/DL4
    SXLGD_MS05LDL5 As Integer       ' SXLGD_測定値05 L/DL5
    SXLGD_MS05DEN1 As Integer       ' SXLGD_測定値05 Den1
    SXLGD_MS05DEN2 As Integer       ' SXLGD_測定値05 Den2
    SXLGD_MS05DEN3 As Integer       ' SXLGD_測定値05 Den3
    SXLGD_MS05DEN4 As Integer       ' SXLGD_測定値05 Den4
    SXLGD_MS05DEN5 As Integer       ' SXLGD_測定値05 Den5
    SXLGD_MS06LDL1 As Integer       ' SXLGD_測定値06 L/DL1
    SXLGD_MS06LDL2 As Integer       ' SXLGD_測定値06 L/DL2
    SXLGD_MS06LDL3 As Integer       ' SXLGD_測定値06 L/DL3
    SXLGD_MS06LDL4 As Integer       ' SXLGD_測定値06 L/DL4
    SXLGD_MS06LDL5 As Integer       ' SXLGD_測定値06 L/DL5
    SXLGD_MS06DEN1 As Integer       ' SXLGD_測定値06 Den1
    SXLGD_MS06DEN2 As Integer       ' SXLGD_測定値06 Den2
    SXLGD_MS06DEN3 As Integer       ' SXLGD_測定値06 Den3
    SXLGD_MS06DEN4 As Integer       ' SXLGD_測定値06 Den4
    SXLGD_MS06DEN5 As Integer       ' SXLGD_測定値06 Den5
    SXLGD_MS07LDL1 As Integer       ' SXLGD_測定値07 L/DL1
    SXLGD_MS07LDL2 As Integer       ' SXLGD_測定値07 L/DL2
    SXLGD_MS07LDL3 As Integer       ' SXLGD_測定値07 L/DL3
    SXLGD_MS07LDL4 As Integer       ' SXLGD_測定値07 L/DL4
    SXLGD_MS07LDL5 As Integer       ' SXLGD_測定値07 L/DL5
    SXLGD_MS07DEN1 As Integer       ' SXLGD_測定値07 Den1
    SXLGD_MS07DEN2 As Integer       ' SXLGD_測定値07 Den2
    SXLGD_MS07DEN3 As Integer       ' SXLGD_測定値07 Den3
    SXLGD_MS07DEN4 As Integer       ' SXLGD_測定値07 Den4
    SXLGD_MS07DEN5 As Integer       ' SXLGD_測定値07 Den5
    SXLGD_MS08LDL1 As Integer       ' SXLGD_測定値08 L/DL1
    SXLGD_MS08LDL2 As Integer       ' SXLGD_測定値08 L/DL2
    SXLGD_MS08LDL3 As Integer       ' SXLGD_測定値08 L/DL3
    SXLGD_MS08LDL4 As Integer       ' SXLGD_測定値08 L/DL4
    SXLGD_MS08LDL5 As Integer       ' SXLGD_測定値08 L/DL5
    SXLGD_MS08DEN1 As Integer       ' SXLGD_測定値08 Den1
    SXLGD_MS08DEN2 As Integer       ' SXLGD_測定値08 Den2
    SXLGD_MS08DEN3 As Integer       ' SXLGD_測定値08 Den3
    SXLGD_MS08DEN4 As Integer       ' SXLGD_測定値08 Den4
    SXLGD_MS08DEN5 As Integer       ' SXLGD_測定値08 Den5
    SXLGD_MS09LDL1 As Integer       ' SXLGD_測定値09 L/DL1
    SXLGD_MS09LDL2 As Integer       ' SXLGD_測定値09 L/DL2
    SXLGD_MS09LDL3 As Integer       ' SXLGD_測定値09 L/DL3
    SXLGD_MS09LDL4 As Integer       ' SXLGD_測定値09 L/DL4
    SXLGD_MS09LDL5 As Integer       ' SXLGD_測定値09 L/DL5
    SXLGD_MS09DEN1 As Integer       ' SXLGD_測定値09 Den1
    SXLGD_MS09DEN2 As Integer       ' SXLGD_測定値09 Den2
    SXLGD_MS09DEN3 As Integer       ' SXLGD_測定値09 Den3
    SXLGD_MS09DEN4 As Integer       ' SXLGD_測定値09 Den4
    SXLGD_MS09DEN5 As Integer       ' SXLGD_測定値09 Den5
    SXLGD_MS10LDL1 As Integer       ' SXLGD_測定値10 L/DL1
    SXLGD_MS10LDL2 As Integer       ' SXLGD_測定値10 L/DL2
    SXLGD_MS10LDL3 As Integer       ' SXLGD_測定値10 L/DL3
    SXLGD_MS10LDL4 As Integer       ' SXLGD_測定値10 L/DL4
    SXLGD_MS10LDL5 As Integer       ' SXLGD_測定値10 L/DL5
    SXLGD_MS10DEN1 As Integer       ' SXLGD_測定値10 Den1
    SXLGD_MS10DEN2 As Integer       ' SXLGD_測定値10 Den2
    SXLGD_MS10DEN3 As Integer       ' SXLGD_測定値10 Den3
    SXLGD_MS10DEN4 As Integer       ' SXLGD_測定値10 Den4
    SXLGD_MS10DEN5 As Integer       ' SXLGD_測定値10 Den5
    SXLGD_MS11LDL1 As Integer       ' SXLGD_測定値11 L/DL1
    SXLGD_MS11LDL2 As Integer       ' SXLGD_測定値11 L/DL2
    SXLGD_MS11LDL3 As Integer       ' SXLGD_測定値11 L/DL3
    SXLGD_MS11LDL4 As Integer       ' SXLGD_測定値11 L/DL4
    SXLGD_MS11LDL5 As Integer       ' SXLGD_測定値11 L/DL5
    SXLGD_MS11DEN1 As Integer       ' SXLGD_測定値11 Den1
    SXLGD_MS11DEN2 As Integer       ' SXLGD_測定値11 Den2
    SXLGD_MS11DEN3 As Integer       ' SXLGD_測定値11 Den3
    SXLGD_MS11DEN4 As Integer       ' SXLGD_測定値11 Den4
    SXLGD_MS11DEN5 As Integer       ' SXLGD_測定値11 Den5
    SXLGD_MS12LDL1 As Integer       ' SXLGD_測定値12 L/DL1
    SXLGD_MS12LDL2 As Integer       ' SXLGD_測定値12 L/DL2
    SXLGD_MS12LDL3 As Integer       ' SXLGD_測定値12 L/DL3
    SXLGD_MS12LDL4 As Integer       ' SXLGD_測定値12 L/DL4
    SXLGD_MS12LDL5 As Integer       ' SXLGD_測定値12 L/DL5
    SXLGD_MS12DEN1 As Integer       ' SXLGD_測定値12 Den1
    SXLGD_MS12DEN2 As Integer       ' SXLGD_測定値12 Den2
    SXLGD_MS12DEN3 As Integer       ' SXLGD_測定値12 Den3
    SXLGD_MS12DEN4 As Integer       ' SXLGD_測定値12 Den4
    SXLGD_MS12DEN5 As Integer       ' SXLGD_測定値12 Den5
    SXLGD_MS13LDL1 As Integer       ' SXLGD_測定値13 L/DL1
    SXLGD_MS13LDL2 As Integer       ' SXLGD_測定値13 L/DL2
    SXLGD_MS13LDL3 As Integer       ' SXLGD_測定値13 L/DL3
    SXLGD_MS13LDL4 As Integer       ' SXLGD_測定値13 L/DL4
    SXLGD_MS13LDL5 As Integer       ' SXLGD_測定値13 L/DL5
    SXLGD_MS13DEN1 As Integer       ' SXLGD_測定値13 Den1
    SXLGD_MS13DEN2 As Integer       ' SXLGD_測定値13 Den2
    SXLGD_MS13DEN3 As Integer       ' SXLGD_測定値13 Den3
    SXLGD_MS13DEN4 As Integer       ' SXLGD_測定値13 Den4
    SXLGD_MS13DEN5 As Integer       ' SXLGD_測定値13 Den5
    SXLGD_MS14LDL1 As Integer       ' SXLGD_測定値14 L/DL1
    SXLGD_MS14LDL2 As Integer       ' SXLGD_測定値14 L/DL2
    SXLGD_MS14LDL3 As Integer       ' SXLGD_測定値14 L/DL3
    SXLGD_MS14LDL4 As Integer       ' SXLGD_測定値14 L/DL4
    SXLGD_MS14LDL5 As Integer       ' SXLGD_測定値14 L/DL5
    SXLGD_MS14DEN1 As Integer       ' SXLGD_測定値14 Den1
    SXLGD_MS14DEN2 As Integer       ' SXLGD_測定値14 Den2
    SXLGD_MS14DEN3 As Integer       ' SXLGD_測定値14 Den3
    SXLGD_MS14DEN4 As Integer       ' SXLGD_測定値14 Den4
    SXLGD_MS14DEN5 As Integer       ' SXLGD_測定値14 Den5
    SXLGD_MS15LDL1 As Integer       ' SXLGD_測定値15 L/DL1
    SXLGD_MS15LDL2 As Integer       ' SXLGD_測定値15 L/DL2
    SXLGD_MS15LDL3 As Integer       ' SXLGD_測定値15 L/DL3
    SXLGD_MS15LDL4 As Integer       ' SXLGD_測定値15 L/DL4
    SXLGD_MS15LDL5 As Integer       ' SXLGD_測定値15 L/DL5
    SXLGD_MS15DEN1 As Integer       ' SXLGD_測定値15 Den1
    SXLGD_MS15DEN2 As Integer       ' SXLGD_測定値15 Den2
    SXLGD_MS15DEN3 As Integer       ' SXLGD_測定値15 Den3
    SXLGD_MS15DEN4 As Integer       ' SXLGD_測定値15 Den4
    SXLGD_MS15DEN5 As Integer       ' SXLGD_測定値15 Den5
    SXLGD_MSRSDEN As Integer        ' SXLGD_測定結果 Den
    SXLGD_MSRSLDL As Integer        ' SXLGD_測定結果 L/DL
    SXLGD_MSRSDVD2 As Integer       ' SXLGD_測定結果 DVD2
    SXLT_SMPPOS As Integer          ' SXLLTｻﾝﾌﾟﾙ測定位置（SXL位置情報）
    SXLLT_MEASPEAK As Integer       ' SXLLT_測定値 ピーク値
    SXLLT_MEAS1 As Integer          ' SXLLT_測定値1
    SXLLT_MEAS2 As Integer          ' SXLLT_測定値2
    SXLLT_MEAS3 As Integer          ' SXLLT_測定値3
    SXLLT_MEAS4 As Integer          ' SXLLT_測定値4
    SXLLT_MEAS5 As Integer          ' SXLLT_測定値5
    SXLLT_CALCMEAS As Integer       ' SXLLT_計算結果
    REGDATE As Date                 ' 登録日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
    SXLOSF1_POS1  As Double         'OSF1ﾊﾟﾀｰﾝ区分１位置
    SXLOSF1_WID1  As Double         'OSF1ﾊﾟﾀｰﾝ区分１幅
    SXLOSF1_RD1   As String * 1     'OSF1ﾊﾟﾀｰﾝ区分１R/D
    SXLOSF1_POS2  As Double         'OSF1ﾊﾟﾀｰﾝ区分２位置
    SXLOSF1_WID2  As Double         'OSF1ﾊﾟﾀｰﾝ区分２幅
    SXLOSF1_RD2   As String * 1     'OSF1ﾊﾟﾀｰﾝ区分２R/D
    SXLOSF1_POS3  As Double         'OSF1ﾊﾟﾀｰﾝ区分３位置
    SXLOSF1_WID3  As Double         'OSF1ﾊﾟﾀｰﾝ区分３幅
    SXLOSF1_RD3   As String * 1     'OSF1ﾊﾟﾀｰﾝ区分３R/D
    SXLOSF2_POS1  As Double         'OSF2ﾊﾟﾀｰﾝ区分１位置
    SXLOSF2_WID1  As Double         'OSF2ﾊﾟﾀｰﾝ区分１幅
    SXLOSF2_RD1   As String * 1     'OSF2ﾊﾟﾀｰﾝ区分１R/D
    SXLOSF2_POS2  As Double         'OSF2ﾊﾟﾀｰﾝ区分２位置
    SXLOSF2_WID2  As Double         'OSF2ﾊﾟﾀｰﾝ区分２幅
    SXLOSF2_RD2   As String * 1     'OSF2ﾊﾟﾀｰﾝ区分２R/D
    SXLOSF2_POS3  As Double         'OSF2ﾊﾟﾀｰﾝ区分３位置
    SXLOSF2_WID3  As Double         'OSF2ﾊﾟﾀｰﾝ区分３幅
    SXLOSF2_RD3   As String * 1     'OSF2ﾊﾟﾀｰﾝ区分３R/D
    SXLOSF3_POS1  As Double         'OSF3ﾊﾟﾀｰﾝ区分１位置
    SXLOSF3_WID1  As Double         'OSF3ﾊﾟﾀｰﾝ区分１幅
    SXLOSF3_RD1   As String * 1     'OSF3ﾊﾟﾀｰﾝ区分１R/D
    SXLOSF3_POS2  As Double         'OSF3ﾊﾟﾀｰﾝ区分２位置
    SXLOSF3_WID2  As Double         'OSF3ﾊﾟﾀｰﾝ区分２幅
    SXLOSF3_RD2   As String * 1     'OSF3ﾊﾟﾀｰﾝ区分２R/D
    SXLOSF3_POS3  As Double         'OSF3ﾊﾟﾀｰﾝ区分３位置
    SXLOSF3_WID3  As Double         'OSF3ﾊﾟﾀｰﾝ区分３幅
    SXLOSF3_RD3   As String * 1     'OSF3ﾊﾟﾀｰﾝ区分３R/D
    SXLOSF4_POS1  As Double         'OSF4ﾊﾟﾀｰﾝ区分１位置
    SXLOSF4_WID1  As Double         'OSF4ﾊﾟﾀｰﾝ区分１幅
    SXLOSF4_RD1   As String * 1     'OSF4ﾊﾟﾀｰﾝ区分１R/D
    SXLOSF4_POS2  As Double         'OSF4ﾊﾟﾀｰﾝ区分２位置
    SXLOSF4_WID2  As Double         'OSF4ﾊﾟﾀｰﾝ区分２幅
    SXLOSF4_RD2   As String * 1     'OSF4ﾊﾟﾀｰﾝ区分２R/D
    SXLOSF4_POS3  As Double         'OSF4ﾊﾟﾀｰﾝ区分３位置
    SXLOSF4_WID3  As Double         'OSF4ﾊﾟﾀｰﾝ区分３幅
    SXLOSF4_RD3   As String * 1     'OSF4ﾊﾟﾀｰﾝ区分３R/D
    SXLGD_MS01DVD2 As Integer       'DVD2測定結果値１
    SXLGD_MS02DVD2 As Integer       'DVD2測定結果値２
    SXLGD_MS03DVD2 As Integer       'DVD2測定結果値３
    SXLGD_MS04DVD2 As Integer       'DVD2測定結果値４
    SXLGD_MS05DVD2 As Integer       'DVD2測定結果値５
    SXLBMD1_MNBCR As Double         'BMD1SXL計算結果面内分布
    SXLBMD2_MNBCR As Double         'BMD2SXL計算結果面内分布
    SXLBMD3_MNBCR As Double         'BMD3SXL計算結果面内分布
End Type


' GD実績(WF)　05/01/31 ooba
Public Type typ_TBCMJ015
    CRYNUM      As String * 12          ' 結晶番号
    POSITION    As Integer              ' 位置
    SMPKBN      As String * 1           ' サンプル区分
    TRANCOND    As String * 1           ' 処理条件
    TRANCNT     As String * 1           ' 処理回数
    HSFLG       As String * 1           ' 保証フラグ
    SMPLNO      As String * 16          ' サンプルＮｏ
    SMPLUMU     As String * 1           ' サンプル有無
    hinban      As String * 8           ' 品番
    REVNUM      As Integer              ' 製品番号改訂番号
    FACTORY     As String * 1           ' 工場
    OPECOND     As String * 1           ' 操業条件
    SXLID       As String * 13          ' SXLID
    KRPROCCD    As String * 5           ' 管理工程コード
    PROCCODE    As String * 5           ' 工程コード
    GOUKI       As String * 3           ' 号機
    OSITEM      As String * 4           ' 評価項目
    MAISU       As Integer              ' 評価枚数
    Spec        As String * 10          ' 規格値
    NETSU       As String * 2           ' 熱処理条件
    ET          As String * 3           ' エッチング条件
    MES         As String * 3           ' 計測方法
    DKAN        As String * 10          ' ＤＫアニール条件
    ETMAE_RYO01 As Double               ' ET前重量01
    ETATO_RYO01 As Double               ' ET後重量01
    MSRSDEN     As Integer              ' 測定結果 Den
    MSRSLDL     As Integer              ' 測定結果 L/DL
    MSRSDVD2    As Integer              ' 測定結果 DVD2
    MS01LDL1    As Integer              ' 測定値01 L/DL1
    MS01LDL2    As Integer              ' 測定値01 L/DL2
    MS01LDL3    As Integer              ' 測定値01 L/DL3
    MS01LDL4    As Integer              ' 測定値01 L/DL4
    MS01LDL5    As Integer              ' 測定値01 L/DL5
    MS01DEN1    As Integer              ' 測定値01 Den1
    MS01DEN2    As Integer              ' 測定値01 Den2
    MS01DEN3    As Integer              ' 測定値01 Den3
    MS01DEN4    As Integer              ' 測定値01 Den4
    MS01DEN5    As Integer              ' 測定値01 Den5
    MS02LDL1    As Integer              ' 測定値02 L/DL1
    MS02LDL2    As Integer              ' 測定値02 L/DL2
    MS02LDL3    As Integer              ' 測定値02 L/DL3
    MS02LDL4    As Integer              ' 測定値02 L/DL4
    MS02LDL5    As Integer              ' 測定値02 L/DL5
    MS02DEN1    As Integer              ' 測定値02 Den1
    MS02DEN2    As Integer              ' 測定値02 Den2
    MS02DEN3    As Integer              ' 測定値02 Den3
    MS02DEN4    As Integer              ' 測定値02 Den4
    MS02DEN5    As Integer              ' 測定値02 Den5
    MS03LDL1    As Integer              ' 測定値03 L/DL1
    MS03LDL2    As Integer              ' 測定値03 L/DL2
    MS03LDL3    As Integer              ' 測定値03 L/DL3
    MS03LDL4    As Integer              ' 測定値03 L/DL4
    MS03LDL5    As Integer              ' 測定値03 L/DL5
    MS03DEN1    As Integer              ' 測定値03 Den1
    MS03DEN2    As Integer              ' 測定値03 Den2
    MS03DEN3    As Integer              ' 測定値03 Den3
    MS03DEN4    As Integer              ' 測定値03 Den4
    MS03DEN5    As Integer              ' 測定値03 Den5
    MS04LDL1    As Integer              ' 測定値04 L/DL1
    MS04LDL2    As Integer              ' 測定値04 L/DL2
    MS04LDL3    As Integer              ' 測定値04 L/DL3
    MS04LDL4    As Integer              ' 測定値04 L/DL4
    MS04LDL5    As Integer              ' 測定値04 L/DL5
    MS04DEN1    As Integer              ' 測定値04 Den1
    MS04DEN2    As Integer              ' 測定値04 Den2
    MS04DEN3    As Integer              ' 測定値04 Den3
    MS04DEN4    As Integer              ' 測定値04 Den4
    MS04DEN5    As Integer              ' 測定値04 Den5
    MS05LDL1    As Integer              ' 測定値05 L/DL1
    MS05LDL2    As Integer              ' 測定値05 L/DL2
    MS05LDL3    As Integer              ' 測定値05 L/DL3
    MS05LDL4    As Integer              ' 測定値05 L/DL4
    MS05LDL5    As Integer              ' 測定値05 L/DL5
    MS05DEN1    As Integer              ' 測定値05 Den1
    MS05DEN2    As Integer              ' 測定値05 Den2
    MS05DEN3    As Integer              ' 測定値05 Den3
    MS05DEN4    As Integer              ' 測定値05 Den4
    MS05DEN5    As Integer              ' 測定値05 Den5
    MS06LDL1    As Integer              ' 測定値06 L/DL1
    MS06LDL2    As Integer              ' 測定値06 L/DL2
    MS06LDL3    As Integer              ' 測定値06 L/DL3
    MS06LDL4    As Integer              ' 測定値06 L/DL4
    MS06LDL5    As Integer              ' 測定値06 L/DL5
    MS06DEN1    As Integer              ' 測定値06 Den1
    MS06DEN2    As Integer              ' 測定値06 Den2
    MS06DEN3    As Integer              ' 測定値06 Den3
    MS06DEN4    As Integer              ' 測定値06 Den4
    MS06DEN5    As Integer              ' 測定値06 Den5
    MS07LDL1    As Integer              ' 測定値07 L/DL1
    MS07LDL2    As Integer              ' 測定値07 L/DL2
    MS07LDL3    As Integer              ' 測定値07 L/DL3
    MS07LDL4    As Integer              ' 測定値07 L/DL4
    MS07LDL5    As Integer              ' 測定値07 L/DL5
    MS07DEN1    As Integer              ' 測定値07 Den1
    MS07DEN2    As Integer              ' 測定値07 Den2
    MS07DEN3    As Integer              ' 測定値07 Den3
    MS07DEN4    As Integer              ' 測定値07 Den4
    MS07DEN5    As Integer              ' 測定値07 Den5
    MS08LDL1    As Integer              ' 測定値08 L/DL1
    MS08LDL2    As Integer              ' 測定値08 L/DL2
    MS08LDL3    As Integer              ' 測定値08 L/DL3
    MS08LDL4    As Integer              ' 測定値08 L/DL4
    MS08LDL5    As Integer              ' 測定値08 L/DL5
    MS08DEN1    As Integer              ' 測定値08 Den1
    MS08DEN2    As Integer              ' 測定値08 Den2
    MS08DEN3    As Integer              ' 測定値08 Den3
    MS08DEN4    As Integer              ' 測定値08 Den4
    MS08DEN5    As Integer              ' 測定値08 Den5
    MS09LDL1    As Integer              ' 測定値09 L/DL1
    MS09LDL2    As Integer              ' 測定値09 L/DL2
    MS09LDL3    As Integer              ' 測定値09 L/DL3
    MS09LDL4    As Integer              ' 測定値09 L/DL4
    MS09LDL5    As Integer              ' 測定値09 L/DL5
    MS09DEN1    As Integer              ' 測定値09 Den1
    MS09DEN2    As Integer              ' 測定値09 Den2
    MS09DEN3    As Integer              ' 測定値09 Den3
    MS09DEN4    As Integer              ' 測定値09 Den4
    MS09DEN5    As Integer              ' 測定値09 Den5
    MS10LDL1    As Integer              ' 測定値10 L/DL1
    MS10LDL2    As Integer              ' 測定値10 L/DL2
    MS10LDL3    As Integer              ' 測定値10 L/DL3
    MS10LDL4    As Integer              ' 測定値10 L/DL4
    MS10LDL5    As Integer              ' 測定値10 L/DL5
    MS10DEN1    As Integer              ' 測定値10 Den1
    MS10DEN2    As Integer              ' 測定値10 Den2
    MS10DEN3    As Integer              ' 測定値10 Den3
    MS10DEN4    As Integer              ' 測定値10 Den4
    MS10DEN5    As Integer              ' 測定値10 Den5
    MS11LDL1    As Integer              ' 測定値11 L/DL1
    MS11LDL2    As Integer              ' 測定値11 L/DL2
    MS11LDL3    As Integer              ' 測定値11 L/DL3
    MS11LDL4    As Integer              ' 測定値11 L/DL4
    MS11LDL5    As Integer              ' 測定値11 L/DL5
    MS11DEN1    As Integer              ' 測定値11 Den1
    MS11DEN2    As Integer              ' 測定値11 Den2
    MS11DEN3    As Integer              ' 測定値11 Den3
    MS11DEN4    As Integer              ' 測定値11 Den4
    MS11DEN5    As Integer              ' 測定値11 Den5
    MS12LDL1    As Integer              ' 測定値12 L/DL1
    MS12LDL2    As Integer              ' 測定値12 L/DL2
    MS12LDL3    As Integer              ' 測定値12 L/DL3
    MS12LDL4    As Integer              ' 測定値12 L/DL4
    MS12LDL5    As Integer              ' 測定値12 L/DL5
    MS12DEN1    As Integer              ' 測定値12 Den1
    MS12DEN2    As Integer              ' 測定値12 Den2
    MS12DEN3    As Integer              ' 測定値12 Den3
    MS12DEN4    As Integer              ' 測定値12 Den4
    MS12DEN5    As Integer              ' 測定値12 Den5
    MS13LDL1    As Integer              ' 測定値13 L/DL1
    MS13LDL2    As Integer              ' 測定値13 L/DL2
    MS13LDL3    As Integer              ' 測定値13 L/DL3
    MS13LDL4    As Integer              ' 測定値13 L/DL4
    MS13LDL5    As Integer              ' 測定値13 L/DL5
    MS13DEN1    As Integer              ' 測定値13 Den1
    MS13DEN2    As Integer              ' 測定値13 Den2
    MS13DEN3    As Integer              ' 測定値13 Den3
    MS13DEN4    As Integer              ' 測定値13 Den4
    MS13DEN5    As Integer              ' 測定値13 Den5
    MS14LDL1    As Integer              ' 測定値14 L/DL1
    MS14LDL2    As Integer              ' 測定値14 L/DL2
    MS14LDL3    As Integer              ' 測定値14 L/DL3
    MS14LDL4    As Integer              ' 測定値14 L/DL4
    MS14LDL5    As Integer              ' 測定値14 L/DL5
    MS14DEN1    As Integer              ' 測定値14 Den1
    MS14DEN2    As Integer              ' 測定値14 Den2
    MS14DEN3    As Integer              ' 測定値14 Den3
    MS14DEN4    As Integer              ' 測定値14 Den4
    MS14DEN5    As Integer              ' 測定値14 Den5
    MS15LDL1    As Integer              ' 測定値15 L/DL1
    MS15LDL2    As Integer              ' 測定値15 L/DL2
    MS15LDL3    As Integer              ' 測定値15 L/DL3
    MS15LDL4    As Integer              ' 測定値15 L/DL4
    MS15LDL5    As Integer              ' 測定値15 L/DL5
    MS15DEN1    As Integer              ' 測定値15 Den1
    MS15DEN2    As Integer              ' 測定値15 Den2
    MS15DEN3    As Integer              ' 測定値15 Den3
    MS15DEN4    As Integer              ' 測定値15 Den4
    MS15DEN5    As Integer              ' 測定値15 Den5
    MS01DVD2    As Integer              ' 測定値01 DVD2
    MS02DVD2    As Integer              ' 測定値02 DVD2
    MS03DVD2    As Integer              ' 測定値03 DVD2
    MS04DVD2    As Integer              ' 測定値04 DVD2
    MS05DVD2    As Integer              ' 測定値05 DVD2
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    MSZEROMN    As Integer              ' L/DL0連続数最小値
    MSZEROMX    As Integer              ' L/DL0連続数最大値
    PTNJUDGRES  As String * 1           ' パターン判定結果
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    TSTAFFID    As String * 8           ' 登録社員ID
    REGDATE     As Date                 ' 登録日付
    KSTAFFID    As String * 8           ' 更新社員ID
    UPDDATE     As Date                 ' 更新日付
    SENDFLAG    As String * 1           ' 送信フラグ
    SENDDATE    As Date                 ' 送信日付
End Type

''Upd start 2005/06/21 (TCS)T.Terauchi  SPV9点対応  SPV実績ﾃｰﾌﾞﾙ
Public Type typ_TBCMJ016
    CRYNUM          As String * 12          ' 結晶番号
    POSITION        As Integer              ' 位置
    SMPKBN          As String * 1           ' サンプル区分
    TRANCOND        As String * 1           ' 処理条件
    TRANCNT         As Integer              ' 処理回数
    HSFLG           As String * 1           ' 保証フラグ
    SMPLNO          As String * 16          ' サンプルＮｏ
    SMPLUMU         As String * 1           ' サンプル有無
    hinban          As String * 8           ' 品番
    REVNUM          As Integer              ' 製品番号改訂番号
    FACTORY         As String * 1           ' 工場
    OPECOND         As String * 1           ' 操業条件
    SXLID           As String * 13          ' SXLID
    KRPROCCD        As String * 5           ' 管理工程コード
    PROCCODE        As String * 5           ' 工程コード
    GOUKI           As String * 3           ' 号機
    OSITEM          As String * 4           ' 評価項目
    MAISU           As Integer              ' 評価枚数
    Spec            As String * 10          ' 規格値
    NETSU           As String * 2           ' 熱処理条件
    ET              As String * 3           ' エッチング条件
    MES             As String * 3           ' 計測方法
    DKAN            As String * 10          ' ＤＫアニール条件
    SPV_Fe_MAX      As Double               ' SPV_Fe_MAX
    SPV_Fe_AVE      As Double               ' SPV_Fe_AVE
    SPV_Fe_MIN      As Double               ' SPV_Fe_MIN
    ms01_SPV_Fe     As Double               ' 測定値01 SPV_Fe
    ms02_SPV_Fe     As Double               ' 測定値02 SPV_Fe
    ms03_SPV_Fe     As Double               ' 測定値03 SPV_Fe
    ms04_SPV_Fe     As Double               ' 測定値04 SPV_Fe
    ms05_SPV_Fe     As Double               ' 測定値05 SPV_Fe
    ms06_SPV_Fe     As Double               ' 測定値06 SPV_Fe
    ms07_SPV_Fe     As Double               ' 測定値07 SPV_Fe
    ms08_SPV_Fe     As Double               ' 測定値08 SPV_Fe
    ms09_SPV_Fe     As Double               ' 測定値09 SPV_Fe
    SPV_Diff_MAX    As Double               ' SPV_拡散長_MAX
    SPV_Diff_AVE    As Double               ' SPV_拡散長_AVE
    SPV_Diff_MIN    As Double               ' SPV_拡散長_MIN
    ms01_SPV_Diff   As Double               ' 測定値01 SPV_拡散長
    ms02_SPV_Diff   As Double               ' 測定値02 SPV_拡散長
    ms03_SPV_Diff   As Double               ' 測定値03 SPV_拡散長
    ms04_SPV_Diff   As Double               ' 測定値04 SPV_拡散長
    ms05_SPV_Diff   As Double               ' 測定値05 SPV_拡散長
    ms06_SPV_Diff   As Double               ' 測定値06 SPV_拡散長
    ms07_SPV_Diff   As Double               ' 測定値07 SPV_拡散長
    ms08_SPV_Diff   As Double               ' 測定値08 SPV_拡散長
    ms09_SPV_Diff   As Double               ' 測定値09 SPV_拡散長
    TSTAFFID        As String * 8           ' 登録社員ID
    REGDATE         As Date                 ' 登録日付
    KSTAFFID        As String * 8           ' 更新社員ID
    UPDDATE         As Date                 ' 更新日付
    SENDFLAG        As String * 1           ' 送信フラグ
    SENDDATE        As Date                 ' 送信日付
    MAX_FE          As Double               ' FE濃度　最大値(表示、判定用)
    MIN_FE          As Double               ' FE濃度　最小値(表示、判定用)
    AVE_FE          As Double               ' FE濃度　平均(表示、判定用)
    CENTER_FE       As Double               ' FE濃度　中心(表示、判定用)
    MAX_DIFF        As Double               ' 拡散長　最大値(表示、判定用)
    MIN_DIFF        As Double               ' 拡散長　最小値(表示、判定用)
    AVE_DIFF        As Double               ' 拡散長　平均(表示、判定用)
    CENTER_DIFF     As Double               ' 拡散長　中心(表示、判定用)
    ''==SPV判定　20060529 SMP桜井
    SPV_Fe_PUA      As Double               'SPV_Fe PUA値
    SPV_Fe_PUAP     As Double               'SPV_Fe PUA％値
    SPV_Fe_STD      As Double               'SPV_Fe STD
    SPV_Diff_PUA    As Double               'SPV_拡散長 PUA値
    SPV_Diff_PUAP   As Double               'SPV_拡散長 PUA％値
    SPV_Nr_MAX      As Double               'SPV_OtherRecords_MAX
    SPV_Nr_AVE      As Double               'SPV_OtherRecords_AVE
    SPV_Nr_STD      As Double               'SPV_OtherRecords_STD
    SPV_Nr_PUA      As Double               'SPV_OtherRecords_PUA値
    SPV_Nr_PUAP     As Double               'SPV_OtherRecords_PUA％値
    ''==============================
    ''==SPV判定　20060612 SMP)kondoh
    PUA_FE      As Double                   ' FE濃度  PUA値(表示、判定用)
    PUAP_FE     As Double                   ' FE濃度  PUA％値(表示、判定用)
    STD_FE      As Double                   ' FE濃度  STD(表示、判定用)
    PUA_DIFF    As Double                   ' 拡散長  PUA値(表示、判定用)
    PUAP_DIFF   As Double                   ' 拡散長  PUA％値(表示、判定用)
    MAX_NR      As Double                   ' NR濃度  最大値(表示、判定用)
    MIN_NR      As Double                   ' NR濃度  最小値(表示、判定用)
    AVE_NR      As Double                   ' NR濃度  平均(表示、判定用)
    CENTER_NR   As Double                   ' NR濃度  中心(表示、判定用)
    PUA_NR      As Double                   ' NR濃度  PUA値(表示、判定用)
    PUAP_NR     As Double                   ' NR濃度  PUA％値(表示、判定用)
    STD_NR      As Double                   ' NR濃度  STD(表示、判定用)
    ''==============================
End Type
''Upd end   2005/06/21 (TCS)T.Terauchi  SPV9点対応  SPV実績ﾃｰﾌﾞﾙ


' X線実績　2009/08 SUMCO Akizuki追加
Public Type typ_TBCMJ021
    CRYNUM As String * 12           ' 結晶番号
    POSITION As Integer             ' 位置
    SMPKBN As String * 1            ' サンプル区分
    TRANCOND As String * 1          ' 処理条件
    TRANCNT As Integer              ' 処理回数
    SMPLNO As Long                  ' サンプルＮｏ
    SMPLUMU As String * 1           ' サンプル有無
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    hinban As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    XRAYX As Single                 ' 結晶面傾き 横方向(X)
    XRAYY As Single                 ' 結晶面傾き 縦方向(Y)
    XRAYXY As Single                ' 結晶面傾き 合成(複合)
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type

''↓Add 2010/01/12 SIRD対応 Y.Hitomi
Public Type typ_TBCMJ022
    CRYNUM          As String * 12          ' 結晶番号
    POSITION        As Integer              ' 位置
    SMPKBN          As String * 1           ' サンプル区分
    TRANCOND        As String * 1           ' 処理条件
    TRANCNT         As Integer              ' 処理回数
    HSFLG           As String * 1           ' 保証フラグ
    SMPLNO          As String * 16          ' サンプルＮｏ
    SMPLUMU         As String * 1           ' サンプル有無
    BLOCKID         As String * 12          ' ブロックID
    SXLID           As String * 13          ' SXLID
    hinban          As String * 8           ' 品番
    REVNUM          As Integer              ' 製品番号改訂番号
    FACTORY         As String * 1           ' 工場
    OPECOND         As String * 1           ' 操業条件
    KRPROCCD        As String * 5           ' 管理工程コード
    PROCCODE        As String * 5           ' 工程コード
    GOUKI           As String * 3           ' 号機
    OSITEM          As String * 4           ' 評価項目
    MAISU           As Integer              ' 評価枚数
    Spec            As String * 10          ' 規格値
    NETSU           As String * 2           ' 熱処理条件
    ET              As String * 3           ' エッチング条件
    MES             As String * 3           ' 計測方法
    DKAN            As String * 10          ' ＤＫアニール条件
    SIRDCNT         As Integer              ' 面内個数
    PLANTCAT        As String * 2           ' 事業所区分
    OSWAFID         As String * 6           ' OSウェハーID
    TSTAFFID        As String * 8           ' 登録社員ID
    REGDATE         As Date                 ' 登録日付
    KSTAFFID        As String * 8           ' 更新社員ID
    UPDDATE         As Date                 ' 更新日付
    SENDFLAG        As String * 1           ' 送信フラグ
    SENDDATE        As Date                 ' 送信日付
End Type
''↑Add 2010/01/12 SIRD対応 Y.Hitomi

'Add Start 2010/12/17 SMPK Miyata
Public Type typ_TBCMJ023
    CRYNUM          As String * 12          ' 結晶番号
    POSITION        As Integer              ' 位置
    SMPKBN          As String * 1           ' サンプル区分
    TRANCOND        As String * 1           ' 処理条件
    TRANCNT         As Integer              ' 処理回数
    SMPLNO          As Long                 ' サンプルＮｏ
    SMPLUMUC        As String               ' サンプル有無(C)
    SMPLUMUCJ       As String               ' サンプル有無(CJ)
    SMPLUMUCJLT     As String               ' サンプル有無(CJLT)
    SMPLUMUCJ2      As String               ' サンプル有無(CJ2)
    hinban          As String * 8           ' 品番
    REVNUM          As Integer              ' 製品番号改訂番号
    factory         As String * 1           ' 工場
    opecond         As String * 1           ' 操業条件
    KRPROCCD        As String * 5           ' 管理工程コード
    PROCCODE        As String * 5           ' 工程コード
    GOUKI           As String * 3           ' 号機
    CPTNJSK         As String * 1           ' C パターン実績
    CDISKJSK        As Integer              ' C Disk半径実績
    CRINGNKJSK      As Integer              ' C Ring内径実績
    CRINGGKJSK      As Integer              ' C Ring外径実績
    C_SZ            As String * 1           ' C 測定条件
    CHANTEI         As String               ' C 判定結果
    CJPTNJSK        As String               ' CJ パターン実績
    CJDISKJSK       As Integer              ' CJ Disk半径実績
    CJRINGNKJSK     As Integer              ' CJ Ring内径実績
    CJRINGGKJSK     As Integer              ' CJ Ring外径実績
    CJBANDNKJSK     As Integer              ' CJ Band内径実績
    CJBANDGKJSK     As Integer              ' CJ Band外径実績
    CJRINGCALC      As Integer              ' CJ Ring幅計算
    CJPICALC        As Integer              ' CJ Pi幅計算
    CJ_NETU         As String * 2           ' CJ 熱処理法
    CJHANTEI        As String               ' CJ 判定結果
    CJBUIUMU        As String               ' CJ 部位別判定有無
    CJDMAXPIC5      As Integer              ' CJ Diskのみパターン Pi幅上限値
    CJRMAXPIC5      As Integer              ' CJ Ringのみパターン Pi幅上限値
    CJDRMAXPIC5     As Integer              ' CJ DiskRingパターン Pi幅上限値
    CJALLMAXDIC5    As Integer              ' CJ 共通Disk半径上限値
    CJALLMINRINC5   As Integer              ' CJ 共通Ring内径下限値
    CJALLMAXRIGC5   As Integer              ' CJ 共通Ring外径上限値
    CJLTPTNJSK      As String               ' CJ(LT) パターン実績
    CJLTDISKJSK     As Integer              ' CJ(LT) Disk半径実績
    CJLTRINGNKJSK   As Integer              ' CJ(LT) Ring内径実績
    CJLTRINGGKJSK   As Integer              ' CJ(LT) Ring外径実績
    CJLTBANDNKJSK   As Integer              ' CJ(LT) Band内径実績
    CJLTBANDGKJSK   As Integer              ' CJ(LT) Band外径実績
    CJLTRINGCALC    As Integer              ' CJ(LT) Ring幅計算
    CJLTPICALC      As Integer              ' CJ(LT) Pi幅計算
    CJLTBANDCALC    As Integer              ' CJ(LT) Band幅計算
    HSXCJLTBND      As Integer              ' CJ(LT) Band幅上限値
    CJLT_NETU       As String * 2           ' CJ(LT) 熱処理法
    CJLTHANTEI      As String               ' CJ(LT) 判定結果
    CJ2PTNJSK       As String               ' CJ2 パターン実績
    CJ2DISKJSK      As Integer              ' CJ2 Disk半径実績
    CJ2RINGNKJSK    As Integer              ' CJ2 Ring内径実績
    CJ2RINGGKJSK    As Integer              ' CJ2 Ring外径実績
    CJ2PICALC       As Integer              ' CJ2 Pi幅計算
    CJ2_NETU        As String * 2           ' CJ2 熱処理法
    CJ2HANTEI       As String               ' CJ2 判定結果
    CJ2BUIUMU       As String               ' CJ2 部位別判定有無
    CJ2DMAXPIC5     As Integer              ' CJ2 Diskのみパターン Pi幅上限値
    CJ2RMAXPIC5     As Integer              ' CJ2 Ringのみパターン Pi幅上限値
    CJ2RMINRINC5    As Integer              ' CJ2 Ringのみパターン Ring内径下限値
    CJ2RMAXRIGC5    As Integer              ' CJ2 Ringのみパターン Ring外径上限値
    CJ2DRMAXPIC5    As Integer              ' CJ2 DiskRingパターン Pi幅上限値
    CJ2DRMINRINC5   As Integer              ' CJ2 DiskRingパターン Ring内径下限値
    CJ2DRMAXRIGC5   As Integer              ' CJ2 DiskRingパターン Ring外径上限値
    TSTAFFID        As String               ' 登録社員ID
    REGDATE         As Date                 ' 登録日付
    TSTAFFIDC       As String               ' 登録社員ID (C)
    REGDATEC        As String               ' 登録日付   (C)
    TSTAFFIDCJ      As String               ' 登録社員ID (CJ)
    REGDATECJ       As String               ' 登録日付   (CJ)
    TSTAFFIDCJLT    As String               ' 登録社員ID (CJLT)
    REGDATECJLT     As String               ' 登録日付   (CJLT)
    TSTAFFIDCJ2     As String               ' 登録社員ID (CJ2)
    REGDATECJ2      As String               ' 登録日付   (CJ2)
    KSTAFFID        As String               ' 更新社員ID
    UPDDATE         As String               ' 更新日付
    KSTAFFIDC       As String               ' 更新社員ID (C)
    UPDDATEC        As String               ' 更新日付   (C)
    KSTAFFIDCJ      As String               ' 更新社員ID (CJ)
    UPDDATECJ       As String               ' 更新日付   (CJ)
    KSTAFFIDCJLT    As String               ' 更新社員ID (CJLT)
    UPDDATECJLT     As String               ' 更新日付   (CJLT)
    KSTAFFIDCJ2     As String               ' 更新社員ID (CJ2)
    UPDDATECJ2      As String               ' 更新日付   (CJ2)
    SENDFLAG        As String               ' 送信フラグ
    SENDDATE        As Date                 ' 送信日付
End Type
'Add End   2010/12/17 SMPK Miyata

' 抜試指示実績
Public Type typ_TBCMW001
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' インゴット位置
    TRANCNT As Integer              ' 処理回数
    CRYLEN As Integer               ' 長さ
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    BLOCKID As String * 12          ' ブロックID
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 結晶情報変更
Public Type typ_TBCMW002
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' インゴット位置
    TRANCNT As Integer              ' 処理回数
    CRYLEN As Integer               ' 長さ
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    BLOCKID As String * 12          ' ブロックID
    DELFLG As String * 1            ' 削除区分
    TOPBDLN As Integer              ' TOP不良長さ
    TOPBDCS As String * 3           ' TOP不良理由
    TAILBDLN As Integer             ' TAIL不良長さ
    TAILBDCS As String * 3          ' TAIL不良理由
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SUMMITSENDFLAG As String * 1    ' SUMMIT送信フラグ
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 抜試変更指示実績
Public Type typ_TBCMW003
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' インゴット位置
    TRANCNT As Integer              ' 処理回数
    CRYLEN As Integer               ' 長さ
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    BLOCKID As String * 12          ' ブロックID
    DELFLG As String * 1            ' 削除区分
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 再抜試指示実績
Public Type typ_TBCMW004
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' インゴット位置
    TRANCNT As Integer              ' 処理回数
    CRYLEN As Integer               ' 長さ
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    BLOCKID As String * 12          ' ブロックID
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' WF総合判定実績
Public Type typ_TBCMW005
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' インゴット位置
    TRANCNT As Integer              ' 処理回数
    CRYLEN As Integer               ' 長さ
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    SXLID As String * 13            ' SXLID
    CODE As String * 1              ' 区分コード
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 振替廃棄実績
Public Type typ_TBCMW006
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' インゴット位置
    TRANCNT As Integer              ' 処理回数
    CRYLEN As Integer               ' 長さ
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    TRANCLS As String * 1           ' 処理区分
    DUNWNUM As String * 8           ' 転用先品番
    DUNWREV As Integer              ' 転用先品番 製品番号改訂番号
    DUNWFACT As String * 1          ' 転用先品番 工場
    DUNWOPCD As String * 1          ' 転用先品番 操業条件
    DUOGNUM As String * 8           ' 転用元品番
    DUOGREV As Integer              ' 転用元品番 製品番号改訂番号
    DUOGFACT As String * 1          ' 転用元品番 工場
    DUOGOPCD As String * 1          ' 転用元品番 操業条件
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    MUKESAKI As String              ' 07/09/05 SPK Tsutsumi Add
End Type


' シングル確定実績
Public Type typ_TBCMW007
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' インゴット位置
    CRYLEN As Integer               ' 長さ
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    SXLID As String * 13            ' シングルID
    SAMPLE_FROM As String * 16      ' サンプルID (From)
    SAMPLE_TO As String * 16        ' サンプルID (To)
    BLOCKID As String * 12          ' ブロックID
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' WFホールド（解除）実績
Public Type typ_TBCMW008
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' インゴット位置
    TRANCNT As Integer              ' 処理回数
    CRYLEN As Integer               ' 長さ
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    SNGLID As String * 13           ' シングルID
    HLDCLASS As String * 1          ' ホールド処理区分
    HLDCAUSE As String * 3          ' ホールド理由
    HLDCMNT As String               ' ホールドコメント
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' WFセンター総合判定測定値
Public Type typ_TBCMW009
    SXLID As String * 13            ' SXLID
    FROMTOKBN As String * 1         ' FROMTO区分
    SAMPLE_FROM As String * 16      ' サンプルID (From)
    SAMPLE_TO As String * 16        ' サンプルID (To)
    BLOCKID As String * 12          ' ブロックID
    KRPROCCD As String * 5          ' 管理工程コード
    PROCCODE As String * 5          ' 工程コード
    WFOI_SMPPOS As Integer          ' WFOIｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFOI_NETSU As String * 2        ' WFOI_熱処理条件
    WFOI_ET As String * 3           ' WFOI_エッチング条件
    WFOI_MES As String * 3          ' WFOI_計測方法
    WFOI_MESDATA1 As Double         ' WFOI_測定データその１
    WFOI_MESDATA2 As Double         ' WFOI_測定データその２
    WFOI_MESDATA3 As Double         ' WFOI_測定データその３
    WFOI_MESDATA4 As Double         ' WFOI_測定データその４
    WFOI_MESDATA5 As Double         ' WFOI_測定データその５
    WFOI_MESDATA6 As Double         ' WFOI_測定データその６
    WFOI_MESDATA7 As Double         ' WFOI_測定データその７
    WFOI_MESDATA8 As Double         ' WFOI_測定データその８
    WFOI_MESDATA9 As Double         ' WFOI_測定データその９
    WFOI_MESDATA10 As Double        ' WFOI_測定データその１０
    WFOI_ORG As Double              ' WFOI_ORG計算結果
    WFRS_SMPPOS As Integer          ' WFRSｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFRS_NETSU As String * 2        ' WFRS_熱処理条件
    WFRS_ET As String * 3           ' WFRS_エッチング条件
    WFRS_MES As String * 3          ' WFRS_計測方法
    WFRS_MESDATA1 As Double         ' WFRS_測定データその１
    WFRS_MESDATA2 As Double         ' WFRS_測定データその２
    WFRS_MESDATA3 As Double         ' WFRS_測定データその３
    WFRS_MESDATA4 As Double         ' WFRS_測定データその４
    WFRS_MESDATA5 As Double         ' WFRS_測定データその５
    WFRS_RRG As Double              ' WFRS_RRG計算結果
    WFDOI_SMPPOS As Integer         ' WFDOIｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）　number(4)
    WFDOI_NETU_1 As String * 2      ' WFDOI_熱処理条件_1
    WFDOI_MES_1 As String * 3       ' WFDOI_計測方法_1
    WFDOI_MESDATA1_1 As Double      ' WFDOI_(ｲﾆｼｬﾙOi-AfterOi）１_1
    WFDOI_MESDATA2_1 As Double      ' WFDOI_(ｲﾆｼｬﾙOi-AfterOi)２_1
    WFDOI_MESDATA3_1 As Double      ' WFDOI_(ｲﾆｼｬﾙOi-AfterOi)３_1
    WFDOI_NETU_2 As String * 2      ' WFDOI_熱処理条件_２
    WFDOI_MES_2 As String * 3       ' WFDOI_計測方法_２
    WFDOI_MESDATA1_2 As Double      ' WFDOI_(ｲﾆｼｬﾙOi-AfterOi）１_２
    WFDOI_MESDATA2_2 As Double      ' WFDOI_(ｲﾆｼｬﾙOi-AfterOi)２_２
    WFDOI_MESDATA3_2 As Double      ' WFDOI_(ｲﾆｼｬﾙOi-AfterOi)３_２
    WFDOI_NETU_3 As String * 2      ' WFDOI_熱処理条件_３
    WFDOI_MES_3 As String * 3       ' WFDOI_計測方法_３
    WFDOI_MESDATA1_3 As Double      ' WFDOI_(ｲﾆｼｬﾙOi-AfterOi）１_３
    WFDOI_MESDATA2_3 As Double      ' WFDOI_(ｲﾆｼｬﾙOi-AfterOi)２_３
    WFDOI_MESDATA3_3 As Double      ' WFDOI_(ｲﾆｼｬﾙOi-AfterOi)３_３
    WFOSF1_SMPPOS As Integer        ' WFOSF1ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFOSF1_NETSU As String * 2      ' WFOSF1_熱処理条件
    WFOSF1_ET As String * 3         ' WFOSF1_エッチング条件
    WFOSF1_MES As String * 3        ' WFOSF1_計測方法
    WFOSF1_MAX As Double            ' WFOSF1_判定時のMAX値_1
    WFOSF1_AVE As Double            ' WFOSF1_判定時のAVE値_1
    WFOSF2_SMPPOS As Integer        ' WFOSF２ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）　number(4)
    WFOSF2_NETSU As String * 2      ' WFOSF2_熱処理条件_２
    WFOSF2_ET As String * 3         ' WFOSF2_エッチング条件_２
    WFOSF2_MES As String * 3        ' WFOSF2_計測方法_２
    WFOSF2_MAX As Double            ' WFOSF2_判定時のMAX値_２
    WFOSF2_AVE As Double            ' WFOSF2_判定時のAVE値_２
    WFOSF3_SMPPOS As Integer        ' WFOSF３ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFOSF3_NETSU As String * 2      ' WFOSF3_熱処理条件_３
    WFOSF3_ET As String * 3         ' WFOSF3_エッチング条件_３
    WFOSF3_MES As String * 3        ' WFOSF3_計測方法_３
    WFOSF3_MAX As Double            ' WFOSF3_判定時のMAX値_３
    WFOSF3_AVE As Double            ' WFOSF3_判定時のAVE値_３
    WFOSF4_SMPPOS As Integer        ' WFOSF４ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFOSF4_NETSU As String * 2      ' WFOSF4_熱処理条件_４
    WFOSF4_ET As String * 3         ' WFOSF4_エッチング条件_４
    WFOSF4_MES As String * 3        ' WFOSF4_計測方法_４
    WFOSF4_MAX As Double            ' WFOSF4_判定時のMAX値_４
    WFOSF4_AVE As Double            ' WFOSF4_判定時のAVE値_４
    WFBMD1_SMPPOS As Integer        ' WFBMD1ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFBMD1_NETSU As String * 2      ' WFBMD1_熱処理条件_1
    WFBMD1_ET As String * 3         ' WFBMD1_エッチング条件_1
    WFBMD1_MES As String * 3        ' WFBMD1_計測方法_1
    WFBMD1_MAX As Double            ' WFBMD1_判定時のMAX値_1
    WFBMD1_AVE As Double            ' WFBMD1_判定時のAVE値_1
    WFBMD2_SMPPOS As Integer        ' WFBMD２ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFBMD2_NETSU As String * 2      ' WFBMD2_熱処理条件_２
    WFBMD2_ET As String * 3         ' WFBMD2_エッチング条件_２
    WFBMD2_MES As String * 3        ' WFBMD2_計測方法_２
    WFBMD2_MAX As Double            ' WFBMD2_判定時のMAX値_２
    WFBMD2_AVE As Double            ' WFBMD2_判定時のAVE値_２
    WFBMD3_SMPPOS As Integer        ' WFBMD３ｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFBMD3_NETSU As String * 2      ' WFBMD3_熱処理条件_３
    WFBMD3_ET As String * 3         ' WFBMD3_エッチング条件_３
    WFBMD3_MES As String * 3        ' WFBMD3_計測方法_３
    WFBMD3_MAX As Double            ' WFBMD3_判定時のMAX値_３
    WFBMD3_AVE As Double            ' WFBMD3_判定時のAVE値_３
    WFDSOD_SMPPOS As Integer        ' WFDSODｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFDSOD_NETSU As String * 2      ' WFDSOD_熱処理条件
    WFDSOD_ET As String * 3         ' WFDSOD_エッチング条件
    WFDSOD_MES As String * 3        ' WFDSOD_計測方法
    WFDSOD_TOTAL As Integer         ' WFDSOD_判定時のTOTAL値
    WFSPV_SMPPOS As Integer         ' WFSPVｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFSPV_NETSU As String * 2       ' WFSVP_熱処理条件
    WFSPV_ET As String * 3          ' WFSPV_エッチング条件
    WFSPV_MES As String * 3         ' WFSPV_計測方法
    WFSPV_KST_MAX As Double         ' WFSPV_拡散長判定時のMAX値
    WFSPV_KST_AVE As Double         ' WFSPV_拡散長判定時のAVE値
    WFSPV_KST_MIN As Double         ' WFSPV_拡散長判定時のMIN値
    WFSPV_FE_MAX As Double          ' WFSPV_Fe濃度判定時のMAX値
    WFSPV_FE_AVE As Double          ' WFSPV_Fe濃度判定時のAVE値
    WFSPV_FE_MIN As Double          ' WFSPV_Fe濃度判定時のMIN値
    WFDZ_SMPPOS As Integer          ' WFDZｻﾝﾌﾟﾙ-ID測定位置（SXL位置情報）
    WFDZ_NETSU As String * 2        ' WFDZ_熱処理条件
    WFDZ_ET As String * 3           ' WFDZ_エッチング条件
    WFDZ_MES As String * 3          ' WFDZ_計測方法
    WFDZ_MAX As Double              ' WFDZ_判定時のMAX値_
    WFDZ_AVE As Double              ' WFDZ_判定時のAVE値
    TSTAFFID As String * 8          ' 登録社員ID
    REGDATE As Date                 ' 登録日付
    KSTAFFID As String * 8          ' 更新社員ID
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
    WFBMD1_MIN As Double            ' WFBMD1_判定時のMIN値_1　▼2003/05/14 ooba
    WFBMD1_MBNP As Double           ' WFBMD1_判定時の面内分布
    WFBMD2_MIN As Double            ' WFBMD2_判定時のMIN値_2
    WFBMD2_MBNP As Double           ' WFBMD2_判定時の面内分布
    WFBMD3_MIN As Double            ' WFBMD3_判定時のMIN値_3
    WFBMD3_MBNP As Double           ' WFBMD3_判定時の面内分布
    WFDZ_MIN As Double              ' WFDZ_判定時のMIN値
    WFOSF1_PATKBNP1 As Double       ' WF_OSF1_パターン区分１位置
    WFOSF1_PATKBNWID1 As Double     ' WF_OSF1_パターン区分１幅
    WFOSF1_PATKBNRD1 As String * 1  ' WF_OSF1_パターン区分１Ring/Disk
    WFOSF1_PATKBNP2 As Double       ' WF_OSF1_パターン区分２位置
    WFOSF1_PATKBNWID2 As Double     ' WF_OSF1_パターン区分２幅
    WFOSF1_PATKBNRD2 As String * 1  ' WF_OSF1_パターン区分２Ring/Disk
    WFOSF1_PATKBNP3 As Double       ' WF_OSF1_パターン区分３位置
    WFOSF1_PATKBNWID3 As Double     ' WF_OSF1_パターン区分３幅
    WFOSF1_PATKBNRD3 As String * 1  ' WF_OSF1_パターン区分３Ring/Disk
    WFOSF2_PATKBNP1 As Double       ' WF_OSF2_パターン区分１位置
    WFOSF2_PATKBNWID1 As Double     ' WF_OSF2_パターン区分１幅
    WFOSF2_PATKBNRD1 As String * 1  ' WF_OSF2_パターン区分１Ring/Disk
    WFOSF2_PATKBNP2 As Double       ' WF_OSF2_パターン区分２位置
    WFOSF2_PATKBNWID2 As Double     ' WF_OSF2_パターン区分２幅
    WFOSF2_PATKBNRD2 As String * 1  ' WF_OSF2_パターン区分２Ring/Disk
    WFOSF2_PATKBNP3 As Double       ' WF_OSF2_パターン区分３位置
    WFOSF2_PATKBNWID3 As Double     ' WF_OSF2_パターン区分３幅
    WFOSF2_PATKBNRD3 As String * 1  ' WF_OSF2_パターン区分３Ring/Disk
    WFOSF3_PATKBNP1 As Double       ' WF_OSF3_パターン区分１位置
    WFOSF3_PATKBNWID1 As Double     ' WF_OSF3_パターン区分１幅
    WFOSF3_PATKBNRD1 As String * 1  ' WF_OSF3_パターン区分１Ring/Disk
    WFOSF3_PATKBNP2 As Double       ' WF_OSF3_パターン区分２位置
    WFOSF3_PATKBNWID2 As Double     ' WF_OSF3_パターン区分２幅
    WFOSF3_PATKBNRD2 As String * 1  ' WF_OSF3_パターン区分２Ring/Disk
    WFOSF3_PATKBNP3 As Double       ' WF_OSF3_パターン区分３位置
    WFOSF3_PATKBNWID3 As Double     ' WF_OSF3_パターン区分３幅
    WFOSF3_PATKBNRD3 As String * 1  ' WF_OSF3_パターン区分３Ring/Disk
    WFOSF4_PATKBNP1 As Double       ' WF_OSF4_パターン区分１位置
    WFOSF4_PATKBNWID1 As Double     ' WF_OSF4_パターン区分１幅
    WFOSF4_PATKBNRD1 As String * 1  ' WF_OSF4_パターン区分１Ring/Disk
    WFOSF4_PATKBNP2 As Double       ' WF_OSF4_パターン区分２位置
    WFOSF4_PATKBNWID2 As Double     ' WF_OSF4_パターン区分２幅
    WFOSF4_PATKBNRD2 As String * 1  ' WF_OSF4_パターン区分２Ring/Disk
    WFOSF4_PATKBNP3 As Double       ' WF_OSF4_パターン区分３位置
    WFOSF4_PATKBNWID3 As Double     ' WF_OSF4_パターン区分３幅
    WFOSF4_PATKBNRD3 As String * 1  ' WF_OSF4_パターン区分３Ring/Disk　▲2003/05/14 ooba
End Type


' 顧客仕様管理
Public Type typ_TBCME001
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KMGSRRNO As String * 9          ' 購管理仕様登録依頼番号
    KMGSHN As String * 7            ' 購管理社内品名
    KMGSNRNO As Integer             ' 購管理社内品名改訂番号
    KMGSTFNO As String * 8          ' 購管理社員Ｎｏ
    KMGCSGRP As String * 3          ' 購管理顧客グループ
    KMGCSCOD As String * 8          ' 購管理顧客コード
    KMGCSHN As String               ' 購管理顧客品名
    KMGKBNNO As String * 3          ' 購管理区分Ｎｏ
    COPYMSHN As String * 9          ' コピー元社内品名
    CONFLAG As String * 1           ' 確認フラグ
    REINFLAG As String * 1          ' 再付与フラグ
    KMGCSGNO As String              ' 購管理顧客一般仕様番号
    KMGCSKNO As String              ' 購管理顧客個別仕様番号
    KMGCSBNO As String              ' 購管理顧客部品番号
    KMGTTBNO As Integer             ' 購管理対応テーブル枝番
    DIHYSHN As String * 7           ' 代表社内品名
    KMGSHKBN As String * 1          ' 購管理製品区分
    KMGWKBN As String * 2           ' 購管理製法区分
    KMGYKBN As String * 6           ' 購管理用途区分
    KMGDMKBN As String * 1          ' 購管理直径区分
    KMGRSKBN As String * 1          ' 購管理ＲＳ区分
    KMGSNBEF As String * 7          ' 購管理社内品名前回
    KMGSDATE As Date                ' 購管理適用開始日
    KMGEDATE As Date                ' 購管理適用終了日
    KMGTDATE As Date                ' 購管理登録月日
    KMGRKBNK As String * 1          ' 購管理比抵抗区分種別
    KMGAKBUM As String * 1          ' 購管理厚み区分有無
    KMGOXKBN As String * 1          ' 購管理酸素区分
    KMGIDKBU As String * 1          ' 購管理ＩＤ区分有無
    KMGIGKBN As String * 1          ' 購管理ＩＧ区分
    KMGWRBKU As String * 1          ' 購管理ＷＡＲＰ分布規格有無
    KMGSBKUM As String * 1          ' 購管理反り分布規格有無
    KMGFBKUM As String * 1          ' 購管理平坦分布規格有無
    KMGFPSUM As String * 1          ' 購管理平坦Ｐサイト有無
    KMGFOFUM As String * 1          ' 購管理平坦オフセット有無
    KMGNCKBN As String * 1          ' 購管理ノッチ区分
    KMGMKBKU As String * 1          ' 購管理面検欠陥分布規格有無
    KMGCMPKU As String * 1          ' 購管理ＣＭＰ加工有無
    KMGSZKBN As String * 1          ' 購管理支給材料区分
    KMGSZMUM As String * 1          ' 購管理支給材料面取有無
    KMGEPKBN As String              ' 購管理ＥＰ基板名
    KMGEPSKN As String              ' 購管理ＥＰ製品型名
    KMGEPSKB As String * 2          ' 購管理ＥＰ仕上区分
    KMGEPRKU As String * 1          ' 購管理ＥＰ比抵抗区分有無
    KMGEPAKU As String * 1          ' 購管理ＥＰ厚み区分有無
    KMGEPIKU As String * 1          ' 購管理ＥＰＩＤ区分有無
    KMGKZKBN As String * 1          ' 購管理購入材料区分
    KMGSKBN As String * 1           ' 購管理製作区分
    KMGTRKSI As String * 1          ' 購管理ＴＲＫ＃指定
    KMGHNCDS As String * 1          ' 購管理品名コード＿製
    KMGHNCDT As String * 1          ' 購管理品名コード＿直
    KMGHNCDD As String * 1          ' 購管理品名コード＿ド
    KMGHNCDK As String * 1          ' 購管理品名コード＿結
    KMGHNCDF As String * 1          ' 購管理品名コード＿フ
    KMGHNCDN As String * 1          ' 購管理品名コード＿内
    KMGHNCDH As String * 1          ' 購管理品名コード＿薄
    KMGNOTE As String               ' 購管理特記
    KMGRS1N As String               ' 購管理予備１＿内
    KMGRS1Y As String               ' 購管理予備１＿用
    KMGRS2N As String               ' 購管理予備２＿内
    KMGRS2Y As String               ' 購管理予備２＿用
    KMGRS3N As String               ' 購管理予備３＿内
    KMGRS3Y As String               ' 購管理予備３＿用
    KMGRS4N As String               ' 購管理予備４＿内
    KMGRS4Y As String               ' 購管理予備４＿用
    KMGRS5N As String               ' 購管理予備５＿内
    KMGRS5Y As String               ' 購管理予備５＿用
    KMGRS6N As String               ' 購管理予備６＿内
    KMGRS6Y As String               ' 購管理予備６＿用
    KMGRS7N As String               ' 購管理予備７＿内
    KMGRS7Y As String               ' 購管理予備７＿用
    KMGRS8N As String               ' 購管理予備８＿内
    KMGRS8Y As String               ' 購管理予備８＿用
    KMGRS9N As String               ' 購管理予備９＿内
    KMGRS9Y As String               ' 購管理予備９＿用
    KMGRS10N As String              ' 購管理予備１０＿内
    KMGRS10Y As String              ' 購管理予備１０＿用
    KMGWFLVS As Integer             ' 購管理ＷＦ水準数
    KMGWFLVN As String * 2          ' 購管理ＷＦ水準名
    KMGWFLVC As String              ' 購管理ＷＦ水準内容
    KMGEPLVS As Integer             ' 購管理ＥＰ水準数
    KMGEPLVN As String * 2          ' 購管理ＥＰ水準名
    KMGEPLVY As String              ' 購管理ＥＰ水準内容
    KMGSSREC As String * 3          ' 購管理仕様設定記録
    SSMGKBN As String * 1           ' 生産管理確認区分
    HSHSKBN As String * 1           ' 品質保証確認区分
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 顧客仕様納入ﾃﾞｰﾀ
Public Type typ_TBCME002
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KMGSRRNO As String * 9          ' 購管理仕様登録依頼番号
    KNURSMAX As Integer             ' 購納入ロットサイズ上限
    KNURSMIN As Integer             ' 購納入ロットサイズ下限
    KNUKRMAX As Integer             ' 購納入構成ロット上限
    KNUNPACK As String * 1          ' 購納入内装パック＿種
    KNUNPACT As String * 1          ' 購納入内装パック＿直
    KNUNPACS As String * 1          ' 購納入内装パック＿連
    KNUNPACR As String * 1          ' 購納入内装パック＿リ
    KNUNZWAY As String * 1          ' 購納入内装充填方法＿方
    KNUNZWYH As String * 1          ' 購納入内装充填方法＿端
    KNUNZWYT As String * 1          ' 購納入内装充填方法＿単
    KNUNZWYW As String * 1          ' 購納入内装充填方法＿Ｗ
    KNUNPAC As String * 1           ' 購納入内装包装
    KNUPACKZ As Integer             ' 購納入パック充填数
    KNUSEALA As String * 1          ' 購納入シール貼付
    KNUSEALK As String * 1          ' 購納入シール種類
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 顧客仕様SXLﾃﾞｰﾀ添付
Public Type typ_TBCME003
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KMGSHN As String * 7            ' 購管理社内品名
    KMGSNRNO As Integer             ' 購管理社内品名改訂番号
    KNUSXDTS As String * 1          ' 購納入ＳＸデータ添付形式＿添
    KNUSXDTT As String * 1          ' 購納入ＳＸデータ添付形式＿特
    KNUSXDTY As String * 1          ' 購納入ＳＸデータ添付形式＿予
    KTSXRTK As String * 1           ' 購添ＳＸ比抵抗値＿コ
    KTSXRTD As String * 1           ' 購添ＳＸ比抵抗値＿デ
    KTSXRTS As String * 1           ' 購添ＳＸ比抵抗値＿サ
    KTSXRMK As String * 1           ' 購添ＳＸ比抵抗面内分布＿コ
    KTSXRMD As String * 1           ' 購添ＳＸ比抵抗面内分布＿デ
    KTSXRMS As String * 1           ' 購添ＳＸ比抵抗面内分布＿サ
    KTSXRM2K As String * 1          ' 購添ＳＸ比抵抗面内分布２＿コ
    KTSXRM2D As String * 1          ' 購添ＳＸ比抵抗面内分布２＿デ
    KTSXRM2S As String * 1          ' 購添ＳＸ比抵抗面内分布２＿サ
    KTSXDIMK As String * 1          ' 購添ＳＸ直径＿コ
    KTSXDIMD As String * 1          ' 購添ＳＸ直径＿デ
    KTSXDIMS As String * 1          ' 購添ＳＸ直径＿サ
    KTSXTMK As String * 1           ' 購添ＳＸ転位密度＿コ
    KTSXTMD As String * 1           ' 購添ＳＸ転位密度＿デ
    KTSXTMS As String * 1           ' 購添ＳＸ転位密度＿サ
    KTSXLTK As String * 1           ' 購添ＳＸライフタイム＿コ
    KTSXLTD As String * 1           ' 購添ＳＸライフタイム＿デ
    KTSXLTS As String * 1           ' 購添ＳＸライフタイム＿サ
    KTSXCNK As String * 1           ' 購添ＳＸ炭素濃度＿コ
    KTSXCND As String * 1           ' 購添ＳＸ炭素濃度＿デ
    KTSXCNS As String * 1           ' 購添ＳＸ炭素濃度＿サ
    KTSXONK As String * 1           ' 購添ＳＸ酸素濃度＿コ
    KTSXOND As String * 1           ' 購添ＳＸ酸素濃度＿デ
    KTSXONS As String * 1           ' 購添ＳＸ酸素濃度＿サ
    KTSXOS1K As String * 1          ' 購添ＳＸＯＳＦ１＿コ
    KTSXOS1D As String * 1          ' 購添ＳＸＯＳＦ１＿デ
    KTSXOS1S As String * 1          ' 購添ＳＸＯＳＦ１＿サ
    KTSXOS2K As String * 1          ' 購添ＳＸＯＳＦ２＿コ
    KTSXOS2D As String * 1          ' 購添ＳＸＯＳＦ２＿デ
    KTSXOS2S As String * 1          ' 購添ＳＸＯＳＦ２＿サ
    KTSXBM1K As String * 1          ' 購添ＳＸＢＭＤ１＿コ
    KTSXBM1D As String * 1          ' 購添ＳＸＢＭＤ１＿デ
    KTSXBM1S As String * 1          ' 購添ＳＸＢＭＤ１＿サ
    KTSXBM2K As String * 1          ' 購添ＳＸＢＭＤ２＿コ
    KTSXBM2D As String * 1          ' 購添ＳＸＢＭＤ２＿デ
    KTSXBM2S As String * 1          ' 購添ＳＸＢＭＤ２＿サ
    KTSXDSOK As String * 1          ' 購添ＳＸＤＳＯＤ＿コ
    KTSXDSOD As String * 1          ' 購添ＳＸＤＳＯＤ＿デ
    KTSXDSOS As String * 1          ' 購添ＳＸＤＳＯＤ＿サ
    KTSXFPDK As String * 1          ' 購添ＳＸＦＰＤ＿コ
    KTSXFPDD As String * 1          ' 購添ＳＸＦＰＤ＿デ
    KTSXFPDS As String * 1          ' 購添ＳＸＦＰＤ＿サ
    KTSXSRK As String * 1           ' 購添ＳＸＳＲ＿コ
    KTSXSRD As String * 1           ' 購添ＳＸＳＲ＿デ
    KTSXSRS As String * 1           ' 購添ＳＸＳＲ＿サ
    KTSXBNK As String * 1           ' 購添ＳＸＢ濃度＿コ
    KTSXBND As String * 1           ' 購添ＳＸＢ濃度＿デ
    KTSXBNS As String * 1           ' 購添ＳＸＢ濃度＿サ
    KTSXSMP As String * 1           ' 購添ＳＸ製品位置＿コ
    KTSXMPD As String * 1           ' 購添ＳＸ製品位置＿デ
    KTSXMPS As String * 1           ' 購添ＳＸ製品位置＿サ
    KTSXODNK As String * 1          ' 購添ＳＸ酸素析出濃度＿コ
    KTSXODND As String * 1          ' 購添ＳＸ酸素析出濃度＿デ
    KTSXODNS As String * 1          ' 購添ＳＸ酸素析出濃度＿サ
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 顧客仕様WFﾃﾞｰﾀ添付
Public Type typ_TBCME004
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KMGSHN As String * 7            ' 購管理社内品名
    KMGSNRNO As Integer             ' 購管理社内品名改訂番号
    KNWFDTFS As String * 1          ' 購納入ＷＦデータ添付形式＿添
    KNWFDTFT As String * 1          ' 購納入ＷＦデータ添付形式＿特
    KNWFDTFY As String * 1          ' 購納入ＷＦデータ添付形式＿予
    KTWFRTK As String * 1           ' 購添ＷＦ比抵抗値＿コ
    KTWFRTD As String * 1           ' 購添ＷＦ比抵抗値＿デ
    KTWFRTS As String * 1           ' 購添ＷＦ比抵抗値＿サ
    KTWFRMK As String * 1           ' 購添ＷＦ比抵抗面内分布＿コ
    KTWFRMD As String * 1           ' 購添ＷＦ比抵抗面内分布＿デ
    KTWFRMS As String * 1           ' 購添ＷＦ比抵抗面内分布＿サ
    KTWFRM2K As String * 1          ' 購添ＷＦ比抵抗面内分布２＿コ
    KTWFRM2D As String * 1          ' 購添ＷＦ比抵抗面内分布２＿デ
    KTWFRM2S As String * 1          ' 購添ＷＦ比抵抗面内分布２＿サ
    KTWFDIMK As String * 1          ' 購添ＷＦ直径＿コ
    KTWFDIMD As String * 1          ' 購添ＷＦ直径＿デ
    KTWFDIMS As String * 1          ' 購添ＷＦ直径＿サ
    KTWFSAK As String * 1           ' 購添ＷＦ仕上厚＿コ
    KTWFSAD As String * 1           ' 購添ＷＦ仕上厚＿デ
    KTWFSAS As String * 1           ' 購添ＷＦ仕上厚＿サ
    KTWFARK As String * 1           ' 購添ＷＦ厚面内範囲＿コ
    KTWFARD As String * 1           ' 購添ＷＦ厚面内範囲＿デ
    KTWFARS As String * 1           ' 購添ＷＦ厚面内範囲＿サ
    KTWFWARK As String * 1          ' 購添ＷＦＷＡＲＰ＿コ
    KTWFWARD As String * 1          ' 購添ＷＦＷＡＲＰ＿デ
    KTWFWARS As String * 1          ' 購添ＷＦＷＡＲＰ＿サ
    KTWFSK As String * 1            ' 購添ＷＦ反り＿コ
    KTWFSD As String * 1            ' 購添ＷＦ反り＿デ
    KTWFSS As String * 1            ' 購添ＷＦ反り＿サ
    KTWFGBK As String * 1           ' 購添ＷＦ平坦ＧＢ＿コ
    KTWFGBD As String * 1           ' 購添ＷＦ平坦ＧＢ＿デ
    KTWFGBS As String * 1           ' 購添ＷＦ平坦ＧＢ＿サ
    KTWFGFRK As String * 1          ' 購添ＷＦ平坦ＧＦＲ＿コ
    KTWFGFRD As String * 1          ' 購添ＷＦ平坦ＧＦＲ＿デ
    KTWFGFRS As String * 1          ' 購添ＷＦ平坦ＧＦＲ＿サ
    KTWFGFDK As String * 1          ' 購添ＷＦ平坦ＧＦＤ＿コ
    KTWFGFDD As String * 1          ' 購添ＷＦ平坦ＧＦＤ＿デ
    KTWFGFDS As String * 1          ' 購添ＷＦ平坦ＧＦＤ＿サ
    KTWFSBK As String * 1           ' 購添ＷＦ平坦ＳＢ＿コ
    KTWFSBD As String * 1           ' 購添ＷＦ平坦ＳＢ＿デ
    KTWFSBS As String * 1           ' 購添ＷＦ平坦ＳＢ＿サ
    KTWFSFK As String * 1           ' 購添ＷＦ平坦ＳＦ＿コ
    KTWFSFD As String * 1           ' 購添ＷＦ平坦ＳＦ＿デ
    KTWFSFS As String * 1           ' 購添ＷＦ平坦ＳＦ＿サ
    KTWFGBPK As String * 1          ' 購添ＷＦ平坦ＧＢＰＵＡ＿コ
    KTWFGBPD As String * 1          ' 購添ＷＦ平坦ＧＢＰＵＡ＿デ
    KTWFGBPS As String * 1          ' 購添ＷＦ平坦ＧＢＰＵＡ＿サ
    KTWFGFPK As String * 1          ' 購添ＷＦ平坦ＧＦＲＰＵＡ＿コ
    KTWFGFPD As String * 1          ' 購添ＷＦ平坦ＧＦＲＰＵＡ＿デ
    KTWFGFPS As String * 1          ' 購添ＷＦ平坦ＧＦＲＰＵＡ＿サ
    KTWFGDPK As String * 1          ' 購添ＷＦ平坦ＧＦＤＰＵＡ＿コ
    KTWFGDPD As String * 1          ' 購添ＷＦ平坦ＧＦＤＰＵＡ＿デ
    KTWFGDPS As String * 1          ' 購添ＷＦ平坦ＧＦＤＰＵＡ＿サ
    KTWFSBPK As String * 1          ' 購添ＷＦ平坦ＳＢＰＵＡ＿コ
    KTWFSBPD As String * 1          ' 購添ＷＦ平坦ＳＢＰＵＡ＿デ
    KTWFSBPS As String * 1          ' 購添ＷＦ平坦ＳＢＰＵＡ＿サ
    KTWFSFPK As String * 1          ' 購添ＷＦ平坦ＳＦＰＵＡ＿コ
    KTWFSFPD As String * 1          ' 購添ＷＦ平坦ＳＦＰＵＡ＿デ
    KTWFSFPS As String * 1          ' 購添ＷＦ平坦ＳＦＰＵＡ＿サ
    KTWFBDK As String * 1           ' 購添ＷＦＢＤ＿コ
    KTWFBDD As String * 1           ' 購添ＷＦＢＤ＿デ
    KTWFBDS As String * 1           ' 購添ＷＦＢＤ＿サ
    KTWFMKK As String * 1           ' 購添ＷＦ面検欠陥＿コ
    KTWFMKD As String * 1           ' 購添ＷＦ面検欠陥＿デ
    KTWFMKS As String * 1           ' 購添ＷＦ面検欠陥＿サ
    KTWFOTAK As String * 1          ' 購添ＷＦ酸化膜耐圧＿コ
    KTWFOTAD As String * 1          ' 購添ＷＦ酸化膜耐圧＿デ
    KTWFOTAS As String * 1          ' 購添ＷＦ酸化膜耐圧＿サ
    KTWFARAK As String * 1          ' 購添ＷＦ表面粗さ＿コ
    KTWFARAD As String * 1          ' 購添ＷＦ表面粗さ＿デ
    KTWFARAS As String * 1          ' 購添ＷＦ表面粗さ＿サ
    KTWFLTK As String * 1           ' 購添ＷＦライフタイム＿コ
    KTWFLTD As String * 1           ' 購添ＷＦライフタイム＿デ
    KTWFLTS As String * 1           ' 購添ＷＦライフタイム＿サ
    KTWFCNK As String * 1           ' 購添ＷＦ炭素濃度＿コ
    KTWFCND As String * 1           ' 購添ＷＦ炭素濃度＿デ
    KTWFCNS As String * 1           ' 購添ＷＦ炭素濃度＿サ
    KTWFONK As String * 1           ' 購添ＷＦ酸素濃度＿コ
    KTWFOND As String * 1           ' 購添ＷＦ酸素濃度＿デ
    KTWFONS As String * 1           ' 購添ＷＦ酸素濃度＿サ
    KTWFOBK As String * 1           ' 購添ＷＦ酸素面内分布＿コ
    KTWFOBD As String * 1           ' 購添ＷＦ酸素面内分布＿デ
    KTWFOBS As String * 1           ' 購添ＷＦ酸素面内分布＿サ
    KTWFOS1K As String * 1          ' 購添ＷＦＯＳＦ１＿コ
    KTWFOS1D As String * 1          ' 購添ＷＦＯＳＦ１＿デ
    KTWFOS1S As String * 1          ' 購添ＷＦＯＳＦ１＿サ
    KTWFOS2K As String * 1          ' 購添ＷＦＯＳＦ２＿コ
    KTWFOS2D As String * 1          ' 購添ＷＦＯＳＦ２＿デ
    KTWFOS2S As String * 1          ' 購添ＷＦＯＳＦ２＿サ
    KTWFOS3K As String * 1          ' 購添ＷＦＯＳＦ３＿コ
    KTWFOS3D As String * 1          ' 購添ＷＦＯＳＦ３＿デ
    KTWFOS3S As String * 1          ' 購添ＷＦＯＳＦ３＿サ
    KTWFOS4K As String * 1          ' 購添ＷＦＯＳＦ４＿コ
    KTWFOS4D As String * 1          ' 購添ＷＦＯＳＦ４＿デ
    KTWFOS4S As String * 1          ' 購添ＷＦＯＳＦ４＿サ
    KTWFBM1K As String * 1          ' 購添ＷＦＢＭＤ１＿コ
    KTWFBM1D As String * 1          ' 購添ＷＦＢＭＤ１＿デ
    KTWFBM1S As String * 1          ' 購添ＷＦＢＭＤ１＿サ
    KTWFBM2K As String * 1          ' 購添ＷＦＢＭＤ２＿コ
    KTWFBM2D As String * 1          ' 購添ＷＦＢＭＤ２＿デ
    KTWFBM2S As String * 1          ' 購添ＷＦＢＭＤ２＿サ
    KTWFBM3K As String * 1          ' 購添ＷＦＢＭＤ３＿コ
    KTWFBM3D As String * 1          ' 購添ＷＦＢＭＤ３＿デ
    KTWFBM3S As String * 1          ' 購添ＷＦＢＭＤ３＿サ
    KTWFOSPK As String * 1          ' 購添ＷＦＯＳＰ＿コ
    KTWFOSPD As String * 1          ' 購添ＷＦＯＳＰ＿デ
    KTWFOSPS As String * 1          ' 購添ＷＦＯＳＰ＿サ
    KTWFDZOK As String * 1          ' 購添ＷＦＤＺ析出酸素濃度＿コ
    KTWFDZOD As String * 1          ' 購添ＷＦＤＺ析出酸素濃度＿デ
    KTWFDZOS As String * 1          ' 購添ＷＦＤＺ析出酸素濃度＿サ
    KTWFKMHK As String * 1          ' 購添ＷＦ結晶面傾方向＿コ
    KTWFKMHD As String * 1          ' 購添ＷＦ結晶面傾方向＿デ
    KTWFKMHS As String * 1          ' 購添ＷＦ結晶面傾方向＿サ
    KTWFOFKL As String * 1          ' 購添ＷＦＯＦ１長さ＿コ
    KTWFOSDL As String * 1          ' 購添ＷＦＯＦ１長さ＿デ
    KTWFOFSL As String * 1          ' 購添ＷＦＯＦ１長さ＿サ
    KTWFMWK As String * 1           ' 購添ＷＦ面取巾＿コ
    KTWFMWD As String * 1           ' 購添ＷＦ面取巾＿デ
    KTWFMWS As String * 1           ' 購添ＷＦ面取巾＿サ
    KTWFMKWK As String * 1          ' 購添ＷＦ無欠陥層巾＿コ
    KTWFMKWD As String * 1          ' 購添ＷＦ無欠陥層巾＿デ
    KTWFMKWS As String * 1          ' 購添ＷＦ無欠陥層巾＿サ
    KTWFNOXK As String * 1          ' 購添ＷＦ熱酸化膜厚＿コ
    KTWFNOXD As String * 1          ' 購添ＷＦ熱酸化膜厚＿デ
    KTWFNOXS As String * 1          ' 購添ＷＦ熱酸化膜厚＿サ
    KTWFPSK As String * 1           ' 購添ＷＦポリシリ厚＿コ
    KTWFPSD As String * 1           ' 購添ＷＦポリシリ厚＿デ
    KTWFPSS As String * 1           ' 購添ＷＦポリシリ厚＿サ
    KTWFCVDK As String * 1          ' 購添ＷＦＣＶＤ厚＿コ
    KTWFCVDD As String * 1          ' 購添ＷＦＣＶＤ厚＿デ
    KTWFCVDS As String * 1          ' 購添ＷＦＣＶＤ厚＿サ
    KTWFSONK As String * 1          ' 購添ＷＦ析出酸素濃度＿コ
    KTWFSOND As String * 1          ' 購添ＷＦ析出酸素濃度＿デ
    KTWFSONS As String * 1          ' 購添ＷＦ析出酸素濃度＿サ
    KTWFOFPK As String * 1          ' 購添ＷＦＯＦ１位置＿コ
    KTWFOFPD As String * 1          ' 購添ＷＦＯＦ１位置＿デ
    KTWFOFPS As String * 1          ' 購添ＷＦＯＦ１位置＿サ
    KTWFGDK As String * 1           ' 購添ＷＦＧＤ＿コ
    KTWFGDD As String * 1           ' 購添ＷＦＧＤ＿デ
    KTWFGDS As String * 1           ' 購添ＷＦＧＤ＿サ
    KTWFDSOK As String * 1          ' 購添ＷＦＤＳＯＤ＿コ
    KTWFDSOD As String * 1          ' 購添ＷＦＤＳＯＤ＿デ
    KTWFDSOS As String * 1          ' 購添ＷＦＤＳＯＤ＿サ
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 顧客仕様SXLﾃﾞｰﾀ１
Public Type typ_TBCME005
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KMGSRRNO As String * 9          ' 購管理仕様登録依頼番号
    CONFLAG As String * 1           ' 確認フラグ
    REINFLAG As String * 1          ' 再付与フラグ
    KSXTYPE As String * 1           ' 購ＳＸタイプ
    KSXTYPKW As String * 1          ' 購ＳＸタイプ検査方法
    KSXTYPKB As String * 1          ' 購ＳＸタイプ検査区分
    KSXDOP As String * 1            ' 購ＳＸドーパント
    KSXRMIN As Double               ' 購ＳＸ比抵抗下限
    KSXRMAX As Double               ' 購ＳＸ比抵抗上限
    KSXRUNIT As String * 1          ' 購ＳＸ比抵抗単位
    KSXRSPOH As String * 1          ' 購ＳＸ比抵抗測定位置＿方
    KSXRSPOT As String * 1          ' 購ＳＸ比抵抗測定位置＿点
    KSXRSPOI As String * 1          ' 購ＳＸ比抵抗測定位置＿位
    KSXRHWYT As String * 1          ' 購ＳＸ比抵抗保証方法＿対
    KSXRHWYS As String * 1          ' 購ＳＸ比抵抗保証方法＿処
    KSXRKKBN As String * 1          ' 購ＳＸ比抵抗検査区分
    KSXRKWAY As String * 2          ' 購ＳＸ比抵抗検査方法
    KSXRKHNM As String * 1          ' 購ＳＸ比抵抗検査頻度＿枚
    KSXRKHNI As String * 1          ' 購ＳＸ比抵抗検査頻度＿位
    KSXRKHNH As String * 1          ' 購ＳＸ比抵抗検査頻度＿保
    KSXRKHNS As String * 1          ' 購ＳＸ比抵抗検査頻度＿試
    KSXRMCAL As String * 1          ' 購ＳＸ比抵抗面内計算
    KSXRMBNP As Double              ' 購ＳＸ比抵抗面内分布
    KSXRMCL2 As String * 1          ' 購ＳＸ比抵抗面内計算２
    KSXRMBP2 As Double              ' 購ＳＸ比抵抗面内分布２
    KSXRSDEV As Double              ' 購ＳＸ比抵抗標準偏差
    KSXRAMIN As Double              ' 購ＳＸ比抵抗平均下限
    KSXRAMAX As Double              ' 購ＳＸ比抵抗平均上限
    KSXFORM As String * 1           ' 購ＳＸ形状
    KSXD1CEN As Double              ' 購ＳＸ直径１中心
    KSXD1MIN As Double              ' 購ＳＸ直径１下限
    KSXD1MAX As Double              ' 購ＳＸ直径１上限
    KSXD1KBN As String * 1          ' 購ＳＸ直径１検査区分
    KSXD2CEN As Double              ' 購ＳＸ直径２中心
    KSXD2MIN As Double              ' 購ＳＸ直径２下限
    KSXD2MAX As Double              ' 購ＳＸ直径２上限
    KSXD2KBN As String * 1          ' 購ＳＸ直径２検査区分
    KSXDUNIT As String * 1          ' 購ＳＸ直径単位
    KSXCDIR As String * 1           ' 購ＳＸ結晶面方位
    KSXCSCEN As Double              ' 購ＳＸ結晶面傾中心
    KSXCSMIN As Double              ' 購ＳＸ結晶面傾下限
    KSXCSMAX As Double              ' 購ＳＸ結晶面傾上限
    KSXCKWAY As String * 2          ' 購ＳＸ結晶面検査方法
    KSXCKHNM As String * 1          ' 購ＳＸ結晶面検査頻度＿枚
    KSXCKHNI As String * 1          ' 購ＳＸ結晶面検査頻度＿位
    KSXCKHNH As String * 1          ' 購ＳＸ結晶面検査頻度＿保
    KSXCKHNS As String * 1          ' 購ＳＸ結晶面検査頻度＿試
    KSXCSDIR As String * 2          ' 購ＳＸ結晶面傾方位
    KSXCSDIS As String * 1          ' 購ＳＸ結晶面傾方位指定
    KSXCTDIR As String * 2          ' 購ＳＸ結晶面傾縦方位
    KSXCTCEN As Double              ' 購ＳＸ結晶面傾縦中心
    KSXCTMIN As Double              ' 購ＳＸ結晶面傾縦下限
    KSXCTMAX As Double              ' 購ＳＸ結晶面傾縦上限
    KSXCYDIR As String * 2          ' 購ＳＸ結晶面傾横方位
    KSXCYCEN As Double              ' 購ＳＸ結晶面傾横中心
    KSXCYMIN As Double              ' 購ＳＸ結晶面傾横下限
    KSXCYMAX As Double              ' 購ＳＸ結晶面傾横上限
    KSXOF1PD As String * 2          ' 購ＳＸＯＦ１位置方位
    KSXOF1PN As Double              ' 購ＳＸＯＦ１位置下限
    KSXOF1PX As Double              ' 購ＳＸＯＦ１位置上限
    KSXOF1PK As String * 1          ' 購ＳＸＯＦ１位置検査区分
    KSXOF1PW As String * 2          ' 購ＳＸＯＦ１位置検査方法
    KSXOF1LC As Double              ' 購ＳＸＯＦ１長中心
    KSXOF1LN As Double              ' 購ＳＸＯＦ１長下限
    KSXOF1LX As Double              ' 購ＳＸＯＦ１長上限
    KSXOF1LK As String * 1          ' 購ＳＸＯＦ１長検査区分
    KSXOF1DC As Double              ' 購ＳＸＯＦ１直径中心
    KSXOF1DN As Double              ' 購ＳＸＯＦ１直径下限
    KSXOF1DX As Double              ' 購ＳＸＯＦ１直径上限
    KSXOF1DK As String * 1          ' 購ＳＸＯＦ１直径検査区分
    KSXDFORM As String * 1          ' 購ＳＸ溝形状
    KSXDFKBN As String * 1          ' 購ＳＸ溝形状検査区分
    KSXDPDRC As String * 1          ' 購ＳＸ溝位置方向
    KSXDPACN As Integer             ' 購ＳＸ溝位置角度中心
    KSXDPAMN As Integer             ' 購ＳＸ溝位置角度下限
    KSXDPAMX As Integer             ' 購ＳＸ溝位置角度上限
    KSXDPKWY As String * 2          ' 購ＳＸ溝位置検査方法
    KSXDPKBN As String * 1          ' 購ＳＸ溝位置検査区分
    KSXDPDIR As String * 2          ' 購ＳＸ溝位置方位
    KSXDPMIN As Double              ' 購ＳＸ溝位置下限
    KSXDPMAX As Double              ' 購ＳＸ溝位置上限
    KSXDWCEN As Double              ' 購ＳＸ溝巾中心
    KSXDWMIN As Double              ' 購ＳＸ溝巾下限
    DSXDWMAX As Double              ' 購ＳＸ溝巾上限
    KSXDWKBN As String * 1          ' 購ＳＸ溝巾検査区分
    KSXDDCEN As Double              ' 購ＳＸ溝深中心
    KSXDDMIN As Double              ' 購ＳＸ溝深下限
    KSXDDMAX As Double              ' 購ＳＸ溝深上限
    KSXDDKBN As String * 1          ' 購ＳＸ溝深検査区分
    KSXDACEN As Double              ' 購ＳＸ溝角度中心
    KSXDAMIN As Double              ' 購ＳＸ溝角度下限
    KSXDAMAX As Double              ' 購ＳＸ溝角度上限
    KSXDAKBN As String * 1          ' 購ＳＸ溝角度検査区分
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 顧客仕様SXLﾃﾞｰﾀ２
Public Type typ_TBCME006
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KMGSRRNO As String * 9          ' 購管理仕様登録依頼番号
    KSXTMMAX As Long                ' 購ＳＸ転位密度上限
    KSXTMSPH As String * 1          ' 購ＳＸ転位密度測定位置＿方
    KSXTMSPT As String * 1          ' 購ＳＸ転位密度測定位置＿点
    KSXTMSPR As String * 1          ' 購ＳＸ転位密度測定位置＿領
    KSXTMKBN As String * 1          ' 購ＳＸ転位密度検査区分
    KSXTMKHM As String * 1          ' 購ＳＸ転位密度検査頻度＿枚
    KSXTMKHI As String * 1          ' 購ＳＸ転位密度検査頻度＿位
    KSXTMKHH As String * 1          ' 購ＳＸ転位密度検査頻度＿保
    KSXTMKHS As String * 1          ' 購ＳＸ転位密度検査頻度＿試
    KSXLTMIN As Integer             ' 購ＳＸＬタイム下限
    KSXLTMAX As Integer             ' 購ＳＸＬタイム上限
    KSXLTUNT As String * 1          ' 購ＳＸＬタイム単位
    KSXLTSPH As String * 1          ' 購ＳＸＬタイム測定位置＿方
    KSXLTSPT As String * 1          ' 購ＳＸＬタイム測定位置＿点
    KSXLTSPI As String * 1          ' 購ＳＸＬタイム測定位置＿位
    KSXLTHWT As String * 1          ' 購ＳＸＬタイム保証方法＿対
    KSXLTHWS As String * 1          ' 購ＳＸＬタイム保証方法＿処
    KSXLTKWY As String * 2          ' 購ＳＸＬタイム検査方法
    KSXLTNSW As String * 2          ' 購ＳＸＬタイム熱処理法
    KSXLTKBN As String * 1          ' 購ＳＸＬタイム検査区分
    KSXLTKHM As String * 1          ' 購ＳＸＬタイム検査頻度＿枚
    KSXLTKHI As String * 1          ' 購ＳＸＬタイム検査頻度＿位
    KSXLTKHH As String * 1          ' 購ＳＸＬタイム検査頻度＿保
    KSXLTKHS As String * 1          ' 購ＳＸＬタイム検査頻度＿試
    KSXLTMBP As Double              ' 購ＳＸＬタイム面内分布
    KSXLTMCL As String * 1          ' 購ＳＸＬタイム面内計算
    KSXCNMIN As Double              ' 購ＳＸ炭素濃度下限
    KSXCNMAX As Double              ' 購ＳＸ炭素濃度上限
    KSXCNIND As String * 2          ' 購ＳＸ炭素濃度指数
    KSXCNUNT As String * 1          ' 購ＳＸ炭素濃度単位
    KSXCNSPH As String * 1          ' 購ＳＸ炭素濃度測定位置＿方
    KSXCNSPT As String * 1          ' 購ＳＸ炭素濃度測定位置＿点
    KSXCNSPI As String * 1          ' 購ＳＸ炭素濃度測定位置＿位
    KSXCNHWT As String * 1          ' 購ＳＸ炭素濃度保証方法＿対
    KSXCNHWS As String * 1          ' 購ＳＸ炭素濃度保証方法＿処
    KSXCNKWY As String * 2          ' 購ＳＸ炭素濃度検査方法
    KSXCNKBN As String * 1          ' 購ＳＸ炭素濃度検査区分
    KSXONMIN As Double              ' 購ＳＸ酸素濃度下限
    KSXONMAX As Double              ' 購ＳＸ酸素濃度上限
    KSXONIND As String * 2          ' 購ＳＸ酸素濃度指数
    KSXONUNT As String * 1          ' 購ＳＸ酸素濃度単位
    KSXONSPH As String * 1          ' 購ＳＸ酸素濃度測定位置＿方
    KSXONSPT As String * 1          ' 購ＳＸ酸素濃度測定位置＿点
    KSXONSPI As String * 1          ' 購ＳＸ酸素濃度測定位置＿位
    KSXONHWT As String * 1          ' 購ＳＸ酸素濃度保証方法＿対
    KSXONHWS As String * 1          ' 購ＳＸ酸素濃度保証方法＿処
    KSXONKWY As String * 2          ' 購ＳＸ酸素濃度検査方法
    KSXONKBN As String * 1          ' 購ＳＸ酸素濃度検査区分
    KSXONKHM As String * 1          ' 購ＳＸ酸素濃度検査頻度＿枚
    KSXONKHI As String * 1          ' 購ＳＸ酸素濃度検査頻度＿位
    KSXONKHH As String * 1          ' 購ＳＸ酸素濃度検査頻度＿保
    KSXONKHS As String * 1          ' 購ＳＸ酸素濃度検査頻度＿試
    KSXONMBP As Double              ' 購ＳＸ酸素濃度面内分布
    KSXONMCL As String * 1          ' 購ＳＸ酸素濃度面内計算
    KSXONLTB As Double              ' 購ＳＸ酸素濃度ＬＴ分布
    KSXONLTC As String * 1          ' 購ＳＸ酸素濃度ＬＴ計算
    KSXONSDV As Double              ' 購ＳＸ酸素濃度標準偏差
    KSXONAMN As Double              ' 購ＳＸ酸素濃度平均下限
    KSXONAMX As Double              ' 購ＳＸ酸素濃度平均上限
    KSXONMNH As Double              ' 購ＳＸ酸素濃度下限補正
    KSXONMXH As Double              ' 購ＳＸ酸素濃度上限補正
    KSXONHCL As String * 2          ' 購ＳＸ酸素濃度補正計算
    KSXGSFIN As String * 1          ' 購ＳＸ外周仕上げ
    KSXCLMIN As Integer             ' 購ＳＸ結晶長下限
    KSXCLMAX As Integer             ' 購ＳＸ結晶長上限
    KSXCLPMN As Integer             ' 購ＳＸ結晶長許容下限
    KSXCLPR As Double               ' 購ＳＸ結晶長許容比率
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type


' 顧客仕様SXLﾃﾞｰﾀ３
Public Type typ_TBCME007
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    FACTORY As String * 1           ' 工場
    OPECOND As String * 1           ' 操業条件
    KMGSRRNO As String * 9          ' 購管理仕様登録依頼番号
    KSXOF1MAX As Double             ' 購ＳＸＯＳＦ上限
    KSXOF1AMX As Double             ' 購ＳＸＯＳＦ平均上限
    KSXOF1SPH As String * 1         ' 購ＳＸＯＳＦ測定位置＿方
    KSXOF1SPT As String * 1         ' 購ＳＸＯＳＦ測定位置＿点
    KSXOF1SPR As String * 1         ' 購ＳＸＯＳＦ測定位置＿領
    KSXOF1HWT As String * 1         ' 購ＳＸＯＳＦ保証方法＿対
    KSXOF1HWS As String * 1         ' 購ＳＸＯＳＦ保証方法＿処
    KSXOF1SZY As String * 1         ' 購ＳＸＯＳＦ測定条件
    KSXOF1KBN As String * 1         ' 購ＳＸＯＳＦ検査区分
    KSXOF1KHM As String * 1         ' 購ＳＸＯＳＦ検査頻度＿枚
    KSXOF1KHI As String * 1         ' 購ＳＸＯＳＦ検査頻度＿位
    KSXOF1KHH As String * 1         ' 購ＳＸＯＳＦ検査頻度＿保
    KSXOF1KHS As String * 1         ' 購ＳＸＯＳＦ検査頻度＿試
    KSXOF1FGS As String * 1         ' 購ＳＸＯＳＦ雰囲気ガス
    KSXOF1CET As Integer            ' 購ＳＸＯＳＦ選択ＥＴ代
    KSXOF1NSW As String * 2         ' 購ＳＸＯＳＦ熱処理法
    KSXOF1SO1 As Integer            ' 購ＳＸＯＳＦ処理温度１
    KSXOF1ST1 As Integer            ' 購ＳＸＯＳＦ処理時間１
    KSXOF2MX As Double              ' 購ＳＸＯＳＦ２上限
    KSXOF2AX As Double              ' 購ＳＸＯＳＦ２平均上限
    KSXOF2SH As String * 1          ' 購ＳＸＯＳＦ２測定位置＿方
    KSXOF2ST As String * 1          ' 購ＳＸＯＳＦ２測定位置＿点
    KSXOF2SR As String * 1          ' 購ＳＸＯＳＦ２測定位置＿領
    KSXOF2HT As String * 1          ' 購ＳＸＯＳＦ２保証方法＿対
    KSXOF2HS As String * 1          ' 購ＳＸＯＳＦ２保証方法＿処
    KSXOF2SZ As String * 1          ' 購ＳＸＯＳＦ２測定条件
    KSXOF2KB As String * 1          ' 購ＳＸＯＳＦ２検査区分
    KSXOF2KM As String * 1          ' 購ＳＸＯＳＦ２検査頻度＿枚
    KSXOF2KI As String * 1          ' 購ＳＸＯＳＦ２検査頻度＿位
    KSXOF2KH As String * 1          ' 購ＳＸＯＳＦ２検査頻度＿保
    KSXOF2KS As String * 1          ' 購ＳＸＯＳＦ２検査頻度＿試
    KSXOF2GS As String * 1          ' 購ＳＸＯＳＦ２雰囲気ガス
    KSXOF2ET As Integer             ' 購ＳＸＯＳＦ２選択ＥＴ代
    KSXOF2NS As String * 2          ' 購ＳＸＯＳＦ２熱処理法
    KSXOF2O1 As Integer             ' 購ＳＸＯＳＦ２処理温度１
    KSXOF2T1 As Integer             ' 購ＳＸＯＳＦ２処理時間１
    KSXBMMAX As Double              ' 購ＳＸＢＭＤ平均下限
    KSXBMMIN As Double              ' 購ＳＸＢＭＤ平均上限
    KSXBMSPH As String * 1          ' 購ＳＸＢＭＤ測定位置＿方
    KSXBMSPT As String * 1          ' 購ＳＸＢＭＤ測定位置＿点
    KSXBMSPR As String * 1          ' 購ＳＸＢＭＤ測定位置＿領
    KSXBMHWT As String * 1          ' 購ＳＸＢＭＤ保証方法＿対
    KSXBMHWS As String * 1          ' 購ＳＸＢＭＤ保証方法＿処
    KSXBMSZY As String * 1          ' 購ＳＸＢＭＤ測定条件
    KSXBMKBN As String * 1          ' 購ＳＸＢＭＤ検査区分
    KSXBMKHM As String * 1          ' 購ＳＸＢＭＤ検査頻度＿枚
    KSXBMKHI As String * 1          ' 購ＳＸＢＭＤ検査頻度＿位
    KSXBMKHH As String * 1          ' 購ＳＸＢＭＤ検査頻度＿保
    KSXBMKHS As String * 1          ' 購ＳＸＢＭＤ検査頻度＿試
    KSXBMFGS As String * 1          ' 購ＳＸＢＭＤ雰囲気ガス
    KSXBMCET As Integer             ' 購ＳＸＢＭＤ選択ＥＴ代
    KSXBMNS As String * 2           ' 購ＳＸＢＭＤ熱処理法
    KSXBM2AN As Double              ' 購ＳＸＢＭＤ２平均下限
    KSXBM2AX As Double              ' 購ＳＸＢＭＤ２平均上限
    KSXBM2SH As String * 1          ' 購ＳＸＢＭＤ２測定位置＿方
    KSXBM2ST As String * 1          ' 購ＳＸＢＭＤ２測定位置＿点
    KSXBM2SR As String * 1          ' 購ＳＸＢＭＤ２測定位置＿領
    KSXBM2HT As String * 1          ' 購ＳＸＢＭＤ２保証方法＿対
    KSXBM2HS As String * 1          ' 購ＳＸＢＭＤ２保証方法＿処
    KSXBM2SZ As String * 1          ' 購ＳＸＢＭＤ２測定条件
    KSXBM2KB As String * 1          ' 購ＳＸＢＭＤ２検査区分
    KSXBM2KM As String * 1          ' 購ＳＸＢＭＤ２検査頻度＿枚
    KSXBM2KI As String * 1          ' 購ＳＸＢＭＤ２検査頻度＿位
    KSXBM2KH As String * 1          ' 購ＳＸＢＭＤ２検査頻度＿保
    KSXBM2KS As String * 1          ' 購ＳＸＢＭＤ２検査頻度＿試
    KSXBM2GS As String * 1          ' 購ＳＸＢＭＤ２雰囲気ガス
    KSXBM2ET As Integer             ' 購ＳＸＢＭＤ２選択ＥＴ代
    KSXBM2NS As String * 2          ' 購ＳＸＢＭＤ２熱処理法
    KSXDENKU As String * 1          ' 購ＳＸＤｅｎ検査有無
    KSXDENMX As Integer             ' 購ＳＸＤｅｎ上限
    KSXDENMN As Integer             ' 購ＳＸＤｅｎ下限
    KSXDENHT As String * 1          ' 購ＳＸＤｅｎ保証方法＿対
    KSXDENHS As String * 1          ' 購ＳＸＤｅｎ保証方法＿処
    KSXLDLKU As String * 1          ' 購ＳＸＬ／ＤＬ検査有無
    KSXLDLMX As Integer             ' 購ＳＸＬ／ＤＬ上限
    KSXLDLMN As Integer             ' 購ＳＸＬ／ＤＬ下限
    KSXLDLHT As String * 1          ' 購ＳＸＬ／ＤＬ保証方法＿対
    KSXLDLHS As String * 1          ' 購ＳＸＬ／ＤＬ保証方法＿処
    KSXDVDKU As String * 1          ' 購ＳＸＤＶＤ２検査有無
    KSXDVDMX As Integer             ' 購ＳＸＤＶＤ２上限
    KSXDVDMN As Integer             ' 購ＳＸＤＶＤ２下限
    KSXDVDHT As String * 1          ' 購ＳＸＤＶＤ２保証方法＿対
    KSXDVDHS As String * 1          ' 購ＳＸＤＶＤ２保証方法＿処
    KSXGDSPH As String * 1          ' 購ＳＸＧＤ測定位置＿方
    KSXGDSPT As String * 1          ' 購ＳＸＧＤ測定位置＿点
    KSXGDSPR As String * 1          ' 購ＳＸＧＤ測定位置＿領
    KSXGDSZY As String * 1          ' 購ＳＸＧＤ測定条件
    KSXGDZAR As Integer             ' 購ＳＸＧＤ除外領域
    KSXGDKHM As String * 1          ' 購ＳＸＧＤ検査頻度＿枚
    KSXGDKHI As String * 1          ' 購ＳＸＧＤ検査頻度＿位
    KSXGDKHH As String * 1          ' 購ＳＸＧＤ検査頻度＿保
    KSXGDKHS As String * 1          ' 購ＳＸＧＤ検査頻度＿試
    KSXDSOKE As String * 1          ' 購ＳＸＤＳＯＤ検査
    KSXDSOMX As Long                ' 購ＳＸＤＳＯＤ上限
    KSXDSOMN As Long                ' 購ＳＸＤＳＯＤ下限
    KSXDSOAX As Integer             ' 購ＳＸＤＳＯＤ領域上限
    KSXDSOAN As Integer             ' 購ＳＸＤＳＯＤ領域下限
    KSXDSOHT As String * 1          ' 購ＳＸＤＳＯＤ保証方法＿対
    KSXDSOHS As String * 1          ' 購ＳＸＤＳＯＤ保証方法＿処
    KSXDSOKM As String * 1          ' 購ＳＸＤＳＯＤ検査頻度＿枚
    KSXDSOKI As String * 1          ' 購ＳＸＤＳＯＤ検査頻度＿位
    KSXDSOKH As String * 1          ' 購ＳＸＤＳＯＤ検査頻度＿保
    KSXDSOKS As String * 1          ' 購ＳＸＤＳＯＤ検査頻度＿試
    KSXCDOP As String * 1           ' 購ＳＸ結晶ドープ
    IFKBN As String * 4             ' Ｉ／Ｆ区分
    SYORIKBN As String * 1          ' 処理区分
    SPECRRNO As String * 9          ' 仕様登録依頼番号
    SXLMCNO As String * 12          ' ＳＸＬ製作条件番号
    WFMCNO As String * 12           ' ＷＦ製作条件番号
    StaffID As String * 8           ' 社員ID
    REGDATE As Date                 ' 登録日付
    UPDDATE As Date                 ' 更新日付
    SENDFLAG As String * 1          ' 送信フラグ
    SENDDATE As Date                ' 送信日付
End Type

' 欠落ブロック一覧
Public Type typ_LackBlk
    BLOCKID As String * 12      ' ブロックID
    INGOTPOS As Integer         ' 結晶内開始位置
    REALLEN As Integer          ' 長さ
    REJDTTM As Date             ' 欠落日
    NYUKO As Integer            ' そのブロックが入庫済か(1:入庫済 0:未入庫)
    sBlockId As String * 12     ' 払出し単位の先頭ブロックID
    MINYUKO As Integer          ' 払出し単位での未入庫ブロック数
    PUPTN As String             ' 引上ﾊﾟﾀｰﾝ     2004/12/08 追加
    HOLDFLG As String * 1       ' ﾎｰﾙﾄﾞ区分　05/01/31 ooba
    WFHOLDFLG As String * 1     ' WFﾎｰﾙﾄﾞ区分　05/01/31 ooba
    WFHUFLG As String * 1       ' WF振替FLG　06/02/06 ooba
    MUKESAKI As String          ' 向先 07/09/03 SPK Tsutsumi Add
    Koutei As String * 5        ' 工程(XSDCB)　08/01/31 ooba
    KANREN As String * 1        ' 関連ﾌﾞﾛｯｸ有無　08/01/31 ooba
    AGRSTATUS As String             ' 承認確認区分 add SETkimizuka
    STOP    As String               ' 停止 add SETkimizuka
    CAUSE   As String               ' 停止理由 add SETkimizuka
    PRINTNO As String               ' 先行評価 add SETkimizuka
    'Add Start 2010/07/08 SMPK Nakamura
    HINCNT  As String           'ブロック内品番数
    hinban  As String           'ブロック内品番
    CW740STS As String          'CW740ステータス
    'Add End 2010/07/08 SMPK Nakamura
End Type

'add start 2003/04/18 hitec)後藤　--------
' 欠落ブロック一覧(表示用）
Public Type tbl_DispLack
'Chg Start 2010/07/08 SMPK Nakamura
'    BLOCKID As String * 12      ' ブロックID
    SELECTED As Long            ' 選択項目
    BLOCKID As String           ' ブロックID
'Chg End 2010/07/08 SMPK Nakamura
    INGOTPOS As Integer         ' 結晶内開始位置
    REALLEN As Integer          ' 長さ
    REJDTTM As String             ' 欠落日
    PUPTN As String             ' 引上ﾊﾟﾀｰﾝ     2004/12/08 追加
    HOLDFLG As String * 1       ' ﾎｰﾙﾄﾞ区分　05/01/31 ooba
    WFHOLDFLG As String * 1     ' WFﾎｰﾙﾄﾞ区分　05/01/31 ooba
    WFHUFLG As String * 1       ' WF振替FLG　06/02/06 ooba
    MUKESAKI As String          ' 向先 07/09/03 SPK Tsutsumi Add
    Koutei As String * 5        ' 工程(XSDCB)　08/01/31 ooba
    KANREN As String * 1        ' 関連ﾌﾞﾛｯｸ有無　08/01/31 ooba
    AGRSTATUS As String             ' 承認確認区分 add SETkimizuka
    STOP    As String               ' 停止 add SETkimizuka
    CAUSE   As String               ' 停止理由 add SETkimizuka
    PRINTNO As String               ' 先行評価 add SETkimizuka
    'Add Start 2010/07/08 SMPK Nakamura
    hinban  As String           'ブロック内品番
    WAITBLOCK As String         '関連待ちブロック有無
    CW740STS As String          'CW740ステータス
    'Add End 2010/07/08 SMPK Nakamura
End Type
Public tblDispLack() As tbl_DispLack
'add end 2003/04/18 hitec)後藤　--------

'add 2005/11/11 高崎->
'10Ω換算値取得構造体
Public Type typ_OumConvSet
    CTR01A9 As String          '傾き
    CTR02A9 As String          '切片
    CTR03A9 As String          '設定値
End Type
'add 2005/11/11 高崎<-

'add 2009/07/22 SETsw Nakada -->
Public Type typ_TBCMJ020
    CRYNUM     As String * 12       ' 結晶番号
    POSITION   As Integer           ' 位置
    SMPKBN     As String * 1        ' サンプル区分
    TRANCNT    As Integer           ' 処理回数
    TRANCOND   As String * 1        ' 処理条件
    BLOCKID    As String * 12       ' ブロックID
    SMPLNO     As Long              ' サンプルＮｏ
    SMPLUMU    As String * 1        ' サンプル有無
    KRPROCCD   As String * 5        ' 管理工程コード
    PROCCODE   As String * 5        ' 工程コード
    N2NOUDO    As Double            ' Ｎ２濃度仮数
    N2NI       As Integer           ' Ｎ２濃度指数
    TSTAFFID   As String * 8        ' 登録社員ID
    REGDATE    As Date              ' 登録日付
    KSTAFFID   As String * 8        ' 更新社員ID
    UPDDATE    As Date              ' 更新日付
    SENDFLAG   As String * 1        ' 送信フラグ
    SENDDATE   As Date              ' 送信日付
End Type
'add 2009/07/22 SETsw Nakada <--


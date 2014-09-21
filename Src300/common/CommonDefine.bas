Attribute VB_Name = "CommonDefine"
Option Explicit

''
'' システム共通の固有Define変数定義
''f

''　関数戻り値定義
Public Enum FUNCTION_RETURN         ''関数の戻り値
    FUNCTION_RETURN_SUCCESS = 0     '' 正常
    FUNCTION_RETURN_FAILURE = -1    '' 異常
End Enum

'' コントロール状態設定種類
Public Enum enm_CtrlStateKind
    CTRL_ENABLE             '' コントロール有効
    CTRL_ENABLE_GRAY        '' コントロール有効（表示項目色表示・グレー）
    CTRL_ENABLE_YELLOW      '' コントロール有効（表示項目色表示・イエロー）
    CTRL_ENABLE_RED         '' コントロール有効（表示項目色表示・レッド）
    CTRL_DISABLE            '' コントロール無効（表示項目色表示）
    CTRL_DISABLE_GRAY       '' コントロール無効（表示項目色表示・グレー）
    CTRL_WARNING            '' コントロール有効（警告色表示）
    CTRL_DISABLE_WARNING    '' コントロール有効（警告色表示）
    CTRL_DISABLE_YELLOW     '' コントロール有効（表示項目色表示・イエロー）
    CTRL_SELECTED           '' コントロール無効（選択色）
    CTRL_DISABLE_RED        '' コントロール無効（表示項目色表示・レッド）
    CTRL_DISABLE_SKY        '' コントロール無効（表示項目色表示・空色）
    'Add Start 2010/08/04 SMPK Nakamura
    CTRL_ENABLE_SKY         '' コントロール有効（表示項目色表示・空色）
    'Add End 2010/08/04 SMPK Nakamura
End Enum

'' 色定義
Public Enum FIELD_COLOR
    COLOR_OK = vbWindowBackground       '' 正常値の背景色
    COLOR_ENABLE = vbWindowBackground   '' 有効(通常)の背景色
    COLOR_NG = &HC0C0FF                 '' 異常値の背景色
    COLOR_WARNING = &HC0C0FF            '' 警告の背景色
    COLOR_DISABLE = &H80FF80            '' 入力不可のフィールドの背景色(一部で利用)
    COLOR_GRAY = vbButtonFace   '' 入力不可のフィールドの背景色(一部で利用)
    COLOR_SAMPLE = vbCyan               '' サンプルが非デフォルト状態の背景色(一部で利用)
    COLOR_SELECTED = &HFFC0C0           '' 選択色
    COLOR_YELLOW = vbYellow
    COLOR_RED = &HC0C0FF
    COLOR_SKY = &HFFFF80    '' 推定値
End Enum


'' DB内の各種区分コード
Public Const USECLS_PLAN = "0"          ''*設計 使用区分 設計時
Public Const USECLS_ONUSE = "1"         ''*設計 使用区分 加工払出済
Public Const DELCLS_NORMAL = "0"        ''削除区分 通常
Public Const DELCLS_DELETE = "1"        ''削除区分 削除
Public Const LSTATCLS_NORMAL = "0"      ''最終状態区分 通常
Public Const LSTATCLS_HAIKI = "H"       ''最終状態区分 廃棄
Public Const LSTATCLS_SXL = "S"         ''最終状態区分 シングル確定済
Public Const HOLDCLS_NORMAL = "0"       ''ホールド区分 ホールド解除
Public Const HOLDCLS_HOLD = "1"         ''ホールド区分 ホールド
Public Const KTKBN_NORMAL = "0"         ''確定区分 通常
Public Const KTKBN_KAKUTEI = "1"        ''確定区分 確定済


'' 品番型
Public Type tFullHinban
    hinban As String * 8            ' 品番
    mnorevno As Integer             ' 製品番号改訂番号
    factory As String * 1           ' 工場
    opecond As String * 1           ' 操業条件
'Ｚ品番対応区分追加　　　Ｚ品番及び　品番の種類を登録及び処理実施
    Hinkubun As String * 1
    sMukesaki As String             ' 向先 2007/08/16 SPK Tsutsumi Add
End Type

'' 構成品番情報
Public Type KosHin
    hinban As String
    HinPosTop As Integer
    HinPosTail As Integer
    HinLen As Integer
End Type


'' 連番種別管理コード(4桁)
Public Const SEQ_HIKIAGE_SIJI = "SIJI"                  '引上指示No
Public Const SEQ_RMLT_GENRYO = "RMLT"                   'リメルト原料番号
Public Const SEQ_CRYNUM = "CRYN"                        '結晶番号
Public Const SEQ_SAMPLENO = "CSMP"                      'サンプルNo


' 工程コード    ( 工程コードがついているもの以外は、権限管理用のID  ＊印のもの）
Public Const PROCD_SIYOU_UKEIRE = "CC002"               ' 仕様受け入れ処理                  *
Public Const PROCD_SIYOU_NYUURYOKU = "CC003"            ' 製品仕様入力処理                  *
Public Const PROCD_TAKESSYOU_UKEIRE = "CB100"           ' 多結晶受入搬入処理
Public Const PROCD_RIMERUTO_UKEIRE = "CB210"            ' リメルト受入れ切断処理
Public Const PROCD_RIMERUTO_HARAIDASI = "CB220"         ' リメルト洗浄払出処理
Public Const PROCD_GENRYOU_ZAIKO_SYUUSEI = "CB240"      ' 原料在庫修正処理
Public Const PROCD_KAKUAGE = "CB320"                    ' クリスタルカタログ検索格上げ処理
Public Const PROCD_HIKIAGE_SIJI = "CC100"               ' 引上げ指示処理処理
Public Const PROCD_HIKIAGE_TOUNYUU = "CC200"            ' 引上げ投入実績処理
Public Const PROCD_HIKIAGE_SYUURYOU = "CC300"           ' 引上げ終了実績処理
Public Const PROCD_KAKOU_HARAIDASI = "CC310"            ' 加工払出処理
Public Const PROCD_KENNSAKU_KAKOU = "CC400"             ' 研削加工実績処理
Public Const PROCD_SETUDAN = "CC450"                    ' 切断実績処理
''↓追加 START SPT用実績作成方法変更 2006/05/12 SMP-OKAMOTO
Public Const PROCD_SAISETUDAN = "CC460"                 ' 再切断実績処理
''↑追加 END   SPT用実績作成方法変更 2006/05/12 SMP-OKAMOTO
Public Const PROCD_KESSYOU_HOLD = "CC004"               ' 結晶ホールド処理                  *
Public Const PROCD_KESSYOU_HOLD_KAIJO = "CC005"         ' 結晶ホールド解除処理              *


Public Const PROCD_EPD = "CC500"                        ' EPD検査実績処理
Public Const PROCD_TEIKOU = "CC501"                     ' 抵抗検査実績処理
Public Const PROCD_FTIR = "CC502"                       ' FTIR検査実績処理
Public Const PROCD_GFA = "CC503"                        ' GFA検査実績処理
Public Const PROCD_GD = "CC504"                         ' GD検査実績処理
Public Const PROCD_OSF = "CC505"                        ' OSF検査実績処理
Public Const PROCD_BMD = "CC506"                        ' BMD検査実績処理
Public Const PROCD_LIFETIME = "CC507"                   ' ライフタイム検査実績処理
Public Const PROCD_X = "CC508"                          ' X線測定入力実績処理   2009/08 Sumco Akizuki
Public Const PROCD_CUDECO = "CC509"                     ' Cu-deco検査実績処理   Add 2010/12/23 SMPK Miyata


Public Const PROCD_KOUNYU_TAN_KESSYOU = "CB410"         ' 購入単結晶受入処理
Public Const PROCD_KESSYOU_SOUGOUHANTEI = "CC600"       ' 結晶総合判定処理
Public Const PROCD_KESSYOU_SAISYUU_HARAIDASI = "CC700"  ' 結晶最終払出入力処理
Public Const PROCD_NUKISI_SIJI = "CC710"                ' 抜試指示入力処理
Public Const PROCD_BLOCK = "CC711"                      ' ブロックラベル発行処理  '4/16 Yam
Public Const PROCD_WFC_HARAIDASI = "CC720"              ' WFセンター払出処理
Public Const PROCD_KESSYOU_SIYOUJOUHOU_HENKOU = "CC730" ' 結晶仕様情報変更処理
Public Const PROCD_NUKISI_HENKOU = "CW740"              ' 抜試変更指示処理
Public Const PROCD_WFC_SOUGOUHANTEI = "CW750"           ' WFセンター総合判定処理
Public Const PROCD_SXL_KAKUTEI = "CW800"                ' シングル確定処理
Public Const PROCD_TEIKO_HENSEKI_KEISSAN = "     "      ' 抵抗偏析計算処理
Public Const PROCD_PGID_MNT = "CC110"                   ' PG-IDメンテナンス
Public Const PROCD_SEISAKUJOUKEN_MNT = "CC001"          ' 製作条件メンテナンス              *
Public Const PROCD_GFA_KOUSEIJOHO = "CC006"              ' GFA校正情報                       *
'vvvv SUMCO様にて追加 vvvv
Public Const PROCD_SYAIN_MNT = "CC007"                  ' 社員マスタメンテナンス  '02/3/30 Yam
Public Const PROCD_KENGEN_MNT = "CC008"                 ' 権限マスタメンテナンス  '02/3/30 Yam
Public Const PROCD_CODE_MNT = "CC009"                   ' コードマスタメンテナンス'02/3/30 Yam
Public Const PROCD_PRINTER_MNT = "CC010"                ' ラベルプリンタマスタメンテナンス'02/3/30 Yam
'^^^^ SUMCO様にて追加 ^^^^
' 特採用を追加  2003/09/12 SystemBrain ===================> START
Public Const PROCD_TOKUSAI_KENGEN = "CC011"                ' 特採権限
' 特採用を追加  2003/09/12 SystemBrain ===================> END
''Add Start 2011/04/06 SMPK Nakamura
Public Const PROCD_FRS_KOUSEIJOHO = "CC006"              ' FRS校正情報                       *
''Add End 2011/04/06 SMPK Nakamura

' 管理工程コード
Public Const MGPRCD_SIYOU_UKEIRE = "     "              ' 仕様受け入れ処理
Public Const MGPRCD_SIYOU_NYUURYOKU = "     "           ' 製品仕様入力処理
Public Const MGPRCD_TAKESSYOU_UKEIRE = "     "          ' 多結晶受入搬入処理
Public Const MGPRCD_RIMERUTO_UKEIRE = "     "           ' リメルト受入れ切断処理
Public Const MGPRCD_RIMERUTO_HARAIDASI = "     "        ' リメルト洗浄払出処理
Public Const MGPRCD_GENRYOU_ZAIKO_SYUUSEI = "     "     ' 原料在庫修正処理
Public Const MGPRCD_KAKUAGE = "     "                   ' クリスタルカタログ検索格上げ処理
Public Const MGPRCD_HIKIAGE_SIJI = "     "              ' 引上げ指示処理処理
Public Const MGPRCD_HIKIAGE_TOUNYUU = "     "           ' 引上げ投入実績処理
Public Const MGPRCD_HIKIAGE_SYUURYOU = "     "          ' 引上げ終了実績処理
Public Const MGPRCD_KAKOU_HARAIDASI = "     "           ' 加工払出処理
Public Const MGPRCD_KENNSAKU_KAKOU = "     "            ' 研削加工実績処理
Public Const MGPRCD_SETUDAN = "     "                   ' 切断実績処理
''↓追加 START SPT用実績作成方法変更 2006/05/12 SMP-OKAMOTO
Public Const MGPRCD_SAISETUDAN = "     "                ' 再切断実績処理
''↑追加 END   SPT用実績作成方法変更 2006/05/12 SMP-OKAMOTO
Public Const MGPRCD_KESSYOU_HOLD = "     "              ' 結晶ホールド処理
Public Const MGPRCD_KESSYOU_HOLD_KAIJO = "     "        ' 結晶ホールド解除処理
Public Const MGPRCD_FTIR = "     "                      ' FTIR検査実績処理
Public Const MGPRCD_GFA = "     "                       ' GFA検査実績処理
Public Const MGPRCD_TEIKOU = "     "                    ' 抵抗検査実績処理
Public Const MGPRCD_BMD = "     "                       ' BMD検査実績処理
Public Const MGPRCD_OSF = "     "                       ' OSF検査実績処理
Public Const MGPRCD_GD = "     "                        ' GD検査実績処理
Public Const MGPRCD_LIFETIME = "     "                  ' ライフタイム検査実績処理
Public Const MGPRCD_EPD = "     "                       ' EPD検査実績処理
Public Const MGPRCD_X = "     "                         ' X線測定入力実績   2009/08 SUMCO Akizuki
Public Const MGPRCD_CU_DECO = "     "                   ' Cu-deco検査実績処理   'Add 2010/12/23 SMPK Miyata
Public Const MGPRCD_KOUNYU_TAN_KESSYOU = "     "        ' 購入単結晶受入処理
Public Const MGPRCD_KESSYOU_SOUGOUHANTEI = "     "      ' 結晶総合判定処理
Public Const MGPRCD_KESSYOU_SAISYUU_HARAIDASI = "     " '結晶最終払出入力処理
Public Const MGPRCD_NUKISI_SIJI = "     "               ' 抜試指示入力処理
Public Const MGPRCD_WFC_HARAIDASI = "     "             ' WFセンター払出処理
Public Const MGPRCD_KESSYOU_SIYOUJOUHOU_HENKOU = "     " ' 結晶仕様情報変更処理
Public Const MGPRCD_NUKISI_HENKOU = "     "             '抜試変更指示処理
Public Const MGPRCD_WFC_SOUGOUHANTEI = "     "          ' WFセンター総合判定処理
Public Const MGPRCD_SXL_KAKUTEI = "     "               ' シングル確定処理
Public Const MGPRCD_TEIKO_HENSEKI_KEISSAN = "     "     ' 抵抗偏析計算処理
Public Const MGPRCD_PGID_MNT = "     "                  ' PG-IDメンテナンス
Public Const MGPRCD_SEISAKUJOUKEN_MNT = "     "         ' 製作条件メンテナンス
Public Const MGPRCD_GFA_KOUSEIJOHO = "     "              ' GFA校正情報
'vvvv SUMCO様にて追加 vvvv
Public Const MGPRCD_SYAIN_MNT = "     "                 ' 社員マスタメンテナンス  '02/3/30 Yam
Public Const MGPRCD_KENGEN_MNT = "     "                ' 権限マスタメンテナンス  '02/3/30 Yam
Public Const MGPRCD_CODE_MNT = "     "                  ' コードマスタメンテナンス'02/3/30 Yam
Public Const MGPRCD_PRINTER_MNT = "     "               ' ラベルプリンタマスタメンテナンス'02/3/30 Yam
'^^^^ SUMCO様にて追加 ^^^^
Public Const MGPRCD_SETUDANSIJI_OYA = "65   "           ' 切断指示(親情報)　05/09/16 ooba
Public Const MGPRCD_SETUDANSIJI_HURYO = "65   "         ' 切断指示(不良情報)　05/09/16 ooba
Public Const MGPRCD_SETUDANSIJI_HURIKAE = "40   "       ' 切断指示(振替情報)　05/09/16 ooba


'' メインメニュープログラム名
Public Const MAINMENU_PGNAME = "cmac001.exe"


Public Const HIJU_SILICONE = 0.00233               ''シリコンの比重(g/mm3)
Public Const cdblPI As Double = 3.1416             ''π

Public Type type_Coefficient   ''偏析係数計算構造体
    DUNMENSEKI      As Double   ''断面積
    TOPSMPLPOS      As Double   ''トップサンプル位置
    BOTSMPLPOS      As Double   ''ボトムサンプル位置
    CHARGEWEIGHT    As Double   ''チャージ量
    TOPWEIGHT       As Double   ''トップ重量
    TOPRES          As Double   ''トップ側比抵抗中央値
    BOTRES          As Double   ''ボトム側比抵抗中央値
End Type

'=================================================================================
' 2011/01/17 tkimura ADD START
Public Type type_Coefficient_new2   ''推定抵抗,推定引上率計算構造体
    DUNMENSEKI      As Double   ''断面積
    TOPSMPLPOS      As Double   ''トップサンプル位置
    BOTSMPLPOS      As Double   ''ボトムサンプル位置
    SMPLPOS         As Double   ''比抵抗位置(推定位置)
    CHARGEWEIGHT    As Double   ''チャージ量
    CHARGEWEIGHTA   As Double   ''A結晶のチャージ量                         2011/01/17 tkimura ADD
    TOPWEIGHT       As Double   ''トップ重量
    TOPRES          As Double   ''トップ側比抵抗中央値
    BOTRES          As Double   ''ボトム側比抵抗中央値
    HIKIFLG         As Integer  ''引上げフラグ(1=通常orマルチ、2=B,C結晶)   2011/01/17 tkimura ADD
    Henseki         As Double   ''実効偏析                                  2011/01/17 tkimura ADD
    KIJUNTEIKOU     As Double   ''基準抵抗値                                2011/01/17 tkimura ADD
    SUITEIHIKIRITU  As Double   ''推定対象引上げ率                          2011/01/17 tkimura ADD
    SUITEITEIKOU    As Double   ''推定位置比抵抗率                          2011/01/17 tkimura ADD
    GT              As Double   ''ρTop位置引上げ率                         2011/01/17 tkimura ADD
    GB              As Double   ''ρBot位置引上げ率                         2011/01/17 tkimura ADD
    HOSEICHO        As Double   ''補正結晶長                                2011/04/25 Marushita ADD　Micronｲﾝｺﾞｯﾄ位置管理追加対応
End Type


' 2011/01/17 tkimura ADD END
'=================================================================================

Public Type type_ResPosCal     ''推定計算構造体
    COEFFICIENT     As Double   ''偏析係数
    DUNMENSEKI      As Double   ''断面積
    CHARGEWEIGHT    As Double   ''チャージ量
    TOPWEIGHT       As Double   ''トップ重量
    TOPSMPLPOS      As Double   ''トップサンプル位置
    TOPRES          As Double   ''トップ側比抵抗中央値
    target          As Double   ''計算対象抵抗値 or 位置
End Type

Public Type type_CbonCal       ''炭素濃度推定計算構造体
    CHARGEWEIGHT    As Double   ''チャージ量
    TOPWEIGHT       As Double   ''トップ重量
    DUNMENSEKI      As Double   ''断面積
    SMPLPOS         As Double   ''サンプル位置
    CSDATA          As Double   ''炭素濃度
    target          As Double   ''推定位置
End Type

'' メッセージコード定義
Public Const MSG_NOTFOUND_CRYNUM = "ECRY0"      '' 該当する結晶番号が見つかりません。
Public Const MSG_GETERROR_DBDATA = "EGET2"       '' DBデータ取得エラー
Public Const MSG_DISPLAY_ERROR = "EDISP"        '' 表示エラー
Public Const MSG_NOTINPUT_ERROR = "EINIM"       '' 未入力エラー
Public Const MSG_CALC_ERROR = "ECALC"           '' 計算エラー
Public Const MSG_NOTFOUND_HINBAN_ERR = "EHIN0"  '' 品番エラー
Public Const MSG_HINBAN_ERROR = "EHIN1"         '' 品番桁数エラー


'' グローバル変数
Public g_PrSpSXLData1 As typ_TBCME018   '' 製品仕様ＳＸＬデータ１
Public g_PlupEndRslt As typ_TBCMH004    '' 引上げ終了実績
Public g_tblRs()   As typ_TBCMJ002      '' 比抵抗実績
Public XSDC3_StaffID   As String      '' 担当者コード

Public Const REST_WT_CRYCODE = "ABCDEFGHIJKLNMOPQRSTUVWXYZ"  '' 残量引き結晶コード

Public nowCd As String                  '現在工程  2002/11/24 工程コードロジック統一
Public nextCd As String                 '次工程  2002/11/24 工程コードロジック統一
Public Tokusai As String

''2011/01/20 tkimura ADD START ==========================================================>
'WFﾏｯﾌﾟ管理ﾃｰﾌﾞﾙ構造体にインゴット引上率と枚葉推定抵抗値(Center)を追加した。
'add start 2003/03/15 hitec)matsumoto ---------------

'WFﾏｯﾌﾟ管理ﾃｰﾌﾞﾙ構造体
Public Type typeWFmap
    LOTID       As String       'ブロックID
    BLOCKSEQ    As Integer      'ブロック内連番
    INDTM       As Variant      'ＷＦセンター入庫日時
    BASKETID    As String       'バスケットID
    SLOTNO      As Integer      'スロットNO
    CURRWPCS    As Integer      'ＷＦ枚数
    EXISTFLG    As String       '存在フラグ
    TOP_POS     As Integer      'ブロックのＴＯＰからの位置
    REJCAT      As String       '欠落理由
    TXID        As String       'トランザクションID
    REGDATE     As Variant      '登録日付
    SUMMITSENDFLG   As String   'SUMIT送信フラグ
    SENDFLG     As String       '送信フラグ
    SENDDATE    As Variant      '送信日時
    WFSTA       As String       'WF状態
    HREJCODE    As String       '不良理由コード
    UPDPROC     As String       '更新工程
    UPDDATE     As Variant      '更新日時
    SXLID       As String       'SXLID
    hinban      As String       '品番
    REVNUM      As Integer      '製品番号改訂番号
    factory     As String       '工場
    opecond     As String       '操業条件
    KANKBN      As String       '完了区分
    SMPLEID     As String       '抜試位置
    NREJCODE    As String       '抜試返答理由コード
    SMPLEFLG    As String       'サンプルフラグ
    RTOP_POS    As Double       '論理ブロック内位置
    RITOP_POS   As Double       '論理結晶内位置
    up_Ratio    As String       'インゴット引上率           '2011/01/20 tkimura add
    rs_Meas     As String       '枚葉推定抵抗値(Center)     '2011/01/20 tkimura add
    HTOP_POS    As Double       '補正結晶長                 '2011/04/27 Marushita ADD
End Type

Public gtWFmap() As typeWFmap
'add end 2003/03/15 hitec)matsumoto ---------------------
''2011/01/20 tkimura ADD END ==========================================================>

'add start 2003/04/30 hitec)matsumoto ------------------------
'WFﾏｯﾌﾟ対応画面データ格納構造体
Public Type typeSprWFmap
    LOTID       As Variant      'ブロックID
    hinban      As Variant      '品番
    REVNUM      As Variant      '製品番号改訂番号
    factory     As Variant      '工場
    opecond     As Variant      '操業条件
    HINUP       As tFullHinban        ' 上品番
    HINDN       As tFullHinban        ' 下品番
    blockp      As Variant      'ブロックP
''''    BLOCKP_T    As Variant      'ブロックP（上）
''''    BLOCKP_B    As Variant      'ブロックP（下）
    KESSYOUP    As Variant      '結晶P
''''    KESSYOUP_T  As Variant      '結晶P（上）
''''    KESSYOUP_B  As Variant      '結晶P（下）
    BLOCKSEQ    As Integer      'マップ位置
''''    BLOCKSEQ_T  As Integer      'マップ位置（上）
''''    BLOCKSEQ_B  As Integer      'マップ位置（下）
    wfnum       As Integer      'ＷＦ枚数
    WFSTA       As Variant      'WF状態
''''    WFSTA_T     As Variant      'WF状態（上）
''''    WFSTA_B     As Variant      'WF状態（下）
    REJCODE     As Integer      '不良区分
    SAMPLEID    As Variant      'サンプルID
''''    SAMPLEID_T  As Variant      'サンプルID（上）
''''    SAMPLEID_B  As Variant      'サンプルID（下）
    WFSMP_Rs    As Variant      '検査項目（Rs）
    WFSMP_Oi    As Variant      '検査項目（Oi）
    WFSMP_B1    As Variant      '検査項目（B1）
    WFSMP_B2    As Variant      '検査項目（B2）
    WFSMP_B3    As Variant      '検査項目（B3）
    WFSMP_L1    As Variant      '検査項目（L1）
    WFSMP_L2    As Variant      '検査項目（L2）
    WFSMP_L3    As Variant      '検査項目（L3）
    WFSMP_L4    As Variant      '検査項目（L4）
    WFSMP_DS    As Variant      '検査項目（DS）
    WFSMP_DZ    As Variant      '検査項目（DZ）
    WFSMP_SP    As Variant      '検査項目（SP）
    WFSMP_D1    As Variant      '検査項目（D1）
    WFSMP_D2    As Variant      '検査項目（D2）
    WFSMP_D3    As Variant      '検査項目（D3）
    SHAFLAG     As Variant      'サンプルフラグ
''''    SHAFLAG_T   As Variant      'サンプルフラグ（上）
''''    SHAFLAG_B   As Variant      'サンプルフラグ（下）
    ADD_FLG     As String       '0：既存抜試行，1：追加抜試行
End Type
Public gtSprWfMap() As typeSprWFmap

'WF状態
Public Const gsWF_STA_0 As String = "0"       '通常
Public Const gsWF_STA_1 As String = "1"       '共有
'Public Const gsWF_STA_2 As String = "2"      '指示待ち
'Public Const gsWF_STA_3 As String = "3"      '指示OK
Public Const gsWF_STA_4 As String = "4"       '欠落
'Public Const gsWF_STA_5 As String = "5"      '結果

'サンプルフラグ
Public Const gsWF_SMPL_0 As String = "0"      '欠落
Public Const gsWF_SMPL_1 As String = "1"      '指示待ち
Public Const gsWF_SMPL_2 As String = "2"      '指示OK
Public Const gsWF_SMPL_3 As String = "3"      '指示NG
Public Const gsWF_SMPL_4 As String = "4"      '結果

'WF状態（画面表示）
Public Const gsWF_STA_NORMAL       As String = "通常"      '通常
Public Const gsWF_STA_STA_K       As String = "欠落"      '欠落
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'抜試画面のレイアウト調整にて、文言を変更
'Public Const gsWF_STA_SIJI        As String = "指示待ち"  '指示待ち
'Public Const gsWF_STA_SIJI_OK     As String = "指示OK"    '指示OK
'Public Const gsWF_STA_SIJI_NG     As String = "指示NG"    '指示NG
Public Const gsWF_STA_SIJI        As String = "待ち"      '指示待ち
Public Const gsWF_STA_SIJI_OK     As String = "OK"    '指示OK
Public Const gsWF_STA_SIJI_NG     As String = "NG"    '指示NG
'↑変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
Public Const gsWF_STA_SIJI_KEKKA  As String = "結果"     '結果

Public Const COLOR_CryJitsu = &HC0C0C0          '結晶実績ｾﾙ色　05/01/31 ooba
'サンプルフラグ（画面表示）
Public Const gsWF_SMPL_JOINT    As String = "共有"  '共有
'add end   2003/04/30 hitec)matsumoto ------------------------

Public Const BLOCKLEN_MAX As Integer = 400          'ﾌﾞﾛｯｸ制限長さ　06/06/02 ooba
Public Const GS_BLOCKLEN_MAX As Integer = 420       'ﾌﾞﾛｯｸ制限長さ(ｶﾞﾗｽ接着品)　06/06/02 ooba

Public Const SAMPLENO_HEAD As String = "1"      'ｻﾝﾌﾟﾙ�ｂﾌ頭1桁目 ｻﾝﾌﾟﾙ��6桁対応 2007/05/25 SETsw kubota

'--------------- 2008/08/25 INSERT START  By Systech ---------------
Public Const DKTMP_TBCMB005SYS  As String = "SB"        ' TBCMB005 DK温度判定のシステム区分
Public Const DKTMP_TBCMB005CLS  As String = "DK"        ' TBCMB005 DK温度判定の区分
Public Const DKTMP_TBCME033CODE As String = "0036-04"   ' TBCME033 DK温度名称のコード��
Public Const DKTMP_650_20OV     As String = "1"         ' DK温度 650℃ 20Ω以上
Public Const DKTMP_650_20LO     As String = "2"         ' DK温度 650℃ 20Ω未満
Public Const DKTMP_1100         As String = "3"         ' DK温度 1100℃
'--------------- 2008/08/25 INSERT  END   By Systech ---------------


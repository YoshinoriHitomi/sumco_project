Attribute VB_Name = "s_cmbc032_SQL"
Option Explicit

'結晶最終払出入力


Type cmkc001b_LockWait
    flag                As Boolean
    Grp                 As Integer
End Type

Type cmkc001b_Wait3_HINBAN
    hinban              As String * 8                           ' 品番
    REVNUM              As Integer                              ' 製品番号改訂番号
    factory             As String * 1                           ' 工場
    opecond             As String * 1                           ' 操業条件
End Type

Type cmkc001b_Wait3_BLK
    BLOCKID             As String * 12                          ' ブロックID
    INGOTPOS            As Integer                              ' 結晶内開始位置
    LENGTH              As Integer                              ' 長さ
    NOWPROC             As String * 5                           ' 現在工程
    HOLDCLS             As String * 1                           ' ホールド区分 ---kuramoto 追加 2001/09/19----
    GRPFLG1             As Integer                              ' グループ情報
    GRPFLG2             As Integer                              ' グループ情報
    COLORFLG            As Boolean
    topHin              As cmkc001b_Wait3_HINBAN
    botHin              As cmkc001b_Wait3_HINBAN
End Type

Type cmkc001b_Wait3
    CRYNUM              As String * 12           ' 結晶番号
    blkInfo()           As cmkc001b_Wait3_BLK
End Type

'ブロック管理
Public Type typ_cmkc001f_Block
    'E040 ブロック管理
    INGOTPOS            As Integer              ' 結晶内開始位置
    LENGTH              As Integer              ' 長さ
    REALLEN             As Integer              ' 実長さ
    KRPROCCD            As String * 5           ' 現在管理工程
    NOWPROC             As String * 5           ' 現在工程
    LPKRPROCCD          As String * 5           ' 最終通過管理工程
    LASTPASS            As String * 5           ' 最終通過工程
    DELCLS              As String * 1           ' 削除区分
    RSTATCLS            As String * 1           ' 流動状態区分
    LSTATCLS            As String * 1           ' 最終状態区分 */
    'E037 結晶情報管理
    SEED                As String               'SEED
End Type


'仕様取得用
Public Type typ_cmkc001f_Disp
    '品番管理
    hinban              As String * 8            ' 品番
    INGOTPOS            As Integer               ' 結晶内開始位置
    REVNUM              As Integer               ' 製品番号改訂番号
    factory             As String * 1            ' 工場
    opecond             As String * 1            ' 操業条件
    LENGTH              As Integer               ' 長さ
    '製品仕様SXLデータ
    HSXD1CEN            As Double                ' 品ＳＸ直径１中心
    HSXRMIN             As Double                ' 品ＳＸ比抵抗下限
    HSXRMAX             As Double                ' 品ＳＸ比抵抗上限
    HSXRMBNP            As Double                ' 品ＳＸ比抵抗面内分布
    HSXRHWYS            As String * 1            ' 品ＳＸ比抵抗保証方法＿処
    HSXONMIN            As Double                ' 品ＳＸ酸素濃度下限
    HSXONMAX            As Double                ' 品ＳＸ酸素濃度上限
    HSXONMBP            As Double                ' 品ＳＸ酸素濃度面内分布
    HSXONHWS            As String * 1            ' 品ＳＸ酸素濃度保証方法＿処
    HSXCNMIN            As Double                ' 品ＳＸ炭素濃度下限
    HSXCNMAX            As Double                ' 品ＳＸ炭素濃度上限
    HSXCNHWS            As String * 1            ' 品ＳＸ炭素濃度保証方法＿処
    HSXTMMAX            As Double                ' 品ＳＸ転位密度上限         項目追加，修正対応 2003.05.20 yakimura
    HSXBMnAN(1 To 3)    As Double                ' 品ＳＸＢＭＤn 平均下限
    HSXBMnAX(1 To 3)    As Double                ' 品ＳＸＢＭＤn 平均上限
    HSXBMnHS(1 To 3)    As String * 1            ' 品ＳＸＢＭＤn 保証方法＿処
    HSXOFnAX(1 To 4)    As Double                ' 品ＳＸＯＳＦn平均上限
    HSXOFnMX(1 To 4)    As Double                ' 品ＳＸＯＳＦn上限
    HSXOFnHS(1 To 4)    As String * 1            ' 品ＳＸＯＳＦn 保証方法＿処
    HSXDENMX            As Integer               ' 品ＳＸＤｅｎ上限
    HSXDENMN            As Integer               ' 品ＳＸＤｅｎ下限
    HSXDENHS            As String * 1            ' 品ＳＸＤｅｎ保証方法＿処
    HSXDVDMX            As Integer               ' 品ＳＸＤＶＤ２上限
    HSXDVDMN            As Integer               ' 品ＳＸＤＶＤ２下限
    HSXDVDHS            As String * 1            ' 品ＳＸＤＶＤ２保証方法＿処
    HSXLDLMX            As Integer               ' 品ＳＸＬ／ＤＬ上限
    HSXLDLMN            As Integer               ' 品ＳＸＬ／ＤＬ下限
    HSXLDLHS            As String * 1            ' 品ＳＸＬ／ＤＬ保証方法＿処
    HSXLTMIN            As Integer               ' 品ＳＸＬタイム下限
    HSXLTMAX            As Integer               ' 品ＳＸＬタイム上限
    HSXLTHWS            As String * 1            ' 品ＳＸＬタイム保証方法＿処
    HSXDPDIR            As String * 2            ' 品ＳＸ溝位置方位
    HSXDPDRC            As String * 1            ' 品ＳＸ溝位置方向
    HSXDWMIN            As Double                ' 品ＳＸ溝巾下限
    HSXDWMAX            As Double                ' 品ＳＸ溝巾上限
    HSXDDMIN            As Double                ' 品ＳＸ溝深下限
    HSXDDMAX            As Double                ' 品ＳＸ溝深上限
    HSXD1MIN            As Double                ' 品ＳＸ直径１下限
    HSXD1MAX            As Double                ' 品ＳＸ直径１上限
    HSXCTCEN            As Double                ' 品ＳＸ結晶面傾縦中心
    HSXCYCEN            As Double                ' 品ＳＸ結晶面傾横中心
    EPDUP               As Integer               ' 結晶内側管理 EPD　上限
End Type


'実行時入力用
Public Type typ_cmkc001f_ExecCryIn
    CRYNUM              As String * 12       ' 結晶番号(IN)
    INGOTPOS            As Integer         ' インゴット内位置(IN)
End Type


'結晶最終検査
Public Type typ_cmkc001f_ExecFts
    LENGTH              As Integer           ' 長さ
    KRPROCCD            As String * 5      ' 管理工程コード
    PROCCODE            As String * 5      ' 工程コード
    PAYCLASS            As String * 1      ' 払い出し区分
    OUTLENGTH           As Integer        ' 出荷長さ
    PART(1 To 5)        As Integer     ' 部位n
    BDLEN(1 To 5)       As Integer    ' 部位n 不良長さ
    BDCAUS(1 To 5)      As String * 3 ' 部位n 不良理由
    TSTAFFID            As String * 8      ' 登録社員ID
End Type


'クリスタルカタログ受入
Public Type typ_cmkc001f_ExecCatalog
    CRYNUM              As String * 12       ' 結晶番号
    KRPROCCD            As String * 5      ' 管理工程コード
    PROCCODE            As String * 5      ' 工程コード
    BDCODE              As String * 3        ' 不良理由コード
    PALTNUM             As String * 4       ' パレット番号
    TSTAFFID            As String * 8      ' 登録社員ID
End Type

'------------------------------------------------------------------------
Type type_cmkc001b_SmpMng
    CRYNUM              As String * 12
    INGOTPOS            As Integer
    SMPKBN              As String * 1
    
    hinban              As String * 8            ' 品番
    REVNUM              As Integer               ' 製品番号改訂番号
    factory             As String * 1           ' 工場
    opecond             As String * 1           ' 操業条件
    
    
    CRYINDRS            As String * 1
    CRYRESRS            As String * 1
    CRYINDOI            As String * 1
    CRYRESOI            As String * 1
    CRYINDB1            As String * 1
    CRYRESB1            As String * 1
    CRYINDB2            As String * 1
    CRYRESB2            As String * 1
    CRYINDB3            As String * 1
    CRYRESB3            As String * 1
    CRYINDL1            As String * 1
    CRYRESL1            As String * 1
    CRYINDL2            As String * 1
    CRYRESL2            As String * 1
    CRYINDL3            As String * 1
    CRYRESL3            As String * 1
    CRYINDL4            As String * 1
    CRYRESL4            As String * 1
    CRYINDCS            As String * 1
    CRYRESCS            As String * 1
    CRYINDGD            As String * 1
    CRYRESGD            As String * 1
    CRYINDT             As String * 1
    CRYREST             As String * 1
    CRYINDEP            As String * 1
    CRYRESEP            As String * 1
    
    HSXCNHWS            As String * 1          ' 品ＳＸ炭素濃度保証方法＿処
    HSXLTHWS            As String * 1          ' 品ＳＸＬタイム保証方法＿処
    EPD                 As String * 1               ' EPD
End Type

Private Type tSmpMng
    BLOCKID             As String * 12
    TOPPOS              As Integer
    BOTPOS              As Integer
    
    CRYNUM              As String * 12
    INGOTPOS            As Integer
    SMPKBN              As String * 1
    
    hinban              As String * 8            ' 品番
    REVNUM              As Integer               ' 製品番号改訂番号
    factory             As String * 1           ' 工場
    opecond             As String * 1           ' 操業条件
    
    CRYINDRS            As String * 1
    CRYRESRS            As String * 1
    CRYINDOI            As String * 1
    CRYRESOI            As String * 1
    CRYINDB1            As String * 1
    CRYRESB1            As String * 1
    CRYINDB2            As String * 1
    CRYRESB2            As String * 1
    CRYINDB3            As String * 1
    CRYRESB3            As String * 1
    CRYINDL1            As String * 1
    CRYRESL1            As String * 1
    CRYINDL2            As String * 1
    CRYRESL2            As String * 1
    CRYINDL3            As String * 1
    CRYRESL3            As String * 1
    CRYINDL4            As String * 1
    CRYRESL4            As String * 1
    CRYINDCS            As String * 1
    CRYRESCS            As String * 1
    CRYINDGD            As String * 1
    CRYRESGD            As String * 1
    CRYINDT             As String * 1
    CRYREST             As String * 1
    CRYINDEP            As String * 1
    CRYRESEP            As String * 1
End Type

'待ち一覧
Public Type typ_HinMap      '2006/02
    HIN         As tFullHinban                  ' 品番
    LENGTH      As Integer                      ' 長さ
    Weight      As Double                       ' 重量
End Type


'初期表示用
Public Type type_DBDRV_scmzc_fcmkc001b_Disp
    CRYNUM              As String * 12          ' 結晶番号
    INGOTPOS            As Integer              ' 結晶内開始位置
    INPOS               As Integer              ' TOP位置
    LENGTH              As Integer              ' 長さ              '2001/11/8
    BLOCKID             As String * 12          ' ブロックID
    HSXTYPE             As String * 1           ' 品ＳＸタイプ
    HSXCDIR             As String * 1           ' 品ＳＸ結晶面方位
    UPDDATE             As Date                 ' 更新日付
    Judg                As String               ' 判定
    hinM()              As typ_HinMap           ' 品番(full)   '2006/02
    HOLDCLS             As String * 1           ' ホールド区分 ---kuramoto 追加 2001/09/25----
    SMP()               As type_cmkc001b_SmpMng ' サンプル管理
    PUPTN               As String               ' 引上ﾊﾟﾀｰﾝ   ---kubota 追加 2004/12/21----
    HOLDB               As String               ' 2005/08
    HOLDC               As String               ' 2005/08
    HOLDKT              As String               ' 2005/08
    LBLFLG              As String               ' ラベル発行フラグ  2005/11 ADD
    DIA                 As Integer              ' 直径 2006/02
    KIKBN               As String               '期判別区分 2006/11/14 SETsw kubota
    AGRSTATUS           As String               ' 承認確認区分 add SETkimizuka
    STOP                As String               ' 停止     add SETkimizuka
    CAUSE               As String               ' 停止理由 add SETkimizuka
    PRINTNO             As String               ' 先行評価 add SETkimizuka
End Type


'一覧表示用(出荷前一覧)    2008/05/26 SHINDOH
Public Type type_DBDRV_scmzc_fcmkc001b_Disp52
    PLANT               As String               ' 向先
    CRYNUM              As String * 12          ' 結晶番号
    INGOTPOS            As Integer              ' 結晶内開始位置
    INPOS               As Integer              ' TOP位置
    LENGTH              As Integer              ' 長さ
    BLOCKID             As String * 12          ' ブロックID
    HSXTYPE             As String * 1           ' 品ＳＸタイプ
    HSXCDIR             As String * 1           ' 品ＳＸ結晶面方位
    UPDDATE             As Date                 ' 更新日付
    hinM()              As typ_HinMap           ' 品番(full)
    HOLDCLS             As String * 1           ' ホールド区分
    PUPTN               As String               ' 引上ﾊﾟﾀｰﾝ
    HOLDB               As String               ' 2005/08
    HOLDC               As String               ' 2005/08
    HOLDKT              As String               ' 2005/08
End Type
'初期表示用         2008/05/26 SHINDOH
Public Type type_DBDRV_scmzc_fcmkc001b_Disp5
    CRYNUM              As String * 12                          ' 結晶番号
    INGOTPOS            As Integer                              ' 結晶内開始位置
'   LENGTH              As Integer                              ' 長さ              '2001/11/8
    BLOCKID             As String * 12                          ' ブロックID
    HSXTYPE             As String * 1                           ' 品ＳＸタイプ
    HSXCDIR             As String * 1                           ' 品ＳＸ結晶面方位
    UPDDATE             As Date                                 ' 更新日付
    Judg                As String                               ' 判定
    HIN()               As tFullHinban                          ' 品番(full)
    HOLDCLS             As String * 1                           ' ホールド区分 ---kuramoto 追加 2001/09/25----
    SMP()               As type_cmkc001b_SmpMng                 ' サンプル管理
    PUPTN               As String                               ' 引上ﾊﾟﾀｰﾝ    ---kubota 追加 2004/12/08----
    WFCUTT              As Integer                              ' WFｶｯﾄ単位　05/04/19 ooba
    BLOCKHFLAG          As String * 1                           ' ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞ　05/04/19 ooba
    HOLDBCA             As String * 1                           ' ﾎｰﾙﾄﾞ区分(XSDCA)　05/04/19 ooba
    HOLDB               As String               '2006/01
    HOLDC               As String               '2006/01
    HOLDKT              As String               '2006/01
    AGRSTATUS           As String           ' 承認確認区分 add SETkimizuka
    STOP                As String           ' 停止     add SETkimizuka
    CAUSE               As String           ' 停止理由 add SETkimizuka
    PRINTNO             As String           ' 先行評価 add SETkimizuka
End Type
'ブロックの情報 '2008/05/28 SHINDOH
Public Type typ_BlkData
    CRYNUM              As String * 12      ' 結晶番号
    BLOCKID             As String * 12      ' ブロックID
    INGOTPOS            As Integer          ' インゴット内位置
    LENGTH              As Integer          ' ブロック長さ
    REALLEN             As Integer          ' ブロック実長さ
    sBlockId            As String * 12      ' 払出先頭ブロックID
    BLOCKORDER          As Integer          ' ブロック順序
    DIAMETER            As Double           ' 直径 2002/05/01 S.Sano
    WFINDDATE           As String * 10      ' 最終抜試日付
    HOLDCLS             As String * 1       ' ホールド状態
End Type

'ブロック内品番情報 '2008/05/28 SHINDOH
Public Type typ_BlkHinMap
    BLOCKID             As String * 12      ' ブロックID
    HIN                 As tFullHinban      ' 品番
    REALLEN             As Integer          ' 品番実長さ
    HinLen              As Integer          ' 製品長
    PASSFLAG            As String * 1       ' 通過フラグ
    INPOSCA             As Integer          ' 結晶内開始位置　--- 2007/07/17 shindo 追加 ---
    PLANTCATCA          As String           ' 向先 2007/09/12 SPK Tsutsumi Add
End Type

' ブロック一覧  '2008/05/28 SHINDOH
Public Type typ_BlkMap
    BLOCKID             As String * 12      ' ブロックID
    HIN(1 To 5)         As tFullHinban      ' 品番
    WFINDDATE           As String * 10      ' 最終抜試日付
    CRYNUM              As String * 12      ' 結晶番号
    INGOTPOS            As Integer          ' インゴット内位置
    LENGTH              As Integer          ' ブロック長さ
    REALLEN             As Integer          ' ブロック実長さ
    HINREALLEN(1 To 5)  As Integer          ' 品番実長さ
    HinLen(1 To 5)      As Integer          ' 品番長さ
    DIAMETER            As Double           ' 直径 2002/05/01 S.Sano
    sBlockId            As String * 12      ' 先頭ブロックID
    BLOCKORDER          As Integer          ' ブロック順序
    HOLDCLS             As String * 1       ' ホールド状態  --- 2001/09/19 kuramoto 追加 ---
    PASSFLAG            As String * 1       ' 通過フラグ　　--- 200/04/16 Yam
End Type
''ブロック内品番情報(構成品番取得用)　　--- '2008/05/28 SHINDOH
Public Type typ_WkBlkMap
    BLOCKID             As String * 12      ' ブロックID
    HINCNT As Integer
    HIN()         As tFullHinban      ' 品番
    HINREALLEN()  As Integer          ' 品番実長さ
    HinLen()      As Integer          ' 品番長さ
    INPOSCA() As Integer '結晶内開始位置
End Type

'品番情報--- '2008/05/28 SHINDOH
Public Wk_tblBlkMap() As typ_WkBlkMap


'品番、仕様、結晶内側取得用 (TOP,TAIL順で２レコード取得)
Public Type type_DBDRV_scmzc_fcmkc001c_Siyou
    'ブロック管理
    CRYNUM              As String * 12          ' 結晶番号
    INGOTPOS            As Integer              ' 結晶内開始位置
    LENGTH              As Integer              ' 長さ
    '品番管理
    HIN                 As tFullHinban          ' 品番(full)
        
    '結晶情報
    PRODCOND            As String * 4           ' 製作条件
    PGID                As String * 8           ' ＰＧ－ＩＤ
    UPLENGTH            As Integer              ' 引上げ長さ
    FREELENG            As Integer              ' フリー長
    DIAMETER            As Integer              ' 直径 2002/05/01 S.Sano
    CHARGE              As Double               ' チャージ量
    SEED                As String * 4           ' シード
    ADDDPPOS            As Integer              ' 追加ドープ位置

    '製品仕様
    HSXTYPE             As String * 1           ' 品ＳＸタイプ
    HSXD1CEN            As Double               ' 品ＳＸ直径１中心
    HSXCDIR             As String * 1           ' 品ＳＸ結晶面方位
    HSXRMIN             As Double               ' 品ＳＸ比抵抗下限
    HSXRMAX             As Double               ' 品ＳＸ比抵抗上限
    HSXRAMIN            As Double               ' 品ＳＸ比抵抗平均下限
    HSXRAMAX            As Double               ' 品ＳＸ比抵抗平均上限
    HSXRMBNP            As Double               ' 品ＳＸ比抵抗面内分布
    HSXRSPOH            As String * 1           ' 品ＳＸ比抵抗測定位置＿方
    HSXRSPOT            As String * 1           ' 品ＳＸ比抵抗測定位置＿点
    HSXRSPOI            As String * 1           ' 品ＳＸ比抵抗測定位置＿位
    HSXRHWYT            As String * 1           ' 品ＳＸ比抵抗保証方法＿対
    HSXRHWYS            As String * 1           ' 品ＳＸ比抵抗保証方法＿処

    HSXONMIN            As Double               ' 品ＳＸ酸素濃度下限
    HSXONMAX            As Double               ' 品ＳＸ酸素濃度上限
    HSXONAMN            As Double               ' 品ＳＸ酸素濃度平均下限
    HSXONAMX            As Double               ' 品ＳＸ酸素濃度平均上限
    HSXONMBP            As Double               ' 品ＳＸ酸素濃度面内分布
    HSXONSPH            As String * 1           ' 品ＳＸ酸素濃度測定位置＿方
    HSXONSPT            As String * 1           ' 品ＳＸ酸素濃度測定位置＿点
    HSXONSPI            As String * 1           ' 品ＳＸ酸素濃度測定位置＿位
    HSXONHWT            As String * 1           ' 品ＳＸ酸素濃度保証方法＿対
    HSXONHWS            As String * 1           ' 品ＳＸ酸素濃度保証方法＿処

    HSXBM1AN            As Double               ' 品ＳＸＢＭＤ１平均下限
    HSXBM1AX            As Double               ' 品ＳＸＢＭＤ１平均上限
    HSXBM2AN            As Double               ' 品ＳＸＢＭＤ２平均下限
    HSXBM2AX            As Double               ' 品ＳＸＢＭＤ２平均上限
    HSXBM3AN            As Double               ' 品ＳＸＢＭＤ３平均下限
    HSXBM3AX            As Double               ' 品ＳＸＢＭＤ３平均上限
    HSXBM1SH            As String * 1           ' 品ＳＸＢＭＤ１測定位置＿方
    HSXBM1ST            As String * 1           ' 品ＳＸＢＭＤ１測定位置＿点
    HSXBM1SR            As String * 1           ' 品ＳＸＢＭＤ１測定位置＿領
    HSXBM1HT            As String * 1           ' 品ＳＸＢＭＤ１保証方法＿対
    HSXBM1HS            As String * 1           ' 品ＳＸＢＭＤ１保証方法＿処
    HSXBM2SH            As String * 1           ' 品ＳＸＢＭＤ２測定位置＿方
    HSXBM2ST            As String * 1           ' 品ＳＸＢＭＤ２測定位置＿点
    HSXBM2SR            As String * 1           ' 品ＳＸＢＭＤ２測定位置＿領
    HSXBM2HT            As String * 1           ' 品ＳＸＢＭＤ２保証方法＿対
    HSXBM2HS            As String * 1           ' 品ＳＸＢＭＤ２保証方法＿処
    HSXBM3SH            As String * 1           ' 品ＳＸＢＭＤ３測定位置＿方
    HSXBM3ST            As String * 1           ' 品ＳＸＢＭＤ３測定位置＿点
    HSXBM3SR            As String * 1           ' 品ＳＸＢＭＤ３測定位置＿領
    HSXBM3HT            As String * 1           ' 品ＳＸＢＭＤ３保証方法＿対
    HSXBM3HS            As String * 1           ' 品ＳＸＢＭＤ３保証方法＿処

    HSXOS1AX            As Double               ' 品ＳＸＯＳＦ１平均上限
    HSXOS1MX            As Double               ' 品ＳＸＯＳＦ１上限
    HSXOS2AX            As Double               ' 品ＳＸＯＳＦ２平均上限
    HSXOS2MX            As Double               ' 品ＳＸＯＳＦ２上限
    HSXOS3AX            As Double               ' 品ＳＸＯＳＦ３平均上限
    HSXOS3MX            As Double               ' 品ＳＸＯＳＦ３上限
    HSXOS4AX            As Double               ' 品ＳＸＯＳＦ４平均上限
    HSXOS4MX            As Double               ' 品ＳＸＯＳＦ４上限
    HSXOS1SH            As String * 1           ' 品ＳＸＯＳＦ１測定位置＿方
    HSXOS1ST            As String * 1           ' 品ＳＸＯＳＦ１測定位置＿点
    HSXOS1SR            As String * 1           ' 品ＳＸＯＳＦ１測定位置＿領
    HSXOS1HT            As String * 1           ' 品ＳＸＯＳＦ１保証方法＿対
    HSXOS1HS            As String * 1           ' 品ＳＸＯＳＦ１保証方法＿処
    HSXOS2SH            As String * 1           ' 品ＳＸＯＳＦ２測定位置＿方
    HSXOS2ST            As String * 1           ' 品ＳＸＯＳＦ２測定位置＿点
    HSXOS2SR            As String * 1           ' 品ＳＸＯＳＦ２測定位置＿領
    HSXOS2HT            As String * 1           ' 品ＳＸＯＳＦ２保証方法＿対
    HSXOS2HS            As String * 1           ' 品ＳＸＯＳＦ２保証方法＿処
    HSXOS3SH            As String * 1           ' 品ＳＸＯＳＦ３測定位置＿方
    HSXOS3ST            As String * 1           ' 品ＳＸＯＳＦ３測定位置＿点
    HSXOS3SR            As String * 1           ' 品ＳＸＯＳＦ３測定位置＿領
    HSXOS3HT            As String * 1           ' 品ＳＸＯＳＦ３保証方法＿対
    HSXOS3HS            As String * 1           ' 品ＳＸＯＳＦ３保証方法＿処
    HSXOS4SH            As String * 1           ' 品ＳＸＯＳＦ４測定位置＿方
    HSXOS4ST            As String * 1           ' 品ＳＸＯＳＦ４測定位置＿点
    HSXOS4SR            As String * 1           ' 品ＳＸＯＳＦ４測定位置＿領
    HSXOS4HT            As String * 1           ' 品ＳＸＯＳＦ４保証方法＿対
    HSXOS4HS            As String * 1           ' 品ＳＸＯＳＦ４保証方法＿処
    HSXOS1NS            As String * 2           ' 品ＳＸＯＳＦ１熱処理法
    HSXOS2NS            As String * 2           ' 品ＳＸＯＳＦ２熱処理法
    HSXOS3NS            As String * 2           ' 品ＳＸＯＳＦ３熱処理法
    HSXOS4NS            As String * 2           ' 品ＳＸＯＳＦ４熱処理法
    HSXBM1NS            As String * 2           ' 品ＳＸＢＭＤ１熱処理法
    HSXBM2NS            As String * 2           ' 品ＳＸＢＭＤ２熱処理法
    HSXBM3NS            As String * 2           ' 品ＳＸＢＭＤ３熱処理法

    HSXCNMIN            As Double               ' 品ＳＸ炭素濃度下限
    HSXCNMAX            As Double               ' 品ＳＸ炭素濃度上限
    HSXCNSPH            As String * 1           ' 品ＳＸ炭素濃度測定位置＿方
    HSXCNSPT            As String * 1           ' 品ＳＸ炭素濃度測定位置＿点
    HSXCNSPI            As String * 1           ' 品ＳＸ炭素濃度測定位置＿位
    HSXCNHWT            As String * 1           ' 品ＳＸ炭素濃度保証方法＿対
    HSXCNHWS            As String * 1           ' 品ＳＸ炭素濃度保証方法＿処

    HSXDENMX            As Integer              ' 品ＳＸＤｅｎ上限
    HSXDENMN            As Integer              ' 品ＳＸＤｅｎ下限
    HSXLDLMX            As Integer              ' 品ＳＸＬ／ＤＬ上限
    HSXLDLMN            As Integer              ' 品ＳＸＬ／ＤＬ下限
    HSXDVDMX            As Integer              ' 品ＳＸＤＶＤ２上限
    HSXDVDMN            As Integer              ' 品ＳＸＤＶＤ２下限
    HSXDENHT            As String * 1           ' 品ＳＸＤｅｎ保証方法＿対
    HSXDENHS            As String * 1           ' 品ＳＸＤｅｎ保証方法＿処
    HSXLDLHT            As String * 1           ' 品ＳＸＬ／ＤＬ保証方法＿対
    HSXLDLHS            As String * 1           ' 品ＳＸＬ／ＤＬ保証方法＿処
    HSXDVDHT            As String * 1           ' 品ＳＸＤＶＤ２保証方法＿対
    HSXDVDHS            As String * 1           ' 品ＳＸＤＶＤ２保証方法＿処
    HSXDENKU            As String * 1           ' 品ＳＸＤｅｎ検査有無
    HSXDVDKU            As String * 1           ' 品ＳＸＤＶＤ２検査有無
    HSXLDLKU            As String * 1           ' 品ＳＸＬ／ＤＬ検査有無

    HSXLTMIN            As Integer              ' 品ＳＸＬタイム下限
    HSXLTMAX            As Integer              ' 品ＳＸＬタイム上限
    HSXLTSPH            As String * 1           ' 品ＳＸＬタイム測定位置＿方
    HSXLTSPT            As String * 1           ' 品ＳＸＬタイム測定位置＿点
    HSXLTSPI            As String * 1           ' 品ＳＸＬタイム測定位置＿位
    HSXLTHWT            As String * 1           ' 品ＳＸＬタイム保証方法＿対
    HSXLTHWS            As String * 1           ' 品ＳＸＬタイム保証方法＿処
    '結晶内側管理
    EPDUP               As Integer              ' EPD　上限
End Type


' 結晶サンプル管理取得用 (TOP,TAIL順で２レコード取得)
Public Type type_DBDRV_scmzc_fcmkc001c_CrySmp
    CRYNUM              As String * 12          ' 結晶番号
    INGOTPOS            As Integer              ' 結晶内位置
    LENGTH              As Integer              ' 長さ
    BLOCKID             As String * 12          ' ブロックID
    SMPKBN              As String * 1           ' サンプル区分
    SMPLNO              As Integer              ' サンプルNo
    hinban              As String * 12          ' 品番
    REVNUM              As Integer              ' 製品番号改訂番号
    factory             As String * 1           ' 工場
    opecond             As String * 1           ' 操業条件
    KTKBN               As String * 1           ' 確定区分
    CRYINDRS            As String * 1           ' 結晶検査指示（Rs)
    CRYINDOI            As String * 1           ' 結晶検査指示（Oi)
    CRYINDB1            As String * 1           ' 結晶検査指示（B1)
    CRYINDB2            As String * 1           ' 結晶検査指示（B2）
    CRYINDB3            As String * 1           ' 結晶検査指示（B3)
    CRYINDL1            As String * 1           ' 結晶検査指示（L1)
    CRYINDL2            As String * 1           ' 結晶検査指示（L2)
    CRYINDL3            As String * 1           ' 結晶検査指示（L3)
    CRYINDL4            As String * 1           ' 結晶検査指示（L4)
    CRYINDCS            As String * 1           ' 結晶検査指示（Cs)
    CRYINDGD            As String * 1           ' 結晶検査指示（GD)
    CRYINDT             As String * 1           ' 結晶検査指示（T)
    CRYINDEP            As String * 1           ' 結晶検査指示（EPD)
End Type


'結晶抵抗実績
Public Type type_DBDRV_scmzc_fcmkc001c_CryR
    CRYNUM              As String * 12          ' 結晶番号
    POSITION            As Integer              ' 位置
    SMPKBN              As String * 1           ' サンプル区分
    SMPLNO              As Integer              ' サンプルＮｏ
    SMPLUMU             As String * 1           ' サンプル有無
    TRANCOND            As String * 1           ' 処理条件
    MEAS1               As Double               ' 測定値１
    MEAS2               As Double               ' 測定値２
    MEAS3               As Double               ' 測定値３
    MEAS4               As Double               ' 測定値４
    MEAS5               As Double               ' 測定値５
    RRG                 As Double               ' ＲＲＧ
    REGDATE             As Date                 ' 登録日付
End Type


'Oi実績
Public Type type_DBDRV_scmzc_fcmkc001c_Oi
    CRYNUM              As String * 12          ' 結晶番号
    POSITION            As Integer              ' 位置
    SMPKBN              As String * 1           ' サンプル区分
    SMPLNO              As Integer              ' サンプルＮｏ
    SMPLUMU             As String * 1           ' サンプル有無
    TRANCOND            As String * 1           ' 処理条件
    OIMEAS1             As Double               ' Ｏｉ測定値１
    OIMEAS2             As Double               ' Ｏｉ測定値２
    OIMEAS3             As Double               ' Ｏｉ測定値３
    OIMEAS4             As Double               ' Ｏｉ測定値４
    OIMEAS5             As Double               ' Ｏｉ測定値５
    ORGRES              As Double               ' ＯＲＧ結果
    AVE                 As Double               ' ＡＶＥ
    FTIRCONV            As Double               ' ＦＴＩＲ換算
    INSPECTWAY          As String * 2           ' 検査方法
    REGDATE             As Date                 ' 登録日付
End Type


'BMD1～3実績
Public Type type_DBDRV_scmzc_fcmkc001c_BMD
    CRYNUM              As String * 12          ' 結晶番号
    POSITION            As Integer              ' 位置
    SMPKBN              As String * 1           ' サンプル区分
    SMPLNO              As Integer              ' サンプルＮｏ
    SMPLUMU             As String * 1           ' サンプル有無
    HTPRC               As String * 2           ' 熱処理方法
    KKSP                As String * 3           ' 結晶欠陥測定位置
    KKSET               As String * 3           ' 結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    TRANCOND            As String * 1           ' 処理条件
    MEAS1               As Double               ' 測定値１
    MEAS2               As Double               ' 測定値２
    MEAS3               As Double               ' 測定値３
    MEAS4               As Double               ' 測定値４
    MEAS5               As Double               ' 測定値５
    Min                 As Double               ' MIN
    max                 As Double               ' MAX
    AVE                 As Double               ' AVE
    REGDATE             As Date                 ' 登録日付
End Type


'OSF1～4実績
Public Type type_DBDRV_scmzc_fcmkc001c_OSF
    CRYNUM              As String * 12          ' 結晶番号
    POSITION            As Integer              ' 位置
    SMPKBN              As String * 1           ' サンプル区分
    SMPLNO              As Integer              ' サンプルＮｏ
    SMPLUMU             As String * 1           ' サンプル有無
    HTPRC               As String * 2           ' 熱処理方法
    KKSP                As String * 3           ' 結晶欠陥測定位置
    KKSET               As String * 3           ' 結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    TRANCOND            As String * 1           ' 処理条件
    CALCMAX             As Double               ' 計算結果 Max
    CALCAVE             As Double               ' 計算結果 Ave
    MEAS1               As Double               ' 測定値１
    MEAS2               As Double               ' 測定値２
    MEAS3               As Double               ' 測定値３
    MEAS4               As Double               ' 測定値４
    MEAS5               As Double               ' 測定値５
    MEAS6               As Double               ' 測定値６
    MEAS7               As Double               ' 測定値７
    MEAS8               As Double               ' 測定値８
    MEAS9               As Double               ' 測定値９
    MEAS10              As Double               ' 測定値１０
    MEAS11              As Double               ' 測定値１１
    MEAS12              As Double               ' 測定値１２
    MEAS13              As Double               ' 測定値１３
    MEAS14              As Double               ' 測定値１４
    MEAS15              As Double               ' 測定値１５
    MEAS16              As Double               ' 測定値１６
    MEAS17              As Double               ' 測定値１７
    MEAS18              As Double               ' 測定値１８
    MEAS19              As Double               ' 測定値１９
    MEAS20              As Double               ' 測定値２０
    REGDATE             As Date                 ' 登録日付
End Type


'CS実績
Public Type type_DBDRV_scmzc_fcmkc001c_CS
    CRYNUM              As String * 12          ' 結晶番号
    POSITION            As Integer              ' 位置
    SMPKBN              As String * 1           ' サンプル区分
    SMPLNO              As Integer              ' サンプルＮｏ
    SMPLUMU             As String * 1           ' サンプル有無
    TRANCOND            As String * 1           ' 処理条件
    CSMEAS              As Double               ' Cs実測値
    PRE70P              As Double               ' ７０％推定値
    REGDATE             As Date                 ' 登録日付
End Type


'GD実績
Public Type type_DBDRV_scmzc_fcmkc001c_GD
    CRYNUM              As String * 12          ' 結晶番号
    POSITION            As Integer              ' 位置
    SMPKBN              As String * 1           ' サンプル区分
    SMPLNO              As Integer              ' サンプルＮｏ
    SMPLUMU             As String * 1           ' サンプル有無
    TRANCOND            As String * 1           ' 処理条件
    MSRSDEN             As Integer              ' 測定結果 Den
    MSRSLDL             As Integer              ' 測定結果 L/DL
    MSRSDVD2            As Integer              ' 測定結果 DVD2
    MS01LDL1            As Integer              ' 測定値01 L/DL1
    MS01LDL2            As Integer              ' 測定値01 L/DL2
    MS01LDL3            As Integer              ' 測定値01 L/DL3
    MS01LDL4            As Integer              ' 測定値01 L/DL4
    MS01LDL5            As Integer              ' 測定値01 L/DL5
    MS01DEN1            As Integer              ' 測定値01 Den1
    MS01DEN2            As Integer              ' 測定値01 Den2
    MS01DEN3            As Integer              ' 測定値01 Den3
    MS01DEN4            As Integer              ' 測定値01 Den4
    MS01DEN5            As Integer              ' 測定値01 Den5
    MS02LDL1            As Integer              ' 測定値02 L/DL1
    MS02LDL2            As Integer              ' 測定値02 L/DL2
    MS02LDL3            As Integer              ' 測定値02 L/DL3
    MS02LDL4            As Integer              ' 測定値02 L/DL4
    MS02LDL5            As Integer              ' 測定値02 L/DL5
    MS02DEN1            As Integer              ' 測定値02 Den1
    MS02DEN2            As Integer              ' 測定値02 Den2
    MS02DEN3            As Integer              ' 測定値02 Den3
    MS02DEN4            As Integer              ' 測定値02 Den4
    MS02DEN5            As Integer              ' 測定値02 Den5
    MS03LDL1            As Integer              ' 測定値03 L/DL1
    MS03LDL2            As Integer              ' 測定値03 L/DL2
    MS03LDL3            As Integer              ' 測定値03 L/DL3
    MS03LDL4            As Integer              ' 測定値03 L/DL4
    MS03LDL5            As Integer              ' 測定値03 L/DL5
    MS03DEN1            As Integer              ' 測定値03 Den1
    MS03DEN2            As Integer              ' 測定値03 Den2
    MS03DEN3            As Integer              ' 測定値03 Den3
    MS03DEN4            As Integer              ' 測定値03 Den4
    MS03DEN5            As Integer              ' 測定値03 Den5
    MS04LDL1            As Integer              ' 測定値04 L/DL1
    MS04LDL2            As Integer              ' 測定値04 L/DL2
    MS04LDL3            As Integer              ' 測定値04 L/DL3
    MS04LDL4            As Integer              ' 測定値04 L/DL4
    MS04LDL5            As Integer              ' 測定値04 L/DL5
    MS04DEN1            As Integer              ' 測定値04 Den1
    MS04DEN2            As Integer              ' 測定値04 Den2
    MS04DEN3            As Integer              ' 測定値04 Den3
    MS04DEN4            As Integer              ' 測定値04 Den4
    MS04DEN5            As Integer              ' 測定値04 Den5
    MS05LDL1            As Integer              ' 測定値05 L/DL1
    MS05LDL2            As Integer              ' 測定値05 L/DL2
    MS05LDL3            As Integer              ' 測定値05 L/DL3
    MS05LDL4            As Integer              ' 測定値05 L/DL4
    MS05LDL5            As Integer              ' 測定値05 L/DL5
    MS05DEN1            As Integer              ' 測定値05 Den1
    MS05DEN2            As Integer              ' 測定値05 Den2
    MS05DEN3            As Integer              ' 測定値05 Den3
    MS05DEN4            As Integer              ' 測定値05 Den4
    MS05DEN5            As Integer              ' 測定値05 Den5
    MS06LDL1            As Integer              ' 測定値06 L/DL1
    MS06LDL2            As Integer              ' 測定値06 L/DL2
    MS06LDL3            As Integer              ' 測定値06 L/DL3
    MS06LDL4            As Integer              ' 測定値06 L/DL4
    MS06LDL5            As Integer              ' 測定値06 L/DL5
    MS06DEN1            As Integer              ' 測定値06 Den1
    MS06DEN2            As Integer              ' 測定値06 Den2
    MS06DEN3            As Integer              ' 測定値06 Den3
    MS06DEN4            As Integer              ' 測定値06 Den4
    MS06DEN5            As Integer              ' 測定値06 Den5
    MS07LDL1            As Integer              ' 測定値07 L/DL1
    MS07LDL2            As Integer              ' 測定値07 L/DL2
    MS07LDL3            As Integer              ' 測定値07 L/DL3
    MS07LDL4            As Integer              ' 測定値07 L/DL4
    MS07LDL5            As Integer              ' 測定値07 L/DL5
    MS07DEN1            As Integer              ' 測定値07 Den1
    MS07DEN2            As Integer              ' 測定値07 Den2
    MS07DEN3            As Integer              ' 測定値07 Den3
    MS07DEN4            As Integer              ' 測定値07 Den4
    MS07DEN5            As Integer              ' 測定値07 Den5
    MS08LDL1            As Integer              ' 測定値08 L/DL1
    MS08LDL2            As Integer              ' 測定値08 L/DL2
    MS08LDL3            As Integer              ' 測定値08 L/DL3
    MS08LDL4            As Integer              ' 測定値08 L/DL4
    MS08LDL5            As Integer              ' 測定値08 L/DL5
    MS08DEN1            As Integer              ' 測定値08 Den1
    MS08DEN2            As Integer              ' 測定値08 Den2
    MS08DEN3            As Integer              ' 測定値08 Den3
    MS08DEN4            As Integer              ' 測定値08 Den4
    MS08DEN5            As Integer              ' 測定値08 Den5
    MS09LDL1            As Integer              ' 測定値09 L/DL1
    MS09LDL2            As Integer              ' 測定値09 L/DL2
    MS09LDL3            As Integer              ' 測定値09 L/DL3
    MS09LDL4            As Integer              ' 測定値09 L/DL4
    MS09LDL5            As Integer              ' 測定値09 L/DL5
    MS09DEN1            As Integer              ' 測定値09 Den1
    MS09DEN2            As Integer              ' 測定値09 Den2
    MS09DEN3            As Integer              ' 測定値09 Den3
    MS09DEN4            As Integer              ' 測定値09 Den4
    MS09DEN5            As Integer              ' 測定値09 Den5
    MS10LDL1            As Integer              ' 測定値10 L/DL1
    MS10LDL2            As Integer              ' 測定値10 L/DL2
    MS10LDL3            As Integer              ' 測定値10 L/DL3
    MS10LDL4            As Integer              ' 測定値10 L/DL4
    MS10LDL5            As Integer              ' 測定値10 L/DL5
    MS10DEN1            As Integer              ' 測定値10 Den1
    MS10DEN2            As Integer              ' 測定値10 Den2
    MS10DEN3            As Integer              ' 測定値10 Den3
    MS10DEN4            As Integer              ' 測定値10 Den4
    MS10DEN5            As Integer              ' 測定値10 Den5
    MS11LDL1            As Integer              ' 測定値11 L/DL1
    MS11LDL2            As Integer              ' 測定値11 L/DL2
    MS11LDL3            As Integer              ' 測定値11 L/DL3
    MS11LDL4            As Integer              ' 測定値11 L/DL4
    MS11LDL5            As Integer              ' 測定値11 L/DL5
    MS11DEN1            As Integer              ' 測定値11 Den1
    MS11DEN2            As Integer              ' 測定値11 Den2
    MS11DEN3            As Integer              ' 測定値11 Den3
    MS11DEN4            As Integer              ' 測定値11 Den4
    MS11DEN5            As Integer              ' 測定値11 Den5
    MS12LDL1            As Integer              ' 測定値12 L/DL1
    MS12LDL2            As Integer              ' 測定値12 L/DL2
    MS12LDL3            As Integer              ' 測定値12 L/DL3
    MS12LDL4            As Integer              ' 測定値12 L/DL4
    MS12LDL5            As Integer              ' 測定値12 L/DL5
    MS12DEN1            As Integer              ' 測定値12 Den1
    MS12DEN2            As Integer              ' 測定値12 Den2
    MS12DEN3            As Integer              ' 測定値12 Den3
    MS12DEN4            As Integer              ' 測定値12 Den4
    MS12DEN5            As Integer              ' 測定値12 Den5
    MS13LDL1            As Integer              ' 測定値13 L/DL1
    MS13LDL2            As Integer              ' 測定値13 L/DL2
    MS13LDL3            As Integer              ' 測定値13 L/DL3
    MS13LDL4            As Integer              ' 測定値13 L/DL4
    MS13LDL5            As Integer              ' 測定値13 L/DL5
    MS13DEN1            As Integer              ' 測定値13 Den1
    MS13DEN2            As Integer              ' 測定値13 Den2
    MS13DEN3            As Integer              ' 測定値13 Den3
    MS13DEN4            As Integer              ' 測定値13 Den4
    MS13DEN5            As Integer              ' 測定値13 Den5
    MS14LDL1            As Integer              ' 測定値14 L/DL1
    MS14LDL2            As Integer              ' 測定値14 L/DL2
    MS14LDL3            As Integer              ' 測定値14 L/DL3
    MS14LDL4            As Integer              ' 測定値14 L/DL4
    MS14LDL5            As Integer              ' 測定値14 L/DL5
    MS14DEN1            As Integer              ' 測定値14 Den1
    MS14DEN2            As Integer              ' 測定値14 Den2
    MS14DEN3            As Integer              ' 測定値14 Den3
    MS14DEN4            As Integer              ' 測定値14 Den4
    MS14DEN5            As Integer              ' 測定値14 Den5
    MS15LDL1            As Integer              ' 測定値15 L/DL1
    MS15LDL2            As Integer              ' 測定値15 L/DL2
    MS15LDL3            As Integer              ' 測定値15 L/DL3
    MS15LDL4            As Integer              ' 測定値15 L/DL4
    MS15LDL5            As Integer              ' 測定値15 L/DL5
    MS15DEN1            As Integer              ' 測定値15 Den1
    MS15DEN2            As Integer              ' 測定値15 Den2
    MS15DEN3            As Integer              ' 測定値15 Den3
    MS15DEN4            As Integer              ' 測定値15 Den4
    MS15DEN5            As Integer              ' 測定値15 Den5
    REGDATE             As Date                 ' 登録日付
End Type


'ライフタイム実績取得関数
Public Type type_DBDRV_scmzc_fcmkc001c_LT
    CRYNUM              As String * 12          ' 結晶番号
    POSITION            As Integer              ' 位置
    SMPKBN              As String * 1           ' サンプル区分
    SMPLNO              As Integer              ' サンプルＮｏ
    SMPLUMU             As String * 1           ' サンプル有無
    MEAS1               As Integer              ' 測定値１
    MEAS2               As Integer              ' 測定値２
    MEAS3               As Integer              ' 測定値３
    MEAS4               As Integer              ' 測定値４
    MEAS5               As Integer              ' 測定値５
    TRANCOND            As String * 1           ' 処理条件
    MEASPEAK            As Integer              ' 測定値 ピーク値
    CALCMEAS            As Integer              ' 計算結果
    REGDATE             As Date                 ' 登録日付
    LTSPI               As String               ' 測定位置コード
End Type


'EPD実績取得関数
Public Type type_DBDRV_scmzc_fcmkc001c_EPD
    CRYNUM              As String * 12          ' 結晶番号
    POSITION            As Integer              ' 位置
    SMPKBN              As String * 1           ' サンプル区分
    SMPLNO              As Integer              ' サンプルＮｏ
    SMPLUMU             As String * 1           ' サンプル有無
    TRANCOND            As String * 1           ' 処理条件
    MEASURE             As Integer              ' 測定値
    REGDATE             As Date                 ' 登録日付
End Type


'実績をまとめた構造体
Public Type type_DBDRV_scmzc_fcmkc001c_Zisseki
    CRYRZ()             As type_DBDRV_scmzc_fcmkc001c_CryR
    OIZ()               As type_DBDRV_scmzc_fcmkc001c_Oi
    BMD1Z()             As type_DBDRV_scmzc_fcmkc001c_BMD
    BMD2Z()             As type_DBDRV_scmzc_fcmkc001c_BMD
    BMD3Z()             As type_DBDRV_scmzc_fcmkc001c_BMD
    OSF1Z()             As type_DBDRV_scmzc_fcmkc001c_OSF
    OSF2Z()             As type_DBDRV_scmzc_fcmkc001c_OSF
    OSF3Z()             As type_DBDRV_scmzc_fcmkc001c_OSF
    OSF4Z()             As type_DBDRV_scmzc_fcmkc001c_OSF
    csz()               As type_DBDRV_scmzc_fcmkc001c_CS
    GDZ()               As type_DBDRV_scmzc_fcmkc001c_GD
    LTZ()               As type_DBDRV_scmzc_fcmkc001c_LT
    EPDZ()              As type_DBDRV_scmzc_fcmkc001c_EPD
    SURSZ()             As type_DBDRV_scmzc_fcmkc001c_CryR
End Type

'*** UPDATE START T.TERAUCHI 2004/12/19 CX作成用構造体設定
Public Type type_DBDRV_fcmkc001c_InsXodcx
    BLOCKID     As String               ' ブロックID
    CRYNUM      As String * 12          ' 結晶番号
    INGOTPOS    As Integer              ' 結晶内開始位置
    LASTPASS    As String * 5           ' 最終通過工程
    LENGTH      As Integer              ' 長さ
    Weight      As Double               ' 重量
    STAFFID     As String               ' 担当者コード
    STAFFNAME   As String               ' 担当者名
End Type
'*** UPDATE END   T.TERAUCHI 2004/12/19

' 2007/08/30 SPK Tsutsumi Add Start
Public Type typ_Mukesaki
    sMukeCode As String     '' 向先コード
    sMukeName As String     '' 向先名
End Type

Public s_CmbMukesaki() As typ_Mukesaki
Public s_Mukesaki() As typ_Mukesaki
' 2007/08/30 SPK Tsutsumi Add End
'2008/05/30 SHINDOH----------------------------------------------
' 配列初期化値
Public Const DEF_PARAM_VALUE_LT = -1
' ライフタイム測定点数（新データは１０点固定）
Public Const SS_SOKUETI_TENSU = 10
' ライフタイム測定点数（旧データは５点固定）
Public Const SS_SOKUETI_TENSU_OLD = 5
'2008/05/30 SHINDOH----------------------------------------------


Public Type typ_TBCMX011
BLOCKID As String
FROMTOKBN As String
TRANCNT As Integer
STCID As String
hinban As String
REVNUM As String
factory As String
opecond As String
STCKNNUM As String
CRYNUM As String
CRYDECDATE As Date
PLUPDATE As Date
UPLENGTH As Integer
FREELENG As Integer
INGOTPOS As Integer
BlkLen As Integer
BLKWGHT As Long
LENGTH As Integer
Weight As Long
MCNO As String
PGID As String
DM1 As Integer
DM2 As Integer
NCHDPTH As Integer
CHARGE As Double
SEED As String
SXL_RS_SMPPOS As Integer
SXLRS_MEAS1 As Double
SXLRS_MEAS2 As Double
SXLRS_MEAS3 As Double
SXLRS_MEAS4 As Double
SXLRS_MEAS5 As Double
SXLRS_EFEHS As Double
SXLRS_RRG As Double
SXL_OI_SMPPOS As Integer
SXLOI_OIMEAS1 As Double
SXLOI_OIMEAS2 As Double
SXLOI_OIMEAS3 As Double
SXLOI_OIMEAS4 As Double
SXLOI_OIMEAS5 As Double
SXLOI_ORGRES As Double
SXLOI_INSPECTWAY As String
SXL_CS_SMPPOS As Integer
SXLCS_CSMEAS As Double
SXLCS_70PPRE As Double
SXLCS_BSUIME As Double
SXLOSF_SMPPOS As Integer
SXLOSF1_KKSP As String
SXLOSF1_NETU As String
SXLOSF1_KKSET As String
SXLOSF1_CALCMAX As Double
SXLOSF1_CALCAVE As Double
SXLOSF2_KKSP As String
SXLOSF2_NETU As String
SXLOSF2_KKSET As String
SXLOSF2_CALCMAX As Double
SXLOSF2_CALCAVE As Double
SXLOSF3_KKSP As String
SXLOSF3_NETU As String
SXLOSF3_KKSET As String
SXLOSF3_CALCMAX As Double
SXLOSF3_CALCAVE As Double
SXLOSF4_KKSP As String
SXLOSF4_NETU As String
SXLOSF4_KKSET As String
SXLOSF4_CALCMAX As Double
SXLOSF4_CALCAVE As Double
SXLBMD_SMPPOS As Integer
SXLBMD1_KKSP As String
SXLBMD1_NETU As String
SXLBMD1_KKSET As String
SXLBMD1_CALCMAX As Double
SXLBMD1_CALCAVE As Double
SXLBMD1_CALCMIN As Double
SXLBMD1_CALCMB As Double
SXLBMD2_KKSP As String
SXLBMD2_NETU As String
SXLBMD2_KKSET As String
SXLBMD2_CALCMAX As Double
SXLBMD2_CALCAVE As Double
SXLBMD2_CALCMIN As Double
SXLBMD2_CALCMB As Double
SXLBMD3_KKSP As String
SXLBMD3_NETU As String
SXLBMD3_KKSET As String
SXLBMD3_CALCMAX As Double
SXLBMD3_CALCAVE As Double
SXLBMD3_CALCMIN As Double
SXLBMD3_CALCMB As Double
SXLGD_SMPPOS As Integer
SXLGD_MSRSDEN As Integer
SXLGD_MSRSLDL As Integer
SXLGD_MSRSDVD2 As Integer
SXLLT_SMPPOS As Integer
SXLLT_MEASPEAK As Integer
SXLLT_CALCMEAS As Integer
REGDATE As Date
SENDFLAG As String
SENDDATE As Date
SNDKDWH As String
SDAYDWH As Date
SNDKSPC As String
SDAYSPC As Date
End Type


Public Type typ_TBCMX012

BLOCKID As String
FROMTOKBN As String
STCID As String
hinban As String
REVNUM As String
factory As String
opecond As String
STCKNNUM As String
CRYNUM As String
SXLOSF1_SMPPOS As Integer
SXLOSF1_KKSP As String
SXLOSF1_NETU As String
SXLOSF1_KKSET As String
SXLOSF1_MEAS1 As Double
SXLOSF1_MEAS2 As Double
SXLOSF1_MEAS3 As Double
SXLOSF1_MEAS4 As Double
SXLOSF1_MEAS5 As Double
SXLOSF1_MEAS6 As Double
SXLOSF1_MEAS7 As Double
SXLOSF1_MEAS8 As Double
SXLOSF1_MEAS9 As Double
SXLOSF1_MEAS10 As Double
SXLOSF1_MEAS11 As Double
SXLOSF1_MEAS12 As Double
SXLOSF1_MEAS13 As Double
SXLOSF1_MEAS14 As Double
SXLOSF1_MEAS15 As Double
SXLOSF1_MEAS16 As Double
SXLOSF1_MEAS17 As Double
SXLOSF1_MEAS18 As Double
SXLOSF1_MEAS19 As Double
SXLOSF1_MEAS20 As Double
SXLOSF2_KKSP As String
SXLOSF2_NETU As String
SXLOSF2_KKSET As String
SXLOSF2_MEAS1 As Double
SXLOSF2_MEAS2 As Double
SXLOSF2_MEAS3 As Double
SXLOSF2_MEAS4 As Double
SXLOSF2_MEAS5 As Double
SXLOSF2_MEAS6 As Double
SXLOSF2_MEAS7 As Double
SXLOSF2_MEAS8 As Double
SXLOSF2_MEAS9 As Double
SXLOSF2_MEAS10 As Double
SXLOSF2_MEAS11 As Double
SXLOSF2_MEAS12 As Double
SXLOSF2_MEAS13 As Double
SXLOSF2_MEAS14 As Double
SXLOSF2_MEAS15 As Double
SXLOSF2_MEAS16 As Double
SXLOSF2_MEAS17 As Double
SXLOSF2_MEAS18 As Double
SXLOSF2_MEAS19 As Double
SXLOSF2_MEAS20 As Double
SXLOSF3_KKSP As String
SXLOSF3_NETU As String
SXLOSF3_KKSET As String
SXLOSF3_MEAS1 As Double
SXLOSF3_MEAS2 As Double
SXLOSF3_MEAS3 As Double
SXLOSF3_MEAS4 As Double
SXLOSF3_MEAS5 As Double
SXLOSF3_MEAS6 As Double
SXLOSF3_MEAS7 As Double
SXLOSF3_MEAS8 As Double
SXLOSF3_MEAS9 As Double
SXLOSF3_MEAS10 As Double
SXLOSF3_MEAS11 As Double
SXLOSF3_MEAS12 As Double
SXLOSF3_MEAS13 As Double
SXLOSF3_MEAS14 As Double
SXLOSF3_MEAS15 As Double
SXLOSF3_MEAS16 As Double
SXLOSF3_MEAS17 As Double
SXLOSF3_MEAS18 As Double
SXLOSF3_MEAS19 As Double
SXLOSF3_MEAS20 As Double
SXLOSF4_KKSP As String
SXLOSF4_NETU As String
SXLOSF4_KKSET As String
SXLOSF4_MEAS1  As Double
SXLOSF4_MEAS2 As Double
SXLOSF4_MEAS3 As Double
SXLOSF4_MEAS4 As Double
SXLOSF4_MEAS5 As Double
SXLOSF4_MEAS6 As Double
SXLOSF4_MEAS7 As Double
SXLOSF4_MEAS8 As Double
SXLOSF4_MEAS9 As Double
SXLOSF4_MEAS10 As Double
SXLOSF4_MEAS11 As Double
SXLOSF4_MEAS12 As Double
SXLOSF4_MEAS13 As Double
SXLOSF4_MEAS14 As Double
SXLOSF4_MEAS15 As Double
SXLOSF4_MEAS16 As Double
SXLOSF4_MEAS17 As Double
SXLOSF4_MEAS18 As Double
SXLOSF4_MEAS19 As Double
SXLOSF4_MEAS20 As Double
SXLBMD_SMPPOS As Integer
SXLBMD1_KKSP As String
SXLBMD1_NETU As String
SXLBMD1_KKSET As String
SXLBMD1_MEAS1 As Double
SXLBMD1_MEAS2 As Double
SXLBMD1_MEAS3 As Double
SXLBMD1_MEAS4 As Double
SXLBMD1_MEAS5 As Double
SXLBMD2_KKSP As String
SXLBMD2_NETU As String
SXLBMD2_KKSET As String
SXLBMD2_MEAS1 As Double
SXLBMD2_MEAS2 As Double
SXLBMD2_MEAS3 As Double
SXLBMD2_MEAS4 As Double
SXLBMD2_MEAS5 As Double
SXLBMD3_KKSP As String
SXLBMD3_NETU As String
SXLBMD3_KKSET As String
SXLBMD3_MEAS1 As Double
SXLBMD3_MEAS2 As Double
SXLBMD3_MEAS3 As Double
SXLBMD3_MEAS4 As Double
SXLBMD3_MEAS5 As Double
SXLT_SMPPOS As Integer
SXLLT_MEASPEAK As Integer
SXLLT_MEAS1 As Integer
SXLLT_MEAS2 As Integer
SXLLT_MEAS3 As Integer
SXLLT_MEAS4 As Integer
SXLLT_MEAS5 As Integer
REGDATE As Date
SENDFLAG As String
SENDDATE As Date
SNDKDWH As String
SDAYDWH As Date
SNDKSPC As String
SDAYSPC As Date
End Type

Public Type typ_TBCMX013
BLOCKID As String
FROMTOKBN As String
STCID As String
hinban As String
REVNUM As String
factory As String
opecond As String
STCKNNUM As String
CRYNUM As String
SXLGD_SMPPOS As Integer
SXLGD_MS01DEN1 As Integer
SXLGD_MS02DEN1 As Integer
SXLGD_MS03DEN1 As Integer
SXLGD_MS04DEN1 As Integer
SXLGD_MS05DEN1 As Integer
SXLGD_MS06DEN1 As Integer
SXLGD_MS07DEN1 As Integer
SXLGD_MS08DEN1 As Integer
SXLGD_MS09DEN1 As Integer
SXLGD_MS10DEN1 As Integer
SXLGD_MS11DEN1 As Integer
SXLGD_MS12DEN1 As Integer
SXLGD_MS13DEN1 As Integer
SXLGD_MS14DEN1 As Integer
SXLGD_MS15DEN1 As Integer
SXLGD_MS01DEN2 As Integer
SXLGD_MS02DEN2 As Integer
SXLGD_MS03DEN2 As Integer
SXLGD_MS04DEN2 As Integer
SXLGD_MS05DEN2 As Integer
SXLGD_MS06DEN2 As Integer
SXLGD_MS07DEN2 As Integer
SXLGD_MS08DEN2 As Integer
SXLGD_MS09DEN2 As Integer
SXLGD_MS10DEN2 As Integer
SXLGD_MS11DEN2 As Integer
SXLGD_MS12DEN2 As Integer
SXLGD_MS13DEN2 As Integer
SXLGD_MS14DEN2 As Integer
SXLGD_MS15DEN2 As Integer
SXLGD_MS01DEN3 As Integer
SXLGD_MS02DEN3 As Integer
SXLGD_MS03DEN3 As Integer
SXLGD_MS04DEN3 As Integer
SXLGD_MS05DEN3 As Integer
SXLGD_MS06DEN3 As Integer
SXLGD_MS07DEN3 As Integer
SXLGD_MS08DEN3 As Integer
SXLGD_MS09DEN3 As Integer
SXLGD_MS10DEN3 As Integer
SXLGD_MS11DEN3 As Integer
SXLGD_MS12DEN3 As Integer
SXLGD_MS13DEN3 As Integer
SXLGD_MS14DEN3 As Integer
SXLGD_MS15DEN3 As Integer
SXLGD_MS01DEN4 As Integer
SXLGD_MS02DEN4 As Integer
SXLGD_MS03DEN4 As Integer
SXLGD_MS04DEN4 As Integer
SXLGD_MS05DEN4 As Integer
SXLGD_MS06DEN4 As Integer
SXLGD_MS07DEN4 As Integer
SXLGD_MS08DEN4 As Integer
SXLGD_MS09DEN4 As Integer
SXLGD_MS10DEN4 As Integer
SXLGD_MS11DEN4 As Integer
SXLGD_MS12DEN4 As Integer
SXLGD_MS13DEN4 As Integer
SXLGD_MS14DEN4 As Integer
SXLGD_MS15DEN4 As Integer
SXLGD_MS01DEN5 As Integer
SXLGD_MS02DEN5 As Integer
SXLGD_MS03DEN5 As Integer
SXLGD_MS04DEN5 As Integer
SXLGD_MS05DEN5 As Integer
SXLGD_MS06DEN5 As Integer
SXLGD_MS07DEN5 As Integer
SXLGD_MS08DEN5 As Integer
SXLGD_MS09DEN5 As Integer
SXLGD_MS10DEN5 As Integer
SXLGD_MS11DEN5 As Integer
SXLGD_MS12DEN5 As Integer
SXLGD_MS13DEN5 As Integer
SXLGD_MS14DEN5 As Integer
SXLGD_MS15DEN5 As Integer
SXLGD_MS01LDL1 As Integer
SXLGD_MS02LDL1 As Integer
SXLGD_MS03LDL1 As Integer
SXLGD_MS04LDL1 As Integer
SXLGD_MS05LDL1 As Integer
SXLGD_MS06LDL1 As Integer
SXLGD_MS07LDL1 As Integer
SXLGD_MS08LDL1 As Integer
SXLGD_MS09LDL1 As Integer
SXLGD_MS10LDL1 As Integer
SXLGD_MS11LDL1 As Integer
SXLGD_MS12LDL1 As Integer
SXLGD_MS13LDL1 As Integer
SXLGD_MS14LDL1 As Integer
SXLGD_MS15LDL1 As Integer
SXLGD_MS01LDL2 As Integer
SXLGD_MS02LDL2 As Integer
SXLGD_MS03LDL2 As Integer
SXLGD_MS04LDL2 As Integer
SXLGD_MS05LDL2 As Integer
SXLGD_MS06LDL2 As Integer
SXLGD_MS07LDL2 As Integer
SXLGD_MS08LDL2 As Integer
SXLGD_MS09LDL2 As Integer
SXLGD_MS10LDL2 As Integer
SXLGD_MS11LDL2 As Integer
SXLGD_MS12LDL2 As Integer
SXLGD_MS13LDL2 As Integer
SXLGD_MS14LDL2 As Integer
SXLGD_MS15LDL2 As Integer
SXLGD_MS01LDL3 As Integer
SXLGD_MS02LDL3 As Integer
SXLGD_MS03LDL3 As Integer
SXLGD_MS04LDL3 As Integer
SXLGD_MS05LDL3 As Integer
SXLGD_MS06LDL3 As Integer
SXLGD_MS07LDL3 As Integer
SXLGD_MS08LDL3 As Integer
SXLGD_MS09LDL3 As Integer
SXLGD_MS10LDL3 As Integer
SXLGD_MS11LDL3 As Integer
SXLGD_MS12LDL3 As Integer
SXLGD_MS13LDL3 As Integer
SXLGD_MS14LDL3 As Integer
SXLGD_MS15LDL3 As Integer
SXLGD_MS01LDL4 As Integer
SXLGD_MS02LDL4 As Integer
SXLGD_MS03LDL4 As Integer
SXLGD_MS04LDL4 As Integer
SXLGD_MS05LDL4 As Integer
SXLGD_MS06LDL4 As Integer
SXLGD_MS07LDL4 As Integer
SXLGD_MS08LDL4 As Integer
SXLGD_MS09LDL4 As Integer
SXLGD_MS10LDL4 As Integer
SXLGD_MS11LDL4 As Integer
SXLGD_MS12LDL4 As Integer
SXLGD_MS13LDL4 As Integer
SXLGD_MS14LDL4 As Integer
SXLGD_MS15LDL4 As Integer
SXLGD_MS01LDL5 As Integer
SXLGD_MS02LDL5 As Integer
SXLGD_MS03LDL5 As Integer
SXLGD_MS04LDL5 As Integer
SXLGD_MS05LDL5 As Integer
SXLGD_MS06LDL5 As Integer
SXLGD_MS07LDL5 As Integer
SXLGD_MS08LDL5 As Integer
SXLGD_MS09LDL5 As Integer
SXLGD_MS10LDL5 As Integer
SXLGD_MS11LDL5 As Integer
SXLGD_MS12LDL5 As Integer
SXLGD_MS13LDL5 As Integer
SXLGD_MS14LDL5 As Integer
SXLGD_MS15LDL5 As Integer
SXLGD_MS01DVD21 As Integer
SXLGD_MS01DVD22 As Integer
SXLGD_MS01DVD23 As Integer
SXLGD_MS01DVD24 As Integer
SXLGD_MS01DVD25 As Integer
REGDATE As Date
SENDFLAG As String
SENDDATE As Date
SNDKDWH As String
SDAYDWH As Date
SNDKSPC As String
SDAYSPC As Date
End Type

Public recX011() As typ_TBCMX011
Public recX012() As typ_TBCMX012
Public recX013() As typ_TBCMX013




'概要      :結晶最終払出入力 表示用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型                   ,説明
'      　　:BlockID_in　 ,I  ,String               ,ブロックID
'      　　:blkInfo　　　,O  ,typ_cmkc001f_Block   ,ブロック情報
'      　　:records　　　,O  ,typ_cmkc001f_Disp    ,製品仕様取得用
'      　　:戻り値       ,O  ,FUNCTION_RETURN      ,読み込みの成否
Public Function DBDRV_fcmkc001f_Disp(BlockID_in As String, blkInfo As typ_cmkc001f_Block, records() As typ_cmkc001f_Disp) As FUNCTION_RETURN
    Dim sql As String
    Dim rs As OraDynaset
    Dim recCnt As Integer
    Dim i As Long
    Dim n As Integer
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function DBDRV_fcmkc001f_Disp"
    
    DBDRV_fcmkc001f_Disp = FUNCTION_RETURN_SUCCESS
    
    ''ブロック情報を得る
    sql = "Select BLK.INGOTPOS, BLK.LENGTH, BLK.REALLEN, BLK.KRPROCCD, BLK.NOWPROC, BLK.LPKRPROCCD, " & _
          "BLK.LASTPASS, BLK.DELCLS, BLK.RSTATCLS, BLK.LSTATCLS, CRY.SEED " & _
          "From TBCME040 BLK, TBCME037 CRY " & _
          "Where (BLOCKID='" & BlockID_in & "') and (BLK.CRYNUM=CRY.CRYNUM)"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        DBDRV_fcmkc001f_Disp = FUNCTION_RETURN_FAILURE
        rs.Close
        GoTo proc_exit
    End If
    With blkInfo
        .INGOTPOS = rs("INGOTPOS")          ' 結晶内開始位置
        .LENGTH = rs("LENGTH")              ' 長さ
        .REALLEN = rs("REALLEN")            ' 実長さ
        .KRPROCCD = rs("KRPROCCD")          ' 現在管理工程
        .NOWPROC = rs("NOWPROC")            ' 現在工程
        .LPKRPROCCD = rs("LPKRPROCCD")      ' 最終通過管理工程
        .LASTPASS = rs("LASTPASS")          ' 最終通過工程
        .DELCLS = rs("DELCLS")              ' 削除区分
        .RSTATCLS = rs("RSTATCLS")          ' 流動状態区分
        .LSTATCLS = rs("LSTATCLS")          ' 最終状態区分
        .SEED = rs("SEED")                  ' SEED
    End With
    rs.Close
    
    
    
    ''製品仕様を得る
    sql = "select "
    sql = sql & "BH.E041HINBAN, "           ' 品番
    sql = sql & "BH.E041INGOTPOS, "         ' 結晶内開始位置
    sql = sql & "BH.E041REVNUM, "           ' 製品番号改訂番号
    sql = sql & "BH.E041FACTORY, "          ' 工場
    sql = sql & "BH.E041OPECOND, "          ' 操業条件
    sql = sql & "BH.E041LENGTH, "           ' 長さ
    '製品仕様SXLデータ
    sql = sql & "S.E018HSXD1CEN, "          ' 品ＳＸ直径１中心
    sql = sql & "S.E018HSXRMIN, "           ' 品ＳＸ比抵抗下限
    sql = sql & "S.E018HSXRMAX, "           ' 品ＳＸ比抵抗上限
    sql = sql & "S.E018HSXRMBNP, "          ' 品ＳＸ比抵抗面内分布
    sql = sql & "S.E018HSXRHWYS, "          ' 品ＳＸ比抵抗保証方法＿処
    sql = sql & "S.E019HSXONMIN, "          ' 品ＳＸ酸素濃度下限
    sql = sql & "S.E019HSXONMAX, "          ' 品ＳＸ酸素濃度上限
    sql = sql & "S.E019HSXONMBP, "          ' 品ＳＸ酸素濃度面内分布
    sql = sql & "S.E019HSXONHWS, "          ' 品ＳＸ酸素濃度保証方法＿処
    sql = sql & "S.E019HSXCNMIN, "          ' 品ＳＸ炭素濃度下限
    sql = sql & "S.E019HSXCNMAX, "          ' 品ＳＸ炭素濃度上限
    sql = sql & "S.E019HSXCNHWS, "          ' 品ＳＸ炭素濃度保証方法＿処
    sql = sql & "S.E019HSXTMMAXN, "         ' 品ＳＸ転位密度上限        項目追加，修正対応 2003.05.20 yakimura
    sql = sql & "S.E020HSXBM1AN, "          ' 品ＳＸＢＭＤ１平均下限
    sql = sql & "S.E020HSXBM1AX, "          ' 品ＳＸＢＭＤ１平均上限
    sql = sql & "S.E020HSXBM1HS, "          ' 品ＳＸＢＭＤ１保証方法＿処
    sql = sql & "S.E020HSXBM2AN, "          ' 品ＳＸＢＭＤ２平均下限
    sql = sql & "S.E020HSXBM2AX, "          ' 品ＳＸＢＭＤ２平均上限
    sql = sql & "S.E020HSXBM2HS, "          ' 品ＳＸＢＭＤ２保証方法＿処
    sql = sql & "S.E020HSXBM3AN, "          ' 品ＳＸＢＭＤ３平均下限
    sql = sql & "S.E020HSXBM3AX, "          ' 品ＳＸＢＭＤ３平均上限
    sql = sql & "S.E020HSXBM3HS, "          ' 品ＳＸＢＭＤ３保証方法＿処
    sql = sql & "S.E020HSXOF1AX, "          ' 品ＳＸＯＳＦ１平均上限
    sql = sql & "S.E020HSXOF1MX, "          ' 品ＳＸＯＳＦ１上限
    sql = sql & "S.E020HSXOF1HS, "          ' 品ＳＸＯＳＦ１ 保証方法＿処
    sql = sql & "S.E020HSXOF2AX, "          ' 品ＳＸＯＳＦ２平均上限
    sql = sql & "S.E020HSXOF2MX, "          ' 品ＳＸＯＳＦ２上限
    sql = sql & "S.E020HSXOF2HS, "          ' 品ＳＸＯＳＦ２ 保証方法＿処
    sql = sql & "S.E020HSXOF3AX, "          ' 品ＳＸＯＳＦ３平均上限
    sql = sql & "S.E020HSXOF3MX, "          ' 品ＳＸＯＳＦ３上限
    sql = sql & "S.E020HSXOF3HS, "          ' 品ＳＸＯＳＦ３ 保証方法＿処
    sql = sql & "S.E020HSXOF4AX, "          ' 品ＳＸＯＳＦ４平均上限
    sql = sql & "S.E020HSXOF4MX, "          ' 品ＳＸＯＳＦ４上限
    sql = sql & "S.E020HSXOF4HS, "          ' 品ＳＸＯＳＦ４ 保証方法＿処
    sql = sql & "S.E020HSXDENMX, "          ' 品ＳＸＤｅｎ上限
    sql = sql & "S.E020HSXDENMN, "          ' 品ＳＸＤｅｎ下限
    sql = sql & "S.E020HSXDENHS, "          ' 品ＳＸＤｅｎ保証方法＿処
    sql = sql & "S.E020HSXDVDMXN, "         ' 品ＳＸＤＶＤ２上限       項目追加，修正対応 2003.05.20 yakimura
    sql = sql & "S.E020HSXDVDMNN, "         ' 品ＳＸＤＶＤ２下限       項目追加，修正対応 2003.05.20 yakimura
    sql = sql & "S.E020HSXDVDHS, "          ' 品ＳＸＤＶＤ２保証方法＿処
    sql = sql & "S.E020HSXLDLMX, "          ' 品ＳＸＬ／ＤＬ上限
    sql = sql & "S.E020HSXLDLMN, "          ' 品ＳＸＬ／ＤＬ下限
    sql = sql & "S.E020HSXLDLHS, "          ' 品ＳＸＬ／ＤＬ保証方法＿処
    sql = sql & "S.E019HSXLTMIN, "          ' 品ＳＸＬタイム下限
    sql = sql & "S.E019HSXLTMAX, "          ' 品ＳＸＬタイム上限
    sql = sql & "S.E019HSXLTHWS, "          ' 品ＳＸＬタイム保証方法＿処
    sql = sql & "S.E018HSXDPDIR, "          ' 品ＳＸ溝位置方位
    sql = sql & "S.E018HSXDPDRC, "          ' 品ＳＸ溝位置方向
    sql = sql & "S.E018HSXDWMIN, "          ' 品ＳＸ溝巾下限
    sql = sql & "S.E018HSXDWMAX, "          ' 品ＳＸ溝巾上限
    sql = sql & "S.E018HSXDDMIN, "          ' 品ＳＸ溝深下限
    sql = sql & "S.E018HSXDDMAX, "          ' 品ＳＸ溝深上限
    sql = sql & "S.E018HSXD1MIN, "          ' 品ＳＸ直径１下限
    sql = sql & "S.E018HSXD1MAX, "          ' 品ＳＸ直径１上限
    sql = sql & "S.E018HSXCTCEN, "          ' 品ＳＸ結晶面傾縦中心
    sql = sql & "S.E018HSXCYCEN, "          ' 品ＳＸ結晶面傾横中心
    sql = sql & "U.EPDUP "                  ' 結晶内側管理 EPD　上限
    sql = sql & " from VECME009 BH, VECME001 S, TBCME036 U "
    sql = sql & " where BH.E040BLOCKID='" & BlockID_in & "' "
    sql = sql & " and S.E018HINBAN=BH.E041HINBAN "
    sql = sql & " and S.E018MNOREVNO=BH.E041REVNUM "
    sql = sql & " and S.E018FACTORY=BH.E041FACTORY "
    sql = sql & " and S.E018OPECOND=BH.E041OPECOND "
    sql = sql & " and U.HINBAN=BH.E041HINBAN "
    sql = sql & " and U.MNOREVNO=BH.E041REVNUM "
    sql = sql & " and U.FACTORY=BH.E041FACTORY "
    sql = sql & " and U.OPECOND=BH.E041OPECOND "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount = 0 Then
        ReDim records(0)
        rs.Close
        GoTo proc_exit
    End If
    
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            '品番管理
            .hinban = rs("E041HINBAN")                          ' 品番
            .INGOTPOS = rs("E041INGOTPOS")                      ' 結晶内開始位置
            .REVNUM = rs("E041REVNUM")                          ' 製品番号改訂番号
            .factory = rs("E041FACTORY")                        ' 工場
            .opecond = rs("E041OPECOND")                        ' 操業条件
            .LENGTH = rs("E041LENGTH")                          ' 長さ
            '製品仕様SXLデータ
            .HSXD1CEN = fncNullCheck(rs("E018HSXD1CEN"))                      ' 品ＳＸ直径１中心
            .HSXRMIN = fncNullCheck(rs("E018HSXRMIN"))                        ' 品ＳＸ比抵抗下限
            .HSXRMAX = fncNullCheck(rs("E018HSXRMAX"))                        ' 品ＳＸ比抵抗上限
            .HSXRMBNP = fncNullCheck(rs("E018HSXRMBNP"))                      ' 品ＳＸ比抵抗面内分布
            .HSXRHWYS = rs("E018HSXRHWYS")                      ' 品ＳＸ比抵抗保証方法＿処
            .HSXONMIN = fncNullCheck(rs("E019HSXONMIN"))                      ' 品ＳＸ酸素濃度下限  'NULL対応
            .HSXONMAX = fncNullCheck(rs("E019HSXONMAX"))                      ' 品ＳＸ酸素濃度上限
            .HSXONMBP = fncNullCheck(rs("E019HSXONMBP"))                      ' 品ＳＸ酸素濃度面内分布  'NULL対応
            .HSXONHWS = rs("E019HSXONHWS")                      ' 品ＳＸ酸素濃度保証方法＿処
            .HSXCNMIN = fncNullCheck(rs("E019HSXCNMIN"))                      ' 品ＳＸ炭素濃度下限  'NULL対応
            .HSXCNMAX = fncNullCheck(rs("E019HSXCNMAX"))                      ' 品ＳＸ炭素濃度上限  'NULL対応
            .HSXCNHWS = rs("E019HSXCNHWS")                      ' 品ＳＸ炭素濃度保証方法＿処
            .HSXTMMAX = rs("E019HSXTMMAXN")                     ' 品ＳＸ転位密度上限       項目追加，修正対応 2003.05.20 yakimura
            For n = 1 To 3 'NULL対応
                If IsNull(rs("E020HSXBM" & n & "AN")) = False Then
                    .HSXBMnAN(n) = rs("E020HSXBM" & n & "AN") * 10  ' 品ＳＸＢＭＤn 平均下限
                Else
                    .HSXBMnAN(n) = -1
                End If
                
                If IsNull(rs("E020HSXBM" & n & "AX")) = False Then
                    .HSXBMnAX(n) = rs("E020HSXBM" & n & "AX") * 10 ' 品ＳＸＢＭＤn 平均上限
                Else
                    .HSXBMnAX(n) = -1
                End If
                .HSXBMnHS(n) = rs("E020HSXBM" & n & "HS")       ' 品ＳＸＢＭＤn 保証方法＿処
            Next
            For n = 1 To 4
                .HSXOFnAX(n) = fncNullCheck(rs("E020HSXOF" & n & "AX"))       ' 品ＳＸＯＳＦn 平均上限  'NULL対応
                .HSXOFnMX(n) = fncNullCheck(rs("E020HSXOF" & n & "MX"))       ' 品ＳＸＯＳＦn 上限      'NULL対応
                .HSXOFnHS(n) = rs("E020HSXOF" & n & "HS")       ' 品ＳＸＯＳＦn 保証方法＿処
            Next
            .HSXDENMX = fncNullCheck(rs("E020HSXDENMX"))                      ' 品ＳＸＤｅｎ上限    'NULL対応
            .HSXDENMN = fncNullCheck(rs("E020HSXDENMN"))                      ' 品ＳＸＤｅｎ下限    'NULL対応
            .HSXDENHS = rs("E020HSXDENHS")                      ' 品ＳＸＤｅｎ保証方法＿処
            .HSXDVDMX = fncNullCheck(rs("E020HSXDVDMXN"))                     ' 品ＳＸＤＶＤ２上限      項目追加，修正対応 2003.05.20 yakimura 'NULL対応
            .HSXDVDMN = fncNullCheck(rs("E020HSXDVDMNN"))                     ' 品ＳＸＤＶＤ２下限      項目追加，修正対応 2003.05.20 yakimura  'NULL対応
            .HSXDVDHS = rs("E020HSXDVDHS")                      ' 品ＳＸＤＶＤ２保証方法＿処
            .HSXLDLMX = fncNullCheck(rs("E020HSXLDLMX"))                      ' 品ＳＸＬ／ＤＬ上限  'NULL対応
            .HSXLDLMN = fncNullCheck(rs("E020HSXLDLMN"))                      ' 品ＳＸＬ／ＤＬ下限  'NULL対応
            .HSXLDLHS = rs("E020HSXLDLHS")                      ' 品ＳＸＬ／ＤＬ保証方法＿処
            .HSXLTMIN = fncNullCheck(rs("E019HSXLTMIN"))                      ' 品ＳＸＬタイム下限  'NULL対応
            .HSXLTMAX = fncNullCheck(rs("E019HSXLTMAX"))                      ' 品ＳＸＬタイム上限  'NULL対応
            .HSXLTHWS = rs("E019HSXLTHWS")                      ' 品ＳＸＬタイム保証方法＿処
            .HSXDPDIR = rs("E018HSXDPDIR")                      ' 品ＳＸ溝位置方位
            .HSXDPDRC = rs("E018HSXDPDRC")                      ' 品ＳＸ溝位置方向
            .HSXDWMIN = fncNullCheck(rs("E018HSXDWMIN"))                      ' 品ＳＸ溝巾下限  'NULL対応
            .HSXDWMAX = fncNullCheck(rs("E018HSXDWMAX"))                      ' 品ＳＸ溝巾上限  'NULL対応
            .HSXDDMIN = fncNullCheck(rs("E018HSXDDMIN"))                      ' 品ＳＸ溝深下限  'NULL対応
            .HSXDDMAX = fncNullCheck(rs("E018HSXDDMAX"))                      ' 品ＳＸ溝深上限  'NULL対応
            .HSXD1MIN = fncNullCheck(rs("E018HSXD1MIN"))                      ' 品ＳＸ直径１下限    'NULL対応
            .HSXD1MAX = fncNullCheck(rs("E018HSXD1MAX"))                      ' 品ＳＸ直径１上限    'NULL対応
            .HSXCTCEN = fncNullCheck(rs("E018HSXCTCEN"))                      ' 品ＳＸ結晶面傾縦中心    'NULL対応
            .HSXCYCEN = fncNullCheck(rs("E018HSXCYCEN"))                      ' 品ＳＸ結晶面傾横中心    'NULL対応
            .EPDUP = fncNullCheck(rs("EPDUP"))                                ' 結晶内側管理 EPD　上限  'NULL対応
        End With
        rs.MoveNext
    Next
    rs.Close


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_fcmkc001f_Disp = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function


'結晶最終検査への挿入（内部関数）
Private Function fcmkc001f_ExecFts(CryIn As typ_cmkc001f_ExecCryIn, CryFTest As typ_cmkc001f_ExecFts) As FUNCTION_RETURN
Dim sql As String
Dim n As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Sub fcmkc001f_ExecFts"

    fcmkc001f_ExecFts = FUNCTION_RETURN_SUCCESS
        
    '結晶最終検査への挿入
    With CryFTest
        sql = "insert into TBCMJ010 ( "
        sql = sql & "CRYNUM, "                      ' 結晶番号
        sql = sql & "INGOTPOS, "                    ' インゴット内位置
        sql = sql & "TRANCNT, "                     ' 処理回数
        sql = sql & "LENGTH, "                      ' 長さ
        sql = sql & "KRPROCCD, "                    ' 管理工程コード
        sql = sql & "PROCCODE, "                    ' 工程コード
        sql = sql & "PAYCLASS, "                    ' 払い出し区分
        sql = sql & "OUTLENGTH, "                   ' 出荷長さ
        For n = 1 To 5
            sql = sql & "PART" & n & ", "           ' 部位n
            sql = sql & "P" & n & "BDLEN, "         ' 部位n 不良長さ
            sql = sql & "P" & n & "BDCAUS, "        ' 部位n 不良理由
        Next
        sql = sql & "TSTAFFID, "                    ' 登録社員ID
        sql = sql & "REGDATE, "                     ' 登録日付
        sql = sql & "KSTAFFID, "                    ' 更新社員ID
        sql = sql & "UPDDATE, "                     ' 更新日付
        sql = sql & "SUMMITSENDFLAG, "              ' SUMMIT送信フラグ
        sql = sql & "SENDFLAG, "                    ' 送信フラグ
        sql = sql & "SENDDATE ) "                   ' 送信日付
        
        sql = sql & "select "
        sql = sql & " '" & CryIn.CRYNUM & "', "     ' 結晶番号
        sql = sql & CryIn.INGOTPOS & ", "           ' インゴット内位置
        sql = sql & "nvl(max(TRANCNT),0)+1, "       ' 処理回数
        sql = sql & .LENGTH & ", "                  ' 長さ
        sql = sql & " '" & .KRPROCCD & "', "        ' 管理工程コード
        sql = sql & " '" & .PROCCODE & "', "        ' 工程コード
        sql = sql & " '" & .PAYCLASS & "', "        ' 払い出し区分
        sql = sql & .OUTLENGTH & ", "               ' 出荷長さ
        
        For n = 1 To 5
            sql = sql & .PART(n) & ", "             ' 部位n
            sql = sql & .BDLEN(n) & ", "            ' 部位n 不良長さ
            sql = sql & " '" & .BDCAUS(n) & "', "   ' 部位n 不良理由
        Next
        sql = sql & " '" & .TSTAFFID & "', "        ' 登録社員ID
        sql = sql & "sysdate, "                     ' 登録日付
        sql = sql & " '" & .TSTAFFID & "', "        ' 更新社員ID
        sql = sql & "sysdate, "                     ' 更新日付
        sql = sql & "'0', "                         ' SUMMIT送信フラグ
        sql = sql & "'0', "                         ' 送信フラグ
        sql = sql & "sysdate "                      ' 送信日付
        sql = sql & " From TBCMJ010 "
        sql = sql & " where CRYNUM='" & CryIn.CRYNUM & "' and INGOTPOS=" & CryIn.INGOTPOS
    End With
    
    If 0 >= OraDB.ExecuteSQL(sql) Then
        fcmkc001f_ExecFts = FUNCTION_RETURN_FAILURE
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    fcmkc001f_ExecFts = FUNCTION_RETURN_FAILURE
    Debug.Print "==== ERROR"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function


'クリスタルカタログ受入実績への挿入（内部関数）
Private Function fcmkc001f_ExecCatalog(CryIn As typ_cmkc001f_ExecCryIn, CryCatalog As typ_cmkc001f_ExecCatalog) As FUNCTION_RETURN
Dim sql As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function fcmkc001f_ExecCatalog"

    fcmkc001f_ExecCatalog = FUNCTION_RETURN_SUCCESS


    ' クリスタルカタログ受入実績への挿入
    sql = "insert into TBCMG007 ( "
    sql = sql & "CRYNUM, "            ' 結晶番号
    sql = sql & "TRANCNT, "           ' 処理回数
    sql = sql & "KRPROCCD, "          ' 管理工程コード
    sql = sql & "PROCCODE, "          ' 工程コード
    sql = sql & "BDCODE, "            ' 不良理由コード
    sql = sql & "PALTNUM, "           ' パレット番号
    sql = sql & "TSTAFFID, "          ' 登録社員ID
    sql = sql & "REGDATE, "           ' 登録日付
    sql = sql & "KSTAFFID, "          ' 更新社員ID
    sql = sql & "UPDDATE, "           ' 更新日付
    sql = sql & "SENDFLAG, "          ' 送信フラグ
    sql = sql & "SENDDATE) "          ' 送信日付

    With CryCatalog
        sql = sql & "Select "
        sql = sql & " '" & .CRYNUM & "', "                          ' 結晶番号
        sql = sql & "nvl(max(TRANCNT),0)+1, "                       ' 処理回数
        sql = sql & " '" & MGPRCD_KESSYOU_SAISYUU_HARAIDASI & "', " ' 管理工程コード
        sql = sql & " '" & PROCD_KESSYOU_SAISYUU_HARAIDASI & "', "  ' 工程コード
        sql = sql & " '" & .BDCODE & "', "                          ' 不良理由コード
        sql = sql & " '" & .PALTNUM & "', "                         ' パレット番号
        sql = sql & " '" & .TSTAFFID & "', "                        ' 登録社員ID
        sql = sql & "sysdate, "                                     ' 登録日付
        sql = sql & " '" & .TSTAFFID & "', "                        ' 更新社員ID
        sql = sql & "sysdate, "                                     ' 更新日付
        sql = sql & "'0', "                                         ' 送信フラグ
        sql = sql & "sysdate "                                      ' 送信日付
        sql = sql & "From TBCMG007 " & _
              "Where (CRYNUM='" & .CRYNUM & "')"
    End With

    If 0 >= OraDB.ExecuteSQL(sql) Then
        fcmkc001f_ExecCatalog = FUNCTION_RETURN_FAILURE
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    fcmkc001f_ExecCatalog = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function


' ブロック管理の更新（内部関数）
Private Function fcmkc001f_ExecBlock(CryIn As typ_cmkc001f_ExecCryIn, BlockMan As typ_cmkc001f_Block, Optional BDCAUS$ = vbNullString) As FUNCTION_RETURN
Dim sql As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function fcmkc001f_ExecBlock"

    fcmkc001f_ExecBlock = FUNCTION_RETURN_SUCCESS

    ' ブロック管理の更新
    With BlockMan
        sql = "update TBCME040 set "
        sql = sql & "REALLEN=" & .REALLEN & ", "            ' 実長さ
        sql = sql & "KRPROCCD='" & .KRPROCCD & "', "        ' 現在管理工程
        sql = sql & "NOWPROC='" & .NOWPROC & "', "          ' 現在工程
        sql = sql & "LPKRPROCCD='" & .LPKRPROCCD & "', "    ' 最終通過管理工程
        sql = sql & "LASTPASS='" & .LASTPASS & "', "        ' 最終通過工程
        sql = sql & "DELCLS='" & .DELCLS & "', "            ' 削除区分
        sql = sql & "RSTATCLS='" & .RSTATCLS & "', "        ' 流動状態区分
        sql = sql & "LSTATCLS='" & .LSTATCLS & "', "        ' 最終状態区分
        If BDCAUS <> vbNullString Then
            sql = sql & "BDCAUS='" & BDCAUS & "', "         ' 最終状態区分
        End If
        sql = sql & "UPDDATE=SYSDATE, "                     ' 更新日
        sql = sql & "SENDFLAG='0' "
        sql = sql & " where  "
        sql = sql & "CRYNUM='" & CryIn.CRYNUM & "' and INGOTPOS=" & CryIn.INGOTPOS
    End With
    
    If 0 >= OraDB.ExecuteSQL(sql) Then
        fcmkc001f_ExecBlock = FUNCTION_RETURN_FAILURE
    End If
    

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    fcmkc001f_ExecBlock = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function


'2008/06/01 SHINDOH--------------------------------------
'Public Function fcmkc001f_Exec_037(sCryNum As String) As FUNCTION_RETURN
Public Function fcmkc001f_Exec_037(sCrynum As String, sProcCd As String) As FUNCTION_RETURN
'2008/06/01 SHINDOH--------------------------------------
    Dim sDbName As String
    Dim sql     As String
    Dim sErrMsg As String
    
    
    '' 結晶情報の更新
    sDbName = "E037"
    sql = "update TBCME037 set "
    sql = sql & "KRPROCCD  ='" & MGPRCD_WFC_HARAIDASI & "', "
'2008/06/01 SHINDOH--------------------------------------
'    sql = sql & "PROCCD    ='" & PROCD_WFC_HARAIDASI & "', "
    sql = sql & "PROCCD    ='" & sProcCd & "', "
'2008/06/01 SHINDOH--------------------------------------
    sql = sql & "LPKRPROCCD='" & MGPRCD_KESSYOU_SAISYUU_HARAIDASI & "', "
    sql = sql & "LASTPASS  ='" & PROCD_KESSYOU_SAISYUU_HARAIDASI & "', "
    sql = sql & "UPDDATE   = sysdate, "
    sql = sql & "SENDFLAG  ='0'"
    sql = sql & " where CRYNUM='" & sCrynum & "'"
    '' WriteDBLog sql, sDbName
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        fcmkc001f_Exec_037 = FUNCTION_RETURN_FAILURE
    Else
        fcmkc001f_Exec_037 = FUNCTION_RETURN_SUCCESS
    End If
End Function


'実行時メイン
Public Function DBDRV_fcmkc001f_Exec( _
  CryIn As typ_cmkc001f_ExecCryIn, CryFTest As typ_cmkc001f_ExecFts, _
  BlockMan As typ_cmkc001f_Block, blkID As String, STAFFID As String) As FUNCTION_RETURN

Dim skipNukishi As Boolean
Dim sql$
Dim sqlWhere$
Dim INGOTPOS%
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001f_SQL.bas -- Function DBDRV_fcmkc001f_Exec"

    DBDRV_fcmkc001f_Exec = FUNCTION_RETURN_SUCCESS
    BlockMan.RSTATCLS = "T" '???

'ＷＦサンプル処理変更 2003.05.20 yakimura
'    '(P+/N+で)抜試指示工程を飛ばすかどうかのフラグを設定
'    If CryFTest.PROCCODE = PROCD_WFC_HARAIDASI Then
'        skipNukishi = True
'    Else
'        skipNukishi = False
'    End If
'ＷＦサンプル処理変更 2003.05.20 yakimura
    
    'ブロック管理テーブルの更新(なかったら挿入???)
    With BlockMan
        .KRPROCCD = CryFTest.KRPROCCD
        .NOWPROC = CryFTest.PROCCODE

'ＷＦサンプル処理変更 2003.05.20 yakimura
'        If skipNukishi Then
'            .LPKRPROCCD = MGPRCD_NUKISI_SIJI
'            .LASTPASS = PROCD_NUKISI_SIJI
'        Else
            .LPKRPROCCD = MGPRCD_KESSYOU_SAISYUU_HARAIDASI
            .LASTPASS = PROCD_KESSYOU_SAISYUU_HARAIDASI
'        End If
'ＷＦサンプル処理変更 2003.05.20 yakimura
        
        .REALLEN = CryFTest.OUTLENGTH
    End With
        
    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecBlock(CryIn, BlockMan) Then
        DBDRV_fcmkc001f_Exec = FUNCTION_RETURN_FAILURE
    End If
    
    '結晶最終検査へインサート
    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecFts(CryIn, CryFTest) Then
        DBDRV_fcmkc001f_Exec = FUNCTION_RETURN_FAILURE
    End If
    
'ＷＦサンプル処理変更 2003.05.20 yakimura
    '(P+/N+で)抜試指示工程を飛ばす場合の追加処理
'    If skipNukishi Then
'        sqlWhere = " where CRYNUM='" & CryIn.CRYNUM & "' and INGOTPOS=" & BlockMan.INGOTPOS
        
        'SXL管理を作る
'        sql = "insert into TBCME042 ("
'        sql = sql & "CRYNUM,INGOTPOS,LENGTH,SXLID,KRPROCCD,NOWPROC,LPKRPROCCD,LASTPASS,DELCLS,LSTATCLS,HOLDCLS"
'        sql = sql & ",HINBAN,REVNUM,FACTORY,OPECOND,BDCAUS,COUNT"
'        sql = sql & ",REGDATE,UPDDATE,SUMMITSENDFLAG,SENDFLAG,SENDDATE,PASSFLAG"
'        sql = sql & ") select"
'        sql = sql & " CRYNUM, HINFROM, HINTO-HINFROM"
'        sql = sql & ", substr(BLOCKID,1,10) || substr('0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ',HINFROM/100+1,1) || to_char(mod(HINFROM,100),'FM00') as SXLID"
'        sql = sql & ", ' ', 'CC720', ' ', 'CC710', '0', 'T', '0'"
'        sql = sql & ", HINBAN, REVNUM, FACTORY, OPECOND"
'        sql = sql & ", ' ', 0, sysdate, sysdate, '0', '0', sysdate, ' ' "
'        sql = sql & "from"
'        sql = sql & "("
'        sql = sql & " select BLK.CRYNUM, BLK.BLOCKID, HIN.HINBAN, HIN.REVNUM, HIN.FACTORY, HIN.OPECOND"
'        sql = sql & " , greatest(BLK.INGOTPOS,HIN.INGOTPOS) as HINFROM"
'        sql = sql & " , least(BLK.INGOTPOS+BLK.LENGTH,HIN.INGOTPOS+HIN.LENGTH) as HINTO"
'        sql = sql & " from TBCME041 HIN, TBCME040 BLK"
'        sql = sql & " where BLK.CRYNUM='" & CryIn.CRYNUM & "' and BLK.INGOTPOS=" & BlockMan.INGOTPOS
'        sql = sql & "  and HIN.CRYNUM=BLK.CRYNUM"
'        sql = sql & "  and HIN.INGOTPOS<BLK.INGOTPOS+BLK.LENGTH"
'        sql = sql & "  and HIN.INGOTPOS+HIN.LENGTH>BLK.INGOTPOS"
'        sql = sql & ") HINS"
'        If (OraDB.ExecuteSQL(sql) < 1) Then
'            DBDRV_fcmkc001f_Exec = FUNCTION_RETURN_FAILURE
'            GoTo proc_exit
'        End If
        
'        '抜試指示実績を作る
'        sql = "insert into TBCMW001 ("
'        sql = sql & " CRYNUM, INGOTPOS, TRANCNT, CRYLEN, KRPROCCD, PROCCODE,"
'        sql = sql & " BLOCKID, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE"
'        sql = sql & ") values ("
'        sql = sql & "'" & CryIn.CRYNUM & "', " & BlockMan.INGOTPOS & ","
'        sql = sql & " (select nvl(max(trancnt),0)+1 from TBCMW001" & sqlWhere & "), "
'        sql = sql & BlockMan.LENGTH & ","
'        sql = sql & " '" & MGPRCD_KESSYOU_SAISYUU_HARAIDASI & "', '" & PROCD_KESSYOU_SAISYUU_HARAIDASI & "','"
'        sql = sql & blkID & "', "
'        sql = sql & "'" & STAFFID & "', sysdate, ' ', sysdate, '0', sysdate)"
'        If (OraDB.ExecuteSQL(sql) < 1) Then
'            DBDRV_fcmkc001f_Exec = FUNCTION_RETURN_FAILURE
'            GoTo proc_exit
'        End If
'    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_fcmkc001f_Exec = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'BAR出荷のとき
Public Function DBDRV_fcmkc001f_ExecBar(CryIn As typ_cmkc001f_ExecCryIn, CryFTest As typ_cmkc001f_ExecFts, BlockMan As typ_cmkc001f_Block) As FUNCTION_RETURN

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001f_SQL.bas -- Function DBDRV_fcmkc001f_ExecBar"

    DBDRV_fcmkc001f_ExecBar = FUNCTION_RETURN_SUCCESS
'OraDB.BeginTrans
    'ブロック管理テーブルの更新(なかったら挿入???)
    With BlockMan
        .DELCLS = "1"
        .LSTATCLS = "B"
        .KRPROCCD = "     "
        .NOWPROC = "CC705"
        .LPKRPROCCD = MGPRCD_KESSYOU_SAISYUU_HARAIDASI
        .LASTPASS = PROCD_KESSYOU_SAISYUU_HARAIDASI
        .REALLEN = CryFTest.OUTLENGTH
    End With
    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecBlock(CryIn, BlockMan) Then
        DBDRV_fcmkc001f_ExecBar = FUNCTION_RETURN_FAILURE
    End If
    
    '結晶最終検査へインサート
    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecFts(CryIn, CryFTest) Then
        DBDRV_fcmkc001f_ExecBar = FUNCTION_RETURN_FAILURE
    End If
    
'OraDB.CommitTrans
    
    'ブロック管理テーブルの更新(なかったら挿入???)(ブロック管理???)
'    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecBlock(CryIn, BlockMan) Then
'        DBDRV_fcmkc001f_ExecBar = FUNCTION_RETURN_FAILURE
'    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_fcmkc001f_ExecBar = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'クリスタルカタログ格下
Public Function DBDRV_fcmkc001f_ExecCatalog(CryIn As typ_cmkc001f_ExecCryIn, CryFTest As typ_cmkc001f_ExecFts, _
                                           BlockMan As typ_cmkc001f_Block, CryCatalog As typ_cmkc001f_ExecCatalog) As FUNCTION_RETURN
Dim HIN As tFullHinban

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001f_SQL.bas -- Function DBDRV_fcmkc001f_ExecCatalog"

    DBDRV_fcmkc001f_ExecCatalog = FUNCTION_RETURN_SUCCESS
    CryFTest.PAYCLASS = "2"
    BlockMan.RSTATCLS = "G" '???
    BlockMan.NOWPROC = PROCD_KAKUAGE
    BlockMan.KRPROCCD = MGPRCD_KAKUAGE

    'ブロック管理テーブルの更新(なかったら挿入???)
    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecBlock(CryIn, BlockMan, CryCatalog.BDCODE) Then
        DBDRV_fcmkc001f_ExecCatalog = FUNCTION_RETURN_FAILURE
    End If

    '品番管理テーブルの更新
    With HIN
        .hinban = "G"
        .mnorevno = 0
        .factory = " "
        .opecond = " "
    End With
    With BlockMan
        If ChangeAreaHinban(CryIn.CRYNUM, .INGOTPOS, .LENGTH, HIN) = FUNCTION_RETURN_FAILURE Then
            DBDRV_fcmkc001f_ExecCatalog = FUNCTION_RETURN_FAILURE
        End If
    End With

    '結晶最終検査へインサート
    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecFts(CryIn, CryFTest) Then
        DBDRV_fcmkc001f_ExecCatalog = FUNCTION_RETURN_FAILURE
    End If

    'クリスタルカタログ受入実績へインサート
    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecCatalog(CryIn, CryCatalog) Then
        DBDRV_fcmkc001f_ExecCatalog = FUNCTION_RETURN_FAILURE
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_fcmkc001f_ExecCatalog = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'リメルト
Public Function DBDRV_fcmkc001f_ExecRemelt(CryIn As typ_cmkc001f_ExecCryIn, CryFTest As typ_cmkc001f_ExecFts, _
                                           BlockMan As typ_cmkc001f_Block) As FUNCTION_RETURN
Dim HIN As tFullHinban

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001f_SQL.bas -- Function DBDRV_fcmkc001f_ExecRemelt"

    DBDRV_fcmkc001f_ExecRemelt = FUNCTION_RETURN_SUCCESS
    CryFTest.PAYCLASS = "3" '???
    BlockMan.RSTATCLS = "M"
    BlockMan.NOWPROC = PROCD_RIMERUTO_UKEIRE
    BlockMan.KRPROCCD = MGPRCD_RIMERUTO_UKEIRE

    'ブロック管理テーブルの更新(なかったら挿入???)
    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecBlock(CryIn, BlockMan) Then
        DBDRV_fcmkc001f_ExecRemelt = FUNCTION_RETURN_FAILURE
    End If
    
    '品番管理テーブルの更新
    With HIN
        .hinban = "Z"
        .mnorevno = 0
        .factory = " "
        .opecond = " "
    End With
    With BlockMan
        If ChangeAreaHinban(CryIn.CRYNUM, .INGOTPOS, .LENGTH, HIN) = FUNCTION_RETURN_FAILURE Then
            DBDRV_fcmkc001f_ExecRemelt = FUNCTION_RETURN_FAILURE
        End If
    End With

    '結晶最終検査へインサート
    If FUNCTION_RETURN_FAILURE = fcmkc001f_ExecFts(CryIn, CryFTest) Then
        DBDRV_fcmkc001f_ExecRemelt = FUNCTION_RETURN_FAILURE
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_fcmkc001f_ExecRemelt = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


''''''(2002/07 s_cmzcF_cmhc001d_SQL.basより移動)
'''''Private Function AreaStr(cnd$, v1, v2, fmt$) As String
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo PROC_ERR
'''''    gErr.Push "s_cmzcF_cmhc001d_SQL.bas -- Function AreaStr"
'''''
'''''    If Trim$(cnd) = vbNullString Then                           ''保証方法が空欄なら規格なし
'''''        AreaStr = vbNullString
'''''    Else
'''''        AreaStr = Format$(v1, fmt) & " - " & Format$(v2, fmt)   ''指定の書式で範囲文字列を作成
'''''    End If
'''''
'''''PROC_EXIT:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''PROC_ERR:
'''''    'エラーハンドラ
'''''    gErr.HandleError
'''''    Resume PROC_EXIT
'''''End Function


'''''Public Function cmkc001b_DBDataCheck1(LWD() As cmkc001b_LockWait, Wd1() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
''''''    Dim typ_A As typ_AllTypes        '全情報構造体
''''''    Dim c0 As Integer
''''''    Dim sErrMsg As String
''''''    Dim NothingFlag As Boolean
''''''    Dim FuncAns As FUNCTION_RETURN
''''''    For c0 = 1 To UBound(Wd1())
''''''        NothingFlag = False
''''''        FuncAns = DBDRV_scmzc_fcmkc001b_Disp(Wd1(c0).BLOCKID, typ_A.typ_si, typ_A.typ_cr, typ_A.typ_zi, sErrMsg, NothingFlag)
''''''        LWD(c0).flag = NothingFlag
''''''    Next
'''''
'''''
'''''    Dim l       As Long
'''''    Dim m       As Long
'''''    Dim sql     As String
'''''    Dim rs      As OraDynaset    'RecordSet
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo PROC_ERR
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function cmkc001b_DBDataCheck1"
'''''
'''''
'''''    cmkc001b_DBDataCheck1 = FUNCTION_RETURN_SUCCESS
'''''
'''''    Set rs = Nothing
'''''
'''''#If SPEEDUP Then   '高速化実験 02.1.28-2.15 野村
''''''高速化メモ
''''''候補となるブロックとその両端サンプルについて、検査状態をまとめて取得
''''''SQLの発行回数を抑制してメモリ内での処理に切り換える
'''''Dim SMP()   As tSmpMng
'''''Dim idx     As Integer
'''''Dim topIdx  As Integer
'''''Dim botIdx  As Integer
'''''
'''''Debug.Print " 1:" & Time
'''''    sql = vbNullString
''''''    sql = sql & "select"
''''''    sql = sql & "  B.BLOCKID, B.INGOTPOS as TOPPOS, B.INGOTPOS+LENGTH as BOTPOS"
''''''    sql = sql & ", S.CRYNUM, S.INGOTPOS, SMPKBN, HINBAN, REVNUM, FACTORY, OPECOND"
''''''    sql = sql & ", CRYINDRS, CRYRESRS, CRYINDOI, CRYRESOI"
''''''    sql = sql & ", CRYINDB1, CRYRESB1, CRYINDB2, CRYRESB2, CRYINDB3, CRYRESB3"
''''''    sql = sql & ", CRYINDL1, CRYRESL1, CRYINDL2, CRYRESL2, CRYINDL3, CRYRESL3, CRYINDL4, CRYRESL4"
''''''    sql = sql & ", CRYINDCS, CRYRESCS, CRYINDGD, CRYRESGD, CRYINDT, CRYREST, CRYINDEP, CRYRESEP "
''''''    sql = sql & "from TBCME043 S, TBCME040 B "
''''''    sql = sql & "where S.CRYNUM=B.CRYNUM"
''''''    sql = sql & "  and B.INGOTPOS>=0"
''''''    sql = sql & "  and B.DELCLS='0'"
''''''    sql = sql & "  and B.NOWPROC in ('CC600','CC700', 'CC710')"
''''''    sql = sql & "  and B.RSTATCLS='T'"
''''''    sql = sql & "  and B.HOLDCLS='0'"
''''''    sql = sql & "  and ((S.INGOTPOS=B.INGOTPOS) or (S.INGOTPOS=B.INGOTPOS+B.LENGTH)) "
''''''    sql = sql & "order by B.BLOCKID, S.INGOTPOS, S.SMPKBN"
'''''    sql = sql & "select"
'''''    sql = sql & "  B.BLOCKID, B.INGOTPOS as TOPPOS, B.INGOTPOS+LENGTH as BOTPOS"
'''''    sql = sql & ", S.XTALCS, S.INPOSCS, SMPKBNCS, HINBCS, REVNUMCS, FACTORYCS, OPECS"
'''''    sql = sql & ", CRYINDRSCS, CRYRESRS1CS, CRYINDOICS, CRYRESOICS"
'''''    sql = sql & ", CRYINDB1CS, CRYRESB1CS, CRYINDB2CS, CRYRESB2CS, CRYINDB3CS, CRYRESB3CS"
'''''    sql = sql & ", CRYINDL1CS, CRYRESL1CS, CRYINDL2CS, CRYRESL2CS, CRYINDL3CS, CRYRESL3CS, CRYINDL4CS, CRYRESL4CS"
'''''    sql = sql & ", CRYINDCSCS, CRYRESCSCS, CRYINDGDCS, CRYRESGDCS, CRYINDTCS, CRYRESTCS, CRYINDEPCS, CRYRESEPCS "
'''''    sql = sql & "from XSDCS S, TBCME040 B "
'''''    sql = sql & "where S.XTALCS=B.CRYNUM"
'''''    sql = sql & "  and B.INGOTPOS>=0"
'''''    sql = sql & "  and B.DELCLS='0'"
'''''    sql = sql & "  and B.NOWPROC in ('CC600','CC700', 'CC710')"
'''''    sql = sql & "  and B.RSTATCLS='T'"
'''''    sql = sql & "  and B.HOLDCLS='0'"
'''''    sql = sql & "  and ((S.INPOSCS=B.INGOTPOS) or (S.INPOSCS=B.INGOTPOS+B.LENGTH)) "
'''''    sql = sql & "order by B.BLOCKID, S.INPOSCS, S.SMPKBNCS"
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
'''''    ReDim SMP(rs.RecordCount)
'''''    With SMP(0)
'''''        .BLOCKID = " "
'''''        .CRYNUM = " "
'''''        .SMPKBN = " "
'''''        .hinban = " "
'''''        .factory = " "
'''''        .opecond = " "
'''''        .CRYINDRS = " "
'''''        .CRYRESRS = " "
'''''        .CRYINDOI = " "
'''''        .CRYRESOI = " "
'''''        .CRYINDB1 = " "
'''''        .CRYRESB1 = " "
'''''        .CRYINDB2 = " "
'''''        .CRYRESB2 = " "
'''''        .CRYINDB3 = " "
'''''        .CRYRESB3 = " "
'''''        .CRYINDL1 = " "
'''''        .CRYRESL1 = " "
'''''        .CRYINDL2 = " "
'''''        .CRYRESL2 = " "
'''''        .CRYINDL3 = " "
'''''        .CRYRESL3 = " "
'''''        .CRYINDL4 = " "
'''''        .CRYRESL4 = " "
'''''        .CRYINDCS = " "
'''''        .CRYRESCS = " "
'''''        .CRYINDGD = " "
'''''        .CRYRESGD = " "
'''''        .CRYINDT = " "
'''''        .CRYREST = " "
'''''        .CRYINDEP = " "
'''''        .CRYRESEP = " "
'''''    End With
'''''
'''''    For l = 1 To rs.RecordCount
'''''        With SMP(l)
'''''            .BLOCKID = rs("BLOCKID")
'''''            .TOPPOS = rs("TOPPOS")
'''''            .BOTPOS = rs("BOTPOS")
'''''            .CRYNUM = rs("CRYNUM")
'''''            .INGOTPOS = rs("INGOTPOS")
'''''            .SMPKBN = rs("SMPKBN")
'''''            .hinban = rs("HINBAN")
'''''            .REVNUM = rs("REVNUM")
'''''            .factory = rs("FACTORY")
'''''            .opecond = rs("OPECOND")
'''''            .CRYINDRS = rs("CRYINDRS")
'''''            .CRYRESRS = rs("CRYRESRS")
'''''            .CRYINDOI = rs("CRYINDOI")
'''''            .CRYRESOI = rs("CRYRESOI")
'''''            .CRYINDB1 = rs("CRYINDB1")
'''''            .CRYRESB1 = rs("CRYRESB1")
'''''            .CRYINDB2 = rs("CRYINDB2")
'''''            .CRYRESB2 = rs("CRYRESB2")
'''''            .CRYINDB3 = rs("CRYINDB3")
'''''            .CRYRESB3 = rs("CRYRESB3")
'''''            .CRYINDL1 = rs("CRYINDL1")
'''''            .CRYRESL1 = rs("CRYRESL1")
'''''            .CRYINDL2 = rs("CRYINDL2")
'''''            .CRYRESL2 = rs("CRYRESL2")
'''''            .CRYINDL3 = rs("CRYINDL3")
'''''            .CRYRESL3 = rs("CRYRESL3")
'''''            .CRYINDL4 = rs("CRYINDL4")
'''''            .CRYRESL4 = rs("CRYRESL4")
'''''            .CRYINDCS = rs("CRYINDCS")
'''''            .CRYRESCS = rs("CRYRESCS")
'''''            .CRYINDGD = rs("CRYINDGD")
'''''            .CRYRESGD = rs("CRYRESGD")
'''''            .CRYINDT = rs("CRYINDT")
'''''            .CRYREST = rs("CRYREST")
'''''            .CRYINDEP = rs("CRYINDEP")
'''''            .CRYRESEP = rs("CRYRESEP")
'''''        End With
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''    Set rs = Nothing
'''''Debug.Print " 2:" & Time
'''''#End If
'''''
'''''    For l = 1 To UBound(Wd1())
'''''        DoEvents
'''''        LWD(l).flag = False
''''''Debug.Print " " & l & ":" & Time
'''''
'''''        With Wd1(l)
'''''
'''''        ' 購入単結晶のブロックは無条件でＯＫ
'''''        If Mid$(.BLOCKID, 1, 1) <> "8" Then
'''''
'''''            ReDim .SMP(2)
'''''
'''''            ' 上下のサンプル情報取得
'''''#If SPEEDUP Then   '高速化実験 02.1.28-2.15 野村
''''''高速化メモ
''''''一括取得した検査状態配列から、データを取得するように改造
'''''            For m = 1 To 2
'''''                DoEvents
'''''
'''''                topIdx = 0
'''''                botIdx = 0
'''''                For idx = 1 To UBound(SMP)
'''''                    If (SMP(idx).BLOCKID = .BLOCKID) Then
'''''                        If (SMP(idx).SMPKBN = "T") Then
'''''                            topIdx = idx
'''''                        Else
'''''                            botIdx = idx
'''''                        End If
'''''                    ElseIf SMP(idx).BLOCKID > .BLOCKID Then
'''''                        Exit For
'''''                    End If
'''''                Next
'''''                If m = 1 Then
'''''                    If topIdx > 0 Then
'''''                        idx = topIdx
'''''                    Else
'''''                        idx = botIdx
'''''                    End If
'''''                Else
'''''                    If botIdx > 0 Then
'''''                        idx = botIdx
'''''                    Else
'''''                        idx = topIdx
'''''                    End If
'''''                End If
'''''
'''''                With .SMP(m)
'''''                    .CRYNUM = SMP(idx).CRYNUM
'''''                    .INGOTPOS = SMP(idx).INGOTPOS
'''''                    .SMPKBN = SMP(idx).SMPKBN
'''''                    .hinban = SMP(idx).hinban
'''''                    .REVNUM = SMP(idx).REVNUM
'''''                    .factory = SMP(idx).factory
'''''                    .opecond = SMP(idx).opecond
'''''                    .CRYINDRS = SMP(idx).CRYINDRS
'''''                    .CRYRESRS = SMP(idx).CRYRESRS
'''''                    .CRYINDOI = SMP(idx).CRYINDOI
'''''                    .CRYRESOI = SMP(idx).CRYRESOI
'''''                    .CRYINDB1 = SMP(idx).CRYINDB1
'''''                    .CRYRESB1 = SMP(idx).CRYRESB1
'''''                    .CRYINDB2 = SMP(idx).CRYINDB2
'''''                    .CRYRESB2 = SMP(idx).CRYRESB2
'''''                    .CRYINDB3 = SMP(idx).CRYINDB3
'''''                    .CRYRESB3 = SMP(idx).CRYRESB3
'''''                    .CRYINDL1 = SMP(idx).CRYINDL1
'''''                    .CRYRESL1 = SMP(idx).CRYRESL1
'''''                    .CRYINDL2 = SMP(idx).CRYINDL2
'''''                    .CRYRESL2 = SMP(idx).CRYRESL2
'''''                    .CRYINDL3 = SMP(idx).CRYINDL3
'''''                    .CRYRESL3 = SMP(idx).CRYRESL3
'''''                    .CRYINDL4 = SMP(idx).CRYINDL4
'''''                    .CRYRESL4 = SMP(idx).CRYRESL4
'''''                    .CRYINDCS = SMP(idx).CRYINDCS
'''''                    .CRYRESCS = SMP(idx).CRYRESCS
'''''                    .CRYINDGD = SMP(idx).CRYINDGD
'''''                    .CRYRESGD = SMP(idx).CRYRESGD
'''''                    .CRYINDT = SMP(idx).CRYINDT
'''''                    .CRYREST = SMP(idx).CRYREST
'''''                    .CRYINDEP = SMP(idx).CRYINDEP
'''''                    .CRYRESEP = SMP(idx).CRYRESEP
'''''                End With
'''''            Next m
'''''
'''''#Else
'''''            sql = " select "
'''''            sql = sql & " V.E043CRYNUM, "
'''''            sql = sql & " V.E043INGOTPOS, "
'''''            sql = sql & " V.E043SMPKBN, "
'''''            sql = sql & " V.E043HINBAN, "
'''''            sql = sql & " V.E043REVNUM, "
'''''            sql = sql & " V.E043FACTORY, "
'''''            sql = sql & " V.E043OPECOND, "
'''''            sql = sql & " V.E043CRYINDRS, "
'''''            sql = sql & " V.E043CRYRESRS, "
'''''            sql = sql & " V.E043CRYINDOI, "
'''''            sql = sql & " V.E043CRYRESOI, "
'''''            sql = sql & " V.E043CRYINDB1, "
'''''            sql = sql & " V.E043CRYRESB1, "
'''''            sql = sql & " V.E043CRYINDB2, "
'''''            sql = sql & " V.E043CRYRESB2, "
'''''            sql = sql & " V.E043CRYINDB3, "
'''''            sql = sql & " V.E043CRYRESB3, "
'''''            sql = sql & " V.E043CRYINDL1, "
'''''            sql = sql & " V.E043CRYRESL1, "
'''''            sql = sql & " V.E043CRYINDL2, "
'''''            sql = sql & " V.E043CRYRESL2, "
'''''            sql = sql & " V.E043CRYINDL3, "
'''''            sql = sql & " V.E043CRYRESL3, "
'''''            sql = sql & " V.E043CRYINDL4, "
'''''            sql = sql & " V.E043CRYRESL4, "
'''''            sql = sql & " V.E043CRYINDCS, "
'''''            sql = sql & " V.E043CRYRESCS, "
'''''            sql = sql & " V.E043CRYINDGD, "
'''''            sql = sql & " V.E043CRYRESGD, "
'''''            sql = sql & " V.E043CRYINDT, "
'''''            sql = sql & " V.E043CRYREST, "
'''''            sql = sql & " V.E043CRYINDEP, "
'''''            sql = sql & " V.E043CRYRESEP "
'''''            sql = sql & " from VECME010 V "
'''''            sql = sql & " where E040CRYNUM = '" & .CRYNUM & "' "
'''''            sql = sql & " and   E040INGOTPOS = '" & .INGOTPOS & "' "
'''''            sql = sql & " order by E043INGOTPOS"
'''''
'''''            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''            For m = 1 To 2
'''''                DoEvents
'''''                .SMP(m).CRYNUM = rs("E043CRYNUM")
'''''                .SMP(m).INGOTPOS = rs("E043INGOTPOS")
'''''                .SMP(m).SMPKBN = rs("E043SMPKBN")
'''''                .SMP(m).hinban = rs("E043HINBAN")
'''''                .SMP(m).REVNUM = rs("E043REVNUM")
'''''                .SMP(m).factory = rs("E043FACTORY")
'''''                .SMP(m).opecond = rs("E043OPECOND")
'''''                .SMP(m).CRYINDRS = rs("E043CRYINDRS")
'''''                .SMP(m).CRYRESRS = rs("E043CRYRESRS")
'''''                .SMP(m).CRYINDOI = rs("E043CRYINDOI")
'''''                .SMP(m).CRYRESOI = rs("E043CRYRESOI")
'''''                .SMP(m).CRYINDB1 = rs("E043CRYINDB1")
'''''                .SMP(m).CRYRESB1 = rs("E043CRYRESB1")
'''''                .SMP(m).CRYINDB2 = rs("E043CRYINDB2")
'''''                .SMP(m).CRYRESB2 = rs("E043CRYRESB2")
'''''                .SMP(m).CRYINDB3 = rs("E043CRYINDB3")
'''''                .SMP(m).CRYRESB3 = rs("E043CRYRESB3")
'''''                .SMP(m).CRYINDL1 = rs("E043CRYINDL1")
'''''                .SMP(m).CRYRESL1 = rs("E043CRYRESL1")
'''''                .SMP(m).CRYINDL2 = rs("E043CRYINDL2")
'''''                .SMP(m).CRYRESL2 = rs("E043CRYRESL2")
'''''                .SMP(m).CRYINDL3 = rs("E043CRYINDL3")
'''''                .SMP(m).CRYRESL3 = rs("E043CRYRESL3")
'''''                .SMP(m).CRYINDL4 = rs("E043CRYINDL4")
'''''                .SMP(m).CRYRESL4 = rs("E043CRYRESL4")
'''''                .SMP(m).CRYINDCS = rs("E043CRYINDCS")
'''''                .SMP(m).CRYRESCS = rs("E043CRYRESCS")
'''''                .SMP(m).CRYINDGD = rs("E043CRYINDGD")
'''''                .SMP(m).CRYRESGD = rs("E043CRYRESGD")
'''''                .SMP(m).CRYINDT = rs("E043CRYINDT")
'''''                .SMP(m).CRYREST = rs("E043CRYREST")
'''''                .SMP(m).CRYINDEP = rs("E043CRYINDEP")
'''''                .SMP(m).CRYRESEP = rs("E043CRYRESEP")
'''''
'''''                rs.MoveNext
'''''            Next m
'''''            rs.Close
'''''            Set rs = Nothing
'''''#End If
'''''
''''''高速化メモ
''''''品番仕様/Cs/EPD/LTはまだブロック毎にSQLを投げている
''''''ここをまとめていけば、あと5秒程度縮むのではないかと思われる
''''''ただし、Cs/LTについては結果取得の方法が変わるので、その後の検討が必要
''''''いずれにせよ、対象結晶全てについてCs/LT/EPD指示のあるサンプルを抜き出せばよいはず
'''''
'''''            ' 品番の仕様情報取得
'''''            For m = 1 To 2
'''''                If Trim$(.SMP(m).hinban) = "G" Or Trim$(.SMP(m).hinban) = "Z" Then
'''''                    .SMP(m).HSXCNHWS = "S"
'''''                    .SMP(m).HSXLTHWS = "S"
'''''                    .SMP(m).EPD = "S"
'''''                ElseIf Len(Trim$(.SMP(m).hinban)) Then
'''''                    sql = " select "
'''''                    sql = sql & " S.HSXCNHWS, "
'''''                    sql = sql & " S.HSXLTHWS, "
'''''                    sql = sql & " 'H' as EPD "
'''''                    sql = sql & " from TBCME019 S "
'''''                    sql = sql & " where S.HINBAN = '" & .SMP(m).hinban & "' "
'''''                    sql = sql & " and S.MNOREVNO = " & .SMP(m).REVNUM & " "
'''''                    sql = sql & " and S.FACTORY = '" & .SMP(m).factory & "' "
'''''                    sql = sql & " and S.OPECOND = '" & .SMP(m).opecond & "' "
'''''
'''''                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                    .SMP(m).HSXCNHWS = rs("HSXCNHWS")
'''''                    .SMP(m).HSXLTHWS = rs("HSXLTHWS")
'''''                    .SMP(m).EPD = rs("EPD")
'''''
'''''                    rs.Close
'''''                    Set rs = Nothing
'''''                Else
'''''                    '空品番の場合
'''''                    .SMP(m).HSXCNHWS = " "
'''''                    .SMP(m).HSXLTHWS = " "
'''''                    .SMP(m).EPD = " "
'''''                End If
'''''            Next m
'''''
'''''            ' チェック
'''''            For m = 1 To 2
'''''                DoEvents
'''''                ' CSのチェック
''''''                If (.SMP(m).HSXCNHWS = "H" Or .SMP(m).HSXCNHWS = "S") And .SMP(m).CRYINDCS = "0" Then  ' 参考評価はなくてもＯＫ
'''''                If .SMP(m).HSXCNHWS = "H" And .SMP(m).CRYINDCS = "0" Then
'''''
'''''                    sql = "select CRYRESCSCS as RES from XSDCS "
'''''                    sql = sql & "where CRYNUM = '" & .SMP(m).CRYNUM & "' "
'''''                    sql = sql & "  and INPOSCS >= " & .SMP(m).INGOTPOS
'''''                    sql = sql & "  and CRYINDCSCS<>'0'"
'''''                    sql = sql & " order by INPOSCS"
'''''
'''''
'''''                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                    If rs.RecordCount Then
'''''                        If rs("RES") = "0" Then LWD(l).flag = True
'''''                    End If
'''''
'''''                    rs.Close
'''''                    Set rs = Nothing
'''''                End If
'''''
'''''                ' LTのチェック
''''''                If (.SMP(m).HSXLTHWS = "H" Or .SMP(m).HSXLTHWS = "S") And .SMP(m).CRYINDT = "0" And LWD(l).flag = False Then ' 参考評価はなくてもＯＫ
'''''                If .SMP(m).HSXLTHWS = "H" And .SMP(m).CRYINDT = "0" And LWD(l).flag = False Then
'''''
'''''                    sql = "select CRYRESTCS as RES from XSDCS "
'''''                    sql = sql & "where XTALCS = '" & .SMP(m).CRYNUM & "' "
'''''                    sql = sql & "  and INPOSCS >= " & .SMP(m).INGOTPOS
'''''                    sql = sql & "  and CRYINDTCS<>'0'"
'''''                    sql = sql & " order by INPOSCS"
'''''
'''''                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                    If rs.RecordCount Then
'''''                        If rs("RES") = "0" Then LWD(l).flag = True
'''''                    End If
'''''
'''''                    rs.Close
'''''                    Set rs = Nothing
'''''                End If
'''''
'''''                ' EPDのチェック
''''''                If (.SMP(m).EPD = "H" Or .SMP(m).EPD = "S") And .SMP(m).CRYINDEP = "0" And LWD(l).flag = False Then ' Sはありえなけど統一
'''''                If .SMP(m).EPD = "H" And .SMP(m).CRYINDEP = "0" And LWD(l).flag = False Then ' Sはありえなけど統一
'''''
'''''                    sql = "select CRYRESEPCS as RES from XSDCS "
'''''                    sql = sql & "where XTALCS = '" & .SMP(m).CRYNUM & "' "
'''''                    sql = sql & "  and INPOSCS >= " & .SMP(m).INGOTPOS
'''''                    sql = sql & "  and CRYINDEPCS<>'0'"
'''''                    sql = sql & " order by INPOSCS"
'''''
'''''                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                    If rs.RecordCount Then
'''''                        If rs("RES") = "0" Then LWD(l).flag = True
'''''                    End If
'''''
'''''                    rs.Close
'''''                    Set rs = Nothing
'''''                End If
''''''                If LWD(l).flag = True Then
''''''                    Exit For
''''''                End If
'''''            Next m
'''''        End If
'''''
'''''        End With    ' .Wd1()
'''''
'''''    Next l
'''''
'''''PROC_EXIT:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''PROC_ERR:
'''''    'エラーハンドラ
'''''    gErr.HandleError
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    cmkc001b_DBDataCheck1 = FUNCTION_RETURN_FAILURE
'''''    Resume PROC_EXIT
'''''End Function


'''''Public Function cmkc001b_DBDataCheck3(LWD() As cmkc001b_LockWait, _
'''''                                 Wd3() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
'''''    Dim c0              As Integer
'''''    Dim c1              As Integer
'''''    Dim c2              As Integer
'''''    Dim MaxRec          As Integer
'''''    Dim RecCount        As Integer
'''''    Dim EQFlag          As Boolean
'''''    Dim sql             As String       'SQL全体
'''''    Dim rs              As OraDynaset    'RecordSet
'''''    Dim GrpCount1       As Integer
'''''    Dim GrpCount2       As Integer
'''''    Dim ColorFlag       As Boolean
'''''    Dim TotalBlk        As Integer
'''''    Dim CheckPoint      As Integer
'''''    Dim CheckEnd        As Integer
'''''    Dim tempGrpFlag     As String * 1
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo PROC_ERR
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp"
'''''
'''''    cmkc001b_DBDataCheck3 = FUNCTION_RETURN_SUCCESS
'''''    TotalBlk = UBound(Wd3())
'''''
'''''Debug.Print " 1:" & Time
'''''
'''''    'CC700のブロックの結晶一覧を作る
'''''    ReDim GrpInfo(1) As cmkc001b_Wait3
'''''    GrpInfo(1).CRYNUM = vbNullString
'''''    c1 = 0
'''''    For c0 = 1 To TotalBlk
'''''        DoEvents
'''''        If c1 = 0 Then
'''''            GrpInfo(1).CRYNUM = Wd3(c0).CRYNUM
'''''        End If
'''''        MaxRec = UBound(GrpInfo())
'''''        EQFlag = False
'''''        c1 = 1
'''''        Do While c1 <= MaxRec
'''''            DoEvents
'''''            If Wd3(c0).CRYNUM = GrpInfo(c1).CRYNUM Then
'''''                EQFlag = True
'''''                Exit Do
'''''            End If
'''''            c1 = c1 + 1
'''''        Loop
'''''        If Not EQFlag Then
'''''            ReDim Preserve GrpInfo(MaxRec + 1) As cmkc001b_Wait3
'''''            GrpInfo(MaxRec + 1).CRYNUM = Wd3(c0).CRYNUM
'''''        End If
'''''    Next
'''''
'''''Debug.Print " 2:" & Time
'''''
'''''    '結晶に含まれる全てのブロックを求める
'''''    MaxRec = UBound(GrpInfo())
'''''    For c0 = 1 To MaxRec
'''''        sql = "select "
'''''        sql = sql & "BLOCKID, "
'''''        sql = sql & "INGOTPOS, "
'''''        sql = sql & "LENGTH, "
'''''        sql = sql & "NOWPROC, "
'''''        sql = sql & "HOLDCLS "
'''''        sql = sql & "from TBCME040 "
'''''        sql = sql & "where CRYNUM='" & GrpInfo(c0).CRYNUM & "' "
''''''2001/11/14 S.Sano        sql = sql & "and LSTATCLS='T' "
''''''2001/11/14 S.Sano        sql = sql & "and RSTATCLS='T' "
''''''2001/11/14 S.Sano        sql = sql & "and DELCLS='0' "
'''''        'sql = sql & "and HOLDCLS='0' "
'''''        sql = sql & "order by BLOCKID "
'''''
'''''
'''''        'データを抽出する
'''''        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''        RecCount = rs.RecordCount
'''''        If RecCount = 0 Then
'''''            rs.Close
'''''            GoTo PROC_EXIT
'''''        End If
'''''        ReDim GrpInfo(c0).blkInfo(RecCount) As cmkc001b_Wait3_BLK
'''''        For c1 = 1 To RecCount
'''''            GrpInfo(c0).blkInfo(c1).BLOCKID = rs("BLOCKID")
'''''            GrpInfo(c0).blkInfo(c1).INGOTPOS = rs("INGOTPOS")
'''''            GrpInfo(c0).blkInfo(c1).LENGTH = rs("LENGTH")
'''''            GrpInfo(c0).blkInfo(c1).NOWPROC = rs("NOWPROC")
'''''            GrpInfo(c0).blkInfo(c1).HOLDCLS = rs("HOLDCLS")
'''''            rs.MoveNext
'''''        Next
'''''        rs.Close
'''''    Next
'''''
'''''Debug.Print " 3:" & Time
'''''    'ブロックの上下品番を求める
'''''#If SPEEDUP Then   '高速化実験 02.1.28-2.15 野村
''''''高速化メモ
''''''ブロックの上下品番を求めるだけなら、1回のSQLでまとめて情報を取得できるはず
'''''Dim blkID() As String
'''''Dim topHin() As tFullHinban
'''''Dim botHin() As tFullHinban
'''''Dim idx As Integer
'''''Dim rsCount As Integer
'''''Dim found As Boolean
'''''
'''''    sql = vbNullString
'''''    sql = sql & "select"
'''''    sql = sql & "  b.BLOCKID"
'''''    sql = sql & ", TOP.HINBAN as THINBAN, TOP.REVNUM as TREVNUM, TOP.FACTORY as TFACTORY, TOP.OPECOND as TOPECOND"
'''''    sql = sql & ", BOT.HINBAN as BHINBAN, BOT.REVNUM as BREVNUM, BOT.FACTORY as BFACTORY, BOT.OPECOND as BOPECOND "
'''''    sql = sql & "from TBCME040 B, TBCME041 TOP, TBCME041 BOT "
'''''    sql = sql & "Where b.CRYNUM = Top.CRYNUM"
'''''    sql = sql & "  and B.CRYNUM=BOT.CRYNUM"
'''''    sql = sql & "  and B.INGOTPOS>=0"
'''''    sql = sql & "  and B.DELCLS='0'"
'''''    sql = sql & "  and B.NOWPROC in ('CC600','CC700', 'CC710')"
'''''    sql = sql & "  and B.RSTATCLS='T'"
'''''    sql = sql & "  and B.HOLDCLS='0'"
'''''    sql = sql & "  and B.INGOTPOS>=TOP.INGOTPOS"
'''''    sql = sql & "  and B.INGOTPOS<TOP.INGOTPOS+TOP.LENGTH"
'''''    sql = sql & "  and B.INGOTPOS+B.LENGTH>BOT.INGOTPOS"
'''''    sql = sql & "  and B.INGOTPOS+B.LENGTH<=BOT.INGOTPOS+BOT.LENGTH "
'''''    sql = sql & "order by B.BLOCKID"
'''''
'''''    'データを抽出する
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    rsCount = rs.RecordCount
'''''    ReDim blkID(1 To rsCount)
'''''    ReDim topHin(1 To rsCount)
'''''    ReDim botHin(1 To rsCount)
'''''    For c0 = 1 To rsCount
'''''        blkID(c0) = rs!BLOCKID
'''''        topHin(c0).hinban = rs!THINBAN
'''''        topHin(c0).mnorevno = rs!TREVNUM
'''''        topHin(c0).factory = rs!TFACTORY
'''''        topHin(c0).opecond = rs!TOPECOND
'''''        botHin(c0).hinban = rs!BHINBAN
'''''        botHin(c0).mnorevno = rs!BREVNUM
'''''        botHin(c0).factory = rs!BFACTORY
'''''        botHin(c0).opecond = rs!BOPECOND
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''
'''''    For c0 = 1 To MaxRec
'''''        RecCount = UBound(GrpInfo(c0).blkInfo())
'''''        For c1 = 1 To RecCount
'''''            found = False
'''''            For idx = 1 To rsCount
'''''                If blkID(idx) = GrpInfo(c0).blkInfo(c1).BLOCKID Then
'''''                    found = True
'''''                    Exit For
'''''                ElseIf blkID(idx) > GrpInfo(c0).blkInfo(c1).BLOCKID Then
'''''                    Exit For
'''''                End If
'''''            Next
'''''
'''''            If found Then
'''''                GrpInfo(c0).blkInfo(c1).topHin.hinban = topHin(idx).hinban
'''''                GrpInfo(c0).blkInfo(c1).topHin.factory = topHin(idx).factory
'''''                GrpInfo(c0).blkInfo(c1).topHin.opecond = topHin(idx).opecond
'''''                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = topHin(idx).mnorevno
'''''            Else
'''''                GrpInfo(c0).blkInfo(c1).topHin.hinban = ""
'''''                GrpInfo(c0).blkInfo(c1).topHin.factory = ""
'''''                GrpInfo(c0).blkInfo(c1).topHin.opecond = ""
'''''                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = 0
'''''            End If
'''''
'''''            If found Then
'''''                GrpInfo(c0).blkInfo(c1).botHin.hinban = botHin(idx).hinban
'''''                GrpInfo(c0).blkInfo(c1).botHin.factory = botHin(idx).factory
'''''                GrpInfo(c0).blkInfo(c1).botHin.opecond = botHin(idx).opecond
'''''                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = botHin(idx).mnorevno
'''''            Else
'''''                GrpInfo(c0).blkInfo(c1).botHin.hinban = ""
'''''                GrpInfo(c0).blkInfo(c1).botHin.factory = ""
'''''                GrpInfo(c0).blkInfo(c1).botHin.opecond = ""
'''''                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = 0
'''''            End If
'''''        Next
'''''    Next
'''''#Else
'''''    For c0 = 1 To MaxRec
'''''        RecCount = UBound(GrpInfo(c0).blkInfo())
'''''        For c1 = 1 To RecCount
'''''            sql = "select "
'''''            sql = sql & "HINBAN, "
'''''            sql = sql & "REVNUM, "
'''''            sql = sql & "FACTORY, "
'''''            sql = sql & "OPECOND "
'''''            sql = sql & "from TBCME041 "
'''''            sql = sql & "where CRYNUM='" & GrpInfo(c0).CRYNUM & "' "
''''''2001/11/14 S.Sano            sql = sql & "and INGOTPOS <= " & GrpInfo(c0).blkInfo(c1).INGOTPOS & " "
'''''            sql = sql & "and INGOTPOS = " & GrpInfo(c0).blkInfo(c1).INGOTPOS & " " '2001/11/14 S.Sano
''''''2001/11/14 S.Sano            sql = sql & "and (INGOTPOS + LENGTH) > " & GrpInfo(c0).blkInfo(c1).INGOTPOS & " "
'''''
'''''            'データを抽出する
'''''            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''            RecCount = rs.RecordCount
'''''            If RecCount = 0 Then
'''''                GrpInfo(c0).blkInfo(c1).topHin.hinban = ""
'''''                GrpInfo(c0).blkInfo(c1).topHin.factory = ""
'''''                GrpInfo(c0).blkInfo(c1).topHin.opecond = ""
'''''                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = 0
'''''            Else
'''''                GrpInfo(c0).blkInfo(c1).topHin.hinban = rs("HINBAN")
'''''                GrpInfo(c0).blkInfo(c1).topHin.factory = rs("FACTORY")
'''''                GrpInfo(c0).blkInfo(c1).topHin.opecond = rs("OPECOND")
'''''                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = rs("REVNUM")
'''''            End If
'''''            rs.Close
'''''
'''''            sql = "select "
'''''            sql = sql & "HINBAN, "
'''''            sql = sql & "REVNUM, "
'''''            sql = sql & "FACTORY, "
'''''            sql = sql & "OPECOND "
'''''            sql = sql & "from TBCME041 "
'''''            sql = sql & "where CRYNUM='" & GrpInfo(c0).CRYNUM & "' "
'''''            sql = sql & "and INGOTPOS < " & GrpInfo(c0).blkInfo(c1).INGOTPOS + GrpInfo(c0).blkInfo(c1).LENGTH & " "
'''''            sql = sql & "and (INGOTPOS + LENGTH) >= " & GrpInfo(c0).blkInfo(c1).INGOTPOS + GrpInfo(c0).blkInfo(c1).LENGTH & " "
'''''
'''''            'データを抽出する
'''''            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''            RecCount = rs.RecordCount
'''''            If RecCount = 0 Then
'''''                GrpInfo(c0).blkInfo(c1).botHin.hinban = ""
'''''                GrpInfo(c0).blkInfo(c1).botHin.factory = ""
'''''                GrpInfo(c0).blkInfo(c1).botHin.opecond = ""
'''''                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = 0
'''''            Else
'''''                GrpInfo(c0).blkInfo(c1).botHin.hinban = rs("HINBAN")
'''''                GrpInfo(c0).blkInfo(c1).botHin.factory = rs("FACTORY")
'''''                GrpInfo(c0).blkInfo(c1).botHin.opecond = rs("OPECOND")
'''''                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = rs("REVNUM")
'''''            End If
'''''            rs.Close
'''''        Next
'''''    Next
'''''#End If
'''''
'''''Debug.Print " 4:" & Time
'''''    '求めた情報からグループを求める
'''''    GrpCount1 = 0
'''''    GrpCount2 = 0
'''''    For c0 = 1 To MaxRec
'''''        GrpCount1 = GrpCount1 + 1
'''''        GrpCount2 = GrpCount2 + 1
'''''        RecCount = UBound(GrpInfo(c0).blkInfo())
'''''        For c1 = 1 To RecCount
'''''            'ブロック切れ目で品番が変われば別グループと判断する
'''''            Select Case c1
'''''            Case 1
'''''                GrpInfo(c0).blkInfo(c1).GRPFLG1 = GrpCount1
'''''            Case Else
'''''                If (GrpInfo(c0).blkInfo(c1).topHin.factory <> GrpInfo(c0).blkInfo(c1 - 1).botHin.factory) Or _
'''''                   (GrpInfo(c0).blkInfo(c1).topHin.hinban <> GrpInfo(c0).blkInfo(c1 - 1).botHin.hinban) Or _
'''''                   (GrpInfo(c0).blkInfo(c1).topHin.opecond <> GrpInfo(c0).blkInfo(c1 - 1).botHin.opecond) Or _
'''''                   (GrpInfo(c0).blkInfo(c1).topHin.REVNUM <> GrpInfo(c0).blkInfo(c1 - 1).botHin.REVNUM) Then
'''''                    GrpCount1 = GrpCount1 + 1
'''''                End If
'''''                GrpInfo(c0).blkInfo(c1).GRPFLG1 = GrpCount1
'''''            End Select
'''''
'''''            '同一グループ内で、工程違いのブロックが存在した場合、同一グループ内の
'''''            '小グループとしてグループ分けする。
'''''            'CC710以外なら対象外としグループ判定をしない
'''''            If GrpInfo(c0).blkInfo(c1).NOWPROC = PROCD_NUKISI_SIJI And GrpInfo(c0).blkInfo(c1).HOLDCLS = "0" Then
'''''                Select Case c1
'''''                Case 1
'''''                    GrpInfo(c0).blkInfo(c1).GRPFLG2 = GrpCount2
'''''                Case Else
'''''                    If (GrpInfo(c0).blkInfo(c1).topHin.factory <> GrpInfo(c0).blkInfo(c1 - 1).botHin.factory) Or _
'''''                       (GrpInfo(c0).blkInfo(c1).topHin.hinban <> GrpInfo(c0).blkInfo(c1 - 1).botHin.hinban) Or _
'''''                       (GrpInfo(c0).blkInfo(c1).topHin.opecond <> GrpInfo(c0).blkInfo(c1 - 1).botHin.opecond) Or _
'''''                       (GrpInfo(c0).blkInfo(c1).topHin.REVNUM <> GrpInfo(c0).blkInfo(c1 - 1).botHin.REVNUM) Then
'''''                        GrpCount2 = GrpCount2 + 1
'''''                    End If
'''''                    GrpInfo(c0).blkInfo(c1).GRPFLG2 = GrpCount2
'''''                End Select
'''''            Else
'''''                GrpCount2 = GrpCount2 + 1
'''''                GrpInfo(c0).blkInfo(c1).GRPFLG2 = 0
'''''            End If
'''''        Next
'''''    Next
'''''
'''''Debug.Print " 5:" & Time
'''''    '求めた情報から表示色を求める
'''''    For c0 = 1 To MaxRec
'''''        RecCount = UBound(GrpInfo(c0).blkInfo())
'''''        ColorFlag = False
'''''        CheckPoint = 0
'''''        For c1 = 1 To RecCount
'''''            If CheckPoint > 0 Then
'''''                If GrpInfo(c0).blkInfo(c1).GRPFLG1 <> GrpInfo(c0).blkInfo(CheckPoint).GRPFLG1 Then
'''''                    For c2 = CheckPoint To c1 - 1
'''''                        GrpInfo(c0).blkInfo(c2).COLORFLG = ColorFlag
'''''                    Next
'''''                    ColorFlag = False
'''''                    CheckPoint = c1
'''''                End If
'''''            Else
'''''                CheckPoint = c1
'''''            End If
'''''            If CheckPoint > 0 Then
'''''                If (GrpInfo(c0).blkInfo(c1).NOWPROC = PROCD_SETUDAN) Or _
'''''                   (GrpInfo(c0).blkInfo(c1).NOWPROC = PROCD_KESSYOU_SOUGOUHANTEI) Or _
'''''                   (GrpInfo(c0).blkInfo(c1).NOWPROC = PROCD_KESSYOU_SAISYUU_HARAIDASI) Or _
'''''                   (GrpInfo(c0).blkInfo(c1).HOLDCLS = "1") Then
'''''                    ColorFlag = True
'''''                End If
'''''            End If
'''''        Next
'''''        For c1 = CheckPoint To RecCount
'''''            GrpInfo(c0).blkInfo(c1).COLORFLG = ColorFlag
'''''        Next
'''''    Next
'''''
'''''Debug.Print " 6:" & Time
'''''    For c0 = 1 To MaxRec
'''''        RecCount = UBound(GrpInfo(c0).blkInfo())
'''''        For c1 = 1 To RecCount
'''''            For c2 = 1 To TotalBlk
'''''                If Wd3(c2).BLOCKID = GrpInfo(c0).blkInfo(c1).BLOCKID Then
'''''                    LWD(c2).flag = GrpInfo(c0).blkInfo(c1).COLORFLG
'''''                    LWD(c2).Grp = GrpInfo(c0).blkInfo(c1).GRPFLG2
'''''                    Exit For
'''''                End If
'''''            Next
'''''        Next
'''''    Next
''''''    Debug.Print Now
'''''
'''''PROC_EXIT:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''PROC_ERR:
'''''    'エラーハンドラ
'''''    cmkc001b_DBDataCheck3 = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume PROC_EXIT
'''''End Function


''''''概要    :待ち一覧 初期表示用ＤＢドライバ（検査待ち）
''''''ﾊﾟﾗﾒｰﾀ  :変数名       ,IO  ,型                                 ,説明
''''''        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,初期表示用
''''''        :戻ﾘ値        ,O   ,FUNCTION_RETURN                    ,読み込み成否
''''''説明    :
''''''履歴    :2001/07/06 蔵本 作成
'''''Public Function DBDRV_scmzc_fcmkc001b_Disp1(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
'''''
'''''    Dim sql As String       'SQL全体
'''''    Dim rs As OraDynaset    'RecordSet
'''''    Dim recCnt As Long      'ブロック管理のレコード数
'''''    Dim i As Long
'''''    Dim j As Long
'''''    Dim k As Long
'''''    Dim BlockIdBuf As String
'''''
'''''    '<検査待ち＞
'''''    'ブロック管理テーブルからブロックID、更新日付取得（検査実績が未検査のもの）
'''''
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo PROC_ERR
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp1"
'''''
'''''    DBDRV_scmzc_fcmkc001b_Disp1 = FUNCTION_RETURN_SUCCESS
'''''
'''''    'ブロックID、更新日付の取得
'''''    sql = "select distinct "
'''''    sql = sql & " V.E040CRYNUM, "
'''''    sql = sql & " V.E040INGOTPOS, "
'''''    sql = sql & " V.E040BLOCKID, "
'''''    sql = sql & " V.E040UPDDATE, "
'''''    sql = sql & " V.E040HOLDCLS, "
'''''    sql = sql & " H.HINBAN, "            ' 品番
'''''    sql = sql & " H.REVNUM, "            ' 製品番号改訂番号
'''''    sql = sql & " H.FACTORY, "           ' 工場
'''''    sql = sql & " H.OPECOND, "           ' 操業条件
'''''    sql = sql & " S.HSXTYPE, "           ' 品ＳＸタイプ
'''''    sql = sql & " S.HSXCDIR, "            ' 品ＳＸ結晶面方位
'''''    sql = sql & " H.INGOTPOS "
'''''    sql = sql & " from "
'''''    sql = sql & " VECME010 V, TBCME041 H, TBCME018 S "
'''''    sql = sql & " where "
'''''    sql = sql & " V.E040CRYNUM = H.CRYNUM "
'''''    sql = sql & " and H.HINBAN = S.HINBAN "
'''''    sql = sql & " and H.REVNUM = S.MNOREVNO "
'''''    sql = sql & " and H.FACTORY = S.FACTORY "
'''''    sql = sql & " and H.OPECOND = S.OPECOND "
'''''                'ブロック内の品番検索
'''''    sql = sql & " and (( V.E040INGOTPOS >= H.INGOTPOS "
'''''    sql = sql & " and V.E040INGOTPOS < H.INGOTPOS + H.LENGTH ) "
'''''    sql = sql & " or ( V.E040INGOTPOS + V.E040LENGTH > H.INGOTPOS "
'''''    sql = sql & " and V.E040INGOTPOS + V.E040LENGTH < H.INGOTPOS + H.LENGTH  ) "
'''''    sql = sql & " or ( H.INGOTPOS >= V.E040INGOTPOS "
'''''    sql = sql & " and H.INGOTPOS < V.E040INGOTPOS + V.E040LENGTH ) "
'''''    sql = sql & " or ( H.INGOTPOS + H.LENGTH > V.E040INGOTPOS "
'''''    sql = sql & " and H.INGOTPOS + H.LENGTH < V.E040INGOTPOS + V.E040LENGTH )) "
'''''                '工程コード、状態、区分の条件指定
'''''    sql = sql & " and V.E040NOWPROC='CC600' "
'''''    sql = sql & " and V.E040LSTATCLS='T' "
'''''    sql = sql & " and V.E040RSTATCLS='T' "
'''''    sql = sql & " and V.E040DELCLS='0' "
'''''    'sql = sql & " and V.E040HOLDCLS='0' " ' ホールドブロックも取得
'''''                '指示が0でなく実績が0
'''''    sql = sql & " and ((V.E043CRYINDRS<>'0' and V.E043CRYRESRS='0') "         ' 結晶検査実績（Rs)
'''''    sql = sql & " or (V.E043CRYINDOI<>'0' and V.E043CRYRESOI='0') "         ' 結晶検査実績（Oi)
'''''    sql = sql & " or (V.E043CRYINDB1<>'0' and V.E043CRYRESB1='0')"          ' 結晶検査実績（B1)
'''''    sql = sql & " or (V.E043CRYINDB2<>'0' and V.E043CRYRESB2='0') "         ' 結晶検査実績（B2）
'''''    sql = sql & " or (V.E043CRYINDB3<>'0' and V.E043CRYRESB3='0') "         ' 結晶検査実績（B3)
'''''    sql = sql & " or (V.E043CRYINDL1<>'0' and V.E043CRYRESL1='0') "         ' 結晶検査実績（L1)
'''''    sql = sql & " or (V.E043CRYINDL2<>'0' and V.E043CRYRESL2='0') "         ' 結晶検査実績（L2)
'''''    sql = sql & " or (V.E043CRYINDL3<>'0' and V.E043CRYRESL3='0') "         ' 結晶検査実績（L3)
'''''    sql = sql & " or (V.E043CRYINDL4<>'0' and V.E043CRYRESL4='0') "         ' 結晶検査実績（L4)
'''''    sql = sql & " or (V.E043CRYINDCS<>'0' and V.E043CRYRESCS='0') "         ' 結晶検査実績（Cs)
'''''    sql = sql & " or (V.E043CRYINDGD<>'0' and V.E043CRYRESGD='0') "         ' 結晶検査実績（GD)
'''''    sql = sql & " or (V.E043CRYINDT<>'0' and V.E043CRYREST='0') "           ' 結晶検査実績（T)
'''''    sql = sql & " or (V.E043CRYINDEP<>'0' and V.E043CRYRESEP='0')) "         ' 結晶検査実績（EPD)
'''''    sql = sql & " order by V.E040BLOCKID, H.INGOTPOS "
'''''
'''''    'データを抽出する
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''    'レコード0件時は正常
'''''    If rs.RecordCount = 0 Then
'''''        rs.Close
'''''        ReDim records(0)
'''''    Else
'''''        BlockIdBuf = vbNullString
'''''        recCnt = rs.RecordCount
'''''        j = 0
'''''        For i = 1 To recCnt
'''''            DoEvents
'''''        'ブロックID等の格納
'''''            If rs("E040BLOCKID") <> BlockIdBuf Then
'''''
'''''                j = j + 1
'''''                ReDim Preserve records(j)
'''''
'''''                With records(j)
'''''                    .CRYNUM = rs("E040CRYNUM")
'''''                    .INGOTPOS = rs("E040INGOTPOS")
'''''                    .BLOCKID = rs("E040BLOCKID")   ' ブロックID
'''''                    .UPDDATE = rs("E040UPDDATE")   ' 更新日付
'''''                    .HOLDCLS = rs("E040HOLDCLS")   ' ホールド区分
'''''                    BlockIdBuf = records(j).BLOCKID
'''''                    .HSXTYPE = rs("HSXTYPE")
'''''                    .HSXCDIR = rs("HSXCDIR")
'''''                    .Judg = " "
'''''                End With
'''''
'''''                k = 1
'''''            End If
'''''
'''''            '品番の格納
'''''            ReDim Preserve records(j).HIN(k)
'''''            records(j).HIN(k).hinban = rs("HINBAN")
'''''            records(j).HIN(k).mnorevno = rs("REVNUM")
'''''            records(j).HIN(k).factory = rs("FACTORY")
'''''            records(j).HIN(k).opecond = rs("OPECOND")
'''''            k = k + 1
'''''            rs.MoveNext
'''''        Next i
'''''        rs.Close
'''''
'''''    End If
'''''
'''''
'''''
'''''PROC_EXIT:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''PROC_ERR:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    DBDRV_scmzc_fcmkc001b_Disp1 = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume PROC_EXIT
'''''End Function


''''''概要    :待ち一覧 初期表示用ＤＢドライバ（判定待ち）
''''''ﾊﾟﾗﾒｰﾀ  :変数名       ,IO  ,型                                 ,説明
''''''        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,初期表示用
''''''        :戻ﾘ値        ,O   ,FUNCTION_RETURN                    ,読み込み成否
''''''説明    :
''''''履歴    :2001/07/06 蔵本 作成
'''''Public Function DBDRV_scmzc_fcmkc001b_Disp2(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
'''''
'''''    '＜判定待ち＞
'''''    '検査待ちが押されている場合と逆で０が一つもないもの
'''''    Dim sql As String       'SQL全体
'''''    Dim rs As OraDynaset    'RecordSet
'''''    Dim recCnt As Long      'ブロック管理のレコード数
'''''    Dim i As Long
'''''    Dim j As Long
'''''    Dim k As Long
'''''    Dim BlockIdBuf As String
'''''
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo PROC_ERR
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp2"
'''''
'''''    DBDRV_scmzc_fcmkc001b_Disp2 = FUNCTION_RETURN_SUCCESS
'''''
'''''    sql = "select distinct "
'''''    sql = sql & " B.CRYNUM, "
'''''    sql = sql & " B.INGOTPOS as ss, "
''''''    sql = sql & " B.LENGTH, "             ' 長さ追加 2001/11/8
'''''    sql = sql & " B.BLOCKID, "
'''''    sql = sql & " B.UPDDATE, "
'''''    sql = sql & " B.HOLDCLS, "
'''''    sql = sql & " H.HINBAN, "            ' 品番
'''''    sql = sql & " H.REVNUM, "            ' 製品番号改訂番号
'''''    sql = sql & " H.FACTORY, "           ' 工場
'''''    sql = sql & " H.OPECOND, "           ' 操業条件
'''''    sql = sql & " S.HSXTYPE, "           ' 品ＳＸタイプ
'''''    sql = sql & " S.HSXCDIR, "            ' 品ＳＸ結晶面方位
'''''    sql = sql & " H.INGOTPOS, "
'''''                '判定NGがあるかどうか
'''''    sql = sql & " (select count(*) from VECME010 V1 "
'''''    sql = sql & "  where V1.E040BLOCKID=B.BLOCKID "
'''''    sql = sql & "  and ((V1.E043CRYINDRS<>'0' and V1.E043CRYRESRS='2') "         ' 結晶検査実績（Rs)
'''''    sql = sql & "  or (V1.E043CRYINDOI<>'0' and V1.E043CRYRESOI='2') "         ' 結晶検査実績（Oi)
'''''    sql = sql & "  or (V1.E043CRYINDB1<>'0' and V1.E043CRYRESB1='2')"          ' 結晶検査実績（B1)
'''''    sql = sql & "  or (V1.E043CRYINDB2<>'0' and V1.E043CRYRESB2='2') "         ' 結晶検査実績（B2）
'''''    sql = sql & "  or (V1.E043CRYINDB3<>'0' and V1.E043CRYRESB3='2') "         ' 結晶検査実績（B3)
'''''    sql = sql & "  or (V1.E043CRYINDL1<>'0' and V1.E043CRYRESL1='2') "         ' 結晶検査実績（L1)
'''''    sql = sql & "  or (V1.E043CRYINDL2<>'0' and V1.E043CRYRESL2='2') "         ' 結晶検査実績（L2)
'''''    sql = sql & "  or (V1.E043CRYINDL3<>'0' and V1.E043CRYRESL3='2') "         ' 結晶検査実績（L3)
'''''    sql = sql & "  or (V1.E043CRYINDL4<>'0' and V1.E043CRYRESL4='2') "         ' 結晶検査実績（L4)
'''''    sql = sql & "  or (V1.E043CRYINDCS<>'0' and V1.E043CRYRESCS='2') "         ' 結晶検査実績（Cs)
'''''    sql = sql & "  or (V1.E043CRYINDGD<>'0' and V1.E043CRYRESGD='2') "         ' 結晶検査実績（GD)
'''''    sql = sql & "  or (V1.E043CRYINDT<>'0' and V1.E043CRYREST='2') "           ' 結晶検査実績（T)
'''''    sql = sql & "  or (V1.E043CRYINDEP<>'0' and V1.E043CRYRESEP='2')) ) as J "         ' 結晶検査実績（EPD)
'''''    sql = sql & " from "
'''''    sql = sql & " TBCME040 B, TBCME041 H, TBCME018 S"
'''''    sql = sql & " where "
'''''    sql = sql & " B.CRYNUM = H.CRYNUM "
'''''    sql = sql & " and H.HINBAN = S.HINBAN "
'''''    sql = sql & " and H.REVNUM = S.MNOREVNO "
'''''    sql = sql & " and H.FACTORY = S.FACTORY "
'''''    sql = sql & " and H.OPECOND = S.OPECOND "
'''''
'''''                '工程コード、状態、区分の条件指定
'''''    sql = sql & " and B.NOWPROC='CC600' "
'''''    sql = sql & " and B.LSTATCLS='T' "
'''''    sql = sql & " and B.RSTATCLS='T' "
'''''    sql = sql & " and B.DELCLS='0' "
'''''    'sql = sql & " and B.HOLDCLS='0' " ' ホールドブロックも取得
'''''                'ブロック内に含まれる品番を検索
'''''    sql = sql & " and (( B.INGOTPOS >= H.INGOTPOS "
'''''    sql = sql & " and B.INGOTPOS < H.INGOTPOS + H.LENGTH ) "
'''''    sql = sql & " or ( B.INGOTPOS + B.LENGTH > H.INGOTPOS "
'''''    sql = sql & " and B.INGOTPOS + B.LENGTH < H.INGOTPOS + H.LENGTH  ) "
'''''    sql = sql & " or ( H.INGOTPOS >= B.INGOTPOS "
'''''    sql = sql & " and H.INGOTPOS < B.INGOTPOS + B.LENGTH ) "
'''''    sql = sql & " or ( H.INGOTPOS + H.LENGTH > B.INGOTPOS "
'''''    sql = sql & " and H.INGOTPOS + H.LENGTH < B.INGOTPOS + B.LENGTH )) "
'''''                '指示が0でなく実績が0でないサンプルが上下２枚あるか
'''''    sql = sql & " and 2=( select count(*) "
'''''    sql = sql & "  from VECME010 V2 "
'''''    sql = sql & "  where "
'''''    sql = sql & "  B.BLOCKID=V2.E040BLOCKID"
'''''    sql = sql & "  and (V2.E043CRYINDRS='0' or V2.E043CRYRESRS<>'0') "         ' 結晶検査実績（Rs)
'''''    sql = sql & "  and (V2.E043CRYINDOI='0' or V2.E043CRYRESOI<>'0') "         ' 結晶検査実績（Oi)
'''''    sql = sql & "  and (V2.E043CRYINDB1='0' or V2.E043CRYRESB1<>'0')"          ' 結晶検査実績（B1)
'''''    sql = sql & "  and (V2.E043CRYINDB2='0' or V2.E043CRYRESB2<>'0') "         ' 結晶検査実績（B2）
'''''    sql = sql & "  and (V2.E043CRYINDB3='0' or V2.E043CRYRESB3<>'0') "         ' 結晶検査実績（B3)
'''''    sql = sql & "  and (V2.E043CRYINDL1='0' or V2.E043CRYRESL1<>'0') "         ' 結晶検査実績（L1)
'''''    sql = sql & "  and (V2.E043CRYINDL2='0' or V2.E043CRYRESL2<>'0') "         ' 結晶検査実績（L2)
'''''    sql = sql & "  and (V2.E043CRYINDL3='0' or V2.E043CRYRESL3<>'0') "         ' 結晶検査実績（L3)
'''''    sql = sql & "  and (V2.E043CRYINDL4='0' or V2.E043CRYRESL4<>'0') "         ' 結晶検査実績（L4)
'''''    sql = sql & "  and (V2.E043CRYINDCS='0' or V2.E043CRYRESCS<>'0') "         ' 結晶検査実績（Cs)
'''''    sql = sql & "  and (V2.E043CRYINDGD='0' or V2.E043CRYRESGD<>'0') "         ' 結晶検査実績（GD)
'''''    sql = sql & "  and (V2.E043CRYINDT='0' or V2.E043CRYREST<>'0') "           ' 結晶検査実績（T)
'''''    sql = sql & "  and (V2.E043CRYINDEP='0' or V2.E043CRYRESEP<>'0') )"         ' 結晶検査実績（EPD)
'''''    sql = sql & " order by B.BLOCKID, H.INGOTPOS "
'''''
'''''    'データを抽出する
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''    'レコード0件時は正常
'''''    If rs.RecordCount = 0 Then
'''''        rs.Close
'''''        ReDim records(0)
'''''    Else
'''''        BlockIdBuf = vbNullString
'''''        recCnt = rs.RecordCount
'''''        j = 0
'''''        For i = 1 To recCnt
'''''            DoEvents
'''''        'ブロックID等の格納
'''''            If rs("BLOCKID") <> BlockIdBuf Then
'''''
'''''                j = j + 1
'''''                ReDim Preserve records(j)
'''''
'''''                With records(j)
'''''                    .CRYNUM = rs("CRYNUM")
'''''                    .INGOTPOS = rs("ss")
''''''                    .LENGTH = rs("LENGTH")      ' 長さ
'''''                    .BLOCKID = rs("BLOCKID")   ' ブロックID
'''''                    .UPDDATE = rs("UPDDATE")   ' 更新日付
'''''                    .HOLDCLS = rs("HOLDCLS")   ' ホールド区分
'''''                    BlockIdBuf = records(j).BLOCKID
'''''                    .HSXTYPE = rs("HSXTYPE")
'''''                    .HSXCDIR = rs("HSXCDIR")
'''''                    If rs("J") > 0 Then
'''''
'''''                        .Judg = "2"
'''''                    Else
'''''                        .Judg = "1"
'''''                    End If
'''''
'''''                End With
'''''                k = 1
'''''            End If
'''''
'''''            '品番の格納
'''''            ReDim Preserve records(j).HIN(k)
'''''            records(j).HIN(k).hinban = rs("HINBAN")
'''''            records(j).HIN(k).mnorevno = rs("REVNUM")
'''''            records(j).HIN(k).factory = rs("FACTORY")
'''''            records(j).HIN(k).opecond = rs("OPECOND")
'''''            k = k + 1
'''''            rs.MoveNext
'''''        Next i
'''''        rs.Close
'''''
'''''    End If
'''''
'''''
'''''    '購入単結晶実績取得
'''''    If getKouBlock(records(), "CC600") = FUNCTION_RETURN_FAILURE Then
'''''       DBDRV_scmzc_fcmkc001b_Disp2 = FUNCTION_RETURN_FAILURE
'''''       GoTo PROC_EXIT
'''''    End If
'''''
'''''
'''''PROC_EXIT:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''PROC_ERR:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    DBDRV_scmzc_fcmkc001b_Disp2 = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume PROC_EXIT
'''''End Function



'概要    :待ち一覧 初期表示用ＤＢドライバ（払出待ち）
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO  ,型                                 ,説明
'        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,初期表示用
'        :戻ﾘ値        ,O   ,FUNCTION_RETURN                    ,読み込み成否
'説明    :
'履歴    :2001/07/06 蔵本 作成
Public Function DBDRV_scmzc_fcmkc001b_Disp3(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN

    '＜払出待ち＞
    'CC700のもの
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp3"


    DBDRV_scmzc_fcmkc001b_Disp3 = FUNCTION_RETURN_SUCCESS
    
    'ブロックID､更新日付、品番等取得
    If getBlockID(records(), "CC700") = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmkc001b_Disp3 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If


'    '購入単結晶実績取得
'    If getKouBlock(records(), "CC700") = FUNCTION_RETURN_FAILURE Then
'       DBDRV_scmzc_fcmkc001b_Disp3 = FUNCTION_RETURN_FAILURE
'       GoTo proc_exit
'    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    DBDRV_scmzc_fcmkc001b_Disp3 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function



''''''概要    :待ち一覧 初期表示用ＤＢドライバ（抜試指示待ち）
''''''ﾊﾟﾗﾒｰﾀ  :変数名       ,IO  ,型                                 ,説明
''''''        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,初期表示用
''''''        :戻ﾘ値        ,O   ,FUNCTION_RETURN                    ,読み込み成否
''''''説明    :
''''''履歴    :2001/07/06 蔵本 作成
'''''Public Function DBDRV_scmzc_fcmkc001b_Disp4(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
'''''
'''''    '＜抜試指示待ち＞
'''''    'CC710のもの
'''''
'''''    'ブロックID､更新日付取得
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo PROC_ERR
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp4"
'''''
'''''    DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_SUCCESS
'''''
'''''
'''''    'ブロックID､更新日付、品番等取得
'''''    If getBlockID(records(), "CC710") = FUNCTION_RETURN_FAILURE Then
'''''        DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_FAILURE
'''''        GoTo PROC_EXIT
'''''    End If
'''''
''''''2000/08/24 S.Sano Start
''''''    '購入単結晶実績取得
''''''    If getKouBlock(records(), "CC710") = FUNCTION_RETURN_FAILURE Then
''''''       DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_FAILURE
''''''       GoTo proc_exit
''''''    End If
''''''2000/08/24 S.Sano End
'''''
'''''
'''''PROC_EXIT:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''PROC_ERR:
'''''    'エラーハンドラ
'''''    gErr.HandleError
'''''    DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_FAILURE
'''''    Resume PROC_EXIT
'''''End Function



''''''購入単結晶用
'''''Private Function getKouBlock(records() As type_DBDRV_scmzc_fcmkc001b_Disp, NOWPROC As String) As FUNCTION_RETURN
'''''
'''''    Dim sql As String       'SQL全体
'''''    Dim rs As OraDynaset    'RecordSet
'''''    Dim recCnt As Long
'''''    Dim motoRecCnt As Long
'''''    Dim i As Long
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo PROC_ERR
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function getKouBlock"
'''''
'''''    getKouBlock = FUNCTION_RETURN_SUCCESS
'''''
'''''    sql = " select "
'''''    sql = sql & " B.BLOCKID, "
'''''    sql = sql & " B.UPDDATE, "
'''''    sql = sql & " B.HOLDCLS, "
'''''    sql = sql & " K.HINBAN, "
'''''    sql = sql & " K.MNOREVNO, "
'''''    sql = sql & " K.FACTORY, "
'''''    sql = sql & " K.OPECOND "
'''''    sql = sql & " from TBCME040 B,TBCMG002 K "
'''''    sql = sql & " where B.BLOCKID=K.CRYNUM "
'''''    sql = sql & " and substr(B.BLOCKID,1,1)='8' "
'''''    sql = sql & " and B.NOWPROC='" & NOWPROC & "' "
'''''    sql = sql & " and B.LSTATCLS='T' "
'''''    sql = sql & " and B.RSTATCLS='T' "
'''''    sql = sql & " and B.DELCLS='0' "
'''''    'sql = sql & " and B.HOLDCLS='0' " ' ホールドブロックも取得
'''''    sql = sql & " and K.TRANCNT=any(select max(TRANCNT) from TBCMG002 where CRYNUM=B.BLOCKID ) "
'''''    sql = sql & " order by B.BLOCKID "
'''''
'''''
'''''    'データを抽出する
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''    If rs.RecordCount = 0 Then
'''''        rs.Close
'''''        GoTo PROC_EXIT
'''''    End If
'''''
'''''    motoRecCnt = UBound(records)
'''''    recCnt = rs.RecordCount
'''''    ReDim Preserve records(UBound(records) + recCnt)
'''''
'''''    For i = motoRecCnt + 1 To UBound(records)
'''''        DoEvents
'''''        ReDim records(i).HIN(1)
'''''        With records(i)
'''''            .BLOCKID = rs("BLOCKID")     ' ブロックID
'''''            .UPDDATE = rs("UPDDATE")     ' 更新日付
'''''            .HOLDCLS = rs("HOLDCLS")     ' ホールド区分
'''''            .HIN(1).hinban = rs("HINBAN")       ' 品番
'''''            .HIN(1).mnorevno = rs("MNOREVNO")   ' 製品番号改訂番号
'''''            .HIN(1).factory = rs("FACTORY")     ' 工場
'''''            .HIN(1).opecond = rs("OPECOND")     ' 操業条件
'''''        End With
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''
'''''PROC_EXIT:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''PROC_ERR:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    getKouBlock = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume PROC_EXIT
'''''
'''''End Function



'内部関数 ブロックID、更新日付取得（払出待ち、抜試指示待ち用）
Private Function getBlockID(records() As type_DBDRV_scmzc_fcmkc001b_Disp, _
                            NOWPROC As String) As FUNCTION_RETURN

    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim recCnt      As Long         'レコード数
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim BlockIdBuf  As String
    Dim lLp         As Long         '2007/08/30 SPK Tsutsumi Add
    Dim sBakPos     As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function getBlockID"

    getBlockID = FUNCTION_RETURN_SUCCESS

    'sql = "select "
    'sql = sql & " X.INPOSC2, "
    'sql = sql & " X.GNLC2, "
    'sql = sql & " X.HOLDBC2, "      '2005/08
    'sql = sql & " X.HOLDCC2, "      '2005/08
    'sql = sql & " X.HOLDKTC2, "      '2005/08
    'sql = sql & " X.LBLFLGC2,"       '2005/11
    'sql = sql & " V.E040CRYNUM, "
    'sql = sql & " V.E040BLOCKID, "
    'sql = sql & " V.E040INGOTPOS, "
    'sql = sql & " V.E040UPDDATE, "
    'sql = sql & " V.E040HOLDCLS, "
    'sql = sql & " V.E041HINBAN, "            ' 品番
    'sql = sql & " V.E041REVNUM, "            ' 製品番号改訂番号
    'sql = sql & " V.E041FACTORY, "           ' 工場
    'sql = sql & " V.E041OPECOND, "           ' 操業条件
    'sql = sql & " V.E041LENGTH, "            ' 長さ  2006/02
    'sql = sql & " V.E037DIAMETER, "          ' 直径  2006/02
    'sql = sql & " S.HSXTYPE, "           ' 品ＳＸタイプ
    'sql = sql & " S.HSXCDIR "            ' 品ＳＸ結晶面方位
    'sql = sql & ",XC1.PUPTNC1 "          ' 引上ﾊﾟﾀｰﾝ(2004/12/21) kubota
    'sql = sql & " from "
    ''sql = sql & " VECME009 V, TBCME018 S "
    'sql = sql & " VECME009 V, TBCME018 S , XSDC2 X "
    
    'sql = sql & ",XSDC1 XC1 "                       '引上ﾊﾟﾀｰﾝ追加対応(2004/12/21) kubota
    
    'sql = sql & " where "
    'sql = sql & " V.E040BLOCKID = X.CRYNUMC2 "
    'sql = sql & " and V.E041HINBAN = S.HINBAN "
    'sql = sql & " and V.E041REVNUM = S.MNOREVNO "
    'sql = sql & " and V.E041FACTORY = S.FACTORY "
    'sql = sql & " and V.E041OPECOND = S.OPECOND "
    'sql = sql & " and V.E040NOWPROC='" & NOWPROC & "' "
    'sql = sql & " and V.E040LSTATCLS='T' "
    'sql = sql & " and V.E040RSTATCLS='T' "
    'sql = sql & " and V.E040DELCLS='0' "
    ''sql = sql & " and V.E040HOLDCLS='0' " ' ホールドブロックも取得
    'sql = sql & " and X.XTALC2 = XC1.XTALC1(+) "    '引上ﾊﾟﾀｰﾝ追加対応(2004/12/21) kubota
    'sql = sql & " order by V.E040BLOCKID, V.E041INGOTPOS "
    
    'VIEW --> CA 変更  2006/02
    sql = "select "
    sql = sql & " X.XTALC2, "
    sql = sql & " X.INPOSC2, "
    sql = sql & " X.GNLC2, "
    sql = sql & " X.HOLDBC2, "      '2005/08
    sql = sql & " X.HOLDCC2, "      '2005/08
    sql = sql & " X.HOLDKTC2, "      '2005/08
    sql = sql & " X.LBLFLGC2,"       '2005/11
    sql = sql & " CA.CRYNUMCA, "
    sql = sql & " CA.INPOSCA, "
    sql = sql & " CA.KDAYCA, "
    sql = sql & " CA.HOLDBCA, "
    sql = sql & " CA.HINBCA, "           ' 品番
    sql = sql & " CA.REVNUMCA, "         ' 製品番号改訂番号
    sql = sql & " CA.FACTORYCA, "        ' 工場
    sql = sql & " CA.OPECA, "            ' 操業条件
    sql = sql & " CA.GNLCA, "            ' 長さ  2006/02
    sql = sql & " CA.GNWCA, "            ' 重量  2006/02
    sql = sql & " S.HSXTYPE, "           ' 品ＳＸタイプ
    sql = sql & " S.HSXCDIR "            ' 品ＳＸ結晶面方位
    sql = sql & ",XC1.PUPTNC1 "          ' 引上ﾊﾟﾀｰﾝ(2004/12/21) kubota
    sql = sql & ",X.KIKBNC2 "            ' 期判別区分 2006/11/14 SETsw kubota
    sql = sql & ",X.PLANTCATC2 "         ' 向先 2007/08/30 SPK Tsutsumi Add
    ' 流動監視SQL修正 upd SETkimizuka Start  09/07/01
    '' 流動停止項目追加 add SETkimizuka Start  09/03/26
    'sql = sql & " , NVL(TO_CHAR(Y4.AGRSTATUS),' ') as AGRSTATUS "
    'sql = sql & " , NVL(TO_CHAR(Y4.STOP),'0') as STOP "
    'sql = sql & " , NVL(Y4.CAUSE,' ') as CAUSE "
    'sql = sql & " , NVL(Y4.PRINTKIND || Y4.PRINTNO,' ') as PRINTNO "
    '' 流動停止項目追加 add SETkimizuka End    09/03/26
    sql = sql & " , NVL(TO_CHAR(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)),' ') as AGRSTATUS "
    sql = sql & " , NVL(TO_CHAR(Y4.STOPY4),' ') as STOP "
    sql = sql & " , DECODE(TRIM(Y4.CAUSEY4),NULL,' ',TRIM(Y4.CAUSEY4) || ':' || A9.NAMEJA9) as CAUSE "
    sql = sql & " , NVL(Y4.PRINTKINDY4 || Y4.PRINTNOY4,' ') as PRINTNO "
    sql = sql & " , NVL(Y4.WKKTY4,'0') as WKKTY4 "
    ' 流動監視SQL修正 upd SETkimizuka End  09/07/01
    sql = sql & " from "
    'sql = sql & " VECME009 V, TBCME018 S , XSDC2 X "
    sql = sql & " XSDCA CA, TBCME018 S , XSDC2 X "
    
    sql = sql & ",XSDC1 XC1 "                       '引上ﾊﾟﾀｰﾝ追加対応(2004/12/21) kubota
    
    ' 流動監視SQL修正 upd SETkimizuka Start  09/07/01
    sql = sql & "    ,XODY3 Y3,XODY4 Y4,KODA9 A9 "
    '' 流動停止項目追加 add SETkimizuka Start  09/03/26
    'sql = sql & "    ,(SELECT XTALNOY3 as XTALNO,MIN(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",AGRSTATUSY4)) as AGRSTATUS  "
    'sql = sql & "      ,MAX(STOPY4) as STOP,DECODE(TRIM(CAUSEY4),'',TRIM(CAUSEY4),TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSE ,Y5.PRINTNO,Y5.PRINTKIND "
    'sql = sql & "      FROM XODY3  "
    'sql = sql & "           LEFT OUTER JOIN XODY4 on ( XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND STOPY4 <> '2'  AND WKKTY4 in " & CreateWkktSQL(WATCH_PROCCD) & ") "
    'sql = sql & "           LEFT OUTER JOIN KODA9 on ( SYSCA9 = 'X' AND SHUCA9 = '30' AND CAUSEY4 = CODEA9 ) "
    'sql = sql & "           LEFT OUTER JOIN (SELECT XTALNOY4 as XTALNO,SXLIDY4 as SXLID,PRINTNOY5 as PRINTNO,PRINTKINDY5 as PRINTKIND "
    'sql = sql & "                FROM XODY3,XODY4,XODY5 "
    'sql = sql & "              WHERE XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND LIVKY3 = '0' "
    'sql = sql & "                AND PRINTKINDY4 = PRINTKINDY5 AND PRINTNOY4 = PRINTNOY5  "
    'sql = sql & "                AND HKBNY5 ='0' GROUP BY XTALNOY4,SXLIDY4,PRINTNOY5,PRINTKINDY5) Y5 ON (XTALNOY3 = XTALNO AND SXLIDY3 = SXLID ) "
    'sql = sql & "      WHERE  "
    'sql = sql & "       LIVKY3    = '0' "
    'sql = sql & "       GROUP BY XTALNOY3,AGRSTATUSY4,CAUSEY4,Y5.PRINTNO,Y5.PRINTKIND,NAMEJA9) Y4 "
    '' 流動停止項目追加 add SETkimizuka End  09/03/26
    ' 流動監視SQL修正 upd SETkimizuka End  09/07/01
    sql = sql & " where "
    sql = sql & " CA.CRYNUMCA = X.CRYNUMC2 "
    sql = sql & " and CA.HINBCA = S.HINBAN "
    sql = sql & " and CA.REVNUMCA = S.MNOREVNO "
    sql = sql & " and CA.FACTORYCA = S.FACTORY "
    sql = sql & " and CA.OPECA = S.OPECOND "
    sql = sql & " and CA.GNWKNTCA ='" & NOWPROC & "' "
    sql = sql & " and CA.LIVKCA='0' "
    sql = sql & " and X.XTALC2 = XC1.XTALC1(+) "    '引上ﾊﾟﾀｰﾝ追加対応(2004/12/21) kubota
    ' 流動監視SQL修正 upd SETkimizuka Start  09/07/01
    'sql = sql & " AND CA.CRYNUMCA    = Y4.XTALNO(+) "            'add 09/03/26 SETkimizuka
    sql = sql & " AND CA.CRYNUMCA = Y3.XTALNOY3(+) "
    sql = sql & " AND Y3.LIVKY3(+) = '0' "
    sql = sql & " AND Y4.LIVKY4(+) = '0' "
    sql = sql & " AND Y3.XTALNOY3 = Y4.XTALNOY4(+) "
    sql = sql & " AND Y3.RCNTY3 = Y4.RCNTY4(+) "
    sql = sql & " AND A9.SYSCA9(+) = 'X' AND A9.SHUCA9(+) = '30' AND Y4.CAUSEY4 = A9.CODEA9(+) "
    ' 流動監視SQL修正 upd SETkimizuka End  09/07/01

    sql = sql & " order by CA.CRYNUMCA, CA.INPOSCA "

    'データを抽出する
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    'レコードがない場合正常終了
    If rs.RecordCount = 0 Then
        rs.Close
        ReDim records(0)
        GoTo proc_exit
    End If
    
    BlockIdBuf = vbNullString
    recCnt = rs.RecordCount
    j = 0
    For i = 1 To recCnt
        DoEvents
        'ブロックID等の格納
        If rs("CRYNUMCA") <> BlockIdBuf Then
        
            j = j + 1
            ReDim Preserve records(j)
            
            With records(j)
                '.CRYNUM = rs("E040CRYNUM")
                '.INGOTPOS = rs("E040INGOTPOS")
                '.BLOCKID = rs("E040BLOCKID")   ' ブロックID
                '.UPDDATE = rs("E040UPDDATE")   ' 更新日付
                '.HOLDCLS = rs("E040HOLDCLS")   ' ホールド区分
                .CRYNUM = rs("XTALC2")
                .INGOTPOS = rs("INPOSCA")
                .BLOCKID = rs("CRYNUMCA")   ' ブロックID
                .UPDDATE = rs("KDAYCA")   ' 更新日付
                .HOLDCLS = rs("HOLDBCA")   ' ホールド区分
                BlockIdBuf = records(j).BLOCKID
                .HSXTYPE = rs("HSXTYPE")
                .HSXCDIR = rs("HSXCDIR")
                .Judg = " "
                .INPOS = rs("INPOSC2")
                .LENGTH = rs("GNLC2")
                .PUPTN = rs("PUPTNC1")         ' 引上ﾊﾟﾀｰﾝ(2004/12/21) kubota
                .HOLDB = rs("HOLDBC2")
                .HOLDC = rs("HOLDCC2")
                If IsNull(rs("HOLDKTC2")) = False Then .HOLDKT = rs("HOLDKTC2")    '2005/08
                If IsNull(rs("LBLFLGC2")) = False Then .LBLFLG = rs("LBLFLGC2")    '2005/11
                '.DIA = rs("E037DIAMETER")   '2006/02
                If IsNull(rs("KIKBNC2")) = False Then .KIKBN = rs("KIKBNC2")       ' 期判別区分 2006/11/14 SETsw kubota
            End With
            
            k = 1
            sBakPos = ""    'add 09/03/26 SETkimizuka
        End If
        
        
        If InStr(sBakPos, Trim(rs("INPOSCA"))) = 0 Then '流動監視項目追加に伴う修正 add
            '品番の格納
            ReDim Preserve records(j).hinM(k)
            records(j).hinM(k).HIN.hinban = rs("HINBCA")
            records(j).hinM(k).HIN.mnorevno = rs("REVNUMCA")
            records(j).hinM(k).HIN.factory = rs("FACTORYCA")
            records(j).hinM(k).HIN.opecond = rs("OPECA")
            records(j).hinM(k).LENGTH = rs("GNLCA")
            records(j).hinM(k).Weight = rs("GNWCA")
            
            ' 向先 2007/08/30 SPK Tsutsumi Start
            If IsNull(rs("PLANTCATC2")) = False Then
                For lLp = 0 To UBound(s_Mukesaki)
                    If rs("PLANTCATC2") = s_Mukesaki(lLp).sMukeCode Then
                        records(j).hinM(k).HIN.sMukesaki = s_Mukesaki(lLp).sMukeName
                    End If
                Next lLp
            End If
            sBakPos = sBakPos & Trim(rs("INPOSCA")) & " "
            k = k + 1
        End If
'        '品番の格納
'        ReDim Preserve records(j).hinM(k)
'        records(j).hinM(k).HIN.hinban = rs("HINBCA")
'        records(j).hinM(k).HIN.mnorevno = rs("REVNUMCA")
'        records(j).hinM(k).HIN.Factory = rs("FACTORYCA")
'        records(j).hinM(k).HIN.OpeCond = rs("OPECA")
'        records(j).hinM(k).LENGTH = rs("GNLCA")
'        records(j).hinM(k).Weight = rs("GNWCA")
        
        ' 向先 2007/08/30 SPK Tsutsumi End

        ' 流動監視SQL修正 upd SETkimizuka Start  09/07/01
        ' 流動停止項目追加 add SETkimizuka Start  09/03/26
        'records(j).STOP = rs("STOP")                   '停止区分
        'records(j).AGRSTATUS = rs("AGRSTATUS")       '承認確認区分
        'If Trim(rs("CAUSE")) <> "" And InStr(records(j).CAUSE, Trim(rs("CAUSE"))) = 0 Then
        '    records(j).CAUSE = records(j).CAUSE & rs("CAUSE") & vbTab       '停止理由
        'End If
        If rs("STOP") <> "2" And rs("WKKTY4") = "CC700" Then
           If Trim(records(j).AGRSTATUS) = "" Or (rs("AGRSTATUS") < records(j).AGRSTATUS) Then
                records(j).STOP = rs("STOP")                   '停止区分
                records(j).AGRSTATUS = rs("AGRSTATUS")       '承認確認区分
           End If
            If Trim(rs("CAUSE")) <> "" And InStr(records(j).CAUSE, Trim(rs("CAUSE"))) = 0 Then
                records(j).CAUSE = records(j).CAUSE & rs("CAUSE") & vbTab       '停止理由
            End If
        End If
        If Trim(rs("PRINTNO")) <> "" And InStr(records(j).PRINTNO, Trim(rs("PRINTNO"))) = 0 Then
            records(j).PRINTNO = records(j).PRINTNO & rs("PRINTNO") & vbTab       '先行評価
        End If
        ' 流動停止項目追加 add SETkimizuka End    09/03/26

'        k = k + 1
        rs.MoveNext
    Next i
    rs.Close
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getBlockID = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'---- ADD [精製原料管理、原料工程実績作成処理追加] 以下追加関数　START ---- TCS)T.TERAUCHI

'概要      :精製原料管理作成用
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                           ,説明
'          :Xodcx         ,I  ,type_DBDRV_fcmkc001c_InsXodcx,ブロック管理の現在管理工程、現在工程、最終通過管理工程、最終通過工程更新用
'          :戻り値        ,O  ,FUNCTION_RETURN              ,
'説明      :
Public Function DBDRV_fcmkc001c_InsXODCX(Xodcx As type_DBDRV_fcmkc001c_InsXodcx) As FUNCTION_RETURN

    Dim sSql        As String
    Dim objDS       As Object
    Dim sDbName     As String
    Dim sErrMsg     As String
    Dim dCyokkei    As Double
    Dim sDopType    As String
    Dim sCSDop      As String       'CSドープ有無
    Dim sNDop       As String       '窒素ドープ有無
    Dim sLTUmu      As String
    Dim sSCNTRL     As String       '識別ｺﾝﾄﾛｰﾙｺｰﾄﾞ ADD 2011/03/24 TSMC品識別対応
    
'*** UPDATE START TAGAWA 2004/12/16
    Dim sFlag       As String
'*** UPDATE END  TAGAWA 2004/12/16

'エラーハンドラの設定
On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function DBDRV_fcmkc001c_InsXODCX"
    
    DBDRV_fcmkc001c_InsXODCX = FUNCTION_RETURN_FAILURE
    
    '****** 登録情報の取得 ******
    '' 精製原料基本情報取得SQLの作成
    sDbName = "XSDC1"
    Call GetAssistSQL_300(sSql, Xodcx.CRYNUM)
    If DynSet2(objDS, sSql) = False Then
        If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
        sErrMsg = GetMsgStr("EGET", sDbName)
        Exit Function
    End If
    '該当データ無しの場合
    If objDS.EOF = True Then
        If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
        sErrMsg = GetMsgStr("ESM04", sDbName)
        Exit Function
    End If
    
    '' 重量取得
    dCyokkei = objDS.Fields("PRODMCX").Value                'MID引上直径
    Xodcx.Weight = WeightOfCylinder(dCyokkei, Xodcx.LENGTH) 'ブロック重量
    
    '' CSドープ有無、窒素ドープ有無設定
    sDopType = UCase(NulltoStr(objDS.Fields("DTYPEC1").Value))
'*** UPDATE START TAGAWA 2004/12/16**************************
''    If sDopType = " " Or sDopType = "P" Then
''        sCSDop = "2"
''        sNDop = "1"
''    ElseIf sDopType = "N" Then
''        sCSDop = "1"
''        sNDop = "2"
''    Else
''        sCSDop = " "
''        sNDop = " "
''    End If
        ''結晶ドープ取得
        sFlag = UCase(Trim(NulltoStr(objDS.Fields("DPNTCLS").Value)))
        ''Csドープの時
        If sFlag = "C" Then
            sCSDop = "2"
            sNDop = "1"
        ''窒素ドープの時
        ElseIf sFlag = "N" Then
            sCSDop = "1"
            sNDop = "2"
        ''Ｗドープの時
        ElseIf sFlag = "M" Then
            sCSDop = "2"
            sNDop = "2"
        ''その他
        Else
            sCSDop = "1"
            sNDop = "1"
        End If
'*** UPDATE END TAGAWA 2004/12/16**************************

    ''ライフタイム仕様有無
    If objDS.Fields("HSXLTHWS").Value = "H" Then
        ''有り
        sLTUmu = "2"
    Else
        ''無し
        sLTUmu = "1"
    End If
    '*** UPDATE START Marushita 2011/03/24 TSMC品識別対応
    ''精製原料チェックフラグの判断
    If NulltoStr(objDS.Fields("MTRLCHKFLG").Value) = "1" Then
        ''品番NULL時の識別コントロールコードセット(空白3桁)
        If NulltoStr(objDS.Fields("HINBCX").Value) = "" Then
            sSCNTRL = "   "
        Else
            ''識別コントロールコードセット(品番3桁)
            sSCNTRL = left(objDS.Fields("HINBCX").Value, 3)
        End If
    Else
        ''識別コントロールコードセット(空白3桁)
        sSCNTRL = "   "
    End If
    '*** UPDATE END   Marushita 2011/03/24
    
    '****** 精製原料管理作成処理 ******
    sSql = ""
    sSql = sSql & "INSERT INTO xodcx(" & vbLf
    sSql = sSql & "crynumcx" & vbLf    ''ブロックID
    sSql = sSql & ",mtrlnumcx" & vbLf   ''原料№
    sSql = sSql & ",wkktcx" & vbLf      ''工程コード
    sSql = sSql & ",workcx" & vbLf      ''工場コード
    sSql = sSql & ",hdaycx" & vbLf      ''発生日時
    sSql = sSql & ",weightcx" & vbLf    ''重量
    sSql = sSql & ",htkbncx" & vbLf     ''廃棄/適合区分
    sSql = sSql & ",divumucx" & vbLf    ''分割有無
    sSql = sSql & ",toworkcx" & vbLf    ''払出先工場コード
    sSql = sSql & ",frworkcx" & vbLf    ''発生工場コード
    sSql = sSql & ",hinbcx" & vbLf      ''品番
    sSql = sSql & ",typecx" & vbLf      ''タイプ
    sSql = sSql & ",dptypecx" & vbLf    ''ドープタイプ
    sSql = sSql & ",tposcx" & vbLf      ''位置L(トップ側)
    sSql = sSql & ",lencx" & vbLf       ''ブロック長さ
    sSql = sSql & ",siweightcx" & vbLf  ''仕込み重量
    sSql = sSql & ",updmcx" & vbLf      ''引上AV径
    sSql = sSql & ",prodmcx" & vbLf     ''製品径
    sSql = sSql & ",tdopposcx" & vbLf   ''追加ドープ投入位置L
    sSql = sSql & ",wdopumucx" & vbLf   ''Wドープ(P/N混合)有無
    sSql = sSql & ",csdopumucx" & vbLf  ''CSドープ有無
    sSql = sSql & ",ndopumucx" & vbLf   ''窒素ドープ有無
    sSql = sSql & ",ltspecumucx" & vbLf ''ライフタイム仕様有無
    sSql = sSql & ",csspecumucx" & vbLf ''CS仕様有無
    sSql = sSql & ",topwcx" & vbLf      ''トップWT
    sSql = sSql & ",dmkcx" & vbLf       ''直径区分
    sSql = sSql & ",xtalcx" & vbLf      ''結晶番号
    sSql = sSql & ",livkcx" & vbLf      ''生死区分
    sSql = sSql & ",unifgcx" & vbLf     ''結合FLG
    sSql = sSql & ",twarifgcx" & vbLf   ''縦割FLG
    sSql = sSql & ",refusefgcx" & vbLf  ''受入可否FLG
    sSql = sSql & ",tstafidcx" & vbLf   ''登録社員ID
    sSql = sSql & ",tdaycx" & vbLf      ''登録日付
    sSql = sSql & ",kstafidcx" & vbLf   ''更新者
    sSql = sSql & ",kdaycx" & vbLf      ''更新日時
    sSql = sSql & ",crydopcx" & vbLf    ''結晶ドープ
    sSql = sSql & ",crydopvlcx" & vbLf  ''結晶ドープ量
    sSql = sSql & ",bkformcx" & vbLf    ''ブロック形状
    sSql = sSql & ",pgidcx" & vbLf      ''PG-ID
    sSql = sSql & ",blktypcx" & vbLf    ''ブロック種別
    sSql = sSql & ",tkacutwcx" & vbLf   ''Tサンプル前重量
'*** UPDATE START TAGAWA 2004/12/16***************
    sSql = sSql & ",denflgcx" & vbLf     ''電極材フラグ
'*** UPDATE END   TAGAWA 2004/12/16***************
    sSql = sSql & ",toptwcx" & vbLf     ''ﾄｯﾌﾟ取出しWT
'*** UPDATE START Marushita 2011/03/24 TSMC品識別対応
    sSql = sSql & ",scntrlcx" & vbLf    ''識別ｺﾝﾄﾛｰﾙｺｰﾄﾞ
'*** UPDATE END   Marushita 2011/03/24
    sSql = sSql & ")values(" & vbLf
    sSql = sSql & "'" & Xodcx.BLOCKID & "0" & "'" & vbLf                    ''ﾌﾞﾛｯｸID
    sSql = sSql & ",' '" & vbLf                                             ''原料番号
    sSql = sSql & ",'" & Right(PROCD_KOUNYU_TAN_KESSYOU, 4) & "'" & vbLf    ''工程コード('B410')
    sSql = sSql & ",'" & gsFactryCd & "'" & vbLf                            ''工場ｺｰﾄﾞ
    sSql = sSql & ",sysdate" & vbLf                                         ''発生日時
    sSql = sSql & "," & Xodcx.Weight & vbLf                                 ''重量
    sSql = sSql & ",'1'" & vbLf                                             ''廃棄・適合区分
    sSql = sSql & ",'1'" & vbLf                                             ''分割有無
    sSql = sSql & ",'" & gsFactryCd & "'" & vbLf                            ''払出先工場ｺｰﾄﾞ
    sSql = sSql & ",'" & gsFactryCd & "'" & vbLf                            ''発生工場ｺｰﾄﾞ
    sSql = sSql & ",'" & objDS.Fields("HINBCX").Value & "'" & vbLf          ''品番
    sSql = sSql & ",'" & objDS.Fields("HSXTYPE").Value & "'  " & vbLf       ''タイプ
    sSql = sSql & ",'" & sDopType & "'" & vbLf                              ''ドープタイプ
    sSql = sSql & ", " & Xodcx.INGOTPOS & vbLf                              ''ブロック管理･結晶内開始位置
    sSql = sSql & ", " & Xodcx.LENGTH & vbLf                                ''ブロック管理･長さ
    sSql = sSql & ", " & ConvNum(objDS.Fields("SUICHARGE").Value) & vbLf    ''仕込み重量
    sSql = sSql & ", " & ConvNum(objDS.Fields("UPDMCX").Value) & vbLf       ''引上AV径
    sSql = sSql & ", " & ConvNum(objDS.Fields("PRODMCX").Value) & vbLf      ''製品径
    sSql = sSql & ", " & ConvNum(objDS.Fields("ADDOPPC1").Value) & vbLf     ''追加ドープ投入位置L
    sSql = sSql & ",'1'" & vbLf                                             ''Wドープ(P/N混合)有無
    sSql = sSql & ",'" & sCSDop & "'" & vbLf                                ''CSドープ有無
    sSql = sSql & ",'" & sNDop & "'" & vbLf                                 ''窒素ドープ有無
    sSql = sSql & ",'" & sLTUmu & "'" & vbLf                                ''ライフタイム使用有無
    sSql = sSql & ",'2'" & vbLf                                             ''CS使用有無
    sSql = sSql & "," & ConvNum(objDS.Fields("CTR01A9").Value) & vbLf       '’トップWT
    sSql = sSql & ",'300'" & vbLf                                           ''直径区分
    sSql = sSql & ",'" & Xodcx.CRYNUM & "'" & vbLf                          ''結晶番号
    sSql = sSql & ",'0'" & vbLf                                             ''生死区分
    sSql = sSql & ",'0'" & vbLf                                             ''結合FLG
    sSql = sSql & ",'0'" & vbLf                                             ''縦割FLG
    sSql = sSql & ",'0'" & vbLf                                             ''受入可否FLG
    sSql = sSql & ",'" & Xodcx.STAFFID & "'" & vbLf                         ''登録社員ID
    sSql = sSql & ",sysdate" & vbLf                                         ''登録日付
    sSql = sSql & ",'" & Xodcx.STAFFID & "'" & vbLf                         ''更新社員ID
    sSql = sSql & ",sysdate" & vbLf                                         ''更新日付
    sSql = sSql & ",'" & objDS.Fields("DPNTCLS").Value & "'" & vbLf         ''結晶ドープ
    sSql = sSql & "," & ConvNum(objDS.Fields("DOPANT").Value) & vbLf        ''結晶ドープ量
    sSql = sSql & ",'3'" & vbLf                                             ''ブロック形状
    sSql = sSql & ",'" & objDS.Fields("PGID").Value & "'" & vbLf            ''PG-ID
    sSql = sSql & ",'A'" & vbLf                                             ''ブロック種別
    sSql = sSql & ",0" & vbLf                                               ''Tサンプル前重量
'*** UPDATE START TAGAWA 2004/12/16***************
    sSql = sSql & ",'1'" & vbLf                                             ''電極材フラグ
'*** UPDATE END   TAGAWA 2004/12/16***************
    sSql = sSql & "," & ConvNum(objDS.Fields("PUTCUTWC1").Value) & vbLf     ''ﾄｯﾌﾟ取出しWT
'*** UPDATE START Marushita 2011/03/24 TSMC品識別対応
    sSql = sSql & ",'" & sSCNTRL & "'" & vbLf                               ''識別ｺﾝﾄﾛｰﾙｺｰﾄﾞ
'*** UPDATE END   Marushita 2011/03/24
    
    sSql = sSql & ")"
        
    If 0 >= OraDB.ExecuteSQL(sSql) Then
        DBDRV_fcmkc001c_InsXODCX = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_fcmkc001c_InsXODCX = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    DBDRV_fcmkc001c_InsXODCX = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :原料工程実績作成用
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                           ,説明
'          :Xodcx         ,I   ,type_DBDRV_scmzc_fcmkc001c_InsXodcx         ,ブロック管理の現在管理工程、現在工程、最終通過管理工程、最終通過工程更新用
'          :戻り値        ,O  ,FUNCTION_RETURN              ,
'説明      :
'履歴      :2004/12/04 新規作成 TCS)T.TERAUCHI
Public Function DBDRV_fcmkc001c_InsXODB3(Xodcx As type_DBDRV_fcmkc001c_InsXodcx) As FUNCTION_RETURN
    Dim objDS       As Object
    Dim sSql        As String
    Dim iRenban     As Integer
    Dim sYear       As String
    Dim sMonth      As String
    Dim sDay        As String
    Dim sHour       As String
    Dim sMin        As String
    Dim sNowdate    As String
    Dim sCyoku      As String

'' エラーハンドラの設定
On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function DBDRV_fcmkc001c_InsXODB3"

    DBDRV_fcmkc001c_InsXODB3 = FUNCTION_RETURN_FAILURE
    
    '***** 登録情報の取得 *****
    '' 工程コード設定
    nowCd = Xodcx.LASTPASS
    
    '' システム日付、実績日付等の設定
    If Not GetSysdate Then
        GoTo proc_exit
    End If
    sNowdate = gsSysdate
    
    'サーバーシステム日付を実績日に変更
    sNowdate = GetJITUDATE(Format(sNowdate, "yyyymmddhhmmss"))
    
    '実績日より直区分を判定
    sCyoku = GetCYOKU(gsSysdate)
    
    '実績日から切り取り
    sYear = Mid(sNowdate, 1, 4)     '年
    sMonth = Mid(sNowdate, 5, 2)    '月
    sDay = Mid(sNowdate, 7, 2)      '日
    sHour = Mid(sNowdate, 9, 2)     '時
    sMin = Mid(sNowdate, 11, 2)     '分

    iRenban = 0

    '' 工程連番の取得
    sSql = ""
    sSql = sSql & " SELECT NVL(MAX(kcntb3),0) maxcnt        " & vbLf   '工程連番
    sSql = sSql & " FROM   xodb3                            " & vbLf
    sSql = sSql & " WHERE  polnob3 = '" & Xodcx.BLOCKID & "0" & "'" & vbLf
    
    'SQL文実行
    If DynSet2(objDS, sSql) = False Then
        If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
        GoTo proc_exit
    End If
    
    '取得したデータを格納
    iRenban = objDS.Fields("maxcnt").Value + 1
    If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing

    '****** 原料工程実績(XODB3)更新 ******
    sSql = ""
    sSql = sSql & "insert into XODB3(                       " & vbLf
    sSql = sSql & "             POLNOB3                     " & vbLf   '原料番号
    sSql = sSql & "            ,KCNTB3                      " & vbLf   '工程連番
    sSql = sSql & "            ,CRSEQB3                     " & vbLf   '処理連番
    sSql = sSql & "            ,TDAYB3                      " & vbLf   '登録日付
    sSql = sSql & "            ,RDAYB3                      " & vbLf   '修正日付
    sSql = sSql & "            ,SDAYB3                      " & vbLf   '送信日付
    sSql = sSql & "            ,SNDKB3                      " & vbLf   '送信区分
    sSql = sSql & "            ,SAKJB3                      " & vbLf   '削除区分
    sSql = sSql & "            ,POKUBB3                     " & vbLf   '原料区分
    sSql = sSql & "            ,POKIDCB3                    " & vbLf   '原料種類コード
    sSql = sSql & "            ,POLTNB3                     " & vbLf   '原料ロットNo
    sSql = sSql & "            ,MODKBB3                     " & vbLf   '赤黒区分
    sSql = sSql & "            ,SUMKBB3                     " & vbLf   '集計区分
    sSql = sSql & "            ,WKKTB3                      " & vbLf   '工程コード
    sSql = sSql & "            ,PLACB3                      " & vbLf   'ラインコード
    sSql = sSql & "            ,FRWB3                       " & vbLf   '受入重量
    sSql = sSql & "            ,TOWB3                       " & vbLf   '払出重量
    sSql = sSql & "            ,LOSWB3                      " & vbLf   'ロス重量
    sSql = sSql & "            ,FRWKKTB3                    " & vbLf   '受入工程コード
    sSql = sSql & "            ,TOWKKTB3                    " & vbLf   '払出工程コード
    sSql = sSql & "            ,TOWKKBB3                    " & vbLf   '払出区分
    sSql = sSql & "            ,TOWORKB3                    " & vbLf   '払出工場コード
    sSql = sSql & "            ,TOPLACB3                    " & vbLf   '払出ラインコード
    sSql = sSql & "            ,CHGNB3                      " & vbLf   'チャージNo
    sSql = sSql & "            ,EYYB3                       " & vbLf   '実績日付(年)
    sSql = sSql & "            ,EMMB3                       " & vbLf   '実績日付(月)
    sSql = sSql & "            ,EDDB3                       " & vbLf   '実績日付(日)
    sSql = sSql & "            ,ECYOKB3                     " & vbLf   '直区分
    sSql = sSql & "            ,EHHB3                       " & vbLf   '実績時間(時)
    sSql = sSql & "            ,EMIB3　                     " & vbLf   '実績時間(分)
    sSql = sSql & "            ,MANB3                       " & vbLf   '担当者
    sSql = sSql & "            ,MANJB3                      " & vbLf   '担当者名
    sSql = sSql & "            ,DENKB3                      " & vbLf   '濃度区分
    sSql = sSql & "            ,DENSITYB3                   " & vbLf   '濃度値
    sSql = sSql & "            ,GSNDFLGB3                   " & vbLf   '原料送信フラグ
    sSql = sSql & "            ,HFLGB3                      " & vbLf   '発生フラグ
    sSql = sSql & "            ,htkbnb3                     " & vbLf   '現品区分
    sSql = sSql & "            ,plworkb3                    " & vbLf   '使用予定工場
    sSql = sSql & "            ,mdensityb3                  " & vbLf   '元濃度値
    sSql = sSql & "            ,gsdayb3                     " & vbLf
    sSql = sSql & ")VALUES(                                 " & vbLf
    sSql = sSql & " '" & Xodcx.BLOCKID & "0" & "'                 " & vbLf  ' 原料番号
    sSql = sSql & "," & iRenban & "                         " & vbLf   '工程連番
    sSql = sSql & ",1                                       " & vbLf   '処理連番
    sSql = sSql & ",to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss') " & vbLf '登録日付
    sSql = sSql & ",null                                    " & vbLf   '修正日付
    sSql = sSql & ",null                                    " & vbLf   '送信日付
    sSql = sSql & ",' '                                     " & vbLf   '送信区分
    sSql = sSql & ",'0'                                     " & vbLf   '削除区分
    sSql = sSql & ",'2'                                     " & vbLf   '原料区分
    sSql = sSql & ",'888'                                   " & vbLf   '原料種類コード
    sSql = sSql & ",' '                                     " & vbLf   '原料ロット番号
    sSql = sSql & ",' '                                     " & vbLf   '赤黒区分
    sSql = sSql & ",' '                                     " & vbLf   '集計区分
    sSql = sSql & ",'" & Right(nowCd, 4) & "'               " & vbLf   '工程コード
    sSql = sSql & ",' '                                     " & vbLf   'ラインコード
    sSql = sSql & "," & Xodcx.Weight & "                    " & vbLf   '受入重量
    sSql = sSql & "," & Xodcx.Weight & "                    " & vbLf   '払出重量
    sSql = sSql & ",0                                       " & vbLf   'ロス重量
    sSql = sSql & ",'" & Right(nowCd, 4) & "'               " & vbLf   '受入工程コード
    sSql = sSql & ",'" & Right(PROCD_KOUNYU_TAN_KESSYOU, 4) & "'" & vbLf '払出工程コード('B410')
    sSql = sSql & ",' '                                     " & vbLf   '払出区分
    sSql = sSql & ",'" & gsFactryCd & "'                    " & vbLf   '払出工場コード
    sSql = sSql & ",' '                                     " & vbLf   '払出ラインコード
    sSql = sSql & ",' '                                     " & vbLf   'チャージNo
    sSql = sSql & ",'" & sYear & "'                         " & vbLf   '実績日付(年)
    sSql = sSql & ",'" & sMonth & "'                        " & vbLf   '実績日付(月)
    sSql = sSql & ",'" & sDay & "'                          " & vbLf   '実績日付(日)
    sSql = sSql & ",'" & sCyoku & "'                        " & vbLf   '直区分
    sSql = sSql & ",'" & sHour & "'                         " & vbLf   '実績時間(時)
    sSql = sSql & ",'" & sMin & "'                          " & vbLf   '実績時間(分)
    sSql = sSql & ",'" & Xodcx.STAFFID & "'                 " & vbLf   '担当者
    sSql = sSql & ",'" & Xodcx.STAFFNAME & "'               " & vbLf   '担当者名
    sSql = sSql & ",' '                                     " & vbLf   '濃度区分
    sSql = sSql & ",NULL                                    " & vbLf   '濃度値
    sSql = sSql & ",'7'                                     " & vbLf   '原料送信フラグ
    sSql = sSql & ",'0'                                     " & vbLf   '発生フラグ
    sSql = sSql & ",'1'                                     " & vbLf   '現品区分
    sSql = sSql & ",'" & gsFactryCd & "'                    " & vbLf   '使用予定工場
    sSql = sSql & ",NULL                                    " & vbLf   '元濃度
    sSql = sSql & ",NULL                                    " & vbLf '
    sSql = sSql & ")"
    
    If SqlExec2(sSql) = -1 Then
        GoTo proc_exit
    End If
    
    DBDRV_fcmkc001c_InsXODB3 = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    DBDRV_fcmkc001c_InsXODB3 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

' @(f)
' 機能      : SQL数値変換関数
'
' 返り値    : <入力数値> or NULL
'
' 引き数    : 変換対象数値
'
' 機能説明  : 渡された数値がNULLであれば"NULL"をそうでなければそのまま出力する
Private Function ConvNum(vinput) As String
    If IsNull(vinput) Or vinput = "NULL" Then
        vinput = ""
    End If
    
    If vinput = "" Then
        ConvNum = "NULL"
    Else
        ConvNum = vinput
    End If
End Function

' @(f)
' 機能      : 直区分判定
'
' 返り値    : 直区分
'
' 引き数    : nowdate  -  日付データ
'
' 機能説明  : 渡された日付データの時刻から直区分を判定する
'
Private Function GetCYOKU(nowdate As String) As String
    
    Dim jitutime As String     '判定用の時刻を格納する変数

    '判定用に時刻を切り出す
    jitutime = Format(nowdate, "hhnnss")

    '直区分を設定する
    '3直 00:00から07:59
    If jitutime >= "000000" And jitutime < "080000" Then
        GetCYOKU = "3"
    '1直 08:00から15:59
    ElseIf jitutime >= "080000" And jitutime < "160000" Then
        GetCYOKU = "1"
    '2直 16:00から23:59
    ElseIf jitutime >= "160000" And jitutime < "240000" Then
        GetCYOKU = "2"
    End If
End Function
'---- ADD [精製原料管理、原料工程実績作成処理追加] 以上追加関数　END ---- TCS)T.TERAUCHI
Public Function DBDRV_SELECT_HOLD(gTblDispData As typ_TBCMJ012, pCrynum As String) As FUNCTION_RETURN

    Dim sql As String       'SQL全体
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      'レコード数
    Dim i As Long

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_XSDC1_SQL.bas -- Function DBDRV_SELECT_HOLD"
     ''SQLを組み立てる
     'sql = "SELECT PROCCODE, HLDTRCLS, HLDCAUSE, HLDCMNT, UPDDATE, KSTAFFID, HOLDKT FROM TBCMJ012,XSDC2 "
     sql = "SELECT HLDCMNT FROM TBCMJ012,XSDC2 "
     sql = sql & " WHERE CRYNUMC2 = '" & pCrynum & "'"
     sql = sql & " AND   XTALC2 = CRYNUM   "
     sql = sql & " AND   INPOSC2 = INGOTPOS   "
     'sql = sql & " AND TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMJ012 WHERE CRYNUM = '" & pCrynum & "')"
     sql = sql & " ORDER BY TRANCNT"
    ''データを抽出する
     Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
     
     If rs Is Nothing Then
         DBDRV_SELECT_HOLD = FUNCTION_RETURN_FAILURE
         Exit Function
     End If
     With gTblDispData
        If rs.RecordCount > 0 Then
           rs.MoveLast
           If IsNull(rs("HLDCMNT")) = False Then .HLDCMNT = rs("HLDCMNT")
        End If
    End With
    rs.Close

    DBDRV_SELECT_HOLD = FUNCTION_RETURN_SUCCESS
    

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_SELECT_HOLD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要   ：ﾗﾍﾞﾙ出力処理
'説明   ：「出力済」にした場合ストッカ工程の最後の工程に「出荷」を追加する
'引数   ：strCrynum     I   結晶番号
'       ：strStaffID    I   担当者名
'返り値 ：TRUE  正常
'       ：FALSE 異常
'履歴   ：2006/10/27 SETsw 高崎　新規作成
'備考   ：トランザクションは呼び出し側でかけてください
Public Function StockerShip(StrCryNum As String, StrStaffId As String) As Boolean

    '関数をなるべく独立させるために定義をここに持たせる
    Const SHIP As String = "08"         '出荷
    Const SGEN As String = "09"         '精製原料払出
    Const KAKUSITA As String = "10"     '格下品出庫
    Const STOCKER As String = "11"      'ストッカー
    Const OLDSTOCKER As String = "12"   '既存ストッカー
    Const DELETE As String = "13"       '削除
    Const HAIKI As String = "14"        '廃棄
    
    Dim sSql As String
    Dim rs As OraDynaset    'RecordSet
    
    Dim sProcNum As String  '工程番号
    Dim sProcKbn As String  '工程区分
    
    Dim sLastProcNum As String  '最終工程番号
    
    Dim sTranNum As String      '処理回数
    Dim sNowProcNum As String   '現在工程番号
    
    Dim bUpdateFlg As Boolean   'UPDATEを行ったか
    
    StockerShip = False
    
    '実行条件
    '最終工程が「ストッカー」の場合には「ストッカー」を「出荷」に変更する
    '最終工程が「ストッカー」以外の場合には最終工程に「出荷」を登録する
    'ただし、「精製原料払出」「格下品出庫」「既存ストッカー」「削除」「廃棄」の場合はエラーとする
    'TBCMF005に対象の結晶番号がない場合は何もせず正常終了とする。
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "既存画面.bas -- Function StockerShip"
    
    '前準備
        '処理回数、現在工程番号を取得する
        sSql = ""
        sSql = sSql & "SELECT TRANNUM, PROCNUM AS NOWPROCNUM FROM TBCMF005 WHERE CRYNUM = '" & StrCryNum & "' "
        sSql = sSql & "AND DELCLS = '0'"
        ''データを抽出する
        Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
        If rs Is Nothing Then
            GoTo proc_exit
        End If
        ''抽出結果を格納する
        sTranNum = NulltoStr(rs("TRANNUM"))
        sNowProcNum = NulltoStr(rs("NOWPROCNUM"))
        
        '1件も取得できない場合は正常終了
        If rs.RecordCount = 0 Then
            StockerShip = True
            Exit Function
        End If
        
        rs.Close
        
        '現在工程番号が取得できない場合、現在工程が"0"は異常終了
        If sNowProcNum = "" Or sNowProcNum = "0" Then Exit Function

        '初期化
        Set rs = Nothing
    
        
        '結晶番号に対する削除された工程を含めた最終工程番号を取得する
        sSql = ""
        sSql = sSql & "SELECT NVL(MAX(PROCNUM),0) AS LASTPROCNUM FROM TBCMF006 WHERE CRYNUM = '" & StrCryNum & "' "
        ''データを抽出する
        Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
        If rs Is Nothing Then
            GoTo proc_exit
        End If
        ''抽出結果を格納する
        sLastProcNum = NulltoStr(rs("LASTPROCNUM"))
        
        '仕掛が取得できていて工程が1件も取得できない場合は終了
        If rs.RecordCount = 0 Then
            StockerShip = True
            Exit Function
        End If
        rs.Close
        
        '最終工程番号が取得できない場合は終了
        If sLastProcNum = "0" Then
            StockerShip = True
            Exit Function
        End If

        '初期化
        Set rs = Nothing
    
    '方法
    '1. 結晶番号から最終工程番号(未削除)と工程区分を割り出す
    sSql = ""
    sSql = sSql & "SELECT PROCKBN, PROCNUM FROM TBCMF006 "
    sSql = sSql & "WHERE "
    sSql = sSql & "CRYNUM = '" & StrCryNum & "' "
    sSql = sSql & "AND PROCNUM = ("
        sSql = sSql & "SELECT MAX(PROCNUM) FROM TBCMF006 WHERE CRYNUM = '" & StrCryNum & "' "
        sSql = sSql & "AND DELCLS = '0' "
        sSql = sSql & ") "
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    sProcKbn = NulltoStr(rs("PROCKBN"))
    sProcNum = NulltoStr(rs("PROCNUM"))
        
    '1件も取得できない場合は終了
    If rs.RecordCount = 0 Then
        StockerShip = True
        Exit Function
    End If
    
    rs.Close
    
    '初期化
    Set rs = Nothing
    
    '2. 工程区分を判断する
    '   2.1 最終工程(未削除)が「精製原料払出」「格下品出庫」「既存ストッカー」「削除」「廃棄」場合にはエラー
        If sProcKbn = SGEN Or sProcKbn = KAKUSITA Or sProcKbn = OLDSTOCKER Or sProcKbn = DELETE Or sProcKbn = HAIKI Then
            Exit Function
    '   2.2 最終工程(未削除)が「ストッカー」の場合には結晶番号と割り出した工程番号を用いてUPDATE
        ElseIf sProcKbn = STOCKER Then
        'ここではSQLを作成するだけ。実行はif文を抜けた後。
            sSql = ""
            sSql = sSql & "UPDATE TBCMF006 SET "
            sSql = sSql & "PROCKBN = '" & SHIP & "', "
            sSql = sSql & "KSTAFFID = '" & StrStaffId & "', "
            sSql = sSql & "UPDDATE = SYSDATE "
            sSql = sSql & "WHERE "
            sSql = sSql & "CRYNUM = '" & StrCryNum & "' "
            sSql = sSql & "AND PROCNUM = " & sProcNum
            
            'UPDATE実行フラグを立てておく
            bUpdateFlg = True
    '   2.3 2.1,2.2以外の場合、結晶番号と割り出した最終工程番号+1を用いてINSERT
        'ここではSQLを作成するだけ。実行はif文を抜けた後。
        Else
            sSql = ""
            sSql = sSql & "INSERT INTO TBCMF006("
            sSql = sSql & "CRYNUM,"
            sSql = sSql & "PROCNUM,"
            sSql = sSql & "PROCKBN,"
            sSql = sSql & "PROCSTAT,"
            sSql = sSql & "HOLDFLG,"
            sSql = sSql & "PRIORITY,"
            sSql = sSql & "KSIYOUFLG,"
            sSql = sSql & "DELIVFLG,"
            sSql = sSql & "STOCKFLG,"
            sSql = sSql & "RESLTFLG,"
            sSql = sSql & "INSPCTFLG,"
            sSql = sSql & "DELCLS,"
            sSql = sSql & "TSTAFFID,"
            sSql = sSql & "REGDATE) "
            sSql = sSql & "VALUES("
            sSql = sSql & "'" & StrCryNum & "',"
            sSql = sSql & CInt(sLastProcNum) + 1 & ","
            sSql = sSql & "'" & SHIP & "',"
            sSql = sSql & "'0',"
            sSql = sSql & "'0',"
            sSql = sSql & "'1',"
            sSql = sSql & "'0',"
            sSql = sSql & "'0',"
            sSql = sSql & "'0',"
            sSql = sSql & "'0',"
            sSql = sSql & "'0',"
            sSql = sSql & "'0',"
            sSql = sSql & "'" & StrStaffId & "',"
            sSql = sSql & "SYSDATE)"
        End If
    '実行
    OraDB.ExecuteSQL (sSql)

    'UPDATEの場合、UPDATEした工程が現在工程番号であった場合は加工仕掛テーブルも更新する
    If bUpdateFlg Then
        If sProcNum = sNowProcNum Then
            sSql = ""
            sSql = sSql & "UPDATE TBCMF005 SET "
            sSql = sSql & "PROCKBN = '" & SHIP & "', "
            sSql = sSql & "KSTAFFID = '" & StrStaffId & "', "
            sSql = sSql & "UPDDATE = SYSDATE "
            sSql = sSql & "WHERE "
            sSql = sSql & "CRYNUM = '" & StrCryNum & "' "
            sSql = sSql & "AND TRANNUM = " & sTranNum
            
            '実行
            OraDB.ExecuteSQL (sSql)
        End If
    End If
    
    StockerShip = True
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    
    Resume proc_exit

End Function

' 2007/08/30 SPK Tsutsumi Add Start
Public Function GetMukeCode() As FUNCTION_RETURN
    Dim sql As String
    Dim rs As OraDynaset
    Dim recCnt As Long      'レコード数
    Dim i  As Long
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmbc016_0.frm -- Function Getstaffauthority"
    
    GetMukeCode = FUNCTION_RETURN_FAILURE
    
    sql = "Select CODEA9,NAMEJA9 "
    sql = sql & "from KODA9 "
    sql = sql & "where SYSCA9 = 'X' "
    sql = sql & "and SHUCA9 = '20' "
    sql = sql & "and (CODEA9 = '14' "
    sql = sql & "or CODEA9 = '15' "
    sql = sql & "or CODEA9 = '16' "
'2008/05/26 SHINDOH UPD
'    sql = sql & "or CODEA9 = 'ZZ') "
'------------------------------------
    sql = sql & "or CODEA9 = 'ZZ' "
    sql = sql & "or CODEA9 = 'ZX') "
'------------------------------------
    sql = sql & "order by CODEA9 "      '向先不具合対応 2009/01/04 SETsw kubota

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If
    
    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim s_Mukesaki(recCnt)
    
    If recCnt = 0 Then
        Exit Function
    End If
    
    For i = 1 To recCnt
        With s_Mukesaki(i)
            If IsNull(rs.Fields("CODEA9")) = False Then .sMukeCode = rs.Fields("CODEA9")    ' 向先コード
            If IsNull(rs.Fields("NAMEJA9")) = False Then .sMukeName = rs.Fields("NAMEJA9")  ' 向先名
        End With
        rs.MoveNext
    Next
    rs.Close

    GetMukeCode = FUNCTION_RETURN_SUCCESS
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function
'2007/08/30 SPK Tsutsumi Add End

'2008/01/25 SETsw kubota Add Start
Public Function GetSiyoHaraiLen(ByRef sHSXCLMIN As String _
                              , ByRef sHSXCLMAX As String _
                              ) As Boolean
    
    Dim sql As String
    Dim rs As OraDynaset
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL -- Function GetSiyoHaraiLen"
    
    GetSiyoHaraiLen = False
    
    ''下限値の最大(最も厳しい仕様)を取得
    sql = "Select max(HSXCLMIN) HSXCLMIN"   '下限値の最大
    sql = sql & "  from TBCME020,XSDCA"
    sql = sql & " where CRYNUMCA = '" & f_cmbc032_2.txtBlkID & "'"
    sql = sql & "   and HINBAN   = HINBCA"
    sql = sql & "   and MNOREVNO = REVNUMCA"
    sql = sql & "   and FACTORY  = FACTORYCA"
    sql = sql & "   and OPECOND  = OPECA"
    sql = sql & "   and HSXCLMIN <> 0"
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.EOF = True Then
        GoTo proc_exit
    End If
    sHSXCLMIN = NulltoStr(rs.Fields("HSXCLMIN"))
    
    rs.Close
    
    ''上限値の最小(最も厳しい仕様)を取得
    sql = "Select min(HSXCLMAX) HSXCLMAX"   '上限値の最小
    sql = sql & "  from TBCME020,XSDCA"
    sql = sql & " where CRYNUMCA = '" & f_cmbc032_2.txtBlkID & "'"
    sql = sql & "   and HINBAN   = HINBCA"
    sql = sql & "   and MNOREVNO = REVNUMCA"
    sql = sql & "   and FACTORY  = FACTORYCA"
    sql = sql & "   and OPECOND  = OPECA"
    sql = sql & "   and HSXCLMAX <> 0"
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.EOF = True Then
        GoTo proc_exit
    End If
    sHSXCLMAX = NulltoStr(rs.Fields("HSXCLMAX"))
    
    rs.Close

    GetSiyoHaraiLen = True

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function
'2008/01/25 SETsw kubota Add End

''***********************************************************************************************
''SHINDOH ADD
''
''***********************************************************************************************
'------------------------------------------------------------------------------------------------------------(作りり直しSTR)
'概要    :WF出荷待ち一覧 初期表示用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO  ,型                                 ,説明
'        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,初期表示用
'        :戻ﾘ値        ,O   ,FUNCTION_RETURN                    ,読み込み成否
'説明    :
'履歴    :
'@'Public Function DBDRV_scmzc_fcmkc001b_Disp5(records() As type_DBDRV_scmzc_fcmkc001b_Disp5) As FUNCTION_RETURN
'@'
    '＜WFセンタ払出待ち＞
    'CC720のもの
    'エラーハンドラの設定
'@'    On Error GoTo proc_err
'@'    gErr.Push "s_cmbc032_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp5"

'@'    DBDRV_scmzc_fcmkc001b_Disp5 = FUNCTION_RETURN_SUCCESS


    'ブロックID､更新日付、品番等取得
'@'     If GetListData(records(), "CC720") = FUNCTION_RETURN_FAILURE Then
'@'        DBDRV_scmzc_fcmkc001b_Disp5 = FUNCTION_RETURN_FAILURE
'@'        GoTo proc_exit
'@'    End If


'@'proc_exit:
    '終了
'@'    gErr.Pop
'@'    Exit Function

'@'proc_err:
    'エラーハンドラ
'@'    gErr.HandleError
'@'    DBDRV_scmzc_fcmkc001b_Disp5 = FUNCTION_RETURN_FAILURE
'@'    Resume proc_exit
'@'End Function
Public Function DBDRV_scmzc_fcmkc001b_Disp5(records() As type_DBDRV_scmzc_fcmkc001b_Disp5, tmpBlkData() As typ_BlkData) As FUNCTION_RETURN

    'ブロックID､更新日付取得

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp"

    DBDRV_scmzc_fcmkc001b_Disp5 = FUNCTION_RETURN_SUCCESS

    'ブロックID､更新日付、品番等取得
    If getBlockID2(records(), tmpBlkData(), "CC720", 1, 0, "") = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmkc001b_Disp5 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    DBDRV_scmzc_fcmkc001b_Disp5 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'------------------------------------------------------------------------------------------------------------(作りり直しSTR)
'概要    :SXL出荷待ち一覧 初期表示用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO  ,型                                 ,説明
'        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,初期表示用
'        :戻ﾘ値        ,O   ,FUNCTION_RETURN                    ,読み込み成否
'説明    :
'履歴    :
Public Function DBDRV_scmzc_fcmkc001b_Disp6(records() As type_DBDRV_scmzc_fcmkc001b_Disp5, tmpBlkData() As typ_BlkData, scmbKosei As Integer, stxtblk As String) As FUNCTION_RETURN

    '＜SXL出荷前＞
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp6"

    DBDRV_scmzc_fcmkc001b_Disp6 = FUNCTION_RETURN_SUCCESS

    'ブロックID､更新日付、品番等取得
     If getBlockID2(records(), tmpBlkData(), "CC705", 2, scmbKosei, stxtblk) = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmkc001b_Disp6 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    DBDRV_scmzc_fcmkc001b_Disp6 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'内部関数 ブロックID、更新日付取得（払出待ち、抜試指示待ち用）
Private Function getBlockID2(records() As type_DBDRV_scmzc_fcmkc001b_Disp5, _
                            pBlkData() As typ_BlkData, _
                            NOWPROC As String, formNum As Integer, scmbKosei As Integer, stxtblk As String) As FUNCTION_RETURN

    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim recCnt      As Long         'レコード数
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim BlockIdBuf  As String
    Dim sBlkID      As String
    Dim blkOrder    As Integer
    Dim Jiltuseki   As Judg_Kakou
    Dim nowtime     As Date         '現在日付
    Dim sBakPos     As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function getBlockID2"

    getBlockID2 = FUNCTION_RETURN_SUCCESS

'    sql = "select X.XTALCA    as CRYNUM, "
'    sql = sql & " X.CRYNUMCA  as BLOCKID, "
'    sql = sql & " B.INGOTPOS, "
'    sql = sql & " XC2.KDAYC2, "
'    sql = sql & " B.HOLDCLS, "
'    sql = sql & " X.HINBCA    as HINBAN, "      ' 品番
'    sql = sql & " X.REVNUMCA  as REVNUM, "      ' 製品番号改訂番号
'    sql = sql & " X.FACTORYCA as FACTORY, "     ' 工場
'    sql = sql & " X.OPECA     as OPECOND, "     ' 操業条件
'    sql = sql & " S.HSXTYPE, "                  ' 品ＳＸタイプ
'    sql = sql & " S.HSXCDIR, "                  ' 品ＳＸ結晶面方位
'    sql = sql & " B.REALLEN, "
'    sql = sql & " XC2.GNLC2 as LEN, "
'    sql = sql & " B2.BLOCKID  as SBLOCKID, "
'    sql = sql & " nvl("
'    sql = sql & "    (select DMTOP1 from TBCMI002 I2"
'    sql = sql & "     where CRYNUM=B.CRYNUM"
'    sql = sql & "       and INGOTPOS=(select max(INGOTPOS) from TBCMI002 where CRYNUM=B.CRYNUM  and INGOTPOS<=B.INGOTPOS)"
'    sql = sql & "       and TRANCNT =(select max(TRANCNT)  from TBCMI002 where CRYNUM=I2.CRYNUM and INGOTPOS=I2.INGOTPOS)"
'    sql = sql & "    )"
'    sql = sql & "    , (select DIAMETER from TBCME037 where CRYNUM=B.CRYNUM)"
'    sql = sql & "  ) as DIAM, "
'    sql = sql & " (select max(UPDDATE) from TBCMW001 where CRYNUM=B2.CRYNUM and INGOTPOS=B2.INGOTPOS) as NUKISHI_AT "
'    sql = sql & ",XC1.PUPTNC1 as PUPTN "            '引上ﾊﾟﾀｰﾝ追加対応
'    sql = sql & ",X.HOLDBCA "                       'ﾎｰﾙﾄﾞ区分(XSDCA)
'    sql = sql & ",E36.WFCUTT "                      'WFｶｯﾄ単位
'    sql = sql & ",E36.BLOCKHFLAG "                  'ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞ
'    sql = sql & ",X.HOLDCCA "                       'ﾎｰﾙﾄﾞ理由
'    sql = sql & ",X.HOLDKTCA "                       'ﾎｰﾙﾄﾞ工程
'    sql = sql & ",X.PLANTCATCA "                    '向先
'    sql = sql & " From  XSDCA X, TBCME018 S, TBCMJ010 J, TBCME040 B, TBCME040 B2 "
'    sql = sql & "     , XSDC1 XC1 "                 '引上ﾊﾟﾀｰﾝ追加対応
'    sql = sql & "     , TBCME036 E36 "
'    sql = sql & "     , XSDC2 XC2 "
'    sql = sql & " where X.CRYNUMCA = B.BLOCKID "
'    sql = sql & "   and X.HINBCA   = S.HINBAN "
'    sql = sql & "   and X.REVNUMCA = S.MNOREVNO "
'    sql = sql & "   and X.FACTORYCA= S.FACTORY "
'    sql = sql & "   and X.OPECA    = S.OPECOND "
'    sql = sql & "   and X.HINBCA   = E36.HINBAN "
'    sql = sql & "   and X.REVNUMCA = E36.MNOREVNO "
'    sql = sql & "   and X.FACTORYCA= E36.FACTORY "
'    sql = sql & "   and X.OPECA    = E36.OPECOND "
'    sql = sql & "   and X.GNWKNTCA = '" & NOWPROC & "' "
'    sql = sql & "   and B.LSTATCLS = 'T' "
'    sql = sql & "   and B.RSTATCLS = 'T' "
'    sql = sql & "   and X.LIVKCA   = '0' "
'    sql = sql & "   and B.DELCLS   = '0' "
'    sql = sql & "   and J.CRYNUM   = B.CRYNUM "
'    sql = sql & "   and J.INGOTPOS = B.INGOTPOS"
'    sql = sql & "   and J.TRANCNT  = (select max(TRANCNT) from TBCMJ010 where CRYNUM=J.CRYNUM and INGOTPOS=J.INGOTPOS)"
'    sql = sql & "   and B2.CRYNUM  = B.CRYNUM"
'    sql = sql & "   and B2.INGOTPOS = B.INGOTPOS"
'    sql = sql & "   and X.XTALCA   = XC1.XTALC1(+) "   '引上ﾊﾟﾀｰﾝ追加対応
'    sql = sql & "   and XC2.CRYNUMC2 = X.CRYNUMCA "
    sql = "select  X.XTALCA  as CRYNUM,"
    sql = sql & " X.CRYNUMCA  as BLOCKID,"
    sql = sql & " XC2.INPOSC2,"
    sql = sql & " XC2.KDAYC2,"
    sql = sql & " XC2.HOLDBC2,"
    sql = sql & " X.HINBCA    as HINBAN,"
    sql = sql & " X.REVNUMCA  as REVNUM,"
    sql = sql & " X.FACTORYCA as FACTORY,"
    sql = sql & " X.OPECA     as OPECOND,"
    sql = sql & " S.HSXTYPE,"
    sql = sql & " S.HSXCDIR,"
    sql = sql & " XC2.REALLC2,"
    sql = sql & " XC2.GNLC2 as LEN,"
    sql = sql & " XC1.PUPTNC1 as PUPTN ,"
    sql = sql & " X.HOLDBCA ,"
    sql = sql & " X.HOLDCCA ,"
    sql = sql & " X.HOLDKTCA ,"
    sql = sql & " X.PLANTCATCA"
    ' 流動監視SQL修正 upd SETkimizuka Start  09/07/01
    '' 流動停止項目追加 add SETkimizuka Start  09/03/26
    'sql = sql & " , NVL(TO_CHAR(Y4.AGRSTATUS),' ') as AGRSTATUS "
    'sql = sql & " , NVL(TO_CHAR(Y4.STOP),'0') as STOP "
    'sql = sql & " , NVL(Y4.CAUSE,' ') as CAUSE "
    'sql = sql & " , NVL(Y4.PRINTKIND || Y4.PRINTNO,' ') as PRINTNO "
    '' 流動停止項目追加 add SETkimizuka End    09/03/26
    sql = sql & " , NVL(TO_CHAR(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)),' ') as AGRSTATUS "
    sql = sql & " , NVL(TO_CHAR(Y4.STOPY4),' ') as STOP "
    sql = sql & " , DECODE(TRIM(Y4.CAUSEY4),NULL,' ',TRIM(Y4.CAUSEY4) || ':' || A9.NAMEJA9) as CAUSE "
    sql = sql & " , NVL(Y4.PRINTKINDY4 || Y4.PRINTNOY4,' ') as PRINTNO "
    sql = sql & " , NVL(Y4.WKKTY4,'0') as WKKTY4 "
    ' 流動監視SQL修正 upd SETkimizuka End  09/07/01
    sql = sql & " From  XSDCA X, TBCME018 S,XSDC1 XC1  ,XSDC2 XC2"
    ' 流動監視SQL修正 upd SETkimizuka Start  09/07/01
    sql = sql & "    ,XODY3 Y3,XODY4 Y4,KODA9 A9 "
    '' 流動停止項目追加 add SETkimizuka Start  09/03/26
    'sql = sql & "    ,(SELECT XTALNOY3 as XTALNO,MIN(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",AGRSTATUSY4)) as AGRSTATUS  "
    'sql = sql & "      ,MAX(STOPY4) as STOP,DECODE(TRIM(CAUSEY4),'',TRIM(CAUSEY4),TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSE ,Y5.PRINTNO,Y5.PRINTKIND "
    'sql = sql & "      FROM XODY3  "
    'sql = sql & "           LEFT OUTER JOIN XODY4 on ( XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND STOPY4 <> '2'  AND WKKTY4 in " & IIf(formNum = 1, CreateWkktSQL(WATCH_PROCCD_WF), IIf(formNum = 2, CreateWkktSQL(WATCH_PROCCD_BAR), "(' ')")) & ")"
    'sql = sql & "           LEFT OUTER JOIN KODA9 on ( SYSCA9 = 'X' AND SHUCA9 = '30' AND CAUSEY4 = CODEA9 ) "
    'sql = sql & "           LEFT OUTER JOIN (SELECT XTALNOY4 as XTALNO,SXLIDY4 as SXLID,PRINTNOY5 as PRINTNO,PRINTKINDY5 as PRINTKIND "
    'sql = sql & "                FROM XODY3,XODY4,XODY5 "
    'sql = sql & "              WHERE XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND LIVKY3 = '0' "
    'sql = sql & "                AND PRINTKINDY4 = PRINTKINDY5 AND PRINTNOY4 = PRINTNOY5  "
    'sql = sql & "                AND HKBNY5 ='0' GROUP BY XTALNOY4,SXLIDY4,PRINTNOY5,PRINTKINDY5) Y5 ON (XTALNOY3 = XTALNO AND SXLIDY3 = SXLID ) "
    'sql = sql & "      WHERE  "
    'sql = sql & "       LIVKY3    = '0' "
    'sql = sql & "       GROUP BY XTALNOY3,AGRSTATUSY4,CAUSEY4,Y5.PRINTNO,Y5.PRINTKIND,NAMEJA9) Y4 "
    ' 流動停止項目追加 add SETkimizuka End  09/03/26
    ' 流動監視SQL修正 upd SETkimizuka End  09/07/01
    sql = sql & " Where"
    sql = sql & " x.HINBCA = s.hinban"
    sql = sql & " and X.REVNUMCA = S.MNOREVNO"
    sql = sql & " and X.FACTORYCA= S.FACTORY"
    sql = sql & " and X.OPECA    = S.OPECOND"
    sql = sql & " and X.XTALCA   = XC1.XTALC1(+)"
    If formNum = 1 Then
        sql = sql & " and XC2.GNWKNTC2 = 'CC720'"
    '    sql = sql & " and XC2.LSTATBC2 = 'T'"
    ElseIf formNum = 2 Then
        sql = sql & " and XC2.GNWKNTC2 = 'CC705'"
        sql = sql & " and XC2.LSTATBC2 = 'B'"
    End If
    sql = sql & " and XC2.RSTATBC2 = 'T'"
    sql = sql & " and XC2.LIVKC2   = '0'"
    sql = sql & " and X.LIVKCA     = '0'"
    sql = sql & " and XC2.CRYNUMC2 = X.CRYNUMCA"
    ' 流動監視SQL修正 upd SETkimizuka Start  09/07/01
    'sql = sql & " AND X.CRYNUMCA    = Y4.XTALNO(+) "            'add 09/03/26 SETkimizuka
    sql = sql & " AND X.CRYNUMCA = Y3.XTALNOY3(+) "
    sql = sql & " AND Y3.LIVKY3(+) = '0' "
    sql = sql & " AND Y4.LIVKY4(+) = '0' "
    sql = sql & " AND Y3.XTALNOY3 = Y4.XTALNOY4(+) "
    sql = sql & " AND Y3.RCNTY3 = Y4.RCNTY4(+) "
    sql = sql & " AND A9.SYSCA9(+) = 'X' AND A9.SHUCA9(+) = '30' AND Y4.CAUSEY4 = A9.CODEA9(+) "
    ' 流動監視SQL修正 upd SETkimizuka End  09/07/01

    If formNum = 2 Then
        sql = sql & "   and XC2.CRYNUMC2 like '" & stxtblk & "' "
        If scmbKosei = 0 Then
            sql = sql & "   and trim(XC2.GNWKKBC2) is null"
        Else
            sql = sql & "   and XC2.GNWKKBC2 ='" & scmbKosei & "'"
        End If
    End If

    sql = sql & " order by X.CRYNUMCA , X.INPOSCA"

Debug.Print "GetBlk " & sql
    'データを抽出する
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    'レコードがない場合正常終了
    If rs.RecordCount = 0 Then
        rs.Close
        ReDim records(0)
        ReDim pBlkData(0)
        GoTo proc_exit
    End If

    BlockIdBuf = vbNullString
    recCnt = rs.RecordCount

    ReDim pBlkData(1 To recCnt)
    sBlkID = vbNullString
    blkOrder = 0
    j = 0
    For i = 1 To recCnt
        DoEvents
        'ブロックID等の格納
        If rs("BLOCKID") <> BlockIdBuf Then

            j = j + 1
            ReDim Preserve records(j)
            With records(j)
                .CRYNUM = rs("CRYNUM")
                .INGOTPOS = rs("INPOSC2")
                .BLOCKID = rs("BLOCKID")   ' ブロックID
                .UPDDATE = rs("KDAYC2")   ' 更新日付
                .HOLDCLS = rs("HOLDBC2")   ' ホールド区分
                BlockIdBuf = records(j).BLOCKID
                .HSXTYPE = rs("HSXTYPE")
                .HSXCDIR = rs("HSXCDIR")
                .Judg = " "
                .PUPTN = rs("PUPTN")
                If IsNull(rs("HOLDBCA")) = False Then .HOLDBCA = rs("HOLDBCA") Else .HOLDBCA = " "  'ﾎｰﾙﾄﾞ区分(XSDCA)
                If IsNull(rs("HOLDCCA")) = False Then .HOLDC = rs("HOLDCCA") Else .HOLDC = " "  'ﾎｰﾙﾄﾞ理由
                If IsNull(rs("HOLDKTCA")) = False Then .HOLDKT = rs("HOLDKTCA") Else .HOLDKT = " "  'ﾎｰﾙﾄﾞ工程
            End With

            k = 1
            sBakPos = ""    'add 09/03/26 SETkimizuka
        End If


        With pBlkData(i)
            .CRYNUM = rs("CRYNUM")
            .BLOCKID = rs("BLOCKID")
            .INGOTPOS = rs("INPOSC2")
            .LENGTH = rs("LEN")
            .REALLEN = rs("REALLC2")
            '.sBlockID = rs("sBLOCKID")
            .sBlockId = rs("BLOCKID")
            If sBlkID <> .sBlockId Then
                sBlkID = .sBlockId
                blkOrder = 1
            Else
                blkOrder = blkOrder + 1
            End If
            .BLOCKORDER = blkOrder
            '.DIAMETER = rs("DIAM")

            '最終抜試日付に現在日付を表示
            nowtime = getSvrTime()
            .WFINDDATE = Format$(nowtime, "YYYY/MM/DD")
            .HOLDCLS = rs("HOLDBC2")
        End With


        If InStr(sBakPos, Trim(rs("INPOSC2"))) = 0 Then '流動監視項目追加に伴う修正 upd
            '品番の格納
            ReDim Preserve records(j).HIN(k)
            records(j).HIN(k).hinban = rs("HINBAN")
            records(j).HIN(k).mnorevno = rs("REVNUM")
            records(j).HIN(k).factory = rs("FACTORY")
            records(j).HIN(k).opecond = rs("OPECOND")
            
            If IsNull(rs("PLANTCATCA")) = False Then
                records(j).HIN(k).sMukesaki = rs("PLANTCATCA")
            End If
            sBakPos = sBakPos & Trim(rs("INPOSC2")) & " "
            k = k + 1
        End If

        ''品番の格納
        'ReDim Preserve records(j).HIN(k)
        'records(j).HIN(k).hinban = rs("HINBAN")
        'records(j).HIN(k).mnorevno = rs("REVNUM")
        'records(j).HIN(k).Factory = rs("FACTORY")
        'records(j).HIN(k).OpeCond = rs("OPECOND")
        
        'If IsNull(rs("PLANTCATCA")) = False Then
        '    records(j).HIN(k).sMukesaki = rs("PLANTCATCA")
        'End If

'        If k = 1 Then
'            '品番１WFｶｯﾄ単位
'            If IsNull(rs("WFCUTT")) = False Then
'                records(j).WFCUTT = rs("WFCUTT")
'            Else
'                records(j).WFCUTT = -1
'            End If
'            '品番１ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞ
'            If IsNull(rs("BLOCKHFLAG")) = False Then
'                records(j).BLOCKHFLAG = rs("BLOCKHFLAG")
'            Else
'                records(j).BLOCKHFLAG = " "
'            End If
'        End If
'

        ' 流動監視SQL修正 upd SETkimizuka Start  09/07/01
        '' 流動停止項目追加 add SETkimizuka Start  09/03/26
        'records(j).STOP = rs("STOP")                   '停止区分
        'records(j).AGRSTATUS = rs("AGRSTATUS")       '承認確認区分
        'If Trim(rs("CAUSE")) <> "" And InStr(records(j).CAUSE, Trim(rs("CAUSE"))) = 0 Then
        '    records(j).CAUSE = records(j).CAUSE & rs("CAUSE") & vbTab       '停止理由
        'End If
        
        'IIf(formNum = 1, CreateWkktSQL(CC720), IIf(formNum = 2, CreateWkktSQL(CC705)
        
        If rs("STOP") <> "2" And rs("WKKTY4") = IIf(formNum = 1, "CC720", IIf(formNum = 2, "CC705", "")) Then
           If Trim(records(j).AGRSTATUS) = "" Or (rs("AGRSTATUS") < records(j).AGRSTATUS) Then
                records(j).STOP = rs("STOP")                   '停止区分
                records(j).AGRSTATUS = rs("AGRSTATUS")         '承認確認区分
           End If
            If Trim(rs("CAUSE")) <> "" And InStr(records(j).CAUSE, Trim(rs("CAUSE"))) = 0 Then
                records(j).CAUSE = records(j).CAUSE & rs("CAUSE") & vbTab       '停止理由
            End If
        End If
        ' 流動監視SQL修正 upd SETkimizuka End  09/07/01
        If Trim(rs("PRINTNO")) <> "" And InStr(records(j).PRINTNO, Trim(rs("PRINTNO"))) = 0 Then
            records(j).PRINTNO = records(j).PRINTNO & rs("PRINTNO") & vbTab       '先行評価
        End If
        ' 流動停止項目追加 add SETkimizuka End    09/03/26

        'k = k + 1
        rs.MoveNext
    Next i
    rs.Close

    For i = 1 To recCnt
        With pBlkData(i)
            If scmzc_getKakouJiltuseki(.BLOCKID, Jiltuseki) = FUNCTION_RETURN_SUCCESS Then
                .DIAMETER = (Jiltuseki.TAIL(1) + Jiltuseki.TAIL(2) + Jiltuseki.top(1) + Jiltuseki.top(2)) / 4
            End If
        End With
    Next

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getBlockID2 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

Public Function cmkc001b_DBDataCheck5(LWD() As cmkc001b_LockWait, _
                                      Wd3() As type_DBDRV_scmzc_fcmkc001b_Disp5) As FUNCTION_RETURN
    Dim c0          As Integer
    Dim c1          As Integer
    Dim c2          As Integer
    Dim MaxRec      As Integer
    Dim RecCount    As Integer
    Dim EQFlag      As Boolean
    Dim sql         As String       ' SQL全体
    Dim rs          As OraDynaset   ' RecordSet
    Dim GrpCount1   As Integer
    Dim GrpCount2   As Integer
    Dim ColorFlag   As Boolean
    Dim TotalBlk    As Integer
    Dim CheckPoint  As Integer
    Dim CheckEnd    As Integer
    Dim tempGrpFlag As String * 1

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function cmkc001b_DBDataCheck5"

    cmkc001b_DBDataCheck5 = FUNCTION_RETURN_SUCCESS
    TotalBlk = UBound(Wd3())

Debug.Print " 1:" & Time

    'CC700のブロックの結晶一覧を作る
    ReDim GrpInfo(1) As cmkc001b_Wait3
    GrpInfo(1).CRYNUM = vbNullString
    c1 = 0
    For c0 = 1 To TotalBlk
        DoEvents
        If c1 = 0 Then
            GrpInfo(1).CRYNUM = Wd3(c0).CRYNUM
        End If
        MaxRec = UBound(GrpInfo())
        EQFlag = False
        c1 = 1
        Do While c1 <= MaxRec
            DoEvents
            If Wd3(c0).CRYNUM = GrpInfo(c1).CRYNUM Then
                EQFlag = True
                Exit Do
            End If
            c1 = c1 + 1
        Loop
        If Not EQFlag Then
            ReDim Preserve GrpInfo(MaxRec + 1) As cmkc001b_Wait3
            GrpInfo(MaxRec + 1).CRYNUM = Wd3(c0).CRYNUM
        End If
    Next

Debug.Print " 2:" & Time

    '結晶に含まれる全てのブロックを求める
    'Z品番ブロック（下３桁中に'$'を含むもの）は除く

    MaxRec = UBound(GrpInfo())
    For c0 = 1 To MaxRec
        sql = "select BLOCKID, INGOTPOS, LENGTH, NOWPROC, HOLDCLS "
        sql = sql & "from  TBCME040 "
        sql = sql & "where CRYNUM='" & GrpInfo(c0).CRYNUM & "' "
        sql = sql & "  and 0     = INSTR(BLOCKID,'$',10,1)"
        sql = sql & "order by INGOTPOS, BLOCKID "

        'データを抽出する
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        RecCount = rs.RecordCount
        If RecCount = 0 Then
            rs.Close
            GoTo proc_exit
        End If
        ReDim GrpInfo(c0).blkInfo(RecCount) As cmkc001b_Wait3_BLK
        For c1 = 1 To RecCount
            GrpInfo(c0).blkInfo(c1).BLOCKID = rs("BLOCKID")
            GrpInfo(c0).blkInfo(c1).INGOTPOS = rs("INGOTPOS")
            GrpInfo(c0).blkInfo(c1).LENGTH = rs("LENGTH")
            GrpInfo(c0).blkInfo(c1).NOWPROC = rs("NOWPROC")
            GrpInfo(c0).blkInfo(c1).HOLDCLS = rs("HOLDCLS")
            rs.MoveNext
        Next
        rs.Close
    Next

Debug.Print " 3:" & Time

    'ブロックの上下品番を求める
#If SPEEDUP Then   '高速化実験 02.1.28-2.15 野村
'高速化メモ
'ブロックの上下品番を求めるだけなら、1回のSQLでまとめて情報を取得できるはず
Dim blkID()     As String
Dim topHin()    As tFullHinban
Dim botHin()    As tFullHinban
Dim idx         As Integer
Dim rsCount     As Integer
Dim found       As Boolean

    sql = vbNullString
    sql = sql & "select"
    sql = sql & "  b.BLOCKID"
    sql = sql & ", TOP.HINBAN as THINBAN, TOP.REVNUM as TREVNUM, TOP.FACTORY as TFACTORY, TOP.OPECOND as TOPECOND"
    sql = sql & ", BOT.HINBAN as BHINBAN, BOT.REVNUM as BREVNUM, BOT.FACTORY as BFACTORY, BOT.OPECOND as BOPECOND "

    sql = sql & "from TBCME040 B, TBCME041 TOP, TBCME041 BOT "

    sql = sql & "Where B.CRYNUM            =  TOP.CRYNUM"
    sql = sql & "  and B.CRYNUM            =  BOT.CRYNUM"

    sql = sql & "  and B.INGOTPOS          >=  0"
    sql = sql & "  and B.DELCLS            =  '0'"

    sql = sql & "  and B.NOWPROC           in ('CC600','CC700', 'CC710', 'CC720')"
    sql = sql & "  and B.RSTATCLS          =  'T'"
    sql = sql & "  and B.HOLDCLS           =  '0'"

    sql = sql & "  and B.INGOTPOS          >= TOP.INGOTPOS"
    sql = sql & "  and B.INGOTPOS          <  TOP.INGOTPOS+TOP.LENGTH"
    sql = sql & "  and B.INGOTPOS+B.LENGTH >  BOT.INGOTPOS"
    sql = sql & "  and B.INGOTPOS+B.LENGTH <= BOT.INGOTPOS+BOT.LENGTH "

    sql = sql & "order by B.BLOCKID"

    'データを抽出する
    ' 上下品番をフル品番で取得
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    rsCount = rs.RecordCount

    ReDim blkID(1 To rsCount)
    ReDim topHin(1 To rsCount)
    ReDim botHin(1 To rsCount)
    For c0 = 1 To rsCount
        blkID(c0) = rs!BLOCKID

        topHin(c0).hinban = rs!THINBAN
        topHin(c0).mnorevno = rs!TREVNUM
        topHin(c0).factory = rs!TFACTORY
        topHin(c0).opecond = rs!TOPECOND

        botHin(c0).hinban = rs!BHINBAN
        botHin(c0).mnorevno = rs!BREVNUM
        botHin(c0).factory = rs!BFACTORY
        botHin(c0).opecond = rs!BOPECOND
        rs.MoveNext
    Next
    rs.Close

    For c0 = 1 To MaxRec
        RecCount = UBound(GrpInfo(c0).blkInfo())
        For c1 = 1 To RecCount
            With GrpInfo(c0).blkInfo(c1)
                found = False
                For idx = 1 To rsCount
                    If blkID(idx) = .BLOCKID Then
                        found = True
                        Exit For
                    ElseIf blkID(idx) > .BLOCKID Then
                        Exit For
                    End If
                Next

                If found Then
                    .topHin.hinban = topHin(idx).hinban
                    .topHin.factory = topHin(idx).factory
                    .topHin.opecond = topHin(idx).opecond
                    .topHin.REVNUM = topHin(idx).mnorevno
                Else
                    .topHin.hinban = ""
                    .topHin.factory = ""
                    .topHin.opecond = ""
                    .topHin.REVNUM = 0
                End If

                If found Then
                    .botHin.hinban = botHin(idx).hinban
                    .botHin.factory = botHin(idx).factory
                    .botHin.opecond = botHin(idx).opecond
                    .botHin.REVNUM = botHin(idx).mnorevno
                Else
                    .botHin.hinban = ""
                    .botHin.factory = ""
                    .botHin.opecond = ""
                    .botHin.REVNUM = 0
                End If
            End With
        Next
    Next
#Else
' #IF で処理されないのでコメントとしておく（コード判読UPのため）
'    For c0 = 1 To MaxRec
'        RecCount = UBound(GrpInfo(c0).blkInfo())
'        For c1 = 1 To RecCount
'            sql = "select "
'            sql = sql & "HINBAN, "
'            sql = sql & "REVNUM, "
'            sql = sql & "FACTORY, "
'            sql = sql & "OPECOND "
'            sql = sql & "from TBCME041 "
'            sql = sql & "where CRYNUM='" & GrpInfo(c0).Crynum & "' "
''2001/11/14 S.Sano            sql = sql & "and INGOTPOS <= " & GrpInfo(c0).blkInfo(c1).INGOTPOS & " "
'            sql = sql & "and INGOTPOS = " & GrpInfo(c0).blkInfo(c1).INGOTPOS & " " '2001/11/14 S.Sano
''2001/11/14 S.Sano            sql = sql & "and (INGOTPOS + LENGTH) > " & GrpInfo(c0).blkInfo(c1).INGOTPOS & " "
'
'            'データを抽出する
'            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'            RecCount = rs.RecordCount
'            If RecCount = 0 Then
'                GrpInfo(c0).blkInfo(c1).topHin.hinban = ""
'                GrpInfo(c0).blkInfo(c1).topHin.factory = ""
'                GrpInfo(c0).blkInfo(c1).topHin.opecond = ""
'                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = 0
'            Else
'                GrpInfo(c0).blkInfo(c1).topHin.hinban = rs("HINBAN")
'                GrpInfo(c0).blkInfo(c1).topHin.factory = rs("FACTORY")
'                GrpInfo(c0).blkInfo(c1).topHin.opecond = rs("OPECOND")
'                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = rs("REVNUM")
'            End If
'            rs.Close
'
'
'            sql = "select HINBAN, REVNUM, FACTORY, OPECOND "
'            sql = sql & "from  TBCME041 "
'            sql = sql & "where CRYNUM='" & GrpInfo(c0).Crynum & "' "
'            sql = sql & "  and INGOTPOS            <  " & GrpInfo(c0).blkInfo(c1).INGOTPOS + GrpInfo(c0).blkInfo(c1).LENGTH & " "
'            sql = sql & "  and (INGOTPOS + LENGTH) >= " & GrpInfo(c0).blkInfo(c1).INGOTPOS + GrpInfo(c0).blkInfo(c1).LENGTH & " "
'
'            'データを抽出する
'            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'            RecCount = rs.RecordCount
'            If RecCount = 0 Then
'                GrpInfo(c0).blkInfo(c1).botHin.hinban = ""
'                GrpInfo(c0).blkInfo(c1).botHin.factory = ""
'                GrpInfo(c0).blkInfo(c1).botHin.opecond = ""
'                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = 0
'            Else
'                GrpInfo(c0).blkInfo(c1).botHin.hinban = rs("HINBAN")
'                GrpInfo(c0).blkInfo(c1).botHin.factory = rs("FACTORY")
'                GrpInfo(c0).blkInfo(c1).botHin.opecond = rs("OPECOND")
'                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = rs("REVNUM")
'            End If
'            rs.Close
'        Next
'    Next
#End If

Debug.Print " 4:" & Time

    '求めた情報からグループを求める
    GrpCount1 = 0
    GrpCount2 = 0
    For c0 = 1 To MaxRec
        GrpCount1 = GrpCount1 + 1
        GrpCount2 = GrpCount2 + 1
        RecCount = UBound(GrpInfo(c0).blkInfo())
        For c1 = 1 To RecCount
            Dim wTopHin As cmkc001b_Wait3_HINBAN
            Dim wBotHin As cmkc001b_Wait3_HINBAN

            wTopHin = GrpInfo(c0).blkInfo(c1).topHin
            wBotHin = GrpInfo(c0).blkInfo(c1 - 1).botHin

            'ブロック切れ目で品番が変われば別グループと判断する
            Select Case c1
            Case 1
                GrpInfo(c0).blkInfo(c1).GRPFLG1 = GrpCount1
            Case Else
                If (wTopHin.factory <> wBotHin.factory) Or (wTopHin.hinban <> wBotHin.hinban) Or _
                   (wTopHin.opecond <> wBotHin.opecond) Or (wTopHin.REVNUM <> wBotHin.REVNUM) Then
                    GrpCount1 = GrpCount1 + 1
                End If
                GrpInfo(c0).blkInfo(c1).GRPFLG1 = GrpCount1
            End Select

            '同一グループ内で、工程違いのブロックが存在した場合、同一グループ内の
            '小グループとしてグループ分けする。
            'CC710以外なら対象外としグループ判定をしない
            If GrpInfo(c0).blkInfo(c1).NOWPROC = PROCD_WFC_HARAIDASI And GrpInfo(c0).blkInfo(c1).HOLDCLS = "0" Then
                Select Case c1
                Case 1
                    GrpInfo(c0).blkInfo(c1).GRPFLG2 = GrpCount2
                Case Else
                    If (wTopHin.factory <> wBotHin.factory) Or (wTopHin.hinban <> wBotHin.hinban) Or _
                       (wTopHin.opecond <> wBotHin.opecond) Or (wTopHin.REVNUM <> wBotHin.REVNUM) Then
                        GrpCount2 = GrpCount2 + 1
                    End If
                    GrpInfo(c0).blkInfo(c1).GRPFLG2 = GrpCount2
                End Select
            Else
                GrpCount2 = GrpCount2 + 1
                GrpInfo(c0).blkInfo(c1).GRPFLG2 = 0
            End If
        Next
    Next

Debug.Print " 5:" & Time

    '求めた情報から表示色を求める
    For c0 = 1 To MaxRec
        RecCount = UBound(GrpInfo(c0).blkInfo())
        ColorFlag = False
        CheckPoint = 0
        For c1 = 1 To RecCount
            If CheckPoint > 0 Then
                If GrpInfo(c0).blkInfo(c1).GRPFLG1 <> GrpInfo(c0).blkInfo(CheckPoint).GRPFLG1 Then
                    For c2 = CheckPoint To c1 - 1
                        GrpInfo(c0).blkInfo(c2).COLORFLG = ColorFlag
                    Next
                    ColorFlag = False
                    CheckPoint = c1
                End If
            Else
                CheckPoint = c1
            End If
            If CheckPoint > 0 Then
                If (GrpInfo(c0).blkInfo(c1).NOWPROC = PROCD_SETUDAN) Or _
                   (GrpInfo(c0).blkInfo(c1).NOWPROC = PROCD_KESSYOU_SOUGOUHANTEI) Or _
                   (GrpInfo(c0).blkInfo(c1).NOWPROC = PROCD_KESSYOU_SAISYUU_HARAIDASI) Or _
                   (GrpInfo(c0).blkInfo(c1).HOLDCLS = "1") Then
                    ColorFlag = True
                End If
            End If
        Next
        For c1 = CheckPoint To RecCount
            GrpInfo(c0).blkInfo(c1).COLORFLG = ColorFlag
        Next
    Next

Debug.Print " 6:" & Time

    For c0 = 1 To MaxRec
        RecCount = UBound(GrpInfo(c0).blkInfo())
        For c1 = 1 To RecCount
            For c2 = 1 To TotalBlk
                If Wd3(c2).BLOCKID = GrpInfo(c0).blkInfo(c1).BLOCKID Then
                    LWD(c2).flag = GrpInfo(c0).blkInfo(c1).COLORFLG
                    LWD(c2).Grp = GrpInfo(c0).blkInfo(c1).GRPFLG2
                    Exit For
                End If
            Next
        Next
    Next

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    cmkc001b_DBDataCheck5 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function



Public Function GetListData(records() As type_DBDRV_scmzc_fcmkc001b_Disp52, NOWPROC As String) As FUNCTION_RETURN

'内部関数 ブロックID、更新日付取得（払出待ち、抜試指示待ち用）

    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim recCnt      As Long         'レコード数
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim BlockIdBuf  As String
    Dim lLp         As Long         '2007/08/30 SPK Tsutsumi Add
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function GetListData"

    GetListData = FUNCTION_RETURN_SUCCESS

    sql = "select "
    sql = sql & " XC2.XTALC2, "
    sql = sql & " XC2.INPOSC2, "
    sql = sql & " XC2.GNLC2, "
    sql = sql & " XC2.HOLDBC2, "
    sql = sql & " XC2.HOLDCC2, "
    sql = sql & " XC2.HOLDKTC2, "
    sql = sql & " XC2.LBLFLGC2,"
    sql = sql & " XCA.CRYNUMCA, "
    sql = sql & " XCA.INPOSCA, "
    sql = sql & " XCA.KDAYCA, "
    sql = sql & " XCA.HOLDBCA, "
    sql = sql & " XCA.HINBCA, "             ' 品番
    sql = sql & " XCA.REVNUMCA, "           ' 製品番号改訂番号
    sql = sql & " XCA.FACTORYCA, "          ' 工場
    sql = sql & " XCA.OPECA, "              ' 操業条件
    sql = sql & " XCA.GNLCA, "              ' 長さ
    sql = sql & " XCA.GNWCA, "              ' 重量
    sql = sql & " S.HSXTYPE, "              ' 品ＳＸタイプ
    sql = sql & " S.HSXCDIR "               ' 品ＳＸ結晶面方位
    sql = sql & ",XC1.PUPTNC1 "             ' 引上ﾊﾟﾀｰﾝ
    sql = sql & ",XC2.KIKBNC2 "             ' 期判別区分
    sql = sql & ",XC2.PLANTCATC2 "          ' 向先
    sql = sql & " from "
    sql = sql & " XSDCA XCA, TBCME018 S , XSDC2 XC2 "
    sql = sql & ",XSDC1 XC1 "                       '引上ﾊﾟﾀｰﾝ追加対応
    sql = sql & " where "
    sql = sql & " XCA.CRYNUMCA = XC2.CRYNUMC2 "
    sql = sql & " and XCA.HINBCA = S.HINBAN "
    sql = sql & " and XCA.REVNUMCA = S.MNOREVNO "
    sql = sql & " and XCA.FACTORYCA = S.FACTORY "
    sql = sql & " and XCA.OPECA = S.OPECOND "
    sql = sql & " and XCA.GNWKNTCA ='" & NOWPROC & "' "
    sql = sql & " and XCA.LIVKCA='0' "
    sql = sql & " and XC2.XTALC2 = XC1.XTALC1(+) "    '引上ﾊﾟﾀｰﾝ追加対応
    
    If NOWPROC = "CC705" Then
    
    End If
    sql = sql & " order by XCA.CRYNUMCA, XCA.INPOSCA "

    'データを抽出する
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    'レコードがない場合正常終了
    If rs.RecordCount = 0 Then
       rs.Close
       ReDim records(0)
       GoTo proc_exit
    End If
    
    BlockIdBuf = vbNullString
    recCnt = rs.RecordCount
    j = 0
    For i = 1 To recCnt
        DoEvents
        'ブロックID等の格納
        If rs("CRYNUMCA") <> BlockIdBuf Then
        
            j = j + 1
            ReDim Preserve records(j)
            With records(j)
                .CRYNUM = rs("XTALC2")
                .INGOTPOS = rs("INPOSCA")
                .BLOCKID = rs("CRYNUMCA")   ' ブロックID
                .UPDDATE = rs("KDAYCA")     ' 更新日付
                .HOLDCLS = rs("HOLDBCA")    ' ホールド区分
                BlockIdBuf = records(j).BLOCKID
                .HSXTYPE = rs("HSXTYPE")
                .HSXCDIR = rs("HSXCDIR")
                .INPOS = rs("INPOSC2")
                .LENGTH = rs("GNLC2")
                .PUPTN = rs("PUPTNC1")      ' 引上ﾊﾟﾀｰﾝ
                .HOLDB = rs("HOLDBC2")
                .HOLDC = rs("HOLDCC2")
                If IsNull(rs("HOLDKTC2")) = False Then .HOLDKT = rs("HOLDKTC2")
            End With
            
            k = 1
        End If
        
        '品番の格納
        ReDim Preserve records(j).hinM(1)
        records(j).hinM(1).HIN.hinban = rs("HINBCA")
        records(j).hinM(1).HIN.mnorevno = rs("REVNUMCA")
        records(j).hinM(1).HIN.factory = rs("FACTORYCA")
        records(j).hinM(1).HIN.opecond = rs("OPECA")
        records(j).hinM(1).LENGTH = rs("GNLCA")
        records(j).hinM(1).Weight = rs("GNWCA")
        
        ' 向先
        If IsNull(rs("PLANTCATC2")) = False Then
            For lLp = 0 To UBound(s_Mukesaki)
                If rs("PLANTCATC2") = s_Mukesaki(lLp).sMukeCode Then
                    records(j).hinM(1).HIN.sMukesaki = s_Mukesaki(lLp).sMukeName
                End If
            Next lLp
        End If

        k = k + 1
        rs.MoveNext
    Next i
    rs.Close
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    GetListData = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'-------------------------------------------------------------------------CC720から
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:pBlkHinMap() ,O  ,typ_BlkHinMap    ,ブロック品番情報
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,読み込みの成否
'説明      :
'履歴      :2002/04/22 作成 野村
Public Function DBDRV_scmzc_fcmkc001h_Disp22(pBlkHinMap() As typ_BlkHinMap, formNum As Integer) As FUNCTION_RETURN
Dim sql     As String
Dim rs      As OraDynaset
Dim recCnt  As Long
Dim i       As Long
Dim j       As Long ' 2007/09/12 SPK Tsutsumi Add
Dim sBuff   As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc034_SQL.bas -- Function DBDRV_scmzc_fcmkc001h_Disp2"

    ''ブロック内の品番構成を取得する (ブロックID, 品番, 製品長, 実長さ)
'    sql = "select X.CRYNUMCA  as BLOCKID, "
'    sql = sql & " B.PASSFLAG, "
'    sql = sql & " X.HINBCA    as HINBAN, "
'    sql = sql & " X.REVNUMCA  as REVNUM, "
'    sql = sql & " X.FACTORYCA as FACTORY, "
'    sql = sql & " X.OPECA     as OPECOND, "
'    sql = sql & " X.GNLCA     as HINLEN, "
'    sql = sql & " X.GNLCA + case when X.INPOSCA = (select max(INPOSCA) from XSDCA where CRYNUMCA = B.BLOCKID and LIVKCA = '0')"
'    sql = sql & "           then (J.P1BDLEN+J.P2BDLEN+J.P3BDLEN+J.P4BDLEN+J.P5BDLEN) "
'    sql = sql & "           else 0 end as REALLEN ,"
'    sql = sql & " X.INPOSCA as INPOSCA"
'    sql = sql & " ,X.PLANTCATCA as PLANTCATCA"
'    sql = sql & " from  XSDCA X, TBCME040 B, TBCMJ010 J "
'
'    sql = sql & " where X.CRYNUMCA = B.BLOCKID "
'    If formNum = 1 Then
'        sql = sql & "   and X.GNWKNTCA = 'CC720' "
'    ElseIf formNum = 2 Then
'        sql = sql & "   and X.NEWKNTCA = 'CC705' "
'    End If
'    sql = sql & "   and X.LIVKCA   = '0' "
'    sql = sql & "   and B.DELCLS   = '0'"
'    sql = sql & "   and J.CRYNUM   = B.CRYNUM "
'    sql = sql & "   and J.INGOTPOS = B.INGOTPOS"
'    sql = sql & "   and J.TRANCNT  = (select max(TRANCNT) from TBCMJ010 where CRYNUM=J.CRYNUM and INGOTPOS=J.INGOTPOS)"
'    If formNum = 1 Then
'        sql = sql & " and (X.PLANTCATCA =14  or X.PLANTCATCA =15 or X.PLANTCATCA =16)"
'    ElseIf formNum = 2 Then
'        sql = sql & " and (X.PLANTCATCA ='ZZ'  or X.PLANTCATCA ='ZX')"
'    End If
'    sql = sql & " order by B.BLOCKID, X.INPOSCA"

    sql = "select X.CRYNUMCA  as BLOCKID, "
    sql = sql & " X.HINBCA    as HINBAN, "
    sql = sql & " X.REVNUMCA  as REVNUM, "
    sql = sql & " X.FACTORYCA as FACTORY, "
    sql = sql & " X.OPECA     as OPECOND, "
    sql = sql & " X.GNLCA     as HINLEN, "
    sql = sql & " X.INPOSCA as INPOSCA, "
    sql = sql & " X.PLANTCATCA as PLANTCATCA"
    sql = sql & " from  XSDCA X, XSDC2 B"
    sql = sql & " where X.CRYNUMCA = B.CRYNUMC2 "
    If formNum = 1 Then
        sql = sql & "   and X.GNWKNTCA = 'CC720' "
    ElseIf formNum = 2 Then
        sql = sql & "   and X.NEWKNTCA = 'CC700' "
    End If
    sql = sql & "   and X.LIVKCA   = '0' "
    If formNum = 1 Then
'        sql = sql & " and (X.PLANTCATCA =14  or X.PLANTCATCA =15 or X.PLANTCATCA =16)"
        sql = sql & " and (X.PLANTCATCA ='14'  or X.PLANTCATCA ='15' or X.PLANTCATCA ='16')"
    ElseIf formNum = 2 Then
        sql = sql & " and (X.PLANTCATCA ='ZZ'  or X.PLANTCATCA ='ZX')"
    End If
    sql = sql & " order by X.CRYNUMCA, X.INPOSCA"

Debug.Print "Disp22 " & sql
    Set rs = OraDB.CreateDynaset(sql, ORADB_DEFAULT)
    recCnt = rs.RecordCount
    If recCnt <= 0 Then
        ReDim pBlkHinMap(0)
    Else
        ReDim pBlkHinMap(1 To recCnt)
        For i = 1 To recCnt
            With pBlkHinMap(i)
                .BLOCKID = rs("BLOCKID")
                .HIN.hinban = rs("HINBAN")
                .HIN.mnorevno = rs("REVNUM")
                .HIN.factory = rs("FACTORY")
                .HIN.opecond = rs("OPECOND")
                .HinLen = rs("HINLEN")
                '.REALLEN = rs("REALLEN")
                .INPOSCA = rs("INPOSCA")

'                If IsNull(rs("PASSFLAG")) = True Then
'                    sBuff = ""
'                Else
'                    sBuff = rs("PASSFLAG")
'                End If
'                .PASSFLAG = vbNullString & sBuff
                
                If IsNull(rs("PLANTCATCA")) = True Then
                    .PLANTCATCA = ""
                Else
                    For j = 0 To UBound(s_Mukesaki)
                        If s_Mukesaki(j).sMukeCode = rs("PLANTCATCA") Then
                            .PLANTCATCA = ""
                            Exit For
                        End If
                    Next j
                End If
            End With
            rs.MoveNext
        Next
    End If
    rs.Close

    DBDRV_scmzc_fcmkc001h_Disp22 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmkc001h_Disp22 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'------------------------------------------------------------------------
''2008/05/30 SHINDOH 40の処理
'------------------------------------------------------------------------

'全面変更 2003/10/17 SystemBrain
'履歴    :2008/03/31 青柳 ①現行SXL確定実行時からSXLマップ受信時に送信ﾀｲﾐﾝｸﾞ変更する。
'                         ②サンプルIDを反映元のサンプルID(結晶サンプルID含む)ではなく、代表サンプルIDに変更する。
'                         ③GB7/GB8/GB9のSXL確定日付を一致させる。
Public Function WriteX01n(ByVal DoProc%, ByVal blkID$, ByVal WfCnt%, errmsg$) As FUNCTION_RETURN
    Dim recX001(1 To 2)     As c_cmzcrec
    Dim recX002(1 To 2)     As c_cmzcrec
    Dim recX003(1 To 2)     As c_cmzcrec        'GD検査測定点データ
    Dim recX004(1 To 2)     As c_cmzcrec        'EP検査書
    Dim recX005(1 To 2)     As c_cmzcrec        'EP測定点ﾃﾞｰﾀ
    Dim i                   As Integer
    Dim j                   As Integer
    Dim rs                  As OraDynaset
    Dim sql                 As String
    Dim XlSmpPos(1 To 2)    As Integer
    Dim CRYNUM              As String
    Dim sBlkID(1 To 2)      As String       'XSDCW BLOCKID格納
    Dim smpId(2)            As String
    Dim HIN                 As tFullHinban
    Dim iX011cnt            As Integer      '08/09/12 ooba
    Dim recXSDCS(1 To 2)    As c_cmzcrec        '新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ)
    Dim recXSDCW(1 To 2)    As c_cmzcrec        '新ｻﾝﾌﾟﾙ管理(SXL)
    Dim recE037             As c_cmzcrec        '結晶情報
    Dim recXSDC1            As c_cmzcrec        '結晶引上
    Dim recXSDC2            As c_cmzcrec
    Dim recX011(1 To 2)     As c_cmzcrec
    Dim recX012(1 To 2)     As c_cmzcrec
    Dim recX013(1 To 2)     As c_cmzcrec
    Dim Jiltuseki           As Judg_Kakou
    Dim sKMGCSHN            As String

    Dim RsHIN       As tFullHinban  '比抵抗(Rs)仕様取得品番
    Dim sRsData(10) As String       '比抵抗(Rs)ﾃﾞｰﾀ
'    Dim sRsPtn      As String       '比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝ
    Dim sRsPtn(2)   As String       '比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝ
    Dim sPos        As String       'SXL位置(TOP/BOT)
    Dim gSmpID(2)   As String       'TBCMX003用サンプルID
    Dim sErrMsg     As String       'ｴﾗｰﾒｯｾｰｼﾞ
    Dim nowtime     As Date  '③GB7/GB8/GB9のSXL確定日付を一致させる。
    
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function WriteX00n"
    
    WriteX01n = FUNCTION_RETURN_FAILURE
    
    '③GB7/GB8/GB9のSXL確定日付を一致させる。
    nowtime = getSvrTime()    'サーバーの時間を取得

    ''SXLの品番を取得する
    sql = ""
    sql = sql & " SELECT"
    sql = sql & "   HINBCA as HINBAN"           ''品番
    sql = sql & "  ,REVNUMCA as REVNUM"         ''製品番号改訂番号
    sql = sql & "  ,FACTORYCA as FACTORY"       ''工場
    sql = sql & "  ,OPECA as OPECOND"           ''操業条件
    sql = sql & "  ,PLANTCATCA as PLANTCAT"     ''向先  2007/09/04 SPK Tsutsumi Add
    sql = sql & " FROM"
    sql = sql & "   XSDCA"
    sql = sql & " WHERE CRYNUMCA = '" & blkID$ & "'"
    sql = sql & " and LIVKCA = '0'"             '08/07/22 ooba
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount < 1 Then
        errmsg = "XSDCA:" & rs.RecordCount
        rs.Close
        GoTo proc_exit
    End If
    HIN.hinban = rs!hinban
    HIN.mnorevno = rs!REVNUM
    HIN.factory = rs!factory
    HIN.opecond = rs!opecond
    HIN.sMukesaki = rs!PLANTCAT
    Set rs = Nothing

    
    '-------------------- XSDCSの読み込み ----------------------------------------
    For j = 1 To 2
        If j = 1 Then
            '近いXL測定位置(FROM)を求める
            sql = "select * from XSDCS where CRYNUMCS = '" & blkID & "' and "
            sql = sql & "TBKBNCS = 'T' and LIVKCS = '0'"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount <> 1 Then
                errmsg = "XSDCS:From"
                Set rs = Nothing
                GoTo proc_exit
            End If
            Set recXSDCS(1) = New c_cmzcrec
            recXSDCS(1).CopyFromRs "XSDCS", rs
            Set rs = Nothing
            XlSmpPos(1) = recXSDCS(1)("INPOSCS").Value
        ElseIf j = 2 Then
            '近いXL測定位置(TO)を求める
            sql = "select * from XSDCS where CRYNUMCS = '" & blkID & "' and "
            sql = sql & "TBKBNCS = 'B' and LIVKCS = '0'"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount <> 1 Then
                errmsg = "XSDCS:To"
                Set rs = Nothing
                GoTo proc_exit
            End If
            Set recXSDCS(2) = New c_cmzcrec
            recXSDCS(2).CopyFromRs "XSDCS", rs
            Set rs = Nothing
            XlSmpPos(2) = recXSDCS(2)("INPOSCS").Value
        End If
    Next j

    '-------------------- TBCME037の読み込み ----------------------------------------
    CRYNUM = left$(blkID, 9) & "000"        ' 結晶番号
    sql = "select * from TBCME037 where (CRYNUM='" & CRYNUM & "')"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        errmsg = "TBCME037"
        Set rs = Nothing
        GoTo proc_exit
    End If
    Set recE037 = New c_cmzcrec
    recE037.CopyFromRs "TBCME037", rs
    Set rs = Nothing

    '-------------------- XSDC1の読み込み ----------------------------------------
    sql = "select * from XSDC1 where (XTALC1='" & CRYNUM & "')"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        errmsg = "XSDC1"
        Set rs = Nothing
        GoTo proc_exit
    End If
    Set recXSDC1 = New c_cmzcrec
    recXSDC1.CopyFromRs "XSDC1", rs
    Set rs = Nothing
    
    '-------------------- XSDC2の読み込み ----------------------------------------
    sql = "select * from XSDC2 where (CRYNUMC2='" & blkID & "')"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        errmsg = "XSDC2"
        Set rs = Nothing
        GoTo proc_exit
    End If
    Set recXSDC2 = New c_cmzcrec
    recXSDC2.CopyFromRs "XSDC2", rs
    Set rs = Nothing

    '-------------------- TBCME001の読み込み ----------------------------------------
    sql = "SELECT KMGCSHN FROM TBCME001 WHERE HINBAN = '" & HIN.hinban & "'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        errmsg = "TBCME001"
        Set rs = Nothing
        GoTo proc_exit
    Else
        sKMGCSHN = rs.Fields("KMGCSHN")
    End If
    Set rs = Nothing
                    
    '-------------------- TBCMI002の読み込み ----------------------------------------
    If scmzc_getKakouJiltuseki(blkID, Jiltuseki) = FUNCTION_RETURN_FAILURE Then
        errmsg = "TBCMI002"
        GoTo proc_exit
    End If

    'TBCMX011処理回数取得　08/09/12 ooba
    If GetTBCMX011cnt(blkID, iX011cnt) = FUNCTION_RETURN_FAILURE Then
        errmsg = "TBCMX011"
        GoTo proc_exit
    End If
    
    '==============================================
    '　各種実績データの取得・設定
    '==============================================
    For i = 1 To 2
        '-------------------- TBCMX011固定情報データ設定 ----------------------------------------
        Set recX011(i) = New c_cmzcrec
        recX011(i).TABLENAME = "TBCMX011"
        recX011(i).SetRecDefault
        
        With recX011(i)
            .Fields("BLOCKID").Value = blkID                                'BLOCKID
            .Fields("FROMTOKBN").Value = CStr(i)                            'FROMTO区分
''            If DoProc = 0 Then
''                .Fields("TRANCNT").Value = 1                                '実績入力は１固定
''            Else
''                .Fields("TRANCNT").Value = "(SELECT NVL(MAX(TRANCNT),0) + 1 FROM TBCMX011" & _
''                                              " WHERE BLOCKID = '" & .Fields("BLOCKID").Value & "'" & _
''                                                " AND FROMTOKBN = '" & .Fields("FROMTOKBN").Value & "')"
''            End If
            .Fields("TRANCNT").Value = iX011cnt     '08/09/12 ooba
            .Fields("STCID").Value = IIf(Trim(f_cmbc032_2.lblSTCID.Caption) = "", "", f_cmbc032_2.lblSTCID.Caption)
            .Fields("HINBAN").Value = recXSDCS(i)("HINBCS").Value
            .Fields("REVNUM").Value = recXSDCS(i)("REVNUMCS").Value
            .Fields("FACTORY").Value = recXSDCS(i)("FACTORYCS").Value
            .Fields("OPECOND").Value = recXSDCS(i)("OPECS").Value
            If Trim(f_cmbc032_2.lblSTCID.Caption) = "" Then
                .Fields("STCKNNUM").Value = ""
            Else
                .Fields("STCKNNUM").Value = sKMGCSHN
            End If
            .Fields("CRYNUM").Value = recXSDCS(i)("XTALCS").Value
            .Fields("CRYDECDATE").Value = nowtime
            .nowtime = nowtime
            .Fields("PLUPDATE").Value = recXSDC1("TDAYC1").Value
            .Fields("UPLENGTH").Value = recE037("UPLENGTH").Value
            .Fields("FREELENG").Value = recE037("FREELENG").Value
            .Fields("INGOTPOS").Value = recXSDCS(i)("INPOSCS").Value
            .Fields("BLKLEN").Value = recXSDC2("REALLC2").Value
            .Fields("BLKWGHT").Value = recXSDC2("REALWC2").Value
            .Fields("LENGTH").Value = recXSDC2("GNLC2").Value
            .Fields("WEIGHT").Value = recXSDC2("GNWC2").Value
            .Fields("MCNO").Value = recE037("PRODCOND").Value
            .Fields("PGID").Value = recE037("PGID").Value
            If i = 1 Then
                .Fields("DM1").Value = Jiltuseki.top(1)
                .Fields("DM2").Value = Jiltuseki.top(2)
            Else
                .Fields("DM1").Value = Jiltuseki.TAIL(1)
                .Fields("DM2").Value = Jiltuseki.TAIL(2)
            End If
            .Fields("NCHDPTH").Value = Jiltuseki.DPTH(1)
'            .Fields("CHARGE").Value = recE037("CHARGE").Value / 1000
            .Fields("CHARGE").Value = recXSDC1("SUICHARGE").Value / 1000
            .Fields("SEED").Value = recE037("SEED").Value
            .Fields("REGDATE").Value = "SYSDATE"
            .Fields("SENDFLAG").Value = 0
        End With
        
        '-------------------- TBCMX012固定情報データ設定 ----------------------------------------
        Set recX012(i) = New c_cmzcrec
        recX012(i).TABLENAME = "TBCMX012"
        recX012(i).SetRecDefault
        
        With recX012(i)
            .Fields("BLOCKID").Value = blkID                                'BLOCKID
            .Fields("FROMTOKBN").Value = CStr(i)                            'FROMTO区分
            .Fields("TRANCNT").Value = iX011cnt                             '08/09/12 ooba
            .Fields("STCID").Value = IIf(Trim(f_cmbc032_2.lblSTCID.Caption) = "", "", f_cmbc032_2.lblSTCID.Caption)
            .Fields("HINBAN").Value = recXSDCS(i)("HINBCS").Value
            .Fields("REVNUM").Value = recXSDCS(i)("REVNUMCS").Value
            .Fields("FACTORY").Value = recXSDCS(i)("FACTORYCS").Value
            .Fields("OPECOND").Value = recXSDCS(i)("OPECS").Value
            If Trim(f_cmbc032_2.lblSTCID.Caption) = "" Then
                .Fields("STCKNNUM").Value = ""
            Else
                .Fields("STCKNNUM").Value = sKMGCSHN
            End If
            .Fields("CRYNUM").Value = recXSDCS(i)("XTALCS").Value
            .Fields("REGDATE").Value = "SYSDATE"
            .Fields("SENDFLAG").Value = 0
        End With
        
        '-------------------- TBCMX013固定情報データ設定 ----------------------------------------
        Set recX013(i) = New c_cmzcrec
        recX013(i).TABLENAME = "TBCMX013"
        recX013(i).SetRecDefault
        
        With recX013(i)
            .Fields("BLOCKID").Value = blkID                                'BLOCKID
            .Fields("FROMTOKBN").Value = CStr(i)                            'FROMTO区分
            .Fields("TRANCNT").Value = iX011cnt                             '08/09/12 ooba
            .Fields("STCID").Value = IIf(Trim(f_cmbc032_2.lblSTCID.Caption) = "", "", f_cmbc032_2.lblSTCID.Caption)
            .Fields("HINBAN").Value = recXSDCS(i)("HINBCS").Value
            .Fields("REVNUM").Value = recXSDCS(i)("REVNUMCS").Value
            .Fields("FACTORY").Value = recXSDCS(i)("FACTORYCS").Value
            .Fields("OPECOND").Value = recXSDCS(i)("OPECS").Value
            If Trim(f_cmbc032_2.lblSTCID.Caption) = "" Then
                .Fields("STCKNNUM").Value = ""
            Else
                .Fields("STCKNNUM").Value = sKMGCSHN
            End If
            .Fields("CRYNUM").Value = recXSDCS(i)("XTALCS").Value
            .Fields("REGDATE").Value = "SYSDATE"
            .Fields("SENDFLAG").Value = 0
        End With
                
        If i = 1 Then sPos = "TOP" Else sPos = "BOT"
        
        '-------------------- (結晶Rs)結晶抵抗実績(TBCMJ002)データ取得設定 ----------------------------------------
        If getTBCMJ002(CRYNUM, recXSDCS(), i, HIN, recX011(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "J002:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '-------------------- (結晶Oi)結晶Oi実績(TBCMJ003)データ取得設定 ----------------------------------------
        If getTBCMJ003(CRYNUM, recXSDCS(i), HIN, recX011(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "J003:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '-------------------- (Cs)Cs実績(TBCMJ004)データ取得設定 ----------------------------------------
        '品番＆ｴﾗｰﾒｯｾｰｼﾞ追加
        If getTBCMJ004(CRYNUM, recXSDCS(i), HIN, recX011(i), sErrMsg) = FUNCTION_RETURN_FAILURE Then
            If sErrMsg = "" Then
                errmsg = "J004:" & XlSmpPos(i)
            Else
                errmsg = sErrMsg
            End If
            GoTo proc_exit
        End If

        '-------------------- (結晶OSF1～4)結晶OSF実績(TBCMJ005)データ取得設定 ----------------------------------------
        For j = 1 To 4
            If getTBCMJ005(CRYNUM, recXSDCS(i), j, recX011(i), recX012(i)) = FUNCTION_RETURN_FAILURE Then
                errmsg = "J005-" & j & ":" & XlSmpPos(i)
                GoTo proc_exit
            End If
        Next

        '-------------------- (結晶BMD1～3)結晶BMD実績(TBCMJ008)データ取得設定 ----------------------------------------
        For j = 1 To 3
            If getTBCMJ008(CRYNUM, recXSDCS(i), j, recX011(i), recX012(i)) = FUNCTION_RETURN_FAILURE Then
                errmsg = "J008-" & j & ":" & XlSmpPos(i)
                GoTo proc_exit
            End If
        Next

        '-------------------- (GD)GD実績(TBCMJ006)データ取得設定 ----------------------------------------
        If getTBCMJ006(CRYNUM, recXSDCS(i), recX011(i), recX013(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "J006:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '-------------------- (LT)LT実績(TBCMJ007)データ取得設定 ----------------------------------------
        If getTBCMJ007(CRYNUM, recXSDCS(i), HIN, i, recX011(i), recX012(i)) = FUNCTION_RETURN_FAILURE Then  '05/12/05 ooba
            errmsg = "J007:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '==============================================
        '　TBCMX011 に書き込む
        '==============================================
        With recX011(i)
            sql = .SqlInsert
            
            If 0 >= OraDB.ExecuteSQL(sql) Then
                WriteX01n = FUNCTION_RETURN_FAILURE
            End If
            
        End With

        '==============================================
        '　TBCMX012 に書き込む
        '==============================================
        '変更時も登録　08/09/12 ooba
''        If DoProc = 0 Then
          With recX012(i)
            sql = .SqlInsert
            
            If 0 >= OraDB.ExecuteSQL(sql) Then
                WriteX01n = FUNCTION_RETURN_FAILURE
            End If
          End With
''        End If
            
            
        '==============================================
        '　TBCMX013 に書き込む
        '==============================================
        '変更時も登録　08/09/12 ooba
''        If DoProc = 0 Then
          With recX013(i)
            sql = .SqlInsert
            
            If 0 >= OraDB.ExecuteSQL(sql) Then
                WriteX01n = FUNCTION_RETURN_FAILURE
            End If
          End With
''        End If
    
    Next
    
    WriteX01n = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    WriteX01n = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :共有サンプルチェック処理
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :inSXLID         , I  ,String            , SXL-ID
'          :inSMPLID        , I  ,String            , ｻﾝﾌﾟﾙID
'          :outSMPLID       , O  ,String            , 共有ｻﾝﾌﾟﾙID(共有でない場合、inSMPLIDを返す)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :指定されたｻﾝﾌﾟﾙIDが全共有かどうかをﾁｪｯｸし、全共有の場合、共有ｻﾝﾌﾟﾙIDを取得し返す
'履歴      :成
Private Function chkComSAMPL(inSXLID As String, inSMPLID As String, outSMPLID As String) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim wXTALCW     As String
    Dim wINPOSCW    As Integer
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function chkComSAMPL"
    
    chkComSAMPL = FUNCTION_RETURN_SUCCESS
    outSMPLID = inSMPLID
    
    '-------------------- 全共有確認(XSDCW) ----------------------------------------
    sql = "select XTALCW, INPOSCW from XSDCW "
    sql = sql & "where SXLIDCW = '" & inSXLID & "' and "
    sql = sql & "      REPSMPLIDCW = '" & inSMPLID & "' and "
    sql = sql & "      (WFINDRSCW = '2' or WFINDRSCW = '0' or WFINDRSCW = ' ' or WFINDRSCW is null) and "
    sql = sql & "      (WFINDOICW = '2' or WFINDOICW = '0' or WFINDOICW = ' ' or WFINDOICW is null) and "
    sql = sql & "      (WFINDB1CW = '2' or WFINDB1CW = '0' or WFINDB1CW = ' ' or WFINDB1CW is null) and "
    sql = sql & "      (WFINDB2CW = '2' or WFINDB2CW = '0' or WFINDB2CW = ' ' or WFINDB2CW is null) and "
    sql = sql & "      (WFINDB2CW = '2' or WFINDB3CW = '0' or WFINDB3CW = ' ' or WFINDB3CW is null) and "
    sql = sql & "      (WFINDL1CW = '2' or WFINDL1CW = '0' or WFINDL1CW = ' ' or WFINDL1CW is null) and "
    sql = sql & "      (WFINDL2CW = '2' or WFINDL2CW = '0' or WFINDL2CW = ' ' or WFINDL2CW is null) and "
    sql = sql & "      (WFINDL3CW = '2' or WFINDL3CW = '0' or WFINDL3CW = ' ' or WFINDL3CW is null) and "
    sql = sql & "      (WFINDL4CW = '2' or WFINDL4CW = '0' or WFINDL4CW = ' ' or WFINDL4CW is null) and "
    sql = sql & "      (WFINDDSCW = '2' or WFINDDSCW = '0' or WFINDDSCW = ' ' or WFINDDSCW is null) and "
    sql = sql & "      (WFINDDZCW = '2' or WFINDDZCW = '0' or WFINDDZCW = ' ' or WFINDDZCW is null) and "
    sql = sql & "      (WFINDSPCW = '2' or WFINDSPCW = '0' or WFINDSPCW = ' ' or WFINDSPCW is null) and "
    sql = sql & "      (WFINDDO1CW = '2' or WFINDDO1CW = '0' or WFINDDO1CW = ' ' or WFINDDO1CW is null) and "
    sql = sql & "      (WFINDDO2CW = '2' or WFINDDO2CW = '0' or WFINDDO2CW = ' ' or WFINDDO2CW is null) and "
    sql = sql & "      (WFINDDO3CW = '2' or WFINDDO3CW = '0' or WFINDDO3CW = ' ' or WFINDDO3CW is null) and "
    sql = sql & "      (WFINDOT1CW = '2' or WFINDOT1CW = '0' or WFINDOT1CW = ' ' or WFINDOT1CW is null) and "
    sql = sql & "      (WFINDOT2CW = '2' or WFINDOT2CW = '0' or WFINDOT2CW = ' ' or WFINDOT2CW is null) and "
    sql = sql & "      (WFINDAOICW = '2' or WFINDAOICW = '0' or WFINDAOICW = ' ' or WFINDAOICW is null) and "
    sql = sql & "      (((WFINDGDCW = '2' or WFINDGDCW = '0' or WFINDGDCW = ' ' or WFINDGDCW is null) and WFHSGDCW = '0') or WFHSGDCW = '1') "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    wXTALCW = rs("XTALCW")      '結晶番号
    wINPOSCW = rs("INPOSCW")    '結晶内位置
    Set rs = Nothing
    
    '-------------------- 共有ｻﾝﾌﾟﾙIDの取得(XSDCW) ----------------------------------------
    sql = "select REPSMPLIDCW from XSDCW "
    sql = sql & "where XTALCW = '" & wXTALCW & "' and "
    sql = sql & "      INPOSCW = '" & wINPOSCW & "' and "
    sql = sql & "      SXLIDCW != '" & inSXLID & "' and "
    sql = sql & "      REPSMPLIDCW != '" & inSMPLID & "' "
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    outSMPLID = rs("REPSMPLIDCW")       '代表ｻﾝﾌﾟﾙID(共有)
    Set rs = Nothing

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    chkComSAMPL = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :結晶抵抗実績(TBCMJ002)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCS()      , I  ,c_cmzcrec         , XSDCS構造体   (新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ))
'          :i               , I  ,Integer           , Top/Bot種別(1:Top, 2:Bot)
'          :hin             , I  ,tFullHinban       , 品番(全品番構造体)
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :結晶抵抗実績(TBCMJ002)からﾃﾞｰﾀを取得し、SXL検査書構造体にｾｯﾄする
'履歴      :
Private Function getTBCMJ002(CRYNUM As String, recXSDCS() As c_cmzcrec, i As Integer, HIN As tFullHinban, recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim k           As Integer
    Dim wMeas1(2)   As Double
    Dim wgtCharge   As Long                 '偏析計算用パラメータ
    Dim wgtTop      As Double               '偏析計算用パラメータ
    Dim wgtTopCut   As Double               '偏析計算用パラメータ
    Dim DM          As Double               '偏析計算用パラメータ
    Dim cc          As type_Coefficient
    Dim CRes        As C_RES                '結晶RS判定構造体
    Dim wComp       As Double
    Dim wHSXRHWYS   As String               '保証方法＿処
    Dim RET As FUNCTION_RETURN
    Dim wStaff      As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ002"
    
    getTBCMJ002 = FUNCTION_RETURN_FAILURE

    With recX001
        .Fields("SXL_RS_SMPPOS").Value = -1                 'SXLRSサンプル測定位置(SXL測定情報)
        .Fields("SXLRS_MEAS1").Value = -1                   'SXLRS_測定値1
        .Fields("SXLRS_MEAS2").Value = -1                   'SXLRS_測定値2
        .Fields("SXLRS_MEAS3").Value = -1                   'SXLRS_測定値3
        .Fields("SXLRS_MEAS4").Value = -1                   'SXLRS_測定値4
        .Fields("SXLRS_MEAS5").Value = -1                   'SXLRS_測定値5
        .Fields("SXLRS_EFEHS").Value = -1                   'SXLRS_実効偏析
        .Fields("SXLRS_RRG").Value = -1                     'SXLRS_RRG
    
        '-------------------- TBCMJ002の読み込み(Rs) ----------------------------------------
        If (recXSDCS(i)("CRYINDRSCS").Value <> "0") And (recXSDCS(i)("CRYRESRS1CS").Value <> "0") Then
            '実効偏析算出の為、Top/Botの両方を取得
            For k = 1 To 2
                sql = "select * from TBCMJ002 "
                sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
                sql = sql & "      SMPLNO = " & recXSDCS(k)("CRYSMPLIDRSCS").Value & " "
                sql = sql & "order by TRANCNT desc"
                sql = "select * from (" & sql & ") where rownum = 1"
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                If k = i Then
                    .Fields("SXL_RS_SMPPOS").Value = rs("POSITION")             'SXLRSサンプル測定位置(SXL測定情報)
                    .Fields("SXLRS_MEAS1").Value = rs("MEAS1")                  'SXLRS_測定値1
                    .Fields("SXLRS_MEAS2").Value = rs("MEAS2")                  'SXLRS_測定値2
                    .Fields("SXLRS_MEAS3").Value = rs("MEAS3")                  'SXLRS_測定値3
                    .Fields("SXLRS_MEAS4").Value = rs("MEAS4")                  'SXLRS_測定値4
                    .Fields("SXLRS_MEAS5").Value = rs("MEAS5")                  'SXLRS_測定値5
                    wStaff = rs("KSTAFFID")                                     '---TEST2004/10
                End If
                wMeas1(k) = rs("MEAS1")                             '実効偏析算出用
                Set rs = Nothing
            Next k
            
            'SXLRS_EFEHS
            If GetCoeffParams(CRYNUM, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
            cc.DUNMENSEKI = AreaOfCircle(DM)
            cc.TOPSMPLPOS = recXSDCS(1)("INPOSCS").Value
            cc.BOTSMPLPOS = recXSDCS(2)("INPOSCS").Value
            cc.CHARGEWEIGHT = wgtCharge
            cc.TOPWEIGHT = wgtTop + wgtTopCut
            cc.TOPRES = wMeas1(1)
            cc.BOTRES = wMeas1(2)
            wComp = CoefficientCalculation(cc)
        
            If wComp = -9999 Then
                wComp = 0                                       'SXLRS_実効偏析
            End If
            .Fields("SXLRS_EFEHS").Value = wComp                'SXLRS_実効偏析
            
            'SXLRS_RRG
            'Cng Start 2011/10/13 Y.Hitomi
            'Cng Start 2011/09/19 Y.Hitomi
            sql = "select HSXRHWYS, HSXRSPOH, HSXRSPOT, HSXRSPOI,HSXRMCAL,HSXRHWYT from TBCME018 where "
'            sql = "select HSXRHWYS, HSXRSPOH, HSXRSPOT, HSXRSPOI,HSXRMCAL from TBCME018 where "
'            sql = "select HSXRHWYS, HSXRSPOH, HSXRSPOT, HSXRSPOI from TBCME018 where "
            'Cng End 2011/09/19 Y.Hitomi
            'Cng End 2011/10/13 Y.Hitomi
            
            sql = sql & " HINBAN = '" & HIN.hinban & "' and "
            sql = sql & " MNOREVNO = " & HIN.mnorevno & " and "
            sql = sql & " FACTORY = '" & HIN.factory & "' and "
            sql = sql & " OPECOND = '" & HIN.opecond & "' "
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
                
            'Cng Start 2011/09/19 Y.Hitomi
            CRes.GuaranteeRes.cBunp = rs("HSXRMCAL")                     ' 品ＳＸ比抵抗分布計算
            'CRes.GuaranteeRes.cBunp = rs("HSXRSPOH")                    ' 品ＳＸ比抵抗測定位置＿方
            'Cng End   2011/09/19 Y.Hitomi
            CRes.GuaranteeRes.cCount = rs("HSXRSPOT")                   ' 品ＳＸ比抵抗測定位置＿点
            CRes.GuaranteeRes.cPos = rs("HSXRSPOI")                     ' 品ＳＸ比抵抗測定位置＿位
            wHSXRHWYS = rs("HSXRHWYS")                                  ' 品ＳＸ比抵抗保証方法＿処
            'Add Start 2011/10/12 Y.Hitomi
            CRes.GuaranteeRes.cObj = rs("HSXRHWYT")                     ' 品ＳＸ比抵抗保証方法＿対
            'Add End 2011/10/12 Y.Hitomi
            Set rs = Nothing
            
            CRes.Res(0) = NtoZ2(.Fields("SXLRS_MEAS1").Value)           'Rs測定値1
            CRes.Res(1) = NtoZ2(.Fields("SXLRS_MEAS2").Value)           'Rs測定値2
            CRes.Res(2) = NtoZ2(.Fields("SXLRS_MEAS3").Value)           'Rs測定値3
            CRes.Res(3) = NtoZ2(.Fields("SXLRS_MEAS4").Value)           'Rs測定値4
            CRes.Res(4) = NtoZ2(.Fields("SXLRS_MEAS5").Value)           'Rs測定値5
            
            ''-----> 2006/06 測定位置による計算は必要なためコメントを外し測定順にデータを戻す処理を追加する
            If Trim(wStaff) <> KSTAFF_J002 Then   '新測定データの場合だけ処理する
                RET = Set_Rs_Ichi(CRes.GuaranteeRes.cCount, CRes.GuaranteeRes.cPos, CRes.Res(0), CRes.Res(1), CRes.Res(2), _
                               CRes.Res(3), CRes.Res(4))
            End If
            
            .Fields("SXLRS_RRG").Value = CryRES_Judg(CRes.Res(), CRes.GuaranteeRes)     'SXLRS_RRG
            
            CRes.Res(0) = NtoZ2(.Fields("SXLRS_MEAS1").Value)           'Rs測定値1
            CRes.Res(1) = NtoZ2(.Fields("SXLRS_MEAS2").Value)           'Rs測定値2
            CRes.Res(2) = NtoZ2(.Fields("SXLRS_MEAS3").Value)           'Rs測定値3
            CRes.Res(3) = NtoZ2(.Fields("SXLRS_MEAS4").Value)           'Rs測定値4
            CRes.Res(4) = NtoZ2(.Fields("SXLRS_MEAS5").Value)           'Rs測定値5

'Cng Start 2011/10/25 Y.Hitomi
            '保証方法="H"、かつ、SXLRS_RRG計算結果が-2(=分布計算未定義）の場合、エラーとする。
            If (wHSXRHWYS = "H") And (.Fields("SXLRS_RRG").Value = -2) Then GoTo proc_exit
                        
'            '保証方法="H"、かつ、SXLRS_RRG計算結果が-1の場合、エラーとする
'            'Cng Start 2011/10/12 Y.Hitomi
'            'If (wHSXRHWYS = "H") And (.Fields("SXLRS_RRG").Value = -1) Then GoTo proc_exit
'            If (wHSXRHWYS = "H") And (.Fields("SXLRS_RRG").Value = -1) And CRes.GuaranteeRes.cObj <> "1" Then GoTo proc_exit
'            'Cng End 2011/10/12 Y.Hitomi
'Cng End 2011/10/25 Y.Hitomi
        End If
    End With

    getTBCMJ002 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ002 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :結晶Oi実績(TBCMJ003)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS構造体   (新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ))
'          :hin             , I  ,tFullHinban       , 品番(全品番構造体)
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :結晶Oi実績(TBCMJ003)からﾃﾞｰﾀを取得し、SXL検査書構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
Private Function getTBCMJ003(CRYNUM As String, recXSDCS As c_cmzcrec, HIN As tFullHinban, recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim COi         As C_Oi                 '結晶Oi判定構造体
    Dim wHSXONHWS   As String               '保証方法＿処
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ003"
    
    getTBCMJ003 = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    With recX001
        .Fields("SXL_OI_SMPPOS").Value = -1                 'SXLOIサンプル測定位置(SXL測定情報)
        .Fields("SXLOI_OIMEAS1").Value = -1                 'SXLOI_Oi測定値1
        .Fields("SXLOI_OIMEAS2").Value = -1                 'SXLOI_Oi測定値2
        .Fields("SXLOI_OIMEAS3").Value = -1                 'SXLOI_Oi測定値3
        .Fields("SXLOI_OIMEAS4").Value = -1                 'SXLOI_Oi測定値4
        .Fields("SXLOI_OIMEAS5").Value = -1                 'SXLOI_Oi測定値5
        .Fields("SXLOI_ORGRES").Value = -1                  'SXLOI_ORG結果
        .Fields("SXLOI_INSPECTWAY").Value = -1              'SXLOI検査方法
    
        '-------------------- TBCMJ003の読み込み(Oi) ----------------------------------------
        If (recXSDCS("CRYINDOICS").Value <> "0") And (recXSDCS("CRYRESOICS").Value <> "0") Then
            sql = "select * from TBCMJ003 "
            sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
            sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDOICS").Value & " "
            sql = sql & "  and TRANCOND = 0 "   'GFAのFTIR換算値取得異常対応 2011/02/28 SETsw kubota
            sql = sql & "order by TRANCNT desc"
            sql = "select * from (" & sql & ") where rownum = 1"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            .Fields("SXL_OI_SMPPOS").Value = rs("POSITION")             'SXLOIサンプル測定位置(SXL測定情報)
''''            .Fields("SXLOI_OIMEAS1").Value = rs("OIMEAS1")              'SXLOI_Oi測定値1
''''            .Fields("SXLOI_OIMEAS2").Value = rs("OIMEAS2")              'SXLOI_Oi測定値2
''''            .Fields("SXLOI_OIMEAS3").Value = rs("OIMEAS3")              'SXLOI_Oi測定値3
''''            .Fields("SXLOI_OIMEAS4").Value = rs("OIMEAS4")              'SXLOI_Oi測定値4
''''            .Fields("SXLOI_OIMEAS5").Value = rs("OIMEAS5")              'SXLOI_Oi測定値5
            'OI_NULL対応　2005/03/08 TUKU START --------------------------------------------------
            If IsNull(rs("OIMEAS1")) = False Then .Fields("SXLOI_OIMEAS1").Value = rs("OIMEAS1") Else .Fields("SXLOI_OIMEAS1").Value = -1  'SXLOI_Oi測定値1
            If IsNull(rs("OIMEAS2")) = False Then .Fields("SXLOI_OIMEAS2").Value = rs("OIMEAS2") Else .Fields("SXLOI_OIMEAS2").Value = -1  'SXLOI_Oi測定値2
            If IsNull(rs("OIMEAS3")) = False Then .Fields("SXLOI_OIMEAS3").Value = rs("OIMEAS3") Else .Fields("SXLOI_OIMEAS3").Value = -1  'SXLOI_Oi測定値3
            If IsNull(rs("OIMEAS4")) = False Then .Fields("SXLOI_OIMEAS4").Value = rs("OIMEAS4") Else .Fields("SXLOI_OIMEAS4").Value = -1  'SXLOI_Oi測定値4
            If IsNull(rs("OIMEAS5")) = False Then .Fields("SXLOI_OIMEAS5").Value = rs("OIMEAS5") Else .Fields("SXLOI_OIMEAS5").Value = -1  'SXLOI_Oi測定値5
            'OI_NULL対応　2005/03/08 TUKU END   --------------------------------------------------
            .Fields("SXLOI_INSPECTWAY").Value = rs("INSPECTWAY")        'SXLOI検査方法
            Set rs = Nothing
        
            'SXLOI_ORG
            'Cng Start 2011/10/13 Y.Hitomi
            'Cng Start 2011/09/19 Y.Hitomi
            sql = "select HSXONHWS, HSXONSPH, HSXONSPT, HSXONSPI,HSXONMCL,HSXONHWT from TBCME019 where "
'            sql = "select HSXONHWS, HSXONSPH, HSXONSPT, HSXONSPI,HSXONMCL from TBCME019 where "
            'sql = "select HSXONHWS, HSXONSPH, HSXONSPT, HSXONSPI from TBCME019 where "
            'Cng End   2011/09/19 Y.Hitomi
            'Cng Start 2011/10/13 Y.Hitomi
            sql = sql & " HINBAN = '" & HIN.hinban & "' and "
            sql = sql & " MNOREVNO = " & HIN.mnorevno & " and "
            sql = sql & " FACTORY = '" & HIN.factory & "' and "
            sql = sql & " OPECOND = '" & HIN.opecond & "' "
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            
            ReDim COi.Oi(4) As Double
            'Cng Start 2011/09/19 Y.Hitomi
            COi.GuaranteeOi.cBunp = rs("HSXONMCL")                      ' 品ＳＸ酸素濃度分布計算
            'COi.GuaranteeOi.cBunp = rs("HSXONSPH")                      ' 品ＳＸ酸素濃度測定位置＿方
            'Cng End   2011/09/19 Y.Hitomi
            
            COi.GuaranteeOi.cCount = rs("HSXONSPT")                     ' 品ＳＸ酸素濃度測定位置＿点
            COi.GuaranteeOi.cPos = rs("HSXONSPI")                       ' 品ＳＸ酸素濃度測定位置＿位
            wHSXONHWS = rs("HSXONHWS")                                  ' 品ＳＸ酸素濃度保証方法＿処
            'Add Start 2011/10/12 Y.Hitomi
            COi.GuaranteeOi.cObj = rs("HSXONHWT")                       ' 品ＳＸ酸素濃度保証方法＿対
            'Add End 2011/10/12 Y.Hitomi
            Set rs = Nothing

            COi.Oi(0) = NtoZ2(.Fields("SXLOI_OIMEAS1").Value)           'Oi測定値1
            COi.Oi(1) = NtoZ2(.Fields("SXLOI_OIMEAS2").Value)           'Oi測定値2
            COi.Oi(2) = NtoZ2(.Fields("SXLOI_OIMEAS3").Value)           'Oi測定値3
            COi.Oi(3) = NtoZ2(.Fields("SXLOI_OIMEAS4").Value)           'Oi測定値4
            COi.Oi(4) = NtoZ2(.Fields("SXLOI_OIMEAS5").Value)           'Oi測定値5
            
            .Fields("SXLOI_ORGRES").Value = CryOi_Judg(COi.Oi(), COi.GuaranteeOi)       'SXLOI_ORG結果
            
'Cng Start 2011/10/25 Y.Hitomi
            '保証方法="H"、かつ、SXLOI_ORG計算結果が-2(=分布計算未定義）の場合、エラーとする。
            If (wHSXONHWS = "H") And (.Fields("SXLOI_ORGRES").Value = -2) Then GoTo proc_exit
            
'            '保証方法="H"、かつ、SXLOI_ORG計算結果が-1の場合、エラーとする。2003/11/21 SystemBrain
''            If (wHSXONHWS = "H") And (.Fields("SXLOI_ORGRES").Value = -1) Then GoTo proc_exit
'            'Cng Start 2011/10/12 Y.Hitomi
'            If (wHSXONHWS = "H") And (.Fields("SXLOI_ORGRES").Value = -1) And COi.GuaranteeOi.cObj <> "1" Then GoTo proc_exit
'            'If (wHSXONHWS = "H") And (.Fields("SXLOI_ORGRES").Value = -1) Then GoTo proc_exit
'            'Cng End 2011/10/12 Y.Hitomi
'Cng End 2011/10/25 Y.Hitomi

        End If
    End With

    getTBCMJ003 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ003 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :Cs実績(TBCMJ004)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS構造体   (新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ))
'          :HIN             , I  ,tFullHinban       , 品番　06/04/20 ooba
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'          :sErrMsg         , O  ,String            , ｴﾗｰﾒｯｾｰｼﾞ　06/04/20 ooba
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :Cs実績(TBCMJ004)からﾃﾞｰﾀを取得し、SXL検査書構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
Private Function getTBCMJ004(CRYNUM As String, recXSDCS As c_cmzcrec, HIN As tFullHinban, _
                             recX001 As c_cmzcrec, sErrMsg As String) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    

    Dim rs2         As OraDynaset
    Dim dCmax       As Double           '仕様(上限値)
    Dim dCmin       As Double           '仕様(下限値)
    Dim iSmpNo      As Long             '推定元ｻﾝﾌﾟﾙNo
    Dim tCsSuitei   As CS_SUITEI_TYPE   'CS推定計算用構造体
    Dim dCsSuitei   As Double           'Cs推定値
    
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ004"
    
    getTBCMJ004 = FUNCTION_RETURN_FAILURE

    sErrMsg = ""        '06/04/20 ooba
    
    '-------------------- 初期ｸﾘｱ ----------------------------------------
    With recX001
        .Fields("SXL_CS_SMPPOS").Value = -1                 'SXLCSサンプル測定位置(SXL測定情報)
        .Fields("SXLCS_CSMEAS").Value = -1                  'SXLCS_Cs実測値
        .Fields("SXLCS_70PPRE").Value = -1                  'SXLCS_70%推定値
        .Fields("SXLCS_BSUIMEAS").Value = -1                'SXLCS_Csﾌﾞﾛｯｸ推定値
    
        '-------------------- TBCMJ004の読み込み(Cs) ----------------------------------------
        If (recXSDCS("CRYINDCSCS").Value <> "0") And (recXSDCS("CRYRESCSCS").Value <> "0") Then
            sql = "select * from TBCMJ004 "
            sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
            sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDCSCS").Value & " "
            sql = sql & "order by TRANCNT desc"
            sql = "select * from (" & sql & ") where rownum = 1"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            .Fields("SXL_CS_SMPPOS").Value = rs("POSITION")             'SXLCSサンプル測定位置(SXL測定情報)
''''            .Fields("SXLCS_CSMEAS").Value = rs("CSMEAS")                'SXLCS_Cs実測値
''''            .Fields("SXLCS_70PPRE").Value = rs("PRE70P")                'SXLCS_70%推定値
            'OI_NULL対応　2005/03/08 TUKU START --------------------------------------------------
            If IsNull(rs("CSMEAS")) = False Then .Fields("SXLCS_CSMEAS").Value = rs("CSMEAS") Else .Fields("SXLCS_CSMEAS").Value = -1  'SXLCS_Cs実測値
            If IsNull(rs("PRE70P")) = False Then .Fields("SXLCS_70PPRE").Value = rs("PRE70P") Else .Fields("SXLCS_70PPRE").Value = -1  'SXLCS_70%推定値
            'OI_NULL対応　2005/03/08 TUKU START --------------------------------------------------
            
            Set rs = Nothing
            
            ''Csﾌﾞﾛｯｸ推定値計算対応　06/04/20 ooba START ======================================>
        
            '実測の場合は｢ﾌﾞﾛｯｸ推定値＝実測値｣
            If recXSDCS("CRYINDCSCS").Value = "1" Then
                .Fields("SXLCS_BSUIMEAS").Value = .Fields("SXLCS_CSMEAS").Value
            Else
                '①推定位置
                tCsSuitei.sInfPos = CStr(recXSDCS("INPOSCS").Value)
                
                '②ｻﾝﾌﾟﾙ位置
                '③ｻﾝﾌﾟﾙ測定値
                '推定元ｻﾝﾌﾟﾙNo取得
                iSmpNo = recXSDCS("CRYSMPLIDCSCS").Value
                
                'ｻﾝﾌﾟﾙ位置＆ｻﾝﾌﾟﾙ測定値取得
                sql = "select POSITION, CSMEAS from TBCMJ004 "
                sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
                sql = sql & "      SMPLNO = " & iSmpNo & " "
                sql = sql & "order by TRANCNT desc"
                
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    sErrMsg = GetMsgStr("ECLC3") & " <ｻﾝﾌﾟﾙ測定値> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                tCsSuitei.sSamplePos = rs("POSITION")       'ｻﾝﾌﾟﾙ位置
                tCsSuitei.sResCs = rs("CSMEAS")             'ｻﾝﾌﾟﾙ測定値
                Set rs = Nothing
                
                '④ﾁｬｰｼﾞ量
                '⑤TOP重量
                sql = "select SUICHARGE, WGHTTOC1, PUTCUTWC1 from XSDC1 "
                sql = sql & "where XTALC1 = '" & CRYNUM & "' "
                
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    sErrMsg = GetMsgStr("ECLC3") & " <ﾁｬｰｼﾞ量> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                'ﾃﾞｰﾀ不正
                If (IsNull(rs("SUICHARGE")) Or IsNull(rs("WGHTTOC1")) Or IsNull(rs("PUTCUTWC1"))) Then
                    sErrMsg = GetMsgStr("ECLC3") & " <ﾁｬｰｼﾞ量> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                
                tCsSuitei.sSiWeight = rs("SUICHARGE")       '推定ﾁｬｰｼﾞ量
                tCsSuitei.sTopWT = CLng(rs("WGHTTOC1")) + CLng(rs("PUTCUTWC1"))     'TOP重量
                Set rs = Nothing
                '｢推定ﾁｬｰｼﾞ量=0｣or｢推定ﾁｬｰｼﾞ量≦TOP重量｣の場合はｴﾗｰとする
                If CLng(tCsSuitei.sSiWeight) = 0 Or _
                   (CLng(tCsSuitei.sSiWeight) <= CLng(tCsSuitei.sTopWT)) Then
                    sErrMsg = GetMsgStr("ECLC3") & " <ﾁｬｰｼﾞ量> "
                    GoTo proc_exit
                End If
                
                '⑥直径
                sql = "select HSXD1CEN from TBCME018 "
                sql = sql & "where HINBAN = '" & HIN.hinban & "' "
                sql = sql & "and MNOREVNO = " & HIN.mnorevno & " "
                sql = sql & "and FACTORY = '" & HIN.factory & "' "
                sql = sql & "and OPECOND = '" & HIN.opecond & "' "
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    sErrMsg = GetMsgStr("ECLC3") & " <直径> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                tCsSuitei.sUpDm = rs("HSXD1CEN")            '品SX直径1中心
                
                '⑦ｶｰﾎﾞﾝ偏析係数
                sql = "select CTR01A9 from KODA9 "
                sql = sql & "where SYSCA9 = 'K' "
                sql = sql & "and SHUCA9 = 'AP' "
                sql = sql & "and CODEA9 = '1' "
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    sErrMsg = GetMsgStr("ECLC3") & " <ｶｰﾎﾞﾝ偏析係数> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                tCsSuitei.sCsHenseki = rs("CTR01A9")        'ｶｰﾎﾞﾝ偏析係数
                
                '⑧Csﾌﾞﾛｯｸ推定値計算
                If Not GetCsSuiteiMain(tCsSuitei, dCsSuitei) Then
                    sErrMsg = GetMsgStr("ECLC3")
                    GoTo proc_exit
                End If
                .Fields("SXLCS_BSUIMEAS").Value = dCsSuitei
            End If
            ''Csﾌﾞﾛｯｸ推定値計算対応　06/04/20 ooba END ========================================>
        End If
    End With

    getTBCMJ004 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ004 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :結晶OSF実績(TBCMJ005)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS構造体   (新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ))
'          :j               , I  ,Integer           , OSF No
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002構造体(測定点ﾃﾞｰﾀ)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :結晶OSF実績(TBCMJ005)からﾃﾞｰﾀを取得し、SXL検査書・測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
Private Function getTBCMJ005(CRYNUM As String, recXSDCS As c_cmzcrec, j As Integer, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ005"
    
    getTBCMJ005 = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    'TBCMX001
    With recX001
        If j = 1 Then
            .Fields("SXLOSF_SMPPOS").Value = -1             'OSFサンプル測定位置(SXL測定情報)
        End If
        .Fields("SXLOSF" & j & "_KKSP").Value = ""          'OSFx結晶欠陥測定位置
        .Fields("SXLOSF" & j & "_NETU").Value = ""          'OSFx熱処理法
        .Fields("SXLOSF" & j & "_KKSET").Value = ""         'OSFx結晶欠陥測定条件+選択ET代
        .Fields("SXLOSF" & j & "_CALCMAX").Value = -1       'OSFxSXL計算結果 Max_x
        .Fields("SXLOSF" & j & "_CALCAVE").Value = -1       'OSFxSXL計算結果 Ave_x
    End With
        
    'TBCMX002
    With recX002
        If j = 1 Then
            .Fields("SXLOSF1_SMPPOS").Value = -1            'SXLOSFサンプル測定位置(SXL位置情報)
        End If
        .Fields("SXLOSF" & j & "_KKSP").Value = ""          'SXLOSFx結晶欠陥確定位置
        .Fields("SXLOSF" & j & "_NETU").Value = ""          'SXLOSFx熱処理法
        .Fields("SXLOSF" & j & "_KKSET").Value = ""         'SXLOSFx結晶欠陥測定条件+選択ET代
        .Fields("SXLOSF" & j & "_MEAS1").Value = -1         'SXLOSFx測定点1
        .Fields("SXLOSF" & j & "_MEAS2").Value = -1         'SXLOSFx測定点2
        .Fields("SXLOSF" & j & "_MEAS3").Value = -1         'SXLOSFx測定点3
        .Fields("SXLOSF" & j & "_MEAS4").Value = -1         'SXLOSFx測定点4
        .Fields("SXLOSF" & j & "_MEAS5").Value = -1         'SXLOSFx測定点5
        .Fields("SXLOSF" & j & "_MEAS6").Value = -1         'SXLOSFx測定点6
        .Fields("SXLOSF" & j & "_MEAS7").Value = -1         'SXLOSFx測定点7
        .Fields("SXLOSF" & j & "_MEAS8").Value = -1         'SXLOSFx測定点8
        .Fields("SXLOSF" & j & "_MEAS9").Value = -1         'SXLOSFx測定点9
        .Fields("SXLOSF" & j & "_MEAS10").Value = -1        'SXLOSFx測定点10
        .Fields("SXLOSF" & j & "_MEAS11").Value = -1        'SXLOSFx測定点11
        .Fields("SXLOSF" & j & "_MEAS12").Value = -1        'SXLOSFx測定点12
        .Fields("SXLOSF" & j & "_MEAS13").Value = -1        'SXLOSFx測定点13
        .Fields("SXLOSF" & j & "_MEAS14").Value = -1        'SXLOSFx測定点14
        .Fields("SXLOSF" & j & "_MEAS15").Value = -1        'SXLOSFx測定点15
        .Fields("SXLOSF" & j & "_MEAS16").Value = -1        'SXLOSFx測定点16
        .Fields("SXLOSF" & j & "_MEAS17").Value = -1        'SXLOSFx測定点17
        .Fields("SXLOSF" & j & "_MEAS18").Value = -1        'SXLOSFx測定点18
        .Fields("SXLOSF" & j & "_MEAS19").Value = -1        'SXLOSFx測定点19
        .Fields("SXLOSF" & j & "_MEAS20").Value = -1        'SXLOSFx測定点20
    End With
    
    '-------------------- TBCMJ005の読み込み(OSF1～4) ----------------------------------------
    If (recXSDCS("CRYINDL" & j & "CS").Value <> "0") And (recXSDCS("CRYRESL" & j & "CS").Value <> "0") Then
        sql = "select * from TBCMJ005 "
        sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
        sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDL" & j & "CS").Value & " and "
        sql = sql & "      TRANCOND = '" & j & "' "
        sql = sql & "order by TRANCNT desc"
        sql = "select * from (" & sql & ") where rownum = 1"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
    
        'TBCMX001
        With recX001
            If .Fields("SXLOSF_SMPPOS").Value = -1 Then
                .Fields("SXLOSF_SMPPOS").Value = rs("POSITION")         'OSFサンプル測定位置(SXL測定情報)
            End If
            .Fields("SXLOSF" & j & "_KKSP").Value = rs("KKSP")          'OSFx結晶欠陥測定位置
            .Fields("SXLOSF" & j & "_NETU").Value = rs("HTPRC")         'OSFx熱処理法
            .Fields("SXLOSF" & j & "_KKSET").Value = rs("KKSET")        'OSFx結晶欠陥測定条件+選択ET代
            .Fields("SXLOSF" & j & "_CALCMAX").Value = rs("CALCMAX")    'OSFxSXL計算結果 Max_x
            .Fields("SXLOSF" & j & "_CALCAVE").Value = rs("CALCAVE")    'OSFxSXL計算結果 Ave_x
        End With
            
        'TBCMX002
        With recX002
            If .Fields("SXLOSF1_SMPPOS").Value = -1 Then
                .Fields("SXLOSF1_SMPPOS").Value = rs("POSITION")        'SXLOSFサンプル測定位置(SXL位置情報)
            End If
            .Fields("SXLOSF" & j & "_KKSP").Value = rs("KKSP")          'SXLOSFx結晶欠陥確定位置
            .Fields("SXLOSF" & j & "_NETU").Value = rs("HTPRC")         'SXLOSFx熱処理法
            .Fields("SXLOSF" & j & "_KKSET").Value = rs("KKSET")        'SXLOSFx結晶欠陥測定条件+選択ET代
            .Fields("SXLOSF" & j & "_MEAS1").Value = rs("MEAS1")        'SXLOSFx測定点1
            .Fields("SXLOSF" & j & "_MEAS2").Value = rs("MEAS2")        'SXLOSFx測定点2
            .Fields("SXLOSF" & j & "_MEAS3").Value = rs("MEAS3")        'SXLOSFx測定点3
            .Fields("SXLOSF" & j & "_MEAS4").Value = rs("MEAS4")        'SXLOSFx測定点4
            .Fields("SXLOSF" & j & "_MEAS5").Value = rs("MEAS5")        'SXLOSFx測定点5
            .Fields("SXLOSF" & j & "_MEAS6").Value = rs("MEAS6")        'SXLOSFx測定点6
            .Fields("SXLOSF" & j & "_MEAS7").Value = rs("MEAS7")        'SXLOSFx測定点7
            .Fields("SXLOSF" & j & "_MEAS8").Value = rs("MEAS8")        'SXLOSFx測定点8
            .Fields("SXLOSF" & j & "_MEAS9").Value = rs("MEAS9")        'SXLOSFx測定点9
            .Fields("SXLOSF" & j & "_MEAS10").Value = rs("MEAS10")      'SXLOSFx測定点10
            .Fields("SXLOSF" & j & "_MEAS11").Value = rs("MEAS11")      'SXLOSFx測定点11
            .Fields("SXLOSF" & j & "_MEAS12").Value = rs("MEAS12")      'SXLOSFx測定点12
            .Fields("SXLOSF" & j & "_MEAS13").Value = rs("MEAS13")      'SXLOSFx測定点13
            .Fields("SXLOSF" & j & "_MEAS14").Value = rs("MEAS14")      'SXLOSFx測定点14
            .Fields("SXLOSF" & j & "_MEAS15").Value = rs("MEAS15")      'SXLOSFx測定点15
            .Fields("SXLOSF" & j & "_MEAS16").Value = rs("MEAS16")      'SXLOSFx測定点16
            .Fields("SXLOSF" & j & "_MEAS17").Value = rs("MEAS17")      'SXLOSFx測定点17
            .Fields("SXLOSF" & j & "_MEAS18").Value = rs("MEAS18")      'SXLOSFx測定点18
            .Fields("SXLOSF" & j & "_MEAS19").Value = rs("MEAS19")      'SXLOSFx測定点19
            .Fields("SXLOSF" & j & "_MEAS20").Value = rs("MEAS20")      'SXLOSFx測定点20
        End With
        Set rs = Nothing
    End If

    getTBCMJ005 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ005 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :結晶BMD実績(TBCMJ008)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS構造体   (新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ))
'          :j               , I  ,Integer           , BMD No
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002構造体(測定点ﾃﾞｰﾀ)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :結晶BMD実績(TBCMJ008)からﾃﾞｰﾀを取得し、SXL検査書・測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :
Private Function getTBCMJ008(CRYNUM As String, recXSDCS As c_cmzcrec, j As Integer, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim dMeas(9)    As Double
    Dim strMeasPos  As String
    Dim iRet        As Integer
    Dim wComp       As Double
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ008"
    
    getTBCMJ008 = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    'TBCMX001
    With recX001
        If j = 1 Then
            .Fields("SXLBMD_SMPPOS").Value = -1             'BMDサンプル測定位置(SXL位置情報)
        End If
        .Fields("SXLBMD" & j & "_KKSP").Value = ""          'BMDx結晶欠陥測定位置
        .Fields("SXLBMD" & j & "_NETU").Value = ""          'BMDx熱処理法
        .Fields("SXLBMD" & j & "_KKSET").Value = ""         'BMDx結晶欠陥測定条件＋選択ET代
        .Fields("SXLBMD" & j & "_CALCMAX").Value = -1       'BMDxSXL計算結果 Max
        .Fields("SXLBMD" & j & "_CALCAVE").Value = -1       'BMDxSXL計算結果 Ave
        .Fields("SXLBMD" & j & "_CALCMIN").Value = -1       'BMDxSXL計算結果 Min
        .Fields("SXLBMD" & j & "_CALCMB").Value = -1        'BMDxSXL計算結果 面内分布
    End With
        
    'TBCMX002
    With recX002
        If j = 1 Then
            .Fields("SXLBMD_SMPPOS").Value = -1             'SXLBMDサンプル測定位置(SXL位置情報)
        End If
        .Fields("SXLBMD" & j & "_KKSP").Value = ""          'SXLBMD1結晶欠陥測定位置
        .Fields("SXLBMD" & j & "_NETU").Value = ""          'SXLBMD1熱処理法
        .Fields("SXLBMD" & j & "_KKSET").Value = ""         'SXLBMD1結晶欠陥測定条件+選択ET代
        .Fields("SXLBMD" & j & "_MEAS1").Value = -1         'SXLBMD1測定点1
        .Fields("SXLBMD" & j & "_MEAS2").Value = -1         'SXLBMD1測定点2
        .Fields("SXLBMD" & j & "_MEAS3").Value = -1         'SXLBMD1測定点3
        .Fields("SXLBMD" & j & "_MEAS4").Value = -1         'SXLBMD1測定点4
        .Fields("SXLBMD" & j & "_MEAS5").Value = -1         'SXLBMD1測定点5
    End With
    
    '-------------------- TBCMJ008の読み込み(BMD1～3) ----------------------------------------
    If (recXSDCS("CRYINDB" & j & "CS").Value <> "0") And (recXSDCS("CRYRESB" & j & "CS").Value <> "0") Then
        sql = "select * from TBCMJ008 "
        sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
        sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDB" & j & "CS").Value & " and "
        sql = sql & "      TRANCOND = '" & j & "' "
        sql = sql & "order by TRANCNT desc"
        sql = "select * from (" & sql & ") where rownum = 1"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            If .Fields("SXLBMD_SMPPOS").Value = -1 Then
                .Fields("SXLBMD_SMPPOS").Value = rs("POSITION")         'BMDサンプル測定位置(SXL位置情報)
            End If
            .Fields("SXLBMD" & j & "_KKSP").Value = rs("KKSP")          'BMDx結晶欠陥測定位置
            .Fields("SXLBMD" & j & "_NETU").Value = rs("HTPRC")         'BMDx熱処理法
            .Fields("SXLBMD" & j & "_KKSET").Value = rs("KKSET")        'BMDx結晶欠陥測定条件＋選択ET代
            .Fields("SXLBMD" & j & "_CALCMAX").Value = rs("MEASMAX")    'BMDxSXL計算結果 Max
            .Fields("SXLBMD" & j & "_CALCAVE").Value = rs("MEASAVE")    'BMDxSXL計算結果 Ave
'            .Fields("SXLBMD" & j & "_CALCMB").Value = rs("BMDMNBUNP")   'BMDxSXL計算結果 面内分布
            If IsNull(rs("BMDMNBUNP")) = False Then .Fields("SXLBMD" & j & "_CALCMB").Value = rs("BMDMNBUNP")   'BMDxSXL計算結果 面内分布
        End With
            
        'TBCMX002
        With recX002
            If .Fields("SXLBMD_SMPPOS").Value = -1 Then
                .Fields("SXLBMD_SMPPOS").Value = rs("POSITION")         'SXLBMDサンプル測定位置(SXL位置情報)
            End If
            .Fields("SXLBMD" & j & "_KKSP").Value = rs("KKSP")          'SXLBMDx結晶欠陥測定位置
            .Fields("SXLBMD" & j & "_NETU").Value = rs("HTPRC")         'SXLBMDx熱処理法
            .Fields("SXLBMD" & j & "_KKSET").Value = rs("KKSET")        'SXLBMDx結晶欠陥測定条件+選択ET代
            .Fields("SXLBMD" & j & "_MEAS1").Value = rs("MEAS1")        'SXLBMDx測定点1
            .Fields("SXLBMD" & j & "_MEAS2").Value = rs("MEAS2")        'SXLBMDx測定点2
            .Fields("SXLBMD" & j & "_MEAS3").Value = rs("MEAS3")        'SXLBMDx測定点3
            .Fields("SXLBMD" & j & "_MEAS4").Value = rs("MEAS4")        'SXLBMDx測定点4
            .Fields("SXLBMD" & j & "_MEAS5").Value = rs("MEAS5")        'SXLBMDx測定点5
        End With
        Set rs = Nothing
    
        'BMD最小値の取得 2003/05/31 tuku                START
        dMeas(0) = recX002.Fields("SXLBMD" & j & "_MEAS1").Value
        dMeas(1) = recX002.Fields("SXLBMD" & j & "_MEAS2").Value
        dMeas(2) = recX002.Fields("SXLBMD" & j & "_MEAS3").Value
        dMeas(3) = recX002.Fields("SXLBMD" & j & "_MEAS4").Value
        dMeas(4) = recX002.Fields("SXLBMD" & j & "_MEAS5").Value
        ''結晶欠陥測定位置コード
        strMeasPos = Trim(recX002.Fields("SXLBMD" & j & "_KKSP").Value)
        ''最小値を計算する。
        iRet = getSXLBMDMIN(wComp, strMeasPos, dMeas)
        ''計算結果を格納する
        recX001.Fields("SXLBMD" & j & "_CALCMIN").Value = wComp
    End If

    getTBCMJ008 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ008 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :GD実績(TBCMJ006)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS構造体   (新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ))
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002構造体(測定点ﾃﾞｰﾀ)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :GD実績(TBCMJ006)からﾃﾞｰﾀを取得し、SXL検査書・測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :
Private Function getTBCMJ006(CRYNUM As String, recXSDCS As c_cmzcrec, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ006"
    
    getTBCMJ006 = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("SXLGD_SMPPOS").Value = -1                  'GDサンプル測定位置(SXL位置情報)
        .Fields("SXLGD_MSRSDEN").Value = -1                 'SXLGD_測定結果 Den
        .Fields("SXLGD_MSRSLDL").Value = -1                 'SXLGD_測定結果 L/DL
        .Fields("SXLGD_MSRSDVD2").Value = -1                'SXLGD_測定結果 DVD2
    End With
        
    'TBCMX002
    With recX002
        .Fields("SXLGD_SMPPOS").Value = -1                                  'SXLGDサンプル測定位置(SXL位置情報)
        For i = 1 To 15
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL1").Value = -1       'SXLGD_測定値xx L/DL1
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL2").Value = -1       'SXLGD_測定値xx L/DL2
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL3").Value = -1       'SXLGD_測定値xx L/DL3
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL4").Value = -1       'SXLGD_測定値xx L/DL4
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL5").Value = -1       'SXLGD_測定値xx L/DL5
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN1").Value = -1       'SXLGD_測定値xx Den1
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN2").Value = -1       'SXLGD_測定値xx Den2
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN3").Value = -1       'SXLGD_測定値xx Den3
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN4").Value = -1       'SXLGD_測定値xx Den4
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN5").Value = -1       'SXLGD_測定値xx Den5
        Next
        For i = 1 To 5
            .Fields("SXLGD_MS01DVD2" & i).Value = -1                        'SXLGD_測定値xx DVD2
        Next
    End With
        
    '-------------------- TBCMJ006の読み込み(GD) ----------------------------------------
    If (recXSDCS("CRYINDGDCS").Value <> "0") And (recXSDCS("CRYRESGDCS").Value <> "0") Then
        sql = "select * from TBCMJ006 "
        sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
        sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDGDCS").Value & " "
        sql = sql & "order by TRANCNT desc"
        sql = "select * from (" & sql & ") where rownum = 1"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            .Fields("SXLGD_SMPPOS").Value = rs("POSITION")              'GDサンプル測定位置(SXL位置情報)
            .Fields("SXLGD_MSRSDEN").Value = rs("MSRSDEN")              'SXLGD_測定結果 Den
            .Fields("SXLGD_MSRSLDL").Value = rs("MSRSLDL")              'SXLGD_測定結果 L/DL
            .Fields("SXLGD_MSRSDVD2").Value = rs("MSRSDVD2")            'SXLGD_測定結果 DVD2
        End With
            
        'TBCMX002
        With recX002
            .Fields("SXLGD_SMPPOS").Value = rs("POSITION")                                                      'SXLGDサンプル測定位置(SXL位置情報)
            For i = 1 To 15
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL1").Value = rs("MS" & Format(i, "00") & "LDL1")      'SXLGD_測定値xx L/DL1
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL2").Value = rs("MS" & Format(i, "00") & "LDL2")      'SXLGD_測定値xx L/DL2
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL3").Value = rs("MS" & Format(i, "00") & "LDL3")      'SXLGD_測定値xx L/DL3
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL4").Value = rs("MS" & Format(i, "00") & "LDL4")      'SXLGD_測定値xx L/DL4
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL5").Value = rs("MS" & Format(i, "00") & "LDL5")      'SXLGD_測定値xx L/DL5
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN1").Value = rs("MS" & Format(i, "00") & "DEN1")      'SXLGD_測定値xx Den1
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN2").Value = rs("MS" & Format(i, "00") & "DEN2")      'SXLGD_測定値xx Den2
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN3").Value = rs("MS" & Format(i, "00") & "DEN3")      'SXLGD_測定値xx Den3
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN4").Value = rs("MS" & Format(i, "00") & "DEN4")      'SXLGD_測定値xx Den4
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN5").Value = rs("MS" & Format(i, "00") & "DEN5")      'SXLGD_測定値xx Den5
            Next
            
'--------------- 208/06/24 INSERT START  By Systech ---------------
            For i = 1 To 5
                If rs("MS0" & i & "DVD2") <> -1 Then
                    .Fields("SXLGD_MS01DVD2" & i).Value = rs("MS0" & i & "DVD2")                                'SXLGD_測定値xx DVD2
                End If
            Next
'--------------- 208/06/24 INSERT  END   By Systech ---------------
        
        End With
        Set rs = Nothing
    End If

    getTBCMJ006 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ006 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :LT実績(TBCMJ007)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS構造体   (新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ))
'          :ChkHin          , I  ,tFullHinban       , LT仕様取得用品番　05/12/05 ooba
'          :i               , I  ,Integer           , BMD No
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002構造体(測定点ﾃﾞｰﾀ)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :LT実績(TBCMJ007)からﾃﾞｰﾀを取得し、SXL検査書・測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :
Private Function getTBCMJ007(CRYNUM As String, recXSDCS As c_cmzcrec, ChkHin As tFullHinban, i As Integer, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim j           As Integer      '
    Dim rs2         As OraDynaset   '
    Dim sql2        As String       '
    Dim iRet        As Integer      '
    Dim iTmpMes(9)  As Integer      'LT実績ﾃﾞｰﾀ(1～10)
    Dim iCalcMeas   As Integer      'LT計算結果
    Dim sIchi       As String       '品SXLﾀｲﾑ測定位置_位
    Dim iOldFlg     As Integer      '旧ﾃﾞｰﾀﾌﾗｸﾞ
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ007"
    
    getTBCMJ007 = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("SXLLT_SMPPOS").Value = -1                  'LTサンプル測定位置(SXL位置情報)
        .Fields("SXLLT_MEASPEAK").Value = -1                'SXLLT_測定値 ピーク値
        .Fields("SXLLT_CALCMEAS").Value = -1                'SXLLT_計算結果
    End With
        
    'TBCMX002
    With recX002
        .Fields("SXLT_SMPPOS").Value = -1                   'SXLLTサンプル測定位置(SXL位置情報)
        .Fields("SXLLT_MEASPEAK").Value = -1                'SXLLT_測定値 ピーク値
        .Fields("SXLLT_MEAS1").Value = -1                   'SXLLT_測定値1
        .Fields("SXLLT_MEAS2").Value = -1                   'SXLLT_測定値2
        .Fields("SXLLT_MEAS3").Value = -1                   'SXLLT_測定値3
        .Fields("SXLLT_MEAS4").Value = -1                   'SXLLT_測定値4
        .Fields("SXLLT_MEAS5").Value = -1                   'SXLLT_測定値5
    End With
        
    'BOT側のみﾃﾞｰﾀ取得
    If i <> 1 Then
        '-------------------- TBCMJ007の読み込み(LT) ----------------------------------------
        If (recXSDCS("CRYINDTCS").Value <> "0") And (recXSDCS("CRYRESTCS").Value <> "0") Then
            sql = "select * from TBCMJ007 "
            sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
            sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDTCS").Value & " "
            sql = sql & "order by TRANCNT desc"
            sql = "select * from (" & sql & ") where rownum = 1"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            
            If IsNull(rs("LTSPIFLG")) Then iOldFlg = 1 Else iOldFlg = 0
            
            '初期化
            iCalcMeas = -1
            For j = 0 To 9
                iTmpMes(j) = -1
            Next j
            
            If Not IsNull(rs("MEAS1")) Then iTmpMes(0) = rs("MEAS1")
            If Not IsNull(rs("MEAS2")) Then iTmpMes(1) = rs("MEAS2")
            If Not IsNull(rs("MEAS3")) Then iTmpMes(2) = rs("MEAS3")
            If Not IsNull(rs("MEAS4")) Then iTmpMes(3) = rs("MEAS4")
            If Not IsNull(rs("MEAS5")) Then iTmpMes(4) = rs("MEAS5")
            If Not IsNull(rs("MEAS6")) Then iTmpMes(5) = rs("MEAS6")
            If Not IsNull(rs("MEAS7")) Then iTmpMes(6) = rs("MEAS7")
            If Not IsNull(rs("MEAS8")) Then iTmpMes(7) = rs("MEAS8")
            If Not IsNull(rs("MEAS9")) Then iTmpMes(8) = rs("MEAS9")
            If Not IsNull(rs("MEAS10")) Then iTmpMes(9) = rs("MEAS10")
            
            '10点測定の場合
            If iOldFlg = 0 Then
                sql2 = "select HSXLTSPI from TBCME019"
                sql2 = sql2 & " where HINBAN = '" & ChkHin.hinban & "'"
                sql2 = sql2 & " and MNOREVNO = " & ChkHin.mnorevno
                sql2 = sql2 & " and FACTORY = '" & ChkHin.factory & "'"
                sql2 = sql2 & " and OPECOND = '" & ChkHin.opecond & "'"
                Set rs2 = OraDB.DBCreateDynaset(sql2, ORADYN_NO_BLANKSTRIP)
                If rs2.RecordCount = 0 Then
                    Set rs2 = Nothing
                    GoTo proc_exit
                End If
                If Not IsNull(rs2("HSXLTSPI")) Then sIchi = rs2("HSXLTSPI") Else sIchi = ""
                Set rs2 = Nothing
            End If
            
            '計算結果取得
            iRet = KNS_CalculateMeasResult_LT(iCalcMeas, iTmpMes(), sIchi, iOldFlg)

            
            'TBCMX001
            With recX001
                .Fields("SXLLT_SMPPOS").Value = rs("POSITION")          'LTサンプル測定位置(SXL位置情報)
                .Fields("SXLLT_MEASPEAK").Value = rs("MEASPEAK")        'SXLLT_測定値 ピーク値

                .Fields("SXLLT_CALCMEAS").Value = iCalcMeas             'SXLLT_計算結果
            End With
                
            'TBCMX002
            With recX002
                .Fields("SXLT_SMPPOS").Value = rs("POSITION")           'SXLLTサンプル測定位置(SXL位置情報)
                .Fields("SXLLT_MEASPEAK").Value = rs("MEASPEAK")        'SXLLT_測定値 ピーク値

                '旧ﾃﾞｰﾀ
                If iOldFlg = 1 Then
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_測定値1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(1)           'SXLLT_測定値2
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(2)           'SXLLT_測定値3
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(3)           'SXLLT_測定値4
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(4)           'SXLLT_測定値5
                '3:CE,Inside3mm
                ElseIf sIchi = "3" Then
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_測定値1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(0)           'SXLLT_測定値1
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(7)           'SXLLT_測定値8
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(8)           'SXLLT_測定値9
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(9)           'SXLLT_測定値10
                '5:CE,Inside5mm
                ElseIf sIchi = "5" Then
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_測定値1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(0)           'SXLLT_測定値1
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(4)           'SXLLT_測定値5
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(5)           'SXLLT_測定値6
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(6)           'SXLLT_測定値7
                'A:CE,Inside10mm
                ElseIf sIchi = "A" Then
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_測定値1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(0)           'SXLLT_測定値1
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(1)           'SXLLT_測定値2
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(2)           'SXLLT_測定値3
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(3)           'SXLLT_測定値4
                'その他
                Else
                    'その他の場合は｢A:CE,Inside10mm｣とする
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_測定値1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(0)           'SXLLT_測定値1
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(1)           'SXLLT_測定値2
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(2)           'SXLLT_測定値3
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(3)           'SXLLT_測定値4
                End If
            End With
            Set rs = Nothing
        End If
    End If

    getTBCMJ007 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ007 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function



'概要      :結晶GD実績(TBCMJ006)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS構造体   (新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ))
'          :recX003         , O  ,c_cmzcrec         , TBCMX003構造体(GD検査測定点ﾃﾞｰﾀ)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :結晶GD実績(TBCMJ006)からﾃﾞｰﾀを取得し、GD検査測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'          :結晶GD実績(TBCMJ006)の測定データの初期値である-1をNULLに変更してTBCMX003に登録する。
'履歴      :2005/02/15 ffc)tanabe
Private Function getTBCMJ006GD(CRYNUM As String, recXSDCW As c_cmzcrec, recX003 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ006GD"
    
    getTBCMJ006GD = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
        
    'TBCMX003
    With recX003
            .Fields("SXLGD_HSFLG").Value = vbNullString                              'SXLGDGD測定結果保証フラグ
            .Fields("SXLGD_SMPPOS").Value = vbNullString                             'SXLGDGDサンプル測定位置(SXL位置情報)
            .Fields("SXLGD_MSRSDEN").Value = vbNullString                            'SXLGDGD_測定結果 Den
            .Fields("SXLGD_MSRSLDL").Value = vbNullString                            'SXLGDGD_測定結果 L/DL
            .Fields("SXLGD_MSRSDVD2").Value = vbNullString                           'SXLGDGD_測定結果 DVD2
            .Fields("WFGD_HSFLG").Value = vbNullString                               'WFGD測定結果保証フラグ
            .Fields("WFGD_SMPPOS").Value = vbNullString                              'WFGDサンプル測定位置(SXL位置情報)
            .Fields("WFGD_MSRSDEN").Value = vbNullString                             'WFGD_測定結果 Den
            .Fields("WFGD_MSRSLDL").Value = vbNullString                             'WFGD_測定結果 L/DL
            .Fields("WFGD_MSRSDVD2").Value = vbNullString                            'WFGD_測定結果 DVD2
            
        For i = 1 To 15
            .Fields("WFGD_MS" & Format(i, "00") & "LDL1").Value = vbNullString       'WFGD_測定値xx L/DL1
            .Fields("WFGD_MS" & Format(i, "00") & "LDL2").Value = vbNullString       'WFGD_測定値xx L/DL2
            .Fields("WFGD_MS" & Format(i, "00") & "LDL3").Value = vbNullString       'WFGD_測定値xx L/DL3
            .Fields("WFGD_MS" & Format(i, "00") & "LDL4").Value = vbNullString       'WFGD_測定値xx L/DL4
            .Fields("WFGD_MS" & Format(i, "00") & "LDL5").Value = vbNullString       'WFGD_測定値xx L/DL5
            .Fields("WFGD_MS" & Format(i, "00") & "DEN1").Value = vbNullString       'WFGD_測定値xx Den1
            .Fields("WFGD_MS" & Format(i, "00") & "DEN2").Value = vbNullString       'WFGD_測定値xx Den2
            .Fields("WFGD_MS" & Format(i, "00") & "DEN3").Value = vbNullString       'WFGD_測定値xx Den3
            .Fields("WFGD_MS" & Format(i, "00") & "DEN4").Value = vbNullString       'WFGD_測定値xx Den4
            .Fields("WFGD_MS" & Format(i, "00") & "DEN5").Value = vbNullString       'WFGD_測定値xx Den5
        Next
        
        For i = 1 To 5
            .Fields("WFGD_MS01DVD2" & i).Value = vbNullString                        'WFGD_測定値xx DVD2
        Next
        
    End With
        
    '-------------------- TBCMJ006の読み込み(GD) ----------------------------------------
    sql = "select * from TBCMJ006 "
    sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
    sql = sql & "      SMPLNO = " & Trim(recXSDCW("WFSMPLIDGDCW").Value)
    sql = sql & " order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    
    'TBCMX003
    With recX003
        .Fields("SXLGD_HSFLG").Value = "1"                          'SXLGD測定結果保証フラグ
        .Fields("SXLGD_SMPPOS").Value = rs("POSITION")              'SXLGDサンプル測定位置(SXL位置情報)
        If rs("MSRSDEN") <> -1 Then
            .Fields("SXLGD_MSRSDEN").Value = rs("MSRSDEN")          'SXLGD_測定結果 Den
        End If
        If rs("MSRSLDL") <> -1 Then
            .Fields("SXLGD_MSRSLDL").Value = rs("MSRSLDL")          'SXLGD_測定結果 L/DL
        End If
        If rs("MSRSDVD2") <> -1 Then
            .Fields("SXLGD_MSRSDVD2").Value = rs("MSRSDVD2")        'SXLGD_測定結果 DVD2
        End If
        
        For i = 1 To 15
            If rs("MS" & Format(i, "00") & "DEN1") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN1").Value = rs("MS" & Format(i, "00") & "DEN1")      'SXLGD_測定値xx Den1
            End If
            If rs("MS" & Format(i, "00") & "DEN2") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN2").Value = rs("MS" & Format(i, "00") & "DEN2")      'SXLGD_測定値xx Den2
            End If
            If rs("MS" & Format(i, "00") & "DEN3") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN3").Value = rs("MS" & Format(i, "00") & "DEN3")      'SXLGD_測定値xx Den3
            End If
            If rs("MS" & Format(i, "00") & "DEN4") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN4").Value = rs("MS" & Format(i, "00") & "DEN4")      'SXLGD_測定値xx Den4
            End If
            If rs("MS" & Format(i, "00") & "DEN5") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN5").Value = rs("MS" & Format(i, "00") & "DEN5")      'SXLGD_測定値xx Den5
            End If
            If rs("MS" & Format(i, "00") & "LDL1") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL1").Value = rs("MS" & Format(i, "00") & "LDL1")      'SXLGD_測定値xx L/DL1
            End If
            If rs("MS" & Format(i, "00") & "LDL2") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL2").Value = rs("MS" & Format(i, "00") & "LDL2")      'SXLGD_測定値xx L/DL2
            End If
            If rs("MS" & Format(i, "00") & "LDL3") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL3").Value = rs("MS" & Format(i, "00") & "LDL3")      'SXLGD_測定値xx L/DL3
            End If
            If rs("MS" & Format(i, "00") & "LDL4") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL4").Value = rs("MS" & Format(i, "00") & "LDL4")      'SXLGD_測定値xx L/DL4
            End If
            If rs("MS" & Format(i, "00") & "LDL5") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL5").Value = rs("MS" & Format(i, "00") & "LDL5")      'SXLGD_測定値xx L/DL5
            End If
        Next
        
        For i = 1 To 5
            If rs("MS0" & i & "DVD2") <> -1 Then
                .Fields("WFGD_MS01DVD2" & i).Value = rs("MS0" & i & "DVD2")         'SXLGD_測定値xx DVD2
            End If
        Next
        
    End With
    Set rs = Nothing

    getTBCMJ006GD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ006GD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :WFGD実績(TBCMJ015)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW構造体   (新ｻﾝﾌﾟﾙ管理(SXL))
'          :recX003         , O  ,c_cmzcrec         , TBCMX003構造体(GD検査測定点ﾃﾞｰﾀ)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :WFGD実績(TBCMJ015)からﾃﾞｰﾀを取得し、GD検査測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :2005/02/15 ffc)tanabe
Private Function getTBCMJ015WFGD(CRYNUM As String, recXSDCW As c_cmzcrec, recX003 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ015WFGD"
    
    getTBCMJ015WFGD = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
        
    'TBCMX003
    With recX003
            .Fields("SXLGD_HSFLG").Value = vbNullString                              'SXLGDGD測定結果保証フラグ
            .Fields("SXLGD_SMPPOS").Value = vbNullString                             'SXLGDGDサンプル測定位置(SXL位置情報)
            .Fields("SXLGD_MSRSDEN").Value = vbNullString                            'SXLGDGD_測定結果 Den
            .Fields("SXLGD_MSRSLDL").Value = vbNullString                            'SXLGDGD_測定結果 L/DL
            .Fields("SXLGD_MSRSDVD2").Value = vbNullString                           'SXLGDGD_測定結果 DVD2
            .Fields("WFGD_HSFLG").Value = vbNullString                               'WFGD測定結果保証フラグ
            .Fields("WFGD_SMPPOS").Value = vbNullString                              'WFGDサンプル測定位置(SXL位置情報)
            .Fields("WFGD_MSRSDEN").Value = vbNullString                             'WFGD_測定結果 Den
            .Fields("WFGD_MSRSLDL").Value = vbNullString                             'WFGD_測定結果 L/DL
            .Fields("WFGD_MSRSDVD2").Value = vbNullString                            'WFGD_測定結果 DVD2
            
        For i = 1 To 15
            .Fields("WFGD_MS" & Format(i, "00") & "LDL1").Value = vbNullString       'WFGD_測定値xx L/DL1
            .Fields("WFGD_MS" & Format(i, "00") & "LDL2").Value = vbNullString       'WFGD_測定値xx L/DL2
            .Fields("WFGD_MS" & Format(i, "00") & "LDL3").Value = vbNullString       'WFGD_測定値xx L/DL3
            .Fields("WFGD_MS" & Format(i, "00") & "LDL4").Value = vbNullString       'WFGD_測定値xx L/DL4
            .Fields("WFGD_MS" & Format(i, "00") & "LDL5").Value = vbNullString       'WFGD_測定値xx L/DL5
            .Fields("WFGD_MS" & Format(i, "00") & "DEN1").Value = vbNullString       'WFGD_測定値xx Den1
            .Fields("WFGD_MS" & Format(i, "00") & "DEN2").Value = vbNullString       'WFGD_測定値xx Den2
            .Fields("WFGD_MS" & Format(i, "00") & "DEN3").Value = vbNullString       'WFGD_測定値xx Den3
            .Fields("WFGD_MS" & Format(i, "00") & "DEN4").Value = vbNullString       'WFGD_測定値xx Den4
            .Fields("WFGD_MS" & Format(i, "00") & "DEN5").Value = vbNullString       'WFGD_測定値xx Den5
        Next
        
        For i = 1 To 5
            .Fields("WFGD_MS01DVD2" & i).Value = vbNullString                        'WFGD_測定値xx DVD2
        Next
        
    End With
        
    '-------------------- TBCMJ015の読み込み(GD) ----------------------------------------
    sql = "select * from TBCMJ015 "
    sql = sql & " where CRYNUM = '" & CRYNUM & "'"
    sql = sql & " and   SMPLNO = '" & recXSDCW("WFSMPLIDGDCW").Value & "'"
    sql = sql & " and   HSFLG = '1'"
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    
    'TBCMX003
    With recX003
        .Fields("WFGD_HSFLG").Value = "1"                                                                 'WFGD測定結果保証フラグ
        .Fields("WFGD_SMPPOS").Value = rs("POSITION")                                                     'WFGDサンプル測定位置(SXL位置情報)
        .Fields("WFGD_MSRSDEN").Value = rs("MSRSDEN")                                                     'WFGD_測定結果 Den
        .Fields("WFGD_MSRSLDL").Value = rs("MSRSLDL")                                                     'WFGD_測定結果 L/DL
        .Fields("WFGD_MSRSDVD2").Value = rs("MSRSDVD2")                                                   'WFGD_測定結果 DVD2
        
        For i = 1 To 15
            .Fields("WFGD_MS" & Format(i, "00") & "DEN1").Value = rs("MS" & Format(i, "00") & "DEN1")      'WFGD_測定値xx Den1
            .Fields("WFGD_MS" & Format(i, "00") & "DEN2").Value = rs("MS" & Format(i, "00") & "DEN2")      'WFGD_測定値xx Den2
            .Fields("WFGD_MS" & Format(i, "00") & "DEN3").Value = rs("MS" & Format(i, "00") & "DEN3")      'WFGD_測定値xx Den3
            .Fields("WFGD_MS" & Format(i, "00") & "DEN4").Value = rs("MS" & Format(i, "00") & "DEN4")      'WFGD_測定値xx Den4
            .Fields("WFGD_MS" & Format(i, "00") & "DEN5").Value = rs("MS" & Format(i, "00") & "DEN5")      'WFGD_測定値xx Den5
            .Fields("WFGD_MS" & Format(i, "00") & "LDL1").Value = rs("MS" & Format(i, "00") & "LDL1")      'WFGD_測定値xx L/DL1
            .Fields("WFGD_MS" & Format(i, "00") & "LDL2").Value = rs("MS" & Format(i, "00") & "LDL2")      'WFGD_測定値xx L/DL2
            .Fields("WFGD_MS" & Format(i, "00") & "LDL3").Value = rs("MS" & Format(i, "00") & "LDL3")      'WFGD_測定値xx L/DL3
            .Fields("WFGD_MS" & Format(i, "00") & "LDL4").Value = rs("MS" & Format(i, "00") & "LDL4")      'WFGD_測定値xx L/DL4
            .Fields("WFGD_MS" & Format(i, "00") & "LDL5").Value = rs("MS" & Format(i, "00") & "LDL5")      'WFGD_測定値xx L/DL5
        Next
        
        For i = 1 To 5
            .Fields("WFGD_MS01DVD2" & i).Value = rs("MS0" & i & "DVD2")                                    'WFGD_測定値xx DVD2
        Next
        
    End With
    Set rs = Nothing

    getTBCMJ015WFGD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ015WFGD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :SXL確定指示(TBCMY007)ﾃｰﾌﾞﾙにｾｯﾄするSXLの比抵抗ﾃﾞｰﾀを取得する。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO  ,型                :説明
'          :SXLID          ,I   ,String            ,SXLID
'　　      :sPos  　　　    ,I   ,String 　         ,SXL位置(TOP/BOT)   04/04/15 ooba
'          :sPattern       ,I   ,String            ,比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝ
'                                                   ●ﾊﾟﾀｰﾝA : WF実績ﾃﾞｰﾀ取得
'                                                   ●ﾊﾟﾀｰﾝB : 結晶実績ﾃﾞｰﾀ取得
'                                                   ●ﾊﾟﾀｰﾝC : 取得ﾃﾞｰﾀなし
'          :mesdata()      ,O   ,String            ,比抵抗ﾃﾞｰﾀ
'          :戻り値          ,O   ,FUNCTION_RETURN   ,抽出の成否
'説明      :
'履歴      :04/02/12 ooba　作成
Public Function cmbc040_GetSxlRsData(SXLID As String, sPos As String, sPattern As String, mesdata() As String) As FUNCTION_RETURN
    
    Dim sTBkbn As String        'T/B区分
    Dim i As Integer
    Dim j As Integer
    Dim sSql As String
    Dim rs As OraDynaset
    Dim dTmpData(10) As Double   '比抵抗(Rs)ﾃﾞｰﾀ
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function cmbc040_GetSxlRsData"
    cmbc040_GetSxlRsData = FUNCTION_RETURN_FAILURE
    
    If sPos = "TOP" Then sTBkbn = "T" Else sTBkbn = "B"  '04/04/15 ooba
    
    '比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝが『A』の場合、WF実績ﾃﾞｰﾀ(TBCMY013)を取得する。
    If sPattern = "A" Then
'''        For i = 1 To 2
'''            If i = 1 Then sTBkbn = "T" Else sTBkbn = "B"
        '該当SXLより、新ｻﾝﾌﾟﾙ管理-WF<XSDCW>のｻﾝﾌﾟﾙID_Rsを取得。
        'ｻﾝﾌﾟﾙID_Rsから、測定評価結果<TBCMY013>の比抵抗実績ﾃﾞｰﾀ(TOP側/BOT側)を取得する。
        sSql = "select MESDATA1, MESDATA2, MESDATA3, MESDATA4, MESDATA5 "
        sSql = sSql & "from TBCMY013 "
        sSql = sSql & "where OSITEM = 'RES' "
        sSql = sSql & "and SAMPLEID in ( "
        sSql = sSql & "         select WFSMPLIDRSCW from XSDCW "
        sSql = sSql & "         where TBKBNCW = '" & sTBkbn & "' "
        sSql = sSql & "         and SXLIDCW = '" & SXLID & "') "
        
        Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
        
        If rs.RecordCount > 0 Then
            'TOP側実績ﾃﾞｰﾀ
            If sTBkbn = "T" Then
                mesdata(1) = rs("MESDATA1")
                mesdata(2) = rs("MESDATA2")
                mesdata(3) = rs("MESDATA3")
                mesdata(4) = rs("MESDATA4")
                mesdata(5) = rs("MESDATA5")
            'BOT側実績ﾃﾞｰﾀ
            ElseIf sTBkbn = "B" Then
                mesdata(6) = rs("MESDATA1")
                mesdata(7) = rs("MESDATA2")
                mesdata(8) = rs("MESDATA3")
                mesdata(9) = rs("MESDATA4")
                mesdata(10) = rs("MESDATA5")
            End If
        Else
            '実績ﾃﾞｰﾀがない場合はｴﾗｰ
            Set rs = Nothing
            GoTo proc_exit
        End If
        Set rs = Nothing
'''        Next
    '比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝが『B』の場合、結晶実績ﾃﾞｰﾀ(TBCMJ002)を取得する。
    ElseIf sPattern = "B" Then
'''        For i = 1 To 2
'''            If i = 1 Then sTBkbn = "T" Else sTBkbn = "B"
        '該当SXLより、新ｻﾝﾌﾟﾙ管理-WF<XSDCW>のT/B区分、ｻﾝﾌﾟﾙﾌﾞﾛｯｸIDを取得。
        'T/B区分、ｻﾝﾌﾟﾙﾌﾞﾛｯｸIDから、新ｻﾝﾌﾟﾙ管理-ﾌﾞﾛｯｸ<XSDCS>の結晶番号、ｻﾝﾌﾟﾙID_Rsを取得。
        '結晶番号、ｻﾝﾌﾟﾙID_Rsから、結晶抵抗実績<TBCMJ002>の比抵抗実績ﾃﾞｰﾀ(TOP側/BOT側)を取得する。
        sSql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 "
        sSql = sSql & "from TBCMJ002 "
        sSql = sSql & "where (CRYNUM, SMPLNO) in ( "
        sSql = sSql & "         select XTALCS, CRYSMPLIDRSCS "
        sSql = sSql & "         from XSDCS "
        sSql = sSql & "         where (TBKBNCS, CRYNUMCS) in ( "
        sSql = sSql & "                  select TBKBNCW, SMCRYNUMCW "
        sSql = sSql & "                  from XSDCW "
        sSql = sSql & "                  where TBKBNCW = '" & sTBkbn & "' "
        sSql = sSql & "                  and SXLIDCW = '" & SXLID & "')) "
        sSql = sSql & "and TRANCNT = ( "
        sSql = sSql & "         select max(TRANCNT) "
        sSql = sSql & "         from TBCMJ002 "
        sSql = sSql & "         where (CRYNUM, SMPLNO) in ( "
        sSql = sSql & "                  select XTALCS, CRYSMPLIDRSCS "
        sSql = sSql & "                  from XSDCS "
        sSql = sSql & "                  where (TBKBNCS, CRYNUMCS) in ( "
        sSql = sSql & "                           select TBKBNCW, SMCRYNUMCW "
        sSql = sSql & "                           from XSDCW "
        sSql = sSql & "                           where TBKBNCW = '" & sTBkbn & "' "
        sSql = sSql & "                           and SXLIDCW = '" & SXLID & "'))) "
    
        Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
    
        If rs.RecordCount > 0 Then
            'TOP側実績ﾃﾞｰﾀ
            If sTBkbn = "T" Then
                dTmpData(1) = rs("MEAS1")
                dTmpData(2) = rs("MEAS2")
                dTmpData(3) = rs("MEAS3")
                dTmpData(4) = rs("MEAS4")
                dTmpData(5) = rs("MEAS5")
                '型変換
                For j = 1 To 5
                    mesdata(j) = CStr(dTmpData(j))
                Next
            'BOT側実績ﾃﾞｰﾀ
            ElseIf sTBkbn = "B" Then
                dTmpData(6) = rs("MEAS1")
                dTmpData(7) = rs("MEAS2")
                dTmpData(8) = rs("MEAS3")
                dTmpData(9) = rs("MEAS4")
                dTmpData(10) = rs("MEAS5")
                '型変換
                For j = 6 To 10
                    mesdata(j) = CStr(dTmpData(j))
                Next
            End If
        Else
            '実績ﾃﾞｰﾀがない場合はｴﾗｰ
            Set rs = Nothing
            GoTo proc_exit
        End If
        Set rs = Nothing
'''        Next
    '比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝが『C』の場合、取得実績ﾃﾞｰﾀなし。
    ElseIf sPattern = "C" Then
    
    End If
    
    '取得ﾃﾞｰﾀが空白/-1/NULLの時はｽﾍﾟｰｽをｾｯﾄする。
'''    For i = 1 To 10
'''        If mesdata(i) = "" Or mesdata(i) = "-1" Or mesdata(i) = vbNullString Then
'''            mesdata(i) = " "
'''        End If
'''    Next
    For i = 1 To 5
        If sTBkbn = "T" Then j = i Else j = i + 5
        If mesdata(j) = "" Or mesdata(j) = "-1" Or mesdata(j) = vbNullString Then
            mesdata(j) = " "
        End If
    Next
    
    cmbc040_GetSxlRsData = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    cmbc040_GetSxlRsData = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


Private Function NtoZ2(strWk As String) As Double
    If Trim(strWk) = "" Then
        NtoZ2 = -1
        Exit Function
    End If
    NtoZ2 = CDbl(strWk)
End Function


Private Function CryRES_Judg(CRs() As Double, GarRes As Guarantee) As Double
    Dim pt As Integer

    ''RRG判定
    Select Case GarRes.cPos
      Case "B", "C", "D", "E", "F", "K", "S", "Y"
          Select Case GarRes.cBunp
          Case "A", "B", "C", "M"
             ''RRG計算
             CryRES_Judg = MENNAI_Cal(RES_JUDG, CRs(), GarRes, GarRes.cBunp)

          Case "", " "                                          'ｽﾍﾟｰｽ追加　05/07/05 ooba
'Add Start 2011/10/13 Y.Hitomi
                CryRES_Judg = -1
                Exit Function
'Add End 2011/10/13 Y.Hitomi

'Del Start 2011/10/13 Y.Hitomi
''Cng Start 2011/09/19 Y.Hitomi
'            If GarRes.cCount = "1" Then
'                CryRES_Judg = 0  '１点測定の場合のみ０とする
'            Else
'                CryRES_Judg = -1
'                Exit Function
'            End If
''                CryRES_Judg = -1
''                Exit Function
''Cng End   2011/09/19 Y.Hitomi
'Del End 2011/10/13 Y.Hitomi

          Case Else
'Cng Start 2011/10/25 Y.Hitomi
                CryRES_Judg = -2
                Exit Function
'             ''RRG計算　　　コード "A" にて計算
'             If Trim(GarRes.cCount) = "" Then
'                pt = 3
'             Else
'                pt = val(GarRes.cCount)
'             End If
'             CryRES_Judg = RoundUp((RGCal(CRs(), pt)), 4)
'Cng End 2011/10/25 Y.Hitomi
         End Select
      Case Else
         Select Case GarRes.cBunp
         Case "A", "B", "C", "D", "E", "M", "N"
             ''RRG計算
             CryRES_Judg = MENNAI_Cal(RES_JUDG, CRs(), GarRes, GarRes.cBunp)

         Case "", " " 'ｽﾍﾟｰｽ追加　05/07/05 ooba
'Add Start 2011/10/13 Y.Hitomi
                CryRES_Judg = -1
                Exit Function
'Add End 2011/10/13 Y.Hitomi

'Del Start 2011/10/13 Y.Hitomi
''Cng Start 2011/09/19 Y.Hitomi
'            If GarRes.cCount = "1" Then
'                CryRES_Judg = 0  '１点測定の場合のみ０とする
'            Else
'                CryRES_Judg = -1
'                Exit Function
'            End If
''                CryRES_Judg = -1
''                Exit Function
''Cng End   2011/09/19 Y.Hitomi
'Del End 2011/10/13 Y.Hitomi
         Case Else
'Cng Start 2011/10/25 Y.Hitomi
                CryRES_Judg = -2
                Exit Function
'             ''RRG計算　　　コード "A" にて計算
'             If Trim(GarRes.cCount) = "" Then
'                pt = 3
'             Else
'                pt = val(GarRes.cCount)
'             End If
'             CryRES_Judg = RoundUp((RGCal(CRs(), pt)), 4)
'Cng End 2011/10/25 Y.Hitomi
         End Select
    End Select
Cal_Escp:
        
End Function

Private Function CryOi_Judg(COi() As Double, GarOi As Guarantee) As Double
    Dim pt As Integer
    ReDim JData(UBound(COi())) As Double
    
    ''ORG判定
    
    Select Case GarOi.cPos
      Case "B", "C", "D", "E", "F", "K", "Y"
          Select Case GarOi.cBunp
          Case "A", "B", "C"
             ''ORG計算
             CryOi_Judg = MENNAI_Cal(OI_JUDG, COi(), GarOi, GarOi.cBunp)

          Case "", " "                                              'ｽﾍﾟｰｽ追加　05/07/05 ooba
'Add Start 2011/10/13 Y.Hitomi
                CryOi_Judg = -1
                Exit Function
'Add End 2011/10/13 Y.Hitomi
             
             ''計算区分がスペースの場合は、計算，判定を行わない
'             If GarOi.cBunp = "" Or GarOi.cBunp = " " Then         '→ｺﾒﾝﾄ化　05/07/05 ooba
'                    GoTo Cal_Escp
'Del Start 2011/10/13 Y.Hitomi
''Cng Start 2011/09/19 Y.Hitomi
'            If GarOi.cCount = "1" Then
'                CryOi_Judg = 0  '１点測定の場合のみ０とする
'            Else
'                CryOi_Judg = -1
'                Exit Function
'            End If
''                CryOi_Judg = -1
''                Exit Function
''Cng End   2011/09/19 Y.Hitomi
'Del End 2011/10/13 Y.Hitomi

'             End If                                                '→ｺﾒﾝﾄ化　05/07/05 ooba

          Case Else
'Cng Start 2011/10/25 Y.Hitomi
                CryOi_Judg = -2
                Exit Function
'             ''ORG計算　　　コード "A" にて計算
'             If Trim(GarOi.cCount) = "" Then
'                pt = 3
'             Else
'                pt = val(GarOi.cCount)
'             End If
'             CryOi_Judg = RoundUp((RGCal(COi(), pt)), 4)
'Cng End 2011/10/25 Y.Hitomi

         End Select

      Case Else

         Select Case GarOi.cBunp
         Case "A", "B", "C", "D", "E", "N"
             ''ORG計算
             CryOi_Judg = MENNAI_Cal(OI_JUDG, COi(), GarOi, GarOi.cBunp)

         Case "", " "                                               'ｽﾍﾟｰｽ追加　05/07/05 ooba
'Add Start 2011/10/13 Y.Hitomi
                CryOi_Judg = -1
                Exit Function
'Add End 2011/10/13 Y.Hitomi
             
             ''計算区分がスペースの場合は、計算，判定を行わない
'             If GarOi.cBunp = "" Or GarOi.cBunp = " " Then         '→ｺﾒﾝﾄ化　05/07/05 ooba
'                    GoTo Cal_Escp
'Del Start 2011/10/13 Y.Hitomi
''Cng Start 2011/09/19 Y.Hitomi
'            If GarOi.cCount = "1" Then
'                CryOi_Judg = 0  '１点測定の場合のみ０とする
'            Else
'                CryOi_Judg = -1
'                Exit Function
'            End If
''                CryOi_Judg = -1
''                Exit Function
''Cng End   2011/09/19 Y.Hitomi
'Del End 2011/10/13 Y.Hitomi
'Cng End   2011/09/19 Y.Hitomi
'             End If                                                '→ｺﾒﾝﾄ化　05/07/05 ooba

         Case Else
'Cng Start 2011/10/25 Y.Hitomi
                CryOi_Judg = -2
                Exit Function
'             ''ORG計算　　　コード "A" にて計算
'             If Trim(GarOi.cCount) = "" Then
'                pt = 3
'             Else
'                pt = val(GarOi.cCount)
'             End If
'             CryOi_Judg = RoundUp((RGCal(COi(), pt)), 4)
'Cng End 2011/10/25 Y.Hitomi

         End Select
    End Select
Cal_Escp:

End Function

'概要      :BMD実績のMin値を計算する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :dMin          ,O   ,Double    ,Min値
'          :strMeasPos    ,I   ,String    ,結晶欠陥測定位置コード（3byte）
'          :dMeas()       ,I   ,Double    ,測定位置配列
'          :戻り値        ,O   ,Integer     ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Private Function getSXLBMDMIN(dMin As Double, strMeasPos As String, dMeas() As Double) As Integer
    Dim dConv       As Double
    Dim iMeasNum    As Integer
    Dim Index       As Integer
    Dim dForMin()   As Double
    Dim strParam    As String

    On Error GoTo Err
    getSXLBMDMIN = FUNCTION_RETURN_FAILURE

    If strMeasPos = "" Then
        dMin = -1
        getSXLBMDMIN = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If

    '' 結晶欠陥測定位置（測定方法）より換算係数を取得
    strParam = GetCodeField("GP", "01", Mid(strMeasPos, 1, 1), "INFO8")
    If strParam = vbNullString Then strParam = "1"
    dConv = val(strParam)

    '' 結晶欠陥測定位置（測定点）の取得
    iMeasNum = GetMeasureNum(Mid(strMeasPos, 2, 1), 1)
    If iMeasNum < 1 Then Exit Function

    '' Min値計算
    ReDim dForMin(iMeasNum - 1)
    For Index = 0 To UBound(dForMin)
        dForMin(Index) = dMeas(Index)
    Next Index
    dMin = GetMin(dForMin) * dConv / 10000

    getSXLBMDMIN = FUNCTION_RETURN_SUCCESS
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
End Function



'概要      :測定結果を計算する（ライフタイム実績）
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :iResult       ,O   ,Integer   ,計算結果
'          :iParam()      ,I   ,Integer   ,測定値配列
'          :sHsxLtspi     ,I   ,String    ,測定位置         (新データ[10点測定]は3,5,Aのどれかを設定する)
'          :iOldFlg       ,I   ,Integer   ,旧データフラグ   (旧データ[5点測定]は1を設定する)
'          :戻り値        ,O   ,Integer   ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :2005/11/07 牧野 変更　10点測定対応
Public Function KNS_CalculateMeasResult_LT(iResult As Integer, iParam() As Integer, _
                    sHsxLtspi As String, iOldFlg As Integer) As Integer
    Dim Index   As Integer
    Dim iAve    As Integer

    On Error GoTo Err
    KNS_CalculateMeasResult_LT = FUNCTION_RETURN_FAILURE
    
    '' 旧データの場合（５点測定）
    If iOldFlg = 1 Then
        '' パラメータ入力チェック
        For Index = 0 To KNS_GetMeasureNum_LT(iOldFlg) - 1
            If iParam(Index) = DEF_PARAM_VALUE_LT Then
                Exit Function
            End If
        Next Index
        ''３，４，５点の測定点のAVEを求める
        iAve = RoundDown((iParam(2) + iParam(3) + iParam(4)) / 3#, 0)

        '' 測定点２とAVE値を比較、値の小さい方を測定結果とする
        If iAve < iParam(1) Then
            iResult = iAve
        Else
            iResult = iParam(1)
        End If

    '' 新データの場合（１０点測定）
    Else
        '' パラメータ入力チェック
        For Index = 0 To KNS_GetMeasureNum_LT(iOldFlg) - 1
            If iParam(Index) = DEF_PARAM_VALUE_LT Then
                Exit Function
            End If
        Next Index

        ''' [A:Ce,Inside3mm]の場合
        If Trim(sHsxLtspi) = "3" Then
            ''８，９，１０点の測定点のAVEを求める
            iAve = RoundDown((iParam(7) + iParam(8) + iParam(9)) / 3#, 0)

        ''' [A:Ce,Inside5mm]の場合
        ElseIf Trim(sHsxLtspi) = "5" Then
            ''５，６，７点の測定点のAVEを求める
            iAve = RoundDown((iParam(4) + iParam(5) + iParam(6)) / 3#, 0)

        ''' [A:Ce,Inside10mm]の場合
        ElseIf Trim(sHsxLtspi) = "A" Then
            ''２，３，４点の測定点のAVEを求める
            iAve = RoundDown((iParam(1) + iParam(2) + iParam(3)) / 3#, 0)

        ''' その他の場合は[A:Ce,Inside10mm]の仕様とする
        Else
            ''２，３，４点の測定点のAVEを求める
            iAve = RoundDown((iParam(1) + iParam(2) + iParam(3)) / 3#, 0)

        End If
    
        '' 測定点１とAVE値を比較、値の小さい方を測定結果とする
        If iAve < iParam(0) Then
            iResult = iAve
        Else
            iResult = iParam(0)
        End If
    End If

    KNS_CalculateMeasResult_LT = FUNCTION_RETURN_SUCCESS
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
End Function

'2008/05/30 SHINDOH------------------------------------------------------------------
'●赤黒処理の赤部分作成●
    ' ブロック管理の更新（内部関数）
Public Function DBDRV_redXSDC3(records As typ_XSDC3_Update) As FUNCTION_RETURN

    Dim sql As String
    Dim strErrMsg As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function DBDRV_redXSDC3"

    DBDRV_redXSDC3 = FUNCTION_RETURN_SUCCESS

    ' ブロック管理の更新
    With records
        .KCNTC3 = .KCNTC3 + 1
        .MODKBC3 = 1
        .FRWC3 = .FRWC3 * (-1)
        .FRLC3 = .FRLC3 * (-1)
        .FUWC3 = .FUWC3 * (-1)
        .FULC3 = .FULC3 * (-1)
        .TOLC3 = .TOLC3 * (-1)
        .TOWC3 = .TOWC3 * (-1)
        .SNDKC3 = 0
        .SUMITBC3 = 0
    End With
    If CreateXSDC3(records, strErrMsg) = FUNCTION_RETURN_FAILURE Then
        DBDRV_redXSDC3 = FUNCTION_RETURN_FAILURE
    End If
   
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_redXSDC3 = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'-----------------------------------------------------------------------------
'概要      :テーブル「XSDC3」から条件に赤黒処理用レコードの抽出
'説明      :
'-----------------------------------------------------------------------------
Public Function DBDRV_GetredXSDC3(records As typ_XSDC3_Update, t_CryNum As String, t_INPOS As Integer) As FUNCTION_RETURN
    
    Dim sql As String       'SQL全体
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      'レコード数
    Dim i As Long
     'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function DBDRV_GetredXSDC3"

    sql = ""
    sql = sql & "select * from XSDC3"
    sql = sql & " Where CRYNUMC3='" & t_CryNum & "'"
    sql = sql & " and INPOSC3=" & t_INPOS
    sql = sql & " and KCNTC3= ( SELECT MAX(KCNTC3) FROM XSDC3 WHERE "
    sql = sql & " CRYNUMC3='" & t_CryNum & "'"
    sql = sql & " and INPOSC3=" & t_INPOS
    sql = sql & " Group by CRYNUMC3, INPOSC3 )"
    
Debug.Print "getredXSDC3 " & sql
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        DBDRV_GetredXSDC3 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        Exit Function
    End If
    With records
        If IsNull(rs.Fields("CRYNUMC3")) = False Then .CRYNUMC3 = rs.Fields("CRYNUMC3")
        If IsNull(rs.Fields("INPOSC3")) = False Then .INPOSC3 = rs.Fields("INPOSC3")
        If IsNull(rs.Fields("KCNTC3")) = False Then .KCNTC3 = rs.Fields("KCNTC3")
        If IsNull(rs.Fields("HINBC3")) = False Then .HINBC3 = rs.Fields("HINBC3")
        If IsNull(rs.Fields("REVNUMC3")) = False Then .REVNUMC3 = rs.Fields("REVNUMC3")
        If IsNull(rs.Fields("FACTORYC3")) = False Then .FACTORYC3 = rs.Fields("FACTORYC3")
        If IsNull(rs.Fields("OPEC3")) = False Then .OPEC3 = rs.Fields("OPEC3")
        If IsNull(rs.Fields("LENC3")) = False Then .LENC3 = rs.Fields("LENC3")
        If IsNull(rs.Fields("XTALC3")) = False Then .XTALC3 = rs.Fields("XTALC3")
        If IsNull(rs.Fields("SXLIDC3")) = False Then .SXLIDC3 = rs.Fields("SXLIDC3")
        If IsNull(rs.Fields("KNKTC3")) = False Then .KNKTC3 = rs.Fields("KNKTC3")
        If IsNull(rs.Fields("WKKTC3")) = False Then .WKKTC3 = rs.Fields("WKKTC3")
        If IsNull(rs.Fields("WKKBC3")) = False Then .WKKBC3 = rs.Fields("WKKBC3")
        If IsNull(rs.Fields("MACOC3")) = False Then .MACOC3 = rs.Fields("MACOC3")
        If IsNull(rs.Fields("MODKBC3")) = False Then .MODKBC3 = rs.Fields("MODKBC3")
        If IsNull(rs.Fields("SUMKBC3")) = False Then .SUMKBC3 = rs.Fields("SUMKBC3")
        If IsNull(rs.Fields("FRKNKTC3")) = False Then .FRKNKTC3 = rs.Fields("FRKNKTC3")
        If IsNull(rs.Fields("FRWKKTC3")) = False Then .FRWKKTC3 = rs.Fields("FRWKKTC3")
        If IsNull(rs.Fields("FRWKKBC3")) = False Then .FRWKKBC3 = rs.Fields("FRWKKBC3")
        If IsNull(rs.Fields("FRMACOC3")) = False Then .FRMACOC3 = rs.Fields("FRMACOC3")
        If IsNull(rs.Fields("TOWNKTC3")) = False Then .TOWNKTC3 = rs.Fields("TOWNKTC3")
        If IsNull(rs.Fields("TOWKKTC3")) = False Then .TOWKKTC3 = rs.Fields("TOWKKTC3")
        If IsNull(rs.Fields("TOMACOC3")) = False Then .TOMACOC3 = rs.Fields("TOMACOC3")
        If IsNull(rs.Fields("FRLC3")) = False Then .FRLC3 = rs.Fields("FRLC3")
        If IsNull(rs.Fields("FRWC3")) = False Then .FRWC3 = rs.Fields("FRWC3")
        If IsNull(rs.Fields("FRMC3")) = False Then .FRMC3 = rs.Fields("FRMC3")
        If IsNull(rs.Fields("FULC3")) = False Then .FULC3 = rs.Fields("FULC3")
        If IsNull(rs.Fields("FUWC3")) = False Then .FUWC3 = rs.Fields("FUWC3")
        If IsNull(rs.Fields("FUMC3")) = False Then .FUMC3 = rs.Fields("FUMC3")
        If IsNull(rs.Fields("LOSWC3")) = False Then .LOSWC3 = rs.Fields("LOSWC3")
        If IsNull(rs.Fields("LOSLC3")) = False Then .LOSLC3 = rs.Fields("LOSLC3")
        If IsNull(rs.Fields("LOSMC3")) = False Then .LOSMC3 = rs.Fields("LOSMC3")
        If IsNull(rs.Fields("TOLC3")) = False Then .TOLC3 = rs.Fields("TOLC3")
        If IsNull(rs.Fields("TOWC3")) = False Then .TOWC3 = rs.Fields("TOWC3")
        If IsNull(rs.Fields("TOMC3")) = False Then .TOMC3 = rs.Fields("TOMC3")
        If IsNull(rs.Fields("SUMITLC3")) = False Then .SUMITLC3 = rs.Fields("SUMITLC3")
        If IsNull(rs.Fields("SUMITWC3")) = False Then .SUMITWC3 = rs.Fields("SUMITWC3")
        If IsNull(rs.Fields("SUMITMC3")) = False Then .SUMITMC3 = rs.Fields("SUMITMC3")
        If IsNull(rs.Fields("MOTHINC3")) = False Then .MOTHINC3 = rs.Fields("MOTHINC3")
        If IsNull(rs.Fields("XTWORKC3")) = False Then .XTWORKC3 = rs.Fields("XTWORKC3")
        If IsNull(rs.Fields("WFWORKC3")) = False Then .WFWORKC3 = rs.Fields("WFWORKC3")
        If IsNull(rs.Fields("STATIMEC3")) = False Then .STATIMEC3 = rs.Fields("STATIMEC3")
        If IsNull(rs.Fields("STOTIMEC3")) = False Then .STOTIMEC3 = rs.Fields("STOTIMEC3")
        If IsNull(rs.Fields("ETIMEC3")) = False Then .ETIMEC3 = rs.Fields("ETIMEC3")
        If IsNull(rs.Fields("HOLDCC3")) = False Then .HOLDCC3 = rs.Fields("HOLDCC3")
        If IsNull(rs.Fields("HOLDBC3")) = False Then .HOLDBC3 = rs.Fields("HOLDBC3")
        If IsNull(rs.Fields("LDFRCC3")) = False Then .LDFRCC3 = rs.Fields("LDFRCC3")
        If IsNull(rs.Fields("LDFRBC3")) = False Then .LDFRBC3 = rs.Fields("LDFRBC3")
        If IsNull(rs.Fields("TSTAFFC3")) = False Then .TSTAFFC3 = rs.Fields("TSTAFFC3")
        If IsNull(rs.Fields("TDAYC3")) = False Then .TDAYC3 = rs.Fields("TDAYC3")
        If IsNull(rs.Fields("KSTAFFC3")) = False Then .KSTAFFC3 = rs.Fields("KSTAFFC3")
        If IsNull(rs.Fields("KDAYC3")) = False Then .KDAYC3 = rs.Fields("KDAYC3")
        If IsNull(rs.Fields("SUMITBC3")) = False Then .SUMITBC3 = rs.Fields("SUMITBC3")
        If IsNull(rs.Fields("SNDKC3")) = False Then .SNDKC3 = rs.Fields("SNDKC3")
        If IsNull(rs.Fields("SNDDAYC3")) = False Then .SNDDAYC3 = rs.Fields("SNDDAYC3")
        If IsNull(rs.Fields("SUMDAYC3")) = False Then .SUMDAYC3 = rs.Fields("SUMDAYC3")
        If IsNull(rs.Fields("PAYCLASSC3")) = False Then .PAYCLASSC3 = rs.Fields("PAYCLASSC3")
        If IsNull(rs.Fields("CUTCNTC3")) = False Then .CUTCNTC3 = rs.Fields("CUTCNTC3")
        If IsNull(rs.Fields("HINBFLGC3")) = False Then .HINBFLGC3 = rs.Fields("HINBFLGC3")
        If IsNull(rs.Fields("RPCRYNUMC3")) = False Then .RPCRYNUMC3 = rs.Fields("RPCRYNUMC3")
        If IsNull(rs.Fields("PLANTCATC3")) = False Then .PLANTCATC3 = rs.Fields("PLANTCATC3")
        End With
        rs.MoveNext
    rs.Close

    DBDRV_GetredXSDC3 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_GetredXSDC3 = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function

'不良長さﾁｪｯｸ　08/06/27 ooba
Public Function chkFuryoLen(lJitsuLen As Long, lSeiLen As Long, lMaxLen As Long) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset

    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function chkFuryoLen"
    
    chkFuryoLen = FUNCTION_RETURN_FAILURE
    lMaxLen = 0
    
    sql = "select CTR01A9 from KODA9 "
    sql = sql & "where SYSCA9 = 'X' "
    sql = sql & "and SHUCA9 = '19' "
    sql = sql & "and CODEA9 = 'HARAISXL' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    'ﾃﾞｰﾀなし
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If

    If IsNull(rs("CTR01A9")) = False Then lMaxLen = rs("CTR01A9")       '不良長さ上限
    
    rs.Close
    
    
    '入力不良長さ(実長さ－製品長さ)が上限値以下の場合はOK。
    If lMaxLen = 0 Or (lJitsuLen - lSeiLen) <= lMaxLen Then
        chkFuryoLen = FUNCTION_RETURN_SUCCESS
    End If
    
proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function

'長さ修正時の範囲ﾁｪｯｸ　08/06/30 ooba
Public Function chkRange(lUkeLen As Long, lHaraiLen As Long, _
                                    lMaxLen As Long, lMinLen As Long) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset

    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function chkRange"
    
    chkRange = FUNCTION_RETURN_FAILURE
    lMaxLen = 0
    lMinLen = 0
    
    sql = "select CTR01A9,CTR02A9 from KODA9 "
    sql = sql & "where SYSCA9 = 'X' "
    sql = sql & "and SHUCA9 = '19' "
    sql = sql & "and CODEA9 = 'HARAIALT' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    'ﾃﾞｰﾀなし
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If

    If IsNull(rs("CTR01A9")) = False Then lMaxLen = rs("CTR01A9")       '上限
    If IsNull(rs("CTR02A9")) = False Then lMinLen = rs("CTR02A9")       '下限
    
    rs.Close
    
    
    '払出実長さと受入実長さの差が上限／下限に含まれる場合はOK。
    '上限ﾁｪｯｸ
    If lMaxLen = 0 Or (lHaraiLen - lUkeLen) <= lMaxLen Then
        '下限ﾁｪｯｸ
        If lMinLen = 0 Or (lUkeLen - lHaraiLen) <= lMinLen Then
            chkRange = FUNCTION_RETURN_SUCCESS
        End If
    End If
    
proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function

'TBCST001ﾚｺｰﾄﾞ件数取得＆更新　08/06/30 ooba
Public Function GetTBCST001cnt(sBlkID As String, iCnt As Integer) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset

    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function GetTBCST001cnt"
    
    GetTBCST001cnt = FUNCTION_RETURN_FAILURE
    iCnt = 0
    
    sql = "select BLOCKID from TBCST001 "
    sql = sql & "where BLOCKID = '" & sBlkID & "' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    'ﾚｺｰﾄﾞ件数
    iCnt = rs.RecordCount
    
    rs.Close
    
    '既存ﾃﾞｰﾀを送信対象外とする
    If iCnt > 0 Then
        sql = "update TBCST001 "
        sql = sql & "set SENDFLAG = '5' "
        sql = sql & "where BLOCKID = '" & sBlkID & "' "
        
        If OraDB.ExecuteSQL(sql) <= 0 Then
            GoTo proc_exit
        End If
    End If
    
    GetTBCST001cnt = FUNCTION_RETURN_SUCCESS
    
    
proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function

'TBCMX011処理回数取得　08/09/12 ooba
Public Function GetTBCMX011cnt(sBlkID As String, iCnt As Integer) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset

    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function GetTBCMX011cnt"
    
    GetTBCMX011cnt = FUNCTION_RETURN_FAILURE
    
    iCnt = 1
    
    sql = "select MAX(TRANCNT) TRANCNT from TBCMX011 "
    sql = sql & "where BLOCKID = '" & sBlkID & "' "
    sql = sql & "group by BLOCKID "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    If rs.RecordCount > 0 Then
        '処理回数
        iCnt = rs("TRANCNT") + 1
    End If
    
    rs.Close

    
    GetTBCMX011cnt = FUNCTION_RETURN_SUCCESS
    
    
proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function

'Add Start 2012/01/31 Y.Hitomi
'WF払い出し長さチェック
Public Function fncChkWFLen(lUkeLen As Long, lHaraiLen As Long, _
                                    lMaxLen As Long, lMinLen As Long) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset

    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function chkRange"
    
    fncChkWFLen = FUNCTION_RETURN_FAILURE
    lMaxLen = 0
    lMinLen = 0
    
    sql = "select CTR01A9,CTR02A9 from KODA9 "
    sql = sql & "where SYSCA9 = 'X' "
    sql = sql & "and SHUCA9 = '19' "
    sql = sql & "and CODEA9 = 'HARAIWFLEN' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    'ﾃﾞｰﾀなし
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If

    If IsNull(rs("CTR01A9")) = False Then lMinLen = rs("CTR01A9")       '下限
    If IsNull(rs("CTR02A9")) = False Then lMaxLen = rs("CTR02A9")       '上限
    
    rs.Close
    
    '払出実長さと上下限チェック
    If lHaraiLen >= lMinLen And lHaraiLen <= lMaxLen Then
        fncChkWFLen = FUNCTION_RETURN_SUCCESS
    End If
    
proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function
'SIRD先行評価ブロックSXL出荷流動チェック
Public Function FncChkSird(sBlockId As String) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset

    On Error GoTo proc_err
    gErr.Push "s_cmbc032_SQL.bas -- Function chkRange"
    
    FncChkSird = FUNCTION_RETURN_FAILURE
    
    sql = "select SIRDKBNY3 from XODY3 "
    sql = sql & "where XTALNOY3 = '" & sBlockId & "' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    'ﾃﾞｰﾀなし
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If

    
    '対象ブロックが、SIRD評価ブロックでなければチェックOK
    If rs("SIRDKBNY3") <> "2" Then
        FncChkSird = FUNCTION_RETURN_SUCCESS
    End If
    
    rs.Close
    

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function
'Add End 2012/01/31 Y.Hitomi


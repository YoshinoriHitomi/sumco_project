Attribute VB_Name = "s_cmbc033_SQL"
Option Explicit

' 抜試指示

Public lStfMst As Long
Public intEnCmd As Integer
Public Const MAXCNT As Integer = 16                             ' 最大件数
Public Const BlkTop As Integer = 1                                 ' TOP側
Public Const BlkTail As Integer = 2                                ' TAIL側
Public Const KSYSCLASS As String = "GP"                         ' システム区分
Public Const MSYSCLASS As String = "NM"                         ' システム区分
Public Const KCLASS As String = "01"                            ' クラス
Public Const KCODE As String = "1"                              ' コード

' ブロック情報
Public Type typ_BlkInf1
    BLOCKID As String * 12      ' ブロックID
    LENGTH As Integer           ' 長さ
    REALLEN As Integer          ' 実長さ
    KRPROCCD As String * 5      ' 現在管理工程
    NOWPROC As String * 5       ' 現在工程
    LPKRPROCCD As String * 5    ' 最終通過管理工程
    LASTPASS As String * 5      ' 最終通過工程
    RSTATCLS As String * 1      ' 流動状態区分
    BDCODE As String * 3        ' 不良理由コード
    PALTNUM As String * 4       ' パレット番号
    SEED As String * 4          ' シード
    COF As type_Coefficient     ' 偏析係数計算
    SAMPFLAG As Boolean         ' サンプル取得フラグ
End Type

Type cmkc001b_LockWait
    flag As Boolean
    Grp As Integer
End Type
Type cmkc001b_Wait3_HINBAN
    HINBAN As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    factory As String * 1           ' 工場
    opecond As String * 1           ' 操業条件
End Type
Type cmkc001b_Wait3_BLK
    BLOCKID As String * 12          ' ブロックID
    IngotPos As Integer             ' 結晶内開始位置
    LENGTH As Integer               ' 長さ
    NOWPROC As String * 5           ' 現在工程
    HOLDCLS As String * 1           ' ホールド区分 ---kuramoto 追加 2001/09/19----
    GRPFLG1 As Integer           ' グループ情報
    GRPFLG2 As Integer           ' グループ情報
    COLORFLG As Boolean
    topHin As cmkc001b_Wait3_HINBAN
    botHin As cmkc001b_Wait3_HINBAN
End Type
Type cmkc001b_Wait3
    CRYNUM As String * 12           ' 結晶番号
    blkInfo() As cmkc001b_Wait3_BLK
End Type

Type type_cmkc001b_SmpMng
    CRYNUM As String * 12
    IngotPos As Integer
    SMPKBN As String * 1
    
    HINBAN As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    factory As String * 1           ' 工場
    opecond As String * 1           ' 操業条件
    
    
    CRYINDRS As String * 1
    CRYRESRS As String * 1
    CRYINDOI As String * 1
    CRYRESOI As String * 1
    CRYINDB1 As String * 1
    CRYRESB1 As String * 1
    CRYINDB2 As String * 1
    CRYRESB2 As String * 1
    CRYINDB3 As String * 1
    CRYRESB3 As String * 1
    CRYINDL1 As String * 1
    CRYRESL1 As String * 1
    CRYINDL2 As String * 1
    CRYRESL2 As String * 1
    CRYINDL3 As String * 1
    CRYRESL3 As String * 1
    CRYINDL4 As String * 1
    CRYRESL4 As String * 1
    CRYINDCS As String * 1
    CRYRESCS As String * 1
    CRYINDGD As String * 1
    CRYRESGD As String * 1
    CRYINDT As String * 1
    CRYREST As String * 1
    CRYINDEP As String * 1
    CRYRESEP As String * 1
    
    HSXCNHWS As String * 1          ' 品ＳＸ炭素濃度保証方法＿処
    HSXLTHWS As String * 1          ' 品ＳＸＬタイム保証方法＿処
    EPD As String * 1               ' EPD
End Type

#If SPEEDUP Then   '高速化実験 02.1.28-2.15 野村
Private Type tSmpMng
    BLOCKID As String * 12
    TOPPOS As Integer
    BOTPOS As Integer
    
    CRYNUM As String * 12
    IngotPos As Integer
    SMPKBN As String * 1
    
    HINBAN As String * 8            ' 品番
    REVNUM As Integer               ' 製品番号改訂番号
    factory As String * 1           ' 工場
    opecond As String * 1           ' 操業条件
    
    CRYINDRS As String * 1
    CRYRESRS As String * 1
    CRYINDOI As String * 1
    CRYRESOI As String * 1
    CRYINDB1 As String * 1
    CRYRESB1 As String * 1
    CRYINDB2 As String * 1
    CRYRESB2 As String * 1
    CRYINDB3 As String * 1
    CRYRESB3 As String * 1
    CRYINDL1 As String * 1
    CRYRESL1 As String * 1
    CRYINDL2 As String * 1
    CRYRESL2 As String * 1
    CRYINDL3 As String * 1
    CRYRESL3 As String * 1
    CRYINDL4 As String * 1
    CRYRESL4 As String * 1
    CRYINDCS As String * 1
    CRYRESCS As String * 1
    CRYINDGD As String * 1
    CRYRESGD As String * 1
    CRYINDT As String * 1
    CRYREST As String * 1
    CRYINDEP As String * 1
    CRYRESEP As String * 1
End Type
#End If


'待ち一覧

'初期表示用
Public Type type_DBDRV_scmzc_fcmkc001b_Disp
    CRYNUM As String * 12           ' 結晶番号
    IngotPos As Integer             ' 結晶内開始位置
'    LENGTH As Integer               ' 長さ              '2001/11/8
    BLOCKID As String * 12          ' ブロックID
    HSXTYPE As String * 1           ' 品ＳＸタイプ
    HSXCDIR As String * 1           ' 品ＳＸ結晶面方位
    UPDDATE As Date                 ' 更新日付
    Judg As String                  ' 判定
    hin() As tFullHinban            ' 品番(full)
    HOLDCLS As String * 1           ' ホールド区分 ---kuramoto 追加 2001/09/25----
    SMP() As type_cmkc001b_SmpMng   ' サンプル管理
End Type

'品番、仕様、結晶内側取得用 (TOP,TAIL順で２レコード取得)
Public Type type_DBDRV_scmzc_fcmkc001c_Siyou
    'ブロック管理
    CRYNUM As String * 12             ' 結晶番号
    IngotPos As Integer               ' 結晶内開始位置
    LENGTH As Integer                 ' 長さ
    '品番管理
    hin As tFullHinban                ' 品番(full)
        
        '結晶情報
    PRODCOND As String * 4            ' 製作条件
    PGID As String * 8                ' ＰＧ－ＩＤ
    UPLENGTH As Integer               ' 引上げ長さ
    FREELENG As Integer               ' フリー長
    DIAMETER As Integer               ' 直径 2002/05/01 S.Sano
    CHARGE As Double                  ' チャージ量
    SEED As String * 4                ' シード
    ADDDPPOS As Integer                 ' 追加ドープ位置

    '製品仕様
    HSXTYPE As String * 1             ' 品ＳＸタイプ
    HSXD1CEN As Double                ' 品ＳＸ直径１中心
    HSXCDIR As String * 1             ' 品ＳＸ結晶面方位
    HSXRMIN As Double                 ' 品ＳＸ比抵抗下限
    HSXRMAX As Double                 ' 品ＳＸ比抵抗上限
    HSXRAMIN As Double                ' 品ＳＸ比抵抗平均下限
    HSXRAMAX As Double                ' 品ＳＸ比抵抗平均上限
    HSXRMBNP As Double                ' 品ＳＸ比抵抗面内分布
    HSXRSPOH As String * 1            ' 品ＳＸ比抵抗測定位置＿方
    HSXRSPOT As String * 1            ' 品ＳＸ比抵抗測定位置＿点
    HSXRSPOI As String * 1            ' 品ＳＸ比抵抗測定位置＿位
    HSXRHWYT As String * 1            ' 品ＳＸ比抵抗保証方法＿対
    HSXRHWYS As String * 1            ' 品ＳＸ比抵抗保証方法＿処

    HSXONMIN As Double                ' 品ＳＸ酸素濃度下限
    HSXONMAX As Double                ' 品ＳＸ酸素濃度上限
    HSXONAMN As Double                ' 品ＳＸ酸素濃度平均下限
    HSXONAMX As Double                ' 品ＳＸ酸素濃度平均上限
    HSXONMBP As Double                ' 品ＳＸ酸素濃度面内分布
    HSXONSPH As String * 1            ' 品ＳＸ酸素濃度測定位置＿方
    HSXONSPT As String * 1            ' 品ＳＸ酸素濃度測定位置＿点
    HSXONSPI As String * 1            ' 品ＳＸ酸素濃度測定位置＿位
    HSXONHWT As String * 1            ' 品ＳＸ酸素濃度保証方法＿対
    HSXONHWS As String * 1            ' 品ＳＸ酸素濃度保証方法＿処

    HSXBM1AN As Double                ' 品ＳＸＢＭＤ１平均下限
    HSXBM1AX As Double                ' 品ＳＸＢＭＤ１平均上限
    HSXBM2AN As Double                ' 品ＳＸＢＭＤ２平均下限
    HSXBM2AX As Double                ' 品ＳＸＢＭＤ２平均上限
    HSXBM3AN As Double                ' 品ＳＸＢＭＤ３平均下限
    HSXBM3AX As Double                ' 品ＳＸＢＭＤ３平均上限
    HSXBM1SH As String * 1            ' 品ＳＸＢＭＤ１測定位置＿方
    HSXBM1ST As String * 1            ' 品ＳＸＢＭＤ１測定位置＿点
    HSXBM1SR As String * 1            ' 品ＳＸＢＭＤ１測定位置＿領
    HSXBM1HT As String * 1            ' 品ＳＸＢＭＤ１保証方法＿対
    HSXBM1HS As String * 1            ' 品ＳＸＢＭＤ１保証方法＿処
    HSXBM2SH As String * 1            ' 品ＳＸＢＭＤ２測定位置＿方
    HSXBM2ST As String * 1            ' 品ＳＸＢＭＤ２測定位置＿点
    HSXBM2SR As String * 1            ' 品ＳＸＢＭＤ２測定位置＿領
    HSXBM2HT As String * 1            ' 品ＳＸＢＭＤ２保証方法＿対
    HSXBM2HS As String * 1            ' 品ＳＸＢＭＤ２保証方法＿処
    HSXBM3SH As String * 1            ' 品ＳＸＢＭＤ３測定位置＿方
    HSXBM3ST As String * 1            ' 品ＳＸＢＭＤ３測定位置＿点
    HSXBM3SR As String * 1            ' 品ＳＸＢＭＤ３測定位置＿領
    HSXBM3HT As String * 1            ' 品ＳＸＢＭＤ３保証方法＿対
    HSXBM3HS As String * 1            ' 品ＳＸＢＭＤ３保証方法＿処

    HSXOS1AX As Double                ' 品ＳＸＯＳＦ１平均上限
    HSXOS1MX As Double                ' 品ＳＸＯＳＦ１上限
    HSXOS2AX As Double                ' 品ＳＸＯＳＦ２平均上限
    HSXOS2MX As Double                ' 品ＳＸＯＳＦ２上限
    HSXOS3AX As Double                ' 品ＳＸＯＳＦ３平均上限
    HSXOS3MX As Double                ' 品ＳＸＯＳＦ３上限
    HSXOS4AX As Double                ' 品ＳＸＯＳＦ４平均上限
    HSXOS4MX As Double                ' 品ＳＸＯＳＦ４上限
    HSXOS1SH As String * 1            ' 品ＳＸＯＳＦ１測定位置＿方
    HSXOS1ST As String * 1            ' 品ＳＸＯＳＦ１測定位置＿点
    HSXOS1SR As String * 1            ' 品ＳＸＯＳＦ１測定位置＿領
    HSXOS1HT As String * 1            ' 品ＳＸＯＳＦ１保証方法＿対
    HSXOS1HS As String * 1            ' 品ＳＸＯＳＦ１保証方法＿処
    HSXOS2SH As String * 1            ' 品ＳＸＯＳＦ２測定位置＿方
    HSXOS2ST As String * 1            ' 品ＳＸＯＳＦ２測定位置＿点
    HSXOS2SR As String * 1            ' 品ＳＸＯＳＦ２測定位置＿領
    HSXOS2HT As String * 1            ' 品ＳＸＯＳＦ２保証方法＿対
    HSXOS2HS As String * 1            ' 品ＳＸＯＳＦ２保証方法＿処
    HSXOS3SH As String * 1            ' 品ＳＸＯＳＦ３測定位置＿方
    HSXOS3ST As String * 1            ' 品ＳＸＯＳＦ３測定位置＿点
    HSXOS3SR As String * 1            ' 品ＳＸＯＳＦ３測定位置＿領
    HSXOS3HT As String * 1            ' 品ＳＸＯＳＦ３保証方法＿対
    HSXOS3HS As String * 1            ' 品ＳＸＯＳＦ３保証方法＿処
    HSXOS4SH As String * 1            ' 品ＳＸＯＳＦ４測定位置＿方
    HSXOS4ST As String * 1            ' 品ＳＸＯＳＦ４測定位置＿点
    HSXOS4SR As String * 1            ' 品ＳＸＯＳＦ４測定位置＿領
    HSXOS4HT As String * 1            ' 品ＳＸＯＳＦ４保証方法＿対
    HSXOS4HS As String * 1            ' 品ＳＸＯＳＦ４保証方法＿処
    HSXOS1NS As String * 2            ' 品ＳＸＯＳＦ１熱処理法
    HSXOS2NS As String * 2            ' 品ＳＸＯＳＦ２熱処理法
    HSXOS3NS As String * 2            ' 品ＳＸＯＳＦ３熱処理法
    HSXOS4NS As String * 2            ' 品ＳＸＯＳＦ４熱処理法
    HSXBM1NS As String * 2            ' 品ＳＸＢＭＤ１熱処理法
    HSXBM2NS As String * 2            ' 品ＳＸＢＭＤ２熱処理法
    HSXBM3NS As String * 2            ' 品ＳＸＢＭＤ３熱処理法

    HSXCNMIN As Double                ' 品ＳＸ炭素濃度下限
    HSXCNMAX As Double                ' 品ＳＸ炭素濃度上限
    HSXCNSPH As String * 1            ' 品ＳＸ炭素濃度測定位置＿方
    HSXCNSPT As String * 1            ' 品ＳＸ炭素濃度測定位置＿点
    HSXCNSPI As String * 1            ' 品ＳＸ炭素濃度測定位置＿位
    HSXCNHWT As String * 1            ' 品ＳＸ炭素濃度保証方法＿対
    HSXCNHWS As String * 1            ' 品ＳＸ炭素濃度保証方法＿処

    HSXDENMX As Integer               ' 品ＳＸＤｅｎ上限
    HSXDENMN As Integer               ' 品ＳＸＤｅｎ下限
    HSXLDLMX As Integer               ' 品ＳＸＬ／ＤＬ上限
    HSXLDLMN As Integer               ' 品ＳＸＬ／ＤＬ下限
    HSXDVDMX As Integer               ' 品ＳＸＤＶＤ２上限
    HSXDVDMN As Integer               ' 品ＳＸＤＶＤ２下限
    HSXDENHT As String * 1            ' 品ＳＸＤｅｎ保証方法＿対
    HSXDENHS As String * 1            ' 品ＳＸＤｅｎ保証方法＿処
    HSXLDLHT As String * 1            ' 品ＳＸＬ／ＤＬ保証方法＿対
    HSXLDLHS As String * 1            ' 品ＳＸＬ／ＤＬ保証方法＿処
    HSXDVDHT As String * 1            ' 品ＳＸＤＶＤ２保証方法＿対
    HSXDVDHS As String * 1            ' 品ＳＸＤＶＤ２保証方法＿処
    HSXDENKU As String * 1            ' 品ＳＸＤｅｎ検査有無
    HSXDVDKU As String * 1            ' 品ＳＸＤＶＤ２検査有無
    HSXLDLKU As String * 1            ' 品ＳＸＬ／ＤＬ検査有無

    HSXLTMIN As Integer               ' 品ＳＸＬタイム下限
    HSXLTMAX As Integer               ' 品ＳＸＬタイム上限
    HSXLTSPH As String * 1            ' 品ＳＸＬタイム測定位置＿方
    HSXLTSPT As String * 1            ' 品ＳＸＬタイム測定位置＿点
    HSXLTSPI As String * 1            ' 品ＳＸＬタイム測定位置＿位
    HSXLTHWT As String * 1            ' 品ＳＸＬタイム保証方法＿対
    HSXLTHWS As String * 1            ' 品ＳＸＬタイム保証方法＿処
    '結晶内側管理
    EPDUP As Integer                  ' EPD　上限
End Type


' 結晶サンプル管理取得用 (TOP,TAIL順で２レコード取得)
Public Type type_DBDRV_scmzc_fcmkc001c_CrySmp
    CRYNUM As String * 12             ' 結晶番号
    IngotPos As Integer               ' 結晶内位置
    LENGTH As Integer                 ' 長さ
    BLOCKID As String * 12            ' ブロックID
    SMPKBN As String * 1              ' サンプル区分
    SMPLNO As Integer                 ' サンプルNo
    HINBAN As String * 12             ' 品番
    REVNUM As Integer                 ' 製品番号改訂番号
    factory As String * 1             ' 工場
    opecond As String * 1             ' 操業条件
    KTKBN  As String * 1              ' 確定区分
    CRYINDRS As String * 1            ' 結晶検査指示（Rs)
    CRYINDOI As String * 1            ' 結晶検査指示（Oi)
    CRYINDB1 As String * 1            ' 結晶検査指示（B1)
    CRYINDB2 As String * 1            ' 結晶検査指示（B2）
    CRYINDB3 As String * 1            ' 結晶検査指示（B3)
    CRYINDL1 As String * 1            ' 結晶検査指示（L1)
    CRYINDL2 As String * 1            ' 結晶検査指示（L2)
    CRYINDL3 As String * 1            ' 結晶検査指示（L3)
    CRYINDL4 As String * 1            ' 結晶検査指示（L4)
    CRYINDCS As String * 1            ' 結晶検査指示（Cs)
    CRYINDGD As String * 1            ' 結晶検査指示（GD)
    CRYINDT As String * 1             ' 結晶検査指示（T)
    CRYINDEP As String * 1            ' 結晶検査指示（EPD)
End Type


'結晶抵抗実績
Public Type type_DBDRV_scmzc_fcmkc001c_CryR
    CRYNUM As String * 12             ' 結晶番号
    POSITION As Integer               ' 位置
    SMPKBN As String * 1              ' サンプル区分
    SMPLNO As Integer                 ' サンプルＮｏ
    SMPLUMU As String * 1             ' サンプル有無
    TRANCOND As String * 1            ' 処理条件
    MEAS1 As Double                   ' 測定値１
    MEAS2 As Double                   ' 測定値２
    MEAS3 As Double                   ' 測定値３
    MEAS4 As Double                   ' 測定値４
    MEAS5 As Double                   ' 測定値５
    RRG As Double                     ' ＲＲＧ
    REGDATE As Date                   ' 登録日付
End Type


'Oi実績
Public Type type_DBDRV_scmzc_fcmkc001c_Oi
    CRYNUM As String * 12             ' 結晶番号
    POSITION As Integer               ' 位置
    SMPKBN As String * 1              ' サンプル区分
    SMPLNO As Integer                 ' サンプルＮｏ
    SMPLUMU As String * 1             ' サンプル有無
    TRANCOND As String * 1            ' 処理条件
    OIMEAS1 As Double                 ' Ｏｉ測定値１
    OIMEAS2 As Double                 ' Ｏｉ測定値２
    OIMEAS3 As Double                 ' Ｏｉ測定値３
    OIMEAS4 As Double                 ' Ｏｉ測定値４
    OIMEAS5 As Double                 ' Ｏｉ測定値５
    ORGRES As Double                  ' ＯＲＧ結果
    AVE As Double                     ' ＡＶＥ
    FTIRCONV As Double                ' ＦＴＩＲ換算
    INSPECTWAY As String * 2          ' 検査方法
    REGDATE As Date                   ' 登録日付
End Type


'BMD1～3実績
Public Type type_DBDRV_scmzc_fcmkc001c_BMD
    CRYNUM As String * 12             ' 結晶番号
    POSITION As Integer               ' 位置
    SMPKBN As String * 1              ' サンプル区分
    SMPLNO As Integer                 ' サンプルＮｏ
    SMPLUMU As String * 1             ' サンプル有無
    HTPRC As String * 2               ' 熱処理方法
    KKSP As String * 3                ' 結晶欠陥測定位置
    KKSET As String * 3               ' 結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    TRANCOND As String * 1            ' 処理条件
    MEAS1 As Double                   ' 測定値１
    MEAS2 As Double                   ' 測定値２
    MEAS3 As Double                   ' 測定値３
    MEAS4 As Double                   ' 測定値４
    MEAS5 As Double                   ' 測定値５
    Min As Double                     ' MIN
    max As Double                     ' MAX
    AVE As Double                     ' AVE
    REGDATE As Date                   ' 登録日付
End Type


'OSF1～4実績
Public Type type_DBDRV_scmzc_fcmkc001c_OSF
    CRYNUM As String * 12             ' 結晶番号
    POSITION As Integer               ' 位置
    SMPKBN As String * 1              ' サンプル区分
    SMPLNO As Integer                 ' サンプルＮｏ
    SMPLUMU As String * 1             ' サンプル有無
    HTPRC As String * 2               ' 熱処理方法
    KKSP As String * 3                ' 結晶欠陥測定位置
    KKSET As String * 3               ' 結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    TRANCOND As String * 1            ' 処理条件
    CALCMAX As Double                 ' 計算結果 Max
    CALCAVE As Double                 ' 計算結果 Ave
    MEAS1 As Double                   ' 測定値１
    MEAS2 As Double                   ' 測定値２
    MEAS3 As Double                   ' 測定値３
    MEAS4 As Double                   ' 測定値４
    MEAS5 As Double                   ' 測定値５
    MEAS6 As Double                   ' 測定値６
    MEAS7 As Double                   ' 測定値７
    MEAS8 As Double                   ' 測定値８
    MEAS9 As Double                   ' 測定値９
    MEAS10 As Double                  ' 測定値１０
    MEAS11 As Double                  ' 測定値１１
    MEAS12 As Double                  ' 測定値１２
    MEAS13 As Double                  ' 測定値１３
    MEAS14 As Double                  ' 測定値１４
    MEAS15 As Double                  ' 測定値１５
    MEAS16 As Double                  ' 測定値１６
    MEAS17 As Double                  ' 測定値１７
    MEAS18 As Double                  ' 測定値１８
    MEAS19 As Double                  ' 測定値１９
    MEAS20 As Double                  ' 測定値２０
    REGDATE As Date                   ' 登録日付
End Type


'CS実績
Public Type type_DBDRV_scmzc_fcmkc001c_CS
    CRYNUM As String * 12             ' 結晶番号
    POSITION As Integer               ' 位置
    SMPKBN As String * 1              ' サンプル区分
    SMPLNO As Integer                 ' サンプルＮｏ
    SMPLUMU As String * 1             ' サンプル有無
    TRANCOND As String * 1            ' 処理条件
    CSMEAS As Double                  ' Cs実測値
    PRE70P As Double                  ' ７０％推定値
    REGDATE As Date                   ' 登録日付
End Type


'GD実績
Public Type type_DBDRV_scmzc_fcmkc001c_GD
    CRYNUM As String * 12             ' 結晶番号
    POSITION As Integer               ' 位置
    SMPKBN As String * 1              ' サンプル区分
    SMPLNO As Integer                 ' サンプルＮｏ
    SMPLUMU As String * 1             ' サンプル有無
    TRANCOND As String * 1            ' 処理条件
    MSRSDEN As Integer                ' 測定結果 Den
    MSRSLDL As Integer                ' 測定結果 L/DL
    MSRSDVD2 As Integer               ' 測定結果 DVD2
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
    REGDATE As Date                   ' 登録日付
End Type


'ライフタイム実績取得関数
Public Type type_DBDRV_scmzc_fcmkc001c_LT
    CRYNUM As String * 12             ' 結晶番号
    POSITION As Integer               ' 位置
    SMPKBN As String * 1              ' サンプル区分
    SMPLNO As Integer                 ' サンプルＮｏ
    SMPLUMU As String * 1             ' サンプル有無
    MEAS1 As Integer                  ' 測定値１
    MEAS2 As Integer                  ' 測定値２
    MEAS3 As Integer                  ' 測定値３
    MEAS4 As Integer                  ' 測定値４
    MEAS5 As Integer                  ' 測定値５
    TRANCOND As String * 1            ' 処理条件
    MEASPEAK As Integer               ' 測定値 ピーク値
    CALCMEAS As Integer               ' 計算結果
    REGDATE As Date                   ' 登録日付
    LTSPI As String                 '測定位置コード
End Type


'EPD実績取得関数
Public Type type_DBDRV_scmzc_fcmkc001c_EPD
    CRYNUM As String * 12             ' 結晶番号
    POSITION As Integer               ' 位置
    SMPKBN As String * 1              ' サンプル区分
    SMPLNO As Integer                 ' サンプルＮｏ
    SMPLUMU As String * 1             ' サンプル有無
    TRANCOND As String * 1            ' 処理条件
    MEASURE As Integer                ' 測定値
    REGDATE As Date                   ' 登録日付
End Type


'実績をまとめた構造体
Public Type type_DBDRV_scmzc_fcmkc001c_Zisseki
    CRYRZ() As type_DBDRV_scmzc_fcmkc001c_CryR
    OIZ() As type_DBDRV_scmzc_fcmkc001c_Oi
    BMD1Z() As type_DBDRV_scmzc_fcmkc001c_BMD
    BMD2Z() As type_DBDRV_scmzc_fcmkc001c_BMD
    BMD3Z() As type_DBDRV_scmzc_fcmkc001c_BMD
    OSF1Z() As type_DBDRV_scmzc_fcmkc001c_OSF
    OSF2Z() As type_DBDRV_scmzc_fcmkc001c_OSF
    OSF3Z() As type_DBDRV_scmzc_fcmkc001c_OSF
    OSF4Z() As type_DBDRV_scmzc_fcmkc001c_OSF
    csz() As type_DBDRV_scmzc_fcmkc001c_CS
    GDZ() As type_DBDRV_scmzc_fcmkc001c_GD
    LTZ() As type_DBDRV_scmzc_fcmkc001c_LT
    EPDZ() As type_DBDRV_scmzc_fcmkc001c_EPD
    SURSZ() As type_DBDRV_scmzc_fcmkc001c_CryR
End Type


'ブロック管理更新用（現在工程、最終通過工程）
Public Type type_DBDRV_scmzc_fcmkc001c_UpdBlock1
    CRYNUM As String * 12           ' 結晶番号
    IngotPos As Integer             ' 結晶内開始位置
    NOWPROC As String * 5           ' 現在工程
    LASTPASS As String * 5          ' 最終通過工程
End Type


'ブロック管理更新用（削除区分、最終状態区分、流動状態区分）
Public Type type_DBDRV_scmzc_fcmkc001c_UpdBlock2
    CRYNUM As String * 12           ' 結晶番号
    IngotPos As Integer             ' 結晶内開始位置
    DELCLS As String * 1            ' 削除区分
    LSTATCLS As String * 1          ' 最終状態区分
    RSTATCLS As String * 1          ' 流動状態区分
End Type

'ブロック管理更新用（クリスタルカタログ、リメルト用）
Public Type typ_DBDRV_fcmkc001c_UpdBlkCR
    CRYNUM As String * 12           ' 結晶番号
    IngotPos As Integer             ' 結晶内開始位置
    NOWPROC As String * 5           ' 現在工程
'    LASTPASS As String * 5          ' 最終通過工程
    DELCLS As String * 1            ' 削除区分
    BDCAUS As String * 3            ' 不良理由
    LSTATCLS As String * 1          ' 最終状態区分
    RSTATCLS As String * 1          ' 流動状態区分
End Type



'結晶サンプル管理更新用
Public Type type_DBDRV_scmzc_fcmkc001c_UpdCrySmp
    CRYNUM As String * 12           ' 結晶番号
    IngotPos As Integer             ' 結晶内位置
    SMPKBN As String * 1            ' サンプル区分
End Type

'測定結果のJ014書込要否構造体
Public Type Judg_Spec_Cry
    Enable As Boolean           '有効な品番である
    rs As Boolean               'Rsは要書込
    Oi As Boolean               'Oiは要書込
    B1 As Boolean               'BMD1は要書込
    B2 As Boolean               'BMD2は要書込
    B3 As Boolean               'BMD3は要書込
    L1 As Boolean               'OSF1は要書込
    L2 As Boolean               'OSF2は要書込
    L3 As Boolean               'OSF3は要書込
    L4 As Boolean               'OSF4は要書込
    Cs As Boolean               'Csは要書込
    GD As Boolean               'GDは要書込
    Lt As Boolean               'LTは要書込
    EPD As Boolean              'EPDは要書込
End Type

' 仕様の指示がたっている判断用
Public Const SIJI = "H"
Public Const SANKOU = "S"

'2002/08/01 M.Tomita------------------------------------------------------

'===========================================
' ＷＦ加工用共通テーブル
'===========================================

' 抜試指示
Public Type typ_WafInd
    BLOCKID As String * 12      ' ブロックID
    BlockPos As Integer         ' ブロックＰ
    IngotPos As Integer         ' 結晶Ｐ
    LENGTH As Integer           ' 長さ
    HINUP As tFullHinban        ' 上品番
    HINDN As tFullHinban        ' 下品番
    SMP As typ_WFSample         ' 検査項目
    HINFLG As Boolean           ' 品番区切りフラグ
    SMPFLG As Boolean           ' WFサンプル区切りフラグ
    ERRDNFLG As Boolean         ' 下品番エラーフラグ
    SMPLKBN1 As String * 1      ' サンプル区分１
    SMPLKBN2 As String * 1      ' サンプル区分２
End Type

' 製品仕様
Public Type typ_HinSpec
    hin As tFullHinban          ' 品番
    IngotPos As Integer         ' 結晶内開始位置
    LENGTH As Integer           ' 長さ
    HWFRMIN As Double           ' 比抵抗下限
    HWFRMAX As Double           ' 比抵抗上限
    HWFRHWYS As String * 1      ' 検査有無(Rs)
    HWFONHWS As String * 1      ' 検査有無(Oi)
    HWFBM1HS As String * 1      ' 検査有無(B1)
    HWFBM2HS As String * 1      ' 検査有無(B2)
    HWFBM3HS As String * 1      ' 検査有無(B3)
    HWFOF1HS As String * 1      ' 検査有無(L1)
    HWFOF2HS As String * 1      ' 検査有無(L2)
    HWFOF3HS As String * 1      ' 検査有無(L3)
    HWFOF4HS As String * 1      ' 検査有無(L4)
    HWFDSOHS As String * 1      ' 検査有無(DS)
    HWFMKHWS As String * 1      ' 検査有無(DZ)
    HWFSPVHS As String * 1      ' 検査有無(SP/Fe濃度)
    HWFDLHWS As String * 1      ' 検査有無(SP/拡散長)
    HWFOS1HS As String * 1      ' 検査有無(D1)
    HWFOS2HS As String * 1      ' 検査有無(D2)
    HWFOS3HS As String * 1      ' 検査有無(D3)
    HWFOTHER1 As String * 1     ' 検査有無(OT2) ''Add.03/05/20 後藤
    HWFOTHER2 As String * 1     ' 検査有無(OT1) ''Add.03/05/20
End Type

' 欠落ウェハー
Public Type typ_LackMap
    BLOCKID As String * 12      ' ブロックID
    LACKPOSS As Double          ' 欠落位置(From)
    LACKPOSE As Double          ' 欠落位置(To)
    REJCAT As String * 1        ' 欠落理由
    LACKCNTS As Integer         ' 欠落枚目(From)
    LACKCNTE As Integer         ' 欠落枚目(To)
End Type

'各実績情報
Public Type typ_ALLRSLT
    pos As Integer                    ' 結晶内開始位置
    NAIYO As String                   ' 内容
    INFO1 As String                   ' 情報１
    INFO2 As String                   ' 情報２
    INFO3 As String                   ' 情報３
    INFO4 As String                   ' 情報４
    OKNG  As String                   ' 判定結果
    SMPLNO As Integer                 ' サンプルＮｏ
    BLOCKNG As Boolean                'GDエラーとなる品番を含むか判別
End Type

'全情報構造体
Type typ_AllTypes
    intPFlg As Integer                              ' 表示フラグ
    strStaffID As String                            ' スタッフID
    strStaffName As String                          ' スタッフ名
    BLOCKID  As String * 12                         ' ブロックID
    Cut(2) As Double                                ' 再カット位置
    COEF(2) As Double                               ' 偏析係数
    CRCOEF As Double                                ' 結晶偏析係数
    OKNG(2) As Boolean                              ' 比抵抗判定
    Henseki As Boolean                              ' 比抵抗実績有無(結晶全体TOP/TAIL)
    JudgRes(2) As Boolean                              ' 比抵抗判定    2001/10/02 S.Sano
    JudgRrg(2) As Boolean                              ' RRG判定       2001/10/02 S.Sano
    typ_rsz() As typ_TBCMJ002                       ' 結晶抵抗実績(結晶全体TOP/TAIL)
    typ_hage(2) As typ_TBCMH004                     ' 引上げ終了実績
    typ_rslt(2, MAXCNT) As typ_ALLRSLT              ' 各実績情報
    typ_zi As type_DBDRV_scmzc_fcmkc001c_Zisseki    ' 実績をまとめた構造体
    typ_si() As type_DBDRV_scmzc_fcmkc001c_Siyou    ' 仕様
    typ_cr() As type_DBDRV_scmzc_fcmkc001c_CrySmp   ' 結晶サンプル管理取得用 (TOP,TAIL順で２レコード取得)
    blYONE As Boolean                               ' 米沢フラグ
End Type

Public typ_A As typ_AllTypes        '全情報構造体
Public JudgSC(2) As Judg_Spec_Cry        '仕様検査支持構造体
Public TotalJudg As Boolean         'トータル判定
Public MeasFlag(2) As Judg_Spec_Cry        '仕様検査支持構造体
Public Kakou As type_KakouJudg      '加工実績判定構造体


'ブロックラベル払出し  4/16 Yam作成

' ブロック一覧
Public Type typ_BlkLbl
    BLOCKID As String * 12      ' ブロックID
    hin(5) As tFullHinban       ' 品番
    WFINDDATE As String * 10    ' 最終抜試日付
    CRYNUM As String * 12       ' 結晶番号
    IngotPos As Integer         ' インゴット内位置
    LENGTH As Integer           ' ブロック長さ
    REALLEN As Integer          ' ブロック実長さ
    HINLEN(5) As Integer        ' 品番長さ
    DIAMETER As Integer         ' 直径
    SBLOCKID As String * 12     ' 先頭ブロックID
    BLOCKORDER As Integer       ' ブロック順序
    HOLDCLS As String * 1       ' ホールド状態  --- 2001/09/19 kuramoto 追加 ---
    PASSFLAG As String * 1      ' 通過フラグ　　--- 200/04/16 Yam
End Type



'概要      :抜試指示用 画面表示時ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:sBlockID　　　,I  ,String         　,ブロックID
'      　　:pCryInf 　　　,O  ,typ_TBCME037   　,結晶情報
'      　　:pHinDsn 　　　,O  ,typ_TBCME039   　,品番設計
'      　　:pHinMng 　　　,O  ,typ_TBCME041   　,品番管理
'      　　:pBlkInf 　　　,O  ,typ_BlkInf1    　,ブロック情報
'      　　:pHinSpec　　　,O  ,typ_HinSpec    　,製品仕様
'      　　:dNeraiRes 　　,O  ,Double         　,ねらい品番の比抵抗上限値（P+の判断用）
'      　　:sErrMsg 　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,読み込みの成否
Public Function DBDRV_scmzc_fcmkc001g_Disp(ByVal SBLOCKID As String, pCryInf As typ_TBCME037, _
                                           pHinDsn() As typ_TBCME039, pHinMng() As typ_TBCME041, _
                                           pBlkInf() As typ_BlkInf1, pHinSpec() As typ_HinSpec, _
                                           dNeraiRes As Double, sErrMsg As String) As FUNCTION_RETURN

    Dim tmpCryInf() As typ_TBCME037
    Dim rs As OraDynaset
    Dim sql As String
    Dim sDbName As String
    Dim sCryNum As String
    Dim sHin As String
    Dim sSeed As String
    Dim dMenseki As Double
    Dim dTopWght As Double
    Dim dCharge As Double
    Dim dMeas(4) As Double
    Dim bFlag As Boolean
    Dim recCnt As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001g_SQL.bas -- Function DBDRV_scmzc_fcmkc001g_Disp"
    sErrMsg = ""

    '' ブロック管理の取得
    sDbName = "E040"
    sCryNum = Left(SBLOCKID, 9) & "000"
    sql = "select "
    sql = sql & "INGOTPOS, LENGTH, REALLEN, BLOCKID, "
    sql = sql & "KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, RSTATCLS"
    sql = sql & " from TBCME040 where CRYNUM='" & sCryNum & "'"
    sql = sql & " and INGOTPOS>=0 and LENGTH>0 order by INGOTPOS"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    bFlag = False
    ReDim pBlkInf(recCnt)
    For i = 1 To recCnt
        With pBlkInf(i)
            .COF.TOPSMPLPOS = rs("INGOTPOS")
            .LENGTH = rs("LENGTH")
            .REALLEN = rs("REALLEN")
            .BLOCKID = rs("BLOCKID")
            .KRPROCCD = rs("KRPROCCD")
            .NOWPROC = rs("NOWPROC")
            .LPKRPROCCD = rs("LPKRPROCCD")
            .LASTPASS = rs("LASTPASS")
            .RSTATCLS = rs("RSTATCLS")
            .COF.BOTSMPLPOS = .COF.TOPSMPLPOS + .LENGTH
            .SAMPFLAG = False
            If .BLOCKID = SBLOCKID Then
                bFlag = True
            End If
        End With
        rs.MoveNext
    Next i
    rs.Close

    '' ブロックID存在チェック
    If bFlag = False Then
        sErrMsg = GetMsgStr("EBLK0")
        DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' 結晶情報の取得(s_cmzcTBCME037_SQL.bas が必要)
    sDbName = "E037"
    sql = " where CRYNUM='" & sCryNum & "'"
    If DBDRV_GetTBCME037(tmpCryInf(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpCryInf) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pCryInf = tmpCryInf(1)

    '' 品番設計の取得(s_cmzcTBCME039_SQL.bas が必要)
    sDbName = "E039"
    sql = " where substr(CRYNUM,1,7)='" & Left(sCryNum, 7) & "' and LENGTH>0 order by INGOTPOS"
    If DBDRV_GetTBCME039(pHinDsn(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' 品番管理の取得(s_cmzcTBCME041_SQL.bas が必要)
    sDbName = "E041"
    sql = " where CRYNUM='" & sCryNum & "'　and LENGTH>0 order by INGOTPOS"
    If DBDRV_GetTBCME041(pHinMng(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(pHinMng) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' 引上げ終了実績の取得
    sDbName = "H004"
    sql = "select (DM1+DM2+DM3)/3.0 as DM, WGHTTOP, CHARGE, SEED from TBCMH004 where CRYNUM='" & sCryNum & "'"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount > 0 Then
        dMenseki = AreaOfCircle(rs("DM"))
        dTopWght = rs("WGHTTOP")
        dCharge = rs("CHARGE")
        sSeed = rs("SEED")
    Else
        dMenseki = 0
        dTopWght = 0
        dCharge = 0
        sSeed = ""
    End If
    rs.Close

    '' 結晶抵抗実績の取得
    sDbName = "J002"
    recCnt = UBound(pBlkInf)
    For i = 1 To recCnt
        With pBlkInf(i)
            .SEED = sSeed                   ' シード
            .COF.DUNMENSEKI = dMenseki      ' 断面積
            .COF.CHARGEWEIGHT = dCharge     ' チャージ量
            .COF.TOPWEIGHT = dTopWght       ' トップ重量

            '' トップ側比抵抗中央値の取得
            sql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 from TBCMJ002"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and POSITION=" & .COF.TOPSMPLPOS & " and SMPKBN='T'"
            sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT) from TBCMJ002"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and POSITION=" & .COF.TOPSMPLPOS & " and SMPKBN='T')"
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount > 0 Then
                dMeas(0) = rs("MEAS1")
                dMeas(1) = rs("MEAS2")
                dMeas(2) = rs("MEAS3")
                dMeas(3) = rs("MEAS4")
                dMeas(4) = rs("MEAS5")
                .COF.TOPRES = JudgCenter(dMeas())
            Else
                .COF.TOPRES = -9999
            End If
            rs.Close

            '' ボトム側比抵抗中央値の取得
            sql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 from TBCMJ002"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='B'"
            sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT) from TBCMJ002"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='B')"
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                rs.Close
                sql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 from TBCMJ002"
                sql = sql & " where CRYNUM='" & sCryNum & "'"
                sql = sql & " and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='T'"
                sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT) from TBCMJ002"
                sql = sql & " where CRYNUM='" & sCryNum & "'"
                sql = sql & " and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='T')"
                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            End If
            If rs.RecordCount > 0 Then
                dMeas(0) = rs("MEAS1")
                dMeas(1) = rs("MEAS2")
                dMeas(2) = rs("MEAS3")
                dMeas(3) = rs("MEAS4")
                dMeas(4) = rs("MEAS5")
                .COF.BOTRES = JudgCenter(dMeas())
            Else
                .COF.BOTRES = -9999
            End If
            rs.Close
        End With
    Next i

    '' 製品仕様の取得
    sDbName = "VE004"
    recCnt = UBound(pHinMng)
    ReDim pHinSpec(recCnt)
    k = 0
    For i = 1 To recCnt
        With pHinMng(i)
            sHin = RTrim$(.HINBAN)
            If sHin <> "" And sHin <> "G" And sHin <> "Z" Then
                For j = 1 To k
                    If pHinSpec(j).hin.HINBAN = .HINBAN Then
                        pHinSpec(j).LENGTH = pHinSpec(j).LENGTH + .LENGTH
                        Exit For
                    End If
                Next j
                If j > k Then
                    k = k + 1
                    pHinSpec(k).IngotPos = .IngotPos
                    pHinSpec(k).hin.HINBAN = .HINBAN
                    pHinSpec(k).hin.mnorevno = .REVNUM
                    pHinSpec(k).hin.factory = .factory
                    pHinSpec(k).hin.opecond = .opecond
                    pHinSpec(k).LENGTH = .LENGTH
                    If DBDRV_scmzc_fcmkc001g_GetSpec(pHinSpec(k)) = FUNCTION_RETURN_FAILURE Then
                        sErrMsg = GetMsgStr("EGET2", sDbName)
                        DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                End If
            End If
        End With
    Next i
    ReDim Preserve pHinSpec(k)

    '' ねらい品番の比抵抗上限値を取得
    sql = "select HSXRMAX"
    sql = sql & " from TBCME037 E37, TBCME018 E18"
    sql = sql & " where (E37.CRYNUM='" & Left$(SBLOCKID, 9) & "000')"
    sql = sql & " and (E37.RPHINBAN=E18.HINBAN) and (E37.RPREVNUM=E18.MNOREVNO)"
    sql = sql & " and (E37.RPFACT=E18.FACTORY) and (E37.RPOPCOND=E18.OPECOND)"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount > 0 Then
        dNeraiRes = rs("HSXRMAX")
    Else
        dNeraiRes = 0#      'ここまではこないはず
    End If
    rs.Close

    DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("EGET2", sDbName)
    DBDRV_scmzc_fcmkc001g_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :抜試指示用 製品仕様専用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:pHinSpec　　　,IO ,typ_HinSpec    　,製品仕様
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,読み込みの成否
Public Function DBDRV_scmzc_fcmkc001g_GetSpec(pHinSpec As typ_HinSpec) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String
    Dim sOT1    As String   '03/05/21
    Dim sOT2    As String
    Dim rtn     As FUNCTION_RETURN

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc033_SQL.bas -- Function DBDRV_scmzc_fcmkc001g_GetSpec"

    '' 製品仕様の取得
    With pHinSpec
        sql = "select "
        sql = sql & "E021HWFRMIN, E021HWFRMAX, E021HWFRHWYS, "
        sql = sql & "E024HWFMKHWS, E025HWFONHWS, E025HWFOS1HS, E025HWFOS2HS, E025HWFOS3HS, "
        sql = sql & "E026HWFDSOHS, E028HWFSPVHS, E028HWFDLHWS, E029HWFOF1HS, E029HWFOF2HS, "
        sql = sql & "E029HWFOF3HS, E029HWFOF4HS, E029HWFBM1HS, E029HWFBM2HS, E029HWFBM3HS"
        sql = sql & " from VECME004"
        sql = sql & " where E018HINBAN='" & .hin.HINBAN & "'"
        sql = sql & " and E018MNOREVNO=" & .hin.mnorevno
        sql = sql & " and E018FACTORY='" & .hin.factory & "'"
        sql = sql & " and E018OPECOND='" & .hin.opecond & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        .HWFRMIN = fncNullCheck(rs("E021HWFRMIN"))
        .HWFRMAX = fncNullCheck(rs("E021HWFRMAX"))
        .HWFRHWYS = rs("E021HWFRHWYS")
        .HWFMKHWS = rs("E024HWFMKHWS")
        .HWFONHWS = rs("E025HWFONHWS")
        .HWFOS1HS = rs("E025HWFOS1HS")
        .HWFOS2HS = rs("E025HWFOS2HS")
        .HWFOS3HS = rs("E025HWFOS3HS")
        .HWFDSOHS = rs("E026HWFDSOHS")
        .HWFSPVHS = rs("E028HWFSPVHS")
        .HWFDLHWS = rs("E028HWFDLHWS")
        .HWFOF1HS = rs("E029HWFOF1HS")
        .HWFOF2HS = rs("E029HWFOF2HS")
        .HWFOF3HS = rs("E029HWFOF3HS")
        .HWFOF4HS = rs("E029HWFOF4HS")
        .HWFBM1HS = rs("E029HWFBM1HS")
        .HWFBM2HS = rs("E029HWFBM2HS")
        .HWFBM3HS = rs("E029HWFBM3HS")
        rtn = scmzc_getE036(pHinSpec.hin, sOT1, sOT2)   '03/05/21
        If rtn = FUNCTION_RETURN_FAILURE Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        .HWFOTHER1 = sOT1 '### 03/05/21
        .HWFOTHER2 = sOT2
 
        rs.Close
    End With

    DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :抜試指示用 実行時ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:sCryNum　　　,I  ,String         　,結晶番号
'      　　:pBlkInf　　　,I  ,typ_BlkInf1    　,ブロック情報
'      　　:pSXLMng　　　,I  ,typ_TBCME042   　,SXL管理
'      　　:pWafSmp　　　,I  ,typ_XSDCW   　   ,新サンプル管理（SXL）
'      　　:pCryCat　　　,I  ,typ_TBCMG007   　,クリスタルカタログ受入実績
'      　　:pBsInd 　　　,I  ,typ_TBCMW001   　,抜試指示実績
'      　　:pMesInd　　　,I  ,typ_TBCMY003   　,測定評価方法指示
'      　　:sErrMsg　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,書き込みの成否
Public Function DBDRV_scmzc_fcmkc001g_Exec(ByVal sCryNum As String, pBlkInf() As typ_BlkInf1, _
                                           pSXLMng() As typ_TBCME042, pWafSmp() As typ_XSDCW, pCryCat() As typ_TBCMG007, _
                                           pBsInd() As typ_TBCMW001, pMesInd() As typ_TBCMY003, sErrMsg As String) As FUNCTION_RETURN

Dim sql As String
Dim sDbName As String
Dim recCnt As Long
Dim i As Long
Dim hin As tFullHinban

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc033_SQL.bas -- Function DBDRV_scmzc_fcmkc001g_Exec"
    sErrMsg = ""

    '' WriteDBLog " ", "Start"

    '' 結晶情報の更新
    sDbName = "E037"
    sql = "update TBCME037 set "
    sql = sql & "KRPROCCD='" & MGPRCD_WFC_HARAIDASI & "', "
    sql = sql & "PROCCD='" & PROCD_WFC_HARAIDASI & "', "
    sql = sql & "LPKRPROCCD='" & MGPRCD_NUKISI_SIJI & "', "
    sql = sql & "LASTPASS='" & PROCD_NUKISI_SIJI & "', "
    sql = sql & "UPDDATE=sysdate, "
    sql = sql & "SENDFLAG='0'"
    sql = sql & " where CRYNUM='" & sCryNum & "'"
    '' WriteDBLog sql, sDbName
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' ブロック管理の更新
    sDbName = "E040"
    recCnt = UBound(pBlkInf)
    For i = 1 To recCnt
        With pBlkInf(i)
            sql = "update TBCME040 set "
            sql = sql & "KRPROCCD='" & .KRPROCCD & "', "
            sql = sql & "NOWPROC='" & .NOWPROC & "', "
            sql = sql & "LPKRPROCCD='" & .LPKRPROCCD & "', "
            sql = sql & "LASTPASS='" & .LASTPASS & "', "
            sql = sql & "RSTATCLS='" & .RSTATCLS & "', "
            sql = sql & "UPDDATE=sysdate, "
            sql = sql & "SENDFLAG='0' "
            sql = sql & " where CRYNUM='" & sCryNum & "' and INGOTPOS=" & .COF.TOPSMPLPOS
        End With
        '' WriteDBLog sql, sDbName
        If OraDB.ExecuteSQL(sql) <= 0 Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    
    '品番管理テーブルの更新
    sDbName = "E041"
    recCnt = UBound(pBlkInf)
    With hin
        .mnorevno = 0
        .factory = " "
        .opecond = " "
    End With
    For i = 1 To recCnt
        With pBlkInf(i)
            If .RSTATCLS = "G" Then
                'G品番に変更
                hin.HINBAN = "G"
                If ChangeAreaHinban(sCryNum, CInt(.COF.TOPSMPLPOS), .LENGTH, hin) = FUNCTION_RETURN_FAILURE Then
                    DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
                End If
            ElseIf .RSTATCLS = "M" Then
                'Z品番に変更
                hin.HINBAN = "Z"
                If ChangeAreaHinban(sCryNum, CInt(.COF.TOPSMPLPOS), .LENGTH, hin) = FUNCTION_RETURN_FAILURE Then
                    DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
                End If
            End If
        End With
    Next

    '' SXL管理の挿入(s_cmzcDBdriverCOM_SQL.bas が必要)
    sDbName = "E042"
    If DBDRV_SXL_INS(pSXLMng()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' WFサンプル管理の挿入(s_cmzcDBdriverCOM_SQL.bas が必要)
    sDbName = "E044"
'''' --TEST--
''''If DBDRV_WfSmp_INS(pWafSmp()) = FUNCTION_RETURN_FAILURE Then
    If DBDRV_WfSmp_INS(pWafSmp(), i) = FUNCTION_RETURN_FAILURE Then
        
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' クリスタルカタログ受入実績の挿入
    sDbName = "G007"
    recCnt = UBound(pCryCat)
    For i = 1 To recCnt
        With pCryCat(i)
            sql = "insert into TBCMG007 "
            sql = sql & "(CRYNUM, TRANCNT, KRPROCCD, PROCCODE, BDCODE, PALTNUM, "
            sql = sql & "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE)"
            sql = sql & " select '"
            sql = sql & .CRYNUM & "', "             ' 結晶番号
            sql = sql & "nvl(max(TRANCNT),0)+1, '"  ' 処理回数
            sql = sql & MGPRCD_NUKISI_SIJI & "', '" ' 管理工程コード
            sql = sql & PROCD_NUKISI_SIJI & "', '"  ' 工程コード
            sql = sql & .BDCODE & "', '"            ' 不良理由コード
            sql = sql & .PALTNUM & "', '"           ' パレット番号
            sql = sql & .TSTAFFID & "', "           ' 登録社員ID
            sql = sql & "sysdate, '"                ' 登録日付
            sql = sql & .KSTAFFID & "', "           ' 更新社員ID
            sql = sql & "sysdate, "                 ' 更新日付
            sql = sql & "'0', "                     ' 送信フラグ
            sql = sql & "sysdate"                   ' 送信日付
            sql = sql & " from TBCMG007"
            sql = sql & " where CRYNUM='" & .CRYNUM & "'"
        End With
        '' WriteDBLog sql, sDbName
        If OraDB.ExecuteSQL(sql) <= 0 Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    '' 抜試指示実績の挿入
    sDbName = "W001"
    recCnt = UBound(pBsInd)
    For i = 1 To recCnt
        With pBsInd(i)
            sql = "insert into TBCMW001 "
            sql = sql & "(CRYNUM, INGOTPOS, TRANCNT, "
            sql = sql & "CRYLEN, KRPROCCD, PROCCODE, BLOCKID, "
            sql = sql & "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE)"
            sql = sql & " select '"
            sql = sql & .CRYNUM & "', "             ' 結晶番号
            sql = sql & .IngotPos & ", "            ' インゴット位置
            sql = sql & "nvl(max(TRANCNT),0)+1, "   ' 処理回数
            sql = sql & .CRYLEN & ", '"             ' 長さ
            sql = sql & MGPRCD_NUKISI_SIJI & "', '" ' 管理工程コード
            sql = sql & PROCD_NUKISI_SIJI & "', '"  ' 工程コード
            sql = sql & .BLOCKID & "', '"           ' ブロックID
            sql = sql & .TSTAFFID & "', "           ' 登録社員ID
            sql = sql & "sysdate, '"                ' 登録日付
            sql = sql & .TSTAFFID & "', "           ' 更新社員ID
            sql = sql & "sysdate, "                 ' 更新日付
            sql = sql & "'0', "                     ' 送信フラグ
            sql = sql & "sysdate"                   ' 送信日付
            sql = sql & " from TBCMW001"
            sql = sql & " where CRYNUM='" & sCryNum & "' and INGOTPOS=" & .IngotPos
        End With
        '' WriteDBLog sql, sDbName
        If OraDB.ExecuteSQL(sql) <= 0 Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    '' 測定評価方法指示の挿入(s_cmzcDBdriverCOM_SQL.bas が必要)
    sDbName = "Y003"
    If DBDRV_SokuSizi_Ins(pMesInd()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    '' WriteDBLog " ", "End"
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
    DBDRV_scmzc_fcmkc001g_Exec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCME037」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME037 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcF_TBCME037_SQL.basより移動)
Public Function DBDRV_GetTBCME037(records() As typ_TBCME037, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, DELCLS, KRPROCCD, PROCCD, LPKRPROCCD, LASTPASS, RPHINBAN, RPREVNUM, RPFACT, RPOPCOND, PRODCOND," & _
              " PGID, UPLENGTH, TOPLENG, BODYLENG, BOTLENG, FREELENG, DIAMETER, CHARGE, SEED, ADDDPCLS, ADDDPPOS, ADDDPVAL," & _
              " REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME037"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME037 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .DELCLS = rs("DELCLS")           ' 削除区分
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCD = rs("PROCCD")           ' 工程コード
            .LPKRPROCCD = rs("LPKRPROCCD")   ' 最終通過管理工程
            .LASTPASS = rs("LASTPASS")       ' 最終通過工程
            .RPHINBAN = rs("RPHINBAN")       ' ねらい品番
            .RPREVNUM = rs("RPREVNUM")       ' ねらい品番製品番号改訂番号
            .RPFACT = rs("RPFACT")           ' ねらい品番工場
            .RPOPCOND = rs("RPOPCOND")       ' ねらい品番操業条件
            .PRODCOND = rs("PRODCOND")       ' 製作条件
            .PGID = rs("PGID")               ' ＰＧ－ＩＤ
            .UPLENGTH = rs("UPLENGTH")       ' 引上げ長さ
            .TOPLENG = rs("TOPLENG")         ' ＴＯＰ長さ
            .BODYLENG = rs("BODYLENG")       ' 直胴長さ
            .BOTLENG = rs("BOTLENG")         ' ＢＯＴ長さ
            .FREELENG = rs("FREELENG")       ' フリー長
            .DIAMETER = rs("DIAMETER")       ' 直径
            .CHARGE = rs("CHARGE")           ' チャージ量
            .SEED = rs("SEED")               ' シード
            .ADDDPCLS = rs("ADDDPCLS")       ' 追加ドープ種類
            .ADDDPPOS = rs("ADDDPPOS")       ' 追加ドープ位置
            .ADDDPVAL = rs("ADDDPVAL")       ' 追加ドープ量
            .REGDATE = rs("REGDATE")         ' 登録日付
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME037 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCME039」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME039 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcF_TBCME039_SQL.basより移動)
Public Function DBDRV_GetTBCME039(records() As typ_TBCME039, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, INGOTPOS, HINBAN, REVNUM, FACT, OPCOND, LENGTH, USECLASS, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME039"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME039 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .IngotPos = rs("INGOTPOS")       ' 結晶内開始位置
            .HINBAN = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 改訂番号
            .FACT = rs("FACT")               ' 工場
            .OPCOND = rs("OPCOND")           ' 操業条件
            .LENGTH = rs("LENGTH")           ' 長さ
            .USECLASS = rs("USECLASS")       ' 使用区分
            .REGDATE = rs("REGDATE")         ' 登録日付
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME039 = FUNCTION_RETURN_SUCCESS
End Function


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
'履歴      :2001/08/24作成　野村  (2002/07 s_cmzcTBCME041_SQL.basより移動)
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
            .IngotPos = rs("INGOTPOS")       ' 結晶内開始位置
            .HINBAN = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
            .factory = rs("FACTORY")         ' 工場
            .opecond = rs("OPECOND")         ' 操業条件
            .LENGTH = rs("LENGTH")           ' 長さ
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


'概要    :待ち一覧 初期表示用ＤＢドライバ（検査待ち）
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO  ,型                                 ,説明
'        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,初期表示用
'        :戻ﾘ値        ,O   ,FUNCTION_RETURN                    ,読み込み成否
'説明    :
'履歴    :2001/07/06 蔵本 作成
Public Function DBDRV_scmzc_fcmkc001b_Disp1(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN

    Dim sql As String       'SQL全体
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      'ブロック管理のレコード数
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim BlockIdBuf As String
    
    '<検査待ち＞
    'ブロック管理テーブルからブロックID、更新日付取得（検査実績が未検査のもの）
    

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc033_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp1"

    DBDRV_scmzc_fcmkc001b_Disp1 = FUNCTION_RETURN_SUCCESS

    'ブロックID、更新日付の取得
    sql = "select distinct "
    sql = sql & " V.E040CRYNUM, "
    sql = sql & " V.E040INGOTPOS, "
    sql = sql & " V.E040BLOCKID, "
    sql = sql & " V.E040UPDDATE, "
    sql = sql & " V.E040HOLDCLS, "
    sql = sql & " H.HINBAN, "            ' 品番
    sql = sql & " H.REVNUM, "            ' 製品番号改訂番号
    sql = sql & " H.FACTORY, "           ' 工場
    sql = sql & " H.OPECOND, "           ' 操業条件
    sql = sql & " S.HSXTYPE, "           ' 品ＳＸタイプ
    sql = sql & " S.HSXCDIR, "            ' 品ＳＸ結晶面方位
    sql = sql & " H.INGOTPOS "
    sql = sql & " from "
    sql = sql & " VECME010 V, TBCME041 H, TBCME018 S "
    sql = sql & " where "
    sql = sql & " V.E040CRYNUM = H.CRYNUM "
    sql = sql & " and H.HINBAN = S.HINBAN "
    sql = sql & " and H.REVNUM = S.MNOREVNO "
    sql = sql & " and H.FACTORY = S.FACTORY "
    sql = sql & " and H.OPECOND = S.OPECOND "
                'ブロック内の品番検索
    sql = sql & " and (( V.E040INGOTPOS >= H.INGOTPOS "
    sql = sql & " and V.E040INGOTPOS < H.INGOTPOS + H.LENGTH ) "
    sql = sql & " or ( V.E040INGOTPOS + V.E040LENGTH > H.INGOTPOS "
    sql = sql & " and V.E040INGOTPOS + V.E040LENGTH < H.INGOTPOS + H.LENGTH  ) "
    sql = sql & " or ( H.INGOTPOS >= V.E040INGOTPOS "
    sql = sql & " and H.INGOTPOS < V.E040INGOTPOS + V.E040LENGTH ) "
    sql = sql & " or ( H.INGOTPOS + H.LENGTH > V.E040INGOTPOS "
    sql = sql & " and H.INGOTPOS + H.LENGTH < V.E040INGOTPOS + V.E040LENGTH )) "
                '工程コード、状態、区分の条件指定
    sql = sql & " and V.E040NOWPROC='CC600' "
    sql = sql & " and V.E040LSTATCLS='T' "
    sql = sql & " and V.E040RSTATCLS='T' "
    sql = sql & " and V.E040DELCLS='0' "
    'sql = sql & " and V.E040HOLDCLS='0' " ' ホールドブロックも取得
                '指示が0でなく実績が0
    sql = sql & " and ((V.E043CRYINDRS<>'0' and V.E043CRYRESRS='0') "         ' 結晶検査実績（Rs)
    sql = sql & " or (V.E043CRYINDOI<>'0' and V.E043CRYRESOI='0') "         ' 結晶検査実績（Oi)
    sql = sql & " or (V.E043CRYINDB1<>'0' and V.E043CRYRESB1='0')"          ' 結晶検査実績（B1)
    sql = sql & " or (V.E043CRYINDB2<>'0' and V.E043CRYRESB2='0') "         ' 結晶検査実績（B2）
    sql = sql & " or (V.E043CRYINDB3<>'0' and V.E043CRYRESB3='0') "         ' 結晶検査実績（B3)
    sql = sql & " or (V.E043CRYINDL1<>'0' and V.E043CRYRESL1='0') "         ' 結晶検査実績（L1)
    sql = sql & " or (V.E043CRYINDL2<>'0' and V.E043CRYRESL2='0') "         ' 結晶検査実績（L2)
    sql = sql & " or (V.E043CRYINDL3<>'0' and V.E043CRYRESL3='0') "         ' 結晶検査実績（L3)
    sql = sql & " or (V.E043CRYINDL4<>'0' and V.E043CRYRESL4='0') "         ' 結晶検査実績（L4)
    sql = sql & " or (V.E043CRYINDCS<>'0' and V.E043CRYRESCS='0') "         ' 結晶検査実績（Cs)
    sql = sql & " or (V.E043CRYINDGD<>'0' and V.E043CRYRESGD='0') "         ' 結晶検査実績（GD)
    sql = sql & " or (V.E043CRYINDT<>'0' and V.E043CRYREST='0') "           ' 結晶検査実績（T)
    sql = sql & " or (V.E043CRYINDEP<>'0' and V.E043CRYRESEP='0')) "         ' 結晶検査実績（EPD)
    sql = sql & " order by V.E040BLOCKID, H.INGOTPOS "

    'データを抽出する
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    'レコード0件時は正常
    If rs.RecordCount = 0 Then
        rs.Close
        ReDim records(0)
    Else
        BlockIdBuf = vbNullString
        recCnt = rs.RecordCount
        j = 0
        For i = 1 To recCnt
            DoEvents
        'ブロックID等の格納
            If rs("E040BLOCKID") <> BlockIdBuf Then
            
                j = j + 1
                ReDim Preserve records(j)
                
                With records(j)
                    .CRYNUM = rs("E040CRYNUM")
                    .IngotPos = rs("E040INGOTPOS")
                    .BLOCKID = rs("E040BLOCKID")   ' ブロックID
                    .UPDDATE = rs("E040UPDDATE")   ' 更新日付
                    .HOLDCLS = rs("E040HOLDCLS")   ' ホールド区分
                    BlockIdBuf = records(j).BLOCKID
                    .HSXTYPE = rs("HSXTYPE")
                    .HSXCDIR = rs("HSXCDIR")
                    .Judg = " "
                End With
                
                k = 1
            End If
            
            '品番の格納
            ReDim Preserve records(j).hin(k)
            records(j).hin(k).HINBAN = rs("HINBAN")
            records(j).hin(k).mnorevno = rs("REVNUM")
            records(j).hin(k).factory = rs("FACTORY")
            records(j).hin(k).opecond = rs("OPECOND")
            k = k + 1
            rs.MoveNext
        Next i
        rs.Close
            
    End If
        
    

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmkc001b_Disp1 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'概要    :待ち一覧 初期表示用ＤＢドライバ（判定待ち）
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO  ,型                                 ,説明
'        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,初期表示用
'        :戻ﾘ値        ,O   ,FUNCTION_RETURN                    ,読み込み成否
'説明    :
'履歴    :2001/07/06 蔵本 作成
Public Function DBDRV_scmzc_fcmkc001b_Disp2(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN

    '＜判定待ち＞
    '検査待ちが押されている場合と逆で０が一つもないもの
    Dim sql As String       'SQL全体
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      'ブロック管理のレコード数
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim BlockIdBuf As String


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc033_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp2"

    DBDRV_scmzc_fcmkc001b_Disp2 = FUNCTION_RETURN_SUCCESS
    
    sql = "select distinct "
    sql = sql & " B.CRYNUM, "
    sql = sql & " B.INGOTPOS as ss, "
'    sql = sql & " B.LENGTH, "             ' 長さ追加 2001/11/8
    sql = sql & " B.BLOCKID, "
    sql = sql & " B.UPDDATE, "
    sql = sql & " B.HOLDCLS, "
    sql = sql & " H.HINBAN, "            ' 品番
    sql = sql & " H.REVNUM, "            ' 製品番号改訂番号
    sql = sql & " H.FACTORY, "           ' 工場
    sql = sql & " H.OPECOND, "           ' 操業条件
    sql = sql & " S.HSXTYPE, "           ' 品ＳＸタイプ
    sql = sql & " S.HSXCDIR, "            ' 品ＳＸ結晶面方位
    sql = sql & " H.INGOTPOS, "
                '判定NGがあるかどうか
    sql = sql & " (select count(*) from VECME010 V1 "
    sql = sql & "  where V1.E040BLOCKID=B.BLOCKID "
    sql = sql & "  and ((V1.E043CRYINDRS<>'0' and V1.E043CRYRESRS='2') "         ' 結晶検査実績（Rs)
    sql = sql & "  or (V1.E043CRYINDOI<>'0' and V1.E043CRYRESOI='2') "         ' 結晶検査実績（Oi)
    sql = sql & "  or (V1.E043CRYINDB1<>'0' and V1.E043CRYRESB1='2')"          ' 結晶検査実績（B1)
    sql = sql & "  or (V1.E043CRYINDB2<>'0' and V1.E043CRYRESB2='2') "         ' 結晶検査実績（B2）
    sql = sql & "  or (V1.E043CRYINDB3<>'0' and V1.E043CRYRESB3='2') "         ' 結晶検査実績（B3)
    sql = sql & "  or (V1.E043CRYINDL1<>'0' and V1.E043CRYRESL1='2') "         ' 結晶検査実績（L1)
    sql = sql & "  or (V1.E043CRYINDL2<>'0' and V1.E043CRYRESL2='2') "         ' 結晶検査実績（L2)
    sql = sql & "  or (V1.E043CRYINDL3<>'0' and V1.E043CRYRESL3='2') "         ' 結晶検査実績（L3)
    sql = sql & "  or (V1.E043CRYINDL4<>'0' and V1.E043CRYRESL4='2') "         ' 結晶検査実績（L4)
    sql = sql & "  or (V1.E043CRYINDCS<>'0' and V1.E043CRYRESCS='2') "         ' 結晶検査実績（Cs)
    sql = sql & "  or (V1.E043CRYINDGD<>'0' and V1.E043CRYRESGD='2') "         ' 結晶検査実績（GD)
    sql = sql & "  or (V1.E043CRYINDT<>'0' and V1.E043CRYREST='2') "           ' 結晶検査実績（T)
    sql = sql & "  or (V1.E043CRYINDEP<>'0' and V1.E043CRYRESEP='2')) ) as J "         ' 結晶検査実績（EPD)
    sql = sql & " from "
    sql = sql & " TBCME040 B, TBCME041 H, TBCME018 S"
    sql = sql & " where "
    sql = sql & " B.CRYNUM = H.CRYNUM "
    sql = sql & " and H.HINBAN = S.HINBAN "
    sql = sql & " and H.REVNUM = S.MNOREVNO "
    sql = sql & " and H.FACTORY = S.FACTORY "
    sql = sql & " and H.OPECOND = S.OPECOND "
    
                '工程コード、状態、区分の条件指定
    sql = sql & " and B.NOWPROC='CC600' "
    sql = sql & " and B.LSTATCLS='T' "
    sql = sql & " and B.RSTATCLS='T' "
    sql = sql & " and B.DELCLS='0' "
    'sql = sql & " and B.HOLDCLS='0' " ' ホールドブロックも取得
                'ブロック内に含まれる品番を検索
    sql = sql & " and (( B.INGOTPOS >= H.INGOTPOS "
    sql = sql & " and B.INGOTPOS < H.INGOTPOS + H.LENGTH ) "
    sql = sql & " or ( B.INGOTPOS + B.LENGTH > H.INGOTPOS "
    sql = sql & " and B.INGOTPOS + B.LENGTH < H.INGOTPOS + H.LENGTH  ) "
    sql = sql & " or ( H.INGOTPOS >= B.INGOTPOS "
    sql = sql & " and H.INGOTPOS < B.INGOTPOS + B.LENGTH ) "
    sql = sql & " or ( H.INGOTPOS + H.LENGTH > B.INGOTPOS "
    sql = sql & " and H.INGOTPOS + H.LENGTH < B.INGOTPOS + B.LENGTH )) "
                '指示が0でなく実績が0でないサンプルが上下２枚あるか
    sql = sql & " and 2=( select count(*) "
    sql = sql & "  from VECME010 V2 "
    sql = sql & "  where "
    sql = sql & "  B.BLOCKID=V2.E040BLOCKID"
    sql = sql & "  and (V2.E043CRYINDRS='0' or V2.E043CRYRESRS<>'0') "         ' 結晶検査実績（Rs)
    sql = sql & "  and (V2.E043CRYINDOI='0' or V2.E043CRYRESOI<>'0') "         ' 結晶検査実績（Oi)
    sql = sql & "  and (V2.E043CRYINDB1='0' or V2.E043CRYRESB1<>'0')"          ' 結晶検査実績（B1)
    sql = sql & "  and (V2.E043CRYINDB2='0' or V2.E043CRYRESB2<>'0') "         ' 結晶検査実績（B2）
    sql = sql & "  and (V2.E043CRYINDB3='0' or V2.E043CRYRESB3<>'0') "         ' 結晶検査実績（B3)
    sql = sql & "  and (V2.E043CRYINDL1='0' or V2.E043CRYRESL1<>'0') "         ' 結晶検査実績（L1)
    sql = sql & "  and (V2.E043CRYINDL2='0' or V2.E043CRYRESL2<>'0') "         ' 結晶検査実績（L2)
    sql = sql & "  and (V2.E043CRYINDL3='0' or V2.E043CRYRESL3<>'0') "         ' 結晶検査実績（L3)
    sql = sql & "  and (V2.E043CRYINDL4='0' or V2.E043CRYRESL4<>'0') "         ' 結晶検査実績（L4)
    sql = sql & "  and (V2.E043CRYINDCS='0' or V2.E043CRYRESCS<>'0') "         ' 結晶検査実績（Cs)
    sql = sql & "  and (V2.E043CRYINDGD='0' or V2.E043CRYRESGD<>'0') "         ' 結晶検査実績（GD)
    sql = sql & "  and (V2.E043CRYINDT='0' or V2.E043CRYREST<>'0') "           ' 結晶検査実績（T)
    sql = sql & "  and (V2.E043CRYINDEP='0' or V2.E043CRYRESEP<>'0') )"         ' 結晶検査実績（EPD)
    sql = sql & " order by B.BLOCKID, H.INGOTPOS "
    
    'データを抽出する
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    'レコード0件時は正常
    If rs.RecordCount = 0 Then
        rs.Close
        ReDim records(0)
    Else
        BlockIdBuf = vbNullString
        recCnt = rs.RecordCount
        j = 0
        For i = 1 To recCnt
            DoEvents
        'ブロックID等の格納
            If rs("BLOCKID") <> BlockIdBuf Then
            
                j = j + 1
                ReDim Preserve records(j)
                
                With records(j)
                    .CRYNUM = rs("CRYNUM")
                    .IngotPos = rs("ss")
'                    .LENGTH = rs("LENGTH")      ' 長さ
                    .BLOCKID = rs("BLOCKID")   ' ブロックID
                    .UPDDATE = rs("UPDDATE")   ' 更新日付
                    .HOLDCLS = rs("HOLDCLS")   ' ホールド区分
                    BlockIdBuf = records(j).BLOCKID
                    .HSXTYPE = rs("HSXTYPE")
                    .HSXCDIR = rs("HSXCDIR")
                    If rs("J") > 0 Then
                        
                        .Judg = "2"
                    Else
                        .Judg = "1"
                    End If
                
                End With
                k = 1
            End If
            
            '品番の格納
            ReDim Preserve records(j).hin(k)
            records(j).hin(k).HINBAN = rs("HINBAN")
            records(j).hin(k).mnorevno = rs("REVNUM")
            records(j).hin(k).factory = rs("FACTORY")
            records(j).hin(k).opecond = rs("OPECOND")
            k = k + 1
            rs.MoveNext
        Next i
        rs.Close
            
    End If

    
    '購入単結晶実績取得
    If getKouBlock(records(), "CC600") = FUNCTION_RETURN_FAILURE Then
       DBDRV_scmzc_fcmkc001b_Disp2 = FUNCTION_RETURN_FAILURE
       GoTo proc_exit
    End If
    
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmkc001b_Disp2 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function



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
    gErr.Push "s_cmbc033_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp3"


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



'概要    :待ち一覧 初期表示用ＤＢドライバ（抜試指示待ち）
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO  ,型                                 ,説明
'        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,初期表示用
'        :戻ﾘ値        ,O   ,FUNCTION_RETURN                    ,読み込み成否
'説明    :
'履歴    :2001/07/06 蔵本 作成
Public Function DBDRV_scmzc_fcmkc001b_Disp4(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN

    '＜抜試指示待ち＞
    'CC710のもの
    
    'ブロックID､更新日付取得

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc033_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp4"

    DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_SUCCESS


    'ブロックID､更新日付、品番等取得
    If getBlockID(records(), "CC710") = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
'2000/08/24 S.Sano Start
'    '購入単結晶実績取得
'    If getKouBlock(records(), "CC710") = FUNCTION_RETURN_FAILURE Then
'       DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_FAILURE
'       GoTo proc_exit
'    End If
'2000/08/24 S.Sano End


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function
Public Function cmkc001b_DBDataCheck1(LWD() As cmkc001b_LockWait, Wd1() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
'    Dim typ_A As typ_AllTypes        '全情報構造体
'    Dim c0 As Integer
'    Dim sErrMsg As String
'    Dim NothingFlag As Boolean
'    Dim FuncAns As FUNCTION_RETURN
'    For c0 = 1 To UBound(Wd1())
'        NothingFlag = False
'        FuncAns = DBDRV_scmzc_fcmkc001b_Disp(Wd1(c0).BLOCKID, typ_A.typ_si, typ_A.typ_cr, typ_A.typ_zi, sErrMsg, NothingFlag)
'        LWD(c0).flag = NothingFlag
'    Next
    
   
    Dim l As Long, m As Long
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc033_SQL.bas -- Function cmkc001b_DBDataCheck1"

    
    cmkc001b_DBDataCheck1 = FUNCTION_RETURN_SUCCESS
    
    Set rs = Nothing
    
#If SPEEDUP Then   '高速化実験 02.1.28-2.15 野村
'高速化メモ
'候補となるブロックとその両端サンプルについて、検査状態をまとめて取得
'SQLの発行回数を抑制してメモリ内での処理に切り換える
Dim SMP() As tSmpMng
Dim idx As Integer
Dim topIdx As Integer
Dim botIdx As Integer

Debug.Print " 1:" & Time
    sql = vbNullString
'    sql = sql & "select"
'    sql = sql & "  B.BLOCKID, B.INGOTPOS as TOPPOS, B.INGOTPOS+LENGTH as BOTPOS"
'    sql = sql & ", S.CRYNUM, S.INGOTPOS, SMPKBN, HINBAN, REVNUM, FACTORY, OPECOND"
'    sql = sql & ", CRYINDRS, CRYRESRS, CRYINDOI, CRYRESOI"
'    sql = sql & ", CRYINDB1, CRYRESB1, CRYINDB2, CRYRESB2, CRYINDB3, CRYRESB3"
'    sql = sql & ", CRYINDL1, CRYRESL1, CRYINDL2, CRYRESL2, CRYINDL3, CRYRESL3, CRYINDL4, CRYRESL4"
'    sql = sql & ", CRYINDCS, CRYRESCS, CRYINDGD, CRYRESGD, CRYINDT, CRYREST, CRYINDEP, CRYRESEP "
'    sql = sql & "from TBCME043 S, TBCME040 B "
'    sql = sql & "where S.CRYNUM=B.CRYNUM"
'    sql = sql & "  and B.INGOTPOS>=0"
'    sql = sql & "  and B.DELCLS='0'"
'    sql = sql & "  and B.NOWPROC in ('CC600','CC700', 'CC710')"
'    sql = sql & "  and B.RSTATCLS='T'"
'    sql = sql & "  and B.HOLDCLS='0'"
'    sql = sql & "  and ((S.INGOTPOS=B.INGOTPOS) or (S.INGOTPOS=B.INGOTPOS+B.LENGTH)) "
'    sql = sql & "order by B.BLOCKID, S.INGOTPOS, S.SMPKBN"
    sql = sql & "select"
    sql = sql & "  B.BLOCKID, B.INGOTPOS as TOPPOS, B.INGOTPOS+LENGTH as BOTPOS"
    sql = sql & ", S.XTALCS, S.INPOSCS, SMPKBNCS, HINBCS, REVNUMCS, FACTORYCS, OPECS"
    sql = sql & ", CRYINDRSCS, CRYRESRS1CS, CRYINDOICS, CRYRESOICS"
    sql = sql & ", CRYINDB1CS, CRYRESB1CS, CRYINDB2CS, CRYRESB2CS, CRYINDB3CS, CRYRESB3CS"
    sql = sql & ", CRYINDL1CS, CRYRESL1CS, CRYINDL2CS, CRYRESL2CS, CRYINDL3CS, CRYRESL3CS, CRYINDL4CS, CRYRESL4CS"
    sql = sql & ", CRYINDCSCS, CRYRESCSCS, CRYINDGDCS, CRYRESGDCS, CRYINDTCS, CRYRESTCS, CRYINDEPCS, CRYRESEPCS "
    sql = sql & "from XSDCS S, TBCME040 B "
    sql = sql & "where S.XTALCS=B.CRYNUM"
    sql = sql & "  and B.INGOTPOS>=0"
    sql = sql & "  and B.DELCLS='0'"
    sql = sql & "  and B.NOWPROC in ('CC600','CC700', 'CC710')"
    sql = sql & "  and B.RSTATCLS='T'"
    sql = sql & "  and B.HOLDCLS='0'"
    sql = sql & "  and ((S.INPOSCS=B.INGOTPOS) or (S.INPOSCS=B.INGOTPOS+B.LENGTH)) "
    sql = sql & "order by B.BLOCKID, S.INPOSCS, S.SMPKBNCS"

    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    ReDim SMP(rs.RecordCount)
    With SMP(0)
        .BLOCKID = " "
        .CRYNUM = " "
        .SMPKBN = " "
        .HINBAN = " "
        .factory = " "
        .opecond = " "
        .CRYINDRS = " "
        .CRYRESRS = " "
        .CRYINDOI = " "
        .CRYRESOI = " "
        .CRYINDB1 = " "
        .CRYRESB1 = " "
        .CRYINDB2 = " "
        .CRYRESB2 = " "
        .CRYINDB3 = " "
        .CRYRESB3 = " "
        .CRYINDL1 = " "
        .CRYRESL1 = " "
        .CRYINDL2 = " "
        .CRYRESL2 = " "
        .CRYINDL3 = " "
        .CRYRESL3 = " "
        .CRYINDL4 = " "
        .CRYRESL4 = " "
        .CRYINDCS = " "
        .CRYRESCS = " "
        .CRYINDGD = " "
        .CRYRESGD = " "
        .CRYINDT = " "
        .CRYREST = " "
        .CRYINDEP = " "
        .CRYRESEP = " "
    End With

    For l = 1 To rs.RecordCount
        With SMP(l)
            .BLOCKID = rs("BLOCKID")
            .TOPPOS = rs("TOPPOS")
            .BOTPOS = rs("BOTPOS")
            .CRYNUM = rs("CRYNUM")
            .IngotPos = rs("INGOTPOS")
            .SMPKBN = rs("SMPKBN")
            .HINBAN = rs("HINBAN")
            .REVNUM = rs("REVNUM")
            .factory = rs("FACTORY")
            .opecond = rs("OPECOND")
            .CRYINDRS = rs("CRYINDRS")
            .CRYRESRS = rs("CRYRESRS")
            .CRYINDOI = rs("CRYINDOI")
            .CRYRESOI = rs("CRYRESOI")
            .CRYINDB1 = rs("CRYINDB1")
            .CRYRESB1 = rs("CRYRESB1")
            .CRYINDB2 = rs("CRYINDB2")
            .CRYRESB2 = rs("CRYRESB2")
            .CRYINDB3 = rs("CRYINDB3")
            .CRYRESB3 = rs("CRYRESB3")
            .CRYINDL1 = rs("CRYINDL1")
            .CRYRESL1 = rs("CRYRESL1")
            .CRYINDL2 = rs("CRYINDL2")
            .CRYRESL2 = rs("CRYRESL2")
            .CRYINDL3 = rs("CRYINDL3")
            .CRYRESL3 = rs("CRYRESL3")
            .CRYINDL4 = rs("CRYINDL4")
            .CRYRESL4 = rs("CRYRESL4")
            .CRYINDCS = rs("CRYINDCS")
            .CRYRESCS = rs("CRYRESCS")
            .CRYINDGD = rs("CRYINDGD")
            .CRYRESGD = rs("CRYRESGD")
            .CRYINDT = rs("CRYINDT")
            .CRYREST = rs("CRYREST")
            .CRYINDEP = rs("CRYINDEP")
            .CRYRESEP = rs("CRYRESEP")
        End With
        rs.MoveNext
    Next
    rs.Close
    Set rs = Nothing
Debug.Print " 2:" & Time
#End If
    
    For l = 1 To UBound(Wd1())
        DoEvents
        LWD(l).flag = False
'Debug.Print " " & l & ":" & Time
        
        With Wd1(l)
        
        ' 購入単結晶のブロックは無条件でＯＫ
        If Mid$(.BLOCKID, 1, 1) <> "8" Then
        
            ReDim .SMP(2)
                        
            ' 上下のサンプル情報取得
#If SPEEDUP Then   '高速化実験 02.1.28-2.15 野村
'高速化メモ
'一括取得した検査状態配列から、データを取得するように改造
            For m = 1 To 2
                DoEvents
                
                topIdx = 0
                botIdx = 0
                For idx = 1 To UBound(SMP)
                    If (SMP(idx).BLOCKID = .BLOCKID) Then
                        If (SMP(idx).SMPKBN = "T") Then
                            topIdx = idx
                        Else
                            botIdx = idx
                        End If
                    ElseIf SMP(idx).BLOCKID > .BLOCKID Then
                        Exit For
                    End If
                Next
                If m = 1 Then
                    If topIdx > 0 Then
                        idx = topIdx
                    Else
                        idx = botIdx
                    End If
                Else
                    If botIdx > 0 Then
                        idx = botIdx
                    Else
                        idx = topIdx
                    End If
                End If
                
                With .SMP(m)
                    .CRYNUM = SMP(idx).CRYNUM
                    .IngotPos = SMP(idx).IngotPos
                    .SMPKBN = SMP(idx).SMPKBN
                    .HINBAN = SMP(idx).HINBAN
                    .REVNUM = SMP(idx).REVNUM
                    .factory = SMP(idx).factory
                    .opecond = SMP(idx).opecond
                    .CRYINDRS = SMP(idx).CRYINDRS
                    .CRYRESRS = SMP(idx).CRYRESRS
                    .CRYINDOI = SMP(idx).CRYINDOI
                    .CRYRESOI = SMP(idx).CRYRESOI
                    .CRYINDB1 = SMP(idx).CRYINDB1
                    .CRYRESB1 = SMP(idx).CRYRESB1
                    .CRYINDB2 = SMP(idx).CRYINDB2
                    .CRYRESB2 = SMP(idx).CRYRESB2
                    .CRYINDB3 = SMP(idx).CRYINDB3
                    .CRYRESB3 = SMP(idx).CRYRESB3
                    .CRYINDL1 = SMP(idx).CRYINDL1
                    .CRYRESL1 = SMP(idx).CRYRESL1
                    .CRYINDL2 = SMP(idx).CRYINDL2
                    .CRYRESL2 = SMP(idx).CRYRESL2
                    .CRYINDL3 = SMP(idx).CRYINDL3
                    .CRYRESL3 = SMP(idx).CRYRESL3
                    .CRYINDL4 = SMP(idx).CRYINDL4
                    .CRYRESL4 = SMP(idx).CRYRESL4
                    .CRYINDCS = SMP(idx).CRYINDCS
                    .CRYRESCS = SMP(idx).CRYRESCS
                    .CRYINDGD = SMP(idx).CRYINDGD
                    .CRYRESGD = SMP(idx).CRYRESGD
                    .CRYINDT = SMP(idx).CRYINDT
                    .CRYREST = SMP(idx).CRYREST
                    .CRYINDEP = SMP(idx).CRYINDEP
                    .CRYRESEP = SMP(idx).CRYRESEP
                End With
            Next m
            
#Else
            sql = " select "
            sql = sql & " V.E043CRYNUM, "
            sql = sql & " V.E043INGOTPOS, "
            sql = sql & " V.E043SMPKBN, "
            sql = sql & " V.E043HINBAN, "
            sql = sql & " V.E043REVNUM, "
            sql = sql & " V.E043FACTORY, "
            sql = sql & " V.E043OPECOND, "
            sql = sql & " V.E043CRYINDRS, "
            sql = sql & " V.E043CRYRESRS, "
            sql = sql & " V.E043CRYINDOI, "
            sql = sql & " V.E043CRYRESOI, "
            sql = sql & " V.E043CRYINDB1, "
            sql = sql & " V.E043CRYRESB1, "
            sql = sql & " V.E043CRYINDB2, "
            sql = sql & " V.E043CRYRESB2, "
            sql = sql & " V.E043CRYINDB3, "
            sql = sql & " V.E043CRYRESB3, "
            sql = sql & " V.E043CRYINDL1, "
            sql = sql & " V.E043CRYRESL1, "
            sql = sql & " V.E043CRYINDL2, "
            sql = sql & " V.E043CRYRESL2, "
            sql = sql & " V.E043CRYINDL3, "
            sql = sql & " V.E043CRYRESL3, "
            sql = sql & " V.E043CRYINDL4, "
            sql = sql & " V.E043CRYRESL4, "
            sql = sql & " V.E043CRYINDCS, "
            sql = sql & " V.E043CRYRESCS, "
            sql = sql & " V.E043CRYINDGD, "
            sql = sql & " V.E043CRYRESGD, "
            sql = sql & " V.E043CRYINDT, "
            sql = sql & " V.E043CRYREST, "
            sql = sql & " V.E043CRYINDEP, "
            sql = sql & " V.E043CRYRESEP "
            sql = sql & " from VECME010 V "
            sql = sql & " where E040CRYNUM = '" & .CRYNUM & "' "
            sql = sql & " and   E040INGOTPOS = '" & .IngotPos & "' "
            sql = sql & " order by E043INGOTPOS"
            
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            
            For m = 1 To 2
                DoEvents
                .SMP(m).CRYNUM = rs("E043CRYNUM")
                .SMP(m).IngotPos = rs("E043INGOTPOS")
                .SMP(m).SMPKBN = rs("E043SMPKBN")
                .SMP(m).HINBAN = rs("E043HINBAN")
                .SMP(m).REVNUM = rs("E043REVNUM")
                .SMP(m).factory = rs("E043FACTORY")
                .SMP(m).opecond = rs("E043OPECOND")
                .SMP(m).CRYINDRS = rs("E043CRYINDRS")
                .SMP(m).CRYRESRS = rs("E043CRYRESRS")
                .SMP(m).CRYINDOI = rs("E043CRYINDOI")
                .SMP(m).CRYRESOI = rs("E043CRYRESOI")
                .SMP(m).CRYINDB1 = rs("E043CRYINDB1")
                .SMP(m).CRYRESB1 = rs("E043CRYRESB1")
                .SMP(m).CRYINDB2 = rs("E043CRYINDB2")
                .SMP(m).CRYRESB2 = rs("E043CRYRESB2")
                .SMP(m).CRYINDB3 = rs("E043CRYINDB3")
                .SMP(m).CRYRESB3 = rs("E043CRYRESB3")
                .SMP(m).CRYINDL1 = rs("E043CRYINDL1")
                .SMP(m).CRYRESL1 = rs("E043CRYRESL1")
                .SMP(m).CRYINDL2 = rs("E043CRYINDL2")
                .SMP(m).CRYRESL2 = rs("E043CRYRESL2")
                .SMP(m).CRYINDL3 = rs("E043CRYINDL3")
                .SMP(m).CRYRESL3 = rs("E043CRYRESL3")
                .SMP(m).CRYINDL4 = rs("E043CRYINDL4")
                .SMP(m).CRYRESL4 = rs("E043CRYRESL4")
                .SMP(m).CRYINDCS = rs("E043CRYINDCS")
                .SMP(m).CRYRESCS = rs("E043CRYRESCS")
                .SMP(m).CRYINDGD = rs("E043CRYINDGD")
                .SMP(m).CRYRESGD = rs("E043CRYRESGD")
                .SMP(m).CRYINDT = rs("E043CRYINDT")
                .SMP(m).CRYREST = rs("E043CRYREST")
                .SMP(m).CRYINDEP = rs("E043CRYINDEP")
                .SMP(m).CRYRESEP = rs("E043CRYRESEP")
                
                rs.MoveNext
            Next m
            rs.Close
            Set rs = Nothing
#End If
            
'高速化メモ
'品番仕様/Cs/EPD/LTはまだブロック毎にSQLを投げている
'ここをまとめていけば、あと5秒程度縮むのではないかと思われる
'ただし、Cs/LTについては結果取得の方法が変わるので、その後の検討が必要
'いずれにせよ、対象結晶全てについてCs/LT/EPD指示のあるサンプルを抜き出せばよいはず
            
            ' 品番の仕様情報取得
            For m = 1 To 2
                If Trim$(.SMP(m).HINBAN) = "G" Or Trim$(.SMP(m).HINBAN) = "Z" Then
                    .SMP(m).HSXCNHWS = "S"
                    .SMP(m).HSXLTHWS = "S"
                    .SMP(m).EPD = "S"
                ElseIf Len(Trim$(.SMP(m).HINBAN)) Then
                    sql = " select "
                    sql = sql & " S.HSXCNHWS, "
                    sql = sql & " S.HSXLTHWS, "
                    sql = sql & " 'H' as EPD "
                    sql = sql & " from TBCME019 S "
                    sql = sql & " where S.HINBAN = '" & .SMP(m).HINBAN & "' "
                    sql = sql & " and S.MNOREVNO = " & .SMP(m).REVNUM & " "
                    sql = sql & " and S.FACTORY = '" & .SMP(m).factory & "' "
                    sql = sql & " and S.OPECOND = '" & .SMP(m).opecond & "' "
        
                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                    .SMP(m).HSXCNHWS = rs("HSXCNHWS")
                    .SMP(m).HSXLTHWS = rs("HSXLTHWS")
                    .SMP(m).EPD = rs("EPD")
                    
                    rs.Close
                    Set rs = Nothing
                Else
                    '空品番の場合
                    .SMP(m).HSXCNHWS = " "
                    .SMP(m).HSXLTHWS = " "
                    .SMP(m).EPD = " "
                End If
            Next m
        
            ' チェック
            For m = 1 To 2
                DoEvents
                ' CSのチェック
'                If (.SMP(m).HSXCNHWS = "H" Or .SMP(m).HSXCNHWS = "S") And .SMP(m).CRYINDCS = "0" Then  ' 参考評価はなくてもＯＫ
                If .SMP(m).HSXCNHWS = "H" And .SMP(m).CRYINDCS = "0" Then
                
                    sql = "select CRYRESCSCS as RES from XSDCS "
                    sql = sql & "where XTALCS = '" & .SMP(m).CRYNUM & "' "
                    sql = sql & "  and INPOSCS >= " & .SMP(m).IngotPos
                    sql = sql & "  and CRYINDCSCS<>'0'"
                    sql = sql & " order by INPOSCS"
                    
                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                    If rs.RecordCount Then
                        If rs("RES") = "0" Then LWD(l).flag = True
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                End If
                
                ' LTのチェック
'                If (.SMP(m).HSXLTHWS = "H" Or .SMP(m).HSXLTHWS = "S") And .SMP(m).CRYINDT = "0" And LWD(l).flag = False Then ' 参考評価はなくてもＯＫ
                If .SMP(m).HSXLTHWS = "H" And .SMP(m).CRYINDT = "0" And LWD(l).flag = False Then
                    
                    sql = "select CRYRESTCS as RES from XSDCS "
                    sql = sql & "where XTALCS = '" & .SMP(m).CRYNUM & "' "
                    sql = sql & "  and INPOSCS >= " & .SMP(m).IngotPos
                    sql = sql & "  and CRYINDTCS<>'0'"
                    sql = sql & " order by INPOSCS"
                    
                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                    If rs.RecordCount Then
                        If rs("RES") = "0" Then LWD(l).flag = True
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                End If
                
                ' EPDのチェック
'                If (.SMP(m).EPD = "H" Or .SMP(m).EPD = "S") And .SMP(m).CRYINDEP = "0" And LWD(l).flag = False Then ' Sはありえなけど統一
                If .SMP(m).EPD = "H" And .SMP(m).CRYINDEP = "0" And LWD(l).flag = False Then ' Sはありえなけど統一
                   
                    sql = "select CRYRESEPCS as RES from XSDCS "
                    sql = sql & "where XTALCS = '" & .SMP(m).CRYNUM & "' "
                    sql = sql & "  and INPOSCS >= " & .SMP(m).IngotPos
                    sql = sql & "  and CRYINDEP<>'0'"
                    sql = sql & " order by INPOSCS"
                    
                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                    If rs.RecordCount Then
                        If rs("RES") = "0" Then LWD(l).flag = True
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                End If
'                If LWD(l).flag = True Then
'                    Exit For
'                End If
            Next m
        End If
        
        End With    ' .Wd1()
        
    Next l
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    cmkc001b_DBDataCheck1 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

Public Function cmkc001b_DBDataCheck3(LWD() As cmkc001b_LockWait, _
                                 Wd3() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
    Dim c0 As Integer
    Dim c1 As Integer
    Dim c2 As Integer
    Dim MaxRec As Integer
    Dim recCount As Integer
    Dim EQFlag As Boolean
    Dim sql As String       'SQL全体
    Dim rs As OraDynaset    'RecordSet
    Dim GrpCount1 As Integer
    Dim GrpCount2 As Integer
    Dim ColorFlag As Boolean
    Dim TotalBlk As Integer
    Dim CheckPoint As Integer
    Dim CheckEnd As Integer
    Dim tempGrpFlag As String * 1
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc033_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp"

    cmkc001b_DBDataCheck3 = FUNCTION_RETURN_SUCCESS
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
    MaxRec = UBound(GrpInfo())
    For c0 = 1 To MaxRec
        sql = "select "
        sql = sql & "BLOCKID, "
        sql = sql & "INGOTPOS, "
        sql = sql & "LENGTH, "
        sql = sql & "NOWPROC, "
        sql = sql & "HOLDCLS "
        sql = sql & "from TBCME040 "
        sql = sql & "where CRYNUM='" & GrpInfo(c0).CRYNUM & "' "
'2001/11/14 S.Sano        sql = sql & "and LSTATCLS='T' "
'2001/11/14 S.Sano        sql = sql & "and RSTATCLS='T' "
'2001/11/14 S.Sano        sql = sql & "and DELCLS='0' "
        'sql = sql & "and HOLDCLS='0' "
        sql = sql & "order by BLOCKID "
    
        
        'データを抽出する
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        recCount = rs.RecordCount
        If recCount = 0 Then
            rs.Close
            GoTo proc_exit
        End If
        ReDim GrpInfo(c0).blkInfo(recCount) As cmkc001b_Wait3_BLK
        For c1 = 1 To recCount
            GrpInfo(c0).blkInfo(c1).BLOCKID = rs("BLOCKID")
            GrpInfo(c0).blkInfo(c1).IngotPos = rs("INGOTPOS")
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
Dim blkID() As String
Dim topHin() As tFullHinban
Dim botHin() As tFullHinban
Dim idx As Integer
Dim rsCount As Integer
Dim found As Boolean

    sql = vbNullString
    sql = sql & "select"
    sql = sql & "  b.BLOCKID"
    sql = sql & ", TOP.HINBAN as THINBAN, TOP.REVNUM as TREVNUM, TOP.FACTORY as TFACTORY, TOP.OPECOND as TOPECOND"
    sql = sql & ", BOT.HINBAN as BHINBAN, BOT.REVNUM as BREVNUM, BOT.FACTORY as BFACTORY, BOT.OPECOND as BOPECOND "
    sql = sql & "from TBCME040 B, TBCME041 TOP, TBCME041 BOT "
    sql = sql & "Where b.CRYNUM = Top.CRYNUM"
    sql = sql & "  and B.CRYNUM=BOT.CRYNUM"
    sql = sql & "  and B.INGOTPOS>=0"
    sql = sql & "  and B.DELCLS='0'"
    sql = sql & "  and B.NOWPROC in ('CC600','CC700', 'CC710')"
    sql = sql & "  and B.RSTATCLS='T'"
    sql = sql & "  and B.HOLDCLS='0'"
    sql = sql & "  and B.INGOTPOS>=TOP.INGOTPOS"
    sql = sql & "  and B.INGOTPOS<TOP.INGOTPOS+TOP.LENGTH"
    sql = sql & "  and B.INGOTPOS+B.LENGTH>BOT.INGOTPOS"
    sql = sql & "  and B.INGOTPOS+B.LENGTH<=BOT.INGOTPOS+BOT.LENGTH "
    sql = sql & "order by B.BLOCKID"
    
    'データを抽出する
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    rsCount = rs.RecordCount
    ReDim blkID(1 To rsCount)
    ReDim topHin(1 To rsCount)
    ReDim botHin(1 To rsCount)
    For c0 = 1 To rsCount
        blkID(c0) = rs!BLOCKID
        topHin(c0).HINBAN = rs!THINBAN
        topHin(c0).mnorevno = rs!TREVNUM
        topHin(c0).factory = rs!TFACTORY
        topHin(c0).opecond = rs!TOPECOND
        botHin(c0).HINBAN = rs!BHINBAN
        botHin(c0).mnorevno = rs!BREVNUM
        botHin(c0).factory = rs!BFACTORY
        botHin(c0).opecond = rs!BOPECOND
        rs.MoveNext
    Next
    rs.Close

    For c0 = 1 To MaxRec
        recCount = UBound(GrpInfo(c0).blkInfo())
        For c1 = 1 To recCount
            found = False
            For idx = 1 To rsCount
                If blkID(idx) = GrpInfo(c0).blkInfo(c1).BLOCKID Then
                    found = True
                    Exit For
                ElseIf blkID(idx) > GrpInfo(c0).blkInfo(c1).BLOCKID Then
                    Exit For
                End If
            Next
        
            If found Then
                GrpInfo(c0).blkInfo(c1).topHin.HINBAN = topHin(idx).HINBAN
                GrpInfo(c0).blkInfo(c1).topHin.factory = topHin(idx).factory
                GrpInfo(c0).blkInfo(c1).topHin.opecond = topHin(idx).opecond
                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = topHin(idx).mnorevno
            Else
                GrpInfo(c0).blkInfo(c1).topHin.HINBAN = ""
                GrpInfo(c0).blkInfo(c1).topHin.factory = ""
                GrpInfo(c0).blkInfo(c1).topHin.opecond = ""
                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = 0
            End If
            
            If found Then
                GrpInfo(c0).blkInfo(c1).botHin.HINBAN = botHin(idx).HINBAN
                GrpInfo(c0).blkInfo(c1).botHin.factory = botHin(idx).factory
                GrpInfo(c0).blkInfo(c1).botHin.opecond = botHin(idx).opecond
                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = botHin(idx).mnorevno
            Else
                GrpInfo(c0).blkInfo(c1).botHin.HINBAN = ""
                GrpInfo(c0).blkInfo(c1).botHin.factory = ""
                GrpInfo(c0).blkInfo(c1).botHin.opecond = ""
                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = 0
            End If
        Next
    Next
#Else
    For c0 = 1 To MaxRec
        recCount = UBound(GrpInfo(c0).blkInfo())
        For c1 = 1 To recCount
            sql = "select "
            sql = sql & "HINBAN, "
            sql = sql & "REVNUM, "
            sql = sql & "FACTORY, "
            sql = sql & "OPECOND "
            sql = sql & "from TBCME041 "
            sql = sql & "where CRYNUM='" & GrpInfo(c0).CRYNUM & "' "
'2001/11/14 S.Sano            sql = sql & "and INGOTPOS <= " & GrpInfo(c0).blkInfo(c1).INGOTPOS & " "
            sql = sql & "and INGOTPOS = " & GrpInfo(c0).blkInfo(c1).IngotPos & " " '2001/11/14 S.Sano
'2001/11/14 S.Sano            sql = sql & "and (INGOTPOS + LENGTH) > " & GrpInfo(c0).blkInfo(c1).INGOTPOS & " "
            
            'データを抽出する
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            recCount = rs.RecordCount
            If recCount = 0 Then
                GrpInfo(c0).blkInfo(c1).topHin.HINBAN = ""
                GrpInfo(c0).blkInfo(c1).topHin.factory = ""
                GrpInfo(c0).blkInfo(c1).topHin.opecond = ""
                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = 0
            Else
                GrpInfo(c0).blkInfo(c1).topHin.HINBAN = rs("HINBAN")
                GrpInfo(c0).blkInfo(c1).topHin.factory = rs("FACTORY")
                GrpInfo(c0).blkInfo(c1).topHin.opecond = rs("OPECOND")
                GrpInfo(c0).blkInfo(c1).topHin.REVNUM = rs("REVNUM")
            End If
            rs.Close
        
            sql = "select "
            sql = sql & "HINBAN, "
            sql = sql & "REVNUM, "
            sql = sql & "FACTORY, "
            sql = sql & "OPECOND "
            sql = sql & "from TBCME041 "
            sql = sql & "where CRYNUM='" & GrpInfo(c0).CRYNUM & "' "
            sql = sql & "and INGOTPOS < " & GrpInfo(c0).blkInfo(c1).IngotPos + GrpInfo(c0).blkInfo(c1).LENGTH & " "
            sql = sql & "and (INGOTPOS + LENGTH) >= " & GrpInfo(c0).blkInfo(c1).IngotPos + GrpInfo(c0).blkInfo(c1).LENGTH & " "
            
            'データを抽出する
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            recCount = rs.RecordCount
            If recCount = 0 Then
                GrpInfo(c0).blkInfo(c1).botHin.HINBAN = ""
                GrpInfo(c0).blkInfo(c1).botHin.factory = ""
                GrpInfo(c0).blkInfo(c1).botHin.opecond = ""
                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = 0
            Else
                GrpInfo(c0).blkInfo(c1).botHin.HINBAN = rs("HINBAN")
                GrpInfo(c0).blkInfo(c1).botHin.factory = rs("FACTORY")
                GrpInfo(c0).blkInfo(c1).botHin.opecond = rs("OPECOND")
                GrpInfo(c0).blkInfo(c1).botHin.REVNUM = rs("REVNUM")
            End If
            rs.Close
        Next
    Next
#End If
    
Debug.Print " 4:" & Time
    '求めた情報からグループを求める
    GrpCount1 = 0
    GrpCount2 = 0
    For c0 = 1 To MaxRec
        GrpCount1 = GrpCount1 + 1
        GrpCount2 = GrpCount2 + 1
        recCount = UBound(GrpInfo(c0).blkInfo())
        For c1 = 1 To recCount
            'ブロック切れ目で品番が変われば別グループと判断する
            Select Case c1
            Case 1
                GrpInfo(c0).blkInfo(c1).GRPFLG1 = GrpCount1
            Case Else
                If (GrpInfo(c0).blkInfo(c1).topHin.factory <> GrpInfo(c0).blkInfo(c1 - 1).botHin.factory) Or _
                   (GrpInfo(c0).blkInfo(c1).topHin.HINBAN <> GrpInfo(c0).blkInfo(c1 - 1).botHin.HINBAN) Or _
                   (GrpInfo(c0).blkInfo(c1).topHin.opecond <> GrpInfo(c0).blkInfo(c1 - 1).botHin.opecond) Or _
                   (GrpInfo(c0).blkInfo(c1).topHin.REVNUM <> GrpInfo(c0).blkInfo(c1 - 1).botHin.REVNUM) Then
                    GrpCount1 = GrpCount1 + 1
                End If
                GrpInfo(c0).blkInfo(c1).GRPFLG1 = GrpCount1
            End Select
            
            '同一グループ内で、工程違いのブロックが存在した場合、同一グループ内の
            '小グループとしてグループ分けする。
            'CC710以外なら対象外としグループ判定をしない
            If GrpInfo(c0).blkInfo(c1).NOWPROC = PROCD_NUKISI_SIJI And GrpInfo(c0).blkInfo(c1).HOLDCLS = "0" Then
                Select Case c1
                Case 1
                    GrpInfo(c0).blkInfo(c1).GRPFLG2 = GrpCount2
                Case Else
                    If (GrpInfo(c0).blkInfo(c1).topHin.factory <> GrpInfo(c0).blkInfo(c1 - 1).botHin.factory) Or _
                       (GrpInfo(c0).blkInfo(c1).topHin.HINBAN <> GrpInfo(c0).blkInfo(c1 - 1).botHin.HINBAN) Or _
                       (GrpInfo(c0).blkInfo(c1).topHin.opecond <> GrpInfo(c0).blkInfo(c1 - 1).botHin.opecond) Or _
                       (GrpInfo(c0).blkInfo(c1).topHin.REVNUM <> GrpInfo(c0).blkInfo(c1 - 1).botHin.REVNUM) Then
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
        recCount = UBound(GrpInfo(c0).blkInfo())
        ColorFlag = False
        CheckPoint = 0
        For c1 = 1 To recCount
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
        For c1 = CheckPoint To recCount
            GrpInfo(c0).blkInfo(c1).COLORFLG = ColorFlag
        Next
    Next
Debug.Print " 6:" & Time
    For c0 = 1 To MaxRec
        recCount = UBound(GrpInfo(c0).blkInfo())
        For c1 = 1 To recCount
            For c2 = 1 To TotalBlk
                If Wd3(c2).BLOCKID = GrpInfo(c0).blkInfo(c1).BLOCKID Then
                    LWD(c2).flag = GrpInfo(c0).blkInfo(c1).COLORFLG
                    LWD(c2).Grp = GrpInfo(c0).blkInfo(c1).GRPFLG2
                    Exit For
                End If
            Next
        Next
    Next
'    Debug.Print Now

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    cmkc001b_DBDataCheck3 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'購入単結晶用
Private Function getKouBlock(records() As type_DBDRV_scmzc_fcmkc001b_Disp, NOWPROC As String) As FUNCTION_RETURN

    Dim sql As String       'SQL全体
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long
    Dim motoRecCnt As Long
    Dim i As Long

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc033_SQL.bas -- Function getKouBlock"

    getKouBlock = FUNCTION_RETURN_SUCCESS

    sql = " select "
    sql = sql & " B.BLOCKID, "
    sql = sql & " B.UPDDATE, "
    sql = sql & " B.HOLDCLS, "
    sql = sql & " K.HINBAN, "
    sql = sql & " K.MNOREVNO, "
    sql = sql & " K.FACTORY, "
    sql = sql & " K.OPECOND "
    sql = sql & " from TBCME040 B,TBCMG002 K "
    sql = sql & " where B.BLOCKID=K.CRYNUM "
    sql = sql & " and substr(B.BLOCKID,1,1)='8' "
    sql = sql & " and B.NOWPROC='" & NOWPROC & "' "
    sql = sql & " and B.LSTATCLS='T' "
    sql = sql & " and B.RSTATCLS='T' "
    sql = sql & " and B.DELCLS='0' "
    'sql = sql & " and B.HOLDCLS='0' " ' ホールドブロックも取得
    sql = sql & " and K.TRANCNT=any(select max(TRANCNT) from TBCMG002 where CRYNUM=B.BLOCKID ) "
    sql = sql & " order by B.BLOCKID "

    
    'データを抽出する
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount = 0 Then
        rs.Close
        GoTo proc_exit
    End If
    
    motoRecCnt = UBound(records)
    recCnt = rs.RecordCount
    ReDim Preserve records(UBound(records) + recCnt)
    
    For i = motoRecCnt + 1 To UBound(records)
        DoEvents
        ReDim records(i).hin(1)
        With records(i)
            .BLOCKID = rs("BLOCKID")     ' ブロックID
            .UPDDATE = rs("UPDDATE")     ' 更新日付
            .HOLDCLS = rs("HOLDCLS")     ' ホールド区分
            .hin(1).HINBAN = rs("HINBAN")       ' 品番
            .hin(1).mnorevno = rs("MNOREVNO")   ' 製品番号改訂番号
            .hin(1).factory = rs("FACTORY")     ' 工場
            .hin(1).opecond = rs("OPECOND")     ' 操業条件
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
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getKouBlock = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
    
End Function

'内部関数 ブロックID、更新日付取得（払出待ち、抜試指示待ち用）
Private Function getBlockID(records() As type_DBDRV_scmzc_fcmkc001b_Disp, _
                            NOWPROC As String) As FUNCTION_RETURN

    Dim sql As String       'SQL全体
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      'レコード数
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim BlockIdBuf As String


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function getBlockID"

    getBlockID = FUNCTION_RETURN_SUCCESS

    sql = "select "
    sql = sql & " V.E040CRYNUM, "
    sql = sql & " V.E040BLOCKID, "
    sql = sql & " V.E040INGOTPOS, "
    sql = sql & " V.E040UPDDATE, "
    sql = sql & " V.E040HOLDCLS, "
    sql = sql & " V.E041HINBAN, "            ' 品番
    sql = sql & " V.E041REVNUM, "            ' 製品番号改訂番号
    sql = sql & " V.E041FACTORY, "           ' 工場
    sql = sql & " V.E041OPECOND, "           ' 操業条件
    sql = sql & " S.HSXTYPE, "           ' 品ＳＸタイプ
    sql = sql & " S.HSXCDIR "            ' 品ＳＸ結晶面方位
    sql = sql & " from "
    sql = sql & " VECME009 V, TBCME018 S "
    sql = sql & " where "
    sql = sql & " V.E041HINBAN = S.HINBAN "
    sql = sql & " and V.E041REVNUM = S.MNOREVNO "
    sql = sql & " and V.E041FACTORY = S.FACTORY "
    sql = sql & " and V.E041OPECOND = S.OPECOND "
    sql = sql & " and V.E040NOWPROC='" & NOWPROC & "' "
    sql = sql & " and V.E040LSTATCLS='T' "
    sql = sql & " and V.E040RSTATCLS='T' "
    sql = sql & " and V.E040DELCLS='0' "
    'sql = sql & " and V.E040HOLDCLS='0' " ' ホールドブロックも取得
    sql = sql & " order by V.E040BLOCKID, V.E041INGOTPOS "

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
        If rs("E040BLOCKID") <> BlockIdBuf Then
        
            j = j + 1
            ReDim Preserve records(j)
            
            With records(j)
                .CRYNUM = rs("E040CRYNUM")
                .IngotPos = rs("E040INGOTPOS")
                .BLOCKID = rs("E040BLOCKID")   ' ブロックID
                .UPDDATE = rs("E040UPDDATE")   ' 更新日付
                .HOLDCLS = rs("E040HOLDCLS")   ' ホールド区分
                BlockIdBuf = records(j).BLOCKID
                .HSXTYPE = rs("HSXTYPE")
                .HSXCDIR = rs("HSXCDIR")
                .Judg = " "
            End With
            
            k = 1
        End If
        
        '品番の格納
        ReDim Preserve records(j).hin(k)
        records(j).hin(k).HINBAN = rs("E041HINBAN")
        records(j).hin(k).mnorevno = rs("E041REVNUM")
        records(j).hin(k).factory = rs("E041FACTORY")
        records(j).hin(k).opecond = rs("E041OPECOND")
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
    getBlockID = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function



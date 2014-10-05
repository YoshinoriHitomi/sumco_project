Attribute VB_Name = "s_cmbc030_SQL"
Option Explicit

Type cmkc001b_LockWait
    flag        As Boolean
    Grp         As Integer
End Type

Type cmkc001b_Wait3_HINBAN
    hinban      As String * 8               ' 品番
    REVNUM      As Integer                  ' 製品番号改訂番号
    factory     As String * 1               ' 工場
    opecond     As String * 1               ' 操業条件
End Type

Type cmkc001b_Wait3_BLK
    BLOCKID     As String * 12              ' ブロックID
    IngotPos    As Integer                  ' 結晶内開始位置
    LENGTH      As Integer                  ' 長さ
    NOWPROC     As String * 5               ' 現在工程
    HOLDCLS     As String * 1               ' ホールド区分 ---kuramoto 追加 2001/09/19----
    GRPFLG1     As Integer                  ' グループ情報
    GRPFLG2     As Integer                  ' グループ情報
    COLORFLG    As Boolean
    topHin      As cmkc001b_Wait3_HINBAN
    botHin      As cmkc001b_Wait3_HINBAN
End Type

Type cmkc001b_Wait3
    CRYNUM      As String * 12              ' 結晶番号
    blkInfo()   As cmkc001b_Wait3_BLK
End Type

Type type_cmkc001b_SmpMng
    CRYNUM      As String * 12
    IngotPos    As Integer
    SMPKBN      As String * 1
    
    hinban      As String * 8               ' 品番
    REVNUM      As Integer                  ' 製品番号改訂番号
    factory     As String * 1               ' 工場
    opecond     As String * 1               ' 操業条件
        
    CRYINDRS    As String * 1
    CRYRESRS    As String * 1
    CRYINDOI    As String * 1
    CRYRESOI    As String * 1
    CRYINDB1    As String * 1
    CRYRESB1    As String * 1
    CRYINDB2    As String * 1
    CRYRESB2    As String * 1
    CRYINDB3    As String * 1
    CRYRESB3    As String * 1
    CRYINDL1    As String * 1
    CRYRESL1    As String * 1
    CRYINDL2    As String * 1
    CRYRESL2    As String * 1
    CRYINDL3    As String * 1
    CRYRESL3    As String * 1
    CRYINDL4    As String * 1
    CRYRESL4    As String * 1
    CRYINDCS    As String * 1
    CRYRESCS    As String * 1
    CRYINDGD    As String * 1
    CRYRESGD    As String * 1
    CRYINDT     As String * 1
    CRYREST     As String * 1
    CRYINDEP    As String * 1
    CRYRESEP    As String * 1
    
    HSXCNHWS    As String * 1               ' 品ＳＸ炭素濃度保証方法＿処
    HSXLTHWS    As String * 1               ' 品ＳＸＬタイム保証方法＿処
    EPD         As String * 1               ' EPD
End Type

'''''#If SPEEDUP Then   '高速化実験 02.1.28-2.15 野村
'''''Private Type tSmpMng
'''''    BLOCKID     As String * 12
'''''    TOPPOS      As Integer
'''''    BOTPOS      As Integer
'''''
'''''    CRYNUM      As String * 12
'''''    IngotPos    As Integer
'''''    SMPKBN      As String * 1
'''''
'''''    hinban      As String * 8           ' 品番
'''''    REVNUM      As Integer              ' 製品番号改訂番号
'''''    factory     As String * 1           ' 工場
'''''    opecond     As String * 1           ' 操業条件
'''''
'''''    CRYINDRS    As String * 1
'''''    CRYRESRS    As String * 1
'''''    CRYINDOI    As String * 1
'''''    CRYRESOI    As String * 1
'''''    CRYINDB1    As String * 1
'''''    CRYRESB1    As String * 1
'''''    CRYINDB2    As String * 1
'''''    CRYRESB2    As String * 1
'''''    CRYINDB3    As String * 1
'''''    CRYRESB3    As String * 1
'''''    CRYINDL1    As String * 1
'''''    CRYRESL1    As String * 1
'''''    CRYINDL2    As String * 1
'''''    CRYRESL2    As String * 1
'''''    CRYINDL3    As String * 1
'''''    CRYRESL3    As String * 1
'''''    CRYINDL4    As String * 1
'''''    CRYRESL4    As String * 1
'''''    CRYINDCS    As String * 1
'''''    CRYRESCS    As String * 1
'''''    CRYINDGD    As String * 1
'''''    CRYRESGD    As String * 1
'''''    CRYINDT     As String * 1
'''''    CRYREST     As String * 1
'''''    CRYINDEP    As String * 1
'''''    CRYRESEP    As String * 1
'''''End Type
'''''#End If


'待ち一覧

'初期表示用
Public Type type_DBDRV_scmzc_fcmkc001b_Disp
    CRYNUM      As String * 12              ' 結晶番号
    IngotPos    As Integer                  ' 結晶内開始位置
'   LENGTH      As Integer                  ' 長さ              '2001/11/8
    BLOCKID     As String * 12              ' ブロックID
    HSXTYPE     As String * 1               ' 品ＳＸタイプ
    HSXCDIR     As String * 1               ' 品ＳＸ結晶面方位
    UPDDATE     As Date                     ' 更新日付
    Judg        As String                   ' 判定
    hin()       As tFullHinban              ' 品番(full)
    HOLDCLS     As String * 1               ' ホールド区分 ---kuramoto 追加 2001/09/25----
    SMP()       As type_cmkc001b_SmpMng     ' サンプル管理
End Type


''''''品番、仕様、結晶内側取得用 (TOP,TAIL順で２レコード取得)
'''''Public Type type_DBDRV_scmzc_fcmkc001c_Siyou
'''''    'ブロック管理
'''''    CRYNUM As String * 12             ' 結晶番号
'''''    INGOTPOS As Integer               ' 結晶内開始位置
'''''    LENGTH As Integer                 ' 長さ
'''''    '品番管理
'''''    hin As tFullHinban                ' 品番(full)
'''''
'''''        '結晶情報
'''''    PRODCOND As String * 4            ' 製作条件
'''''    PGID As String * 8                ' ＰＧ－ＩＤ
'''''    UPLENGTH As Integer               ' 引上げ長さ
'''''    FREELENG As Integer               ' フリー長
'''''    DIAMETER As Integer               ' 直径 2002/05/01 S.Sano
'''''    CHARGE As Double                  ' チャージ量
'''''    SEED As String * 4                ' シード
'''''    ADDDPPOS As Integer                 ' 追加ドープ位置
'''''
'''''    '製品仕様
'''''    HSXTYPE As String * 1             ' 品ＳＸタイプ
'''''    HSXD1CEN As Double                ' 品ＳＸ直径１中心
'''''    HSXCDIR As String * 1             ' 品ＳＸ結晶面方位
'''''    HSXRMIN As Double                 ' 品ＳＸ比抵抗下限
'''''    HSXRMAX As Double                 ' 品ＳＸ比抵抗上限
'''''    HSXRAMIN As Double                ' 品ＳＸ比抵抗平均下限
'''''    HSXRAMAX As Double                ' 品ＳＸ比抵抗平均上限
'''''    HSXRMCAL As String * 1            ' 品ＳＸ比抵抗面内計算　　　　'' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06
'''''    HSXRMBNP As Double                ' 品ＳＸ比抵抗面内分布
'''''    HSXRSPOH As String * 1            ' 品ＳＸ比抵抗測定位置＿方
'''''    HSXRSPOT As String * 1            ' 品ＳＸ比抵抗測定位置＿点
'''''    HSXRSPOI As String * 1            ' 品ＳＸ比抵抗測定位置＿位
'''''    HSXRHWYT As String * 1            ' 品ＳＸ比抵抗保証方法＿対
'''''    HSXRHWYS As String * 1            ' 品ＳＸ比抵抗保証方法＿処
'''''
'''''    HSXONMIN As Double                ' 品ＳＸ酸素濃度下限
'''''    HSXONMAX As Double                ' 品ＳＸ酸素濃度上限
'''''    HSXONAMN As Double                ' 品ＳＸ酸素濃度平均下限
'''''    HSXONAMX As Double                ' 品ＳＸ酸素濃度平均上限
'''''    HSXONMCL As String * 1            ' 品ＳＸ酸素濃度面内計算　　　　'' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06
'''''    HSXONMBP As Double                ' 品ＳＸ酸素濃度面内分布
'''''    HSXONSPH As String * 1            ' 品ＳＸ酸素濃度測定位置＿方
'''''    HSXONSPT As String * 1            ' 品ＳＸ酸素濃度測定位置＿点
'''''    HSXONSPI As String * 1            ' 品ＳＸ酸素濃度測定位置＿位
'''''    HSXONHWT As String * 1            ' 品ＳＸ酸素濃度保証方法＿対
'''''    HSXONHWS As String * 1            ' 品ＳＸ酸素濃度保証方法＿処
'''''
'''''    HSXBM1AN As Double                ' 品ＳＸＢＭＤ１平均下限
'''''    HSXBM1AX As Double                ' 品ＳＸＢＭＤ１平均上限
'''''    HSXBM2AN As Double                ' 品ＳＸＢＭＤ２平均下限
'''''    HSXBM2AX As Double                ' 品ＳＸＢＭＤ２平均上限
'''''    HSXBM3AN As Double                ' 品ＳＸＢＭＤ３平均下限
'''''    HSXBM3AX As Double                ' 品ＳＸＢＭＤ３平均上限
'''''    HSXBM1SH As String * 1            ' 品ＳＸＢＭＤ１測定位置＿方
'''''    HSXBM1ST As String * 1            ' 品ＳＸＢＭＤ１測定位置＿点
'''''    HSXBM1SR As String * 1            ' 品ＳＸＢＭＤ１測定位置＿領
'''''    HSXBM1HT As String * 1            ' 品ＳＸＢＭＤ１保証方法＿対
'''''    HSXBM1HS As String * 1            ' 品ＳＸＢＭＤ１保証方法＿処
'''''    HSXBM2SH As String * 1            ' 品ＳＸＢＭＤ２測定位置＿方
'''''    HSXBM2ST As String * 1            ' 品ＳＸＢＭＤ２測定位置＿点
'''''    HSXBM2SR As String * 1            ' 品ＳＸＢＭＤ２測定位置＿領
'''''    HSXBM2HT As String * 1            ' 品ＳＸＢＭＤ２保証方法＿対
'''''    HSXBM2HS As String * 1            ' 品ＳＸＢＭＤ２保証方法＿処
'''''    HSXBM3SH As String * 1            ' 品ＳＸＢＭＤ３測定位置＿方
'''''    HSXBM3ST As String * 1            ' 品ＳＸＢＭＤ３測定位置＿点
'''''    HSXBM3SR As String * 1            ' 品ＳＸＢＭＤ３測定位置＿領
'''''    HSXBM3HT As String * 1            ' 品ＳＸＢＭＤ３保証方法＿対
'''''    HSXBM3HS As String * 1            ' 品ＳＸＢＭＤ３保証方法＿処
'''''
'''''    HSXOS1AX As Double                ' 品ＳＸＯＳＦ１平均上限
'''''    HSXOS1MX As Double                ' 品ＳＸＯＳＦ１上限
'''''    HSXOS2AX As Double                ' 品ＳＸＯＳＦ２平均上限
'''''    HSXOS2MX As Double                ' 品ＳＸＯＳＦ２上限
'''''    HSXOS3AX As Double                ' 品ＳＸＯＳＦ３平均上限
'''''    HSXOS3MX As Double                ' 品ＳＸＯＳＦ３上限
'''''    HSXOS4AX As Double                ' 品ＳＸＯＳＦ４平均上限
'''''    HSXOS4MX As Double                ' 品ＳＸＯＳＦ４上限
'''''    HSXOS1SH As String * 1            ' 品ＳＸＯＳＦ１測定位置＿方
'''''    HSXOS1ST As String * 1            ' 品ＳＸＯＳＦ１測定位置＿点
'''''    HSXOS1SR As String * 1            ' 品ＳＸＯＳＦ１測定位置＿領
'''''    HSXOS1HT As String * 1            ' 品ＳＸＯＳＦ１保証方法＿対
'''''    HSXOS1HS As String * 1            ' 品ＳＸＯＳＦ１保証方法＿処
'''''    HSXOS2SH As String * 1            ' 品ＳＸＯＳＦ２測定位置＿方
'''''    HSXOS2ST As String * 1            ' 品ＳＸＯＳＦ２測定位置＿点
'''''    HSXOS2SR As String * 1            ' 品ＳＸＯＳＦ２測定位置＿領
'''''    HSXOS2HT As String * 1            ' 品ＳＸＯＳＦ２保証方法＿対
'''''    HSXOS2HS As String * 1            ' 品ＳＸＯＳＦ２保証方法＿処
'''''    HSXOS3SH As String * 1            ' 品ＳＸＯＳＦ３測定位置＿方
'''''    HSXOS3ST As String * 1            ' 品ＳＸＯＳＦ３測定位置＿点
'''''    HSXOS3SR As String * 1            ' 品ＳＸＯＳＦ３測定位置＿領
'''''    HSXOS3HT As String * 1            ' 品ＳＸＯＳＦ３保証方法＿対
'''''    HSXOS3HS As String * 1            ' 品ＳＸＯＳＦ３保証方法＿処
'''''    HSXOS4SH As String * 1            ' 品ＳＸＯＳＦ４測定位置＿方
'''''    HSXOS4ST As String * 1            ' 品ＳＸＯＳＦ４測定位置＿点
'''''    HSXOS4SR As String * 1            ' 品ＳＸＯＳＦ４測定位置＿領
'''''    HSXOS4HT As String * 1            ' 品ＳＸＯＳＦ４保証方法＿対
'''''    HSXOS4HS As String * 1            ' 品ＳＸＯＳＦ４保証方法＿処
'''''    HSXOS1NS As String * 2            ' 品ＳＸＯＳＦ１熱処理法
'''''    HSXOS2NS As String * 2            ' 品ＳＸＯＳＦ２熱処理法
'''''    HSXOS3NS As String * 2            ' 品ＳＸＯＳＦ３熱処理法
'''''    HSXOS4NS As String * 2            ' 品ＳＸＯＳＦ４熱処理法
'''''    HSXBM1NS As String * 2            ' 品ＳＸＢＭＤ１熱処理法
'''''    HSXBM2NS As String * 2            ' 品ＳＸＢＭＤ２熱処理法
'''''    HSXBM3NS As String * 2            ' 品ＳＸＢＭＤ３熱処理法
'''''
'''''    HSXCNMIN As Double                ' 品ＳＸ炭素濃度下限
'''''    HSXCNMAX As Double                ' 品ＳＸ炭素濃度上限
'''''    HSXCNSPH As String * 1            ' 品ＳＸ炭素濃度測定位置＿方
'''''    HSXCNSPT As String * 1            ' 品ＳＸ炭素濃度測定位置＿点
'''''    HSXCNSPI As String * 1            ' 品ＳＸ炭素濃度測定位置＿位
'''''    HSXCNHWT As String * 1            ' 品ＳＸ炭素濃度保証方法＿対
'''''    HSXCNHWS As String * 1            ' 品ＳＸ炭素濃度保証方法＿処
'''''
'''''    HSXDENMX As Integer               ' 品ＳＸＤｅｎ上限
'''''    HSXDENMN As Integer               ' 品ＳＸＤｅｎ下限
'''''    HSXLDLMX As Integer               ' 品ＳＸＬ／ＤＬ上限
'''''    HSXLDLMN As Integer               ' 品ＳＸＬ／ＤＬ下限
'''''    HSXDVDMX As Integer               ' 品ＳＸＤＶＤ２上限
'''''    HSXDVDMN As Integer               ' 品ＳＸＤＶＤ２下限
'''''    HSXDENHT As String * 1            ' 品ＳＸＤｅｎ保証方法＿対
'''''    HSXDENHS As String * 1            ' 品ＳＸＤｅｎ保証方法＿処
'''''    HSXLDLHT As String * 1            ' 品ＳＸＬ／ＤＬ保証方法＿対
'''''    HSXLDLHS As String * 1            ' 品ＳＸＬ／ＤＬ保証方法＿処
'''''    HSXDVDHT As String * 1            ' 品ＳＸＤＶＤ２保証方法＿対
'''''    HSXDVDHS As String * 1            ' 品ＳＸＤＶＤ２保証方法＿処
'''''    HSXDENKU As String * 1            ' 品ＳＸＤｅｎ検査有無
'''''    HSXDVDKU As String * 1            ' 品ＳＸＤＶＤ２検査有無
'''''    HSXLDLKU As String * 1            ' 品ＳＸＬ／ＤＬ検査有無
'''''
'''''    HSXLTMIN As Integer               ' 品ＳＸＬタイム下限
'''''    HSXLTMAX As Integer               ' 品ＳＸＬタイム上限
'''''    HSXLTSPH As String * 1            ' 品ＳＸＬタイム測定位置＿方
'''''    HSXLTSPT As String * 1            ' 品ＳＸＬタイム測定位置＿点
'''''    HSXLTSPI As String * 1            ' 品ＳＸＬタイム測定位置＿位
'''''    HSXLTHWT As String * 1            ' 品ＳＸＬタイム保証方法＿対
'''''    HSXLTHWS As String * 1            ' 品ＳＸＬタイム保証方法＿処
'''''    '結晶内側管理
'''''    EPDUP As Integer                  ' EPD　上限
'''''
'''''' 払出規制項目追加対応 yakimura 2002.12.01 start
'''''    TOPREG  As Integer                ' TOP規制
'''''    TAILREG As Double                 ' TAIL規制
'''''    BTMSPRT As Integer                ' ボトム析出規制
'''''' 払出規制項目追加対応 yakimura 2002.12.01 end
'''''
'''''' OSF，BMD項目追加対応  2002.04.02 yakimura
'''''    HSXOSF1PTK As String * 1          ' 品ＳＸＯＳＦ１パタン区分
'''''    HSXOSF2PTK As String * 1          ' 品ＳＸＯＳＦ２パタン区分
'''''    HSXOSF3PTK As String * 1          ' 品ＳＸＯＳＦ３パタン区分
'''''    HSXOSF4PTK As String * 1          ' 品ＳＸＯＳＦ４パタン区分
'''''    HSXBMD1MBP As Double              ' 品ＳＸＢＭＤ１面内分布
'''''    HSXBMD2MBP As Double              ' 品ＳＸＢＭＤ２面内分布
'''''    HSXBMD3MBP As Double              ' 品ＳＸＢＭＤ３面内分布
'''''' OSF，BMD項目追加対応  2002.04.02 yakimura
'''''End Type


'''''' 結晶サンプル管理取得用 (TOP,TAIL順で２レコード取得)
'''''Public Type type_DBDRV_scmzc_fcmkc001c_CrySmp
'''''    CRYNUM As String * 12             ' 結晶番号
'''''    INGOTPOS As Integer               ' 結晶内位置
'''''    LENGTH As Integer                 ' 長さ
'''''    BLOCKID As String * 12            ' ブロックID
'''''    SMPKBN As String * 1              ' サンプル区分
'''''    SMPLNO As Integer                 ' サンプルNo
'''''    hinban As String * 12             ' 品番
'''''    REVNUM As Integer                 ' 製品番号改訂番号
'''''    factory As String * 1             ' 工場
'''''    opecond As String * 1             ' 操業条件
'''''    KTKBN  As String * 1              ' 確定区分
'''''    CRYINDRS As String * 1            ' 結晶検査指示（Rs)
'''''    CRYINDOI As String * 1            ' 結晶検査指示（Oi)
'''''    CRYINDB1 As String * 1            ' 結晶検査指示（B1)
'''''    CRYINDB2 As String * 1            ' 結晶検査指示（B2）
'''''    CRYINDB3 As String * 1            ' 結晶検査指示（B3)
'''''    CRYINDL1 As String * 1            ' 結晶検査指示（L1)
'''''    CRYINDL2 As String * 1            ' 結晶検査指示（L2)
'''''    CRYINDL3 As String * 1            ' 結晶検査指示（L3)
'''''    CRYINDL4 As String * 1            ' 結晶検査指示（L4)
'''''    CRYINDCS As String * 1            ' 結晶検査指示（Cs)
'''''    CRYINDGD As String * 1            ' 結晶検査指示（GD)
'''''    CRYINDT As String * 1             ' 結晶検査指示（T)
'''''    CRYINDEP As String * 1            ' 結晶検査指示（EPD)
'''''End Type


''''''結晶抵抗実績
'''''Public Type type_DBDRV_scmzc_fcmkc001c_CryR
'''''    CRYNUM As String * 12             ' 結晶番号
'''''    POSITION As Integer               ' 位置
'''''    SMPKBN As String * 1              ' サンプル区分
'''''    SMPLNO As Integer                 ' サンプルＮｏ
'''''    SMPLUMU As String * 1             ' サンプル有無
'''''    TRANCOND As String * 1            ' 処理条件
'''''    MEAS1 As Double                   ' 測定値１
'''''    MEAS2 As Double                   ' 測定値２
'''''    MEAS3 As Double                   ' 測定値３
'''''    MEAS4 As Double                   ' 測定値４
'''''    MEAS5 As Double                   ' 測定値５
'''''    RRG As Double                     ' ＲＲＧ
'''''    REGDATE As Date                   ' 登録日付
'''''End Type


''''''Oi実績
'''''Public Type type_DBDRV_scmzc_fcmkc001c_Oi
'''''    CRYNUM As String * 12             ' 結晶番号
'''''    POSITION As Integer               ' 位置
'''''    SMPKBN As String * 1              ' サンプル区分
'''''    SMPLNO As Integer                 ' サンプルＮｏ
'''''    SMPLUMU As String * 1             ' サンプル有無
'''''    TRANCOND As String * 1            ' 処理条件
'''''    OIMEAS1 As Double                 ' Ｏｉ測定値１
'''''    OIMEAS2 As Double                 ' Ｏｉ測定値２
'''''    OIMEAS3 As Double                 ' Ｏｉ測定値３
'''''    OIMEAS4 As Double                 ' Ｏｉ測定値４
'''''    OIMEAS5 As Double                 ' Ｏｉ測定値５
'''''    ORGRES As Double                  ' ＯＲＧ結果
'''''    AVE As Double                     ' ＡＶＥ
'''''    FTIRCONV As Double                ' ＦＴＩＲ換算
'''''    INSPECTWAY As String * 2          ' 検査方法
'''''    REGDATE As Date                   ' 登録日付
'''''End Type
'''''
'''''
''''''BMD1～3実績
'''''Public Type type_DBDRV_scmzc_fcmkc001c_BMD
'''''    CRYNUM As String * 12             ' 結晶番号
'''''    POSITION As Integer               ' 位置
'''''    SMPKBN As String * 1              ' サンプル区分
'''''    SMPLNO As Integer                 ' サンプルＮｏ
'''''    SMPLUMU As String * 1             ' サンプル有無
'''''    HTPRC As String * 2               ' 熱処理方法
'''''    KKSP As String * 3                ' 結晶欠陥測定位置
'''''    KKSET As String * 3               ' 結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
'''''    TRANCOND As String * 1            ' 処理条件
'''''    MEAS1 As Double                   ' 測定値１
'''''    MEAS2 As Double                   ' 測定値２
'''''    MEAS3 As Double                   ' 測定値３
'''''    MEAS4 As Double                   ' 測定値４
'''''    MEAS5 As Double                   ' 測定値５
'''''    Min As Double                     ' MIN
'''''    max As Double                     ' MAX
'''''    AVE As Double                     ' AVE
'''''    REGDATE As Date                   ' 登録日付
'''''' OSF，BMD項目追加対応  2002.04.02 yakimura
'''''    BMDMNBUNP As Double               ' BMD面内分布
'''''' OSF，BMD項目追加対応  2002.04.02 yakimura
'''''End Type
'''''
'''''
''''''OSF1～4実績
'''''Public Type type_DBDRV_scmzc_fcmkc001c_OSF
'''''    CRYNUM As String * 12             ' 結晶番号
'''''    POSITION As Integer               ' 位置
'''''    SMPKBN As String * 1              ' サンプル区分
'''''    SMPLNO As Integer                 ' サンプルＮｏ
'''''    SMPLUMU As String * 1             ' サンプル有無
'''''    HTPRC As String * 2               ' 熱処理方法
'''''    KKSP As String * 3                ' 結晶欠陥測定位置
'''''    KKSET As String * 3               ' 結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
'''''    TRANCOND As String * 1            ' 処理条件
'''''    CALCMAX As Double                 ' 計算結果 Max
'''''    CALCAVE As Double                 ' 計算結果 Ave
'''''    MEAS1 As Double                   ' 測定値１
'''''    MEAS2 As Double                   ' 測定値２
'''''    MEAS3 As Double                   ' 測定値３
'''''    MEAS4 As Double                   ' 測定値４
'''''    MEAS5 As Double                   ' 測定値５
'''''    MEAS6 As Double                   ' 測定値６
'''''    MEAS7 As Double                   ' 測定値７
'''''    MEAS8 As Double                   ' 測定値８
'''''    MEAS9 As Double                   ' 測定値９
'''''    MEAS10 As Double                  ' 測定値１０
'''''    MEAS11 As Double                  ' 測定値１１
'''''    MEAS12 As Double                  ' 測定値１２
'''''    MEAS13 As Double                  ' 測定値１３
'''''    MEAS14 As Double                  ' 測定値１４
'''''    MEAS15 As Double                  ' 測定値１５
'''''    MEAS16 As Double                  ' 測定値１６
'''''    MEAS17 As Double                  ' 測定値１７
'''''    MEAS18 As Double                  ' 測定値１８
'''''    MEAS19 As Double                  ' 測定値１９
'''''    MEAS20 As Double                  ' 測定値２０
'''''    REGDATE As Date                   ' 登録日付
'''''' OSF，BMD項目追加対応  2002.04.02 yakimura
'''''    OSFPOS1    As Double              ' ﾊﾟﾀｰﾝ区分１位置
'''''    OSFWID1    As Double              ' ﾊﾟﾀｰﾝ区分１幅
'''''    OSFRD1     As String * 1          ' ﾊﾟﾀｰﾝ区分１R/D
'''''    OSFPOS2    As Double              ' ﾊﾟﾀｰﾝ区分２位置
'''''    OSFWID2    As Double              ' ﾊﾟﾀｰﾝ区分２幅
'''''    OSFRD2     As String * 1          ' ﾊﾟﾀｰﾝ区分２R/D
'''''    OSFPOS3    As Double              ' ﾊﾟﾀｰﾝ区分３位置
'''''    OSFWID3    As Double              ' ﾊﾟﾀｰﾝ区分３幅
'''''    OSFRD3     As String * 1          ' ﾊﾟﾀｰﾝ区分３R/D
'''''' OSF，BMD項目追加対応  2002.04.02 yakimura
'''''End Type
'''''
'''''
''''''CS実績
'''''Public Type type_DBDRV_scmzc_fcmkc001c_CS
'''''    CRYNUM As String * 12             ' 結晶番号
'''''    POSITION As Integer               ' 位置
'''''    SMPKBN As String * 1              ' サンプル区分
'''''    SMPLNO As Integer                 ' サンプルＮｏ
'''''    SMPLUMU As String * 1             ' サンプル有無
'''''    TRANCOND As String * 1            ' 処理条件
'''''    CSMEAS As Double                  ' Cs実測値
'''''    PRE70P As Double                  ' ７０％推定値
'''''    REGDATE As Date                   ' 登録日付
'''''End Type
'''''
'''''
''''''GD実績
'''''Public Type type_DBDRV_scmzc_fcmkc001c_GD
'''''    CRYNUM As String * 12             ' 結晶番号
'''''    POSITION As Integer               ' 位置
'''''    SMPKBN As String * 1              ' サンプル区分
'''''    SMPLNO As Integer                 ' サンプルＮｏ
'''''    SMPLUMU As String * 1             ' サンプル有無
'''''    TRANCOND As String * 1            ' 処理条件
'''''    MSRSDEN As Integer                ' 測定結果 Den
'''''    MSRSLDL As Integer                ' 測定結果 L/DL
'''''    MSRSDVD2 As Integer               ' 測定結果 DVD2
'''''    MS01LDL1 As Integer             ' 測定値01 L/DL1
'''''    MS01LDL2 As Integer             ' 測定値01 L/DL2
'''''    MS01LDL3 As Integer             ' 測定値01 L/DL3
'''''    MS01LDL4 As Integer             ' 測定値01 L/DL4
'''''    MS01LDL5 As Integer             ' 測定値01 L/DL5
'''''    MS01DEN1 As Integer             ' 測定値01 Den1
'''''    MS01DEN2 As Integer             ' 測定値01 Den2
'''''    MS01DEN3 As Integer             ' 測定値01 Den3
'''''    MS01DEN4 As Integer             ' 測定値01 Den4
'''''    MS01DEN5 As Integer             ' 測定値01 Den5
'''''    MS02LDL1 As Integer             ' 測定値02 L/DL1
'''''    MS02LDL2 As Integer             ' 測定値02 L/DL2
'''''    MS02LDL3 As Integer             ' 測定値02 L/DL3
'''''    MS02LDL4 As Integer             ' 測定値02 L/DL4
'''''    MS02LDL5 As Integer             ' 測定値02 L/DL5
'''''    MS02DEN1 As Integer             ' 測定値02 Den1
'''''    MS02DEN2 As Integer             ' 測定値02 Den2
'''''    MS02DEN3 As Integer             ' 測定値02 Den3
'''''    MS02DEN4 As Integer             ' 測定値02 Den4
'''''    MS02DEN5 As Integer             ' 測定値02 Den5
'''''    MS03LDL1 As Integer             ' 測定値03 L/DL1
'''''    MS03LDL2 As Integer             ' 測定値03 L/DL2
'''''    MS03LDL3 As Integer             ' 測定値03 L/DL3
'''''    MS03LDL4 As Integer             ' 測定値03 L/DL4
'''''    MS03LDL5 As Integer             ' 測定値03 L/DL5
'''''    MS03DEN1 As Integer             ' 測定値03 Den1
'''''    MS03DEN2 As Integer             ' 測定値03 Den2
'''''    MS03DEN3 As Integer             ' 測定値03 Den3
'''''    MS03DEN4 As Integer             ' 測定値03 Den4
'''''    MS03DEN5 As Integer             ' 測定値03 Den5
'''''    MS04LDL1 As Integer             ' 測定値04 L/DL1
'''''    MS04LDL2 As Integer             ' 測定値04 L/DL2
'''''    MS04LDL3 As Integer             ' 測定値04 L/DL3
'''''    MS04LDL4 As Integer             ' 測定値04 L/DL4
'''''    MS04LDL5 As Integer             ' 測定値04 L/DL5
'''''    MS04DEN1 As Integer             ' 測定値04 Den1
'''''    MS04DEN2 As Integer             ' 測定値04 Den2
'''''    MS04DEN3 As Integer             ' 測定値04 Den3
'''''    MS04DEN4 As Integer             ' 測定値04 Den4
'''''    MS04DEN5 As Integer             ' 測定値04 Den5
'''''    MS05LDL1 As Integer             ' 測定値05 L/DL1
'''''    MS05LDL2 As Integer             ' 測定値05 L/DL2
'''''    MS05LDL3 As Integer             ' 測定値05 L/DL3
'''''    MS05LDL4 As Integer             ' 測定値05 L/DL4
'''''    MS05LDL5 As Integer             ' 測定値05 L/DL5
'''''    MS05DEN1 As Integer             ' 測定値05 Den1
'''''    MS05DEN2 As Integer             ' 測定値05 Den2
'''''    MS05DEN3 As Integer             ' 測定値05 Den3
'''''    MS05DEN4 As Integer             ' 測定値05 Den4
'''''    MS05DEN5 As Integer             ' 測定値05 Den5
'''''    MS06LDL1 As Integer             ' 測定値06 L/DL1
'''''    MS06LDL2 As Integer             ' 測定値06 L/DL2
'''''    MS06LDL3 As Integer             ' 測定値06 L/DL3
'''''    MS06LDL4 As Integer             ' 測定値06 L/DL4
'''''    MS06LDL5 As Integer             ' 測定値06 L/DL5
'''''    MS06DEN1 As Integer             ' 測定値06 Den1
'''''    MS06DEN2 As Integer             ' 測定値06 Den2
'''''    MS06DEN3 As Integer             ' 測定値06 Den3
'''''    MS06DEN4 As Integer             ' 測定値06 Den4
'''''    MS06DEN5 As Integer             ' 測定値06 Den5
'''''    MS07LDL1 As Integer             ' 測定値07 L/DL1
'''''    MS07LDL2 As Integer             ' 測定値07 L/DL2
'''''    MS07LDL3 As Integer             ' 測定値07 L/DL3
'''''    MS07LDL4 As Integer             ' 測定値07 L/DL4
'''''    MS07LDL5 As Integer             ' 測定値07 L/DL5
'''''    MS07DEN1 As Integer             ' 測定値07 Den1
'''''    MS07DEN2 As Integer             ' 測定値07 Den2
'''''    MS07DEN3 As Integer             ' 測定値07 Den3
'''''    MS07DEN4 As Integer             ' 測定値07 Den4
'''''    MS07DEN5 As Integer             ' 測定値07 Den5
'''''    MS08LDL1 As Integer             ' 測定値08 L/DL1
'''''    MS08LDL2 As Integer             ' 測定値08 L/DL2
'''''    MS08LDL3 As Integer             ' 測定値08 L/DL3
'''''    MS08LDL4 As Integer             ' 測定値08 L/DL4
'''''    MS08LDL5 As Integer             ' 測定値08 L/DL5
'''''    MS08DEN1 As Integer             ' 測定値08 Den1
'''''    MS08DEN2 As Integer             ' 測定値08 Den2
'''''    MS08DEN3 As Integer             ' 測定値08 Den3
'''''    MS08DEN4 As Integer             ' 測定値08 Den4
'''''    MS08DEN5 As Integer             ' 測定値08 Den5
'''''    MS09LDL1 As Integer             ' 測定値09 L/DL1
'''''    MS09LDL2 As Integer             ' 測定値09 L/DL2
'''''    MS09LDL3 As Integer             ' 測定値09 L/DL3
'''''    MS09LDL4 As Integer             ' 測定値09 L/DL4
'''''    MS09LDL5 As Integer             ' 測定値09 L/DL5
'''''    MS09DEN1 As Integer             ' 測定値09 Den1
'''''    MS09DEN2 As Integer             ' 測定値09 Den2
'''''    MS09DEN3 As Integer             ' 測定値09 Den3
'''''    MS09DEN4 As Integer             ' 測定値09 Den4
'''''    MS09DEN5 As Integer             ' 測定値09 Den5
'''''    MS10LDL1 As Integer             ' 測定値10 L/DL1
'''''    MS10LDL2 As Integer             ' 測定値10 L/DL2
'''''    MS10LDL3 As Integer             ' 測定値10 L/DL3
'''''    MS10LDL4 As Integer             ' 測定値10 L/DL4
'''''    MS10LDL5 As Integer             ' 測定値10 L/DL5
'''''    MS10DEN1 As Integer             ' 測定値10 Den1
'''''    MS10DEN2 As Integer             ' 測定値10 Den2
'''''    MS10DEN3 As Integer             ' 測定値10 Den3
'''''    MS10DEN4 As Integer             ' 測定値10 Den4
'''''    MS10DEN5 As Integer             ' 測定値10 Den5
'''''    MS11LDL1 As Integer             ' 測定値11 L/DL1
'''''    MS11LDL2 As Integer             ' 測定値11 L/DL2
'''''    MS11LDL3 As Integer             ' 測定値11 L/DL3
'''''    MS11LDL4 As Integer             ' 測定値11 L/DL4
'''''    MS11LDL5 As Integer             ' 測定値11 L/DL5
'''''    MS11DEN1 As Integer             ' 測定値11 Den1
'''''    MS11DEN2 As Integer             ' 測定値11 Den2
'''''    MS11DEN3 As Integer             ' 測定値11 Den3
'''''    MS11DEN4 As Integer             ' 測定値11 Den4
'''''    MS11DEN5 As Integer             ' 測定値11 Den5
'''''    MS12LDL1 As Integer             ' 測定値12 L/DL1
'''''    MS12LDL2 As Integer             ' 測定値12 L/DL2
'''''    MS12LDL3 As Integer             ' 測定値12 L/DL3
'''''    MS12LDL4 As Integer             ' 測定値12 L/DL4
'''''    MS12LDL5 As Integer             ' 測定値12 L/DL5
'''''    MS12DEN1 As Integer             ' 測定値12 Den1
'''''    MS12DEN2 As Integer             ' 測定値12 Den2
'''''    MS12DEN3 As Integer             ' 測定値12 Den3
'''''    MS12DEN4 As Integer             ' 測定値12 Den4
'''''    MS12DEN5 As Integer             ' 測定値12 Den5
'''''    MS13LDL1 As Integer             ' 測定値13 L/DL1
'''''    MS13LDL2 As Integer             ' 測定値13 L/DL2
'''''    MS13LDL3 As Integer             ' 測定値13 L/DL3
'''''    MS13LDL4 As Integer             ' 測定値13 L/DL4
'''''    MS13LDL5 As Integer             ' 測定値13 L/DL5
'''''    MS13DEN1 As Integer             ' 測定値13 Den1
'''''    MS13DEN2 As Integer             ' 測定値13 Den2
'''''    MS13DEN3 As Integer             ' 測定値13 Den3
'''''    MS13DEN4 As Integer             ' 測定値13 Den4
'''''    MS13DEN5 As Integer             ' 測定値13 Den5
'''''    MS14LDL1 As Integer             ' 測定値14 L/DL1
'''''    MS14LDL2 As Integer             ' 測定値14 L/DL2
'''''    MS14LDL3 As Integer             ' 測定値14 L/DL3
'''''    MS14LDL4 As Integer             ' 測定値14 L/DL4
'''''    MS14LDL5 As Integer             ' 測定値14 L/DL5
'''''    MS14DEN1 As Integer             ' 測定値14 Den1
'''''    MS14DEN2 As Integer             ' 測定値14 Den2
'''''    MS14DEN3 As Integer             ' 測定値14 Den3
'''''    MS14DEN4 As Integer             ' 測定値14 Den4
'''''    MS14DEN5 As Integer             ' 測定値14 Den5
'''''    MS15LDL1 As Integer             ' 測定値15 L/DL1
'''''    MS15LDL2 As Integer             ' 測定値15 L/DL2
'''''    MS15LDL3 As Integer             ' 測定値15 L/DL3
'''''    MS15LDL4 As Integer             ' 測定値15 L/DL4
'''''    MS15LDL5 As Integer             ' 測定値15 L/DL5
'''''    MS15DEN1 As Integer             ' 測定値15 Den1
'''''    MS15DEN2 As Integer             ' 測定値15 Den2
'''''    MS15DEN3 As Integer             ' 測定値15 Den3
'''''    MS15DEN4 As Integer             ' 測定値15 Den4
'''''    MS15DEN5 As Integer             ' 測定値15 Den5
'''''    REGDATE As Date                   ' 登録日付
'''''End Type
'''''
'''''
''''''ライフタイム実績取得関数
'''''Public Type type_DBDRV_scmzc_fcmkc001c_LT
'''''    CRYNUM As String * 12             ' 結晶番号
'''''    POSITION As Integer               ' 位置
'''''    SMPKBN As String * 1              ' サンプル区分
'''''    SMPLNO As Integer                 ' サンプルＮｏ
'''''    SMPLUMU As String * 1             ' サンプル有無
'''''    MEAS1 As Integer                  ' 測定値１
'''''    MEAS2 As Integer                  ' 測定値２
'''''    MEAS3 As Integer                  ' 測定値３
'''''    MEAS4 As Integer                  ' 測定値４
'''''    MEAS5 As Integer                  ' 測定値５
'''''    TRANCOND As String * 1            ' 処理条件
'''''    MEASPEAK As Integer               ' 測定値 ピーク値
'''''    CALCMEAS As Integer               ' 計算結果
'''''    REGDATE As Date                   ' 登録日付
'''''    LTSPI As String                 '測定位置コード
'''''End Type
'''''
'''''
''''''EPD実績取得関数
'''''Public Type type_DBDRV_scmzc_fcmkc001c_EPD
'''''    CRYNUM As String * 12             ' 結晶番号
'''''    POSITION As Integer               ' 位置
'''''    SMPKBN As String * 1              ' サンプル区分
'''''    SMPLNO As Integer                 ' サンプルＮｏ
'''''    SMPLUMU As String * 1             ' サンプル有無
'''''    TRANCOND As String * 1            ' 処理条件
'''''    MEASURE As Integer                ' 測定値
'''''    REGDATE As Date                   ' 登録日付
'''''End Type


''''''実績をまとめた構造体
'''''Public Type type_DBDRV_scmzc_fcmkc001c_Zisseki
'''''    CRYRZ() As type_DBDRV_scmzc_fcmkc001c_CryR
'''''    OIZ() As type_DBDRV_scmzc_fcmkc001c_Oi
'''''    BMD1Z() As type_DBDRV_scmzc_fcmkc001c_BMD
'''''    BMD2Z() As type_DBDRV_scmzc_fcmkc001c_BMD
'''''    BMD3Z() As type_DBDRV_scmzc_fcmkc001c_BMD
'''''    OSF1Z() As type_DBDRV_scmzc_fcmkc001c_OSF
'''''    OSF2Z() As type_DBDRV_scmzc_fcmkc001c_OSF
'''''    OSF3Z() As type_DBDRV_scmzc_fcmkc001c_OSF
'''''    OSF4Z() As type_DBDRV_scmzc_fcmkc001c_OSF
'''''    csz() As type_DBDRV_scmzc_fcmkc001c_CS
'''''    GDZ() As type_DBDRV_scmzc_fcmkc001c_GD
'''''    LTZ() As type_DBDRV_scmzc_fcmkc001c_LT
'''''    EPDZ() As type_DBDRV_scmzc_fcmkc001c_EPD
'''''    SURSZ() As type_DBDRV_scmzc_fcmkc001c_CryR
'''''End Type


'ブロック管理更新用（現在工程、最終通過工程）
Public Type type_DBDRV_scmzc_fcmkc001c_UpdBlock1
    CRYNUM      As String * 12          ' 結晶番号
    IngotPos    As Integer              ' 結晶内開始位置
    NOWPROC     As String * 5           ' 現在工程
    LASTPASS    As String * 5           ' 最終通過工程
End Type



'ブロック管理更新用（クリスタルカタログ、リメルト用）
Public Type typ_DBDRV_fcmkc001c_UpdBlkCR
    CRYNUM      As String * 12          ' 結晶番号
    IngotPos    As Integer              ' 結晶内開始位置
    NOWPROC     As String * 5           ' 現在工程
'   LASTPASS    As String * 5           ' 最終通過工程
    DELCLS      As String * 1           ' 削除区分
    BDCAUS      As String * 3           ' 不良理由
    LSTATCLS    As String * 1           ' 最終状態区分
    RSTATCLS    As String * 1           ' 流動状態区分
End Type



'結晶サンプル管理更新用
Public Type type_DBDRV_scmzc_fcmkc001c_UpdCrySmp
    CRYNUM      As String * 12          ' 結晶番号
    IngotPos    As Integer              ' 結晶内位置
    SMPKBN      As String * 1           ' サンプル区分
End Type


''''''測定結果のJ014書込要否構造体
'''''Public Type Judg_Spec_Cry
'''''    Enable As Boolean           '有効な品番である
'''''    rs As Boolean               'Rsは要書込
'''''    Oi As Boolean               'Oiは要書込
'''''    B1 As Boolean               'BMD1は要書込
'''''    B2 As Boolean               'BMD2は要書込
'''''    B3 As Boolean               'BMD3は要書込
'''''    L1 As Boolean               'OSF1は要書込
'''''    L2 As Boolean               'OSF2は要書込
'''''    L3 As Boolean               'OSF3は要書込
'''''    L4 As Boolean               'OSF4は要書込
'''''    Cs As Boolean               'Csは要書込
'''''    GD As Boolean               'GDは要書込
'''''    Lt As Boolean               'LTは要書込
'''''    EPD As Boolean              'EPDは要書込
'''''End Type


'''''' 仕様の指示がたっている判断用
'''''Public Const SIJI = "H"
'''''Public Const SANKOU = "S"




''''''内部関数 ブロックID、更新日付取得（払出待ち、抜試指示待ち用）
'''''Private Function getBlockID(records() As type_DBDRV_scmzc_fcmkc001b_Disp, _
'''''                            NOWPROC As String) As FUNCTION_RETURN
'''''
'''''    Dim sql         As String       'SQL全体
'''''    Dim rs          As OraDynaset   'RecordSet
'''''    Dim recCnt      As Long         'レコード数
'''''    Dim i           As Long
'''''    Dim j           As Long
'''''    Dim k           As Long
'''''    Dim BlockIdBuf  As String
'''''
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function getBlockID"
'''''
'''''    getBlockID = FUNCTION_RETURN_SUCCESS
'''''
'''''    sql = "select V.E040CRYNUM, "
'''''    sql = sql & " V.E040BLOCKID, "
'''''    sql = sql & " V.E040INGOTPOS, "
'''''    sql = sql & " V.E040UPDDATE, "
'''''    sql = sql & " V.E040HOLDCLS, "
'''''    sql = sql & " V.E041HINBAN, "           ' 品番
'''''    sql = sql & " V.E041REVNUM, "           ' 製品番号改訂番号
'''''    sql = sql & " V.E041FACTORY, "          ' 工場
'''''    sql = sql & " V.E041OPECOND, "          ' 操業条件
'''''    sql = sql & " S.HSXTYPE, "              ' 品ＳＸタイプ
'''''    sql = sql & " S.HSXCDIR "               ' 品ＳＸ結晶面方位
'''''    sql = sql & " from VECME009 V, TBCME018 S "
'''''    sql = sql & " where V.E041HINBAN  = S.HINBAN "
'''''    sql = sql & "   and V.E041REVNUM  = S.MNOREVNO "
'''''    sql = sql & "   and V.E041FACTORY = S.FACTORY "
'''''    sql = sql & "   and V.E041OPECOND = S.OPECOND "
'''''    sql = sql & "   and V.E040NOWPROC ='" & NOWPROC & "' "
'''''    sql = sql & "   and V.E040LSTATCLS='T' "
'''''    sql = sql & "   and V.E040RSTATCLS='T' "
'''''    sql = sql & "   and V.E040DELCLS  ='0' "
'''''   'sql = sql & "   and V.E040HOLDCLS ='0' " ' ホールドブロックも取得
'''''    sql = sql & " order by V.E040BLOCKID, V.E041INGOTPOS "
'''''
'''''    'データを抽出する
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''    'レコードがない場合正常終了
'''''    If rs.RecordCount = 0 Then
'''''        rs.Close
'''''        ReDim records(0)
'''''        GoTo proc_exit
'''''    End If
'''''
'''''    BlockIdBuf = vbNullString
'''''    recCnt = rs.RecordCount
'''''    j = 0
'''''    For i = 1 To recCnt
'''''        DoEvents
'''''        'ブロックID等の格納
'''''        If rs("E040BLOCKID") <> BlockIdBuf Then
'''''
'''''            j = j + 1
'''''            ReDim Preserve records(j)
'''''
'''''            With records(j)
'''''                .CRYNUM = rs("E040CRYNUM")
'''''                .IngotPos = rs("E040INGOTPOS")
'''''                .BLOCKID = rs("E040BLOCKID")   ' ブロックID
'''''                .UPDDATE = rs("E040UPDDATE")   ' 更新日付
'''''                .HOLDCLS = rs("E040HOLDCLS")   ' ホールド区分
'''''                BlockIdBuf = records(j).BLOCKID
'''''                .HSXTYPE = rs("HSXTYPE")
'''''                .HSXCDIR = rs("HSXCDIR")
'''''                .Judg = " "
'''''            End With
'''''
'''''            k = 1
'''''        End If
'''''
'''''        '品番の格納
'''''        ReDim Preserve records(j).hin(k)
'''''        records(j).hin(k).hinban = rs("E041HINBAN")
'''''        records(j).hin(k).mnorevno = rs("E041REVNUM")
'''''        records(j).hin(k).factory = rs("E041FACTORY")
'''''        records(j).hin(k).opecond = rs("E041OPECOND")
'''''        k = k + 1
'''''        rs.MoveNext
'''''    Next i
'''''    rs.Close
'''''
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    getBlockID = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


''''''購入単結晶用
'''''Private Function getKouBlock(records() As type_DBDRV_scmzc_fcmkc001b_Disp, NOWPROC As String) As FUNCTION_RETURN
'''''
'''''    Dim sql         As String       'SQL全体
'''''    Dim rs          As OraDynaset   'RecordSet
'''''    Dim recCnt      As Long
'''''    Dim motoRecCnt  As Long
'''''    Dim i           As Long
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
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
'''''    sql = sql & " from  TBCME040 B,TBCMG002 K "
'''''    sql = sql & " where B.BLOCKID=K.CRYNUM "
'''''    sql = sql & "   and substr(B.BLOCKID,1,1)='8' "
'''''    sql = sql & "   and B.NOWPROC ='" & NOWPROC & "' "
'''''    sql = sql & "   and B.LSTATCLS='T' "
'''''    sql = sql & "   and B.RSTATCLS='T' "
'''''    sql = sql & "   and B.DELCLS  ='0' "
'''''   'sql = sql & "   and B.HOLDCLS ='0' " ' ホールドブロックも取得
'''''    sql = sql & "   and K.TRANCNT =any(select max(TRANCNT) from TBCMG002 where CRYNUM=B.BLOCKID) "
'''''    sql = sql & " order by B.BLOCKID "
'''''
'''''
'''''    'データを抽出する
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''    If rs.RecordCount = 0 Then
'''''        rs.Close
'''''        GoTo proc_exit
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
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    getKouBlock = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''
'''''End Function



Public Function DBDRV_scmzc_fcmkc001b_Disp00(record0() As type_DBDRV_scmzc_fcmkc001b_Disp, _
                                             record1() As type_DBDRV_scmzc_fcmkc001b_Disp, _
                                             LWD() As cmkc001b_LockWait) As FUNCTION_RETURN

    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim recCnt      As Long         'ブロック管理のレコード数
    Dim i           As Long
    
    Dim j1          As Long
    Dim k1          As Long
    Dim j2          As Long
    Dim k2          As Long
    
    Dim BlockIdBuf1 As String
    Dim BlockIdBuf2 As String
    
    '<検査待ち>
    '<判定待ち>
    
    'ブロック管理テーブルからブロックID、更新日付取得（検査実績が未検査のもの）
    

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp1"
    DBDRV_scmzc_fcmkc001b_Disp00 = FUNCTION_RETURN_SUCCESS

    'ブロックID、更新日付の取得
    sql = "select distinct "
    
    sql = sql & " X.XTALCA       as CRYNUM,"
    sql = sql & " B.INGOTPOS,"
    sql = sql & " X.CRYNUMCA     as BLOCKID,"
    sql = sql & " B.UPDDATE,"
    sql = sql & " B.HOLDCLS,"
    sql = sql & " X.HINBCA       as HINBAN,"        ' 品番
    sql = sql & " X.REVNUMCA     as REVNUM,"        ' 製品番号改訂番号
    sql = sql & " X.FACTORYCA    as FACTORY,"       ' 工場
    sql = sql & " X.OPECA        as OPECOND,"       ' 操業条件
    sql = sql & " S.HSXTYPE,"                       ' 品ＳＸタイプ
    sql = sql & " S.HSXCDIR,"                       ' 品ＳＸ結晶面方位
    sql = sql & " X.INPOSCA,"
    
    sql = sql & " XT.CRYINDRSCS  as T_CRYINDRSCS,"  ' 上サンプル
    sql = sql & " XT.CRYINDOICS  as T_CRYINDOICS,"
    sql = sql & " XT.CRYINDB1CS  as T_CRYINDB1CS,"
    sql = sql & " XT.CRYINDB2CS  as T_CRYINDB2CS,"
    sql = sql & " XT.CRYINDB3CS  as T_CRYINDB3CS,"
    sql = sql & " XT.CRYINDL1CS  as T_CRYINDL1CS,"
    sql = sql & " XT.CRYINDL2CS  as T_CRYINDL2CS,"
    sql = sql & " XT.CRYINDL3CS  as T_CRYINDL3CS,"
    sql = sql & " XT.CRYINDL4CS  as T_CRYINDL4CS,"
    sql = sql & " XT.CRYINDCSCS  as T_CRYINDCSCS,"
    sql = sql & " XT.CRYINDGDCS  as T_CRYINDGDCS,"
    sql = sql & " XT.CRYINDTCS   as T_CRYINDT_CS,"
    sql = sql & " XT.CRYINDEPCS  as T_CRYINDEPCS,"
    sql = sql & " XT.CRYRESRS1CS as T_CRYRESR1CS,"
    sql = sql & " XT.CRYRESRS2CS as T_CRYRESR2CS,"
    sql = sql & " XT.CRYRESOICS  as T_CRYRESOICS,"
    sql = sql & " XT.CRYRESB1CS  as T_CRYRESB1CS,"
    sql = sql & " XT.CRYRESB2CS  as T_CRYRESB2CS,"
    sql = sql & " XT.CRYRESB3CS  as T_CRYRESB3CS,"
    sql = sql & " XT.CRYRESL1CS  as T_CRYRESL1CS,"
    sql = sql & " XT.CRYRESL2CS  as T_CRYRESL2CS,"
    sql = sql & " XT.CRYRESL3CS  as T_CRYRESL3CS,"
    sql = sql & " XT.CRYRESL4CS  as T_CRYRESL4CS,"
    sql = sql & " XT.CRYRESCSCS  as T_CRYRESCSCS,"
    sql = sql & " XT.CRYRESGDCS  as T_CRYRESGDCS,"
    sql = sql & " XT.CRYRESTCS   as T_CRYREST_CS,"
    sql = sql & " XT.CRYRESEPCS  as T_CRYRESEPCS,"
                                    
    sql = sql & " XB.CRYINDRSCS  as B_CRYINDRSCS,"  ' 下サンプル
    sql = sql & " XB.CRYINDOICS  as B_CRYINDOICS,"
    sql = sql & " XB.CRYINDB1CS  as B_CRYINDB1CS,"
    sql = sql & " XB.CRYINDB2CS  as B_CRYINDB2CS,"
    sql = sql & " XB.CRYINDB3CS  as B_CRYINDB3CS,"
    sql = sql & " XB.CRYINDL1CS  as B_CRYINDL1CS,"
    sql = sql & " XB.CRYINDL2CS  as B_CRYINDL2CS,"
    sql = sql & " XB.CRYINDL3CS  as B_CRYINDL3CS,"
    sql = sql & " XB.CRYINDL4CS  as B_CRYINDL4CS,"
    sql = sql & " XB.CRYINDCSCS  as B_CRYINDCSCS,"
    sql = sql & " XB.CRYINDGDCS  as B_CRYINDGDCS,"
    sql = sql & " XB.CRYINDTCS   as B_CRYINDT_CS,"
    sql = sql & " XB.CRYINDEPCS  as B_CRYINDEPCS,"
    sql = sql & " XB.CRYRESRS1CS as B_CRYRESR1CS,"
    sql = sql & " XB.CRYRESRS2CS as B_CRYRESR2CS,"
    sql = sql & " XB.CRYRESOICS  as B_CRYRESOICS,"
    sql = sql & " XB.CRYRESB1CS  as B_CRYRESB1CS,"
    sql = sql & " XB.CRYRESB2CS  as B_CRYRESB2CS,"
    sql = sql & " XB.CRYRESB3CS  as B_CRYRESB3CS,"
    sql = sql & " XB.CRYRESL1CS  as B_CRYRESL1CS,"
    sql = sql & " XB.CRYRESL2CS  as B_CRYRESL2CS,"
    sql = sql & " XB.CRYRESL3CS  as B_CRYRESL3CS,"
    sql = sql & " XB.CRYRESL4CS  as B_CRYRESL4CS,"
    sql = sql & " XB.CRYRESCSCS  as B_CRYRESCSCS,"
    sql = sql & " XB.CRYRESGDCS  as B_CRYRESGDCS,"
    sql = sql & " XB.CRYRESTCS   as B_CRYREST_CS,"
    sql = sql & " XB.CRYRESEPCS  as B_CRYRESEPCS,"
    
    sql = sql & " (select count(*) From XSDCS X2"                                   ' 指示が実測(1)で実績が無し(0)が１カ所でもあれば検査待ち
    sql = sql & "   where X2.CRYNUMCS= X.CRYNUMCA"                                  '
    sql = sql & "     and X2.LIVKCS  ='0'"                                          ' 生死区分
    sql = sql & "     and ((X2.CRYINDRSCS='1' and X2.CRYRESRS1CS='0')"              ' 結晶検査実績（Rs)
    sql = sql & "      or  (X2.CRYINDOICS='1' and X2.CRYRESOICS ='0')"              ' 結晶検査実績（Oi)
    sql = sql & "      or  (X2.CRYINDB1CS='1' and X2.CRYRESB1CS ='0')"              ' 結晶検査実績（B1)
    sql = sql & "      or  (X2.CRYINDB2CS='1' and X2.CRYRESB2CS ='0')"              ' 結晶検査実績（B2）
    sql = sql & "      or  (X2.CRYINDB3CS='1' and X2.CRYRESB3CS ='0')"              ' 結晶検査実績（B3)
    sql = sql & "      or  (X2.CRYINDL1CS='1' and X2.CRYRESL1CS ='0')"              ' 結晶検査実績（L1)
    sql = sql & "      or  (X2.CRYINDL2CS='1' and X2.CRYRESL2CS ='0')"              ' 結晶検査実績（L2)
    sql = sql & "      or  (X2.CRYINDL3CS='1' and X2.CRYRESL3CS ='0')"              ' 結晶検査実績（L3)
    sql = sql & "      or  (X2.CRYINDL4CS='1' and X2.CRYRESL4CS ='0')"              ' 結晶検査実績（L4)
    sql = sql & "      or  (X2.CRYINDCSCS='1' and X2.CRYRESCSCS ='0')"              ' 結晶検査実績（Cs)
    sql = sql & "      or  (X2.CRYINDGDCS='1' and X2.CRYRESGDCS ='0')"              ' 結晶検査実績（GD)
    sql = sql & "      or  (X2.CRYINDTCS ='1' and X2.CRYRESTCS  ='0')"              ' 結晶検査実績（T)
    sql = sql & "      or  (X2.CRYINDEPCS='1' and X2.CRYRESEPCS ='0')) ) as DTTYPE" ' 結晶検査実績（EPD)
    
    sql = sql & " from  XSDCA X, TBCME040 B, TBCME018 S, XSDCS XT, XSDCS XB"
    
    sql = sql & " where X.CRYNUMCA = B.BLOCKID "
    sql = sql & "   and X.GNWKNTCA ='CC600' "
    sql = sql & "   and X.LSTATBCA ='T' "
''''sql = sql & "   and X.RSTATBCA ='T' "           '' 格上げ格下げで表示されなくなる（？）のでコメント
    sql = sql & "   and X.LIVKCA   ='0' "
    sql = sql & "   and B.DELCLS   ='0' "
   'sql = sql & "   and B.HOLDCLS  ='0' "           ' ホールドブロックも取得
    
    sql = sql & "   and X.HINBCA   = S.HINBAN "
    sql = sql & "   and X.REVNUMCA = S.MNOREVNO "
    sql = sql & "   and X.FACTORYCA= S.FACTORY "
    sql = sql & "   and X.OPECA    = S.OPECOND "
                                                    ' サンプルは必ず上下１件づつ存在する事
    sql = sql & "   and XT.CRYNUMCS= X.CRYNUMCA"    ' 上サンプル条件
    sql = sql & "   and XT.TBKBNCS ='T'"
    sql = sql & "   and XT.LIVKCS  ='0'"
    sql = sql & "   and XB.CRYNUMCS= X.CRYNUMCA"    ' 下サンプル条件
    sql = sql & "   and XB.TBKBNCS ='B'"
    sql = sql & "   and XB.LIVKCS  ='0'"
    
    sql = sql & " order by X.CRYNUMCA, X.INPOSCA "

    'データを抽出する
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    'レコード0件時は正常
    If rs.RecordCount = 0 Then
        rs.Close
        ReDim record0(0)
        ReDim record1(0)
        ReDim LWD(0)
    Else
        BlockIdBuf1 = vbNullString:  j1 = 0     ' 検査待ち
        BlockIdBuf2 = vbNullString:  j2 = 0     ' 判定待ち
        
        recCnt = rs.RecordCount
        For i = 1 To recCnt
            'ブロックID等の格納
            DoEvents
            
            If rs("DTTYPE") > 0 Then
            '<検査待ち>
                If rs("BLOCKID") <> BlockIdBuf1 Then
                    j1 = j1 + 1
                    ReDim Preserve record0(j1)
                    
                    With record0(j1)
                        .CRYNUM = rs("CRYNUM")
                        .IngotPos = rs("INGOTPOS")
                        .BLOCKID = rs("BLOCKID")   ' ブロックID
                        .UPDDATE = rs("UPDDATE")   ' 更新日付
                        .HOLDCLS = rs("HOLDCLS")   ' ホールド区分
                        .HSXTYPE = rs("HSXTYPE")
                        .HSXCDIR = rs("HSXCDIR")
                        .Judg = " "
                        BlockIdBuf1 = record0(j1).BLOCKID
                    End With
                    
                    k1 = 1
                End If
                
                '品番の格納
                ReDim Preserve record0(j1).hin(k1)
                With record0(j1).hin(k1)
                    .hinban = rs("HINBAN")
                    .mnorevno = rs("REVNUM")
                    .factory = rs("FACTORY")
                    .opecond = rs("OPECOND")
                End With
                k1 = k1 + 1
            
            Else
            '<判定待ち>
                If rs("BLOCKID") <> BlockIdBuf2 Then
                    j2 = j2 + 1
                    ReDim Preserve record1(j2)
                    ReDim Preserve LWD(j2)
                    
                    With record1(j2)
                        .CRYNUM = rs("CRYNUM")
                        .IngotPos = rs("INGOTPOS")
                        .BLOCKID = rs("BLOCKID")   ' ブロックID
                        .UPDDATE = rs("UPDDATE")   ' 更新日付
                        .HOLDCLS = rs("HOLDCLS")   ' ホールド区分
                        .HSXTYPE = rs("HSXTYPE")
                        .HSXCDIR = rs("HSXCDIR")
                        .Judg = " "
                        BlockIdBuf2 = record1(j2).BLOCKID
                    End With
                    
                    ' 実測(1)に対する実績がすべて設定されているので
                    ' 反映(2)、推定(3)に対しても実績がすべて設定されているかチェック
                    LWD(j2).flag = False
                    ' 上サンプル
                    If (rs("T_CRYINDRSCS") = "3" And rs("T_CRYRESR2CS") = "0") _
                    Or (rs("T_CRYINDRSCS") > "1" And rs("T_CRYRESR1CS") = "0") _
                    Or (rs("T_CRYINDOICS") > "1" And rs("T_CRYRESOICS") = "0") _
                    Or (rs("T_CRYINDB1CS") > "1" And rs("T_CRYRESB1CS") = "0") _
                    Or (rs("T_CRYINDB2CS") > "1" And rs("T_CRYRESB2CS") = "0") _
                    Or (rs("T_CRYINDB3CS") > "1" And rs("T_CRYRESB3CS") = "0") _
                    Or (rs("T_CRYINDL1CS") > "1" And rs("T_CRYRESL1CS") = "0") _
                    Or (rs("T_CRYINDL2CS") > "1" And rs("T_CRYRESL2CS") = "0") _
                    Or (rs("T_CRYINDL3CS") > "1" And rs("T_CRYRESL3CS") = "0") _
                    Or (rs("T_CRYINDL4CS") > "1" And rs("T_CRYRESL4CS") = "0") _
                    Or (rs("T_CRYINDCSCS") > "1" And rs("T_CRYRESCSCS") = "0") _
                    Or (rs("T_CRYINDGDCS") > "1" And rs("T_CRYRESGDCS") = "0") _
                    Or (rs("T_CRYINDT_CS") > "1" And rs("T_CRYREST_CS") = "0") _
                    Or (rs("T_CRYINDEPCS") > "1" And rs("T_CRYRESEPCS") = "0") Then
                        LWD(j2).flag = True
                    End If
                    ' 下サンプル
                    If (rs("B_CRYINDRSCS") = "3" And rs("B_CRYRESR2CS") = "0") _
                    Or (rs("B_CRYINDRSCS") > "1" And rs("B_CRYRESR1CS") = "0") _
                    Or (rs("B_CRYINDOICS") > "1" And rs("B_CRYRESOICS") = "0") _
                    Or (rs("B_CRYINDB1CS") > "1" And rs("B_CRYRESB1CS") = "0") _
                    Or (rs("B_CRYINDB2CS") > "1" And rs("B_CRYRESB2CS") = "0") _
                    Or (rs("B_CRYINDB3CS") > "1" And rs("B_CRYRESB3CS") = "0") _
                    Or (rs("B_CRYINDL1CS") > "1" And rs("B_CRYRESL1CS") = "0") _
                    Or (rs("B_CRYINDL2CS") > "1" And rs("B_CRYRESL2CS") = "0") _
                    Or (rs("B_CRYINDL3CS") > "1" And rs("B_CRYRESL3CS") = "0") _
                    Or (rs("B_CRYINDL4CS") > "1" And rs("B_CRYRESL4CS") = "0") _
                    Or (rs("B_CRYINDCSCS") > "1" And rs("B_CRYRESCSCS") = "0") _
                    Or (rs("B_CRYINDGDCS") > "1" And rs("B_CRYRESGDCS") = "0") _
                    Or (rs("B_CRYINDT_CS") > "1" And rs("B_CRYREST_CS") = "0") _
                    Or (rs("B_CRYINDEPCS") > "1" And rs("B_CRYRESEPCS") = "0") Then
                        LWD(j2).flag = True
                    End If
                    
                    ' ※type_DBDRV_scmzc_fcmkc001b_Dispにはサンプル情報を保持する構造体も定義されているが
                    ' 　フラグ設定の判定にしか使用していないので設定しないでおく
                    ' 　設定する場合はcmkc001b_DBDataCheck1を参照（処理時間が長いので関数は不使用の事）
                    
                    k2 = 1
                End If
                
                '品番の格納
                ReDim Preserve record1(j2).hin(k2)
                With record1(j2).hin(k2)
                    .hinban = rs("HINBAN")
                    .mnorevno = rs("REVNUM")
                    .factory = rs("FACTORY")
                    .opecond = rs("OPECOND")
                End With
                k2 = k2 + 1
            End If
            
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
    DBDRV_scmzc_fcmkc001b_Disp00 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function



''''''概要    :待ち一覧 初期表示用ＤＢドライバ（検査待ち）
''''''ﾊﾟﾗﾒｰﾀ  :変数名       ,IO  ,型                                 ,説明
''''''        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,初期表示用
''''''        :戻ﾘ値        ,O   ,FUNCTION_RETURN                    ,読み込み成否
''''''説明    :
''''''履歴    :2001/07/06 蔵本 作成
'''''Public Function DBDRV_scmzc_fcmkc001b_Disp1(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
'''''
'''''    Dim sql         As String       'SQL全体
'''''    Dim rs          As OraDynaset   'RecordSet
'''''    Dim recCnt      As Long         'ブロック管理のレコード数
'''''    Dim i           As Long
'''''    Dim j           As Long
'''''    Dim k           As Long
'''''    Dim BlockIdBuf  As String
'''''
'''''    '<検査待ち＞
'''''    'ブロック管理テーブルからブロックID、更新日付取得（検査実績が未検査のもの）
'''''
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp1"
'''''
'''''    DBDRV_scmzc_fcmkc001b_Disp1 = FUNCTION_RETURN_SUCCESS
'''''
'''''    'ブロックID、更新日付の取得
'''''    sql = "select distinct "
'''''
'''''    sql = sql & " X.XTALCA    as CRYNUM, "
'''''    sql = sql & " B.INGOTPOS, "
'''''    sql = sql & " X.CRYNUMCA  as BLOCKID, "
'''''    sql = sql & " B.UPDDATE, "
'''''    sql = sql & " B.HOLDCLS, "
'''''    sql = sql & " X.HINBCA    as HINBAN, "          ' 品番
'''''    sql = sql & " X.REVNUMCA  as REVNUM, "          ' 製品番号改訂番号
'''''    sql = sql & " X.FACTORYCA as FACTORY, "         ' 工場
'''''    sql = sql & " X.OPECA     as OPECOND, "         ' 操業条件
'''''    sql = sql & " S.HSXTYPE, "                      ' 品ＳＸタイプ
'''''    sql = sql & " S.HSXCDIR, "                      ' 品ＳＸ結晶面方位
'''''    sql = sql & " X.INPOSCA   as INGOTPOS "
'''''
'''''
'''''    sql = sql & " from  XSDCA X, TBCME040 B, TBCME018 S, XSDCS X2 "
'''''
'''''    sql = sql & " where X.CRYNUMCA = B.BLOCKID "
'''''    sql = sql & "   and X.CRYNUMCA = X2.CRYNUMCS "
'''''    sql = sql & "   and X.HINBCA   = S.HINBAN "
'''''    sql = sql & "   and X.REVNUMCA = S.MNOREVNO "
'''''    sql = sql & "   and X.FACTORYCA= S.FACTORY "
'''''    sql = sql & "   and X.OPECA    = S.OPECOND "
'''''
'''''    sql = sql & "   and X.GNWKNTCA='CC600' "
'''''    sql = sql & "   and X.LSTATBCA='T' "
'''''    sql = sql & "   and X.RSTATBCA='T' "
'''''    sql = sql & "   and X.LIVKCA  ='0' "
'''''    sql = sql & "   and B.DELCLS  ='0' "
'''''   'sql = sql & "   and B.HOLDCLS ='0' " ' ホールドブロックも取得
'''''
'''''    '指示が0でなく実績が0
'''''    sql = sql & " and ((X2.CRYINDRSCS<>'0' and X2.CRYRESRS1CS='0')"       ' 結晶検査実績（Rs)
'''''    sql = sql & "   or (X2.CRYINDOICS<>'0' and X2.CRYRESOICS ='0')"       ' 結晶検査実績（Oi)
'''''    sql = sql & "   or (X2.CRYINDB1CS<>'0' and X2.CRYRESB1CS ='0')"       ' 結晶検査実績（B1)
'''''    sql = sql & "   or (X2.CRYINDB2CS<>'0' and X2.CRYRESB2CS ='0')"       ' 結晶検査実績（B2）
'''''    sql = sql & "   or (X2.CRYINDB3CS<>'0' and X2.CRYRESB3CS ='0')"       ' 結晶検査実績（B3)
'''''    sql = sql & "   or (X2.CRYINDL1CS<>'0' and X2.CRYRESL1CS ='0')"       ' 結晶検査実績（L1)
'''''    sql = sql & "   or (X2.CRYINDL2CS<>'0' and X2.CRYRESL2CS ='0')"       ' 結晶検査実績（L2)
'''''    sql = sql & "   or (X2.CRYINDL3CS<>'0' and X2.CRYRESL3CS ='0')"       ' 結晶検査実績（L3)
'''''    sql = sql & "   or (X2.CRYINDL4CS<>'0' and X2.CRYRESL4CS ='0')"       ' 結晶検査実績（L4)
'''''    sql = sql & "   or (X2.CRYINDCSCS<>'0' and X2.CRYRESCSCS ='0')"       ' 結晶検査実績（Cs)
'''''    sql = sql & "   or (X2.CRYINDGDCS<>'0' and X2.CRYRESGDCS ='0')"       ' 結晶検査実績（GD)
'''''    sql = sql & "   or (X2.CRYINDTCS <>'0' and X2.CRYRESTCS  ='0')"       ' 結晶検査実績（T)
'''''    sql = sql & "   or (X2.CRYINDEPCS<>'0' and X2.CRYRESEPCS ='0'))"      ' 結晶検査実績（EPD)
'''''
'''''    sql = sql & " order by X.CRYNUMCA, X.INPOSCA "
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
'''''                    .IngotPos = rs("INGOTPOS")
'''''                    .BLOCKID = rs("BLOCKID")   ' ブロックID
'''''                    .UPDDATE = rs("UPDDATE")   ' 更新日付
'''''                    .HOLDCLS = rs("HOLDCLS")   ' ホールド区分
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
'''''            ReDim Preserve records(j).hin(k)
'''''            records(j).hin(k).hinban = rs("HINBAN")
'''''            records(j).hin(k).mnorevno = rs("REVNUM")
'''''            records(j).hin(k).factory = rs("FACTORY")
'''''            records(j).hin(k).opecond = rs("OPECOND")
'''''            k = k + 1
'''''            rs.MoveNext
'''''        Next i
'''''        rs.Close
'''''
'''''    End If
'''''
'''''
'''''
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    DBDRV_scmzc_fcmkc001b_Disp1 = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
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
'''''    Dim sql         As String       'SQL全体
'''''    Dim rs          As OraDynaset   'RecordSet
'''''    Dim recCnt      As Long         'ブロック管理のレコード数
'''''    Dim i           As Long
'''''    Dim j           As Long
'''''    Dim k           As Long
'''''    Dim BlockIdBuf  As String
'''''
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp2"
'''''
'''''    DBDRV_scmzc_fcmkc001b_Disp2 = FUNCTION_RETURN_SUCCESS
'''''
'''''    sql = "select distinct "
'''''
'''''    sql = sql & " X.XTALCA    as CRYNUM,  "
'''''    sql = sql & " B.INGOTPOS  as ss, "
''''''   sql = sql & " B.LENGTH, "                   ' 長さ追加 2001/11/8
'''''    sql = sql & " X.CRYNUMCA  as BLOCKID, "
'''''    sql = sql & " B.UPDDATE, "
'''''    sql = sql & " B.HOLDCLS, "
'''''    sql = sql & " X.HINBCA    as HINBAN,  "     ' 品番
'''''    sql = sql & " X.REVNUMCA  as REVNUM,  "     ' 製品番号改訂番号
'''''    sql = sql & " X.FACTORYCA as FACTORY, "     ' 工場
'''''    sql = sql & " X.OPECA     as OPECOND, "     ' 操業条件
'''''    sql = sql & " S.HSXTYPE, "                  ' 品ＳＸタイプ
'''''    sql = sql & " S.HSXCDIR, "                  ' 品ＳＸ結晶面方位
'''''    sql = sql & " X.INPOSCA   as INGOTPOS "
'''''
''''''''''                '判定NGがあるかどうか
''''''''''    sql = sql & " (select count(*) from XSDCS X21 "
''''''''''    sql = sql & "  where   X21.CRYNUMCS=X.CRYNUMCA"
''''''''''    sql = sql & "    and ((X21.CRYINDRSCS<>'0' and X21.CRYRESRS1CS='2')"            ' 結晶検査実績（Rs)
''''''''''    sql = sql & "      or (X21.CRYINDOICS<>'0' and X21.CRYRESOICS ='2')"            ' 結晶検査実績（Oi)
''''''''''    sql = sql & "      or (X21.CRYINDB1CS<>'0' and X21.CRYRESB1CS ='2')"            ' 結晶検査実績（B1)
''''''''''    sql = sql & "      or (X21.CRYINDB2CS<>'0' and X21.CRYRESB2CS ='2')"            ' 結晶検査実績（B2）
''''''''''    sql = sql & "      or (X21.CRYINDB3CS<>'0' and X21.CRYRESB3CS ='2')"            ' 結晶検査実績（B3)
''''''''''    sql = sql & "      or (X21.CRYINDL1CS<>'0' and X21.CRYRESL1CS ='2')"            ' 結晶検査実績（L1)
''''''''''    sql = sql & "      or (X21.CRYINDL2CS<>'0' and X21.CRYRESL2CS ='2')"            ' 結晶検査実績（L2)
''''''''''    sql = sql & "      or (X21.CRYINDL3CS<>'0' and X21.CRYRESL3CS ='2')"            ' 結晶検査実績（L3)
''''''''''    sql = sql & "      or (X21.CRYINDL4CS<>'0' and X21.CRYRESL4CS ='2')"            ' 結晶検査実績（L4)
''''''''''    sql = sql & "      or (X21.CRYINDCSCS<>'0' and X21.CRYRESCSCS ='2')"            ' 結晶検査実績（Cs)
''''''''''    sql = sql & "      or (X21.CRYINDGDCS<>'0' and X21.CRYRESGDCS ='2')"            ' 結晶検査実績（GD)
''''''''''    sql = sql & "      or (X21.CRYINDTCS <>'0' and X21.CRYRESTCS  ='2')"            ' 結晶検査実績（T)
''''''''''    sql = sql & "      or (X21.CRYINDEPCS<>'0' and X21.CRYRESEPCS ='2')) ) as J "   ' 結晶検査実績（EPD)
'''''
'''''    sql = sql & " from  XSDCA X, TBCME040 B, TBCME018 S"
'''''    sql = sql & " where X.CRYNUMCA = B.BLOCKID "
'''''    sql = sql & "   and X.HINBCA   = S.HINBAN "
'''''    sql = sql & "   and X.REVNUMCA = S.MNOREVNO "
'''''    sql = sql & "   and X.FACTORYCA= S.FACTORY "
'''''    sql = sql & "   and X.OPECA    = S.OPECOND "
'''''
'''''    '工程コード、状態、区分の条件指定
'''''
'''''    sql = sql & "   and X.GNWKNTCA='CC600' "
'''''    sql = sql & "   and X.LSTATBCA='T' "
'''''    sql = sql & "   and X.RSTATBCA='T' "
'''''    sql = sql & "   and X.LIVKCA  ='0' "
'''''    sql = sql & "   and B.DELCLS  ='0' "
'''''   'sql = sql & "   and B.HOLDCLS ='0' " ' ホールドブロックも取得
'''''
'''''''                'ブロック内に含まれる品番を検索
'''''''    sql = sql & " and (( B.INGOTPOS >= H.INGOTPOS "
'''''''    sql = sql & " and B.INGOTPOS < H.INGOTPOS + H.LENGTH ) "
'''''''    sql = sql & " or ( B.INGOTPOS + B.LENGTH > H.INGOTPOS "
'''''''    sql = sql & " and B.INGOTPOS + B.LENGTH < H.INGOTPOS + H.LENGTH  ) "
'''''''    sql = sql & " or ( H.INGOTPOS >= B.INGOTPOS "
'''''''    sql = sql & " and H.INGOTPOS < B.INGOTPOS + B.LENGTH ) "
'''''''    sql = sql & " or ( H.INGOTPOS + H.LENGTH > B.INGOTPOS "
'''''''    sql = sql & " and H.INGOTPOS + H.LENGTH < B.INGOTPOS + B.LENGTH )) "
'''''' ブロックをキーにデータを取得するので範囲外は無い
'''''
'''''                '指示が0でなく実績が0でないサンプルが上下２枚あるか
'''''    sql = sql & " and 2=(select count(*) From XSDCS X22"
'''''    sql = sql & "        where  X22.CRYNUMCS=X.CRYNUMCA"
'''''
'''''    sql = sql & "         and ((X22.CRYINDRSCS<>'0' and X22.CRYRESRS1CS<>'0')"          ' 結晶検査実績（Rs)
'''''    sql = sql & "          or  (X22.CRYINDOICS<>'0' and X22.CRYRESOICS <>'0')"          ' 結晶検査実績（Oi)
'''''    sql = sql & "          or  (X22.CRYINDB1CS<>'0' and X22.CRYRESB1CS <>'0')"          ' 結晶検査実績（B1)
'''''    sql = sql & "          or  (X22.CRYINDB2CS<>'0' and X22.CRYRESB2CS <>'0')"          ' 結晶検査実績（B2）
'''''    sql = sql & "          or  (X22.CRYINDB3CS<>'0' and X22.CRYRESB3CS <>'0')"          ' 結晶検査実績（B3)
'''''    sql = sql & "          or  (X22.CRYINDL1CS<>'0' and X22.CRYRESL1CS <>'0')"          ' 結晶検査実績（L1)
'''''    sql = sql & "          or  (X22.CRYINDL2CS<>'0' and X22.CRYRESL2CS <>'0')"          ' 結晶検査実績（L2)
'''''    sql = sql & "          or  (X22.CRYINDL3CS<>'0' and X22.CRYRESL3CS <>'0')"          ' 結晶検査実績（L3)
'''''    sql = sql & "          or  (X22.CRYINDL4CS<>'0' and X22.CRYRESL4CS <>'0')"          ' 結晶検査実績（L4)
'''''    sql = sql & "          or  (X22.CRYINDCSCS<>'0' and X22.CRYRESCSCS <>'0')"          ' 結晶検査実績（Cs)
'''''    sql = sql & "          or  (X22.CRYINDGDCS<>'0' and X22.CRYRESGDCS <>'0')"          ' 結晶検査実績（GD)
'''''    sql = sql & "          or  (X22.CRYINDTCS <>'0' and X22.CRYRESTCS  <>'0')"          ' 結晶検査実績（T)
'''''    sql = sql & "          or  (X22.CRYINDEPCS<>'0' and X22.CRYRESEPCS <>'0')) )"       ' 結晶検査実績（EPD)
''''''''''    sql = sql & "          and (X22.CRYINDRSCS='0' or X22.CRYRESRS1CS<>'0')"          ' 結晶検査実績（Rs)
''''''''''    sql = sql & "          and (X22.CRYINDOICS='0' or X22.CRYRESOICS <>'0')"          ' 結晶検査実績（Oi)
''''''''''    sql = sql & "          and (X22.CRYINDB1CS='0' or X22.CRYRESB1CS <>'0')"          ' 結晶検査実績（B1)
''''''''''    sql = sql & "          and (X22.CRYINDB2CS='0' or X22.CRYRESB2CS <>'0')"          ' 結晶検査実績（B2）
''''''''''    sql = sql & "          and (X22.CRYINDB3CS='0' or X22.CRYRESB3CS <>'0')"          ' 結晶検査実績（B3)
''''''''''    sql = sql & "          and (X22.CRYINDL1CS='0' or X22.CRYRESL1CS <>'0')"          ' 結晶検査実績（L1)
''''''''''    sql = sql & "          and (X22.CRYINDL2CS='0' or X22.CRYRESL2CS <>'0')"          ' 結晶検査実績（L2)
''''''''''    sql = sql & "          and (X22.CRYINDL3CS='0' or X22.CRYRESL3CS <>'0')"          ' 結晶検査実績（L3)
''''''''''    sql = sql & "          and (X22.CRYINDL4CS='0' or X22.CRYRESL4CS <>'0')"          ' 結晶検査実績（L4)
''''''''''    sql = sql & "          and (X22.CRYINDCSCS='0' or X22.CRYRESCSCS <>'0')"          ' 結晶検査実績（Cs)
''''''''''    sql = sql & "          and (X22.CRYINDGDCS='0' or X22.CRYRESGDCS <>'0')"          ' 結晶検査実績（GD)
''''''''''    sql = sql & "          and (X22.CRYINDTCS ='0' or X22.CRYRESTCS  <>'0')"          ' 結晶検査実績（T)
''''''''''    sql = sql & "          and (X22.CRYINDEPCS='0' or X22.CRYRESEPCS <>'0') )"        ' 結晶検査実績（EPD)
'''''
'''''''    sql = sql & " order by B.BLOCKID, H.INGOTPOS "
'''''    sql = sql & " order by X.CRYNUMCA, X.INPOSCA "
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
'''''                    .IngotPos = rs("ss")
''''''                   .LENGTH = rs("LENGTH")      ' 長さ
'''''                    .BLOCKID = rs("BLOCKID")    ' ブロックID
'''''                    .UPDDATE = rs("UPDDATE")    ' 更新日付
'''''                    .HOLDCLS = rs("HOLDCLS")    ' ホールド区分
'''''                    BlockIdBuf = records(j).BLOCKID
'''''                    .HSXTYPE = rs("HSXTYPE")
'''''                    .HSXCDIR = rs("HSXCDIR")
''''''                    If rs("J") > 0 Then
''''''
''''''                        .Judg = "2"
''''''                    Else
'''''                        .Judg = "1"
''''''                    End If
'''''
'''''                End With
'''''                k = 1
'''''            End If
'''''
'''''            '品番の格納
'''''            ReDim Preserve records(j).hin(k)
'''''            records(j).hin(k).hinban = rs("HINBAN")
'''''            records(j).hin(k).mnorevno = rs("REVNUM")
'''''            records(j).hin(k).factory = rs("FACTORY")
'''''            records(j).hin(k).opecond = rs("OPECOND")
'''''            k = k + 1
'''''            rs.MoveNext
'''''        Next i
'''''        rs.Close
'''''
'''''    End If
'''''
'''''
''''''''''    '購入単結晶実績取得
''''''''''    If getKouBlock(records(), "CC600") = FUNCTION_RETURN_FAILURE Then
''''''''''       DBDRV_scmzc_fcmkc001b_Disp2 = FUNCTION_RETURN_FAILURE
''''''''''       GoTo proc_exit
''''''''''    End If
'''''
'''''
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    DBDRV_scmzc_fcmkc001b_Disp2 = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function



'''''概要    :待ち一覧 初期表示用ＤＢドライバ（払出待ち）
'''''ﾊﾟﾗﾒｰﾀ  :変数名       ,IO  ,型                                 ,説明
'''''        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,初期表示用
'''''        :戻ﾘ値        ,O   ,FUNCTION_RETURN                    ,読み込み成否
'''''説明    :
'''''履歴    :2001/07/06 蔵本 作成
''''Public Function DBDRV_scmzc_fcmkc001b_Disp3(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
''''
''''    '＜払出待ち＞
''''    'CC700のもの
''''
''''    'エラーハンドラの設定
''''    On Error GoTo proc_err
''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp3"
''''
''''
''''    DBDRV_scmzc_fcmkc001b_Disp3 = FUNCTION_RETURN_SUCCESS
''''
''''    'ブロックID､更新日付、品番等取得
''''    If getBlockID(records(), "CC700") = FUNCTION_RETURN_FAILURE Then
''''        DBDRV_scmzc_fcmkc001b_Disp3 = FUNCTION_RETURN_FAILURE
''''        GoTo proc_exit
''''    End If
''''
''''
'''''    '購入単結晶実績取得
'''''    If getKouBlock(records(), "CC700") = FUNCTION_RETURN_FAILURE Then
'''''       DBDRV_scmzc_fcmkc001b_Disp3 = FUNCTION_RETURN_FAILURE
'''''       GoTo proc_exit
'''''    End If
''''
''''proc_exit:
''''    '終了
''''    gErr.Pop
''''    Exit Function
''''
''''proc_err:
''''    'エラーハンドラ
''''    gErr.HandleError
''''    DBDRV_scmzc_fcmkc001b_Disp3 = FUNCTION_RETURN_FAILURE
''''    Resume proc_exit
''''End Function
''''
''''
''''
'''''概要    :待ち一覧 初期表示用ＤＢドライバ（抜試指示待ち）
'''''ﾊﾟﾗﾒｰﾀ  :変数名       ,IO  ,型                                 ,説明
'''''        :records      ,IO  ,type_DBDRV_scmzc_fcmkc001b_Disp    ,初期表示用
'''''        :戻ﾘ値        ,O   ,FUNCTION_RETURN                    ,読み込み成否
'''''説明    :
'''''履歴    :2001/07/06 蔵本 作成
''''Public Function DBDRV_scmzc_fcmkc001b_Disp4(records() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
''''
''''    '＜抜試指示待ち＞
''''    'CC710のもの
''''
''''    'ブロックID､更新日付取得
''''
''''    'エラーハンドラの設定
''''    On Error GoTo proc_err
''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001b_Disp4"
''''
''''    DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_SUCCESS
''''
''''
''''    'ブロックID､更新日付、品番等取得
''''    If getBlockID(records(), "CC710") = FUNCTION_RETURN_FAILURE Then
''''        DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_FAILURE
''''        GoTo proc_exit
''''    End If
''''
'''''2000/08/24 S.Sano Start
'''''    '購入単結晶実績取得
'''''    If getKouBlock(records(), "CC710") = FUNCTION_RETURN_FAILURE Then
'''''       DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_FAILURE
'''''       GoTo proc_exit
'''''    End If
'''''2000/08/24 S.Sano End
''''
''''
''''proc_exit:
''''    '終了
''''    gErr.Pop
''''    Exit Function
''''
''''proc_err:
''''    'エラーハンドラ
''''    gErr.HandleError
''''    DBDRV_scmzc_fcmkc001b_Disp4 = FUNCTION_RETURN_FAILURE
''''    Resume proc_exit
''''End Function


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
'''''    Dim l   As Long
'''''    Dim m   As Long
'''''    Dim sql As String
'''''    Dim rs  As OraDynaset    'RecordSet
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
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
'''''
'''''    sql = sql & "select "
'''''    sql = sql & "B.BLOCKID,  B.INGOTPOS as TOPPOS,    B.INGOTPOS+LENGTH as BOTPOS, "
'''''
'''''    sql = sql & "S.XTALCS,   S.INPOSCS,   SMPKBNCS,   HINBCS,     REVNUMCS,   FACTORYCS,  OPECS, "
'''''    sql = sql & "CRYINDRSCS, CRYRESRS1CS, CRYINDOICS, CRYRESOICS, CRYINDB1CS, CRYRESB1CS, CRYINDB2CS,"
'''''    sql = sql & "CRYRESB2CS, CRYINDB3CS,  CRYRESB3CS, CRYINDL1CS, CRYRESL1CS, CRYINDL2CS, CRYRESL2CS,"
'''''    sql = sql & "CRYINDL3CS, CRYRESL3CS,  CRYINDL4CS, CRYRESL4CS, CRYINDCSCS, CRYRESCSCS, CRYINDGDCS,"
'''''    sql = sql & "CRYRESGDCS, CRYINDTCS,   CRYRESTCS,  CRYINDEPCS, CRYRESEPCS "
'''''
'''''    sql = sql & "from  XSDCS S, TBCME040 B "
'''''
'''''    sql = sql & "where S.XTALCS  = B.CRYNUM"
'''''
'''''    sql = sql & "  and B.INGOTPOS>=0"
'''''    sql = sql & "  and B.DELCLS  = '0'"
'''''    sql = sql & "  and B.NOWPROC in ('CC600','CC700', 'CC710')"
'''''    sql = sql & "  and B.RSTATCLS='T'"
'''''    sql = sql & "  and B.HOLDCLS ='0'"
'''''    sql = sql & "  and ((S.INPOSCS=B.INGOTPOS) or (S.INPOSCS=B.INGOTPOS+B.LENGTH)) "
'''''
'''''    sql = sql & "order by B.BLOCKID, S.INPOSCS, S.SMPKBNCS"
'''''
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
'''''            .CRYNUM = rs("XTALCS")
'''''            .IngotPos = rs("INPOSCS")
'''''            .SMPKBN = rs("SMPKBNCS")
'''''            .hinban = rs("HINBCS")
'''''            .REVNUM = rs("REVNUMCS")
'''''            .factory = rs("FACTORYCS")
'''''            .opecond = rs("OPECS")
'''''            .CRYINDRS = rs("CRYINDRSCS")
'''''            .CRYRESRS = rs("CRYRESRS1CS")
'''''            .CRYINDOI = rs("CRYINDOICS")
'''''            .CRYRESOI = rs("CRYRESOICS")
'''''            .CRYINDB1 = rs("CRYINDB1CS")
'''''            .CRYRESB1 = rs("CRYRESB1CS")
'''''            .CRYINDB2 = rs("CRYINDB2CS")
'''''            .CRYRESB2 = rs("CRYRESB2CS")
'''''            .CRYINDB3 = rs("CRYINDB3CS")
'''''            .CRYRESB3 = rs("CRYRESB3CS")
'''''            .CRYINDL1 = rs("CRYINDL1CS")
'''''            .CRYRESL1 = rs("CRYRESL1CS")
'''''            .CRYINDL2 = rs("CRYINDL2CS")
'''''            .CRYRESL2 = rs("CRYRESL2CS")
'''''            .CRYINDL3 = rs("CRYINDL3CS")
'''''            .CRYRESL3 = rs("CRYRESL3CS")
'''''            .CRYINDL4 = rs("CRYINDL4CS")
'''''            .CRYRESL4 = rs("CRYRESL4CS")
'''''            .CRYINDCS = rs("CRYINDCSCS")
'''''            .CRYRESCS = rs("CRYRESCSCS")
'''''            .CRYINDGD = rs("CRYINDGDCS")
'''''            .CRYRESGD = rs("CRYRESGDCS")
'''''            .CRYINDT = rs("CRYINDTCS")
'''''            .CRYREST = rs("CRYRESTCS")
'''''            .CRYINDEP = rs("CRYINDEPCS")
'''''            .CRYRESEP = rs("CRYRESEPCS")
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
'''''                    .IngotPos = SMP(idx).IngotPos
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
'''''            sql = "select "
'''''            sql = sql & " XTALCS, "
'''''            sql = sql & " INPOSCS, "
'''''            sql = sql & " SMPKBNCS, "
'''''            sql = sql & " HINBCS, "
'''''            sql = sql & " REVNUMCS, "
'''''            sql = sql & " FACTORYCS, "
'''''            sql = sql & " OPECS, "
'''''            sql = sql & " CRYINDRSCS, "
'''''            sql = sql & " CRYRESRSCS, "
'''''            sql = sql & " CRYINDOICS, "
'''''            sql = sql & " CRYRESOICS, "
'''''            sql = sql & " CRYINDB1CS, "
'''''            sql = sql & " CRYRESB1CS, "
'''''            sql = sql & " CRYINDB2CS, "
'''''            sql = sql & " CRYRESB2CS, "
'''''            sql = sql & " CRYINDB3CS, "
'''''            sql = sql & " CRYRESB3CS, "
'''''            sql = sql & " CRYINDL1CS, "
'''''            sql = sql & " CRYRESL1CS, "
'''''            sql = sql & " CRYINDL2CS, "
'''''            sql = sql & " CRYRESL2CS, "
'''''            sql = sql & " CRYINDL3CS, "
'''''            sql = sql & " CRYRESL3CS, "
'''''            sql = sql & " CRYINDL4CS, "
'''''            sql = sql & " CRYRESL4CS, "
'''''            sql = sql & " CRYINDCSCS, "
'''''            sql = sql & " CRYRESCSCS, "
'''''            sql = sql & " CRYINDGDCS, "
'''''            sql = sql & " CRYRESGDCS, "
'''''            sql = sql & " CRYINDTCS, "
'''''            sql = sql & " CRYRESTCS, "
'''''            sql = sql & " CRYINDEPCS, "
'''''            sql = sql & " CRYRESEPCS "
'''''
'''''''            sql = sql & " from VECME010 V "
'''''''            sql = sql & " where E040CRYNUM = '" & .Crynum & "' "
'''''''            sql = sql & " and   E040INGOTPOS = '" & .IngotPos & "' "
'''''''            sql = sql & " order by E043INPOSCS"
'''''
'''''            sql = sql & " from XSDCS "
'''''            sql = sql & " where CRYNUMCS = '" & .BLOCKID & "' "
'''''            sql = sql & " order by INPOSCS"
'''''
'''''            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''            For m = 1 To 2
'''''                DoEvents
'''''                .SMP(m).CRYNUM = rs("E043XTALCS")
'''''                .SMP(m).IngotPos = rs("E043INPOSCS")
'''''                .SMP(m).SMPKBN = rs("E043SMPKBNCS")
'''''                .SMP(m).hinban = rs("E043HINBCS")
'''''                .SMP(m).REVNUM = rs("E043REVNUMCS")
'''''                .SMP(m).factory = rs("E043FACTORYCS")
'''''                .SMP(m).opecond = rs("E043OPECS")
'''''                .SMP(m).CRYINDRS = rs("E043CRYINDRSCS")
'''''                .SMP(m).CRYRESRS = rs("E043CRYRESRS1CS")
'''''                .SMP(m).CRYINDOI = rs("E043CRYINDOICS")
'''''                .SMP(m).CRYRESOI = rs("E043CRYRESOICS")
'''''                .SMP(m).CRYINDB1 = rs("E043CRYINDB1CS")
'''''                .SMP(m).CRYRESB1 = rs("E043CRYRESB1CS")
'''''                .SMP(m).CRYINDB2 = rs("E043CRYINDB2CS")
'''''                .SMP(m).CRYRESB2 = rs("E043CRYRESB2CS")
'''''                .SMP(m).CRYINDB3 = rs("E043CRYINDB3CS")
'''''                .SMP(m).CRYRESB3 = rs("E043CRYRESB3CS")
'''''                .SMP(m).CRYINDL1 = rs("E043CRYINDL1CS")
'''''                .SMP(m).CRYRESL1 = rs("E043CRYRESL1CS")
'''''                .SMP(m).CRYINDL2 = rs("E043CRYINDL2CS")
'''''                .SMP(m).CRYRESL2 = rs("E043CRYRESL2CS")
'''''                .SMP(m).CRYINDL3 = rs("E043CRYINDL3CS")
'''''                .SMP(m).CRYRESL3 = rs("E043CRYRESL3CS")
'''''                .SMP(m).CRYINDL4 = rs("E043CRYINDL4CS")
'''''                .SMP(m).CRYRESL4 = rs("E043CRYRESL4CS")
'''''                .SMP(m).CRYINDCS = rs("E043CRYINDCSCS")
'''''                .SMP(m).CRYRESCS = rs("E043CRYRESCSCS")
'''''                .SMP(m).CRYINDGD = rs("E043CRYINDGDCS")
'''''                .SMP(m).CRYRESGD = rs("E043CRYRESGDCS")
'''''                .SMP(m).CRYINDT = rs("E043CRYINDTCS")
'''''                .SMP(m).CRYREST = rs("E043CRYRESTCS")
'''''                .SMP(m).CRYINDEP = rs("E043CRYINDEPCS")
'''''                .SMP(m).CRYRESEP = rs("E043CRYRESEPCS")
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
'''''                    sql = sql & " S.HSXCNHWS,"
'''''                    sql = sql & " S.HSXLTHWS,"
'''''                    sql = sql & " 'H' as EPD "
'''''                    sql = sql & " from  TBCME019 S "
'''''                    sql = sql & " where S.HINBAN   = '" & .SMP(m).hinban & "'"
'''''                    sql = sql & "   and S.MNOREVNO =  " & .SMP(m).REVNUM
'''''                    sql = sql & "   and S.FACTORY  = '" & .SMP(m).factory & "'"
'''''                    sql = sql & "   and S.OPECOND  = '" & .SMP(m).opecond & "'"
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
''''''                    sql = "select CRYRESCS as RES from TBCME043 "
''''''                    sql = sql & "where CRYNUM = '" & .SMP(m).CRYNUM & "' "
''''''                    sql = sql & "  and INGOTPOS >= " & .SMP(m).INGOTPOS
''''''                    sql = sql & "  and CRYINDCS<>'0'"
''''''                    sql = sql & " order by INGOTPOS"
'''''                    sql = "select CRYRESCSCS as RES from XSDCS "
'''''                    sql = sql & "where XTALCS = '" & .SMP(m).CRYNUM & "' "
'''''                    sql = sql & "  and INPOSCS >= " & .SMP(m).IngotPos
'''''                    sql = sql & "  and CRYINDCSCS<>'0'"
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
'''''                ' LTのチェック
''''''                If (.SMP(m).HSXLTHWS = "H" Or .SMP(m).HSXLTHWS = "S") And .SMP(m).CRYINDT = "0" And LWD(l).flag = False Then ' 参考評価はなくてもＯＫ
'''''                If .SMP(m).HSXLTHWS = "H" And .SMP(m).CRYINDT = "0" And LWD(l).flag = False Then
'''''
''''''                    sql = "select CRYREST as RES from TBCME043 "
''''''                    sql = sql & "where CRYNUM = '" & .SMP(m).CRYNUM & "' "
''''''                    sql = sql & "  and INGOTPOS >= " & .SMP(m).INGOTPOS
''''''                    sql = sql & "  and CRYINDT<>'0'"
''''''                    sql = sql & " order by INGOTPOS"
'''''
'''''                    sql = "select CRYRESTCS as RES from XSDCS "
'''''                    sql = sql & "where XTALCS = '" & .SMP(m).CRYNUM & "' "
'''''                    sql = sql & "  and INPOSCS >= " & .SMP(m).IngotPos
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
''''''                    sql = "select CRYRESEP as RES from TBCME043 "
''''''                    sql = sql & "where CRYNUM = '" & .SMP(m).CRYNUM & "' "
''''''                    sql = sql & "  and INGOTPOS >= " & .SMP(m).INGOTPOS
''''''                    sql = sql & "  and CRYINDEP<>'0'"
''''''                    sql = sql & " order by INGOTPOS"
'''''
'''''                    sql = "select CRYRESEPCS as RES from XSDCS "
'''''                    sql = sql & "where XTALCS = '" & .SMP(m).CRYNUM & "' "
'''''                    sql = sql & "  and INPOSCS >= " & .SMP(m).IngotPos
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
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''
'''''
'''''    gErr.HandleError
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    cmkc001b_DBDataCheck1 = FUNCTION_RETURN_FAILURE
'''''    Resume proc_exit
'''''End Function


'''''Public Function cmkc001b_DBDataCheck3(LWD() As cmkc001b_LockWait, _
'''''                                 Wd3() As type_DBDRV_scmzc_fcmkc001b_Disp) As FUNCTION_RETURN
'''''    Dim c0 As Integer
'''''    Dim c1 As Integer
'''''    Dim c2 As Integer
'''''    Dim MaxRec As Integer
'''''    Dim RecCount As Integer
'''''    Dim EQFlag As Boolean
'''''    Dim sql As String       'SQL全体
'''''    Dim rs As OraDynaset    'RecordSet
'''''    Dim GrpCount1 As Integer
'''''    Dim GrpCount2 As Integer
'''''    Dim ColorFlag As Boolean
'''''    Dim TotalBlk As Integer
'''''    Dim CheckPoint As Integer
'''''    Dim CheckEnd As Integer
'''''    Dim tempGrpFlag As String * 1
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
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
'''''            GoTo proc_exit
'''''        End If
'''''        ReDim GrpInfo(c0).blkInfo(RecCount) As cmkc001b_Wait3_BLK
'''''        For c1 = 1 To RecCount
'''''            GrpInfo(c0).blkInfo(c1).BLOCKID = rs("BLOCKID")
'''''            GrpInfo(c0).blkInfo(c1).IngotPos = rs("INGOTPOS")
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
'''''Dim BLKID() As String
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
'''''    ReDim BLKID(1 To rsCount)
'''''    ReDim topHin(1 To rsCount)
'''''    ReDim botHin(1 To rsCount)
'''''    For c0 = 1 To rsCount
'''''        BLKID(c0) = rs!BLOCKID
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
'''''                If BLKID(idx) = GrpInfo(c0).blkInfo(c1).BLOCKID Then
'''''                    found = True
'''''                    Exit For
'''''                ElseIf BLKID(idx) > GrpInfo(c0).blkInfo(c1).BLOCKID Then
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
'''''            sql = sql & "and INGOTPOS = " & GrpInfo(c0).blkInfo(c1).IngotPos & " " '2001/11/14 S.Sano
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
'''''            sql = sql & "and INGOTPOS < " & GrpInfo(c0).blkInfo(c1).IngotPos + GrpInfo(c0).blkInfo(c1).LENGTH & " "
'''''            sql = sql & "and (INGOTPOS + LENGTH) >= " & GrpInfo(c0).blkInfo(c1).IngotPos + GrpInfo(c0).blkInfo(c1).LENGTH & " "
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
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    cmkc001b_DBDataCheck3 = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function




'総合判定

'12桁品番はommonDefine.BAS で定義されている type tFullHinban を使用しました
'Public Type tFullHinban
'    hinban As String * 8            ' 品番
'    MNOREVNO As Integer             ' 製品番号改訂番号
'    FACTORY As String * 1           ' 工場
'    OPECOND As String * 1           ' 操業条件
'End Type


'
'-*-*- 人見　20010807 ５推定、６引継ぎ、EPD,Cs,LT　修正
' ＜判定フロー＞
' 仕様保証方法＿処 --+--なし(H以外) --実績（該当位置）--あってもなくても判定OK
'　　　　　　　　　　|
'                   +--あり(H) --実績（該当位置) --+--あり -- 判定チェック --+-- OK
'                                                 |                        |
'                                                 |                        +-- MG
'                                                 |
'                                                 +--なし --+-- 検査指示５・６以外の場合 --+--EPD、Cs、LTの場合下を探す --+-- なし -- NG
'                                                           |                            |                          　 |
'                                                           |                            +--EPD,Cs、LT以外 -- NG       +-- あり -- 判定チェック --+-- OK
'                                                           |                                                                                    |
'                                                           |                                                                                    +-- NG
'                                                           |
'                                                           +-- 検査指示５の場合 (Rs, Cs) なら推定なので全体から実績を探す --+-- 判定チェック --+-- OK
'                                                           |  　　　　　　　　　                                          |                 |
'                                                           |                                                             |                 +-- NG
'                                                           +-- 検査指示６の場合 TOPなら上へ、TAILなら下へ実績を探す       --+
'
'　　　　　　　　　　　　　　　　　　　　　　　　　　　　   　（検査指示は、指示を立てる側が正常に立てていると考えている）



''''''概要      :内部関数　SQLに推定での検索条件を付加する
'''''Private Sub AddSQL_SUITEI(sql As String, Cry As String, Table As String, TorB As Integer, Optional subSQL As String = "")
'''''    sql = sql & " from " & Table & " T1 "
'''''    sql = sql & " where T1.CRYNUM='" & Cry & "' "
'''''    sql = sql & " and   T1.TRANCNT=ANY( select max(TRANCNT) from " & Table & " T2 "
'''''    sql = sql & "                       where T2.CRYNUM=T1.CRYNUM  and T2.POSITION=T1.POSITION and T2.SMPKBN=T1.SMPKBN " & subSQL & ") "
'''''    sql = sql & " and T1.SMPLUMU = '0' "
'''''    sql = sql & subSQL
'''''    If TorB = 1 Then
'''''        sql = sql & " order by T1.POSITION asc "     ' １回目は最初から
'''''    Else
'''''        sql = sql & " order by T1.POSITION desc "    ' ２回目は後ろから
'''''    End If
'''''End Sub
'''''
'''''
''''''概要      :内部関数　SQLに引継ぎでの検索条件を付加する
'''''Private Sub AddSQL_HIKITUGI(sql As String, Cry As String, pos As Integer, Table As String, TorB As Integer, Optional subSQL = "")
'''''    sql = sql & " from " & Table & " T1 "
'''''    sql = sql & " where T1.CRYNUM='" & Cry & "' "
'''''    sql = sql & " and   T1.TRANCNT=ANY( select max(TRANCNT) from " & Table & " T2 "
'''''    sql = sql & "                       where T2.CRYNUM=T1.CRYNUM  and T2.POSITION=T1.POSITION and T2.SMPKBN=T1.SMPKBN " & subSQL & ") "
'''''    sql = sql & " and T1.SMPLUMU = '0' "
'''''    sql = sql & subSQL
'''''    If TorB = 1 Then                        ' TOP側は上に探す
'''''        sql = sql & " and T1.POSITION < " & CStr(pos)
'''''        sql = sql & " order by T1.POSITION asc, SMPKBN asc "
'''''    Else                                    ' BOT側は下に探す
'''''        sql = sql & " and T1.POSITION > " & CStr(pos)
'''''        sql = sql & " order by T1.POSITION desc, SMPKBN desc "
'''''    End If
'''''End Sub
'''''
''''''概要      :内部関数　SQLに引継ぎでの検索条件を付加する
'''''Private Sub AddSQL_HIKITUGI2(sql As String, Cry As String, pos As Integer, Table As String, TorB As Integer, Optional subSQL = "")
'''''    sql = sql & " from " & Table & " T1 "
'''''    sql = sql & " where T1.CRYNUM='" & Cry & "' "
'''''    sql = sql & " and   T1.TRANCNT=ANY( select max(TRANCNT) from " & Table & " T2 "
'''''    sql = sql & "                       where T2.CRYNUM=T1.CRYNUM  and T2.POSITION=T1.POSITION and T2.SMPKBN=T1.SMPKBN " & subSQL & ") "
'''''    sql = sql & " and T1.SMPLUMU = '0' "
'''''    sql = sql & subSQL
'''''    If TorB = 1 Then                        ' TOP側は上に探す
'''''        sql = sql & " and T1.POSITION < " & CStr(pos)
'''''        sql = sql & " order by T1.POSITION desc, SMPKBN desc "
'''''    Else                                    ' BOT側は下に探す
'''''        sql = sql & " and T1.POSITION > " & CStr(pos)
'''''        sql = sql & " order by T1.POSITION asc, SMPKBN asc "
'''''    End If
'''''End Sub
'''''
''''''概要      :内部関数　SQLに下に実績を検索する検索条件を付加する
'''''Private Sub AddSQL_Down(sql As String, Cry As String, pos As Integer, Table As String, Optional subSQL = "")
'''''    sql = sql & " from " & Table & " T1 "
'''''    sql = sql & " where T1.CRYNUM='" & Cry & "' "
'''''    sql = sql & " and   T1.TRANCNT=ANY( select max(TRANCNT) from " & Table & " T2 "
'''''    sql = sql & "                       where T2.CRYNUM=T1.CRYNUM  and T2.POSITION=T1.POSITION and T2.SMPKBN=T1.SMPKBN  " & subSQL & ") "
'''''    sql = sql & " and T1.SMPLUMU = '0' "
'''''    sql = sql & " and T1.POSITION > " & CStr(pos)
'''''    sql = sql & subSQL
'''''    sql = sql & " order by POSITION asc, SMPKBN asc "
'''''End Sub


'''''Private Sub AddSQL_Default(sql As String, Cry As String, pos As Integer, Spk As String, Table As String, Optional subSQL = "")
'''''    sql = sql & " from " & Table
'''''    sql = sql & " where CRYNUM='" & Cry & "' " & _
'''''                " and POSITION=" & pos & _
'''''                " and SMPKBN='" & Spk & "' " & _
'''''                subSQL & _
'''''                " and TRANCNT=ANY( select max(TRANCNT) from " & Table & _
'''''                                   " where CRYNUM='" & Cry & "' " & _
'''''                                   " and POSITION=" & pos & _
'''''                                   " and SMPKBN='" & Spk & "' " & _
'''''                                   subSQL & " ) "
'''''End Sub
'''''
'''''Private Sub AddSQL_Default2(sql As String, Cry As String, pos As Integer, Spk As String, Table As String, Optional subSQL = "")
'''''    sql = sql & " from " & Table
'''''    sql = sql & " where CRYNUM='" & Cry & "' " & _
'''''                " and POSITION=" & pos & _
'''''                subSQL & _
'''''                " and TRANCNT=ANY( select max(TRANCNT) from " & Table & _
'''''                                   " B where B.CRYNUM='" & Cry & "' " & _
'''''                                   " and B.POSITION=" & pos & _
'''''                                   " and B.SMPKBN=SMPKBN " & _
'''''                                   subSQL & " ) "
'''''    If (Spk = "T") Then
'''''        sql = sql & "order by SMPKBN desc "
'''''    Else
'''''        sql = sql & "order by SMPKBN "
'''''    End If
'''''End Sub


''''''概要      :内部関数 結晶抵抗実績取得用オブジェクトコピー関数(レコードセットからのコピー)
'''''Private Sub CryR_ObjCpy(CryR As type_DBDRV_scmzc_fcmkc001c_CryR, rs As OraDynaset)
'''''    With CryR
'''''        .CRYNUM = rs("CRYNUM")         ' 結晶番号
'''''        .POSITION = rs("POSITION")     ' 位置
'''''        .SMPKBN = rs("SMPKBN")         ' サンプル区分
'''''        .SMPLNO = rs("SMPLNO")         ' サンプルＮｏ
'''''        .SMPLUMU = rs("SMPLUMU")       ' サンプル有無
'''''        .TRANCOND = rs("TRANCOND")     ' 処理条件
'''''        .MEAS1 = rs("MEAS1")           ' 測定値１
'''''        .MEAS2 = rs("MEAS2")           ' 測定値２
'''''        .MEAS3 = rs("MEAS3")           ' 測定値３
'''''        .MEAS4 = rs("MEAS4")           ' 測定値4
'''''        .MEAS5 = rs("MEAS5")           ' 測定値５
'''''        .RRG = rs("RRG")               ' RRG
'''''        .REGDATE = rs("REGDATE")       ' 登録日付
'''''    End With
'''''End Sub
'''''
''''''概要      :内部関数 結晶抵抗実績取得用ベースSQL
'''''Private Sub CryR_SetBaseSQL(sql As String)
'''''    sql = "select "
'''''    sql = sql & "CRYNUM, "        ' 結晶番号
'''''    sql = sql & "POSITION, "      ' 位置
'''''    sql = sql & "SMPKBN, "        ' サンプル区分
'''''    sql = sql & "SMPLNO, "        ' サンプルＮｏ
'''''    sql = sql & "SMPLUMU, "       ' サンプル有無
'''''    sql = sql & "TRANCOND, "      ' 処理条件
'''''    sql = sql & "MEAS1, "         ' 測定値１
'''''    sql = sql & "MEAS2, "         ' 測定値２
'''''    sql = sql & "MEAS3, "         ' 測定値３
'''''    sql = sql & "MEAS4, "         ' 測定値４
'''''    sql = sql & "MEAS5, "         ' 測定値５
'''''    sql = sql & "RRG, "            ' RRG
'''''    sql = sql & "REGDATE "         '　登録日付
'''''End Sub


''''''概要      :内部関数 結晶抵抗実績取得用
'''''Private Function CryR_Zisseki(Siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                              Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
'''''                              CryR As type_DBDRV_scmzc_fcmkc001c_CryR, _
'''''                              SuCryR As type_DBDRV_scmzc_fcmkc001c_CryR, _
'''''                              TorB As Integer, _
'''''                              Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
'''''
'''''    Dim sql As String
'''''    Dim rs As OraDynaset
'''''    Dim i As Long
'''''    Dim recCnt As Integer
'''''    Dim NothingFlag As Boolean
'''''
'''''    NothingFlag = False
'''''
'''''    ' 結晶抵抗実績テーブルから値を取得
'''''    Dim Tname As String
'''''    Tname = "TBCMJ002"
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function CryR_Zisseki"
'''''
'''''    CryR_Zisseki = FUNCTION_RETURN_SUCCESS
'''''
'''''    Set rs = Nothing
'''''
'''''    ' 検査指示にかかわらず実績を取得する　（品番を何回も振り替えた・手動で検査指示を立てた　などのため）
'''''    DoEvents
'''''    Call CryR_SetBaseSQL(sql)
'''''    DoEvents
'''''    Call AddSQL_Default2(sql, Samp.CRYNUM, Samp.INGOTPOS, Samp.SMPKBN, Tname)
'''''    DoEvents
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    DoEvents
'''''
'''''    If rs.RecordCount <> 0 Then          ' 実績があったら、それを採用する
'''''        DoEvents
'''''        Call CryR_ObjCpy(CryR, rs)
'''''        SuCryR = CryR   '2001/10/24 S.Sano　推定でも実績が存在した場合の処理
'''''        rs.Close
'''''        Set rs = Nothing
'''''    Else
'''''        rs.Close
'''''        Set rs = Nothing
'''''
''''''----- 8/12 野村 修正（比抵抗は仕様に関わらず結果を表示したい）
''''''元        If Siyou.HSXRHWYS = SIJI Then   ' 仕様の指示がたっている
''''''        If Siyou.HSXRHWYS = SIJI Then   ' 仕様の指示がたっている
''''''
'''''            If Samp.CRYINDRS = "5" Then       ' 推定なら          ' 本当ならTOP／BOTで１回でいいはず ---１回にした(蔵本)
''''''                For i = 1 To 2
'''''                DoEvents
'''''                Call CryR_SetBaseSQL(sql)
'''''                DoEvents
'''''                Call AddSQL_SUITEI(sql, Samp.CRYNUM, Tname, TorB)
'''''                DoEvents
'''''
'''''                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                DoEvents
'''''                If rs.RecordCount <> 0 Then
'''''                    DoEvents
'''''                    Call CryR_ObjCpy(SuCryR, rs)
'''''                Else
'''''                    NothingFlag = True
'''''                End If
'''''                rs.Close
'''''                Set rs = Nothing
''''''                Next i
'''''            ElseIf Samp.CRYINDRS = "6" Then       ' 引継ぎなら
'''''                DoEvents
'''''                Call CryR_SetBaseSQL(sql)
'''''                DoEvents
'''''                Call AddSQL_HIKITUGI2(sql, Samp.CRYNUM, Samp.INGOTPOS, Tname, TorB)
'''''                DoEvents
'''''                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                DoEvents
'''''                If rs.RecordCount <> 0 Then
'''''                    recCnt = rs.RecordCount
'''''                    For i = 1 To recCnt
'''''                        If i = 1 Then                                     ' 一回目は保持
'''''                            DoEvents
'''''                            Call CryR_ObjCpy(CryR, rs)
'''''                            Exit For    '１レコード目だけでOK
'''''                        Else
'''''                            If CryR.POSITION = rs("POSITION") And CryR.REGDATE < rs("REGDATE") Then   ' 前の位置と同じだったら登録日付が新しいものをとる
'''''                                DoEvents
'''''                                Call CryR_ObjCpy(CryR, rs)
'''''                            End If
'''''                            Exit For
'''''                        End If
'''''                        rs.MoveNext
'''''                    Next
'''''                Else
'''''                    NothingFlag = True
'''''                End If
'''''                If Not rs Is Nothing Then
'''''                    rs.Close
'''''                    Set rs = Nothing
'''''                End If
'''''            End If  ' 検査指示が5 or 6 なら
''''''----- 8/12 野村 修正（比抵抗は仕様に関わらず結果を表示したい）
''''''元        End If  ' 指示がたっている
''''''        End If  ' 指示がたっている
''''''-----
'''''    End If ' 実績がある
'''''    If NothingFlagStr <> vbNullString Then
'''''        If NothingFlag Then
'''''            NothingFlagStr = "1"
'''''        End If
'''''    End If
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    CryR_Zisseki = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


''''''概要      :内部関数 Oi実績取得用オブジェクトコピー関数(レコードセットからのコピー)
'''''Private Sub Oi_ObjCpy(Oi As type_DBDRV_scmzc_fcmkc001c_Oi, rs As OraDynaset)
'''''    With Oi
'''''        .CRYNUM = rs("CRYNUM")         ' 結晶番号
'''''        .POSITION = rs("POSITION")     ' 位置
'''''        .SMPKBN = rs("SMPKBN")         ' サンプル区分
'''''        .SMPLNO = rs("SMPLNO")         ' サンプルＮｏ
'''''        .SMPLUMU = rs("SMPLUMU")       ' サンプル有無
'''''        .TRANCOND = rs("TRANCOND")     ' 処理条件
'''''        .OIMEAS1 = rs("OIMEAS1")       ' Ｏｉ測定値１
'''''        .OIMEAS2 = rs("OIMEAS2")       ' Ｏｉ測定値２
'''''        .OIMEAS3 = rs("OIMEAS3")       ' Ｏｉ測定値３
'''''        .OIMEAS4 = rs("OIMEAS4")       ' Ｏｉ測定値４
'''''        .OIMEAS5 = rs("OIMEAS5")       ' Ｏｉ測定値５
'''''        .ORGRES = rs("ORGRES")         ' ＯＲＧ結果
'''''        .AVE = rs("AVE")               ' ＡＶＥ
'''''        .FTIRCONV = rs("FTIRCONV")     ' ＦＴＩＲ換算
'''''        .INSPECTWAY = rs("INSPECTWAY") ' 検査方法
'''''        .REGDATE = rs("REGDATE")       ' 登録日付
'''''    End With
'''''End Sub
'''''
''''''概要      :内部関数 Oi実績取得用ベースSQL
'''''Private Sub Oi_SetBaseSQL(sql As String)
'''''    sql = "select "
'''''    sql = sql & "CRYNUM, "          ' 結晶番号
'''''    sql = sql & "POSITION, "        ' 位置
'''''    sql = sql & "SMPKBN, "          ' サンプル区分
'''''    sql = sql & "SMPLNO, "          ' サンプルＮｏ
'''''    sql = sql & "SMPLUMU, "         ' サンプル有無
'''''    sql = sql & "TRANCOND, "        ' 処理条件
'''''    sql = sql & "OIMEAS1, "         ' Ｏｉ測定値１
'''''    sql = sql & "OIMEAS2, "         ' Ｏｉ測定値２
'''''    sql = sql & "OIMEAS3, "         ' Ｏｉ測定値３
'''''    sql = sql & "OIMEAS4, "         ' Ｏｉ測定値４
'''''    sql = sql & "OIMEAS5, "         ' Ｏｉ測定値５
'''''    sql = sql & "ORGRES, "          ' ＯＲＧ結果
'''''    sql = sql & "AVE, "             ' ＡＶＥ
'''''    sql = sql & "FTIRCONV, "        ' ＦＴＩＲ換算
'''''    sql = sql & "INSPECTWAY, "      ' 検査方法
'''''    sql = sql & "REGDATE "          ' 登録日付
'''''End Sub


''''''概要      :内部関数 Oi実績取得用
'''''Private Function Oi_Zisseki(Siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                              Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
'''''                              Oi As type_DBDRV_scmzc_fcmkc001c_Oi, _
'''''                              TorB As Integer, _
'''''                              Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
'''''
'''''    Dim sql As String
'''''    Dim rs As OraDynaset
'''''    Dim i As Long
'''''    Dim recCnt As Integer
'''''    Dim Tname As String
'''''    Dim NothingFlag As Boolean
'''''
'''''    NothingFlag = False
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function Oi_Zisseki"
'''''
'''''    Oi_Zisseki = FUNCTION_RETURN_SUCCESS
'''''
'''''    Tname = "TBCMJ003"
'''''    Set rs = Nothing
'''''
'''''    ' Oi実績テーブルから値を取得
'''''    DoEvents
'''''    Call Oi_SetBaseSQL(sql)
'''''    DoEvents
'''''    Call AddSQL_Default2(sql, Samp.CRYNUM, Samp.INGOTPOS, Samp.SMPKBN, Tname)
'''''
'''''    DoEvents
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    DoEvents
'''''
'''''    If rs.RecordCount <> 0 Then
'''''        DoEvents
'''''        Call Oi_ObjCpy(Oi, rs)
'''''        rs.Close
'''''        Set rs = Nothing
'''''    Else
''''''----- 8/12 野村 修正（引継ぎのときは結果を表示したい）
''''''元        If Siyou.HSXONHWS = SIJI Then   ' 仕様の指示がたっている
''''''-----
'''''            If Samp.CRYINDOI = "6" Then       ' 引継ぎなら
'''''                DoEvents
'''''                Call Oi_SetBaseSQL(sql)
'''''                DoEvents
'''''                Call AddSQL_HIKITUGI2(sql, Samp.CRYNUM, Samp.INGOTPOS, Tname, TorB)
'''''
'''''                DoEvents
'''''                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                DoEvents
'''''
'''''                If rs.RecordCount <> 0 Then
'''''                    recCnt = rs.RecordCount
'''''                    For i = 1 To recCnt
'''''                        If i = 1 Then                                     ' 一回目は保持
'''''                            DoEvents
'''''                            Call Oi_ObjCpy(Oi, rs)
'''''                            Exit For    '１レコード目だけでOK
'''''                        Else
'''''                            If Oi.POSITION = rs("POSITION") And Oi.REGDATE < rs("REGDATE") Then   ' 前の位置と同じだったら登録日付が新しいものをとる
'''''                                DoEvents
'''''                                Call Oi_ObjCpy(Oi, rs)
'''''                            End If
'''''                            Exit For
'''''                        End If
'''''
'''''                        rs.MoveNext
'''''                    Next
'''''                Else
'''''                    NothingFlag = True
'''''                End If
'''''                If Not rs Is Nothing Then
'''''                    rs.Close
'''''                    Set rs = Nothing
'''''                End If
'''''            End If  ' 検査指示が 6 なら
''''''----- 8/12 野村 修正（引継ぎのときは結果を表示したい）
''''''元        End If ' 仕様の指示がたっている
''''''-----
'''''    End If ' 実績がある
'''''    If NothingFlagStr <> vbNullString Then
'''''        If NothingFlag Then
'''''            NothingFlagStr = "1"
'''''        End If
'''''    End If
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    Oi_Zisseki = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


''''''概要      :内部関数 BMD実績取得用オブジェクトコピー関数(レコードセットからのコピー)
'''''Private Sub BMD_ObjCpy(BMD As type_DBDRV_scmzc_fcmkc001c_BMD, rs As OraDynaset)
'''''    With BMD
'''''        .CRYNUM = rs("CRYNUM")          ' 結晶番号
'''''        .POSITION = rs("POSITION")      ' 位置
'''''        .SMPKBN = rs("SMPKBN")          ' サンプル区分
'''''        .SMPLNO = rs("SMPLNO")          ' サンプルＮｏ
'''''        .SMPLUMU = rs("SMPLUMU")        ' サンプル有無
'''''        .HTPRC = rs("HTPRC")            ' 熱処理方法
'''''        .KKSP = rs("KKSP")              ' 結晶欠陥測定位置
'''''        .KKSET = rs("KKSET")            ' 結晶欠陥測定条件＋選択ET代
'''''        .TRANCOND = rs("TRANCOND")      ' 処理条件
'''''        .MEAS1 = rs("MEAS1")            ' 測定値１
'''''        .MEAS2 = rs("MEAS2")            ' 測定値２
'''''        .MEAS3 = rs("MEAS3")            ' 測定値３
'''''        .MEAS4 = rs("MEAS4")            ' 測定値４
'''''        .MEAS5 = rs("MEAS5")            ' 測定値５
'''''        .Min = rs("MEASMIN")            ' MIN
'''''        .max = rs("MEASMAX")            ' MAX
'''''        .AVE = rs("MEASAVE")            ' AVE
'''''        .REGDATE = rs("REGDATE")        ' 登録日付
'''''
'''''' OSF，BMD項目追加対応  2002.04.02 yakimura
'''''         If IsNull(rs("BMDMNBUNP")) = False Then .BMDMNBUNP = rs("BMDMNBUNP")
'''''' OSF，BMD項目追加対応  2002.04.02 yakimura
'''''    End With
'''''End Sub


''''''概要      :内部関数 BMD実績取得用ベースSQL
'''''Private Sub BMD_SetBaseSQL(sql As String)
'''''    sql = "select "
'''''    sql = sql & "CRYNUM, "              ' 結晶番号
'''''    sql = sql & "POSITION, "            ' 位置
'''''    sql = sql & "SMPKBN, "              ' サンプル区分
'''''    sql = sql & "SMPLNO, "              ' サンプルＮｏ
'''''    sql = sql & "SMPLUMU, "             ' サンプル有無
'''''    sql = sql & "HTPRC,"                ' 熱処理方法"
'''''    sql = sql & "KKSP,"                 ' 結晶欠陥測定位置"
'''''    sql = sql & "KKSET,"                ' 結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)"
'''''    sql = sql & "TRANCOND, "            ' 処理条件
'''''    sql = sql & "MEAS1, "               ' 測定値１
'''''    sql = sql & "MEAS2, "               ' 測定値２
'''''    sql = sql & "MEAS3, "               ' 測定値３
'''''    sql = sql & "MEAS4, "               ' 測定値４
'''''    sql = sql & "MEAS5, "               ' 測定値５
'''''    sql = sql & "MEASMIN, "             ' MIN
'''''    sql = sql & "MEASMAX, "             ' MAX
'''''    sql = sql & "MEASAVE, "             ' AVE
'''''    sql = sql & "REGDATE,"              ' 登録日付
'''''' OSF，BMD項目追加対応  2002.04.02 yakimura
'''''    sql = sql & "BMDMNBUNP "            ' BMD面内分布
'''''' OSF，BMD項目追加対応  2002.04.02 yakimura
'''''End Sub



''''''概要      :内部関数 BMD実績取得用
'''''Private Function BMD_Zisseki(Siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                              Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
'''''                              inTRANCOND As Integer, _
'''''                              BMD As type_DBDRV_scmzc_fcmkc001c_BMD, _
'''''                              TorB As Integer, _
'''''                              Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
'''''
'''''    Dim sql As String
'''''    Dim rs As OraDynaset
'''''    Dim recCnt As Integer
'''''    Dim i As Long
'''''    Dim Tname As String
'''''    Dim NothingFlag As Boolean
'''''
'''''    NothingFlag = False
'''''
'''''    ' BMD実績テーブルから値を取得
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function BMD_Zisseki"
'''''
'''''    BMD_Zisseki = FUNCTION_RETURN_SUCCESS
'''''
'''''    Tname = "TBCMJ008"
'''''    Set rs = Nothing
'''''
'''''    DoEvents
'''''    Call BMD_SetBaseSQL(sql)
'''''    DoEvents
'''''    Call AddSQL_Default2(sql, Samp.CRYNUM, Samp.INGOTPOS, Samp.SMPKBN, Tname, " and TRANCOND='" & inTRANCOND & "' ")
'''''    DoEvents
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    DoEvents
'''''
'''''    If rs.RecordCount <> 0 Then
'''''        DoEvents
'''''        Call BMD_ObjCpy(BMD, rs)
'''''        rs.Close
'''''        Set rs = Nothing
'''''    Else
'''''        If (inTRANCOND = 1 And ((Siyou.HSXBM1HS = SIJI) Or (Siyou.HSXBM1HS = SANKOU))) _
'''''           Or (inTRANCOND = 2 And ((Siyou.HSXBM2HS = SIJI) Or (Siyou.HSXBM2HS = SANKOU))) _
'''''           Or (inTRANCOND = 3 And ((Siyou.HSXBM3HS = SIJI) Or (Siyou.HSXBM3HS = SANKOU))) Then           ' 仕様の指示がたっている
'''''            If (inTRANCOND = 1 And Samp.CRYINDB1 = "6") _
'''''               Or (inTRANCOND = 2 And Samp.CRYINDB2 = "6") _
'''''               Or (inTRANCOND = 3 And Samp.CRYINDB3 = "6") Then       ' 引継ぎなら
'''''                DoEvents
'''''                Call BMD_SetBaseSQL(sql)
'''''                DoEvents
'''''                Call AddSQL_HIKITUGI2(sql, Samp.CRYNUM, Samp.INGOTPOS, Tname, TorB, " and TRANCOND='" & inTRANCOND & "' ")
'''''
'''''                DoEvents
'''''                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                DoEvents
'''''
'''''                If rs.RecordCount <> 0 Then
'''''                    recCnt = rs.RecordCount
'''''                    For i = 1 To recCnt
'''''                        If i = 1 Then                                     ' 一回目は保持
'''''                            DoEvents
'''''                            Call BMD_ObjCpy(BMD, rs)
'''''                            Exit For    '１レコード目だけでOK
'''''                        Else
'''''                            If BMD.POSITION = rs("POSITION") And BMD.REGDATE < rs("REGDATE") Then   ' 前の位置と同じだったら登録日付が新しいものをとる
'''''                                DoEvents
'''''                                Call BMD_ObjCpy(BMD, rs)
'''''                            End If
'''''                            Exit For
'''''                        End If
'''''
'''''                        rs.MoveNext
'''''                    Next
'''''                Else
'''''                    NothingFlag = True
'''''                End If
'''''                If Not rs Is Nothing Then
'''''                    rs.Close
'''''                    Set rs = Nothing
'''''                End If
'''''            End If  ' 検査指示が5 or 6 なら
'''''        End If ' 指示がたっている
'''''    End If ' 実績があるかどうか
'''''
'''''    If NothingFlagStr <> vbNullString Then
'''''        If NothingFlag Then
'''''            NothingFlagStr = "1"
'''''        End If
'''''    End If
'''''
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    BMD_Zisseki = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


'''''' レコードセットからのコピー
''''''概要      :内部関数 BMD実績取得用オブジェクトコピー関数(レコードセットからのコピー)
'''''Private Sub OSF_ObjCpy(OSF As type_DBDRV_scmzc_fcmkc001c_OSF, rs As OraDynaset)
'''''    With OSF
'''''        .CRYNUM = rs("CRYNUM")         ' 結晶番号
'''''        .POSITION = rs("POSITION")     ' 位置
'''''        .SMPKBN = rs("SMPKBN")         ' サンプル区分
'''''        .SMPLNO = rs("SMPLNO")         ' サンプルＮｏ
'''''        .SMPLUMU = rs("SMPLUMU")       ' サンプル有無
'''''        .HTPRC = rs("HTPRC")           ' 熱処理方法
'''''        .KKSP = rs("KKSP")             ' 結晶欠陥測定位置
'''''        .KKSET = rs("KKSET")           ' 結晶欠陥測定条件＋選択ET代
'''''        .TRANCOND = rs("TRANCOND")     ' 処理条件
'''''        .CALCMAX = rs("CALCMAX")       ' 計算結果 Max
'''''        .CALCAVE = rs("CALCAVE")       ' 計算結果 Ave
'''''        .MEAS1 = rs("MEAS1")           ' 測定値１
'''''        .MEAS2 = rs("MEAS2")           ' 測定値２
'''''        .MEAS3 = rs("MEAS3")           ' 測定値３
'''''        .MEAS4 = rs("MEAS4")           ' 測定値４
'''''        .MEAS5 = rs("MEAS5")           ' 測定値５
'''''        .MEAS6 = rs("MEAS6")           ' 測定値６
'''''        .MEAS7 = rs("MEAS7")           ' 測定値７
'''''        .MEAS8 = rs("MEAS8")           ' 測定値８
'''''        .MEAS9 = rs("MEAS9")           ' 測定値９
'''''        .MEAS10 = rs("MEAS10")         ' 測定値１０
'''''        .MEAS11 = rs("MEAS11")         ' 測定値１１
'''''        .MEAS12 = rs("MEAS12")         ' 測定値１２
'''''        .MEAS13 = rs("MEAS13")         ' 測定値１３
'''''        .MEAS14 = rs("MEAS14")         ' 測定値１４
'''''        .MEAS15 = rs("MEAS15")         ' 測定値１５
'''''        .MEAS16 = rs("MEAS16")         ' 測定値１６
'''''        .MEAS17 = rs("MEAS17")         ' 測定値１７
'''''        .MEAS18 = rs("MEAS18")         ' 測定値１８
'''''        .MEAS19 = rs("MEAS19")         ' 測定値１９
'''''        .MEAS20 = rs("MEAS20")         ' 測定値２０
'''''        .REGDATE = rs("REGDATE")       ' 登録日付
'''''
'''''' OSF，BMD項目追加対応  2002.04.02 yakimura
'''''         If IsNull(rs("OSFPOS1")) = False Then .OSFPOS1 = rs("OSFPOS1")   'ﾊﾟﾀｰﾝ区分１位置
'''''         If IsNull(rs("OSFWID1")) = False Then .OSFWID1 = rs("OSFWID1")   'ﾊﾟﾀｰﾝ区分１幅
'''''         If IsNull(rs("OSFRD1")) = False Then .OSFRD1 = rs("OSFRD1")      'ﾊﾟﾀｰﾝ区分１R/D
'''''         If IsNull(rs("OSFPOS2")) = False Then .OSFPOS2 = rs("OSFPOS2")   'ﾊﾟﾀｰﾝ区分２位置
'''''         If IsNull(rs("OSFWID2")) = False Then .OSFWID2 = rs("OSFWID2")   'ﾊﾟﾀｰﾝ区分２幅
'''''         If IsNull(rs("OSFRD2")) = False Then .OSFRD2 = rs("OSFRD2")      'ﾊﾟﾀｰﾝ区分２R/D
'''''         If IsNull(rs("OSFPOS3")) = False Then .OSFPOS3 = rs("OSFPOS3")   'ﾊﾟﾀｰﾝ区分３位置
'''''         If IsNull(rs("OSFWID3")) = False Then .OSFWID3 = rs("OSFWID3")   'ﾊﾟﾀｰﾝ区分３幅
'''''         If IsNull(rs("OSFRD3")) = False Then .OSFRD3 = rs("OSFRD3")      'ﾊﾟﾀｰﾝ区分３R/D
'''''' OSF，BMD項目追加対応  2002.04.02 yakimura
'''''    End With
'''''End Sub
'''''
''''''概要      :内部関数 BMD実績取得用ベースSQL
'''''Private Sub OSF_SetBaseSQL(sql As String)
'''''    sql = "select "
'''''    sql = sql & "CRYNUM, "        ' 結晶番号
'''''    sql = sql & "POSITION, "      ' 位置
'''''    sql = sql & "SMPKBN, "        ' サンプル区分
'''''    sql = sql & "SMPLNO, "        ' サンプルＮｏ
'''''    sql = sql & "SMPLUMU, "       ' サンプル有無
'''''    sql = sql & "HTPRC,"          ' 熱処理方法"
'''''    sql = sql & "KKSP,"           ' 結晶欠陥測定位置"
'''''    sql = sql & "KKSET,"          ' 結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)"
'''''    sql = sql & "TRANCOND, "      ' 処理条件
'''''    sql = sql & "CALCMAX, "       ' 計算結果 Max
'''''    sql = sql & "CALCAVE, "       ' 計算結果 Ave
'''''    sql = sql & "MEAS1, "         ' 測定値１
'''''    sql = sql & "MEAS2, "         ' 測定値２
'''''    sql = sql & "MEAS3, "         ' 測定値３
'''''    sql = sql & "MEAS4, "         ' 測定値４
'''''    sql = sql & "MEAS5, "         ' 測定値５
'''''    sql = sql & "MEAS6, "         ' 測定値６
'''''    sql = sql & "MEAS7, "         ' 測定値７
'''''    sql = sql & "MEAS8, "         ' 測定値８
'''''    sql = sql & "MEAS9, "         ' 測定値９
'''''    sql = sql & "MEAS10, "        ' 測定値１０
'''''    sql = sql & "MEAS11, "        ' 測定値１１
'''''    sql = sql & "MEAS12, "        ' 測定値１２
'''''    sql = sql & "MEAS13, "        ' 測定値１３
'''''    sql = sql & "MEAS14, "        ' 測定値１４
'''''    sql = sql & "MEAS15, "        ' 測定値１５
'''''    sql = sql & "MEAS16, "        ' 測定値１６
'''''    sql = sql & "MEAS17, "        ' 測定値１７
'''''    sql = sql & "MEAS18, "        ' 測定値１８
'''''    sql = sql & "MEAS19, "        ' 測定値１９
'''''    sql = sql & "MEAS20, "        ' 測定値２０
'''''    sql = sql & "REGDATE, "       ' 登録日付
'''''' OSF，BMD項目追加対応  2002.04.02 yakimura
'''''    sql = sql & "OSFPOS1, "       ' ﾊﾟﾀｰﾝ区分１位置
'''''    sql = sql & "OSFWID1, "       ' ﾊﾟﾀｰﾝ区分１幅
'''''    sql = sql & "OSFRD1, "        ' ﾊﾟﾀｰﾝ区分１R/D
'''''    sql = sql & "OSFPOS2, "       ' ﾊﾟﾀｰﾝ区分２位置
'''''    sql = sql & "OSFWID2, "       ' ﾊﾟﾀｰﾝ区分２幅
'''''    sql = sql & "OSFRD2, "        ' ﾊﾟﾀｰﾝ区分２R/D
'''''    sql = sql & "OSFPOS3, "       ' ﾊﾟﾀｰﾝ区分３位置
'''''    sql = sql & "OSFWID3, "       ' ﾊﾟﾀｰﾝ区分３幅
'''''    sql = sql & "OSFRD3 "         ' ﾊﾟﾀｰﾝ区分３R/D
'''''' OSF，BMD項目追加対応  2002.04.02 yakimura
'''''End Sub


''''''概要      :内部関数 OSF実績取得用
'''''Private Function OSF_Zisseki(Siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                              Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
'''''                              inTRANCOND As Integer, _
'''''                              OSF As type_DBDRV_scmzc_fcmkc001c_OSF, _
'''''                              TorB As Integer, _
'''''                              Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
'''''    Dim sql As String
'''''    Dim rs As OraDynaset
'''''    Dim recCnt As Integer
'''''    Dim i As Long
'''''    Dim Tname As String
'''''    Dim NothingFlag As Boolean
'''''
'''''    NothingFlag = False
'''''
'''''
'''''    ' OSF実績テーブルから値を取得
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function OSF_Zisseki"
'''''
'''''    OSF_Zisseki = FUNCTION_RETURN_SUCCESS
'''''
'''''    Tname = "TBCMJ005"
'''''    Set rs = Nothing
'''''
'''''    DoEvents
'''''    Call OSF_SetBaseSQL(sql)
'''''    DoEvents
'''''    Call AddSQL_Default2(sql, Samp.CRYNUM, Samp.INGOTPOS, Samp.SMPKBN, Tname, " and TRANCOND='" & inTRANCOND & "' ")
'''''    DoEvents
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    DoEvents
'''''
'''''    If rs.RecordCount <> 0 Then
'''''        DoEvents
'''''        Call OSF_ObjCpy(OSF, rs)
'''''        rs.Close
'''''        Set rs = Nothing
'''''    Else
'''''        rs.Close
'''''        Set rs = Nothing
'''''        If (inTRANCOND = 1 And ((Siyou.HSXOS1HS = SIJI) Or (Siyou.HSXOS1HS = SANKOU))) _
'''''           Or (inTRANCOND = 2 And ((Siyou.HSXOS2HS = SIJI) Or (Siyou.HSXOS2HS = SANKOU))) _
'''''           Or (inTRANCOND = 3 And ((Siyou.HSXOS3HS = SIJI) Or (Siyou.HSXOS3HS = SANKOU))) Then          ' 仕様の指示がたっている
'''''           If (inTRANCOND = 1 And Samp.CRYINDL1 = "6") _
'''''              Or (inTRANCOND = 2 And Samp.CRYINDL2 = "6") _
'''''              Or (inTRANCOND = 3 And Samp.CRYINDL3 = "6") Then       ' 引継ぎなら
'''''                DoEvents
'''''                Call OSF_SetBaseSQL(sql)
'''''                DoEvents
'''''                Call AddSQL_HIKITUGI2(sql, Samp.CRYNUM, Samp.INGOTPOS, Tname, TorB, " and TRANCOND='" & inTRANCOND & "' ")
'''''                DoEvents
'''''                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                DoEvents
'''''
'''''                If rs.RecordCount <> 0 Then
'''''                    recCnt = rs.RecordCount
'''''                    For i = 1 To recCnt
'''''                        If i = 1 Then                                     ' 一回目は保持
'''''                            DoEvents
'''''                            Call OSF_ObjCpy(OSF, rs)
'''''                            Exit For    '１レコード目だけでOK
'''''                        Else
'''''                            If OSF.POSITION = rs("POSITION") And OSF.REGDATE < rs("REGDATE") Then   ' 前の位置と同じだったら登録日付が新しいものをとる
'''''                                DoEvents
'''''                                Call OSF_ObjCpy(OSF, rs)
'''''                            End If
'''''                        End If
'''''                        rs.MoveNext
'''''                    Next
'''''                Else
'''''                    NothingFlag = True
'''''                End If
'''''                If Not rs Is Nothing Then
'''''                    rs.Close
'''''                    Set rs = Nothing
'''''                End If
'''''            End If  ' 検査指示が5 or 6 なら
'''''        End If  ' 指示がたっている
'''''    End If ' 実績がある
'''''
'''''    If NothingFlagStr <> vbNullString Then
'''''        If NothingFlag Then
'''''            NothingFlagStr = "1"
'''''        End If
'''''    End If
'''''
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    OSF_Zisseki = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


'''''Private Sub Cs_ObjCpy(Cs As type_DBDRV_scmzc_fcmkc001c_CS, rs As OraDynaset)
'''''    With Cs
'''''        .CRYNUM = rs("CRYNUM")         ' 結晶番号
'''''        .POSITION = rs("POSITION")     ' 位置
'''''        .SMPKBN = rs("SMPKBN")         ' サンプル区分
'''''        .SMPLNO = rs("SMPLNO")         ' サンプルＮｏ
'''''        .SMPLUMU = rs("SMPLUMU")       ' サンプル有無
'''''        .TRANCOND = rs("TRANCOND")     ' 処理条件
'''''        .CSMEAS = rs("CSMEAS")         ' Cs実測値
'''''        .PRE70P = rs("PRE70P")         ' ７０％推定値
'''''        .REGDATE = rs("REGDATE")        ' 登録日付
'''''    End With
'''''
'''''End Sub


'''''Private Sub Cs_SetBaseSQL(sql As String)
'''''    sql = "select "
'''''    sql = sql & "CRYNUM, "        ' 結晶番号
'''''    sql = sql & "POSITION, "      ' 位置
'''''    sql = sql & "SMPKBN, "        ' サンプル区分
'''''    sql = sql & "SMPLNO, "        ' サンプルＮｏ
'''''    sql = sql & "SMPLUMU, "       ' サンプル有無
'''''    sql = sql & "TRANCOND, "      ' 処理条件
'''''    sql = sql & "CSMEAS, "        ' Cs実測値
'''''    sql = sql & "PRE70P, "         ' ７０％推定値
'''''    sql = sql & "REGDATE "        ' 登録日付
'''''
'''''End Sub


''''''内部関数 Cs実績取得用
'''''Private Function CS_Zisseki(CRYNUM As String, Samp() As type_DBDRV_scmzc_fcmkc001c_CrySmp, Cs() As type_DBDRV_scmzc_fcmkc001c_CS) As FUNCTION_RETURN
'''''Dim sql As String
'''''Dim rs As OraDynaset
'''''Dim recCnt As Integer
'''''Dim i As Long
'''''Dim Tname As String
'''''Dim jCs As String
'''''Dim jCsFromTo As String
'''''Dim tt As Integer
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function CS_Zisseki"
'''''    CS_Zisseki = FUNCTION_RETURN_FAILURE
'''''
'''''    'ブロック内の品番についてCs仕様の有無をチェックする
'''''    If DBDRV_scmzc_fcmkc001c_CheckSpecCs(CRYNUM, Samp(1).INGOTPOS, Samp(2).INGOTPOS, jCs, jCsFromTo) = FUNCTION_RETURN_FAILURE Then
'''''        GoTo proc_err
'''''    End If
'''''
'''''    For tt = 1 To 2
'''''        With Cs(tt)
'''''            .CRYNUM = vbNullString
'''''            .CSMEAS = -1
'''''            .POSITION = Samp(tt).INGOTPOS
'''''            .PRE70P = -1
'''''            .SMPLNO = -1
'''''            .SMPLUMU = "0"
'''''        End With
'''''    Next
'''''    If (jCsFromTo = SIJI) Or (jCsFromTo = SANKOU) Then 'FromTo仕様を含む品番があるため、引継不可
'''''        For tt = 1 To 2
'''''            Tname = "TBCMJ004"
'''''            Call Cs_SetBaseSQL(sql)
'''''            Call AddSQL_Default2(sql, CRYNUM, Samp(tt).INGOTPOS, Samp(tt).SMPKBN, Tname)
'''''            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''            If rs.RecordCount > 0 Then
'''''                Call Cs_ObjCpy(Cs(tt), rs)
'''''            End If
'''''            rs.Close
'''''            Set rs = Nothing
'''''        Next
'''''    ElseIf (jCs = SIJI) Or (jCs = SANKOU) Then         'Cs仕様を含む品番がある。Tail側のみ下側から実績取得
'''''        tt = 2
'''''        Tname = "TBCMJ004"
'''''        Call Cs_SetBaseSQL(sql)
'''''        sql = sql & " from TBCMJ004 T1"
'''''        sql = sql & " where T1.CRYNUM='" & CRYNUM & "'"
'''''        sql = sql & " and T1.TRANCNT=(select max(TRANCNT) from TBCMJ004 where CRYNUM=T1.CRYNUM and POSITION=T1.POSITION and SMPKBN=T1.SMPKBN)"
'''''        sql = sql & " and T1.SMPLUMU='0'"
'''''        sql = sql & " and T1.POSITION>=" & Samp(tt).INGOTPOS
'''''        sql = sql & " order by POSITION asc, SMPKBN asc"
'''''        sql = "select * from (" & sql & ") where rownum=1"
'''''        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''        If rs.RecordCount > 0 Then
'''''            Call Cs_ObjCpy(Cs(tt), rs)
'''''        End If
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''
'''''    CS_Zisseki = FUNCTION_RETURN_SUCCESS
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


'''''Private Sub GD_ObjCpy(GD As type_DBDRV_scmzc_fcmkc001c_GD, rs As OraDynaset)
'''''    With GD
'''''        .CRYNUM = rs("CRYNUM")         ' 結晶番号
'''''        .POSITION = rs("POSITION")     ' 位置
'''''        .SMPKBN = rs("SMPKBN")         ' サンプル区分
'''''        .SMPLNO = rs("SMPLNO")         ' サンプルＮｏ
'''''        .SMPLUMU = rs("SMPLUMU")       ' サンプル有無
'''''        .TRANCOND = rs("TRANCOND")     ' 処理条件
'''''        .MSRSDEN = rs("MSRSDEN")       ' 測定結果 Den
'''''        .MSRSLDL = rs("MSRSLDL")       ' 測定結果 L/DL
'''''        .MSRSDVD2 = rs("MSRSDVD2")     ' 測定結果 DVD2
'''''        .MS01LDL1 = rs("MS01LDL1")            ' 測定値01 L/DL1
'''''        .MS01LDL2 = rs("MS01LDL2")            ' 測定値01 L/DL2
'''''        .MS01LDL3 = rs("MS01LDL3")            ' 測定値01 L/DL3
'''''        .MS01LDL4 = rs("MS01LDL4")            ' 測定値01 L/DL4
'''''        .MS01LDL5 = rs("MS01LDL5")            ' 測定値01 L/DL5
'''''        .MS01DEN1 = rs("MS01DEN1")            ' 測定値01 Den1
'''''        .MS01DEN2 = rs("MS01DEN2")            ' 測定値01 Den2
'''''        .MS01DEN3 = rs("MS01DEN3")            ' 測定値01 Den3
'''''        .MS01DEN4 = rs("MS01DEN4")            ' 測定値01 Den4
'''''        .MS01DEN5 = rs("MS01DEN5")            ' 測定値01 Den5
'''''        .MS02LDL1 = rs("MS02LDL1")            ' 測定値02 L/DL1
'''''        .MS02LDL2 = rs("MS02LDL2")            ' 測定値02 L/DL2
'''''        .MS02LDL3 = rs("MS02LDL3")            ' 測定値02 L/DL3
'''''        .MS02LDL4 = rs("MS02LDL4")            ' 測定値02 L/DL4
'''''        .MS02LDL5 = rs("MS02LDL5")            ' 測定値02 L/DL5
'''''        .MS02DEN1 = rs("MS02DEN1")            ' 測定値02 Den1
'''''        .MS02DEN2 = rs("MS02DEN2")            ' 測定値02 Den2
'''''        .MS02DEN3 = rs("MS02DEN3")            ' 測定値02 Den3
'''''        .MS02DEN4 = rs("MS02DEN4")            ' 測定値02 Den4
'''''        .MS02DEN5 = rs("MS02DEN5")            ' 測定値02 Den5
'''''        .MS03LDL1 = rs("MS03LDL1")            ' 測定値03 L/DL1
'''''        .MS03LDL2 = rs("MS03LDL2")            ' 測定値03 L/DL2
'''''        .MS03LDL3 = rs("MS03LDL3")            ' 測定値03 L/DL3
'''''        .MS03LDL4 = rs("MS03LDL4")            ' 測定値03 L/DL4
'''''        .MS03LDL5 = rs("MS03LDL5")            ' 測定値03 L/DL5
'''''        .MS03DEN1 = rs("MS03DEN1")            ' 測定値03 Den1
'''''        .MS03DEN2 = rs("MS03DEN2")            ' 測定値03 Den2
'''''        .MS03DEN3 = rs("MS03DEN3")            ' 測定値03 Den3
'''''        .MS03DEN4 = rs("MS03DEN4")            ' 測定値03 Den4
'''''        .MS03DEN5 = rs("MS03DEN5")            ' 測定値03 Den5
'''''        .MS04LDL1 = rs("MS04LDL1")            ' 測定値04 L/DL1
'''''        .MS04LDL2 = rs("MS04LDL2")            ' 測定値04 L/DL2
'''''        .MS04LDL3 = rs("MS04LDL3")            ' 測定値04 L/DL3
'''''        .MS04LDL4 = rs("MS04LDL4")            ' 測定値04 L/DL4
'''''        .MS04LDL5 = rs("MS04LDL5")            ' 測定値04 L/DL5
'''''        .MS04DEN1 = rs("MS04DEN1")            ' 測定値04 Den1
'''''        .MS04DEN2 = rs("MS04DEN2")            ' 測定値04 Den2
'''''        .MS04DEN3 = rs("MS04DEN3")            ' 測定値04 Den3
'''''        .MS04DEN4 = rs("MS04DEN4")            ' 測定値04 Den4
'''''        .MS04DEN5 = rs("MS04DEN5")            ' 測定値04 Den5
'''''        .MS05LDL1 = rs("MS05LDL1")            ' 測定値05 L/DL1
'''''        .MS05LDL2 = rs("MS05LDL2")            ' 測定値05 L/DL2
'''''        .MS05LDL3 = rs("MS05LDL3")            ' 測定値05 L/DL3
'''''        .MS05LDL4 = rs("MS05LDL4")            ' 測定値05 L/DL4
'''''        .MS05LDL5 = rs("MS05LDL5")            ' 測定値05 L/DL5
'''''        .MS05DEN1 = rs("MS05DEN1")            ' 測定値05 Den1
'''''        .MS05DEN2 = rs("MS05DEN2")            ' 測定値05 Den2
'''''        .MS05DEN3 = rs("MS05DEN3")            ' 測定値05 Den3
'''''        .MS05DEN4 = rs("MS05DEN4")            ' 測定値05 Den4
'''''        .MS05DEN5 = rs("MS05DEN5")            ' 測定値05 Den5
'''''        .MS06LDL1 = rs("MS06LDL1")            ' 測定値06 L/DL1
'''''        .MS06LDL2 = rs("MS06LDL2")            ' 測定値06 L/DL2
'''''        .MS06LDL3 = rs("MS06LDL3")            ' 測定値06 L/DL3
'''''        .MS06LDL4 = rs("MS06LDL4")            ' 測定値06 L/DL4
'''''        .MS06LDL5 = rs("MS06LDL5")            ' 測定値06 L/DL5
'''''        .MS06DEN1 = rs("MS06DEN1")            ' 測定値06 Den1
'''''        .MS06DEN2 = rs("MS06DEN2")            ' 測定値06 Den2
'''''        .MS06DEN3 = rs("MS06DEN3")            ' 測定値06 Den3
'''''        .MS06DEN4 = rs("MS06DEN4")            ' 測定値06 Den4
'''''        .MS06DEN5 = rs("MS06DEN5")            ' 測定値06 Den5
'''''        .MS07LDL1 = rs("MS07LDL1")            ' 測定値07 L/DL1
'''''        .MS07LDL2 = rs("MS07LDL2")            ' 測定値07 L/DL2
'''''        .MS07LDL3 = rs("MS07LDL3")            ' 測定値07 L/DL3
'''''        .MS07LDL4 = rs("MS07LDL4")            ' 測定値07 L/DL4
'''''        .MS07LDL5 = rs("MS07LDL5")            ' 測定値07 L/DL5
'''''        .MS07DEN1 = rs("MS07DEN1")            ' 測定値07 Den1
'''''        .MS07DEN2 = rs("MS07DEN2")            ' 測定値07 Den2
'''''        .MS07DEN3 = rs("MS07DEN3")            ' 測定値07 Den3
'''''        .MS07DEN4 = rs("MS07DEN4")            ' 測定値07 Den4
'''''        .MS07DEN5 = rs("MS07DEN5")            ' 測定値07 Den5
'''''        .MS08LDL1 = rs("MS08LDL1")            ' 測定値08 L/DL1
'''''        .MS08LDL2 = rs("MS08LDL2")            ' 測定値08 L/DL2
'''''        .MS08LDL3 = rs("MS08LDL3")            ' 測定値08 L/DL3
'''''        .MS08LDL4 = rs("MS08LDL4")            ' 測定値08 L/DL4
'''''        .MS08LDL5 = rs("MS08LDL5")            ' 測定値08 L/DL5
'''''        .MS08DEN1 = rs("MS08DEN1")            ' 測定値08 Den1
'''''        .MS08DEN2 = rs("MS08DEN2")            ' 測定値08 Den2
'''''        .MS08DEN3 = rs("MS08DEN3")            ' 測定値08 Den3
'''''        .MS08DEN4 = rs("MS08DEN4")            ' 測定値08 Den4
'''''        .MS08DEN5 = rs("MS08DEN5")            ' 測定値08 Den5
'''''        .MS09LDL1 = rs("MS09LDL1")            ' 測定値09 L/DL1
'''''        .MS09LDL2 = rs("MS09LDL2")            ' 測定値09 L/DL2
'''''        .MS09LDL3 = rs("MS09LDL3")            ' 測定値09 L/DL3
'''''        .MS09LDL4 = rs("MS09LDL4")            ' 測定値09 L/DL4
'''''        .MS09LDL5 = rs("MS09LDL5")            ' 測定値09 L/DL5
'''''        .MS09DEN1 = rs("MS09DEN1")            ' 測定値09 Den1
'''''        .MS09DEN2 = rs("MS09DEN2")            ' 測定値09 Den2
'''''        .MS09DEN3 = rs("MS09DEN3")            ' 測定値09 Den3
'''''        .MS09DEN4 = rs("MS09DEN4")            ' 測定値09 Den4
'''''        .MS09DEN5 = rs("MS09DEN5")            ' 測定値09 Den5
'''''        .MS10LDL1 = rs("MS10LDL1")            ' 測定値10 L/DL1
'''''        .MS10LDL2 = rs("MS10LDL2")            ' 測定値10 L/DL2
'''''        .MS10LDL3 = rs("MS10LDL3")            ' 測定値10 L/DL3
'''''        .MS10LDL4 = rs("MS10LDL4")            ' 測定値10 L/DL4
'''''        .MS10LDL5 = rs("MS10LDL5")            ' 測定値10 L/DL5
'''''        .MS10DEN1 = rs("MS10DEN1")            ' 測定値10 Den1
'''''        .MS10DEN2 = rs("MS10DEN2")            ' 測定値10 Den2
'''''        .MS10DEN3 = rs("MS10DEN3")            ' 測定値10 Den3
'''''        .MS10DEN4 = rs("MS10DEN4")            ' 測定値10 Den4
'''''        .MS10DEN5 = rs("MS10DEN5")            ' 測定値10 Den5
'''''        .MS11LDL1 = rs("MS11LDL1")            ' 測定値11 L/DL1
'''''        .MS11LDL2 = rs("MS11LDL2")            ' 測定値11 L/DL2
'''''        .MS11LDL3 = rs("MS11LDL3")            ' 測定値11 L/DL3
'''''        .MS11LDL4 = rs("MS11LDL4")            ' 測定値11 L/DL4
'''''        .MS11LDL5 = rs("MS11LDL5")            ' 測定値11 L/DL5
'''''        .MS11DEN1 = rs("MS11DEN1")            ' 測定値11 Den1
'''''        .MS11DEN2 = rs("MS11DEN2")            ' 測定値11 Den2
'''''        .MS11DEN3 = rs("MS11DEN3")            ' 測定値11 Den3
'''''        .MS11DEN4 = rs("MS11DEN4")            ' 測定値11 Den4
'''''        .MS11DEN5 = rs("MS11DEN5")            ' 測定値11 Den5
'''''        .MS12LDL1 = rs("MS12LDL1")            ' 測定値12 L/DL1
'''''        .MS12LDL2 = rs("MS12LDL2")            ' 測定値12 L/DL2
'''''        .MS12LDL3 = rs("MS12LDL3")            ' 測定値12 L/DL3
'''''        .MS12LDL4 = rs("MS12LDL4")            ' 測定値12 L/DL4
'''''        .MS12LDL5 = rs("MS12LDL5")            ' 測定値12 L/DL5
'''''        .MS12DEN1 = rs("MS12DEN1")            ' 測定値12 Den1
'''''        .MS12DEN2 = rs("MS12DEN2")            ' 測定値12 Den2
'''''        .MS12DEN3 = rs("MS12DEN3")            ' 測定値12 Den3
'''''        .MS12DEN4 = rs("MS12DEN4")            ' 測定値12 Den4
'''''        .MS12DEN5 = rs("MS12DEN5")            ' 測定値12 Den5
'''''        .MS13LDL1 = rs("MS13LDL1")            ' 測定値13 L/DL1
'''''        .MS13LDL2 = rs("MS13LDL2")            ' 測定値13 L/DL2
'''''        .MS13LDL3 = rs("MS13LDL3")            ' 測定値13 L/DL3
'''''        .MS13LDL4 = rs("MS13LDL4")            ' 測定値13 L/DL4
'''''        .MS13LDL5 = rs("MS13LDL5")            ' 測定値13 L/DL5
'''''        .MS13DEN1 = rs("MS13DEN1")            ' 測定値13 Den1
'''''        .MS13DEN2 = rs("MS13DEN2")            ' 測定値13 Den2
'''''        .MS13DEN3 = rs("MS13DEN3")            ' 測定値13 Den3
'''''        .MS13DEN4 = rs("MS13DEN4")            ' 測定値13 Den4
'''''        .MS13DEN5 = rs("MS13DEN5")            ' 測定値13 Den5
'''''        .MS14LDL1 = rs("MS14LDL1")            ' 測定値14 L/DL1
'''''        .MS14LDL2 = rs("MS14LDL2")            ' 測定値14 L/DL2
'''''        .MS14LDL3 = rs("MS14LDL3")            ' 測定値14 L/DL3
'''''        .MS14LDL4 = rs("MS14LDL4")            ' 測定値14 L/DL4
'''''        .MS14LDL5 = rs("MS14LDL5")            ' 測定値14 L/DL5
'''''        .MS14DEN1 = rs("MS14DEN1")            ' 測定値14 Den1
'''''        .MS14DEN2 = rs("MS14DEN2")            ' 測定値14 Den2
'''''        .MS14DEN3 = rs("MS14DEN3")            ' 測定値14 Den3
'''''        .MS14DEN4 = rs("MS14DEN4")            ' 測定値14 Den4
'''''        .MS14DEN5 = rs("MS14DEN5")            ' 測定値14 Den5
'''''        .MS15LDL1 = rs("MS15LDL1")            ' 測定値15 L/DL1
'''''        .MS15LDL2 = rs("MS15LDL2")            ' 測定値15 L/DL2
'''''        .MS15LDL3 = rs("MS15LDL3")            ' 測定値15 L/DL3
'''''        .MS15LDL4 = rs("MS15LDL4")            ' 測定値15 L/DL4
'''''        .MS15LDL5 = rs("MS15LDL5")            ' 測定値15 L/DL5
'''''        .MS15DEN1 = rs("MS15DEN1")            ' 測定値15 Den1
'''''        .MS15DEN2 = rs("MS15DEN2")            ' 測定値15 Den2
'''''        .MS15DEN3 = rs("MS15DEN3")            ' 測定値15 Den3
'''''        .MS15DEN4 = rs("MS15DEN4")            ' 測定値15 Den4
'''''        .MS15DEN5 = rs("MS15DEN5")            ' 測定値15 Den5
'''''        .REGDATE = rs("REGDATE")              ' 登録日付
'''''    End With
'''''
'''''End Sub
'''''
'''''Private Sub GD_SetBaseSQL(sql As String)
'''''    ' GD実績テーブルから値を取得
'''''    sql = "select "
'''''    sql = sql & "CRYNUM, "               ' 結晶番号
'''''    sql = sql & "POSITION, "             ' 位置
'''''    sql = sql & "SMPKBN, "               ' サンプル区分
'''''    sql = sql & "SMPLNO, "               ' サンプルＮｏ
'''''    sql = sql & "SMPLUMU, "              ' サンプル有無
'''''    sql = sql & "TRANCOND, "             ' 処理条件
'''''    sql = sql & "MSRSDEN, "              ' 測定結果 Den
'''''    sql = sql & "MSRSLDL, "              ' 測定結果 L/DL
'''''    sql = sql & "MSRSDVD2, "             ' 測定結果 DVD2
'''''    sql = sql & "MS01LDL1, "             ' 測定値01 L/DL1"
'''''    sql = sql & "MS01LDL2, "             ' 測定値01 L/DL2"
'''''    sql = sql & "MS01LDL3, "             ' 測定値01 L/DL3"
'''''    sql = sql & "MS01LDL4, "             ' 測定値01 L/DL4"
'''''    sql = sql & "MS01LDL5, "             ' 測定値01 L/DL5"
'''''    sql = sql & "MS01DEN1, "             ' 測定値01 Den1"
'''''    sql = sql & "MS01DEN2, "             ' 測定値01 Den2"
'''''    sql = sql & "MS01DEN3, "             ' 測定値01 Den3"
'''''    sql = sql & "MS01DEN4, "             ' 測定値01 Den4"
'''''    sql = sql & "MS01DEN5, "             ' 測定値01 Den5"
'''''    sql = sql & "MS02LDL1, "             ' 測定値02 L/DL1"
'''''    sql = sql & "MS02LDL2, "             ' 測定値02 L/DL2"
'''''    sql = sql & "MS02LDL3, "             ' 測定値02 L/DL3"
'''''    sql = sql & "MS02LDL4, "             ' 測定値02 L/DL4"
'''''    sql = sql & "MS02LDL5, "             ' 測定値02 L/DL5"
'''''    sql = sql & "MS02DEN1, "             ' 測定値02 Den1"
'''''    sql = sql & "MS02DEN2, "             ' 測定値02 Den2"
'''''    sql = sql & "MS02DEN3, "             ' 測定値02 Den3"
'''''    sql = sql & "MS02DEN4, "             ' 測定値02 Den4"
'''''    sql = sql & "MS02DEN5, "             ' 測定値02 Den5"
'''''    sql = sql & "MS03LDL1, "             ' 測定値03 L/DL1"
'''''    sql = sql & "MS03LDL2, "             ' 測定値03 L/DL2"
'''''    sql = sql & "MS03LDL3, "             ' 測定値03 L/DL3"
'''''    sql = sql & "MS03LDL4, "             ' 測定値03 L/DL4"
'''''    sql = sql & "MS03LDL5, "             ' 測定値03 L/DL5"
'''''    sql = sql & "MS03DEN1, "             ' 測定値03 Den1"
'''''    sql = sql & "MS03DEN2, "             ' 測定値03 Den2"
'''''    sql = sql & "MS03DEN3, "             ' 測定値03 Den3"
'''''    sql = sql & "MS03DEN4, "             ' 測定値03 Den4"
'''''    sql = sql & "MS03DEN5, "             ' 測定値03 Den5"
'''''    sql = sql & "MS04LDL1, "             ' 測定値04 L/DL1"
'''''    sql = sql & "MS04LDL2, "             ' 測定値04 L/DL2"
'''''    sql = sql & "MS04LDL3, "             ' 測定値04 L/DL3"
'''''    sql = sql & "MS04LDL4, "             ' 測定値04 L/DL4"
'''''    sql = sql & "MS04LDL5, "             ' 測定値04 L/DL5"
'''''    sql = sql & "MS04DEN1, "             ' 測定値04 Den1"
'''''    sql = sql & "MS04DEN2, "             ' 測定値04 Den2"
'''''    sql = sql & "MS04DEN3, "             ' 測定値04 Den3"
'''''    sql = sql & "MS04DEN4, "             ' 測定値04 Den4"
'''''    sql = sql & "MS04DEN5, "             ' 測定値04 Den5"
'''''    sql = sql & "MS05LDL1, "             ' 測定値05 L/DL1"
'''''    sql = sql & "MS05LDL2, "             ' 測定値05 L/DL2"
'''''    sql = sql & "MS05LDL3, "             ' 測定値05 L/DL3"
'''''    sql = sql & "MS05LDL4, "             ' 測定値05 L/DL4"
'''''    sql = sql & "MS05LDL5, "             ' 測定値05 L/DL5"
'''''    sql = sql & "MS05DEN1, "             ' 測定値05 Den1"
'''''    sql = sql & "MS05DEN2, "             ' 測定値05 Den2"
'''''    sql = sql & "MS05DEN3, "             ' 測定値05 Den3"
'''''    sql = sql & "MS05DEN4, "             ' 測定値05 Den4"
'''''    sql = sql & "MS05DEN5, "             ' 測定値05 Den5"
'''''    sql = sql & "MS06LDL1, "             ' 測定値06 L/DL1"
'''''    sql = sql & "MS06LDL2, "             ' 測定値06 L/DL2"
'''''    sql = sql & "MS06LDL3, "             ' 測定値06 L/DL3"
'''''    sql = sql & "MS06LDL4, "             ' 測定値06 L/DL4"
'''''    sql = sql & "MS06LDL5, "             ' 測定値06 L/DL5"
'''''    sql = sql & "MS06DEN1, "             ' 測定値06 Den1"
'''''    sql = sql & "MS06DEN2, "             ' 測定値06 Den2"
'''''    sql = sql & "MS06DEN3, "             ' 測定値06 Den3"
'''''    sql = sql & "MS06DEN4, "             ' 測定値06 Den4"
'''''    sql = sql & "MS06DEN5, "             ' 測定値06 Den5"
'''''    sql = sql & "MS07LDL1, "             ' 測定値07 L/DL1"
'''''    sql = sql & "MS07LDL2, "             ' 測定値07 L/DL2"
'''''    sql = sql & "MS07LDL3, "             ' 測定値07 L/DL3"
'''''    sql = sql & "MS07LDL4, "             ' 測定値07 L/DL4"
'''''    sql = sql & "MS07LDL5, "             ' 測定値07 L/DL5"
'''''    sql = sql & "MS07DEN1, "             ' 測定値07 Den1"
'''''    sql = sql & "MS07DEN2, "             ' 測定値07 Den2"
'''''    sql = sql & "MS07DEN3, "             ' 測定値07 Den3"
'''''    sql = sql & "MS07DEN4, "             ' 測定値07 Den4"
'''''    sql = sql & "MS07DEN5, "             ' 測定値07 Den5"
'''''    sql = sql & "MS08LDL1, "             ' 測定値08 L/DL1"
'''''    sql = sql & "MS08LDL2, "             ' 測定値08 L/DL2"
'''''    sql = sql & "MS08LDL3, "             ' 測定値08 L/DL3"
'''''    sql = sql & "MS08LDL4, "             ' 測定値08 L/DL4"
'''''    sql = sql & "MS08LDL5, "             ' 測定値08 L/DL5"
'''''    sql = sql & "MS08DEN1, "             ' 測定値08 Den1"
'''''    sql = sql & "MS08DEN2, "             ' 測定値08 Den2"
'''''    sql = sql & "MS08DEN3, "             ' 測定値08 Den3"
'''''    sql = sql & "MS08DEN4, "             ' 測定値08 Den4"
'''''    sql = sql & "MS08DEN5, "             ' 測定値08 Den5"
'''''    sql = sql & "MS09LDL1, "             ' 測定値09 L/DL1"
'''''    sql = sql & "MS09LDL2, "             ' 測定値09 L/DL2"
'''''    sql = sql & "MS09LDL3, "             ' 測定値09 L/DL3"
'''''    sql = sql & "MS09LDL4, "             ' 測定値09 L/DL4"
'''''    sql = sql & "MS09LDL5, "             ' 測定値09 L/DL5"
'''''    sql = sql & "MS09DEN1, "             ' 測定値09 Den1"
'''''    sql = sql & "MS09DEN2, "             ' 測定値09 Den2"
'''''    sql = sql & "MS09DEN3, "             ' 測定値09 Den3"
'''''    sql = sql & "MS09DEN4, "             ' 測定値09 Den4"
'''''    sql = sql & "MS09DEN5, "             ' 測定値09 Den5"
'''''    sql = sql & "MS10LDL1, "             ' 測定値10 L/DL1"
'''''    sql = sql & "MS10LDL2, "             ' 測定値10 L/DL2"
'''''    sql = sql & "MS10LDL3, "             ' 測定値10 L/DL3"
'''''    sql = sql & "MS10LDL4, "             ' 測定値10 L/DL4"
'''''    sql = sql & "MS10LDL5, "             ' 測定値10 L/DL5"
'''''    sql = sql & "MS10DEN1, "             ' 測定値10 Den1"
'''''    sql = sql & "MS10DEN2, "             ' 測定値10 Den2"
'''''    sql = sql & "MS10DEN3, "             ' 測定値10 Den3"
'''''    sql = sql & "MS10DEN4, "             ' 測定値10 Den4"
'''''    sql = sql & "MS10DEN5, "             ' 測定値10 Den5"
'''''    sql = sql & "MS11LDL1, "             ' 測定値11 L/DL1"
'''''    sql = sql & "MS11LDL2, "             ' 測定値11 L/DL2"
'''''    sql = sql & "MS11LDL3, "             ' 測定値11 L/DL3"
'''''    sql = sql & "MS11LDL4, "             ' 測定値11 L/DL4"
'''''    sql = sql & "MS11LDL5, "             ' 測定値11 L/DL5"
'''''    sql = sql & "MS11DEN1, "             ' 測定値11 Den1"
'''''    sql = sql & "MS11DEN2, "             ' 測定値11 Den2"
'''''    sql = sql & "MS11DEN3, "             ' 測定値11 Den3"
'''''    sql = sql & "MS11DEN4, "             ' 測定値11 Den4"
'''''    sql = sql & "MS11DEN5, "             ' 測定値11 Den5"
'''''    sql = sql & "MS12LDL1, "             ' 測定値12 L/DL1"
'''''    sql = sql & "MS12LDL2, "             ' 測定値12 L/DL2"
'''''    sql = sql & "MS12LDL3, "             ' 測定値12 L/DL3"
'''''    sql = sql & "MS12LDL4, "             ' 測定値12 L/DL4"
'''''    sql = sql & "MS12LDL5, "             ' 測定値12 L/DL5"
'''''    sql = sql & "MS12DEN1, "             ' 測定値12 Den1"
'''''    sql = sql & "MS12DEN2, "             ' 測定値12 Den2"
'''''    sql = sql & "MS12DEN3, "             ' 測定値12 Den3"
'''''    sql = sql & "MS12DEN4, "             ' 測定値12 Den4"
'''''    sql = sql & "MS12DEN5, "             ' 測定値12 Den5"
'''''    sql = sql & "MS13LDL1, "             ' 測定値13 L/DL1"
'''''    sql = sql & "MS13LDL2, "             ' 測定値13 L/DL2"
'''''    sql = sql & "MS13LDL3, "             ' 測定値13 L/DL3"
'''''    sql = sql & "MS13LDL4, "             ' 測定値13 L/DL4"
'''''    sql = sql & "MS13LDL5, "             ' 測定値13 L/DL5"
'''''    sql = sql & "MS13DEN1, "             ' 測定値13 Den1"
'''''    sql = sql & "MS13DEN2, "             ' 測定値13 Den2"
'''''    sql = sql & "MS13DEN3, "             ' 測定値13 Den3"
'''''    sql = sql & "MS13DEN4, "             ' 測定値13 Den4"
'''''    sql = sql & "MS13DEN5, "             ' 測定値13 Den5"
'''''    sql = sql & "MS14LDL1, "             ' 測定値14 L/DL1"
'''''    sql = sql & "MS14LDL2, "             ' 測定値14 L/DL2"
'''''    sql = sql & "MS14LDL3, "             ' 測定値14 L/DL3"
'''''    sql = sql & "MS14LDL4, "             ' 測定値14 L/DL4"
'''''    sql = sql & "MS14LDL5, "             ' 測定値14 L/DL5"
'''''    sql = sql & "MS14DEN1, "             ' 測定値14 Den1"
'''''    sql = sql & "MS14DEN2, "             ' 測定値14 Den2"
'''''    sql = sql & "MS14DEN3, "             ' 測定値14 Den3"
'''''    sql = sql & "MS14DEN4, "             ' 測定値14 Den4"
'''''    sql = sql & "MS14DEN5, "             ' 測定値14 Den5"
'''''    sql = sql & "MS15LDL1, "             ' 測定値15 L/DL1"
'''''    sql = sql & "MS15LDL2, "             ' 測定値15 L/DL2"
'''''    sql = sql & "MS15LDL3, "             ' 測定値15 L/DL3"
'''''    sql = sql & "MS15LDL4, "             ' 測定値15 L/DL4"
'''''    sql = sql & "MS15LDL5, "             ' 測定値15 L/DL5"
'''''    sql = sql & "MS15DEN1, "             ' 測定値15 Den1"
'''''    sql = sql & "MS15DEN2, "             ' 測定値15 Den2"
'''''    sql = sql & "MS15DEN3, "             ' 測定値15 Den3"
'''''    sql = sql & "MS15DEN4, "             ' 測定値15 Den4"
'''''    sql = sql & "MS15DEN5, "              ' 測定値15 Den5"
'''''    sql = sql & "REGDATE "               ' 登録日付
'''''
'''''End Sub


''''''内部関数 GD実績取得用
'''''Private Function GD_Zisseki(Siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                              Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
'''''                              GD As type_DBDRV_scmzc_fcmkc001c_GD, _
'''''                              TorB As Integer, _
'''''                              Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
'''''    Dim sql As String
'''''    Dim rs As OraDynaset
'''''    Dim recCnt As Integer
'''''    Dim i As Long
'''''    Dim Tname As String
'''''    Dim NothingFlag As Boolean
'''''
'''''    NothingFlag = False
'''''
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function GD_Zisseki"
'''''
'''''    GD_Zisseki = FUNCTION_RETURN_SUCCESS
'''''
'''''    Tname = "TBCMJ006"
'''''    Set rs = Nothing
'''''
'''''    DoEvents
'''''    Call GD_SetBaseSQL(sql)
'''''    DoEvents
'''''    Call AddSQL_Default2(sql, Samp.CRYNUM, Samp.INGOTPOS, Samp.SMPKBN, Tname)
'''''    DoEvents
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    DoEvents
'''''
'''''    If rs.RecordCount <> 0 Then
'''''        DoEvents
'''''        Call GD_ObjCpy(GD, rs)
'''''        rs.Close
'''''        Set rs = Nothing
'''''    Else
'''''        rs.Close
'''''        Set rs = Nothing
'''''        If Siyou.HSXDENKU = "1" Or Siyou.HSXDVDKU = "1" Or Siyou.HSXLDLKU = "1" Then
'''''           If Samp.CRYINDGD = "6" Then       ' 引継ぎなら
'''''                DoEvents
'''''                Call GD_SetBaseSQL(sql)
'''''                DoEvents
'''''                Call AddSQL_HIKITUGI2(sql, Samp.CRYNUM, Samp.INGOTPOS, Tname, TorB)
'''''
'''''                DoEvents
'''''                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                DoEvents
'''''
'''''                If rs.RecordCount <> 0 Then
'''''                    recCnt = rs.RecordCount
'''''                    For i = 1 To recCnt
'''''                        If i = 1 Then                                     ' 一回目は保持
'''''                            DoEvents
'''''                            Call GD_ObjCpy(GD, rs)
'''''                            Exit For    '１レコード目だけでOK
'''''                        Else
'''''                            If GD.POSITION = rs("POSITION") And GD.REGDATE < rs("REGDATE") Then   ' 前の位置と同じだったら登録日付が新しいものをとる
'''''                                DoEvents
'''''                                Call GD_ObjCpy(GD, rs)
'''''                            End If
'''''                        End If
'''''
'''''                        rs.MoveNext
'''''                    Next
'''''                Else
'''''                    NothingFlag = True
'''''                End If
'''''                If Not rs Is Nothing Then
'''''                    rs.Close
'''''                    Set rs = Nothing
'''''                End If
'''''            End If  ' 検査指示が5 or 6 なら
'''''        End If  ' 指示がたっている
'''''    End If ' 実績がある
'''''
'''''    If NothingFlagStr <> vbNullString Then
'''''        If NothingFlag Then
'''''            NothingFlagStr = "1"
'''''        End If
'''''    End If
'''''
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    GD_Zisseki = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function



'''''Private Sub LT_ObjCpy(Lt As type_DBDRV_scmzc_fcmkc001c_LT, rs As OraDynaset)
'''''    With Lt
'''''        .CRYNUM = rs("CRYNUM")         ' 結晶番号
'''''        .POSITION = rs("POSITION")     ' 位置
'''''        .SMPKBN = rs("SMPKBN")         ' サンプル区分
'''''        .SMPLNO = rs("SMPLNO")         ' サンプルＮｏ
'''''        .SMPLUMU = rs("SMPLUMU")       ' サンプル有無
'''''        .MEAS1 = rs("MEAS1")           ' 測定値１
'''''        .MEAS2 = rs("MEAS2")           ' 測定値２
'''''        .MEAS3 = rs("MEAS3")           ' 測定値３
'''''        .MEAS4 = rs("MEAS4")           ' 測定値４
'''''        .MEAS5 = rs("MEAS5")           ' 測定値５
'''''        .TRANCOND = rs("TRANCOND")     ' 処理条件
'''''        .MEASPEAK = rs("MEASPEAK")     ' 測定値 ピーク値
'''''        .CALCMEAS = rs("CALCMEAS")     ' 計算結果
'''''        .REGDATE = rs("REGDATE")        '　登録日付
'''''        .LTSPI = rs("HSXLTSPI")            '測定位置コード
'''''    End With
'''''
'''''End Sub
'''''
'''''Private Sub LT_SetBaseSQL(sql As String)
'''''    sql = "select "
'''''    sql = sql & "CRYNUM, "        ' 結晶番号
'''''    sql = sql & "POSITION, "      ' 位置
'''''    sql = sql & "SMPKBN, "        ' サンプル区分
'''''    sql = sql & "SMPLNO, "        ' サンプルＮｏ
'''''    sql = sql & "SMPLUMU, "       ' サンプル有無
'''''    sql = sql & "MEAS1,"          ' 測定値１"
'''''    sql = sql & "MEAS2,"          ' 測定値２"
'''''    sql = sql & "MEAS3,"          ' 測定値３"
'''''    sql = sql & "MEAS4,"          ' 測定値４"
'''''    sql = sql & "MEAS5,"          ' 測定値５"
'''''    sql = sql & "TRANCOND, "      ' 処理条件
'''''    sql = sql & "MEASPEAK, "      ' 測定値 ピーク値
'''''    sql = sql & "CALCMEAS, "       ' 計算結果
'''''    sql = sql & "REGDATE "        ' 登録日付
'''''
'''''End Sub


''''''内部関数 ライフタイム実績取得用
'''''Private Function LT_Zisseki(Siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                              Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
'''''                              Lt As type_DBDRV_scmzc_fcmkc001c_LT, _
'''''                              TorB As Integer, _
'''''                              Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
'''''    Dim sql As String
'''''    Dim rs As OraDynaset
'''''    Dim recCnt As Integer
'''''    Dim i As Long
'''''    Dim Tname As String
'''''    Dim NothingFlag As Boolean
'''''    Dim hin As tFullHinban
'''''    Dim LTSPI As String
'''''
'''''    NothingFlag = False
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function LT_Zisseki"
'''''
'''''    ' ライフタイム実績テーブルから値を取得
'''''    LT_Zisseki = FUNCTION_RETURN_SUCCESS
'''''
'''''    Tname = "TBCMJ007"
'''''
'''''    If NothingFlagStr <> vbNullString Then NothingFlagStr = "0"
'''''    With Lt
'''''        .CRYNUM = vbNullString
'''''        .POSITION = -1
'''''        .SMPLNO = -1
'''''    End With
'''''
'''''    'サンプル位置にLT実績があれば、それを (B側優先、最後の実績)
'''''    sql = "select * from ("
'''''    sql = sql & "  select CRYNUM, POSITION, SMPKBN, TRANCOND, SMPLNO, SMPLUMU"
'''''    sql = sql & "  , MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASPEAK, CALCMEAS, REGDATE"
'''''    sql = sql & "  , (select HSXLTSPI from TBCME019 where HINBAN=LT.HINBAN and MNOREVNO=LT.REVNUM and FACTORY=LT.FACTORY and OPECOND=LT.OPECOND) as HSXLTSPI"
'''''    sql = sql & "  from TBCMJ007 LT"
'''''    sql = sql & "  where CRYNUM='" & Samp.CRYNUM & "' and POSITION=" & Samp.INGOTPOS
'''''    sql = sql & "  order by SMPKBN, TRANCNT desc"
'''''    sql = sql & ") where rownum=1"
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    If rs.RecordCount > 0 Then
'''''        Call LT_ObjCpy(Lt, rs)
'''''        rs.Close
'''''        Set rs = Nothing
'''''        LT_Zisseki = FUNCTION_RETURN_SUCCESS
'''''        GoTo proc_exit
'''''    End If
'''''
'''''    'ブロック内で最も厳しいLT仕様を持つ品番と、その測定位置を求める
'''''    If DBDRV_getLtHinbanInBlock(Samp.CRYNUM, Samp.INGOTPOS, hin, LTSPI) = FUNCTION_RETURN_FAILURE Then
'''''        LT_Zisseki = FUNCTION_RETURN_FAILURE
'''''        GoTo proc_exit
'''''    End If
'''''    If LTSPI = vbNullString Then    'ブロック内にLT仕様なし
'''''        If NothingFlagStr <> vbNullString Then NothingFlagStr = "1"
'''''        LT_Zisseki = FUNCTION_RETURN_SUCCESS
'''''        GoTo proc_exit
'''''    End If
'''''
'''''    'LT測定結果を検索する (同一測定位置があればそれ、なければより厳しい検査結果を、なるべく近い下側から取る)
'''''    sql = "select * from ("
'''''    sql = sql & "  select LT.CRYNUM, POSITION, TRANCOND, SMPKBN, SMPLNO, SMPLUMU"
'''''    sql = sql & "  , MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASPEAK, CALCMEAS, LT.REGDATE"
'''''    sql = sql & "  , SIYO.HSXLTSPI"
'''''    sql = sql & "  from TBCMJ007 LT, TBCME019 SIYO"
'''''    sql = sql & "  where LT.HINBAN=SIYO.HINBAN and LT.REVNUM=SIYO.MNOREVNO and LT.FACTORY=SIYO.FACTORY and LT.OPECOND=SIYO.OPECOND"
'''''    sql = sql & "    and LT.CRYNUM='" & Samp.CRYNUM & "'"
'''''    sql = sql & "    and POSITION>=" & Samp.INGOTPOS
'''''    sql = sql & "    and decode(SIYO.HSXLTSPI,' ','ZZ',SIYO.HSXLTSPI)<='" & LTSPI & "'"
'''''    sql = sql & "  order by decode(HSXLTSPI,'" & LTSPI & "',1,0) desc, POSITION, SMPKBN,TRANCNT desc"
'''''    sql = sql & ") where rownum=1"
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    If rs.RecordCount > 0 Then
'''''        Call LT_ObjCpy(Lt, rs)
'''''    Else
'''''        If NothingFlagStr <> vbNullString Then NothingFlagStr = "1"
'''''    End If
'''''    rs.Close
'''''    Set rs = Nothing
'''''
'''''    LT_Zisseki = FUNCTION_RETURN_SUCCESS
'''''
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    LT_Zisseki = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


''''''概要      :
'''''Private Sub EPD_ObjCpy(EPD As type_DBDRV_scmzc_fcmkc001c_EPD, rs As OraDynaset)
'''''    With EPD
'''''        .CRYNUM = rs("CRYNUM")      ' 結晶番号
'''''        .POSITION = rs("POSITION")  ' 位置
'''''        .SMPKBN = rs("SMPKBN")      ' サンプル区分
'''''        .SMPLNO = rs("SMPLNO")      ' サンプルＮｏ
'''''        .SMPLUMU = rs("SMPLUMU")    ' サンプル有無
'''''        .TRANCOND = rs("TRANCOND")  ' 処理条件
'''''        .MEASURE = rs("MEASURE")    ' 測定値
'''''        .REGDATE = rs("REGDATE")    ' 登録日付
'''''    End With
'''''End Sub
'''''
''''''概要      :
'''''Private Sub EPD_SetBaseSQL(sql As String)
'''''    sql = "select "
'''''    sql = sql & "CRYNUM, "          ' 結晶番号
'''''    sql = sql & "POSITION, "        ' 位置
'''''    sql = sql & "SMPKBN, "          ' サンプル区分
'''''    sql = sql & "SMPLNO, "          ' サンプルＮｏ
'''''    sql = sql & "SMPLUMU, "         ' サンプル有無
'''''    sql = sql & "TRANCOND, "        ' 処理条件
'''''    sql = sql & "MEASURE, "         ' 測定値
'''''    sql = sql & "REGDATE "          ' 登録日付
'''''End Sub


''''''概要      :内部関数 EPD実績取得用
'''''Private Function EPD_Zisseki(Siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                              Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
'''''                              EPD As type_DBDRV_scmzc_fcmkc001c_EPD, _
'''''                              TorB As Integer, _
'''''                              Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
'''''    Dim sql As String
'''''    Dim rs As OraDynaset
'''''    Dim recCnt As Integer
'''''    Dim i As Long
'''''    Dim Tname As String
'''''    Dim NothingFlag As Boolean
'''''
'''''    NothingFlag = False
'''''
'''''    ' EPD実績テーブルから値を取得
'''''    ' VECME010（ブロック管理を検索し、そのブロックに対するサンプルを表示するビュー）からサンプル区分を取得（where 結晶番号、位置）
'''''    ' 結晶番号、位置、サンプル区分、処理回数最大を検索条件とし実績テーブルから値を取得
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function EPD_Zisseki"
'''''
'''''    EPD_Zisseki = FUNCTION_RETURN_SUCCESS
'''''
'''''    Tname = "TBCMJ001"
'''''    Set rs = Nothing
'''''
'''''    DoEvents
'''''    Call EPD_SetBaseSQL(sql)
'''''    DoEvents
'''''    Call AddSQL_Default2(sql, Samp.CRYNUM, Samp.INGOTPOS, Samp.SMPKBN, Tname)
'''''
'''''    DoEvents
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    DoEvents
'''''
'''''    If rs.RecordCount <> 0 Then
'''''        DoEvents
'''''        Call EPD_ObjCpy(EPD, rs)
'''''        rs.Close
'''''        Set rs = Nothing
'''''    Else
'''''        rs.Close
'''''        Set rs = Nothing
'''''
'''''
'''''        ' 下方向に実績を探し、同じ位置にあった場合　新しい日付の方を取得する
'''''        DoEvents
'''''        Call EPD_SetBaseSQL(sql)
'''''        DoEvents
'''''        Call AddSQL_Down(sql, Samp.CRYNUM, Samp.INGOTPOS, Tname)
'''''
'''''        DoEvents
'''''        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''        DoEvents
'''''
'''''        If rs.RecordCount <> 0 Then
'''''            recCnt = rs.RecordCount
'''''            For i = 1 To recCnt
'''''                If i = 1 Then                                     ' 一回目は保持
'''''                    DoEvents
'''''                    Call EPD_ObjCpy(EPD, rs)
'''''                Else
'''''                    If EPD.POSITION = rs("POSITION") And EPD.REGDATE < rs("REGDATE") Then   ' 前の位置と同じだったら登録日付が新しいものをとる
'''''                        DoEvents
'''''                        Call EPD_ObjCpy(EPD, rs)
'''''                    End If
'''''                End If
'''''
'''''                rs.MoveNext
'''''            Next
'''''        Else
'''''            NothingFlag = True
'''''        End If
'''''        If Not rs Is Nothing Then
'''''            rs.Close
'''''            Set rs = Nothing
'''''        End If
'''''    End If
'''''
'''''    If NothingFlagStr <> vbNullString Then
'''''        If NothingFlag Then
'''''            NothingFlagStr = "1"
'''''        End If
'''''    End If
'''''
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    If Not rs Is Nothing Then
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''    EPD_Zisseki = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


''''''概要      :内部関数 テーブル「TBCMG002」から条件にあったレコードを抽出する
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
''''''          :records       ,O  ,typ_TBCMG002 ,抽出レコード
''''''          :[sqlWhere]    ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
''''''          :[sqlOrder]    ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
''''''          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
''''''説明      :
''''''履歴      :
'''''Private Function Kounyu_Zisseki(records As typ_TBCMG002, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
'''''Dim sql As String       'SQL全体
'''''Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
'''''Dim rs As OraDynaset    'RecordSet
'''''Dim recCnt As Long      'レコード数
'''''Dim i As Long
'''''
'''''    ''SQLを組み立てる
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function Kounyu_Zisseki"
'''''
'''''    Kounyu_Zisseki = FUNCTION_RETURN_SUCCESS
'''''
'''''    sqlBase = "Select CRYNUM, TRANCNT, KRPROCCD, PROCCODE, HINBAN, MNOREVNO, FACTORY, OPECOND, REPCCL, RBATCHNO, DMTOP1, DMTOP2," & _
'''''              " DMTAIL1, DMTAIL2, NCHDPTH1, NCHDPTH2, UPLENGTH, SXLPOS, BLKLEN, BLKWGHT, CMPTOP1, CMPTOP2, CMPTOP3, CMPTOP4," & _
'''''              " CMPTOP5, CMPTOPR, CMPTAIL1, CMPTAIL2, CMPTAIL3, CMPTAIL4, CMPTAIL5, CMPTAILR, OITOP1, OITOP2, OITOP3, OITOP4," & _
'''''              " OITOP5, OITOPR, OITAIL1, OITAIL2, OITAIL3, OITAIL4, OITAIL5, OITAILR, CSTOP, CSTAIL, LD1TOPMX, LD1TOPAV, LD1TAILM," & _
'''''              " LD1TAILA, LD2TOPMM, LD2TOPAV, LD2TAILM, LD2TAILA, BMDTOPMX, BMDTOPAV, BMDTAILM, BMDTAILA, GD1TOP, GD1TAIL," & _
'''''              " GD2TOP, GD2TAIL, DIA1TOP, DIA1TAIL, DIA2TOP, DIA2TAIL, LTFTOP, LTFTAIL, EPD, TSTAFFID, REGDATE, KSTAFFID," & _
'''''              " UPDDATE, SENDFLAG, SENDDATE "
'''''    sqlBase = sqlBase & "From TBCMG002"
'''''    sql = sqlBase
'''''    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
'''''        sql = sql & sqlWhere & sqlOrder
'''''    End If
'''''
'''''    ''データを抽出する
'''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''    If rs.RecordCount = 0 Then
'''''        GoTo proc_exit
'''''    End If
'''''
'''''    ''抽出結果を格納する
'''''    With records
'''''        .CRYNUM = rs("CRYNUM")           ' 結晶番号
'''''        .TRANCNT = rs("TRANCNT")         ' 処理回数
'''''        .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
'''''        .PROCCODE = rs("PROCCODE")       ' 工程コード
'''''        .hinban = rs("HINBAN")           ' 品番
'''''        .mnorevno = rs("MNOREVNO")       ' 製品番号改訂番号
'''''        .factory = rs("FACTORY")         ' 工場
'''''        .opecond = rs("OPECOND")         ' 操業条件
'''''        .REPCCL = rs("REPCCL")           ' 受入取消区分
'''''        .RBATCHNO = rs("RBATCHNO")       ' 炉バッチＮｏ
'''''        .DMTOP1 = rs("DMTOP1")           ' 直径ＴＯＰ１
'''''        .DMTOP2 = rs("DMTOP2")           ' 直径ＴＯＰ２
'''''        .DMTAIL1 = rs("DMTAIL1")         ' 直径ＴＡＩＬ１
'''''        .DMTAIL2 = rs("DMTAIL2")         ' 直径ＴＡＩＬ２
'''''        .NCHDPTH1 = rs("NCHDPTH1")       ' ノッチ深さ１
'''''        .NCHDPTH2 = rs("NCHDPTH2")       ' ノッチ深さ２
'''''        .UPLENGTH = rs("UPLENGTH")       ' 引上げ長
'''''        .SXLPOS = rs("SXLPOS")           ' ＳＸＬ位置
'''''        .BlkLen = rs("BLKLEN")           ' ブロック長さ
'''''        .BLKWGHT = rs("BLKWGHT")         ' ブロック重量
'''''        .CMPTOP1 = rs("CMPTOP1")         ' 比抵抗TOP　１
'''''        .CMPTOP2 = rs("CMPTOP2")         ' 比抵抗TOP　２
'''''        .CMPTOP3 = rs("CMPTOP3")         ' 比抵抗TOP　３
'''''        .CMPTOP4 = rs("CMPTOP4")         ' 比抵抗TOP　４
'''''        .CMPTOP5 = rs("CMPTOP5")         ' 比抵抗TOP　５
'''''        .CMPTOPR = rs("CMPTOPR")         ' 比抵抗TOP　RRG
'''''        .CMPTAIL1 = rs("CMPTAIL1")       ' 比抵抗TAIL　１
'''''        .CMPTAIL2 = rs("CMPTAIL2")       ' 比抵抗TAIL　２
'''''        .CMPTAIL3 = rs("CMPTAIL3")       ' 比抵抗TAIL　３
'''''        .CMPTAIL4 = rs("CMPTAIL4")       ' 比抵抗TAIL　４
'''''        .CMPTAIL5 = rs("CMPTAIL5")       ' 比抵抗TAIL　５
'''''        .CMPTAILR = rs("CMPTAILR")       ' 比抵抗TAIL　RRG
'''''        .OITOP1 = rs("OITOP1")           ' Oi　TOP　１
'''''        .OITOP2 = rs("OITOP2")           ' Oi　TOP　２
'''''        .OITOP3 = rs("OITOP3")           ' Oi　TOP　３
'''''        .OITOP4 = rs("OITOP4")           ' Oi　TOP　４
'''''        .OITOP5 = rs("OITOP5")           ' Oi　TOP　５
'''''        .OITOPR = rs("OITOPR")           ' Oi　TOP　ROG
'''''        .OITAIL1 = rs("OITAIL1")         ' Oi　TAIL　１
'''''        .OITAIL2 = rs("OITAIL2")         ' Oi　TAIL　２
'''''        .OITAIL3 = rs("OITAIL3")         ' Oi　TAIL　３
'''''        .OITAIL4 = rs("OITAIL4")         ' Oi　TAIL　４
'''''        .OITAIL5 = rs("OITAIL5")         ' Oi　TAIL　５
'''''        .OITAILR = rs("OITAILR")         ' Oi　TAIL　ROG
'''''        .CSTOP = rs("CSTOP")             ' Cs　TOP
'''''        .CSTAIL = rs("CSTAIL")           ' Cs　TAIL
'''''        .LD1TOPMX = rs("LD1TOPMX")       ' LD-1　TOP　MAX
'''''        .LD1TOPAV = rs("LD1TOPAV")       ' LD-1　TOP　AVE
'''''        .LD1TAILM = rs("LD1TAILM")       ' LD-1　TAIL　MAX
'''''        .LD1TAILA = rs("LD1TAILA")       ' LD-1　TAIL　AVE
'''''        .LD2TOPMM = rs("LD2TOPMM")       ' LD-2　TOP　MAX
'''''        .LD2TOPAV = rs("LD2TOPAV")       ' LD-2　TOP　AVE
'''''        .LD2TAILM = rs("LD2TAILM")       ' LD-2　TAIL　MAX
'''''        .LD2TAILA = rs("LD2TAILA")       ' LD-2　TAIL　AVE
'''''        .BMDTOPMX = rs("BMDTOPMX")       ' BMD　TOP　MAX
'''''        .BMDTOPAV = rs("BMDTOPAV")       ' BMD　TOP　AVE
'''''        .BMDTAILM = rs("BMDTAILM")       ' BMD　TAIL　MAX
'''''        .BMDTAILA = rs("BMDTAILA")       ' BMD　TAIL　AVE
'''''        .GD1TOP = rs("GD1TOP")           ' GD1 TOP
'''''        .GD1TAIL = rs("GD1TAIL")         ' GD1 TAIL
'''''        .GD2TOP = rs("GD2TOP")           ' GD2 TOP
'''''        .GD2TAIL = rs("GD2TAIL")         ' GD2 TAIL
'''''        .DIA1TOP = rs("DIA1TOP")         ' DIA1 TOP
'''''        .DIA1TAIL = rs("DIA1TAIL")       ' DIA1 TAIL
'''''        .DIA2TOP = rs("DIA2TOP")         ' DIA2 TOP
'''''        .DIA2TAIL = rs("DIA2TAIL")       ' DIA2 TAIL
'''''        .LTFTOP = rs("LTFTOP")           ' LIFETIME from TOP
'''''        .LTFTAIL = rs("LTFTAIL")         ' LIFETIME from TAIL
'''''        .EPD = rs("EPD")                 ' EPD
'''''        .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
'''''        .REGDATE = rs("REGDATE")         ' 登録日付
'''''        .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
'''''        .UPDDATE = rs("UPDDATE")         ' 更新日付
'''''        .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
'''''        .SENDDATE = rs("SENDDATE")       ' 送信日付
'''''    End With
'''''    rs.Close
'''''
'''''
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    Kounyu_Zisseki = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function



'概要      :内部関数 品番、仕様を取得する
Public Function getHinSiyou30(inBlockID As String, Siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou) As FUNCTION_RETURN

    Dim sql         As String
    Dim rs          As OraDynaset
    Dim recCnt      As Integer
    Dim i           As Long
    Dim Jiltuseki   As Judg_Kakou

    '品番、SXL仕様からデータの取得
' 払出規制項目追加対応 yakimura 2002.12.01 start

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function getHinSiyou"

    getHinSiyou30 = FUNCTION_RETURN_SUCCESS

    sql = "select "
    sql = sql & "distinct "                   ''重複データ削除  2003/09/09 ooba
    sql = sql & "BH.E040CRYNUM, "             ' 結晶番号
    sql = sql & "BH.E040INGOTPOS, "           ' 結晶内開始位置
    sql = sql & "BH.E040LENGTH, "             ' 長さ
    sql = sql & "BH.E041HINBAN, "             ' 品番
    sql = sql & "BH.E041REVNUM, "             ' 製品番号改訂番号
    sql = sql & "BH.E041FACTORY, "            ' 工場
    sql = sql & "BH.E041OPECOND, "            ' 操業条件

    sql = sql & "BH.E037PRODCOND, "           ' 製作条件
    sql = sql & "BH.E037PGID, "               ' ＰＧ－ＩＤ
    sql = sql & "BH.E037UPLENGTH, "           ' 引上げ長さ
    sql = sql & "BH.E037FREELENG, "           ' フリー長
    sql = sql & "BH.E037DIAMETER, "           ' 直径
    sql = sql & "BH.E037CHARGE, "             ' チャージ量
    sql = sql & "BH.E037SEED, "               ' シード
    sql = sql & "BH.E037ADDDPPOS, "           ' 追加ドープ位置

    sql = sql & "S.E018HSXTYPE, "             ' 品ＳＸタイプ
    sql = sql & "S.E018HSXD1CEN, "            ' 品ＳＸ直径１中心
    sql = sql & "S.E018HSXCDIR, "             ' 品ＳＸ結晶面方位

    sql = sql & "S.E018HSXRMIN, "             ' 品ＳＸ比抵抗下限
    sql = sql & "S.E018HSXRMAX, "             ' 品ＳＸ比抵抗上限
    sql = sql & "S.E018HSXRAMIN, "            ' 品ＳＸ比抵抗平均下限
    sql = sql & "S.E018HSXRAMAX, "            ' 品ＳＸ比抵抗平均上限
    sql = sql & "S.E018HSXRMCAL, "            ' 品ＳＸ比抵抗面内計算　　'' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06
    sql = sql & "S.E018HSXRMBNP, "            ' 品ＳＸ比抵抗面内分布
    sql = sql & "S.E018HSXRSPOH, "            ' 品ＳＸ比抵抗測定位置＿方
    sql = sql & "S.E018HSXRSPOT, "            ' 品ＳＸ比抵抗測定位置＿点
    sql = sql & "S.E018HSXRSPOI, "            ' 品ＳＸ比抵抗測定位置＿位
    sql = sql & "S.E018HSXRHWYT, "            ' 品ＳＸ比抵抗保証方法＿対
    sql = sql & "S.E018HSXRHWYS, "            ' 品ＳＸ比抵抗保証方法＿処

    sql = sql & "S.E019HSXONMIN, "            ' 品ＳＸ酸素濃度下限
    sql = sql & "S.E019HSXONMAX, "            ' 品ＳＸ酸素濃度上限
    sql = sql & "S.E019HSXONAMN, "            ' 品ＳＸ酸素濃度平均下限
    sql = sql & "S.E019HSXONAMX, "            ' 品ＳＸ酸素濃度平均上限
    sql = sql & "S.E019HSXONMCL, "            ' 品ＳＸ酸素濃度面内計算　　'' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06
    sql = sql & "S.E019HSXONMBP, "            ' 品ＳＸ酸素濃度面内分布
    sql = sql & "S.E019HSXONSPH, "            ' 品ＳＸ酸素濃度測定位置＿方
    sql = sql & "S.E019HSXONSPT, "            ' 品ＳＸ酸素濃度測定位置＿点
    sql = sql & "S.E019HSXONSPI, "            ' 品ＳＸ酸素濃度測定位置＿位
    sql = sql & "S.E019HSXONHWT, "            ' 品ＳＸ酸素濃度保証方法＿対
    sql = sql & "S.E019HSXONHWS, "            ' 品ＳＸ酸素濃度保証方法＿処

    sql = sql & "S.E020HSXBM1AN, "            ' 品ＳＸＢＭＤ１平均下限
    sql = sql & "S.E020HSXBM1AX, "            ' 品ＳＸＢＭＤ１平均上限
    sql = sql & "S.E020HSXBM2AN, "            ' 品ＳＸＢＭＤ２平均下限
    sql = sql & "S.E020HSXBM2AX, "            ' 品ＳＸＢＭＤ２平均上限
    sql = sql & "S.E020HSXBM3AN, "            ' 品ＳＸＢＭＤ３平均下限
    sql = sql & "S.E020HSXBM3AX, "            ' 品ＳＸＢＭＤ３平均上限
    sql = sql & "S.E020HSXBM1SH, "            ' 品ＳＸＢＭＤ１測定位置＿方
    sql = sql & "S.E020HSXBM1ST, "            ' 品ＳＸＢＭＤ１測定位置＿点
    sql = sql & "S.E020HSXBM1SR, "            ' 品ＳＸＢＭＤ１測定位置＿領
    sql = sql & "S.E020HSXBM1HT, "            ' 品ＳＸＢＭＤ１保証方法＿対
    sql = sql & "S.E020HSXBM1HS, "            ' 品ＳＸＢＭＤ１保証方法＿処
    sql = sql & "S.E020HSXBM2SH, "            ' 品ＳＸＢＭＤ２測定位置＿方
    sql = sql & "S.E020HSXBM2ST, "            ' 品ＳＸＢＭＤ２測定位置＿点
    sql = sql & "S.E020HSXBM2SR, "            ' 品ＳＸＢＭＤ２測定位置＿領
    sql = sql & "S.E020HSXBM2HT, "            ' 品ＳＸＢＭＤ２保証方法＿対
    sql = sql & "S.E020HSXBM2HS, "            ' 品ＳＸＢＭＤ２保証方法＿処
    sql = sql & "S.E020HSXBM3SH, "            ' 品ＳＸＢＭＤ３測定位置＿方
    sql = sql & "S.E020HSXBM3ST, "            ' 品ＳＸＢＭＤ３測定位置＿点
    sql = sql & "S.E020HSXBM3SR, "            ' 品ＳＸＢＭＤ３測定位置＿領
    sql = sql & "S.E020HSXBM3HT, "            ' 品ＳＸＢＭＤ３保証方法＿対
    sql = sql & "S.E020HSXBM3HS, "            ' 品ＳＸＢＭＤ３保証方法＿処

    sql = sql & "S.E020HSXOF1AX, "            ' 品ＳＸＯＳＦ１平均上限
    sql = sql & "S.E020HSXOF1MX, "            ' 品ＳＸＯＳＦ１上限
    sql = sql & "S.E020HSXOF2AX, "            ' 品ＳＸＯＳＦ２平均上限
    sql = sql & "S.E020HSXOF2MX, "            ' 品ＳＸＯＳＦ２上限
    sql = sql & "S.E020HSXOF3AX, "            ' 品ＳＸＯＳＦ３平均上限
    sql = sql & "S.E020HSXOF3MX, "            ' 品ＳＸＯＳＦ３上限
    sql = sql & "S.E020HSXOF4AX, "            ' 品ＳＸＯＳＦ４平均上限
    sql = sql & "S.E020HSXOF4MX, "            ' 品ＳＸＯＳＦ４上限
    sql = sql & "S.E020HSXOF1SH, "            ' 品ＳＸＯＳＦ１測定位置＿方
    sql = sql & "S.E020HSXOF1ST, "            ' 品ＳＸＯＳＦ１測定位置＿点
    sql = sql & "S.E020HSXOF1SR, "            ' 品ＳＸＯＳＦ１測定位置＿領
    sql = sql & "S.E020HSXOF1HT, "            ' 品ＳＸＯＳＦ１保証方法＿対
    sql = sql & "S.E020HSXOF1HS, "            ' 品ＳＸＯＳＦ１保証方法＿処
    sql = sql & "S.E020HSXOF2SH, "            ' 品ＳＸＯＳＦ２測定位置＿方
    sql = sql & "S.E020HSXOF2ST, "            ' 品ＳＸＯＳＦ２測定位置＿点
    sql = sql & "S.E020HSXOF2SR, "            ' 品ＳＸＯＳＦ２測定位置＿領
    sql = sql & "S.E020HSXOF2HT, "            ' 品ＳＸＯＳＦ２保証方法＿対
    sql = sql & "S.E020HSXOF2HS, "            ' 品ＳＸＯＳＦ２保証方法＿処
    sql = sql & "S.E020HSXOF3SH, "            ' 品ＳＸＯＳＦ３測定位置＿方
    sql = sql & "S.E020HSXOF3ST, "            ' 品ＳＸＯＳＦ３測定位置＿点
    sql = sql & "S.E020HSXOF3SR, "            ' 品ＳＸＯＳＦ３測定位置＿領
    sql = sql & "S.E020HSXOF3HT, "            ' 品ＳＸＯＳＦ３保証方法＿対
    sql = sql & "S.E020HSXOF3HS, "            ' 品ＳＸＯＳＦ３保証方法＿処
    sql = sql & "S.E020HSXOF4SH, "            ' 品ＳＸＯＳＦ４測定位置＿方
    sql = sql & "S.E020HSXOF4ST, "            ' 品ＳＸＯＳＦ４測定位置＿点
    sql = sql & "S.E020HSXOF4SR, "            ' 品ＳＸＯＳＦ４測定位置＿領
    sql = sql & "S.E020HSXOF4HT, "            ' 品ＳＸＯＳＦ４保証方法＿対
    sql = sql & "S.E020HSXOF4HS, "            ' 品ＳＸＯＳＦ４保証方法＿処
    sql = sql & "S.E020HSXOF1NS, "            ' 品ＳＸＯＳＦ１熱処理法
    sql = sql & "S.E020HSXOF2NS, "            ' 品ＳＸＯＳＦ２熱処理法
    sql = sql & "S.E020HSXOF3NS, "            ' 品ＳＸＯＳＦ３熱処理法
    sql = sql & "S.E020HSXOF4NS, "            ' 品ＳＸＯＳＦ４熱処理法
    sql = sql & "S.E020HSXBM1NS, "            ' 品ＳＸＢＭＤ１熱処理法
    sql = sql & "S.E020HSXBM2NS, "            ' 品ＳＸＢＭＤ２熱処理法
    sql = sql & "S.E020HSXBM3NS, "            ' 品ＳＸＢＭＤ３熱処理法

    sql = sql & "S.E019HSXCNMIN, "            ' 品ＳＸ炭素濃度下限
    sql = sql & "S.E019HSXCNMAX, "            ' 品ＳＸ炭素濃度上限
    sql = sql & "S.E019HSXCNSPH, "            ' 品ＳＸ炭素濃度測定位置＿方
    sql = sql & "S.E019HSXCNSPT, "            ' 品ＳＸ炭素濃度測定位置＿点
    sql = sql & "S.E019HSXCNSPI, "            ' 品ＳＸ炭素濃度測定位置＿位
    sql = sql & "S.E019HSXCNHWT, "            ' 品ＳＸ炭素濃度保証方法＿対
    sql = sql & "S.E019HSXCNHWS, "            ' 品ＳＸ炭素濃度保証方法＿処

    sql = sql & "S.E020HSXDENMX, "            ' 品ＳＸＤｅｎ上限
    sql = sql & "S.E020HSXDENMN, "            ' 品ＳＸＤｅｎ下限
    sql = sql & "S.E020HSXLDLMX, "            ' 品ＳＸＬ／ＤＬ上限
    sql = sql & "S.E020HSXLDLMN, "            ' 品ＳＸＬ／ＤＬ下限
    sql = sql & "S.E020HSXDVDMXN, "           ' 品ＳＸＤＶＤ２上限   項目追加，修正対応 2003.05.20 yakimura
    sql = sql & "S.E020HSXDVDMNN, "           ' 品ＳＸＤＶＤ２下限   項目追加，修正対応 2003.05.20 yakimura
    sql = sql & "S.E020HSXDENHT, "            ' 品ＳＸＤｅｎ保証方法＿対
    sql = sql & "S.E020HSXDENHS, "            ' 品ＳＸＤｅｎ保証方法＿処
    sql = sql & "S.E020HSXLDLHT, "            ' 品ＳＸＬ／ＤＬ保証方法＿対
    sql = sql & "S.E020HSXLDLHS, "            ' 品ＳＸＬ／ＤＬ保証方法＿処
    sql = sql & "S.E020HSXDVDHT, "            ' 品ＳＸＤＶＤ２保証方法＿対
    sql = sql & "S.E020HSXDVDHS, "            ' 品ＳＸＤＶＤ２保証方法＿処
    sql = sql & "S.E020HSXDENKU, "            ' 品ＳＸＤｅｎ検査有無
    sql = sql & "S.E020HSXDVDKU, "            ' 品ＳＸＤＶＤ２検査有無
    sql = sql & "S.E020HSXLDLKU, "            ' 品ＳＸＬ／ＤＬ検査有無

    sql = sql & "S.E019HSXLTMIN, "            ' 品ＳＸＬタイム下限
    sql = sql & "S.E019HSXLTMAX, "            ' 品ＳＸＬタイム上限
    sql = sql & "S.E019HSXLTSPH, "            ' 品ＳＸＬタイム測定位置＿方
    sql = sql & "S.E019HSXLTSPT, "            ' 品ＳＸＬタイム測定位置＿点
    sql = sql & "S.E019HSXLTSPI, "            ' 品ＳＸＬタイム測定位置＿位
    sql = sql & "S.E019HSXLTHWT, "            ' 品ＳＸＬタイム保証方法＿対
    sql = sql & "S.E019HSXLTHWS, "            ' 品ＳＸＬタイム保証方法＿処
    sql = sql & "U.EPDUP, "                   ' EPD 上限
    sql = sql & "U.TOPREG, "                  ' TOP規制
    sql = sql & "U.TAILREG, "                 ' TAIL規制
    sql = sql & "U.BTMSPRT, "                 ' ボトム析出規制

' OSF，BMD項目追加対応  2002.04.02 yakimura
    sql = sql & "S.E020HSXOSF1PTK, "          ' 品ＳＸＯＳＦ１パタン区分
    sql = sql & "S.E020HSXOSF2PTK, "          ' 品ＳＸＯＳＦ２パタン区分
    sql = sql & "S.E020HSXOSF3PTK, "          ' 品ＳＸＯＳＦ３パタン区分
    sql = sql & "S.E020HSXOSF4PTK, "          ' 品ＳＸＯＳＦ４パタン区分
    sql = sql & "S.E020HSXBMD1MBP, "          ' 品ＳＸＢＭＤ１面内分布
    sql = sql & "S.E020HSXBMD2MBP, "          ' 品ＳＸＢＭＤ２面内分布
    sql = sql & "S.E020HSXBMD3MBP  "          ' 品ＳＸＢＭＤ３面内分布
' OSF，BMD項目追加対応  2002.04.02 yakimura


' 結晶内位置で品番（複数品番の場合）をソートするために取得する必要有り
    sql = sql & ", BH.E041INGOTPOS "

    sql = sql & " from VECME009 BH, VECME001 S, TBCME036 U "
    ' 品番TOPの取得
    sql = sql & " where BH.E040BLOCKID='" & inBlockID & "' " & _
                "   and S.E018HINBAN  =BH.E041HINBAN  and S.E018MNOREVNO=BH.E041REVNUM " & _
                "   and S.E018FACTORY =BH.E041FACTORY and S.E018OPECOND =BH.E041OPECOND" & _
                "   and U.HINBAN      =BH.E041HINBAN  and U.MNOREVNO    =BH.E041REVNUM " & _
                "   and U.FACTORY     =BH.E041FACTORY and U.OPECOND     =BH.E041OPECOND"

' 結晶内位置で品番（複数品番の場合）をソート
    sql = sql & " order by BH.E041INGOTPOS ASC"


    ''-------------↓コメント化【ブロック内全品番仕様取得】 2003/09/09 ooba-------------
'                " and BH.E041INGOTPOS=ANY(select min(E041INGOTPOS) from VECME009 where E040BLOCKID='" & inBlockID & "' ) "

'    sql = sql & " union all "
'    sql = sql & " select "
'    sql = sql & "BH.E040CRYNUM, "             ' 結晶番号
'    sql = sql & "BH.E040INGOTPOS, "           ' 結晶内開始位置
'    sql = sql & "BH.E040LENGTH, "             ' 長さ
'    sql = sql & "BH.E041HINBAN, "             ' 品番
'    sql = sql & "BH.E041REVNUM, "             ' 製品番号改訂番号
'    sql = sql & "BH.E041FACTORY, "            ' 工場
'    sql = sql & "BH.E041OPECOND, "            ' 操業条件
'
'    sql = sql & "BH.E037PRODCOND, "           ' 製作条件
'    sql = sql & "BH.E037PGID, "               ' ＰＧ－ＩＤ
'    sql = sql & "BH.E037UPLENGTH, "           ' 引上げ長さ
'    sql = sql & "BH.E037FREELENG, "           ' フリー長
'    sql = sql & "BH.E037DIAMETER, "           ' 直径
'    sql = sql & "BH.E037CHARGE, "             ' チャージ量
'    sql = sql & "BH.E037SEED, "               ' シード
'    sql = sql & "BH.E037ADDDPPOS, "           ' 追加ドープ位置
'
'    sql = sql & "S.E018HSXTYPE, "             ' 品ＳＸタイプ
'    sql = sql & "S.E018HSXD1CEN, "            ' 品ＳＸ直径１中心
'    sql = sql & "S.E018HSXCDIR, "             ' 品ＳＸ結晶面方位
'
'    sql = sql & "S.E018HSXRMIN, "             ' 品ＳＸ比抵抗下限
'    sql = sql & "S.E018HSXRMAX, "             ' 品ＳＸ比抵抗上限
'    sql = sql & "S.E018HSXRAMIN, "            ' 品ＳＸ比抵抗平均下限
'    sql = sql & "S.E018HSXRAMAX, "            ' 品ＳＸ比抵抗平均上限
'    sql = sql & "S.E018HSXRMCAL, "            ' 品ＳＸ比抵抗面内計算    '' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06
'    sql = sql & "S.E018HSXRMBNP, "            ' 品ＳＸ比抵抗面内分布
'    sql = sql & "S.E018HSXRSPOH, "            ' 品ＳＸ比抵抗測定位置＿方
'    sql = sql & "S.E018HSXRSPOT, "            ' 品ＳＸ比抵抗測定位置＿点
'    sql = sql & "S.E018HSXRSPOI, "            ' 品ＳＸ比抵抗測定位置＿位
'    sql = sql & "S.E018HSXRHWYT, "            ' 品ＳＸ比抵抗保証方法＿対
'    sql = sql & "S.E018HSXRHWYS, "            ' 品ＳＸ比抵抗保証方法＿処
'
'    sql = sql & "S.E019HSXONMIN, "            ' 品ＳＸ酸素濃度下限
'    sql = sql & "S.E019HSXONMAX, "            ' 品ＳＸ酸素濃度上限
'    sql = sql & "S.E019HSXONAMN, "            ' 品ＳＸ酸素濃度平均下限
'    sql = sql & "S.E019HSXONAMX, "            ' 品ＳＸ酸素濃度平均上限
'    sql = sql & "S.E019HSXONMCL, "            ' 品ＳＸ酸素濃度面内計算　 '' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06
'    sql = sql & "S.E019HSXONMBP, "            ' 品ＳＸ酸素濃度面内分布
'    sql = sql & "S.E019HSXONSPH, "            ' 品ＳＸ酸素濃度測定位置＿方
'    sql = sql & "S.E019HSXONSPT, "            ' 品ＳＸ酸素濃度測定位置＿点
'    sql = sql & "S.E019HSXONSPI, "            ' 品ＳＸ酸素濃度測定位置＿位
'    sql = sql & "S.E019HSXONHWT, "            ' 品ＳＸ酸素濃度保証方法＿対
'    sql = sql & "S.E019HSXONHWS, "            ' 品ＳＸ酸素濃度保証方法＿処
'
'    sql = sql & "S.E020HSXBM1AN, "            ' 品ＳＸＢＭＤ１平均下限
'    sql = sql & "S.E020HSXBM1AX, "            ' 品ＳＸＢＭＤ１平均上限
'    sql = sql & "S.E020HSXBM2AN, "            ' 品ＳＸＢＭＤ２平均下限
'    sql = sql & "S.E020HSXBM2AX, "            ' 品ＳＸＢＭＤ２平均上限
'    sql = sql & "S.E020HSXBM3AN, "            ' 品ＳＸＢＭＤ３平均下限
'    sql = sql & "S.E020HSXBM3AX, "            ' 品ＳＸＢＭＤ３平均上限
'    sql = sql & "S.E020HSXBM1SH, "            ' 品ＳＸＢＭＤ１測定位置＿方
'    sql = sql & "S.E020HSXBM1ST, "            ' 品ＳＸＢＭＤ１測定位置＿点
'    sql = sql & "S.E020HSXBM1SR, "            ' 品ＳＸＢＭＤ１測定位置＿領
'    sql = sql & "S.E020HSXBM1HT, "            ' 品ＳＸＢＭＤ１保証方法＿対
'    sql = sql & "S.E020HSXBM1HS, "            ' 品ＳＸＢＭＤ１保証方法＿処
'    sql = sql & "S.E020HSXBM2SH, "            ' 品ＳＸＢＭＤ２測定位置＿方
'    sql = sql & "S.E020HSXBM2ST, "            ' 品ＳＸＢＭＤ２測定位置＿点
'    sql = sql & "S.E020HSXBM2SR, "            ' 品ＳＸＢＭＤ２測定位置＿領
'    sql = sql & "S.E020HSXBM2HT, "            ' 品ＳＸＢＭＤ２保証方法＿対
'    sql = sql & "S.E020HSXBM2HS, "            ' 品ＳＸＢＭＤ２保証方法＿処
'    sql = sql & "S.E020HSXBM3SH, "            ' 品ＳＸＢＭＤ３測定位置＿方
'    sql = sql & "S.E020HSXBM3ST, "            ' 品ＳＸＢＭＤ３測定位置＿点
'    sql = sql & "S.E020HSXBM3SR, "            ' 品ＳＸＢＭＤ３測定位置＿領
'    sql = sql & "S.E020HSXBM3HT, "            ' 品ＳＸＢＭＤ３保証方法＿対
'    sql = sql & "S.E020HSXBM3HS, "            ' 品ＳＸＢＭＤ３保証方法＿処
'
'    sql = sql & "S.E020HSXOF1AX, "            ' 品ＳＸＯＳＦ１平均上限
'    sql = sql & "S.E020HSXOF1MX, "            ' 品ＳＸＯＳＦ１上限
'    sql = sql & "S.E020HSXOF2AX, "            ' 品ＳＸＯＳＦ２平均上限
'    sql = sql & "S.E020HSXOF2MX, "            ' 品ＳＸＯＳＦ２上限
'    sql = sql & "S.E020HSXOF3AX, "            ' 品ＳＸＯＳＦ３平均上限
'    sql = sql & "S.E020HSXOF3MX, "            ' 品ＳＸＯＳＦ３上限
'    sql = sql & "S.E020HSXOF4AX, "            ' 品ＳＸＯＳＦ４平均上限
'    sql = sql & "S.E020HSXOF4MX, "            ' 品ＳＸＯＳＦ４上限
'    sql = sql & "S.E020HSXOF1SH, "            ' 品ＳＸＯＳＦ１測定位置＿方
'    sql = sql & "S.E020HSXOF1ST, "            ' 品ＳＸＯＳＦ１測定位置＿点
'    sql = sql & "S.E020HSXOF1SR, "            ' 品ＳＸＯＳＦ１測定位置＿領
'    sql = sql & "S.E020HSXOF1HT, "            ' 品ＳＸＯＳＦ１保証方法＿対
'    sql = sql & "S.E020HSXOF1HS, "            ' 品ＳＸＯＳＦ１保証方法＿処
'    sql = sql & "S.E020HSXOF2SH, "            ' 品ＳＸＯＳＦ２測定位置＿方
'    sql = sql & "S.E020HSXOF2ST, "            ' 品ＳＸＯＳＦ２測定位置＿点
'    sql = sql & "S.E020HSXOF2SR, "            ' 品ＳＸＯＳＦ２測定位置＿領
'    sql = sql & "S.E020HSXOF2HT, "            ' 品ＳＸＯＳＦ２保証方法＿対
'    sql = sql & "S.E020HSXOF2HS, "            ' 品ＳＸＯＳＦ２保証方法＿処
'    sql = sql & "S.E020HSXOF3SH, "            ' 品ＳＸＯＳＦ３測定位置＿方
'    sql = sql & "S.E020HSXOF3ST, "            ' 品ＳＸＯＳＦ３測定位置＿点
'    sql = sql & "S.E020HSXOF3SR, "            ' 品ＳＸＯＳＦ３測定位置＿領
'    sql = sql & "S.E020HSXOF3HT, "            ' 品ＳＸＯＳＦ３保証方法＿対
'    sql = sql & "S.E020HSXOF3HS, "            ' 品ＳＸＯＳＦ３保証方法＿処
'    sql = sql & "S.E020HSXOF4SH, "            ' 品ＳＸＯＳＦ４測定位置＿方
'    sql = sql & "S.E020HSXOF4ST, "            ' 品ＳＸＯＳＦ４測定位置＿点
'    sql = sql & "S.E020HSXOF4SR, "            ' 品ＳＸＯＳＦ４測定位置＿領
'    sql = sql & "S.E020HSXOF4HT, "            ' 品ＳＸＯＳＦ４保証方法＿対
'    sql = sql & "S.E020HSXOF4HS, "            ' 品ＳＸＯＳＦ４保証方法＿処
'    sql = sql & "S.E020HSXOF1NS, "            ' 品ＳＸＯＳＦ１熱処理法
'    sql = sql & "S.E020HSXOF2NS, "            ' 品ＳＸＯＳＦ２熱処理法
'    sql = sql & "S.E020HSXOF3NS, "            ' 品ＳＸＯＳＦ３熱処理法
'    sql = sql & "S.E020HSXOF4NS, "            ' 品ＳＸＯＳＦ４熱処理法
'    sql = sql & "S.E020HSXBM1NS, "            ' 品ＳＸＢＭＤ１熱処理法
'    sql = sql & "S.E020HSXBM2NS, "            ' 品ＳＸＢＭＤ２熱処理法
'    sql = sql & "S.E020HSXBM3NS, "            ' 品ＳＸＢＭＤ３熱処理法
'
'    sql = sql & "S.E019HSXCNMIN, "            ' 品ＳＸ炭素濃度下限
'    sql = sql & "S.E019HSXCNMAX, "            ' 品ＳＸ炭素濃度上限
'    sql = sql & "S.E019HSXCNSPH, "            ' 品ＳＸ炭素濃度測定位置＿方
'    sql = sql & "S.E019HSXCNSPT, "            ' 品ＳＸ炭素濃度測定位置＿点
'    sql = sql & "S.E019HSXCNSPI, "            ' 品ＳＸ炭素濃度測定位置＿位
'    sql = sql & "S.E019HSXCNHWT, "            ' 品ＳＸ炭素濃度保証方法＿対
'    sql = sql & "S.E019HSXCNHWS, "            ' 品ＳＸ炭素濃度保証方法＿処
'
'    sql = sql & "S.E020HSXDENMX, "            ' 品ＳＸＤｅｎ上限
'    sql = sql & "S.E020HSXDENMN, "            ' 品ＳＸＤｅｎ下限
'    sql = sql & "S.E020HSXLDLMX, "            ' 品ＳＸＬ／ＤＬ上限
'    sql = sql & "S.E020HSXLDLMN, "            ' 品ＳＸＬ／ＤＬ下限
'    sql = sql & "S.E020HSXDVDMXN, "           ' 品ＳＸＤＶＤ２上限   項目追加，修正対応 2003.05.20 yakimura
'    sql = sql & "S.E020HSXDVDMNN, "           ' 品ＳＸＤＶＤ２下限   項目追加，修正対応 2003.05.20 yakimura
'    sql = sql & "S.E020HSXDENHT, "            ' 品ＳＸＤｅｎ保証方法＿対
'    sql = sql & "S.E020HSXDENHS, "            ' 品ＳＸＤｅｎ保証方法＿処
'    sql = sql & "S.E020HSXLDLHT, "            ' 品ＳＸＬ／ＤＬ保証方法＿対
'    sql = sql & "S.E020HSXLDLHS, "            ' 品ＳＸＬ／ＤＬ保証方法＿処
'    sql = sql & "S.E020HSXDVDHT, "            ' 品ＳＸＤＶＤ２保証方法＿対
'    sql = sql & "S.E020HSXDVDHS, "            ' 品ＳＸＤＶＤ２保証方法＿処
'    sql = sql & "S.E020HSXDENKU, "            ' 品ＳＸＤｅｎ検査有無
'    sql = sql & "S.E020HSXDVDKU, "            ' 品ＳＸＤＶＤ２検査有無
'    sql = sql & "S.E020HSXLDLKU, "            ' 品ＳＸＬ／ＤＬ検査有無
'
'    sql = sql & "S.E019HSXLTMIN, "            ' 品ＳＸＬタイム下限
'    sql = sql & "S.E019HSXLTMAX, "            ' 品ＳＸＬタイム上限
'    sql = sql & "S.E019HSXLTSPH, "            ' 品ＳＸＬタイム測定位置＿方
'    sql = sql & "S.E019HSXLTSPT, "            ' 品ＳＸＬタイム測定位置＿点
'    sql = sql & "S.E019HSXLTSPI, "            ' 品ＳＸＬタイム測定位置＿位
'    sql = sql & "S.E019HSXLTHWT, "            ' 品ＳＸＬタイム保証方法＿対
'    sql = sql & "S.E019HSXLTHWS, "            ' 品ＳＸＬタイム保証方法＿処
'    sql = sql & "U.EPDUP, "                   ' EPD 上限
'    sql = sql & "U.TOPREG, "                  ' TOP規制
'    sql = sql & "U.TAILREG, "                 ' TAIL規制
'    sql = sql & "U.BTMSPRT, "                 ' ボトム析出規制
'
'' OSF，BMD項目追加対応  2002.04.02 yakimura
'    sql = sql & "S.E020HSXOSF1PTK, "          ' 品ＳＸＯＳＦ１パタン区分
'    sql = sql & "S.E020HSXOSF2PTK, "          ' 品ＳＸＯＳＦ２パタン区分
'    sql = sql & "S.E020HSXOSF3PTK, "          ' 品ＳＸＯＳＦ３パタン区分
'    sql = sql & "S.E020HSXOSF4PTK, "          ' 品ＳＸＯＳＦ４パタン区分
'    sql = sql & "S.E020HSXBMD1MBP, "          ' 品ＳＸＢＭＤ１面内分布
'    sql = sql & "S.E020HSXBMD2MBP, "          ' 品ＳＸＢＭＤ２面内分布
'    sql = sql & "S.E020HSXBMD3MBP  "          ' 品ＳＸＢＭＤ３面内分布
'' OSF，BMD項目追加対応  2002.04.02 yakimura
'
'    sql = sql & " from VECME009 BH, VECME001 S, TBCME036 U "
'    '品番TAILの取得
'    sql = sql & " where BH.E040BLOCKID='" & inBlockID & "' " & _
'                " and BH.E041INGOTPOS=ANY(select max(E041INGOTPOS) from VECME009 where E040BLOCKID='" & inBlockID & "' ) " & _
'                " and S.E018HINBAN=BH.E041HINBAN and S.E018MNOREVNO=BH.E041REVNUM" & _
'                " and S.E018FACTORY=BH.E041FACTORY and S.E018OPECOND=BH.E041OPECOND " & _
'                " and U.HINBAN=BH.E041HINBAN and U.MNOREVNO=BH.E041REVNUM" & _
'                " and U.FACTORY=BH.E041FACTORY and U.OPECOND=BH.E041OPECOND "
    ''-------------↑コメント化【ブロック内全品番仕様取得】 2003/09/09 ooba-------------

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
    End If

    recCnt = rs.RecordCount
    BlkHinCNT = rs.RecordCount             ''仕様表示品番数取得  2003/09/09 ooba

    ReDim Siyou(recCnt)
    For i = 1 To recCnt

        With Siyou(i)
            .CRYNUM = rs("E040CRYNUM")                ' 結晶番号
            .IngotPos = rs("E040INGOTPOS")            ' 結晶内開始位置
            .LENGTH = rs("E040LENGTH")                ' 長さ
            .hin.hinban = rs("E041HINBAN")            ' 品番
            .hin.mnorevno = rs("E041REVNUM")          ' 製品番号改訂番号
            .hin.factory = rs("E041FACTORY")          ' 工場
            .hin.opecond = rs("E041OPECOND")          ' 操業条件

            .PRODCOND = rs("E037PRODCOND")            ' 製作条件
            .PGID = rs("E037PGID")                    ' ＰＧ－ＩＤ
            .UPLENGTH = rs("E037UPLENGTH")            ' 引上げ長さ
            .FREELENG = rs("E037FREELENG")            ' フリー長
            .DIAMETER = rs("E037DIAMETER")            ' 直径
            .CHARGE = rs("E037CHARGE")                ' チャージ量
            .SEED = rs("E037SEED")                    ' シード
            .ADDDPPOS = rs("E037ADDDPPOS")            ' 追加ドープ位置

            .HSXTYPE = rs("E018HSXTYPE")              ' 品ＳＸタイプ"
            .HSXD1CEN = fncNullCheck(rs("E018HSXD1CEN"))            ' 品ＳＸ直径１中心"
            .HSXCDIR = rs("E018HSXCDIR")              ' 品ＳＸ結晶面方位"

            .HSXRMIN = fncNullCheck(rs("E018HSXRMIN"))              ' 品ＳＸ比抵抗下限  'NULL対応
            .HSXRMAX = fncNullCheck(rs("E018HSXRMAX"))              ' 品ＳＸ比抵抗上限　'NULL対応
            .HSXRAMIN = fncNullCheck(rs("E018HSXRAMIN"))            ' 品ＳＸ比抵抗平均下限　'NULL対応
            .HSXRAMAX = fncNullCheck(rs("E018HSXRAMAX"))            ' 品ＳＸ比抵抗平均上限  'NULL対応
            .HSXRMCAL = rs("E018HSXRMCAL")            ' 品ＳＸ比抵抗面内計算     '' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06
            .HSXRMBNP = fncNullCheck(rs("E018HSXRMBNP"))            ' 品ＳＸ比抵抗面内分布  'NULL対応
            .HSXRSPOH = rs("E018HSXRSPOH")            ' 品ＳＸ比抵抗測定位置＿方
            .HSXRSPOT = rs("E018HSXRSPOT")            ' 品ＳＸ比抵抗測定位置＿点
            .HSXRSPOI = rs("E018HSXRSPOI")            ' 品ＳＸ比抵抗測定位置＿位
            .HSXRHWYT = rs("E018HSXRHWYT")            ' 品ＳＸ比抵抗保証方法＿対
            .HSXRHWYS = rs("E018HSXRHWYS")            ' 品ＳＸ比抵抗保証方法＿処

            .HSXONMIN = fncNullCheck(rs("E019HSXONMIN"))            ' 品ＳＸ酸素濃度下限
            .HSXONMAX = fncNullCheck(rs("E019HSXONMAX"))            ' 品ＳＸ酸素濃度上限
            .HSXONAMN = fncNullCheck(rs("E019HSXONAMN"))            ' 品ＳＸ酸素濃度平均下限
            .HSXONAMX = fncNullCheck(rs("E019HSXONAMX"))            ' 品ＳＸ酸素濃度平均上限
            .HSXONMCL = rs("E019HSXONMCL")            ' 品ＳＸ酸素濃度面内計算   '' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06
            .HSXONMBP = fncNullCheck(rs("E019HSXONMBP"))            ' 品ＳＸ酸素濃度面内分布
            .HSXONSPH = rs("E019HSXONSPH")            ' 品ＳＸ酸素濃度測定位置＿方
            .HSXONSPT = rs("E019HSXONSPT")            ' 品ＳＸ酸素濃度測定位置＿点
            .HSXONSPI = rs("E019HSXONSPI")            ' 品ＳＸ酸素濃度測定位置＿位
            .HSXONHWT = rs("E019HSXONHWT")            ' 品ＳＸ酸素濃度保証方法＿対
            .HSXONHWS = rs("E019HSXONHWS")            ' 品ＳＸ酸素濃度保証方法＿処

           '.HSXBM1AN = rs("E020HSXBM1AN") * 10       ' 品ＳＸＢＭＤ１平均下限
           '.HSXBM1AX = rs("E020HSXBM1AX") * 10       ' 品ＳＸＢＭＤ１平均上限
           '.HSXBM2AN = rs("E020HSXBM2AN") * 10       ' 品ＳＸＢＭＤ２平均下限
           '.HSXBM2AX = rs("E020HSXBM2AX") * 10       ' 品ＳＸＢＭＤ２平均上限
           '.HSXBM3AN = rs("E020HSXBM3AN") * 10       ' 品ＳＸＢＭＤ３平均下限
           '.HSXBM3AX = rs("E020HSXBM3AX") * 10       ' 品ＳＸＢＭＤ３平均上限
            'BMDべき乗数　変更対応　2003/05/17 osawa
            .HSXBM1AN = fncNullCheck(rs("E020HSXBM1AN"))            ' 品ＳＸＢＭＤ１平均下限
            .HSXBM1AX = fncNullCheck(rs("E020HSXBM1AX"))            ' 品ＳＸＢＭＤ１平均上限
            .HSXBM2AN = fncNullCheck(rs("E020HSXBM2AN"))            ' 品ＳＸＢＭＤ２平均下限
            .HSXBM2AX = fncNullCheck(rs("E020HSXBM2AX"))            ' 品ＳＸＢＭＤ２平均上限
            .HSXBM3AN = fncNullCheck(rs("E020HSXBM3AN"))            ' 品ＳＸＢＭＤ３平均下限
            .HSXBM3AX = fncNullCheck(rs("E020HSXBM3AX"))            ' 品ＳＸＢＭＤ３平均上限
            '
            .HSXBM1SH = rs("E020HSXBM1SH")            ' 品ＳＸＢＭＤ１測定位置＿方
            .HSXBM1ST = rs("E020HSXBM1ST")            ' 品ＳＸＢＭＤ１測定位置＿点
            .HSXBM1SR = rs("E020HSXBM1SR")            ' 品ＳＸＢＭＤ１測定位置＿領
            .HSXBM1HT = rs("E020HSXBM1HT")            ' 品ＳＸＢＭＤ１保証方法＿対
            .HSXBM1HS = rs("E020HSXBM1HS")            ' 品ＳＸＢＭＤ１保証方法＿処
            .HSXBM2SH = rs("E020HSXBM2SH")            ' 品ＳＸＢＭＤ２測定位置＿方
            .HSXBM2ST = rs("E020HSXBM2ST")            ' 品ＳＸＢＭＤ２測定位置＿点
            .HSXBM2SR = rs("E020HSXBM2SR")            ' 品ＳＸＢＭＤ２測定位置＿領
            .HSXBM2HT = rs("E020HSXBM2HT")            ' 品ＳＸＢＭＤ２保証方法＿対
            .HSXBM2HS = rs("E020HSXBM2HS")            ' 品ＳＸＢＭＤ２保証方法＿処
            .HSXBM3SH = rs("E020HSXBM3SH")            ' 品ＳＸＢＭＤ３測定位置＿方
            .HSXBM3ST = rs("E020HSXBM3ST")            ' 品ＳＸＢＭＤ３測定位置＿点
            .HSXBM3SR = rs("E020HSXBM3SR")            ' 品ＳＸＢＭＤ３測定位置＿領
            .HSXBM3HT = rs("E020HSXBM3HT")            ' 品ＳＸＢＭＤ３保証方法＿対
            .HSXBM3HS = rs("E020HSXBM3HS")            ' 品ＳＸＢＭＤ３保証方法＿処

            .HSXOF1AX = fncNullCheck(rs("E020HSXOF1AX"))            ' 品ＳＸＯＳＦ１平均上限
            .HSXOF1MX = fncNullCheck(rs("E020HSXOF1MX"))            ' 品ＳＸＯＳＦ１上限
            .HSXOF2AX = fncNullCheck(rs("E020HSXOF2AX"))            ' 品ＳＸＯＳＦ２平均上限
            .HSXOF2MX = fncNullCheck(rs("E020HSXOF2MX"))            ' 品ＳＸＯＳＦ２上限
            .HSXOF3AX = fncNullCheck(rs("E020HSXOF3AX"))            ' 品ＳＸＯＳＦ３平均上限
            .HSXOF3MX = fncNullCheck(rs("E020HSXOF3MX"))            ' 品ＳＸＯＳＦ３上限
            .HSXOF4AX = fncNullCheck(rs("E020HSXOF4AX"))            ' 品ＳＸＯＳＦ４平均上限
            .HSXOF4MX = fncNullCheck(rs("E020HSXOF4MX"))            ' 品ＳＸＯＳＦ４上限
            .HSXOF1SH = rs("E020HSXOF1SH")            ' 品ＳＸＯＳＦ１測定位置＿方
            .HSXOF1ST = rs("E020HSXOF1ST")            ' 品ＳＸＯＳＦ１測定位置＿点
            .HSXOF1SR = rs("E020HSXOF1SR")            ' 品ＳＸＯＳＦ１測定位置＿領
            .HSXOF1HT = rs("E020HSXOF1HT")            ' 品ＳＸＯＳＦ１保証方法＿対
            .HSXOF1HS = rs("E020HSXOF1HS")            ' 品ＳＸＯＳＦ１保証方法＿処
            .HSXOF2SH = rs("E020HSXOF2SH")            ' 品ＳＸＯＳＦ２測定位置＿方
            .HSXOF2ST = rs("E020HSXOF2ST")            ' 品ＳＸＯＳＦ２測定位置＿点
            .HSXOF2SR = rs("E020HSXOF2SR")            ' 品ＳＸＯＳＦ２測定位置＿領
            .HSXOF2HT = rs("E020HSXOF2HT")            ' 品ＳＸＯＳＦ２保証方法＿対
            .HSXOF2HS = rs("E020HSXOF2HS")            ' 品ＳＸＯＳＦ２保証方法＿処
            .HSXOF3SH = rs("E020HSXOF3SH")            ' 品ＳＸＯＳＦ３測定位置＿方
            .HSXOF3ST = rs("E020HSXOF3ST")            ' 品ＳＸＯＳＦ３測定位置＿点
            .HSXOF3SR = rs("E020HSXOF3SR")            ' 品ＳＸＯＳＦ３測定位置＿領
            .HSXOF3HT = rs("E020HSXOF3HT")            ' 品ＳＸＯＳＦ３保証方法＿対
            .HSXOF3HS = rs("E020HSXOF3HS")            ' 品ＳＸＯＳＦ３保証方法＿処
            .HSXOF4SH = rs("E020HSXOF4SH")            ' 品ＳＸＯＳＦ４測定位置＿方
            .HSXOF4ST = rs("E020HSXOF4ST")            ' 品ＳＸＯＳＦ４測定位置＿点
            .HSXOF4SR = rs("E020HSXOF4SR")            ' 品ＳＸＯＳＦ４測定位置＿領
            .HSXOF4HT = rs("E020HSXOF4HT")            ' 品ＳＸＯＳＦ４保証方法＿対
            .HSXOF4HS = rs("E020HSXOF4HS")            ' 品ＳＸＯＳＦ４保証方法＿処
            .HSXOF1NS = rs("E020HSXOF1NS")            ' 品ＳＸＯＳＦ１熱処理法
            .HSXOF2NS = rs("E020HSXOF2NS")            ' 品ＳＸＯＳＦ２熱処理法
            .HSXOF3NS = rs("E020HSXOF3NS")            ' 品ＳＸＯＳＦ３熱処理法
            .HSXOF4NS = rs("E020HSXOF4NS")            ' 品ＳＸＯＳＦ４熱処理法
            .HSXBM1NS = rs("E020HSXBM1NS")            ' 品ＳＸＢＭＤ１熱処理法
            .HSXBM2NS = rs("E020HSXBM2NS")            ' 品ＳＸＢＭＤ２熱処理法
            .HSXBM3NS = rs("E020HSXBM3NS")            ' 品ＳＸＢＭＤ３熱処理法

            .HSXCNMIN = fncNullCheck(rs("E019HSXCNMIN"))            ' 品ＳＸ炭素濃度下限
            .HSXCNMAX = fncNullCheck(rs("E019HSXCNMAX"))            ' 品ＳＸ炭素濃度上限
            .HSXCNSPH = rs("E019HSXCNSPH")            ' 品ＳＸ炭素濃度測定位置＿方
            .HSXCNSPT = rs("E019HSXCNSPT")            ' 品ＳＸ炭素濃度測定位置＿点
            .HSXCNSPI = rs("E019HSXCNSPI")            ' 品ＳＸ炭素濃度測定位置＿位
            .HSXCNHWT = rs("E019HSXCNHWT")            ' 品ＳＸ炭素濃度保証方法＿対
            .HSXCNHWS = rs("E019HSXCNHWS")            ' 品ＳＸ炭素濃度保証方法＿処

            .HSXDENMX = fncNullCheck(rs("E020HSXDENMX"))            ' 品ＳＸＤｅｎ上限
            .HSXDENMN = fncNullCheck(rs("E020HSXDENMN"))            ' 品ＳＸＤｅｎ下限
            .HSXLDLMX = fncNullCheck(rs("E020HSXLDLMX"))            ' 品ＳＸＬ／ＤＬ上限
            .HSXLDLMN = fncNullCheck(rs("E020HSXLDLMN"))            ' 品ＳＸＬ／ＤＬ下限
            .HSXDVDMX = fncNullCheck(rs("E020HSXDVDMXN"))           ' 品ＳＸＤＶＤ２上限   項目追加，修正対応 2003.05.20 yakimura
            .HSXDVDMN = fncNullCheck(rs("E020HSXDVDMNN"))           ' 品ＳＸＤＶＤ２下限   項目追加，修正対応 2003.05.20 yakimura
            .HSXDENHT = rs("E020HSXDENHT")            ' 品ＳＸＤｅｎ保証方法＿対
            .HSXDENHS = rs("E020HSXDENHS")            ' 品ＳＸＤｅｎ保証方法＿処
            .HSXLDLHT = rs("E020HSXLDLHT")            ' 品ＳＸＬ／ＤＬ保証方法＿対
            .HSXLDLHS = rs("E020HSXLDLHS")            ' 品ＳＸＬ／ＤＬ保証方法＿処
            .HSXDVDHT = rs("E020HSXDVDHT")            ' 品ＳＸＤＶＤ２保証方法＿対
            .HSXDVDHS = rs("E020HSXDVDHS")            ' 品ＳＸＤＶＤ２保証方法＿処
            .HSXDENKU = rs("E020HSXDENKU")            ' 品ＳＸＤｅｎ検査有無
            .HSXDVDKU = rs("E020HSXDVDKU")            ' 品ＳＸＤＶＤ２検査有無
            .HSXLDLKU = rs("E020HSXLDLKU")            ' 品ＳＸＬ／ＤＬ検査有無

            .HSXLTMIN = fncNullCheck(rs("E019HSXLTMIN"))            ' 品ＳＸＬタイム下限
            .HSXLTMAX = fncNullCheck(rs("E019HSXLTMAX"))            ' 品ＳＸＬタイム上限
            .HSXLTSPH = rs("E019HSXLTSPH")            ' 品ＳＸＬタイム測定位置＿方
            .HSXLTSPT = rs("E019HSXLTSPT")            ' 品ＳＸＬタイム測定位置＿点
            .HSXLTSPI = rs("E019HSXLTSPI")            ' 品ＳＸＬタイム測定位置＿位
            .HSXLTHWT = rs("E019HSXLTHWT")            ' 品ＳＸＬタイム保証方法＿対
            .HSXLTHWS = rs("E019HSXLTHWS")            ' 品ＳＸＬタイム保証方法＿処
'''''       .EPDUP = rs("EPDUP")                      ' EPD上限
'''''       .TOPREG = rs("TOPREG")                    ' TOP規制
'''''       .TAILREG = rs("TAILREG")                  ' TAIL規制
'''''       .BTMSPRT = rs("BTMSPRT")                  ' ボトム析出規制
'''''       --TEST--
'''''       NULL対応仮セット
            .EPDUP = IIf(IsNull(rs("EPDUP")) = True, 0, rs("EPDUP"))
            .TOPREG = IIf(IsNull(rs("TOPREG")) = True, 0, rs("TOPREG"))
            .TAILREG = IIf(IsNull(rs("TAILREG")) = True, 0, rs("TAILREG"))
            .BTMSPRT = IIf(IsNull(rs("BTMSPRT")) = True, 0, rs("BTMSPRT"))

' OSF，BMD項目追加対応  2002.04.02 yakimura
            If IsNull(rs("E020HSXOSF1PTK")) = False Then .HSXOSF1PTK = rs("E020HSXOSF1PTK")   ' 品ＳＸＯＳＦ１パタン区分
            If IsNull(rs("E020HSXOSF2PTK")) = False Then .HSXOSF2PTK = rs("E020HSXOSF2PTK")   ' 品ＳＸＯＳＦ２パタン区分
            If IsNull(rs("E020HSXOSF3PTK")) = False Then .HSXOSF3PTK = rs("E020HSXOSF3PTK")   ' 品ＳＸＯＳＦ３パタン区分
            If IsNull(rs("E020HSXOSF4PTK")) = False Then .HSXOSF4PTK = rs("E020HSXOSF4PTK")   ' 品ＳＸＯＳＦ４パタン区分
            If IsNull(rs("E020HSXBMD1MBP")) = False Then .HSXBMD1MBP = rs("E020HSXBMD1MBP")   ' 品ＳＸＢＭＤ１面内分布
            If IsNull(rs("E020HSXBMD2MBP")) = False Then .HSXBMD2MBP = rs("E020HSXBMD2MBP")   ' 品ＳＸＢＭＤ２面内分布
            If IsNull(rs("E020HSXBMD3MBP")) = False Then .HSXBMD3MBP = rs("E020HSXBMD3MBP")   ' 品ＳＸＢＭＤ３面内分布
' OSF，BMD項目追加対応  2002.04.02 yakimura

        End With
        rs.MoveNext
    Next

    If scmzc_getKakouJiltuseki(inBlockID, Jiltuseki) = FUNCTION_RETURN_FAILURE Then
        getHinSiyou30 = FUNCTION_RETURN_FAILURE
        ReDim Siyou(0)
        GoTo proc_exit
    End If
    For i = 1 To recCnt
        Siyou(i).DIAMETER = (Jiltuseki.Top(1) + Jiltuseki.Top(2) + Jiltuseki.TAIL(1) + Jiltuseki.TAIL(2)) / 4 ' 直径
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
    gErr.HandleError
    Resume proc_exit
End Function


''''''概要      :内部関数 サンプル番号を取得する
'''''Private Function getCrySmp(inCryNum As String, inIngotPos, _
'''''                           CrySmp() As type_DBDRV_scmzc_fcmkc001c_CrySmp _
'''''                           ) As FUNCTION_RETURN
'''''
'''''
'''''    Dim sql     As String
'''''    Dim rs      As OraDynaset
'''''    Dim recCnt  As Integer
'''''    Dim i       As Long
'''''
'''''    'サンプル番号取得
'''''    'VECME010（ブロック管理を検索し、そのブロックに対するサンプルを表示するビュー）から値を取得
'''''
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function getCrySmp"
'''''
'''''    sql = "select "
'''''    sql = sql & "XTALCS, "              ' 結晶番号
'''''    sql = sql & "INPOSCS, "             ' 結晶内位置
'''''    sql = sql & "B.LENGTH, "            ' 長さ
'''''    sql = sql & "B.BLOCKID, "           ' ブロックID
'''''    sql = sql & "SMPKBNCS, "            ' サンプル区分
'''''    sql = sql & "REPSMPLIDCS, "         ' サンプルNo
'''''    sql = sql & "HINBCS, "              ' 品番
'''''    sql = sql & "REVNUMCS, "            ' 製品番号改訂番号
'''''    sql = sql & "FACTORYCS, "           ' 工場
'''''    sql = sql & "OPECS, "               ' 操業条件
'''''    sql = sql & "KTKBNCS, "             ' 確定区分
'''''    sql = sql & "CRYINDRSCS, "          ' 結晶検査指示（Rs)
'''''    sql = sql & "CRYINDOICS, "          ' 結晶検査指示（Oi)
'''''    sql = sql & "CRYINDB1CS, "          ' 結晶検査指示（B1)
'''''    sql = sql & "CRYINDB2CS, "          ' 結晶検査指示（B2）
'''''    sql = sql & "CRYINDB3CS, "          ' 結晶検査指示（B3)
'''''    sql = sql & "CRYINDL1CS, "          ' 結晶検査指示（L1)
'''''    sql = sql & "CRYINDL2CS, "          ' 結晶検査指示（L2)
'''''    sql = sql & "CRYINDL3CS, "          ' 結晶検査指示（L3)
'''''    sql = sql & "CRYINDL4CS, "          ' 結晶検査指示（L4)
'''''    sql = sql & "CRYINDCSCS, "          ' 結晶検査指示（Cs)
'''''    sql = sql & "CRYINDGDCS, "          ' 結晶検査指示（GD)
'''''    sql = sql & "CRYINDTCS, "           ' 結晶検査指示（T)
'''''    sql = sql & "CRYINDEPCS "           ' 結晶検査指示（EPD)
'''''
'''''    sql = sql & " from  TBCME040 B, XSDCS X "
'''''    sql = sql & " where B.CRYNUM  ='" & inCryNum & "' "
'''''    sql = sql & "   and B.INGOTPOS= " & inIngotPos
'''''    sql = sql & "   and B.BLOCKID = X.CRYNUMCS "
'''''    sql = sql & " order by X.INPOSCS "  ' TOP TAIL順
'''''
'''''
''''''''''    sql = sql & " from VECME010"
''''''''''    sql = sql & " where E040CRYNUM='" & inCryNum & "' and E040INGOTPOS=" & inIngotPos
''''''''''    sql = sql & " order by E043INPOSCS " ' TOP TAIL順
'''''
'''''
'''''
'''''    ' SQL実行
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''Debug.Print sql
'''''
'''''    If rs.RecordCount = 0 Then
'''''        getCrySmp = FUNCTION_RETURN_FAILURE
'''''        ReDim CrySmp(0)
'''''        GoTo proc_exit
'''''    End If
'''''
'''''    recCnt = rs.RecordCount
'''''    ReDim CrySmp(recCnt)
'''''    For i = 1 To recCnt
'''''        With CrySmp(i)
'''''
''''''            .CRYNUM = rs("XTALCS")              ' 結晶番号
''''''            .INGOTPOS = rs("INPOSCS")           ' 結晶内位置
''''''            .LENGTH = rs("LENGTH")              ' 長さ
''''''            .BLOCKID = rs("BLOCKID")            ' ブロックID
''''''            .SMPKBN = rs("SMPKBNCS")            ' サンプル区分
''''''            .SMPLNO = rs("REPSMPLIDCS")         ' サンプルNo
''''''            .hinban = rs("HINBCS")              ' 品番
''''''            .REVNUM = rs("REVNUMCS")            ' 製品番号改訂番号
''''''            .factory = rs("FACTORYCS")          ' 工場
''''''            .opecond = rs("OPECS")              ' 操業条件
''''''            .KTKBN = rs("KTKBNCS")              ' 確定区分
''''''            .CRYINDRS = rs("CRYINDRSCS")        ' 結晶検査指示（Rs)
''''''            .CRYINDOI = rs("CRYINDOICS")        ' 結晶検査指示（Oi)
''''''            .CRYINDB1 = rs("CRYINDB1CS")        ' 結晶検査指示（B1)
''''''            .CRYINDB2 = rs("CRYINDB2CS")        ' 結晶検査指示（B2)
''''''            .CRYINDB3 = rs("CRYINDB3CS")        ' 結晶検査指示（B3)
''''''            .CRYINDL1 = rs("CRYINDL1CS")        ' 結晶検査指示（L1)
''''''            .CRYINDL2 = rs("CRYINDL2CS")        ' 結晶検査指示（L2)
''''''            .CRYINDL3 = rs("CRYINDL3CS")        ' 結晶検査指示（L3)
''''''            .CRYINDL4 = rs("CRYINDL4CS")        ' 結晶検査指示（L4)
''''''            .CRYINDCS = rs("CRYINDCSCS")        ' 結晶検査指示（Cs)
''''''            .CRYINDGD = rs("CRYINDGDCS")        ' 結晶検査指示（GD)
''''''            .CRYINDT = rs("CRYINDTCS")          ' 結晶検査指示（T)
''''''            .CRYINDEP = rs("CRYINDEPCS")        ' 結晶検査指示（EPD)
'''''
'''''        End With
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''
'''''    getCrySmp = FUNCTION_RETURN_SUCCESS
'''''
'''''
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    getCrySmp = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


'概要      :総合判定 表示用ＤＢドライバ（ブロックＩＤが米沢の場合）
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                 ,説明
'          :inBlockID     ,I  ,String                             ,対象ブロックID
'          :Siyou()       ,O  ,type_DBDRV_scmzc_fcmkc001c_Siyou   ,品番、仕様、結晶内側取得用
'          :CrySmp()      ,O  ,type_DBDRV_scmzc_fcmkc001c_CrySmp  ,結晶サンプル管理取得用
'          :Zisseki       ,O  ,type_DBDRV_scmzc_fcmkc001c_Zisseki ,実績用
'          :sErrMsg       ,O  ,String                             ,
'          :戻り値        ,O  ,FUNCTION_RETURN                    ,読み込み成否
'説明      :
'履歴      :2001/06/26 蔵本 作成
'''''Public Function DBDRV_scmzc_fcmkc001c_Disp(inBlockID As String, _
'''''                                           Siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                                           CrySmp() As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
'''''                                           Zisseki As type_DBDRV_scmzc_fcmkc001c_Zisseki, _
'''''                                           sErrMsg As String) As FUNCTION_RETURN
Public Function DBDRV_scmzc_fcmkc001c_Disp(inBlockID As String, _
                                           Siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                                           sErrMsg As String) As FUNCTION_RETURN
''''Dim i       As Integer
''''Dim recCnt  As Integer
    Dim sDbName As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_Disp"

    DBDRV_scmzc_fcmkc001c_Disp = FUNCTION_RETURN_FAILURE

    sDbName = "V011"
    '品番、SXL仕様からデータの取得（レコード0件の場合もエラー）
    If FUNCTION_RETURN_FAILURE = getHinSiyou30(inBlockID, Siyou()) Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If


'''' -TEST-
'''''    sDbName = "V010"
'''''    '結晶サンプルの取得(レコード0件の場合もエラー)
'''''    If FUNCTION_RETURN_FAILURE = getCrySmp(Siyou(1).CRYNUM, Siyou(1).INGOTPOS, CrySmp()) Then
'''''        sErrMsg = GetMsgStr("EGET2", sDbName)
'''''        DBDRV_scmzc_fcmkc001c_Disp = FUNCTION_RETURN_FAILURE
'''''        GoTo proc_exit
'''''    End If
'''''
'''''
'''''
'''''
'''''    With Zisseki
'''''        ReDim .CRYRZ(2)
'''''        ReDim .OIZ(2)
'''''        ReDim .BMD1Z(2)
'''''        ReDim .BMD2Z(2)
'''''        ReDim .BMD3Z(2)
'''''        ReDim .OSF1Z(2)
'''''        ReDim .OSF2Z(2)
'''''        ReDim .OSF3Z(2)
'''''        ReDim .OSF4Z(2)
'''''        ReDim .csz(2)
'''''        ReDim .GDZ(2)
'''''        ReDim .LTZ(2)
'''''        ReDim .EPDZ(2)
'''''        ReDim .SURSZ(2)
'''''    End With
'''''
'''''    'recCnt = UBound(CrySmp)
'''''    '結晶サンプルの指示を見て実績を取る
'''''    For i = 1 To 2 'recCnt
'''''        'サンプル管理の品番を鵜呑みにしない
'''''        CrySmp(i).hinban = Siyou(i).hin.hinban
'''''        CrySmp(i).REVNUM = Siyou(i).hin.mnorevno
'''''        CrySmp(i).factory = Siyou(i).hin.factory
'''''        CrySmp(i).opecond = Siyou(i).hin.opecond
'''''
'''''        sDbName = "J002"
'''''        If CryR_Zisseki(Siyou(i), CrySmp(i), Zisseki.CRYRZ(i), Zisseki.SURSZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J003"
'''''        If Oi_Zisseki(Siyou(i), CrySmp(i), Zisseki.OIZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J008"
'''''        If BMD_Zisseki(Siyou(i), CrySmp(i), "1", Zisseki.BMD1Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J008"
'''''        If BMD_Zisseki(Siyou(i), CrySmp(i), "2", Zisseki.BMD2Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J008"
'''''        If BMD_Zisseki(Siyou(i), CrySmp(i), "3", Zisseki.BMD3Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J005"
'''''        If OSF_Zisseki(Siyou(i), CrySmp(i), "1", Zisseki.OSF1Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J005"
'''''        If OSF_Zisseki(Siyou(i), CrySmp(i), "2", Zisseki.OSF2Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J005"
'''''        If OSF_Zisseki(Siyou(i), CrySmp(i), "3", Zisseki.OSF3Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J005"
'''''        If OSF_Zisseki(Siyou(i), CrySmp(i), "4", Zisseki.OSF4Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J006"
'''''        If GD_Zisseki(Siyou(i), CrySmp(i), Zisseki.GDZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J007"
'''''        If LT_Zisseki(Siyou(i), CrySmp(i), Zisseki.LTZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''        sDbName = "J001"
'''''        If EPD_Zisseki(Siyou(i), CrySmp(i), Zisseki.EPDZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'''''    Next
'''''    sDbName = "J004"
'''''    If CS_Zisseki(Siyou(1).CRYNUM, CrySmp, Zisseki.csz) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit

    sDbName = ""
    DBDRV_scmzc_fcmkc001c_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    If Trim$(sDbName) <> "" Then sErrMsg = GetMsgStr("EGET2", sDbName)
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_scmzc_fcmkc001c_Disp = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function



'概要      :総合判定 表示用ＤＢドライバ（ブロックＩＤが米沢以外の場合）
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                ,説明
'          :inBlockID     ,I  ,String                            ,対象ブロックID
'          :Siyou()       ,O  ,type_DBDRV_scmzc_fcmkc001c_Siyou  ,品番、仕様、結晶内側取得用
'          :CrySmp()      ,O  ,type_DBDRV_scmzc_fcmkc001c_CrySmp ,結晶サンプル管理取得用
'          :Zisseki       ,O  ,typ_TBCMG002                      ,購入単結晶受入実績用
'          :sErrMsg       ,O  ,String                            ,
'          :戻り値        ,O  ,FUNCTION_RETURN                   ,読み込み成否
'説明      :
'履歴      :2001/06/28 蔵本 作成
'''''Public Function DBDRV_scmzc_fcmkc001c_Disp2(inBlockID As String, _
'''''                                           Siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                                           CrySmp() As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
'''''                                           Zisseki As typ_TBCMG002, _
'''''                                           sErrMsg As String) As FUNCTION_RETURN
Public Function DBDRV_scmzc_fcmkc001c_Disp2(inBlockID As String, _
                                           Siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                                           sErrMsg As String) As FUNCTION_RETURN

    Dim sDbName As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_Disp2"

    DBDRV_scmzc_fcmkc001c_Disp2 = FUNCTION_RETURN_SUCCESS

    sDbName = "V011"
    '品番、SXL仕様からデータの取得（レコード0件時エラー）
    If getHinSiyou30(inBlockID, Siyou()) = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmkc001c_Disp2 = FUNCTION_RETURN_FAILURE
        sErrMsg = GetMsgStr("EGET2", sDbName)
        GoTo proc_exit
    End If

'    sDBName = "V010"
'    '結晶サンプルの取得（レコード0件時エラー）
'    If getCrySmp(Siyou(1).Crynum, Siyou(1).IngotPos, CrySmp()) = FUNCTION_RETURN_FAILURE Then
'        DBDRV_scmzc_fcmkc001c_Disp2 = FUNCTION_RETURN_FAILURE
'        sErrMsg = GetMsgStr("EGET2", sDBName)
'        GoTo proc_exit
'    End If

'''''    sDbName = "G002"
'''''    '購入単結晶受入実績取得（レコード0件時エラー）
'''''    If Kounyu_Zisseki(Zisseki, " where CRYNUM = '" & inBlockID & _
'''''                               "' and TRANCNT=ANY(select MAX(TRANCNT) from TBCMG002 " & _
'''''                               " where CRYNUM = '" & inBlockID & "' )") = FUNCTION_RETURN_FAILURE Then
'''''        DBDRV_scmzc_fcmkc001c_Disp2 = FUNCTION_RETURN_FAILURE
'''''        GoTo proc_exit
'''''    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_scmzc_fcmkc001c_Disp2 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


''''''概要      :共用の場合には、既存レコードから必要項目を取得する
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
''''''          :Soku          ,IO ,typ_TBCMJ014 ,測定評価実績データ
''''''説明      :実績が記録されていない項目について、既存の共用レコードに値があればそれを採用する
''''''履歴      :2002/4/26 野村 作成
'''''Private Sub UpdateFromOrgJ014(Soku As typ_TBCMJ014)
'''''Dim sql As String
'''''Dim rs As OraDynaset
'''''
'''''    With Soku
'''''        '既存の共用レコードから値を取得する
'''''        '「共用レコード」の判断は、その位置に総合判定測定値が1レコードだけあることを条件とする
'''''        '※ T/Bの両側があるときは、既存の値を取らなくても Sokuに入っているものを採用してよい
'''''        '※ 共用サンプルでも、既存レコードがなければ問題ない
'''''        sql = sql & "select POSITION,SMPKBN"
'''''        sql = sql & ",SXL_RS_SMPPOS,SXL_OI_SMPPOS,SXL_CS_SMPPOS,SXLOSF1_SMPPOS,SXLBMD_SMPPOS,SXLGD_SMPPOS,SXLT_SMPPOS"
'''''        sql = sql & ",SXLRS_MEAS1,SXLRS_MEAS2,SXLRS_MEAS3,SXLRS_MEAS4,SXLRS_MEAS5,SXLRS_EFEHS,SXLRS_RRG"
'''''        sql = sql & ",SXLOI_OIMEAS1,SXLOI_OIMEAS2,SXLOI_OIMEAS3,SXLOI_OIMEAS4,SXLOI_OIMEAS5,SXLOI_ORGRES,SXLOI_INSPECTWAY"
'''''        sql = sql & ",SXLCS_CSMEAS,SXLCS_70PPRE"
'''''        sql = sql & ",SXLOSF1_KKSP,SXLOSF1_NETU,SXLOSF1_KKSET,SXLOSF1_MEAS1,SXLOSF1_MEAS2,SXLOSF1_MEAS3,SXLOSF1_MEAS4,SXLOSF1_MEAS5,SXLOSF1_MEAS6,SXLOSF1_MEAS7,SXLOSF1_MEAS8,SXLOSF1_MEAS9,SXLOSF1_MEAS10"
'''''        sql = sql & ",SXLOSF1_MEAS11,SXLOSF1_MEAS12,SXLOSF1_MEAS13,SXLOSF1_MEAS14,SXLOSF1_MEAS15,SXLOSF1_MEAS16,SXLOSF1_MEAS17,SXLOSF1_MEAS18,SXLOSF1_MEAS19,SXLOSF1_MEAS20,SXLOSF1_CALCMAX,SXLOSF1_CALCAVE"
'''''        sql = sql & ",SXLOSF2_KKSP,SXLOSF2_NETU,SXLOSF2_KKSET,SXLOSF2_MEAS1,SXLOSF2_MEAS2,SXLOSF2_MEAS3,SXLOSF2_MEAS4,SXLOSF2_MEAS5,SXLOSF2_MEAS6,SXLOSF2_MEAS7,SXLOSF2_MEAS8,SXLOSF2_MEAS9,SXLOSF2_MEAS10"
'''''        sql = sql & ",SXLOSF2_MEAS11,SXLOSF2_MEAS12,SXLOSF2_MEAS13,SXLOSF2_MEAS14,SXLOSF2_MEAS15,SXLOSF2_MEAS16,SXLOSF2_MEAS17,SXLOSF2_MEAS18,SXLOSF2_MEAS19,SXLOSF2_MEAS20,SXLOSF2_CALCMAX,SXLOSF2_CALCAVE"
'''''        sql = sql & ",SXLOSF3_KKSP,SXLOSF3_NETU,SXLOSF3_KKSET,SXLOSF3_MEAS1,SXLOSF3_MEAS2,SXLOSF3_MEAS3,SXLOSF3_MEAS4,SXLOSF3_MEAS5,SXLOSF3_MEAS6,SXLOSF3_MEAS7,SXLOSF3_MEAS8,SXLOSF3_MEAS9,SXLOSF3_MEAS10"
'''''        sql = sql & ",SXLOSF3_MEAS11,SXLOSF3_MEAS12,SXLOSF3_MEAS13,SXLOSF3_MEAS14,SXLOSF3_MEAS15,SXLOSF3_MEAS16,SXLOSF3_MEAS17,SXLOSF3_MEAS18,SXLOSF3_MEAS19,SXLOSF3_MEAS20,SXLOSF3_CALCMAX,SXLOSF3_CALCAVE"
'''''        sql = sql & ",SXLOSF4_KKSP,SXLOSF4_NETU,SXLOSF4_KKSET,SXLOSF4_MEAS1,SXLOSF4_MEAS2,SXLOSF4_MEAS3,SXLOSF4_MEAS4,SXLOSF4_MEAS5,SXLOSF4_MEAS6,SXLOSF4_MEAS7,SXLOSF4_MEAS8,SXLOSF4_MEAS9,SXLOSF4_MEAS10"
'''''        sql = sql & ",SXLOSF4_MEAS11,SXLOSF4_MEAS12,SXLOSF4_MEAS13,SXLOSF4_MEAS14,SXLOSF4_MEAS15,SXLOSF4_MEAS16,SXLOSF4_MEAS17,SXLOSF4_MEAS18,SXLOSF4_MEAS19,SXLOSF4_MEAS20,SXLOSF4_CALCMAX,SXLOSF4_CALCAVE"
'''''        sql = sql & ",SXLBMD1_KKSP,SXLBMD1_NETU,SXLBMD1_KKSET,SXLBMD1_MEAS1,SXLBMD1_MEAS2,SXLBMD1_MEAS3,SXLBMD1_MEAS4,SXLBMD1_MEAS5,SXLBMD1_CALCMAX,SXLBMD1_CALCAVE"
'''''        sql = sql & ",SXLBMD2_KKSP,SXLBMD2_NETU,SXLBMD2_KKSET,SXLBMD2_MEAS1,SXLBMD2_MEAS2,SXLBMD2_MEAS3,SXLBMD2_MEAS4,SXLBMD2_MEAS5,SXLBMD2_CALCMAX,SXLBMD2_CALCAVE"
'''''        sql = sql & ",SXLBMD3_KKSP,SXLBMD3_NETU,SXLBMD3_KKSET,SXLBMD3_MEAS1,SXLBMD3_MEAS2,SXLBMD3_MEAS3,SXLBMD3_MEAS4,SXLBMD3_MEAS5,SXLBMD3_CALCMAX,SXLBMD3_CALCAVE"
'''''        sql = sql & ",SXLGD_MS01LDL1,SXLGD_MS01LDL2,SXLGD_MS01LDL3,SXLGD_MS01LDL4,SXLGD_MS01LDL5,SXLGD_MS01DEN1,SXLGD_MS01DEN2,SXLGD_MS01DEN3,SXLGD_MS01DEN4,SXLGD_MS01DEN5"
'''''        sql = sql & ",SXLGD_MS02LDL1,SXLGD_MS02LDL2,SXLGD_MS02LDL3,SXLGD_MS02LDL4,SXLGD_MS02LDL5,SXLGD_MS02DEN1,SXLGD_MS02DEN2,SXLGD_MS02DEN3,SXLGD_MS02DEN4,SXLGD_MS02DEN5"
'''''        sql = sql & ",SXLGD_MS03LDL1,SXLGD_MS03LDL2,SXLGD_MS03LDL3,SXLGD_MS03LDL4,SXLGD_MS03LDL5,SXLGD_MS03DEN1,SXLGD_MS03DEN2,SXLGD_MS03DEN3,SXLGD_MS03DEN4,SXLGD_MS03DEN5"
'''''        sql = sql & ",SXLGD_MS04LDL1,SXLGD_MS04LDL2,SXLGD_MS04LDL3,SXLGD_MS04LDL4,SXLGD_MS04LDL5,SXLGD_MS04DEN1,SXLGD_MS04DEN2,SXLGD_MS04DEN3,SXLGD_MS04DEN4,SXLGD_MS04DEN5"
'''''        sql = sql & ",SXLGD_MS05LDL1,SXLGD_MS05LDL2,SXLGD_MS05LDL3,SXLGD_MS05LDL4,SXLGD_MS05LDL5,SXLGD_MS05DEN1,SXLGD_MS05DEN2,SXLGD_MS05DEN3,SXLGD_MS05DEN4,SXLGD_MS05DEN5"
'''''        sql = sql & ",SXLGD_MS06LDL1,SXLGD_MS06LDL2,SXLGD_MS06LDL3,SXLGD_MS06LDL4,SXLGD_MS06LDL5,SXLGD_MS06DEN1,SXLGD_MS06DEN2,SXLGD_MS06DEN3,SXLGD_MS06DEN4,SXLGD_MS06DEN5"
'''''        sql = sql & ",SXLGD_MS07LDL1,SXLGD_MS07LDL2,SXLGD_MS07LDL3,SXLGD_MS07LDL4,SXLGD_MS07LDL5,SXLGD_MS07DEN1,SXLGD_MS07DEN2,SXLGD_MS07DEN3,SXLGD_MS07DEN4,SXLGD_MS07DEN5"
'''''        sql = sql & ",SXLGD_MS08LDL1,SXLGD_MS08LDL2,SXLGD_MS08LDL3,SXLGD_MS08LDL4,SXLGD_MS08LDL5,SXLGD_MS08DEN1,SXLGD_MS08DEN2,SXLGD_MS08DEN3,SXLGD_MS08DEN4,SXLGD_MS08DEN5"
'''''        sql = sql & ",SXLGD_MS09LDL1,SXLGD_MS09LDL2,SXLGD_MS09LDL3,SXLGD_MS09LDL4,SXLGD_MS09LDL5,SXLGD_MS09DEN1,SXLGD_MS09DEN2,SXLGD_MS09DEN3,SXLGD_MS09DEN4,SXLGD_MS09DEN5"
'''''        sql = sql & ",SXLGD_MS10LDL1,SXLGD_MS10LDL2,SXLGD_MS10LDL3,SXLGD_MS10LDL4,SXLGD_MS10LDL5,SXLGD_MS10DEN1,SXLGD_MS10DEN2,SXLGD_MS10DEN3,SXLGD_MS10DEN4,SXLGD_MS10DEN5"
'''''        sql = sql & ",SXLGD_MS11LDL1,SXLGD_MS11LDL2,SXLGD_MS11LDL3,SXLGD_MS11LDL4,SXLGD_MS11LDL5,SXLGD_MS11DEN1,SXLGD_MS11DEN2,SXLGD_MS11DEN3,SXLGD_MS11DEN4,SXLGD_MS11DEN5"
'''''        sql = sql & ",SXLGD_MS12LDL1,SXLGD_MS12LDL2,SXLGD_MS12LDL3,SXLGD_MS12LDL4,SXLGD_MS12LDL5,SXLGD_MS12DEN1,SXLGD_MS12DEN2,SXLGD_MS12DEN3,SXLGD_MS12DEN4,SXLGD_MS12DEN5"
'''''        sql = sql & ",SXLGD_MS13LDL1,SXLGD_MS13LDL2,SXLGD_MS13LDL3,SXLGD_MS13LDL4,SXLGD_MS13LDL5,SXLGD_MS13DEN1,SXLGD_MS13DEN2,SXLGD_MS13DEN3,SXLGD_MS13DEN4,SXLGD_MS13DEN5"
'''''        sql = sql & ",SXLGD_MS14LDL1,SXLGD_MS14LDL2,SXLGD_MS14LDL3,SXLGD_MS14LDL4,SXLGD_MS14LDL5,SXLGD_MS14DEN1,SXLGD_MS14DEN2,SXLGD_MS14DEN3,SXLGD_MS14DEN4,SXLGD_MS14DEN5"
'''''        sql = sql & ",SXLGD_MS15LDL1,SXLGD_MS15LDL2,SXLGD_MS15LDL3,SXLGD_MS15LDL4,SXLGD_MS15LDL5,SXLGD_MS15DEN1,SXLGD_MS15DEN2,SXLGD_MS15DEN3,SXLGD_MS15DEN4,SXLGD_MS15DEN5"
'''''        sql = sql & ",SXLGD_MSRSDEN,SXLGD_MSRSLDL,SXLGD_MSRSDVD2"
'''''        sql = sql & ",SXLLT_MEASPEAK,SXLLT_MEAS1,SXLLT_MEAS2,SXLLT_MEAS3,SXLLT_MEAS4,SXLLT_MEAS5,SXLLT_CALCMEAS"
'''''        sql = sql & ",SXLOSF1_POS1,SXLOSF1_WID1,SXLOSF1_RD1"
'''''        sql = sql & ",SXLOSF1_POS2,SXLOSF1_WID2,SXLOSF1_RD2"
'''''        sql = sql & ",SXLOSF1_POS3,SXLOSF1_WID3,SXLOSF1_RD3"
'''''        sql = sql & ",SXLOSF2_POS1,SXLOSF2_WID1,SXLOSF2_RD1"
'''''        sql = sql & ",SXLOSF2_POS2,SXLOSF2_WID2,SXLOSF2_RD2"
'''''        sql = sql & ",SXLOSF2_POS3,SXLOSF2_WID3,SXLOSF2_RD3"
'''''        sql = sql & ",SXLOSF3_POS1,SXLOSF3_WID1,SXLOSF3_RD1"
'''''        sql = sql & ",SXLOSF3_POS2,SXLOSF3_WID2,SXLOSF3_RD2"
'''''        sql = sql & ",SXLOSF3_POS3,SXLOSF3_WID3,SXLOSF3_RD3"
'''''        sql = sql & ",SXLOSF4_POS1,SXLOSF4_WID1,SXLOSF4_RD1"
'''''        sql = sql & ",SXLOSF4_POS2,SXLOSF4_WID2,SXLOSF4_RD2"
'''''        sql = sql & ",SXLOSF4_POS3,SXLOSF4_WID3,SXLOSF4_RD3"
'''''        sql = sql & ",SXLGD_MS01DVD2,SXLGD_MS02DVD2,SXLGD_MS03DVD2,SXLGD_MS04DVD2,SXLGD_MS05DVD2"
'''''        sql = sql & ",SXLBMD1_MNBCR,SXLBMD2_MNBCR,SXLBMD3_MNBCR"
'''''            'ここまでのSQLは "select *" の方が適切だと思う (nomura:必要カラムがほぼ全カラムのため)
'''''        sql = sql & " from TBCMJ014"
'''''        sql = sql & " where CRYNUM='" & .CRYNUM & "' and POSITION=" & .POSITION & " and SMPKBN='" & .SMPKBN & "'"
'''''        sql = sql & " and 1=(select count(*) from TBCMJ014 where CRYNUM='" & .CRYNUM & "' and POSITION=" & .POSITION & ")"
'''''        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''        If rs.RecordCount > 0 Then
'''''            '検査項目毎に、実績をコピーする (既存レコードに値があって、現ブロックに実績がない場合)
'''''            If (.SXL_RS_SMPPOS = -1) And (rs("SXL_RS_SMPPOS") <> -1) Then
'''''                'Rs実績をコピー
'''''                .SXL_RS_SMPPOS = rs("SXL_RS_SMPPOS")
'''''                .SXLRS_MEAS1 = rs("SXLRS_MEAS1")
'''''                .SXLRS_MEAS2 = rs("SXLRS_MEAS2")
'''''                .SXLRS_MEAS3 = rs("SXLRS_MEAS3")
'''''                .SXLRS_MEAS4 = rs("SXLRS_MEAS4")
'''''                .SXLRS_MEAS5 = rs("SXLRS_MEAS5")
'''''                .SXLRS_EFEHS = rs("SXLRS_EFEHS")
'''''                .SXLRS_RRG = rs("SXLRS_RRG")
'''''            End If
'''''            If (.SXL_OI_SMPPOS = -1) And (rs("SXL_OI_SMPPOS") <> -1) Then
'''''                'Oi実績をコピー
'''''                .SXL_OI_SMPPOS = rs("SXL_OI_SMPPOS")
'''''                .SXLOI_OIMEAS1 = rs("SXLOI_OIMEAS1")
'''''                .SXLOI_OIMEAS2 = rs("SXLOI_OIMEAS2")
'''''                .SXLOI_OIMEAS3 = rs("SXLOI_OIMEAS3")
'''''                .SXLOI_OIMEAS4 = rs("SXLOI_OIMEAS4")
'''''                .SXLOI_OIMEAS5 = rs("SXLOI_OIMEAS5")
'''''                .SXLOI_ORGRES = rs("SXLOI_ORGRES")
'''''                .SXLOI_INSPECTWAY = rs("SXLOI_INSPECTWAY")
'''''            End If
'''''            If (.SXLCS_CSMEAS = -1) And (rs("SXLCS_CSMEAS") <> -1) Then
'''''                'Cs実績をコピー
'''''                .SXL_CS_SMPPOS = rs("SXL_CS_SMPPOS")
'''''                .SXLCS_CSMEAS = rs("SXLCS_CSMEAS")
'''''                .SXLCS_70PPRE = rs("SXLCS_70PPRE")
'''''            End If
'''''            If (.SXLOSF1_SMPPOS = -1) And (rs("SXLOSF1_SMPPOS") <> -1) Then
'''''                'OSF実績をコピー
'''''                .SXLOSF1_SMPPOS = rs("SXLOSF1_SMPPOS")
'''''                .SXLOSF1_KKSP = rs("SXLOSF1_KKSP")
'''''                .SXLOSF1_NETU = rs("SXLOSF1_NETU")
'''''                .SXLOSF1_KKSET = rs("SXLOSF1_KKSET")
'''''                .SXLOSF1_MEAS1 = rs("SXLOSF1_MEAS1")
'''''                .SXLOSF1_MEAS2 = rs("SXLOSF1_MEAS2")
'''''                .SXLOSF1_MEAS3 = rs("SXLOSF1_MEAS3")
'''''                .SXLOSF1_MEAS4 = rs("SXLOSF1_MEAS4")
'''''                .SXLOSF1_MEAS5 = rs("SXLOSF1_MEAS5")
'''''                .SXLOSF1_MEAS6 = rs("SXLOSF1_MEAS6")
'''''                .SXLOSF1_MEAS7 = rs("SXLOSF1_MEAS7")
'''''                .SXLOSF1_MEAS8 = rs("SXLOSF1_MEAS8")
'''''                .SXLOSF1_MEAS9 = rs("SXLOSF1_MEAS9")
'''''                .SXLOSF1_MEAS10 = rs("SXLOSF1_MEAS10")
'''''                .SXLOSF1_MEAS11 = rs("SXLOSF1_MEAS11")
'''''                .SXLOSF1_MEAS12 = rs("SXLOSF1_MEAS12")
'''''                .SXLOSF1_MEAS13 = rs("SXLOSF1_MEAS13")
'''''                .SXLOSF1_MEAS14 = rs("SXLOSF1_MEAS14")
'''''                .SXLOSF1_MEAS15 = rs("SXLOSF1_MEAS15")
'''''                .SXLOSF1_MEAS16 = rs("SXLOSF1_MEAS16")
'''''                .SXLOSF1_MEAS17 = rs("SXLOSF1_MEAS17")
'''''                .SXLOSF1_MEAS18 = rs("SXLOSF1_MEAS18")
'''''                .SXLOSF1_MEAS19 = rs("SXLOSF1_MEAS19")
'''''                .SXLOSF1_MEAS20 = rs("SXLOSF1_MEAS20")
'''''                .SXLOSF1_CALCMAX = rs("SXLOSF1_CALCMAX")
'''''                .SXLOSF1_CALCAVE = rs("SXLOSF1_CALCAVE")
'''''                .SXLOSF2_KKSP = rs("SXLOSF2_KKSP")
'''''                .SXLOSF2_NETU = rs("SXLOSF2_NETU")
'''''                .SXLOSF2_KKSET = rs("SXLOSF2_KKSET")
'''''                .SXLOSF2_MEAS1 = rs("SXLOSF2_MEAS1")
'''''                .SXLOSF2_MEAS2 = rs("SXLOSF2_MEAS2")
'''''                .SXLOSF2_MEAS3 = rs("SXLOSF2_MEAS3")
'''''                .SXLOSF2_MEAS4 = rs("SXLOSF2_MEAS4")
'''''                .SXLOSF2_MEAS5 = rs("SXLOSF2_MEAS5")
'''''                .SXLOSF2_MEAS6 = rs("SXLOSF2_MEAS6")
'''''                .SXLOSF2_MEAS7 = rs("SXLOSF2_MEAS7")
'''''                .SXLOSF2_MEAS8 = rs("SXLOSF2_MEAS8")
'''''                .SXLOSF2_MEAS9 = rs("SXLOSF2_MEAS9")
'''''                .SXLOSF2_MEAS10 = rs("SXLOSF2_MEAS10")
'''''                .SXLOSF2_MEAS11 = rs("SXLOSF2_MEAS11")
'''''                .SXLOSF2_MEAS12 = rs("SXLOSF2_MEAS12")
'''''                .SXLOSF2_MEAS13 = rs("SXLOSF2_MEAS13")
'''''                .SXLOSF2_MEAS14 = rs("SXLOSF2_MEAS14")
'''''                .SXLOSF2_MEAS15 = rs("SXLOSF2_MEAS15")
'''''                .SXLOSF2_MEAS16 = rs("SXLOSF2_MEAS16")
'''''                .SXLOSF2_MEAS17 = rs("SXLOSF2_MEAS17")
'''''                .SXLOSF2_MEAS18 = rs("SXLOSF2_MEAS18")
'''''                .SXLOSF2_MEAS19 = rs("SXLOSF2_MEAS19")
'''''                .SXLOSF2_MEAS20 = rs("SXLOSF2_MEAS20")
'''''                .SXLOSF2_CALCMAX = rs("SXLOSF2_CALCMAX")
'''''                .SXLOSF2_CALCAVE = rs("SXLOSF2_CALCAVE")
'''''                .SXLOSF3_KKSP = rs("SXLOSF3_KKSP")
'''''                .SXLOSF3_NETU = rs("SXLOSF3_NETU")
'''''                .SXLOSF3_KKSET = rs("SXLOSF3_KKSET")
'''''                .SXLOSF3_MEAS1 = rs("SXLOSF3_MEAS1")
'''''                .SXLOSF3_MEAS2 = rs("SXLOSF3_MEAS2")
'''''                .SXLOSF3_MEAS3 = rs("SXLOSF3_MEAS3")
'''''                .SXLOSF3_MEAS4 = rs("SXLOSF3_MEAS4")
'''''                .SXLOSF3_MEAS5 = rs("SXLOSF3_MEAS5")
'''''                .SXLOSF3_MEAS6 = rs("SXLOSF3_MEAS6")
'''''                .SXLOSF3_MEAS7 = rs("SXLOSF3_MEAS7")
'''''                .SXLOSF3_MEAS8 = rs("SXLOSF3_MEAS8")
'''''                .SXLOSF3_MEAS9 = rs("SXLOSF3_MEAS9")
'''''                .SXLOSF3_MEAS10 = rs("SXLOSF3_MEAS10")
'''''                .SXLOSF3_MEAS11 = rs("SXLOSF3_MEAS11")
'''''                .SXLOSF3_MEAS12 = rs("SXLOSF3_MEAS12")
'''''                .SXLOSF3_MEAS13 = rs("SXLOSF3_MEAS13")
'''''                .SXLOSF3_MEAS14 = rs("SXLOSF3_MEAS14")
'''''                .SXLOSF3_MEAS15 = rs("SXLOSF3_MEAS15")
'''''                .SXLOSF3_MEAS16 = rs("SXLOSF3_MEAS16")
'''''                .SXLOSF3_MEAS17 = rs("SXLOSF3_MEAS17")
'''''                .SXLOSF3_MEAS18 = rs("SXLOSF3_MEAS18")
'''''                .SXLOSF3_MEAS19 = rs("SXLOSF3_MEAS19")
'''''                .SXLOSF3_MEAS20 = rs("SXLOSF3_MEAS20")
'''''                .SXLOSF3_CALCMAX = rs("SXLOSF3_CALCMAX")
'''''                .SXLOSF3_CALCAVE = rs("SXLOSF3_CALCAVE")
'''''                .SXLOSF4_KKSP = rs("SXLOSF4_KKSP")
'''''                .SXLOSF4_NETU = rs("SXLOSF4_NETU")
'''''                .SXLOSF4_KKSET = rs("SXLOSF4_KKSET")
'''''                .SXLOSF4_MEAS1 = rs("SXLOSF4_MEAS1")
'''''                .SXLOSF4_MEAS2 = rs("SXLOSF4_MEAS2")
'''''                .SXLOSF4_MEAS3 = rs("SXLOSF4_MEAS3")
'''''                .SXLOSF4_MEAS4 = rs("SXLOSF4_MEAS4")
'''''                .SXLOSF4_MEAS5 = rs("SXLOSF4_MEAS5")
'''''                .SXLOSF4_MEAS6 = rs("SXLOSF4_MEAS6")
'''''                .SXLOSF4_MEAS7 = rs("SXLOSF4_MEAS7")
'''''                .SXLOSF4_MEAS8 = rs("SXLOSF4_MEAS8")
'''''                .SXLOSF4_MEAS9 = rs("SXLOSF4_MEAS9")
'''''                .SXLOSF4_MEAS10 = rs("SXLOSF4_MEAS10")
'''''                .SXLOSF4_MEAS11 = rs("SXLOSF4_MEAS11")
'''''                .SXLOSF4_MEAS12 = rs("SXLOSF4_MEAS12")
'''''                .SXLOSF4_MEAS13 = rs("SXLOSF4_MEAS13")
'''''                .SXLOSF4_MEAS14 = rs("SXLOSF4_MEAS14")
'''''                .SXLOSF4_MEAS15 = rs("SXLOSF4_MEAS15")
'''''                .SXLOSF4_MEAS16 = rs("SXLOSF4_MEAS16")
'''''                .SXLOSF4_MEAS17 = rs("SXLOSF4_MEAS17")
'''''                .SXLOSF4_MEAS18 = rs("SXLOSF4_MEAS18")
'''''                .SXLOSF4_MEAS19 = rs("SXLOSF4_MEAS19")
'''''                .SXLOSF4_MEAS20 = rs("SXLOSF4_MEAS20")
'''''                .SXLOSF4_CALCMAX = rs("SXLOSF4_CALCMAX")
'''''                .SXLOSF4_CALCAVE = rs("SXLOSF4_CALCAVE")
'''''                If IsNull(rs("SXLOSF1_POS1")) = False Then .SXLOSF1_POS1 = rs("SXLOSF1_POS1")
'''''                If IsNull(rs("SXLOSF1_WID1")) = False Then .SXLOSF1_WID1 = rs("SXLOSF1_WID1")
'''''                If IsNull(rs("SXLOSF1_RD1")) = False Then .SXLOSF1_RD1 = rs("SXLOSF1_RD1")
'''''                If IsNull(rs("SXLOSF1_POS2")) = False Then .SXLOSF1_POS2 = rs("SXLOSF1_POS2")
'''''                If IsNull(rs("SXLOSF1_WID2")) = False Then .SXLOSF1_WID2 = rs("SXLOSF1_WID2")
'''''                If IsNull(rs("SXLOSF1_RD2")) = False Then .SXLOSF1_RD2 = rs("SXLOSF1_RD2")
'''''                If IsNull(rs("SXLOSF1_POS3")) = False Then .SXLOSF1_POS3 = rs("SXLOSF1_POS3")
'''''                If IsNull(rs("SXLOSF1_WID3")) = False Then .SXLOSF1_WID3 = rs("SXLOSF1_WID3")
'''''                If IsNull(rs("SXLOSF1_RD3")) = False Then .SXLOSF1_RD3 = rs("SXLOSF1_RD3")
'''''                If IsNull(rs("SXLOSF2_POS1")) = False Then .SXLOSF2_POS1 = rs("SXLOSF2_POS1")
'''''                If IsNull(rs("SXLOSF2_WID1")) = False Then .SXLOSF2_WID1 = rs("SXLOSF2_WID1")
'''''                If IsNull(rs("SXLOSF2_RD1")) = False Then .SXLOSF2_RD1 = rs("SXLOSF2_RD1")
'''''                If IsNull(rs("SXLOSF2_POS2")) = False Then .SXLOSF2_POS2 = rs("SXLOSF2_POS2")
'''''                If IsNull(rs("SXLOSF2_WID2")) = False Then .SXLOSF2_WID2 = rs("SXLOSF2_WID2")
'''''                If IsNull(rs("SXLOSF2_RD2")) = False Then .SXLOSF2_RD2 = rs("SXLOSF2_RD2")
'''''                If IsNull(rs("SXLOSF2_POS3")) = False Then .SXLOSF2_POS3 = rs("SXLOSF2_POS3")
'''''                If IsNull(rs("SXLOSF2_WID3")) = False Then .SXLOSF2_WID3 = rs("SXLOSF2_WID3")
'''''                If IsNull(rs("SXLOSF2_RD3")) = False Then .SXLOSF2_RD3 = rs("SXLOSF2_RD3")
'''''                If IsNull(rs("SXLOSF3_POS1")) = False Then .SXLOSF3_POS1 = rs("SXLOSF3_POS1")
'''''                If IsNull(rs("SXLOSF3_WID1")) = False Then .SXLOSF3_WID1 = rs("SXLOSF3_WID1")
'''''                If IsNull(rs("SXLOSF3_RD1")) = False Then .SXLOSF3_RD1 = rs("SXLOSF3_RD1")
'''''                If IsNull(rs("SXLOSF3_POS2")) = False Then .SXLOSF3_POS2 = rs("SXLOSF3_POS2")
'''''                If IsNull(rs("SXLOSF3_WID2")) = False Then .SXLOSF3_WID2 = rs("SXLOSF3_WID2")
'''''                If IsNull(rs("SXLOSF3_RD2")) = False Then .SXLOSF3_RD2 = rs("SXLOSF3_RD2")
'''''                If IsNull(rs("SXLOSF3_POS3")) = False Then .SXLOSF3_POS3 = rs("SXLOSF3_POS3")
'''''                If IsNull(rs("SXLOSF3_WID3")) = False Then .SXLOSF3_WID3 = rs("SXLOSF3_WID3")
'''''                If IsNull(rs("SXLOSF3_RD3")) = False Then .SXLOSF3_RD3 = rs("SXLOSF3_RD3")
'''''                If IsNull(rs("SXLOSF4_POS1")) = False Then .SXLOSF4_POS1 = rs("SXLOSF4_POS1")
'''''                If IsNull(rs("SXLOSF4_WID1")) = False Then .SXLOSF4_WID1 = rs("SXLOSF4_WID1")
'''''                If IsNull(rs("SXLOSF4_RD1")) = False Then .SXLOSF4_RD1 = rs("SXLOSF4_RD1")
'''''                If IsNull(rs("SXLOSF4_POS2")) = False Then .SXLOSF4_POS2 = rs("SXLOSF4_POS2")
'''''                If IsNull(rs("SXLOSF4_WID2")) = False Then .SXLOSF4_WID2 = rs("SXLOSF4_WID2")
'''''                If IsNull(rs("SXLOSF4_RD2")) = False Then .SXLOSF4_RD2 = rs("SXLOSF4_RD2")
'''''                If IsNull(rs("SXLOSF4_POS3")) = False Then .SXLOSF4_POS3 = rs("SXLOSF4_POS3")
'''''                If IsNull(rs("SXLOSF4_WID3")) = False Then .SXLOSF4_WID3 = rs("SXLOSF4_WID3")
'''''                If IsNull(rs("SXLOSF4_RD3")) = False Then .SXLOSF4_RD3 = rs("SXLOSF4_RD3")
'''''            End If
'''''            If (.SXLBMD_SMPPOS = -1) And (rs("SXLBMD_SMPPOS") <> -1) Then
'''''                'BMD実績をコピー
'''''                .SXLBMD_SMPPOS = rs("SXLBMD_SMPPOS")
'''''                .SXLBMD1_KKSP = rs("SXLBMD1_KKSP")
'''''                .SXLBMD1_NETU = rs("SXLBMD1_NETU")
'''''                .SXLBMD1_KKSET = rs("SXLBMD1_KKSET")
'''''                .SXLBMD1_MEAS1 = rs("SXLBMD1_MEAS1")
'''''                .SXLBMD1_MEAS2 = rs("SXLBMD1_MEAS2")
'''''                .SXLBMD1_MEAS3 = rs("SXLBMD1_MEAS3")
'''''                .SXLBMD1_MEAS4 = rs("SXLBMD1_MEAS4")
'''''                .SXLBMD1_MEAS5 = rs("SXLBMD1_MEAS5")
'''''                .SXLBMD1_CALCMAX = rs("SXLBMD1_CALCMAX")
'''''                .SXLBMD1_CALCAVE = rs("SXLBMD1_CALCAVE")
'''''                .SXLBMD2_KKSP = rs("SXLBMD2_KKSP")
'''''                .SXLBMD2_NETU = rs("SXLBMD2_NETU")
'''''                .SXLBMD2_KKSET = rs("SXLBMD2_KKSET")
'''''                .SXLBMD2_MEAS1 = rs("SXLBMD2_MEAS1")
'''''                .SXLBMD2_MEAS2 = rs("SXLBMD2_MEAS2")
'''''                .SXLBMD2_MEAS3 = rs("SXLBMD2_MEAS3")
'''''                .SXLBMD2_MEAS4 = rs("SXLBMD2_MEAS4")
'''''                .SXLBMD2_MEAS5 = rs("SXLBMD2_MEAS5")
'''''                .SXLBMD2_CALCMAX = rs("SXLBMD2_CALCMAX")
'''''                .SXLBMD2_CALCAVE = rs("SXLBMD2_CALCAVE")
'''''                .SXLBMD3_KKSP = rs("SXLBMD3_KKSP")
'''''                .SXLBMD3_NETU = rs("SXLBMD3_NETU")
'''''                .SXLBMD3_KKSET = rs("SXLBMD3_KKSET")
'''''                .SXLBMD3_MEAS1 = rs("SXLBMD3_MEAS1")
'''''                .SXLBMD3_MEAS2 = rs("SXLBMD3_MEAS2")
'''''                .SXLBMD3_MEAS3 = rs("SXLBMD3_MEAS3")
'''''                .SXLBMD3_MEAS4 = rs("SXLBMD3_MEAS4")
'''''                .SXLBMD3_MEAS5 = rs("SXLBMD3_MEAS5")
'''''                .SXLBMD3_CALCMAX = rs("SXLBMD3_CALCMAX")
'''''                .SXLBMD3_CALCAVE = rs("SXLBMD3_CALCAVE")
'''''                If IsNull(rs("SXLBMD1_MNBCR")) = False Then .SXLBMD1_MNBCR = rs("SXLBMD1_MNBCR")
'''''                If IsNull(rs("SXLBMD2_MNBCR")) = False Then .SXLBMD2_MNBCR = rs("SXLBMD2_MNBCR")
'''''                If IsNull(rs("SXLBMD3_MNBCR")) = False Then .SXLBMD3_MNBCR = rs("SXLBMD3_MNBCR")
'''''            End If
'''''            If (.SXLGD_SMPPOS = -1) And (rs("SXLGD_SMPPOS") <> -1) Then
'''''                'GD実績をコピー
'''''                .SXLGD_SMPPOS = rs("SXLGD_SMPPOS")
'''''                .SXLGD_MS01LDL1 = rs("SXLGD_MS01LDL1")
'''''                .SXLGD_MS01LDL2 = rs("SXLGD_MS01LDL2")
'''''                .SXLGD_MS01LDL3 = rs("SXLGD_MS01LDL3")
'''''                .SXLGD_MS01LDL4 = rs("SXLGD_MS01LDL4")
'''''                .SXLGD_MS01LDL5 = rs("SXLGD_MS01LDL5")
'''''                .SXLGD_MS01DEN1 = rs("SXLGD_MS01DEN1")
'''''                .SXLGD_MS01DEN2 = rs("SXLGD_MS01DEN2")
'''''                .SXLGD_MS01DEN3 = rs("SXLGD_MS01DEN3")
'''''                .SXLGD_MS01DEN4 = rs("SXLGD_MS01DEN4")
'''''                .SXLGD_MS01DEN5 = rs("SXLGD_MS01DEN5")
'''''                .SXLGD_MS02LDL1 = rs("SXLGD_MS02LDL1")
'''''                .SXLGD_MS02LDL2 = rs("SXLGD_MS02LDL2")
'''''                .SXLGD_MS02LDL3 = rs("SXLGD_MS02LDL3")
'''''                .SXLGD_MS02LDL4 = rs("SXLGD_MS02LDL4")
'''''                .SXLGD_MS02LDL5 = rs("SXLGD_MS02LDL5")
'''''                .SXLGD_MS02DEN1 = rs("SXLGD_MS02DEN1")
'''''                .SXLGD_MS02DEN2 = rs("SXLGD_MS02DEN2")
'''''                .SXLGD_MS02DEN3 = rs("SXLGD_MS02DEN3")
'''''                .SXLGD_MS02DEN4 = rs("SXLGD_MS02DEN4")
'''''                .SXLGD_MS02DEN5 = rs("SXLGD_MS02DEN5")
'''''                .SXLGD_MS03LDL1 = rs("SXLGD_MS03LDL1")
'''''                .SXLGD_MS03LDL2 = rs("SXLGD_MS03LDL2")
'''''                .SXLGD_MS03LDL3 = rs("SXLGD_MS03LDL3")
'''''                .SXLGD_MS03LDL4 = rs("SXLGD_MS03LDL4")
'''''                .SXLGD_MS03LDL5 = rs("SXLGD_MS03LDL5")
'''''                .SXLGD_MS03DEN1 = rs("SXLGD_MS03DEN1")
'''''                .SXLGD_MS03DEN2 = rs("SXLGD_MS03DEN2")
'''''                .SXLGD_MS03DEN3 = rs("SXLGD_MS03DEN3")
'''''                .SXLGD_MS03DEN4 = rs("SXLGD_MS03DEN4")
'''''                .SXLGD_MS03DEN5 = rs("SXLGD_MS03DEN5")
'''''                .SXLGD_MS04LDL1 = rs("SXLGD_MS04LDL1")
'''''                .SXLGD_MS04LDL2 = rs("SXLGD_MS04LDL2")
'''''                .SXLGD_MS04LDL3 = rs("SXLGD_MS04LDL3")
'''''                .SXLGD_MS04LDL4 = rs("SXLGD_MS04LDL4")
'''''                .SXLGD_MS04LDL5 = rs("SXLGD_MS04LDL5")
'''''                .SXLGD_MS04DEN1 = rs("SXLGD_MS04DEN1")
'''''                .SXLGD_MS04DEN2 = rs("SXLGD_MS04DEN2")
'''''                .SXLGD_MS04DEN3 = rs("SXLGD_MS04DEN3")
'''''                .SXLGD_MS04DEN4 = rs("SXLGD_MS04DEN4")
'''''                .SXLGD_MS04DEN5 = rs("SXLGD_MS04DEN5")
'''''                .SXLGD_MS05LDL1 = rs("SXLGD_MS05LDL1")
'''''                .SXLGD_MS05LDL2 = rs("SXLGD_MS05LDL2")
'''''                .SXLGD_MS05LDL3 = rs("SXLGD_MS05LDL3")
'''''                .SXLGD_MS05LDL4 = rs("SXLGD_MS05LDL4")
'''''                .SXLGD_MS05LDL5 = rs("SXLGD_MS05LDL5")
'''''                .SXLGD_MS05DEN1 = rs("SXLGD_MS05DEN1")
'''''                .SXLGD_MS05DEN2 = rs("SXLGD_MS05DEN2")
'''''                .SXLGD_MS05DEN3 = rs("SXLGD_MS05DEN3")
'''''                .SXLGD_MS05DEN4 = rs("SXLGD_MS05DEN4")
'''''                .SXLGD_MS05DEN5 = rs("SXLGD_MS05DEN5")
'''''                .SXLGD_MS06LDL1 = rs("SXLGD_MS06LDL1")
'''''                .SXLGD_MS06LDL2 = rs("SXLGD_MS06LDL2")
'''''                .SXLGD_MS06LDL3 = rs("SXLGD_MS06LDL3")
'''''                .SXLGD_MS06LDL4 = rs("SXLGD_MS06LDL4")
'''''                .SXLGD_MS06LDL5 = rs("SXLGD_MS06LDL5")
'''''                .SXLGD_MS06DEN1 = rs("SXLGD_MS06DEN1")
'''''                .SXLGD_MS06DEN2 = rs("SXLGD_MS06DEN2")
'''''                .SXLGD_MS06DEN3 = rs("SXLGD_MS06DEN3")
'''''                .SXLGD_MS06DEN4 = rs("SXLGD_MS06DEN4")
'''''                .SXLGD_MS06DEN5 = rs("SXLGD_MS06DEN5")
'''''                .SXLGD_MS07LDL1 = rs("SXLGD_MS07LDL1")
'''''                .SXLGD_MS07LDL2 = rs("SXLGD_MS07LDL2")
'''''                .SXLGD_MS07LDL3 = rs("SXLGD_MS07LDL3")
'''''                .SXLGD_MS07LDL4 = rs("SXLGD_MS07LDL4")
'''''                .SXLGD_MS07LDL5 = rs("SXLGD_MS07LDL5")
'''''                .SXLGD_MS07DEN1 = rs("SXLGD_MS07DEN1")
'''''                .SXLGD_MS07DEN2 = rs("SXLGD_MS07DEN2")
'''''                .SXLGD_MS07DEN3 = rs("SXLGD_MS07DEN3")
'''''                .SXLGD_MS07DEN4 = rs("SXLGD_MS07DEN4")
'''''                .SXLGD_MS07DEN5 = rs("SXLGD_MS07DEN5")
'''''                .SXLGD_MS08LDL1 = rs("SXLGD_MS08LDL1")
'''''                .SXLGD_MS08LDL2 = rs("SXLGD_MS08LDL2")
'''''                .SXLGD_MS08LDL3 = rs("SXLGD_MS08LDL3")
'''''                .SXLGD_MS08LDL4 = rs("SXLGD_MS08LDL4")
'''''                .SXLGD_MS08LDL5 = rs("SXLGD_MS08LDL5")
'''''                .SXLGD_MS08DEN1 = rs("SXLGD_MS08DEN1")
'''''                .SXLGD_MS08DEN2 = rs("SXLGD_MS08DEN2")
'''''                .SXLGD_MS08DEN3 = rs("SXLGD_MS08DEN3")
'''''                .SXLGD_MS08DEN4 = rs("SXLGD_MS08DEN4")
'''''                .SXLGD_MS08DEN5 = rs("SXLGD_MS08DEN5")
'''''                .SXLGD_MS09LDL1 = rs("SXLGD_MS09LDL1")
'''''                .SXLGD_MS09LDL2 = rs("SXLGD_MS09LDL2")
'''''                .SXLGD_MS09LDL3 = rs("SXLGD_MS09LDL3")
'''''                .SXLGD_MS09LDL4 = rs("SXLGD_MS09LDL4")
'''''                .SXLGD_MS09LDL5 = rs("SXLGD_MS09LDL5")
'''''                .SXLGD_MS09DEN1 = rs("SXLGD_MS09DEN1")
'''''                .SXLGD_MS09DEN2 = rs("SXLGD_MS09DEN2")
'''''                .SXLGD_MS09DEN3 = rs("SXLGD_MS09DEN3")
'''''                .SXLGD_MS09DEN4 = rs("SXLGD_MS09DEN4")
'''''                .SXLGD_MS09DEN5 = rs("SXLGD_MS09DEN5")
'''''                .SXLGD_MS10LDL1 = rs("SXLGD_MS10LDL1")
'''''                .SXLGD_MS10LDL2 = rs("SXLGD_MS10LDL2")
'''''                .SXLGD_MS10LDL3 = rs("SXLGD_MS10LDL3")
'''''                .SXLGD_MS10LDL4 = rs("SXLGD_MS10LDL4")
'''''                .SXLGD_MS10LDL5 = rs("SXLGD_MS10LDL5")
'''''                .SXLGD_MS10DEN1 = rs("SXLGD_MS10DEN1")
'''''                .SXLGD_MS10DEN2 = rs("SXLGD_MS10DEN2")
'''''                .SXLGD_MS10DEN3 = rs("SXLGD_MS10DEN3")
'''''                .SXLGD_MS10DEN4 = rs("SXLGD_MS10DEN4")
'''''                .SXLGD_MS10DEN5 = rs("SXLGD_MS10DEN5")
'''''                .SXLGD_MS11LDL1 = rs("SXLGD_MS11LDL1")
'''''                .SXLGD_MS11LDL2 = rs("SXLGD_MS11LDL2")
'''''                .SXLGD_MS11LDL3 = rs("SXLGD_MS11LDL3")
'''''                .SXLGD_MS11LDL4 = rs("SXLGD_MS11LDL4")
'''''                .SXLGD_MS11LDL5 = rs("SXLGD_MS11LDL5")
'''''                .SXLGD_MS11DEN1 = rs("SXLGD_MS11DEN1")
'''''                .SXLGD_MS11DEN2 = rs("SXLGD_MS11DEN2")
'''''                .SXLGD_MS11DEN3 = rs("SXLGD_MS11DEN3")
'''''                .SXLGD_MS11DEN4 = rs("SXLGD_MS11DEN4")
'''''                .SXLGD_MS11DEN5 = rs("SXLGD_MS11DEN5")
'''''                .SXLGD_MS12LDL1 = rs("SXLGD_MS12LDL1")
'''''                .SXLGD_MS12LDL2 = rs("SXLGD_MS12LDL2")
'''''                .SXLGD_MS12LDL3 = rs("SXLGD_MS12LDL3")
'''''                .SXLGD_MS12LDL4 = rs("SXLGD_MS12LDL4")
'''''                .SXLGD_MS12LDL5 = rs("SXLGD_MS12LDL5")
'''''                .SXLGD_MS12DEN1 = rs("SXLGD_MS12DEN1")
'''''                .SXLGD_MS12DEN2 = rs("SXLGD_MS12DEN2")
'''''                .SXLGD_MS12DEN3 = rs("SXLGD_MS12DEN3")
'''''                .SXLGD_MS12DEN4 = rs("SXLGD_MS12DEN4")
'''''                .SXLGD_MS12DEN5 = rs("SXLGD_MS12DEN5")
'''''                .SXLGD_MS13LDL1 = rs("SXLGD_MS13LDL1")
'''''                .SXLGD_MS13LDL2 = rs("SXLGD_MS13LDL2")
'''''                .SXLGD_MS13LDL3 = rs("SXLGD_MS13LDL3")
'''''                .SXLGD_MS13LDL4 = rs("SXLGD_MS13LDL4")
'''''                .SXLGD_MS13LDL5 = rs("SXLGD_MS13LDL5")
'''''                .SXLGD_MS13DEN1 = rs("SXLGD_MS13DEN1")
'''''                .SXLGD_MS13DEN2 = rs("SXLGD_MS13DEN2")
'''''                .SXLGD_MS13DEN3 = rs("SXLGD_MS13DEN3")
'''''                .SXLGD_MS13DEN4 = rs("SXLGD_MS13DEN4")
'''''                .SXLGD_MS13DEN5 = rs("SXLGD_MS13DEN5")
'''''                .SXLGD_MS14LDL1 = rs("SXLGD_MS14LDL1")
'''''                .SXLGD_MS14LDL2 = rs("SXLGD_MS14LDL2")
'''''                .SXLGD_MS14LDL3 = rs("SXLGD_MS14LDL3")
'''''                .SXLGD_MS14LDL4 = rs("SXLGD_MS14LDL4")
'''''                .SXLGD_MS14LDL5 = rs("SXLGD_MS14LDL5")
'''''                .SXLGD_MS14DEN1 = rs("SXLGD_MS14DEN1")
'''''                .SXLGD_MS14DEN2 = rs("SXLGD_MS14DEN2")
'''''                .SXLGD_MS14DEN3 = rs("SXLGD_MS14DEN3")
'''''                .SXLGD_MS14DEN4 = rs("SXLGD_MS14DEN4")
'''''                .SXLGD_MS14DEN5 = rs("SXLGD_MS14DEN5")
'''''                .SXLGD_MS15LDL1 = rs("SXLGD_MS15LDL1")
'''''                .SXLGD_MS15LDL2 = rs("SXLGD_MS15LDL2")
'''''                .SXLGD_MS15LDL3 = rs("SXLGD_MS15LDL3")
'''''                .SXLGD_MS15LDL4 = rs("SXLGD_MS15LDL4")
'''''                .SXLGD_MS15LDL5 = rs("SXLGD_MS15LDL5")
'''''                .SXLGD_MS15DEN1 = rs("SXLGD_MS15DEN1")
'''''                .SXLGD_MS15DEN2 = rs("SXLGD_MS15DEN2")
'''''                .SXLGD_MS15DEN3 = rs("SXLGD_MS15DEN3")
'''''                .SXLGD_MS15DEN4 = rs("SXLGD_MS15DEN4")
'''''                .SXLGD_MS15DEN5 = rs("SXLGD_MS15DEN5")
'''''                .SXLGD_MSRSDEN = rs("SXLGD_MSRSDEN")
'''''                .SXLGD_MSRSLDL = rs("SXLGD_MSRSLDL")
'''''                .SXLGD_MSRSDVD2 = rs("SXLGD_MSRSDVD2")
'''''                If IsNull(rs("SXLGD_MS01DVD2")) = False Then .SXLGD_MS01DVD2 = rs("SXLGD_MS01DVD2")
'''''                If IsNull(rs("SXLGD_MS02DVD2")) = False Then .SXLGD_MS02DVD2 = rs("SXLGD_MS02DVD2")
'''''                If IsNull(rs("SXLGD_MS03DVD2")) = False Then .SXLGD_MS03DVD2 = rs("SXLGD_MS03DVD2")
'''''                If IsNull(rs("SXLGD_MS04DVD2")) = False Then .SXLGD_MS04DVD2 = rs("SXLGD_MS04DVD2")
'''''                If IsNull(rs("SXLGD_MS05DVD2")) = False Then .SXLGD_MS05DVD2 = rs("SXLGD_MS05DVD2")
'''''            End If
'''''            If (.SXLT_SMPPOS = -1) And (rs("SXLT_SMPPOS") <> -1) Then
'''''                'LT実績をコピー
'''''                .SXLT_SMPPOS = rs("SXLT_SMPPOS")
'''''                .SXLLT_MEASPEAK = rs("SXLLT_MEASPEAK")
'''''                .SXLLT_MEAS1 = rs("SXLLT_MEAS1")
'''''                .SXLLT_MEAS2 = rs("SXLLT_MEAS2")
'''''                .SXLLT_MEAS3 = rs("SXLLT_MEAS3")
'''''                .SXLLT_MEAS4 = rs("SXLLT_MEAS4")
'''''                .SXLLT_MEAS5 = rs("SXLLT_MEAS5")
'''''                .SXLLT_CALCMEAS = rs("SXLLT_CALCMEAS")
'''''            End If
'''''        End If
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End With
'''''End Sub


''''''概要      :総合判定 結晶総合判定測定値挿入用ドライバ
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
''''''          :BLOCKID       ,I  ,String       ,ブロックID
''''''          :YoneFlg       ,I  ,Boolean      ,ブロックIDが米沢かどうかのフラグ（Trueは米沢）
''''''          :Soku          ,I  ,typ_TBCMJ014 ,結晶総合判定測定値テーブルへの挿入用
''''''          :戻り値        ,O  ,FUNCTION_RETURN,読み込み成否
''''''説明      :
''''''履歴      :2001/06/27 蔵本 作成
'''''Public Function DBDRV_scmzc_fcmkc001c_InsSoku(BLOCKID As String, _
'''''                                           YoneFlg As Boolean, _
'''''                                           Soku As typ_TBCMJ014 _
'''''                                           ) As FUNCTION_RETURN
'''''
'''''    Dim sql As String
'''''    Dim rs As OraDynaset
'''''    Dim recCnt As Integer
'''''    Dim i As Long
'''''    Dim PLUPDATE As String
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_InsSoku"
'''''
'''''    DBDRV_scmzc_fcmkc001c_InsSoku = FUNCTION_RETURN_SUCCESS
'''''
'''''
'''''    If YoneFlg = True Then
'''''        '引上げ終了実績から引上日付取得 処理回数はなくなる
'''''        sql = "select "
'''''        sql = sql & " to_char(REGDATE,'YYYYMMDDHH24MISS') as cDate "
'''''        sql = sql & " from TBCMH004 "
'''''        sql = sql & " where CRYNUM='" & Soku.CRYNUM & "' "
'''''
'''''        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''        If rs.RecordCount = 0 Then
'''''            DBDRV_scmzc_fcmkc001c_InsSoku = FUNCTION_RETURN_FAILURE
'''''            GoTo proc_exit
'''''        End If
'''''
'''''        PLUPDATE = rs("cDate")
'''''        rs.Close
'''''    Else
'''''        '購入単結晶実績から引上げ日付取得
'''''        sql = "select "
'''''        sql = sql & " to_char(REGDATE,'YYYYMMDDHH24MISS') as cDate "
'''''        sql = sql & " from TBCMG002 "
'''''        sql = sql & " where CRYNUM='" & BLOCKID & "' "
'''''        sql = sql & " and TRANCNT=any(select max(TRANCNT) from TBCMG002 where CRYNUM='" & BLOCKID & "' ) "
'''''        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''
'''''        If rs.RecordCount = 0 Then
'''''            DBDRV_scmzc_fcmkc001c_InsSoku = FUNCTION_RETURN_FAILURE
'''''            GoTo proc_exit
'''''        End If
'''''
'''''        PLUPDATE = rs("cDate")
'''''        rs.Close
'''''    End If
'''''
'''''    '共用の場合には、既存レコードから必要項目を取得する
'''''    UpdateFromOrgJ014 Soku
'''''
'''''    '共用の場合に既にレコードが存在する場合があるので、まず削除する
'''''    With Soku
'''''        sql = "delete from TBCMJ014 "
'''''        sql = sql & "where (CRYNUM='" & .CRYNUM & "')"       ' 結晶番号
'''''        sql = sql & " and (POSITION=" & .POSITION & ")"        ' 位置
'''''        sql = sql & " and (SMPKBN='" & .SMPKBN & "') "        ' サンプル区分
'''''        If 0 > OraDB.ExecuteSQL(sql) Then
'''''            DBDRV_scmzc_fcmkc001c_InsSoku = FUNCTION_RETURN_FAILURE
'''''            GoTo proc_exit
'''''        End If
'''''    End With
'''''
'''''    '結晶総合判定測定値への挿入（TBCMJ014）
'''''    sql = "insert into TBCMJ014 ( "
'''''    sql = sql & "CRYNUM, "           ' 結晶番号
'''''    sql = sql & "POSITION, "         ' 位置
'''''    sql = sql & "SMPKBN, "           ' サンプル区分
'''''    sql = sql & "LENGTH, "           ' 長さ
'''''    sql = sql & "UBLOCKID, "         ' UブロックID
'''''    sql = sql & "DBLOCKID, "         ' DブロックID
'''''    sql = sql & "HINBAN, "           ' 品番
'''''    sql = sql & "REVNUM, "           ' 製品番号改訂番号
'''''    sql = sql & "FACTORY, "          ' 工場
'''''    sql = sql & "OPECOND, "          ' 操業条件
'''''    sql = sql & "PRODCOND, "         ' 製作条件
'''''    sql = sql & "PGID, "             ' ＰＧ－ＩＤ
'''''    sql = sql & "UPLENGTH, "         ' 引上げ長さ
'''''    sql = sql & "PLUPDATE, "         ' 引上日付
'''''    sql = sql & "FREELENG, "         ' フリー長
'''''    sql = sql & "DIAMETER, "         ' 直径
'''''    sql = sql & "CHARGE, "           ' チャージ量
'''''    sql = sql & "SEED, "             ' シード
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXL_RS_SMPPOS, "   ' SXLRSｻﾝﾌﾟﾙ測定位置（SXL測定情報）
'''''    sql = sql & "SXLRS_MEAS1, "      ' SXLRS_測定値１
'''''    sql = sql & "SXLRS_MEAS2, "      ' SXLRS_測定値２
'''''    sql = sql & "SXLRS_MEAS3, "      ' SXLRS_測定値３
'''''    sql = sql & "SXLRS_MEAS4, "      ' SXLRS_測定値４
'''''    sql = sql & "SXLRS_MEAS5, "      ' SXLRS_測定値５
'''''    sql = sql & "SXLRS_EFEHS, "      ' SXLRS_実効偏析
'''''    sql = sql & "SXLRS_RRG, "        ' SXLRS_ＲＲＧ
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXL_OI_SMPPOS, "    ' SXLOIｻﾝﾌﾟﾙ測定位置（SXL測定情報）
'''''    sql = sql & "SXLOI_OIMEAS1, "    ' SXLOI_Ｏｉ測定値１
'''''    sql = sql & "SXLOI_OIMEAS2, "    ' SXLOI_Ｏｉ測定値２
'''''    sql = sql & "SXLOI_OIMEAS3, "    ' SXLOI_Ｏｉ測定値３
'''''    sql = sql & "SXLOI_OIMEAS4, "    ' SXLOI_Ｏｉ測定値４
'''''    sql = sql & "SXLOI_OIMEAS5, "    ' SXLOI_Ｏｉ測定値５
'''''    sql = sql & "SXLOI_ORGRES, "     ' SXLOI_ＯＲＧ結果
'''''    sql = sql & "SXLOI_INSPECTWAY, " ' SXLOI_検査方法
'''''    sql = sql & "SXL_CS_SMPPOS, "   ' SXLCSｻﾝﾌﾟﾙ測定位置（SXL測定情報）
'''''    sql = sql & "SXLCS_CSMEAS, "     ' SXLCS_Cs実測値
'''''    sql = sql & "SXLCS_70PPRE, "     ' SXLCS_７０％推定値
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF1_SMPPOS, "   ' SXLOSFｻﾝﾌﾟﾙ測定位置（SXL位置情報）
'''''    sql = sql & "SXLOSF1_KKSP, "     ' SXLOSF1結晶欠陥測定位置
'''''    sql = sql & "SXLOSF1_NETU, "     ' SXLOSF1熱処理法
'''''    sql = sql & "SXLOSF1_KKSET, "    ' SXLOSF1結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
'''''    sql = sql & "SXLOSF1_MEAS1, "    ' SXLOSF1測定点１
'''''    sql = sql & "SXLOSF1_MEAS2, "    ' SXLOSF1測定点2
'''''    sql = sql & "SXLOSF1_MEAS3, "    ' SXLOSF1測定点3
'''''    sql = sql & "SXLOSF1_MEAS4, "    ' SXLOSF1測定点4
'''''    sql = sql & "SXLOSF1_MEAS5, "    ' SXLOSF1測定点5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF1_MEAS6, "    ' SXLOSF1測定点6
'''''    sql = sql & "SXLOSF1_MEAS7, "    ' SXLOSF1測定点7
'''''    sql = sql & "SXLOSF1_MEAS8, "    ' SXLOSF1測定点8
'''''    sql = sql & "SXLOSF1_MEAS9, "    ' SXLOSF1測定点9
'''''    sql = sql & "SXLOSF1_MEAS10, "   ' SXLOSF1測定点10
'''''    sql = sql & "SXLOSF1_MEAS11, "   ' SXLOSF1測定点11
'''''    sql = sql & "SXLOSF1_MEAS12, "   ' SXLOSF1測定点12
'''''    sql = sql & "SXLOSF1_MEAS13, "   ' SXLOSF1測定点13
'''''    sql = sql & "SXLOSF1_MEAS14, "   ' SXLOSF1測定点14
'''''    sql = sql & "SXLOSF1_MEAS15, "   ' SXLOSF1測定点15
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF1_MEAS16, "   ' SXLOSF1測定点16
'''''    sql = sql & "SXLOSF1_MEAS17, "   ' SXLOSF1測定点17
'''''    sql = sql & "SXLOSF1_MEAS18, "   ' SXLOSF1測定点18
'''''    sql = sql & "SXLOSF1_MEAS19, "   ' SXLOSF1測定点19
'''''    sql = sql & "SXLOSF1_MEAS20, "   ' SXLOSF1測定点20
'''''    sql = sql & "SXLOSF1_CALCMAX, "  ' OSF1SXL計算結果 Max_1
'''''    sql = sql & "SXLOSF1_CALCAVE, "  ' OSF1SXL計算結果 Ave_1
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF2_KKSP, "    ' SXLOSF２結晶欠陥測定位置
'''''    sql = sql & "SXLOSF2_NETU, "     ' SXLOSF２熱処理法
'''''    sql = sql & "SXLOSF2_KKSET, "    ' SXLOSF２結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
'''''    sql = sql & "SXLOSF2_MEAS1, "    ' SXLOSF2測定点１
'''''    sql = sql & "SXLOSF2_MEAS2, "    ' SXLOSF2測定点2
'''''    sql = sql & "SXLOSF2_MEAS3, "    ' SXLOSF2測定点3
'''''    sql = sql & "SXLOSF2_MEAS4, "    ' SXLOSF2測定点4
'''''    sql = sql & "SXLOSF2_MEAS5, "    ' SXLOSF2測定点5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF2_MEAS6, "    ' SXLOSF2測定点6
'''''    sql = sql & "SXLOSF2_MEAS7, "    ' SXLOSF2測定点7
'''''    sql = sql & "SXLOSF2_MEAS8, "    ' SXLOSF2測定点8
'''''    sql = sql & "SXLOSF2_MEAS9, "    ' SXLOSF2測定点9
'''''    sql = sql & "SXLOSF2_MEAS10, "   ' SXLOSF2測定点10
'''''    sql = sql & "SXLOSF2_MEAS11, "   ' SXLOSF2測定点11
'''''    sql = sql & "SXLOSF2_MEAS12, "   ' SXLOSF2測定点12
'''''    sql = sql & "SXLOSF2_MEAS13, "   ' SXLOSF2測定点13
'''''    sql = sql & "SXLOSF2_MEAS14, "   ' SXLOSF2測定点14
'''''    sql = sql & "SXLOSF2_MEAS15, "   ' SXLOSF2測定点15
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF2_MEAS16, "   ' SXLOSF2測定点16
'''''    sql = sql & "SXLOSF2_MEAS17, "   ' SXLOSF2測定点17
'''''    sql = sql & "SXLOSF2_MEAS18, "   ' SXLOSF2測定点18
'''''    sql = sql & "SXLOSF2_MEAS19, "   ' SXLOSF2測定点19
'''''    sql = sql & "SXLOSF2_MEAS20, "   ' SXLOSF2測定点20
'''''    sql = sql & "SXLOSF2_CALCMAX, "  ' OSF２SXL計算結果 Max_2
'''''    sql = sql & "SXLOSF2_CALCAVE, "  ' OSF２SXL計算結果 Ave_2
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF3_KKSP, "   ' SXLOSF３結晶欠陥測定位置
'''''    sql = sql & "SXLOSF3_NETU, "     ' SXLOSF３熱処理法
'''''    sql = sql & "SXLOSF3_KKSET, "    ' SXLOSF３結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
'''''    sql = sql & "SXLOSF3_MEAS1, "    ' SXLOSF3測定点１
'''''    sql = sql & "SXLOSF3_MEAS2, "    ' SXLOSF3測定点2
'''''    sql = sql & "SXLOSF3_MEAS3, "    ' SXLOSF3測定点3
'''''    sql = sql & "SXLOSF3_MEAS4, "    ' SXLOSF3測定点4
'''''    sql = sql & "SXLOSF3_MEAS5, "    ' SXLOSF3測定点5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF3_MEAS6, "    ' SXLOSF3測定点6
'''''    sql = sql & "SXLOSF3_MEAS7, "    ' SXLOSF3測定点7
'''''    sql = sql & "SXLOSF3_MEAS8, "    ' SXLOSF3測定点8
'''''    sql = sql & "SXLOSF3_MEAS9, "    ' SXLOSF3測定点9
'''''    sql = sql & "SXLOSF3_MEAS10, "   ' SXLOSF3測定点10
'''''    sql = sql & "SXLOSF3_MEAS11, "   ' SXLOSF3測定点11
'''''    sql = sql & "SXLOSF3_MEAS12, "   ' SXLOSF3測定点12
'''''    sql = sql & "SXLOSF3_MEAS13, "   ' SXLOSF3測定点13
'''''    sql = sql & "SXLOSF3_MEAS14, "   ' SXLOSF3測定点14
'''''    sql = sql & "SXLOSF3_MEAS15, "   ' SXLOSF3測定点15
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF3_MEAS16, "   ' SXLOSF3測定点16
'''''    sql = sql & "SXLOSF3_MEAS17, "   ' SXLOSF3測定点17
'''''    sql = sql & "SXLOSF3_MEAS18, "   ' SXLOSF3測定点18
'''''    sql = sql & "SXLOSF3_MEAS19, "   ' SXLOSF3測定点19
'''''    sql = sql & "SXLOSF3_MEAS20, "   ' SXLOSF3測定点20
'''''    sql = sql & "SXLOSF3_CALCMAX, "  ' OSF３SXL計算結果 Max_3
'''''    sql = sql & "SXLOSF3_CALCAVE, "  ' OSF３SXL計算結果 Ave_3
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF4_KKSP, "   ' SXLOSF４結晶欠陥測定位置
'''''    sql = sql & "SXLOSF4_NETU, "     ' SXLOSF４熱処理法
'''''    sql = sql & "SXLOSF4_KKSET, "    ' SXLOSF４結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
'''''    sql = sql & "SXLOSF4_MEAS1, "    ' SXLOSF4測定点１
'''''    sql = sql & "SXLOSF4_MEAS2, "    ' SXLOSF4測定点2
'''''    sql = sql & "SXLOSF4_MEAS3, "    ' SXLOSF4測定点3
'''''    sql = sql & "SXLOSF4_MEAS4, "    ' SXLOSF4測定点4
'''''    sql = sql & "SXLOSF4_MEAS5, "    ' SXLOSF4測定点5
'''''    sql = sql & "SXLOSF4_MEAS6, "    ' SXLOSF4測定点6
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF4_MEAS7, "    ' SXLOSF4測定点7
'''''    sql = sql & "SXLOSF4_MEAS8, "    ' SXLOSF4測定点8
'''''    sql = sql & "SXLOSF4_MEAS9, "    ' SXLOSF4測定点9
'''''    sql = sql & "SXLOSF4_MEAS10, "   ' SXLOSF4測定点10
'''''    sql = sql & "SXLOSF4_MEAS11, "   ' SXLOSF4測定点11
'''''    sql = sql & "SXLOSF4_MEAS12, "   ' SXLOSF4測定点12
'''''    sql = sql & "SXLOSF4_MEAS13, "   ' SXLOSF4測定点13
'''''    sql = sql & "SXLOSF4_MEAS14, "   ' SXLOSF4測定点14
'''''    sql = sql & "SXLOSF4_MEAS15, "   ' SXLOSF4測定点15
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF4_MEAS16, "   ' SXLOSF4測定点16
'''''    sql = sql & "SXLOSF4_MEAS17, "   ' SXLOSF4測定点17
'''''    sql = sql & "SXLOSF4_MEAS18, "   ' SXLOSF4測定点18
'''''    sql = sql & "SXLOSF4_MEAS19, "   ' SXLOSF4測定点19
'''''    sql = sql & "SXLOSF4_MEAS20, "   ' SXLOSF4測定点20
'''''    sql = sql & "SXLOSF4_CALCMAX, "  ' OSF４SXL計算結果 Max_4
'''''    sql = sql & "SXLOSF4_CALCAVE, "  ' OSF４SXL計算結果 Ave_4
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLBMD_SMPPOS, "   ' SXLBMDｻﾝﾌﾟﾙ測定位置（SXL位置情報）
'''''    sql = sql & "SXLBMD1_KKSP, "     ' SXLBMD1結晶欠陥測定位置
'''''    sql = sql & "SXLBMD1_NETU, "     ' SXLBMD1熱処理法
'''''    sql = sql & "SXLBMD1_KKSET, "    ' SXLBMD1結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
'''''    sql = sql & "SXLBMD1_MEAS1, "    ' SXLBMD1測定点１
'''''    sql = sql & "SXLBMD1_MEAS2, "    ' SXLBMD1測定点2
'''''    sql = sql & "SXLBMD1_MEAS3, "    ' SXLBMD1測定点3
'''''    sql = sql & "SXLBMD1_MEAS4, "    ' SXLBMD1測定点4
'''''    sql = sql & "SXLBMD1_MEAS5, "    ' SXLBMD1測定点5
'''''    sql = sql & "SXLBMD1_CALCMAX, "  ' BMD1SXL計算結果 Max
'''''    sql = sql & "SXLBMD1_CALCAVE, "  ' BMD1SXL計算結果 Ave
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLBMD2_KKSP, "    ' SXLBMD２結晶欠陥測定位置
'''''    sql = sql & "SXLBMD2_NETU, "     ' SXLBMD２熱処理法
'''''    sql = sql & "SXLBMD2_KKSET, "    ' SXLBMD２結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
'''''    sql = sql & "SXLBMD2_MEAS1, "    ' SXLBMD2測定点１
'''''    sql = sql & "SXLBMD2_MEAS2, "    ' SXLBMD2測定点2
'''''    sql = sql & "SXLBMD2_MEAS3, "    ' SXLBMD2測定点3
'''''    sql = sql & "SXLBMD2_MEAS4, "    ' SXLBMD2測定点4
'''''    sql = sql & "SXLBMD2_MEAS5, "    ' SXLBMD2測定点5
'''''    sql = sql & "SXLBMD2_CALCMAX, "  ' BMD２SXL計算結果 Max
'''''    sql = sql & "SXLBMD2_CALCAVE, "  ' BMD２SXL計算結果 Ave
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLBMD3_KKSP, "   ' SXLBMD３結晶欠陥測定位置
'''''    sql = sql & "SXLBMD3_NETU, "     ' SXLBMD３熱処理法
'''''    sql = sql & "SXLBMD3_KKSET, "    ' SXLBMD３結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
'''''    sql = sql & "SXLBMD3_MEAS1, "    ' SXLBMD3測定点１
'''''    sql = sql & "SXLBMD3_MEAS2, "    ' SXLBMD3測定点2
'''''    sql = sql & "SXLBMD3_MEAS3, "    ' SXLBMD3測定点3
'''''    sql = sql & "SXLBMD3_MEAS4, "    ' SXLBMD3測定点4
'''''    sql = sql & "SXLBMD3_MEAS5, "    ' SXLBMD3測定点5
'''''    sql = sql & "SXLBMD3_CALCMAX, "  ' BMD３SXL計算結果 Max
'''''    sql = sql & "SXLBMD3_CALCAVE, "  ' BMD３SXL計算結果 Ave
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_SMPPOS, "     ' SXLGDｻﾝﾌﾟﾙ測定位置（SXL位置情報）
'''''    sql = sql & "SXLGD_MS01LDL1, "   ' SXLGD_測定値01 L/DL1
'''''    sql = sql & "SXLGD_MS01LDL2, "   ' SXLGD_測定値01 L/DL2
'''''    sql = sql & "SXLGD_MS01LDL3, "   ' SXLGD_測定値01 L/DL3
'''''    sql = sql & "SXLGD_MS01LDL4, "   ' SXLGD_測定値01 L/DL4
'''''    sql = sql & "SXLGD_MS01LDL5, "   ' SXLGD_測定値01 L/DL5
'''''    sql = sql & "SXLGD_MS01DEN1, "   ' SXLGD_測定値01 Den1
'''''    sql = sql & "SXLGD_MS01DEN2, "   ' SXLGD_測定値01 Den2
'''''    sql = sql & "SXLGD_MS01DEN3, "   ' SXLGD_測定値01 Den3
'''''    sql = sql & "SXLGD_MS01DEN4, "   ' SXLGD_測定値01 Den4
'''''    sql = sql & "SXLGD_MS01DEN5, "   ' SXLGD_測定値01 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS02LDL1, "   ' SXLGD_測定値02 L/DL1
'''''    sql = sql & "SXLGD_MS02LDL2, "   ' SXLGD_測定値02 L/DL2
'''''    sql = sql & "SXLGD_MS02LDL3, "   ' SXLGD_測定値02 L/DL3
'''''    sql = sql & "SXLGD_MS02LDL4, "   ' SXLGD_測定値02 L/DL4
'''''    sql = sql & "SXLGD_MS02LDL5, "   ' SXLGD_測定値02 L/DL5
'''''    sql = sql & "SXLGD_MS02DEN1, "   ' SXLGD_測定値02 Den1
'''''    sql = sql & "SXLGD_MS02DEN2, "   ' SXLGD_測定値02 Den2
'''''    sql = sql & "SXLGD_MS02DEN3, "   ' SXLGD_測定値02 Den3
'''''    sql = sql & "SXLGD_MS02DEN4, "   ' SXLGD_測定値02 Den4
'''''    sql = sql & "SXLGD_MS02DEN5, "   ' SXLGD_測定値02 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS03LDL1, "   ' SXLGD_測定値03 L/DL1
'''''    sql = sql & "SXLGD_MS03LDL2, "   ' SXLGD_測定値03 L/DL2
'''''    sql = sql & "SXLGD_MS03LDL3, "   ' SXLGD_測定値03 L/DL3
'''''    sql = sql & "SXLGD_MS03LDL4, "   ' SXLGD_測定値03 L/DL4
'''''    sql = sql & "SXLGD_MS03LDL5, "   ' SXLGD_測定値03 L/DL5
'''''    sql = sql & "SXLGD_MS03DEN1, "   ' SXLGD_測定値03 Den1
'''''    sql = sql & "SXLGD_MS03DEN2, "   ' SXLGD_測定値03 Den2
'''''    sql = sql & "SXLGD_MS03DEN3, "   ' SXLGD_測定値03 Den3
'''''    sql = sql & "SXLGD_MS03DEN4, "   ' SXLGD_測定値03 Den4
'''''    sql = sql & "SXLGD_MS03DEN5, "  ' SXLGD_測定値03 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS04LDL1, "   ' SXLGD_測定値04 L/DL1
'''''    sql = sql & "SXLGD_MS04LDL2, "   ' SXLGD_測定値04 L/DL2
'''''    sql = sql & "SXLGD_MS04LDL3, "   ' SXLGD_測定値04 L/DL3
'''''    sql = sql & "SXLGD_MS04LDL4, "   ' SXLGD_測定値04 L/DL4
'''''    sql = sql & "SXLGD_MS04LDL5, "   ' SXLGD_測定値04 L/DL5
'''''    sql = sql & "SXLGD_MS04DEN1, "   ' SXLGD_測定値04 Den1
'''''    sql = sql & "SXLGD_MS04DEN2, "   ' SXLGD_測定値04 Den2
'''''    sql = sql & "SXLGD_MS04DEN3, "   ' SXLGD_測定値04 Den3
'''''    sql = sql & "SXLGD_MS04DEN4, "   ' SXLGD_測定値04 Den4
'''''    sql = sql & "SXLGD_MS04DEN5, "   ' SXLGD_測定値04 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS05LDL1, "   ' SXLGD_測定値05 L/DL1
'''''    sql = sql & "SXLGD_MS05LDL2, "   ' SXLGD_測定値05 L/DL2
'''''    sql = sql & "SXLGD_MS05LDL3, "   ' SXLGD_測定値05 L/DL3
'''''    sql = sql & "SXLGD_MS05LDL4, "   ' SXLGD_測定値05 L/DL4
'''''    sql = sql & "SXLGD_MS05LDL5, "   ' SXLGD_測定値05 L/DL5
'''''    sql = sql & "SXLGD_MS05DEN1, "   ' SXLGD_測定値05 Den1
'''''    sql = sql & "SXLGD_MS05DEN2, "   ' SXLGD_測定値05 Den2
'''''    sql = sql & "SXLGD_MS05DEN3, "   ' SXLGD_測定値05 Den3
'''''    sql = sql & "SXLGD_MS05DEN4, "   ' SXLGD_測定値05 Den4
'''''    sql = sql & "SXLGD_MS05DEN5, "   ' SXLGD_測定値05 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS06LDL1, "   ' SXLGD_測定値06 L/DL1
'''''    sql = sql & "SXLGD_MS06LDL2, "   ' SXLGD_測定値06 L/DL2
'''''    sql = sql & "SXLGD_MS06LDL3, "   ' SXLGD_測定値06 L/DL3
'''''    sql = sql & "SXLGD_MS06LDL4, "   ' SXLGD_測定値06 L/DL4
'''''    sql = sql & "SXLGD_MS06LDL5, "   ' SXLGD_測定値06 L/DL5
'''''    sql = sql & "SXLGD_MS06DEN1, "   ' SXLGD_測定値06 Den1
'''''    sql = sql & "SXLGD_MS06DEN2, "   ' SXLGD_測定値06 Den2
'''''    sql = sql & "SXLGD_MS06DEN3, "   ' SXLGD_測定値06 Den3
'''''    sql = sql & "SXLGD_MS06DEN4, "   ' SXLGD_測定値06 Den4
'''''    sql = sql & "SXLGD_MS06DEN5, "   ' SXLGD_測定値06 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS07LDL1, "   ' SXLGD_測定値07 L/DL1
'''''    sql = sql & "SXLGD_MS07LDL2, "   ' SXLGD_測定値07 L/DL2
'''''    sql = sql & "SXLGD_MS07LDL3, "   ' SXLGD_測定値07 L/DL3
'''''    sql = sql & "SXLGD_MS07LDL4, "   ' SXLGD_測定値07 L/DL4
'''''    sql = sql & "SXLGD_MS07LDL5, "   ' SXLGD_測定値07 L/DL5
'''''    sql = sql & "SXLGD_MS07DEN1, "   ' SXLGD_測定値07 Den1
'''''    sql = sql & "SXLGD_MS07DEN2, "   ' SXLGD_測定値07 Den2
'''''    sql = sql & "SXLGD_MS07DEN3, "   ' SXLGD_測定値07 Den3
'''''    sql = sql & "SXLGD_MS07DEN4, "   ' SXLGD_測定値07 Den4
'''''    sql = sql & "SXLGD_MS07DEN5, "   ' SXLGD_測定値07 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS08LDL1, "   ' SXLGD_測定値08 L/DL1
'''''    sql = sql & "SXLGD_MS08LDL2, "   ' SXLGD_測定値08 L/DL2
'''''    sql = sql & "SXLGD_MS08LDL3, "   ' SXLGD_測定値08 L/DL3
'''''    sql = sql & "SXLGD_MS08LDL4, "   ' SXLGD_測定値08 L/DL4
'''''    sql = sql & "SXLGD_MS08LDL5, "   ' SXLGD_測定値08 L/DL5
'''''    sql = sql & "SXLGD_MS08DEN1, "   ' SXLGD_測定値08 Den1
'''''    sql = sql & "SXLGD_MS08DEN2, "   ' SXLGD_測定値08 Den2
'''''    sql = sql & "SXLGD_MS08DEN3, "   ' SXLGD_測定値08 Den3
'''''    sql = sql & "SXLGD_MS08DEN4, "   ' SXLGD_測定値08 Den4
'''''    sql = sql & "SXLGD_MS08DEN5, "   ' SXLGD_測定値08 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS09LDL1, "   ' SXLGD_測定値09 L/DL1
'''''    sql = sql & "SXLGD_MS09LDL2, "   ' SXLGD_測定値09 L/DL2
'''''    sql = sql & "SXLGD_MS09LDL3, "   ' SXLGD_測定値09 L/DL3
'''''    sql = sql & "SXLGD_MS09LDL4, "   ' SXLGD_測定値09 L/DL4
'''''    sql = sql & "SXLGD_MS09LDL5, "   ' SXLGD_測定値09 L/DL5
'''''    sql = sql & "SXLGD_MS09DEN1, "   ' SXLGD_測定値09 Den1
'''''    sql = sql & "SXLGD_MS09DEN2, "   ' SXLGD_測定値09 Den2
'''''    sql = sql & "SXLGD_MS09DEN3, "   ' SXLGD_測定値09 Den3
'''''    sql = sql & "SXLGD_MS09DEN4, "   ' SXLGD_測定値09 Den4
'''''    sql = sql & "SXLGD_MS09DEN5, "   ' SXLGD_測定値09 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS10LDL1, "   ' SXLGD_測定値10 L/DL1
'''''    sql = sql & "SXLGD_MS10LDL2, "   ' SXLGD_測定値10 L/DL2
'''''    sql = sql & "SXLGD_MS10LDL3, "   ' SXLGD_測定値10 L/DL3
'''''    sql = sql & "SXLGD_MS10LDL4, "   ' SXLGD_測定値10 L/DL4
'''''    sql = sql & "SXLGD_MS10LDL5, "   ' SXLGD_測定値10 L/DL5
'''''    sql = sql & "SXLGD_MS10DEN1, "   ' SXLGD_測定値10 Den1
'''''    sql = sql & "SXLGD_MS10DEN2, "   ' SXLGD_測定値10 Den2
'''''    sql = sql & "SXLGD_MS10DEN3, "   ' SXLGD_測定値10 Den3
'''''    sql = sql & "SXLGD_MS10DEN4, "   ' SXLGD_測定値10 Den4
'''''    sql = sql & "SXLGD_MS10DEN5, "   ' SXLGD_測定値10 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS11LDL1, "   ' SXLGD_測定値11 L/DL1
'''''    sql = sql & "SXLGD_MS11LDL2, "   ' SXLGD_測定値11 L/DL2
'''''    sql = sql & "SXLGD_MS11LDL3, "   ' SXLGD_測定値11 L/DL3
'''''    sql = sql & "SXLGD_MS11LDL4, "   ' SXLGD_測定値11 L/DL4
'''''    sql = sql & "SXLGD_MS11LDL5, "   ' SXLGD_測定値11 L/DL5
'''''    sql = sql & "SXLGD_MS11DEN1, "   ' SXLGD_測定値11 Den1
'''''    sql = sql & "SXLGD_MS11DEN2, "   ' SXLGD_測定値11 Den2
'''''    sql = sql & "SXLGD_MS11DEN3, "   ' SXLGD_測定値11 Den3
'''''    sql = sql & "SXLGD_MS11DEN4, "   ' SXLGD_測定値11 Den4
'''''    sql = sql & "SXLGD_MS11DEN5, "   ' SXLGD_測定値11 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS12LDL1, "   ' SXLGD_測定値12 L/DL1
'''''    sql = sql & "SXLGD_MS12LDL2, "   ' SXLGD_測定値12 L/DL2
'''''    sql = sql & "SXLGD_MS12LDL3, "   ' SXLGD_測定値12 L/DL3
'''''    sql = sql & "SXLGD_MS12LDL4, "   ' SXLGD_測定値12 L/DL4
'''''    sql = sql & "SXLGD_MS12LDL5, "   ' SXLGD_測定値12 L/DL5
'''''    sql = sql & "SXLGD_MS12DEN1, "   ' SXLGD_測定値12 Den1
'''''    sql = sql & "SXLGD_MS12DEN2, "   ' SXLGD_測定値12 Den2
'''''    sql = sql & "SXLGD_MS12DEN3, "   ' SXLGD_測定値12 Den3
'''''    sql = sql & "SXLGD_MS12DEN4, "   ' SXLGD_測定値12 Den4
'''''    sql = sql & "SXLGD_MS12DEN5, "   ' SXLGD_測定値12 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS13LDL1, "   ' SXLGD_測定値13 L/DL1
'''''    sql = sql & "SXLGD_MS13LDL2, "   ' SXLGD_測定値13 L/DL2
'''''    sql = sql & "SXLGD_MS13LDL3, "   ' SXLGD_測定値13 L/DL3
'''''    sql = sql & "SXLGD_MS13LDL4, "   ' SXLGD_測定値13 L/DL4
'''''    sql = sql & "SXLGD_MS13LDL5, "   ' SXLGD_測定値13 L/DL5
'''''    sql = sql & "SXLGD_MS13DEN1, "   ' SXLGD_測定値13 Den1
'''''    sql = sql & "SXLGD_MS13DEN2, "   ' SXLGD_測定値13 Den2
'''''    sql = sql & "SXLGD_MS13DEN3, "   ' SXLGD_測定値13 Den3
'''''    sql = sql & "SXLGD_MS13DEN4, "   ' SXLGD_測定値13 Den4
'''''    sql = sql & "SXLGD_MS13DEN5, "   ' SXLGD_測定値13 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS14LDL1, "   ' SXLGD_測定値14 L/DL1
'''''    sql = sql & "SXLGD_MS14LDL2, "   ' SXLGD_測定値14 L/DL2
'''''    sql = sql & "SXLGD_MS14LDL3, "   ' SXLGD_測定値14 L/DL3
'''''    sql = sql & "SXLGD_MS14LDL4, "   ' SXLGD_測定値14 L/DL4
'''''    sql = sql & "SXLGD_MS14LDL5, "   ' SXLGD_測定値14 L/DL5
'''''    sql = sql & "SXLGD_MS14DEN1, "   ' SXLGD_測定値14 Den1
'''''    sql = sql & "SXLGD_MS14DEN2, "   ' SXLGD_測定値14 Den2
'''''    sql = sql & "SXLGD_MS14DEN3, "   ' SXLGD_測定値14 Den3
'''''    sql = sql & "SXLGD_MS14DEN4, "   ' SXLGD_測定値14 Den4
'''''    sql = sql & "SXLGD_MS14DEN5, "   ' SXLGD_測定値14 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MS15LDL1, "   ' SXLGD_測定値15 L/DL1
'''''    sql = sql & "SXLGD_MS15LDL2, "   ' SXLGD_測定値15 L/DL2
'''''    sql = sql & "SXLGD_MS15LDL3, "   ' SXLGD_測定値15 L/DL3
'''''    sql = sql & "SXLGD_MS15LDL4, "   ' SXLGD_測定値15 L/DL4
'''''    sql = sql & "SXLGD_MS15LDL5, "   ' SXLGD_測定値15 L/DL5
'''''    sql = sql & "SXLGD_MS15DEN1, "   ' SXLGD_測定値15 Den1
'''''    sql = sql & "SXLGD_MS15DEN2, "   ' SXLGD_測定値15 Den2
'''''    sql = sql & "SXLGD_MS15DEN3, "   ' SXLGD_測定値15 Den3
'''''    sql = sql & "SXLGD_MS15DEN4, "   ' SXLGD_測定値15 Den4
'''''    sql = sql & "SXLGD_MS15DEN5, "   ' SXLGD_測定値15 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLGD_MSRSDEN, "   ' SXLGD_測定結果 Den
'''''    sql = sql & "SXLGD_MSRSLDL, "    ' SXLGD_測定結果 L/DL
'''''    sql = sql & "SXLGD_MSRSDVD2, "   ' SXLGD_測定結果 DVD2
'''''    sql = sql & "SXLT_SMPPOS, "      ' SXLLTｻﾝﾌﾟﾙ測定位置（SXL位置情報）
'''''    sql = sql & "SXLLT_MEASPEAK, "   ' SXLLT_測定値 ピーク値
'''''    sql = sql & "SXLLT_MEAS1, "      ' SXLLT_測定値1
'''''    sql = sql & "SXLLT_MEAS2, "      ' SXLLT_測定値2
'''''    sql = sql & "SXLLT_MEAS3, "      ' SXLLT_測定値3
'''''    sql = sql & "SXLLT_MEAS4, "      ' SXLLT_測定値4
'''''    sql = sql & "SXLLT_MEAS5, "      ' SXLLT_測定値5
'''''    sql = sql & "SXLLT_CALCMEAS, "   ' SXLLT_計算結果
'''''    sql = sql & "REGDATE, "          ' 登録日付
'''''    sql = sql & "SENDFLAG, "         ' 送信フラグ
'''''    sql = sql & "SENDDATE,  "         ' 送信日付
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''    sql = sql & "SXLOSF1_POS1, "      ' OSF1ﾊﾟﾀｰﾝ区分１位置
'''''    sql = sql & "SXLOSF1_WID1, "      ' OSF1ﾊﾟﾀｰﾝ区分１幅
'''''    sql = sql & "SXLOSF1_RD1, "       ' OSF1ﾊﾟﾀｰﾝ区分１R/D
'''''    sql = sql & "SXLOSF1_POS2, "      ' OSF1ﾊﾟﾀｰﾝ区分２位置
'''''    sql = sql & "SXLOSF1_WID2, "      ' OSF1ﾊﾟﾀｰﾝ区分２幅
'''''    sql = sql & "SXLOSF1_RD2, "       ' OSF1ﾊﾟﾀｰﾝ区分２R/D
'''''    sql = sql & "SXLOSF1_POS3, "      ' OSF1ﾊﾟﾀｰﾝ区分３位置
'''''    sql = sql & "SXLOSF1_WID3, "      ' OSF1ﾊﾟﾀｰﾝ区分３幅
'''''    sql = sql & "SXLOSF1_RD3, "       ' OSF1ﾊﾟﾀｰﾝ区分３R/D
'''''    sql = sql & "SXLOSF2_POS1, "      ' OSF2ﾊﾟﾀｰﾝ区分１位置
'''''    sql = sql & "SXLOSF2_WID1, "      ' OSF2ﾊﾟﾀｰﾝ区分１幅
'''''    sql = sql & "SXLOSF2_RD1, "       ' OSF2ﾊﾟﾀｰﾝ区分１R/D
'''''    sql = sql & "SXLOSF2_POS2, "      ' OSF2ﾊﾟﾀｰﾝ区分２位置
'''''    sql = sql & "SXLOSF2_WID2, "      ' OSF2ﾊﾟﾀｰﾝ区分２幅
'''''    sql = sql & "SXLOSF2_RD2, "       ' OSF2ﾊﾟﾀｰﾝ区分２R/D
'''''    sql = sql & "SXLOSF2_POS3, "      ' OSF2ﾊﾟﾀｰﾝ区分３位置
'''''    sql = sql & "SXLOSF2_WID3, "      ' OSF2ﾊﾟﾀｰﾝ区分３幅
'''''    sql = sql & "SXLOSF2_RD3, "       ' OSF2ﾊﾟﾀｰﾝ区分３R/D
'''''    sql = sql & "SXLOSF3_POS1, "      ' OSF3ﾊﾟﾀｰﾝ区分１位置
'''''    sql = sql & "SXLOSF3_WID1, "      ' OSF3ﾊﾟﾀｰﾝ区分１幅
'''''    sql = sql & "SXLOSF3_RD1, "       ' OSF3ﾊﾟﾀｰﾝ区分１R/D
'''''    sql = sql & "SXLOSF3_POS2, "      ' OSF3ﾊﾟﾀｰﾝ区分２位置
'''''    sql = sql & "SXLOSF3_WID2, "      ' OSF3ﾊﾟﾀｰﾝ区分２幅
'''''    sql = sql & "SXLOSF3_RD2, "       ' OSF3ﾊﾟﾀｰﾝ区分２R/D
'''''    sql = sql & "SXLOSF3_POS3, "      ' OSF3ﾊﾟﾀｰﾝ区分３位置
'''''    sql = sql & "SXLOSF3_WID3, "      ' OSF3ﾊﾟﾀｰﾝ区分３幅
'''''    sql = sql & "SXLOSF3_RD3, "       ' OSF3ﾊﾟﾀｰﾝ区分３R/D
'''''    sql = sql & "SXLOSF4_POS1, "      ' OSF4ﾊﾟﾀｰﾝ区分１位置
'''''    sql = sql & "SXLOSF4_WID1, "      ' OSF4ﾊﾟﾀｰﾝ区分１幅
'''''    sql = sql & "SXLOSF4_RD1, "       ' OSF4ﾊﾟﾀｰﾝ区分１R/D
'''''    sql = sql & "SXLOSF4_POS2, "      ' OSF4ﾊﾟﾀｰﾝ区分２位置
'''''    sql = sql & "SXLOSF4_WID2, "      ' OSF4ﾊﾟﾀｰﾝ区分２幅
'''''    sql = sql & "SXLOSF4_RD2, "       ' OSF4ﾊﾟﾀｰﾝ区分２R/D
'''''    sql = sql & "SXLOSF4_POS3, "      ' OSF4ﾊﾟﾀｰﾝ区分３位置
'''''    sql = sql & "SXLOSF4_WID3, "      ' OSF4ﾊﾟﾀｰﾝ区分３幅
'''''    sql = sql & "SXLOSF4_RD3, "       ' OSF4ﾊﾟﾀｰﾝ区分３R/D
'''''    sql = sql & "SXLGD_MS01DVD2, "    ' DVD2測定結果値１
'''''    sql = sql & "SXLGD_MS02DVD2, "    ' DVD2測定結果値２
'''''    sql = sql & "SXLGD_MS03DVD2, "    ' DVD2測定結果値３
'''''    sql = sql & "SXLGD_MS04DVD2, "    ' DVD2測定結果値４
'''''    sql = sql & "SXLGD_MS05DVD2, "    ' DVD2測定結果値５
'''''    sql = sql & "SXLBMD1_MNBCR, "     ' BMD1SXL計算結果面内分布
'''''    sql = sql & "SXLBMD2_MNBCR, "     ' BMD2SXL計算結果面内分布
'''''    sql = sql & "SXLBMD3_MNBCR ) "    ' BMD3SXL計算結果面内分布
'''''    With Soku
'''''        sql = sql & " values ( "
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " '" & .CRYNUM & "', "       ' 結晶番号
'''''        sql = sql & " " & .POSITION & ", "        ' 位置
'''''        sql = sql & " '" & .SMPKBN & "', "        ' サンプル区分
'''''        sql = sql & " " & .LENGTH & ", "          ' 長さ
'''''        sql = sql & " '" & .UBLOCKID & "', "      ' UブロックID
'''''        sql = sql & " '" & .DBLOCKID & "', "      ' DブロックID
'''''        sql = sql & " '" & .hinban & "', "        ' 品番
'''''        sql = sql & " " & .REVNUM & ", "          ' 製品番号改訂番号
'''''        sql = sql & " '" & .factory & "', "       ' 工場
'''''        sql = sql & " '" & .opecond & "', "       ' 操業条件
'''''        sql = sql & " '" & .PRODCOND & "', "      ' 製作条件
'''''        sql = sql & " '" & Mid(.PGID, 1, 8) & "', "        ' ＰＧ－ＩＤ
'''''        sql = sql & " " & .UPLENGTH & ", "        ' 引上げ長さ
'''''        sql = sql & " to_date('" & PLUPDATE & "','YYYYMMDDHH24MISS'), "    ' 引上日付
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " " & .FREELENG & ", "        ' フリー長
'''''        sql = sql & " " & .DIAMETER & ", "        ' 直径
'''''        sql = sql & " '" & .CHARGE & "', "        ' チャージ量
'''''        sql = sql & " '" & .SEED & "', "          ' シード
'''''        sql = sql & " " & .SXL_RS_SMPPOS & ", "  ' SXLRSｻﾝﾌﾟﾙ測定位置（SXL測定情報）
'''''        sql = sql & " " & .SXLRS_MEAS1 & ", "     ' SXLRS_測定値１
'''''        sql = sql & " " & .SXLRS_MEAS2 & ", "     ' SXLRS_測定値２
'''''        sql = sql & " " & .SXLRS_MEAS3 & ", "     ' SXLRS_測定値３
'''''        sql = sql & " " & .SXLRS_MEAS4 & ", "     ' SXLRS_測定値４
'''''        sql = sql & " " & .SXLRS_MEAS5 & ", "     ' SXLRS_測定値５
'''''        sql = sql & " " & .SXLRS_EFEHS & ", "     ' SXLRS_実効偏析
'''''        sql = sql & " " & .SXLRS_RRG & ", "       ' SXLRS_ＲＲＧ
'''''        sql = sql & " " & .SXL_OI_SMPPOS & ", "   ' SXLOIｻﾝﾌﾟﾙ測定位置（SXL測定情報）
'''''        sql = sql & " " & .SXLOI_OIMEAS1 & ", "   ' SXLOI_Ｏｉ測定値１
'''''        sql = sql & " " & .SXLOI_OIMEAS2 & ", "   ' SXLOI_Ｏｉ測定値２
'''''        sql = sql & " " & .SXLOI_OIMEAS3 & ", "   ' SXLOI_Ｏｉ測定値３
'''''        sql = sql & " " & .SXLOI_OIMEAS4 & ", "   ' SXLOI_Ｏｉ測定値４
'''''        sql = sql & " " & .SXLOI_OIMEAS5 & ", "   ' SXLOI_Ｏｉ測定値５
'''''        sql = sql & " " & .SXLOI_ORGRES & ", "    ' SXLOI_ＯＲＧ結果
'''''        sql = sql & " '" & .SXLOI_INSPECTWAY & "', " ' SXLOI_検査方法
'''''        sql = sql & " " & .SXL_CS_SMPPOS & ", "      ' SXLCSｻﾝﾌﾟﾙ測定位置（SXL測定情報）
'''''        sql = sql & " " & .SXLCS_CSMEAS & ", "       ' SXLCS_Cs実測値
'''''        sql = sql & " " & .SXLCS_70PPRE & ", "       ' SXLCS_７０％推定値
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " " & .SXLOSF1_SMPPOS & ", "    ' SXLOSFｻﾝﾌﾟﾙ測定位置（SXL位置情報）
'''''        sql = sql & " '" & .SXLOSF1_KKSP & "', "    ' SXLOSF1結晶欠陥測定位置
'''''        sql = sql & " '" & .SXLOSF1_NETU & "', "    ' SXLOSF1熱処理法
'''''        sql = sql & " '" & .SXLOSF1_KKSET & "', "   ' SXLOSF1結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
'''''        sql = sql & " " & .SXLOSF1_MEAS1 & ", "   ' SXLOSF1測定点１
'''''        sql = sql & " " & .SXLOSF1_MEAS2 & ", "   ' SXLOSF1測定点2
'''''        sql = sql & " " & .SXLOSF1_MEAS3 & ", "   ' SXLOSF1測定点3
'''''        sql = sql & " " & .SXLOSF1_MEAS4 & ", "   ' SXLOSF1測定点4
'''''        sql = sql & " " & .SXLOSF1_MEAS5 & ", "   ' SXLOSF1測定点5
'''''        sql = sql & " " & .SXLOSF1_MEAS6 & ", "   ' SXLOSF1測定点6
'''''        sql = sql & " " & .SXLOSF1_MEAS7 & ", "   ' SXLOSF1測定点7
'''''        sql = sql & " " & .SXLOSF1_MEAS8 & ", "   ' SXLOSF1測定点8
'''''        sql = sql & " " & .SXLOSF1_MEAS9 & ", "   ' SXLOSF1測定点9
'''''        sql = sql & " " & .SXLOSF1_MEAS10 & ", " ' SXLOSF1測定点10
'''''        sql = sql & " " & .SXLOSF1_MEAS11 & ", "  ' SXLOSF1測定点11
'''''        sql = sql & " " & .SXLOSF1_MEAS12 & ", "  ' SXLOSF1測定点12
'''''        sql = sql & " " & .SXLOSF1_MEAS13 & ", "  ' SXLOSF1測定点13
'''''        sql = sql & " " & .SXLOSF1_MEAS14 & ", "  ' SXLOSF1測定点14
'''''        sql = sql & " " & .SXLOSF1_MEAS15 & ", "  ' SXLOSF1測定点15
'''''        sql = sql & " " & .SXLOSF1_MEAS16 & ", "  ' SXLOSF1測定点16
'''''        sql = sql & " " & .SXLOSF1_MEAS17 & ", "  ' SXLOSF1測定点17
'''''        sql = sql & " " & .SXLOSF1_MEAS18 & ", "  ' SXLOSF1測定点18
'''''        sql = sql & " " & .SXLOSF1_MEAS19 & ", "  ' SXLOSF1測定点19
'''''        sql = sql & " " & .SXLOSF1_MEAS20 & ", "  ' SXLOSF1測定点20
'''''        sql = sql & " " & .SXLOSF1_CALCMAX & ", " ' OSF1SXL計算結果 Max_1
'''''        sql = sql & " " & .SXLOSF1_CALCAVE & ", "  ' OSF1SXL計算結果 Ave_1
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " '" & .SXLOSF2_KKSP & "', "    ' SXLOSF２結晶欠陥測定位置
'''''        sql = sql & " '" & .SXLOSF2_NETU & "', "    ' SXLOSF２熱処理法
'''''        sql = sql & " '" & .SXLOSF2_KKSET & "', "   ' SXLOSF２結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
'''''        sql = sql & " " & .SXLOSF2_MEAS1 & ", "   ' SXLOSF2測定点１
'''''        sql = sql & " " & .SXLOSF2_MEAS2 & ", "   ' SXLOSF2測定点2
'''''        sql = sql & " " & .SXLOSF2_MEAS3 & ", "   ' SXLOSF2測定点3
'''''        sql = sql & " " & .SXLOSF2_MEAS4 & ", "   ' SXLOSF2測定点4
'''''        sql = sql & " " & .SXLOSF2_MEAS5 & ", "   ' SXLOSF2測定点5
'''''        sql = sql & " " & .SXLOSF2_MEAS6 & ", "   ' SXLOSF2測定点6
'''''        sql = sql & " " & .SXLOSF2_MEAS7 & ", "   ' SXLOSF2測定点7
'''''        sql = sql & " " & .SXLOSF2_MEAS8 & ", "   ' SXLOSF2測定点8
'''''        sql = sql & " " & .SXLOSF2_MEAS9 & ", "   ' SXLOSF2測定点9
'''''        sql = sql & " " & .SXLOSF2_MEAS10 & ", "  ' SXLOSF2測定点10
'''''        sql = sql & " " & .SXLOSF2_MEAS11 & ", "  ' SXLOSF2測定点11
'''''        sql = sql & " " & .SXLOSF2_MEAS12 & ", "  ' SXLOSF2測定点12
'''''        sql = sql & " " & .SXLOSF2_MEAS13 & ", "  ' SXLOSF2測定点13
'''''        sql = sql & " " & .SXLOSF2_MEAS14 & ", "  ' SXLOSF2測定点14
'''''        sql = sql & " " & .SXLOSF2_MEAS15 & ", "  ' SXLOSF2測定点15
'''''        sql = sql & " " & .SXLOSF2_MEAS16 & ", "  ' SXLOSF2測定点16
'''''        sql = sql & " " & .SXLOSF2_MEAS17 & ", "  ' SXLOSF2測定点17
'''''        sql = sql & " " & .SXLOSF2_MEAS18 & ", "  ' SXLOSF2測定点18
'''''        sql = sql & " " & .SXLOSF2_MEAS19 & ", "  ' SXLOSF2測定点19
'''''        sql = sql & " " & .SXLOSF2_MEAS20 & ", "  ' SXLOSF2測定点20
'''''        sql = sql & " " & .SXLOSF2_CALCMAX & ", " ' OSF２SXL計算結果 Max_2
'''''        sql = sql & " " & .SXLOSF2_CALCAVE & ", " ' OSF２SXL計算結果 Ave_2
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " '" & .SXLOSF3_KKSP & "', "    ' SXLOSF３結晶欠陥測定位置
'''''        sql = sql & " '" & .SXLOSF3_NETU & "', "    ' SXLOSF３熱処理法
'''''        sql = sql & " '" & .SXLOSF3_KKSET & "', "   ' SXLOSF３結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
'''''        sql = sql & " " & .SXLOSF3_MEAS1 & ", "   ' SXLOSF3測定点１
'''''        sql = sql & " " & .SXLOSF3_MEAS2 & ", "   ' SXLOSF3測定点2
'''''        sql = sql & " " & .SXLOSF3_MEAS3 & ", "   ' SXLOSF3測定点3
'''''        sql = sql & " " & .SXLOSF3_MEAS4 & ", "   ' SXLOSF3測定点4
'''''        sql = sql & " " & .SXLOSF3_MEAS5 & ", "   ' SXLOSF3測定点5
'''''        sql = sql & " " & .SXLOSF3_MEAS6 & ", "   ' SXLOSF3測定点6
'''''        sql = sql & " " & .SXLOSF3_MEAS7 & ", "   ' SXLOSF3測定点7
'''''        sql = sql & " " & .SXLOSF3_MEAS8 & ", "   ' SXLOSF3測定点8
'''''        sql = sql & " " & .SXLOSF3_MEAS9 & ", "   ' SXLOSF3測定点9
'''''        sql = sql & " " & .SXLOSF3_MEAS10 & ", "  ' SXLOSF3測定点10
'''''        sql = sql & " " & .SXLOSF3_MEAS11 & ", "  ' SXLOSF3測定点11
'''''        sql = sql & " " & .SXLOSF3_MEAS12 & ", "  ' SXLOSF3測定点12
'''''        sql = sql & " " & .SXLOSF3_MEAS13 & ", "  ' SXLOSF3測定点13
'''''        sql = sql & " " & .SXLOSF3_MEAS14 & ", "  ' SXLOSF3測定点14
'''''        sql = sql & " " & .SXLOSF3_MEAS15 & ", "  ' SXLOSF3測定点15
'''''        sql = sql & " " & .SXLOSF3_MEAS16 & ", "  ' SXLOSF3測定点16
'''''        sql = sql & " " & .SXLOSF3_MEAS17 & ", "  ' SXLOSF3測定点17
'''''        sql = sql & " " & .SXLOSF3_MEAS18 & ", "  ' SXLOSF3測定点18
'''''        sql = sql & " " & .SXLOSF3_MEAS19 & ", "  ' SXLOSF3測定点19
'''''        sql = sql & " " & .SXLOSF3_MEAS20 & ", "  ' SXLOSF3測定点20
'''''        sql = sql & " " & .SXLOSF3_CALCMAX & ", " ' OSF３SXL計算結果 Max_3
'''''        sql = sql & " " & .SXLOSF3_CALCAVE & ", " ' OSF３SXL計算結果 Ave_3
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " '" & .SXLOSF4_KKSP & "', "    ' SXLOSF４結晶欠陥測定位置
'''''        sql = sql & " '" & .SXLOSF4_NETU & "', "    ' SXLOSF４熱処理法
'''''        sql = sql & " '" & .SXLOSF4_KKSET & "', "   ' SXLOSF４結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
'''''        sql = sql & " " & .SXLOSF4_MEAS1 & ", "   ' SXLOSF4測定点１
'''''        sql = sql & " " & .SXLOSF4_MEAS2 & ", "   ' SXLOSF4測定点2
'''''        sql = sql & " " & .SXLOSF4_MEAS3 & ", "   ' SXLOSF4測定点3
'''''        sql = sql & " " & .SXLOSF4_MEAS4 & ", "   ' SXLOSF4測定点4
'''''        sql = sql & " " & .SXLOSF4_MEAS5 & ", "   ' SXLOSF4測定点5
'''''        sql = sql & " " & .SXLOSF4_MEAS6 & ", "   ' SXLOSF4測定点6
'''''        sql = sql & " " & .SXLOSF4_MEAS7 & ", "   ' SXLOSF4測定点7
'''''        sql = sql & " " & .SXLOSF4_MEAS8 & ", "   ' SXLOSF4測定点8
'''''        sql = sql & " " & .SXLOSF4_MEAS9 & ", "   ' SXLOSF4測定点9
'''''        sql = sql & " " & .SXLOSF4_MEAS10 & ", "  ' SXLOSF4測定点10
'''''        sql = sql & " " & .SXLOSF4_MEAS11 & ", " ' SXLOSF4測定点11
'''''        sql = sql & " " & .SXLOSF4_MEAS12 & ", "  ' SXLOSF4測定点12
'''''        sql = sql & " " & .SXLOSF4_MEAS13 & ", "  ' SXLOSF4測定点13
'''''        sql = sql & " " & .SXLOSF4_MEAS14 & ", "  ' SXLOSF4測定点14
'''''        sql = sql & " " & .SXLOSF4_MEAS15 & ", "  ' SXLOSF4測定点15
'''''        sql = sql & " " & .SXLOSF4_MEAS16 & ", "  ' SXLOSF4測定点16
'''''        sql = sql & " " & .SXLOSF4_MEAS17 & ", "  ' SXLOSF4測定点17
'''''        sql = sql & " " & .SXLOSF4_MEAS18 & ", "  ' SXLOSF4測定点18
'''''        sql = sql & " " & .SXLOSF4_MEAS19 & ", "  ' SXLOSF4測定点19
'''''        sql = sql & " " & .SXLOSF4_MEAS20 & ", "  ' SXLOSF4測定点20
'''''        sql = sql & " " & .SXLOSF4_CALCMAX & ", " ' OSF４SXL計算結果 Max_4
'''''        sql = sql & " " & .SXLOSF4_CALCAVE & ", " ' OSF４SXL計算結果 Ave_4
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " " & .SXLBMD_SMPPOS & ", "   ' SXLBMDｻﾝﾌﾟﾙ測定位置（SXL位置情報）
'''''        sql = sql & " '" & .SXLBMD1_KKSP & "', "    ' SXLBMD1結晶欠陥測定位置
'''''        sql = sql & " '" & .SXLBMD1_NETU & "', "    ' SXLBMD1熱処理法
'''''        sql = sql & " '" & .SXLBMD1_KKSET & "', "   ' SXLBMD1結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
'''''        sql = sql & " " & .SXLBMD1_MEAS1 & ", "   ' SXLBMD1測定点１
'''''        sql = sql & " " & .SXLBMD1_MEAS2 & ", "   ' SXLBMD1測定点2
'''''        sql = sql & " " & .SXLBMD1_MEAS3 & ", "   ' SXLBMD1測定点3
'''''        sql = sql & " " & .SXLBMD1_MEAS4 & ", "   ' SXLBMD1測定点4
'''''        sql = sql & " " & .SXLBMD1_MEAS5 & ", "   ' SXLBMD1測定点5
'''''        sql = sql & " " & .SXLBMD1_CALCMAX & ", " ' BMD1SXL計算結果 Max
'''''        sql = sql & " " & .SXLBMD1_CALCAVE & ", " ' BMD1SXL計算結果 Ave
'''''        sql = sql & " '" & .SXLBMD2_KKSP & "', "    ' SXLBMD２結晶欠陥測定位置
'''''        sql = sql & " '" & .SXLBMD2_NETU & "', "    ' SXLBMD２熱処理法
'''''        sql = sql & " '" & .SXLBMD2_KKSET & "', "   ' SXLBMD２結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
'''''        sql = sql & " " & .SXLBMD2_MEAS1 & ", "   ' SXLBMD2測定点１
'''''        sql = sql & " " & .SXLBMD2_MEAS2 & ", "   ' SXLBMD2測定点2
'''''        sql = sql & " " & .SXLBMD2_MEAS3 & ", "   ' SXLBMD2測定点3
'''''        sql = sql & " " & .SXLBMD2_MEAS4 & ", "   ' SXLBMD2測定点4
'''''        sql = sql & " " & .SXLBMD2_MEAS5 & ", "   ' SXLBMD2測定点5
'''''        sql = sql & " " & .SXLBMD2_CALCMAX & ", " ' BMD２SXL計算結果 Max
'''''        sql = sql & " " & .SXLBMD2_CALCAVE & ", " ' BMD２SXL計算結果 Ave
'''''        sql = sql & " '" & .SXLBMD3_KKSP & "', "    ' SXLBMD３結晶欠陥測定位置
'''''        sql = sql & " '" & .SXLBMD3_NETU & "', "    ' SXLBMD３熱処理法
'''''        sql = sql & " '" & .SXLBMD3_KKSET & "', "  ' SXLBMD３結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
'''''        sql = sql & " " & .SXLBMD3_MEAS1 & ", "   ' SXLBMD3測定点１
'''''        sql = sql & " " & .SXLBMD3_MEAS2 & ", "   ' SXLBMD3測定点2
'''''        sql = sql & " " & .SXLBMD3_MEAS3 & ", "   ' SXLBMD3測定点3
'''''        sql = sql & " " & .SXLBMD3_MEAS4 & ", "   ' SXLBMD3測定点4
'''''        sql = sql & " " & .SXLBMD3_MEAS5 & ", "   ' SXLBMD3測定点5
'''''        sql = sql & " " & .SXLBMD3_CALCMAX & ", " ' BMD３SXL計算結果 Max
'''''        sql = sql & " " & .SXLBMD3_CALCAVE & ", " ' BMD３SXL計算結果 Ave
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " " & .SXLGD_SMPPOS & ", "    ' SXLGDｻﾝﾌﾟﾙ測定位置（SXL位置情報）
'''''        sql = sql & " " & .SXLGD_MS01LDL1 & ", "  ' SXLGD_測定値01 L/DL1
'''''        sql = sql & " " & .SXLGD_MS01LDL2 & ", "  ' SXLGD_測定値01 L/DL2
'''''        sql = sql & " " & .SXLGD_MS01LDL3 & ", "  ' SXLGD_測定値01 L/DL3
'''''        sql = sql & " " & .SXLGD_MS01LDL4 & ", "  ' SXLGD_測定値01 L/DL4
'''''        sql = sql & " " & .SXLGD_MS01LDL5 & ", "  ' SXLGD_測定値01 L/DL5
'''''        sql = sql & " " & .SXLGD_MS01DEN1 & ", "  ' SXLGD_測定値01 Den1
'''''        sql = sql & " " & .SXLGD_MS01DEN2 & ", "  ' SXLGD_測定値01 Den2
'''''        sql = sql & " " & .SXLGD_MS01DEN3 & ", "  ' SXLGD_測定値01 Den3
'''''        sql = sql & " " & .SXLGD_MS01DEN4 & ", "  ' SXLGD_測定値01 Den4
'''''        sql = sql & " " & .SXLGD_MS01DEN5 & ", "  ' SXLGD_測定値01 Den5
'''''        sql = sql & " " & .SXLGD_MS02LDL1 & ", "  ' SXLGD_測定値02 L/DL1
'''''        sql = sql & " " & .SXLGD_MS02LDL2 & ", "  ' SXLGD_測定値02 L/DL2
'''''        sql = sql & " " & .SXLGD_MS02LDL3 & ", "  ' SXLGD_測定値02 L/DL3
'''''        sql = sql & " " & .SXLGD_MS02LDL4 & ", "  ' SXLGD_測定値02 L/DL4
'''''        sql = sql & " " & .SXLGD_MS02LDL5 & ", "  ' SXLGD_測定値02 L/DL5
'''''        sql = sql & " " & .SXLGD_MS02DEN1 & ", "  ' SXLGD_測定値02 Den1
'''''        sql = sql & " " & .SXLGD_MS02DEN2 & ", "  ' SXLGD_測定値02 Den2
'''''        sql = sql & " " & .SXLGD_MS02DEN3 & ", "  ' SXLGD_測定値02 Den3
'''''        sql = sql & " " & .SXLGD_MS02DEN4 & ", "  ' SXLGD_測定値02 Den4
'''''        sql = sql & " " & .SXLGD_MS02DEN5 & ", "  ' SXLGD_測定値02 Den5
'''''        sql = sql & " " & .SXLGD_MS03LDL1 & ", "  ' SXLGD_測定値03 L/DL1
'''''        sql = sql & " " & .SXLGD_MS03LDL2 & ", "  ' SXLGD_測定値03 L/DL2
'''''        sql = sql & " " & .SXLGD_MS03LDL3 & ", "  ' SXLGD_測定値03 L/DL3
'''''        sql = sql & " " & .SXLGD_MS03LDL4 & ", "  ' SXLGD_測定値03 L/DL4
'''''        sql = sql & " " & .SXLGD_MS03LDL5 & ", "  ' SXLGD_測定値03 L/DL5
'''''        sql = sql & " " & .SXLGD_MS03DEN1 & ", "  ' SXLGD_測定値03 Den1
'''''        sql = sql & " " & .SXLGD_MS03DEN2 & ", "  ' SXLGD_測定値03 Den2
'''''        sql = sql & " " & .SXLGD_MS03DEN3 & ", "  ' SXLGD_測定値03 Den3
'''''        sql = sql & " " & .SXLGD_MS03DEN4 & ", "  ' SXLGD_測定値03 Den4
'''''        sql = sql & " " & .SXLGD_MS03DEN5 & ", "  ' SXLGD_測定値03 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " " & .SXLGD_MS04LDL1 & ", "  ' SXLGD_測定値04 L/DL1
'''''        sql = sql & " " & .SXLGD_MS04LDL2 & ", "  ' SXLGD_測定値04 L/DL2
'''''        sql = sql & " " & .SXLGD_MS04LDL3 & ", "  ' SXLGD_測定値04 L/DL3
'''''        sql = sql & " " & .SXLGD_MS04LDL4 & ", "  ' SXLGD_測定値04 L/DL4
'''''        sql = sql & " " & .SXLGD_MS04LDL5 & ", "  ' SXLGD_測定値04 L/DL5
'''''        sql = sql & " " & .SXLGD_MS04DEN1 & ", "  ' SXLGD_測定値04 Den1
'''''        sql = sql & " " & .SXLGD_MS04DEN2 & ", "  ' SXLGD_測定値04 Den2
'''''        sql = sql & " " & .SXLGD_MS04DEN3 & ", "  ' SXLGD_測定値04 Den3
'''''        sql = sql & " " & .SXLGD_MS04DEN4 & ", "  ' SXLGD_測定値04 Den4
'''''        sql = sql & " " & .SXLGD_MS04DEN5 & ", "  ' SXLGD_測定値04 Den5
'''''        sql = sql & " " & .SXLGD_MS05LDL1 & ", "  ' SXLGD_測定値05 L/DL1
'''''        sql = sql & " " & .SXLGD_MS05LDL2 & ", "  ' SXLGD_測定値05 L/DL2
'''''        sql = sql & " " & .SXLGD_MS05LDL3 & ", "  ' SXLGD_測定値05 L/DL3
'''''        sql = sql & " " & .SXLGD_MS05LDL4 & ", "  ' SXLGD_測定値05 L/DL4
'''''        sql = sql & " " & .SXLGD_MS05LDL5 & ", "  ' SXLGD_測定値05 L/DL5
'''''        sql = sql & " " & .SXLGD_MS05DEN1 & ", "  ' SXLGD_測定値05 Den1
'''''        sql = sql & " " & .SXLGD_MS05DEN2 & ", "  ' SXLGD_測定値05 Den2
'''''        sql = sql & " " & .SXLGD_MS05DEN3 & ", "  ' SXLGD_測定値05 Den3
'''''        sql = sql & " " & .SXLGD_MS05DEN4 & ", "  ' SXLGD_測定値05 Den4
'''''        sql = sql & " " & .SXLGD_MS05DEN5 & ", "  ' SXLGD_測定値05 Den5
'''''        sql = sql & " " & .SXLGD_MS06LDL1 & ", "  ' SXLGD_測定値06 L/DL1
'''''        sql = sql & " " & .SXLGD_MS06LDL2 & ", "  ' SXLGD_測定値06 L/DL2
'''''        sql = sql & " " & .SXLGD_MS06LDL3 & ", "  ' SXLGD_測定値06 L/DL3
'''''        sql = sql & " " & .SXLGD_MS06LDL4 & ", "  ' SXLGD_測定値06 L/DL4
'''''        sql = sql & " " & .SXLGD_MS06LDL5 & ", "  ' SXLGD_測定値06 L/DL5
'''''        sql = sql & " " & .SXLGD_MS06DEN1 & ", "  ' SXLGD_測定値06 Den1
'''''        sql = sql & " " & .SXLGD_MS06DEN2 & ", "  ' SXLGD_測定値06 Den2
'''''        sql = sql & " " & .SXLGD_MS06DEN3 & ", "  ' SXLGD_測定値06 Den3
'''''        sql = sql & " " & .SXLGD_MS06DEN4 & ", "  ' SXLGD_測定値06 Den4
'''''        sql = sql & " " & .SXLGD_MS06DEN5 & ", "  ' SXLGD_測定値06 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " " & .SXLGD_MS07LDL1 & ", "  ' SXLGD_測定値07 L/DL1
'''''        sql = sql & " " & .SXLGD_MS07LDL2 & ", "  ' SXLGD_測定値07 L/DL2
'''''        sql = sql & " " & .SXLGD_MS07LDL3 & ", "  ' SXLGD_測定値07 L/DL3
'''''        sql = sql & " " & .SXLGD_MS07LDL4 & ", "  ' SXLGD_測定値07 L/DL4
'''''        sql = sql & " " & .SXLGD_MS07LDL5 & ", "  ' SXLGD_測定値07 L/DL5
'''''        sql = sql & " " & .SXLGD_MS07DEN1 & ", "  ' SXLGD_測定値07 Den1
'''''        sql = sql & " " & .SXLGD_MS07DEN2 & ", "  ' SXLGD_測定値07 Den2
'''''        sql = sql & " " & .SXLGD_MS07DEN3 & ", "  ' SXLGD_測定値07 Den3
'''''        sql = sql & " " & .SXLGD_MS07DEN4 & ", "  ' SXLGD_測定値07 Den4
'''''        sql = sql & " " & .SXLGD_MS07DEN5 & ", "  ' SXLGD_測定値07 Den5
'''''        sql = sql & " " & .SXLGD_MS08LDL1 & ", "  ' SXLGD_測定値08 L/DL1
'''''        sql = sql & " " & .SXLGD_MS08LDL2 & ", "  ' SXLGD_測定値08 L/DL2
'''''        sql = sql & " " & .SXLGD_MS08LDL3 & ", "  ' SXLGD_測定値08 L/DL3
'''''        sql = sql & " " & .SXLGD_MS08LDL4 & ", "  ' SXLGD_測定値08 L/DL4
'''''        sql = sql & " " & .SXLGD_MS08LDL5 & ", "  ' SXLGD_測定値08 L/DL5
'''''        sql = sql & " " & .SXLGD_MS08DEN1 & ", "  ' SXLGD_測定値08 Den1
'''''        sql = sql & " " & .SXLGD_MS08DEN2 & ", "  ' SXLGD_測定値08 Den2
'''''        sql = sql & " " & .SXLGD_MS08DEN3 & ", "  ' SXLGD_測定値08 Den3
'''''        sql = sql & " " & .SXLGD_MS08DEN4 & ", "  ' SXLGD_測定値08 Den4
'''''        sql = sql & " " & .SXLGD_MS08DEN5 & ", "  ' SXLGD_測定値08 Den5
'''''        sql = sql & " " & .SXLGD_MS09LDL1 & ", "  ' SXLGD_測定値09 L/DL1
'''''        sql = sql & " " & .SXLGD_MS09LDL2 & ", "  ' SXLGD_測定値09 L/DL2
'''''        sql = sql & " " & .SXLGD_MS09LDL3 & ", "  ' SXLGD_測定値09 L/DL3
'''''        sql = sql & " " & .SXLGD_MS09LDL4 & ", "  ' SXLGD_測定値09 L/DL4
'''''        sql = sql & " " & .SXLGD_MS09LDL5 & ", "  ' SXLGD_測定値09 L/DL5
'''''        sql = sql & " " & .SXLGD_MS09DEN1 & ", "  ' SXLGD_測定値09 Den1
'''''        sql = sql & " " & .SXLGD_MS09DEN2 & ", " ' SXLGD_測定値09 Den2
'''''        sql = sql & " " & .SXLGD_MS09DEN3 & ", "  ' SXLGD_測定値09 Den3
'''''        sql = sql & " " & .SXLGD_MS09DEN4 & ", "  ' SXLGD_測定値09 Den4
'''''        sql = sql & " " & .SXLGD_MS09DEN5 & ", "  ' SXLGD_測定値09 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " " & .SXLGD_MS10LDL1 & ", "  ' SXLGD_測定値10 L/DL1
'''''        sql = sql & " " & .SXLGD_MS10LDL2 & ", "  ' SXLGD_測定値10 L/DL2
'''''        sql = sql & " " & .SXLGD_MS10LDL3 & ", "  ' SXLGD_測定値10 L/DL3
'''''        sql = sql & " " & .SXLGD_MS10LDL4 & ", "  ' SXLGD_測定値10 L/DL4
'''''        sql = sql & " " & .SXLGD_MS10LDL5 & ", "  ' SXLGD_測定値10 L/DL5
'''''        sql = sql & " " & .SXLGD_MS10DEN1 & ", "  ' SXLGD_測定値10 Den1
'''''        sql = sql & " " & .SXLGD_MS10DEN2 & ", "  ' SXLGD_測定値10 Den2
'''''        sql = sql & " " & .SXLGD_MS10DEN3 & ", "  ' SXLGD_測定値10 Den3
'''''        sql = sql & " " & .SXLGD_MS10DEN4 & ", "  ' SXLGD_測定値10 Den4
'''''        sql = sql & " " & .SXLGD_MS10DEN5 & ", "  ' SXLGD_測定値10 Den5
'''''        sql = sql & " " & .SXLGD_MS11LDL1 & ", "  ' SXLGD_測定値11 L/DL1
'''''        sql = sql & " " & .SXLGD_MS11LDL2 & ", "  ' SXLGD_測定値11 L/DL2
'''''        sql = sql & " " & .SXLGD_MS11LDL3 & ", "  ' SXLGD_測定値11 L/DL3
'''''        sql = sql & " " & .SXLGD_MS11LDL4 & ", " ' SXLGD_測定値11 L/DL4
'''''        sql = sql & " " & .SXLGD_MS11LDL5 & ", "  ' SXLGD_測定値11 L/DL5
'''''        sql = sql & " " & .SXLGD_MS11DEN1 & ", "  ' SXLGD_測定値11 Den1
'''''        sql = sql & " " & .SXLGD_MS11DEN2 & ", "  ' SXLGD_測定値11 Den2
'''''        sql = sql & " " & .SXLGD_MS11DEN3 & ", "  ' SXLGD_測定値11 Den3
'''''        sql = sql & " " & .SXLGD_MS11DEN4 & ", "  ' SXLGD_測定値11 Den4
'''''        sql = sql & " " & .SXLGD_MS11DEN5 & ", "  ' SXLGD_測定値11 Den5
'''''        sql = sql & " " & .SXLGD_MS12LDL1 & ", "  ' SXLGD_測定値12 L/DL1
'''''        sql = sql & " " & .SXLGD_MS12LDL2 & ", "  ' SXLGD_測定値12 L/DL2
'''''        sql = sql & " " & .SXLGD_MS12LDL3 & ", "  ' SXLGD_測定値12 L/DL3
'''''        sql = sql & " " & .SXLGD_MS12LDL4 & ", "  ' SXLGD_測定値12 L/DL4
'''''        sql = sql & " " & .SXLGD_MS12LDL5 & ", "  ' SXLGD_測定値12 L/DL5
'''''        sql = sql & " " & .SXLGD_MS12DEN1 & ", "  ' SXLGD_測定値12 Den1
'''''        sql = sql & " " & .SXLGD_MS12DEN2 & ", "  ' SXLGD_測定値12 Den2
'''''        sql = sql & " " & .SXLGD_MS12DEN3 & ", "  ' SXLGD_測定値12 Den3
'''''        sql = sql & " " & .SXLGD_MS12DEN4 & ", "  ' SXLGD_測定値12 Den4
'''''        sql = sql & " " & .SXLGD_MS12DEN5 & ", "  ' SXLGD_測定値12 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " " & .SXLGD_MS13LDL1 & ", "  ' SXLGD_測定値13 L/DL1
'''''        sql = sql & " " & .SXLGD_MS13LDL2 & ", "  ' SXLGD_測定値13 L/DL2
'''''        sql = sql & " " & .SXLGD_MS13LDL3 & ", "  ' SXLGD_測定値13 L/DL3
'''''        sql = sql & " " & .SXLGD_MS13LDL4 & ", "  ' SXLGD_測定値13 L/DL4
'''''        sql = sql & " " & .SXLGD_MS13LDL5 & ", "  ' SXLGD_測定値13 L/DL5
'''''        sql = sql & " " & .SXLGD_MS13DEN1 & ", "  ' SXLGD_測定値13 Den1
'''''        sql = sql & " " & .SXLGD_MS13DEN2 & ", "  ' SXLGD_測定値13 Den2
'''''        sql = sql & " " & .SXLGD_MS13DEN3 & ", "  ' SXLGD_測定値13 Den3
'''''        sql = sql & " " & .SXLGD_MS13DEN4 & ", "  ' SXLGD_測定値13 Den4
'''''        sql = sql & " " & .SXLGD_MS13DEN5 & ", "  ' SXLGD_測定値13 Den5
'''''        sql = sql & " " & .SXLGD_MS14LDL1 & ", "  ' SXLGD_測定値14 L/DL1
'''''        sql = sql & " " & .SXLGD_MS14LDL2 & ", "  ' SXLGD_測定値14 L/DL2
'''''        sql = sql & " " & .SXLGD_MS14LDL3 & ", "  ' SXLGD_測定値14 L/DL3
'''''        sql = sql & " " & .SXLGD_MS14LDL4 & ", "  ' SXLGD_測定値14 L/DL4
'''''        sql = sql & " " & .SXLGD_MS14LDL5 & ", "  ' SXLGD_測定値14 L/DL5
'''''        sql = sql & " " & .SXLGD_MS14DEN1 & ", "  ' SXLGD_測定値14 Den1
'''''        sql = sql & " " & .SXLGD_MS14DEN2 & ", "  ' SXLGD_測定値14 Den2
'''''        sql = sql & " " & .SXLGD_MS14DEN3 & ", "  ' SXLGD_測定値14 Den3
'''''        sql = sql & " " & .SXLGD_MS14DEN4 & ", "  ' SXLGD_測定値14 Den4
'''''        sql = sql & " " & .SXLGD_MS14DEN5 & ", "  ' SXLGD_測定値14 Den5
'''''        sql = sql & " " & .SXLGD_MS15LDL1 & ", "  ' SXLGD_測定値15 L/DL1
'''''        sql = sql & " " & .SXLGD_MS15LDL2 & ", "  ' SXLGD_測定値15 L/DL2
'''''        sql = sql & " " & .SXLGD_MS15LDL3 & ", "  ' SXLGD_測定値15 L/DL3
'''''        sql = sql & " " & .SXLGD_MS15LDL4 & ", "  ' SXLGD_測定値15 L/DL4
'''''        sql = sql & " " & .SXLGD_MS15LDL5 & ", "  ' SXLGD_測定値15 L/DL5
'''''        sql = sql & " " & .SXLGD_MS15DEN1 & ", "  ' SXLGD_測定値15 Den1
'''''        sql = sql & " " & .SXLGD_MS15DEN2 & ", "  ' SXLGD_測定値15 Den2
'''''        sql = sql & " " & .SXLGD_MS15DEN3 & ", "  ' SXLGD_測定値15 Den3
'''''        sql = sql & " " & .SXLGD_MS15DEN4 & ", "  ' SXLGD_測定値15 Den4
'''''        sql = sql & " " & .SXLGD_MS15DEN5 & ", "  ' SXLGD_測定値15 Den5
'''''#If PRNSQL > 0 Then
'''''sql = sql & vbCrLf
'''''#End If
'''''        sql = sql & " " & .SXLGD_MSRSDEN & ", "   ' SXLGD_測定結果 Den
'''''        sql = sql & " " & .SXLGD_MSRSLDL & ", "   ' SXLGD_測定結果 L/DL
'''''        sql = sql & " " & .SXLGD_MSRSDVD2 & ", "  ' SXLGD_測定結果 DVD2
'''''        sql = sql & " " & .SXLT_SMPPOS & ", "     ' SXLLTｻﾝﾌﾟﾙ測定位置（SXL位置情報）
'''''        sql = sql & " " & .SXLLT_MEASPEAK & ", "  ' SXLLT_測定値 ピーク値
'''''        sql = sql & " " & .SXLLT_MEAS1 & ", "     ' SXLLT_測定値1
'''''        sql = sql & " " & .SXLLT_MEAS2 & ", "     ' SXLLT_測定値2
'''''        sql = sql & " " & .SXLLT_MEAS3 & ", "     ' SXLLT_測定値3
'''''        sql = sql & " " & .SXLLT_MEAS4 & ", "     ' SXLLT_測定値4
'''''        sql = sql & " " & .SXLLT_MEAS5 & ", "     ' SXLLT_測定値5
'''''        sql = sql & " " & .SXLLT_CALCMEAS & ", "  ' SXLLT_計算結果
'''''        sql = sql & "sysdate, "
'''''        sql = sql & "'0', "
'''''        sql = sql & "sysdate, "
'''''        sql = sql & " " & .SXLOSF1_POS1 & ", "     'OSF1ﾊﾟﾀｰﾝ区分１位置
'''''        sql = sql & " " & .SXLOSF1_WID1 & ", "     'OSF1ﾊﾟﾀｰﾝ区分１幅
'''''        sql = sql & " '" & .SXLOSF1_RD1 & "', "      'OSF1ﾊﾟﾀｰﾝ区分１R/D
'''''        sql = sql & " " & .SXLOSF1_POS2 & ", "     'OSF1ﾊﾟﾀｰﾝ区分２位置
'''''        sql = sql & " " & .SXLOSF1_WID2 & ", "     'OSF1ﾊﾟﾀｰﾝ区分２幅
'''''        sql = sql & " '" & .SXLOSF1_RD2 & "', "      'OSF1ﾊﾟﾀｰﾝ区分２R/D
'''''        sql = sql & " " & .SXLOSF1_POS3 & ", "     'OSF1ﾊﾟﾀｰﾝ区分３位置
'''''        sql = sql & " " & .SXLOSF1_WID3 & ", "     'OSF1ﾊﾟﾀｰﾝ区分３幅
'''''        sql = sql & " '" & .SXLOSF1_RD3 & "', "      'OSF1ﾊﾟﾀｰﾝ区分３R/D
'''''        sql = sql & " " & .SXLOSF2_POS1 & ", "     'OSF2ﾊﾟﾀｰﾝ区分１位置
'''''        sql = sql & " " & .SXLOSF2_WID1 & ", "     'OSF2ﾊﾟﾀｰﾝ区分１幅
'''''        sql = sql & " '" & .SXLOSF2_RD1 & "', "      'OSF2ﾊﾟﾀｰﾝ区分１R/D
'''''        sql = sql & " " & .SXLOSF2_POS2 & ", "     'OSF2ﾊﾟﾀｰﾝ区分２位置
'''''        sql = sql & " " & .SXLOSF2_WID2 & ", "     'OSF2ﾊﾟﾀｰﾝ区分２幅
'''''        sql = sql & " '" & .SXLOSF2_RD2 & "', "      'OSF2ﾊﾟﾀｰﾝ区分２R/D
'''''        sql = sql & " " & .SXLOSF2_POS3 & ", "     'OSF2ﾊﾟﾀｰﾝ区分３位置
'''''        sql = sql & " " & .SXLOSF2_WID3 & ", "     'OSF2ﾊﾟﾀｰﾝ区分３幅
'''''        sql = sql & " '" & .SXLOSF2_RD3 & "', "      'OSF2ﾊﾟﾀｰﾝ区分３R/D
'''''        sql = sql & " " & .SXLOSF3_POS1 & ", "     'OSF3ﾊﾟﾀｰﾝ区分１位置
'''''        sql = sql & " " & .SXLOSF3_WID1 & ", "     'OSF3ﾊﾟﾀｰﾝ区分１幅
'''''        sql = sql & " '" & .SXLOSF3_RD1 & "', "      'OSF3ﾊﾟﾀｰﾝ区分１R/D
'''''        sql = sql & " " & .SXLOSF3_POS2 & ", "     'OSF3ﾊﾟﾀｰﾝ区分２位置
'''''        sql = sql & " " & .SXLOSF3_WID2 & ", "     'OSF3ﾊﾟﾀｰﾝ区分２幅
'''''        sql = sql & " '" & .SXLOSF3_RD2 & "', "      'OSF3ﾊﾟﾀｰﾝ区分２R/D
'''''        sql = sql & " " & .SXLOSF3_POS3 & ", "     'OSF3ﾊﾟﾀｰﾝ区分３位置
'''''        sql = sql & " " & .SXLOSF3_WID3 & ", "     'OSF3ﾊﾟﾀｰﾝ区分３幅
'''''        sql = sql & " '" & .SXLOSF3_RD3 & "', "      'OSF3ﾊﾟﾀｰﾝ区分３R/D
'''''        sql = sql & " " & .SXLOSF4_POS1 & ", "     'OSF4ﾊﾟﾀｰﾝ区分１位置
'''''        sql = sql & " " & .SXLOSF4_WID1 & ", "     'OSF4ﾊﾟﾀｰﾝ区分１幅
'''''        sql = sql & " '" & .SXLOSF4_RD1 & "', "      'OSF4ﾊﾟﾀｰﾝ区分１R/D
'''''        sql = sql & " " & .SXLOSF4_POS2 & ", "     'OSF4ﾊﾟﾀｰﾝ区分２位置
'''''        sql = sql & " " & .SXLOSF4_WID2 & ", "     'OSF4ﾊﾟﾀｰﾝ区分２幅
'''''        sql = sql & " '" & .SXLOSF4_RD2 & "', "      'OSF4ﾊﾟﾀｰﾝ区分２R/D
'''''        sql = sql & " " & .SXLOSF4_POS3 & ", "     'OSF4ﾊﾟﾀｰﾝ区分３位置
'''''        sql = sql & " " & .SXLOSF4_WID3 & ", "     'OSF4ﾊﾟﾀｰﾝ区分３幅
'''''        sql = sql & " '" & .SXLOSF4_RD3 & "', "      'OSF4ﾊﾟﾀｰﾝ区分３R/D
'''''        sql = sql & " " & .SXLGD_MS01DVD2 & ", "   'DVD2測定結果値１
'''''        sql = sql & " " & .SXLGD_MS02DVD2 & ", "   'DVD2測定結果値２
'''''        sql = sql & " " & .SXLGD_MS03DVD2 & ", "   'DVD2測定結果値３
'''''        sql = sql & " " & .SXLGD_MS04DVD2 & ", "   'DVD2測定結果値４
'''''        sql = sql & " " & .SXLGD_MS05DVD2 & ", "   'DVD2測定結果値５
'''''        sql = sql & " " & .SXLBMD1_MNBCR & ", "    'BMD1SXL計算結果面内分布
'''''        sql = sql & " " & .SXLBMD2_MNBCR & ", "    'BMD2SXL計算結果面内分布
'''''        sql = sql & " " & .SXLBMD3_MNBCR & ") "    'BMD3SXL計算結果面内分布
'''''    End With
'''''#If PRNSQL > 0 Then
'''''Debug.Print sql
'''''Stop
'''''#End If
'''''
'''''    If 0 >= OraDB.ExecuteSQL(sql) Then
'''''        DBDRV_scmzc_fcmkc001c_InsSoku = FUNCTION_RETURN_FAILURE
'''''        GoTo proc_exit
'''''    End If
'''''
'''''
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    DBDRV_scmzc_fcmkc001c_InsSoku = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


'概要      :総合判定 総合判定実績更新用ドライバ
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :Zis           ,I  ,typ_TBCMJ009 ,総合判定実績テーブルへの挿入用
'          :戻り値        ,O  ,FUNCTION_RETURN,読み込み成否
'説明      :
'履歴      :2001/06/27 蔵本 作成
Public Function DBDRV_scmzc_fcmkc001c_InsZis(Zis As typ_TBCMJ009) As FUNCTION_RETURN

    Dim sql As String
                                          
    '総合判定実績への挿入（TBCMJ009）

    ' 総合判定実績

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_InsZis"
    
    sql = "insert into TBCMJ009 ( "
    sql = sql & "CRYNUM, "           ' 結晶番号
    sql = sql & "INGOTPOS, "         ' インゴット内位置
    sql = sql & "TRANCNT, "          ' 処理回数
    sql = sql & "LENGTH, "           ' 長さ
    sql = sql & "KRPROCCD, "         ' 管理工程コード
    sql = sql & "PROCCODE, "         ' 工程コード
    sql = sql & "CODE, "             ' 区分コード
    sql = sql & "TSTAFFID, "         ' 登録社員ID
    sql = sql & "REGDATE, "          ' 登録日付
    sql = sql & "KSTAFFID, "         ' 更新社員ID
    sql = sql & "UPDDATE, "          ' 更新日付
    sql = sql & "SENDFLAG, "         ' 送信フラグ
    sql = sql & "SENDDATE )"         ' 送信日付
    
    sql = sql & " select "
    With Zis
    '工程コード設定ロジックの統一　2002/11/28 hama
    '工程コード設定
        nextCd = .PROCCODE
        sql = sql & " '" & .CRYNUM & "', "           ' 結晶番号
        sql = sql & " '" & .IngotPos & "', "         ' インゴット内位置
        sql = sql & " nvl(max(TRANCNT),0)+1, "       ' 処理回数
        sql = sql & " " & .LENGTH & ", "             ' 長さ
        sql = sql & " '" & .KRPROCCD & "', "         ' 管理工程コード
        sql = sql & " '" & nextCd & "', "            ' 工程コード
        sql = sql & " '" & .CODE & "', "             ' 区分コード
        sql = sql & " '" & .TSTAFFID & "', "         ' 登録社員ID
        sql = sql & "sysdate, "
        sql = sql & " '" & .TSTAFFID & "', "         ' 更新社員ID
        sql = sql & "sysdate, "
        sql = sql & "'0', "
        sql = sql & "sysdate "
        sql = sql & " from TBCMJ009 "
        sql = sql & " where CRYNUM  ='" & Zis.CRYNUM & "'"
        sql = sql & "   and INGOTPOS= " & Zis.IngotPos
    End With
    
    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_scmzc_fcmkc001c_InsZis = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_scmzc_fcmkc001c_InsZis = FUNCTION_RETURN_SUCCESS
    End If
    

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmkc001c_InsZis = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :総合判定 ブロック管理更新用（現在工程、最終通過工程）
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   ,説明
'          :Block         ,I  ,type_DBDRV_scmzc_fcmkc001c_UpdBlock1 ,ブロック管理の現在管理工程、現在工程、最終通過管理工程、最終通過工程更新用
'          :戻り値        ,O  ,FUNCTION_RETURN                      ,読み込み成否
'説明      :
'履歴      :2001/06/27 蔵本 作成
Public Function DBDRV_scmzc_fcmkc001c_UpdBlock1(Block As type_DBDRV_scmzc_fcmkc001c_UpdBlock1) As FUNCTION_RETURN

    Dim sql As String

    ' ブロック管理の更新

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_UpdBlock1"
    '工程コード設定ロジックの統一　2002/11/28 hama
    '工程コード設定
    nextCd = Block.NOWPROC
    nowCd = Block.LASTPASS
    
    sql = "update TBCME040 set "
    sql = sql & "  NOWPROC ='" & nextCd & "' "      '現在工程
    sql = sql & ", LASTPASS='" & nowCd & "' "       '最終通過工程
    sql = sql & ", UPDDATE =sysdate "
    sql = sql & ", SENDFLAG='0' "
    
    sql = sql & " where CRYNUM  ='" & Block.CRYNUM & "'"
    sql = sql & "   and INGOTPOS= " & Block.IngotPos
        
    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_scmzc_fcmkc001c_UpdBlock1 = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_scmzc_fcmkc001c_UpdBlock1 = FUNCTION_RETURN_SUCCESS
    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmkc001c_UpdBlock1 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function



'概要      :総合判定 結晶サンプル管理更新用（確定区分を１に更新）
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   ,説明
'          :CrySmp()      ,I  ,type_DBDRV_scmzc_fcmkc001c_UpdCrySmp ,結晶サンプル管理更新用
'          :戻り値        ,O  ,FUNCTION_RETURN                      ,読み込み成否
'説明      :
'履歴      :2001/07/26 蔵本 作成
Public Function DBDRV_scmzc_fcmkc001c_UpdCrySmp(CrySmp() As type_DBDRV_scmzc_fcmkc001c_UpdCrySmp) As FUNCTION_RETURN

    Dim sql As String
    Dim i As Long
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_UpdCrySmp"

    For i = 1 To UBound(CrySmp)
        ' 結晶サンプル管理の更新
'        sql = "update TBCME043 set "
'        sql = sql & "  KTKBN='1' "          '確定区分
'        sql = sql & ", UPDDATE=sysdate "
'        sql = sql & ", SENDFLAG='0' "
'        sql = sql & " where CRYNUM='" & CrySmp(i).CRYNUM & "' "
'        sql = sql & " and INGOTPOS=" & CrySmp(i).INGOTPOS & " "
'        sql = sql & " and SMPKBN='" & CrySmp(i).SMPKBN & "' "

        sql = "update XSDCS set "
        sql = sql & "  KTKBNCS='1' "          '確定区分
        sql = sql & ", KDAYCS =sysdate "
        sql = sql & ", SNDKCS ='0' "
        
        sql = sql & " where XTALCS  ='" & CrySmp(i).CRYNUM & "' "
        sql = sql & "   and INPOSCS = " & CrySmp(i).IngotPos & " "
        sql = sql & "   and SMPKBNCS='" & CrySmp(i).SMPKBN & "' "
        
        If 0 >= OraDB.ExecuteSQL(sql) Then
            DBDRV_scmzc_fcmkc001c_UpdCrySmp = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
     Next
     
     DBDRV_scmzc_fcmkc001c_UpdCrySmp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmkc001c_UpdCrySmp = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :総合判定 ブロック管理更新用（クリスタルカタログ、リメルト用）
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                           ,説明
'          :Block         ,I   ,type_DBDRV_scmzc_fcmkc001c_UpdBlock         ,ブロック管理の現在管理工程、現在工程、最終通過管理工程、最終通過工程更新用
'          :戻り値        ,O  ,FUNCTION_RETURN              ,
'説明      :
'履歴      :2001/06/27 蔵本 作成
Public Function DBDRV_fcmkc001c_UpdBlkCR(Block As typ_DBDRV_fcmkc001c_UpdBlkCR) As FUNCTION_RETURN

    Dim sql As String

    ' ブロック管理の更新

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_fcmkc001c_UpdBlkCR"
    '工程コード設定ロジックの統一　2002/11/28 hama
    '工程コード設定
    nextCd = Block.NOWPROC
    nowCd = PROCD_KESSYOU_SOUGOUHANTEI

    sql = "update TBCME040 set "
    sql = sql & " NOWPROC  ='" & nextCd & "' "              '現在工程
    sql = sql & ", LASTPASS='" & nowCd & "' "               '最終通過工程
    sql = sql & ", DELCLS  ='" & Block.DELCLS & "' "        '削除区分
    sql = sql & ", LSTATCLS='" & Block.LSTATCLS & "' "      '最終状態区分
    sql = sql & ", RSTATCLS='" & Block.RSTATCLS & "' "      '流動状態区分
    sql = sql & ", BDCAUS  ='" & Block.BDCAUS & "' "        '不良理由
    sql = sql & ", UPDDATE =sysdate "
    sql = sql & ", SENDFLAG='0' "
    sql = sql & " where CRYNUM  ='" & Block.CRYNUM & "'"
    sql = sql & "   and INGOTPOS= " & Block.IngotPos
        
    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_fcmkc001c_UpdBlkCR = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_fcmkc001c_UpdBlkCR = FUNCTION_RETURN_SUCCESS
    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_fcmkc001c_UpdBlkCR = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :上下の品番について検査項目_処を調べる
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
''''''          :CRYNUM        ,I  ,String       ,対象の結晶番号
''''''          :INGOTPOS      ,I  ,Integer      ,対象の位置
''''''          :smpShared     ,O  ,Boolean      ,サンプルは共用(上下品番とも Z/G/空 でないときのみTrueになりうる)
''''''          :WSpec(2)      ,O  ,Judg_Spec_Cry,上下品番の検査項目_処(1:上品番 2:下品番)
''''''説明      :総合判定測定値(TBCMJ014)に書き込むべき検査項目を調べるために利用する
''''''          :Z/G/空品番の場合は全て検査不要とする
''''''履歴      :2002/2/20 野村 作成
''''''概要      :
'''''Public Function GetHinbanSpec(CRYNUM As String, INGOTPOS As Integer, smpShared As Boolean, WSpec() As Judg_Spec_Cry) As FUNCTION_RETURN
'''''Dim sql$
'''''Dim rs As OraDynaset
'''''Dim i As Integer
'''''Dim loopTo As Integer
'''''
'''''    'エラーハンドラの設定
'''''    GetHinbanSpec = FUNCTION_RETURN_FAILURE
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function GetHinbanSpec"
'''''
'''''    '初期化
'''''    For i = 1 To 2
'''''        With WSpec(i)
'''''            .Enable = False
'''''            .rs = False
'''''            .Oi = False
'''''            .Cs = False
'''''            .Lt = False
'''''            .EPD = False
'''''            .B1 = False
'''''            .B2 = False
'''''            .B3 = False
'''''            .L1 = False
'''''            .L2 = False
'''''            .L3 = False
'''''            .L4 = False
'''''            .GD = False
'''''        End With
'''''    Next
'''''
'''''    '上下品番の検査項目_処を調べる
'''''    sql = "select HIN.INGOTPOS as HinFrom, HIN.INGOTPOS+HIN.LENGTH as HinTo"
'''''    sql = sql & ", E018.HSXRHWYS, E019.HSXONHWS, E019.HSXCNHWS, E019.HSXLTHWS"
'''''    sql = sql & ", E020.HSXOF1HS, E020.HSXOF2HS, E020.HSXOF3HS, E020.HSXOF4HS"
'''''    sql = sql & ", E020.HSXBM1HS, E020.HSXBM2HS, E020.HSXBM3HS"
'''''    sql = sql & ", E020.HSXDENHS, E020.HSXLDLHS, E020.HSXDVDHS "
'''''    sql = sql & "from TBCME041 HIN, TBCME018 E018, TBCME019 E019, TBCME020 E020 "
'''''    sql = sql & "where HIN.CRYNUM='" & CRYNUM & "'"
'''''    sql = sql & "  and HIN.INGOTPOS <= " & INGOTPOS
'''''    sql = sql & "  and HIN.INGOTPOS+HIN.LENGTH >= " & INGOTPOS
'''''    sql = sql & "  and HIN.HINBAN=E018.HINBAN and HIN.REVNUM=E018.MNOREVNO and HIN.FACTORY=E018.FACTORY and HIN.OPECOND=E018.OPECOND"
'''''    sql = sql & "  and HIN.HINBAN=E019.HINBAN and HIN.REVNUM=E019.MNOREVNO and HIN.FACTORY=E019.FACTORY and HIN.OPECOND=E019.OPECOND"
'''''    sql = sql & "  and HIN.HINBAN=E020.HINBAN and HIN.REVNUM=E020.MNOREVNO and HIN.FACTORY=E020.FACTORY and HIN.OPECOND=E020.OPECOND "
'''''    sql = sql & "order by HIN.INGOTPOS"
'''''    Set rs = OraDB.CreateDynaset(sql, ORADB_DEFAULT)
'''''
'''''    Do While Not rs.EOF
'''''        If rs("HinFrom") < INGOTPOS Then
'''''            '上品番の仕様
'''''            With WSpec(1)
'''''                .Enable = True
'''''                If ((rs("HSXRHWYS") = SIJI) Or (rs("HSXRHWYS") = SANKOU)) Then .rs = True
'''''                If ((rs("HSXONHWS") = SIJI) Or (rs("HSXONHWS") = SANKOU)) Then .Oi = True
'''''                If ((rs("HSXOF1HS") = SIJI) Or (rs("HSXOF1HS") = SANKOU)) Then .L1 = True
'''''                If ((rs("HSXOF2HS") = SIJI) Or (rs("HSXOF2HS") = SANKOU)) Then .L2 = True
'''''                If ((rs("HSXOF3HS") = SIJI) Or (rs("HSXOF3HS") = SANKOU)) Then .L3 = True
'''''                If ((rs("HSXOF4HS") = SIJI) Or (rs("HSXOF4HS") = SANKOU)) Then .L4 = True
'''''                If ((rs("HSXBM1HS") = SIJI) Or (rs("HSXBM1HS") = SANKOU)) Then .B1 = True
'''''                If ((rs("HSXBM2HS") = SIJI) Or (rs("HSXBM2HS") = SANKOU)) Then .B2 = True
'''''                If ((rs("HSXBM3HS") = SIJI) Or (rs("HSXBM3HS") = SANKOU)) Then .B3 = True
'''''                If ((rs("HSXDENHS") = SIJI) Or (rs("HSXDENHS") = SANKOU)) Or _
'''''                   ((rs("HSXLDLHS") = SIJI) Or (rs("HSXLDLHS") = SANKOU)) Or _
'''''                   ((rs("HSXDVDHS") = SIJI) Or (rs("HSXDVDHS") = SANKOU)) Then .GD = True
'''''                If ((rs("HSXCNHWS") = SIJI) Or (rs("HSXCNHWS") = SANKOU)) Then .Cs = True
'''''                If ((rs("HSXLTHWS") = SIJI) Or (rs("HSXLTHWS") = SANKOU)) Then .Lt = True
'''''                .EPD = True
'''''            End With
'''''        End If
'''''
'''''        If rs("HinTo") > INGOTPOS Then
'''''            '下品番の仕様
'''''            With WSpec(2)
'''''                .Enable = True
'''''                If ((rs("HSXRHWYS") = SIJI) Or (rs("HSXRHWYS") = SANKOU)) Then .rs = True
'''''                If ((rs("HSXONHWS") = SIJI) Or (rs("HSXONHWS") = SANKOU)) Then .Oi = True
'''''                If ((rs("HSXOF1HS") = SIJI) Or (rs("HSXOF1HS") = SANKOU)) Then .L1 = True
'''''                If ((rs("HSXOF2HS") = SIJI) Or (rs("HSXOF2HS") = SANKOU)) Then .L2 = True
'''''                If ((rs("HSXOF3HS") = SIJI) Or (rs("HSXOF3HS") = SANKOU)) Then .L3 = True
'''''                If ((rs("HSXOF4HS") = SIJI) Or (rs("HSXOF4HS") = SANKOU)) Then .L4 = True
'''''                If ((rs("HSXBM1HS") = SIJI) Or (rs("HSXBM1HS") = SANKOU)) Then .B1 = True
'''''                If ((rs("HSXBM2HS") = SIJI) Or (rs("HSXBM2HS") = SANKOU)) Then .B2 = True
'''''                If ((rs("HSXBM3HS") = SIJI) Or (rs("HSXBM3HS") = SANKOU)) Then .B3 = True
'''''                If ((rs("HSXDENHS") = SIJI) Or (rs("HSXDENHS") = SANKOU)) Or _
'''''                   ((rs("HSXLDLHS") = SIJI) Or (rs("HSXLDLHS") = SANKOU)) Or _
'''''                   ((rs("HSXDVDHS") = SIJI) Or (rs("HSXDVDHS") = SANKOU)) Then .GD = True
'''''                If ((rs("HSXCNHWS") = SIJI) Or (rs("HSXCNHWS") = SANKOU)) Then .Cs = True
'''''                If ((rs("HSXLTHWS") = SIJI) Or (rs("HSXLTHWS") = SANKOU)) Then .Lt = True
'''''                .EPD = True
'''''            End With
'''''        End If
'''''
'''''        rs.MoveNext
'''''    Loop
'''''    rs.Close
'''''    Set rs = Nothing
'''''
'''''    '共用サンプルであるかを調べる
'''''    If WSpec(1).Enable And WSpec(2).Enable Then
''''''       sql = "select count(*) as SMPCNT from TBCME043 where CRYNUM='" & CRYNUM & "' and INGOTPOS=" & INGOTPOS
'''''        sql = "select count(*) as SMPCNT from XSDCS where  XTALCS='" & CRYNUM & "' and INPOSCS=" & INGOTPOS
'''''
'''''        Set rs = OraDB.CreateDynaset(sql, ORADB_DEFAULT)
'''''        If rs.RecordCount > 0 Then
'''''            If rs("SMPCNT") = 1 Then
'''''                smpShared = True
'''''            End If
'''''        End If
'''''        rs.Close
'''''        Set rs = Nothing
'''''    End If
'''''
'''''    GetHinbanSpec = FUNCTION_RETURN_SUCCESS
'''''
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    GetHinbanSpec = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


''''''概要      :GD仕様を取得する
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                               ,説明
''''''          :Gd_Siyou()    ,   ,type_DBDRV_scmzc_fcmkc001c_Siyou ,
''''''          :BLOCKID       ,   ,String                           ,
''''''          :戻り値        ,O  ,FUNCTION_RETURN                  ,
''''''説明      :
''''''履歴      :
'''''Public Function getGDsiyou(Gd_Siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou, _
'''''                           BLOCKID As String) As FUNCTION_RETURN
'''''
'''''    Dim sql As String
'''''    Dim rs As OraDynaset
'''''    Dim recCnt As Integer
'''''    Dim i As Long
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function getGDsiyou"
'''''
'''''    getGDsiyou = FUNCTION_RETURN_SUCCESS
'''''
'''''    sql = sql & "select "
'''''    sql = sql & "E020.HINBAN,"                   ' 品番
'''''    sql = sql & "HSXDENMX,"                     ' 品ＳＸＤｅｎ上限
'''''    sql = sql & "HSXDENMN,"                     ' 品ＳＸＤｅｎ下限
'''''    sql = sql & "HSXLDLMX,"                     ' 品ＳＸＬ／ＤＬ上限
'''''    sql = sql & "HSXLDLMN,"                     ' 品ＳＸＬ／ＤＬ下限
'''''    sql = sql & "HSXDVDMXN,"                     ' 品ＳＸＤＶＤ２上限   項目追加，修正対応 2003.05.20 yakimura
'''''    sql = sql & "HSXDVDMNN,"                     ' 品ＳＸＤＶＤ２下限   項目追加，修正対応 2003.05.20 yakimura
'''''    sql = sql & "HSXDENHT, "                    ' 品ＳＸＤｅｎ保証方法＿対
'''''    sql = sql & "HSXDENHS,"                     ' 品ＳＸＤｅｎ保証方法＿処
'''''    sql = sql & "HSXLDLHT,"                     ' 品ＳＸＬ／ＤＬ保証方法＿対
'''''    sql = sql & "HSXLDLHS,"                     ' 品ＳＸＬ／ＤＬ保証方法＿処
'''''    sql = sql & "HSXDVDHT,"                     ' 品ＳＸＤＶＤ２保証方法＿対
'''''    sql = sql & "HSXDVDHS,"                     ' 品ＳＸＤＶＤ２保証方法＿処
'''''    sql = sql & "HSXDENKU,"                     ' 品ＳＸＤｅｎ検査有無
'''''    sql = sql & "HSXDVDKU,"                     ' 品ＳＸＤＶＤ２検査有無
'''''    sql = sql & "HSXLDLKU "                      ' 品ＳＸＬ／ＤＬ検査有無
'''''    sql = sql & "from TBCME020 E020,TBCME041 E041,TBCME040  E040 "
'''''    sql = sql & "where E040.BLOCKID = '" & BLOCKID & "'"
'''''    sql = sql & "   and  E041.CRYNUM = E040.CRYNUM"
'''''    sql = sql & "   and E040.INGOTPOS<=E041.INGOTPOS"
'''''    sql = sql & "   and E041.INGOTPOS< E040.INGOTPOS+E040.LENGTH"
'''''    sql = sql & "   and E041.HINBAN = E020.HINBAN"
'''''    sql = sql & "   and E041.OPECOND = E020.OPECOND"
'''''
'''''
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    If rs.RecordCount = 0 Then
'''''        getGDsiyou = FUNCTION_RETURN_FAILURE
'''''        ReDim Gd_Siyou(0)
'''''        rs.Close
'''''        GoTo proc_exit
'''''    End If
'''''
'''''
'''''    recCnt = rs.RecordCount
'''''    ReDim Gd_Siyou(recCnt)
'''''
'''''    For i = 1 To recCnt
'''''        With Gd_Siyou(i)
'''''
'''''            .hin.hinban = rs("HINBAN")            ' 品番
'''''            .HSXDENMX = rs("HSXDENMX")            ' 品ＳＸＤｅｎ上限
'''''            .HSXDENMN = rs("HSXDENMN")             ' 品ＳＸＤｅｎ下限
'''''            .HSXLDLMX = rs("HSXLDLMX")            ' 品ＳＸＬ／ＤＬ上限
'''''            .HSXLDLMN = rs("HSXLDLMN")            ' 品ＳＸＬ／ＤＬ下限
'''''            .HSXDVDMX = rs("HSXDVDMXN")            ' 品ＳＸＤＶＤ２上限   項目追加，修正対応 2003.05.20 yakimura
'''''            .HSXDVDMN = rs("HSXDVDMNN")            ' 品ＳＸＤＶＤ２下限   項目追加，修正対応 2003.05.20 yakimura
'''''            .HSXDENHT = rs("HSXDENHT")            ' 品ＳＸＤｅｎ保証方法＿対
'''''            .HSXDENHS = rs("HSXDENHS")            ' 品ＳＸＤｅｎ保証方法＿処
'''''            .HSXLDLHT = rs("HSXLDLHT")            ' 品ＳＸＬ／ＤＬ保証方法＿対
'''''            .HSXLDLHS = rs("HSXLDLHS")            ' 品ＳＸＬ／ＤＬ保証方法＿処
'''''            .HSXDVDHT = rs("HSXDVDHT")            ' 品ＳＸＤＶＤ２保証方法＿対
'''''            .HSXDVDHS = rs("HSXDVDHS")            ' 品ＳＸＤＶＤ２保証方法＿処
'''''            .HSXDENKU = rs("HSXDENKU")            ' 品ＳＸＤｅｎ検査有無
'''''            .HSXDVDKU = rs("HSXDVDKU")            ' 品ＳＸＤＶＤ２検査有無
'''''            .HSXLDLKU = rs("HSXLDLKU")            ' 品ＳＸＬ／ＤＬ検査有無
'''''        End With
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


''''''概要      :加工実績判定に構造体に値をセットする
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型             ,説明
''''''          :BLOCKID       ,   ,String         ,ブロックID
''''''          :Kakou         ,   ,type_KakouJudg ,加工実績判定構造体
''''''          :戻り値        ,O  ,FUNCTION_RETURN,
''''''説明      :ブロック内全品番の仕様と実績を求める
''''''履歴      :2002/4/16 佐野 作成
'''''Public Function DBDRV_scmzc_fcmkc001c_Kakou(BLOCKID As String, Kakou As type_KakouJudg) As FUNCTION_RETURN
'''''    Dim sql As String
'''''    Dim sql1 As String
'''''    Dim rs As OraDynaset
'''''    Dim recCnt As Integer
'''''    Dim c0 As Integer
'''''    Dim tHIN() As tFullHinban
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo PROC_ERR
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_Kakou"
'''''
'''''    DBDRV_scmzc_fcmkc001c_Kakou = FUNCTION_RETURN_FAILURE
'''''
'''''    'ブロック内の全品番を求める
'''''    sql = "select HINBAN, REVNUM, FACTORY, OPECOND from TBCME040 E40, TBCME041 E41 "
'''''    sql = sql & "Where E41.CRYNUM = E40.CRYNUM and "
'''''    sql = sql & "E40.BLOCKID = '" & BLOCKID & "' and "
'''''    sql = sql & "E40.INGOTPOS < E41.INGOTPOS+E41.LENGTH and "
'''''    sql = sql & "E40.INGOTPOS+E40.LENGTH > E41.INGOTPOS"
'''''
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    recCnt = rs.RecordCount
'''''    ReDim tHIN(recCnt)
'''''    If recCnt = 0 Then
'''''        rs.Close
'''''        GoTo PROC_EXIT
'''''    End If
'''''    For c0 = 1 To recCnt
'''''        tHIN(c0).hinban = rs("HINBAN")
'''''        tHIN(c0).mnorevno = rs("REVNUM")
'''''        tHIN(c0).factory = rs("FACTORY")
'''''        tHIN(c0).opecond = rs("OPECOND")
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''
'''''    '求めた全品番の加工仕様を求める
'''''    If scmzc_getKakouSpec(tHIN(), Kakou.Spec()) = FUNCTION_RETURN_FAILURE Then
'''''        GoTo PROC_EXIT
'''''    End If
'''''
'''''    '対象ブロックの加工実績を求める
'''''    If scmzc_getKakouJiltuseki(BLOCKID, Kakou.Jiltuseki) = FUNCTION_RETURN_FAILURE Then
'''''        GoTo PROC_EXIT
'''''    End If
'''''
'''''    DBDRV_scmzc_fcmkc001c_Kakou = FUNCTION_RETURN_SUCCESS
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
'''''    gErr.HandleError
'''''    Resume PROC_EXIT
'''''End Function


''''''概要      :ブロック内の品番についてCs仕様の有無をチェックする
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
''''''          :crynum        ,I  ,String    ,結晶番号
''''''          :blkFrom       ,I  ,Integer   ,ブロック開始位置
''''''          :blkTo         ,I  ,Integer   ,ブロック終了位置
''''''          :hasCs         ,O  ,String    ,ブロック内の品番にCs仕様を持つものがあるか(='H':保証あり ='S':参考あり 他:仕様なし)
''''''          :hasCsFromTo   ,O  ,String    ,ブロック内の品番にFromToのCs仕様を持つものがあるか(='H':保証あり ='S':参考あり 他:仕様なし)
''''''          :戻り値        ,O  ,FUNCTION_RETURN,
''''''説明      :判定画面のTop側/Bot側について表示・判定を行うかどうかを決定するために利用する
''''''履歴      :2002/4/16 野村 作成
'''''Public Function DBDRV_scmzc_fcmkc001c_CheckSpecCs(CRYNUM As String, BlkFrom As Integer, BlkTo As Integer, jCs As String, jCsFromTo As String) As FUNCTION_RETURN
'''''Dim sql As String
'''''Dim rs As OraDynaset
'''''Dim HSXCNHWS As String
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo PROC_ERR
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_CheckSpecCs"
'''''
'''''    DBDRV_scmzc_fcmkc001c_CheckSpecCs = FUNCTION_RETURN_FAILURE
'''''
'''''    jCs = " "
'''''    jCsFromTo = " "
'''''    sql = "select HSXCNHWS, HSXCNMIN "
'''''    sql = sql & "from TBCME041 HIN, TBCME019 SPEC "
'''''    sql = sql & "where HIN.CRYNUM='" & CRYNUM & "'"
'''''    sql = sql & "  and HIN.INGOTPOS<" & BlkTo
'''''    sql = sql & "  and HIN.INGOTPOS+LENGTH>" & BlkFrom
'''''    sql = sql & "  and SPEC.HINBAN=HIN.HINBAN"
'''''    sql = sql & "  and SPEC.MNOREVNO=HIN.REVNUM"
'''''    sql = sql & "  and SPEC.FACTORY=HIN.FACTORY"
'''''    sql = sql & "  and SPEC.OPECOND=HIN.OPECOND"
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    '品番毎にCs仕様・CsFromTo仕様の'H','S'を調べる
'''''    Do While rs.EOF = False
'''''        HSXCNHWS = rs("HSXCNHWS")
'''''        If HSXCNHWS = SIJI Then
'''''            jCs = HSXCNHWS
'''''            If rs("HSXCNMIN") > 0# Then
'''''                jCsFromTo = HSXCNHWS
'''''            End If
'''''        ElseIf HSXCNHWS = SANKOU Then
'''''            If jCs <> SIJI Then jCs = HSXCNHWS
'''''            If rs("HSXCNMIN") > 0# Then
'''''                If jCsFromTo <> SIJI Then jCsFromTo = HSXCNHWS
'''''            End If
'''''        End If
'''''        rs.MoveNext
'''''    Loop
'''''    rs.Close
'''''    Set rs = Nothing
'''''    DBDRV_scmzc_fcmkc001c_CheckSpecCs = FUNCTION_RETURN_SUCCESS
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
'''''    gErr.HandleError
'''''    Resume PROC_EXIT
'''''End Function


''''''概要      :ブロック内の品番についてCs仕様を取得する('H'or'S'のもののみ)
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
''''''          :crynum        ,   ,String    ,
''''''          :blkFrom       ,   ,Integer   ,
''''''          :blkTo         ,   ,Integer   ,
''''''          :SpecCs()      ,   ,C_Cs      ,
''''''          :戻り値        ,O  ,FUNCTION_R,
''''''説明      :
''''''履歴      :
'''''Public Function DBDRV_scmzc_fcmkc001c_GetSpecCs(CRYNUM As String, BlkFrom As Integer, BlkTo As Integer, SpecCs() As C_Cs) As FUNCTION_RETURN
'''''Dim sql As String
'''''Dim rs As OraDynaset
'''''Dim recCnt As Long
'''''Dim i As Long
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_GetSpecCs"
'''''
'''''    DBDRV_scmzc_fcmkc001c_GetSpecCs = FUNCTION_RETURN_FAILURE
'''''    sql = "select HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT,HSXCNHWS, HSXCNMIN, HSXCNMAX "
'''''    sql = sql & "from TBCME041 HIN, TBCME019 SPEC "
'''''    sql = sql & "where HIN.CRYNUM='" & CRYNUM & "'"
'''''    sql = sql & "  and HIN.INGOTPOS<" & BlkTo
''''''野村氏の指示により変更
''''''2002/05/11 S.Sano    sql = sql & "  and HIN.INGOTPOS+LENGTH>=" & BlkFrom
'''''    sql = sql & "  and HIN.INGOTPOS+LENGTH>" & BlkFrom '2002/05/11 S.Sano
'''''    sql = sql & "  and SPEC.HINBAN=HIN.HINBAN"
'''''    sql = sql & "  and SPEC.MNOREVNO=HIN.REVNUM"
'''''    sql = sql & "  and SPEC.FACTORY=HIN.FACTORY"
'''''    sql = sql & "  and SPEC.OPECOND=HIN.OPECOND"
'''''    sql = sql & "  and SPEC.HSXCNHWS in ('H','S')"
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    recCnt = rs.RecordCount
'''''    If recCnt = 0 Then
'''''        ReDim SpecCs(0)
'''''    Else
'''''        ReDim SpecCs(1 To recCnt)
'''''        For i = 1 To recCnt
'''''            With SpecCs(i)
'''''                .GuaranteeCs.cMeth = rs("HSXCNSPH")
'''''                .GuaranteeCs.cCount = rs("HSXCNSPT")
'''''                .GuaranteeCs.cPos = rs("HSXCNSPI")
'''''                .GuaranteeCs.cObj = rs("HSXCNHWT")
'''''                .GuaranteeCs.cJudg = rs("HSXCNHWS")
'''''                .SpecCsMin = rs("HSXCNMIN")
'''''                .SpecCsMax = rs("HSXCNMAX")
'''''            End With
'''''            rs.MoveNext
'''''        Next
'''''    End If
'''''    rs.Close
'''''    Set rs = Nothing
'''''
'''''    DBDRV_scmzc_fcmkc001c_GetSpecCs = FUNCTION_RETURN_SUCCESS
'''''
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


''''''概要      :ブロック内の品番についてLt仕様を取得する('H'or'S'のもののみ)
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
''''''          :crynum        ,   ,String    ,
''''''          :blkFrom       ,   ,Integer   ,
''''''          :blkTo         ,   ,Integer   ,
''''''          :SpecLt()      ,   ,C_Lt      ,
''''''          :戻り値        ,O  ,FUNCTION_R,
''''''説明      :
''''''履歴      :
'''''Public Function DBDRV_scmzc_fcmkc001c_GetSpecLt(CRYNUM As String, BlkFrom As Integer, BlkTo As Integer, SpecLt() As C_LT) As FUNCTION_RETURN
'''''Dim sql As String
'''''Dim rs As OraDynaset
'''''Dim recCnt As Long
'''''Dim i As Long
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_GetSpecLt"
'''''
'''''    DBDRV_scmzc_fcmkc001c_GetSpecLt = FUNCTION_RETURN_FAILURE
'''''    sql = "select HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT,HSXLTHWS, HSXLTMIN, HSXLTMAX "
'''''    sql = sql & "from TBCME041 HIN, TBCME019 SPEC "
'''''    sql = sql & "where HIN.CRYNUM='" & CRYNUM & "'"
'''''    sql = sql & "  and HIN.INGOTPOS<" & BlkTo
'''''    sql = sql & "  and HIN.INGOTPOS+LENGTH>=" & BlkFrom
'''''    sql = sql & "  and SPEC.HINBAN=HIN.HINBAN"
'''''    sql = sql & "  and SPEC.MNOREVNO=HIN.REVNUM"
'''''    sql = sql & "  and SPEC.FACTORY=HIN.FACTORY"
'''''    sql = sql & "  and SPEC.OPECOND=HIN.OPECOND"
'''''    sql = sql & "  and SPEC.HSXLTHWS in ('H','S')"
'''''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    recCnt = rs.RecordCount
'''''    If recCnt = 0 Then
'''''        ReDim SpecLt(0)
'''''    Else
'''''        ReDim SpecLt(1 To recCnt)
'''''        For i = 1 To recCnt
'''''            With SpecLt(i)
'''''                .GuaranteeLt.cMeth = rs("HSXLTSPH")
'''''                .GuaranteeLt.cCount = rs("HSXLTSPT")
'''''                .GuaranteeLt.cPos = rs("HSXLTSPI")
'''''                .GuaranteeLt.cObj = rs("HSXLTHWT")
'''''                .GuaranteeLt.cJudg = rs("HSXLTHWS")
'''''                .SpecLtMin = rs("HSXLTMIN")
'''''                .SpecLtMax = rs("HSXLTMAX")
'''''            End With
'''''            rs.MoveNext
'''''        Next
'''''    End If
'''''    rs.Close
'''''    Set rs = Nothing
'''''
'''''    DBDRV_scmzc_fcmkc001c_GetSpecLt = FUNCTION_RETURN_SUCCESS
'''''
'''''proc_exit:
'''''    '終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    'エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    gErr.HandleError
'''''    Resume proc_exit
'''''End Function


''''''概要      :ブロック内にLT「保証」はあるか？なければ「参考」はあるか？
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型          ,説明
''''''          :Crynum        ,I  ,String      ,結晶番号
''''''          :BlkFrom       ,I  ,Integer     ,ブロックの開始位置
''''''          :BlkTo         ,I  ,Integer     ,ブロックの終了位置
''''''          :戻り値        ,O  ,String      ,'H':保証あり 'S':参考あり vbNullString:なし
''''''説明      :
''''''履歴      :2002/10/10 野村 作成
'''''Public Function DBDRV_getLtGuaranteeInBlock(CRYNUM As String, BlkFrom As Integer, BlkTo As Integer) As String
'''''Dim sql As String
'''''Dim rs As OraDynaset
'''''
'''''    'エラーハンドラの設定
'''''    On Error GoTo PROC_ERR
'''''    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_getLtGuaranteeInBlock"
'''''    DBDRV_getLtGuaranteeInBlock = vbNullString
'''''
'''''    sql = "select SIYO.HSXLTHWS "
'''''    sql = sql & "from TBCME041 HIN, TBCME019 SIYO "
'''''    sql = sql & "where HIN.CRYNUM='" & CRYNUM & "'"
'''''    sql = sql & "  and HIN.INGOTPOS<" & BlkTo & " and HIN.INGOTPOS+HIN.LENGTH>" & BlkFrom
'''''    sql = sql & "  and SIYO.HINBAN=HIN.HINBAN and SIYO.MNOREVNO=HIN.REVNUM and SIYO.FACTORY=HIN.FACTORY and SIYO.OPECOND=HIN.OPECOND"
'''''    sql = sql & "  and SIYO.HSXLTHWS in ('H','S') "
'''''    sql = sql & "order by SIYO.HSXLTHWS"
'''''    sql = "select * from (" & sql & ") where rownum=1"
'''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    If rs.RecordCount > 0 Then
'''''        DBDRV_getLtGuaranteeInBlock = rs("HSXLTHWS")
'''''    Else
'''''        DBDRV_getLtGuaranteeInBlock = vbNullString
'''''    End If
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
'''''    gErr.HandleError
'''''    Resume PROC_EXIT
'''''End Function
'''''============================================================================================================================


'概要      :推定算出データの登録
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                       ,説明
'          :RS            ,I  ,typ_TBCMJ002             ,結晶抵抗実績テーブルへの挿入用
'          :戻り値        ,O  ,FUNCTION_RETURN          ,成否
'説明      :
'履歴      :2001/06/27 蔵本 作成
Public Function DBDRV_SuiteiZis_InsRS(rs() As typ_TBCMJ002) As FUNCTION_RETURN

    Dim lcnt    As Integer
    Dim sql     As String
                                          
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc030_SQL.bas -- Function DBDRV_SuiteiZis_InsRS"
    
    '推定算出データを結晶抵抗実績テーブル(TBCMJ002)へ登録する。
    For lcnt = 1 To UBound(rs)
        With rs(lcnt)
            sql = "insert into TBCMJ002 ("
            sql = sql & "CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, KRPROCCD, PROCCODE, "
            sql = sql & "HINBAN, REVNUM, factory, opecond, GOUKI, TYPE, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, "
            sql = sql & "EFEHS, RRG, JudgData, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE, SUIFLG) "
            
            sql = sql & "values ('"
            sql = sql & .CRYNUM & "', "             '結晶番号
            sql = sql & .POSITION & ", '"           '位置
            sql = sql & .SMPKBN & "', '"            'サンプル区分"
            sql = sql & .TRANCOND & "', "           '処理条件
            sql = sql & .TRANCNT & ", "             '処理回数
            sql = sql & .SMPLNO & ", '"             'サンプル№
            sql = sql & .SMPLUMU & "', '"           'サンプル有無
            sql = sql & .KRPROCCD & "', '"          '管理工程コード
            sql = sql & .PROCCODE & "', '"          '工程コード
            sql = sql & .hinban & "', "             '品番
            sql = sql & .REVNUM & ", '"             '製品番号改訂番号
            sql = sql & .factory & "', '"           '工場
            sql = sql & .opecond & "', '"           '操業条件
            sql = sql & .GOUKI & "', '"             '号機
            sql = sql & .Type & "', "               'タイプ
            sql = sql & .MEAS1 & ", "               '測定値１
            sql = sql & .MEAS2 & ", "               '測定値２
            sql = sql & .MEAS3 & ", "               '測定値３
            sql = sql & .MEAS4 & ", "               '測定値４
            sql = sql & .MEAS5 & ", "               '測定値５
            sql = sql & .EFEHS & ", "               '実行偏析
            sql = sql & .RRG & ", "                 'ＲＲＧ
            sql = sql & .JudgData & ", '"           '検索対象値
            sql = sql & .TSTAFFID & "', "           '登録社員ID
            sql = sql & "sysdate, '"                '登録日付
            sql = sql & .KSTAFFID & "', "           '更新社員ID
            sql = sql & "sysdate, '"                '更新日付
            sql = sql & .SENDFLAG & "', "           '送信フラグ
            sql = sql & "sysdate, '"                '送信日付
            sql = sql & .SUIFLG & "')"              '推定FLG"
        End With
        
        '' WriteDBLog sql
        If OraDB.ExecuteSQL(sql) <= 0 Then GoTo proc_err
    
    Next lcnt

    DBDRV_SuiteiZis_InsRS = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_SuiteiZis_InsRS = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


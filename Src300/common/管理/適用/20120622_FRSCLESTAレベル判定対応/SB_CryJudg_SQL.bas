Attribute VB_Name = "SB_CryJudg_SQL"
Option Explicit

'品番、仕様、結晶内側取得用(TOP,TAIL順で２レコード取得)
Public Type type_DBDRV_scmzc_fcmkc001c_Siyou
    
    'ブロック管理
    CRYNUM      As String * 12        ' 結晶番号
    INGOTPOS    As Integer            ' 結晶内開始位置
    Length      As Integer            ' 長さ
    
    '品番管理
    HIN As tFullHinban                ' 品番(full)
        
    '結晶情報
    PRODCOND    As String * 4         ' 製作条件
    PGID        As String * 8         ' ＰＧ－ＩＤ
    UPLENGTH    As Integer            ' 引上げ長さ
    FREELENG    As Integer            ' フリー長
    DIAMETER    As Integer            ' 直径 2002/05/01 S.Sano
    CHARGE      As Double             ' チャージ量
    SEED        As String * 4         ' シード
    ADDDPPOS    As Integer            ' 追加ドープ位置

    '製品仕様
    HSXTYPE  As String * 1            ' 品ＳＸタイプ
    HSXD1CEN As Double                ' 品ＳＸ直径１中心
    HSXCDIR  As String * 1            ' 品ＳＸ結晶面方位
    HSXRMIN  As Double                ' 品ＳＸ比抵抗下限
    HSXRMAX  As Double                ' 品ＳＸ比抵抗上限
    HSXRAMIN As Double                ' 品ＳＸ比抵抗平均下限
    HSXRAMAX As Double                ' 品ＳＸ比抵抗平均上限
    HSXRMCAL As String * 1            ' 品ＳＸ比抵抗面内計算　　　　'' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06
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
    HSXONMCL As String * 1            ' 品ＳＸ酸素濃度面内計算　　　　'' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06
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

    HSXOF1AX As Double                ' 品ＳＸＯＳＦ１平均上限
    HSXOF1MX As Double                ' 品ＳＸＯＳＦ１上限
    HSXOF2AX As Double                ' 品ＳＸＯＳＦ２平均上限
    HSXOF2MX As Double                ' 品ＳＸＯＳＦ２上限
    HSXOF3AX As Double                ' 品ＳＸＯＳＦ３平均上限
    HSXOF3MX As Double                ' 品ＳＸＯＳＦ３上限
    HSXOF4AX As Double                ' 品ＳＸＯＳＦ４平均上限
    HSXOF4MX As Double                ' 品ＳＸＯＳＦ４上限
    HSXOF1SH As String * 1            ' 品ＳＸＯＳＦ１測定位置＿方
    HSXOF1ST As String * 1            ' 品ＳＸＯＳＦ１測定位置＿点
    HSXOF1SR As String * 1            ' 品ＳＸＯＳＦ１測定位置＿領
    HSXOF1HT As String * 1            ' 品ＳＸＯＳＦ１保証方法＿対
    HSXOF1HS As String * 1            ' 品ＳＸＯＳＦ１保証方法＿処
    HSXOF2SH As String * 1            ' 品ＳＸＯＳＦ２測定位置＿方
    HSXOF2ST As String * 1            ' 品ＳＸＯＳＦ２測定位置＿点
    HSXOF2SR As String * 1            ' 品ＳＸＯＳＦ２測定位置＿領
    HSXOF2HT As String * 1            ' 品ＳＸＯＳＦ２保証方法＿対
    HSXOF2HS As String * 1            ' 品ＳＸＯＳＦ２保証方法＿処
    HSXOF3SH As String * 1            ' 品ＳＸＯＳＦ３測定位置＿方
    HSXOF3ST As String * 1            ' 品ＳＸＯＳＦ３測定位置＿点
    HSXOF3SR As String * 1            ' 品ＳＸＯＳＦ３測定位置＿領
    HSXOF3HT As String * 1            ' 品ＳＸＯＳＦ３保証方法＿対
    HSXOF3HS As String * 1            ' 品ＳＸＯＳＦ３保証方法＿処
    HSXOF4SH As String * 1            ' 品ＳＸＯＳＦ４測定位置＿方
    HSXOF4ST As String * 1            ' 品ＳＸＯＳＦ４測定位置＿点
    HSXOF4SR As String * 1            ' 品ＳＸＯＳＦ４測定位置＿領
    HSXOF4HT As String * 1            ' 品ＳＸＯＳＦ４保証方法＿対
    HSXOF4HS As String * 1            ' 品ＳＸＯＳＦ４保証方法＿処
    HSXOF1NS As String * 2            ' 品ＳＸＯＳＦ１熱処理法
    HSXOF2NS As String * 2            ' 品ＳＸＯＳＦ２熱処理法
    HSXOF3NS As String * 2            ' 品ＳＸＯＳＦ３熱処理法
    HSXOF4NS As String * 2            ' 品ＳＸＯＳＦ４熱処理法
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
    HSXCNKHI As String * 1            ' 品ＳＸ炭素濃度検査頻度＿位 09/01/08 ooba

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
''Add Start 2011/07/22 LT10Ω判定追加対応 T.Koi(SETsw)
    HSXLT10MIN As Integer             ' 品ＳＸＬタイム10Ω換算下限値
''Add End   2011/07/22 LT10Ω判定追加対応 T.Koi(SETsw)
    HSXLTSPH As String * 1            ' 品ＳＸＬタイム測定位置＿方
    HSXLTSPT As String * 1            ' 品ＳＸＬタイム測定位置＿点
    HSXLTSPI As String * 1            ' 品ＳＸＬタイム測定位置＿位
    HSXLTHWT As String * 1            ' 品ＳＸＬタイム保証方法＿対
    HSXLTHWS As String * 1            ' 品ＳＸＬタイム保証方法＿処
    '結晶内側管理
    EPDUP As Integer                  ' EPD　上限
    
    'WF仕様(結晶判定用)　08/4/15 ooba START ==========================>
    HWFRHWYS As String * 1          ' 品ＷＦ比抵抗保証方法＿処
    HWFONHWS As String * 1          ' 品ＷＦ酸素濃度保証方法＿処
    HWFOF1HS As String * 1          ' 品ＷＦＯＳＦ１保証方法＿処
    HWFOF2HS As String * 1          ' 品ＷＦＯＳＦ２保証方法＿処
    HWFOF3HS As String * 1          ' 品ＷＦＯＳＦ３保証方法＿処
    HWFOF4HS As String * 1          ' 品ＷＦＯＳＦ４保証方法＿処
    HWFBM1HS As String * 1          ' 品ＷＦＢＭＤ１保証方法＿処
    HWFBM2HS As String * 1          ' 品ＷＦＢＭＤ２保証方法＿処
    HWFBM3HS As String * 1          ' 品ＷＦＢＭＤ３保証方法＿処
    HWFDENHS As String * 1          ' 品ＷＦＤｅｎ保証方法＿処
    HWFDVDHS As String * 1          ' 品ＷＦＤＶＤ２保証方法＿処
    HWFLDLHS As String * 1          ' 品ＷＦＬ／ＤＬ保証方法＿処
    HWFRKHNN As String * 1          ' 品ＷＦ比抵抗検査頻度＿抜
    HWFONKHN As String * 1          ' 品ＷＦ酸素濃度検査頻度＿抜
    HWFOF1KN As String * 1          ' 品ＷＦＯＳＦ１検査頻度＿抜
    HWFOF2KN As String * 1          ' 品ＷＦＯＳＦ２検査頻度＿抜
    HWFOF3KN As String * 1          ' 品ＷＦＯＳＦ３検査頻度＿抜
    HWFOF4KN As String * 1          ' 品ＷＦＯＳＦ４検査頻度＿抜
    HWFBM1KN As String * 1          ' 品ＷＦＢＭＤ１検査頻度＿抜
    HWFBM2KN As String * 1          ' 品ＷＦＢＭＤ２検査頻度＿抜
    HWFBM3KN As String * 1          ' 品ＷＦＢＭＤ３検査頻度＿抜
    HWFGDKHN As String * 1          ' 品ＷＦＧＤ検査頻度＿抜
    'WF仕様(結晶判定用)　08/4/15 ooba END ============================>
    
' 払出規制項目追加対応 yakimura 2002.12.01 start
    TOPREG  As Integer                ' TOP規制
    TAILREG As Double                 ' TAIL規制
    BTMSPRT As Integer                ' ボトム析出規制
' 払出規制項目追加対応 yakimura 2002.12.01 end

' OSF，BMD項目追加対応  2002.04.02 yakimura
    HSXOSF1PTK As String * 1          ' 品ＳＸＯＳＦ１パタン区分
    HSXOSF2PTK As String * 1          ' 品ＳＸＯＳＦ２パタン区分
    HSXOSF3PTK As String * 1          ' 品ＳＸＯＳＦ３パタン区分
    HSXOSF4PTK As String * 1          ' 品ＳＸＯＳＦ４パタン区分
    HSXBMD1MBP As Double              ' 品ＳＸＢＭＤ１面内分布
    HSXBMD2MBP As Double              ' 品ＳＸＢＭＤ２面内分布
    HSXBMD3MBP As Double              ' 品ＳＸＢＭＤ３面内分布
' OSF，BMD項目追加対応  2002.04.02 yakimura
    BLOCKHFLAG As String * 1
''Upd Start (TCS)T.Terauchi 2005/10/12  GDﾗｲﾝ数表示対応
    HSXGDLINE   As String * 3         ' GDﾗｲﾝ数
''Upd End   (TCS)T.Terauchi 2005/10/12  GDﾗｲﾝ数表示対応

'C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
    COSF3FLAG As String * 1
'C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---

'--------------- 2008/08/25 INSERT START  By Systech ---------------
    HSXDKTMP    As String * 1              ' DK温度（仕様）
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
    HSXGDPTK As String * 1          ' 品ＳＸＧＤパタン区分
    HWFGDPTK    As String * 1       ' 品ＷＦＧＤパタン区分
    WFHSGDCW    As String * 1       ' 保証FLG（GD)
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
''2009/07/13 add Kameda 窒素 -----------------------
    HSXCDOPMN As Double
    HSXCDOPMX As Double
    HSXCDPNI As String
    HSXCDOPN As Double
''---------------------------------------------------
''2009/08/12 add Kameda 結晶面傾
    HSXCSCEN As Double
    HSXCSMIN As Double
    HSXCSMAX As Double
''2009/09/01 add Kameda 結晶面傾
    HSXCYCEN As Double
    HSXCYMIN As Double
    HSXCYMAX As Double
    HSXCTCEN As Double
    HSXCTMIN As Double
    HSXCTMAX As Double
''2010/02/04 add Kameda SIRD
    HWFSIRDMX As Double
    
  'Add Start 2011/01/17 SMPK A.Nagamine : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の仕様項目追加
    HSXCPK      As String * 1       ' 品ＳＸＣパターン区分
    HSXCSZ      As String * 1       ' 品ＳＸＣ測定条件
    HSXCHT      As String * 1       ' 品ＳＸＣ保証方法＿対
    HSXCHS      As String * 1       ' 品ＳＸＣ保証方法＿処
    HSXCJPK     As String * 1       ' 品ＳＸＣＪパターン区分
    HSXCJNS     As String * 2       ' 品ＳＸＣＪ熱処理法
    HSXCJHT     As String * 1       ' 品ＳＸＣＪ保証方法＿対
    HSXCJHS     As String * 1       ' 品ＳＸＣＪ保証方法＿処
    HSXCJLTPK   As String * 1       ' 品ＳＸＣＪＬＴパターン区分
    HSXCJLTNS   As String * 2       ' 品ＳＸＣＪＬＴ熱処理法
    HSXCJLTHT   As String * 1       ' 品ＳＸＣＪＬＴ保証方法＿対
    HSXCJLTHS   As String * 1       ' 品ＳＸＣＪＬＴ保証方法＿処
    HSXCJ2PK    As String * 1       ' 品ＳＸＣＪ２パターン区分
    HSXCJ2NS    As String * 2       ' 品ＳＸＣＪ２熱処理法
    HSXCJ2HT    As String * 1       ' 品ＳＸＣＪ２保証方法＿対
    HSXCJ2HS    As String * 1       ' 品ＳＸＣＪ２保証方法＿処
    HSXCJLTBND  As Integer          ' 品SXL/CJLTバンド幅 Number(3,0)
  'Add End   2011/01/17 SMPK A.Nagamine

'Add Start 2011/02/28 SMPK H.Ohkubo
    HSXONKHI As String * 1          ' 品ＳＸ酸素濃度検査頻度＿位
    FRSFLG   As String * 1          ' FRS測定有無
'Add End 2011/02/28 SMPK H.Ohkubo
'Add Start 2012/06/01 SMPK H.Ohkubo
    HSXCOSF3PK   As String * 1      ' 品ＳＸＣＯＳＦ３パターン区分
'Add Start 2012/06/01 SMPK H.Ohkubo
End Type

' 新サンプル管理(ﾌﾞﾛｯｸ)取得用(TOP,TAIL順で２レコード取得)
Public Type type_DBDRV_scmzc_fcmkc001c_CrySmp
    CRYNUMCS        As String * 12      'ブロックID
    Length          As Integer          ' 長さ
    SMPKBNCS        As String * 1       'サンプル区分
    TBKBNCS         As String * 1       'T/B区分
    REPSMPLIDCS     As Long             '代表サンプルID         Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    XTALCS          As String * 12      '結晶番号
    INPOSCS         As Integer          '結晶内位置
    HINBCS          As String * 8       '品番
    REVNUMCS        As Integer          '製品番号改訂番号
    FACTORYCS       As String * 1       '工場
    OPECS           As String * 1       '操業条件
    KTKBNCS         As String * 1       '確定区分
    BLKKTFLAGCS     As String * 1       'ブロック確定フラグ
    CRYSMPLIDRSCS   As Long             'サンプルID(Rs)         Integer→Long サンプル№6桁対応
    CRYSMPLIDRS1CS  As Long             '推定サンプルID1(Rs)    Integer→Long サンプル№6桁対応
    CRYSMPLIDRS2CS  As Long             '推定サンプルID2(Rs)    Integer→Long サンプル№6桁対応
    CRYINDRSCS      As String * 1       '状態FLG(Rs)
    CRYRESRS1CS     As String * 1       '実績FLG1(Rs)
    CRYRESRS2CS     As String * 1       '実績FLG2(Rs)
    CRYSMPLIDOICS   As Long             'サンプルID(Oi)         Integer→Long サンプル№6桁対応
    CRYINDOICS      As String * 1       '状態FLG(Oi)
    CRYRESOICS      As String * 1       '実績FLG(Oi)
    CRYSMPLIDB1CS   As Long             'サンプルID(B1)         Integer→Long サンプル№6桁対応
    CRYINDB1CS      As String * 1       '状態FLG(B1)
    CRYRESB1CS      As String * 1       '実績FLG(B1)
    CRYSMPLIDB2CS   As Long             'サンプルID(B2)         Integer→Long サンプル№6桁対応
    CRYINDB2CS      As String * 1       '状態FLG(B2)
    CRYRESB2CS      As String * 1       '実績FLG(B2)
    CRYSMPLIDB3CS   As Long             'サンプルID(B3)         Integer→Long サンプル№6桁対応
    CRYINDB3CS      As String * 1       '状態FLG(B3)
    CRYRESB3CS      As String * 1       '実績FLG(B3)
    CRYSMPLIDL1CS   As Long             'サンプルID(L1)         Integer→Long サンプル№6桁対応
    CRYINDL1CS      As String * 1       '状態FLG(L1)
    CRYRESL1CS      As String * 1       '実績FLG(L1)
    CRYSMPLIDL2CS   As Long             'サンプルID(L2)         Integer→Long サンプル№6桁対応
    CRYINDL2CS      As String * 1       '状態FLG(L2)
    CRYRESL2CS      As String * 1       '実績FLG(L2)
    CRYSMPLIDL3CS   As Long             'サンプルID(L3)         Integer→Long サンプル№6桁対応
    CRYINDL3CS      As String * 1       '状態FLG(L3)
    CRYRESL3CS      As String * 1       '実績FLG(L3)
    CRYSMPLIDL4CS   As Long             'サンプルID(L4)         Integer→Long サンプル№6桁対応
    CRYINDL4CS      As String * 1       '状態FLG(L4)
    CRYRESL4CS      As String * 1       '実績FLG(L4)
    CRYSMPLIDCSCS   As Long             'サンプルID(Cs)         Integer→Long サンプル№6桁対応
    CRYINDCSCS      As String * 1       '状態FLG(Cs)
    CRYRESCSCS      As String * 1       '実績FLG(Cs)
    CRYSMPLIDGDCS   As Long             'サンプルID(GD)         Integer→Long サンプル№6桁対応
    CRYINDGDCS      As String * 1       '状態FLG(GD)
    CRYRESGDCS      As String * 1       '実績FLG(GD)
    CRYSMPLIDTCS    As Long             'サンプルID(T)          Integer→Long サンプル№6桁対応
    CRYINDTCS       As String * 1       '状態FLG(T)
    CRYRESTCS       As String * 1       '実績FLG(T)
''Add Start 2011/07/25 LT10Ω判定追加対応 T.Koi(SETsw)
    CRYREST10CS     As String * 1       '実績FLG(T10)
''Add End   2011/07/25 LT10Ω判定追加対応 T.Koi(SETsw)
    CRYSMPLIDEPCS   As Long             'サンプルID(EPD)        Integer→Long サンプル№6桁対応
    CRYINDEPCS      As String * 1       '状態FLG(EPD)
    CRYRESEPCS      As String * 1       '実績FLG(EPD)
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    HSXDKTMP        As String * 1       'DK温度(実績)
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    CRYINDXC1       As String * 1       '状態FLG(X)     2009/08/12 Kameda
    CRYRESXC1       As String * 1       '実績FLG(X)     2009/08/12 Kameda
    SIRDKBNY3       As String * 1       '状態FLG(SIRD)  2010/02/04 Kameda
    
  'Add Start 2011/01/17 SMPK A.Nagamine : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の仕様項目追加
    CRYSMPLIDCCS    As Long             ' サンプルID(C)
    CRYINDCCS       As String * 1       ' 状態FLG(C)
    CRYRESCCS       As String * 1       ' 実績FLG(C)
    CRYSMPLIDCJCS   As Long             ' サンプルID(CJ)
    CRYINDCJCS      As String * 1       ' 状態FLG(CJ)
    CRYRESCJCS      As String * 1       ' 実績FLG(CJ)
    CRYSMPLIDCJLTCS As Long             ' サンプルID(CJ[LT])
    CRYINDCJLTCS    As String * 1       ' 状態FLG(CJ[LT])
    CRYRESCJLTCS    As String * 1       ' 実績FLG(CJ[LT])
    CRYSMPLIDCJ2CS  As Long             ' サンプルID(CJ2)
    CRYINDCJ2CS     As String * 1       ' 状態FLG(CJ2)
    CRYRESCJ2CS     As String * 1       ' 実績FLG(CJ2)
  'Add End   2011/01/17 SMPK A.Nagamine
End Type

'実績をまとめた構造体
Public Type type_DBDRV_scmzc_fcmkc001c_Zisseki
    CRYRZ() As type_DBDRV_scmzc_fcmkc001c_CryR
    OIZ()   As type_DBDRV_scmzc_fcmkc001c_Oi
    BMD1Z() As type_DBDRV_scmzc_fcmkc001c_BMD
    BMD2Z() As type_DBDRV_scmzc_fcmkc001c_BMD
    BMD3Z() As type_DBDRV_scmzc_fcmkc001c_BMD
    OSF1Z() As type_DBDRV_scmzc_fcmkc001c_OSF
    OSF2Z() As type_DBDRV_scmzc_fcmkc001c_OSF
    OSF3Z() As type_DBDRV_scmzc_fcmkc001c_OSF
    OSF4Z() As type_DBDRV_scmzc_fcmkc001c_OSF
    CSZ()   As type_DBDRV_scmzc_fcmkc001c_CS
    GDZ()   As type_DBDRV_scmzc_fcmkc001c_GD
    LTZ()   As type_DBDRV_scmzc_fcmkc001c_LT
    EPDZ()  As type_DBDRV_scmzc_fcmkc001c_EPD
    SURSZ() As type_DBDRV_scmzc_fcmkc001c_CryR
    XZ As type_DBDRV_scmzc_fcmkc001c_X
    SIRD As type_DBDRV_scmzc_fcmkc001c_SIRD
    
  'Add Start 2011/01/17 SMPK A.Nagamine : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の項目追加 SB_CryHanSui.bas
    CuC()       As type_DBDRV_scmzc_fcmkc001c_C     'C     実績
    CuCJ()      As type_DBDRV_scmzc_fcmkc001c_CJ    'CJ    実績
    CuCJLT()    As type_DBDRV_scmzc_fcmkc001c_CJLT  'CJ(LT)実績
    CuCJ2()     As type_DBDRV_scmzc_fcmkc001c_CJ2   'CJ2   実績
  'Add End   2011/01/17 SMPK A.Nagamine
End Type

'測定結果のJ014書込要否構造体
Public Type Judg_Spec_Cry
    Enable  As Boolean          '有効な品番である
    rs      As Boolean          'Rsは要書込
    Oi      As Boolean          'Oiは要書込
    B1      As Boolean          'BMD1は要書込
    B2      As Boolean          'BMD2は要書込
    B3      As Boolean          'BMD3は要書込
    L1      As Boolean          'OSF1は要書込
    L2      As Boolean          'OSF2は要書込
    L3      As Boolean          'OSF3は要書込
    L4      As Boolean          'OSF4は要書込
'C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
    COSF3   As Boolean          'C-OSF3ﾌﾗｸﾞ
'C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---
    Cs      As Boolean          'Csは要書込
    GD      As Boolean          'GDは要書込
    Lt      As Boolean          'LTは要書込
    EPD     As Boolean          'EPDは要書込
    
  'Add Start 2011/01/17 SMPK A.Nagamine : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の項目追加
    CuC     As Boolean          'Cは要書込
    CuCJ    As Boolean          'CJは要書込
    CuCJLT  As Boolean          'CJ(LT)は要書込
    CuCJ2   As Boolean          'CJ2は要書込
  'Add End   2011/01/17 SMPK A.Nagamine
End Type

' 仕様の指示がたっている判断用
Public Const SIJI = "H"
Public Const SANKOU = "S"

'概要      :総合判定 各種データ取得
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                 ,説明
'          :inBlockID     ,I  ,String                             ,対象ブロックID
'          :tNew_Hinban   ,I  ,tFullHinban                        ,対象品番(構造体)
'          :Siyou()       ,O  ,type_DBDRV_scmzc_fcmkc001c_Siyou   ,品番、仕様、結晶内側取得用
'          :CrySmp()      ,O  ,type_DBDRV_scmzc_fcmkc001c_CrySmp  ,結晶サンプル管理取得用
'          :Zisseki       ,O  ,type_DBDRV_scmzc_fcmkc001c_Zisseki ,実績用
'          :sErrMsg       ,O  ,String                             ,
'          :iSmpGetFlg    ,I  ,Integer                            :ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ(0:ｻﾝﾌﾟﾙ指定なし, 1:ｻﾝﾌﾟﾙ指定あり)
'          :iSamplID1     ,I  ,Long                               :TOPｻﾝﾌﾟﾙID(省略可)   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :iSamplID2     ,I  ,Long                               :BOTｻﾝﾌﾟﾙID(省略可)   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :戻り値        ,O  ,FUNCTION_RETURN                    ,読み込み成否
'説明      :
'履歴      :2001/06/26 蔵本 作成
Public Function funCryGetDataEtc(inBlockID As String, tNew_Hinban As tFullHinban, _
                                 siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                                 CrySmp() As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                                 Zisseki As type_DBDRV_scmzc_fcmkc001c_Zisseki, _
                                 sErrMsg As String, _
                                 iSmpGetFlg As Integer, iSamplID1 As Long, iSamplID2 As Long) As FUNCTION_RETURN

    Dim chk_cnt As Integer
    Dim i       As Integer
    Dim recCnt  As Integer
    Dim sDbName As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function funCryGetDataEtc"

    funCryGetDataEtc = FUNCTION_RETURN_FAILURE

    sDbName = "V011"
    '品番、SXL仕様からデータの取得（レコード0件の場合もエラー）
    If getHinSiyou(inBlockID, tNew_Hinban, siyou()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        funCryGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '同一品番をコピー
    chk_cnt = UBound(siyou)
    If chk_cnt = 1 Then
        ReDim Preserve siyou(chk_cnt + 1)
        siyou(chk_cnt + 1) = siyou(chk_cnt)
    End If
    
    sDbName = "V010"
    '結晶サンプルの取得(レコード0件の場合もエラー)
    If getCrySmp(inBlockID, CrySmp(), iSmpGetFlg, iSamplID1, iSamplID2) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        funCryGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
  'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA評価対応(Cu-deco) :
    recCnt = UBound(CrySmp)
  'Add End   2011/01/17 SMPK A.Nagamine

    With Zisseki
  'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA評価対応(Cu-deco) :
        'ReDim .CRYRZ(2)
        'ReDim .OIZ(2)
        'ReDim .BMD1Z(2)
        'ReDim .BMD2Z(2)
        'ReDim .BMD3Z(2)
        'ReDim .OSF1Z(2)
        'ReDim .OSF2Z(2)
        'ReDim .OSF3Z(2)
        'ReDim .OSF4Z(2)
        'ReDim .CSZ(2)
        'ReDim .GDZ(2)
        'ReDim .LTZ(2)
        'ReDim .EPDZ(2)
        'ReDim .SURSZ(2)
        
        ReDim .CRYRZ(recCnt)
        ReDim .OIZ(recCnt)
        ReDim .BMD1Z(recCnt)
        ReDim .BMD2Z(recCnt)
        ReDim .BMD3Z(recCnt)
        ReDim .OSF1Z(recCnt)
        ReDim .OSF2Z(recCnt)
        ReDim .OSF3Z(recCnt)
        ReDim .OSF4Z(recCnt)
        ReDim .CSZ(recCnt)
        ReDim .GDZ(recCnt)
        ReDim .LTZ(recCnt)
        ReDim .EPDZ(recCnt)
        ReDim .SURSZ(recCnt)
  'Add End   2011/01/17 SMPK A.Nagamine
    
  'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の項目追加
        ReDim .CuC(recCnt)
        ReDim .CuCJ(recCnt)
        ReDim .CuCJLT(recCnt)
        ReDim .CuCJ2(recCnt)
  'Add End   2011/01/17 SMPK A.Nagamine
    End With
    
  'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA評価対応(Cu-deco) :
'    recCnt = UBound(CrySmp)
  'Add End   2011/01/17 SMPK A.Nagamine

    '結晶サンプルの指示を見て実績を取る
    For i = 1 To recCnt
        
        sDbName = "J002"
        If CryR_Zisseki(siyou(i), CrySmp(i), Zisseki.CRYRZ(i), Zisseki.SURSZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J003"
        If Oi_Zisseki(siyou(i), CrySmp(i), Zisseki.OIZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J008"
        If BMD_Zisseki(siyou(i), CrySmp(i), "1", Zisseki.BMD1Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J008"
        If BMD_Zisseki(siyou(i), CrySmp(i), "2", Zisseki.BMD2Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J008"
        If BMD_Zisseki(siyou(i), CrySmp(i), "3", Zisseki.BMD3Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J005"
        If OSF_Zisseki(siyou(i), CrySmp(i), "1", Zisseki.OSF1Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J005"
        If OSF_Zisseki(siyou(i), CrySmp(i), "2", Zisseki.OSF2Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J005"
        If OSF_Zisseki(siyou(i), CrySmp(i), "3", Zisseki.OSF3Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J005"
        If OSF_Zisseki(siyou(i), CrySmp(i), "4", Zisseki.OSF4Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J004"
        If CS_Zisseki(siyou(i), CrySmp(i), Zisseki.CSZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J006"
        If GD_Zisseki(siyou(i), CrySmp(i), Zisseki.GDZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J007"
        If LT_Zisseki(siyou(i), CrySmp(i), Zisseki.LTZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J001"
        If EPD_Zisseki(siyou(i), CrySmp(i), Zisseki.EPDZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        
      'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の項目追加
        sDbName = "J023"
        If CuDeco_C_Zisseki(siyou(i), CrySmp(i), Zisseki.CuC(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J023"
        If CuDeco_CJ_Zisseki(siyou(i), CrySmp(i), Zisseki.CuCJ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J023"
        If CuDeco_CJLT_Zisseki(siyou(i), CrySmp(i), Zisseki.CuCJLT(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J023"
        If CuDeco_CJ2_Zisseki(siyou(i), CrySmp(i), Zisseki.CuCJ2(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
      'Add End   2011/01/17 SMPK A.Nagamine
        
    Next
    '2009/08/12 Kameda
    'X線検査測定フラグの取得
    If GetXSDC1_XRAY(CrySmp(recCnt)) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", "XSDC1_XRAY")
        funCryGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    sDbName = "J021"
    If X_Zisseki(CrySmp(recCnt).XTALCS, Zisseki.XZ) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
    
    '2010/02/04 Kameda
    'SIRD評価区分取得
    If GetXODY3_SIRD(CrySmp(1)) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", "XODY3_SIRD")
        funCryGetDataEtc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    sDbName = "J022"
    If SIRD_Zisseki(CrySmp(1), Zisseki.SIRD) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
    
    
    
    sDbName = ""
    funCryGetDataEtc = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    If Trim$(sDbName) <> "" Then sErrMsg = GetMsgStr("EGET2", sDbName)
    If recCnt > 2 Then
        sErrMsg = "0"
    End If
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    funCryGetDataEtc = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :総合判定 各種データ取得(結晶総合判定：反映データの合否判定を行わない用)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                 ,説明
'          :inBlockID     ,I  ,String                             ,対象ブロックID
'          :Top_Hinban      ,I  ,tFullHinban                      ,TOP品番
'          :Tail_Hinban     ,I  ,tFullHinban                      ,TAIL品番
'          :Siyou()       ,O  ,type_DBDRV_scmzc_fcmkc001c_Siyou   ,品番、仕様、結晶内側取得用
'          :CrySmp()      ,O  ,type_DBDRV_scmzc_fcmkc001c_CrySmp  ,結晶サンプル管理取得用
'          :Zisseki       ,O  ,type_DBDRV_scmzc_fcmkc001c_Zisseki ,実績用
'          :sErrMsg       ,O  ,String                             ,ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :iSmpGetFlg    ,I  ,Integer                            ,ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ(0:ｻﾝﾌﾟﾙ指定なし, 1:ｻﾝﾌﾟﾙ指定あり)
'          :iSamplID1     ,I  ,Long                               ,TOPｻﾝﾌﾟﾙID(省略可)   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :iSamplID2     ,I  ,Long                               ,BOTｻﾝﾌﾟﾙID(省略可)   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :戻り値        ,O  ,FUNCTION_RETURN                    ,読み込み成否
'説明      :
'履歴      :2005/02/08 作成  ffc)tanabe
Public Function funCryGetDataEtc2(inBlockID As String, Top_Hinban As tFullHinban, Tail_Hinban As tFullHinban, _
                                 siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                                 CrySmp() As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                                 Zisseki As type_DBDRV_scmzc_fcmkc001c_Zisseki, _
                                 sErrMsg As String, _
                                 iSmpGetFlg As Integer, iSamplID1 As Long, iSamplID2 As Long) As FUNCTION_RETURN

    Dim i       As Integer                              'for文用変数
    Dim recCnt  As Integer                              '結晶サンプル指示件数
    Dim t_Siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou   '仕様構造体
    Dim sDbName As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function funCryGetDataEtc2"

    funCryGetDataEtc2 = FUNCTION_RETURN_FAILURE

    '仕様配列の初期化
    ReDim siyou(2)

    sDbName = "V011"
    'TOP側
    '品番、SXL仕様からデータの取得（レコード0件の場合もエラー）
    If getHinSiyou(inBlockID, Top_Hinban, t_Siyou()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        funCryGetDataEtc2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    'TOP側の仕様データを格納する。
    siyou(1) = t_Siyou(1)

    'TAIL側
    '品番、SXL仕様からデータの取得（レコード0件の場合もエラー）
    If getHinSiyou(inBlockID, Tail_Hinban, t_Siyou()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        funCryGetDataEtc2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    'TAIL側の仕様データを格納する。
    siyou(2) = t_Siyou(1)
    
    sDbName = "V010"
    '結晶サンプルの取得(レコード0件の場合もエラー)
    If getCrySmp(inBlockID, CrySmp(), iSmpGetFlg, iSamplID1, iSamplID2) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        funCryGetDataEtc2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
  'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA評価対応(Cu-deco) :
    recCnt = UBound(CrySmp)
  'Add End   2011/01/17 SMPK A.Nagamine

    With Zisseki
  'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA評価対応(Cu-deco) :
        'ReDim .CRYRZ(2)
        'ReDim .OIZ(2)
        'ReDim .BMD1Z(2)
        'ReDim .BMD2Z(2)
        'ReDim .BMD3Z(2)
        'ReDim .OSF1Z(2)
        'ReDim .OSF2Z(2)
        'ReDim .OSF3Z(2)
        'ReDim .OSF4Z(2)
        'ReDim .CSZ(2)
        'ReDim .GDZ(2)
        'ReDim .LTZ(2)
        'ReDim .EPDZ(2)
        'ReDim .SURSZ(2)
  
        ReDim .CRYRZ(recCnt)
        ReDim .OIZ(recCnt)
        ReDim .BMD1Z(recCnt)
        ReDim .BMD2Z(recCnt)
        ReDim .BMD3Z(recCnt)
        ReDim .OSF1Z(recCnt)
        ReDim .OSF2Z(recCnt)
        ReDim .OSF3Z(recCnt)
        ReDim .OSF4Z(recCnt)
        ReDim .CSZ(recCnt)
        ReDim .GDZ(recCnt)
        ReDim .LTZ(recCnt)
        ReDim .EPDZ(recCnt)
        ReDim .SURSZ(recCnt)
  'Add End   2011/01/17 SMPK A.Nagamine
    
  'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の項目追加
        ReDim .CuC(recCnt)
        ReDim .CuCJ(recCnt)
        ReDim .CuCJLT(recCnt)
        ReDim .CuCJ2(recCnt)
  'Add End   2011/01/17 SMPK A.Nagamine
    End With
    
  'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA評価対応(Cu-deco) :
'    recCnt = UBound(CrySmp)
  'Add End   2011/01/17 SMPK A.Nagamine

    '結晶サンプルの指示を見て実績を取る
    For i = 1 To recCnt
        
        sDbName = "J002"
        If CryR_Zisseki(siyou(i), CrySmp(i), Zisseki.CRYRZ(i), Zisseki.SURSZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J003"
        If Oi_Zisseki(siyou(i), CrySmp(i), Zisseki.OIZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J008"
        If BMD_Zisseki(siyou(i), CrySmp(i), "1", Zisseki.BMD1Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J008"
        If BMD_Zisseki(siyou(i), CrySmp(i), "2", Zisseki.BMD2Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J008"
        If BMD_Zisseki(siyou(i), CrySmp(i), "3", Zisseki.BMD3Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J005"
        If OSF_Zisseki(siyou(i), CrySmp(i), "1", Zisseki.OSF1Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J005"
        If OSF_Zisseki(siyou(i), CrySmp(i), "2", Zisseki.OSF2Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J005"
        If OSF_Zisseki(siyou(i), CrySmp(i), "3", Zisseki.OSF3Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J005"
        If OSF_Zisseki(siyou(i), CrySmp(i), "4", Zisseki.OSF4Z(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J004"
        If CS_Zisseki(siyou(i), CrySmp(i), Zisseki.CSZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J006"
        If GD_Zisseki(siyou(i), CrySmp(i), Zisseki.GDZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J007"
        If LT_Zisseki(siyou(i), CrySmp(i), Zisseki.LTZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J001"
        If EPD_Zisseki(siyou(i), CrySmp(i), Zisseki.EPDZ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        
      'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の項目追加
        sDbName = "J023"
        If CuDeco_C_Zisseki(siyou(i), CrySmp(i), Zisseki.CuC(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J023"
        If CuDeco_CJ_Zisseki(siyou(i), CrySmp(i), Zisseki.CuCJ(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J023"
        If CuDeco_CJLT_Zisseki(siyou(i), CrySmp(i), Zisseki.CuCJLT(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
        sDbName = "J023"
        If CuDeco_CJ2_Zisseki(siyou(i), CrySmp(i), Zisseki.CuCJ2(i), i) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
      'Add End   2011/01/17 SMPK A.Nagamine
        
    Next
    '2009/08/12 Kameda
    'X線検査測定フラグの取得
    If GetXSDC1_XRAY(CrySmp(recCnt)) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", "XSDC1_XRAY")
        funCryGetDataEtc2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    sDbName = "J021"
    If X_Zisseki(CrySmp(recCnt).XTALCS, Zisseki.XZ) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
    
    '2010/02/04 Kameda
    'SIRD評価区分取得
    If GetXODY3_SIRD(CrySmp(1)) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", "XODY3_SIRD")
        funCryGetDataEtc2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    sDbName = "J022"
    If SIRD_Zisseki(CrySmp(1), Zisseki.SIRD) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
    
    sDbName = ""
    funCryGetDataEtc2 = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    If Trim$(sDbName) <> "" Then sErrMsg = GetMsgStr("EGET2", sDbName)
    If recCnt > 2 Then
        sErrMsg = "0"
    End If
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    funCryGetDataEtc2 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :内部関数 品番、仕様を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                 ,説明
'          :inBlockID     ,I  ,String                             ,対象ブロックID
'          :tNew_Hinban   ,I  ,tFullHinban                        ,対象品番(構造体)
'          :Siyou()       ,O  ,type_DBDRV_scmzc_fcmkc001c_Siyou   ,品番、仕様、結晶内側取得用
'          :戻り値        ,O  ,FUNCTION_RETURN                    ,読み込み成否
'説明      :
'履歴      :
Public Function getHinSiyou(inBlockID As String, tNew_Hinban As tFullHinban, _
                            siyou() As type_DBDRV_scmzc_fcmkc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim recCnt      As Integer
    Dim i           As Long
    Dim Jiltuseki   As Judg_Kakou
    Dim iIngotPos   As Integer          '結晶内位置
    Dim iLength     As Integer          '長さ
    Dim sCryNum     As String           '結晶番号
    
    '品番、SXL仕様からデータの取得
' 払出規制項目追加対応 yakimura 2002.12.01 start
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function getHinSiyou"

    getHinSiyou = FUNCTION_RETURN_SUCCESS

    If ciSmpGetFlg = 0 Then
        sql = "select "
        'ﾌﾞﾛｯｸ管理(TBCME040)参照停止　05/10/04 ooba START ===================================>
        sql = sql & "CSTOP.XTALCS as CRYNUM, "                      ' 結晶番号
        sql = sql & "CSTOP.INPOSCS as INGOTPOS, "                   ' 結晶内開始位置
        sql = sql & "CSBOT.INPOSCS - CSTOP.INPOSCS as LENGTH, "     ' 長さ
        'ﾌﾞﾛｯｸ管理(TBCME040)参照停止　05/10/04 ooba END =====================================>
    Else
        '工程実績ﾃﾞｰﾀ取得関数からﾃﾞｰﾀを取得し設定する
        If GET_hurikaeC3(inBlockID, ciKcnt, iIngotPos, iLength, sCryNum) = FUNCTION_RETURN_FAILURE Then
            getHinSiyou = FUNCTION_RETURN_FAILURE
            ReDim siyou(0)
            GoTo proc_exit
        End If
            
        sql = "select "
        sql = sql & sCryNum & " as CRYNUM, "        ' 結晶番号
        sql = sql & iIngotPos & " as INGOTPOS, "    ' 結晶内開始位置
        sql = sql & iLength & " as LENGTH, "        ' 長さ
    End If
    
    sql = sql & "E037.PRODCOND, "           ' 製作条件
    sql = sql & "E037.PGID, "               ' ＰＧ－ＩＤ
    sql = sql & "E037.UPLENGTH, "           ' 引上げ長さ
    sql = sql & "E037.FREELENG, "           ' フリー長
    sql = sql & "E037.DIAMETER, "           ' 直径
    sql = sql & "E037.CHARGE, "             ' チャージ量
    sql = sql & "E037.SEED, "               ' シード
    sql = sql & "E037.ADDDPPOS, "           ' 追加ドープ位置
    
    sql = sql & "E018.HSXTYPE, "             ' 品ＳＸタイプ
    sql = sql & "E018.HSXD1CEN, "            ' 品ＳＸ直径１中心
    sql = sql & "E018.HSXCDIR, "             ' 品ＳＸ結晶面方位
    
    sql = sql & "E018.HSXRMIN, "             ' 品ＳＸ比抵抗下限
    sql = sql & "E018.HSXRMAX, "             ' 品ＳＸ比抵抗上限
    sql = sql & "E018.HSXRAMIN, "            ' 品ＳＸ比抵抗平均下限
    sql = sql & "E018.HSXRAMAX, "            ' 品ＳＸ比抵抗平均上限
    sql = sql & "E018.HSXRMCAL, "            ' 品ＳＸ比抵抗面内計算　　'' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06
    sql = sql & "E018.HSXRMBNP, "            ' 品ＳＸ比抵抗面内分布
    sql = sql & "E018.HSXRSPOH, "            ' 品ＳＸ比抵抗測定位置＿方
    sql = sql & "E018.HSXRSPOT, "            ' 品ＳＸ比抵抗測定位置＿点
    sql = sql & "E018.HSXRSPOI, "            ' 品ＳＸ比抵抗測定位置＿位
    sql = sql & "E018.HSXRHWYT, "            ' 品ＳＸ比抵抗保証方法＿対
    sql = sql & "E018.HSXRHWYS, "            ' 品ＳＸ比抵抗保証方法＿処

    sql = sql & "E019.HSXONMIN, "            ' 品ＳＸ酸素濃度下限
    sql = sql & "E019.HSXONMAX, "            ' 品ＳＸ酸素濃度上限
    sql = sql & "E019.HSXONAMN, "            ' 品ＳＸ酸素濃度平均下限
    sql = sql & "E019.HSXONAMX, "            ' 品ＳＸ酸素濃度平均上限
    sql = sql & "E019.HSXONMCL, "            ' 品ＳＸ酸素濃度面内計算　　'' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06
    sql = sql & "E019.HSXONMBP, "            ' 品ＳＸ酸素濃度面内分布
    sql = sql & "E019.HSXONSPH, "            ' 品ＳＸ酸素濃度測定位置＿方
    sql = sql & "E019.HSXONSPT, "            ' 品ＳＸ酸素濃度測定位置＿点
    sql = sql & "E019.HSXONSPI, "            ' 品ＳＸ酸素濃度測定位置＿位
    sql = sql & "E019.HSXONHWT, "            ' 品ＳＸ酸素濃度保証方法＿対
    sql = sql & "E019.HSXONHWS, "            ' 品ＳＸ酸素濃度保証方法＿処

    sql = sql & "E020.HSXBM1AN, "            ' 品ＳＸＢＭＤ１平均下限
    sql = sql & "E020.HSXBM1AX, "            ' 品ＳＸＢＭＤ１平均上限
    sql = sql & "E020.HSXBM2AN, "            ' 品ＳＸＢＭＤ２平均下限
    sql = sql & "E020.HSXBM2AX, "            ' 品ＳＸＢＭＤ２平均上限
    sql = sql & "E020.HSXBM3AN, "            ' 品ＳＸＢＭＤ３平均下限
    sql = sql & "E020.HSXBM3AX, "            ' 品ＳＸＢＭＤ３平均上限
    sql = sql & "E020.HSXBM1SH, "            ' 品ＳＸＢＭＤ１測定位置＿方
    sql = sql & "E020.HSXBM1ST, "            ' 品ＳＸＢＭＤ１測定位置＿点
    sql = sql & "E020.HSXBM1SR, "            ' 品ＳＸＢＭＤ１測定位置＿領
    sql = sql & "E020.HSXBM1HT, "            ' 品ＳＸＢＭＤ１保証方法＿対
    sql = sql & "E020.HSXBM1HS, "            ' 品ＳＸＢＭＤ１保証方法＿処
    sql = sql & "E020.HSXBM2SH, "            ' 品ＳＸＢＭＤ２測定位置＿方
    sql = sql & "E020.HSXBM2ST, "            ' 品ＳＸＢＭＤ２測定位置＿点
    sql = sql & "E020.HSXBM2SR, "            ' 品ＳＸＢＭＤ２測定位置＿領
    sql = sql & "E020.HSXBM2HT, "            ' 品ＳＸＢＭＤ２保証方法＿対
    sql = sql & "E020.HSXBM2HS, "            ' 品ＳＸＢＭＤ２保証方法＿処
    sql = sql & "E020.HSXBM3SH, "            ' 品ＳＸＢＭＤ３測定位置＿方
    sql = sql & "E020.HSXBM3ST, "            ' 品ＳＸＢＭＤ３測定位置＿点
    sql = sql & "E020.HSXBM3SR, "            ' 品ＳＸＢＭＤ３測定位置＿領
    sql = sql & "E020.HSXBM3HT, "            ' 品ＳＸＢＭＤ３保証方法＿対
    sql = sql & "E020.HSXBM3HS, "            ' 品ＳＸＢＭＤ３保証方法＿処

    sql = sql & "E020.HSXOF1AX, "            ' 品ＳＸＯＳＦ１平均上限
    sql = sql & "E020.HSXOF1MX, "            ' 品ＳＸＯＳＦ１上限
    sql = sql & "E020.HSXOF2AX, "            ' 品ＳＸＯＳＦ２平均上限
    sql = sql & "E020.HSXOF2MX, "            ' 品ＳＸＯＳＦ２上限
    sql = sql & "E020.HSXOF3AX, "            ' 品ＳＸＯＳＦ３平均上限
    sql = sql & "E020.HSXOF3MX, "            ' 品ＳＸＯＳＦ３上限
    sql = sql & "E020.HSXOF4AX, "            ' 品ＳＸＯＳＦ４平均上限
    sql = sql & "E020.HSXOF4MX, "            ' 品ＳＸＯＳＦ４上限
    sql = sql & "E020.HSXOF1SH, "            ' 品ＳＸＯＳＦ１測定位置＿方
    sql = sql & "E020.HSXOF1ST, "            ' 品ＳＸＯＳＦ１測定位置＿点
    sql = sql & "E020.HSXOF1SR, "            ' 品ＳＸＯＳＦ１測定位置＿領
    sql = sql & "E020.HSXOF1HT, "            ' 品ＳＸＯＳＦ１保証方法＿対
    sql = sql & "E020.HSXOF1HS, "            ' 品ＳＸＯＳＦ１保証方法＿処
    sql = sql & "E020.HSXOF2SH, "            ' 品ＳＸＯＳＦ２測定位置＿方
    sql = sql & "E020.HSXOF2ST, "            ' 品ＳＸＯＳＦ２測定位置＿点
    sql = sql & "E020.HSXOF2SR, "            ' 品ＳＸＯＳＦ２測定位置＿領
    sql = sql & "E020.HSXOF2HT, "            ' 品ＳＸＯＳＦ２保証方法＿対
    sql = sql & "E020.HSXOF2HS, "            ' 品ＳＸＯＳＦ２保証方法＿処
    sql = sql & "E020.HSXOF3SH, "            ' 品ＳＸＯＳＦ３測定位置＿方
    sql = sql & "E020.HSXOF3ST, "            ' 品ＳＸＯＳＦ３測定位置＿点
    sql = sql & "E020.HSXOF3SR, "            ' 品ＳＸＯＳＦ３測定位置＿領
    sql = sql & "E020.HSXOF3HT, "            ' 品ＳＸＯＳＦ３保証方法＿対
    sql = sql & "E020.HSXOF3HS, "            ' 品ＳＸＯＳＦ３保証方法＿処
    sql = sql & "E020.HSXOF4SH, "            ' 品ＳＸＯＳＦ４測定位置＿方
    sql = sql & "E020.HSXOF4ST, "            ' 品ＳＸＯＳＦ４測定位置＿点
    sql = sql & "E020.HSXOF4SR, "            ' 品ＳＸＯＳＦ４測定位置＿領
    sql = sql & "E020.HSXOF4HT, "            ' 品ＳＸＯＳＦ４保証方法＿対
    sql = sql & "E020.HSXOF4HS, "            ' 品ＳＸＯＳＦ４保証方法＿処
    sql = sql & "E020.HSXOF1NS, "            ' 品ＳＸＯＳＦ１熱処理法
    sql = sql & "E020.HSXOF2NS, "            ' 品ＳＸＯＳＦ２熱処理法
    sql = sql & "E020.HSXOF3NS, "            ' 品ＳＸＯＳＦ３熱処理法
    sql = sql & "E020.HSXOF4NS, "            ' 品ＳＸＯＳＦ４熱処理法
    sql = sql & "E020.HSXBM1NS, "            ' 品ＳＸＢＭＤ１熱処理法
    sql = sql & "E020.HSXBM2NS, "            ' 品ＳＸＢＭＤ２熱処理法
    sql = sql & "E020.HSXBM3NS, "            ' 品ＳＸＢＭＤ３熱処理法

    sql = sql & "E019.HSXCNMIN, "            ' 品ＳＸ炭素濃度下限
    sql = sql & "E019.HSXCNMAX, "            ' 品ＳＸ炭素濃度上限
    sql = sql & "E019.HSXCNSPH, "            ' 品ＳＸ炭素濃度測定位置＿方
    sql = sql & "E019.HSXCNSPT, "            ' 品ＳＸ炭素濃度測定位置＿点
    sql = sql & "E019.HSXCNSPI, "            ' 品ＳＸ炭素濃度測定位置＿位
    sql = sql & "E019.HSXCNHWT, "            ' 品ＳＸ炭素濃度保証方法＿対
    sql = sql & "E019.HSXCNHWS, "            ' 品ＳＸ炭素濃度保証方法＿処
    sql = sql & "E019.HSXCNKHI, "            ' 品ＳＸ炭素濃度検査頻度＿位 09/01/08 ooba

    sql = sql & "E020.HSXDENMX, "            ' 品ＳＸＤｅｎ上限
    sql = sql & "E020.HSXDENMN, "            ' 品ＳＸＤｅｎ下限
    sql = sql & "E020.HSXLDLMX, "            ' 品ＳＸＬ／ＤＬ上限
    sql = sql & "E020.HSXLDLMN, "            ' 品ＳＸＬ／ＤＬ下限
    sql = sql & "E020.HSXDVDMXN, "           ' 品ＳＸＤＶＤ２上限   項目追加，修正対応 2003.05.20 yakimura
    sql = sql & "E020.HSXDVDMNN, "           ' 品ＳＸＤＶＤ２下限   項目追加，修正対応 2003.05.20 yakimura
    sql = sql & "E020.HSXDENHT, "            ' 品ＳＸＤｅｎ保証方法＿対
    sql = sql & "E020.HSXDENHS, "            ' 品ＳＸＤｅｎ保証方法＿処
    sql = sql & "E020.HSXLDLHT, "            ' 品ＳＸＬ／ＤＬ保証方法＿対
    sql = sql & "E020.HSXLDLHS, "            ' 品ＳＸＬ／ＤＬ保証方法＿処
    sql = sql & "E020.HSXDVDHT, "            ' 品ＳＸＤＶＤ２保証方法＿対
    sql = sql & "E020.HSXDVDHS, "            ' 品ＳＸＤＶＤ２保証方法＿処
    sql = sql & "E020.HSXDENKU, "            ' 品ＳＸＤｅｎ検査有無
    sql = sql & "E020.HSXDVDKU, "            ' 品ＳＸＤＶＤ２検査有無
    sql = sql & "E020.HSXLDLKU, "            ' 品ＳＸＬ／ＤＬ検査有無

    sql = sql & "E019.HSXLTMIN, "            ' 品ＳＸＬタイム下限
    sql = sql & "E019.HSXLTMAX, "            ' 品ＳＸＬタイム上限
''Add Start 2011/07/25 LT10Ω判定追加対応 T.Koi(SETsw)
    sql = sql & "E036.LTCONVAL, "            ' 品ＳＸＬLT10下限
''Add End   2011/07/25 LT10Ω判定追加対応 T.Koi(SETsw)
    sql = sql & "E019.HSXLTSPH, "            ' 品ＳＸＬタイム測定位置＿方
    sql = sql & "E019.HSXLTSPT, "            ' 品ＳＸＬタイム測定位置＿点
    sql = sql & "E019.HSXLTSPI, "            ' 品ＳＸＬタイム測定位置＿位
    sql = sql & "E019.HSXLTHWT, "            ' 品ＳＸＬタイム保証方法＿対
    sql = sql & "E019.HSXLTHWS, "            ' 品ＳＸＬタイム保証方法＿処
    sql = sql & "E036.EPDUP, "               ' EPD 上限
    sql = sql & "E036.TOPREG, "              ' TOP規制
    sql = sql & "E036.TAILREG, "             ' TAIL規制
    sql = sql & "E036.BTMSPRT, "             ' ボトム析出規制
'*** UPDATE ↓ Y.SIMIZU 2005/10/12 GDﾗｲﾝ数追加
    sql = sql & "E036.HSXGDLINE, "           ' 品ＳＸＬＧＤライン数
'*** UPDATE ↑ Y.SIMIZU 2005/10/12 GDﾗｲﾝ数追加

'C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
    sql = sql & "E036.COSF3FLAG, "           ' C-OSF3ﾌﾗｸﾞ
'C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---

'--------------- 2008/08/25 INSERT START  By Systech ---------------
    sql = sql & "NVL(E036.HSXDKTMP,' ') HSXDKTMP, "
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    sql = sql & "E036.HSXLDLRMN HSXLDLRMN, "
    sql = sql & "E036.HSXLDLRMX HSXLDLRMX, "
    sql = sql & "E036.HWFLDLRMN HWFLDLRMN, "
    sql = sql & "E036.HWFLDLRMX HWFLDLRMX, "
    sql = sql & "E036.HSXOF1ARPTK HSXOF1ARPTK, "
    sql = sql & "E036.HSXOFARMIN HSXOFARMIN, "
    sql = sql & "E036.HSXOFARMAX HSXOFARMAX, "
    sql = sql & "E036.HSXOFARMHMX HSXOFARMHMX, "
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End

' OSF，BMD項目追加対応  2002.04.02 yakimura
    sql = sql & "E020.HSXOSF1PTK, "          ' 品ＳＸＯＳＦ１パタン区分
    sql = sql & "E020.HSXOSF2PTK, "          ' 品ＳＸＯＳＦ２パタン区分
    sql = sql & "E020.HSXOSF3PTK, "          ' 品ＳＸＯＳＦ３パタン区分
    sql = sql & "E020.HSXOSF4PTK, "          ' 品ＳＸＯＳＦ４パタン区分
    sql = sql & "E020.HSXBMD1MBP, "          ' 品ＳＸＢＭＤ１面内分布
    sql = sql & "E020.HSXBMD2MBP, "          ' 品ＳＸＢＭＤ２面内分布
    sql = sql & "E020.HSXBMD3MBP, "          ' 品ＳＸＢＭＤ３面内分布
' OSF，BMD項目追加対応  2002.04.02 yakimura
    
    'WF仕様取得　08/04/15 ooba START ===========================================>
    sql = sql & "E021.HWFRHWYS, "            ' 品ＷＦ比抵抗保証方法＿処
    sql = sql & "E025.HWFONHWS, "            ' 品ＷＦ酸素濃度保証方法＿処
    sql = sql & "E029.HWFOF1HS, "            ' 品ＷＦＯＳＦ１保証方法＿処
    sql = sql & "E029.HWFOF2HS, "            ' 品ＷＦＯＳＦ２保証方法＿処
    sql = sql & "E029.HWFOF3HS, "            ' 品ＷＦＯＳＦ３保証方法＿処
    sql = sql & "E029.HWFOF4HS, "            ' 品ＷＦＯＳＦ４保証方法＿処
    sql = sql & "E029.HWFBM1HS, "            ' 品ＷＦＢＭＤ１保証方法＿処
    sql = sql & "E029.HWFBM2HS, "            ' 品ＷＦＢＭＤ２保証方法＿処
    sql = sql & "E029.HWFBM3HS, "            ' 品ＷＦＢＭＤ３保証方法＿処
    sql = sql & "E026.HWFDENHS, "            ' 品ＷＦＤｅｎ保証方法＿処
    sql = sql & "E026.HWFDVDHS, "            ' 品ＷＦＤＶＤ２保証方法＿処
    sql = sql & "E026.HWFLDLHS, "            ' 品ＷＦＬ／ＤＬ保証方法＿処
    sql = sql & "E021.HWFRKHNN, "            ' 品ＷＦ比抵抗検査頻度＿抜
    sql = sql & "E025.HWFONKHN, "            ' 品ＷＦ酸素濃度検査頻度＿抜
    sql = sql & "E029.HWFOF1KN, "            ' 品ＷＦＯＳＦ１検査頻度＿抜
    sql = sql & "E029.HWFOF2KN, "            ' 品ＷＦＯＳＦ２検査頻度＿抜
    sql = sql & "E029.HWFOF3KN, "            ' 品ＷＦＯＳＦ３検査頻度＿抜
    sql = sql & "E029.HWFOF4KN, "            ' 品ＷＦＯＳＦ４検査頻度＿抜
    sql = sql & "E029.HWFBM1KN, "            ' 品ＷＦＢＭＤ１検査頻度＿抜
    sql = sql & "E029.HWFBM2KN, "            ' 品ＷＦＢＭＤ２検査頻度＿抜
    sql = sql & "E029.HWFBM3KN, "            ' 品ＷＦＢＭＤ３検査頻度＿抜
    sql = sql & "E026.HWFGDKHN "             ' 品ＷＦＧＤ検査頻度＿抜
    'WF仕様取得　08/04/15 ooba END =============================================>

'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    sql = sql & ",E020.HSXGDPTK "            ' 品ＳＸＧＤパタン区分
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    sql = sql & ",E018.HSXCSCEN "            ' 品ＳＸ面傾き中心  2009/08/12 Kameda
    sql = sql & ",E018.HSXCSMIN "            ' 品ＳＸ面傾き下限  2009/08/12 Kameda
    sql = sql & ",E018.HSXCSMAX "            ' 品ＳＸ面傾き上限  2009/08/12 Kameda
    sql = sql & ",E018.HSXCYCEN "            ' 品ＳＸ面傾き中心  2009/09/01 Kameda
    sql = sql & ",E018.HSXCYMIN "            ' 品ＳＸ面傾き下限  2009/09/01 Kameda
    sql = sql & ",E018.HSXCYMAX "            ' 品ＳＸ面傾き上限  2009/09/01 Kameda
    sql = sql & ",E018.HSXCTCEN "            ' 品ＳＸ面傾き中心  2009/09/01 Kameda
    sql = sql & ",E018.HSXCTMIN "            ' 品ＳＸ面傾き下限  2009/09/01 Kameda
    sql = sql & ",E018.HSXCTMAX "            ' 品ＳＸ面傾き上限  2009/09/01 Kameda
    sql = sql & ",E048.HWFSIRDMX "           ' 品WF面内個数上限  2010/02/04 Kameda
    
  'Add Start 2011/01/17 SMPK A.Nagamine : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の仕様項目追加
    sql = sql & ",E020.HSXCPK,    E020.HSXCSZ,    E020.HSXCHT,    E020.HSXCHS,    E020.HSXCJPK   "
    sql = sql & ",E020.HSXCJNS,   E020.HSXCJHT,   E020.HSXCJHS,   E020.HSXCJLTPK, E020.HSXCJLTNS "
    sql = sql & ",E020.HSXCJLTHT, E020.HSXCJLTHS, E020.HSXCJ2PK,  E020.HSXCJ2NS,  E020.HSXCJ2HT  "
    sql = sql & ",E020.HSXCJ2HS,  E036.HSXCJLTBND "
  'Add End   2011/01/17 SMPK A.Nagamine
  'Add Start 2012/06/01 SMPK H.Ohkubo
    sql = sql & ",NVL(E020.HSXCOSF3PK,'4') as HSXCOSF3PK"    '品ＳＸＣＯＳＦ３パターン区分"
  'Add End 2012/06/01 SMPK H.Ohkubo
    If ciSmpGetFlg = 0 Then
        'ﾌﾞﾛｯｸ管理(TBCME040)参照停止　05/10/04 ooba START ===================================>
        sql = sql & " from TBCME037 E037, TBCME018 E018, TBCME019 E019, TBCME020 E020, TBCME036 E036, "
        sql = sql & "      TBCME021 E021, TBCME025 E025, TBCME026 E026, TBCME029 E029, TBCME048 E048, "  '08/04/15 ooba, 2010/02/04 Kameda addE048
        sql = sql & " (select CRYNUMCS, XTALCS, INPOSCS from XSDCS "
        sql = sql & " where TBKBNCS = 'T' and CRYNUMCS = '" & inBlockID & "' "
        sql = sql & " ) CSTOP, "
        sql = sql & " (select CRYNUMCS, XTALCS, INPOSCS from XSDCS "
        sql = sql & " where TBKBNCS = 'B' and CRYNUMCS = '" & inBlockID & "' "
        sql = sql & " ) CSBOT "
        sql = sql & " where CSTOP.CRYNUMCS = CSBOT.CRYNUMCS and "
        sql = sql & "       E037.CRYNUM = '" & left(inBlockID, 9) & "000' and "
        'ﾌﾞﾛｯｸ管理(TBCME040)参照停止　05/10/04 ooba END =====================================>
    Else
        sql = sql & " from TBCME037 E037, TBCME018 E018, TBCME019 E019, TBCME020 E020, TBCME036 E036, "
        sql = sql & "      TBCME021 E021, TBCME025 E025, TBCME026 E026, TBCME029 E029, TBCME048 E048 "   '08/04/15 ooba, 2010/02/04 Kameda addE048
        sql = sql & " where E037.CRYNUM = '" & left(inBlockID, 9) & "000' and "
    End If
    sql = sql & "       E018.HINBAN = '" & tNew_Hinban.hinban & "' and "
    sql = sql & "       E018.MNOREVNO = " & tNew_Hinban.mnorevno & " and "
    sql = sql & "       E018.FACTORY = '" & tNew_Hinban.factory & "' and "
    sql = sql & "       E018.OPECOND = '" & tNew_Hinban.opecond & "' and "
    sql = sql & "       E019.HINBAN = E018.HINBAN and E019.MNOREVNO = E018.MNOREVNO and E019.FACTORY = E018.FACTORY and E019.OPECOND = E018.OPECOND and "
    sql = sql & "       E020.HINBAN = E018.HINBAN and E020.MNOREVNO = E018.MNOREVNO and E020.FACTORY = E018.FACTORY and E020.OPECOND = E018.OPECOND and "
    sql = sql & "       E021.HINBAN = E018.HINBAN and E021.MNOREVNO = E018.MNOREVNO and E021.FACTORY = E018.FACTORY and E021.OPECOND = E018.OPECOND and "   '08/04/15 ooba
    sql = sql & "       E025.HINBAN = E018.HINBAN and E025.MNOREVNO = E018.MNOREVNO and E025.FACTORY = E018.FACTORY and E025.OPECOND = E018.OPECOND and "   '08/04/15 ooba
    sql = sql & "       E026.HINBAN = E018.HINBAN and E026.MNOREVNO = E018.MNOREVNO and E026.FACTORY = E018.FACTORY and E026.OPECOND = E018.OPECOND and "   '08/04/15 ooba
    sql = sql & "       E029.HINBAN = E018.HINBAN and E029.MNOREVNO = E018.MNOREVNO and E029.FACTORY = E018.FACTORY and E029.OPECOND = E018.OPECOND and "   '08/04/15 ooba
    sql = sql & "       E036.HINBAN = E018.HINBAN and E036.MNOREVNO = E018.MNOREVNO and E036.FACTORY = E018.FACTORY and E036.OPECOND = E018.OPECOND and "
    sql = sql & "       E048.HINBAN = E018.HINBAN and E048.MNOREVNO = E018.MNOREVNO and E048.FACTORY = E018.FACTORY and E048.OPECOND = E018.OPECOND "       '2010/02/04 Kameda
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount = 0 Then
    End If

    recCnt = rs.RecordCount
    ReDim siyou(recCnt)
    For i = 1 To recCnt
    
        With siyou(i)
            .CRYNUM = rs("CRYNUM")                  ' 結晶番号
            .INGOTPOS = rs("INGOTPOS")              ' 結晶内開始位置
            .Length = rs("LENGTH")                  ' 長さ
            .HIN.hinban = tNew_Hinban.hinban        ' 品番
            .HIN.mnorevno = tNew_Hinban.mnorevno    ' 製品番号改訂番号
            .HIN.factory = tNew_Hinban.factory      ' 工場
            .HIN.opecond = tNew_Hinban.opecond      ' 操業条件
            
            .PRODCOND = rs("PRODCOND")              ' 製作条件
            .PGID = rs("PGID")                      ' ＰＧ－ＩＤ
            .UPLENGTH = rs("UPLENGTH")              ' 引上げ長さ
            .FREELENG = rs("FREELENG")              ' フリー長
            .DIAMETER = rs("DIAMETER")              ' 直径
            .CHARGE = rs("CHARGE")                  ' チャージ量
            .SEED = rs("SEED")                      ' シード
            .ADDDPPOS = rs("ADDDPPOS")              ' 追加ドープ位置
    
            .HSXTYPE = rs("HSXTYPE")                        ' 品ＳＸタイプ"
            .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))        ' 品ＳＸ直径１中心"         2003/12/10 SystemBrain Null対応
            .HSXCDIR = rs("HSXCDIR")                        ' 品ＳＸ結晶面方位"

            .HSXRMIN = fncNullCheck(rs("HSXRMIN"))          ' 品ＳＸ比抵抗下限          2003/12/10 SystemBrain Null対応
            .HSXRMAX = fncNullCheck(rs("HSXRMAX"))          ' 品ＳＸ比抵抗上限          2003/12/10 SystemBrain Null対応
            .HSXRAMIN = fncNullCheck(rs("HSXRAMIN"))        ' 品ＳＸ比抵抗平均下限      2003/12/10 SystemBrain Null対応
            .HSXRAMAX = fncNullCheck(rs("HSXRAMAX"))        ' 品ＳＸ比抵抗平均上限      2003/12/10 SystemBrain Null対応
            .HSXRMCAL = rs("HSXRMCAL")                      ' 品ＳＸ比抵抗面内計算     '' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06
            .HSXRMBNP = fncNullCheck(rs("HSXRMBNP"))        ' 品ＳＸ比抵抗面内分布      2003/12/10 SystemBrain Null対応
            .HSXRSPOH = rs("HSXRSPOH")                      ' 品ＳＸ比抵抗測定位置＿方
            .HSXRSPOT = rs("HSXRSPOT")                      ' 品ＳＸ比抵抗測定位置＿点
            .HSXRSPOI = rs("HSXRSPOI")                      ' 品ＳＸ比抵抗測定位置＿位
            .HSXRHWYT = rs("HSXRHWYT")                      ' 品ＳＸ比抵抗保証方法＿対
            .HSXRHWYS = rs("HSXRHWYS")                      ' 品ＳＸ比抵抗保証方法＿処

            .HSXONMIN = fncNullCheck(rs("HSXONMIN"))        ' 品ＳＸ酸素濃度下限        2003/12/10 SystemBrain Null対応
            .HSXONMAX = fncNullCheck(rs("HSXONMAX"))        ' 品ＳＸ酸素濃度上限        2003/12/10 SystemBrain Null対応
            .HSXONAMN = fncNullCheck(rs("HSXONAMN"))        ' 品ＳＸ酸素濃度平均下限    2003/12/10 SystemBrain Null対応
            .HSXONAMX = fncNullCheck(rs("HSXONAMX"))        ' 品ＳＸ酸素濃度平均上限    2003/12/10 SystemBrain Null対応
            .HSXONMCL = rs("HSXONMCL")                      ' 品ＳＸ酸素濃度面内計算   '' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06
            .HSXONMBP = fncNullCheck(rs("HSXONMBP"))        ' 品ＳＸ酸素濃度面内分布    2003/12/10 SystemBrain Null対応
            .HSXONSPH = rs("HSXONSPH")                      ' 品ＳＸ酸素濃度測定位置＿方
            .HSXONSPT = rs("HSXONSPT")                      ' 品ＳＸ酸素濃度測定位置＿点
            .HSXONSPI = rs("HSXONSPI")                      ' 品ＳＸ酸素濃度測定位置＿位
            .HSXONHWT = rs("HSXONHWT")                      ' 品ＳＸ酸素濃度保証方法＿対
            .HSXONHWS = rs("HSXONHWS")                      ' 品ＳＸ酸素濃度保証方法＿処

            .HSXBM1AN = fncNullCheck(rs("HSXBM1AN"))        ' 品ＳＸＢＭＤ１平均下限    2003/12/10 SystemBrain Null対応
            .HSXBM1AX = fncNullCheck(rs("HSXBM1AX"))        ' 品ＳＸＢＭＤ１平均上限    2003/12/10 SystemBrain Null対応
            .HSXBM1SH = rs("HSXBM1SH")                      ' 品ＳＸＢＭＤ１測定位置＿方
            .HSXBM1ST = rs("HSXBM1ST")                      ' 品ＳＸＢＭＤ１測定位置＿点
            .HSXBM1SR = rs("HSXBM1SR")                      ' 品ＳＸＢＭＤ１測定位置＿領
            .HSXBM1HT = rs("HSXBM1HT")                      ' 品ＳＸＢＭＤ１保証方法＿対
            .HSXBM1HS = rs("HSXBM1HS")                      ' 品ＳＸＢＭＤ１保証方法＿処
            .HSXBM1NS = rs("HSXBM1NS")                      ' 品ＳＸＢＭＤ１熱処理法
            .HSXBM2AN = fncNullCheck(rs("HSXBM2AN"))        ' 品ＳＸＢＭＤ２平均下限    2003/12/10 SystemBrain Null対応
            .HSXBM2AX = fncNullCheck(rs("HSXBM2AX"))        ' 品ＳＸＢＭＤ２平均上限    2003/12/10 SystemBrain Null対応
            .HSXBM2SH = rs("HSXBM2SH")                      ' 品ＳＸＢＭＤ２測定位置＿方
            .HSXBM2ST = rs("HSXBM2ST")                      ' 品ＳＸＢＭＤ２測定位置＿点
            .HSXBM2SR = rs("HSXBM2SR")                      ' 品ＳＸＢＭＤ２測定位置＿領
            .HSXBM2HT = rs("HSXBM2HT")                      ' 品ＳＸＢＭＤ２保証方法＿対
            .HSXBM2HS = rs("HSXBM2HS")                      ' 品ＳＸＢＭＤ２保証方法＿処
            .HSXBM2NS = rs("HSXBM2NS")                      ' 品ＳＸＢＭＤ２熱処理法
            .HSXBM3AN = fncNullCheck(rs("HSXBM3AN"))        ' 品ＳＸＢＭＤ３平均下限    2003/12/10 SystemBrain Null対応
            .HSXBM3AX = fncNullCheck(rs("HSXBM3AX"))        ' 品ＳＸＢＭＤ３平均上限    2003/12/10 SystemBrain Null対応
            .HSXBM3SH = rs("HSXBM3SH")                      ' 品ＳＸＢＭＤ３測定位置＿方
            .HSXBM3ST = rs("HSXBM3ST")                      ' 品ＳＸＢＭＤ３測定位置＿点
            .HSXBM3SR = rs("HSXBM3SR")                      ' 品ＳＸＢＭＤ３測定位置＿領
            .HSXBM3HT = rs("HSXBM3HT")                      ' 品ＳＸＢＭＤ３保証方法＿対
            .HSXBM3HS = rs("HSXBM3HS")                      ' 品ＳＸＢＭＤ３保証方法＿処
            .HSXBM3NS = rs("HSXBM3NS")                      ' 品ＳＸＢＭＤ３熱処理法
            
            .HSXOF1AX = fncNullCheck(rs("HSXOF1AX"))        ' 品ＳＸＯＳＦ１平均上限    2003/12/10 SystemBrain Null対応
            .HSXOF1MX = fncNullCheck(rs("HSXOF1MX"))        ' 品ＳＸＯＳＦ１上限        2003/12/10 SystemBrain Null対応
            .HSXOF1SH = rs("HSXOF1SH")                      ' 品ＳＸＯＳＦ１測定位置＿方
            .HSXOF1ST = rs("HSXOF1ST")                      ' 品ＳＸＯＳＦ１測定位置＿点
            .HSXOF1SR = rs("HSXOF1SR")                      ' 品ＳＸＯＳＦ１測定位置＿領
            .HSXOF1HT = rs("HSXOF1HT")                      ' 品ＳＸＯＳＦ１保証方法＿対
            .HSXOF1HS = rs("HSXOF1HS")                      ' 品ＳＸＯＳＦ１保証方法＿処
            .HSXOF1NS = rs("HSXOF1NS")                      ' 品ＳＸＯＳＦ１熱処理法
            .HSXOF2AX = fncNullCheck(rs("HSXOF2AX"))        ' 品ＳＸＯＳＦ２平均上限    2003/12/10 SystemBrain Null対応
            .HSXOF2MX = fncNullCheck(rs("HSXOF2MX"))        ' 品ＳＸＯＳＦ２上限        2003/12/10 SystemBrain Null対応
            .HSXOF2SH = rs("HSXOF2SH")                      ' 品ＳＸＯＳＦ２測定位置＿方
            .HSXOF2ST = rs("HSXOF2ST")                      ' 品ＳＸＯＳＦ２測定位置＿点
            .HSXOF2SR = rs("HSXOF2SR")                      ' 品ＳＸＯＳＦ２測定位置＿領
            .HSXOF2HT = rs("HSXOF2HT")                      ' 品ＳＸＯＳＦ２保証方法＿対
            .HSXOF2HS = rs("HSXOF2HS")                      ' 品ＳＸＯＳＦ２保証方法＿処
            .HSXOF2NS = rs("HSXOF2NS")                      ' 品ＳＸＯＳＦ２熱処理法
            .HSXOF3AX = fncNullCheck(rs("HSXOF3AX"))        ' 品ＳＸＯＳＦ３平均上限    2003/12/10 SystemBrain Null対応
            .HSXOF3MX = fncNullCheck(rs("HSXOF3MX"))        ' 品ＳＸＯＳＦ３上限        2003/12/10 SystemBrain Null対応
            .HSXOF3SH = rs("HSXOF3SH")                      ' 品ＳＸＯＳＦ３測定位置＿方
            .HSXOF3ST = rs("HSXOF3ST")                      ' 品ＳＸＯＳＦ３測定位置＿点
            .HSXOF3SR = rs("HSXOF3SR")                      ' 品ＳＸＯＳＦ３測定位置＿領
            .HSXOF3HT = rs("HSXOF3HT")                      ' 品ＳＸＯＳＦ３保証方法＿対
            .HSXOF3HS = rs("HSXOF3HS")                      ' 品ＳＸＯＳＦ３保証方法＿処
            .HSXOF3NS = rs("HSXOF3NS")                      ' 品ＳＸＯＳＦ３熱処理法
            .HSXOF4AX = fncNullCheck(rs("HSXOF4AX"))        ' 品ＳＸＯＳＦ４平均上限    2003/12/10 SystemBrain Null対応
            .HSXOF4MX = fncNullCheck(rs("HSXOF4MX"))        ' 品ＳＸＯＳＦ４上限        2003/12/10 SystemBrain Null対応
            .HSXOF4SH = rs("HSXOF4SH")                      ' 品ＳＸＯＳＦ４測定位置＿方
            .HSXOF4ST = rs("HSXOF4ST")                      ' 品ＳＸＯＳＦ４測定位置＿点
            .HSXOF4SR = rs("HSXOF4SR")                      ' 品ＳＸＯＳＦ４測定位置＿領
            .HSXOF4HT = rs("HSXOF4HT")                      ' 品ＳＸＯＳＦ４保証方法＿対
            .HSXOF4HS = rs("HSXOF4HS")                      ' 品ＳＸＯＳＦ４保証方法＿処
            .HSXOF4NS = rs("HSXOF4NS")                      ' 品ＳＸＯＳＦ４熱処理法
            
            .HSXCNMIN = fncNullCheck(rs("HSXCNMIN"))        ' 品ＳＸ炭素濃度下限        2003/12/10 SystemBrain Null対応
            .HSXCNMAX = fncNullCheck(rs("HSXCNMAX"))        ' 品ＳＸ炭素濃度上限        2003/12/10 SystemBrain Null対応
            .HSXCNSPH = rs("HSXCNSPH")                      ' 品ＳＸ炭素濃度測定位置＿方
            .HSXCNSPT = rs("HSXCNSPT")                      ' 品ＳＸ炭素濃度測定位置＿点
            .HSXCNSPI = rs("HSXCNSPI")                      ' 品ＳＸ炭素濃度測定位置＿位
            .HSXCNHWT = rs("HSXCNHWT")                      ' 品ＳＸ炭素濃度保証方法＿対
            .HSXCNHWS = rs("HSXCNHWS")                      ' 品ＳＸ炭素濃度保証方法＿処
            .HSXCNKHI = rs("HSXCNKHI")                      ' 品ＳＸ炭素濃度検査頻度＿位 09/01/08 ooba

            .HSXDENMX = fncNullCheck(rs("HSXDENMX"))        ' 品ＳＸＤｅｎ上限          2003/12/10 SystemBrain Null対応
            .HSXDENMN = fncNullCheck(rs("HSXDENMN"))        ' 品ＳＸＤｅｎ下限          2003/12/10 SystemBrain Null対応
            .HSXLDLMX = fncNullCheck(rs("HSXLDLMX"))        ' 品ＳＸＬ／ＤＬ上限        2003/12/10 SystemBrain Null対応
            .HSXLDLMN = fncNullCheck(rs("HSXLDLMN"))        ' 品ＳＸＬ／ＤＬ下限        2003/12/10 SystemBrain Null対応
            .HSXDVDMX = fncNullCheck(rs("HSXDVDMXN"))       ' 品ＳＸＤＶＤ２上限   項目追加，修正対応 2003.05.20 yakimura   2003/12/10 SystemBrain Null対応
            .HSXDVDMN = fncNullCheck(rs("HSXDVDMNN"))       ' 品ＳＸＤＶＤ２下限   項目追加，修正対応 2003.05.20 yakimura   2003/12/10 SystemBrain Null対応
            .HSXDENHT = rs("HSXDENHT")                      ' 品ＳＸＤｅｎ保証方法＿対
            .HSXDENHS = rs("HSXDENHS")                      ' 品ＳＸＤｅｎ保証方法＿処
            .HSXLDLHT = rs("HSXLDLHT")                      ' 品ＳＸＬ／ＤＬ保証方法＿対
            .HSXLDLHS = rs("HSXLDLHS")                      ' 品ＳＸＬ／ＤＬ保証方法＿処
            .HSXDVDHT = rs("HSXDVDHT")                      ' 品ＳＸＤＶＤ２保証方法＿対
            .HSXDVDHS = rs("HSXDVDHS")                      ' 品ＳＸＤＶＤ２保証方法＿処
            .HSXDENKU = rs("HSXDENKU")                      ' 品ＳＸＤｅｎ検査有無
            .HSXDVDKU = rs("HSXDVDKU")                      ' 品ＳＸＤＶＤ２検査有無
            .HSXLDLKU = rs("HSXLDLKU")                      ' 品ＳＸＬ／ＤＬ検査有無
        '*** UPDATE ↓ Y.SIMIZU 2005/10/12 GDﾗｲﾝ数追加
            .HSXGDLINE = fncNullCheck(rs("HSXGDLINE"))      ' 品ＳＸＬＧＤライン数
        '*** UPDATE ↑ Y.SIMIZU 2005/10/12 GDﾗｲﾝ数追加
            .HSXLTMIN = fncNullCheck(rs("HSXLTMIN"))        ' 品ＳＸＬタイム下限        2003/12/10 SystemBrain Null対応
            .HSXLTMAX = fncNullCheck(rs("HSXLTMAX"))        ' 品ＳＸＬタイム上限        2003/12/10 SystemBrain Null対応
''Add Start 2011/07/25 LT10Ω判定追加対応 T.Koi(SETsw)
            .HSXLT10MIN = fncNullCheck(rs("LTCONVAL"))      ' 品ＳＸＬLT10下限
''Add End   2011/07/25 LT10Ω判定追加対応 T.Koi(SETsw)
            .HSXLTSPH = rs("HSXLTSPH")                      ' 品ＳＸＬタイム測定位置＿方
            .HSXLTSPT = rs("HSXLTSPT")                      ' 品ＳＸＬタイム測定位置＿点
            .HSXLTSPI = rs("HSXLTSPI")                      ' 品ＳＸＬタイム測定位置＿位
            .HSXLTHWT = rs("HSXLTHWT")                      ' 品ＳＸＬタイム保証方法＿対
            .HSXLTHWS = rs("HSXLTHWS")                      ' 品ＳＸＬタイム保証方法＿処
            
            'Null対応 2003/10/22 SystemBrain ↓
            .EPDUP = fncNullCheck(rs("EPDUP"))              ' EPD上限                   2003/12/10 SystemBrain Null対応
            .TOPREG = fncNullCheck(rs("TOPREG"))            ' TOP規制                   2003/12/10 SystemBrain Null対応
            .TAILREG = fncNullCheck(rs("TAILREG"))          ' TAIL規制                  2003/12/10 SystemBrain Null対応
            .BTMSPRT = fncNullCheck(rs("BTMSPRT"))          ' ボトム析出規制            2003/12/10 SystemBrain Null対応
            'Null対応 2003/10/22 SystemBrain ↑

'C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
            If IsNull(rs("COSF3FLAG")) = False Then .COSF3FLAG = rs("COSF3FLAG") Else .COSF3FLAG = " "            'C-OSF3ﾌﾗｸﾞ
'C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---

'--------------- 2008/08/25 INSERT START  By Systech ---------------
            .HSXDKTMP = rs("HSXDKTMP")
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
            .HSXLDLRMN = fncNullCheck(rs("HSXLDLRMN"))      ' 品SXL/DL連続0下限
            .HSXLDLRMX = fncNullCheck(rs("HSXLDLRMX"))      ' 品SXL/DL連続0上限
            .HWFLDLRMN = fncNullCheck(rs("HWFLDLRMN"))      ' 品WFL/DL連続0下限
            .HWFLDLRMX = fncNullCheck(rs("HWFLDLRMX"))      ' 品WFL/DL連続0上限
            If IsNull(rs("HSXOF1ARPTK")) = False Then .HSXOF1ARPTK = rs("HSXOF1ARPTK") Else .HSXOF1ARPTK = " "  ' 品SXOSF1(ArAN)パタン区分
            .HSXOFARMIN = fncNullCheck(rs("HSXOFARMIN"))    ' 品SXOSF(ArAN)下限
            .HSXOFARMAX = fncNullCheck(rs("HSXOFARMAX"))    ' 品SXOSF(ArAN)上限
            .HSXOFARMHMX = fncNullCheck(rs("HSXOFARMHMX"))  ' 品SXOSF(ArAN)面内比上限
            If IsNull(rs("HSXGDPTK")) = False Then .HSXGDPTK = rs("HSXGDPTK") Else .HSXGDPTK = " "  ' 品ＳＸＧＤパタン区分
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End

' OSF，BMD項目追加対応  2002.04.02 yakimura
            If IsNull(rs("HSXOSF1PTK")) = False Then .HSXOSF1PTK = rs("HSXOSF1PTK")   ' 品ＳＸＯＳＦ１パタン区分
            If IsNull(rs("HSXOSF2PTK")) = False Then .HSXOSF2PTK = rs("HSXOSF2PTK")   ' 品ＳＸＯＳＦ２パタン区分
            If IsNull(rs("HSXOSF3PTK")) = False Then .HSXOSF3PTK = rs("HSXOSF3PTK")   ' 品ＳＸＯＳＦ３パタン区分
            If IsNull(rs("HSXOSF4PTK")) = False Then .HSXOSF4PTK = rs("HSXOSF4PTK")   ' 品ＳＸＯＳＦ４パタン区分
            
            .HSXBMD1MBP = fncNullCheck(rs("HSXBMD1MBP"))    ' 品ＳＸＢＭＤ１面内分布    2003/12/10 SystemBrain Null対応
            .HSXBMD2MBP = fncNullCheck(rs("HSXBMD2MBP"))    ' 品ＳＸＢＭＤ２面内分布    2003/12/10 SystemBrain Null対応
            .HSXBMD3MBP = fncNullCheck(rs("HSXBMD3MBP"))    ' 品ＳＸＢＭＤ３面内分布    2003/12/10 SystemBrain Null対応
' OSF，BMD項目追加対応  2002.04.02 yakimura
            
            'WF仕様取得　08/04/15 ooba START ============================================>
            .HWFRHWYS = rs("HWFRHWYS")                      ' 品ＷＦ比抵抗保証方法＿処
            .HWFONHWS = rs("HWFONHWS")                      ' 品ＷＦ酸素濃度保証方法＿処
            .HWFOF1HS = rs("HWFOF1HS")                      ' 品ＷＦＯＳＦ１保証方法＿処
            .HWFOF2HS = rs("HWFOF2HS")                      ' 品ＷＦＯＳＦ２保証方法＿処
            .HWFOF3HS = rs("HWFOF3HS")                      ' 品ＷＦＯＳＦ３保証方法＿処
            .HWFOF4HS = rs("HWFOF4HS")                      ' 品ＷＦＯＳＦ４保証方法＿処
            .HWFBM1HS = rs("HWFBM1HS")                      ' 品ＷＦＢＭＤ１保証方法＿処
            .HWFBM2HS = rs("HWFBM2HS")                      ' 品ＷＦＢＭＤ２保証方法＿処
            .HWFBM3HS = rs("HWFBM3HS")                      ' 品ＷＦＢＭＤ３保証方法＿処
            .HWFDENHS = rs("HWFDENHS")                      ' 品ＷＦＤｅｎ保証方法＿処
            .HWFDVDHS = rs("HWFDVDHS")                      ' 品ＷＦＤＶＤ２保証方法＿処
            .HWFLDLHS = rs("HWFLDLHS")                      ' 品ＷＦＬ／ＤＬ保証方法＿処
            .HWFRKHNN = rs("HWFRKHNN")                      ' 品ＷＦ比抵抗検査頻度＿抜
            .HWFONKHN = rs("HWFONKHN")                      ' 品ＷＦ酸素濃度検査頻度＿抜
            .HWFOF1KN = rs("HWFOF1KN")                      ' 品ＷＦＯＳＦ１検査頻度＿抜
            .HWFOF2KN = rs("HWFOF2KN")                      ' 品ＷＦＯＳＦ２検査頻度＿抜
            .HWFOF3KN = rs("HWFOF3KN")                      ' 品ＷＦＯＳＦ３検査頻度＿抜
            .HWFOF4KN = rs("HWFOF4KN")                      ' 品ＷＦＯＳＦ４検査頻度＿抜
            .HWFBM1KN = rs("HWFBM1KN")                      ' 品ＷＦＢＭＤ１検査頻度＿抜
            .HWFBM2KN = rs("HWFBM2KN")                      ' 品ＷＦＢＭＤ２検査頻度＿抜
            .HWFBM3KN = rs("HWFBM3KN")                      ' 品ＷＦＢＭＤ３検査頻度＿抜
            .HWFGDKHN = rs("HWFGDKHN")                      ' 品ＷＦＧＤ検査頻度＿抜
            'WF仕様取得　08/04/15 ooba END ==============================================>
            .HSXCSCEN = fncNullCheck(rs("HSXCSCEN"))        ' 品ＳＸ面傾き中心  2009/08/12 Kameda
            .HSXCSMIN = fncNullCheck(rs("HSXCSMIN"))        ' 品ＳＸ面傾き下限  2009/08/12 Kameda
            .HSXCSMAX = fncNullCheck(rs("HSXCSMAX"))        ' 品ＳＸ面傾き上限  2009/08/12 Kameda
            .HSXCTCEN = fncNullCheck(rs("HSXCTCEN"))        ' 品ＳＸ面傾き上限  2009/09/01 Kameda
            .HSXCTMIN = fncNullCheck(rs("HSXCTMIN"))        ' 品ＳＸ面傾き下限  2009/09/01 Kameda
            .HSXCTMAX = fncNullCheck(rs("HSXCTMAX"))        ' 品ＳＸ面傾き上限  2009/09/01 Kameda
            .HSXCYCEN = fncNullCheck(rs("HSXCYCEN"))        ' 品ＳＸ面傾き上限  2009/09/01 Kameda
            .HSXCYMIN = fncNullCheck(rs("HSXCYMIN"))        ' 品ＳＸ面傾き下限  2009/09/01 Kameda
            .HSXCYMAX = fncNullCheck(rs("HSXCYMAX"))        ' 品ＳＸ面傾き上限  2009/09/01 Kameda
            .HWFSIRDMX = fncNullCheck(rs("HWFSIRDMX"))      ' 品面内個数上限    2010/02/04 Kameda
            
  'Add Start 2011/01/17 SMPK A.Nagamine : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の仕様項目追加
            If IsNull(rs("HSXCPK")) = False Then .HSXCPK = rs("HSXCPK") Else .HSXCPK = " "              ' 品ＳＸＣパターン区分
            If IsNull(rs("HSXCSZ")) = False Then .HSXCSZ = rs("HSXCSZ") Else .HSXCSZ = " "              ' 品ＳＸＣ測定条件
            If IsNull(rs("HSXCHT")) = False Then .HSXCHT = rs("HSXCHT") Else .HSXCHT = " "              ' 品ＳＸＣ保証方法＿対
            If IsNull(rs("HSXCHS")) = False Then .HSXCHS = rs("HSXCHS") Else .HSXCHS = " "              ' 品ＳＸＣ保証方法＿処
            If IsNull(rs("HSXCJPK")) = False Then .HSXCJPK = rs("HSXCJPK") Else .HSXCJPK = " "          ' 品ＳＸＣＪパターン区分
            If IsNull(rs("HSXCJNS")) = False Then .HSXCJNS = rs("HSXCJNS") Else .HSXCJNS = " "          ' 品ＳＸＣＪ熱処理法
            If IsNull(rs("HSXCJHT")) = False Then .HSXCJHT = rs("HSXCJHT") Else .HSXCJHT = " "          ' 品ＳＸＣＪ保証方法＿対
            If IsNull(rs("HSXCJHS")) = False Then .HSXCJHS = rs("HSXCJHS") Else .HSXCJHS = " "          ' 品ＳＸＣＪ保証方法＿処
            If IsNull(rs("HSXCJLTPK")) = False Then .HSXCJLTPK = rs("HSXCJLTPK") Else .HSXCJLTPK = " "  ' 品ＳＸＣＪＬＴパターン区分
            If IsNull(rs("HSXCJLTNS")) = False Then .HSXCJLTNS = rs("HSXCJLTNS") Else .HSXCJLTNS = " "  ' 品ＳＸＣＪＬＴ熱処理法
            If IsNull(rs("HSXCJLTHT")) = False Then .HSXCJLTHT = rs("HSXCJLTHT") Else .HSXCJLTHT = " "  ' 品ＳＸＣＪＬＴ保証方法＿対
            If IsNull(rs("HSXCJLTHS")) = False Then .HSXCJLTHS = rs("HSXCJLTHS") Else .HSXCJLTHS = " "  ' 品ＳＸＣＪＬＴ保証方法＿処
            If IsNull(rs("HSXCJ2PK")) = False Then .HSXCJ2PK = rs("HSXCJ2PK") Else .HSXCJ2PK = " "      ' 品ＳＸＣＪ２パターン区分
            If IsNull(rs("HSXCJ2NS")) = False Then .HSXCJ2NS = rs("HSXCJ2NS") Else .HSXCJ2NS = " "      ' 品ＳＸＣＪ２熱処理法
            If IsNull(rs("HSXCJ2HT")) = False Then .HSXCJ2HT = rs("HSXCJ2HT") Else .HSXCJ2HT = " "      ' 品ＳＸＣＪ２保証方法＿対
            If IsNull(rs("HSXCJ2HS")) = False Then .HSXCJ2HS = rs("HSXCJ2HS") Else .HSXCJ2HS = " "      ' 品ＳＸＣＪ２保証方法＿処
            .HSXCJLTBND = fncNullCheck(rs("HSXCJLTBND"))                                                ' 品SXL/CJLTバンド幅 Number(3,0)
    
  'Add End   2011/01/17 SMPK A.Nagamine
  
  'Add Start 2012/06/01 SMPK H.Ohkubo
            If IsNull(rs("HSXCOSF3PK")) = False Then .HSXCOSF3PK = rs("HSXCOSF3PK") Else .HSXCOSF3PK = " "  '品ＳＸＣＯＳＦ３パターン区分"
  'Add End 2012/06/01 SMPK H.Ohkubo
  
        End With
        rs.MoveNext
    Next

    If scmzc_getKakouJiltuseki(inBlockID, Jiltuseki) = FUNCTION_RETURN_FAILURE Then
        getHinSiyou = FUNCTION_RETURN_FAILURE
        ReDim siyou(0)
        GoTo proc_exit
    End If
    For i = 1 To recCnt
        siyou(i).DIAMETER = (Jiltuseki.top(1) + Jiltuseki.top(2) + Jiltuseki.TAIL(1) + Jiltuseki.TAIL(2)) / 4 ' 直径
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

'概要      :内部関数 サンプル番号を取得する
'概要      :総合判定 各種データ取得
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                 ,説明
'          :inBlockID     ,I  ,String                             ,対象ブロックID
'          :CrySmp()      ,O  ,type_DBDRV_scmzc_fcmkc001c_CrySmp  ,結晶サンプル管理取得用
'          :iSmpGetFlg    ,I  ,Integer                            :ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ(0:ｻﾝﾌﾟﾙ指定なし, 1:ｻﾝﾌﾟﾙ指定あり)
'          :iSamplID1     ,I  ,Long                               :TOPｻﾝﾌﾟﾙID(省略可)   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :iSamplID2     ,I  ,Long                               :BOTｻﾝﾌﾟﾙID(省略可)   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :戻り値        ,O  ,FUNCTION_RETURN                    ,読み込み成否
'説明      :
'履歴      :2001/06/26 蔵本 作成
Private Function getCrySmp(inBlockID As String, _
                           CrySmp() As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                           iSmpGetFlg As Integer, iSamplID1 As Long, iSamplID2 As Long) As FUNCTION_RETURN
    
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim recCnt      As Integer
    Dim i           As Long
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    Dim wkXsdcs     As typ_XSDCS
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function getCrySmp"

    If iSmpGetFlg = 0 Then          'ﾌﾞﾛｯｸIDで検索(生死区分=生ﾛｯﾄ)
        'ﾌﾞﾛｯｸ管理(TBCME040)参照停止　05/10/03 ooba
        sql = "select CS.CRYNUMCS, CSBOT.INPOSCS - CSTOP.INPOSCS as LENGTH, CS.SMPKBNCS, CS.TBKBNCS, CS.REPSMPLIDCS, CS.XTALCS, CS.INPOSCS, "
        sql = sql & "CS.HINBCS, CS.REVNUMCS, CS.FACTORYCS, CS.OPECS, CS.KTKBNCS, CS.BLKKTFLAGCS, "
        sql = sql & "CS.CRYSMPLIDRSCS, CS.CRYSMPLIDRS1CS, CS.CRYSMPLIDRS2CS, CS.CRYINDRSCS, CS.CRYRESRS1CS, CS.CRYRESRS2CS, "
        sql = sql & "CS.CRYSMPLIDOICS, CS.CRYINDOICS, CS.CRYRESOICS, CS.CRYSMPLIDB1CS, CS.CRYINDB1CS, CS.CRYRESB1CS, "
        sql = sql & "CS.CRYSMPLIDB2CS, CS.CRYINDB2CS, CS.CRYRESB2CS, CS.CRYSMPLIDB3CS, CS.CRYINDB3CS, CS.CRYRESB3CS, "
        sql = sql & "CS.CRYSMPLIDL1CS, CS.CRYINDL1CS, CS.CRYRESL1CS, CS.CRYSMPLIDL2CS, CS.CRYINDL2CS, CS.CRYRESL2CS, "
        sql = sql & "CS.CRYSMPLIDL3CS, CS.CRYINDL3CS, CS.CRYRESL3CS, CS.CRYSMPLIDL4CS, CS.CRYINDL4CS, CS.CRYRESL4CS, "
        sql = sql & "CS.CRYSMPLIDCSCS, CS.CRYINDCSCS, CS.CRYRESCSCS, CS.CRYSMPLIDGDCS, CS.CRYINDGDCS, CS.CRYRESGDCS, "
        sql = sql & "CS.CRYSMPLIDTCS, CS.CRYINDTCS, CS.CRYRESTCS, CS.CRYREST10CS, CS.CRYSMPLIDEPCS, CS.CRYINDEPCS, CS.CRYRESEPCS "
      'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の仕様項目追加
        sql = sql & ", CS.CRYSMPLIDCCS, CS.CRYINDCCS, CS.CRYRESCCS, CS.CRYSMPLIDCJCS, CS.CRYINDCJCS"
        sql = sql & ", CS.CRYRESCJCS, CS.CRYSMPLIDCJLTCS, CS.CRYINDCJLTCS, CS.CRYRESCJLTCS, CS.CRYSMPLIDCJ2CS"
        sql = sql & ", CS.CRYINDCJ2CS, CS.CRYRESCJ2CS "
      'Add End   2011/01/17 SMPK A.Nagamine
        sql = sql & "from XSDCS CS, "
        sql = sql & "(select CRYNUMCS, XTALCS, INPOSCS from XSDCS "
        sql = sql & "where TBKBNCS = 'T' and CRYNUMCS = '" & inBlockID & "' "
        sql = sql & ") CSTOP, "
        sql = sql & "(select CRYNUMCS, XTALCS, INPOSCS from XSDCS "
        sql = sql & "where TBKBNCS = 'B' and CRYNUMCS = '" & inBlockID & "' "
        sql = sql & ") CSBOT "
        sql = sql & "where CSTOP.CRYNUMCS = CSBOT.CRYNUMCS and "
        
        sql = sql & "CS.CRYNUMCS = '" & inBlockID & "' and "
        sql = sql & "CS.LIVKCS = '0'"
    
    Else                            '結晶番号とｻﾝﾌﾟﾙIDで検索
        sql = "select CS.CRYNUMCS, 0 as LENGTH, CS.SMPKBNCS, CS.TBKBNCS, CS.REPSMPLIDCS, CS.XTALCS, CS.INPOSCS, "
        sql = sql & "CS.HINBCS, CS.REVNUMCS, CS.FACTORYCS, CS.OPECS, CS.KTKBNCS, CS.BLKKTFLAGCS, "
        sql = sql & "CS.CRYSMPLIDRSCS, CS.CRYSMPLIDRS1CS, CS.CRYSMPLIDRS2CS, CS.CRYINDRSCS, CS.CRYRESRS1CS, CS.CRYRESRS2CS, "
        sql = sql & "CS.CRYSMPLIDOICS, CS.CRYINDOICS, CS.CRYRESOICS, CS.CRYSMPLIDB1CS, CS.CRYINDB1CS, CS.CRYRESB1CS, "
        sql = sql & "CS.CRYSMPLIDB2CS, CS.CRYINDB2CS, CS.CRYRESB2CS, CS.CRYSMPLIDB3CS, CS.CRYINDB3CS, CS.CRYRESB3CS, "
        sql = sql & "CS.CRYSMPLIDL1CS, CS.CRYINDL1CS, CS.CRYRESL1CS, CS.CRYSMPLIDL2CS, CS.CRYINDL2CS, CS.CRYRESL2CS, "
        sql = sql & "CS.CRYSMPLIDL3CS, CS.CRYINDL3CS, CS.CRYRESL3CS, CS.CRYSMPLIDL4CS, CS.CRYINDL4CS, CS.CRYRESL4CS, "
        sql = sql & "CS.CRYSMPLIDCSCS, CS.CRYINDCSCS, CS.CRYRESCSCS, CS.CRYSMPLIDGDCS, CS.CRYINDGDCS, CS.CRYRESGDCS, "
        sql = sql & "CS.CRYSMPLIDTCS, CS.CRYINDTCS, CS.CRYRESTCS, CS.CRYREST10CS, CS.CRYSMPLIDEPCS, CS.CRYINDEPCS, CS.CRYRESEPCS "
      'Add Start 2011/01/17 SMPK A.Nagamine     : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の仕様項目追加
        sql = sql & ", CS.CRYSMPLIDCCS, CS.CRYINDCCS, CS.CRYRESCCS, CS.CRYSMPLIDCJCS, CS.CRYINDCJCS"
        sql = sql & ", CS.CRYRESCJCS, CS.CRYSMPLIDCJLTCS, CS.CRYINDCJLTCS, CS.CRYRESCJLTCS, CS.CRYSMPLIDCJ2CS"
        sql = sql & ", CS.CRYINDCJ2CS, CS.CRYRESCJ2CS "
      'Add End   2011/01/17 SMPK A.Nagamine
        sql = sql & "from XSDCS CS "
        sql = sql & "where substr(CS.CRYNUMCS, 1, 10) = substr('" & inBlockID & "', 1, 10) and "
        sql = sql & "CS.REPSMPLIDCS in (" & iSamplID1 & ", " & iSamplID2 & ")"
    End If
    
    sql = sql & "order by CS.INPOSCS "  ' TOP TAIL順
    ' SQL実行
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount = 0 Then
        getCrySmp = FUNCTION_RETURN_FAILURE
        ReDim CrySmp(0)
        GoTo proc_exit
    End If
    
    recCnt = rs.RecordCount
    ReDim CrySmp(recCnt)
    For i = 1 To recCnt
        With CrySmp(i)
            .CRYNUMCS = rs("CRYNUMCS")          'ブロックID
            .Length = rs("LENGTH")              ' 長さ
            If IsNull(rs("SMPKBNCS")) = False Then .SMPKBNCS = rs("SMPKBNCS")                   ' サンプル区分
            .TBKBNCS = rs("TBKBNCS")            'T/B区分
            .REPSMPLIDCS = rs("REPSMPLIDCS")    ' 代表サンプルID
            
            If IsNull(rs("XTALCS")) = False Then .XTALCS = rs("XTALCS")                         ' 結晶番号
            If IsNull(rs("INPOSCS")) = False Then .INPOSCS = rs("INPOSCS")                      ' 結晶内位置
            If IsNull(rs("HINBCS")) = False Then .HINBCS = rs("HINBCS")                         ' 品番
            If IsNull(rs("REVNUMCS")) = False Then .REVNUMCS = rs("REVNUMCS")                   ' 製品番号改訂番号
            If IsNull(rs("FACTORYCS")) = False Then .FACTORYCS = rs("FACTORYCS")                ' 工場
            If IsNull(rs("OPECS")) = False Then .OPECS = rs("OPECS")                            ' 操業条件
            If IsNull(rs("KTKBNCS")) = False Then .KTKBNCS = rs("KTKBNCS")                      ' 確定区分
            If IsNull(rs("BLKKTFLAGCS")) = False Then .BLKKTFLAGCS = rs("BLKKTFLAGCS")          ' ブロック確定フラグ
            If IsNull(rs("CRYSMPLIDRSCS")) = False Then .CRYSMPLIDRSCS = rs("CRYSMPLIDRSCS")    ' サンプルID(Rs)
            If IsNull(rs("CRYSMPLIDRS1CS")) = False Then .CRYSMPLIDRS1CS = rs("CRYSMPLIDRS1CS") ' 推定サンプルID1(Rs)
            If IsNull(rs("CRYSMPLIDRS2CS")) = False Then .CRYSMPLIDRS2CS = rs("CRYSMPLIDRS2CS") ' 推定サンプルID2(Rs)
            If IsNull(rs("CRYINDRSCS")) = False Then .CRYINDRSCS = rs("CRYINDRSCS")             ' 状態FLG(Rs)
            If IsNull(rs("CRYRESRS1CS")) = False Then .CRYRESRS1CS = rs("CRYRESRS1CS")          ' 実績FLG1(Rs)
            If IsNull(rs("CRYRESRS2CS")) = False Then .CRYRESRS2CS = rs("CRYRESRS2CS")          ' 実績FLG2(Rs)
            If IsNull(rs("CRYSMPLIDOICS")) = False Then .CRYSMPLIDOICS = rs("CRYSMPLIDOICS")    ' サンプルID(Oi)
            If IsNull(rs("CRYINDOICS")) = False Then .CRYINDOICS = rs("CRYINDOICS")             ' 状態FLG(Oi)
            If IsNull(rs("CRYRESOICS")) = False Then .CRYRESOICS = rs("CRYRESOICS")             ' 実績FLG(Oi)
            If IsNull(rs("CRYSMPLIDB1CS")) = False Then .CRYSMPLIDB1CS = rs("CRYSMPLIDB1CS")    ' サンプルID(B1)
            If IsNull(rs("CRYINDB1CS")) = False Then .CRYINDB1CS = rs("CRYINDB1CS")             ' 状態FLG(B1)
            If IsNull(rs("CRYRESB1CS")) = False Then .CRYRESB1CS = rs("CRYRESB1CS")             ' 実績FLG(B1)
            If IsNull(rs("CRYSMPLIDB2CS")) = False Then .CRYSMPLIDB2CS = rs("CRYSMPLIDB2CS")    ' サンプルID(B2)
            If IsNull(rs("CRYINDB2CS")) = False Then .CRYINDB2CS = rs("CRYINDB2CS")             ' 状態FLG(B2)
            If IsNull(rs("CRYRESB2CS")) = False Then .CRYRESB2CS = rs("CRYRESB2CS")             ' 実績FLG(B2)
            If IsNull(rs("CRYSMPLIDB3CS")) = False Then .CRYSMPLIDB3CS = rs("CRYSMPLIDB3CS")    ' サンプルID(B3)
            If IsNull(rs("CRYINDB3CS")) = False Then .CRYINDB3CS = rs("CRYINDB3CS")             ' 状態FLG(B3)
            If IsNull(rs("CRYRESB3CS")) = False Then .CRYRESB3CS = rs("CRYRESB3CS")             ' 実績FLG(B3)
            If IsNull(rs("CRYSMPLIDL1CS")) = False Then .CRYSMPLIDL1CS = rs("CRYSMPLIDL1CS")    ' サンプルID(L1)
            If IsNull(rs("CRYINDL1CS")) = False Then .CRYINDL1CS = rs("CRYINDL1CS")             ' 状態FLG(L1)
            If IsNull(rs("CRYRESL1CS")) = False Then .CRYRESL1CS = rs("CRYRESL1CS")             ' 実績FLG(L1)
            If IsNull(rs("CRYSMPLIDL2CS")) = False Then .CRYSMPLIDL2CS = rs("CRYSMPLIDL2CS")    ' サンプルID(L2)
            If IsNull(rs("CRYINDL2CS")) = False Then .CRYINDL2CS = rs("CRYINDL2CS")             ' 状態FLG(L2)
            If IsNull(rs("CRYRESL2CS")) = False Then .CRYRESL2CS = rs("CRYRESL2CS")             ' 実績FLG(L2)
            If IsNull(rs("CRYSMPLIDL3CS")) = False Then .CRYSMPLIDL3CS = rs("CRYSMPLIDL3CS")    ' サンプルID(L3)
            If IsNull(rs("CRYINDL3CS")) = False Then .CRYINDL3CS = rs("CRYINDL3CS")             ' 状態FLG(L3)
            If IsNull(rs("CRYRESL3CS")) = False Then .CRYRESL3CS = rs("CRYRESL3CS")             ' 実績FLG(L3)
            If IsNull(rs("CRYSMPLIDL4CS")) = False Then .CRYSMPLIDL4CS = rs("CRYSMPLIDL4CS")    ' サンプルID(L4)
            If IsNull(rs("CRYINDL4CS")) = False Then .CRYINDL4CS = rs("CRYINDL4CS")             ' 状態FLG(L4)
            If IsNull(rs("CRYRESL4CS")) = False Then .CRYRESL4CS = rs("CRYRESL4CS")             ' 実績FLG(L4)
            If IsNull(rs("CRYSMPLIDCSCS")) = False Then .CRYSMPLIDCSCS = rs("CRYSMPLIDCSCS")    ' サンプルID(Cs)
            If IsNull(rs("CRYINDCSCS")) = False Then .CRYINDCSCS = rs("CRYINDCSCS")             ' 状態FLG(Cs)
            If IsNull(rs("CRYRESCSCS")) = False Then .CRYRESCSCS = rs("CRYRESCSCS")             ' 実績FLG(Cs)
            If IsNull(rs("CRYSMPLIDGDCS")) = False Then .CRYSMPLIDGDCS = rs("CRYSMPLIDGDCS")    ' サンプルID(GD)
            If IsNull(rs("CRYINDGDCS")) = False Then .CRYINDGDCS = rs("CRYINDGDCS")             ' 状態FLG(GD)
            If IsNull(rs("CRYRESGDCS")) = False Then .CRYRESGDCS = rs("CRYRESGDCS")             ' 実績FLG(GD)
            If IsNull(rs("CRYSMPLIDTCS")) = False Then .CRYSMPLIDTCS = rs("CRYSMPLIDTCS")       ' サンプルID(T)
            If IsNull(rs("CRYINDTCS")) = False Then .CRYINDTCS = rs("CRYINDTCS")                ' 状態FLG(T)
            If IsNull(rs("CRYRESTCS")) = False Then .CRYRESTCS = rs("CRYRESTCS")                ' 実績FLG(T)
''Add Start 2011/07/25 LT10Ω判定追加対応 T.Koi(SETsw)
            If IsNull(rs("CRYREST10CS")) = False Then .CRYREST10CS = rs("CRYREST10CS")                ' 実績FLG(T)
''Add End   2011/07/25 LT10Ω判定追加対応 T.Koi(SETsw)
            If IsNull(rs("CRYSMPLIDEPCS")) = False Then .CRYSMPLIDEPCS = rs("CRYSMPLIDEPCS")    ' サンプルID(EPD)
            If IsNull(rs("CRYINDEPCS")) = False Then .CRYINDEPCS = rs("CRYINDEPCS")             ' 状態FLG(EPD)
            If IsNull(rs("CRYRESEPCS")) = False Then .CRYRESEPCS = rs("CRYRESEPCS")             ' 実績FLG(EPD)
'--------------- 2008/08/25 INSERT START  By Systech ---------------
            ' DK温度（実績）
            wkXsdcs.HINBCS = .HINBCS
            wkXsdcs.REVNUMCS = .REVNUMCS
            wkXsdcs.FACTORYCS = .FACTORYCS
            wkXsdcs.OPECS = .OPECS
            wkXsdcs.XTALCS = .XTALCS
            wkXsdcs.CRYSMPLIDRSCS = .CRYSMPLIDRSCS
            wkXsdcs.CRYINDRSCS = .CRYINDRSCS
            .HSXDKTMP = GetDKTmpCode(False, wkXsdcs)
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
            
          'Add Start 2011/01/17 SMPK A.Nagamine : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の仕様項目追加
            If IsNull(rs("CRYSMPLIDCCS")) = False Then .CRYSMPLIDCCS = rs("CRYSMPLIDCCS")           ' サンプルID(C)
            If IsNull(rs("CRYINDCCS")) = False Then .CRYINDCCS = rs("CRYINDCCS")                    ' 状態FLG(C)
            If IsNull(rs("CRYRESCCS")) = False Then .CRYRESCCS = rs("CRYRESCCS")                    ' 実績FLG(C)
            If IsNull(rs("CRYSMPLIDCJCS")) = False Then .CRYSMPLIDCJCS = rs("CRYSMPLIDCJCS")        ' サンプルID(CJ)
            If IsNull(rs("CRYINDCJCS")) = False Then .CRYINDCJCS = rs("CRYINDCJCS")                 ' 状態FLG(CJ)
            If IsNull(rs("CRYRESCJCS")) = False Then .CRYRESCJCS = rs("CRYRESCJCS")                 ' 実績FLG(CJ)
            If IsNull(rs("CRYSMPLIDCJLTCS")) = False Then .CRYSMPLIDCJLTCS = rs("CRYSMPLIDCJLTCS")  ' サンプルID(CJ[LT])
            If IsNull(rs("CRYINDCJLTCS")) = False Then .CRYINDCJLTCS = rs("CRYINDCJLTCS")           ' 状態FLG(CJ[LT])
            If IsNull(rs("CRYRESCJLTCS")) = False Then .CRYRESCJLTCS = rs("CRYRESCJLTCS")           ' 実績FLG(CJ[LT])
            If IsNull(rs("CRYSMPLIDCJ2CS")) = False Then .CRYSMPLIDCJ2CS = rs("CRYSMPLIDCJ2CS")     ' サンプルID(CJ2)
            If IsNull(rs("CRYINDCJ2CS")) = False Then .CRYINDCJ2CS = rs("CRYINDCJ2CS")              ' 状態FLG(CJ2)
            If IsNull(rs("CRYRESCJ2CS")) = False Then .CRYRESCJ2CS = rs("CRYRESCJ2CS")              ' 実績FLG(CJ2)
          'Add End   2011/01/17 SMPK A.Nagamine
            
        End With
        rs.MoveNext
    Next
    rs.Close
    
    getCrySmp = FUNCTION_RETURN_SUCCESS
    

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getCrySmp = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :内部関数 結晶抵抗実績取得用
Private Function CryR_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                              Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                              CryR As type_DBDRV_scmzc_fcmkc001c_CryR, _
                              SuCryR As type_DBDRV_scmzc_fcmkc001c_CryR, _
                              TorB As Integer, _
                              Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    Dim wkXsdcs     As typ_XSDCS
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

    NothingFlag = False

    ' 結晶抵抗実績テーブルから値を取得

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function CryR_Zisseki"

    CryR_Zisseki = FUNCTION_RETURN_SUCCESS

    Set rs = Nothing

    ' 推定データの確認と推定データ作成
    If (Samp.CRYINDRSCS = "3") And (Samp.KTKBNCS = "0") And (ciSmpGetFlg = 0) Then
        If (Samp.CRYRESRS1CS <> "0") And (Samp.CRYRESRS2CS <> "0") Then     ' 推定元実績が両方あり
    
            ' 推定データ作成
            If funComputeSuitei(siyou, Samp, CryR) <> 0 Then
                NothingFlag = True
                CryR_Zisseki = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
    
        Else                                                                ' 推定元実績が無い
            NothingFlag = True
            CryR_Zisseki = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    
    ' 指示(仕様)と実績FLGを確認
    ElseIf (Samp.CRYINDRSCS <> "0") And (Samp.CRYRESRS1CS <> "0") And (Samp.KTKBNCS <> "9") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, "
        '----TEST2004/10
        sql = sql & "MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, EFEHS, RRG, REGDATE, KSTAFFID "
        sql = sql & "from TBCMJ002 "
        sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "      SMPLNO = " & Samp.CRYSMPLIDRSCS & " and "
        sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ002 "
        sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "                       SMPLNO = " & Samp.CRYSMPLIDRSCS & ")"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With CryR
                .CRYNUM = rs("CRYNUM")          ' 結晶番号
                .POSITION = rs("POSITION")      ' 位置
                .SMPKBN = rs("SMPKBN")          ' サンプル区分
                .TRANCOND = rs("TRANCOND")      ' 処理条件
                .TRANCNT = rs("TRANCNT")        ' 処理回数
                .SMPLNO = rs("SMPLNO")          ' サンプルＮｏ
                .SMPLUMU = rs("SMPLUMU")        ' サンプル有無
                .MEAS1 = rs("MEAS1")            ' 測定値１
                .MEAS2 = rs("MEAS2")            ' 測定値２
                .MEAS3 = rs("MEAS3")            ' 測定値３
                .MEAS4 = rs("MEAS4")            ' 測定値４
                .MEAS5 = rs("MEAS5")            ' 測定値５
                .EFEHS = rs("EFEHS")            ' 実効偏析
                .RRG = rs("RRG")                ' RRG
                .REGDATE = rs("REGDATE")        ' 登録日付
                '---TEST2004/10
                .KSTAFFID = rs("KSTAFFID")
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

'--------------- 2008/08/25 INSERT START  By Systech ---------------
    ' DK温度（実績）
    wkXsdcs.XTALCS = Samp.XTALCS
    wkXsdcs.CRYSMPLIDRSCS = Samp.CRYSMPLIDRSCS
    wkXsdcs.CRYINDRSCS = "0"
    CryR.HSXDKTMP = GetDKTmpCode(False, wkXsdcs)
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    
    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    CryR_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' 推定データ作成
'------------------------------------------------
'概要      :指定された情報から、推定計算を行ない、推定実績値を作成する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :Siyou         ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :仕様情報構造体
'          :Samp          ,I  ,type_DBDRV_scmzc_fcmkc001c_CrySmp    :新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ)構造体
'          :CryR          ,O  ,type_DBDRV_scmzc_fcmkc001c_CryR      :RS実績構造体
'          :戻り値        ,O  ,Integer                              :結果(0:正常, 1:異常)
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Private Function funComputeSuitei(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                                  Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                                  CryR As type_DBDRV_scmzc_fcmkc001c_CryR) As Integer
    
    Dim tSuiHin         As tFullHinban
    Dim tCryRs(2)       As type_DBDRV_scmzc_fcmkc001c_CryR          '(0)→推定元Top, (1)→推定元Bot, (2)→推定先
    Dim getPtrn1        As String                                   'TOP位置ﾊﾟﾀｰﾝｺｰﾄﾞ
    Dim getPtrn2        As String                                   'BOT位置ﾊﾟﾀｰﾝｺｰﾄﾞ

    Dim retCode         As Integer
    Dim wGetSPtrn1      As String
    Dim wGetSPtrn2      As String
    Dim wcnt            As Integer
    Dim wMeasTop(4)     As Double                   'Top測定値
    Dim wMeasBot(4)     As Double                   'Bot測定値
    Dim wMeasSui()      As Double                   '算出推定値
    Dim retJudg         As Boolean
    
    '新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ)の品番設定
    tSuiHin.hinban = Samp.HINBCS
    tSuiHin.mnorevno = Samp.REVNUMCS
    tSuiHin.factory = Samp.FACTORYCS
    tSuiHin.opecond = Samp.OPECS
    
    '新ｻﾝﾌﾟﾙ管理(XSDCS)の推定元ｻﾝﾌﾟﾙID1から、推定元RS実績値を取得する。
    If funGetCryRsJisseki(Samp.XTALCS, Samp.CRYSMPLIDRS1CS, tCryRs(0)) <> 0 Then GoTo ComputeSuiteiNG

    '新ｻﾝﾌﾟﾙ管理(XSDCS)の推定元ｻﾝﾌﾟﾙID2から、推定元RS実績値を取得する。
    If funGetCryRsJisseki(Samp.XTALCS, Samp.CRYSMPLIDRS2CS, tCryRs(1)) <> 0 Then GoTo ComputeSuiteiNG

    '結晶抵抗実績の処理回数取得
    retCode = funGetTrancntRS(Samp)
    If retCode < 0 Then GoTo ComputeSuiteiSonotaErr

    '推定先の実績データ編集
    With tCryRs(2)
        .CRYNUM = Samp.XTALCS               '結晶番号
        .POSITION = Samp.INPOSCS            '位置
        .SMPKBN = Samp.TBKBNCS              'ｻﾝﾌﾟﾙ区分
        .TRANCOND = "0"                     '処理条件
        .TRANCNT = retCode                  '処理回数
        .SMPLNO = Samp.CRYSMPLIDRSCS        'ｻﾝﾌﾟﾙNo
        .SMPLUMU = "0"                      'ｻﾝﾌﾟﾙ有無
    End With
    
    'Top/Bot測定値を推定値算出用にセット
        wMeasTop(0) = tCryRs(0).MEAS1
        wMeasTop(1) = tCryRs(0).MEAS2
        wMeasTop(2) = tCryRs(0).MEAS3
        wMeasTop(3) = tCryRs(0).MEAS4
        wMeasTop(4) = tCryRs(0).MEAS5
    
        wMeasBot(0) = tCryRs(1).MEAS1
        wMeasBot(1) = tCryRs(1).MEAS2
        wMeasBot(2) = tCryRs(1).MEAS3
        wMeasBot(3) = tCryRs(1).MEAS4
        wMeasBot(4) = tCryRs(1).MEAS5
    
    '推定先の測定点数分、推定値を算出する
    ReDim wMeasSui(4)
    For wcnt = 0 To 4
        
        '推定値の算出
        retCode = new_ResSuitei(Samp.XTALCS, wMeasTop(wcnt), tCryRs(0).POSITION, wMeasBot(wcnt), tCryRs(1).POSITION, Samp.INPOSCS, wMeasSui(wcnt))
        If retCode = FUNCTION_RETURN_FAILURE Then GoTo ComputeSuiteiNG
    
    Next wcnt
    
    '推定値の設定
    tCryRs(2).MEAS1 = wMeasSui(0)
    tCryRs(2).MEAS2 = wMeasSui(1)
    tCryRs(2).MEAS3 = wMeasSui(2)
    tCryRs(2).MEAS4 = wMeasSui(3)
    tCryRs(2).MEAS5 = wMeasSui(4)
    
    CryR = tCryRs(2)
    funComputeSuitei = 0
    Exit Function

ComputeSuiteiNG:
    funComputeSuitei = 0
    Exit Function

ComputeSuiteiSonotaErr:
    funComputeSuitei = -2
End Function

'------------------------------------------------
' 比抵抗推定パターンコード取得
'------------------------------------------------
'概要      :結晶番号と推定元ｻﾝﾌﾟﾙID1と推定元ｻﾝﾌﾟﾙID2から、新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ)(XSDCS)を検索し、それぞれの品番を取得する。
'           推定元1,推定元2,推定先の品番から比抵抗仕様値を取得し、比抵抗推定ﾊﾟﾀｰﾝｺｰﾄﾞを取得する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                                   :説明
'          :sCryNum       ,I  ,String                               :結晶番号
'          :tSuiHin       ,I  ,tFullHinban                          :推定先品番(構造体)
'          :iSmplID1      ,I  ,Integer                              :推定元サンプルＩＤ１
'          :iSmplID2      ,I  ,Integer                              :推定元サンプルＩＤ２
'          :sHSXRSPOT     ,I  ,String                               :推定先RS測定点数
'          :tCryRs()      ,I  ,type_DBDRV_scmzc_fcmkc001c_CryR      :RS実績 (0)→推定元Top, (1)→推定元Bot, (2)→推定先
'          :iGetPCode1    ,O  ,String                               :推定元パターン１('A' or 'B')
'          :iGetPCode2    ,O  ,String                               :推定元パターン２('A' or 'B')
'          :戻り値        ,O  ,Integer                              :取得結果 = 0 : 正常終了
'                                                                               1 : 正常終了(該当サンプルなし)
'                                                                              -1 : 異常終了
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Private Function funGetPcodeRS(sCryNum As String, tSuiHin As tFullHinban, iSmplID1 As Integer, iSmplID2 As Integer, _
                                                    sHSXRSPOT As String, tCryRs() As type_DBDRV_scmzc_fcmkc001c_CryR, _
                                                    iGetPCode1 As String, iGetPCode2 As String) As Integer
    
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim getNewSpec  As String       '新ｻﾝﾌﾟﾙ位置比抵抗仕様値
    Dim wcnt        As Integer
    Dim getTopHin   As tFullHinban  'TOP位置品番
    Dim getTopSpec  As String       'TOP位置比抵抗仕様値
    Dim getTopPtrn  As String       'TOP位置ﾊﾟﾀｰﾝｺｰﾄﾞ
    Dim getBotHin   As tFullHinban  'BOT位置品番
    Dim getBotSpec  As String       'BOT位置比抵抗仕様値
    Dim getBotPtrn  As String       'BOT位置ﾊﾟﾀｰﾝｺｰﾄﾞ
    
    '-------------------- 推定先 --------------------
    '各品番の比抵抗仕様値取得
    '≪指定された新サンプル位置≫
    getNewSpec = funGetSuiSpecRS(tSuiHin)
    If getNewSpec = " " Then GoTo GetPcodeRSEmpty
    
    '-------------------- 推定元１ --------------------
    '指定された情報を元に、新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ)(XSDCS)を検索する。
    '≪推定元サンプルＩＤ１(TOP位置)の取得≫
    sql = "select HINBCS, REVNUMCS, FACTORYCS, OPECS from XSDCS "
'' 09/03/02 FAE)akiyama start
'    sql = sql & "where XTALCS = '" & sCryNum & "' and "
    sql = sql & "where CRYNUMCS LIKE '" & left(sCryNum, 9) & "%' and "
'' 09/03/02 FAE)akiyama end
    sql = sql & "      REPSMPLIDCS = " & iSmplID1
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        Set rs = Nothing
        GoTo GetPcodeRSEmpty
    End If
    
    'TOP位置データの設定
    getTopHin.hinban = rs("HINBCS")         'TOP位置品番
    getTopHin.mnorevno = rs("REVNUMCS")     'TOP位置製品番号改訂番号
    getTopHin.factory = rs("FACTORYCS")     'TOP位置工場
    getTopHin.opecond = rs("OPECS")         'TOP位置操業条件
    Set rs = Nothing
    
    '≪推定元サンプルＩＤ１(TOP位置)≫
    getTopSpec = funGetSuiSpecRS(getTopHin)
    If getTopSpec <> " " Then
        'コードDB取得関数を呼び出し､コードテーブルから比抵抗推定パターンコードを取得する｡
        getTopPtrn = "A"
    Else
        '実績ﾃﾞｰﾀから、件数を算出する
        wcnt = funGetRsCnt(tCryRs(0))
        If wcnt < 1 Then GoTo GetPcodeRSEmpty

        If wcnt = sHSXRSPOT Then
            getTopPtrn = "A"
        ElseIf wcnt > sHSXRSPOT Then
            getTopPtrn = "B"
        Else
            GoTo GetPcodeRSEmpty
        End If
    End If
    
    '-------------------- 推定元２ --------------------
    '指定された情報を元に、新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ)(XSDCS)を検索する。
    '≪推定元サンプルＩＤ２(BOT位置)の取得≫
    sql = "select HINBCS, REVNUMCS, FACTORYCS, OPECS from XSDCS "
'' 09/03/02 FAE)akiyama start
'    sql = sql & "where XTALCS = '" & sCryNum & "' and "
    sql = sql & "where CRYNUMCS LIKE '" & left(sCryNum, 9) & "%' and "
'' 09/03/02 FAE)akiyama end
    sql = sql & "      REPSMPLIDCS = " & iSmplID2
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        Set rs = Nothing
        GoTo GetPcodeRSEmpty
    End If
    
    'BOT位置データの設定
    getBotHin.hinban = rs("HINBCS")         'BOT位置品番
    getBotHin.mnorevno = rs("REVNUMCS")     'BOT位置製品番号改訂番号
    getBotHin.factory = rs("FACTORYCS")     'BOT位置工場
    getBotHin.opecond = rs("OPECS")         'BOT位置操業条件
    Set rs = Nothing
    
    '≪推定元サンプルＩＤ２(BOT位置)≫
    getBotSpec = funGetSuiSpecRS(getBotHin)
    If getBotSpec <> " " Then
        'コードDB取得関数を呼び出し､コードテーブルから比抵抗推定パターンコードを取得する｡
        getBotPtrn = "A"
    Else
        '実績ﾃﾞｰﾀから、件数を算出する
        wcnt = funGetRsCnt(tCryRs(1))
        If wcnt < 1 Then GoTo GetPcodeRSEmpty

        If wcnt = sHSXRSPOT Then
            getBotPtrn = "A"
        ElseIf wcnt > sHSXRSPOT Then
            getBotPtrn = "B"
        Else
            GoTo GetPcodeRSEmpty
        End If
    End If
    
    '呼び出し元への結果通知
    iGetPCode1 = getTopPtrn         '推定元パターン１('A' or 'B')
    iGetPCode2 = getBotPtrn         '推定元パターン２('A' or 'B')
    
    funGetPcodeRS = 0
    Exit Function

GetPcodeRSEmpty:
    funGetPcodeRS = 1
    Exit Function

GetPcodeRSParameterErr:
    funGetPcodeRS = -1
End Function

'------------------------------------------------
' 結晶抵抗実績のデータ件数取得
'------------------------------------------------
'概要      :結晶抵抗実績(構造体)に存在するデータ件数を取得する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :tCryRs        ,I  ,type_DBDRV_scmzc_fcmkc001c_CryR      :結晶抵抗実績構造体
'          :戻り値        ,O  ,Integer                              :データ件数
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Private Function funGetRsCnt(tCryRs As type_DBDRV_scmzc_fcmkc001c_CryR) As Integer
    
    Dim sql         As String
    Dim rs          As OraDynaset

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function funGetRsCnt"

    funGetRsCnt = 0
    
    If tCryRs.MEAS1 = -1 Then GoTo proc_exit
    funGetRsCnt = funGetRsCnt + 1
    If tCryRs.MEAS2 = -1 Then GoTo proc_exit
    funGetRsCnt = funGetRsCnt + 1
    If tCryRs.MEAS3 = -1 Then GoTo proc_exit
    funGetRsCnt = funGetRsCnt + 1
    If tCryRs.MEAS4 = -1 Then GoTo proc_exit
    funGetRsCnt = funGetRsCnt + 1
    If tCryRs.MEAS5 = -1 Then GoTo proc_exit
    funGetRsCnt = funGetRsCnt + 1

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    funGetRsCnt = -1
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' 結晶抵抗実績の処理回数取得
'------------------------------------------------
'概要      :結晶抵抗実績(TBCMJ002)から該当するデータの処理回数を取得する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :Samp          ,I  ,type_DBDRV_scmzc_fcmkc001c_CrySmp    :新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ)構造体
'          :戻り値        ,O  ,Integer                              :処理回数(最大値＋１)
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Private Function funGetTrancntRS(Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp) As Integer
    
    Dim sql         As String
    Dim rs          As OraDynaset

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function funGetTrancntRS"

    Set rs = Nothing

    ' 結晶抵抗実績テーブルから値を取得
    sql = "select TRANCNT+1 MAXCNT from TBCMJ002 "
    sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
    sql = sql & "      SMPLNO = " & Samp.REPSMPLIDCS & " and "
    sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ002 "
    sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
    sql = sql & "                 SMPLNO = " & Samp.REPSMPLIDCS & ")"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    If rs.EOF Or rs.RecordCount = 0 Then
        funGetTrancntRS = 1
    Else
        funGetTrancntRS = rs("MAXCNT")
    End If
    Set rs = Nothing

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    funGetTrancntRS = -1
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :内部関数 Oi実績取得用
Private Function Oi_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                            Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                            Oi As type_DBDRV_scmzc_fcmkc001c_Oi, _
                            TorB As Integer, _
                            Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean

    NothingFlag = False

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function Oi_Zisseki"

    Oi_Zisseki = FUNCTION_RETURN_SUCCESS

    ' 指示(仕様)と実績FLGを確認
    If (Samp.CRYINDOICS <> "0") And (Samp.CRYRESOICS <> "0") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, "
        sql = sql & "OIMEAS1, OIMEAS2, OIMEAS3, OIMEAS4, OIMEAS5, ORGRES, AVE, FTIRCONV, INSPECTWAY, REGDATE "
        sql = sql & "from TBCMJ003 "
        sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "      SMPLNO = " & Samp.CRYSMPLIDOICS & " and "
        sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ003 "
        sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "                       SMPLNO = " & Samp.CRYSMPLIDOICS & ")"
        sql = sql & "  and TRANCOND = 0 "       'GFAのFTIR換算値表示異常対応 2011/01/20追加 SETsw kubota
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With Oi
                .CRYNUM = rs("CRYNUM")          ' 結晶番号
                .POSITION = rs("POSITION")      ' 位置
                .SMPKBN = rs("SMPKBN")          ' サンプル区分
                .TRANCOND = rs("TRANCOND")      ' 処理条件
                .TRANCNT = rs("TRANCNT")        ' 処理回数
                .SMPLNO = rs("SMPLNO")          ' サンプルＮｏ
                .SMPLUMU = rs("SMPLUMU")        ' サンプル有無
'OI_NULL対応　2005/03/08 TUKU START --------------------------------------------------
                If IsNull(rs("OIMEAS1")) = False Then .OIMEAS1 = rs("OIMEAS1") Else .OIMEAS1 = -1  'Ｏｉ測定値1
                If IsNull(rs("OIMEAS2")) = False Then .OIMEAS2 = rs("OIMEAS2") Else .OIMEAS2 = -1  'Ｏｉ測定値2
                If IsNull(rs("OIMEAS3")) = False Then .OIMEAS3 = rs("OIMEAS3") Else .OIMEAS3 = -1  'Ｏｉ測定値3
                If IsNull(rs("OIMEAS4")) = False Then .OIMEAS4 = rs("OIMEAS4") Else .OIMEAS4 = -1  'Ｏｉ測定値4
                If IsNull(rs("OIMEAS5")) = False Then .OIMEAS5 = rs("OIMEAS5") Else .OIMEAS5 = -1  'Ｏｉ測定値5
                If IsNull(rs("ORGRES")) = False Then .ORGRES = rs("ORGRES") Else .ORGRES = -1    ' ＯＲＧ結果
'OI_NULL対応　2005/03/08 TUKU END   --------------------------------------------------
                .AVE = rs("AVE")                ' ＡＶＥ
                .FTIRCONV = rs("FTIRCONV")      ' ＦＴＩＲ換算
                .INSPECTWAY = rs("INSPECTWAY")  ' 検査方法
                .REGDATE = rs("REGDATE")        ' 登録日付
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Oi_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :内部関数 BMD実績取得用
Private Function BMD_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                             Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                             inTRANCOND As Integer, _
                             BMD As type_DBDRV_scmzc_fcmkc001c_BMD, _
                             TorB As Integer, _
                             Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean
    Dim wHSX_HS     As String
    Dim wCryIND     As String
    Dim wCryRES     As String
    Dim wCrySMPL    As Long         'Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    
    NothingFlag = False

    ' BMD実績テーブルから値を取得

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function BMD_Zisseki"

    BMD_Zisseki = FUNCTION_RETURN_SUCCESS

    If inTRANCOND = 1 Then
        wHSX_HS = siyou.HSXBM1HS
        wCryIND = Samp.CRYINDB1CS
        wCryRES = Samp.CRYRESB1CS
        wCrySMPL = Samp.CRYSMPLIDB1CS
    ElseIf inTRANCOND = 2 Then
        wHSX_HS = siyou.HSXBM2HS
        wCryIND = Samp.CRYINDB2CS
        wCryRES = Samp.CRYRESB2CS
        wCrySMPL = Samp.CRYSMPLIDB2CS
    ElseIf inTRANCOND = 3 Then
        wHSX_HS = siyou.HSXBM3HS
        wCryIND = Samp.CRYINDB3CS
        wCryRES = Samp.CRYRESB3CS
        wCrySMPL = Samp.CRYSMPLIDB3CS
    End If

    ' 指示(仕様)と実績FLGを確認
    If (wCryIND <> "0") And (wCryRES <> "0") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HTPRC, KKSP, KKSET, "
        sql = sql & "MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASMIN, MEASMAX, MEASAVE, BMDMNBUNP, REGDATE "
        sql = sql & "from TBCMJ008 "
        sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "      SMPLNO = " & wCrySMPL & " and "
        sql = sql & "      TRANCOND = '" & inTRANCOND & "' and "
        sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ008 "
        sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "                       SMPLNO = " & wCrySMPL & " and "
        sql = sql & "                       TRANCOND = '" & inTRANCOND & "')"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With BMD
                .CRYNUM = rs("CRYNUM")          ' 結晶番号
                .POSITION = rs("POSITION")      ' 位置
                .SMPKBN = rs("SMPKBN")          ' サンプル区分
                .TRANCOND = rs("TRANCOND")      ' 処理条件
                .TRANCNT = rs("TRANCNT")        ' 処理回数
                .SMPLNO = rs("SMPLNO")          ' サンプルＮｏ
                .SMPLUMU = rs("SMPLUMU")        ' サンプル有無
                .HTPRC = rs("HTPRC")            ' 熱処理方法
                .KKSP = rs("KKSP")              ' 結晶欠陥測定位置
                .KKSET = rs("KKSET")            ' 結晶欠陥測定条件＋選択ET代
                .MEAS1 = rs("MEAS1")            ' 測定値１
                .MEAS2 = rs("MEAS2")            ' 測定値２
                .MEAS3 = rs("MEAS3")            ' 測定値３
                .MEAS4 = rs("MEAS4")            ' 測定値４
                .MEAS5 = rs("MEAS5")            ' 測定値５
                .MEASMIN = rs("MEASMIN")        ' MIN
                .MEASMAX = rs("MEASMAX")        ' MAX
                .MEASAVE = rs("MEASAVE")        ' AVE
                 If IsNull(rs("BMDMNBUNP")) = False Then .BMDMNBUNP = rs("BMDMNBUNP")       ' BMD面内分布
                .REGDATE = rs("REGDATE")        ' 登録日付
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    BMD_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :内部関数 GD実績取得用
Private Function OSF_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                             Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                             inTRANCOND As Integer, _
                             OSF As type_DBDRV_scmzc_fcmkc001c_OSF, _
                             TorB As Integer, _
                             Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean
    Dim wHSX_HS     As String
    Dim wCryIND     As String
    Dim wCryRES     As String
    Dim wCrySMPL    As Long     'Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota

    NothingFlag = False

    ' OSF実績テーブルから値を取得

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function OSF_Zisseki"

    OSF_Zisseki = FUNCTION_RETURN_SUCCESS

    If inTRANCOND = 1 Then
        wHSX_HS = siyou.HSXOF1HS
        wCryIND = Samp.CRYINDL1CS
        wCryRES = Samp.CRYRESL1CS
        wCrySMPL = Samp.CRYSMPLIDL1CS
    ElseIf inTRANCOND = 2 Then
        wHSX_HS = siyou.HSXOF2HS
        wCryIND = Samp.CRYINDL2CS
        wCryRES = Samp.CRYRESL2CS
        wCrySMPL = Samp.CRYSMPLIDL2CS
    ElseIf inTRANCOND = 3 Then
        wHSX_HS = siyou.HSXOF3HS
        wCryIND = Samp.CRYINDL3CS
        wCryRES = Samp.CRYRESL3CS
        wCrySMPL = Samp.CRYSMPLIDL3CS
    ElseIf inTRANCOND = 4 Then
        wHSX_HS = siyou.HSXOF4HS
        wCryIND = Samp.CRYINDL4CS
        wCryRES = Samp.CRYRESL4CS
        wCrySMPL = Samp.CRYSMPLIDL4CS
    End If

    ' 指示(仕様)と実績FLGを確認
    If (wCryIND <> "0") And (wCryRES <> "0") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HTPRC, KKSP, KKSET, CALCMAX, CALCAVE, "
        sql = sql & "MEAS1, MEAS2,  MEAS3,  MEAS4,  MEAS5,  MEAS6,  MEAS7,  MEAS8,  MEAS9,  MEAS10, "
        sql = sql & "MEAS11,MEAS12, MEAS13, MEAS14, MEAS15, MEAS16, MEAS17, MEAS18, MEAS19, MEAS20, "
        sql = sql & "OSFPOS1, OSFWID1, OSFRD1, OSFPOS2, OSFWID2, OSFRD2, OSFPOS3, OSFWID3, OSFRD3, REGDATE "
        
        sql = sql & ",CALCMH "  '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
        'Add Start 2012/06/01 SMPK H.Ohkubo
        sql = sql & ",COSF3PTNJSK "
        'Add End 2012/06/01 SMPK H.Ohkubo
        sql = sql & "from TBCMJ005 "
        sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "      SMPLNO = " & wCrySMPL & " and "
        sql = sql & "      TRANCOND = '" & inTRANCOND & "' and "
        sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ005 "
        sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "                       SMPLNO = " & wCrySMPL & " and "
        sql = sql & "                       TRANCOND = '" & inTRANCOND & "')"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With OSF
                .CRYNUM = rs("CRYNUM")          ' 結晶番号
                .POSITION = rs("POSITION")      ' 位置
                .SMPKBN = rs("SMPKBN")          ' サンプル区分
                .TRANCOND = rs("TRANCOND")      ' 処理条件
                .TRANCNT = rs("TRANCNT")        ' 処理回数
                .SMPLNO = rs("SMPLNO")          ' サンプルＮｏ
                .SMPLUMU = rs("SMPLUMU")        ' サンプル有無
                .HTPRC = rs("HTPRC")            ' 熱処理方法
                .KKSP = rs("KKSP")              ' 結晶欠陥測定位置
                .KKSET = rs("KKSET")            ' 結晶欠陥測定条件＋選択ET代
                .CALCMAX = rs("CALCMAX")       ' 計算結果 Max
                .CALCAVE = rs("CALCAVE")       ' 計算結果 Ave
                .MEAS1 = rs("MEAS1")           ' 測定値１
                .MEAS2 = rs("MEAS2")           ' 測定値２
                .MEAS3 = rs("MEAS3")           ' 測定値３
                .MEAS4 = rs("MEAS4")           ' 測定値４
                .MEAS5 = rs("MEAS5")           ' 測定値５
                .MEAS6 = rs("MEAS6")           ' 測定値６
                .MEAS7 = rs("MEAS7")           ' 測定値７
                .MEAS8 = rs("MEAS8")           ' 測定値８
                .MEAS9 = rs("MEAS9")           ' 測定値９
                .MEAS10 = rs("MEAS10")         ' 測定値１０
                .MEAS11 = rs("MEAS11")         ' 測定値１１
                .MEAS12 = rs("MEAS12")         ' 測定値１２
                .MEAS13 = rs("MEAS13")         ' 測定値１３
                .MEAS14 = rs("MEAS14")         ' 測定値１４
                .MEAS15 = rs("MEAS15")         ' 測定値１５
                .MEAS16 = rs("MEAS16")         ' 測定値１６
                .MEAS17 = rs("MEAS17")         ' 測定値１７
                .MEAS18 = rs("MEAS18")         ' 測定値１８
                .MEAS19 = rs("MEAS19")         ' 測定値１９
                .MEAS20 = rs("MEAS20")         ' 測定値２０
                 If IsNull(rs("OSFPOS1")) = False Then .OSFPOS1 = rs("OSFPOS1")   'ﾊﾟﾀｰﾝ区分１位置
                 If IsNull(rs("OSFWID1")) = False Then .OSFWID1 = rs("OSFWID1")   'ﾊﾟﾀｰﾝ区分１幅
                 If IsNull(rs("OSFRD1")) = False Then .OSFRD1 = rs("OSFRD1")      'ﾊﾟﾀｰﾝ区分１R/D
                 If IsNull(rs("OSFPOS2")) = False Then .OSFPOS2 = rs("OSFPOS2")   'ﾊﾟﾀｰﾝ区分２位置
                 If IsNull(rs("OSFWID2")) = False Then .OSFWID2 = rs("OSFWID2")   'ﾊﾟﾀｰﾝ区分２幅
                 If IsNull(rs("OSFRD2")) = False Then .OSFRD2 = rs("OSFRD2")      'ﾊﾟﾀｰﾝ区分２R/D
                 If IsNull(rs("OSFPOS3")) = False Then .OSFPOS3 = rs("OSFPOS3")   'ﾊﾟﾀｰﾝ区分３位置
                 If IsNull(rs("OSFWID3")) = False Then .OSFWID3 = rs("OSFWID3")   'ﾊﾟﾀｰﾝ区分３幅
                 If IsNull(rs("OSFRD3")) = False Then .OSFRD3 = rs("OSFRD3")      'ﾊﾟﾀｰﾝ区分３R/D
                 If IsNull(rs("CALCMH")) = False Then .CALCMH = rs("CALCMH")      '面内比(MAX/MIN)  '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
                .REGDATE = rs("REGDATE")       ' 登録日付
                
                'Add Start 2012/06/01 SMPK H.Ohkubo
                If Not IsNull(rs("COSF3PTNJSK")) Then
                    .COSF3PTNJSK = rs("COSF3PTNJSK")   ' パターン区分実績
                Else
                    'パターン実績なし
                    .COSF3PTNJSK = "0"
                End If
                'Add End 2012/06/01 SMPK H.Ohkubo
                
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    OSF_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'内部関数 Cs実績取得用
Private Function CS_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                            Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                            Cs As type_DBDRV_scmzc_fcmkc001c_CS, _
                            TorB As Integer, _
                            Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean

    NothingFlag = False

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function CS_Zisseki"
    
    CS_Zisseki = FUNCTION_RETURN_SUCCESS

    ' 指示(仕様)と実績FLGを確認
    If (Samp.CRYINDCSCS <> "0") And (Samp.CRYRESCSCS <> "0") Then

        sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, "
        sql = sql & "CSMEAS, PRE70P, INSPECTWAY, REGDATE "
        sql = sql & "from TBCMJ004 "
        sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "      SMPLNO = " & Samp.CRYSMPLIDCSCS & " and "
        sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ004 "
        sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "                       SMPLNO = " & Samp.CRYSMPLIDCSCS & ")"

        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With Cs
                .CRYNUM = rs("CRYNUM")          ' 結晶番号
                .POSITION = rs("POSITION")      ' 位置
                .SMPKBN = rs("SMPKBN")          ' サンプル区分
                .TRANCOND = rs("TRANCOND")      ' 処理条件
                .TRANCNT = rs("TRANCNT")        ' 処理回数
                .SMPLNO = rs("SMPLNO")          ' サンプルＮｏ
                .SMPLUMU = rs("SMPLUMU")        ' サンプル有無
'OI_NULL対応　2005/03/08 TUKU START --------------------------------------------------
                If IsNull(rs("CSMEAS")) = False Then .CSMEAS = rs("CSMEAS") Else .CSMEAS = -1  ' Cs実測値
                If IsNull(rs("PRE70P")) = False Then .PRE70P = rs("PRE70P") Else .PRE70P = -1  ' ７０％推定値
'OI_NULL対応　2005/03/08 TUKU START --------------------------------------------------
                .INSPECTWAY = rs("INSPECTWAY")  ' 検査方法
                .REGDATE = rs("REGDATE")        ' 登録日付
            End With
        Else
            NothingFlag = True
        End If

        Set rs = Nothing
    End If
    
    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    gErr.HandleError
    Resume proc_exit
End Function

'内部関数 GD実績取得用
Private Function GD_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                            Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                            GD As type_DBDRV_scmzc_fcmkc001c_GD, _
                            TorB As Integer, _
                            Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean
    
    NothingFlag = False

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function GD_Zisseki"

    GD_Zisseki = FUNCTION_RETURN_SUCCESS

    ' 指示(仕様)と実績FLGを確認
    If (Samp.CRYINDGDCS <> "0") And (Samp.CRYRESGDCS <> "0") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, MSRSDEN, MSRSLDL, MSRSDVD2, "
        sql = sql & "MS01LDL1, MS01LDL2, MS01LDL3, MS01LDL4, MS01LDL5, MS01DEN1, MS01DEN2, MS01DEN3, MS01DEN4, MS01DEN5, "
        sql = sql & "MS02LDL1, MS02LDL2, MS02LDL3, MS02LDL4, MS02LDL5, MS02DEN1, MS02DEN2, MS02DEN3, MS02DEN4, MS02DEN5, "
        sql = sql & "MS03LDL1, MS03LDL2, MS03LDL3, MS03LDL4, MS03LDL5, MS03DEN1, MS03DEN2, MS03DEN3, MS03DEN4, MS03DEN5, "
        sql = sql & "MS04LDL1, MS04LDL2, MS04LDL3, MS04LDL4, MS04LDL5, MS04DEN1, MS04DEN2, MS04DEN3, MS04DEN4, MS04DEN5, "
        sql = sql & "MS05LDL1, MS05LDL2, MS05LDL3, MS05LDL4, MS05LDL5, MS05DEN1, MS05DEN2, MS05DEN3, MS05DEN4, MS05DEN5, "
        sql = sql & "MS06LDL1, MS06LDL2, MS06LDL3, MS06LDL4, MS06LDL5, MS06DEN1, MS06DEN2, MS06DEN3, MS06DEN4, MS06DEN5, "
        sql = sql & "MS07LDL1, MS07LDL2, MS07LDL3, MS07LDL4, MS07LDL5, MS07DEN1, MS07DEN2, MS07DEN3, MS07DEN4, MS07DEN5, "
        sql = sql & "MS08LDL1, MS08LDL2, MS08LDL3, MS08LDL4, MS08LDL5, MS08DEN1, MS08DEN2, MS08DEN3, MS08DEN4, MS08DEN5, "
        sql = sql & "MS09LDL1, MS09LDL2, MS09LDL3, MS09LDL4, MS09LDL5, MS09DEN1, MS09DEN2, MS09DEN3, MS09DEN4, MS09DEN5, "
        sql = sql & "MS10LDL1, MS10LDL2, MS10LDL3, MS10LDL4, MS10LDL5, MS10DEN1, MS10DEN2, MS10DEN3, MS10DEN4, MS10DEN5, "
        sql = sql & "MS11LDL1, MS11LDL2, MS11LDL3, MS11LDL4, MS11LDL5, MS11DEN1, MS11DEN2, MS11DEN3, MS11DEN4, MS11DEN5, "
        sql = sql & "MS12LDL1, MS12LDL2, MS12LDL3, MS12LDL4, MS12LDL5, MS12DEN1, MS12DEN2, MS12DEN3, MS12DEN4, MS12DEN5, "
        sql = sql & "MS13LDL1, MS13LDL2, MS13LDL3, MS13LDL4, MS13LDL5, MS13DEN1, MS13DEN2, MS13DEN3, MS13DEN4, MS13DEN5, "
        sql = sql & "MS14LDL1, MS14LDL2, MS14LDL3, MS14LDL4, MS14LDL5, MS14DEN1, MS14DEN2, MS14DEN3, MS14DEN4, MS14DEN5, "
        sql = sql & "MS15LDL1, MS15LDL2, MS15LDL3, MS15LDL4, MS15LDL5, MS15DEN1, MS15DEN2, MS15DEN3, MS15DEN4, MS15DEN5, "
        sql = sql & "MS01DVD2, MS02DVD2, MS03DVD2, MS04DVD2, MS05DVD2, REGDATE "
        
        sql = sql & ",MSZEROMN, MSZEROMX "  '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
        
        sql = sql & "from TBCMJ006 "
        sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "      SMPLNO = " & Samp.CRYSMPLIDGDCS & " and "
        sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ006 "
        sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "                       SMPLNO = " & Samp.CRYSMPLIDGDCS & ")"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With GD
                .CRYNUM = rs("CRYNUM")          ' 結晶番号
                .POSITION = rs("POSITION")      ' 位置
                .SMPKBN = rs("SMPKBN")          ' サンプル区分
                .TRANCOND = rs("TRANCOND")      ' 処理条件
                .TRANCNT = rs("TRANCNT")        ' 処理回数
                .SMPLNO = rs("SMPLNO")          ' サンプルＮｏ
                .SMPLUMU = rs("SMPLUMU")        ' サンプル有無
                .MSRSDEN = rs("MSRSDEN")        ' 測定結果 Den
                .MSRSLDL = rs("MSRSLDL")        ' 測定結果 L/DL
                .MSRSDVD2 = rs("MSRSDVD2")      ' 測定結果 DVD2
                .MS01LDL1 = rs("MS01LDL1")      ' 測定値01 L/DL1
                .MS01LDL2 = rs("MS01LDL2")      ' 測定値01 L/DL2
                .MS01LDL3 = rs("MS01LDL3")      ' 測定値01 L/DL3
                .MS01LDL4 = rs("MS01LDL4")      ' 測定値01 L/DL4
                .MS01LDL5 = rs("MS01LDL5")      ' 測定値01 L/DL5
                .MS01DEN1 = rs("MS01DEN1")      ' 測定値01 Den1
                .MS01DEN2 = rs("MS01DEN2")      ' 測定値01 Den2
                .MS01DEN3 = rs("MS01DEN3")      ' 測定値01 Den3
                .MS01DEN4 = rs("MS01DEN4")      ' 測定値01 Den4
                .MS01DEN5 = rs("MS01DEN5")      ' 測定値01 Den5
                .MS02LDL1 = rs("MS02LDL1")      ' 測定値02 L/DL1
                .MS02LDL2 = rs("MS02LDL2")      ' 測定値02 L/DL2
                .MS02LDL3 = rs("MS02LDL3")      ' 測定値02 L/DL3
                .MS02LDL4 = rs("MS02LDL4")      ' 測定値02 L/DL4
                .MS02LDL5 = rs("MS02LDL5")      ' 測定値02 L/DL5
                .MS02DEN1 = rs("MS02DEN1")      ' 測定値02 Den1
                .MS02DEN2 = rs("MS02DEN2")      ' 測定値02 Den2
                .MS02DEN3 = rs("MS02DEN3")      ' 測定値02 Den3
                .MS02DEN4 = rs("MS02DEN4")      ' 測定値02 Den4
                .MS02DEN5 = rs("MS02DEN5")      ' 測定値02 Den5
                .MS03LDL1 = rs("MS03LDL1")      ' 測定値03 L/DL1
                .MS03LDL2 = rs("MS03LDL2")      ' 測定値03 L/DL2
                .MS03LDL3 = rs("MS03LDL3")      ' 測定値03 L/DL3
                .MS03LDL4 = rs("MS03LDL4")      ' 測定値03 L/DL4
                .MS03LDL5 = rs("MS03LDL5")      ' 測定値03 L/DL5
                .MS03DEN1 = rs("MS03DEN1")      ' 測定値03 Den1
                .MS03DEN2 = rs("MS03DEN2")      ' 測定値03 Den2
                .MS03DEN3 = rs("MS03DEN3")      ' 測定値03 Den3
                .MS03DEN4 = rs("MS03DEN4")      ' 測定値03 Den4
                .MS03DEN5 = rs("MS03DEN5")      ' 測定値03 Den5
                .MS04LDL1 = rs("MS04LDL1")      ' 測定値04 L/DL1
                .MS04LDL2 = rs("MS04LDL2")      ' 測定値04 L/DL2
                .MS04LDL3 = rs("MS04LDL3")      ' 測定値04 L/DL3
                .MS04LDL4 = rs("MS04LDL4")      ' 測定値04 L/DL4
                .MS04LDL5 = rs("MS04LDL5")      ' 測定値04 L/DL5
                .MS04DEN1 = rs("MS04DEN1")      ' 測定値04 Den1
                .MS04DEN2 = rs("MS04DEN2")      ' 測定値04 Den2
                .MS04DEN3 = rs("MS04DEN3")      ' 測定値04 Den3
                .MS04DEN4 = rs("MS04DEN4")      ' 測定値04 Den4
                .MS04DEN5 = rs("MS04DEN5")      ' 測定値04 Den5
                .MS05LDL1 = rs("MS05LDL1")      ' 測定値05 L/DL1
                .MS05LDL2 = rs("MS05LDL2")      ' 測定値05 L/DL2
                .MS05LDL3 = rs("MS05LDL3")      ' 測定値05 L/DL3
                .MS05LDL4 = rs("MS05LDL4")      ' 測定値05 L/DL4
                .MS05LDL5 = rs("MS05LDL5")      ' 測定値05 L/DL5
                .MS05DEN1 = rs("MS05DEN1")      ' 測定値05 Den1
                .MS05DEN2 = rs("MS05DEN2")      ' 測定値05 Den2
                .MS05DEN3 = rs("MS05DEN3")      ' 測定値05 Den3
                .MS05DEN4 = rs("MS05DEN4")      ' 測定値05 Den4
                .MS05DEN5 = rs("MS05DEN5")      ' 測定値05 Den5
                .MS06LDL1 = rs("MS06LDL1")      ' 測定値06 L/DL1
                .MS06LDL2 = rs("MS06LDL2")      ' 測定値06 L/DL2
                .MS06LDL3 = rs("MS06LDL3")      ' 測定値06 L/DL3
                .MS06LDL4 = rs("MS06LDL4")      ' 測定値06 L/DL4
                .MS06LDL5 = rs("MS06LDL5")      ' 測定値06 L/DL5
                .MS06DEN1 = rs("MS06DEN1")      ' 測定値06 Den1
                .MS06DEN2 = rs("MS06DEN2")      ' 測定値06 Den2
                .MS06DEN3 = rs("MS06DEN3")      ' 測定値06 Den3
                .MS06DEN4 = rs("MS06DEN4")      ' 測定値06 Den4
                .MS06DEN5 = rs("MS06DEN5")      ' 測定値06 Den5
                .MS07LDL1 = rs("MS07LDL1")      ' 測定値07 L/DL1
                .MS07LDL2 = rs("MS07LDL2")      ' 測定値07 L/DL2
                .MS07LDL3 = rs("MS07LDL3")      ' 測定値07 L/DL3
                .MS07LDL4 = rs("MS07LDL4")      ' 測定値07 L/DL4
                .MS07LDL5 = rs("MS07LDL5")      ' 測定値07 L/DL5
                .MS07DEN1 = rs("MS07DEN1")      ' 測定値07 Den1
                .MS07DEN2 = rs("MS07DEN2")      ' 測定値07 Den2
                .MS07DEN3 = rs("MS07DEN3")      ' 測定値07 Den3
                .MS07DEN4 = rs("MS07DEN4")      ' 測定値07 Den4
                .MS07DEN5 = rs("MS07DEN5")      ' 測定値07 Den5
                .MS08LDL1 = rs("MS08LDL1")      ' 測定値08 L/DL1
                .MS08LDL2 = rs("MS08LDL2")      ' 測定値08 L/DL2
                .MS08LDL3 = rs("MS08LDL3")      ' 測定値08 L/DL3
                .MS08LDL4 = rs("MS08LDL4")      ' 測定値08 L/DL4
                .MS08LDL5 = rs("MS08LDL5")      ' 測定値08 L/DL5
                .MS08DEN1 = rs("MS08DEN1")      ' 測定値08 Den1
                .MS08DEN2 = rs("MS08DEN2")      ' 測定値08 Den2
                .MS08DEN3 = rs("MS08DEN3")      ' 測定値08 Den3
                .MS08DEN4 = rs("MS08DEN4")      ' 測定値08 Den4
                .MS08DEN5 = rs("MS08DEN5")      ' 測定値08 Den5
                .MS09LDL1 = rs("MS09LDL1")      ' 測定値09 L/DL1
                .MS09LDL2 = rs("MS09LDL2")      ' 測定値09 L/DL2
                .MS09LDL3 = rs("MS09LDL3")      ' 測定値09 L/DL3
                .MS09LDL4 = rs("MS09LDL4")      ' 測定値09 L/DL4
                .MS09LDL5 = rs("MS09LDL5")      ' 測定値09 L/DL5
                .MS09DEN1 = rs("MS09DEN1")      ' 測定値09 Den1
                .MS09DEN2 = rs("MS09DEN2")      ' 測定値09 Den2
                .MS09DEN3 = rs("MS09DEN3")      ' 測定値09 Den3
                .MS09DEN4 = rs("MS09DEN4")      ' 測定値09 Den4
                .MS09DEN5 = rs("MS09DEN5")      ' 測定値09 Den5
                .MS10LDL1 = rs("MS10LDL1")      ' 測定値10 L/DL1
                .MS10LDL2 = rs("MS10LDL2")      ' 測定値10 L/DL2
                .MS10LDL3 = rs("MS10LDL3")      ' 測定値10 L/DL3
                .MS10LDL4 = rs("MS10LDL4")      ' 測定値10 L/DL4
                .MS10LDL5 = rs("MS10LDL5")      ' 測定値10 L/DL5
                .MS10DEN1 = rs("MS10DEN1")      ' 測定値10 Den1
                .MS10DEN2 = rs("MS10DEN2")      ' 測定値10 Den2
                .MS10DEN3 = rs("MS10DEN3")      ' 測定値10 Den3
                .MS10DEN4 = rs("MS10DEN4")      ' 測定値10 Den4
                .MS10DEN5 = rs("MS10DEN5")      ' 測定値10 Den5
                .MS11LDL1 = rs("MS11LDL1")      ' 測定値11 L/DL1
                .MS11LDL2 = rs("MS11LDL2")      ' 測定値11 L/DL2
                .MS11LDL3 = rs("MS11LDL3")      ' 測定値11 L/DL3
                .MS11LDL4 = rs("MS11LDL4")      ' 測定値11 L/DL4
                .MS11LDL5 = rs("MS11LDL5")      ' 測定値11 L/DL5
                .MS11DEN1 = rs("MS11DEN1")      ' 測定値11 Den1
                .MS11DEN2 = rs("MS11DEN2")      ' 測定値11 Den2
                .MS11DEN3 = rs("MS11DEN3")      ' 測定値11 Den3
                .MS11DEN4 = rs("MS11DEN4")      ' 測定値11 Den4
                .MS11DEN5 = rs("MS11DEN5")      ' 測定値11 Den5
                .MS12LDL1 = rs("MS12LDL1")      ' 測定値12 L/DL1
                .MS12LDL2 = rs("MS12LDL2")      ' 測定値12 L/DL2
                .MS12LDL3 = rs("MS12LDL3")      ' 測定値12 L/DL3
                .MS12LDL4 = rs("MS12LDL4")      ' 測定値12 L/DL4
                .MS12LDL5 = rs("MS12LDL5")      ' 測定値12 L/DL5
                .MS12DEN1 = rs("MS12DEN1")      ' 測定値12 Den1
                .MS12DEN2 = rs("MS12DEN2")      ' 測定値12 Den2
                .MS12DEN3 = rs("MS12DEN3")      ' 測定値12 Den3
                .MS12DEN4 = rs("MS12DEN4")      ' 測定値12 Den4
                .MS12DEN5 = rs("MS12DEN5")      ' 測定値12 Den5
                .MS13LDL1 = rs("MS13LDL1")      ' 測定値13 L/DL1
                .MS13LDL2 = rs("MS13LDL2")      ' 測定値13 L/DL2
                .MS13LDL3 = rs("MS13LDL3")      ' 測定値13 L/DL3
                .MS13LDL4 = rs("MS13LDL4")      ' 測定値13 L/DL4
                .MS13LDL5 = rs("MS13LDL5")      ' 測定値13 L/DL5
                .MS13DEN1 = rs("MS13DEN1")      ' 測定値13 Den1
                .MS13DEN2 = rs("MS13DEN2")      ' 測定値13 Den2
                .MS13DEN3 = rs("MS13DEN3")      ' 測定値13 Den3
                .MS13DEN4 = rs("MS13DEN4")      ' 測定値13 Den4
                .MS13DEN5 = rs("MS13DEN5")      ' 測定値13 Den5
                .MS14LDL1 = rs("MS14LDL1")      ' 測定値14 L/DL1
                .MS14LDL2 = rs("MS14LDL2")      ' 測定値14 L/DL2
                .MS14LDL3 = rs("MS14LDL3")      ' 測定値14 L/DL3
                .MS14LDL4 = rs("MS14LDL4")      ' 測定値14 L/DL4
                .MS14LDL5 = rs("MS14LDL5")      ' 測定値14 L/DL5
                .MS14DEN1 = rs("MS14DEN1")      ' 測定値14 Den1
                .MS14DEN2 = rs("MS14DEN2")      ' 測定値14 Den2
                .MS14DEN3 = rs("MS14DEN3")      ' 測定値14 Den3
                .MS14DEN4 = rs("MS14DEN4")      ' 測定値14 Den4
                .MS14DEN5 = rs("MS14DEN5")      ' 測定値14 Den5
                .MS15LDL1 = rs("MS15LDL1")      ' 測定値15 L/DL1
                .MS15LDL2 = rs("MS15LDL2")      ' 測定値15 L/DL2
                .MS15LDL3 = rs("MS15LDL3")      ' 測定値15 L/DL3
                .MS15LDL4 = rs("MS15LDL4")      ' 測定値15 L/DL4
                .MS15LDL5 = rs("MS15LDL5")      ' 測定値15 L/DL5
                .MS15DEN1 = rs("MS15DEN1")      ' 測定値15 Den1
                .MS15DEN2 = rs("MS15DEN2")      ' 測定値15 Den2
                .MS15DEN3 = rs("MS15DEN3")      ' 測定値15 Den3
                .MS15DEN4 = rs("MS15DEN4")      ' 測定値15 Den4
                .MS15DEN5 = rs("MS15DEN5")      ' 測定値15 Den5
                If IsNull(rs("MS01DVD2")) = False Then .MS01DVD2 = rs("MS01DVD2")   '測定値01 DVD2
                If IsNull(rs("MS02DVD2")) = False Then .MS02DVD2 = rs("MS02DVD2")   '測定値02 DVD2
                If IsNull(rs("MS03DVD2")) = False Then .MS03DVD2 = rs("MS03DVD2")   '測定値03 DVD2
                If IsNull(rs("MS04DVD2")) = False Then .MS04DVD2 = rs("MS04DVD2")   '測定値04 DVD2
                If IsNull(rs("MS05DVD2")) = False Then .MS05DVD2 = rs("MS05DVD2")   '測定値05 DVD2
                
                If IsNull(rs("MSZEROMN")) = False Then .MSZEROMN = rs("MSZEROMN")   'L/DL0連続数最小値  '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
                If IsNull(rs("MSZEROMX")) = False Then .MSZEROMX = rs("MSZEROMX")   'L/DL0連続数最大値  '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
                
                .REGDATE = rs("REGDATE")        ' 登録日付
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    GD_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'内部関数 ライフタイム実績取得用
Private Function LT_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                            Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                            Lt As type_DBDRV_scmzc_fcmkc001c_LT, _
                            TorB As Integer, _
                            Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean
    
    NothingFlag = False

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function LT_Zisseki"

    ' ライフタイム実績テーブルから値を取得
    LT_Zisseki = FUNCTION_RETURN_SUCCESS

    ' 指示(仕様)と実績FLGを確認
    If (Samp.CRYINDTCS <> "0") And (Samp.CRYRESTCS <> "0") Then
        
        '2005/12/02 mod SET高崎 測定値１～５カラムNULL許可につきNVL使用 ->
        '                    測定値６～１０カラム追加
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, MEASPEAK, CALCMEAS, REGDATE, "
        sql = sql & "NVL(MEAS1, -1) MEAS1, "
        sql = sql & "NVL(MEAS2, -1) MEAS2, "
        sql = sql & "NVL(MEAS3, -1) MEAS3, "
        sql = sql & "NVL(MEAS4, -1) MEAS4, "
        sql = sql & "NVL(MEAS5, -1) MEAS5, "
        sql = sql & " NVL(MEAS6,-1) MEAS6, "
        sql = sql & " NVL(MEAS7,-1) MEAS7, "
        sql = sql & " NVL(MEAS8,-1) MEAS8, "
        sql = sql & " NVL(MEAS9,-1) MEAS9, "
        sql = sql & " NVL(MEAS10,-1) MEAS10, "
        sql = sql & " LTSPIFLG "
        sql = sql & ",NVL(CONVAL,-1) CONVAL "
        sql = sql & "from TBCMJ007 "
        sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "      SMPLNO = " & Samp.CRYSMPLIDTCS & " and "
        sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ007 "
        sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "                       SMPLNO = " & Samp.CRYSMPLIDTCS & ")"
        
        '2005/12/02 mod SET高崎 測定値１～５カラムNULL許可につきNVL使用
        '                    測定値６～１０カラム追加               <-
        Set rs = OraDB.CreateDynaset(sql, ORADYN_READONLY)
        If rs.RecordCount > 0 Then
            With Lt
                .CRYNUM = rs("CRYNUM")          ' 結晶番号
                .POSITION = rs("POSITION")      ' 位置
                .SMPKBN = rs("SMPKBN")          ' サンプル区分
                .TRANCOND = rs("TRANCOND")      ' 処理条件
                .TRANCNT = rs("TRANCNT")        ' 処理回数
                .SMPLNO = rs("SMPLNO")          ' サンプルＮｏ
                .SMPLUMU = rs("SMPLUMU")        ' サンプル有無
                .MEAS1 = rs("MEAS1")            ' 測定値１
                .MEAS2 = rs("MEAS2")            ' 測定値２
                .MEAS3 = rs("MEAS3")            ' 測定値３
                .MEAS4 = rs("MEAS4")            ' 測定値４
                .MEAS5 = rs("MEAS5")            ' 測定値５
                .MEASPEAK = rs("MEASPEAK")      ' 測定値 ピーク値
                .CALCMEAS = rs("CALCMEAS")      ' 計算結果
                .REGDATE = rs("REGDATE")        ' 登録日付
''Add Start 2011/07/22 LT10Ω判定追加対応 T.Koi(SETsw)
                .CONVAL = rs("CONVAL")          ' 10Ω換算値
''Add End   2011/07/22 LT10Ω判定追加対応 T.Koi(SETsw)
                '2005/12/02 add SET高崎 測定値６～１０カラム追加のため追加 ->
                .MEAS6 = rs("MEAS6")            ' 測定値６
                .MEAS7 = rs("MEAS7")            ' 測定値７
                .MEAS8 = rs("MEAS8")            ' 測定値８
                .MEAS9 = rs("MEAS9")            ' 測定値９
                .MEAS10 = rs("MEAS10")          ' 測定値１０
                .LTSPIFLG = Trim(CStr(NulltoStr(rs.Fields("LTSPIFLG").Value)))  '測定位置判定フラグ
                '2005/12/02 add SET高崎 測定値６～１０カラム追加のため追加 <-
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    LT_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :内部関数 EPD実績取得用
Private Function EPD_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                             Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                             EPD As type_DBDRV_scmzc_fcmkc001c_EPD, _
                             TorB As Integer, _
                             Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean

    NothingFlag = False

    ' EPD実績テーブルから値を取得

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function EPD_Zisseki"

    EPD_Zisseki = FUNCTION_RETURN_SUCCESS

    ' 指示(仕様)と実績FLGを確認
    If (Samp.CRYINDEPCS <> "0") And (Samp.CRYRESEPCS <> "0") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, MEASURE, REGDATE "
        sql = sql & "from TBCMJ001 "
        sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "      SMPLNO = " & Samp.CRYSMPLIDEPCS & " and "
        sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ001 "
        sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "                       SMPLNO = " & Samp.CRYSMPLIDEPCS & ")"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With EPD
                .CRYNUM = rs("CRYNUM")          ' 結晶番号
                .POSITION = rs("POSITION")      ' 位置
                .SMPKBN = rs("SMPKBN")          ' サンプル区分
                .TRANCOND = rs("TRANCOND")      ' 処理条件
                .TRANCNT = rs("TRANCNT")        ' 処理回数
                .SMPLNO = rs("SMPLNO")          ' サンプルＮｏ
                .SMPLUMU = rs("SMPLUMU")        ' サンプル有無
                .MEASURE = rs("MEASURE")        ' 測定値
                .REGDATE = rs("REGDATE")        ' 登録日付
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    EPD_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'概要      :内部関数 X線実績取得用    2009/08/12 Kameda
Private Function X_Zisseki(XTALCS As String, x As type_DBDRV_scmzc_fcmkc001c_X, _
                             Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean

    NothingFlag = False

    ' EPD実績テーブルから値を取得

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function X_Zisseki"

    X_Zisseki = FUNCTION_RETURN_SUCCESS

        
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, XRAYX,XRAYY,XRAYXY, REGDATE "
    sql = sql & "from TBCMJ021 "
    sql = sql & "where CRYNUM = '" & XTALCS & "' and "
    'sql = sql & "      SMPLNO = " & Samp.CRYSMPLIDEPCS & " and "
    sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ021 "
    'sql = sql & "                 where CRYNUM = '" & Samp.XTALCS & "' and "
    'sql = sql & "                       SMPLNO = " & Samp.CRYSMPLIDEPCS & ")"
    sql = sql & "                 where CRYNUM = '" & XTALCS & "' )"
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount <> 0 Then
        With x
            .CRYNUM = rs("CRYNUM")          ' 結晶番号
            .POSITION = rs("POSITION")      ' 位置
            .SMPKBN = rs("SMPKBN")          ' サンプル区分
            .TRANCOND = rs("TRANCOND")      ' 処理条件
            .TRANCNT = rs("TRANCNT")        ' 処理回数
            .SMPLNO = rs("SMPLNO")          ' サンプルＮｏ
            .SMPLUMU = rs("SMPLUMU")        ' サンプル有無
            .XX = rs("XRAYX")               ' 測定値X
            .XY = rs("XRAYY")               ' 測定値Y
            .XXY = rs("XRAYXY")             ' 測定値XY
            .REGDATE = rs("REGDATE")        ' 登録日付
        End With
    Else
        NothingFlag = True
    End If
    
    Set rs = Nothing
    
    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    X_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'概要      :内部関数 SIRD実績取得用    2010/02/04 Kameda
Private Function SIRD_Zisseki(Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, SIRD As type_DBDRV_scmzc_fcmkc001c_SIRD, _
                             Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean

    NothingFlag = False
    SIRD.NothingFlg = ""      '2010/02/18 Kameda
    ' SIRD実績テーブルから値を取得

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function SIRD_Zisseki"

    SIRD_Zisseki = FUNCTION_RETURN_SUCCESS

    If Samp.SIRDKBNY3 = "1" Then
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, SIRDCNT, REGDATE "
        sql = sql & "from TBCMJ022 "
        sql = sql & "where CRYNUM = '" & Samp.XTALCS & "' and "
        sql = sql & "      TRANCNT = '0'"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
        If rs.RecordCount <> 0 Then
            With SIRD
                .CRYNUM = rs("CRYNUM")          ' 結晶番号
                .POSITION = rs("POSITION")      ' 位置
                .SMPKBN = rs("SMPKBN")          ' サンプル区分
                .TRANCOND = rs("TRANCOND")      ' 処理条件
                .TRANCNT = rs("TRANCNT")        ' 処理回数
                .SMPLNO = rs("SMPLNO")          ' サンプルＮｏ
                .SMPLUMU = rs("SMPLUMU")        ' サンプル有無
                .SIRDCNT = rs("SIRDCNT")        ' 測定値
                .REGDATE = rs("REGDATE")        ' 登録日付
            End With
        Else
            NothingFlag = True
            SIRD.NothingFlg = "1"    '2010/02/18 Kameda
        End If
        
        Set rs = Nothing
    End If
    
    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    SIRD_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :加工実績判定に構造体に値をセットする
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型             ,説明
'          :BLOCKID       ,   ,String         ,ブロックID
'          :Kakou         ,   ,type_KakouJudg ,加工実績判定構造体
'          :戻り値        ,O  ,FUNCTION_RETURN,
'説明      :ブロック内全品番の仕様と実績を求める
'履歴      :2002/4/16 佐野 作成
Public Function DBDRV_scmzc_fcmkc001c_Kakou(BLOCKID As String, Kakou As type_KakouJudg) As FUNCTION_RETURN
    Dim sql     As String
    Dim sql1    As String
    Dim rs      As OraDynaset
    Dim recCnt  As Integer
    Dim c0      As Integer
    Dim tHIN()  As tFullHinban

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function DBDRV_scmzc_fcmkc001c_Kakou"

    DBDRV_scmzc_fcmkc001c_Kakou = FUNCTION_RETURN_FAILURE

    'ブロック内の全品番を求める
    'ﾌﾞﾛｯｸ管理(TBCME040)参照停止　05/10/03 ooba START ======================================>
    sql = "select HINBAN, REVNUM, FACTORY, OPECOND from XSDC2 C2, TBCME041 E41 "
    sql = sql & "Where E41.CRYNUM = C2.XTALC2 and "
    sql = sql & "C2.CRYNUMC2 = '" & BLOCKID & "' and "
    sql = sql & "C2.INPOSC2 < E41.INGOTPOS+E41.LENGTH and "
    sql = sql & "C2.INPOSC2+C2.GNLC2 > E41.INGOTPOS"
    'ﾌﾞﾛｯｸ管理(TBCME040)参照停止　05/10/03 ooba END ========================================>

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    ReDim tHIN(recCnt)
    If recCnt = 0 Then
        rs.Close
        GoTo proc_exit
    End If
    For c0 = 1 To recCnt
        tHIN(c0).hinban = rs("HINBAN")
        tHIN(c0).mnorevno = rs("REVNUM")
        tHIN(c0).factory = rs("FACTORY")
        tHIN(c0).opecond = rs("OPECOND")
        rs.MoveNext
    Next
    rs.Close
    
    '求めた全品番の加工仕様を求める
    If scmzc_getKakouSpec(tHIN(), Kakou.Spec()) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    
    '対象ブロックの加工実績を求める
    If scmzc_getKakouJiltuseki(BLOCKID, Kakou.Jiltuseki) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    
    DBDRV_scmzc_fcmkc001c_Kakou = FUNCTION_RETURN_SUCCESS

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
'概要      :面傾き判定X線検査状態、実績フラグ取得
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :CrySmp        ,IO  ,Double       ,
'履歴      :2009/08/12
Private Function GetXSDC1_XRAY(CrySmp As type_DBDRV_scmzc_fcmkc001c_CrySmp) As FUNCTION_RETURN
    Dim sql             As String
    Dim rs              As OraDynaset
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function GetXSDC1_XRAY"

    GetXSDC1_XRAY = FUNCTION_RETURN_FAILURE
    
    sql = "select "
    sql = sql & "NVL(CRYINDXC1,'0') as CRYINDXC1 "         ' 状態FLG(X線)
    sql = sql & ",NVL(CRYRESXC1,'0') as CRYRESXC1 "        ' 実績FLG(X線)
    sql = sql & " from XSDC1"
    sql = sql & " where XTALC1 = '" & CrySmp.XTALCS & "'"
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    
    If rs.RecordCount <> 0 Then
        CrySmp.CRYINDXC1 = rs("CRYINDXC1")
        CrySmp.CRYRESXC1 = rs("CRYRESXC1")
    End If
    
    rs.Close

    GetXSDC1_XRAY = FUNCTION_RETURN_SUCCESS
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
'概要      :SIRD評価区分取得
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :CrySmp        ,IO  ,Double       ,
'履歴      :2010/02/04
Private Function GetXODY3_SIRD(CrySmp As type_DBDRV_scmzc_fcmkc001c_CrySmp) As FUNCTION_RETURN
    Dim sql             As String
    Dim rs              As OraDynaset
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function GetXODY3_SIRD"

    GetXODY3_SIRD = FUNCTION_RETURN_FAILURE
    
    sql = "select "
    sql = sql & "NVL(SIRDKBNY3,'0') as SIRDKBNY3 "         '
    sql = sql & " from XODY3"
    sql = sql & " where XTALNOY3 = '" & CrySmp.CRYNUMCS & "'"
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    
    If rs.RecordCount <> 0 Then
        CrySmp.SIRDKBNY3 = rs("SIRDKBNY3")
    End If
    
    rs.Close

    GetXODY3_SIRD = FUNCTION_RETURN_SUCCESS
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

'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の実績情報取得
'概要      :内部関数 Cu-Deco C 実績取得用
Private Function CuDeco_C_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                             Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                             Cu_deco_C As type_DBDRV_scmzc_fcmkc001c_C, _
                             TorB As Integer, _
                             Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean
    Dim wCryIND     As String
    Dim wCryRES     As String
    Dim wCrySMPL    As Long

    NothingFlag = False

    ' Cu_deco実績テーブル(TBCMJ023)から値を取得

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function CuDeco_C_Zisseki"

    CuDeco_C_Zisseki = FUNCTION_RETURN_SUCCESS

    wCryIND = Samp.CRYINDCCS        ' 状態フラグ C
    wCryRES = Samp.CRYRESCCS        ' 実績フラグ C
    wCrySMPL = Samp.CRYSMPLIDCCS    ' サンプルID C

    ' 指示(仕様)と実績FLGを確認
    If (wCryIND <> "0") And (wCryRES <> "0") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCNT, SMPLNO"
        sql = sql & ", SMPLUMUC, REGDATEC"
        sql = sql & ", CPTNJSK, CDISKJSK, CRINGNKJSK, CRINGGKJSK, CHANTEI"
        
        sql = sql & " from TBCMJ023"
        sql = sql & " where CRYNUM = '" & Samp.XTALCS & "' and"
        sql = sql & "       SMPLNO = " & wCrySMPL & " and"
        sql = sql & "       TRANCNT = (select max(TRANCNT) from TBCMJ023"
        sql = sql & "                  where CRYNUM = '" & Samp.XTALCS & "' and"
        sql = sql & "                        SMPLNO = " & wCrySMPL & ")"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With Cu_deco_C
                
                If IsNull(rs("CRYNUM")) = False Then .CRYNUM = rs("CRYNUM")             ' 結晶番号
                .POSITION = CInt(fncNullCheck(rs("POSITION")))                          ' 位置
                If IsNull(rs("SMPKBN")) = False Then .SMPKBN = rs("SMPKBN")             ' サンプル区分
                .TRANCNT = CInt(fncNullCheck(rs("TRANCNT")))                            ' 処理回数
                .SMPLNO = CLng(fncNullCheck(rs("SMPLNO")))                              ' サンプルＮｏ
                If IsNull(rs("SMPLUMUC")) = False Then .SMPLUMUC = rs("SMPLUMUC")       ' サンプル有無 C
                
                If IsNull(rs("CPTNJSK")) = False Then .CPTNJSK = rs("CPTNJSK")          ' C パターン実績
                
                .CDISKJSK = CInt(fncNullCheck(rs("CDISKJSK")))                          ' C Disk半径実績
                .CRINGNKJSK = CInt(fncNullCheck(rs("CRINGNKJSK")))                      ' C Ring内径実績
                .CRINGGKJSK = CInt(fncNullCheck(rs("CRINGGKJSK")))                      ' C Ring外径実績
                
                If IsNull(rs("CHANTEI")) = False Then .CHANTEI = rs("CHANTEI")          ' C 判定結果
                
                If IsNull(rs("REGDATEC")) = False Then .REGDATE = rs("REGDATEC")        ' 登録日付
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    CuDeco_C_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'Add End   2011/01/17 SMPK A.Nagamine


'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の実績情報取得
'概要      :内部関数 Cu-Deco CJ 実績取得用
Private Function CuDeco_CJ_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                             Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                             Cu_deco_CJ As type_DBDRV_scmzc_fcmkc001c_CJ, _
                             TorB As Integer, _
                             Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean
    Dim wCryIND     As String
    Dim wCryRES     As String
    Dim wCrySMPL    As Long

    NothingFlag = False

    ' Cu_deco実績テーブル(TBCMJ023)から値を取得

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function CuDeco_CJ_Zisseki"

    CuDeco_CJ_Zisseki = FUNCTION_RETURN_SUCCESS

    wCryIND = Samp.CRYINDCJCS           ' 状態フラグ CJ
    wCryRES = Samp.CRYRESCJCS           ' 実績フラグ CJ
    wCrySMPL = Samp.CRYSMPLIDCJCS       ' サンプルID CJ

    ' 指示(仕様)と実績FLGを確認
    If (wCryIND <> "0") And (wCryRES <> "0") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCNT, SMPLNO"
        sql = sql & ", SMPLUMUCJ, REGDATECJ"
        sql = sql & ", CJPTNJSK, CJDISKJSK, CJRINGNKJSK, CJRINGGKJSK, CJBANDNKJSK"
        sql = sql & ", CJBANDGKJSK, CJRINGCALC, CJPICALC, CJHANTEI, CJDMAXPIC5"
'Chg Start 2012/05/22 SMPK H.Ohkubo CLESTA評価レベル判定対応
'        sql = sql & ", CJRMAXPIC5, CJDRMAXPIC5, CJALLMAXDIC5, CJALLMINRINC5, CJALLMAXRIGC5"
        sql = sql & ", CJRMAXPIC5, CJDRMAXPIC5, CJALLMINDIC5, CJALLMAXDIC5, CJALLMINRINC5, CJALLMAXRIGC5"
'Chg End 2012/05/22 SMPK H.Ohkubo CLESTA評価レベル判定対応
        
        sql = sql & " from TBCMJ023"
        sql = sql & " where CRYNUM = '" & Samp.XTALCS & "' and"
        sql = sql & "       SMPLNO = " & wCrySMPL & " and"
        sql = sql & "       TRANCNT = (select max(TRANCNT) from TBCMJ023"
        sql = sql & "                  where CRYNUM = '" & Samp.XTALCS & "' and"
        sql = sql & "                        SMPLNO = " & wCrySMPL & ")"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With Cu_deco_CJ
                
                If IsNull(rs("CRYNUM")) = False Then .CRYNUM = rs("CRYNUM")             ' 結晶番号
                .POSITION = CInt(fncNullCheck(rs("POSITION")))                          ' 位置
                If IsNull(rs("SMPKBN")) = False Then .SMPKBN = rs("SMPKBN")             ' サンプル区分
                .TRANCNT = CInt(fncNullCheck(rs("TRANCNT")))                            ' 処理回数
                .SMPLNO = CLng(fncNullCheck(rs("SMPLNO")))                              ' サンプルＮｏ
                If IsNull(rs("SMPLUMUCJ")) = False Then .SMPLUMUCJ = rs("SMPLUMUCJ")    ' サンプル有無 CJ
                
                If IsNull(rs("CJPTNJSK")) = False Then .CJPTNJSK = rs("CJPTNJSK")                   ' CJ パターン実績
                
                .CJDISKJSK = CInt(fncNullCheck(rs("CJDISKJSK")))                                    ' CJ Disk半径実績
                .CJRINGNKJSK = CInt(fncNullCheck(rs("CJRINGNKJSK")))                                ' CJ Ring内径実績
                .CJRINGGKJSK = CInt(fncNullCheck(rs("CJRINGGKJSK")))                                ' CJ Ring外径実績
                .CJBANDNKJSK = CInt(fncNullCheck(rs("CJBANDNKJSK")))                                ' CJ Band内径実績
                .CJBANDGKJSK = CInt(fncNullCheck(rs("CJBANDGKJSK")))                                ' CJ Band外径実績
                .CJRINGCALC = CInt(fncNullCheck(rs("CJRINGCALC")))                                  ' CJ Ring幅計算
                .CJPICALC = CInt(fncNullCheck(rs("CJPICALC")))                                      ' CJ Pi幅計算
                
                If IsNull(rs("CJHANTEI")) = False Then .CJHANTEI = rs("CJHANTEI")                   ' CJ 判定結果
                
                .CJDMAXPIC5 = CInt(fncNullCheck(rs("CJDMAXPIC5")))                                  ' CJ Diskのみパターン Pi幅上限値
                .CJRMAXPIC5 = CInt(fncNullCheck(rs("CJRMAXPIC5")))                                  ' CJ Ringのみパターン Pi幅上限値
                .CJDRMAXPIC5 = CInt(fncNullCheck(rs("CJDRMAXPIC5")))                                ' CJ DiskRingパターン Pi幅上限値
'Add Start 2012/05/22 SMPK H.Ohkubo CLESTA評価レベル判定対応
                .CJALLMINDIC5 = CInt(fncNullCheck(rs("CJALLMINDIC5")))                              ' CJ 共通Disk半径下限値
'Add End 2012/05/22 SMPK H.Ohkubo CLESTA評価レベル判定対応
                .CJALLMAXDIC5 = CInt(fncNullCheck(rs("CJALLMAXDIC5")))                              ' CJ 共通Disk半径上限値
                .CJALLMINRINC5 = CInt(fncNullCheck(rs("CJALLMINRINC5")))                            ' CJ 共通Ring内径下限値
                .CJALLMAXRIGC5 = CInt(fncNullCheck(rs("CJALLMAXRIGC5")))                            ' CJ 共通Ring外径上限値
                
                If IsNull(rs("REGDATECJ")) = False Then .REGDATE = rs("REGDATECJ")       ' 登録日付
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    CuDeco_CJ_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'Add End   2011/01/17 SMPK A.Nagamine


'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の実績情報取得
'概要      :内部関数 Cu-Deco CJ(LT) 実績取得用
Private Function CuDeco_CJLT_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                             Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                             Cu_deco_CJLT As type_DBDRV_scmzc_fcmkc001c_CJLT, _
                             TorB As Integer, _
                             Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean
    Dim wCryIND     As String
    Dim wCryRES     As String
    Dim wCrySMPL    As Long

    NothingFlag = False

    ' Cu_deco実績テーブル(TBCMJ023)から値を取得

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function CuDeco_CJLT_Zisseki"

    CuDeco_CJLT_Zisseki = FUNCTION_RETURN_SUCCESS

    wCryIND = Samp.CRYINDCJLTCS         ' 状態フラグ CJ(LT)
    wCryRES = Samp.CRYRESCJLTCS         ' 実績フラグ CJ(LT)
    wCrySMPL = Samp.CRYSMPLIDCJLTCS     ' サンプルID CJ(LT)

    ' 指示(仕様)と実績FLGを確認
    If (wCryIND <> "0") And (wCryRES <> "0") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCNT, SMPLNO"
        sql = sql & ", SMPLUMUCJLT, REGDATECJLT"
        sql = sql & ", CJLTPTNJSK, CJLTDISKJSK, CJLTRINGNKJSK, CJLTRINGGKJSK, CJLTBANDNKJSK"
        sql = sql & ", CJLTBANDGKJSK, CJLTRINGCALC, CJLTPICALC, CJLTHANTEI"
        
        sql = sql & " from TBCMJ023"
        sql = sql & " where CRYNUM = '" & Samp.XTALCS & "' and"
        sql = sql & "       SMPLNO = " & wCrySMPL & " and"
        sql = sql & "       TRANCNT = (select max(TRANCNT) from TBCMJ023"
        sql = sql & "                  where CRYNUM = '" & Samp.XTALCS & "' and"
        sql = sql & "                        SMPLNO = " & wCrySMPL & ")"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With Cu_deco_CJLT
                
                If IsNull(rs("CRYNUM")) = False Then .CRYNUM = rs("CRYNUM")             ' 結晶番号
                .POSITION = CInt(fncNullCheck(rs("POSITION")))                          ' 位置
                If IsNull(rs("SMPKBN")) = False Then .SMPKBN = rs("SMPKBN")             ' サンプル区分
                .TRANCNT = CInt(fncNullCheck(rs("TRANCNT")))                            ' 処理回数
                .SMPLNO = CLng(fncNullCheck(rs("SMPLNO")))                              ' サンプルＮｏ
                If IsNull(rs("SMPLUMUCJLT")) = False Then .SMPLUMUCJLT = rs("SMPLUMUCJLT")          ' サンプル有無 CJ(LT)
                
                If IsNull(rs("CJLTPTNJSK")) = False Then .CJLTPTNJSK = rs("CJLTPTNJSK")             ' CJ(LT) パターン実績
                
                .CJLTDISKJSK = CInt(fncNullCheck(rs("CJLTDISKJSK")))                                ' CJ(LT) Disk半径実績
                .CJLTRINGNKJSK = CInt(fncNullCheck(rs("CJLTRINGNKJSK")))                            ' CJ(LT) Ring内径実績
                .CJLTRINGGKJSK = CInt(fncNullCheck(rs("CJLTRINGGKJSK")))                            ' CJ(LT) Ring外径実績
                .CJLTBANDNKJSK = CInt(fncNullCheck(rs("CJLTBANDNKJSK")))                            ' CJ(LT) Band内径実績
                .CJLTBANDGKJSK = CInt(fncNullCheck(rs("CJLTBANDGKJSK")))                            ' CJ(LT) Band外径実績
                .CJLTRINGCALC = CInt(fncNullCheck(rs("CJLTRINGCALC")))                              ' CJ(LT) Ring幅計算
                .CJLTPICALC = CInt(fncNullCheck(rs("CJLTPICALC")))                                  ' CJ(LT) Pi幅計算
                
                If IsNull(rs("CJLTHANTEI")) = False Then .CJLTHANTEI = rs("CJLTHANTEI")             ' CJ(LT) 判定結果
                
                If IsNull(rs("REGDATECJLT")) = False Then .REGDATE = rs("REGDATECJLT")       ' 登録日付
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    CuDeco_CJLT_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'Add End   2011/01/17 SMPK A.Nagamine


'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の実績情報取得
'概要      :内部関数 Cu-Deco CJ2 実績取得用
Private Function CuDeco_CJ2_Zisseki(siyou As type_DBDRV_scmzc_fcmkc001c_Siyou, _
                             Samp As type_DBDRV_scmzc_fcmkc001c_CrySmp, _
                             Cu_deco_CJ2 As type_DBDRV_scmzc_fcmkc001c_CJ2, _
                             TorB As Integer, _
                             Optional NothingFlagStr$ = vbNullString) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim NothingFlag As Boolean
    Dim wCryIND     As String
    Dim wCryRES     As String
    Dim wCrySMPL    As Long

    NothingFlag = False

    ' Cu_deco実績テーブル(TBCMJ023)から値を取得

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryJudg_SQL.bas -- Function CuDeco_CJ2_Zisseki"

    CuDeco_CJ2_Zisseki = FUNCTION_RETURN_SUCCESS

    wCryIND = Samp.CRYINDCJ2CS          ' 状態フラグ CJ2
    wCryRES = Samp.CRYRESCJ2CS          ' 実績フラグ CJ2
    wCrySMPL = Samp.CRYSMPLIDCJ2CS      ' サンプルID CJ2

    ' 指示(仕様)と実績FLGを確認
    If (wCryIND <> "0") And (wCryRES <> "0") Then
        
        sql = "select CRYNUM, POSITION, SMPKBN, TRANCNT, SMPLNO"
        sql = sql & ", SMPLUMUCJ2, REGDATECJ2"
        sql = sql & ", CJ2PTNJSK, CJ2DISKJSK, CJ2RINGNKJSK, CJ2RINGGKJSK, CJ2PICALC"
        sql = sql & ", CJ2HANTEI, CJ2DMAXPIC5, CJ2RMAXPIC5, CJ2RMINRINC5, CJ2RMAXRIGC5"
        sql = sql & ", CJ2DRMAXPIC5, CJ2DRMINRINC5, CJ2DRMAXRIGC5"
        
        sql = sql & " from TBCMJ023"
        sql = sql & " where CRYNUM = '" & Samp.XTALCS & "' and"
        sql = sql & "       SMPLNO = " & wCrySMPL & " and"
        sql = sql & "       TRANCNT = (select max(TRANCNT) from TBCMJ023"
        sql = sql & "                  where CRYNUM = '" & Samp.XTALCS & "' and"
        sql = sql & "                        SMPLNO = " & wCrySMPL & ")"
        
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount <> 0 Then
            With Cu_deco_CJ2
                
                If IsNull(rs("CRYNUM")) = False Then .CRYNUM = rs("CRYNUM")             ' 結晶番号
                .POSITION = CInt(fncNullCheck(rs("POSITION")))                          ' 位置
                If IsNull(rs("SMPKBN")) = False Then .SMPKBN = rs("SMPKBN")             ' サンプル区分
                .TRANCNT = CInt(fncNullCheck(rs("TRANCNT")))                            ' 処理回数
                .SMPLNO = CLng(fncNullCheck(rs("SMPLNO")))                              ' サンプルＮｏ
                If IsNull(rs("SMPLUMUCJ2")) = False Then .SMPLUMUCJ2 = rs("SMPLUMUCJ2")             ' サンプル有無CJ2
                
                If IsNull(rs("CJ2PTNJSK")) = False Then .CJ2PTNJSK = rs("CJ2PTNJSK")                ' CJ2 パターン実績
                
                .CJ2DISKJSK = CInt(fncNullCheck(rs("CJ2DISKJSK")))                                  ' CJ2 Disk半径実績
                .CJ2RINGNKJSK = CInt(fncNullCheck(rs("CJ2RINGNKJSK")))                              ' CJ2 Ring内径実績
                .CJ2RINGGKJSK = CInt(fncNullCheck(rs("CJ2RINGGKJSK")))                              ' CJ2 Ring外径実績
                .CJ2PICALC = CInt(fncNullCheck(rs("CJ2PICALC")))                                    ' CJ2 Pi幅計算
                
                If IsNull(rs("CJ2HANTEI")) = False Then .CJ2HANTEI = rs("CJ2HANTEI")                ' CJ2 判定結果
                
                .CJ2DMAXPIC5 = CInt(fncNullCheck(rs("CJ2DMAXPIC5")))                                ' CJ2 Diskのみパターン Pi幅下限値
                .CJ2RMAXPIC5 = CInt(fncNullCheck(rs("CJ2RMAXPIC5")))                                ' CJ2 Ringのみパターン Pi幅下限値
                .CJ2RMINRINC5 = CInt(fncNullCheck(rs("CJ2RMINRINC5")))                              ' CJ2 Ringのみパターン Ring内径下限値
                .CJ2RMAXRIGC5 = CInt(fncNullCheck(rs("CJ2RMAXRIGC5")))                              ' CJ2 Ringのみパターン Ring外径上限値
                .CJ2DRMAXPIC5 = CInt(fncNullCheck(rs("CJ2DRMAXPIC5")))                              ' CJ2 DiskRingパターン Pi幅下限値
                .CJ2DRMINRINC5 = CInt(fncNullCheck(rs("CJ2DRMINRINC5")))                            ' CJ2 DiskRingパターン Ring内径下限値
                .CJ2DRMAXRIGC5 = CInt(fncNullCheck(rs("CJ2DRMAXRIGC5")))                            ' CJ2 DiskRingパターン Ring外径上限値
                
                If IsNull(rs("REGDATECJ2")) = False Then .REGDATE = rs("REGDATECJ2")       ' 登録日付
            End With
        Else
            NothingFlag = True
        End If
        
        Set rs = Nothing
    End If

    If NothingFlagStr <> vbNullString Then
        If NothingFlag Then
            NothingFlagStr = "1"
        End If
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    CuDeco_CJ2_Zisseki = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'Add End   2011/01/17 SMPK A.Nagamine


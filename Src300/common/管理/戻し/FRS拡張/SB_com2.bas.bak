Attribute VB_Name = "SB_Com2"
Option Explicit

'Add Start 2010/12/23 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : C-OSF3判定テーブルのキー変更(Class 14 -> 22), Cu-deco(C, CJ, CJ(LT), CJ2)の仕様判定追加
' CLESTA評価対応(Cu-deco)
' 品番振替チェック（仕様チェック２－２、実績チェック２－１）を改造する。
'Add End   2010/12/23 SMPK A.Nagamine

'Public Type t_FullHinban
'    hinban As String * 12        ' 振替元品番
'End Type
'振替元品番
Public tOld_Hinban As tFullHinban           ' 振替元品番データ
'振替先品番
Public tNew_Hinban As tFullHinban           ' 振替先品番データ

'仕様取得構造体
Public Type typ_chk1_1
    'HWFTYPE     As String * 1       'タイプ
    HSXFTYPE     As String * 1       'タイプ 2004/12/21 名称変更
    BLOCKHFLAG  As String * 1       'ブロック単位保証フラグ
    HSXSDSLP    As String * 1       'シード傾き     2009/08/06追加 SETsw kubota
End Type
Public tbl_chk1_1(1) As typ_chk1_1

Public Type typ_chk1_2
    HSXCDIR     As String * 1       '結晶面方位
    HSXCSCEN    As Double           '結晶面傾き中心
    HSXDOP      As String * 1       'ドーパント
    HWFCDOP     As String * 1       '結晶ドープ
'    HSXSDSLP    As String * 1       'シード傾き    2009/08/06削除 SETsw kubota
    HSXDPDIR    As String * 2       '溝位置方位
    MCNO1       As String * 1       '品種
    MCNO2       As String * 1       '引上げ速度
    MCNO3       As String * 1       'HZタイプ
    DCHYUUBU    As String * 1       'ドローチューブ
    NDOPHUFLG   As String * 1       '窒素ドープ振替可能チェック add 0108
    CDOPHUFLG   As String * 1       'Cドープ振替可能チェック    add 0108
    HWFSIRDHS   As String * 1       'SIRD保証方法 処 2010/05/24 SIRD対応 Y.Hitomi
End Type
Public tbl_chk1_2(1) As typ_chk1_2

Public Type typ_chk1_3
    HSXD1MIN    As Double           '品ＳＸ直径１下限
    HSXD1MAX    As Double           '品ＳＸ直径１上限
    HSXDWMIN    As Double           '品ＳＸ溝巾下限
    HSXDWMAX    As Double           '品ＳＸ溝巾上限
    HSXDDMIN    As Double           '品ＳＸ溝深下限
    HSXDDMAX    As Double           '品ＳＸ溝深上限
    HWFWARPR    As String * 1       'Warpランク
End Type
Public tbl_chk1_3(1) As typ_chk1_3

Public Type typ_chk1_4
    HSXRHWYS    As String * 1       '保証方法_対象
    HSXONHWS    As String * 1       '保証方法_対象
    HSXONSPT    As String * 1       '測定位置_点        '08/01/29 ooba
    HSXONSPI    As String * 1       '測定位置_位
    HSXONKWY    As String * 2       '検査方法
    HSXOF1HS    As String * 1       '保証方法_対象
    HSXOF1SH    As String * 1       '測定位置_方
    HSXOF1ST    As String * 1       '測定位置_点
    HSXOF1SR    As String * 1       '測定位置_領
    HSXOF1NS    As String * 2       '熱処理法
    HSXOF1SZ    As String * 1       '測定条件
    HSXOF1ET    As Integer          '選択ET代
    HSXOSF1PTK  As String * 1       'パターン区分
    HSXOF2HS    As String * 1       '保証方法_対象
    HSXOF2SH    As String * 1       '測定位置_方
    HSXOF2ST    As String * 1       '測定位置_点
    HSXOF2SR    As String * 1       '測定位置_領
    HSXOF2NS    As String * 2       '熱処理法
    HSXOF2SZ    As String * 1       '測定条件
    HSXOF2ET    As Integer          '選択ET代
    HSXOSF2PTK  As String * 1       'パターン区分
    HSXOF3HS    As String * 1       '保証方法_対象
    HSXOF3SH    As String * 1       '測定位置_方
    HSXOF3ST    As String * 1       '測定位置_点
    HSXOF3SR    As String * 1       '測定位置_領
    HSXOF3NS    As String * 2       '熱処理法
    HSXOF3SZ    As String * 1       '測定条件
    HSXOF3ET    As Integer          '選択ET代
    HSXOSF3PTK  As String * 1       'パターン区分
    HSXOF4HS    As String * 1       '保証方法_対象    <--- ArANでこのｴﾘｱを使用！
    HSXOF4SH    As String * 1       '測定位置_方
    HSXOF4ST    As String * 1       '測定位置_点
    HSXOF4SR    As String * 1       '測定位置_領
    HSXOF4NS    As String * 2       '熱処理法
    HSXOF4SZ    As String * 1       '測定条件
    HSXOF4ET    As Integer          '選択ET代
    HSXOSF4PTK  As String * 1       'パターン区分    <--- ArANでこのｴﾘｱを使用！
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    HWFSIRDMX   As Integer          '軸状転位上限(SIRD)
    HWFSIRDSZ   As String * 1       '軸状転位測定条件(SIRD)
    HWFSIRDHT   As String * 1       '軸状転位保証方法＿対(SIRD)
    HWFSIRDHS   As String * 1       '軸状転位保証方法＿処(SIRD)
    HWFSIRDKM   As String * 1       '軸状転位検査頻度＿枚(SIRD)
    HWFSIRDKH   As String * 1       '軸状転位検査頻度＿保(SIRD)
    HWFSIRDKU   As String * 1       '軸状転位検査頻度＿ウ(SIRD)
    HWFSIRDPS   As String * 2       '軸状転位TB保証位置(SIRD)
    HWFSIRDKN   As String * 1       '軸状転位検査頻度_抜(SIRD)
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
    HSXBM1HS    As String * 1       '保証方法_対象
    HSXBM1SH    As String * 1       '測定位置_方
    HSXBM1ST    As String * 1       '測定位置_点
    HSXBM1SR    As String * 1       '測定位置_領
    HSXBM1NS    As String * 2       '熱処理法
    HSXBM1SZ    As String * 1       '測定条件
    HSXBM1ET    As Integer          '選択ET代
    HSXBM2HS    As String * 1       '保証方法_対象
    HSXBM2SH    As String * 1       '測定位置_方
    HSXBM2ST    As String * 1       '測定位置_点
    HSXBM2SR    As String * 1       '測定位置_領
    HSXBM2NS    As String * 2       '熱処理法
    HSXBM2SZ    As String * 1       '測定条件
    HSXBM2ET    As Integer          '選択ET代
    HSXBM3HS    As String * 1       '保証方法_対象
    HSXBM3SH    As String * 1       '測定位置_方
    HSXBM3ST    As String * 1       '測定位置_点
    HSXBM3SR    As String * 1       '測定位置_領
    HSXBM3NS    As String * 2       '熱処理法
    HSXBM3SZ    As String * 1       '測定条件
    HSXBM3ET    As Integer          '選択ET代
    HSXTMMAX    As Long             '上限
    HSXLTHWS    As String * 1       '保証方法_対象
    HSXCNHWS    As String * 1       '保証方法_対象
    HSXCNKWY    As String * 2       '検査方法
    HSXDENHS    As String * 1       '保証方法_対象
    HSXDENMN    As Integer          '下限
    HSXDENMX    As Integer          '上限
    HSXDVDHS    As String * 1       '保証方法_対象
    HSXDVDMNN   As Integer          '下限
    HSXDVDMXN   As Integer          '上限
    HSXLDLHS    As String * 1       '保証方法_対象
    HSXLDLMN    As Integer          '下限
    HSXLDLMX    As Integer          '上限
'*** UPDATE ↓ Y.SIMIZU 2005/10/12 GDﾗｲﾝ数追加
    HSXGDLINE   As String           '測定条件
'*** UPDATE ↑ Y.SIMIZU 2005/10/12 GDﾗｲﾝ数追加

'--------------- 2008/08/25 INSERT START  By Systech ---------------
    HSXDKTMP    As String * 1       ' DK温度
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    HSXCNKHI    As String * 1       ' 品SX炭素濃度検査頻度＿位   '' add 0108
    
    'Add Start 2010/12/23 SMPK A.Nagamine       : CLESTA評価対応(Cu-deco) : C-OSF3判定テーブルのキー変更(Class 14 -> 22), Cu-deco(C, CJ, CJ(LT), CJ2)の仕様判定追加
    
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
    
    'Add End   2010/12/23 SMPK A.Nagamine
End Type
Public tbl_chk1_4(1) As typ_chk1_4

Public Type typ_chk1_4_1
    HOSYOU      As String * 1       '保証方法＿対象
    Min         As Integer          '下限
    max         As Integer          '上限
    SOKU_HOU    As String * 1       '測定位置＿方
    SOKU_TEN    As String * 1       '測定位置＿点
    SOKU_ICHI   As String * 1       '測定位置＿位
    SOKU_RYOU   As String * 1       '測定位置＿領
    UMU         As String * 1       '検査有無           ????????????????(桁数）
    NETSU       As String * 2       '熱処理法
    JOUKEN      As String * 1       '測定条件
    ET          As Integer          '選択ＥＴ代
    KENSA       As String * 2       '検査方法
'*** UPDATE ↓ Y.SIMIZU 2005/10/12 STRING型に変更
'    LINE        As Integer          'ライン数           ????????????????(桁数）
    LINE        As String          'ライン数           ????????????????(桁数）
'*** UPDATE ↑ Y.SIMIZU 2005/10/12 STRING型に変更
    PATTERN     As String * 1       'パターン区分
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    HSXDKTMP    As String * 1       ' DK温度
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    HSXCNKHI    As String * 1       ' 品SX炭素濃度検査頻度＿位   '' add 0108
    
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    HWFSIRDMX   As Integer          '軸状転位上限(SIRD)
    HWFSIRDHT   As String * 1       '軸状転位保証方法＿対(SIRD)
    HWFSIRDKM   As String * 1       '軸状転位検査頻度＿枚(SIRD)
    HWFSIRDKH   As String * 1       '軸状転位検査頻度＿保(SIRD)
    HWFSIRDKU   As String * 1       '軸状転位検査頻度＿ウ(SIRD)
    HWFSIRDKN   As String * 1       '軸状転位検査頻度_抜(SIRD)
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
End Type
Public tbl_chk1_4_1(1) As typ_chk1_4_1

Public Type typ_chk1_5
    BLOCKHFLAG  As String * 1       'ブロック単位保証フラグ　05/07/29 ooba
    HWFRHWYS    As String * 1       '保証方法＿対象
    HWFONHWS    As String * 1       '保証方法＿対象
    HWFONSPT    As String * 1       '測定位置＿点       '08/01/29 ooba
    HWFOF1HS    As String * 1       '保証方法＿対象
    HWFOF1SH    As String * 1       '測定位置＿方
    HWFOF1SR    As String * 1       '測定位置＿領
    HWFOF1NS    As String * 2       '熱処理法
    HWFOF1SZ    As String * 1       '測定条件
    HWFOF1ET    As Integer          '選択ＥＴ代
    HWFOSF1PTK  As String * 1       'パターン区分
    HWFOF2HS    As String * 1       '保証方法＿対象
    HWFOF2SH    As String * 1       '測定位置＿方
    HWFOF2SR    As String * 1       '測定位置＿領
    HWFOF2NS    As String * 2       '熱処理法
    HWFOF2SZ    As String * 1       '測定条件
    HWFOF2ET    As Integer          '選択ＥＴ代
    HWFOSF2PTK  As String * 1       'パターン区分
    HWFOF3HS    As String * 1       '保証方法＿対象
    HWFOF3SH    As String * 1       '測定位置＿方
    HWFOF3SR    As String * 1       '測定位置＿領
    HWFOF3NS    As String * 2       '熱処理法
    HWFOF3SZ    As String * 1       '測定条件
    HWFOF3ET    As Integer          '選択ＥＴ代
    HWFOSF3PTK  As String * 1       'パターン区分
    
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''    HWFOF4HS    As String * 1       '保証方法＿対象
'''    HWFOF4SH    As String * 1       '測定位置＿方
'''    HWFOF4SR    As String * 1       '測定位置＿領
'''    HWFOF4NS    As String * 2       '熱処理法
'''    HWFOF4SZ    As String * 1       '測定条件
'''    HWFOF4ET    As Integer          '選択ＥＴ代
'''    HWFOSF4PTK  As String * 1       'パターン区分

    HWFSIRDMX   As Integer          '軸状転位上限(SIRD)
    HWFSIRDSZ   As String * 1       '軸状転位測定条件(SIRD)
    HWFSIRDHT   As String * 1       '軸状転位保証方法＿対(SIRD)
    HWFSIRDHS   As String * 1       '軸状転位保証方法＿処(SIRD)
    HWFSIRDKM   As String * 1       '軸状転位検査頻度＿枚(SIRD)
    HWFSIRDKH   As String * 1       '軸状転位検査頻度＿保(SIRD)
    HWFSIRDKU   As String * 1       '軸状転位検査頻度＿ウ(SIRD)
    HWFSIRDPS   As String * 2       '軸状転位TB保証位置(SIRD)
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)

    HWFBM1HS    As String * 1       '保証方法＿対象
    HWFBM1SH    As String * 1       '測定位置＿方
    HWFBM1ST    As String * 1       '測定位置＿点
    HWFBM1SR    As String * 1       '測定位置＿領
    HWFBM1NS    As String * 2       '熱処理法
    HWFBM1SZ    As String * 1       '測定条件
    HWFBM1ET    As Integer          '選択ＥＴ代
    HWFBM2HS    As String * 1       '保証方法＿対象
    HWFBM2SH    As String * 1       '測定位置＿方
    HWFBM2ST    As String * 1       '測定位置＿点
    HWFBM2SR    As String * 1       '測定位置＿領
    HWFBM2NS    As String * 2       '熱処理法
    HWFBM2SZ    As String * 1       '測定条件
    HWFBM2ET    As Integer          '選択ＥＴ代
    HWFBM3HS    As String * 1       '保証方法＿対象
    HWFBM3SH    As String * 1       '測定位置＿方
    HWFBM3ST    As String * 1       '測定位置＿点
    HWFBM3SR    As String * 1       '測定位置＿領
    HWFBM3NS    As String * 2       '熱処理法
    HWFBM3SZ    As String * 1       '測定条件
    HWFBM3ET    As Integer          '選択ＥＴ代
    HWFOS1HS    As String * 1       '保証方法＿対象
    HWFOS1NS    As String * 2       '熱処理法
    HWFOS2HS    As String * 1       '保証方法＿対象
    HWFOS2NS    As String * 2       '熱処理法
    HWFOS3HS    As String * 1       '保証方法＿対象
    HWFOS3NS    As String * 2       '熱処理法
    HWFDSOHS    As String * 1       '保証方法＿対象
    HWFDSONWY   As String * 2       '熱処理法
    HWFDSOPTK   As String * 1       'パターン区分       'DSODﾊﾟﾀｰﾝ区分追加　04/07/28 ooba
    HWFMKHWS    As String * 1       '保証方法＿対象
    HWFMKSPH    As String * 1       '測定位置＿方
    HWFMKSPT    As String * 1       '測定位置＿点
    HWFMKSPR    As String * 1       '測定位置＿領
    HWFMKNSW    As String * 2       '熱処理法
    HWFMKSZY    As String * 1       '測定条件
    HWFMKCET    As Integer          '選択ＥＴ代
    HWFSPVHS    As String * 1       '保証方法＿対象
    HWFSPVST    As String * 1       '測定位置＿点
    HWFDLHWS    As String * 1       '保証方法＿対象
    HWFZOHWS    As String * 1       '保証方法＿対象     ''残存酸素追加　03/12/09 ooba
    HWFZONSW    As String * 2       '熱処理法           ''残存酸素追加　03/12/09 ooba
    
    HWFDENHS    As String * 1       '保証方法＿対象     'GD追加　05/01/27 ooba START ====>
    HWFDENMN    As Integer          '下限
    HWFDENMX    As Integer          '上限
    HWFDVDHS    As String * 1       '保証方法＿対象
    HWFDVDMNN   As Integer          '下限
    HWFDVDMXN   As Integer          '上限
    HWFLDLHS    As String * 1       '保証方法＿対象
    HWFLDLMN    As Integer          '下限
    HWFLDLMX    As Integer          '上限
    HWFGDKHN    As String * 1       '検査頻度_抜(GD)    'GD追加　05/01/27 ooba END ======>
    
    HWFRKHNN    As String * 1       ' 検査頻度_抜(Rs)   '追加　04/04/13 ooba START ====>
    HWFONKHN    As String * 1       ' 検査頻度_抜(Oi)
    HWFOF1KN    As String * 1       ' 検査頻度_抜(L1)
    HWFOF2KN    As String * 1       ' 検査頻度_抜(L2)
    HWFOF3KN    As String * 1       ' 検査頻度_抜(L3)
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''    HWFOF4KN    As String * 1       ' 検査頻度_抜(L4)
    HWFSIRDKN   As String * 1       ' 検査頻度_抜(SIRD)
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
    HWFBM1KN    As String * 1       ' 検査頻度_抜(B1)
    HWFBM2KN    As String * 1       ' 検査頻度_抜(B2)
    HWFBM3KN    As String * 1       ' 検査頻度_抜(B3)
    HWFOS1KN    As String * 1       ' 検査頻度_抜(D1)
    HWFOS2KN    As String * 1       ' 検査頻度_抜(D2)
    HWFOS3KN    As String * 1       ' 検査頻度_抜(D3)
    HWFDSOKN    As String * 1       ' 検査頻度_抜(DS)
    HWFMKKHN    As String * 1       ' 検査頻度_抜(DZ)
    HWFSPVKN    As String * 1       ' 検査頻度_抜(SP/Fe濃度)
    HWFDLKHN    As String * 1       ' 検査頻度_抜(SP/拡散長)
    HWFZOKHN    As String * 1       ' 検査頻度_抜(AO)   '追加　04/04/13 ooba END ======>

''Upd Start 2005/06/16 (TCS)T.Terauchi      SPV9点対応
    HWFSPVSH    As String * 1       ' 測定位置＿方(SPVFE)
    HWFSPVSI    As String * 1       ' 測定位置＿位(SPVFE)
    HWFDLSPH    As String * 1       ' 測定位置＿方(拡散長)
    HWFDLSPT    As String * 1       ' 測定位置＿点(拡散長)
    HWFDLSPI    As String * 1       ' 測定位置＿位(拡散長)
''Upd End   2005/06/16 (TCS)T.Terauchi      SPV9点対応
'*** UPDATE ↓ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数追加
    HWFGDLINE   As String           ' 測定条件(GDﾗｲﾝ数)
'*** UPDATE ↑ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数追加
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    HWFGDSZY    As String * 1       'GD測定条件
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---

'↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.1 AN温度 振替可否チェック追加
    HWFANTNP    As Integer          ' 品ＷＦＡＮ温度
'↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    HWFSPVPUG   As Double           ' PUA限(SPVFE)      '追加　06/05/31 ooba START =======>
    HWFSPVPUR   As Double           ' PUA率(SPVFE)
    HWFDLPUG    As Double           ' PUA限(拡散長)
    HWFDLPUR    As Double           ' PUA率(拡散長)
    HWFNRHS     As String * 1       ' 保証方法＿対象(SPVNR)
    HWFNRSH     As String * 1       ' 測定位置＿方(SPVNR)
    HWFNRST     As String * 1       ' 測定位置＿点(SPVNR)
    HWFNRSI     As String * 1       ' 測定位置＿位(SPVNR)
    HWFNRKN     As String * 1       ' 検査頻度＿抜(SPVNR)
    HWFNRPUG    As Double           ' PUA限(SPVNR)
    HWFNRPUR    As Double           ' PUA率(SPVNR)      '追加　06/05/31 ooba END =========>
End Type
Public tbl_chk1_5(1) As typ_chk1_5
Public tbl_chk1_5_SXGD As typ_chk1_5    '結晶GD仕様格納用　05/07/29 ooba

Public Type typ_chk1_5_1
    HOSYOU      As String * 1       '保証方法＿対象
    Min         As Integer          '下限
    max         As Integer          '上限
    SOKU_HOU    As String * 1       '測定位置＿方
    SOKU_TEN    As String * 1       '測定位置＿点
    SOKU_ICHI   As String * 1       '測定位置＿位
    SOKU_RYOU   As String * 1       '測定位置＿領
    UMU         As String * 1       '検査有無           ????????????????(桁数）
    NETSU       As String * 2       '熱処理法
    JOUKEN      As String * 1       '測定条件
    ET          As Integer          '選択ＥＴ代
    KENSA       As String * 2       '検査方法
    PATTERN     As String * 1       'パターン区分
    KENH_NUKI   As String * 1       '検査頻度_抜　04/04/13 ooba
'*** UPDATE ↓ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数追加
    LINE        As String           ' 測定条件(GDﾗｲﾝ数)
'*** UPDATE ↑ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数追加
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    HWFGDSZY    As String * 1       'GD測定条件
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
'↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.1 AN温度 振替可否チェック追加
    HWFANTNP    As Integer          ' 品ＷＦＡＮ温度
'↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    PUAGEN      As Double           'PUA限          '追加　06/05/31 ooba
    PUAPER      As Double           'PUA率          '追加　06/05/31 ooba
    
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    HWFSIRDMX   As Integer          '軸状転位上限(SIRD)
    HWFSIRDHT   As String * 1       '軸状転位保証方法＿対(SIRD)
    HWFSIRDKM   As String * 1       '軸状転位検査頻度＿枚(SIRD)
    HWFSIRDKH   As String * 1       '軸状転位検査頻度＿保(SIRD)
    HWFSIRDKU   As String * 1       '軸状転位検査頻度＿ウ(SIRD)
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
    
End Type
Public tbl_chk1_5_1(1) As typ_chk1_5_1

Public Type typ_chk1_6
    HWFNP1AR    As Double           '品WFナノトポ１エリア
    HWFNP1MAX   As Double           '品WFナノトポ１上限
    HWFNP2AR    As Double           '品WFナノトポ２エリア
    HWFNP2MAX   As Double           '品WFナノトポ２上限
    HSXCSCEN    As Double           '結晶面傾き中心
End Type
Public tbl_chk1_6(1) As typ_chk1_6

'品番組合せﾁｪｯｸ1　06/04/25 ooba
Public Type typ_chk1_7
    HSXTYPE     As String * 1       '品SXタイプ
    HSXCDIR     As String * 1       '品SX結晶面方位
    HSXCSCEN    As Double           '品SX結晶面傾中心
    HSXDOP      As String * 1       '品SXドーパント
    HWFCDOP     As String * 1       '品WF結晶ドープ
    HSXSDSLP    As String * 1       '品SXシード傾
    HSXDPDIR    As String * 2       '品SX溝位置方位
End Type
Public tbl_chk1_7(1) As typ_chk1_7

'品番組合せﾁｪｯｸ2　06/04/25 ooba
Public Type typ_chk1_8
    HSXCDOP     As String * 1       '品SX結晶ドープ
    GLASS       As String * 1       'ガラス接着
    SLICEATU    As Double           'SL厚み
    HSXCSMIN    As Double           '品SX結晶面傾下限
    HSXCSMAX    As Double           '品SX結晶面傾上限
    HSXWFWAR    As String * 1       '品SXWFWarpランク
    KUMIDOP    As String * 1       '組合せドープフラグ 2006/07/21 SMP)kondoh Add
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    HSXDKTMP    As String * 1       '品SXDK温度
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
End Type
Public tbl_chk1_8(1) As typ_chk1_8

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
Public Type typ_chk1_9
    HEPOF1HS    As String * 1       '保証方法＿処
    HEPOF1SH    As String * 1       '測定位置＿方
    HEPOF1ST    As String * 1       '測定位置＿点
    HEPOF1SR    As String * 1       '測定位置＿領
    HEPOF1NS    As String * 2       '熱処理法
    HEPOF1SZ    As String * 1       '測定条件
    HEPOF1ET    As Integer          '選択ＥＴ代
    HEPOSF1PTK  As String * 1       'パターン区分
    HEPOF1KN    As String * 1       '検査頻度＿抜
    HEPOF2HS    As String * 1       '保証方法＿処
    HEPOF2SH    As String * 1       '測定位置＿方
    HEPOF2ST    As String * 1       '測定位置＿点
    HEPOF2SR    As String * 1       '測定位置＿領
    HEPOF2NS    As String * 2       '熱処理法
    HEPOF2SZ    As String * 1       '測定条件
    HEPOF2ET    As Integer          '選択ＥＴ代
    HEPOSF2PTK  As String * 1       'パターン区分
    HEPOF2KN    As String * 1       '検査頻度＿抜
    HEPOF3HS    As String * 1       '保証方法＿処
    HEPOF3SH    As String * 1       '測定位置＿方
    HEPOF3ST    As String * 1       '測定位置＿点
    HEPOF3SR    As String * 1       '測定位置＿領
    HEPOF3NS    As String * 2       '熱処理法
    HEPOF3SZ    As String * 1       '測定条件
    HEPOF3ET    As Integer          '選択ＥＴ代
    HEPOSF3PTK  As String * 1       'パターン区分
    HEPOF3KN    As String * 1       '検査頻度＿抜
    HEPBM1HS    As String * 1       '保証方法＿処
    HEPBM1SH    As String * 1       '測定位置＿方
    HEPBM1ST    As String * 1       '測定位置＿点
    HEPBM1SR    As String * 1       '測定位置＿領
    HEPBM1NS    As String * 2       '熱処理法
    HEPBM1SZ    As String * 1       '測定条件
    HEPBM1ET    As Integer          '選択ＥＴ代
    HEPBM1KN    As String * 1       '検査頻度＿抜
    HEPBM2HS    As String * 1       '保証方法＿処
    HEPBM2SH    As String * 1       '測定位置＿方
    HEPBM2ST    As String * 1       '測定位置＿点
    HEPBM2SR    As String * 1       '測定位置＿領
    HEPBM2NS    As String * 2       '熱処理法
    HEPBM2SZ    As String * 1       '測定条件
    HEPBM2ET    As Integer          '選択ＥＴ代
    HEPBM2KN    As String * 1       '検査頻度＿抜
    HEPBM3HS    As String * 1       '保証方法＿処
    HEPBM3SH    As String * 1       '測定位置＿方
    HEPBM3ST    As String * 1       '測定位置＿点
    HEPBM3SR    As String * 1       '測定位置＿領
    HEPBM3NS    As String * 2       '熱処理法
    HEPBM3SZ    As String * 1       '測定条件
    HEPBM3ET    As Integer          '選択ＥＴ代
    HEPBM3KN    As String * 1       '検査頻度＿抜
    HEPANTNP    As Integer          '品ＥＰＡＮ温度
    HEPACEN     As Double           '品ＥＰ厚中心
End Type
Public tbl_chk1_9(1) As typ_chk1_9
Public Type typ_chk1_9_1
    HOSYOU      As String * 1       '保証方法＿処
    MIN_LIMIT   As Integer          '下限
    MAX_LIMIT   As Integer          '上限
    SOKU_HOU    As String * 1       '測定位置＿方
    SOKU_TEN    As String * 1       '測定位置＿点
    SOKU_ICHI   As String * 1       '測定位置＿位
    SOKU_RYOU   As String * 1       '測定位置＿領
    UMU         As String * 1       '検査有無
    NETSU       As String * 2       '熱処理法
    JOUKEN      As String * 1       '測定条件
    ET          As Integer          '選択ＥＴ代
    KENSA       As String * 2       '検査方法
    PATTERN     As String * 1       'パターン区分
    KENH_NUKI   As String * 1       '検査頻度＿抜
    ANTMP       As String           'AN温度
    EPATU       As Double           'エピ厚
End Type
Public tbl_chk1_9_1(1) As typ_chk1_9_1
Public RET_3_4  As Integer          '3-4 ＷＦＣ評価実績(エピ)チェック(funChkFurikae3_4)の戻り値
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

'常識仕様ﾁｪｯｸ2　06/10/05 ooba
Public Type typ_chk1_10
    HSXCDIR     As String * 1       '結晶面方位
    HSXCSCEN    As Double           '結晶面傾き中心 ''2008/11/27 結晶面傾中心チェック緩和(2) ADD By Systech
    HSXDOP      As String * 1       'ドーパント
    HWFCDOP     As String * 1       '結晶ドープ
    HSXDPDIR    As String * 2       '溝位置方位
    MCNO1       As String * 1       '品種
    MCNO2       As String * 1       '引上げ速度
    MCNO3       As String * 1       'HZタイプ
    DCHYUUBU    As String * 1       'ドローチューブ
    NDOPHUFLG   As String * 1       '窒素ドープ振替可能チェック add 0108
    CDOPHUFLG   As String * 1       'Cドープ振替可能チェック    add 0108
End Type
Public tbl_chk1_10(1) As typ_chk1_10

'品番組合せﾁｪｯｸ3　11/04/14 kameda
Public Type typ_chk1_11
    hinban As tFullHinban
    NHINCHKFLG  As String           '狙い品番チェックフラグ
End Type

Public tbl_chk1_11(1) As typ_chk1_11

'Add Start 2011/04/20 SMPK Miyata
'中間抜試仕様チェック
Public Type typ_chk1_12
    MSMPFLG     As String * 1       '中間抜試フラグ
    MSMPTANIMAI As Integer          '中間抜試単位(枚数)
End Type
Public tbl_chk1_12(1) As typ_chk1_12
'Add End   2011/04/20 SMPK Miyata

'マルチ引上げ適用可否チェック　11/05/19 kameda
Public Type typ_chk1_13
    hinban As tFullHinban
    MLTHTFLG  As String           'マルチ引上げ適用可否フラグ
    SIJICNT As Integer             'グループ指示数
    RENBAN As Integer          'グループ内引上順番
    MLTJDG() As String
End Type

Public tbl_chk1_13(1) As typ_chk1_13
Public tbl_chk2_5 As typ_chk1_13

'Add Start 2011/05/11 SMPK Nakamura FRSシステム化対応
'FRS仕様ﾁｪｯｸ
Public Type typ_chk1_14
    FRSFLG  As String               'FRS測定フラグ
End Type
'Add End 2011/05/11 SMPK Nakamura FRSシステム化対応

Public tbl_chk1_14(1) As typ_chk1_14

'Add Start 2011/07/22 SMPK Nakamura 結晶面傾きチェック追加対応
'結晶面傾きチェック
Public Type typ_chk1_15
    HSXCDIR     As String * 1       'ＳＸＬ結晶面方位
    HSXCSCEN    As Double           'ＳＸＬ結晶面傾き中心
    HSXCSMIN    As Double           'ＳＸＬ結晶面傾き下限
    HSXCSMAX    As Double           'ＳＸＬ結晶面傾き上限
    HSXCKWAY    As String * 2       'ＳＸＬ結晶面検査方法
    HSXCKHNM    As String * 1       'ＳＸＬ結晶面検査頻度_枚
    HSXCKHNI    As String * 1       'ＳＸＬ結晶面検査頻度_位
    HSXCKHNH    As String * 1       'ＳＸＬ結晶面検査頻度_保
    HSXCKHNS    As String * 1       'ＳＸＬ結晶面検査頻度_試
    HSXCSDIR    As String * 2       'ＳＸＬ結晶面傾き方位
    HSXCSDIS    As String * 1       'ＳＸＬ結晶面傾き方位指定
    HSXCTDIR    As String * 2       'ＳＸＬ結晶面傾き縦方位
    HSXCTCEN    As Double           'ＳＸＬ結晶面傾き縦中心
    HSXCTMIN    As Double           'ＳＸＬ結晶面傾き縦下限
    HSXCTMAX    As Double           'ＳＸＬ結晶面傾き縦上限
    HSXCYDIR    As String * 2       'ＳＸＬ結晶面傾き横方位
    HSXCYCEN    As Double           'ＳＸＬ結晶面傾き横中心
    HSXCYMIN    As Double           'ＳＸＬ結晶面傾き横下限
    HSXCYMAX    As Double           'ＳＸＬ結晶面傾き横上限
    HWFCSGCEN   As Double           'ＷＦ結晶面操合成角中心
    HWFCSGMIN   As Double           'ＷＦ結晶面操合成角下限
    HWFCSGMAX   As Double           'ＷＦ結晶面操合成角上限
    HWFCSXCEN   As Double           'ＷＦ結晶面操Ｘ方位中心
    HWFCSXMIN   As Double           'ＷＦ結晶面操Ｘ方位下限
    HWFCSXMAX   As Double           'ＷＦ結晶面操Ｘ方位上限
    HWFCSYCEN   As Double           'ＷＦ結晶面操Ｙ方位中心
    HWFCSYMIN   As Double           'ＷＦ結晶面操Ｙ方位下限
    HWFCSYMAX   As Double           'ＷＦ結晶面操Ｙ方位上限
End Type
Public tbl_chk1_15(1) As typ_chk1_15

'結晶面傾き組合せチェック
Public Type typ_chk1_16
    HSXCDIR     As String * 1       'ＳＸＬ結晶面方位
    HSXCSCEN    As Double           'ＳＸＬ結晶面傾き中心
    HSXCSMIN    As Double           'ＳＸＬ結晶面傾き下限
    HSXCSMAX    As Double           'ＳＸＬ結晶面傾き上限
    HSXCKWAY    As String * 2       'ＳＸＬ結晶面検査方法
    HSXCKHNM    As String * 1       'ＳＸＬ結晶面検査頻度_枚
    HSXCKHNI    As String * 1       'ＳＸＬ結晶面検査頻度_位
    HSXCKHNH    As String * 1       'ＳＸＬ結晶面検査頻度_保
    HSXCKHNS    As String * 1       'ＳＸＬ結晶面検査頻度_試
    HSXCSDIR    As String * 2       'ＳＸＬ結晶面傾き方位
    HSXCSDIS    As String * 1       'ＳＸＬ結晶面傾き方位指定
    HSXCTDIR    As String * 2       'ＳＸＬ結晶面傾き縦方位
    HSXCTCEN    As Double           'ＳＸＬ結晶面傾き縦中心
    HSXCTMIN    As Double           'ＳＸＬ結晶面傾き縦下限
    HSXCTMAX    As Double           'ＳＸＬ結晶面傾き縦上限
    HSXCYDIR    As String * 2       'ＳＸＬ結晶面傾き横方位
    HSXCYCEN    As Double           'ＳＸＬ結晶面傾き横中心
    HSXCYMIN    As Double           'ＳＸＬ結晶面傾き横下限
    HSXCYMAX    As Double           'ＳＸＬ結晶面傾き横上限
    HWFCSGCEN   As Double           'ＷＦ結晶面操合成角中心
    HWFCSGMIN   As Double           'ＷＦ結晶面操合成角下限
    HWFCSGMAX   As Double           'ＷＦ結晶面操合成角上限
    HWFCSXCEN   As Double           'ＷＦ結晶面操Ｘ方位中心
    HWFCSXMIN   As Double           'ＷＦ結晶面操Ｘ方位下限
    HWFCSXMAX   As Double           'ＷＦ結晶面操Ｘ方位上限
    HWFCSYCEN   As Double           'ＷＦ結晶面操Ｙ方位中心
    HWFCSYMIN   As Double           'ＷＦ結晶面操Ｙ方位下限
    HWFCSYMAX   As Double           'ＷＦ結晶面操Ｙ方位上限
End Type
Public tbl_chk1_16(1) As typ_chk1_16
'Add End 2011/07/22 SMPK Nakamura 結晶面傾きチェック追加対応

'合否判定有無フラグ（0:反映データの合否判定を行う, 1:反映データの合否判定を行わない） '2005/02/08 ffc)tanabe
Public JudgChgFlg           As String
'SB_Com2→SB_Comに移動 08/12/24 ooba
'Public JudgKoutei           As String       '工程(結晶実績ﾁｪｯｸ用)　08/04/15 ooba

'--------------- 2008/07/25 INSERT START  By Systech ---------------
'SB_Com2→SB_Comに移動 08/12/24 ooba
'Public gsTbcmy028ErrCode    As String           ' 振替チェックエラーコード
'--------------- 2008/07/25 INSERT  END   By Systech ---------------

'認定炉判定用   2008/08/20  Info.Kameda  -------
Public Type type_Ninteiro_xodfa
    ROID        As String               ' 認定炉ID
    REV         As Integer              ' 改訂番号
    GOUKI       As String               ' 号機ID
    KUBUN       As String               ' 区分
    FRCHG       As Long                 ' ﾁｬｰｼﾞ量From
    TOCHG       As Long                 ' ﾁｬｰｼﾞ量To
    JUDGRO      As String               ' 判定
    SUICHG      As Long                 ' 推定ﾁｬｰｼﾞ
    CHKSXL      As String               ' 判定有無(SXL)
    CHKWFC1     As String               ' 判定有無(WF)
    CHKWFC2     As String               ' 判定有無(WF)
    SYNDAY      As String
End Type
'認定炉情報表示用
Public gNinteiro_Data() As type_Ninteiro_xodfa
'------------------------------------------------
'窒素濃度判定用  2009/07/30  Kameda
Public Type typ_chk2_4
    N2NOUDO As Double
    NJDG() As String
End Type
Public tbl_chk2_4(1) As typ_chk2_4
'Sub MAIN()
'    Dim ret As Integer
'    Dim ErrCode As Integer
'    Dim ErrMsg As String
'    Dim iErr_Code As Integer
'    Dim sErr_Msg As String
'    Dim iCrySmpID As Integer
'    Dim sWfSmpID As String
'    Dim iJudgFlg As Integer
'    gsFactryCd = "42"
'    OraDBOpen
'SEKI:
'    tOld_Hinban.hinban = "SZS0014A"
'    tOld_Hinban.mnorevno = 0
'    tOld_Hinban.factory = "Y"
'    tOld_Hinban.opecond = "1"
'    tNew_Hinban.hinban = "ZZS0014A"
'    tNew_Hinban.mnorevno = 0
'    tNew_Hinban.factory = "Y"
'    tNew_Hinban.opecond = "1"
'    ret = funChkFurikaeShiyou("CC600", "716302010000", tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg, 1)
'    Debug.Print Time, iErr_Code, sErr_Msg
'    GoTo SEKI
'End Sub
    
'------------------------------------------------
' 振替可否チェック（仕様）
'------------------------------------------------

'概要      :パラメータに指定された、振替元品番から振替先品番に振り替えが可能かどうかをチェックし、結果を返す。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型             :説明
'          :sProccd         ,I  ,String         :工程番号
'          :sKeyID          ,I  ,String         :ﾌﾞﾛｯｸID、又は、SXL-ID
'          :tOld_Hinban     ,I  ,tFullHinban    :振替元品番(構造体)
'          :tNew_Hinban     ,I  ,tFullHinban    :振替先品番(構造体)
'          :iErr_Code       ,O  ,Integer        :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String         :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :typ_B           ,O  ,typ_AllTypesB  :結晶総合判定全情報構造体(構造体)
'          :typ_CType       ,O  ,typ_AllTypesC  :WFC総合判定全情報構造体(構造体)
'          :iSmpGetFlg      ,I  ,Integer        :ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ(0:ｻﾝﾌﾟﾙ指定なし, 1:ｻﾝﾌﾟﾙ指定あり)
'          :sSamplID1       ,I  ,String         :TOPｻﾝﾌﾟﾙID(省略可)
'          :sSamplID2       ,I  ,String         :BOTｻﾝﾌﾟﾙID(省略可)
'          :iKcnt           ,I  ,Integer        :工程連番(省略可)
'          :iHcnt           ,I  ,Integer        :複数品番カウント(認定炉・窒素の複数品番判定用に追加)  2009/09/25 Kameda
'          :iCC10           ,I  ,Integer        :結晶設計変更工程フラグ  1:結晶設計変更工程            2011/07/11 Kameda
'          :sPlshMeth       ,I  ,String         :研削方法(加工払出画面) M:MGR(MGRの場合、溝位置方位ロック解除)  2011/10/13 SETsw kubota
'          :戻り値          ,O  ,Integer        :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/19 新規作成　SB
''>>>>>複数品番20060502 SMP桜井======================================================
''Public Function funChkFurikaeShiyou(sProccd As String, sKeyID As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
''                                    iErr_Code As Integer, sErr_Msg As String, _
''                                    typ_B As typ_AllTypesB, typ_CType As typ_AllTypesC, _
''                                    iSmpGetFlg As Integer, _
''                                    Optional sSamplID1 As String = vbNullString, Optional sSamplID2 As String = vbNullString, _
''                                    Optional iKcnt As Integer = 0) As Integer
''<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Public Function funChkFurikaeShiyou(sProccd As String, sKeyID As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                    iErr_Code As Integer, sErr_Msg As String, _
                                    typ_b As typ_AllTypesB, typ_CType As typ_AllTypesC, _
                                    iSmpGetFlg As Integer, _
                                    Optional sSamplID1 As String = vbNullString, Optional sSamplID2 As String = vbNullString, _
                                    Optional iKcnt As Integer = 0, _
                                    Optional iMultiFlg As Integer = 0, _
                                    Optional iELCs_Flg As Integer = 0, _
                                    Optional iHcnt As Integer = 1, _
                                    Optional iCC10 As Integer = 0 _
                                  , Optional sPlshMeth As String = "" _
                                  ) As Integer
    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sResult2    As String       'コードＤＢ取得関数(FE)の取得変数 2011/04/07追加 SETsw kubota
    Dim sResult3    As String       'コードＤＢ取得関数(FF)の取得変数 2011/04/07追加 SETsw kubota
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikaeShiyou = 0
    iErr_Code = 0
    sErr_Msg = ""
    RET_3_4 = 0                     '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
   ' tJudgData = ""
    '------------------------------------------ Ｇ品番、Ｚ品番チェック -------------------------------------------------
    '通常品番 → Ｇ品番、Ｚ品番　⇒　振替ＯＫ
    If (Trim$(tNew_Hinban.hinban) = "G") Or (Trim$(tNew_Hinban.hinban) = "Z") Then GoTo Apl_Exit
    
    'Ｚ品番、Ｇ品番 → 通常品番　⇒　振替ＯＫ
    If (Trim$(tOld_Hinban.hinban) = "Z") Or (Trim$(tOld_Hinban.hinban) = "G") Then GoTo Apl_Exit
    
''    'Ｇ品番 → 通常品番　⇒　振替ＮＧ
''    If (Trim$(tOld_Hinban.HINBAN) = "G") Then
''        funChkFurikaeShiyou = 1
''        iErr_Code = 1100
''        sErr_Msg = "G品番は振替できません。"
''        GoTo Apl_Exit
''    End If
    '------------------------------------------ 入力チェック -------------------------------------------------
    '工程番号のチェック
    If Trim$(sProccd) = "" Then
            funChkFurikaeShiyou = -1
            sErr_Msg = "入力引数値エラー(工程番号指定なし)"
            GoTo Apl_Error
    End If
    JudgKoutei = sProccd        '08/04/15 ooba
    'ﾌﾞﾛｯｸID、SXL-IDのチェック
    If Trim$(sKeyID) = "" Then
            funChkFurikaeShiyou = -1
            sErr_Msg = "入力引数値エラー(ﾌﾞﾛｯｸID or SXL-ID指定なし, 工程番号 : " & sProccd & ")"
            GoTo Apl_Error
    End If
    If (left(sProccd, 4) = "CC31") Or (left(sProccd, 4) = "CC60") Or (left(sProccd, 4) = "CC61") Or (left(sProccd, 4) = "CC73") Or _
       (left(sProccd, 4) = "CW74") Or (left(sProccd, 4) = "CW75") Or (left(sProccd, 4) = "CW76") Then
        If (left(sProccd, 4) = "CW75") Or (left(sProccd, 4) = "CW76") Then
            If Len(sKeyID) <> 13 Then
                funChkFurikaeShiyou = -1
                sErr_Msg = "入力引数値エラー(SXL-ID : " & sKeyID & ")"
                GoTo Apl_Error
            End If
        Else
            If Len(sKeyID) <> 12 Then
                funChkFurikaeShiyou = -1
                sErr_Msg = "入力引数値エラー(ﾌﾞﾛｯｸID : " & sKeyID & ")"
                GoTo Apl_Error
            End If
        End If
    Else
            funChkFurikaeShiyou = -1
            sErr_Msg = "入力引数値エラー(工程番号 : " & sProccd & ")"
            GoTo Apl_Error
    End If
    'ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞの指定チェック
    If (left(sProccd, 4) = "CC60") Or (left(sProccd, 4) = "CW75") Then
       If (IsNull(iSmpGetFlg)) Or (iSmpGetFlg <> 0 And iSmpGetFlg <> 1) Then
          funChkFurikaeShiyou = -1
          sErr_Msg = "入力引数値エラー(ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ : " & iSmpGetFlg & ")"
          GoTo Apl_Error
       End If
       If (iSmpGetFlg = 1) And _
          (IsNull(sSamplID1) Or Trim$(sSamplID1) = "" Or IsNull(sSamplID2) Or Trim$(sSamplID2) = "") Then
          funChkFurikaeShiyou = -1
          sErr_Msg = "入力引数値エラー(ｻﾝﾌﾟﾙID指定なし)"
          GoTo Apl_Error
       End If
    End If
    
    '------------------------------------------ 指示取得 ------------------------------------------------------
    '振替指示データ取得
    sResult = ""
'    RET = funCodeDBGet("SB", "FC", sProccd, 0, " ", sResult)
    RET = funCodeDBGet("SB", "FD", sProccd, 0, " ", sResult)        'FC→FD 2011/04/07修正 SETsw kubota
    If RET <> 0 Then
        funChkFurikaeShiyou = -2
        GoTo Apl_Error
    End If
    
    '振替指示データ取得(FE) 2011/04/07追加 SETsw kubota
    sResult2 = ""
    RET = funCodeDBGet("SB", "FE", sProccd, 0, " ", sResult2)
    If RET <> 0 Then
        funChkFurikaeShiyou = -2
        GoTo Apl_Error
    End If
    
    '振替指示データ取得(FF) 2011/04/07追加 SETsw kubota
    sResult3 = ""
    RET = funCodeDBGet("SB", "FF", sProccd, 0, " ", sResult3)
    If RET <> 0 Then
        funChkFurikaeShiyou = -2
        GoTo Apl_Error
    End If
    '------------------------------------------ Make SQL ------------------------------------------------------
    '1-1 組み合わせ品番チェック
    If Mid(sResult, 1, 1) = "1" Then
        If iMultiFlg = 0 Then ''<<複数品番判定対応20060502 SMP桜井　ブロック品番保障0,2のMultiBlockの真ん中はしない
            If iELCs_Flg = 0 Or iELCs_Flg = 2 Then ''<<複数品番判定対応20060502 SMP桜井　振替元品番チェック振替元以外の品番はしない
                RET = funChkFurikae1_1(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg)
                If RET <> 0 Then
                    funChkFurikaeShiyou = RET
                    If RET > 0 Then GoTo Apl_Exit
                    GoTo Apl_Error
                End If
            End If
        End If
    End If
    '1-2 常識仕様チェック
    If Mid(sResult, 2, 1) = "1" Then
        If iMultiFlg = 0 Then ''<<複数品番判定対応20060502 SMP桜井　ブロック品番保障0,2のMultiBlockの真ん中はしない
            If iELCs_Flg = 0 Or iELCs_Flg = 2 Then  ''<<複数品番判定対応20060502 SMP桜井　振替元品番チェック振替元以外の品番はしない
'                RET = funChkFurikae1_2(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg)
                RET = funChkFurikae1_2(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg, sPlshMeth)
                If RET <> 0 Then
                    funChkFurikaeShiyou = RET
                    If RET > 0 Then GoTo Apl_Exit
                    GoTo Apl_Error
                End If
            End If
        End If
    End If
    '1-3 外観実績を振替先品番チェック
    If Mid(sResult, 3, 1) = "1" Then
        If iMultiFlg = 0 Then ''<<複数品番判定対応20060502 SMP桜井　ブロック品番保障0,2のMultiBlockの真ん中はしない
            If iELCs_Flg = 0 Or iELCs_Flg = 2 Then  ''<<複数品番判定対応20060502 SMP桜井　振替元品番チェック振替元以外の品番はしない
                RET = funChkFurikae1_3(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg)
                If RET <> 0 Then
                    funChkFurikaeShiyou = RET
                    If RET > 0 Then GoTo Apl_Exit
                    GoTo Apl_Error
                End If
            End If
        End If
    End If
    '1-4 結晶評価項目仕様チェック
    If Mid(sResult, 4, 1) = "1" Then
        If iMultiFlg = 0 Then ''<<複数品番判定対応20060502 SMP桜井　ブロック品番保障0,2のMultiBlockの真ん中はしない
            RET = funChkFurikae1_4(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg, iELCs_Flg)
            If RET <> 0 Then
                funChkFurikaeShiyou = RET
                If RET > 0 Then GoTo Apl_Exit
                GoTo Apl_Error
            End If
        End If
    End If
    '1-5 先行評価項目仕様チェック
    If Mid(sResult, 5, 1) = "1" Then
        If iMultiFlg = 0 Then ''<<複数品番判定対応20060502 SMP桜井　ブロック品番保障0,2のMultiBlockの真ん中はしない
            If iELCs_Flg = 0 Or iELCs_Flg = 2 Then  ''<<複数品番判定対応20060502 SMP桜井　振替元品番チェック振替元以外の品番はしない
                RET = funChkFurikae1_5(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg)
                If RET <> 0 Then
                    funChkFurikaeShiyou = RET
                    If RET > 0 Then GoTo Apl_Exit
                    GoTo Apl_Error
                End If
            End If
        End If
    End If
    '1-6 ナノトポ規格チェック
    If Mid(sResult, 6, 1) = "1" Then
        If iMultiFlg = 0 Then ''<<複数品番判定対応20060502 SMP桜井　ブロック品番保障0,2のMultiBlockの真ん中はしない
            If iELCs_Flg = 0 Or iELCs_Flg = 2 Then  ''<<複数品番判定対応20060502 SMP桜井　振替元品番チェック振替元以外の品番はしない
                RET = funChkFurikae1_6(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg)
                If RET <> 0 Then
                    funChkFurikaeShiyou = RET
                    If RET > 0 Then GoTo Apl_Exit
                    GoTo Apl_Error
                End If
            End If
        End If
    End If
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    '1-9 先行評価項目仕様チェック
    If Mid(sResult, 9, 1) = "1" Then
        If iMultiFlg = 0 Then ''ブロック品番保障0,2のMultiBlockの真ん中はしない
            If iELCs_Flg = 0 Or iELCs_Flg = 2 Then  ''振替元品番チェック振替元以外の品番はしない
                RET = funChkFurikae1_9(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg)
                If RET <> 0 Then
                    funChkFurikaeShiyou = RET
                    If RET > 0 Then GoTo Apl_Exit
                    GoTo Apl_Error
                End If
            End If
        End If
    End If
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    '1-10 常識仕様チェック２　06/10/05 ooba
    If Mid(sResult, 10, 1) = "1" Then
        If iMultiFlg = 0 Then
            If iELCs_Flg = 0 Or iELCs_Flg = 2 Then
                RET = funChkFurikae1_10(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg)
                If RET <> 0 Then
                    funChkFurikaeShiyou = RET
                    If RET > 0 Then GoTo Apl_Exit
                    GoTo Apl_Error
                End If
            End If
        End If
    End If
    '1-11 狙い品番チェック　11/04/14 kameda
    If Mid(sResult, 11, 1) = "1" Then
        If iMultiFlg = 0 Then
            If iELCs_Flg = 0 Or iELCs_Flg = 2 Then
                'RET = funChkFurikae1_11(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg)
                RET = funChkFurikae1_11(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iCC10, iErr_Code, sErr_Msg)
                If RET <> 0 Then
                    funChkFurikaeShiyou = RET
                    If RET > 0 Then GoTo Apl_Exit
                    GoTo Apl_Error
                End If
            End If
        End If
    End If
    'Add Start 2011/04/20 SMPK Miyata
    '1-12 中間抜試仕様チェック
    If Mid(sResult, 12, 1) = "1" Then
        RET = funChkFurikae1_12(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg)
        If RET <> 0 Then
            funChkFurikaeShiyou = RET
            If RET > 0 Then GoTo Apl_Exit
            GoTo Apl_Error
        End If
    End If
    'Add End   2011/04/20 SMPK Miyata
    '1-13 マルチ引上げ適用可否チェック　11/05/19 kameda
    If Mid(sResult, 13, 1) = "1" Then
        If iMultiFlg = 0 Then
            If iELCs_Flg = 0 Or iELCs_Flg = 2 Then
                RET = funChkFurikae1_13(sKeyID, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg)
                If RET <> 0 Then
                    funChkFurikaeShiyou = RET
                    If RET > 0 Then GoTo Apl_Exit
                    GoTo Apl_Error
                End If
            End If
        End If
    End If
'Add Start 2011/05/11 SMPK Nakamura FRSシステム化対応
    '1-14 FRS仕様チェック
    If Mid(sResult, 14, 1) = "1" Then
        If iMultiFlg = 0 Then
            If iELCs_Flg = 0 Or iELCs_Flg = 2 Then
                RET = funChkFurikae1_14(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg)
                If RET <> 0 Then
                    funChkFurikaeShiyou = RET
                    If RET > 0 Then GoTo Apl_Exit
                    GoTo Apl_Error
                End If
            End If
        End If
    End If
'Add End 2011/05/11 SMPK Nakamura FRSシステム化対応
'Add Start 2011/07/12 SMPK Nakamura 結晶面傾きチェック追加対応
    '1-15 結晶面傾きチェック
    If Mid(sResult, 15, 1) = "1" Then
        If iMultiFlg = 0 Then
            If iELCs_Flg = 0 Or iELCs_Flg = 2 Then
                RET = funChkFurikae1_15(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg)
                If RET <> 0 Then
                    funChkFurikaeShiyou = RET
                    If RET > 0 Then GoTo Apl_Exit
                    GoTo Apl_Error
                End If
            End If
        End If
    End If
'Add End 2011/07/12 SMPK Nakamura 結晶面傾きチェック追加対応
    '2-1 結晶評価実績チェック
'    If Mid(sResult, 11, 1) = "1" Then
    If Mid(sResult2, 1, 1) = "1" Then       'FCの11桁目→FEの1桁目 2011/04/07 SETsw kubota
        If iELCs_Flg = 0 Or iELCs_Flg = 2 Then  ''<<複数品番判定対応20060502 SMP桜井　振替元品番チェック振替元以外の品番はしない
            RET = funChkFurikae2_1(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg, iSmpGetFlg, sSamplID1, sSamplID2, iKcnt)
            If RET <> 0 Then
                funChkFurikaeShiyou = RET
                '判定(CC600)は抜けない  2008/08/28 修正
                If sProccd = "CC600" Then
                    If RET < 0 Then GoTo Apl_Error
                Else
                    If RET > 0 Then GoTo Apl_Exit
                    GoTo Apl_Error
                End If
            End If
        End If
    End If
    
    ''C－OSF3チェックの変更 2008.04.20 青柳
    '2-2 C－OSF3チェック
'    If Mid(sResult, 12, 1) = "1" Then
    If Mid(sResult2, 2, 1) = "1" Then       'FCの12桁目→FEの2桁目 2011/04/07 SETsw kubota
        If iMultiFlg = 0 Then ''<<複数品番判定対応20060502 SMP桜井　ブロック品番保障0,2のMultiBlockの真ん中はしない
            RET = funChkFurikae2_2(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg, iELCs_Flg)
            If RET <> 0 Then
                funChkFurikaeShiyou = RET
                If RET > 0 Then GoTo Apl_Exit
                GoTo Apl_Error
            End If
        End If
    End If
    
    '2-3 認定炉チェック
'    If Mid(sResult, 13, 1) = "1" Then
    If Mid(sResult2, 3, 1) = "1" Then       'FCの13桁目→FEの3桁目 2011/04/07 SETsw kubota
        'If iMultiFlg = 0 Then     'del 2010/05/07 Kameda
            RET = funChkFurikae2_3(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iHcnt, iErr_Code, sErr_Msg)
            If RET <> 0 Then
                funChkFurikaeShiyou = RET
                '判定(CC600,CW750)は抜けない 2008/08/28 修正
                If sProccd = "CC600" Or sProccd = "CW750" Then
                    If RET < 0 Then GoTo Apl_Error
                Else
                    If RET > 0 Then GoTo Apl_Exit
                    GoTo Apl_Error
                End If
            End If
        'End If
    End If
    
    '2-4 窒素濃度チェック add 2009/07/30 Kameda
'    If Mid(sResult, 14, 1) = "1" Then
    If Mid(sResult2, 4, 1) = "1" Then       'FCの14桁目→FEの4桁目 2011/04/07 SETsw kubota
        'If iMultiFlg = 0 Then     'del 2010/05/07 Kameda
            RET = funChkFurikae2_4(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iHcnt, iErr_Code, sErr_Msg)
            If RET <> 0 Then
                funChkFurikaeShiyou = RET
                '判定(CC600)は抜けない
                If sProccd = "CC600" Then
                    If RET < 0 Then GoTo Apl_Error
                Else
                    If RET > 0 Then GoTo Apl_Exit
                    GoTo Apl_Error
                End If
            End If
        'End If
    End If
    
    '2-5 マルチ引上げ適用チェック add 2011/05/19 Kameda
    If Mid(sResult2, 5, 1) = "1" Then       'FEの5桁目
        'If iMultiFlg = 0 Then     'del 2010/05/07 Kameda
            RET = funChkFurikae2_5(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iHcnt, iErr_Code, sErr_Msg)
            If RET <> 0 Then
                funChkFurikaeShiyou = RET
                '判定(CC600)は抜けない
                If sProccd = "CC600" Or sProccd = "CW750" Then
                    If RET < 0 Then GoTo Apl_Error
                Else
                    If RET > 0 Then GoTo Apl_Exit
                    GoTo Apl_Error
                End If
            End If
        'End If
    End If
    
    '3-1 ＷＦＣ評価実績チェック
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -s-
    RET = 0
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -e-
'    If Mid(sResult, 21, 1) = "1" Then
    If Mid(sResult3, 1, 1) = "1" Then       'FCの21桁目→FFの1桁目 2011/04/07 SETsw kubota
        If iMultiFlg = 0 Then ''<<複数品番判定対応20060502 SMP桜井　ブロック品番保障0,2のMultiBlockの真ん中はしない
            If iELCs_Flg = 0 Or iELCs_Flg = 2 Then  ''<<複数品番判定対応20060502 SMP桜井　振替元品番チェック振替元以外の品番はしない
                RET = funChkFurikae3_1(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg, iSmpGetFlg, sSamplID1, sSamplID2, iKcnt)
                If RET <> 0 Then
                    funChkFurikaeShiyou = RET
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -s-
'' チェック結果がNGの場合でも、3-4 ＷＦＣ評価実績(エピ)チェックは行う
'                    If RET > 0 Then GoTo Apl_Exit
'                    GoTo Apl_Error
                    If RET < 0 Then GoTo Apl_Error
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -e-
                End If
            End If
        End If
    End If
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    '3-4 ＷＦＣ評価実績(エピ)チェック
'    If Mid(sResult, 24, 1) = "1" Then
    If Mid(sResult3, 4, 1) = "1" Then       'FCの24桁目→FFの4桁目 2011/04/07 SETsw kubota
        If iMultiFlg = 0 Then '' ブロック品番保障0,2のMultiBlockの真ん中はしない
            If iELCs_Flg = 0 Or iELCs_Flg = 2 Then  '' 振替元品番チェック振替元以外の品番はしない
                RET_3_4 = funChkFurikae3_4(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg, iSmpGetFlg, sSamplID1, sSamplID2, iKcnt)
                If RET_3_4 <> 0 Then
                    funChkFurikaeShiyou = RET_3_4
                    If RET_3_4 > 0 Then GoTo Apl_Exit
                    GoTo Apl_Error
                End If
            End If
        End If
    End If
    If RET <> 0 Then
        funChkFurikaeShiyou = RET
        If RET > 0 Then GoTo Apl_Exit
        GoTo Apl_Error
    End If
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    '3-2 Warp実績チェック       05/12/28 ooba
'    If Mid(sResult, 22, 1) = "1" Then
    If Mid(sResult3, 2, 1) = "1" Then       'FCの22桁目→FFの2桁目 2011/04/07 SETsw kubota
        If iMultiFlg = 0 Then ''<<複数品番判定対応20060502 SMP桜井　ブロック品番保障0,2のMultiBlockの真ん中はしない
            If iELCs_Flg = 0 Or iELCs_Flg = 2 Then  ''<<複数品番判定対応20060502 SMP桜井　振替元品番チェック振替元以外の品番はしない
                RET = funChkFurikae3_2(tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg)
                If RET <> 0 Then
                    funChkFurikaeShiyou = RET
                    If RET > 0 Then GoTo Apl_Exit
                    GoTo Apl_Error
                End If
            End If
        End If
    End If
    'Chg Start 2011/07/22 SMPK Nakamura 結晶面傾きチェック追加対応
'    '3-3 合成角度実績チェック   05/12/28 ooba
    '3-3 Ｘ線実績チェック
    'Chg End 2011/07/22 SMPK Nakamura 結晶面傾きチェック追加対応
'    If Mid(sResult, 23, 1) = "1" Then
    If Mid(sResult3, 3, 1) = "1" Then       'FCの23桁目→FFの3桁目 2011/04/07 SETsw kubota
        If iMultiFlg = 0 Then ''<<複数品番判定対応20060502 SMP桜井　ブロック品番保障0,2のMultiBlockの真ん中はしない
            If iELCs_Flg = 0 Or iELCs_Flg = 2 Then  ''<<複数品番判定対応20060502 SMP桜井　振替元品番チェック振替元以外の品番はしない
                RET = funChkFurikae3_3(tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg)
                If RET <> 0 Then
                    funChkFurikaeShiyou = RET
                    If RET > 0 Then GoTo Apl_Exit
                    GoTo Apl_Error
                End If
            End If
        End If
    End If

'Add Start 2011/04/25 SMPK Miyata
    '3-5 ＷＦＣ評価実績チェック
    If Mid(sResult3, 5, 1) = "1" Then       'FCの21桁目→FFの1桁目 2011/04/07 SETsw kubota
        RET = funChkFurikae3_5(sProccd, sKeyID, tOld_Hinban, tNew_Hinban, iErr_Code, sErr_Msg, iSmpGetFlg, sSamplID1, sSamplID2, iKcnt)
        If RET <> 0 Then
            funChkFurikaeShiyou = RET
            If RET > 0 Then GoTo Apl_Exit
            GoTo Apl_Error
        End If
    End If
'Add End   2011/04/25 SMPK Miyata

    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Select Case iErr_Code
        Case 0      '正常終了
            sErr_Msg = ""
'        Case 1      '正常終了（該当ﾃﾞｰﾀなし）
'            sErr_Msg = "振替可能な品番はありません。"
'        Case -1
'            sErr_Msg = "入力引数値にｴﾗｰがあります。"
        Case -2
            sErr_Msg = "振替指示ﾃﾞｰﾀ取得ｴﾗｰ"
        Case -3
            sErr_Msg = "DBｱｸｾｽｴﾗｰ(" & sErr_Msg & ")"
        Case -4
            sErr_Msg = "APLｴﾗｰ(" & sErr_Msg & ")"
        Case -5
            sErr_Msg = "想定外の仕様ﾃﾞｰﾀ(" & sErr_Msg & ")"
    End Select
    
    '振替ＮＧの場合、エラーコードを文字列に変換し、エラーメッセージコードとして返す。
''''    If iErr_Code > 1 Then
'''    If funChkFurikaeShiyou = 1 Then
'''        sErr_Msg = "F" & CStr(iErr_Code)
'''    End If
    
    Exit Function
    
Apl_Error:
    iErr_Code = funChkFurikaeShiyou
    GoTo Apl_Exit

Apl_down:
    funChkFurikaeShiyou = -4
    iErr_Code = funChkFurikaeShiyou
    GoTo Apl_Exit
    
End Function

'------------------------------------------------
' 品番組合せチェック (チェック１－７，１－８)
'------------------------------------------------

'概要      :仕掛ロット内の全品番に対して組合せチェックを行い、結果を返す。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型             :説明
'          :sProccd         ,I  ,String         :工程番号
'          :sCrynum         ,I  ,String         :結晶番号
'          :tKumi_Hinban()  ,I  ,tFullHinban    :ﾁｪｯｸ品番
'          :iKumi_Row()     ,I  ,Integer        :品番行位置
'          :iHinPnt         ,O  ,Integer        :ﾁｪｯｸNG品番行位置
'          :iErr_Code       ,O  ,Integer        :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String         :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer        :成否(0:正常終了(ﾁｪｯｸOK),1:正常終了(ﾁｪｯｸNG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :06/04/25 ooba

Public Function funChkKumiHinban(sProccd As String, sCryNum As String, _
                                    tKumi_Hinban() As tFullHinban, iKumi_Row() As Integer, _
                                    iHinPnt As Integer, iErr_Code As Integer, _
                                    sErr_Msg As String) As Integer

    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkKumiHinban = 0
    iHinPnt = 0
    iErr_Code = 0
    sErr_Msg = ""

    '工程ﾁｪｯｸ
    If (left(sProccd, 4) <> "CC10") And (left(sProccd, 4) <> "CC31") And _
       (left(sProccd, 4) <> "CC60") And (left(sProccd, 4) <> "CC61") And _
       (left(sProccd, 4) <> "CC73") And (left(sProccd, 4) <> "CW74") And _
       (left(sProccd, 4) <> "CW75") And (left(sProccd, 4) <> "CW76") Then
        funChkKumiHinban = -1
        sErr_Msg = "工程"
        GoTo Apl_Error
    End If
    '結晶番号ﾁｪｯｸ
    If Len(sCryNum) <> 12 Then
        funChkKumiHinban = -1
        sErr_Msg = "結晶番号"
        GoTo Apl_Error
    End If
    '品番ﾁｪｯｸ
    If UBound(tKumi_Hinban) = 0 Then
        funChkKumiHinban = -1
        sErr_Msg = "品番0"
        GoTo Apl_Error
    End If
    
    '狙い品番取得(引上指示以外)
    If left(sProccd, 4) <> "CC10" Then
        If funNeraiHinGet(sCryNum, tKumi_Hinban(0)) = FUNCTION_RETURN_FAILURE Then
            funChkKumiHinban = -1
            sErr_Msg = "狙い品番"
            GoTo Apl_Error
        End If
    End If
    
    '------------------------------------------ 指示取得 ------------------------------------------------------
    '組合せチェック指示データ取得
    sResult = ""
'    RET = funCodeDBGet("SB", "FC", sProccd, 0, " ", sResult)
    RET = funCodeDBGet("SB", "FD", sProccd, 0, " ", sResult)        'FC→FD 2011/04/07修正 SETsw kubota
    If RET <> 0 Then
        funChkKumiHinban = -2
        GoTo Apl_Error
    End If
    '------------------------------------------ Make SQL ------------------------------------------------------
    '1-7 品番組合せチェック１
    If Mid(sResult, 7, 1) = "1" Then
        RET = funChkFurikae1_7(tKumi_Hinban(), iKumi_Row(), iHinPnt, iErr_Code, sErr_Msg)
        If RET <> 0 Then
            funChkKumiHinban = RET
            If RET > 0 Then GoTo Apl_Exit
            GoTo Apl_Error
        End If
    End If
    '1-8 品番組合せチェック２
    If Mid(sResult, 8, 1) = "1" Then
        RET = funChkFurikae1_8(tKumi_Hinban(), iKumi_Row(), iHinPnt, iErr_Code, sErr_Msg)
        If RET <> 0 Then
            funChkKumiHinban = RET
            If RET > 0 Then GoTo Apl_Exit
            GoTo Apl_Error
        End If
    End If
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Select Case iErr_Code
        Case 0      '正常終了
            sErr_Msg = ""
        Case -1
            sErr_Msg = "入力引数値ｴﾗｰ(" & sErr_Msg & ")"
        Case -2
            sErr_Msg = "組合せﾁｪｯｸ指示ﾃﾞｰﾀ取得ｴﾗｰ"
        Case -3
            sErr_Msg = "DBｱｸｾｽｴﾗｰ(" & sErr_Msg & ")"
        Case -4
            sErr_Msg = "APLｴﾗｰ(" & sErr_Msg & ")"
        Case -5
            sErr_Msg = "想定外の仕様ﾃﾞｰﾀ(" & sErr_Msg & ")"
    End Select
    
    Exit Function
    
Apl_Error:
    iErr_Code = funChkKumiHinban
    GoTo Apl_Exit

Apl_down:
    funChkKumiHinban = -4
    iErr_Code = funChkKumiHinban
    GoTo Apl_Exit
    
End Function

'------------------------------------------------
' マルチブロック組合せ可否チェック（仕様）
'------------------------------------------------

'概要      :パラメータに指定された、ブロック先頭品番からブロック最尾品番の組合せが可能かどうかをチェックし、結果を返す。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型             :説明
'          :sProccd         ,I  ,String         :工程番号
'          :sKeyID          ,I  ,String         :ﾌﾞﾛｯｸID、又は、SXL-ID
'          :tTop_Hinban     ,I  ,tFullHinban    :先頭品番(構造体)
'          :tBtm_Hinban     ,I  ,tFullHinban    :最尾品番(構造体)
'          :iErr_Code       ,O  ,Integer        :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String         :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer        :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :2011/07/13 新規作成　SMPK Nakamura
Public Function funChkMultiShiyou(sProccd As String, sKeyID As String, tTop_Hinban As tFullHinban, tBtm_Hinban As tFullHinban, _
                                  iErr_Code As Integer, sErr_Msg As String) As Integer
    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkMultiShiyou = 0
    iErr_Code = 0
    sErr_Msg = ""
'    '------------------------------------------ Ｇ品番、Ｚ品番チェック -------------------------------------------------
'    '通常品番 → Ｇ品番、Ｚ品番　⇒　振替ＯＫ
'    If (Trim$(tNew_Hinban.hinban) = "G") Or (Trim$(tNew_Hinban.hinban) = "Z") Then GoTo Apl_Exit
'
'    'Ｚ品番、Ｇ品番 → 通常品番　⇒　振替ＯＫ
'    If (Trim$(tOld_Hinban.hinban) = "Z") Or (Trim$(tOld_Hinban.hinban) = "G") Then GoTo Apl_Exit
    
    '------------------------------------------ 入力チェック -------------------------------------------------
    '工程番号のチェック
    If Trim$(sProccd) = "" Then
        funChkMultiShiyou = -1
        sErr_Msg = "入力引数値エラー(工程番号指定なし)"
        GoTo Apl_Error
    End If
    JudgKoutei = sProccd
    'ﾌﾞﾛｯｸID、SXL-IDのチェック
    If Trim$(sKeyID) = "" Then
        funChkMultiShiyou = -1
        sErr_Msg = "入力引数値エラー(ﾌﾞﾛｯｸID or SXL-ID指定なし, 工程番号 : " & sProccd & ")"
        GoTo Apl_Error
    End If
    If (left(sProccd, 4) = "CC31") Or (left(sProccd, 4) = "CC60") Or (left(sProccd, 4) = "CC61") Or (left(sProccd, 4) = "CC73") Or _
       (left(sProccd, 4) = "CW74") Or (left(sProccd, 4) = "CW75") Or (left(sProccd, 4) = "CW76") Then
        If (left(sProccd, 4) = "CW75") Or (left(sProccd, 4) = "CW76") Then
            If Len(sKeyID) <> 13 Then
                funChkMultiShiyou = -1
                sErr_Msg = "入力引数値エラー(SXL-ID : " & sKeyID & ")"
                GoTo Apl_Error
            End If
        Else
            If Len(sKeyID) <> 12 Then
                funChkMultiShiyou = -1
                sErr_Msg = "入力引数値エラー(ﾌﾞﾛｯｸID : " & sKeyID & ")"
                GoTo Apl_Error
            End If
        End If
    Else
        funChkMultiShiyou = -1
        sErr_Msg = "入力引数値エラー(工程番号 : " & sProccd & ")"
        GoTo Apl_Error
    End If
'    'ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞの指定チェック
'    If (left(sProccd, 4) = "CC60") Or (left(sProccd, 4) = "CW75") Then
'       If (IsNull(iSmpGetFlg)) Or (iSmpGetFlg <> 0 And iSmpGetFlg <> 1) Then
'          funChkMultiShiyou = -1
'          sErr_Msg = "入力引数値エラー(ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ : " & iSmpGetFlg & ")"
'          GoTo Apl_Error
'       End If
'       If (iSmpGetFlg = 1) And _
'          (IsNull(sSamplID1) Or Trim$(sSamplID1) = "" Or IsNull(sSamplID2) Or Trim$(sSamplID2) = "") Then
'          funChkMultiShiyou = -1
'          sErr_Msg = "入力引数値エラー(ｻﾝﾌﾟﾙID指定なし)"
'          GoTo Apl_Error
'       End If
'    End If
    
    '------------------------------------------ 指示取得 ------------------------------------------------------
    '振替指示データ取得
    sResult = ""
    RET = funCodeDBGet("SB", "FG", sProccd, 0, " ", sResult)
    If RET <> 0 Then
        funChkMultiShiyou = -2
        GoTo Apl_Error
    End If
    '------------------------------------------ Make SQL ------------------------------------------------------
    '1-16 結晶面傾き組合せチェック
    If Mid(sResult, 16, 1) = "1" Then
        RET = funChkFurikae1_16(sProccd, sKeyID, tTop_Hinban, tBtm_Hinban, iErr_Code, sErr_Msg)
        If RET <> 0 Then
            funChkMultiShiyou = RET
            If RET > 0 Then GoTo Apl_Exit
            GoTo Apl_Error
        End If
    End If

    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Select Case iErr_Code
        Case 0      '正常終了
            sErr_Msg = ""
        Case -2
            sErr_Msg = "振替指示ﾃﾞｰﾀ取得ｴﾗｰ"
        Case -3
            sErr_Msg = "DBｱｸｾｽｴﾗｰ(" & sErr_Msg & ")"
        Case -4
            sErr_Msg = "APLｴﾗｰ(" & sErr_Msg & ")"
        Case -5
            sErr_Msg = "想定外の仕様ﾃﾞｰﾀ(" & sErr_Msg & ")"
    End Select
    
    Exit Function
    
Apl_Error:
    iErr_Code = funChkMultiShiyou
    GoTo Apl_Exit

Apl_down:
    funChkMultiShiyou = -4
    iErr_Code = funChkMultiShiyou
    GoTo Apl_Exit
    
End Function

'------------------------------------------------
' 組み合わせ品番チェック
'------------------------------------------------

'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sKeyID          ,I  ,String       :ﾌﾞﾛｯｸID、又は、SXL-ID
'          :tOld_Hinban     ,I  ,tFullHinban  :振替元品番(構造体)
'          :tNew_Hinban     ,I  ,tFullHinban  :振替先品番(構造体)
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/19 新規作成　SB

Public Function funChkFurikae1_1(sProccd As String, sBlockId As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iErr_Code As Integer, sErr_Msg As String) As Integer


    Dim sql As String       'SQL全体
    Dim rs  As OraDynaset   'RecordSet
    Dim sResult As String   'コードＤＢ取得関数の取得変数   '05/04/04 ooba
    Dim RET     As Integer  '戻り値     '05/04/04 ooba
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae1_1 = 0
    
    Erase tbl_chk1_1
    '------------------------------------------ 振替元品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-1 振替元品番仕様取得(" & tOld_Hinban.hinban & Format(tOld_Hinban.mnorevno, "00") & tOld_Hinban.factory & tOld_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    'sql = sql & "SELECT E021.HWFTYPE,E036.BLOCKHFLAG " & vbCrLf     2004/12/21変更
    'sql = sql & "FROM   TBCME021 E021,TBCME036 E036 " & vbCrLf
    'sql = sql & "WHERE  E021.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    'sql = sql & "       E021.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    'sql = sql & "       E021.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    'sql = sql & "       E021.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    'sql = sql & "       E036.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    'sql = sql & "       E036.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    'sql = sql & "       E036.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    'sql = sql & "       E036.OPECOND   =   '" & tOld_Hinban.opecond & "' " & vbCrLf
    sql = sql & "SELECT E018.HSXTYPE,E036.BLOCKHFLAG " & vbCrLf
    sql = sql & "      ,E020.HSXSDSLP" & vbCrLf                     '2009/08/06追加 SETsw kubota
    sql = sql & "FROM   TBCME018 E018,TBCME036 E036 " & vbCrLf
    sql = sql & "      ,TBCME020 E020" & vbCrLf                     '2009/08/06追加 SETsw kubota
    sql = sql & "WHERE  E018.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E018.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E018.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    '>>>>> TBCME020追加 2009/08/06 SETsw kubota ----------
    sql = sql & "       E020.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E020.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E020.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E020.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    '<<<<< TBCME020追加 2009/08/06 SETsw kubota ----------
    sql = sql & "       E036.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tOld_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    With tbl_chk1_1(0)
        'If IsNull(rs("HWFTYPE")) = False Then .HWFTYPE = rs("HWFTYPE") Else .HWFTYPE = " "                  'ﾀｲﾌﾟ
        If IsNull(rs("HSXTYPE")) = False Then .HSXFTYPE = rs("HSXTYPE") Else .HSXFTYPE = " "                  'ﾀｲﾌﾟ
        If IsNull(rs("BLOCKHFLAG")) = False Then .BLOCKHFLAG = rs("BLOCKHFLAG") Else .BLOCKHFLAG = " "      'ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞ
        If IsNull(rs("HSXSDSLP")) = False Then .HSXSDSLP = rs("HSXSDSLP") Else .HSXSDSLP = " "              'シード傾き   2009/08/06追加 SETsw kubota
    End With
    
    Set rs = Nothing
    '------------------------------------------ 振替先品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-1 振替先品番仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    'sql = sql & "SELECT E021.HWFTYPE,E036.BLOCKHFLAG " & vbCrLf
    'sql = sql & "FROM   TBCME021 E021,TBCME036 E036 " & vbCrLf
    'sql = sql & "WHERE  E021.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    'sql = sql & "       E021.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    'sql = sql & "       E021.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    'sql = sql & "       E021.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    'sql = sql & "       E036.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    'sql = sql & "       E036.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    'sql = sql & "       E036.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    'sql = sql & "       E036.OPECOND   =   '" & tNew_Hinban.opecond & "' " & vbCrLf
    sql = sql & "SELECT E018.HSXTYPE,E036.BLOCKHFLAG " & vbCrLf
    sql = sql & "      ,E020.HSXSDSLP" & vbCrLf                     '2009/08/06追加 SETsw kubota
    sql = sql & "FROM   TBCME018 E018,TBCME036 E036 " & vbCrLf
    sql = sql & "      ,TBCME020 E020" & vbCrLf                     '2009/08/06追加 SETsw kubota
    sql = sql & "WHERE  E018.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E018.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E018.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    '>>>>> TBCME020追加 2009/08/06 SETsw kubota ----------
    sql = sql & "       E020.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E020.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E020.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E020.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    '<<<<< TBCME020追加 2009/08/06 SETsw kubota ----------
    sql = sql & "       E036.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tNew_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    With tbl_chk1_1(1)
        'If IsNull(rs("HWFTYPE")) = False Then .HWFTYPE = rs("HWFTYPE") Else .HWFTYPE = " "                  'ﾀｲﾌﾟ
        If IsNull(rs("HSXTYPE")) = False Then .HSXFTYPE = rs("HSXTYPE") Else .HSXFTYPE = " "                  'ﾀｲﾌﾟ
        If IsNull(rs("BLOCKHFLAG")) = False Then .BLOCKHFLAG = rs("BLOCKHFLAG") Else .BLOCKHFLAG = " "      'ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞ
        If IsNull(rs("HSXSDSLP")) = False Then .HSXSDSLP = rs("HSXSDSLP") Else .HSXSDSLP = " "              'シード傾き   2009/08/06追加 SETsw kubota
    End With
    
    Set rs = Nothing
    '------------------------------------------ 各種チェック ------------------------------------------------------
    On Error GoTo Apl_down
    'タイプのチェック
    sErr_Msg = "1-1 ﾀｲﾌﾟﾁｪｯｸ"
    If Trim$(tbl_chk1_1(0).HSXFTYPE) <> Trim$(tbl_chk1_1(1).HSXFTYPE) Then
        If Trim$(tbl_chk1_1(1).HSXFTYPE) <> "Z" Then    '不問品番への振替はOK 2011/05/11 SETsw kubota
            funChkFurikae1_1 = 1
            iErr_Code = 1101
            sErr_Msg = "CHECK1-1,ﾀｲﾌﾟ不一致の為、振替できません。"
    '--------------- 2008/07/25 INSERT START  By Systech ---------------
            gsTbcmy028ErrCode = "00001"
    '--------------- 2008/07/25 INSERT  END   By Systech ---------------
            GoTo Apl_Exit
        End If
    End If
    'ブロック単位保証のチェック
    sErr_Msg = "1-1 ﾌﾞﾛｯｸ単位保証ﾁｪｯｸ"
'    If Trim$(tbl_chk1_1(0).BLOCKHFLAG) <> Trim$(tbl_chk1_1(1).BLOCKHFLAG) Then
'        funChkFurikae1_1 = 1
'        iErr_Code = 1102
'        sErr_Msg = "CHECK1-1,ﾌﾞﾛｯｸ単位保障不一致の為、振替できません。"
'        GoTo Apl_Exit
'    End If

    ''ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞの振替ﾁｪｯｸ変更  05/04/04 ooba START ======================================>
    sResult = ""
    RET = funCodeDBGet("SB", "BH", tbl_chk1_1(0).BLOCKHFLAG, 1, tbl_chk1_1(1).BLOCKHFLAG, sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→元:" & tbl_chk1_1(0).BLOCKHFLAG & ", 先:" & tbl_chk1_1(1).BLOCKHFLAG
        GoTo CodeDBGet_Error
    End If
    If sResult = 0 Then
        funChkFurikae1_1 = 1
        iErr_Code = 1102
        sErr_Msg = "CHECK1-1,ﾌﾞﾛｯｸ単位保証、振替不可能です。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00002"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    ''ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞの振替ﾁｪｯｸ変更  05/04/04 ooba END ========================================>
    
'>>>>> シード傾きのチェックを1-2から1-1へ移動 2009/08/06 SETsw kubota ------
    'シード傾きのチェック
    sErr_Msg = "1-1 ｼｰﾄﾞ傾きﾁｪｯｸ"
    If Trim$(tbl_chk1_1(0).HSXSDSLP) <> Trim$(tbl_chk1_1(1).HSXSDSLP) Then
        funChkFurikae1_1 = 1
        iErr_Code = 1205
        sErr_Msg = "CHECK1-1,シード傾き不一致の為、振替できません。"
        gsTbcmy028ErrCode = "00007"
        GoTo Apl_Exit
    End If
'<<<<< シード傾きのチェックを1-2から1-1へ移動 2009/08/06 SETsw kubota ------
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae1_1 = 0 Then
        funChkFurikae1_1 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae1_1 = -4
    GoTo Apl_Exit

'05/04/04 ooba
CodeDBGet_Error:
    Set rs = Nothing
    If funChkFurikae1_1 = 0 Then
        funChkFurikae1_1 = -5
    End If
    GoTo Apl_Exit
    
End Function
    
'------------------------------------------------
' 振替先と振替元の常識仕様チェック
'------------------------------------------------

'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sKeyID          ,I  ,String       :ﾌﾞﾛｯｸID、又は、SXL-ID
'          :tOld_Hinban     ,I  ,tFullHinban  :振替元品番(構造体)
'          :tNew_Hinban     ,I  ,tFullHinban  :振替先品番(構造体)
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :sPlshMeth       ,I  ,String       :研削方法(加工払出画面) M:MGR(MGRの場合、溝位置方位ロック解除)  2011/10/13 SETsw kubota
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/19 新規作成　SB

Public Function funChkFurikae1_2(sProccd As String, sBlockId As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iErr_Code As Integer, sErr_Msg As String _
                               , Optional ByVal sPlshMeth As String = "" _
                               ) As Integer


    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim sClass      As String       '区分 '' add 0108
    Dim wXtal()     As String                            '2010/04/16 Kameda
    Dim wINPOS()    As Integer                           '2010/04/16 Kameda
    Dim Xsen        As type_DBDRV_scmzc_fcmkc001c_X      '2010/04/16 Kameda
    Dim Xsiyou      As type_DBDRV_scmzc_fcmkc001c_Siyou  '2010/04/16 Kameda
    Dim JUDGXY     As Boolean                            'X線判定用フラグ追加 2010/04/16
    Dim JUDGX      As Boolean                            'X線判定用フラグ追加 2010/04/16
    Dim JUDGY      As Boolean                            'X線判定用フラグ追加 2010/04/16
    Dim cnt        As Integer
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae1_2 = 0
    
    '------------------------------------------ 振替元品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-2 振替元品番仕様取得(" & tOld_Hinban.hinban & Format(tOld_Hinban.mnorevno, "00") & tOld_Hinban.factory & tOld_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E018.HSXCDIR,E018.HSXCSCEN,E018.HSXDOP,E023.HWFCDOP,E020.HSXSDSLP,E018.HSXDPDIR, " & vbCrLf
    '2010/05/24 SIRD対応 Y.Hitomi
    sql = sql & "       SUBSTR(E018.MCNO,1,1) MCNO1,SUBSTR(E018.MCNO,4,1) MCNO2,SUBSTR(E018.MCNO,3,1) MCNO3,E036.DCHYUUBU,E048.HWFSIRDHS " & vbCrLf
'    sql = sql & "       SUBSTR(E018.MCNO,1,1) MCNO1,SUBSTR(E018.MCNO,4,1) MCNO2,SUBSTR(E018.MCNO,3,1) MCNO3,E036.DCHYUUBU " & vbCrLf
    sql = sql & "FROM   TBCME018 E018,TBCME023 E023,TBCME020 E020,TBCME036 E036,TBCME048 E048 " & vbCrLf
    '2010/05/24 SIRD対応 Y.Hitomi
'    sql = sql & "FROM   TBCME018 E018,TBCME023 E023,TBCME020 E020,TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E018.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E018.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E018.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E023.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E023.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E023.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E023.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E020.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E020.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E020.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E020.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E036.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    '2010/05/24 SIRD対応 Y.Hitomi
    sql = sql & "       E048.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E048.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E048.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E048.OPECOND   =   '" & tOld_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    Erase tbl_chk1_2
    With tbl_chk1_2(0)
        If IsNull(rs("HSXCDIR")) = False Then .HSXCDIR = rs("HSXCDIR") Else .HSXCDIR = " "          ' 結晶面方位
        If IsNull(rs("HSXCSCEN")) = False Then .HSXCSCEN = rs("HSXCSCEN") Else .HSXCSCEN = 0        ' 結晶面傾き中心
        If IsNull(rs("HSXDOP")) = False Then .HSXDOP = rs("HSXDOP") Else .HSXDOP = " "              ' ドーパント
        If IsNull(rs("HWFCDOP")) = False Then .HWFCDOP = rs("HWFCDOP") Else .HWFCDOP = " "          ' 結晶ドープ
'        If IsNull(rs("HSXSDSLP")) = False Then .HSXSDSLP = rs("HSXSDSLP") Else .HSXSDSLP = " "      ' シード傾き   2009/08/06削除 SETsw kubota
        If IsNull(rs("HSXDPDIR")) = False Then .HSXDPDIR = rs("HSXDPDIR") Else .HSXDPDIR = " "      ' 溝位置方位
        If IsNull(rs("MCNO1")) = False Then .MCNO1 = rs("MCNO1") Else .MCNO1 = " "                  ' 品種
        If IsNull(rs("MCNO2")) = False Then .MCNO2 = rs("MCNO2") Else .MCNO2 = " "                  ' 引上げ速度
        If IsNull(rs("MCNO3")) = False Then .MCNO3 = rs("MCNO3") Else .MCNO3 = " "                  ' HZタイプ
        If IsNull(rs("DCHYUUBU")) = False Then .DCHYUUBU = rs("DCHYUUBU") Else .DCHYUUBU = "2"      ' ドローチューブ
        If IsNull(rs("HWFSIRDHS")) = False Then .HWFSIRDHS = rs("HWFSIRDHS") Else .HWFSIRDHS = " "  ' SIRD保証方法 処 2010/05/24 SIRD対応 Y.Hitomi
    End With
    
    Set rs = Nothing
    '------------------------------------------ 振替先品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-2 振替先品番仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E018.HSXCDIR,E018.HSXCSCEN,E018.HSXDOP,E023.HWFCDOP,E020.HSXSDSLP,E018.HSXDPDIR, " & vbCrLf
    sql = sql & "       SUBSTR(E018.MCNO,1,1) MCNO1,SUBSTR(E018.MCNO,4,1) MCNO2,SUBSTR(E018.MCNO,3,1) MCNO3,E036.DCHYUUBU, " & vbCrLf   '' chg 0108
    sql = sql & "       E036.NDOPHUFLG,E036.CDOPHUFLG,E048.HWFSIRDHS, " & vbCrLf    '' 2010/05/24 SIRD対応 Y.Hitomi
    sql = sql & "       E018.HSXCSCEN,E018.HSXCSMIN,E018.HSXCSMAX,E018.HSXCYCEN,E018.HSXCYMIN,E018.HSXCYMAX,E018.HSXCTCEN,E018.HSXCTMIN,E018.HSXCTMAX " & vbCrLf   '2010/04/16 Kameda
    sql = sql & "FROM   TBCME018 E018,TBCME023 E023,TBCME020 E020,TBCME036 E036,TBCME048 E048 " & vbCrLf
'    sql = sql & "FROM   TBCME018 E018,TBCME023 E023,TBCME020 E020,TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E018.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E018.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E018.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E023.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E023.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E023.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E023.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E020.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E020.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E020.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E020.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E036.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    '2010/05/24 SIRD対応 Y.Hitomi
    sql = sql & "       E048.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E048.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E048.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E048.OPECOND   =   '" & tNew_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    With tbl_chk1_2(1)
        If IsNull(rs("HSXCDIR")) = False Then .HSXCDIR = rs("HSXCDIR") Else .HSXCDIR = " "          ' 結晶面方位
        If IsNull(rs("HSXCSCEN")) = False Then .HSXCSCEN = rs("HSXCSCEN") Else .HSXCSCEN = 0        ' 結晶面傾き中心
        If IsNull(rs("HSXDOP")) = False Then .HSXDOP = rs("HSXDOP") Else .HSXDOP = " "              ' ドーパント
        If IsNull(rs("HWFCDOP")) = False Then .HWFCDOP = rs("HWFCDOP") Else .HWFCDOP = " "          ' 結晶ドープ
'        If IsNull(rs("HSXSDSLP")) = False Then .HSXSDSLP = rs("HSXSDSLP") Else .HSXSDSLP = " "      ' シード傾き   2009/08/06削除 SETsw kubota
        If IsNull(rs("HSXDPDIR")) = False Then .HSXDPDIR = rs("HSXDPDIR") Else .HSXDPDIR = " "      ' 溝位置方位
        If IsNull(rs("MCNO1")) = False Then .MCNO1 = rs("MCNO1") Else .MCNO1 = " "                  ' 品種
        If IsNull(rs("MCNO2")) = False Then .MCNO2 = rs("MCNO2") Else .MCNO2 = " "                  ' 引上げ速度
        If IsNull(rs("MCNO3")) = False Then .MCNO3 = rs("MCNO3") Else .MCNO3 = " "                  ' HZタイプ
        If IsNull(rs("DCHYUUBU")) = False Then .DCHYUUBU = rs("DCHYUUBU") Else .DCHYUUBU = "2"      ' ドローチューブ
        If IsNull(rs("NDOPHUFLG")) = False Then .NDOPHUFLG = rs("NDOPHUFLG") Else .NDOPHUFLG = " "  ' 窒素ドープ振替可否フラグ '' add 0108
        If IsNull(rs("CDOPHUFLG")) = False Then .CDOPHUFLG = rs("CDOPHUFLG") Else .CDOPHUFLG = " "  ' Cドープ振替可否フラグ '' add 0108
        If IsNull(rs("HWFSIRDHS")) = False Then .HWFSIRDHS = rs("HWFSIRDHS") Else .HWFSIRDHS = " "  ' SIRD保証方法 処 2010/05/24 SIRD対応 Y.Hitomi
    End With
    With Xsiyou
        .HSXCSCEN = fncNullCheck(rs("HSXCSCEN"))        ' 品ＳＸ面傾き中心    2010/04/16 Kameda
        .HSXCSMIN = fncNullCheck(rs("HSXCSMIN"))        ' 品ＳＸ面傾き下限    2010/04/16 Kameda
        .HSXCSMAX = fncNullCheck(rs("HSXCSMAX"))        ' 品ＳＸ面傾き上限    2010/04/16 Kameda
        .HSXCTCEN = fncNullCheck(rs("HSXCTCEN"))        ' 品ＳＸ面傾き縦中心  2010/04/16 Kameda
        .HSXCTMIN = fncNullCheck(rs("HSXCTMIN"))        ' 品ＳＸ面傾き縦下限  2010/04/16 Kameda
        .HSXCTMAX = fncNullCheck(rs("HSXCTMAX"))        ' 品ＳＸ面傾き縦上限  2010/04/16 Kameda
        .HSXCYCEN = fncNullCheck(rs("HSXCYCEN"))        ' 品ＳＸ面傾き横中心  2010/04/16 Kameda
        .HSXCYMIN = fncNullCheck(rs("HSXCYMIN"))        ' 品ＳＸ面傾き横下限  2010/04/16 Kameda
        .HSXCYMAX = fncNullCheck(rs("HSXCYMAX"))        ' 品ＳＸ面傾き横上限  2010/04/16 Kameda
    End With
    Set rs = Nothing
    '------------------------------------------ 指示取得 ------------------------------------------------------
    On Error GoTo Apl_down
    '結晶面方位のチェック
    sErr_Msg = "1-2 結晶面方位ﾁｪｯｸ"
    If Trim$(tbl_chk1_2(0).HSXCDIR) <> Trim$(tbl_chk1_2(1).HSXCDIR) Then
        funChkFurikae1_2 = 1
        iErr_Code = 1201
        sErr_Msg = "CHECK1-2,結晶面方位不一致の為、振替できません。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00003"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    
''2008/11/27 結晶面傾中心チェック緩和(2) DEL By Systech Start
''        '結晶面傾中心のチェック
''        sErr_Msg = "1-2 結晶面傾中心ﾁｪｯｸ"
''        If (Trim$(tbl_chk1_2(0).HSXCSCEN) = 4) Or (Trim$(tbl_chk1_2(1).HSXCSCEN) = 4) Then
''            If Trim$(tbl_chk1_2(0).HSXCSCEN) <> Trim$(tbl_chk1_2(1).HSXCSCEN) Then
''                funChkFurikae1_2 = 1
''                iErr_Code = 1202
''                sErr_Msg = "CHECK1-2,結晶面傾中心不一致の為、振替できません。"
''    '--------------- 2008/07/25 INSERT START  By Systech ---------------
''                gsTbcmy028ErrCode = "00004"
''    '--------------- 2008/07/25 INSERT  END   By Systech ---------------
''                GoTo Apl_Exit
''            End If
''        End If
''2008/11/27 結晶面傾中心チェック緩和(2) DEL By Systech End
    
    ''2010/04/16 結晶面傾中心仕様判断条件の追加 №100087 Kameda    <----- 1-10へ移動
    'If left(sProccd, 4) = "CW76" Then
    '    '面傾中心仕様0.00度品から0.00度品以外への振替を禁止
    '    sErr_Msg = "1-2 結晶面傾中心ﾁｪｯｸ"
    '    If Trim$(tbl_chk1_2(0).HSXCSCEN) = 0 Then
    '        If Trim$(tbl_chk1_2(1).HSXCSCEN) <> 0 Then
    '            funChkFurikae1_2 = 1
    '            iErr_Code = 1201
    '            sErr_Msg = "CHECK1-2,結晶面傾中心不一致の為、振替できません。"
    '            gsTbcmy028ErrCode = "00004"
    '            GoTo Apl_Exit
    '        End If
    '    End If
    '    '面傾中心仕様1.00度以下品から0.00度品への振替はＸ線実績が振替先の仕様範囲内
    '    If Trim$(tbl_chk1_2(0).HSXCSCEN) < 1 And Trim$(tbl_chk1_2(1).HSXCSCEN) = 0 Then
    '        sql = vbNullString
    '        sql = sql & "SELECT XTALCA,INPOSCA FROM XSDCA " & vbCrLf
    '        sql = sql & "WHERE  SXLIDCA = '" & sBlockId & "' AND " & vbCrLf
    '        sql = sql & "       LIVKCA  = '0' " & vbCrLf
    '
    '        On Error GoTo db_Error
    '        'SQL文の実行
    '        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    '
    '        '該当データなし
    '        If rs.EOF Or rs.RecordCount = 0 Then GoTo db_Error
    '
    '        ReDim wXTAL(rs.RecordCount)
    '        ReDim wINPOS(rs.RecordCount)
    '        For cnt = 1 To rs.RecordCount
    '            wXTAL(cnt) = rs("XTALCA")
    '            wINPOS(cnt) = rs("INPOSCA")
    '            rs.MoveNext
    '        Next
    '        Set rs = Nothing
    '
    '        For cnt = 1 To UBound(wXTAL)
    '            sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, XRAYX,XRAYY,XRAYXY, REGDATE "
    '            sql = sql & "from TBCMJ021 "
    '            sql = sql & "where CRYNUM = '" & wXTAL(cnt) & "' and "
    '            sql = sql & "      POSITION = '" & wINPOS(cnt) & "' and "
    '            sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ021 "
    '            sql = sql & "                 where CRYNUM = '" & wXTAL(cnt) & "' )"
    '
    '            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    '
    '            If rs.RecordCount <> 0 Then
    '                With Xsen
    '                    .CRYNUM = rs("CRYNUM")          ' 結晶番号
    '                    .POSITION = rs("POSITION")      ' 位置
    '                    .SMPKBN = rs("SMPKBN")          ' サンプル区分
    '                    .TRANCOND = rs("TRANCOND")      ' 処理条件
    '                    .TRANCNT = rs("TRANCNT")        ' 処理回数
    '                    .SMPLNO = rs("SMPLNO")          ' サンプルＮｏ
    '                    .SMPLUMU = rs("SMPLUMU")        ' サンプル有無
    '                    .XX = rs("XRAYX")               ' 測定値X
    '                    .XY = rs("XRAYY")               ' 測定値Y
    '                    .XXY = rs("XRAYXY")             ' 測定値XY
    '                    .REGDATE = rs("REGDATE")        ' 登録日付
    '                End With
    '                If CrXjudg(Xsiyou, Xsen, JUDGXY, JUDGX, JUDGY) = True Then
    '                    If JUDGXY = False Then
    '                        funChkFurikae1_2 = 1
    '                        iErr_Code = 1201
    '                        sErr_Msg = "CHECK1-2,結晶面傾中心,Ｘ線実績が範囲外の為、振替できません。"
    '                        gsTbcmy028ErrCode = "00004"
    '                        GoTo Apl_Exit
    '                    End If
    '                End If
    '            End If
    '        Next
    '    End If
    'End If
    ''2010/04/16 結晶面傾中心仕様判断条件の追加 END №100087 Kameda
    
    'ドーパントのチェック
    sErr_Msg = "1-2 ﾄﾞｰﾊﾟﾝﾄﾁｪｯｸ"
    If Trim$(tbl_chk1_2(0).HSXDOP) <> Trim$(tbl_chk1_2(1).HSXDOP) Then
        If Trim$(tbl_chk1_2(1).HSXDOP) <> "Z" Then      '不問品番への振替はOK 2011/05/12 SETsw kubota
            funChkFurikae1_2 = 1
            iErr_Code = 1203
            sErr_Msg = "CHECK1-2,ドーパント不一致の為、振替できません。"
    '--------------- 2008/07/25 INSERT START  By Systech ---------------
            gsTbcmy028ErrCode = "00005"
    '--------------- 2008/07/25 INSERT  END   By Systech ---------------
            GoTo Apl_Exit
        End If
    End If
    '結晶ドープのチェック
    sErr_Msg = "1-2 結晶ﾄﾞｰﾌﾟﾁｪｯｸ"
''    If Trim$(tbl_chk1_2(0).HWFCDOP) <> Trim$(tbl_chk1_2(1).HWFCDOP) Then
''        funChkFurikae1_2 = 1
''        iErr_Code = 1204
''        sErr_Msg = "CHECK1-2,結晶ドープ不一致の為、振替できません。"
''        GoTo Apl_Exit
''    End If
'' add start 0108

    '' 区分判断
    sClass = ""
    '' N振替可/C振替可
    If tbl_chk1_2(1).NDOPHUFLG = "0" And tbl_chk1_2(1).CDOPHUFLG = "0" Then
        sClass = "D0"
    '' N振替可/C振替不可
    ElseIf tbl_chk1_2(1).NDOPHUFLG = "0" And tbl_chk1_2(1).CDOPHUFLG <> "0" Then
        sClass = "D1"
    '' N振替不可/C振替可
    ElseIf tbl_chk1_2(1).NDOPHUFLG <> "0" And tbl_chk1_2(1).CDOPHUFLG = "0" Then
        sClass = "D2"
    '' N振替不可/C振替不可
    ElseIf tbl_chk1_2(1).NDOPHUFLG <> "0" And tbl_chk1_2(1).CDOPHUFLG <> "0" Then
        sClass = "D3"
    End If
'' add end 0108
    
    '06/10/17 ooba START =====================================================================>
    sResult = ""
    RET = funCodeDBGet("SB", sClass, tbl_chk1_2(0).HWFCDOP, 1, tbl_chk1_2(1).HWFCDOP, sResult)  '' chg 0108
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→元:" & tbl_chk1_2(0).HWFCDOP & ", 先:" & tbl_chk1_2(1).HWFCDOP
        GoTo CodeDBGet_Error
    End If
    If sResult = 0 Then
        funChkFurikae1_2 = 1
        iErr_Code = 1204
        sErr_Msg = "CHECK1-2,結晶ドープ、振替不可能です。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00006"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    '06/10/17 ooba END =======================================================================>

'シード傾きのチェックを1-2から1-1へ移動 2009/08/06 SETsw kubota
'    'シード傾きのチェック
'    sErr_Msg = "1-2 ｼｰﾄﾞ傾きﾁｪｯｸ"
'    If Trim$(tbl_chk1_2(0).HSXSDSLP) <> Trim$(tbl_chk1_2(1).HSXSDSLP) Then
'        funChkFurikae1_2 = 1
'        iErr_Code = 1205
'        sErr_Msg = "CHECK1-2,シード傾き不一致の為、振替できません。"
''--------------- 2008/07/25 INSERT START  By Systech ---------------
'        gsTbcmy028ErrCode = "00007"
''--------------- 2008/07/25 INSERT  END   By Systech ---------------
'        GoTo Apl_Exit
'    End If

    '溝位置方位のチェック（同一分類グループなら振替可能）
    sErr_Msg = "1-2 溝位置方位ﾁｪｯｸ"
    sResult = ""
'>>>>> CC310ノッチ方位変更の振替ロック解除 2011/10/13 SETsw kubota -----------------
    If sPlshMeth <> "M" Then     'M:MGRの場合、溝位置方位ロック解除
'<<<<< CC310ノッチ方位変更の振替ロック解除 2011/10/13 SETsw kubota -----------------
        RET = funCodeDBGet("SB", "MZ", tbl_chk1_2(0).HSXDPDIR, 1, tbl_chk1_2(1).HSXDPDIR, sResult)
        If RET <> 0 Then
            sErr_Msg = sErr_Msg & "→元:" & tbl_chk1_2(0).HSXDPDIR & ", 先:" & tbl_chk1_2(1).HSXDPDIR
            GoTo CodeDBGet_Error
        End If
        If sResult = 0 Then
            funChkFurikae1_2 = 1
            iErr_Code = 1206
            sErr_Msg = "CHECK1-2,溝位置方位、振替不可能です。"
    '--------------- 2008/07/25 INSERT START  By Systech ---------------
            gsTbcmy028ErrCode = "00008"
    '--------------- 2008/07/25 INSERT  END   By Systech ---------------
            GoTo Apl_Exit
        End If
'>>>>> CC310ノッチ方位変更の振替ロック解除 2011/10/13 SETsw kubota -----------------
    End If
'<<<<< CC310ノッチ方位変更の振替ロック解除 2011/10/13 SETsw kubota -----------------
    
    '品種のチェック
    sErr_Msg = "1-2 品種ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "HS", tbl_chk1_2(0).MCNO1, 1, tbl_chk1_2(1).MCNO1, sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→元:" & tbl_chk1_2(0).MCNO1 & ", 先:" & tbl_chk1_2(1).MCNO1
        GoTo CodeDBGet_Error
    End If
    If sResult = 0 Then
        funChkFurikae1_2 = 1
        iErr_Code = 1207
        sErr_Msg = "CHECK1-2,品種、振替不可能です。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00010"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    '引上げ速度
    sErr_Msg = "1-2 引上げ速度ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "HK", tbl_chk1_2(0).MCNO2, 1, tbl_chk1_2(1).MCNO2, sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→元:" & tbl_chk1_2(0).MCNO2 & ", 先:" & tbl_chk1_2(1).MCNO2
        GoTo CodeDBGet_Error
    End If
    If sResult = 0 Then
        funChkFurikae1_2 = 1
        iErr_Code = 1208
        sErr_Msg = "CHECK1-2,引上げ速度、振替不可能です。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00011"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'ＨＺタイプチェック
    sErr_Msg = "1-2 HZﾀｲﾌﾟﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "HZ", tbl_chk1_2(0).MCNO3, 1, tbl_chk1_2(1).MCNO3, sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→元:" & tbl_chk1_2(0).MCNO3 & ", 先:" & tbl_chk1_2(1).MCNO3
        GoTo CodeDBGet_Error
    End If
    If sResult = 0 Then
        funChkFurikae1_2 = 1
        iErr_Code = 1209
        sErr_Msg = "CHECK1-2,ＨＺタイプ、振替不可能です。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00012"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'ドローチューブチェック
    sErr_Msg = "1-2 ﾄﾞﾛｰﾁｭｰﾌﾞﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "DC", tbl_chk1_2(0).DCHYUUBU, 1, tbl_chk1_2(1).DCHYUUBU, sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→元:" & tbl_chk1_2(0).DCHYUUBU & ", 先:" & tbl_chk1_2(1).DCHYUUBU
        GoTo CodeDBGet_Error
    End If
    If sResult = 0 Then
        funChkFurikae1_2 = 1
        iErr_Code = 1210
        sErr_Msg = "CHECK1-2,ドローチューブ、振替不可能です。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00009"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'SIRD保証方法のチェック 2010/05/24 Y.Hitomi Add
    sErr_Msg = "1-2 SIRD保証方法ﾁｪｯｸ"
    If Trim$(tbl_chk1_2(0).HWFSIRDHS) = "" Then
        If Trim$(tbl_chk1_2(1).HWFSIRDHS) = "S" Or Trim$(tbl_chk1_2(1).HWFSIRDHS) = "H" Then
            funChkFurikae1_2 = 1
            iErr_Code = 1211
            sErr_Msg = "CHECK1-2,振替先がSIRD保証不一致の為,振替できません"
            gsTbcmy028ErrCode = "00013"
        GoTo Apl_Exit
        End If
    End If
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae1_2 = 0 Then
        funChkFurikae1_2 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae1_2 = -4
    GoTo Apl_Exit
    
CodeDBGet_Error:
    Set rs = Nothing
    If funChkFurikae1_2 = 0 Then
        funChkFurikae1_2 = -5
    End If
    GoTo Apl_Exit

End Function

    
'------------------------------------------------
' 外観実績を振替先品番でチェック
'------------------------------------------------

'概要      :振替先品番の外観実績（直径、溝巾、溝深）を研削加工実績(TBCMI002)から取得する。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sKeyID          ,I  ,String       :ﾌﾞﾛｯｸID、又は、SXL-ID
'          :tOld_Hinban     ,I  ,tFullHinban  :振替元品番(構造体)
'          :tNew_Hinban     ,I  ,tFullHinban  :振替先品番(構造体)
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/19 新規作成　SB

Public Function funChkFurikae1_3(sProccd As String, sBlockId As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iErr_Code As Integer, sErr_Msg As String) As Integer
    
    Dim RET         As Integer          '戻り値
    Dim sResult     As String           'コードＤＢ取得関数の取得変数
    Dim sql         As String           'SQL全体
    Dim rs          As OraDynaset       'RecordSet
    Dim wBLKID()    As String
    Dim Jiltuseki   As Judg_Kakou
    Dim W_AVG       As Double
    Dim cnt         As Integer
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae1_3 = 0
    
    '------------------------------------------ 振替元品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-3 振替元品番仕様取得(" & tOld_Hinban.hinban & Format(tOld_Hinban.mnorevno, "00") & tOld_Hinban.factory & tOld_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E018.HSXD1MIN,E018.HSXD1MAX,E018.HSXDWMIN,E018.HSXDWMAX,E018.HSXDDMIN,E018.HSXDDMAX,E027.HWFWARPR " & vbCrLf
    sql = sql & "FROM   TBCME018 E018,TBCME027 E027 " & vbCrLf
    sql = sql & "WHERE  E018.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E018.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E018.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E027.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E027.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E027.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E027.OPECOND   =   '" & tOld_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    Erase tbl_chk1_3
    With tbl_chk1_3(0)
        If IsNull(rs("HSXD1MIN")) = False Then .HSXD1MIN = rs("HSXD1MIN") Else .HSXD1MIN = 0        '品ＳＸ直径１下限
        If IsNull(rs("HSXD1MAX")) = False Then .HSXD1MAX = rs("HSXD1MAX") Else .HSXD1MAX = 0        '品ＳＸ直径１上限
        If IsNull(rs("HSXDWMIN")) = False Then .HSXDWMIN = rs("HSXDWMIN") Else .HSXDWMIN = 0        '品ＳＸ溝巾下限
        If IsNull(rs("HSXDWMAX")) = False Then .HSXDWMAX = rs("HSXDWMAX") Else .HSXDWMAX = 0        '品ＳＸ溝巾上限
        If IsNull(rs("HSXDDMIN")) = False Then .HSXDDMIN = rs("HSXDDMIN") Else .HSXDDMIN = 0        '品ＳＸ溝深下限
        If IsNull(rs("HSXDDMAX")) = False Then .HSXDDMAX = rs("HSXDDMAX") Else .HSXDDMAX = 0        '品ＳＸ溝深上限
        If IsNull(rs("HWFWARPR")) = False Then .HWFWARPR = rs("HWFWARPR") Else .HWFWARPR = "1"      'Warpランク
    End With
    
    Set rs = Nothing
    '------------------------------------------ 振替先品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-3 振替先品番仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E018.HSXD1MIN,E018.HSXD1MAX,E018.HSXDWMIN,E018.HSXDWMAX,E018.HSXDDMIN,E018.HSXDDMAX,E027.HWFWARPR " & vbCrLf
    sql = sql & "FROM   TBCME018 E018,TBCME027 E027 " & vbCrLf
    sql = sql & "WHERE  E018.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E018.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E018.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E027.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E027.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E027.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E027.OPECOND   =   '" & tNew_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    With tbl_chk1_3(1)
        If IsNull(rs("HSXD1MIN")) = False Then .HSXD1MIN = rs("HSXD1MIN") Else .HSXD1MIN = 0        '品ＳＸ直径１下限
        If IsNull(rs("HSXD1MAX")) = False Then .HSXD1MAX = rs("HSXD1MAX") Else .HSXD1MAX = 0        '品ＳＸ直径１上限
        If IsNull(rs("HSXDWMIN")) = False Then .HSXDWMIN = rs("HSXDWMIN") Else .HSXDWMIN = 0        '品ＳＸ溝巾下限
        If IsNull(rs("HSXDWMAX")) = False Then .HSXDWMAX = rs("HSXDWMAX") Else .HSXDWMAX = 0        '品ＳＸ溝巾上限
        If IsNull(rs("HSXDDMIN")) = False Then .HSXDDMIN = rs("HSXDDMIN") Else .HSXDDMIN = 0        '品ＳＸ溝深下限
        If IsNull(rs("HSXDDMAX")) = False Then .HSXDDMAX = rs("HSXDDMAX") Else .HSXDDMAX = 0        '品ＳＸ溝深上限
        If IsNull(rs("HWFWARPR")) = False Then .HWFWARPR = rs("HWFWARPR") Else .HWFWARPR = "1"      'Warpランク
    End With
    
    Set rs = Nothing
    '------------------------------------------ 指示取得 ------------------------------------------------------
    On Error GoTo Apl_down
    '振替先品番の外観実績（直径、溝巾、溝深）の取得
    
    'CW750,CW760の場合、ﾌﾞﾛｯｸIDを取得
    If (left(sProccd, 4) = "CW75") Or (left(sProccd, 4) = "CW76") Then
        sErr_Msg = "1-3 外観実績BLK取得"
        sql = vbNullString
        sql = sql & "SELECT CRYNUMCA FROM XSDCA " & vbCrLf
        sql = sql & "WHERE  SXLIDCA = '" & sBlockId & "' AND " & vbCrLf
        sql = sql & "       LIVKCA  = '0' " & vbCrLf
    
        On Error GoTo db_Error
        'SQL文の実行
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        
        '該当データなし
        If rs.EOF Or rs.RecordCount = 0 Then GoTo db_Error
        
        ReDim wBLKID(rs.RecordCount)
        For cnt = 1 To rs.RecordCount
            wBLKID(cnt) = rs("CRYNUMCA")
            rs.MoveNext
        Next
        Set rs = Nothing
    Else
        ReDim wBLKID(1)
        wBLKID(1) = sBlockId
    End If
    
    For cnt = 1 To UBound(wBLKID)
        sErr_Msg = "1-3 外観実績取得"
        RET = scmzc_getKakouJiltuseki(wBLKID(cnt), Jiltuseki)
        If RET <> 0 Then
'            funChkFurikae1_3 = -2
            funChkFurikae1_3 = 1
            iErr_Code = 1305
            sErr_Msg = "CHECK1-3,外観実績取得エラー"
            GoTo Apl_Exit
        End If
        '直径実績のチェック
        sErr_Msg = "1-3 直径実績ﾁｪｯｸ"
        W_AVG = Jiltuseki.top(1) + Jiltuseki.top(2) + Jiltuseki.TAIL(1) + Jiltuseki.TAIL(2)
        W_AVG = W_AVG / 4#
        If tbl_chk1_3(1).HSXD1MIN <= W_AVG And _
           tbl_chk1_3(1).HSXD1MAX >= W_AVG Then
        Else
            funChkFurikae1_3 = 1
            iErr_Code = 1301
            sErr_Msg = "CHECK1-3,直径実績が仕様範囲外の為、振替できません。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
            gsTbcmy028ErrCode = "00016"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            GoTo Apl_Exit
        End If
        '溝巾実績のチェック
        sErr_Msg = "1-3 溝巾実績ﾁｪｯｸ"
        If tbl_chk1_3(1).HSXDWMIN <= Jiltuseki.WIDH(1) And _
           tbl_chk1_3(1).HSXDWMAX >= Jiltuseki.WIDH(1) Then
        Else
            funChkFurikae1_3 = 1
            iErr_Code = 1302
            sErr_Msg = "CHECK1-3,溝巾実績が仕様範囲外の為、振替できません。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
            gsTbcmy028ErrCode = "00017"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            GoTo Apl_Exit
        End If
        '溝深実績のチェック
        sErr_Msg = "1-3 溝深実績ﾁｪｯｸ"
        If tbl_chk1_3(1).HSXDDMIN <= Jiltuseki.DPTH(1) And _
           tbl_chk1_3(1).HSXDDMAX >= Jiltuseki.DPTH(1) Then
        Else
            funChkFurikae1_3 = 1
            iErr_Code = 1303
            sErr_Msg = "CHECK1-3,溝深実績が仕様範囲外の為、振替できません。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
            gsTbcmy028ErrCode = "00018"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            GoTo Apl_Exit
        End If
    Next
    
    'Warpランク
    sErr_Msg = "1-3 ﾜｰﾌﾟﾗﾝｸﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "WR", tbl_chk1_3(0).HWFWARPR, 1, tbl_chk1_3(1).HWFWARPR, sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→元:" & tbl_chk1_3(0).HWFWARPR & ", 先:" & tbl_chk1_3(1).HWFWARPR
        GoTo CodeDBGet_Error
    End If
    If sResult = 0 Then
        funChkFurikae1_3 = 1
        iErr_Code = 1304
        sErr_Msg = "CHECK1-3,ワープランク、振替不可能です。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00015"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae1_3 = 0 Then
        funChkFurikae1_3 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae1_3 = -4
    GoTo Apl_Exit
    
CodeDBGet_Error:
    Set rs = Nothing
    If funChkFurikae1_3 = 0 Then
        funChkFurikae1_3 = -5
    End If
    GoTo Apl_Exit

End Function

    
'------------------------------------------------
' 振替元と振替先の結晶評価項目仕様チェック
'------------------------------------------------

'概要      :振替元品番と振替先品番の結晶評価項目仕様をチェックする。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sKeyID          ,I  ,String       :ﾌﾞﾛｯｸID、又は、SXL-ID
'          :tOld_Hinban     ,I  ,tFullHinban  :振替元品番(構造体)
'          :tNew_Hinban     ,I  ,tFullHinban  :振替先品番(構造体)
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :iELCs_Flg       ,O  ,Integer      :0 ･･･ 1-4全項目チェック
'                                              1 ･･･ 1-4(Cs,EPD,LT)のみチェック
'                                              2 ･･･ 1-4(Cs,EPD,LT)以外チェック
'                                              3 ･･･ 1-4(Cs)のみチェック
'                                              4 ･･･ 1-4(EPD)のみチェック
'                                              5 ･･･ 1-4(LT)のみチェック
'履歴      :2003/09/19 新規作成　SB
''            2006/05/09 SMP桜井　複数品番判定対応>>>--変更前
''Public Function funChkFurikae1_4(sProccd As String, sBlockId As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
''                                 iErr_Code As Integer, sErr_Msg As String) As Integer<<<

Public Function funChkFurikae1_4(sProccd As String, sBlockId As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iErr_Code As Integer, sErr_Msg As String, Optional iELCs_Flg As Integer = 0) As Integer

'<<<<<複数品番判定対応
    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql As String               'SQL全体
    Dim rs  As OraDynaset           'RecordSet
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae1_4 = 0
    
    '------------------------------------------ 振替元品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-4 振替元品番仕様取得(" & tOld_Hinban.hinban & Format(tOld_Hinban.mnorevno, "00") & tOld_Hinban.factory & tOld_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E018.HSXRHWYS,E019.HSXONHWS,  E019.HSXONSPT,E019.HSXONSPI,E019.HSXONKWY,E020.HSXOF1HS,E020.HSXOF1SH,  E020.HSXOF1ST,E020.HSXOF1SR,  E020.HSXOF1NS,E020.HSXOF1SZ,   " & vbCrLf
    sql = sql & "       E020.HSXOF1ET,E020.HSXOSF1PTK,E020.HSXOF2HS,E020.HSXOF2SH,E020.HSXOF2ST,E020.HSXOF2SR,E020.HSXOF2NS,  E020.HSXOF2SZ,  E020.HSXOF2ET,E020.HSXOSF2PTK, " & vbCrLf
    sql = sql & "       E020.HSXOF3HS,E020.HSXOF3SH,  E020.HSXOF3ST,E020.HSXOF3SR,E020.HSXOF3NS,E020.HSXOF3SZ,  E020.HSXOF3ET,E020.HSXOSF3PTK,E020.HSXOF4HS,E020.HSXOF4SH,   " & vbCrLf
    sql = sql & "       E020.HSXOF4ST,E020.HSXOF4SR,  E020.HSXOF4NS,E020.HSXOF4SZ,E020.HSXOF4ET,E020.HSXOSF4PTK,E020.HSXBM1HS,E020.HSXBM1SH,  E020.HSXBM1ST,E020.HSXBM1SR,   " & vbCrLf
    sql = sql & "       E020.HSXBM1NS,E020.HSXBM1SZ,  E020.HSXBM1ET,E020.HSXBM2HS,E020.HSXBM2SH,E020.HSXBM2ST,  E020.HSXBM2SR,E020.HSXBM2NS,  E020.HSXBM2SZ,E020.HSXBM2ET,   " & vbCrLf
    sql = sql & "       E020.HSXBM3HS,E020.HSXBM3SH,  E020.HSXBM3ST,E020.HSXBM3SR,E020.HSXBM3NS,E020.HSXBM3SZ,  E020.HSXBM3ET,E019.HSXTMMAX,  E019.HSXLTHWS,E019.HSXCNHWS,   " & vbCrLf
'*** UPDATE ↓ Y.SIMIZU 2005/10/12 GDﾗｲﾝ数取得追加
'    sql = sql & "       E019.HSXCNKWY,E020.HSXDENHS,  E020.HSXDENMN,E020.HSXDENMX,E020.HSXDVDHS,E020.HSXDVDMNN,  E020.HSXDVDMXN,E020.HSXLDLHS,  E020.HSXLDLMN,E020.HSXLDLMX    " & vbCrLf
'    sql = sql & "FROM   TBCME018 E018,TBCME019 E019,TBCME020 E020 " & vbCrLf
'    sql = sql & "WHERE  E018.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
'    sql = sql & "       E018.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
'    sql = sql & "       E018.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
'    sql = sql & "       E018.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
'    sql = sql & "       E019.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
'    sql = sql & "       E019.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
'    sql = sql & "       E019.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
'    sql = sql & "       E019.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
'    sql = sql & "       E020.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
'    sql = sql & "       E020.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
'    sql = sql & "       E020.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
'    sql = sql & "       E020.OPECOND   =   '" & tOld_Hinban.opecond & "' " & vbCrLf
'C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
    sql = sql & "       E019.HSXCNKWY,E020.HSXDENHS,  E020.HSXDENMN,E020.HSXDENMX,E020.HSXDVDHS,E020.HSXDVDMNN,  E020.HSXDVDMXN,E020.HSXLDLHS,  E020.HSXLDLMN,E020.HSXLDLMX,E036.HSXGDLINE,E036.COSF3FLAG " & vbCrLf
'C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---
'--------------- 2008/08/25 INSERT START  By Systech ---------------
'    sql = sql & "       ,NVL(E036.HSXDKTMP,' ') HSXDKTMP " & vbCrLf
    '08/12/21 ooba
    sql = sql & "       ,NVL(E036.HSXDKTMP,' ') HSXDKTMP, E036.HSXOF1ARPTK " & vbCrLf
    sql = sql & "       ,E019.HSXCNKHI " & vbCrLf   '' add 0108
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    sql = sql & "       ,E048.HWFSIRDMX " & vbCrLf                   '軸状転位上限
    sql = sql & "       ,E048.HWFSIRDSZ " & vbCrLf                   '軸状転位測定条件
    sql = sql & "       ,E048.HWFSIRDHT " & vbCrLf                   '軸状転位保証方法＿対
    sql = sql & "       ,E048.HWFSIRDHS " & vbCrLf                   '軸状転位保証方法_処
    sql = sql & "       ,E048.HWFSIRDKM " & vbCrLf                   '軸状転位検査頻度＿枚
    sql = sql & "       ,E048.HWFSIRDKN " & vbCrLf                   '軸状転位検査頻度_抜
    sql = sql & "       ,E048.HWFSIRDKH " & vbCrLf                   '軸状転位検査頻度＿保
    sql = sql & "       ,E048.HWFSIRDKU " & vbCrLf                   '軸状転位検査頻度＿ウ
    sql = sql & "       ,E048.HWFSIRDPS " & vbCrLf                   '軸状転位TB保証位置
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    sql = sql & "FROM   TBCME018 E018,TBCME019 E019,TBCME020 E020,TBCME036 E036 " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    sql = sql & "       ,TBCME048 E048  " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
    sql = sql & "WHERE  E018.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E018.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E018.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E019.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E019.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E019.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E019.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E020.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E020.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E020.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E020.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E036.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tOld_Hinban.opecond & "'     " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    sql = sql & "   AND E048.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E048.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E048.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E048.OPECOND   =   '" & tOld_Hinban.opecond & "' " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
'*** UPDATE ↑ Y.SIMIZU 2005/10/12 GDﾗｲﾝ数取得追加
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    Erase tbl_chk1_4
    With tbl_chk1_4(0)
        'Rs
        If IsNull(rs("HSXRHWYS")) = False Then .HSXRHWYS = rs("HSXRHWYS") Else .HSXRHWYS = " "              '保証方法_対象
        'Oi
        If IsNull(rs("HSXONHWS")) = False Then .HSXONHWS = rs("HSXONHWS") Else .HSXONHWS = " "              '保証方法_対象
        If IsNull(rs("HSXONSPT")) = False Then .HSXONSPT = rs("HSXONSPT") Else .HSXONSPT = " "              '測定位置_点    '08/01/29 ooba
        If IsNull(rs("HSXONSPI")) = False Then .HSXONSPI = rs("HSXONSPI") Else .HSXONSPI = " "              '測定位置_位
        If IsNull(rs("HSXONKWY")) = False Then .HSXONKWY = rs("HSXONKWY") Else .HSXONKWY = " "              '検査方法
        'OSF1
        If IsNull(rs("HSXOF1HS")) = False Then .HSXOF1HS = rs("HSXOF1HS") Else .HSXOF1HS = " "              '保証方法_対象
        If IsNull(rs("HSXOF1SH")) = False Then .HSXOF1SH = rs("HSXOF1SH") Else .HSXOF1SH = " "              '測定位置_方
        If IsNull(rs("HSXOF1ST")) = False Then .HSXOF1ST = rs("HSXOF1ST") Else .HSXOF1ST = " "              '測定位置_点
        If IsNull(rs("HSXOF1SR")) = False Then .HSXOF1SR = rs("HSXOF1SR") Else .HSXOF1SR = " "              '測定位置_領
        If IsNull(rs("HSXOF1NS")) = False Then .HSXOF1NS = rs("HSXOF1NS") Else .HSXOF1NS = " "              '熱処理法
        If IsNull(rs("HSXOF1SZ")) = False Then .HSXOF1SZ = rs("HSXOF1SZ") Else .HSXOF1SZ = " "              '測定条件
        If IsNull(rs("HSXOF1ET")) = False Then .HSXOF1ET = rs("HSXOF1ET") Else .HSXOF1ET = 0                '選択ET代
        If IsNull(rs("HSXOSF1PTK")) = False Then .HSXOSF1PTK = rs("HSXOSF1PTK") Else .HSXOSF1PTK = "4"      'パターン区分
        'OSF2
        If IsNull(rs("HSXOF2HS")) = False Then .HSXOF2HS = rs("HSXOF2HS") Else .HSXOF2HS = " "              '保証方法_対象
        If IsNull(rs("HSXOF2SH")) = False Then .HSXOF2SH = rs("HSXOF2SH") Else .HSXOF2SH = " "              '測定位置_方
        If IsNull(rs("HSXOF2ST")) = False Then .HSXOF2ST = rs("HSXOF2ST") Else .HSXOF2ST = " "              '測定位置_点
        If IsNull(rs("HSXOF2SR")) = False Then .HSXOF2SR = rs("HSXOF2SR") Else .HSXOF2SR = " "              '測定位置_領
        If IsNull(rs("HSXOF2NS")) = False Then .HSXOF2NS = rs("HSXOF2NS") Else .HSXOF2NS = " "              '熱処理法
        If IsNull(rs("HSXOF2SZ")) = False Then .HSXOF2SZ = rs("HSXOF2SZ") Else .HSXOF2SZ = " "              '測定条件
        If IsNull(rs("HSXOF2ET")) = False Then .HSXOF2ET = rs("HSXOF2ET") Else .HSXOF2ET = 0                '選択ET代
        If IsNull(rs("HSXOSF2PTK")) = False Then .HSXOSF2PTK = rs("HSXOSF2PTK") Else .HSXOSF2PTK = "4"      'パターン区分
        'OSF3
        If IsNull(rs("HSXOF3HS")) = False Then .HSXOF3HS = rs("HSXOF3HS") Else .HSXOF3HS = " "              '保証方法_対象
        If IsNull(rs("HSXOF3SH")) = False Then .HSXOF3SH = rs("HSXOF3SH") Else .HSXOF3SH = " "              '測定位置_方
        If IsNull(rs("HSXOF3ST")) = False Then .HSXOF3ST = rs("HSXOF3ST") Else .HSXOF3ST = " "              '測定位置_点
        If IsNull(rs("HSXOF3SR")) = False Then .HSXOF3SR = rs("HSXOF3SR") Else .HSXOF3SR = " "              '測定位置_領
        If IsNull(rs("HSXOF3NS")) = False Then .HSXOF3NS = rs("HSXOF3NS") Else .HSXOF3NS = " "              '熱処理法
        If IsNull(rs("HSXOF3SZ")) = False Then .HSXOF3SZ = rs("HSXOF3SZ") Else .HSXOF3SZ = " "              '測定条件
        If IsNull(rs("HSXOF3ET")) = False Then .HSXOF3ET = rs("HSXOF3ET") Else .HSXOF3ET = 0                '選択ET代
        If IsNull(rs("HSXOSF3PTK")) = False Then .HSXOSF3PTK = rs("HSXOSF3PTK") Else .HSXOSF3PTK = "4"      'パターン区分


''C－OSF3チェックの変更 2008.04.20 青柳
'''C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
''        'If IsNull(rs("HSXOF4HS")) = False Then .HSXOF4HS = rs("HSXOF4HS") Else .HSXOF4HS = " "             '保証方法_対象
''        If IsNull(rs("COSF3FLAG")) = False Then .HSXOF4HS = rs("COSF3FLAG") Else .HSXOF4HS = " "            'C-OSF3ﾌﾗｸﾞ
'''C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---
        
'--------------- 2008/08/25 INSERT START  By Systech ---------------
        .HSXDKTMP = rs("HSXDKTMP")
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
        
        'OSF4
'        If IsNull(rs("HSXOF4SH")) = False Then .HSXOF4SH = rs("HSXOF4SH") Else .HSXOF4SH = " "              '測定位置_方
'        If IsNull(rs("HSXOF4ST")) = False Then .HSXOF4ST = rs("HSXOF4ST") Else .HSXOF4ST = " "              '測定位置_点
'        If IsNull(rs("HSXOF4SR")) = False Then .HSXOF4SR = rs("HSXOF4SR") Else .HSXOF4SR = " "              '測定位置_領
'        If IsNull(rs("HSXOF4NS")) = False Then .HSXOF4NS = rs("HSXOF4NS") Else .HSXOF4NS = " "              '熱処理法
'        If IsNull(rs("HSXOF4SZ")) = False Then .HSXOF4SZ = rs("HSXOF4SZ") Else .HSXOF4SZ = " "              '測定条件
'        If IsNull(rs("HSXOF4ET")) = False Then .HSXOF4ET = rs("HSXOF4ET") Else .HSXOF4ET = 0                '選択ET代
'        If IsNull(rs("HSXOSF4PTK")) = False Then .HSXOSF4PTK = rs("HSXOSF4PTK") Else .HSXOSF4PTK = "4"      'パターン区分
        If IsNull(rs("HSXOF1HS")) = False Then .HSXOF4HS = rs("HSXOF1HS") Else .HSXOF4HS = " "              '保証方法_対象(ArANはOSF1保証) 08/12/21 ooba
        If IsNull(rs("HSXOF1ARPTK")) = False Then .HSXOSF4PTK = rs("HSXOF1ARPTK") Else .HSXOSF4PTK = " "    '(ArAN)ﾊﾟﾀｰﾝ区分 08/12/21 ooba
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
        'SIRD
        If IsNull(rs("HWFSIRDMX")) = False Then .HWFSIRDMX = rs("HWFSIRDMX") Else .HWFSIRDMX = "0"          '軸状転位上限
        If IsNull(rs("HWFSIRDSZ")) = False Then .HWFSIRDSZ = rs("HWFSIRDSZ") Else .HWFSIRDSZ = " "          '軸状転位測定条件
        If IsNull(rs("HWFSIRDHT")) = False Then .HWFSIRDHT = rs("HWFSIRDHT") Else .HWFSIRDHT = " "          '軸状転位保証方法＿対
        If IsNull(rs("HWFSIRDHS")) = False Then .HWFSIRDHS = rs("HWFSIRDHS") Else .HWFSIRDHS = " "          '軸状転位保証方法＿処
        If IsNull(rs("HWFSIRDKM")) = False Then .HWFSIRDKM = rs("HWFSIRDKM") Else .HWFSIRDKM = " "          '軸状転位検査頻度＿枚
        If IsNull(rs("HWFSIRDKN")) = False Then .HWFSIRDKN = rs("HWFSIRDKN") Else .HWFSIRDKN = " "          '軸状転位検査頻度＿抜
        If IsNull(rs("HWFSIRDKH")) = False Then .HWFSIRDKH = rs("HWFSIRDKH") Else .HWFSIRDKH = " "          '軸状転位検査頻度＿保
        If IsNull(rs("HWFSIRDKU")) = False Then .HWFSIRDKU = rs("HWFSIRDKU") Else .HWFSIRDKU = " "          '軸状転位検査頻度＿ウ
        If IsNull(rs("HWFSIRDPS")) = False Then .HWFSIRDPS = Trim(rs("HWFSIRDPS")) Else .HWFSIRDPS = " "    '軸状転位TB保証位置
        
        '「軸状転位TB保証位置」を判定し、「軸状転位検査頻度＿抜」に編集
        Select Case Trim(.HWFSIRDPS)
        Case "T"
            .HWFSIRDKN = "3"
        Case "B"
            .HWFSIRDKN = "4"
        Case "TB"
            .HWFSIRDKN = "6"
        Case Else
            .HWFSIRDKN = " "
        End Select
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
        'BMD1
        If IsNull(rs("HSXBM1HS")) = False Then .HSXBM1HS = rs("HSXBM1HS") Else .HSXBM1HS = " "              '保証方法_対象
        If IsNull(rs("HSXBM1SH")) = False Then .HSXBM1SH = rs("HSXBM1SH") Else .HSXBM1SH = " "              '測定位置_方
        If IsNull(rs("HSXBM1ST")) = False Then .HSXBM1ST = rs("HSXBM1ST") Else .HSXBM1ST = " "              '測定位置_点
        If IsNull(rs("HSXBM1SR")) = False Then .HSXBM1SR = rs("HSXBM1SR") Else .HSXBM1SR = " "              '測定位置_領
        If IsNull(rs("HSXBM1NS")) = False Then .HSXBM1NS = rs("HSXBM1NS") Else .HSXBM1NS = " "              '熱処理法
        If IsNull(rs("HSXBM1SZ")) = False Then .HSXBM1SZ = rs("HSXBM1SZ") Else .HSXBM1SZ = " "              '測定条件
        If IsNull(rs("HSXBM1ET")) = False Then .HSXBM1ET = rs("HSXBM1ET") Else .HSXBM1ET = 0                '選択ET代
        'BMD2
        If IsNull(rs("HSXBM2HS")) = False Then .HSXBM2HS = rs("HSXBM2HS") Else .HSXBM2HS = " "              '保証方法_対象
        If IsNull(rs("HSXBM2SH")) = False Then .HSXBM2SH = rs("HSXBM2SH") Else .HSXBM2SH = " "              '測定位置_方
        If IsNull(rs("HSXBM2ST")) = False Then .HSXBM2ST = rs("HSXBM2ST") Else .HSXBM2ST = " "              '測定位置_点
        If IsNull(rs("HSXBM2SR")) = False Then .HSXBM2SR = rs("HSXBM2SR") Else .HSXBM2SR = " "              '測定位置_領
        If IsNull(rs("HSXBM2NS")) = False Then .HSXBM2NS = rs("HSXBM2NS") Else .HSXBM2NS = " "              '熱処理法
        If IsNull(rs("HSXBM2SZ")) = False Then .HSXBM2SZ = rs("HSXBM2SZ") Else .HSXBM2SZ = " "              '測定条件
        If IsNull(rs("HSXBM2ET")) = False Then .HSXBM2ET = rs("HSXBM2ET") Else .HSXBM2ET = 0                '選択ET代
        'BMD3
        If IsNull(rs("HSXBM3HS")) = False Then .HSXBM3HS = rs("HSXBM3HS") Else .HSXBM3HS = " "              '保証方法_対象
        If IsNull(rs("HSXBM3SH")) = False Then .HSXBM3SH = rs("HSXBM3SH") Else .HSXBM3SH = " "              '測定位置_方
        If IsNull(rs("HSXBM3ST")) = False Then .HSXBM3ST = rs("HSXBM3ST") Else .HSXBM3ST = " "              '測定位置_点
        If IsNull(rs("HSXBM3SR")) = False Then .HSXBM3SR = rs("HSXBM3SR") Else .HSXBM3SR = " "              '測定位置_領
        If IsNull(rs("HSXBM3NS")) = False Then .HSXBM3NS = rs("HSXBM3NS") Else .HSXBM3NS = " "              '熱処理法
        If IsNull(rs("HSXBM3SZ")) = False Then .HSXBM3SZ = rs("HSXBM3SZ") Else .HSXBM3SZ = " "              '測定条件
        If IsNull(rs("HSXBM3ET")) = False Then .HSXBM3ET = rs("HSXBM3ET") Else .HSXBM3ET = 0                '選択ET代
        'EPD
        If IsNull(rs("HSXTMMAX")) = False Then .HSXTMMAX = rs("HSXTMMAX") Else .HSXTMMAX = 0                '上限
        'LT
        If IsNull(rs("HSXLTHWS")) = False Then .HSXLTHWS = rs("HSXLTHWS") Else .HSXLTHWS = " "              '保証方法_対象
        'CS
        If IsNull(rs("HSXCNHWS")) = False Then .HSXCNHWS = rs("HSXCNHWS") Else .HSXCNHWS = " "              '保証方法_対象
        If IsNull(rs("HSXCNKWY")) = False Then .HSXCNKWY = rs("HSXCNKWY") Else .HSXCNKWY = " "              '検査方法
        If IsNull(rs("HSXCNKHI")) = False Then .HSXCNKHI = rs("HSXCNKHI") Else .HSXCNKHI = " "              '検査頻度＿位   '' add 0108
        'DEN
        If IsNull(rs("HSXDENHS")) = False Then .HSXDENHS = rs("HSXDENHS") Else .HSXDENHS = " "              '保証方法_対象
        If IsNull(rs("HSXDENMN")) = False Then .HSXDENMN = rs("HSXDENMN") Else .HSXDENMN = 0                '下限
        If IsNull(rs("HSXDENMX")) = False Then .HSXDENMX = rs("HSXDENMX") Else .HSXDENMX = 0                '上限
        'DVD2
        If IsNull(rs("HSXDVDHS")) = False Then .HSXDVDHS = rs("HSXDVDHS") Else .HSXDVDHS = " "              '保証方法_対象
        If IsNull(rs("HSXDVDMNN")) = False Then .HSXDVDMNN = rs("HSXDVDMNN") Else .HSXDVDMNN = 0            '下限
        If IsNull(rs("HSXDVDMXN")) = False Then .HSXDVDMXN = rs("HSXDVDMXN") Else .HSXDVDMXN = 0            '上限
        'L/DL
        If IsNull(rs("HSXLDLHS")) = False Then .HSXLDLHS = rs("HSXLDLHS") Else .HSXLDLHS = " "              '保証方法_対象
        If IsNull(rs("HSXLDLMN")) = False Then .HSXLDLMN = rs("HSXLDLMN") Else .HSXLDLMN = 0                '下限
        If IsNull(rs("HSXLDLMX")) = False Then .HSXLDLMX = rs("HSXLDLMX") Else .HSXLDLMX = 0                '上限
    '*** UPDATE ↓ Y.SIMIZU 2005/10/12 GDﾗｲﾝ数追加
        If IsNull(rs("HSXGDLINE")) = False Then .HSXGDLINE = rs("HSXGDLINE") Else .HSXGDLINE = " "          'GDﾗｲﾝ数
    '*** UPDATE ↑ Y.SIMIZU 2005/10/12 GDﾗｲﾝ数追加
    End With
    
    Set rs = Nothing
    '------------------------------------------ 振替先品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-4 振替先品番仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E018.HSXRHWYS,E019.HSXONHWS,  E019.HSXONSPT,E019.HSXONSPI,E019.HSXONKWY,E020.HSXOF1HS,E020.HSXOF1SH,  E020.HSXOF1ST,E020.HSXOF1SR,  E020.HSXOF1NS,E020.HSXOF1SZ,   " & vbCrLf
    sql = sql & "       E020.HSXOF1ET,E020.HSXOSF1PTK,E020.HSXOF2HS,E020.HSXOF2SH,E020.HSXOF2ST,E020.HSXOF2SR,E020.HSXOF2NS,  E020.HSXOF2SZ,  E020.HSXOF2ET,E020.HSXOSF2PTK, " & vbCrLf
    sql = sql & "       E020.HSXOF3HS,E020.HSXOF3SH,  E020.HSXOF3ST,E020.HSXOF3SR,E020.HSXOF3NS,E020.HSXOF3SZ,  E020.HSXOF3ET,E020.HSXOSF3PTK,E020.HSXOF4HS,E020.HSXOF4SH,   " & vbCrLf
    sql = sql & "       E020.HSXOF4ST,E020.HSXOF4SR,  E020.HSXOF4NS,E020.HSXOF4SZ,E020.HSXOF4ET,E020.HSXOSF4PTK,E020.HSXBM1HS,E020.HSXBM1SH,  E020.HSXBM1ST,E020.HSXBM1SR,   " & vbCrLf
    sql = sql & "       E020.HSXBM1NS,E020.HSXBM1SZ,  E020.HSXBM1ET,E020.HSXBM2HS,E020.HSXBM2SH,E020.HSXBM2ST,  E020.HSXBM2SR,E020.HSXBM2NS,  E020.HSXBM2SZ,E020.HSXBM2ET,   " & vbCrLf
    sql = sql & "       E020.HSXBM3HS,E020.HSXBM3SH,  E020.HSXBM3ST,E020.HSXBM3SR,E020.HSXBM3NS,E020.HSXBM3SZ,  E020.HSXBM3ET,E019.HSXTMMAX,  E019.HSXLTHWS,E019.HSXCNHWS,   " & vbCrLf
'*** UPDATE ↓ Y.SIMIZU 2005/10/12 GDﾗｲﾝ数取得追加
'    sql = sql & "       E019.HSXCNKWY,E020.HSXDENHS,  E020.HSXDENMN,E020.HSXDENMX,E020.HSXDVDHS,E020.HSXDVDMNN, E020.HSXDVDMXN,E020.HSXLDLHS,  E020.HSXLDLMN,E020.HSXLDLMX    " & vbCrLf
'    sql = sql & "FROM   TBCME018 E018,TBCME019 E019,TBCME020 E020 " & vbCrLf
'    sql = sql & "WHERE  E018.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
'    sql = sql & "       E018.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
'    sql = sql & "       E018.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
'    sql = sql & "       E018.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
'    sql = sql & "       E019.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
'    sql = sql & "       E019.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
'    sql = sql & "       E019.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
'    sql = sql & "       E019.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
'    sql = sql & "       E020.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
'    sql = sql & "       E020.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
'    sql = sql & "       E020.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
'    sql = sql & "       E020.OPECOND   =   '" & tNew_Hinban.opecond & "' " & vbCrLf
'C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
    sql = sql & "       E019.HSXCNKWY,E020.HSXDENHS,  E020.HSXDENMN,E020.HSXDENMX,E020.HSXDVDHS,E020.HSXDVDMNN, E020.HSXDVDMXN,E020.HSXLDLHS,  E020.HSXLDLMN,E020.HSXLDLMX,E036.HSXGDLINE,E036.COSF3FLAG " & vbCrLf
'C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    sql = sql & "       ,NVL(E036.HSXDKTMP,' ') HSXDKTMP " & vbCrLf
    '08/12/21 ooba
    sql = sql & "       ,NVL(E036.HSXDKTMP,' ') HSXDKTMP, E036.HSXOF1ARPTK " & vbCrLf
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    sql = sql & "       ,E019.HSXCNKHI " & vbCrLf   '' add 0108
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    sql = sql & "       ,E048.HWFSIRDMX " & vbCrLf                   '軸状転位上限
    sql = sql & "       ,E048.HWFSIRDSZ " & vbCrLf                   '軸状転位測定条件
    sql = sql & "       ,E048.HWFSIRDHT " & vbCrLf                   '軸状転位保証方法＿対
    sql = sql & "       ,E048.HWFSIRDHS " & vbCrLf                   '軸状転位保証方法_処
    sql = sql & "       ,E048.HWFSIRDKM " & vbCrLf                   '軸状転位検査頻度＿枚
    sql = sql & "       ,E048.HWFSIRDKN " & vbCrLf                   '軸状転位検査頻度_抜
    sql = sql & "       ,E048.HWFSIRDKH " & vbCrLf                   '軸状転位検査頻度＿保
    sql = sql & "       ,E048.HWFSIRDKU " & vbCrLf                   '軸状転位検査頻度＿ウ
    sql = sql & "       ,E048.HWFSIRDPS " & vbCrLf                   '軸状転位TB保証位置
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
    sql = sql & "FROM   TBCME018 E018,TBCME019 E019,TBCME020 E020,TBCME036 E036 " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    sql = sql & "       ,TBCME048 E048  " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
    sql = sql & "WHERE  E018.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E018.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E018.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E019.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E019.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E019.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E019.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E020.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E020.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E020.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E020.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E036.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tNew_Hinban.opecond & "'  " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    sql = sql & "   AND E048.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E048.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E048.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E048.OPECOND   =   '" & tNew_Hinban.opecond & "' " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
'*** UPDATE ↑ Y.SIMIZU 2005/10/12 GDﾗｲﾝ数取得追加
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    With tbl_chk1_4(1)
        'Rs
        If IsNull(rs("HSXRHWYS")) = False Then .HSXRHWYS = rs("HSXRHWYS") Else .HSXRHWYS = " "              '保証方法_対象
        'Oi
        If IsNull(rs("HSXONHWS")) = False Then .HSXONHWS = rs("HSXONHWS") Else .HSXONHWS = " "              '保証方法_対象
        If IsNull(rs("HSXONSPT")) = False Then .HSXONSPT = rs("HSXONSPT") Else .HSXONSPT = " "              '測定位置_点    '08/01/29 ooba
        If IsNull(rs("HSXONSPI")) = False Then .HSXONSPI = rs("HSXONSPI") Else .HSXONSPI = " "              '測定位置_位
        If IsNull(rs("HSXONKWY")) = False Then .HSXONKWY = rs("HSXONKWY") Else .HSXONKWY = " "              '検査方法
        'OSF1
        If IsNull(rs("HSXOF1HS")) = False Then .HSXOF1HS = rs("HSXOF1HS") Else .HSXOF1HS = " "              '保証方法_対象
        If IsNull(rs("HSXOF1SH")) = False Then .HSXOF1SH = rs("HSXOF1SH") Else .HSXOF1SH = " "              '測定位置_方
        If IsNull(rs("HSXOF1ST")) = False Then .HSXOF1ST = rs("HSXOF1ST") Else .HSXOF1ST = " "              '測定位置_点
        If IsNull(rs("HSXOF1SR")) = False Then .HSXOF1SR = rs("HSXOF1SR") Else .HSXOF1SR = " "              '測定位置_領
        If IsNull(rs("HSXOF1NS")) = False Then .HSXOF1NS = rs("HSXOF1NS") Else .HSXOF1NS = " "              '熱処理法
        If IsNull(rs("HSXOF1SZ")) = False Then .HSXOF1SZ = rs("HSXOF1SZ") Else .HSXOF1SZ = " "              '測定条件
        If IsNull(rs("HSXOF1ET")) = False Then .HSXOF1ET = rs("HSXOF1ET") Else .HSXOF1ET = 0                '選択ET代
        If IsNull(rs("HSXOSF1PTK")) = False Then .HSXOSF1PTK = rs("HSXOSF1PTK") Else .HSXOSF1PTK = "4"      'パターン区分
        'OSF2
        If IsNull(rs("HSXOF2HS")) = False Then .HSXOF2HS = rs("HSXOF2HS") Else .HSXOF2HS = " "              '保証方法_対象
        If IsNull(rs("HSXOF2SH")) = False Then .HSXOF2SH = rs("HSXOF2SH") Else .HSXOF2SH = " "              '測定位置_方
        If IsNull(rs("HSXOF2ST")) = False Then .HSXOF2ST = rs("HSXOF2ST") Else .HSXOF2ST = " "              '測定位置_点
        If IsNull(rs("HSXOF2SR")) = False Then .HSXOF2SR = rs("HSXOF2SR") Else .HSXOF2SR = " "              '測定位置_領
        If IsNull(rs("HSXOF2NS")) = False Then .HSXOF2NS = rs("HSXOF2NS") Else .HSXOF2NS = " "              '熱処理法
        If IsNull(rs("HSXOF2SZ")) = False Then .HSXOF2SZ = rs("HSXOF2SZ") Else .HSXOF2SZ = " "              '測定条件
        If IsNull(rs("HSXOF2ET")) = False Then .HSXOF2ET = rs("HSXOF2ET") Else .HSXOF2ET = 0                '選択ET代
        If IsNull(rs("HSXOSF2PTK")) = False Then .HSXOSF2PTK = rs("HSXOSF2PTK") Else .HSXOSF2PTK = "4"      'パターン区分
        'OSF3
        If IsNull(rs("HSXOF3HS")) = False Then .HSXOF3HS = rs("HSXOF3HS") Else .HSXOF3HS = " "              '保証方法_対象
        If IsNull(rs("HSXOF3SH")) = False Then .HSXOF3SH = rs("HSXOF3SH") Else .HSXOF3SH = " "              '測定位置_方
        If IsNull(rs("HSXOF3ST")) = False Then .HSXOF3ST = rs("HSXOF3ST") Else .HSXOF3ST = " "              '測定位置_点
        If IsNull(rs("HSXOF3SR")) = False Then .HSXOF3SR = rs("HSXOF3SR") Else .HSXOF3SR = " "              '測定位置_領
        If IsNull(rs("HSXOF3NS")) = False Then .HSXOF3NS = rs("HSXOF3NS") Else .HSXOF3NS = " "              '熱処理法
        If IsNull(rs("HSXOF3SZ")) = False Then .HSXOF3SZ = rs("HSXOF3SZ") Else .HSXOF3SZ = " "              '測定条件
        If IsNull(rs("HSXOF3ET")) = False Then .HSXOF3ET = rs("HSXOF3ET") Else .HSXOF3ET = 0                '選択ET代
        If IsNull(rs("HSXOSF3PTK")) = False Then .HSXOSF3PTK = rs("HSXOSF3PTK") Else .HSXOSF3PTK = "4"      'パターン区分


''C－OSF3チェックの変更 2008.04.20 青柳
'''C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
''        'If IsNull(rs("HSXOF4HS")) = False Then .HSXOF4HS = rs("HSXOF4HS") Else .HSXOF4HS = " "             '保証方法_対象
''        If IsNull(rs("COSF3FLAG")) = False Then .HSXOF4HS = rs("COSF3FLAG") Else .HSXOF4HS = " "            'C-OSF3ﾌﾗｸﾞ
'''C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---
        
'--------------- 2008/08/25 INSERT START  By Systech ---------------
        .HSXDKTMP = rs("HSXDKTMP")
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
        
        'OSF4
'        If IsNull(rs("HSXOF4SH")) = False Then .HSXOF4SH = rs("HSXOF4SH") Else .HSXOF4SH = " "              '測定位置_方
'        If IsNull(rs("HSXOF4ST")) = False Then .HSXOF4ST = rs("HSXOF4ST") Else .HSXOF4ST = " "              '測定位置_点
'        If IsNull(rs("HSXOF4SR")) = False Then .HSXOF4SR = rs("HSXOF4SR") Else .HSXOF4SR = " "              '測定位置_領
'        If IsNull(rs("HSXOF4NS")) = False Then .HSXOF4NS = rs("HSXOF4NS") Else .HSXOF4NS = " "              '熱処理法
'        If IsNull(rs("HSXOF4SZ")) = False Then .HSXOF4SZ = rs("HSXOF4SZ") Else .HSXOF4SZ = " "              '測定条件
'        If IsNull(rs("HSXOF4ET")) = False Then .HSXOF4ET = rs("HSXOF4ET") Else .HSXOF4ET = 0                '選択ET代
'        If IsNull(rs("HSXOSF4PTK")) = False Then .HSXOSF4PTK = rs("HSXOSF4PTK") Else .HSXOSF4PTK = "4"      'パターン区分
        If IsNull(rs("HSXOF1HS")) = False Then .HSXOF4HS = rs("HSXOF1HS") Else .HSXOF4HS = " "              '保証方法_対象(ArANはOSF1保証) 08/12/21 ooba
        If IsNull(rs("HSXOF1ARPTK")) = False Then .HSXOSF4PTK = rs("HSXOF1ARPTK") Else .HSXOSF4PTK = " "    '(ArAN)ﾊﾟﾀﾝ区分 08/12/21 ooba
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
        'SIRD
        If IsNull(rs("HWFSIRDMX")) = False Then .HWFSIRDMX = rs("HWFSIRDMX") Else .HWFSIRDMX = "0"          '軸状転位上限
        If IsNull(rs("HWFSIRDSZ")) = False Then .HWFSIRDSZ = rs("HWFSIRDSZ") Else .HWFSIRDSZ = " "          '軸状転位測定条件
        If IsNull(rs("HWFSIRDHT")) = False Then .HWFSIRDHT = rs("HWFSIRDHT") Else .HWFSIRDHT = " "          '軸状転位保証方法＿対
        If IsNull(rs("HWFSIRDHS")) = False Then .HWFSIRDHS = rs("HWFSIRDHS") Else .HWFSIRDHS = " "          '軸状転位保証方法＿処
        If IsNull(rs("HWFSIRDKM")) = False Then .HWFSIRDKM = rs("HWFSIRDKM") Else .HWFSIRDKM = " "          '軸状転位検査頻度＿枚
        If IsNull(rs("HWFSIRDKN")) = False Then .HWFSIRDKN = rs("HWFSIRDKN") Else .HWFSIRDKN = " "          '軸状転位検査頻度＿抜
        If IsNull(rs("HWFSIRDKH")) = False Then .HWFSIRDKH = rs("HWFSIRDKH") Else .HWFSIRDKH = " "          '軸状転位検査頻度＿保
        If IsNull(rs("HWFSIRDKU")) = False Then .HWFSIRDKU = rs("HWFSIRDKU") Else .HWFSIRDKU = " "          '軸状転位検査頻度＿ウ
        If IsNull(rs("HWFSIRDPS")) = False Then .HWFSIRDPS = Trim(rs("HWFSIRDPS")) Else .HWFSIRDPS = " "    '軸状転位TB保証位置
        
        '「軸状転位TB保証位置」を判定し、「軸状転位検査頻度＿抜」に編集
        Select Case Trim(.HWFSIRDPS)
        Case "T"
            .HWFSIRDKN = "3"
        Case "B"
            .HWFSIRDKN = "4"
        Case "TB"
            .HWFSIRDKN = "6"
        Case Else
            .HWFSIRDKN = " "
        End Select
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
        'BMD1
        If IsNull(rs("HSXBM1HS")) = False Then .HSXBM1HS = rs("HSXBM1HS") Else .HSXBM1HS = " "              '保証方法_対象
        If IsNull(rs("HSXBM1SH")) = False Then .HSXBM1SH = rs("HSXBM1SH") Else .HSXBM1SH = " "              '測定位置_方
        If IsNull(rs("HSXBM1ST")) = False Then .HSXBM1ST = rs("HSXBM1ST") Else .HSXBM1ST = " "              '測定位置_点
        If IsNull(rs("HSXBM1SR")) = False Then .HSXBM1SR = rs("HSXBM1SR") Else .HSXBM1SR = " "              '測定位置_領
        If IsNull(rs("HSXBM1NS")) = False Then .HSXBM1NS = rs("HSXBM1NS") Else .HSXBM1NS = " "              '熱処理法
        If IsNull(rs("HSXBM1SZ")) = False Then .HSXBM1SZ = rs("HSXBM1SZ") Else .HSXBM1SZ = " "              '測定条件
        If IsNull(rs("HSXBM1ET")) = False Then .HSXBM1ET = rs("HSXBM1ET") Else .HSXBM1ET = 0                '選択ET代
        'BMD2
        If IsNull(rs("HSXBM2HS")) = False Then .HSXBM2HS = rs("HSXBM2HS") Else .HSXBM2HS = " "              '保証方法_対象
        If IsNull(rs("HSXBM2SH")) = False Then .HSXBM2SH = rs("HSXBM2SH") Else .HSXBM2SH = " "              '測定位置_方
        If IsNull(rs("HSXBM2ST")) = False Then .HSXBM2ST = rs("HSXBM2ST") Else .HSXBM2ST = " "              '測定位置_点
        If IsNull(rs("HSXBM2SR")) = False Then .HSXBM2SR = rs("HSXBM2SR") Else .HSXBM2SR = " "              '測定位置_領
        If IsNull(rs("HSXBM2NS")) = False Then .HSXBM2NS = rs("HSXBM2NS") Else .HSXBM2NS = " "              '熱処理法
        If IsNull(rs("HSXBM2SZ")) = False Then .HSXBM2SZ = rs("HSXBM2SZ") Else .HSXBM2SZ = " "              '測定条件
        If IsNull(rs("HSXBM2ET")) = False Then .HSXBM2ET = rs("HSXBM2ET") Else .HSXBM2ET = 0                '選択ET代
        'BMD3
        If IsNull(rs("HSXBM3HS")) = False Then .HSXBM3HS = rs("HSXBM3HS") Else .HSXBM3HS = " "              '保証方法_対象
        If IsNull(rs("HSXBM3SH")) = False Then .HSXBM3SH = rs("HSXBM3SH") Else .HSXBM3SH = " "              '測定位置_方
        If IsNull(rs("HSXBM3ST")) = False Then .HSXBM3ST = rs("HSXBM3ST") Else .HSXBM3ST = " "              '測定位置_点
        If IsNull(rs("HSXBM3SR")) = False Then .HSXBM3SR = rs("HSXBM3SR") Else .HSXBM3SR = " "              '測定位置_領
        If IsNull(rs("HSXBM3NS")) = False Then .HSXBM3NS = rs("HSXBM3NS") Else .HSXBM3NS = " "              '熱処理法
        If IsNull(rs("HSXBM3SZ")) = False Then .HSXBM3SZ = rs("HSXBM3SZ") Else .HSXBM3SZ = " "              '測定条件
        If IsNull(rs("HSXBM3ET")) = False Then .HSXBM3ET = rs("HSXBM3ET") Else .HSXBM3ET = 0                '選択ET代
        'EPD
        If IsNull(rs("HSXTMMAX")) = False Then .HSXTMMAX = rs("HSXTMMAX") Else .HSXTMMAX = 0                '上限
        'LT
        If IsNull(rs("HSXLTHWS")) = False Then .HSXLTHWS = rs("HSXLTHWS") Else .HSXLTHWS = " "              '保証方法_対象
        'CS
        If IsNull(rs("HSXCNHWS")) = False Then .HSXCNHWS = rs("HSXCNHWS") Else .HSXCNHWS = " "              '保証方法_対象
        If IsNull(rs("HSXCNKWY")) = False Then .HSXCNKWY = rs("HSXCNKWY") Else .HSXCNKWY = " "              '検査方法
        If IsNull(rs("HSXCNKHI")) = False Then .HSXCNKHI = rs("HSXCNKHI") Else .HSXCNKHI = " "              '検査頻度＿位   '' add 0108
        'DEN
        If IsNull(rs("HSXDENHS")) = False Then .HSXDENHS = rs("HSXDENHS") Else .HSXDENHS = " "              '保証方法_対象
        If IsNull(rs("HSXDENMN")) = False Then .HSXDENMN = rs("HSXDENMN") Else .HSXDENMN = 0                '下限
        If IsNull(rs("HSXDENMX")) = False Then .HSXDENMX = rs("HSXDENMX") Else .HSXDENMX = 0                '上限
        'DVD2
        If IsNull(rs("HSXDVDHS")) = False Then .HSXDVDHS = rs("HSXDVDHS") Else .HSXDVDHS = " "              '保証方法_対象
        If IsNull(rs("HSXDVDMNN")) = False Then .HSXDVDMNN = rs("HSXDVDMNN") Else .HSXDVDMNN = 0            '下限
        If IsNull(rs("HSXDVDMXN")) = False Then .HSXDVDMXN = rs("HSXDVDMXN") Else .HSXDVDMXN = 0            '上限
        'L/DL
        If IsNull(rs("HSXLDLHS")) = False Then .HSXLDLHS = rs("HSXLDLHS") Else .HSXLDLHS = " "              '保証方法_対象
        If IsNull(rs("HSXLDLMN")) = False Then .HSXLDLMN = rs("HSXLDLMN") Else .HSXLDLMN = 0                '下限
        If IsNull(rs("HSXLDLMX")) = False Then .HSXLDLMX = rs("HSXLDLMX") Else .HSXLDLMX = 0                '上限
    '*** UPDATE ↓ Y.SIMIZU 2005/10/12 GDﾗｲﾝ数追加
        If IsNull(rs("HSXGDLINE")) = False Then .HSXGDLINE = rs("HSXGDLINE") Else .HSXGDLINE = " "          'GDﾗｲﾝ数
    '*** UPDATE ↑ Y.SIMIZU 2005/10/12 GDﾗｲﾝ数追加
    End With
    
    Set rs = Nothing
    '------------------------------------------ 指示取得 ------------------------------------------------------
    On Error GoTo Apl_down
    If iELCs_Flg = 0 Or iELCs_Flg = 2 Then ''<<複数品番判定対応　20060509SMP桜井
        '比抵抗
        sErr_Msg = "1-4 比抵抗ﾁｪｯｸ"
        sResult = ""
        RET = funCodeDBGet("SB", "14", "RS", 0, " ", sResult)
        If RET <> 0 Then
            sErr_Msg = sErr_Msg & "→指示取得"
            GoTo CodeDBGet_Error
        End If
        Erase tbl_chk1_4_1
        tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HSXRHWYS
        tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HSXRHWYS
'--------------- 2008/08/25 INSERT START  By Systech ---------------
        tbl_chk1_4_1(0).HSXDKTMP = tbl_chk1_4(0).HSXDKTMP
        tbl_chk1_4_1(1).HSXDKTMP = tbl_chk1_4(1).HSXDKTMP
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
        RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK1-4,RS")
        If RET <> 0 Then
            funChkFurikae1_4 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
            gsTbcmy028ErrCode = "00030"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            GoTo Apl_Exit
        End If
        '酸素濃度
        sErr_Msg = "1-4 酸素濃度ﾁｪｯｸ"
        sResult = ""
        RET = funCodeDBGet("SB", "14", "OI", 0, " ", sResult)
        If RET <> 0 Then
            sErr_Msg = sErr_Msg & "→指示取得"
            GoTo CodeDBGet_Error
        End If
        Erase tbl_chk1_4_1
        tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HSXONHWS
        tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HSXONHWS
        tbl_chk1_4_1(0).SOKU_TEN = tbl_chk1_4(0).HSXONSPT   '08/01/29 ooba
        tbl_chk1_4_1(1).SOKU_TEN = tbl_chk1_4(1).HSXONSPT   '08/01/29 ooba
        tbl_chk1_4_1(0).SOKU_ICHI = tbl_chk1_4(0).HSXONSPI
        tbl_chk1_4_1(1).SOKU_ICHI = tbl_chk1_4(1).HSXONSPI
        tbl_chk1_4_1(0).KENSA = tbl_chk1_4(0).HSXONKWY
        tbl_chk1_4_1(1).KENSA = tbl_chk1_4(1).HSXONKWY
        RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK1-4,Oi")
        If RET <> 0 Then
            funChkFurikae1_4 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
            gsTbcmy028ErrCode = "00031"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            GoTo Apl_Exit
        End If
        'ＯＳＦ1
        sErr_Msg = "1-4 OSF1ﾁｪｯｸ"
        sResult = ""
        RET = funCodeDBGet("SB", "14", "O1", 0, " ", sResult)
        If RET <> 0 Then
            sErr_Msg = sErr_Msg & "→指示取得"
            GoTo CodeDBGet_Error
        End If
        Erase tbl_chk1_4_1
        tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HSXOF1HS
        tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HSXOF1HS
        tbl_chk1_4_1(0).SOKU_HOU = tbl_chk1_4(0).HSXOF1SH
        tbl_chk1_4_1(1).SOKU_HOU = tbl_chk1_4(1).HSXOF1SH
        tbl_chk1_4_1(0).SOKU_TEN = tbl_chk1_4(0).HSXOF1ST
        tbl_chk1_4_1(1).SOKU_TEN = tbl_chk1_4(1).HSXOF1ST
        tbl_chk1_4_1(0).SOKU_RYOU = tbl_chk1_4(0).HSXOF1SR
        tbl_chk1_4_1(1).SOKU_RYOU = tbl_chk1_4(1).HSXOF1SR
        tbl_chk1_4_1(0).NETSU = tbl_chk1_4(0).HSXOF1NS
        tbl_chk1_4_1(1).NETSU = tbl_chk1_4(1).HSXOF1NS
        tbl_chk1_4_1(0).JOUKEN = tbl_chk1_4(0).HSXOF1SZ
        tbl_chk1_4_1(1).JOUKEN = tbl_chk1_4(1).HSXOF1SZ
        tbl_chk1_4_1(0).ET = tbl_chk1_4(0).HSXOF1ET
        tbl_chk1_4_1(1).ET = tbl_chk1_4(1).HSXOF1ET
        tbl_chk1_4_1(0).PATTERN = tbl_chk1_4(0).HSXOSF1PTK
        tbl_chk1_4_1(1).PATTERN = tbl_chk1_4(1).HSXOSF1PTK
        RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK1-4,OSF1")
        If RET <> 0 Then
            funChkFurikae1_4 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
            gsTbcmy028ErrCode = "00033"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            GoTo Apl_Exit
        End If
        'ＯＳＦ２
        sErr_Msg = "1-4 OSF2ﾁｪｯｸ"
        sResult = ""
        RET = funCodeDBGet("SB", "14", "O2", 0, " ", sResult)
        If RET <> 0 Then
            sErr_Msg = sErr_Msg & "→指示取得"
            GoTo CodeDBGet_Error
        End If
        Erase tbl_chk1_4_1
        tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HSXOF2HS
        tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HSXOF2HS
        tbl_chk1_4_1(0).SOKU_HOU = tbl_chk1_4(0).HSXOF2SH
        tbl_chk1_4_1(1).SOKU_HOU = tbl_chk1_4(1).HSXOF2SH
        tbl_chk1_4_1(0).SOKU_TEN = tbl_chk1_4(0).HSXOF2ST
        tbl_chk1_4_1(1).SOKU_TEN = tbl_chk1_4(1).HSXOF2ST
        tbl_chk1_4_1(0).SOKU_RYOU = tbl_chk1_4(0).HSXOF2SR
        tbl_chk1_4_1(1).SOKU_RYOU = tbl_chk1_4(1).HSXOF2SR
        tbl_chk1_4_1(0).NETSU = tbl_chk1_4(0).HSXOF2NS
        tbl_chk1_4_1(1).NETSU = tbl_chk1_4(1).HSXOF2NS
        tbl_chk1_4_1(0).JOUKEN = tbl_chk1_4(0).HSXOF2SZ
        tbl_chk1_4_1(1).JOUKEN = tbl_chk1_4(1).HSXOF2SZ
        tbl_chk1_4_1(0).ET = tbl_chk1_4(0).HSXOF2ET
        tbl_chk1_4_1(1).ET = tbl_chk1_4(1).HSXOF2ET
        tbl_chk1_4_1(0).PATTERN = tbl_chk1_4(0).HSXOSF2PTK
        tbl_chk1_4_1(1).PATTERN = tbl_chk1_4(1).HSXOSF2PTK
        RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK1-4,OSF2")
        If RET <> 0 Then
            funChkFurikae1_4 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
            gsTbcmy028ErrCode = "00034"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            GoTo Apl_Exit
        End If
        'ＯＳＦ３
        sErr_Msg = "1-4 OSF3ﾁｪｯｸ"
        sResult = ""
        RET = funCodeDBGet("SB", "14", "O3", 0, " ", sResult)
        If RET <> 0 Then
            sErr_Msg = sErr_Msg & "→指示取得"
            GoTo CodeDBGet_Error
        End If
        Erase tbl_chk1_4_1
        tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HSXOF3HS
        tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HSXOF3HS
        tbl_chk1_4_1(0).SOKU_HOU = tbl_chk1_4(0).HSXOF3SH
        tbl_chk1_4_1(1).SOKU_HOU = tbl_chk1_4(1).HSXOF3SH
        tbl_chk1_4_1(0).SOKU_TEN = tbl_chk1_4(0).HSXOF3ST
        tbl_chk1_4_1(1).SOKU_TEN = tbl_chk1_4(1).HSXOF3ST
        tbl_chk1_4_1(0).SOKU_RYOU = tbl_chk1_4(0).HSXOF3SR
        tbl_chk1_4_1(1).SOKU_RYOU = tbl_chk1_4(1).HSXOF3SR
        tbl_chk1_4_1(0).NETSU = tbl_chk1_4(0).HSXOF3NS
        tbl_chk1_4_1(1).NETSU = tbl_chk1_4(1).HSXOF3NS
        tbl_chk1_4_1(0).JOUKEN = tbl_chk1_4(0).HSXOF3SZ
        tbl_chk1_4_1(1).JOUKEN = tbl_chk1_4(1).HSXOF3SZ
        tbl_chk1_4_1(0).ET = tbl_chk1_4(0).HSXOF3ET
        tbl_chk1_4_1(1).ET = tbl_chk1_4(1).HSXOF3ET
        tbl_chk1_4_1(0).PATTERN = tbl_chk1_4(0).HSXOSF3PTK
        tbl_chk1_4_1(1).PATTERN = tbl_chk1_4(1).HSXOSF3PTK
        RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK1-4,OSF3")
        If RET <> 0 Then
            funChkFurikae1_4 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
            gsTbcmy028ErrCode = "00035"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            GoTo Apl_Exit
        End If
        'ＯＳＦ４
        sErr_Msg = "1-4 OSF4ﾁｪｯｸ"
        sResult = ""
        RET = funCodeDBGet("SB", "14", "O4", 0, " ", sResult)
        If RET <> 0 Then
            sErr_Msg = sErr_Msg & "→指示取得"
            GoTo CodeDBGet_Error
        End If
        tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HSXOF4HS
        tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HSXOF4HS
        tbl_chk1_4_1(0).SOKU_HOU = tbl_chk1_4(0).HSXOF4SH
        tbl_chk1_4_1(1).SOKU_HOU = tbl_chk1_4(1).HSXOF4SH
        tbl_chk1_4_1(0).SOKU_TEN = tbl_chk1_4(0).HSXOF4ST
        tbl_chk1_4_1(1).SOKU_TEN = tbl_chk1_4(1).HSXOF4ST
        tbl_chk1_4_1(0).SOKU_RYOU = tbl_chk1_4(0).HSXOF4SR
        tbl_chk1_4_1(1).SOKU_RYOU = tbl_chk1_4(1).HSXOF4SR
        tbl_chk1_4_1(0).NETSU = tbl_chk1_4(0).HSXOF4NS
        tbl_chk1_4_1(1).NETSU = tbl_chk1_4(1).HSXOF4NS
        tbl_chk1_4_1(0).JOUKEN = tbl_chk1_4(0).HSXOF4SZ
        tbl_chk1_4_1(1).JOUKEN = tbl_chk1_4(1).HSXOF4SZ
        tbl_chk1_4_1(0).ET = tbl_chk1_4(0).HSXOF4ET
        tbl_chk1_4_1(1).ET = tbl_chk1_4(1).HSXOF4ET
        tbl_chk1_4_1(0).PATTERN = tbl_chk1_4(0).HSXOSF4PTK
        tbl_chk1_4_1(1).PATTERN = tbl_chk1_4(1).HSXOSF4PTK
'        RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK1-4,C-OSF3")
        RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK1-4,OSF-ArAN")     '08/12/21 ooba
        If RET <> 0 Then
            funChkFurikae1_4 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
            gsTbcmy028ErrCode = "00036"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            GoTo Apl_Exit
        End If
        
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
        'ＳＩＲＤ
        sErr_Msg = "1-4 SIRDﾁｪｯｸ"
        sResult = ""
        RET = funCodeDBGet("SB", "14", "SD", 0, " ", sResult)
        If RET <> 0 Then
            sErr_Msg = sErr_Msg & "→指示取得"
            GoTo CodeDBGet_Error
        End If
        Erase tbl_chk1_4_1                                          '引数ﾃｰﾌﾞﾙｸﾘｱ
        tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HWFSIRDHS            '軸状転位保証方法＿処
        tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HWFSIRDHS            '軸状転位保証方法＿処
        tbl_chk1_4_1(0).JOUKEN = tbl_chk1_4(0).HWFSIRDSZ            '軸状転位測定条件
        tbl_chk1_4_1(1).JOUKEN = tbl_chk1_4(1).HWFSIRDSZ            '軸状転位測定条件
        tbl_chk1_4_1(0).HWFSIRDMX = tbl_chk1_4(0).HWFSIRDMX         '軸状転位上限
        tbl_chk1_4_1(1).HWFSIRDMX = tbl_chk1_4(1).HWFSIRDMX         '軸状転位上限
        tbl_chk1_4_1(0).HWFSIRDHT = tbl_chk1_4(0).HWFSIRDHT         '軸状転位保証方法＿対
        tbl_chk1_4_1(1).HWFSIRDHT = tbl_chk1_4(1).HWFSIRDHT         '軸状転位保証方法＿対
        tbl_chk1_4_1(0).HWFSIRDKM = tbl_chk1_4(0).HWFSIRDKM         '軸状転位検査頻度＿枚
        tbl_chk1_4_1(1).HWFSIRDKM = tbl_chk1_4(1).HWFSIRDKM         '軸状転位検査頻度＿枚
        tbl_chk1_4_1(0).HWFSIRDKH = tbl_chk1_4(0).HWFSIRDKH         '軸状転位検査頻度＿保
        tbl_chk1_4_1(1).HWFSIRDKH = tbl_chk1_4(1).HWFSIRDKH         '軸状転位検査頻度＿保
        tbl_chk1_4_1(0).HWFSIRDKU = tbl_chk1_4(0).HWFSIRDKU         '軸状転位検査頻度＿ウ
        tbl_chk1_4_1(1).HWFSIRDKU = tbl_chk1_4(1).HWFSIRDKU         '軸状転位検査頻度＿ウ
        tbl_chk1_4_1(0).HWFSIRDKN = tbl_chk1_4(0).HWFSIRDKN         '軸状転位検査頻度＿抜
        tbl_chk1_4_1(1).HWFSIRDKN = tbl_chk1_4(1).HWFSIRDKN         '軸状転位検査頻度＿抜
        RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK1-4,SIRD")
        If RET <> 0 Then
            funChkFurikae1_4 = RET
            gsTbcmy028ErrCode = "00036"
            GoTo Apl_Exit
        End If
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
        
        'ＢＭＤ１
        sErr_Msg = "1-4 BMD1ﾁｪｯｸ"
        sResult = ""
        RET = funCodeDBGet("SB", "14", "B1", 0, " ", sResult)
        If RET <> 0 Then
            sErr_Msg = sErr_Msg & "→指示取得"
            GoTo CodeDBGet_Error
        End If
        Erase tbl_chk1_4_1
        tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HSXBM1HS
        tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HSXBM1HS
        tbl_chk1_4_1(0).SOKU_HOU = tbl_chk1_4(0).HSXBM1SH
        tbl_chk1_4_1(1).SOKU_HOU = tbl_chk1_4(1).HSXBM1SH
        tbl_chk1_4_1(0).SOKU_TEN = tbl_chk1_4(0).HSXBM1ST
        tbl_chk1_4_1(1).SOKU_TEN = tbl_chk1_4(1).HSXBM1ST
        tbl_chk1_4_1(0).SOKU_RYOU = tbl_chk1_4(0).HSXBM1SR
        tbl_chk1_4_1(1).SOKU_RYOU = tbl_chk1_4(1).HSXBM1SR
        tbl_chk1_4_1(0).NETSU = tbl_chk1_4(0).HSXBM1NS
        tbl_chk1_4_1(1).NETSU = tbl_chk1_4(1).HSXBM1NS
        tbl_chk1_4_1(0).JOUKEN = tbl_chk1_4(0).HSXBM1SZ
        tbl_chk1_4_1(1).JOUKEN = tbl_chk1_4(1).HSXBM1SZ
        tbl_chk1_4_1(0).ET = tbl_chk1_4(0).HSXBM1ET
        tbl_chk1_4_1(1).ET = tbl_chk1_4(1).HSXBM1ET
        RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK1-4,BMD1")
        If RET <> 0 Then
            funChkFurikae1_4 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
            gsTbcmy028ErrCode = "00037"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            GoTo Apl_Exit
        End If
        'ＢＭＤ２
        sErr_Msg = "1-4 BMD2ﾁｪｯｸ"
        sResult = ""
        RET = funCodeDBGet("SB", "14", "B2", 0, " ", sResult)
        If RET <> 0 Then
            sErr_Msg = sErr_Msg & "→指示取得"
            GoTo CodeDBGet_Error
        End If
        Erase tbl_chk1_4_1
        tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HSXBM2HS
        tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HSXBM2HS
        tbl_chk1_4_1(0).SOKU_HOU = tbl_chk1_4(0).HSXBM2SH
        tbl_chk1_4_1(1).SOKU_HOU = tbl_chk1_4(1).HSXBM2SH
        tbl_chk1_4_1(0).SOKU_TEN = tbl_chk1_4(0).HSXBM2ST
        tbl_chk1_4_1(1).SOKU_TEN = tbl_chk1_4(1).HSXBM2ST
        tbl_chk1_4_1(0).SOKU_RYOU = tbl_chk1_4(0).HSXBM2SR
        tbl_chk1_4_1(1).SOKU_RYOU = tbl_chk1_4(1).HSXBM2SR
        tbl_chk1_4_1(0).NETSU = tbl_chk1_4(0).HSXBM2NS
        tbl_chk1_4_1(1).NETSU = tbl_chk1_4(1).HSXBM2NS
        tbl_chk1_4_1(0).JOUKEN = tbl_chk1_4(0).HSXBM2SZ
        tbl_chk1_4_1(1).JOUKEN = tbl_chk1_4(1).HSXBM2SZ
        tbl_chk1_4_1(0).ET = tbl_chk1_4(0).HSXBM2ET
        tbl_chk1_4_1(1).ET = tbl_chk1_4(1).HSXBM2ET
        RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK1-4,BMD2")
        If RET <> 0 Then
            funChkFurikae1_4 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
            gsTbcmy028ErrCode = "00038"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            GoTo Apl_Exit
        End If
        'ＢＭＤ３
        sErr_Msg = "1-4 BMD3ﾁｪｯｸ"
        sResult = ""
        RET = funCodeDBGet("SB", "14", "B3", 0, " ", sResult)
        If RET <> 0 Then
            sErr_Msg = sErr_Msg & "→指示取得"
            GoTo CodeDBGet_Error
        End If
        Erase tbl_chk1_4_1
        tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HSXBM3HS
        tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HSXBM3HS
        tbl_chk1_4_1(0).SOKU_HOU = tbl_chk1_4(0).HSXBM3SH
        tbl_chk1_4_1(1).SOKU_HOU = tbl_chk1_4(1).HSXBM3SH
        tbl_chk1_4_1(0).SOKU_TEN = tbl_chk1_4(0).HSXBM3ST
        tbl_chk1_4_1(1).SOKU_TEN = tbl_chk1_4(1).HSXBM3ST
        tbl_chk1_4_1(0).SOKU_RYOU = tbl_chk1_4(0).HSXBM3SR
        tbl_chk1_4_1(1).SOKU_RYOU = tbl_chk1_4(1).HSXBM3SR
        tbl_chk1_4_1(0).NETSU = tbl_chk1_4(0).HSXBM3NS
        tbl_chk1_4_1(1).NETSU = tbl_chk1_4(1).HSXBM3NS
        tbl_chk1_4_1(0).JOUKEN = tbl_chk1_4(0).HSXBM3SZ
        tbl_chk1_4_1(1).JOUKEN = tbl_chk1_4(1).HSXBM3SZ
        tbl_chk1_4_1(0).ET = tbl_chk1_4(0).HSXBM3ET
        tbl_chk1_4_1(1).ET = tbl_chk1_4(1).HSXBM3ET
        RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK1-4,BMD3")
        If RET <> 0 Then
            funChkFurikae1_4 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
            gsTbcmy028ErrCode = "00039"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            GoTo Apl_Exit
        End If
    End If ''<<<複数品番判定対応
    
    
    Select Case iELCs_Flg   ''<<複数品番判定対応　SMP近藤 06/07/04
    Case 0, 1, 4            ''<<複数品番判定対応　SMP近藤 06/07/04
    
    'ＥＰＤ
    sErr_Msg = "1-4 EPDﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "14", "EPD", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_4_1
    tbl_chk1_4_1(0).max = tbl_chk1_4(0).HSXTMMAX
    tbl_chk1_4_1(1).max = tbl_chk1_4(1).HSXTMMAX
    RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK1-4,EPD")
    If RET <> 0 Then
        funChkFurikae1_4 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00032"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    
    Case 0, 1, 5            ''<<複数品番判定対応　SMP近藤 06/07/04
    
    'ライフタイム
    sErr_Msg = "1-4 ﾗｲﾌﾀｲﾑﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "14", "LT", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_4_1
    tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HSXLTHWS
    tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HSXLTHWS
    RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK1-4,LT")
    If RET <> 0 Then
        funChkFurikae1_4 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00040"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    
    Case 0, 1, 3            ''<<複数品番判定対応　SMP近藤 06/07/04
    
    '炭素濃度
    sErr_Msg = "1-4 炭素濃度ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "14", "CS", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_4_1
    tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HSXCNHWS
    tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HSXCNHWS
    tbl_chk1_4_1(0).KENSA = tbl_chk1_4(0).HSXCNKWY
    tbl_chk1_4_1(1).KENSA = tbl_chk1_4(1).HSXCNKWY
    'add start 0108
    tbl_chk1_4_1(0).HSXCNKHI = tbl_chk1_4(0).HSXCNKHI
    tbl_chk1_4_1(1).HSXCNKHI = tbl_chk1_4(1).HSXCNKHI
    'add end 0108
    RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK1-4,CS")
    If RET <> 0 Then
        funChkFurikae1_4 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00041"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If

    End Select              ''<<複数品番判定対応　SMP近藤 06/07/04

    
    If iELCs_Flg = 0 Or iELCs_Flg = 2 Then ''<<複数品番判定対応　SMP桜井
        'ＤＥＮ
        sErr_Msg = "1-4 DENﾁｪｯｸ"
        sResult = ""
        RET = funCodeDBGet("SB", "14", "DEN", 0, " ", sResult)
        If RET <> 0 Then
            sErr_Msg = sErr_Msg & "→指示取得"
            GoTo CodeDBGet_Error
        End If
        Erase tbl_chk1_4_1
        tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HSXDENHS
        tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HSXDENHS
        tbl_chk1_4_1(0).Min = tbl_chk1_4(0).HSXDENMN
        tbl_chk1_4_1(1).Min = tbl_chk1_4(1).HSXDENMN
        tbl_chk1_4_1(0).max = tbl_chk1_4(0).HSXDENMX
        tbl_chk1_4_1(1).max = tbl_chk1_4(1).HSXDENMX
    '*** UPDATE ↓ Y.SIMIZU 2005/10/12 ﾗｲﾝ数追加
        tbl_chk1_4_1(0).LINE = tbl_chk1_4(0).HSXGDLINE
        tbl_chk1_4_1(1).LINE = tbl_chk1_4(1).HSXGDLINE
    '*** UPDATE ↑ Y.SIMIZU 2005/10/12 ﾗｲﾝ数追加
        RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK1-4,DEN")
        If RET <> 0 Then
            funChkFurikae1_4 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
            If iErr_Code = 1413 Then
                gsTbcmy028ErrCode = "00042"
            Else
                gsTbcmy028ErrCode = "00043"
            End If
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            GoTo Apl_Exit
        End If
        'ＤＶＤ２
        sErr_Msg = "1-4 DVD2ﾁｪｯｸ"
        sResult = ""
        RET = funCodeDBGet("SB", "14", "DVD", 0, " ", sResult)
        If RET <> 0 Then
            sErr_Msg = sErr_Msg & "→指示取得"
            GoTo CodeDBGet_Error
        End If
        Erase tbl_chk1_4_1
        tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HSXDVDHS
        tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HSXDVDHS
        tbl_chk1_4_1(0).Min = tbl_chk1_4(0).HSXDVDMNN
        tbl_chk1_4_1(1).Min = tbl_chk1_4(1).HSXDVDMNN
        tbl_chk1_4_1(0).max = tbl_chk1_4(0).HSXDVDMXN
        tbl_chk1_4_1(1).max = tbl_chk1_4(1).HSXDVDMXN
    '*** UPDATE ↓ Y.SIMIZU 2005/10/12 ﾗｲﾝ数追加
        tbl_chk1_4_1(0).LINE = tbl_chk1_4(0).HSXGDLINE
        tbl_chk1_4_1(1).LINE = tbl_chk1_4(1).HSXGDLINE
    '*** UPDATE ↑ Y.SIMIZU 2005/10/12 ﾗｲﾝ数追加
        RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK1-4,DVD")
        If RET <> 0 Then
            funChkFurikae1_4 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
            gsTbcmy028ErrCode = "00044"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            GoTo Apl_Exit
        End If
        'Ｌ／ＤＬ
        sErr_Msg = "1-4 L/DLﾁｪｯｸ"
        sResult = ""
        RET = funCodeDBGet("SB", "14", "LDL", 0, " ", sResult)
        If RET <> 0 Then
            sErr_Msg = sErr_Msg & "→指示取得"
            GoTo CodeDBGet_Error
        End If
        Erase tbl_chk1_4_1
        tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HSXLDLHS
        tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HSXLDLHS
        tbl_chk1_4_1(0).Min = tbl_chk1_4(0).HSXLDLMN
        tbl_chk1_4_1(1).Min = tbl_chk1_4(1).HSXLDLMN
        tbl_chk1_4_1(0).max = tbl_chk1_4(0).HSXLDLMX
        tbl_chk1_4_1(1).max = tbl_chk1_4(1).HSXLDLMX
    '*** UPDATE ↓ Y.SIMIZU 2005/10/12 ﾗｲﾝ数追加
        tbl_chk1_4_1(0).LINE = tbl_chk1_4(0).HSXGDLINE
        tbl_chk1_4_1(1).LINE = tbl_chk1_4(1).HSXGDLINE
    '*** UPDATE ↑ Y.SIMIZU 2005/10/12 ﾗｲﾝ数追加
        RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK1-4,LDL")
        If RET <> 0 Then
            funChkFurikae1_4 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
            gsTbcmy028ErrCode = "00045"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
            GoTo Apl_Exit
        End If
    End If ''<<<複数品番判定対応
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae1_4 = 0 Then
        funChkFurikae1_4 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae1_4 = -4
    GoTo Apl_Exit
    
CodeDBGet_Error:
    Set rs = Nothing
    If funChkFurikae1_4 = 0 Then
        funChkFurikae1_4 = -5
    End If
    GoTo Apl_Exit

End Function

'------------------------------------------------
' 結晶評価項目仕様詳細チェック
'------------------------------------------------

'概要      :指定されたﾁｪｯｸ内容詳細に基づき、該当する仕様値のチェックを行なう。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型             :説明
'          :sChkCode        ,I  ,String         :チェック内容詳細
'          :tbl_chk1_4_1    ,I  ,typ_chk1_4_1   :仕様値構造体配列
'          :iErr_Code       ,O  ,Integer        :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String         :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :sAdd_Msg        ,I  ,String         :添付ｴﾗｰﾒｯｾｰｼﾞ
'          :戻り値          ,O  ,Integer        :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/10 新規作成　SB

Public Function funChkFurikae1_4_1(sChkCode As String, tbl_chk1_4_1() As typ_chk1_4_1, _
                                   iErr_Code As Integer, sErr_Msg As String, sAdd_Msg As String) As Integer

    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim wHOSYOU_0   As String       '保証方法＿対象
    Dim wHOSYOU_1   As String       '保証方法＿対象

    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae1_4_1 = 0
    iErr_Code = 0
    '------------------------------------------ 保証方法チェック ------------------------------------------------------
    If tbl_chk1_4_1(1).HOSYOU <> "H" And tbl_chk1_4_1(1).HOSYOU <> "S" Then GoTo Apl_Exit
    
    '------------------------------------------ 各種チェック ------------------------------------------------------
    '保証方法＿対象
    sErr_Msg = "保証方法_対象ﾁｪｯｸ"
    If Mid(sChkCode, 1, 1) = "2" Then
        '振替元と振替先が等しければ振替ＯＫ
        If tbl_chk1_4_1(0).HOSYOU <> tbl_chk1_4_1(1).HOSYOU Then
            
            wHOSYOU_0 = tbl_chk1_4_1(0).HOSYOU
            If tbl_chk1_4_1(0).HOSYOU <> "H" And tbl_chk1_4_1(0).HOSYOU <> "S" Then wHOSYOU_0 = "-"
            wHOSYOU_1 = tbl_chk1_4_1(1).HOSYOU
            If tbl_chk1_4_1(1).HOSYOU <> "H" And tbl_chk1_4_1(1).HOSYOU <> "S" Then wHOSYOU_1 = "-"
            
            'マトリクス取得
            sResult = ""
'            ret = funCodeDBGet("SB", "SH", tbl_chk1_4_1(0).HOSYOU, 1, tbl_chk1_4_1(1).HOSYOU, sResult)
            RET = funCodeDBGet("SB", "SH", wHOSYOU_0, 1, wHOSYOU_1, sResult)
            If RET <> 0 Then
                sErr_Msg = sAdd_Msg & sErr_Msg & "→元:" & tbl_chk1_4_1(0).HOSYOU & ", 先:" & tbl_chk1_4_1(1).HOSYOU
                GoTo CodeDBGet_Error
            End If
            If sResult = 0 Then
                funChkFurikae1_4_1 = 1
                iErr_Code = 1401
                GoTo Apl_Exit
            End If
        End If
    End If
    '下限
    sErr_Msg = "下限ﾁｪｯｸ"
    If Mid(sChkCode, 2, 1) = "1" Then
        If tbl_chk1_4_1(0).Min <> tbl_chk1_4_1(1).Min Then
            funChkFurikae1_4_1 = 1
            iErr_Code = 1402
            GoTo Apl_Exit
        End If
    End If
    '上限
    sErr_Msg = "上限ﾁｪｯｸ"
    If Mid(sChkCode, 3, 1) = "1" Then
        If tbl_chk1_4_1(0).max <> tbl_chk1_4_1(1).max Then
            funChkFurikae1_4_1 = 1
            iErr_Code = 1403
            GoTo Apl_Exit
        End If
    End If
    '測定位置＿方
    sErr_Msg = "測定位置_方ﾁｪｯｸ"
    If Mid(sChkCode, 4, 1) = "1" Then
        If Trim$(tbl_chk1_4_1(0).SOKU_HOU) <> Trim$(tbl_chk1_4_1(1).SOKU_HOU) Then
            funChkFurikae1_4_1 = 1
            iErr_Code = 1404
            GoTo Apl_Exit
        End If
    End If
    '測定位置＿点
    sErr_Msg = "測定位置_点ﾁｪｯｸ"
    If Mid(sChkCode, 5, 1) = "1" Then
        If Trim$(tbl_chk1_4_1(0).SOKU_TEN) <> Trim$(tbl_chk1_4_1(1).SOKU_TEN) Then
            funChkFurikae1_4_1 = 1
            iErr_Code = 1405
            GoTo Apl_Exit
        End If
    ElseIf Mid(sChkCode, 5, 1) = "2" Then   '08/01/29 ooba
        If Trim$(tbl_chk1_4_1(0).SOKU_TEN) = "" Or _
           Trim$(tbl_chk1_4_1(1).SOKU_TEN) = "" Or _
           Trim$(tbl_chk1_4_1(0).SOKU_TEN) < Trim$(tbl_chk1_4_1(1).SOKU_TEN) Then
            funChkFurikae1_4_1 = 1
            iErr_Code = 1405
            GoTo Apl_Exit
        End If
    End If
    '測定位置＿位
    sErr_Msg = "測定位置_位ﾁｪｯｸ"
    If Mid(sChkCode, 6, 1) = "2" Then
        'マトリクス取得
        sResult = ""
        RET = funCodeDBGet("SB", "OI", tbl_chk1_4_1(0).SOKU_ICHI, 1, tbl_chk1_4_1(1).SOKU_ICHI, sResult)
        If RET <> 0 Then
            sErr_Msg = sAdd_Msg & sErr_Msg & "→元:" & tbl_chk1_4_1(0).SOKU_ICHI & ", 先:" & tbl_chk1_4_1(1).SOKU_ICHI
            GoTo CodeDBGet_Error
        End If
        If sResult = 0 Then
            funChkFurikae1_4_1 = 1
            iErr_Code = 1406
            GoTo Apl_Exit
        End If
    End If
    '測定位置＿領
    sErr_Msg = "測定位置_領ﾁｪｯｸ"
    If Mid(sChkCode, 7, 1) = "1" Then
        If Trim$(tbl_chk1_4_1(0).SOKU_RYOU) <> Trim$(tbl_chk1_4_1(1).SOKU_RYOU) Then
            funChkFurikae1_4_1 = 1
            iErr_Code = 1407
            GoTo Apl_Exit
        End If
    End If
    '検査有無
    sErr_Msg = "検査有無ﾁｪｯｸ"
    If Mid(sChkCode, 8, 1) = "1" Then
        If Trim$(tbl_chk1_4_1(0).UMU) <> Trim$(tbl_chk1_4_1(1).UMU) Then
            funChkFurikae1_4_1 = 1
            iErr_Code = 1408
            GoTo Apl_Exit
        End If
    End If
    '熱処理法
    sErr_Msg = "熱処理法ﾁｪｯｸ"
    If Mid(sChkCode, 9, 1) = "1" Then
        If Trim$(tbl_chk1_4_1(0).NETSU) <> Trim$(tbl_chk1_4_1(1).NETSU) Then
            funChkFurikae1_4_1 = 1
            iErr_Code = 1409
            GoTo Apl_Exit
        End If
    End If
    '測定条件
    sErr_Msg = "測定条件ﾁｪｯｸ"
    If Mid(sChkCode, 10, 1) = "1" Then
        If Trim$(tbl_chk1_4_1(0).JOUKEN) <> Trim$(tbl_chk1_4_1(1).JOUKEN) Then
            funChkFurikae1_4_1 = 1
            iErr_Code = 1410
            GoTo Apl_Exit
        End If
    End If
    '選択ＥＴ代
    sErr_Msg = "選択ET代ﾁｪｯｸ"
    If Mid(sChkCode, 11, 1) = "1" Then
        If tbl_chk1_4_1(0).ET <> tbl_chk1_4_1(1).ET Then
            funChkFurikae1_4_1 = 1
            iErr_Code = 1411
            GoTo Apl_Exit
        End If
    End If
    '検査方法
    sErr_Msg = "検査方法ﾁｪｯｸ"
    If Mid(sChkCode, 12, 1) = "1" Then
        If Trim$(tbl_chk1_4_1(0).KENSA) <> Trim$(tbl_chk1_4_1(1).KENSA) Then
            funChkFurikae1_4_1 = 1
            iErr_Code = 1412
            GoTo Apl_Exit
        End If
    End If
'*** UPDATE ↓ Y.SIMIZU 2005/10/12 ﾗｲﾝ数ﾁｪｯｸ追加
'    'ライン数
'    sErr_Msg = "ﾗｲﾝ数ﾁｪｯｸ"
'    If Mid(sChkCode, 13, 1) = "1" Then
'        If tbl_chk1_4_1(0).LINE <> tbl_chk1_4_1(1).LINE Then
'            funChkFurikae1_4_1 = 1
'            iErr_Code = 1413
'            GoTo Apl_Exit
'        End If
'    End If
    'ライン数
    sErr_Msg = "ﾗｲﾝ数ﾁｪｯｸ"
    If Mid(sChkCode, 13, 1) = "1" Then
        If tbl_chk1_4_1(0).LINE <> tbl_chk1_4_1(1).LINE Then
            funChkFurikae1_4_1 = 1
            iErr_Code = 1413
            GoTo Apl_Exit
        End If
    ElseIf Mid(sChkCode, 13, 1) = "2" Then
        'マトリクス取得
        sResult = ""
           
        RET = funCodeDBGet("SB", "LN", tbl_chk1_4_1(0).LINE, 1, tbl_chk1_4_1(1).LINE, sResult)
        If RET <> 0 Then
            sErr_Msg = sAdd_Msg & sErr_Msg & "→元:" & tbl_chk1_4_1(0).LINE & ", 先:" & tbl_chk1_4_1(1).LINE
            GoTo CodeDBGet_Error
        End If
        If sResult = 0 Then
            funChkFurikae1_4_1 = 1
            iErr_Code = 1415
            GoTo Apl_Exit
        End If
    End If
'*** UPDATE ↑ Y.SIMIZU 2005/10/12 ﾗｲﾝ数ﾁｪｯｸ追加
    'パターン区分
    sErr_Msg = "ﾊﾟﾀｰﾝ区分ﾁｪｯｸ"
    If Mid(sChkCode, 14, 1) = "2" Then
        'ArANﾊﾟﾀｰﾝ区分ﾁｪｯｸ 08/12/21 ooba
        If InStr(sAdd_Msg, "ArAN") > 0 Then
            If Trim$(tbl_chk1_4_1(0).PATTERN) = "" And Trim$(tbl_chk1_4_1(1).PATTERN) <> "" Then
                funChkFurikae1_4_1 = 1
                iErr_Code = 1414
                GoTo Apl_Exit
            End If
        Else
            'マトリクス取得
            sResult = ""
            RET = funCodeDBGet("SB", "OS", tbl_chk1_4_1(0).PATTERN, 1, tbl_chk1_4_1(1).PATTERN, sResult)
            If RET <> 0 Then
                sErr_Msg = sAdd_Msg & sErr_Msg & "→元:" & tbl_chk1_4_1(0).PATTERN & ", 先:" & tbl_chk1_4_1(1).PATTERN
                GoTo CodeDBGet_Error
            End If
            If sResult = 0 Then
                funChkFurikae1_4_1 = 1
                iErr_Code = 1414
                GoTo Apl_Exit
            End If
        End If
    End If
        
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    'DK温度
    sErr_Msg = "DK温度ﾁｪｯｸ"
    If Mid(sChkCode, 15, 1) = "2" Then
        If Trim(tbl_chk1_4_1(0).HSXDKTMP) = "" And Trim(tbl_chk1_4_1(1).HSXDKTMP) = "" Then
        Else
            'マトリクス取得
            sResult = ""
            RET = funCodeDBGet(DKTMP_TBCMB005SYS, DKTMP_TBCMB005CLS, tbl_chk1_4_1(0).HSXDKTMP, 1, tbl_chk1_4_1(1).HSXDKTMP, sResult)
            If RET <> 0 Then
                sErr_Msg = sAdd_Msg & sErr_Msg & "→元:" & tbl_chk1_4_1(0).HSXDKTMP & ", 先:" & tbl_chk1_4_1(1).HSXDKTMP
                GoTo CodeDBGet_Error
            End If
            If sResult = 0 Then
                funChkFurikae1_4_1 = 1
                iErr_Code = 1416
                GoTo Apl_Exit
            End If
        End If
    End If
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

    'Cs保証位置チェック  '' add 0108
    sErr_Msg = "Cs保証位置ﾁｪｯｸ"
    If Mid(sChkCode, 16, 1) = "2" Then
        If (Trim(tbl_chk1_4_1(0).HSXCNKHI) <> "6" And Trim(tbl_chk1_4_1(0).HSXCNKHI) <> "9") And _
           (Trim(tbl_chk1_4_1(1).HSXCNKHI) = "6" Or Trim(tbl_chk1_4_1(1).HSXCNKHI) = "9") Then
            ''B保証品をT/B保証品に振り替えはエラー
            funChkFurikae1_4_1 = 1
            iErr_Code = 1417
            GoTo Apl_Exit
        Else
            ''上記以外は振替OK
        End If
    End If

    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Select Case iErr_Code
        Case 1401
            sErr_Msg = sAdd_Msg & "の保証方法が不一致の為、振替できません。"
        Case 1402
            sErr_Msg = sAdd_Msg & "の下限が不一致の為、振替できません。"
        Case 1403
            sErr_Msg = sAdd_Msg & "の上限が不一致の為、振替できません。"
        Case 1404
            sErr_Msg = sAdd_Msg & "の測定位置＿方が不一致の為、振替できません。"
        Case 1405
            sErr_Msg = sAdd_Msg & "の測定位置＿点が不一致の為、振替できません。"
        Case 1406
            sErr_Msg = sAdd_Msg & "の測定位置＿位が振替不可能です。"
        Case 1407
            sErr_Msg = sAdd_Msg & "の測定位置＿領が不一致の為、振替できません。"
        Case 1408
            sErr_Msg = sAdd_Msg & "の検査有無が不一致の為、振替できません。"
        Case 1409
            sErr_Msg = sAdd_Msg & "の熱処理法が不一致の為、振替できません。"
        Case 1410
            sErr_Msg = sAdd_Msg & "の測定条件が不一致の為、振替できません。"
        Case 1411
            sErr_Msg = sAdd_Msg & "の選択ＥＴ代が不一致の為、振替できません。"
        Case 1412
            sErr_Msg = sAdd_Msg & "の検査方法が不一致の為、振替できません。"
        Case 1413
            sErr_Msg = sAdd_Msg & "のライン数が不一致の為、振替できません。"
        Case 1414
            sErr_Msg = sAdd_Msg & "のパターン区分が振替不可能です。"
    '*** UPDATE ↓ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数対応
        Case 1415
            sErr_Msg = sAdd_Msg & "のGDライン数が振替不可能です。"
    '*** UPDATE ↑ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数対応
'--------------- 2008/08/25 INSERT START  By Systech ---------------
        Case 1416
            sErr_Msg = sAdd_Msg & "のDK温度が振替不可能です。"
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
        Case 1417   '' add 0108
            sErr_Msg = sAdd_Msg & "の検査頻度＿位が振替不可能です。"  '' add 0108
    End Select
    
    Exit Function
    
Apl_down:
    funChkFurikae1_4_1 = -4
    GoTo Apl_Exit
    
CodeDBGet_Error:
    If funChkFurikae1_4_1 = 0 Then
        funChkFurikae1_4_1 = -5
    End If
    GoTo Apl_Exit

End Function

'------------------------------------------------
' 振替元と振替先の先行評価項目仕様チェック
'------------------------------------------------

'概要      :振替元品番と振替先品番の先行評価項目仕様をチェックする。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sKeyID          ,I  ,String       :ﾌﾞﾛｯｸID、又は、SXL-ID
'          :tOld_Hinban     ,I  ,tFullHinban  :振替元品番(構造体)
'          :tNew_Hinban     ,I  ,tFullHinban  :振替先品番(構造体)
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/19 新規作成　SB

Public Function funChkFurikae1_5(sProccd As String, sBlockId As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iErr_Code As Integer, sErr_Msg As String) As Integer



    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql As String               'SQL全体
    Dim rs  As OraDynaset           'RecordSet

    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae1_5 = 0
    
    '------------------------------------------ 振替元品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-5 振替元品番仕様取得(" & tOld_Hinban.hinban & Format(tOld_Hinban.mnorevno, "00") & tOld_Hinban.factory & tOld_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E021.HWFRHWYS,E025.HWFONHWS,E025.HWFONSPT,  E029.HWFOF1HS,E029.HWFOF1SH,E029.HWFOF1SR,  E029.HWFOF1NS,E029.HWFOF1SZ,E029.HWFOF1ET,  E029.HWFOSF1PTK, E029.HWFOF2HS,   " & vbCrLf
    sql = sql & "       E029.HWFOF2SH,E029.HWFOF2SR,E029.HWFOF2NS,  E029.HWFOF2SZ,E029.HWFOF2ET,E029.HWFOSF2PTK,E029.HWFOF3HS,E029.HWFOF3SH,E029.HWFOF3SR,  E029.HWFOF3NS,   " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''    sql = sql & "       E029.HWFOF3SZ,E029.HWFOF3ET,E029.HWFOSF3PTK,E029.HWFOF4HS,E029.HWFOF4SH,E029.HWFOF4SR,  E029.HWFOF4NS,E029.HWFOF4SZ,E029.HWFOF4ET,  E029.HWFOSF4PTK, " & vbCrLf
    sql = sql & "       E029.HWFOF3SZ,E029.HWFOF3ET,E029.HWFOSF3PTK, " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
    ''残存酸素仕様取得追加　03/12/09 ooba   ''DSODﾊﾟﾀｰﾝ区分取得追加　04/07/29 ooba
    sql = sql & "       E025.HWFZOHWS,E025.HWFZONSW,E026.HWFDSOPTK, " & vbCrLf
    sql = sql & "       E029.HWFBM1HS,E029.HWFBM1SH,E029.HWFBM1ST,  E029.HWFBM1SR,E029.HWFBM1NS,E029.HWFBM1SZ,  E029.HWFBM1ET,E029.HWFBM2HS,E029.HWFBM2SH,  E029.HWFBM2ST,   " & vbCrLf
    sql = sql & "       E029.HWFBM2SR,E029.HWFBM2NS,E029.HWFBM2SZ,  E029.HWFBM2ET,E029.HWFBM3HS,E029.HWFBM3SH,  E029.HWFBM3ST,E029.HWFBM3SR,E029.HWFBM3NS,  E029.HWFBM3SZ,   " & vbCrLf
    sql = sql & "       E029.HWFBM3ET,E025.HWFOS1HS,E025.HWFOS1NS,  E025.HWFOS2HS,E025.HWFOS2NS,E025.HWFOS3HS,  E025.HWFOS3NS,E026.HWFDSOHS,E026.HWFDSONWY, E024.HWFMKHWS,   " & vbCrLf
    sql = sql & "       E024.HWFMKSPH,E024.HWFMKSPT,E024.HWFMKSPR,  E024.HWFMKNSW,E024.HWFMKSZY,E024.HWFMKCET,  E028.HWFSPVHS,E028.HWFSPVST,E028.HWFDLHWS,                   " & vbCrLf

''Upd Start 2005/06/16 (TCS)T.Terauchi  SPV9点対応
    sql = sql & "       E028.HWFSPVSH,E028.HWFSPVSI," & vbCrLf                    ''SPVFE
    sql = sql & "       E028.HWFDLSPH,E028.HWFDLSPT,E028.HWFDLSPI," & vbCrLf      ''拡散長
''Upd End   2005/06/16 (TCS)T.Terauchi  SPV9点対応

'*** UPDATE ↓ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数取得追加
    sql = sql & "       E036.HSXGDLINE,E036.HWFGDLINE," & vbCrLf
'*** UPDATE ↑ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数取得追加

    'ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞ／結晶GD仕様取得追加　05/07/29 ooba
    sql = sql & "       E036.BLOCKHFLAG," & vbCrLf
    sql = sql & "       E020.HSXDENHS,E020.HSXDENMN,E020.HSXDENMX,  E020.HSXDVDHS,E020.HSXDVDMNN,E020.HSXDVDMXN,E020.HSXLDLHS,E020.HSXLDLMN,E020.HSXLDLMX,                   " & vbCrLf
    'GD仕様取得追加　05/01/27 ooba
    
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    sql = sql & "       E026.HWFDENHS,E026.HWFDENMN,E026.HWFDENMX,  E026.HWFDVDHS,E026.HWFDVDMNN,E026.HWFDVDMXN,E026.HWFLDLHS,E026.HWFLDLMN,E026.HWFLDLMX,  E026.HWFGDKHN, E026.HWFGDSZY,  " & vbCrLf
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---

    ''検査頻度_抜ﾃﾞｰﾀ取得　04/04/13 ooba
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''    sql = sql & "       E021.HWFRKHNN, E025.HWFONKHN, E029.HWFOF1KN, E029.HWFOF2KN, E029.HWFOF3KN, E029.HWFOF4KN, E029.HWFBM1KN, E029.HWFBM2KN, E029.HWFBM3KN,               " & vbCrLf
    sql = sql & "       E021.HWFRKHNN, E025.HWFONKHN, E029.HWFOF1KN, E029.HWFOF2KN, E029.HWFOF3KN, E029.HWFBM1KN, E029.HWFBM2KN, E029.HWFBM3KN,               " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
    sql = sql & "       E025.HWFOS1KN, E025.HWFOS2KN, E025.HWFOS3KN, E026.HWFDSOKN, E024.HWFMKKHN, E028.HWFSPVKN, E028.HWFDLKHN, E025.HWFZOKHN                               " & vbCrLf
'↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.1 AN温度 振替可否チェック追加
    sql = sql & "       ,E025.HWFANTNP " & vbCrLf ' 品ＷＦＡＮ温度
'↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    'SPV仕様項目追加(PUA限,PUA率,Nr濃度仕様)　06/05/31 ooba
    sql = sql & "       ,E048.HWFSPVPUG,E048.HWFSPVPUR,E048.HWFDLPUG,E048.HWFDLPUR          " & vbCrLf
    sql = sql & "       ,E048.HWFNRHS,E048.HWFNRSH,E048.HWFNRST,E048.HWFNRSI,E048.HWFNRKN   " & vbCrLf
    sql = sql & "       ,E048.HWFNRPUG,E048.HWFNRPUR                                        " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    sql = sql & "       ,E048.HWFSIRDMX " & vbCrLf                   '軸状転位上限
    sql = sql & "       ,E048.HWFSIRDSZ " & vbCrLf                   '軸状転位測定条件
    sql = sql & "       ,E048.HWFSIRDHT " & vbCrLf                   '軸状転位保証方法＿対
    sql = sql & "       ,E048.HWFSIRDHS " & vbCrLf                   '軸状転位保証方法_処
    sql = sql & "       ,E048.HWFSIRDKM " & vbCrLf                   '軸状転位検査頻度＿枚
    sql = sql & "       ,E048.HWFSIRDKN " & vbCrLf                   '軸状転位検査頻度_抜
    sql = sql & "       ,E048.HWFSIRDKH " & vbCrLf                   '軸状転位検査頻度＿保
    sql = sql & "       ,E048.HWFSIRDKU " & vbCrLf                   '軸状転位検査頻度＿ウ
    sql = sql & "       ,E048.HWFSIRDPS " & vbCrLf                   '軸状転位TB保証位置
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
    sql = sql & "FROM   TBCME021 E021,TBCME025 E025,TBCME029 E029,TBCME028 E028,TBCME026 E026,TBCME024 E024,TBCME036 E036,TBCME020 E020,TBCME048 E048 " & vbCrLf
    sql = sql & "WHERE  E021.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E021.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E021.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E021.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E025.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E025.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E025.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E025.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E029.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E029.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E029.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E029.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E028.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E028.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E028.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E028.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E026.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E026.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E026.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E026.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E036.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E020.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E020.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E020.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E020.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E024.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E024.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E024.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E024.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E048.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E048.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E048.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E048.OPECOND   =   '" & tOld_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    '検査頻度_抜ﾃﾞｰﾀ追加　04/04/13 ooba
    Erase tbl_chk1_5
    With tbl_chk1_5(0)
        If IsNull(rs("BLOCKHFLAG")) = False Then .BLOCKHFLAG = rs("BLOCKHFLAG") Else .BLOCKHFLAG = " "      'ブロック単位保証フラグ　05/07/29 ooba
        'Rs
        If IsNull(rs("HWFRHWYS")) = False Then .HWFRHWYS = rs("HWFRHWYS") Else .HWFRHWYS = " "              '保証方法_対象
        If IsNull(rs("HWFRKHNN")) = False Then .HWFRKHNN = rs("HWFRKHNN") Else .HWFRKHNN = " "              '検査頻度_抜
        'Oi
        If IsNull(rs("HWFONHWS")) = False Then .HWFONHWS = rs("HWFONHWS") Else .HWFONHWS = " "              '保証方法_対象
        If IsNull(rs("HWFONSPT")) = False Then .HWFONSPT = rs("HWFONSPT") Else .HWFONSPT = " "              '測定位置_点    '08/01/29 ooba
        If IsNull(rs("HWFONKHN")) = False Then .HWFONKHN = rs("HWFONKHN") Else .HWFONKHN = " "              '検査頻度_抜
        'OSF1
        If IsNull(rs("HWFOF1HS")) = False Then .HWFOF1HS = rs("HWFOF1HS") Else .HWFOF1HS = " "              '保証方法_対象
        If IsNull(rs("HWFOF1SH")) = False Then .HWFOF1SH = rs("HWFOF1SH") Else .HWFOF1SH = " "              '測定位置_方
        If IsNull(rs("HWFOF1SR")) = False Then .HWFOF1SR = rs("HWFOF1SR") Else .HWFOF1SR = " "              '測定位置_領
        If IsNull(rs("HWFOF1NS")) = False Then .HWFOF1NS = rs("HWFOF1NS") Else .HWFOF1NS = " "              '熱処理法
        If IsNull(rs("HWFOF1SZ")) = False Then .HWFOF1SZ = rs("HWFOF1SZ") Else .HWFOF1SZ = " "              '測定条件
        If IsNull(rs("HWFOF1ET")) = False Then .HWFOF1ET = rs("HWFOF1ET") Else .HWFOF1ET = 0                '選択ET代
        If IsNull(rs("HWFOSF1PTK")) = False Then .HWFOSF1PTK = rs("HWFOSF1PTK") Else .HWFOSF1PTK = "4"      'パターン区分
        If IsNull(rs("HWFOF1KN")) = False Then .HWFOF1KN = rs("HWFOF1KN") Else .HWFOF1KN = " "              '検査頻度_抜
        'OSF2
        If IsNull(rs("HWFOF2HS")) = False Then .HWFOF2HS = rs("HWFOF2HS") Else .HWFOF2HS = " "              '保証方法_対象
        If IsNull(rs("HWFOF2SH")) = False Then .HWFOF2SH = rs("HWFOF2SH") Else .HWFOF2SH = " "              '測定位置_方
        If IsNull(rs("HWFOF2SR")) = False Then .HWFOF2SR = rs("HWFOF2SR") Else .HWFOF2SR = " "              '測定位置_領
        If IsNull(rs("HWFOF2NS")) = False Then .HWFOF2NS = rs("HWFOF2NS") Else .HWFOF2NS = " "              '熱処理法
        If IsNull(rs("HWFOF2SZ")) = False Then .HWFOF2SZ = rs("HWFOF2SZ") Else .HWFOF2SZ = " "              '測定条件
        If IsNull(rs("HWFOF2ET")) = False Then .HWFOF2ET = rs("HWFOF2ET") Else .HWFOF2ET = 0                '選択ET代
        If IsNull(rs("HWFOSF2PTK")) = False Then .HWFOSF2PTK = rs("HWFOSF2PTK") Else .HWFOSF2PTK = "4"      'パターン区分
        If IsNull(rs("HWFOF2KN")) = False Then .HWFOF2KN = rs("HWFOF2KN") Else .HWFOF2KN = " "              '検査頻度_抜
        'OSF3
        If IsNull(rs("HWFOF3HS")) = False Then .HWFOF3HS = rs("HWFOF3HS") Else .HWFOF3HS = " "              '保証方法_対象
        If IsNull(rs("HWFOF3SH")) = False Then .HWFOF3SH = rs("HWFOF3SH") Else .HWFOF3SH = " "              '測定位置_方
        If IsNull(rs("HWFOF3SR")) = False Then .HWFOF3SR = rs("HWFOF3SR") Else .HWFOF3SR = " "              '測定位置_領
        If IsNull(rs("HWFOF3NS")) = False Then .HWFOF3NS = rs("HWFOF3NS") Else .HWFOF3NS = " "              '熱処理法
        If IsNull(rs("HWFOF3SZ")) = False Then .HWFOF3SZ = rs("HWFOF3SZ") Else .HWFOF3SZ = " "              '測定条件
        If IsNull(rs("HWFOF3ET")) = False Then .HWFOF3ET = rs("HWFOF3ET") Else .HWFOF3ET = 0                '選択ET代
        If IsNull(rs("HWFOSF3PTK")) = False Then .HWFOSF3PTK = rs("HWFOSF3PTK") Else .HWFOSF3PTK = "4"      'パターン区分
        If IsNull(rs("HWFOF3KN")) = False Then .HWFOF3KN = rs("HWFOF3KN") Else .HWFOF3KN = " "              '検査頻度_抜
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''        'OSF4
'''        If IsNull(rs("HWFOF4HS")) = False Then .HWFOF4HS = rs("HWFOF4HS") Else .HWFOF4HS = " "              '保証方法_対象
'''        If IsNull(rs("HWFOF4SH")) = False Then .HWFOF4SH = rs("HWFOF4SH") Else .HWFOF4SH = " "              '測定位置_方
'''        If IsNull(rs("HWFOF4SR")) = False Then .HWFOF4SR = rs("HWFOF4SR") Else .HWFOF4SR = " "              '測定位置_領
'''        If IsNull(rs("HWFOF4NS")) = False Then .HWFOF4NS = rs("HWFOF4NS") Else .HWFOF4NS = " "              '熱処理法
'''        If IsNull(rs("HWFOF4SZ")) = False Then .HWFOF4SZ = rs("HWFOF4SZ") Else .HWFOF4SZ = " "              '測定条件
'''        If IsNull(rs("HWFOF4ET")) = False Then .HWFOF4ET = rs("HWFOF4ET") Else .HWFOF4ET = 0                '選択ET代
'''        If IsNull(rs("HWFOSF4PTK")) = False Then .HWFOSF4PTK = rs("HWFOSF4PTK") Else .HWFOSF4PTK = "4"      'パターン区分
'''        If IsNull(rs("HWFOF4KN")) = False Then .HWFOF4KN = rs("HWFOF4KN") Else .HWFOF4KN = " "              '検査頻度_抜

        'SIRD
        If IsNull(rs("HWFSIRDMX")) = False Then .HWFSIRDMX = rs("HWFSIRDMX") Else .HWFSIRDMX = "0"          '軸状転位上限
        If IsNull(rs("HWFSIRDSZ")) = False Then .HWFSIRDSZ = rs("HWFSIRDSZ") Else .HWFSIRDSZ = " "          '軸状転位測定条件
        If IsNull(rs("HWFSIRDHT")) = False Then .HWFSIRDHT = rs("HWFSIRDHT") Else .HWFSIRDHT = " "          '軸状転位保証方法＿対
        If IsNull(rs("HWFSIRDHS")) = False Then .HWFSIRDHS = rs("HWFSIRDHS") Else .HWFSIRDHS = " "          '軸状転位保証方法＿処
        If IsNull(rs("HWFSIRDKM")) = False Then .HWFSIRDKM = rs("HWFSIRDKM") Else .HWFSIRDKM = " "          '軸状転位検査頻度＿枚
        If IsNull(rs("HWFSIRDKN")) = False Then .HWFSIRDKN = rs("HWFSIRDKN") Else .HWFSIRDKN = " "          '軸状転位検査頻度＿抜
        If IsNull(rs("HWFSIRDKH")) = False Then .HWFSIRDKH = rs("HWFSIRDKH") Else .HWFSIRDKH = " "          '軸状転位検査頻度＿保
        If IsNull(rs("HWFSIRDKU")) = False Then .HWFSIRDKU = rs("HWFSIRDKU") Else .HWFSIRDKU = " "          '軸状転位検査頻度＿ウ
        If IsNull(rs("HWFSIRDPS")) = False Then .HWFSIRDPS = Trim(rs("HWFSIRDPS")) Else .HWFSIRDPS = " "    '軸状転位TB保証位置
        
        '「軸状転位TB保証位置」を判定し、「軸状転位検査頻度＿抜」に編集
        Select Case Trim(.HWFSIRDPS)
        Case "T"
            .HWFSIRDKN = "3"
        Case "B"
            .HWFSIRDKN = "4"
        Case "TB"
            .HWFSIRDKN = "6"
        Case Else
            .HWFSIRDKN = " "
        End Select
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
        'BMD1
        If IsNull(rs("HWFBM1HS")) = False Then .HWFBM1HS = rs("HWFBM1HS") Else .HWFBM1HS = " "              '保証方法_対象
        If IsNull(rs("HWFBM1SH")) = False Then .HWFBM1SH = rs("HWFBM1SH") Else .HWFBM1SH = " "              '測定位置_方
        If IsNull(rs("HWFBM1ST")) = False Then .HWFBM1ST = rs("HWFBM1ST") Else .HWFBM1ST = " "              '測定位置_点
        If IsNull(rs("HWFBM1SR")) = False Then .HWFBM1SR = rs("HWFBM1SR") Else .HWFBM1SR = " "              '測定位置_領
        If IsNull(rs("HWFBM1NS")) = False Then .HWFBM1NS = rs("HWFBM1NS") Else .HWFBM1NS = " "              '熱処理法
        If IsNull(rs("HWFBM1SZ")) = False Then .HWFBM1SZ = rs("HWFBM1SZ") Else .HWFBM1SZ = " "              '測定条件
        If IsNull(rs("HWFBM1ET")) = False Then .HWFBM1ET = rs("HWFBM1ET") Else .HWFBM1ET = 0                '選択ET代
        If IsNull(rs("HWFBM1KN")) = False Then .HWFBM1KN = rs("HWFBM1KN") Else .HWFBM1KN = " "              '検査頻度_抜
        'BMD2
        If IsNull(rs("HWFBM2HS")) = False Then .HWFBM2HS = rs("HWFBM2HS") Else .HWFBM2HS = " "              '保証方法_対象
        If IsNull(rs("HWFBM2SH")) = False Then .HWFBM2SH = rs("HWFBM2SH") Else .HWFBM2SH = " "              '測定位置_方
        If IsNull(rs("HWFBM2ST")) = False Then .HWFBM2ST = rs("HWFBM2ST") Else .HWFBM2ST = " "              '測定位置_点
        If IsNull(rs("HWFBM2SR")) = False Then .HWFBM2SR = rs("HWFBM2SR") Else .HWFBM2SR = " "              '測定位置_領
        If IsNull(rs("HWFBM2NS")) = False Then .HWFBM2NS = rs("HWFBM2NS") Else .HWFBM2NS = " "              '熱処理法
        If IsNull(rs("HWFBM2SZ")) = False Then .HWFBM2SZ = rs("HWFBM2SZ") Else .HWFBM2SZ = " "              '測定条件
        If IsNull(rs("HWFBM2ET")) = False Then .HWFBM2ET = rs("HWFBM2ET") Else .HWFBM2ET = 0                '選択ET代
        If IsNull(rs("HWFBM2KN")) = False Then .HWFBM2KN = rs("HWFBM2KN") Else .HWFBM2KN = " "              '検査頻度_抜
        'BMD3
        If IsNull(rs("HWFBM3HS")) = False Then .HWFBM3HS = rs("HWFBM3HS") Else .HWFBM3HS = " "              '保証方法_対象
        If IsNull(rs("HWFBM3SH")) = False Then .HWFBM3SH = rs("HWFBM3SH") Else .HWFBM3SH = " "              '測定位置_方
        If IsNull(rs("HWFBM3ST")) = False Then .HWFBM3ST = rs("HWFBM3ST") Else .HWFBM3ST = " "              '測定位置_点
        If IsNull(rs("HWFBM3SR")) = False Then .HWFBM3SR = rs("HWFBM3SR") Else .HWFBM3SR = " "              '測定位置_領
        If IsNull(rs("HWFBM3NS")) = False Then .HWFBM3NS = rs("HWFBM3NS") Else .HWFBM3NS = " "              '熱処理法
        If IsNull(rs("HWFBM3SZ")) = False Then .HWFBM3SZ = rs("HWFBM3SZ") Else .HWFBM3SZ = " "              '測定条件
        If IsNull(rs("HWFBM3ET")) = False Then .HWFBM3ET = rs("HWFBM3ET") Else .HWFBM3ET = 0                '選択ET代
        If IsNull(rs("HWFBM3KN")) = False Then .HWFBM3KN = rs("HWFBM3KN") Else .HWFBM3KN = " "              '検査頻度_抜
        'DOI1
        If IsNull(rs("HWFOS1HS")) = False Then .HWFOS1HS = rs("HWFOS1HS") Else .HWFOS1HS = " "              '保証方法_対象
        If IsNull(rs("HWFOS1NS")) = False Then .HWFOS1NS = rs("HWFOS1NS") Else .HWFOS1NS = " "              '熱処理法
        If IsNull(rs("HWFOS1KN")) = False Then .HWFOS1KN = rs("HWFOS1KN") Else .HWFOS1KN = " "              '検査頻度_抜
        'DOI2
        If IsNull(rs("HWFOS2HS")) = False Then .HWFOS2HS = rs("HWFOS2HS") Else .HWFOS2HS = " "              '保証方法_対象
        If IsNull(rs("HWFOS2NS")) = False Then .HWFOS2NS = rs("HWFOS2NS") Else .HWFOS2NS = " "              '熱処理法
        If IsNull(rs("HWFOS2KN")) = False Then .HWFOS2KN = rs("HWFOS2KN") Else .HWFOS2KN = " "              '検査頻度_抜
        'DOI3
        If IsNull(rs("HWFOS3HS")) = False Then .HWFOS3HS = rs("HWFOS3HS") Else .HWFOS3HS = " "              '保証方法_対象
        If IsNull(rs("HWFOS3NS")) = False Then .HWFOS3NS = rs("HWFOS3NS") Else .HWFOS3NS = " "              '熱処理法
        If IsNull(rs("HWFOS3KN")) = False Then .HWFOS3KN = rs("HWFOS3KN") Else .HWFOS3KN = " "              '検査頻度_抜
        'DSOD
        If IsNull(rs("HWFDSOHS")) = False Then .HWFDSOHS = rs("HWFDSOHS") Else .HWFDSOHS = " "              '保証方法_対象
        If IsNull(rs("HWFDSONWY")) = False Then .HWFDSONWY = rs("HWFDSONWY") Else .HWFDSONWY = " "          '熱処理法
        If IsNull(rs("HWFDSOKN")) = False Then .HWFDSOKN = rs("HWFDSOKN") Else .HWFDSOKN = " "              '検査頻度_抜
        If IsNull(rs("HWFDSOPTK")) = False Then .HWFDSOPTK = rs("HWFDSOPTK") Else .HWFDSOPTK = " "          'パターン区分　04/07/29 ooba
        'DZ
        If IsNull(rs("HWFMKHWS")) = False Then .HWFMKHWS = rs("HWFMKHWS") Else .HWFMKHWS = " "              '保証方法_対象
        If IsNull(rs("HWFMKSPH")) = False Then .HWFMKSPH = rs("HWFMKSPH") Else .HWFMKSPH = " "              '測定位置_方
        If IsNull(rs("HWFMKSPT")) = False Then .HWFMKSPT = rs("HWFMKSPT") Else .HWFMKSPT = " "              '測定位置_点
        If IsNull(rs("HWFMKSPR")) = False Then .HWFMKSPR = rs("HWFMKSPR") Else .HWFMKSPR = " "              '測定位置_領
        If IsNull(rs("HWFMKNSW")) = False Then .HWFMKNSW = rs("HWFMKNSW") Else .HWFMKNSW = " "              '熱処理法
        If IsNull(rs("HWFMKSZY")) = False Then .HWFMKSZY = rs("HWFMKSZY") Else .HWFMKSZY = " "              '測定条件
        If IsNull(rs("HWFMKCET")) = False Then .HWFMKCET = rs("HWFMKCET") Else .HWFMKCET = 0                '選択ET代
        If IsNull(rs("HWFMKKHN")) = False Then .HWFMKKHN = rs("HWFMKKHN") Else .HWFMKKHN = " "              '検査頻度_抜
        'SPVFE
        If IsNull(rs("HWFSPVHS")) = False Then .HWFSPVHS = rs("HWFSPVHS") Else .HWFSPVHS = " "              '保証方法_対象
        If IsNull(rs("HWFSPVST")) = False Then .HWFSPVST = rs("HWFSPVST") Else .HWFSPVST = " "              '測定位置＿点
        If IsNull(rs("HWFSPVKN")) = False Then .HWFSPVKN = rs("HWFSPVKN") Else .HWFSPVKN = " "              '検査頻度_抜
        '拡散長
        If IsNull(rs("HWFDLHWS")) = False Then .HWFDLHWS = rs("HWFDLHWS") Else .HWFDLHWS = " "              '保証方法_対象
        If IsNull(rs("HWFDLKHN")) = False Then .HWFDLKHN = rs("HWFDLKHN") Else .HWFDLKHN = " "              '検査頻度_抜
        
    ''Upd Start 2005/06/16 (TCS)T.Terauchi  SPV9点対応
        'SPVFE
        If IsNull(rs("HWFSPVSH")) = False Then .HWFSPVSH = rs("HWFSPVSH") Else .HWFSPVSH = " "              '測定位置＿方
        If IsNull(rs("HWFSPVSI")) = False Then .HWFSPVSI = rs("HWFSPVSI") Else .HWFSPVSI = " "              '測定位置＿位
        '拡散長
        If IsNull(rs("HWFDLSPH")) = False Then .HWFDLSPH = rs("HWFDLSPH") Else .HWFDLSPH = " "              '測定位置＿方
        If IsNull(rs("HWFDLSPT")) = False Then .HWFDLSPT = rs("HWFDLSPT") Else .HWFDLSPT = " "              '測定位置＿点
        If IsNull(rs("HWFDLSPI")) = False Then .HWFDLSPI = rs("HWFDLSPI") Else .HWFDLSPI = " "              '測定位置＿位
    ''Upd End   2005/06/16 (TCS)T.Terauchi  SPV9点対応
        
        ''06/05/31 ooba START ==================================================================>
        'SPVFE
        If IsNull(rs("HWFSPVPUG")) = False Then .HWFSPVPUG = rs("HWFSPVPUG") Else .HWFSPVPUG = -1           'PUA限
        If IsNull(rs("HWFSPVPUR")) = False Then .HWFSPVPUR = rs("HWFSPVPUR") Else .HWFSPVPUR = -1           'PUA率
        '拡散長
        If IsNull(rs("HWFDLPUG")) = False Then .HWFDLPUG = rs("HWFDLPUG") Else .HWFDLPUG = -1               'PUA限
        If IsNull(rs("HWFDLPUR")) = False Then .HWFDLPUR = rs("HWFDLPUR") Else .HWFDLPUR = -1               'PUA率
        'SPVNR
        If IsNull(rs("HWFNRHS")) = False Then .HWFNRHS = rs("HWFNRHS") Else .HWFNRHS = " "                  '保証方法＿対象
        If IsNull(rs("HWFNRSH")) = False Then .HWFNRSH = rs("HWFNRSH") Else .HWFNRSH = " "                  '測定位置＿方
        If IsNull(rs("HWFNRST")) = False Then .HWFNRST = rs("HWFNRST") Else .HWFNRST = " "                  '測定位置＿点
        If IsNull(rs("HWFNRSI")) = False Then .HWFNRSI = rs("HWFNRSI") Else .HWFNRSI = " "                  '測定位置＿位
        If IsNull(rs("HWFNRKN")) = False Then .HWFNRKN = rs("HWFNRKN") Else .HWFNRKN = " "                  '検査頻度＿抜
        If IsNull(rs("HWFNRPUG")) = False Then .HWFNRPUG = rs("HWFNRPUG") Else .HWFNRPUG = -1               'PUA限
        If IsNull(rs("HWFNRPUR")) = False Then .HWFNRPUR = rs("HWFNRPUR") Else .HWFNRPUR = -1               'PUA率
        ''06/05/31 ooba END ====================================================================>
        
        'AOi        '残存酸素追加　03/12/09 ooba
        If IsNull(rs("HWFZOHWS")) = False Then .HWFZOHWS = rs("HWFZOHWS") Else .HWFZOHWS = " "              '保証方法_対象
        If IsNull(rs("HWFZONSW")) = False Then .HWFZONSW = rs("HWFZONSW") Else .HWFZONSW = " "              '熱処理法
        If IsNull(rs("HWFZOKHN")) = False Then .HWFZOKHN = rs("HWFZOKHN") Else .HWFZOKHN = " "              '検査頻度_抜
        'DEN        'DEN追加　05/01/27 ooba
        If IsNull(rs("HWFDENHS")) = False Then .HWFDENHS = rs("HWFDENHS") Else .HWFDENHS = " "              '保証方法_対象
        If IsNull(rs("HWFDENMN")) = False Then .HWFDENMN = rs("HWFDENMN") Else .HWFDENMN = 0                '下限
        If IsNull(rs("HWFDENMX")) = False Then .HWFDENMX = rs("HWFDENMX") Else .HWFDENMX = 0                '上限
        'DVD2       'DVD2追加　05/01/27 ooba
        If IsNull(rs("HWFDVDHS")) = False Then .HWFDVDHS = rs("HWFDVDHS") Else .HWFDVDHS = " "              '保証方法_対象
        If IsNull(rs("HWFDVDMNN")) = False Then .HWFDVDMNN = rs("HWFDVDMNN") Else .HWFDVDMNN = 0            '下限
        If IsNull(rs("HWFDVDMXN")) = False Then .HWFDVDMXN = rs("HWFDVDMXN") Else .HWFDVDMXN = 0            '上限
        'L/DL       'L/DL追加　05/01/27 ooba
        If IsNull(rs("HWFLDLHS")) = False Then .HWFLDLHS = rs("HWFLDLHS") Else .HWFLDLHS = " "              '保証方法_対象
        If IsNull(rs("HWFLDLMN")) = False Then .HWFLDLMN = rs("HWFLDLMN") Else .HWFLDLMN = 0                '下限
        If IsNull(rs("HWFLDLMX")) = False Then .HWFLDLMX = rs("HWFLDLMX") Else .HWFLDLMX = 0                '上限
        If IsNull(rs("HWFGDKHN")) = False Then .HWFGDKHN = rs("HWFGDKHN") Else .HWFGDKHN = " "              '検査頻度_抜
    '*** UPDATE ↓ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数追加
        If IsNull(rs("HWFGDLINE")) = False Then .HWFGDLINE = rs("HWFGDLINE") Else .HWFGDLINE = " "               '測定条件
    '*** UPDATE ↑ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数追加
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
        If IsNull(rs("HWFGDSZY")) = False Then .HWFGDSZY = rs("HWFGDSZY") Else .HWFGDSZY = " "               '測定条件
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
    '↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    '2.1.1 AN温度 振替可否チェック追加
        If IsNull(rs("HWFANTNP")) = False Then .HWFANTNP = rs("HWFANTNP") Else .HWFANTNP = 0                '品ＷＦＡＮ温度
    '↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    End With
    
    '結晶GD仕様データセット　05/07/29 ooba
    With tbl_chk1_5_SXGD
        'DEN
        If IsNull(rs("HSXDENHS")) = False Then .HWFDENHS = rs("HSXDENHS") Else .HWFDENHS = " "              '保証方法_対象
        If IsNull(rs("HSXDENMN")) = False Then .HWFDENMN = rs("HSXDENMN") Else .HWFDENMN = 0                '下限
        If IsNull(rs("HSXDENMX")) = False Then .HWFDENMX = rs("HSXDENMX") Else .HWFDENMX = 0                '上限
        'DVD2
        If IsNull(rs("HSXDVDHS")) = False Then .HWFDVDHS = rs("HSXDVDHS") Else .HWFDVDHS = " "              '保証方法_対象
        If IsNull(rs("HSXDVDMNN")) = False Then .HWFDVDMNN = rs("HSXDVDMNN") Else .HWFDVDMNN = 0            '下限
        If IsNull(rs("HSXDVDMXN")) = False Then .HWFDVDMXN = rs("HSXDVDMXN") Else .HWFDVDMXN = 0            '上限
        'L/DL
        If IsNull(rs("HSXLDLHS")) = False Then .HWFLDLHS = rs("HSXLDLHS") Else .HWFLDLHS = " "              '保証方法_対象
        If IsNull(rs("HSXLDLMN")) = False Then .HWFLDLMN = rs("HSXLDLMN") Else .HWFLDLMN = 0                '下限
        If IsNull(rs("HSXLDLMX")) = False Then .HWFLDLMX = rs("HSXLDLMX") Else .HWFLDLMX = 0                '上限
        
        If IsNull(rs("HSXGDLINE")) = False Then .HWFGDLINE = rs("HSXGDLINE") Else .HWFGDLINE = " "          'ライン数
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
        If IsNull(rs("HWFGDSZY")) = False Then .HWFGDSZY = rs("HWFGDSZY") Else .HWFGDSZY = " "               '測定条件
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
    
    End With
    
    Set rs = Nothing
    '------------------------------------------ 振替先品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-5 振替先品番仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E021.HWFRHWYS,E025.HWFONHWS,E025.HWFONSPT,  E029.HWFOF1HS,E029.HWFOF1SH,E029.HWFOF1SR,  E029.HWFOF1NS,E029.HWFOF1SZ,E029.HWFOF1ET,  E029.HWFOSF1PTK, E029.HWFOF2HS,   " & vbCrLf
    sql = sql & "       E029.HWFOF2SH,E029.HWFOF2SR,E029.HWFOF2NS,  E029.HWFOF2SZ,E029.HWFOF2ET,E029.HWFOSF2PTK,E029.HWFOF3HS,E029.HWFOF3SH,E029.HWFOF3SR,  E029.HWFOF3NS,   " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''    sql = sql & "       E029.HWFOF3SZ,E029.HWFOF3ET,E029.HWFOSF3PTK,E029.HWFOF4HS,E029.HWFOF4SH,E029.HWFOF4SR,  E029.HWFOF4NS,E029.HWFOF4SZ,E029.HWFOF4ET,  E029.HWFOSF4PTK, " & vbCrLf
    sql = sql & "       E029.HWFOF3SZ,E029.HWFOF3ET,E029.HWFOSF3PTK, " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
    ''残存酸素仕様取得追加　03/12/09 ooba   ''DSODﾊﾟﾀｰﾝ区分取得追加　04/07/29 ooba
    sql = sql & "       E025.HWFZOHWS,E025.HWFZONSW,E026.HWFDSOPTK, " & vbCrLf
    sql = sql & "       E029.HWFBM1HS,E029.HWFBM1SH,E029.HWFBM1ST,  E029.HWFBM1SR,E029.HWFBM1NS,E029.HWFBM1SZ,  E029.HWFBM1ET,E029.HWFBM2HS,E029.HWFBM2SH,  E029.HWFBM2ST,   " & vbCrLf
    sql = sql & "       E029.HWFBM2SR,E029.HWFBM2NS,E029.HWFBM2SZ,  E029.HWFBM2ET,E029.HWFBM3HS,E029.HWFBM3SH,  E029.HWFBM3ST,E029.HWFBM3SR,E029.HWFBM3NS,  E029.HWFBM3SZ,   " & vbCrLf
    sql = sql & "       E029.HWFBM3ET,E025.HWFOS1HS,E025.HWFOS1NS,  E025.HWFOS2HS,E025.HWFOS2NS,E025.HWFOS3HS,  E025.HWFOS3NS,E026.HWFDSOHS,E026.HWFDSONWY, E024.HWFMKHWS,   " & vbCrLf
    sql = sql & "       E024.HWFMKSPH,E024.HWFMKSPT,E024.HWFMKSPR,  E024.HWFMKNSW,E024.HWFMKSZY,E024.HWFMKCET,  E028.HWFSPVHS,E028.HWFSPVST,E028.HWFDLHWS,                   " & vbCrLf
    
''Upd Start 2005/06/16 (TCS)T.Terauchi  SPV9点対応
    sql = sql & "       E028.HWFSPVSH,E028.HWFSPVSI," & vbCrLf                    ''SPVFE
    sql = sql & "       E028.HWFDLSPH,E028.HWFDLSPT,E028.HWFDLSPI," & vbCrLf      ''拡散長
''Upd End   2005/06/16 (TCS)T.Terauchi  SPV9点対応
'*** UPDATE ↓ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数取得追加
    sql = sql & "       E036.HWFGDLINE," & vbCrLf
'*** UPDATE ↑ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数取得追加
    'ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞ取得追加　05/07/29 ooba
    sql = sql & "       E036.BLOCKHFLAG," & vbCrLf
    'GD仕様取得追加　05/01/27 ooba
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    sql = sql & "       E026.HWFDENHS,E026.HWFDENMN,E026.HWFDENMX,  E026.HWFDVDHS,E026.HWFDVDMNN,E026.HWFDVDMXN,E026.HWFLDLHS,E026.HWFLDLMN,E026.HWFLDLMX,  E026.HWFGDKHN, E026.HWFGDSZY,  " & vbCrLf
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
    ''検査頻度_抜ﾃﾞｰﾀ取得　04/04/13 ooba
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''    sql = sql & "       E021.HWFRKHNN, E025.HWFONKHN, E029.HWFOF1KN, E029.HWFOF2KN, E029.HWFOF3KN, E029.HWFOF4KN, E029.HWFBM1KN, E029.HWFBM2KN, E029.HWFBM3KN,               " & vbCrLf
    sql = sql & "       E021.HWFRKHNN, E025.HWFONKHN, E029.HWFOF1KN, E029.HWFOF2KN, E029.HWFOF3KN, E029.HWFBM1KN, E029.HWFBM2KN, E029.HWFBM3KN,               " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
    sql = sql & "       E025.HWFOS1KN, E025.HWFOS2KN, E025.HWFOS3KN, E026.HWFDSOKN, E024.HWFMKKHN, E028.HWFSPVKN, E028.HWFDLKHN, E025.HWFZOKHN                               " & vbCrLf
'↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.1 AN温度 振替可否チェック追加
    sql = sql & "       ,E025.HWFANTNP " & vbCrLf ' 品ＷＦＡＮ温度
'↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    'SPV仕様項目追加(PUA限,PUA率,Nr濃度仕様)　06/05/31 ooba
    sql = sql & "       ,E048.HWFSPVPUG,E048.HWFSPVPUR,E048.HWFDLPUG,E048.HWFDLPUR          " & vbCrLf
    sql = sql & "       ,E048.HWFNRHS,E048.HWFNRSH,E048.HWFNRST,E048.HWFNRSI,E048.HWFNRKN   " & vbCrLf
    sql = sql & "       ,E048.HWFNRPUG,E048.HWFNRPUR                                        " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    sql = sql & "       ,E048.HWFSIRDMX " & vbCrLf                   '軸状転位上限
    sql = sql & "       ,E048.HWFSIRDSZ " & vbCrLf                   '軸状転位測定条件
    sql = sql & "       ,E048.HWFSIRDHT " & vbCrLf                   '軸状転位保証方法＿対
    sql = sql & "       ,E048.HWFSIRDHS " & vbCrLf                   '軸状転位保証方法_処
    sql = sql & "       ,E048.HWFSIRDKM " & vbCrLf                   '軸状転位検査頻度＿枚
    sql = sql & "       ,E048.HWFSIRDKN " & vbCrLf                   '軸状転位検査頻度_抜
    sql = sql & "       ,E048.HWFSIRDKH " & vbCrLf                   '軸状転位検査頻度＿保
    sql = sql & "       ,E048.HWFSIRDKU " & vbCrLf                   '軸状転位検査頻度＿ウ
    sql = sql & "       ,E048.HWFSIRDPS " & vbCrLf                   '軸状転位TB保証位置
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
    sql = sql & "FROM   TBCME021 E021,TBCME025 E025,TBCME029 E029,TBCME028 E028,TBCME026 E026,TBCME024 E024,TBCME036 E036,TBCME048 E048 " & vbCrLf
    sql = sql & "WHERE  E021.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E021.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E021.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E021.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E025.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E025.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E025.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E025.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E029.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E029.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E029.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E029.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E028.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E028.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E028.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E028.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E026.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E026.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E026.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E026.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E036.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E024.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E024.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E024.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E024.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E048.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E048.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E048.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E048.OPECOND   =   '" & tNew_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    With tbl_chk1_5(1)
        If IsNull(rs("BLOCKHFLAG")) = False Then .BLOCKHFLAG = rs("BLOCKHFLAG") Else .BLOCKHFLAG = " "      'ブロック単位保証フラグ　05/07/29 ooba
        'Rs
        If IsNull(rs("HWFRHWYS")) = False Then .HWFRHWYS = rs("HWFRHWYS") Else .HWFRHWYS = " "              '保証方法_対象
        If IsNull(rs("HWFRKHNN")) = False Then .HWFRKHNN = rs("HWFRKHNN") Else .HWFRKHNN = " "              '検査頻度_抜
        'Oi
        If IsNull(rs("HWFONHWS")) = False Then .HWFONHWS = rs("HWFONHWS") Else .HWFONHWS = " "              '保証方法_対象
        If IsNull(rs("HWFONSPT")) = False Then .HWFONSPT = rs("HWFONSPT") Else .HWFONSPT = " "              '測定位置_点    '08/01/29 ooba
        If IsNull(rs("HWFONKHN")) = False Then .HWFONKHN = rs("HWFONKHN") Else .HWFONKHN = " "              '検査頻度_抜
        'OSF1
        If IsNull(rs("HWFOF1HS")) = False Then .HWFOF1HS = rs("HWFOF1HS") Else .HWFOF1HS = " "              '保証方法_対象
        If IsNull(rs("HWFOF1SH")) = False Then .HWFOF1SH = rs("HWFOF1SH") Else .HWFOF1SH = " "              '測定位置_方
        If IsNull(rs("HWFOF1SR")) = False Then .HWFOF1SR = rs("HWFOF1SR") Else .HWFOF1SR = " "              '測定位置_領
        If IsNull(rs("HWFOF1NS")) = False Then .HWFOF1NS = rs("HWFOF1NS") Else .HWFOF1NS = " "              '熱処理法
        If IsNull(rs("HWFOF1SZ")) = False Then .HWFOF1SZ = rs("HWFOF1SZ") Else .HWFOF1SZ = " "              '測定条件
        If IsNull(rs("HWFOF1ET")) = False Then .HWFOF1ET = rs("HWFOF1ET") Else .HWFOF1ET = 0                '選択ET代
        If IsNull(rs("HWFOSF1PTK")) = False Then .HWFOSF1PTK = rs("HWFOSF1PTK") Else .HWFOSF1PTK = "4"      'パターン区分
        If IsNull(rs("HWFOF1KN")) = False Then .HWFOF1KN = rs("HWFOF1KN") Else .HWFOF1KN = " "              '検査頻度_抜
        'OSF2
        If IsNull(rs("HWFOF2HS")) = False Then .HWFOF2HS = rs("HWFOF2HS") Else .HWFOF2HS = " "              '保証方法_対象
        If IsNull(rs("HWFOF2SH")) = False Then .HWFOF2SH = rs("HWFOF2SH") Else .HWFOF2SH = " "              '測定位置_方
        If IsNull(rs("HWFOF2SR")) = False Then .HWFOF2SR = rs("HWFOF2SR") Else .HWFOF2SR = " "              '測定位置_領
        If IsNull(rs("HWFOF2NS")) = False Then .HWFOF2NS = rs("HWFOF2NS") Else .HWFOF2NS = " "              '熱処理法
        If IsNull(rs("HWFOF2SZ")) = False Then .HWFOF2SZ = rs("HWFOF2SZ") Else .HWFOF2SZ = " "              '測定条件
        If IsNull(rs("HWFOF2ET")) = False Then .HWFOF2ET = rs("HWFOF2ET") Else .HWFOF2ET = 0                '選択ET代
        If IsNull(rs("HWFOSF2PTK")) = False Then .HWFOSF2PTK = rs("HWFOSF2PTK") Else .HWFOSF2PTK = "4"      'パターン区分
        If IsNull(rs("HWFOF2KN")) = False Then .HWFOF2KN = rs("HWFOF2KN") Else .HWFOF2KN = " "              '検査頻度_抜
        'OSF3
        If IsNull(rs("HWFOF3HS")) = False Then .HWFOF3HS = rs("HWFOF3HS") Else .HWFOF3HS = " "              '保証方法_対象
        If IsNull(rs("HWFOF3SH")) = False Then .HWFOF3SH = rs("HWFOF3SH") Else .HWFOF3SH = " "              '測定位置_方
        If IsNull(rs("HWFOF3SR")) = False Then .HWFOF3SR = rs("HWFOF3SR") Else .HWFOF3SR = " "              '測定位置_領
        If IsNull(rs("HWFOF3NS")) = False Then .HWFOF3NS = rs("HWFOF3NS") Else .HWFOF3NS = " "              '熱処理法
        If IsNull(rs("HWFOF3SZ")) = False Then .HWFOF3SZ = rs("HWFOF3SZ") Else .HWFOF3SZ = " "              '測定条件
        If IsNull(rs("HWFOF3ET")) = False Then .HWFOF3ET = rs("HWFOF3ET") Else .HWFOF3ET = 0                '選択ET代
        If IsNull(rs("HWFOSF3PTK")) = False Then .HWFOSF3PTK = rs("HWFOSF3PTK") Else .HWFOSF3PTK = "4"      'パターン区分
        If IsNull(rs("HWFOF3KN")) = False Then .HWFOF3KN = rs("HWFOF3KN") Else .HWFOF3KN = " "              '検査頻度_抜
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''        'OSF4
'''        If IsNull(rs("HWFOF4HS")) = False Then .HWFOF4HS = rs("HWFOF4HS") Else .HWFOF4HS = " "              '保証方法_対象
'''        If IsNull(rs("HWFOF4SH")) = False Then .HWFOF4SH = rs("HWFOF4SH") Else .HWFOF4SH = " "              '測定位置_方
'''        If IsNull(rs("HWFOF4SR")) = False Then .HWFOF4SR = rs("HWFOF4SR") Else .HWFOF4SR = " "              '測定位置_領
'''        If IsNull(rs("HWFOF4NS")) = False Then .HWFOF4NS = rs("HWFOF4NS") Else .HWFOF4NS = " "              '熱処理法
'''        If IsNull(rs("HWFOF4SZ")) = False Then .HWFOF4SZ = rs("HWFOF4SZ") Else .HWFOF4SZ = " "              '測定条件
'''        If IsNull(rs("HWFOF4ET")) = False Then .HWFOF4ET = rs("HWFOF4ET") Else .HWFOF4ET = 0                '選択ET代
'''        If IsNull(rs("HWFOSF4PTK")) = False Then .HWFOSF4PTK = rs("HWFOSF4PTK") Else .HWFOSF4PTK = "4"      'パターン区分
'''        If IsNull(rs("HWFOF4KN")) = False Then .HWFOF4KN = rs("HWFOF4KN") Else .HWFOF4KN = " "              '検査頻度_抜

        'SIRD
        If IsNull(rs("HWFSIRDMX")) = False Then .HWFSIRDMX = rs("HWFSIRDMX") Else .HWFSIRDMX = "0"          '軸状転位上限
        If IsNull(rs("HWFSIRDSZ")) = False Then .HWFSIRDSZ = rs("HWFSIRDSZ") Else .HWFSIRDSZ = " "          '軸状転位測定条件
        If IsNull(rs("HWFSIRDHT")) = False Then .HWFSIRDHT = rs("HWFSIRDHT") Else .HWFSIRDHT = " "          '軸状転位保証方法＿対
        If IsNull(rs("HWFSIRDHS")) = False Then .HWFSIRDHS = rs("HWFSIRDHS") Else .HWFSIRDHS = " "          '軸状転位保証方法＿処
        If IsNull(rs("HWFSIRDKM")) = False Then .HWFSIRDKM = rs("HWFSIRDKM") Else .HWFSIRDKM = " "          '軸状転位検査頻度＿枚
        If IsNull(rs("HWFSIRDKN")) = False Then .HWFSIRDKN = rs("HWFSIRDKN") Else .HWFSIRDKN = " "          '軸状転位検査頻度＿抜
        If IsNull(rs("HWFSIRDKH")) = False Then .HWFSIRDKH = rs("HWFSIRDKH") Else .HWFSIRDKH = " "          '軸状転位検査頻度＿保
        If IsNull(rs("HWFSIRDKU")) = False Then .HWFSIRDKU = rs("HWFSIRDKU") Else .HWFSIRDKU = " "          '軸状転位検査頻度＿ウ
        If IsNull(rs("HWFSIRDPS")) = False Then .HWFSIRDPS = Trim(rs("HWFSIRDPS")) Else .HWFSIRDPS = " "    '軸状転位TB保証位置
        
        '「軸状転位TB保証位置」を判定し、「軸状転位検査頻度＿抜」に編集
        Select Case Trim(.HWFSIRDPS)
        Case "T"
            .HWFSIRDKN = "3"
        Case "B"
            .HWFSIRDKN = "4"
        Case "TB"
            .HWFSIRDKN = "6"
        Case Else
            .HWFSIRDKN = " "
        End Select
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
        'BMD1
        If IsNull(rs("HWFBM1HS")) = False Then .HWFBM1HS = rs("HWFBM1HS") Else .HWFBM1HS = " "              '保証方法_対象
        If IsNull(rs("HWFBM1SH")) = False Then .HWFBM1SH = rs("HWFBM1SH") Else .HWFBM1SH = " "              '測定位置_方
        If IsNull(rs("HWFBM1ST")) = False Then .HWFBM1ST = rs("HWFBM1ST") Else .HWFBM1ST = " "              '測定位置_点
        If IsNull(rs("HWFBM1SR")) = False Then .HWFBM1SR = rs("HWFBM1SR") Else .HWFBM1SR = " "              '測定位置_領
        If IsNull(rs("HWFBM1NS")) = False Then .HWFBM1NS = rs("HWFBM1NS") Else .HWFBM1NS = " "              '熱処理法
        If IsNull(rs("HWFBM1SZ")) = False Then .HWFBM1SZ = rs("HWFBM1SZ") Else .HWFBM1SZ = " "              '測定条件
        If IsNull(rs("HWFBM1ET")) = False Then .HWFBM1ET = rs("HWFBM1ET") Else .HWFBM1ET = 0                '選択ET代
        If IsNull(rs("HWFBM1KN")) = False Then .HWFBM1KN = rs("HWFBM1KN") Else .HWFBM1KN = " "              '検査頻度_抜
        'BMD2
        If IsNull(rs("HWFBM2HS")) = False Then .HWFBM2HS = rs("HWFBM2HS") Else .HWFBM2HS = " "              '保証方法_対象
        If IsNull(rs("HWFBM2SH")) = False Then .HWFBM2SH = rs("HWFBM2SH") Else .HWFBM2SH = " "              '測定位置_方
        If IsNull(rs("HWFBM2ST")) = False Then .HWFBM2ST = rs("HWFBM2ST") Else .HWFBM2ST = " "              '測定位置_点
        If IsNull(rs("HWFBM2SR")) = False Then .HWFBM2SR = rs("HWFBM2SR") Else .HWFBM2SR = " "              '測定位置_領
        If IsNull(rs("HWFBM2NS")) = False Then .HWFBM2NS = rs("HWFBM2NS") Else .HWFBM2NS = " "              '熱処理法
        If IsNull(rs("HWFBM2SZ")) = False Then .HWFBM2SZ = rs("HWFBM2SZ") Else .HWFBM2SZ = " "              '測定条件
        If IsNull(rs("HWFBM2ET")) = False Then .HWFBM2ET = rs("HWFBM2ET") Else .HWFBM2ET = 0                '選択ET代
        If IsNull(rs("HWFBM2KN")) = False Then .HWFBM2KN = rs("HWFBM2KN") Else .HWFBM2KN = " "              '検査頻度_抜
        'BMD3
        If IsNull(rs("HWFBM3HS")) = False Then .HWFBM3HS = rs("HWFBM3HS") Else .HWFBM3HS = " "              '保証方法_対象
        If IsNull(rs("HWFBM3SH")) = False Then .HWFBM3SH = rs("HWFBM3SH") Else .HWFBM3SH = " "              '測定位置_方
        If IsNull(rs("HWFBM3ST")) = False Then .HWFBM3ST = rs("HWFBM3ST") Else .HWFBM3ST = " "              '測定位置_点
        If IsNull(rs("HWFBM3SR")) = False Then .HWFBM3SR = rs("HWFBM3SR") Else .HWFBM3SR = " "              '測定位置_領
        If IsNull(rs("HWFBM3NS")) = False Then .HWFBM3NS = rs("HWFBM3NS") Else .HWFBM3NS = " "              '熱処理法
        If IsNull(rs("HWFBM3SZ")) = False Then .HWFBM3SZ = rs("HWFBM3SZ") Else .HWFBM3SZ = " "              '測定条件
        If IsNull(rs("HWFBM3ET")) = False Then .HWFBM3ET = rs("HWFBM3ET") Else .HWFBM3ET = 0                '選択ET代
        If IsNull(rs("HWFBM3KN")) = False Then .HWFBM3KN = rs("HWFBM3KN") Else .HWFBM3KN = " "              '検査頻度_抜
        'DOI1
        If IsNull(rs("HWFOS1HS")) = False Then .HWFOS1HS = rs("HWFOS1HS") Else .HWFOS1HS = " "              '保証方法_対象
        If IsNull(rs("HWFOS1NS")) = False Then .HWFOS1NS = rs("HWFOS1NS") Else .HWFOS1NS = " "              '熱処理法
        If IsNull(rs("HWFOS1KN")) = False Then .HWFOS1KN = rs("HWFOS1KN") Else .HWFOS1KN = " "              '検査頻度_抜
        'DOI2
        If IsNull(rs("HWFOS2HS")) = False Then .HWFOS2HS = rs("HWFOS2HS") Else .HWFOS2HS = " "              '保証方法_対象
        If IsNull(rs("HWFOS2NS")) = False Then .HWFOS2NS = rs("HWFOS2NS") Else .HWFOS2NS = " "              '熱処理法
        If IsNull(rs("HWFOS2KN")) = False Then .HWFOS2KN = rs("HWFOS2KN") Else .HWFOS2KN = " "              '検査頻度_抜
        'DOI3
        If IsNull(rs("HWFOS3HS")) = False Then .HWFOS3HS = rs("HWFOS3HS") Else .HWFOS3HS = " "              '保証方法_対象
        If IsNull(rs("HWFOS3NS")) = False Then .HWFOS3NS = rs("HWFOS3NS") Else .HWFOS3NS = " "              '熱処理法
        If IsNull(rs("HWFOS3KN")) = False Then .HWFOS3KN = rs("HWFOS3KN") Else .HWFOS3KN = " "              '検査頻度_抜
        'DSOD
        If IsNull(rs("HWFDSOHS")) = False Then .HWFDSOHS = rs("HWFDSOHS") Else .HWFDSOHS = " "              '保証方法_対象
        If IsNull(rs("HWFDSONWY")) = False Then .HWFDSONWY = rs("HWFDSONWY") Else .HWFDSONWY = " "          '熱処理法
        If IsNull(rs("HWFDSOKN")) = False Then .HWFDSOKN = rs("HWFDSOKN") Else .HWFDSOKN = " "              '検査頻度_抜
        If IsNull(rs("HWFDSOPTK")) = False Then .HWFDSOPTK = rs("HWFDSOPTK") Else .HWFDSOPTK = " "          'パターン区分　04/07/29 ooba
        'DZ
        If IsNull(rs("HWFMKHWS")) = False Then .HWFMKHWS = rs("HWFMKHWS") Else .HWFMKHWS = " "              '保証方法_対象
        If IsNull(rs("HWFMKSPH")) = False Then .HWFMKSPH = rs("HWFMKSPH") Else .HWFMKSPH = " "              '測定位置_方
        If IsNull(rs("HWFMKSPT")) = False Then .HWFMKSPT = rs("HWFMKSPT") Else .HWFMKSPT = " "              '測定位置_点
        If IsNull(rs("HWFMKSPR")) = False Then .HWFMKSPR = rs("HWFMKSPR") Else .HWFMKSPR = " "              '測定位置_領
        If IsNull(rs("HWFMKNSW")) = False Then .HWFMKNSW = rs("HWFMKNSW") Else .HWFMKNSW = " "              '熱処理法
        If IsNull(rs("HWFMKSZY")) = False Then .HWFMKSZY = rs("HWFMKSZY") Else .HWFMKSZY = " "              '測定条件
        If IsNull(rs("HWFMKCET")) = False Then .HWFMKCET = rs("HWFMKCET") Else .HWFMKCET = 0                '選択ET代
        If IsNull(rs("HWFMKKHN")) = False Then .HWFMKKHN = rs("HWFMKKHN") Else .HWFMKKHN = " "              '検査頻度_抜
        'SPVFE
        If IsNull(rs("HWFSPVHS")) = False Then .HWFSPVHS = rs("HWFSPVHS") Else .HWFSPVHS = " "              '保証方法_対象
        If IsNull(rs("HWFSPVST")) = False Then .HWFSPVST = rs("HWFSPVST") Else .HWFSPVST = " "              '測定位置＿点
        If IsNull(rs("HWFSPVKN")) = False Then .HWFSPVKN = rs("HWFSPVKN") Else .HWFSPVKN = " "              '検査頻度_抜
        '拡散長
        If IsNull(rs("HWFDLHWS")) = False Then .HWFDLHWS = rs("HWFDLHWS") Else .HWFDLHWS = " "              '保証方法_対象
        If IsNull(rs("HWFDLKHN")) = False Then .HWFDLKHN = rs("HWFDLKHN") Else .HWFDLKHN = " "              '検査頻度_抜
        
    ''Upd Start 2005/06/16 (TCS)T.Terauchi  SPV9点対応
        'SPVFE
        If IsNull(rs("HWFSPVSH")) = False Then .HWFSPVSH = rs("HWFSPVSH") Else .HWFSPVSH = " "              '測定位置＿方
        If IsNull(rs("HWFSPVSI")) = False Then .HWFSPVSI = rs("HWFSPVSI") Else .HWFSPVSI = " "              '測定位置＿位
        '拡散長
        If IsNull(rs("HWFDLSPH")) = False Then .HWFDLSPH = rs("HWFDLSPH") Else .HWFDLSPH = " "              '測定位置＿方
        If IsNull(rs("HWFDLSPT")) = False Then .HWFDLSPT = rs("HWFDLSPT") Else .HWFDLSPT = " "              '測定位置＿点
        If IsNull(rs("HWFDLSPI")) = False Then .HWFDLSPI = rs("HWFDLSPI") Else .HWFDLSPI = " "              '測定位置＿位
    ''Upd End   2005/06/16 (TCS)T.Terauchi  SPV9点対応
        
        ''06/05/31 ooba START ==================================================================>
        'SPVFE
        If IsNull(rs("HWFSPVPUG")) = False Then .HWFSPVPUG = rs("HWFSPVPUG") Else .HWFSPVPUG = -1           'PUA限
        If IsNull(rs("HWFSPVPUR")) = False Then .HWFSPVPUR = rs("HWFSPVPUR") Else .HWFSPVPUR = -1           'PUA率
        '拡散長
        If IsNull(rs("HWFDLPUG")) = False Then .HWFDLPUG = rs("HWFDLPUG") Else .HWFDLPUG = -1               'PUA限
        If IsNull(rs("HWFDLPUR")) = False Then .HWFDLPUR = rs("HWFDLPUR") Else .HWFDLPUR = -1               'PUA率
        'SPVNR
        If IsNull(rs("HWFNRHS")) = False Then .HWFNRHS = rs("HWFNRHS") Else .HWFNRHS = " "                  '保証方法＿対象
        If IsNull(rs("HWFNRSH")) = False Then .HWFNRSH = rs("HWFNRSH") Else .HWFNRSH = " "                  '測定位置＿方
        If IsNull(rs("HWFNRST")) = False Then .HWFNRST = rs("HWFNRST") Else .HWFNRST = " "                  '測定位置＿点
        If IsNull(rs("HWFNRSI")) = False Then .HWFNRSI = rs("HWFNRSI") Else .HWFNRSI = " "                  '測定位置＿位
        If IsNull(rs("HWFNRKN")) = False Then .HWFNRKN = rs("HWFNRKN") Else .HWFNRKN = " "                  '検査頻度＿抜
        If IsNull(rs("HWFNRPUG")) = False Then .HWFNRPUG = rs("HWFNRPUG") Else .HWFNRPUG = -1               'PUA限
        If IsNull(rs("HWFNRPUR")) = False Then .HWFNRPUR = rs("HWFNRPUR") Else .HWFNRPUR = -1               'PUA率
        ''06/05/31 ooba END ====================================================================>
        
        'AOi        残存酸素追加　03/12/09 ooba
        If IsNull(rs("HWFZOHWS")) = False Then .HWFZOHWS = rs("HWFZOHWS") Else .HWFZOHWS = " "              '保証方法_対象
        If IsNull(rs("HWFZONSW")) = False Then .HWFZONSW = rs("HWFZONSW") Else .HWFZONSW = " "              '熱処理法
        If IsNull(rs("HWFZOKHN")) = False Then .HWFZOKHN = rs("HWFZOKHN") Else .HWFZOKHN = " "              '検査頻度_抜
        'DEN        'DEN追加　05/01/27 ooba
        If IsNull(rs("HWFDENHS")) = False Then .HWFDENHS = rs("HWFDENHS") Else .HWFDENHS = " "              '保証方法_対象
        If IsNull(rs("HWFDENMN")) = False Then .HWFDENMN = rs("HWFDENMN") Else .HWFDENMN = 0                '下限
        If IsNull(rs("HWFDENMX")) = False Then .HWFDENMX = rs("HWFDENMX") Else .HWFDENMX = 0                '上限
        'DVD2       'DVD2追加　05/01/27 ooba
        If IsNull(rs("HWFDVDHS")) = False Then .HWFDVDHS = rs("HWFDVDHS") Else .HWFDVDHS = " "              '保証方法_対象
        If IsNull(rs("HWFDVDMNN")) = False Then .HWFDVDMNN = rs("HWFDVDMNN") Else .HWFDVDMNN = 0            '下限
        If IsNull(rs("HWFDVDMXN")) = False Then .HWFDVDMXN = rs("HWFDVDMXN") Else .HWFDVDMXN = 0            '上限
        'L/DL       'L/DL追加　05/01/27 ooba
        If IsNull(rs("HWFLDLHS")) = False Then .HWFLDLHS = rs("HWFLDLHS") Else .HWFLDLHS = " "              '保証方法_対象
        If IsNull(rs("HWFLDLMN")) = False Then .HWFLDLMN = rs("HWFLDLMN") Else .HWFLDLMN = 0                '下限
        If IsNull(rs("HWFLDLMX")) = False Then .HWFLDLMX = rs("HWFLDLMX") Else .HWFLDLMX = 0                '上限
        If IsNull(rs("HWFGDKHN")) = False Then .HWFGDKHN = rs("HWFGDKHN") Else .HWFGDKHN = " "              '検査頻度_抜
    '*** UPDATE ↓ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数追加
        If IsNull(rs("HWFGDLINE")) = False Then .HWFGDLINE = rs("HWFGDLINE") Else .HWFGDLINE = " "               '測定条件
    '*** UPDATE ↑ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数追加
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
        If IsNull(rs("HWFGDSZY")) = False Then .HWFGDSZY = rs("HWFGDSZY") Else .HWFGDSZY = " "               '測定条件
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
    '↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    '2.1.1 AN温度 振替可否チェック追加
        If IsNull(rs("HWFANTNP")) = False Then .HWFANTNP = rs("HWFANTNP") Else .HWFANTNP = 0                '品ＷＦＡＮ温度
    '↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    End With
    
    Set rs = Nothing
    
    ''ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞ｢1｣→｢2｣の振替時は、(元)結晶GDと(先)WF_GDのﾁｪｯｸとする。　05/07/29 ooba
    If tbl_chk1_5(0).BLOCKHFLAG = "1" And tbl_chk1_5(1).BLOCKHFLAG = "2" Then
        'DEN
        tbl_chk1_5(0).HWFDENHS = tbl_chk1_5_SXGD.HWFDENHS               '保証方法_対象
        tbl_chk1_5(0).HWFDENMN = tbl_chk1_5_SXGD.HWFDENMN               '下限
        tbl_chk1_5(0).HWFDENMX = tbl_chk1_5_SXGD.HWFDENMX               '上限
        'DVD2
        tbl_chk1_5(0).HWFDVDHS = tbl_chk1_5_SXGD.HWFDVDHS               '保証方法_対象
        tbl_chk1_5(0).HWFDVDMNN = tbl_chk1_5_SXGD.HWFDVDMNN             '下限
        tbl_chk1_5(0).HWFDVDMXN = tbl_chk1_5_SXGD.HWFDVDMXN             '上限
        'L/DL
        tbl_chk1_5(0).HWFLDLHS = tbl_chk1_5_SXGD.HWFLDLHS               '保証方法_対象
        tbl_chk1_5(0).HWFLDLMN = tbl_chk1_5_SXGD.HWFLDLMN               '下限
        tbl_chk1_5(0).HWFLDLMX = tbl_chk1_5_SXGD.HWFLDLMX               '上限
        'GD抜取位置 (結晶はT/B保証とする)
        tbl_chk1_5(0).HWFGDKHN = "6"                                    '検査頻度_抜
        'GDライン数
        tbl_chk1_5(0).HWFGDLINE = tbl_chk1_5_SXGD.HWFGDLINE             'ライン数
    End If
    
    '------------------------------------------ 指示取得 ------------------------------------------------------
    On Error GoTo Apl_down
    '比抵抗
    sErr_Msg = "1-5 比抵抗ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "RS", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFRHWYS
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFRHWYS
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFRKHNN          '04/04/13 ooba
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFRKHNN          '04/04/13 ooba
'↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.1 AN温度 振替可否チェック追加
    '振替チェックにAN温度を加える
    tbl_chk1_5_1(0).HWFANTNP = tbl_chk1_5(0).HWFANTNP
    tbl_chk1_5_1(1).HWFANTNP = tbl_chk1_5(1).HWFANTNP
'↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,RS")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00060"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    '酸素濃度
    sErr_Msg = "1-5 酸素濃度ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "OI", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFONHWS
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFONHWS
    tbl_chk1_5_1(0).SOKU_TEN = tbl_chk1_5(0).HWFONSPT           '08/01/29 ooba
    tbl_chk1_5_1(1).SOKU_TEN = tbl_chk1_5(1).HWFONSPT           '08/01/29 ooba
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFONKHN          '04/04/13 ooba
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFONKHN          '04/04/13 ooba
'↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.1 AN温度 振替可否チェック追加
    '振替チェックにAN温度を加える
    tbl_chk1_5_1(0).HWFANTNP = tbl_chk1_5(0).HWFANTNP
    tbl_chk1_5_1(1).HWFANTNP = tbl_chk1_5(1).HWFANTNP
'↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,Oi")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00061"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'ＯＳＦ1
    sErr_Msg = "1-5 OSF1ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "O1", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFOF1HS
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFOF1HS
    tbl_chk1_5_1(0).SOKU_HOU = tbl_chk1_5(0).HWFOF1SH
    tbl_chk1_5_1(1).SOKU_HOU = tbl_chk1_5(1).HWFOF1SH
    tbl_chk1_5_1(0).SOKU_RYOU = tbl_chk1_5(0).HWFOF1SR
    tbl_chk1_5_1(1).SOKU_RYOU = tbl_chk1_5(1).HWFOF1SR
    tbl_chk1_5_1(0).NETSU = tbl_chk1_5(0).HWFOF1NS
    tbl_chk1_5_1(1).NETSU = tbl_chk1_5(1).HWFOF1NS
    tbl_chk1_5_1(0).JOUKEN = tbl_chk1_5(0).HWFOF1SZ
    tbl_chk1_5_1(1).JOUKEN = tbl_chk1_5(1).HWFOF1SZ
    tbl_chk1_5_1(0).ET = tbl_chk1_5(0).HWFOF1ET
    tbl_chk1_5_1(1).ET = tbl_chk1_5(1).HWFOF1ET
    tbl_chk1_5_1(0).PATTERN = tbl_chk1_5(0).HWFOSF1PTK
    tbl_chk1_5_1(1).PATTERN = tbl_chk1_5(1).HWFOSF1PTK
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFOF1KN          '04/04/13 ooba
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFOF1KN          '04/04/13 ooba
'↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.1 AN温度 振替可否チェック追加
    '振替チェックにAN温度を加える
    tbl_chk1_5_1(0).HWFANTNP = tbl_chk1_5(0).HWFANTNP
    tbl_chk1_5_1(1).HWFANTNP = tbl_chk1_5(1).HWFANTNP
'↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,OSF1")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00062"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'ＯＳＦ２
    sErr_Msg = "1-5 OSF2ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "O2", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFOF2HS
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFOF2HS
    tbl_chk1_5_1(0).SOKU_HOU = tbl_chk1_5(0).HWFOF2SH
    tbl_chk1_5_1(1).SOKU_HOU = tbl_chk1_5(1).HWFOF2SH
    tbl_chk1_5_1(0).SOKU_RYOU = tbl_chk1_5(0).HWFOF2SR
    tbl_chk1_5_1(1).SOKU_RYOU = tbl_chk1_5(1).HWFOF2SR
    tbl_chk1_5_1(0).NETSU = tbl_chk1_5(0).HWFOF2NS
    tbl_chk1_5_1(1).NETSU = tbl_chk1_5(1).HWFOF2NS
    tbl_chk1_5_1(0).JOUKEN = tbl_chk1_5(0).HWFOF2SZ
    tbl_chk1_5_1(1).JOUKEN = tbl_chk1_5(1).HWFOF2SZ
    tbl_chk1_5_1(0).ET = tbl_chk1_5(0).HWFOF2ET
    tbl_chk1_5_1(1).ET = tbl_chk1_5(1).HWFOF2ET
    tbl_chk1_5_1(0).PATTERN = tbl_chk1_5(0).HWFOSF2PTK
    tbl_chk1_5_1(1).PATTERN = tbl_chk1_5(1).HWFOSF2PTK
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFOF2KN          '04/04/13 ooba
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFOF2KN          '04/04/13 ooba
'↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.1 AN温度 振替可否チェック追加
    '振替チェックにAN温度を加える
    tbl_chk1_5_1(0).HWFANTNP = tbl_chk1_5(0).HWFANTNP
    tbl_chk1_5_1(1).HWFANTNP = tbl_chk1_5(1).HWFANTNP
'↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,OSF2")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00063"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'ＯＳＦ３
    sErr_Msg = "1-5 OSF3ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "O3", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFOF3HS
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFOF3HS
    tbl_chk1_5_1(0).SOKU_HOU = tbl_chk1_5(0).HWFOF3SH
    tbl_chk1_5_1(1).SOKU_HOU = tbl_chk1_5(1).HWFOF3SH
    tbl_chk1_5_1(0).SOKU_RYOU = tbl_chk1_5(0).HWFOF3SR
    tbl_chk1_5_1(1).SOKU_RYOU = tbl_chk1_5(1).HWFOF3SR
    tbl_chk1_5_1(0).NETSU = tbl_chk1_5(0).HWFOF3NS
    tbl_chk1_5_1(1).NETSU = tbl_chk1_5(1).HWFOF3NS
    tbl_chk1_5_1(0).JOUKEN = tbl_chk1_5(0).HWFOF3SZ
    tbl_chk1_5_1(1).JOUKEN = tbl_chk1_5(1).HWFOF3SZ
    tbl_chk1_5_1(0).ET = tbl_chk1_5(0).HWFOF3ET
    tbl_chk1_5_1(1).ET = tbl_chk1_5(1).HWFOF3ET
    tbl_chk1_5_1(0).PATTERN = tbl_chk1_5(0).HWFOSF3PTK
    tbl_chk1_5_1(1).PATTERN = tbl_chk1_5(1).HWFOSF3PTK
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFOF3KN          '04/04/13 ooba
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFOF3KN          '04/04/13 ooba
'↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.1 AN温度 振替可否チェック追加
    '振替チェックにAN温度を加える
    tbl_chk1_5_1(0).HWFANTNP = tbl_chk1_5(0).HWFANTNP
    tbl_chk1_5_1(1).HWFANTNP = tbl_chk1_5(1).HWFANTNP
'↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,OSF3")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00064"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''    'ＯＳＦ４
'''    sErr_Msg = "1-5 OSF4ﾁｪｯｸ"
'''    sResult = ""
'''    RET = funCodeDBGet("SB", "15", "O4", 0, " ", sResult)
'''    If RET <> 0 Then
'''        sErr_Msg = sErr_Msg & "→指示取得"
'''        GoTo CodeDBGet_Error
'''    End If
'''    Erase tbl_chk1_5_1
'''    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFOF4HS
'''    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFOF4HS
'''    tbl_chk1_5_1(0).SOKU_HOU = tbl_chk1_5(0).HWFOF4SH
'''    tbl_chk1_5_1(1).SOKU_HOU = tbl_chk1_5(1).HWFOF4SH
'''    tbl_chk1_5_1(0).SOKU_RYOU = tbl_chk1_5(0).HWFOF4SR
'''    tbl_chk1_5_1(1).SOKU_RYOU = tbl_chk1_5(1).HWFOF4SR
'''    tbl_chk1_5_1(0).NETSU = tbl_chk1_5(0).HWFOF4NS
'''    tbl_chk1_5_1(1).NETSU = tbl_chk1_5(1).HWFOF4NS
'''    tbl_chk1_5_1(0).JOUKEN = tbl_chk1_5(0).HWFOF4SZ
'''    tbl_chk1_5_1(1).JOUKEN = tbl_chk1_5(1).HWFOF4SZ
'''    tbl_chk1_5_1(0).ET = tbl_chk1_5(0).HWFOF4ET
'''    tbl_chk1_5_1(1).ET = tbl_chk1_5(1).HWFOF4ET
'''    tbl_chk1_5_1(0).PATTERN = tbl_chk1_5(0).HWFOSF4PTK
'''    tbl_chk1_5_1(1).PATTERN = tbl_chk1_5(1).HWFOSF4PTK
'''    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFOF4KN          '04/04/13 ooba
'''    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFOF4KN          '04/04/13 ooba
''''↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
''''2.1.1 AN温度 振替可否チェック追加
'''    '振替チェックにAN温度を加える
'''    tbl_chk1_5_1(0).HWFANTNP = tbl_chk1_5(0).HWFANTNP
'''    tbl_chk1_5_1(1).HWFANTNP = tbl_chk1_5(1).HWFANTNP
''''↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'''    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,OSF4")
'''    If RET <> 0 Then
'''        funChkFurikae1_5 = RET
''''--------------- 2008/07/25 INSERT START  By Systech ---------------
'''        gsTbcmy028ErrCode = "00065"
''''--------------- 2008/07/25 INSERT  END   By Systech ---------------
'''        GoTo Apl_Exit
'''    End If

    'ＳＩＲＤ
    sErr_Msg = "1-5 SIRDﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "SD", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1                                        '引数ﾃｰﾌﾞﾙｸﾘｱ
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFSIRDHS          '軸状転位保証方法＿処
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFSIRDHS          '軸状転位保証方法＿処
    tbl_chk1_5_1(0).JOUKEN = tbl_chk1_5(0).HWFSIRDSZ          '軸状転位測定条件
    tbl_chk1_5_1(1).JOUKEN = tbl_chk1_5(1).HWFSIRDSZ          '軸状転位測定条件
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFSIRDKN       '軸状転位検査頻度＿抜
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFSIRDKN       '軸状転位検査頻度＿抜
    tbl_chk1_5_1(0).HWFSIRDMX = tbl_chk1_5(0).HWFSIRDMX       '軸状転位上限
    tbl_chk1_5_1(1).HWFSIRDMX = tbl_chk1_5(1).HWFSIRDMX       '軸状転位上限
    tbl_chk1_5_1(0).HWFSIRDHT = tbl_chk1_5(0).HWFSIRDHT       '軸状転位保証方法＿対
    tbl_chk1_5_1(1).HWFSIRDHT = tbl_chk1_5(1).HWFSIRDHT       '軸状転位保証方法＿対
    tbl_chk1_5_1(0).HWFSIRDKM = tbl_chk1_5(0).HWFSIRDKM       '軸状転位検査頻度＿枚
    tbl_chk1_5_1(1).HWFSIRDKM = tbl_chk1_5(1).HWFSIRDKM       '軸状転位検査頻度＿枚
    tbl_chk1_5_1(0).HWFSIRDKH = tbl_chk1_5(0).HWFSIRDKH       '軸状転位検査頻度＿保
    tbl_chk1_5_1(1).HWFSIRDKH = tbl_chk1_5(1).HWFSIRDKH       '軸状転位検査頻度＿保
    tbl_chk1_5_1(0).HWFSIRDKU = tbl_chk1_5(0).HWFSIRDKU       '軸状転位検査頻度＿ウ
    tbl_chk1_5_1(1).HWFSIRDKU = tbl_chk1_5(1).HWFSIRDKU       '軸状転位検査頻度＿ウ
    tbl_chk1_5_1(0).HWFANTNP = tbl_chk1_5(0).HWFANTNP         '2.1.1 AN温度 振替可否チェック
    tbl_chk1_5_1(1).HWFANTNP = tbl_chk1_5(1).HWFANTNP         '2.1.1 AN温度 振替可否チェック
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,SIRD")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
        gsTbcmy028ErrCode = "00065"
        GoTo Apl_Exit
    End If
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
    
    'ＢＭＤ１
    sErr_Msg = "1-5 BMD1ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "B1", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFBM1HS
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFBM1HS
    tbl_chk1_5_1(0).SOKU_HOU = tbl_chk1_5(0).HWFBM1SH
    tbl_chk1_5_1(1).SOKU_HOU = tbl_chk1_5(1).HWFBM1SH
    tbl_chk1_5_1(0).SOKU_TEN = tbl_chk1_5(0).HWFBM1ST
    tbl_chk1_5_1(1).SOKU_TEN = tbl_chk1_5(1).HWFBM1ST
    tbl_chk1_5_1(0).SOKU_RYOU = tbl_chk1_5(0).HWFBM1SR
    tbl_chk1_5_1(1).SOKU_RYOU = tbl_chk1_5(1).HWFBM1SR
    tbl_chk1_5_1(0).NETSU = tbl_chk1_5(0).HWFBM1NS
    tbl_chk1_5_1(1).NETSU = tbl_chk1_5(1).HWFBM1NS
    tbl_chk1_5_1(0).JOUKEN = tbl_chk1_5(0).HWFBM1SZ
    tbl_chk1_5_1(1).JOUKEN = tbl_chk1_5(1).HWFBM1SZ
    tbl_chk1_5_1(0).ET = tbl_chk1_5(0).HWFBM1ET
    tbl_chk1_5_1(1).ET = tbl_chk1_5(1).HWFBM1ET
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFBM1KN          '04/04/13 ooba
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFBM1KN          '04/04/13 ooba
'↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.1 AN温度 振替可否チェック追加
    '振替チェックにAN温度を加える
    tbl_chk1_5_1(0).HWFANTNP = tbl_chk1_5(0).HWFANTNP
    tbl_chk1_5_1(1).HWFANTNP = tbl_chk1_5(1).HWFANTNP
'↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,BMD1")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00066"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'ＢＭＤ２
    sErr_Msg = "1-5 BMD2ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "B2", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFBM2HS
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFBM2HS
    tbl_chk1_5_1(0).SOKU_HOU = tbl_chk1_5(0).HWFBM2SH
    tbl_chk1_5_1(1).SOKU_HOU = tbl_chk1_5(1).HWFBM2SH
    tbl_chk1_5_1(0).SOKU_TEN = tbl_chk1_5(0).HWFBM2ST
    tbl_chk1_5_1(1).SOKU_TEN = tbl_chk1_5(1).HWFBM2ST
    tbl_chk1_5_1(0).SOKU_RYOU = tbl_chk1_5(0).HWFBM2SR
    tbl_chk1_5_1(1).SOKU_RYOU = tbl_chk1_5(1).HWFBM2SR
    tbl_chk1_5_1(0).NETSU = tbl_chk1_5(0).HWFBM2NS
    tbl_chk1_5_1(1).NETSU = tbl_chk1_5(1).HWFBM2NS
    tbl_chk1_5_1(0).JOUKEN = tbl_chk1_5(0).HWFBM2SZ
    tbl_chk1_5_1(1).JOUKEN = tbl_chk1_5(1).HWFBM2SZ
    tbl_chk1_5_1(0).ET = tbl_chk1_5(0).HWFBM2ET
    tbl_chk1_5_1(1).ET = tbl_chk1_5(1).HWFBM2ET
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFBM2KN          '04/04/13 ooba
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFBM2KN          '04/04/13 ooba
'↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.1 AN温度 振替可否チェック追加
    '振替チェックにAN温度を加える
    tbl_chk1_5_1(0).HWFANTNP = tbl_chk1_5(0).HWFANTNP
    tbl_chk1_5_1(1).HWFANTNP = tbl_chk1_5(1).HWFANTNP
'↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,BMD2")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00067"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'ＢＭＤ３
    sErr_Msg = "1-5 BMD3ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "B3", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFBM3HS
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFBM3HS
    tbl_chk1_5_1(0).SOKU_HOU = tbl_chk1_5(0).HWFBM3SH
    tbl_chk1_5_1(1).SOKU_HOU = tbl_chk1_5(1).HWFBM3SH
    tbl_chk1_5_1(0).SOKU_TEN = tbl_chk1_5(0).HWFBM3ST
    tbl_chk1_5_1(1).SOKU_TEN = tbl_chk1_5(1).HWFBM3ST
    tbl_chk1_5_1(0).SOKU_RYOU = tbl_chk1_5(0).HWFBM3SR
    tbl_chk1_5_1(1).SOKU_RYOU = tbl_chk1_5(1).HWFBM3SR
    tbl_chk1_5_1(0).NETSU = tbl_chk1_5(0).HWFBM3NS
    tbl_chk1_5_1(1).NETSU = tbl_chk1_5(1).HWFBM3NS
    tbl_chk1_5_1(0).JOUKEN = tbl_chk1_5(0).HWFBM3SZ
    tbl_chk1_5_1(1).JOUKEN = tbl_chk1_5(1).HWFBM3SZ
    tbl_chk1_5_1(0).ET = tbl_chk1_5(0).HWFBM3ET
    tbl_chk1_5_1(1).ET = tbl_chk1_5(1).HWFBM3ET
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFBM3KN          '04/04/13 ooba
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFBM3KN          '04/04/13 ooba
'↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.1 AN温度 振替可否チェック追加
    '振替チェックにAN温度を加える
    tbl_chk1_5_1(0).HWFANTNP = tbl_chk1_5(0).HWFANTNP
    tbl_chk1_5_1(1).HWFANTNP = tbl_chk1_5(1).HWFANTNP
'↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,BMD3")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00068"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    '酸素析出１
    sErr_Msg = "1-5 酸素析出1ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "D1", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFOS1HS
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFOS1HS
    tbl_chk1_5_1(0).NETSU = tbl_chk1_5(0).HWFOS1NS
    tbl_chk1_5_1(1).NETSU = tbl_chk1_5(1).HWFOS1NS
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFOS1KN          '04/04/13 ooba
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFOS1KN          '04/04/13 ooba
'↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.1 AN温度 振替可否チェック追加
    '振替チェックにAN温度を加える
    tbl_chk1_5_1(0).HWFANTNP = tbl_chk1_5(0).HWFANTNP
    tbl_chk1_5_1(1).HWFANTNP = tbl_chk1_5(1).HWFANTNP
'↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,DO1")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00069"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    '酸素析出２
    sErr_Msg = "1-5 酸素析出2ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "D2", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFOS2HS
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFOS2HS
    tbl_chk1_5_1(0).NETSU = tbl_chk1_5(0).HWFOS2NS
    tbl_chk1_5_1(1).NETSU = tbl_chk1_5(1).HWFOS2NS
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFOS2KN          '04/04/13 ooba
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFOS2KN          '04/04/13 ooba
'↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.1 AN温度 振替可否チェック追加
    '振替チェックにAN温度を加える
    tbl_chk1_5_1(0).HWFANTNP = tbl_chk1_5(0).HWFANTNP
    tbl_chk1_5_1(1).HWFANTNP = tbl_chk1_5(1).HWFANTNP
'↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,DO2")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00070"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    '酸素析出３
    sErr_Msg = "1-5 酸素析出3ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "D3", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFOS3HS
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFOS3HS
    tbl_chk1_5_1(0).NETSU = tbl_chk1_5(0).HWFOS3NS
    tbl_chk1_5_1(1).NETSU = tbl_chk1_5(1).HWFOS3NS
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFOS3KN          '04/04/13 ooba
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFOS3KN          '04/04/13 ooba
'↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.1 AN温度 振替可否チェック追加
    '振替チェックにAN温度を加える
    tbl_chk1_5_1(0).HWFANTNP = tbl_chk1_5(0).HWFANTNP
    tbl_chk1_5_1(1).HWFANTNP = tbl_chk1_5(1).HWFANTNP
'↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,DO3")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00071"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'ＤＳＯＤ
    sErr_Msg = "1-5 DSﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "DS", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFDSOHS
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFDSOHS
    tbl_chk1_5_1(0).NETSU = tbl_chk1_5(0).HWFDSONWY
    tbl_chk1_5_1(1).NETSU = tbl_chk1_5(1).HWFDSONWY
    tbl_chk1_5_1(0).PATTERN = tbl_chk1_5(0).HWFDSOPTK           'ﾊﾟﾀｰﾝ区分追加　04/07/29 ooba
    tbl_chk1_5_1(1).PATTERN = tbl_chk1_5(1).HWFDSOPTK           'ﾊﾟﾀｰﾝ区分追加　04/07/29 ooba
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFDSOKN          '04/04/13 ooba
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFDSOKN          '04/04/13 ooba
    'GD/DSOD熱処理条件追加　06/12/22 ooba START =========================================>
    tbl_chk1_5_1(0).HWFANTNP = tbl_chk1_5(0).HWFANTNP
    tbl_chk1_5_1(1).HWFANTNP = tbl_chk1_5(1).HWFANTNP
    'GD/DSOD熱処理条件追加　06/12/22 ooba END ===========================================>
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,DS")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00073"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'ＤＺ
    sErr_Msg = "1-5 DZﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "DZ", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFMKHWS
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFMKHWS
    tbl_chk1_5_1(0).SOKU_HOU = tbl_chk1_5(0).HWFMKSPH
    tbl_chk1_5_1(1).SOKU_HOU = tbl_chk1_5(1).HWFMKSPH
    tbl_chk1_5_1(0).SOKU_TEN = tbl_chk1_5(0).HWFMKSPT
    tbl_chk1_5_1(1).SOKU_TEN = tbl_chk1_5(1).HWFMKSPT
    tbl_chk1_5_1(0).SOKU_RYOU = tbl_chk1_5(0).HWFMKSPR
    tbl_chk1_5_1(1).SOKU_RYOU = tbl_chk1_5(1).HWFMKSPR
    tbl_chk1_5_1(0).NETSU = tbl_chk1_5(0).HWFMKNSW
    tbl_chk1_5_1(1).NETSU = tbl_chk1_5(1).HWFMKNSW
    tbl_chk1_5_1(0).JOUKEN = tbl_chk1_5(0).HWFMKSZY
    tbl_chk1_5_1(1).JOUKEN = tbl_chk1_5(1).HWFMKSZY
    tbl_chk1_5_1(0).ET = tbl_chk1_5(0).HWFMKCET
    tbl_chk1_5_1(1).ET = tbl_chk1_5(1).HWFMKCET
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFMKKHN          '04/04/13 ooba
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFMKKHN          '04/04/13 ooba
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,DZ")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00074"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'ＳＰＶＦＥ
    sErr_Msg = "1-5 SPVFEﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "SP", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFSPVHS
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFSPVHS
    tbl_chk1_5_1(0).SOKU_TEN = tbl_chk1_5(0).HWFSPVST
    tbl_chk1_5_1(1).SOKU_TEN = tbl_chk1_5(1).HWFSPVST
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFSPVKN          '04/04/13 ooba
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFSPVKN          '04/04/13 ooba
    
''Upd Start 2005/06/16 (TCS)T.Terauchi      SPV9点対応
    tbl_chk1_5_1(0).SOKU_HOU = tbl_chk1_5(0).HWFSPVSH       ''測定位置＿方(振替元)
    tbl_chk1_5_1(1).SOKU_HOU = tbl_chk1_5(1).HWFSPVSH       ''測定位置＿方(振替先)
    tbl_chk1_5_1(0).SOKU_ICHI = tbl_chk1_5(0).HWFSPVSI      ''測定位置＿位(振替元)
    tbl_chk1_5_1(1).SOKU_ICHI = tbl_chk1_5(1).HWFSPVSI      ''測定位置＿位(振替先)
''Upd End   2005/06/16 (TCS)T.Terauchi      SPV9点対応
    
    ''06/05/31 ooba START ============================================>
    tbl_chk1_5_1(0).PUAGEN = tbl_chk1_5(0).HWFSPVPUG
    tbl_chk1_5_1(1).PUAGEN = tbl_chk1_5(1).HWFSPVPUG
    tbl_chk1_5_1(0).PUAPER = tbl_chk1_5(0).HWFSPVPUR
    tbl_chk1_5_1(1).PUAPER = tbl_chk1_5(1).HWFSPVPUR
    ''06/05/31 ooba END ==============================================>
    
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,SPVFE")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00075"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    '拡散長
    sErr_Msg = "1-5 拡散長ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "KL", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFDLHWS
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFDLHWS
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFDLKHN          '04/04/13 ooba
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFDLKHN          '04/04/13 ooba
    
''Upd Start 2005/06/16 (TCS)T.Terauchi      SPV9点対応
    tbl_chk1_5_1(0).SOKU_HOU = tbl_chk1_5(0).HWFDLSPH       ''測定位置＿方(振替元)
    tbl_chk1_5_1(1).SOKU_HOU = tbl_chk1_5(1).HWFDLSPH       ''測定位置＿方(振替先)
    tbl_chk1_5_1(0).SOKU_TEN = tbl_chk1_5(0).HWFDLSPT       ''測定位置＿点(振替元)
    tbl_chk1_5_1(1).SOKU_TEN = tbl_chk1_5(1).HWFDLSPT       ''測定位置＿点(振替先)
    tbl_chk1_5_1(0).SOKU_ICHI = tbl_chk1_5(0).HWFDLSPI      ''測定位置＿位(振替元)
    tbl_chk1_5_1(1).SOKU_ICHI = tbl_chk1_5(1).HWFDLSPI      ''測定位置＿位(振替先)
''Upd End   2005/06/16 (TCS)T.Terauchi      SPV9点対応
    
    ''06/05/31 ooba START ============================================>
    tbl_chk1_5_1(0).PUAGEN = tbl_chk1_5(0).HWFDLPUG
    tbl_chk1_5_1(1).PUAGEN = tbl_chk1_5(1).HWFDLPUG
    tbl_chk1_5_1(0).PUAPER = tbl_chk1_5(0).HWFDLPUR
    tbl_chk1_5_1(1).PUAPER = tbl_chk1_5(1).HWFDLPUR
    ''06/05/31 ooba END ==============================================>
    
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,拡散長")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00076"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    
    ''残存酸素追加　03/12/09 ooba START ============================================>
    '残存酸素
    sErr_Msg = "1-5 残存酸素ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "AO", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFZOHWS
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFZOHWS
    tbl_chk1_5_1(0).NETSU = tbl_chk1_5(0).HWFZONSW
    tbl_chk1_5_1(1).NETSU = tbl_chk1_5(1).HWFZONSW
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFZOKHN          '04/04/13 ooba
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFZOKHN          '04/04/13 ooba
'↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.1 AN温度 振替可否チェック追加
    '振替チェックにAN温度を加える
    tbl_chk1_5_1(0).HWFANTNP = tbl_chk1_5(0).HWFANTNP
    tbl_chk1_5_1(1).HWFANTNP = tbl_chk1_5(1).HWFANTNP
'↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,AOi")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00072"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    ''残存酸素追加　03/12/09 ooba END ==============================================>
    
    ''GD追加　05/01/27 ooba START =================================================>
    'ＤＥＮ
    sErr_Msg = "1-5 DENﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "DEN", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFDENHS
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFDENHS
    tbl_chk1_5_1(0).Min = tbl_chk1_5(0).HWFDENMN
    tbl_chk1_5_1(1).Min = tbl_chk1_5(1).HWFDENMN
    tbl_chk1_5_1(0).max = tbl_chk1_5(0).HWFDENMX
    tbl_chk1_5_1(1).max = tbl_chk1_5(1).HWFDENMX
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFGDKHN
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFGDKHN
'*** UPDATE ↓ Y.SIMIZU 2005/10/7 ﾗｲﾝ数追加
    tbl_chk1_5_1(0).LINE = tbl_chk1_5(0).HWFGDLINE
    tbl_chk1_5_1(1).LINE = tbl_chk1_5(1).HWFGDLINE
'*** UPDATE ↑ Y.SIMIZU 2005/10/7 ﾗｲﾝ数追加
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    tbl_chk1_5_1(0).HWFGDSZY = tbl_chk1_5(0).HWFGDSZY
    tbl_chk1_5_1(1).HWFGDSZY = tbl_chk1_5(1).HWFGDSZY
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
    'GD/DSOD熱処理条件追加　06/12/22 ooba START =========================================>
    tbl_chk1_5_1(0).HWFANTNP = tbl_chk1_5(0).HWFANTNP
    tbl_chk1_5_1(1).HWFANTNP = tbl_chk1_5(1).HWFANTNP
    'GD/DSOD熱処理条件追加　06/12/22 ooba END ===========================================>
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,DEN")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        If iErr_Code = 1515 Then
            gsTbcmy028ErrCode = "00078"
        Else
            gsTbcmy028ErrCode = "00079"
        End If
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'ＤＶＤ２
    sErr_Msg = "1-5 DVD2ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "DVD", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFDVDHS
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFDVDHS
    tbl_chk1_5_1(0).Min = tbl_chk1_5(0).HWFDVDMNN
    tbl_chk1_5_1(1).Min = tbl_chk1_5(1).HWFDVDMNN
    tbl_chk1_5_1(0).max = tbl_chk1_5(0).HWFDVDMXN
    tbl_chk1_5_1(1).max = tbl_chk1_5(1).HWFDVDMXN
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFGDKHN
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFGDKHN
'*** UPDATE ↓ Y.SIMIZU 2005/10/7 ﾗｲﾝ数追加
    tbl_chk1_5_1(0).LINE = tbl_chk1_5(0).HWFGDLINE
    tbl_chk1_5_1(1).LINE = tbl_chk1_5(1).HWFGDLINE
'*** UPDATE ↑ Y.SIMIZU 2005/10/7 ﾗｲﾝ数追加
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    tbl_chk1_5_1(0).HWFGDSZY = tbl_chk1_5(0).HWFGDSZY
    tbl_chk1_5_1(1).HWFGDSZY = tbl_chk1_5(1).HWFGDSZY
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
    'GD/DSOD熱処理条件追加　06/12/22 ooba START =========================================>
    tbl_chk1_5_1(0).HWFANTNP = tbl_chk1_5(0).HWFANTNP
    tbl_chk1_5_1(1).HWFANTNP = tbl_chk1_5(1).HWFANTNP
    'GD/DSOD熱処理条件追加　06/12/22 ooba END ===========================================>
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,DVD")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00080"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'Ｌ／ＤＬ
    sErr_Msg = "1-5 L/DLﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "LDL", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFLDLHS
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFLDLHS
    tbl_chk1_5_1(0).Min = tbl_chk1_5(0).HWFLDLMN
    tbl_chk1_5_1(1).Min = tbl_chk1_5(1).HWFLDLMN
    tbl_chk1_5_1(0).max = tbl_chk1_5(0).HWFLDLMX
    tbl_chk1_5_1(1).max = tbl_chk1_5(1).HWFLDLMX
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFGDKHN
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFGDKHN
'*** UPDATE ↓ Y.SIMIZU 2005/10/7 ﾗｲﾝ数追加
    tbl_chk1_5_1(0).LINE = tbl_chk1_5(0).HWFGDLINE
    tbl_chk1_5_1(1).LINE = tbl_chk1_5(1).HWFGDLINE
'*** UPDATE ↑ Y.SIMIZU 2005/10/7 ﾗｲﾝ数追加
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    tbl_chk1_5_1(0).HWFGDSZY = tbl_chk1_5(0).HWFGDSZY
    tbl_chk1_5_1(1).HWFGDSZY = tbl_chk1_5(1).HWFGDSZY
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
    'GD/DSOD熱処理条件追加　06/12/22 ooba START =========================================>
    tbl_chk1_5_1(0).HWFANTNP = tbl_chk1_5(0).HWFANTNP
    tbl_chk1_5_1(1).HWFANTNP = tbl_chk1_5(1).HWFANTNP
    'GD/DSOD熱処理条件追加　06/12/22 ooba END ===========================================>
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,LDL")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00081"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    ''GD追加　05/01/27 ooba END ===================================================>
    
    ''06/05/31 ooba START ============================================>
    'ＳＰＶＮＲ
    sErr_Msg = "1-5 SPVNRﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "15", "NR", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_5_1
    tbl_chk1_5_1(0).HOSYOU = tbl_chk1_5(0).HWFNRHS      '保証方法＿対象
    tbl_chk1_5_1(1).HOSYOU = tbl_chk1_5(1).HWFNRHS
    tbl_chk1_5_1(0).SOKU_HOU = tbl_chk1_5(0).HWFNRSH    '測定位置＿方
    tbl_chk1_5_1(1).SOKU_HOU = tbl_chk1_5(1).HWFNRSH
    tbl_chk1_5_1(0).SOKU_TEN = tbl_chk1_5(0).HWFNRST    '測定位置＿点
    tbl_chk1_5_1(1).SOKU_TEN = tbl_chk1_5(1).HWFNRST
    tbl_chk1_5_1(0).SOKU_ICHI = tbl_chk1_5(0).HWFNRSI   '測定位置＿位
    tbl_chk1_5_1(1).SOKU_ICHI = tbl_chk1_5(1).HWFNRSI
    tbl_chk1_5_1(0).KENH_NUKI = tbl_chk1_5(0).HWFNRKN   '検査頻度＿抜
    tbl_chk1_5_1(1).KENH_NUKI = tbl_chk1_5(1).HWFNRKN
    tbl_chk1_5_1(0).PUAGEN = tbl_chk1_5(0).HWFNRPUG     'PUA限
    tbl_chk1_5_1(1).PUAGEN = tbl_chk1_5(1).HWFNRPUG
    tbl_chk1_5_1(0).PUAPER = tbl_chk1_5(0).HWFNRPUR     'PUA率
    tbl_chk1_5_1(1).PUAPER = tbl_chk1_5(1).HWFNRPUR
    
    RET = funChkFurikae1_5_1(sResult, tbl_chk1_5_1(), iErr_Code, sErr_Msg, "CHECK1-5,SPVNR")
    If RET <> 0 Then
        funChkFurikae1_5 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00077"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    ''06/05/31 ooba END ==============================================>
    
'------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae1_5 = 0 Then
        funChkFurikae1_5 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae1_5 = -4
    GoTo Apl_Exit
    
CodeDBGet_Error:
    Set rs = Nothing
    If funChkFurikae1_5 = 0 Then
        funChkFurikae1_5 = -5
    End If
    GoTo Apl_Exit

End Function

'------------------------------------------------
' 先行評価項目仕様詳細チェック
'------------------------------------------------

'概要      :指定されたﾁｪｯｸ内容詳細に基づき、該当する仕様値のチェックを行なう。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型                 :説明
'          :sChkCode        ,I  ,String             :チェック内容詳細
'          :tbl_chk1_5_1()  ,I  ,typ_chk1_5_1       :仕様値構造体配列
'          :iErr_Code       ,O  ,Integer            :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String             :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :sAdd_Msg        ,I  ,String             :添付ｴﾗｰﾒｯｾｰｼﾞ
'          :戻り値          ,O  ,Integer            :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/19 新規作成　SB

Public Function funChkFurikae1_5_1(sChkCode As String, tbl_chk1_5_1() As typ_chk1_5_1, _
                                   iErr_Code As Integer, sErr_Msg As String, sAdd_Msg As String) As Integer



    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim wHOSYOU_0   As String       '保証方法＿対象
    Dim wHOSYOU_1   As String       '保証方法＿対象
    Dim iCnt        As Integer      '04/04/13 ooba
    Dim sNum(2)     As String       '04/04/13 ooba
'↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.1 AN温度 振替可否チェック追加
    Dim lsCodeList() As String       'コードDBのコードのリスト
    Dim liNumCnt    As Integer
'↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------

    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae1_5_1 = 0
    iErr_Code = 0
    '------------------------------------------ 保証方法チェック ------------------------------------------------------
    If tbl_chk1_5_1(1).HOSYOU <> "H" And tbl_chk1_5_1(1).HOSYOU <> "S" Then GoTo Apl_Exit
    
    '------------------------------------------ 各種チェック ------------------------------------------------------
    '保証方法＿対象
    sErr_Msg = "保証方法_対象ﾁｪｯｸ"
    If Mid(sChkCode, 1, 1) = "2" Then
        '振替元と振替先が等しければ振替ＯＫ
        If tbl_chk1_5_1(0).HOSYOU <> tbl_chk1_5_1(1).HOSYOU Then
            
            wHOSYOU_0 = tbl_chk1_5_1(0).HOSYOU
            If tbl_chk1_5_1(0).HOSYOU <> "H" And tbl_chk1_5_1(0).HOSYOU <> "S" Then wHOSYOU_0 = "-"
            wHOSYOU_1 = tbl_chk1_5_1(1).HOSYOU
            If tbl_chk1_5_1(1).HOSYOU <> "H" And tbl_chk1_5_1(1).HOSYOU <> "S" Then wHOSYOU_1 = "-"
            
            'マトリクス取得
            sResult = ""
'            ret = funCodeDBGet("SB", "SH", tbl_chk1_5_1(0).HOSYOU, 1, tbl_chk1_5_1(1).HOSYOU, sResult)
            RET = funCodeDBGet("SB", "SH", wHOSYOU_0, 1, wHOSYOU_1, sResult)
            If RET <> 0 Then
                sErr_Msg = sAdd_Msg & sErr_Msg & "→元:" & tbl_chk1_5_1(0).HOSYOU & ", 先:" & tbl_chk1_5_1(1).HOSYOU
                GoTo CodeDBGet_Error
            End If
            If sResult = 0 Then
                funChkFurikae1_5_1 = 1
                iErr_Code = 1501
                GoTo Apl_Exit
            End If
        End If
    End If
    '下限
    sErr_Msg = "下限ﾁｪｯｸ"
    If Mid(sChkCode, 2, 1) = "1" Then
        If tbl_chk1_5_1(0).Min <> tbl_chk1_5_1(1).Min Then
            funChkFurikae1_5_1 = 1
            iErr_Code = 1502
            GoTo Apl_Exit
        End If
    End If
    '上限
    sErr_Msg = "上限ﾁｪｯｸ"
    If Mid(sChkCode, 3, 1) = "1" Then
        If tbl_chk1_5_1(0).max <> tbl_chk1_5_1(1).max Then
            funChkFurikae1_5_1 = 1
            iErr_Code = 1503
            GoTo Apl_Exit
        End If
    End If
    '測定位置＿方
    sErr_Msg = "測定位置_方ﾁｪｯｸ"
    If Mid(sChkCode, 4, 1) = "1" Then
        If Trim$(tbl_chk1_5_1(0).SOKU_HOU) <> Trim$(tbl_chk1_5_1(1).SOKU_HOU) Then
            funChkFurikae1_5_1 = 1
            iErr_Code = 1504
            GoTo Apl_Exit
        End If
    End If
    '測定位置＿点
    sErr_Msg = "測定位置_点ﾁｪｯｸ"
    If Mid(sChkCode, 5, 1) = "1" Then
        If Trim$(tbl_chk1_5_1(0).SOKU_TEN) <> Trim$(tbl_chk1_5_1(1).SOKU_TEN) Then
            funChkFurikae1_5_1 = 1
            iErr_Code = 1505
            GoTo Apl_Exit
        End If
    ElseIf Mid(sChkCode, 5, 1) = "2" Then   '08/01/29 ooba
        If Trim$(tbl_chk1_5_1(0).SOKU_TEN) = "" Or _
           Trim$(tbl_chk1_5_1(1).SOKU_TEN) = "" Or _
           Trim$(tbl_chk1_5_1(0).SOKU_TEN) < Trim$(tbl_chk1_5_1(1).SOKU_TEN) Then
            funChkFurikae1_5_1 = 1
            iErr_Code = 1505
            GoTo Apl_Exit
        End If
    End If
    '測定位置＿位
    sErr_Msg = "測定位置_位ﾁｪｯｸ"
    If Mid(sChkCode, 6, 1) = "2" Then
       'マトリクス取得
        sResult = ""
        RET = funCodeDBGet("SB", "OI", tbl_chk1_5_1(0).SOKU_ICHI, 1, tbl_chk1_5_1(1).SOKU_ICHI, sResult)
        If RET <> 0 Then
            sErr_Msg = sAdd_Msg & sErr_Msg & "→元:" & tbl_chk1_5_1(0).SOKU_ICHI & ", 先:" & tbl_chk1_5_1(1).SOKU_ICHI
            GoTo CodeDBGet_Error
        End If
        If sResult = 0 Then
            funChkFurikae1_5_1 = 1
            iErr_Code = 1506
            GoTo Apl_Exit
        End If
    End If
    
''Upd Start 2005/06/16 (TCS)T.Terauchi      SPV9点対応
    '測定位置＿位
    If Mid(sChkCode, 6, 1) = "1" Then
        If Trim$(tbl_chk1_5_1(0).SOKU_ICHI) <> Trim$(tbl_chk1_5_1(1).SOKU_ICHI) Then
            funChkFurikae1_5_1 = 1
            iErr_Code = 1506
            GoTo Apl_Exit
        End If
    End If
''Upd End   2005/06/16 (TCS)T.Terauchi      SPV9点対応
    
    '測定位置＿領
    sErr_Msg = "測定位置_領ﾁｪｯｸ"
    If Mid(sChkCode, 7, 1) = "1" Then
        If Trim$(tbl_chk1_5_1(0).SOKU_RYOU) <> Trim$(tbl_chk1_5_1(1).SOKU_RYOU) Then
            funChkFurikae1_5_1 = 1
            iErr_Code = 1507
            GoTo Apl_Exit
        End If
    End If
    '検査有無
    sErr_Msg = "検査有無ﾁｪｯｸ"
    If Mid(sChkCode, 8, 1) = "1" Then
        If Trim$(tbl_chk1_5_1(0).UMU) <> Trim$(tbl_chk1_5_1(1).UMU) Then
            funChkFurikae1_5_1 = 1
            iErr_Code = 1508
            GoTo Apl_Exit
        End If
    End If
    '熱処理法
    sErr_Msg = "熱処理法ﾁｪｯｸ"
    If Mid(sChkCode, 9, 1) = "1" Then
        If Trim$(tbl_chk1_5_1(0).NETSU) <> Trim$(tbl_chk1_5_1(1).NETSU) Then
            funChkFurikae1_5_1 = 1
            iErr_Code = 1509
            GoTo Apl_Exit
        End If
    End If
    '測定条件
    sErr_Msg = "測定条件ﾁｪｯｸ"
    If Mid(sChkCode, 10, 1) = "1" Then
        If Trim$(tbl_chk1_5_1(0).JOUKEN) <> Trim$(tbl_chk1_5_1(1).JOUKEN) Then
            funChkFurikae1_5_1 = 1
            iErr_Code = 1510
            GoTo Apl_Exit
        End If
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    ElseIf Mid(sChkCode, 10, 1) = "2" Then
        If Trim$(tbl_chk1_5_1(0).HWFGDSZY) = "F" And Trim$(tbl_chk1_5_1(1).HWFGDSZY) = "G" Then
            funChkFurikae1_5_1 = 1
            iErr_Code = 1510
            GoTo Apl_Exit
        End If
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
    End If
    '選択ＥＴ代
    sErr_Msg = "選択ET代ﾁｪｯｸ"
    If Mid(sChkCode, 11, 1) = "1" Then
        If tbl_chk1_5_1(0).ET <> tbl_chk1_5_1(1).ET Then
            funChkFurikae1_5_1 = 1
            iErr_Code = 1511
            GoTo Apl_Exit
        End If
    End If
    '検査方法
    sErr_Msg = "検査方法ﾁｪｯｸ"
    If Mid(sChkCode, 12, 1) = "1" Then
        If Trim$(tbl_chk1_5_1(0).KENSA) <> Trim$(tbl_chk1_5_1(1).KENSA) Then
            funChkFurikae1_5_1 = 1
            iErr_Code = 1512
            GoTo Apl_Exit
        End If
    End If
    'パターン区分
    sErr_Msg = "ﾊﾟﾀｰﾝ区分ﾁｪｯｸ"
    If Mid(sChkCode, 13, 1) = "2" Then
        'マトリクス取得
        sResult = ""
        RET = funCodeDBGet("SB", "OS", tbl_chk1_5_1(0).PATTERN, 1, tbl_chk1_5_1(1).PATTERN, sResult)
        If RET <> 0 Then
            sErr_Msg = sAdd_Msg & sErr_Msg & "→元:" & tbl_chk1_5_1(0).PATTERN & ", 先:" & tbl_chk1_5_1(1).PATTERN
            GoTo CodeDBGet_Error
        End If
        If sResult = 0 Then
            funChkFurikae1_5_1 = 1
            iErr_Code = 1513
            GoTo Apl_Exit
        End If
    End If
    ''検査頻度＿抜　04/04/13 ooba
    sErr_Msg = "検査頻度_抜ﾁｪｯｸ"
    If Mid(sChkCode, 14, 1) = "2" Then
        'マトリクス取得
        sResult = ""
        
        For iCnt = 0 To 1
            Select Case tbl_chk1_5_1(iCnt).KENH_NUKI
            Case "3", "4", "6"
                sNum(iCnt) = tbl_chk1_5_1(iCnt).KENH_NUKI
            Case Else
                sNum(iCnt) = "ETC"
            End Select
        Next
        
        RET = funCodeDBGet("SB", "HO", sNum(0), 1, sNum(1), sResult)
        If RET <> 0 Then
            sErr_Msg = sAdd_Msg & sErr_Msg & "→元:" & tbl_chk1_5_1(0).KENH_NUKI & ", 先:" & tbl_chk1_5_1(1).KENH_NUKI
            GoTo CodeDBGet_Error
        End If
        If sResult = 0 Then
            funChkFurikae1_5_1 = 1
            iErr_Code = 1514
            GoTo Apl_Exit
        End If
    End If
    
'*** UPDATE ↓ Y.SIMIZU 2005/10/7 ﾗｲﾝ数追加
    ''ﾗｲﾝ数
    sErr_Msg = "ﾗｲﾝ数"
    If Mid(sChkCode, 15, 1) = "2" Then
        'マトリクス取得
        sResult = ""
        
        For iCnt = 0 To 1
            sNum(iCnt) = tbl_chk1_5_1(iCnt).LINE
        Next
        
        RET = funCodeDBGet("SB", "LN", sNum(0), 1, sNum(1), sResult)
        If RET <> 0 Then
            sErr_Msg = sAdd_Msg & sErr_Msg & "→元:" & tbl_chk1_5_1(0).LINE & ", 先:" & tbl_chk1_5_1(1).LINE
            GoTo CodeDBGet_Error
        End If
        If sResult = 0 Then
            funChkFurikae1_5_1 = 1
            iErr_Code = 1515
            GoTo Apl_Exit
        End If
    End If
'*** UPDATE ↑ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数追加
'↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'2.1.1 AN温度 振替可否チェック追加
    ''AN温度
    sErr_Msg = "AN温度ﾁｪｯｸ"
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -s-
    If Mid(sChkCode, 16, 1) = "2" Then
''    If Mid(sChkCode, 16, 1) = "1" Then
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -s-
        'マトリクス取得
        sResult = ""
        
        For iCnt = 0 To 1
            sNum(iCnt) = CStr(Trim(tbl_chk1_5_1(iCnt).HWFANTNP))
        Next
        '' コードマスタのコードの一覧を取得
        RET = funCodeDBGetCodeList("SB", "AE", lsCodeList)
        If RET <> 0 Then
            sErr_Msg = sAdd_Msg & sErr_Msg & "→元:" & tbl_chk1_5_1(0).HWFANTNP & ", 先:" & tbl_chk1_5_1(1).HWFANTNP
            GoTo CodeDBGet_Error
        End If
        ''コードマスタに登録されていないコードはスペースに変換する
        For liNumCnt = 0 To 1
            RET = 0
            For iCnt = 1 To UBound(lsCodeList)
                If Trim(lsCodeList(iCnt)) = Trim(sNum(liNumCnt)) Then
                    RET = 1
                    Exit For
                End If
            Next iCnt
            If RET = 0 Then
                sNum(liNumCnt) = "     "
            End If
        Next liNumCnt
        
        ''項目により使用マトリックスが違うので場合分けする
        If Trim(Right(sAdd_Msg, 2)) = "RS" Then     '比抵抗チェック
            RET = funCodeDBGet("SB", "AR", sNum(1), 1, sNum(0), sResult)
        ElseIf Trim(Right(sAdd_Msg, 2)) = "Oi" Then '酸素濃度チェック
            RET = funCodeDBGet("SB", "AO", sNum(1), 1, sNum(0), sResult)
        ElseIf Trim(Right(sAdd_Msg, 2)) = "DS" Then     'DSODチェック　06/12/22 ooba
            RET = funCodeDBGet("SB", "AD", sNum(1), 1, sNum(0), sResult)
        ElseIf Trim(Right(sAdd_Msg, 3)) = "DEN" Or _
               Trim(Right(sAdd_Msg, 3)) = "DVD" Or _
               Trim(Right(sAdd_Msg, 3)) = "LDL" Then    'GDチェック　06/12/22 ooba
            RET = funCodeDBGet("SB", "AG", sNum(1), 1, sNum(0), sResult)
        Else                                        'その他
            RET = funCodeDBGet("SB", "AE", sNum(1), 1, sNum(0), sResult)
        End If
        
        If RET <> 0 Then
            sErr_Msg = sAdd_Msg & sErr_Msg & "→元:" & tbl_chk1_5_1(0).HWFANTNP & ", 先:" & tbl_chk1_5_1(1).HWFANTNP
            GoTo CodeDBGet_Error
        End If
        If sResult = 0 Then
            funChkFurikae1_5_1 = 1
            iErr_Code = 1516
            ''メッセージの中に温度を入れたいので、エラーメッセージはここで作成する
            sAdd_Msg = sAdd_Msg & "のAN温度が振替不可能です。(" & tbl_chk1_5_1(0).HWFANTNP & "℃ → " & tbl_chk1_5_1(1).HWFANTNP & "℃)"
            GoTo Apl_Exit
        End If
    End If
'↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    
    'PUA限　06/05/31 ooba
    sErr_Msg = "PUA限ﾁｪｯｸ"
    If Mid(sChkCode, 17, 1) = "1" Then
        If tbl_chk1_5_1(1).SOKU_HOU & tbl_chk1_5_1(1).SOKU_TEN & tbl_chk1_5_1(1).SOKU_ICHI = "AMX" Then
            If tbl_chk1_5_1(0).PUAGEN <> tbl_chk1_5_1(1).PUAGEN And tbl_chk1_5_1(1).PUAGEN <> -1 Then
                funChkFurikae1_5_1 = 1
                iErr_Code = 1517
                GoTo Apl_Exit
            End If
        End If
    End If
    'PUA率　06/05/31 ooba
    sErr_Msg = "PUA率ﾁｪｯｸ"
    If Mid(sChkCode, 18, 1) = "1" Then
        If tbl_chk1_5_1(1).SOKU_HOU & tbl_chk1_5_1(1).SOKU_TEN & tbl_chk1_5_1(1).SOKU_ICHI = "AMX" Then
            If tbl_chk1_5_1(0).PUAPER <> tbl_chk1_5_1(1).PUAPER And tbl_chk1_5_1(1).PUAPER <> -1 Then
                funChkFurikae1_5_1 = 1
                iErr_Code = 1518
                GoTo Apl_Exit
            End If
        End If
    End If
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Select Case iErr_Code
        Case 1501
            sErr_Msg = sAdd_Msg & "の保証方法が不一致の為、振替できません。"
        Case 1502
            sErr_Msg = sAdd_Msg & "の下限が不一致の為、振替できません。"
        Case 1503
            sErr_Msg = sAdd_Msg & "の上限が不一致の為、振替できません。"
        Case 1504
            sErr_Msg = sAdd_Msg & "の測定位置＿方が不一致の為、振替できません。"
        Case 1505
            sErr_Msg = sAdd_Msg & "の測定位置＿点が不一致の為、振替できません。"
        Case 1506
        
        ''Upd Start 2005/06/16 (TCS)T.Terauchi      SPV9点対応
        ''    sErr_Msg = sAdd_Msg & "の測定位置＿位が振替不可能です。"
            If Mid(sChkCode, 6, 1) = "2" Then
                sErr_Msg = sAdd_Msg & "の測定位置＿位が振替不可能です。"
            Else
                sErr_Msg = sAdd_Msg & "の測定位置＿位が不一致の為、振替できません。"
            End If
        ''Upd End   2005/06/16 (TCS)T.Terauchi      SPV9点対応
        
        Case 1507
            sErr_Msg = sAdd_Msg & "の測定位置＿領が不一致の為、振替できません。"
        Case 1508
            sErr_Msg = sAdd_Msg & "の検査有無が不一致の為、振替できません。"
        Case 1509
            sErr_Msg = sAdd_Msg & "の熱処理法が不一致の為、振替できません。"
        Case 1510
            sErr_Msg = sAdd_Msg & "の測定条件が不一致の為、振替できません。"
        Case 1511
            sErr_Msg = sAdd_Msg & "の選択ＥＴ代が不一致の為、振替できません。"
        Case 1512
            sErr_Msg = sAdd_Msg & "の検査方法が不一致の為、振替できません。"
        Case 1513
            sErr_Msg = sAdd_Msg & "のパターン区分が振替不可能です。"
        Case 1514
            sErr_Msg = sAdd_Msg & "の検査頻度＿抜が振替不可能です。"
    '*** UPDATE ↓ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数対応
        Case 1515
            sErr_Msg = sAdd_Msg & "のGDライン数が振替不可能です。"
    '*** UPDATE ↑ Y.SIMIZU 2005/10/7 GDﾗｲﾝ数対応
    
    '↓↓↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    '2.1.1 AN温度 振替可否チェック追加
        Case 1516
            sErr_Msg = sAdd_Msg
    '↑↑↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    
        Case 1517   '06/05/31 ooba
            sErr_Msg = sAdd_Msg & "のPUA限が不一致の為、振替できません。"
        Case 1518   '06/05/31 ooba
            sErr_Msg = sAdd_Msg & "のPUA率が不一致の為、振替できません。"
                
    End Select
    
    Exit Function
    
Apl_down:
    funChkFurikae1_5_1 = -4
    GoTo Apl_Exit
    
CodeDBGet_Error:
    If funChkFurikae1_5_1 = 0 Then
        funChkFurikae1_5_1 = -5
    End If
    GoTo Apl_Exit
    
End Function

'------------------------------------------------
' ナノトポ規格チェック
'------------------------------------------------

'概要      :振替元品番と振替先品番が、ガラス接着品かどうかを判断する。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sKeyID          ,I  ,String       :ﾌﾞﾛｯｸID、又は、SXL-ID
'          :tOld_Hinban     ,I  ,tFullHinban  :振替元品番(構造体)
'          :tNew_Hinban     ,I  ,tFullHinban  :振替先品番(構造体)
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/19 新規作成　SB

Public Function funChkFurikae1_6(sProccd As String, sBlockId As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iErr_Code As Integer, sErr_Msg As String) As Integer
    
    Dim W_OLD_FLG As Integer
    Dim W_NEW_FLG As Integer
    Dim sql As String               'SQL全体
    Dim rs  As OraDynaset           'RecordSet

    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae1_6 = 0
    
    '------------------------------------------ 振替元品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-6 振替元品番仕様取得(" & tOld_Hinban.hinban & Format(tOld_Hinban.mnorevno, "00") & tOld_Hinban.factory & tOld_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E026.HWFNP1AR,E026.HWFNP1MAX,E026.HWFNP2AR,E026.HWFNP2MAX,E018.HSXCSCEN " & vbCrLf
    sql = sql & "FROM   TBCME026 E026,TBCME018 E018 " & vbCrLf
    sql = sql & "WHERE  E026.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E026.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E026.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E026.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E018.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E018.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E018.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND   =   '" & tOld_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    Erase tbl_chk1_6
    With tbl_chk1_6(0)
        If IsNull(rs("HWFNP1AR")) = False Then .HWFNP1AR = rs("HWFNP1AR") Else .HWFNP1AR = 0            '品WFナノトポ１エリア
        If IsNull(rs("HWFNP1MAX")) = False Then .HWFNP1MAX = rs("HWFNP1MAX") Else .HWFNP1MAX = 0        '品WFナノトポ１上限
        If IsNull(rs("HWFNP2AR")) = False Then .HWFNP2AR = rs("HWFNP2AR") Else .HWFNP2AR = 0            '品WFナノトポ２エリア
        If IsNull(rs("HWFNP2MAX")) = False Then .HWFNP2MAX = rs("HWFNP2MAX") Else .HWFNP2MAX = 0        '品WFナノトポ２上限
        If IsNull(rs("HSXCSCEN")) = False Then .HSXCSCEN = rs("HSXCSCEN") Else .HSXCSCEN = 0            '結晶面傾き中心
    End With
    
    Set rs = Nothing
    '------------------------------------------ 振替先品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-6 振替先品番仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E026.HWFNP1AR,E026.HWFNP1MAX,E026.HWFNP2AR,E026.HWFNP2MAX,E018.HSXCSCEN " & vbCrLf
    sql = sql & "FROM   TBCME026 E026,TBCME018 E018 " & vbCrLf
    sql = sql & "WHERE  E026.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E026.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E026.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E026.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E018.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E018.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E018.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND   =   '" & tNew_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    With tbl_chk1_6(1)
        If IsNull(rs("HWFNP1AR")) = False Then .HWFNP1AR = rs("HWFNP1AR") Else .HWFNP1AR = 0            '品WFナノトポ１エリア
        If IsNull(rs("HWFNP1MAX")) = False Then .HWFNP1MAX = rs("HWFNP1MAX") Else .HWFNP1MAX = 0        '品WFナノトポ１上限
        If IsNull(rs("HWFNP2AR")) = False Then .HWFNP2AR = rs("HWFNP2AR") Else .HWFNP2AR = 0            '品WFナノトポ２エリア
        If IsNull(rs("HWFNP2MAX")) = False Then .HWFNP2MAX = rs("HWFNP2MAX") Else .HWFNP2MAX = 0        '品WFナノトポ２上限
        If IsNull(rs("HSXCSCEN")) = False Then .HSXCSCEN = rs("HSXCSCEN") Else .HSXCSCEN = 0            '結晶面傾き中心
    End With
    
    Set rs = Nothing
    '------------------------------------------ 各種チェック ------------------------------------------------------
    On Error GoTo Apl_down
    W_OLD_FLG = 0
    If tbl_chk1_6(0).HWFNP1AR = 2 And tbl_chk1_6(0).HWFNP1MAX <= 17 Or _
       tbl_chk1_6(0).HWFNP2AR = 10 And tbl_chk1_6(0).HWFNP2MAX <= 50 Then
        W_OLD_FLG = 1
    End If
    W_NEW_FLG = 0
    If tbl_chk1_6(1).HWFNP1AR = 2 And tbl_chk1_6(1).HWFNP1MAX <= 17 Or _
       tbl_chk1_6(1).HWFNP2AR = 10 And tbl_chk1_6(1).HWFNP2MAX <= 50 Then
        W_NEW_FLG = 1
    End If
    'ガラス接着品のチェック
    If W_OLD_FLG = 0 And W_NEW_FLG = 1 Then
        funChkFurikae1_6 = 1
        iErr_Code = 1601
        sErr_Msg = "CHECK1-6,ナノトポ規格外の為、振替できません。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00014"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If

'Del Start 2011/10/03 Y.Hitomi
'    '結晶面傾中心のチェック
'    sErr_Msg = "1-6 結晶面傾中心ﾁｪｯｸ"
'    If Trim$(tbl_chk1_6(0).HSXCSCEN) <> Trim$(tbl_chk1_6(1).HSXCSCEN) Then
'        funChkFurikae1_6 = 1
'        iErr_Code = 1602
'        sErr_Msg = "CHECK1-6,結晶面傾中心不一致の為、振替できません。"
''--------------- 2008/07/25 INSERT START  By Systech ---------------
'        gsTbcmy028ErrCode = "00004"
''--------------- 2008/07/25 INSERT  END   By Systech ---------------
'        GoTo Apl_Exit
'    End If
'Del End   2011/10/03 Y.Hitomi

'------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae1_6 = 0 Then
        funChkFurikae1_6 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae1_6 = -4
    GoTo Apl_Exit
    
End Function

'------------------------------------------------
' 品番組合せチェック１
'------------------------------------------------

'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :tKumi_Hinban()  ,I  ,tFullHinban  :ﾁｪｯｸ品番
'          :iKumi_Row()     ,I  ,Integer      :品番行位置
'          :iHinPnt         ,O  ,Integer      :ﾁｪｯｸNG品番位置
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :06/04/25 ooba

Public Function funChkFurikae1_7(tKumi_Hinban() As tFullHinban, iKumi_Row() As Integer, _
                                 iHinPnt As Integer, iErr_Code As Integer, _
                                 sErr_Msg As String) As Integer

    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim i           As Integer
    
    Dim sTmpType    As String       'タイプ                             2011/05/12
    Dim sTmpHinban  As String       '品番(タイプ組合せチェック用)       2011/05/12
    Dim sTmpDope    As String       'ドーパント                         2011/05/12
    Dim sTmpDpHinb  As String       '品番(ドーパント組合せチェック用)   2011/05/12
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae1_7 = 0
    
    If UBound(tKumi_Hinban) = 1 Then GoTo Apl_Exit
    
    '------------------------------------------ チェック元品番仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-7 チェック元品番仕様取得(" & tKumi_Hinban(1).hinban & Format(tKumi_Hinban(1).mnorevno, "00") & tKumi_Hinban(1).factory & tKumi_Hinban(1).opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E018.HSXTYPE,E018.HSXCDIR,E018.HSXCSCEN,E018.HSXDOP, " & vbCrLf
    sql = sql & "       E023.HWFCDOP,E020.HSXSDSLP " & vbCrLf
    sql = sql & "FROM   TBCME018 E018,TBCME023 E023,TBCME020 E020 " & vbCrLf
    sql = sql & "WHERE  E018.HINBAN    =   '" & tKumi_Hinban(1).hinban & "'  AND " & vbCrLf
    sql = sql & "       E018.MNOREVNO  =    " & tKumi_Hinban(1).mnorevno & " AND " & vbCrLf
    sql = sql & "       E018.FACTORY   =   '" & tKumi_Hinban(1).factory & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND   =   '" & tKumi_Hinban(1).opecond & "' AND " & vbCrLf
    sql = sql & "       E023.HINBAN    =   '" & tKumi_Hinban(1).hinban & "'  AND " & vbCrLf
    sql = sql & "       E023.MNOREVNO  =    " & tKumi_Hinban(1).mnorevno & " AND " & vbCrLf
    sql = sql & "       E023.FACTORY   =   '" & tKumi_Hinban(1).factory & "' AND " & vbCrLf
    sql = sql & "       E023.OPECOND   =   '" & tKumi_Hinban(1).opecond & "' AND " & vbCrLf
    sql = sql & "       E020.HINBAN    =   '" & tKumi_Hinban(1).hinban & "'  AND " & vbCrLf
    sql = sql & "       E020.MNOREVNO  =    " & tKumi_Hinban(1).mnorevno & " AND " & vbCrLf
    sql = sql & "       E020.FACTORY   =   '" & tKumi_Hinban(1).factory & "' AND " & vbCrLf
    sql = sql & "       E020.OPECOND   =   '" & tKumi_Hinban(1).opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    Erase tbl_chk1_7
    With tbl_chk1_7(0)
        If IsNull(rs("HSXTYPE")) = False Then .HSXTYPE = rs("HSXTYPE") Else .HSXTYPE = " "          ' 結晶面方位
        If IsNull(rs("HSXCDIR")) = False Then .HSXCDIR = rs("HSXCDIR") Else .HSXCDIR = " "          ' 結晶面方位
        If IsNull(rs("HSXCSCEN")) = False Then .HSXCSCEN = rs("HSXCSCEN") Else .HSXCSCEN = 0        ' 結晶面傾き中心
        If IsNull(rs("HSXDOP")) = False Then .HSXDOP = rs("HSXDOP") Else .HSXDOP = " "              ' ドーパント
        If IsNull(rs("HWFCDOP")) = False Then .HWFCDOP = rs("HWFCDOP") Else .HWFCDOP = " "          ' 結晶ドープ
        If IsNull(rs("HSXSDSLP")) = False Then .HSXSDSLP = rs("HSXSDSLP") Else .HSXSDSLP = " "      ' シード傾き
    End With
    
'>>>>> PN不問品番の組合せチェックを解除 2011/05/12 SETsw kubota ------------
    '一番目の品番を保存
    sTmpType = Trim$(tbl_chk1_7(0).HSXTYPE)
    sTmpHinban = tKumi_Hinban(1).hinban
    sTmpDope = Trim$(tbl_chk1_7(0).HSXDOP)
    sTmpDpHinb = tKumi_Hinban(1).hinban
'<<<<<< PN不問品番の組合せチェックを解除 2011/05/12 SETsw kubota ------------
    
    Set rs = Nothing
    
    For i = 2 To UBound(tKumi_Hinban)
        iHinPnt = iKumi_Row(i)      '品番位置ｾｯﾄ
        If tKumi_Hinban(1).hinban <> tKumi_Hinban(i).hinban Or _
           tKumi_Hinban(1).mnorevno <> tKumi_Hinban(i).mnorevno Or _
           tKumi_Hinban(1).factory <> tKumi_Hinban(i).factory Or _
           tKumi_Hinban(1).opecond <> tKumi_Hinban(i).opecond Then
           
            '---------------------------------- チェック先品番仕様データ取得 ------------------------------------------------------
            sErr_Msg = "1-7 チェック先品番仕様取得(" & tKumi_Hinban(i).hinban & Format(tKumi_Hinban(i).mnorevno, "00") & tKumi_Hinban(i).factory & tKumi_Hinban(i).opecond & ")"
            'SQL文の作成
            sql = vbNullString
            sql = sql & "SELECT E018.HSXTYPE,E018.HSXCDIR,E018.HSXCSCEN,E018.HSXDOP, " & vbCrLf
            sql = sql & "       E023.HWFCDOP,E020.HSXSDSLP " & vbCrLf
            sql = sql & "FROM   TBCME018 E018,TBCME023 E023,TBCME020 E020 " & vbCrLf
            sql = sql & "WHERE  E018.HINBAN    =   '" & tKumi_Hinban(i).hinban & "'  AND " & vbCrLf
            sql = sql & "       E018.MNOREVNO  =    " & tKumi_Hinban(i).mnorevno & " AND " & vbCrLf
            sql = sql & "       E018.FACTORY   =   '" & tKumi_Hinban(i).factory & "' AND " & vbCrLf
            sql = sql & "       E018.OPECOND   =   '" & tKumi_Hinban(i).opecond & "' AND " & vbCrLf
            sql = sql & "       E023.HINBAN    =   '" & tKumi_Hinban(i).hinban & "'  AND " & vbCrLf
            sql = sql & "       E023.MNOREVNO  =    " & tKumi_Hinban(i).mnorevno & " AND " & vbCrLf
            sql = sql & "       E023.FACTORY   =   '" & tKumi_Hinban(i).factory & "' AND " & vbCrLf
            sql = sql & "       E023.OPECOND   =   '" & tKumi_Hinban(i).opecond & "' AND " & vbCrLf
            sql = sql & "       E020.HINBAN    =   '" & tKumi_Hinban(i).hinban & "'  AND " & vbCrLf
            sql = sql & "       E020.MNOREVNO  =    " & tKumi_Hinban(i).mnorevno & " AND " & vbCrLf
            sql = sql & "       E020.FACTORY   =   '" & tKumi_Hinban(i).factory & "' AND " & vbCrLf
            sql = sql & "       E020.OPECOND   =   '" & tKumi_Hinban(i).opecond & "' " & vbCrLf
            
            On Error GoTo db_Error
            'SQL文の実行
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            
            '該当データなし
            If rs.EOF Or rs.RecordCount > 1 Then
                GoTo db_Error
            End If
            
            '取得データセット
            With tbl_chk1_7(1)
                If IsNull(rs("HSXTYPE")) = False Then .HSXTYPE = rs("HSXTYPE") Else .HSXTYPE = " "          ' タイプ
                If IsNull(rs("HSXCDIR")) = False Then .HSXCDIR = rs("HSXCDIR") Else .HSXCDIR = " "          ' 結晶面方位
                If IsNull(rs("HSXCSCEN")) = False Then .HSXCSCEN = rs("HSXCSCEN") Else .HSXCSCEN = 0        ' 結晶面傾き中心
                If IsNull(rs("HSXDOP")) = False Then .HSXDOP = rs("HSXDOP") Else .HSXDOP = " "              ' ドーパント
                If IsNull(rs("HWFCDOP")) = False Then .HWFCDOP = rs("HWFCDOP") Else .HWFCDOP = " "          ' WF結晶ドープ
                If IsNull(rs("HSXSDSLP")) = False Then .HSXSDSLP = rs("HSXSDSLP") Else .HSXSDSLP = " "      ' シード傾き
            End With
            
            Set rs = Nothing
            '---------------------------------- 指示取得 ------------------------------------------------------
            On Error GoTo Apl_down
            'タイプのチェック
            sErr_Msg = "1-7 ﾀｲﾌﾟﾁｪｯｸ"
'>>>>> PN不問品番の組合せチェックを解除 2011/05/12 SETsw kubota ------------
'            If Trim$(tbl_chk1_7(0).HSXTYPE) <> Trim$(tbl_chk1_7(1).HSXTYPE) Then
'                funChkFurikae1_7 = 1
'                iErr_Code = 1701
'                sErr_Msg = "ﾀｲﾌﾟ ⇒ " & tKumi_Hinban(1).hinban & "：" & tbl_chk1_7(0).HSXTYPE & "，" & tKumi_Hinban(i).hinban & "：" & tbl_chk1_7(1).HSXTYPE
'                GoTo Apl_Exit
'            End If
            'タイプが異なるか比較
            If sTmpType <> Trim$(tbl_chk1_7(1).HSXTYPE) Then
                '異なっても、どちらかがZ:不問の場合はエラーとしない
                If sTmpType <> "Z" _
                And Trim$(tbl_chk1_7(1).HSXTYPE) <> "Z" Then
                    funChkFurikae1_7 = 1
                    iErr_Code = 1701
                    sErr_Msg = "ﾀｲﾌﾟ ⇒ " & sTmpHinban & "：" & sTmpType & "，" & tKumi_Hinban(i).hinban & "：" & tbl_chk1_7(1).HSXTYPE
                    GoTo Apl_Exit
                End If
            End If
            '次回比較するタイプと品番を保存
            If Trim$(tbl_chk1_7(1).HSXTYPE) <> "Z" Then
                sTmpType = Trim$(tbl_chk1_7(1).HSXTYPE)
                sTmpHinban = tKumi_Hinban(i).hinban         '品番(エラーメッセージ用)
            End If
'<<<<< PN不問品番の組合せチェックを解除 2011/05/12 SETsw kubota ------------
            '結晶面方位のチェック
            sErr_Msg = "1-7 結晶面方位ﾁｪｯｸ"
            If Trim$(tbl_chk1_7(0).HSXCDIR) <> Trim$(tbl_chk1_7(1).HSXCDIR) Then
                funChkFurikae1_7 = 1
                iErr_Code = 1702
                sErr_Msg = "結晶面方位 ⇒ " & tKumi_Hinban(1).hinban & "：" & tbl_chk1_7(0).HSXCDIR & "，" & tKumi_Hinban(i).hinban & "：" & tbl_chk1_7(1).HSXCDIR
                GoTo Apl_Exit
            End If
            '結晶面傾中心のチェック
            sErr_Msg = "1-7 結晶面傾中心ﾁｪｯｸ"
            If (Trim$(tbl_chk1_7(0).HSXCSCEN) = 4) Or (Trim$(tbl_chk1_7(1).HSXCSCEN) = 4) Then
                If Trim$(tbl_chk1_7(0).HSXCSCEN) <> Trim$(tbl_chk1_7(1).HSXCSCEN) Then
                    funChkFurikae1_7 = 1
                    iErr_Code = 1703
                    sErr_Msg = "結晶面傾中心 ⇒ " & tKumi_Hinban(1).hinban & "：" & tbl_chk1_7(0).HSXCSCEN & "，" & tKumi_Hinban(i).hinban & "：" & tbl_chk1_7(1).HSXCSCEN
                    GoTo Apl_Exit
                End If
            End If
            'ドーパントのチェック
            sErr_Msg = "1-7 ﾄﾞｰﾊﾟﾝﾄﾁｪｯｸ"
'>>>>> ドーパント不問品番の組合せチェックを解除 2011/05/12 SETsw kubota ------------
'            If Trim$(tbl_chk1_7(0).HSXDOP) <> Trim$(tbl_chk1_7(1).HSXDOP) Then
'                funChkFurikae1_7 = 1
'                iErr_Code = 1704
'                sErr_Msg = "ﾄﾞｰﾊﾟﾝﾄ ⇒ " & tKumi_Hinban(1).hinban & "：" & tbl_chk1_7(0).HSXDOP & "，" & tKumi_Hinban(i).hinban & "：" & tbl_chk1_7(1).HSXDOP
'                GoTo Apl_Exit
'            End If
            'ドーパントが異なるか比較
            If sTmpDope <> Trim$(tbl_chk1_7(1).HSXDOP) Then
                '異なっても、どちらかがZ:不問の場合はエラーとしない
                If sTmpDope <> "Z" _
                And Trim$(tbl_chk1_7(1).HSXDOP) <> "Z" Then
                    funChkFurikae1_7 = 1
                    iErr_Code = 1704
                    sErr_Msg = "ﾄﾞｰﾊﾟﾝﾄ ⇒ " & sTmpDpHinb & "：" & sTmpDope & "，" & tKumi_Hinban(i).hinban & "：" & tbl_chk1_7(1).HSXDOP
                    GoTo Apl_Exit
                End If
            End If
            '次回比較するドーパントと品番を保存
            If Trim$(tbl_chk1_7(1).HSXDOP) <> "Z" Then
                sTmpDope = Trim$(tbl_chk1_7(1).HSXDOP)
                sTmpDpHinb = tKumi_Hinban(i).hinban         '品番(エラーメッセージ用)
            End If
'<<<<< ドーパント不問品番の組合せチェックを解除 2011/05/12 SETsw kubota ------------
            
            '↓ﾁｪｯｸ停止　06/07/28 ooba START ====================================================>
''            'WF結晶ドープのチェック
''            sErr_Msg = "1-7 WF結晶ﾄﾞｰﾌﾟﾁｪｯｸ"
''            If Trim$(tbl_chk1_7(0).HWFCDOP) <> Trim$(tbl_chk1_7(1).HWFCDOP) Then
''                funChkFurikae1_7 = 1
''                iErr_Code = 1705
''                sErr_Msg = "WF結晶ﾄﾞｰﾌﾟ ⇒ " & tKumi_Hinban(1).hinban & "：" & tbl_chk1_7(0).HWFCDOP & "，" & tKumi_Hinban(i).hinban & "：" & tbl_chk1_7(1).HWFCDOP
''                GoTo Apl_Exit
''            End If
            '↑ﾁｪｯｸ停止　06/07/28 ooba END ======================================================>
            
            'シード傾きのチェック
            sErr_Msg = "1-7 ｼｰﾄﾞ傾きﾁｪｯｸ"
            If Trim$(tbl_chk1_7(0).HSXSDSLP) <> Trim$(tbl_chk1_7(1).HSXSDSLP) Then
                funChkFurikae1_7 = 1
                iErr_Code = 1706
                sErr_Msg = "ｼｰﾄﾞ傾き ⇒ " & tKumi_Hinban(1).hinban & "：" & tbl_chk1_7(0).HSXSDSLP & "，" & tKumi_Hinban(i).hinban & "：" & tbl_chk1_7(1).HSXSDSLP
                GoTo Apl_Exit
            End If
        End If
    Next i
    
    '溝位置方位のチェック
    For i = 2 To UBound(tKumi_Hinban)
        iHinPnt = iKumi_Row(i)      '品番位置ｾｯﾄ
        '｢溝位置方位｣はﾌﾞﾛｯｸ単位のﾁｪｯｸ
        If tKumi_Hinban(i).Hinkubun = tKumi_Hinban(i - 1).Hinkubun Then
        
            '---------------------------------- チェック元品番仕様データ取得 ------------------------------------------------------
            sErr_Msg = "1-7 チェック元品番仕様取得(" & tKumi_Hinban(i - 1).hinban & Format(tKumi_Hinban(i - 1).mnorevno, "00") & tKumi_Hinban(i - 1).factory & tKumi_Hinban(i - 1).opecond & ")"
            'SQL文の作成
            sql = vbNullString
            sql = sql & "SELECT HSXDPDIR " & vbCrLf
            sql = sql & "FROM   TBCME018 " & vbCrLf
            sql = sql & "WHERE  HINBAN    =   '" & tKumi_Hinban(i - 1).hinban & "'  AND " & vbCrLf
            sql = sql & "       MNOREVNO  =    " & tKumi_Hinban(i - 1).mnorevno & " AND " & vbCrLf
            sql = sql & "       FACTORY   =   '" & tKumi_Hinban(i - 1).factory & "' AND " & vbCrLf
            sql = sql & "       OPECOND   =   '" & tKumi_Hinban(i - 1).opecond & "' " & vbCrLf
            
            On Error GoTo db_Error
            'SQL文の実行
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            '該当データなし
            If rs.EOF Or rs.RecordCount > 1 Then
                GoTo db_Error
            End If
            '取得データセット
            With tbl_chk1_7(0)
                If IsNull(rs("HSXDPDIR")) = False Then .HSXDPDIR = rs("HSXDPDIR") Else .HSXDPDIR = " "      ' 溝位置方位
            End With
            Set rs = Nothing
            
            '---------------------------------- チェック先品番仕様データ取得 ------------------------------------------------------
            sErr_Msg = "1-7 チェック先品番仕様取得(" & tKumi_Hinban(i).hinban & Format(tKumi_Hinban(i).mnorevno, "00") & tKumi_Hinban(i).factory & tKumi_Hinban(i).opecond & ")"
            'SQL文の作成
            sql = vbNullString
            sql = sql & "SELECT HSXDPDIR " & vbCrLf
            sql = sql & "FROM   TBCME018 " & vbCrLf
            sql = sql & "WHERE  HINBAN    =   '" & tKumi_Hinban(i).hinban & "'  AND " & vbCrLf
            sql = sql & "       MNOREVNO  =    " & tKumi_Hinban(i).mnorevno & " AND " & vbCrLf
            sql = sql & "       FACTORY   =   '" & tKumi_Hinban(i).factory & "' AND " & vbCrLf
            sql = sql & "       OPECOND   =   '" & tKumi_Hinban(i).opecond & "' " & vbCrLf
            
            On Error GoTo db_Error
            'SQL文の実行
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            '該当データなし
            If rs.EOF Or rs.RecordCount > 1 Then
                GoTo db_Error
            End If
            '取得データセット
            With tbl_chk1_7(1)
                If IsNull(rs("HSXDPDIR")) = False Then .HSXDPDIR = rs("HSXDPDIR") Else .HSXDPDIR = " "      ' 溝位置方位
            End With
            Set rs = Nothing
            
            '---------------------------------- 指示取得 ------------------------------------------------------
            On Error GoTo Apl_down
            '溝位置方位のチェック（同一分類グループなら組合せ可能）
            sErr_Msg = "1-7 溝位置方位ﾁｪｯｸ"
            sResult = ""
            RET = funCodeDBGet("SB", "MZ", tbl_chk1_7(0).HSXDPDIR, 1, tbl_chk1_7(1).HSXDPDIR, sResult)
            If RET <> 0 Then
                sErr_Msg = sErr_Msg & "→" & tKumi_Hinban(i - 1).hinban & "：" & tbl_chk1_7(0).HSXDPDIR & "，" & tKumi_Hinban(i).hinban & "：" & tbl_chk1_7(1).HSXDPDIR
                GoTo CodeDBGet_Error
            End If
            If sResult = 0 Then
                funChkFurikae1_7 = 1
                iErr_Code = 1707
                sErr_Msg = "溝位置方位 ⇒ " & tKumi_Hinban(i - 1).hinban & "：" & tbl_chk1_7(0).HSXDPDIR & "，" & tKumi_Hinban(i).hinban & "：" & tbl_chk1_7(1).HSXDPDIR
                GoTo Apl_Exit
            End If
        End If
    Next i
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae1_7 = 0 Then
        funChkFurikae1_7 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae1_7 = -4
    GoTo Apl_Exit
    
CodeDBGet_Error:
    Set rs = Nothing
    If funChkFurikae1_7 = 0 Then
        funChkFurikae1_7 = -5
    End If
    GoTo Apl_Exit

End Function

'------------------------------------------------
' 品番組合せチェック２
'------------------------------------------------

'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :tKumi_Hinban()  ,I  ,tFullHinban  :ﾁｪｯｸ品番
'          :iKumi_Row()     ,I  ,Integer      :品番行位置
'          :iHinPnt         ,O  ,Integer      :ﾁｪｯｸNG品番位置
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :06/04/25 ooba
'          :06/07/19 SMP)kondoh

Public Function funChkFurikae1_8(tKumi_Hinban() As tFullHinban, iKumi_Row() As Integer, _
                                 iHinPnt As Integer, iErr_Code As Integer, _
                                 sErr_Msg As String) As Integer

    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim i           As Integer

''06/07/19 SMP)kondoh START Add =========================================================>
' チェック元品番を 基準品番(狙い品番) から ブロック内の代表品番 に変更
    Dim l           As Integer
    Dim m           As Integer
    Dim SQLHIN      As String
''06/07/19 SMP)kondoh END Add =========================================================>

    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae1_8 = 0

''06/07/19 SMP)kondoh START Del =========================================================>
' チェック元品番を 基準品番(狙い品番) から ブロック内の代表品番 に変更したため、取得位置を移動

''    '------------------------------------------ チェック元品番仕様データ取得 ------------------------------------------------------
''    sErr_Msg = "1-8 チェック元品番仕様取得(" & tKumi_Hinban(0).hinban & Format(tKumi_Hinban(0).mnorevno, "00") & tKumi_Hinban(0).factory & tKumi_Hinban(0).opecond & ")"
''    'SQL文の作成
''    sql = vbNullString
''    sql = sql & "SELECT E020.HSXCDOP,E036.GLASS,E036.SLICEATU, " & vbCrLf
''    sql = sql & "       E018.HSXCSMIN,E018.HSXCSMAX,E020.HSXWFWAR " & vbCrLf
''    sql = sql & "FROM   TBCME018 E018,TBCME020 E020,TBCME036 E036 " & vbCrLf
''    sql = sql & "WHERE  E018.HINBAN    =   '" & tKumi_Hinban(0).hinban & "'  AND " & vbCrLf
''    sql = sql & "       E018.MNOREVNO  =    " & tKumi_Hinban(0).mnorevno & " AND " & vbCrLf
''    sql = sql & "       E018.FACTORY   =   '" & tKumi_Hinban(0).factory & "' AND " & vbCrLf
''    sql = sql & "       E018.OPECOND   =   '" & tKumi_Hinban(0).opecond & "' AND " & vbCrLf
''    sql = sql & "       E020.HINBAN    =   '" & tKumi_Hinban(0).hinban & "'  AND " & vbCrLf
''    sql = sql & "       E020.MNOREVNO  =    " & tKumi_Hinban(0).mnorevno & " AND " & vbCrLf
''    sql = sql & "       E020.FACTORY   =   '" & tKumi_Hinban(0).factory & "' AND " & vbCrLf
''    sql = sql & "       E020.OPECOND   =   '" & tKumi_Hinban(0).opecond & "' AND " & vbCrLf
''    sql = sql & "       E036.HINBAN    =   '" & tKumi_Hinban(0).hinban & "'  AND " & vbCrLf
''    sql = sql & "       E036.MNOREVNO  =    " & tKumi_Hinban(0).mnorevno & " AND " & vbCrLf
''    sql = sql & "       E036.FACTORY   =   '" & tKumi_Hinban(0).factory & "' AND " & vbCrLf
''    sql = sql & "       E036.OPECOND   =   '" & tKumi_Hinban(0).opecond & "' " & vbCrLf
''
''    On Error GoTo db_Error
''    'SQL文の実行
''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''
''    '該当データなし
''    If rs.EOF Or rs.RecordCount > 1 Then
''        GoTo db_Error
''    End If
''
''    '取得データセット
''    Erase tbl_chk1_8
''    With tbl_chk1_8(0)
''        If IsNull(rs("HSXCDOP")) = False Then .HSXCDOP = rs("HSXCDOP") Else .HSXCDOP = " "          ' SX結晶ドープ
''        'C/N/M/Z以外はその他扱い
''        If (.HSXCDOP <> "C" And .HSXCDOP <> "N" And .HSXCDOP <> "M" And .HSXCDOP <> "Z") Then
''            .HSXCDOP = " "
''        End If
''        If IsNull(rs("GLASS")) = False Then .GLASS = rs("GLASS") Else .GLASS = " "                  ' ガラス接着
''        If IsNull(rs("SLICEATU")) = False Then .SLICEATU = rs("SLICEATU") Else .SLICEATU = 0        ' SL厚み
''        If IsNull(rs("HSXCSMIN")) = False Then .HSXCSMIN = rs("HSXCSMIN") Else .HSXCSMIN = 0        ' 結晶面傾下限(合成角度)
''        If IsNull(rs("HSXCSMAX")) = False Then .HSXCSMAX = rs("HSXCSMAX") Else .HSXCSMAX = 0        ' 結晶面傾上限(合成角度)
''        If IsNull(rs("HSXWFWAR")) = False Then .HSXWFWAR = rs("HSXWFWAR") Else .HSXWFWAR = " "      ' Warpランク
''    End With
''
''    Set rs = Nothing
''06/07/19 SMP)kondoh END Del =========================================================>


    For i = 1 To UBound(tKumi_Hinban)
        iHinPnt = iKumi_Row(i)      '品番位置ｾｯﾄ


''06/07/19 SMP)kondoh START Add =========================================================>
' チェック元品番を 基準品番(狙い品番) から ブロック内の代表品番 に変更

        'ﾌﾞﾛｯｸの切れ目でﾌﾞﾛｯｸ内の代表品番を取得する
        If tKumi_Hinban(i).Hinkubun <> tKumi_Hinban(i - 1).Hinkubun Then
            
            SQLHIN = vbNullString
            For l = 1 To UBound(tKumi_Hinban)
                If tKumi_Hinban(i).Hinkubun = tKumi_Hinban(l).Hinkubun Then
                    SQLHIN = SQLHIN & "(HINBAN='" & tKumi_Hinban(l).hinban & "'"
                    SQLHIN = SQLHIN & " and MNOREVNO=" & tKumi_Hinban(l).mnorevno
                    SQLHIN = SQLHIN & " and FACTORY='" & tKumi_Hinban(l).factory & "'"
                    SQLHIN = SQLHIN & " and OPECOND='" & tKumi_Hinban(l).opecond & "') or "
                End If
            Next l
            SQLHIN = "(" & left(SQLHIN, Len(SQLHIN) - 4) & ")"
            
            sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, HWFWARPR from TBCME027 where "
            '①ﾅﾉﾄﾎﾟﾌﾗｸﾞ(0:ｶﾞﾗｽ接着無し,1:ｶﾞﾗｽ接着有り)が最大な品番の中で
            '合成角の規格幅(結晶面傾上限-結晶面傾下限)が最小の品番の中で
            'ﾜｰﾌﾟﾗﾝｸが最大の品番
            sql = sql & "HWFWARPR = (select MAX(HWFWARPR) from TBCME027 "
            sql = sql & "            where (HINBAN, MNOREVNO, FACTORY, OPECOND) in "
            sql = sql & "               ("
            sql = sql & "               select HINBAN, MNOREVNO, FACTORY, OPECOND "
            sql = sql & "               from TBCME018 "
            sql = sql & "               where ABS(HSXCSMAX - HSXCSMIN) = "
            sql = sql & "                       (select MIN(ABS(HSXCSMAX - HSXCSMIN)) from TBCME018 "
            sql = sql & "                      where (HINBAN, MNOREVNO, FACTORY, OPECOND) in "
            sql = sql & "                           (select HINBAN, MNOREVNO, FACTORY, OPECOND from TBCME036 "
            sql = sql & "                           where decode(GLASS,null,'0',' ','0',GLASS) = "
            sql = sql & "                               (select MAX(decode(GLASS,null,'0',' ','0',GLASS)) "
            sql = sql & "                               from TBCME036 where " & SQLHIN
            sql = sql & "                               ) "
            sql = sql & "                           and " & SQLHIN
            sql = sql & "                           ) "
            sql = sql & "                       ) "
            sql = sql & "                and (HINBAN, MNOREVNO, FACTORY, OPECOND) in "
            sql = sql & "                  (select HINBAN, MNOREVNO, FACTORY, OPECOND from TBCME036 "
            sql = sql & "                   where decode(GLASS,null,'0',' ','0',GLASS) = "
            sql = sql & "                         (select MAX(decode(GLASS,null,'0',' ','0',GLASS)) "
            sql = sql & "                          from TBCME036 where " & SQLHIN
            sql = sql & "                         )"
            sql = sql & "                   and " & SQLHIN
            sql = sql & "                  ) "
            sql = sql & "               ) "
            sql = sql & "           ) "
            sql = sql & "and (HINBAN, MNOREVNO, FACTORY, OPECOND) in "
            sql = sql & "               ("
            sql = sql & "               select HINBAN, MNOREVNO, FACTORY, OPECOND "
            sql = sql & "               from TBCME018 "
            sql = sql & "               where ABS(HSXCSMAX - HSXCSMIN) = "
            sql = sql & "                       (select MIN(ABS(HSXCSMAX - HSXCSMIN)) from TBCME018 "
            sql = sql & "                      where (HINBAN, MNOREVNO, FACTORY, OPECOND) in "
            sql = sql & "                           (select HINBAN, MNOREVNO, FACTORY, OPECOND from TBCME036 "
            sql = sql & "                           where decode(GLASS,null,'0',' ','0',GLASS) = "
            sql = sql & "                               (select MAX(decode(GLASS,null,'0',' ','0',GLASS)) "
            sql = sql & "                               from TBCME036 where " & SQLHIN
            sql = sql & "                               ) "
            sql = sql & "                           and " & SQLHIN
            sql = sql & "                           ) "
            sql = sql & "                       ) "
            sql = sql & "                and (HINBAN, MNOREVNO, FACTORY, OPECOND) in "
            sql = sql & "                  (select HINBAN, MNOREVNO, FACTORY, OPECOND from TBCME036 "
            sql = sql & "                   where decode(GLASS,null,'0',' ','0',GLASS) = "
            sql = sql & "                         (select MAX(decode(GLASS,null,'0',' ','0',GLASS)) "
            sql = sql & "                          from TBCME036 where " & SQLHIN
            sql = sql & "                         )"
            sql = sql & "                   and " & SQLHIN
            sql = sql & "                  ) "
            sql = sql & "               ) "
        
            On Error GoTo db_Error
            'SQL文の実行
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        
            If rs.RecordCount <= 0 Then
                rs.Close
                GoTo db_Error
            End If

            '' 複数存在する場合は、先頭品番を代表品番とする
            For l = 1 To UBound(tKumi_Hinban)
                If tKumi_Hinban(i).Hinkubun = tKumi_Hinban(l).Hinkubun Then
                    rs.MoveFirst
                    For m = 1 To rs.RecordCount
                        If rs("HINBAN") = tKumi_Hinban(l).hinban And _
                            rs("FACTORY") = tKumi_Hinban(l).factory And _
                            rs("MNOREVNO") = tKumi_Hinban(l).mnorevno And _
                            rs("OPECOND") = tKumi_Hinban(l).opecond Then
                            tKumi_Hinban(0).hinban = tKumi_Hinban(l).hinban
                            tKumi_Hinban(0).factory = tKumi_Hinban(l).factory
                            tKumi_Hinban(0).mnorevno = tKumi_Hinban(l).mnorevno
                            tKumi_Hinban(0).opecond = tKumi_Hinban(l).opecond
                            l = UBound(tKumi_Hinban)
                            Exit For
                        End If
                        rs.MoveNext
                    Next m
                End If
            Next l
            Set rs = Nothing

        
            ' 代表品番の仕様を取得する
            '------------------------------------------ チェック元品番仕様データ取得 ------------------------------------------------------
            sErr_Msg = "1-8 チェック元品番仕様取得(" & tKumi_Hinban(0).hinban & Format(tKumi_Hinban(0).mnorevno, "00") & tKumi_Hinban(0).factory & tKumi_Hinban(0).opecond & ")"
            'SQL文の作成
            sql = vbNullString
            sql = sql & "SELECT E020.HSXCDOP,E036.GLASS,E036.SLICEATU, " & vbCrLf
            sql = sql & "       E018.HSXCSMIN,E018.HSXCSMAX,E020.HSXWFWAR " & vbCrLf
            sql = sql & "       ,E036.KUMIDOP " & vbCrLf                                   '' 組合せドープフラグ 2006/07/21 SMP)kondoh Add
'--------------- 2008/08/25 INSERT START  By Systech ---------------
            sql = sql & "       ,NVL(E036.HSXDKTMP, ' ') AS HSXDKTMP " & vbCrLf                                   '' 組合せドープフラグ 2006/07/21 SMP)kondoh Add
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
            sql = sql & "FROM   TBCME018 E018,TBCME020 E020,TBCME036 E036 " & vbCrLf
            sql = sql & "WHERE  E018.HINBAN    =   '" & tKumi_Hinban(0).hinban & "'  AND " & vbCrLf
            sql = sql & "       E018.MNOREVNO  =    " & tKumi_Hinban(0).mnorevno & " AND " & vbCrLf
            sql = sql & "       E018.FACTORY   =   '" & tKumi_Hinban(0).factory & "' AND " & vbCrLf
            sql = sql & "       E018.OPECOND   =   '" & tKumi_Hinban(0).opecond & "' AND " & vbCrLf
            sql = sql & "       E020.HINBAN    =   '" & tKumi_Hinban(0).hinban & "'  AND " & vbCrLf
            sql = sql & "       E020.MNOREVNO  =    " & tKumi_Hinban(0).mnorevno & " AND " & vbCrLf
            sql = sql & "       E020.FACTORY   =   '" & tKumi_Hinban(0).factory & "' AND " & vbCrLf
            sql = sql & "       E020.OPECOND   =   '" & tKumi_Hinban(0).opecond & "' AND " & vbCrLf
            sql = sql & "       E036.HINBAN    =   '" & tKumi_Hinban(0).hinban & "'  AND " & vbCrLf
            sql = sql & "       E036.MNOREVNO  =    " & tKumi_Hinban(0).mnorevno & " AND " & vbCrLf
            sql = sql & "       E036.FACTORY   =   '" & tKumi_Hinban(0).factory & "' AND " & vbCrLf
            sql = sql & "       E036.OPECOND   =   '" & tKumi_Hinban(0).opecond & "' " & vbCrLf
            
            On Error GoTo db_Error
            'SQL文の実行
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            
            '該当データなし
            If rs.EOF Or rs.RecordCount > 1 Then
                GoTo db_Error
            End If
            
            '取得データセット
            Erase tbl_chk1_8
            With tbl_chk1_8(0)
                If IsNull(rs("HSXCDOP")) = False Then .HSXCDOP = rs("HSXCDOP") Else .HSXCDOP = " "          ' SX結晶ドープ
                'C/N/M/Z以外はその他扱い
                If (.HSXCDOP <> "C" And .HSXCDOP <> "N" And .HSXCDOP <> "M" And .HSXCDOP <> "Z") Then
                    .HSXCDOP = " "
                End If
                If IsNull(rs("GLASS")) = False Then .GLASS = rs("GLASS") Else .GLASS = " "                  ' ガラス接着
                If IsNull(rs("SLICEATU")) = False Then .SLICEATU = rs("SLICEATU") Else .SLICEATU = 0        ' SL厚み
                If IsNull(rs("HSXCSMIN")) = False Then .HSXCSMIN = rs("HSXCSMIN") Else .HSXCSMIN = 0        ' 結晶面傾下限(合成角度)
                If IsNull(rs("HSXCSMAX")) = False Then .HSXCSMAX = rs("HSXCSMAX") Else .HSXCSMAX = 0        ' 結晶面傾上限(合成角度)
                If IsNull(rs("HSXWFWAR")) = False Then .HSXWFWAR = rs("HSXWFWAR") Else .HSXWFWAR = " "      ' Warpランク
                '' 2006/07/21 SMP)kondoh START Add
                If IsNull(rs("KUMIDOP")) = False Then .KUMIDOP = rs("KUMIDOP") Else .KUMIDOP = "0"      ' 組合せドープフラグ
                '1(C)/2(N)/3(M)/4(Z)/5(N可)以外は0(選択なし)扱い
                If (.KUMIDOP <> "1" And .KUMIDOP <> "2" And .KUMIDOP <> "3" And .KUMIDOP <> "4" And .KUMIDOP <> "5") Then
                    .KUMIDOP = "0"
                End If
                '' 2006/07/21 SMP)kondoh END Add
'--------------- 2008/08/25 INSERT START  By Systech ---------------
                .HSXDKTMP = rs("HSXDKTMP")
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
            End With
            
            Set rs = Nothing
        
        End If
''06/07/19 SMP)kondoh END Add =========================================================>

        If tKumi_Hinban(0).hinban <> tKumi_Hinban(i).hinban Or _
           tKumi_Hinban(0).mnorevno <> tKumi_Hinban(i).mnorevno Or _
           tKumi_Hinban(0).factory <> tKumi_Hinban(i).factory Or _
           tKumi_Hinban(0).opecond <> tKumi_Hinban(i).opecond Then
           
            '---------------------------------- チェック先品番仕様データ取得 ------------------------------------------------------
            sErr_Msg = "1-8 チェック先品番仕様取得(" & tKumi_Hinban(i).hinban & Format(tKumi_Hinban(i).mnorevno, "00") & tKumi_Hinban(i).factory & tKumi_Hinban(i).opecond & ")"
            'SQL文の作成
            sql = vbNullString
            sql = sql & "SELECT E020.HSXCDOP,E036.GLASS,E036.SLICEATU, " & vbCrLf
            sql = sql & "       E018.HSXCSMIN,E018.HSXCSMAX,E020.HSXWFWAR " & vbCrLf
            sql = sql & "       ,E036.KUMIDOP " & vbCrLf                                   '' 組合せドープフラグ 2006/07/21 SMP)kondoh Add
'--------------- 2008/08/25 INSERT START  By Systech ---------------
            sql = sql & "       ,NVL(E036.HSXDKTMP, ' ') AS HSXDKTMP " & vbCrLf                                   '' 組合せドープフラグ 2006/07/21 SMP)kondoh Add
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
            sql = sql & "FROM   TBCME018 E018,TBCME020 E020,TBCME036 E036 " & vbCrLf
            sql = sql & "WHERE  E018.HINBAN    =   '" & tKumi_Hinban(i).hinban & "'  AND " & vbCrLf
            sql = sql & "       E018.MNOREVNO  =    " & tKumi_Hinban(i).mnorevno & " AND " & vbCrLf
            sql = sql & "       E018.FACTORY   =   '" & tKumi_Hinban(i).factory & "' AND " & vbCrLf
            sql = sql & "       E018.OPECOND   =   '" & tKumi_Hinban(i).opecond & "' AND " & vbCrLf
            sql = sql & "       E020.HINBAN    =   '" & tKumi_Hinban(i).hinban & "'  AND " & vbCrLf
            sql = sql & "       E020.MNOREVNO  =    " & tKumi_Hinban(i).mnorevno & " AND " & vbCrLf
            sql = sql & "       E020.FACTORY   =   '" & tKumi_Hinban(i).factory & "' AND " & vbCrLf
            sql = sql & "       E020.OPECOND   =   '" & tKumi_Hinban(i).opecond & "' AND " & vbCrLf
            sql = sql & "       E036.HINBAN    =   '" & tKumi_Hinban(i).hinban & "'  AND " & vbCrLf
            sql = sql & "       E036.MNOREVNO  =    " & tKumi_Hinban(i).mnorevno & " AND " & vbCrLf
            sql = sql & "       E036.FACTORY   =   '" & tKumi_Hinban(i).factory & "' AND " & vbCrLf
            sql = sql & "       E036.OPECOND   =   '" & tKumi_Hinban(i).opecond & "' " & vbCrLf
            
            On Error GoTo db_Error
            'SQL文の実行
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            
            '該当データなし
            If rs.EOF Or rs.RecordCount > 1 Then
                GoTo db_Error
            End If
            
            '取得データセット
            With tbl_chk1_8(1)
                If IsNull(rs("HSXCDOP")) = False Then .HSXCDOP = rs("HSXCDOP") Else .HSXCDOP = " "          ' SX結晶ドープ
                'C/N/M/Z以外はその他扱い
                If (.HSXCDOP <> "C" And .HSXCDOP <> "N" And .HSXCDOP <> "M" And .HSXCDOP <> "Z") Then
                    .HSXCDOP = " "
                End If
                If IsNull(rs("GLASS")) = False Then .GLASS = rs("GLASS") Else .GLASS = " "                  ' ガラス接着
                If IsNull(rs("SLICEATU")) = False Then .SLICEATU = rs("SLICEATU") Else .SLICEATU = 0        ' SL厚み
                If IsNull(rs("HSXCSMIN")) = False Then .HSXCSMIN = rs("HSXCSMIN") Else .HSXCSMIN = 0        ' 結晶面傾下限(合成角度)
                If IsNull(rs("HSXCSMAX")) = False Then .HSXCSMAX = rs("HSXCSMAX") Else .HSXCSMAX = 0        ' 結晶面傾上限(合成角度)
                If IsNull(rs("HSXWFWAR")) = False Then .HSXWFWAR = rs("HSXWFWAR") Else .HSXWFWAR = " "      ' Warpランク
                '' 2006/07/21 SMP)kondoh START Add
                If IsNull(rs("KUMIDOP")) = False Then .KUMIDOP = rs("KUMIDOP") Else .KUMIDOP = "0"      ' 組合せドープフラグ
                '1(C)/2(N)/3(M)/4(Z)/5(N可)以外は0(選択なし)扱い
                If (.KUMIDOP <> "1" And .KUMIDOP <> "2" And .KUMIDOP <> "3" And .KUMIDOP <> "4" And .KUMIDOP <> "5") Then
                    .KUMIDOP = "0"
                End If
                '' 2006/07/21 SMP)kondoh END Add
'--------------- 2008/08/25 INSERT START  By Systech ---------------
                .HSXDKTMP = rs("HSXDKTMP")
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
            End With
            
            Set rs = Nothing
            '---------------------------------- 指示取得 ------------------------------------------------------
            On Error GoTo Apl_down
            
''06/07/19 SMP)kondoh START Del =========================================================>
''            'SX結晶ドープのチェック
''            sErr_Msg = "1-8 SX結晶ﾄﾞｰﾌﾟﾁｪｯｸ"
''            sResult = ""
''            RET = funCodeDBGet("SB", "DP", tbl_chk1_8(0).HSXCDOP, 1, tbl_chk1_8(1).HSXCDOP, sResult)
''            If RET <> 0 Then
''                sErr_Msg = sErr_Msg & "→元:" & tbl_chk1_8(0).HSXCDOP & ", 先:" & tbl_chk1_8(1).HSXCDOP
''                GoTo CodeDBGet_Error
''            End If
''            If sResult = 0 Then
''                funChkFurikae1_8 = 1
''                iErr_Code = 1801
''                sErr_Msg = "SX結晶ﾄﾞｰﾌﾟ ⇒ 元：" & tbl_chk1_8(0).HSXCDOP & "，先：" & tbl_chk1_8(1).HSXCDOP
''                GoTo Apl_Exit
''            End If
''            'ガラス接着のチェック
''            sErr_Msg = "1-8 ｶﾞﾗｽ接着ﾁｪｯｸ"
''            If Trim$(tbl_chk1_8(0).GLASS) <> "1" And Trim$(tbl_chk1_8(1).GLASS) = "1" Then
''                funChkFurikae1_8 = 1
''                iErr_Code = 1802
''                sErr_Msg = "ｶﾞﾗｽ接着 ⇒ 元：" & tbl_chk1_8(0).GLASS & ", 先:" & tbl_chk1_8(1).GLASS
''                sErr_Msg = "ｶﾞﾗｽ接着 ⇒ " & tKumi_Hinban(i).hinban & Format(tKumi_Hinban(i).mnorevno, "00") & tKumi_Hinban(i).factory & tKumi_Hinban(i).opecond & "：" & tbl_chk1_8(1).GLASS
''                GoTo Apl_Exit
''            End If
''06/07/19 SMP)kondoh END Del =========================================================>
            'SL厚みのチェック
            sErr_Msg = "1-8 SL厚みﾁｪｯｸ"
            If tbl_chk1_8(0).SLICEATU <> tbl_chk1_8(1).SLICEATU Then
                funChkFurikae1_8 = 1
                iErr_Code = 1803
''06/07/19 SMP)kondoh START Cng =========================================================>
''                sErr_Msg = "SL厚み ⇒ 元：" & tbl_chk1_8(0).SLICEATU & ", 先:" & tbl_chk1_8(1).SLICEATU
                sErr_Msg = "SL厚み ⇒ " & tKumi_Hinban(0).hinban & "：" & tbl_chk1_8(0).SLICEATU & "，" & tKumi_Hinban(i).hinban & "：" & tbl_chk1_8(1).SLICEATU
''06/07/19 SMP)kondoh END Cng =========================================================>
                GoTo Apl_Exit
            End If
            
            '↓ﾁｪｯｸ停止　06/08/22 ooba START ==================================================>
''            '合成角度のチェック
''            sErr_Msg = "1-8 合成角度ﾁｪｯｸ"
''            If Abs(tbl_chk1_8(0).HSXCSMAX - tbl_chk1_8(0).HSXCSMIN) > _
''               Abs(tbl_chk1_8(1).HSXCSMAX - tbl_chk1_8(1).HSXCSMIN) Then
''                funChkFurikae1_8 = 1
''                iErr_Code = 1804
''''06/07/19 SMP)kondoh START Cng =========================================================>
''''                sErr_Msg = "合成角度 ⇒ 元：" & Abs(tbl_chk1_8(0).HSXCSMAX - tbl_chk1_8(0).HSXCSMIN) & _
''''                            ", 先:" & Abs(tbl_chk1_8(1).HSXCSMAX - tbl_chk1_8(1).HSXCSMIN)
''                sErr_Msg = "合成角度 ⇒ " & tKumi_Hinban(0).hinban & "：" & Abs(tbl_chk1_8(0).HSXCSMAX - tbl_chk1_8(0).HSXCSMIN) & _
''                            "，" & tKumi_Hinban(i).hinban & "：" & Abs(tbl_chk1_8(1).HSXCSMAX - tbl_chk1_8(1).HSXCSMIN)
''''06/07/19 SMP)kondoh END Cng =========================================================>
''                GoTo Apl_Exit
''            End If
            '↑ﾁｪｯｸ停止　06/08/22 ooba END ====================================================>
            
            'Warpランクのチェック
            sErr_Msg = "1-8 Warpﾗﾝｸﾁｪｯｸ"
            If IsNumeric(tbl_chk1_8(0).HSXWFWAR) And IsNumeric(tbl_chk1_8(1).HSXWFWAR) Then
                If CInt(tbl_chk1_8(0).HSXWFWAR) < CInt(tbl_chk1_8(1).HSXWFWAR) Then
                    funChkFurikae1_8 = 1
                    iErr_Code = 1805
''06/07/19 SMP)kondoh START Cng =========================================================>
''                    sErr_Msg = "Warpﾗﾝｸ ⇒ 元：" & tbl_chk1_8(0).HSXWFWAR & ", 先:" & tbl_chk1_8(1).HSXWFWAR
                    sErr_Msg = "Warpﾗﾝｸ ⇒ " & tKumi_Hinban(0).hinban & "：" & tbl_chk1_8(0).HSXWFWAR & "，" & tKumi_Hinban(i).hinban & "：" & tbl_chk1_8(1).HSXWFWAR
''06/07/19 SMP)kondoh END Cng =========================================================>
                    GoTo Apl_Exit
                End If
            End If

''06/07/21 SMP)kondoh START Add =========================================================>
            '組合せドープフラグのチェック
            sErr_Msg = "1-8 組合せﾄﾞｰﾌﾟﾌﾗｸﾞﾁｪｯｸ"
            sResult = ""
            RET = funCodeDBGet("SB", "DP", tbl_chk1_8(0).KUMIDOP, 1, tbl_chk1_8(1).KUMIDOP, sResult)
            If RET <> 0 Then
                sErr_Msg = "組合せﾄﾞｰﾌﾟﾌﾗｸﾞ⇒" & tKumi_Hinban(0).hinban & "：" & tbl_chk1_8(0).KUMIDOP & "，" & tKumi_Hinban(i).hinban & "：" & tbl_chk1_8(1).KUMIDOP
                GoTo CodeDBGet_Error
            End If
            If sResult = 0 Then
                funChkFurikae1_8 = 1
                iErr_Code = 1801
                sErr_Msg = "組合せﾄﾞｰﾌﾟﾌﾗｸﾞ⇒" & tKumi_Hinban(0).hinban & "：" & tbl_chk1_8(0).KUMIDOP & "，" & tKumi_Hinban(i).hinban & "：" & tbl_chk1_8(1).KUMIDOP
                GoTo Apl_Exit
            End If
''06/07/21 SMP)kondoh END Del =========================================================>

'--------------- 2008/08/25 INSERT START  By Systech ---------------
            ' DK温度のチェック
            sErr_Msg = "1-8 DK温度ﾁｪｯｸ"
            sResult = ""
            If ((tbl_chk1_8(0).HSXDKTMP = DKTMP_650_20OV Or tbl_chk1_8(0).HSXDKTMP = DKTMP_650_20LO) And _
                (tbl_chk1_8(1).HSXDKTMP = DKTMP_650_20OV Or tbl_chk1_8(1).HSXDKTMP = DKTMP_650_20LO)) Or _
               (tbl_chk1_8(0).HSXDKTMP = DKTMP_1100 And tbl_chk1_8(1).HSXDKTMP = DKTMP_1100) Or _
               (Trim(tbl_chk1_8(0).HSXDKTMP) = "" And Trim(tbl_chk1_8(1).HSXDKTMP) = "") Then
            Else
               ' 温度が異なる場合は、ＮＧ
                funChkFurikae1_8 = 1
                iErr_Code = 1806
                sErr_Msg = "DK温度⇒" & _
                            tKumi_Hinban(0).hinban & "：" & GetDKTmpDispName("" & GetGPCodeCont(DKTMP_TBCME033CODE, tbl_chk1_8(0).HSXDKTMP)) & "℃，" & _
                            tKumi_Hinban(i).hinban & "：" & GetDKTmpDispName("" & GetGPCodeCont(DKTMP_TBCME033CODE, tbl_chk1_8(1).HSXDKTMP)) & "℃"
                GoTo Apl_Exit
            End If
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
        End If
    Next i
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae1_8 = 0 Then
        funChkFurikae1_8 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae1_8 = -4
    GoTo Apl_Exit
    
CodeDBGet_Error:
    Set rs = Nothing
    If funChkFurikae1_8 = 0 Then
        funChkFurikae1_8 = -5
    End If
    GoTo Apl_Exit

End Function


'------------------------------------------------
' 振替元と振替先のエピ先行評価項目仕様チェック
'------------------------------------------------

'概要      :振替元品番と振替先品番の先行評価項目仕様をチェックする。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sKeyID          ,I  ,String       :ﾌﾞﾛｯｸID、又は、SXL-ID
'          :tOld_Hinban     ,I  ,tFullHinban  :振替元品番(構造体)
'          :tNew_Hinban     ,I  ,tFullHinban  :振替先品番(構造体)
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :2006/08/15 新規作成 エピ先行評価追加対応 SMP)kondoh
Public Function funChkFurikae1_9(sProccd As String, sBlockId As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iErr_Code As Integer, sErr_Msg As String) As Integer



    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql As String               'SQL全体
    Dim rs  As OraDynaset           'RecordSet

    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae1_9 = 0
    
    '------------------------------------------ 振替元品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-9 振替元品番仕様取得(" & tOld_Hinban.hinban & Format(tOld_Hinban.mnorevno, "00") & tOld_Hinban.factory & tOld_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E050.HEPOF1HS,E050.HEPOF1SH,E050.HEPOF1ST,E050.HEPOF1SR,E050.HEPOF1NS,E050.HEPOF1SZ,E050.HEPOF1ET,E050.HEPOSF1PTK,E050.HEPOF1KN,   " & vbCrLf
    sql = sql & "       E050.HEPOF2HS,E050.HEPOF2SH,E050.HEPOF2ST,E050.HEPOF2SR,E050.HEPOF2NS,E050.HEPOF2SZ,E050.HEPOF2ET,E050.HEPOSF2PTK,E050.HEPOF2KN,   " & vbCrLf
    sql = sql & "       E050.HEPOF3HS,E050.HEPOF3SH,E050.HEPOF3ST,E050.HEPOF3SR,E050.HEPOF3NS,E050.HEPOF3SZ,E050.HEPOF3ET,E050.HEPOSF3PTK,E050.HEPOF3KN,   " & vbCrLf
    sql = sql & "       E050.HEPBM1HS,E050.HEPBM1SH,E050.HEPBM1ST,E050.HEPBM1SR,E050.HEPBM1NS,E050.HEPBM1SZ,E050.HEPBM1ET,E050.HEPBM1KN,   " & vbCrLf
    sql = sql & "       E050.HEPBM2HS,E050.HEPBM2SH,E050.HEPBM2ST,E050.HEPBM2SR,E050.HEPBM2NS,E050.HEPBM2SZ,E050.HEPBM2ET,E050.HEPBM2KN,   " & vbCrLf
    sql = sql & "       E050.HEPBM3HS,E050.HEPBM3SH,E050.HEPBM3ST,E050.HEPBM3SR,E050.HEPBM3NS,E050.HEPBM3SZ,E050.HEPBM3ET,E050.HEPBM3KN,   " & vbCrLf
    sql = sql & "       E050.HEPANTNP,E050.HEPACEN " & vbCrLf   'AN温度
    sql = sql & "FROM   TBCME050 E050 " & vbCrLf
    sql = sql & "WHERE  E050.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E050.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E050.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E050.OPECOND   =   '" & tOld_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If

    '取得データセット
    Erase tbl_chk1_9
    With tbl_chk1_9(0)
        'OSF1E
        If IsNull(rs("HEPOF1HS")) = False Then .HEPOF1HS = rs("HEPOF1HS") Else .HEPOF1HS = " "              '保証方法_処
        If IsNull(rs("HEPOF1SH")) = False Then .HEPOF1SH = rs("HEPOF1SH") Else .HEPOF1SH = " "              '測定位置_方
        If IsNull(rs("HEPOF1ST")) = False Then .HEPOF1ST = rs("HEPOF1ST") Else .HEPOF1ST = " "              '測定位置_点
        If IsNull(rs("HEPOF1SR")) = False Then .HEPOF1SR = rs("HEPOF1SR") Else .HEPOF1SR = " "              '測定位置_領
        If IsNull(rs("HEPOF1NS")) = False Then .HEPOF1NS = rs("HEPOF1NS") Else .HEPOF1NS = " "              '熱処理法
        If IsNull(rs("HEPOF1SZ")) = False Then .HEPOF1SZ = rs("HEPOF1SZ") Else .HEPOF1SZ = " "              '測定条件
        If IsNull(rs("HEPOSF1PTK")) = False Then .HEPOSF1PTK = rs("HEPOSF1PTK") Else .HEPOSF1PTK = " "      'パターン区分
        If IsNull(rs("HEPOF1ET")) = False Then .HEPOF1ET = rs("HEPOF1ET") Else .HEPOF1ET = 0                '選択ET代
        If IsNull(rs("HEPOF1KN")) = False Then .HEPOF1KN = rs("HEPOF1KN") Else .HEPOF1KN = " "              '測定位置_抜
        'OSF2E
        If IsNull(rs("HEPOF2HS")) = False Then .HEPOF2HS = rs("HEPOF2HS") Else .HEPOF2HS = " "              '保証方法_処
        If IsNull(rs("HEPOF2SH")) = False Then .HEPOF2SH = rs("HEPOF2SH") Else .HEPOF2SH = " "              '測定位置_方
        If IsNull(rs("HEPOF2ST")) = False Then .HEPOF2ST = rs("HEPOF2ST") Else .HEPOF2ST = " "              '測定位置_点
        If IsNull(rs("HEPOF2SR")) = False Then .HEPOF2SR = rs("HEPOF2SR") Else .HEPOF2SR = " "              '測定位置_領
        If IsNull(rs("HEPOF2NS")) = False Then .HEPOF2NS = rs("HEPOF2NS") Else .HEPOF2NS = " "              '熱処理法
        If IsNull(rs("HEPOF2SZ")) = False Then .HEPOF2SZ = rs("HEPOF2SZ") Else .HEPOF2SZ = " "              '測定条件
        If IsNull(rs("HEPOSF2PTK")) = False Then .HEPOSF2PTK = rs("HEPOSF2PTK") Else .HEPOSF2PTK = " "      'パターン区分
        If IsNull(rs("HEPOF2ET")) = False Then .HEPOF2ET = rs("HEPOF2ET") Else .HEPOF2ET = 0                '選択ET代
        If IsNull(rs("HEPOF2KN")) = False Then .HEPOF2KN = rs("HEPOF2KN") Else .HEPOF2KN = " "              '測定位置_抜
        'OSF3E
        If IsNull(rs("HEPOF3HS")) = False Then .HEPOF3HS = rs("HEPOF3HS") Else .HEPOF3HS = " "              '保証方法_処
        If IsNull(rs("HEPOF3SH")) = False Then .HEPOF3SH = rs("HEPOF3SH") Else .HEPOF3SH = " "              '測定位置_方
        If IsNull(rs("HEPOF3ST")) = False Then .HEPOF3ST = rs("HEPOF3ST") Else .HEPOF3ST = " "              '測定位置_点
        If IsNull(rs("HEPOF3SR")) = False Then .HEPOF3SR = rs("HEPOF3SR") Else .HEPOF3SR = " "              '測定位置_領
        If IsNull(rs("HEPOF3NS")) = False Then .HEPOF3NS = rs("HEPOF3NS") Else .HEPOF3NS = " "              '熱処理法
        If IsNull(rs("HEPOF3SZ")) = False Then .HEPOF3SZ = rs("HEPOF3SZ") Else .HEPOF3SZ = " "              '測定条件
        If IsNull(rs("HEPOSF3PTK")) = False Then .HEPOSF3PTK = rs("HEPOSF3PTK") Else .HEPOSF3PTK = " "      'パターン区分
        If IsNull(rs("HEPOF3ET")) = False Then .HEPOF3ET = rs("HEPOF3ET") Else .HEPOF3ET = 0                '選択ET代
        If IsNull(rs("HEPOF3KN")) = False Then .HEPOF3KN = rs("HEPOF3KN") Else .HEPOF3KN = " "              '測定位置_抜
        'BMD1E
        If IsNull(rs("HEPBM1HS")) = False Then .HEPBM1HS = rs("HEPBM1HS") Else .HEPBM1HS = " "              '保証方法_処
        If IsNull(rs("HEPBM1SH")) = False Then .HEPBM1SH = rs("HEPBM1SH") Else .HEPBM1SH = " "              '測定位置_方
        If IsNull(rs("HEPBM1ST")) = False Then .HEPBM1ST = rs("HEPBM1ST") Else .HEPBM1ST = " "              '測定位置_点
        If IsNull(rs("HEPBM1SR")) = False Then .HEPBM1SR = rs("HEPBM1SR") Else .HEPBM1SR = " "              '測定位置_領
        If IsNull(rs("HEPBM1NS")) = False Then .HEPBM1NS = rs("HEPBM1NS") Else .HEPBM1NS = " "              '熱処理法
        If IsNull(rs("HEPBM1SZ")) = False Then .HEPBM1SZ = rs("HEPBM1SZ") Else .HEPBM1SZ = " "              '測定条件
        If IsNull(rs("HEPBM1ET")) = False Then .HEPBM1ET = rs("HEPBM1ET") Else .HEPBM1ET = 0                '選択ET代
        If IsNull(rs("HEPBM1KN")) = False Then .HEPBM1KN = rs("HEPBM1KN") Else .HEPBM1KN = " "              '測定位置_抜
        'BMD2E
        If IsNull(rs("HEPBM2HS")) = False Then .HEPBM2HS = rs("HEPBM2HS") Else .HEPBM2HS = " "              '保証方法_処
        If IsNull(rs("HEPBM2SH")) = False Then .HEPBM2SH = rs("HEPBM2SH") Else .HEPBM2SH = " "              '測定位置_方
        If IsNull(rs("HEPBM2ST")) = False Then .HEPBM2ST = rs("HEPBM2ST") Else .HEPBM2ST = " "              '測定位置_点
        If IsNull(rs("HEPBM2SR")) = False Then .HEPBM2SR = rs("HEPBM2SR") Else .HEPBM2SR = " "              '測定位置_領
        If IsNull(rs("HEPBM2NS")) = False Then .HEPBM2NS = rs("HEPBM2NS") Else .HEPBM2NS = " "              '熱処理法
        If IsNull(rs("HEPBM2SZ")) = False Then .HEPBM2SZ = rs("HEPBM2SZ") Else .HEPBM2SZ = " "              '測定条件
        If IsNull(rs("HEPBM2ET")) = False Then .HEPBM2ET = rs("HEPBM2ET") Else .HEPBM2ET = 0                '選択ET代
        If IsNull(rs("HEPBM2KN")) = False Then .HEPBM2KN = rs("HEPBM2KN") Else .HEPBM2KN = " "              '測定位置_抜
        'BMD3E
        If IsNull(rs("HEPBM3HS")) = False Then .HEPBM3HS = rs("HEPBM3HS") Else .HEPBM3HS = " "              '保証方法_処
        If IsNull(rs("HEPBM3SH")) = False Then .HEPBM3SH = rs("HEPBM3SH") Else .HEPBM3SH = " "              '測定位置_方
        If IsNull(rs("HEPBM3ST")) = False Then .HEPBM3ST = rs("HEPBM3ST") Else .HEPBM3ST = " "              '測定位置_点
        If IsNull(rs("HEPBM3SR")) = False Then .HEPBM3SR = rs("HEPBM3SR") Else .HEPBM3SR = " "              '測定位置_領
        If IsNull(rs("HEPBM3NS")) = False Then .HEPBM3NS = rs("HEPBM3NS") Else .HEPBM3NS = " "              '熱処理法
        If IsNull(rs("HEPBM3SZ")) = False Then .HEPBM3SZ = rs("HEPBM3SZ") Else .HEPBM3SZ = " "              '測定条件
        If IsNull(rs("HEPBM3ET")) = False Then .HEPBM3ET = rs("HEPBM3ET") Else .HEPBM3ET = 0                '選択ET代
        If IsNull(rs("HEPBM3KN")) = False Then .HEPBM3KN = rs("HEPBM3KN") Else .HEPBM3KN = " "              '測定位置_抜
        'エピAN温度
        If IsNull(rs("HEPANTNP")) = False Then .HEPANTNP = rs("HEPANTNP") Else .HEPANTNP = 0                'AN温度
        'エピ厚
        If IsNull(rs("HEPACEN")) = False Then .HEPACEN = rs("HEPACEN") Else .HEPACEN = 0                    'エピ厚
    End With
    
    Set rs = Nothing

    '------------------------------------------ 振替先品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-9 振替先品番仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E050.HEPOF1HS,E050.HEPOF1SH,E050.HEPOF1ST,E050.HEPOF1SR,E050.HEPOF1NS,E050.HEPOF1SZ,E050.HEPOF1ET,E050.HEPOSF1PTK,E050.HEPOF1KN,   " & vbCrLf
    sql = sql & "       E050.HEPOF2HS,E050.HEPOF2SH,E050.HEPOF2ST,E050.HEPOF2SR,E050.HEPOF2NS,E050.HEPOF2SZ,E050.HEPOF2ET,E050.HEPOSF2PTK,E050.HEPOF2KN,   " & vbCrLf
    sql = sql & "       E050.HEPOF3HS,E050.HEPOF3SH,E050.HEPOF3ST,E050.HEPOF3SR,E050.HEPOF3NS,E050.HEPOF3SZ,E050.HEPOF3ET,E050.HEPOSF3PTK,E050.HEPOF3KN,   " & vbCrLf
    sql = sql & "       E050.HEPBM1HS,E050.HEPBM1SH,E050.HEPBM1ST,E050.HEPBM1SR,E050.HEPBM1NS,E050.HEPBM1SZ,E050.HEPBM1ET,E050.HEPBM1KN,   " & vbCrLf
    sql = sql & "       E050.HEPBM2HS,E050.HEPBM2SH,E050.HEPBM2ST,E050.HEPBM2SR,E050.HEPBM2NS,E050.HEPBM2SZ,E050.HEPBM2ET,E050.HEPBM2KN,   " & vbCrLf
    sql = sql & "       E050.HEPBM3HS,E050.HEPBM3SH,E050.HEPBM3ST,E050.HEPBM3SR,E050.HEPBM3NS,E050.HEPBM3SZ,E050.HEPBM3ET,E050.HEPBM3KN,   " & vbCrLf
    sql = sql & "       E050.HEPANTNP,E050.HEPACEN " & vbCrLf   'AN温度
    sql = sql & "FROM   TBCME050 E050 " & vbCrLf
    sql = sql & "WHERE  E050.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E050.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E050.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E050.OPECOND   =   '" & tNew_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    With tbl_chk1_9(1)
        'OSF1E
        If IsNull(rs("HEPOF1HS")) = False Then .HEPOF1HS = rs("HEPOF1HS") Else .HEPOF1HS = " "              '保証方法_処
        If IsNull(rs("HEPOF1SH")) = False Then .HEPOF1SH = rs("HEPOF1SH") Else .HEPOF1SH = " "              '測定位置_方
        If IsNull(rs("HEPOF1ST")) = False Then .HEPOF1ST = rs("HEPOF1ST") Else .HEPOF1ST = " "              '測定位置_点
        If IsNull(rs("HEPOF1SR")) = False Then .HEPOF1SR = rs("HEPOF1SR") Else .HEPOF1SR = " "              '測定位置_領
        If IsNull(rs("HEPOF1NS")) = False Then .HEPOF1NS = rs("HEPOF1NS") Else .HEPOF1NS = " "              '熱処理法
        If IsNull(rs("HEPOF1SZ")) = False Then .HEPOF1SZ = rs("HEPOF1SZ") Else .HEPOF1SZ = " "              '測定条件
        If IsNull(rs("HEPOSF1PTK")) = False Then .HEPOSF1PTK = rs("HEPOSF1PTK") Else .HEPOSF1PTK = " "      'パターン区分
        If IsNull(rs("HEPOF1ET")) = False Then .HEPOF1ET = rs("HEPOF1ET") Else .HEPOF1ET = 0                '選択ET代
        If IsNull(rs("HEPOF1KN")) = False Then .HEPOF1KN = rs("HEPOF1KN") Else .HEPOF1KN = " "              '測定位置_抜
        'OSF2E
        If IsNull(rs("HEPOF2HS")) = False Then .HEPOF2HS = rs("HEPOF2HS") Else .HEPOF2HS = " "              '保証方法_処
        If IsNull(rs("HEPOF2SH")) = False Then .HEPOF2SH = rs("HEPOF2SH") Else .HEPOF2SH = " "              '測定位置_方
        If IsNull(rs("HEPOF2ST")) = False Then .HEPOF2ST = rs("HEPOF2ST") Else .HEPOF2ST = " "              '測定位置_点
        If IsNull(rs("HEPOF2SR")) = False Then .HEPOF2SR = rs("HEPOF2SR") Else .HEPOF2SR = " "              '測定位置_領
        If IsNull(rs("HEPOF2NS")) = False Then .HEPOF2NS = rs("HEPOF2NS") Else .HEPOF2NS = " "              '熱処理法
        If IsNull(rs("HEPOF2SZ")) = False Then .HEPOF2SZ = rs("HEPOF2SZ") Else .HEPOF2SZ = " "              '測定条件
        If IsNull(rs("HEPOSF2PTK")) = False Then .HEPOSF2PTK = rs("HEPOSF2PTK") Else .HEPOSF2PTK = " "      'パターン区分
        If IsNull(rs("HEPOF2ET")) = False Then .HEPOF2ET = rs("HEPOF2ET") Else .HEPOF2ET = 0                '選択ET代
        If IsNull(rs("HEPOF2KN")) = False Then .HEPOF2KN = rs("HEPOF2KN") Else .HEPOF2KN = " "              '測定位置_抜
        'OSF3E
        If IsNull(rs("HEPOF3HS")) = False Then .HEPOF3HS = rs("HEPOF3HS") Else .HEPOF3HS = " "              '保証方法_処
        If IsNull(rs("HEPOF3SH")) = False Then .HEPOF3SH = rs("HEPOF3SH") Else .HEPOF3SH = " "              '測定位置_方
        If IsNull(rs("HEPOF3ST")) = False Then .HEPOF3ST = rs("HEPOF3ST") Else .HEPOF3ST = " "              '測定位置_点
        If IsNull(rs("HEPOF3SR")) = False Then .HEPOF3SR = rs("HEPOF3SR") Else .HEPOF3SR = " "              '測定位置_領
        If IsNull(rs("HEPOF3NS")) = False Then .HEPOF3NS = rs("HEPOF3NS") Else .HEPOF3NS = " "              '熱処理法
        If IsNull(rs("HEPOF3SZ")) = False Then .HEPOF3SZ = rs("HEPOF3SZ") Else .HEPOF3SZ = " "              '測定条件
        If IsNull(rs("HEPOSF3PTK")) = False Then .HEPOSF3PTK = rs("HEPOSF3PTK") Else .HEPOSF3PTK = " "      'パターン区分
        If IsNull(rs("HEPOF3ET")) = False Then .HEPOF3ET = rs("HEPOF3ET") Else .HEPOF3ET = 0                '選択ET代
        If IsNull(rs("HEPOF3KN")) = False Then .HEPOF3KN = rs("HEPOF3KN") Else .HEPOF3KN = " "              '測定位置_抜
        'BMD1E
        If IsNull(rs("HEPBM1HS")) = False Then .HEPBM1HS = rs("HEPBM1HS") Else .HEPBM1HS = " "              '保証方法_処
        If IsNull(rs("HEPBM1SH")) = False Then .HEPBM1SH = rs("HEPBM1SH") Else .HEPBM1SH = " "              '測定位置_方
        If IsNull(rs("HEPBM1ST")) = False Then .HEPBM1ST = rs("HEPBM1ST") Else .HEPBM1ST = " "              '測定位置_点
        If IsNull(rs("HEPBM1SR")) = False Then .HEPBM1SR = rs("HEPBM1SR") Else .HEPBM1SR = " "              '測定位置_領
        If IsNull(rs("HEPBM1NS")) = False Then .HEPBM1NS = rs("HEPBM1NS") Else .HEPBM1NS = " "              '熱処理法
        If IsNull(rs("HEPBM1SZ")) = False Then .HEPBM1SZ = rs("HEPBM1SZ") Else .HEPBM1SZ = " "              '測定条件
        If IsNull(rs("HEPBM1ET")) = False Then .HEPBM1ET = rs("HEPBM1ET") Else .HEPBM1ET = 0                '選択ET代
        If IsNull(rs("HEPBM1KN")) = False Then .HEPBM1KN = rs("HEPBM1KN") Else .HEPBM1KN = " "              '測定位置_抜
        'BMD2E
        If IsNull(rs("HEPBM2HS")) = False Then .HEPBM2HS = rs("HEPBM2HS") Else .HEPBM2HS = " "              '保証方法_処
        If IsNull(rs("HEPBM2SH")) = False Then .HEPBM2SH = rs("HEPBM2SH") Else .HEPBM2SH = " "              '測定位置_方
        If IsNull(rs("HEPBM2ST")) = False Then .HEPBM2ST = rs("HEPBM2ST") Else .HEPBM2ST = " "              '測定位置_点
        If IsNull(rs("HEPBM2SR")) = False Then .HEPBM2SR = rs("HEPBM2SR") Else .HEPBM2SR = " "              '測定位置_領
        If IsNull(rs("HEPBM2NS")) = False Then .HEPBM2NS = rs("HEPBM2NS") Else .HEPBM2NS = " "              '熱処理法
        If IsNull(rs("HEPBM2SZ")) = False Then .HEPBM2SZ = rs("HEPBM2SZ") Else .HEPBM2SZ = " "              '測定条件
        If IsNull(rs("HEPBM2ET")) = False Then .HEPBM2ET = rs("HEPBM2ET") Else .HEPBM2ET = 0                '選択ET代
        If IsNull(rs("HEPBM2KN")) = False Then .HEPBM2KN = rs("HEPBM2KN") Else .HEPBM2KN = " "              '測定位置_抜
        'BMD3E
        If IsNull(rs("HEPBM3HS")) = False Then .HEPBM3HS = rs("HEPBM3HS") Else .HEPBM3HS = " "              '保証方法_処
        If IsNull(rs("HEPBM3SH")) = False Then .HEPBM3SH = rs("HEPBM3SH") Else .HEPBM3SH = " "              '測定位置_方
        If IsNull(rs("HEPBM3ST")) = False Then .HEPBM3ST = rs("HEPBM3ST") Else .HEPBM3ST = " "              '測定位置_点
        If IsNull(rs("HEPBM3SR")) = False Then .HEPBM3SR = rs("HEPBM3SR") Else .HEPBM3SR = " "              '測定位置_領
        If IsNull(rs("HEPBM3NS")) = False Then .HEPBM3NS = rs("HEPBM3NS") Else .HEPBM3NS = " "              '熱処理法
        If IsNull(rs("HEPBM3SZ")) = False Then .HEPBM3SZ = rs("HEPBM3SZ") Else .HEPBM3SZ = " "              '測定条件
        If IsNull(rs("HEPBM3ET")) = False Then .HEPBM3ET = rs("HEPBM3ET") Else .HEPBM3ET = 0                '選択ET代
        If IsNull(rs("HEPBM3KN")) = False Then .HEPBM3KN = rs("HEPBM3KN") Else .HEPBM3KN = " "              '測定位置_抜
        'エピAN温度
        If IsNull(rs("HEPANTNP")) = False Then .HEPANTNP = rs("HEPANTNP") Else .HEPANTNP = 0                'AN温度
        'エピ厚
        If IsNull(rs("HEPACEN")) = False Then .HEPACEN = rs("HEPACEN") Else .HEPACEN = 0                    'エピ厚
    End With
    
    Set rs = Nothing
    
    '------------------------------------------ 指示取得 ------------------------------------------------------
    On Error GoTo Apl_down
    'OSF1E
    sErr_Msg = "1-9 OSF1Eﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "19", "O1E", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_9_1
    tbl_chk1_9_1(0).HOSYOU = tbl_chk1_9(0).HEPOF1HS
    tbl_chk1_9_1(1).HOSYOU = tbl_chk1_9(1).HEPOF1HS
    tbl_chk1_9_1(0).SOKU_HOU = tbl_chk1_9(0).HEPOF1SH
    tbl_chk1_9_1(1).SOKU_HOU = tbl_chk1_9(1).HEPOF1SH
    tbl_chk1_9_1(0).SOKU_TEN = tbl_chk1_9(0).HEPOF1ST
    tbl_chk1_9_1(1).SOKU_TEN = tbl_chk1_9(1).HEPOF1ST
    tbl_chk1_9_1(0).SOKU_RYOU = tbl_chk1_9(0).HEPOF1SR
    tbl_chk1_9_1(1).SOKU_RYOU = tbl_chk1_9(1).HEPOF1SR
    tbl_chk1_9_1(0).NETSU = tbl_chk1_9(0).HEPOF1NS
    tbl_chk1_9_1(1).NETSU = tbl_chk1_9(1).HEPOF1NS
    tbl_chk1_9_1(0).JOUKEN = tbl_chk1_9(0).HEPOF1SZ
    tbl_chk1_9_1(1).JOUKEN = tbl_chk1_9(1).HEPOF1SZ
    tbl_chk1_9_1(0).ET = tbl_chk1_9(0).HEPOF1ET
    tbl_chk1_9_1(1).ET = tbl_chk1_9(1).HEPOF1ET
    tbl_chk1_9_1(0).PATTERN = tbl_chk1_9(0).HEPOSF1PTK
    tbl_chk1_9_1(1).PATTERN = tbl_chk1_9(1).HEPOSF1PTK
    tbl_chk1_9_1(0).KENH_NUKI = tbl_chk1_9(0).HEPOF1KN
    tbl_chk1_9_1(1).KENH_NUKI = tbl_chk1_9(1).HEPOF1KN
    tbl_chk1_9_1(0).ANTMP = tbl_chk1_9(0).HEPANTNP
    tbl_chk1_9_1(1).ANTMP = tbl_chk1_9(1).HEPANTNP
    tbl_chk1_9_1(0).EPATU = tbl_chk1_9(0).HEPACEN
    tbl_chk1_9_1(1).EPATU = tbl_chk1_9(1).HEPACEN
    RET = funChkFurikae1_9_1(sResult, tbl_chk1_9_1(), iErr_Code, sErr_Msg, "CHECK1-9,OSF1E")
    If RET <> 0 Then
        funChkFurikae1_9 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00082"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'OSF2E
    sErr_Msg = "1-9 OSF2Eﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "19", "O2E", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_9_1
    tbl_chk1_9_1(0).HOSYOU = tbl_chk1_9(0).HEPOF2HS
    tbl_chk1_9_1(1).HOSYOU = tbl_chk1_9(1).HEPOF2HS
    tbl_chk1_9_1(0).SOKU_HOU = tbl_chk1_9(0).HEPOF2SH
    tbl_chk1_9_1(1).SOKU_HOU = tbl_chk1_9(1).HEPOF2SH
    tbl_chk1_9_1(0).SOKU_TEN = tbl_chk1_9(0).HEPOF2ST
    tbl_chk1_9_1(1).SOKU_TEN = tbl_chk1_9(1).HEPOF2ST
    tbl_chk1_9_1(0).SOKU_RYOU = tbl_chk1_9(0).HEPOF2SR
    tbl_chk1_9_1(1).SOKU_RYOU = tbl_chk1_9(1).HEPOF2SR
    tbl_chk1_9_1(0).NETSU = tbl_chk1_9(0).HEPOF2NS
    tbl_chk1_9_1(1).NETSU = tbl_chk1_9(1).HEPOF2NS
    tbl_chk1_9_1(0).JOUKEN = tbl_chk1_9(0).HEPOF2SZ
    tbl_chk1_9_1(1).JOUKEN = tbl_chk1_9(1).HEPOF2SZ
    tbl_chk1_9_1(0).ET = tbl_chk1_9(0).HEPOF2ET
    tbl_chk1_9_1(1).ET = tbl_chk1_9(1).HEPOF2ET
    tbl_chk1_9_1(0).PATTERN = tbl_chk1_9(0).HEPOSF2PTK
    tbl_chk1_9_1(1).PATTERN = tbl_chk1_9(1).HEPOSF2PTK
    tbl_chk1_9_1(0).KENH_NUKI = tbl_chk1_9(0).HEPOF2KN
    tbl_chk1_9_1(1).KENH_NUKI = tbl_chk1_9(1).HEPOF2KN
    tbl_chk1_9_1(0).ANTMP = tbl_chk1_9(0).HEPANTNP
    tbl_chk1_9_1(1).ANTMP = tbl_chk1_9(1).HEPANTNP
    tbl_chk1_9_1(0).EPATU = tbl_chk1_9(0).HEPACEN
    tbl_chk1_9_1(1).EPATU = tbl_chk1_9(1).HEPACEN
    RET = funChkFurikae1_9_1(sResult, tbl_chk1_9_1(), iErr_Code, sErr_Msg, "CHECK1-9,OSF2E")
    If RET <> 0 Then
        funChkFurikae1_9 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00083"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'OSF3E
    sErr_Msg = "1-9 OSF3Eﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "19", "O3E", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_9_1
    tbl_chk1_9_1(0).HOSYOU = tbl_chk1_9(0).HEPOF3HS
    tbl_chk1_9_1(1).HOSYOU = tbl_chk1_9(1).HEPOF3HS
    tbl_chk1_9_1(0).SOKU_HOU = tbl_chk1_9(0).HEPOF3SH
    tbl_chk1_9_1(1).SOKU_HOU = tbl_chk1_9(1).HEPOF3SH
    tbl_chk1_9_1(0).SOKU_TEN = tbl_chk1_9(0).HEPOF3ST
    tbl_chk1_9_1(1).SOKU_TEN = tbl_chk1_9(1).HEPOF3ST
    tbl_chk1_9_1(0).SOKU_RYOU = tbl_chk1_9(0).HEPOF3SR
    tbl_chk1_9_1(1).SOKU_RYOU = tbl_chk1_9(1).HEPOF3SR
    tbl_chk1_9_1(0).NETSU = tbl_chk1_9(0).HEPOF3NS
    tbl_chk1_9_1(1).NETSU = tbl_chk1_9(1).HEPOF3NS
    tbl_chk1_9_1(0).JOUKEN = tbl_chk1_9(0).HEPOF3SZ
    tbl_chk1_9_1(1).JOUKEN = tbl_chk1_9(1).HEPOF3SZ
    tbl_chk1_9_1(0).ET = tbl_chk1_9(0).HEPOF3ET
    tbl_chk1_9_1(1).ET = tbl_chk1_9(1).HEPOF3ET
    tbl_chk1_9_1(0).PATTERN = tbl_chk1_9(0).HEPOSF3PTK
    tbl_chk1_9_1(1).PATTERN = tbl_chk1_9(1).HEPOSF3PTK
    tbl_chk1_9_1(0).KENH_NUKI = tbl_chk1_9(0).HEPOF3KN
    tbl_chk1_9_1(1).KENH_NUKI = tbl_chk1_9(1).HEPOF3KN
    tbl_chk1_9_1(0).ANTMP = tbl_chk1_9(0).HEPANTNP
    tbl_chk1_9_1(1).ANTMP = tbl_chk1_9(1).HEPANTNP
    tbl_chk1_9_1(0).EPATU = tbl_chk1_9(0).HEPACEN
    tbl_chk1_9_1(1).EPATU = tbl_chk1_9(1).HEPACEN
    RET = funChkFurikae1_9_1(sResult, tbl_chk1_9_1(), iErr_Code, sErr_Msg, "CHECK1-9,OSF3E")
    If RET <> 0 Then
        funChkFurikae1_9 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00084"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'BMD1E
    sErr_Msg = "1-9 BMD1Eﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "19", "B1E", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_9_1
    tbl_chk1_9_1(0).HOSYOU = tbl_chk1_9(0).HEPBM1HS
    tbl_chk1_9_1(1).HOSYOU = tbl_chk1_9(1).HEPBM1HS
    tbl_chk1_9_1(0).SOKU_HOU = tbl_chk1_9(0).HEPBM1SH
    tbl_chk1_9_1(1).SOKU_HOU = tbl_chk1_9(1).HEPBM1SH
    tbl_chk1_9_1(0).SOKU_TEN = tbl_chk1_9(0).HEPBM1ST
    tbl_chk1_9_1(1).SOKU_TEN = tbl_chk1_9(1).HEPBM1ST
    tbl_chk1_9_1(0).SOKU_RYOU = tbl_chk1_9(0).HEPBM1SR
    tbl_chk1_9_1(1).SOKU_RYOU = tbl_chk1_9(1).HEPBM1SR
    tbl_chk1_9_1(0).NETSU = tbl_chk1_9(0).HEPBM1NS
    tbl_chk1_9_1(1).NETSU = tbl_chk1_9(1).HEPBM1NS
    tbl_chk1_9_1(0).JOUKEN = tbl_chk1_9(0).HEPBM1SZ
    tbl_chk1_9_1(1).JOUKEN = tbl_chk1_9(1).HEPBM1SZ
    tbl_chk1_9_1(0).ET = tbl_chk1_9(0).HEPBM1ET
    tbl_chk1_9_1(1).ET = tbl_chk1_9(1).HEPBM1ET
    tbl_chk1_9_1(0).KENH_NUKI = tbl_chk1_9(0).HEPBM1KN
    tbl_chk1_9_1(1).KENH_NUKI = tbl_chk1_9(1).HEPBM1KN
    tbl_chk1_9_1(0).ANTMP = tbl_chk1_9(0).HEPANTNP
    tbl_chk1_9_1(1).ANTMP = tbl_chk1_9(1).HEPANTNP
    tbl_chk1_9_1(0).EPATU = tbl_chk1_9(0).HEPACEN
    tbl_chk1_9_1(1).EPATU = tbl_chk1_9(1).HEPACEN
    RET = funChkFurikae1_9_1(sResult, tbl_chk1_9_1(), iErr_Code, sErr_Msg, "CHECK1-9,BMD1E")
    If RET <> 0 Then
        funChkFurikae1_9 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00085"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'BMD2E
    sErr_Msg = "1-9 BMD2Eﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "19", "B2E", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_9_1
    tbl_chk1_9_1(0).HOSYOU = tbl_chk1_9(0).HEPBM2HS
    tbl_chk1_9_1(1).HOSYOU = tbl_chk1_9(1).HEPBM2HS
    tbl_chk1_9_1(0).SOKU_HOU = tbl_chk1_9(0).HEPBM2SH
    tbl_chk1_9_1(1).SOKU_HOU = tbl_chk1_9(1).HEPBM2SH
    tbl_chk1_9_1(0).SOKU_TEN = tbl_chk1_9(0).HEPBM2ST
    tbl_chk1_9_1(1).SOKU_TEN = tbl_chk1_9(1).HEPBM2ST
    tbl_chk1_9_1(0).SOKU_RYOU = tbl_chk1_9(0).HEPBM2SR
    tbl_chk1_9_1(1).SOKU_RYOU = tbl_chk1_9(1).HEPBM2SR
    tbl_chk1_9_1(0).NETSU = tbl_chk1_9(0).HEPBM2NS
    tbl_chk1_9_1(1).NETSU = tbl_chk1_9(1).HEPBM2NS
    tbl_chk1_9_1(0).JOUKEN = tbl_chk1_9(0).HEPBM2SZ
    tbl_chk1_9_1(1).JOUKEN = tbl_chk1_9(1).HEPBM2SZ
    tbl_chk1_9_1(0).ET = tbl_chk1_9(0).HEPBM2ET
    tbl_chk1_9_1(1).ET = tbl_chk1_9(1).HEPBM2ET
    tbl_chk1_9_1(0).KENH_NUKI = tbl_chk1_9(0).HEPBM2KN
    tbl_chk1_9_1(1).KENH_NUKI = tbl_chk1_9(1).HEPBM2KN
    tbl_chk1_9_1(0).ANTMP = tbl_chk1_9(0).HEPANTNP
    tbl_chk1_9_1(1).ANTMP = tbl_chk1_9(1).HEPANTNP
    tbl_chk1_9_1(0).EPATU = tbl_chk1_9(0).HEPACEN
    tbl_chk1_9_1(1).EPATU = tbl_chk1_9(1).HEPACEN
    RET = funChkFurikae1_9_1(sResult, tbl_chk1_9_1(), iErr_Code, sErr_Msg, "CHECK1-9,BMD2E")
    If RET <> 0 Then
        funChkFurikae1_9 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00086"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'BMD3E
    sErr_Msg = "1-9 BMD3Eﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "19", "B3E", 0, " ", sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→指示取得"
        GoTo CodeDBGet_Error
    End If
    Erase tbl_chk1_9_1
    tbl_chk1_9_1(0).HOSYOU = tbl_chk1_9(0).HEPBM3HS
    tbl_chk1_9_1(1).HOSYOU = tbl_chk1_9(1).HEPBM3HS
    tbl_chk1_9_1(0).SOKU_HOU = tbl_chk1_9(0).HEPBM3SH
    tbl_chk1_9_1(1).SOKU_HOU = tbl_chk1_9(1).HEPBM3SH
    tbl_chk1_9_1(0).SOKU_TEN = tbl_chk1_9(0).HEPBM3ST
    tbl_chk1_9_1(1).SOKU_TEN = tbl_chk1_9(1).HEPBM3ST
    tbl_chk1_9_1(0).SOKU_RYOU = tbl_chk1_9(0).HEPBM3SR
    tbl_chk1_9_1(1).SOKU_RYOU = tbl_chk1_9(1).HEPBM3SR
    tbl_chk1_9_1(0).NETSU = tbl_chk1_9(0).HEPBM3NS
    tbl_chk1_9_1(1).NETSU = tbl_chk1_9(1).HEPBM3NS
    tbl_chk1_9_1(0).JOUKEN = tbl_chk1_9(0).HEPBM3SZ
    tbl_chk1_9_1(1).JOUKEN = tbl_chk1_9(1).HEPBM3SZ
    tbl_chk1_9_1(0).ET = tbl_chk1_9(0).HEPBM3ET
    tbl_chk1_9_1(1).ET = tbl_chk1_9(1).HEPBM3ET
    tbl_chk1_9_1(0).KENH_NUKI = tbl_chk1_9(0).HEPBM3KN
    tbl_chk1_9_1(1).KENH_NUKI = tbl_chk1_9(1).HEPBM3KN
    tbl_chk1_9_1(0).ANTMP = tbl_chk1_9(0).HEPANTNP
    tbl_chk1_9_1(1).ANTMP = tbl_chk1_9(1).HEPANTNP
    tbl_chk1_9_1(0).EPATU = tbl_chk1_9(0).HEPACEN
    tbl_chk1_9_1(1).EPATU = tbl_chk1_9(1).HEPACEN
    RET = funChkFurikae1_9_1(sResult, tbl_chk1_9_1(), iErr_Code, sErr_Msg, "CHECK1-9,BMD3E")
    If RET <> 0 Then
        funChkFurikae1_9 = RET
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00087"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If

'------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae1_9 = 0 Then
        funChkFurikae1_9 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae1_9 = -4
    GoTo Apl_Exit
    
CodeDBGet_Error:
    Set rs = Nothing
    If funChkFurikae1_9 = 0 Then
        funChkFurikae1_9 = -5
    End If
    GoTo Apl_Exit

End Function

'------------------------------------------------
' 先行評価項目仕様詳細チェック
'------------------------------------------------

'概要      :指定されたﾁｪｯｸ内容詳細に基づき、該当する仕様値のチェックを行なう。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型                 :説明
'          :sChkCode        ,I  ,String             :チェック内容詳細
'          :tbl_chk1_9_1()  ,I  ,typ_chk1_9_1       :仕様値構造体配列
'          :iErr_Code       ,O  ,Integer            :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String             :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :sAdd_Msg        ,I  ,String             :添付ｴﾗｰﾒｯｾｰｼﾞ
'          :戻り値          ,O  ,Integer            :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :2006/08/15 新規作成 エピ先行評価追加対応 SMP)kondoh

Public Function funChkFurikae1_9_1(sChkCode As String, tbl_chk1_9_1() As typ_chk1_9_1, _
                                   iErr_Code As Integer, sErr_Msg As String, sAdd_Msg As String) As Integer

    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim wHOSYOU_0   As String       '保証方法＿対象
    Dim wHOSYOU_1   As String       '保証方法＿対象
    Dim iCnt        As Integer
    Dim sNum(2)     As String
    Dim lsCodeList() As String       'コードDBのコードのリスト
    Dim liNumCnt    As Integer

    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae1_9_1 = 0
    iErr_Code = 0
    '------------------------------------------ 保証方法チェック ------------------------------------------------------
    If tbl_chk1_9_1(1).HOSYOU <> "H" And tbl_chk1_9_1(1).HOSYOU <> "S" Then GoTo Apl_Exit
    
    '------------------------------------------ 各種チェック ------------------------------------------------------
    '保証方法＿対象
    sErr_Msg = "保証方法_対象ﾁｪｯｸ"
    If Mid(sChkCode, 1, 1) = "2" Then
        '振替元と振替先が等しければ振替ＯＫ
        If tbl_chk1_9_1(0).HOSYOU <> tbl_chk1_9_1(1).HOSYOU Then
            
            wHOSYOU_0 = tbl_chk1_9_1(0).HOSYOU
            If tbl_chk1_9_1(0).HOSYOU <> "H" And tbl_chk1_9_1(0).HOSYOU <> "S" Then wHOSYOU_0 = "-"
            wHOSYOU_1 = tbl_chk1_9_1(1).HOSYOU
            If tbl_chk1_9_1(1).HOSYOU <> "H" And tbl_chk1_9_1(1).HOSYOU <> "S" Then wHOSYOU_1 = "-"
            
            'マトリクス取得
            sResult = ""
            RET = funCodeDBGet("SB", "SH", wHOSYOU_0, 1, wHOSYOU_1, sResult)
            If RET <> 0 Then
                sErr_Msg = sAdd_Msg & sErr_Msg & "→元:" & tbl_chk1_9_1(0).HOSYOU & ", 先:" & tbl_chk1_9_1(1).HOSYOU
                GoTo CodeDBGet_Error
            End If
            If sResult = 0 Then
                funChkFurikae1_9_1 = 1
                iErr_Code = 1901
                GoTo Apl_Exit
            End If
        End If
    End If
''    '下限
''    sErr_Msg = "下限ﾁｪｯｸ"
''    If Mid(sChkCode, 2, 1) = "1" Then
''        If tbl_chk1_9_1(0).MIN_LIMIT <> tbl_chk1_9_1(1).MIN_LIMIT Then
''            funChkFurikae1_9_1 = 1
''            iErr_Code = 1902
''            GoTo Apl_Exit
''        End If
''    End If
''    '上限
''    sErr_Msg = "上限ﾁｪｯｸ"
''    If Mid(sChkCode, 3, 1) = "1" Then
''        If tbl_chk1_9_1(0).MAX_LIMIT <> tbl_chk1_9_1(1).MAX_LIMIT Then
''            funChkFurikae1_9_1 = 1
''            iErr_Code = 1903
''            GoTo Apl_Exit
''        End If
''    End If
    '測定位置＿方
    sErr_Msg = "測定位置_方ﾁｪｯｸ"
    If Mid(sChkCode, 4, 1) = "1" Then
        If Trim$(tbl_chk1_9_1(0).SOKU_HOU) <> Trim$(tbl_chk1_9_1(1).SOKU_HOU) Then
            funChkFurikae1_9_1 = 1
            iErr_Code = 1904
            GoTo Apl_Exit
        End If
    End If
    '測定位置＿点
    sErr_Msg = "測定位置_点ﾁｪｯｸ"
    If Mid(sChkCode, 5, 1) = "1" Then
        If Trim$(tbl_chk1_9_1(0).SOKU_TEN) <> Trim$(tbl_chk1_9_1(1).SOKU_TEN) Then
            funChkFurikae1_9_1 = 1
            iErr_Code = 1905
            GoTo Apl_Exit
        End If
    End If
''    '測定位置＿位
''    sErr_Msg = "測定位置_位ﾁｪｯｸ"
''    If Mid(sChkCode, 6, 1) = "2" Then
''       'マトリクス取得
''        sResult = ""
''        RET = funCodeDBGet("SB", "OI", tbl_chk1_9_1(0).SOKU_ICHI, 1, tbl_chk1_9_1(1).SOKU_ICHI, sResult)
''        If RET <> 0 Then
''            sErr_Msg = sAdd_Msg & sErr_Msg & "→元:" & tbl_chk1_9_1(0).SOKU_ICHI & ", 先:" & tbl_chk1_9_1(1).SOKU_ICHI
''            GoTo CodeDBGet_Error
''        End If
''        If sResult = 0 Then
''            funChkFurikae1_9_1 = 1
''            iErr_Code = 1906
''            GoTo Apl_Exit
''        End If
''    End If
''    '測定位置＿位
''    If Mid(sChkCode, 6, 1) = "1" Then
''        If Trim$(tbl_chk1_9_1(0).SOKU_ICHI) <> Trim$(tbl_chk1_9_1(1).SOKU_ICHI) Then
''            funChkFurikae1_9_1 = 1
''            iErr_Code = 1906
''            GoTo Apl_Exit
''        End If
''    End If
    
    '測定位置＿領
    sErr_Msg = "測定位置_領ﾁｪｯｸ"
    If Mid(sChkCode, 7, 1) = "1" Then
        If Trim$(tbl_chk1_9_1(0).SOKU_RYOU) <> Trim$(tbl_chk1_9_1(1).SOKU_RYOU) Then
            funChkFurikae1_9_1 = 1
            iErr_Code = 1907
            GoTo Apl_Exit
        End If
    End If
''    '検査有無
''    sErr_Msg = "検査有無ﾁｪｯｸ"
''    If Mid(sChkCode, 8, 1) = "1" Then
''        If Trim$(tbl_chk1_9_1(0).UMU) <> Trim$(tbl_chk1_9_1(1).UMU) Then
''            funChkFurikae1_9_1 = 1
''            iErr_Code = 1908
''            GoTo Apl_Exit
''        End If
''    End If
    '熱処理法
    sErr_Msg = "熱処理法ﾁｪｯｸ"
    If Mid(sChkCode, 9, 1) = "1" Then
        If Trim$(tbl_chk1_9_1(0).NETSU) <> Trim$(tbl_chk1_9_1(1).NETSU) Then
            funChkFurikae1_9_1 = 1
            iErr_Code = 1909
            GoTo Apl_Exit
        End If
    End If
    '測定条件
    sErr_Msg = "測定条件ﾁｪｯｸ"
    If Mid(sChkCode, 10, 1) = "1" Then
        If Trim$(tbl_chk1_9_1(0).JOUKEN) <> Trim$(tbl_chk1_9_1(1).JOUKEN) Then
            funChkFurikae1_9_1 = 1
            iErr_Code = 1910
            GoTo Apl_Exit
        End If
    End If
    '選択ＥＴ代
    sErr_Msg = "選択ET代ﾁｪｯｸ"
    If Mid(sChkCode, 11, 1) = "1" Then
        If tbl_chk1_9_1(0).ET <> tbl_chk1_9_1(1).ET Then
            funChkFurikae1_9_1 = 1
            iErr_Code = 1911
            GoTo Apl_Exit
        End If
    End If
''    '検査方法
''    sErr_Msg = "検査方法ﾁｪｯｸ"
''    If Mid(sChkCode, 12, 1) = "1" Then
''        If Trim$(tbl_chk1_9_1(0).KENSA) <> Trim$(tbl_chk1_9_1(1).KENSA) Then
''            funChkFurikae1_9_1 = 1
''            iErr_Code = 1912
''            GoTo Apl_Exit
''        End If
''    End If
    'パターン区分
    sErr_Msg = "ﾊﾟﾀｰﾝ区分ﾁｪｯｸ"
    If Mid(sChkCode, 13, 1) = "2" Then
        'マトリクス取得
        sResult = ""
        RET = funCodeDBGet("SB", "OS", tbl_chk1_9_1(0).PATTERN, 1, tbl_chk1_9_1(1).PATTERN, sResult)
        If RET <> 0 Then
            sErr_Msg = sAdd_Msg & sErr_Msg & "→元:" & tbl_chk1_9_1(0).PATTERN & ", 先:" & tbl_chk1_9_1(1).PATTERN
            GoTo CodeDBGet_Error
        End If
        If sResult = 0 Then
            funChkFurikae1_9_1 = 1
            iErr_Code = 1913
            GoTo Apl_Exit
        End If
    End If
    ''検査頻度＿抜
    sErr_Msg = "検査頻度_抜ﾁｪｯｸ"
    If Mid(sChkCode, 14, 1) = "2" Then
        'マトリクス取得
        sResult = ""
        
        For iCnt = 0 To 1
            Select Case tbl_chk1_9_1(iCnt).KENH_NUKI
            Case "3", "4", "6"
                sNum(iCnt) = tbl_chk1_9_1(iCnt).KENH_NUKI
            Case Else
                sNum(iCnt) = "ETC"
            End Select
        Next
        
        RET = funCodeDBGet("SB", "HO", sNum(0), 1, sNum(1), sResult)
        
        If RET <> 0 Then
            sErr_Msg = sAdd_Msg & sErr_Msg & "→元:" & tbl_chk1_9_1(0).KENH_NUKI & ", 先:" & tbl_chk1_9_1(1).KENH_NUKI
            GoTo CodeDBGet_Error
        End If
        If sResult = 0 Then
            funChkFurikae1_9_1 = 1
            iErr_Code = 1914
            GoTo Apl_Exit
        End If
    End If
    ''AN温度
    sErr_Msg = "AN温度ﾁｪｯｸ"
    If Mid(sChkCode, 15, 1) = "2" Then
        'マトリクス取得
        sResult = ""
        
        For iCnt = 0 To 1
            sNum(iCnt) = CStr(Trim(tbl_chk1_9_1(iCnt).ANTMP))
        Next
        '' コードマスタのコードの一覧を取得
        RET = funCodeDBGetCodeList("SB", "AE", lsCodeList)
        If RET <> 0 Then
            sErr_Msg = sAdd_Msg & sErr_Msg & "→元:" & tbl_chk1_9_1(0).ANTMP & ", 先:" & tbl_chk1_9_1(1).ANTMP
            GoTo CodeDBGet_Error
        End If
        ''コードマスタに登録されていないコードはスペースに変換する
        For liNumCnt = 0 To 1
            RET = 0
            For iCnt = 1 To UBound(lsCodeList)
                If Trim(lsCodeList(iCnt)) = Trim(sNum(liNumCnt)) Then
                    RET = 1
                    Exit For
                End If
            Next iCnt
            If RET = 0 Then
                sNum(liNumCnt) = "     "
            End If
        Next liNumCnt
        
        RET = funCodeDBGet("SB", "AE", sNum(1), 1, sNum(0), sResult)
        
        If RET <> 0 Then
            sErr_Msg = sAdd_Msg & sErr_Msg & "→元:" & tbl_chk1_9_1(0).ANTMP & ", 先:" & tbl_chk1_9_1(1).ANTMP
            GoTo CodeDBGet_Error
        End If
        If sResult = 0 Then
            funChkFurikae1_9_1 = 1
            iErr_Code = 1915
            ''メッセージの中に温度を入れたいので、エラーメッセージはここで作成する
            sAdd_Msg = sAdd_Msg & "のAN温度が振替不可能です。(" & tbl_chk1_9_1(0).ANTMP & "℃ → " & tbl_chk1_9_1(1).ANTMP & "℃)"
            GoTo Apl_Exit
        End If
    End If
    'エピ厚 まだまだ(等号が入るかPending中)
    sErr_Msg = "選択ET代ﾁｪｯｸ"
    If Mid(sChkCode, 16, 1) = "2" Then
        If tbl_chk1_9_1(0).EPATU > tbl_chk1_9_1(1).EPATU Then
            funChkFurikae1_9_1 = 1
            iErr_Code = 1916
            sAdd_Msg = sAdd_Msg & "のＥ１厚中心が振替不可能です。(" & tbl_chk1_9_1(0).EPATU & " → " & tbl_chk1_9_1(1).EPATU & ")"
            GoTo Apl_Exit
        End If
    End If

    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Select Case iErr_Code
        Case 1901
            sErr_Msg = sAdd_Msg & "の保証方法が不一致の為、振替できません。"
        Case 1902
            sErr_Msg = sAdd_Msg & "の下限が不一致の為、振替できません。"
        Case 1903
            sErr_Msg = sAdd_Msg & "の上限が不一致の為、振替できません。"
        Case 1904
            sErr_Msg = sAdd_Msg & "の測定位置＿方が不一致の為、振替できません。"
        Case 1905
            sErr_Msg = sAdd_Msg & "の測定位置＿点が不一致の為、振替できません。"
        Case 1906
            If Mid(sChkCode, 6, 1) = "2" Then
                sErr_Msg = sAdd_Msg & "の測定位置＿位が振替不可能です。"
            Else
                sErr_Msg = sAdd_Msg & "の測定位置＿位が不一致の為、振替できません。"
            End If
        Case 1907
            sErr_Msg = sAdd_Msg & "の測定位置＿領が不一致の為、振替できません。"
        Case 1908
            sErr_Msg = sAdd_Msg & "の検査有無が不一致の為、振替できません。"
        Case 1909
            sErr_Msg = sAdd_Msg & "の熱処理法が不一致の為、振替できません。"
        Case 1910
            sErr_Msg = sAdd_Msg & "の測定条件が不一致の為、振替できません。"
        Case 1911
            sErr_Msg = sAdd_Msg & "の選択ＥＴ代が不一致の為、振替できません。"
        Case 1912
            sErr_Msg = sAdd_Msg & "の検査方法が不一致の為、振替できません。"
        Case 1913
            sErr_Msg = sAdd_Msg & "のパターン区分が振替不可能です。"
        Case 1914
            sErr_Msg = sAdd_Msg & "の検査頻度＿抜が振替不可能です。"
        Case 1915
            sErr_Msg = sAdd_Msg
        Case 1916
            sErr_Msg = sAdd_Msg
    End Select
    
    Exit Function
    
Apl_down:
    funChkFurikae1_9_1 = -4
    GoTo Apl_Exit
    
CodeDBGet_Error:
    If funChkFurikae1_9_1 = 0 Then
        funChkFurikae1_9_1 = -5
    End If
    GoTo Apl_Exit
    
End Function

'------------------------------------------------
' 振替先と振替元の常識仕様チェック２
'------------------------------------------------

'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sKeyID          ,I  ,String       :ﾌﾞﾛｯｸID、又は、SXL-ID
'          :tOld_Hinban     ,I  ,tFullHinban  :振替元品番(構造体)
'          :tNew_Hinban     ,I  ,tFullHinban  :振替先品番(構造体)
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :06/10/05 ooba

Public Function funChkFurikae1_10(sProccd As String, sBlockId As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                  iErr_Code As Integer, sErr_Msg As String) As Integer


    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim sClass      As String       '区分  ''add 0108
    Dim wXtal       As String                            '2010/04/16 Kameda
    Dim Xsen        As type_DBDRV_scmzc_fcmkc001c_X      '2010/04/16 Kameda
    Dim Xsiyou      As type_DBDRV_scmzc_fcmkc001c_Siyou  '2010/04/16 Kameda
    Dim JUDGXY     As Boolean                            'X線判定用フラグ追加 2010/04/16
    Dim JUDGX      As Boolean                            'X線判定用フラグ追加 2010/04/16
    Dim JUDGY      As Boolean                            'X線判定用フラグ追加 2010/04/16
    Dim cnt        As Integer
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae1_10 = 0
    
    '------------------------------------------ 振替元品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-10 振替元品番仕様取得(" & tOld_Hinban.hinban & Format(tOld_Hinban.mnorevno, "00") & tOld_Hinban.factory & tOld_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E018.HSXCDIR,E018.HSXDOP,E023.HWFCDOP,E018.HSXDPDIR, " & vbCrLf
    sql = sql & "       E018.HSXCSCEN, "    ''2008/11/27 結晶面傾中心チェック緩和(2) ADD By Systech
    sql = sql & "       SUBSTR(E018.MCNO,1,1) MCNO1,SUBSTR(E018.MCNO,4,1) MCNO2,SUBSTR(E018.MCNO,3,1) MCNO3,E036.DCHYUUBU " & vbCrLf
    sql = sql & "FROM   TBCME018 E018,TBCME023 E023,TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E018.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E018.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E018.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E023.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E023.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E023.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E023.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E036.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tOld_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    Erase tbl_chk1_10
    With tbl_chk1_10(0)
        If IsNull(rs("HSXCDIR")) = False Then .HSXCDIR = rs("HSXCDIR") Else .HSXCDIR = " "          ' 結晶面方位
        If IsNull(rs("HSXCSCEN")) = False Then .HSXCSCEN = rs("HSXCSCEN") Else .HSXCSCEN = 0        ' 結晶面傾き中心    ''2008/11/27 結晶面傾中心チェック緩和(2) ADD By Systech
        If IsNull(rs("HSXDOP")) = False Then .HSXDOP = rs("HSXDOP") Else .HSXDOP = " "              ' ドーパント
        If IsNull(rs("HWFCDOP")) = False Then .HWFCDOP = rs("HWFCDOP") Else .HWFCDOP = " "          ' 結晶ドープ
        If IsNull(rs("HSXDPDIR")) = False Then .HSXDPDIR = rs("HSXDPDIR") Else .HSXDPDIR = " "      ' 溝位置方位
        If IsNull(rs("MCNO1")) = False Then .MCNO1 = rs("MCNO1") Else .MCNO1 = " "                  ' 品種
        If IsNull(rs("MCNO2")) = False Then .MCNO2 = rs("MCNO2") Else .MCNO2 = " "                  ' 引上げ速度
        If IsNull(rs("MCNO3")) = False Then .MCNO3 = rs("MCNO3") Else .MCNO3 = " "                  ' HZタイプ
        If IsNull(rs("DCHYUUBU")) = False Then .DCHYUUBU = rs("DCHYUUBU") Else .DCHYUUBU = "2"      ' ドローチューブ
    End With
    
    Set rs = Nothing
    '------------------------------------------ 振替先品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-10 振替先品番仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E018.HSXCDIR,E018.HSXDOP,E023.HWFCDOP,E018.HSXDPDIR, " & vbCrLf
    sql = sql & "       E018.HSXCSCEN, "    ''2008/11/27 結晶面傾中心チェック緩和(2) ADD By Systech
    sql = sql & "       SUBSTR(E018.MCNO,1,1) MCNO1,SUBSTR(E018.MCNO,4,1) MCNO2,SUBSTR(E018.MCNO,3,1) MCNO3,E036.DCHYUUBU " & vbCrLf
    sql = sql & "       ,E036.NDOPHUFLG,E036.CDOPHUFLG " & vbCrLf    '' add 0108
    sql = sql & "       ,E018.HSXCSCEN,E018.HSXCSMIN,E018.HSXCSMAX,E018.HSXCYCEN,E018.HSXCYMIN,E018.HSXCYMAX,E018.HSXCTCEN,E018.HSXCTMIN,E018.HSXCTMAX " & vbCrLf   '2010/04/16 Kameda
    sql = sql & "FROM   TBCME018 E018,TBCME023 E023,TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E018.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E018.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E018.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E023.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E023.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E023.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E023.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E036.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tNew_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    With tbl_chk1_10(1)
        If IsNull(rs("HSXCDIR")) = False Then .HSXCDIR = rs("HSXCDIR") Else .HSXCDIR = " "          ' 結晶面方位
        If IsNull(rs("HSXCSCEN")) = False Then .HSXCSCEN = rs("HSXCSCEN") Else .HSXCSCEN = 0        ' 結晶面傾き中心    ''2008/11/27 結晶面傾中心チェック緩和(2) ADD By Systech
        If IsNull(rs("HSXDOP")) = False Then .HSXDOP = rs("HSXDOP") Else .HSXDOP = " "              ' ドーパント
        If IsNull(rs("HWFCDOP")) = False Then .HWFCDOP = rs("HWFCDOP") Else .HWFCDOP = " "          ' 結晶ドープ
        If IsNull(rs("HSXDPDIR")) = False Then .HSXDPDIR = rs("HSXDPDIR") Else .HSXDPDIR = " "      ' 溝位置方位
        If IsNull(rs("MCNO1")) = False Then .MCNO1 = rs("MCNO1") Else .MCNO1 = " "                  ' 品種
        If IsNull(rs("MCNO2")) = False Then .MCNO2 = rs("MCNO2") Else .MCNO2 = " "                  ' 引上げ速度
        If IsNull(rs("MCNO3")) = False Then .MCNO3 = rs("MCNO3") Else .MCNO3 = " "                  ' HZタイプ
        If IsNull(rs("DCHYUUBU")) = False Then .DCHYUUBU = rs("DCHYUUBU") Else .DCHYUUBU = "2"      ' ドローチューブ
        If IsNull(rs("NDOPHUFLG")) = False Then .NDOPHUFLG = rs("NDOPHUFLG") Else .NDOPHUFLG = " "  ' 窒素ドープ振替可否フラグ '' add 0108
        If IsNull(rs("CDOPHUFLG")) = False Then .CDOPHUFLG = rs("CDOPHUFLG") Else .CDOPHUFLG = " "  ' Cドープ振替可否フラグ '' add 0108
    End With
    With Xsiyou
        .HSXCSCEN = fncNullCheck(rs("HSXCSCEN"))        ' 品ＳＸ面傾き中心    2010/04/16 Kameda
        .HSXCSMIN = fncNullCheck(rs("HSXCSMIN"))        ' 品ＳＸ面傾き下限    2010/04/16 Kameda
        .HSXCSMAX = fncNullCheck(rs("HSXCSMAX"))        ' 品ＳＸ面傾き上限    2010/04/16 Kameda
        .HSXCTCEN = fncNullCheck(rs("HSXCTCEN"))        ' 品ＳＸ面傾き縦中心  2010/04/16 Kameda
        .HSXCTMIN = fncNullCheck(rs("HSXCTMIN"))        ' 品ＳＸ面傾き縦下限  2010/04/16 Kameda
        .HSXCTMAX = fncNullCheck(rs("HSXCTMAX"))        ' 品ＳＸ面傾き縦上限  2010/04/16 Kameda
        .HSXCYCEN = fncNullCheck(rs("HSXCYCEN"))        ' 品ＳＸ面傾き横中心  2010/04/16 Kameda
        .HSXCYMIN = fncNullCheck(rs("HSXCYMIN"))        ' 品ＳＸ面傾き横下限  2010/04/16 Kameda
        .HSXCYMAX = fncNullCheck(rs("HSXCYMAX"))        ' 品ＳＸ面傾き横上限  2010/04/16 Kameda
    End With
    
    Set rs = Nothing
    '------------------------------------------ 指示取得 ------------------------------------------------------
    On Error GoTo Apl_down
    '結晶面方位のチェック
    sErr_Msg = "1-10 結晶面方位ﾁｪｯｸ"
    If Trim$(tbl_chk1_10(0).HSXCDIR) <> Trim$(tbl_chk1_10(1).HSXCDIR) Then
        funChkFurikae1_10 = 1
        iErr_Code = 1001
        sErr_Msg = "CHECK1-10,結晶面方位不一致の為、振替できません。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00003"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
''2008/11/27 結晶面傾中心チェック緩和(2) ADD By Systech Start
    '結晶面傾中心のチェック
    sErr_Msg = "1-10 結晶面傾中心ﾁｪｯｸ"
    If Abs(tbl_chk1_10(0).HSXCSCEN - tbl_chk1_10(1).HSXCSCEN) > 1 Then
        funChkFurikae1_10 = 1
        iErr_Code = 1009
        sErr_Msg = "CHECK1-10,結晶面傾中心不一致の為、振替できません。"
        gsTbcmy028ErrCode = "00004"
        GoTo Apl_Exit
    End If
''2008/11/27 結晶面傾中心チェック緩和(2) ADD By Systech End
    'ドーパントのチェック
    sErr_Msg = "1-10 ﾄﾞｰﾊﾟﾝﾄﾁｪｯｸ"
    If Trim$(tbl_chk1_10(0).HSXDOP) <> Trim$(tbl_chk1_10(1).HSXDOP) Then
        funChkFurikae1_10 = 1
        iErr_Code = 1002
        sErr_Msg = "CHECK1-10,ドーパント不一致の為、振替できません。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00005"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    
    '2010/04/16 結晶面傾中心仕様判断条件の追加 №100087 Kameda    <----- 1-2より移動
    'If left(sProccd, 4) = "CW76" Then
        '面傾中心仕様0.00度品から0.00度品以外への振替を禁止
        sErr_Msg = "1-10 結晶面傾中心ﾁｪｯｸ"
        If Trim$(tbl_chk1_10(0).HSXCSCEN) = 0 Then
            If Trim$(tbl_chk1_10(1).HSXCSCEN) <> 0 Then
                funChkFurikae1_10 = 1
                iErr_Code = 1201
                sErr_Msg = "CHECK1-10,結晶面傾中心不一致の為、振替できません。"
                gsTbcmy028ErrCode = "00004"
                GoTo Apl_Exit
            End If
        End If
        '面傾中心仕様1.00度以下品から0.00度品への振替はＸ線実績が振替先の仕様範囲内
        If Trim$(tbl_chk1_10(0).HSXCSCEN) < 1 And Trim$(tbl_chk1_10(1).HSXCSCEN) = 0 Then
            wXtal = left(sBlockId, 9) & "000"
            sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, XRAYX,XRAYY,XRAYXY, REGDATE "
            sql = sql & "from TBCMJ021 "
            sql = sql & "where CRYNUM = '" & wXtal & "' and "
            sql = sql & "      TRANCNT = (select max(TRANCNT) from TBCMJ021 "
            sql = sql & "                 where CRYNUM = '" & wXtal & "' )"
            
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.EOF Or rs.RecordCount = 0 Then
            Else
                With Xsen
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
                If CrXjudg(Xsiyou, Xsen, JUDGXY, JUDGX, JUDGY) = True Then
                    If JUDGXY = False Then
                        funChkFurikae1_10 = 1
                        iErr_Code = 1201
                        sErr_Msg = "CHECK1-10,結晶面傾中心,Ｘ線実績が範囲外の為、振替できません。"
                        gsTbcmy028ErrCode = "00004"
                        GoTo Apl_Exit
                    End If
                End If
            End If
        End If
    'End If
    '2010/04/16 結晶面傾中心仕様判断条件の追加 END №100087 Kameda
    
    '結晶ドープのチェック
    sErr_Msg = "1-10 結晶ﾄﾞｰﾌﾟﾁｪｯｸ"
'' add start 0108
    '' 区分判断
    sClass = ""
    '' N振替可/C振替可
    If tbl_chk1_10(1).NDOPHUFLG = "0" And tbl_chk1_10(1).CDOPHUFLG = "0" Then
        sClass = "D0"
    '' N振替可/C振替不可
    ElseIf tbl_chk1_10(1).NDOPHUFLG = "0" And tbl_chk1_10(1).CDOPHUFLG <> "0" Then
        sClass = "D1"
    '' N振替不可/C振替可
    ElseIf tbl_chk1_10(1).NDOPHUFLG <> "0" And tbl_chk1_10(1).CDOPHUFLG = "0" Then
        sClass = "D2"
    '' N振替不可/C振替不可
    ElseIf tbl_chk1_10(1).NDOPHUFLG <> "0" And tbl_chk1_10(1).CDOPHUFLG <> "0" Then
        sClass = "D3"
    End If
'' add end 0108
    
    sResult = ""
    RET = funCodeDBGet("SB", sClass, tbl_chk1_10(0).HWFCDOP, 1, tbl_chk1_10(1).HWFCDOP, sResult) '' chg 0108
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→元:" & tbl_chk1_10(0).HWFCDOP & ", 先:" & tbl_chk1_10(1).HWFCDOP
        GoTo CodeDBGet_Error
    End If
    If sResult = 0 Then
        funChkFurikae1_10 = 1
        iErr_Code = 1003
        sErr_Msg = "CHECK1-10,結晶ドープ、振替不可能です。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00006"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    '溝位置方位のチェック（同一分類グループなら振替可能）
    sErr_Msg = "1-10 溝位置方位ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "MZ", tbl_chk1_10(0).HSXDPDIR, 1, tbl_chk1_10(1).HSXDPDIR, sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→元:" & tbl_chk1_10(0).HSXDPDIR & ", 先:" & tbl_chk1_10(1).HSXDPDIR
        GoTo CodeDBGet_Error
    End If
    If sResult = 0 Then
        funChkFurikae1_10 = 1
        iErr_Code = 1004
        sErr_Msg = "CHECK1-10,溝位置方位、振替不可能です。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00008"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    '品種のチェック
    sErr_Msg = "1-10 品種ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "HS", tbl_chk1_10(0).MCNO1, 1, tbl_chk1_10(1).MCNO1, sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→元:" & tbl_chk1_10(0).MCNO1 & ", 先:" & tbl_chk1_10(1).MCNO1
        GoTo CodeDBGet_Error
    End If
    If sResult = 0 Then
        funChkFurikae1_10 = 1
        iErr_Code = 1005
        sErr_Msg = "CHECK1-10,品種、振替不可能です。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00010"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    '引上げ速度
    sErr_Msg = "1-10 引上げ速度ﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "HK", tbl_chk1_10(0).MCNO2, 1, tbl_chk1_10(1).MCNO2, sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→元:" & tbl_chk1_10(0).MCNO2 & ", 先:" & tbl_chk1_10(1).MCNO2
        GoTo CodeDBGet_Error
    End If
    If sResult = 0 Then
        funChkFurikae1_10 = 1
        iErr_Code = 1006
        sErr_Msg = "CHECK1-10,引上げ速度、振替不可能です。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00011"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'ＨＺタイプチェック
    sErr_Msg = "1-10 HZﾀｲﾌﾟﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "HZ", tbl_chk1_10(0).MCNO3, 1, tbl_chk1_10(1).MCNO3, sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→元:" & tbl_chk1_10(0).MCNO3 & ", 先:" & tbl_chk1_10(1).MCNO3
        GoTo CodeDBGet_Error
    End If
    If sResult = 0 Then
        funChkFurikae1_10 = 1
        iErr_Code = 1007
        sErr_Msg = "CHECK1-10,ＨＺタイプ、振替不可能です。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00012"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    'ドローチューブチェック
    sErr_Msg = "1-10 ﾄﾞﾛｰﾁｭｰﾌﾞﾁｪｯｸ"
    sResult = ""
    RET = funCodeDBGet("SB", "DC", tbl_chk1_10(0).DCHYUUBU, 1, tbl_chk1_10(1).DCHYUUBU, sResult)
    If RET <> 0 Then
        sErr_Msg = sErr_Msg & "→元:" & tbl_chk1_10(0).DCHYUUBU & ", 先:" & tbl_chk1_10(1).DCHYUUBU
        GoTo CodeDBGet_Error
    End If
    If sResult = 0 Then
        funChkFurikae1_10 = 1
        iErr_Code = 1008
        sErr_Msg = "CHECK1-10,ドローチューブ、振替不可能です。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00009"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae1_10 = 0 Then
        funChkFurikae1_10 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae1_10 = -4
    GoTo Apl_Exit
    
CodeDBGet_Error:
    Set rs = Nothing
    If funChkFurikae1_10 = 0 Then
        funChkFurikae1_10 = -5
    End If
    GoTo Apl_Exit

End Function

'------------------------------------------------
' 狙い品番チェック
'------------------------------------------------

'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :tKumi_Hinban()  ,I  ,tFullHinban  :ﾁｪｯｸ品番
'          :iKumi_Row()     ,I  ,Integer      :品番行位置
'          :iCC10           ,I  ,Integer      :結晶設計変更工程だったら１
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :

Public Function funChkFurikae1_11(sProccd As String, sKeyID As String, _
                                 tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iCC10 As Integer, iErr_Code As Integer, sErr_Msg As String) As Integer

    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim i           As Integer
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae1_11 = 0
    
   
    '------------------------------------------ チェック先品番仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-11 チェック先品番仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E036.NHINCHKFLG " & vbCrLf
    sql = sql & "FROM   TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E036.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tNew_Hinban.opecond & "'  " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    Erase tbl_chk1_11
    With tbl_chk1_11(0)
'>>>>> NHINCHKFLGが空白の場合にチェックNGとなるため修正 2011/05/11 SETsw kubota ------------
'        If IsNull(rs("NHINCHKFLG")) = False Then .NHINCHKFLG = rs("NHINCHKFLG") Else .NHINCHKFLG = "0"          ' 狙い品番チェックフラグ
        .NHINCHKFLG = NulltoStr(rs("NHINCHKFLG"))
        If .NHINCHKFLG <> "1" Then
            .NHINCHKFLG = "0"
        End If
'<<<<< NHINCHKFLGが空白の場合にチェックNGとなるため修正 2011/05/11 SETsw kubota ------------
    
        Set rs = Nothing
        '判定有無
        If .NHINCHKFLG = "0" Then GoTo Apl_Exit
    
    End With
           
    '---------------------------------- 狙い品番データ取得 ------------------------------------------------------
    sErr_Msg = "1-11 狙い品番取得(" & sKeyID & ")"
    'SQL文の作成
    sql = vbNullString
    If iCC10 = 1 Then      '結晶設計変更工程
        sql = sql & "SELECT HINBAN PUHINBC1,NMNOREVNO PUREVNUMC1,NFACTORY PUFACTORYC1,NOPECOND PUOPEC1 " & vbCrLf
        sql = sql & "FROM   TBCMH001 " & vbCrLf
        sql = sql & "WHERE  UPINDNO    =   '" & Mid(sKeyID, 1, 7) & "00" & "'   " & vbCrLf
    Else
        sql = sql & "SELECT PUHINBC1,PUREVNUMC1,PUFACTORYC1,PUOPEC1 " & vbCrLf
        sql = sql & "FROM   XSDC1,XSDCA " & vbCrLf
        sql = sql & "WHERE  XTALCA    =   XTALC1  AND " & vbCrLf
        sql = sql & "       LIVKCA    =   '0'     AND " & vbCrLf
        sql = sql & "       ROWNUM    =   1       AND " & vbCrLf
        If sProccd = "CW761" Then
            sql = sql & "       SXLIDCA  =    '" & sKeyID & "' " & vbCrLf
        Else
            sql = sql & "       CRYNUMCA  =    '" & sKeyID & "' " & vbCrLf
        End If
    End If
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    '取得データセット
    With tbl_chk1_11(1)
        If IsNull(rs("PUHINBC1")) = False Then .hinban.hinban = rs("PUHINBC1") Else .hinban.hinban = " "
        If IsNull(rs("PUREVNUMC1")) = False Then .hinban.mnorevno = rs("PUREVNUMC1") Else .hinban.mnorevno = " "
        If IsNull(rs("PUFACTORYC1")) = False Then .hinban.factory = rs("PUFACTORYC1") Else .hinban.factory = " "
        If IsNull(rs("PUOPEC1")) = False Then .hinban.opecond = rs("PUOPEC1") Else .hinban.opecond = " "
    End With
    
    Set rs = Nothing
    On Error GoTo Apl_down
    '品番３桁チェック
    sErr_Msg = "1-11 品番ﾁｪｯｸ"
    If left(tNew_Hinban.hinban, 3) <> left(tbl_chk1_11(1).hinban.hinban, 3) Then
        funChkFurikae1_11 = 1
        iErr_Code = 1101
        sErr_Msg = "CHECK1-11,狙い品番(3桁)不一致の為、振替できません。"
        gsTbcmy028ErrCode = "00130"
        GoTo Apl_Exit
    End If
    
    'フラグチェック   '2011/06/20 Kameda
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT NVL(E036.NHINCHKFLG,' ')  NHINCHKFLG " & vbCrLf
    sql = sql & "FROM   TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E036.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tOld_Hinban.opecond & "'  " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    Erase tbl_chk1_11
    With tbl_chk1_11(1)
        .NHINCHKFLG = rs("NHINCHKFLG")
        If .NHINCHKFLG <> "1" Then
            .NHINCHKFLG = "0"
        End If
    
        Set rs = Nothing
    
        sErr_Msg = "1-11 品番ﾁｪｯｸ"
        If .NHINCHKFLG <> "1" Then
            funChkFurikae1_11 = 1
            iErr_Code = 1101
            sErr_Msg = "CHECK1-11,品番チェックフラグ不一致、振替できません。"
            gsTbcmy028ErrCode = "00131"
            GoTo Apl_Exit
        End If
    End With
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae1_11 = 0 Then
        funChkFurikae1_11 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae1_11 = -4
    GoTo Apl_Exit
    
CodeDBGet_Error:
    Set rs = Nothing
    If funChkFurikae1_11 = 0 Then
        funChkFurikae1_11 = -5
    End If
    GoTo Apl_Exit

End Function

'Add Start 2011/04/20 SMPK Miyata
'------------------------------------------------
' 振替先と振替元の中間抜試仕様チェック
'------------------------------------------------

'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sKeyID          ,I  ,String       :ﾌﾞﾛｯｸID、又は、SXL-ID
'          :tOld_Hinban     ,I  ,tFullHinban  :振替元品番(構造体)
'          :tNew_Hinban     ,I  ,tFullHinban  :振替先品番(構造体)
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :

Public Function funChkFurikae1_12(sProccd As String, sBlockId As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                  iErr_Code As Integer, sErr_Msg As String) As Integer


    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim iSxlWfCnt   As Integer      'WF枚数

    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae1_12 = 0
    
    '------------------------------------------ 振替元品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-12 振替元品番仕様取得(" & tOld_Hinban.hinban & Format(tOld_Hinban.mnorevno, "00") & tOld_Hinban.factory & tOld_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E036.MSMPFLG, E036.MSMPTANIMAI " & vbCrLf
    sql = sql & "FROM   TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E036.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tOld_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If

    '取得データセット
    Erase tbl_chk1_12
    With tbl_chk1_12(0)
        If IsNull(rs("MSMPFLG")) = False Then .MSMPFLG = rs("MSMPFLG") Else .MSMPFLG = "0"                  '中間抜試フラグ
        If IsNull(rs("MSMPTANIMAI")) = False Then .MSMPTANIMAI = rs("MSMPTANIMAI") Else .MSMPTANIMAI = 0    '中間抜試単位(枚数)
    End With
    
    Set rs = Nothing
    '------------------------------------------ 振替先品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-12 振替先品番仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E036.MSMPFLG, E036.MSMPTANIMAI " & vbCrLf
    sql = sql & "FROM   TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E036.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tNew_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    With tbl_chk1_12(1)
        If IsNull(rs("MSMPFLG")) = False Then .MSMPFLG = rs("MSMPFLG") Else .MSMPFLG = "0"                  '中間抜試フラグ
        If IsNull(rs("MSMPTANIMAI")) = False Then .MSMPTANIMAI = rs("MSMPTANIMAI") Else .MSMPTANIMAI = 0    '中間抜試単位(枚数)
    End With
    
    Set rs = Nothing
    '------------------------------------------ 分割結晶(SXL)データ取得 ------------------------------------------------------
    sErr_Msg = "1-12 振替元分割結晶(SXL)データ取得(" & sBlockId & ")"

    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT"
    sql = sql & " MAICB" & vbCrLf           '実枚数
    sql = sql & "FROM XSDCB " & vbCrLf
    sql = sql & "WHERE SXLIDCB = '" & sBlockId & "'" & vbCrLf

    On Error GoTo db_Error

    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    '該当データなし
    If rs.EOF Or rs.RecordCount <> 1 Then
        GoTo db_Error
    End If

    '取得データセット
    If IsNull(rs("MAICB")) = False Then iSxlWfCnt = rs("MAICB") Else iSxlWfCnt = 0   '実枚数

    Set rs = Nothing
    '------------------------------------------ 指示取得 ------------------------------------------------------
    On Error GoTo Apl_down

    '中間抜試単位のチェック
    sErr_Msg = "1-12 中間抜試単位ﾁｪｯｸ"

'Cng Start 2011/08/11 Y.Hitomi
    '中間抜試無し品から有り品か？
'    If tbl_chk1_12(0).MSMPFLG = "0" And tbl_chk1_12(1).MSMPFLG = "1" Then
    If tbl_chk1_12(0).MSMPFLG = "0" And (tbl_chk1_12(1).MSMPFLG = "1" Or tbl_chk1_12(0).MSMPFLG = "3") Then
'Cng End   2011/08/11 Y.Hitomi
    
        'Cng Start 2011/07/19 Y.Hitomi   中間無⇒有は実績チェックに委ねるが、実績なし品は、ここでチェックする
        '　　　　　　　　　　　　　　　　但し、製品仕様（保証）品のみとする
        If iSxlWfCnt >= tbl_chk1_12(1).MSMPTANIMAI Then
'Cng Start 2011/09/26 Y.Hitomi
            If ChkSXL_XSDCW_1(sBlockId) <> FUNCTION_RETURN_SUCCESS Then
'            If ChkSXL_XSDCW_1(sBlockId) <> FUNCTION_RETURN_SUCCESS _
'               And ChkMidSpec(tNew_Hinban.hinban, tNew_Hinban.opecond) Then
'Cng End   2011/09/26 Y.Hitomi
                funChkFurikae1_12 = 1
                iErr_Code = 11201
                sErr_Msg = "CHECK1-12,中間抜試実績無し⇒有りは振替できません。"
                gsTbcmy028ErrCode = "00131"
                GoTo Apl_Exit
            End If
        End If
   End If
        'Cng End   2011/07/19  Y.Hitomi
        
'Del Start 2011/07/28 Y.Hitomi
'    '中間抜試有り品から有り品か？
'    ElseIf tbl_chk1_12(0).MSMPFLG = "1" And tbl_chk1_12(1).MSMPFLG = "1" Then
'
'        '振替元品番構成長さ(SXL長さ)が振替先の中間抜試単位より長い?
'        If iSxlWfCnt >= tbl_chk1_12(1).MSMPTANIMAI Then
'
'            '中間抜試単位が振替元より振替先の方が短い?
'            If Trim$(tbl_chk1_12(0).MSMPTANIMAI) > Trim$(tbl_chk1_12(1).MSMPTANIMAI) Then
'                funChkFurikae1_12 = 1
'                iErr_Code = 11202
'                sErr_Msg = "CHECK1-12,振替先の方が中間抜試単位が短い為、振替できません。"
'                gsTbcmy028ErrCode = "00132"
'                GoTo Apl_Exit
'            End If
'        End If
'Del End 2011/07/28 Y.Hitomi
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae1_12 = 0 Then
        funChkFurikae1_12 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae1_12 = -4
    GoTo Apl_Exit
    
CodeDBGet_Error:
    Set rs = Nothing
    If funChkFurikae1_12 = 0 Then
        funChkFurikae1_12 = -5
    End If
    GoTo Apl_Exit

End Function
'Add End   2011/04/20 SMPK Miyata

'------------------------------------------------
' マルチ引上げ適用可否チェック
'------------------------------------------------

'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sBlockId        ,I  ,String       :ﾌﾞﾛｯｸID、又は、SXL-ID
'          :tOld_Hinban     ,I  ,tFullHinban  :振替元品番(構造体)
'          :tNew_Hinban     ,I  ,tFullHinban  :振替先品番(構造体)
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :11/05/19 Kameda

Public Function funChkFurikae1_13(sBlockId As String, _
                                 tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iErr_Code As Integer, sErr_Msg As String) As Integer

    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim i           As Integer
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae1_13 = 0
    
    '------------------------------------------ 判定データ取得(XSDC1) ------------------------------------------
    '連続コード取得
    sql = "SELECT NVL(SIJICNT,0) SIJICNT,NVL(RENBAN,0) RENBAN " & vbCrLf
    sql = sql & "FROM XSDC1,TBCMH001 " & vbCrLf
    sql = sql & "WHERE XTALC1 = '" & left(sBlockId, 9) & "000" & "' " & vbCrLf
    sql = sql & " AND  HISIJIC1 = UPINDNO "
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    Erase tbl_chk1_13
    With tbl_chk1_13(0)
        
        .SIJICNT = rs("SIJICNT")
        .RENBAN = rs("RENBAN")
        
        Set rs = Nothing
        
        '判定有無
        If .SIJICNT <= 1 Then GoTo Apl_Exit   '当該結晶がマルチ引上バッチの場合のみチェック
        If .RENBAN <= 1 Then GoTo Apl_Exit    'マルチ引上げバッチでも１本目は対象外とします
    
    End With
    
    '-------------------------------- 振替元マルチ引上げ適用可否仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-13 振替元品番仕様取得(" & tOld_Hinban.hinban & Format(tOld_Hinban.mnorevno, "00") & tOld_Hinban.factory & tOld_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT NVL(E036.MLTHTFLG,' ')  MLTHTFLG " & vbCrLf
    sql = sql & "FROM   TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E036.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tOld_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    With tbl_chk1_13(0)
    ' マルチ引上げ適用可否フラグ
        If Trim(rs("MLTHTFLG")) <> "" Then
            .MLTHTFLG = rs("MLTHTFLG")
        Else
            .MLTHTFLG = "0"
        End If
    End With
    
    Set rs = Nothing
    
    '--------------------------------- 振替先マルチ引上げ適用可否仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-13 振替先品番仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT NVL(E036.MLTHTFLG,' ')  MLTHTFLG  " & vbCrLf
    sql = sql & "FROM   TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E036.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tNew_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    With tbl_chk1_13(1)
    ' マルチ引上げ適用可否フラグ
        If Trim(rs("MLTHTFLG")) <> "" Then
            .MLTHTFLG = rs("MLTHTFLG")
        Else
            .MLTHTFLG = "0"
        End If
    End With
    
    Set rs = Nothing

    '＜振替条件＞
    '--振替元品番--        --振替先品番--         --振替結果--
    'マルチ引上げ適用可    マルチ引上げ適用否       振替可能
    'マルチ引上げ適用否    マルチ引上げ適用可       振替不可
    
    '[マルチ引上げ可否フラグ：0＝可　1＝不可]
    
    On Error GoTo Apl_down
    sErr_Msg = "1-13 マルチ引上げ適用ﾁｪｯｸ"
    If tbl_chk1_13(0).MLTHTFLG = "1" Then
        If tbl_chk1_13(1).MLTHTFLG = "0" Then
            funChkFurikae1_13 = 1
            iErr_Code = 1301
            sErr_Msg = "CHECK1-13,マルチ引上げ適用可否エラーの為、振替できません。"
            gsTbcmy028ErrCode = "01301"
            GoTo Apl_Exit
        End If
    End If
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae1_13 = 0 Then
        funChkFurikae1_13 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae1_13 = -4
    GoTo Apl_Exit

End Function

'Add Start 2011/05/11 SMPK Nakamura FRSシステム化対応
'------------------------------------------------
' FRS仕様チェック
'------------------------------------------------

'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sKeyID          ,I  ,String       :ﾌﾞﾛｯｸID、又は、SXL-ID
'          :tOld_Hinban     ,I  ,tFullHinban  :振替元品番(構造体)
'          :tNew_Hinban     ,I  ,tFullHinban  :振替先品番(構造体)
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :11/05/11 SMPK Nakamura

Public Function funChkFurikae1_14(sProccd As String, sKeyID As String, _
                                 tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iErr_Code As Integer, sErr_Msg As String) As Integer

    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim i           As Integer
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae1_14 = 0

    '------------------------------------------ 振替元FRS仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-14 振替元品番仕様取得(" & tOld_Hinban.hinban & Format(tOld_Hinban.mnorevno, "00") & tOld_Hinban.factory & tOld_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E036.FRSFLG " & vbCrLf
    sql = sql & "FROM   TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E036.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tOld_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    Erase tbl_chk1_14
    With tbl_chk1_14(0)
        ' FRS測定フラグ
        If IsNull(rs("FRSFLG")) = False Then
            If Trim(rs("FRSFLG")) <> "" Then
                .FRSFLG = rs("FRSFLG")
            Else
                .FRSFLG = "0"
            End If
        Else
            .FRSFLG = "0"
        End If
    End With
    
    Set rs = Nothing
    
    '------------------------------------------ 振替先FRS仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-14 振替先品番仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E036.FRSFLG " & vbCrLf
    sql = sql & "FROM   TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E036.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tNew_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    'Del Start 2011/07/12 SMPK Nakamura
'    Erase tbl_chk1_14
    'Del End 2011/07/12 SMPK Nakamura
    With tbl_chk1_14(1)
        ' FRS測定フラグ
        If IsNull(rs("FRSFLG")) = False Then
            If Trim(rs("FRSFLG")) <> "" Then
                .FRSFLG = rs("FRSFLG")
            Else
                .FRSFLG = "0"
            End If
        Else
            .FRSFLG = "0"
        End If
    End With
    
    Set rs = Nothing

    '判定有無(振替元品番FRSフラグが"1"の場合は振替OK
    If tbl_chk1_14(0).FRSFLG = "1" Then GoTo Apl_Exit

    On Error GoTo Apl_down
    '品番３桁チェック
    sErr_Msg = "1-14 品番ﾁｪｯｸ"
    If tbl_chk1_14(1).FRSFLG = "1" Then
        funChkFurikae1_14 = 1
        iErr_Code = 1401
        sErr_Msg = "CHECK1-14,FRS測定無し→有りには、振替できません。" '
        gsTbcmy028ErrCode = "00140"
        GoTo Apl_Exit
    End If
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae1_14 = 0 Then
        funChkFurikae1_14 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae1_14 = -4
    GoTo Apl_Exit

End Function
'Add End 2011/05/11 SMPK Nakamura FRSシステム化対応

'Add Start 2011/07/12 SMPK Nakamura 結晶面傾きチェック追加対応
'------------------------------------------------
' 結晶面傾きチェック
'------------------------------------------------

'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sKeyID          ,I  ,String       :ﾌﾞﾛｯｸID、又は、SXL-ID
'          :tOld_Hinban     ,I  ,tFullHinban  :振替元品番(構造体)
'          :tNew_Hinban     ,I  ,tFullHinban  :振替先品番(構造体)
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :11/07/12 SMPK Nakamura

Public Function funChkFurikae1_15(sProccd As String, sKeyID As String, _
                                  tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                  iErr_Code As Integer, sErr_Msg As String) As Integer

    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim i           As Integer
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae1_15 = 0
    
    '取得データセット初期化
    Erase tbl_chk1_15
    
    '------------------------------------------ 振替元品番仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-15 振替元品番仕様取得(" & tOld_Hinban.hinban & Format(tOld_Hinban.mnorevno, "00") & tOld_Hinban.factory & tOld_Hinban.opecond & ")"
    'SQL文の作成
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E018.HSXCDIR,E018.HSXCSCEN,E018.HSXCSMIN,E018.HSXCSMAX, " & vbCrLf
    sql = sql & "       E018.HSXCKWAY,E018.HSXCKHNM,E018.HSXCKHNI,E018.HSXCKHNH,E018.HSXCKHNS, " & vbCrLf
    sql = sql & "       E018.HSXCSDIR,E018.HSXCSDIS, " & vbCrLf
    sql = sql & "       E018.HSXCTDIR,E018.HSXCTCEN,E018.HSXCTMIN,E018.HSXCTMAX, " & vbCrLf
    sql = sql & "       E018.HSXCYDIR,E018.HSXCYCEN,E018.HSXCYMIN,E018.HSXCYMAX, " & vbCrLf
    sql = sql & "       E027.HWFCSGCEN,E027.HWFCSGMIN,E027.HWFCSGMAX, " & vbCrLf
    sql = sql & "       E027.HWFCSXCEN,E027.HWFCSXMIN,E027.HWFCSXMAX, " & vbCrLf
    sql = sql & "       E027.HWFCSYCEN,E027.HWFCSYMIN,E027.HWFCSYMAX  " & vbCrLf
    sql = sql & "FROM   TBCME018 E018, TBCME027 E027 " & vbCrLf
    sql = sql & "WHERE  E018.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E018.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E018.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E027.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E027.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E027.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E027.OPECOND   =   '" & tOld_Hinban.opecond & "'"
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    With tbl_chk1_15(0)
        If IsNull(rs("HSXCDIR")) = False Then .HSXCDIR = rs("HSXCDIR") Else .HSXCDIR = " "      ' ＳＸＬ結晶面方位
        .HSXCSCEN = fncNullCheck(rs("HSXCSCEN"))                                                ' ＳＸＬ結晶面傾き中心
        .HSXCSMIN = fncNullCheck(rs("HSXCSMIN"))                                                ' ＳＸＬ結晶面傾き下限
        .HSXCSMAX = fncNullCheck(rs("HSXCSMAX"))                                                ' ＳＸＬ結晶面傾き上限
        If IsNull(rs("HSXCKWAY")) = False Then .HSXCKWAY = rs("HSXCKWAY") Else .HSXCKWAY = " "  ' ＳＸＬ結晶面検査方法
        If IsNull(rs("HSXCKHNM")) = False Then .HSXCKHNM = rs("HSXCKHNM") Else .HSXCKHNM = " "  ' ＳＸＬ結晶面検査頻度_枚
        If IsNull(rs("HSXCKHNI")) = False Then .HSXCKHNI = rs("HSXCKHNI") Else .HSXCKHNI = " "  ' ＳＸＬ結晶面検査頻度_位
        If IsNull(rs("HSXCKHNH")) = False Then .HSXCKHNH = rs("HSXCKHNH") Else .HSXCKHNH = " "  ' ＳＸＬ結晶面検査頻度_保
        If IsNull(rs("HSXCKHNS")) = False Then .HSXCKHNS = rs("HSXCKHNS") Else .HSXCKHNS = " "  ' ＳＸＬ結晶面検査頻度_試
        If IsNull(rs("HSXCSDIR")) = False Then .HSXCSDIR = rs("HSXCSDIR") Else .HSXCSDIR = " "  ' ＳＸＬ結晶面傾き方位
        If IsNull(rs("HSXCSDIS")) = False Then .HSXCSDIS = rs("HSXCSDIS") Else .HSXCSDIS = " "  ' ＳＸＬ結晶面傾き方位指定
        If IsNull(rs("HSXCTDIR")) = False Then .HSXCTDIR = rs("HSXCTDIR") Else .HSXCTDIR = " "  ' ＳＸＬ結晶面傾き縦方位
        .HSXCTCEN = fncNullCheck(rs("HSXCTCEN"))                                                ' ＳＸＬ結晶面傾き縦中心
        .HSXCTMIN = fncNullCheck(rs("HSXCTMIN"))                                                ' ＳＸＬ結晶面傾き縦下限
        .HSXCTMAX = fncNullCheck(rs("HSXCTMAX"))                                                ' ＳＸＬ結晶面傾き縦上限
        If IsNull(rs("HSXCYDIR")) = False Then .HSXCYDIR = rs("HSXCYDIR") Else .HSXCYDIR = " "  ' ＳＸＬ結晶面傾き横方位
        .HSXCYCEN = fncNullCheck(rs("HSXCYCEN"))                                                ' ＳＸＬ結晶面傾き横中心
        .HSXCYMIN = fncNullCheck(rs("HSXCYMIN"))                                                ' ＳＸＬ結晶面傾き横下限
        .HSXCYMAX = fncNullCheck(rs("HSXCYMAX"))                                                ' ＳＸＬ結晶面傾き横上限
        .HWFCSGCEN = fncNullCheck(rs("HWFCSGCEN"))                                              ' ＷＦ結晶面操合成角中心
        .HWFCSGMIN = fncNullCheck(rs("HWFCSGMIN"))                                              ' ＷＦ結晶面操合成角下限
        .HWFCSGMAX = fncNullCheck(rs("HWFCSGMAX"))                                              ' ＷＦ結晶面操合成角上限
        .HWFCSXCEN = fncNullCheck(rs("HWFCSXCEN"))                                              ' ＷＦ結晶面操Ｘ方位中心
        .HWFCSXMIN = fncNullCheck(rs("HWFCSXMIN"))                                              ' ＷＦ結晶面操Ｘ方位下限
        .HWFCSXMAX = fncNullCheck(rs("HWFCSXMAX"))                                              ' ＷＦ結晶面操Ｘ方位上限
        .HWFCSYCEN = fncNullCheck(rs("HWFCSYCEN"))                                              ' ＷＦ結晶面操Ｙ方位中心
        .HWFCSYMIN = fncNullCheck(rs("HWFCSYMIN"))                                              ' ＷＦ結晶面操Ｙ方位下限
        .HWFCSYMAX = fncNullCheck(rs("HWFCSYMAX"))                                              ' ＷＦ結晶面操Ｙ方位上限
    End With
    
    Set rs = Nothing
    
    '------------------------------------------ 振替先品番仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-15 振替先品番仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E018.HSXCDIR,E018.HSXCSCEN,E018.HSXCSMIN,E018.HSXCSMAX, " & vbCrLf
    sql = sql & "       E018.HSXCKWAY,E018.HSXCKHNM,E018.HSXCKHNI,E018.HSXCKHNH,E018.HSXCKHNS, " & vbCrLf
    sql = sql & "       E018.HSXCSDIR,E018.HSXCSDIS, " & vbCrLf
    sql = sql & "       E018.HSXCTDIR,E018.HSXCTCEN,E018.HSXCTMIN,E018.HSXCTMAX, " & vbCrLf
    sql = sql & "       E018.HSXCYDIR,E018.HSXCYCEN,E018.HSXCYMIN,E018.HSXCYMAX, " & vbCrLf
    sql = sql & "       E027.HWFCSGCEN,E027.HWFCSGMIN,E027.HWFCSGMAX, " & vbCrLf
    sql = sql & "       E027.HWFCSXCEN,E027.HWFCSXMIN,E027.HWFCSXMAX, " & vbCrLf
    sql = sql & "       E027.HWFCSYCEN,E027.HWFCSYMIN,E027.HWFCSYMAX  " & vbCrLf
    sql = sql & "FROM   TBCME018 E018, TBCME027 E027 " & vbCrLf
    sql = sql & "WHERE  E018.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E018.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E018.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E027.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E027.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E027.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E027.OPECOND   =   '" & tNew_Hinban.opecond & "'"

    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    With tbl_chk1_15(1)
        If IsNull(rs("HSXCDIR")) = False Then .HSXCDIR = rs("HSXCDIR") Else .HSXCDIR = " "      ' ＳＸＬ結晶面方位
        .HSXCSCEN = fncNullCheck(rs("HSXCSCEN"))                                                ' ＳＸＬ結晶面傾き中心
        .HSXCSMIN = fncNullCheck(rs("HSXCSMIN"))                                                ' ＳＸＬ結晶面傾き下限
        .HSXCSMAX = fncNullCheck(rs("HSXCSMAX"))                                                ' ＳＸＬ結晶面傾き上限
        If IsNull(rs("HSXCKWAY")) = False Then .HSXCKWAY = rs("HSXCKWAY") Else .HSXCKWAY = " "  ' ＳＸＬ結晶面検査方法
        If IsNull(rs("HSXCKHNM")) = False Then .HSXCKHNM = rs("HSXCKHNM") Else .HSXCKHNM = " "  ' ＳＸＬ結晶面検査頻度_枚
        If IsNull(rs("HSXCKHNI")) = False Then .HSXCKHNI = rs("HSXCKHNI") Else .HSXCKHNI = " "  ' ＳＸＬ結晶面検査頻度_位
        If IsNull(rs("HSXCKHNH")) = False Then .HSXCKHNH = rs("HSXCKHNH") Else .HSXCKHNH = " "  ' ＳＸＬ結晶面検査頻度_保
        If IsNull(rs("HSXCKHNS")) = False Then .HSXCKHNS = rs("HSXCKHNS") Else .HSXCKHNS = " "  ' ＳＸＬ結晶面検査頻度_試
        If IsNull(rs("HSXCSDIR")) = False Then .HSXCSDIR = rs("HSXCSDIR") Else .HSXCSDIR = " "  ' ＳＸＬ結晶面傾き方位
        If IsNull(rs("HSXCSDIS")) = False Then .HSXCSDIS = rs("HSXCSDIS") Else .HSXCSDIS = " "  ' ＳＸＬ結晶面傾き方位指定
        If IsNull(rs("HSXCTDIR")) = False Then .HSXCTDIR = rs("HSXCTDIR") Else .HSXCTDIR = " "  ' ＳＸＬ結晶面傾き縦方位
        .HSXCTCEN = fncNullCheck(rs("HSXCTCEN"))                                                ' ＳＸＬ結晶面傾き縦中心
        .HSXCTMIN = fncNullCheck(rs("HSXCTMIN"))                                                ' ＳＸＬ結晶面傾き縦下限
        .HSXCTMAX = fncNullCheck(rs("HSXCTMAX"))                                                ' ＳＸＬ結晶面傾き縦上限
        If IsNull(rs("HSXCYDIR")) = False Then .HSXCYDIR = rs("HSXCYDIR") Else .HSXCYDIR = " "  ' ＳＸＬ結晶面傾き横方位
        .HSXCYCEN = fncNullCheck(rs("HSXCYCEN"))                                                ' ＳＸＬ結晶面傾き横中心
        .HSXCYMIN = fncNullCheck(rs("HSXCYMIN"))                                                ' ＳＸＬ結晶面傾き横下限
        .HSXCYMAX = fncNullCheck(rs("HSXCYMAX"))                                                ' ＳＸＬ結晶面傾き横上限
        .HWFCSGCEN = fncNullCheck(rs("HWFCSGCEN"))                                              ' ＷＦ結晶面操合成角中心
        .HWFCSGMIN = fncNullCheck(rs("HWFCSGMIN"))                                              ' ＷＦ結晶面操合成角下限
        .HWFCSGMAX = fncNullCheck(rs("HWFCSGMAX"))                                              ' ＷＦ結晶面操合成角上限
        .HWFCSXCEN = fncNullCheck(rs("HWFCSXCEN"))                                              ' ＷＦ結晶面操Ｘ方位中心
        .HWFCSXMIN = fncNullCheck(rs("HWFCSXMIN"))                                              ' ＷＦ結晶面操Ｘ方位下限
        .HWFCSXMAX = fncNullCheck(rs("HWFCSXMAX"))                                              ' ＷＦ結晶面操Ｘ方位上限
        .HWFCSYCEN = fncNullCheck(rs("HWFCSYCEN"))                                              ' ＷＦ結晶面操Ｙ方位中心
        .HWFCSYMIN = fncNullCheck(rs("HWFCSYMIN"))                                              ' ＷＦ結晶面操Ｙ方位下限
        .HWFCSYMAX = fncNullCheck(rs("HWFCSYMAX"))                                              ' ＷＦ結晶面操Ｙ方位上限
    End With
    
    Set rs = Nothing
    On Error GoTo Apl_down
    
    If left(sProccd, 2) = "CC" Then
        '◆ＳＸＬ結晶面傾き中心仕様チェック
        ' 振替元・先一致チェック
        sErr_Msg = "1-15 結晶面傾中心ﾁｪｯｸ"
        If tbl_chk1_15(0).HSXCSCEN <> tbl_chk1_15(1).HSXCSCEN Then
            ' 振替元・先の傾き中心が1.00度以内かをチェック(例外)
            If tbl_chk1_15(0).HSXCSCEN > 1 Or tbl_chk1_15(1).HSXCSCEN > 1 Then
                ' 結晶面傾チェックNG
                funChkFurikae1_15 = 1
                iErr_Code = 1501
                sErr_Msg = "CHECK1-15,結晶面傾中心不一致の為、振替できません。"
                gsTbcmy028ErrCode = "00150"
                GoTo Apl_Exit
            End If
        End If
    ElseIf left(sProccd, 2) = "CW" Then
        '◆ＷＦ結晶面傾き中心仕様チェック
        '◆スライスターゲット(WF結晶面傾中心、WF結晶面傾縦中心、WF結晶面傾横中心)チェック
        Dim blnSliceTarget As Boolean
        blnSliceTarget = True
        ' WF結晶面傾き中心仕様一致チェック
        If tbl_chk1_15(0).HWFCSGCEN = tbl_chk1_15(1).HWFCSGCEN Then
            ' 振替元のWF結晶面傾き縦中心、縦下限、縦上限が全て設定なしかをチェック
            If tbl_chk1_15(0).HWFCSYCEN = -1 And _
               tbl_chk1_15(0).HWFCSYMIN = -1 And _
               tbl_chk1_15(0).HWFCSYMAX = -1 Then
                blnSliceTarget = True
            Else
                ' 振替先のWF結晶面傾き縦中心、縦下限、縦上限が全て設定なしかをチェック
                If tbl_chk1_15(1).HWFCSYCEN = -1 And _
                   tbl_chk1_15(1).HWFCSYMIN = -1 And _
                   tbl_chk1_15(1).HWFCSYMAX = -1 Then
                    blnSliceTarget = True
                Else
                    ' WF結晶面傾き縦中心仕様一致チェック
                    If tbl_chk1_15(0).HWFCSYCEN = tbl_chk1_15(1).HWFCSYCEN Then
                        blnSliceTarget = True
                    Else
                        blnSliceTarget = False
                    End If
                End If
            End If
            If blnSliceTarget = True Then
                ' 振替元のWF結晶面傾き横中心、横下限、横上限が全て設定なしかをチェック
                If tbl_chk1_15(0).HWFCSXCEN = -1 And _
                   tbl_chk1_15(0).HWFCSXMIN = -1 And _
                   tbl_chk1_15(0).HWFCSXMAX = -1 Then
                    blnSliceTarget = True
                Else
                    ' 振替先のWF結晶面傾き横中心、横下限、横上限が全て設定なしかをチェック
                    If tbl_chk1_15(1).HWFCSXCEN = -1 And _
                       tbl_chk1_15(1).HWFCSXMIN = -1 And _
                       tbl_chk1_15(1).HWFCSXMAX = -1 Then
                        blnSliceTarget = True
                    Else
                        ' WF結晶面傾き横中心仕様一致チェック
                        If tbl_chk1_15(0).HWFCSXCEN = tbl_chk1_15(1).HWFCSXCEN Then
                            blnSliceTarget = True
                        Else
                            blnSliceTarget = False
                        End If
                    End If
                End If
            End If
        Else
            blnSliceTarget = False
        End If
        
        'ターゲット不一致(マルチブロックの場合は例外処理は実施しない)
        If blnSliceTarget = False Then
            ' 振替先のスライスターゲットが0.00度品かをチェック
'Add Start 2011/10/3 Y.Hitomi
            If tbl_chk1_15(1).HWFCSGCEN <> 0 Or _
               (tbl_chk1_15(1).HWFCSXCEN <> -1 And tbl_chk1_15(1).HWFCSXCEN <> 0) Or _
               (tbl_chk1_15(1).HWFCSYCEN <> -1 And tbl_chk1_15(1).HWFCSYCEN <> 0) Then
'            If tbl_chk1_15(1).HWFCSGCEN <> 0 Or _
'               tbl_chk1_15(1).HWFCSXCEN <> 0 Or _
'               tbl_chk1_15(1).HWFCSYCEN <> 0 Then
'Add Start 2011/10/3 Y.Hitomi
                ' 結晶面傾チェックNG
                funChkFurikae1_15 = 1
                iErr_Code = 1503
                sErr_Msg = "CHECK1-15,WF結晶面傾中心仕様が異なる為、振替できません。"
                gsTbcmy028ErrCode = "00152"
                GoTo Apl_Exit
            End If
        End If
    End If
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:

    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae1_15 = 0 Then
        funChkFurikae1_15 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae1_15 = -4
    GoTo Apl_Exit

End Function

'------------------------------------------------
' 結晶面傾き組合せチェック
'------------------------------------------------

'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sKeyID          ,I  ,String       :ﾌﾞﾛｯｸID、又は、SXL-ID
'          :tTop_Hinban     ,I  ,tFullHinban  :ﾏﾙﾁﾌﾞﾛｯｸ先頭品番(構造体)
'          :tBtm_Hinban     ,I  ,tFullHinban  :ﾏﾙﾁﾌﾞﾛｯｸ最尾品番(構造体)
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :11/07/12 SMPK Nakamura

Public Function funChkFurikae1_16(sProccd As String, sKeyID As String, _
                                  tTop_Hinban As tFullHinban, tBtm_Hinban As tFullHinban, _
                                  iErr_Code As Integer, sErr_Msg As String) As Integer

    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim i           As Integer
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae1_16 = 0
    
    '取得データセット初期化
    Erase tbl_chk1_16
    
    '------------------------------------------ 先頭品番仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-16 先頭品番仕様取得(" & tTop_Hinban.hinban & Format(tTop_Hinban.mnorevno, "00") & tTop_Hinban.factory & tTop_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E018.HSXCDIR,E018.HSXCSCEN,E018.HSXCSMIN,E018.HSXCSMAX, " & vbCrLf
    sql = sql & "       E018.HSXCKWAY,E018.HSXCKHNM,E018.HSXCKHNI,E018.HSXCKHNH,E018.HSXCKHNS, " & vbCrLf
    sql = sql & "       E018.HSXCSDIR,E018.HSXCSDIS, " & vbCrLf
    sql = sql & "       E018.HSXCTDIR,E018.HSXCTCEN,E018.HSXCTMIN,E018.HSXCTMAX, " & vbCrLf
    sql = sql & "       E018.HSXCYDIR,E018.HSXCYCEN,E018.HSXCYMIN,E018.HSXCYMAX, " & vbCrLf
    sql = sql & "       E027.HWFCSGCEN,E027.HWFCSGMIN,E027.HWFCSGMAX, " & vbCrLf
    sql = sql & "       E027.HWFCSXCEN,E027.HWFCSXMIN,E027.HWFCSXMAX, " & vbCrLf
    sql = sql & "       E027.HWFCSYCEN,E027.HWFCSYMIN,E027.HWFCSYMAX  " & vbCrLf
    sql = sql & "FROM   TBCME018 E018, TBCME027 E027 " & vbCrLf
    sql = sql & "WHERE  E018.HINBAN    =   '" & tTop_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E018.MNOREVNO  =    " & tTop_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E018.FACTORY   =   '" & tTop_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND   =   '" & tTop_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E027.HINBAN    =   '" & tTop_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E027.MNOREVNO  =    " & tTop_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E027.FACTORY   =   '" & tTop_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E027.OPECOND   =   '" & tTop_Hinban.opecond & "'"

    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    With tbl_chk1_16(0)
        If IsNull(rs("HSXCDIR")) = False Then .HSXCDIR = rs("HSXCDIR") Else .HSXCDIR = " "      ' ＳＸＬ結晶面方位
        .HSXCSCEN = fncNullCheck(rs("HSXCSCEN"))                                                ' ＳＸＬ結晶面傾き中心
        .HSXCSMIN = fncNullCheck(rs("HSXCSMIN"))                                                ' ＳＸＬ結晶面傾き下限
        .HSXCSMAX = fncNullCheck(rs("HSXCSMAX"))                                                ' ＳＸＬ結晶面傾き上限
        If IsNull(rs("HSXCKWAY")) = False Then .HSXCKWAY = rs("HSXCKWAY") Else .HSXCKWAY = " "  ' ＳＸＬ結晶面検査方法
        If IsNull(rs("HSXCKHNM")) = False Then .HSXCKHNM = rs("HSXCKHNM") Else .HSXCKHNM = " "  ' ＳＸＬ結晶面検査頻度_枚
        If IsNull(rs("HSXCKHNI")) = False Then .HSXCKHNI = rs("HSXCKHNI") Else .HSXCKHNI = " "  ' ＳＸＬ結晶面検査頻度_位
        If IsNull(rs("HSXCKHNH")) = False Then .HSXCKHNH = rs("HSXCKHNH") Else .HSXCKHNH = " "  ' ＳＸＬ結晶面検査頻度_保
        If IsNull(rs("HSXCKHNS")) = False Then .HSXCKHNS = rs("HSXCKHNS") Else .HSXCKHNS = " "  ' ＳＸＬ結晶面検査頻度_試
        If IsNull(rs("HSXCSDIR")) = False Then .HSXCSDIR = rs("HSXCSDIR") Else .HSXCSDIR = " "  ' ＳＸＬ結晶面傾き方位
        If IsNull(rs("HSXCSDIS")) = False Then .HSXCSDIS = rs("HSXCSDIS") Else .HSXCSDIS = " "  ' ＳＸＬ結晶面傾き方位指定
        If IsNull(rs("HSXCTDIR")) = False Then .HSXCTDIR = rs("HSXCTDIR") Else .HSXCTDIR = " "  ' ＳＸＬ結晶面傾き縦方位
        .HSXCTCEN = fncNullCheck(rs("HSXCTCEN"))                                                ' ＳＸＬ結晶面傾き縦中心
        .HSXCTMIN = fncNullCheck(rs("HSXCTMIN"))                                                ' ＳＸＬ結晶面傾き縦下限
        .HSXCTMAX = fncNullCheck(rs("HSXCTMAX"))                                                ' ＳＸＬ結晶面傾き縦上限
        If IsNull(rs("HSXCYDIR")) = False Then .HSXCYDIR = rs("HSXCYDIR") Else .HSXCYDIR = " "  ' ＳＸＬ結晶面傾き横方位
        .HSXCYCEN = fncNullCheck(rs("HSXCYCEN"))                                                ' ＳＸＬ結晶面傾き横中心
        .HSXCYMIN = fncNullCheck(rs("HSXCYMIN"))                                                ' ＳＸＬ結晶面傾き横下限
        .HSXCYMAX = fncNullCheck(rs("HSXCYMAX"))                                                ' ＳＸＬ結晶面傾き横上限
        .HWFCSGCEN = fncNullCheck(rs("HWFCSGCEN"))                                              ' ＷＦ結晶面操合成角中心
        .HWFCSGMIN = fncNullCheck(rs("HWFCSGMIN"))                                              ' ＷＦ結晶面操合成角下限
        .HWFCSGMAX = fncNullCheck(rs("HWFCSGMAX"))                                              ' ＷＦ結晶面操合成角上限
        .HWFCSXCEN = fncNullCheck(rs("HWFCSXCEN"))                                              ' ＷＦ結晶面操Ｘ方位中心
        .HWFCSXMIN = fncNullCheck(rs("HWFCSXMIN"))                                              ' ＷＦ結晶面操Ｘ方位下限
        .HWFCSXMAX = fncNullCheck(rs("HWFCSXMAX"))                                              ' ＷＦ結晶面操Ｘ方位上限
        .HWFCSYCEN = fncNullCheck(rs("HWFCSYCEN"))                                              ' ＷＦ結晶面操Ｙ方位中心
        .HWFCSYMIN = fncNullCheck(rs("HWFCSYMIN"))                                              ' ＷＦ結晶面操Ｙ方位下限
        .HWFCSYMAX = fncNullCheck(rs("HWFCSYMAX"))                                              ' ＷＦ結晶面操Ｙ方位上限
    End With
    
    Set rs = Nothing
    
    '------------------------------------------ 最尾品番仕様データ取得 ------------------------------------------------------
    sErr_Msg = "1-16 最尾品番仕様取得(" & tBtm_Hinban.hinban & Format(tBtm_Hinban.mnorevno, "00") & tBtm_Hinban.factory & tBtm_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E018.HSXCDIR,E018.HSXCSCEN,E018.HSXCSMIN,E018.HSXCSMAX, " & vbCrLf
    sql = sql & "       E018.HSXCKWAY,E018.HSXCKHNM,E018.HSXCKHNI,E018.HSXCKHNH,E018.HSXCKHNS, " & vbCrLf
    sql = sql & "       E018.HSXCSDIR,E018.HSXCSDIS, " & vbCrLf
    sql = sql & "       E018.HSXCTDIR,E018.HSXCTCEN,E018.HSXCTMIN,E018.HSXCTMAX, " & vbCrLf
    sql = sql & "       E018.HSXCYDIR,E018.HSXCYCEN,E018.HSXCYMIN,E018.HSXCYMAX, " & vbCrLf
    sql = sql & "       E027.HWFCSGCEN,E027.HWFCSGMIN,E027.HWFCSGMAX, " & vbCrLf
    sql = sql & "       E027.HWFCSXCEN,E027.HWFCSXMIN,E027.HWFCSXMAX, " & vbCrLf
    sql = sql & "       E027.HWFCSYCEN,E027.HWFCSYMIN,E027.HWFCSYMAX  " & vbCrLf
    sql = sql & "FROM   TBCME018 E018, TBCME027 E027 " & vbCrLf
    sql = sql & "WHERE  E018.HINBAN    =   '" & tBtm_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E018.MNOREVNO  =    " & tBtm_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E018.FACTORY   =   '" & tBtm_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND   =   '" & tBtm_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E027.HINBAN    =   '" & tBtm_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E027.MNOREVNO  =    " & tBtm_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E027.FACTORY   =   '" & tBtm_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E027.OPECOND   =   '" & tBtm_Hinban.opecond & "'"
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    With tbl_chk1_16(1)
        If IsNull(rs("HSXCDIR")) = False Then .HSXCDIR = rs("HSXCDIR") Else .HSXCDIR = " "      ' ＳＸＬ結晶面方位
        .HSXCSCEN = fncNullCheck(rs("HSXCSCEN"))                                                ' ＳＸＬ結晶面傾き中心
        .HSXCSMIN = fncNullCheck(rs("HSXCSMIN"))                                                ' ＳＸＬ結晶面傾き下限
        .HSXCSMAX = fncNullCheck(rs("HSXCSMAX"))                                                ' ＳＸＬ結晶面傾き上限
        If IsNull(rs("HSXCKWAY")) = False Then .HSXCKWAY = rs("HSXCKWAY") Else .HSXCKWAY = " "  ' ＳＸＬ結晶面検査方法
        If IsNull(rs("HSXCKHNM")) = False Then .HSXCKHNM = rs("HSXCKHNM") Else .HSXCKHNM = " "  ' ＳＸＬ結晶面検査頻度_枚
        If IsNull(rs("HSXCKHNI")) = False Then .HSXCKHNI = rs("HSXCKHNI") Else .HSXCKHNI = " "  ' ＳＸＬ結晶面検査頻度_位
        If IsNull(rs("HSXCKHNH")) = False Then .HSXCKHNH = rs("HSXCKHNH") Else .HSXCKHNH = " "  ' ＳＸＬ結晶面検査頻度_保
        If IsNull(rs("HSXCKHNS")) = False Then .HSXCKHNS = rs("HSXCKHNS") Else .HSXCKHNS = " "  ' ＳＸＬ結晶面検査頻度_試
        If IsNull(rs("HSXCSDIR")) = False Then .HSXCSDIR = rs("HSXCSDIR") Else .HSXCSDIR = " "  ' ＳＸＬ結晶面傾き方位
        If IsNull(rs("HSXCSDIS")) = False Then .HSXCSDIS = rs("HSXCSDIS") Else .HSXCSDIS = " "  ' ＳＸＬ結晶面傾き方位指定
        If IsNull(rs("HSXCTDIR")) = False Then .HSXCTDIR = rs("HSXCTDIR") Else .HSXCTDIR = " "  ' ＳＸＬ結晶面傾き縦方位
        .HSXCTCEN = fncNullCheck(rs("HSXCTCEN"))                                                ' ＳＸＬ結晶面傾き縦中心
        .HSXCTMIN = fncNullCheck(rs("HSXCTMIN"))                                                ' ＳＸＬ結晶面傾き縦下限
        .HSXCTMAX = fncNullCheck(rs("HSXCTMAX"))                                                ' ＳＸＬ結晶面傾き縦上限
        If IsNull(rs("HSXCYDIR")) = False Then .HSXCYDIR = rs("HSXCYDIR") Else .HSXCYDIR = " "  ' ＳＸＬ結晶面傾き横方位
        .HSXCYCEN = fncNullCheck(rs("HSXCYCEN"))                                                ' ＳＸＬ結晶面傾き横中心
        .HSXCYMIN = fncNullCheck(rs("HSXCYMIN"))                                                ' ＳＸＬ結晶面傾き横下限
        .HSXCYMAX = fncNullCheck(rs("HSXCYMAX"))                                                ' ＳＸＬ結晶面傾き横上限
        .HWFCSGCEN = fncNullCheck(rs("HWFCSGCEN"))                                              ' ＷＦ結晶面操合成角中心
        .HWFCSGMIN = fncNullCheck(rs("HWFCSGMIN"))                                              ' ＷＦ結晶面操合成角下限
        .HWFCSGMAX = fncNullCheck(rs("HWFCSGMAX"))                                              ' ＷＦ結晶面操合成角上限
        .HWFCSXCEN = fncNullCheck(rs("HWFCSXCEN"))                                              ' ＷＦ結晶面操Ｘ方位中心
        .HWFCSXMIN = fncNullCheck(rs("HWFCSXMIN"))                                              ' ＷＦ結晶面操Ｘ方位下限
        .HWFCSXMAX = fncNullCheck(rs("HWFCSXMAX"))                                              ' ＷＦ結晶面操Ｘ方位上限
        .HWFCSYCEN = fncNullCheck(rs("HWFCSYCEN"))                                              ' ＷＦ結晶面操Ｙ方位中心
        .HWFCSYMIN = fncNullCheck(rs("HWFCSYMIN"))                                              ' ＷＦ結晶面操Ｙ方位下限
        .HWFCSYMAX = fncNullCheck(rs("HWFCSYMAX"))                                              ' ＷＦ結晶面操Ｙ方位上限
    End With
    
    Set rs = Nothing
    On Error GoTo Apl_down
    
    If left(sProccd, 2) = "CC" Then
        '◆ＳＸＬ結晶面傾き中心仕様チェック
        ' 振替元・先一致チェック
        sErr_Msg = "1-16 結晶面傾中心ﾁｪｯｸ"
        If tbl_chk1_16(0).HSXCSCEN <> tbl_chk1_16(1).HSXCSCEN Then
            ' 組合せ品番の傾き中心が1.00度以内かをチェック(例外)
            If tbl_chk1_16(0).HSXCSCEN > 1 Or tbl_chk1_16(1).HSXCSCEN > 1 Then
                ' 結晶面傾チェックNG
                funChkFurikae1_16 = 1
                iErr_Code = 1601
                sErr_Msg = "CHECK1-16,結晶面傾中心不一致の為、組合せできません。"
                gsTbcmy028ErrCode = "00160"
                GoTo Apl_Exit
            End If
        End If
    End If
    
    '◆ＷＦ結晶面傾き中心仕様チェック
    '◆スライスターゲット(WF結晶面傾中心、WF結晶面傾縦中心、WF結晶面傾横中心)チェック
    Dim blnSliceTarget As Boolean
    blnSliceTarget = True
    ' WF結晶面傾き中心仕様一致チェック
    If tbl_chk1_16(0).HWFCSGCEN = tbl_chk1_16(1).HWFCSGCEN Then
        ' 振替元のWF結晶面傾き縦中心、縦下限、縦上限が全て設定なしかをチェック
        If tbl_chk1_16(0).HWFCSYCEN = -1 And _
           tbl_chk1_16(0).HWFCSYMIN = -1 And _
           tbl_chk1_16(0).HWFCSYMAX = -1 Then
            blnSliceTarget = True
        Else
            ' 振替先のWF結晶面傾き縦中心、縦下限、縦上限が全て設定なしかをチェック
            If tbl_chk1_16(1).HWFCSYCEN = -1 And _
               tbl_chk1_16(1).HWFCSYMIN = -1 And _
               tbl_chk1_16(1).HWFCSYMAX = -1 Then
                blnSliceTarget = True
            Else
                ' WF結晶面傾き縦中心仕様一致チェック
                If tbl_chk1_16(0).HWFCSYCEN = tbl_chk1_16(1).HWFCSYCEN Then
                    blnSliceTarget = True
                Else
                    blnSliceTarget = False
                End If
            End If
        End If
        If blnSliceTarget = True Then
            ' 振替元のWF結晶面傾き横中心、横下限、横上限が全て設定なしかをチェック
            If tbl_chk1_16(0).HWFCSXCEN = -1 And _
               tbl_chk1_16(0).HWFCSXMIN = -1 And _
               tbl_chk1_16(0).HWFCSXMAX = -1 Then
                blnSliceTarget = True
            Else
                ' 振替先のWF結晶面傾き横中心、横下限、横上限が全て設定なしかをチェック
                If tbl_chk1_16(1).HWFCSXCEN = -1 And _
                   tbl_chk1_16(1).HWFCSXMIN = -1 And _
                   tbl_chk1_16(1).HWFCSXMAX = -1 Then
                    blnSliceTarget = True
                Else
                    ' WF結晶面傾き横中心仕様一致チェック
                    If tbl_chk1_16(0).HWFCSXCEN = tbl_chk1_16(1).HWFCSXCEN Then
                        blnSliceTarget = True
                    Else
                        blnSliceTarget = False
                    End If
                End If
            End If
        End If
    Else
        blnSliceTarget = False
    End If
    
    'ターゲット不一致(マルチブロックの場合は例外処理は実施しない)
    If blnSliceTarget = False Then
        ' 結晶面傾チェックNG
        funChkFurikae1_16 = 1
        iErr_Code = 1603
        sErr_Msg = "CHECK1-16,WF結晶面傾中心仕様が異なる為、組合せできません。"
        gsTbcmy028ErrCode = "00162"
        GoTo Apl_Exit
    End If
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae1_16 = 0 Then
        funChkFurikae1_16 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae1_16 = -4
    GoTo Apl_Exit

End Function
'Add End 2011/07/12 SMPK Nakamura 結晶面傾きチェック追加対応

'------------------------------------------------
' 結晶評価実績チェック
'------------------------------------------------

'概要      :指定された振替元品番から振替先品番への振り替えが、可能かどうかを結晶評価実績を元にチェックする。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sBlockId        ,I  ,String       :ﾌﾞﾛｯｸID
'          :tOld_Hinban     ,I  ,tFullHinban  :振替元品番(構造体)
'          :tNew_Hinban     ,I  ,tFullHinban  :振替先品番(構造体)
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :iSmpGetFlg      ,I  ,Integer      :ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ(0:ｻﾝﾌﾟﾙ指定なし, 1:ｻﾝﾌﾟﾙ指定あり)
'          :sSamplID1       ,I  ,String       :TOPｻﾝﾌﾟﾙID
'          :sSamplID2       ,I  ,String       :BOTｻﾝﾌﾟﾙID
'          :iKcnt           ,I  ,Integer      :工程連番(省略可)
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/19 新規作成　SB

Public Function funChkFurikae2_1(sProccd As String, sBlockId As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iErr_Code As Integer, sErr_Msg As String, _
                                 iSmpGetFlg As Integer, sSamplID1 As String, sSamplID2 As String, iKcnt As Integer) As Integer
    Dim sql         As String               'SQL全体
    Dim rs          As OraDynaset           'RecordSet
    Dim wBLKID()    As String               '総合判定対象ﾌﾞﾛｯｸID
    Dim cnt         As Integer              '複数ﾌﾞﾛｯｸｶｳﾝﾀ
    Dim TotalJudg   As Boolean              '総合判定結果
    Dim tb          As Integer              'Top/Botｶｳﾝﾀ
    Dim ks          As Integer              '検査項目ｶｳﾝﾀ
    Const MAXCNT    As Integer = 16         ' 最大件数
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae2_1 = 0
    
'    '------------------------------------------ CW760の場合、ﾌﾞﾛｯｸIDを取得 ------------------------------------------------
    If (left(sProccd, 4) = "CW76") Then
        sErr_Msg = "2-1 BLK-ID取得"
        sql = vbNullString
        sql = sql & "SELECT CRYNUMCA FROM XSDCA " & vbCrLf
        sql = sql & "WHERE  SXLIDCA = '" & sBlockId & "' AND " & vbCrLf
        sql = sql & "       LIVKCA  = '0' " & vbCrLf
    
        On Error GoTo db_Error
        'SQL文の実行
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        
        '該当データなし
        If rs.EOF Or rs.RecordCount = 0 Then GoTo db_Error
        
        ReDim wBLKID(rs.RecordCount)
        For cnt = 1 To rs.RecordCount
            If IsNull(rs("CRYNUMCA")) = False Then wBLKID(cnt) = rs("CRYNUMCA") Else wBLKID(cnt) = " "
            rs.MoveNext
        Next cnt
        Set rs = Nothing
    Else
        ReDim wBLKID(1)
        wBLKID(1) = sBlockId
    End If
    
    For cnt = 1 To UBound(wBLKID)
        '------------------------------------------ 結晶総合判定共通関数 ------------------------------------------------------
        '---------------------------- 2005/02/07 ffc)tanabe 追加 start --------------------------------
        '処理工程="CC600"の場合
        If (left(sProccd, 4) = "CC60") Then
        
            '反映データの合否判定を行う。
            If JudgChgFlg = "0" Then
            
                If iSmpGetFlg = 0 Then
                    If funCrySogoHantei(wBLKID(cnt), tNew_Hinban, TotalJudg, iErr_Code, sErr_Msg, typ_b, iSmpGetFlg) <> 0 Then GoTo Apl_down
                Else
                    If funCrySogoHantei(wBLKID(cnt), tNew_Hinban, TotalJudg, iErr_Code, sErr_Msg, typ_b, iSmpGetFlg, _
                                                                                CInt(sSamplID1), CInt(sSamplID2), iKcnt) <> 0 Then GoTo Apl_down
                End If
                
            '反映データの合否判定を行わない。
            Else
                ''==複数品番判定対応 20060501SMP桜井
                '--Before
''                If iSmpGetFlg = 0 Then
''                    If funCrySogoHantei2(wBLKID(cnt), tOld_Hinban, tNew_Hinban, TotalJudg, iErr_Code, sErr_Msg, typ_B, iSmpGetFlg) <> 0 Then GoTo Apl_down
''                Else
''                    If funCrySogoHantei2(wBLKID(cnt), tOld_Hinban, tNew_Hinban, TotalJudg, iErr_Code, sErr_Msg, typ_B, iSmpGetFlg, _
''                                                                                CInt(sSamplID1), CInt(sSamplID2), iKcnt) <> 0 Then GoTo Apl_down
''                End If
                '--<<
                If iSmpGetFlg = 0 Then
                    If funCrySogoHantei_CC600Multi(wBLKID(cnt), tOld_Hinban, tNew_Hinban, TotalJudg, iErr_Code, sErr_Msg, typ_b, iSmpGetFlg) <> 0 Then GoTo Apl_down
                Else
                    If funCrySogoHantei_CC600Multi(wBLKID(cnt), tOld_Hinban, tNew_Hinban, TotalJudg, iErr_Code, sErr_Msg, typ_b, iSmpGetFlg, _
                                                                                CInt(sSamplID1), CInt(sSamplID2), iKcnt) <> 0 Then GoTo Apl_down
                End If
                ''====================<<<<
            End If
        
        '処理工程="CC600"以外の場合
        Else
        
            If iSmpGetFlg = 0 Then
                If funCrySogoHantei(wBLKID(cnt), tNew_Hinban, TotalJudg, iErr_Code, sErr_Msg, typ_b, iSmpGetFlg) <> 0 Then GoTo Apl_down
            Else
                If funCrySogoHantei(wBLKID(cnt), tNew_Hinban, TotalJudg, iErr_Code, sErr_Msg, typ_b, iSmpGetFlg, _
                                                                            CInt(sSamplID1), CInt(sSamplID2), iKcnt) <> 0 Then GoTo Apl_down
            End If
        
        End If
        
        '---------------------------- 2005/02/07 ffc)tanabe 追加 end -----------------------------------
        
        If Not TotalJudg Then
            '振替不可の内容を取得(最初にNGとなった項目)
            With typ_b
                For tb = 1 To 2     'TOP/BOT 2回分ﾙｰﾌﾟ
                    If .OKNG(tb) = False Then
                        sErr_Msg = "RS-" & IIf(tb = 1, "TOP", "BOT") & "⇒NG"
                        Exit For
                    Else
                        For ks = 0 To MAXCNT     '検査項目最大件数分ﾙｰﾌﾟ
                            If .typ_rslt(tb, ks).OKNG = "NG" Then
                                sErr_Msg = .typ_rslt(tb, ks).NAIYO & "-" & IIf(tb = 1, "TOP", "BOT") & "⇒NG"
                                Exit For
                            End If
                            If (left(sProccd, 4) = "CC60") Then ''<<複数品番判定対応
                                ''検査をしない項目があるので途中で抜けてしまうのを回避
                                ''If .typ_rslt(tb, ks).OKNG = "" Then Exit For
                            Else
                                If .typ_rslt(tb, ks).OKNG = "" Then Exit For
                            End If
                        Next ks
                    End If
                Next tb
            End With
            
            funChkFurikae2_1 = 1
            iErr_Code = 2101
'            sErr_Msg = "CHECK2-1,結晶評価実績で不合格の為、振り替えできません。"
            sErr_Msg = "CHECK2-1,結晶評価実績,振替不可[" & wBLKID(cnt) & "](" & sErr_Msg & ")"
            GoTo Apl_Exit
        End If
    Next cnt

'------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function

db_Error:
    Set rs = Nothing
    If funChkFurikae2_1 = 0 Then
        funChkFurikae2_1 = -3
    End If
    GoTo Apl_Exit
    
Apl_down:
    funChkFurikae2_1 = -4
    GoTo Apl_Exit
    
End Function

'------------------------------------------------
' 認定炉チェック  2008/08/20 追加  Info.Kameda
'------------------------------------------------

'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sKeyID          ,I  ,String       :ﾌﾞﾛｯｸID、又は、SXL-ID
'          :tOld_Hinban     ,I  ,tFullHinban  :振替元品番(構造体)
'          :tNew_Hinban     ,I  ,tFullHinban  :振替先品番(構造体)
'          :iHinCnt         ,I  ,Integer      :複数品番カウント
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :

Public Function funChkFurikae2_3(sProccd As String, sBlockId As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iHinCnt As Integer, iErr_Code As Integer, sErr_Msg As String) As Integer
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim s
    Dim sBLIDedt    As String
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae2_3 = 0
    
    ReDim Preserve gNinteiro_Data(iHinCnt)
    '------------------------------------------ 判定データ取得(XODFA_1) ----------------------------------------
    '認定炉ID,改訂番号
    '製作条件Noより認定炉ID取得
    ' SQL作成
    sql = "SELECT IDFA1,REVFA1,CHK_SXL,CHK_WFC1,CHK_WFC2, "
    sql = sql & " TO_CHAR(SYN_DATE,'YYYY/MM/DD HH24:MI:SS') SDATE"
    sql = sql & " FROM XODFA_1 WHERE trim(HINBAN) = '" & left(tNew_Hinban.hinban, 3) & "' "     '2008/09/03 追加
    sql = sql & "                            and trim(MCNO) = (select trim(MCNO) from tbcme036 where hinban = '" & tNew_Hinban.hinban & "' "
    sql = sql & "                            and mnorevno = '" & tNew_Hinban.mnorevno & "' "
    sql = sql & "                            and factory = '" & tNew_Hinban.factory & "' "
    sql = sql & "                            and opecond = '" & tNew_Hinban.opecond & "') "
    sql = sql & " AND REMOVE = '0' "
    sql = sql & " AND FLAG = '1' "
    ' 実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなしは判定OK
    If rs.EOF Or rs.RecordCount = 0 Then GoTo Apl_Exit
    
    With gNinteiro_Data(iHinCnt)
        .ROID = rs("IDFA1")
        .REV = rs("REVFA1")
        .CHKSXL = rs("CHK_SXL")
        .CHKWFC1 = rs("CHK_WFC1")
        .CHKWFC2 = rs("CHK_WFC2")
        .SYNDAY = rs("SDATE")
        
        Set rs = Nothing
    
        '判定有無
        If sProccd = "CC600" Then
            If .CHKSXL = "0" Then GoTo Apl_Exit
        'ElseIf sProccd = "CW731" Then  <---------------- 判定無し
        '    If .CHKWFC1 = "0" Then GoTo Apl_Exit
        ElseIf sProccd = "CW750" Then
            If .CHKWFC2 = "0" Then GoTo Apl_Exit
        End If
    
        .JUDGRO = "-1"
    '------------------------------------------ 判定データ取得(XODFA_2) ----------------------------------------
    '区分,チャージ量(From),チャージ量(To)
    
        sql = "SELECT * FROM XODFA_2 WHERE IDFA2 ='" & .ROID & "' AND REVFA2='" & .REV & "' AND ROIDFA2 = '" & Mid(sBlockId, 1, 3) & "'"
        
        ' 実行
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        
        '該当データなし
        If rs.EOF Or rs.RecordCount = 0 Then '2008/08/28 修正
            funChkFurikae2_3 = 1
            iErr_Code = 2301
            sErr_Msg = "CHECK2-3,認定炉判定エラー、振替できません。"
            GoTo Apl_Exit
        End If
        
        'データセット
        .GOUKI = rs("ROIDFA2")                                            ' 号機№
        If IsNull(rs("KUBUNFA2")) = False Then .KUBUN = rs("KUBUNFA2")    ' 区分
        If IsNull(rs("FRCHGFA2")) = False Then .FRCHG = rs("FRCHGFA2")    ' ﾁｬｰｼﾞ量From
        If IsNull(rs("TOCHGFA2")) = False Then .TOCHG = rs("TOCHGFA2")    ' ﾁｬｰｼﾞ量To
        
        Set rs = Nothing
    '------------------------------------------ 判定データ取得(XSDC1) ------------------------------------------
    ''推定チャージ     2008/10/02 XSDC1-->TBCMH001変更 Kameda
        'sql = "SELECT SUICHARGE FROM XSDC1 "
        'sql = sql & "WHERE XTALC1 = '" & left(sBlockID, 9) & "000" & "' "
    
    ''引上げ指示チャージ    2008/12/04 変更 Kameda
        'sql = "SELECT CHARGE FROM TBCMH001 "
        'sql = sql & "WHERE substr(UPINDNO,1,7) = '" & left(sBlockID, 7) & "' "
        
        '1本引の場合は（⑨が０の場合）及びリチャージ（⑨が1,2,3～の場合）は
        'TBCMH001が9桁なので8桁目を0にして9桁でCHARGE項目を取ってくる｡
        'リチャージの場合は（⑨がA・B　の場合）
        'TBCMH001が9桁なので8桁目9桁目を0にしてCHARGE項目を取ってくる｡
        
        '2009/06/04 Kameda
        '⑨桁目がCに対応
        'If Mid(sBlockId, 9, 1) = "A" Or Mid(sBlockId, 9, 1) = "B" Then
        '    sBLIDedt = Mid(sBlockId, 1, 7) & "00"
        'Else
        '    sBLIDedt = Mid(sBlockId, 1, 7) & "0" & Mid(sBlockId, 9, 1)
        'End If
        
        '2009/12/25 Kameda 全ロット８桁目をゼロに
        'If IsNumeric(Mid(sBlockId, 9, 1)) = True Then
            sBLIDedt = Mid(sBlockId, 1, 7) & "0" & Mid(sBlockId, 9, 1)
        'Else
        '    sBLIDedt = Mid(sBlockId, 1, 7) & "00"
        'End If
        
        ''引上げ指示チャージ取得
        sql = "SELECT CHARGE FROM TBCMH001 "
        sql = sql & "WHERE UPINDNO = '" & sBLIDedt & "' "
        
        ' 実行
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        
        '該当データなし
        If rs.EOF Or rs.RecordCount = 0 Then  '2008/08/28 修正
            funChkFurikae2_3 = 1
            iErr_Code = 2301
            sErr_Msg = "CHECK2-3,認定炉判定エラー、振替できません。"
            GoTo Apl_Exit
        End If
        'データセット
        If IsNull(rs("CHARGE")) = False Then .SUICHG = rs("CHARGE")    ' 推定ﾁｬｰｼﾞ量
    
    '------------------------------------------ 判定 -----------------------------------------------------------
        '号機番号が登録されているかKUBUN=1
        If .KUBUN = 0 Then
            funChkFurikae2_3 = 1
            iErr_Code = 2301
            sErr_Msg = "CHECK2-3,認定炉判定エラー、振替できません。"
            GoTo Apl_Exit
        End If
        '推定ﾁｬｰｼﾞ量がﾁｬｰｼﾞFrom, To範囲内か
        If .FRCHG > .SUICHG Then
            funChkFurikae2_3 = 1
            iErr_Code = 2301
            sErr_Msg = "CHECK2-3,認定炉判定エラー、振替できません。"
            GoTo Apl_Exit
        End If
        If .TOCHG < .SUICHG Then
            funChkFurikae2_3 = 1
            iErr_Code = 2301
            sErr_Msg = "CHECK2-3,認定炉判定エラー、振替できません。"
            GoTo Apl_Exit
        End If
    
        .JUDGRO = "0"     '判定ＯＫ
    
    End With
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae2_3 = 0 Then
        funChkFurikae2_3 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae2_3 = -4
    GoTo Apl_Exit
    
End Function

'------------------------------------------------
' ＷＦＣ評価実績チェック
'------------------------------------------------

'概要      :指定された振替元品番から振替先品番への振り替えが、可能かどうかをＷＦＣ評価実績を元にチェックする。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sSXL_ID         ,I  ,String       :ﾌﾞﾛｯｸID
'          :tOld_Hinban     ,I  ,String       :振替元品番
'          :tNew_Hinban     ,I  ,String       :振替候補品番
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :iSmpGetFlg      ,I  ,Integer      :ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ(0:ｻﾝﾌﾟﾙ指定なし, 1:ｻﾝﾌﾟﾙ指定あり)
'          :sSamplID1       ,I  ,String       :TOPｻﾝﾌﾟﾙID
'          :sSamplID2       ,I  ,String       :BOTｻﾝﾌﾟﾙID
'          :iKcnt           ,I  ,Integer      :工程連番(省略可)
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/19 新規作成　SB

Public Function funChkFurikae3_1(sProccd As String, sSXL_ID As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iErr_Code As Integer, sErr_Msg As String, _
                                 iSmpGetFlg As Integer, sSamplID1 As String, sSamplID2 As String, iKcnt As Integer) As Integer
    
    Dim TotalJudg   As Boolean
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae3_1 = 0
    
'    '------------------------------------------ 結晶総合判定共通関数 ------------------------------------------------------
    If iSmpGetFlg = 0 Then
        If funWfcSogoHantei(sSXL_ID, tNew_Hinban, TotalJudg, iErr_Code, sErr_Msg, typ_CType, iSmpGetFlg) <> 0 Then GoTo Apl_down
    Else
        If funWfcSogoHantei(sSXL_ID, tNew_Hinban, TotalJudg, iErr_Code, sErr_Msg, typ_CType, iSmpGetFlg, _
                                                                                   sSamplID1, sSamplID2, iKcnt) <> 0 Then GoTo Apl_down
    End If
    
    If Not TotalJudg Then
        funChkFurikae3_1 = 1
        iErr_Code = 3101
        sErr_Msg = "CHECK3-1,WFC評価実績で不合格の為、振り替えできません。"
        GoTo Apl_Exit
    End If

'------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funChkFurikae3_1 = -4
    GoTo Apl_Exit
    
End Function

'------------------------------------------------
' Warp実績チェック
'------------------------------------------------

'概要      :指定された振替元品番から振替先品番への振り替えが、可能かどうかをWarp実績を元にチェックする。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :tOld_Hinban     ,I  ,String       :振替元品番
'          :tNew_Hinban     ,I  ,String       :振替候補品番
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :05/12/28 ooba

Public Function funChkFurikae3_2(tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iErr_Code As Integer, sErr_Msg As String) As Integer

    Dim i, j            As Integer
    Dim iPoint          As Integer              '品番境界のWarpﾃﾞｰﾀ算出用
    Dim iLoop           As Integer              'Warpﾃﾞｰﾀ配列のﾙｰﾌﾟ開始位置
    Dim iCntW           As Integer              'Warpﾃﾞｰﾀ数ｶｳﾝﾀ
    Dim dWarpMaxT       As Double               'Warp上限値
    Dim bWarpAllJudg    As Boolean              '全ﾃﾞｰﾀのWarp判定
    'Add 2010/03/30 Y.Hitomi Warpｴﾗｰ緩和対応
    Dim iWarpErrCount   As Integer              'Warpｴﾗｰ個数
    
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae3_2 = 0
    
    
    '振替ﾁｪｯｸ実施済の場合は処理を抜ける
    If tMapHinG.WARPFLG Then GoTo Apl_Exit
    
    '初期化
    iLoop = 1
    j = 1
    sErr_Msg = ""
    bWarpAllJudg = True
    iCntW = UBound(tWarpMeasG)
    'Add 2010/03/30 Y.Hitomi Warpｴﾗｰ緩和対応
    iWarpErrCount = 0
    
    'Warp仕様値の取得
    If funGetSiyou_Warp(tNew_Hinban, dWarpMaxT) = FUNCTION_RETURN_FAILURE Then
        sErr_Msg = "3-2 品番Warp仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond & ")"
        GoTo db_Error
    End If
    
    '品番境界のﾃﾞｰﾀ取得(TOP側)
    iPoint = 0
    For i = iLoop To UBound(tWarpInitG)
        'WarpﾃﾞｰﾀとﾌﾞﾛｯｸIDが一致
        If tMapHinG.BLOCKID = tWarpInitG(i).BLOCKID Then
            'Warpﾃﾞｰﾀの位置より上
            If tMapHinG.BLKSEQ_S < tWarpInitG(i).WAFID Then
                iPoint = i
                Exit For
            'Warpﾃﾞｰﾀの位置と同じ
            ElseIf tMapHinG.BLKSEQ_S = tWarpInitG(i).WAFID Then
                iPoint = 0
                Exit For
            'Warpﾃﾞｰﾀの位置より下
            ElseIf tMapHinG.BLKSEQ_S > tWarpInitG(i).WAFID Then
                '次のWarpﾃﾞｰﾀが無い(関連ﾌﾞﾛｯｸの最後)
                If i = UBound(tWarpInitG) Then
                    iPoint = i
                    Exit For
                '次のWarpﾃﾞｰﾀが無い(ﾌﾞﾛｯｸの最後)
                ElseIf tMapHinG.BLOCKID <> tWarpInitG(i + 1).BLOCKID Then
                    iPoint = i
                    Exit For
                '次のWarpﾃﾞｰﾀの位置より上
                ElseIf tMapHinG.BLKSEQ_S < tWarpInitG(i + 1).WAFID Then
                    '上下のWarp測定値を比較して厳しい(大きい)方を採用
                    iPoint = IIf(tWarpInitG(i).MEASDATA > tWarpInitG(i + 1).MEASDATA, i, i + 1)
                    Exit For
                End If
            End If
        End If
    Next i
    If iPoint > 0 Then
        'Warpﾃﾞｰﾀｾｯﾄ
        iCntW = iCntW + 1
        ReDim Preserve tWarpMeasG(iCntW)
        With tWarpMeasG(iCntW)
            .BLOCKID = tWarpInitG(iPoint).BLOCKID               'ﾌﾞﾛｯｸID
            .WAFID = tMapHinG.BLKSEQ_S                          'ｳｪﾊｰID
            .MEASDATA = tWarpInitG(iPoint).MEASDATA             '測定値
            .HIN = tMapHinG.HIN                                 '品番
            .max = dWarpMaxT                                    '仕様Max値
            'Warp判定
            .Judg = WfWarpJudg(.max, .MEASDATA)                 '判定
            If Not .Judg Then
                bWarpAllJudg = .Judg
                'Add 2010/03/30 Y.Hitomi Warpｴﾗｰ緩和対応
                iWarpErrCount = iWarpErrCount + 1
            End If
            .EXISTFLG = 0                                       '存在ﾌﾗｸﾞ(実ﾃﾞｰﾀ無)
        End With
        iLoop = iPoint
    End If
    
    For i = iLoop To UBound(tWarpInitG)
        'WFﾏｯﾌﾟ上の品番ﾃﾞｰﾀ範囲内にあれは処理する
        '存在ﾌﾗｸﾞ条件追加 07/03/16 ooba
        If tWarpInitG(i).EXISTFLG = 1 And _
           tWarpInitG(i).BLOCKID = tMapHinG.BLOCKID And _
           tWarpInitG(i).WAFID >= tMapHinG.BLKSEQ_S And _
           tWarpInitG(i).WAFID <= tMapHinG.BLKSEQ_E Then
            'Warpﾃﾞｰﾀｾｯﾄ
            iCntW = iCntW + 1
            ReDim Preserve tWarpMeasG(iCntW)
            With tWarpMeasG(iCntW)
                .BLOCKID = tWarpInitG(i).BLOCKID                'ﾌﾞﾛｯｸID
                .WAFID = tWarpInitG(i).WAFID                    'ｳｪﾊｰID
                .MEASDATA = tWarpInitG(i).MEASDATA              '測定値
                .HIN = tMapHinG.HIN                             '品番
                .max = dWarpMaxT                                '仕様Max値
                'Warp判定
                .Judg = WfWarpJudg(.max, .MEASDATA)             '判定
                If Not .Judg Then
                    bWarpAllJudg = .Judg
                    'Add 2010/03/30 Y.Hitomi Warpｴﾗｰ緩和対応
                    iWarpErrCount = iWarpErrCount + 1
                End If
                .EXISTFLG = 1                                   '存在ﾌﾗｸﾞ(実ﾃﾞｰﾀ有)

            End With
            j = i
        End If
    Next i
    iLoop = j
    
    '品番境界のﾃﾞｰﾀ取得(BOT側)
    iPoint = 0
    For i = iLoop To UBound(tWarpInitG)
        'WarpﾃﾞｰﾀとﾌﾞﾛｯｸIDが一致
        If tMapHinG.BLOCKID = tWarpInitG(i).BLOCKID Then
            'Warpﾃﾞｰﾀの位置より上
            If tMapHinG.BLKSEQ_E < tWarpInitG(i).WAFID Then
                iPoint = i
                Exit For
            'Warpﾃﾞｰﾀの位置と同じ
            ElseIf tMapHinG.BLKSEQ_E = tWarpInitG(i).WAFID Then
                iPoint = 0
                Exit For
            'Warpﾃﾞｰﾀの位置より下
            ElseIf tMapHinG.BLKSEQ_E > tWarpInitG(i).WAFID Then
                '次のWarpﾃﾞｰﾀが無い(関連ﾌﾞﾛｯｸの最後)
                If i = UBound(tWarpInitG) Then
                    iPoint = i
                    Exit For
                '次のWarpﾃﾞｰﾀが無い(ﾌﾞﾛｯｸの最後)
                ElseIf tMapHinG.BLOCKID <> tWarpInitG(i + 1).BLOCKID Then
                    iPoint = i
                    Exit For
                '次のWarpﾃﾞｰﾀの位置より上
                ElseIf tMapHinG.BLKSEQ_E < tWarpInitG(i + 1).WAFID Then
                    '上下のWarp測定値を比較して厳しい(大きい)方を採用
                    iPoint = IIf(tWarpInitG(i).MEASDATA > tWarpInitG(i + 1).MEASDATA, i, i + 1)
                    Exit For
                End If
            End If
        End If
    Next i
    If iPoint > 0 Then
        'Warpﾃﾞｰﾀｾｯﾄ
        iCntW = iCntW + 1
        ReDim Preserve tWarpMeasG(iCntW)
        With tWarpMeasG(iCntW)
            .BLOCKID = tWarpInitG(iPoint).BLOCKID               'ﾌﾞﾛｯｸID
            .WAFID = tMapHinG.BLKSEQ_E                          'ｳｪﾊｰID
            .MEASDATA = tWarpInitG(iPoint).MEASDATA             '測定値
            .HIN = tMapHinG.HIN                                 '品番
            .max = dWarpMaxT                                    '仕様Max値
            'Warp判定
            .Judg = WfWarpJudg(.max, .MEASDATA)                 '判定
            If Not .Judg Then
                bWarpAllJudg = .Judg
                'Add 2010/03/30 Y.Hitomi Warpｴﾗｰ緩和対応
                iWarpErrCount = iWarpErrCount + 1
            End If
            
            .EXISTFLG = 0                                       '存在ﾌﾗｸﾞ(実ﾃﾞｰﾀ無)
        End With
    End If
    
    tMapHinG.WARPFLG = True     'Warp振替ﾁｪｯｸ済
    
    'Change 2010/03/30 Y.Hitomi Warpｴﾗｰ緩和対応
    'Change 2010/05/31 Y.Hitomi Warpｴﾗｰ緩和対応(9→10枚）
'    If Not bWarpAllJudg Then
    If iWarpErrCount > 10 Then
        funChkFurikae3_2 = 1
        iErr_Code = 3201
        sErr_Msg = "CHECK3-2,Warp実績で不合格の為、振り替えできません。"
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00128"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    
'------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:

    Exit Function

db_Error:
    If funChkFurikae3_2 = 0 Then
        funChkFurikae3_2 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae3_2 = -4
    GoTo Apl_Exit

End Function

'------------------------------------------------
' 合成角度実績チェック
'------------------------------------------------

'概要      :指定された振替元品番から振替先品番への振り替えが、可能かどうかを合成角度実績を元にチェックする。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :tOld_Hinban     ,I  ,String       :振替元品番
'          :tNew_Hinban     ,I  ,String       :振替候補品番
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :05/12/29 ooba

Public Function funChkFurikae3_3(tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iErr_Code As Integer, sErr_Msg As String) As Integer

    Dim i               As Integer
    Dim iCntK           As Integer              '合成角度ﾃﾞｰﾀ数ｶｳﾝﾀ
    'Add Start 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応
    Dim iCntKX          As Integer              '横(X)角度ﾃﾞｰﾀ数ｶｳﾝﾀ
    Dim iCntKY          As Integer              '縦(Y)角度ﾃﾞｰﾀ数ｶｳﾝﾀ
    'Add End 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応
    Dim dKakuMinT       As Double               '合成角度下限値
    Dim dKakuMaxT       As Double               '合成角度上限値
    Dim bKakuAllJudg    As Boolean              '全ﾃﾞｰﾀの合成角度判定
    
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae3_3 = 0
    
    '振替ﾁｪｯｸ実施済の場合は処理を抜ける
    If tMapHinG.KAKUFLG Then GoTo Apl_Exit
    
    '初期化
    sErr_Msg = ""
    bKakuAllJudg = True
    iCntK = UBound(tKakuMeasG)
    'Add Start 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応
    iCntKX = UBound(tKakuXMeasG)
    iCntKY = UBound(tKakuYMeasG)
    'Add End 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応
    
'Chg Start 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応
'    '合成角度仕様値の取得
'    If funGetSiyou_Kaku(tNew_Hinban, dKakuMinT, dKakuMaxT) = FUNCTION_RETURN_FAILURE Then
'        sErr_Msg = "3-3 品番合成角度仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.FACTORY & tNew_Hinban.OPECOND & ")"
'        GoTo db_Error
'    End If
'
'    For i = 1 To UBound(tKakuInitG)
'        If tKakuInitG(i).BLOCKID = tMapHinG.BLOCKID Then
'            '合成角度ﾃﾞｰﾀｾｯﾄ
'            iCntK = iCntK + 1
'            ReDim Preserve tKakuMeasG(iCntK)
'            With tKakuMeasG(iCntK)
'                .BLOCKID = tKakuInitG(i).BLOCKID                'ﾌﾞﾛｯｸID
'                .MEASDATA = tKakuInitG(i).MEASDATA              '測定値
'                .hin = tMapHinG.hin                             '品番
'                .Min = dKakuMinT                                '仕様Min値
'                .max = dKakuMaxT                                '仕様Max値
'                '合成角度判定
'                .Judg = WfKakuJudg(.Min, .max, .MEASDATA)       '判定
'                If Not .Judg Then bKakuAllJudg = .Judg
'            End With
'        End If
'    Next i
    '合成角度仕様値の取得
    If funGetSiyou_WFXtalInclination("XY", tNew_Hinban, dKakuMinT, dKakuMaxT) = FUNCTION_RETURN_FAILURE Then
        sErr_Msg = "3-3 品番合成角度仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond & ")"
        GoTo db_Error
    End If
    
    For i = 1 To UBound(tKakuInitG)
        If tKakuInitG(i).BLOCKID = tMapHinG.BLOCKID Then
            '合成角度ﾃﾞｰﾀｾｯﾄ
            iCntK = iCntK + 1
            ReDim Preserve tKakuMeasG(iCntK)
            With tKakuMeasG(iCntK)
                .BLOCKID = tKakuInitG(i).BLOCKID                'ﾌﾞﾛｯｸID
                .MEASDATA = tKakuInitG(i).MEASDATA              '測定値
                .HIN = tMapHinG.HIN                             '品番
                .Min = dKakuMinT                                '仕様Min値
                .max = dKakuMaxT                                '仕様Max値
                '合成角度判定
                .Judg = WfKakuJudg(.Min, .max, .MEASDATA)       '判定
                If Not .Judg Then bKakuAllJudg = .Judg
            End With
        End If
    Next i
    '横(X)角度仕様値の取得
    If funGetSiyou_WFXtalInclination("X", tNew_Hinban, dKakuMinT, dKakuMaxT) = FUNCTION_RETURN_FAILURE Then
        sErr_Msg = "3-3 品番横(X)角度仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond & ")"
        GoTo db_Error
    End If

    For i = 1 To UBound(tKakuXInitG)
        If tKakuXInitG(i).BLOCKID = tMapHinG.BLOCKID Then
            '横(X)角度ﾃﾞｰﾀｾｯﾄ
            iCntKX = iCntKX + 1
            ReDim Preserve tKakuXMeasG(iCntKX)
            With tKakuXMeasG(iCntKX)
                .BLOCKID = tKakuXInitG(i).BLOCKID               'ﾌﾞﾛｯｸID
                .MEASDATA = tKakuXInitG(i).MEASDATA             '測定値
                .HIN = tMapHinG.HIN                             '品番
                .Min = dKakuMinT                                '仕様Min値
                .max = dKakuMaxT                                '仕様Max値
                '横(X)角度判定
                .Judg = WfKakuJudg(.Min, .max, .MEASDATA)       '判定
                If Not .Judg Then bKakuAllJudg = .Judg
            End With
        End If
    Next i
    '縦(Y)角度仕様値の取得
    If funGetSiyou_WFXtalInclination("Y", tNew_Hinban, dKakuMinT, dKakuMaxT) = FUNCTION_RETURN_FAILURE Then
        sErr_Msg = "3-3 品番横(Y)角度仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond & ")"
        GoTo db_Error
    End If

    For i = 1 To UBound(tKakuYInitG)
        If tKakuYInitG(i).BLOCKID = tMapHinG.BLOCKID Then
            '縦(Y)角度ﾃﾞｰﾀｾｯﾄ
            iCntKY = iCntKY + 1
            ReDim Preserve tKakuYMeasG(iCntKY)
            With tKakuYMeasG(iCntKY)
                .BLOCKID = tKakuYInitG(i).BLOCKID               'ﾌﾞﾛｯｸID
                .MEASDATA = tKakuYInitG(i).MEASDATA             '測定値
                .HIN = tMapHinG.HIN                             '品番
                .Min = dKakuMinT                                '仕様Min値
                .max = dKakuMaxT                                '仕様Max値
                '縦(Y)角度判定
                .Judg = WfKakuJudg(.Min, .max, .MEASDATA)       '判定
                If Not .Judg Then bKakuAllJudg = .Judg
            End With
        End If
    Next i
'Chg End 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応
    
    tMapHinG.KAKUFLG = True     '合成角度振替ﾁｪｯｸ済
    
    If Not bKakuAllJudg Then
        funChkFurikae3_3 = 1
        iErr_Code = 3301
'Chg Start 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応
'        sErr_Msg = "CHECK3-3,合成角度実績で不合格の為、振り替えできません。"
        sErr_Msg = "CHECK3-3,X線実績で不合格の為、振り替えできません。"
'Chg End 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応
'--------------- 2008/07/25 INSERT START  By Systech ---------------
        gsTbcmy028ErrCode = "00129"
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
        GoTo Apl_Exit
    End If
    
'------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:

    Exit Function

db_Error:
    If funChkFurikae3_3 = 0 Then
        funChkFurikae3_3 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae3_3 = -4
    GoTo Apl_Exit

End Function

'------------------------------------------------
' ＷＦＣ評価実績(エピ)チェック
'------------------------------------------------

'概要      :指定された振替元品番から振替先品番への振り替えが、可能かどうかをＷＦＣ評価実績(エピ)を元にチェックする。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sSXL_ID         ,I  ,String       :ﾌﾞﾛｯｸID
'          :tOld_Hinban     ,I  ,String       :振替元品番
'          :tNew_Hinban     ,I  ,String       :振替候補品番
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :iSmpGetFlg      ,I  ,Integer      :ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ(0:ｻﾝﾌﾟﾙ指定なし, 1:ｻﾝﾌﾟﾙ指定あり)
'          :sSamplID1       ,I  ,String       :TOPｻﾝﾌﾟﾙID
'          :sSamplID2       ,I  ,String       :BOTｻﾝﾌﾟﾙID
'          :iKcnt           ,I  ,Integer      :工程連番(省略可)
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :2006/08/15 新規作成 エピ先行評価追加対応 SMP)kondoh

Public Function funChkFurikae3_4(sProccd As String, sSXL_ID As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iErr_Code As Integer, sErr_Msg As String, _
                                 iSmpGetFlg As Integer, sSamplID1 As String, sSamplID2 As String, iKcnt As Integer) As Integer
    
    Dim TotalJudg   As Boolean
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae3_4 = 0
    
'    '------------------------------------------ 結晶総合判定共通関数 ------------------------------------------------------
    If iSmpGetFlg = 0 Then
        If funWfcSogoHantei_EP(sSXL_ID, tNew_Hinban, TotalJudg, iErr_Code, sErr_Msg, typ_CType, typ_CType_EP, iSmpGetFlg) <> 0 Then GoTo Apl_down
    Else
        If funWfcSogoHantei_EP(sSXL_ID, tNew_Hinban, TotalJudg, iErr_Code, sErr_Msg, typ_CType, typ_CType_EP, iSmpGetFlg, _
                                                                                   sSamplID1, sSamplID2, iKcnt) <> 0 Then GoTo Apl_down
    End If
    
    If Not TotalJudg Then
        funChkFurikae3_4 = 1
        iErr_Code = 3401
        sErr_Msg = "CHECK3-4,WFC評価実績(エピ)で不合格の為、振り替えできません。"
        GoTo Apl_Exit
    End If

'------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funChkFurikae3_4 = -4
    GoTo Apl_Exit
    
End Function

'Add Start 2011/04/25 SMPK Miyata
'------------------------------------------------
' ＷＦＣ中間抜試実績チェック
'------------------------------------------------

'概要      :指定された振替元品番から振替先品番への振り替えが、可能かどうかを中間抜試実績を元にチェックする。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sSXL_ID         ,I  ,String       :ﾌﾞﾛｯｸID
'          :tOld_Hinban     ,I  ,String       :振替元品番
'          :tNew_Hinban     ,I  ,String       :振替候補品番
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :iSmpGetFlg      ,I  ,Integer      :ｻﾝﾌﾟﾙ管理取得ﾌﾗｸﾞ(0:ｻﾝﾌﾟﾙ指定なし, 1:ｻﾝﾌﾟﾙ指定あり)
'          :sSamplID1       ,I  ,String       :TOPｻﾝﾌﾟﾙID
'          :sSamplID2       ,I  ,String       :BOTｻﾝﾌﾟﾙID
'          :iKcnt           ,I  ,Integer      :工程連番(省略可)
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :

Public Function funChkFurikae3_5(sProccd As String, sSXL_ID As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iErr_Code As Integer, sErr_Msg As String, _
                                 iSmpGetFlg As Integer, sSamplID1 As String, sSamplID2 As String, iKcnt As Integer) As Integer

    Dim TotalJudg   As Boolean
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae3_5 = 0
    
'------------------------------------------ 結晶総合判定共通関数 ------------------------------------------------------
    If iSmpGetFlg = 0 Then
        If funWfcMidleHantei(sSXL_ID, tNew_Hinban, TotalJudg, iErr_Code, sErr_Msg, typ_CType, iSmpGetFlg) <> 0 Then GoTo Apl_down
    Else
        If funWfcMidleHantei(sSXL_ID, tNew_Hinban, TotalJudg, iErr_Code, sErr_Msg, typ_CType, iSmpGetFlg, _
                                                                                   sSamplID1, sSamplID2, iKcnt) <> 0 Then GoTo Apl_down
    End If
    
    If Not TotalJudg Then
        funChkFurikae3_5 = 1
        iErr_Code = 3101
        sErr_Msg = "CHECK3-5,WFC中間抜試実績チェックで不合格の為、振り替えできません。"
        GoTo Apl_Exit
    End If

'------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funChkFurikae3_5 = -4
    GoTo Apl_Exit
    
End Function
'Add End   2011/04/25 SMPK Miyata


'
'概要      :振替元品番と振替先品番の結晶評価項目仕様をチェックする。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sKeyID          ,I  ,String       :ﾌﾞﾛｯｸID、又は、SXL-ID
'          :tOld_Hinban     ,I  ,tFullHinban  :振替元品番(構造体)
'          :tNew_Hinban     ,I  ,tFullHinban  :振替先品番(構造体)
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :iELCs_Flg       ,O  ,Integer      :0 ･･･ 1-4全項目チェック
'                                              1 ･･･ 1-4(Cs,EPD,LT)のみチェック
'                                              2 ･･･ 1-4(Cs,EPD,LT)以外チェック
'                                              3 ･･･ 1-4(Cs)のみチェック
'                                              4 ･･･ 1-4(EPD)のみチェック
'                                              5 ･･･ 1-4(LT)のみチェック
'
'     funChkFurikae1_4を流用
'
''    チェック２－２
''    振替元と振替先のCOSF3仕様チェック
''    結晶評価項目COSF3の仕様チェックを行う｡
''                    元品番
''                    H   S   その他
''      先品番    H   ○  ○  ×
''                S   ○  ○  ×           ○ ： 振替OK
''            その他  ○  ○  ○           × ： 振替NG
''    仕様 (COSF3フラグ)
''      テーブル名          テーブル            カラム
''      結晶内側管理        TBCME036            COSF3FLAG
''
'履歴      :2008/04/20 新規作成　青柳
'Add Start 2010/12/23 SMPK A.Nagamine
' CLESTA評価対応(Cu-deco) : C-OSF3判定テーブルのキー変更(Class 14 -> 22), Cu-deco(C, CJ, CJ(LT), CJ2)の仕様判定追加
'Add End   2010/12/23 SMPK A.Nagamine
'
Public Function funChkFurikae2_2(sProccd As String, sBlockId As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iErr_Code As Integer, sErr_Msg As String, Optional iELCs_Flg As Integer = 0) As Integer

    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql As String               'SQL全体
    Dim rs  As OraDynaset           'RecordSet
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae2_2 = 0
    
    '------------------------------------------ 振替元品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "2-2 振替元品番仕様取得(" & tOld_Hinban.hinban & Format(tOld_Hinban.mnorevno, "00") & tOld_Hinban.factory & tOld_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
'Add Start 2010/12/23 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : C-OSF3判定テーブルのキー変更(Class 14 -> 22), Cu-deco(C, CJ, CJ(LT), CJ2)の仕様判定追加
    'sql = sql & "SELECT E018.HSXRHWYS,E019.HSXONHWS,  E019.HSXONSPT,E019.HSXONSPI,E019.HSXONKWY,E020.HSXOF1HS,E020.HSXOF1SH,  E020.HSXOF1ST,E020.HSXOF1SR,  E020.HSXOF1NS,E020.HSXOF1SZ,   " & vbCrLf
    'sql = sql & "       E020.HSXOF1ET,E020.HSXOSF1PTK,E020.HSXOF2HS,E020.HSXOF2SH,E020.HSXOF2ST,E020.HSXOF2SR,E020.HSXOF2NS,  E020.HSXOF2SZ,  E020.HSXOF2ET,E020.HSXOSF2PTK, " & vbCrLf
    'sql = sql & "       E020.HSXOF3HS,E020.HSXOF3SH,  E020.HSXOF3ST,E020.HSXOF3SR,E020.HSXOF3NS,E020.HSXOF3SZ,  E020.HSXOF3ET,E020.HSXOSF3PTK,E020.HSXOF4HS,E020.HSXOF4SH,   " & vbCrLf
    'sql = sql & "       E020.HSXOF4ST,E020.HSXOF4SR,  E020.HSXOF4NS,E020.HSXOF4SZ,E020.HSXOF4ET,E020.HSXOSF4PTK,E020.HSXBM1HS,E020.HSXBM1SH,  E020.HSXBM1ST,E020.HSXBM1SR,   " & vbCrLf
    'sql = sql & "       E020.HSXBM1NS,E020.HSXBM1SZ,  E020.HSXBM1ET,E020.HSXBM2HS,E020.HSXBM2SH,E020.HSXBM2ST,  E020.HSXBM2SR,E020.HSXBM2NS,  E020.HSXBM2SZ,E020.HSXBM2ET,   " & vbCrLf
    'sql = sql & "       E020.HSXBM3HS,E020.HSXBM3SH,  E020.HSXBM3ST,E020.HSXBM3SR,E020.HSXBM3NS,E020.HSXBM3SZ,  E020.HSXBM3ET,E019.HSXTMMAX,  E019.HSXLTHWS,E019.HSXCNHWS,   " & vbCrLf
    '
    ''C－OSF3判定機能 ---
    'sql = sql & "       E019.HSXCNKWY,E020.HSXDENHS,  E020.HSXDENMN,E020.HSXDENMX,E020.HSXDVDHS,E020.HSXDVDMNN,  E020.HSXDVDMXN,E020.HSXLDLHS,  E020.HSXLDLMN,E020.HSXLDLMX,E036.HSXGDLINE,E036.COSF3FLAG " & vbCrLf
    
    sql = sql & "SELECT E020.HSXOF4HS,  E020.HSXOF4SH,  E020.HSXOF4ST,  E020.HSXOF4SR,  E020.HSXOF4NS,  " & vbCrLf
    sql = sql & "       E020.HSXOF4SZ,  E020.HSXOF4ET,  E020.HSXOSF4PTK,E020.HSXBM1NS,  E020.HSXBM1SZ,  " & vbCrLf

    'C－OSF3判定機能  ---
    sql = sql & "       E036.COSF3FLAG " & vbCrLf
'Add End   2010/12/23 SMPK A.Nagamine
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    sql = sql & "       ,NVL(E036.HSXDKTMP,' ') HSXDKTMP " & vbCrLf
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    
'Add Start 2010/12/23 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : C-OSF3判定テーブルのキー変更(Class 14 -> 22), Cu-deco(C, CJ, CJ(LT), CJ2)の仕様判定追加
    
    sql = sql & ",      E020.HSXCOSF3HS,E020.HSXCOSF3SH,E020.HSXCOSF3ST,E020.HSXCOSF3SR,E020.HSXCOSF3NS," & vbCrLf
    sql = sql & "       E020.HSXCOSF3SZ,E036.HSXCOSF3ET,E020.HSXCOSF3PK,                                " & vbCrLf
    sql = sql & "       E020.HSXCPK,    E020.HSXCSZ,    E020.HSXCHT,    E020.HSXCHS,    E020.HSXCJPK,   " & vbCrLf
    sql = sql & "       E020.HSXCJNS,   E020.HSXCJHT,   E020.HSXCJHS,   E020.HSXCJLTPK, E020.HSXCJLTNS, " & vbCrLf
    sql = sql & "       E020.HSXCJLTHT, E020.HSXCJLTHS, E020.HSXCJ2PK,  E020.HSXCJ2NS,  E020.HSXCJ2HT,  " & vbCrLf
    sql = sql & "       E020.HSXCJ2HS,  E036.HSXCJLTBND " & vbCrLf
    
'Add End   2010/12/23 SMPK A.Nagamine
    
    sql = sql & "FROM   TBCME020 E020,TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E020.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E020.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E020.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E020.OPECOND   =   '" & tOld_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E036.HINBAN    =   '" & tOld_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tOld_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tOld_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tOld_Hinban.opecond & "'     " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    Erase tbl_chk1_4
    With tbl_chk1_4(0)
    
        ''C－OSF3判定機能   ---
        If IsNull(rs("COSF3FLAG")) = False Then .HSXOF4HS = rs("COSF3FLAG") Else .HSXOF4HS = " "            'C-OSF3ﾌﾗｸﾞ
    
'--------------- 2008/08/25 INSERT START  By Systech ---------------
        .HSXDKTMP = rs("HSXDKTMP")
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
        
        'OSF4
    'Add Start 2010/12/23 SMPK A.Nagamine       : CLESTA評価対応(Cu-deco) : C-OSF3判定テーブルのキー変更(Class 14 -> 22), Cu-deco(C, CJ, CJ(LT), CJ2)の仕様判定追加
        'If IsNull(rs("HSXCOSF3HS")) = False Then .HSXOF4HS = rs("HSXCOSF3HS") Else .HSXOF4HS = " "              'C-OSF3 保証方法_処 2010/12/24 Add
        
        'If IsNull(rs("HSXOF4SH")) = False Then .HSXOF4SH = rs("HSXOF4SH") Else .HSXOF4SH = " "              '測定位置_方
        'If IsNull(rs("HSXOF4ST")) = False Then .HSXOF4ST = rs("HSXOF4ST") Else .HSXOF4ST = " "              '測定位置_点
        'If IsNull(rs("HSXOF4SR")) = False Then .HSXOF4SR = rs("HSXOF4SR") Else .HSXOF4SR = " "              '測定位置_領
        'If IsNull(rs("HSXOF4NS")) = False Then .HSXOF4NS = rs("HSXOF4NS") Else .HSXOF4NS = " "              '熱処理法
        'If IsNull(rs("HSXOF4SZ")) = False Then .HSXOF4SZ = rs("HSXOF4SZ") Else .HSXOF4SZ = " "              '測定条件
        'If IsNull(rs("HSXOF4ET")) = False Then .HSXOF4ET = rs("HSXOF4ET") Else .HSXOF4ET = 0                '選択ET代
        'If IsNull(rs("HSXOSF4PTK")) = False Then .HSXOSF4PTK = rs("HSXOSF4PTK") Else .HSXOSF4PTK = "4"      'パターン区分
        
        If IsNull(rs("HSXCOSF3SH")) = False Then .HSXOF4SH = rs("HSXCOSF3SH") Else .HSXOF4SH = " "              'C-OSF3 測定位置_方     2010/12/24 Add
        If IsNull(rs("HSXCOSF3ST")) = False Then .HSXOF4ST = rs("HSXCOSF3ST") Else .HSXOF4ST = " "              'C-OSF3 測定位置_点     2010/12/24 Add
        If IsNull(rs("HSXCOSF3SR")) = False Then .HSXOF4SR = rs("HSXCOSF3SR") Else .HSXOF4SR = " "              'C-OSF3 測定位置_領     2010/12/24 Add
        If IsNull(rs("HSXCOSF3NS")) = False Then .HSXOF4NS = rs("HSXCOSF3NS") Else .HSXOF4NS = " "              'C-OSF3 熱処理法        2010/12/24 Add
        If IsNull(rs("HSXCOSF3SZ")) = False Then .HSXOF4SZ = rs("HSXCOSF3SZ") Else .HSXOF4SZ = " "              'C-OSF3 測定条件        2010/12/24 Add
        If IsNull(rs("HSXCOSF3ET")) = False Then .HSXOF4ET = rs("HSXCOSF3ET") Else .HSXOF4ET = 0                'C-OSF3 選択ET代        2010/12/24 Add
        If IsNull(rs("HSXCOSF3PK")) = False Then .HSXOSF4PTK = rs("HSXCOSF3PK") Else .HSXOSF4PTK = "4"          'C-OSF3 パターン区分    2010/12/24 Add
        
        If IsNull(rs("HSXCPK")) = False Then .HSXCPK = rs("HSXCPK") Else .HSXCPK = " "                  '/* 品ＳＸＣパターン区分 */
        If IsNull(rs("HSXCSZ")) = False Then .HSXCSZ = rs("HSXCSZ") Else .HSXCSZ = " "                  '/* 品ＳＸＣ測定条件     */
        If IsNull(rs("HSXCHT")) = False Then .HSXCHT = rs("HSXCHT") Else .HSXCHT = " "                  '/* 品ＳＸＣ保証方法＿対 */
        If IsNull(rs("HSXCHS")) = False Then .HSXCHS = rs("HSXCHS") Else .HSXCHS = " "                  '/* 品ＳＸＣ保証方法＿処 */
        
        If IsNull(rs("HSXCJPK")) = False Then .HSXCJPK = rs("HSXCJPK") Else .HSXCJPK = " "              '/* 品ＳＸＣＪパターン区分 */
        If IsNull(rs("HSXCJNS")) = False Then .HSXCJNS = rs("HSXCJNS") Else .HSXCJNS = "  "             '/* 品ＳＸＣＪ熱処理法     */
        If IsNull(rs("HSXCJHT")) = False Then .HSXCJHT = rs("HSXCJHT") Else .HSXCJHT = " "              '/* 品ＳＸＣＪ保証方法＿対 */
        If IsNull(rs("HSXCJHS")) = False Then .HSXCJHS = rs("HSXCJHS") Else .HSXCJHS = " "              '/* 品ＳＸＣＪ保証方法＿処 */
        
        If IsNull(rs("HSXCJLTPK")) = False Then .HSXCJLTPK = rs("HSXCJLTPK") Else .HSXCJLTPK = " "      '/* 品ＳＸＣＪＬＴパターン区分 */
        If IsNull(rs("HSXCJLTNS")) = False Then .HSXCJLTNS = rs("HSXCJLTNS") Else .HSXCJLTNS = "  "     '/* 品ＳＸＣＪＬＴ熱処理法     */
        If IsNull(rs("HSXCJLTHT")) = False Then .HSXCJLTHT = rs("HSXCJLTHT") Else .HSXCJLTHT = " "      '/* 品ＳＸＣＪＬＴ保証方法＿対 */
        If IsNull(rs("HSXCJLTHS")) = False Then .HSXCJLTHS = rs("HSXCJLTHS") Else .HSXCJLTHS = " "      '/* 品ＳＸＣＪＬＴ保証方法＿処 */
        
        If IsNull(rs("HSXCJ2PK")) = False Then .HSXCJ2PK = rs("HSXCJ2PK") Else .HSXCJ2PK = " "          '/* 品ＳＸＣＪ２パターン区分 */
        If IsNull(rs("HSXCJ2NS")) = False Then .HSXCJ2NS = rs("HSXCJ2NS") Else .HSXCJ2NS = "  "         '/* 品ＳＸＣＪ２熱処理法     */
        If IsNull(rs("HSXCJ2HT")) = False Then .HSXCJ2HT = rs("HSXCJ2HT") Else .HSXCJ2HT = " "          '/* 品ＳＸＣＪ２保証方法＿対 */
        If IsNull(rs("HSXCJ2HS")) = False Then .HSXCJ2HS = rs("HSXCJ2HS") Else .HSXCJ2HS = " "          '/* 品ＳＸＣＪ２保証方法＿処 */
        
        If IsNull(rs("HSXCJLTBND")) = False Then .HSXCJLTBND = rs("HSXCJLTBND") Else .HSXCJLTBND = 0    '/* 品SXL/CJLTバンド幅 Number(3,0) */
        
    'Add End 2010/12/23 SMPK A.Nagamine
    
    End With
    
    Set rs = Nothing
    '------------------------------------------ 振替先品種仕様データ取得 ------------------------------------------------------
    sErr_Msg = "2-2 振替先品番仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
'Add Start 2010/12/23 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : C-OSF3判定テーブルのキー変更(Class 14 -> 22), Cu-deco(C, CJ, CJ(LT), CJ2)の仕様判定追加
    'sql = sql & "SELECT E018.HSXRHWYS,E019.HSXONHWS,  E019.HSXONSPT,E019.HSXONSPI,E019.HSXONKWY,E020.HSXOF1HS,E020.HSXOF1SH,  E020.HSXOF1ST,E020.HSXOF1SR,  E020.HSXOF1NS,E020.HSXOF1SZ,   " & vbCrLf
    'sql = sql & "       E020.HSXOF1ET,E020.HSXOSF1PTK,E020.HSXOF2HS,E020.HSXOF2SH,E020.HSXOF2ST,E020.HSXOF2SR,E020.HSXOF2NS,  E020.HSXOF2SZ,  E020.HSXOF2ET,E020.HSXOSF2PTK, " & vbCrLf
    'sql = sql & "       E020.HSXOF3HS,E020.HSXOF3SH,  E020.HSXOF3ST,E020.HSXOF3SR,E020.HSXOF3NS,E020.HSXOF3SZ,  E020.HSXOF3ET,E020.HSXOSF3PTK,E020.HSXOF4HS,E020.HSXOF4SH,   " & vbCrLf
    'sql = sql & "       E020.HSXOF4ST,E020.HSXOF4SR,  E020.HSXOF4NS,E020.HSXOF4SZ,E020.HSXOF4ET,E020.HSXOSF4PTK,E020.HSXBM1HS,E020.HSXBM1SH,  E020.HSXBM1ST,E020.HSXBM1SR,   " & vbCrLf
    'sql = sql & "       E020.HSXBM1NS,E020.HSXBM1SZ,  E020.HSXBM1ET,E020.HSXBM2HS,E020.HSXBM2SH,E020.HSXBM2ST,  E020.HSXBM2SR,E020.HSXBM2NS,  E020.HSXBM2SZ,E020.HSXBM2ET,   " & vbCrLf
    'sql = sql & "       E020.HSXBM3HS,E020.HSXBM3SH,  E020.HSXBM3ST,E020.HSXBM3SR,E020.HSXBM3NS,E020.HSXBM3SZ,  E020.HSXBM3ET,E019.HSXTMMAX,  E019.HSXLTHWS,E019.HSXCNHWS,   " & vbCrLf
    '
    ''C－OSF3判定機能  ---
    'sql = sql & "       E019.HSXCNKWY,E020.HSXDENHS,  E020.HSXDENMN,E020.HSXDENMX,E020.HSXDVDHS,E020.HSXDVDMNN, E020.HSXDVDMXN,E020.HSXLDLHS,  E020.HSXLDLMN,E020.HSXLDLMX,E036.HSXGDLINE,E036.COSF3FLAG " & vbCrLf

    sql = sql & "SELECT E020.HSXOF4HS,  E020.HSXOF4SH,  E020.HSXOF4ST,  E020.HSXOF4SR,  E020.HSXOF4NS,  " & vbCrLf
    sql = sql & "       E020.HSXOF4SZ,  E020.HSXOF4ET,  E020.HSXOSF4PTK,E020.HSXBM1NS,  E020.HSXBM1SZ,  " & vbCrLf

    'C－OSF3判定機能  ---
    sql = sql & "       E036.COSF3FLAG " & vbCrLf
'Add End   2010/12/23 SMPK A.Nagamine
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    sql = sql & "       ,NVL(E036.HSXDKTMP,' ') HSXDKTMP " & vbCrLf
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    
'Add Start 2010/12/23 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : C-OSF3判定テーブルのキー変更(Class 14 -> 22), Cu-deco(C, CJ, CJ(LT), CJ2)の仕様判定追加
    
    sql = sql & ",      E020.HSXCOSF3HS,E020.HSXCOSF3SH,E020.HSXCOSF3ST,E020.HSXCOSF3SR,E020.HSXCOSF3NS," & vbCrLf
    sql = sql & "       E020.HSXCOSF3SZ,E036.HSXCOSF3ET,E020.HSXCOSF3PK,                                " & vbCrLf
    sql = sql & "       E020.HSXCPK,    E020.HSXCSZ,    E020.HSXCHT,    E020.HSXCHS,    E020.HSXCJPK,   " & vbCrLf
    sql = sql & "       E020.HSXCJNS,   E020.HSXCJHT,   E020.HSXCJHS,   E020.HSXCJLTPK, E020.HSXCJLTNS, " & vbCrLf
    sql = sql & "       E020.HSXCJLTHT, E020.HSXCJLTHS, E020.HSXCJ2PK,  E020.HSXCJ2NS,  E020.HSXCJ2HT,  " & vbCrLf
    sql = sql & "       E020.HSXCJ2HS,  E036.HSXCJLTBND " & vbCrLf
    
'Add End   2010/12/23 SMPK A.Nagamine
    
    sql = sql & "FROM   TBCME020 E020,TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E020.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E020.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E020.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E020.OPECOND   =   '" & tNew_Hinban.opecond & "' AND " & vbCrLf
    sql = sql & "       E036.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tNew_Hinban.opecond & "'  " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    With tbl_chk1_4(1)
    
        'C－OSF3判定機能   ---
        If IsNull(rs("COSF3FLAG")) = False Then .HSXOF4HS = rs("COSF3FLAG") Else .HSXOF4HS = " "            'C-OSF3ﾌﾗｸﾞ

'--------------- 2008/08/25 INSERT START  By Systech ---------------
        .HSXDKTMP = rs("HSXDKTMP")
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
        
        'OSF4
    'Add Start 2010/12/23 SMPK A.Nagamine       : CLESTA評価対応(Cu-deco) : C-OSF3判定テーブルのキー変更(Class 14 -> 22), Cu-deco(C, CJ, CJ(LT), CJ2)の仕様判定追加
        'If IsNull(rs("HSXCOSF3HS")) = False Then .HSXOF4HS = rs("HSXCOSF3HS") Else .HSXOF4HS = " "              'C-OSF3 保証方法_処 2010/12/24 Add
        
        'If IsNull(rs("HSXOF4SH")) = False Then .HSXOF4SH = rs("HSXOF4SH") Else .HSXOF4SH = " "              '測定位置_方
        'If IsNull(rs("HSXOF4ST")) = False Then .HSXOF4ST = rs("HSXOF4ST") Else .HSXOF4ST = " "              '測定位置_点
        'If IsNull(rs("HSXOF4SR")) = False Then .HSXOF4SR = rs("HSXOF4SR") Else .HSXOF4SR = " "              '測定位置_領
        'If IsNull(rs("HSXOF4NS")) = False Then .HSXOF4NS = rs("HSXOF4NS") Else .HSXOF4NS = " "              '熱処理法
        'If IsNull(rs("HSXOF4SZ")) = False Then .HSXOF4SZ = rs("HSXOF4SZ") Else .HSXOF4SZ = " "              '測定条件
        'If IsNull(rs("HSXOF4ET")) = False Then .HSXOF4ET = rs("HSXOF4ET") Else .HSXOF4ET = 0                '選択ET代
        'If IsNull(rs("HSXOSF4PTK")) = False Then .HSXOSF4PTK = rs("HSXOSF4PTK") Else .HSXOSF4PTK = "4"      'パターン区分
        
        If IsNull(rs("HSXCOSF3SH")) = False Then .HSXOF4SH = rs("HSXCOSF3SH") Else .HSXOF4SH = " "              'C-OSF3 測定位置_方     2010/12/24 Add
        If IsNull(rs("HSXCOSF3ST")) = False Then .HSXOF4ST = rs("HSXCOSF3ST") Else .HSXOF4ST = " "              'C-OSF3 測定位置_点     2010/12/24 Add
        If IsNull(rs("HSXCOSF3SR")) = False Then .HSXOF4SR = rs("HSXCOSF3SR") Else .HSXOF4SR = " "              'C-OSF3 測定位置_領     2010/12/24 Add
        If IsNull(rs("HSXCOSF3NS")) = False Then .HSXOF4NS = rs("HSXCOSF3NS") Else .HSXOF4NS = " "              'C-OSF3 熱処理法        2010/12/24 Add
        If IsNull(rs("HSXCOSF3SZ")) = False Then .HSXOF4SZ = rs("HSXCOSF3SZ") Else .HSXOF4SZ = " "              'C-OSF3 測定条件        2010/12/24 Add
        If IsNull(rs("HSXCOSF3ET")) = False Then .HSXOF4ET = rs("HSXCOSF3ET") Else .HSXOF4ET = 0                'C-OSF3 選択ET代        2010/12/24 Add
        If IsNull(rs("HSXCOSF3PK")) = False Then .HSXOSF4PTK = rs("HSXCOSF3PK") Else .HSXOSF4PTK = "4"          'C-OSF3 パターン区分    2010/12/24 Add
        
        If IsNull(rs("HSXCPK")) = False Then .HSXCPK = rs("HSXCPK") Else .HSXCPK = " "                  '/* 品ＳＸＣパターン区分 */
        If IsNull(rs("HSXCSZ")) = False Then .HSXCSZ = rs("HSXCSZ") Else .HSXCSZ = " "                  '/* 品ＳＸＣ測定条件     */
        If IsNull(rs("HSXCHT")) = False Then .HSXCHT = rs("HSXCHT") Else .HSXCHT = " "                  '/* 品ＳＸＣ保証方法＿対 */
        If IsNull(rs("HSXCHS")) = False Then .HSXCHS = rs("HSXCHS") Else .HSXCHS = " "                  '/* 品ＳＸＣ保証方法＿処 */
        
        If IsNull(rs("HSXCJPK")) = False Then .HSXCJPK = rs("HSXCJPK") Else .HSXCJPK = " "              '/* 品ＳＸＣＪパターン区分 */
        If IsNull(rs("HSXCJNS")) = False Then .HSXCJNS = rs("HSXCJNS") Else .HSXCJNS = "  "             '/* 品ＳＸＣＪ熱処理法     */
        If IsNull(rs("HSXCJHT")) = False Then .HSXCJHT = rs("HSXCJHT") Else .HSXCJHT = " "              '/* 品ＳＸＣＪ保証方法＿対 */
        If IsNull(rs("HSXCJHS")) = False Then .HSXCJHS = rs("HSXCJHS") Else .HSXCJHS = " "              '/* 品ＳＸＣＪ保証方法＿処 */
        
        If IsNull(rs("HSXCJLTPK")) = False Then .HSXCJLTPK = rs("HSXCJLTPK") Else .HSXCJLTPK = " "      '/* 品ＳＸＣＪＬＴパターン区分 */
        If IsNull(rs("HSXCJLTNS")) = False Then .HSXCJLTNS = rs("HSXCJLTNS") Else .HSXCJLTNS = "  "     '/* 品ＳＸＣＪＬＴ熱処理法     */
        If IsNull(rs("HSXCJLTHT")) = False Then .HSXCJLTHT = rs("HSXCJLTHT") Else .HSXCJLTHT = " "      '/* 品ＳＸＣＪＬＴ保証方法＿対 */
        If IsNull(rs("HSXCJLTHS")) = False Then .HSXCJLTHS = rs("HSXCJLTHS") Else .HSXCJLTHS = " "      '/* 品ＳＸＣＪＬＴ保証方法＿処 */
        
        If IsNull(rs("HSXCJ2PK")) = False Then .HSXCJ2PK = rs("HSXCJ2PK") Else .HSXCJ2PK = " "          '/* 品ＳＸＣＪ２パターン区分 */
        If IsNull(rs("HSXCJ2NS")) = False Then .HSXCJ2NS = rs("HSXCJ2NS") Else .HSXCJ2NS = "  "         '/* 品ＳＸＣＪ２熱処理法     */
        If IsNull(rs("HSXCJ2HT")) = False Then .HSXCJ2HT = rs("HSXCJ2HT") Else .HSXCJ2HT = " "          '/* 品ＳＸＣＪ２保証方法＿対 */
        If IsNull(rs("HSXCJ2HS")) = False Then .HSXCJ2HS = rs("HSXCJ2HS") Else .HSXCJ2HS = " "          '/* 品ＳＸＣＪ２保証方法＿処 */
        
        If IsNull(rs("HSXCJLTBND")) = False Then .HSXCJLTBND = rs("HSXCJLTBND") Else .HSXCJLTBND = 0    '/* 品SXL/CJLTバンド幅 Number(3,0) */
        
    'Add End 2010/12/23 SMPK A.Nagamine
    
    End With
    
    Set rs = Nothing
    
    '------------------------------------------ 指示取得 ------------------------------------------------------
    On Error GoTo Apl_down
    If iELCs_Flg = 0 Or iELCs_Flg = 2 Then
        
        'ＯＳＦ４
        'Add Start 2010/12/23 SMPK A.Nagamine
        'sErr_Msg = "2-2 OSF4ﾁｪｯｸ"
        sErr_Msg = "2-2 C-OSF3ﾁｪｯｸ"
        'Add End   2010/12/23 SMPK A.Nagamine
        sResult = ""
        
        'Add Start 2010/12/23 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : C-OSF3判定テーブルのキー変更(Class 14 -> 22), Cu-deco(C, CJ, CJ(LT), CJ2)の仕様判定追加
        'RET = funCodeDBGet("SB", "14", "O4", 0, " ", sResult)
        RET = funCodeDBGet("SB", "22", "O4", 0, " ", sResult)
        'Add End   2010/12/23 SMPK A.Nagamine
        If RET <> 0 Then
            sErr_Msg = sErr_Msg & "→指示取得"
            GoTo CodeDBGet_Error
        End If
        tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HSXOF4HS
        tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HSXOF4HS
        tbl_chk1_4_1(0).SOKU_HOU = tbl_chk1_4(0).HSXOF4SH
        tbl_chk1_4_1(1).SOKU_HOU = tbl_chk1_4(1).HSXOF4SH
        tbl_chk1_4_1(0).SOKU_TEN = tbl_chk1_4(0).HSXOF4ST
        tbl_chk1_4_1(1).SOKU_TEN = tbl_chk1_4(1).HSXOF4ST
        tbl_chk1_4_1(0).SOKU_RYOU = tbl_chk1_4(0).HSXOF4SR
        tbl_chk1_4_1(1).SOKU_RYOU = tbl_chk1_4(1).HSXOF4SR
        tbl_chk1_4_1(0).NETSU = tbl_chk1_4(0).HSXOF4NS
        tbl_chk1_4_1(1).NETSU = tbl_chk1_4(1).HSXOF4NS
        tbl_chk1_4_1(0).JOUKEN = tbl_chk1_4(0).HSXOF4SZ
        tbl_chk1_4_1(1).JOUKEN = tbl_chk1_4(1).HSXOF4SZ
        tbl_chk1_4_1(0).ET = tbl_chk1_4(0).HSXOF4ET
        tbl_chk1_4_1(1).ET = tbl_chk1_4(1).HSXOF4ET
        tbl_chk1_4_1(0).PATTERN = tbl_chk1_4(0).HSXOSF4PTK
        tbl_chk1_4_1(1).PATTERN = tbl_chk1_4(1).HSXOSF4PTK
        RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK2-2,C-OSF3")
        If RET <> 0 Then
            funChkFurikae2_2 = RET
            GoTo Apl_Exit
        End If
        
        
        'Add Start 2010/12/23 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : C-OSF3判定テーブルのキー変更(Class 14 -> 22), Cu-deco(C, CJ, CJ(LT), CJ2)の仕様判定追加
        
        RET = funCodeDBGet("SB", "22", "C", 0, " ", sResult)
        
        tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HSXCHS
        tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HSXCHS
        tbl_chk1_4_1(0).Min = 0
        tbl_chk1_4_1(1).Min = 0
        tbl_chk1_4_1(0).max = 0
        tbl_chk1_4_1(1).max = 0
        tbl_chk1_4_1(0).SOKU_HOU = " "
        tbl_chk1_4_1(1).SOKU_HOU = " "
        tbl_chk1_4_1(0).SOKU_TEN = " "
        tbl_chk1_4_1(1).SOKU_TEN = " "
        tbl_chk1_4_1(0).SOKU_ICHI = " "
        tbl_chk1_4_1(1).SOKU_ICHI = " "
        tbl_chk1_4_1(0).SOKU_RYOU = " "
        tbl_chk1_4_1(1).SOKU_RYOU = " "
        tbl_chk1_4_1(0).UMU = " "
        tbl_chk1_4_1(1).UMU = " "
        tbl_chk1_4_1(0).NETSU = "  "
        tbl_chk1_4_1(1).NETSU = "  "
        tbl_chk1_4_1(0).JOUKEN = tbl_chk1_4(0).HSXCSZ
        tbl_chk1_4_1(1).JOUKEN = tbl_chk1_4(1).HSXCSZ
        tbl_chk1_4_1(0).ET = 0
        tbl_chk1_4_1(1).ET = 0
        tbl_chk1_4_1(0).KENSA = "  "
        tbl_chk1_4_1(1).KENSA = "  "
        tbl_chk1_4_1(0).LINE = " "
        tbl_chk1_4_1(1).LINE = " "
        tbl_chk1_4_1(0).PATTERN = tbl_chk1_4(0).HSXCPK
        tbl_chk1_4_1(1).PATTERN = tbl_chk1_4(1).HSXCPK
        tbl_chk1_4_1(0).HSXDKTMP = " "
        tbl_chk1_4_1(1).HSXDKTMP = " "
        tbl_chk1_4_1(0).HSXCNKHI = " "
        tbl_chk1_4_1(1).HSXCNKHI = " "
        RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK2-2,Cu-deco,C")
        If RET <> 0 Then
            funChkFurikae2_2 = RET
            GoTo Apl_Exit
        End If
        
        RET = funCodeDBGet("SB", "22", "CJ", 0, " ", sResult)
        
        tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HSXCJHS
        tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HSXCJHS
        tbl_chk1_4_1(0).Min = 0
        tbl_chk1_4_1(1).Min = 0
        tbl_chk1_4_1(0).max = 0
        tbl_chk1_4_1(1).max = 0
        tbl_chk1_4_1(0).SOKU_HOU = " "
        tbl_chk1_4_1(1).SOKU_HOU = " "
        tbl_chk1_4_1(0).SOKU_TEN = " "
        tbl_chk1_4_1(1).SOKU_TEN = " "
        tbl_chk1_4_1(0).SOKU_ICHI = " "
        tbl_chk1_4_1(1).SOKU_ICHI = " "
        tbl_chk1_4_1(0).SOKU_RYOU = " "
        tbl_chk1_4_1(1).SOKU_RYOU = " "
        tbl_chk1_4_1(0).UMU = " "
        tbl_chk1_4_1(1).UMU = " "
        tbl_chk1_4_1(0).NETSU = tbl_chk1_4(0).HSXCJNS
        tbl_chk1_4_1(1).NETSU = tbl_chk1_4(1).HSXCJNS
        tbl_chk1_4_1(0).JOUKEN = " "
        tbl_chk1_4_1(1).JOUKEN = " "
        tbl_chk1_4_1(0).ET = 0
        tbl_chk1_4_1(1).ET = 0
        tbl_chk1_4_1(0).KENSA = "  "
        tbl_chk1_4_1(1).KENSA = "  "
        tbl_chk1_4_1(0).LINE = " "
        tbl_chk1_4_1(1).LINE = " "
        tbl_chk1_4_1(0).PATTERN = tbl_chk1_4(0).HSXCJPK
        tbl_chk1_4_1(1).PATTERN = tbl_chk1_4(1).HSXCJPK
        tbl_chk1_4_1(0).HSXDKTMP = " "
        tbl_chk1_4_1(1).HSXDKTMP = " "
        tbl_chk1_4_1(0).HSXCNKHI = " "
        tbl_chk1_4_1(1).HSXCNKHI = " "
        RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK2-2,Cu-deco,CJ")
        If RET <> 0 Then
            funChkFurikae2_2 = RET
            GoTo Apl_Exit
        End If
        
        RET = funCodeDBGet("SB", "22", "CJLT", 0, " ", sResult)
        
        tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HSXCJLTHS
        tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HSXCJLTHS
        tbl_chk1_4_1(0).Min = 0
        tbl_chk1_4_1(1).Min = 0
        tbl_chk1_4_1(0).max = 0
        tbl_chk1_4_1(1).max = 0
        tbl_chk1_4_1(0).SOKU_HOU = " "
        tbl_chk1_4_1(1).SOKU_HOU = " "
        tbl_chk1_4_1(0).SOKU_TEN = " "
        tbl_chk1_4_1(1).SOKU_TEN = " "
        tbl_chk1_4_1(0).SOKU_ICHI = " "
        tbl_chk1_4_1(1).SOKU_ICHI = " "
        tbl_chk1_4_1(0).SOKU_RYOU = " "
        tbl_chk1_4_1(1).SOKU_RYOU = " "
        tbl_chk1_4_1(0).UMU = " "
        tbl_chk1_4_1(1).UMU = " "
        tbl_chk1_4_1(0).NETSU = tbl_chk1_4(0).HSXCJLTNS
        tbl_chk1_4_1(1).NETSU = tbl_chk1_4(1).HSXCJLTNS
        tbl_chk1_4_1(0).JOUKEN = " "
        tbl_chk1_4_1(1).JOUKEN = " "
        tbl_chk1_4_1(0).ET = 0
        tbl_chk1_4_1(1).ET = 0
        tbl_chk1_4_1(0).KENSA = "  "
        tbl_chk1_4_1(1).KENSA = "  "
        tbl_chk1_4_1(0).LINE = " "
        tbl_chk1_4_1(1).LINE = " "
        tbl_chk1_4_1(0).PATTERN = tbl_chk1_4(0).HSXCJLTPK
        tbl_chk1_4_1(1).PATTERN = tbl_chk1_4(1).HSXCJLTPK
        tbl_chk1_4_1(0).HSXDKTMP = " "
        tbl_chk1_4_1(1).HSXDKTMP = " "
        tbl_chk1_4_1(0).HSXCNKHI = " "
        tbl_chk1_4_1(1).HSXCNKHI = " "
        RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK2-2,Cu-deco,CJLT")
        If RET <> 0 Then
            funChkFurikae2_2 = RET
            GoTo Apl_Exit
        End If
        
        RET = funCodeDBGet("SB", "22", "CJ2", 0, " ", sResult)
        
        tbl_chk1_4_1(0).HOSYOU = tbl_chk1_4(0).HSXCJ2HS
        tbl_chk1_4_1(1).HOSYOU = tbl_chk1_4(1).HSXCJ2HS
        tbl_chk1_4_1(0).Min = 0
        tbl_chk1_4_1(1).Min = 0
        tbl_chk1_4_1(0).max = 0
        tbl_chk1_4_1(1).max = 0
        tbl_chk1_4_1(0).SOKU_HOU = " "
        tbl_chk1_4_1(1).SOKU_HOU = " "
        tbl_chk1_4_1(0).SOKU_TEN = " "
        tbl_chk1_4_1(1).SOKU_TEN = " "
        tbl_chk1_4_1(0).SOKU_ICHI = " "
        tbl_chk1_4_1(1).SOKU_ICHI = " "
        tbl_chk1_4_1(0).SOKU_RYOU = " "
        tbl_chk1_4_1(1).SOKU_RYOU = " "
        tbl_chk1_4_1(0).UMU = " "
        tbl_chk1_4_1(1).UMU = " "
        tbl_chk1_4_1(0).NETSU = tbl_chk1_4(0).HSXCJ2NS
        tbl_chk1_4_1(1).NETSU = tbl_chk1_4(1).HSXCJ2NS
        tbl_chk1_4_1(0).JOUKEN = " "
        tbl_chk1_4_1(1).JOUKEN = " "
        tbl_chk1_4_1(0).ET = 0
        tbl_chk1_4_1(1).ET = 0
        tbl_chk1_4_1(0).KENSA = "  "
        tbl_chk1_4_1(1).KENSA = "  "
        tbl_chk1_4_1(0).LINE = " "
        tbl_chk1_4_1(1).LINE = " "
        tbl_chk1_4_1(0).PATTERN = tbl_chk1_4(0).HSXCJ2PK
        tbl_chk1_4_1(1).PATTERN = tbl_chk1_4(1).HSXCJ2PK
        tbl_chk1_4_1(0).HSXDKTMP = " "
        tbl_chk1_4_1(1).HSXDKTMP = " "
        tbl_chk1_4_1(0).HSXCNKHI = " "
        tbl_chk1_4_1(1).HSXCNKHI = " "
        RET = funChkFurikae1_4_1(sResult, tbl_chk1_4_1(), iErr_Code, sErr_Msg, "CHECK2-2,Cu-deco,CJ2")
        If RET <> 0 Then
            funChkFurikae2_2 = RET
            GoTo Apl_Exit
        End If
        
        'Add End   2010/12/23 SMPK A.Nagamine
        
    End If
    
    
    '------------------------------------------ 終了処理  ------------------------------------------------------


Apl_Exit:
    
    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae2_2 = 0 Then
        funChkFurikae2_2 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae2_2 = -4
    GoTo Apl_Exit
    
CodeDBGet_Error:
    Set rs = Nothing
    If funChkFurikae2_2 = 0 Then
        funChkFurikae2_2 = -5
    End If
    GoTo Apl_Exit

End Function
'------------------------------------------------
'   窒素濃度チェック  2009/07/30 追加  Kameda
'------------------------------------------------
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sKeyID          ,I  ,String       :ﾌﾞﾛｯｸID、又は、SXL-ID
'          :tOld_Hinban     ,I  ,tFullHinban  :振替元品番(構造体)
'          :tNew_Hinban     ,I  ,tFullHinban  :振替先品番(構造体)
'          :iHinCnt         ,I  ,Integer      :複数品番カウント
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :
'履歴      :2009/09/01 仕様あり（上限、下限ともに０ではない場合）で実測無しは判定NG
'　　      :2009/09/01 仕様なしは判定OK
'　　      :2009/09/03 実測値0は判定OK
'　　      :2009/09/28 仕様上下ともに０は判定なし→結晶ドープ種類で判定
Public Function funChkFurikae2_4(sProccd As String, sBlockId As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iHinCnt As Integer, iErr_Code As Integer, sErr_Msg As String) As Integer
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim s
    Dim sClData As typ_chk2_4
    Dim dblNMin As Double
    Dim dblNMax As Double
    Dim wBLKID()    As String               '総合判定対象ﾌﾞﾛｯｸID
    Dim cnt As Integer
    Dim ErrFlg(1) As Boolean
    Dim strCdop As String
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae2_4 = 0
    
    'tbl_chk2_4(0) = sClData
    'tbl_chk2_4(1) = sClData
    dblNMin = -1
    dblNMax = -1
    ReDim Preserve tbl_chk2_4(0).NJDG(iHinCnt)
    ReDim Preserve tbl_chk2_4(1).NJDG(iHinCnt)
'    '------------------------------------------ CWの場合、ﾌﾞﾛｯｸIDを取得 ------------------------------------------------
    If (left(sProccd, 4) = "CW76") Then
        sErr_Msg = "2-4 BLK-ID取得"
        sql = vbNullString
        sql = sql & "SELECT CRYNUMCA FROM XSDCA " & vbCrLf
        sql = sql & "WHERE  SXLIDCA = '" & sBlockId & "' AND " & vbCrLf
        sql = sql & "       LIVKCA  = '0' " & vbCrLf
    
        On Error GoTo db_Error
        'SQL文の実行
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        
        '該当データなし
        If rs.EOF Or rs.RecordCount = 0 Then GoTo db_Error
        
        ReDim wBLKID(rs.RecordCount)
        For cnt = 1 To rs.RecordCount
            If IsNull(rs("CRYNUMCA")) = False Then wBLKID(cnt) = rs("CRYNUMCA") Else wBLKID(cnt) = " "
            rs.MoveNext
        Next cnt
        Set rs = Nothing
    Else
        ReDim wBLKID(1)
        wBLKID(1) = sBlockId
    End If
    '------------------------------------------ 判定データ取得(仕様TBCME020) ----------------------------------------
        sql = "SELECT nvl(HSXCDOPMN,0) as HSXCDOPMN " & vbCrLf
        sql = sql & ",nvl(HSXCDOPMX,0) as HSXCDOPMX " & vbCrLf
        sql = sql & ",nvl(HSXCDOP,' ') as HSXCDOP " & vbCrLf
        sql = sql & "FROM   TBCME020 " & vbCrLf
        sql = sql & "WHERE  HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
        sql = sql & "       MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
        sql = sql & "       FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
        sql = sql & "       OPECOND   =   '" & tNew_Hinban.opecond & "'  " & vbCrLf
        
        On Error GoTo db_Error
        'SQL文の実行
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        
        '該当データなし
        If rs.EOF Or rs.RecordCount > 1 Then
            GoTo db_Error
        End If
        
        '取得データセット
         
        dblNMin = rs("HSXCDOPMN")
        dblNMax = rs("HSXCDOPMX")
        strCdop = Trim(rs("HSXCDOP"))
        'If dblNMin = 0 And dblNMax = 0 Then   2009/09/28 Kameda
        '    '判定ＯＫ
        '    tbl_chk2_4(0).N2NOUDO = -1
        '    tbl_chk2_4(1).N2NOUDO = -1
        If strCdop = "" Then
            '判定ＯＫ
            tbl_chk2_4(0).N2NOUDO = -1
            tbl_chk2_4(1).N2NOUDO = -1
            Set rs = Nothing
            Exit Function
        End If
        Set rs = Nothing
    
    '------------------------------------------ 判定データ取得(TBCMJ020) ----------------------------------------
    For cnt = 1 To UBound(wBLKID)
        '窒素濃度取得
        ' SQL作成
        sql = "SELECT nvl(N2NOUDO,-1) TOPNOUDO "
        sql = sql & " FROM TBCMJ020  "
        sql = sql & " WHERE  BLOCKID = '" & wBLKID(cnt) & "'  "
        sql = sql & " AND SMPKBN = 'T' "
        sql = sql & " order by TRANCNT desc "
        ' 実行
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        
        '該当データなしは判定NG
        If rs.EOF Or rs.RecordCount = 0 Then
            tbl_chk2_4(0).N2NOUDO = -1
        Else
            tbl_chk2_4(0).N2NOUDO = rs("TOPNOUDO")
        End If
        Set rs = Nothing
        
        'Tail
        sql = "SELECT nvl(N2NOUDO,-1) BOTNOUDO "
        sql = sql & " FROM TBCMJ020  "
        sql = sql & " WHERE  BLOCKID = '" & wBLKID(cnt) & "'  "
        sql = sql & " AND SMPKBN = 'B' "
        sql = sql & " order by TRANCNT desc "
        ' 実行
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        
        '該当データなしは判定NG
        If rs.EOF Or rs.RecordCount = 0 Then
            tbl_chk2_4(1).N2NOUDO = -1
        Else
            tbl_chk2_4(1).N2NOUDO = rs("BOTNOUDO")
        End If
        Set rs = Nothing
        
    '------------------------------------------ 判定 -----------------------------------------------------------
        tbl_chk2_4(0).NJDG(iHinCnt) = "-1"
        tbl_chk2_4(1).NJDG(iHinCnt) = "-1"
        ErrFlg(0) = False
        ErrFlg(1) = False
        
        If tbl_chk2_4(0).N2NOUDO <> 0 And tbl_chk2_4(1).N2NOUDO <> 0 Then '実測値がT,Bとも0はOK
            If strCdop <> "Z" Then
                If tbl_chk2_4(0).N2NOUDO = -1 Then
                    funChkFurikae2_4 = 1
                    iErr_Code = 2301
                    sErr_Msg = "CHECK2-4,窒素濃度エラー、振替できません。"
                    ErrFlg(0) = True
                '濃度From, To範囲内か
                ElseIf tbl_chk2_4(0).N2NOUDO < dblNMin Then
                    funChkFurikae2_4 = 1
                    iErr_Code = 2301
                    sErr_Msg = "CHECK2-4,窒素濃度エラー、振替できません。"
                    ErrFlg(0) = True
                ElseIf tbl_chk2_4(0).N2NOUDO > dblNMax Then
                    funChkFurikae2_4 = 1
                    iErr_Code = 2301
                    sErr_Msg = "CHECK2-4,窒素濃度エラー、振替できません。"
                    ErrFlg(0) = True
                End If
                
                If tbl_chk2_4(1).N2NOUDO = -1 Then
                    funChkFurikae2_4 = 1
                    iErr_Code = 2301
                    sErr_Msg = "CHECK2-4,窒素濃度エラー、振替できません。"
                    ErrFlg(1) = True
                '濃度From, To範囲内か
                
                ElseIf tbl_chk2_4(1).N2NOUDO < dblNMin Then
                    funChkFurikae2_4 = 1
                    iErr_Code = 2301
                    sErr_Msg = "CHECK2-4,窒素濃度エラー、振替できません。"
                    ErrFlg(1) = True
                ElseIf tbl_chk2_4(1).N2NOUDO > dblNMax Then
                    funChkFurikae2_4 = 1
                    iErr_Code = 2301
                    sErr_Msg = "CHECK2-4,窒素濃度エラー、振替できません。"
                    ErrFlg(1) = True
                End If
            End If
        End If
    
    Next
    '判定ＯＫ
    If ErrFlg(0) = False Then tbl_chk2_4(0).NJDG(iHinCnt) = "0"
    If ErrFlg(1) = False Then tbl_chk2_4(1).NJDG(iHinCnt) = "0"
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae2_4 = 0 Then
        funChkFurikae2_4 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae2_4 = -4
    GoTo Apl_Exit
    
End Function
'------------------------------------------------
'   マルチ引上げ適用チェック  2011/05/19 追加  Kameda
'------------------------------------------------
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sBlockId        ,I  ,String       :ﾌﾞﾛｯｸID、又は、SXL-ID
'          :tOld_Hinban     ,I  ,tFullHinban  :振替元品番(構造体)
'          :tNew_Hinban     ,I  ,tFullHinban  :振替先品番(構造体)
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :成否(0:正常終了(振替OK),1:正常終了(振替NG),-2:取得ｴﾗｰ)
'説明      :対象工程のブロック、またはシングルの品番がマルチ引上げ適用不可の場合
'           該当ブロック､シングルがマルチ2本目以降の場合はエラーとします｡
'           または、該当ブロック、シングルがマルチ２本目以降の場合
'           含まれるの品番がマルチ適用不可があった場合､エラーとし､流動不可とします｡
'             ※リチャージ、残引き（計画）の2本目以降とは,連続コード3桁目が2以上
'履歴      :2011/05/19 Kameda
Public Function funChkFurikae2_5(sProccd As String, sBlockId As String, tOld_Hinban As tFullHinban, tNew_Hinban As tFullHinban, _
                                 iHinCnt As Integer, iErr_Code As Integer, sErr_Msg As String) As Integer
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim s
    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim i           As Integer
    Dim sXtal       As String
    On Error GoTo Apl_down
    
    '戻り値初期化
    funChkFurikae2_5 = 0
    
    '-------------------------------- 振替元マルチ引上げ適用可否仕様データ取得 ------------------------------------------------------
    sErr_Msg = "2-5 品番仕様取得(" & tNew_Hinban.hinban & Format(tNew_Hinban.mnorevno, "00") & tNew_Hinban.factory & tNew_Hinban.opecond & ")"
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT NVL(E036.MLTHTFLG,' ') MLTHTFLG " & vbCrLf
    sql = sql & "FROM   TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E036.HINBAN    =   '" & tNew_Hinban.hinban & "'  AND " & vbCrLf
    sql = sql & "       E036.MNOREVNO  =    " & tNew_Hinban.mnorevno & " AND " & vbCrLf
    sql = sql & "       E036.FACTORY   =   '" & tNew_Hinban.factory & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND   =   '" & tNew_Hinban.opecond & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
        
    '
    If Trim(rs("MLTHTFLG")) <> "" Then
        tbl_chk2_5.MLTHTFLG = rs("MLTHTFLG")
    Else
        tbl_chk2_5.MLTHTFLG = "0"
    End If
    
    Set rs = Nothing

'    '------------------------------------------ CWの場合、結晶番号を取得 ------------------------------------------------
    If (left(sProccd, 4) >= "CW75") Then
        sErr_Msg = "2_5 BLK-ID取得"
        sql = vbNullString
        sql = sql & "SELECT XTALCA FROM XSDCA " & vbCrLf
        sql = sql & "WHERE  SXLIDCA = '" & sBlockId & "' AND " & vbCrLf
        sql = sql & "       LIVKCA  = '0' " & vbCrLf
    Else
        sql = vbNullString
        sql = sql & "SELECT XTALCA FROM XSDCA " & vbCrLf
        sql = sql & "WHERE  CRYNUMCA = '" & sBlockId & "' AND " & vbCrLf
        sql = sql & "       LIVKCA  = '0' " & vbCrLf
    End If
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount = 0 Then GoTo db_Error
    
    If IsNull(rs("XTALCA")) = False Then sXtal = rs("XTALCA") Else sXtal = " "
    
    Set rs = Nothing
    
    
    '------------------------------------------ 判定データ取得(XSDC1) ------------------------------------------
    '連続コード取得
    sql = "SELECT NVL(SIJICNT,0) SIJICNT,NVL(RENBAN,0) RENBAN " & vbCrLf
    sql = sql & "FROM XSDC1,TBCMH001 " & vbCrLf
    sql = sql & "WHERE XTALC1 = '" & sXtal & "' " & vbCrLf
    sql = sql & " AND  HISIJIC1 = UPINDNO "
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo db_Error
    End If
    
    '取得データセット
    With tbl_chk2_5
        .SIJICNT = rs("SIJICNT")
        .RENBAN = rs("RENBAN")
        
        Set rs = Nothing
        
    End With
    
    '[マルチ引上げ可否フラグ：０＝可　１＝不可]
    'マルチ引上げ適用不可の場合
    '該当ブロック､シングルがマルチ2本目以降の場合はエラー
    On Error GoTo Apl_down
    
    tbl_chk2_5.MLTJDG(iHinCnt) = "-1"
        
    sErr_Msg = "2-5 マルチ引上げ適用ﾁｪｯｸ"
    If tbl_chk2_5.MLTHTFLG = "1" Then
        If tbl_chk2_5.RENBAN > 1 Then
            funChkFurikae2_5 = 1
            iErr_Code = 2501
            sErr_Msg = "CHECK2-5,マルチ引上げ適用可否エラー" '
            gsTbcmy028ErrCode = "02501"
            GoTo Apl_Exit
        End If
    End If
    
    tbl_chk2_5.MLTJDG(iHinCnt) = "0"
    
    '該当ブロック､シングルがマルチ2本目以降の場合
    '含まれる品番がマルチ適用不可があった場合､エラー
    'If tbl_chk2_5.RENBAN > 1 Then
    '    If tbl_chk2_5.MLTHTFLG = "1" Then
    '        funChkFurikae2_5 = 1
    '        iErr_Code = 2501
    '        sErr_Msg = "CHECK2-5,マルチ引上げ適用可否エラー" '
    '        gsTbcmy028ErrCode = "02501"
    '        GoTo Apl_Exit
    '    End If
    'End If
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
db_Error:
    Set rs = Nothing
    If funChkFurikae2_5 = 0 Then
        funChkFurikae2_5 = -3
    End If
    GoTo Apl_Exit

Apl_down:
    funChkFurikae2_5 = -4
    GoTo Apl_Exit

End Function



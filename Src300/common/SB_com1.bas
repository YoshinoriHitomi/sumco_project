Attribute VB_Name = "SB_Com1"
Option Explicit

'振替候補品番
Public Type tbl_KouhoHin
    GETHINBAN As String * 12        ' 振替候補品番
End Type
Public KouhoHinban() As tbl_KouhoHin           ' 振替候補品番データ
'redimをやる

'仕様取得構造体
Public Type typ_Spec1_1
'    HWFTYPE     As String * 1       'タイプ    'HWFTYPE→HSXTYPE 振替チェックと合わせた 2011/05/11 SETsw kubota
    HSXTYPE     As String * 1       'タイプ
    BLOCKHFLAG  As String * 1       'ブロック単位保証フラグ
End Type
Public tbl_spec1_1(1) As typ_Spec1_1

Public Type typ_Spec1_2
    HSXCDIR     As String * 1       '結晶面方位
    HSXCSCEN    As Double           '結晶面傾き中心
    HSXDOP      As String * 1       'ドーパント
    HWFCDOP     As String * 1       '結晶ドープ
    HSXSDSLP    As String * 1       'シード傾き
    HSXDPDIR    As String * 2       '溝位置方位
    MCNO1       As String * 1       '品種
    MCNO2       As String * 1       '引上げ速度
    MCNO3       As String * 1       'HZタイプ
    DCHYUUBU    As String * 1       'ドローチューブ
End Type
Public tbl_spec1_2(1) As typ_Spec1_2

Public Type typ_Spec1_3
    HWFWARPR    As String * 1       'Warpランク
End Type
Public tbl_spec1_3(1) As typ_Spec1_3

Public Type typ_Spec1_4
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
    HSXOF4HS    As String * 1       '保証方法_対象
    HSXOF4SH    As String * 1       '測定位置_方
    HSXOF4ST    As String * 1       '測定位置_点
    HSXOF4SR    As String * 1       '測定位置_領
    HSXOF4NS    As String * 2       '熱処理法
    HSXOF4SZ    As String * 1       '測定条件
    HSXOF4ET    As Integer          '選択ET代
    HSXOSF4PTK  As String * 1       'パターン区分
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
    HWFSIRDMX   As Integer          '軸状転位上限(SIRD)
    HWFSIRDSZ   As String * 1       '軸状転位測定条件(SIRD)
    HWFSIRDHT   As String * 1       '軸状転位保証方法＿対(SIRD)
    HWFSIRDHS   As String * 1       '軸状転位保証方法＿処(SIRD)
    HWFSIRDKM   As String * 1       '軸状転位検査頻度＿枚(SIRD)
    HWFSIRDKH   As String * 1       '軸状転位検査頻度＿保(SIRD)
    HWFSIRDKU   As String * 1       '軸状転位検査頻度＿ウ(SIRD)
    HWFSIRDPS   As String * 2       '軸状転位TB保証位置(SIRD)
    HWFSIRDKN   As String * 1       '軸状転位検査頻度＿抜(SIRD)
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
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
    HSXDVDMN    As Integer          '下限
    HSXDVDMX    As Integer          '上限
    HSXLDLHS    As String * 1       '保証方法_対象
    HSXLDLMN    As Integer          '下限
    HSXLDLMX    As Integer          '上限
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
    HSXGDLINE   As String           'ﾗｲﾝ数
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
'C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
    COSF3FLAG   As String
'C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---

'--------------- 2008/08/25 INSERT START  By Systech ---------------
    HSXDKTMP    As String * 1       ' DK温度
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

End Type
Public tbl_spec1_4(1) As typ_Spec1_4

Public Type typ_Spec1_4_1
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
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 STRING型に変更
'    LINE        As Integer          'ライン数           ????????????????(桁数）
    LINE        As String           'ライン数           ????????????????(桁数）
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 STRING型に変更
    PATTERN     As String * 1       'パターン区分
    HOSYOU1     As String           '保証方法＿対象
    Min1        As String           '下限
    Max1        As String           '上限
    SOKU_HOU1   As String           '測定位置＿方
    SOKU_TEN1   As String           '測定位置＿点
    SOKU_ICHI1  As String           '測定位置＿位
    SOKU_RYOU1  As String           '測定位置＿領
    UMU1        As String           '検査有無           ????????????????(桁数）
    NETSU1      As String           '熱処理法
    JOUKEN1     As String           '測定条件
    ET1         As String           '選択ＥＴ代
    KENSA1      As String           '検査方法
    Line1       As String           'ライン数           ????????????????(桁数）
    PATTERN1    As String           'パターン区分
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    HSXDKTMP    As String * 1       ' DK温度
    HSXDKTMP1   As String           ' DK温度カラム名
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    HWFSIRDMX   As Integer          '軸状転位上限(SIRD)
    HWFSIRDHT   As String * 1       '軸状転位保証方法＿対(SIRD)
    HWFSIRDKM   As String * 1       '軸状転位検査頻度＿枚(SIRD)
    HWFSIRDKH   As String * 1       '軸状転位検査頻度＿保(SIRD)
    HWFSIRDKU   As String * 1       '軸状転位検査頻度＿ウ(SIRD)
    HWFSIRDKN   As String * 1       '軸状転位検査頻度＿抜(SIRD)

    HWFSIRDMX1  As String           '軸状転位上限(SIRD)
    HWFSIRDHT1  As String           '軸状転位保証方法＿対(SIRD)
    HWFSIRDKM1  As String           '軸状転位検査頻度＿枚(SIRD)
    HWFSIRDKH1  As String           '軸状転位検査頻度＿保(SIRD)
    HWFSIRDKU1  As String           '軸状転位検査頻度＿ウ(SIRD)
    HWFSIRDKN1  As String           '軸状転位検査頻度＿抜(SIRD)
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
End Type
Public tbl_spec1_4_1(0) As typ_Spec1_4_1

Public Type typ_Spec1_5
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
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
    HWFGDLINE   As String           'ﾗｲﾝ数
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    HWFGDSZY    As String * 1       'GD測定条件
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    HWFANTNP    As String           'AN温度
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
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
Public tbl_spec1_5(1) As typ_Spec1_5

Public Type typ_Spec1_5_1
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
    HOSYOU1     As String           '保証方法＿対象
    Min1        As String           '下限
    Max1        As String           '上限
    SOKU_HOU1   As String           '測定位置＿方
    SOKU_TEN1   As String           '測定位置＿点
    SOKU_ICHI1  As String           '測定位置＿位
    SOKU_RYOU1  As String           '測定位置＿領
    UMU1        As String           '検査有無           ????????????????(桁数）
    NETSU1      As String           '熱処理法
    JOUKEN1     As String           '測定条件
    ET1         As String           '選択ＥＴ代
    KENSA1      As String           '検査方法
    PATTERN1    As String           'パターン区分
    KENH_NUKI1  As String           '検査頻度_抜　04/04/13 ooba
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ライン数追加
    LINE        As String           'ライン数
    Line1       As String           'ライン数
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ライン数追加
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    Antnp       As String           'AN温度
    ANTNP1       As String          'AN温度
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    PUAGEN      As Double           'PUA限          '追加　06/05/31 ooba START =======>
    PUAPER      As Double           'PUA率
    PUAGEN1     As String           'PUA限
    PUAPER1     As String           'PUA率          '追加　06/05/31 ooba END =========>
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    HWFGDSZY    As String * 1       'GD測定条件
    HWFGDSZY1   As String           'GD測定条件
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---

'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    HWFSIRDMX   As Integer          '軸状転位上限(SIRD)
    HWFSIRDHT   As String * 1       '軸状転位保証方法＿対(SIRD)
    HWFSIRDKM   As String * 1       '軸状転位検査頻度＿枚(SIRD)
    HWFSIRDKH   As String * 1       '軸状転位検査頻度＿保(SIRD)
    HWFSIRDKU   As String * 1       '軸状転位検査頻度＿ウ(SIRD)

    HWFSIRDMX1  As String           '軸状転位上限(SIRD)
    HWFSIRDHT1  As String           '軸状転位保証方法＿対(SIRD)
    HWFSIRDKM1  As String           '軸状転位検査頻度＿枚(SIRD)
    HWFSIRDKH1  As String           '軸状転位検査頻度＿保(SIRD)
    HWFSIRDKU1  As String           '軸状転位検査頻度＿ウ(SIRD)
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)

End Type
Public tbl_spec1_5_1(0) As typ_Spec1_5_1

Public Type typ_Spec1_6
    HWFNP1AR    As Double           '品WFナノトポ１エリア
    HWFNP1MAX   As Double           '品WFナノトポ１上限
    HWFNP2AR    As Double           '品WFナノトポ２エリア
    HWFNP2MAX   As Double           '品WFナノトポ２上限
    HSXCSCEN    As Double           '結晶面傾き中心
End Type
Public tbl_spec1_6(1) As typ_Spec1_6

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
Public Type typ_Spec1_9
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
Public tbl_spec1_9(1) As typ_Spec1_9
Public Type typ_Spec1_9_1
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
    Antnp       As String           'AN温度
    EPATU       As Double           'エピ厚
    HOSYOU1     As String           '保証方法＿対象(カラム名)
    MIN_LIMIT1  As String           '下限(カラム名)
    MAX_LIMIT1  As String           '上限(カラム名)
    SOKU_HOU1   As String           '測定位置＿方(カラム名)
    SOKU_TEN1   As String           '測定位置＿点(カラム名)
    SOKU_ICHI1  As String           '測定位置＿位(カラム名)
    SOKU_RYOU1  As String           '測定位置＿領(カラム名)
    UMU1        As String           '検査有無(カラム名)
    NETSU1      As String           '熱処理法(カラム名)
    JOUKEN1     As String           '測定条件(カラム名)
    ET1         As String           '選択ＥＴ代(カラム名)
    KENSA1      As String           '検査方法(カラム名)
    PATTERN1    As String           'パターン区分(カラム名)
    KENH_NUKI1  As String           '検査頻度＿抜(カラム名)
    ANTNP1      As String           'AN温度(カラム名)
    EPATU1      As String           'エピ厚(カラム名)
End Type
Public tbl_spec1_9_1(0) As typ_Spec1_9_1
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

'常識仕様ﾁｪｯｸ2　06/10/05 ooba
Public Type typ_Spec1_10
    HSXCDIR     As String * 1       '結晶面方位
    HSXCSCEN    As Double           '結晶面傾き中心     ''2008/11/27 結晶面傾中心チェック緩和(2) ADD By Systech
    HSXDOP      As String * 1       'ドーパント
    HWFCDOP     As String * 1       '結晶ドープ
    HSXDPDIR    As String * 2       '溝位置方位
    MCNO1       As String * 1       '品種
    MCNO2       As String * 1       '引上げ速度
    MCNO3       As String * 1       'HZタイプ
    DCHYUUBU    As String * 1       'ドローチューブ
End Type
Public tbl_spec1_10(1) As typ_Spec1_10

''C－OSF3チェックの変更 2008.04.20 青柳
Public sFlg_2_2 As String

    
'------------------------------------------------
' 振替候補品番取得（仕様チェック）
'------------------------------------------------

'概要      :パラメータに指定された、振替元品番から振り替えが可能な品番を検索し、結果を返す。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sBlockid        ,I  ,String       :ﾌﾞﾛｯｸID、又は、結晶番号
'          :sOld_Hinban     ,I  ,String       :振替元品番
'          :sGet_Hinban()   ,O  ,String       :振替候補品番
'          :iErr_Code       ,O  ,Integer      :ｴﾗｰｺｰﾄﾞ(戻り値と同一)
'          :sErr_Msg        ,O  ,String       :ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ
'          :戻り値          ,O  ,Integer      :取得の成否(0:正常取得, -1:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/10 新規作成　SB

Public Function fncGetKouhoHinbanShiyou(sProccd As String, sBlockId As String, sOld_Hinban As String, sGet_Hinban() As tbl_KouhoHin, iErr_Code As Integer, sErr_Msg As String) As Integer
    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sResult2    As String       'コードＤＢ取得関数の取得変数(FE) 2011/04/07追加 SETsw kubota
    Dim sMakesql    As String       '呼び出しファンクションSQL作成
    Dim sMakesql1   As String       '呼び出しファンクションSQL作成
    Dim sMakesql2   As String       '呼び出しファンクションSQL作成
    Dim sMakesql3   As String       '呼び出しファンクションSQL作成
    Dim sMakesql4   As String       '呼び出しファンクションSQL作成
    Dim sMakesql5   As String       '呼び出しファンクションSQL作成
    Dim sMakesql6   As String       '呼び出しファンクションSQL作成
    Dim w_i         As Long         'カウンタ
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    Dim sMakesql9   As String       '呼び出しファンクションSQL作成
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    Dim sMakesql10  As String       '呼び出しファンクションSQL作成　06/10/05 ooba
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    fncGetKouhoHinbanShiyou = 0
    
    '------------------------------------------ 入力チェック -------------------------------------------------
    'ﾌﾞﾛｯｸID、又は、結晶番号の桁数チェック
    If Trim$(sBlockId) = "" Then
            fncGetKouhoHinbanShiyou = -1
            GoTo Apl_Error
    End If
    If Len(sBlockId) <> 12 Then
            fncGetKouhoHinbanShiyou = -1
            GoTo Apl_Error
    End If
    '振替元品番の桁数チェック
    If Trim$(sOld_Hinban) = "" Then
            fncGetKouhoHinbanShiyou = -1
            GoTo Apl_Error
    End If
    If Len(sOld_Hinban) <> 12 Then
            fncGetKouhoHinbanShiyou = -1
            GoTo Apl_Error
    End If
    
    '------------------------------------------ 指示取得 ------------------------------------------------------
    '振替指示データ取得
    sResult = ""
'    RET = funCodeDBGet("SB", "FC", sProccd, 0, " ", sResult)
    RET = funCodeDBGet("SB", "FD", sProccd, 0, " ", sResult)        'FC→FD 2011/04/07修正 SETsw kubota
    If RET <> 0 Then
        fncGetKouhoHinbanShiyou = -2
        GoTo Apl_Error
    End If
    
    '振替指示データ取得(FE) 2011/04/07追加 SETsw kubota
    sResult2 = ""
    RET = funCodeDBGet("SB", "FE", sProccd, 0, " ", sResult2)
    If RET <> 0 Then
        fncGetKouhoHinbanShiyou = -2
        GoTo Apl_Error
    End If
    '------------------------------------------ Make SQL ------------------------------------------------------
    sMakesql1 = ""
    sMakesql2 = ""
    sMakesql3 = ""
    sMakesql4 = ""
    sMakesql5 = ""
    sMakesql6 = ""
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    sMakesql9 = ""
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    sMakesql10 = ""     '06/10/05 ooba
    
    ''C－OSF3チェックの変更 2008.04.20 青柳
'>>>>> FCの11桁～20桁をFEに移動 2011/04/07 SETsw kubota ----------
'    sFlg_2_2 = Mid(sResult, 12, 1)
    sFlg_2_2 = Mid(sResult2, 2, 1)
'<<<<< FCの11桁～20桁をFEに移動 2011/04/07 SETsw kubota ----------
    
Debug.Print "(1-1 Make) "; Now
    '組み合わせ品番比較SQL文作成
    If Mid(sResult, 1, 1) = "1" Then
        sMakesql = ""
        RET = funGetKouhoHinban1_1(sProccd, sBlockId, sOld_Hinban, sMakesql)
        If RET <> 0 Then
            fncGetKouhoHinbanShiyou = RET
            GoTo Apl_Error
        End If
        sMakesql1 = "AND EXISTS (" & sMakesql & ") " & vbCrLf
    End If
Debug.Print "(1-2 Make) "; Now
    '常識仕様比較SQL文作成
    If Mid(sResult, 2, 1) = "1" Then
        sMakesql = ""
        RET = funGetKouhoHinban1_2(sProccd, sBlockId, sOld_Hinban, sMakesql)
        If RET <> 0 Then
            fncGetKouhoHinbanShiyou = RET
            GoTo Apl_Error
        End If
        sMakesql2 = "AND EXISTS (" & sMakesql & ") " & vbCrLf
    End If
Debug.Print "(1-3 Make) "; Now
    '外観実績を振替先品番比較SQL文作成
    If Mid(sResult, 3, 1) = "1" Then
        sMakesql = ""
        RET = funGetKouhoHinban1_3(sProccd, sBlockId, sOld_Hinban, sMakesql)
        If RET <> 0 Then
            fncGetKouhoHinbanShiyou = RET
            GoTo Apl_Error
        End If
        sMakesql3 = "AND EXISTS (" & sMakesql & ") " & vbCrLf
    End If
Debug.Print "(1-4 Make) "; Now
    '結晶評価項目仕様比較SQL文作成
    If Mid(sResult, 4, 1) = "1" Then
        sMakesql = ""
        RET = funGetKouhoHinban1_4(sProccd, sBlockId, sOld_Hinban, sMakesql)
        If RET <> 0 Then
            fncGetKouhoHinbanShiyou = RET
            GoTo Apl_Error
        End If
        sMakesql4 = "AND EXISTS (" & sMakesql & ") " & vbCrLf
    End If
Debug.Print "(1-5 Make) "; Now
    '先行評価項目仕様比較SQL文作成
    If Mid(sResult, 5, 1) = "1" Then
        sMakesql = ""
        RET = funGetKouhoHinban1_5(sProccd, sBlockId, sOld_Hinban, sMakesql)
        If RET <> 0 Then
            fncGetKouhoHinbanShiyou = RET
            GoTo Apl_Error
        End If
        sMakesql5 = "AND EXISTS (" & sMakesql & ") " & vbCrLf
    End If
Debug.Print "(1-6 Make) "; Now
    'ナノトポ規格比較SQL文作成
    If Mid(sResult, 6, 1) = "1" Then
        sMakesql = ""
        RET = funGetKouhoHinban1_6(sProccd, sBlockId, sOld_Hinban, sMakesql)
        If RET <> 0 Then
            fncGetKouhoHinbanShiyou = RET
            GoTo Apl_Error
        End If
        sMakesql6 = "AND EXISTS (" & sMakesql & ") " & vbCrLf
    End If
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
Debug.Print "(1-9 Make) "; Now
    'エピ先行評価項目仕様比較SQL文作成
    If Mid(sResult, 9, 1) = "1" Then
        sMakesql = ""
        RET = funGetKouhoHinban1_9(sProccd, sBlockId, sOld_Hinban, sMakesql)
        If RET <> 0 Then
            fncGetKouhoHinbanShiyou = RET
            GoTo Apl_Error
        End If
        sMakesql9 = "AND EXISTS (" & sMakesql & ") " & vbCrLf
    End If
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

Debug.Print "(1-10 Make) "; Now
    '常識仕様２比較SQL文作成　06/10/05 ooba
    If Mid(sResult, 10, 1) = "1" Then
        sMakesql = ""
        RET = funGetKouhoHinban1_10(sProccd, sBlockId, sOld_Hinban, sMakesql)
        If RET <> 0 Then
            fncGetKouhoHinbanShiyou = RET
            GoTo Apl_Error
        End If
        sMakesql10 = "AND EXISTS (" & sMakesql & ") " & vbCrLf
    End If
    
Debug.Print "(SQL文 Make) "; Now
    '------------------------------------------ SQL実行  ------------------------------------------------------
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E018A.HINBAN || TO_CHAR(E018A.MNOREVNO,'FM00')  || E018A.FACTORY || E018A.OPECOND AS HINBAN " & vbCrLf
    sql = sql & "FROM   TBCME018 E018A " & vbCrLf
    sql = sql & "WHERE  E018A.HINBAN || TO_CHAR(E018A.MNOREVNO, 'FM00') || E018A.FACTORY || E018A.OPECOND IN ( " & vbCrLf
    sql = sql & "       SELECT E018B.HINBAN || MAX(TO_CHAR(E018B.MNOREVNO, 'FM00') || E018B.FACTORY || E018B.OPECOND) " & vbCrLf
    sql = sql & "       FROM   TBCME018 E018B  " & vbCrLf
    sql = sql & "       WHERE  (E018B.SYNFLAG IS NULL OR E018B.SYNFLAG='1') " & vbCrLf
    sql = sql & "       GROUP BY E018B.HINBAN) AND " & vbCrLf
'    sql = sql & "      (E018A.SYNFLAG IS NULL OR E018A.SYNFLAG='1') AND " & vbCrLf
    sql = sql & "       E018A.HINBAN || TO_CHAR(E018A.MNOREVNO, 'FM00') || E018A.FACTORY || E018A.OPECOND   <>   '" & sOld_Hinban & "' " & vbCrLf
    sql = sql & sMakesql1
    sql = sql & sMakesql2
    sql = sql & sMakesql3
    sql = sql & sMakesql4
    sql = sql & sMakesql5
    sql = sql & sMakesql6
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    sql = sql & sMakesql9
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    sql = sql & sMakesql10  '06/10/05 ooba
    '2008/04/30 品番ソート追加 Kameda
    sql = sql & "ORDER BY HINBAN"
    On Error GoTo db_Error

Debug.Print "(SQL文 Start) "; Now
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
Debug.Print "(SQL文 End) "; Now
    
    '該当データなし
    If rs.EOF Then
        fncGetKouhoHinbanShiyou = 1
        GoTo db_Error
    End If
    
    '動的配列変数に対するメモリ領域の再割り当て
    ReDim sGet_Hinban(0 To rs.RecordCount - 1)
    
    '取得データセット
    For w_i = 0 To rs.RecordCount - 1
        With sGet_Hinban(w_i)
            .GETHINBAN = rs("HINBAN")          ' 振替候補品番
        End With
        rs.MoveNext
    Next
    Set rs = Nothing

    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Select Case iErr_Code
        Case 0      '正常終了
            sErr_Msg = ""
        Case 1      '正常終了（該当データなし）
            sErr_Msg = "振替可能な品番はありません。"
        Case -1
            sErr_Msg = "入力引数値にエラーがあります。"
        Case -2
            sErr_Msg = "振替指示データ取得でエラーが発生しました。"
        Case -3
            sErr_Msg = "データベースアクセスでエラーが発生しました。"
        Case -4
            sErr_Msg = "アプリケーションでエラーが発生しました。"
'------------------------------------------------------------------------------------
        Case 11001
            sErr_Msg = "1-1 仕様データ取得でエラーが発生しました。"
        Case 11002      '05/04/04 ooba
            sErr_Msg = "1-1 想定外の仕様データです。(ブロック単位保証フラグ)"
        Case 11003      '05/04/04 ooba
            sErr_Msg = "1-1 SQL文の編集でエラーが発生しました。(ブロック単位保証フラグ)"
'------------------------------------------------------------------------------------
        Case 12001
            sErr_Msg = "1-2 仕様データ取得でエラーが発生しました。"
        Case 12002
            sErr_Msg = "1-2 想定外の仕様データです。(溝位置方位)"
        Case 12003
            sErr_Msg = "1-2 SQL文の編集でエラーが発生しました。(溝位置方位)"
        Case 12004
            sErr_Msg = "1-2 想定外の仕様データです。(品種)"
        Case 12005
            sErr_Msg = "1-2 SQL文の編集でエラーが発生しました。(品種)"
        Case 12006
            sErr_Msg = "1-2 想定外の仕様データです。(引上げ速度)"
        Case 12007
            sErr_Msg = "1-2 SQL文の編集でエラーが発生しました。(引上げ速度)"
        Case 12008
            sErr_Msg = "1-2 想定外の仕様データです。(ＨＺタイプ)"
        Case 12009
            sErr_Msg = "1-2 SQL文の編集でエラーが発生しました。(ＨＺタイプ)"
        Case 12010
            sErr_Msg = "1-2 想定外の仕様データです。(ドローチューブ)"
        Case 12011
            sErr_Msg = "1-2 SQL文の編集でエラーが発生しました。(ドローチューブ)"
        Case 12012      '06/10/17 ooba
            sErr_Msg = "1-2 想定外の仕様データです。(結晶ドープ)"
        Case 12013      '06/10/17 ooba
            sErr_Msg = "1-2 SQL文の編集でエラーが発生しました。(結晶ドープ)"
'------------------------------------------------------------------------------------
        Case 13001
            sErr_Msg = "1-3 仕様データ取得でエラーが発生しました。"
        Case 13002
            sErr_Msg = "1-3 想定外の仕様データです。(Warpランク)"
        Case 13003
            sErr_Msg = "1-3 SQL文の編集でエラーが発生しました。(Warpランク)"
'------------------------------------------------------------------------------------
        Case 14001
            sErr_Msg = "1-4 仕様データ取得でエラーが発生しました。"
        Case 14010 To 14019
            sErr_Msg = "1-4 想定外の仕様データです。【比抵抗→" & funErrMsgGet(iErr_Code - 14010) & "】"
        Case 14020 To 14029
            sErr_Msg = "1-4 想定外の仕様データです。【酸素濃度→" & funErrMsgGet(iErr_Code - 14020) & "】"
        Case 14030 To 14039
            sErr_Msg = "1-4 想定外の仕様データです。【OSF1→" & funErrMsgGet(iErr_Code - 14030) & "】"
        Case 14040 To 14049
            sErr_Msg = "1-4 想定外の仕様データです。【OSF2→" & funErrMsgGet(iErr_Code - 14040) & "】"
        Case 14050 To 14059
            sErr_Msg = "1-4 想定外の仕様データです。【OSF3→" & funErrMsgGet(iErr_Code - 14050) & "】"
        Case 14060 To 14069
            sErr_Msg = "1-4 想定外の仕様データです。【OSF4→" & funErrMsgGet(iErr_Code - 14060) & "】"
        Case 14070 To 14079
            sErr_Msg = "1-4 想定外の仕様データです。【BMD1→" & funErrMsgGet(iErr_Code - 14070) & "】"
        Case 14080 To 14089
            sErr_Msg = "1-4 想定外の仕様データです。【BMD2→" & funErrMsgGet(iErr_Code - 14080) & "】"
        Case 14090 To 14099
            sErr_Msg = "1-4 想定外の仕様データです。【BMD3→" & funErrMsgGet(iErr_Code - 14090) & "】"
        Case 14100 To 14109
            sErr_Msg = "1-4 想定外の仕様データです。【EPD→" & funErrMsgGet(iErr_Code - 14100) & "】"
        Case 14110 To 14119
            sErr_Msg = "1-4 想定外の仕様データです。【ﾗｲﾌﾀｲﾑ→" & funErrMsgGet(iErr_Code - 14110) & "】"
        Case 14120 To 14129
            sErr_Msg = "1-4 想定外の仕様データです。【炭素濃度→" & funErrMsgGet(iErr_Code - 14120) & "】"
        Case 14130 To 14139
            sErr_Msg = "1-4 想定外の仕様データです。【DEN→" & funErrMsgGet(iErr_Code - 14130) & "】"
        Case 14140 To 14149
            sErr_Msg = "1-4 想定外の仕様データです。【DVD2→" & funErrMsgGet(iErr_Code - 14140) & "】"
        Case 14150 To 14159
            sErr_Msg = "1-4 想定外の仕様データです。【L/DL→" & funErrMsgGet(iErr_Code - 14150) & "】"
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
        Case 14160 To 14169
            sErr_Msg = "1-4 想定外の仕様データです。【SIRD→" & funErrMsgGet(iErr_Code - 14160) & "】"
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
'------------------------------------------------------------------------------------
        Case 15001
            sErr_Msg = "1-5 仕様データ取得でエラーが発生しました。"
        Case 15010 To 15019
            sErr_Msg = "1-5 想定外の仕様データです。【比抵抗→" & funErrMsgGet(iErr_Code - 15010) & "】"
        Case 15020 To 15029
            sErr_Msg = "1-5 想定外の仕様データです。【酸素濃度→" & funErrMsgGet(iErr_Code - 15020) & "】"
        Case 15030 To 15039
            sErr_Msg = "1-5 想定外の仕様データです。【OSF1→" & funErrMsgGet(iErr_Code - 15030) & "】"
        Case 15040 To 15049
            sErr_Msg = "1-5 想定外の仕様データです。【OSF2→" & funErrMsgGet(iErr_Code - 15040) & "】"
        Case 15050 To 15059
            sErr_Msg = "1-5 想定外の仕様データです。【OSF3→" & funErrMsgGet(iErr_Code - 15050) & "】"
        Case 15060 To 15069
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''            sErr_Msg = "1-5 想定外の仕様データです。【OSF4→" & funErrMsgGet(iErr_Code - 15060) & "】"
            sErr_Msg = "1-5 想定外の仕様データです。【SIRD→" & funErrMsgGet(iErr_Code - 15060) & "】"
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
        Case 15070 To 15079
            sErr_Msg = "1-5 想定外の仕様データです。【BMD1→" & funErrMsgGet(iErr_Code - 15070) & "】"
        Case 15080 To 15089
            sErr_Msg = "1-5 想定外の仕様データです。【BMD2→" & funErrMsgGet(iErr_Code - 15080) & "】"
        Case 15090 To 15099
            sErr_Msg = "1-5 想定外の仕様データです。【BMD3→" & funErrMsgGet(iErr_Code - 15080) & "】"
        Case 15100 To 15109
            sErr_Msg = "1-5 想定外の仕様データです。【酸素析出1→" & funErrMsgGet(iErr_Code - 15100) & "】"
        Case 15110 To 15119
            sErr_Msg = "1-5 想定外の仕様データです。【酸素析出2→" & funErrMsgGet(iErr_Code - 15110) & "】"
        Case 15120 To 15129
            sErr_Msg = "1-5 想定外の仕様データです。【酸素析出3→" & funErrMsgGet(iErr_Code - 15120) & "】"
        Case 15130 To 15139
            sErr_Msg = "1-5 想定外の仕様データです。【DSOD→" & funErrMsgGet(iErr_Code - 15130) & "】"
        Case 15140 To 15149
            sErr_Msg = "1-5 想定外の仕様データです。【DZ→" & funErrMsgGet(iErr_Code - 15140) & "】"
        Case 15150 To 15159
            sErr_Msg = "1-5 想定外の仕様データです。【SPVFE→" & funErrMsgGet(iErr_Code - 15150) & "】"
        Case 15160 To 15169
            sErr_Msg = "1-5 想定外の仕様データです。【拡散長→" & funErrMsgGet(iErr_Code - 15160) & "】"
        Case 15170 To 15179         '残存酸素追加　03/12/09 ooba
            sErr_Msg = "1-5 想定外の仕様データです。【残存酸素→" & funErrMsgGet(iErr_Code - 15170) & "】"
        Case 15180 To 15189         'GD-Den追加　05/01/27 ooba
            sErr_Msg = "1-5 想定外の仕様データです。【DEN→" & funErrMsgGet(iErr_Code - 15180) & "】"
        Case 15190 To 15199         'GD-DVD2追加　05/01/27 ooba
            sErr_Msg = "1-5 想定外の仕様データです。【DVD2→" & funErrMsgGet(iErr_Code - 15190) & "】"
        Case 15200 To 15209         'GD-L/DL追加　05/01/27 ooba
            sErr_Msg = "1-5 想定外の仕様データです。【L/DL→" & funErrMsgGet(iErr_Code - 15200) & "】"
        Case 15210 To 15219         'SPVNR追加　06/05/31 ooba
            sErr_Msg = "1-5 想定外の仕様データです。【SPVNR→" & funErrMsgGet(iErr_Code - 15210) & "】"
'------------------------------------------------------------------------------------
        Case 16001
            sErr_Msg = "1-6 仕様データ取得でエラーが発生しました。"
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
        Case 19001
            sErr_Msg = "1-9 仕様データ取得でエラーが発生しました。"
        Case 19010 To 19019
            sErr_Msg = "1-9 想定外の仕様データです。【OSF1E→" & funErrMsgGet(iErr_Code - 19010) & "】"
        Case 19020 To 19029
            sErr_Msg = "1-9 想定外の仕様データです。【OSF2E→" & funErrMsgGet(iErr_Code - 19020) & "】"
        Case 19030 To 19039
            sErr_Msg = "1-9 想定外の仕様データです。【OSF3E→" & funErrMsgGet(iErr_Code - 19030) & "】"
        Case 19040 To 19049
            sErr_Msg = "1-9 想定外の仕様データです。【BMD1E→" & funErrMsgGet(iErr_Code - 19040) & "】"
        Case 19050 To 19059
            sErr_Msg = "1-9 想定外の仕様データです。【BMD2E→" & funErrMsgGet(iErr_Code - 19050) & "】"
        Case 19060 To 19069
            sErr_Msg = "1-9 想定外の仕様データです。【BMD3E→" & funErrMsgGet(iErr_Code - 19060) & "】"
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
'------------------------------------------------------------------------------------
        '10001～10013　06/10/05 ooba
        Case 10001
            sErr_Msg = "1-10 仕様データ取得でエラーが発生しました。"
        Case 10002
            sErr_Msg = "1-10 想定外の仕様データです。(溝位置方位)"
        Case 10003
            sErr_Msg = "1-10 SQL文の編集でエラーが発生しました。(溝位置方位)"
        Case 10004
            sErr_Msg = "1-10 想定外の仕様データです。(品種)"
        Case 10005
            sErr_Msg = "1-10 SQL文の編集でエラーが発生しました。(品種)"
        Case 10006
            sErr_Msg = "1-10 想定外の仕様データです。(引上げ速度)"
        Case 10007
            sErr_Msg = "1-10 SQL文の編集でエラーが発生しました。(引上げ速度)"
        Case 10008
            sErr_Msg = "1-10 想定外の仕様データです。(ＨＺタイプ)"
        Case 10009
            sErr_Msg = "1-10 SQL文の編集でエラーが発生しました。(ＨＺタイプ)"
        Case 10010
            sErr_Msg = "1-10 想定外の仕様データです。(ドローチューブ)"
        Case 10011
            sErr_Msg = "1-10 SQL文の編集でエラーが発生しました。(ドローチューブ)"
        Case 10012
            sErr_Msg = "1-10 想定外の仕様データです。(結晶ドープ)"
        Case 10013
            sErr_Msg = "1-10 SQL文の編集でエラーが発生しました。(結晶ドープ)"
'------------------------------------------------------------------------------------
    End Select

    If iErr_Code > 10000 Then sErr_Msg = sErr_Msg & "(" & sOld_Hinban & ")"

    Exit Function
    
Apl_Error:
    iErr_Code = fncGetKouhoHinbanShiyou
    GoTo Apl_Exit

Apl_down:
    fncGetKouhoHinbanShiyou = -4
    iErr_Code = fncGetKouhoHinbanShiyou
    GoTo Apl_Exit

db_Error:
    Set rs = Nothing
    If fncGetKouhoHinbanShiyou = 0 Then
        fncGetKouhoHinbanShiyou = -3
    End If
    iErr_Code = fncGetKouhoHinbanShiyou
    GoTo Apl_Exit
    
End Function

'------------------------------------------------
' 組み合わせ品番比較SQL文作成
'------------------------------------------------

'概要      :指定された振替元品番と、以下の仕様値が一致している品番を抽出するSQL文を作成し、呼び出し元に返す。
'           （タイプ、ブロック単位保障フラグ）
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sBlockid        ,I  ,String       :ﾌﾞﾛｯｸID、又は、結晶番号
'          :sOld_Hinban     ,I  ,String       :振替元品番
'          :sMakeSql        ,O  ,String       :作成SQL文
'          :戻り値          ,O  ,Integer      :取得の成否(0:正常取得, -1:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/10 新規作成　SB

Public Function funGetKouhoHinban1_1(sProccd As String, BLOCKID As String, sOld_Hinban As String, sMakesql As String) As Integer


    Dim sql As String       'SQL全体
    Dim rs  As OraDynaset   'RecordSet
    Dim RET     As Integer  '戻り値                         '05/04/04 ooba START ============>
    Dim sResult As String   'コードＤＢ取得関数の取得変数
    Dim sinstr  As String   'ＳＱＬin句用文字列
    Dim sinstr1 As String   'ＳＱＬin句用文字列              '05/04/04 ooba END ==============>
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funGetKouhoHinban1_1 = 0
    
    '------------------------------------------ 振替元品種仕様データ取得 ------------------------------------------------------
    'SQL文の作成
'>>>>> TBCME021.HWFTYPE→TBCME018.HSXTYPE 2011/05/11 SETsw kubota ---------
'    sql = vbNullString
'    sql = sql & "SELECT E021.HWFTYPE,E036.BLOCKHFLAG " & vbCrLf
'    sql = sql & "FROM   TBCME021 E021,TBCME036 E036 " & vbCrLf
'    sql = sql & "WHERE  E021.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
'    sql = sql & "       TO_CHAR(E021.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
'    sql = sql & "       E021.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
'    sql = sql & "       E021.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' AND " & vbCrLf
'    sql = sql & "       E036.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
'    sql = sql & "       TO_CHAR(E036.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
'    sql = sql & "       E036.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
'    sql = sql & "       E036.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' " & vbCrLf
''    sql = sql & "WHERE  E021.HINBAN || TO_CHAR(E021.MNOREVNO, 'FM00') || E021.FACTORY || E021.OPECOND   =   '" & sOld_Hinban & "' " & vbCrLf
''    sql = sql & "       E036.HINBAN || TO_CHAR(E036.MNOREVNO, 'FM00') || E036.FACTORY || E036.OPECOND   =   '" & sOld_Hinban & "' " & vbCrLf
    sql = vbNullString
    sql = sql & "SELECT E018.HSXTYPE,E036.BLOCKHFLAG " & vbCrLf
    sql = sql & "FROM   TBCME018 E018,TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E018.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E018.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' AND " & vbCrLf
    sql = sql & "       E036.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E036.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E036.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' " & vbCrLf
'<<<<< TBCME021.HWFTYPE→TBCME018.HSXTYPE 2011/05/11 SETsw kubota ---------
    
    On Error GoTo db_Error
    'SQL文の実行
    Debug.Print sql
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        funGetKouhoHinban1_1 = 11001
        GoTo db_Error
    End If
    
    '取得データセット
    Erase tbl_spec1_1
    With tbl_spec1_1(0)
'        If IsNull(rs("HWFTYPE")) = False Then .HWFTYPE = rs("HWFTYPE") Else .HWFTYPE = " "                  'ﾀｲﾌﾟ
        If IsNull(rs("HSXTYPE")) = False Then .HSXTYPE = rs("HSXTYPE") Else .HSXTYPE = " "                  'ﾀｲﾌﾟ
        If IsNull(rs("BLOCKHFLAG")) = False Then .BLOCKHFLAG = rs("BLOCKHFLAG") Else .BLOCKHFLAG = " "      'ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞ
    End With
    
    Set rs = Nothing
    On Error GoTo Apl_down
    
    ''ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞの振替ﾁｪｯｸ変更  05/04/04 ooba START ======================================>
    sResult = ""
    sinstr = ""
    RET = funCodeDBGet("SB", "BH", tbl_spec1_1(0).BLOCKHFLAG, 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_1 = 11002
        GoTo Apl_Exit
    End If
    RET = funinfo2(sResult, sinstr)
    If RET <> 0 Then
        funGetKouhoHinban1_1 = 11003
        GoTo Apl_Exit
    End If
    RET = funCodeinGet("SB", "BH", sinstr, sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_1 = 11003
        GoTo Apl_Exit
    End If
    sinstr1 = sResult
    ''ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞの振替ﾁｪｯｸ変更  05/04/04 ooba END ========================================>
    
'>>>>> PN不問品番への振替ロック解除 2011/05/11 SETsw kubota ---------
'    '------------------------------------------ 振替元品種と同一仕様の品番を取得 ------------------------------------------------------
'    'SQL文の作成
'    sql = vbNullString
'    sql = sql & "SELECT 'X' " & vbCrLf
''    sql = sql & "SELECT E021.HINBAN || TO_CHAR(E021.MNOREVNO, 'FM00') || E021.FACTORY || E021.OPECOND HINBAN " & vbCrLf
'    sql = sql & "FROM   TBCME021 E021, TBCME036 E036 " & vbCrLf
'    sql = sql & "WHERE  E018A.HINBAN                    = E021.HINBAN                       AND " & vbCrLf
'    sql = sql & "       TO_CHAR(E018A.MNOREVNO, 'FM00') = TO_CHAR(E021.MNOREVNO, 'FM00')    AND " & vbCrLf
'    sql = sql & "       E018A.FACTORY                   = E021.FACTORY                      AND " & vbCrLf
'    sql = sql & "       E018A.OPECOND                   = E021.OPECOND                      AND " & vbCrLf
'    sql = sql & "       E021.HINBAN                     = E036.HINBAN                       AND " & vbCrLf
'    sql = sql & "       TO_CHAR(E021.MNOREVNO, 'FM00')  = TO_CHAR(E036.MNOREVNO, 'FM00')    AND " & vbCrLf
'    sql = sql & "       E021.FACTORY                    = E036.FACTORY                      AND " & vbCrLf
'    sql = sql & "       E021.OPECOND                    = E036.OPECOND                      AND " & vbCrLf
''    sql = sql & "WHERE  E021.HINBAN || TO_CHAR(E021.MNOREVNO, 'FM00') || E021.FACTORY || E021.OPECOND <> '" & sOld_Hinban & "' AND " & vbCrLf
''    sql = sql & "       E021.HINBAN || TO_CHAR(E021.MNOREVNO, 'FM00') || E021.FACTORY || E021.OPECOND = " & vbCrLf
''    sql = sql & "       E036.HINBAN || TO_CHAR(E036.MNOREVNO, 'FM00') || E036.FACTORY || E036.OPECOND AND " & vbCrLf
'    sql = sql & "       E021.HWFTYPE                    = '" & tbl_spec1_1(0).HWFTYPE & "'  AND " & vbCrLf
''    sql = sql & "       E036.BLOCKHFLAG                 = '" & tbl_spec1_1(0).BLOCKHFLAG & "' " & vbCrLf
'    sql = sql & "       E036.BLOCKHFLAG                 IN (" & sinstr1 & ") " & vbCrLf     '05/04/04 ooba

    '----------------------- 振替元品種と同一仕様or不問の品番を取得 ------------------------------------------------------
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT 'X' " & vbCrLf
    sql = sql & "FROM   TBCME018 E018, TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E018A.HINBAN                    = E018.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018A.MNOREVNO, 'FM00') = TO_CHAR(E018.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E018A.FACTORY                   = E018.FACTORY                      AND " & vbCrLf
    sql = sql & "       E018A.OPECOND                   = E018.OPECOND                      AND " & vbCrLf
    sql = sql & "       E018.HINBAN                     = E036.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018.MNOREVNO, 'FM00')  = TO_CHAR(E036.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E018.FACTORY                    = E036.FACTORY                      AND " & vbCrLf
    sql = sql & "       E018.OPECOND                    = E036.OPECOND                      AND " & vbCrLf
    
    'タイプが同一又は不問の品番を取得
    sql = sql & "     ( E018.HSXTYPE = '" & tbl_spec1_1(0).HSXTYPE & "'" & vbCrLf
    sql = sql & "    OR E018.HSXTYPE = 'Z' )  AND " & vbCrLf
    
    sql = sql & "       E036.BLOCKHFLAG                 IN (" & sinstr1 & ") " & vbCrLf     '05/04/04 ooba
'<<<<< PN不問品番への振替ロック解除 2011/05/11 SETsw kubota ---------

    sMakesql = sql

'    On Error GoTo db_Error
'    'SQL文の実行
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'
'    '該当データなし
'    If rs.EOF Then
'        funGetKouhoHinban1_1 = 1
'        GoTo db_Error
'    Else
'        sMakesql = sql
'    End If
'
'    Set rs = Nothing
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funGetKouhoHinban1_1 = -4
    GoTo Apl_Exit
    
db_Error:
    Set rs = Nothing
    If funGetKouhoHinban1_1 = 0 Then
        funGetKouhoHinban1_1 = -3
    End If
    GoTo Apl_Exit

End Function
    
'------------------------------------------------
' 振替先と振替元の常識仕様チェック
'------------------------------------------------

'概要      :指定された振替元品番と、以下の仕様値が一致している品番を抽出するSQL文を作成し、呼び出し元に返す。
'           （結晶面方位、結晶面傾中心、ドーパント、結晶ドープ、シード傾き）
'           指定された振替元品番と、以下の仕様値がマトリクスで一致している品番を抽出するSQL文を作成する。
'           （溝位置方位、品種、引上速度、ＨＺタイプ、ドローチューブ）
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sBlockid        ,I  ,String       :ﾌﾞﾛｯｸID、又は、結晶番号
'          :sOld_Hinban     ,I  ,String       :振替元品番
'          :sMakeSql        ,O  ,String       :作成SQL文
'          :戻り値          ,O  ,Integer      :取得の成否(0:正常取得, -1:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/10 新規作成　SB

Public Function funGetKouhoHinban1_2(sProccd As String, sBlockId As String, sOld_Hinban As String, sMakesql As String) As Integer


    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql         As String       'SQL全体
    Dim rs         As OraDynaset   'RecordSet
    Dim sinstr     As String       'ＳＱＬin句用文字列
    Dim sinstr1    As String       'ＳＱＬin句用文字列
    Dim sinstr2    As String       'ＳＱＬin句用文字列
    Dim sinstr3    As String       'ＳＱＬin句用文字列
    Dim sinstr4    As String       'ＳＱＬin句用文字列
    Dim sinstr5    As String       'ＳＱＬin句用文字列
    Dim sinstr6    As String       'ＳＱＬin句用文字列
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funGetKouhoHinban1_2 = 0
    
    '------------------------------------------ 振替元品種仕様データ取得 ------------------------------------------------------
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E018.HSXCDIR,E018.HSXCSCEN,E018.HSXDOP,E023.HWFCDOP,E020.HSXSDSLP,E018.HSXDPDIR, " & vbCrLf
    sql = sql & "       SUBSTR(E018.MCNO,1,1) MCNO1,SUBSTR(E018.MCNO,4,1) MCNO2,SUBSTR(E018.MCNO,3,1) MCNO3,E036.DCHYUUBU " & vbCrLf
    sql = sql & "FROM   TBCME018 E018,TBCME023 E023,TBCME020 E020,TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E018.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E018.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' AND " & vbCrLf
    sql = sql & "       E023.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E023.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E023.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E023.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' AND " & vbCrLf
    sql = sql & "       E020.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E020.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E020.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E020.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' AND " & vbCrLf
    sql = sql & "       E036.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E036.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E036.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' " & vbCrLf
'    sql = sql & "WHERE  E018.HINBAN || TO_CHAR(E018.MNOREVNO, 'FM00') || E018.FACTORY || E018.OPECOND   =   '" & sOld_Hinban & "' AND " & vbCrLf
'    sql = sql & "       E023.HINBAN || TO_CHAR(E023.MNOREVNO, 'FM00') || E023.FACTORY || E023.OPECOND   =   '" & sOld_Hinban & "' AND " & vbCrLf
'    sql = sql & "       E020.HINBAN || TO_CHAR(E020.MNOREVNO, 'FM00') || E020.FACTORY || E020.OPECOND   =   '" & sOld_Hinban & "' AND " & vbCrLf
'    sql = sql & "       E036.HINBAN || TO_CHAR(E036.MNOREVNO, 'FM00') || E036.FACTORY || E036.OPECOND   =   '" & sOld_Hinban & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        funGetKouhoHinban1_2 = 12001
        GoTo db_Error
    End If
    
    '取得データセット
    Erase tbl_spec1_2
    With tbl_spec1_2(0)
        If IsNull(rs("HSXCDIR")) = False Then .HSXCDIR = rs("HSXCDIR") Else .HSXCDIR = " "          ' 結晶面方位
        If IsNull(rs("HSXCSCEN")) = False Then .HSXCSCEN = rs("HSXCSCEN") Else .HSXCSCEN = -1       ' 結晶面傾き中心
        If IsNull(rs("HSXDOP")) = False Then .HSXDOP = rs("HSXDOP") Else .HSXDOP = " "              ' ドーパント
        If IsNull(rs("HWFCDOP")) = False Then .HWFCDOP = rs("HWFCDOP") Else .HWFCDOP = " "          ' 結晶ドープ
        If IsNull(rs("HSXSDSLP")) = False Then .HSXSDSLP = rs("HSXSDSLP") Else .HSXSDSLP = " "      ' シード傾き
        If IsNull(rs("HSXDPDIR")) = False Then .HSXDPDIR = rs("HSXDPDIR") Else .HSXDPDIR = " "      ' 溝位置方位
        If IsNull(rs("MCNO1")) = False Then .MCNO1 = rs("MCNO1") Else .MCNO1 = " "                  ' 品種
        If IsNull(rs("MCNO2")) = False Then .MCNO2 = rs("MCNO2") Else .MCNO2 = " "                  ' 引上げ速度
        If IsNull(rs("MCNO3")) = False Then .MCNO3 = rs("MCNO3") Else .MCNO3 = " "                  ' HZタイプ
        If IsNull(rs("DCHYUUBU")) = False Then .DCHYUUBU = rs("DCHYUUBU") Else .DCHYUUBU = " "      ' ドローチューブ
    End With
    
    Set rs = Nothing
    On Error GoTo Apl_down
    '------------------------------------------ 指示取得 ------------------------------------------------------
    sinstr1 = ""
    sinstr2 = ""
    sinstr3 = ""
    sinstr4 = ""
    sinstr5 = ""
    sinstr6 = ""
    '溝位置方位
    sResult = ""
    sinstr = ""
    RET = funCodeDBGet("SB", "MZ", tbl_spec1_2(0).HSXDPDIR, 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_2 = 12002
        GoTo Apl_Exit
    End If
    RET = funinfo2(sResult, sinstr)
    If RET <> 0 Then
        funGetKouhoHinban1_2 = 12003
        GoTo Apl_Exit
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↓追加 2003/10/21
    RET = funCodeinGet("SB", "MZ", sinstr, sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_2 = 12003
        GoTo Apl_Exit
    End If
    sinstr1 = sResult
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↑追加 2003/10/21
'    sinstr1 = sinstr
    '品種
    sResult = ""
    sinstr = ""
    RET = funCodeDBGet("SB", "HS", tbl_spec1_2(0).MCNO1, 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_2 = 12004
        GoTo Apl_Exit
    End If
    RET = funinfo2(sResult, sinstr)
    If RET <> 0 Then
        funGetKouhoHinban1_2 = 12005
        GoTo Apl_Exit
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↓追加 2003/10/21
    RET = funCodeinGet("SB", "HS", sinstr, sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_2 = 12005
        GoTo Apl_Exit
    End If
    sinstr2 = sResult
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↑追加 2003/10/21
'    sinstr2 = sinstr
    '引上げ速度
    sResult = ""
    sinstr = ""
    RET = funCodeDBGet("SB", "HK", tbl_spec1_2(0).MCNO2, 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_2 = 12006
        GoTo Apl_Exit
    End If
    RET = funinfo2(sResult, sinstr)
    If RET <> 0 Then
        funGetKouhoHinban1_2 = 12007
        GoTo Apl_Exit
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↓追加 2003/10/21
    RET = funCodeinGet("SB", "HK", sinstr, sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_2 = 12007
        GoTo Apl_Exit
    End If
    sinstr3 = sResult
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↑追加 2003/10/21
    'sinstr3 = sinstr
    'ＨＺタイプ
    sResult = ""
    sinstr = ""
    RET = funCodeDBGet("SB", "HZ", tbl_spec1_2(0).MCNO3, 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_2 = 12008
        GoTo Apl_Exit
    End If
    RET = funinfo2(sResult, sinstr)
    If RET <> 0 Then
        funGetKouhoHinban1_2 = 12009
        GoTo Apl_Exit
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↓追加 2003/10/21
    RET = funCodeinGet("SB", "HZ", sinstr, sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_2 = 12009
        GoTo Apl_Exit
    End If
    sinstr4 = sResult
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↑追加 2003/10/21
'    sinstr4 = sinstr
    'ドローチューブ
    sResult = ""
    sinstr = ""
    If tbl_spec1_2(0).DCHYUUBU <> " " Then
        RET = funCodeDBGet("SB", "DC", tbl_spec1_2(0).DCHYUUBU, 0, " ", sResult)
    Else
        RET = funCodeDBGet("SB", "DC", "2", 0, " ", sResult)
    End If
    If RET <> 0 Then
        funGetKouhoHinban1_2 = 12010
        GoTo Apl_Exit
    End If
    RET = funinfo2(sResult, sinstr)
    If RET <> 0 Then
        funGetKouhoHinban1_2 = 12011
        GoTo Apl_Exit
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↓追加 2003/10/21
    RET = funCodeinGet("SB", "DC", sinstr, sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_2 = 12011
        GoTo Apl_Exit
    End If
    sinstr5 = sResult
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↑追加 2003/10/21
'    sinstr5 = sinstr
    '結晶ドープ
    sResult = ""
    sinstr = ""
    RET = funCodeDBGet("SB", "SD", tbl_spec1_2(0).HWFCDOP, 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_2 = 12012
        GoTo Apl_Exit
    End If
    RET = funinfo2(sResult, sinstr)
    If RET <> 0 Then
        funGetKouhoHinban1_2 = 12013
        GoTo Apl_Exit
    End If
    RET = funCodeinGet("SB", "SD", sinstr, sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_2 = 12013
        GoTo Apl_Exit
    End If
    sinstr6 = sResult
    
'    sinstr5 = sinstr
    '------------------------------------------ 振替元品種と同一仕様の品番を取得 ------------------------------------------------------
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT 'X' " & vbCrLf
'    sql = sql & "SELECT E018.HINBAN || TO_CHAR(E018.MNOREVNO, 'FM00') || E018.FACTORY | |E018.OPECOND HINBAN " & vbCrLf
    sql = sql & "FROM   TBCME018 E018, TBCME023 E023, TBCME036 E036, TBCME020 E020 " & vbCrLf
    sql = sql & "WHERE  E018A.HINBAN                    = E018.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018A.MNOREVNO, 'FM00') = TO_CHAR(E018.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E018A.FACTORY                   = E018.FACTORY                      AND " & vbCrLf
    sql = sql & "       E018A.OPECOND                   = E018.OPECOND                      AND " & vbCrLf
    sql = sql & "       E018.HINBAN                     = E023.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018.MNOREVNO, 'FM00')  = TO_CHAR(E023.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E018.FACTORY                    = E023.FACTORY                      AND " & vbCrLf
    sql = sql & "       E018.OPECOND                    = E023.OPECOND                      AND " & vbCrLf
    sql = sql & "       E018.HINBAN                     = E036.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018.MNOREVNO, 'FM00')  = TO_CHAR(E036.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E018.FACTORY                    = E036.FACTORY                      AND " & vbCrLf
    sql = sql & "       E018.OPECOND                    = E036.OPECOND                      AND " & vbCrLf
    sql = sql & "       E018.HINBAN                     = E020.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018.MNOREVNO, 'FM00')  = TO_CHAR(E020.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E018.FACTORY                    = E020.FACTORY                      AND " & vbCrLf
    sql = sql & "       E018.OPECOND                    = E020.OPECOND                      AND " & vbCrLf
'    sql = sql & "WHERE  E018.HINBAN || TO_CHAR(E018.MNOREVNO, 'FM00') || E018.FACTORY || E018.OPECOND <> '" & sOld_Hinban & "' AND " & vbCrLf
'    sql = sql & "       E018.HINBAN || TO_CHAR(E018.MNOREVNO, 'FM00') || E018.FACTORY || E018.OPECOND = E023.HINBAN || TO_CHAR(E023.MNOREVNO, 'FM00') || E023.FACTORY || E023.OPECOND AND " & vbCrLf
'    sql = sql & "       E018.HINBAN || TO_CHAR(E018.MNOREVNO, 'FM00') || E018.FACTORY || E018.OPECOND = E036.HINBAN || TO_CHAR(E036.MNOREVNO, 'FM00') || E036.FACTORY || E036.OPECOND AND " & vbCrLf
'    sql = sql & "       E018.HINBAN || TO_CHAR(E018.MNOREVNO, 'FM00') || E018.FACTORY || E018.OPECOND = E020.HINBAN || TO_CHAR(E020.MNOREVNO, 'FM00') || E020.FACTORY || E020.OPECOND AND " & vbCrLf
    sql = sql & "       E018.HSXCDIR                    = '" & tbl_spec1_2(0).HSXCDIR & "'  AND " & vbCrLf
'    sql = sql & "       E023.HWFCDOP                    = '" & tbl_spec1_2(0).HWFCDOP & "'  AND " & vbCrLf
'    sql = sql & "       E018.HSXCSCEN                   =  " & tbl_spec1_2(0).HSXCSCEN & "  AND " & vbCrLf
''2008/11/27 結晶面傾中心チェック緩和(2) DEL By Systech Start
''    If tbl_spec1_2(0).HSXCSCEN = -1 Then
''        sql = sql & "       (E018.HSXCSCEN is null OR E018.HSXCSCEN = 0 OR E018.HSXCSCEN = 0.43) AND " & vbCrLf
''    ElseIf (tbl_spec1_2(0).HSXCSCEN = 0) Or (tbl_spec1_2(0).HSXCSCEN = 0.43) Then
''        sql = sql & "       (E018.HSXCSCEN = 0 OR E018.HSXCSCEN = 0.43) AND " & vbCrLf
''    Else
''        sql = sql & "       E018.HSXCSCEN                   =  " & tbl_spec1_6(0).HSXCSCEN & "  AND " & vbCrLf
''    End If
''2008/11/27 結晶面傾中心チェック緩和(2) DEL By Systech End
    
    sql = sql & "       E020.HSXSDSLP                   = '" & tbl_spec1_2(0).HSXSDSLP & "' AND " & vbCrLf
    
'>>>>> ドーパント不問品番への振替ロック解除 2011/05/12 SETsw kubota ---------
'    sql = sql & "       E018.HSXDOP                     = '" & tbl_spec1_2(0).HSXDOP & "'   AND " & vbCrLf
    'ドーパントが同一又は不問の品番を取得
    sql = sql & "     ( E018.HSXDOP = '" & tbl_spec1_2(0).HSXDOP & "'" & vbCrLf
    sql = sql & "    OR E018.HSXDOP = 'Z' )  AND " & vbCrLf
'<<<<< ドーパント不問品番への振替ロック解除 2011/05/12 SETsw kubota ---------

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↓追加 2003/10/21
'    sql = sql & "       E018.HSXDPDIR               IN (SELECT NVL(TRIM(CODE),CHR(32)) FROM TBCMB005 WHERE SYSCLASS = 'SB' AND CLASS = 'MZ' AND INFO2 in (" & sinstr1 & ")) AND " & vbCrLf
'    sql = sql & "       substr(E018.MCNO, 1, 1)     IN (SELECT NVL(TRIM(CODE),CHR(32)) FROM TBCMB005 WHERE SYSCLASS = 'SB' AND CLASS = 'HS' AND INFO2 in (" & sinstr2 & ")) AND " & vbCrLf
'    sql = sql & "       substr(E018.MCNO, 4, 1)     IN (SELECT NVL(TRIM(CODE),CHR(32)) FROM TBCMB005 WHERE SYSCLASS = 'SB' AND CLASS = 'HK' AND INFO2 in (" & sinstr3 & ")) AND " & vbCrLf
'    sql = sql & "       substr(E018.MCNO, 3, 1)     IN (SELECT NVL(TRIM(CODE),CHR(32)) FROM TBCMB005 WHERE SYSCLASS = 'SB' AND CLASS = 'HZ' AND INFO2 in (" & sinstr4 & ")) AND " & vbCrLf
'    sql = sql & "       E036.DCHYUUBU               IN (SELECT NVL(TRIM(CODE),CHR(32)) FROM TBCMB005 WHERE SYSCLASS = 'SB' AND CLASS = 'DC' AND INFO2 in (" & sinstr5 & ")) " & vbCrLf
    sql = sql & "       E018.HSXDPDIR               IN (" & sinstr1 & ") AND " & vbCrLf
    sql = sql & "       substr(E018.MCNO, 1, 1)     IN (" & sinstr2 & ") AND " & vbCrLf
    sql = sql & "       substr(E018.MCNO, 4, 1)     IN (" & sinstr3 & ") AND " & vbCrLf
    sql = sql & "       substr(E018.MCNO, 3, 1)     IN (" & sinstr4 & ") AND " & vbCrLf
    If tbl_spec1_2(0).DCHYUUBU = " " Then
        sql = sql & "       E036.DCHYUUBU is null OR E036.DCHYUUBU IN (" & sinstr5 & ") " & vbCrLf
    Else
        sql = sql & "       E036.DCHYUUBU               IN (" & sinstr5 & ")     " & vbCrLf
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↑追加 2003/10/21
    sql = sql & " AND   E023.HWFCDOP                IN (" & sinstr6 & ")     " & vbCrLf     '06/10/17 ooba
    
    sMakesql = sql
    
'    On Error GoTo db_Error
'    'SQL文の実行
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'
'    '該当データなし
'    If rs.EOF Then
'        funGetKouhoHinban1_2 = 1
'        GoTo db_Error
'    Else
'        sMakesql = sql
'    End If
'
'    Set rs = Nothing
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funGetKouhoHinban1_2 = -4
    GoTo Apl_Exit
    
db_Error:
    Set rs = Nothing
    If funGetKouhoHinban1_2 = 0 Then
        funGetKouhoHinban1_2 = -3
    End If
    GoTo Apl_Exit

End Function

'------------------------------------------------
' SQLin句文作成
'------------------------------------------------

'概要      :マトリクス取得データよりSQLin句文を作成し、呼び出し元に返す。
'           （ワープランク）
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sResult         ,I  ,String       :マトリクス取得データ
'          :sinstr          ,O  ,String       :SQLin句文
'          :戻り値          ,O  ,Integer      :取得の成否(0:正常取得, -1:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/10 新規作成　SB

Public Function funinfo2(sResult, sinstr) As Integer
    Dim W_POS As Long             'ポジション取得
    Dim W_STARTPOS  As Long       'スタートポジション
    W_STARTPOS = 1
    Do Until W_STARTPOS > Len(sResult)
        W_POS = InStr(W_STARTPOS, sResult, "1")
        If W_POS = 0 Then
            W_STARTPOS = W_STARTPOS + 1
        Else
            If sinstr = "" Then
                sinstr = "'" & W_POS & "'"
            Else
                sinstr = sinstr & "," & "'" & W_POS & "'"
            End If
            W_STARTPOS = W_POS + 1
        End If
    Loop

End Function
    
'------------------------------------------------
' 外観実績を振替先品番で比較SQL文作成
'------------------------------------------------

'概要      :指定された振替元品番と、以下の仕様値が一致している品番を抽出するSQL文を作成し、呼び出し元に返す。
'           （ワープランク）
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sBlockid        ,I  ,String       :ﾌﾞﾛｯｸID、又は、結晶番号
'          :sOld_Hinban     ,I  ,String       :振替元品番
'          :sMakeSql        ,O  ,String       :作成SQL文
'          :戻り値          ,O  ,Integer      :取得の成否(0:正常取得, -1:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/10 新規作成　SB

Public Function funGetKouhoHinban1_3(sProccd As String, sBlockId As String, sOld_Hinban As String, sMakesql As String) As Integer
    
    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql As String           'SQL全体
    Dim rs  As OraDynaset       'RecordSet
    Dim sinstr     As String    'ＳＱＬin句用文字列
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funGetKouhoHinban1_3 = 0
    
    '------------------------------------------ 振替元品種仕様データ取得 ------------------------------------------------------
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E027.HWFWARPR " & vbCrLf
    sql = sql & "FROM   TBCME027 E027 " & vbCrLf
    sql = sql & "WHERE  E027.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E027.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E027.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E027.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' " & vbCrLf
'    sql = sql & "WHERE  E027.HINBAN || TO_CHAR(E027.MNOREVNO, 'FM00') || E027.FACTORY || E027.OPECOND   =   '" & sOld_Hinban & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        funGetKouhoHinban1_3 = 13001
        GoTo db_Error
    End If
    
    '取得データセット
    Erase tbl_spec1_3
    With tbl_spec1_3(0)
        If IsNull(rs("HWFWARPR")) = False Then .HWFWARPR = rs("HWFWARPR") Else .HWFWARPR = " "      'Warpランク
    End With
    
    Set rs = Nothing
    On Error GoTo Apl_down
    '------------------------------------------ 指示取得 ------------------------------------------------------
    'Warpランク
    sResult = ""
    sinstr = ""
    If tbl_spec1_3(0).HWFWARPR <> " " Then
        RET = funCodeDBGet("SB", "WR", tbl_spec1_3(0).HWFWARPR, 0, " ", sResult)
    Else
        RET = funCodeDBGet("SB", "WR", "1", 0, " ", sResult)
    End If
    If RET <> 0 Then
        funGetKouhoHinban1_3 = 13002
        GoTo Apl_Exit
    End If
    RET = funinfo2(sResult, sinstr)
    If RET <> 0 Then
        funGetKouhoHinban1_3 = 13003
        GoTo Apl_Exit
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↓追加 2003/10/21
    RET = funCodeinGet("SB", "WR", sinstr, sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_3 = 13003
        GoTo Apl_Exit
    End If
    sinstr = sResult
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↑追加 2003/10/21
    '------------------------------------------ 振替元品種と同一仕様の品番を取得 ------------------------------------------------------
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT 'X' " & vbCrLf
'    sql = sql & "SELECT E027.HINBAN || TO_CHAR(E027.MNOREVNO, 'FM00') || E027.FACTORY || E027.OPECOND HINBAN " & vbCrLf
    sql = sql & "FROM   TBCME027 E027 " & vbCrLf
    sql = sql & "WHERE  E018A.HINBAN                    = E027.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018A.MNOREVNO, 'FM00') = TO_CHAR(E027.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E018A.FACTORY                   = E027.FACTORY                      AND " & vbCrLf
    sql = sql & "       E018A.OPECOND                   = E027.OPECOND                      AND " & vbCrLf
'    sql = sql & "WHERE  E027.HINBAN || TO_CHAR(E027.MNOREVNO, 'FM00') || E027.FACTORY || E027.OPECOND <> '" & sOld_Hinban & "' AND " & vbCrLf
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↓追加 2003/10/21
'    sql = sql & "       E027.HWFWARPR IN (SELECT NVL(TRIM(CODE),CHR(32)) FROM TBCMB005 WHERE SYSCLASS = 'SB' AND CLASS = 'WR' AND INFO2 IN (" & sinstr & ")) " & vbCrLf
    If tbl_spec1_3(0).HWFWARPR = " " Then
        sql = sql & "       E027.HWFWARPR is null OR E027.HWFWARPR IN (" & sinstr & ") " & vbCrLf
    Else
        sql = sql & "       E027.HWFWARPR IN (" & sinstr & ") " & vbCrLf
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↑追加 2003/10/21
    
    sMakesql = sql
    
'    On Error GoTo db_Error
'    'SQL文の実行
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'
'    '該当データなし
'    If rs.EOF Then
'        funGetKouhoHinban1_3 = 1
'        GoTo db_Error
'    Else
'        sMakesql = sql
'    End If
'
'    Set rs = Nothing
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funGetKouhoHinban1_3 = -4
    GoTo Apl_Exit
    
db_Error:
    Set rs = Nothing
    If funGetKouhoHinban1_3 = 0 Then
        funGetKouhoHinban1_3 = -3
    End If
    GoTo Apl_Exit

End Function

    
'------------------------------------------------
' 結晶評価項目仕様比較SQL文作成
'------------------------------------------------

'概要      :指定された振替元品番と、以下の仕様値が一致している品番を抽出するSQL文を作成し、呼び出し元に返す。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sBlockid        ,I  ,String       :ﾌﾞﾛｯｸID、又は、結晶番号
'          :sOld_Hinban     ,I  ,String       :振替元品番
'          :sMakeSql        ,O  ,String       :作成SQL文
'          :戻り値          ,O  ,Integer      :取得の成否(0:正常取得, -1:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/10 新規作成　SB

Public Function funGetKouhoHinban1_4(sProccd As String, sBlockId As String, sOld_Hinban As String, sMakesql As String) As Integer


    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql As String               'SQL全体
    Dim rs  As OraDynaset           'RecordSet
    Dim sMakesql1   As String       '呼び出しファンクションSQL作成
    Dim sMakesql2   As String       '呼び出しファンクションSQL作成
    Dim sMakesql3   As String       '呼び出しファンクションSQL作成
    Dim sMakesql4   As String       '呼び出しファンクションSQL作成
    Dim sMakesql5   As String       '呼び出しファンクションSQL作成
    Dim sMakesql6   As String       '呼び出しファンクションSQL作成
    Dim sMakesql7   As String       '呼び出しファンクションSQL作成
    Dim sMakesql8   As String       '呼び出しファンクションSQL作成
    Dim sMakesql9   As String       '呼び出しファンクションSQL作成
    Dim sMakesql10  As String       '呼び出しファンクションSQL作成
    Dim sMakesql11  As String       '呼び出しファンクションSQL作成
    Dim sMakesql12  As String       '呼び出しファンクションSQL作成
    Dim sMakesql13  As String       '呼び出しファンクションSQL作成
    Dim sMakesql14  As String       '呼び出しファンクションSQL作成
    Dim sMakesql15  As String       '呼び出しファンクションSQL作成
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funGetKouhoHinban1_4 = 0
    
    '------------------------------------------ 振替元品種仕様データ取得 ------------------------------------------------------
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E018.HSXRHWYS,E019.HSXONHWS,  E019.HSXONSPT,E019.HSXONSPI,E019.HSXONKWY,E020.HSXOF1HS,E020.HSXOF1SH,  E020.HSXOF1ST,E020.HSXOF1SR,  E020.HSXOF1NS,E020.HSXOF1SZ,   " & vbCrLf
    sql = sql & "       E020.HSXOF1ET,E020.HSXOSF1PTK,E020.HSXOF2HS,E020.HSXOF2SH,E020.HSXOF2ST,E020.HSXOF2SR,E020.HSXOF2NS,  E020.HSXOF2SZ,  E020.HSXOF2ET,E020.HSXOSF2PTK, " & vbCrLf
    sql = sql & "       E020.HSXOF3HS,E020.HSXOF3SH,  E020.HSXOF3ST,E020.HSXOF3SR,E020.HSXOF3NS,E020.HSXOF3SZ,  E020.HSXOF3ET,E020.HSXOSF3PTK,E020.HSXOF4HS,E020.HSXOF4SH,   " & vbCrLf
    sql = sql & "       E020.HSXOF4ST,E020.HSXOF4SR,  E020.HSXOF4NS,E020.HSXOF4SZ,E020.HSXOF4ET,E020.HSXOSF4PTK,E020.HSXBM1HS,E020.HSXBM1SH,  E020.HSXBM1ST,E020.HSXBM1SR,   " & vbCrLf
    sql = sql & "       E020.HSXBM1NS,E020.HSXBM1SZ,  E020.HSXBM1ET,E020.HSXBM2HS,E020.HSXBM2SH,E020.HSXBM2ST,  E020.HSXBM2SR,E020.HSXBM2NS,  E020.HSXBM2SZ,E020.HSXBM2ET,   " & vbCrLf
    sql = sql & "       E020.HSXBM3HS,E020.HSXBM3SH,  E020.HSXBM3ST,E020.HSXBM3SR,E020.HSXBM3NS,E020.HSXBM3SZ,  E020.HSXBM3ET,E019.HSXTMMAX,  E019.HSXLTHWS,E019.HSXCNHWS,   " & vbCrLf
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数取得追加
'    sql = sql & "       E019.HSXCNKWY,E020.HSXDENHS,  E020.HSXDENMN,E020.HSXDENMX,E020.HSXDVDHS,E020.HSXDVDMN,  E020.HSXDVDMX,E020.HSXLDLHS,  E020.HSXLDLMN,E020.HSXLDLMX    " & vbCrLf
'    sql = sql & "FROM   TBCME018 E018,TBCME019 E019,TBCME020 E020 " & vbCrLf
    sql = sql & "       E019.HSXCNKWY,E020.HSXDENHS,  E020.HSXDENMN,E020.HSXDENMX,E020.HSXDVDHS,E020.HSXDVDMN,  E020.HSXDVDMX,E020.HSXLDLHS,  E020.HSXLDLMN,E020.HSXLDLMX,E036.HSXGDLINE,E036.COSF3FLAG " & vbCrLf
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    sql = sql & "       ,NVL(E036.HSXDKTMP,' ') HSXDKTMP " & vbCrLf
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
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
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数取得追加
    sql = sql & "WHERE  E018.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E018.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' AND " & vbCrLf
    sql = sql & "       E019.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E019.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E019.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E019.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' AND " & vbCrLf
    sql = sql & "       E020.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E020.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E020.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数取得追加
'    sql = sql & "       E020.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' " & vbCrLf
    sql = sql & "       E020.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' AND " & vbCrLf
    sql = sql & "       E036.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E036.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E036.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    sql = sql & "   AND E048.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E048.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E048.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E048.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
'    sql = sql & "WHERE  E018.HINBAN || TO_CHAR(E018.MNOREVNO, 'FM00') || E018.FACTORY || E018.OPECOND   =   '" & sOld_Hinban & "' AND " & vbCrLf
'    sql = sql & "       E019.HINBAN || TO_CHAR(E019.MNOREVNO, 'FM00') || E019.FACTORY || E019.OPECOND   =   '" & sOld_Hinban & "' AND " & vbCrLf
'    sql = sql & "       E020.HINBAN || TO_CHAR(E020.MNOREVNO, 'FM00') || E020.FACTORY || E020.OPECOND   =   '" & sOld_Hinban & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        funGetKouhoHinban1_4 = 14001
        GoTo db_Error
    End If
    
    '取得データセット
    Erase tbl_spec1_4
    With tbl_spec1_4(0)
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
        If IsNull(rs("HSXOSF1PTK")) = False Then .HSXOSF1PTK = rs("HSXOSF1PTK") Else .HSXOSF1PTK = " "      'パターン区分
        'OSF2
        If IsNull(rs("HSXOF2HS")) = False Then .HSXOF2HS = rs("HSXOF2HS") Else .HSXOF2HS = " "              '保証方法_対象
        If IsNull(rs("HSXOF2SH")) = False Then .HSXOF2SH = rs("HSXOF2SH") Else .HSXOF2SH = " "              '測定位置_方
        If IsNull(rs("HSXOF2ST")) = False Then .HSXOF2ST = rs("HSXOF2ST") Else .HSXOF2ST = " "              '測定位置_点
        If IsNull(rs("HSXOF2SR")) = False Then .HSXOF2SR = rs("HSXOF2SR") Else .HSXOF2SR = " "              '測定位置_領
        If IsNull(rs("HSXOF2NS")) = False Then .HSXOF2NS = rs("HSXOF2NS") Else .HSXOF2NS = " "              '熱処理法
        If IsNull(rs("HSXOF2SZ")) = False Then .HSXOF2SZ = rs("HSXOF2SZ") Else .HSXOF2SZ = " "              '測定条件
        If IsNull(rs("HSXOF2ET")) = False Then .HSXOF2ET = rs("HSXOF2ET") Else .HSXOF2ET = 0                '選択ET代
        If IsNull(rs("HSXOSF2PTK")) = False Then .HSXOSF2PTK = rs("HSXOSF2PTK") Else .HSXOSF2PTK = " "      'パターン区分
        'OSF3
        If IsNull(rs("HSXOF3HS")) = False Then .HSXOF3HS = rs("HSXOF3HS") Else .HSXOF3HS = " "              '保証方法_対象
        If IsNull(rs("HSXOF3SH")) = False Then .HSXOF3SH = rs("HSXOF3SH") Else .HSXOF3SH = " "              '測定位置_方
        If IsNull(rs("HSXOF3ST")) = False Then .HSXOF3ST = rs("HSXOF3ST") Else .HSXOF3ST = " "              '測定位置_点
        If IsNull(rs("HSXOF3SR")) = False Then .HSXOF3SR = rs("HSXOF3SR") Else .HSXOF3SR = " "              '測定位置_領
        If IsNull(rs("HSXOF3NS")) = False Then .HSXOF3NS = rs("HSXOF3NS") Else .HSXOF3NS = " "              '熱処理法
        If IsNull(rs("HSXOF3SZ")) = False Then .HSXOF3SZ = rs("HSXOF3SZ") Else .HSXOF3SZ = " "              '測定条件
        If IsNull(rs("HSXOF3ET")) = False Then .HSXOF3ET = rs("HSXOF3ET") Else .HSXOF3ET = 0                '選択ET代
        If IsNull(rs("HSXOSF3PTK")) = False Then .HSXOSF3PTK = rs("HSXOSF3PTK") Else .HSXOSF3PTK = " "      'パターン区分
        'OSF4

''C－OSF3チェックの変更 2008.04.20 青柳
If sFlg_2_2 = "1" Then
'C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
        'If IsNull(rs("HSXOF4HS")) = False Then .HSXOF4HS = rs("HSXOF4HS") Else .HSXOF4HS = " "             '保証方法_対象
        If IsNull(rs("COSF3FLAG")) = False Then .HSXOF4HS = rs("COSF3FLAG") Else .HSXOF4HS = " "            'C-OSF3ﾌﾗｸﾞ
        If IsNull(rs("COSF3FLAG")) = False Then .COSF3FLAG = rs("COSF3FLAG") Else .COSF3FLAG = " "            'C-OSF3ﾌﾗｸﾞ
'C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---
End If

'--------------- 2008/08/25 INSERT START  By Systech ---------------
        .HSXDKTMP = rs("HSXDKTMP")
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

        If IsNull(rs("HSXOF4SH")) = False Then .HSXOF4SH = rs("HSXOF4SH") Else .HSXOF4SH = " "              '測定位置_方
        If IsNull(rs("HSXOF4ST")) = False Then .HSXOF4ST = rs("HSXOF4ST") Else .HSXOF4ST = " "              '測定位置_点
        If IsNull(rs("HSXOF4SR")) = False Then .HSXOF4SR = rs("HSXOF4SR") Else .HSXOF4SR = " "              '測定位置_領
        If IsNull(rs("HSXOF4NS")) = False Then .HSXOF4NS = rs("HSXOF4NS") Else .HSXOF4NS = " "              '熱処理法
        If IsNull(rs("HSXOF4SZ")) = False Then .HSXOF4SZ = rs("HSXOF4SZ") Else .HSXOF4SZ = " "              '測定条件
        If IsNull(rs("HSXOF4ET")) = False Then .HSXOF4ET = rs("HSXOF4ET") Else .HSXOF4ET = 0                '選択ET代
        If IsNull(rs("HSXOSF4PTK")) = False Then .HSXOSF4PTK = rs("HSXOSF4PTK") Else .HSXOSF4PTK = " "      'パターン区分

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
        'DEN
        If IsNull(rs("HSXDENHS")) = False Then .HSXDENHS = rs("HSXDENHS") Else .HSXDENHS = " "              '保証方法_対象
        If IsNull(rs("HSXDENMN")) = False Then .HSXDENMN = rs("HSXDENMN") Else .HSXDENMN = 0                '下限
        If IsNull(rs("HSXDENMX")) = False Then .HSXDENMX = rs("HSXDENMX") Else .HSXDENMX = 0                '上限
        'DVD2
        If IsNull(rs("HSXDVDHS")) = False Then .HSXDVDHS = rs("HSXDVDHS") Else .HSXDVDHS = " "              '保証方法_対象
        If IsNull(rs("HSXDVDMN")) = False Then .HSXDVDMN = rs("HSXDVDMN") Else .HSXDVDMN = 0                '下限
        If IsNull(rs("HSXDVDMX")) = False Then .HSXDVDMX = rs("HSXDVDMX") Else .HSXDVDMX = 0                '上限
        'L/DL
        If IsNull(rs("HSXLDLHS")) = False Then .HSXLDLHS = rs("HSXLDLHS") Else .HSXLDLHS = " "              '保証方法_対象
        If IsNull(rs("HSXLDLMN")) = False Then .HSXLDLMN = rs("HSXLDLMN") Else .HSXLDLMN = 0                '下限
        If IsNull(rs("HSXLDLMX")) = False Then .HSXLDLMX = rs("HSXLDLMX") Else .HSXLDLMX = 0                '上限
    '*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
        'GDﾗｲﾝ数
        If IsNull(rs("HSXGDLINE")) = False Then .HSXGDLINE = rs("HSXGDLINE") Else .HSXGDLINE = " "          'ﾗｲﾝ数
    '*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
    End With
    
    Set rs = Nothing
    On Error GoTo Apl_down
    '------------------------------------------ 指示取得 ------------------------------------------------------
    sMakesql1 = ""
    sMakesql2 = ""
    sMakesql3 = ""
    sMakesql4 = ""
    sMakesql5 = ""
    sMakesql6 = ""
    sMakesql7 = ""
    sMakesql8 = ""
    sMakesql9 = ""
    sMakesql10 = ""
    sMakesql11 = ""
    sMakesql12 = ""
    sMakesql13 = ""
    sMakesql14 = ""
    sMakesql15 = ""
    '比抵抗
    sResult = ""
    RET = funCodeDBGet("SB", "14", "RS", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14010
        GoTo Apl_Exit
    End If
    Erase tbl_spec1_4_1
    sMakesql = ""
    tbl_spec1_4_1(0).HOSYOU = tbl_spec1_4(0).HSXRHWYS
    tbl_spec1_4_1(0).HOSYOU1 = "HSXRHWYS"
'--------------- 2008/08/25 UPDATE START  By Systech ---------------
'    RET = funGetKouhoHinban1_4_1(sResult, tbl_spec1_4_1(), "E018", sMakesql)
    tbl_spec1_4_1(0).HSXDKTMP = tbl_spec1_4(0).HSXDKTMP
    tbl_spec1_4_1(0).HSXDKTMP1 = "HSXDKTMP"
    RET = funGetKouhoHinban1_4_1(sResult, tbl_spec1_4_1(), "E018", sMakesql, "E036")
'--------------- 2008/08/25 UPDATE  END   By Systech ---------------
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14010 + RET
        GoTo Apl_Exit
    End If
    sMakesql1 = sMakesql
    '酸素濃度
    sResult = ""
    RET = funCodeDBGet("SB", "14", "OI", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14020
        GoTo Apl_Exit
    End If
    Erase tbl_spec1_4_1
    sMakesql = ""
    tbl_spec1_4_1(0).HOSYOU = tbl_spec1_4(0).HSXONHWS
    tbl_spec1_4_1(0).HOSYOU1 = "HSXONHWS"
    tbl_spec1_4_1(0).SOKU_TEN = tbl_spec1_4(0).HSXONSPT     '08/01/29 ooba
    tbl_spec1_4_1(0).SOKU_TEN1 = "HSXONSPT"                 '08/01/29 ooba
    tbl_spec1_4_1(0).SOKU_ICHI = tbl_spec1_4(0).HSXONSPI
    tbl_spec1_4_1(0).SOKU_ICHI1 = "HSXONSPI"
    tbl_spec1_4_1(0).KENSA = tbl_spec1_4(0).HSXONKWY
    tbl_spec1_4_1(0).KENSA1 = "HSXONKWY"
    RET = funGetKouhoHinban1_4_1(sResult, tbl_spec1_4_1(), "E019", sMakesql)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14020 + RET
        GoTo Apl_Exit
    End If
    sMakesql2 = sMakesql
    'ＯＳＦ1
    sResult = ""
    RET = funCodeDBGet("SB", "14", "O1", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14030
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_4_1
    tbl_spec1_4_1(0).HOSYOU = tbl_spec1_4(0).HSXOF1HS
    tbl_spec1_4_1(0).HOSYOU1 = "HSXOF1HS"
    tbl_spec1_4_1(0).SOKU_HOU = tbl_spec1_4(0).HSXOF1SH
    tbl_spec1_4_1(0).SOKU_HOU1 = "HSXOF1SH"
    tbl_spec1_4_1(0).SOKU_TEN = tbl_spec1_4(0).HSXOF1ST
    tbl_spec1_4_1(0).SOKU_TEN1 = "HSXOF1ST"
    tbl_spec1_4_1(0).SOKU_RYOU = tbl_spec1_4(0).HSXOF1SR
    tbl_spec1_4_1(0).SOKU_RYOU1 = "HSXOF1SR"
    tbl_spec1_4_1(0).NETSU = tbl_spec1_4(0).HSXOF1NS
    tbl_spec1_4_1(0).NETSU1 = "HSXOF1NS"
    tbl_spec1_4_1(0).JOUKEN = tbl_spec1_4(0).HSXOF1SZ
    tbl_spec1_4_1(0).JOUKEN1 = "HSXOF1SZ"
    tbl_spec1_4_1(0).ET = tbl_spec1_4(0).HSXOF1ET
    tbl_spec1_4_1(0).ET1 = "HSXOF1ET"
    tbl_spec1_4_1(0).PATTERN = tbl_spec1_4(0).HSXOSF1PTK
    tbl_spec1_4_1(0).PATTERN1 = "HSXOSF1PTK"
    RET = funGetKouhoHinban1_4_1(sResult, tbl_spec1_4_1(), "E020", sMakesql)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14030 + RET
        GoTo Apl_Exit
    End If
    sMakesql3 = sMakesql
    'ＯＳＦ２
    sResult = ""
    RET = funCodeDBGet("SB", "14", "O2", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14040
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_4_1
    tbl_spec1_4_1(0).HOSYOU = tbl_spec1_4(0).HSXOF2HS
    tbl_spec1_4_1(0).HOSYOU1 = "HSXOF2HS"
    tbl_spec1_4_1(0).SOKU_HOU = tbl_spec1_4(0).HSXOF2SH
    tbl_spec1_4_1(0).SOKU_HOU1 = "HSXOF2SH"
    tbl_spec1_4_1(0).SOKU_TEN = tbl_spec1_4(0).HSXOF2ST
    tbl_spec1_4_1(0).SOKU_TEN1 = "HSXOF2ST"
    tbl_spec1_4_1(0).SOKU_RYOU = tbl_spec1_4(0).HSXOF2SR
    tbl_spec1_4_1(0).SOKU_RYOU1 = "HSXOF2SR"
    tbl_spec1_4_1(0).NETSU = tbl_spec1_4(0).HSXOF2NS
    tbl_spec1_4_1(0).NETSU1 = "HSXOF2NS"
    tbl_spec1_4_1(0).JOUKEN = tbl_spec1_4(0).HSXOF2SZ
    tbl_spec1_4_1(0).JOUKEN1 = "HSXOF2SZ"
    tbl_spec1_4_1(0).ET = tbl_spec1_4(0).HSXOF2ET
    tbl_spec1_4_1(0).ET1 = "HSXOF2ET"
    tbl_spec1_4_1(0).PATTERN = tbl_spec1_4(0).HSXOSF2PTK
    tbl_spec1_4_1(0).PATTERN1 = "HSXOSF2PTK"
    RET = funGetKouhoHinban1_4_1(sResult, tbl_spec1_4_1(), "E020", sMakesql)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14040 + RET
        GoTo Apl_Exit
    End If
    sMakesql4 = sMakesql
    'ＯＳＦ３
    sResult = ""
    RET = funCodeDBGet("SB", "14", "O3", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14050
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_4_1
    tbl_spec1_4_1(0).HOSYOU = tbl_spec1_4(0).HSXOF3HS
    tbl_spec1_4_1(0).HOSYOU1 = "HSXOF3HS"
    tbl_spec1_4_1(0).SOKU_HOU = tbl_spec1_4(0).HSXOF3SH
    tbl_spec1_4_1(0).SOKU_HOU1 = "HSXOF3SH"
    tbl_spec1_4_1(0).SOKU_TEN = tbl_spec1_4(0).HSXOF3ST
    tbl_spec1_4_1(0).SOKU_TEN1 = "HSXOF3ST"
    tbl_spec1_4_1(0).SOKU_RYOU = tbl_spec1_4(0).HSXOF3SR
    tbl_spec1_4_1(0).SOKU_RYOU1 = "HSXOF3SR"
    tbl_spec1_4_1(0).NETSU = tbl_spec1_4(0).HSXOF3NS
    tbl_spec1_4_1(0).NETSU1 = "HSXOF3NS"
    tbl_spec1_4_1(0).JOUKEN = tbl_spec1_4(0).HSXOF3SZ
    tbl_spec1_4_1(0).JOUKEN1 = "HSXOF3SZ"
    tbl_spec1_4_1(0).ET = tbl_spec1_4(0).HSXOF3ET
    tbl_spec1_4_1(0).ET1 = "HSXOF3ET"
    tbl_spec1_4_1(0).PATTERN = tbl_spec1_4(0).HSXOSF3PTK
    tbl_spec1_4_1(0).PATTERN1 = "HSXOSF3PTK"
    RET = funGetKouhoHinban1_4_1(sResult, tbl_spec1_4_1(), "E020", sMakesql)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14050 + RET
        GoTo Apl_Exit
    End If
    sMakesql5 = sMakesql
    
    'ＯＳＦ４
    sResult = ""
    RET = funCodeDBGet("SB", "14", "O4", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14060
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_4_1
    tbl_spec1_4_1(0).HOSYOU = tbl_spec1_4(0).COSF3FLAG
    'tbl_spec1_4_1(0).HOSYOU = tbl_spec1_4(0).HSXOF4HS
    'tbl_spec1_4_1(0).HOSYOU1 = "HSXOF4HS"
    tbl_spec1_4_1(0).HOSYOU1 = "COSF3FLAG"
    tbl_spec1_4_1(0).SOKU_HOU = tbl_spec1_4(0).HSXOF4SH
    tbl_spec1_4_1(0).SOKU_HOU1 = "HSXOF4SH"
    tbl_spec1_4_1(0).SOKU_TEN = tbl_spec1_4(0).HSXOF4ST
    tbl_spec1_4_1(0).SOKU_TEN1 = "HSXOF4ST"
    tbl_spec1_4_1(0).SOKU_RYOU = tbl_spec1_4(0).HSXOF4SR
    tbl_spec1_4_1(0).SOKU_RYOU1 = "HSXOF4SR"
    tbl_spec1_4_1(0).NETSU = tbl_spec1_4(0).HSXOF4NS
    tbl_spec1_4_1(0).NETSU1 = "HSXOF4NS"
    tbl_spec1_4_1(0).JOUKEN = tbl_spec1_4(0).HSXOF4SZ
    tbl_spec1_4_1(0).JOUKEN1 = "HSXOF4SZ"
    tbl_spec1_4_1(0).ET = tbl_spec1_4(0).HSXOF4ET
    tbl_spec1_4_1(0).ET1 = "HSXOF4ET"
    tbl_spec1_4_1(0).PATTERN = tbl_spec1_4(0).HSXOSF4PTK
    tbl_spec1_4_1(0).PATTERN1 = "HSXOSF4PTK"
    RET = funGetKouhoHinban1_4_1(sResult, tbl_spec1_4_1(), "E020", sMakesql, "E036")
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14060 + RET
        GoTo Apl_Exit
    End If
    sMakesql6 = sMakesql
    
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    'ＳＩＲＤ
    sResult = ""
    RET = funCodeDBGet("SB", "14", "SD", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14160
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_4_1                                         '引数ﾃｰﾌﾞﾙｸﾘｱ
    tbl_spec1_4_1(0).HOSYOU = tbl_spec1_4(0).HWFSIRDHS          '軸状転位保証方法＿処
    tbl_spec1_4_1(0).HOSYOU1 = "HWFSIRDHS"                      '軸状転位保証方法＿処
    tbl_spec1_4_1(0).JOUKEN = tbl_spec1_4(0).HWFSIRDSZ          '軸状転位測定条件
    tbl_spec1_4_1(0).JOUKEN1 = "HWFSIRDSZ"                      '軸状転位測定条件
    tbl_spec1_4_1(0).HWFSIRDMX = tbl_spec1_4(0).HWFSIRDMX       '軸状転位上限
    tbl_spec1_4_1(0).HWFSIRDMX1 = "HWFSIRDMX"                   '軸状転位上限
    tbl_spec1_4_1(0).HWFSIRDHT = tbl_spec1_4(0).HWFSIRDHT       '軸状転位保証方法＿対
    tbl_spec1_4_1(0).HWFSIRDHT1 = "HWFSIRDHT"                   '軸状転位保証方法＿対
    tbl_spec1_4_1(0).HWFSIRDKM = tbl_spec1_4(0).HWFSIRDKM       '軸状転位検査頻度＿枚
    tbl_spec1_4_1(0).HWFSIRDKM1 = "HWFSIRDKM"                   '軸状転位検査頻度＿枚
    tbl_spec1_4_1(0).HWFSIRDKH = tbl_spec1_4(0).HWFSIRDKH       '軸状転位検査頻度＿保
    tbl_spec1_4_1(0).HWFSIRDKH1 = "HWFSIRDKH"                   '軸状転位検査頻度＿保
    tbl_spec1_4_1(0).HWFSIRDKU = tbl_spec1_4(0).HWFSIRDKU       '軸状転位検査頻度＿ウ
    tbl_spec1_4_1(0).HWFSIRDKU1 = "HWFSIRDKU"                   '軸状転位検査頻度＿ウ
    tbl_spec1_4_1(0).HWFSIRDKN = tbl_spec1_4(0).HWFSIRDKN       '軸状転位検査頻度＿抜
    tbl_spec1_4_1(0).HWFSIRDKN1 = "HWFSIRDKN"                   '軸状転位検査頻度＿抜
    RET = funGetKouhoHinban1_4_1(sResult, tbl_spec1_4_1(), "E048", sMakesql)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14160 + RET
        GoTo Apl_Exit
    End If
    sMakesql6 = sMakesql
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)

    'ＢＭＤ１
    sResult = ""
    RET = funCodeDBGet("SB", "14", "B1", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14070
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_4_1
    tbl_spec1_4_1(0).HOSYOU = tbl_spec1_4(0).HSXBM1HS
    tbl_spec1_4_1(0).HOSYOU1 = "HSXBM1HS"
    tbl_spec1_4_1(0).SOKU_HOU = tbl_spec1_4(0).HSXBM1SH
    tbl_spec1_4_1(0).SOKU_HOU1 = "HSXBM1SH"
    tbl_spec1_4_1(0).SOKU_TEN = tbl_spec1_4(0).HSXBM1ST
    tbl_spec1_4_1(0).SOKU_TEN1 = "HSXBM1ST"
    tbl_spec1_4_1(0).SOKU_RYOU = tbl_spec1_4(0).HSXBM1SR
    tbl_spec1_4_1(0).SOKU_RYOU1 = "HSXBM1SR"
    tbl_spec1_4_1(0).NETSU = tbl_spec1_4(0).HSXBM1NS
    tbl_spec1_4_1(0).NETSU1 = "HSXBM1NS"
    tbl_spec1_4_1(0).JOUKEN = tbl_spec1_4(0).HSXBM1SZ
    tbl_spec1_4_1(0).JOUKEN1 = "HSXBM1SZ"
    tbl_spec1_4_1(0).ET = tbl_spec1_4(0).HSXBM1ET
    tbl_spec1_4_1(0).ET1 = "HSXBM1ET"
    RET = funGetKouhoHinban1_4_1(sResult, tbl_spec1_4_1(), "E020", sMakesql)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14070 + RET
        GoTo Apl_Exit
    End If
    sMakesql7 = sMakesql
    'ＢＭＤ２
    sResult = ""
    RET = funCodeDBGet("SB", "14", "B2", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14080
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_4_1
    tbl_spec1_4_1(0).HOSYOU = tbl_spec1_4(0).HSXBM2HS
    tbl_spec1_4_1(0).HOSYOU1 = "HSXBM2HS"
    tbl_spec1_4_1(0).SOKU_HOU = tbl_spec1_4(0).HSXBM2SH
    tbl_spec1_4_1(0).SOKU_HOU1 = "HSXBM2SH"
    tbl_spec1_4_1(0).SOKU_TEN = tbl_spec1_4(0).HSXBM2ST
    tbl_spec1_4_1(0).SOKU_TEN1 = "HSXBM2ST"
    tbl_spec1_4_1(0).SOKU_RYOU = tbl_spec1_4(0).HSXBM2SR
    tbl_spec1_4_1(0).SOKU_RYOU1 = "HSXBM2SR"
    tbl_spec1_4_1(0).NETSU = tbl_spec1_4(0).HSXBM2NS
    tbl_spec1_4_1(0).NETSU1 = "HSXBM2NS"
    tbl_spec1_4_1(0).JOUKEN = tbl_spec1_4(0).HSXBM2SZ
    tbl_spec1_4_1(0).JOUKEN1 = "HSXBM2SZ"
    tbl_spec1_4_1(0).ET = tbl_spec1_4(0).HSXBM2ET
    tbl_spec1_4_1(0).ET1 = "HSXBM2ET"
    RET = funGetKouhoHinban1_4_1(sResult, tbl_spec1_4_1(), "E020", sMakesql)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14080 + RET
        GoTo Apl_Exit
    End If
    sMakesql8 = sMakesql
    'ＢＭＤ３
    sResult = ""
    RET = funCodeDBGet("SB", "14", "B3", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14090
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_4_1
    tbl_spec1_4_1(0).HOSYOU = tbl_spec1_4(0).HSXBM3HS
    tbl_spec1_4_1(0).HOSYOU1 = "HSXBM3HS"
    tbl_spec1_4_1(0).SOKU_HOU = tbl_spec1_4(0).HSXBM3SH
    tbl_spec1_4_1(0).SOKU_HOU1 = "HSXBM3SH"
    tbl_spec1_4_1(0).SOKU_TEN = tbl_spec1_4(0).HSXBM3ST
    tbl_spec1_4_1(0).SOKU_TEN1 = "HSXBM3ST"
    tbl_spec1_4_1(0).SOKU_RYOU = tbl_spec1_4(0).HSXBM3SR
    tbl_spec1_4_1(0).SOKU_RYOU1 = "HSXBM3SR"
    tbl_spec1_4_1(0).NETSU = tbl_spec1_4(0).HSXBM3NS
    tbl_spec1_4_1(0).NETSU1 = "HSXBM3NS"
    tbl_spec1_4_1(0).JOUKEN = tbl_spec1_4(0).HSXBM3SZ
    tbl_spec1_4_1(0).JOUKEN1 = "HSXBM3SZ"
    tbl_spec1_4_1(0).ET = tbl_spec1_4(0).HSXBM3ET
    tbl_spec1_4_1(0).ET1 = "HSXBM3ET"
    RET = funGetKouhoHinban1_4_1(sResult, tbl_spec1_4_1(), "E020", sMakesql)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14090 + RET
        GoTo Apl_Exit
    End If
    sMakesql9 = sMakesql
    'ＥＰＤ
    sResult = ""
    RET = funCodeDBGet("SB", "14", "EPD", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14100
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_4_1
    tbl_spec1_4_1(0).max = tbl_spec1_4(0).HSXTMMAX
    tbl_spec1_4_1(0).Max1 = "HSXTMMAX"
    RET = funGetKouhoHinban1_4_1(sResult, tbl_spec1_4_1(), "E019", sMakesql)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14100 + RET
        GoTo Apl_Exit
    End If
    sMakesql10 = sMakesql
    'ライフタイム
    sResult = ""
    RET = funCodeDBGet("SB", "14", "LT", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14110
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_4_1
    tbl_spec1_4_1(0).HOSYOU = tbl_spec1_4(0).HSXLTHWS
    tbl_spec1_4_1(0).HOSYOU1 = "HSXLTHWS"
    RET = funGetKouhoHinban1_4_1(sResult, tbl_spec1_4_1(), "E019", sMakesql)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14110 + RET
        GoTo Apl_Exit
    End If
    sMakesql11 = sMakesql
    '炭素濃度
    sResult = ""
    RET = funCodeDBGet("SB", "14", "CS", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14120
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_4_1
    tbl_spec1_4_1(0).HOSYOU = tbl_spec1_4(0).HSXCNHWS
    tbl_spec1_4_1(0).HOSYOU1 = "HSXCNHWS"
    tbl_spec1_4_1(0).KENSA = tbl_spec1_4(0).HSXCNKWY
    tbl_spec1_4_1(0).KENSA1 = "HSXCNKWY"
    RET = funGetKouhoHinban1_4_1(sResult, tbl_spec1_4_1(), "E019", sMakesql)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14120 + RET
        GoTo Apl_Exit
    End If
    sMakesql12 = sMakesql
    'ＤＥＮ
    sResult = ""
    RET = funCodeDBGet("SB", "14", "DEN", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14130
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_4_1
    tbl_spec1_4_1(0).HOSYOU = tbl_spec1_4(0).HSXDENHS
    tbl_spec1_4_1(0).HOSYOU1 = "HSXDENHS"
    tbl_spec1_4_1(0).Min = tbl_spec1_4(0).HSXDENMN
    tbl_spec1_4_1(0).Min1 = "HSXDENMN"
    tbl_spec1_4_1(0).max = tbl_spec1_4(0).HSXDENMX
    tbl_spec1_4_1(0).Max1 = "HSXDENMX"
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
    tbl_spec1_4_1(0).LINE = tbl_spec1_4(0).HSXGDLINE
    tbl_spec1_4_1(0).Line1 = "HSXGDLINE"
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
    RET = funGetKouhoHinban1_4_1(sResult, tbl_spec1_4_1(), "E020", sMakesql, "E036")
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14130 + RET
        GoTo Apl_Exit
    End If
    sMakesql13 = sMakesql
    'ＤＶＤ２
    sResult = ""
    RET = funCodeDBGet("SB", "14", "DVD", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14140
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_4_1
    tbl_spec1_4_1(0).HOSYOU = tbl_spec1_4(0).HSXDVDHS
    tbl_spec1_4_1(0).HOSYOU1 = "HSXDVDHS"
    tbl_spec1_4_1(0).Min = tbl_spec1_4(0).HSXDVDMN
    tbl_spec1_4_1(0).Min1 = "HSXDVDMN"
    tbl_spec1_4_1(0).max = tbl_spec1_4(0).HSXDVDMX
    tbl_spec1_4_1(0).Max1 = "HSXDVDMX"
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
    tbl_spec1_4_1(0).LINE = tbl_spec1_4(0).HSXGDLINE
    tbl_spec1_4_1(0).Line1 = "HSXGDLINE"
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
    RET = funGetKouhoHinban1_4_1(sResult, tbl_spec1_4_1(), "E020", sMakesql, "E036")
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14140 + RET
        GoTo Apl_Exit
    End If
    sMakesql14 = sMakesql
    'Ｌ／ＤＬ
    sResult = ""
    RET = funCodeDBGet("SB", "14", "LDL", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14150
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_4_1
    tbl_spec1_4_1(0).HOSYOU = tbl_spec1_4(0).HSXLDLHS
    tbl_spec1_4_1(0).HOSYOU1 = "HSXLDLHS"
    tbl_spec1_4_1(0).Min = tbl_spec1_4(0).HSXLDLMN
    tbl_spec1_4_1(0).Min1 = "HSXLDLMN"
    tbl_spec1_4_1(0).max = tbl_spec1_4(0).HSXLDLMX
    tbl_spec1_4_1(0).Max1 = "HSXLDLMX"
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
    tbl_spec1_4_1(0).LINE = tbl_spec1_4(0).HSXGDLINE
    tbl_spec1_4_1(0).Line1 = "HSXGDLINE"
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
    RET = funGetKouhoHinban1_4_1(sResult, tbl_spec1_4_1(), "E020", sMakesql, "E036")
    If RET <> 0 Then
        funGetKouhoHinban1_4 = 14150 + RET
        GoTo Apl_Exit
    End If
    sMakesql15 = sMakesql
    '------------------------------------------ 振替元品種と同一仕様の品番を取得 ------------------------------------------------------
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT 'X' " & vbCrLf
'    sql = sql & "SELECT E018.HINBAN || TO_CHAR(E018.MNOREVNO, 'FM00') || E018.FACTORY || E018.OPECOND HINBAN " & vbCrLf
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
'    sql = sql & "FROM   TBCME018 E018, TBCME019 E019, TBCME020 E020 " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''    sql = sql & "FROM   TBCME018 E018, TBCME019 E019, TBCME020 E020, TBCME036 E036 " & vbCrLf
    sql = sql & "FROM   TBCME018 E018, TBCME019 E019, TBCME020 E020, TBCME036 E036 , TBCME048 E048 " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
    sql = sql & "WHERE  E018A.HINBAN                    = E018.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018A.MNOREVNO, 'FM00') = TO_CHAR(E018.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E018A.FACTORY                   = E018.FACTORY                      AND " & vbCrLf
    sql = sql & "       E018A.OPECOND                   = E018.OPECOND                      AND " & vbCrLf
    sql = sql & "       E018.HINBAN                     = E019.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018.MNOREVNO, 'FM00')  = TO_CHAR(E019.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E018.FACTORY                    = E019.FACTORY                      AND " & vbCrLf
    sql = sql & "       E018.OPECOND                    = E019.OPECOND                      AND " & vbCrLf
    sql = sql & "       E018.HINBAN                     = E020.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018.MNOREVNO, 'FM00')  = TO_CHAR(E020.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E018.FACTORY                    = E020.FACTORY                      AND " & vbCrLf
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
'    sql = sql & "       E018.OPECOND                    = E020.OPECOND " & vbCrLf
    sql = sql & "       E018.OPECOND                    = E020.OPECOND                      AND " & vbCrLf
    sql = sql & "       E018.HINBAN                     = E036.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018.MNOREVNO, 'FM00')  = TO_CHAR(E036.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E018.FACTORY                    = E036.FACTORY                      AND " & vbCrLf
    sql = sql & "       E018.OPECOND                    = E036.OPECOND                       " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    sql = sql & "   AND E018.HINBAN                     = E048.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018.MNOREVNO, 'FM00')  = TO_CHAR(E048.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E018.FACTORY                    = E048.FACTORY                      AND " & vbCrLf
    sql = sql & "       E018.OPECOND                    = E048.OPECOND                       " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
'    sql = sql & "WHERE  E018.HINBAN || TO_CHAR(E018.MNOREVNO, 'FM00') || E018.FACTORY || E018.OPECOND <> '" & sOld_Hinban & "' AND " & vbCrLf
'    sql = sql & "       E018.HINBAN || TO_CHAR(E018.MNOREVNO, 'FM00') || E018.FACTORY || E018.OPECOND = E019.HINBAN || TO_CHAR(E019.MNOREVNO, 'FM00') || E019.FACTORY || E019.OPECOND AND " & vbCrLf
'    sql = sql & "       E018.HINBAN || TO_CHAR(E018.MNOREVNO, 'FM00') || E018.FACTORY || E018.OPECOND = E020.HINBAN || TO_CHAR(E020.MNOREVNO, 'FM00') || E020.FACTORY || E020.OPECOND " & vbCrLf
    sql = sql & sMakesql1
    sql = sql & sMakesql2
    sql = sql & sMakesql3
    sql = sql & sMakesql4
    sql = sql & sMakesql5
    sql = sql & sMakesql6
    sql = sql & sMakesql7
    sql = sql & sMakesql8
    sql = sql & sMakesql9
    sql = sql & sMakesql10
    sql = sql & sMakesql11
    sql = sql & sMakesql12
    sql = sql & sMakesql13
    sql = sql & sMakesql14
    sql = sql & sMakesql15
    
    sMakesql = sql
    
'    On Error GoTo db_Error
'    'SQL文の実行
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'
'    '該当データなし
'    If rs.EOF Then
'        funGetKouhoHinban1_4 = 1
'        GoTo db_Error
'    Else
'        sMakesql = sql
'    End If
'
'    Set rs = Nothing
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funGetKouhoHinban1_4 = -4
    GoTo Apl_Exit
    
db_Error:
    Set rs = Nothing
    If funGetKouhoHinban1_4 = 0 Then
        funGetKouhoHinban1_4 = -3
    End If
    GoTo Apl_Exit

End Function

'------------------------------------------------
' 結晶評価項目仕様比較詳細SQL文作成
'------------------------------------------------

'概要      :指定されたﾁｪｯｸ内容詳細に基づき、該当する仕様値が一致している、または、マトリクスで一致している品番を抽出するSQL文を作成する。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型                 :説明
'          :sChkCode        ,I  ,String             :工程番号
'          :tbl_spec1_4_1   ,I  ,typ_ChkFurikae1-4  :ﾌﾞﾛｯｸID、又は、結晶番号
'          :sChkTable       ,I  ,String             :振替元品番
'          :sMakeSql        ,O  ,String             :作成SQL文
'          :sChkTable2      ,I  ,String             :
'          :戻り値          ,O  ,Integer            :取得の成否(0:正常取得, -1:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/10 新規作成　SB
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾁｪｯｸするﾃｰﾌﾞﾙが2つある場合に対応する為、引数にsChkTable2を追加
'Public Function funGetKouhoHinban1_4_1(sChkCode As String, tbl_spec1_4_1() As typ_Spec1_4_1, sChkTable As String, sMakesql As String) As Integer
Public Function funGetKouhoHinban1_4_1(sChkCode As String, tbl_spec1_4_1() As typ_Spec1_4_1, sChkTable As String, sMakesql As String, Optional sChkTable2 As String = "") As Integer
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾁｪｯｸするﾃｰﾌﾞﾙが2つある場合に対応する為、引数にsChkTable2を追加
    Dim RET         As Integer      '戻り値
    Dim sql As String       'SQL全体
    Dim rs  As OraDynaset   'RecordSet
    Dim sinstr     As String       'ＳＱＬin句用文字列
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    Dim lsDKCodeListWork()  As String   'Code一覧
    Dim lsDKCodeList()  As String      'Code一覧
    Dim iCnt            As Integer
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

    On Error GoTo Apl_down
    
    '戻り値初期化
    funGetKouhoHinban1_4_1 = 0
    '------------------------------------------ SQL文生成 ------------------------------------------------------
    'SQL文の作成
    sql = vbNullString
    '保証方法＿対象
    If Mid(sChkCode, 1, 1) = "2" Then
'        If tbl_spec1_4_1(0).HOSYOU = "H" Or tbl_spec1_4_1(0).HOSYOU = "S" Then
'            'マトリクス取得
'            sResult = ""
'            sinstr = ""
'            ret = funCodeDBGet("SB", "SH", tbl_spec1_4_1(0).HOSYOU, 0, " ", sResult)
'            If ret <> 0 Then
'                funGetKouhoHinban1_4_1 = 1
'                GoTo Apl_Exit
'            End If
'            ret = funinfo2(sResult, sinstr)
'            If ret <> 0 Then
'                funGetKouhoHinban1_4_1 = 1
'                GoTo Apl_Exit
'            End If
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↓追加 2003/10/21
'            ret = funCodeinGet("SB", "SH", sinstr, sResult)
'            If ret <> 0 Then
'                funGetKouhoHinban1_4_1 = 1
'                GoTo Apl_Exit
'            End If
'            sinstr = sResult
'    '        sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).HOSYOU1 & " IN  (SELECT NVL(TRIM(CODE),CHR(32)) FROM TBCMB005 WHERE SYSCLASS = 'SB' AND CLASS = 'SH' AND INFO2 in (" & sinstr & ")) " & vbCrLf
'            sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).HOSYOU1 & " IN  (" & sinstr & ") " & vbCrLf
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↑追加 2003/10/21
'        Else
'            sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).HOSYOU1 & " NOT IN  ('H', 'S') " & vbCrLf
'        End If

''C－OSF3チェックの変更 2008.04.20 青柳
If sFlg_2_2 = "1" Then
'C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
        If Trim(tbl_spec1_4_1(0).HOSYOU1) = "COSF3FLAG" Then
            If tbl_spec1_4_1(0).HOSYOU = "S" Then
                sql = sql & " AND " & sChkTable2 & "." & tbl_spec1_4_1(0).HOSYOU1 & " NOT IN  ('H') " & vbCrLf
            ElseIf tbl_spec1_4_1(0).HOSYOU <> "H" And tbl_spec1_4_1(0).HOSYOU <> "S" Then
                sql = sql & " AND (" & sChkTable2 & "." & tbl_spec1_4_1(0).HOSYOU1 & " NOT IN  ('H', 'S') " & vbCrLf
                sql = sql & " OR " & sChkTable2 & "." & tbl_spec1_4_1(0).HOSYOU1 & " IS NULL)" & vbCrLf
            End If
        Else
            If tbl_spec1_4_1(0).HOSYOU = "S" Then
                sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).HOSYOU1 & " NOT IN  ('H') " & vbCrLf
            ElseIf tbl_spec1_4_1(0).HOSYOU <> "H" And tbl_spec1_4_1(0).HOSYOU <> "S" Then
                sql = sql & " AND (" & sChkTable & "." & tbl_spec1_4_1(0).HOSYOU1 & " NOT IN  ('H', 'S') " & vbCrLf
                sql = sql & " OR " & sChkTable & "." & tbl_spec1_4_1(0).HOSYOU1 & " IS NULL)" & vbCrLf
            End If
        End If
End If

    End If
'C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---
    '------------------------------------------ 保証方法チェック ------------------------------------------------------
    If tbl_spec1_4_1(0).HOSYOU <> "H" And tbl_spec1_4_1(0).HOSYOU <> "S" Then GoTo Make_Exit
    
    '下限
    If Mid(sChkCode, 2, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).Min1 & " = " & tbl_spec1_4_1(0).Min & " " & vbCrLf
    End If
    '上限
    If Mid(sChkCode, 3, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).Max1 & " = " & tbl_spec1_4_1(0).max & " " & vbCrLf
    End If
    '測定位置＿方
    If Mid(sChkCode, 4, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).SOKU_HOU1 & " = '" & tbl_spec1_4_1(0).SOKU_HOU & "' " & vbCrLf
    End If
    '測定位置＿点
    If Mid(sChkCode, 5, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).SOKU_TEN1 & " = '" & tbl_spec1_4_1(0).SOKU_TEN & "' " & vbCrLf
    ElseIf Mid(sChkCode, 5, 1) = "2" Then   '08/01/29 ooba
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).SOKU_TEN1 & " <= '" & tbl_spec1_4_1(0).SOKU_TEN & "' " & vbCrLf
    End If
    '測定位置＿位
    If Mid(sChkCode, 6, 1) = "2" Then
        'マトリクス取得
        sResult = ""
        sinstr = ""
        RET = funCodeDBGet("SB", "OI", tbl_spec1_4_1(0).SOKU_ICHI, 0, " ", sResult)
        If RET <> 0 Then
            funGetKouhoHinban1_4_1 = 2
            GoTo Apl_Exit
        End If
        RET = funinfo2(sResult, sinstr)
        If RET <> 0 Then
            funGetKouhoHinban1_4_1 = 2
            GoTo Apl_Exit
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↓追加 2003/10/21
        RET = funCodeinGet("SB", "OI", sinstr, sResult)
        If RET <> 0 Then
            funGetKouhoHinban1_4_1 = 2
            GoTo Apl_Exit
        End If
        sinstr = sResult
'        sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).SOKU_ICHI1 & " IN  (SELECT NVL(TRIM(CODE),CHR(32)) FROM TBCMB005 WHERE SYSCLASS = 'SB' AND CLASS = 'OI' AND INFO2 in (" & sinstr & ")) " & vbCrLf
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).SOKU_ICHI1 & " IN  (" & sinstr & ") " & vbCrLf
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↑追加 2003/10/21
    End If
    '測定位置＿領
    If Mid(sChkCode, 7, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).SOKU_RYOU1 & " = '" & tbl_spec1_4_1(0).SOKU_RYOU & "' " & vbCrLf
    End If
    '検査有無
    If Mid(sChkCode, 8, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).UMU1 & " = '" & tbl_spec1_4_1(0).UMU & "' " & vbCrLf
    End If
    '熱処理法
    If Mid(sChkCode, 9, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).NETSU1 & " = '" & tbl_spec1_4_1(0).NETSU & "' " & vbCrLf
    End If
    '測定条件
    If Mid(sChkCode, 10, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).JOUKEN1 & " = '" & tbl_spec1_4_1(0).JOUKEN & "' " & vbCrLf
    End If
    '選択ＥＴ代
    If Mid(sChkCode, 11, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).ET1 & " = " & tbl_spec1_4_1(0).ET & " " & vbCrLf
    End If
    '検査方法
    If Mid(sChkCode, 12, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).KENSA1 & " = '" & tbl_spec1_4_1(0).KENSA & "' " & vbCrLf
    End If
    'ライン数
    If Mid(sChkCode, 13, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).Line1 & " = '" & tbl_spec1_4_1(0).LINE & "' " & vbCrLf
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
    ElseIf Mid(sChkCode, 13, 1) = "2" Then
        'マトリクス取得
        sResult = ""
        sinstr = ""
        
        RET = funCodeDBGet("SB", "LN", tbl_spec1_4_1(0).LINE, 0, " ", sResult)
        
        If RET <> 0 Then
            funGetKouhoHinban1_4_1 = 4
            GoTo Apl_Exit
        End If
        RET = funinfo2(sResult, sinstr)
        If RET <> 0 Then
            funGetKouhoHinban1_4_1 = 4
            GoTo Apl_Exit
        End If
                
        RET = funCodeinGet("SB", "LN", sinstr, sResult)
        If RET <> 0 Then
            funGetKouhoHinban1_4_1 = 4
            GoTo Apl_Exit
        End If
        sinstr = sResult
                
        If InStr(sinstr, "' '") > 0 Then
            'DBのﾗｲﾝ数ｶﾗﾑは整数型で，ｽﾍﾟｰｽを条件に入れるとｴﾗｰになるのでその対応
            If InStr(sinstr, ",' '") > 0 Then
                sinstr = Replace(sinstr, ",' '", "")
            ElseIf InStr(sinstr, "' ',") > 0 Then
                sinstr = Replace(sinstr, "' ',", "")
            Else
                sinstr = Replace(sinstr, "' '", "")
            End If
            sql = sql & " AND (" & sChkTable2 & "." & tbl_spec1_4_1(0).Line1 & " IS NULL" & vbCrLf
            sql = sql & " OR   " & sChkTable2 & "." & tbl_spec1_4_1(0).Line1 & " IN  (" & sinstr & "))" & vbCrLf
        Else
            sql = sql & " AND (" & sChkTable2 & "." & tbl_spec1_4_1(0).Line1 & " IN  (" & sinstr & ") " & vbCrLf
        End If
        
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
    End If
    'パターン区分
    If Mid(sChkCode, 14, 1) = "2" Then
        'マトリクス取得
        sResult = ""
        sinstr = ""
        If tbl_spec1_4_1(0).PATTERN <> " " Then
            RET = funCodeDBGet("SB", "OS", tbl_spec1_4_1(0).PATTERN, 0, " ", sResult)
        Else
            RET = funCodeDBGet("SB", "OS", "4", 0, " ", sResult)
        End If
        If RET <> 0 Then
            funGetKouhoHinban1_4_1 = 3
            GoTo Apl_Exit
        End If
        RET = funinfo2(sResult, sinstr)
        If RET <> 0 Then
            funGetKouhoHinban1_4_1 = 3
            GoTo Apl_Exit
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↓追加 2003/10/21
        RET = funCodeinGet("SB", "OS", sinstr, sResult)
        If RET <> 0 Then
            funGetKouhoHinban1_4_1 = 3
            GoTo Apl_Exit
        End If
        sinstr = sResult
'        sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).PATTERN1 & " IN  (SELECT NVL(TRIM(CODE),CHR(32)) FROM TBCMB005 WHERE SYSCLASS = 'SB' AND CLASS = 'OS' AND INFO2 in (" & sinstr & ")) " & vbCrLf
        If tbl_spec1_4_1(0).PATTERN = " " Then
            sql = sql & " AND (" & sChkTable & "." & tbl_spec1_4_1(0).PATTERN1 & " IN  (" & sinstr & ") " & vbCrLf
            sql = sql & " OR " & sChkTable & "." & tbl_spec1_4_1(0).PATTERN1 & " IS NULL)" & vbCrLf
        Else
            sql = sql & " AND " & sChkTable & "." & tbl_spec1_4_1(0).PATTERN1 & " IN  (" & sinstr & ") " & vbCrLf
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↑追加 2003/10/21
    End If
    
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    'DK温度
    If Mid(sChkCode, 15, 1) = "2" Then
        ReDim lsDKCodeListWork(0) As String
        ReDim lsDKCodeList(0) As String
        
        If Trim(tbl_spec1_4_1(0).HSXDKTMP) <> "" Then
            'DK温度マトリックスよりCodeの一覧を取得する
            RET = funCodeDBGetCodeList(DKTMP_TBCMB005SYS, DKTMP_TBCMB005CLS, lsDKCodeListWork)
            If RET < 0 Then
                funGetKouhoHinban1_4_1 = 4
                GoTo Apl_Exit
            End If
            
            For iCnt = 1 To UBound(lsDKCodeListWork)
                RET = funCodeDBGetMatrixReturn(DKTMP_TBCMB005SYS, DKTMP_TBCMB005CLS, tbl_spec1_4_1(0).HSXDKTMP, lsDKCodeListWork(iCnt))
                If RET < 0 Then
                    funGetKouhoHinban1_4_1 = 4
                    GoTo Apl_Exit
                ElseIf RET = 0 Then
                    ' DK温度チェックNGの値を保持する
                    ReDim Preserve lsDKCodeList(UBound(lsDKCodeList) + 1) As String
                    lsDKCodeList(UBound(lsDKCodeList)) = lsDKCodeListWork(iCnt)
                End If
            Next iCnt
                
            'DK温度チェックNG以外のデータを取得する
            If UBound(lsDKCodeList) <> 0 Then
                sql = sql & " AND (" & sChkTable2 & "." & tbl_spec1_4_1(0).HSXDKTMP1 & " NOT IN (" & vbCrLf
                For iCnt = 1 To UBound(lsDKCodeList)
                    If iCnt <> 1 Then
                        sql = sql & ","
                    End If
                    sql = sql & "'" & lsDKCodeList(iCnt) & "'"
                Next iCnt
                sql = sql & "))" & vbCrLf
            End If
        End If
    End If
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

Make_Exit:
    sMakesql = sql
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funGetKouhoHinban1_4_1 = -4
    GoTo Apl_Exit
    
End Function

'------------------------------------------------
' 先行評価項目仕様比較SQL文作成
'------------------------------------------------

'概要      :指定された振替元品番と先行評価項目仕様値が一致している、または、マトリクスで一致している品番を抽出するSQL文を作成する。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sBlockid        ,I  ,String       :ﾌﾞﾛｯｸID、又は、結晶番号
'          :sOld_Hinban     ,I  ,String       :振替元品番
'          :sMakeSql        ,O  ,String       :作成SQL文
'          :戻り値          ,O  ,Integer      :取得の成否(0:正常取得, -1:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/10 新規作成　SB

Public Function funGetKouhoHinban1_5(sProccd As String, sBlockId As String, sOld_Hinban As String, sMakesql As String) As Integer



    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql As String               'SQL全体
    Dim rs  As OraDynaset           'RecordSet
    Dim sMakesql1   As String       '呼び出しファンクションSQL作成
    Dim sMakesql2   As String       '呼び出しファンクションSQL作成
    Dim sMakesql3   As String       '呼び出しファンクションSQL作成
    Dim sMakesql4   As String       '呼び出しファンクションSQL作成
    Dim sMakesql5   As String       '呼び出しファンクションSQL作成
    Dim sMakesql6   As String       '呼び出しファンクションSQL作成
    Dim sMakesql7   As String       '呼び出しファンクションSQL作成
    Dim sMakesql8   As String       '呼び出しファンクションSQL作成
    Dim sMakesql9   As String       '呼び出しファンクションSQL作成
    Dim sMakesql10  As String       '呼び出しファンクションSQL作成
    Dim sMakesql11  As String       '呼び出しファンクションSQL作成
    Dim sMakesql12  As String       '呼び出しファンクションSQL作成
    Dim sMakesql13  As String       '呼び出しファンクションSQL作成
    Dim sMakesql14  As String       '呼び出しファンクションSQL作成
    Dim sMakesql15  As String       '呼び出しファンクションSQL作成
    Dim sMakesql16  As String       '呼び出しファンクションSQL作成
    Dim sMakesql17  As String       '呼び出しファンクションSQL作成  '残存酸素追加　03/12/09 ooba
    Dim sMakesql18  As String       '呼び出しファンクションSQL作成  'GD-Den追加　05/01/27 ooba
    Dim sMakesql19  As String       '呼び出しファンクションSQL作成  'GD-DVD2追加　05/01/27 ooba
    Dim sMakesql20  As String       '呼び出しファンクションSQL作成  'GD-L/DL追加　05/01/27 ooba
    Dim sMakesql21  As String       '呼び出しファンクションSQL作成  'SPVNR追加　06/05/31 ooba

    On Error GoTo Apl_down
    
    '戻り値初期化
    funGetKouhoHinban1_5 = 0
    
    '------------------------------------------ 振替元品種仕様データ取得 ------------------------------------------------------
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
    
    'GD仕様取得追加　05/01/27 ooba
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    sql = sql & "       E026.HWFDENHS,E026.HWFDENMN,E026.HWFDENMX,  E026.HWFDVDHS,E026.HWFDVDMNN,E026.HWFDVDMXN,E026.HWFLDLHS,E026.HWFLDLMN,E026.HWFLDLMX,  E026.HWFGDKHN, E026.HWFGDSZY,  " & vbCrLf
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
    ''検査頻度_抜ﾃﾞｰﾀ取得　04/04/13 ooba
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''    sql = sql & "       E021.HWFRKHNN, E025.HWFONKHN, E029.HWFOF1KN, E029.HWFOF2KN, E029.HWFOF3KN, E029.HWFOF4KN, E029.HWFBM1KN, E029.HWFBM2KN, E029.HWFBM3KN,               " & vbCrLf
    sql = sql & "       E021.HWFRKHNN, E025.HWFONKHN, E029.HWFOF1KN, E029.HWFOF2KN, E029.HWFOF3KN, E029.HWFBM1KN, E029.HWFBM2KN, E029.HWFBM3KN,               " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数取得追加
'    sql = sql & "       E025.HWFOS1KN, E025.HWFOS2KN, E025.HWFOS3KN, E026.HWFDSOKN, E024.HWFMKKHN, E028.HWFSPVKN, E028.HWFDLKHN, E025.HWFZOKHN                               " & vbCrLf
'    sql = sql & "FROM   TBCME021 E021,TBCME025 E025,TBCME029 E029,TBCME028 E028,TBCME026 E026,TBCME024 E024 " & vbCrLf
    sql = sql & "       E025.HWFOS1KN, E025.HWFOS2KN, E025.HWFOS3KN, E026.HWFDSOKN, E024.HWFMKKHN, E028.HWFSPVKN, E028.HWFDLKHN, E025.HWFZOKHN,E036.HWFGDLINE                " & vbCrLf
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    sql = sql & "       ,E025.HWFANTNP " & vbCrLf   'AN温度
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
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
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数取得追加
    sql = sql & "WHERE  E021.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E021.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E021.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E021.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' AND " & vbCrLf
    sql = sql & "       E025.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E025.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E025.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E025.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' AND " & vbCrLf
    sql = sql & "       E029.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E029.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E029.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E029.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' AND " & vbCrLf
    sql = sql & "       E028.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E028.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E028.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E028.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' AND " & vbCrLf
    sql = sql & "       E026.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E026.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E026.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E026.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' AND " & vbCrLf
    sql = sql & "       E024.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E024.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E024.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数取得追加
'    sql = sql & "       E024.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' " & vbCrLf
    sql = sql & "       E024.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' AND " & vbCrLf
    sql = sql & "       E036.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E036.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E036.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' AND " & vbCrLf
    sql = sql & "       E048.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E048.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E048.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E048.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "'  " & vbCrLf
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
'    sql = sql & "WHERE  E021.HINBAN || TO_CHAR(E021.MNOREVNO, 'FM00') || E021.FACTORY || E021.OPECOND   =   '" & sOld_Hinban & "' AND " & vbCrLf
'    sql = sql & "       E025.HINBAN || TO_CHAR(E025.MNOREVNO, 'FM00') || E025.FACTORY || E025.OPECOND   =   '" & sOld_Hinban & "' AND " & vbCrLf
'    sql = sql & "       E029.HINBAN || TO_CHAR(E029.MNOREVNO, 'FM00') || E029.FACTORY || E029.OPECOND   =   '" & sOld_Hinban & "' AND " & vbCrLf
'    sql = sql & "       E028.HINBAN || TO_CHAR(E028.MNOREVNO, 'FM00') || E028.FACTORY || E028.OPECOND   =   '" & sOld_Hinban & "' AND " & vbCrLf
'    sql = sql & "       E026.HINBAN || TO_CHAR(E026.MNOREVNO, 'FM00') || E026.FACTORY || E026.OPECOND   =   '" & sOld_Hinban & "' AND " & vbCrLf
'    sql = sql & "       E024.HINBAN || TO_CHAR(E024.MNOREVNO, 'FM00') || E024.FACTORY || E024.OPECOND   =   '" & sOld_Hinban & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        funGetKouhoHinban1_5 = 15001
        GoTo db_Error
    End If
    
    '取得データセット
    '検査頻度_抜ﾃﾞｰﾀ追加　04/04/13 ooba
    Erase tbl_spec1_5
    With tbl_spec1_5(0)
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
        If IsNull(rs("HWFOSF1PTK")) = False Then .HWFOSF1PTK = rs("HWFOSF1PTK") Else .HWFOSF1PTK = " "      'パターン区分
        If IsNull(rs("HWFOF1KN")) = False Then .HWFOF1KN = rs("HWFOF1KN") Else .HWFOF1KN = " "              '検査頻度_抜
        'OSF2
        If IsNull(rs("HWFOF2HS")) = False Then .HWFOF2HS = rs("HWFOF2HS") Else .HWFOF2HS = " "              '保証方法_対象
        If IsNull(rs("HWFOF2SH")) = False Then .HWFOF2SH = rs("HWFOF2SH") Else .HWFOF2SH = " "              '測定位置_方
        If IsNull(rs("HWFOF2SR")) = False Then .HWFOF2SR = rs("HWFOF2SR") Else .HWFOF2SR = " "              '測定位置_領
        If IsNull(rs("HWFOF2NS")) = False Then .HWFOF2NS = rs("HWFOF2NS") Else .HWFOF2NS = " "              '熱処理法
        If IsNull(rs("HWFOF2SZ")) = False Then .HWFOF2SZ = rs("HWFOF2SZ") Else .HWFOF2SZ = " "              '測定条件
        If IsNull(rs("HWFOF2ET")) = False Then .HWFOF2ET = rs("HWFOF2ET") Else .HWFOF2ET = 0                '選択ET代
        If IsNull(rs("HWFOSF2PTK")) = False Then .HWFOSF2PTK = rs("HWFOSF2PTK") Else .HWFOSF2PTK = " "      'パターン区分
        If IsNull(rs("HWFOF2KN")) = False Then .HWFOF2KN = rs("HWFOF2KN") Else .HWFOF2KN = " "              '検査頻度_抜
        'OSF3
        If IsNull(rs("HWFOF3HS")) = False Then .HWFOF3HS = rs("HWFOF3HS") Else .HWFOF3HS = " "              '保証方法_対象
        If IsNull(rs("HWFOF3SH")) = False Then .HWFOF3SH = rs("HWFOF3SH") Else .HWFOF3SH = " "              '測定位置_方
        If IsNull(rs("HWFOF3SR")) = False Then .HWFOF3SR = rs("HWFOF3SR") Else .HWFOF3SR = " "              '測定位置_領
        If IsNull(rs("HWFOF3NS")) = False Then .HWFOF3NS = rs("HWFOF3NS") Else .HWFOF3NS = " "              '熱処理法
        If IsNull(rs("HWFOF3SZ")) = False Then .HWFOF3SZ = rs("HWFOF3SZ") Else .HWFOF3SZ = " "              '測定条件
        If IsNull(rs("HWFOF3ET")) = False Then .HWFOF3ET = rs("HWFOF3ET") Else .HWFOF3ET = 0                '選択ET代
        If IsNull(rs("HWFOSF3PTK")) = False Then .HWFOSF3PTK = rs("HWFOSF3PTK") Else .HWFOSF3PTK = " "      'パターン区分
        If IsNull(rs("HWFOF3KN")) = False Then .HWFOF3KN = rs("HWFOF3KN") Else .HWFOF3KN = " "              '検査頻度_抜
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''        'OSF4
'''        If IsNull(rs("HWFOF4HS")) = False Then .HWFOF4HS = rs("HWFOF4HS") Else .HWFOF4HS = " "              '保証方法_対象
'''        If IsNull(rs("HWFOF4SH")) = False Then .HWFOF4SH = rs("HWFOF4SH") Else .HWFOF4SH = " "              '測定位置_方
'''        If IsNull(rs("HWFOF4SR")) = False Then .HWFOF4SR = rs("HWFOF4SR") Else .HWFOF4SR = " "              '測定位置_領
'''        If IsNull(rs("HWFOF4NS")) = False Then .HWFOF4NS = rs("HWFOF4NS") Else .HWFOF4NS = " "              '熱処理法
'''        If IsNull(rs("HWFOF4SZ")) = False Then .HWFOF4SZ = rs("HWFOF4SZ") Else .HWFOF4SZ = " "              '測定条件
'''        If IsNull(rs("HWFOF4ET")) = False Then .HWFOF4ET = rs("HWFOF4ET") Else .HWFOF4ET = 0                '選択ET代
'''        If IsNull(rs("HWFOSF4PTK")) = False Then .HWFOSF4PTK = rs("HWFOSF4PTK") Else .HWFOSF4PTK = " "      'パターン区分
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
    '*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
        'ﾗｲﾝ数
        If IsNull(rs("HWFGDLINE")) = False Then .HWFGDLINE = rs("HWFGDLINE") Else .HWFGDLINE = " "          'ﾗｲﾝ数
    '*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
        If IsNull(rs("HWFGDSZY")) = False Then .HWFGDSZY = rs("HWFGDSZY") Else .HWFGDSZY = " "               '測定条件
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        If IsNull(rs("HWFANTNP")) = False Then .HWFANTNP = rs("HWFANTNP") Else .HWFANTNP = " "              'AN温度
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    End With
    
    Set rs = Nothing
'    On Error GoTo Apl_down
    '------------------------------------------ 指示取得 ------------------------------------------------------
    sMakesql1 = ""
    sMakesql2 = ""
    sMakesql3 = ""
    sMakesql4 = ""
    sMakesql5 = ""
    sMakesql6 = ""
    sMakesql7 = ""
    sMakesql8 = ""
    sMakesql9 = ""
    sMakesql10 = ""
    sMakesql11 = ""
    sMakesql12 = ""
    sMakesql13 = ""
    sMakesql14 = ""
    sMakesql15 = ""
    sMakesql16 = ""
    sMakesql17 = ""         '残存酸素仕様取得SQL追加　03/12/09 ooba
    sMakesql18 = ""         'GD-Den仕様取得SQL追加　05/01/27 ooba
    sMakesql19 = ""         'GD-DVD2仕様取得SQL追加　05/01/27 ooba
    sMakesql20 = ""         'GD-L/DL仕様取得SQL追加　05/01/27 ooba
    sMakesql21 = ""         'SPVNR仕様取得SQL追加　06/05/31 ooba
    
    '比抵抗
    sResult = ""
    RET = funCodeDBGet("SB", "15", "RS", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15010
        GoTo Apl_Exit
    End If
    Erase tbl_spec1_5_1
    sMakesql = ""
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFRHWYS
    tbl_spec1_5_1(0).HOSYOU1 = "HWFRHWYS"
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFRKHNN        '04/04/13 ooba
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFRKHNN"                    '04/04/13 ooba
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    tbl_spec1_5_1(0).Antnp = tbl_spec1_5(0).HWFANTNP
    tbl_spec1_5_1(0).ANTNP1 = "HWFANTNP"
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'引数にどのチェックから呼んだか、渡す
'    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E021", sMakesql)
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E021", sMakesql, "", "Rs")
'↑変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15010 + RET
        GoTo Apl_Exit
    End If
    sMakesql1 = sMakesql
    '酸素濃度
    sResult = ""
    RET = funCodeDBGet("SB", "15", "OI", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15020
        GoTo Apl_Exit
    End If
    Erase tbl_spec1_5_1
    sMakesql = ""
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFONHWS
    tbl_spec1_5_1(0).HOSYOU1 = "HWFONHWS"
    tbl_spec1_5_1(0).SOKU_TEN = tbl_spec1_5(0).HWFONSPT         '08/01/29 ooba
    tbl_spec1_5_1(0).SOKU_TEN1 = "HWFONSPT"                     '08/01/29 ooba
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFONKHN        '04/04/13 ooba
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFONKHN"                    '04/04/13 ooba
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    tbl_spec1_5_1(0).Antnp = tbl_spec1_5(0).HWFANTNP
    tbl_spec1_5_1(0).ANTNP1 = "HWFANTNP"
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'引数にどのチェックから呼んだか、渡す
'    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E025", sMakesql)
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E025", sMakesql, "", "Oi")
'↑変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15020 + RET
        GoTo Apl_Exit
    End If
    sMakesql2 = sMakesql
    'ＯＳＦ1
    sResult = ""
    RET = funCodeDBGet("SB", "15", "O1", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15030
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_5_1
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFOF1HS
    tbl_spec1_5_1(0).HOSYOU1 = "HWFOF1HS"
    tbl_spec1_5_1(0).SOKU_HOU = tbl_spec1_5(0).HWFOF1SH
    tbl_spec1_5_1(0).SOKU_HOU1 = "HWFOF1SH"
    tbl_spec1_5_1(0).SOKU_RYOU = tbl_spec1_5(0).HWFOF1SR
    tbl_spec1_5_1(0).SOKU_RYOU1 = "HWFOF1SR"
    tbl_spec1_5_1(0).NETSU = tbl_spec1_5(0).HWFOF1NS
    tbl_spec1_5_1(0).NETSU1 = "HWFOF1NS"
    tbl_spec1_5_1(0).JOUKEN = tbl_spec1_5(0).HWFOF1SZ
    tbl_spec1_5_1(0).JOUKEN1 = "HWFOF1SZ"
    tbl_spec1_5_1(0).ET = tbl_spec1_5(0).HWFOF1ET
    tbl_spec1_5_1(0).ET1 = "HWFOF1ET"
    tbl_spec1_5_1(0).PATTERN = tbl_spec1_5(0).HWFOSF1PTK
    tbl_spec1_5_1(0).PATTERN1 = "HWFOSF1PTK"
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFOF1KN        '04/04/13 ooba
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFOF1KN"                    '04/04/13 ooba
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    tbl_spec1_5_1(0).Antnp = tbl_spec1_5(0).HWFANTNP
    tbl_spec1_5_1(0).ANTNP1 = "HWFANTNP"
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'引数にどのチェックから呼んだか、渡す
'    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E029", sMakesql)
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E029", sMakesql, "", "L1")
'↑変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15030 + RET
        GoTo Apl_Exit
    End If
    sMakesql3 = sMakesql
    'ＯＳＦ２
    sResult = ""
    RET = funCodeDBGet("SB", "15", "O2", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15040
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_5_1
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFOF2HS
    tbl_spec1_5_1(0).HOSYOU1 = "HWFOF2HS"
    tbl_spec1_5_1(0).SOKU_HOU = tbl_spec1_5(0).HWFOF2SH
    tbl_spec1_5_1(0).SOKU_HOU1 = "HWFOF2SH"
    tbl_spec1_5_1(0).SOKU_RYOU = tbl_spec1_5(0).HWFOF2SR
    tbl_spec1_5_1(0).SOKU_RYOU1 = "HWFOF2SR"
    tbl_spec1_5_1(0).NETSU = tbl_spec1_5(0).HWFOF2NS
    tbl_spec1_5_1(0).NETSU1 = "HWFOF2NS"
    tbl_spec1_5_1(0).JOUKEN = tbl_spec1_5(0).HWFOF2SZ
    tbl_spec1_5_1(0).JOUKEN1 = "HWFOF2SZ"
    tbl_spec1_5_1(0).ET = tbl_spec1_5(0).HWFOF2ET
    tbl_spec1_5_1(0).ET1 = "HWFOF2ET"
    tbl_spec1_5_1(0).PATTERN = tbl_spec1_5(0).HWFOSF2PTK
    tbl_spec1_5_1(0).PATTERN1 = "HWFOSF2PTK"
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFOF2KN        '04/04/13 ooba
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFOF2KN"                    '04/04/13 ooba
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    tbl_spec1_5_1(0).Antnp = tbl_spec1_5(0).HWFANTNP
    tbl_spec1_5_1(0).ANTNP1 = "HWFANTNP"
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'引数にどのチェックから呼んだか、渡す
'    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E029", sMakesql)
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E029", sMakesql, "", "L2")
'↑変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15040 + RET
        GoTo Apl_Exit
    End If
    sMakesql4 = sMakesql
    'ＯＳＦ３
    sResult = ""
    RET = funCodeDBGet("SB", "15", "O3", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15050
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_5_1
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFOF3HS
    tbl_spec1_5_1(0).HOSYOU1 = "HWFOF3HS"
    tbl_spec1_5_1(0).SOKU_HOU = tbl_spec1_5(0).HWFOF3SH
    tbl_spec1_5_1(0).SOKU_HOU1 = "HWFOF3SH"
    tbl_spec1_5_1(0).SOKU_RYOU = tbl_spec1_5(0).HWFOF3SR
    tbl_spec1_5_1(0).SOKU_RYOU1 = "HWFOF3SR"
    tbl_spec1_5_1(0).NETSU = tbl_spec1_5(0).HWFOF3NS
    tbl_spec1_5_1(0).NETSU1 = "HWFOF3NS"
    tbl_spec1_5_1(0).JOUKEN = tbl_spec1_5(0).HWFOF3SZ
    tbl_spec1_5_1(0).JOUKEN1 = "HWFOF3SZ"
    tbl_spec1_5_1(0).ET = tbl_spec1_5(0).HWFOF3ET
    tbl_spec1_5_1(0).ET1 = "HWFOF3ET"
    tbl_spec1_5_1(0).PATTERN = tbl_spec1_5(0).HWFOSF3PTK
    tbl_spec1_5_1(0).PATTERN1 = "HWFOSF3PTK"
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFOF3KN        '04/04/13 ooba
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFOF3KN"                    '04/04/13 ooba
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    tbl_spec1_5_1(0).Antnp = tbl_spec1_5(0).HWFANTNP
    tbl_spec1_5_1(0).ANTNP1 = "HWFANTNP"
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'引数にどのチェックから呼んだか、渡す
'    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E029", sMakesql)
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E029", sMakesql, "", "L3")
'↑変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15050 + RET
        GoTo Apl_Exit
    End If
    sMakesql5 = sMakesql
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''    'ＯＳＦ４
'''    sResult = ""
'''    RET = funCodeDBGet("SB", "15", "O4", 0, " ", sResult)
'''    If RET <> 0 Then
'''        funGetKouhoHinban1_5 = 15060
'''        GoTo Apl_Exit
'''    End If
'''    sMakesql = ""
'''    Erase tbl_spec1_5_1
'''    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFOF4HS
'''    tbl_spec1_5_1(0).HOSYOU1 = "HWFOF4HS"
'''    tbl_spec1_5_1(0).SOKU_HOU = tbl_spec1_5(0).HWFOF4SH
'''    tbl_spec1_5_1(0).SOKU_HOU1 = "HWFOF4SH"
'''    tbl_spec1_5_1(0).SOKU_RYOU = tbl_spec1_5(0).HWFOF4SR
'''    tbl_spec1_5_1(0).SOKU_RYOU1 = "HWFOF4SR"
'''    tbl_spec1_5_1(0).NETSU = tbl_spec1_5(0).HWFOF4NS
'''    tbl_spec1_5_1(0).NETSU1 = "HWFOF4NS"
'''    tbl_spec1_5_1(0).JOUKEN = tbl_spec1_5(0).HWFOF4SZ
'''    tbl_spec1_5_1(0).JOUKEN1 = "HWFOF4SZ"
'''    tbl_spec1_5_1(0).ET = tbl_spec1_5(0).HWFOF4ET
'''    tbl_spec1_5_1(0).ET1 = "HWFOF4ET"
'''    tbl_spec1_5_1(0).PATTERN = tbl_spec1_5(0).HWFOSF4PTK
'''    tbl_spec1_5_1(0).PATTERN1 = "HWFOSF4PTK"
'''    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFOF4KN        '04/04/13 ooba
'''    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFOF4KN"                    '04/04/13 ooba
''''↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'''    tbl_spec1_5_1(0).Antnp = tbl_spec1_5(0).HWFANTNP
'''    tbl_spec1_5_1(0).ANTNP1 = "HWFANTNP"
''''↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
''''↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
''''引数にどのチェックから呼んだか、渡す
'''    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E029", sMakesql, "", "L4")
''''↑変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'''    If RET <> 0 Then
'''        funGetKouhoHinban1_5 = 15060 + RET
'''        GoTo Apl_Exit
'''    End If
'''    sMakesql6 = sMakesql
    
    'ＳＩＲＤ
    sResult = ""
    RET = funCodeDBGet("SB", "15", "SD", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15060
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_5_1                                         '引数ﾃｰﾌﾞﾙｸﾘｱ
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFSIRDHS          '軸状転位保証方法＿処
    tbl_spec1_5_1(0).HOSYOU1 = "HWFSIRDHS"                      '軸状転位保証方法＿処
    tbl_spec1_5_1(0).JOUKEN = tbl_spec1_5(0).HWFSIRDSZ          '軸状転位測定条件
    tbl_spec1_5_1(0).JOUKEN1 = "HWFSIRDSZ"                      '軸状転位測定条件
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFSIRDKN       '軸状転位検査頻度＿抜
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFSIRDKN"                   '軸状転位検査頻度＿抜
    tbl_spec1_5_1(0).HWFSIRDMX = tbl_spec1_5(0).HWFSIRDMX       '軸状転位上限
    tbl_spec1_5_1(0).HWFSIRDMX1 = "HWFSIRDMX"                   '軸状転位上限
    tbl_spec1_5_1(0).HWFSIRDHT = tbl_spec1_5(0).HWFSIRDHT       '軸状転位保証方法＿対
    tbl_spec1_5_1(0).HWFSIRDHT1 = "HWFSIRDHT"                   '軸状転位保証方法＿対
    tbl_spec1_5_1(0).HWFSIRDKM = tbl_spec1_5(0).HWFSIRDKM       '軸状転位検査頻度＿枚
    tbl_spec1_5_1(0).HWFSIRDKM1 = "HWFSIRDKM"                   '軸状転位検査頻度＿枚
    tbl_spec1_5_1(0).HWFSIRDKH = tbl_spec1_5(0).HWFSIRDKH       '軸状転位検査頻度＿保
    tbl_spec1_5_1(0).HWFSIRDKH1 = "HWFSIRDKH"                   '軸状転位検査頻度＿保
    tbl_spec1_5_1(0).HWFSIRDKU = tbl_spec1_5(0).HWFSIRDKU       '軸状転位検査頻度＿ウ
    tbl_spec1_5_1(0).HWFSIRDKU1 = "HWFSIRDKU"                   '軸状転位検査頻度＿ウ
    tbl_spec1_5_1(0).Antnp = tbl_spec1_5(0).HWFANTNP            '2.1.1 AN温度 振替可否チェック
    tbl_spec1_5_1(0).ANTNP1 = "HWFANTNP"                        '2.1.1 AN温度 振替可否チェック
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E048", sMakesql, "", "SD")
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15060 + RET
        GoTo Apl_Exit
    End If
    sMakesql6 = sMakesql
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
    'ＢＭＤ１
    sResult = ""
    RET = funCodeDBGet("SB", "15", "B1", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15070
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_5_1
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFBM1HS
    tbl_spec1_5_1(0).HOSYOU1 = "HWFBM1HS"
    tbl_spec1_5_1(0).SOKU_HOU = tbl_spec1_5(0).HWFBM1SH
    tbl_spec1_5_1(0).SOKU_HOU1 = "HWFBM1SH"
    tbl_spec1_5_1(0).SOKU_TEN = tbl_spec1_5(0).HWFBM1ST
    tbl_spec1_5_1(0).SOKU_TEN1 = "HWFBM1ST"
    tbl_spec1_5_1(0).SOKU_RYOU = tbl_spec1_5(0).HWFBM1SR
    tbl_spec1_5_1(0).SOKU_RYOU1 = "HWFBM1SR"
    tbl_spec1_5_1(0).NETSU = tbl_spec1_5(0).HWFBM1NS
    tbl_spec1_5_1(0).NETSU1 = "HWFBM1NS"
    tbl_spec1_5_1(0).JOUKEN = tbl_spec1_5(0).HWFBM1SZ
    tbl_spec1_5_1(0).JOUKEN1 = "HWFBM1SZ"
    tbl_spec1_5_1(0).ET = tbl_spec1_5(0).HWFBM1ET
    tbl_spec1_5_1(0).ET1 = "HWFBM1ET"
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFBM1KN        '04/04/13 ooba
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFBM1KN"                    '04/04/13 ooba
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    tbl_spec1_5_1(0).Antnp = tbl_spec1_5(0).HWFANTNP
    tbl_spec1_5_1(0).ANTNP1 = "HWFANTNP"
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'引数にどのチェックから呼んだか、渡す
'    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E029", sMakesql)
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E029", sMakesql, "", "B1")
'↑変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15070 + RET
        GoTo Apl_Exit
    End If
    sMakesql7 = sMakesql
    'ＢＭＤ２
    sResult = ""
    RET = funCodeDBGet("SB", "15", "B2", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15080
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_5_1
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFBM2HS
    tbl_spec1_5_1(0).HOSYOU1 = "HWFBM2HS"
    tbl_spec1_5_1(0).SOKU_HOU = tbl_spec1_5(0).HWFBM2SH
    tbl_spec1_5_1(0).SOKU_HOU1 = "HWFBM2SH"
    tbl_spec1_5_1(0).SOKU_TEN = tbl_spec1_5(0).HWFBM2ST
    tbl_spec1_5_1(0).SOKU_TEN1 = "HWFBM2ST"
    tbl_spec1_5_1(0).SOKU_RYOU = tbl_spec1_5(0).HWFBM2SR
    tbl_spec1_5_1(0).SOKU_RYOU1 = "HWFBM2SR"
    tbl_spec1_5_1(0).NETSU = tbl_spec1_5(0).HWFBM2NS
    tbl_spec1_5_1(0).NETSU1 = "HWFBM2NS"
    tbl_spec1_5_1(0).JOUKEN = tbl_spec1_5(0).HWFBM2SZ
    tbl_spec1_5_1(0).JOUKEN1 = "HWFBM2SZ"
    tbl_spec1_5_1(0).ET = tbl_spec1_5(0).HWFBM2ET
    tbl_spec1_5_1(0).ET1 = "HWFBM2ET"
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFBM2KN        '04/04/13 ooba
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFBM2KN"                    '04/04/13 ooba
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    tbl_spec1_5_1(0).Antnp = tbl_spec1_5(0).HWFANTNP
    tbl_spec1_5_1(0).ANTNP1 = "HWFANTNP"
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'引数にどのチェックから呼んだか、渡す
'    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E029", sMakesql)
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E029", sMakesql, "", "B2")
'↑変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15080 + RET
        GoTo Apl_Exit
    End If
    sMakesql8 = sMakesql
    'ＢＭＤ３
    sResult = ""
    RET = funCodeDBGet("SB", "15", "B3", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15090
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_5_1
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFBM3HS
    tbl_spec1_5_1(0).HOSYOU1 = "HWFBM3HS"
    tbl_spec1_5_1(0).SOKU_HOU = tbl_spec1_5(0).HWFBM3SH
    tbl_spec1_5_1(0).SOKU_HOU1 = "HWFBM3SH"
    tbl_spec1_5_1(0).SOKU_TEN = tbl_spec1_5(0).HWFBM3ST
    tbl_spec1_5_1(0).SOKU_TEN1 = "HWFBM3ST"
    tbl_spec1_5_1(0).SOKU_RYOU = tbl_spec1_5(0).HWFBM3SR
    tbl_spec1_5_1(0).SOKU_RYOU1 = "HWFBM3SR"
    tbl_spec1_5_1(0).NETSU = tbl_spec1_5(0).HWFBM3NS
    tbl_spec1_5_1(0).NETSU1 = "HWFBM3NS"
    tbl_spec1_5_1(0).JOUKEN = tbl_spec1_5(0).HWFBM3SZ
    tbl_spec1_5_1(0).JOUKEN1 = "HWFBM3SZ"
    tbl_spec1_5_1(0).ET = tbl_spec1_5(0).HWFBM3ET
    tbl_spec1_5_1(0).ET1 = "HWFBM3ET"
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFBM3KN        '04/04/13 ooba
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFBM3KN"                    '04/04/13 ooba
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    tbl_spec1_5_1(0).Antnp = tbl_spec1_5(0).HWFANTNP
    tbl_spec1_5_1(0).ANTNP1 = "HWFANTNP"
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'引数にどのチェックから呼んだか、渡す
'    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E029", sMakesql)
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E029", sMakesql, "", "B3")
'↑変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15090 + RET
        GoTo Apl_Exit
    End If
    sMakesql9 = sMakesql
    '酸素析出１
    sResult = ""
    RET = funCodeDBGet("SB", "15", "D1", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15100
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_5_1
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFOS1HS
    tbl_spec1_5_1(0).HOSYOU1 = "HWFOS1HS"
    tbl_spec1_5_1(0).NETSU = tbl_spec1_5(0).HWFOS1NS
    tbl_spec1_5_1(0).NETSU1 = "HWFOS1NS"
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFOS1KN        '04/04/13 ooba
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFOS1KN"                    '04/04/13 ooba
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    tbl_spec1_5_1(0).Antnp = tbl_spec1_5(0).HWFANTNP
    tbl_spec1_5_1(0).ANTNP1 = "HWFANTNP"
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'引数にどのチェックから呼んだか、渡す
'    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E025", sMakesql)
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E025", sMakesql, "", "D1")
'↑変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15100 + RET
        GoTo Apl_Exit
    End If
    sMakesql10 = sMakesql
    '酸素析出２
    sResult = ""
    RET = funCodeDBGet("SB", "15", "D2", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15110
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_5_1
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFOS2HS
    tbl_spec1_5_1(0).HOSYOU1 = "HWFOS2HS"
    tbl_spec1_5_1(0).NETSU = tbl_spec1_5(0).HWFOS2NS
    tbl_spec1_5_1(0).NETSU1 = "HWFOS2NS"
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFOS2KN        '04/04/13 ooba
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFOS2KN"                    '04/04/13 ooba
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    tbl_spec1_5_1(0).Antnp = tbl_spec1_5(0).HWFANTNP
    tbl_spec1_5_1(0).ANTNP1 = "HWFANTNP"
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'引数にどのチェックから呼んだか、渡す
'    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E025", sMakesql)
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E025", sMakesql, "", "D2")
'↑変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15110 + RET
        GoTo Apl_Exit
    End If
    sMakesql11 = sMakesql
    '酸素析出３
    sResult = ""
    RET = funCodeDBGet("SB", "15", "D3", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15120
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_5_1
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFOS3HS
    tbl_spec1_5_1(0).HOSYOU1 = "HWFOS3HS"
    tbl_spec1_5_1(0).NETSU = tbl_spec1_5(0).HWFOS3NS
    tbl_spec1_5_1(0).NETSU1 = "HWFOS3NS"
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFOS3KN        '04/04/13 ooba
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFOS3KN"                    '04/04/13 ooba
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    tbl_spec1_5_1(0).Antnp = tbl_spec1_5(0).HWFANTNP
    tbl_spec1_5_1(0).ANTNP1 = "HWFANTNP"
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'引数にどのチェックから呼んだか、渡す
'    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E025", sMakesql)
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E025", sMakesql, "", "D3")
'↑変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15120 + RET
        GoTo Apl_Exit
    End If
    sMakesql12 = sMakesql
    'ＤＳＯＤ
    sResult = ""
    RET = funCodeDBGet("SB", "15", "DS", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15130
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_5_1
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFDSOHS
    tbl_spec1_5_1(0).HOSYOU1 = "HWFDSOHS"
    tbl_spec1_5_1(0).NETSU = tbl_spec1_5(0).HWFDSONWY
    tbl_spec1_5_1(0).NETSU1 = "HWFDSONWY"
    tbl_spec1_5_1(0).PATTERN = tbl_spec1_5(0).HWFDSOPTK         'ﾊﾟﾀｰﾝ区分追加　04/07/29 ooba
    tbl_spec1_5_1(0).PATTERN1 = "HWFDSOPTK"                     'ﾊﾟﾀｰﾝ区分追加　04/07/29 ooba
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFDSOKN        '04/04/13 ooba
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFDSOKN"                    '04/04/13 ooba
'    ret = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E026", sMakesql)
    'GD/DSOD熱処理条件追加　06/12/22 ooba START =========================================>
    tbl_spec1_5_1(0).Antnp = tbl_spec1_5(0).HWFANTNP
    tbl_spec1_5_1(0).ANTNP1 = "HWFANTNP"
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E026", sMakesql, "", "DS")
    'GD/DSOD熱処理条件追加　06/12/22 ooba END ===========================================>

    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15130 + RET
        GoTo Apl_Exit
    End If
    sMakesql13 = sMakesql
    'ＤＺ
    sResult = ""
    RET = funCodeDBGet("SB", "15", "DZ", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15140
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_5_1
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFMKHWS
    tbl_spec1_5_1(0).HOSYOU1 = "HWFMKHWS"
    tbl_spec1_5_1(0).SOKU_HOU = tbl_spec1_5(0).HWFMKSPH
    tbl_spec1_5_1(0).SOKU_HOU1 = "HWFMKSPH"
    tbl_spec1_5_1(0).SOKU_TEN = tbl_spec1_5(0).HWFMKSPT
    tbl_spec1_5_1(0).SOKU_TEN1 = "HWFMKSPT"
    tbl_spec1_5_1(0).SOKU_RYOU = tbl_spec1_5(0).HWFMKSPR
    tbl_spec1_5_1(0).SOKU_RYOU1 = "HWFMKSPR"
    tbl_spec1_5_1(0).NETSU = tbl_spec1_5(0).HWFMKNSW
    tbl_spec1_5_1(0).NETSU1 = "HWFMKNSW"
    tbl_spec1_5_1(0).JOUKEN = tbl_spec1_5(0).HWFMKSZY
    tbl_spec1_5_1(0).JOUKEN1 = "HWFMKSZY"
    tbl_spec1_5_1(0).ET = tbl_spec1_5(0).HWFMKCET
    tbl_spec1_5_1(0).ET1 = "HWFMKCET"
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFMKKHN        '04/04/13 ooba
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFMKKHN"                    '04/04/13 ooba
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E024", sMakesql)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15140 + RET
        GoTo Apl_Exit
    End If
    sMakesql14 = sMakesql
    'ＳＰＶＦＥ
    sResult = ""
    RET = funCodeDBGet("SB", "15", "SP", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15150
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_5_1
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFSPVHS
    tbl_spec1_5_1(0).HOSYOU1 = "HWFSPVHS"
    tbl_spec1_5_1(0).SOKU_TEN = tbl_spec1_5(0).HWFSPVST
    tbl_spec1_5_1(0).SOKU_TEN1 = "HWFSPVST"
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFSPVKN        '04/04/13 ooba
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFSPVKN"                    '04/04/13 ooba
    
''Upd Start 2005/06/16 (TCS)T.Terauchi      SPV9点対応
    tbl_spec1_5_1(0).SOKU_HOU = tbl_spec1_5(0).HWFSPVSH         ''測定位置＿方
    tbl_spec1_5_1(0).SOKU_HOU1 = "HWFSPVSH"                     ''測定位置＿方
    tbl_spec1_5_1(0).SOKU_ICHI = tbl_spec1_5(0).HWFSPVSI        ''測定位置＿位
    tbl_spec1_5_1(0).SOKU_ICHI1 = "HWFSPVSI"                    ''測定位置＿位
''Upd End   2005/06/16 (TCS)T.Terauchi      SPV9点対応
    
    ''06/05/31 ooba START ============================================>
    tbl_spec1_5_1(0).PUAGEN = tbl_spec1_5(0).HWFSPVPUG
    tbl_spec1_5_1(0).PUAGEN1 = "HWFSPVPUG"
    tbl_spec1_5_1(0).PUAPER = tbl_spec1_5(0).HWFSPVPUR
    tbl_spec1_5_1(0).PUAPER1 = "HWFSPVPUR"
    ''06/05/31] ooba END ==============================================>
    
    '取得ﾃｰﾌﾞﾙ2追加　06/05/31 ooba
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E028", sMakesql, "E048")
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15150 + RET
        GoTo Apl_Exit
    End If
    sMakesql15 = sMakesql
    '拡散長
    sResult = ""
    RET = funCodeDBGet("SB", "15", "KL", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15160
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_5_1
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFDLHWS
    tbl_spec1_5_1(0).HOSYOU1 = "HWFDLHWS"
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFDLKHN        '04/04/13 ooba
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFDLKHN"                    '04/04/13 ooba
    
''Upd Start 2005/06/16 (TCS)T.Terauchi      SPV9点対応
    tbl_spec1_5_1(0).SOKU_HOU = tbl_spec1_5(0).HWFDLSPH         ''測定位置＿方
    tbl_spec1_5_1(0).SOKU_HOU1 = "HWFDLSPH"                     ''測定位置＿方
    tbl_spec1_5_1(0).SOKU_TEN = tbl_spec1_5(0).HWFDLSPT         ''測定位置＿点
    tbl_spec1_5_1(0).SOKU_TEN1 = "HWFDLSPT"                     ''測定位置＿点
    tbl_spec1_5_1(0).SOKU_ICHI = tbl_spec1_5(0).HWFDLSPI        ''測定位置＿位
    tbl_spec1_5_1(0).SOKU_ICHI1 = "HWFDLSPI"                    ''測定位置＿位
''Upd End   2005/06/16 (TCS)T.Terauchi      SPV9点対応
    
    ''06/05/31 ooba START ============================================>
    tbl_spec1_5_1(0).PUAGEN = tbl_spec1_5(0).HWFDLPUG
    tbl_spec1_5_1(0).PUAGEN1 = "HWFDLPUG"
    tbl_spec1_5_1(0).PUAPER = tbl_spec1_5(0).HWFDLPUR
    tbl_spec1_5_1(0).PUAPER1 = "HWFDLPUR"
    ''06/05/31 ooba END ==============================================>
    
    '取得ﾃｰﾌﾞﾙ2追加　06/05/31 ooba
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E028", sMakesql, "E048")
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15160 + RET
        GoTo Apl_Exit
    End If
    sMakesql16 = sMakesql
    
    ''残存酸素追加　03/12/09 ooba START ============================================>
    '残存酸素
    sResult = ""
    RET = funCodeDBGet("SB", "15", "AO", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15170
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_5_1
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFZOHWS
    tbl_spec1_5_1(0).HOSYOU1 = "HWFZOHWS"
    tbl_spec1_5_1(0).NETSU = tbl_spec1_5(0).HWFZONSW
    tbl_spec1_5_1(0).NETSU1 = "HWFZONSW"
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFZOKHN        '04/04/13 ooba
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFZOKHN"                    '04/04/13 ooba
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    tbl_spec1_5_1(0).Antnp = tbl_spec1_5(0).HWFANTNP
    tbl_spec1_5_1(0).ANTNP1 = "HWFANTNP"
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'引数にどのチェックから呼んだか、渡す
'    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E025", sMakesql)
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E025", sMakesql, "", "AO")
'↑変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15170 + RET
        GoTo Apl_Exit
    End If
    sMakesql17 = sMakesql
    ''残存酸素追加　03/12/09 ooba END ==============================================>
    
    ''GD追加　05/01/27 ooba START =================================================>
    'ＤＥＮ
    sResult = ""
    RET = funCodeDBGet("SB", "15", "DEN", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15180
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_5_1
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFDENHS
    tbl_spec1_5_1(0).HOSYOU1 = "HWFDENHS"
    tbl_spec1_5_1(0).Min = tbl_spec1_5(0).HWFDENMN
    tbl_spec1_5_1(0).Min1 = "HWFDENMN"
    tbl_spec1_5_1(0).max = tbl_spec1_5(0).HWFDENMX
    tbl_spec1_5_1(0).Max1 = "HWFDENMX"
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFGDKHN
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFGDKHN"
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
    tbl_spec1_5_1(0).LINE = tbl_spec1_5(0).HWFGDLINE
    tbl_spec1_5_1(0).Line1 = "HWFGDLINE"
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    tbl_spec1_5_1(0).HWFGDSZY = tbl_spec1_5(0).HWFGDSZY
    tbl_spec1_5_1(0).HWFGDSZY1 = "HWFGDSZY"
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
'    ret = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E026", sMakesql, "E036")
    'GD/DSOD熱処理条件追加　06/12/22 ooba START =========================================>
    tbl_spec1_5_1(0).Antnp = tbl_spec1_5(0).HWFANTNP
    tbl_spec1_5_1(0).ANTNP1 = "HWFANTNP"
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E026", sMakesql, "E036", "GD")
    'GD/DSOD熱処理条件追加　06/12/22 ooba END ===========================================>
    
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15180 + RET
        GoTo Apl_Exit
    End If
    sMakesql18 = sMakesql
    'ＤＶＤ２
    sResult = ""
    RET = funCodeDBGet("SB", "15", "DVD", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15190
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_5_1
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFDVDHS
    tbl_spec1_5_1(0).HOSYOU1 = "HWFDVDHS"
    tbl_spec1_5_1(0).Min = tbl_spec1_5(0).HWFDVDMNN
    tbl_spec1_5_1(0).Min1 = "HWFDVDMNN"
    tbl_spec1_5_1(0).max = tbl_spec1_5(0).HWFDVDMXN
    tbl_spec1_5_1(0).Max1 = "HWFDVDMXN"
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFGDKHN
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFGDKHN"
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
    tbl_spec1_5_1(0).LINE = tbl_spec1_5(0).HWFGDLINE
    tbl_spec1_5_1(0).Line1 = "HWFGDLINE"
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    tbl_spec1_5_1(0).HWFGDSZY = tbl_spec1_5(0).HWFGDSZY
    tbl_spec1_5_1(0).HWFGDSZY1 = "HWFGDSZY"
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
'    ret = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E026", sMakesql, "E036")
    'GD/DSOD熱処理条件追加　06/12/22 ooba START =========================================>
    tbl_spec1_5_1(0).Antnp = tbl_spec1_5(0).HWFANTNP
    tbl_spec1_5_1(0).ANTNP1 = "HWFANTNP"
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E026", sMakesql, "E036", "GD")
    'GD/DSOD熱処理条件追加　06/12/22 ooba END ===========================================>
    
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15190 + RET
        GoTo Apl_Exit
    End If
    sMakesql19 = sMakesql
    'Ｌ／ＤＬ
    sResult = ""
    RET = funCodeDBGet("SB", "15", "LDL", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15200
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_5_1
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFLDLHS
    tbl_spec1_5_1(0).HOSYOU1 = "HWFLDLHS"
    tbl_spec1_5_1(0).Min = tbl_spec1_5(0).HWFLDLMN
    tbl_spec1_5_1(0).Min1 = "HWFLDLMN"
    tbl_spec1_5_1(0).max = tbl_spec1_5(0).HWFLDLMX
    tbl_spec1_5_1(0).Max1 = "HWFLDLMX"
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFGDKHN
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFGDKHN"
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
    tbl_spec1_5_1(0).LINE = tbl_spec1_5(0).HWFGDLINE
    tbl_spec1_5_1(0).Line1 = "HWFGDLINE"
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    tbl_spec1_5_1(0).HWFGDSZY = tbl_spec1_5(0).HWFGDSZY
    tbl_spec1_5_1(0).HWFGDSZY1 = "HWFGDSZY"
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
'    ret = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E026", sMakesql, "E036")
    'GD/DSOD熱処理条件追加　06/12/22 ooba START =========================================>
    tbl_spec1_5_1(0).Antnp = tbl_spec1_5(0).HWFANTNP
    tbl_spec1_5_1(0).ANTNP1 = "HWFANTNP"
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E026", sMakesql, "E036", "GD")
    'GD/DSOD熱処理条件追加　06/12/22 ooba END ===========================================>
    
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15200 + RET
        GoTo Apl_Exit
    End If
    sMakesql20 = sMakesql
    ''GD追加　05/01/27 ooba END ===================================================>
    
    ''06/05/31 ooba START ============================================>
    'ＳＰＶＮＲ
    sResult = ""
    RET = funCodeDBGet("SB", "15", "NR", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15210
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_5_1
    tbl_spec1_5_1(0).HOSYOU = tbl_spec1_5(0).HWFNRHS        '保証方法＿対象
    tbl_spec1_5_1(0).HOSYOU1 = "HWFNRHS"
    tbl_spec1_5_1(0).SOKU_HOU = tbl_spec1_5(0).HWFNRSH      '測定位置＿方
    tbl_spec1_5_1(0).SOKU_HOU1 = "HWFNRSH"
    tbl_spec1_5_1(0).SOKU_TEN = tbl_spec1_5(0).HWFNRST      '測定位置＿点
    tbl_spec1_5_1(0).SOKU_TEN1 = "HWFNRST"
    tbl_spec1_5_1(0).SOKU_ICHI = tbl_spec1_5(0).HWFNRSI     '測定位置＿位
    tbl_spec1_5_1(0).SOKU_ICHI1 = "HWFNRSI"
    tbl_spec1_5_1(0).KENH_NUKI = tbl_spec1_5(0).HWFNRKN     '検査頻度＿抜
    tbl_spec1_5_1(0).KENH_NUKI1 = "HWFNRKN"
    tbl_spec1_5_1(0).PUAGEN = tbl_spec1_5(0).HWFNRPUG       'PUA限
    tbl_spec1_5_1(0).PUAGEN1 = "HWFNRPUG"
    tbl_spec1_5_1(0).PUAPER = tbl_spec1_5(0).HWFNRPUR       'PUA率
    tbl_spec1_5_1(0).PUAPER1 = "HWFNRPUR"
    
    RET = funGetKouhoHinban1_5_1(sResult, tbl_spec1_5_1(), "E048", sMakesql, "E048")
    If RET <> 0 Then
        funGetKouhoHinban1_5 = 15210 + RET
        GoTo Apl_Exit
    End If
    sMakesql21 = sMakesql
    ''06/05/31 ooba END ==============================================>
    
    '------------------------------------------ 振替元品種と同一仕様の品番を取得 ------------------------------------------------------
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT 'X' " & vbCrLf
'    sql = sql & "SELECT E021.HINBAN || TO_CHAR(E021.MNOREVNO, 'FM00') || E021.FACTORY || E021.OPECOND HINBAN " & vbCrLf
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
'    sql = sql & "FROM   TBCME021 E021, TBCME024 E024, TBCME025 E025, TBCME026 E026, TBCME029 E029, TBCME028 E028 " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''    sql = sql & "FROM   TBCME021 E021, TBCME024 E024, TBCME025 E025, TBCME026 E026, TBCME029 E029, TBCME028 E028,TBCME036 E036" & vbCrLf
    sql = sql & "FROM   TBCME021 E021, TBCME024 E024, TBCME025 E025, TBCME026 E026, TBCME029 E029, TBCME028 E028,TBCME036 E036,TBCME048 E048" & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki REP END(OSF4->SIRD)
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
    sql = sql & "WHERE  E018A.HINBAN                    = E021.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018A.MNOREVNO, 'FM00') = TO_CHAR(E021.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E018A.FACTORY                   = E021.FACTORY                      AND " & vbCrLf
    sql = sql & "       E018A.OPECOND                   = E021.OPECOND                      AND " & vbCrLf
    sql = sql & "       E021.HINBAN                     = E024.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E021.MNOREVNO, 'FM00')  = TO_CHAR(E024.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E021.FACTORY                    = E024.FACTORY                      AND " & vbCrLf
    sql = sql & "       E021.OPECOND                    = E024.OPECOND                      AND " & vbCrLf
    sql = sql & "       E021.HINBAN                     = E025.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E021.MNOREVNO, 'FM00')  = TO_CHAR(E025.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E021.FACTORY                    = E025.FACTORY                      AND " & vbCrLf
    sql = sql & "       E021.OPECOND                    = E025.OPECOND                      AND " & vbCrLf
    sql = sql & "       E021.HINBAN                     = E026.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E021.MNOREVNO, 'FM00')  = TO_CHAR(E026.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E021.FACTORY                    = E026.FACTORY                      AND " & vbCrLf
    sql = sql & "       E021.OPECOND                    = E026.OPECOND                      AND " & vbCrLf
    sql = sql & "       E021.HINBAN                     = E029.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E021.MNOREVNO, 'FM00')  = TO_CHAR(E029.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E021.FACTORY                    = E029.FACTORY                      AND " & vbCrLf
    sql = sql & "       E021.OPECOND                    = E029.OPECOND                      AND " & vbCrLf
    sql = sql & "       E021.HINBAN                     = E028.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E021.MNOREVNO, 'FM00')  = TO_CHAR(E028.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E021.FACTORY                    = E028.FACTORY                      AND " & vbCrLf
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
'    sql = sql & "       E021.OPECOND                    = E028.OPECOND " & vbCrLf
    sql = sql & "       E021.OPECOND                    = E028.OPECOND                      AND " & vbCrLf
    sql = sql & "       E021.HINBAN                     = E036.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E021.MNOREVNO, 'FM00')  = TO_CHAR(E036.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E021.FACTORY                    = E036.FACTORY                      AND " & vbCrLf
    sql = sql & "       E021.OPECOND                    = E036.OPECOND                      AND " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(OSF4->SIRD)
    sql = sql & "       E021.HINBAN                     = E048.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E021.MNOREVNO, 'FM00')  = TO_CHAR(E048.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E021.FACTORY                    = E048.FACTORY                      AND " & vbCrLf
    sql = sql & "       E021.OPECOND                    = E048.OPECOND                       " & vbCrLf
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END (OSF4->SIRD)
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
'    sql = sql & "WHERE  E021.HINBAN || TO_CHAR(E021.MNOREVNO, 'FM00') || E021.FACTORY || E021.OPECOND <> '" & sOld_Hinban & "' AND " & vbCrLf
'    sql = sql & "       E021.HINBAN || TO_CHAR(E021.MNOREVNO, 'FM00') || E021.FACTORY || E021.OPECOND = E024.HINBAN || TO_CHAR(E024.MNOREVNO, 'FM00') || E024.FACTORY || E024.OPECOND AND " & vbCrLf
'    sql = sql & "       E021.HINBAN || TO_CHAR(E021.MNOREVNO, 'FM00') || E021.FACTORY || E021.OPECOND = E025.HINBAN || TO_CHAR(E025.MNOREVNO, 'FM00') || E025.FACTORY || E025.OPECOND AND " & vbCrLf
'    sql = sql & "       E021.HINBAN || TO_CHAR(E021.MNOREVNO, 'FM00') || E021.FACTORY || E021.OPECOND = E026.HINBAN || TO_CHAR(E026.MNOREVNO, 'FM00') || E026.FACTORY || E026.OPECOND AND " & vbCrLf
'    sql = sql & "       E021.HINBAN || TO_CHAR(E021.MNOREVNO, 'FM00') || E021.FACTORY || E021.OPECOND = E029.HINBAN || TO_CHAR(E029.MNOREVNO, 'FM00') || E029.FACTORY || E029.OPECOND AND " & vbCrLf
'    sql = sql & "       E021.HINBAN || TO_CHAR(E021.MNOREVNO, 'FM00') || E021.FACTORY || E021.OPECOND = E028.HINBAN || TO_CHAR(E028.MNOREVNO, 'FM00') || E028.FACTORY || E028.OPECOND " & vbCrLf
    sql = sql & sMakesql1
    sql = sql & sMakesql2
    sql = sql & sMakesql3
    sql = sql & sMakesql4
    sql = sql & sMakesql5
    sql = sql & sMakesql6
    sql = sql & sMakesql7
    sql = sql & sMakesql8
    sql = sql & sMakesql9
    sql = sql & sMakesql10
    sql = sql & sMakesql11
    sql = sql & sMakesql12
    sql = sql & sMakesql13
    sql = sql & sMakesql14
    sql = sql & sMakesql15
    sql = sql & sMakesql16
    sql = sql & sMakesql17      '残存酸素仕様取得SQL追加　03/12/09 ooba
    sql = sql & sMakesql18      'GD-Den仕様取得SQL追加　05/01/27 ooba
    sql = sql & sMakesql19      'GD-DVD2仕様取得SQL追加　05/01/27 ooba
    sql = sql & sMakesql20      'GD-L/DL仕様取得SQL追加　05/01/27 ooba
    sql = sql & sMakesql21      'SPVNR仕様取得SQL追加　06/05/31 ooba
    
    sMakesql = sql
    
''    On Error GoTo db_Error
'    'SQL文の実行
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'
'    '該当データなし
'    If rs.EOF Then
'        funGetKouhoHinban1_5 = 1
'        GoTo db_Error
'    Else
'        sMakesql = sql
'    End If
'
'    Set rs = Nothing
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funGetKouhoHinban1_5 = -4
    GoTo Apl_Exit
    
db_Error:
    Set rs = Nothing
    If funGetKouhoHinban1_5 = 0 Then
        funGetKouhoHinban1_5 = -3
    End If
    GoTo Apl_Exit

End Function

'------------------------------------------------
' 先行評価項目仕様比較詳細SQL文作成
'------------------------------------------------

'概要      :指定されたﾁｪｯｸ内容詳細に基づき、該当する仕様値が一致している、または、マトリクスで一致している品番を抽出するSQL文を作成する。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型                 :説明
'          :sChkCode        ,I  ,String             :工程番号
'          :tbl_spec1_5_1() ,I  ,typ_Spec1_5_1      :ﾌﾞﾛｯｸID、又は、結晶番号
'          :sChkTable       ,I  ,String             :振替元品番
'          :sMakeSql        ,O  ,String             :作成SQL文
'          :戻り値          ,O  ,Integer            :取得の成否(0:正常取得, -1:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/10 新規作成　SB
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾁｪｯｸするﾃｰﾌﾞﾙが2つある場合に対応する為、引数にsChkTable2を追加
'Public Function funGetKouhoHinban1_5_1(sChkCode As String, tbl_spec1_5_1() As typ_Spec1_5_1, sChkTable As String, sMakesql As String) As Integer
'↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'AN温度振替チェック：引数にチェック項目を追加
'Public Function funGetKouhoHinban1_5_1(sChkCode As String, tbl_spec1_5_1() As typ_Spec1_5_1, sChkTable As String, sMakesql As String, Optional sChkTable2 As String = "") As Integer
Public Function funGetKouhoHinban1_5_1(sChkCode As String, tbl_spec1_5_1() As typ_Spec1_5_1, sChkTable As String, sMakesql As String, Optional sChkTable2 As String = "", Optional sChkCode2 As String = "") As Integer
'↑変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾁｪｯｸするﾃｰﾌﾞﾙが2つある場合に対応する為、引数にsChkTable2を追加
    Dim RET         As Integer      '戻り値
    Dim sql As String       'SQL全体
    Dim rs  As OraDynaset   'RecordSet
    Dim sinstr     As String       'ＳＱＬin句用文字列
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim iCnt        As Integer      '04/04/13 ooba
    Dim sNum        As String       '04/04/13 ooba
'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    Dim lsANCodeListWork()  As String      'Code一覧
    Dim lsANCodeList()  As String      'Code一覧
    Dim lsANCode        As String      'チェック種類
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------

    On Error GoTo Apl_down
    
    '戻り値初期化
    funGetKouhoHinban1_5_1 = 0
    '------------------------------------------ SQL文生成 ------------------------------------------------------
    'SQL文の作成
    sql = vbNullString
    '保証方法＿対象
    If Mid(sChkCode, 1, 1) = "2" Then
'        If tbl_spec1_5_1(0).HOSYOU = "H" Or tbl_spec1_5_1(0).HOSYOU = "S" Then
'            'マトリクス取得
'            sResult = ""
'            sinstr = ""
'            ret = funCodeDBGet("SB", "SH", tbl_spec1_5_1(0).HOSYOU, 0, " ", sResult)
'            If ret <> 0 Then
'                funGetKouhoHinban1_5_1 = 1
'                GoTo Apl_Exit
'            End If
'            ret = funinfo2(sResult, sinstr)
'            If ret <> 0 Then
'                funGetKouhoHinban1_5_1 = 1
'                GoTo Apl_Exit
'            End If
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↓追加 2003/10/21
'            ret = funCodeinGet("SB", "SH", sinstr, sResult)
'            If ret <> 0 Then
'                funGetKouhoHinban1_5_1 = 1
'                GoTo Apl_Exit
'            End If
'            sinstr = sResult
'    '        sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).HOSYOU1 & " IN  (SELECT NVL(TRIM(CODE),CHR(32)) FROM TBCMB005 WHERE SYSCLASS = 'SB' AND CLASS = 'SH' AND INFO2 in (" & sinstr & ")) " & vbCrLf
'            sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).HOSYOU1 & " IN  (" & sinstr & ") " & vbCrLf
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↑追加 2003/10/21
'        Else
'            sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).HOSYOU1 & " NOT IN  ('H', 'S') " & vbCrLf
'        End If
        If tbl_spec1_5_1(0).HOSYOU = "S" Then
            sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).HOSYOU1 & " NOT IN  ('H') " & vbCrLf
        ElseIf tbl_spec1_5_1(0).HOSYOU <> "H" And tbl_spec1_5_1(0).HOSYOU <> "S" Then
            sql = sql & " AND (" & sChkTable & "." & tbl_spec1_5_1(0).HOSYOU1 & " NOT IN  ('H', 'S') " & vbCrLf
            sql = sql & " OR " & sChkTable & "." & tbl_spec1_5_1(0).HOSYOU1 & " IS NULL)" & vbCrLf
        End If
    End If
    
    '------------------------------------------ 保証方法チェック ------------------------------------------------------
    If tbl_spec1_5_1(0).HOSYOU <> "H" And tbl_spec1_5_1(0).HOSYOU <> "S" Then GoTo Make_Exit
    
    '下限
    If Mid(sChkCode, 2, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).Min1 & " = " & tbl_spec1_5_1(0).Min & " " & vbCrLf
    End If
    '上限
    If Mid(sChkCode, 3, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).Max1 & " = " & tbl_spec1_5_1(0).max & " " & vbCrLf
    End If
    '測定位置＿方
    If Mid(sChkCode, 4, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).SOKU_HOU1 & " = '" & tbl_spec1_5_1(0).SOKU_HOU & "' " & vbCrLf
    End If
    '測定位置＿点
    If Mid(sChkCode, 5, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).SOKU_TEN1 & " = '" & tbl_spec1_5_1(0).SOKU_TEN & "' " & vbCrLf
    ElseIf Mid(sChkCode, 5, 1) = "2" Then   '08/01/29 ooba
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).SOKU_TEN1 & " <= '" & tbl_spec1_5_1(0).SOKU_TEN & "' " & vbCrLf
    End If
    '測定位置＿位
    If Mid(sChkCode, 6, 1) = "2" Then
        'マトリクス取得
        sResult = ""
        sinstr = ""
        RET = funCodeDBGet("SB", "OI", tbl_spec1_5_1(0).SOKU_ICHI, 0, " ", sResult)
        If RET <> 0 Then
            funGetKouhoHinban1_5_1 = 2
            GoTo Apl_Exit
        End If
        RET = funinfo2(sResult, sinstr)
        If RET <> 0 Then
            funGetKouhoHinban1_5_1 = 2
            GoTo Apl_Exit
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↓追加 2003/10/21
        RET = funCodeinGet("SB", "OI", sinstr, sResult)
        If RET <> 0 Then
            funGetKouhoHinban1_5_1 = 2
            GoTo Apl_Exit
        End If
        sinstr = sResult
'        sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).SOKU_ICHI1 & " IN  (SELECT NVL(TRIM(CODE),CHR(32)) FROM TBCMB005 WHERE SYSCLASS = 'SB' AND CLASS = 'OI' AND INFO2 in (" & sinstr & ")) " & vbCrLf
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).SOKU_ICHI1 & " IN  (" & sinstr & ") " & vbCrLf
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↑追加 2003/10/21
    End If
    
''Upd Start 2005/06/16 (TCS)T.Terauchi      SPV9点対応
    '測定位置＿位
    If Mid(sChkCode, 6, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).SOKU_ICHI1 & " = '" & tbl_spec1_5_1(0).SOKU_ICHI & "' " & vbCrLf
    End If
''Upd End   2005/06/16 (TCS)T.Terauchi      SPV9点対応
    
    '測定位置＿領
    If Mid(sChkCode, 7, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).SOKU_RYOU1 & " = '" & tbl_spec1_5_1(0).SOKU_RYOU & "' " & vbCrLf
    End If
    '検査有無
    If Mid(sChkCode, 8, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).UMU1 & " = '" & tbl_spec1_5_1(0).UMU & "' " & vbCrLf
    End If
    '熱処理法
    If Mid(sChkCode, 9, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).NETSU1 & " = '" & tbl_spec1_5_1(0).NETSU & "' " & vbCrLf
    End If
    '測定条件
    If Mid(sChkCode, 10, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).JOUKEN1 & " = '" & tbl_spec1_5_1(0).JOUKEN & "' " & vbCrLf
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    ElseIf Mid(sChkCode, 10, 1) = "2" Then
        If Trim(tbl_spec1_5_1(0).HWFGDSZY) = "F" Then
            sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).HWFGDSZY1 & " NOT IN ('G')" & vbCrLf
        End If
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    End If
    '選択ＥＴ代
    If Mid(sChkCode, 11, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).ET1 & " = " & tbl_spec1_5_1(0).ET & " " & vbCrLf
    End If
    '検査方法
    If Mid(sChkCode, 12, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).KENSA1 & " = '" & tbl_spec1_5_1(0).KENSA & "' " & vbCrLf
    End If
    'パターン区分
    If Mid(sChkCode, 13, 1) = "2" Then
        'マトリクス取得
        sResult = ""
        sinstr = ""
        If tbl_spec1_5_1(0).PATTERN <> " " Then
            RET = funCodeDBGet("SB", "OS", tbl_spec1_5_1(0).PATTERN, 0, " ", sResult)
        Else
            RET = funCodeDBGet("SB", "OS", "4", 0, " ", sResult)
        End If
        If RET <> 0 Then
            funGetKouhoHinban1_5_1 = 3
            GoTo Apl_Exit
        End If
        RET = funinfo2(sResult, sinstr)
        If RET <> 0 Then
            funGetKouhoHinban1_5_1 = 3
            GoTo Apl_Exit
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↓追加 2003/10/21
        RET = funCodeinGet("SB", "OS", sinstr, sResult)
        If RET <> 0 Then
            funGetKouhoHinban1_5_1 = 3
            GoTo Apl_Exit
        End If
        sinstr = sResult
'        sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).PATTERN1 & " IN  (SELECT NVL(TRIM(CODE),CHR(32)) FROM TBCMB005 WHERE SYSCLASS = 'SB' AND CLASS = 'OS' AND INFO2 in (" & sinstr & ")) " & vbCrLf
        If tbl_spec1_5_1(0).PATTERN = " " Then
            sql = sql & " AND (" & sChkTable & "." & tbl_spec1_5_1(0).PATTERN1 & " IN  (" & sinstr & ") " & vbCrLf
            sql = sql & " OR " & sChkTable & "." & tbl_spec1_5_1(0).PATTERN1 & " IS NULL)" & vbCrLf
        Else
            sql = sql & " AND " & sChkTable & "." & tbl_spec1_5_1(0).PATTERN1 & " IN  (" & sinstr & ") " & vbCrLf
        End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''↑追加 2003/10/21
    End If
    '検査頻度＿抜　04/04/13 ooba
    If Mid(sChkCode, 14, 1) = "2" Then
        'マトリクス取得
        sResult = ""
        sinstr = ""
        If tbl_spec1_5_1(0).KENH_NUKI = "3" Or tbl_spec1_5_1(0).KENH_NUKI = "4" _
                Or tbl_spec1_5_1(0).KENH_NUKI = "6" Then
            RET = funCodeDBGet("SB", "HO", tbl_spec1_5_1(0).KENH_NUKI, 0, " ", sResult)
        Else
            RET = funCodeDBGet("SB", "HO", "ETC", 0, " ", sResult)
        End If
        If RET <> 0 Then
            funGetKouhoHinban1_5_1 = 4
            GoTo Apl_Exit
        End If
        For iCnt = 1 To 3
            If iCnt = 1 Then sNum = "3"
            If iCnt = 2 Then sNum = "4"
            If iCnt = 3 Then sNum = "6"
            If Mid(sResult, iCnt, 1) = "1" Then
                If sinstr = "" Then sinstr = "'" & sNum & "'" Else sinstr = sinstr & ", '" & sNum & "'"
            End If
        Next
        If sinstr <> "" Then
            If Mid(sResult, 4, 1) = "1" Then sql = sql & " AND (" Else sql = sql & " AND "
            sql = sql & sChkTable & "." & tbl_spec1_5_1(0).KENH_NUKI1 & " IN  (" & sinstr & ") " & vbCrLf
        End If
        If Mid(sResult, 4, 1) = "1" Then
            If sinstr <> "" Then sql = sql & " OR " Else sql = sql & " AND "
            sql = sql & "(" & sChkTable & "." & tbl_spec1_5_1(0).KENH_NUKI1 & " IS NULL" & vbCrLf
            sql = sql & " OR " & sChkTable & "." & tbl_spec1_5_1(0).KENH_NUKI1 & " NOT IN ('3', '4', '6'))" & vbCrLf
            If sinstr <> "" Then sql = sql & ")" & vbCrLf Else sql = sql & vbCrLf
        End If
    End If
    
        
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加
    'ライン数
    If Mid(sChkCode, 15, 1) = "2" Then
        'マトリクス取得
        sResult = ""
        sinstr = ""
        
        RET = funCodeDBGet("SB", "LN", tbl_spec1_5_1(0).LINE, 0, " ", sResult)
        
        If RET <> 0 Then
            funGetKouhoHinban1_5_1 = 4
            GoTo Apl_Exit
        End If
        RET = funinfo2(sResult, sinstr)
        If RET <> 0 Then
            funGetKouhoHinban1_5_1 = 4
            GoTo Apl_Exit
        End If
                
        RET = funCodeinGet("SB", "LN", sinstr, sResult)
        If RET <> 0 Then
            funGetKouhoHinban1_5_1 = 4
            GoTo Apl_Exit
        End If
        sinstr = sResult
                
        If InStr(sinstr, "' '") > 0 Then
            'DBのﾗｲﾝ数ｶﾗﾑは整数型で，ｽﾍﾟｰｽを条件に入れるとｴﾗｰになるのでその対応
            If InStr(sinstr, ",' '") > 0 Then
                sinstr = Replace(sinstr, ",' '", "")
            ElseIf InStr(sinstr, "' ',") > 0 Then
                sinstr = Replace(sinstr, "' ',", "")
            Else
                sinstr = Replace(sinstr, "' '", "")
            End If
            sql = sql & " AND (" & sChkTable2 & "." & tbl_spec1_5_1(0).Line1 & " IS NULL" & vbCrLf
            sql = sql & " OR   " & sChkTable2 & "." & tbl_spec1_5_1(0).Line1 & " IN  (" & sinstr & "))" & vbCrLf
        Else
            sql = sql & " AND (" & sChkTable2 & "." & tbl_spec1_5_1(0).Line1 & " IN  (" & sinstr & ") " & vbCrLf
        End If
        
    End If
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数追加

'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'AN温度振替チェック
    'AN温度
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -s-
''    If Mid(sChkCode, 16, 1) = "1" Then
    If Mid(sChkCode, 16, 1) = "2" Then
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -e-
        ReDim lsANCodeListWork(0) As String
        ReDim lsANCodeList(0) As String
        'チェック内容により、使用マトリックスを変える
        Select Case sChkCode2
            Case "Rs"
                lsANCode = "AR"
            Case "Oi"
                lsANCode = "AO"
            Case "DS"               'DSOD追加　06/12/22 ooba
                lsANCode = "AD"
            Case "GD"               'GD追加　06/12/22 ooba
                lsANCode = "AG"
            Case Else
                lsANCode = "AE"
        End Select
        
        'AN温度マトリックスよりCodeの一覧を取得する
        RET = funCodeDBGetCodeList("SB", lsANCode, lsANCodeListWork)
        If RET < 0 Then
            funGetKouhoHinban1_5_1 = 4
            GoTo Apl_Exit
        End If
        
        For iCnt = 1 To UBound(lsANCodeListWork)
            RET = funCodeDBGetMatrixReturn("SB", lsANCode, lsANCodeListWork(iCnt), tbl_spec1_5_1(0).Antnp)
            If RET < 0 Then
                funGetKouhoHinban1_5_1 = 4
                GoTo Apl_Exit
            ElseIf RET = 0 Then
                ' AN温度チェックNGの値を保持する
                ReDim Preserve lsANCodeList(UBound(lsANCodeList) + 1) As String
                lsANCodeList(UBound(lsANCodeList)) = lsANCodeListWork(iCnt)
            End If
        Next iCnt
            
        'AN温度チェックNG以外のデータを取得する
        If UBound(lsANCodeList) <> 0 Then
            sql = sql & " AND (E025." & tbl_spec1_5_1(0).ANTNP1 & " NOT IN (" & vbCrLf
            For iCnt = 1 To UBound(lsANCodeList)
                If iCnt <> 1 Then
                    sql = sql & ","
                End If
                sql = sql & "'" & lsANCodeList(iCnt) & "'"
            Next iCnt
            sql = sql & "))" & vbCrLf
        End If
    
    End If
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------

    'PUA限　06/05/31 ooba
    If Mid(sChkCode, 17, 1) = "1" Then
        If tbl_spec1_5_1(0).SOKU_HOU & tbl_spec1_5_1(0).SOKU_TEN & tbl_spec1_5_1(0).SOKU_ICHI = "AMX" Then
            If tbl_spec1_5_1(0).PUAGEN <> -1 Then
                sql = sql & " AND " & sChkTable2 & "." & tbl_spec1_5_1(0).PUAGEN1 & " = " & tbl_spec1_5_1(0).PUAGEN & " " & vbCrLf
            Else
                sql = sql & " AND " & sChkTable2 & "." & tbl_spec1_5_1(0).PUAGEN1 & " IS NULL " & vbCrLf
            End If
        End If
    End If
    'PUA率　06/05/31 ooba
    If Mid(sChkCode, 18, 1) = "1" Then
        If tbl_spec1_5_1(0).SOKU_HOU & tbl_spec1_5_1(0).SOKU_TEN & tbl_spec1_5_1(0).SOKU_ICHI = "AMX" Then
            If tbl_spec1_5_1(0).PUAPER <> -1 Then
                sql = sql & " AND " & sChkTable2 & "." & tbl_spec1_5_1(0).PUAPER1 & " = " & tbl_spec1_5_1(0).PUAPER & " " & vbCrLf
            Else
                sql = sql & " AND " & sChkTable2 & "." & tbl_spec1_5_1(0).PUAPER1 & " IS NULL " & vbCrLf
            End If
        End If
    End If

Make_Exit:
    sMakesql = sql
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funGetKouhoHinban1_5_1 = -4
    GoTo Apl_Exit
    
End Function

'------------------------------------------------
' ナノトポ規格比較SQL文作成
'------------------------------------------------

'概要      :振替元品番が、ガラス接着品かどうかを判断する。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sBlockid        ,I  ,String       :ﾌﾞﾛｯｸID、又は、結晶番号
'          :sOld_Hinban     ,I  ,String       :振替元品番
'          :sMakeSql        ,O  ,String       :作成SQL文
'          :戻り値          ,O  ,Integer      :取得の成否(0:正常取得, -1:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/10 新規作成　SB

Public Function funGetKouhoHinban1_6(sProccd As String, sBlockId As String, sOld_Hinban As String, sMakesql As String) As Integer

    Dim sql As String       'SQL全体
    Dim rs  As OraDynaset   'RecordSet
    Dim w_i         As Long         'カウンタ
    Dim w_x         As Long         'カウンタ
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funGetKouhoHinban1_6 = 0
    
    '------------------------------------------ 振替元品種仕様データ取得 ------------------------------------------------------
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E026.HWFNP1AR,E026.HWFNP1MAX,E026.HWFNP2AR,E026.HWFNP2MAX,E018.HSXCSCEN " & vbCrLf
    sql = sql & "FROM   TBCME026 E026,TBCME018 E018 " & vbCrLf
    sql = sql & "WHERE  E026.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E026.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E026.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E026.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' AND " & vbCrLf
    sql = sql & "       E018.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E018.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' " & vbCrLf
'    sql = sql & "WHERE  E026.HINBAN || TO_CHAR(E026.MNOREVNO, 'FM00') || E026.FACTORY || E026.OPECOND   =   '" & sOld_Hinban & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        funGetKouhoHinban1_6 = 16001
        GoTo db_Error
    End If
    
    '取得データセット
    Erase tbl_spec1_6
    With tbl_spec1_6(0)
        If IsNull(rs("HWFNP1AR")) = False Then .HWFNP1AR = rs("HWFNP1AR") Else .HWFNP1AR = 0            '品WFナノトポ１エリア
        If IsNull(rs("HWFNP1MAX")) = False Then .HWFNP1MAX = rs("HWFNP1MAX") Else .HWFNP1MAX = 0        '品WFナノトポ１上限
        If IsNull(rs("HWFNP2AR")) = False Then .HWFNP2AR = rs("HWFNP2AR") Else .HWFNP2AR = 0            '品WFナノトポ２エリア
        If IsNull(rs("HWFNP2MAX")) = False Then .HWFNP2MAX = rs("HWFNP2MAX") Else .HWFNP2MAX = 0        '品WFナノトポ２上限
        If IsNull(rs("HSXCSCEN")) = False Then .HSXCSCEN = rs("HSXCSCEN") Else .HSXCSCEN = -1           '結晶面傾き中心
    End With
    'Double
    Set rs = Nothing
    On Error GoTo Apl_down
    '------------------------------------------ 振替元品種と同一仕様の品番を取得 ------------------------------------------------------
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT 'X' " & vbCrLf
'    sql = sql & "SELECT E026.HINBAN || TO_CHAR(E026.MNOREVNO, 'FM00') || E026.FACTORY | |E026.OPECOND HINBAN " & vbCrLf
    sql = sql & "FROM   TBCME026 E026,TBCME018 E018 " & vbCrLf
    sql = sql & "WHERE  E018A.HINBAN                    = E018.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018A.MNOREVNO, 'FM00') = TO_CHAR(E018.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E018A.FACTORY                   = E018.FACTORY                      AND " & vbCrLf
    sql = sql & "       E018A.OPECOND                   = E018.OPECOND                      AND " & vbCrLf
    sql = sql & "       E018A.HINBAN                    = E026.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018A.MNOREVNO, 'FM00') = TO_CHAR(E026.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E018A.FACTORY                   = E026.FACTORY                      AND " & vbCrLf
    sql = sql & "       E018A.OPECOND                   = E026.OPECOND                      AND " & vbCrLf
    If tbl_spec1_6(0).HSXCSCEN = -1 Then
        sql = sql & "       (E018.HSXCSCEN is null OR E018.HSXCSCEN = 0) " & vbCrLf
    Else
        sql = sql & "       E018.HSXCSCEN                   =  " & tbl_spec1_6(0).HSXCSCEN & " " & vbCrLf
    End If
    If tbl_spec1_6(0).HWFNP1AR = 2 And tbl_spec1_6(0).HWFNP1MAX <= 17 Or _
       tbl_spec1_6(0).HWFNP2AR = 10 And tbl_spec1_6(0).HWFNP2MAX <= 50 Then
        'ガラス接着品
'        sql = sql & "WHERE  E026.HINBAN || TO_CHAR(E026.MNOREVNO, 'FM00') || E026.FACTORY || E026.OPECOND <> '" & sOld_Hinban & "' " & vbCrLf
    Else
'        sql = sql & "WHERE  E026.HINBAN || TO_CHAR(E026.MNOREVNO, 'FM00') || E026.FACTORY || E026.OPECOND <> '" & sOld_Hinban & "' AND " & vbCrLf
        sql = sql & "AND    ((E026.HWFNP1AR <> 2    OR  " & vbCrLf
        sql = sql & "         E026.HWFNP1MAX > 17)  OR  " & vbCrLf
        sql = sql & "        (E026.HWFNP2AR <> 10   OR  " & vbCrLf
        sql = sql & "         E026.HWFNP2MAX > 50))     " & vbCrLf
    End If
    
    sMakesql = sql
    
'    On Error GoTo db_Error
'    'SQL文の実行
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'
'    '該当データなし
'    If rs.EOF Then
'        funGetKouhoHinban1_6 = 1
'        GoTo db_Error
'    Else
'        sMakesql = sql
'    End If
'
'    Set rs = Nothing
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funGetKouhoHinban1_6 = -4
    GoTo Apl_Exit
    
db_Error:
    Set rs = Nothing
    If funGetKouhoHinban1_6 = 0 Then
        funGetKouhoHinban1_6 = -3
    End If
    GoTo Apl_Exit

End Function

'------------------------------------------------
' コードＤＢよりＩＮ句取得
'------------------------------------------------

'概要      :指定された項目をキーに、コードマスター(TBCMB005)から該当するデータを取得する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :sSysclass     ,I  ,String       :ｼｽﾃﾑ区分('SB'固定)
'          :sClass        ,I  ,String       :区分
'          :sinstr        ,I  ,String       :INFO2
'          :sResult       ,O  ,String       :取得ﾃﾞｰﾀ
'          :戻り値        ,O  ,Integer      :取得の成否(0:正常取得, -1:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/04 新規作成　システムブレイン

Public Function funCodeinGet(sSysclass As String, sClass As String, sinstr As String, sResult As String) As Integer


    Dim sql As String       'SQL全体
    Dim rs  As OraDynaset   'RecordSet
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funCodeinGet = 0
    sResult = ""
    '------------------------------------------ 振替元品種仕様データ取得 ------------------------------------------------------
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT NVL(TRIM(CODE),CHR(32))  WCODE      " & vbCrLf
    sql = sql & "FROM   TBCMB005                            " & vbCrLf
    sql = sql & "WHERE  SYSCLASS = '" & sSysclass & "' AND  " & vbCrLf
    sql = sql & "       CLASS    = '" & sClass & "'    AND  " & vbCrLf
    sql = sql & "       INFO2 in (" & sinstr & ")           " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Debug.Print sql
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        GoTo db_Error
    End If
    With rs
        .dbMoveFirst
        Do Until .EOF
            If sResult = "" Then
                sResult = "'" & .Fields(0).Value & "'"
            Else
                sResult = sResult & ",'" & .Fields(0).Value & "'"
            End If
            .DbMoveNext
        Loop
    End With
    
    Set rs = Nothing
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funCodeinGet = -4
    GoTo Apl_Exit
    
db_Error:
    Set rs = Nothing
    If funCodeinGet = 0 Then
        funCodeinGet = -3
    End If
    GoTo Apl_Exit

End Function

'------------------------------------------------
' ﾊﾟﾀｰﾝ1_4,1_5のｴﾗｰ文字列取得
'------------------------------------------------

'概要      :指定されたｴﾗｰｺｰﾄﾞに該当するｴﾗｰ文字列を返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :iErr_Code     ,I  ,Integer      :ｴﾗｰｺｰﾄﾞ
'          :戻り値        ,O  ,String       :ｴﾗｰ文字列
'説明      :
'履歴      :2003/10/26 新規作成　システムブレイン

Private Function funErrMsgGet(iErr_Code As Integer) As String
    
    '戻り値初期化
    funErrMsgGet = ""
    
    If iErr_Code = 1 Then
        funErrMsgGet = "保証方法_対象"
    ElseIf iErr_Code = 2 Then
        funErrMsgGet = "測定位置_位"
    ElseIf iErr_Code = 3 Then
        funErrMsgGet = "ﾊﾟﾀｰﾝ区分"
'*** UPDATE ↓ Y.SIMIZU 2005/10/24 ﾗｲﾝ数のｴﾗｰﾒｯｾｰｼﾞ追加
    ElseIf iErr_Code = 4 Then
        funErrMsgGet = "ﾗｲﾝ数"
'*** UPDATE ↑ Y.SIMIZU 2005/10/24 ﾗｲﾝ数のｴﾗｰﾒｯｾｰｼﾞ追加
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    ElseIf iErr_Code = 5 Then
        funErrMsgGet = "検査頻度_抜"
    ElseIf iErr_Code = 6 Then
        funErrMsgGet = "AN温度"
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    End If

End Function


'------------------------------------------------
' エピ先行評価項目仕様比較SQL文作成
'------------------------------------------------
'概要      :指定された振替元品番とエピ先行評価項目仕様値が一致している、
'           または、マトリクスで一致している品番を抽出するSQL文を作成する。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sBlockid        ,I  ,String       :ﾌﾞﾛｯｸID、又は、結晶番号
'          :sOld_Hinban     ,I  ,String       :振替元品番
'          :sMakeSql        ,O  ,String       :作成SQL文
'          :戻り値          ,O  ,Integer      :取得の成否(0:正常取得, -1:取得ｴﾗｰ)
'説明      :funGetKouhoHinban1_5(先行評価項目仕様比較SQL文作成)を元に作成
'履歴      :2006/08/15 新規作成 エピ先行評価追加対応 SMP)kondoh

Public Function funGetKouhoHinban1_9(sProccd As String, sBlockId As String, sOld_Hinban As String, sMakesql As String) As Integer



    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql As String               'SQL全体
    Dim rs  As OraDynaset           'RecordSet
    Dim sMakesql1   As String       '呼び出しファンクションSQL作成
    Dim sMakesql2   As String       '呼び出しファンクションSQL作成
    Dim sMakesql3   As String       '呼び出しファンクションSQL作成
    Dim sMakesql4   As String       '呼び出しファンクションSQL作成
    Dim sMakesql5   As String       '呼び出しファンクションSQL作成
    Dim sMakesql6   As String       '呼び出しファンクションSQL作成

    On Error GoTo Apl_down
    
    '戻り値初期化
    funGetKouhoHinban1_9 = 0
    
    '------------------------------------------ 振替元品種仕様データ取得 ------------------------------------------------------
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E050.HEPOF1HS,E050.HEPOF1SH,E050.HEPOF1ST,E050.HEPOF1SR,E050.HEPOF1NS,E050.HEPOF1SZ,E050.HEPOF1ET,E050.HEPOSF1PTK,E050.HEPOF1KN,   " & vbCrLf
    sql = sql & "       E050.HEPOF2HS,E050.HEPOF2SH,E050.HEPOF2ST,E050.HEPOF2SR,E050.HEPOF2NS,E050.HEPOF2SZ,E050.HEPOF2ET,E050.HEPOSF2PTK,E050.HEPOF2KN,   " & vbCrLf
    sql = sql & "       E050.HEPOF3HS,E050.HEPOF3SH,E050.HEPOF3ST,E050.HEPOF3SR,E050.HEPOF3NS,E050.HEPOF3SZ,E050.HEPOF3ET,E050.HEPOSF3PTK,E050.HEPOF3KN,   " & vbCrLf
    sql = sql & "       E050.HEPBM1HS,E050.HEPBM1SH,E050.HEPBM1ST,E050.HEPBM1SR,E050.HEPBM1NS,E050.HEPBM1SZ,E050.HEPBM1ET,E050.HEPBM1KN,   " & vbCrLf
    sql = sql & "       E050.HEPBM2HS,E050.HEPBM2SH,E050.HEPBM2ST,E050.HEPBM2SR,E050.HEPBM2NS,E050.HEPBM2SZ,E050.HEPBM2ET,E050.HEPBM2KN,   " & vbCrLf
    sql = sql & "       E050.HEPBM3HS,E050.HEPBM3SH,E050.HEPBM3ST,E050.HEPBM3SR,E050.HEPBM3NS,E050.HEPBM3SZ,E050.HEPBM3ET,E050.HEPBM3KN,   " & vbCrLf
    sql = sql & "       E050.HEPANTNP,E050.HEPACEN " & vbCrLf
    sql = sql & "FROM   TBCME050 E050 " & vbCrLf
    sql = sql & "WHERE  E050.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E050.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E050.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E050.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' " & vbCrLf
   
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        funGetKouhoHinban1_9 = 19001
        GoTo db_Error
    End If
    
    '取得データセット
    '検査頻度_抜ﾃﾞｰﾀ追加　04/04/13 ooba
    Erase tbl_spec1_9
    With tbl_spec1_9(0)
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
    sMakesql1 = ""
    sMakesql2 = ""
    sMakesql3 = ""
    sMakesql4 = ""
    sMakesql5 = ""
    sMakesql6 = ""

    'OSF1E
    sResult = ""
    RET = funCodeDBGet("SB", "19", "O1E", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_9 = 19010
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_9_1
    tbl_spec1_9_1(0).HOSYOU = tbl_spec1_9(0).HEPOF1HS
    tbl_spec1_9_1(0).HOSYOU1 = "HEPOF1HS"
    tbl_spec1_9_1(0).SOKU_HOU = tbl_spec1_9(0).HEPOF1SH
    tbl_spec1_9_1(0).SOKU_HOU1 = "HEPOF1SH"
    tbl_spec1_9_1(0).SOKU_TEN = tbl_spec1_9(0).HEPOF1ST
    tbl_spec1_9_1(0).SOKU_TEN1 = "HEPOF1ST"
    tbl_spec1_9_1(0).SOKU_RYOU = tbl_spec1_9(0).HEPOF1SR
    tbl_spec1_9_1(0).SOKU_RYOU1 = "HEPOF1SR"
    tbl_spec1_9_1(0).NETSU = tbl_spec1_9(0).HEPOF1NS
    tbl_spec1_9_1(0).NETSU1 = "HEPOF1NS"
    tbl_spec1_9_1(0).JOUKEN = tbl_spec1_9(0).HEPOF1SZ
    tbl_spec1_9_1(0).JOUKEN1 = "HEPOF1SZ"
    tbl_spec1_9_1(0).ET = tbl_spec1_9(0).HEPOF1ET
    tbl_spec1_9_1(0).ET1 = "HEPOF1ET"
    tbl_spec1_9_1(0).PATTERN = tbl_spec1_9(0).HEPOSF1PTK
    tbl_spec1_9_1(0).PATTERN1 = "HEPOSF1PTK"
    tbl_spec1_9_1(0).KENH_NUKI = tbl_spec1_9(0).HEPOF1KN
    tbl_spec1_9_1(0).KENH_NUKI1 = "HEPOF1KN"
    tbl_spec1_9_1(0).Antnp = tbl_spec1_9(0).HEPANTNP
    tbl_spec1_9_1(0).ANTNP1 = "HEPANTNP"
    tbl_spec1_9_1(0).EPATU = tbl_spec1_9(0).HEPACEN
    tbl_spec1_9_1(0).EPATU1 = "HEPACEN"
    RET = funGetKouhoHinban1_9_1(sResult, tbl_spec1_9_1(), "E050", sMakesql)
    If RET <> 0 Then
        funGetKouhoHinban1_9 = 19010 + RET
        GoTo Apl_Exit
    End If
    sMakesql1 = sMakesql
    
    'OSF2E
    sResult = ""
    RET = funCodeDBGet("SB", "19", "O2E", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_9 = 19020
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_9_1
    tbl_spec1_9_1(0).HOSYOU = tbl_spec1_9(0).HEPOF2HS
    tbl_spec1_9_1(0).HOSYOU1 = "HEPOF2HS"
    tbl_spec1_9_1(0).SOKU_HOU = tbl_spec1_9(0).HEPOF2SH
    tbl_spec1_9_1(0).SOKU_HOU1 = "HEPOF2SH"
    tbl_spec1_9_1(0).SOKU_TEN = tbl_spec1_9(0).HEPOF2ST
    tbl_spec1_9_1(0).SOKU_TEN1 = "HEPOF2ST"
    tbl_spec1_9_1(0).SOKU_RYOU = tbl_spec1_9(0).HEPOF2SR
    tbl_spec1_9_1(0).SOKU_RYOU1 = "HEPOF2SR"
    tbl_spec1_9_1(0).NETSU = tbl_spec1_9(0).HEPOF2NS
    tbl_spec1_9_1(0).NETSU1 = "HEPOF2NS"
    tbl_spec1_9_1(0).JOUKEN = tbl_spec1_9(0).HEPOF2SZ
    tbl_spec1_9_1(0).JOUKEN1 = "HEPOF2SZ"
    tbl_spec1_9_1(0).ET = tbl_spec1_9(0).HEPOF2ET
    tbl_spec1_9_1(0).ET1 = "HEPOF2ET"
    tbl_spec1_9_1(0).PATTERN = tbl_spec1_9(0).HEPOSF2PTK
    tbl_spec1_9_1(0).PATTERN1 = "HEPOSF2PTK"
    tbl_spec1_9_1(0).KENH_NUKI = tbl_spec1_9(0).HEPOF2KN
    tbl_spec1_9_1(0).KENH_NUKI1 = "HEPOF2KN"
    tbl_spec1_9_1(0).Antnp = tbl_spec1_9(0).HEPANTNP
    tbl_spec1_9_1(0).ANTNP1 = "HEPANTNP"
    tbl_spec1_9_1(0).EPATU = tbl_spec1_9(0).HEPACEN
    tbl_spec1_9_1(0).EPATU1 = "HEPACEN"
    RET = funGetKouhoHinban1_9_1(sResult, tbl_spec1_9_1(), "E050", sMakesql)
    If RET <> 0 Then
        funGetKouhoHinban1_9 = 19020 + RET
        GoTo Apl_Exit
    End If
    sMakesql2 = sMakesql
    
    'OSF3E
    sResult = ""
    RET = funCodeDBGet("SB", "19", "O3E", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_9 = 19030
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_9_1
    tbl_spec1_9_1(0).HOSYOU = tbl_spec1_9(0).HEPOF3HS
    tbl_spec1_9_1(0).HOSYOU1 = "HEPOF3HS"
    tbl_spec1_9_1(0).SOKU_HOU = tbl_spec1_9(0).HEPOF3SH
    tbl_spec1_9_1(0).SOKU_HOU1 = "HEPOF3SH"
    tbl_spec1_9_1(0).SOKU_TEN = tbl_spec1_9(0).HEPOF3ST
    tbl_spec1_9_1(0).SOKU_TEN1 = "HEPOF3ST"
    tbl_spec1_9_1(0).SOKU_RYOU = tbl_spec1_9(0).HEPOF3SR
    tbl_spec1_9_1(0).SOKU_RYOU1 = "HEPOF3SR"
    tbl_spec1_9_1(0).NETSU = tbl_spec1_9(0).HEPOF3NS
    tbl_spec1_9_1(0).NETSU1 = "HEPOF3NS"
    tbl_spec1_9_1(0).JOUKEN = tbl_spec1_9(0).HEPOF3SZ
    tbl_spec1_9_1(0).JOUKEN1 = "HEPOF3SZ"
    tbl_spec1_9_1(0).ET = tbl_spec1_9(0).HEPOF3ET
    tbl_spec1_9_1(0).ET1 = "HEPOF3ET"
    tbl_spec1_9_1(0).PATTERN = tbl_spec1_9(0).HEPOSF3PTK
    tbl_spec1_9_1(0).PATTERN1 = "HEPOSF3PTK"
    tbl_spec1_9_1(0).KENH_NUKI = tbl_spec1_9(0).HEPOF3KN
    tbl_spec1_9_1(0).KENH_NUKI1 = "HEPOF3KN"
    tbl_spec1_9_1(0).Antnp = tbl_spec1_9(0).HEPANTNP
    tbl_spec1_9_1(0).ANTNP1 = "HEPANTNP"
    tbl_spec1_9_1(0).EPATU = tbl_spec1_9(0).HEPACEN
    tbl_spec1_9_1(0).EPATU1 = "HEPACEN"
    RET = funGetKouhoHinban1_9_1(sResult, tbl_spec1_9_1(), "E050", sMakesql)
    If RET <> 0 Then
        funGetKouhoHinban1_9 = 19030 + RET
        GoTo Apl_Exit
    End If
    sMakesql3 = sMakesql
    
    'BMD1E
    sResult = ""
    RET = funCodeDBGet("SB", "19", "B1E", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_9 = 19040
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_9_1
    tbl_spec1_9_1(0).HOSYOU = tbl_spec1_9(0).HEPBM1HS
    tbl_spec1_9_1(0).HOSYOU1 = "HEPBM1HS"
    tbl_spec1_9_1(0).SOKU_HOU = tbl_spec1_9(0).HEPBM1SH
    tbl_spec1_9_1(0).SOKU_HOU1 = "HEPBM1SH"
    tbl_spec1_9_1(0).SOKU_TEN = tbl_spec1_9(0).HEPBM1ST
    tbl_spec1_9_1(0).SOKU_TEN1 = "HEPBM1ST"
    tbl_spec1_9_1(0).SOKU_RYOU = tbl_spec1_9(0).HEPBM1SR
    tbl_spec1_9_1(0).SOKU_RYOU1 = "HEPBM1SR"
    tbl_spec1_9_1(0).NETSU = tbl_spec1_9(0).HEPBM1NS
    tbl_spec1_9_1(0).NETSU1 = "HEPBM1NS"
    tbl_spec1_9_1(0).JOUKEN = tbl_spec1_9(0).HEPBM1SZ
    tbl_spec1_9_1(0).JOUKEN1 = "HEPBM1SZ"
    tbl_spec1_9_1(0).ET = tbl_spec1_9(0).HEPBM1ET
    tbl_spec1_9_1(0).ET1 = "HEPBM1ET"
    tbl_spec1_9_1(0).KENH_NUKI = tbl_spec1_9(0).HEPBM1KN
    tbl_spec1_9_1(0).KENH_NUKI1 = "HEPBM1KN"
    tbl_spec1_9_1(0).Antnp = tbl_spec1_9(0).HEPANTNP
    tbl_spec1_9_1(0).ANTNP1 = "HEPANTNP"
    tbl_spec1_9_1(0).EPATU = tbl_spec1_9(0).HEPACEN
    tbl_spec1_9_1(0).EPATU1 = "HEPACEN"
    RET = funGetKouhoHinban1_9_1(sResult, tbl_spec1_9_1(), "E050", sMakesql)
    If RET <> 0 Then
        funGetKouhoHinban1_9 = 19040 + RET
        GoTo Apl_Exit
    End If
    sMakesql4 = sMakesql
    
    'BMD2E
    sResult = ""
    RET = funCodeDBGet("SB", "19", "B2E", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_9 = 19050
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_9_1
    tbl_spec1_9_1(0).HOSYOU = tbl_spec1_9(0).HEPBM2HS
    tbl_spec1_9_1(0).HOSYOU1 = "HEPBM2HS"
    tbl_spec1_9_1(0).SOKU_HOU = tbl_spec1_9(0).HEPBM2SH
    tbl_spec1_9_1(0).SOKU_HOU1 = "HEPBM2SH"
    tbl_spec1_9_1(0).SOKU_TEN = tbl_spec1_9(0).HEPBM2ST
    tbl_spec1_9_1(0).SOKU_TEN1 = "HEPBM2ST"
    tbl_spec1_9_1(0).SOKU_RYOU = tbl_spec1_9(0).HEPBM2SR
    tbl_spec1_9_1(0).SOKU_RYOU1 = "HEPBM2SR"
    tbl_spec1_9_1(0).NETSU = tbl_spec1_9(0).HEPBM2NS
    tbl_spec1_9_1(0).NETSU1 = "HEPBM2NS"
    tbl_spec1_9_1(0).JOUKEN = tbl_spec1_9(0).HEPBM2SZ
    tbl_spec1_9_1(0).JOUKEN1 = "HEPBM2SZ"
    tbl_spec1_9_1(0).ET = tbl_spec1_9(0).HEPBM2ET
    tbl_spec1_9_1(0).ET1 = "HEPBM2ET"
    tbl_spec1_9_1(0).KENH_NUKI = tbl_spec1_9(0).HEPBM2KN
    tbl_spec1_9_1(0).KENH_NUKI1 = "HEPBM2KN"
    tbl_spec1_9_1(0).Antnp = tbl_spec1_9(0).HEPANTNP
    tbl_spec1_9_1(0).ANTNP1 = "HEPANTNP"
    tbl_spec1_9_1(0).EPATU = tbl_spec1_9(0).HEPACEN
    tbl_spec1_9_1(0).EPATU1 = "HEPACEN"
    RET = funGetKouhoHinban1_9_1(sResult, tbl_spec1_9_1(), "E050", sMakesql)
    If RET <> 0 Then
        funGetKouhoHinban1_9 = 19050 + RET
        GoTo Apl_Exit
    End If
    sMakesql5 = sMakesql

    'BMD3E
    sResult = ""
    RET = funCodeDBGet("SB", "19", "B3E", 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_9 = 19060
        GoTo Apl_Exit
    End If
    sMakesql = ""
    Erase tbl_spec1_9_1
    tbl_spec1_9_1(0).HOSYOU = tbl_spec1_9(0).HEPBM3HS
    tbl_spec1_9_1(0).HOSYOU1 = "HEPBM3HS"
    tbl_spec1_9_1(0).SOKU_HOU = tbl_spec1_9(0).HEPBM3SH
    tbl_spec1_9_1(0).SOKU_HOU1 = "HEPBM3SH"
    tbl_spec1_9_1(0).SOKU_TEN = tbl_spec1_9(0).HEPBM3ST
    tbl_spec1_9_1(0).SOKU_TEN1 = "HEPBM3ST"
    tbl_spec1_9_1(0).SOKU_RYOU = tbl_spec1_9(0).HEPBM3SR
    tbl_spec1_9_1(0).SOKU_RYOU1 = "HEPBM3SR"
    tbl_spec1_9_1(0).NETSU = tbl_spec1_9(0).HEPBM3NS
    tbl_spec1_9_1(0).NETSU1 = "HEPBM3NS"
    tbl_spec1_9_1(0).JOUKEN = tbl_spec1_9(0).HEPBM3SZ
    tbl_spec1_9_1(0).JOUKEN1 = "HEPBM3SZ"
    tbl_spec1_9_1(0).ET = tbl_spec1_9(0).HEPBM3ET
    tbl_spec1_9_1(0).ET1 = "HEPBM3ET"
    tbl_spec1_9_1(0).KENH_NUKI = tbl_spec1_9(0).HEPBM3KN
    tbl_spec1_9_1(0).KENH_NUKI1 = "HEPBM3KN"
    tbl_spec1_9_1(0).Antnp = tbl_spec1_9(0).HEPANTNP
    tbl_spec1_9_1(0).ANTNP1 = "HEPANTNP"
    tbl_spec1_9_1(0).EPATU = tbl_spec1_9(0).HEPACEN
    tbl_spec1_9_1(0).EPATU1 = "HEPACEN"
    RET = funGetKouhoHinban1_9_1(sResult, tbl_spec1_9_1(), "E050", sMakesql)
    If RET <> 0 Then
        funGetKouhoHinban1_9 = 19060 + RET
        GoTo Apl_Exit
    End If
    sMakesql6 = sMakesql

    
    '------------------------------------------ 振替元品種と同一仕様の品番を取得 ------------------------------------------------------
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT 'X' " & vbCrLf
    sql = sql & "FROM   TBCME050 E050 " & vbCrLf
    sql = sql & "WHERE  E018A.HINBAN                    = E050.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018A.MNOREVNO, 'FM00') = TO_CHAR(E050.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E018A.FACTORY                   = E050.FACTORY                      AND " & vbCrLf
    sql = sql & "       E018A.OPECOND                   = E050.OPECOND                      " & vbCrLf
    sql = sql & sMakesql1
    sql = sql & sMakesql2
    sql = sql & sMakesql3
    sql = sql & sMakesql4
    sql = sql & sMakesql5
    sql = sql & sMakesql6
    
    sMakesql = sql
        
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funGetKouhoHinban1_9 = -4
    GoTo Apl_Exit
    
db_Error:
    Set rs = Nothing
    If funGetKouhoHinban1_9 = 0 Then
        funGetKouhoHinban1_9 = -3
    End If
    GoTo Apl_Exit

End Function

'------------------------------------------------
' エピ先行評価項目仕様比較詳細SQL文作成
'------------------------------------------------
'概要      :指定されたﾁｪｯｸ内容詳細に基づき、該当する仕様値が一致している、
'           または、マトリクスで一致している品番を抽出するSQL文を作成する。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型                 :説明
'          :sChkCode        ,I  ,String             :工程番号
'          :tbl_spec1_9_1() ,I  ,typ_Spec1_9_1      :ﾌﾞﾛｯｸID、又は、結晶番号
'          :sChkTable       ,I  ,String             :振替元品番
'          :sMakeSql        ,O  ,String             :作成SQL文
'          :戻り値          ,O  ,Integer            :取得の成否(0:正常取得, -1:取得ｴﾗｰ)
'説明      :
'履歴      :2006/08/15 新規作成 エピ先行評価追加対応 SMP)kondoh
Public Function funGetKouhoHinban1_9_1(sChkCode As String, tbl_spec1_9_1() As typ_Spec1_9_1, _
                                        sChkTable As String, sMakesql As String) As Integer
    Dim RET         As Integer          '戻り値
    Dim sql         As String           'SQL全体
    Dim rs          As OraDynaset       'RecordSet
    Dim sinstr      As String           'ＳＱＬin句用文字列
    Dim sResult     As String           'コードＤＢ取得関数の取得変数
    Dim iCnt        As Integer
    Dim sNum        As String
    Dim lsANCodeListWork()  As String   'Code一覧
    Dim lsANCodeList()  As String       'Code一覧
    Dim lsANCode        As String       'チェック種類

    On Error GoTo Apl_down
    
    '戻り値初期化
    funGetKouhoHinban1_9_1 = 0
    
    'SQL文の作成
    sql = vbNullString
    '保証方法＿処
    If Mid(sChkCode, 1, 1) = "2" Then
        If tbl_spec1_9_1(0).HOSYOU = "S" Then
            sql = sql & " AND " & sChkTable & "." & tbl_spec1_9_1(0).HOSYOU1 & " NOT IN  ('H') " & vbCrLf
        ElseIf tbl_spec1_9_1(0).HOSYOU <> "H" And tbl_spec1_9_1(0).HOSYOU <> "S" Then
            sql = sql & " AND (" & sChkTable & "." & tbl_spec1_9_1(0).HOSYOU1 & " NOT IN  ('H', 'S') " & vbCrLf
            sql = sql & " OR " & sChkTable & "." & tbl_spec1_9_1(0).HOSYOU1 & " IS NULL)" & vbCrLf
        End If
    End If
    
    If tbl_spec1_9_1(0).HOSYOU <> "H" And tbl_spec1_9_1(0).HOSYOU <> "S" Then GoTo Make_Exit
    
''    '下限
''    If Mid(sChkCode, 2, 1) = "1" Then
''        sql = sql & " AND " & sChkTable & "." & tbl_spec1_9_1(0).MIN_LIMIT1 & " = " & tbl_spec1_9_1(0).MIN_LIMIT & " " & vbCrLf
''    End If
''    '上限
''    If Mid(sChkCode, 3, 1) = "1" Then
''        sql = sql & " AND " & sChkTable & "." & tbl_spec1_9_1(0).MAX_LIMIT1 & " = " & tbl_spec1_9_1(0).MAX_LIMIT & " " & vbCrLf
''    End If
    '測定位置＿方
    If Mid(sChkCode, 4, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_9_1(0).SOKU_HOU1 & " = '" & tbl_spec1_9_1(0).SOKU_HOU & "' " & vbCrLf
    End If
    '測定位置＿点
    If Mid(sChkCode, 5, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_9_1(0).SOKU_TEN1 & " = '" & tbl_spec1_9_1(0).SOKU_TEN & "' " & vbCrLf
    End If
''    '測定位置＿位
''    If Mid(sChkCode, 6, 1) = "2" Then
''        'マトリクス取得
''        sResult = ""
''        sinstr = ""
''        RET = funCodeDBGet("SB", "OI", tbl_spec1_9_1(0).SOKU_ICHI, 0, " ", sResult)
''        If RET <> 0 Then
''            funGetKouhoHinban1_9_1 = 2
''            GoTo Apl_Exit
''        End If
''        RET = funinfo2(sResult, sinstr)
''        If RET <> 0 Then
''            funGetKouhoHinban1_9_1 = 2
''            GoTo Apl_Exit
''        End If
''        RET = funCodeinGet("SB", "OI", sinstr, sResult)
''        If RET <> 0 Then
''            funGetKouhoHinban1_9_1 = 2
''            GoTo Apl_Exit
''        End If
''        sinstr = sResult
''        sql = sql & " AND " & sChkTable & "." & tbl_spec1_9_1(0).SOKU_ICHI1 & " IN  (" & sinstr & ") " & vbCrLf
''    End If
''    '測定位置＿位
''    If Mid(sChkCode, 6, 1) = "1" Then
''        sql = sql & " AND " & sChkTable & "." & tbl_spec1_9_1(0).SOKU_ICHI1 & " = '" & tbl_spec1_9_1(0).SOKU_ICHI & "' " & vbCrLf
''    End If
    '測定位置＿領
    If Mid(sChkCode, 7, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_9_1(0).SOKU_RYOU1 & " = '" & tbl_spec1_9_1(0).SOKU_RYOU & "' " & vbCrLf
    End If
''    '検査有無
''    If Mid(sChkCode, 8, 1) = "1" Then
''        sql = sql & " AND " & sChkTable & "." & tbl_spec1_9_1(0).UMU1 & " = '" & tbl_spec1_9_1(0).UMU & "' " & vbCrLf
''    End If
    '熱処理法
    If Mid(sChkCode, 9, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_9_1(0).NETSU1 & " = '" & tbl_spec1_9_1(0).NETSU & "' " & vbCrLf
    End If
    '測定条件
    If Mid(sChkCode, 10, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_9_1(0).JOUKEN1 & " = '" & tbl_spec1_9_1(0).JOUKEN & "' " & vbCrLf
    End If
    '選択ＥＴ代
    If Mid(sChkCode, 11, 1) = "1" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_9_1(0).ET1 & " = " & tbl_spec1_9_1(0).ET & " " & vbCrLf
    End If
''    '検査方法
''    If Mid(sChkCode, 12, 1) = "1" Then
''        sql = sql & " AND " & sChkTable & "." & tbl_spec1_9_1(0).KENSA1 & " = '" & tbl_spec1_9_1(0).KENSA & "' " & vbCrLf
''    End If
    'パターン区分
    If Mid(sChkCode, 13, 1) = "2" Then
        'マトリクス取得
        sResult = ""
        sinstr = ""
        If tbl_spec1_9_1(0).PATTERN <> " " Then
            RET = funCodeDBGet("SB", "OS", tbl_spec1_9_1(0).PATTERN, 0, " ", sResult)
        Else
            RET = funCodeDBGet("SB", "OS", "4", 0, " ", sResult)
        End If
        If RET <> 0 Then
            funGetKouhoHinban1_9_1 = 3
            GoTo Apl_Exit
        End If
        RET = funinfo2(sResult, sinstr)
        If RET <> 0 Then
            funGetKouhoHinban1_9_1 = 3
            GoTo Apl_Exit
        End If
        RET = funCodeinGet("SB", "OS", sinstr, sResult)
        If RET <> 0 Then
            funGetKouhoHinban1_9_1 = 3
            GoTo Apl_Exit
        End If
        sinstr = sResult
        If tbl_spec1_9_1(0).PATTERN = " " Then
            sql = sql & " AND (" & sChkTable & "." & tbl_spec1_9_1(0).PATTERN1 & " IN  (" & sinstr & ") " & vbCrLf
            sql = sql & " OR " & sChkTable & "." & tbl_spec1_9_1(0).PATTERN1 & " IS NULL)" & vbCrLf
        Else
            sql = sql & " AND " & sChkTable & "." & tbl_spec1_9_1(0).PATTERN1 & " IN  (" & sinstr & ") " & vbCrLf
        End If
    End If
    '検査頻度＿抜
    If Mid(sChkCode, 14, 1) = "2" Then
        'マトリクス取得
        sResult = ""
        sinstr = ""
        If tbl_spec1_9_1(0).KENH_NUKI = "3" Or tbl_spec1_9_1(0).KENH_NUKI = "4" _
                Or tbl_spec1_9_1(0).KENH_NUKI = "6" Then
            RET = funCodeDBGet("SB", "HO", tbl_spec1_9_1(0).KENH_NUKI, 0, " ", sResult)
        Else
            RET = funCodeDBGet("SB", "HO", "ETC", 0, " ", sResult)
        End If
        If RET <> 0 Then
            funGetKouhoHinban1_9_1 = 5
            GoTo Apl_Exit
        End If
        For iCnt = 1 To 3
            If iCnt = 1 Then sNum = "3"
            If iCnt = 2 Then sNum = "4"
            If iCnt = 3 Then sNum = "6"
            If Mid(sResult, iCnt, 1) = "1" Then
                If sinstr = "" Then sinstr = "'" & sNum & "'" Else sinstr = sinstr & ", '" & sNum & "'"
            End If
        Next
        If sinstr <> "" Then
            If Mid(sResult, 4, 1) = "1" Then sql = sql & " AND (" Else sql = sql & " AND "
            sql = sql & sChkTable & "." & tbl_spec1_9_1(0).KENH_NUKI1 & " IN  (" & sinstr & ") " & vbCrLf
        End If
        If Mid(sResult, 4, 1) = "1" Then
            If sinstr <> "" Then sql = sql & " OR " Else sql = sql & " AND "
            sql = sql & "(" & sChkTable & "." & tbl_spec1_9_1(0).KENH_NUKI1 & " IS NULL" & vbCrLf
            sql = sql & " OR " & sChkTable & "." & tbl_spec1_9_1(0).KENH_NUKI1 & " NOT IN ('3', '4', '6'))" & vbCrLf
            If sinstr <> "" Then sql = sql & ")" & vbCrLf Else sql = sql & vbCrLf
        End If
    End If
    'AN温度
    If Mid(sChkCode, 15, 1) = "2" Then
        ReDim lsANCodeListWork(0) As String
        ReDim lsANCodeList(0) As String
        lsANCode = "AE"
        'AN温度マトリックスよりCodeの一覧を取得する
        RET = funCodeDBGetCodeList("SB", lsANCode, lsANCodeListWork)
        If RET < 0 Then
            funGetKouhoHinban1_9_1 = 6
            GoTo Apl_Exit
        End If
        
        For iCnt = 1 To UBound(lsANCodeListWork)
            RET = funCodeDBGetMatrixReturn("SB", lsANCode, lsANCodeListWork(iCnt), tbl_spec1_9_1(0).Antnp)
            If RET < 0 Then
                funGetKouhoHinban1_9_1 = 6
                GoTo Apl_Exit
            ElseIf RET = 0 Then
                ' AN温度チェックNGの値を保持する
                ReDim Preserve lsANCodeList(UBound(lsANCodeList) + 1) As String
                lsANCodeList(UBound(lsANCodeList)) = lsANCodeListWork(iCnt)
            End If
        Next iCnt

        'AN温度チェックNG以外のデータを取得する
        If UBound(lsANCodeList) <> 0 Then
            sql = sql & " AND (E050." & tbl_spec1_9_1(0).ANTNP1 & " NOT IN (" & vbCrLf
            For iCnt = 1 To UBound(lsANCodeList)
                If iCnt <> 1 Then
                    sql = sql & ","
                End If
                sql = sql & "'" & lsANCodeList(iCnt) & "'"
            Next iCnt
            sql = sql & "))" & vbCrLf
        End If
    End If
    'エピ厚中心 まだまだ(等号が入るか否かPending中)
    If Mid(sChkCode, 16, 1) = "2" Then
        sql = sql & " AND " & sChkTable & "." & tbl_spec1_9_1(0).EPATU1 & " >= " & tbl_spec1_9_1(0).EPATU & " " & vbCrLf
    End If

Make_Exit:
    sMakesql = sql
    
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funGetKouhoHinban1_9_1 = -4
    GoTo Apl_Exit
    
End Function

'------------------------------------------------
' 振替先と振替元の常識仕様チェック２
'------------------------------------------------

'概要      :指定された振替元品番と、以下の仕様値が一致している品番を抽出するSQL文を作成し、呼び出し元に返す。
'           （結晶面方位、ドーパント、結晶ドープ）
'           指定された振替元品番と、以下の仕様値がマトリクスで一致している品番を抽出するSQL文を作成する。
'           （溝位置方位、品種、引上速度、ＨＺタイプ、ドローチューブ）
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型           :説明
'          :sProccd         ,I  ,String       :工程番号
'          :sBlockid        ,I  ,String       :ﾌﾞﾛｯｸID、又は、結晶番号
'          :sOld_Hinban     ,I  ,String       :振替元品番
'          :sMakeSql        ,O  ,String       :作成SQL文
'          :戻り値          ,O  ,Integer      :取得の成否(0:正常取得, -1:取得ｴﾗｰ)
'説明      :
'履歴      :06/10/05 ooba

Public Function funGetKouhoHinban1_10(sProccd As String, sBlockId As String, sOld_Hinban As String, sMakesql As String) As Integer


    Dim RET         As Integer      '戻り値
    Dim sResult     As String       'コードＤＢ取得関数の取得変数
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim sinstr      As String       'ＳＱＬin句用文字列
    Dim sinstr1     As String       'ＳＱＬin句用文字列
    Dim sinstr2     As String       'ＳＱＬin句用文字列
    Dim sinstr3     As String       'ＳＱＬin句用文字列
    Dim sinstr4     As String       'ＳＱＬin句用文字列
    Dim sinstr5     As String       'ＳＱＬin句用文字列
    Dim sinstr6     As String       'ＳＱＬin句用文字列
    
    On Error GoTo Apl_down
    
    '戻り値初期化
    funGetKouhoHinban1_10 = 0
    
    '------------------------------------------ 振替元品種仕様データ取得 ------------------------------------------------------
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT E018.HSXCDIR,E018.HSXDOP,E023.HWFCDOP,E018.HSXDPDIR, " & vbCrLf
    sql = sql & "       E018.HSXCSCEN,"     ''2008/11/27 結晶面傾中心チェック緩和(2) ADD By Systech
    sql = sql & "       SUBSTR(E018.MCNO,1,1) MCNO1,SUBSTR(E018.MCNO,4,1) MCNO2,SUBSTR(E018.MCNO,3,1) MCNO3,E036.DCHYUUBU " & vbCrLf
    sql = sql & "FROM   TBCME018 E018,TBCME023 E023,TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E018.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E018.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E018.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' AND " & vbCrLf
    sql = sql & "       E023.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E023.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E023.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E023.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' AND " & vbCrLf
    sql = sql & "       E036.HINBAN                     =   '" & Mid(sOld_Hinban, 1, 8) & "'  AND " & vbCrLf
    sql = sql & "       TO_CHAR(E036.MNOREVNO, 'FM00')  =   '" & Mid(sOld_Hinban, 9, 2) & "'  AND " & vbCrLf
    sql = sql & "       E036.FACTORY                    =   '" & Mid(sOld_Hinban, 11, 1) & "' AND " & vbCrLf
    sql = sql & "       E036.OPECOND                    =   '" & Mid(sOld_Hinban, 12, 1) & "' " & vbCrLf
    
    On Error GoTo db_Error
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        funGetKouhoHinban1_10 = 10001
        GoTo db_Error
    End If
    
    '取得データセット
    Erase tbl_spec1_10
    With tbl_spec1_10(0)
        If IsNull(rs("HSXCDIR")) = False Then .HSXCDIR = rs("HSXCDIR") Else .HSXCDIR = " "          ' 結晶面方位
        If IsNull(rs("HSXCSCEN")) = False Then .HSXCSCEN = rs("HSXCSCEN") Else .HSXCSCEN = -1       ' 結晶面傾き中心    ''2008/11/27 結晶面傾中心チェック緩和(2) ADD By Systech
        If IsNull(rs("HSXDOP")) = False Then .HSXDOP = rs("HSXDOP") Else .HSXDOP = " "              ' ドーパント
        If IsNull(rs("HWFCDOP")) = False Then .HWFCDOP = rs("HWFCDOP") Else .HWFCDOP = " "          ' 結晶ドープ
        If IsNull(rs("HSXDPDIR")) = False Then .HSXDPDIR = rs("HSXDPDIR") Else .HSXDPDIR = " "      ' 溝位置方位
        If IsNull(rs("MCNO1")) = False Then .MCNO1 = rs("MCNO1") Else .MCNO1 = " "                  ' 品種
        If IsNull(rs("MCNO2")) = False Then .MCNO2 = rs("MCNO2") Else .MCNO2 = " "                  ' 引上げ速度
        If IsNull(rs("MCNO3")) = False Then .MCNO3 = rs("MCNO3") Else .MCNO3 = " "                  ' HZタイプ
        If IsNull(rs("DCHYUUBU")) = False Then .DCHYUUBU = rs("DCHYUUBU") Else .DCHYUUBU = " "      ' ドローチューブ
    End With
    
    Set rs = Nothing
    On Error GoTo Apl_down
    '------------------------------------------ 指示取得 ------------------------------------------------------
    sinstr1 = ""
    sinstr2 = ""
    sinstr3 = ""
    sinstr4 = ""
    sinstr5 = ""
    sinstr6 = ""
    '溝位置方位
    sResult = ""
    sinstr = ""
    RET = funCodeDBGet("SB", "MZ", tbl_spec1_10(0).HSXDPDIR, 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_10 = 10002
        GoTo Apl_Exit
    End If
    RET = funinfo2(sResult, sinstr)
    If RET <> 0 Then
        funGetKouhoHinban1_10 = 10003
        GoTo Apl_Exit
    End If
    RET = funCodeinGet("SB", "MZ", sinstr, sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_10 = 10003
        GoTo Apl_Exit
    End If
    sinstr1 = sResult
    '品種
    sResult = ""
    sinstr = ""
    RET = funCodeDBGet("SB", "HS", tbl_spec1_10(0).MCNO1, 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_10 = 10004
        GoTo Apl_Exit
    End If
    RET = funinfo2(sResult, sinstr)
    If RET <> 0 Then
        funGetKouhoHinban1_10 = 10005
        GoTo Apl_Exit
    End If
    RET = funCodeinGet("SB", "HS", sinstr, sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_10 = 10005
        GoTo Apl_Exit
    End If
    sinstr2 = sResult
    '引上げ速度
    sResult = ""
    sinstr = ""
    RET = funCodeDBGet("SB", "HK", tbl_spec1_10(0).MCNO2, 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_10 = 10006
        GoTo Apl_Exit
    End If
    RET = funinfo2(sResult, sinstr)
    If RET <> 0 Then
        funGetKouhoHinban1_10 = 10007
        GoTo Apl_Exit
    End If
    RET = funCodeinGet("SB", "HK", sinstr, sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_10 = 10007
        GoTo Apl_Exit
    End If
    sinstr3 = sResult
    'ＨＺタイプ
    sResult = ""
    sinstr = ""
    RET = funCodeDBGet("SB", "HZ", tbl_spec1_10(0).MCNO3, 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_10 = 10008
        GoTo Apl_Exit
    End If
    RET = funinfo2(sResult, sinstr)
    If RET <> 0 Then
        funGetKouhoHinban1_10 = 10009
        GoTo Apl_Exit
    End If
    RET = funCodeinGet("SB", "HZ", sinstr, sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_10 = 10009
        GoTo Apl_Exit
    End If
    sinstr4 = sResult
    'ドローチューブ
    sResult = ""
    sinstr = ""
    If tbl_spec1_10(0).DCHYUUBU <> " " Then
        RET = funCodeDBGet("SB", "DC", tbl_spec1_10(0).DCHYUUBU, 0, " ", sResult)
    Else
        RET = funCodeDBGet("SB", "DC", "2", 0, " ", sResult)
    End If
    If RET <> 0 Then
        funGetKouhoHinban1_10 = 10010
        GoTo Apl_Exit
    End If
    RET = funinfo2(sResult, sinstr)
    If RET <> 0 Then
        funGetKouhoHinban1_10 = 10011
        GoTo Apl_Exit
    End If
    RET = funCodeinGet("SB", "DC", sinstr, sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_10 = 10011
        GoTo Apl_Exit
    End If
    sinstr5 = sResult
    '結晶ドープ
    sResult = ""
    sinstr = ""
    RET = funCodeDBGet("SB", "SD", tbl_spec1_10(0).HWFCDOP, 0, " ", sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_10 = 10012
        GoTo Apl_Exit
    End If
    RET = funinfo2(sResult, sinstr)
    If RET <> 0 Then
        funGetKouhoHinban1_10 = 10013
        GoTo Apl_Exit
    End If
    RET = funCodeinGet("SB", "SD", sinstr, sResult)
    If RET <> 0 Then
        funGetKouhoHinban1_10 = 10013
        GoTo Apl_Exit
    End If
    sinstr6 = sResult

    '------------------------------------------ 振替元品種と同一仕様の品番を取得 ------------------------------------------------------
    'SQL文の作成
    sql = vbNullString
    sql = sql & "SELECT 'X' " & vbCrLf
    sql = sql & "FROM   TBCME018 E018, TBCME023 E023, TBCME036 E036 " & vbCrLf
    sql = sql & "WHERE  E018A.HINBAN                    = E018.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018A.MNOREVNO, 'FM00') = TO_CHAR(E018.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E018A.FACTORY                   = E018.FACTORY                      AND " & vbCrLf
    sql = sql & "       E018A.OPECOND                   = E018.OPECOND                      AND " & vbCrLf
    sql = sql & "       E018.HINBAN                     = E023.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018.MNOREVNO, 'FM00')  = TO_CHAR(E023.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E018.FACTORY                    = E023.FACTORY                      AND " & vbCrLf
    sql = sql & "       E018.OPECOND                    = E023.OPECOND                      AND " & vbCrLf
    sql = sql & "       E018.HINBAN                     = E036.HINBAN                       AND " & vbCrLf
    sql = sql & "       TO_CHAR(E018.MNOREVNO, 'FM00')  = TO_CHAR(E036.MNOREVNO, 'FM00')    AND " & vbCrLf
    sql = sql & "       E018.FACTORY                    = E036.FACTORY                      AND " & vbCrLf
    sql = sql & "       E018.OPECOND                    = E036.OPECOND                      AND " & vbCrLf
    
    sql = sql & "       E018.HSXCDIR                    = '" & tbl_spec1_10(0).HSXCDIR & "' AND " & vbCrLf
    sql = sql & "       E018.HSXDOP                     = '" & tbl_spec1_10(0).HSXDOP & "'  AND " & vbCrLf
    sql = sql & "       E018.HSXDPDIR               IN (" & sinstr1 & ") AND " & vbCrLf
    sql = sql & "       substr(E018.MCNO, 1, 1)     IN (" & sinstr2 & ") AND " & vbCrLf
    sql = sql & "       substr(E018.MCNO, 4, 1)     IN (" & sinstr3 & ") AND " & vbCrLf
    sql = sql & "       substr(E018.MCNO, 3, 1)     IN (" & sinstr4 & ") AND " & vbCrLf
    If tbl_spec1_10(0).DCHYUUBU = " " Then
        sql = sql & "       E036.DCHYUUBU is null OR E036.DCHYUUBU IN (" & sinstr5 & ") " & vbCrLf
    Else
        sql = sql & "       E036.DCHYUUBU               IN (" & sinstr5 & ")     " & vbCrLf
    End If
    sql = sql & " AND   E023.HWFCDOP                IN (" & sinstr6 & ")     " & vbCrLf
''2008/11/27 結晶面傾中心チェック緩和(2) ADD By Systech Start
    If tbl_spec1_10(0).HSXCSCEN = -1 Then
    Else
        sql = sql & " AND   ABS(" & tbl_spec1_10(0).HSXCSCEN & " - E018.HSXCSCEN ) <= 1.0 "
    End If
''2008/11/27 結晶面傾中心チェック緩和(2) ADD By Systech End
        
    sMakesql = sql
        
    '------------------------------------------ 終了処理  ------------------------------------------------------
Apl_Exit:
    
    Exit Function
    
Apl_down:
    funGetKouhoHinban1_10 = -4
    GoTo Apl_Exit
    
db_Error:
    Set rs = Nothing
    If funGetKouhoHinban1_10 = 0 Then
        funGetKouhoHinban1_10 = -3
    End If
    GoTo Apl_Exit

End Function


Attribute VB_Name = "s_cmzcF_cmkc001WF"
Option Explicit

' WFサンプル仕様(*は未チェックのパラメータ)
Public Type typ_SpWFSamp
    hin As tFullHinban      ' 品番

    HWFRHWYS As String * 1  ' 処理方法(Rs)
    HWFRSPOH As String * 1  ' 測定方法(Rs)*
    HWFRSPOT As String * 1  ' 測定点数(Rs) -> Heavy
    HWFRSPOI As String * 1  ' 測定位置(Rs)*

    HWFONHWS As String * 1  ' 処理方法(Oi)
    HWFONKWY As String * 2  ' 検査方法(Oi)
    HWFONSPH As String * 1  ' 測定方法(Oi)
    HWFONSPT As String * 1  ' 測定点数(Oi) -> Heavy
    HWFONSPI As String * 1  ' 測定位置(Oi)

    HWFBM1HS As String * 1  ' 処理方法(B1)
    HWFBM1SH As String * 1  ' 測定方法(B1)
    HWFBM1ST As String * 1  ' 測定点数(B1)
    HWFBM1SR As String * 1  ' 除外領域(B1)
    HWFBM1NS As String * 2  ' 熱処理法(B1)
    HWFBM1SZ As String * 1  ' 測定条件(B1)
    HWFBM1ET As Integer     ' 選択エッチ(B1)

    HWFBM2HS As String * 1  ' 処理方法(B2)
    HWFBM2SH As String * 1  ' 測定方法(B2)
    HWFBM2ST As String * 1  ' 測定点数(B2)
    HWFBM2SR As String * 1  ' 除外領域(B2)
    HWFBM2NS As String * 2  ' 熱処理法(B2)
    HWFBM2SZ As String * 1  ' 測定条件(B2)
    HWFBM2ET As Integer     ' 選択エッチ(B2)

    HWFBM3HS As String * 1  ' 処理方法(B3)
    HWFBM3SH As String * 1  ' 測定方法(B3)
    HWFBM3ST As String * 1  ' 測定点数(B3)
    HWFBM3SR As String * 1  ' 除外領域(B3)
    HWFBM3NS As String * 2  ' 熱処理法(B3)
    HWFBM3SZ As String * 1  ' 測定条件(B3)
    HWFBM3ET As Integer     ' 選択エッチ(B3)

    HWFOF1HS As String * 1  ' 処理方法(L1)
    HWFOF1SH As String * 1  ' 測定方法(L1)
    HWFOF1ST As String * 1  ' 測定点数(L1)
    HWFOF1SR As String * 1  ' 除外領域(L1)
    HWFOF1NS As String * 2  ' 熱処理法(L1)
    HWFOF1SZ As String * 1  ' 測定条件(L1)
    HWFOF1ET As Integer     ' 選択エッチ(L1)

    HWFOF2HS As String * 1  ' 処理方法(L2)
    HWFOF2SH As String * 1  ' 測定方法(L2)
    HWFOF2ST As String * 1  ' 測定点数(L2)
    HWFOF2SR As String * 1  ' 除外領域(L2)
    HWFOF2NS As String * 2  ' 熱処理法(L2)
    HWFOF2SZ As String * 1  ' 測定条件(L2)
    HWFOF2ET As Integer     ' 選択エッチ(L2)

    HWFOF3HS As String * 1  ' 処理方法(L3)
    HWFOF3SH As String * 1  ' 測定方法(L3)
    HWFOF3ST As String * 1  ' 測定点数(L3)
    HWFOF3SR As String * 1  ' 除外領域(L3)
    HWFOF3NS As String * 2  ' 熱処理法(L3)
    HWFOF3SZ As String * 1  ' 測定条件(L3)
    HWFOF3ET As Integer     ' 選択エッチ(L3)

'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''    HWFOF4HS As String * 1  ' 処理方法(L4)
'''    HWFOF4SH As String * 1  ' 測定方法(L4)
'''    HWFOF4ST As String * 1  ' 測定点数(L4)
'''    HWFOF4SR As String * 1  ' 除外領域(L4)
'''    HWFOF4NS As String * 2  ' 熱処理法(L4)
'''    HWFOF4SZ As String * 1  ' 測定条件(L4)
'''    HWFOF4ET As Integer     ' 選択エッチ(L4)
    
    HWFSIRDMX As Integer       '軸状転位上限(SIRD)
    HWFSIRDSZ As String * 1    '軸状転位測定条件(SIRD)
    HWFSIRDHT As String * 1    '軸状転位保証方法＿対(SIRD)
    HWFSIRDHS As String * 1    '軸状転位保証方法＿処(SIRD)
    HWFSIRDKM As String * 1    '軸状転位検査頻度＿枚(SIRD)
    HWFSIRDKH As String * 1    '軸状転位検査頻度＿保(SIRD)
    HWFSIRDKU As String * 1    '軸状転位検査頻度＿ウ(SIRD)
    HWFSIRDPS As String * 2    '軸状転位TB保証位置(SIRD)
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)

    HWFDSOHS As String * 1  ' 処理方法(DS)

    HWFMKHWS As String * 1  ' 処理方法(DZ)
    HWFMKSPH As String * 1  ' 測定方法(DZ)
    HWFMKSPT As String * 1  ' 測定点数(DZ)
    HWFMKSPR As String * 1  ' 除外領域(DZ)
    HWFMKNSW As String * 2  ' 熱処理法(DZ)
    HWFMKSZY As String * 1  ' 測定条件(DZ)
    HWFMKCET As Integer     ' 選択エッチ(DZ)

    HWFSPVHS As String * 1  ' 処理方法(SP/Fe濃度)
    HWFSPVSH As String * 1  ' 測定方法(SP/Fe濃度)*
    HWFSPVST As String * 1  ' 測定点数(SP/Fe濃度)*
    HWFSPVSI As String * 1  ' 測定位置(SP/Fe濃度)*
    HWFDLHWS As String * 1  ' 処理方法(SP/拡散長)
    HWFDLSPH As String * 1  ' 測定方法(SP/拡散長)*
    HWFDLSPT As String * 1  ' 測定点数(SP/拡散長)*
    HWFDLSPI As String * 1  ' 測定位置(SP/拡散長)*
    HWFNRHS  As String * 1  ' 処理方法(SP/Nr濃度)               06/06/08 ooba START ======>
    HWFNRSH  As String * 1  ' 測定方法(SP/Nr濃度)*
    HWFNRST  As String * 1  ' 測定点数(SP/Nr濃度)*
    HWFNRSI  As String * 1  ' 測定位置(SP/Nr濃度)*
    HWFSPVPUG   As String * 10      ' PUA限(SP/Fe濃度)*
    HWFSPVPUR   As String * 10      ' PUA率(SP/Fe濃度)*
    HWFSPVSTD   As String * 10      ' 標準偏差(SP/Fe濃度)*
    HWFDLPUG    As String * 10      ' PUA限(SP/拡散長)*
    HWFDLPUR    As String * 10      ' PUA率(SP/拡散長)*
    HWFNRPUG    As String * 10      ' PUA限(SP/Nr濃度)*
    HWFNRPUR    As String * 10      ' PUA率(SP/Nr濃度)*
    HWFNRSTD    As String * 10      ' 標準偏差(SP/Nr濃度)*      06/06/08 ooba END ========>

    HWFOS1HS As String * 1  ' 処理方法(D1)
    HWFOS1SH As String * 1  ' 測定方法(D1)*
    HWFOS1ST As String * 1  ' 測定点数(D1)*
    HWFOS1SI As String * 1  ' 測定位置(D1)*
    HWFOS1NS As String * 2  ' 熱処理法(D1)

    HWFOS2HS As String * 1  ' 処理方法(D2)
    HWFOS2SH As String * 1  ' 測定方法(D2)*
    HWFOS2ST As String * 1  ' 測定点数(D2)*
    HWFOS2SI As String * 1  ' 測定位置(D2)*
    HWFOS2NS As String * 2  ' 熱処理法(D2)

    HWFOS3HS As String * 1  ' 処理方法(D3)
    HWFOS3SH As String * 1  ' 測定方法(D3)*
    HWFOS3ST As String * 1  ' 測定点数(D3)*
    HWFOS3SI As String * 1  ' 測定位置(D3)*
    HWFOS3NS As String * 2  ' 熱処理法(D3)
    
    HWOTHER1 As String * 1  ' 検査項目(OT1) '03/05/21
    HWOTHER2 As String * 1  ' 検査項目(OT1) '03/05/21
    
    HWFZOHWS As String * 1  ' 処理方法(AO)  ''追加 03/12/05 ooba START ======>
    HWFZOSPH As String * 1  ' 測定方法(AO)*
    HWFZOSPT As String * 1  ' 測定点数(AO)*
    HWFZOSPI As String * 1  ' 測定位置(AO)*
    HWFZONSW As String * 2  ' 熱処理法(AO)  ''追加 03/12/05 ooba END ========>
    
    HWFDENHS As String * 1  ' 処理方法(GD/DEN)  '追加　05/01/18 ooba START ====>
    HWFLDLHS As String * 1  ' 処理方法(GD/LDL)
    HWFDVDHS As String * 1  ' 処理方法(GD/DVD2) '追加　05/01/18 ooba END ======>
    HWFGDSPH As String * 1  ' 測定方法(GD)　    '05/10/25 ooba
    HWFGDSPT As String * 1  ' 測定点数(GD)　    '05/10/25 ooba
    HWFGDZAR As String * 1  ' 除外領域(GD)　    '05/10/25 ooba
    
    HWFRKHNN As String * 1  ' 検査頻度_抜(Rs)   '追加　04/04/08 ooba START ====>
    HWFONKHN As String * 1  ' 検査頻度_抜(Oi)
    HWFOF1KN As String * 1  ' 検査頻度_抜(L1)
    HWFOF2KN As String * 1  ' 検査頻度_抜(L2)
    HWFOF3KN As String * 1  ' 検査頻度_抜(L3)
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''    HWFOF4KN As String * 1  ' 検査頻度_抜(L4)
    HWFSIRDKN As String * 1  ' 検査頻度_抜(SIRD)
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
    HWFBM1KN As String * 1  ' 検査頻度_抜(B1)
    HWFBM2KN As String * 1  ' 検査頻度_抜(B2)
    HWFBM3KN As String * 1  ' 検査頻度_抜(B3)
    HWFOS1KN As String * 1  ' 検査頻度_抜(D1)
    HWFOS2KN As String * 1  ' 検査頻度_抜(D2)
    HWFOS3KN As String * 1  ' 検査頻度_抜(D3)
    HWFDSOKN As String * 1  ' 検査頻度_抜(DS)
    HWFMKKHN As String * 1  ' 検査頻度_抜(DZ)
    HWFSPVKN As String * 1  ' 検査頻度_抜(SP/Fe濃度)
    HWFDLKHN As String * 1  ' 検査頻度_抜(SP/拡散長)
    HWFZOKHN As String * 1  ' 検査頻度_抜(AO)   '追加　04/04/08 ooba END ======>
    HWFGDKHN As String * 1  ' 検査頻度_抜(GD)　05/01/18 ooba
    HWFNRKN  As String * 1  ' 検査頻度_抜(SP/Nr濃度)  06/06/08 ooba
    
    HWFIGKBN As String * 1  ' IG区分
    HWFANTNP As Integer     ' DKアニール条件(温度)
    HWFANTIM As Integer     ' DKアニール条件(時間)
    HWFANGZY As String * 1  ' DKアニール条件(ガス)　04/07/23 ooba
    
    HWOTHER1MAI As String * 1  ' サンプル枚数(OT1) '04/06/23
    HWOTHER2MAI As String * 1  ' サンプル枚数(OT2) '04/06/23

''Upd Start (TCS)T.Terauchi 2005/10/05  抜試指示4.5ﾗｲﾝ対応
    HWFGDLINE   As String * 3   '品WFGDﾗｲﾝ数(TBCME036)
''Upd End   (TCS)T.Terauchi 2005/10/05  抜試指示4.5ﾗｲﾝ対応

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    HEPOF1NS As String * 2  ' 品熱処理法(OSF1E)
    HEPOF1SZ As String * 1  ' 品測定条件(OSF1E)
    HEPOF1ET As Integer     ' 品選択ET代(OSF1E)
    HEPOF1HS As String * 1  ' 品保証方法_処(OSF1E)
    HEPOF1SH As String * 1  ' 品測定位置_方(OSF1E)
    HEPOF1ST As String * 1  ' 品測定位置_点(OSF1E)
    HEPOF1SR As String * 1  ' 品測定位置_領(OSF1E)
    HEPOF1KN As String * 1  ' 品検査頻度_抜(OSF1E)
    HEPOF2NS As String * 2  ' 品熱処理法(OSF2E)
    HEPOF2SZ As String * 1  ' 品測定条件(OSF2E)
    HEPOF2ET As Integer     ' 品選択ET代(OSF2E)
    HEPOF2HS As String * 1  ' 品保証方法_処(OSF2E)
    HEPOF2SH As String * 1  ' 品測定位置_方(OSF2E)
    HEPOF2ST As String * 1  ' 品測定位置_点(OSF2E)
    HEPOF2SR As String * 1  ' 品測定位置_領(OSF2E)
    HEPOF2KN As String * 1  ' 品検査頻度_抜(OSF2E)
    HEPOF3NS As String * 2  ' 品熱処理法(OSF3E)
    HEPOF3SZ As String * 1  ' 品測定条件(OSF3E)
    HEPOF3ET As Integer     ' 品選択ET代(OSF3E)
    HEPOF3HS As String * 1  ' 品保証方法_処(OSF3E)
    HEPOF3SH As String * 1  ' 品測定位置_方(OSF3E)
    HEPOF3ST As String * 1  ' 品測定位置_点(OSF3E)
    HEPOF3SR As String * 1  ' 品測定位置_領(OSF3E)
    HEPOF3KN As String * 1  ' 品検査頻度_抜(OSF3E)
    HEPBM1NS As String * 2  ' 品熱処理法(BMD1E)
    HEPBM1SZ As String * 1  ' 品測定条件(BMD1E)
    HEPBM1ET As Integer     ' 品選択ET代(BMD1E)
    HEPBM1HS As String * 1  ' 品保証方法_処(BMD1E)
    HEPBM1SH As String * 1  ' 品測定位置_方(BMD1E)
    HEPBM1ST As String * 1  ' 品測定位置_点(BMD1E)
    HEPBM1SR As String * 1  ' 品測定位置_領(BMD1E)
    HEPBM1KN As String * 1  ' 品検査頻度_抜(BMD1E)
    HEPBM2NS As String * 2  ' 品熱処理法(BMD2E)
    HEPBM2SZ As String * 1  ' 品測定条件(BMD2E)
    HEPBM2ET As Integer     ' 品選択ET代(BMD2E)
    HEPBM2HS As String * 1  ' 品保証方法_処(BMD2E)
    HEPBM2SH As String * 1  ' 品測定位置_方(BMD2E)
    HEPBM2ST As String * 1  ' 品測定位置_点(BMD2E)
    HEPBM2SR As String * 1  ' 品測定位置_領(BMD2E)
    HEPBM2KN As String * 1  ' 品検査頻度_抜(BMD2E)
    HEPBM3NS As String * 2  ' 品熱処理法(BMD3E)
    HEPBM3SZ As String * 1  ' 品測定条件(BMD3E)
    HEPBM3ET As Integer     ' 品選択ET代(BMD3E)
    HEPBM3HS As String * 1  ' 品保証方法_処(BMD3E)
    HEPBM3SH As String * 1  ' 品測定位置_方(BMD3E)
    HEPBM3ST As String * 1  ' 品測定位置_点(BMD3E)
    HEPBM3SR As String * 1  ' 品測定位置_領(BMD3E)
    HEPBM3KN As String * 1  ' 品検査頻度_抜(BMD3E)
    HEPACEN  As Double      ' 品E1厚中心
    HEPANTNP As Integer     ' 品EPAN温度
    HEPANTIM As Integer     ' 品EPAN時間
    HEPIGKBN As String * 1  ' 品EPIG区分
    HEPANGZY As String * 1  ' 品EP高温ANガス条件
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    HWFGDSZY As String * 1  ' 品ＷＦＧＤ測定条件
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
End Type

' WFサンプルテーブル
Public Type typ_WFSample
    CRYINDRS As String * 1  ' 検査項目(Rs)
    CRYINDOI As String * 1  ' 検査項目(Oi)
    CRYINDB1 As String * 1  ' 検査項目(B1)
    CRYINDB2 As String * 1  ' 検査項目(B2）
    CRYINDB3 As String * 1  ' 検査項目(B3)
    CRYINDL1 As String * 1  ' 検査項目(L1)
    CRYINDL2 As String * 1  ' 検査項目(L2)
    CRYINDL3 As String * 1  ' 検査項目(L3)
    CRYINDL4 As String * 1  ' 検査項目(L4)
    CRYINDDS As String * 1  ' 検査項目(DS)
    CRYINDDZ As String * 1  ' 検査項目(DZ)
    CRYINDSP As String * 1  ' 検査項目(SP)
    CRYINDD1 As String * 1  ' 検査項目(D1)
    CRYINDD2 As String * 1  ' 検査項目(D2)
    CRYINDD3 As String * 1  ' 検査項目(D3)
    CRYINDOT1 As String * 1 ' 検査項目(OT1) 'Add.03/05/20
    CRYINDOT2 As String * 1 ' 検査項目(OT2) 'Add.03/05/20
    CRYINDAO As String * 1  ' 検査有無(AO)  '追加 03/12/05 ooba
    CRYINDGD As String * 1  ' 検査有無(GD)  '追加 05/01/18 ooba
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    CRYINDGD2 As String * 1  ' 検査有無(GD測定条件用)
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
    WFHSGD As String * 1    ' 保証FLG(GD)   '追加 05/01/18 ooba
' 06/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    EPIINDL1 As String * 1 ' 検査有無(OSF1E)
    EPIINDL2 As String * 1 ' 検査有無(OSF2E)
    EPIINDL3 As String * 1 ' 検査有無(OSF3E)
    EPIINDB1 As String * 1 ' 検査有無(BMD1E)
    EPIINDB2 As String * 1 ' 検査有無(BMD2E)
    EPIINDB3 As String * 1 ' 検査有無(BMD3E)
' 06/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
End Type

'概要      :製品仕様WFデータの取得ドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型               ,説明
'　　      :pSpWFSamp　　　,IO ,typ_SpWFSamp   　,WFサンプル仕様
'　　      :戻り値         ,O  ,FUNCTION_RETURN　,読み込みの成否
Public Function scmzc_getWF(pSpWFSamp As typ_SpWFSamp) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String
    Dim sOT1    As String '03/05/21 後藤
    Dim sOT2    As String '03/05/21 後藤
    Dim sMAI1    As String '04/06/23
    Dim sMAI2    As String '04/06/23
    Dim rtn     As FUNCTION_RETURN
    '' エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcF_cmkc001WF.bas -- Function scmzc_getWF"

    '' 製品仕様の取得
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''    sql = "select " & _
'''          "E021HWFRSPOH, E021HWFRSPOT, E021HWFRSPOI, E021HWFRHWYS, E024HWFMKSPH, " & _
'''          "E024HWFMKSPT, E024HWFMKSPR, E024HWFMKHWS, E024HWFMKSZY, E024HWFMKNSW, " & _
'''          "E024HWFMKCET, E025HWFONSPH, E025HWFONSPT, E025HWFONSPI, E025HWFONHWS, " & _
'''          "E025HWFONKWY, E025HWFOS1NS, E025HWFOS1SH, E025HWFOS1ST, E025HWFOS1SI, " & _
'''          "E025HWFOS1HS, E025HWFOS2NS, E025HWFOS2SH, E025HWFOS2ST, E025HWFOS2SI, " & _
'''          "E025HWFOS2HS, E025HWFOS3NS, E025HWFOS3SH, E025HWFOS3ST, E025HWFOS3SI, " & _
'''          "E025HWFOS3HS, E025HWFANTNP, E025HWFANTIM, E026HWFDSOHS, E028HWFSPVSH, " & _
'''          "E028HWFSPVST, E028HWFSPVSI, E028HWFSPVHS, E028HWFDLSPH, E028HWFDLSPT, " & _
'''          "E028HWFDLSPI, E028HWFDLHWS, E029HWFOF1ET, E029HWFOF1NS, E029HWFOF1SZ, " & _
'''          "E029HWFOF1SH, E029HWFOF1ST, E029HWFOF1SR, E029HWFOF1HS, E029HWFOF2ET, " & _
'''          "E029HWFOF2NS, E029HWFOF2SZ, E029HWFOF2SH, E029HWFOF2ST, E029HWFOF2SR, " & _
'''          "E029HWFOF2HS, E029HWFOF3ET, E029HWFOF3NS, E029HWFOF3SZ, E029HWFOF3SH, " & _
'''          "E029HWFOF3ST, E029HWFOF3SR, E029HWFOF3HS, E029HWFOF4ET, E029HWFOF4NS, " & _
'''          "E029HWFOF4SZ, E029HWFOF4SH, E029HWFOF4ST, E029HWFOF4SR, E029HWFOF4HS, " & _
'''          "E029HWFBM1ET, E029HWFBM1NS, E029HWFBM1SZ, E029HWFBM1SH, E029HWFBM1ST, " & _
'''          "E029HWFBM1SR, E029HWFBM1HS, E029HWFBM2ET, E029HWFBM2NS, E029HWFBM2SZ, " & _
'''          "E029HWFBM2SH, E029HWFBM2ST, E029HWFBM2SR, E029HWFBM2HS, E029HWFBM3ET, " & _
'''          "E029HWFBM3NS, E029HWFBM3SZ, E029HWFBM3SH, E029HWFBM3ST, E029HWFBM3SR, E029HWFBM3HS" & _
'''          " from VECME001" & _
'''          " where E018HINBAN='" & pSpWFSamp.hin.hinban & "' and E018MNOREVNO=" & pSpWFSamp.hin.mnorevno & _
'''          " and E018FACTORY='" & pSpWFSamp.hin.FACTORY & "' and E018OPECOND='" & pSpWFSamp.hin.OPECOND & "'"
          
    sql = "select " & _
          "E021HWFRSPOH, E021HWFRSPOT, E021HWFRSPOI, E021HWFRHWYS, E024HWFMKSPH, " & _
          "E024HWFMKSPT, E024HWFMKSPR, E024HWFMKHWS, E024HWFMKSZY, E024HWFMKNSW, " & _
          "E024HWFMKCET, E025HWFONSPH, E025HWFONSPT, E025HWFONSPI, E025HWFONHWS, " & _
          "E025HWFONKWY, E025HWFOS1NS, E025HWFOS1SH, E025HWFOS1ST, E025HWFOS1SI, " & _
          "E025HWFOS1HS, E025HWFOS2NS, E025HWFOS2SH, E025HWFOS2ST, E025HWFOS2SI, " & _
          "E025HWFOS2HS, E025HWFOS3NS, E025HWFOS3SH, E025HWFOS3ST, E025HWFOS3SI, " & _
          "E025HWFOS3HS, E025HWFANTNP, E025HWFANTIM, E026HWFDSOHS, E028HWFSPVSH, " & _
          "E028HWFSPVST, E028HWFSPVSI, E028HWFSPVHS, E028HWFDLSPH, E028HWFDLSPT, " & _
          "E028HWFDLSPI, E028HWFDLHWS, E029HWFOF1ET, E029HWFOF1NS, E029HWFOF1SZ, " & _
          "E029HWFOF1SH, E029HWFOF1ST, E029HWFOF1SR, E029HWFOF1HS, E029HWFOF2ET, " & _
          "E029HWFOF2NS, E029HWFOF2SZ, E029HWFOF2SH, E029HWFOF2ST, E029HWFOF2SR, " & _
          "E029HWFOF2HS, E029HWFOF3ET, E029HWFOF3NS, E029HWFOF3SZ, E029HWFOF3SH, " & _
          "E029HWFOF3ST, E029HWFOF3SR, E029HWFOF3HS, " & _
          "E029HWFBM1ET, E029HWFBM1NS, E029HWFBM1SZ, E029HWFBM1SH, E029HWFBM1ST, " & _
          "E029HWFBM1SR, E029HWFBM1HS, E029HWFBM2ET, E029HWFBM2NS, E029HWFBM2SZ, " & _
          "E029HWFBM2SH, E029HWFBM2ST, E029HWFBM2SR, E029HWFBM2HS, E029HWFBM3ET, " & _
          "E029HWFBM3NS, E029HWFBM3SZ, E029HWFBM3SH, E029HWFBM3ST, E029HWFBM3SR, E029HWFBM3HS" & _
          " from VECME001" & _
          " where E018HINBAN='" & pSpWFSamp.hin.hinban & "' and E018MNOREVNO=" & pSpWFSamp.hin.mnorevno & _
          " and E018FACTORY='" & pSpWFSamp.hin.factory & "' and E018OPECOND='" & pSpWFSamp.hin.opecond & "'"
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo PROC_EXIT
    End If

    With pSpWFSamp
        .HWFRSPOH = rs("E021HWFRSPOH")
        .HWFRSPOT = rs("E021HWFRSPOT")
        .HWFRSPOI = rs("E021HWFRSPOI")
        .HWFRHWYS = rs("E021HWFRHWYS")
        .HWFMKSPH = rs("E024HWFMKSPH")
        .HWFMKSPT = rs("E024HWFMKSPT")
        .HWFMKSPR = rs("E024HWFMKSPR")
        .HWFMKHWS = rs("E024HWFMKHWS")
        .HWFMKSZY = rs("E024HWFMKSZY")
        .HWFMKNSW = rs("E024HWFMKNSW")
        .HWFMKCET = fncNullCheck(rs("E024HWFMKCET"))
        .HWFONSPH = rs("E025HWFONSPH")
        .HWFONSPT = rs("E025HWFONSPT")
        .HWFONSPI = rs("E025HWFONSPI")
        .HWFONHWS = rs("E025HWFONHWS")
        .HWFONKWY = rs("E025HWFONKWY")
        .HWFOS1NS = rs("E025HWFOS1NS")
        .HWFOS1HS = rs("E025HWFOS1HS")
        .HWFOS1SH = rs("E025HWFOS1SH")
        .HWFOS1ST = rs("E025HWFOS1ST")
        .HWFOS1SI = rs("E025HWFOS1SI")
        .HWFOS2NS = rs("E025HWFOS2NS")
        .HWFOS2SH = rs("E025HWFOS2SH")
        .HWFOS2ST = rs("E025HWFOS2ST")
        .HWFOS2SI = rs("E025HWFOS2SI")
        .HWFOS2HS = rs("E025HWFOS2HS")
        .HWFOS3NS = rs("E025HWFOS3NS")
        .HWFOS3SH = rs("E025HWFOS3SH")
        .HWFOS3ST = rs("E025HWFOS3ST")
        .HWFOS3SI = rs("E025HWFOS3SI")
        .HWFOS3HS = rs("E025HWFOS3HS")
        .HWFANTNP = fncNullCheck(rs("E025HWFANTNP"))
        .HWFANTIM = fncNullCheck(rs("E025HWFANTIM"))
        .HWFDSOHS = rs("E026HWFDSOHS")
        .HWFSPVSH = rs("E028HWFSPVSH")
        .HWFSPVST = rs("E028HWFSPVST")
        .HWFSPVSI = rs("E028HWFSPVSI")
        .HWFSPVHS = rs("E028HWFSPVHS")
        .HWFDLSPH = rs("E028HWFDLSPH")
        .HWFDLSPT = rs("E028HWFDLSPT")
        .HWFDLSPI = rs("E028HWFDLSPI")
        .HWFDLHWS = rs("E028HWFDLHWS")
        .HWFOF1ET = fncNullCheck(rs("E029HWFOF1ET"))
        .HWFOF1NS = rs("E029HWFOF1NS")
        .HWFOF1SZ = rs("E029HWFOF1SZ")
        .HWFOF1SH = rs("E029HWFOF1SH")
        .HWFOF1ST = rs("E029HWFOF1ST")
        .HWFOF1SR = rs("E029HWFOF1SR")
        .HWFOF1HS = rs("E029HWFOF1HS")
        .HWFOF2ET = fncNullCheck(rs("E029HWFOF2ET"))
        .HWFOF2NS = rs("E029HWFOF2NS")
        .HWFOF2SZ = rs("E029HWFOF2SZ")
        .HWFOF2SH = rs("E029HWFOF2SH")
        .HWFOF2ST = rs("E029HWFOF2ST")
        .HWFOF2SR = rs("E029HWFOF2SR")
        .HWFOF2HS = rs("E029HWFOF2HS")
        .HWFOF3ET = fncNullCheck(rs("E029HWFOF3ET"))
        .HWFOF3NS = rs("E029HWFOF3NS")
        .HWFOF3SZ = rs("E029HWFOF3SZ")
        .HWFOF3SH = rs("E029HWFOF3SH")
        .HWFOF3ST = rs("E029HWFOF3ST")
        .HWFOF3SR = rs("E029HWFOF3SR")
        .HWFOF3HS = rs("E029HWFOF3HS")
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''        .HWFOF4ET = fncNullCheck(rs("E029HWFOF4ET"))
'''        .HWFOF4NS = rs("E029HWFOF4NS")
'''        .HWFOF4SZ = rs("E029HWFOF4SZ")
'''        .HWFOF4SH = rs("E029HWFOF4SH")
'''        .HWFOF4ST = rs("E029HWFOF4ST")
'''        .HWFOF4SR = rs("E029HWFOF4SR")
'''        .HWFOF4HS = rs("E029HWFOF4HS")
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
        .HWFBM1ET = fncNullCheck(rs("E029HWFBM1ET"))
        .HWFBM1NS = rs("E029HWFBM1NS")
        .HWFBM1SZ = rs("E029HWFBM1SZ")
        .HWFBM1SH = rs("E029HWFBM1SH")
        .HWFBM1ST = rs("E029HWFBM1ST")
        .HWFBM1SR = rs("E029HWFBM1SR")
        .HWFBM1HS = rs("E029HWFBM1HS")
        .HWFBM2ET = fncNullCheck(rs("E029HWFBM2ET"))
        .HWFBM2NS = rs("E029HWFBM2NS")
        .HWFBM2SZ = rs("E029HWFBM2SZ")
        .HWFBM2SH = rs("E029HWFBM2SH")
        .HWFBM2ST = rs("E029HWFBM2ST")
        .HWFBM2SR = rs("E029HWFBM2SR")
        .HWFBM2HS = rs("E029HWFBM2HS")
        .HWFBM3ET = fncNullCheck(rs("E029HWFBM3ET"))
        .HWFBM3NS = rs("E029HWFBM3NS")
        .HWFBM3SZ = rs("E029HWFBM3SZ")
        .HWFBM3SH = rs("E029HWFBM3SH")
        .HWFBM3ST = rs("E029HWFBM3ST")
        .HWFBM3SR = rs("E029HWFBM3SR")
        .HWFBM3HS = rs("E029HWFBM3HS")
        'rtn = scmzc_getE036(pSpWFSamp.HIN, sOT1, sOT2)   2004/06/23
        'rtn = scmzc_getE036(pSpWFSamp.HIN, sOT1, sOT2)    '2004/07/12 koyama update
        rtn = scmzc_getE036(pSpWFSamp.hin, sOT1, sOT2, sMAI1, sMAI2) '2004/07/12 koyama update
        If rtn = FUNCTION_RETURN_FAILURE Then
            rs.Close
            scmzc_getWF = FUNCTION_RETURN_FAILURE
            GoTo PROC_EXIT
        End If
        .HWOTHER1 = sOT1 '### 03/05/20
        .HWOTHER2 = sOT2
        .HWOTHER1MAI = sMAI1   '04/06/23
        .HWOTHER2MAI = sMAI2   '04/06/23
    End With
    rs.Close
    
    '検査頻度_抜ﾃﾞｰﾀ取得　04/04/08 ooba START ==========================================>
    sql = "select "
    sql = sql & "TBCME026.HWFGDKHN, "   '検査頻度_抜(GD)　05/01/18 ooba
    sql = sql & "TBCME024.HWFANGZY, "   '品ＷＦ高温ＡＮガス条件　04/07/23 ooba
    sql = sql & "TBCME021.HWFRKHNN, "
    sql = sql & "TBCME025.HWFONKHN, "
    sql = sql & "TBCME029.HWFOF1KN, "
    sql = sql & "TBCME029.HWFOF2KN, "
    sql = sql & "TBCME029.HWFOF3KN, "
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''    sql = sql & "TBCME029.HWFOF4KN, "
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
    sql = sql & "TBCME029.HWFBM1KN, "
    sql = sql & "TBCME029.HWFBM2KN, "
    sql = sql & "TBCME029.HWFBM3KN, "
    sql = sql & "TBCME025.HWFOS1KN, "
    sql = sql & "TBCME025.HWFOS2KN, "
    sql = sql & "TBCME025.HWFOS3KN, "
    sql = sql & "TBCME026.HWFDSOKN, "
    sql = sql & "TBCME024.HWFMKKHN, "
    sql = sql & "TBCME028.HWFSPVKN, "
    sql = sql & "TBCME028.HWFDLKHN, "
    sql = sql & "TBCME025.HWFZOKHN "
    sql = sql & "from TBCME021, TBCME024, TBCME025, TBCME026, TBCME028, TBCME029 "
    sql = sql & "where TBCME021.HINBAN = TBCME024.HINBAN "
    sql = sql & "and TBCME021.MNOREVNO = TBCME024.MNOREVNO "
    sql = sql & "and TBCME021.FACTORY = TBCME024.FACTORY "
    sql = sql & "and TBCME021.OPECOND = TBCME024.OPECOND "
    sql = sql & "and TBCME021.HINBAN = TBCME025.HINBAN "
    sql = sql & "and TBCME021.MNOREVNO = TBCME025.MNOREVNO "
    sql = sql & "and TBCME021.FACTORY = TBCME025.FACTORY "
    sql = sql & "and TBCME021.OPECOND = TBCME025.OPECOND "
    sql = sql & "and TBCME021.HINBAN = TBCME026.HINBAN "
    sql = sql & "and TBCME021.MNOREVNO = TBCME026.MNOREVNO "
    sql = sql & "and TBCME021.FACTORY = TBCME026.FACTORY "
    sql = sql & "and TBCME021.OPECOND = TBCME026.OPECOND "
    sql = sql & "and TBCME021.HINBAN = TBCME028.HINBAN "
    sql = sql & "and TBCME021.MNOREVNO = TBCME028.MNOREVNO "
    sql = sql & "and TBCME021.FACTORY = TBCME028.FACTORY "
    sql = sql & "and TBCME021.OPECOND = TBCME028.OPECOND "
    sql = sql & "and TBCME021.HINBAN = TBCME029.HINBAN "
    sql = sql & "and TBCME021.MNOREVNO = TBCME029.MNOREVNO "
    sql = sql & "and TBCME021.FACTORY = TBCME029.FACTORY "
    sql = sql & "and TBCME021.OPECOND = TBCME029.OPECOND "
    sql = sql & "and TBCME021.HINBAN = '" & pSpWFSamp.hin.hinban & "' "
    sql = sql & "and TBCME021.MNOREVNO = " & pSpWFSamp.hin.mnorevno & " "
    sql = sql & "and TBCME021.FACTORY = '" & pSpWFSamp.hin.factory & "' "
    sql = sql & "and TBCME021.OPECOND = '" & pSpWFSamp.hin.opecond & "' "

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo PROC_EXIT
    End If
    
    With pSpWFSamp
        If IsNull(rs("HWFGDKHN")) = False Then .HWFGDKHN = rs("HWFGDKHN") Else .HWFGDKHN = " "  '05/01/18 ooba
        If IsNull(rs("HWFANGZY")) = False Then .HWFANGZY = rs("HWFANGZY") Else .HWFANGZY = " "  '04/07/23 ooba
        If IsNull(rs("HWFRKHNN")) = False Then .HWFRKHNN = rs("HWFRKHNN") Else .HWFRKHNN = " "
        If IsNull(rs("HWFONKHN")) = False Then .HWFONKHN = rs("HWFONKHN") Else .HWFONKHN = " "
        If IsNull(rs("HWFOF1KN")) = False Then .HWFOF1KN = rs("HWFOF1KN") Else .HWFOF1KN = " "
        If IsNull(rs("HWFOF2KN")) = False Then .HWFOF2KN = rs("HWFOF2KN") Else .HWFOF2KN = " "
        If IsNull(rs("HWFOF3KN")) = False Then .HWFOF3KN = rs("HWFOF3KN") Else .HWFOF3KN = " "
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''        If IsNull(rs("HWFOF4KN")) = False Then .HWFOF4KN = rs("HWFOF4KN") Else .HWFOF4KN = " "
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
        If IsNull(rs("HWFBM1KN")) = False Then .HWFBM1KN = rs("HWFBM1KN") Else .HWFBM1KN = " "
        If IsNull(rs("HWFBM2KN")) = False Then .HWFBM2KN = rs("HWFBM2KN") Else .HWFBM2KN = " "
        If IsNull(rs("HWFBM3KN")) = False Then .HWFBM3KN = rs("HWFBM3KN") Else .HWFBM3KN = " "
        If IsNull(rs("HWFOS1KN")) = False Then .HWFOS1KN = rs("HWFOS1KN") Else .HWFOS1KN = " "
        If IsNull(rs("HWFOS2KN")) = False Then .HWFOS2KN = rs("HWFOS2KN") Else .HWFOS2KN = " "
        If IsNull(rs("HWFOS3KN")) = False Then .HWFOS3KN = rs("HWFOS3KN") Else .HWFOS3KN = " "
        If IsNull(rs("HWFDSOKN")) = False Then .HWFDSOKN = rs("HWFDSOKN") Else .HWFDSOKN = " "
        If IsNull(rs("HWFMKKHN")) = False Then .HWFMKKHN = rs("HWFMKKHN") Else .HWFMKKHN = " "
        If IsNull(rs("HWFSPVKN")) = False Then .HWFSPVKN = rs("HWFSPVKN") Else .HWFSPVKN = " "
        If IsNull(rs("HWFDLKHN")) = False Then .HWFDLKHN = rs("HWFDLKHN") Else .HWFDLKHN = " "
        If IsNull(rs("HWFZOKHN")) = False Then .HWFZOKHN = rs("HWFZOKHN") Else .HWFZOKHN = " "
    End With
    rs.Close
    '検査頻度_抜ﾃﾞｰﾀ取得　04/04/08 ooba END ============================================>
    
    '' 残存酸素仕様取得　03/12/05 ooba START ===========================================>
    sql = "select HWFZOHWS, HWFZOSPH, HWFZOSPT, HWFZOSPI, HWFZONSW from TBCME025 "
    sql = sql & "where HINBAN = '" & pSpWFSamp.hin.hinban & "' "
    sql = sql & "and MNOREVNO = " & pSpWFSamp.hin.mnorevno & " "
    sql = sql & "and FACTORY = '" & pSpWFSamp.hin.factory & "' "
    sql = sql & "and OPECOND = '" & pSpWFSamp.hin.opecond & "' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo PROC_EXIT
    End If
    
    If IsNull(rs("HWFZOHWS")) = False Then pSpWFSamp.HWFZOHWS = rs("HWFZOHWS") Else pSpWFSamp.HWFZOHWS = " "
    If IsNull(rs("HWFZOSPH")) = False Then pSpWFSamp.HWFZOSPH = rs("HWFZOSPH") Else pSpWFSamp.HWFZOSPH = " "
    If IsNull(rs("HWFZOSPT")) = False Then pSpWFSamp.HWFZOSPT = rs("HWFZOSPT") Else pSpWFSamp.HWFZOSPT = " "
    If IsNull(rs("HWFZOSPI")) = False Then pSpWFSamp.HWFZOSPI = rs("HWFZOSPI") Else pSpWFSamp.HWFZOSPI = " "
    If IsNull(rs("HWFZONSW")) = False Then pSpWFSamp.HWFZONSW = rs("HWFZONSW") Else pSpWFSamp.HWFZONSW = " "
    
    rs.Close
    '' 残存酸素仕様取得　03/12/05 ooba END =============================================>
    
    '' GD仕様取得　05/01/18 ooba START ================================================>
    
''Upd start (TCS)T.Terauchi 2005/10/05  抜試指示4.5ﾗｲﾝ対応
'    sql = "select "
'    sql = sql & "HWFDENHS, "        '処理方法(GD/DEN)
'    sql = sql & "HWFLDLHS, "        '処理方法(GD/LDL)
'    sql = sql & "HWFDVDHS "         '処理方法(GD/DVD2)
'    sql = sql & "from TBCME026 "
'    sql = sql & "where HINBAN = '" & pSpWFSamp.HIN.hinban & "' "
'    sql = sql & "and MNOREVNO = " & pSpWFSamp.HIN.mnorevno & " "
'    sql = sql & "and FACTORY = '" & pSpWFSamp.HIN.factory & "' "
'    sql = sql & "and OPECOND = '" & pSpWFSamp.HIN.opecond & "' "
    sql = "select "
    sql = sql & "T1.HWFGDSPH AS HWFGDSPH, "         '測定方法(GD)　05/10/25 ooba
    sql = sql & "T1.HWFGDSPT AS HWFGDSPT, "         '測定点数(GD)　05/10/25 ooba
    sql = sql & "T1.HWFGDZAR AS HWFGDZAR, "         '除外領域(GD)　05/10/25 ooba
    sql = sql & "T1.HWFDENHS AS HWFDENHS, "         '処理方法(GD/DEN)
    sql = sql & "T1.HWFLDLHS AS HWFLDLHS, "         '処理方法(GD/LDL)
    sql = sql & "T1.HWFDVDHS AS HWFDVDHS"           '処理方法(GD/DVD2)
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    sql = sql & ",T1.HWFGDSZY AS HWFGDSZY"          '測定条件(GD)
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
    sql = sql & ",T2.HWFGDLINE AS HWFGDLINE "       'ﾗｲﾝ数
    sql = sql & "from TBCME026 T1,TBCME036 T2 "
    sql = sql & "where T1.HINBAN = '" & pSpWFSamp.hin.hinban & "' "
    sql = sql & "and T1.MNOREVNO = " & pSpWFSamp.hin.mnorevno & " "
    sql = sql & "and T1.FACTORY = '" & pSpWFSamp.hin.factory & "' "
    sql = sql & "and T1.OPECOND = '" & pSpWFSamp.hin.opecond & "' "
    sql = sql & "and T1.HINBAN = T2.HINBAN "
    sql = sql & "and T1.MNOREVNO = T2.MNOREVNO "
    sql = sql & "and T1.FACTORY = T2.FACTORY "
    sql = sql & "and T1.OPECOND = T2.OPECOND "
''Upd end   (TCS)T.Terauchi 2005/10/05  抜試指示4.5ﾗｲﾝ対応

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo PROC_EXIT
    End If
    
    If IsNull(rs("HWFGDSPH")) = False Then pSpWFSamp.HWFGDSPH = rs("HWFGDSPH") Else pSpWFSamp.HWFGDSPH = " "  '05/10/25 ooba
    If IsNull(rs("HWFGDSPT")) = False Then pSpWFSamp.HWFGDSPT = rs("HWFGDSPT") Else pSpWFSamp.HWFGDSPT = " "  '05/10/25 ooba
    If IsNull(rs("HWFGDZAR")) = False Then pSpWFSamp.HWFGDZAR = rs("HWFGDZAR") Else pSpWFSamp.HWFGDZAR = " "  '05/10/25 ooba
    If IsNull(rs("HWFDENHS")) = False Then pSpWFSamp.HWFDENHS = rs("HWFDENHS") Else pSpWFSamp.HWFDENHS = " "
    If IsNull(rs("HWFLDLHS")) = False Then pSpWFSamp.HWFLDLHS = rs("HWFLDLHS") Else pSpWFSamp.HWFLDLHS = " "
    If IsNull(rs("HWFDVDHS")) = False Then pSpWFSamp.HWFDVDHS = rs("HWFDVDHS") Else pSpWFSamp.HWFDVDHS = " "
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
    If IsNull(rs("HWFGDSZY")) = False Then pSpWFSamp.HWFGDSZY = rs("HWFGDSZY") Else pSpWFSamp.HWFGDSZY = " "
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
''Upd Start (TCS)T.Terauchi 2005/10/05  抜試指示4.5ﾗｲﾝ対応
    If IsNull(rs("HWFGDLINE")) = False Then pSpWFSamp.HWFGDLINE = CStr(rs("HWFGDLINE"))
''Upd End   (TCS)T.Terauchi 2005/10/05  抜試指示4.5ﾗｲﾝ対応
    
    rs.Close
    '' GD仕様取得　05/01/18 ooba END ==================================================>
    
    '' SPV仕様取得　06/06/08 ooba START ===============================================>
    sql = "select HWFNRHS, "                    '品WFSPVNR保証方法_処
    sql = sql & "HWFNRSH, "                     '品WFSPVNR測定位置_方
    sql = sql & "HWFNRST, "                     '品WFSPVNR測定位置_点
    sql = sql & "HWFNRSI, "                     '品WFSPVNR測定位置_位
    sql = sql & "HWFNRKN, "                     '品WFSPVNR検査頻度_抜
    sql = sql & "HWFSPVPUG, "                   '品WFSPVFEPUA限
    sql = sql & "HWFSPVPUR, "                   '品WFSPVFEPUA率
    sql = sql & "HWFSPVSTD, "                   '品WFSPVFE標準偏差
    sql = sql & "HWFDLPUG, "                    '品WF拡散長PUA限
    sql = sql & "HWFDLPUR, "                    '品WF拡散長PUA率
    sql = sql & "HWFNRPUG, "                    '品WFSPVNRPUA限
    sql = sql & "HWFNRPUR, "                    '品WFSPVNRPUA率
    sql = sql & "HWFNRSTD "                     '品WFSPVNR標準偏差
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
    sql = sql & ",HWFSIRDMX, "                  '軸状転位上限
    sql = sql & "HWFSIRDSZ, "                   '軸状転位測定条件
    sql = sql & "HWFSIRDHT, "                   '軸状転位保証方法＿対
    sql = sql & "HWFSIRDHS, "                   '軸状転位保証方法_処
    sql = sql & "HWFSIRDKM, "                   '軸状転位検査頻度＿枚
    sql = sql & "HWFSIRDKN, "                   '軸状転位検査頻度_抜
    sql = sql & "HWFSIRDKH, "                   '軸状転位検査頻度＿保
    sql = sql & "HWFSIRDKU, "                   '軸状転位検査頻度＿ウ
    sql = sql & "HWFSIRDPS  "                   '軸状転位TB保証位置
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
    sql = sql & "from TBCME048 "
    sql = sql & "where HINBAN = '" & pSpWFSamp.hin.hinban & "' "
    sql = sql & "and MNOREVNO = " & pSpWFSamp.hin.mnorevno & " "
    sql = sql & "and FACTORY = '" & pSpWFSamp.hin.factory & "' "
    sql = sql & "and OPECOND = '" & pSpWFSamp.hin.opecond & "' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo PROC_EXIT
    End If
    
    If IsNull(rs("HWFNRHS")) = False Then pSpWFSamp.HWFNRHS = rs("HWFNRHS") Else pSpWFSamp.HWFNRHS = " "
    If IsNull(rs("HWFNRSH")) = False Then pSpWFSamp.HWFNRSH = rs("HWFNRSH") Else pSpWFSamp.HWFNRSH = " "
    If IsNull(rs("HWFNRST")) = False Then pSpWFSamp.HWFNRST = rs("HWFNRST") Else pSpWFSamp.HWFNRST = " "
    If IsNull(rs("HWFNRSI")) = False Then pSpWFSamp.HWFNRSI = rs("HWFNRSI") Else pSpWFSamp.HWFNRSI = " "
    If IsNull(rs("HWFNRKN")) = False Then pSpWFSamp.HWFNRKN = rs("HWFNRKN") Else pSpWFSamp.HWFNRKN = " "
    If IsNull(rs("HWFSPVPUG")) = False Then pSpWFSamp.HWFSPVPUG = rs("HWFSPVPUG") Else pSpWFSamp.HWFSPVPUG = " "
    If IsNull(rs("HWFSPVPUR")) = False Then pSpWFSamp.HWFSPVPUR = rs("HWFSPVPUR") Else pSpWFSamp.HWFSPVPUR = " "
    If IsNull(rs("HWFSPVSTD")) = False Then pSpWFSamp.HWFSPVSTD = rs("HWFSPVSTD") Else pSpWFSamp.HWFSPVSTD = " "
    If IsNull(rs("HWFDLPUG")) = False Then pSpWFSamp.HWFDLPUG = rs("HWFDLPUG") Else pSpWFSamp.HWFDLPUG = " "
    If IsNull(rs("HWFDLPUR")) = False Then pSpWFSamp.HWFDLPUR = rs("HWFDLPUR") Else pSpWFSamp.HWFDLPUR = " "
    If IsNull(rs("HWFNRPUG")) = False Then pSpWFSamp.HWFNRPUG = rs("HWFNRPUG") Else pSpWFSamp.HWFNRPUG = " "
    If IsNull(rs("HWFNRPUR")) = False Then pSpWFSamp.HWFNRPUR = rs("HWFNRPUR") Else pSpWFSamp.HWFNRPUR = " "
    If IsNull(rs("HWFNRSTD")) = False Then pSpWFSamp.HWFNRSTD = rs("HWFNRSTD") Else pSpWFSamp.HWFNRSTD = " "
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
    If IsNull(rs("HWFSIRDMX")) = False Then pSpWFSamp.HWFSIRDMX = rs("HWFSIRDMX") Else pSpWFSamp.HWFSIRDMX = "0"    '軸状転位上限
    If IsNull(rs("HWFSIRDSZ")) = False Then pSpWFSamp.HWFSIRDSZ = rs("HWFSIRDSZ") Else pSpWFSamp.HWFSIRDSZ = " "    '軸状転位測定条件
    If IsNull(rs("HWFSIRDHT")) = False Then pSpWFSamp.HWFSIRDHT = rs("HWFSIRDHT") Else pSpWFSamp.HWFSIRDHT = " "    '軸状転位保証方法＿対
    If IsNull(rs("HWFSIRDHS")) = False Then pSpWFSamp.HWFSIRDHS = rs("HWFSIRDHS") Else pSpWFSamp.HWFSIRDHS = " "    '軸状転位保証方法＿処
    If IsNull(rs("HWFSIRDKM")) = False Then pSpWFSamp.HWFSIRDKM = rs("HWFSIRDKM") Else pSpWFSamp.HWFSIRDKM = " "    '軸状転位検査頻度＿枚
    If IsNull(rs("HWFSIRDKN")) = False Then pSpWFSamp.HWFSIRDKN = rs("HWFSIRDKN") Else pSpWFSamp.HWFSIRDKN = " "    '軸状転位検査頻度＿抜
    If IsNull(rs("HWFSIRDKH")) = False Then pSpWFSamp.HWFSIRDKH = rs("HWFSIRDKH") Else pSpWFSamp.HWFSIRDKH = " "    '軸状転位検査頻度＿保
    If IsNull(rs("HWFSIRDKU")) = False Then pSpWFSamp.HWFSIRDKU = rs("HWFSIRDKU") Else pSpWFSamp.HWFSIRDKU = " "    '軸状転位検査頻度＿ウ
    If IsNull(rs("HWFSIRDPS")) = False Then pSpWFSamp.HWFSIRDPS = Trim(rs("HWFSIRDPS")) Else pSpWFSamp.HWFSIRDPS = " "    '軸状転位TB保証位置
    
    '「軸状転位TB保証位置」を判定し、「軸状転位検査頻度＿抜」に編集（仮対応）
    Select Case Trim(pSpWFSamp.HWFSIRDPS)
    Case "T"
        pSpWFSamp.HWFSIRDKN = "3"
    Case "B"
        pSpWFSamp.HWFSIRDKN = "4"
    Case "TB"
        pSpWFSamp.HWFSIRDKN = "6"
    Case Else
        pSpWFSamp.HWFSIRDKN = " "
    End Select
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
    
    rs.Close
    '' SPV仕様取得　06/06/08 ooba END =================================================>
    
    '' 製品仕様管理の取得
    sql = "select HWFIGKBN from TBCME017" & _
          " where HINBAN='" & pSpWFSamp.hin.hinban & "' and MNOREVNO=" & pSpWFSamp.hin.mnorevno & _
          " and FACTORY='" & pSpWFSamp.hin.factory & "'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo PROC_EXIT
    End If
    pSpWFSamp.HWFIGKBN = rs("HWFIGKBN")
    rs.Close

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    '' エピ仕様取得(BMD1E～BMD3E,OSF1E～OSF3E)
    sql = "select HEPOF1NS, "                   ' 品熱処理法(OSF1E)
    sql = sql & "HEPOF1SZ, "                    ' 品測定条件(OSF1E)
    sql = sql & "HEPOF1ET, "                    ' 品選択ET代(OSF1E)
    sql = sql & "HEPOF1HS, "                    ' 品保証方法_処(OSF1E)
    sql = sql & "HEPOF1SH, "                    ' 品測定位置_方(OSF1E)
    sql = sql & "HEPOF1ST, "                    ' 品測定位置_点(OSF1E)
    sql = sql & "HEPOF1SR, "                    ' 品測定位置_領(OSF1E)
    sql = sql & "HEPOF1KN, "                    ' 品検査頻度_抜(OSF1E)
    sql = sql & "HEPOF2NS, "                    ' 品熱処理法(OSF2E)
    sql = sql & "HEPOF2SZ, "                    ' 品測定条件(OSF2E)
    sql = sql & "HEPOF2ET, "                    ' 品選択ET代(OSF2E)
    sql = sql & "HEPOF2HS, "                    ' 品保証方法_処(OSF2E)
    sql = sql & "HEPOF2SH, "                    ' 品測定位置_方(OSF2E)
    sql = sql & "HEPOF2ST, "                    ' 品測定位置_点(OSF2E)
    sql = sql & "HEPOF2SR, "                    ' 品測定位置_領(OSF2E)
    sql = sql & "HEPOF2KN, "                    ' 品検査頻度_抜(OSF2E)
    sql = sql & "HEPOF3NS, "                    ' 品熱処理法(OSF3E)
    sql = sql & "HEPOF3SZ, "                    ' 品測定条件(OSF3E)
    sql = sql & "HEPOF3ET, "                    ' 品選択ET代(OSF3E)
    sql = sql & "HEPOF3HS, "                    ' 品保証方法_処(OSF3E)
    sql = sql & "HEPOF3SH, "                    ' 品測定位置_方(OSF3E)
    sql = sql & "HEPOF3ST, "                    ' 品測定位置_点(OSF3E)
    sql = sql & "HEPOF3SR, "                    ' 品測定位置_領(OSF3E)
    sql = sql & "HEPOF3KN, "                    ' 品検査頻度_抜(OSF3E)
    sql = sql & "HEPBM1NS, "                    ' 品熱処理法(BMD1E)
    sql = sql & "HEPBM1SZ, "                    ' 品測定条件(BMD1E)
    sql = sql & "HEPBM1ET, "                    ' 品選択ET代(BMD1E)
    sql = sql & "HEPBM1HS, "                    ' 品保証方法_処(BMD1E)
    sql = sql & "HEPBM1SH, "                    ' 品測定位置_方(BMD1E)
    sql = sql & "HEPBM1ST, "                    ' 品測定位置_点(BMD1E)
    sql = sql & "HEPBM1SR, "                    ' 品測定位置_領(BMD1E)
    sql = sql & "HEPBM1KN, "                    ' 品検査頻度_抜(BMD1E)
    sql = sql & "HEPBM2NS, "                    ' 品熱処理法(BMD2E)
    sql = sql & "HEPBM2SZ, "                    ' 品測定条件(BMD2E)
    sql = sql & "HEPBM2ET, "                    ' 品選択ET代(BMD2E)
    sql = sql & "HEPBM2HS, "                    ' 品保証方法_処(BMD2E)
    sql = sql & "HEPBM2SH, "                    ' 品測定位置_方(BMD2E)
    sql = sql & "HEPBM2ST, "                    ' 品測定位置_点(BMD2E)
    sql = sql & "HEPBM2SR, "                    ' 品測定位置_領(BMD2E)
    sql = sql & "HEPBM2KN, "                    ' 品検査頻度_抜(BMD2E)
    sql = sql & "HEPBM3NS, "                    ' 品熱処理法(BMD3E)
    sql = sql & "HEPBM3SZ, "                    ' 品測定条件(BMD3E)
    sql = sql & "HEPBM3ET, "                    ' 品選択ET代(BMD3E)
    sql = sql & "HEPBM3HS, "                    ' 品保証方法_処(BMD3E)
    sql = sql & "HEPBM3SH, "                    ' 品測定位置_方(BMD3E)
    sql = sql & "HEPBM3ST, "                    ' 品測定位置_点(BMD3E)
    sql = sql & "HEPBM3SR, "                    ' 品測定位置_領(BMD3E)
    sql = sql & "HEPBM3KN, "                    ' 品検査頻度_抜(BMD3E)
    sql = sql & "HEPACEN, "                     ' 品E1厚中心
    sql = sql & "HEPANTNP, "                    ' 品EPAN温度
    sql = sql & "HEPANTIM, "                    ' 品EPAN時間
    sql = sql & "HEPIGKBN, "                    ' 品EPIG区分
    sql = sql & "HEPANGZY "                     ' 品EP高温ANガス条件
    sql = sql & "from TBCME050 "
    sql = sql & "where HINBAN = '" & pSpWFSamp.hin.hinban & "' "
    sql = sql & "and MNOREVNO = " & pSpWFSamp.hin.mnorevno & " "
    sql = sql & "and FACTORY = '" & pSpWFSamp.hin.factory & "' "
    sql = sql & "and OPECOND = '" & pSpWFSamp.hin.opecond & "' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        scmzc_getWF = FUNCTION_RETURN_FAILURE
        GoTo PROC_EXIT
    End If
    
    If IsNull(rs("HEPOF1NS")) = False Then pSpWFSamp.HEPOF1NS = rs("HEPOF1NS") Else pSpWFSamp.HEPOF1NS = " "
    If IsNull(rs("HEPOF1SZ")) = False Then pSpWFSamp.HEPOF1SZ = rs("HEPOF1SZ") Else pSpWFSamp.HEPOF1SZ = " "
    pSpWFSamp.HEPOF1ET = fncNullCheck(rs("HEPOF1ET"))
    If IsNull(rs("HEPOF1HS")) = False Then pSpWFSamp.HEPOF1HS = rs("HEPOF1HS") Else pSpWFSamp.HEPOF1HS = " "
    If IsNull(rs("HEPOF1SH")) = False Then pSpWFSamp.HEPOF1SH = rs("HEPOF1SH") Else pSpWFSamp.HEPOF1SH = " "
    If IsNull(rs("HEPOF1ST")) = False Then pSpWFSamp.HEPOF1ST = rs("HEPOF1ST") Else pSpWFSamp.HEPOF1ST = " "
    If IsNull(rs("HEPOF1SR")) = False Then pSpWFSamp.HEPOF1SR = rs("HEPOF1SR") Else pSpWFSamp.HEPOF1SR = " "
    If IsNull(rs("HEPOF1KN")) = False Then pSpWFSamp.HEPOF1KN = rs("HEPOF1KN") Else pSpWFSamp.HEPOF1KN = " "
    If IsNull(rs("HEPOF2NS")) = False Then pSpWFSamp.HEPOF2NS = rs("HEPOF2NS") Else pSpWFSamp.HEPOF2NS = " "
    If IsNull(rs("HEPOF2SZ")) = False Then pSpWFSamp.HEPOF2SZ = rs("HEPOF2SZ") Else pSpWFSamp.HEPOF2SZ = " "
    pSpWFSamp.HEPOF2ET = fncNullCheck(rs("HEPOF2ET"))
    If IsNull(rs("HEPOF2HS")) = False Then pSpWFSamp.HEPOF2HS = rs("HEPOF2HS") Else pSpWFSamp.HEPOF2HS = " "
    If IsNull(rs("HEPOF2SH")) = False Then pSpWFSamp.HEPOF2SH = rs("HEPOF2SH") Else pSpWFSamp.HEPOF2SH = " "
    If IsNull(rs("HEPOF2ST")) = False Then pSpWFSamp.HEPOF2ST = rs("HEPOF2ST") Else pSpWFSamp.HEPOF2ST = " "
    If IsNull(rs("HEPOF2SR")) = False Then pSpWFSamp.HEPOF2SR = rs("HEPOF2SR") Else pSpWFSamp.HEPOF2SR = " "
    If IsNull(rs("HEPOF2KN")) = False Then pSpWFSamp.HEPOF2KN = rs("HEPOF2KN") Else pSpWFSamp.HEPOF2KN = " "
    If IsNull(rs("HEPOF3NS")) = False Then pSpWFSamp.HEPOF3NS = rs("HEPOF3NS") Else pSpWFSamp.HEPOF3NS = " "
    If IsNull(rs("HEPOF3SZ")) = False Then pSpWFSamp.HEPOF3SZ = rs("HEPOF3SZ") Else pSpWFSamp.HEPOF3SZ = " "
    pSpWFSamp.HEPOF3ET = fncNullCheck(rs("HEPOF3ET"))
    If IsNull(rs("HEPOF3HS")) = False Then pSpWFSamp.HEPOF3HS = rs("HEPOF3HS") Else pSpWFSamp.HEPOF3HS = " "
    If IsNull(rs("HEPOF3SH")) = False Then pSpWFSamp.HEPOF3SH = rs("HEPOF3SH") Else pSpWFSamp.HEPOF3SH = " "
    If IsNull(rs("HEPOF3ST")) = False Then pSpWFSamp.HEPOF3ST = rs("HEPOF3ST") Else pSpWFSamp.HEPOF3ST = " "
    If IsNull(rs("HEPOF3SR")) = False Then pSpWFSamp.HEPOF3SR = rs("HEPOF3SR") Else pSpWFSamp.HEPOF3SR = " "
    If IsNull(rs("HEPOF3KN")) = False Then pSpWFSamp.HEPOF3KN = rs("HEPOF3KN") Else pSpWFSamp.HEPOF3KN = " "
    If IsNull(rs("HEPBM1NS")) = False Then pSpWFSamp.HEPBM1NS = rs("HEPBM1NS") Else pSpWFSamp.HEPBM1NS = " "
    If IsNull(rs("HEPBM1SZ")) = False Then pSpWFSamp.HEPBM1SZ = rs("HEPBM1SZ") Else pSpWFSamp.HEPBM1SZ = " "
    pSpWFSamp.HEPBM1ET = fncNullCheck(rs("HEPBM1ET"))
    If IsNull(rs("HEPBM1HS")) = False Then pSpWFSamp.HEPBM1HS = rs("HEPBM1HS") Else pSpWFSamp.HEPBM1HS = " "
    If IsNull(rs("HEPBM1SH")) = False Then pSpWFSamp.HEPBM1SH = rs("HEPBM1SH") Else pSpWFSamp.HEPBM1SH = " "
    If IsNull(rs("HEPBM1ST")) = False Then pSpWFSamp.HEPBM1ST = rs("HEPBM1ST") Else pSpWFSamp.HEPBM1ST = " "
    If IsNull(rs("HEPBM1SR")) = False Then pSpWFSamp.HEPBM1SR = rs("HEPBM1SR") Else pSpWFSamp.HEPBM1SR = " "
    If IsNull(rs("HEPBM1KN")) = False Then pSpWFSamp.HEPBM1KN = rs("HEPBM1KN") Else pSpWFSamp.HEPBM1KN = " "
    If IsNull(rs("HEPBM2NS")) = False Then pSpWFSamp.HEPBM2NS = rs("HEPBM2NS") Else pSpWFSamp.HEPBM2NS = " "
    If IsNull(rs("HEPBM2SZ")) = False Then pSpWFSamp.HEPBM2SZ = rs("HEPBM2SZ") Else pSpWFSamp.HEPBM2SZ = " "
    pSpWFSamp.HEPBM2ET = fncNullCheck(rs("HEPBM2ET"))
    If IsNull(rs("HEPBM2HS")) = False Then pSpWFSamp.HEPBM2HS = rs("HEPBM2HS") Else pSpWFSamp.HEPBM2HS = " "
    If IsNull(rs("HEPBM2SH")) = False Then pSpWFSamp.HEPBM2SH = rs("HEPBM2SH") Else pSpWFSamp.HEPBM2SH = " "
    If IsNull(rs("HEPBM2ST")) = False Then pSpWFSamp.HEPBM2ST = rs("HEPBM2ST") Else pSpWFSamp.HEPBM2ST = " "
    If IsNull(rs("HEPBM2SR")) = False Then pSpWFSamp.HEPBM2SR = rs("HEPBM2SR") Else pSpWFSamp.HEPBM2SR = " "
    If IsNull(rs("HEPBM2KN")) = False Then pSpWFSamp.HEPBM2KN = rs("HEPBM2KN") Else pSpWFSamp.HEPBM2KN = " "
    If IsNull(rs("HEPBM3NS")) = False Then pSpWFSamp.HEPBM3NS = rs("HEPBM3NS") Else pSpWFSamp.HEPBM3NS = " "
    If IsNull(rs("HEPBM3SZ")) = False Then pSpWFSamp.HEPBM3SZ = rs("HEPBM3SZ") Else pSpWFSamp.HEPBM3SZ = " "
    pSpWFSamp.HEPBM3ET = fncNullCheck(rs("HEPBM3ET"))
    If IsNull(rs("HEPBM3HS")) = False Then pSpWFSamp.HEPBM3HS = rs("HEPBM3HS") Else pSpWFSamp.HEPBM3HS = " "
    If IsNull(rs("HEPBM3SH")) = False Then pSpWFSamp.HEPBM3SH = rs("HEPBM3SH") Else pSpWFSamp.HEPBM3SH = " "
    If IsNull(rs("HEPBM3ST")) = False Then pSpWFSamp.HEPBM3ST = rs("HEPBM3ST") Else pSpWFSamp.HEPBM3ST = " "
    If IsNull(rs("HEPBM3SR")) = False Then pSpWFSamp.HEPBM3SR = rs("HEPBM3SR") Else pSpWFSamp.HEPBM3SR = " "
    If IsNull(rs("HEPBM3KN")) = False Then pSpWFSamp.HEPBM3KN = rs("HEPBM3KN") Else pSpWFSamp.HEPBM3KN = " "
    pSpWFSamp.HEPACEN = fncNullCheck(rs("HEPACEN"))
    pSpWFSamp.HEPANTNP = fncNullCheck(rs("HEPANTNP"))
    pSpWFSamp.HEPANTIM = fncNullCheck(rs("HEPANTIM"))
    If IsNull(rs("HEPIGKBN")) = False Then pSpWFSamp.HEPIGKBN = rs("HEPIGKBN") Else pSpWFSamp.HEPIGKBN = " "
    If IsNull(rs("HEPANGZY")) = False Then pSpWFSamp.HEPANGZY = rs("HEPANGZY") Else pSpWFSamp.HEPANGZY = " "
    rs.Close
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

    scmzc_getWF = FUNCTION_RETURN_SUCCESS

PROC_EXIT:
    '' 終了
    gErr.Pop
    Exit Function

PROC_ERR:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    scmzc_getWF = FUNCTION_RETURN_FAILURE
    Resume PROC_EXIT

End Function
'----------------------------------------------------------------------------
'概要      :製品仕様WFデータ（OT１、OT2)の取得ドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型               ,説明
'　　      :pSpWFSamp　　　,IO ,typ_SpWFSamp   　,WFサンプル仕様
'　　      :戻り値         ,O  ,FUNCTION_RETURN　,読み込みの成否
'履歴      :03/05/21 後藤     2004/06/23 変更 その他サンプル枚数取得
'----------------------------------------------------------------------------
Public Function scmzc_getE036(pHin As tFullHinban, strOT1 As String, strOT2 As String, _
                              strMAI1 As String, strMAI2 As String) As FUNCTION_RETURN
    Dim sql     As String
    Dim rs As OraDynaset
    
    '' エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcF_cmkc001WF.bas -- Function scmzc_getE036"
    '--- 2004/06/23
    'sql = "select " & _
          "OTHER1, OTHER2, OTHERTIME" & _
          " from TBCME036" & _
          " where HINBAN ='" & pHin.hinban & "' and MNOREVNO=" & pHin.mnorevno & _
          " and FACTORY ='" & pHin.factory & "' and OPECOND ='" & pHin.opecond & "'" & _
          " and OTHERTIME > sysdate"
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -s-
''    sql = "select " & _
''          "OTHER1, OTHER2, OTHERTIME, OTHER1MAI, OTHER2MAI " & _
''          " from TBCME036" & _
''          " where HINBAN ='" & pHin.hinban & "' and MNOREVNO=" & pHin.mnorevno & _
''          " and FACTORY ='" & pHin.factory & "' and OPECOND ='" & pHin.opecond & "'" & _
''          " and OTHERTIME > sysdate"
    sql = ""
    sql = sql & " SELECT"
    sql = sql & "   a.ot1 AS other1"
    sql = sql & "  ,a.ot1m AS other1mai"
    sql = sql & "  ,b.ot2 AS other2"
    sql = sql & "  ,b.ot2m AS other2mai"
    sql = sql & " FROM"
    sql = sql & "   ("
    sql = sql & "    SELECT"
    sql = sql & "      COUNT(other1)"
    sql = sql & "     ,MAX(other1) AS ot1"
    sql = sql & "     ,MAX(other1mai) AS ot1m"
    sql = sql & "    FROM"
    sql = sql & "      tbcme036"
    sql = sql & "    WHERE hinban   = '" & pHin.hinban & "'"
    sql = sql & "      AND mnorevno = " & pHin.mnorevno
    sql = sql & "      AND factory  = '" & pHin.factory & "'"
    sql = sql & "      AND opecond  = '" & pHin.opecond & "'"
    sql = sql & "      AND othertime > SYSDATE"
    sql = sql & "   ) a"
    sql = sql & "  ,("
    sql = sql & "    SELECT"
    sql = sql & "      COUNT(other2)"
    sql = sql & "     ,MAX(other2) AS ot2"
    sql = sql & "     ,MAX(other2mai) AS ot2m"
    sql = sql & "    FROM"
    sql = sql & "      tbcme036"
    sql = sql & "    WHERE hinban   = '" & pHin.hinban & "'"
    sql = sql & "      AND mnorevno = " & pHin.mnorevno
    sql = sql & "      AND factory  = '" & pHin.factory & "'"
    sql = sql & "      AND opecond  = '" & pHin.opecond & "'"
    sql = sql & "      AND othertime2 > SYSDATE"
    sql = sql & "   ) b"
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -e-
    '---------------
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount = 0 Then
        rs.Close
        strOT1 = "0"
        strOT2 = "0"
        strMAI1 = "0"    '2004/06/23
        strMAI2 = "0"    '2004/06/23
''        scmzc_getE036 = FUNCTION_RETURN_FAILURE
        GoTo PROC_EXIT
    End If
    If IsNull(rs("OTHER1")) = True Then
        strOT1 = "0"
    Else
        strOT1 = rs("OTHER1")
    End If
    If IsNull(rs("OTHER2")) = True Then
        strOT2 = "0"
    Else
        strOT2 = rs("OTHER2")
    End If
    '----- 2004/06/23
    If IsNull(rs("OTHER1MAI")) = True Then
        strMAI1 = "0"
    Else
        strMAI1 = rs("OTHER1MAI")
    End If
    If IsNull(rs("OTHER2MAI")) = True Then
        strMAI2 = "0"
    Else
        strMAI2 = rs("OTHER2MAI")
    End If
    '-----------------
    
    scmzc_getE036 = FUNCTION_RETURN_SUCCESS
    rs.Close
    
PROC_EXIT:
    '' 終了
    gErr.Pop
    Exit Function
    
PROC_ERR:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    scmzc_getE036 = FUNCTION_RETURN_FAILURE
    Resume PROC_EXIT
    
End Function


'概要      :酸素析出と残存酸素の仕様チェック
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型              ,説明
'      　　:pHin　　    　,I  ,tFullHinban   　,品番
'      　　:戻り値        ,O  ,Integer       　,仕様チェック結果(-1:ｴﾗｰ，0:AOi仕様無，1:AOi仕様有)
'説明      :酸素析出(Δoi)と残存酸素の両方に仕様が立っていた場合エラーを返す
'履歴      :03/12/05 ooba

Public Function ChkAoiSiyou(pHin As tFullHinban) As Integer

    Dim sSql As String
    Dim rs As OraDynaset
    Dim sDoiSiyou(2) As String  '検査有無(DOi1～3)
    Dim sAoiSiyou As String     '検査有無(AOi)
    Dim iCnt As Integer
    
    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcF_cmkc001WF.bas -- Function ChkAoiSiyou"

    sSql = "select HWFOS1HS, HWFOS2HS, HWFOS3HS, HWFZOHWS from TBCME025 "
    sSql = sSql & "where HINBAN = '" & pHin.hinban & "' "
    sSql = sSql & "and MNOREVNO = " & pHin.mnorevno & " "
    sSql = sSql & "and FACTORY = '" & pHin.factory & "' "
    sSql = sSql & "and OPECOND = '" & pHin.opecond & "' "

    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount = 0 Then
        rs.Close
        ChkAoiSiyou = -1
        GoTo PROC_EXIT
    End If
    
    If IsNull(rs("HWFOS1HS")) = False Then sDoiSiyou(0) = rs("HWFOS1HS") '品WF酸素析出1保証方法_処
    If IsNull(rs("HWFOS2HS")) = False Then sDoiSiyou(1) = rs("HWFOS2HS") '品WF酸素析出2保証方法_処
    If IsNull(rs("HWFOS3HS")) = False Then sDoiSiyou(2) = rs("HWFOS3HS") '品WF酸素析出3保証方法_処
    If IsNull(rs("HWFZOHWS")) = False Then sAoiSiyou = rs("HWFZOHWS")    '品WF残存酸素保証方法_処
    
'--------------- 2008/07/25 INSERT START  By Systech ---------------
    rs.Close
'--------------- 2008/07/25 INSERT  END   By Systech ---------------
    
    '酸素析出と残存酸素の仕様チェック
    ChkAoiSiyou = 0
    For iCnt = 0 To 2
        If sDoiSiyou(iCnt) = "H" Or sDoiSiyou(iCnt) = "S" Then
            '酸素析出(Δoi)と残存酸素の両方に仕様が立っていた場合はエラー
            If sAoiSiyou = "H" Or sAoiSiyou = "S" Then
                ChkAoiSiyou = -1
                Exit For
            End If
        Else
            If sAoiSiyou = "H" Or sAoiSiyou = "S" Then
                ChkAoiSiyou = 1
            End If
        End If
    Next
    
PROC_EXIT:
    '' 終了
    gErr.Pop
    Exit Function
    
PROC_ERR:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    ChkAoiSiyou = -1
    Resume PROC_EXIT
    
End Function


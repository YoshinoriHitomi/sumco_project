Attribute VB_Name = "s_cmzcSPWF"
Option Explicit
'                                     2001/07/04
'===============================================================================
' ＷＦ所要サンプル判定関数
' 概要    :
'===============================================================================

Private tblSampUP As typ_SpWFSamp
Private tblSampDN As typ_SpWFSamp

'概要      :ＷＦサンプルの取得（１ブロック分）
'ﾊﾟﾗﾒｰﾀ　　:変数名          ,IO ,型             ,説明
'　　      :pHinUp    　　　,I  ,tFullHinban  　,上品番テーブル
'　　      :pHinDn    　　　,I  ,tFullHinban  　,下品番テーブル
'　　      :pWFSample　　　 ,O  ,typ_WFSample　 ,ＷＦサンプルテーブル
'説明      :検査指示サンプルデータを取得する
'履歴      :2001/07/04　大塚 作成
Public Sub GetWFSampAll(pHinUp As tFullHinban, pHinDn As tFullHinban, pWFSample As typ_WFSample)

    '' 検査指示サンプルの取得
    With pWFSample
        .CRYINDRS = GetWFSamp(pHinUp, pHinDn, 1)
        .CRYINDOI = GetWFSamp(pHinUp, pHinDn, 2)
        .CRYINDB1 = GetWFSamp(pHinUp, pHinDn, 3)
        .CRYINDB2 = GetWFSamp(pHinUp, pHinDn, 4)
        .CRYINDB3 = GetWFSamp(pHinUp, pHinDn, 5)
        .CRYINDL1 = GetWFSamp(pHinUp, pHinDn, 6)
        .CRYINDL2 = GetWFSamp(pHinUp, pHinDn, 7)
        .CRYINDL3 = GetWFSamp(pHinUp, pHinDn, 8)
        .CRYINDL4 = GetWFSamp(pHinUp, pHinDn, 9)
        .CRYINDDS = GetWFSamp(pHinUp, pHinDn, 10)
        .CRYINDDZ = GetWFSamp(pHinUp, pHinDn, 11)
        .CRYINDSP = GetWFSamp(pHinUp, pHinDn, 12)
        .CRYINDD1 = GetWFSamp(pHinUp, pHinDn, 13)
        .CRYINDD2 = GetWFSamp(pHinUp, pHinDn, 14)
        .CRYINDD3 = GetWFSamp(pHinUp, pHinDn, 15)
        .CRYINDAO = GetWFSamp(pHinUp, pHinDn, 18)   '残存酸素追加　03/12/05 ooba
        .CRYINDGD = GetWFSamp(pHinUp, pHinDn, 19)   'GD追加　05/01/18 ooba
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
        .EPIINDB1 = GetWFSamp(pHinUp, pHinDn, 20)
        .EPIINDB2 = GetWFSamp(pHinUp, pHinDn, 21)
        .EPIINDB3 = GetWFSamp(pHinUp, pHinDn, 22)
        .EPIINDL1 = GetWFSamp(pHinUp, pHinDn, 23)
        .EPIINDL2 = GetWFSamp(pHinUp, pHinDn, 24)
        .EPIINDL3 = GetWFSamp(pHinUp, pHinDn, 25)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    End With

End Sub

'概要      :ＷＦサンプルの取得
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型           ,説明
'　　      :pHinUp 　　　,I  ,tFullHinban　,上品番テーブル
'　　      :pHinDn 　　　,I  ,tFullHinban　,下品番テーブル
'　　      :iCol   　　　,I  ,Integer      ,列
'　　      :戻り値       ,O  ,String       ,検査指示サンプル
'説明      :検査指示サンプルデータを取得する
'履歴      :2001/07/04　大塚 作成
'          :2004/04/08 ooba　WF抜試指示変更(保証方法ﾁｪｯｸの追加)
Public Function GetWFSamp(pHinUp As tFullHinban, pHinDn As tFullHinban, iCol As Integer) As String

    Dim HINBANUP As String
    Dim HINBANDN As String
    Dim iMode As Integer
    Dim m As Integer
    Dim i As Integer
    '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
    '2.1.2 共有チェック追加
    Dim liRet       As Integer
    Dim lsResult    As String       'コードＤＢ取得関数の取得変数
    Dim llCnt           As Long
    Dim lsCode(1)       As String
    Dim liLoopCnt       As Integer
    '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------

    '' 上品番／下品番が共に空でなければ
    HINBANUP = Trim(pHinUp.hinban)
    HINBANDN = Trim(pHinDn.hinban)
    If (HINBANUP = "" Or HINBANUP = "G" Or HINBANUP = "Z") And _
       (HINBANDN = "" Or HINBANDN = "G" Or HINBANDN = "Z") Then
        GetWFSamp = "0"
        Exit Function
    End If

    '' 上品番／下品番状態の分類
    If HINBANUP = "" Or HINBANUP = "G" Or HINBANUP = "Z" Then
        iMode = 1
    ElseIf HINBANDN = "" Or HINBANDN = "G" Or HINBANDN = "Z" Then
        iMode = 2
    ElseIf HINBANUP = HINBANDN Then
        iMode = 3
    Else
        iMode = 4
    End If

    '' 上品番の製品仕様データを取得
    If iMode <> 1 Then
'        If tblSampUP.HIN.hinban <> pHinUp.hinban Then
        '最新品番ﾃﾞｰﾀ取得対応　04/05/24 ooba
        If Not (tblSampUP.hin.hinban = pHinUp.hinban _
                And tblSampUP.hin.mnorevno = pHinUp.mnorevno _
                And tblSampUP.hin.factory = pHinUp.factory _
                And tblSampUP.hin.opecond = pHinUp.opecond) Then
           tblSampUP.hin.hinban = pHinUp.hinban
           tblSampUP.hin.mnorevno = pHinUp.mnorevno
           tblSampUP.hin.factory = pHinUp.factory
           tblSampUP.hin.opecond = pHinUp.opecond
            If scmzc_getWF(tblSampUP) = FUNCTION_RETURN_FAILURE Then
                GetWFSamp = "0"
                Exit Function
            End If
        End If
    End If

    '' 下品番の製品仕様データを取得
    If iMode <> 2 Then
'        If tblSampDN.HIN.hinban <> pHinDn.hinban Then
        '最新品番ﾃﾞｰﾀ取得対応　04/05/24 ooba
        If Not (tblSampDN.hin.hinban = pHinDn.hinban _
                And tblSampDN.hin.mnorevno = pHinDn.mnorevno _
                And tblSampDN.hin.factory = pHinDn.factory _
                And tblSampDN.hin.opecond = pHinDn.opecond) Then
           tblSampDN.hin.hinban = pHinDn.hinban
           tblSampDN.hin.mnorevno = pHinDn.mnorevno
           tblSampDN.hin.factory = pHinDn.factory
           tblSampDN.hin.opecond = pHinDn.opecond
            If scmzc_getWF(tblSampDN) = FUNCTION_RETURN_FAILURE Then
                GetWFSamp = "0"
                Exit Function
            End If
        End If
    End If

    '' 上品番／下品番状態分岐
    Select Case iMode
    Case 1      '' 上品番なし
        Select Case iCol
        Case 1      'Rs
            GetWFSamp = IIf(CheckHWS(tblSampDN.HWFRHWYS) And CheckKHN(tblSampDN.HWFRKHNN, 1, "TOP"), "1", "0")
        Case 2      'Oi
            GetWFSamp = IIf(CheckHWS(tblSampDN.HWFONHWS) And CheckKHN(tblSampDN.HWFONKHN, 2, "TOP"), "1", "0")
        Case 3      'B1
            GetWFSamp = IIf(CheckHWS(tblSampDN.HWFBM1HS) And CheckKHN(tblSampDN.HWFBM1KN, 7, "TOP"), "1", "0")
        Case 4      'B2
            GetWFSamp = IIf(CheckHWS(tblSampDN.HWFBM2HS) And CheckKHN(tblSampDN.HWFBM2KN, 8, "TOP"), "1", "0")
        Case 5      'B3
            GetWFSamp = IIf(CheckHWS(tblSampDN.HWFBM3HS) And CheckKHN(tblSampDN.HWFBM3KN, 9, "TOP"), "1", "0")
        Case 6      'L1
            GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOF1HS) And CheckKHN(tblSampDN.HWFOF1KN, 3, "TOP"), "1", "0")
        Case 7      'L2
            GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOF2HS) And CheckKHN(tblSampDN.HWFOF2KN, 4, "TOP"), "1", "0")
        Case 8      'L3
            GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOF3HS) And CheckKHN(tblSampDN.HWFOF3KN, 5, "TOP"), "1", "0")
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''        Case 9      'L4
'''            GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOF4HS) And CheckKHN(tblSampDN.HWFOF4KN, 6, "TOP"), "1", "0")
        
        Case 9      'SIRD
            GetWFSamp = IIf(CheckHWS(tblSampDN.HWFSIRDHS) And CheckKHN(tblSampDN.HWFSIRDKN, 6, "TOP"), "1", "0")
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
        Case 10     'DS
            GetWFSamp = IIf(CheckHWS(tblSampDN.HWFDSOHS) And CheckKHN(tblSampDN.HWFDSOKN, 13, "TOP"), "1", "0")
        Case 11     'DZ
            GetWFSamp = IIf(CheckHWS(tblSampDN.HWFMKHWS) And CheckKHN(tblSampDN.HWFMKKHN, 14, "TOP"), "1", "0")
        Case 12     'SP     'Nr濃度追加　06/06/08 ooba
            GetWFSamp = IIf((CheckHWS(tblSampDN.HWFSPVHS) And CheckKHN(tblSampDN.HWFSPVKN, 15, "TOP")) _
                            Or (CheckHWS(tblSampDN.HWFDLHWS) And CheckKHN(tblSampDN.HWFDLKHN, 16, "TOP")) _
                            Or (CheckHWS(tblSampDN.HWFNRHS) And CheckKHN(tblSampDN.HWFNRKN, 19, "TOP")), "1", "0")
        Case 13     'D1
            GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOS1HS) And CheckKHN(tblSampDN.HWFOS1KN, 10, "TOP"), "1", "0")
        Case 14     'D2
            GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOS2HS) And CheckKHN(tblSampDN.HWFOS2KN, 11, "TOP"), "1", "0")
        Case 15     'D3
            GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOS3HS) And CheckKHN(tblSampDN.HWFOS3KN, 12, "TOP"), "1", "0")
        Case 16
            GetWFSamp = IIf(CheckOT(tblSampDN.HWOTHER1), "1", "0")  ''03/05/21
        Case 17
            GetWFSamp = IIf(CheckOT(tblSampDN.HWOTHER2), "1", "0")
        Case 18     'AO
            GetWFSamp = IIf(CheckHWS(tblSampDN.HWFZOHWS) And CheckKHN(tblSampDN.HWFZOKHN, 17, "TOP"), "1", "0") ''残存酸素追加　03/12/05 ooba
        Case 19     'GD　05/01/18 ooba
            GetWFSamp = IIf((CheckHWS(tblSampDN.HWFDENHS) Or _
                             CheckHWS(tblSampDN.HWFLDLHS) Or _
                             CheckHWS(tblSampDN.HWFDVDHS)) And _
                            CheckKHN(tblSampDN.HWFGDKHN, 18, "TOP"), "1", "0")
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
        Case 20     'BMD1E
            GetWFSamp = IIf(CheckHWS(tblSampDN.HEPBM1HS) And CheckKHN_EP(tblSampDN.HEPBM1KN, 1, "TOP"), "1", "0")
        Case 21     'BMD2E
            GetWFSamp = IIf(CheckHWS(tblSampDN.HEPBM2HS) And CheckKHN_EP(tblSampDN.HEPBM2KN, 2, "TOP"), "1", "0")
        Case 22     'BMD3E
            GetWFSamp = IIf(CheckHWS(tblSampDN.HEPBM3HS) And CheckKHN_EP(tblSampDN.HEPBM3KN, 3, "TOP"), "1", "0")
        Case 23     'OSF1E
            GetWFSamp = IIf(CheckHWS(tblSampDN.HEPOF1HS) And CheckKHN_EP(tblSampDN.HEPOF1KN, 4, "TOP"), "1", "0")
        Case 24     'OSF2E
            GetWFSamp = IIf(CheckHWS(tblSampDN.HEPOF2HS) And CheckKHN_EP(tblSampDN.HEPOF2KN, 5, "TOP"), "1", "0")
        Case 25     'OSF3E
            GetWFSamp = IIf(CheckHWS(tblSampDN.HEPOF3HS) And CheckKHN_EP(tblSampDN.HEPOF3KN, 6, "TOP"), "1", "0")
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
        End Select
    Case 2      '' 下品番なし
        Select Case iCol
        Case 1      'Rs
            GetWFSamp = IIf(CheckHWS(tblSampUP.HWFRHWYS) And CheckKHN(tblSampUP.HWFRKHNN, 1, "BOT"), "2", "0")
        Case 2      'Oi
            GetWFSamp = IIf(CheckHWS(tblSampUP.HWFONHWS) And CheckKHN(tblSampUP.HWFONKHN, 2, "BOT"), "2", "0")
        Case 3      'B1
            GetWFSamp = IIf(CheckHWS(tblSampUP.HWFBM1HS) And CheckKHN(tblSampUP.HWFBM1KN, 7, "BOT"), "2", "0")
        Case 4      'B2
            GetWFSamp = IIf(CheckHWS(tblSampUP.HWFBM2HS) And CheckKHN(tblSampUP.HWFBM2KN, 8, "BOT"), "2", "0")
        Case 5      'B3
            GetWFSamp = IIf(CheckHWS(tblSampUP.HWFBM3HS) And CheckKHN(tblSampUP.HWFBM3KN, 9, "BOT"), "2", "0")
        Case 6      'L1
            GetWFSamp = IIf(CheckHWS(tblSampUP.HWFOF1HS) And CheckKHN(tblSampUP.HWFOF1KN, 3, "BOT"), "2", "0")
        Case 7      'L2
            GetWFSamp = IIf(CheckHWS(tblSampUP.HWFOF2HS) And CheckKHN(tblSampUP.HWFOF2KN, 4, "BOT"), "2", "0")
        Case 8      'L3
            GetWFSamp = IIf(CheckHWS(tblSampUP.HWFOF3HS) And CheckKHN(tblSampUP.HWFOF3KN, 5, "BOT"), "2", "0")
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''        Case 9      'L4
'''            GetWFSamp = IIf(CheckHWS(tblSampUP.HWFOF4HS) And CheckKHN(tblSampUP.HWFOF4KN, 6, "BOT"), "2", "0")

        Case 9      'SIRD
            GetWFSamp = IIf(CheckHWS(tblSampUP.HWFSIRDHS) And CheckKHN(tblSampUP.HWFSIRDKN, 6, "BOT"), "2", "0")
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
        Case 10     'DS
            GetWFSamp = IIf(CheckHWS(tblSampUP.HWFDSOHS) And CheckKHN(tblSampUP.HWFDSOKN, 13, "BOT"), "2", "0")
        Case 11     'DZ
            GetWFSamp = IIf(CheckHWS(tblSampUP.HWFMKHWS) And CheckKHN(tblSampUP.HWFMKKHN, 14, "BOT"), "2", "0")
        Case 12     'SP     'Nr濃度追加　06/06/08 ooba
            GetWFSamp = IIf((CheckHWS(tblSampUP.HWFSPVHS) And CheckKHN(tblSampUP.HWFSPVKN, 15, "BOT")) _
                            Or (CheckHWS(tblSampUP.HWFDLHWS) And CheckKHN(tblSampUP.HWFDLKHN, 16, "BOT")) _
                            Or (CheckHWS(tblSampUP.HWFNRHS) And CheckKHN(tblSampUP.HWFNRKN, 19, "BOT")), "2", "0")
        Case 13     'D1
            GetWFSamp = IIf(CheckHWS(tblSampUP.HWFOS1HS) And CheckKHN(tblSampUP.HWFOS1KN, 10, "BOT"), "2", "0")
        Case 14     'D2
            GetWFSamp = IIf(CheckHWS(tblSampUP.HWFOS2HS) And CheckKHN(tblSampUP.HWFOS2KN, 11, "BOT"), "2", "0")
        Case 15     'D3
            GetWFSamp = IIf(CheckHWS(tblSampUP.HWFOS3HS) And CheckKHN(tblSampUP.HWFOS3KN, 12, "BOT"), "2", "0")
        Case 16
            GetWFSamp = IIf(CheckOT(tblSampUP.HWOTHER1), "2", "0")  '03/05/21
        Case 17
            GetWFSamp = IIf(CheckOT(tblSampUP.HWOTHER2), "2", "0")
        Case 18     'AO
            GetWFSamp = IIf(CheckHWS(tblSampUP.HWFZOHWS) And CheckKHN(tblSampUP.HWFZOKHN, 17, "BOT"), "2", "0") ''残存酸素追加　03/12/05 ooba
        Case 19     'GD　05/01/18 ooba
            GetWFSamp = IIf((CheckHWS(tblSampUP.HWFDENHS) Or _
                             CheckHWS(tblSampUP.HWFLDLHS) Or _
                             CheckHWS(tblSampUP.HWFDVDHS)) And _
                            CheckKHN(tblSampUP.HWFGDKHN, 18, "BOT"), "2", "0")
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
        Case 20     'BMD1E
            GetWFSamp = IIf(CheckHWS(tblSampUP.HEPBM1HS) And CheckKHN_EP(tblSampUP.HEPBM1KN, 1, "BOT"), "2", "0")
        Case 21     'BMD2E
            GetWFSamp = IIf(CheckHWS(tblSampUP.HEPBM2HS) And CheckKHN_EP(tblSampUP.HEPBM2KN, 2, "BOT"), "2", "0")
        Case 22     'BMD3E
            GetWFSamp = IIf(CheckHWS(tblSampUP.HEPBM3HS) And CheckKHN_EP(tblSampUP.HEPBM3KN, 3, "BOT"), "2", "0")
        Case 23     'OSF1E
            GetWFSamp = IIf(CheckHWS(tblSampUP.HEPOF1HS) And CheckKHN_EP(tblSampUP.HEPOF1KN, 4, "BOT"), "2", "0")
        Case 24     'OSF2E
            GetWFSamp = IIf(CheckHWS(tblSampUP.HEPOF2HS) And CheckKHN_EP(tblSampUP.HEPOF2KN, 5, "BOT"), "2", "0")
        Case 25     'OSF3E
            GetWFSamp = IIf(CheckHWS(tblSampUP.HEPOF3HS) And CheckKHN_EP(tblSampUP.HEPOF3KN, 6, "BOT"), "2", "0")
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
        End Select
    Case 3      '' 上品番＝下品番
        Select Case iCol
        Case 1      'Rs
            If CheckHWS(tblSampUP.HWFRHWYS) And CheckKHN(tblSampUP.HWFRKHNN, 1, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFRHWYS) And CheckKHN(tblSampDN.HWFRKHNN, 1, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFRHWYS) And CheckKHN(tblSampDN.HWFRKHNN, 1, "TOP"), "1", "0")
            End If
        Case 2      'Oi
            If CheckHWS(tblSampUP.HWFONHWS) And CheckKHN(tblSampUP.HWFONKHN, 2, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFONHWS) And CheckKHN(tblSampDN.HWFONKHN, 2, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFONHWS) And CheckKHN(tblSampDN.HWFONKHN, 2, "TOP"), "1", "0")
            End If
        Case 3      'B1
            If CheckHWS(tblSampUP.HWFBM1HS) And CheckKHN(tblSampUP.HWFBM1KN, 7, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFBM1HS) And CheckKHN(tblSampDN.HWFBM1KN, 7, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFBM1HS) And CheckKHN(tblSampDN.HWFBM1KN, 7, "TOP"), "1", "0")
            End If
        Case 4      'B2
            If CheckHWS(tblSampUP.HWFBM2HS) And CheckKHN(tblSampUP.HWFBM2KN, 8, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFBM2HS) And CheckKHN(tblSampDN.HWFBM2KN, 8, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFBM2HS) And CheckKHN(tblSampDN.HWFBM2KN, 8, "TOP"), "1", "0")
            End If
        Case 5      'B3
            If CheckHWS(tblSampUP.HWFBM3HS) And CheckKHN(tblSampUP.HWFBM3KN, 9, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFBM3HS) And CheckKHN(tblSampDN.HWFBM3KN, 9, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFBM3HS) And CheckKHN(tblSampDN.HWFBM3KN, 9, "TOP"), "1", "0")
            End If
        Case 6      'L1
            If CheckHWS(tblSampUP.HWFOF1HS) And CheckKHN(tblSampUP.HWFOF1KN, 3, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOF1HS) And CheckKHN(tblSampDN.HWFOF1KN, 3, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOF1HS) And CheckKHN(tblSampDN.HWFOF1KN, 3, "TOP"), "1", "0")
            End If
        Case 7      'L2
            If CheckHWS(tblSampUP.HWFOF2HS) And CheckKHN(tblSampUP.HWFOF2KN, 4, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOF2HS) And CheckKHN(tblSampDN.HWFOF2KN, 4, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOF2HS) And CheckKHN(tblSampDN.HWFOF2KN, 4, "TOP"), "1", "0")
            End If
        Case 8      'L3
            If CheckHWS(tblSampUP.HWFOF3HS) And CheckKHN(tblSampUP.HWFOF3KN, 5, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOF3HS) And CheckKHN(tblSampDN.HWFOF3KN, 5, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOF3HS) And CheckKHN(tblSampDN.HWFOF3KN, 5, "TOP"), "1", "0")
            End If
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''        Case 9      'L4
'''            If CheckHWS(tblSampUP.HWFOF4HS) And CheckKHN(tblSampUP.HWFOF4KN, 6, "BOT") Then
'''                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOF4HS) And CheckKHN(tblSampDN.HWFOF4KN, 6, "TOP"), "3", "2")
'''            Else
'''                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOF4HS) And CheckKHN(tblSampDN.HWFOF4KN, 6, "TOP"), "1", "0")
'''            End If

        Case 9      'SIRD
            If CheckHWS(tblSampUP.HWFSIRDHS) And CheckKHN(tblSampUP.HWFSIRDKN, 6, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFSIRDHS) And CheckKHN(tblSampDN.HWFSIRDKN, 6, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFSIRDHS) And CheckKHN(tblSampDN.HWFSIRDKN, 6, "TOP"), "1", "0")
            End If
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
        Case 10     'DS
            If CheckHWS(tblSampUP.HWFDSOHS) And CheckKHN(tblSampUP.HWFDSOKN, 13, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFDSOHS) And CheckKHN(tblSampDN.HWFDSOKN, 13, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFDSOHS) And CheckKHN(tblSampDN.HWFDSOKN, 13, "TOP"), "1", "0")
            End If
        Case 11     'DZ
            If CheckHWS(tblSampUP.HWFMKHWS) And CheckKHN(tblSampUP.HWFMKKHN, 14, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFMKHWS) And CheckKHN(tblSampDN.HWFMKKHN, 14, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFMKHWS) And CheckKHN(tblSampDN.HWFMKKHN, 14, "TOP"), "1", "0")
            End If
        Case 12     'SP     'Nr濃度追加　06/06/08 ooba
            If (CheckHWS(tblSampUP.HWFSPVHS) And CheckKHN(tblSampUP.HWFSPVKN, 15, "BOT")) _
                Or (CheckHWS(tblSampUP.HWFDLHWS) And CheckKHN(tblSampUP.HWFDLKHN, 16, "BOT")) _
                Or (CheckHWS(tblSampUP.HWFNRHS) And CheckKHN(tblSampUP.HWFNRKN, 19, "BOT")) Then
                GetWFSamp = IIf((CheckHWS(tblSampDN.HWFSPVHS) And CheckKHN(tblSampDN.HWFSPVKN, 15, "TOP")) _
                                Or (CheckHWS(tblSampDN.HWFDLHWS) And CheckKHN(tblSampDN.HWFDLKHN, 16, "TOP")) _
                                Or (CheckHWS(tblSampDN.HWFNRHS) And CheckKHN(tblSampDN.HWFNRKN, 19, "TOP")), "3", "2")
            Else
                GetWFSamp = IIf((CheckHWS(tblSampDN.HWFSPVHS) And CheckKHN(tblSampDN.HWFSPVKN, 15, "TOP")) _
                                Or (CheckHWS(tblSampDN.HWFDLHWS) And CheckKHN(tblSampDN.HWFDLKHN, 16, "TOP")) _
                                Or (CheckHWS(tblSampDN.HWFNRHS) And CheckKHN(tblSampDN.HWFNRKN, 19, "TOP")), "1", "0")
            End If
        Case 13     'D1
            If CheckHWS(tblSampUP.HWFOS1HS) And CheckKHN(tblSampUP.HWFOS1KN, 10, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOS1HS) And CheckKHN(tblSampDN.HWFOS1KN, 10, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOS1HS) And CheckKHN(tblSampDN.HWFOS1KN, 10, "TOP"), "1", "0")
            End If
        Case 14     'D2
            If CheckHWS(tblSampUP.HWFOS2HS) And CheckKHN(tblSampUP.HWFOS2KN, 11, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOS2HS) And CheckKHN(tblSampDN.HWFOS2KN, 11, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOS2HS) And CheckKHN(tblSampDN.HWFOS2KN, 11, "TOP"), "1", "0")
            End If
        Case 15     'D3
            If CheckHWS(tblSampUP.HWFOS3HS) And CheckKHN(tblSampUP.HWFOS3KN, 12, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOS3HS) And CheckKHN(tblSampDN.HWFOS3KN, 12, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOS3HS) And CheckKHN(tblSampDN.HWFOS3KN, 12, "TOP"), "1", "0")
            End If
        Case 16
            If CheckOT(tblSampUP.HWOTHER1) Then
'                GetWFSamp = IIf(CheckOT(tblSampDN.HWOTHER1), "3", "2")  '03/05/21
                GetWFSamp = IIf(CheckOT(tblSampDN.HWOTHER1), "4", "2")  '必ず実測を立てる　04/05/24 ooba
            Else
                GetWFSamp = IIf(CheckOT(tblSampDN.HWOTHER1), "1", "0")
            End If
        Case 17
            If CheckOT(tblSampUP.HWOTHER2) Then
'                GetWFSamp = IIf(CheckOT(tblSampDN.HWOTHER2), "3", "2")  '03/05/21
                GetWFSamp = IIf(CheckOT(tblSampDN.HWOTHER2), "4", "2")  '必ず実測を立てる　04/05/24 ooba
            Else
                GetWFSamp = IIf(CheckOT(tblSampDN.HWOTHER2), "1", "0")
            End If
        Case 18     'AO                                             ''残存酸素追加　03/12/05 ooba
            If CheckHWS(tblSampUP.HWFZOHWS) And CheckKHN(tblSampUP.HWFZOKHN, 17, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFZOHWS) And CheckKHN(tblSampDN.HWFZOKHN, 17, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFZOHWS) And CheckKHN(tblSampDN.HWFZOKHN, 17, "TOP"), "1", "0")
            End If
        Case 19     'GD　05/01/18 ooba
            If (CheckHWS(tblSampUP.HWFDENHS) Or _
                CheckHWS(tblSampUP.HWFLDLHS) Or _
                CheckHWS(tblSampUP.HWFDVDHS)) And _
               CheckKHN(tblSampUP.HWFGDKHN, 18, "BOT") Then
               
                GetWFSamp = IIf((CheckHWS(tblSampDN.HWFDENHS) Or _
                                 CheckHWS(tblSampDN.HWFLDLHS) Or _
                                 CheckHWS(tblSampDN.HWFDVDHS)) And _
                                CheckKHN(tblSampDN.HWFGDKHN, 18, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf((CheckHWS(tblSampDN.HWFDENHS) Or _
                                 CheckHWS(tblSampDN.HWFLDLHS) Or _
                                 CheckHWS(tblSampDN.HWFDVDHS)) And _
                                CheckKHN(tblSampDN.HWFGDKHN, 18, "TOP"), "1", "0")
            End If
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
        Case 20     'BMD1E
            If CheckHWS(tblSampUP.HEPBM1HS) And CheckKHN_EP(tblSampUP.HEPBM1KN, 1, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HEPBM1HS) And CheckKHN_EP(tblSampDN.HEPBM1KN, 1, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HEPBM1HS) And CheckKHN_EP(tblSampDN.HEPBM1KN, 1, "TOP"), "1", "0")
            End If
        Case 21     'BMD2E
            If CheckHWS(tblSampUP.HEPBM2HS) And CheckKHN_EP(tblSampUP.HEPBM2KN, 2, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HEPBM2HS) And CheckKHN_EP(tblSampDN.HEPBM2KN, 2, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HEPBM2HS) And CheckKHN_EP(tblSampDN.HEPBM2KN, 2, "TOP"), "1", "0")
            End If
        Case 22     'BMD3E
            If CheckHWS(tblSampUP.HEPBM3HS) And CheckKHN_EP(tblSampUP.HEPBM3KN, 3, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HEPBM3HS) And CheckKHN_EP(tblSampDN.HEPBM3KN, 3, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HEPBM3HS) And CheckKHN_EP(tblSampDN.HEPBM3KN, 3, "TOP"), "1", "0")
            End If
        Case 23     'OSF1E
            If CheckHWS(tblSampUP.HEPOF1HS) And CheckKHN_EP(tblSampUP.HEPOF1KN, 4, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HEPOF1HS) And CheckKHN_EP(tblSampDN.HEPOF1KN, 4, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HEPOF1HS) And CheckKHN_EP(tblSampDN.HEPOF1KN, 4, "TOP"), "1", "0")
            End If
        Case 24     'OSF2E
            If CheckHWS(tblSampUP.HEPOF2HS) And CheckKHN_EP(tblSampUP.HEPOF2KN, 5, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HEPOF2HS) And CheckKHN_EP(tblSampDN.HEPOF2KN, 5, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HEPOF2HS) And CheckKHN_EP(tblSampDN.HEPOF2KN, 5, "TOP"), "1", "0")
            End If
        Case 25     'OSF3E
            If CheckHWS(tblSampUP.HEPOF3HS) And CheckKHN_EP(tblSampUP.HEPOF3KN, 6, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HEPOF3HS) And CheckKHN_EP(tblSampDN.HEPOF3KN, 6, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HEPOF3HS) And CheckKHN_EP(tblSampDN.HEPOF3KN, 6, "TOP"), "1", "0")
            End If
''GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
'        Case 26
'                If Trim(tblSampUP.HWFGDSZY) = "G" And Trim(tblSampDN.HWFGDSZY) <> "G" Then
'                    GetWFSamp = "1"
'                Else
'                    GetWFSamp = "2"
'                End If
''GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
            
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

        End Select
    Case 4      '' 上品番＜＞下品番
        Select Case iCol
        Case 1      'Rs
            If CheckHWS(tblSampUP.HWFRHWYS) And CheckKHN(tblSampUP.HWFRKHNN, 1, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFRHWYS) And CheckKHN(tblSampDN.HWFRKHNN, 1, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFRHWYS) And CheckKHN(tblSampDN.HWFRKHNN, 1, "TOP"), "1", "0")
            End If
            
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '2.1.2 共有チェック追加
            '3:両方検査(共通)の場合、AN温度チェックを行い、NGの時は4:両方検査(別)にする
            If GetWFSamp = "3" Then
                liRet = funCodeDBGetMatrixReturn("SB", "AR", CStr(tblSampDN.HWFANTNP), CStr(tblSampUP.HWFANTNP))
                If liRet = -1 Then
'                    GetWFSamp = "4"
                ElseIf liRet = 0 Then
                    GetWFSamp = "4"
                End If
            End If
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        
        Case 2      'Oi
            If Not (CheckHWS(tblSampUP.HWFONHWS) And CheckKHN(tblSampUP.HWFONKHN, 2, "BOT")) Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFONHWS) And CheckKHN(tblSampDN.HWFONKHN, 2, "TOP"), "1", "0")
            ElseIf Not (CheckHWS(tblSampDN.HWFONHWS) And CheckKHN(tblSampDN.HWFONKHN, 2, "TOP")) Then
                GetWFSamp = "2"
            Else
                GetWFSamp = IIf(tblSampUP.HWFONKWY = tblSampDN.HWFONKWY And _
                                tblSampUP.HWFONSPH = tblSampDN.HWFONSPH And _
                                tblSampUP.HWFONSPI = tblSampDN.HWFONSPI, "3", "4")
            End If
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '2.1.2 共有チェック追加
            '3:両方検査(共通)の場合、AN温度チェックを行い、NGの時は4:両方検査(別)にする
            If GetWFSamp = "3" Then
                liRet = funCodeDBGetMatrixReturn("SB", "AO", CStr(tblSampDN.HWFANTNP), CStr(tblSampUP.HWFANTNP))
                If liRet = -1 Then
'                    GetWFSamp = "4"
                ElseIf liRet = 0 Then
                    GetWFSamp = "4"
                End If
            End If
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        Case 3      'B1
            If Not (CheckHWS(tblSampUP.HWFBM1HS) And CheckKHN(tblSampUP.HWFBM1KN, 7, "BOT")) Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFBM1HS) And CheckKHN(tblSampDN.HWFBM1KN, 7, "TOP"), "1", "0")
            ElseIf Not (CheckHWS(tblSampDN.HWFBM1HS) And CheckKHN(tblSampDN.HWFBM1KN, 7, "TOP")) Then
                GetWFSamp = "2"
            Else
                                            '↓コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                                            'AN温度はマトリックスを使用してチェックする
'                                tblSampUP.HWFANTNP = tblSampDN.HWFANTNP And _
                                            '↑コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                GetWFSamp = IIf(tblSampUP.HWFBM1SH = tblSampDN.HWFBM1SH And _
                                tblSampUP.HWFBM1ST = tblSampDN.HWFBM1ST And _
                                tblSampUP.HWFBM1SR = tblSampDN.HWFBM1SR And _
                                tblSampUP.HWFBM1NS = tblSampDN.HWFBM1NS And _
                                tblSampUP.HWFBM1SZ = tblSampDN.HWFBM1SZ And _
                                tblSampUP.HWFBM1ET = tblSampDN.HWFBM1ET And _
                                tblSampUP.HWFANTIM = tblSampDN.HWFANTIM, "3", "4")
            End If
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '2.1.2 共有チェック追加
            '3:両方検査(共通)の場合、AN温度チェックを行い、NGの時は4:両方検査(別)にする
            If GetWFSamp = "3" Then
                liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(tblSampDN.HWFANTNP), CStr(tblSampUP.HWFANTNP))
                If liRet = -1 Then
'                    GetWFSamp = "4"
                ElseIf liRet = 0 Then
                    GetWFSamp = "4"
                End If
            End If
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        Case 4      'B2
            If Not (CheckHWS(tblSampUP.HWFBM2HS) And CheckKHN(tblSampUP.HWFBM2KN, 8, "BOT")) Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFBM2HS) And CheckKHN(tblSampDN.HWFBM2KN, 8, "TOP"), "1", "0")
            ElseIf Not (CheckHWS(tblSampDN.HWFBM2HS) And CheckKHN(tblSampDN.HWFBM2KN, 8, "TOP")) Then
                GetWFSamp = "2"
            Else
                                            '↓コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                                            'AN温度はマトリックスを使用してチェックする
'                                tblSampUP.HWFANTNP = tblSampDN.HWFANTNP And _
                                            '↑コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                GetWFSamp = IIf(tblSampUP.HWFBM2SH = tblSampDN.HWFBM2SH And _
                                tblSampUP.HWFBM2ST = tblSampDN.HWFBM2ST And _
                                tblSampUP.HWFBM2SR = tblSampDN.HWFBM2SR And _
                                tblSampUP.HWFBM2NS = tblSampDN.HWFBM2NS And _
                                tblSampUP.HWFBM2SZ = tblSampDN.HWFBM2SZ And _
                                tblSampUP.HWFBM2ET = tblSampDN.HWFBM2ET And _
                                tblSampUP.HWFANTIM = tblSampDN.HWFANTIM, "3", "4")
            End If
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '2.1.2 共有チェック追加
            '3:両方検査(共通)の場合、AN温度チェックを行い、NGの時は4:両方検査(別)にする
            If GetWFSamp = "3" Then
                liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(tblSampDN.HWFANTNP), CStr(tblSampUP.HWFANTNP))
                If liRet = -1 Then
'                    GetWFSamp = "4"
                ElseIf liRet = 0 Then
                    GetWFSamp = "4"
                End If
            End If
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        Case 5      'B3
            If Not (CheckHWS(tblSampUP.HWFBM3HS) And CheckKHN(tblSampUP.HWFBM3KN, 9, "BOT")) Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFBM3HS) And CheckKHN(tblSampDN.HWFBM3KN, 9, "TOP"), "1", "0")
            ElseIf Not (CheckHWS(tblSampDN.HWFBM3HS) And CheckKHN(tblSampDN.HWFBM3KN, 9, "TOP")) Then
                GetWFSamp = "2"
            Else
                                            '↓コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                                            'AN温度はマトリックスを使用してチェックする
'                                tblSampUP.HWFANTNP = tblSampDN.HWFANTNP And _
                                            '↑コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                GetWFSamp = IIf(tblSampUP.HWFBM3SH = tblSampDN.HWFBM3SH And _
                                tblSampUP.HWFBM3ST = tblSampDN.HWFBM3ST And _
                                tblSampUP.HWFBM3SR = tblSampDN.HWFBM3SR And _
                                tblSampUP.HWFBM3NS = tblSampDN.HWFBM3NS And _
                                tblSampUP.HWFBM3SZ = tblSampDN.HWFBM3SZ And _
                                tblSampUP.HWFBM3ET = tblSampDN.HWFBM3ET And _
                                tblSampUP.HWFANTIM = tblSampDN.HWFANTIM, "3", "4")
            End If
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '2.1.2 共有チェック追加
            '3:両方検査(共通)の場合、AN温度チェックを行い、NGの時は4:両方検査(別)にする
            If GetWFSamp = "3" Then
                liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(tblSampDN.HWFANTNP), CStr(tblSampUP.HWFANTNP))
                If liRet = -1 Then
'                    GetWFSamp = "4"
                ElseIf liRet = 0 Then
                    GetWFSamp = "4"
                End If
            End If
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        Case 6      'L1
            If Not (CheckHWS(tblSampUP.HWFOF1HS) And CheckKHN(tblSampUP.HWFOF1KN, 3, "BOT")) Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOF1HS) And CheckKHN(tblSampDN.HWFOF1KN, 3, "TOP"), "1", "0")
            ElseIf Not (CheckHWS(tblSampDN.HWFOF1HS) And CheckKHN(tblSampDN.HWFOF1KN, 3, "TOP")) Then
                GetWFSamp = "2"
            Else
                                            '↓コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                                            'AN温度はマトリックスを使用してチェックする
'                                tblSampUP.HWFANTNP = tblSampDN.HWFANTNP And _
                                            '↑コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                GetWFSamp = IIf(tblSampUP.HWFOF1SH = tblSampDN.HWFOF1SH And _
                                tblSampUP.HWFOF1ST = tblSampDN.HWFOF1ST And _
                                tblSampUP.HWFOF1SR = tblSampDN.HWFOF1SR And _
                                tblSampUP.HWFOF1NS = tblSampDN.HWFOF1NS And _
                                tblSampUP.HWFOF1SZ = tblSampDN.HWFOF1SZ And _
                                tblSampUP.HWFOF1ET = tblSampDN.HWFOF1ET And _
                                tblSampUP.HWFANTIM = tblSampDN.HWFANTIM, "3", "4")
            End If
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '2.1.2 共有チェック追加
            '3:両方検査(共通)の場合、AN温度チェックを行い、NGの時は4:両方検査(別)にする
            If GetWFSamp = "3" Then
                liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(tblSampDN.HWFANTNP), CStr(tblSampUP.HWFANTNP))
                If liRet = -1 Then
'                    GetWFSamp = "4"
                ElseIf liRet = 0 Then
                    GetWFSamp = "4"
                End If
            End If
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        Case 7      'L2
            If Not (CheckHWS(tblSampUP.HWFOF2HS) And CheckKHN(tblSampUP.HWFOF2KN, 4, "BOT")) Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOF2HS) And CheckKHN(tblSampDN.HWFOF2KN, 4, "TOP"), "1", "0")
            ElseIf Not (CheckHWS(tblSampDN.HWFOF2HS) And CheckKHN(tblSampDN.HWFOF2KN, 4, "TOP")) Then
                GetWFSamp = "2"
            Else
                                            '↓コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                                            'AN温度はマトリックスを使用してチェックする
'                                tblSampUP.HWFANTNP = tblSampDN.HWFANTNP And _
                                            '↑コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                GetWFSamp = IIf(tblSampUP.HWFOF2SH = tblSampDN.HWFOF2SH And _
                                tblSampUP.HWFOF2ST = tblSampDN.HWFOF2ST And _
                                tblSampUP.HWFOF2SR = tblSampDN.HWFOF2SR And _
                                tblSampUP.HWFOF2NS = tblSampDN.HWFOF2NS And _
                                tblSampUP.HWFOF2SZ = tblSampDN.HWFOF2SZ And _
                                tblSampUP.HWFOF2ET = tblSampDN.HWFOF2ET And _
                                tblSampUP.HWFANTIM = tblSampDN.HWFANTIM, "3", "4")
            End If
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '2.1.2 共有チェック追加
            '3:両方検査(共通)の場合、AN温度チェックを行い、NGの時は4:両方検査(別)にする
            If GetWFSamp = "3" Then
                liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(tblSampDN.HWFANTNP), CStr(tblSampUP.HWFANTNP))
                If liRet = -1 Then
'                    GetWFSamp = "4"
                ElseIf liRet = 0 Then
                    GetWFSamp = "4"
                End If
            End If
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        Case 8      'L3
            If Not (CheckHWS(tblSampUP.HWFOF3HS) And CheckKHN(tblSampUP.HWFOF3KN, 5, "BOT")) Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOF3HS) And CheckKHN(tblSampDN.HWFOF3KN, 5, "TOP"), "1", "0")
            ElseIf Not (CheckHWS(tblSampDN.HWFOF3HS) And CheckKHN(tblSampDN.HWFOF3KN, 5, "TOP")) Then
                GetWFSamp = "2"
            Else
                                            '↓コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                                            'AN温度はマトリックスを使用してチェックする
'                                tblSampUP.HWFANTNP = tblSampDN.HWFANTNP And _
                                            '↑コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                GetWFSamp = IIf(tblSampUP.HWFOF3SH = tblSampDN.HWFOF3SH And _
                                tblSampUP.HWFOF3ST = tblSampDN.HWFOF3ST And _
                                tblSampUP.HWFOF3SR = tblSampDN.HWFOF3SR And _
                                tblSampUP.HWFOF3NS = tblSampDN.HWFOF3NS And _
                                tblSampUP.HWFOF3SZ = tblSampDN.HWFOF3SZ And _
                                tblSampUP.HWFOF3ET = tblSampDN.HWFOF3ET And _
                                tblSampUP.HWFANTIM = tblSampDN.HWFANTIM, "3", "4")
            End If
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '2.1.2 共有チェック追加
            '3:両方検査(共通)の場合、AN温度チェックを行い、NGの時は4:両方検査(別)にする
            If GetWFSamp = "3" Then
                liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(tblSampDN.HWFANTNP), CStr(tblSampUP.HWFANTNP))
                If liRet = -1 Then
'                    GetWFSamp = "4"
                ElseIf liRet = 0 Then
                    GetWFSamp = "4"
                End If
            End If
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''        Case 9      'L4
'''            If Not (CheckHWS(tblSampUP.HWFOF4HS) And CheckKHN(tblSampUP.HWFOF4KN, 6, "BOT")) Then
'''                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOF4HS) And CheckKHN(tblSampDN.HWFOF4KN, 6, "TOP"), "1", "0")
'''            ElseIf Not (CheckHWS(tblSampDN.HWFOF4HS) And CheckKHN(tblSampDN.HWFOF4KN, 6, "TOP")) Then
'''                GetWFSamp = "2"
'''            Else
'''                                            '↓コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'''                                            'AN温度はマトリックスを使用してチェックする
''''                                tblSampUP.HWFANTNP = tblSampDN.HWFANTNP And _
'''                                            '↑コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
'''                GetWFSamp = IIf(tblSampUP.HWFOF4SH = tblSampDN.HWFOF4SH And _
'''                                tblSampUP.HWFOF4ST = tblSampDN.HWFOF4ST And _
'''                                tblSampUP.HWFOF4SR = tblSampDN.HWFOF4SR And _
'''                                tblSampUP.HWFOF4NS = tblSampDN.HWFOF4NS And _
'''                                tblSampUP.HWFOF4SZ = tblSampDN.HWFOF4SZ And _
'''                                tblSampUP.HWFOF4ET = tblSampDN.HWFOF4ET And _
'''                                tblSampUP.HWFANTIM = tblSampDN.HWFANTIM, "3", "4")
'''            End If

        Case 9      'SIRD
            If Not (CheckHWS(tblSampUP.HWFSIRDHS) And CheckKHN(tblSampUP.HWFSIRDKN, 6, "BOT")) Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFSIRDHS) And CheckKHN(tblSampDN.HWFSIRDKN, 6, "TOP"), "1", "0")
            ElseIf Not (CheckHWS(tblSampDN.HWFSIRDHS) And CheckKHN(tblSampDN.HWFSIRDKN, 6, "TOP")) Then
                GetWFSamp = "2"
            Else
                                            '↓コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                                            'AN温度はマトリックスを使用してチェックする
'                                tblSampUP.HWFANTNP = tblSampDN.HWFANTNP And _
                                            '↑コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                GetWFSamp = IIf(tblSampUP.HWFSIRDMX = tblSampDN.HWFSIRDMX And _
                                tblSampUP.HWFSIRDSZ = tblSampDN.HWFSIRDSZ And _
                                tblSampUP.HWFSIRDHT = tblSampDN.HWFSIRDHT And _
                                tblSampUP.HWFSIRDHS = tblSampDN.HWFSIRDHS And _
                                tblSampUP.HWFSIRDKM = tblSampDN.HWFSIRDKM And _
                                tblSampUP.HWFSIRDKN = tblSampDN.HWFSIRDKN And _
                                tblSampUP.HWFSIRDKH = tblSampDN.HWFSIRDKH And _
                                tblSampUP.HWFSIRDKU = tblSampDN.HWFSIRDKU And _
                                tblSampUP.HWFANTIM = tblSampDN.HWFANTIM, "3", "4")
            End If
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '2.1.2 共有チェック追加
            '3:両方検査(共通)の場合、AN温度チェックを行い、NGの時は4:両方検査(別)にする
            If GetWFSamp = "3" Then
                liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(tblSampDN.HWFANTNP), CStr(tblSampUP.HWFANTNP))
                If liRet = -1 Then
'                    GetWFSamp = "4"
                ElseIf liRet = 0 Then
                    GetWFSamp = "4"
                End If
            End If
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        Case 10     'DS
            If CheckHWS(tblSampUP.HWFDSOHS) And CheckKHN(tblSampUP.HWFDSOKN, 13, "BOT") Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFDSOHS) And CheckKHN(tblSampDN.HWFDSOKN, 13, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFDSOHS) And CheckKHN(tblSampDN.HWFDSOKN, 13, "TOP"), "1", "0")
            End If
            'GD/DSOD熱処理条件追加　06/12/22 ooba START =====================================>
            If GetWFSamp = "3" Then
                liLoopCnt = funCodeDBGet("SB", "15", "DS", 0, " ", lsResult)
                If liLoopCnt = 0 And Mid(lsResult, 16, 1) = "2" Then
                    liRet = funCodeDBGetMatrixReturn("SB", "AD", CStr(tblSampDN.HWFANTNP), CStr(tblSampUP.HWFANTNP))
                    If liRet = -1 Then
                    ElseIf liRet = 0 Then
                        GetWFSamp = "4"
                    End If
                End If
            End If
            'GD/DSOD熱処理条件追加　06/12/22 ooba END =======================================>
            
        Case 11     'DZ
            If Not (CheckHWS(tblSampUP.HWFMKHWS) And CheckKHN(tblSampUP.HWFMKKHN, 14, "BOT")) Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFMKHWS) And CheckKHN(tblSampDN.HWFMKKHN, 14, "TOP"), "1", "0")
            ElseIf Not (CheckHWS(tblSampDN.HWFMKHWS) And CheckKHN(tblSampDN.HWFMKKHN, 14, "TOP")) Then
                GetWFSamp = "2"
            Else
                                            '↓コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                                            'AN温度はマトリックスを使用してチェックする
'                                tblSampUP.HWFANTNP = tblSampDN.HWFANTNP And _
                                            '↑コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                GetWFSamp = IIf(tblSampUP.HWFMKSPH = tblSampDN.HWFMKSPH And _
                                tblSampUP.HWFMKSPT = tblSampDN.HWFMKSPT And _
                                tblSampUP.HWFMKSPR = tblSampDN.HWFMKSPR And _
                                tblSampUP.HWFMKNSW = tblSampDN.HWFMKNSW And _
                                tblSampUP.HWFMKSZY = tblSampDN.HWFMKSZY And _
                                tblSampUP.HWFMKCET = tblSampDN.HWFMKCET And _
                                tblSampUP.HWFANTIM = tblSampDN.HWFANTIM, "3", "4")
            End If
        Case 12     'SP     'Nr濃度追加　06/06/08 ooba
            If (CheckHWS(tblSampUP.HWFSPVHS) And CheckKHN(tblSampUP.HWFSPVKN, 15, "BOT")) _
                Or (CheckHWS(tblSampUP.HWFDLHWS) And CheckKHN(tblSampUP.HWFDLKHN, 16, "BOT")) _
                Or (CheckHWS(tblSampUP.HWFNRHS) And CheckKHN(tblSampUP.HWFNRKN, 19, "BOT")) Then
                GetWFSamp = IIf((CheckHWS(tblSampDN.HWFSPVHS) And CheckKHN(tblSampDN.HWFSPVKN, 15, "TOP")) _
                                Or (CheckHWS(tblSampDN.HWFDLHWS) And CheckKHN(tblSampDN.HWFDLKHN, 16, "TOP")) _
                                Or (CheckHWS(tblSampDN.HWFNRHS) And CheckKHN(tblSampDN.HWFNRKN, 19, "TOP")), "3", "2")
            Else
                GetWFSamp = IIf((CheckHWS(tblSampDN.HWFSPVHS) And CheckKHN(tblSampDN.HWFSPVKN, 15, "TOP")) _
                                Or (CheckHWS(tblSampDN.HWFDLHWS) And CheckKHN(tblSampDN.HWFDLKHN, 16, "TOP")) _
                                Or (CheckHWS(tblSampDN.HWFNRHS) And CheckKHN(tblSampDN.HWFNRKN, 19, "TOP")), "1", "0")
            End If
            'SPV測定方法ﾁｪｯｸ追加　07/07/04 ooba START ===========================>
            If GetWFSamp = "3" Then
                'Fe濃度測定方法(測定位置_方,測定位置_点,測定位置_位)
                If tblSampUP.HWFSPVSH <> tblSampDN.HWFSPVSH Or _
                   tblSampUP.HWFSPVST <> tblSampDN.HWFSPVST Or _
                   tblSampUP.HWFSPVSI <> tblSampDN.HWFSPVSI Then
                
                    GetWFSamp = "4"
                '拡散長測定方法(測定位置_方,測定位置_点,測定位置_位)
                ElseIf tblSampUP.HWFDLSPH <> tblSampDN.HWFDLSPH Or _
                       tblSampUP.HWFDLSPT <> tblSampDN.HWFDLSPT Or _
                       tblSampUP.HWFDLSPI <> tblSampDN.HWFDLSPI Then
                
                    GetWFSamp = "4"
                'Nr濃度測定方法(測定位置_方,測定位置_点,測定位置_位)
                ElseIf tblSampUP.HWFNRSH <> tblSampDN.HWFNRSH Or _
                       tblSampUP.HWFNRST <> tblSampDN.HWFNRST Or _
                       tblSampUP.HWFNRSI <> tblSampDN.HWFNRSI Then
                   
                    GetWFSamp = "4"
                End If
            End If
            'SPV測定方法ﾁｪｯｸ追加　07/07/04 ooba END =============================>
            
        Case 13     'D1
            If Not (CheckHWS(tblSampUP.HWFOS1HS) And CheckKHN(tblSampUP.HWFOS1KN, 10, "BOT")) Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOS1HS) And CheckKHN(tblSampDN.HWFOS1KN, 10, "TOP"), "1", "0")
            ElseIf Not (CheckHWS(tblSampDN.HWFOS1HS) And CheckKHN(tblSampDN.HWFOS1KN, 10, "TOP")) Then
                GetWFSamp = "2"
            Else
                                            '↓コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                                            'AN温度はマトリックスを使用してチェックする
'                                tblSampUP.HWFANTNP = tblSampDN.HWFANTNP And _
                                            '↑コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                GetWFSamp = IIf(tblSampUP.HWFOS1NS = tblSampDN.HWFOS1NS And _
                                tblSampUP.HWFANTIM = tblSampDN.HWFANTIM, "3", "4")
            End If
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '2.1.2 共有チェック追加
            '3:両方検査(共通)の場合、AN温度チェックを行い、NGの時は4:両方検査(別)にする
            If GetWFSamp = "3" Then
                liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(tblSampDN.HWFANTNP), CStr(tblSampUP.HWFANTNP))
                If liRet = -1 Then
'                    GetWFSamp = "4"
                ElseIf liRet = 0 Then
                    GetWFSamp = "4"
                End If
            End If
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        Case 14     'D2
            If Not (CheckHWS(tblSampUP.HWFOS2HS) And CheckKHN(tblSampUP.HWFOS2KN, 11, "BOT")) Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOS2HS) And CheckKHN(tblSampDN.HWFOS2KN, 11, "TOP"), "1", "0")
            ElseIf Not (CheckHWS(tblSampDN.HWFOS2HS) And CheckKHN(tblSampDN.HWFOS2KN, 11, "TOP")) Then
                GetWFSamp = "2"
            Else
                                            '↓コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                                            'AN温度はマトリックスを使用してチェックする
'                                tblSampUP.HWFANTNP = tblSampDN.HWFANTNP And _
                                            '↑コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                GetWFSamp = IIf(tblSampUP.HWFOS2NS = tblSampDN.HWFOS2NS And _
                                tblSampUP.HWFANTIM = tblSampDN.HWFANTIM, "3", "4")
            End If
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '2.1.2 共有チェック追加
            '3:両方検査(共通)の場合、AN温度チェックを行い、NGの時は4:両方検査(別)にする
            If GetWFSamp = "3" Then
                liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(tblSampDN.HWFANTNP), CStr(tblSampUP.HWFANTNP))
                If liRet = -1 Then
'                    GetWFSamp = "4"
                ElseIf liRet = 0 Then
                    GetWFSamp = "4"
                End If
            End If
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        Case 15     'D3
            If Not (CheckHWS(tblSampUP.HWFOS3HS) And CheckKHN(tblSampUP.HWFOS3KN, 12, "BOT")) Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFOS3HS) And CheckKHN(tblSampDN.HWFOS3KN, 12, "TOP"), "1", "0")
            ElseIf Not (CheckHWS(tblSampDN.HWFOS3HS) And CheckKHN(tblSampDN.HWFOS3KN, 12, "TOP")) Then
                GetWFSamp = "2"
            Else
                                            '↓コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                                            'AN温度はマトリックスを使用してチェックする
'                                tblSampUP.HWFANTNP = tblSampDN.HWFANTNP And _
                                            '↑コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                GetWFSamp = IIf(tblSampUP.HWFOS3NS = tblSampDN.HWFOS3NS And _
                                tblSampUP.HWFANTIM = tblSampDN.HWFANTIM, "3", "4")
            End If
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '2.1.2 共有チェック追加
            '3:両方検査(共通)の場合、AN温度チェックを行い、NGの時は4:両方検査(別)にする
            If GetWFSamp = "3" Then
                liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(tblSampDN.HWFANTNP), CStr(tblSampUP.HWFANTNP))
                If liRet = -1 Then
'                    GetWFSamp = "4"
                ElseIf liRet = 0 Then
                    GetWFSamp = "4"
                End If
            End If
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        Case 16
            If Not CheckOT(tblSampUP.HWOTHER1) Then '03/05/22
                GetWFSamp = IIf(CheckOT(tblSampDN.HWOTHER1), "1", "0") '03/05/21
            ElseIf Not CheckOT(tblSampDN.HWOTHER1) Then
                GetWFSamp = "2"
            Else
'                GetWFSamp = "3"
                GetWFSamp = "4" '必ず実測を立てる　04/05/24 ooba
''                GetWFSamp = IIf(tblSampUP.HWFOS3NS = tblSampDN.HWFOS3NS And _
''                                tblSampUP.HWFANTNP = tblSampDN.HWFANTNP And _
''                                tblSampUP.HWFANTIM = tblSampDN.HWFANTIM, "3", "4")
            End If
        Case 17
            If Not CheckOT(tblSampUP.HWOTHER2) Then    '03/05/22
                GetWFSamp = IIf(CheckOT(tblSampDN.HWOTHER2), "1", "0") '03/05/21
            ElseIf Not CheckOT(tblSampDN.HWOTHER2) Then
                GetWFSamp = "2"
            Else
'                GetWFSamp = "3"
                GetWFSamp = "4" '必ず実測を立てる　04/05/24 ooba
''                GetWFSamp = IIf(tblSampUP.HWFOS3NS = tblSampDN.HWFOS3NS And _
''                                tblSampUP.HWFANTNP = tblSampDN.HWFANTNP And _
''                                tblSampUP.HWFANTIM = tblSampDN.HWFANTIM, "3", "4")
            End If
        Case 18     'AO                                             ''残存酸素追加　03/12/05 ooba
            If Not (CheckHWS(tblSampUP.HWFZOHWS) And CheckKHN(tblSampUP.HWFZOKHN, 17, "BOT")) Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HWFZOHWS) And CheckKHN(tblSampDN.HWFZOKHN, 17, "TOP"), "1", "0")
            ElseIf Not (CheckHWS(tblSampDN.HWFZOHWS) And CheckKHN(tblSampDN.HWFZOKHN, 17, "TOP")) Then
                GetWFSamp = "2"
            Else
                                            '↓コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                                            'AN温度はマトリックスを使用してチェックする
'                                tblSampUP.HWFANTNP = tblSampDN.HWFANTNP And _
                                            '↑コメント 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                GetWFSamp = IIf(tblSampUP.HWFZONSW = tblSampDN.HWFZONSW And _
                                tblSampUP.HWFANTIM = tblSampDN.HWFANTIM, "3", "4")
            End If
            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '2.1.2 共有チェック追加
            '3:両方検査(共通)の場合、AN温度チェックを行い、NGの時は4:両方検査(別)にする
            If GetWFSamp = "3" Then
                liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(tblSampDN.HWFANTNP), CStr(tblSampUP.HWFANTNP))
                If liRet = -1 Then
'                    GetWFSamp = "4"
                ElseIf liRet = 0 Then
                    GetWFSamp = "4"
                End If
            End If
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        Case 19     'GD　05/01/18 ooba
            If (CheckHWS(tblSampUP.HWFDENHS) Or _
                CheckHWS(tblSampUP.HWFLDLHS) Or _
                CheckHWS(tblSampUP.HWFDVDHS)) And _
               CheckKHN(tblSampUP.HWFGDKHN, 18, "BOT") Then
               
                GetWFSamp = IIf((CheckHWS(tblSampDN.HWFDENHS) Or _
                                 CheckHWS(tblSampDN.HWFLDLHS) Or _
                                 CheckHWS(tblSampDN.HWFDVDHS)) And _
                                CheckKHN(tblSampDN.HWFGDKHN, 18, "TOP"), "3", "2")
            Else
                GetWFSamp = IIf((CheckHWS(tblSampDN.HWFDENHS) Or _
                                 CheckHWS(tblSampDN.HWFLDLHS) Or _
                                 CheckHWS(tblSampDN.HWFDVDHS)) And _
                                CheckKHN(tblSampDN.HWFGDKHN, 18, "TOP"), "1", "0")
            End If
            'GD/DSOD熱処理条件追加　06/12/22 ooba START =====================================>
            If GetWFSamp = "3" Then
                'DEN-AN温度ﾁｪｯｸ
                liLoopCnt = funCodeDBGet("SB", "15", "DEN", 0, " ", lsResult)
                If liLoopCnt = 0 And Mid(lsResult, 16, 1) = "2" Then
                    'LDL-AN温度ﾁｪｯｸ
                    liLoopCnt = funCodeDBGet("SB", "15", "LDL", 0, " ", lsResult)
                    If liLoopCnt = 0 And Mid(lsResult, 16, 1) = "2" Then
                        'DVD-AN温度ﾁｪｯｸ
                        liLoopCnt = funCodeDBGet("SB", "15", "DVD", 0, " ", lsResult)
                        If liLoopCnt = 0 And Mid(lsResult, 16, 1) = "2" Then
                            liRet = funCodeDBGetMatrixReturn("SB", "AG", CStr(tblSampDN.HWFANTNP), CStr(tblSampUP.HWFANTNP))
                            If liRet = -1 Then
                            ElseIf liRet = 0 Then
                                GetWFSamp = "4"
                            End If
                        End If
                    End If
                End If
            End If
            'GD/DSOD熱処理条件追加　06/12/22 ooba END =======================================>
            
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
        Case 20     'BMD1E
            If Not (CheckHWS(tblSampUP.HEPBM1HS) And CheckKHN(tblSampUP.HEPBM1KN, 1, "BOT")) Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HEPBM1HS) And CheckKHN(tblSampDN.HEPBM1KN, 1, "TOP"), "1", "0")
            ElseIf Not (CheckHWS(tblSampDN.HEPBM1HS) And CheckKHN(tblSampDN.HEPBM1KN, 1, "TOP")) Then
                GetWFSamp = "2"
            Else
                GetWFSamp = IIf(tblSampUP.HEPBM1SH = tblSampDN.HEPBM1SH And _
                                tblSampUP.HEPBM1ST = tblSampDN.HEPBM1ST And _
                                tblSampUP.HEPBM1SR = tblSampDN.HEPBM1SR And _
                                tblSampUP.HEPBM1NS = tblSampDN.HEPBM1NS And _
                                tblSampUP.HEPBM1SZ = tblSampDN.HEPBM1SZ And _
                                tblSampUP.HEPBM1ET = tblSampDN.HEPBM1ET And _
                                tblSampUP.HEPANTIM = tblSampDN.HEPANTIM, "3", "4")
            End If
            '2.1.2 共有チェック追加
            '3:両方検査(共通)の場合、AN温度チェックを行い、NGの時は4:両方検査(別)にする
            If GetWFSamp = "3" Then
                liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(tblSampDN.HEPANTNP), CStr(tblSampUP.HEPANTNP))
                If liRet = -1 Then
                ElseIf liRet = 0 Then
                    GetWFSamp = "4"
                End If
            End If
        Case 21     'BMD2E
            If Not (CheckHWS(tblSampUP.HEPBM2HS) And CheckKHN(tblSampUP.HEPBM2KN, 2, "BOT")) Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HEPBM2HS) And CheckKHN(tblSampDN.HEPBM2KN, 2, "TOP"), "1", "0")
            ElseIf Not (CheckHWS(tblSampDN.HEPBM2HS) And CheckKHN(tblSampDN.HEPBM2KN, 2, "TOP")) Then
                GetWFSamp = "2"
            Else
                GetWFSamp = IIf(tblSampUP.HEPBM2SH = tblSampDN.HEPBM2SH And _
                                tblSampUP.HEPBM2ST = tblSampDN.HEPBM2ST And _
                                tblSampUP.HEPBM2SR = tblSampDN.HEPBM2SR And _
                                tblSampUP.HEPBM2NS = tblSampDN.HEPBM2NS And _
                                tblSampUP.HEPBM2SZ = tblSampDN.HEPBM2SZ And _
                                tblSampUP.HEPBM2ET = tblSampDN.HEPBM2ET And _
                                tblSampUP.HEPANTIM = tblSampDN.HEPANTIM, "3", "4")
            End If
            '2.1.2 共有チェック追加
            '3:両方検査(共通)の場合、AN温度チェックを行い、NGの時は4:両方検査(別)にする
            If GetWFSamp = "3" Then
                liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(tblSampDN.HEPANTNP), CStr(tblSampUP.HEPANTNP))
                If liRet = -1 Then
                ElseIf liRet = 0 Then
                    GetWFSamp = "4"
                End If
            End If
        Case 22     'BMD3E
            If Not (CheckHWS(tblSampUP.HEPBM3HS) And CheckKHN(tblSampUP.HEPBM3KN, 3, "BOT")) Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HEPBM3HS) And CheckKHN(tblSampDN.HEPBM3KN, 3, "TOP"), "1", "0")
            ElseIf Not (CheckHWS(tblSampDN.HEPBM3HS) And CheckKHN(tblSampDN.HEPBM3KN, 3, "TOP")) Then
                GetWFSamp = "2"
            Else
                GetWFSamp = IIf(tblSampUP.HEPBM3SH = tblSampDN.HEPBM3SH And _
                                tblSampUP.HEPBM3ST = tblSampDN.HEPBM3ST And _
                                tblSampUP.HEPBM3SR = tblSampDN.HEPBM3SR And _
                                tblSampUP.HEPBM3NS = tblSampDN.HEPBM3NS And _
                                tblSampUP.HEPBM3SZ = tblSampDN.HEPBM3SZ And _
                                tblSampUP.HEPBM3ET = tblSampDN.HEPBM3ET And _
                                tblSampUP.HEPANTIM = tblSampDN.HEPANTIM, "3", "4")
            End If
            '2.1.2 共有チェック追加
            '3:両方検査(共通)の場合、AN温度チェックを行い、NGの時は4:両方検査(別)にする
            If GetWFSamp = "3" Then
                liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(tblSampDN.HEPANTNP), CStr(tblSampUP.HEPANTNP))
                If liRet = -1 Then
                ElseIf liRet = 0 Then
                    GetWFSamp = "4"
                End If
            End If
        Case 23     'OSF1E
            If Not (CheckHWS(tblSampUP.HEPOF1HS) And CheckKHN(tblSampUP.HEPOF1KN, 4, "BOT")) Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HEPOF1HS) And CheckKHN(tblSampDN.HEPOF1KN, 4, "TOP"), "1", "0")
            ElseIf Not (CheckHWS(tblSampDN.HEPOF1HS) And CheckKHN(tblSampDN.HEPOF1KN, 4, "TOP")) Then
                GetWFSamp = "2"
            Else
                GetWFSamp = IIf(tblSampUP.HEPOF1SH = tblSampDN.HEPOF1SH And _
                                tblSampUP.HEPOF1ST = tblSampDN.HEPOF1ST And _
                                tblSampUP.HEPOF1SR = tblSampDN.HEPOF1SR And _
                                tblSampUP.HEPOF1NS = tblSampDN.HEPOF1NS And _
                                tblSampUP.HEPOF1SZ = tblSampDN.HEPOF1SZ And _
                                tblSampUP.HEPOF1ET = tblSampDN.HEPOF1ET And _
                                tblSampUP.HEPANTIM = tblSampDN.HEPANTIM, "3", "4")
            End If
            '2.1.2 共有チェック追加
            '3:両方検査(共通)の場合、AN温度チェックを行い、NGの時は4:両方検査(別)にする
            If GetWFSamp = "3" Then
                liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(tblSampDN.HEPANTNP), CStr(tblSampUP.HEPANTNP))
                If liRet = -1 Then
                ElseIf liRet = 0 Then
                    GetWFSamp = "4"
                End If
            End If
        Case 24     'OSF2E
            If Not (CheckHWS(tblSampUP.HEPOF2HS) And CheckKHN(tblSampUP.HEPOF2KN, 5, "BOT")) Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HEPOF2HS) And CheckKHN(tblSampDN.HEPOF2KN, 5, "TOP"), "1", "0")
            ElseIf Not (CheckHWS(tblSampDN.HEPOF2HS) And CheckKHN(tblSampDN.HEPOF2KN, 5, "TOP")) Then
                GetWFSamp = "2"
            Else
                GetWFSamp = IIf(tblSampUP.HEPOF2SH = tblSampDN.HEPOF2SH And _
                                tblSampUP.HEPOF2ST = tblSampDN.HEPOF2ST And _
                                tblSampUP.HEPOF2SR = tblSampDN.HEPOF2SR And _
                                tblSampUP.HEPOF2NS = tblSampDN.HEPOF2NS And _
                                tblSampUP.HEPOF2SZ = tblSampDN.HEPOF2SZ And _
                                tblSampUP.HEPOF2ET = tblSampDN.HEPOF2ET And _
                                tblSampUP.HEPANTIM = tblSampDN.HEPANTIM, "3", "4")
            End If
            '2.1.2 共有チェック追加
            '3:両方検査(共通)の場合、AN温度チェックを行い、NGの時は4:両方検査(別)にする
            If GetWFSamp = "3" Then
                liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(tblSampDN.HEPANTNP), CStr(tblSampUP.HEPANTNP))
                If liRet = -1 Then
                ElseIf liRet = 0 Then
                    GetWFSamp = "4"
                End If
            End If
        Case 25     'OSF3E
            If Not (CheckHWS(tblSampUP.HEPOF3HS) And CheckKHN(tblSampUP.HEPOF3KN, 6, "BOT")) Then
                GetWFSamp = IIf(CheckHWS(tblSampDN.HEPOF3HS) And CheckKHN(tblSampDN.HEPOF3KN, 6, "TOP"), "1", "0")
            ElseIf Not (CheckHWS(tblSampDN.HEPOF3HS) And CheckKHN(tblSampDN.HEPOF3KN, 6, "TOP")) Then
                GetWFSamp = "2"
            Else
                GetWFSamp = IIf(tblSampUP.HEPOF3SH = tblSampDN.HEPOF3SH And _
                                tblSampUP.HEPOF3ST = tblSampDN.HEPOF3ST And _
                                tblSampUP.HEPOF3SR = tblSampDN.HEPOF3SR And _
                                tblSampUP.HEPOF3NS = tblSampDN.HEPOF3NS And _
                                tblSampUP.HEPOF3SZ = tblSampDN.HEPOF3SZ And _
                                tblSampUP.HEPOF3ET = tblSampDN.HEPOF3ET And _
                                tblSampUP.HEPANTIM = tblSampDN.HEPANTIM, "3", "4")
            End If
            '2.1.2 共有チェック追加
            '3:両方検査(共通)の場合、AN温度チェックを行い、NGの時は4:両方検査(別)にする
            If GetWFSamp = "3" Then
                liRet = funCodeDBGetMatrixReturn("SB", "AE", CStr(tblSampDN.HEPANTNP), CStr(tblSampUP.HEPANTNP))
                If liRet = -1 Then
                ElseIf liRet = 0 Then
                    GetWFSamp = "4"
                End If
            End If
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
        Case 26
            If Trim(tblSampUP.HWFGDSZY) = "G" And Trim(tblSampDN.HWFGDSZY) <> "G" Then
                GetWFSamp = "2"
            ElseIf Trim(tblSampDN.HWFGDSZY) = "G" And Trim(tblSampUP.HWFGDSZY) <> "G" Then
                GetWFSamp = "1"
            Else
                GetWFSamp = ""
            End If
'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga END     ---
    
        End Select
    End Select

End Function

'概要      :ＷＦサンプルの取得（Heavy Version）
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型           ,説明
'　　      :pHinUp　　　,I  ,tFullHinban　,上品番テーブル
'　　      :pHinDn　　　,I  ,tFullHinban　,下品番テーブル
'　　      :戻り値      ,O  ,String     　,検査指示サンプル
'説明      :RsとOiの共通検査指示について上下品番のいずれがヘヴィであるかを返す
'履歴      :2001/07/03　大塚 作成
'          :2004/04/08 ooba　WF抜試指示変更(保証方法ﾁｪｯｸの追加)
Public Function GetWFSampHeavy(pHinUp As tFullHinban, pHinDn As tFullHinban) As String

    Dim HINBANUP As String
    Dim HINBANDN As String
    Dim a As Boolean
    Dim b As Boolean

    '' 上品番／下品番が共に空でなければ
    HINBANUP = Trim(pHinUp.hinban)
    HINBANDN = Trim(pHinDn.hinban)
    If (HINBANUP = "" Or HINBANUP = "G" Or HINBANUP = "Z") And _
       (HINBANDN = "" Or HINBANDN = "G" Or HINBANDN = "Z") Then
        GetWFSampHeavy = "T"
        Exit Function
    End If

    '' 共通サンプル以外は除外する
    If HINBANUP = "" Or HINBANUP = "G" Or HINBANUP = "Z" Then
        GetWFSampHeavy = "T"
        Exit Function
    ElseIf HINBANDN = "" Or HINBANDN = "G" Or HINBANDN = "Z" Then
        GetWFSampHeavy = "T"
        Exit Function
    ElseIf HINBANUP <> HINBANDN Then
        GetWFSampHeavy = "T"
        Exit Function
    End If

    '' 上品番の製品仕様データを取得
'    If tblSampUP.HIN.hinban <> pHinUp.hinban Then
    '最新品番ﾃﾞｰﾀ取得対応　04/05/24 ooba
    If Not (tblSampUP.hin.hinban = pHinUp.hinban _
            And tblSampUP.hin.mnorevno = pHinUp.mnorevno _
            And tblSampUP.hin.factory = pHinUp.factory _
            And tblSampUP.hin.opecond = pHinUp.opecond) Then
        tblSampUP.hin.hinban = pHinUp.hinban
        tblSampUP.hin.mnorevno = pHinUp.mnorevno
        tblSampUP.hin.factory = pHinUp.factory
        tblSampUP.hin.opecond = pHinUp.opecond
        If scmzc_getWF(tblSampUP) = FUNCTION_RETURN_FAILURE Then
            GetWFSampHeavy = "T"
            Exit Function
        End If
    End If

    '' 下品番の製品仕様データを取得
'    If tblSampDN.HIN.hinban <> pHinDn.hinban Then
    '最新品番ﾃﾞｰﾀ取得対応　04/05/24 ooba
    If Not (tblSampDN.hin.hinban = pHinDn.hinban _
            And tblSampDN.hin.mnorevno = pHinDn.mnorevno _
            And tblSampDN.hin.factory = pHinDn.factory _
            And tblSampDN.hin.opecond = pHinDn.opecond) Then
        tblSampDN.hin.hinban = pHinDn.hinban
        tblSampDN.hin.mnorevno = pHinDn.mnorevno
        tblSampDN.hin.factory = pHinDn.factory
        tblSampDN.hin.opecond = pHinDn.opecond
        If scmzc_getWF(tblSampDN) = FUNCTION_RETURN_FAILURE Then
            GetWFSampHeavy = "T"
            Exit Function
        End If
    End If

    '' 共通サンプルに対して検査指示があるかチェック
    If (CheckHWS(tblSampUP.HWFRHWYS) And CheckKHN(tblSampUP.HWFRKHNN, 1, "BOT")) _
        And (CheckHWS(tblSampDN.HWFRHWYS) And CheckKHN(tblSampDN.HWFRKHNN, 1, "TOP")) Then
        a = True
    Else
        a = False
    End If
    If (CheckHWS(tblSampUP.HWFONHWS) And CheckKHN(tblSampUP.HWFONKHN, 2, "BOT")) _
        And (CheckHWS(tblSampDN.HWFONHWS) And CheckKHN(tblSampDN.HWFONKHN, 2, "TOP")) Then
        b = True
    Else
        b = False
    End If

    If a = True And b = True Then
        If tblSampUP.HWFRSPOT <= tblSampDN.HWFRSPOT And _
           tblSampUP.HWFONSPT <= tblSampDN.HWFONSPT Then
            GetWFSampHeavy = "T"
        ElseIf tblSampUP.HWFRSPOT >= tblSampDN.HWFRSPOT And _
               tblSampUP.HWFONSPT >= tblSampDN.HWFONSPT Then
            GetWFSampHeavy = "B"
        ElseIf tblSampUP.HWFRSPOT > tblSampDN.HWFRSPOT And _
               tblSampUP.HWFONSPT < tblSampDN.HWFONSPT Then
            GetWFSampHeavy = "X"
        End If
    ElseIf a = True Then
        If tblSampUP.HWFRSPOT <= tblSampDN.HWFRSPOT Then
            GetWFSampHeavy = "T"
        Else
            GetWFSampHeavy = "B"
        End If
    ElseIf b = True Then
        If tblSampUP.HWFONSPT <= tblSampDN.HWFONSPT Then
            GetWFSampHeavy = "T"
        Else
            GetWFSampHeavy = "B"
        End If
    Else
        GetWFSampHeavy = "T"
    End If

End Function

'概要      :処理方法のチェック
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型       ,説明
'　　      :sHWS  　　　,I  ,String 　,処理方法
'      　　:戻り値      ,O  ,Boolean　,検査の有無
'説明      :処理方法をチェックして検査の有無を返す
'履歴      :2001/07/03　大塚 作成
Private Function CheckHWS(ByVal sHWS As String) As Boolean

'ＷＦサンプル処理変更 2003.05.20 yakimura
'    If sHWS = "X" Or sHWS = "S" Then
    If sHWS = "H" Or sHWS = "S" Then
'ＷＦサンプル処理変更 2003.05.20 yakimura
        CheckHWS = True
    Else
        CheckHWS = False
    End If

End Function

'概要      :処理方法のチェック
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型       ,説明
'　　      :sHWS  　　　,I  ,String 　,処理方法
'      　　:戻り値      ,O  ,Boolean　,検査の有無
'説明      :処理方法をチェックして検査の有無を返す
'履歴      :2003/05/21
Private Function CheckOT(ByVal sHWS As String) As Boolean

'ＷＦサンプル処理変更 2003.05.20 yakimura
'    If sHWS = "X" Or sHWS = "S" Then
    If sHWS = "1" Then
'ＷＦサンプル処理変更 2003.05.20 yakimura
        CheckOT = True
    Else
        CheckOT = False
    End If

End Function


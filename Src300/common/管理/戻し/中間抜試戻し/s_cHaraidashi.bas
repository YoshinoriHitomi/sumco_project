Attribute VB_Name = "s_cHaraidashi"
Option Explicit
'===========================================
' ＷＦ加工用共通テーブル
'===========================================

' 抜試指示
Public Type typ_WafInd
    BLOCKID As String * 12      ' ブロックID
    BlockPos As Integer         ' ブロックＰ
    SAMPLEID    As Variant      ' add 2003/03/28 hitec)matsumoto サンプルIDを取得
    SAMPLEID2   As Variant      ' add 2003/03/28 hitec)matsumoto サンプルID2を取得
    INGOTPOS As Integer         ' 結晶Ｐ
    BkIngotPos  As Integer      ' add 2003/03/28 hitec)matsumoto
    Length As Integer           ' 長さ
    HINUP As tFullHinban        ' 上品番
    HINDN As tFullHinban        ' 下品番
    SMP As typ_WFSample         ' 検査項目
    HINFLG As Boolean           ' 品番区切りフラグ
    SMPFLG As Boolean           ' WFサンプル区切りフラグ
    ERRDNFLG As Boolean         ' 下品番エラーフラグ
    SMPLKBN1 As String * 1      ' サンプル区分１
    SMPLKBN2 As String * 1      ' サンプル区分２
    HANEIFLG As Boolean         '反映フラグ-------2003/09/23 追加 iida
End Type

' 製品仕様
Public Type typ_HinSpec
    hin As tFullHinban          ' 品番
    INGOTPOS As Integer         ' 結晶内開始位置
    Length As Integer           ' 長さ
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
    HWFNRHS  As String * 1      ' 検査有無(SP/Nr濃度)  06/06/08 ooba
    HWFOS1HS As String * 1      ' 検査有無(D1)
    HWFOS2HS As String * 1      ' 検査有無(D2)
    HWFOS3HS As String * 1      ' 検査有無(D3)
    HWFOTHER1 As String * 1     ' 検査有無(OT1) '03/05/23
    HWFOTHER2 As String * 1     ' 検査有無(OT2) '03/05/23
    HWFZOHWS As String * 1      ' 検査有無(AO)　'追加 03/12/09 ooba
    HWFDENHS As String * 1      ' 検査有無(GD/DEN)  '追加　05/01/25 ooba START ====>
    HWFLDLHS As String * 1      ' 検査有無(GD/LDL)
    HWFDVDHS As String * 1      ' 検査有無(GD/DVD2) '追加　05/01/25 ooba END ======>
    HWFOTHER1MAI As String * 1  ' 枚数(OT1) '04/06/25
    HWFOTHER2MAI As String * 1  ' 枚数(OT2) '04/06/25
    WFCUTUNIT As String * 4     ' WFカット単位 '追加 2005/04/12 ffc)tanabe
' 06/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    HEPOF1HS As String * 1      ' 検査有無(OSF1E)
    HEPOF2HS As String * 1      ' 検査有無(OSF2E)
    HEPOF3HS As String * 1      ' 検査有無(OSF3E)
    HEPBM1HS As String * 1      ' 検査有無(BMD1E)
    HEPBM2HS As String * 1      ' 検査有無(BMD2E)
    HEPBM3HS As String * 1      ' 検査有無(BMD3E)
' 06/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
' 10/02/16 Add SIRD対応 Y.Hitomi
    HWFSIRDHS As String * 1     ' 検査有無(SIRD)
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

Public tblHinSpec() As typ_HinSpec      ' 製品仕様テーブル
Public tblWafInd() As typ_WafInd        ' 抜試指示テーブル
Public tblNukishi() As typ_XSDCW        '抜試データ構造体作成用　2003/10/02 iida
Public iNowBlkPos As Integer            ' 現在表示ブロック位置
Public iNowBlkCnt As Integer            ' 現在表示ブロックサンプル数
Public tblLackMap() As typ_LackMap      ' 欠落ウェハーテーブル
Public bDispLock As Boolean             ' 画面ロックフラグ

'概要      :Ｕ／Ｄを上下に分割
'説明      :Ｕ／Ｄサンプルを上下に分割する
'履歴      :2001/10/05　大塚 作成
Public Sub SeparateUD()
    'Step3.2にて、機能廃止
End Sub

''概要      :分割サンプルの取得
''ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型       ,説明
''　　      :sSamp 　　　,IO ,String 　,サンプル
''　　      :iMode 　　　,I  ,Integer　,1:上側サンプル, 2:下側サンプル
''説明      :分割サンプルを取得する
''履歴      :2001/10/05　大塚 作成
'Public Sub GetSampleUD(sSamp As String, iMode As Integer)
'
'    Select Case sSamp
'    Case "1"
'        sSamp = IIf(iMode = 1, "0", "1")
'    Case "2"
'        sSamp = IIf(iMode = 1, "2", "0")
'    Case "3", "4"
'        sSamp = IIf(iMode = 1, "2", "1")
'    End Select
'
'End Sub
'---------------------削除------------------------------------------

'概要      :フル品番の取得
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'　　      :pHIN    　　　,IO ,tFullHinban    　,品番テーブル
'　　      :戻り値        ,O  ,FUNCTION_RETURN　,読み込みの成否
'説明      :８桁品番からフル品番を取得する
'履歴      :2001/07/11　大塚 作成
Public Function GetFullHinban(pHin As tFullHinban) As FUNCTION_RETURN

    Dim sHin As String
    Dim m As Integer
    Dim i As Integer

    sHin = Trim(pHin.hinban)
    If sHin = "" Or sHin = "G" Or sHin = "Z" Then
        pHin.mnorevno = 0
        pHin.FACTORY = ""
        pHin.OPECOND = ""
        GetFullHinban = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If
    m = UBound(tblHinSpec)
    For i = 1 To m
        If tblHinSpec(i).hin.hinban = pHin.hinban Then
            pHin = tblHinSpec(i).hin
            GetFullHinban = FUNCTION_RETURN_SUCCESS
            Exit Function
        End If
    Next i
    GetFullHinban = GetLastHinban(pHin.hinban, pHin)

End Function

'概要      :サンプルのトップ側／ボトム側区分の取得
'ﾊﾟﾗﾒｰﾀ　　:変数名　　　　,IO ,型       ,説明
'　　      :sSample      ,I  ,String 　,サンプル
'　　      :bTop         ,O  ,Boolean　,トップ側区分の有無
'　　      :bBot         ,O  ,Boolean　,ボトム側区分の有無
'説明      :サンプル区分の有無を返す
'履歴      :2001/07/11　大塚 作成
Public Sub GetSampleBT(ByVal sSample As String, bTop As Boolean, bBot As Boolean)

    Select Case sSample
    Case "1"
        bTop = True
    Case "2"
        bBot = True
    Case "4"
        bTop = True
        bBot = True
    End Select

End Sub

'概要      :サンプルのスキップ
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型      ,説明
'　　      :sSample　　　,I  ,String　,サンプル
'　　      :sSkip  　　　,I  ,String　,スキップするサンプル
'　　      :戻り値       ,O  ,String　,スキップ後のサンプル
'説明      :指定されたサンプルなら０クリアする
'履歴      :2001/07/03　大塚 作成
Public Function SkipSample(ByVal sSample As String, ByVal sSkip As String) As String

    If sSample = sSkip Then
        SkipSample = "0"
    Else
        SkipSample = sSample
    End If

End Function

'概要      :SXL IDの取得
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型       ,説明
'　　      :sBlockID 　　　,I  ,String 　,ブロックID
'　　      :iIngotPos　　　,I  ,Integer　,結晶内開始位置
'　　      :戻り値         ,O  ,String 　,SXL ID
'説明      :SXL IDを返す
'履歴      :2001/07/11　大塚 作成
Public Function GetSXLID(sBlockId As String, iIngotpos As Integer) As String

    GetSXLID = left(sBlockId, 10) & GetWafPos(iIngotpos)

End Function

'概要      :抜試位置文字列の取得
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型       ,説明
'　　      :iIngotPos　　　,I  ,Integer　,結晶内開始位置
'　　      :戻り値         ,O  ,String 　,抜試位置文字列
'説明      :抜試位置文字列を返す
'履歴      :2001/07/11　大塚 作成
Public Function GetWafPos(iIngotpos As Integer) As String

    Dim i As Integer
    Dim j As Integer

    If iIngotpos >= 1000 Then
        i = Int(iIngotpos / 100)
        j = iIngotpos Mod 100
        GetWafPos = Chr$(i - 10 + Asc("A")) & Format(j, "00")
    Else
        GetWafPos = Format(iIngotpos, "000")
    End If

End Function

'概要      :測定評価方法指示テーブルの作成
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型                ,説明
'　　      :pSXLMng　　　,I  ,typ_TBCME042   　 ,SXL管理
'　　      :pWafSmp　　　,I  ,typ_XSDCW   　    ,新サンプル管理（SXL）
'　　      :pMesInd　　　,O  ,typ_TBCMY003   　 ,測定評価方法指示
'　　      :戻り値       ,O  ,FUNCTION_RETURN　 ,読み込みの成否
'説明      :測定評価方法指示テーブルを作成する
'履歴      :2001/07/23　大塚 作成
'           2006/08/15　SMP)kondoh 修正
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -s-
'Public Function MakeMesIndTbl(pSXLMng() As typ_TBCME042, pWafSmp() As typ_XSDCW, pMesInd() As typ_TBCMY003) As FUNCTION_RETURN
Public Function MakeMesIndTbl(pSXLMng() As typ_TBCME042, pWafSmp() As typ_XSDCW, _
                        pMesInd() As typ_TBCMY003, pEpMesInd() As typ_TBCMY020) As FUNCTION_RETURN
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -e-

    Dim tmpSpWFSamp() As typ_SpWFSamp
    Dim sHin As String
    Dim sDKAN As String
    Dim m As Integer
    Dim n As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim sGdSpec As String       '規格値(GD)　05/01/27 ooba
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    '' エピ先行評価項目用のDKアニール条件
    Dim sDKAN_EP        As String
    Dim l               As Integer
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

    '' 測定評価方法指示用の製品仕様を取得
    j = 0
    m = UBound(pSXLMng)
    ReDim tmpSpWFSamp(m)
    For i = 1 To m
        sHin = RTrim$(pSXLMng(i).hinban)
        If sHin <> "" And sHin <> "G" And sHin <> "Z" Then
            j = j + 1
            tmpSpWFSamp(j).hin.hinban = pSXLMng(i).hinban
            tmpSpWFSamp(j).hin.mnorevno = pSXLMng(i).REVNUM
            tmpSpWFSamp(j).hin.FACTORY = pSXLMng(i).FACTORY
            tmpSpWFSamp(j).hin.OPECOND = pSXLMng(i).OPECOND
            If scmzc_getWF(tmpSpWFSamp(j)) = FUNCTION_RETURN_FAILURE Then
                MakeMesIndTbl = FUNCTION_RETURN_FAILURE
                Exit Function
            End If
        End If
    Next i
    ReDim Preserve tmpSpWFSamp(j)

    '' 測定評価方法指示テーブルの作成
    k = 0
    m = UBound(pWafSmp)
    n = UBound(tmpSpWFSamp)
    ReDim pMesInd(m * 18)   'OTH2を削除 エピ先行評価追加対応 06/08/15 SMP)kondoh
'    ReDim pMesInd(m * 19)   'GD追加　05/01/18 ooba
'    ReDim pMesInd(m * 18)   '残存酸素追加　03/12/05 ooba
'    ReDim pMesInd(m * 17)   '03/05/24
'    ReDim pMesInd(m * 15)

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    l = 0
    ReDim pEpMesInd(m * 7)  ' OTH2、OSF1E～OSF3E、BMD1E～BMD3E
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

    For i = 1 To m
        For j = 1 To n
            If tmpSpWFSamp(j).hin.hinban = pWafSmp(i).HINBCW Then
                Exit For
            End If
        Next j
        If j <= n Then
            With tmpSpWFSamp(j)
'                sDKAN = IIf(.HWFIGKBN = "3", "R ", "V ") & Format(.HWFANTNP, "@@@@") & Format(.HWFANTIM, "@@@@")
                'DKｱﾆｰﾙ条件変更(IG区分"4"→"R",ガス種追加)　04/07/29 ooba
                sDKAN = IIf(.HWFIGKBN = "3" Or .HWFIGKBN = "4", "R ", "V ") & Format(.HWFANTNP, "@@@@") & " " & .HWFANGZY
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                ' エピ先行評価項目用のDKアニール条件
                ' (1桁目：品EPIG区分,3～6桁目：品EPAN温度,8桁目：品EP高温ANガス条件,10桁目：品E1厚中心の整数部1の位)
                sDKAN_EP = IIf(.HEPIGKBN = "3" Or .HEPIGKBN = "4", "R", "V") & " " & _
                            IIf(.HEPANTNP >= 0, Format(.HEPANTNP, "@@@@"), Space(4)) & " " & _
                            .HEPANGZY & " " & _
                            IIf(.HEPACEN >= 0, Mid(Format(.HEPACEN, "000.00"), 3, 1), Space(1))
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
                '↓　以下全ての指示の有無を <>"0" で判定していたが、実測のみ（="1"）で判定するように変更
                If pWafSmp(i).WFINDRSCW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "RES"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "RES"
                    pMesInd(k).NETSU = ""
                    pMesInd(k).ET = ""
                    pMesInd(k).MES = .HWFRSPOH & .HWFRSPOT & .HWFRSPOI
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDOICW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "OI"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "OI"
                    pMesInd(k).NETSU = ""
                    pMesInd(k).ET = ""
                    pMesInd(k).MES = .HWFONSPH & .HWFONSPT & .HWFONSPI
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDB1CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "BMD"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "BMD1"
                    pMesInd(k).NETSU = .HWFBM1NS
                    pMesInd(k).ET = .HWFBM1SZ & Format(.HWFBM1ET, "00")
                    pMesInd(k).MES = .HWFBM1SH & .HWFBM1ST & .HWFBM1SR
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDB2CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "BMD"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "BMD2"
                    pMesInd(k).NETSU = .HWFBM2NS
                    pMesInd(k).ET = .HWFBM2SZ & Format(.HWFBM2ET, "00")
                    pMesInd(k).MES = .HWFBM2SH & .HWFBM2ST & .HWFBM2SR
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDB3CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "BMD"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "BMD3"
                    pMesInd(k).NETSU = .HWFBM3NS
                    pMesInd(k).ET = .HWFBM3SZ & Format(.HWFBM3ET, "00")
                    pMesInd(k).MES = .HWFBM3SH & .HWFBM3ST & .HWFBM3SR
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDL1CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "OSF"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "OSF1"
                    pMesInd(k).NETSU = .HWFOF1NS
                    pMesInd(k).ET = .HWFOF1SZ & Format(.HWFOF1ET, "00")
                    pMesInd(k).MES = .HWFOF1SH & .HWFOF1ST & .HWFOF1SR
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDL2CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "OSF"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "OSF2"
                    pMesInd(k).NETSU = .HWFOF2NS
                    pMesInd(k).ET = .HWFOF2SZ & Format(.HWFOF2ET, "00")
                    pMesInd(k).MES = .HWFOF2SH & .HWFOF2ST & .HWFOF2SR
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDL3CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "OSF"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "OSF3"
                    pMesInd(k).NETSU = .HWFOF3NS
                    pMesInd(k).ET = .HWFOF3SZ & Format(.HWFOF3ET, "00")
                    pMesInd(k).MES = .HWFOF3SH & .HWFOF3ST & .HWFOF3SR
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''                If pWafSmp(i).WFINDL4CW = "1" Then
'''                    k = k + 1
'''                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
'''                    pMesInd(k).OSITEM = "OSF"
'''                    pMesInd(k).SAMPLEKB = "A"
'''                    pMesInd(k).Spec = "OSF4"
'''                    pMesInd(k).NETSU = .HWFOF4NS
'''                    pMesInd(k).ET = .HWFOF4SZ & Format(.HWFOF4ET, "00")
'''                    pMesInd(k).MES = .HWFOF4SH & .HWFOF4ST & .HWFOF4SR
'''                    pMesInd(k).DKAN = sDKAN
'''                    pMesInd(k).MAISU = "1"
'''                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
'''                End If

                If pWafSmp(i).WFINDL4CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
'                    pMesInd(k).OSITEM = "SIRD"
                    pMesInd(k).OSITEM = "TENI"  '2010/05/19 REP Y.HItomi
                    pMesInd(k).SAMPLEKB = "A"
'                    pMesInd(k).Spec = "SIRD"
                    pMesInd(k).Spec = "TENI"    '2010/05/19 REP Y.HItomi
                    pMesInd(k).NETSU = ""
                    pMesInd(k).ET = .HWFSIRDSZ '''& Format(.HWFOF4ET, "00")
                    pMesInd(k).MES = ""
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
'◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)
                If pWafSmp(i).WFINDDSCW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "DSOD"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "DSOD"
                    pMesInd(k).NETSU = "G0"
                    pMesInd(k).ET = ""
                    pMesInd(k).MES = ""
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDDZCW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "DZ"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "DZ"
                    pMesInd(k).NETSU = .HWFMKNSW
                    pMesInd(k).ET = .HWFMKSZY & Format(.HWFMKCET, "00")
                    pMesInd(k).MES = .HWFMKSPH & .HWFMKSPT & .HWFMKSPR
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDSPCW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "SPV"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "SPV"
                    pMesInd(k).NETSU = ""
                    pMesInd(k).ET = ""
'                    pMesInd(k).MES = .HWFSPVSH & .HWFSPVST & .HWFSPVSI
                    '05/10/13 ooba START ==============================================>
                    If .HWFSPVHS = "H" Or .HWFSPVHS = "S" Then
                        pMesInd(k).MES = .HWFSPVSH & .HWFSPVST & .HWFSPVSI
                    ElseIf .HWFDLHWS = "H" Or .HWFDLHWS = "S" Then
                        pMesInd(k).MES = .HWFDLSPH & .HWFDLSPT & .HWFDLSPI
                    Else    'Nr濃度追加　06/06/08 ooba
                        pMesInd(k).MES = .HWFNRSH & .HWFNRST & .HWFNRSI
                    End If
                    '05/10/13 ooba END ================================================>
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    '06/06/08 ooba START ==============================================>
                    pMesInd(k).FEPUA = .HWFSPVPUG           'SPV_Fe_PUA値
                    pMesInd(k).FEPUAPCT = .HWFSPVPUR        'SPV_Fe_PUA％値
                    pMesInd(k).FESTD = .HWFSPVSTD           'SPV_Fe_STD
                    pMesInd(k).DIFFPUA = .HWFDLPUG          'SPV_拡散長_PUA値
                    pMesInd(k).DIFFPUAPCT = .HWFDLPUR       'SPV_拡散長_PUA％値
                    pMesInd(k).NRPUA = .HWFNRPUG            'SPV_NR_PUA値
                    pMesInd(k).NRPUAPCT = .HWFNRPUR         'SPV_NR_PUA%値
                    pMesInd(k).NRSTD = .HWFNRSTD            'SPV_NR_STD
                    '06/06/08 ooba END ================================================>
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDDO1CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "DOI"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "DOI1"
                    pMesInd(k).NETSU = .HWFOS1NS
                    pMesInd(k).ET = ""
                    pMesInd(k).MES = .HWFOS1SH & .HWFOS1ST & .HWFOS1SI
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDDO2CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "DOI"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "DOI2"
                    pMesInd(k).NETSU = .HWFOS2NS
                    pMesInd(k).ET = ""
                    pMesInd(k).MES = .HWFOS2SH & .HWFOS2ST & .HWFOS2SI
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDDO3CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "DOI"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "DOI3"
                    pMesInd(k).NETSU = .HWFOS3NS
                    pMesInd(k).ET = ""
                    pMesInd(k).MES = .HWFOS3SH & .HWFOS3ST & .HWFOS3SI
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                '################################## Add,03/05/23 hitec)matsumoto ##########
                If pWafSmp(i).WFINDOT1CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
''''                pMesInd(k).OSITEM = "OTH"
                    pMesInd(k).OSITEM = "OTH1"  'upd 2003/06/09 hitec)matsumoto 仕様変更
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "OTHER1"
'''                 pMesInd(k).NETSU = .HWFOS3NS
'''                 pMesInd(k).ET = ""
'''                 pMesInd(k).MES = .HWFOS3SH & .HWFOS3ST & .HWFOS3SI
                    pMesInd(k).NETSU = vbNullString
                    pMesInd(k).ET = vbNullString
                    pMesInd(k).MES = vbNullString
''''                pMesInd(k).DKAN = vbNullString  '03/05/22
                    pMesInd(k).DKAN = sDKAN 'upd 2003/06/10 hitec)matsumoto
                    pMesInd(k).MAISU = .HWOTHER1MAI
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDOT2CW = "1" Then
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -s-
''                    k = k + 1
''                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
''''''                pMesInd(k).OSITEM = "OTH"
''                    pMesInd(k).OSITEM = "OTH2"  'upd 2003/06/09 hitec)matsumoto 仕様変更
''                    pMesInd(k).SAMPLEKB = "A"
''                    pMesInd(k).Spec = "OTHER2"
'''''                 pMesInd(k).NETSU = .HWFOS3NS
'''''                 pMesInd(k).ET = ""
'''''                 pMesInd(k).MES = .HWFOS3SH & .HWFOS3ST & .HWFOS3SI
''                    pMesInd(k).NETSU = vbNullString
''                    pMesInd(k).ET = vbNullString
''                    pMesInd(k).MES = vbNullString
''''''                pMesInd(k).DKAN = vbNullString  '03/05/22
''                    pMesInd(k).DKAN = sDKAN 'upd 2003/06/10 hitec)matsumoto
''                    pMesInd(k).MAISU = .HWOTHER2MAI
                    l = l + 1
                    pEpMesInd(l).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pEpMesInd(l).OSITEM = "OTH2"
                    pEpMesInd(l).SAMPLEKB = "A"
                    pEpMesInd(l).Spec = "OTHER2"
                    pEpMesInd(l).NETSU = vbNullString
                    pEpMesInd(l).ET = vbNullString
                    pEpMesInd(l).MES = vbNullString
                    pEpMesInd(l).DKAN = sDKAN_EP
                    pEpMesInd(l).MAISU = .HWOTHER2MAI
                    pEpMesInd(l).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -e-
                End If
                '################################## End,03/05/23 hitec)matsumoto ##########
                
                '' 残存酸素追加　03/12/05 ooba START ===============================>
                If pWafSmp(i).WFINDAOICW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "AOI"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "AOI"
                    pMesInd(k).NETSU = .HWFZONSW
                    pMesInd(k).ET = ""
                    pMesInd(k).MES = .HWFZOSPH & .HWFZOSPT & .HWFZOSPI
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                '' 残存酸素追加　03/12/05 ooba END =================================>
                
                '' GD追加　05/01/18 ooba START =====================================>
                If pWafSmp(i).WFINDGDCW = "1" And pWafSmp(i).WFHSGDCW = "0" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    
                ''Upd Start (TCS)T.Terauchi 2005/10/05  抜試指示4.5ﾗｲﾝ対応
                ''    pMesInd(k).OSITEM = "GD"
                    
                    If Trim(.HWFGDLINE) = "3" Then
                        pMesInd(k).OSITEM = "GD"
                    ElseIf Trim(.HWFGDLINE) = "4.5" Then
                        pMesInd(k).OSITEM = "GD45"
                    ElseIf Trim(.HWFGDLINE) = "5" Then
                        pMesInd(k).OSITEM = "GD50"
                    Else
                        pMesInd(k).OSITEM = "GD"
                    End If
                ''Upd End   (TCS)T.Terauchi 2005/10/05  抜試指示4.5ﾗｲﾝ対応
                    
                    pMesInd(k).SAMPLEKB = "A"
                    
                    '規格値(SPEC) 1桁目:DVD2
                    If .HWFDVDHS = "H" Or .HWFDVDHS = "S" Then sGdSpec = "V" Else sGdSpec = Space(1)
                    sGdSpec = sGdSpec & Space(1)
                    '規格値(SPEC) 3桁目:L/DL
                    If .HWFLDLHS = "H" Or .HWFLDLHS = "S" Then sGdSpec = sGdSpec & "L" Else sGdSpec = sGdSpec & Space(1)
                    sGdSpec = sGdSpec & Space(1)
                    '規格値(SPEC) 5桁目:Den
                    If .HWFDENHS = "H" Or .HWFDENHS = "S" Then sGdSpec = sGdSpec & "D" Else sGdSpec = sGdSpec & Space(1)
                    
                    pMesInd(k).Spec = sGdSpec
                    pMesInd(k).NETSU = ""
                    pMesInd(k).ET = ""
'                    pMesInd(k).MES = ""
                    pMesInd(k).MES = .HWFGDSPH & .HWFGDSPT & .HWFGDZAR      '05/10/25 ooba
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                '' GD追加　05/01/18 ooba END =======================================>
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                '' OSF1E
                If pWafSmp(i).EPINDL1CW = "1" Then
                    l = l + 1
                    pEpMesInd(l).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pEpMesInd(l).OSITEM = "OSF"
                    pEpMesInd(l).SAMPLEKB = "A"
                    pEpMesInd(l).Spec = "OSF1"
                    pEpMesInd(l).NETSU = .HEPOF1NS
                    pEpMesInd(l).ET = .HEPOF1SZ & IIf(.HEPOF1ET >= 0, Format(.HEPOF1ET, "00"), Space(2))
                    pEpMesInd(l).MES = .HEPOF1SH & .HEPOF1ST & .HEPOF1SR
                    pEpMesInd(l).DKAN = sDKAN_EP
                    pEpMesInd(l).MAISU = "1"
                    pEpMesInd(l).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                '' OSF2E
                If pWafSmp(i).EPINDL2CW = "1" Then
                    l = l + 1
                    pEpMesInd(l).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pEpMesInd(l).OSITEM = "OSF"
                    pEpMesInd(l).SAMPLEKB = "A"
                    pEpMesInd(l).Spec = "OSF2"
                    pEpMesInd(l).NETSU = .HEPOF2NS
                    pEpMesInd(l).ET = .HEPOF2SZ & IIf(.HEPOF2ET >= 0, Format(.HEPOF2ET, "00"), Space(2))
                    pEpMesInd(l).MES = .HEPOF2SH & .HEPOF2ST & .HEPOF2SR
                    pEpMesInd(l).DKAN = sDKAN_EP
                    pEpMesInd(l).MAISU = "1"
                    pEpMesInd(l).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                '' OSF3E
                If pWafSmp(i).EPINDL3CW = "1" Then
                    l = l + 1
                    pEpMesInd(l).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pEpMesInd(l).OSITEM = "OSF"
                    pEpMesInd(l).SAMPLEKB = "A"
                    pEpMesInd(l).Spec = "OSF3"
                    pEpMesInd(l).NETSU = .HEPOF3NS
                    pEpMesInd(l).ET = .HEPOF3SZ & IIf(.HEPOF3ET >= 0, Format(.HEPOF3ET, "00"), Space(2))
                    pEpMesInd(l).MES = .HEPOF3SH & .HEPOF3ST & .HEPOF3SR
                    pEpMesInd(l).DKAN = sDKAN_EP
                    pEpMesInd(l).MAISU = "1"
                    pEpMesInd(l).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                '' BMD1E
                If pWafSmp(i).EPINDB1CW = "1" Then
                    l = l + 1
                    pEpMesInd(l).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pEpMesInd(l).OSITEM = "BMD"
                    pEpMesInd(l).SAMPLEKB = "A"
                    pEpMesInd(l).Spec = "BMD1"
                    pEpMesInd(l).NETSU = .HEPBM1NS
                    pEpMesInd(l).ET = .HEPBM1SZ & IIf(.HEPBM1ET >= 0, Format(.HEPBM1ET, "00"), Space(2))
                    pEpMesInd(l).MES = .HEPBM1SH & .HEPBM1ST & .HEPBM1SR
                    pEpMesInd(l).DKAN = sDKAN_EP
                    pEpMesInd(l).MAISU = "1"
                    pEpMesInd(l).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                '' BMD2E
                If pWafSmp(i).EPINDB2CW = "1" Then
                    l = l + 1
                    pEpMesInd(l).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pEpMesInd(l).OSITEM = "BMD"
                    pEpMesInd(l).SAMPLEKB = "A"
                    pEpMesInd(l).Spec = "BMD2"
                    pEpMesInd(l).NETSU = .HEPBM2NS
                    pEpMesInd(l).ET = .HEPBM2SZ & IIf(.HEPBM2ET >= 0, Format(.HEPBM2ET, "00"), Space(2))
                    pEpMesInd(l).MES = .HEPBM2SH & .HEPBM2ST & .HEPBM2SR
                    pEpMesInd(l).DKAN = sDKAN_EP
                    pEpMesInd(l).MAISU = "1"
                    pEpMesInd(l).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                '' BMD3E
                If pWafSmp(i).EPINDB3CW = "1" Then
                    l = l + 1
                    pEpMesInd(l).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pEpMesInd(l).OSITEM = "BMD"
                    pEpMesInd(l).SAMPLEKB = "A"
                    pEpMesInd(l).Spec = "BMD3"
                    pEpMesInd(l).NETSU = .HEPBM3NS
                    pEpMesInd(l).ET = .HEPBM3SZ & IIf(.HEPBM3ET >= 0, Format(.HEPBM3ET, "00"), Space(2))
                    pEpMesInd(l).MES = .HEPBM3SH & .HEPBM3ST & .HEPBM3SR
                    pEpMesInd(l).DKAN = sDKAN_EP
                    pEpMesInd(l).MAISU = "1"
                    pEpMesInd(l).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
            End With
        End If
    Next i
    ReDim Preserve pMesInd(k)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    ReDim Preserve pEpMesInd(l)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

    MakeMesIndTbl = FUNCTION_RETURN_SUCCESS

End Function


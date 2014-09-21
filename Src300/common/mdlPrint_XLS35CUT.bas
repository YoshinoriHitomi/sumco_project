Attribute VB_Name = "mdlPrint_XLS35CUT"
Option Explicit

'******************************************************************************
' @(S)
'           帳票(切断指示書)出力メイン(300mm)
'
' @(h) mdlPrint_XLS35CUT.bas             ver 1.0 ( 2008.09.22 SETsw kubota )
'
'CMBC016加工払出、CMBC030総合判定、CMDC101帳票再発行で同じファイルを使用する
'(片方で変更したらファイルをそのままコピーする)
'******************************************************************************

Private Enum enmSample
    CUT_RS = 0                      '抵抗
    CUT_OI                          'Oi
    CUT_GFA                         'GFA    2012/06/11追加 SETsw kubota
    CUT_O1                          'OSF1
    CUT_O2                          'OSF2
    CUT_O3                          'OSF3
    CUT_CO3                         '2段OS
    CUT_C                           'C      C,CJ,CJ2測定 2010/10/25追加 SETsw kubota
    CUT_CJ                          'CJ
    CUT_CJ2                         'CJ2
    CUT_B1                          'BMD1
    CUT_B2                          'BMD2
    CUT_B3                          'BMD3
    CUT_GD                          'GD
    CUT_LT                          'LT
    CUT_CS                          'CS
    CUT_EPD                         'EPD
    CUT_X                           'X線
    
    CUT_MAXCNT                      'サンプル種類
End Enum

'帳票明細
Private Type typPrintInfo_Meisai
    sBlockNo        As String       'ブロックID(結晶番号の10桁目～)
    sZuban          As String       '図番
    sLen            As String       'ブロック長さ
    sCutPos         As String       '切断位置
    sSmpl(CUT_MAXCNT - 1)   As String     '各測定項目のサンプル指示数(抵抗～EPD)
    sMaisu          As String       'サンプル枚数
    
    'サンプル図(3,3)  4枚(0～3),各1/4に分ける(0:左上,1:右上,2:左下,3:右下)
    sSmplPic(3, 3)  As String
'>>>>> サンプル№表示対応　2009/01/26　Marushita
    sSmpNo          As String       'サンプル№
'<<<<< サンプル№表示対応　2009/01/26　Marushita
'>>>>> トップ・ボトム区別対応　2009/11/12　Marushita
'>>>>> マルチ品番対応　2009/11/18　Marushita
    iSmpKbnT(CUT_MAXCNT - 1)  As Integer      'サンプル区分TOP
    iSmpKbnB(CUT_MAXCNT - 1)  As Integer      'サンプル区分BOT
'<<<<< マルチ品番対応　2009/11/18　Marushita
    sSmpNoT         As String       'サンプル№(TOP)
    sSmpNoB         As String       'サンプル№(BOT)
'<<<<< トップ・ボトム区別対応　2009/11/12　Marushita
'Add Start 2011/03/08 SMPK Nakamura FRSシステム化対応
    sFrsFlg         As String       'FRS状態
    sFrsResult      As String       'FRS実績
'Add End 2011/03/08 SMPK Nakamura FRSシステム化対応

    sNotchPos       As String       'Notch位置  2012/06/08追加 SETsw kubota

End Type

'帳票全体
Private Type typPrintInfo
    'ヘッダ
    sXtalNo         As String       '結晶番号
    sDate           As String       '発行日
    sZuban          As String       '図番(品番)
    sType           As String       '伝導型
    sDia            As String       '直径
    sJiku           As String       '結晶軸
    sRsKikaku       As String       'ρ規格
    sOiKikaku       As String       'Oi規格
    sNeraiRs        As String       'ねらい抵抗
    sCharge         As String       'チャージ量
    sPgid           As String       'PG-ID
    sBottom         As String       'ボトム状況
    sPulWeight      As String       '引上重量
    sTopCutWeight   As String       'トップカット重量
    sFreeLen        As String       'フリー長
    sPulLen         As String       '引上長さ
    sKataLen        As String       '肩カット長さ
    sOiDopPos       As String       '追ドープ位置
    
    '明細
    tMeisai()       As typPrintInfo_Meisai
    
    'サンプル厚みと形状
    sThick(CUT_MAXCNT - 1)  As String   '厚み
    sShape(CUT_MAXCNT - 1)  As String   '形状
    
    sSmplNm(CUT_MAXCNT - 1) As String   'ｻﾝﾌﾟﾙ名
    sPicStr(CUT_MAXCNT - 1) As String   'ｻﾝﾌﾟﾙ図表示文字
    
    sMaisu          As String   'サンプル枚数合計
    
    sBarCode        As String   '結晶番号バーコード　ADD 2009/08/05 SSS.Marushita

End Type
Private mtPrintInfo As typPrintInfo

Private Const PRINTFILENAME     As String = "切断指示書"
Private Const TEMPLATENAME      As String = "XLS35CUT3"
'>>>>>改ページ対応 2009/08/05 SSS.Marushita
Private Const CON_X             As Long = 40        '円の水平長
Private Const CON_Y             As Long = 40        '円の垂直長
'Private Const CON_X             As Long = 50        '円の水平長
'Private Const CON_Y             As Long = 50        '円の垂直長
'<<<<<改ページ対応 2009/08/05 SSS.Marushita

Private Const PIC_DUMMY         As String = "*"     '1/4,1/2で分けるマーク

'コードマスターよりページ管理データ取得用 ADD 2009/08/05 SSS.Marushita
Public lPRINTPAGEROW            As Long          '帳票1ページの行数
Public lPRINTMEISAIROW          As Long          '1ページ内の明細数
Public lSET_MEISAI_CNT          As Long          'マルチ品番調整後の明細数
'Add Start 2011/03/08 SMPK Nakamura FRSシステム化追加対応
Public Const FRSKBN_NONE    As String = "-"        ' FRS状態       -:FRSなし
Public Const FRSKBN_0       As String = "0"        '               0:対象外
Public Const FRSKBN_1       As String = "1"        '               1:評価
Public Const FRSKBN_2       As String = "2"        '               2:引継
Public Const FRSRSL_0       As String = "0"        ' FRS実績       0:未測定
Public Const FRSRSL_1       As String = "1"        '               1:判定OK
Public Const FRSRSL_2       As String = "2"        '               2:判定NG
Public Const FRSRSL_3       As String = "3"        '               3:再判定
Public Const FRSRSL_4       As String = "4"        '               4:測定済
Public Const FRSKBN_0_NAME  As String = "対象外"   ' FRS区分名称   対象外(状態FLG[FRS]：0)
Public Const FRSKBN_11_NAME As String = "評価"     '               評価(状態FLG[FRS]：>0、実績FLG[FRS]：0)
Public Const FRSKBN_12_NAME As String = "判定"     '               再判定(状態FLG[FRS]：>0、実績FLG[FRS]：3)
Public Const FRSKBN_13_NAME As String = "判定済"   '               判定済(状態FLG[FRS]：>0、実績FLG[FRS]：1 or 2)
'Add End 2011/03/08 SMPK Nakamura FRSシステム化追加対応

Private Const NOTCH_ASTER       As String = "****"     'ノッチ位置の*表示   2012/06/08追加 SETsw kubota


'///////////////////////////////////////////////////
' @(f)
' 機能    : Excel編集＆印刷
' 返り値  : なし
' 引き数  :
' 機能説明:
'///////////////////////////////////////////////////
Public Function PrtExec_XLS35CUT(ByVal sXtalNo As String _
                      , ByRef frmInet As Form _
                      ) As Boolean

    Dim lCnt        As Long
    Dim bResult     As Boolean

    'テンプレートダウンロード
    bResult = ActDownLoad(TEMPLATENAME, ".xls", frmInet, frmInet.Inet1)
    If bResult = False Then
        'ダウンロード失敗
        Call MsgOut(0, "帳票ファイルのダウンロードに失敗しました", ERR_DISP)
        Exit Function
    End If
    
    '印刷データ取得
    Call MsgOut(0, "印刷データ取得中", NORMAL_MSG)
    DoEvents
    If GetPrintInfo(sXtalNo) = False Then
        Exit Function
    End If
    Call MsgOut(0, "", NORMAL_MSG)

    'Excel編集＆印刷
    If PrtExec_CutSiji = False Then
        Exit Function
    End If
    
    PrtExec_XLS35CUT = True

End Function


'///////////////////////////////////////////////////
' @(f)
' 機能　　: 帳票出力情報取得処理
' 返り値　: True  - 正常
' 　　　    False - 異常
' 引き数　:
' 機能説明:
'///////////////////////////////////////////////////
Private Function GetPrintInfo(ByVal sXtalNo As String) As Boolean
    
    Dim sSql            As String
    Dim objDynaData     As Object
    Dim sXtalNoCnv      As String
    Dim lCnt            As Long
    Dim sWkZuban        As String
    Dim lBlkCnt         As Long     'ブロックカウント　2009/12/08追加
    Dim sBlock          As String   'プロック判断用　　2009/12/08追加
    Dim sZuban          As String   '図番編集用　　　　2009/12/08追加

    Dim sCsPos          As String   'XODCS位置
    Dim lCsCnt          As Long
    
    Dim lCnt2           As Long     'XODCZブロックチェック用
    Dim iJissoku        As Integer  '実測サンプルチェック用
    Dim iBlock          As Integer  'ブロックチェック用

'>>>>> コードマスターよりページ管理データ取得 ADD 2009/06/25 SSS.Marushita
    Dim tKoda9          As typKoda9Data
    '管理コード取得
    If GetKanriCode("K", "01", TEMPLATENAME, tKoda9) = False Then
        Exit Function
    End If    '帳票1ページの行数のセット
    lPRINTPAGEROW = val(Trim$(tKoda9.sKCODE01A9))
    '1ページ内の明細数のセット
    lPRINTMEISAIROW = val(Trim$(tKoda9.sKCODE02A9))
'<<<<< コードマスターよりページ管理データ取得 ADD 2009/06/25 SSS.Marushita
    
    '■データ取得(ヘッダ)
    sSql = "SELECT  NVL(C1.PUHINBC1 , ' ') PUHINBC1"                        '図番
    sSql = sSql & ",NVL(E018.HSXTYPE ,' ') HSXTYPE"                         '伝導型
    'sSQL = sSQL & ",NVL(TO_CHAR(U001.QCOM_SXLDIADV) , ' ') QCOM_SXLDIADV"   '直径区分？
    sSql = sSql & ",NVL(TO_CHAR(E018.HSXD1CEN) , ' ') HSXD1CEN"             '直径
    sSql = sSql & ",NVL(E018.HSXCDIR ,' ') HSXCDIR"                         '結晶軸
    sSql = sSql & ",NVL(TO_CHAR(E018.HSXRMIN) ,' ') HSXRMIN"                '抵抗規格(Min)
    sSql = sSql & ",NVL(TO_CHAR(E018.HSXRMAX) ,' ') HSXRMAX"                '抵抗規格(Max)
    sSql = sSql & ",NVL(TO_CHAR(E019.HSXONMIN) ,' ') HSXONMIN"              '統一Oi規格(Min)
    sSql = sSql & ",NVL(TO_CHAR(E019.HSXONMAX) ,' ') HSXONMAX"              '統一Oi規格(Max)
    sSql = sSql & ",NVL(TO_CHAR(H001.AMRESIST) ,' ') AMRESIST"              'ねらい抵抗
    sSql = sSql & ",NVL(TO_CHAR(C1.PUCHAGC1) ,' ') PUCHAGC1"                'チャージ量
    sSql = sSql & ",NVL(H001.PGID ,' ') PGID"                               'PG-ID
    sSql = sSql & ",NVL(H004.STATCLS ,' ') STATCLS"                         'ボトム状況
    sSql = sSql & ",NVL(TO_CHAR(C1.WGHTTAC1) ,' ') WGHTTAC1"                '引上重量
    sSql = sSql & ",NVL(TO_CHAR(C1.PUTCUTWC1) ,' ') PUTCUTWC1"              'トップカット重量
    sSql = sSql & ",NVL(TO_CHAR(C1.PUFRELC1) ,' ') PUFRELC1"                'フリー長
    sSql = sSql & ",NVL(TO_CHAR(C1.LENTKC1) ,' ') LENTKC1"                  '引上長さ
    'sSQL = sSQL & ",NVL(TO_CHAR(E8.KACUTLE8) ,' ') KACUTLE8"                '肩カット長さ？
    sSql = sSql & ",NVL(TO_CHAR(C1.ADDOPPC1) ,' ') ADDOPPC1"                '追ドープ位置
    sSql = sSql & ",NVL(C2.GNWKNTC2 ,' ') GNWKNTC2"                '追ドープ位置
    sSql = sSql & "  FROM XSDC2    C2"
    sSql = sSql & "     , XSDC1    C1"
    sSql = sSql & "     , TBCME018 E018"
    sSql = sSql & "     , TBCME019 E019"
    sSql = sSql & "     , TBCMH001 H001"
    sSql = sSql & "     , TBCMH004 H004"
    sSql = sSql & " WHERE C2.CRYNUMC2 = '" & sXtalNo & "'"
    sSql = sSql & "   AND C2.XTALC2 = C1.XTALC1"
    sSql = sSql & "   AND C1.HISIJIC1 = H001.UPINDNO(+)"
    sSql = sSql & "   AND C2.CRYNUMC2 = H004.CRYNUM(+)"
    sSql = sSql & "   AND ( C1.PUHINBC1 = E018.HINBAN(+)"
    sSql = sSql & "   AND   C1.PUREVNUMC1 = E018.MNOREVNO(+)"
    sSql = sSql & "   AND   C1.PUFACTORYC1 = E018.FACTORY(+)"
    sSql = sSql & "   AND   C1.PUOPEC1 = E018.OPECOND(+) )"
    sSql = sSql & "   AND ( C1.PUHINBC1 = E019.HINBAN(+)"
    sSql = sSql & "   AND   C1.PUREVNUMC1 = E019.MNOREVNO(+)"
    sSql = sSql & "   AND   C1.PUFACTORYC1 = E019.FACTORY(+)"
    sSql = sSql & "   AND   C1.PUOPEC1 = E019.OPECOND(+) )"

'Debug.Print sSQL
    
    'SQL実行
    'If mdlCommon.DynSet(objDynaData, sSQL) = False Then
    If mdlCommon.DynSet2(objDynaData, sSql) = False Then
        Call MsgOut(100, sSql, ERR_DISP_LOG, "XSDC2,C1,E018,E019,H001,H004")
        Exit Function
    End If
    If objDynaData.EOF = True Then
        Call MsgOut(0, "(切断指示書)該当データが存在しません", ERR_DISP)
        Exit Function
    End If

    '■ヘッダ編集
    With mtPrintInfo
        
        Call GetXtalHensyu(sXtalNo, 1, sXtalNoCnv)          '結晶番号編集("-"つける)
        .sXtalNo = sXtalNoCnv                               '結晶番号
        .sDate = Format$(Date, "yyyy.mm.dd")                '発行日
        .sZuban = objDynaData("PUHINBC1").Value             '品番
        Call GetHinbanHensyu(objDynaData("PUHINBC1").Value, 1, .sZuban)
            
        .sType = LCase$(objDynaData("HSXTYPE").Value)       '伝導型
        .sDia = objDynaData("HSXD1CEN").Value               '直径
        .sJiku = objDynaData("HSXCDIR").Value               '結晶軸
        .sRsKikaku = objDynaData("HSXRMIN").Value _
           & " - " & objDynaData("HSXRMAX").Value           'ρ規格"
        .sOiKikaku = objDynaData("HSXONMIN").Value _
           & " - " & objDynaData("HSXONMAX").Value          '統一Oi規格
        .sNeraiRs = objDynaData("AMRESIST").Value           'ねらい抵抗
        .sCharge = objDynaData("PUCHAGC1").Value            'チャージ量
        .sPgid = objDynaData("PGID").Value                  'PG-ID
        
        'ボトム状況(0,1以外もある？0から5？)
        .sBottom = objDynaData("STATCLS").Value
        'If objDynaData("STATCLS").Value = "0" Then
        '    .sBottom = "○"
        'ElseIf objDynaData("STATCLS").Value = "1" Then
        '    .sBottom = "×"
        'Else
        '    .sBottom = ""
        'End If
            
        .sPulWeight = objDynaData("WGHTTAC1").Value         '引上重量
        .sTopCutWeight = objDynaData("PUTCUTWC1").Value     'トップカット重量
        .sFreeLen = objDynaData("PUFRELC1").Value           'フリー長
        .sPulLen = objDynaData("LENTKC1").Value             '引上長さ
        '.sKataLen = objDynaData("KACUTLE8").Value           '肩カット長さ
        .sOiDopPos = objDynaData("ADDOPPC1").Value          '追ドープ位置
    
        .sBarCode = "*" & sXtalNo & "*"                     '結晶番号バーコード ADD 2009/06/30 SSS.Marushita
    End With
    objDynaData.Close
    
    '■データ取得(サンプル厚み・形状)
    sSql = "SELECT  NVL(NAMEJA9 ,' ')   NAMEJA9"            'サンプル名
    sSql = sSql & ",NVL(KCODEA9 ,' ')   KCODEA9"            '図表示
    sSql = sSql & ",NVL(KCODE01A9 ,' ') KCODE01A9"          '厚み
    sSql = sSql & ",NVL(KCODE02A9 ,' ') KCODE02A9"          '形状(200mm未満)
    sSql = sSql & ",NVL(KCODE03A9 ,' ') KCODE03A9"          '形状(200mm以上)
    sSql = sSql & "  FROM KODA9"
    sSql = sSql & " WHERE SYSCA9 = 'X'"
    sSql = sSql & "   AND SHUCA9 = 'HE'"
    sSql = sSql & " ORDER BY CTR01A9"

    'SQL実行
    'If mdlCommon.DynSet(objDynaData, sSQL) = False Then
    If mdlCommon.DynSet2(objDynaData, sSql) = False Then
        Call MsgOut(100, sSql, ERR_DISP_LOG, "XODC2,C1,E8,TBSSU001,002,103")
        Exit Function
    End If

    If objDynaData.RecordCount <> CUT_MAXCNT Then
        Call MsgOut(0, "(切断指示書)サンプル厚み・形状コード設定件数異常", ERR_DISP)
        Exit Function
    End If
    
    With mtPrintInfo
        For lCnt = 0 To objDynaData.RecordCount - 1
            '厚み（KODA9の'F','08'で取得？）
            .sThick(lCnt) = objDynaData("KCODE01A9").Value
'>>>>> 米沢厚み1.3mm対応　2008/10/28　SET.Marushita
            'If .sThick(lCnt) <> "1.1" And .sThick(lCnt) <> "1.2" Then
            If .sThick(lCnt) <> "1.1" And .sThick(lCnt) <> "1.2" And .sThick(lCnt) <> "1.3" Then
'<<<<< 米沢厚み1.3mm対応　2008/10/28　SET.Marushita
                Call MsgOut(0, "(切断指示書)サンプル厚みコード設定異常「" & .sThick(lCnt) & "」", ERR_DISP)
                Exit Function
            End If
            '形状(直径区分が無いので300mmのみセット？)
            'If val(mtPrintInfo.sDia) < 200 Then
                '.sShape(lCnt) = objDynaData("KCODE02A9").Value
            'Else
            .sShape(lCnt) = objDynaData("KCODE03A9").Value
            'End If
            If .sShape(lCnt) <> "1/4" And .sShape(lCnt) <> "1/2" And .sShape(lCnt) <> "4/4" Then
                Call MsgOut(0, "(切断指示書)サンプル形状コード設定異常「" & .sShape(lCnt) & "」", ERR_DISP)
                Exit Function
            End If
            .sSmplNm(lCnt) = objDynaData("NAMEJA9").Value       'ｻﾝﾌﾟﾙ名
            .sPicStr(lCnt) = objDynaData("KCODEA9").Value       'ｻﾝﾌﾟﾙ図表示文字
            objDynaData.MoveNext
        Next lCnt
    End With
    
    '■データ取得(明細)
    sSql = "SELECT  NVL(CZ.CRYNUMCZ , ' ') CRYNUMCZ"            '結晶番号
    sSql = sSql & ",NVL(CZ.HINBCZ , ' ') HINBCZ"                '図番
    sSql = sSql & ",NVL(TO_CHAR(CZ.INPOSCZ) , ' ') INPOSCZ"     '結晶部位(Top)
    sSql = sSql & ",NVL(TO_CHAR(CZ.GNLCZ) , ' ') GNLCZ"         '仕掛長さ
'Add Start 2011/03/08 SMPK Nakamura FRSシステム化対応
    sSql = sSql & ",NVL(CS1.CRYINDOIFRSCS1,'-') as INDOIFRS "   'FRS状態
    sSql = sSql & ",NVL(CS1.CRYRESOIFRSCS1,'0') as RESOIFRS "   'FRS実績
'Add End 2011/03/08 SMPK Nakamura FRSシステム化対応
    sSql = sSql & ",NVL(E018.HSXDPDIR ,' ') HSXDPDIR"           '品ＳＸ溝位置方位 2012/06/08追加 SETsw kubota

    sSql = sSql & "  FROM XSDCZ    CZ"
'Add Start 2011/03/08 SMPK Nakamura FRSシステム化対応
    sSql = sSql & ",XSDCS_1    CS1"
'Add End 2011/03/08 SMPK Nakamura FRSシステム化対応
    sSql = sSql & ",TBCME018 E018"                              '2012/06/08追加 SETsw kubota
    
    'sSQL = sSQL & " WHERE CZ.CRYNUMCZ = '" & sXtalNo & "'"
    sSql = sSql & " WHERE CZ.RPCRYNUMCZ = '" & sXtalNo & "'"
    'sSQL = sSQL & "   AND CZ.CUTKCZ = '1'"
'Add Start 2011/03/08 SMPK Nakamura FRSシステム化対応
    sSql = sSql & "   AND CZ.CRYNUMCZ = CS1.CRYNUMCS1(+)"
    sSql = sSql & "   AND CZ.HINBCZ = CS1.HINBCS1(+)"
    sSql = sSql & "   AND CZ.INPOSCZ = CS1.INPOSCS1(+)"
'Add End 2011/03/08 SMPK Nakamura FRSシステム化対応
    
    '2012/06/08追加 SETsw kubota
    sSql = sSql & "   AND CZ.HINBCZ    = E018.HINBAN(+)"
    sSql = sSql & "   AND CZ.REVNUMCZ  = E018.MNOREVNO(+)"
    sSql = sSql & "   AND CZ.FACTORYCZ = E018.FACTORY(+)"
    sSql = sSql & "   AND CZ.OPECZ     = E018.OPECOND(+)"
    
    sSql = sSql & " ORDER BY CZ.INPOSCZ"
    
    'SQL実行
    'If mdlCommon.DynSet(objDynaData, sSQL) = False Then
    If mdlCommon.DynSet2(objDynaData, sSql) = False Then
        Call MsgOut(100, sSql, ERR_DISP_LOG, "XSDCZ")
        Exit Function
    End If
    If objDynaData.EOF = True Then
        Call MsgOut(0, "(切断指示書)該当データが存在しません", ERR_DISP)
        Exit Function
    End If
    
'>>>>> マルチ品番対応改修　2009/12/08　SSS.Marushita
    lBlkCnt = 0
    sBlock = ""
    ReDim mtPrintInfo.tMeisai(objDynaData.RecordCount)      '一つ余分に作っておく
    For lCnt = 1 To objDynaData.RecordCount
        'ブロックが同じ場合は配列はそのまま
        If sBlock = Mid$(objDynaData("CRYNUMCZ").Value, 10, 3) Then
            'ブロックが同じ場合は品番を追加
            sZuban = ""                                           '図番
            Call GetHinbanHensyu(objDynaData("HINBCZ").Value, 1, sZuban)
            mtPrintInfo.tMeisai(lBlkCnt - 1).sZuban = mtPrintInfo.tMeisai(lBlkCnt - 1).sZuban & " " & Trim$(sZuban)
            mtPrintInfo.tMeisai(lBlkCnt - 1).sLen = CStr(CDbl(objDynaData("GNLCZ").Value) + CDbl(mtPrintInfo.tMeisai(lBlkCnt - 1).sLen))                    'ブロック長さ
            'マルチ品番恒久対応 2009/12/25 SSS.Marushita
            '最後のブロックの場合(最後の結晶部位(BOT)をセット) 2009/12/25
            If lCnt = objDynaData.RecordCount Then
                mtPrintInfo.tMeisai(lBlkCnt).sCutPos = CStr(CDbl(objDynaData("GNLCZ").Value) + CDbl(objDynaData("INPOSCZ").Value)) '結晶部位(BOT)
            End If
            
            'Notch位置対応 2012/06/08 SETsw kubota
            If mtPrintInfo.tMeisai(lBlkCnt - 1).sNotchPos <> CnvMizoNotchDisp(objDynaData("HSXDPDIR").Value) Then
                '上品番と仕様が異なる場合、*表示
                mtPrintInfo.tMeisai(lBlkCnt - 1).sNotchPos = NOTCH_ASTER
            End If
            
        Else
            lBlkCnt = lBlkCnt + 1
            With mtPrintInfo.tMeisai(lBlkCnt - 1)
                sZuban = ""                                       '図番判断用をクリア
                Call GetHinbanHensyu(objDynaData("HINBCZ").Value, 1, sZuban)
                .sZuban = Trim$(sZuban)
                .sLen = objDynaData("GNLCZ").Value                    'ブロック長さ
                .sCutPos = objDynaData("INPOSCZ").Value               '結晶部位(Top)
                .sBlockNo = Mid$(objDynaData("CRYNUMCZ").Value, 10, 3)    'ブロック番号？(結晶番号の10桁目～)
'Add Start 2011/03/08 SMPK Nakamura FRSシステム化対応
                .sFrsFlg = objDynaData("INDOIFRS").Value
                .sFrsResult = objDynaData("RESOIFRS").Value
'Add End 2011/03/08 SMPK Nakamura FRSシステム化対応

                'Notch位置対応 2012/06/08 SETsw kubota
                .sNotchPos = CnvMizoNotchDisp(objDynaData("HSXDPDIR").Value)
                If .sNotchPos = "" Then     '空白は*表示
                    .sNotchPos = NOTCH_ASTER
                End If
                
                sBlock = Mid$(objDynaData("CRYNUMCZ").Value, 10, 3)    'ブロック番号比較用
            End With
            '最後のブロックの場合
            If lCnt = objDynaData.RecordCount Then
                '結晶Bot位置(Top位置+長さ)を保存
                mtPrintInfo.tMeisai(lBlkCnt).sCutPos = CStr(val(mtPrintInfo.tMeisai(lBlkCnt - 1).sCutPos) + val(mtPrintInfo.tMeisai(lBlkCnt - 1).sLen))
            End If
        End If
        objDynaData.MoveNext
    Next lCnt
    objDynaData.Close
    
    lSET_MEISAI_CNT = lBlkCnt       'マルチ品番調整後の明細数
'<<<<< マルチ品番対応改修　2009/12/08　SSS.Marushita
    
    '■データ取得(サンプル管理)
    sSql = "SELECT  NVL(CS.CRYNUMCS , ' ') CRYNUMCS"        '結晶番号
    sSql = sSql & ",NVL(CS.TBKBNCS , ' ') TBKBNCS"          'T/B区分
    sSql = sSql & ",NVL(TO_CHAR(CS.INPOSCS) , ' ') INPOSCS" '結晶部位
    '抵抗
    sSql = sSql & ",NVL(CS.CRYINDRSCS , ' ') CRYINDRSCS"    '状態FLG(Rs)
    'OiまたはGFA(どちらを表示？両方表示？)
    sSql = sSql & ",NVL(CS.CRYINDOICS , ' ') CRYINDOICS"    '状態FLG(Oi)
    'OSF・OSF3(L1からL4：すべて使用？)
    sSql = sSql & ",NVL(CS.CRYINDL1CS , ' ') CRYINDL1CS"    '状態FLG(L1)
    sSql = sSql & ",NVL(CS.CRYINDL2CS , ' ') CRYINDL2CS"    '状態FLG(L2)
    sSql = sSql & ",NVL(CS.CRYINDL3CS , ' ') CRYINDL3CS"    '状態FLG(L3)
    sSql = sSql & ",NVL(CS.CRYINDL4CS , ' ') CRYINDL4CS"    '状態FLG(L4)
    'BMD(B1からB3：すべて使用？)
    sSql = sSql & ",NVL(CS.CRYINDB1CS , ' ') CRYINDB1CS"    '状態FLG(B1)
    sSql = sSql & ",NVL(CS.CRYINDB2CS , ' ') CRYINDB2CS"    '状態FLG(B2)
    sSql = sSql & ",NVL(CS.CRYINDB3CS , ' ') CRYINDB3CS"    '状態FLG(B3)
    'DvD2(表示はGD？)
    sSql = sSql & ",NVL(CS.CRYINDGDCS , ' ') CRYINDGDCS"    '状態FLG(GD)
    'LT(T)
    sSql = sSql & ",NVL(CS.CRYINDTCS , ' ') CRYINDTCS"      '状態FLG(T)
    'CS
    sSql = sSql & ",NVL(CS.CRYINDCSCS , ' ') CRYINDCSCS"    '状態FLG(CS)
    'EPD
    sSql = sSql & ",NVL(CS.CRYINDEPCS , ' ') CRYINDEPCS"    '状態FLG(EPD)
    'X線    2009/08/06追加 SETsw kubota
    sSql = sSql & ",NVL(CS.CRYINDXCS , ' ') CRYINDXCS"      '状態FLG(X線)

    'Add Start 2011/02/02 SMPK Miyata
    sSql = sSql & ",NVL(CS.CRYINDCCS , ' ') CRYINDCCS"      '状態FLG(C)
    sSql = sSql & ",NVL(CS.CRYINDCJCS , ' ') CRYINDCJCS"    '状態FLG(CJ)
    sSql = sSql & ",NVL(CS.CRYINDCJLTCS , ' ') CRYINDCJLTCS" '状態FLG(CJLT)
    sSql = sSql & ",NVL(CS.CRYINDCJ2CS , ' ') CRYINDCJ2CS"  '状態FLG(CJ2)
    'Add End   2011/02/02 SMPK Miyata

    '抵抗(FLG1,2のどちらを使用？)
    sSql = sSql & ",NVL(CS.CRYRESRS1CS , ' ') CRYRESRS1CS"  '実績FLG(Rs1)
    sSql = sSql & ",NVL(CS.CRYRESRS2CS , ' ') CRYRESRS2CS"  '実績FLG(Rs2)
    'OiまたはGFA(どちらを表示？両方表示？)
    sSql = sSql & ",NVL(CS.CRYRESOICS , ' ') CRYRESOICS"    '実績FLG(Oi)
    'OSF・OSF3(L1からL4：すべて使用？)
    sSql = sSql & ",NVL(CS.CRYRESL1CS , ' ') CRYRESL1CS"    '実績FLG(L1)
    sSql = sSql & ",NVL(CS.CRYRESL2CS , ' ') CRYRESL2CS"    '実績FLG(L2)
    sSql = sSql & ",NVL(CS.CRYRESL3CS , ' ') CRYRESL3CS"    '実績FLG(L3)
    sSql = sSql & ",NVL(CS.CRYRESL4CS , ' ') CRYRESL4CS"    '実績FLG(L4)
    'BMD(B1からB3：すべて使用？)
    sSql = sSql & ",NVL(CS.CRYRESB1CS , ' ') CRYRESB1CS"    '実績FLG(B1)
    sSql = sSql & ",NVL(CS.CRYRESB2CS , ' ') CRYRESB2CS"    '実績FLG(B2)
    sSql = sSql & ",NVL(CS.CRYRESB3CS , ' ') CRYRESB3CS"    '実績FLG(B3)
    'DvD2(表示はGD？)
    sSql = sSql & ",NVL(CS.CRYRESGDCS , ' ') CRYRESGDCS"    '実績FLG(DvD2)
    'LT(T)
    sSql = sSql & ",NVL(CS.CRYRESTCS , ' ') CRYRESTCS"      '実績FLG(LT)
    'CS
    sSql = sSql & ",NVL(CS.CRYRESCSCS , ' ') CRYRESCSCS"    '実績FLG(CS)
    'EPD
    sSql = sSql & ",NVL(CS.CRYRESEPCS , ' ') CRYRESEPCS"    '実績FLG(EPD)
    'X線    2009/08/06追加 SETsw kubota
    sSql = sSql & ",NVL(CS.CRYRESXCS , ' ') CRYRESXCS"      '実績FLG(X線)
    'Add Start 2011/02/02 SMPK Miyata
    sSql = sSql & ",NVL(CS.CRYRESCCS , ' ') CRYRESCCS"      '実績FLG(C)
    sSql = sSql & ",NVL(CS.CRYRESCJCS , ' ') CRYRESCJCS"    '実績FLG(CJ)
    sSql = sSql & ",NVL(CS.CRYRESCJLTCS , ' ') CRYRESCJLTCS" '実績FLG(CJLT)
    sSql = sSql & ",NVL(CS.CRYRESCJ2CS , ' ') CRYRESCJ2CS"  '実績FLG(CJ2)
    'Add End   2011/02/02 SMPK Miyata
'>>>>> 代表サンプルIDの取得対応　2009/01/26　Marushita
    sSql = sSql & ",NVL(CS.REPSMPLIDCS , 0) REPSMPLIDCS"    '代表サンプルID
'<<<<< 代表サンプルIDの取得対応　2009/01/26　Marushita
    
    'GFA対応 2012/06/11 SETsw kubota
    sSql = sSql & ",NVL(E019.HSXONKWY , ' ') HSXONKWY"      '品ＳＸ酸素濃度検査方法
    
    sSql = sSql & "  FROM XSDCS CS"
    sSql = sSql & "     , XSDC2 C2"
    sSql = sSql & "     , TBCME019 E019"    '2012/06/11追加 SETsw kubota
    sSql = sSql & " WHERE C2.CRYNUMC2 = '" & sXtalNo & "'"
    sSql = sSql & "   AND C2.XTALC2 = CS.XTALCS"
    sSql = sSql & "   AND CS.LIVKCS <> '1'"
    sSql = sSql & "   AND CS.HINBCS = E019.HINBAN(+)"
    sSql = sSql & "   AND CS.REVNUMCS = E019.MNOREVNO(+)"
    sSql = sSql & "   AND CS.FACTORYCS = E019.FACTORY(+)"
    sSql = sSql & "   AND CS.OPECS = E019.OPECOND(+)"
    sSql = sSql & " ORDER BY CS.INPOSCS,CS.TBKBNCS"
    
    'SQL実行
    'If mdlCommon.DynSet(objDynaData, sSQL) = False Then
    If mdlCommon.DynSet2(objDynaData, sSql) = False Then
        Call MsgOut(100, sSql, ERR_DISP_LOG, "XODCS,XODC2")
        Exit Function
    End If
    If objDynaData.EOF = True Then
        Call MsgOut(0, "(切断指示書)該当データが存在しません", ERR_DISP)
        Exit Function
    End If
    
    
    For lCsCnt = 1 To objDynaData.RecordCount
        sCsPos = objDynaData("INPOSCS").Value
        
        'For lCnt = 0 To UBound(mtPrintInfo.tMeisai)
        For lCnt = 0 To lSET_MEISAI_CNT
            With mtPrintInfo.tMeisai(lCnt)
                If sCsPos = .sCutPos Then   '位置が同じ場合
                    '>>>>> 対象ブロックチェック不具合対応  2009/11/18　SSS.Marushita
                    ''>>>>> 対象ブロックチェック対応  2009/11/12　SSS.Marushita
                    iBlock = 0
                    For lCnt2 = 0 To lSET_MEISAI_CNT
                    'For lCnt2 = 0 To UBound(mtPrintInfo.tMeisai)
                        'ブロックが同じ場合のみ対象とする
                        If Mid$(objDynaData("CRYNUMCS").Value, 10, 3) = mtPrintInfo.tMeisai(lCnt2).sBlockNo Then
                            iBlock = 1
                            Exit For
                        End If
                    Next lCnt2
                    'ブロックが同じものがあるときのみ処理
                    If iBlock = 1 Then
                        iJissoku = 0
                        '各測定項目について
                        '状態FLG='1'(実測)、実績FLG='0'(実績なし)の場合、カウントアップする
                        
                        '抵抗
                        If objDynaData("CRYINDRSCS").Value = "1" _
                        And objDynaData("CRYRESRS1CS").Value = "0" Then
                        'And objDynaData("CRYRESRS2CS").Value = "0" Then
                            
                            .sSmpl(CUT_RS) = CStr(val(.sSmpl(CUT_RS)) + 1)
                            '>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                            'CSの区分がトップの時、区分にセット
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_RS) = 1
                                iJissoku = 1
                            'CSの区分がボトムの時、区分にセット
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_RS) = 1
                                iJissoku = 2
                            End If
                            '<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                        End If
                        
                        'Oi
                        If objDynaData("CRYINDOICS").Value = "1" _
                        And objDynaData("CRYRESOICS").Value = "0" Then
                            '>>>>> GFA表示対応 2012/06/11 SETsw kubota ---------------------------
                            '.sSmpl(CUT_OI) = CStr(val(.sSmpl(CUT_OI)) + 1)
                            ''>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                            ''CSの区分がトップの時、区分にセット
                            'If objDynaData("TBKBNCS").Value = "T" Then
                            '    .iSmpKbnT(CUT_OI) = 1
                            '    iJissoku = 1
                            ''CSの区分がボトムの時、区分にセット
                            'ElseIf objDynaData("TBKBNCS").Value = "B" Then
                            '    .iSmpKbnB(CUT_OI) = 1
                            '    iJissoku = 2
                            ''<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                            'End If
                            If objDynaData("HSXONKWY").Value = "CD" Then
                                .sSmpl(CUT_OI) = CStr(val(.sSmpl(CUT_OI)) + 1)
                                'CSの区分がトップの時、区分にセット
                                If objDynaData("TBKBNCS").Value = "T" Then
                                    .iSmpKbnT(CUT_OI) = 1
                                    iJissoku = 1
                                'CSの区分がボトムの時、区分にセット
                                ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                    .iSmpKbnB(CUT_OI) = 1
                                    iJissoku = 2
                                End If
                            ElseIf objDynaData("HSXONKWY").Value = "CG" Then
                                .sSmpl(CUT_GFA) = CStr(val(.sSmpl(CUT_GFA)) + 1)
                                'CSの区分がトップの時、区分にセット
                                If objDynaData("TBKBNCS").Value = "T" Then
                                    .iSmpKbnT(CUT_GFA) = 1
                                    iJissoku = 1
                                'CSの区分がボトムの時、区分にセット
                                ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                    .iSmpKbnB(CUT_GFA) = 1
                                    iJissoku = 2
                                End If
                            End If
                            '<<<<< GFA表示対応 2012/06/11 SETsw kubota ---------------------------
                        End If
                        
                        'OSF(L1)
                        If objDynaData("CRYINDL1CS").Value = "1" _
                        And objDynaData("CRYRESL1CS").Value = "0" Then
                            .sSmpl(CUT_O1) = CStr(val(.sSmpl(CUT_O1)) + 1)
                            '>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                            'CSの区分がトップの時、区分にセット
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_O1) = 1
                                iJissoku = 1
                            'CSの区分がボトムの時、区分にセット
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_O1) = 1
                                iJissoku = 2
                            End If
                            '<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                        End If
                        
                        'OSF(L2)
                        If objDynaData("CRYINDL2CS").Value = "1" _
                        And objDynaData("CRYRESL2CS").Value = "0" Then
                            .sSmpl(CUT_O2) = CStr(val(.sSmpl(CUT_O2)) + 1)
                            '>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                            'CSの区分がトップの時、区分にセット
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_O2) = 1
                                iJissoku = 1
                            'CSの区分がボトムの時、区分にセット
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_O2) = 1
                                iJissoku = 2
                            End If
                            '<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                        End If
                        
                        'OSF(L3)
                        If objDynaData("CRYINDL3CS").Value = "1" _
                        And objDynaData("CRYRESL3CS").Value = "0" Then
                            .sSmpl(CUT_O3) = CStr(val(.sSmpl(CUT_O3)) + 1)
                            '>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                            'CSの区分がトップの時、区分にセット
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_O3) = 1
                                iJissoku = 1
                            'CSの区分がボトムの時、区分にセット
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_O3) = 1
                                iJissoku = 2
                            End If
                            '<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                        End If
                        
                        'BMD(B1)
                        If objDynaData("CRYINDB1CS").Value = "1" _
                        And objDynaData("CRYRESB1CS").Value = "0" Then
                            .sSmpl(CUT_B1) = CStr(val(.sSmpl(CUT_B1)) + 1)
                            '>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                            'CSの区分がトップの時、区分にセット
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_B1) = 1
                                iJissoku = 1
                            'CSの区分がボトムの時、区分にセット
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_B1) = 1
                                iJissoku = 2
                            End If
                            '<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                        End If
                        
                        'BMD(B2)
                        If objDynaData("CRYINDB2CS").Value = "1" _
                        And objDynaData("CRYRESB2CS").Value = "0" Then
                            .sSmpl(CUT_B2) = CStr(val(.sSmpl(CUT_B2)) + 1)
                            '>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                            'CSの区分がトップの時、区分にセット
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_B2) = 1
                                iJissoku = 1
                            'CSの区分がボトムの時、区分にセット
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_B2) = 1
                                iJissoku = 2
                            End If
                            '<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                        End If
                        
                        'BMD(B3)
                        If objDynaData("CRYINDB3CS").Value = "1" _
                        And objDynaData("CRYRESB3CS").Value = "0" Then
                            .sSmpl(CUT_B3) = CStr(val(.sSmpl(CUT_B3)) + 1)
                            '>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                            'CSの区分がトップの時、区分にセット
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_B3) = 1
                                iJissoku = 1
                            'CSの区分がボトムの時、区分にセット
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_B3) = 1
                                iJissoku = 2
                            End If
                            '<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                        End If
                        
                        'DvD2(GD)
                        If objDynaData("CRYINDGDCS").Value = "1" _
                        And objDynaData("CRYRESGDCS").Value = "0" Then
                            .sSmpl(CUT_GD) = CStr(val(.sSmpl(CUT_GD)) + 1)
                            '>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                            'CSの区分がトップの時、区分にセット
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_GD) = 1
                                iJissoku = 1
                            'CSの区分がボトムの時、区分にセット
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_GD) = 1
                                iJissoku = 2
                            End If
                            '<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                        End If
                        
                        'LT
                        If objDynaData("CRYINDTCS").Value = "1" _
                        And objDynaData("CRYRESTCS").Value = "0" Then
                            .sSmpl(CUT_LT) = CStr(val(.sSmpl(CUT_LT)) + 1)
                            '>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                            'CSの区分がトップの時、区分にセット
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_LT) = 1
                                iJissoku = 1
                            'CSの区分がボトムの時、区分にセット
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_LT) = 1
                                iJissoku = 2
                            End If
                            '<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                        End If
                        
                        'CS
                        If objDynaData("CRYINDCSCS").Value = "1" _
                        And objDynaData("CRYRESCSCS").Value = "0" Then
                            .sSmpl(CUT_CS) = CStr(val(.sSmpl(CUT_CS)) + 1)
                            '>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                            'CSの区分がトップの時、区分にセット
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_CS) = 1
                                iJissoku = 1
                            'CSの区分がボトムの時、区分にセット
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_CS) = 1
                                iJissoku = 2
                            End If
                            '<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                        End If
                        
                        'EPD
                        If objDynaData("CRYINDEPCS").Value = "1" _
                        And objDynaData("CRYRESEPCS").Value = "0" Then
                            .sSmpl(CUT_EPD) = CStr(val(.sSmpl(CUT_EPD)) + 1)
                            '>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                            'CSの区分がトップの時、区分にセット
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_EPD) = 1
                                iJissoku = 1
                            'CSの区分がボトムの時、区分にセット
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_EPD) = 1
                                iJissoku = 2
                            End If
                            '<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                        End If
                        
                        '2段OS(L4？)
                        If objDynaData("CRYINDL4CS").Value = "1" _
                        And objDynaData("CRYRESL4CS").Value = "0" Then
                            .sSmpl(CUT_CO3) = CStr(val(.sSmpl(CUT_CO3)) + 1)
                            '>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                            'CSの区分がトップの時、区分にセット
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_CO3) = 1
                                iJissoku = 1
                            'CSの区分がボトムの時、区分にセット
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_CO3) = 1
                                iJissoku = 2
                            End If
                            '<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                        End If
    
                        'X線    2009/08/06追加 SETsw kubota
                        If objDynaData("CRYINDXCS").Value = "1" _
                        And objDynaData("CRYRESXCS").Value = "0" Then
                            .sSmpl(CUT_X) = CStr(val(.sSmpl(CUT_X)) + 1)
                            '>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                            'CSの区分がトップの時、区分にセット
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_X) = 1
                                iJissoku = 1
                            'CSの区分がボトムの時、区分にセット
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_X) = 1
                                iJissoku = 2
                            End If
                            '<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                        End If

                        'Add Start 2011/02/04 SMPK Miyata
                        'C
                        If objDynaData("CRYINDCCS").Value = "1" _
                        And objDynaData("CRYRESCCS").Value = "0" Then
                            .sSmpl(CUT_C) = CStr(val(.sSmpl(CUT_C)) + 1)
                            'CSの区分がトップの時、区分にセット
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_C) = 1
                                iJissoku = 1
                            'CSの区分がボトムの時、区分にセット
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_C) = 1
                                iJissoku = 2
                            End If
                        End If
                        'CJ
                        If objDynaData("CRYINDCJCS").Value = "1" _
                        And objDynaData("CRYRESCJCS").Value = "0" Then
                            .sSmpl(CUT_CJ) = CStr(val(.sSmpl(CUT_CJ)) + 1)
                            'CSの区分がトップの時、区分にセット
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_CJ) = 1
                                iJissoku = 1
                            'CSの区分がボトムの時、区分にセット
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_CJ) = 1
                                iJissoku = 2
                            End If
                        End If
                        'CJ2
                        If objDynaData("CRYINDCJ2CS").Value = "1" _
                        And objDynaData("CRYRESCJ2CS").Value = "0" Then
                            .sSmpl(CUT_CJ2) = CStr(val(.sSmpl(CUT_CJ2)) + 1)
                            'CSの区分がトップの時、区分にセット
                            If objDynaData("TBKBNCS").Value = "T" Then
                                .iSmpKbnT(CUT_CJ2) = 1
                                iJissoku = 1
                            'CSの区分がボトムの時、区分にセット
                            ElseIf objDynaData("TBKBNCS").Value = "B" Then
                                .iSmpKbnB(CUT_CJ2) = 1
                                iJissoku = 2
                            End If
                        End If
                        'Add End   2011/02/04 SMPK Miyata

'>>>>> 代表サンプルIDのセット　2009/01/26　Marushita
                        If Trim(objDynaData("REPSMPLIDCS").Value) <> "0" Then
                            '>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                            'CSの区分がトップで実測の時、トップにセット
                            If objDynaData("TBKBNCS").Value = "T" And iJissoku = 1 Then
                                .sSmpNoT = Format(val(CStr(objDynaData("REPSMPLIDCS").Value)), "000000")
                            'CSの区分がボトムで実測の時、ボトムにセット
                            ElseIf objDynaData("TBKBNCS").Value = "B" And iJissoku = 2 Then
                                .sSmpNoB = Format(val(CStr(objDynaData("REPSMPLIDCS").Value)), "000000")
                            End If
                            '.sSmpNo = Format(val(CStr(objDynaData("REPSMPLIDCS").Value)), "000000")
                            '<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                        End If
'<<<<< 代表サンプルIDのセット　2009/01/26　Marushita
                    End If
                    'Next lCnt2
                    '<<<<< 対象ブロックチェック対応  2009/11/12　SSS.Marushita
                    '<<<<< 対象ブロックチェック不具合対応  2009/11/18　SSS.Marushita
                End If
                
            End With
        
        Next lCnt
        objDynaData.MoveNext
    Next lCsCnt

    '各位置のサンプル図編集、サンプル枚数を数える
    mtPrintInfo.sMaisu = "0"
    For lCnt = 0 To lSET_MEISAI_CNT
    'For lCnt = 0 To UBound(mtPrintInfo.tMeisai)
        Call GetSamplePic(lCnt)
        '枚数合計計算
        mtPrintInfo.sMaisu = CStr(val(mtPrintInfo.sMaisu) _
                                + val(mtPrintInfo.tMeisai(lCnt).sMaisu))
    Next lCnt
    
    '正常終了
    GetPrintInfo = True

End Function


'///////////////////////////////////////////////////
' @(f)
' 機能　　: Excel帳票編集＆印刷処理
' 返り値　: True  - 正常終了
' 　　　　  False - 異常終了
' 引き数  :
' 機能説明:
'///////////////////////////////////////////////////
Private Function PrtExec_CutSiji() As Boolean
    
    '定義をObjectに変更
    Dim xlApp           As Object               'EXCEL関連
    Dim xlBook          As Object               'EXCEL関連
    Dim xlSheet         As Object               'EXCEL関連
    'Dim xlApp           As Excel.Application    'EXCEL関連
    'Dim xlBook          As Excel.Workbook       'EXCEL関連
    'Dim xlSheet         As Excel.Worksheet      'EXCEL関連
    Dim objFSO          As Object               'FSO
    
    Dim szSavePath      As String               '出力ファイルパス
    Dim szTmpFileName   As String               'テンプレートファイル名(パス含)
    Dim szOutFileName   As String               '出力ファイル名(パス含)
    
    Dim szError         As String               'エラーメッセージ
    
    Dim lCnt            As Long                 'ループカウンタ
    Dim lSheetCnt       As Long                 'シート数

    Dim szSCell         As String               '選択セル(開始位置)
    Dim szECell         As String               '選択セル(終了位置)
    Dim szSCellTo       As String               'コピー先セル(開始位置)
    Dim szECellTo       As String               'コピー先セル(終了位置)
    
    Dim lOutputCnt      As String               '出力号機数カウンタ
    Dim sGroupNo        As String               'グループ№
    
    PrtExec_CutSiji = False
    
    Set xlApp = CreateObject("Excel.Application")
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    Call MsgOut(0, "印刷ファイル作成中", NORMAL_MSG)

    'ファイル名取得
    szSavePath = App.Path & "\" & "REPORT"
    szTmpFileName = App.Path & "\" & TEMPLATENAME & ".xls"
    szOutFileName = szSavePath & "\" & PRINTFILENAME & "_" & Format$(Now(), "YYYYMMDDhhmmss") & ".xls"

    'ディレクトリ存在有無チェック
    If Not objFSO.FolderExists(szSavePath) Then
        '無ければ作る
        Call objFSO.CreateFolder(szSavePath)
    End If
    
    'テンプレートをコピー
    objFSO.CopyFile szTmpFileName, szOutFileName
    
    'ファイルを開く
    Set xlBook = xlApp.Workbooks.Open(szOutFileName)
    Set xlSheet = xlBook.Worksheets(1)
    
    '警告メッセージ無し
    xlApp.DisplayAlerts = False
'    xlApp.DisplayAlerts = True

    xlSheet.Activate

    '書出処理
    xlApp.Visible = False                   'Excelを非表示
    On Error GoTo FileDeleteErrorExit
    
    Call SetPrintData(xlSheet)
    
    '左上を表示
    xlSheet.Cells(1, 1).Show
    
    '■印刷
    xlSheet.PrintOut
    
    '■ファイル保存
    xlBook.SaveAs szOutFileName

    Call MsgOut(0, "出力が完了しました", NORMAL_MSG)

    '終了
    xlBook.Close
'    xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    Set objFSO = Nothing

    '正常終了
    PrtExec_CutSiji = True
    Exit Function
    
FileDeleteErrorExit:

    '■終了
'    xlBook.Close
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    Set objFSO = Nothing

    Call MsgOut(0, "", ERR_DISP_LOG, "")
    szError = "ＥＸＣＥＬファイル出力に失敗しました。" & vbCrLf & _
                "(" & Err.Number & ")" & Err.Description
    
    Call MsgBox(szError, vbCritical + vbOKOnly, "EXCELﾌｧｲﾙ出力")

End Function


'///////////////////////////////////////////////////
' @(f)
' 機能　　: Excel帳票編集メイン処理
' 返り値　: True  - 正常終了
' 　　　　  False - 異常終了
' 引き数　: xlSheet  - excelシートオブジェクト
' 　　　　  lPageCnt - ページ数
' 機能説明:
'         : 改ページ対応　2009/06/29 SSS.Marushita
'///////////////////////////////////////////////////
Private Sub SetPrintData(ByRef xlSheet As Object)
'Private Sub SetPrintData(ByRef xlSheet As Excel.Worksheet)  定義をObjectに変更

    Dim lCnt        As Long
    Dim lPicCnt     As Long
    Dim lSmplCnt    As Long
    Dim lBaseCol    As Long
    Dim sStrCell    As String
    Dim sEndCell    As String
    
    Dim szSCell     As String   ' セル位置用　ADD 2009/06/29 SSS.Marushita
    Dim szECell     As String   ' セル位置用　ADD 2009/06/29 SSS.Marushita

    Dim lPageCnt    As Long     ' ページ数用　ADD 2009/06/25 SSS.Marushita
    Dim lBaseRow    As Long     ' Row位置用　 ADD 2009/06/25 SSS.Marushita

    Dim lEditCntMax As Long     ' ○編集数最大 ADD 2009/09/02 SETsw kubota

'>>>>> 色識別管理対応追加 2011/11/30 SET.Abe
    Dim sGetColor   As String       '色管理コード
    Dim typKODA9    As typKoda9Data 'KODA9定義体
'<<<<< 色識別管理対応追加 2011/11/30 SET.Abe

    '追加ページ数取得
    lPageCnt = Int((lSET_MEISAI_CNT - 1) / lPRINTMEISAIROW)
    'lPageCnt = Int((UBound(mtPrintInfo.tMeisai) - 1) / lPRINTMEISAIROW)
    For lCnt = 1 To lPageCnt
        szSCell = "A" & CStr(lCnt * lPRINTPAGEROW) + 1
        szECell = "BC" & CStr((lCnt + 1) * lPRINTPAGEROW)
        'ページのコピー
        Call xlSheet.Range("A1", "BC" & CStr(lPRINTPAGEROW)).Copy(xlSheet.Range(szSCell, szECell))
    Next lCnt
    
    'ヘッダ編集
    For lCnt = 0 To lPageCnt
        lBaseRow = lCnt * lPRINTPAGEROW
        With mtPrintInfo
'>>>>> 色識別管理対応追加 2011/12/01 SET.Abe
            '帳票色分け判別共通関数(300mm用)で色管理コードを取得
            sGetColor = Fnc_GetColor_300(Replace(.sXtalNo, "-", ""))
'            Call MsgOut(0, "DEBUG 取得色管理コード = " & sGetColor, ERR_DISP_LOG)
            '管理コードテーブル取得共通関数(GetKanriCode)により色番号を取得
            Call GetKanriCode("X", "CO", sGetColor, typKODA9)
'            Call MsgOut(0, "DEBUG 取得色番号１ = " & typKODA9.sKCODE01A9, ERR_DISP_LOG)
'            Call MsgOut(0, "DEBUG 取得色番号２ = " & typKODA9.sKCODE02A9, ERR_DISP_LOG)
            '色番号が0(白)でない時
            If typKODA9.sKCODE01A9 <> "0" Then
                '帳票タイトル左側のセル(B1～G2) の背景色を色番号１(KCODE01A9)に設定
                xlSheet.Range("B1:G2").Interior.ColorIndex = typKODA9.sKCODE01A9
            End If
            '色番号が0(白)でない時
            If typKODA9.sKCODE02A9 <> "0" Then
                '帳票タイトル右側のセル(H1～M2) の背景色を色番号２(KCODE02A9)に設定
                xlSheet.Range("H1:M2").Interior.ColorIndex = typKODA9.sKCODE02A9
            End If
'<<<<< 色識別管理対応追加 2011/12/01 SET.Abe
            
            xlSheet.Cells(lBaseRow + 2, 37).Value = "'" & .sDate             '発行日
            xlSheet.Cells(lBaseRow + 4, 5).Value = "'" & .sXtalNo            '結晶番号
            xlSheet.Cells(lBaseRow + 4, 22).Value = "'" & .sZuban            '品番
            xlSheet.Cells(lBaseRow + 5, 22).Value = "'" & .sType             '伝導型
            xlSheet.Cells(lBaseRow + 6, 22).Value = "'" & .sDia              '直径
            xlSheet.Cells(lBaseRow + 7, 22).Value = "'" & .sJiku             '結晶軸
            xlSheet.Cells(lBaseRow + 8, 22).Value = "'" & .sRsKikaku         'ρ規格
            xlSheet.Cells(lBaseRow + 9, 22).Value = "'" & .sOiKikaku         'Oi規格
        
            xlSheet.Cells(lBaseRow + 4, 31).Value = "'" & .sNeraiRs          'ねらい抵抗
            xlSheet.Cells(lBaseRow + 5, 31).Value = "'" & .sCharge           'チャージ量
            xlSheet.Cells(lBaseRow + 6, 31).Value = "'" & .sPgid             'PG-ID
            xlSheet.Cells(lBaseRow + 7, 31).Value = "'" & .sBottom           'ボトム状況
            xlSheet.Cells(lBaseRow + 8, 31).Value = "'" & .sPulWeight        '引上重量
            xlSheet.Cells(lBaseRow + 9, 31).Value = "'" & .sTopCutWeight     'トップカット重量
        
            xlSheet.Cells(lBaseRow + 4, 38).Value = "'" & .sFreeLen          'フリー長
            xlSheet.Cells(lBaseRow + 5, 38).Value = "'" & .sPulLen           '引上長さ
            '表示なし
            xlSheet.Cells(lBaseRow + 6, 35).Value = ""                       '肩カット長さ(タイトル)
            xlSheet.Cells(lBaseRow + 6, 38).Value = "'" & .sKataLen          '肩カット長さ
            xlSheet.Cells(lBaseRow + 7, 38).Value = "'" & .sOiDopPos         '追ドープ位置
            
            For lSmplCnt = CUT_RS To CUT_MAXCNT - 1
                xlSheet.Cells(lBaseRow + 20 + lSmplCnt, 2).Value = "'" & .sSmplNm(lSmplCnt)  'ｻﾝﾌﾟﾙ名
                xlSheet.Cells(lBaseRow + 20 + lSmplCnt, 4).Value = "'" & .sThick(lSmplCnt)   '厚み
                xlSheet.Cells(lBaseRow + 20 + lSmplCnt, 6).Value = "'" & .sShape(lSmplCnt)   '形状
            Next lSmplCnt
            
            xlSheet.Cells(lBaseRow + 40, 6).Value = "'" & .sMaisu            'サンプル枚数合計

            'ページ・バーコードの表示追加 ADD 2009/06/30 SSS.Marushita
            xlSheet.Cells(lBaseRow + 2, 53).Value = "'" & lCnt + 1 & "/" & lPageCnt + 1      'ページ表示
            xlSheet.Cells(lBaseRow + 4, 42).Value = "'" & .sBarCode          '結晶番号バーコード

'        xlSheet.Cells(2, 37).Value = "'" & .sDate               '発行日
'        xlSheet.Cells(4, 5).Value = "'" & .sXtalNo              '結晶番号
'        xlSheet.Cells(4, 22).Value = "'" & .sZuban              '品番
'        xlSheet.Cells(5, 22).Value = "'" & .sType               '伝導型
'        xlSheet.Cells(6, 22).Value = "'" & .sDia                '直径
'        xlSheet.Cells(7, 22).Value = "'" & .sJiku               '結晶軸
'        xlSheet.Cells(8, 22).Value = "'" & .sRsKikaku           'ρ規格
'        xlSheet.Cells(9, 22).Value = "'" & .sOiKikaku           'Oi規格
'
'        xlSheet.Cells(4, 31).Value = "'" & .sNeraiRs            'ねらい抵抗
'        xlSheet.Cells(5, 31).Value = "'" & .sCharge             'チャージ量
'        xlSheet.Cells(6, 31).Value = "'" & .sPgid               'PG-ID
'        xlSheet.Cells(7, 31).Value = "'" & .sBottom             'ボトム状況
'        xlSheet.Cells(8, 31).Value = "'" & .sPulWeight          '引上重量
'        xlSheet.Cells(9, 31).Value = "'" & .sTopCutWeight       'トップカット重量
'
'        xlSheet.Cells(4, 38).Value = "'" & .sFreeLen            'フリー長
'        xlSheet.Cells(5, 38).Value = "'" & .sPulLen             '引上長さ
'        '表示なし
'        xlSheet.Cells(6, 35).Value = ""                         '肩カット長さ(タイトル)
'        xlSheet.Cells(6, 38).Value = "'" & .sKataLen            '肩カット長さ
'        xlSheet.Cells(7, 38).Value = "'" & .sOiDopPos           '追ドープ位置
'
'        For lSmplCnt = CUT_RS To CUT_EPD
'            xlSheet.Cells(19 + lSmplCnt, 2).Value = "'" & .sSmplNm(lSmplCnt)    'ｻﾝﾌﾟﾙ名
'            xlSheet.Cells(19 + lSmplCnt, 4).Value = "'" & .sThick(lSmplCnt)     '厚み
'            xlSheet.Cells(19 + lSmplCnt, 6).Value = "'" & .sShape(lSmplCnt)     '形状
'        Next lSmplCnt
'
''>>>>> サンプル№表示対応　2009/01/26　Marushita
'        xlSheet.Cells(33, 6).Value = "'" & .sMaisu              'サンプル枚数合計
''        xlSheet.Cells(32, 6).Value = "'" & .sMaisu              'サンプル枚数合計
''<<<<< サンプル№表示対応　2009/01/26　Marushita
    
        End With
    Next lCnt
    
    '明細編集
    For lCnt = 0 To lSET_MEISAI_CNT
    'For lCnt = 0 To UBound(mtPrintInfo.tMeisai)
'        With mtPrintInfo.tMeisai(lCnt)     '使用されている箇所(下)に移動 2010/10/25
            '15明細(14明細+最後）
            '開始位置の調整(最初は固定)
            If lCnt = 0 Then
                lBaseRow = 0
                lBaseCol = 8
            Else
                lBaseRow = Int((lCnt - 1) / lPRINTMEISAIROW) * lPRINTPAGEROW
                lBaseCol = (lCnt - Int((lCnt - 1) / lPRINTMEISAIROW) * lPRINTMEISAIROW) * 3 + 8
            End If
            '次ページの先頭の時
            If Int(lCnt / lPRINTMEISAIROW) > 0 And lBaseCol = 11 Then
                '前頁最終位置を先頭にセット
                xlSheet.Cells(lBaseRow + 11, 10).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sBlockNo  'ブロックID
                '>>>>> マルチ品番恒久対応　2009/12/09　SSS.Marushita
                xlSheet.Cells(lBaseRow + 12, 9).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sZuban    '図番
                'xlSheet.Cells(lBaseRow + 12, 10).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sZuban    '図番
                'ブロックID、図番を入れた箇所に色をつける
                If mtPrintInfo.tMeisai(lCnt - 1).sBlockNo <> "" Then
                    sStrCell = ConvXlsNumToA(9) & CStr(lBaseRow + 11)
                    sEndCell = ConvXlsNumToA(11) & CStr(lBaseRow + 12)
                    xlSheet.Range(sStrCell, sEndCell).Interior.Color = &HC0C0C0
                End If
                
                xlSheet.Cells(lBaseRow + 15, 9).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sLen         'ブロック長さ
                xlSheet.Cells(lBaseRow + 14, 8).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sCutPos      '切断位置
                xlSheet.Cells(lBaseRow + 13, 10).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sNotchPos   'Notch位置 2012/06/08追加 SETsw kubota
                xlSheet.Cells(lBaseRow + 19, 8).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sCutPos      '切断位置
'Add Start 2011/03/08 SMPK Nakamura FRSシステム化対応
                'FRS測定
                If mtPrintInfo.tMeisai(lCnt - 1).sFrsFlg = FRSKBN_NONE Then
                    xlSheet.Cells(lBaseRow + 15, 11).Value = ""
                ElseIf mtPrintInfo.tMeisai(lCnt - 1).sFrsFlg = FRSKBN_0 Then
                    xlSheet.Cells(lBaseRow + 15, 11).Value = ""
                Else
                    If mtPrintInfo.tMeisai(lCnt - 1).sFrsResult = FRSRSL_0 Then
                        xlSheet.Cells(lBaseRow + 15, 11).Value = "●"
                    ElseIf mtPrintInfo.tMeisai(lCnt - 1).sFrsResult = FRSRSL_3 Then
                        xlSheet.Cells(lBaseRow + 15, 11).Value = FRSKBN_12_NAME
                    Else
                        xlSheet.Cells(lBaseRow + 15, 11).Value = ""
                    End If
                End If
'Add End 2011/03/08 SMPK Nakamura FRSシステム化対応
                For lSmplCnt = CUT_RS To CUT_MAXCNT - 1
                    '>>>>> サンプル位置トップボトム明示マルチ品番対応  2009/11/18　SSS.Marushita
                    xlSheet.Cells(lBaseRow + 20 + lSmplCnt, 8).Value = "'" & GetSampleStr(CStr(mtPrintInfo.tMeisai(lCnt - 1).iSmpKbnT(lSmplCnt)), _
                                                                                          CStr(mtPrintInfo.tMeisai(lCnt - 1).iSmpKbnB(lSmplCnt)))
                    '>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                    'xlSheet.Cells(lBaseRow + 19 + lSmplCnt, 8).Value = "'" & GetSampleStr(CStr(mtPrintInfo.tMeisai(lCnt - 1).iSmpKbn(lSmplCnt)))
                    'xlSheet.Cells(lBaseRow + 19 + lSmplCnt, 8).Value = "'" & GetSampleStr(mtPrintInfo.tMeisai(lCnt - 1).sSmpl(lSmplCnt))
                    '<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                    '<<<<< サンプル位置トップボトム明示マルチ品番対応  2009/11/18　SSS.Marushita
                Next lSmplCnt
                '>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                xlSheet.Cells(lBaseRow + 38, 8).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sSmpNoB       'サンプル№BOT
                xlSheet.Cells(lBaseRow + 39, 8).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sSmpNoT       'サンプル№TOP
                'xlSheet.Cells(lBaseRow + 36, 8).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sSmpNo       'サンプル№
                '<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                xlSheet.Cells(lBaseRow + 40, 8).Value = "'" & mtPrintInfo.tMeisai(lCnt - 1).sMaisu       'サンプル指示(抵抗)
                For lSmplCnt = 0 To val(mtPrintInfo.tMeisai(lCnt - 1).sMaisu) - 1
                    
'>>>>> X線測定対応 2009/09/02 SETsw kubota ------------------
'                    Call OvalWrite(xlSheet, lBaseRow + 39 + lSmplCnt * 5, 8)
                    If mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, 0) = "X" Then
                        'X線の場合、一ページ目の合計の下に○表示
                        Call OvalWrite(xlSheet, 42, 4)
                        xlSheet.Cells(43, 4).Value = mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, 0)
                        lEditCntMax = 4     '通常の位置にX線を編集しない分、一枚多くループ
                    Else
                        lEditCntMax = 3
                        Call OvalWrite(xlSheet, lBaseRow + 42 + lSmplCnt * 5, 8)
                    End If
'<<<<< X線測定対応 2009/09/02 SETsw kubota ------------------
                    
                    For lPicCnt = 0 To 3
                        If mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, lPicCnt) <> PIC_DUMMY Then
                            Select Case lPicCnt
                            Case 0  '左上
                                If mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, 0) <> "X" Then   'X線は二ページ目以降○表示なし
                                    xlSheet.Cells(lBaseRow + 43 + lSmplCnt * 5, 8).Value = mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, lPicCnt)
                                    
                                    '右上が空でなければ右と下に罫線を引く
                                    If mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, 1) <> "" Then
                                        sStrCell = ConvXlsNumToA(8) & CStr(lBaseRow + 42 + lSmplCnt * 5)
                                        sEndCell = ConvXlsNumToA(8) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                        With xlSheet.Range(sStrCell, sEndCell)
                                            .Borders(xlEdgeRight).LineStyle = xlContinuous  '通常線
                                            .Borders(xlEdgeRight).Weight = xlThin
                                        End With
                                        sStrCell = ConvXlsNumToA(8) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                        sEndCell = ConvXlsNumToA(8) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                        With xlSheet.Range(sStrCell, sEndCell)
                                            .Borders(xlEdgeBottom).LineStyle = xlContinuous  '通常線
                                            .Borders(xlEdgeBottom).Weight = xlThin
                                        End With
                                    End If
                                    '左下が空でなければ下罫線を引く
                                    If mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, 2) <> "" Then
                                        sStrCell = ConvXlsNumToA(8) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                        sEndCell = ConvXlsNumToA(9) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                        With xlSheet.Range(sStrCell, sEndCell)
                                            .Borders(xlEdgeBottom).LineStyle = xlContinuous '通常線
                                            .Borders(xlEdgeBottom).Weight = xlThin
                                        End With
                                    End If
                                End If
                            Case 1
                                xlSheet.Cells(lBaseRow + 43 + lSmplCnt * 5, 9).Value = mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, lPicCnt)
                            Case 2
                                xlSheet.Cells(lBaseRow + 44 + lSmplCnt * 5, 8).Value = mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, lPicCnt)
                                '右下が空でなければ右罫線を引く
                                If mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, 3) <> "" Then
                                    sStrCell = ConvXlsNumToA(8) & CStr(lBaseRow + 44 + lSmplCnt * 5)
                                    sEndCell = ConvXlsNumToA(8) & CStr(lBaseRow + 45 + lSmplCnt * 5)
                                    With xlSheet.Range(sStrCell, sEndCell)
                                        .Borders(xlEdgeRight).LineStyle = xlContinuous  '通常線
                                        .Borders(xlEdgeRight).Weight = xlThin
                                    End With
                                End If
                            Case 3
                                xlSheet.Cells(lBaseRow + 44 + lSmplCnt * 5, 9).Value = mtPrintInfo.tMeisai(lCnt - 1).sSmplPic(lSmplCnt, lPicCnt)
                            End Select
                        End If
                    Next lPicCnt
                    
                    '4枚編集したら抜ける
                    'If lSmplCnt = 2 Then
                    If lSmplCnt = lEditCntMax Then
                        Exit For
                    End If
                Next
                
            End If
            
        With mtPrintInfo.tMeisai(lCnt)     '移動 2010/10/25
            
            '最終位置の時、ブロックID・図番・色つけ・ブロック長さは表示しない
            If (lCnt Mod lPRINTMEISAIROW) = 0 And lCnt > 0 Then
            Else
                xlSheet.Cells(lBaseRow + 11, lBaseCol + 2).Value = "'" & .sBlockNo    'ブロックID
                '>>>>> マルチ品番恒久対応　2009/12/09　SSS.Marushita
                xlSheet.Cells(lBaseRow + 12, lBaseCol + 1).Value = "'" & .sZuban      '図番
                'xlSheet.Cells(lBaseRow + 12, lBaseCol + 2).Value = "'" & .sZuban      '図番
                'lBaseCol = lCnt * 3 + 9
                'xlSheet.Cells(11, lBaseCol + 1).Value = "'" & .sBlockNo    'ブロックID
                'xlSheet.Cells(12, lBaseCol + 1).Value = "'" & .sZuban      '図番
                xlSheet.Cells(lBaseRow + 13, lBaseCol + 2).Value = "'" & .sNotchPos 'Notch位置  2012/06/08追加 SETsw kubota
                
                'ブロックID、図番を入れた箇所に色をつける
                If .sBlockNo <> "" Then
                    sStrCell = ConvXlsNumToA(lBaseCol + 1) & CStr(lBaseRow + 11)
                    sEndCell = ConvXlsNumToA(lBaseCol + 3) & CStr(lBaseRow + 12)
                    'sStrCell = ConvXlsNumToA(lBaseCol) & "11"
                    'sEndCell = ConvXlsNumToA(lBaseCol + 2) & "12"
                    xlSheet.Range(sStrCell, sEndCell).Interior.Color = &HC0C0C0
                End If
                
                xlSheet.Cells(lBaseRow + 15, lBaseCol + 1).Value = "'" & .sLen         'ブロック長さ
                'xlSheet.Cells(15, lBaseCol).Value = "'" & .sLen             'ブロック長さ
'Add Start 2011/03/08 SMPK Nakamura FRSシステム化対応
                'FRS測定
                If .sFrsFlg = FRSKBN_NONE Then
                    xlSheet.Cells(lBaseRow + 17, lBaseCol + 1).Value = ""
                ElseIf .sFrsFlg = FRSKBN_0 Then
                    xlSheet.Cells(lBaseRow + 17, lBaseCol + 1).Value = ""
                Else
                    If .sFrsResult = FRSRSL_0 Then
                        xlSheet.Cells(lBaseRow + 17, lBaseCol + 1).Value = "●"
                    ElseIf .sFrsResult = FRSRSL_3 Then
                        xlSheet.Cells(lBaseRow + 17, lBaseCol + 1).Value = FRSKBN_12_NAME
                    Else
                        xlSheet.Cells(lBaseRow + 17, lBaseCol + 1).Value = ""
                    End If
                End If
'Add End 2011/03/08 SMPK Nakamura FRSシステム化対応
            End If
            xlSheet.Cells(lBaseRow + 14, lBaseCol).Value = "'" & .sCutPos       '切断位置
            xlSheet.Cells(lBaseRow + 19, lBaseCol).Value = "'" & .sCutPos       '切断位置
            'lBaseCol = lCnt * 3 + 8
            'xlSheet.Cells(14, lBaseCol).Value = "'" & .sCutPos          '切断位置
            'xlSheet.Cells(18, lBaseCol).Value = "'" & .sCutPos          '切断位置
            
            For lSmplCnt = CUT_RS To CUT_MAXCNT - 1
                '>>>>> サンプル位置トップボトム明示マルチ品番対応  2009/11/18　SSS.Marushita
                xlSheet.Cells(lBaseRow + 20 + lSmplCnt, lBaseCol).Value = "'" & GetSampleStr(CStr(.iSmpKbnT(lSmplCnt)), _
                                                                                             CStr(.iSmpKbnB(lSmplCnt)))
                '>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                'xlSheet.Cells(lBaseRow + 19 + lSmplCnt, lBaseCol).Value = "'" & GetSampleStr(.sSmpl(lSmplCnt))
                '<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
                'xlSheet.Cells(19 + lSmplCnt, lBaseCol).Value = "'" & GetSampleStr(.sSmpl(lSmplCnt))
                '<<<<< サンプル位置トップボトム明示マルチ品番対応  2009/11/18　SSS.Marushita
            Next lSmplCnt
            
'>>>>> サンプル№の表示　2009/01/26　Marushita
            '>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
            xlSheet.Cells(lBaseRow + 38, lBaseCol).Value = "'" & .sSmpNoB      'サンプル№BOT
            xlSheet.Cells(lBaseRow + 39, lBaseCol).Value = "'" & .sSmpNoT      'サンプル№TOP
            ''''xlSheet.Cells(lBaseRow + 36, lBaseCol).Value = "'" & .sSmpNo       'サンプル№
            '<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
            xlSheet.Cells(lBaseRow + 40, lBaseCol).Value = "'" & .sMaisu       'サンプル指示(枚数)
            'xlSheet.Cells(32, lBaseCol).Value = "'" & .sSmpNo           'サンプル№
            'xlSheet.Cells(33, lBaseCol).Value = "'" & .sMaisu           'サンプル指示(抵抗)
'            xlSheet.Cells(32, lBaseCol).Value = "'" & .sMaisu           'サンプル指示(抵抗)
'<<<<< サンプル№の表示　2009/01/26　Marushita

            For lSmplCnt = 0 To val(.sMaisu) - 1
'>>>>> X線測定対応 2009/09/02 SETsw kubota ------------------
''>>>>> サンプル№表示対応　2009/01/26　Marushita
'                Call OvalWrite(xlSheet, lBaseRow + 39 + lSmplCnt * 5, lBaseCol)
'                'Call OvalWrite(xlSheet, 35 + lSmplCnt * 5, lBaseCol)
'                'Call OvalWrite(xlSheet, 34 + lSmplCnt * 5, lBaseCol)
''<<<<< サンプル№表示対応　2009/01/26　Marushita
                If .sSmplPic(lSmplCnt, 0) = "X" Then
                    'X線の場合、一ページ目の合計の下に○表示
                    Call OvalWrite(xlSheet, 42, 4)
                    xlSheet.Cells(43, 4).Value = .sSmplPic(lSmplCnt, 0)
                    lEditCntMax = 4     '通常の位置にX線を編集しない分、一枚多くループ
                Else
                    Call OvalWrite(xlSheet, lBaseRow + 42 + lSmplCnt * 5, lBaseCol)
                    lEditCntMax = 3
                End If
'<<<<< X線測定対応 2009/09/02 SETsw kubota ------------------
                
                For lPicCnt = 0 To 3
                    If .sSmplPic(lSmplCnt, lPicCnt) <> PIC_DUMMY Then
'                    If .sSmplPic(lSmplCnt, lPicCnt) <> "" Then
                    
                        Select Case lPicCnt
                        Case 0  '左上
'>>>>> X線測定対応 2009/09/02 SETsw kubota ------------------
                            If .sSmplPic(lSmplCnt, 0) <> "X" Then
'<<<<< X線測定対応 2009/09/02 SETsw kubota ------------------
                                xlSheet.Cells(lBaseRow + 43 + lSmplCnt * 5, lBaseCol).Value = .sSmplPic(lSmplCnt, lPicCnt)
                                
                                '右上が空でなければ右と下に罫線を引く
                                If .sSmplPic(lSmplCnt, 1) <> "" Then
                                    sStrCell = ConvXlsNumToA(lBaseCol) & CStr(lBaseRow + 42 + lSmplCnt * 5)
                                    sEndCell = ConvXlsNumToA(lBaseCol) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                    With xlSheet.Range(sStrCell, sEndCell)
                                        .Borders(xlEdgeRight).LineStyle = xlContinuous  '通常線
                                        .Borders(xlEdgeRight).Weight = xlThin
                                    End With
                                    sStrCell = ConvXlsNumToA(lBaseCol) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                    sEndCell = ConvXlsNumToA(lBaseCol) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                    With xlSheet.Range(sStrCell, sEndCell)
                                        .Borders(xlEdgeBottom).LineStyle = xlContinuous  '通常線
                                        .Borders(xlEdgeBottom).Weight = xlThin
                                    End With
                                End If
                                '左下が空でなければ下罫線を引く
                                If .sSmplPic(lSmplCnt, 2) <> "" Then
                                    sStrCell = ConvXlsNumToA(lBaseCol) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                    sEndCell = ConvXlsNumToA(lBaseCol + 1) & CStr(lBaseRow + 43 + lSmplCnt * 5)
                                    With xlSheet.Range(sStrCell, sEndCell)
                                        .Borders(xlEdgeBottom).LineStyle = xlContinuous '通常線
                                        .Borders(xlEdgeBottom).Weight = xlThin
                                    End With
                                End If
'>>>>> X線測定対応 2009/09/02 SETsw kubota ------------------
                            End If
'<<<<< X線測定対応 2009/09/02 SETsw kubota ------------------
                        
                        Case 1
'>>>>> サンプル№表示対応　2009/01/26　Marushita
                            xlSheet.Cells(lBaseRow + 43 + lSmplCnt * 5, lBaseCol + 1).Value = .sSmplPic(lSmplCnt, lPicCnt)
                            'xlSheet.Cells(36 + lSmplCnt * 5, lBaseCol + 1).Value = .sSmplPic(lSmplCnt, lPicCnt)
                            'xlSheet.Cells(35 + lSmplCnt * 5, lBaseCol + 1).Value = .sSmplPic(lSmplCnt, lPicCnt)
'<<<<< サンプル№表示対応　2009/01/26　Marushita
                        Case 2
'>>>>> サンプル№表示対応　2009/01/26　Marushita
                            xlSheet.Cells(lBaseRow + 44 + lSmplCnt * 5, lBaseCol).Value = .sSmplPic(lSmplCnt, lPicCnt)
                            'xlSheet.Cells(37 + lSmplCnt * 5, lBaseCol).Value = .sSmplPic(lSmplCnt, lPicCnt)
                            'xlSheet.Cells(36 + lSmplCnt * 5, lBaseCol).Value = .sSmplPic(lSmplCnt, lPicCnt)
'<<<<< サンプル№表示対応　2009/01/26　Marushita
                            '右下が空でなければ右罫線を引く
                            If .sSmplPic(lSmplCnt, 3) <> "" Then
'>>>>> サンプル№表示対応　2009/01/26　Marushita
                                sStrCell = ConvXlsNumToA(lBaseCol) & CStr(lBaseRow + 44 + lSmplCnt * 5)
                                sEndCell = ConvXlsNumToA(lBaseCol) & CStr(lBaseRow + 45 + lSmplCnt * 5)
                                'sStrCell = ConvXlsNumToA(lBaseCol) & CStr(37 + lSmplCnt * 5)
                                'sEndCell = ConvXlsNumToA(lBaseCol) & CStr(38 + lSmplCnt * 5)
                                'sStrCell = ConvXlsNumToA(lBaseCol) & CStr(36 + lSmplCnt * 5)
                                'sEndCell = ConvXlsNumToA(lBaseCol) & CStr(37 + lSmplCnt * 5)
'<<<<< サンプル№表示対応　2009/01/26　Marushita
                                With xlSheet.Range(sStrCell, sEndCell)
                                    .Borders(xlEdgeRight).LineStyle = xlContinuous  '通常線
                                    .Borders(xlEdgeRight).Weight = xlThin
                                End With
                            End If
                        Case 3
'>>>>> サンプル№表示対応　2009/01/26　Marushita
                            xlSheet.Cells(lBaseRow + 44 + lSmplCnt * 5, lBaseCol + 1).Value = .sSmplPic(lSmplCnt, lPicCnt)
                            'xlSheet.Cells(37 + lSmplCnt * 5, lBaseCol + 1).Value = .sSmplPic(lSmplCnt, lPicCnt)
                            'xlSheet.Cells(36 + lSmplCnt * 5, lBaseCol + 1).Value = .sSmplPic(lSmplCnt, lPicCnt)
'<<<<< サンプル№表示対応　2009/01/26　Marushita
                        End Select
                    
                    End If
                Next lPicCnt
                
''>>>>> サンプル№表示対応　2009/01/26　Marushita
'                '4枚編集したら抜ける⇒3枚編集したら抜ける
'                If lSmplCnt = 2 Then
'                'If lSmplCnt = 3 Then
''<<<<< サンプル№表示対応　2009/01/26　Marushita
                If lSmplCnt = lEditCntMax Then
                    Exit For
                End If
            Next

        End With
    Next lCnt

End Sub


'///////////////////////////////////////////////////
' @(f)
' 機能　　: 円を描く処理
' 返り値　: True  - 正常終了
' 　　　　  False - 異常終了
' 引き数　: xlSheet  - excelシートオブジェクト
' 　　　　  sSell - 基準となるセル
' 機能説明:
'///////////////////////////////////////////////////
Private Function OvalWrite(ByRef xlSheet As Object _
                         , ByVal lRow As Long _
                         , ByVal lCol As Long _
                         ) As Boolean
'Private Function OvalWrite(ByRef xlSheet As Excel.Worksheet _　  定義をObjectに変更

'>>>>>　Excel2007対応(円の位置がずれる問題に対応)　2009/01/29　Marushita
'Dim x As Object, MyZoom As Variant, VW As Variant
'>>>>>　Excelが開いている時にオブジェクトエラーとなる問題に対応　2009/05/18　Marushita
Dim objShape As Object
    
    With xlSheet
    Set objShape = .Shapes.AddShape(msoShapeOval _
                        , .Cells(lRow, lCol).Left _
                        , .Cells(lRow, lCol).Top _
                        , CON_X _
                        , CON_Y _
                       )
    End With
    '塗りつぶし、線の色・太さの指定
    objShape.Fill.Visible = msoFalse
    objShape.Line.ForeColor.SchemeColor = 0
    objShape.Line.Weight = 0.75
    
    'ズーム処理をしない（テンプレートを100%にして対応）　2009/05/18　Marushita
'    '現在のズーム倍率を保存
'    MyZoom = ActiveWindow.Zoom
'    '画面のズーム倍率を100%にする
'    VW = ActiveWindow.View
'    Application.ScreenUpdating = False
'    ActiveWindow.Zoom = 100
'    ActiveWindow.View = xlNormalView
'
'    With xlSheet
'         .Shapes.AddShape(msoShapeOval _
'                        , .Cells(lRow, lCol).Left _
'                        , .Cells(lRow, lCol).Top _
'                        , CON_X _
'                        , CON_Y _
'                       ).Select
'    End With
'    '塗りつぶし、線の色・太さの指定
'    With Selection.ShapeRange
'        .Fill.Visible = msoFalse
'        .Line.ForeColor.SchemeColor = 0
'        .Line.Weight = 0.75
'    End With
'
'    '元のズーム倍率に戻す
'    ActiveWindow.Zoom = MyZoom
'    ActiveWindow.View = VW
'
'    Application.ScreenUpdating = True
'<<<<<　Excel2007対応(円の位置がずれる問題に対応)　2009/01/29　Marushita
'    'セルのLeft、Top、Widthプロパティーを利用して位置決め
'    With xlSheet
'        'Shapeの描画
'        .Shapes.AddShape(msoShapeOval _
'                       , .Cells(lRow, lCol).Left _
'                       , .Cells(lRow, lCol).Top _
'                       , CON_X _
'                       , CON_Y _
'                       ).Fill.Visible = msoFalse
'
''        .Shapes.AddShape(msoShapeOval, BX, BY, EX, EY).Name = "aaa"
''        .Shapes("aaa").Fill.Visible = msoFalse
'    End With
'<<<<<　Excelが開いている時にオブジェクトエラーとなる問題に対応　2009/05/18　Marushita

End Function


'///////////////////////////////////////////////////
' @(f)
' 機能　　: サンプル表示文字列取得
' 返り値　: サンプル表示文字列
' 引き数　: サンプル数
' 機能説明: サンプル表示文字列取得
'///////////////////////////////////////////////////
'マルチ品番対応
'Private Function GetSampleStr(ByVal sSampleNum As String) As String
Private Function GetSampleStr(ByVal sSampleNumT As String, ByVal sSampleNumB As String) As String
    
    'サンプル数だけ"●"を表示する
    '変更の可能性有り、1なら○、2なら◎等も検討
    'GetSampleStr = String(val(sSampleNum), "●")
'>>>>> サンプル位置トップボトム明示マルチ品番対応  2009/11/12　SSS.Marushita
    'トップボトムの判断を追加
    If sSampleNumT = "1" And sSampleNumB = "1" Then         '両方あり
        GetSampleStr = "● ●"
    ElseIf sSampleNumT = "1" Then     'トップのみあり
        GetSampleStr = "　 ●"
    ElseIf sSampleNumB = "1" Then     'ボトムのみあり
        GetSampleStr = "● 　"
    End If
'<<<<< サンプル位置トップボトム明示マルチ品番対応  2009/11/12　SSS.Marushita
''>>>>> サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
'    'トップボトムの判断を追加
'    If val(sSampleNum) = 1 Then         'トップのみあり
'        GetSampleStr = "　 ●"
'    ElseIf val(sSampleNum) = 2 Then     'ボトムのみあり
'        GetSampleStr = "● 　"
'    ElseIf val(sSampleNum) = 3 Then     '両方あり
'        GetSampleStr = "● ●"
'    End If
''<<<<< サンプル位置トップボトム明示対応  2009/11/12　SSS.Marushita
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能　　: サンプル枚数を数える
' 返り値　: サンプル枚数
' 引き数　: 行
' 機能説明: サンプル枚数を数える
'///////////////////////////////////////////////////
Private Sub GetSamplePic(ByVal lRow As Long)
    
    Dim lCnt            As Long
    Dim lCnt2           As Long
    Dim lThick          As Long
    Dim lSmpl_1_4(1)    As Long     '1/4サンプルの数
    Dim lSmpl_1_2(1)    As Long     '1/2サンプルの数
    Dim lSmpl_4_4(1)    As Long     '4/4サンプルの数
    
    Dim lPic_1_4()      As Long     '1/4サンプル描画
    Dim lPic_1_2()      As Long     '1/4サンプル描画
    Dim lPic_4_4()      As Long     '1/4サンプル描画
    
    Dim lPicPos         As Long     'どの場所に書くか
    
    With mtPrintInfo.tMeisai(lRow)
        '厚み、形状が同じサンプルをまとめる
        For lCnt = CUT_RS To CUT_MAXCNT - 1
        
            '厚みは1.1,1.2のみを想定
            If mtPrintInfo.sThick(lCnt) = "1.1" Then
                lThick = 0
'>>>>> 米沢厚み1.3mm対応　2008/10/28　SET.Marushita
            'ElseIf mtPrintInfo.sThick(lCnt) = "1.2" Then
            '厚みが1.2,1.3は同じに
            ElseIf mtPrintInfo.sThick(lCnt) = "1.2" Or _
                   mtPrintInfo.sThick(lCnt) = "1.3" Then
'<<<<< 米沢厚み1.3mm対応　2008/10/28　SET.Marushita
                lThick = 1
            Else
                lThick = -1     'エラーに
                Call MsgOut(0, "(切断指示書)厚みエラー:厚み「" & mtPrintInfo.sThick(lCnt) & "」", ERR_DISP)
            End If
            
            '各形状で指示数をカウント
            If val(.sSmpl(lCnt)) > 0 Then
                Select Case mtPrintInfo.sShape(lCnt)
                Case "1/4"
                    lSmpl_1_4(lThick) = lSmpl_1_4(lThick) + val(.sSmpl(lCnt))
                    
                    '配列を確保
                    If lSmpl_1_4(lThick) >= lSmpl_1_4(0) _
                    And lSmpl_1_4(lThick) >= lSmpl_1_4(1) Then
                        ReDim Preserve lPic_1_4(1, lSmpl_1_4(lThick) - 1)
                    End If
                    
                    For lCnt2 = lSmpl_1_4(lThick) - val(.sSmpl(lCnt)) To lSmpl_1_4(lThick) - 1
                        'どのサンプルだったかを保存
                        lPic_1_4(lThick, lCnt2) = lCnt
                    Next lCnt2
                    
                Case "1/2"
                    lSmpl_1_2(lThick) = lSmpl_1_2(lThick) + val(.sSmpl(lCnt))
                
                    '配列を確保
                    If lSmpl_1_2(lThick) >= lSmpl_1_2(0) _
                    And lSmpl_1_2(lThick) >= lSmpl_1_2(1) Then
                        ReDim Preserve lPic_1_2(1, lSmpl_1_2(lThick) - 1)
                    End If
                    
                    For lCnt2 = lSmpl_1_2(lThick) - val(.sSmpl(lCnt)) To lSmpl_1_2(lThick) - 1
                        'どのサンプルだったかを保存
                        lPic_1_2(lThick, lCnt2) = lCnt
                    Next lCnt2
                
                Case "4/4"
                    lSmpl_4_4(lThick) = lSmpl_4_4(lThick) + val(.sSmpl(lCnt))
                
                    '配列を確保
                    If lSmpl_4_4(lThick) >= lSmpl_4_4(0) _
                    And lSmpl_4_4(lThick) >= lSmpl_4_4(1) Then
                        ReDim Preserve lPic_4_4(1, lSmpl_4_4(lThick) - 1)
                    End If
                    
                    For lCnt2 = lSmpl_4_4(lThick) - val(.sSmpl(lCnt)) To lSmpl_4_4(lThick) - 1
                        'どのサンプルだったかを保存
                        lPic_4_4(lThick, lCnt2) = lCnt
                    Next lCnt2
                
                End Select
            End If
        Next lCnt
        
        'サンプル枚数の計算と描画情報編集
        lPicPos = 0
        For lThick = 0 To 1       '厚み0:1.1、1:1.2
            
            '厚みが変わったら途中からは編集しない
            If lPicPos Mod 4 > 0 Then
                lPicPos = lPicPos + 4 - lPicPos Mod 4
            End If
            
            For lCnt = 0 To lSmpl_1_4(lThick) - 1
                If lPicPos < 16 Then    '編集は4枚まで
                    .sSmplPic(Fix(lPicPos / 4), lPicPos Mod 4) = mtPrintInfo.sPicStr(lPic_1_4(lThick, lCnt))
                End If
                lPicPos = lPicPos + 1
                If lPicPos < 16 Then    '編集は4枚まで
                    .sSmplPic(Fix(lPicPos / 4), lPicPos Mod 4) = PIC_DUMMY  '1/4にするマークをつける
                End If
            Next lCnt
            For lCnt = 0 To lSmpl_1_2(lThick) - 1
                '縦では割らない
                If lPicPos Mod 2 = 1 Then
                    lPicPos = lPicPos + 1
                End If
                If lPicPos < 16 Then    '編集は4枚まで
                    .sSmplPic(Fix(lPicPos / 4), lPicPos Mod 4) = mtPrintInfo.sPicStr(lPic_1_2(lThick, lCnt))
                End If
                lPicPos = lPicPos + 2
                If lPicPos < 16 Then    '編集は4枚まで
                    .sSmplPic(Fix(lPicPos / 4), lPicPos Mod 4) = PIC_DUMMY  '1/2にするマークをつける
                End If
            Next lCnt
            For lCnt = 0 To lSmpl_4_4(lThick) - 1
                '途中からは編集しない
                If lPicPos Mod 4 > 0 Then
                    lPicPos = lPicPos + 4 - lPicPos Mod 4
                End If
                If lPicPos < 16 Then    '編集は4枚まで
                    .sSmplPic(Fix(lPicPos / 4), lPicPos Mod 4) = mtPrintInfo.sPicStr(lPic_4_4(lThick, lCnt))
                End If
                lPicPos = lPicPos + 4
            Next lCnt
            
        Next lThick
        .sMaisu = Fix((lPicPos + 3) / 4)
        
    End With


End Sub


'///////////////////////////////////////////////////
' @(f)
' 機能    : 出力要否チェック
' 返り値  : True - 出力する  False - 出力しない
' 引き数  : なし
' 機能説明:
'///////////////////////////////////////////////////
Public Function ChkPrtYN() As Boolean

    Dim tKoda9          As typKoda9Data

    ChkPrtYN = False

    '管理コードマスタ取得
    If GetKanriCode("K", "01", TEMPLATENAME, tKoda9) = False Then
        Exit Function
    End If
    
    If tKoda9.sKCODE05A9 = "1" Then
        'KCODE05が"1"の場合に出力
        ChkPrtYN = True
    End If

End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : 出力要否チェック(再切断時)
' 返り値  : True - 出力する  False - 出力しない
' 引き数  : なし
' 機能説明:
'///////////////////////////////////////////////////
Public Function ChkPrtYN_S() As Boolean

    Dim tKoda9          As typKoda9Data

    ChkPrtYN_S = False

    '管理コードマスタ取得
    If GetKanriCode("K", "01", TEMPLATENAME, tKoda9) = False Then
        Exit Function
    End If
    
    If tKoda9.sKCODE04A9 = "1" Then
        'KCODE04が"1"の場合に出力
        ChkPrtYN_S = True
    End If

End Function




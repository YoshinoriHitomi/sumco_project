Attribute VB_Name = "s_cmzcXSDC_Basic"
Option Explicit
''加工仕様,加工実績構造体
''次元の低い方にMIN値を代入する
Private Type Judg_Kakou
    DPTH(2) As Double   '引き上げ結晶の場合データはひとつしか存在しない
    WIDH(2) As Double   '引き上げ結晶の場合データはひとつしか存在しない
    TOP(2) As Double
    TAIL(2) As Double
    pos As String * 2
End Type


'基本情報
Public Type typ_BasicCd
    nextCode     As String       '次工程
    nowCode      As String       '現工程
    changeHinban As String       '変更品番
    DIAMETER     As Double       '直径
    UPWEIGHT     As Double       '引上重量　04/09/30 ooba
    LENGFREE     As Integer      'ﾌﾘｰ長さ　04/12/22 ooba
End Type

Private Type typ_XSDCA_c_flg
    Entry As Boolean
    Furyo As Integer            '不良数
    Index_F As Long             '不良構造体ｲﾝﾃﾞｯｸｽ
    FuryoW As Long              '不良重量
End Type

'2002/08/29 追加
Private Type typ_KoteiInf        '工程情報
    Wkkt As String               '工程
    Maco As Integer               '処理回数
End Type
Private msL2Wkkt As String        '前前工程
Private msL2Maco As String        '前前工程処理回数
    


Private strNxtCd As String
Private strNowCd As String
Private strChgHin As String
Private dblDiameter As Double
Private regFLG As String
Private intFuryoLen As Integer          '不良長さ
Private intFuryoWei As Long             '不良重量
Private CC300Flg As Boolean             '現在工程かCC300かどうか
Private SXLflg As Boolean               'SXL管理(XSDCB)への登録(更新)を行うかどうか
Private PutWtFlg As Integer             '『引上重量を品番毎に按分した値』を登録するか　04/09/30 ooba
Private lPutWeight As Long              '引上重量　04/09/30 ooba
Private lTotalPwt As Long               '算出合計重量(引上重量)　04/09/30 ooba
Private iTotalLen As Integer            '品番合計長さ　04/09/30 ooba
Private iFreeLen As Integer             'ﾌﾘｰ長さ　04/12/22 ooba

'2002/09/05　m.tomita
Private CB410Flg As Boolean             '現在工程かCB410かどうか


Public Const FACTORYCD As Integer = 42  '製造工場


'概要     :新ＤＢへの書込み基本パターン処理を行う
'ﾊﾟﾗﾒｰﾀ   :変数名           ,IO  ,型                   ,説明
'         :p_typXSDC2_b     ,I   ,typ_XSDC2            ,分割結晶(ﾌﾞﾛｯｸ)前工程実績情報
'         :p_typXSDCA_b()   ,I   ,typ_XSDCA            ,分割結晶(品番)前工程実績情報
'         :p_typXSDC2_c     ,I   ,typ_XSDC2            ,分割結晶(ﾌﾞﾛｯｸ)登録情報
'         :p_typXSDCA_c()   ,I   ,typ_XSDCA            ,分割結晶(品番)登録情報
'         :p_typXSD4upd()   ,I   ,typ_XSDC4            ,不良内訳登録情報
'         :p_typBasicCd     ,I   ,typ_BasicCd          ,基本情報
'         :p_strErrMsg      ,O   ,String               ,ｴﾗｰﾒｯｾｰｼﾞ
'         :戻り値           ,O    ,FUNCTION_RETURN      ,新ＤＢへの書込みの成否
'説明     :

Public Function ExecBscProcess(p_typXSDC2_b As typ_XSDC2, _
                               p_typXSDCA_b() As typ_XSDCA, _
                               p_typXSDC2_c As typ_XSDC2, _
                               p_typXSDCA_c() As typ_XSDCA, _
                               p_typXSD4() As typ_XSDC4, _
                               p_typBasicCd As typ_BasicCd, _
                               p_strErrMsg As String) As FUNCTION_RETURN
On Error GoTo proc_err
    gErr.Push "s_cmzcXSDC_Basic.bas -- Function ExecBscProcess"
                
    Dim i As Long
    Dim j As Long
    Dim recCnt As Long             'レコード数
    Dim C4Cnt As Long              'レコード数 C4

    Dim typXSDC2upd_b As typ_XSDC2_Update
    Dim typXSDC2upd_c As typ_XSDC2_Update
    Dim typXSDCAupd_b() As typ_XSDCA_Update
    Dim typXSDCAupd_c() As typ_XSDCA_Update
    Dim typXSDC4upd() As typ_XSDC4_Update
    Dim typXSDC3upd_b() As typ_XSDC3_Update
    Dim typXSDC3upd_c() As typ_XSDC3_Update
    
    Dim intBlockLen As Integer         'ﾌﾞﾛｯｸの長さ
    Dim intSyoriKaisu As Integer    '現在処理回数
    
''a.基本情報のセット
    With p_typBasicCd
        strNxtCd = .nextCode
        strNowCd = .nowCode
        strChgHin = .changeHinban
        dblDiameter = .DIAMETER
    End With
    
 '現工程がCC300ならﾌﾗｸﾞｾｯﾄ
    If strNowCd = "CC300" Then
        CC300Flg = True
    Else
        CC300Flg = False
    End If
    
'▼2002/09/05 M.TOMITA 追加
 '現工程がCB410ならﾌﾗｸﾞｾｯﾄ
'    If strNowCd = "CB410" Then
    '受入取消処理対応　2003/09/01 ooba
    If strNowCd = "CB410" And strNxtCd = "CC600" Then
        CB410Flg = True
    Else
        CB410Flg = False
    End If
'▲2002/09/05 M.TOMITA
    
 'SXL登録ﾌﾗｸﾞｾｯﾄ
    SXLflg = False
    If strNowCd = "CC710" Or strNowCd = "CC720" Or strNowCd = "CW730" Then
    '抜試指示入力、WFｾﾝﾀｰ払出、結晶情報変更の場合
        SXLflg = True
    ElseIf strNowCd = "CC700" Then
    '結晶最終払出で SXLID が入ってる場合
        If Left(p_typXSDCA_b(1).SXLIDCA, 1) <> vbNullChar And Trim(p_typXSDCA_b(1).SXLIDCA) <> "" Then
            SXLflg = True
        End If
    End If
 
    '' 『引上重量を品番毎に按分した値』を登録するか判断する　04/09/30 ooba
    If strNowCd = "CC300" Or strNowCd = "CC310" Or (strNowCd = "CC400" And strNxtCd = "CC450") Then
        PutWtFlg = 1
    ElseIf strNowCd = "CC450" And (p_typXSDC2_b.CUTCNTC2 = "0" Or Mid(p_typXSDC2_b.CRYNUMC2, 10, 1) = "0") Then
        PutWtFlg = 2
    Else
        PutWtFlg = 0
    End If
    
    lTotalPwt = 0   '04/09/30 ooba
    iTotalLen = 0   '04/09/30 ooba
    iFreeLen = 0    '04/12/22 ooba
 
 '現在処理回数の取得
 ' 処理回数取得ロジック変更 2002/11/21 tuku
    If CC300Flg = True Or CB410Flg = True Then
        intSyoriKaisu = GetGNMACOC(p_typXSDCA_c(1).CRYNUMCA, strNxtCd) '結晶分割(品番)登録情報のﾌﾞﾛｯｸIDをキーに取得
        lPutWeight = CLng(p_typBasicCd.UPWEIGHT)  '引上重量取得　04/09/30 ooba
        iFreeLen = p_typBasicCd.LENGFREE  'ﾌﾘｰ長さ取得　04/12/22 ooba
        
    ElseIf Left(p_typXSDC2_b.CRYNUMC2, 1) <> vbNullChar Then
        intSyoriKaisu = GetGNMACOC(p_typXSDC2_b.CRYNUMC2, strNxtCd) '結晶分割(ﾌﾞﾛｯｸ)前工程情報のﾌﾞﾛｯｸIDをキーに取得
        lPutWeight = GetPutWeight(p_typXSDC2_b.XTALC2)  '引上重量取得　04/09/30 ooba
        
    ElseIf Left(p_typXSDC2_c.CRYNUMC2, 1) <> vbNullChar Then
        intSyoriKaisu = GetGNMACOC(p_typXSDC2_c.CRYNUMC2, strNxtCd) '結晶分割(ﾌﾞﾛｯｸ)登録情報のﾌﾞﾛｯｸIDをキーに取得
        lPutWeight = GetPutWeight(p_typXSDC2_c.XTALC2)  '引上重量取得　04/09/30 ooba
        
    Else
        If UBound(p_typXSD4) = 0 Then
           intSyoriKaisu = 1
        Else
           intSyoriKaisu = GetGNMACOC(p_typXSD4(1).XTALC4, strNxtCd)   '不良内訳登録情報のﾌﾞﾛｯｸIDをキーに取得
           lPutWeight = GetPutWeight(Left(p_typXSD4(1).XTALC4, 9) & "000")  '引上重量取得　04/09/30 ooba
        End If
    End If
    
  ' 同工程仕掛り対応　2002/11/25 tuku
    If strNxtCd = strNowCd Then
        intSyoriKaisu = intSyoriKaisu + 1
    End If
    
'前前工程情報の取得
'前々工程の取得方法変更　2002/11/22 tuku
    If CC300Flg = True Or CB410Flg = True Then                               'CC300,CB410
        msL2Wkkt = ""                      '前前工程
        msL2Maco = ""                      '前前工程処理回数
    ElseIf Left(p_typXSDCA_b(1).CRYNUMCA, 1) = vbNullChar Then            'bなし
        msL2Wkkt = ""                      '前前工程
        msL2Maco = ""                      '前前工程処理回数
    ElseIf Left(p_typXSDC2_c.CRYNUMC2, 1) = vbNullChar Then        'cなし
        msL2Wkkt = p_typXSDCA_b(1).NEWKNTCA   '前前工程
        msL2Maco = p_typXSDCA_b(1).NEMACOCA   '前前工程処理回数
    ElseIf p_typXSDCA_b(1).CRYNUMCA = p_typXSDC2_c.CRYNUMC2 Then       'b=c
        msL2Wkkt = p_typXSDCA_b(1).NEWKNTCA   '前前工程
        msL2Maco = p_typXSDCA_b(1).NEMACOCA   '前前工程処理回数
    Else                                                            'b<>c
        msL2Wkkt = ""                      '前前工程
        msL2Maco = ""                      '前前工程処理回数
    End If
    
'    If Left(p_typXSDC2_b.CRYNUMC2, 1) = vbNullChar Then            'bなし
'        msL2Wkkt = ""                      '前前工程
'        msL2Maco = ""                      '前前工程処理回数
'    ElseIf Left(p_typXSDC2_c.CRYNUMC2, 1) = vbNullChar Then        'cなし
'        msL2Wkkt = p_typXSDC2_b.NEWKNTC2   '前前工程
'        msL2Maco = p_typXSDC2_b.NEMACOC2   '前前工程処理回数
'    ElseIf p_typXSDC2_b.CRYNUMC2 = p_typXSDC2_c.CRYNUMC2 Then       'b=c
'        msL2Wkkt = p_typXSDC2_b.NEWKNTC2   '前前工程
'        msL2Maco = p_typXSDC2_b.NEMACOC2   '前前工程処理回数
'    Else                                                            'b<>c
'        msL2Wkkt = ""                      '前前工程
'        msL2Maco = ""                      '前前工程処理回数
'    End If
    
    
''b.前工程の実績(長さ等)を取得する
    '分割結晶(ﾌﾞﾛｯｸ)
    With typXSDC2upd_b
        .CRYNUMC2 = p_typXSDC2_b.CRYNUMC2                 'ﾌﾞﾛｯｸID
        .KCNTC2 = p_typXSDC2_b.KCNTC2                     '工程通過連番
        .INPOSC2 = p_typXSDC2_b.INPOSC2
        '.KCNTC2 = p_typXSDC2_b.KCNTC2 + 1                 '工程連番＋１してセット
'2002/10/16-----------------------------------------------------------------------------------▼1-①
'        If p_typXSDC2_b.GNWKNTC2 = strNxtCd Then          '同工程なら＋１してセット
'            '.GNMACOC2 = p_typXSDC2_b.GNMACOC2 + 1         '現在処理回数
'            .GNMACOC2 = intSyoriKaisu                     '現在処理回数
'            .NEWKNTC2 = p_typXSDC2_b.NEWKNTC2             '最終通過工程
'            .NEMACOC2 = p_typXSDC2_b.NEMACOC2             '最終通過処理回数
'        Else
'            .GNMACOC2 = 1
'            .NEWKNTC2 = p_typXSDC2_b.GNWKNTC2
'            .NEMACOC2 = p_typXSDC2_b.GNMACOC2
'        End If
'        .GNWKNTC2 = strNxtCd                                   '現在工程
 '2002/11/21 tuku 処理回数取得ロジック変更
        '前工程情報を更新させない(ここでは)
        .GNMACOC2 = p_typXSDC2_b.GNMACOC2                      '現在処理回数
        .NEWKNTC2 = p_typXSDC2_b.NEWKNTC2                      '最終通過工程
        .NEMACOC2 = p_typXSDC2_b.NEMACOC2                      '最終通過処理回数
        .GNWKNTC2 = p_typXSDC2_b.GNWKNTC2                      '現在工程
'2002/10/16-----------------------------------------------------------------------------------▲1-①
        .GNLC2 = p_typXSDC2_b.GNLC2                            '現在長さ
        .GNWC2 = p_typXSDC2_b.GNWC2                            '現在重量
'       .GNWC2 = WeightOfCylinder(dblDiameter, .GNLC2)         '現在重量
        '' SUMIT長さ／重量取得　04/09/30 ooba
        If PutWtFlg > 0 Then
            .SUMITLC2 = p_typXSDC2_b.SUMITLC2                  'SUMIT長さ
            .SUMITWC2 = p_typXSDC2_b.SUMITWC2                  'SUMIT重量
        End If
        .KAKOUBC2 = p_typXSDC2_b.KAKOUBC2                      '加工区分
        .LSTATBC2 = p_typXSDC2_b.LSTATBC2                      '最終状態区分
        .RSTATBC2 = p_typXSDC2_b.RSTATBC2                      '流動状態区分
        .LDFRBC2 = p_typXSDC2_b.LDFRBC2                        '格下区分
        .HOLDBC2 = p_typXSDC2_b.HOLDBC2                        'ﾎｰﾙﾄﾞ区分
        .KANKC2 = p_typXSDC2_b.KANKC2                          '完了区分
        If p_typXSDC2_b.CHGC2 <> 0 Then
            .CHGC2 = p_typXSDC2_b.CHGC2                        'ﾁｬｰｼﾞ量
            .KEIDAYC2 = p_typXSDC2_b.KEIDAYC2                  '計上日付
        End If
        '2003.06.11 (SPK)Y.katabami　優先度＆新規再切区分情報追加
        .CUTCNTC2 = p_typXSDC2_b.CUTCNTC2                      '新規再切区分
        .PRIORITYC2 = p_typXSDC2_b.PRIORITYC2                  '優先度
        .RPCRYNUMC2 = p_typXSDC2_b.RPCRYNUMC2                  '親ﾌﾞﾛｯｸID   2005/11
        .HOLDCC2 = p_typXSDC2_b.HOLDCC2                        'ﾎｰﾙﾄﾞコード  2006/03
        .HOLDKTC2 = p_typXSDC2_b.HOLDKTC2                      'ﾎｰﾙﾄﾞ工程  2006/03
    End With
                               
    '分割結晶(品番)
    recCnt = UBound(p_typXSDCA_b)
    ReDim typXSDCAupd_b(recCnt)
    For i = 1 To recCnt
        With typXSDCAupd_b(i)
            .CRYNUMCA = p_typXSDCA_b(i).CRYNUMCA             'ﾌﾞﾛｯｸID
            .HINBCA = p_typXSDCA_b(i).HINBCA                 '品番
            .INPOSCA = p_typXSDCA_b(i).INPOSCA               '結晶内開始位置
            .REVNUMCA = p_typXSDCA_b(i).REVNUMCA             '製品番号改訂番号
            .FACTORYCA = p_typXSDCA_b(i).FACTORYCA           '工場
            .OPECA = p_typXSDCA_b(i).OPECA                   '操業条件
            .KCKNTCA = p_typXSDCA_b(i).KCKNTCA               '工程連番
            '.KCKNTCA = typXSDC2upd_b.KCNTC2                  'ﾌﾞﾛｯｸの工程連番をセット
            '.GNMACOCA = intSyoriKaisu + (i - 1)              '現在処理回数(2ﾚｺｰﾄﾞ目から＋１)
            .GNMACOCA = p_typXSDCA_b(i).GNMACOCA              '現在処理回数
'2002/10/16-----------------------------------------------------------------------------------▼1-②
'            If p_typXSDCA_b(i).GNWKNTCA = strNxtCd Then      '同工程なら＋１してセット
'                '.GNMACOCA = p_typXSDCA_b(i).GNMACOCA + 1     '現在処理回数
'                .NEWKNTCA = p_typXSDCA_b(i).NEWKNTCA         '最終通過工程
'                .NEMACOCA = p_typXSDCA_b(i).NEMACOCA         '最終通過処理回数
'            Else
'                '.GNMACOCA = 1
'                .NEWKNTCA = p_typXSDCA_b(i).GNWKNTCA
'                .NEMACOCA = p_typXSDCA_b(i).GNMACOCA
'            End If
'            .GNWKNTCA = strNxtCd                             '現在工程
            '2002/11/21 tuku 処理回数取得ロジック変更
            '前工程情報を更新させない(ここでは)
            .NEWKNTCA = p_typXSDCA_b(i).NEWKNTCA                   '最終通過工程
            .NEMACOCA = p_typXSDCA_b(i).NEMACOCA                   '最終通過処理回数
            .GNWKNTCA = p_typXSDCA_b(i).GNWKNTCA                   '現在工程
'2002/10/16-----------------------------------------------------------------------------------▲1-②
            .SXLIDCA = p_typXSDCA_b(i).SXLIDCA               'SXLID
            .XTALCA = p_typXSDCA_b(i).XTALCA                 '結晶番号
            .GNLCA = p_typXSDCA_b(i).GNLCA                   '現在長さ
            .GNWCA = p_typXSDCA_b(i).GNWCA                   '現在重量
'           .GNWCA = WeightOfCylinder(dblDiameter, .GNLCA)   '現在重量
            '' SUMIT長さ／重量取得　04/09/30 ooba
            If PutWtFlg > 0 Then
                .SUMITLCA = p_typXSDCA_b(i).SUMITLCA         'SUMIT長さ
                .SUMITWCA = p_typXSDCA_b(i).SUMITWCA         'SUMIT重量
            End If
            .KAKOUBCA = p_typXSDCA_b(i).KAKOUBCA             '加工区分
            .LSTATBCA = p_typXSDCA_b(i).LSTATBCA             '最終状態区分
            .RSTATBCA = p_typXSDCA_b(i).RSTATBCA             '流動状態区分
            .LDFRBCA = p_typXSDCA_b(i).LDFRBCA               '格下区分
            .HOLDBCA = p_typXSDCA_b(i).HOLDBCA               'ﾎｰﾙﾄﾞ区分
            .KANKCA = p_typXSDCA_b(i).KANKCA                 '完了区分
            If p_typXSDCA_b(i).CHGCA <> 0 Then
                .CHGCA = p_typXSDCA_b(i).CHGCA               'ﾁｬｰｼﾞ量
                .KEIDAYCA = p_typXSDCA_b(i).KEIDAYCA         '計上日付
            End If
            '2003.06.11 (SPK)Y.katabami　代表品番＆新規再切区分情報追加
            .CUTCNTCA = p_typXSDCA_b(i).CUTCNTCA             '新規再切区分
            .HINBFLGCA = p_typXSDCA_b(i).HINBFLGCA           '代表品番フラグ
            .RPCRYNUMCA = p_typXSDCA_b(i).RPCRYNUMCA         '親ﾌﾞﾛｯｸID    2005/11
            .HOLDCCA = p_typXSDCA_b(i).HOLDCCA               'ﾎｰﾙﾄﾞコード   2006/03
            .HOLDKTCA = p_typXSDCA_b(i).HOLDKTCA             'ﾎｰﾙﾄﾞ工程   2006/03
        End With
    Next
    
    
''c.登録レコード情報作成
    Select Case strNowCd
        Case "CC300", "CC310", "CB410", "CC450", "CC600"
            '現在処理回数の取得(結晶分割(ﾌﾞﾛｯｸ)登録情報のﾌﾞﾛｯｸIDをキーに取得)
            If CC300Flg = False And p_typXSDC2_b.CRYNUMC2 <> p_typXSDC2_c.CRYNUMC2 _
                And Left(p_typXSDC2_c.CRYNUMC2, 1) <> vbNullChar Then  '切断対応
                intSyoriKaisu = GetGNMACOC(p_typXSDC2_c.CRYNUMC2, strNxtCd)
            End If
            

            
            '分割結晶(ﾌﾞﾛｯｸ)情報
            With typXSDC2upd_c
                .CRYNUMC2 = p_typXSDC2_c.CRYNUMC2                          'ﾌﾞﾛｯｸID
'                .KCNTC2 = 1                                                '工程連番に１をセット
                .XTALC2 = p_typXSDC2_c.XTALC2                              '結晶番号
                .INPOSC2 = p_typXSDC2_c.INPOSC2                            '結晶内開始位置
                .NEWKNTC2 = strNowCd                                       '最終通過工程
                .GNWKNTC2 = strNxtCd                                       '現在工程
                '.GNMACOC2 = 1                                              '現在処理回数
                .GNMACOC2 = intSyoriKaisu                                  '現在処理回数
                '2002/11/21 tuku 処理回数取得ロジック変更
                .NEMACOC2 = GetNEMACOC2(typXSDC2upd_c.CRYNUMC2)            '最終通過処理回数
                .GNDAYC2 = Format(Now, "yyyy/mm/dd hh:mm:ss")              '現在処理日時
                .GNLC2 = p_typXSDC2_c.GNLC2                                '現在長さ
'                .GNWC2 = WeightOfCylinder(dblDiameter, p_typXSDC2_c.GNLC2) '現在重量

                '' 重量登録変更　04/09/30 ooba START ===========================================>
                If PutWtFlg = 1 Then
                    '引上重量ｾｯﾄ
                    .GNWC2 = lPutWeight
                    'SUMIT長さ／重量をｾｯﾄ
                    .SUMITLC2 = p_typXSDC2_c.GNLC2                                'SUMIT長さ
                    .SUMITWC2 = WeightOfCylinder(dblDiameter, p_typXSDC2_c.GNLC2) 'SUMIT重量
                Else
                    '長さ計算重量ｾｯﾄ
                    .GNWC2 = WeightOfCylinder(dblDiameter, p_typXSDC2_c.GNLC2)
                End If
                '' 重量登録変更　04/09/30 ooba END =============================================>
                
'                .GNMC2 =                                                  ’現在枚数
                .KAKOUBC2 = p_typXSDC2_c.KAKOUBC2                          '加工区分
                If p_typXSDC2_c.CHGC2 <> 0 Then
                    .CHGC2 = p_typXSDC2_c.CHGC2                                'ﾁｬｰｼﾞ量
                    .KEIDAYC2 = p_typXSDC2_c.KEIDAYC2                          '計上日付
                End If
                '2003.06.11 (SPK)Y.katabami　優先度＆新規再切区分情報追加
                .CUTCNTC2 = p_typXSDC2_c.CUTCNTC2                      '新規再切区分
                .PRIORITYC2 = p_typXSDC2_c.PRIORITYC2                  '優先度
                .RPCRYNUMC2 = p_typXSDC2_c.RPCRYNUMC2                  '親ﾌﾞﾛｯｸID   2005/11
            
                '2006/03 HOLD
                .HOLDBC2 = p_typXSDC2_b.HOLDBC2
                .HOLDCC2 = p_typXSDC2_b.HOLDCC2
                .HOLDKTC2 = p_typXSDC2_b.HOLDKTC2
            End With
            
            
            '分割結晶(品番)情報
            recCnt = UBound(p_typXSDCA_c)
            ReDim typXSDCAupd_c(recCnt)
                        
            '' 品番全体長さｾｯﾄ　04/09/30 ooba
            If PutWtFlg > 0 Then
                For i = 1 To recCnt
                    iTotalLen = iTotalLen + p_typXSDCA_c(i).GNLCA
                Next
'                If iTotalLen = 0 Then
'                    ExecBscProcess = FUNCTION_RETURN_FAILURE
'                    p_strErrMsg = "重量登録失敗(XSDC)"
'                    GoTo proc_exit
'                End If
            End If
            
            For i = 1 To recCnt
                With typXSDCAupd_c(i)
                    .CRYNUMCA = p_typXSDCA_c(i).CRYNUMCA                          'ﾌﾞﾛｯｸID
                    .HINBCA = p_typXSDCA_c(i).HINBCA                              '品番
                    .INPOSCA = p_typXSDCA_c(i).INPOSCA                            '結晶内開始位置
                    .REVNUMCA = p_typXSDCA_c(i).REVNUMCA                          '製品番号改訂番号
                    .FACTORYCA = p_typXSDCA_c(i).FACTORYCA                        '工場
                    .OPECA = p_typXSDCA_c(i).OPECA                                '操業条件
                    .SXLIDCA = p_typXSDCA_c(i).SXLIDCA                            'SXLID
                    .XTALCA = p_typXSDCA_c(i).XTALCA                              '結晶番号
                    .NEWKNTCA = strNowCd                                          '最終通過工程
                    .NEMACOCA = GetNEMACOC(.CRYNUMCA, CInt(.INPOSCA))             '最終通過処理回数　’2002/08/29
                    .GNWKNTCA = strNxtCd                                          '現在工程
                    '.GNMACOCA = 1                                                 '現在処理回数
                    '.GNMACOCA = intSyoriKaisu + (i - 1)              '現在処理回数(2ﾚｺｰﾄﾞ目から＋１)
                    .GNMACOCA = intSyoriKaisu                                     '現在処理回数
                    .GNLCA = p_typXSDCA_c(i).GNLCA                                '現在長さ
'                    .GNWCA = WeightOfCylinder(dblDiameter, p_typXSDCA_c(i).GNLCA) '現在重量

                    '' 重量登録変更　04/09/30 ooba START =========================================>
                    If PutWtFlg = 1 Then
                        '引上重量ｾｯﾄ
                        If i = recCnt Then
                            '計算誤差を最後の重量に計上
                            .GNWCA = lPutWeight - lTotalPwt
                        Else
                            '引上重量を品番長さで按分
                            If iTotalLen <> 0 Then
                                .GNWCA = Int(lPutWeight * (p_typXSDCA_c(i).GNLCA / iTotalLen))
                            Else
                                .GNWCA = 0
                            End If
                            lTotalPwt = lTotalPwt + .GNWCA
                        End If
                        'SUMIT長さ／重量をｾｯﾄ
                        .SUMITLCA = p_typXSDCA_c(i).GNLCA                                'SUMIT長さ
                        .SUMITWCA = WeightOfCylinder(dblDiameter, p_typXSDCA_c(i).GNLCA) 'SUMIT重量
                    Else
                        '長さ計算重量ｾｯﾄ
                        .GNWCA = WeightOfCylinder(dblDiameter, p_typXSDCA_c(i).GNLCA)
                    End If
                    '' 重量登録変更　04/09/30 ooba END ===========================================>
                    
'                    .GNMCA =                                                     ’現在枚数
                    .KAKOUBCA = p_typXSDCA_c(i).KAKOUBCA                          '加工区分
                    If p_typXSDCA_c(i).CHGCA <> 0 Then
                        .CHGCA = p_typXSDCA_c(i).CHGCA                                'ﾁｬｰｼﾞ量
                        .KEIDAYCA = p_typXSDCA_c(i).KEIDAYCA                          '計上日付
                    End If
                    '2003.06.11 (SPK)Y.katabami　代表品番＆新規再切区分情報追加
                    .CUTCNTCA = p_typXSDCA_c(i).CUTCNTCA             '新規再切区分
                    .HINBFLGCA = p_typXSDCA_c(i).HINBFLGCA           '代表品番フラグ
                    .RPCRYNUMCA = p_typXSDCA_c(i).RPCRYNUMCA         '親ﾌﾞﾛｯｸID    2005/11
                
                    '2006/03 HOLD
                    .HOLDBCA = p_typXSDCA_b(1).HOLDBCA
                    .HOLDCCA = p_typXSDCA_b(1).HOLDCCA
                    .HOLDKTCA = p_typXSDCA_b(1).HOLDKTCA
                End With
            Next
            
            '登録レコード(c)がない場合
            If Left(typXSDC2upd_c.CRYNUMC2, 1) = vbNullChar And recCnt = 0 Then
                regFLG = "N"
            Else
                regFLG = "Y"
            End If
        Case "CC400"
            regFLG = "N"
        Case Else
            regFLG = "N"
            
    End Select
    
    
    '不良内訳情報
    intFuryoLen = 0
    intFuryoWei = 0
    
    recCnt = UBound(p_typXSD4)
    ReDim typXSDC4upd(recCnt)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  　CC400の円筒研削重量削除処理を組み込み                            　　   '
'    ただし、本当の不良発生を入力する時は、この処理では対応不可　　　　　　　　 '
'　　画面上で不良位置を入力が必要                                            '                                                    '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    平成１５年５月５日　　　　　　　　　濱　三雄
If strNowCd = "CC400" Then
   recCnt = UBound(p_typXSDCA_b)
   ReDim typXSDC4upd(recCnt)
    For i = 1 To recCnt
         With typXSDC4upd(i)
            .XTALC4 = p_typXSDCA_b(i).CRYNUMCA        'ﾌﾞﾛｯｸID
            .INPOSC4 = p_typXSDCA_b(i).INPOSCA        '結晶内開始位置
            .FCODEC4 = "088"                          '研削ロス
            .MACOC4 = intSyoriKaisu                   '現在処理回数
            .HINBC4 = p_typXSDCA_b(i).HINBCA          '品番
            .PUCUTLC4 = 0                             '不良長さ
'            .PUCUTWC4 = p_typXSDCA_b(i).GNWCA - WeightOfCylinder(dblDiameter, CInt(p_typXSDCA_b(i).GNLCA))
            
            '' 不良重量登録変更　04/09/30 ooba START ===========================================>
            If PutWtFlg = 1 Then
                .PUCUTWC4 = p_typXSDCA_b(i).SUMITWCA - WeightOfCylinder(dblDiameter, CInt(p_typXSDCA_b(i).SUMITLCA))
            Else
                .PUCUTWC4 = p_typXSDCA_b(i).GNWCA - WeightOfCylinder(dblDiameter, CInt(p_typXSDCA_b(i).GNLCA))
            End If
            '' 不良重量登録変更　04/09/30 ooba END =============================================>

        End With
        intFuryoLen = intFuryoLen + typXSDC4upd(i).PUCUTLC4
        intFuryoWei = intFuryoWei + typXSDC4upd(i).PUCUTWC4
    Next i
Else
    For i = 1 To recCnt
        With typXSDC4upd(i)
            .XTALC4 = p_typXSD4(i).XTALC4             'ﾌﾞﾛｯｸID
            .INPOSC4 = p_typXSD4(i).INPOSC4           '結晶内開始位置
            '.MACOC4 = intSyoriKaisu + (i - 1)         '現在処理回数(2ﾚｺｰﾄﾞ目から＋１)
            .MACOC4 = intSyoriKaisu                   '現在処理回数
            .HINBC4 = p_typXSD4(i).HINBC4             '品番
            .PUCUTLC4 = p_typXSD4(i).PUCUTLC4         '不良長さ
            .PUCUTWC4 = WeightOfCylinder(dblDiameter, CInt(p_typXSD4(i).PUCUTLC4))
        End With
        intFuryoLen = intFuryoLen + typXSDC4upd(i).PUCUTLC4
        intFuryoWei = intFuryoWei + typXSDC4upd(i).PUCUTWC4
  
    Next i
End If

        
''e.判断処理(登録値を登録構造体にセットする)
''f.登録処理

    If regFLG = "Y" Then
        'ﾌﾞﾛｯｸ長さを求める
        intBlockLen = 0
        If CC300Flg = True Then '現在工程CC300の場合
            recCnt = UBound(p_typXSDCA_c)
            For i = 1 To recCnt
                intBlockLen = intBlockLen + typXSDCAupd_c(i).GNLCA
            Next
        Else
            intBlockLen = p_typXSDC2_c.GNLC2
        End If
        
        If intBlockLen - intFuryoLen > 0 Then   'c-d
        ''セットパターンⅠ-①
            If SetPattern1(typXSDC2upd_b, typXSDCAupd_b, typXSDC2upd_c, typXSDCAupd_c, typXSDC4upd, _
                                                p_strErrMsg) = FUNCTION_RETURN_FAILURE Then
                ExecBscProcess = FUNCTION_RETURN_FAILURE
                p_strErrMsg = p_strErrMsg
                GoTo proc_exit
            End If
        Else
        ''セットパターンⅡ-① (死ロット処理)
            If SetPattern2(typXSDC2upd_b, typXSDCAupd_b, typXSDC2upd_c, typXSDCAupd_c, typXSDC4upd, _
                                                p_strErrMsg) = FUNCTION_RETURN_FAILURE Then
                ExecBscProcess = FUNCTION_RETURN_FAILURE
                p_strErrMsg = p_strErrMsg
                GoTo proc_exit
            End If
        End If
    Else
        
        If p_typXSDC2_b.GNLC2 - intFuryoLen > 0 Then   'b-d
        ''セットパターンⅠ-②
            If SetPattern1(typXSDC2upd_b, typXSDCAupd_b, typXSDC2upd_c, typXSDCAupd_c, typXSDC4upd, _
                                                p_strErrMsg) = FUNCTION_RETURN_FAILURE Then
                ExecBscProcess = FUNCTION_RETURN_FAILURE
                p_strErrMsg = p_strErrMsg
                GoTo proc_exit
            End If
        Else
        ''セットパターンⅡ-② (死ロット処理)
            If SetPattern2(typXSDC2upd_b, typXSDCAupd_b, typXSDC2upd_c, typXSDCAupd_c, typXSDC4upd, _
                                                    p_strErrMsg) = FUNCTION_RETURN_FAILURE Then
                    ExecBscProcess = FUNCTION_RETURN_FAILURE
                    p_strErrMsg = p_strErrMsg
                    GoTo proc_exit
            End If
        End If
    End If
    

    ExecBscProcess = FUNCTION_RETURN_SUCCESS



proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    gErr.HandleError
    ExecBscProcess = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要     :セットパターンⅠの処理を行う(通常処理)
'ﾊﾟﾗﾒｰﾀ   :変数名           ,IO  ,型                   ,説明
'         :p_Block_b        ,I   ,typ_XSDC2_Update     ,分割結晶(ﾌﾞﾛｯｸ)前工程実績情報
'         :p_Hinban_b()     ,I   ,typ_XSDCA_Update     ,分割結晶(品番)前工程実績情報
'         :p_Block_c        ,I   ,typ_XSDC2_Update     ,分割結晶(ﾌﾞﾛｯｸ)登録情報
'         :p_Hinban_c()     ,I   ,typ_XSDCA_Update     ,分割結晶(品番)登録情報
'         :p_Furyo()        ,I   ,typ_XSDC4_Update     ,不良内訳登録情報
'         :p_Error          ,O   ,String               ,ｴﾗｰﾒｯｾｰｼﾞ
'         :戻り値           ,O    ,FUNCTION_RETURN      ,新ＤＢへの書込みの成否
'説明     :
Private Function SetPattern1(p_Block_b As typ_XSDC2_Update, p_Hinban_b() As typ_XSDCA_Update, _
                        p_Block_c As typ_XSDC2_Update, p_Hinban_c() As typ_XSDCA_Update, _
                        p_Furyo() As typ_XSDC4_Update, p_Error As String) As FUNCTION_RETURN
                        
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcXSDC_Basic.bas -- Function SetPattern1"
                        
    Dim i As Long
    Dim j As Long
    Dim n As Long
    
    Dim strWhere As String                   '更新処理WHERE句
    Dim RET As FUNCTION_RETURN               '戻り値
    Dim strErrMsg As String                  'エラーメッセージ
    Dim recCnt As Long                       'レコード数
    Dim recCnt2 As Long                      'レコード数
    Dim recCnt3 As Long                      'レコード数
    Dim AccFlg As Boolean                    'bﾚｺｰﾄﾞﾌﾗｸﾞ(bとcが一致したかをﾁｪｯｸ)
    Dim XSDCA_c_flg() As typ_XSDCA_c_flg     'cﾚｺｰﾄﾞﾌﾗｸﾞ(登録/更新ﾁｪｯｸと対応する不良情報)
    Dim sDbName As String                    'ﾃｰﾌﾞﾙ名
    Dim typKotei() As typ_XSDC3_Update       '工程実績登録情報
    Dim strMotHin As String                  '振替品番(元)
    Dim typSXL() As typ_XSDCB_Update         '分割結晶(品番)情報
    Dim intLen As Integer                    'SXL長さ
    Dim intSyoriKaisu As Integer    '現在処理回数
    Dim rs2 As OraDynaset           'レコードセット
    Dim sql As String
    Dim fullHinban As tFullHinban            'ｸﾘｽﾀﾙｶﾀﾛｸﾞ格上げ時の最新品番　2003/11/10 ooba
    Dim p_Block_b_bar As typ_XSDC2_Update    'Bar出荷用分割結晶(ﾌﾞﾛｯｸ)登録情報　04/09/27 ooba
    Dim p_Hinban_b_bar() As typ_XSDCA_Update 'Bar出荷用分割結晶(品番)登録情報　04/09/27 ooba
    Dim typKotei_bar() As typ_XSDC3_Update   'Bar出荷用工程実績登録情報　04/09/27 ooba
    
    If regFLG = "Y" Then
''●セットパターンⅠ-①
        
        If CC300Flg = False Then   '現工程CC300はﾌﾞﾛｯｸの登録を行わない
    
    '≪分割結晶(ﾌﾞﾛｯｸ)-XSDC2≫
        
            '不良数があれば減算する
            p_Block_c.GNLC2 = p_Block_c.GNLC2 - intFuryoLen
            p_Block_c.GNWC2 = WeightOfCylinder(dblDiameter, CInt(p_Block_c.GNLC2))
    
            If p_Block_b.CRYNUMC2 = p_Block_c.CRYNUMC2 Then  'b=c
            '登録ﾚｺｰﾄﾞ情報(c)でUPDATE
                sDbName = "(XSDC2)"
                With p_Block_c
                    .KCNTC2 = p_Block_b.KCNTC2 + 1    '工程連番を＋１してセット
'2002/10/16 コメント--------------------------------------------------------------------------------▼1-③
'                    .GNMACOC2 = p_Block_b.GNMACOC2    '現在処理回数
'                    .NEWKNTC2 = p_Block_b.NEWKNTC2    '最終通過工程
'                    .NEMACOC2 = p_Block_b.NEMACOC2    '最終通過処理回数
'2002/10/16 コメント--------------------------------------------------------------------------------▲1-③
                End With
                strWhere = "WHERE CRYNUMC2 = '" & p_Block_b.CRYNUMC2 & "'"
                If UpdateXSDC2(p_Block_c, strWhere) = FUNCTION_RETURN_FAILURE Then
                    SetPattern1 = FUNCTION_RETURN_FAILURE
                    p_Error = GetMsgStr("EAPLY") & sDbName
                    GoTo proc_exit
                End If
                
            Else                                            'b<>c
            
            'bのﾚｺｰﾄﾞを生死ﾌﾗｸﾞ=1でUPDATE
                If Left(p_Block_b.CRYNUMC2, 1) <> vbNullChar Then '(ｂがある場合)
                    p_Block_b.LIVKC2 = "1"
                    p_Block_b.KCNTC2 = p_Block_b.KCNTC2 + 1    '工程連番を＋１してセット
                    strWhere = "WHERE CRYNUMC2 = '" & p_Block_b.CRYNUMC2 & "'"
                    If UpdateXSDC2(p_Block_b, strWhere) = FUNCTION_RETURN_FAILURE Then
                        SetPattern1 = FUNCTION_RETURN_FAILURE
                        p_Error = GetMsgStr("EAPLY") & sDbName
                        GoTo proc_exit
                    End If
                End If
                
            '登録ﾚｺｰﾄﾞ情報(c)をINSERT
                p_Block_c.KCNTC2 = 1    '工程連番に１をセット
                If strNowCd = "CC310" Then
                    p_Block_c.KCNTC2 = 2    '工程連番に２をセット(ﾌﾞﾛｯｸの前工程がない)
                End If
                'p_Block_c.GNMACOCA = 1         '現在処理回数
                If CreateXSDC2(p_Block_c, strErrMsg) = FUNCTION_RETURN_FAILURE Then
                    SetPattern1 = FUNCTION_RETURN_FAILURE
                    p_Error = strErrMsg
                    GoTo proc_exit
                End If
            End If
            
        End If
        
        
    '≪分割結晶(品番)-XSDCA≫
        sDbName = "(XSDCA)"
        
        recCnt = UBound(p_Hinban_b)
        recCnt2 = UBound(p_Hinban_c)
        recCnt3 = UBound(p_Furyo)
        
        ReDim XSDCA_c_flg(recCnt2)
        ReDim typKotei(recCnt2)
        
        '不良数があれば減算する
        For i = 1 To recCnt2
            For j = 1 To recCnt3
                If p_Hinban_c(i).CRYNUMCA = p_Furyo(j).XTALC4 _
                            And p_Hinban_c(i).HINBCA = p_Furyo(j).HINBC4 _
                            And p_Hinban_c(i).INPOSCA = p_Furyo(j).INPOSC4 Then
                    '長さセット
                    p_Hinban_c(i).GNLCA = p_Hinban_c(i).GNLCA - p_Furyo(j).PUCUTLC4
                    '重量セット
'                   p_Hinban_c(i).GNWCA = WeightOfCylinder(dblDiameter, CInt(p_Hinban_c(i).GNLCA))
                    p_Hinban_c(i).GNWCA = p_Hinban_c(i).GNWCA - p_Furyo(j).PUCUTWC4


                    XSDCA_c_flg(i).Furyo = p_Furyo(j).PUCUTLC4   '不良長さｾｯﾄ
                    XSDCA_c_flg(i).Index_F = j                   'indexｾｯﾄ

                    Exit For
                Else
                    XSDCA_c_flg(i).Index_F = -1                   'indexｾｯﾄ
                End If
            Next
        Next
        
        n = 0
        AccFlg = False
        For i = 1 To recCnt
            For j = 1 To recCnt2
                If p_Hinban_b(i).CRYNUMCA = p_Hinban_c(j).CRYNUMCA _
                                    And p_Hinban_b(i).HINBCA = p_Hinban_c(j).HINBCA _
                                    And p_Hinban_b(i).INPOSCA = p_Hinban_c(j).INPOSCA Then  'b=c
                    
                    '登録ﾚｺｰﾄﾞ情報(c)でUPDATE
                    With p_Hinban_c(j)
                        .KCKNTCA = p_Block_c.KCNTC2           'ﾌﾞﾛｯｸの工程連番をセット
'2002/10/16 コメント--------------------------------------------------------------------------------▼1-④
'                        .GNMACOCA = p_Hinban_b(i).GNMACOCA    '現在処理回数
'                        .NEWKNTCA = p_Hinban_b(i).NEWKNTCA    '最終通過工程
'                        .NEMACOCA = p_Hinban_b(i).NEMACOCA    '最終通過処理回数
'2002/10/16 コメント--------------------------------------------------------------------------------▲1-④
                    End With
                    
                    '不良内訳ﾚｺｰﾄﾞに品番情報をセット
                    If XSDCA_c_flg(j).Index_F > 0 Then
                        With p_Furyo(XSDCA_c_flg(j).Index_F)
                            .KCKNTC4 = p_Hinban_c(j).KCKNTCA        '工程連番
                            .REVNUMC4 = p_Hinban_c(j).REVNUMCA      '製品番号改訂番号
                            .FACTORYC4 = p_Hinban_c(j).FACTORYCA    '工場
                            .OPEC4 = p_Hinban_c(j).OPECA            '操業条件
                            .WKKTC4 = strNowCd
                        End With
                    End If
                    
                    strWhere = "WHERE CRYNUMCA = '" & p_Hinban_b(i).CRYNUMCA
                    strWhere = strWhere & "' AND HINBCA = '" & p_Hinban_b(i).HINBCA
                    strWhere = strWhere & "' AND INPOSCA = " & p_Hinban_b(i).INPOSCA
                    If UpdateXSDCA(p_Hinban_c(j), strWhere) = FUNCTION_RETURN_FAILURE Then
                        SetPattern1 = FUNCTION_RETURN_FAILURE
                        p_Error = GetMsgStr("EAPLY") & sDbName
                        GoTo proc_exit
                    End If
                    
                '≪工程実績登録≫
                    Call SetXSDC3(typKotei(n), p_Hinban_c(j), XSDCA_c_flg(j))
                    If CreateXSDC3(typKotei(n), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                        SetPattern1 = FUNCTION_RETURN_FAILURE
                        p_Error = strErrMsg
                        GoTo proc_exit
                    End If
                    
                    n = n + 1
                    AccFlg = True
                    XSDCA_c_flg(j).Entry = True   '更新ﾌﾗｸﾞ(一致するﾚｺｰﾄﾞアリ)
                    Exit For
'                Else
'                    XSDCA_c_flg(j).Entry = False   '更新ﾌﾗｸﾞ(一致するﾚｺｰﾄﾞナシ)
                End If
            Next
            
            '一致しないﾚｺｰﾄﾞ(b)の更新  b<>c
            If Left(p_Hinban_b(i).CRYNUMCA, 1) <> vbNullChar Then 'ｂがある場合
                If AccFlg = False Then
                    
                    'bのﾚｺｰﾄﾞを生死ﾌﾗｸﾞ=1でUPDATE
                    p_Hinban_b(i).LIVKCA = "1"
                    p_Hinban_b(i).KCKNTCA = p_Block_b.KCNTC2 + 1    '工程連番(前工程)を＋１してセット
                    strWhere = "WHERE CRYNUMCA = '" & p_Hinban_b(i).CRYNUMCA
                    strWhere = strWhere & "' AND HINBCA = '" & p_Hinban_b(i).HINBCA
                    strWhere = strWhere & "' AND INPOSCA = " & p_Hinban_b(i).INPOSCA
                    
                    If UpdateXSDCA(p_Hinban_b(i), strWhere) = FUNCTION_RETURN_FAILURE Then
                        SetPattern1 = FUNCTION_RETURN_FAILURE
                        p_Error = GetMsgStr("EAPLY") & sDbName
                        GoTo proc_exit
                    End If
                        
'                    '工程実績登録(ｂのレコード)
'                    Call SetXSDC3(typKotei(n), p_Hinban_b(i), XSDCA_c_flg(i))
'                    If CreateXSDC3(typKotei(n), strErrMsg) = FUNCTION_RETURN_FAILURE Then
'                        SetPattern1 = FUNCTION_RETURN_FAILURE
'                        p_Error = strErrMsg
'                        GoTo proc_exit
'                    End If
'                    n = n + 1
                Else
                    AccFlg = False
                End If
            End If
        Next
        
        ReDim p_Hinban_b_bar(recCnt2)  'CC300引上0用分割結晶(品番)登録情報　04/12/17 ooba
        ReDim typKotei_bar(recCnt2)    'CC300引上0用工程実績登録情報　04/12/17 ooba
        
        '一致しないﾚｺｰﾄﾞ(c)の登録
        For i = 1 To recCnt2
             If XSDCA_c_flg(i).Entry = False Then
                '工程連番取得
                '▼2002/09/05 M.TOMITA CB410時処理追加
'                If CC300Flg = True Then
                If CC300Flg = True Or CB410Flg = True Then
                '▲2002/09/05 M.TOMITA CB410時処理追加
                    p_Hinban_c(i).KCKNTCA = 1                  '工程連番に１をセット
                    
                ElseIf recCnt <> 0 And p_Hinban_b(1).CRYNUMCA = p_Hinban_c(1).CRYNUMCA Then
                    p_Hinban_c(i).KCKNTCA = p_Block_c.KCNTC2   'ﾌﾞﾛｯｸの工程連番をセット
                    
                Else
                    p_Hinban_c(i).KCKNTCA = 1                  '工程連番に１をセット
                End If
                
                '不良内訳ﾚｺｰﾄﾞに品番情報をセット
                If XSDCA_c_flg(i).Index_F > 0 Then
                    With p_Furyo(XSDCA_c_flg(i).Index_F)
                        .KCKNTC4 = p_Hinban_c(i).KCKNTCA       '工程連番
                        .REVNUMC4 = p_Hinban_c(i).REVNUMCA     '製品番号改訂番号
                        .FACTORYC4 = p_Hinban_c(i).FACTORYCA   '工場
                        .OPEC4 = p_Hinban_c(i).OPECA           '操業条件
                        .WKKTC4 = strNowCd
                    End With
                End If
                
'2002/08/29----------------------------------------------------
'                '登録
'                If CreateXSDCA(p_Hinban_c(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
'                    SetPattern1 = FUNCTION_RETURN_FAILURE
'                    p_Error = strErrMsg
'                    GoTo proc_exit
'                End If
                
                With p_Hinban_c(i)
                    If CheckUniqueRecord(.CRYNUMCA, .HINBCA, CInt(.INPOSCA)) = True Then
                        '登録
                        If CreateXSDCA(p_Hinban_c(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                            SetPattern1 = FUNCTION_RETURN_FAILURE
                            p_Error = strErrMsg
                            GoTo proc_exit
                        End If
                    Else
                        '更新
                        .LIVKCA = "0"      '生死フラグに"0"をセット
                        strWhere = "WHERE CRYNUMCA = '" & .CRYNUMCA
                        strWhere = strWhere & "' AND HINBCA = '" & .HINBCA
                        strWhere = strWhere & "' AND INPOSCA = " & .INPOSCA
                        
                        If UpdateXSDCA(p_Hinban_c(i), strWhere) = FUNCTION_RETURN_FAILURE Then
                            SetPattern1 = FUNCTION_RETURN_FAILURE
                            p_Error = GetMsgStr("EAPLY") & sDbName
                            GoTo proc_exit
                        End If
                    End If
                End With
'2002/08/29----------------------------------------------------
                
                
                '≪工程実績登録≫
                Call SetXSDC3(typKotei(n), p_Hinban_c(i), XSDCA_c_flg(i))
                If CreateXSDC3(typKotei(n), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                    SetPattern1 = FUNCTION_RETURN_FAILURE
                    p_Error = strErrMsg
                    GoTo proc_exit
                End If
                
                '' CC310の重量0実績作成　04/12/22 ooba START ================================>
                
                '工程CC300でﾌﾘｰ長が0の場合
                If CC300Flg = True And iFreeLen = 0 Then
                
                    p_Hinban_b_bar(i) = p_Hinban_c(i)
                    typKotei_bar(n) = typKotei(n)
                    
                    With p_Hinban_b_bar(i)
                        .KCKNTCA = CInt(p_Hinban_c(i).KCKNTCA) + 1  '工程連番
'                        .NEKKNTCA = p_Hinban_c(i).GNKKNTCA          '最終通過管理工程
'                        .NEWKNTCA = p_Hinban_c(i).GNWKNTCA          '最終通過工程
'                        .NEWKKBCA = p_Hinban_c(i).GNWKKBCA          '最終通過作業区分
'                        .NEMACOCA = p_Hinban_c(i).GNMACOCA          '最終通過処理回数
                        .GNWKNTCA = "CB210"                         '現在工程
                        .GNMACOCA = 1                               '現在処理回数
                        .LSTATBCA = "R"                             '最終状態区分(ﾘﾒﾙﾄ)
                        .RSTATBCA = "M"                             '流動状態区分(ﾘﾒﾙﾄ受入待ち)
                        .LDFRBCA = "1"                              '格下区分(ﾘﾒﾙﾄ)
                        .LIVKCA = "1"                               '生死区分(死ﾛｯﾄ)
                    End With
                    
                    strWhere = "WHERE CRYNUMCA = '" & p_Hinban_c(i).CRYNUMCA
                    strWhere = strWhere & "' AND HINBCA = '" & p_Hinban_c(i).HINBCA
                    strWhere = strWhere & "' AND INPOSCA = " & p_Hinban_c(i).INPOSCA
                    
                    If UpdateXSDCA(p_Hinban_b_bar(i), strWhere) = FUNCTION_RETURN_FAILURE Then
                        SetPattern1 = FUNCTION_RETURN_FAILURE
                        p_Error = GetMsgStr("EAPLY") & sDbName
                        GoTo proc_exit
                    End If
                    
                    With typKotei_bar(n)
                        .KCNTC3 = p_Hinban_b_bar(i).KCKNTCA         '工程連番
                        .KNKTC3 = Space(5)                          '管理工程
                        .WKKTC3 = Space(5)                          '工程
                        .WKKBC3 = p_Hinban_b_bar(i).NEWKKBCA        '作業区分
                        .MACOC3 = p_Hinban_b_bar(i).NEMACOCA        '処理回数
                        .FRKNKTC3 = p_Hinban_c(i).NEKKNTCA          '(受入)管理工程
                        .FRWKKTC3 = p_Hinban_c(i).NEWKNTCA          '(受入)工程
                        .FRWKKBC3 = p_Hinban_c(i).NEWKKBCA          '(受入)作業区分
                        .FRMACOC3 = p_Hinban_c(i).NEMACOCA          '(受入)処理回数
                        .TOWNKTC3 = p_Hinban_b_bar(i).GNKKNTCA      '(払出)管理工程
                        .TOWKKTC3 = p_Hinban_b_bar(i).GNWKNTCA      '(払出)工程
                        .TOMACOC3 = p_Hinban_b_bar(i).GNMACOCA      '(払出)処理回数
                        .FRLC3 = p_Hinban_c(i).GNLCA                '受入長さ
                        .FRWC3 = p_Hinban_c(i).GNWCA                '受入重量
                        .FULC3 = p_Hinban_c(i).GNLCA                '不良長さ
                        .FUWC3 = p_Hinban_c(i).GNWCA                '不良重量
                        .TOLC3 = "0"                                '払出長さ
                        .TOWC3 = "0"                                '払出重量
                        .SUMITLC3 = "0"                             'SUMIT長さ
                        .SUMITWC3 = "0"                             'SUMIT重量
                    End With
                    
                    If CreateXSDC3(typKotei_bar(n), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                        SetPattern1 = FUNCTION_RETURN_FAILURE
                        p_Error = strErrMsg
                        GoTo proc_exit
                    End If
                End If
                '' CC310の重量0実績作成　04/12/22 ooba END ==================================>
                    
                n = n + 1
            End If
        Next
        
    '≪不良内訳登録≫
        For i = 1 To recCnt3
            If p_Furyo(i).KCKNTC4 = "" Then
            '一致する品番が無かったﾚｺｰﾄﾞの情報をセット
                With p_Furyo(i)
                    If p_Block_b.CRYNUMC2 = p_Block_c.CRYNUMC2 Then  'b=c (ﾌﾞﾛｯｸ)
                        .KCKNTC4 = p_Block_b.KCNTC2 + 1    '工程連番を＋１してセット
                    Else
                        .KCKNTC4 = 1                       '工程連番に１をセット
                    End If
                    .WKKTC4 = strNowCd                     '工程
                End With
            End If
            
            '登録
            If p_Furyo(i).PUCUTLC4 <> 0 Then
                If CreateXSDC4(p_Furyo(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                    SetPattern1 = FUNCTION_RETURN_FAILURE
                    p_Error = strErrMsg
                    GoTo proc_exit
                End If
            End If
        Next
        
    Else
    
''●セットパターンⅠ-②
    
    '≪分割結晶(ﾌﾞﾛｯｸ)-XSDC2≫
        sDbName = "(XSDC2)"
        
        '長さ、重量セット
'        p_Block_b.GNLC2 = p_Block_b.GNLC2 - intFuryoLen
'        p_Block_b.GNWC2 = p_Block_b.GNWC2 - intFuryoWei
''       p_Block_b.GNWC2 = WeightOfCylinder(dblDiameter, CInt(p_Block_b.GNLC2))
        
        '' 長さ／重量登録変更　04/09/30 ooba START ================================>
        If PutWtFlg = 1 Then
            p_Block_b.SUMITLC2 = p_Block_b.SUMITLC2 - intFuryoLen
            p_Block_b.SUMITWC2 = p_Block_b.SUMITWC2 - intFuryoWei
        Else
            p_Block_b.GNLC2 = p_Block_b.GNLC2 - intFuryoLen
            p_Block_b.GNWC2 = p_Block_b.GNWC2 - intFuryoWei
        End If
        '' 長さ／重量登録変更　04/09/30 ooba END ==================================>
        
        '結晶最終払出の場合(最終状態区分ｾｯﾄ)
        With p_Block_b
            If strNowCd = "CC700" Then
                If strNxtCd = "CC705" Then
                    .LSTATBC2 = "B"   '最終状態区分(BAR出荷)
                    .KANKC2 = "2"     '完了区分(終了)
                    .LIVKC2 = "1"
                Else
                    .LSTATBC2 = "W"   '最終状態区分(WF出荷)
                End If
            ElseIf strNowCd = "CB320" Then '格上げ時に流動状態区分を通常に戻す 2002/12/17 tuku
                    .RSTATBC2 = "T"
            End If
        End With
        '工程コードを更新する
        With p_Block_b
            .GNWKNTC2 = strNxtCd         ' 現在工程
            .NEWKNTC2 = strNowCd         ' 最終通過工程
        End With
        
        '更新
        p_Block_b.KCNTC2 = p_Block_b.KCNTC2 + 1    '工程連番を＋１してセット
        
        '処理回数取得ロジック変更 2002/11/21 tuku  START
        intSyoriKaisu = GetGNMACOC(p_Block_b.CRYNUMC2, strNxtCd) '結晶分割(ﾌﾞﾛｯｸ)前工程情報のﾌﾞﾛｯｸIDをキーに取得
        If strNxtCd = strNowCd Then         ' 同工程仕掛り対応　2002/11/25 tuku
            intSyoriKaisu = intSyoriKaisu + 1
        End If
        p_Block_b.GNMACOC2 = intSyoriKaisu
        p_Block_b.NEMACOC2 = GetNEMACOC2(p_Block_b.CRYNUMC2)
        '処理回数取得ロジック変更 2002/11/21 tuku  END
        
        strWhere = "WHERE CRYNUMC2 = '" & p_Block_b.CRYNUMC2 & "'"

        If UpdateXSDC2(p_Block_b, strWhere) = FUNCTION_RETURN_FAILURE Then
            SetPattern1 = FUNCTION_RETURN_FAILURE
            p_Error = GetMsgStr("EAPLY") & sDbName
            GoTo proc_exit
        End If
        
        '' CC705(Bar出荷)実績作成　04/09/27 ooba START ================================>
        If strNowCd = "CC700" And strNxtCd = "CC705" Then
        
            p_Block_b_bar = p_Block_b
            
            With p_Block_b_bar
                .KCNTC2 = CInt(p_Block_b.KCNTC2) + 1        '工程連番
                .NEKKNTC2 = p_Block_b.GNKKNTC2              '最終通過管理工程
                .NEWKNTC2 = p_Block_b.GNWKNTC2              '最終通過工程
                .NEWKKBC2 = p_Block_b.GNWKKBC2              '最終通過作業区分
                .NEMACOC2 = p_Block_b.GNMACOC2              '最終通過処理回数
                .GNWKNTC2 = Space(5)                        '現在工程(ｽﾍﾟｰｽ)
                .GNMACOC2 = 1                               '現在処理回数
            End With
            
            If UpdateXSDC2(p_Block_b_bar, strWhere) = FUNCTION_RETURN_FAILURE Then
                SetPattern1 = FUNCTION_RETURN_FAILURE
                p_Error = GetMsgStr("EAPLY") & sDbName
                GoTo proc_exit
            End If
        End If
        '' CC705(Bar出荷)実績作成　04/09/27 ooba END ==================================>
        
            
    '≪分割結晶(品番)-XSDCA≫
        sDbName = "(XSDCA)"
        
        recCnt = UBound(p_Hinban_b)   '分割結晶(品番)-b
        ReDim typKotei(recCnt)
        ReDim XSDCA_c_flg(recCnt)
        ReDim p_Hinban_b_bar(recCnt)  'Bar出荷用分割結晶(品番)登録情報　04/09/27 ooba
        ReDim typKotei_bar(recCnt)    'Bar出荷用工程実績登録情報　04/09/27 ooba
       
        recCnt3 = UBound(p_Furyo)     '不良実績-d
        
        '長さ、重量セット
        For i = 1 To recCnt
            For j = 1 To recCnt3
                If p_Hinban_b(i).CRYNUMCA = p_Furyo(j).XTALC4 _
                            And p_Hinban_b(i).HINBCA = p_Furyo(j).HINBC4 _
                            And p_Hinban_b(i).INPOSCA = p_Furyo(j).INPOSC4 Then
                            
                    '不良長さをマイナス
'                    p_Hinban_b(i).GNLCA = p_Hinban_b(i).GNLCA - p_Furyo(j).PUCUTLC4
''                   p_Hinban_b(i).GNWCA = WeightOfCylinder(dblDiameter, CInt(p_Hinban_b(i).GNLCA))
'                    p_Hinban_b(i).GNWCA = p_Hinban_b(i).GNWCA - p_Furyo(j).PUCUTWC4
                    
                    '' 長さ／重量登録変更　04/09/30 ooba START ================================>
                    If PutWtFlg = 1 Then
                        p_Hinban_b(i).SUMITLCA = p_Hinban_b(i).SUMITLCA - p_Furyo(j).PUCUTLC4
                        p_Hinban_b(i).SUMITWCA = p_Hinban_b(i).SUMITWCA - p_Furyo(j).PUCUTWC4
                    Else
                        p_Hinban_b(i).GNLCA = p_Hinban_b(i).GNLCA - p_Furyo(j).PUCUTLC4
                        p_Hinban_b(i).GNWCA = p_Hinban_b(i).GNWCA - p_Furyo(j).PUCUTWC4
                    End If
                    '' 長さ／重量登録変更　04/09/30 ooba END ==================================>
                    
                     '不良長さｾｯﾄ
                    XSDCA_c_flg(i).Furyo = p_Furyo(j).PUCUTLC4
                    XSDCA_c_flg(i).FuryoW = p_Furyo(j).PUCUTWC4
                    'indexをｾｯﾄ
                    XSDCA_c_flg(i).Index_F = j
                    Exit For
                Else
                    XSDCA_c_flg(i).Index_F = -1
                End If
            Next
        Next
        
        n = 0
        For i = 1 To recCnt
            'ﾌﾞﾛｯｸの工程連番をセット
            p_Hinban_b(i).KCKNTCA = p_Block_b.KCNTC2
            
            '不良内訳ﾚｺｰﾄﾞに品番情報をセット
            If XSDCA_c_flg(i).Index_F > 0 Then
                With p_Furyo(XSDCA_c_flg(i).Index_F)
                    .KCKNTC4 = p_Hinban_b(i).KCKNTCA      '工程連番
                    .REVNUMC4 = p_Hinban_b(i).REVNUMCA    '製品番号改訂番号
                    .FACTORYC4 = p_Hinban_b(i).FACTORYCA  '工場
                    .OPEC4 = p_Hinban_b(i).OPECA          '操業条件
                    .WKKTC4 = strNowCd                    '工程
                End With
            End If
            
            '更新
            strWhere = "WHERE CRYNUMCA = '" & p_Hinban_b(i).CRYNUMCA
            strWhere = strWhere & "' AND HINBCA = '" & p_Hinban_b(i).HINBCA
            strWhere = strWhere & "' AND INPOSCA = " & p_Hinban_b(i).INPOSCA
            
            'ｸﾘｽﾀﾙｶﾀﾛｸﾞ格上の場合
            If strNowCd = "CB320" And strChgHin <> "" Then
                With p_Hinban_b(i)
                    .LIVKCA = "1"          '生死ﾌﾗｸﾞ1をｾｯﾄ
                End With
            '結晶最終払出の場合(最終状態区分ｾｯﾄ)
            Else
                If strNowCd = "CC700" Then
                    With p_Hinban_b(i)
                        If strNxtCd = "CC705" Then
                            .LSTATBCA = "B"   '最終状態区分(BAR出荷)
                            .KANKCA = "2"     '完了区分(終了)
                            .LIVKCA = "1"     '生死ﾌﾗｸﾞ1をｾｯﾄ
                             typKotei(n).PAYCLASSC3 = "1"
                        Else
                            .LSTATBCA = "W"   '最終状態区分(WF出荷)
                             typKotei(n).PAYCLASSC3 = "0"
                        End If
                    End With
                End If
                
                With p_Hinban_b(i)
                    .GNWKNTCA = strNxtCd         ' 現在工程
                    .NEWKNTCA = strNowCd         ' 最終通過工程
                End With
                
                '処理回数取得ロジック変更 2002/11/21 tuku  START
                intSyoriKaisu = GetGNMACOC(p_Block_b.CRYNUMC2, strNxtCd) '結晶分割(ﾌﾞﾛｯｸ)前工程情報のﾌﾞﾛｯｸIDをキーに取得
                If strNxtCd = strNowCd Then                             ' 同工程仕掛り対応　2002/11/25 tuku
                    intSyoriKaisu = intSyoriKaisu + 1
                End If
                p_Hinban_b(i).GNMACOCA = intSyoriKaisu
                p_Hinban_b(i).NEMACOCA = GetNEMACOC(p_Hinban_b(i).CRYNUMCA, CInt(p_Hinban_b(i).INPOSCA))
                '処理回数取得ロジック変更 2002/11/21 tuku  END
            End If
        
            If UpdateXSDCA(p_Hinban_b(i), strWhere) = FUNCTION_RETURN_FAILURE Then
                SetPattern1 = FUNCTION_RETURN_FAILURE
                p_Error = GetMsgStr("EAPLY") & sDbName
                GoTo proc_exit
            End If
            
            '' CC705(Bar出荷)実績作成　04/09/27 ooba START ================================>
            If strNowCd = "CC700" And strNxtCd = "CC705" Then
            
                p_Hinban_b_bar(i) = p_Hinban_b(i)
                
                With p_Hinban_b_bar(i)
                    .KCKNTCA = CInt(p_Hinban_b(i).KCKNTCA) + 1  '工程連番
                    .NEKKNTCA = p_Hinban_b(i).GNKKNTCA          '最終通過管理工程
                    .NEWKNTCA = p_Hinban_b(i).GNWKNTCA          '最終通過工程
                    .NEWKKBCA = p_Hinban_b(i).GNWKKBCA          '最終通過作業区分
                    .NEMACOCA = p_Hinban_b(i).GNMACOCA          '最終通過処理回数
                    .GNWKNTCA = Space(5)                        '現在工程(ｽﾍﾟｰｽ)
                    .GNMACOCA = 1                               '現在処理回数
                End With
                
                If UpdateXSDCA(p_Hinban_b_bar(i), strWhere) = FUNCTION_RETURN_FAILURE Then
                    SetPattern1 = FUNCTION_RETURN_FAILURE
                    p_Error = GetMsgStr("EAPLY") & sDbName
                    GoTo proc_exit
                End If
            End If
            '' CC705(Bar出荷)実績作成　04/09/27 ooba END ==================================>
        
        
'2002/08/29----------------------------------------------------
            'ｸﾘｽﾀﾙｶﾀﾛｸﾞ格上の場合 品番変更
            If strNowCd = "CB320" And strChgHin <> "" Then
                With p_Hinban_b(i)
                    strMotHin = .HINBCA & .REVNUMCA & .FACTORYCA & .OPECA
'''                    .HINBCA = strChgHin    'changeHinbanをセット

                    ''最新品番を取得するように変更　2003/11/10 ooba　START
                    If GetLastHinban(strChgHin, fullHinban) = FUNCTION_RETURN_FAILURE Then
                        SetPattern1 = FUNCTION_RETURN_FAILURE
                        p_Error = GetMsgStr(EHIN0)
                        GoTo proc_exit
                    End If
                    .HINBCA = fullHinban.hinban
                    .REVNUMCA = fullHinban.mnorevno
                    .FACTORYCA = fullHinban.factory
                    .OPECA = fullHinban.opecond
                    ''最新品番を取得するように変更　2003/11/10 ooba　END
                    
                    .LIVKCA = "0"          '生死ﾌﾗｸﾞ0をｾｯﾄ
'2002/10/16 追加-------------------------------------------------------------------▼3-③
                    .GNWKNTCA = strNxtCd         ' 現在工程
                    .NEWKNTCA = strNowCd         ' 最終通過工程
'2002/10/16 追加-------------------------------------------------------------------▲3-③
                    .RSTATBCA = "T" '格上げ時に流動状態区分を通常に戻す 2002/12/17 tuku

                    '処理回数取得ロジック変更 2002/11/21 tuku  START
                    p_Hinban_b(i).GNMACOCA = GetGNMACOC(p_Block_b.CRYNUMC2, strNxtCd) '結晶分割(ﾌﾞﾛｯｸ)前工程情報のﾌﾞﾛｯｸIDをキーに取得
                    p_Hinban_b(i).NEMACOCA = GetNEMACOC(p_Hinban_b(i).CRYNUMCA, CInt(p_Hinban_b(i).INPOSCA))
                    '処理回数取得ロジック変更 2002/11/21 tuku  END
                    
                    If CheckUniqueRecord(.CRYNUMCA, .HINBCA, CInt(.INPOSCA)) = True Then
                        '登録
                        If CreateXSDCA(p_Hinban_b(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                            SetPattern1 = FUNCTION_RETURN_FAILURE
                            p_Error = strErrMsg
                            GoTo proc_exit
                        End If
                    Else
                        '更新
                        strWhere = "WHERE CRYNUMCA = '" & p_Hinban_b(i).CRYNUMCA
                        strWhere = strWhere & "' AND HINBCA = '" & p_Hinban_b(i).HINBCA
                        strWhere = strWhere & "' AND INPOSCA = " & p_Hinban_b(i).INPOSCA
                        
                        If UpdateXSDCA(p_Hinban_b(i), strWhere) = FUNCTION_RETURN_FAILURE Then
                            SetPattern1 = FUNCTION_RETURN_FAILURE
                            p_Error = GetMsgStr("EAPLY") & sDbName
                            GoTo proc_exit
                        End If
                    End If
                End With
            End If
'2002/08/29----------------------------------------------------

            
    '≪工程実績登録-XSDC3≫
            Call SetXSDC3(typKotei(n), p_Hinban_b(i), XSDCA_c_flg(i))
            If strNowCd = "CC400" Then
'               typKotei(n).FRLC3 = p_Hinban_b(i).GNLCA
'               typKotei(n).FRWC3 = p_Hinban_b(i).GNWCA
'               typKotei(n).FUWC3 = XSDCA_c_flg(i).FuryoW
                '' 不良重量登録変更　04/09/30 ooba START =================================>
                typKotei(n).LOSLC3 = XSDCA_c_flg(i).FuryoW          'ロス重量
                If PutWtFlg <> 1 Then
                    typKotei(n).FUWC3 = XSDCA_c_flg(i).FuryoW       '不良重量
                End If
                '' 不良重量登録変更　04/09/30 ooba END ===================================>
            End If
            If strNowCd = "CB320" And strChgHin <> "" Then
                typKotei(n).MOTHINC3 = strMotHin   '元品番をセット
            End If
            If CreateXSDC3(typKotei(n), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                SetPattern1 = FUNCTION_RETURN_FAILURE
                p_Error = strErrMsg
                GoTo proc_exit
            End If
            
            '' CC705(Bar出荷)実績作成　04/09/27 ooba START ================================>
            If strNowCd = "CC700" And strNxtCd = "CC705" Then
            
                typKotei_bar(n) = typKotei(n)
                
                With typKotei_bar(n)
                    .KCNTC3 = p_Hinban_b_bar(i).KCKNTCA         '工程連番
                    .KNKTC3 = p_Hinban_b_bar(i).NEKKNTCA        '管理工程
                    .WKKTC3 = p_Hinban_b_bar(i).NEWKNTCA        '工程
                    .WKKBC3 = p_Hinban_b_bar(i).NEWKKBCA        '作業区分
                    .MACOC3 = p_Hinban_b_bar(i).NEMACOCA        '処理回数
                    .FRKNKTC3 = p_Hinban_b(i).NEKKNTCA          '(受入)管理工程
                    .FRWKKTC3 = p_Hinban_b(i).NEWKNTCA          '(受入)工程
                    .FRWKKBC3 = p_Hinban_b(i).NEWKKBCA          '(受入)作業区分
                    .FRMACOC3 = p_Hinban_b(i).NEMACOCA          '(受入)処理回数
                    .TOWNKTC3 = p_Hinban_b_bar(i).GNKKNTCA      '(払出)管理工程
                    .TOWKKTC3 = p_Hinban_b_bar(i).GNWKNTCA      '(払出)工程
                    .TOMACOC3 = p_Hinban_b_bar(i).GNMACOCA      '(払出)処理回数
                    .FRLC3 = p_Hinban_b(i).GNLCA                '受入長さ
                    .FRWC3 = p_Hinban_b(i).GNWCA                '受入重量
                    .FULC3 = 0                                  '不良長さ
                    .FUWC3 = 0                                  '不良重量
                    .TOLC3 = p_Hinban_b_bar(i).GNLCA            '払出長さ
                    .TOWC3 = p_Hinban_b_bar(i).GNWCA            '払出重量
                End With
                
                If CreateXSDC3(typKotei_bar(n), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                    SetPattern1 = FUNCTION_RETURN_FAILURE
                    p_Error = strErrMsg
                    GoTo proc_exit
                End If
            End If
            '' CC705(Bar出荷)実績作成　04/09/27 ooba END ==================================>
            
            n = n + 1
        Next
        
        
    '≪不良実績-XSDC4≫
        For i = 1 To recCnt3
            If p_Furyo(i).KCKNTC4 = "" Then
            '一致する品番が無かったﾚｺｰﾄﾞの情報をセット
                With p_Furyo(i)
                    .KCKNTC4 = p_Block_b.KCNTC2        '工程連番
                    .WKKTC4 = strNowCd                 '工程
                End With
            End If
            If strNowCd = "CC400" Then
               If p_Furyo(i).PUCUTWC4 <> 0 Then
                  If CreateXSDC4(p_Furyo(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                      SetPattern1 = FUNCTION_RETURN_FAILURE
                      p_Error = strErrMsg
                      GoTo proc_exit
                  End If
               End If
            Else
               If p_Furyo(i).PUCUTLC4 <> 0 Then
                  If CreateXSDC4(p_Furyo(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                      SetPattern1 = FUNCTION_RETURN_FAILURE
                      p_Error = strErrMsg
                      GoTo proc_exit
                  End If
               End If
            End If
        Next
        
        
    '≪分割結晶(SXL)登録-XSDCB≫      ＊分割結晶(品番)をSXL単位に集約
        If SXLflg = True Then
            sDbName = "(XSDCB)"
            Call MakeSXLinfo(p_Hinban_b, typSXL)  '分割結晶(SXL)情報作成
            recCnt = UBound(typSXL)

            For i = 1 To recCnt
                If typSXL(i).SXLIDCB <> "" Then
                    If typSXL(i).LSTCCB = "W" Then
                       typSXL(i).LSTCCB = "T"   '最終状態区分(通常)をｾｯﾄ
                    End If
                    'Sumit連携処理変更によりＤＢ追加　セット現在工程、最終通過工程
                    If strNowCd = "CC700" Then
                        typSXL(i).NEWKNTCB = "CC700"
                        typSXL(i).GNWKNTCB = "CW750"
                    End If
'変更 SystamBrain 2003/10/09 ---------------------------------------------------> START
''''                    If strNowCd = "CC710" Then
''''                        typSXL(i).NEWKNTCB = "CC710"
''''                        typSXL(i).GNWKNTCB = "CST02"
''''                    End If
                    If strNowCd = "CC720" Then
                        typSXL(i).NEWKNTCB = "CC720"
                        typSXL(i).GNWKNTCB = "CST02"
                    End If
'変更 SystamBrain 2003/10/09 ---------------------------------------------------> END
                    'ＤＢ追加　　　　　　　　濱　三雄　　平成１５年５月１日
                    If CheckSXLrecord(typSXL(i).SXLIDCB, intLen) = 0 Then
                    'ﾚｺｰﾄﾞがない場合(登録)
                    
                        If CreateXSDCB(typSXL(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                            SetPattern1 = FUNCTION_RETURN_FAILURE
                            p_Error = strErrMsg
                            GoTo proc_exit
                        End If
                        
                    Else
                    'ﾚｺｰﾄﾞがある場合(更新)
                    
'                        '抜試指示入力、結晶最終払出入力の場合、長さをプラス
                        typSXL(i).LENCB = typSXL(i).LENCB + intLen

                        
                        strWhere = "WHERE SXLIDCB = '" & typSXL(i).SXLIDCB & "'"
                        If UpdateXSDCB(typSXL(i), strWhere) = FUNCTION_RETURN_FAILURE Then
                            SetPattern1 = FUNCTION_RETURN_FAILURE
                            p_Error = GetMsgStr("EAPLY") & sDbName
                            GoTo proc_exit
                        End If
                    End If
                End If
            Next
            
        End If
    End If
    
    SetPattern1 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    SetPattern1 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'概要     :分割結晶(SXL)情報をセットする
'ﾊﾟﾗﾒｰﾀ   :変数名           ,IO  ,型                   ,説明
'         :p_Hinban_sxl()   ,I   ,typ_XSDCA_Update     ,分割結晶(品番)情報
'         :recSXL()         ,O   ,typ_XSDCB_Update     ,分割結晶(SXL)情報
'説明     :分割結晶(品番)をSXL単位に集約する
Private Sub MakeSXLinfo(p_Hinban_sxl() As typ_XSDCA_Update, recSXL() As typ_XSDCB_Update)
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcXSDC_Basic.bas -- Function MakeSXLinfo"
                        
    Dim recCnt As Long
    Dim intLength As Integer
    Dim setSXLflg As Boolean
    Dim i As Long
    Dim j As Long
    Dim n As Long
    
    recCnt = UBound(p_Hinban_sxl)
    ReDim recSXL(recCnt)
    
    setSXLflg = False
    Call SetXSDCB(p_Hinban_sxl(1), recSXL(1))  '1ﾚｺｰﾄﾞ目をセット
    
    n = 1
    For i = 2 To recCnt
        For j = 1 To recCnt
            If p_Hinban_sxl(i).SXLIDCA = recSXL(j).SXLIDCB Then
            
            'ｾｯﾄ済みのﾚｺｰﾄﾞとSXLIDが一緒なら長さを加算する
                intLength = CInt(p_Hinban_sxl(i).GNLCA) + CInt(recSXL(j).LENCB)
                
                If CInt(p_Hinban_sxl(i).INPOSCA) < CInt(recSXL(j).INPOSCB) Then
                '開始位置の小さい方をセット
                    Call SetXSDCB(p_Hinban_sxl(i), recSXL(j))
                    recSXL(j).LENCB = intLength
                Else
                    recSXL(j).LENCB = intLength
                End If
                setSXLflg = True
                Exit For
            End If
        Next
        '一致しないﾚｺｰﾄﾞをセット
        If setSXLflg = False Then
            n = n + 1
            
            Call SetXSDCB(p_Hinban_sxl(i), recSXL(n))
        Else
            setSXLflg = False
        End If
    Next
    
    
proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit

    
End Sub


'概要     :分割結晶(SXL)情報をセットする
'ﾊﾟﾗﾒｰﾀ   :変数名           ,IO  ,型                   ,説明
'         :p_Hinban_sxl     ,I   ,typ_XSDCA_Update     ,分割結晶(品番)情報
'         :recSXL_rtrn      ,O   ,typ_XSDCB_Update     ,分割結晶(SXL)情報
'説明     :分割結晶(品番)ﾃｰﾌﾞﾙ項目に値をｾｯﾄする

Private Sub SetXSDCB(p_Hinban_sxl As typ_XSDCA_Update, recSXL_rtrn As typ_XSDCB_Update)
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcXSDC_Basic.bas -- Function SetXSDCB"
                        
    
    Dim i As Long
    
    
    '分割結晶(SXL)
    With recSXL_rtrn
        .SXLIDCB = p_Hinban_sxl.SXLIDCA          ' SXLID
        '.KCNTCB = p_Hinban_sxl.KCKNTCA           ' 工程連番(ﾌﾞﾛｯｸを受け継いだ場合)
        .KCNTCB = GetKCNTCB(.SXLIDCB)            ' 工程連番(SXLから振る場合)
        .XTALCB = p_Hinban_sxl.XTALCA            ' 結晶番号
        .INPOSCB = p_Hinban_sxl.INPOSCA          ' 結晶内開始位置
        .LENCB = p_Hinban_sxl.GNLCA              ' 長さ
        .HINBCB = p_Hinban_sxl.HINBCA            ' 品番
        .REVNUMCB = p_Hinban_sxl.REVNUMCA        ' 電話番号改訂番号
        .FACTORYCB = p_Hinban_sxl.FACTORYCA      ' 工場
        .OPECB = p_Hinban_sxl.OPECA              ' 操業条件
        .MAICB = p_Hinban_sxl.GNMCA              ' 実枚数
        .MOTHINCB = p_Hinban_sxl.HINBCA          ' 現在品番　　　追加　濱　平成１５年５月１日
'        .WSRMAICB =                             ' WS洗後枚数
'        .WSNMAICB =                             ' WS洗浄欠落枚数
'        .WFCMAICB =                             ' 受入枚数
'        .SXLRMAICB =                            ' SXL指示(良品)
'        .SXLNMAICB =                            ' SXL指示(不良)
'        .WFCNMAICB =                            ' WFC内欠落枚数
'        .SXLEMAICB =                            ' SXL確定枚数
'        .SRMAICB =                              ' サンプル抜指示(良品)
'        .SNMAICB =                              ' サンプル抜指示(不良)
'        .STMAICB =                              ' サンプル枚数
'        .FURIMAICB =                            ' 振替枚数
'        .XTWORKCB =                             ' 製造工場
'        .WFWORKCB =                             ' ウェーハ製造
'        .FURYCCB =                              ' 不良理由
        .LSTCCB = p_Hinban_sxl.LSTATBCA          ' 最終状態区分
'        .LUFRCCB =                              ' 格上コード
'        .LUFRBCB =                              ' 格上区分
'        .LDERCCB =                              ' 格下コード
        .LDFRBCB = p_Hinban_sxl.LDFRBCA          ' 格下区分
'        .HOLDCCB =                              ' ホールドコード
        .HOLDBCB = p_Hinban_sxl.HOLDBCA          ' ホールド区分
'        .EXKUBCB =                              ' 例外区分
'        .HENPKCB =                              ' 返品区分
        .LIVKCB = p_Hinban_sxl.LIVKCA            ' 生死区分
        .KANKCB = p_Hinban_sxl.KANKCA            ' 完了区分
'        .NFCB =                                 ' 入庫区分
'        .SAKJCB =                               ' 削除区分
'        .TDAYCB =                               ' 登録日付
'        .KDAYCB =                               ' 更新日付
'        .SUMITCB =                              ' SUMIT送信フラグ
'        .SNDKCB =                               ' 返品区分
'        .SNDAYCB =                              ' 送信日付
'        .
    
    End With
    
proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit

    
End Sub

'概要     :工程実績情報をセットする
'ﾊﾟﾗﾒｰﾀ   :変数名           ,IO  ,型                   ,説明
'         :p_Koteij         ,O   ,typ_XSDC3_Update     ,工程実績登録情報
'         :p_Hinban         ,I   ,typ_XSDCA_Update     ,分割結晶(品番)情報
'         :hinbanflg        ,I   ,typ_XSDCA_c_flg      ,分割結晶(品番)情報
'説明     :
Private Sub SetXSDC3(p_Koteij As typ_XSDC3_Update, p_Hinban As typ_XSDCA_Update, _
                                  hinbanflg As typ_XSDCA_c_flg)

    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcXSDC_Basic.bas -- Function SetXSDC3"
                        
    
    Dim i As Long
    
    
    '工程実績登録
    With p_Koteij
        .CRYNUMC3 = p_Hinban.CRYNUMCA                                     'ﾌﾞﾛｯｸID
        .INPOSC3 = p_Hinban.INPOSCA                                       '結晶内開始位置
        .KCNTC3 = p_Hinban.KCKNTCA                                        '工程連番
        .HINBC3 = p_Hinban.HINBCA                                         '品番
        .REVNUMC3 = p_Hinban.REVNUMCA                                     '製品番号改訂番号
        .FACTORYC3 = p_Hinban.FACTORYCA                                   '工場
        .OPEC3 = p_Hinban.OPECA                                           '操業条件
        .LENC3 = p_Hinban.GNLCA                                           '長さ
        .XTALC3 = p_Hinban.XTALCA                                         '結晶番号
        .SXLIDC3 = p_Hinban.SXLIDCA                                       'SXLID
        '.WKKTC3 = p_Hinban.GNWKNTCA                                       '工程
        .WKKTC3 = p_Hinban.NEWKNTCA      '工程(最終通過工程をセット)前工程                  '2002/08/29
        '.MACOC3 = p_Hinban.GNMACOCA                                       '処理回数
        .MACOC3 = p_Hinban.NEMACOCA      '処理回数(最終通過処理回数)前工程                  '2002/08/29
        .FRWKKTC3 = msL2Wkkt             '(受入)工程(前工程の最終通過工程)前前工程●           '2002/08/30
        .FRMACOC3 = msL2Maco             '(受入)処理回数(前工程の最終通過処理回数)前前工程●   '2002/08/30
        .TOWKKTC3 = p_Hinban.GNWKNTCA    '(払出)工程(現在工程)次工程                       '2002/08/29
        .TOMACOC3 = p_Hinban.GNMACOCA    '(払出)処理回数(現在処理回数)次工程               '2002/08/29
        If hinbanflg.Furyo <> 0 Then
            .FULC3 = hinbanflg.Furyo                                      '不良長さ
            .FUWC3 = hinbanflg.FuryoW                                     '不良重量
'           .FUWC3 = WeightOfCylinder(dblDiameter, CInt(hinbanflg.Furyo)) '不良重量
            '.FUMC3 =                                                     '不良枚数
        End If
        .TOLC3 = p_Hinban.GNLCA                                           '払出長さ
        .TOWC3 = p_Hinban.GNWCA                                           '払出重量
'       .TOWC3 = WeightOfCylinder(dblDiameter, CInt(p_Hinban.GNLCA))      '払出重量
        '.TOMC3 =                                                         '払出枚数
        '' SUMIT長さ／重量登録追加　04/09/30 ooba
        If PutWtFlg > 0 Then
            .SUMITLC3 = p_Hinban.SUMITLCA                                 'SUMIT長さ
            .SUMITWC3 = p_Hinban.SUMITWCA                                 'SUMIT重量
        End If
        '2003.06.11 (SPK)Y.katabami　代表品番＆新規再切区分情報追加
        .CUTCNTC3 = p_Hinban.CUTCNTCA             '新規再切区分
        .HINBFLGC3 = p_Hinban.HINBFLGCA           '代表品番フラグ
        ''TEST 2005/11
        .RPCRYNUMC3 = p_Hinban.RPCRYNUMCA
    End With
    
proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit

    
End Sub

'概要     :セットパターンⅡの処理を行う(死ロット処理)
'ﾊﾟﾗﾒｰﾀ   :変数名           ,IO  ,型                   ,説明
'         :p_Block_b        ,I   ,typ_XSDC2_Update     ,分割結晶(ﾌﾞﾛｯｸ)前工程実績情報
'         :p_Hinban_b()     ,I   ,typ_XSDCA_Update     ,分割結晶(品番)前工程実績情報
'         :p_Block_c        ,I   ,typ_XSDC2_Update     ,分割結晶(ﾌﾞﾛｯｸ)登録情報
'         :p_Hinban_c()     ,I   ,typ_XSDCA_Update     ,分割結晶(品番)登録情報
'         :p_Furyo()        ,I   ,typ_XSDC4_Update     ,不良内訳登録情報
'         :p_Error          ,O   ,String               ,ｴﾗｰﾒｯｾｰｼﾞ
'         :戻り値           ,O    ,FUNCTION_RETURN      ,新ＤＢへの書込みの成否
'説明     :
Private Function SetPattern2(p_Block_b As typ_XSDC2_Update, p_Hinban_b() As typ_XSDCA_Update, _
                    p_Block_c As typ_XSDC2_Update, p_Hinban_c() As typ_XSDCA_Update, _
                    p_Furyo() As typ_XSDC4_Update, p_Error As String) As FUNCTION_RETURN
                        
                        
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcXSDC_Basic.bas -- Function SetPattern2"
                        
    Dim recCnt As Long    'レコード数
    Dim recCnt2 As Long   'レコード数
    Dim recCnt3 As Long   'レコード数
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim sDbName As String
    Dim strWhere As String
    Dim strErrMsg As String
    Dim SetRec2 As typ_XSDC2_Update
    Dim SetRecA() As typ_XSDCA_Update
    Dim SetRecCB320 As typ_XSDCA_Update 'CB320 G品番ﾚｺｰﾄﾞ作成用
    Dim XSDCA_c_flg() As typ_XSDCA_c_flg
    Dim typKotei() As typ_XSDC3_Update
    Dim HINFLG As typ_XSDCA_c_flg
    Dim strMotHin As String
    Dim AccFlg As Boolean         '不良登録ﾌﾗｸﾞ(不良内訳に登録したかをﾁｪｯｸ)
    Dim typSXL() As typ_XSDCB_Update
    Dim intLen As Integer         'SXL長さ


'対象ﾚｺｰﾄﾞをセット
    If regFLG = "Y" Then
'        recCnt = UBound(p_Hinban_b)
        recCnt = UBound(p_Hinban_c)
'        ReDim p_Hinban_c(recCnt2)
        ReDim typKotei(recCnt)
        ReDim XSDCA_c_flg(recCnt)

        
    'c のﾚｺｰﾄﾞをｾｯﾄ
        SetRec2 = p_Block_c             'ﾌﾞﾛｯｸﾚｺｰﾄﾞ
        
        ReDim SetRecA(recCnt)
        For i = 1 To recCnt
            SetRecA(i) = p_Hinban_c(i)  '品番ﾚｺｰﾄﾞ
        Next
    Else
        recCnt = UBound(p_Hinban_b)
        ReDim typKotei(recCnt)
        ReDim XSDCA_c_flg(0)
        
    'b のﾚｺｰﾄﾞをｾｯﾄ
        SetRec2 = p_Block_b             'ﾌﾞﾛｯｸﾚｺｰﾄﾞ

        ReDim SetRecA(recCnt)
        For i = 1 To recCnt
            SetRecA(i) = p_Hinban_b(i)  '品番ﾚｺｰﾄﾞ
        Next
    End If
    
    '不良数カウント
    recCnt3 = UBound(p_Furyo)



'・nextCode が CB210, CB320, '     ' 以外の場合
    If strNxtCd <> "CB210" And strNxtCd <> "CB320" And strNxtCd <> "     " Then
    
        strNxtCd = "CB210"
        
'        '現在工程変更
'        SetRec2.GNWKNTC2 = strNxtCd
'        For i = 1 To recCnt
'            SetRecA(i).GNWKNTCA = strNxtCd
'        Next
    End If
    

    
'・区分セット
    Select Case strNxtCd
        Case "CB210"
            '分割結晶(ﾌﾞﾛｯｸ)-XSDC2
            With SetRec2
                .LSTATBC2 = "R"   '最終状態区分(ﾘﾒﾙﾄ)
                .LDFRBC2 = "1"    '格下区分(ﾘﾒﾙﾄ)
                .RSTATBC2 = "M"   '流動状態区分(ﾘﾒﾙﾄ受入待ち)
            End With
            '分割結晶(品番)-XSDCA
            For i = 1 To recCnt
                With SetRecA(i)
                    .LSTATBCA = "R"   '最終状態区分(ﾘﾒﾙﾄ)
                    .LDFRBCA = "1"    '格下区分(ﾘﾒﾙﾄ)
                    .RSTATBCA = "M"   '流動状態区分(ﾘﾒﾙﾄ受入待ち)
                End With
            Next
        Case "CB320"
            '分割結晶(ﾌﾞﾛｯｸ)-XSDC2
            SetRec2.RSTATBC2 = "G"   '流動状態区分(ｸﾘｽﾀﾙｶﾀﾛｸﾞ)
            
            '分割結晶(品番)-XSDCA
            For i = 1 To recCnt
                SetRecA(i).RSTATBCA = "G"   '流動状態区分(ｸﾘｽﾀﾙｶﾀﾛｸﾞ)
            Next
        Case "     "
            '分割結晶(ﾌﾞﾛｯｸ)-XSDC2
            With SetRec2
                .LSTATBC2 = "H"   '最終状態区分(廃棄)
                .LDFRBC2 = "2"    '格下区分(ﾊｲｷ)
            End With
            
            '分割結晶(品番)-XSDCA
            For i = 1 To recCnt
                With SetRecA(i)
                    .LSTATBCA = "H"   '最終状態区分(廃棄)
                    .LDFRBCA = "2"    '格下区分(ﾊｲｷ)
                End With
            Next
    End Select
           
    '死ロット処理時も工程を変更するように変更　2002/12/17 tuku  START
    '現在工程変更
    SetRec2.GNWKNTC2 = strNxtCd
    For i = 1 To recCnt
        SetRecA(i).GNWKNTCA = strNxtCd
    Next
    '最終通過工程変更
    SetRec2.NEWKNTC2 = strNowCd
    For i = 1 To recCnt
        SetRecA(i).NEWKNTCA = strNowCd
    Next
    '死ロット処理時も工程を変更するように変更　2002/12/17 tuku END
        
    If CC300Flg = False Then   '現工程CC300はﾌﾞﾛｯｸの登録を行わない
    
'≪分割結晶(ﾌﾞﾛｯｸ)-XSDC2≫
        sDbName = "(XSDC2)"
        With SetRec2
    '        If regFLG = "Y" Then
    '            .GNLC2 = p_Block_c.GNLC2   '登録長さ(c)
    '        Else
    '            .GNLC2 = p_Block_b.GNLC2   '前工程長さ(b)
    '        End If
            
            '生死区分セット
            If strNxtCd = "CB320" Then
                .LIVKC2 = "0"
'2002/10/16 追加-------------------------------------------------------------------▼
                .GNWKNTC2 = strNxtCd         ' 現在工程
                .NEWKNTC2 = strNowCd         ' 最終通過工程
'2002/10/16 追加-------------------------------------------------------------------▲
            Else
                .LIVKC2 = "1"
            End If
        End With
            
        If regFLG = "Y" Then
        ''●セットパターンⅡ-① (登録情報(c))
            If p_Block_b.CRYNUMC2 = p_Block_c.CRYNUMC2 Then   'b=c
            '更新(登録情報(c))
                With SetRec2
                    .KCNTC2 = p_Block_b.KCNTC2 + 1    '工程連番を＋１してセット
                    .GNMACOC2 = p_Block_b.GNMACOC2    '現在処理回数
                    .NEWKNTC2 = p_Block_b.NEWKNTC2    '最終通過工程
                    .NEMACOC2 = p_Block_b.NEMACOC2    '最終通過処理回数
                End With
                strWhere = "WHERE CRYNUMC2 = '" & p_Block_b.CRYNUMC2 & "'"
                If UpdateXSDC2(SetRec2, strWhere) = FUNCTION_RETURN_FAILURE Then
                    SetPattern2 = FUNCTION_RETURN_FAILURE
                    p_Error = GetMsgStr("EAPLY") & sDbName
                    GoTo proc_exit
                End If
            Else
            '登録(登録情報(c))
                SetRec2.KCNTC2 = 1    '工程連番に１をセット
                If CreateXSDC2(SetRec2, strErrMsg) = FUNCTION_RETURN_FAILURE Then
                    SetPattern2 = FUNCTION_RETURN_FAILURE
                    p_Error = strErrMsg
                    GoTo proc_exit
                End If
            End If
        Else
        
        ''●セットパターンⅡ-② (前工程情報(b))
            '更新(前工程情報(b))
            If Left(p_Block_b.CRYNUMC2, 1) <> vbNullChar Then '(ｂがある場合)
                SetRec2.KCNTC2 = p_Block_b.KCNTC2 + 1    '工程連番を＋１してセット
                strWhere = "WHERE CRYNUMC2 = '" & p_Block_b.CRYNUMC2 & "'"
                If UpdateXSDC2(SetRec2, strWhere) = FUNCTION_RETURN_FAILURE Then
                    SetPattern2 = FUNCTION_RETURN_FAILURE
                    p_Error = GetMsgStr("EAPLY") & sDbName
                    GoTo proc_exit
                End If
            End If
        End If
    
    End If
    
    
'≪分割結晶(品番)-XSDCA≫
    sDbName = "(XSDCA)"
    For i = 1 To recCnt
        With SetRecA(i)
'            If regFLG = "Y" Then
'                .GNLCA = p_Hinban_c(i).GNLCA   '登録長さ(c)
'            Else
'                .GNLCA = p_Hinban_b(i).GNLCA   '前工程長さ(b)
'            End If

            '生死区分セット
            .LIVKCA = "1"
        End With
    Next
    
    n = 0

    If regFLG = "Y" Then
    ''●セットパターンⅡ-① (登録情報(c))
    
        recCnt = UBound(p_Hinban_b)
        recCnt2 = UBound(p_Hinban_c)
        
'        '不良数があればセットする
'        For i = 1 To recCnt2
'            For j = 1 To reccnt3
'                If p_Hinban_c(i).CRYNUMCA = p_Furyo(j).XTALC4 _
'                            And p_Hinban_c(i).HINBCA = p_Furyo(j).HINBC4 _
'                            And p_Hinban_c(i).INPOSCA = p_Furyo(j).INPOSC4 Then
'
'                    XSDCA_c_flg(i).Furyo = p_Furyo(j).PUCUTLC4   '不良長さｾｯﾄ
'                    XSDCA_c_flg(i).Index_F = j                   'indexｾｯﾄ
'
'                Else
'                    XSDCA_c_flg(i).Index_F = -1                   'indexｾｯﾄ
'                End If
'            Next
'        Next
    
        For i = 1 To recCnt
            For j = 1 To recCnt2
                If p_Hinban_b(i).CRYNUMCA = p_Hinban_c(j).CRYNUMCA _
                                    And p_Hinban_b(i).HINBCA = p_Hinban_c(j).HINBCA _
                                    And p_Hinban_b(i).INPOSCA = p_Hinban_c(j).INPOSCA Then  'b=c
                    
                    '登録ﾚｺｰﾄﾞ情報(c)でUPDATE
                    With SetRecA(j)
                        .KCKNTCA = SetRec2.KCNTC2      'ﾌﾞﾛｯｸの工程連番セット
                        .GNMACOCA = p_Hinban_b(i).GNMACOCA    '現在処理回数
                        .NEWKNTCA = p_Hinban_b(i).NEWKNTCA    '最終通過工程
                        .NEMACOCA = p_Hinban_b(i).NEMACOCA    '最終通過処理回数
                    End With
                    
                    strWhere = "WHERE CRYNUMCA = '" & p_Hinban_b(i).CRYNUMCA
                    strWhere = strWhere & "' AND HINBCA = '" & p_Hinban_b(i).HINBCA
                    strWhere = strWhere & "' AND INPOSCA = " & p_Hinban_b(i).INPOSCA
                    If UpdateXSDCA(SetRecA(j), strWhere) = FUNCTION_RETURN_FAILURE Then
                        SetPattern2 = FUNCTION_RETURN_FAILURE
                        p_Error = GetMsgStr("EAPLY") & sDbName
                        GoTo proc_exit
                    End If
                    
                    '工程実績-XSDC3登録
                    Call SetXSDC3(typKotei(n), SetRecA(j), XSDCA_c_flg(0))
                    With typKotei(n)
                        .LENC3 = "0"     '長さ
                        .TOLC3 = .LENC3  '払出長さ
                        .TOWC3 = "0"     '払出重量
                        .FULC3 = SetRecA(j).GNLCA      '不良長さ
                        .FUWC3 = SetRecA(j).GNWCA      '重量
                    End With
                    If CreateXSDC3(typKotei(n), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                        SetPattern2 = FUNCTION_RETURN_FAILURE
                        p_Error = strErrMsg
                        GoTo proc_exit
                    End If
                    n = n + 1
                    XSDCA_c_flg(j).Entry = True   '更新ﾌﾗｸﾞ(一致するﾚｺｰﾄﾞアリ)
                    Exit For
                End If
'                XSDCA_c_flg(j).Entry = False   '更新ﾌﾗｸﾞ(一致するﾚｺｰﾄﾞナシ)
            Next
        Next
        
        '一致しないﾚｺｰﾄﾞ(c)の登録
        For i = 1 To recCnt2
             If XSDCA_c_flg(i).Entry = False Then
             
                '工程連番セット
                If recCnt <> 0 And p_Hinban_b(1).CRYNUMCA = p_Hinban_c(1).CRYNUMCA Then
                    SetRecA(i).KCKNTCA = SetRec2.KCNTC2   'ﾌﾞﾛｯｸの工程連番をセット
                Else
                    SetRecA(i).KCKNTCA = 1                '工程連番に１をセット
                End If
             
             
                If CreateXSDCA(SetRecA(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                    SetPattern2 = FUNCTION_RETURN_FAILURE
                    p_Error = strErrMsg
                    GoTo proc_exit
                End If
                
                    
                '≪工程実績-XSDC3登録≫
                Call SetXSDC3(typKotei(n), SetRecA(i), XSDCA_c_flg(0))
                With typKotei(n)
                    .LENC3 = "0"                   '長さ
                    .TOLC3 = .LENC3                '払出長さ
                    .TOWC3 = "0"                   '払出重量
                    .FULC3 = SetRecA(i).GNLCA      '不良長さ
                    .FUWC3 = SetRecA(i).GNWCA      '重量
                End With
                If CreateXSDC3(typKotei(n), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                    SetPattern2 = FUNCTION_RETURN_FAILURE
                    p_Error = strErrMsg
                    GoTo proc_exit
                End If
                n = n + 1
            End If
        Next
        
    Else
    
    ''●セットパターンⅡ-② (前工程情報(b))
    
'        '不良数があればセットする
'        For i = 1 To recCnt
'            For j = 1 To reccnt3
'                If p_Hinban_b(i).CRYNUMCA = p_Furyo(j).XTALC4 _
'                            And p_Hinban_b(i).HINBCA = p_Furyo(j).HINBC4 _
'                            And p_Hinban_b(i).INPOSCA = p_Furyo(j).INPOSC4 Then
'
'                    XSDCA_c_flg(i).Furyo = p_Furyo(j).PUCUTLC4   '不良長さｾｯﾄ
'                    XSDCA_c_flg(i).Index_F = j                   'indexｾｯﾄ
'
'                Else
'                    XSDCA_c_flg(i).Index_F = -1                   'indexｾｯﾄ
'                End If
'            Next
'        Next
    
        For i = 1 To recCnt
            '更新
            SetRecA(i).KCKNTCA = SetRec2.KCNTC2      'ﾌﾞﾛｯｸの工程連番セット
            strWhere = "WHERE CRYNUMCA = '" & p_Hinban_b(i).CRYNUMCA
            strWhere = strWhere & "' AND HINBCA = '" & p_Hinban_b(i).HINBCA
            strWhere = strWhere & "' AND INPOSCA = " & p_Hinban_b(i).INPOSCA
            If UpdateXSDCA(SetRecA(i), strWhere) = FUNCTION_RETURN_FAILURE Then
                SetPattern2 = FUNCTION_RETURN_FAILURE
                p_Error = GetMsgStr("EAPLY") & sDbName
                GoTo proc_exit
            End If
            
            '≪工程実績-XSDC3登録≫
            Call SetXSDC3(typKotei(n), SetRecA(i), XSDCA_c_flg(0))
            With typKotei(n)
                .LENC3 = "0"                  '長さ
                .TOLC3 = .LENC3               '払出長さ
                .TOWC3 = "0"                  '払出重量
                .FULC3 = SetRecA(i).GNLCA     '不良長さ
                .FUWC3 = SetRecA(i).GNWCA     '重量
'2002/10/17----------------------------------------------------
                .WKKTC3 = SetRecA(i).NEWKNTCA                               '工程 　2002/12/17 tuku
                .MACOC3 = SetRecA(i).NEMACOCA                               '処理回数 2002/12/17 tuku
                .FRWKKTC3 = msL2Wkkt                                        '(受入)工程
                .FRMACOC3 = msL2Maco                                        '(受入)処理回数
                .TOWKKTC3 = strNxtCd                                        '(払出)工程(現在工程)次工程
                If Left(SetRec2.CRYNUMC2, 1) <> vbNullChar Then
                     .TOMACOC3 = GetGNMACOC(SetRec2.CRYNUMC2, strNxtCd)     '(払出)処理回数
                Else
                     .TOMACOC3 = 1                                          '(払出)処理回数
                End If
'2002/10/17----------------------------------------------------
                If strNowCd = "CB320" Then
                   .SUMITBC3 = "2 "
                End If
                If strNowCd = "CC700" And strNxtCd = "CB210" Then .PAYCLASSC3 = "3"
                If strNowCd = "CC700" And strNxtCd = "CB320" Then .PAYCLASSC3 = "2"
            End With
            If CreateXSDC3(typKotei(n), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                SetPattern2 = FUNCTION_RETURN_FAILURE
                p_Error = strErrMsg
                GoTo proc_exit
            End If
            n = n + 1
        Next
    End If
    
    
    '次工程がCB320の場合(新規ﾚｺｰﾄﾞ作成)
    If strNxtCd = "CB320" Then
        SetRecCB320 = SetRecA(1)
        With SetRecCB320
            .GNLCA = SetRec2.GNLC2                                 '長さ(ﾌﾞﾛｯｸ長さ)ｾｯﾄ
            .GNWCA = SetRec2.GNWC2
            .INPOSCA = SetRec2.INPOSC2
            .HINBCA = "G"                                          '品番ｾｯﾄ
            .LIVKCA = "0"                                          '生死区分 0
            .KCKNTCA = SetRec2.KCNTC2                              '工程連番(ﾌﾞﾛｯｸ)ｾｯﾄ
            .CHGCA = 0                                             'チャージ量　0
            .KEIDAYCA = ""                                         '計上日付　未セット
'2002/10/16 追加-------------------------------------------------------------------▼
            .GNWKNTCA = strNxtCd         ' 現在工程
            .NEWKNTCA = strNowCd         ' 最終通過工程
'2002/10/16 追加-------------------------------------------------------------------▲
        End With
        
'2002/08/29----------------------------------------------------
'        If CreateXSDCA(SetRecCB320, strErrMsg) = FUNCTION_RETURN_FAILURE Then
'            SetPattern2 = FUNCTION_RETURN_FAILURE
'            p_Error = strErrMsg
'            GoTo proc_exit
'        End If
                
        With SetRecCB320
            If CheckUniqueRecord(.CRYNUMCA, .HINBCA, CInt(.INPOSCA)) = True Then
                '登録
                If CreateXSDCA(SetRecCB320, strErrMsg) = FUNCTION_RETURN_FAILURE Then
                    SetPattern2 = FUNCTION_RETURN_FAILURE
                    p_Error = strErrMsg
                    GoTo proc_exit
                End If
            Else
                '更新
                strWhere = "WHERE CRYNUMCA = '" & .CRYNUMCA
                strWhere = strWhere & "' AND HINBCA = '" & .HINBCA
                strWhere = strWhere & "' AND INPOSCA = " & .INPOSCA
                
                If UpdateXSDCA(SetRecCB320, strWhere) = FUNCTION_RETURN_FAILURE Then
                    SetPattern2 = FUNCTION_RETURN_FAILURE
                    p_Error = GetMsgStr("EAPLY") & sDbName
                    GoTo proc_exit
                End If
            End If
        End With
'2002/08/29----------------------------------------------------
            
            
'        For i = 1 To recCnt
'           If strNowCd <> "CC700" Then
'            With SetRecA(i)
'                strMotHin = .HINBCA & .REVNUMCA & .FACTORYCA & .OPECA  '元品番取得
'                '.HINBCA = "G"                                          '品番ｾｯﾄ
'                .KCKNTCA = SetRec2.KCNTC2                              '工程連番(ﾌﾞﾛｯｸ)ｾｯﾄ
'            End With
'
'            '≪工程実績-XSDC3登録≫
'            Call SetXSDC3(typKotei(n), SetRecA(i), XSDCA_c_flg(0))
'            With typKotei(n)
'                .HINBC3 = "G"                                           '品番ｾｯﾄ
'                .LENC3 = SetRecA(i).GNLCA                               '長さ(品番長さ)ｾｯﾄ
'                .TOLC3 = .LENC3                                         '払出長さ
'                .TOWC3 = WeightOfCylinder(dblDiameter, CInt(.TOLC3))    '払出重量
'                '.FULC3 = SetRecA(i).GNLCA                               '不良長さ
'                '.FUWC3 = SetRecA(i).GNWCA                               '重量
'                .MOTHINC3 = strMotHin                                   '元品番
'2002/10/17----------------------------------------------------
'                .WKKTC3 = SetRecA(i).GNWKNTCA                               '工程
'                .MACOC3 = SetRecA(i).GNMACOCA                               '処理回数
'                .FRWKKTC3 = msL2Wkkt                                        '(受入)工程
'                .FRMACOC3 = msL2Maco                                        '(受入)処理回数
'                .TOWKKTC3 = strNxtCd                                        '(払出)工程(現在工程)次工程
'                .TOMACOC3 = GetGNMACOC(SetRec2.CRYNUMC2, strNxtCd)          '(払出)処理回数
'2002/10/17----------------------------------------------------
'            End With
'            If CreateXSDC3(typKotei(n), strErrMsg) = FUNCTION_RETURN_FAILURE Then
'                SetPattern2 = FUNCTION_RETURN_FAILURE
'                p_Error = strErrMsg
'                GoTo PROC_EXIT
'            End If
'          End If
'        Next
    End If
        
'    '・工程実績-XSDC3
'    For i = 0 To recCnt
'        typXSDC3upd(i).LENC3 = "0"   '長さ
'    Next
    
    '≪不良実績-XSDC4≫
    AccFlg = False
    For i = 1 To recCnt3
        For j = 1 To recCnt
            If p_Furyo(i).XTALC4 = SetRecA(j).CRYNUMCA And _
                                            p_Furyo(i).HINBC4 = SetRecA(j).HINBCA Then
'        With p_Furyo(i)
'            If regFLG = "Y" Then
'                .PUCUTLC4 = p_Hinban_c(i).GNLCA   '登録長さ(c)
'            Else
'                .PUCUTLC4 = p_Hinban_b(i).GNLCA   '前工程長さ(b)
'            End If
'        End With

                With p_Furyo(i)
                    .INPOSC4 = SetRecA(j).INPOSCA
                    .KCKNTC4 = SetRecA(j).KCKNTCA
                    .REVNUMC4 = SetRecA(j).REVNUMCA
                    .FACTORYC4 = SetRecA(j).FACTORYCA
                    .OPEC4 = SetRecA(j).OPECA
                    '.WKKTC4 = SetRecA(j).GNWKNTCA
                    .WKKTC4 = strNowCd
                    .PUCUTLC4 = SetRecA(j).GNLCA
'                    .PUCUTWC4 = WeightOfCylinder(dblDiameter, CInt(SetRecA(j).GNLCA))
                    .PUCUTWC4 = SetRecA(j).GNWCA
                    '.pucutmc4 =
                End With
                
                If p_Furyo(i).PUCUTLC4 <> 0 Then
                    If CreateXSDC4(p_Furyo(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                        SetPattern2 = FUNCTION_RETURN_FAILURE
                        p_Error = strErrMsg
                        GoTo proc_exit
                    End If
                End If
                
                AccFlg = True
                Exit For
                
            End If
        Next
        
        If AccFlg = False Then
            If p_Furyo(i).KCKNTC4 = "" Then
            '一致する品番が無かったﾚｺｰﾄﾞの情報をセット
                With p_Furyo(i)
                    If p_Block_b.CRYNUMC2 = p_Block_c.CRYNUMC2 Then  'b=c (ﾌﾞﾛｯｸ)
                        .KCKNTC4 = p_Block_b.KCNTC2 + 1    '工程連番を＋１してセット
                    Else
                        .KCKNTC4 = 1                       '工程連番に１をセット
                    End If
                    .WKKTC4 = strNowCd                     '工程
                End With
            End If
            
            '登録
            If p_Furyo(i).PUCUTLC4 <> 0 Then
                If CreateXSDC4(p_Furyo(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                    SetPattern2 = FUNCTION_RETURN_FAILURE
                    p_Error = strErrMsg
                    GoTo proc_exit
                End If
            End If
        Else
            AccFlg = True
        End If
    Next
    
    
    '≪分割結晶(SXL)登録-XSDCB≫      ＊分割結晶(品番)をSXL単位に集約
    If SXLflg = True Then
        sDbName = "(XSDCB)"
        Call MakeSXLinfo(SetRecA, typSXL)  '分割結晶(SXL)情報作成
        recCnt = UBound(typSXL)
'        If strNowCd = "CC710" Then     '抜試指示入力の場合
'            For i = 1 To recCnt
'                '登録
'                If typSXL(i).SXLIDCB <> "" Then
'                    If CreateXSDCB(typSXL(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
'                        SetPattern2 = FUNCTION_RETURN_FAILURE
'                        p_Error = strErrMsg
'                        GoTo proc_exit
'                    End If
'                End If
'            Next
'        Else                           'WFｾﾝﾀｰ払出、結晶情報変更の場合
'            For i = 1 To recCnt
'                '更新
'                If typSXL(i).SXLIDCB <> "" Then
'                    strWhere = "WHERE SXLIDCB = '" & typSXL(i).SXLIDCB & "'"
'                    If UpdateXSDCB(typSXL(i), strWhere) = FUNCTION_RETURN_FAILURE Then
'                        SetPattern2 = FUNCTION_RETURN_FAILURE
'                        p_Error = GetMsgStr("EAPLY") & sDBName
'                        GoTo proc_exit
'                    End If
'                End If
'            Next
'        End If

        For i = 1 To recCnt
            If typSXL(i).SXLIDCB <> "" Then

                If CheckSXLrecord(typSXL(i).SXLIDCB, intLen) = 0 Then
                'ﾚｺｰﾄﾞがない場合(登録)
                
                    If CreateXSDCB(typSXL(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                        SetPattern2 = FUNCTION_RETURN_FAILURE
                        p_Error = strErrMsg
                        GoTo proc_exit
                    End If
                    
                Else
                'ﾚｺｰﾄﾞがある場合(更新)
                
                    '抜試指示入力、結晶最終払出入力の場合、長さをプラス
                    'If strNowCd = "CC710" Or strNowCd = "CC700" Then
                    '    typSXL(i).LSTCCB = typSXL(i).LSTCCB + intLen
                    'End If
                    typSXL(i).LSTCCB = typSXL(i).LSTCCB + intLen
                    
                    strWhere = "WHERE SXLIDCB = '" & typSXL(i).SXLIDCB & "'"
                    If UpdateXSDCB(typSXL(i), strWhere) = FUNCTION_RETURN_FAILURE Then
                        SetPattern2 = FUNCTION_RETURN_FAILURE
                        p_Error = GetMsgStr("EAPLY") & sDbName
                        GoTo proc_exit
                    End If
                End If
            End If
        Next

    End If
        
    SetPattern2 = FUNCTION_RETURN_SUCCESS
        

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    SetPattern2 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :シリコン円柱の重量を求める
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :dblDiameter   ,I  ,Double    ,直径(mm)
'          :dblHeight     ,I  ,Double    ,高さ(mm)
'          :戻り値        ,O  ,Double    ,重量(g)
'説明      :
'履歴      :2001/06/29 作成  野村
'          :2002/08/13 Y.Ohno s_cmmc001z よりコピー
Private Function WeightOfCylinder(ByVal dblDiameter As Double, ByVal dblHeight As Double) As Double
Dim dblRadius As Double

    dblRadius = dblDiameter / 2#
    WeightOfCylinder = Int(HIJU_SILICONE * cdblPI * (dblRadius ^ 2) * dblHeight)
End Function






'概要      :ブロック内の最ボトム品番を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :Pi_CRYNUM       ,I  ,String     ,対象のブロックID
'          :tmpXSDCA        ,O  ,typ_XSDCA  ,最ボトム品番
'説明      :引数には、ブロックID・結晶番号をセットします
'履歴      :2002/08/05 M.Tomita
Public Function GetBottomHinban(Pi_CRYNUM As String, tmpXSDCA As typ_XSDCA) As FUNCTION_RETURN


    '変数の定義
    Dim fndXSDCA()  As typ_XSDCA
    Dim strWhere    As String
    Dim strOrder    As String
    Dim sql         As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcXSDC_Basic.bas -- Function GetBottomHinban" '2002/08/05時点ではbas名保留 M.TOMITA
    
    GetBottomHinban = FUNCTION_RETURN_SUCCESS
    
    'WHERE条件式
    strWhere = "WHERE CRYNUMCA = '" & Pi_CRYNUM & "' AND LIVKCA = '0'"
    'OrderBy
    strOrder = "ORDER BY INPOSCA "
    '該当データの取得
''    If DBDRV_GetXSDCA(fndXSDCA(), strWhere) = FUNCTION_RETURN_SUCCESS And UBound(fndXSDCA) > 0 Then '2002/08/22 修正 in FFC初台
    If DBDRV_GetXSDCA(fndXSDCA(), strWhere, strOrder) = FUNCTION_RETURN_SUCCESS And UBound(fndXSDCA) > 0 Then
        '値セット
        tmpXSDCA = fndXSDCA(UBound(fndXSDCA))
        GetBottomHinban = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    GetBottomHinban = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :狙い品番の存在チェックをする
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :Pi_CRYNUM     ,I  ,String      ,結晶番号
'          :戻り値         ,O狙い品番が存在する場合"0"
'                         　 狙い品番が存在しない場合"-1"
'説明      :引数には結晶番号をセットします
'履歴      :2002/08/05 M.Tomita
Public Function Nerai_Hinban_Existence_check(Pi_CRYNUM As String) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset
Dim RET As String
Dim w_Nerai_Hinban As String


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcXSDC_Basic.bas -- Function Nerai_Hinban_Existence_check"

    Nerai_Hinban_Existence_check = FUNCTION_RETURN_SUCCESS

    '***狙い品番を取得
    sql = ""
    sql = sql & "select PUHINBC1 from XSDC1 "
    sql = sql & "where XTALC1 = '" & Pi_CRYNUM & "'"

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP) 'データを抽出する
    
    If rs.RecordCount = 0 Then 'レコードがない場合は正常終了
        rs.Close
        Nerai_Hinban_Existence_check = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    Else
        w_Nerai_Hinban = rs.Fields("PUHINBC1") '*狙い品番
    End If

    '***分割結晶(品番)に上記で求めた"狙い品番"を含むブロックがあるかチェックする。
    sql = ""
    sql = sql & "select CRYNUMCA from XSDCA "
    sql = sql & "where XTALCA = '" & Pi_CRYNUM & "' and " & _
                      "HINBICA = '" & w_Nerai_Hinban & "' "
    sql = sql & "and KANKCA = '0'"
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP) 'データを抽出する
    
    If rs.RecordCount = 0 Then 'レコードがない場合は正常終了
        rs.Close
        Nerai_Hinban_Existence_check = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    
    Nerai_Hinban_Existence_check = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function
proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    Nerai_Hinban_Existence_check = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'概要      :加工区分(分割結晶ブロック)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :tblXSDC2      ,O  ,typ_XSDC2    ,分割結晶（ブロック）
'          :tblXSDC2_b    ,I  ,typ_XSDC2    ,分割結晶（ブロック）前工程
'          :tblXSDCA      ,I  ,typ_XSDCA    ,分割結晶（品番）現工程
'          :戻り値         ,FUNCTION_RETURN  ,成否
'
'説明      :ﾌﾞﾛｯｸIDから、加工区分の取得
'履歴      :2002/08/05 M.Tomita
Public Function Get_ProcDivide(tblXSDC2 As typ_XSDC2, tblXSDC2_b As typ_XSDC2, _
                                    tblXSDCA() As typ_XSDCA) As FUNCTION_RETURN
    Dim sql As String
    Dim tblXSDC1() As typ_XSDC1
    Dim i As Integer
    Dim j As Integer
    
    '戻り値の初期設定
    Get_ProcDivide = FUNCTION_RETURN_FAILURE
    
    '***狙い品番を取得
    sql = "where XTALC1 = '" & tblXSDC2.XTALC2 & "'"

    If DBDRV_GetXSDC1(tblXSDC1(), sql) = FUNCTION_RETURN_FAILURE Then Exit Function
    'レコードがなければプロシージャから抜ける
    If UBound(tblXSDC1) = 0 Then Exit Function
    
    '分割結晶（品番）現工程に狙い品番が存在するかチェック
    For i = 1 To UBound(tblXSDCA)
        '狙い品番があれば、加工区分１をセットし終了
        If tblXSDCA(i).HINBCA = tblXSDC1(1).PUHINBC1 Then
            tblXSDC2.KAKOUBC2 = "1"
            For j = 1 To UBound(tblXSDCA)
                Get_ProcDivide = FUNCTION_RETURN_SUCCESS
                tblXSDCA(j).KAKOUBCA = "1"
            Next j
            Exit Function
        End If
    Next i
    
    '現工程に狙い品番がない場合、前工程との比較
    '前工程の加工区分が1の場合
    If tblXSDC2_b.KAKOUBC2 = "1" Then
        '加工区分2をセット
        tblXSDC2.KAKOUBC2 = "2"
        For j = 1 To UBound(tblXSDCA)
            tblXSDCA(j).KAKOUBCA = "2"
        Next j
    'それ以外
    Else
        '前工程の加工区分をセット
        tblXSDC2.KAKOUBC2 = tblXSDC2_b.KAKOUBC2
        For j = 1 To UBound(tblXSDCA)
            tblXSDCA(j).KAKOUBCA = tblXSDC2_b.KAKOUBC2
        Next j
    End If
    
    Get_ProcDivide = FUNCTION_RETURN_SUCCESS
    
    
End Function
'
''概要      :加工区分(分割結晶品番)
''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
''          :tblXSDCA      ,I  ,typ_XSDCA    ,分割結晶（品番）現工程
''          :戻り値         ,加工区分
''
''説明      :ﾌﾞﾛｯｸIDから、加工区分の取得
''履歴      :2002/08/05 M.Tomita
'Public Function Get_ProcDivide_Hin(tblXSDCA As typ_XSDCA) As Integer
'    Dim sql As String
'    Dim tblXSDC1() As typ_XSDC1
'    Dim tblXSDCA_b() As typ_XSDCA
'    Dim i As Integer
'    Dim exitFlg As Boolean
'
'    '初期値
'    Get_ProcDivide_Hin = 0
'
'    '***狙い品番を取得
'    sql = "where XTALC1 = '" & tblXSDCA.XTALCA & "'"
'
'    If DBDRV_GetXSDC1(tblXSDC1(), sql) = FUNCTION_RETURN_FAILURE Then Exit Function
'    'レコードがなければプロシージャから抜ける
'    If UBound(tblXSDC1) = 0 Then Exit Function
'
'    '分割結晶（品番）現工程に狙い品番が存在するかチェック
'    'あれば、加工区分１をセットしてプロシージャから抜ける
'    If tblXSDCA.HINBCA = tblXSDC1(1).PUHINBC1 Then
'        Get_ProcDivide_Hin = 1
'        Exit Function
'    End If
'
'    '前工程に狙い品番があれば、加工区分に２をセット
'    sql = "where XTALCA = '" & tblXSDCA.XTALCA & "' AND " & _
'                "HINBCA ='" & tblXSDC1(1).PUHINBC1 & "'"
'
'    If DBDRV_GetXSDCA(tblXSDCA_b(), sql) = FUNCTION_RETURN_FAILURE Then Exit Function
'
'    If UBound(tblXSDCA_b) > 0 Then
'        Get_ProcDivide_Hin = 2
'    Else
'        Get_ProcDivide_Hin = 0
'    End If
'End Function
'概要      :直径の算出
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :crynum        ,I  ,String           ,結晶番号OrブロックID
'          :diameter      ,O  ,Double           ,直径
'          :戻り値        ,O  ,FUNCTION_RETURN  ,成否
'
'説明      :結晶番号・ﾌﾞﾛｯｸIDから、直径を算出
'履歴      :2002/08/09 H.FURUYA
Public Function GetDiameter(CRYNUM As String, DIAMETER As Double) As FUNCTION_RETURN
    Dim JudgKakou As Judg_Kakou
    Dim sumDiameter As Double
    Dim rs As OraDynaset
    Dim sql As String
    
    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcXSDC_Basic -- Function GetDiameter"
    GetDiameter = FUNCTION_RETURN_FAILURE
    
    
    '加工実績の取得
'    If scmzc_getKakouJiltuseki(CRYNUM, JudgKakou) = FUNCTION_RETURN_SUCCESS And _
'            JudgKakou.TOP(1) <> "-1" Then
    If scmzc_getKakouJiltuseki(CRYNUM, JudgKakou) = FUNCTION_RETURN_SUCCESS And _
            CInt(JudgKakou.TOP(1)) <> -1 Then
            
        '平均をセットしプロシージャから抜ける
        sumDiameter = JudgKakou.TOP(1) + JudgKakou.TOP(2) + JudgKakou.TAIL(1) + JudgKakou.TAIL(2)
        DIAMETER = sumDiameter / 4#
        '戻り値に成功をセット
        GetDiameter = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
        
    '取得に失敗した場合、H００４のデータを取得する
    sql = "SELECT DM1, DM2, DM3 FROM TBCMH004 " & _
          "WHERE CRYNUM ='" & CRYNUM & "'"

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'レコードがなければ失敗をセットしてプロシージャから抜ける
    If rs.RecordCount = 0 Then
        rs.Close
        GetDiameter = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    '成功したら、平均をセット
    sumDiameter = CDbl(rs("DM1")) + CDbl(rs("DM2")) + CDbl(rs("DM3"))
    DIAMETER = sumDiameter / 3#
    
    '戻り値に成功をセット
    GetDiameter = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    GetDiameter = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :加工実績の取得ドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :BLOCKID        ,I   ,String            ,結晶番号orブロックID
'          :Jiltuseki      ,O   ,Judg_Kakou        ,加工実績
'      　　:戻り値          , O  , FUNCTION_RETURN　, 読み込みの成否
'説明      :
'履歴      :2002/04/17 佐野 信哉 作成
Private Function scmzc_getKakouJiltuseki(BLOCKID As String, Jiltuseki As Judg_Kakou) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String
    Dim recCnt As Integer
    Dim c0 As Integer
    Dim AGRFlag As Boolean
    Dim ans As String
    Dim tINGOTPOS As Integer
    
    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcXSDC_Basic -- Function scmzc_getKakouJiltuseki"
    scmzc_getKakouJiltuseki = FUNCTION_RETURN_FAILURE
    
    '対象ブロックの加工実績の初期化
    For c0 = 1 To 2
        Jiltuseki.TAIL(c0) = -1
        Jiltuseki.TOP(c0) = -1
        Jiltuseki.DPTH(c0) = -1
        Jiltuseki.WIDH(c0) = -1
    Next
    Jiltuseki.pos = ""
'2003/10/18 削除 SystemBrain -------------------------------------------▽
'    If Left(BLOCKID, 1) = "8" Then
'        '購入単結晶の場合
'        sql = "select DMTOP1, DMTOP2, DMTAIL1, DMTAIL2, NCHPOS, NCHDPTH1, NCHDPTH2, NCHWID1, NCHWID2 from TBCMG002 "
'        sql = sql & "where CRYNUM = '" & BLOCKID & "' and "
'        sql = sql & "TRANCNT = any(select max(TRANCNT) from TBCMG002 where CRYNUM = '" & BLOCKID & "')"
'
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        recCnt = rs.RecordCount
'        If recCnt = 0 Then
'            rs.Close
'            scmzc_getKakouJiltuseki = FUNCTION_RETURN_SUCCESS
'            GoTo proc_exit
'        End If
'        Jiltuseki.TAIL(1) = rs("DMTAIL1")
'        Jiltuseki.TAIL(2) = rs("DMTAIL2")
'        Jiltuseki.TOP(1) = rs("DMTOP1")
'        Jiltuseki.TOP(2) = rs("DMTOP2")
'        Jiltuseki.DPTH(1) = rs("NCHDPTH1")
'        Jiltuseki.DPTH(2) = rs("NCHDPTH2")
'        Jiltuseki.WIDH(1) = rs("NCHWID1")
'        Jiltuseki.WIDH(2) = rs("NCHWID2")
'        Jiltuseki.pos = rs("NCHPOS")
'        rs.Close
'    Else
'2003/10/18 削除 SystemBrain -------------------------------------------△
        '引き上げ結晶の場合
        sql = "select DMTOP1, DMTOP2, DMTAIL1, DMTAIL2, NCHPOS, NCHDPTH, NCHWIDTH from TBCMI002 "
        sql = sql & "where CRYNUM='" & Left(BLOCKID, 9) & "000" & "'"
        sql = sql & " and (select INPOSC2 from XSDC2 where CRYNUMC2='" & BLOCKID & "') between INGOTPOS and INGOTPOS+LENGTH-1 "
        sql = sql & "order by INGOTPOS desc, TRANCNT desc"
        sql = "select * from (" & sql & ") where rownum=1"
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        recCnt = rs.RecordCount
        If recCnt = 0 Then
            rs.Close
            scmzc_getKakouJiltuseki = FUNCTION_RETURN_SUCCESS
            GoTo proc_exit
        End If
        Jiltuseki.TAIL(1) = rs("DMTAIL1")
        Jiltuseki.TAIL(2) = rs("DMTAIL2")
        Jiltuseki.TOP(1) = rs("DMTOP1")
        Jiltuseki.TOP(2) = rs("DMTOP2")
        Jiltuseki.DPTH(1) = rs("NCHDPTH")
        Jiltuseki.DPTH(2) = -1
        Jiltuseki.WIDH(1) = rs("NCHWIDTH")
        Jiltuseki.WIDH(2) = -1
        Jiltuseki.pos = rs("NCHPOS")
        rs.Close
'    End If                         '2003/10/18 削除 SystemBrain

    scmzc_getKakouJiltuseki = FUNCTION_RETURN_SUCCESS
proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    scmzc_getKakouJiltuseki = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'概要      :チャージ量の設定の有無をチェック(結晶クラスなし）
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :blkID         ,I  ,String           ,ブロックID
'          :charge        ,O  ,LOng             ,チャージ量
'          :戻り値        ,O  ,FUNCTION_RETURN  ,成否
'
'説明      :
'履歴      :2002/08/15 H.FURUYA
Public Function chkCharge_2(blkID As String, CHARGE As Long, pHinban As String) As FUNCTION_RETURN
    Dim sqlWhere        As String                   'WHERE条件式
    Dim tblXSDC1()      As typ_XSDC1                '結晶引上げテーブル
    Dim tblXSDC1_Up     As typ_XSDC1_Update         '結晶引上げテーブル(更新用）
    Dim tblXSDCA()      As typ_XSDCA                '分割結晶品番テーブル
    Dim i               As Integer
    Dim nowtime         As Date                     'ｻｰﾊﾞｰ日時　05/08/31 ooba
    
   
    chkCharge_2 = FUNCTION_RETURN_FAILURE
    
    'チャージ量の初期化
    CHARGE = -1
    
    '引上げ実績の取得
    'WHERE条件
'2002/09/04
'    sqlWhere = "WHERE XTALC1 = '" & Left(blkID, 8) & "0000" & "' AND KAKOUBC1 ='0'"
    sqlWhere = "WHERE XTALC1 = '" & Left(blkID, 9) & "000" & "' AND KAKOUBC1 ='0'"
    
    'レコードセットの取得(失敗したらプロシージャから抜ける）
    If DBDRV_GetXSDC1(tblXSDC1, sqlWhere) = FUNCTION_RETURN_FAILURE Then Exit Function
    'データがなければ成功をセットしてプロシージャから抜ける
    If UBound(tblXSDC1) = 0 Then
        chkCharge_2 = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If
    
    '前工程の結晶にねらい品番が存在するかチェック
    'WHERE条件
'2002/09/04
    'sqlWhere = "WHERE XTALCA = '" & Left(blkID, 8) & "0000" & "' AND " & _

    sqlWhere = "WHERE XTALCA = '" & Left(blkID, 9) & "000" & "' AND " & _
                     "HINBCA ='" & tblXSDC1(1).PUHINBC1 & "' AND " & _
                     "CRYNUMCA !='" & blkID & " ' AND " & _
                     "LIVKCA = '0'"
                 
    'レコードセットの取得(失敗したらプロシージャから抜ける）
    If DBDRV_GetXSDCA(tblXSDCA, sqlWhere) = FUNCTION_RETURN_FAILURE Then Exit Function
    'データがあれば成功をセットしてプロシージャから抜ける
    If UBound(tblXSDCA) > 0 Then
        chkCharge_2 = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If
                  
    '更新用テーブルにデータをセット
    tblXSDC1_Up.KAKOUBC1 = "1"      '加工区分
'    tblXSDC1_Up.KEIDAYC1 = Now     '計上日付
    nowtime = getSvrTime()          'ｻｰﾊﾞｰ日時取得　05/08/31 ooba
    tblXSDC1_Up.KEIDAYC1 = nowtime  '計上日付　05/08/31 ooba
    
    'WHERE条件
'2002/09/04
    'sqlWhere = "WHERE XTALC1 = '" & Left(blkID, 8) & "0000" & "' AND KAKOUBC1 ='0'"
    sqlWhere = "WHERE XTALC1 = '" & Left(blkID, 9) & "000" & "' AND KAKOUBC1 ='0'"
    
    'データの更新(失敗したらプロシージャから抜ける）
    If UpdateXSDC1(tblXSDC1_Up, sqlWhere) = FUNCTION_RETURN_FAILURE Then Exit Function
    
    '戻り値にチャージ量と成功をセット
    CHARGE = tblXSDC1(1).PUCHAGC1
    pHinban = tblXSDC1(1).PUHINBC1 '2002/11/19 tuku
    chkCharge_2 = FUNCTION_RETURN_SUCCESS
    


End Function





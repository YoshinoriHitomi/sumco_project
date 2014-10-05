Attribute VB_Name = "s_cmzcXDopeCal"
Option Explicit
'結晶データ構造体
Public Type typ_Select_BlockData
    StrCryNum As String                               'ブロック番号
    StrSijiNum As String                              '引上げ番号
    StrXtalNum As String                              '結晶番号
    dblCHARGE As Double                               'チャージ量
    strNERAIN As String                               '狙い[N]
    StrHinban As String                               '品番
    intMNOREVNO As Integer                            '製品番号改訂番号
    StrFactory As String                              '工場
    StrOpeCond As String                              '操業条件
    intTOPOS As Integer                               'TOP位置
    intBOPOS As Integer                               'BOT位置
    strCRYDOPCL As String                             'ドーパント種類
    dblCRYDOPVL As Double                             'ドーパント量
    dblAIMPOS As Double                               'ねらい位置
    dblWGHTTO As Double                               'トップ重量
    intDOPESISU As Integer                            '指数
    dblDopeRyo As Double                              'ドーパント量（指数なし）
    dblDIA As Double                                  '引上げ直径
    intDISPSISU As Integer                            '表示指数
    dblTOPCUT As Double                               'トップカット重量    2009/10/05 Kameda
End Type
'結晶ドープ計算値構造体
Public Type typ_XDOPE_KeisanData
    dblTopLength      As Double   'トップ長さ
    dblInitLiquid     As Double   '初期融液体積
    dblPulRate        As Double   '狙い位置引上率
    dblNeraiInit      As Double   '狙い初期[N]
    dblNeraiDope      As Double   '狙いドープ(Si3N4)量
    dblSi3N4KibanWt   As Double   '基盤重量
    dblSi3N4Weight_10 As Double   'Si3N4重量(1μｍ)
    dblSi3N4Weight_05 As Double   'Si3N4重量(0.5μｍ)
    dblSi3N4Weight_01 As Double   'Si3N4重量(0.1μｍ)
    dblSi3N4Weight_0015 As Double   'Si3N4重量(0.015μｍ)   2011/08/25 Kameda
    dblMaisu_10       As Double   '1μｍドープWF枚数
    intXDopeRyo_10    As Integer  '1μｍドープ量
    intXDopeRyo_05    As Integer  '0.5μｍドープ量
    intXDopeRyo_01    As Integer  '0.1μｍドープ量
    intXDopeRyo_0015  As Integer  '0.015μｍドープ量    2011/08/25 Kameda
    intXDopeRyoJ_10   As Integer  '1μｍドープ量実績
    intXDopeRyoJ_05   As Integer  '0.5μｍドープ量実績
    intXDopeRyoJ_01   As Integer  '0.1μｍドープ量実績
    intXDopeRyoJ_0015 As Integer  '0.015μｍドープ量実績    2011/08/25 Kameda
    dblDopeKei        As Double   '合計ドープWF重量
    dblDopeRyo        As Double   'ドープ(Si3N4)量
    dblSyokiN         As Double   '初期[N]
End Type
'窒素濃度構造体
Public Type typ_NNOUDO_Data
    intXtalPos As Integer         '結晶位置
    dblNnoudo As Double           '窒素濃度
    dblPulWt As Double            '引上げ重量
    dblPuRitu As Double           '引上げ率
End Type
'窒素規格
Public Type typ_spec_N
    HSXCDOPMN As Double
    HSXCDOPMX As Double
    HSXCDPNI As String
    hinban As tFullHinban
    HSXCDOP As String        '2009/09/28 Kameda
End Type

'結晶ドープ計算用定数設定
Public Const CDOPCALC_DIA As Double = 315#             '引き上げ径
Public Const CDOPCALC_CHARGE As Double = 360#          'チャージ量
Public Const CDOPCALC_TOPWEIGHT As Double = 7#         'トップ重量
Public Const CDOPCALC_NERAIPOS As Double = 0#          '狙い位置
Public Const CDOPCALC_NERAIN  As String = "2.00E+13"   '狙いN
Public Const CDOPCALC_TOPCUT As Double = 0#            'トップカット重量  2009/10/05 Kameda

Public Const CDOPCALC_DOPWFDIA As Double = 150#        'ドープwf直径
Public Const CDOPCALC_DOPWFTHICK As Double = 625#      'ドープwf基板厚
Public Const CDOPCALC_FILMTHICK_10 As Double = 1#      '膜厚(1.0μ)
Public Const CDOPCALC_FILMTHICK_05 As Double = 0.5     '膜厚(0.5μ)
Public Const CDOPCALC_FILMTHICK_01 As Double = 0.1     '膜厚(0.1μ)
Public Const CDOPCALC_FILMTHICK_0015 As Double = 0.015   '膜厚(0.015μ)   2011/08/25 Kameda
Public Const CDOPCALC_K0 As Double = 0.0007            'K0
Public Const CDOPCALC_MOL As Double = 140.283          '分子量
Public Const CDOPCALC_DENSITY As Double = 3.185        '密度
Private Const DIA_KUBUN = "300"
Public Const SISU_14 As Integer = 14                   '指数(14乗固定)
Public gBlock As typ_Select_BlockData
Public gKEISAN As typ_XDOPE_KeisanData
'--------------------------------------------------------------------
'概要      :窒素ドープデータ取得
'ﾊﾟﾗﾒｰﾀ    :変数名      ,IO   ,型          ,説明
'          :BLOCK      ,IO  　　,typ_Select_BlockData   ,ブロック詳細
'          :KEISAN     ,O  　　,typ_XDOPE_KeisanData   ,計算結果
'説明      :計算データ、ドープ量を求める
'履歴      :
'///////////////////////////////////////////////////
Public Function GetXLDopeRyo(Block As typ_Select_BlockData, Keisan As typ_XDOPE_KeisanData) As Boolean
    
    
    Dim dModTmp         As Double   '剰余計算
    Dim dRoundTmp       As Double   '四捨五入計算
    Dim i               As Integer
    Dim iSisu           As Integer
    Dim dblDope         As Double
    
    GetXLDopeRyo = False
    
    
    'EXCEL計算式の通りに計算
    ' トップ長さ          = 3*D7*1000/(3.1416*((D5/10/2)^2)*2.328)*10
    ' 初期融液体積        = $D$6*1000/2.57
    ' 狙い位置引上率      = ($D$7+(($D$5/10/2)^2*3.14*2.328*D8/10)/1000)/$D$6
    ' 狙い初期[N]         = D9/(D27*(1-D33)^(D27-1))
    ' 狙いドープ(Si3N4)量 = D34*D32*D28/(4*6.02*10^23)*1000
    ' Si3N4重量(1μｍ)    = (D21/20)^2*3.1416*D23/10000*D29*1000*2
    ' Si3N4重量(0.5μｍ)  = (E21/20)^2*3.1416*E23/10000*E29*1000*2
    ' Si3N4重量(0.1μｍ)  = (F21/20)^2*3.1416*F23/10000*F29*1000*2
    ' 1μｍ               = ROUNDDOWN(D35/D24,0)
    ' 0.5μｍ             = ROUNDDOWN(MOD(D35,D24)/E24,0)
    ' 0.1μｍ             = ROUND(MOD(MOD(D35,D24),E24)/F24,0)
    ' ドープWF基盤重量(D,E,F) = (D21/20)^2*3.1416*D22/10000*2.328
    ' 合計ドープWF重量    = D36*(D25+D24/1000)
    ' ドープ(Si3N4)量     = D24*B15+E24*C15+F24*D15
    ' 初期[N]             = D39*(4*6.02*10^23)/(D32*D28*1000)
    With Keisan
        '指数変換    2011/08/25 Kameda
        iSisu = Block.intDOPESISU
        dblDope = Block.dblDopeRyo
        While (dblDope < 1 Or dblDope >= 10) And dblDope <> 0
            If dblDope < 1 Then    '1未満の場合
                dblDope = dblDope * 10
                iSisu = iSisu - 1       '指数を-1
            ElseIf dblDope >= 10 Then  '10以上の場合
                dblDope = dblDope / 10
                iSisu = iSisu + 1       '指数を+1
            End If
        Wend
        'トップ長さ
        .dblTopLength = 3 * Block.dblWGHTTO * 1000 / (cdblPI * ((Block.dblDIA / 10 / 2) ^ 2) * 2.328) * 10
        
        '初期融液体積
        .dblInitLiquid = Block.dblCHARGE * 1000 / 2.57
        
        '狙い位置引上率
        .dblPulRate = (Block.dblWGHTTO + ((Block.dblDIA / 10 / 2) ^ 2 * 3.14 * 2.328 * Block.dblAIMPOS / 10) / 1000) / Block.dblCHARGE
        
        '狙い初期[N]
        .dblNeraiInit = (Block.dblDopeRyo * 10 ^ Block.intDOPESISU) / (CDOPCALC_K0 * (1 - .dblPulRate) ^ (CDOPCALC_K0 - 1))
        
        '狙いドープ(Si3N4)量
        .dblNeraiDope = .dblNeraiInit * .dblInitLiquid * CDOPCALC_MOL / (4 * 6.02 * 10 ^ 23) * 1000
        
        
        'Si3N4重量
        .dblSi3N4Weight_10 = (CDOPCALC_DOPWFDIA / 20) ^ 2 * cdblPI * CDOPCALC_FILMTHICK_10 / 10000 * CDOPCALC_DENSITY * 1000 * 2
        .dblSi3N4Weight_05 = (CDOPCALC_DOPWFDIA / 20) ^ 2 * cdblPI * CDOPCALC_FILMTHICK_05 / 10000 * CDOPCALC_DENSITY * 1000 * 2
        .dblSi3N4Weight_01 = (CDOPCALC_DOPWFDIA / 20) ^ 2 * cdblPI * CDOPCALC_FILMTHICK_01 / 10000 * CDOPCALC_DENSITY * 1000 * 2
        .dblSi3N4Weight_0015 = (CDOPCALC_DOPWFDIA / 20) ^ 2 * cdblPI * CDOPCALC_FILMTHICK_0015 / 10000 * CDOPCALC_DENSITY * 1000 * 2
        
        '1 μmドープWF枚数
        .dblMaisu_10 = .dblNeraiDope / .dblSi3N4Weight_10
        
        '基盤重量
        .dblSi3N4KibanWt = (CDOPCALC_DOPWFDIA / 20) ^ 2 * cdblPI * CDOPCALC_DOPWFTHICK / 10000 * 2.328
        If iSisu > 11 Then
            '窒素6"WF1.0μ枚数
            .intXDopeRyo_10 = Int(.dblNeraiDope / .dblSi3N4Weight_10)
            
            '窒素6"WF0.5μ枚数
            'VBのModは整数を返すため小数での余りを求める
            dModTmp = .dblNeraiDope - .dblSi3N4Weight_10 * CDbl(.intXDopeRyo_10)
            .intXDopeRyo_05 = Int(dModTmp / .dblSi3N4Weight_05)
            
            '窒素6"WF0.1μ枚数
            dModTmp = dModTmp - .dblSi3N4Weight_05 * CDbl(.intXDopeRyo_05)
            dRoundTmp = dModTmp / .dblSi3N4Weight_01
            .intXDopeRyo_01 = Int(dRoundTmp + 0.5)    '四捨五入
            ' ドープ(Si3N4)量     = D24*B15+E24*C15+F24*D15
            .dblDopeRyo = .dblSi3N4Weight_10 * .intXDopeRyo_10 + .dblSi3N4Weight_05 * .intXDopeRyo_05 + .dblSi3N4Weight_01 * .intXDopeRyo_01
        
        Else
            '窒素6"WF0.015μ枚数  2011/08/25 Kameda
            .intXDopeRyo_0015 = Int(.dblNeraiDope / .dblSi3N4Weight_0015)
            ' ドープ(Si3N4)量     = D24*B15+E24*C15+F24*D15
            .dblDopeRyo = .dblSi3N4Weight_0015 * .intXDopeRyo_0015
        End If
        
        ' 合計ドープWF重量    = D36*(D25+D24/1000)
        .dblDopeKei = .dblMaisu_10 * (.dblSi3N4KibanWt + .dblSi3N4Weight_10 / 1000)
        
        ''投入実績を求める
        'cmhc001d_SelectXDope Keisan, Block.StrCryNum   C6より取得
        
        '実績枚数から初期濃度を計算する    H001より取得　2012/01/27 test Kame
        If SelectWFCount(Block.StrCryNum, Keisan) Then
            .dblDopeRyo = .dblSi3N4Weight_10 * .intXDopeRyoJ_10 + .dblSi3N4Weight_01 * .intXDopeRyoJ_01 + .dblSi3N4Weight_05 * .intXDopeRyo_05 + .dblSi3N4Weight_0015 * .intXDopeRyoJ_0015
        End If
        
        ' 初期[N]             = D39*(4*6.02*10^23)/(D32*D28*1000)
        .dblSyokiN = .dblDopeRyo * (4 * 6.02 * 10 ^ 23) / (.dblInitLiquid * CDOPCALC_MOL * 1000)
    End With

    GetXLDopeRyo = True
    
End Function
'--------------------------------------------------------------------
'概要      :窒素濃度取得(位置指定)
'ﾊﾟﾗﾒｰﾀ    :変数名      ,IO   ,型          ,説明
'          :WGHTTO      ,I  　double       トップ重量
'          :HDIA        ,I  　double       引上げ径
'          :INPOS       ,I  　integer      結晶位置
'          :CHARGE      ,I  　double       チャージ量
'          :NOUDO       ,I  　double     　初期[N]計算値
'          :TOPCUT      ,I  　double     　トップカット重量    2009/10/05 Kameda
'返り値    :窒素濃度
'履歴      :
'///////////////////////////////////////////////////
Public Function GetNNoudo(WGHTTO As Double, HDIA As Double, INPOS As Integer, CHARGE As Double, _
                          NOUDO As Double, TOPCUT As Double) As Double
    Dim dblPulWt As Double
    Dim dblPuRitu As Double
    
        GetNNoudo = 0
        
        '引上げ重量= トップ重量+((引上げ径/10/2)^2*3.14*2.328*結晶位置/10)/1000
        'dblPulWt = WGHTTO + ((HDIA / 10 / 2) ^ 2 * 3.14 * 2.328 * INPOS / 10) / 1000
        dblPulWt = WGHTTO + TOPCUT + ((HDIA / 10 / 2) ^ 2 * 3.14 * 2.328 * INPOS / 10) / 1000   '2009/10/05 Kameda
        
        '引上げ率 = 引上げ重量/チャージ量
        dblPuRitu = Round(dblPulWt / CHARGE, 6)
        
        '濃度 = 初期N*K0*(1-引上げ率)^(K0-1)
        If dblPuRitu < 1 Then
            GetNNoudo = NOUDO * CDOPCALC_K0 * (1 - dblPuRitu) ^ (CDOPCALC_K0 - 1)
        End If
    
        
End Function
'--------------------------------------------------------------------
'概要      :窒素濃度取得  位置(KANKAKU)毎
'ﾊﾟﾗﾒｰﾀ    :変数名      ,IO   ,型          ,説明
'          :BLOCK      ,I  　　,typ_Select_BlockData   ,ブロック詳細
'          :NOUDO      ,I  　　初期[N]計算値
'          :NNOUDO     ,O  　　濃度
'          :KANKAKU    ,I  　　位置間隔
'          :MINLEN     ,I  　　表示開始位置
'          :MAXLEN     ,I  　　表示終了位置
'説明      :計算データ、ドープ量を求める
'履歴      :
'///////////////////////////////////////////////////
Public Function GetNNoudoALL(Block As typ_Select_BlockData, NNOUDO() As typ_NNOUDO_Data, NOUDO As Double, _
                             KANKAKU As Integer) As Boolean
    Dim sPos As Integer
    Dim i As Integer
    Dim sCnt As Integer
    Dim sAmari As Integer
    
    sCnt = Int((Block.intBOPOS - Block.intTOPOS) / KANKAKU)
    sAmari = (Block.intBOPOS - Block.intTOPOS) Mod KANKAKU
    If sAmari <> 0 Then
        sCnt = sCnt + 1
    End If
    
    ReDim NNOUDO(sCnt)
    
    sPos = Block.intTOPOS
    i = 0
    
    For i = 0 To sCnt
        If i = sCnt Then
            sPos = Block.intBOPOS    '最後の行はボトム位置
        End If
        
        '引上げ重量= トップ重量+((引上げ径/10/2)^2*3.14*2.328*結晶位置/10)/1000
        'NNOUDO(i).dblPulWt = Block.dblWGHTTO + ((Block.dblDIA / 10 / 2) ^ 2 * 3.14 * 2.328 * sPos / 10) / 1000
        '引上げ重量= トップ重量+トップカット重量+((引上げ径/10/2)^2*3.14*2.328*結晶位置/10)/10002009/10/05 Kameda
        NNOUDO(i).dblPulWt = Block.dblWGHTTO + Block.dblTOPCUT + ((Block.dblDIA / 10 / 2) ^ 2 * 3.14 * 2.328 * sPos / 10) / 1000
        
        '引上げ率 = 引上げ重量/チャージ量
        NNOUDO(i).dblPuRitu = Round(NNOUDO(i).dblPulWt / Block.dblCHARGE, 6)
        
        '濃度 = 初期N*K0*(1-引上げ率)^(K0-1)
        If NNOUDO(i).dblPuRitu < 1 Then
            NNOUDO(i).dblNnoudo = NOUDO * CDOPCALC_K0 * (1 - NNOUDO(i).dblPuRitu) ^ (CDOPCALC_K0 - 1)
        End If
        '結晶位置
        NNOUDO(i).intXtalPos = sPos
        sPos = sPos + KANKAKU
        
    Next
    
    GetNNoudoALL = True
    
End Function
'--------------------------------------------------------------------
'概要      :窒素濃度取得(切断指示用)
'ﾊﾟﾗﾒｰﾀ    :変数名      ,IO   ,型          ,説明
'          :INPOS       ,I  　integer      結晶位置
'          :CRYNUM      ,I  　string       結晶番号
'          :SISU        ,I  　integer      濃度指数
'返り値    :窒素濃度
'履歴      :
'///////////////////////////////////////////////////
Public Function GetNNoudoSIJI(CRYNUM As String, INPOS() As Long, Sisu As Integer, NOUDO() As Double) As Boolean
    
    Dim dblPulWt As Double
    Dim dblPuRitu As Double
    Dim Block As typ_Select_BlockData
    Dim sNoudo As Double
    Dim iCnt As Integer
    
        GetNNoudoSIJI = False
        
        '初期N濃度取得
        Block.StrCryNum = CRYNUM
        Block.intDOPESISU = Sisu
        
        sNoudo = GetSyokiNoudo(Block)
        
        iCnt = 1
        ReDim NOUDO(UBound(INPOS))
        For iCnt = 1 To UBound(INPOS)
            '2009/10/05 Kameda
            'NOUDO(iCnt) = GetNNoudo(Block.dblWGHTTO, Block.dblDIA, CInt(INPOS(iCnt)), Block.dblCHARGE, sNoudo) / 10 ^ Sisu
            NOUDO(iCnt) = GetNNoudo(Block.dblWGHTTO, Block.dblDIA, CInt(INPOS(iCnt)), Block.dblCHARGE, sNoudo, Block.dblTOPCUT) / 10 ^ Sisu
        Next
        GetNNoudoSIJI = True
End Function
'--------------------------------------------------------------------
'概要      :狙い初期[N]取得
'ﾊﾟﾗﾒｰﾀ    :変数名     ,IO   ,型          ,説明
'          :BLOCK      ,IO  　　,typ_Select_BlockData   ,ブロック詳細
'          :KEISAN     ,O  　　,typ_XDOPE_KeisanData
'説明      :計算データ、ドープ量を求める
'履歴      :
'///////////////////////////////////////////////////
Public Function GetSyokiNoudo(Block As typ_Select_BlockData) As Double
    
    
    Dim Keisan As typ_XDOPE_KeisanData
    
    GetSyokiNoudo = 0
    
    If SelectBlock(Block) = FUNCTION_RETURN_SUCCESS Then
        If Trim(Block.strCRYDOPCL) = "N" Then
            With Keisan
                'Block.dblDopeRyo = Mid(Block.dblCRYDOPVL, 1, InStr(Block.dblCRYDOPVL, ".") - 1)
                Block.dblDopeRyo = Block.dblCRYDOPVL
                If GetXLDopeRyo(Block, Keisan) Then
                    ' 初期[N]
                    'GetSyokiNoudo = .dblDopeRyo * (4 * 6.02 * 10 ^ 23) / (.dblInitLiquid * CDOPCALC_MOL * 1000)
                    GetSyokiNoudo = .dblSyokiN
                End If
            End With
        End If
    End If
    
End Function

'------------------------------------------------
' DBアクセス関数
'------------------------------------------------
'概要      :テーブル「TBCMH001」,「XSDC1」から抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名      ,IO     ,型                     ,説明
'          :rec         ,O  　　,typ_Select_BlockData   ,ブロック詳細
'          :戻り値      ,O      ,FUNCTION_RETURN        ,抽出の成否
'説明      :
'履歴      :

Public Function SelectBlock(rec As typ_Select_BlockData) As FUNCTION_RETURN

Dim sql As String       'SQL全体
Dim rs As OraDynaset    'RecordSet
Dim cnt As Integer      'ｶｳﾝﾄ
Dim i As Long
        
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmec078_SQL.bas -- Function SelectBlock"

    SelectBlock = FUNCTION_RETURN_FAILURE
    
    '***** 引上指示実績（TBCMH001） *****
    ''SQLを組み立てる
    sql = "Select  nvl(CHARGE,0) CHARGE "
    sql = sql & " ,HINBAN,NMNOREVNO,NFACTORY,NOPECOND"
    sql = sql & " ,nvl(CRYDOPCL,' ') CRYDOPCL,nvl(CRYDOPVL,0) CRYDOPVL"
    sql = sql & " ,nvl(DOPN,0) DOPN,nvl(DPNI,' ') DPNI "
    sql = sql & " ,nvl(AIMPOS,0) AIMPOS "
    sql = sql & " From TBCMH001"
    sql = sql & " Where (UPINDNO ='" & left$(rec.StrCryNum, 7) & "0" & Mid(rec.StrCryNum, 9, 1) & "')"
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.EOF Then
        GoTo proc_exit
    End If
     ''抽出結果を格納する
    With rec
        .dblCHARGE = rs("CHARGE") / 1000  ' チャージ量
        .StrHinban = rs("HINBAN")         ' 品番
        .intMNOREVNO = rs("NMNOREVNO")
        .StrFactory = rs("NFACTORY")
        .StrOpeCond = rs("NOPECOND")
        .strCRYDOPCL = rs("CRYDOPCL")     ' 結晶ドープ種類
        '.dblCRYDOPVL = rs("CRYDOPVL")     ' 結晶ドープ量
        .dblCRYDOPVL = rs("DOPN")         ' 結晶ドープ量
        .dblDopeRyo = rs("DOPN")         ' 結晶ドープ量     '2010/01/29 add Kameda
        .dblAIMPOS = rs("AIMPOS")         ' ねらい位置
        .dblDIA = CDOPCALC_DIA            ' 引上げ径
        If Trim(rs("DPNI")) = "" Then
            .intDOPESISU = 0
        Else
            .intDOPESISU = rs("DPNI")         ' 指数
        End If
    End With
    rs.Close
    
    '2009/10/19 add Kameda
    If rec.dblCHARGE = 0 Then
        GoTo proc_exit
    End If
    
    '***** 肩重量（XSDC1） *****
    sql = "select nvl(WGHTTOC1,0) WGHTTOC1,(nvl(DIA1C1,0)+nvl(DIA2C1,0)+nvl(DIA3C1,0))/3 as DIA "
    sql = sql & ", nvl(PUTCUTWC1,0) PUTCUTWC1 "        '2009/10/05 Kameda
    sql = sql & "from XSDC1 "
    sql = sql & "where XTALC1 = '" & left(rec.StrCryNum, 9) & "000" & "' "
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.EOF Then
        rec.dblWGHTTO = GetKTWeight
        rec.dblDIA = CDOPCALC_DIA
        rec.dblTOPCUT = CDOPCALC_TOPCUT          '2009/10/05 Kameda
    Else
        rec.dblWGHTTO = rs("WGHTTOC1") / 1000
        If rs("DIA") = 0 Then
            rec.dblDIA = CDOPCALC_DIA
        Else
            rec.dblDIA = CDbl(rs("DIA"))
        End If
        rec.dblTOPCUT = rs("PUTCUTWC1") / 1000    '2009/10/05 Kameda
    End If
    
    rs.Close
    
    
    SelectBlock = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    SelectBlock = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'概要      :CODEA9から肩重量を得る
'ﾊﾟﾗﾒｰﾀ    :変数名      ,IO ,型       ,説明
'          :KTWeight()  ,O  ,double   ,肩重量
'          :戻り値      ,O  ,FUNCTION_RETURN,抽出の成否
'説明      :
'履歴      :
Public Function GetKTWeight() As Double
Dim rs      As OraDynaset               '抽出RecordDynaset
Dim rsCnt   As Integer                  'ﾚｺｰﾄﾞｶｳﾝﾄ
Dim sql     As String                   'SQL文
Dim i       As Integer                  'ﾙｰﾌﾟｶｳﾝﾄ

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzc010a.bas -- Function GetKTWeight"
    
    'SQL文の作成
    sql = "select CTR02A9 from KODA9 where SYSCA9='K' and SHUCA9='A7' and CODEA9 = '" & DIA_KUBUN & "' "

    'データの抽出
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    '''抽出レコードが存在しない場合
    If rs.EOF Then
        GetKTWeight = 0
        GoTo proc_exit
    End If

    If IsNull(rs("CTR02A9")) Then
        GetKTWeight = 0
    Else
        GetKTWeight = CDbl(rs("CTR02A9"))
    End If
    
    rs.Close


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function
'概要      :結晶ドーパント投入実績取得
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                   ,説明
'          :rec           ,O  ,typ_XDOPE_KeisanData ,抽出レコード
'          :戻り値        ,O  ,FUNCTION_RETURN      ,抽出の成否
'説明      :
'履歴      :
Public Function cmhc001d_SelectXDope(rec As typ_XDOPE_KeisanData, CRYNUM As String) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim rs As OraDynaset    'RecordSet
Dim i As Long
Dim sCryNum As String   '2010/01/29 add Kameda


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmec078_SQL.bas -- Function cmhc001d_SelectXDope"

    cmhc001d_SelectXDope = FUNCTION_RETURN_FAILURE
    
    ''データを抽出     '2010/01/29 Kameda 処理速度対応
    'sql = "Select MATESYUC6,MATERYOC6 " & _
              "From XODC6_1 " & _
              "where (substr(XTALC6,1,9) ='" & left$(CRYNUM, 9) & "')" & _
              " and MATEKC6 = '3'"
    sCryNum = left$(CRYNUM, 9) & "000"
    sCryNum = left$(CRYNUM, 7)
    sql = "Select MATESYUC6,MATERYOC6 " & _
              "From XODC6_1 " & _
              "where substr(XTALC6,1,7) ='" & sCryNum & "'" & _
              " and MATEKC6 = '3'"
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.EOF Then
        GoTo proc_exit
    End If
    Do Until rs.EOF
        Select Case Right(Trim(rs("MATESYUC6")), 3)
            Case "1.0"
                rec.intXDopeRyoJ_10 = rs("MATERYOC6")
            Case "0.5"
                rec.intXDopeRyoJ_05 = rs("MATERYOC6")
            Case "0.1"
                rec.intXDopeRyoJ_01 = rs("MATERYOC6")
            Case "015"                                   '2011/08/25 Kameda
                rec.intXDopeRyoJ_0015 = rs("MATERYOC6")
        End Select
        rs.MoveNext
    Loop
    
    rs.Close
    
    cmhc001d_SelectXDope = FUNCTION_RETURN_SUCCESS

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

'--------------------------------------------------------------------
'概要      :総合判定指示表示用窒素ドープ計算
'ﾊﾟﾗﾒｰﾀ    :変数名     ,IO   ,型          ,説明
'          :BLOCK      ,IO  　　,typ_Select_BlockData   ,ブロック詳細
'          :KEISAN     ,O  　　,typ_XDOPE_KeisanData
'説明      :計算データ、ドープ量を求める
'履歴      :
'///////////////////////////////////////////////////
Public Function GetNNoudoCC600(Block As typ_Select_BlockData, Keisan As typ_XDOPE_KeisanData) As Double
    
    
    'Dim KEISAN As typ_XDOPE_KeisanData    2010/01/29 del Kameda
    'Dim dblSyokiN As Double
    
    GetNNoudoCC600 = 0
    
    With Keisan
        'Block.dblDopeRyo = Mid(Block.dblCRYDOPVL, 1, InStr(Block.dblCRYDOPVL, ".") - 1)
        Block.dblDopeRyo = Block.dblCRYDOPVL
        'If GetXLDopeRyo(Block, KEISAN) Then
            ' 初期[N]
            'dblSyokiN = .dblDopeRyo * (4 * 6.02 * 10 ^ 23) / (.dblInitLiquid * CDOPCALC_MOL * 1000)
        '    dblSyokiN = .dblSyokiN  2010/01/29 del Kameda
        'End If
        '2009/10/05 Kameda
        'GetNNoudoCC600 = GetNNoudo(Block.dblWGHTTO, Block.dblDIA, Block.intTOPOS, Block.dblCHARGE, dblSyokiN) / 10 ^ val(SISU_14)
        GetNNoudoCC600 = GetNNoudo(Block.dblWGHTTO, Block.dblDIA, Block.intTOPOS, Block.dblCHARGE, .dblSyokiN, Block.dblTOPCUT) / 10 ^ val(SISU_14)
    End With
    
End Function

'概要      :窒素規格取得
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型                                  ,説明
'      　　:spec_N  　　　,IO ,
'      　　:戻り値        ,O  ,FUNCTION_RETURN                   　,読み込みの成否
'説明      :
'履歴      :2009/09/03
Public Function GetSpecN(HinSpecN As typ_spec_N) As FUNCTION_RETURN

    Dim rs  As OraDynaset
    Dim sql As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc031_1_SQL.bas -- Function DBDRV_scmzc_fcmkc001d_GetSpecN"
    sql = "select "
    sql = sql & " HSXCDOPMN "
    sql = sql & ",HSXCDOPMX "
    sql = sql & ",HSXCDPNI "
    sql = sql & ",HSXCDOP "
    sql = sql & " from  TBCME020  "  ''
    sql = sql & " where HINBAN  ='" & HinSpecN.hinban.hinban & "'"
    sql = sql & "   and MNOREVNO= " & HinSpecN.hinban.mnorevno
    sql = sql & "   and FACTORY ='" & HinSpecN.hinban.factory & "'"
    sql = sql & "   and OPECOND ='" & HinSpecN.hinban.opecond & "'"
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        GetSpecN = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    With HinSpecN
        If IsNull(rs("HSXCDOPMN")) Then .HSXCDOPMN = 0 Else .HSXCDOPMN = rs("HSXCDOPMN")
        If IsNull(rs("HSXCDOPMX")) Then .HSXCDOPMX = 0 Else .HSXCDOPMX = rs("HSXCDOPMX")
        If IsNull(rs("HSXCDPNI")) Then .HSXCDPNI = "0" Else .HSXCDPNI = rs("HSXCDPNI")
        If IsNull(rs("HSXCDOP")) Then .HSXCDOP = "" Else .HSXCDOP = rs("HSXCDOP")
    End With
    rs.Close

    GetSpecN = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    GetSpecN = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'------------------------------------------------
' DBアクセス関数
'------------------------------------------------
'概要      :テーブル「TBCMH001」から抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名      ,IO     ,型                     ,説明
'          :rec         ,O  　　,StrCryNum              ,ブロックID
'          :戻り値      ,O      ,WF枚数
'説明      :
'履歴      :

Public Function SelectWFCount(StrCryNum As String, rec As typ_XDOPE_KeisanData) As Boolean

Dim sql As String       'SQL全体
Dim rs As OraDynaset    'RecordSet
Dim cnt As Integer      'ｶｳﾝﾄ
Dim i As Long
        
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc016_SQL.bas -- Function SelectWFCount"

    SelectWFCount = False
    
    '***** 引上指示実績（TBCMH001） *****
    ''SQLを組み立てる
    sql = "Select  WFCOUNT10, WFCOUNT05, WFCOUNT01, WFCOUNT0015 "
    sql = sql & " From TBCMH001"
    sql = sql & " Where (UPINDNO ='" & left$(StrCryNum, 7) & "0" & Mid(StrCryNum, 9, 1) & "')"
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.EOF Then
        GoTo proc_exit
    End If
     
     ''抽出結果を格納する
    With rec
        If IsNull(rs("WFCOUNT10")) Then .intXDopeRyoJ_10 = 0 Else .intXDopeRyoJ_10 = rs("WFCOUNT10")
        If IsNull(rs("WFCOUNT05")) Then .intXDopeRyoJ_05 = 0 Else .intXDopeRyoJ_05 = rs("WFCOUNT05")
        If IsNull(rs("WFCOUNT01")) Then .intXDopeRyoJ_01 = 0 Else .intXDopeRyoJ_01 = rs("WFCOUNT01")
        If IsNull(rs("WFCOUNT0015")) Then .intXDopeRyoJ_0015 = 0 Else .intXDopeRyoJ_0015 = rs("WFCOUNT0015")
    End With
    
    rs.Close
    
    SelectWFCount = True
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


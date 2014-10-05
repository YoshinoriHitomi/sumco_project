Attribute VB_Name = "s_cmzccj"
Option Explicit
''Oi判定構造体
Public Type C_Oi
    GuaranteeOi         As Guarantee    ''品質保証情報構造体
    SpecOiMin           As Double       ''品SX酸素濃度下限
    SpecOiMax           As Double       ''品SX酸素濃度上限
    SpecORG             As Double       ''品SX酸素濃度面内分布
    SpecOiAveMin        As Double       ''品SX酸素濃度平均下限
    SpecOiAveMax        As Double       ''品SX酸素濃度平均上限
    Oi()                As Double       ''Oi測定値
    ORG                 As Double       ''ORG計算値
    JudgData            As Double       ''Oi判定対象データ
    JudgOi              As Boolean      ''Oi判定結果
    JudgOrg             As Boolean      ''ORG判定結果
End Type

''Cs判定構造体
Public Type C_Cs
    GuaranteeCs         As Guarantee    ''品質保証情報構造体
    SpecCsMin           As Double       ''品SX炭素濃度下限
    SpecCsMax           As Double       ''品SX炭素濃度上限
    SpecCsKHI           As String * 1   ''品SX検査頻度_位 09/01/08 ooba
    Cs                  As Double       ''Cs測定値
    JudgCs              As Boolean      ''Cs判定結果
End Type

''結晶FTIR判定構造体
Public Type C_FTIR
    GuaranteeOi         As Guarantee    ''品質保証情報構造体
    GuaranteeCs         As Guarantee    ''品質保証情報構造体
    SpecOiMin           As Double       ''品SX酸素濃度下限
    SpecOiMax           As Double       ''品SX酸素濃度上限
    SpecORG             As Double       ''品SX酸素濃度面内分布
    SpecOiAveMin        As Double       ''品SX酸素濃度平均下限
    SpecOiAveMax        As Double       ''品SX酸素濃度平均上限
    SpecCsMin           As Double       ''品SX炭素濃度下限
    SpecCsMax           As Double       ''品SX炭素濃度上限
    Oi(4)               As Double       ''Oi測定値
    Cs                  As Double       ''Cs測定値
    ORG                 As Double       ''ORG計算値
    JudgData            As Double       ''Oi判定対象データ
    JudgOi              As Boolean      ''Oi判定結果
    JudgOrg             As Boolean      ''ORG判定結果
    JudgCs              As Boolean      ''Cs判定結果
End Type

''結晶GFA判定構造体
Public Type C_GFA
    GuaranteeOi         As Guarantee    ''品質保証情報構造体
    SpecOiMin           As Double       ''品SX酸素濃度下限
    SpecOiMax           As Double       ''品SX酸素濃度上限
    SpecORG             As Double       ''品SX酸素濃度面内分布
    SpecOiAveMin        As Double       ''品SX酸素濃度平均下限
    SpecOiAveMax        As Double       ''品SX酸素濃度平均上限
    Ftir(19)            As Double       ''FTIR換算値
    ORG                 As Double       ''ORG計算値
    JudgData            As Double       ''Oi判定対象データ
    JudgFtir            As Boolean      ''FTIR判定結果
    JudgOrg             As Boolean      ''ORG判定結果
End Type

''結晶比抵抗判定構造体
Public Type C_RES
    GuaranteeRes        As Guarantee    ''品質保証情報構造体
    SpecResMin          As Double       ''品SX比抵抗下限
    SpecResMax          As Double       ''品SX比抵抗上限
    SpecResAveMin       As Double       ''品SX比抵抗平均下限
    SpecResAveMax       As Double       ''品SX比抵抗平均上限
    SpecRrg             As Double       ''品SX比抵抗面内分布
    Res(4)              As Double       ''比抵抗測定値
    RRG                 As Double       ''RRG計算値
    JudgData            As Double       ''比抵抗判定対象データ
    JudgRes             As Boolean      ''比抵抗判定値
    JudgRes1            As Boolean      ''比抵抗判定値
    JudgRrg             As Boolean      ''RRG判定値
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    DkTmpSiyo           As String       ''DK温度（仕様）
    DkTmpJsk            As String       ''DK温度（実績）
    JudgDkTmp           As Boolean      ''DK温度判定値
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
End Type

''結晶BMD判定構造体
Public Type C_BMD
    GuaranteeBmd        As Guarantee    ''品質保証情報構造体
    SpecBmdAveMin       As Double       ''品SXBMD平均下限
    SpecBmdAveMax       As Double       ''品SXBMD平均上限
    BMD(4)              As Double       ''BMD測定値
    Min                 As Double       ''最小値
    max                 As Double       ''最大値
    AVE                 As Double       ''平均値
    JudgBmd             As Boolean      ''BMD判定結果
    Bunpu               As Double       ''面内分布
End Type

''結晶OSF判定構造体
Public Type C_OSF
    GuaranteeOsf        As Guarantee    ''品質保証情報構造体
    SpecOsfAveMax       As Double       ''品SXOSF平均上限
    SpecOsfMax          As Double       ''品SX上限
    OSF(19)             As Double       ''OSF測定値
    max                 As Double       ''最大値
    AVE                 As Double       ''平均値
    JudgOsf             As Boolean      ''OSF判定結果
    RD1                 As String * 1   ''RD1
    RD2                 As String * 1   ''RD2
    RD3                 As String * 1   ''RD3
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    ARPTK               As String * 1   '品SXOSF1(ArAN)パタン区分
    ARMIN               As Double       '品SXOSF(ArAN)下限
    ARMAX               As Double       '品SXOSF(ArAN)上限
    ARMHMX              As Double       '品SXOSF(ArAN)面内比上限
    CALCMH              As Double       '面内比(MAX/MIN)
    ArAveMin            As Double       '
    ArAveMax            As Double       '
    JudgOsfPtn          As Boolean      ''OSFパターン判定結果
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
End Type

''結晶GD判定構造体
Public Type C_GD
    GuaranteeDen        As Guarantee    ''品質保証情報構造体
    GuaranteeLdl        As Guarantee    ''品質保証情報構造体
    GuaranteeDvd2       As Guarantee    ''品質保証情報構造体
    JudgFlagDen         As String * 1   ''品SXDen検査有無
    JudgFlagLdl         As String * 1   ''品SXL/DL検査有無
    JudgFlagDvd2        As String * 1   ''品SXDVD2検査有無
    SpecDenMin          As Double       ''品SXDen下限
    SpecDenMax          As Double       ''品SXDen上限
    SpecLdlMin          As Double       ''品SXLdl下限
    SpecLdlMax          As Double       ''品SXLdl上限
    SpecDvd2Min         As Double       ''品SXDvd2下限
    SpecDvd2Max         As Double       ''品SXDvd2上限
'*** UPDATE ↓ Y.SIMIZU 2005/10/13 品WFGDﾗｲﾝ数
    SpecGdLine          As Single       ''品SXGDﾗｲﾝ数
'*** UPDATE ↑ Y.SIMIZU 2005/10/13 品WFGDﾗｲﾝ数
    Den                 As Double       ''Den計算値
    Ldl                 As Double       ''L/DL計算値
    Dvd2                As Double       ''Dvd2計算値
    JudgDen             As Boolean      ''Den判定結果
    JudgLdl             As Boolean      ''L/DL判定結果
    JudgDvd2            As Boolean      ''Dvd2判定結果
    
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
    GDPTK               As String * 1   ''品ＳＸＧＤパタン区分
    LdlMin              As Integer      ''L/DL連続0MIN
    LdlMax              As Integer      ''L/DL連続0MAX
    ZeroLdlMin          As Integer      ''品SXLdl連続0下限
    ZeroLdlMax          As Integer      ''品SXLdl連続0上限
    JudgLdlPtn          As Boolean      ''L/DLパターン判定結果
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
End Type

''結晶ライフタイム判定構造体
Type C_LT
    GuaranteeLt         As Guarantee    ''品質保証情報構造体
    SpecLtMin           As Double       ''品SXLタイム下限
    SpecLtMax           As Double       ''品SXLタイム上限
    SpecLt10Min         As Double       ''品SXLタイム下限(10Ω換算値)
    Lt                  As Double       ''ライフタイム計算値
    Lt10                As Double       ''計算値(10Ω換算値) Add 2011/07/21 T.Koi(SETsw)
    JudgLt              As Boolean      ''ライフタイム判定結果
    JudgLt10            As Boolean      ''判定結果(10Ω換算値) Add 2011/07/21 T.Koi(SETsw)
    resLt10             As String       ''0:待ち 1:OK 2:NG
End Type

''結晶EPD判定構造体
Type C_EPD
    SpecEpdMax          As Double       ''結晶内側管理､EPD上限
    EPD                 As Double       ''EPD測定値
    JudgEpd             As Boolean      ''EPD判定結果
End Type

'2009/08 SUMCO Akizuki
''Ｘ線測定実績 判定構造体
Type C_XY
    Spec_X              As Double        ''測定値　横方向<X>
    SpecX_Max           As Double        ''        横方向<X>上限
    SpecX_Min           As Double        ''        横方向<X>下限
      
    Spec_Y              As Double        ''測定値  縦方向<Y>
    SpecY_Max           As Double        ''        縦方向<Y>上限
    SpecY_Min           As Double        ''        縦方向<Y>下限
    
    Spec_XY             As Double        ''合成角<複合>
    SpecXY_Max          As Double        ''Ｘ線合成角 <複合>上限
    SpecXY_Min          As Double        ''Ｘ線合成角 <複合>下限
    
    JudgResult_X       As Boolean        ''X方向　判定結果
    JudgResult_Y       As Boolean        ''Y方向　判定結果
    JudgResult_XY       As Boolean       ''XY　判定結果
End Type

'Add Start 2011/01/07 SMPK Miyata
''Cu-deco測定実績 判定構造体
Public Type C_CUDECO
    GuaranteeC          As Guarantee    ''品質保証情報構造体
    GuaranteeCJ         As Guarantee    ''品質保証情報構造体
    GuaranteeCJLT       As Guarantee    ''品質保証情報構造体
    GuaranteeCJ2        As Guarantee    ''品質保証情報構造体
    ''----- C判定 ----------------------
    HSXCPK              As String * 1   ''品ＳＸＣパターン区分
    HSXCSZ              As String * 1   ''品ＳＸＣ測定条件

    CPTNJSK             As String * 1   ''C パターン実績
    CDISKJSK            As Integer      ''C Disk半径実績
    CRINGNKJSK          As Integer      ''C Ring内径実績
    CRINGGKJSK          As Integer      ''C Ring外径実績
    ''----- CJ判定 ---------------------
    HSXCJPK             As String * 1   ''品ＳＸＣＪパターン区分
    HSXCJNS             As String * 2   ''品ＳＸＣＪ熱処理法

    CJPTNJSK            As String       ''CJ パターン実績
    CJDISKJSK           As Integer      ''CJ Disk半径実績
    CJRINGNKJSK         As Integer      ''CJ Ring内径実績
    CJRINGGKJSK         As Integer      ''CJ Ring外径実績
    CJBANDNKJSK         As Integer      ''CJ Band内径実績
    CJBANDGKJSK         As Integer      ''CJ Band外径実績
    CJRINGCALC          As Integer      ''CJ Ring幅計算
    CJPICALC            As Integer      ''CJ Pi幅計算
    CJHANTEI            As String       ''CJ 判定結果
    CJBUIUMU            As String       ''CJ 部位別判定有無
    CJDMAXPIC5          As Integer      ''CJ Diskのみパターン Pi幅上限値
    CJRMAXPIC5          As Integer      ''CJ Ringのみパターン Pi幅上限値
    CJDRMAXPIC5         As Integer      ''CJ DiskRingパターン Pi幅上限値
    CJALLMAXDIC5        As Integer      ''CJ 共通Disk半径上限値
    CJALLMINRINC5       As Integer      ''CJ 共通Ring内径下限値
    CJALLMAXRIGC5       As Integer      ''CJ 共通Ring外径上限値
    ''----- CJLT判定 -------------------
    HSXCJLTPK           As String * 1   ''品ＳＸＣＪＬＴパターン区分
    HSXCJLTNS           As String * 2   ''品ＳＸＣＪＬＴ熱処理法
    
    CJLTPTNJSK          As String       ''CJ(LT) パターン実績
    CJLTDISKJSK         As Integer      ''CJ(LT) Disk半径実績
    CJLTRINGNKJSK       As Integer      ''CJ(LT) Ring内径実績
    CJLTRINGGKJSK       As Integer      ''CJ(LT) Ring外径実績
    CJLTBANDNKJSK       As Integer      ''CJ(LT) Band内径実績
    CJLTBANDGKJSK       As Integer      ''CJ(LT) Band外径実績
    CJLTRINGCALC        As Integer      ''CJ(LT) Ring幅計算
    CJLTPICALC          As Integer      ''CJ(LT) Pi幅計算
    CJLTBANDCALC        As Integer      ''CJ(LT) Band幅計算
    HSXCJLTBND          As Integer      ''CJ(LT) Band幅上限値
    ''----- CJ2判定 --------------------
    HSXCJ2PK            As String * 1   ''品ＳＸＣＪ２パターン区分
    HSXCJ2NS            As String * 2   ''品ＳＸＣＪ２熱処理法

    CJ2PTNJSK           As String       ''CJ2 パターン実績
    CJ2DISKJSK          As Integer      ''CJ2 Disk半径実績
    CJ2RINGNKJSK        As Integer      ''CJ2 Ring内径実績
    CJ2RINGGKJSK        As Integer      ''CJ2 Ring外径実績
    CJ2PICALC           As Integer      ''CJ2 Pi幅計算
    CJ2HANTEI           As String       ''CJ2 判定結果
    CJ2BUIUMU           As String       ''CJ2 部位別判定有無
    CJ2DMAXPIC5         As Integer      ''CJ2 Diskのみパターン Pi幅上限値
    CJ2RMAXPIC5         As Integer      ''CJ2 Ringのみパターン Pi幅上限値
    CJ2RMINRINC5        As Integer      ''CJ2 Ringのみパターン Ring内径下限値
    CJ2RMAXRIGC5        As Integer      ''CJ2 Ringのみパターン Ring外径上限値
    CJ2DRMAXPIC5        As Integer      ''CJ2 DiskRingパターン Pi幅上限値
    CJ2DRMINRINC5       As Integer      ''CJ2 DiskRingパターン Ring内径下限値
    CJ2DRMAXRIGC5       As Integer      ''CJ2 DiskRingパターン Ring外径上限値

    JudgC               As Boolean      ''判定結果C
    JudgCJ              As Boolean      ''判定結果CJ
    JudgCJLT            As Boolean      ''判定結果CJLT
    JudgCJ2             As Boolean      ''判定結果CJ2

End Type

''C実績 判定構造体
Public Type C_C
    GuaranteeC          As Guarantee    ''品質保証情報構造体
    HSXCPK              As String * 1   ''品ＳＸＣパターン区分
    HSXCSZ              As String * 1   ''品ＳＸＣ測定条件

    CPTNJSK             As String * 1   ''C パターン実績
    CDISKJSK            As Integer      ''C Disk半径実績
    CRINGNKJSK          As Integer      ''C Ring内径実績
    CRINGGKJSK          As Integer      ''C Ring外径実績

    JudgC               As Boolean      ''判定結果C
End Type

''CJ実績 判定構造体
Public Type C_CJ
    GuaranteeCJ         As Guarantee    ''品質保証情報構造体
    HSXCJPK             As String * 1   ''品ＳＸＣＪパターン区分
    HSXCJNS             As String * 2   ''品ＳＸＣＪ熱処理法

    CJPTNJSK            As String       ''CJ パターン実績
    CJDISKJSK           As Integer      ''CJ Disk半径実績
    CJRINGNKJSK         As Integer      ''CJ Ring内径実績
    CJRINGGKJSK         As Integer      ''CJ Ring外径実績
    CJBANDNKJSK         As Integer      ''CJ Band内径実績
    CJBANDGKJSK         As Integer      ''CJ Band外径実績
    CJRINGCALC          As Integer      ''CJ Ring幅計算
    CJPICALC            As Integer      ''CJ Pi幅計算
    CJHANTEI            As String       ''CJ 判定結果
    CJBUIUMU            As String       ''CJ 部位別判定有無
    CJDMAXPIC5          As Integer      ''CJ Diskのみパターン Pi幅上限値
    CJRMAXPIC5          As Integer      ''CJ Ringのみパターン Pi幅上限値
    CJDRMAXPIC5         As Integer      ''CJ DiskRingパターン Pi幅上限値
    CJALLMAXDIC5        As Integer      ''CJ 共通Disk半径上限値
    CJALLMINRINC5       As Integer      ''CJ 共通Ring内径下限値
    CJALLMAXRIGC5       As Integer      ''CJ 共通Ring外径上限値
    
    JudgCJ              As Boolean      ''判定結果CJ
End Type

''CJLT実績 判定構造体
Public Type C_CJLT
    GuaranteeCJLT       As Guarantee    ''品質保証情報構造体
    HSXCJLTPK           As String * 1   ''品ＳＸＣＪＬＴパターン区分
    HSXCJLTNS           As String * 2   ''品ＳＸＣＪＬＴ熱処理法
    
    CJLTPTNJSK          As String       ''CJ(LT) パターン実績
    CJLTDISKJSK         As Integer      ''CJ(LT) Disk半径実績
    CJLTRINGNKJSK       As Integer      ''CJ(LT) Ring内径実績
    CJLTRINGGKJSK       As Integer      ''CJ(LT) Ring外径実績
    CJLTBANDNKJSK       As Integer      ''CJ(LT) Band内径実績
    CJLTBANDGKJSK       As Integer      ''CJ(LT) Band外径実績
    CJLTRINGCALC        As Integer      ''CJ(LT) Ring幅計算
    CJLTPICALC          As Integer      ''CJ(LT) Pi幅計算
    CJLTBANDCALC        As Integer      ''CJ(LT) Band幅計算
    HSXCJLTBND          As Integer      ''CJ(LT) Band幅上限値

    JudgCJLT            As Boolean      ''判定結果CJLT
End Type

''CJ2実績 判定構造体
Public Type C_CJ2
    GuaranteeCJ2        As Guarantee    ''品質保証情報構造体
    HSXCJ2PK            As String * 1   ''品ＳＸＣＪ２パターン区分
    HSXCJ2NS            As String * 2   ''品ＳＸＣＪ２熱処理法

    CJ2PTNJSK           As String       ''CJ2 パターン実績
    CJ2DISKJSK          As Integer      ''CJ2 Disk半径実績
    CJ2RINGNKJSK        As Integer      ''CJ2 Ring内径実績
    CJ2RINGGKJSK        As Integer      ''CJ2 Ring外径実績
    CJ2PICALC           As Integer      ''CJ2 Pi幅計算
    CJ2HANTEI           As String       ''CJ2 判定結果
    CJ2BUIUMU           As String       ''CJ2 部位別判定有無
    CJ2DMAXPIC5         As Integer      ''CJ2 Diskのみパターン Pi幅上限値
    CJ2RMAXPIC5         As Integer      ''CJ2 Ringのみパターン Pi幅上限値
    CJ2RMINRINC5        As Integer      ''CJ2 Ringのみパターン Ring内径下限値
    CJ2RMAXRIGC5        As Integer      ''CJ2 Ringのみパターン Ring外径上限値
    CJ2DRMAXPIC5        As Integer      ''CJ2 DiskRingパターン Pi幅上限値
    CJ2DRMINRINC5       As Integer      ''CJ2 DiskRingパターン Ring内径下限値
    CJ2DRMAXRIGC5       As Integer      ''CJ2 DiskRingパターン Ring外径上限値

    JudgCJ2             As Boolean      ''判定結果CJ2
End Type
'Add End   2011/01/07 SMPK Miyata

''結晶ブロック偏析判定構造体  2005/1/11追加
Type C_COEF
    NP                   As String       ''
    COEF                 As Double       ''
    JudgCOEF             As Boolean      ''
End Type
''偏析範囲値
Public Const PminusMin As Double = 0.7
Public Const PminusMax As Double = 0.8
Public Const PplusMin As Double = 0.73
Public Const PplusMax As Double = 0.83
Public Const NMin As Double = 0.3
Public Const NMax As Double = 0.4

''SIRD判定構造体   2010/02/04 add Kameda
Type C_SIRD
    SpecSirdMax         As Double       ''仕様面内個数上限
    SIRDCNT             As Double       ''SIRD測定値
    JudgSird            As Boolean      ''SIRD判定結果
End Type


''加工仕様,加工実績構造体
''次元の低い方にMIN値を代入する
Public Type Judg_Kakou
    top(2) As Double
    TAIL(2) As Double
    
    POS As String * 2
    DPTH(2) As Double   '引き上げ結晶の場合データはひとつしか存在しない
    WIDH(2) As Double   '引き上げ結晶の場合データはひとつしか存在しない
    ANGLE(2) As Double     '2009/09 SUMCO Akizuki
End Type

''加工実績判定結果構造体
Public Type Judg_Kakou_Judg
    top As Boolean
    tTOP(2) As Boolean
    TAIL As Boolean
    tTAIL(2) As Boolean
    
    POS As Boolean          'ノッチ位置
    WIDH As Boolean         'ノッチ幅
    DPTH As Boolean         'ノッチ深さ(TOP)
'    DPTH_BOT As Boolean     'ノッチ深さ(BOT)
    ANGLE As Boolean        'ノッチ角度
End Type

''加工実績判定構造体
Public Type type_KakouJudg
    Spec() As Judg_Kakou
    Jiltuseki As Judg_Kakou
    Judg As Judg_Kakou_Judg
'' 09/01/28 FAE)akiyama start
    BLOCKID As String
'' 09/01/28 FAE)akiyama end
End Type

'振替を行った品番　2003/09/05
Public Type fHinban
    moto As String
    saki As String
End Type


'概要      :結晶FTIR判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Ftir          ,I  ,C_FTIR           ,結晶FTIR判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
'　　      :2001/07/19 佐野 信哉 改造
Public Function CrystalFTIRJudg(Ftir As C_FTIR, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    Dim Oi As C_Oi
    Dim Cs As C_Cs
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Oi.GuaranteeOi = Ftir.GuaranteeOi
    Oi.SpecOiMin = Ftir.SpecOiMin
    Oi.SpecOiMax = Ftir.SpecOiMax
    Oi.SpecORG = Ftir.SpecORG
    Oi.SpecOiAveMin = Ftir.SpecOiAveMin
    Oi.SpecOiAveMax = Ftir.SpecOiAveMax
    ReDim Oi.Oi(UBound(Ftir.Oi)) As Double
    For c0 = 0 To UBound(Ftir.Oi)
        Oi.Oi(c0) = Ftir.Oi(c0)
    Next
    Oi.ORG = Ftir.ORG
    
    FuncAns = CrystalOiJudg(Oi, ErrInfo)

    Ftir.JudgData = Oi.JudgData
    Ftir.JudgOi = Oi.JudgOi
    Ftir.JudgOrg = Oi.JudgOrg

'    If Ftir.GuaranteeOi.cJudg = JudgCodeC01 Then ''Oi判定有り
'
'        ''ORG判定
'        Ftir.JudgOrg = RangeDecision_nl(Ftir.ORG, 0, Ftir.SpecORG)
'
'        ''Oi判定
'        If (InStr(ObjCodeGrp01, Ftir.GuaranteeOi.cObj) <> 0) And (GetCrystalJudgData(Ftir.GuaranteeOi, Ftir.Oi(), JData()) = FUNCTION_RETURN_SUCCESS) Then
'            Select Case Ftir.GuaranteeOi.cObj
'            Case ObjCode01, ObjCode02, ObjCode04 ''中心1点、中央値、R/2
'                Ftir.JudgOi = RangeDecision_nl(JData(0), Ftir.SpecOiMin, Ftir.SpecOiMax)
'            Case ObjCode03 ''全域
'                Ftir.JudgOi = JUDG_OK
'                For c0 = 0 To 4
'                    If JData(c0) <> -1 Then
'                        If RangeDecision_nl(JData(c0), Ftir.SpecOiMin, Ftir.SpecOiMax) = JUDG_NG Then
'                            Ftir.JudgOi = JUDG_NG
'                        End If
'                    End If
'                Next
'            End Select
'        Else
'            ''対象データ無し
'            ''エラー情報構造体に情報を代入。
'            FuncAns = SetErrInfo(ErrInfo, EZJ00, OI_JUDG, Ftir.GuaranteeOi.cObj)
'        End If
'        Ftir.JudgOi = (Ftir.JudgOi And Ftir.JudgOrg)
''    Else
''        If InStr(JudgCodeC02, Ftir.GuaranteeOi.cJudg) = 0 Then
''            ''処理方法データ無し
''            ''エラー情報構造体に情報を代入。
''            FuncAns = SetErrInfo(ErrInfo, ZJ001, OI_JUDG, Ftir.GuaranteeOi.cJudg)
''        End If
'    End If
    
    If FuncAns = FUNCTION_RETURN_FAILURE Then Exit Function
    
    Cs.GuaranteeCs = Ftir.GuaranteeCs
    Cs.SpecCsMin = Ftir.SpecCsMin
    Cs.SpecCsMax = Ftir.SpecCsMax
    Cs.Cs = Ftir.Cs
    
    FuncAns = CrystalCsJudg(Cs, ErrInfo)

    Ftir.JudgCs = Cs.JudgCs

'    If Ftir.GuaranteeOi.cJudg = JudgCodeC01 Then ''Cs判定有り
'        Ftir.JudgCs = RangeDecision_nl(Ftir.Cs, Ftir.SpecCsMin, Ftir.SpecCsMax)
''    Else
''        If InStr(JudgCodeC02, Ftir.GuaranteeCs.cJudg) = 0 Then
''            ''処理方法データ無し
''            ''エラー情報構造体に情報を代入。
''            FuncAns = SetErrInfo(ErrInfo, ZJ001, CS_JUDG, Ftir.GuaranteeCs.cJudg)
''        End If
'    End If

    CrystalFTIRJudg = FuncAns
End Function

'概要      :結晶GFA判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Gfa           ,I  ,C_GFA            ,結晶GFA判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function CrystalGFAJudg(Gfa As C_GFA, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    Dim Oi As C_Oi
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Oi.GuaranteeOi = Gfa.GuaranteeOi
    Oi.SpecOiMin = Gfa.SpecOiMin
    Oi.SpecOiMax = Gfa.SpecOiMax
    Oi.SpecORG = Gfa.SpecORG
    Oi.SpecOiAveMin = Gfa.SpecOiAveMin
    Oi.SpecOiAveMax = Gfa.SpecOiAveMax
    ReDim Oi.Oi(UBound(Gfa.Ftir)) As Double
    For c0 = 0 To UBound(Gfa.Ftir)
        Oi.Oi(c0) = Gfa.Ftir(c0)
    Next
    Oi.ORG = Gfa.ORG
    
    FuncAns = CrystalOiJudg(Oi, ErrInfo)
    
    Gfa.JudgData = Oi.JudgData
    Gfa.JudgFtir = Oi.JudgOi
    Gfa.JudgOrg = Oi.JudgOrg
'    If Gfa.GuaranteeOi.cJudg = JudgCodeC01 Then ''GFA判定有り
'
'        ''ORG判定
'        Gfa.JudgOrg = RangeDecision_nl(Gfa.ORG, 0, Gfa.SpecORG)
'
'        ''FTIR判定
'        If (InStr(ObjCodeGrp01, Gfa.GuaranteeOi.cObj) <> 0) And (GetCrystalJudgData(Gfa.GuaranteeOi, Gfa.Ftir(), JData()) = FUNCTION_RETURN_SUCCESS) Then
'            Select Case Gfa.GuaranteeOi.cObj
'            Case ObjCode01, ObjCode02, ObjCode04 ''中心1点、中央値、R/2
'                Gfa.JudgFtir = RangeDecision_nl(JData(0), Gfa.SpecOiMin, Gfa.SpecOiMax)
'            Case ObjCode03 ''全域
'                Gfa.JudgFtir = JUDG_OK
'                For c0 = 0 To 19
'                    If JData(c0) <> -1 Then
'                        If RangeDecision_nl(JData(c0), Gfa.SpecOiMin, Gfa.SpecOiMax) = JUDG_NG Then
'                            Gfa.JudgFtir = JUDG_NG
'                        End If
'                    End If
'                Next
'            End Select
'        Else
'            ''対象データ無し
'            ''エラー情報構造体に情報を代入。
'            FuncAns = SetErrInfo(ErrInfo, EZJ00, GFA_JUDG, Gfa.GuaranteeOi.cObj)
'        End If
'        Gfa.JudgFtir = (Gfa.JudgFtir And Gfa.JudgOrg)
'    Else
''        If InStr(JudgCodeC02, Gfa.GuaranteeOi.cJudg) = 0 Then
''            ''処理方法データ無し
''            ''エラー情報構造体に情報を代入。
''            FuncAns = SetErrInfo(ErrInfo, ZJ001, GFA_JUDG, Gfa.GuaranteeOi.cJudg)
''        End If
'    End If
    
    CrystalGFAJudg = FuncAns
End Function

'概要      :結晶比抵抗判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Res           ,I  ,C_RES            ,結晶比抵抗判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function CrystalRESJudg(Res As C_RES, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim JData(4) As Double
    Dim c0 As Integer
    Dim pt As Integer
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    Dim iRet    As Integer
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Res.JudgData = -1
    Res.JudgRes = JUDG_NG
    Res.JudgRes1 = JUDG_NG
    Res.JudgRrg = JUDG_NG
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    Res.JudgDkTmp = JUDG_NG
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    If Res.GuaranteeRes.cJudg = JudgCodeC01 Then ''RES判定有り


'' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06 start

'        If Trim(Res.GuaranteeRes.cCount) = "" Then
'            pt = 3
'        Else
'            pt = Val(Res.GuaranteeRes.cCount)
'        End If
'        Res.RRG = RoundUp((RGCal(Res.Res(), pt)), 4)
        
        
        ''RRG判定
        Select Case Res.GuaranteeRes.cPos
          Case "B", "C", "D", "E", "F", "K", "S", "Y"
              Select Case Res.GuaranteeRes.cBunp
              Case "A", "B", "C", "M"
                 ''RRG計算
                 Res.RRG = MENNAI_Cal(RES_JUDG, Res.Res(), Res.GuaranteeRes, Res.GuaranteeRes.cBunp)

              Case "", " "
                 ''計算区分がスペースの場合は、計算，判定を行わない
                 Res.RRG = 0
                 Res.JudgRrg = JUDG_OK
                 GoTo Cal_Escp
              Case Else
                 ''RRG計算　　　コード "A" にて計算
                 If Trim(Res.GuaranteeRes.cCount) = "" Then
                    pt = 3
                 Else
                    pt = val(Res.GuaranteeRes.cCount)
                 End If
                 Res.RRG = RoundUp((RGCal(Res.Res(), pt)), 4)

             End Select

          Case Else
             Select Case Res.GuaranteeRes.cBunp
             Case "A", "B", "C", "D", "E", "M", "N"
                 ''RRG計算
                 Res.RRG = MENNAI_Cal(RES_JUDG, Res.Res(), Res.GuaranteeRes, Res.GuaranteeRes.cBunp)

             Case "", " "
                 ''計算区分がスペースの場合は、計算，判定を行わない
                 Res.RRG = 0
                 Res.JudgRrg = JUDG_OK
                 GoTo Cal_Escp
             Case Else
                 ''RRG計算　　　コード "A" にて計算
                 If Trim(Res.GuaranteeRes.cCount) = "" Then
                    pt = 3
                 Else
                    pt = val(Res.GuaranteeRes.cCount)
                 End If
                 Res.RRG = RoundUp((RGCal(Res.Res(), pt)), 4)

             End Select
        End Select

'' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06 end

'2002/02/27 S.Sano RRGの仕様が0の場合は、判定を行わず必ずOKとする。
'2002/02/27 S.Sano 面内分布計算は行う。
        If Res.SpecRrg = 0 Then                                     '2002/02/27 S.Sano
            Res.JudgRrg = JUDG_OK                                   '2002/02/27 S.Sano
        Else                                                        '2002/02/27 S.Sano
            If Res.RRG = -1 Then
                Res.JudgRrg = JUDG_NG
            Else
                Res.JudgRrg = RangeDecision_nl(Res.RRG, 0, Res.SpecRrg)
            End If
        End If                                                      '2002/02/27 S.Sano
        
'' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06 start
Cal_Escp:
'' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06 end
        
        ''RES判定
        '-----TEST2004/10 N追加
        'If (InStr(ObjCodeGrp01, Res.GuaranteeRes.cObj) <> 0) And (GetCrystalJudgData(Res.GuaranteeRes, Res.Res(), JData()) = FUNCTION_RETURN_SUCCESS) Then
        If (InStr(ObjCodeGrp05, Res.GuaranteeRes.cObj) <> 0) And (GetCrystalJudgData(Res.GuaranteeRes, Res.Res(), JData()) = FUNCTION_RETURN_SUCCESS) Then
            Select Case Res.GuaranteeRes.cObj
            Case ObjCode01, ObjCode02, ObjCode04  ''中心1点、中央値、R/2
                Res.JudgRes1 = RangeDecision_nl(JData(0), Res.SpecResMin, Res.SpecResMax)
                Res.JudgData = JData(0)
            'Case ObjCode03 ''全域
            Case ObjCode03, ObjCode13 ''全域、狙い
                Res.JudgRes = JUDG_OK
                Res.JudgRes1 = JUDG_OK
                For c0 = 0 To 4
                    If JData(c0) <> -1 Then
                        If RangeDecision_nl(JData(c0), Res.SpecResMin, Res.SpecResMax) = JUDG_NG Then
                            Res.JudgRes1 = JUDG_NG
                        End If
                    End If
                Next
                Res.JudgData = JudgMax(JData())
            End Select
        Else
            ''対象データ無し
            ''エラー情報構造体に情報を代入。
            FuncAns = SetErrInfo(ErrInfo, EZJ00, RES_JUDG, Res.GuaranteeRes.cObj)
        End If
        
'--------------- 2008/08/25 INSERT START  By Systech ---------------
        ''DK温度判定
        If Trim(Res.DkTmpJsk) = "" Or Trim(Res.DkTmpSiyo) = "" Then
            Res.JudgDkTmp = JUDG_OK
        Else
            iRet = funCodeDBGetMatrixReturn(DKTMP_TBCMB005SYS, DKTMP_TBCMB005CLS, Res.DkTmpJsk, Res.DkTmpSiyo)
            If iRet = -1 Then
                FuncAns = SetErrInfo(ErrInfo, EZJ00, RES_JUDG, Res.GuaranteeRes.cObj)
            ElseIf iRet = 0 Then
                Res.JudgDkTmp = JUDG_NG
            Else
                Res.JudgDkTmp = JUDG_OK
            End If
        End If
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
        
'--------------- 2008/08/25 UPDATE START  By Systech ---------------
'        Res.JudgRes = (Res.JudgRes1 And Res.JudgRrg)
        Res.JudgRes = (Res.JudgRes1 And Res.JudgRrg And Res.JudgDkTmp)
'--------------- 2008/08/25 UPDATE  END   By Systech ---------------
    Else
        Res.JudgRrg = JUDG_OK
        Res.JudgRes = JUDG_OK
        Res.JudgRes1 = JUDG_OK
'--------------- 2008/08/25 INSERT START  By Systech ---------------
        Res.JudgDkTmp = JUDG_OK
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'        If InStr(JudgCodeC02, Res.GuaranteeRes.cJudg) = 0 Then
'            ''処理方法データ無し
'            ''エラー情報構造体に情報を代入。
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, RES_JUDG, Res.GuaranteeRes.cJudg)
'        End If
    End If

    CrystalRESJudg = FuncAns
End Function

'概要      :結晶BMD判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Bmd           ,I  ,C_BMD            ,結晶BMD判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
'          :2001/07/04 佐野 信哉 修正
Public Function CrystalBMDJudg(BMD As C_BMD, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    BMD.JudgBmd = JUDG_NG
    If BMD.GuaranteeBmd.cJudg = JudgCodeC01 Then ''BMD判定有り
        
        ''BMD判定
        If (InStr(ObjCodeGrp02, BMD.GuaranteeBmd.cObj) <> 0) Then
            Select Case BMD.GuaranteeBmd.cObj
            Case ObjCode05 ''全点の平均値
                BMD.JudgBmd = RangeDecision_nl(BMD.AVE, BMD.SpecBmdAveMin, BMD.SpecBmdAveMax)
            Case ObjCode06, ObjCode10, ObjCode11 ''全点の最大値、全点の最小値、MAX(2,4点目)、MAX(2,3,4点目)
                BMD.JudgBmd = RangeDecision_nl(BMD.max, BMD.SpecBmdAveMin, BMD.SpecBmdAveMax)
            Case ObjCode08 ''全点の最小値 ******************************* 購入単結晶判定不可
                BMD.JudgBmd = RangeDecision_nl(BMD.Min, BMD.SpecBmdAveMin, BMD.SpecBmdAveMax)
            Case ObjCode07 ''全点の平均値と最大値 ************************ 購入単結晶判定不可
                If RangeDecision_nl(BMD.AVE, BMD.SpecBmdAveMin, BMD.SpecBmdAveMax) Then
                    BMD.JudgBmd = RangeDecision_nl(BMD.max, BMD.SpecBmdAveMin, BMD.SpecBmdAveMax)
                Else
                    BMD.JudgBmd = JUDG_NG
                End If
            '2001/09/19 S.Sano Start
            Case ObjCode16 ''全点の最小値と最大値 ************************ 購入単結晶判定不可
                If RangeDecision_nl(BMD.Min, BMD.SpecBmdAveMin, BMD.SpecBmdAveMax) Then
                    BMD.JudgBmd = RangeDecision_nl(BMD.max, BMD.SpecBmdAveMin, BMD.SpecBmdAveMax)
                Else
                    BMD.JudgBmd = JUDG_NG
                End If
            '2001/09/19 S.Sano End
            End Select
        Else
            ''対象データ無し
            ''エラー情報構造体に情報を代入。
            FuncAns = SetErrInfo(ErrInfo, EZJ00, BMD_JUDG, BMD.GuaranteeBmd.cObj)
        End If
    Else
        BMD.JudgBmd = JUDG_OK
'        If InStr(JudgCodeC02, Bmd.GuaranteeBmd.cJudg) = 0 Then
'            ''処理方法データ無し
'            ''エラー情報構造体に情報を代入。
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, BMD_JUDG, Bmd.GuaranteeBmd.cJudg)
'        End If
    End If
    
    CrystalBMDJudg = FuncAns
End Function

'概要      :結晶OSF判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Osf           ,I  ,C_OSF            ,結晶OSF判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
'          :2001/07/04 佐野 信哉 修正
Public Function CrystalOSFJudg(OSF As C_OSF, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    Dim Index           As Integer
    Dim dAve            As Double
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    OSF.JudgOsf = JUDG_NG
    

    If Trim(OSF.GuaranteeOsf.cJudg) = JudgCodeC01 Then ''OSF判定有り
        'OSFは仕様値がNullの場合は他の判定と異なりエラーとなる 2004/12/22
        '仕様値のNullチェックでエラーとせず判定でエラーとする
        'Null対応 08/11/06 ooba
''        If OSF.SpecOsfAveMax = -1 Or OSF.SpecOsfMax = -1 Then
''            Exit Function
''        End If
        ''OSF判定
        If (InStr(ObjCodeGrp03, OSF.GuaranteeOsf.cObj) <> 0) Then
            Select Case OSF.GuaranteeOsf.cObj
            Case ObjCode05  ''全点の平均値
                OSF.JudgOsf = RangeDecision_nl(OSF.AVE, 0, OSF.SpecOsfAveMax)
            Case ObjCode06  ''全点の最大値
                OSF.JudgOsf = RangeDecision_nl(OSF.max, 0, OSF.SpecOsfMax)
            Case ObjCode07 ''全点の平均値と最大値
                If RangeDecision_nl(OSF.AVE, 0, OSF.SpecOsfAveMax) Then
                    OSF.JudgOsf = RangeDecision_nl(OSF.max, 0, OSF.SpecOsfMax)
                Else
                    OSF.JudgOsf = JUDG_NG
                End If
            End Select
        'Null対応(規格がNullの場合は対象ｺｰﾄﾞ不問) 08/11/06 ooba
        ElseIf OSF.SpecOsfAveMax = -1 And OSF.SpecOsfMax = -1 Then
            OSF.JudgOsf = JUDG_OK
        Else
            ''対象データ無し
            ''エラー情報構造体に情報を代入。
            FuncAns = SetErrInfo(ErrInfo, EZJ00, OSF_JUDG, OSF.GuaranteeOsf.cObj)
        End If

    Else
        OSF.JudgOsf = JUDG_OK
'        If InStr(JudgCodeC02, Osf.GuaranteeOsf.cJudg) = 0 Then
'            ''処理方法データ無し
'            ''エラー情報構造体に情報を代入。
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, OSF_JUDG, Osf.GuaranteeOsf.cJudg)
'        End If
    End If
            
    CrystalOSFJudg = FuncAns
    
End Function

'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
'概要      :結晶OSF判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Osf           ,I  ,C_OSF            ,結晶OSF判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2008/10/01 Systech 作成  L/DL,OSF判定ﾛｼﾞｯｸ追加
Public Function CrystalOSFJudg_02(OSF As C_OSF, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    Dim Index           As Integer
    Dim Index2          As Integer
    Dim dAve            As Double
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    OSF.JudgOsfPtn = JUDG_OK
    
    If Trim(OSF.GuaranteeOsf.cJudg) = JudgCodeC01 Then ''OSF判定有り
        'OSFは仕様値がNullの場合は他の判定と異なりエラーとなる 2004/12/22
        '仕様値のNullチェックでエラーとせず判定でエラーとする
        
        If OSF.ARPTK = "1" Then
        'FG7
'            OSF.JudgOsf = JUDG_OK

'            If OSF.ARMIN = -1 Or _
'               OSF.ARMAX = -1 Then
'            Else
                '上下限判定
                For Index = 0 To 19
                    If OSF.OSF(Index) >= 0 Then
                        OSF.JudgOsfPtn = RangeDecision_nl(OSF.OSF(Index), OSF.ARMIN, OSF.ARMAX) '08/11/06 ooba
                        If OSF.JudgOsfPtn = JUDG_NG Then Exit For   '08/11/06 ooba
'                        If OSF.OSF(Index) >= OSF.ARMIN And _
'                           OSF.OSF(Index) <= OSF.ARMAX Then
'                        Else
'                            OSF.JudgOsfPtn = JUDG_NG
'                            Exit For
'                        End If
                    End If
                Next Index
'            End If
            
            If OSF.JudgOsfPtn = True Then
                If OSF.ARMHMX = -1 Then
                Else
                    '面内比(MAX/MIN)判定
''                    dAve = OSF.ArAveMax / OSF.ArAveMin
''                    dAve = (Fix((dAve * 10) + 0.9) / 10)    '小数第2位切り上げ
                    
                    If OSF.CALCMH <= OSF.ARMHMX Then
                    Else
                        OSF.JudgOsfPtn = JUDG_NG
                    End If
                End If
            End If
            
        ElseIf OSF.ARPTK = "2" Then
        '水冷(ArAN)
'            OSF.JudgOsf = JUDG_OK
    
            If OSF.ARMHMX = -1 Then
            Else
                '面内比(MAX/MIN)判定
                If OSF.CALCMH <= OSF.ARMHMX Then
                Else
                    OSF.JudgOsfPtn = JUDG_NG
                End If
            End If
            
            If OSF.JudgOsfPtn = True Then
                If OSF.ARMAX = -1 Then
                Else
                    '上限判定
                    If OSF.CALCMH = -1 Then     '08/11/06 ooba
                    'If OSF.OSF(Index) = 0 Then
                        For Index2 = 0 To 19
                            If OSF.ARMAX >= OSF.OSF(Index2) Then
                            Else
                                OSF.JudgOsfPtn = JUDG_NG
                                Exit For
                            End If
                        Next Index2
                    End If
                End If
            End If
        Else
        '判定なし
'            OSF.JudgOsf = JUDG_OK
        
        End If
    Else
'        OSF.JudgOsf = JUDG_OK
    End If
            
    CrystalOSFJudg_02 = FuncAns
    
End Function
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End

'概要      :結晶GD判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Gd            ,I  ,C_GD             ,結晶GD判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function CrystalGDJudg(GD As C_GD, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    ''Den検査有無判断
    GD.JudgDen = JUDG_OK
    If GD.JudgFlagDen = "1" Then
        If GD.GuaranteeDen.cJudg = JudgCodeC01 Then ''Den判定あり
            GD.JudgDen = RangeDecision_nl(GD.Den, GD.SpecDenMin, GD.SpecDenMax)
        Else
            GD.JudgDen = JUDG_OK
'            If InStr(JudgCodeC02, Gd.GuaranteeDen.cJudg) = 0 Then
'                ''処理方法データ無し
'                ''エラー情報構造体に情報を代入。
'                FuncAns = SetErrInfo(ErrInfo, ZJ001, DEN_JUDG, Gd.GuaranteeDen.cJudg)
'            End If
        End If
    End If
    
    ''L/DL検査有無判断
    GD.JudgLdl = JUDG_OK
    If GD.JudgFlagLdl = "1" Then
        If GD.GuaranteeLdl.cJudg = JudgCodeC01 Then ''L/DL判定あり
            GD.JudgLdl = RangeDecision_nl(GD.Ldl, GD.SpecLdlMin, GD.SpecLdlMax)
        Else
            GD.JudgLdl = JUDG_OK
'            If InStr(JudgCodeC02, Gd.GuaranteeLdl.cJudg) = 0 Then
'                ''処理方法データ無し
'                ''エラー情報構造体に情報を代入。
'                FuncAns = SetErrInfo(ErrInfo, ZJ001, LDL_JUDG, Gd.GuaranteeLdl.cJudg)
'            End If
        End If
    End If
    
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
'    If GD.JudgLdl = JUDG_OK Then
    GD.JudgLdlPtn = JUDG_OK
    If GD.JudgFlagLdl = "1" Then
        If GD.GuaranteeLdl.cJudg = JudgCodeC01 Then ''L/DL判定あり
            If GD.GDPTK = "1" Then
                ' "0"連続数(MIN)　≧　品__L/DL連続0下限(SX/WF)
                If GD.ZeroLdlMin = -1 Then
                    GD.JudgLdlPtn = JUDG_OK
                Else
                    If GD.LdlMin >= GD.ZeroLdlMin Then
                        GD.JudgLdlPtn = JUDG_OK
                    Else
                        GD.JudgLdlPtn = JUDG_NG
                    End If
                End If
            ElseIf GD.GDPTK = "2" Then
                ' "0"連続数(MAX)　≦　品__L/DL連続0上限(SX/WF)
                If GD.ZeroLdlMax = -1 Then
                    GD.JudgLdlPtn = JUDG_OK
                Else
                    If GD.LdlMax <= GD.ZeroLdlMax Then
                        GD.JudgLdlPtn = JUDG_OK
                    Else
                        GD.JudgLdlPtn = JUDG_NG
                    End If
                End If
            Else
                ' 判定無し
                
            End If
        End If
    End If
'    End If
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    
    ''DVD2検査有無判断
    GD.JudgDvd2 = JUDG_OK
    If GD.JudgFlagDvd2 = "1" Then
        If GD.GuaranteeDvd2.cJudg = JudgCodeC01 Then ''Dvd2判定あり
'項目追加，修正対応 2003.05.20 yakimura
'            GD.JudgDvd2 = RangeDecision_nl(GD.Dvd2, GD.SpecDvd2Min * 10!, GD.SpecDvd2Max * 10!)
            GD.JudgDvd2 = RangeDecision_nl(GD.Dvd2, GD.SpecDvd2Min, GD.SpecDvd2Max)
'項目追加，修正対応 2003.05.20 yakimura
        Else
            GD.JudgDvd2 = JUDG_OK
'            If InStr(JudgCodeC02, Gd.GuaranteeDvd2.cJudg) = 0 Then
'                ''処理方法データ無し
'                ''エラー情報構造体に情報を代入。
'                FuncAns = SetErrInfo(ErrInfo, ZJ001, DVD2_JUDG, Gd.GuaranteeDvd2.cJudg)
'            End If
        End If
    End If
    
    CrystalGDJudg = FuncAns
End Function

'概要      :結晶ライフタイム判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Lt            ,I  ,C_LT             ,結晶ライフタイム判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function CrystalLTJudg(Lt As C_LT, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Lt.JudgLt = JUDG_OK
    If Lt.GuaranteeLt.cJudg = JudgCodeC01 Then
        If Lt.Lt < Lt.SpecLtMin Then
            Lt.JudgLt = JUDG_NG
        End If
    Else
'        If InStr(JudgCodeC02, Lt.GuaranteeLt.cJudg) = 0 Then
'            ''処理方法データ無し
'            ''エラー情報構造体に情報を代入。
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, LT_JUDG, Lt.GuaranteeLt.cJudg)
'        End If
    End If
    
    CrystalLTJudg = FuncAns
End Function

'概要      :結晶判定を行う。（１０Ω換算値）
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Lt            ,I  ,C_LT             ,結晶判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2011/07/21 T.Koi(SETsw)
Public Function CrystalLT10Judg(Lt As C_LT, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Lt.JudgLt10 = JUDG_OK
    If Lt.GuaranteeLt.cJudg = JudgCodeC01 Then
        If Lt.Lt10 < Lt.SpecLt10Min Then
            Lt.JudgLt10 = JUDG_NG
        End If
    End If
    
    CrystalLT10Judg = FuncAns
End Function


'概要      :結晶EPD判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Epd           ,I  ,C_EPD            ,結晶EPD判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2001/06/06 佐野 信哉 作成
Public Function CrystalEPDJudg(EPD As C_EPD, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    ''エラー情報構造体初期化
    CrystalEPDJudg = SetErrInfo(ErrInfo)
    
    EPD.JudgEpd = RangeDecision_nl(EPD.EPD, 0, EPD.SpecEpdMax)
    CrystalEPDJudg = FUNCTION_RETURN_SUCCESS
End Function

'概要      :X線 合成角<複合>の判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :XY            ,I  ,C_XY             ,X線 判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2009/08 SUMCO 秋月(EPD測定に倣って、作成)

Public Function CrystalXYJudg(XY As C_XY, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    
    ''エラー情報構造体初期化
    CrystalXYJudg = SetErrInfo(ErrInfo)
    
    ''判定の実施
    
    XY.JudgResult_X = RangeDecision_nl(XY.Spec_X, XY.SpecX_Min, XY.SpecX_Max)
    XY.JudgResult_Y = RangeDecision_nl(XY.Spec_Y, XY.SpecY_Min, XY.SpecY_Max)
    XY.JudgResult_XY = RangeDecision_nl(XY.Spec_XY, XY.SpecXY_Min, XY.SpecXY_Max)
    
    CrystalXYJudg = FUNCTION_RETURN_SUCCESS
    
End Function
'概要      :SIRD判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :SIRD          ,I  ,C_SIRD           ,SIRD判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2010/02/04 Kameda
Public Function CrystalSIRDJudg(SIRD As C_SIRD, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    ''エラー情報構造体初期化
    CrystalSIRDJudg = SetErrInfo(ErrInfo)
    
    SIRD.JudgSird = RangeDecision_nl(SIRD.SIRDCNT, 0, SIRD.SpecSirdMax)
    CrystalSIRDJudg = FUNCTION_RETURN_SUCCESS
End Function

'Add Start 2011/01/24 SMPK Miyata
'概要      :結晶C判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Cudeco        ,I  ,C_CUDECO         ,結晶Cu-deco判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :
Public Function CrystalCJudg(CuDeco As C_C, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns         As FUNCTION_RETURN
    Dim ptnOK(3)        As String
    Dim bJudg           As Boolean


    '' パターン判定情報
    '' １文字目はパターン区分、２文字以降がＯＫとするバターン実績を設定する
    ''
    ''　    バターン実績の種類
    ''          "1" : リング無し指定    "2" : ディスク無し指定      "3" : パターン無し指定
    ''          "4" : 不問 (選択なし)   "5" : バンド無し指定        "6" : Pバンド無し指定
    ''          "7" : Bバンド指定無し
    ''
    ''　    バターン実績の種類
    ''          "0":None    "1":Ring    "2":Disk    "3":Disk & Ring
    ''          "5":PB-band "6":P-band  "7":B-band
    ptnOK(0) = "1 02"        '' リング無し指定
    ptnOK(1) = "2 01"        '' ディスク無し指定
    ptnOK(2) = "3 0"         '' パターン無し指定
    ptnOK(3) = "4 0123"      '' 不問 (選択なし)
    
    ''エラー情報構造体初期化
    CrystalCJudg = SetErrInfo(ErrInfo)
    
    '*** パターン判定 ***
    bJudg = CudecoJudgPattern(CuDeco.HSXCPK, CuDeco.CPTNJSK, ptnOK)
    
    CuDeco.JudgC = bJudg

    CrystalCJudg = FUNCTION_RETURN_SUCCESS

End Function

'概要      :結晶CJ判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Cudeco        ,I  ,C_CUDECO         ,結晶Cu-deco判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :
Public Function CrystalCJJudg(CuDeco As C_CJ, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns         As FUNCTION_RETURN
    Dim ptnOK(3)        As String
    Dim bJudg           As Boolean
    Dim intMax          As Integer

    '' パターン判定情報
    '' １文字目はパターン区分、２文字以降がＯＫとするバターン実績を設定する
    ''
    ''　    バターン実績の種類
    ''          "1" : リング無し指定    "2" : ディスク無し指定      "3" : パターン無し指定
    ''          "4" : 不問 (選択なし)   "5" : バンド無し指定        "6" : Pバンド無し指定
    ''          "7" : Bバンド指定無し
    ''
    ''　    バターン実績の種類
    ''          "0":None    "1":Ring    "2":Disk    "3":Disk & Ring
    ''          "5":PB-band "6":P-band  "7":B-band
    ptnOK(0) = "1 02"        '' リング無し指定
    ptnOK(1) = "2 01"        '' ディスク無し指定
    ptnOK(2) = "3 0"         '' パターン無し指定
    ptnOK(3) = "4 0123"      '' 不問 (選択なし)
    
    ''エラー情報構造体初期化
    CrystalCJJudg = SetErrInfo(ErrInfo)
    
    '*** パターン判定 ***
    bJudg = CudecoJudgPattern(CuDeco.HSXCJPK, CuDeco.CJPTNJSK, ptnOK)

    ' CJ Ring内径・外径の判定
    If bJudg Then
        '' パターン実績が[Ring] or [Disk & Ring]
        If (CuDeco.CJPTNJSK = CudecoJskPtnR) Or (CuDeco.CJPTNJSK = CudecoJskPtnDR) Then
            ' 共通Ring内径下限値が未入力か150より大きい場合
            If (CuDeco.CJALLMINRINC5 = -1) Or (CuDeco.CJALLMINRINC5 > 150) Then
                bJudg = False
            ' Ring内径実績が未入力か150より大きい場合
            ElseIf (CuDeco.CJRINGNKJSK = -1) Or (CuDeco.CJRINGNKJSK > 150) Then
                bJudg = False
            ' 共通Ring外径上限値が未入力か150より大きい場合
            ElseIf (CuDeco.CJALLMAXRIGC5 = -1) Or (CuDeco.CJALLMAXRIGC5 > 150) Then
                bJudg = False
            ' Ring外径実績が未入力か150より大きい場合
            ElseIf (CuDeco.CJRINGGKJSK = -1) Or (CuDeco.CJRINGGKJSK > 150) Then
                bJudg = False
            ' 共通Ring内径下限値 > Ring内径実績の場合
            ElseIf (CuDeco.CJALLMINRINC5 > CuDeco.CJRINGNKJSK) Then
                bJudg = False
            ' 共通Ring外径上限値 > Ring外径実績
            ElseIf (CuDeco.CJALLMAXRIGC5 < CuDeco.CJRINGGKJSK) Then
                bJudg = False
            End If
        End If
    End If
    
    ' CJ Disk半径の判定
    If bJudg Then
        '' パターン実績が[Disk] or [Disk & Ring]
        If (CuDeco.CJPTNJSK = CudecoJskPtnD) Or (CuDeco.CJPTNJSK = CudecoJskPtnDR) Then
            ' 共通Disk半径上限値が未入力か150より大きい場合
            If (CuDeco.CJALLMAXDIC5 = -1) Or (CuDeco.CJALLMAXDIC5 > 150) Then
                bJudg = False
            ' Disk半径実績が未入力か150より大きい場合
            ElseIf (CuDeco.CJDISKJSK = -1) Or (CuDeco.CJDISKJSK > 150) Then
                bJudg = False
            ' 共通Disk半径上限値 < Disk半径実績
            ElseIf (CuDeco.CJALLMAXDIC5 < CuDeco.CJDISKJSK) Then
                bJudg = False
            End If
        End If
    End If

    'CJ 計算Pi幅の判定(上限値チェック)
    If bJudg Then
        '' パターン実績が[Disk] or [Ring] or [Disk & Ring]
        If (CuDeco.CJPTNJSK = CudecoJskPtnD) Or (CuDeco.CJPTNJSK = CudecoJskPtnR) Or (CuDeco.CJPTNJSK = CudecoJskPtnDR) Then
            If (CuDeco.CJPTNJSK = CudecoJskPtnD) Then        '[Disk]
                intMax = CuDeco.CJDMAXPIC5                  'Diskのみパターン Pi幅上限値
            ElseIf (CuDeco.CJPTNJSK = CudecoJskPtnR) Then    '[Ring]
                intMax = CuDeco.CJRMAXPIC5                  'Ringのみパターン Pi幅上限値
            Else                                            '[Disk & Ring]
                intMax = CuDeco.CJDRMAXPIC5                 'DiskRingパターン Pi幅上限値
            End If
            
            'Pi幅上限値が未入力か150より大きい場合
            If (intMax = -1) Or (intMax > 150) Then
                bJudg = False
            'Pi幅計算が未入力か150より大きい場合
            ElseIf (CuDeco.CJPICALC = -1) Or (CuDeco.CJPICALC > 150) Then
                bJudg = False
            ''Pi幅上限値 < Pi幅計算の場合
            ElseIf (intMax < CuDeco.CJPICALC) Then
                bJudg = False
            End If
        End If
    End If

    CuDeco.JudgCJ = bJudg

    CrystalCJJudg = FUNCTION_RETURN_SUCCESS
    
End Function

'概要      :結晶CJLT判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Cudeco        ,I  ,C_CUDECO         ,結晶Cu-deco判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :
Public Function CrystalCJLTJudg(CuDeco As C_CJLT, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns         As FUNCTION_RETURN
    Dim ptnOK(6)        As String
    Dim bJudg           As Boolean

    '' パターン判定情報
    '' １文字目はパターン区分、２文字以降がＯＫとするバターン実績を設定する
    ''
    ''　    バターン実績の種類
    ''          "1" : リング無し指定    "2" : ディスク無し指定      "3" : パターン無し指定
    ''          "4" : 不問 (選択なし)   "5" : バンド無し指定        "6" : Pバンド無し指定
    ''          "7" : Bバンド指定無し
    ''
    ''　    バターン実績の種類
    ''          "0":None    "1":Ring    "2":Disk    "3":Disk & Ring
    ''          "5":PB-band "6":P-band  "7":B-band
    ptnOK(0) = "1 0567"         '' リング無し指定
    ptnOK(1) = "2 0567"         '' ディスク無し指定
    ptnOK(2) = "3 0"            '' パターン無し指定
    ptnOK(3) = "4 0567"         '' 不問 (選択なし)
    ptnOK(4) = "5 0"            '' バンド無し指定
    ptnOK(5) = "6 07"           '' リング無し指定
    ptnOK(6) = "7 06"           '' Bバンド指定無し

    ''エラー情報構造体初期化
    CrystalCJLTJudg = SetErrInfo(ErrInfo)
    
    '*** パターン判定 ***
    bJudg = CudecoJudgPattern(CuDeco.HSXCJLTPK, CuDeco.CJLTPTNJSK, ptnOK)

    ' CJ(LT) 計算Band幅の判定
    If bJudg Then
        '' パターン実績が[PB-band] or [P-band] or [B-band]
        ''Del Start 2011/05/13 Y.Hitomi 全パターン共通化
'        If (CuDeco.CJLTPTNJSK = CudecoJskPtnPB_B) Or (CuDeco.CJLTPTNJSK = CudecoJskPtnP_B) Or (CuDeco.CJLTPTNJSK = CudecoJskPtnB_B) Then
        ''Del End   2011/05/13 Y.Hitomi
            ' Band幅上限値が未入力か150より大きい場合
            If (CuDeco.HSXCJLTBND = -1) Or (CuDeco.HSXCJLTBND > 150) Then
                bJudg = False
            ' Band外径実績が未入力か150より大きい場合
            ElseIf (CuDeco.CJLTBANDGKJSK = -1) Or (CuDeco.CJLTBANDGKJSK > 150) Then
                bJudg = False
            ' Band内径実績が未入力か150より大きい場合
            ElseIf (CuDeco.CJLTBANDNKJSK = -1) Or (CuDeco.CJLTBANDNKJSK > 150) Then
                bJudg = False
            ' Band外径実績 < Band内径実績の場合
            ElseIf (CuDeco.CJLTBANDGKJSK < CuDeco.CJLTBANDNKJSK) Then
                bJudg = False
            ' Band幅上限値 < (Band外径実績-Band内径実績)の場合
            ElseIf (CuDeco.HSXCJLTBND < (CuDeco.CJLTBANDGKJSK - CuDeco.CJLTBANDNKJSK)) Then
                bJudg = False
            End If
'        End If
    End If

    CuDeco.JudgCJLT = bJudg

    CrystalCJLTJudg = FUNCTION_RETURN_SUCCESS
    
End Function

'概要      :結晶CJ2判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Cudeco        ,I  ,C_CUDECO         ,結晶Cu-deco判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :
Public Function CrystalCJ2Judg(CuDeco As C_CJ2, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns         As FUNCTION_RETURN
    Dim ptnOK(3)        As String
    Dim bJudg           As Boolean
    Dim intMin          As Integer

    '' パターン判定情報
    '' １文字目はパターン区分、２文字以降がＯＫとするバターン実績を設定する
    ''
    ''　    バターン実績の種類
    ''          "1" : リング無し指定    "2" : ディスク無し指定      "3" : パターン無し指定
    ''          "4" : 不問 (選択なし)   "5" : バンド無し指定        "6" : Pバンド無し指定
    ''          "7" : Bバンド指定無し
    ''
    ''　    バターン実績の種類
    ''          "0":None    "1":Ring    "2":Disk    "3":Disk & Ring
    ''          "5":PB-band "6":P-band  "7":B-band
    ptnOK(0) = "1 02"        '' リング無し指定
    ptnOK(1) = "2 01"        '' ディスク無し指定
    ptnOK(2) = "3 0"         '' パターン無し指定
    ptnOK(3) = "4 0123"      '' 不問 (選択なし)
    
    ''エラー情報構造体初期化
    CrystalCJ2Judg = SetErrInfo(ErrInfo)
    
    '*** パターン判定 ***
    bJudg = CudecoJudgPattern(CuDeco.HSXCJ2PK, CuDeco.CJ2PTNJSK, ptnOK)

    ' CJ2 Ring内径・外径の判定
    If bJudg Then
        '' パターン実績が[Ring]
        If (CuDeco.CJ2PTNJSK = CudecoJskPtnR) Then
            ' Ringのみパターン Ring内径下限値が未入力か150より大きい場合
            If (CuDeco.CJ2RMINRINC5 = -1) Or (CuDeco.CJ2RMINRINC5 > 150) Then
                bJudg = False
            ' Ring内径実績が未入力か150より大きい場合
            ElseIf (CuDeco.CJ2RINGNKJSK = -1) Or (CuDeco.CJ2RINGNKJSK > 150) Then
                bJudg = False
            ' Ringのみパターン Ring外径上限値が未入力か150より大きい場合
            ElseIf (CuDeco.CJ2RMAXRIGC5 = -1) Or (CuDeco.CJ2RMAXRIGC5 > 150) Then
                bJudg = False
            ' Ring外径実績が未入力か150より大きい場合
            ElseIf (CuDeco.CJ2RINGGKJSK = -1) Or (CuDeco.CJ2RINGGKJSK > 150) Then
                bJudg = False
            ' Ringのみパターン Ring内径下限値 > Ring内径実績の場合
            ElseIf (CuDeco.CJ2RMINRINC5 > CuDeco.CJ2RINGNKJSK) Then
                bJudg = False
            ' Ringのみパターン Ring外径上限値 > Ring外径実績の場合
            ElseIf (CuDeco.CJ2RMAXRIGC5 < CuDeco.CJ2RINGGKJSK) Then
                bJudg = False
            End If
        '' パターン実績が[Disk & Ring]
        ElseIf (CuDeco.CJ2PTNJSK = CudecoJskPtnDR) Then
            ' DiskRingパターン Ring内径下限値が未入力か150より大きい場合
            If (CuDeco.CJ2DRMINRINC5 = -1) Or (CuDeco.CJ2DRMINRINC5 > 150) Then
                bJudg = False
            ' Ring内径実績が未入力か150より大きい場合
            ElseIf (CuDeco.CJ2RINGNKJSK = -1) Or (CuDeco.CJ2RINGNKJSK > 150) Then
                bJudg = False
            ' DiskRingパターン Ring外径上限値が未入力か150より大きい場合
            ElseIf (CuDeco.CJ2DRMAXRIGC5 = -1) Or (CuDeco.CJ2DRMAXRIGC5 > 150) Then
                bJudg = False
            ' Ring外径実績が未入力か150より大きい場合
            ElseIf (CuDeco.CJ2RINGGKJSK = -1) Or (CuDeco.CJ2RINGGKJSK > 150) Then
                bJudg = False
            ' DiskRingパターン Ring内径下限値 > Ring内径実績の場合
            ElseIf (CuDeco.CJ2DRMINRINC5 > CuDeco.CJ2RINGNKJSK) Then
                bJudg = False
            ' DiskRingパターン Ring外径上限値 < Ring外径実績
            ElseIf (CuDeco.CJ2DRMAXRIGC5 < CuDeco.CJ2RINGGKJSK) Then
                bJudg = False
            End If
        End If
    End If

    'CJ2 計算Pi幅の判定(下限値チェック)
    If bJudg Then
        '' パターン実績が[Disk] or [Ring] or [Disk & Ring]
        If (CuDeco.CJ2PTNJSK = CudecoJskPtnD) Or (CuDeco.CJ2PTNJSK = CudecoJskPtnR) Or (CuDeco.CJ2PTNJSK = CudecoJskPtnDR) Then
            If (CuDeco.CJ2PTNJSK = CudecoJskPtnD) Then       '[Disk]
                intMin = CuDeco.CJ2DMAXPIC5     ' Diskのみパターン Pi幅上限値(下限値として使用)
            ElseIf (CuDeco.CJ2PTNJSK = CudecoJskPtnR) Then   '[Ring]
                intMin = CuDeco.CJ2RMAXPIC5     ' Ringのみパターン Pi幅上限値(下限値として使用)
            Else                                            '[Disk & Ring]
                intMin = CuDeco.CJ2DRMAXPIC5    ' DiskRingパターン Pi幅上限値(下限値として使用)
            End If
            
            'Pi幅下限値が未入力か150より大きい場合
            If (intMin = -1) Or (intMin > 150) Then
                bJudg = False
            'Pi幅計算が未入力か150より大きい場合
            ElseIf (CuDeco.CJ2PICALC = -1) Or (CuDeco.CJ2PICALC > 150) Then
                bJudg = False
            'Pi幅下限値 > Pi幅計算の場合
            ElseIf (intMin > CuDeco.CJ2PICALC) Then
                bJudg = False
            End If
        End If
    End If

    CuDeco.JudgCJ2 = bJudg

    CrystalCJ2Judg = FUNCTION_RETURN_SUCCESS
    
End Function

'Add End   2011/01/24 SMPK Miyata

'概要      :対象コードに従って判定対象データを求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :Flag          ,I  ,GUARANTEE ,対象コード
'          :d()           ,I  ,double    ,測定値
'          :d1()          ,O  ,double    ,判定対象データ
'          :戻り値        ,O  ,FUNCTION_RETURN,
'説明      :Flag.cObjの値,
'          :1       ,d1(0)=中心測定値
'          :2       ,d1(0)=中央測定値
'          :3       ,d1()=全測定点
'          :4       ,d1(0)=R/2
'          :A       ,d1(0)=平均値
'          :B       ,d1(0)=最大値
'          :C       ,d1(0)=平均値,d1(1)=最大値
'          :D       ,d1(0)=最小値
'          :E       ,d1(0～3)=内周部2点、外周部2点(5点測定で1,2,4,5)
'          :F       ,d1(0)=2,4点目の内大きい値
'          :G       ,d1(0)=2,3,4点目の内大きい値
'履歴      :2001/06/06 佐野 信哉 作成
Public Function GetCrystalJudgData(flag As Guarantee, d() As Double, d1() As Double) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim c0 As Integer
    Dim COUNT As Integer
    Dim High As Integer
    
    '' 配列の上限を取得します。
    High = UBound(d)
    
    FuncAns = FUNCTION_RETURN_SUCCESS '' 正常
    Select Case flag.cObj
    Case ObjCode01 ''中心測定値
        d1(0) = d(0)
    Case ObjCode02 ''測定値の中央値
        d1(0) = JudgCenter(d())
    Case ObjCode03 ''全測定点
        DataCopy d(), d1()
    Case ObjCode04 ''R/2
        Select Case flag.cPos
        Case PosCode01
            If flag.cCount = "1" Then
                d1(0) = d(0)
            Else
                d1(0) = d(1)
            End If
        Case PosCode02, PosCode03, PosCode04, PosCode05, PosCode06, PosCode07, PosCode08
            d1(0) = d(1)
        Case PosCode09
            d1(0) = d(2)
        Case Else
            FuncAns = FUNCTION_RETURN_FAILURE '' 異常
        End Select
    Case ObjCode05 ''全点の平均値
        d1(0) = JudgAve(d())
    Case ObjCode06 ''全点の最大値
        d1(0) = JudgMax(d())
    Case ObjCode07 ''全点の平均値と最大値
        d1(0) = JudgAve(d())
        d1(1) = JudgMax(d())
    Case ObjCode08 ''全点の最小値
        d1(0) = JudgMin(d())
    Case ObjCode09 ''内周部2点、外周部2点(5点測定で1,2,4,5)
        DataCopy d(), d1()
        COUNT = 0
        For c0 = High To 0 Step -1
            If d1(c0) <> -1 Then
                d1(3 - COUNT) = d1(c0)
                COUNT = COUNT + 1
            End If
            If COUNT = 2 Then Exit For
        Next
    Case ObjCode10 ''MAX(2,4点目)
        If (d(1) <> -1) And (d(3) <> -1) Then
            If d(1) >= d(3) Then
                d1(0) = d(1)
            Else
                d1(0) = d(3)
            End If
        Else
            FuncAns = FUNCTION_RETURN_FAILURE '' 異常
        End If
    Case ObjCode11 ''MAX(2,3,4点目)
        If (d(1) <> -1) And (d(2) <> -1) And (d(3) <> -1) Then
            If d(1) >= d(2) Then
                If d(1) >= d(3) Then
                    d1(0) = d(1)
                Else
                    d1(0) = d(3)
                End If
            Else
                If d(2) >= d(3) Then
                    d1(0) = d(2)
                Else
                    d1(0) = d(3)
                End If
            End If
        Else
            ''あり得ないエラー
            FuncAns = FUNCTION_RETURN_FAILURE '' 異常
        End If
''    Case ObjCode12 ''個数保証
''    Case ObjCode13 ''狙い
''    Case ObjCode14 ''形状測定(平坦度、反返り、WARP)
''    Case ObjCode15 ''規格なし
    Case Else
        FuncAns = FUNCTION_RETURN_FAILURE '' 異常
    End Select
    
    GetCrystalJudgData = FuncAns
End Function

'概要      :結晶Oi判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Oi            ,I  ,C_Oi             ,結晶Oi判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2001/07/19 佐野 信哉 作成
Public Function CrystalOiJudg(Oi As C_Oi, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    Dim JData() As Double
    Dim c0 As Integer
    Dim pt As Integer
    
    ReDim JData(UBound(Oi.Oi())) As Double
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    
    Oi.JudgData = -1
    Oi.JudgOi = JUDG_NG
    If Oi.GuaranteeOi.cJudg = JudgCodeC01 Then ''Oi判定有り
        
'' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06 start

'    If Trim(Oi.GuaranteeOi.cCount) = "" Then
'        pt = 1
'    Else
'        pt = Val(Oi.GuaranteeOi.cCount)
'    End If
'    Oi.ORG = RoundUp((RGCal(Oi.Oi(), pt)), 2)

        ''ORG判定
        
        Select Case Oi.GuaranteeOi.cPos
          Case "B", "C", "D", "E", "F", "K", "Y"
              Select Case Oi.GuaranteeOi.cBunp
              Case "A", "B", "C"
                 ''ORG計算
                 Oi.ORG = MENNAI_Cal(OI_JUDG, Oi.Oi(), Oi.GuaranteeOi, Oi.GuaranteeOi.cBunp)

              Case "", " "
                 ''計算区分がスペースの場合は、計算，判定を行わない
                  Oi.ORG = 0
                  Oi.JudgOrg = JUDG_OK
                  GoTo Cal_Escp
              Case Else
                 ''ORG計算　　　コード "A" にて計算
                 If Trim(Oi.GuaranteeOi.cCount) = "" Then
                    pt = 3
                 Else
                    pt = val(Oi.GuaranteeOi.cCount)
                 End If
                 Oi.ORG = RoundUp((RGCal(Oi.Oi(), pt)), 4)

             End Select

          Case Else

             Select Case Oi.GuaranteeOi.cBunp
             Case "A", "B", "C", "D", "E", "N"
                 ''ORG計算
                 Oi.ORG = MENNAI_Cal(OI_JUDG, Oi.Oi(), Oi.GuaranteeOi, Oi.GuaranteeOi.cBunp)

             Case "", " "
                 ''計算区分がスペースの場合は、計算，判定を行わない
                  Oi.ORG = 0
                  Oi.JudgOrg = JUDG_OK
                  GoTo Cal_Escp
             Case Else
                 ''ORG計算　　　コード "A" にて計算
                 If Trim(Oi.GuaranteeOi.cCount) = "" Then
                    pt = 3
                 Else
                    pt = val(Oi.GuaranteeOi.cCount)
                 End If
                 Oi.ORG = RoundUp((RGCal(Oi.Oi(), pt)), 4)

             End Select
        End Select

'' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06 end

'2002/02/27 S.Sano ORGの仕様が0の場合は、判定を行わず必ずOKとする。
'2002/02/27 S.Sano 面内分布計算は行う。
        If Oi.SpecORG = 0 Then                                      '2002/02/27 S.Sano
            Oi.JudgOrg = JUDG_OK                                    '2002/02/27 S.Sano
        Else                                                        '2002/02/27 S.Sano
            If Oi.ORG = -1 Then
                Oi.JudgOrg = JUDG_NG
            Else
                Oi.JudgOrg = RangeDecision_nl(Oi.ORG, 0, Oi.SpecORG)
            End If
        End If                                                      '2002/02/27 S.Sano
        
'' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06 start
Cal_Escp:
'' Res，Oi 面内分布計算式追加依頼  No.030205  yakimura  2003.06.06 end
        
        ''Oi判定
        If (InStr(ObjCodeGrp01, Oi.GuaranteeOi.cObj) <> 0) And (GetCrystalJudgData(Oi.GuaranteeOi, Oi.Oi(), JData()) = FUNCTION_RETURN_SUCCESS) Then
            Select Case Oi.GuaranteeOi.cObj
            Case ObjCode01, ObjCode02, ObjCode04 ''中心1点、中央値、R/2
                Oi.JudgOi = RangeDecision_nl(JData(0), Oi.SpecOiMin, Oi.SpecOiMax)
                Oi.JudgData = JData(0)
            Case ObjCode03 ''全域
                Oi.JudgOi = JUDG_OK
                For c0 = 0 To UBound(JData())
                    If JData(c0) <> -1 Then
                        If RangeDecision_nl(JData(c0), Oi.SpecOiMin, Oi.SpecOiMax) = JUDG_NG Then
                            Oi.JudgOi = JUDG_NG
                        End If
                    End If
                Next
            End Select
            Oi.JudgData = JudgMax(JData())
        Else
            ''対象データ無し
            ''エラー情報構造体に情報を代入。
            FuncAns = SetErrInfo(ErrInfo, EZJ00, OI_JUDG, Oi.GuaranteeOi.cObj)
        End If
        Oi.JudgOi = (Oi.JudgOi And Oi.JudgOrg)
    Else
        Oi.JudgOrg = JUDG_OK
        Oi.JudgOi = JUDG_OK
'        If InStr(JudgCodeC02, Oi.GuaranteeOi.cJudg) = 0 Then
'            ''処理方法データ無し
'            ''エラー情報構造体に情報を代入。
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, OI_JUDG, Oi.GuaranteeOi.cJudg)
'        End If
    End If
    
    CrystalOiJudg = FuncAns
End Function

'概要      :結晶Cs判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :Cs            ,I  ,C_Cs             ,結晶Cs判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :2001/07/19 佐野 信哉 作成
Public Function CrystalCsJudg(Cs As C_Cs, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim FuncAns As FUNCTION_RETURN
    
    ''エラー情報構造体初期化
    FuncAns = SetErrInfo(ErrInfo)
    FuncAns = FUNCTION_RETURN_SUCCESS
    If Cs.GuaranteeCs.cJudg = JudgCodeC01 Then ''Cs判定有り
        'BOT保証は上限ﾁｪｯｸ,TOP/BOT保証は上下限ﾁｪｯｸ 09/01/08 ooba
        If Cs.SpecCsKHI = "6" Or Cs.SpecCsKHI = "9" Then
            Cs.JudgCs = RangeDecision_nl(Cs.Cs, Cs.SpecCsMin, Cs.SpecCsMax)
        Else
            Cs.JudgCs = RangeDecision_nl(Cs.Cs, -1, Cs.SpecCsMax)
        End If
    Else
        Cs.JudgCs = JUDG_OK
'        If InStr(JudgCodeC02, Cs.GuaranteeCs.cJudg) = 0 Then
'            ''処理方法データ無し
'            ''エラー情報構造体に情報を代入。
'            FuncAns = SetErrInfo(ErrInfo, ZJ001, CS_JUDG, Cs.GuaranteeCs.cJudg)
'        End If
    End If
    
    CrystalCsJudg = FuncAns
End Function


#If NO_FURIKAECHECK = 0 Then
'概要      :抜試以降での品番振替時に、結晶側の判定を行う
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型          ,説明
'          :crynum        ,I  ,String      ,結晶番号
'          :ingotpos      ,I  ,Integer     ,対象範囲の開始位置
'          :length        ,I  ,Integer     ,対象範囲の長さ
'          :hin           ,I  ,tFullHinban ,振替先の品番
'          :judge_ok      ,O  ,Boolean     ,判定結果
'          :itemNG        ,O  ,String      ,判定NGとなった項目
'          :戻り値        ,O  ,FUNCTION_RETURN, 判定の合否
'          :                   FUNCTION_RETURN_SUCCESS: 振替可
'          :                   FUNCTION_RETURN_FAILURE: 振替不可もしくは仕様エラー
'説明      :結晶保証のみの項目である GD/LT/Cs について判定する
'履歴      :2002/03/xx 佐野 信哉 作成
Public Function SXLJudge(CRYNUM$, INGOTPOS%, Length%, HIN As tFullHinban, judge_ok As Boolean, itemNG$) As FUNCTION_RETURN
    Dim GD(1) As C_GD
    Dim Cs(1) As C_Cs
    Dim Lt(1) As C_LT
    Dim ErrInfo As ERROR_INFOMATION
    
    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzccj.bas -- Function SXLJudge"

    SXLJudge = FUNCTION_RETURN_FAILURE
    If scmzc_getSXLGuarantee(HIN, GD(), Cs(), Lt()) = FUNCTION_RETURN_FAILURE Then
        '仕様取得エラー
        GoTo proc_exit
    End If
    
    judge_ok = False
    itemNG$ = ""
    
    'GD判定
    '検査実績を取得する
    If scmzc_getSXLGD(CRYNUM$, INGOTPOS%, Length%, GD()) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    'Top位置の判定を行う
    If CrystalGDJudg(GD(0), ErrInfo) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    'Bot位置の検査実績を取得する
    ElseIf CrystalGDJudg(GD(1), ErrInfo) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    'NGなら抜ける
    SXLJudge = FUNCTION_RETURN_SUCCESS
    If (Not GD(0).JudgDen) Or (Not GD(1).JudgDen) Then
        itemNG$ = "DEN"
        GoTo proc_exit
    ElseIf (Not GD(0).JudgDvd2) Or (Not GD(1).JudgDvd2) Then
        itemNG$ = "DVD2"
        GoTo proc_exit
    ElseIf (Not GD(0).JudgLdl) Or (Not GD(1).JudgLdl) Then
        itemNG$ = "L/DL"
        GoTo proc_exit
    End If

    'Cs判定
    '検査実績を取得する
    SXLJudge = FUNCTION_RETURN_FAILURE
    If scmzc_getSXLCs(CRYNUM$, INGOTPOS%, Length%, Cs()) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    'Top位置の判定を行う
    If (Cs(0).GuaranteeCs.cJudg = "H") And (Cs(0).SpecCsMin > 0#) Then
        'CsがFromTo保証の場合は、Top側判定を行う
        If CrystalCsJudg(Cs(0), ErrInfo) = FUNCTION_RETURN_FAILURE Then
            GoTo proc_exit
        End If
    Else
        Cs(0).JudgCs = True
    End If

    'Bot位置の判定を行う
    If CrystalCsJudg(Cs(1), ErrInfo) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    'NGなら抜ける
    SXLJudge = FUNCTION_RETURN_SUCCESS
    If (Not Cs(0).JudgCs) Or (Not Cs(1).JudgCs) Then
        itemNG$ = "CS"
        GoTo proc_exit
    End If
    
    'LT判定
    '検査実績を取得する
    SXLJudge = FUNCTION_RETURN_FAILURE
    If scmzc_getSXLLt(CRYNUM$, INGOTPOS%, Length%, Lt()) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    'Bot位置の判定を行う(LTはBot側だけの保証)
    Lt(0).JudgLt = True
    If CrystalLTJudg(Lt(1), ErrInfo) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    'NGなら抜ける
    SXLJudge = FUNCTION_RETURN_SUCCESS
    If (Not Lt(1).JudgLt) Then
        itemNG$ = "LT"
        GoTo proc_exit
    End If

    judge_ok = True

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :WF側での品番振替時チェックを行う
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :orgSXL        ,I  ,c_cmzcSxls   ,振替前のSXL構成
'          :wfSmps()      ,I  ,typ_XSDCW    ,新サンプル管理（SXL）
'          :Crynum        ,I  ,String       ,結晶番号
'          :lblMsg        ,I  ,Label        ,メッセージ表示エリア
'          :needPreJudge  ,I  ,Boolean      ,その他の結晶側判定も行う
'          :chkFrom       ,I  ,Integer      ,チェック範囲(mm)
'          :chkTo         ,I  ,Integer      ,チェック範囲(mm)
'          :戻り値        ,O  ,FUNCTION_RETURN,
'説明      :チェック対象は、Cs,GD,LT
'履歴      :
'Public Function FurikaeCheck(orgSXL As c_cmzcSxls, WfSmps() As typ_XSDCW, CRYNUM$, lblMsg As Label, needPreJudge As Boolean, chkFrom As Integer, chkTo As Integer) As FUNCTION_RETURN
Public Function FurikaeCheck(orgSXL As c_cmzcSxls, WfSmps() As typ_XSDCW, CRYNUM$, lblMsg As Label, chkFrom As Integer, chkTo As Integer) As fHinban
    Dim HIN As tFullHinban '2002/03/14 S.Sano
    Dim c0 As Integer
    Dim c1 As Integer
    Dim judge_ok As Boolean
    Dim itemNG$
    Dim eqf As Boolean
    Dim hinban$
    Dim pos1 As Integer
    Dim pos2 As Integer
    Dim nSxl As Integer
    Dim fHin As fHinban

''    FurikaeCheck = FUNCTION_RETURN_FAILURE

    ReDim buff$(0)
'    For c0 = 1 To UBound(WfSmps) - 1
        pos1 = WfSmps(c0).INPOSCW
        pos2 = WfSmps(c0 + 1).INPOSCW
        If (pos1 >= chkFrom) And (pos2 <= chkTo) Then
            hinban = Trim(WfSmps(c0).HINBCW)
            If (hinban <> "Z") And (hinban <> "G") And (hinban <> vbNullString) Then
                '品番が変わっていなければ、スキップする。
                eqf = True
                If (WfSmps(c0).SMPKBNCW = "U") Or (WfSmps(c0).SMPKBNCW = "B") Then
                    nSxl = orgSXL.UpperArea(pos1)
                Else
                    nSxl = orgSXL.LowerArea(pos1)
                End If
                If Abs(nSxl) <> 9999 Then
                    If hinban <> orgSXL(CStr(nSxl)).hinban Then
'                        eqf = False
                '構造体に値を保持する
                fHin.moto = hinban
                fHin.saki = orgSXL(CStr(nSxl)).hinban
                    End If
                End If

        '品番の振替は関数で行うため振替チェックの機能を削除-------start iida 2003/09/05
'                If Not eqf Then
'                    If GetLastHinban(HINBAN$, hin) = FUNCTION_RETURN_FAILURE Then
'                        lblMsg.Caption = GetMsgStr("EHIN8", vbNullString) '03/06/06 後藤
'                        Exit Function
'                    End If

'                    '基本判定
'                    If needPreJudge Then
'                        If SXLPreJudge(CRYNUM$, pos1, pos2 - pos1, hin, judge_ok, itemNG$) = FUNCTION_RETURN_FAILURE Then
'                            lblMsg.Caption = GetMsgStr("EHIN8", "(" & itemNG & ")")   '03/06/06 後藤
'                            Exit Function
'                        End If
'                        If Not judge_ok Then
'                        lblMsg.Caption = GetMsgStr("EHIN9", Trim(HINBAN$) & " " & itemNG$)    '03/06/06 後藤
'                            Exit Function
'                        End If
'                    End If

                    'GD/Cs/LT判定
'                    If SXLJudge(CRYNUM, pos1, pos2 - pos1, hin, judge_ok, itemNG$) = FUNCTION_RETURN_FAILURE Then
'                        lblMsg.Caption = GetMsgStr("EHIN8", vbNullString)   '03/06/06 後藤
'                        Exit Function
'                    End If
'                    If Not judge_ok Then
'                        lblMsg.Caption = GetMsgStr("EHIN9", Trim(HINBAN$) & " " & itemNG$)        '03/06/06 後藤
'                        Exit Function
'                    End If
'                End If
        '品番振替は関数で行うため振替チェックの機能を削除-------end iida 2003/09/05
            End If
        End If
'    Next
''    FurikaeCheck = fHinban
End Function
#End If

'概要      :加工実績の判定を行う
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型              ,説明
'          :Kakou         ,IO ,type_KakouJudg  ,加工実績判定構造体
'          :戻り値        ,O  ,FUNCTION_RETURN, 判定の成否
'説明      :加工実績について判定する
'履歴      :2002/04/17 佐野 信哉 作成
Public Function FormJudg(Kakou As type_KakouJudg) As FUNCTION_RETURN
    Dim c0 As Integer
    
    FormJudg = FUNCTION_RETURN_FAILURE
    
    Kakou.Judg.top = False
    Kakou.Judg.tTOP(1) = False
    Kakou.Judg.tTOP(2) = False
    
    Kakou.Judg.TAIL = False
    Kakou.Judg.tTAIL(1) = False
    Kakou.Judg.tTAIL(2) = False
    
    Kakou.Judg.POS = False
    Kakou.Judg.WIDH = False
    Kakou.Judg.DPTH = False
    Kakou.Judg.ANGLE = False        '2009/09 SUMCO Akizuki
    
    
    ''Notch位置の規格判定
    If InStr("A1", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("A1", Kakou.Spec())
    ElseIf InStr("A2", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("A2", Kakou.Spec())
    ElseIf InStr("B1B2B3B4", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("B1B2B3B4", Kakou.Spec())
    ElseIf InStr("B5B6B7B8", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("B5B6B7B8", Kakou.Spec())
    ElseIf InStr("C1", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("C1", Kakou.Spec())
    ElseIf InStr("C2", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("C2", Kakou.Spec())
    ElseIf InStr("D1D2", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("D1D2", Kakou.Spec())
    ElseIf InStr("D3D4", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("D3D4", Kakou.Spec())
    ElseIf InStr("D5D8", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("D5D8", Kakou.Spec())
    ElseIf InStr("D6D7", Kakou.Jiltuseki.POS) <> 0 Then
        Kakou.Judg.POS = Kakou_Pos_Judg("D6D7", Kakou.Spec())
    ElseIf InStr("ZZ", Kakou.Jiltuseki.POS) <> 0 Then      ''''2005/05/27 ADD
        Kakou.Judg.POS = Kakou_Pos_Judg("ZZ", Kakou.Spec())

    Else
        Exit Function
    End If
    
    
    '' 直径(TOP1,2)の規格チェック
    If (Kakou.Jiltuseki.top(1) = -1) And (Kakou.Jiltuseki.top(2) = -1) Then
        Exit Function
    End If
    Kakou.Judg.tTOP(1) = True
    Kakou.Judg.tTOP(2) = True
    
    '各品番ごとに、仕様規格のチェックを行う
    For c0 = 1 To UBound(Kakou.Spec())
        If Kakou.Spec(c0).top(1) = -1 Or Kakou.Spec(c0).top(2) = -1 Then Exit Function      '2003/12/12 SystemBrain Null対応
        If Kakou.Jiltuseki.top(1) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.top(1), Kakou.Spec(c0).top(1), Kakou.Spec(c0).top(2)) = False Then
                Kakou.Judg.tTOP(1) = False
            End If
        End If
        If Kakou.Jiltuseki.top(2) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.top(2), Kakou.Spec(c0).top(1), Kakou.Spec(c0).top(2)) = False Then
                Kakou.Judg.tTOP(2) = False
            End If
        End If
    Next
    If (Kakou.Jiltuseki.top(1) <> -1) And (Kakou.Jiltuseki.top(2) <> -1) Then
        Kakou.Judg.top = (Kakou.Judg.tTOP(1) And Kakou.Judg.tTOP(2))
    ElseIf Kakou.Jiltuseki.top(1) <> -1 Then
        Kakou.Judg.top = Kakou.Judg.tTOP(1)
    ElseIf Kakou.Jiltuseki.top(2) <> -1 Then
        Kakou.Judg.top = Kakou.Judg.tTOP(2)
    End If
    
    
    
    '' 直径(BOT1,2)の規格チェック
    If (Kakou.Jiltuseki.TAIL(1) = -1) And (Kakou.Jiltuseki.TAIL(2) = -1) Then
        Exit Function
    End If
    
    
    Kakou.Judg.tTAIL(1) = True
    Kakou.Judg.tTAIL(2) = True
    
    For c0 = 1 To UBound(Kakou.Spec())
        If Kakou.Spec(c0).top(1) = -1 Or Kakou.Spec(c0).top(2) = -1 Then Exit Function      '2003/12/12 SystemBrain Null対応
        If Kakou.Jiltuseki.TAIL(1) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.TAIL(1), Kakou.Spec(c0).top(1), Kakou.Spec(c0).top(2)) = False Then
                Kakou.Judg.tTAIL(1) = False
            End If
        End If
        If Kakou.Jiltuseki.TAIL(2) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.TAIL(2), Kakou.Spec(c0).top(1), Kakou.Spec(c0).top(2)) = False Then
                Kakou.Judg.tTAIL(2) = False
            End If
        End If
    Next
    If (Kakou.Jiltuseki.TAIL(1) <> -1) And (Kakou.Jiltuseki.TAIL(2) <> -1) Then
        Kakou.Judg.TAIL = (Kakou.Judg.tTAIL(1) And Kakou.Judg.tTAIL(2))
    ElseIf Kakou.Jiltuseki.TAIL(1) <> -1 Then
        Kakou.Judg.TAIL = Kakou.Judg.tTAIL(1)
    ElseIf Kakou.Jiltuseki.TAIL(2) <> -1 Then
        Kakou.Judg.TAIL = Kakou.Judg.tTAIL(2)
    End If
    
    If (Kakou.Jiltuseki.WIDH(1) = -1) And (Kakou.Jiltuseki.WIDH(2) = -1) Then
        Exit Function
    End If
    
    
    ''Notch幅の規格チェック
    Kakou.Judg.WIDH = True
    
    For c0 = 1 To UBound(Kakou.Spec())
        If Kakou.Spec(c0).WIDH(1) = -1 Or Kakou.Spec(c0).WIDH(2) = -1 Then Exit Function      '2003/12/12 SystemBrain Null対応
        If Kakou.Jiltuseki.WIDH(1) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.WIDH(1), Kakou.Spec(c0).WIDH(1), Kakou.Spec(c0).WIDH(2)) = False Then
                Kakou.Judg.WIDH = False
            End If
        End If
        If Kakou.Jiltuseki.WIDH(2) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.WIDH(2), Kakou.Spec(c0).WIDH(1), Kakou.Spec(c0).WIDH(2)) = False Then
                Kakou.Judg.WIDH = False
            End If
        End If
    Next
    
    If (Kakou.Jiltuseki.DPTH(1) = -1) And (Kakou.Jiltuseki.DPTH(2) = -1) Then
        Exit Function
    End If
    
    
    '' Notch深さ(TOP･BOT)の規格チェック
    Kakou.Judg.DPTH = True
    
    For c0 = 1 To UBound(Kakou.Spec())
        If Kakou.Spec(c0).DPTH(1) = -1 Or Kakou.Spec(c0).DPTH(2) = -1 Then Exit Function      '2003/12/12 SystemBrain Null対応
        If Kakou.Jiltuseki.DPTH(1) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.DPTH(1), Kakou.Spec(c0).DPTH(1), Kakou.Spec(c0).DPTH(2)) = False Then
                Kakou.Judg.DPTH = False
            End If
        End If
        If Kakou.Jiltuseki.DPTH(2) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.DPTH(2), Kakou.Spec(c0).DPTH(1), Kakou.Spec(c0).DPTH(2)) = False Then
                Kakou.Judg.DPTH = False
            End If
        End If
    Next

    
    
    '' Notch角度の規格チェック      2009/09 SUMOCO Akizuki
    Kakou.Judg.ANGLE = True
    
    For c0 = 1 To UBound(Kakou.Spec())
        If Kakou.Spec(c0).ANGLE(1) = -1 Or Kakou.Spec(c0).ANGLE(2) = -1 Then Exit Function      '2003/12/12 SystemBrain Null対応
        If Kakou.Jiltuseki.ANGLE(1) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.ANGLE(1), Kakou.Spec(c0).ANGLE(1), Kakou.Spec(c0).ANGLE(2)) = False Then
                Kakou.Judg.ANGLE = False
            End If
        End If
        If Kakou.Jiltuseki.ANGLE(2) <> -1 Then
            If RangeDecision(Kakou.Jiltuseki.ANGLE(2), Kakou.Spec(c0).ANGLE(1), Kakou.Spec(c0).ANGLE(2)) = False Then
                Kakou.Judg.ANGLE = False
            End If
        End If
    Next

    FormJudg = FUNCTION_RETURN_SUCCESS
    
End Function

'概要      :加工実績の位置判定を行う
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型              ,説明
'          :sPos          ,I  ,String         ,位置グループ文字列
'          :Spec()        ,I  ,type_KakouJudg ,加工仕様構造体
'          :戻り値        ,O  ,Boolean         ,判定結果
'説明      :加工実績位置について判定する内部関数
'履歴      :2002/04/17 佐野 信哉 作成
Public Function Kakou_Pos_Judg(sPos As String, Spec() As Judg_Kakou) As Boolean
    Dim c0 As Integer
    Dim tJudg As Boolean
    tJudg = True
    For c0 = 1 To UBound(Spec())
        If tJudg Then
            tJudg = (InStr(sPos, Spec(c0).POS) <> 0)
        End If
    Next
    Kakou_Pos_Judg = tJudg
End Function
'概要      :結晶ブロック偏析判定を行う。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :COEF          ,I  ,C_COEF            ,判定構造体
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,エラー情報構造体
'          :戻り値        ,O  ,FUNCTION_RETURN  ,
'説明      :
'履歴      :
Public Function CrystalCOEFJudg(COEF As C_COEF, ErrInfo As ERROR_INFOMATION) As FUNCTION_RETURN
    Dim sMin As Double
    Dim sMax As Double
    
    ''エラー情報構造体初期化
    CrystalCOEFJudg = SetErrInfo(ErrInfo)
    Select Case COEF.NP
        Case "p-"
            sMin = PminusMin
            sMax = PminusMax
        Case "p+"
            sMin = PplusMin
            sMax = PplusMax
        Case "n"
            sMin = NMin
            sMax = NMax
    End Select
    COEF.JudgCOEF = RangeDecision_nl(COEF.COEF, sMin, sMax)
    CrystalCOEFJudg = FUNCTION_RETURN_SUCCESS
End Function


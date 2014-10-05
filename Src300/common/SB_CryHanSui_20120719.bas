Attribute VB_Name = "SB_CryHanSui"
Option Explicit

'-------------------------------------------------------------------------------
' 定数定義
'-------------------------------------------------------------------------------
'XSDCS
Private Const cCRYSMPLID    As String = "CRYSMPLID"     'XSDCSのサンプルＩＤ
Private Const cCRYIND       As String = "CRYIND"        'XSDCSの状態FLG
Private Const cCRYRES       As String = "CRYRES"        'XSDCSの実績FLG
Private Const cCS           As String = "CS"            'XSDCSの項目最終文字
Private Const cCRY_RS       As String = "RS"            'XSDCSのRs
Private Const cCRY_OI       As String = "OI"            'XSDCSのOi
Private Const cCRY_B1       As String = "B1"            'XSDCSのBMD1
Private Const cCRY_B2       As String = "B2"            'XSDCSのBMD2
Private Const cCRY_B3       As String = "B3"            'XSDCSのBMD3
Private Const cCRY_O1       As String = "L1"            'XSDCSのOSF1
Private Const cCRY_O2       As String = "L2"            'XSDCSのOSF2
Private Const cCRY_O3       As String = "L3"            'XSDCSのOSF3
Private Const cCRY_O4       As String = "L4"            'XSDCSのOSF4
Private Const cCRY_CS       As String = "CS"            'XSDCSのCs
Private Const cCRY_GD       As String = "GD"            'XSDCSのGD
Private Const cCRY_LT       As String = "T"             'XSDCSのLT
Private Const cCRY_EP       As String = "EP"            'XSDCSのEPD
'Add Start 2011/01/19 SMPK Miyata
Private Const cCRY_C        As String = "C"             'XSDCSのC
Private Const cCRY_CJ       As String = "CJ"            'XSDCSのCJ
Private Const cCRY_CJLT     As String = "CJLT"          'XSDCSのCJLT
Private Const cCRY_CJ2      As String = "CJ2"           'XSDCSのCJ2
'Add End   2011/01/19 SMPK Miyata

'結晶抵抗実績
Public Type type_DBDRV_scmzc_fcmkc001c_CryR
    CRYNUM      As String * 12          ' 結晶番号
    POSITION    As Integer              ' 位置
    SMPKBN      As String * 1           ' サンプル区分
    TRANCOND    As String * 1           ' 処理条件
    TRANCNT     As String * 1           ' 処理回数
    SMPLNO      As Long                 ' サンプルＮｏ  Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU     As String * 1           ' サンプル有無
    MEAS1       As Double               ' 測定値１
    MEAS2       As Double               ' 測定値２
    MEAS3       As Double               ' 測定値３
    MEAS4       As Double               ' 測定値４
    MEAS5       As Double               ' 測定値５
    EFEHS       As Double               ' 実効偏析
    RRG         As Double               ' ＲＲＧ
    REGDATE     As Date                 ' 登録日付
    '-----TEST2004/10
    JMEAS1       As Double               ' 測定値１
    JMEAS2       As Double               ' 測定値２
    JMEAS3       As Double               ' 測定値３
    JMEAS4       As Double               ' 測定値４
    JMEAS5       As Double               ' 測定値５
    KSTAFFID     As String
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    HSXDKTMP     As String * 1          'DK温度(実績)
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
End Type

'Oi実績
Public Type type_DBDRV_scmzc_fcmkc001c_Oi
    CRYNUM      As String * 12          ' 結晶番号
    POSITION    As Integer              ' 位置
    SMPKBN      As String * 1           ' サンプル区分
    TRANCOND    As String * 1           ' 処理条件
    TRANCNT     As String * 1           ' 処理回数
    SMPLNO      As Long                 ' サンプルＮｏ  Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU     As String * 1           ' サンプル有無
    OIMEAS1     As Double               ' Ｏｉ測定値１
    OIMEAS2     As Double               ' Ｏｉ測定値２
    OIMEAS3     As Double               ' Ｏｉ測定値３
    OIMEAS4     As Double               ' Ｏｉ測定値４
    OIMEAS5     As Double               ' Ｏｉ測定値５
    ORGRES      As Double               ' ＯＲＧ結果
    AVE         As Double               ' ＡＶＥ
    FTIRCONV    As Double               ' ＦＴＩＲ換算
    INSPECTWAY  As String * 2           ' 検査方法
    REGDATE     As Date                 ' 登録日付
End Type

'BMD1～3実績
Public Type type_DBDRV_scmzc_fcmkc001c_BMD
    CRYNUM      As String * 12          ' 結晶番号
    POSITION    As Integer              ' 位置
    SMPKBN      As String * 1           ' サンプル区分
    TRANCOND    As String * 1           ' 処理条件
    TRANCNT     As String * 1           ' 処理回数
    SMPLNO      As Long                 ' サンプルＮｏ  Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU     As String * 1           ' サンプル有無
    HTPRC       As String * 2           ' 熱処理方法
    KKSP        As String * 3           ' 結晶欠陥測定位置
    KKSET       As String * 3           ' 結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    MEAS1       As Double               ' 測定値１
    MEAS2       As Double               ' 測定値２
    MEAS3       As Double               ' 測定値３
    MEAS4       As Double               ' 測定値４
    MEAS5       As Double               ' 測定値５
    MEASMIN     As Double               ' MIN
    MEASMAX     As Double               ' MAX
    MEASAVE     As Double               ' AVE
    BMDMNBUNP   As Double               ' BMD面内分布
    REGDATE     As Date                 ' 登録日付
End Type

'OSF1～4実績
Public Type type_DBDRV_scmzc_fcmkc001c_OSF
    CRYNUM      As String * 12          ' 結晶番号
    POSITION    As Integer              ' 位置
    SMPKBN      As String * 1           ' サンプル区分
    TRANCOND    As String * 1           ' 処理条件
    TRANCNT     As Integer              ' 処理回数      String * 1 -> Integer 2008/10/28 L/DL,OSF判定ﾛｼﾞｯｸ追加(IT) UPD By Systech
    SMPLNO      As Long                 ' サンプルＮｏ  Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU     As String * 1           ' サンプル有無
    HTPRC       As String * 2           ' 熱処理方法
    KKSP        As String * 3           ' 結晶欠陥測定位置
    KKSET       As String * 3           ' 結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
    CALCMAX     As Double               ' 計算結果 Max
    CALCAVE     As Double               ' 計算結果 Ave
    MEAS1       As Double               ' 測定値１
    MEAS2       As Double               ' 測定値２
    MEAS3       As Double               ' 測定値３
    MEAS4       As Double               ' 測定値４
    MEAS5       As Double               ' 測定値５
    MEAS6       As Double               ' 測定値６
    MEAS7       As Double               ' 測定値７
    MEAS8       As Double               ' 測定値８
    MEAS9       As Double               ' 測定値９
    MEAS10      As Double               ' 測定値１０
    MEAS11      As Double               ' 測定値１１
    MEAS12      As Double               ' 測定値１２
    MEAS13      As Double               ' 測定値１３
    MEAS14      As Double               ' 測定値１４
    MEAS15      As Double               ' 測定値１５
    MEAS16      As Double               ' 測定値１６
    MEAS17      As Double               ' 測定値１７
    MEAS18      As Double               ' 測定値１８
    MEAS19      As Double               ' 測定値１９
    MEAS20      As Double               ' 測定値２０
    OSFPOS1     As Double               ' ﾊﾟﾀｰﾝ区分１位置
    OSFWID1     As Double               ' ﾊﾟﾀｰﾝ区分１幅
    OSFRD1      As String * 1           ' ﾊﾟﾀｰﾝ区分１R/D
    OSFPOS2     As Double               ' ﾊﾟﾀｰﾝ区分２位置
    OSFWID2     As Double               ' ﾊﾟﾀｰﾝ区分２幅
    OSFRD2      As String * 1           ' ﾊﾟﾀｰﾝ区分２R/D
    OSFPOS3     As Double               ' ﾊﾟﾀｰﾝ区分３位置
    OSFWID3     As Double               ' ﾊﾟﾀｰﾝ区分３幅
    OSFRD3      As String * 1           ' ﾊﾟﾀｰﾝ区分３R/D
    CALCMH      As Double               ' 面内比(MAX/MIN)   '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
    PTNJUDGRES  As String               ' パターン判定結果   '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
    REGDATE     As Date                 ' 登録日付
End Type

'CS実績
Public Type type_DBDRV_scmzc_fcmkc001c_CS
    CRYNUM      As String * 12          ' 結晶番号
    POSITION    As Integer              ' 位置
    SMPKBN      As String * 1           ' サンプル区分
    TRANCOND    As String * 1           ' 処理条件
    TRANCNT     As String * 1           ' 処理回数
    SMPLNO      As Long                 ' サンプルＮｏ  Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU     As String * 1           ' サンプル有無
    CSMEAS      As Double               ' Cs実測値
    PRE70P      As Double               ' ７０％推定値
    INSPECTWAY  As String * 2           ' 検査方法
    REGDATE     As Date                 ' 登録日付
End Type

'GD実績
Public Type type_DBDRV_scmzc_fcmkc001c_GD
    CRYNUM      As String * 12          ' 結晶番号
    POSITION    As Integer              ' 位置
    SMPKBN      As String * 1           ' サンプル区分
    TRANCOND    As String * 1           ' 処理条件
    TRANCNT     As Integer              ' 処理回数      String * 1 -> Integer 2008/10/28 L/DL,OSF判定ﾛｼﾞｯｸ追加(IT) UPD By Systech
    SMPLNO      As Long                 ' サンプルＮｏ  Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU     As String * 1           ' サンプル有無
    MSRSDEN     As Integer              ' 測定結果 Den
    MSRSLDL     As Integer              ' 測定結果 L/DL
    MSRSDVD2    As Integer              ' 測定結果 DVD2
    MS01LDL1    As Integer              ' 測定値01 L/DL1
    MS01LDL2    As Integer              ' 測定値01 L/DL2
    MS01LDL3    As Integer              ' 測定値01 L/DL3
    MS01LDL4    As Integer              ' 測定値01 L/DL4
    MS01LDL5    As Integer              ' 測定値01 L/DL5
    MS01DEN1    As Integer              ' 測定値01 Den1
    MS01DEN2    As Integer              ' 測定値01 Den2
    MS01DEN3    As Integer              ' 測定値01 Den3
    MS01DEN4    As Integer              ' 測定値01 Den4
    MS01DEN5    As Integer              ' 測定値01 Den5
    MS02LDL1    As Integer              ' 測定値02 L/DL1
    MS02LDL2    As Integer              ' 測定値02 L/DL2
    MS02LDL3    As Integer              ' 測定値02 L/DL3
    MS02LDL4    As Integer              ' 測定値02 L/DL4
    MS02LDL5    As Integer              ' 測定値02 L/DL5
    MS02DEN1    As Integer              ' 測定値02 Den1
    MS02DEN2    As Integer              ' 測定値02 Den2
    MS02DEN3    As Integer              ' 測定値02 Den3
    MS02DEN4    As Integer              ' 測定値02 Den4
    MS02DEN5    As Integer              ' 測定値02 Den5
    MS03LDL1    As Integer              ' 測定値03 L/DL1
    MS03LDL2    As Integer              ' 測定値03 L/DL2
    MS03LDL3    As Integer              ' 測定値03 L/DL3
    MS03LDL4    As Integer              ' 測定値03 L/DL4
    MS03LDL5    As Integer              ' 測定値03 L/DL5
    MS03DEN1    As Integer              ' 測定値03 Den1
    MS03DEN2    As Integer              ' 測定値03 Den2
    MS03DEN3    As Integer              ' 測定値03 Den3
    MS03DEN4    As Integer              ' 測定値03 Den4
    MS03DEN5    As Integer              ' 測定値03 Den5
    MS04LDL1    As Integer              ' 測定値04 L/DL1
    MS04LDL2    As Integer              ' 測定値04 L/DL2
    MS04LDL3    As Integer              ' 測定値04 L/DL3
    MS04LDL4    As Integer              ' 測定値04 L/DL4
    MS04LDL5    As Integer              ' 測定値04 L/DL5
    MS04DEN1    As Integer              ' 測定値04 Den1
    MS04DEN2    As Integer              ' 測定値04 Den2
    MS04DEN3    As Integer              ' 測定値04 Den3
    MS04DEN4    As Integer              ' 測定値04 Den4
    MS04DEN5    As Integer              ' 測定値04 Den5
    MS05LDL1    As Integer              ' 測定値05 L/DL1
    MS05LDL2    As Integer              ' 測定値05 L/DL2
    MS05LDL3    As Integer              ' 測定値05 L/DL3
    MS05LDL4    As Integer              ' 測定値05 L/DL4
    MS05LDL5    As Integer              ' 測定値05 L/DL5
    MS05DEN1    As Integer              ' 測定値05 Den1
    MS05DEN2    As Integer              ' 測定値05 Den2
    MS05DEN3    As Integer              ' 測定値05 Den3
    MS05DEN4    As Integer              ' 測定値05 Den4
    MS05DEN5    As Integer              ' 測定値05 Den5
    MS06LDL1    As Integer              ' 測定値06 L/DL1
    MS06LDL2    As Integer              ' 測定値06 L/DL2
    MS06LDL3    As Integer              ' 測定値06 L/DL3
    MS06LDL4    As Integer              ' 測定値06 L/DL4
    MS06LDL5    As Integer              ' 測定値06 L/DL5
    MS06DEN1    As Integer              ' 測定値06 Den1
    MS06DEN2    As Integer              ' 測定値06 Den2
    MS06DEN3    As Integer              ' 測定値06 Den3
    MS06DEN4    As Integer              ' 測定値06 Den4
    MS06DEN5    As Integer              ' 測定値06 Den5
    MS07LDL1    As Integer              ' 測定値07 L/DL1
    MS07LDL2    As Integer              ' 測定値07 L/DL2
    MS07LDL3    As Integer              ' 測定値07 L/DL3
    MS07LDL4    As Integer              ' 測定値07 L/DL4
    MS07LDL5    As Integer              ' 測定値07 L/DL5
    MS07DEN1    As Integer              ' 測定値07 Den1
    MS07DEN2    As Integer              ' 測定値07 Den2
    MS07DEN3    As Integer              ' 測定値07 Den3
    MS07DEN4    As Integer              ' 測定値07 Den4
    MS07DEN5    As Integer              ' 測定値07 Den5
    MS08LDL1    As Integer              ' 測定値08 L/DL1
    MS08LDL2    As Integer              ' 測定値08 L/DL2
    MS08LDL3    As Integer              ' 測定値08 L/DL3
    MS08LDL4    As Integer              ' 測定値08 L/DL4
    MS08LDL5    As Integer              ' 測定値08 L/DL5
    MS08DEN1    As Integer              ' 測定値08 Den1
    MS08DEN2    As Integer              ' 測定値08 Den2
    MS08DEN3    As Integer              ' 測定値08 Den3
    MS08DEN4    As Integer              ' 測定値08 Den4
    MS08DEN5    As Integer              ' 測定値08 Den5
    MS09LDL1    As Integer              ' 測定値09 L/DL1
    MS09LDL2    As Integer              ' 測定値09 L/DL2
    MS09LDL3    As Integer              ' 測定値09 L/DL3
    MS09LDL4    As Integer              ' 測定値09 L/DL4
    MS09LDL5    As Integer              ' 測定値09 L/DL5
    MS09DEN1    As Integer              ' 測定値09 Den1
    MS09DEN2    As Integer              ' 測定値09 Den2
    MS09DEN3    As Integer              ' 測定値09 Den3
    MS09DEN4    As Integer              ' 測定値09 Den4
    MS09DEN5    As Integer              ' 測定値09 Den5
    MS10LDL1    As Integer              ' 測定値10 L/DL1
    MS10LDL2    As Integer              ' 測定値10 L/DL2
    MS10LDL3    As Integer              ' 測定値10 L/DL3
    MS10LDL4    As Integer              ' 測定値10 L/DL4
    MS10LDL5    As Integer              ' 測定値10 L/DL5
    MS10DEN1    As Integer              ' 測定値10 Den1
    MS10DEN2    As Integer              ' 測定値10 Den2
    MS10DEN3    As Integer              ' 測定値10 Den3
    MS10DEN4    As Integer              ' 測定値10 Den4
    MS10DEN5    As Integer              ' 測定値10 Den5
    MS11LDL1    As Integer              ' 測定値11 L/DL1
    MS11LDL2    As Integer              ' 測定値11 L/DL2
    MS11LDL3    As Integer              ' 測定値11 L/DL3
    MS11LDL4    As Integer              ' 測定値11 L/DL4
    MS11LDL5    As Integer              ' 測定値11 L/DL5
    MS11DEN1    As Integer              ' 測定値11 Den1
    MS11DEN2    As Integer              ' 測定値11 Den2
    MS11DEN3    As Integer              ' 測定値11 Den3
    MS11DEN4    As Integer              ' 測定値11 Den4
    MS11DEN5    As Integer              ' 測定値11 Den5
    MS12LDL1    As Integer              ' 測定値12 L/DL1
    MS12LDL2    As Integer              ' 測定値12 L/DL2
    MS12LDL3    As Integer              ' 測定値12 L/DL3
    MS12LDL4    As Integer              ' 測定値12 L/DL4
    MS12LDL5    As Integer              ' 測定値12 L/DL5
    MS12DEN1    As Integer              ' 測定値12 Den1
    MS12DEN2    As Integer              ' 測定値12 Den2
    MS12DEN3    As Integer              ' 測定値12 Den3
    MS12DEN4    As Integer              ' 測定値12 Den4
    MS12DEN5    As Integer              ' 測定値12 Den5
    MS13LDL1    As Integer              ' 測定値13 L/DL1
    MS13LDL2    As Integer              ' 測定値13 L/DL2
    MS13LDL3    As Integer              ' 測定値13 L/DL3
    MS13LDL4    As Integer              ' 測定値13 L/DL4
    MS13LDL5    As Integer              ' 測定値13 L/DL5
    MS13DEN1    As Integer              ' 測定値13 Den1
    MS13DEN2    As Integer              ' 測定値13 Den2
    MS13DEN3    As Integer              ' 測定値13 Den3
    MS13DEN4    As Integer              ' 測定値13 Den4
    MS13DEN5    As Integer              ' 測定値13 Den5
    MS14LDL1    As Integer              ' 測定値14 L/DL1
    MS14LDL2    As Integer              ' 測定値14 L/DL2
    MS14LDL3    As Integer              ' 測定値14 L/DL3
    MS14LDL4    As Integer              ' 測定値14 L/DL4
    MS14LDL5    As Integer              ' 測定値14 L/DL5
    MS14DEN1    As Integer              ' 測定値14 Den1
    MS14DEN2    As Integer              ' 測定値14 Den2
    MS14DEN3    As Integer              ' 測定値14 Den3
    MS14DEN4    As Integer              ' 測定値14 Den4
    MS14DEN5    As Integer              ' 測定値14 Den5
    MS15LDL1    As Integer              ' 測定値15 L/DL1
    MS15LDL2    As Integer              ' 測定値15 L/DL2
    MS15LDL3    As Integer              ' 測定値15 L/DL3
    MS15LDL4    As Integer              ' 測定値15 L/DL4
    MS15LDL5    As Integer              ' 測定値15 L/DL5
    MS15DEN1    As Integer              ' 測定値15 Den1
    MS15DEN2    As Integer              ' 測定値15 Den2
    MS15DEN3    As Integer              ' 測定値15 Den3
    MS15DEN4    As Integer              ' 測定値15 Den4
    MS15DEN5    As Integer              ' 測定値15 Den5
    MS01DVD2    As Integer              ' 測定値01 DVD2
    MS02DVD2    As Integer              ' 測定値02 DVD2
    MS03DVD2    As Integer              ' 測定値03 DVD2
    MS04DVD2    As Integer              ' 測定値04 DVD2
    MS05DVD2    As Integer              ' 測定値05 DVD2
    MSZEROMN    As Integer              ' L/DL0連続数最小値 '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
    MSZEROMX    As Integer              ' L/DL0連続数最大値 '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
    PTNJUDGRES  As String               ' パターン判定結果   '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
    REGDATE     As Date                 ' 登録日付
End Type

'ライフタイム実績取得関数
Public Type type_DBDRV_scmzc_fcmkc001c_LT
    CRYNUM      As String * 12          ' 結晶番号
    POSITION    As Integer              ' 位置
    SMPKBN      As String * 1           ' サンプル区分
    TRANCOND    As String * 1           ' 処理条件
    TRANCNT     As String * 1           ' 処理回数
    SMPLNO      As Long                 ' サンプルＮｏ  Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU     As String * 1           ' サンプル有無
    MEAS1       As Integer              ' 測定値１
    MEAS2       As Integer              ' 測定値２
    MEAS3       As Integer              ' 測定値３
    MEAS4       As Integer              ' 測定値４
    MEAS5       As Integer              ' 測定値５
    MEASPEAK    As Integer              ' 測定値 ピーク値
    CALCMEAS    As Integer              ' 計算結果
    REGDATE     As Date                 ' 登録日付
    LTSPI       As String               ' 測定位置コード
'2005/12/02 add SET高崎 測定値6～10追加、判定フラグ->
    MEAS6       As Integer              ' 測定値６
    MEAS7       As Integer              ' 測定値７
    MEAS8       As Integer              ' 測定値８
    MEAS9       As Integer              ' 測定値９
    MEAS10      As Integer              ' 測定値１０
    LTSPIFLG    As String               ' 判定フラグ
'2005/12/02 add SET高崎 測定値6～10追加、判定フラグ<-
''Add Start 2011/07/22 LT10Ω判定追加対応 T.Koi(SETsw)
    CONVAL      As Integer               ' LT10Ω換算
''Add End   2011/07/22 LT10Ω判定追加対応 T.Koi(SETsw)

End Type


'EPD実績取得関数
Public Type type_DBDRV_scmzc_fcmkc001c_EPD
    CRYNUM      As String * 12          ' 結晶番号
    POSITION    As Integer              ' 位置
    SMPKBN      As String * 1           ' サンプル区分
    TRANCOND    As String * 1           ' 処理条件
    TRANCNT     As String * 1           ' 処理回数
    SMPLNO      As Long                 ' サンプルＮｏ  Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPLUMU     As String * 1           ' サンプル有無
    MEASURE     As Integer              ' 測定値
    REGDATE     As Date                 ' 登録日付
End Type
'X線実績取得関数   2009/08/12 add Kameda
Public Type type_DBDRV_scmzc_fcmkc001c_X
    CRYNUM      As String * 12          ' 結晶番号
    POSITION    As Integer              ' 位置
    SMPKBN      As String * 1           ' サンプル区分
    TRANCOND    As String * 1           ' 処理条件
    TRANCNT     As String * 1           ' 処理回数
    SMPLNO      As Long                 ' サンプルＮｏ
    SMPLUMU     As String * 1           ' サンプル有無
    XX          As Double               ' 測定値X
    XY          As Double               ' 測定値Y
    XXY         As Double               ' 測定値合成
    REGDATE     As Date                 ' 登録日付
    'JUDG        As String              ' 判定項目追加   2009/10/22 Kameda
    JUDGXY       As String              ' 判定項目追加   2009/10/22 Kameda
    JUDGX        As String              ' 判定項目追加   2009/10/22 Kameda
    JUDGY        As String              ' 判定項目追加   2009/10/22 Kameda
End Type
'SIRD実績取得関数   2010/02/04 add Kameda
Public Type type_DBDRV_scmzc_fcmkc001c_SIRD
    CRYNUM      As String * 12          ' 結晶番号
    POSITION    As Integer              ' 位置
    SMPKBN      As String * 1           ' サンプル区分
    TRANCOND    As String * 1           ' 処理条件
    TRANCNT     As String * 1           ' 処理回数
    'SMPLNO      As Long                 ' サンプルＮｏ
    SMPLNO      As String                ' サンプルＮｏ
    SMPLUMU     As String * 1           ' サンプル有無
    SIRDCNT     As Double               ' 評価結果
    REGDATE     As Date                 ' 登録日付
    NothingFlg  As String               ' データなしフラグ   2010/02/18 Kameda
End Type

'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : Cu-deco(C, CJ, CJ(LT), CJ2)の項目追加
Public Type type_DBDRV_scmzc_fcmkc001c_C
    CRYNUM          As String * 12      ' 結晶番号
    POSITION        As Integer          ' 位置
    SMPKBN          As String * 1       ' サンプル区分
    TRANCNT         As Integer          ' 処理回数
'    TRANCOND        As String * 1       '処理条件
    SMPLNO          As Long             ' サンプルＮｏ
    SMPLUMUC        As String * 1       ' サンプル有無（C）
    
    CPTNJSK         As String * 1       ' C パターン実績
    CDISKJSK        As Integer          ' C Disk半径実績
    CRINGNKJSK      As Integer          ' C Ring内径実績
    CRINGGKJSK      As Integer          ' C Ring外径実績
    CHANTEI         As String * 1       ' C 判定結果
    REGDATE         As Date             ' C 登録日付
End Type

Public Type type_DBDRV_scmzc_fcmkc001c_CJ
    CRYNUM          As String * 12      ' 結晶番号
    POSITION        As Integer          ' 位置
    SMPKBN          As String * 1       ' サンプル区分
    TRANCNT         As Integer          ' 処理回数
'    TRANCOND        As String * 1       '処理条件
    SMPLNO          As Long             ' サンプルＮｏ
    SMPLUMUCJ       As String * 1       ' サンプル有無（CJ）
    
    CJPTNJSK        As String * 1       ' CJ パターン実績
    CJDISKJSK       As Integer          ' CJ Disk半径実績
    CJRINGNKJSK     As Integer          ' CJ Ring内径実績
    CJRINGGKJSK     As Integer          ' CJ Ring外径実績
    CJBANDNKJSK     As Integer          ' CJ Band内径実績
    CJBANDGKJSK     As Integer          ' CJ Band外径実績
    CJRINGCALC      As Integer          ' CJ Ring幅計算
    CJPICALC        As Integer          ' CJ Pi幅計算
    CJHANTEI        As String * 1       ' CJ 判定結果
'    CJBUIUMU        As String * 1       ' CJ 部位別判定有無
    CJDMAXPIC5      As Integer          ' CJ Diskのみパターン Pi幅上限値
    CJRMAXPIC5      As Integer          ' CJ Ringのみパターン Pi幅上限値
    CJDRMAXPIC5     As Integer          ' CJ DiskRingパターン Pi幅上限値
    CJALLMAXDIC5    As Integer          ' CJ 共通Disk半径上限値
    CJALLMINRINC5   As Integer          ' CJ 共通Ring内径下限値
    CJALLMAXRIGC5   As Integer          ' CJ 共通Ring外径上限値
    REGDATE         As Date             ' CJ 登録日付
End Type

Public Type type_DBDRV_scmzc_fcmkc001c_CJLT
    CRYNUM          As String * 12      ' 結晶番号
    POSITION        As Integer          ' 位置
    SMPKBN          As String * 1       ' サンプル区分
    TRANCNT         As Integer          ' 処理回数
'    TRANCOND        As String * 1       '処理条件
    SMPLNO          As Long             ' サンプルＮｏ
    SMPLUMUCJLT     As String * 1       ' サンプル有無（CJ(LT)）
    
    CJLTPTNJSK      As String * 1       ' CJ(LT) パターン実績
    CJLTDISKJSK     As Integer          ' CJ(LT) Disk半径実績
    CJLTRINGNKJSK   As Integer          ' CJ(LT) Ring内径実績
    CJLTRINGGKJSK   As Integer          ' CJ(LT) Ring外径実績
    CJLTBANDNKJSK   As Integer          ' CJ(LT) Band内径実績
    CJLTBANDGKJSK   As Integer          ' CJ(LT) Band外径実績
    CJLTRINGCALC    As Integer          ' CJ(LT) Ring幅計算
    CJLTPICALC      As Integer          ' CJ(LT) Pi幅計算
    CJLTBANDCALC    As Integer          ' CJ(LT) Band幅計算
    CJLTHANTEI      As String * 1       ' CJ(LT) 判定結果
    HSXCJLTBND      As Integer          ' CJ(LT) Band幅上限値
    REGDATE         As Date             ' CJ(LT) 登録日付
End Type

Public Type type_DBDRV_scmzc_fcmkc001c_CJ2
    CRYNUM          As String * 12      ' 結晶番号
    POSITION        As Integer          ' 位置
    SMPKBN          As String * 1       ' サンプル区分
    TRANCNT         As Integer          ' 処理回数
'    TRANCOND        As String * 1       '処理条件
    SMPLNO          As Long             ' サンプルＮｏ
    SMPLUMUCJ2      As String * 1       ' サンプル有無（CJ2）
    
    CJ2PTNJSK       As String * 1       ' CJ2 パターン実績
    CJ2DISKJSK      As Integer          ' CJ2 Disk半径実績
    CJ2RINGNKJSK    As Integer          ' CJ2 Ring内径実績
    CJ2RINGGKJSK    As Integer          ' CJ2 Ring外径実績
    CJ2PICALC       As Integer          ' CJ2 Pi幅計算
    CJ2HANTEI       As String * 1       ' CJ2 判定結果
'    CJ2BUIUMU       As String * 1       ' CJ2 部位別判定有無
    CJ2DMAXPIC5     As Integer          ' CJ2 Diskのみパターン Pi幅下限値(MAXだが下限です)
    CJ2RMAXPIC5     As Integer          ' CJ2 Ringのみパターン Pi幅下限値(MAXだが下限です)
    CJ2RMINRINC5    As Integer          ' CJ2 Ringのみパターン Ring内径下限値
    CJ2RMAXRIGC5    As Integer          ' CJ2 Ringのみパターン Ring外径上限値
    CJ2DRMAXPIC5    As Integer          ' CJ2 DiskRingパターン Pi幅下限値(MAXだが下限です)
    CJ2DRMINRINC5   As Integer          ' CJ2 DiskRingパターン Ring内径下限値
    CJ2DRMAXRIGC5   As Integer          ' CJ2 DiskRingパターン Ring外径上限値
    REGDATE         As Date             ' CJ2 登録日付
End Type
'Add End   2011/01/17 SMPK A.Nagamine



'------------------------------------------------
' 結晶反映/推定チェック共通関数
'------------------------------------------------

'概要      :指定された評価項目№により、反映か推定かを判断し、結晶反映チェック、または、結晶推定チェックを呼び出す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :sBlockid      ,I  ,String       :ﾌﾞﾛｯｸID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :結晶番号
'          :tFullHin      ,I  ,tFullHinban  :品番(構造体)
'          :iSmplPos      ,I  ,Integer      :新サンプル位置(mm)
'          :iItemNo       ,I  ,Integer      :評価項目№ =  1 RS     ← 推定偏析計算
'                                                       =  2 Oi     ← ﾊﾟﾀｰﾝ1
'                                                       =  3 BMD1   ← ﾊﾟﾀｰﾝ1
'                                                       =  4 BMD2   ← ﾊﾟﾀｰﾝ1
'                                                       =  5 BMD3   ← ﾊﾟﾀｰﾝ1
'                                                       =  6 OSF1   ← ﾊﾟﾀｰﾝ1
'                                                       =  7 OSF2   ← ﾊﾟﾀｰﾝ1
'                                                       =  8 OSF3   ← ﾊﾟﾀｰﾝ1
'                                                       =  9 OSF4   ← ﾊﾟﾀｰﾝ1
'                                                       = 10 CS     ← ﾊﾟﾀｰﾝ2(上限値,下限値共0より大(0<)の場合,ﾊﾟﾀｰﾝ1)
'                                                       = 11 GD     ← ﾊﾟﾀｰﾝ1
'                                                       = 12 LT     ← ﾊﾟﾀｰﾝ3
'                                                       = 13 EPD    ← ﾊﾟﾀｰﾝ2
'          :iFromPos      ,I  ,Integer      :検索範囲From
'          :iToPos        ,I  ,Integer      :検索範囲To
'          :iHanSuiKBN    ,O  ,Integer      :反映/推定区分(0:反映,1:推定)
'          :iGetSmplID1   ,O  ,long         :元サンプルID1                  Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :iGetSmplID2   ,O  ,long         :元サンプルID2 (反映時未使用)   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :戻り値        ,O  ,Integer      :チェック結果 = 0 : 正常終了(反映/推定OK)
'                                                           1 : 正常終了(反映/推定NG)
'                                                          -1 : 入力引数値エラー
'                                                          -2 : 上記以外のエラー
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funChkSxlHanSui(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                                iItemNo As Integer, iFromPos As Integer, iToPos As Integer, iHanSuiKBN As Integer, _
                                iGetSmplID1 As Long, iGetSmplID2 As Long, tFullhin2 As tFullHinban) As Integer
    Dim retCode As Integer
    
    '元サンプルID初期化
    iGetSmplID1 = -1
    iGetSmplID2 = -1
    
    'パラメータチェック
    If (Len(sBlockId) <> 12) Then GoTo ChkSxlHanSuiParameterErr
    If (Len(sCryNum) <> 12) Then GoTo ChkSxlHanSuiParameterErr
    
    '指定された評価項目№により、反映か推定かを判断し、結晶反映チェック、または、結晶推定チェックを呼び出す。
    Select Case iItemNo
    Case 1              'RS(比抵抗)
        retCode = funChkSxlSuitei(sBlockId, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, iGetSmplID1, iGetSmplID2, tFullhin2)
        iHanSuiKBN = 1
    'Chg Start 2011/01/31 SMPK Miyata
    'Case 2 To 13        'Oi(酸素濃度),BMD1,BMD2,BMD3,OSF1,OSF2,OSF3,OSF4,CS(炭素濃度),GD,LT(ﾗｲﾌﾀｲﾑ),EPD
    Case 2 To 13, 15 To 18  'Oi(酸素濃度),BMD1,BMD2,BMD3,OSF1,OSF2,OSF3,OSF4,CS(炭素濃度),GD,LT(ﾗｲﾌﾀｲﾑ),EPD,C,CJ,CJLT,CJ2
    'Chg End   2011/01/31 SMPK Miyata
        retCode = funChkSxlHanei(sBlockId, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, iFromPos, iToPos, iGetSmplID1)
        iHanSuiKBN = 0
    Case Else
        GoTo ChkSxlHanSuiParameterErr
    End Select
    
    '共通関数のチェック結果を当関数の結果として、呼び出し元へ返す。
    funChkSxlHanSui = retCode
    Exit Function

ChkSxlHanSuiParameterErr:
    funChkSxlHanSui = -1
    Exit Function

ChkSxlHanSuiSonotaErr:
    funChkSxlHanSui = -2
End Function

'------------------------------------------------
' 結晶反映チェック
'------------------------------------------------

'概要      :指定された情報から、結晶反映チェックを行ない結果を返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :sBlockid      ,I  ,String       :ﾌﾞﾛｯｸID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :結晶番号
'          :tFullHin      ,I  ,tFullHinban  :品番(構造体)
'          :iSmplPos      ,I  ,Integer      :新サンプル位置(mm)
'          :iItemNo       ,I  ,Integer      :評価項目№ =  1 RS
'                                                       =  2 Oi     ←対象
'                                                       =  3 BMD1   ←対象
'                                                       =  4 BMD2   ←対象
'                                                       =  5 BMD3   ←対象
'                                                       =  6 OSF1   ←対象
'                                                       =  7 OSF2   ←対象
'                                                       =  8 OSF3   ←対象
'                                                       =  9 OSF4   ←対象
'                                                       = 10 CS     ←対象
'                                                       = 11 GD     ←対象
'                                                       = 12 LT     ←対象
'                                                       = 13 EPD    ←対象
'          :iFromPos      ,I  ,Integer      :検索範囲From
'          :iToPos        ,I  ,Integer      :検索範囲To
'          :iGetSmplID    ,O  ,long         :反映元サンプルID   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :戻り値        ,O  ,Integer      :チェック結果 = 0 : 正常終了(反映OK)
'                                                           1 : 正常終了(反映NG)
'                                                          -1 : 入力引数値エラー
'                                                          -2 : 上記以外のエラー
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funChkSxlHanei(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                               iItemNo As Integer, iFromPos As Integer, iToPos As Integer, iGetSmplID As Long) As Integer
    Dim wHPtrn          As Integer
    Dim tSiyou          As type_DBDRV_scmzc_fcmkc001c_Siyou
'C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
    Dim strOSFRD1       As String
    Dim strOSFRD2       As String
    Dim lngOSFWID1      As Long
    Dim lngOSFWID2      As Long
    Dim strJDGEIDC      As String
    Dim strSynFlagc5    As String
    Dim strYmkFlagc5    As String
    Dim lSmpPos         As Long
    Dim strRMAXC5       As String
    Dim strDMAXC5       As String
    Dim strDRRMAXC5     As String
    Dim strDRDMAXC5     As String
    Dim lRMaxc5         As Long
    Dim lDMaxc5         As Long
    Dim lDrrMaxc5       As Long
    Dim lDrdMaxc5       As Long
    
    Dim tSiyou2         As type_DBDRV_scmzc_fcmkc001c_Siyou
'C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---
    Dim wGetBlockid     As String
    Dim wGetSmpKbn      As String
    Dim wGetSmplID      As Long     'Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    
    Dim tCryOi          As type_DBDRV_scmzc_fcmkc001c_Oi
    Dim tCryBMD         As type_DBDRV_scmzc_fcmkc001c_BMD
    Dim tCryOSF         As type_DBDRV_scmzc_fcmkc001c_OSF
    Dim tCryCS          As type_DBDRV_scmzc_fcmkc001c_CS
    Dim tCryGD          As type_DBDRV_scmzc_fcmkc001c_GD
    Dim tCryLT          As type_DBDRV_scmzc_fcmkc001c_LT
    Dim tCryEPD         As type_DBDRV_scmzc_fcmkc001c_EPD
    'Add Start 2011/01/31 SMPK Miyata
    Dim tCryC           As type_DBDRV_scmzc_fcmkc001c_C
    Dim tCryCJ          As type_DBDRV_scmzc_fcmkc001c_CJ
    Dim tCryCJLT        As type_DBDRV_scmzc_fcmkc001c_CJLT
    Dim tCryCJ2         As type_DBDRV_scmzc_fcmkc001c_CJ2
    'Add End   2011/01/31 SMPK Miyata

    Dim retJudg         As Boolean
    Dim wIdFlg          As Integer
    
    Dim dShiyo()        As Double       '2003/12/11 Null対応追加
    Dim sHosyo          As String       '2003/12/11 Null対応追加
    
    '初期化
    wGetSmplID = -1
    
    'パラメータチェック
    If (Len(sBlockId) <> 12) Then GoTo ChkSxlHaneiParameterErr
    If (Len(sCryNum) <> 12) Then GoTo ChkSxlHaneiParameterErr
    
    '指定された評価項目№毎に必要な品番仕様値を取得し、結晶反映値取得パターンを決定する。（指定された評価項目№により、処理が分かれる。）
    Select Case iItemNo
    Case 1                      'RS(比抵抗)
        GoTo ChkSxlHaneiNG
    Case 2                      'Oi(酸素濃度)
        If funGet_TBCME019(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
        wHPtrn = 1
        
        'Null対応処理追加 2003/12/11 SystenBrain ▽
        ReDim dShiyo(5)
        dShiyo(1) = tSiyou.HSXONMIN         ' 品ＳＸ酸素濃度下限
        dShiyo(2) = tSiyou.HSXONMAX         ' 品ＳＸ酸素濃度上限
        dShiyo(3) = tSiyou.HSXONAMN         ' 品ＳＸ酸素濃度平均下限
        dShiyo(4) = tSiyou.HSXONAMX         ' 品ＳＸ酸素濃度平均上限
        dShiyo(5) = tSiyou.HSXONMBP         ' 品ＳＸ酸素濃度面内分布
        'NULLは不問(NULLﾁｪｯｸ関数ｺﾒﾝﾄｱｳﾄ) 09/03/13 ooba
'        If fncJissekiHantei_nl(tSiyou.HSXONHWS, dShiyo) = False Then GoTo ChkSxlHaneiSonotaErr
        'Null対応処理追加 2003/12/11 SystenBrain △
        
    Case 3, 4, 5                'BMD1,BMD2,BMD3
        If funGet_TBCME020(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
        wHPtrn = 1
                
    Case 6, 7, 8, 9             'OSF1,OSF2,OSF3,OSF4
        If funGet_TBCME020(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
        wHPtrn = 1
        
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
        If funGet_TBCME036(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
            
        'Null対応処理追加 2003/12/11 SystenBrain ▽
        ReDim dShiyo(2)
        If iItemNo = 6 Then         'OSF1
            sHosyo = tSiyou.HSXOF1HS            ' 品ＳＸＯＳＦ1保証方法＿処
            dShiyo(1) = tSiyou.HSXOF1AX         ' 品ＳＸＯＳＦ1平均上限
            dShiyo(2) = tSiyou.HSXOF1MX         ' 品ＳＸＯＳＦ1上限
        ElseIf iItemNo = 7 Then     'OSF2
            sHosyo = tSiyou.HSXOF2HS            ' 品ＳＸＯＳＦ2保証方法＿処
            dShiyo(1) = tSiyou.HSXOF2AX         ' 品ＳＸＯＳＦ2平均上限
            dShiyo(2) = tSiyou.HSXOF2MX         ' 品ＳＸＯＳＦ2上限
        ElseIf iItemNo = 8 Then     'OSF3
            sHosyo = tSiyou.HSXOF3HS            ' 品ＳＸＯＳＦ3保証方法＿処
            dShiyo(1) = tSiyou.HSXOF3AX         ' 品ＳＸＯＳＦ3平均上限
            dShiyo(2) = tSiyou.HSXOF3MX         ' 品ＳＸＯＳＦ3上限
        ElseIf iItemNo = 9 Then     'OSF4
'C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
            'C-OSF3ﾌﾗｸﾞ獲得
            If funGet_TBCME036(tFullHin, tSiyou2) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
            wHPtrn = 1
            'ﾌﾗｸﾞの見先を変更
            sHosyo = tSiyou2.COSF3FLAG          ' Ｃ－ＯＳＦ３ﾌﾗｸﾞ
'C－OSF3判定機能追加 2007/04/23 M.Kaga END  ---
            dShiyo(1) = tSiyou.HSXOF4AX         ' 品ＳＸＯＳＦ4平均上限
            dShiyo(2) = tSiyou.HSXOF4MX         ' 品ＳＸＯＳＦ4上限
        End If
        
'C－OSF3判定機能追加 2007/06/14 M.Kaga STRAT ---
        If iItemNo <> 9 Then
            'NULLは不問(NULLﾁｪｯｸ関数ｺﾒﾝﾄｱｳﾄ) 09/03/13 ooba
'            If fncJissekiHantei_nl(sHosyo, dShiyo) = False Then GoTo ChkSxlHaneiSonotaErr
            'Null対応処理追加 2003/12/11 SystenBrain △
        End If
'C－OSF3判定機能追加 2007/06/14 M.Kaga END ---

    Case 10                     'CS(炭素濃度)
        If funGet_TBCME019(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
        'TOP/BOT保証は反映ﾊﾟﾀｰﾝ1,BOT保証は反映ﾊﾟﾀｰﾝ2 09/01/08 ooba
        If tSiyou.HSXCNKHI = "6" Or tSiyou.HSXCNKHI = "9" Then
            wHPtrn = 1
        Else
            wHPtrn = 2
        End If
        
        'Null対応処理追加 2003/12/11 SystenBrain ▽
        ReDim dShiyo(2)
        dShiyo(1) = tSiyou.HSXCNMIN         ' 品ＳＸ炭素濃度下限
        dShiyo(2) = tSiyou.HSXCNMAX         ' 品ＳＸ炭素濃度上限
        'NULLは不問(NULLﾁｪｯｸ関数ｺﾒﾝﾄｱｳﾄ) 09/03/13 ooba
'        If fncJissekiHantei_nl(tSiyou.HSXCNHWS, dShiyo) = False Then GoTo ChkSxlHaneiSonotaErr
        'Null対応処理追加 2003/12/11 SystenBrain △
        
    Case 11                     'GD
        If funGet_TBCME020(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
    '*** UPDATE ↓ Y.SIMIZU 2005/10/12 GDﾗｲﾝ取得追加
        If funGet_TBCME036(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
    '*** UPDATE ↑ Y.SIMIZU 2005/10/12 GDﾗｲﾝ取得追加
        wHPtrn = 1
                
    Case 12                     'LT(ﾗｲﾌﾀｲﾑ)
        If funGet_TBCME019(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
        If funGet_TBCME036(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
        wHPtrn = 3
    
        'Null対応処理追加 2003/12/11 SystenBrain ▽
        ReDim dShiyo(2)
        dShiyo(1) = tSiyou.HSXLTMIN         ' 品ＳＸＬタイム下限
        dShiyo(2) = tSiyou.HSXLTMAX         ' 品ＳＸＬタイム上限
        'NULLは不問(NULLﾁｪｯｸ関数ｺﾒﾝﾄｱｳﾄ) 09/03/13 ooba
'        If fncJissekiHantei_nl(tSiyou.HSXLTHWS, dShiyo) = False Then GoTo ChkSxlHaneiSonotaErr
        'Null対応処理追加 2003/12/11 SystenBrain △
        
    Case 13                     'EPD
        If funGet_TBCME036(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
        wHPtrn = 2
    
        'Null対応処理追加 2003/12/11 SystenBrain ▽
        If tSiyou.EPDUP = -1 Then GoTo ChkSxlHaneiSonotaErr     ' EPD上限
        'Null対応処理追加 2003/12/11 SystenBrain △

    'Add Start 2011/01/31 SMPK Miyata
    Case 15, 16, 17, 18         'C,CJ,CJLT,CJ2
        If funGet_TBCME020(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
        wHPtrn = 1
        
        If iItemNo = 17 Then
            If funGet_TBCME036(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlHaneiNG
        End If
    'Add End   2011/01/31 SMPK Miyata

    Case Else
        GoTo ChkSxlHaneiParameterErr
    End Select

    '結晶反映元サンプルＩＤの取得
    If wHPtrn = 1 Then              '結晶反映値取得パターン１
        If funGetSxlHanei1(sBlockId, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, iFromPos, iToPos, _
                                                                    wGetBlockid, wGetSmpKbn, wGetSmplID) <> 0 Then GoTo ChkSxlHaneiNG
    
    ElseIf wHPtrn = 2 Then          '結晶反映値取得パターン２
        If funGetSxlHanei2(sBlockId, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, _
                                                                    wGetBlockid, wGetSmpKbn, wGetSmplID) <> 0 Then GoTo ChkSxlHaneiNG
    
    ElseIf wHPtrn = 3 Then          '結晶反映値取得パターン３
        If funGetSxlHanei3(sBlockId, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, tSiyou.HSXLTSPI, _
                                                                    wGetBlockid, wGetSmpKbn, wGetSmplID) <> 0 Then GoTo ChkSxlHaneiNG
    
    End If
    
    '結晶反映元ｻﾝﾌﾟﾙIDから、結晶反映値（実績値）を取得する。（指定された評価項目№により、処理が分かれる。）
    Select Case iItemNo
    Case 2                      'Oi(酸素濃度)
        'Oiの実績値を取得する
        If funGetCryOiJisseki(sCryNum, wGetSmplID, tCryOi) <> 0 Then GoTo ChkSxlHaneiNG
        'Oi総合判定を行なう
        If Not CrOiJudg(tSiyou, tCryOi, retJudg) Then GoTo ChkSxlHaneiNG
    
    Case 3, 4, 5                'BMD1, BMD2, BMD3
        If iItemNo = 3 Then
            'BMD1の実績値を取得する
            If funGetCryBMDJisseki(sCryNum, wGetSmplID, 1, tCryBMD) <> 0 Then GoTo ChkSxlHaneiNG
            wIdFlg = 1
        ElseIf iItemNo = 4 Then
            'BMD2の実績値を取得する
            If funGetCryBMDJisseki(sCryNum, wGetSmplID, 2, tCryBMD) <> 0 Then GoTo ChkSxlHaneiNG
            wIdFlg = 2
        ElseIf iItemNo = 5 Then
            'BMD3の実績値を取得する
            If funGetCryBMDJisseki(sCryNum, wGetSmplID, 3, tCryBMD) <> 0 Then GoTo ChkSxlHaneiNG
            wIdFlg = 3
        End If
        'BMDの総合判定を行なう
        If Not CrBmdJudg(tSiyou, tCryBMD, retJudg, wIdFlg) Then GoTo ChkSxlHaneiNG
    
    Case 6, 7, 8, 9             'OSF1, OSF2, OSF3, OSF4
        If iItemNo = 6 Then
            'OSF1の実績値を取得する
            If funGetCryOSFJisseki(sCryNum, wGetSmplID, 1, tCryOSF) <> 0 Then GoTo ChkSxlHaneiNG
            wIdFlg = 1
        ElseIf iItemNo = 7 Then
            'OSF2の実績値を取得する
            If funGetCryOSFJisseki(sCryNum, wGetSmplID, 2, tCryOSF) <> 0 Then GoTo ChkSxlHaneiNG
            wIdFlg = 2
        ElseIf iItemNo = 8 Then
            'OSF3の実績値を取得する
            If funGetCryOSFJisseki(sCryNum, wGetSmplID, 3, tCryOSF) <> 0 Then GoTo ChkSxlHaneiNG
            wIdFlg = 3
        ElseIf iItemNo = 9 Then
            'OSF4の実績値を取得する
            If funGetCryOSFJisseki(sCryNum, wGetSmplID, 4, tCryOSF) <> 0 Then GoTo ChkSxlHaneiNG
            wIdFlg = 4
        End If
        
'C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
        'OSF1,2,3の場合
        If iItemNo = 6 Or iItemNo = 7 Or iItemNo = 8 Then
            'OSFの総合判定を行なう
            If Not CrOsfJudg(tSiyou, tCryOSF, retJudg, wIdFlg) Then GoTo ChkSxlHaneiNG
        Else
            'OSF実績入力判定処理
            '判定ﾊﾟﾀｰﾝ&実績値退避
            If Trim(tCryOSF.OSFRD1) = "R" Or Trim(tCryOSF.OSFRD1) = "D" Then
                strOSFRD1 = Trim(tCryOSF.OSFRD1)
            Else
                strOSFRD1 = "-"
            End If
            If Trim(tCryOSF.OSFRD2) = "D" Then
                strOSFRD2 = Trim(tCryOSF.OSFRD2)
            Else
                strOSFRD2 = "-"
            End If
            If IsNull(tCryOSF.OSFWID1) = True Then
               lngOSFWID1 = -1
            ElseIf IsNumeric(tCryOSF.OSFWID1) = False Then
               lngOSFWID1 = -1
            Else
               lngOSFWID1 = Trim(tCryOSF.OSFWID1)
            End If
            If IsNull(tCryOSF.OSFWID2) = True Then
               lngOSFWID2 = -1
            ElseIf IsNumeric(tCryOSF.OSFWID2) = False Then
               lngOSFWID2 = -1
            Else
               lngOSFWID2 = Trim(tCryOSF.OSFWID2)
            End If
            
            '-1以外の数値考慮
            If lngOSFWID1 < 0 Then
               lngOSFWID1 = -1
            End If
            If lngOSFWID2 < 0 Then
               lngOSFWID2 = -1
            End If

            'ﾊﾟﾀｰﾝ区分、実績値がNULLの場合
            If strOSFRD1 = "-" And strOSFRD2 <> "-" Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            ElseIf strOSFRD1 <> "-" And lngOSFWID1 = -1 Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            ElseIf strOSFRD2 <> "-" And lngOSFWID2 = -1 Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            ElseIf strOSFRD2 = "-" And lngOSFWID2 > 0 Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
        
             '結晶番号をｷｰとしてXSDC1よりC－OSF3判定IDを獲得する
            If GetCOSF3ID(strJDGEIDC, Trim(sCryNum)) <> FUNCTION_RETURN_SUCCESS Then
                '該当ﾚｺｰﾄﾞ無しの場合
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
            'C－OSF3判定IDがNULLの場合
            If Trim(strJDGEIDC) = "" Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            Else
                '獲得したC-OSF3判定IDでXODC5_OSF30より承認ﾌﾗｸﾞの獲得
                If GetSYNFLAGC5(strSynFlagc5, strYmkFlagc5, Trim(strJDGEIDC)) <> FUNCTION_RETURN_SUCCESS Then
                    '該当ﾚｺｰﾄﾞ無しの場合
                    retJudg = False
                    GoTo ChkSxlHaneiNG
                End If
                '承認ﾌﾗｸﾞ:0　未承認の場合
                If Trim(strSynFlagc5) = "0" Or Trim(strSynFlagc5) = "" Or IsNull(strSynFlagc5) Then
                    retJudg = False
                    GoTo ChkSxlHaneiNG
                '削除ﾌﾗｸﾞ:1　無効の場合
                ElseIf Trim(strYmkFlagc5) = "1" Or Trim(strYmkFlagc5) = "" Or IsNull(strYmkFlagc5) Then
                    retJudg = False
                    GoTo ChkSxlHaneiNG
                Else
                
                    'ｻﾝﾌﾟﾙ位置取得(反映元ｻﾝﾌﾟﾙNOに紐付く位置)
                    lSmpPos = Trim(iSmplPos)
                    
                    'ﾊﾟﾀｰﾝ区分により処理分岐
                    'Rのみの場合
                    If strOSFRD1 = "R" And strOSFRD2 = "-" Then
                        'Rのみ上限値の獲得を行う
                        If GetCOSF3PTN(strJDGEIDC, lSmpPos, strOSFRD1, strOSFRD2, strRMAXC5, strDMAXC5, strDRRMAXC5, strDRDMAXC5) <> FUNCTION_RETURN_SUCCESS Then
                            '該当ﾚｺｰﾄﾞ無しの場合
                            retJudg = False
                            GoTo ChkSxlHaneiNG
                        End If
                        'ﾚｺｰﾄﾞ無：VBエラー(後で考える)
                        If Trim(strRMAXC5) = "" Then
                            retJudg = False
                            GoTo ChkSxlHaneiNG
                        Else
                            lRMaxc5 = Trim(strRMAXC5)
                            '実績値の判定
                            If lngOSFWID1 <= lRMaxc5 Then
                                '反映OK
                                retJudg = True
                            ElseIf lngOSFWID1 > lRMaxc5 Then
                                '反映NG
                                retJudg = False
                            End If
                        End If
                    'Dのみの場合
                    ElseIf strOSFRD1 = "D" Then
                        'Dのみ上限値の獲得を行う
                        If GetCOSF3PTN(strJDGEIDC, lSmpPos, strOSFRD1, strOSFRD2, strRMAXC5, strDMAXC5, strDRRMAXC5, strDRDMAXC5) <> FUNCTION_RETURN_SUCCESS Then
                            '該当ﾚｺｰﾄﾞ無しの場合
                            retJudg = False
                            GoTo ChkSxlHaneiNG
                        End If
                                                
                        'ﾚｺｰﾄﾞ無又はﾏｽﾀの実績値がNULL：VBエラー(後で考える)
                        If Trim(strDMAXC5) = "" Then
                            retJudg = False
                            GoTo ChkSxlHaneiNG
                        Else
                            lDMaxc5 = Trim(strDMAXC5)
                            '実績値の判定
                            If lngOSFWID1 <= lDMaxc5 Then
                                '反映OK
                                retJudg = True
                            ElseIf lngOSFWID1 > lDMaxc5 Then
                                '反映NG
                                retJudg = False
                            End If
                        End If
                    'R&Dの場合
                    ElseIf strOSFRD1 = "R" And strOSFRD2 = "D" Then
                        'D共存上限値並びR共存上限値の獲得を行う
                        If GetCOSF3PTN(strJDGEIDC, lSmpPos, strOSFRD1, strOSFRD2, strRMAXC5, strDMAXC5, strDRRMAXC5, strDRDMAXC5) <> FUNCTION_RETURN_SUCCESS Then
                         '該当ﾚｺｰﾄﾞ無しの場合
                            retJudg = False
                            GoTo ChkSxlHaneiNG
                        End If
                                                
                        'ﾚｺｰﾄﾞ無又はﾏｽﾀの実績値がNULL：VBエラー(後で考える)
                        If Trim(strDRRMAXC5) = "" Or Trim(strDRDMAXC5) = "" Then
                            retJudg = False
                            GoTo ChkSxlHaneiNG
                        Else
                            lDrrMaxc5 = Trim(strDRRMAXC5)
                            lDrdMaxc5 = Trim(strDRDMAXC5)
                            '実績値の判定
                            If lngOSFWID1 <= lDrrMaxc5 And lngOSFWID2 <= lDrdMaxc5 Then
                                '反映OK
                                retJudg = True
                            ElseIf lngOSFWID1 > lDrrMaxc5 Or lngOSFWID2 > lDrdMaxc5 Then
                                '反映NG
                                retJudg = False
                            End If
                        End If
                    Else
                        '実績値無、ﾊﾟﾀｰﾝ区分無の場合反映OK
                         retJudg = True
                    End If
                End If
            End If
        End If
        
'C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---
    
    Case 10             'CS(炭素濃度)
        'CSの実績値を取得する
        If funGetCryCSJisseki(sCryNum, wGetSmplID, tCryCS) <> 0 Then GoTo ChkSxlHaneiNG
        'CS総合判定を行なう
        If Not CrCsjudg(tSiyou, tCryCS, retJudg) Then GoTo ChkSxlHaneiNG
    
    Case 11             'GD
        'GDの実績値を取得する
        If funGetCryGDJisseki(sCryNum, wGetSmplID, tCryGD) <> 0 Then GoTo ChkSxlHaneiNG
        'GD総合判定を行なう
        If Not CrGdjudg(tSiyou, tCryGD, retJudg) Then GoTo ChkSxlHaneiNG
    
    Case 12             'LT(ﾗｲﾌﾀｲﾑ)
        'LTの実績値を取得する
        If funGetCryLTJisseki(sCryNum, wGetSmplID, tCryLT) <> 0 Then GoTo ChkSxlHaneiNG
        '2005/12/02 add SET高崎 LT計算関数をcallする ->
        'ライフタイム値を計算しなおす
        Call Sub_LTReCalc(tSiyou, tCryLT)
        '2005/12/02 add SET高崎 LT計算関数をcallする <-
        'LT総合判定を行なう
        If Not CrLtjudg(tSiyou, tCryLT, retJudg) Then GoTo ChkSxlHaneiNG
    
    Case 13             'EPD
        'EPDの実績値を取得する
        If funGetCryEPDJisseki(sCryNum, wGetSmplID, tCryEPD) <> 0 Then GoTo ChkSxlHaneiNG
        'EPD総合判定を行なう
        If Not CrEpdjudg(tSiyou, tCryEPD, retJudg) Then GoTo ChkSxlHaneiNG

    'Add Start 2011/01/31 SMPK Miyata
    Case 15             'C
        If funGetCryCJisseki(sCryNum, wGetSmplID, tCryC) <> 0 Then GoTo ChkSxlHaneiNG

         '結晶番号をｷｰとしてXSDC1よりC－OSF3判定IDを獲得する
        If GetCOSF3ID(strJDGEIDC, Trim(sCryNum)) <> FUNCTION_RETURN_SUCCESS Then
            '該当ﾚｺｰﾄﾞ無しの場合
            retJudg = False
            GoTo ChkSxlHaneiNG
        End If
        'C－OSF3判定IDがNULLの場合
        If Trim(strJDGEIDC) = "" Then
            retJudg = False
            GoTo ChkSxlHaneiNG
        Else
            '獲得したC-OSF3判定IDでXODC5_OSF30より承認ﾌﾗｸﾞの獲得
            If GetSYNFLAGC5(strSynFlagc5, strYmkFlagc5, Trim(strJDGEIDC)) <> FUNCTION_RETURN_SUCCESS Then
                '該当ﾚｺｰﾄﾞ無しの場合
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
            '承認ﾌﾗｸﾞ:0　未承認の場合
            If Trim(strSynFlagc5) = "0" Or Trim(strSynFlagc5) = "" Or IsNull(strSynFlagc5) Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            '削除ﾌﾗｸﾞ:1　無効の場合
            ElseIf Trim(strYmkFlagc5) = "1" Or Trim(strYmkFlagc5) = "" Or IsNull(strYmkFlagc5) Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
        End If

        'C総合判定を行なう
        If Not CrCjudg(tSiyou, tCryC, retJudg) Then GoTo ChkSxlHaneiNG

    Case 16             'CJ
        If funGetCryCJJisseki(sCryNum, wGetSmplID, tCryCJ) <> 0 Then GoTo ChkSxlHaneiNG

         '結晶番号をｷｰとしてXSDC1よりC－OSF3判定IDを獲得する
        If GetCOSF3ID(strJDGEIDC, Trim(sCryNum)) <> FUNCTION_RETURN_SUCCESS Then
            '該当ﾚｺｰﾄﾞ無しの場合
            retJudg = False
            GoTo ChkSxlHaneiNG
        End If
        'C－OSF3判定IDがNULLの場合
        If Trim(strJDGEIDC) = "" Then
            retJudg = False
            GoTo ChkSxlHaneiNG
        Else
            '獲得したC-OSF3判定IDでXODC5_OSF30より承認ﾌﾗｸﾞの獲得
            If GetSYNFLAGC5(strSynFlagc5, strYmkFlagc5, Trim(strJDGEIDC)) <> FUNCTION_RETURN_SUCCESS Then
                '該当ﾚｺｰﾄﾞ無しの場合
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
            '承認ﾌﾗｸﾞ:0　未承認の場合
            If Trim(strSynFlagc5) = "0" Or Trim(strSynFlagc5) = "" Or IsNull(strSynFlagc5) Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            '削除ﾌﾗｸﾞ:1　無効の場合
            ElseIf Trim(strYmkFlagc5) = "1" Or Trim(strYmkFlagc5) = "" Or IsNull(strYmkFlagc5) Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
        End If

        'CJ総合判定を行なう
        If Not CrCJjudg(tSiyou, tCryCJ, retJudg) Then GoTo ChkSxlHaneiNG
    
    Case 17             'CJLT
        If funGetCryCJLTJisseki(sCryNum, wGetSmplID, tCryCJLT) <> 0 Then GoTo ChkSxlHaneiNG

         '結晶番号をｷｰとしてXSDC1よりC－OSF3判定IDを獲得する
        If GetCOSF3ID(strJDGEIDC, Trim(sCryNum)) <> FUNCTION_RETURN_SUCCESS Then
            '該当ﾚｺｰﾄﾞ無しの場合
            retJudg = False
            GoTo ChkSxlHaneiNG
        End If
        'C－OSF3判定IDがNULLの場合
        If Trim(strJDGEIDC) = "" Then
            retJudg = False
            GoTo ChkSxlHaneiNG
        Else
            '獲得したC-OSF3判定IDでXODC5_OSF30より承認ﾌﾗｸﾞの獲得
            If GetSYNFLAGC5(strSynFlagc5, strYmkFlagc5, Trim(strJDGEIDC)) <> FUNCTION_RETURN_SUCCESS Then
                '該当ﾚｺｰﾄﾞ無しの場合
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
            '承認ﾌﾗｸﾞ:0　未承認の場合
            If Trim(strSynFlagc5) = "0" Or Trim(strSynFlagc5) = "" Or IsNull(strSynFlagc5) Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            '削除ﾌﾗｸﾞ:1　無効の場合
            ElseIf Trim(strYmkFlagc5) = "1" Or Trim(strYmkFlagc5) = "" Or IsNull(strYmkFlagc5) Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
        End If

        'CJLT総合判定を行なう
        If Not CrCJLTjudg(tSiyou, tCryCJLT, retJudg) Then GoTo ChkSxlHaneiNG

    Case 18             'CJ2
        If funGetCryCJ2Jisseki(sCryNum, wGetSmplID, tCryCJ2) <> 0 Then GoTo ChkSxlHaneiNG

         '結晶番号をｷｰとしてXSDC1よりC－OSF3判定IDを獲得する
        If GetCOSF3ID(strJDGEIDC, Trim(sCryNum)) <> FUNCTION_RETURN_SUCCESS Then
            '該当ﾚｺｰﾄﾞ無しの場合
            retJudg = False
            GoTo ChkSxlHaneiNG
        End If
        'C－OSF3判定IDがNULLの場合
        If Trim(strJDGEIDC) = "" Then
            retJudg = False
            GoTo ChkSxlHaneiNG
        Else
            '獲得したC-OSF3判定IDでXODC5_OSF30より承認ﾌﾗｸﾞの獲得
            If GetSYNFLAGC5(strSynFlagc5, strYmkFlagc5, Trim(strJDGEIDC)) <> FUNCTION_RETURN_SUCCESS Then
                '該当ﾚｺｰﾄﾞ無しの場合
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
            '承認ﾌﾗｸﾞ:0　未承認の場合
            If Trim(strSynFlagc5) = "0" Or Trim(strSynFlagc5) = "" Or IsNull(strSynFlagc5) Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            '削除ﾌﾗｸﾞ:1　無効の場合
            ElseIf Trim(strYmkFlagc5) = "1" Or Trim(strYmkFlagc5) = "" Or IsNull(strYmkFlagc5) Then
                retJudg = False
                GoTo ChkSxlHaneiNG
            End If
        End If

        'CJ2総合判定を行なう
        If Not CrCJ2judg(tSiyou, tCryCJ2, retJudg) Then GoTo ChkSxlHaneiNG
    'Add End   2011/01/31 SMPK Miyata

    End Select

    '指定された評価項目№の総合判定がOKの場合、反映元サンプルIDを設定し、戻り値に'0'(正常終了(反映OK))を設定し、処理を終了する。
    '総合判定がNGの場合、戻り値に'1'(正常終了(反映NG))を設定し、処理を終了する。
    If retJudg = False Then GoTo ChkSxlHaneiNG
        
    iGetSmplID = wGetSmplID
    funChkSxlHanei = 0
    Exit Function

ChkSxlHaneiNG:
    iGetSmplID = wGetSmplID
    funChkSxlHanei = 1
    Exit Function

ChkSxlHaneiParameterErr:
    funChkSxlHanei = -1
    Exit Function

ChkSxlHaneiSonotaErr:
    funChkSxlHanei = -2
    Exit Function
    
End Function

'------------------------------------------------
' 結晶推定チェック
'------------------------------------------------

'概要      :指定された情報から、結晶推定チェックを行ない結果を返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :sBlockid      ,I  ,String       :ﾌﾞﾛｯｸID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :結晶番号
'          :tFullHin      ,I  ,tFullHinban  :品番(構造体)
'          :iSmplPos      ,I  ,Integer      :新サンプル位置(mm)
'          :iItemNo       ,I  ,Integer      :評価項目№ =  1 RS     ←対象
'                                                       =  2 Oi
'                                                       =  3 BMD1
'                                                       =  4 BMD2
'                                                       =  5 BMD3
'                                                       =  6 OSF1
'                                                       =  7 OSF2
'                                                       =  8 OSF3
'                                                       =  9 OSF4
'                                                       = 10 CS
'                                                       = 11 GD
'                                                       = 12 LT
'                                                       = 13 EPD
'          :iGetSmplID1   ,O  ,long         :推定元サンプルID1      Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :iGetSmplID2   ,O  ,long         :推定元サンプルID2      Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :戻り値        ,O  ,Integer      :チェック結果 = 0 : 正常終了(推定OK)
'                                                           1 : 正常終了(推定NG)
'                                                          -1 : 入力引数値エラー
'                                                          -2 : 上記以外のエラー
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funChkSxlSuitei(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                                iItemNo As Integer, iGetSmplID1 As Long, iGetSmplID2 As Long, tfullhin1 As tFullHinban) As Integer
    Dim retCode         As Integer
    Dim tSiyou          As type_DBDRV_scmzc_fcmkc001c_Siyou
    Dim tTBCME037       As c_cmzcXl
    Dim sqlWhere        As String
    
    Dim wGetBlockidTop  As String
    Dim wGetSmpKbnTop   As String
    Dim wGetSmplIDTop   As Long         'Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    Dim wGetSPtrnTop    As String
    Dim wGetBlockidBot  As String
    Dim wGetSmpKbnBot   As String
    Dim wGetSmplIDBot   As Long         'Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    Dim wGetSPtrnBot    As String
    Dim wGetPosTop      As Integer    '2005/1/11
    Dim wGetPosBot      As Integer    '2005/1/11
    Dim tCryRs(2)       As type_DBDRV_scmzc_fcmkc001c_CryR        '(0)→推定元Top, (1)→推定元Bot, (2)→推定先
    Dim wcnt            As Integer
    Dim wMeasTop(4)     As Double                   'Top測定値
    Dim wMeasBot(4)     As Double                   'Bot測定値
    Dim wMeasSui()      As Double                   '算出推定値
    Dim retJudg         As Boolean
    
    Dim dShiyo(5)       As Double       '2003/12/11 Null対応追加
    
    Dim i      As Integer  'TEST2004/10
    Dim i2     As Integer  'TEST2004/10
    Dim sCnt   As Integer  'TEST2004/10
    '初期化
    wGetSmplIDTop = -1
    wGetSmplIDBot = -1
    sCnt = UBound(SuiteiData) + 1
    ReDim Preserve SuiteiData(sCnt)
    For i2 = 0 To 2
        SuiteiData(sCnt).SuiData(i2).MEAS1 = 0
        SuiteiData(sCnt).SuiData(i2).MEAS2 = 0
        SuiteiData(sCnt).SuiData(i2).MEAS3 = 0
        SuiteiData(sCnt).SuiData(i2).MEAS4 = 0
        SuiteiData(sCnt).SuiData(i2).MEAS5 = 0
    Next
    'パラメータチェック
    If (Len(sBlockId) <> 12) Then GoTo ChkSxlSuiteiParameterErr
    If (Len(sCryNum) <> 12) Then GoTo ChkSxlSuiteiParameterErr
    
    '指定された評価項目№毎に必要な品番仕様値を取得する。（指定された評価項目№により、処理が分かれる。）
    Select Case iItemNo
    Case 1              'RS(比抵抗)
        If Trim(tFullHin.hinban) <> "Z" And Trim(tFullHin.hinban) <> "G" Then
            If funGet_TBCME018(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlSuiteiNG
        End If
        'Null対応処理追加 2003/12/11 SystenBrain ▽
        dShiyo(1) = tSiyou.HSXRMIN          ' 品ＳＸ比抵抗下限
        dShiyo(2) = tSiyou.HSXRMAX          ' 品ＳＸ比抵抗上限
        dShiyo(3) = tSiyou.HSXRAMIN         ' 品ＳＸ比抵抗平均下限
        dShiyo(4) = tSiyou.HSXRAMAX         ' 品ＳＸ比抵抗平均上限
        dShiyo(5) = tSiyou.HSXRMBNP         ' 品ＳＸ比抵抗面内分布
        'NULLは不問(NULLﾁｪｯｸ関数ｺﾒﾝﾄｱｳﾄ) 09/03/13 ooba
'        If fncJissekiHantei_nl(tSiyou.HSXRHWYS, dShiyo) = False Then GoTo ChkSxlSuiteiSonotaErr
        'Null対応処理追加 2003/12/11 SystenBrain △
        
    Case 2 To 13        'Oi(酸素濃度),BMD1,BMD2,BMD3,OSF1,OSF2,OSF3,OSF4,CS(炭素濃度),GD,LT(ﾗｲﾌﾀｲﾑ),EPD
        GoTo ChkSxlSuiteiNG
    Case Else
        GoTo ChkSxlSuiteiParameterErr
    End Select
    '結晶推定元サンプルＩＤの取得    '2005/1/11 修正　位置追加
    If funGetSuitei(sBlockId, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, _
                    wGetBlockidTop, wGetSmpKbnTop, wGetSmplIDTop, wGetSPtrnTop, _
                    wGetBlockidBot, wGetSmpKbnBot, wGetSmplIDBot, wGetSPtrnBot, wGetPosTop, wGetPosBot) <> 0 Then GoTo ChkSxlSuiteiNG
    
    '結晶推定元ﾌﾞﾛｯｸID1、結晶推定元ｻﾝﾌﾟﾙ区分1、結晶推定元ｻﾝﾌﾟﾙID1から、推定元実績値を取得する。
    'RSの実績値を取得する
    If funGetCryRsJisseki(sCryNum, wGetSmplIDTop, tCryRs(0)) <> 0 Then GoTo ChkSxlSuiteiNG
    '結晶推定元ブロックID2､結晶推定元サンプル区分2､結晶推定元サンプルID2から､推定元実績値を取得する｡
    'RSの実績値を取得する
    If funGetCryRsJisseki(sCryNum, wGetSmplIDBot, tCryRs(1)) <> 0 Then GoTo ChkSxlSuiteiNG
    
    With SuiteiData(sCnt)
        .SuiSpec = tSiyou
        .SuiData(0) = tCryRs(0)
        .SuiData(1) = tCryRs(1)
        Debug.Print .SuiSpec.HIN.hinban
        '実測値のチェック（各品番の製作条件をクリアしないと推定できない）
        .RsJudg(1) = True
        .RsJudg(2) = True
        If Trim(tFullHin.hinban) <> "Z" And Trim(tFullHin.hinban) <> "G" Then
            If funChkJissoku(tFullHin, tCryRs(0)) = False Then .RsJudg(1) = False
            If funChkJissoku(tFullHin, tCryRs(1)) = False Then .RsJudg(2) = False
        End If
        If Trim(tfullhin1.hinban) <> "Z" And Trim(tfullhin1.hinban) <> "G" Then
            If funChkJissoku(tfullhin1, tCryRs(0)) = False Then .RsJudg(1) = False
            If funChkJissoku(tfullhin1, tCryRs(1)) = False Then .RsJudg(2) = False
        End If
        If .RsJudg(1) = False Or .RsJudg(2) = False Then GoTo ChkSxlSuiteiNG
    End With
    '--------------
    '推定先の実績データ編集
    With tCryRs(2)
        .CRYNUM = sCryNum
        .POSITION = iSmplPos
        .SMPKBN = sTB
        .TRANCOND = "0"
        .TRANCNT = 1
        .SMPLNO = -1
        .SMPLUMU = "1"
    End With
    
    '------TEST2004/10
    If tCryRs(0).KSTAFFID = KSTAFF_J002 Then
        GoTo ChkSxlSuiteiNG
    End If
    wcnt = funGetRsCnt(tCryRs(0))
    If wcnt <> 5 Then GoTo ChkSxlSuiteiNG
    
    'Top/Bot測定値を推定値算出用にセット
    If wGetSPtrnTop = "A" Then                  '推定パターンA
        wMeasTop(0) = tCryRs(0).MEAS1
        wMeasTop(1) = tCryRs(0).MEAS2
        wMeasTop(2) = tCryRs(0).MEAS3
        wMeasTop(3) = tCryRs(0).MEAS4
        wMeasTop(4) = tCryRs(0).MEAS5
    ElseIf wGetSPtrnTop = "B" Then              '推定パターンB
        wMeasTop(0) = tCryRs(0).MEAS1
        wMeasTop(1) = tCryRs(0).MEAS4
        wMeasTop(2) = tCryRs(0).MEAS5
        wMeasTop(3) = 0
        wMeasTop(4) = 0
    End If
    
    '------TEST2004/10
    If tCryRs(1).KSTAFFID = KSTAFF_J002 Then
        GoTo ChkSxlSuiteiNG
    End If
    wcnt = funGetRsCnt(tCryRs(1))
    If wcnt <> 5 Then GoTo ChkSxlSuiteiNG
    
    If wGetSPtrnBot = "A" Then                  '推定パターンA
        wMeasBot(0) = tCryRs(1).MEAS1
        wMeasBot(1) = tCryRs(1).MEAS2
        wMeasBot(2) = tCryRs(1).MEAS3
        wMeasBot(3) = tCryRs(1).MEAS4
        wMeasBot(4) = tCryRs(1).MEAS5
    ElseIf wGetSPtrnBot = "B" Then              '推定パターンB
        wMeasBot(0) = tCryRs(1).MEAS1
        wMeasBot(1) = tCryRs(1).MEAS4
        wMeasBot(2) = tCryRs(1).MEAS5
        wMeasBot(3) = 0
        wMeasBot(4) = 0
    End If
    
    ReDim wMeasSui(5 - 1)
    For wcnt = 0 To 5 - 1
        '推定値の算出
        retCode = new_ResSuitei(sCryNum, wMeasTop(wcnt), tCryRs(0).POSITION, wMeasBot(wcnt), tCryRs(1).POSITION, iSmplPos, wMeasSui(wcnt))
        If retCode = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlSuiteiNG
    
    Next wcnt
    '結晶情報(TBCME037)データの取得
    sqlWhere = " Where (CRYNUM='" & sCryNum & "')"
    If GetTBCME037(tTBCME037, sqlWhere) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlSuiteiSonotaErr
    
    '-----TEST2004/10
    '推定値の設定
    tCryRs(2).MEAS1 = wMeasSui(0)
    tCryRs(2).MEAS2 = wMeasSui(1)
    tCryRs(2).MEAS3 = wMeasSui(2)
    tCryRs(2).MEAS4 = wMeasSui(3)
    tCryRs(2).MEAS5 = wMeasSui(4)
    '-----TEST2004/10
    SuiteiData(sCnt).SuiData(2) = tCryRs(2)
    '抵抗データを測定位置により並べ替える
    If Trim(tFullHin.hinban) = "Z" Or Trim(tFullHin.hinban) = "G" Then
        retJudg = True
        SuiteiData(sCnt).SuiSpec.HIN.hinban = Trim(tFullHin.hinban)
        SuiteiData(sCnt).COEFflg = True
        SuiteiData(sCnt).DOPEflg = True
    Else
        If Set_Rs_Ichi(tSiyou.HSXRSPOT, tSiyou.HSXRSPOI, tCryRs(2).MEAS1, tCryRs(2).MEAS2, _
                           tCryRs(2).MEAS3, tCryRs(2).MEAS4, tCryRs(2).MEAS5) Then GoTo ChkSxlSuiteiNG
        
        '推定計算で算出した推定値でRS総合判定を行なう
        If Not CrResJudg(0, tSiyou, tCryRs(2), retJudg, 1) Then GoTo ChkSxlSuiteiNG
        '2005/1/11 ブロック偏析値範囲外,追ﾄﾞｰﾌﾟ位置を含むブロックは推定不可
        If HenDopeJudg(wGetPosTop, iSmplPos, tCryRs(0).MEAS1, tCryRs(2).MEAS1, sCryNum, tFullHin) = False Then
            GoTo ChkSxlSuiteiNG
        End If
        If HenDopeJudg(iSmplPos, wGetPosBot, tCryRs(2).MEAS1, tCryRs(1).MEAS1, sCryNum, tFullHin) = False Then
            GoTo ChkSxlSuiteiNG
        End If
    End If
    '指定された評価項目№の総合判定がOKの場合、推定元サンプルID1と推定元サンプルID2を設定し、戻り値に'0'(正常終了(推定OK))を設定し、処理を終了する。
    '総合判定がNGの場合、戻り値に'1'(正常終了(推定NG))を設定し、処理を終了する。
    
    If retJudg = False Then GoTo ChkSxlSuiteiNG
        
    iGetSmplID1 = wGetSmplIDTop
    iGetSmplID2 = wGetSmplIDBot
    funChkSxlSuitei = 0
    Exit Function

ChkSxlSuiteiNG:
    iGetSmplID1 = wGetSmplIDTop
    iGetSmplID2 = wGetSmplIDBot
    funChkSxlSuitei = 1
    Exit Function

ChkSxlSuiteiParameterErr:
    funChkSxlSuitei = -1
    Exit Function

ChkSxlSuiteiSonotaErr:
    funChkSxlSuitei = -2
End Function

'------------------------------------------------
' 結晶反映値取得（パターン１）
'------------------------------------------------

'概要      :指定された新サンプル位置情報から、結晶反映元サンプルＩＤを新サンプル管理(ﾌﾞﾛｯｸ)(XSDCS)より検索し、結果を返す。
'           反映しようとする新サンプル位置が、TOPの場合とBOTの場合で検索方法(方向)が異なる。
'           反映元サンプルＩＤを検索する場合、基本的には、新サンプル位置から見て、上下サンプルの中で近いほうのサンプルＩＤを抽出する。
'           検索する際の検索範囲は、指定された範囲内のみ有効とし、検索範囲内にみつからない場合、「該当ｻﾝﾌﾟﾙなし」とする。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :sBlockid      ,I  ,String       :ﾌﾞﾛｯｸID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :結晶番号
'          :tFullHin      ,I  ,tFullHinban  :品番(構造体)
'          :iSmplPos      ,I  ,Integer      :新サンプル位置(mm)
'          :iItemNo       ,I  ,Integer      :評価項目№ =  1 RS
'                                                       =  2 Oi     ←対象
'                                                       =  3 BMD1   ←対象
'                                                       =  4 BMD2   ←対象
'                                                       =  5 BMD3   ←対象
'                                                       =  6 OSF1   ←対象
'                                                       =  7 OSF2   ←対象
'                                                       =  8 OSF3   ←対象
'                                                       =  9 OSF4   ←対象
'                                                       = 10 CS
'                                                       = 11 GD     ←対象
'                                                       = 12 LT
'                                                       = 13 EPD
'          :iFromPos      ,I  ,Integer      :検索範囲From
'          :iToPos        ,I  ,Integer      :検索範囲To
'          :sGetBlockid   ,O  ,String       :反映元ブロックＩＤ
'          :sGetSmpKbn    ,O  ,String       :反映元サンプル区分
'          :iGetSmplID    ,O  ,long         :反映元サンプルＩＤ     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :戻り値        ,O  ,Integer      :取得結果 = 0 : 正常終了
'                                                       1 : 正常終了(該当サンプルなし)
'                                                      -1 : 異常終了
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funGetSxlHanei1(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                              iItemNo As Integer, iFromPos As Integer, iToPos As Integer, _
                               sGetBlockid As String, sGetSmpKbn As String, iGetSmplID As Long) As Integer
    Dim kName       As String
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim ediSmpid    As String       'ｻﾝﾌﾟﾙID名称
    Dim ediInd      As String       '状態FLG名称
    Dim ediRes      As String       '実績FLG名称
    
    'パラメータチェック
    If (Len(sBlockId) <> 12) Then GoTo GetSxlHanei1ParameterErr
    If (Len(sCryNum) <> 12) Then GoTo GetSxlHanei1ParameterErr
    
    '指定された評価項目№から、検索対照評価項目名を決定する。
    kName = funGetCryKensaName(iItemNo)
    If kName = " " Then GoTo GetSxlHanei1ParameterErr
    
    'SQL文内で使用する名称に編集
    ediSmpid = cCRYSMPLID & kName & cCS     'ｻﾝﾌﾟﾙID
    ediInd = cCRYIND & kName & cCS          '状態FLG
    ediRes = cCRYRES & kName & cCS          '実績FLG
    
    '指定された情報を元に、新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ)(XSDCS)を検索する。
    sql = "select CRYNUMCS, SMPKBNCS, " & ediSmpid & " as SMPLID from XSDCS "

    'TOP位置(T/B区分='T')の検索
    If sTB = "T" Then
        sql = sql & "where tbkbncs = '" & sTB & "' and "
        sql = sql & "      xtalcs = '" & sCryNum & "' and "
        sql = sql & "      inposcs <= " & iSmplPos & " and "
        sql = sql & "      (" & ediInd & " = '1' or " & ediInd & " = '2') and "
        sql = sql & "  " & ediRes & " <> '0' and "
        sql = sql & "      inposcs >= " & iFromPos & " and "
        sql = sql & "      inposcs <= " & iToPos & " "
        sql = sql & "order by inposcs desc"
    
    'BOT位置(T/B区分='B')の検索
    ElseIf sTB = "B" Then
        sql = sql & "where tbkbncs = '" & sTB & "' and "
        sql = sql & "      xtalcs = '" & sCryNum & "' and "
        sql = sql & "      inposcs >= " & iSmplPos & " and "
        sql = sql & "      (" & ediInd & " = '1' or " & ediInd & " = '2') and "
        sql = sql & "  " & ediRes & " <> '0' and "
        sql = sql & "      inposcs >= " & iFromPos & " and "
        sql = sql & "      inposcs <= " & iToPos & " "
        sql = sql & "order by inposcs asc"
    Else
        GoTo GetSxlHanei1ParameterErr
    End If
    
    'SQL文の実行
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetSxlHanei1 = 1
        Set rs = Nothing
        Exit Function
    End If
    
    '呼び出し元への結果通知
    sGetBlockid = rs("CRYNUMCS")
    sGetSmpKbn = rs("SMPKBNCS")
    iGetSmplID = rs("SMPLID")
    Set rs = Nothing
    
    funGetSxlHanei1 = 0
    Exit Function

GetSxlHanei1ParameterErr:
    funGetSxlHanei1 = -1
End Function

'------------------------------------------------
' 結晶反映値取得（パターン２）
'------------------------------------------------

'概要      :指定された新サンプル位置情報から、結晶反映元サンプルＩＤを新サンプル管理(ﾌﾞﾛｯｸ)(XSDCS)より検索し、結果を返す。
'           反映元サンプルＩＤを検索する場合、基本的には、新サンプル位置から見て、下サンプルの中で近いほうのサンプルＩＤを抽出する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :sBlockid      ,I  ,String       :ﾌﾞﾛｯｸID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :結晶番号
'          :tFullHin      ,I  ,tFullHinban  :品番(構造体)
'          :iSmplPos      ,I  ,Integer      :新サンプル位置(mm)
'          :iItemNo       ,I  ,Integer      :評価項目№ =  1 RS
'                                                       =  2 Oi
'                                                       =  3 BMD1
'                                                       =  4 BMD2
'                                                       =  5 BMD3
'                                                       =  6 OSF1
'                                                       =  7 OSF2
'                                                       =  8 OSF3
'                                                       =  9 OSF4
'                                                       = 10 CS     ←対象
'                                                       = 11 GD
'                                                       = 12 LT
'                                                       = 13 EPD    ←対象
'          :sGetBlockid   ,O  ,String       :反映元ブロックＩＤ
'          :sGetSmpKbn    ,O  ,String       :反映元サンプル区分
'          :iGetSmplID    ,O  ,Long         :反映元サンプルＩＤ     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :戻り値        ,O  ,Integer      :取得結果 = 0 : 正常終了
'                                                       1 : 正常終了(該当サンプルなし)
'                                                      -1 : 異常終了
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funGetSxlHanei2(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                               iItemNo As Integer, sGetBlockid As String, sGetSmpKbn As String, iGetSmplID As Long) As Integer
    Dim kName       As String
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim ediSmpid    As String       'ｻﾝﾌﾟﾙID名称
    Dim ediInd      As String       '状態FLG名称
    Dim ediRes      As String       '実績FLG名称
    
    'パラメータチェック
    If (Len(sBlockId) <> 12) Then GoTo GetSxlHanei2ParameterErr
    If (Len(sCryNum) <> 12) Then GoTo GetSxlHanei2ParameterErr
    
    '指定された評価項目№から、検索対照評価項目名を決定する。
    kName = funGetCryKensaName(iItemNo)
    If kName = " " Then GoTo GetSxlHanei2ParameterErr
    
    'SQL文内で使用する名称に編集
    ediSmpid = cCRYSMPLID & kName & cCS     'ｻﾝﾌﾟﾙID
    ediInd = cCRYIND & kName & cCS          '状態FLG
    ediRes = cCRYRES & kName & cCS          '実績FLG
    
    '指定された情報を元に、新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ)(XSDCS)を検索する。
    sql = "select CRYNUMCS, SMPKBNCS, " & ediSmpid & " as SMPLID from XSDCS "
'' 09/03/02 FAE)akiyama start
'    sql = sql & "where xtalcs = '" & sCryNum & "' and "
    sql = sql & "where CRYNUMCS LIKE '" & left(sCryNum, 9) & "%' and "
'' 09/03/02 FAE)akiyama end
    sql = sql & "      inposcs > " & iSmplPos & " and "
    sql = sql & "      (" & ediInd & " = '1' or " & ediInd & " = '2') and "
    sql = sql & "  " & ediRes & " <> '0' "
    sql = sql & "order by inposcs asc"
    
    'SQL文の実行
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetSxlHanei2 = 1
        Set rs = Nothing
        Exit Function
    End If
    
    '呼び出し元への結果通知
    sGetBlockid = rs("CRYNUMCS")
    sGetSmpKbn = rs("SMPKBNCS")
    iGetSmplID = rs("SMPLID")
    Set rs = Nothing
    
    funGetSxlHanei2 = 0
    Exit Function

GetSxlHanei2ParameterErr:
    funGetSxlHanei2 = -1
End Function


'------------------------------------------------
' 結晶反映値取得 (パターン3)
'------------------------------------------------
'
'概要      :指定された新サンプル位置情報から、結晶反映元サンプルＩＤを新サンプル管理(ﾌﾞﾛｯｸ)(XSDCS)より検索し、結果を返す。
'           反映元サンプルＩＤを検索する場合、結晶内で一番厳しい仕様を持つ品番のサンプルを、
'            新サンプル位置から見て、下サンプルの中で近いほうのサンプルＩＤを抽出する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :sBlockid      ,I  ,String       :ﾌﾞﾛｯｸID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :結晶番号
'          :tFullHin      ,I  ,tFullHinban  :品番(構造体)
'          :iSmplPos      ,I  ,Integer      :新サンプル位置(mm)
'          :iItemNo       ,I  ,Integer      :評価項目№ =  1 RS
'                                                       =  2 Oi
'                                                       =  3 BMD1
'                                                       =  4 BMD2
'                                                       =  5 BMD3
'                                                       =  6 OSF1
'                                                       =  7 OSF2
'                                                       =  8 OSF3
'                                                       =  9 OSF4
'                                                       = 10 CS
'                                                       = 11 GD
'                                                       = 12 LT     ←対象
'                                                       = 13 EPD
'          :sGetBlockid   ,O  ,String       :反映元ブロックＩＤ
'          :sGetSmpKbn    ,O  ,String       :反映元サンプル区分
'          :iGetSmplID    ,O  ,Long         :反映元サンプルＩＤ     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :戻り値        ,O  ,Integer      :取得結果 = 0 : 正常終了
'                                                       1 : 正常終了(該当サンプルなし)
'                                                      -1 : 異常終了
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン
'
Public Function funGetSxlHanei3(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                                iItemNo As Integer, sHsxLtspi As String, _
                                sGetBlockid As String, sGetSmpKbn As String, iGetSmplID As Long) As Integer
    Dim kName       As String
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim ediSmpid    As String       'ｻﾝﾌﾟﾙID名称
    Dim ediInd      As String       '状態FLG名称
    Dim ediRes      As String       '実績FLG名称
    Dim GETHINBAN   As tFullHinban  '一番厳しい仕様の品番
    Dim LTsmpid    As String       '実績の検索結果サンプルNO
    
    'パラメータチェック
    If (Len(sBlockId) <> 12) Then GoTo GetSxlHanei3ParameterErr
    If (Len(sCryNum) <> 12) Then GoTo GetSxlHanei3ParameterErr
    
    '指定された評価項目№から、検索対照評価項目名を決定する。
    kName = funGetCryKensaName(iItemNo)
    If kName = " " Then GoTo GetSxlHanei3ParameterErr
    
    'SQL文内で使用する名称に編集
    ediSmpid = cCRYSMPLID & kName & cCS     'ｻﾝﾌﾟﾙID
    ediInd = cCRYIND & kName & cCS          '状態FLG
    ediRes = cCRYRES & kName & cCS          '実績FLG
    
    ''仕様が同じ場合、より近いほうのｻﾝﾌﾟﾙIDに対する品番を取得する　2004/01/26 ooba START ===========>
    sql = sql & "select E019.HINBAN, E019.MNOREVNO, E019.FACTORY, E019.OPECOND, J007.SMPLNO "
    sql = sql & "from TBCME019 E019, TBCMJ007 J007, XSDCS CS "
    sql = sql & "where E019.HINBAN = J007.HINBAN "
    sql = sql & "and E019.MNOREVNO = J007.REVNUM "
    sql = sql & "and E019.FACTORY = J007.FACTORY "
    sql = sql & "and E019.OPECOND = J007.OPECOND "
    sql = sql & "and J007.CRYNUM = CS.XTALCS "
    sql = sql & "and J007.SMPLNO = CS." & ediSmpid & " "
    sql = sql & "and CRYNUM = '" & sCryNum & "' "
    sql = sql & "and J007.TRANCNT = (select max(TRANCNT) from TBCMJ007 "
    sql = sql & "where CRYNUM = J007.CRYNUM "
    sql = sql & "and SMPLNO = J007.SMPLNO) "
    sql = sql & "and CS." & ediInd & " = '1' "
    sql = sql & "and CS." & ediRes & " <> '0' "
    sql = sql & "and E019.HSXLTSPI != ' ' "
    sql = sql & "and E019.HSXLTSPI <= '" & sHsxLtspi & "' "
    sql = sql & "order by E019.HSXLTSPI asc, J007.POSITION asc "
    ''仕様が同じ場合、より近いほうのｻﾝﾌﾟﾙIDに対する品番を取得する　2004/01/26 ooba END =============>
    
    'SQL文の実行
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetSxlHanei3 = 1
        Set rs = Nothing
        Exit Function
    End If
    
    '取得品番の設定
    GETHINBAN.hinban = rs("HINBAN")
    GETHINBAN.mnorevno = rs("MNOREVNO")
    GETHINBAN.factory = rs("FACTORY")
    GETHINBAN.opecond = rs("OPECOND")
    LTsmpid = rs("SMPLNO")              '取得品番のサンプルIDを設定するように変更　04/02/06 tuku
    
    Set rs = Nothing
    
    '指定された情報を元に、新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ)(XSDCS)を検索する。
    '取得したJ007の実績のサンプルIDを元に検索するように変更 04/02/06 tuku
    sql = "select CRYNUMCS, SMPKBNCS, " & ediSmpid & " as SMPLID from XSDCS "
'' 09/03/02 FAE)akiyama start
'    sql = sql & "where XTALCS = '" & sCryNum & "' and "
    sql = sql & "where CRYNUMCS LIKE '" & left(sCryNum, 9) & "%' and "
'' 09/03/02 FAE)akiyama end
    sql = sql & "      INPOSCS > " & iSmplPos & " and "
    sql = sql & "      (" & ediInd & " = '1' or " & ediInd & " = '2') and "
    sql = sql & "  " & ediRes & " <> '0' and "
    sql = sql & "  " & ediSmpid & " = '" & LTsmpid & "'  "
    sql = sql & "order by inposcs asc"
    
    'SQL文の実行
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetSxlHanei3 = 1
        Set rs = Nothing
        Exit Function
    End If
    
    '呼び出し元への結果通知
    sGetBlockid = rs("CRYNUMCS")
    sGetSmpKbn = rs("SMPKBNCS")
    iGetSmplID = rs("SMPLID")
    Set rs = Nothing
    
    funGetSxlHanei3 = 0
    Exit Function

GetSxlHanei3ParameterErr:
    funGetSxlHanei3 = -1
End Function

'------------------------------------------------
' 結晶推定値取得
'------------------------------------------------

'概要      :指定された新ｻﾝﾌﾟﾙ位置情報から、結晶推定元ｻﾝﾌﾟﾙID1と結晶推定元ｻﾝﾌﾟﾙID2を新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ)(XSDCS)より検索し、結果を返す。
'           推定元ｻﾝﾌﾟﾙID1と推定元ｻﾝﾌﾟﾙID2を検索する場合、結晶内の最TOPと最BOT位置の実測データを対象とする。
'           新ｻﾝﾌﾟﾙ位置の品番仕様と結晶内の最TOP／最BOT位置の品番仕様が、それぞれ「3点測定」「5点測定」でのﾊﾟﾀｰﾝが考えられるが、このﾊﾟﾀｰﾝによっても推定可否を判断する。
'           上記の測定点数パターンの組み合わせにより、取得すべき測定点データの位置(場所)が異なる。
'           推定可否の判断として、XSDC1のSUIFLGC1の値(0:推定許可,1:推定禁止)も考慮する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :sBlockid      ,I  ,String       :ﾌﾞﾛｯｸID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :結晶番号
'          :tFullHin      ,I  ,tFullHinban  :品番(構造体)
'          :iSmplPos      ,I  ,Integer      :新サンプル位置(mm)
'          :iItemNo       ,I  ,Integer      :評価項目№ =  1 RS     ←対象
'                                                       =  2 Oi
'                                                       =  3 BMD1
'                                                       =  4 BMD2
'                                                       =  5 BMD3
'                                                       =  6 OSF1
'                                                       =  7 OSF2
'                                                       =  8 OSF3
'                                                       =  9 OSF4
'                                                       = 10 CS
'                                                       = 11 GD
'                                                       = 12 LT
'                                                       = 13 EPD
'          :sGetBlockid1  ,O  ,String       :推定元ブロックＩＤ１
'          :sGetSmpKbn1   ,O  ,String       :推定元サンプル区分１
'          :iGetSmplID1   ,O  ,Long         :推定元サンプルＩＤ１   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :iGetPCode1    ,O  ,String       :推定元パターン１('A' or 'B')
'          :iGetPos1      ,O  ,Integr       :推定元サンプル位置1  2005/1/11
'          :sGetBlockid2  ,O  ,String       :推定元ブロックＩＤ２
'          :sGetSmpKbn2   ,O  ,String       :推定元サンプル区分２
'          :iGetSmplID2   ,O  ,Long         :推定元サンプルＩＤ２   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :iGetPCode2    ,O  ,String       :推定元パターン２('A' or 'B')
'          :iGetPos2      ,O  ,Integr       :推定元サンプル位置2  2005/1/11
'          :戻り値        ,O  ,Integer      :取得結果 = 0 : 正常終了
'                                                       1 : 正常終了(該当サンプルなし)
'                                                      -1 : 異常終了
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン
'          :TEST2004/10 引数追加,2005/1/11 引数追加
Public Function funGetSuitei(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, iItemNo As Integer, _
                             sGetBlockid1 As String, sGetSmpKbn1 As String, iGetSmplID1 As Long, iGetPCode1 As String, _
                             sGetBlockid2 As String, sGetSmpKbn2 As String, iGetSmplID2 As Long, iGetPCode2 As String, _
                             iGetPos1 As Integer, iGetPos2 As Integer) As Integer
    Dim kName       As String
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim ediSmpid    As String       'ｻﾝﾌﾟﾙID名称
    Dim ediInd      As String       '状態FLG名称
    Dim ediRes      As String       '実績FLG名称
    Dim getNewSpec  As String       '新ｻﾝﾌﾟﾙ位置比抵抗仕様値
    Dim getTopBlkID As String       'TOP位置ﾌﾞﾛｯｸID
    Dim getTopSmpK  As String       'TOP位置ｻﾝﾌﾟﾙ区分
    Dim getTopSmpID As String       'TOP位置ｻﾝﾌﾟﾙID
    Dim getTopHin   As tFullHinban  'TOP位置品番
    Dim getTopSpec  As String       'TOP位置比抵抗仕様値
    Dim getTopPtrn  As String       'TOP位置ﾊﾟﾀｰﾝｺｰﾄﾞ
    Dim getTopPos   As Integer      'TOP位置結晶内位置
    
    Dim getBotBlkID As String       'BOT位置ﾌﾞﾛｯｸID
    Dim getBotSmpK  As String       'BOT位置ｻﾝﾌﾟﾙ区分
    Dim getBotSmpID As String       'BOT位置ｻﾝﾌﾟﾙID
    Dim getBotHin   As tFullHinban  'BOT位置品番
    Dim getBotSpec  As String       'BOT位置比抵抗仕様値
    Dim getBotPtrn  As String       'BOT位置ﾊﾟﾀｰﾝｺｰﾄﾞ
    Dim getBotPos   As Integer      'BOT位置結晶内位置
    'パラメータチェック
    If (Len(sBlockId) <> 12) Then GoTo GetSuiteiParameterErr
    If (Len(sCryNum) <> 12) Then GoTo GetSuiteiParameterErr
    
    '指定された評価項目№から、検索対照評価項目名を決定する。
    kName = funGetCryKensaName(iItemNo)
    If kName = " " Then GoTo GetSuiteiParameterErr
    
    'SQL文内で使用する名称に編集
    ediSmpid = cCRYSMPLID & kName & cCS     'ｻﾝﾌﾟﾙID
    ediInd = cCRYIND & kName & cCS          '状態FLG
    ediRes = cCRYRES & kName & "1" & cCS    '実績FLG
    
    '指定された情報を元に、新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ)(XSDCS)を検索する。
    '≪推定元サンプルＩＤ１(TOP位置)の取得≫
    sql = "select CS.CRYNUMCS, CS.SMPKBNCS, CS." & ediSmpid & " as SMPLID, CS.HINBCS, CS.REVNUMCS, CS.FACTORYCS, CS.OPECS, CS.INPOSCS "  '2005/1/11
    sql = sql & "from XSDCS CS, XSDC1 C1 "
    sql = sql & "where CS.XTALCS = '" & sCryNum & "' and "
    sql = sql & "      CS.INPOSCS < " & iSmplPos & " and "
    sql = sql & "      CS." & ediInd & " = '1' and "
    sql = sql & "      CS." & ediInd & " <> '0' and "
    sql = sql & "      CS." & ediRes & " <> '0' and "
    sql = sql & "      C1.XTALC1 = CS.XTALCS and "
    sql = sql & "      C1.SUIFLG = '0' "
    sql = sql & "order by CS.INPOSCS desc"
    
    'SQL文の実行
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        Set rs = Nothing
        GoTo GetSuiteiEmpty
    End If
    
    'TOP位置データの設定
    getTopBlkID = rs("CRYNUMCS")            'TOP位置ﾌﾞﾛｯｸID
    getTopSmpK = rs("SMPKBNCS")             'TOP位置ｻﾝﾌﾟﾙ区分
    getTopSmpID = rs("SMPLID")                'TOP位置ｻﾝﾌﾟﾙID
    getTopHin.hinban = rs("HINBCS")         'TOP位置品番
    getTopHin.mnorevno = rs("REVNUMCS")     'TOP位置製品番号改訂番号
    getTopHin.factory = rs("FACTORYCS")     'TOP位置工場
    getTopHin.opecond = rs("OPECS")         'TOP位置操業条件
    getTopPos = rs("INPOSCS")               'TOP位置   2005/1/11
    Set rs = Nothing
    
    '≪推定元サンプルＩＤ２(BOT位置)の取得≫
    sql = "select CS.CRYNUMCS, CS.SMPKBNCS, CS." & ediSmpid & " as SMPLID, CS.HINBCS, CS.REVNUMCS, CS.FACTORYCS, CS.OPECS, CS.INPOSCS "   '2005/1/11
    sql = sql & "from XSDCS CS, XSDC1 C1 "
    sql = sql & "where CS.XTALCS = '" & sCryNum & "' and "
    sql = sql & "      CS.INPOSCS > " & iSmplPos & " and "
    sql = sql & "      CS." & ediInd & " = '1' and "
    sql = sql & "      CS." & ediInd & " <> '0' and "
    sql = sql & "      CS." & ediRes & " <> '0' and "
    sql = sql & "      C1.XTALC1 = CS.XTALCS and "
    sql = sql & "      C1.SUIFLG = '0' "
    sql = sql & "order by CS.INPOSCS asc"
    
    'SQL文の実行
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        Set rs = Nothing
        GoTo GetSuiteiEmpty
    End If
    
    'BOT位置データの設定
    getBotBlkID = rs("CRYNUMCS")            'BOT位置ﾌﾞﾛｯｸID
    getBotSmpK = rs("SMPKBNCS")             'BOT位置ｻﾝﾌﾟﾙ区分
    getBotSmpID = rs("SMPLID")                'BOT位置ｻﾝﾌﾟﾙID
    getBotHin.hinban = rs("HINBCS")         'BOT位置品番
    getBotHin.mnorevno = rs("REVNUMCS")     'BOT位置製品番号改訂番号
    getBotHin.factory = rs("FACTORYCS")     'BOT位置工場
    getBotHin.opecond = rs("OPECS")         'BOT位置操業条件
    getBotPos = rs("INPOSCS")               '2005/1/11
    Set rs = Nothing
    
    '各品番の比抵抗仕様値取得
    getTopPtrn = "A"
    getBotPtrn = "A"
    '------------------------------------------------------
    
    '呼び出し元への結果通知
    sGetBlockid1 = getTopBlkID      '推定元ブロックＩＤ１
    sGetSmpKbn1 = getTopSmpK        '推定元サンプル区分１
    iGetSmplID1 = getTopSmpID       '推定元サンプルＩＤ１
    iGetPCode1 = getTopPtrn         '推定元パターン１('A' or 'B')
    iGetPos1 = getTopPos            '推定元位置  2005/1/11
    
    sGetBlockid2 = getBotBlkID      '推定元ブロックＩＤ２
    sGetSmpKbn2 = getBotSmpK        '推定元サンプル区分２
    iGetSmplID2 = getBotSmpID       '推定元サンプルＩＤ２
    iGetPCode2 = getBotPtrn         '推定元パターン２('A' or 'B')
    iGetPos2 = getBotPos            '推定元位置  2005/1/11
    
    funGetSuitei = 0
    Exit Function

GetSuiteiEmpty:
    funGetSuitei = 1
    Exit Function

GetSuiteiParameterErr:
    funGetSuitei = -1
End Function

'------------------------------------------------
' 結晶推定 比抵抗仕様値取得関数
'------------------------------------------------

'概要      :指定された品番から、TBCME018を検索し、比抵抗仕様値(品SX比抵抗測定位置_位)を取得し返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :tFullHin      ,I  ,tFullHinban  :品番(構造体)
'          :戻り値        ,O  ,Sting        :比抵抗仕様値(品SX比抵抗測定位置_位)
'                                            (取得できない場合は、空白を返す)
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funGetSuiSpecRS(tFullHin As tFullHinban) As String
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    
    '指定された品番からTBCME018のHSXRSPOI(品SX比抵抗測定位置_位)を検索する。
    sql = "select HSXRSPOI from TBCME018 "
    sql = sql & "where HINBAN = '" & Trim(tFullHin.hinban) & "' and "
    sql = sql & "      MNOREVNO = " & tFullHin.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tFullHin.factory & "' and "
    sql = sql & "      OPECOND = '" & tFullHin.opecond & "'"
    
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetSuiSpecRS = " "
        Set rs = Nothing
        Exit Function
    End If
    
    'TOP位置データの設定
    funGetSuiSpecRS = rs("HSXRSPOI")
    Set rs = Nothing

End Function

'------------------------------------------------
' 結晶検査対象評価項目名取得
'------------------------------------------------

'概要      :評価項目№から、結晶検査対象評価項目名を返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :iItemNo       ,I  ,Integer      :評価項目№ ⇒ ﾌﾞﾛｯｸ =  1 RS
'                                                                =  2 Oi
'                                                                =  3 BMD1
'                                                                =  4 BMD2
'                                                                =  5 BMD3
'                                                                =  6 OSF1
'                                                                =  7 OSF2
'                                                                =  8 OSF3
'                                                                =  9 OSF4
'                                                                = 10 CS
'                                                                = 11 GD
'                                                                = 12 LT
'                                                                = 13 EPD
'          :戻り値        ,O  ,Sting        :検査対象項目名(ﾊﾟﾗﾒｰﾀｴﾗｰ時は、空白を返す)
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funGetCryKensaName(iItemNo As Integer) As String
    
    'パラメータチェック
    'If iItemNo < 1 Or iItemNo > 13 Then GoTo GetCryKensaNameParameterErr    'Chg 2011/01/19 SMPK Miyata
    If iItemNo < 1 Or iItemNo > 18 Then GoTo GetCryKensaNameParameterErr

    'ﾌﾞﾛｯｸ
    Select Case iItemNo
    Case 1:     funGetCryKensaName = cCRY_RS       'RS(比抵抗)
    Case 2:     funGetCryKensaName = cCRY_OI       'Oi(酸素濃度)
    Case 3:     funGetCryKensaName = cCRY_B1       'BMD1
    Case 4:     funGetCryKensaName = cCRY_B2       'BMD2
    Case 5:     funGetCryKensaName = cCRY_B3       'BMD3
    Case 6:     funGetCryKensaName = cCRY_O1       'OSF1
    Case 7:     funGetCryKensaName = cCRY_O2       'OSF2
    Case 8:     funGetCryKensaName = cCRY_O3       'OSF3
    Case 9:     funGetCryKensaName = cCRY_O4       'OSF4
    Case 10:    funGetCryKensaName = cCRY_CS       'CS(炭素濃度)
    Case 11:    funGetCryKensaName = cCRY_GD       'GD
    Case 12:    funGetCryKensaName = cCRY_LT       'LT(ﾗｲﾌﾀｲﾑ)
    Case 13:    funGetCryKensaName = cCRY_EP       'EPD
'Add Start 2011/01/19 SMPK Miyata
    Case 15:    funGetCryKensaName = cCRY_C        'C
    Case 16:    funGetCryKensaName = cCRY_CJ       'CJ
    Case 17:    funGetCryKensaName = cCRY_CJLT     'CJLT
    Case 18:    funGetCryKensaName = cCRY_CJ2      'CJ2
'Add End   2011/01/19 SMPK Miyata
    End Select
    
    Exit Function

GetCryKensaNameParameterErr:
    funGetCryKensaName = " "
End Function

'------------------------------------------------
' 結晶抵抗実績取得関数
'------------------------------------------------

'概要      :結晶番号、サンプルＩＤから、TBCMJ002を検索し、結晶抵抗実績値を取得し返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                               :説明
'          :sCryNum       ,I  ,String                           :結晶番号
'          :iSmplID       ,I  ,Long                             :サンプルＩＤ   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :tCryRs        ,O  ,type_DBDRV_scmzc_fcmkc001c_CryR  :結晶抵抗実績(構造体)
'          :戻り値        ,O  ,Integer                          :取得結果 = 0 : 正常
'                                                                          -1 : 異常
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funGetCryRsJisseki(sCryNum As String, iSmplID As Long, tCryRs As type_DBDRV_scmzc_fcmkc001c_CryR) As Integer

    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    Dim wkXsdcs     As typ_XSDCS
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryRsJisseki"
    
    '結晶番号、サンプルＩＤからTBCMJ002の結晶抵抗実績値を検索する。
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, "
    sql = sql & "MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, EFEHS, RRG, REGDATE, KSTAFFID "    '----TEST2004/10
    sql = sql & "from TBCMJ002 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetCryRsJisseki = -1
        GoTo proc_exit
    End If
    
     ''抽出結果を格納する
    With tCryRs
        .CRYNUM = rs("CRYNUM")              ' 結晶番号
        .POSITION = rs("POSITION")          ' 位置
        .SMPKBN = rs("SMPKBN")              ' サンプル区分
        .TRANCOND = rs("TRANCOND")          ' 処理条件
        .TRANCNT = rs("TRANCNT")            ' 処理回数
        .SMPLNO = rs("SMPLNO")              ' サンプルＮｏ
        .SMPLUMU = rs("SMPLUMU")            ' サンプル有無
        .MEAS1 = rs("MEAS1")                ' 測定値１
        .MEAS2 = rs("MEAS2")                ' 測定値２
        .MEAS3 = rs("MEAS3")                ' 測定値３
        .MEAS4 = rs("MEAS4")                ' 測定値４
        .MEAS5 = rs("MEAS5")                ' 測定値５
        .EFEHS = rs("EFEHS")                ' 実効偏析
        .RRG = rs("RRG")                    ' ＲＲＧ
        .REGDATE = rs("REGDATE")            ' 登録日付
        .KSTAFFID = rs("KSTAFFID")          '----TEST2004/10
'--------------- 2008/08/25 INSERT START  By Systech ---------------
        ' DK温度（実績）
        wkXsdcs.XTALCS = .CRYNUM
        wkXsdcs.CRYSMPLIDRSCS = .SMPLNO
        wkXsdcs.CRYINDRSCS = "0"
        .HSXDKTMP = GetDKTmpCode(False, wkXsdcs)
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
    End With
    
    funGetCryRsJisseki = 0

proc_exit:
    '終了
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' 結晶Oi実績取得関数
'------------------------------------------------

'概要      :結晶番号、サンプルＩＤから、TBCMJ003を検索し、結晶Oi実績値を取得し返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                               :説明
'          :sCryNum       ,I  ,String                           :結晶番号
'          :iSmplID       ,I  ,Long                             :サンプルＩＤ   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :tCryOi        ,O  ,type_DBDRV_scmzc_fcmkc001c_Oi    :結晶Oi実績(構造体)
'          :戻り値        ,O  ,Integer                          :取得結果 = 0 : 正常
'                                                                          -1 : 異常
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funGetCryOiJisseki(sCryNum As String, iSmplID As Long, tCryOi As type_DBDRV_scmzc_fcmkc001c_Oi) As Integer

    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryOiJisseki"
    
    '結晶番号、サンプルＩＤからTBCMJ003の結晶Oi実績値を検索する。
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, "
    sql = sql & "OIMEAS1, OIMEAS2, OIMEAS3, OIMEAS4, OIMEAS5, ORGRES, AVE, FTIRCONV, INSPECTWAY, REGDATE "
    sql = sql & "from TBCMJ003 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "  and TRANCOND = 0 "       'GFAのFTIR換算値表示異常対応 2011/01/20追加 SETsw kubota
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetCryOiJisseki = -1
        GoTo proc_exit
    End If
    
     ''抽出結果を格納する
    With tCryOi
        .CRYNUM = rs("CRYNUM")              ' 結晶番号
        .POSITION = rs("POSITION")          ' 位置
        .SMPKBN = rs("SMPKBN")              ' サンプル区分
        .TRANCOND = rs("TRANCOND")          ' 処理条件
        .TRANCNT = rs("TRANCNT")            ' 処理回数
        .SMPLNO = rs("SMPLNO")              ' サンプルＮｏ
        .SMPLUMU = rs("SMPLUMU")            ' サンプル有無
'OI_NULL対応　2005/03/08 TUKU START --------------------------------------------------
        If IsNull(rs("OIMEAS1")) = False Then .OIMEAS1 = rs("OIMEAS1") Else .OIMEAS1 = -1  'Ｏｉ測定値1
        If IsNull(rs("OIMEAS2")) = False Then .OIMEAS2 = rs("OIMEAS2") Else .OIMEAS2 = -1  'Ｏｉ測定値2
        If IsNull(rs("OIMEAS3")) = False Then .OIMEAS3 = rs("OIMEAS3") Else .OIMEAS3 = -1  'Ｏｉ測定値3
        If IsNull(rs("OIMEAS4")) = False Then .OIMEAS4 = rs("OIMEAS4") Else .OIMEAS4 = -1  'Ｏｉ測定値4
        If IsNull(rs("OIMEAS5")) = False Then .OIMEAS5 = rs("OIMEAS5") Else .OIMEAS5 = -1  'Ｏｉ測定値5
        If IsNull(rs("ORGRES")) = False Then .ORGRES = rs("ORGRES") Else .ORGRES = -1    ' ＯＲＧ結果
'OI_NULL対応　2005/03/08 TUKU END   --------------------------------------------------
        .AVE = rs("AVE")                    ' AVE
        .FTIRCONV = rs("FTIRCONV")          ' FTIR換算
        .INSPECTWAY = rs("INSPECTWAY")      ' 検査方法
        .REGDATE = rs("REGDATE")            ' 登録日付
    End With
    
    funGetCryOiJisseki = 0

proc_exit:
    '終了
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' 結晶BMD実績取得関数
'------------------------------------------------

'概要      :結晶番号、サンプルＩＤから、TBCMJ008を検索し、結晶BMD実績値を取得し返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                               :説明
'          :sCryNum       ,I  ,String                           :結晶番号
'          :iSmplID       ,I  ,Long                             :サンプルＩＤ   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :iTranCond     ,I  ,Integer                          :処理条件(1:BMD1, 2:BMD2, 3:BMD3)
'          :tCryBMD       ,O  ,type_DBDRV_scmzc_fcmkc001c_BMD   :結晶BMD実績(構造体)
'          :戻り値        ,O  ,Integer                          :取得結果 = 0 : 正常
'                                                                          -1 : 異常
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funGetCryBMDJisseki(sCryNum As String, iSmplID As Long, iTranCond As Integer, tCryBMD As type_DBDRV_scmzc_fcmkc001c_BMD) As Integer

    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryBMDJisseki"
    
    '結晶番号、サンプルＩＤからTBCMJ008の結晶BMD実績値を検索する。
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, "
    sql = sql & "HTPRC, KKSP, KKSET, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEASMIN, MEASMAX, MEASAVE, BMDMNBUNP, REGDATE "
    sql = sql & "from TBCMJ008 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      TRANCOND = '" & iTranCond & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetCryBMDJisseki = -1
        GoTo proc_exit
    End If
    
     ''抽出結果を格納する
    With tCryBMD
        .CRYNUM = rs("CRYNUM")              ' 結晶番号
        .POSITION = rs("POSITION")          ' 位置
        .SMPKBN = rs("SMPKBN")              ' サンプル区分
        .TRANCOND = rs("TRANCOND")          ' 処理条件
        .TRANCNT = rs("TRANCNT")            ' 処理回数
        .SMPLNO = rs("SMPLNO")              ' サンプルＮｏ
        .SMPLUMU = rs("SMPLUMU")            ' サンプル有無
        .HTPRC = rs("HTPRC")                ' 熱処理方法
        .KKSP = rs("KKSP")                  ' 結晶欠陥測定位置
        .KKSET = rs("KKSET")                ' 結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
        .MEAS1 = rs("MEAS1")                ' 測定値1
        .MEAS2 = rs("MEAS2")                ' 測定値2
        .MEAS3 = rs("MEAS3")                ' 測定値3
        .MEAS4 = rs("MEAS4")                ' 測定値4
        .MEAS5 = rs("MEAS5")                ' 測定値5
        .MEASMIN = rs("MEASMIN")            ' Min
        .MEASMAX = rs("MEASMAX")            ' max
        .MEASAVE = rs("MEASAVE")            ' AVE
        If Not IsNull(rs("BMDMNBUNP")) Then .BMDMNBUNP = rs("BMDMNBUNP")      ' BMD面内分布
        .REGDATE = rs("REGDATE")            ' 登録日付
    End With
    
    funGetCryBMDJisseki = 0

proc_exit:
    '終了
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' 結晶OSF実績取得関数
'------------------------------------------------

'概要      :結晶番号、サンプルＩＤから、TBCMJ005を検索し、結晶OSF実績値を取得し返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                               :説明
'          :sCryNum       ,I  ,String                           :結晶番号
'          :iSmplID       ,I  ,Long                             :サンプルＩＤ   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :iTranCond     ,I  ,Integer                          :処理条件(1:OSF1, 2:OSF2, 3:OSF3, 4:OSF4)
'          :tCryOSF       ,O  ,type_DBDRV_scmzc_fcmkc001c_OSF   :結晶OSF実績(構造体)
'          :戻り値        ,O  ,Integer                          :取得結果 = 0 : 正常
'                                                                          -1 : 異常
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funGetCryOSFJisseki(sCryNum As String, iSmplID As Long, iTranCond As Integer, tCryOSF As type_DBDRV_scmzc_fcmkc001c_OSF) As Integer

    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryOSFJisseki"
    
    '結晶番号、サンプルＩＤからTBCMJ005の結晶OSF実績値を検索する。
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, "
    sql = sql & "HTPRC, KKSP, KKSET, CALCMAX, CALCAVE, "
    sql = sql & "MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, MEAS6, MEAS7, MEAS8, MEAS9, MEAS10, "
    sql = sql & "MEAS11, MEAS12, MEAS13, MEAS14, MEAS15, MEAS16, MEAS17, MEAS18, MEAS19, MEAS20, "
    sql = sql & "OSFPOS1, OSFWID1, OSFRD1, OSFPOS2, OSFWID2, OSFRD2, OSFPOS3, OSFWID3, OSFRD3, REGDATE "
    sql = sql & ",CALCMH "  '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
    sql = sql & "from TBCMJ005 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      TRANCOND = '" & iTranCond & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetCryOSFJisseki = -1
        GoTo proc_exit
    End If

     ''抽出結果を格納する
    With tCryOSF
        .CRYNUM = rs("CRYNUM")              ' 結晶番号
        .POSITION = rs("POSITION")          ' 位置
        .SMPKBN = rs("SMPKBN")              ' サンプル区分
        .TRANCOND = rs("TRANCOND")          ' 処理条件
        .TRANCNT = rs("TRANCNT")            ' 処理回数
        .SMPLNO = rs("SMPLNO")              ' サンプルＮｏ
        .SMPLUMU = rs("SMPLUMU")            ' サンプル有無
        .HTPRC = rs("HTPRC")                ' 熱処理方法
        .KKSP = rs("KKSP")                  ' 結晶欠陥測定位置
        .KKSET = rs("KKSET")                ' 結晶欠陥測定条件＋選択ET代　　char(1)＋number(2)
        .CALCMAX = rs("CALCMAX")            ' 計算結果 Max
        .CALCAVE = rs("CALCAVE")            ' 計算結果 Ave
        .MEAS1 = rs("MEAS1")                ' 測定値1
        .MEAS2 = rs("MEAS2")                ' 測定値2
        .MEAS3 = rs("MEAS3")                ' 測定値3
        .MEAS4 = rs("MEAS4")                ' 測定値4
        .MEAS5 = rs("MEAS5")                ' 測定値5
        .MEAS6 = rs("MEAS6")                ' 測定値6
        .MEAS7 = rs("MEAS7")                ' 測定値7
        .MEAS8 = rs("MEAS8")                ' 測定値8
        .MEAS9 = rs("MEAS9")                ' 測定値9
        .MEAS10 = rs("MEAS10")              ' 測定値10
        .MEAS11 = rs("MEAS11")              ' 測定値11
        .MEAS12 = rs("MEAS12")              ' 測定値12
        .MEAS13 = rs("MEAS13")              ' 測定値13
        .MEAS14 = rs("MEAS14")              ' 測定値14
        .MEAS15 = rs("MEAS15")              ' 測定値15
        .MEAS16 = rs("MEAS16")              ' 測定値16
        .MEAS17 = rs("MEAS17")              ' 測定値17
        .MEAS18 = rs("MEAS18")              ' 測定値18
        .MEAS19 = rs("MEAS19")              ' 測定値19
        .MEAS20 = rs("MEAS20")              ' 測定値20
        If Not IsNull(rs("OSFPOS1")) Then .OSFPOS1 = rs("OSFPOS1")      ' パターン区分1位置
        If Not IsNull(rs("OSFWID1")) Then .OSFWID1 = rs("OSFWID1")      ' パターン区分1幅
        If Not IsNull(rs("OSFRD1")) Then .OSFRD1 = rs("OSFRD1")         ' パターン区分1R / D
        If Not IsNull(rs("OSFPOS2")) Then .OSFPOS2 = rs("OSFPOS2")      ' パターン区分2位置
        If Not IsNull(rs("OSFWID2")) Then .OSFWID2 = rs("OSFWID2")      ' パターン区分2幅
        If Not IsNull(rs("OSFRD2")) Then .OSFRD2 = rs("OSFRD2")         ' パターン区分2R / D
        If Not IsNull(rs("OSFPOS3")) Then .OSFPOS3 = rs("OSFPOS3")      ' パターン区分3位置
        If Not IsNull(rs("OSFWID3")) Then .OSFWID3 = rs("OSFWID3")      ' パターン区分3幅
        If Not IsNull(rs("OSFRD3")) Then .OSFRD3 = rs("OSFRD3")         ' パターン区分3R / D
        
        .CALCMH = fncNullCheck(rs("CALCMH"))    ' 面内比(MAX/MIN)   '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
        
        .REGDATE = rs("REGDATE")            ' 登録日付
    End With
    
    funGetCryOSFJisseki = 0

proc_exit:
    '終了
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' 結晶CS実績取得関数
'------------------------------------------------

'概要      :結晶番号、サンプルＩＤから、TBCMJ004を検索し、結晶CS実績値を取得し返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                               :説明
'          :sCryNum       ,I  ,String                           :結晶番号
'          :iSmplID       ,I  ,Long                             :サンプルＩＤ   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :tCryCS        ,O  ,type_DBDRV_scmzc_fcmkc001c_CS    :結晶CS実績(構造体)
'          :戻り値        ,O  ,Integer                          :取得結果 = 0 : 正常
'                                                                          -1 : 異常
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funGetCryCSJisseki(sCryNum As String, iSmplID As Long, tCryCS As type_DBDRV_scmzc_fcmkc001c_CS) As Integer

    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryCSJisseki"
    
    '結晶番号、サンプルＩＤからTBCMJ004の結晶CS実績値を検索する。
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, "
    sql = sql & "CSMEAS, PRE70P, INSPECTWAY, REGDATE "
    sql = sql & "from TBCMJ004 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetCryCSJisseki = -1
        GoTo proc_exit
    End If
    
     ''抽出結果を格納する
    With tCryCS
        .CRYNUM = rs("CRYNUM")              ' 結晶番号
        .POSITION = rs("POSITION")          ' 位置
        .SMPKBN = rs("SMPKBN")              ' サンプル区分
        .TRANCOND = rs("TRANCOND")          ' 処理条件
        .TRANCNT = rs("TRANCNT")            ' 処理回数
        .SMPLNO = rs("SMPLNO")              ' サンプルＮｏ
        .SMPLUMU = rs("SMPLUMU")            ' サンプル有無
'OI_NULL対応　2005/03/08 TUKU START --------------------------------------------------
            If IsNull(rs("CSMEAS")) = False Then .CSMEAS = rs("CSMEAS") Else .CSMEAS = -1  ' Cs実測値
            If IsNull(rs("PRE70P")) = False Then .PRE70P = rs("PRE70P") Else .PRE70P = -1  ' ７０％推定値
'OI_NULL対応　2005/03/08 TUKU START --------------------------------------------------
        .INSPECTWAY = rs("INSPECTWAY")      ' 検査方法
        .REGDATE = rs("REGDATE")            ' 登録日付
    End With
    
    funGetCryCSJisseki = 0

proc_exit:
    '終了
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' 結晶GD実績取得関数
'------------------------------------------------

'概要      :結晶番号、サンプルＩＤから、TBCMJ006を検索し、結晶GD実績値を取得し返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                               :説明
'          :sCryNum       ,I  ,String                           :結晶番号
'          :iSmplID       ,I  ,Long                             :サンプルＩＤ   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :tCryGD        ,O  ,type_DBDRV_scmzc_fcmkc001c_GD    :結晶GD実績(構造体)
'          :戻り値        ,O  ,Integer                          :取得結果 = 0 : 正常
'                                                                          -1 : 異常
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funGetCryGDJisseki(sCryNum As String, iSmplID As Long, tCryGD As type_DBDRV_scmzc_fcmkc001c_GD) As Integer

    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryGDJisseki"
    
    '結晶番号、サンプルＩＤからTBCMJ006の結晶GD実績値を検索する。
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, "
    sql = sql & "MSRSDEN, MSRSLDL, MSRSDVD2, "
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
    
    sql = sql & ", MSZEROMN, MSZEROMX " '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
    
    sql = sql & "from TBCMJ006 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetCryGDJisseki = -1
        GoTo proc_exit
    End If
    
     ''抽出結果を格納する
    With tCryGD
        .CRYNUM = rs("CRYNUM")              ' 結晶番号
        .POSITION = rs("POSITION")          ' 位置
        .SMPKBN = rs("SMPKBN")              ' サンプル区分
        .TRANCOND = rs("TRANCOND")          ' 処理条件
        .TRANCNT = rs("TRANCNT")            ' 処理回数
        .SMPLNO = rs("SMPLNO")              ' サンプルＮｏ
        .SMPLUMU = rs("SMPLUMU")            ' サンプル有無
        .MSRSDEN = rs("MSRSDEN")            ' 測定結果 Den
        .MSRSLDL = rs("MSRSLDL")            ' 測定結果 L/DL
        .MSRSDVD2 = rs("MSRSDVD2")          ' 測定結果 DVD2
        .MS01LDL1 = rs("MS01LDL1")          ' 測定値01 L/DL1
        .MS01LDL2 = rs("MS01LDL2")          ' 測定値01 L/DL2
        .MS01LDL3 = rs("MS01LDL3")          ' 測定値01 L/DL3
        .MS01LDL4 = rs("MS01LDL4")          ' 測定値01 L/DL4
        .MS01LDL5 = rs("MS01LDL5")          ' 測定値01 L/DL5
        .MS01DEN1 = rs("MS01DEN1")          ' 測定値01 Den1
        .MS01DEN2 = rs("MS01DEN2")          ' 測定値01 Den2
        .MS01DEN3 = rs("MS01DEN3")          ' 測定値01 Den3
        .MS01DEN4 = rs("MS01DEN4")          ' 測定値01 Den4
        .MS01DEN5 = rs("MS01DEN5")          ' 測定値01 Den5
        .MS02LDL1 = rs("MS02LDL1")          ' 測定値02 L/DL1
        .MS02LDL2 = rs("MS02LDL2")          ' 測定値02 L/DL2
        .MS02LDL3 = rs("MS02LDL3")          ' 測定値02 L/DL3
        .MS02LDL4 = rs("MS02LDL4")          ' 測定値02 L/DL4
        .MS02LDL5 = rs("MS02LDL5")          ' 測定値02 L/DL5
        .MS02DEN1 = rs("MS02DEN1")          ' 測定値02 Den1
        .MS02DEN2 = rs("MS02DEN2")          ' 測定値02 Den2
        .MS02DEN3 = rs("MS02DEN3")          ' 測定値02 Den3
        .MS02DEN4 = rs("MS02DEN4")          ' 測定値02 Den4
        .MS02DEN5 = rs("MS02DEN5")          ' 測定値02 Den5
        .MS03LDL1 = rs("MS03LDL1")          ' 測定値03 L/DL1
        .MS03LDL2 = rs("MS03LDL2")          ' 測定値03 L/DL2
        .MS03LDL3 = rs("MS03LDL3")          ' 測定値03 L/DL3
        .MS03LDL4 = rs("MS03LDL4")          ' 測定値03 L/DL4
        .MS03LDL5 = rs("MS03LDL5")          ' 測定値03 L/DL5
        .MS03DEN1 = rs("MS03DEN1")          ' 測定値03 Den1
        .MS03DEN2 = rs("MS03DEN2")          ' 測定値03 Den2
        .MS03DEN3 = rs("MS03DEN3")          ' 測定値03 Den3
        .MS03DEN4 = rs("MS03DEN4")          ' 測定値03 Den4
        .MS03DEN5 = rs("MS03DEN5")          ' 測定値03 Den5
        .MS04LDL1 = rs("MS04LDL1")          ' 測定値04 L/DL1
        .MS04LDL2 = rs("MS04LDL2")          ' 測定値04 L/DL2
        .MS04LDL3 = rs("MS04LDL3")          ' 測定値04 L/DL3
        .MS04LDL4 = rs("MS04LDL4")          ' 測定値04 L/DL4
        .MS04LDL5 = rs("MS04LDL5")          ' 測定値04 L/DL5
        .MS04DEN1 = rs("MS04DEN1")          ' 測定値04 Den1
        .MS04DEN2 = rs("MS04DEN2")          ' 測定値04 Den2
        .MS04DEN3 = rs("MS04DEN3")          ' 測定値04 Den3
        .MS04DEN4 = rs("MS04DEN4")          ' 測定値04 Den4
        .MS04DEN5 = rs("MS04DEN5")          ' 測定値04 Den5
        .MS05LDL1 = rs("MS05LDL1")          ' 測定値05 L/DL1
        .MS05LDL2 = rs("MS05LDL2")          ' 測定値05 L/DL2
        .MS05LDL3 = rs("MS05LDL3")          ' 測定値05 L/DL3
        .MS05LDL4 = rs("MS05LDL4")          ' 測定値05 L/DL4
        .MS05LDL5 = rs("MS05LDL5")          ' 測定値05 L/DL5
        .MS05DEN1 = rs("MS05DEN1")          ' 測定値05 Den1
        .MS05DEN2 = rs("MS05DEN2")          ' 測定値05 Den2
        .MS05DEN3 = rs("MS05DEN3")          ' 測定値05 Den3
        .MS05DEN4 = rs("MS05DEN4")          ' 測定値05 Den4
        .MS05DEN5 = rs("MS05DEN5")          ' 測定値05 Den5
        .MS06LDL1 = rs("MS06LDL1")          ' 測定値06 L/DL1
        .MS06LDL2 = rs("MS06LDL2")          ' 測定値06 L/DL2
        .MS06LDL3 = rs("MS06LDL3")          ' 測定値06 L/DL3
        .MS06LDL4 = rs("MS06LDL4")          ' 測定値06 L/DL4
        .MS06LDL5 = rs("MS06LDL5")          ' 測定値06 L/DL5
        .MS06DEN1 = rs("MS06DEN1")          ' 測定値06 Den1
        .MS06DEN2 = rs("MS06DEN2")          ' 測定値06 Den2
        .MS06DEN3 = rs("MS06DEN3")          ' 測定値06 Den3
        .MS06DEN4 = rs("MS06DEN4")          ' 測定値06 Den4
        .MS06DEN5 = rs("MS06DEN5")          ' 測定値06 Den5
        .MS07LDL1 = rs("MS07LDL1")          ' 測定値07 L/DL1
        .MS07LDL2 = rs("MS07LDL2")          ' 測定値07 L/DL2
        .MS07LDL3 = rs("MS07LDL3")          ' 測定値07 L/DL3
        .MS07LDL4 = rs("MS07LDL4")          ' 測定値07 L/DL4
        .MS07LDL5 = rs("MS07LDL5")          ' 測定値07 L/DL5
        .MS07DEN1 = rs("MS07DEN1")          ' 測定値07 Den1
        .MS07DEN2 = rs("MS07DEN2")          ' 測定値07 Den2
        .MS07DEN3 = rs("MS07DEN3")          ' 測定値07 Den3
        .MS07DEN4 = rs("MS07DEN4")          ' 測定値07 Den4
        .MS07DEN5 = rs("MS07DEN5")          ' 測定値07 Den5
        .MS08LDL1 = rs("MS08LDL1")          ' 測定値08 L/DL1
        .MS08LDL2 = rs("MS08LDL2")          ' 測定値08 L/DL2
        .MS08LDL3 = rs("MS08LDL3")          ' 測定値08 L/DL3
        .MS08LDL4 = rs("MS08LDL4")          ' 測定値08 L/DL4
        .MS08LDL5 = rs("MS08LDL5")          ' 測定値08 L/DL5
        .MS08DEN1 = rs("MS08DEN1")          ' 測定値08 Den1
        .MS08DEN2 = rs("MS08DEN2")          ' 測定値08 Den2
        .MS08DEN3 = rs("MS08DEN3")          ' 測定値08 Den3
        .MS08DEN4 = rs("MS08DEN4")          ' 測定値08 Den4
        .MS08DEN5 = rs("MS08DEN5")          ' 測定値08 Den5
        .MS09LDL1 = rs("MS09LDL1")          ' 測定値09 L/DL1
        .MS09LDL2 = rs("MS09LDL2")          ' 測定値09 L/DL2
        .MS09LDL3 = rs("MS09LDL3")          ' 測定値09 L/DL3
        .MS09LDL4 = rs("MS09LDL4")          ' 測定値09 L/DL4
        .MS09LDL5 = rs("MS09LDL5")          ' 測定値09 L/DL5
        .MS09DEN1 = rs("MS09DEN1")          ' 測定値09 Den1
        .MS09DEN2 = rs("MS09DEN2")          ' 測定値09 Den2
        .MS09DEN3 = rs("MS09DEN3")          ' 測定値09 Den3
        .MS09DEN4 = rs("MS09DEN4")          ' 測定値09 Den4
        .MS09DEN5 = rs("MS09DEN5")          ' 測定値09 Den5
        .MS10LDL1 = rs("MS10LDL1")          ' 測定値10 L/DL1
        .MS10LDL2 = rs("MS10LDL2")          ' 測定値10 L/DL2
        .MS10LDL3 = rs("MS10LDL3")          ' 測定値10 L/DL3
        .MS10LDL4 = rs("MS10LDL4")          ' 測定値10 L/DL4
        .MS10LDL5 = rs("MS10LDL5")          ' 測定値10 L/DL5
        .MS10DEN1 = rs("MS10DEN1")          ' 測定値10 Den1
        .MS10DEN2 = rs("MS10DEN2")          ' 測定値10 Den2
        .MS10DEN3 = rs("MS10DEN3")          ' 測定値10 Den3
        .MS10DEN4 = rs("MS10DEN4")          ' 測定値10 Den4
        .MS10DEN5 = rs("MS10DEN5")          ' 測定値10 Den5
        .MS11LDL1 = rs("MS11LDL1")          ' 測定値11 L/DL1
        .MS11LDL2 = rs("MS11LDL2")          ' 測定値11 L/DL2
        .MS11LDL3 = rs("MS11LDL3")          ' 測定値11 L/DL3
        .MS11LDL4 = rs("MS11LDL4")          ' 測定値11 L/DL4
        .MS11LDL5 = rs("MS11LDL5")          ' 測定値11 L/DL5
        .MS11DEN1 = rs("MS11DEN1")          ' 測定値11 Den1
        .MS11DEN2 = rs("MS11DEN2")          ' 測定値11 Den2
        .MS11DEN3 = rs("MS11DEN3")          ' 測定値11 Den3
        .MS11DEN4 = rs("MS11DEN4")          ' 測定値11 Den4
        .MS11DEN5 = rs("MS11DEN5")          ' 測定値11 Den5
        .MS12LDL1 = rs("MS12LDL1")          ' 測定値12 L/DL1
        .MS12LDL2 = rs("MS12LDL2")          ' 測定値12 L/DL2
        .MS12LDL3 = rs("MS12LDL3")          ' 測定値12 L/DL3
        .MS12LDL4 = rs("MS12LDL4")          ' 測定値12 L/DL4
        .MS12LDL5 = rs("MS12LDL5")          ' 測定値12 L/DL5
        .MS12DEN1 = rs("MS12DEN1")          ' 測定値12 Den1
        .MS12DEN2 = rs("MS12DEN2")          ' 測定値12 Den2
        .MS12DEN3 = rs("MS12DEN3")          ' 測定値12 Den3
        .MS12DEN4 = rs("MS12DEN4")          ' 測定値12 Den4
        .MS12DEN5 = rs("MS12DEN5")          ' 測定値12 Den5
        .MS13LDL1 = rs("MS13LDL1")          ' 測定値13 L/DL1
        .MS13LDL2 = rs("MS13LDL2")          ' 測定値13 L/DL2
        .MS13LDL3 = rs("MS13LDL3")          ' 測定値13 L/DL3
        .MS13LDL4 = rs("MS13LDL4")          ' 測定値13 L/DL4
        .MS13LDL5 = rs("MS13LDL5")          ' 測定値13 L/DL5
        .MS13DEN1 = rs("MS13DEN1")          ' 測定値13 Den1
        .MS13DEN2 = rs("MS13DEN2")          ' 測定値13 Den2
        .MS13DEN3 = rs("MS13DEN3")          ' 測定値13 Den3
        .MS13DEN4 = rs("MS13DEN4")          ' 測定値13 Den4
        .MS13DEN5 = rs("MS13DEN5")          ' 測定値13 Den5
        .MS14LDL1 = rs("MS14LDL1")          ' 測定値14 L/DL1
        .MS14LDL2 = rs("MS14LDL2")          ' 測定値14 L/DL2
        .MS14LDL3 = rs("MS14LDL3")          ' 測定値14 L/DL3
        .MS14LDL4 = rs("MS14LDL4")          ' 測定値14 L/DL4
        .MS14LDL5 = rs("MS14LDL5")          ' 測定値14 L/DL5
        .MS14DEN1 = rs("MS14DEN1")          ' 測定値14 Den1
        .MS14DEN2 = rs("MS14DEN2")          ' 測定値14 Den2
        .MS14DEN3 = rs("MS14DEN3")          ' 測定値14 Den3
        .MS14DEN4 = rs("MS14DEN4")          ' 測定値14 Den4
        .MS14DEN5 = rs("MS14DEN5")          ' 測定値14 Den5
        .MS15LDL1 = rs("MS15LDL1")          ' 測定値15 L/DL1
        .MS15LDL2 = rs("MS15LDL2")          ' 測定値15 L/DL2
        .MS15LDL3 = rs("MS15LDL3")          ' 測定値15 L/DL3
        .MS15LDL4 = rs("MS15LDL4")          ' 測定値15 L/DL4
        .MS15LDL5 = rs("MS15LDL5")          ' 測定値15 L/DL5
        .MS15DEN1 = rs("MS15DEN1")          ' 測定値15 Den1
        .MS15DEN2 = rs("MS15DEN2")          ' 測定値15 Den2
        .MS15DEN3 = rs("MS15DEN3")          ' 測定値15 Den3
        .MS15DEN4 = rs("MS15DEN4")          ' 測定値15 Den4
        .MS15DEN5 = rs("MS15DEN5")          ' 測定値15 Den5
        If Not IsNull(rs("MS01DVD2")) Then .MS01DVD2 = rs("MS01DVD2")      ' 測定値01 DVD2
        If Not IsNull(rs("MS02DVD2")) Then .MS02DVD2 = rs("MS02DVD2")      ' 測定値02 DVD2
        If Not IsNull(rs("MS03DVD2")) Then .MS03DVD2 = rs("MS03DVD2")      ' 測定値03 DVD2
        If Not IsNull(rs("MS04DVD2")) Then .MS04DVD2 = rs("MS04DVD2")      ' 測定値04 DVD2
        If Not IsNull(rs("MS05DVD2")) Then .MS05DVD2 = rs("MS05DVD2")      ' 測定値05 DVD2
        
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
        .MSZEROMN = fncNullCheck(rs("MSZEROMN"))    ' L/DL0連続数最小値
        .MSZEROMX = fncNullCheck(rs("MSZEROMX"))    ' L/DL0連続数最大値
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
        
        .REGDATE = rs("REGDATE")            ' 登録日付
    End With
    
    funGetCryGDJisseki = 0

proc_exit:
    '終了
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' 結晶LT実績取得関数
'------------------------------------------------

'概要      :結晶番号、サンプルＩＤから、TBCMJ007を検索し、結晶LT実績値を取得し返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                               :説明
'          :sCryNum       ,I  ,String                           :結晶番号
'          :iSmplID       ,I  ,Long                             :サンプルＩＤ   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :tCryLT        ,O  ,type_DBDRV_scmzc_fcmkc001c_LT    :結晶LT実績(構造体)
'          :戻り値        ,O  ,Integer                          :取得結果 = 0 : 正常
'                                                                          -1 : 異常
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funGetCryLTJisseki(sCryNum As String, iSmplID As Long, tCryLT As type_DBDRV_scmzc_fcmkc001c_LT) As Integer

    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryLTJisseki"
    
    '結晶番号、サンプルＩＤからTBCMJ007の結晶LT実績値を検索する。

    '2005/12/02 mod SET高崎 LT測定値10点、NULL化対応 ->
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, MEASPEAK, CALCMEAS, REGDATE, "
    sql = sql & "NVL(MEAS1, -1) MEAS1,"
    sql = sql & "NVL(MEAS2, -1) MEAS2,"
    sql = sql & "NVL(MEAS3, -1) MEAS3,"
    sql = sql & "NVL(MEAS4, -1) MEAS4,"
    sql = sql & "NVL(MEAS5, -1) MEAS5,"
    sql = sql & "NVL(MEAS6, -1) MEAS6,"
    sql = sql & "NVL(MEAS7, -1) MEAS7,"
    sql = sql & "NVL(MEAS8, -1) MEAS8,"
    sql = sql & "NVL(MEAS9, -1) MEAS9,"
    sql = sql & "NVL(MEAS10, -1) MEAS10,"
    sql = sql & "LTSPIFLG "
    sql = sql & ",NVL(CONVAL, -1) CONVAL "
    sql = sql & "from TBCMJ007 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    '2005/12/02 mod SET高崎 LT測定値10点、NULL化対応 <-
    
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY)
    
    '該当データなし
    If rs.EOF Then
        funGetCryLTJisseki = -1
        GoTo proc_exit
    End If
    
     ''抽出結果を格納する
    With tCryLT
        .CRYNUM = rs("CRYNUM")              ' 結晶番号
        .POSITION = rs("POSITION")          ' 位置
        .SMPKBN = rs("SMPKBN")              ' サンプル区分
        .TRANCOND = rs("TRANCOND")          ' 処理条件
        .TRANCNT = rs("TRANCNT")            ' 処理回数
        .SMPLNO = rs("SMPLNO")              ' サンプルＮｏ
        .SMPLUMU = rs("SMPLUMU")            ' サンプル有無
        .MEAS1 = rs("MEAS1")                ' 測定値１
        .MEAS2 = rs("MEAS2")                ' 測定値２
        .MEAS3 = rs("MEAS3")                ' 測定値３
        .MEAS4 = rs("MEAS4")                ' 測定値４
        .MEAS5 = rs("MEAS5")                ' 測定値５
        .MEASPEAK = rs("MEASPEAK")          ' 測定値 ピーク値
        .CALCMEAS = rs("CALCMEAS")          ' 計算結果
        .REGDATE = rs("REGDATE")            ' 登録日付
''Add Start 2011/07/22 LT10Ω判定追加対応 T.Koi(SETsw)
        .CONVAL = rs.Fields("CONVAL")       '10Ω換算値
''Add End   2011/07/22 LT10Ω判定追加対応 T.Koi(SETsw)
        '2005/12/02 add SET高崎 測定値６～１０カラム追加のため追加 ->
        '                       判定フラグカラム追加
        .MEAS6 = rs("MEAS6")            ' 測定値６
        .MEAS7 = rs("MEAS7")            ' 測定値７
        .MEAS8 = rs("MEAS8")            ' 測定値８
        .MEAS9 = rs("MEAS9")            ' 測定値９
        .MEAS10 = rs("MEAS10")          ' 測定値１０
        .LTSPIFLG = Trim(CStr(NulltoStr(rs.Fields("LTSPIFLG").Value)))  '測定位置判定フラグ
        
        '2005/12/02 add SET高崎 測定値６～１０カラム追加のため追加 <-
        '                       判定フラグカラム追加
    End With
    
    funGetCryLTJisseki = 0

proc_exit:
    '終了
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' 結晶EPD実績取得関数
'------------------------------------------------

'概要      :結晶番号、サンプルＩＤから、TBCMJ001を検索し、結晶EPD実績値を取得し返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                               :説明
'          :sCryNum       ,I  ,String                           :結晶番号
'          :iSmplID       ,I  ,Long                             :サンプルＩＤ   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'          :tCryEPD       ,O  ,type_DBDRV_scmzc_fcmkc001c_EPD   :結晶EPD実績(構造体)
'          :戻り値        ,O  ,Integer                          :取得結果 = 0 : 正常
'                                                                          -1 : 異常
'説明      :
'履歴      :2003/09/05 新規作成　システムブレイン

Public Function funGetCryEPDJisseki(sCryNum As String, iSmplID As Long, tCryEPD As type_DBDRV_scmzc_fcmkc001c_EPD) As Integer

    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryEPDJisseki"
    
    '結晶番号、サンプルＩＤからTBCMJ001の結晶EPD実績値を検索する。
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, MEASURE, REGDATE "
    sql = sql & "from TBCMJ001 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetCryEPDJisseki = -1
        GoTo proc_exit
    End If
    
     ''抽出結果を格納する
    With tCryEPD
        .CRYNUM = rs("CRYNUM")              ' 結晶番号
        .POSITION = rs("POSITION")          ' 位置
        .SMPKBN = rs("SMPKBN")              ' サンプル区分
        .TRANCOND = rs("TRANCOND")          ' 処理条件
        .TRANCNT = rs("TRANCNT")            ' 処理回数
        .SMPLNO = rs("SMPLNO")              ' サンプルＮｏ
        .SMPLUMU = rs("SMPLUMU")            ' サンプル有無
        .MEASURE = rs("MEASURE")            ' 測定値
        .REGDATE = rs("REGDATE")            ' 登録日付
    End With
    
    funGetCryEPDJisseki = 0

proc_exit:
    '終了
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
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
' 結晶実測値チェック
'------------------------------------------------

'概要      :結晶実測チェックを行ない結果を返す。
'説明      :
'履歴      :TEST2004/10

Public Function funChkJissoku(tFullHin As tFullHinban, tCryRs As type_DBDRV_scmzc_fcmkc001c_CryR) As Boolean
    Dim tSiyou          As type_DBDRV_scmzc_fcmkc001c_Siyou
    Dim retJudg         As Boolean
    Dim sCryRs As type_DBDRV_scmzc_fcmkc001c_CryR
    
    If funGet_TBCME018(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then
        funChkJissoku = False
        Exit Function
    End If
    sCryRs = tCryRs
    '総合判定を行なう
    '測定点数、保証は全点に固定(全ての値をチェックするため）
    tSiyou.HSXRSPOT = "5"
    tSiyou.HSXRHWYT = "3"
    If Not CrResJudg(0, tSiyou, sCryRs, retJudg, 1) Then
        funChkJissoku = False
        Exit Function
    End If
    If retJudg = False Then
        funChkJissoku = False
        Exit Function
    End If
    funChkJissoku = True
    
End Function

'Add Start 2011/01/31 SMPK Miyata
'------------------------------------------------
' 結晶C実績取得関数
'------------------------------------------------
'概要      :結晶番号、サンプルＩＤから、TBCMJ023を検索し、結晶C実績値を取得し返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                               :説明
'          :sCryNum       ,I  ,String                           :結晶番号
'          :iSmplID       ,I  ,Long                             :サンプルＩＤ
'          :tCryC         ,O  ,type_DBDRV_scmzc_fcmkc001c_C     :結晶C実績(構造体)
'          :戻り値        ,O  ,Integer                          :取得結果 = 0 : 正常
'                                                                          -1 : 異常
'説明      :
'履歴      :
Public Function funGetCryCJisseki(sCryNum As String, iSmplID As Long, tCryC As type_DBDRV_scmzc_fcmkc001c_C) As Integer

    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryCJisseki"
    
    '結晶番号、サンプルＩＤからTBCMJ023の結晶C実績値を検索する。
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCNT, SMPLNO, "
    sql = sql & "SMPLUMUC, CPTNJSK, CDISKJSK, CRINGNKJSK, CRINGGKJSK, CHANTEI, "
    sql = sql & "REGDATE "
    sql = sql & "from TBCMJ023 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetCryCJisseki = -1
        GoTo proc_exit
    End If
    
     ''抽出結果を格納する
    With tCryC
        .CRYNUM = rs("CRYNUM")              ' 結晶番号
        .POSITION = rs("POSITION")          ' 位置
        .SMPKBN = rs("SMPKBN")              ' サンプル区分
        .TRANCNT = rs("TRANCNT")            ' 処理回数

        If IsNull(rs("SMPLNO")) = False Then .SMPLNO = rs("SMPLNO")             ' サンプルＮｏ
        If IsNull(rs("SMPLUMUC")) = False Then .SMPLUMUC = rs("SMPLUMUC")       ' サンプル有無（C）
        If IsNull(rs("CPTNJSK")) = False Then .CPTNJSK = rs("CPTNJSK")          ' C パターン実績
        If IsNull(rs("CDISKJSK")) = False Then .CDISKJSK = rs("CDISKJSK")       ' C Disk半径実績
        If IsNull(rs("CRINGNKJSK")) = False Then .CRINGNKJSK = rs("CRINGNKJSK") ' C Ring内径実績
        If IsNull(rs("CRINGGKJSK")) = False Then .CRINGGKJSK = rs("CRINGGKJSK") ' C Ring外径実績
        If IsNull(rs("CHANTEI")) = False Then .CHANTEI = rs("CHANTEI")          ' C 判定結果
        If IsNull(rs("REGDATE")) = False Then .REGDATE = rs("REGDATE")          ' 登録日付
    End With

    funGetCryCJisseki = 0

proc_exit:
    '終了
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' 結晶CJ実績取得関数
'------------------------------------------------
'概要      :結晶番号、サンプルＩＤから、TBCMJ023を検索し、結晶CJ実績値を取得し返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                               :説明
'          :sCryNum       ,I  ,String                           :結晶番号
'          :iSmplID       ,I  ,Long                             :サンプルＩＤ
'          :tCryCJ        ,O  ,type_DBDRV_scmzc_fcmkc001c_CJ    :結晶CJ実績(構造体)
'          :戻り値        ,O  ,Integer                          :取得結果 = 0 : 正常
'                                                                          -1 : 異常
'説明      :
'履歴      :
Public Function funGetCryCJJisseki(sCryNum As String, iSmplID As Long, tCryCJ As type_DBDRV_scmzc_fcmkc001c_CJ) As Integer

    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryCJJisseki"
    
    '結晶番号、サンプルＩＤからTBCMJ023の結晶C実績値を検索する。
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCNT, SMPLNO, "
    sql = sql & "SMPLUMUCJ, CJPTNJSK, CJDISKJSK, CJRINGNKJSK, CJRINGGKJSK, CJBANDNKJSK, CJBANDGKJSK, CJRINGCALC, "
    sql = sql & "CJPICALC , CJHANTEI, CJDMAXPIC5, CJRMAXPIC5, CJDRMAXPIC5, CJALLMAXDIC5, CJALLMINRINC5, CJALLMAXRIGC5, "
    sql = sql & "REGDATE "
    sql = sql & "from TBCMJ023 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    
    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetCryCJJisseki = -1
        GoTo proc_exit
    End If
    
     ''抽出結果を格納する
    With tCryCJ
        .CRYNUM = rs("CRYNUM")              ' 結晶番号
        .POSITION = rs("POSITION")          ' 位置
        .SMPKBN = rs("SMPKBN")              ' サンプル区分
        .TRANCNT = rs("TRANCNT")            ' 処理回数

        If IsNull(rs("SMPLNO")) = False Then .SMPLNO = rs("SMPLNO")                 ' サンプルＮｏ
        If IsNull(rs("SMPLUMUCJ")) = False Then .SMPLUMUCJ = rs("SMPLUMUCJ")        ' サンプル有無（CJ）
        If IsNull(rs("CJPTNJSK")) = False Then .CJPTNJSK = rs("CJPTNJSK")           ' CJ パターン実績
        If IsNull(rs("CJDISKJSK")) = False Then .CJDISKJSK = rs("CJDISKJSK")        ' CJ Disk半径実績
        If IsNull(rs("CJRINGNKJSK")) = False Then .CJRINGNKJSK = rs("CJRINGNKJSK")  ' CJ Ring内径実績
        If IsNull(rs("CJRINGGKJSK")) = False Then .CJRINGGKJSK = rs("CJRINGGKJSK")  ' CJ Ring外径実績
        If IsNull(rs("CJBANDNKJSK")) = False Then .CJBANDNKJSK = rs("CJBANDNKJSK")  ' CJ Band内径実績
        If IsNull(rs("CJBANDGKJSK")) = False Then .CJBANDGKJSK = rs("CJBANDGKJSK")  ' CJ Band外径実績
        If IsNull(rs("CJRINGCALC")) = False Then .CJRINGCALC = rs("CJRINGCALC")     ' CJ Ring幅計算
        If IsNull(rs("CJPICALC")) = False Then .CJPICALC = rs("CJPICALC")           ' CJ Pi幅計算
        If IsNull(rs("CJHANTEI")) = False Then .CJHANTEI = rs("CJHANTEI")           ' CJ 判定結果
        If IsNull(rs("CJDMAXPIC5")) = False Then .CJDMAXPIC5 = rs("CJDMAXPIC5")     ' CJ Diskのみパターン Pi幅上限値
        If IsNull(rs("CJRMAXPIC5")) = False Then .CJRMAXPIC5 = rs("CJRMAXPIC5")     ' CJ Ringのみパターン Pi幅上限値
        If IsNull(rs("CJDRMAXPIC5")) = False Then .CJDRMAXPIC5 = rs("CJDRMAXPIC5")          ' CJ DiskRingパターン Pi幅上限値
        If IsNull(rs("CJALLMAXDIC5")) = False Then .CJALLMAXDIC5 = rs("CJALLMAXDIC5")       ' CJ 共通Disk半径上限値
        If IsNull(rs("CJALLMINRINC5")) = False Then .CJALLMINRINC5 = rs("CJALLMINRINC5")    ' CJ 共通Ring内径下限値
        If IsNull(rs("CJALLMAXRIGC5")) = False Then .CJALLMAXRIGC5 = rs("CJALLMAXRIGC5")    ' CJ 共通Ring外径上限値
        If IsNull(rs("REGDATE")) = False Then .REGDATE = rs("REGDATE")                      ' 登録日付
    End With

    funGetCryCJJisseki = 0

proc_exit:
    '終了
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' 結晶CJLT実績取得関数
'------------------------------------------------
'概要      :結晶番号、サンプルＩＤから、TBCMJ023を検索し、結晶CJLT実績値を取得し返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                               :説明
'          :sCryNum       ,I  ,String                           :結晶番号
'          :iSmplID       ,I  ,Long                             :サンプルＩＤ
'          :tCryCJLT      ,O  ,type_DBDRV_scmzc_fcmkc001c_CJLT  :結晶CJLT実績(構造体)
'          :戻り値        ,O  ,Integer                          :取得結果 = 0 : 正常
'                                                                          -1 : 異常
'説明      :
'履歴      :
Public Function funGetCryCJLTJisseki(sCryNum As String, iSmplID As Long, tCryCJLT As type_DBDRV_scmzc_fcmkc001c_CJLT) As Integer

    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryCJLTJisseki"
    
    '結晶番号、サンプルＩＤからTBCMJ023の結晶C実績値を検索する。
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCNT, SMPLNO, "
    sql = sql & "SMPLUMUCJLT, CJLTPTNJSK, CJLTDISKJSK, CJLTRINGNKJSK, CJLTRINGGKJSK, "
    sql = sql & "CJLTBANDNKJSK , CJLTBANDGKJSK, CJLTRINGCALC, CJLTPICALC, CJLTHANTEI, "
    sql = sql & "REGDATE "
    sql = sql & "from TBCMJ023 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"


    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetCryCJLTJisseki = -1
        GoTo proc_exit
    End If
    
     ''抽出結果を格納する
    With tCryCJLT
        .CRYNUM = rs("CRYNUM")              ' 結晶番号
        .POSITION = rs("POSITION")          ' 位置
        .SMPKBN = rs("SMPKBN")              ' サンプル区分
        .TRANCNT = rs("TRANCNT")            ' 処理回数

        If IsNull(rs("SMPLNO")) = False Then .SMPLNO = rs("SMPLNO")                 ' サンプルＮｏ
        If IsNull(rs("SMPLUMUCJLT")) = False Then .SMPLUMUCJLT = rs("SMPLUMUCJLT")    ' サンプル有無（CJ(LT)）
        
        If IsNull(rs("CJLTPTNJSK")) = False Then .CJLTPTNJSK = rs("CJLTPTNJSK")     ' CJ(LT) パターン実績
        If IsNull(rs("CJLTDISKJSK")) = False Then .CJLTDISKJSK = rs("CJLTDISKJSK")  ' CJ(LT) Disk半径実績
        If IsNull(rs("CJLTRINGNKJSK")) = False Then .CJLTRINGNKJSK = rs("CJLTRINGNKJSK")  ' CJ(LT) Ring内径実績
        If IsNull(rs("CJLTRINGGKJSK")) = False Then .CJLTRINGGKJSK = rs("CJLTRINGGKJSK")  ' CJ(LT) Ring外径実績
        If IsNull(rs("CJLTBANDNKJSK")) = False Then .CJLTBANDNKJSK = rs("CJLTBANDNKJSK")  ' CJ(LT) Band内径実績
        If IsNull(rs("CJLTBANDGKJSK")) = False Then .CJLTBANDGKJSK = rs("CJLTBANDGKJSK")  ' CJ(LT) Band外径実績
        If IsNull(rs("CJLTRINGCALC")) = False Then .CJLTRINGCALC = rs("CJLTRINGCALC")     ' CJ(LT) Ring幅計算
        If IsNull(rs("CJLTPICALC")) = False Then .CJLTPICALC = rs("CJLTPICALC")           ' CJ(LT) Pi幅計算
        If IsNull(rs("CJLTHANTEI")) = False Then .CJLTHANTEI = rs("CJLTHANTEI")           ' CJ(LT) 判定結果
        If IsNull(rs("REGDATE")) = False Then .REGDATE = rs("REGDATE")                    ' 登録日付
    End With

    funGetCryCJLTJisseki = 0

proc_exit:
    '終了
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'------------------------------------------------
' 結晶CJ2実績取得関数
'------------------------------------------------
'概要      :結晶番号、サンプルＩＤから、TBCMJ023を検索し、結晶CJLT実績値を取得し返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                               :説明
'          :sCryNum       ,I  ,String                           :結晶番号
'          :iSmplID       ,I  ,Long                             :サンプルＩＤ
'          :tCryCJ2       ,O  ,type_DBDRV_scmzc_fcmkc001c_CJ2   :結晶CJ2実績(構造体)
'          :戻り値        ,O  ,Integer                          :取得結果 = 0 : 正常
'                                                                          -1 : 異常
'説明      :
'履歴      :
Public Function funGetCryCJ2Jisseki(sCryNum As String, iSmplID As Long, tCryCJ2 As type_DBDRV_scmzc_fcmkc001c_CJ2) As Integer

    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_CryHanSui.bas -- Function funGetCryCJ2Jisseki"
    
    '結晶番号、サンプルＩＤからTBCMJ023の結晶C実績値を検索する。
    sql = "select CRYNUM, POSITION, SMPKBN, TRANCNT, SMPLNO, "
    sql = sql & "SMPLUMUCJ2, CJ2PTNJSK, CJ2DISKJSK, CJ2RINGNKJSK, CJ2RINGGKJSK,CJ2PICALC, "
    sql = sql & "CJ2HANTEI , CJ2DMAXPIC5, CJ2RMAXPIC5, CJ2RMINRINC5, CJ2RMAXRIGC5, CJ2DRMAXPIC5, "
    sql = sql & "CJ2DRMINRINC5, CJ2DRMAXRIGC5, "
    sql = sql & "REGDATE "
    sql = sql & "from TBCMJ023 "
    sql = sql & "where CRYNUM = '" & sCryNum & "' and "
    sql = sql & "      SMPLNO = " & iSmplID & " "
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"

''
''
''

    'SQL文の実行
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Then
        funGetCryCJ2Jisseki = -1
        GoTo proc_exit
    End If
    
     ''抽出結果を格納する
    With tCryCJ2
        .CRYNUM = rs("CRYNUM")              ' 結晶番号
        .POSITION = rs("POSITION")          ' 位置
        .SMPKBN = rs("SMPKBN")              ' サンプル区分
        .TRANCNT = rs("TRANCNT")            ' 処理回数

        If IsNull(rs("SMPLNO")) = False Then .SMPLNO = rs("SMPLNO")                 ' サンプルＮｏ
        If IsNull(rs("SMPLUMUCJ2")) = False Then .SMPLUMUCJ2 = rs("SMPLUMUCJ2")     ' サンプル有無（CJ2）
        
        If IsNull(rs("CJ2PTNJSK")) = False Then .CJ2PTNJSK = rs("CJ2PTNJSK")        ' CJ2 パターン実績
        If IsNull(rs("CJ2DISKJSK")) = False Then .CJ2DISKJSK = rs("CJ2DISKJSK")     ' CJ2 Disk半径実績
        If IsNull(rs("CJ2RINGNKJSK")) = False Then .CJ2RINGNKJSK = rs("CJ2RINGNKJSK")   ' CJ2 Ring内径実績
        If IsNull(rs("CJ2RINGGKJSK")) = False Then .CJ2RINGGKJSK = rs("CJ2RINGGKJSK")   ' CJ2 Ring外径実績
        If IsNull(rs("CJ2PICALC")) = False Then .CJ2PICALC = rs("CJ2PICALC")            ' CJ2 Pi幅計算
        If IsNull(rs("CJ2HANTEI")) = False Then .CJ2HANTEI = rs("CJ2HANTEI")            ' CJ2 判定結果
        If IsNull(rs("CJ2DMAXPIC5")) = False Then .CJ2DMAXPIC5 = rs("CJ2DMAXPIC5")      ' CJ2 Diskのみパターン Pi幅下限値(MAXだが下限です)
        If IsNull(rs("CJ2RMAXPIC5")) = False Then .CJ2RMAXPIC5 = rs("CJ2RMAXPIC5")      ' CJ2 Ringのみパターン Pi幅下限値(MAXだが下限です)
        If IsNull(rs("CJ2RMINRINC5")) = False Then .CJ2RMINRINC5 = rs("CJ2RMINRINC5")   ' CJ2 Ringのみパターン Ring内径下限値
        If IsNull(rs("CJ2RMAXRIGC5")) = False Then .CJ2RMAXRIGC5 = rs("CJ2RMAXRIGC5")   ' CJ2 Ringのみパターン Ring外径上限値
        
        If IsNull(rs("CJ2DRMAXPIC5")) = False Then .CJ2DRMAXPIC5 = rs("CJ2DRMAXPIC5")   ' CJ2 DiskRingパターン Pi幅下限値(MAXだが下限です)
        If IsNull(rs("CJ2DRMINRINC5")) = False Then .CJ2DRMINRINC5 = rs("CJ2DRMINRINC5") ' CJ2 DiskRingパターン Ring内径下限値
        If IsNull(rs("CJ2DRMAXRIGC5")) = False Then .CJ2DRMAXRIGC5 = rs("CJ2DRMAXRIGC5") ' CJ2 DiskRingパターン Ring外径上限値
        If IsNull(rs("REGDATE")) = False Then .REGDATE = rs("REGDATE")                  ' 登録日付
    End With

    funGetCryCJ2Jisseki = 0

proc_exit:
    '終了
    Set rs = Nothing
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function
'Add End   2011/01/31 SMPK Miyata

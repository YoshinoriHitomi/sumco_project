Attribute VB_Name = "s_cmzcSPXL"
Option Explicit
'                                     2001/07/03
'===============================================================================
' 結晶所要サンプル判定関数
' 概要    :
'===============================================================================

Private tblSampUP As typ_SpSXLSamp
Private tblSampDN As typ_SpSXLSamp
Private tblSampTG As typ_SpSXLSamp

'概要      :結晶サンプルの取得（１ブロック分）
'ﾊﾟﾗﾒｰﾀ　　:変数名          ,IO ,型             ,説明
'　　      :pHinUp    　　　,I  ,tFullHinban  　,上品番テーブル
'　　      :pHinDn    　　　,I  ,tFullHinban  　,下品番テーブル
'      　　:pHinTg    　　　,I  ,tFullHinban  　,ねらい品番テーブル
'　　      :pSXLSample　　　,O  ,typ_SXLSample　,結晶サンプルテーブル
'　　      :戻り値          ,O  ,Integer      　,サンプル枚数
'説明      :検査指示サンプルデータを取得する
'履歴      :2001/07/03　大塚 作成
Public Function GetSXLSampAll(pHinUp As tFullHinban, pHinDn As tFullHinban, pHinTg As tFullHinban, pSXLSample As typ_SXLSample) As Integer

    '' 検査指示サンプルの取得
    With pSXLSample
        .CRYINDRS = GetSXLSamp(pHinUp, pHinDn, pHinTg, 1)
        .CRYINDOI = GetSXLSamp(pHinUp, pHinDn, pHinTg, 2)
        .CRYINDB1 = GetSXLSamp(pHinUp, pHinDn, pHinTg, 3)
        .CRYINDB2 = GetSXLSamp(pHinUp, pHinDn, pHinTg, 4)
        .CRYINDB3 = GetSXLSamp(pHinUp, pHinDn, pHinTg, 5)
        .CRYINDL1 = GetSXLSamp(pHinUp, pHinDn, pHinTg, 6)
        .CRYINDL2 = GetSXLSamp(pHinUp, pHinDn, pHinTg, 7)
        .CRYINDL3 = GetSXLSamp(pHinUp, pHinDn, pHinTg, 8)
        .CRYINDL4 = GetSXLSamp(pHinUp, pHinDn, pHinTg, 9)
        .CRYINDCS = GetSXLSamp(pHinUp, pHinDn, pHinTg, 10)
        .CRYINDGD = GetSXLSamp(pHinUp, pHinDn, pHinTg, 11)
        .CRYINDT = GetSXLSamp(pHinUp, pHinDn, pHinTg, 12)
        .CRYINDEP = GetSXLSamp(pHinUp, pHinDn, pHinTg, 13)
        .CRYINDX = GetSXLSamp(pHinUp, pHinDn, pHinTg, 14)       'X線測定 2009/07/24追加 SETsw kubota
        'Add Start 2010/12/10 SMPK Miyata
        .CRYINDC = GetSXLSamp(pHinUp, pHinDn, pHinTg, 15)       ' 検査項目(C)
        .CRYINDCJ = GetSXLSamp(pHinUp, pHinDn, pHinTg, 16)      ' 検査項目(CJ)
        .CRYINDCJLT = GetSXLSamp(pHinUp, pHinDn, pHinTg, 17)    ' 検査項目(CJ LT)
        .CRYINDCJ2 = GetSXLSamp(pHinUp, pHinDn, pHinTg, 18)     ' 検査項目(CJ2)
        'Add End   2010/12/10 SMPK Miyata
    End With

    '' サンプル枚数の取得
    GetSXLSampAll = GetSXLSampNum(pSXLSample)

End Function

'概要      :結晶サンプルの取得
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型           ,説明
'　　      :pHinUp　　　,I  ,tFullHinban　,上品番テーブル
'　　      :pHinDn　　　,I  ,tFullHinban　,下品番テーブル
'      　　:pHinTg　　　,I  ,tFullHinban　,ねらい品番テーブル
'　　      :iCol  　　　,I  ,Integer    　,列
'　　      :戻り値      ,O  ,String     　,検査指示サンプル
'説明      :検査指示サンプルデータを取得する
'履歴      :2001/07/03　大塚 作成
Public Function GetSXLSamp(pHinUp As tFullHinban, pHinDn As tFullHinban, pHinTg As tFullHinban, iCol As Integer) As String

    Dim HINBANUP As String
    Dim HINBANDN As String
    Dim iMode As Integer

    '' 上品番／下品番状態の分類
    HINBANUP = Trim(pHinUp.hinban)
    HINBANDN = Trim(pHinDn.hinban)
   
   If HINBANDN = "" Then
        iMode = 2
    ElseIf HINBANUP = "" Then
        iMode = 1
    ElseIf HINBANUP & pHinUp.Hinkubun = HINBANDN & pHinDn.Hinkubun Then
        iMode = 3
    Else
        iMode = 4
    End If

    '' 上品番の製品仕様データを取得
     If tblSampUP.HIN.hinban & tblSampUP.HIN.Hinkubun <> pHinUp.hinban & pHinUp.Hinkubun Then
        tblSampUP.HIN.hinban = pHinUp.hinban
        tblSampUP.HIN.mnorevno = pHinUp.mnorevno
        tblSampUP.HIN.Factory = pHinUp.Factory
        tblSampUP.HIN.OpeCond = pHinUp.OpeCond
        tblSampUP.HIN.Hinkubun = pHinUp.Hinkubun
        If tblSampUP.HIN.Hinkubun = "1" Or HINBANUP = "" Then
           Call GetSpecZ2(tblSampUP, pHinTg)
        ElseIf HINBANUP = "G" Or HINBANUP = "Z" Then
               Call GetSpecGZ(tblSampUP, pHinTg)
             Else
               If scmzc_getSXL(tblSampUP) = FUNCTION_RETURN_FAILURE Then
                 GetSXLSamp = "0"
                 Exit Function
              End If
        End If
     End If

    '' 下品番の製品仕様データを取得
     If tblSampDN.HIN.hinban & tblSampDN.HIN.Hinkubun <> pHinDn.hinban & pHinDn.Hinkubun Then
        tblSampDN.HIN.hinban = pHinDn.hinban
        tblSampDN.HIN.mnorevno = pHinDn.mnorevno
        tblSampDN.HIN.Factory = pHinDn.Factory
        tblSampDN.HIN.OpeCond = pHinDn.OpeCond
        tblSampDN.HIN.Hinkubun = pHinDn.Hinkubun
        If tblSampDN.HIN.Hinkubun = "1" Or HINBANDN = "" Then
           Call GetSpecZ2(tblSampDN, pHinTg)
        ElseIf HINBANDN = "G" Or HINBANDN = "Z" Then
               Call GetSpecGZ(tblSampDN, pHinTg)
             Else
               If scmzc_getSXL(tblSampDN) = FUNCTION_RETURN_FAILURE Then
                 GetSXLSamp = "0"
                 Exit Function
              End If
         End If
     End If

    '' 上品番／下品番状態分岐
    Select Case iMode
    Case 1      '' 上品番なし
        Select Case iCol
        Case 1      'Rs
            GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXRHWYS), "1", "0")
        Case 2      'Oi
            GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXONHWS), "1", "0")
        Case 3      'B1
            GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXBM1HS), "1", "0")
        Case 4      'B2
            GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXBM2HS), "1", "0")
        Case 5      'B3
            GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXBM3HS), "1", "0")
        Case 6      'L1
            GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXOF1HS), "1", "0")
        Case 7      'L2
            GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXOF2HS), "1", "0")
        Case 8      'L3
            GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXOF3HS), "1", "0")
        Case 9      'L4
            GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXOF4HS), "1", "0")
        Case 10     'Cs
'            GetSXLSamp = "0"
            'TOP/BOT保証対応 09/01/06 ooba
            GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCNHWS) And tblSampDN.CS_FROMTO, "1", "0")
        Case 11     'GD
            GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXDENHS) Or _
                             CheckHWS(tblSampDN.HSXLDLHS) Or _
                             CheckHWS(tblSampDN.HSXDVDHS), "1", "0")
        Case 12     'T
            GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXLTHWS), "1", "0")
        Case 13     'EPD
            GetSXLSamp = "0"
        Case 14     'X      '2009/07/24追加 SETsw kubota
            GetSXLSamp = "0"
        'Add Start 2010/12/10 SMPK Miyata
        Case 15     'C
            GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCHS), "1", "0")
        Case 16     'CJ
            GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCJHS), "1", "0")
        Case 17     'CJ LT
            GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCJLTHS), "1", "0")
        Case 18     'CJ2
            GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCJ2HS), "1", "0")
        'Add End   2010/12/10 SMPK Miyata

        End Select
    Case 2      '' 下品番なし
        Select Case iCol
        Case 1      'Rs
            GetSXLSamp = IIf(CheckHWS(tblSampUP.HSXRHWYS), "2", "0")
        Case 2      'Oi
            GetSXLSamp = IIf(CheckHWS(tblSampUP.HSXONHWS), "2", "0")
        Case 3      'B1
            GetSXLSamp = IIf(CheckHWS(tblSampUP.HSXBM1HS), "2", "0")
        Case 4      'B2
            GetSXLSamp = IIf(CheckHWS(tblSampUP.HSXBM2HS), "2", "0")
        Case 5      'B3
            GetSXLSamp = IIf(CheckHWS(tblSampUP.HSXBM3HS), "2", "0")
        Case 6      'L1
            GetSXLSamp = IIf(CheckHWS(tblSampUP.HSXOF1HS), "2", "0")
        Case 7      'L2
            GetSXLSamp = IIf(CheckHWS(tblSampUP.HSXOF2HS), "2", "0")
        Case 8      'L3
            GetSXLSamp = IIf(CheckHWS(tblSampUP.HSXOF3HS), "2", "0")
        Case 9      'L4
            GetSXLSamp = IIf(CheckHWS(tblSampUP.HSXOF4HS), "2", "0")
        Case 10     'Cs
'            GetSXLSamp = "2"
            'TOP/BOT保証対応 09/01/06 ooba
            GetSXLSamp = IIf(CheckHWS(tblSampUP.HSXCNHWS), "2", "0")
        Case 11     'GD
            GetSXLSamp = IIf(CheckHWS(tblSampUP.HSXDENHS) Or _
                             CheckHWS(tblSampUP.HSXLDLHS) Or _
                             CheckHWS(tblSampUP.HSXDVDHS), "2", "0")
        Case 12     'T
            GetSXLSamp = IIf(CheckHWS(tblSampUP.HSXLTHWS), "2", "0")
        Case 13     'EPD
            GetSXLSamp = "2"
        Case 14     'X      '2009/07/24追加 SETsw kubota
            GetSXLSamp = "2"
        'Add Start 2010/12/10 SMPK Miyata
        Case 15     'C
            GetSXLSamp = IIf(CheckHWS(tblSampUP.HSXCHS), "2", "0")
        Case 16     'CJ
            GetSXLSamp = IIf(CheckHWS(tblSampUP.HSXCJHS), "2", "0")
        Case 17     'CJ LT
            GetSXLSamp = IIf(CheckHWS(tblSampUP.HSXCJLTHS), "2", "0")
        Case 18     'CJ2
            GetSXLSamp = IIf(CheckHWS(tblSampUP.HSXCJ2HS), "2", "0")
        'Add End   2010/12/10 SMPK Miyata

        End Select
    Case 3      '' 上品番＝下品番
        Select Case iCol
        Case 1      'Rs
            If CheckHWS(tblSampUP.HSXRHWYS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXRHWYS), "3", "2")
            Else
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXRHWYS), "1", "0")
            End If
        Case 2      'Oi
            If CheckHWS(tblSampUP.HSXONHWS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXONHWS), "3", "2")
            Else
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXONHWS), "1", "0")
            End If
        Case 3      'B1
            If CheckHWS(tblSampUP.HSXBM1HS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXBM1HS), "3", "2")
            Else
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXBM1HS), "1", "0")
            End If
        Case 4      'B2
            If CheckHWS(tblSampUP.HSXBM2HS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXBM2HS), "3", "2")
            Else
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXBM2HS), "1", "0")
            End If
        Case 5      'B3
            If CheckHWS(tblSampUP.HSXBM3HS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXBM3HS), "3", "2")
            Else
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXBM3HS), "1", "0")
            End If
        Case 6      'L1
            If CheckHWS(tblSampUP.HSXOF1HS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXOF1HS), "3", "2")
            Else
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXOF1HS), "1", "0")
            End If
        Case 7      'L2
            If CheckHWS(tblSampUP.HSXOF2HS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXOF2HS), "3", "2")
            Else
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXOF2HS), "1", "0")
            End If
        Case 8      'L3
            If CheckHWS(tblSampUP.HSXOF3HS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXOF3HS), "3", "2")
            Else
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXOF3HS), "1", "0")
            End If
        Case 9      'L4
            If CheckHWS(tblSampUP.HSXOF4HS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXOF4HS), "3", "2")
            Else
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXOF4HS), "1", "0")
            End If
        Case 10     'Cs
'            GetSXLSamp = "0"
            'TOP/BOT保証対応 09/01/06 ooba
            If CheckHWS(tblSampUP.HSXCNHWS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCNHWS) And tblSampDN.CS_FROMTO, "3", "2")
            Else
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCNHWS) And tblSampDN.CS_FROMTO, "1", "0")
            End If
        Case 11     'GD
            If CheckHWS(tblSampUP.HSXDENHS) Or _
               CheckHWS(tblSampUP.HSXLDLHS) Or _
               CheckHWS(tblSampUP.HSXDVDHS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXDENHS) Or _
                                 CheckHWS(tblSampDN.HSXLDLHS) Or _
                                 CheckHWS(tblSampDN.HSXDVDHS), "3", "2")
            Else
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXDENHS) Or _
                                 CheckHWS(tblSampDN.HSXLDLHS) Or _
                                 CheckHWS(tblSampDN.HSXDVDHS), "1", "0")
            End If
        Case 12     'T
            If CheckHWS(tblSampUP.HSXLTHWS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXLTHWS), "3", "2")
            Else
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXLTHWS), "1", "0")
            End If
        Case 13     'EPD
            GetSXLSamp = "0"
        Case 14     'X      '2009/07/24追加 SETsw kubota
            GetSXLSamp = "0"
        'Add Start 2010/12/10 SMPK Miyata
        Case 15     'C
            If CheckHWS(tblSampUP.HSXCHS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCHS), "3", "2")
            Else
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCHS), "1", "0")
            End If
        Case 16     'CJ
            If CheckHWS(tblSampUP.HSXCJHS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCJHS), "3", "2")
            Else
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCJHS), "1", "0")
            End If
        Case 17     'CJ LT
            If CheckHWS(tblSampUP.HSXCJLTHS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCJLTHS), "3", "2")
            Else
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCJLTHS), "1", "0")
            End If
        Case 18     'CJ2
            If CheckHWS(tblSampUP.HSXCJ2HS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCJ2HS), "3", "2")
            Else
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCJ2HS), "1", "0")
            End If
        'Add End   2010/12/10 SMPK Miyata

        End Select
    Case 4      '' 上品番＜＞下品番
        Select Case iCol
        Case 1      'Rs
            If CheckHWS(tblSampUP.HSXRHWYS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXRHWYS), "3", "2")
            Else
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXRHWYS), "1", "0")
            End If
        Case 2      'Oi
            If Not CheckHWS(tblSampUP.HSXONHWS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXONHWS), "1", "0")
            ElseIf Not CheckHWS(tblSampDN.HSXONHWS) Then
                GetSXLSamp = "2"
            Else
                GetSXLSamp = IIf(tblSampUP.HSXONKWY = tblSampDN.HSXONKWY And _
                                 tblSampUP.HSXONSPH = tblSampDN.HSXONSPH And _
                                 tblSampUP.HSXONSPI = tblSampDN.HSXONSPI, "3", "4")
            End If
        Case 3      'B1
            If Not CheckHWS(tblSampUP.HSXBM1HS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXBM1HS), "1", "0")
            ElseIf Not CheckHWS(tblSampDN.HSXBM1HS) Then
                GetSXLSamp = "2"
            Else
                GetSXLSamp = IIf(tblSampUP.HSXBM1SH = tblSampDN.HSXBM1SH And _
                                 tblSampUP.HSXBM1ST = tblSampDN.HSXBM1ST And _
                                 tblSampUP.HSXBM1SR = tblSampDN.HSXBM1SR And _
                                 tblSampUP.HSXBM1NS = tblSampDN.HSXBM1NS And _
                                 tblSampUP.HSXBM1SZ = tblSampDN.HSXBM1SZ And _
                                 tblSampUP.HSXBM1ET = tblSampDN.HSXBM1ET, "3", "4")
            End If
        Case 4      'B2
            If Not CheckHWS(tblSampUP.HSXBM2HS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXBM2HS), "1", "0")
            ElseIf Not CheckHWS(tblSampDN.HSXBM2HS) Then
                GetSXLSamp = "2"
            Else
                GetSXLSamp = IIf(tblSampUP.HSXBM2SH = tblSampDN.HSXBM2SH And _
                                 tblSampUP.HSXBM2ST = tblSampDN.HSXBM2ST And _
                                 tblSampUP.HSXBM2SR = tblSampDN.HSXBM2SR And _
                                 tblSampUP.HSXBM2NS = tblSampDN.HSXBM2NS And _
                                 tblSampUP.HSXBM2SZ = tblSampDN.HSXBM2SZ And _
                                 tblSampUP.HSXBM2ET = tblSampDN.HSXBM2ET, "3", "4")
            End If
        Case 5      'B3
            If Not CheckHWS(tblSampUP.HSXBM3HS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXBM3HS), "1", "0")
            ElseIf Not CheckHWS(tblSampDN.HSXBM3HS) Then
                GetSXLSamp = "2"
            Else
                GetSXLSamp = IIf(tblSampUP.HSXBM3SH = tblSampDN.HSXBM3SH And _
                                 tblSampUP.HSXBM3ST = tblSampDN.HSXBM3ST And _
                                 tblSampUP.HSXBM3SR = tblSampDN.HSXBM3SR And _
                                 tblSampUP.HSXBM3NS = tblSampDN.HSXBM3NS And _
                                 tblSampUP.HSXBM3SZ = tblSampDN.HSXBM3SZ And _
                                 tblSampUP.HSXBM3ET = tblSampDN.HSXBM3ET, "3", "4")
            End If
        Case 6      'L1
            If Not CheckHWS(tblSampUP.HSXOF1HS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXOF1HS), "1", "0")
            ElseIf Not CheckHWS(tblSampDN.HSXOF1HS) Then
                GetSXLSamp = "2"
            Else
                GetSXLSamp = IIf(tblSampUP.HSXOF1SH = tblSampDN.HSXOF1SH And _
                                 tblSampUP.HSXOF1ST = tblSampDN.HSXOF1ST And _
                                 tblSampUP.HSXOF1SR = tblSampDN.HSXOF1SR And _
                                 tblSampUP.HSXOF1NS = tblSampDN.HSXOF1NS And _
                                 tblSampUP.HSXOF1SZ = tblSampDN.HSXOF1SZ And _
                                 tblSampUP.HSXOF1ET = tblSampDN.HSXOF1ET, "3", "4")
            End If
        Case 7      'L2
            If Not CheckHWS(tblSampUP.HSXOF2HS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXOF2HS), "1", "0")
            ElseIf Not CheckHWS(tblSampDN.HSXOF2HS) Then
                GetSXLSamp = "2"
            Else
                GetSXLSamp = IIf(tblSampUP.HSXOF2SH = tblSampDN.HSXOF2SH And _
                                 tblSampUP.HSXOF2ST = tblSampDN.HSXOF2ST And _
                                 tblSampUP.HSXOF2SR = tblSampDN.HSXOF2SR And _
                                 tblSampUP.HSXOF2NS = tblSampDN.HSXOF2NS And _
                                 tblSampUP.HSXOF2SZ = tblSampDN.HSXOF2SZ And _
                                 tblSampUP.HSXOF2ET = tblSampDN.HSXOF2ET, "3", "4")
            End If
        Case 8      'L3
            If Not CheckHWS(tblSampUP.HSXOF3HS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXOF3HS), "1", "0")
            ElseIf Not CheckHWS(tblSampDN.HSXOF3HS) Then
                GetSXLSamp = "2"
            Else
                GetSXLSamp = IIf(tblSampUP.HSXOF3SH = tblSampDN.HSXOF3SH And _
                                 tblSampUP.HSXOF3ST = tblSampDN.HSXOF3ST And _
                                 tblSampUP.HSXOF3SR = tblSampDN.HSXOF3SR And _
                                 tblSampUP.HSXOF3NS = tblSampDN.HSXOF3NS And _
                                 tblSampUP.HSXOF3SZ = tblSampDN.HSXOF3SZ And _
                                 tblSampUP.HSXOF3ET = tblSampDN.HSXOF3ET, "3", "4")
            End If
        Case 9      'L4
            If Not CheckHWS(tblSampUP.HSXOF4HS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXOF4HS), "1", "0")
            ElseIf Not CheckHWS(tblSampDN.HSXOF4HS) Then
                GetSXLSamp = "2"
            Else
                GetSXLSamp = IIf(tblSampUP.HSXOF4SH = tblSampDN.HSXOF4SH And _
                                 tblSampUP.HSXOF4ST = tblSampDN.HSXOF4ST And _
                                 tblSampUP.HSXOF4SR = tblSampDN.HSXOF4SR And _
                                 tblSampUP.HSXOF4NS = tblSampDN.HSXOF4NS And _
                                 tblSampUP.HSXOF4SZ = tblSampDN.HSXOF4SZ And _
                                 tblSampUP.HSXOF4ET = tblSampDN.HSXOF4ET, "3", "4")
            End If
        Case 10     'Cs
'            GetSXLSamp = "0"
            'TOP/BOT保証対応 09/01/06 ooba
            If CheckHWS(tblSampUP.HSXCNHWS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCNHWS) And tblSampDN.CS_FROMTO, "3", "2")
            Else
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCNHWS) And tblSampDN.CS_FROMTO, "1", "0")
            End If
        Case 11     'GD
            If Not CheckHWS(tblSampUP.HSXDENHS) And _
               Not CheckHWS(tblSampUP.HSXLDLHS) And _
               Not CheckHWS(tblSampUP.HSXDVDHS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXDENHS) Or _
                                 CheckHWS(tblSampDN.HSXLDLHS) Or _
                                 CheckHWS(tblSampDN.HSXDVDHS), "1", "0")
            ElseIf Not CheckHWS(tblSampDN.HSXDENHS) And _
                   Not CheckHWS(tblSampDN.HSXLDLHS) And _
                   Not CheckHWS(tblSampDN.HSXDVDHS) Then
                GetSXLSamp = "2"
            Else
                GetSXLSamp = "3"
            End If
        Case 12     'T
            If CheckHWS(tblSampUP.HSXLTHWS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXLTHWS), "3", "2")
            Else
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXLTHWS), "1", "0")
            End If
        Case 13     'EPD
            GetSXLSamp = "0"
        Case 14     'X      '2009/07/24追加 SETsw kubota
            GetSXLSamp = "0"
    'Add Start 2010/12/10 SMPK Miyata
        Case 15     'C
            If Not CheckHWS(tblSampUP.HSXCHS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCHS), "1", "0")
            ElseIf Not CheckHWS(tblSampDN.HSXCHS) Then
                GetSXLSamp = "2"
            Else
                GetSXLSamp = IIf(tblSampUP.HSXCSZ = tblSampDN.HSXCSZ, "3", "4")
            End If
        Case 16     'CJ
            If Not CheckHWS(tblSampUP.HSXCJHS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCJHS), "1", "0")
            ElseIf Not CheckHWS(tblSampDN.HSXCJHS) Then
                GetSXLSamp = "2"
            Else
                GetSXLSamp = IIf(tblSampUP.HSXCJNS = tblSampDN.HSXCJNS, "3", "4")
            End If
        Case 17     'CJ LT
            If Not CheckHWS(tblSampUP.HSXCJLTHS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCJLTHS), "1", "0")
            ElseIf Not CheckHWS(tblSampDN.HSXCJLTHS) Then
                GetSXLSamp = "2"
            Else
                GetSXLSamp = IIf(tblSampUP.HSXCJLTNS = tblSampDN.HSXCJLTNS, "3", "4")
            End If
        Case 18     'CJ2
            If Not CheckHWS(tblSampUP.HSXCJ2HS) Then
                GetSXLSamp = IIf(CheckHWS(tblSampDN.HSXCJ2HS), "1", "0")
            ElseIf Not CheckHWS(tblSampDN.HSXCJ2HS) Then
                GetSXLSamp = "2"
            Else
                GetSXLSamp = IIf(tblSampUP.HSXCJ2NS = tblSampDN.HSXCJ2NS, "3", "4")
            End If
    'Add End   2010/12/10 SMPK Miyata
        End Select
    End Select

End Function


'概要      :結晶サンプルの取得（Heavy Version）
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型           ,説明
'　　      :pHinUp　　　,I  ,tFullHinban　,上品番テーブル
'　　      :pHinDn　　　,I  ,tFullHinban　,下品番テーブル
'　　      :戻り値      ,O  ,String     　,検査指示サンプル
'説明      :RsとOiの共通検査指示について上下品番のいずれがヘヴィであるかを返す
'履歴      :2001/07/03　大塚 作成
Public Function GetSXLSampHeavy(pHinUp As tFullHinban, pHinDn As tFullHinban) As String

    Dim HINBANUP As String
    Dim HINBANDN As String
    Dim a As Boolean
    Dim b As Boolean

    '' 上品番／下品番が共に空でなければ
    HINBANUP = Trim(pHinUp.hinban)
    HINBANDN = Trim(pHinDn.hinban)
    If (HINBANUP = "" Or HINBANUP = "G" Or HINBANUP = "Z") And _
       (HINBANDN = "" Or HINBANDN = "G" Or HINBANDN = "Z") Then
        GetSXLSampHeavy = "T"
        Exit Function
    End If

    '' 共通サンプル以外は除外する
    If HINBANUP = "" Or HINBANUP = "G" Or HINBANUP = "Z" Then
        GetSXLSampHeavy = "T"
        Exit Function
    ElseIf HINBANDN = "" Or HINBANDN = "G" Or HINBANDN = "Z" Then
        GetSXLSampHeavy = "T"
        Exit Function
    ElseIf HINBANUP <> HINBANDN Then
        GetSXLSampHeavy = "T"
        Exit Function
    End If

    '' 上品番の製品仕様データを取得
    If tblSampUP.HIN.hinban <> pHinUp.hinban Then
        tblSampUP.HIN.hinban = pHinUp.hinban
        tblSampUP.HIN.mnorevno = pHinUp.mnorevno
        tblSampUP.HIN.Factory = pHinUp.Factory
        tblSampUP.HIN.OpeCond = pHinUp.OpeCond
        If scmzc_getSXL(tblSampUP) = FUNCTION_RETURN_FAILURE Then
            GetSXLSampHeavy = "T"
            Exit Function
        End If
    End If

    '' 下品番の製品仕様データを取得
    If tblSampDN.HIN.hinban <> pHinDn.hinban Then
        tblSampDN.HIN.hinban = pHinDn.hinban
        tblSampDN.HIN.mnorevno = pHinDn.mnorevno
        tblSampDN.HIN.Factory = pHinDn.Factory
        tblSampDN.HIN.OpeCond = pHinDn.OpeCond
        If scmzc_getSXL(tblSampDN) = FUNCTION_RETURN_FAILURE Then
            GetSXLSampHeavy = "T"
            Exit Function
        End If
    End If

    '' 共通サンプルに対して検査指示があるかチェック
    If CheckHWS(tblSampUP.HSXRHWYS) And CheckHWS(tblSampDN.HSXRHWYS) Then
        a = True
    Else
        a = False
    End If
    If CheckHWS(tblSampUP.HSXONHWS) And CheckHWS(tblSampDN.HSXONHWS) Then
        b = True
    Else
        b = False
    End If

    If a = True And b = True Then
        If tblSampUP.HSXRSPOT <= tblSampDN.HSXRSPOT And _
           tblSampUP.HSXONSPT <= tblSampDN.HSXONSPT Then
            GetSXLSampHeavy = "T"
        ElseIf tblSampUP.HSXRSPOT >= tblSampDN.HSXRSPOT And _
               tblSampUP.HSXONSPT >= tblSampDN.HSXONSPT Then
            GetSXLSampHeavy = "B"
        ElseIf tblSampUP.HSXRSPOT > tblSampDN.HSXRSPOT And _
               tblSampUP.HSXONSPT < tblSampDN.HSXONSPT Then
            GetSXLSampHeavy = "X"
        End If
    ElseIf a = True Then
        If tblSampUP.HSXRSPOT <= tblSampDN.HSXRSPOT Then
            GetSXLSampHeavy = "T"
        Else
            GetSXLSampHeavy = "B"
        End If
    ElseIf b = True Then
        If tblSampUP.HSXONSPT <= tblSampDN.HSXONSPT Then
            GetSXLSampHeavy = "T"
        Else
            GetSXLSampHeavy = "B"
        End If
    Else
        GetSXLSampHeavy = "T"
    End If

End Function


'概要      :結晶サンプル枚数の取得
'ﾊﾟﾗﾒｰﾀ　　:変数名          ,IO ,型            ,説明
'　　      :pSXLSample　　　,I  ,typ_SXLSample ,結晶サンプルテーブル
'　　      :戻り値          ,O  ,Integer       ,結晶サンプル枚数
'説明      :検査指示サンプル枚数を計算する
'履歴      :2001/07/03　大塚 作成
Public Function GetSXLSampNum(pSXLSample As typ_SXLSample) As Integer

    Dim sFTIR As String
    Dim bBot As Boolean
    Dim bTop As Boolean
    Dim iBot As Integer
    Dim iTop As Integer

    With pSXLSample
        bBot = False
        bTop = False
        iBot = 0
        iTop = 0
        'Rs
        Call CountSXLNum(.CRYINDRS, iBot, iTop, 1)
        'Oi
        Select Case .CRYINDOI
        Case "1"
            iTop = iTop + 1
            sFTIR = Trim$(tblSampDN.HSXONKWY)
            If sFTIR = "CA" Or sFTIR = "CD" Or sFTIR = "" Then
                bTop = True
            End If
        Case "2", "3"
            iBot = iBot + 1
            sFTIR = Trim$(tblSampUP.HSXONKWY)
            If sFTIR = "CA" Or sFTIR = "CD" Or sFTIR = "" Then
                bBot = True
            End If
        Case "4"
            iBot = iBot + 1
            sFTIR = Trim$(tblSampUP.HSXONKWY)
            If sFTIR = "CA" Or sFTIR = "CD" Or sFTIR = "" Then
                bBot = True
            End If
            iTop = iTop + 1
            sFTIR = Trim$(tblSampDN.HSXONKWY)
            If sFTIR = "CA" Or sFTIR = "CD" Or sFTIR = "" Then
                bTop = True
            End If
        End Select
        'B1
        Call CountSXLNum(.CRYINDB1, iBot, iTop, 1)
        'B2
        Call CountSXLNum(.CRYINDB2, iBot, iTop, 1)
        'B3
        Call CountSXLNum(.CRYINDB3, iBot, iTop, 1)
        'L1
        Call CountSXLNum(.CRYINDL1, iBot, iTop, 1)
        'L2
        Call CountSXLNum(.CRYINDL2, iBot, iTop, 1)
        'L3
        Call CountSXLNum(.CRYINDL3, iBot, iTop, 1)
        'L4
        Call CountSXLNum(.CRYINDL4, iBot, iTop, 1)
        'Cs
        Select Case .CRYINDCS
        Case "1"
            If bTop = False Then
                iTop = iTop + 1
            End If
        Case "2", "3"
            If bBot = False Then
                iBot = iBot + 1
            End If
        Case "4"
            If bBot = False Then
                iBot = iBot + 1
            End If
            If bTop = False Then
                iTop = iTop + 1
            End If
        End Select
        'GD
        Call CountSXLNum(.CRYINDGD, iBot, iTop, 1)
        'T
        Call CountSXLNum(.CRYINDT, iBot, iTop, 1)
        'EPD
        Call CountSXLNum(.CRYINDEP, iBot, iTop, 4)
        
        'X      'X線測定追加 2009/07/27 SETsw kubota
        Call CountSXLNum(.CRYINDX, iBot, iTop, 4)

        'Add Start 2010/12/13 SMPK Miyata
        'C
        Call CountSXLNum(.CRYINDC, iBot, iTop, 1)
        'CJ
        Call CountSXLNum(.CRYINDCJ, iBot, iTop, 1)
        'CJLT   CJがｻﾝﾌﾟﾙ有りの時はｶｳﾝﾄしない
        If .CRYINDCJ <> "1" Then
            Call CountSXLNum(.CRYINDCJLT, iBot, iTop, 1)
        End If
        'CJ2
        Call CountSXLNum(.CRYINDCJ2, iBot, iTop, 1)
        'Add End   2010/12/13 SMPK Miyata

        'Sum
        GetSXLSampNum = Int((iBot + iTop) / 4)
        If (iBot + iTop) Mod 4 > 0 Then
            GetSXLSampNum = GetSXLSampNum + 1
        End If
    End With

End Function

'概要      :結晶サンプル枚数の個別カウント
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型             ,説明
'　　      :sSamp   　　　,I  ,String       　,結晶検査項目
'　　      :iBotNum 　　　,O  ,Integer      　,ブロックボトム側枚数
'　　      :iTopNum 　　　,O  ,Integer      　,ブロックトップ側枚数
'　　      :iCountUp　　　,I  ,Integer      　,カウントアップ値
'説明      :結晶サンプル枚数をカウントアップする
'履歴      :2001/07/03　大塚 作成
Public Sub CountSXLNum(ByVal sSamp As String, iBotNum As Integer, iTopNum As Integer, iCountUp As Integer)

    Select Case sSamp
    Case "1"
        iTopNum = iTopNum + iCountUp
    Case "2", "3"
        iBotNum = iBotNum + iCountUp
    Case "4"
        iBotNum = iBotNum + iCountUp
        iTopNum = iTopNum + iCountUp
    End Select

End Sub

'概要      :処理方法のチェック
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型       ,説明
'　　      :sHWS  　　　,I  ,String 　,処理方法
'      　　:戻り値      ,O  ,Boolean　,検査の有無
'説明      :処理方法をチェックして検査の有無を返す
'履歴      :2001/07/03　大塚 作成
Private Function CheckHWS(ByVal sHWS As String) As Boolean

    If sHWS = "H" Or sHWS = "S" Then
        CheckHWS = True
    Else
        CheckHWS = False
    End If

End Function

'概要      :Ｇ、Ｚ品番用製品仕様SXLデータの取得
'ﾊﾟﾗﾒｰﾀ　　:変数名          ,IO ,型             ,説明
'      　　:pSpSXLSamp　　　,O  ,typ_SpSXLSamp　,結晶サンプル仕様
'      　　:pHinTg    　　　,I  ,tFullHinban  　,ねらい品番テーブル
'説明      :Ｇ、Ｚ品番のデフォルト製品仕様を返す
'履歴      :2001/09/11　大塚 作成
Public Sub GetSpecGZ(pSpSXLSamp As typ_SpSXLSamp, pHinTg As tFullHinban)

    With pSpSXLSamp
        If Trim$(.HIN.hinban) = "G" Then
            If tblSampTG.HIN.hinban <> pHinTg.hinban Then
                tblSampTG.HIN.hinban = pHinTg.hinban
                tblSampTG.HIN.mnorevno = pHinTg.mnorevno
                tblSampTG.HIN.Factory = pHinTg.Factory
                tblSampTG.HIN.OpeCond = pHinTg.OpeCond
                If scmzc_getSXL(tblSampTG) = FUNCTION_RETURN_FAILURE Then
                    tblSampTG.HSXONKWY = "CD"
                    tblSampTG.HSXONSPH = "S"
                    tblSampTG.HSXONSPI = "T"
                End If
            End If
            .HSXONKWY = tblSampTG.HSXONKWY
            .HSXONSPH = tblSampTG.HSXONSPH
            .HSXONSPI = tblSampTG.HSXONSPI
            .HSXONHWS = "S"
            .HSXONSPT = "3"
            .HSXLTHWS = "S"
        Else
            .HSXONHWS = ""
            .HSXONKWY = ""
            .HSXONSPH = ""
            .HSXONSPT = ""
            .HSXONSPI = ""
            .HSXLTHWS = ""
        End If
        .HSXRHWYS = "S"
        .HSXRSPOT = "3"
        .HSXCNHWS = "S"
        .CS_FROMTO = False      '09/01/12 ooba
        .HSXBM1HS = ""
        .HSXBM1SH = ""
        .HSXBM1ST = ""
        .HSXBM1SR = ""
        .HSXBM1NS = ""
        .HSXBM1SZ = ""
        .HSXBM1ET = 0
        .HSXBM2HS = ""
        .HSXBM2SH = ""
        .HSXBM2ST = ""
        .HSXBM2SR = ""
        .HSXBM2NS = ""
        .HSXBM2SZ = ""
        .HSXBM2ET = 0
        .HSXBM3HS = ""
        .HSXBM3SH = ""
        .HSXBM3ST = ""
        .HSXBM3SR = ""
        .HSXBM3NS = ""
        .HSXBM3SZ = ""
        .HSXBM3ET = 0
        .HSXOF1HS = ""
        .HSXOF1SH = ""
        .HSXOF1ST = ""
        .HSXOF1SR = ""
        .HSXOF1NS = ""
        .HSXOF1SZ = ""
        .HSXOF1ET = 0
        .HSXOF2HS = ""
        .HSXOF2SH = ""
        .HSXOF2ST = ""
        .HSXOF2SR = ""
        .HSXOF2NS = ""
        .HSXOF2SZ = ""
        .HSXOF2ET = 0
        .HSXOF3HS = ""
        .HSXOF3SH = ""
        .HSXOF3ST = ""
        .HSXOF3SR = ""
        .HSXOF3NS = ""
        .HSXOF3SZ = ""
        .HSXOF3ET = 0
        .HSXOF4HS = ""
        .HSXOF4SH = ""
        .HSXOF4ST = ""
        .HSXOF4SR = ""
        .HSXOF4NS = ""
        .HSXOF4SZ = ""
        .HSXOF4ET = 0
        .HSXDENHS = ""
        .HSXDVDHS = ""
        .HSXLDLHS = ""
        'Add Start 2010/12/10 SMPK Miyata
        .HSXCHS = ""        ' 処理方法(C)
        .HSXCSZ = ""        ' 測定条件(C)
        .HSXCJHS = ""       ' 処理方法(CJ)
        .HSXCJNS = ""       ' 熱処理法(CJ)
        .HSXCJLTHS = ""     ' 処理方法(CJ LT)
        .HSXCJLTNS = ""     ' 熱処理法(CJ LT)
        .HSXCJ2HS = ""      ' 処理方法(CJ2)
        .HSXCJ2NS = ""      ' 熱処理法(CJ2)
        'Add End   2010/12/10 SMPK Miyata
    End With

End Sub
'概要      :Ｇ、Ｚ品番用製品仕様SXLデータの取得
'ﾊﾟﾗﾒｰﾀ　　:変数名          ,IO ,型             ,説明
'      　　:pSpSXLSamp　　　,O  ,typ_SpSXLSamp　,結晶サンプル仕様
'      　　:pHinTg    　　　,I  ,tFullHinban  　,ねらい品番テーブル
'説明      :Ｇ、Ｚ品番のデフォルト製品仕様を返す
'履歴      :2001/09/11　大塚 作成
Public Sub GetSpecZ2(pSpSXLSamp As typ_SpSXLSamp, pHinTg As tFullHinban)

    With pSpSXLSamp
        .HSXONHWS = ""
        .HSXONKWY = ""
        .HSXONSPH = ""
        .HSXONSPT = ""
        .HSXONSPI = ""
        .HSXLTHWS = ""
        
        .HSXRHWYS = ""
        .HSXRSPOT = ""
        .HSXCNHWS = ""
        .CS_FROMTO = False      '09/01/12 ooba
        .HSXRHWYS = ""
        .HSXRSPOT = ""
        .HSXCNHWS = ""
        .HSXBM1HS = ""
        .HSXBM1SH = ""
        .HSXBM1ST = ""
        .HSXBM1SR = ""
        .HSXBM1NS = ""
        .HSXBM1SZ = ""
        .HSXBM1ET = 0
        .HSXBM2HS = ""
        .HSXBM2SH = ""
        .HSXBM2ST = ""
        .HSXBM2SR = ""
        .HSXBM2NS = ""
        .HSXBM2SZ = ""
        .HSXBM2ET = 0
        .HSXBM3HS = ""
        .HSXBM3SH = ""
        .HSXBM3ST = ""
        .HSXBM3SR = ""
        .HSXBM3NS = ""
        .HSXBM3SZ = ""
        .HSXBM3ET = 0
        .HSXOF1HS = ""
        .HSXOF1SH = ""
        .HSXOF1ST = ""
        .HSXOF1SR = ""
        .HSXOF1NS = ""
        .HSXOF1SZ = ""
        .HSXOF1ET = 0
        .HSXOF2HS = ""
        .HSXOF2SH = ""
        .HSXOF2ST = ""
        .HSXOF2SR = ""
        .HSXOF2NS = ""
        .HSXOF2SZ = ""
        .HSXOF2ET = 0
        .HSXOF3HS = ""
        .HSXOF3SH = ""
        .HSXOF3ST = ""
        .HSXOF3SR = ""
        .HSXOF3NS = ""
        .HSXOF3SZ = ""
        .HSXOF3ET = 0
        .HSXOF4HS = ""
        .HSXOF4SH = ""
        .HSXOF4ST = ""
        .HSXOF4SR = ""
        .HSXOF4NS = ""
        .HSXOF4SZ = ""
        .HSXOF4ET = 0
        .HSXDENHS = ""
        .HSXDVDHS = ""
        .HSXLDLHS = ""
        'Add Start 2010/12/10 SMPK Miyata
        .HSXCHS = ""        ' 処理方法(C)
        .HSXCSZ = ""        ' 測定条件(C)
        .HSXCJHS = ""       ' 処理方法(CJ)
        .HSXCJNS = ""       ' 熱処理法(CJ)
        .HSXCJLTHS = ""     ' 処理方法(CJ LT)
        .HSXCJLTNS = ""     ' 熱処理法(CJ LT)
        .HSXCJ2HS = ""      ' 処理方法(CJ2)
        .HSXCJ2NS = ""      ' 熱処理法(CJ2)
        'Add End   2010/12/10 SMPK Miyata
    End With

End Sub


'概要      :結晶サンプル枚数の取得
'ﾊﾟﾗﾒｰﾀ　　:変数名          ,IO ,型            ,説明
'　　      :pSXLSample　　　,I  ,typ_SXLSample ,結晶サンプルテーブル
'　　      :sBT       　　　,I  ,String      　,T:TOP,B:BOT
'          :iBotCnt         ,I  ,Integer       ,n:ボトム側結晶サンプル数
'          :iKbn            ,I  ,Integer       ,0:結晶サンプル数、1:結晶サンプル枚数
'　　      :戻り値          ,O  ,Integer       ,結晶サンプル数、結晶サンプル枚数
'説明      :検査指示サンプル枚数を計算する
'履歴      :2001/07/03　大塚 作成
Public Function GetSXLSampNum_2(pSXLSample As typ_SXLSample, sBT As String, iBotCnt As Integer, ikbn As Integer) As Integer

    Dim sFTIR As String
    Dim bBTFlg As Boolean
    Dim iBTCnt As Integer

    With pSXLSample
        bBTFlg = False
        iBTCnt = 0
        
        'Rs
        Call CountSXLNum_2(.CRYINDRS, iBTCnt, 1)
        'Oi
        Call CountSXLNum_2(.CRYINDOI, iBTCnt, 1)
        Select Case .CRYINDOI
        Case "1"
            If sBT = "B" Then
                sFTIR = Trim$(tblSampUP.HSXONKWY)
                If sFTIR = "CA" Or sFTIR = "CD" Or sFTIR = "" Then
                    bBTFlg = True
                End If
            Else
                sFTIR = Trim$(tblSampDN.HSXONKWY)
                If sFTIR = "CA" Or sFTIR = "CD" Or sFTIR = "" Then
                    bBTFlg = True
                End If
            End If
        End Select
        'B1
        Call CountSXLNum_2(.CRYINDB1, iBTCnt, 1)
        'B2
        Call CountSXLNum_2(.CRYINDB2, iBTCnt, 1)
        'B3
        Call CountSXLNum_2(.CRYINDB3, iBTCnt, 1)
        'L1
        Call CountSXLNum_2(.CRYINDL1, iBTCnt, 1)
        'L2
        Call CountSXLNum_2(.CRYINDL2, iBTCnt, 1)
        'L3
        Call CountSXLNum_2(.CRYINDL3, iBTCnt, 1)
        'L4
        Call CountSXLNum_2(.CRYINDL4, iBTCnt, 1)
        'Cs
        Select Case .CRYINDCS
        Case "1"
            If bBTFlg = False Then
                iBTCnt = iBTCnt + 1
            End If
        End Select
        'GD
        Call CountSXLNum_2(.CRYINDGD, iBTCnt, 1)
        'T
        Call CountSXLNum_2(.CRYINDT, iBTCnt, 1)
        'EPD
        Call CountSXLNum_2(.CRYINDEP, iBTCnt, 4)
        
        'X      'X線測定追加 2009/07/27 SETsw kubota
        Call CountSXLNum_2(.CRYINDX, iBTCnt, 4)

        'Add Start 2010/12/13 SMPK Miyata
        'C
        Call CountSXLNum_2(.CRYINDC, iBTCnt, 1)
        'CJ
        Call CountSXLNum_2(.CRYINDCJ, iBTCnt, 1)
        'CJLT   CJがｻﾝﾌﾟﾙ有りの時はｶｳﾝﾄしない
        If .CRYINDCJ <> "1" Then
            Call CountSXLNum_2(.CRYINDCJLT, iBTCnt, 1)
        End If
        'CJ2
        Call CountSXLNum_2(.CRYINDCJ2, iBTCnt, 1)
        'Add End   2010/12/13 SMPK Miyata

        'Sum
        If ikbn = 0 Then
            GetSXLSampNum_2 = iBTCnt
        Else
            GetSXLSampNum_2 = Int((iBTCnt + iBotCnt) / 4)
            If (iBTCnt + iBotCnt) Mod 4 > 0 Then
                GetSXLSampNum_2 = GetSXLSampNum_2 + 1
            End If
        End If
    End With

End Function

'概要      :結晶サンプル枚数の個別カウント
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型             ,説明
'　　      :sSamp   　　　,I  ,String       　,結晶検査項目
'　　      :iBTNum  　　　,O  ,Integer      　,ブロック数
'　　      :iCountUp　　　,I  ,Integer      　,カウントアップ値
'説明      :結晶サンプル枚数をカウントアップする
'履歴      :2001/07/03　大塚 作成
Public Sub CountSXLNum_2(ByVal sSamp As String, iBTNum As Integer, iCountUp As Integer)

    Select Case sSamp
    Case "1"
        iBTNum = iBTNum + iCountUp
    End Select

End Sub


Attribute VB_Name = "s_HaraiKisei"
'*******************************************************************************
'*    モジュール名  : s_HaraiKisei
'*
'*    処理概要      :払出し規制追加
'*    ﾊﾟﾗﾒｰﾀ        :typ_CType.typ_Param
'*    説明          :結晶総合判定より流用
'*    履歴          :2010/02/15 Kameda
'*******************************************************************************
Option Explicit

'規制値計算用
Private Type typ_Kisei
    TOPREG         As Integer      '計算用TOP規制
    TAILREG        As Double       '計算用TAIL規制
    BTMSPRT        As Integer      '計算用析出規制
    BTMSPRTD       As Integer      '表示用析出規制
    KISEIP         As Integer
    KISEID         As Integer
    hinban         As tFullHinban
    CRYNUM         As String
    TANJU          As Double
    top            As Integer
    TAIL           As Integer
    LENGTKDO       As Integer
    WGHTTOP        As Long
    WGTOPCUT       As Long
    CHARGE         As Long
    TBFLGT         As Integer      'エラー位置 TOP
    TBFLGB         As Integer      'エラー位置 BOT
    PULENTKC1 As Long              ' 引上直胴長さ  2010/09/06 add Kameda
    PUWGHTTKC1 As Long             ' 引上直胴重量  2010/09/06 add Kameda
End Type

Public Sub PutAllData_Haraidashi()

    Dim wHinban         As tFullHinban
    Dim wJudg           As Boolean
    Dim KISEI As typ_Kisei
    
    With f_cmbc039_2
        
        If .txtSXLTop.text = "" Or .txtSXLTail.text = "" Then                   ' Top , Tail 値が不明確の場合処理しない
           Exit Sub
        End If
        'TOP,TAIL位置の取得
        KISEI.top = .txtSXLTop.text
        KISEI.TAIL = .txtSXLTail.text
            
        wJudg = True
        
        KISEI.KISEIP = 0
            
        KISEI.hinban.hinban = typ_CType.typ_Param.hinban
        KISEI.hinban.mnorevno = typ_CType.typ_Param.REVNUM
        KISEI.hinban.factory = typ_CType.typ_Param.factory
        KISEI.hinban.opecond = typ_CType.typ_Param.opecond
            
        KISEI.CRYNUM = typ_CType.typ_Param.CRYNUM
        
        '規制ﾁｪｯｸ
    
        If HaraidashiKisei(KISEI) Then
            .txtKisei.text = "OK"
        Else
            .txtKisei.text = "NG"
            .txtKisei.backColor = COLOR_RED
            wJudg = False
        End If
    
        '規制値仕様表示
        .spdHinbanTop.col = 6
        .spdHinbanTop.Value = KISEI.TOPREG
        .spdHinbanTop.col = 7
        If KISEI.KISEID <> 0 Then
            .spdHinbanTop.Value = KISEI.KISEID
        End If
        .spdHinbanTop.col = 8
        If KISEI.BTMSPRTD <> 0 Then
            .spdHinbanTop.Value = KISEI.BTMSPRTD
        End If
    End With
    ''払出規制も総合判定に付加する
    If wJudg = False Then
        TotalJudg039 = False
    End If
        
End Sub
'払出規制値を算出するための値を取得
'2010/09/06      払出規制計算項目変更に付き取得項目追加 LENTKC1→PULENTKC1,WGHTTKC1→PUWGHTTKC1 Kameda
Sub PutAllData_Haraidashi_Kisei(wBlockID As String, wTANJU As Double, wLENGTKDO As Integer, wWGHTTOP As Long, wWGTOPCUT As Long, wCHARGE As Long)
    Dim sql             As String
    Dim rs              As OraDynaset
    Dim wWGHTTKDO       As Long
    
    
' TBCMH004⇒XSDC1へ変更 2005/04/15 ffc)tanabe ==============> START
    sql = "select "
    sql = sql & "LENTKC1, "         ' 長さ（直胴）
    sql = sql & "WGHTTKC1, "        ' 重量（直胴）
    sql = sql & "PUTCUTWC1, "       ' TOP取り重量
    sql = sql & "WGHTTOC1, "        ' 重量（TOP）
    sql = sql & "SUICHARGE "        ' 推定チャージ
    sql = sql & ",PULENTKC1"        ' 引上げ直胴長さ　2010/09/06 add Kameda
    sql = sql & ",PUWGHTTKC1"       ' 引上げ直胴重量　2010/09/06 add Kameda
    sql = sql & " from XSDC1"
    sql = sql & " where XTALC1 = '" & wBlockID & "'"
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    
    'wLENGTKDO = rs("LENTKC1")           ' 長さ（直胴）  2010/09/06 Kameda
    'wWGHTTKDO = rs("WGHTTKC1")          ' 重量（直胴）  2010/09/06 Kameda
    wLENGTKDO = rs("PULENTKC1")          ' 引上げ直胴長さ　2010/09/06 add Kameda
    wWGHTTKDO = rs("PUWGHTTKC1")         ' 引上げ直胴重量　2010/09/06 add Kameda
    wWGTOPCUT = rs("PUTCUTWC1")          ' TOP取り重量
    wWGHTTOP = rs("WGHTTOC1")            ' 重量（TOP）
    wCHARGE = rs("SUICHARGE")            ' 推定チャージ量
    rs.Close
' TBCMH004⇒XSDC1へ変更 2005/04/15 ffc)tanabe ==============> END

    wTANJU = Round((wWGHTTKDO / wLENGTKDO), 2)
End Sub

' 規制値取得
Public Sub PutSeihinAll_Reg(wHinban As tFullHinban, wTOPREG As Integer, wTAILREG As Double, wBTMSPRT As Integer)
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim recCnt      As Integer

    sql = "select nvl(TOPREG,0) as TOPREG, nvl(TAILREG,0) as TAILREG,nvl(BTMSPRT,0) as BTMSPRT "
    sql = sql & " from  TBCME036"
    sql = sql & " where HINBAN  = '" & wHinban.hinban & "'"
    sql = sql & "   and MNOREVNO=  " & wHinban.mnorevno
    sql = sql & "   and FACTORY = '" & wHinban.factory & "'"
    sql = sql & "   and OPECOND = '" & wHinban.opecond & "'"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt > 0 Then
          wTOPREG = rs("TOPREG")
          wTAILREG = rs("TAILREG")
          wBTMSPRT = rs("BTMSPRT")
    Else
          ' --TEST-- 仮セット：本来０件はあり得ない
          wTOPREG = 0
          wTAILREG = 0
          wBTMSPRT = 0
    End If
    rs.Close
End Sub
'再抜試の払出し規制のチェックを全て行う
Public Function F_HaraiKisei() As Boolean
  Dim i_Lp1 As Integer
  Dim i_Lp2 As Integer
  Dim i_Sec As Integer
  Dim bEndFlg  As Boolean
  Dim KISEI As typ_Kisei
  
  bEndFlg = True

    With f_cmbc039_3.sprExamine
        For i_Lp1 = 1 To .MaxRows - 1 Step 2
            .row = i_Lp1: .col = 2
            KISEI.hinban.hinban = Trim(.Value)
            .row = i_Lp1: .col = 3
            KISEI.hinban.mnorevno = left(Trim(.Value), 2)
            KISEI.hinban.factory = Mid(Trim(.Value), 3, 1)
            KISEI.hinban.opecond = Mid(Trim(.Value), 4)
            
            '製品品番のみチェックを行う
            If Trim(KISEI.hinban.hinban) <> "Z" And Trim(KISEI.hinban.hinban) <> "G" And Trim(KISEI.hinban.hinban) <> "" Then
                'TOP,TAIL位置の取得
                .row = i_Lp1: .col = 5
                KISEI.top = Trim(.Value)
                .row = i_Lp1 + 1: .col = 5
                KISEI.TAIL = Trim(.Value)
                
                KISEI.CRYNUM = f_cmbc039_3.txtCryNum.text
                
            '規制ﾁｪｯｸ
                If HaraidashiKisei(KISEI) = False Then
                    If KISEI.TBFLGT = 1 Then
                        .row = i_Lp1
                        .col = 5
                        .backColor = COLOR_RED
                        DoEvents
                        bEndFlg = False
                    End If
                    If KISEI.TBFLGB = 1 Then
                        .row = i_Lp1 + 1
                        .col = 5
                        .backColor = COLOR_RED
                        DoEvents
                        bEndFlg = False
                    End If
                End If
            End If
        Next i_Lp1
    End With

    F_HaraiKisei = bEndFlg

End Function
Private Function HaraidashiKisei(KISEI As typ_Kisei) As Boolean
    
HaraidashiKisei = True
With KISEI

    .TBFLGT = 0
    .TBFLGB = 0
    
    '規制値仕様取得
    Call PutSeihinAll_Reg(.hinban, .TOPREG, .TAILREG, .BTMSPRT)
    
    If .TOPREG = 0 And .TAILREG = 0 And .BTMSPRT = 0 Then
        '規制がない場合は処理しない
    Else

        '払出規制値を算出するための値を取得
        Call PutAllData_Haraidashi_Kisei(.CRYNUM, .TANJU, .LENGTKDO, .WGHTTOP, .WGTOPCUT, .CHARGE)
        
        If .TAILREG <> 0 Then
           .TAILREG = .TAILREG / 100            ' Tail規制(%)
           .KISEIP = Int(((.CHARGE * .TAILREG - .WGTOPCUT - .WGHTTOP) / .TANJU) + 0.9)  '規制値を算出
           .KISEID = .KISEIP      '表示用データ
        End If
        
        If .BTMSPRT <> 0 Then                                  '析出0は規制をかけない add 2010/03/31 Kameda
            .BTMSPRT = .LENGTKDO - .BTMSPRT                    '直胴長から規制のエリアを算出
            If .KISEIP = 0 Or .KISEIP > .BTMSPRT Then          '規制エリアを比較
               .KISEIP = .BTMSPRT
            End If
            .BTMSPRTD = .BTMSPRT                                '表示用データ Add 2010/3/15 Y.Hitomi
        End If
        
       
        'If .TAILREG <> 0 Then      2010/03/04 Kameda
        '   If .TOPREG <= .top And .KISEIP >= .TAIL Then
        '   Else
        '     HaraidashiKisei = False
        '   End If
        'End If
        
        If .TOPREG > .top Then
            HaraidashiKisei = False
            .TBFLGT = 1
        End If
        If .KISEIP <> 0 Then
           If .KISEIP < .TAIL Then
                HaraidashiKisei = False
                .TBFLGB = 1
           End If
        End If
    End If
End With

End Function


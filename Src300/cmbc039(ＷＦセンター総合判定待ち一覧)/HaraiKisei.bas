Attribute VB_Name = "s_HaraiKisei"
'*******************************************************************************
'*    ���W���[����  : s_HaraiKisei
'*
'*    �����T�v      :���o���K���ǉ�
'*    ���Ұ�        :typ_CType.typ_Param
'*    ����          :�������������藬�p
'*    ����          :2010/02/15 Kameda
'*******************************************************************************
Option Explicit

'�K���l�v�Z�p
Private Type typ_Kisei
    TOPREG         As Integer      '�v�Z�pTOP�K��
    TAILREG        As Double       '�v�Z�pTAIL�K��
    BTMSPRT        As Integer      '�v�Z�p�͏o�K��
    BTMSPRTD       As Integer      '�\���p�͏o�K��
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
    TBFLGT         As Integer      '�G���[�ʒu TOP
    TBFLGB         As Integer      '�G���[�ʒu BOT
    PULENTKC1 As Long              ' ���㒼������  2010/09/06 add Kameda
    PUWGHTTKC1 As Long             ' ���㒼���d��  2010/09/06 add Kameda
End Type

Public Sub PutAllData_Haraidashi()

    Dim wHinban         As tFullHinban
    Dim wJudg           As Boolean
    Dim KISEI As typ_Kisei
    
    With f_cmbc039_2
        
        If .txtSXLTop.text = "" Or .txtSXLTail.text = "" Then                   ' Top , Tail �l���s���m�̏ꍇ�������Ȃ�
           Exit Sub
        End If
        'TOP,TAIL�ʒu�̎擾
        KISEI.top = .txtSXLTop.text
        KISEI.TAIL = .txtSXLTail.text
            
        wJudg = True
        
        KISEI.KISEIP = 0
            
        KISEI.hinban.hinban = typ_CType.typ_Param.hinban
        KISEI.hinban.mnorevno = typ_CType.typ_Param.REVNUM
        KISEI.hinban.factory = typ_CType.typ_Param.factory
        KISEI.hinban.opecond = typ_CType.typ_Param.opecond
            
        KISEI.CRYNUM = typ_CType.typ_Param.CRYNUM
        
        '�K������
    
        If HaraidashiKisei(KISEI) Then
            .txtKisei.text = "OK"
        Else
            .txtKisei.text = "NG"
            .txtKisei.backColor = COLOR_RED
            wJudg = False
        End If
    
        '�K���l�d�l�\��
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
    ''���o�K������������ɕt������
    If wJudg = False Then
        TotalJudg039 = False
    End If
        
End Sub
'���o�K���l���Z�o���邽�߂̒l���擾
'2010/09/06      ���o�K���v�Z���ڕύX�ɕt���擾���ڒǉ� LENTKC1��PULENTKC1,WGHTTKC1��PUWGHTTKC1 Kameda
Sub PutAllData_Haraidashi_Kisei(wBlockID As String, wTANJU As Double, wLENGTKDO As Integer, wWGHTTOP As Long, wWGTOPCUT As Long, wCHARGE As Long)
    Dim sql             As String
    Dim rs              As OraDynaset
    Dim wWGHTTKDO       As Long
    
    
' TBCMH004��XSDC1�֕ύX 2005/04/15 ffc)tanabe ==============> START
    sql = "select "
    sql = sql & "LENTKC1, "         ' �����i�����j
    sql = sql & "WGHTTKC1, "        ' �d�ʁi�����j
    sql = sql & "PUTCUTWC1, "       ' TOP���d��
    sql = sql & "WGHTTOC1, "        ' �d�ʁiTOP�j
    sql = sql & "SUICHARGE "        ' ����`���[�W
    sql = sql & ",PULENTKC1"        ' ���グ���������@2010/09/06 add Kameda
    sql = sql & ",PUWGHTTKC1"       ' ���グ�����d�ʁ@2010/09/06 add Kameda
    sql = sql & " from XSDC1"
    sql = sql & " where XTALC1 = '" & wBlockID & "'"
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    
    'wLENGTKDO = rs("LENTKC1")           ' �����i�����j  2010/09/06 Kameda
    'wWGHTTKDO = rs("WGHTTKC1")          ' �d�ʁi�����j  2010/09/06 Kameda
    wLENGTKDO = rs("PULENTKC1")          ' ���グ���������@2010/09/06 add Kameda
    wWGHTTKDO = rs("PUWGHTTKC1")         ' ���グ�����d�ʁ@2010/09/06 add Kameda
    wWGTOPCUT = rs("PUTCUTWC1")          ' TOP���d��
    wWGHTTOP = rs("WGHTTOC1")            ' �d�ʁiTOP�j
    wCHARGE = rs("SUICHARGE")            ' ����`���[�W��
    rs.Close
' TBCMH004��XSDC1�֕ύX 2005/04/15 ffc)tanabe ==============> END

    wTANJU = Round((wWGHTTKDO / wLENGTKDO), 2)
End Sub

' �K���l�擾
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
          ' --TEST-- ���Z�b�g�F�{���O���͂��蓾�Ȃ�
          wTOPREG = 0
          wTAILREG = 0
          wBTMSPRT = 0
    End If
    rs.Close
End Sub
'�Ĕ����̕��o���K���̃`�F�b�N��S�čs��
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
            
            '���i�i�Ԃ̂݃`�F�b�N���s��
            If Trim(KISEI.hinban.hinban) <> "Z" And Trim(KISEI.hinban.hinban) <> "G" And Trim(KISEI.hinban.hinban) <> "" Then
                'TOP,TAIL�ʒu�̎擾
                .row = i_Lp1: .col = 5
                KISEI.top = Trim(.Value)
                .row = i_Lp1 + 1: .col = 5
                KISEI.TAIL = Trim(.Value)
                
                KISEI.CRYNUM = f_cmbc039_3.txtCryNum.text
                
            '�K������
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
    
    '�K���l�d�l�擾
    Call PutSeihinAll_Reg(.hinban, .TOPREG, .TAILREG, .BTMSPRT)
    
    If .TOPREG = 0 And .TAILREG = 0 And .BTMSPRT = 0 Then
        '�K�����Ȃ��ꍇ�͏������Ȃ�
    Else

        '���o�K���l���Z�o���邽�߂̒l���擾
        Call PutAllData_Haraidashi_Kisei(.CRYNUM, .TANJU, .LENGTKDO, .WGHTTOP, .WGTOPCUT, .CHARGE)
        
        If .TAILREG <> 0 Then
           .TAILREG = .TAILREG / 100            ' Tail�K��(%)
           .KISEIP = Int(((.CHARGE * .TAILREG - .WGTOPCUT - .WGHTTOP) / .TANJU) + 0.9)  '�K���l���Z�o
           .KISEID = .KISEIP      '�\���p�f�[�^
        End If
        
        If .BTMSPRT <> 0 Then                                  '�͏o0�͋K���������Ȃ� add 2010/03/31 Kameda
            .BTMSPRT = .LENGTKDO - .BTMSPRT                    '����������K���̃G���A���Z�o
            If .KISEIP = 0 Or .KISEIP > .BTMSPRT Then          '�K���G���A���r
               .KISEIP = .BTMSPRT
            End If
            .BTMSPRTD = .BTMSPRT                                '�\���p�f�[�^ Add 2010/3/15 Y.Hitomi
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


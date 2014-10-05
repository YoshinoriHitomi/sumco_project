Attribute VB_Name = "s_cmzcXDopeCal"
Option Explicit
'�����f�[�^�\����
Public Type typ_Select_BlockData
    StrCryNum As String                               '�u���b�N�ԍ�
    StrSijiNum As String                              '���グ�ԍ�
    StrXtalNum As String                              '�����ԍ�
    dblCHARGE As Double                               '�`���[�W��
    strNERAIN As String                               '�_��[N]
    StrHinban As String                               '�i��
    intMNOREVNO As Integer                            '���i�ԍ������ԍ�
    StrFactory As String                              '�H��
    StrOpeCond As String                              '���Ə���
    intTOPOS As Integer                               'TOP�ʒu
    intBOPOS As Integer                               'BOT�ʒu
    strCRYDOPCL As String                             '�h�[�p���g���
    dblCRYDOPVL As Double                             '�h�[�p���g��
    dblAIMPOS As Double                               '�˂炢�ʒu
    dblWGHTTO As Double                               '�g�b�v�d��
    intDOPESISU As Integer                            '�w��
    dblDopeRyo As Double                              '�h�[�p���g�ʁi�w���Ȃ��j
    dblDIA As Double                                  '���グ���a
    intDISPSISU As Integer                            '�\���w��
    dblTOPCUT As Double                               '�g�b�v�J�b�g�d��    2009/10/05 Kameda
End Type
'�����h�[�v�v�Z�l�\����
Public Type typ_XDOPE_KeisanData
    dblTopLength      As Double   '�g�b�v����
    dblInitLiquid     As Double   '�����Z�t�̐�
    dblPulRate        As Double   '�_���ʒu���㗦
    dblNeraiInit      As Double   '�_������[N]
    dblNeraiDope      As Double   '�_���h�[�v(Si3N4)��
    dblSi3N4KibanWt   As Double   '��Տd��
    dblSi3N4Weight_10 As Double   'Si3N4�d��(1�ʂ�)
    dblSi3N4Weight_05 As Double   'Si3N4�d��(0.5�ʂ�)
    dblSi3N4Weight_01 As Double   'Si3N4�d��(0.1�ʂ�)
    dblSi3N4Weight_0015 As Double   'Si3N4�d��(0.015�ʂ�)   2011/08/25 Kameda
    dblMaisu_10       As Double   '1�ʂ��h�[�vWF����
    intXDopeRyo_10    As Integer  '1�ʂ��h�[�v��
    intXDopeRyo_05    As Integer  '0.5�ʂ��h�[�v��
    intXDopeRyo_01    As Integer  '0.1�ʂ��h�[�v��
    intXDopeRyo_0015  As Integer  '0.015�ʂ��h�[�v��    2011/08/25 Kameda
    intXDopeRyoJ_10   As Integer  '1�ʂ��h�[�v�ʎ���
    intXDopeRyoJ_05   As Integer  '0.5�ʂ��h�[�v�ʎ���
    intXDopeRyoJ_01   As Integer  '0.1�ʂ��h�[�v�ʎ���
    intXDopeRyoJ_0015 As Integer  '0.015�ʂ��h�[�v�ʎ���    2011/08/25 Kameda
    dblDopeKei        As Double   '���v�h�[�vWF�d��
    dblDopeRyo        As Double   '�h�[�v(Si3N4)��
    dblSyokiN         As Double   '����[N]
End Type
'���f�Z�x�\����
Public Type typ_NNOUDO_Data
    intXtalPos As Integer         '�����ʒu
    dblNnoudo As Double           '���f�Z�x
    dblPulWt As Double            '���グ�d��
    dblPuRitu As Double           '���グ��
End Type
'���f�K�i
Public Type typ_spec_N
    HSXCDOPMN As Double
    HSXCDOPMX As Double
    HSXCDPNI As String
    hinban As tFullHinban
    HSXCDOP As String        '2009/09/28 Kameda
End Type

'�����h�[�v�v�Z�p�萔�ݒ�
Public Const CDOPCALC_DIA As Double = 315#             '�����グ�a
Public Const CDOPCALC_CHARGE As Double = 360#          '�`���[�W��
Public Const CDOPCALC_TOPWEIGHT As Double = 7#         '�g�b�v�d��
Public Const CDOPCALC_NERAIPOS As Double = 0#          '�_���ʒu
Public Const CDOPCALC_NERAIN  As String = "2.00E+13"   '�_��N
Public Const CDOPCALC_TOPCUT As Double = 0#            '�g�b�v�J�b�g�d��  2009/10/05 Kameda

Public Const CDOPCALC_DOPWFDIA As Double = 150#        '�h�[�vwf���a
Public Const CDOPCALC_DOPWFTHICK As Double = 625#      '�h�[�vwf���
Public Const CDOPCALC_FILMTHICK_10 As Double = 1#      '����(1.0��)
Public Const CDOPCALC_FILMTHICK_05 As Double = 0.5     '����(0.5��)
Public Const CDOPCALC_FILMTHICK_01 As Double = 0.1     '����(0.1��)
Public Const CDOPCALC_FILMTHICK_0015 As Double = 0.015   '����(0.015��)   2011/08/25 Kameda
Public Const CDOPCALC_K0 As Double = 0.0007            'K0
Public Const CDOPCALC_MOL As Double = 140.283          '���q��
Public Const CDOPCALC_DENSITY As Double = 3.185        '���x
Private Const DIA_KUBUN = "300"
Public Const SISU_14 As Integer = 14                   '�w��(14��Œ�)
Public gBlock As typ_Select_BlockData
Public gKEISAN As typ_XDOPE_KeisanData
'--------------------------------------------------------------------
'�T�v      :���f�h�[�v�f�[�^�擾
'���Ұ�    :�ϐ���      ,IO   ,�^          ,����
'          :BLOCK      ,IO  �@�@,typ_Select_BlockData   ,�u���b�N�ڍ�
'          :KEISAN     ,O  �@�@,typ_XDOPE_KeisanData   ,�v�Z����
'����      :�v�Z�f�[�^�A�h�[�v�ʂ����߂�
'����      :
'///////////////////////////////////////////////////
Public Function GetXLDopeRyo(Block As typ_Select_BlockData, Keisan As typ_XDOPE_KeisanData) As Boolean
    
    
    Dim dModTmp         As Double   '��]�v�Z
    Dim dRoundTmp       As Double   '�l�̌ܓ��v�Z
    Dim i               As Integer
    Dim iSisu           As Integer
    Dim dblDope         As Double
    
    GetXLDopeRyo = False
    
    
    'EXCEL�v�Z���̒ʂ�Ɍv�Z
    ' �g�b�v����          = 3*D7*1000/(3.1416*((D5/10/2)^2)*2.328)*10
    ' �����Z�t�̐�        = $D$6*1000/2.57
    ' �_���ʒu���㗦      = ($D$7+(($D$5/10/2)^2*3.14*2.328*D8/10)/1000)/$D$6
    ' �_������[N]         = D9/(D27*(1-D33)^(D27-1))
    ' �_���h�[�v(Si3N4)�� = D34*D32*D28/(4*6.02*10^23)*1000
    ' Si3N4�d��(1�ʂ�)    = (D21/20)^2*3.1416*D23/10000*D29*1000*2
    ' Si3N4�d��(0.5�ʂ�)  = (E21/20)^2*3.1416*E23/10000*E29*1000*2
    ' Si3N4�d��(0.1�ʂ�)  = (F21/20)^2*3.1416*F23/10000*F29*1000*2
    ' 1�ʂ�               = ROUNDDOWN(D35/D24,0)
    ' 0.5�ʂ�             = ROUNDDOWN(MOD(D35,D24)/E24,0)
    ' 0.1�ʂ�             = ROUND(MOD(MOD(D35,D24),E24)/F24,0)
    ' �h�[�vWF��Տd��(D,E,F) = (D21/20)^2*3.1416*D22/10000*2.328
    ' ���v�h�[�vWF�d��    = D36*(D25+D24/1000)
    ' �h�[�v(Si3N4)��     = D24*B15+E24*C15+F24*D15
    ' ����[N]             = D39*(4*6.02*10^23)/(D32*D28*1000)
    With Keisan
        '�w���ϊ�    2011/08/25 Kameda
        iSisu = Block.intDOPESISU
        dblDope = Block.dblDopeRyo
        While (dblDope < 1 Or dblDope >= 10) And dblDope <> 0
            If dblDope < 1 Then    '1�����̏ꍇ
                dblDope = dblDope * 10
                iSisu = iSisu - 1       '�w����-1
            ElseIf dblDope >= 10 Then  '10�ȏ�̏ꍇ
                dblDope = dblDope / 10
                iSisu = iSisu + 1       '�w����+1
            End If
        Wend
        '�g�b�v����
        .dblTopLength = 3 * Block.dblWGHTTO * 1000 / (cdblPI * ((Block.dblDIA / 10 / 2) ^ 2) * 2.328) * 10
        
        '�����Z�t�̐�
        .dblInitLiquid = Block.dblCHARGE * 1000 / 2.57
        
        '�_���ʒu���㗦
        .dblPulRate = (Block.dblWGHTTO + ((Block.dblDIA / 10 / 2) ^ 2 * 3.14 * 2.328 * Block.dblAIMPOS / 10) / 1000) / Block.dblCHARGE
        
        '�_������[N]
        .dblNeraiInit = (Block.dblDopeRyo * 10 ^ Block.intDOPESISU) / (CDOPCALC_K0 * (1 - .dblPulRate) ^ (CDOPCALC_K0 - 1))
        
        '�_���h�[�v(Si3N4)��
        .dblNeraiDope = .dblNeraiInit * .dblInitLiquid * CDOPCALC_MOL / (4 * 6.02 * 10 ^ 23) * 1000
        
        
        'Si3N4�d��
        .dblSi3N4Weight_10 = (CDOPCALC_DOPWFDIA / 20) ^ 2 * cdblPI * CDOPCALC_FILMTHICK_10 / 10000 * CDOPCALC_DENSITY * 1000 * 2
        .dblSi3N4Weight_05 = (CDOPCALC_DOPWFDIA / 20) ^ 2 * cdblPI * CDOPCALC_FILMTHICK_05 / 10000 * CDOPCALC_DENSITY * 1000 * 2
        .dblSi3N4Weight_01 = (CDOPCALC_DOPWFDIA / 20) ^ 2 * cdblPI * CDOPCALC_FILMTHICK_01 / 10000 * CDOPCALC_DENSITY * 1000 * 2
        .dblSi3N4Weight_0015 = (CDOPCALC_DOPWFDIA / 20) ^ 2 * cdblPI * CDOPCALC_FILMTHICK_0015 / 10000 * CDOPCALC_DENSITY * 1000 * 2
        
        '1 ��m�h�[�vWF����
        .dblMaisu_10 = .dblNeraiDope / .dblSi3N4Weight_10
        
        '��Տd��
        .dblSi3N4KibanWt = (CDOPCALC_DOPWFDIA / 20) ^ 2 * cdblPI * CDOPCALC_DOPWFTHICK / 10000 * 2.328
        If iSisu > 11 Then
            '���f6"WF1.0�ʖ���
            .intXDopeRyo_10 = Int(.dblNeraiDope / .dblSi3N4Weight_10)
            
            '���f6"WF0.5�ʖ���
            'VB��Mod�͐�����Ԃ����ߏ����ł̗]������߂�
            dModTmp = .dblNeraiDope - .dblSi3N4Weight_10 * CDbl(.intXDopeRyo_10)
            .intXDopeRyo_05 = Int(dModTmp / .dblSi3N4Weight_05)
            
            '���f6"WF0.1�ʖ���
            dModTmp = dModTmp - .dblSi3N4Weight_05 * CDbl(.intXDopeRyo_05)
            dRoundTmp = dModTmp / .dblSi3N4Weight_01
            .intXDopeRyo_01 = Int(dRoundTmp + 0.5)    '�l�̌ܓ�
            ' �h�[�v(Si3N4)��     = D24*B15+E24*C15+F24*D15
            .dblDopeRyo = .dblSi3N4Weight_10 * .intXDopeRyo_10 + .dblSi3N4Weight_05 * .intXDopeRyo_05 + .dblSi3N4Weight_01 * .intXDopeRyo_01
        
        Else
            '���f6"WF0.015�ʖ���  2011/08/25 Kameda
            .intXDopeRyo_0015 = Int(.dblNeraiDope / .dblSi3N4Weight_0015)
            ' �h�[�v(Si3N4)��     = D24*B15+E24*C15+F24*D15
            .dblDopeRyo = .dblSi3N4Weight_0015 * .intXDopeRyo_0015
        End If
        
        ' ���v�h�[�vWF�d��    = D36*(D25+D24/1000)
        .dblDopeKei = .dblMaisu_10 * (.dblSi3N4KibanWt + .dblSi3N4Weight_10 / 1000)
        
        ''�������т����߂�
        'cmhc001d_SelectXDope Keisan, Block.StrCryNum   C6���擾
        
        '���і������珉���Z�x���v�Z����    H001���擾�@2012/01/27 test Kame
        If SelectWFCount(Block.StrCryNum, Keisan) Then
            .dblDopeRyo = .dblSi3N4Weight_10 * .intXDopeRyoJ_10 + .dblSi3N4Weight_01 * .intXDopeRyoJ_01 + .dblSi3N4Weight_05 * .intXDopeRyo_05 + .dblSi3N4Weight_0015 * .intXDopeRyoJ_0015
        End If
        
        ' ����[N]             = D39*(4*6.02*10^23)/(D32*D28*1000)
        .dblSyokiN = .dblDopeRyo * (4 * 6.02 * 10 ^ 23) / (.dblInitLiquid * CDOPCALC_MOL * 1000)
    End With

    GetXLDopeRyo = True
    
End Function
'--------------------------------------------------------------------
'�T�v      :���f�Z�x�擾(�ʒu�w��)
'���Ұ�    :�ϐ���      ,IO   ,�^          ,����
'          :WGHTTO      ,I  �@double       �g�b�v�d��
'          :HDIA        ,I  �@double       ���グ�a
'          :INPOS       ,I  �@integer      �����ʒu
'          :CHARGE      ,I  �@double       �`���[�W��
'          :NOUDO       ,I  �@double     �@����[N]�v�Z�l
'          :TOPCUT      ,I  �@double     �@�g�b�v�J�b�g�d��    2009/10/05 Kameda
'�Ԃ�l    :���f�Z�x
'����      :
'///////////////////////////////////////////////////
Public Function GetNNoudo(WGHTTO As Double, HDIA As Double, INPOS As Integer, CHARGE As Double, _
                          NOUDO As Double, TOPCUT As Double) As Double
    Dim dblPulWt As Double
    Dim dblPuRitu As Double
    
        GetNNoudo = 0
        
        '���グ�d��= �g�b�v�d��+((���グ�a/10/2)^2*3.14*2.328*�����ʒu/10)/1000
        'dblPulWt = WGHTTO + ((HDIA / 10 / 2) ^ 2 * 3.14 * 2.328 * INPOS / 10) / 1000
        dblPulWt = WGHTTO + TOPCUT + ((HDIA / 10 / 2) ^ 2 * 3.14 * 2.328 * INPOS / 10) / 1000   '2009/10/05 Kameda
        
        '���グ�� = ���グ�d��/�`���[�W��
        dblPuRitu = Round(dblPulWt / CHARGE, 6)
        
        '�Z�x = ����N*K0*(1-���グ��)^(K0-1)
        If dblPuRitu < 1 Then
            GetNNoudo = NOUDO * CDOPCALC_K0 * (1 - dblPuRitu) ^ (CDOPCALC_K0 - 1)
        End If
    
        
End Function
'--------------------------------------------------------------------
'�T�v      :���f�Z�x�擾  �ʒu(KANKAKU)��
'���Ұ�    :�ϐ���      ,IO   ,�^          ,����
'          :BLOCK      ,I  �@�@,typ_Select_BlockData   ,�u���b�N�ڍ�
'          :NOUDO      ,I  �@�@����[N]�v�Z�l
'          :NNOUDO     ,O  �@�@�Z�x
'          :KANKAKU    ,I  �@�@�ʒu�Ԋu
'          :MINLEN     ,I  �@�@�\���J�n�ʒu
'          :MAXLEN     ,I  �@�@�\���I���ʒu
'����      :�v�Z�f�[�^�A�h�[�v�ʂ����߂�
'����      :
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
            sPos = Block.intBOPOS    '�Ō�̍s�̓{�g���ʒu
        End If
        
        '���グ�d��= �g�b�v�d��+((���グ�a/10/2)^2*3.14*2.328*�����ʒu/10)/1000
        'NNOUDO(i).dblPulWt = Block.dblWGHTTO + ((Block.dblDIA / 10 / 2) ^ 2 * 3.14 * 2.328 * sPos / 10) / 1000
        '���グ�d��= �g�b�v�d��+�g�b�v�J�b�g�d��+((���グ�a/10/2)^2*3.14*2.328*�����ʒu/10)/10002009/10/05 Kameda
        NNOUDO(i).dblPulWt = Block.dblWGHTTO + Block.dblTOPCUT + ((Block.dblDIA / 10 / 2) ^ 2 * 3.14 * 2.328 * sPos / 10) / 1000
        
        '���グ�� = ���グ�d��/�`���[�W��
        NNOUDO(i).dblPuRitu = Round(NNOUDO(i).dblPulWt / Block.dblCHARGE, 6)
        
        '�Z�x = ����N*K0*(1-���グ��)^(K0-1)
        If NNOUDO(i).dblPuRitu < 1 Then
            NNOUDO(i).dblNnoudo = NOUDO * CDOPCALC_K0 * (1 - NNOUDO(i).dblPuRitu) ^ (CDOPCALC_K0 - 1)
        End If
        '�����ʒu
        NNOUDO(i).intXtalPos = sPos
        sPos = sPos + KANKAKU
        
    Next
    
    GetNNoudoALL = True
    
End Function
'--------------------------------------------------------------------
'�T�v      :���f�Z�x�擾(�ؒf�w���p)
'���Ұ�    :�ϐ���      ,IO   ,�^          ,����
'          :INPOS       ,I  �@integer      �����ʒu
'          :CRYNUM      ,I  �@string       �����ԍ�
'          :SISU        ,I  �@integer      �Z�x�w��
'�Ԃ�l    :���f�Z�x
'����      :
'///////////////////////////////////////////////////
Public Function GetNNoudoSIJI(CRYNUM As String, INPOS() As Long, Sisu As Integer, NOUDO() As Double) As Boolean
    
    Dim dblPulWt As Double
    Dim dblPuRitu As Double
    Dim Block As typ_Select_BlockData
    Dim sNoudo As Double
    Dim iCnt As Integer
    
        GetNNoudoSIJI = False
        
        '����N�Z�x�擾
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
'�T�v      :�_������[N]�擾
'���Ұ�    :�ϐ���     ,IO   ,�^          ,����
'          :BLOCK      ,IO  �@�@,typ_Select_BlockData   ,�u���b�N�ڍ�
'          :KEISAN     ,O  �@�@,typ_XDOPE_KeisanData
'����      :�v�Z�f�[�^�A�h�[�v�ʂ����߂�
'����      :
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
                    ' ����[N]
                    'GetSyokiNoudo = .dblDopeRyo * (4 * 6.02 * 10 ^ 23) / (.dblInitLiquid * CDOPCALC_MOL * 1000)
                    GetSyokiNoudo = .dblSyokiN
                End If
            End With
        End If
    End If
    
End Function

'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------
'�T�v      :�e�[�u���uTBCMH001�v,�uXSDC1�v���璊�o����
'���Ұ�    :�ϐ���      ,IO     ,�^                     ,����
'          :rec         ,O  �@�@,typ_Select_BlockData   ,�u���b�N�ڍ�
'          :�߂�l      ,O      ,FUNCTION_RETURN        ,���o�̐���
'����      :
'����      :

Public Function SelectBlock(rec As typ_Select_BlockData) As FUNCTION_RETURN

Dim sql As String       'SQL�S��
Dim rs As OraDynaset    'RecordSet
Dim cnt As Integer      '����
Dim i As Long
        
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmec078_SQL.bas -- Function SelectBlock"

    SelectBlock = FUNCTION_RETURN_FAILURE
    
    '***** ����w�����сiTBCMH001�j *****
    ''SQL��g�ݗ��Ă�
    sql = "Select  nvl(CHARGE,0) CHARGE "
    sql = sql & " ,HINBAN,NMNOREVNO,NFACTORY,NOPECOND"
    sql = sql & " ,nvl(CRYDOPCL,' ') CRYDOPCL,nvl(CRYDOPVL,0) CRYDOPVL"
    sql = sql & " ,nvl(DOPN,0) DOPN,nvl(DPNI,' ') DPNI "
    sql = sql & " ,nvl(AIMPOS,0) AIMPOS "
    sql = sql & " From TBCMH001"
    sql = sql & " Where (UPINDNO ='" & left$(rec.StrCryNum, 7) & "0" & Mid(rec.StrCryNum, 9, 1) & "')"
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.EOF Then
        GoTo proc_exit
    End If
     ''���o���ʂ��i�[����
    With rec
        .dblCHARGE = rs("CHARGE") / 1000  ' �`���[�W��
        .StrHinban = rs("HINBAN")         ' �i��
        .intMNOREVNO = rs("NMNOREVNO")
        .StrFactory = rs("NFACTORY")
        .StrOpeCond = rs("NOPECOND")
        .strCRYDOPCL = rs("CRYDOPCL")     ' �����h�[�v���
        '.dblCRYDOPVL = rs("CRYDOPVL")     ' �����h�[�v��
        .dblCRYDOPVL = rs("DOPN")         ' �����h�[�v��
        .dblDopeRyo = rs("DOPN")         ' �����h�[�v��     '2010/01/29 add Kameda
        .dblAIMPOS = rs("AIMPOS")         ' �˂炢�ʒu
        .dblDIA = CDOPCALC_DIA            ' ���グ�a
        If Trim(rs("DPNI")) = "" Then
            .intDOPESISU = 0
        Else
            .intDOPESISU = rs("DPNI")         ' �w��
        End If
    End With
    rs.Close
    
    '2009/10/19 add Kameda
    If rec.dblCHARGE = 0 Then
        GoTo proc_exit
    End If
    
    '***** ���d�ʁiXSDC1�j *****
    sql = "select nvl(WGHTTOC1,0) WGHTTOC1,(nvl(DIA1C1,0)+nvl(DIA2C1,0)+nvl(DIA3C1,0))/3 as DIA "
    sql = sql & ", nvl(PUTCUTWC1,0) PUTCUTWC1 "        '2009/10/05 Kameda
    sql = sql & "from XSDC1 "
    sql = sql & "where XTALC1 = '" & left(rec.StrCryNum, 9) & "000" & "' "
    ''�f�[�^�𒊏o����
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
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    SelectBlock = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'�T�v      :CODEA9���猨�d�ʂ𓾂�
'���Ұ�    :�ϐ���      ,IO ,�^       ,����
'          :KTWeight()  ,O  ,double   ,���d��
'          :�߂�l      ,O  ,FUNCTION_RETURN,���o�̐���
'����      :
'����      :
Public Function GetKTWeight() As Double
Dim rs      As OraDynaset               '���oRecordDynaset
Dim rsCnt   As Integer                  'ں��޶���
Dim sql     As String                   'SQL��
Dim i       As Integer                  'ٰ�߶���

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc010a.bas -- Function GetKTWeight"
    
    'SQL���̍쐬
    sql = "select CTR02A9 from KODA9 where SYSCA9='K' and SHUCA9='A7' and CODEA9 = '" & DIA_KUBUN & "' "

    '�f�[�^�̒��o
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    '''���o���R�[�h�����݂��Ȃ��ꍇ
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
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function
'�T�v      :�����h�[�p���g�������ю擾
'���Ұ�    :�ϐ���        ,IO ,�^                   ,����
'          :rec           ,O  ,typ_XDOPE_KeisanData ,���o���R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN      ,���o�̐���
'����      :
'����      :
Public Function cmhc001d_SelectXDope(rec As typ_XDOPE_KeisanData, CRYNUM As String) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim rs As OraDynaset    'RecordSet
Dim i As Long
Dim sCryNum As String   '2010/01/29 add Kameda


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmec078_SQL.bas -- Function cmhc001d_SelectXDope"

    cmhc001d_SelectXDope = FUNCTION_RETURN_FAILURE
    
    ''�f�[�^�𒊏o     '2010/01/29 Kameda �������x�Ή�
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
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'--------------------------------------------------------------------
'�T�v      :��������w���\���p���f�h�[�v�v�Z
'���Ұ�    :�ϐ���     ,IO   ,�^          ,����
'          :BLOCK      ,IO  �@�@,typ_Select_BlockData   ,�u���b�N�ڍ�
'          :KEISAN     ,O  �@�@,typ_XDOPE_KeisanData
'����      :�v�Z�f�[�^�A�h�[�v�ʂ����߂�
'����      :
'///////////////////////////////////////////////////
Public Function GetNNoudoCC600(Block As typ_Select_BlockData, Keisan As typ_XDOPE_KeisanData) As Double
    
    
    'Dim KEISAN As typ_XDOPE_KeisanData    2010/01/29 del Kameda
    'Dim dblSyokiN As Double
    
    GetNNoudoCC600 = 0
    
    With Keisan
        'Block.dblDopeRyo = Mid(Block.dblCRYDOPVL, 1, InStr(Block.dblCRYDOPVL, ".") - 1)
        Block.dblDopeRyo = Block.dblCRYDOPVL
        'If GetXLDopeRyo(Block, KEISAN) Then
            ' ����[N]
            'dblSyokiN = .dblDopeRyo * (4 * 6.02 * 10 ^ 23) / (.dblInitLiquid * CDOPCALC_MOL * 1000)
        '    dblSyokiN = .dblSyokiN  2010/01/29 del Kameda
        'End If
        '2009/10/05 Kameda
        'GetNNoudoCC600 = GetNNoudo(Block.dblWGHTTO, Block.dblDIA, Block.intTOPOS, Block.dblCHARGE, dblSyokiN) / 10 ^ val(SISU_14)
        GetNNoudoCC600 = GetNNoudo(Block.dblWGHTTO, Block.dblDIA, Block.intTOPOS, Block.dblCHARGE, .dblSyokiN, Block.dblTOPCUT) / 10 ^ val(SISU_14)
    End With
    
End Function

'�T�v      :���f�K�i�擾
'���Ұ��@�@:�ϐ���        ,IO ,�^                                  ,����
'      �@�@:spec_N  �@�@�@,IO ,
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN                   �@,�ǂݍ��݂̐���
'����      :
'����      :2009/09/03
Public Function GetSpecN(HinSpecN As typ_spec_N) As FUNCTION_RETURN

    Dim rs  As OraDynaset
    Dim sql As String

    '' �G���[�n���h���̐ݒ�
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
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    GetSpecN = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------
'�T�v      :�e�[�u���uTBCMH001�v���璊�o����
'���Ұ�    :�ϐ���      ,IO     ,�^                     ,����
'          :rec         ,O  �@�@,StrCryNum              ,�u���b�NID
'          :�߂�l      ,O      ,WF����
'����      :
'����      :

Public Function SelectWFCount(StrCryNum As String, rec As typ_XDOPE_KeisanData) As Boolean

Dim sql As String       'SQL�S��
Dim rs As OraDynaset    'RecordSet
Dim cnt As Integer      '����
Dim i As Long
        
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc016_SQL.bas -- Function SelectWFCount"

    SelectWFCount = False
    
    '***** ����w�����сiTBCMH001�j *****
    ''SQL��g�ݗ��Ă�
    sql = "Select  WFCOUNT10, WFCOUNT05, WFCOUNT01, WFCOUNT0015 "
    sql = sql & " From TBCMH001"
    sql = sql & " Where (UPINDNO ='" & left$(StrCryNum, 7) & "0" & Mid(StrCryNum, 9, 1) & "')"
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.EOF Then
        GoTo proc_exit
    End If
     
     ''���o���ʂ��i�[����
    With rec
        If IsNull(rs("WFCOUNT10")) Then .intXDopeRyoJ_10 = 0 Else .intXDopeRyoJ_10 = rs("WFCOUNT10")
        If IsNull(rs("WFCOUNT05")) Then .intXDopeRyoJ_05 = 0 Else .intXDopeRyoJ_05 = rs("WFCOUNT05")
        If IsNull(rs("WFCOUNT01")) Then .intXDopeRyoJ_01 = 0 Else .intXDopeRyoJ_01 = rs("WFCOUNT01")
        If IsNull(rs("WFCOUNT0015")) Then .intXDopeRyoJ_0015 = 0 Else .intXDopeRyoJ_0015 = rs("WFCOUNT0015")
    End With
    
    rs.Close
    
    SelectWFCount = True
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


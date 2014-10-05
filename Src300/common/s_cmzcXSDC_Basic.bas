Attribute VB_Name = "s_cmzcXSDC_Basic"
Option Explicit
''���H�d�l,���H���э\����
''�����̒Ⴂ����MIN�l��������
Private Type Judg_Kakou
    DPTH(2) As Double   '�����グ�����̏ꍇ�f�[�^�͂ЂƂ������݂��Ȃ�
    WIDH(2) As Double   '�����グ�����̏ꍇ�f�[�^�͂ЂƂ������݂��Ȃ�
    TOP(2) As Double
    TAIL(2) As Double
    pos As String * 2
End Type


'��{���
Public Type typ_BasicCd
    nextCode     As String       '���H��
    nowCode      As String       '���H��
    changeHinban As String       '�ύX�i��
    DIAMETER     As Double       '���a
    UPWEIGHT     As Double       '����d�ʁ@04/09/30 ooba
    LENGFREE     As Integer      '�ذ�����@04/12/22 ooba
End Type

Private Type typ_XSDCA_c_flg
    Entry As Boolean
    Furyo As Integer            '�s�ǐ�
    Index_F As Long             '�s�Ǎ\���̲��ޯ��
    FuryoW As Long              '�s�Ǐd��
End Type

'2002/08/29 �ǉ�
Private Type typ_KoteiInf        '�H�����
    Wkkt As String               '�H��
    Maco As Integer               '������
End Type
Private msL2Wkkt As String        '�O�O�H��
Private msL2Maco As String        '�O�O�H��������
    


Private strNxtCd As String
Private strNowCd As String
Private strChgHin As String
Private dblDiameter As Double
Private regFLG As String
Private intFuryoLen As Integer          '�s�ǒ���
Private intFuryoWei As Long             '�s�Ǐd��
Private CC300Flg As Boolean             '���ݍH����CC300���ǂ���
Private SXLflg As Boolean               'SXL�Ǘ�(XSDCB)�ւ̓o�^(�X�V)���s�����ǂ���
Private PutWtFlg As Integer             '�w����d�ʂ�i�Ԗ��Ɉ������l�x��o�^���邩�@04/09/30 ooba
Private lPutWeight As Long              '����d�ʁ@04/09/30 ooba
Private lTotalPwt As Long               '�Z�o���v�d��(����d��)�@04/09/30 ooba
Private iTotalLen As Integer            '�i�ԍ��v�����@04/09/30 ooba
Private iFreeLen As Integer             '�ذ�����@04/12/22 ooba

'2002/09/05�@m.tomita
Private CB410Flg As Boolean             '���ݍH����CB410���ǂ���


Public Const FACTORYCD As Integer = 42  '�����H��


'�T�v     :�V�c�a�ւ̏����݊�{�p�^�[���������s��
'���Ұ�   :�ϐ���           ,IO  ,�^                   ,����
'         :p_typXSDC2_b     ,I   ,typ_XSDC2            ,��������(��ۯ�)�O�H�����я��
'         :p_typXSDCA_b()   ,I   ,typ_XSDCA            ,��������(�i��)�O�H�����я��
'         :p_typXSDC2_c     ,I   ,typ_XSDC2            ,��������(��ۯ�)�o�^���
'         :p_typXSDCA_c()   ,I   ,typ_XSDCA            ,��������(�i��)�o�^���
'         :p_typXSD4upd()   ,I   ,typ_XSDC4            ,�s�Ǔ���o�^���
'         :p_typBasicCd     ,I   ,typ_BasicCd          ,��{���
'         :p_strErrMsg      ,O   ,String               ,�װү����
'         :�߂�l           ,O    ,FUNCTION_RETURN      ,�V�c�a�ւ̏����݂̐���
'����     :

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
    Dim recCnt As Long             '���R�[�h��
    Dim C4Cnt As Long              '���R�[�h�� C4

    Dim typXSDC2upd_b As typ_XSDC2_Update
    Dim typXSDC2upd_c As typ_XSDC2_Update
    Dim typXSDCAupd_b() As typ_XSDCA_Update
    Dim typXSDCAupd_c() As typ_XSDCA_Update
    Dim typXSDC4upd() As typ_XSDC4_Update
    Dim typXSDC3upd_b() As typ_XSDC3_Update
    Dim typXSDC3upd_c() As typ_XSDC3_Update
    
    Dim intBlockLen As Integer         '��ۯ��̒���
    Dim intSyoriKaisu As Integer    '���ݏ�����
    
''a.��{���̃Z�b�g
    With p_typBasicCd
        strNxtCd = .nextCode
        strNowCd = .nowCode
        strChgHin = .changeHinban
        dblDiameter = .DIAMETER
    End With
    
 '���H����CC300�Ȃ��׸޾��
    If strNowCd = "CC300" Then
        CC300Flg = True
    Else
        CC300Flg = False
    End If
    
'��2002/09/05 M.TOMITA �ǉ�
 '���H����CB410�Ȃ��׸޾��
'    If strNowCd = "CB410" Then
    '�����������Ή��@2003/09/01 ooba
    If strNowCd = "CB410" And strNxtCd = "CC600" Then
        CB410Flg = True
    Else
        CB410Flg = False
    End If
'��2002/09/05 M.TOMITA
    
 'SXL�o�^�׸޾��
    SXLflg = False
    If strNowCd = "CC710" Or strNowCd = "CC720" Or strNowCd = "CW730" Then
    '�����w�����́AWF�������o�A�������ύX�̏ꍇ
        SXLflg = True
    ElseIf strNowCd = "CC700" Then
    '�����ŏI���o�� SXLID �������Ă�ꍇ
        If Left(p_typXSDCA_b(1).SXLIDCA, 1) <> vbNullChar And Trim(p_typXSDCA_b(1).SXLIDCA) <> "" Then
            SXLflg = True
        End If
    End If
 
    '' �w����d�ʂ�i�Ԗ��Ɉ������l�x��o�^���邩���f����@04/09/30 ooba
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
 
 '���ݏ����񐔂̎擾
 ' �����񐔎擾���W�b�N�ύX 2002/11/21 tuku
    If CC300Flg = True Or CB410Flg = True Then
        intSyoriKaisu = GetGNMACOC(p_typXSDCA_c(1).CRYNUMCA, strNxtCd) '��������(�i��)�o�^������ۯ�ID���L�[�Ɏ擾
        lPutWeight = CLng(p_typBasicCd.UPWEIGHT)  '����d�ʎ擾�@04/09/30 ooba
        iFreeLen = p_typBasicCd.LENGFREE  '�ذ�����擾�@04/12/22 ooba
        
    ElseIf Left(p_typXSDC2_b.CRYNUMC2, 1) <> vbNullChar Then
        intSyoriKaisu = GetGNMACOC(p_typXSDC2_b.CRYNUMC2, strNxtCd) '��������(��ۯ�)�O�H��������ۯ�ID���L�[�Ɏ擾
        lPutWeight = GetPutWeight(p_typXSDC2_b.XTALC2)  '����d�ʎ擾�@04/09/30 ooba
        
    ElseIf Left(p_typXSDC2_c.CRYNUMC2, 1) <> vbNullChar Then
        intSyoriKaisu = GetGNMACOC(p_typXSDC2_c.CRYNUMC2, strNxtCd) '��������(��ۯ�)�o�^������ۯ�ID���L�[�Ɏ擾
        lPutWeight = GetPutWeight(p_typXSDC2_c.XTALC2)  '����d�ʎ擾�@04/09/30 ooba
        
    Else
        If UBound(p_typXSD4) = 0 Then
           intSyoriKaisu = 1
        Else
           intSyoriKaisu = GetGNMACOC(p_typXSD4(1).XTALC4, strNxtCd)   '�s�Ǔ���o�^������ۯ�ID���L�[�Ɏ擾
           lPutWeight = GetPutWeight(Left(p_typXSD4(1).XTALC4, 9) & "000")  '����d�ʎ擾�@04/09/30 ooba
        End If
    End If
    
  ' ���H���d�|��Ή��@2002/11/25 tuku
    If strNxtCd = strNowCd Then
        intSyoriKaisu = intSyoriKaisu + 1
    End If
    
'�O�O�H�����̎擾
'�O�X�H���̎擾���@�ύX�@2002/11/22 tuku
    If CC300Flg = True Or CB410Flg = True Then                               'CC300,CB410
        msL2Wkkt = ""                      '�O�O�H��
        msL2Maco = ""                      '�O�O�H��������
    ElseIf Left(p_typXSDCA_b(1).CRYNUMCA, 1) = vbNullChar Then            'b�Ȃ�
        msL2Wkkt = ""                      '�O�O�H��
        msL2Maco = ""                      '�O�O�H��������
    ElseIf Left(p_typXSDC2_c.CRYNUMC2, 1) = vbNullChar Then        'c�Ȃ�
        msL2Wkkt = p_typXSDCA_b(1).NEWKNTCA   '�O�O�H��
        msL2Maco = p_typXSDCA_b(1).NEMACOCA   '�O�O�H��������
    ElseIf p_typXSDCA_b(1).CRYNUMCA = p_typXSDC2_c.CRYNUMC2 Then       'b=c
        msL2Wkkt = p_typXSDCA_b(1).NEWKNTCA   '�O�O�H��
        msL2Maco = p_typXSDCA_b(1).NEMACOCA   '�O�O�H��������
    Else                                                            'b<>c
        msL2Wkkt = ""                      '�O�O�H��
        msL2Maco = ""                      '�O�O�H��������
    End If
    
'    If Left(p_typXSDC2_b.CRYNUMC2, 1) = vbNullChar Then            'b�Ȃ�
'        msL2Wkkt = ""                      '�O�O�H��
'        msL2Maco = ""                      '�O�O�H��������
'    ElseIf Left(p_typXSDC2_c.CRYNUMC2, 1) = vbNullChar Then        'c�Ȃ�
'        msL2Wkkt = p_typXSDC2_b.NEWKNTC2   '�O�O�H��
'        msL2Maco = p_typXSDC2_b.NEMACOC2   '�O�O�H��������
'    ElseIf p_typXSDC2_b.CRYNUMC2 = p_typXSDC2_c.CRYNUMC2 Then       'b=c
'        msL2Wkkt = p_typXSDC2_b.NEWKNTC2   '�O�O�H��
'        msL2Maco = p_typXSDC2_b.NEMACOC2   '�O�O�H��������
'    Else                                                            'b<>c
'        msL2Wkkt = ""                      '�O�O�H��
'        msL2Maco = ""                      '�O�O�H��������
'    End If
    
    
''b.�O�H���̎���(������)���擾����
    '��������(��ۯ�)
    With typXSDC2upd_b
        .CRYNUMC2 = p_typXSDC2_b.CRYNUMC2                 '��ۯ�ID
        .KCNTC2 = p_typXSDC2_b.KCNTC2                     '�H���ʉߘA��
        .INPOSC2 = p_typXSDC2_b.INPOSC2
        '.KCNTC2 = p_typXSDC2_b.KCNTC2 + 1                 '�H���A�ԁ{�P���ăZ�b�g
'2002/10/16-----------------------------------------------------------------------------------��1-�@
'        If p_typXSDC2_b.GNWKNTC2 = strNxtCd Then          '���H���Ȃ�{�P���ăZ�b�g
'            '.GNMACOC2 = p_typXSDC2_b.GNMACOC2 + 1         '���ݏ�����
'            .GNMACOC2 = intSyoriKaisu                     '���ݏ�����
'            .NEWKNTC2 = p_typXSDC2_b.NEWKNTC2             '�ŏI�ʉߍH��
'            .NEMACOC2 = p_typXSDC2_b.NEMACOC2             '�ŏI�ʉߏ�����
'        Else
'            .GNMACOC2 = 1
'            .NEWKNTC2 = p_typXSDC2_b.GNWKNTC2
'            .NEMACOC2 = p_typXSDC2_b.GNMACOC2
'        End If
'        .GNWKNTC2 = strNxtCd                                   '���ݍH��
 '2002/11/21 tuku �����񐔎擾���W�b�N�ύX
        '�O�H�������X�V�����Ȃ�(�����ł�)
        .GNMACOC2 = p_typXSDC2_b.GNMACOC2                      '���ݏ�����
        .NEWKNTC2 = p_typXSDC2_b.NEWKNTC2                      '�ŏI�ʉߍH��
        .NEMACOC2 = p_typXSDC2_b.NEMACOC2                      '�ŏI�ʉߏ�����
        .GNWKNTC2 = p_typXSDC2_b.GNWKNTC2                      '���ݍH��
'2002/10/16-----------------------------------------------------------------------------------��1-�@
        .GNLC2 = p_typXSDC2_b.GNLC2                            '���ݒ���
        .GNWC2 = p_typXSDC2_b.GNWC2                            '���ݏd��
'       .GNWC2 = WeightOfCylinder(dblDiameter, .GNLC2)         '���ݏd��
        '' SUMIT�����^�d�ʎ擾�@04/09/30 ooba
        If PutWtFlg > 0 Then
            .SUMITLC2 = p_typXSDC2_b.SUMITLC2                  'SUMIT����
            .SUMITWC2 = p_typXSDC2_b.SUMITWC2                  'SUMIT�d��
        End If
        .KAKOUBC2 = p_typXSDC2_b.KAKOUBC2                      '���H�敪
        .LSTATBC2 = p_typXSDC2_b.LSTATBC2                      '�ŏI��ԋ敪
        .RSTATBC2 = p_typXSDC2_b.RSTATBC2                      '������ԋ敪
        .LDFRBC2 = p_typXSDC2_b.LDFRBC2                        '�i���敪
        .HOLDBC2 = p_typXSDC2_b.HOLDBC2                        'ΰ��ދ敪
        .KANKC2 = p_typXSDC2_b.KANKC2                          '�����敪
        If p_typXSDC2_b.CHGC2 <> 0 Then
            .CHGC2 = p_typXSDC2_b.CHGC2                        '����ޗ�
            .KEIDAYC2 = p_typXSDC2_b.KEIDAYC2                  '�v����t
        End If
        '2003.06.11 (SPK)Y.katabami�@�D��x���V�K�Đ؋敪���ǉ�
        .CUTCNTC2 = p_typXSDC2_b.CUTCNTC2                      '�V�K�Đ؋敪
        .PRIORITYC2 = p_typXSDC2_b.PRIORITYC2                  '�D��x
        .RPCRYNUMC2 = p_typXSDC2_b.RPCRYNUMC2                  '�e��ۯ�ID   2005/11
        .HOLDCC2 = p_typXSDC2_b.HOLDCC2                        'ΰ��ރR�[�h  2006/03
        .HOLDKTC2 = p_typXSDC2_b.HOLDKTC2                      'ΰ��ލH��  2006/03
    End With
                               
    '��������(�i��)
    recCnt = UBound(p_typXSDCA_b)
    ReDim typXSDCAupd_b(recCnt)
    For i = 1 To recCnt
        With typXSDCAupd_b(i)
            .CRYNUMCA = p_typXSDCA_b(i).CRYNUMCA             '��ۯ�ID
            .HINBCA = p_typXSDCA_b(i).HINBCA                 '�i��
            .INPOSCA = p_typXSDCA_b(i).INPOSCA               '�������J�n�ʒu
            .REVNUMCA = p_typXSDCA_b(i).REVNUMCA             '���i�ԍ������ԍ�
            .FACTORYCA = p_typXSDCA_b(i).FACTORYCA           '�H��
            .OPECA = p_typXSDCA_b(i).OPECA                   '���Ə���
            .KCKNTCA = p_typXSDCA_b(i).KCKNTCA               '�H���A��
            '.KCKNTCA = typXSDC2upd_b.KCNTC2                  '��ۯ��̍H���A�Ԃ��Z�b�g
            '.GNMACOCA = intSyoriKaisu + (i - 1)              '���ݏ�����(2ں��ޖڂ���{�P)
            .GNMACOCA = p_typXSDCA_b(i).GNMACOCA              '���ݏ�����
'2002/10/16-----------------------------------------------------------------------------------��1-�A
'            If p_typXSDCA_b(i).GNWKNTCA = strNxtCd Then      '���H���Ȃ�{�P���ăZ�b�g
'                '.GNMACOCA = p_typXSDCA_b(i).GNMACOCA + 1     '���ݏ�����
'                .NEWKNTCA = p_typXSDCA_b(i).NEWKNTCA         '�ŏI�ʉߍH��
'                .NEMACOCA = p_typXSDCA_b(i).NEMACOCA         '�ŏI�ʉߏ�����
'            Else
'                '.GNMACOCA = 1
'                .NEWKNTCA = p_typXSDCA_b(i).GNWKNTCA
'                .NEMACOCA = p_typXSDCA_b(i).GNMACOCA
'            End If
'            .GNWKNTCA = strNxtCd                             '���ݍH��
            '2002/11/21 tuku �����񐔎擾���W�b�N�ύX
            '�O�H�������X�V�����Ȃ�(�����ł�)
            .NEWKNTCA = p_typXSDCA_b(i).NEWKNTCA                   '�ŏI�ʉߍH��
            .NEMACOCA = p_typXSDCA_b(i).NEMACOCA                   '�ŏI�ʉߏ�����
            .GNWKNTCA = p_typXSDCA_b(i).GNWKNTCA                   '���ݍH��
'2002/10/16-----------------------------------------------------------------------------------��1-�A
            .SXLIDCA = p_typXSDCA_b(i).SXLIDCA               'SXLID
            .XTALCA = p_typXSDCA_b(i).XTALCA                 '�����ԍ�
            .GNLCA = p_typXSDCA_b(i).GNLCA                   '���ݒ���
            .GNWCA = p_typXSDCA_b(i).GNWCA                   '���ݏd��
'           .GNWCA = WeightOfCylinder(dblDiameter, .GNLCA)   '���ݏd��
            '' SUMIT�����^�d�ʎ擾�@04/09/30 ooba
            If PutWtFlg > 0 Then
                .SUMITLCA = p_typXSDCA_b(i).SUMITLCA         'SUMIT����
                .SUMITWCA = p_typXSDCA_b(i).SUMITWCA         'SUMIT�d��
            End If
            .KAKOUBCA = p_typXSDCA_b(i).KAKOUBCA             '���H�敪
            .LSTATBCA = p_typXSDCA_b(i).LSTATBCA             '�ŏI��ԋ敪
            .RSTATBCA = p_typXSDCA_b(i).RSTATBCA             '������ԋ敪
            .LDFRBCA = p_typXSDCA_b(i).LDFRBCA               '�i���敪
            .HOLDBCA = p_typXSDCA_b(i).HOLDBCA               'ΰ��ދ敪
            .KANKCA = p_typXSDCA_b(i).KANKCA                 '�����敪
            If p_typXSDCA_b(i).CHGCA <> 0 Then
                .CHGCA = p_typXSDCA_b(i).CHGCA               '����ޗ�
                .KEIDAYCA = p_typXSDCA_b(i).KEIDAYCA         '�v����t
            End If
            '2003.06.11 (SPK)Y.katabami�@��\�i�ԁ��V�K�Đ؋敪���ǉ�
            .CUTCNTCA = p_typXSDCA_b(i).CUTCNTCA             '�V�K�Đ؋敪
            .HINBFLGCA = p_typXSDCA_b(i).HINBFLGCA           '��\�i�ԃt���O
            .RPCRYNUMCA = p_typXSDCA_b(i).RPCRYNUMCA         '�e��ۯ�ID    2005/11
            .HOLDCCA = p_typXSDCA_b(i).HOLDCCA               'ΰ��ރR�[�h   2006/03
            .HOLDKTCA = p_typXSDCA_b(i).HOLDKTCA             'ΰ��ލH��   2006/03
        End With
    Next
    
    
''c.�o�^���R�[�h���쐬
    Select Case strNowCd
        Case "CC300", "CC310", "CB410", "CC450", "CC600"
            '���ݏ����񐔂̎擾(��������(��ۯ�)�o�^������ۯ�ID���L�[�Ɏ擾)
            If CC300Flg = False And p_typXSDC2_b.CRYNUMC2 <> p_typXSDC2_c.CRYNUMC2 _
                And Left(p_typXSDC2_c.CRYNUMC2, 1) <> vbNullChar Then  '�ؒf�Ή�
                intSyoriKaisu = GetGNMACOC(p_typXSDC2_c.CRYNUMC2, strNxtCd)
            End If
            

            
            '��������(��ۯ�)���
            With typXSDC2upd_c
                .CRYNUMC2 = p_typXSDC2_c.CRYNUMC2                          '��ۯ�ID
'                .KCNTC2 = 1                                                '�H���A�ԂɂP���Z�b�g
                .XTALC2 = p_typXSDC2_c.XTALC2                              '�����ԍ�
                .INPOSC2 = p_typXSDC2_c.INPOSC2                            '�������J�n�ʒu
                .NEWKNTC2 = strNowCd                                       '�ŏI�ʉߍH��
                .GNWKNTC2 = strNxtCd                                       '���ݍH��
                '.GNMACOC2 = 1                                              '���ݏ�����
                .GNMACOC2 = intSyoriKaisu                                  '���ݏ�����
                '2002/11/21 tuku �����񐔎擾���W�b�N�ύX
                .NEMACOC2 = GetNEMACOC2(typXSDC2upd_c.CRYNUMC2)            '�ŏI�ʉߏ�����
                .GNDAYC2 = Format(Now, "yyyy/mm/dd hh:mm:ss")              '���ݏ�������
                .GNLC2 = p_typXSDC2_c.GNLC2                                '���ݒ���
'                .GNWC2 = WeightOfCylinder(dblDiameter, p_typXSDC2_c.GNLC2) '���ݏd��

                '' �d�ʓo�^�ύX�@04/09/30 ooba START ===========================================>
                If PutWtFlg = 1 Then
                    '����d�ʾ��
                    .GNWC2 = lPutWeight
                    'SUMIT�����^�d�ʂ��
                    .SUMITLC2 = p_typXSDC2_c.GNLC2                                'SUMIT����
                    .SUMITWC2 = WeightOfCylinder(dblDiameter, p_typXSDC2_c.GNLC2) 'SUMIT�d��
                Else
                    '�����v�Z�d�ʾ��
                    .GNWC2 = WeightOfCylinder(dblDiameter, p_typXSDC2_c.GNLC2)
                End If
                '' �d�ʓo�^�ύX�@04/09/30 ooba END =============================================>
                
'                .GNMC2 =                                                  �f���ݖ���
                .KAKOUBC2 = p_typXSDC2_c.KAKOUBC2                          '���H�敪
                If p_typXSDC2_c.CHGC2 <> 0 Then
                    .CHGC2 = p_typXSDC2_c.CHGC2                                '����ޗ�
                    .KEIDAYC2 = p_typXSDC2_c.KEIDAYC2                          '�v����t
                End If
                '2003.06.11 (SPK)Y.katabami�@�D��x���V�K�Đ؋敪���ǉ�
                .CUTCNTC2 = p_typXSDC2_c.CUTCNTC2                      '�V�K�Đ؋敪
                .PRIORITYC2 = p_typXSDC2_c.PRIORITYC2                  '�D��x
                .RPCRYNUMC2 = p_typXSDC2_c.RPCRYNUMC2                  '�e��ۯ�ID   2005/11
            
                '2006/03 HOLD
                .HOLDBC2 = p_typXSDC2_b.HOLDBC2
                .HOLDCC2 = p_typXSDC2_b.HOLDCC2
                .HOLDKTC2 = p_typXSDC2_b.HOLDKTC2
            End With
            
            
            '��������(�i��)���
            recCnt = UBound(p_typXSDCA_c)
            ReDim typXSDCAupd_c(recCnt)
                        
            '' �i�ԑS�̒�����ā@04/09/30 ooba
            If PutWtFlg > 0 Then
                For i = 1 To recCnt
                    iTotalLen = iTotalLen + p_typXSDCA_c(i).GNLCA
                Next
'                If iTotalLen = 0 Then
'                    ExecBscProcess = FUNCTION_RETURN_FAILURE
'                    p_strErrMsg = "�d�ʓo�^���s(XSDC)"
'                    GoTo proc_exit
'                End If
            End If
            
            For i = 1 To recCnt
                With typXSDCAupd_c(i)
                    .CRYNUMCA = p_typXSDCA_c(i).CRYNUMCA                          '��ۯ�ID
                    .HINBCA = p_typXSDCA_c(i).HINBCA                              '�i��
                    .INPOSCA = p_typXSDCA_c(i).INPOSCA                            '�������J�n�ʒu
                    .REVNUMCA = p_typXSDCA_c(i).REVNUMCA                          '���i�ԍ������ԍ�
                    .FACTORYCA = p_typXSDCA_c(i).FACTORYCA                        '�H��
                    .OPECA = p_typXSDCA_c(i).OPECA                                '���Ə���
                    .SXLIDCA = p_typXSDCA_c(i).SXLIDCA                            'SXLID
                    .XTALCA = p_typXSDCA_c(i).XTALCA                              '�����ԍ�
                    .NEWKNTCA = strNowCd                                          '�ŏI�ʉߍH��
                    .NEMACOCA = GetNEMACOC(.CRYNUMCA, CInt(.INPOSCA))             '�ŏI�ʉߏ����񐔁@�f2002/08/29
                    .GNWKNTCA = strNxtCd                                          '���ݍH��
                    '.GNMACOCA = 1                                                 '���ݏ�����
                    '.GNMACOCA = intSyoriKaisu + (i - 1)              '���ݏ�����(2ں��ޖڂ���{�P)
                    .GNMACOCA = intSyoriKaisu                                     '���ݏ�����
                    .GNLCA = p_typXSDCA_c(i).GNLCA                                '���ݒ���
'                    .GNWCA = WeightOfCylinder(dblDiameter, p_typXSDCA_c(i).GNLCA) '���ݏd��

                    '' �d�ʓo�^�ύX�@04/09/30 ooba START =========================================>
                    If PutWtFlg = 1 Then
                        '����d�ʾ��
                        If i = recCnt Then
                            '�v�Z�덷���Ō�̏d�ʂɌv��
                            .GNWCA = lPutWeight - lTotalPwt
                        Else
                            '����d�ʂ�i�Ԓ����ň�
                            If iTotalLen <> 0 Then
                                .GNWCA = Int(lPutWeight * (p_typXSDCA_c(i).GNLCA / iTotalLen))
                            Else
                                .GNWCA = 0
                            End If
                            lTotalPwt = lTotalPwt + .GNWCA
                        End If
                        'SUMIT�����^�d�ʂ��
                        .SUMITLCA = p_typXSDCA_c(i).GNLCA                                'SUMIT����
                        .SUMITWCA = WeightOfCylinder(dblDiameter, p_typXSDCA_c(i).GNLCA) 'SUMIT�d��
                    Else
                        '�����v�Z�d�ʾ��
                        .GNWCA = WeightOfCylinder(dblDiameter, p_typXSDCA_c(i).GNLCA)
                    End If
                    '' �d�ʓo�^�ύX�@04/09/30 ooba END ===========================================>
                    
'                    .GNMCA =                                                     �f���ݖ���
                    .KAKOUBCA = p_typXSDCA_c(i).KAKOUBCA                          '���H�敪
                    If p_typXSDCA_c(i).CHGCA <> 0 Then
                        .CHGCA = p_typXSDCA_c(i).CHGCA                                '����ޗ�
                        .KEIDAYCA = p_typXSDCA_c(i).KEIDAYCA                          '�v����t
                    End If
                    '2003.06.11 (SPK)Y.katabami�@��\�i�ԁ��V�K�Đ؋敪���ǉ�
                    .CUTCNTCA = p_typXSDCA_c(i).CUTCNTCA             '�V�K�Đ؋敪
                    .HINBFLGCA = p_typXSDCA_c(i).HINBFLGCA           '��\�i�ԃt���O
                    .RPCRYNUMCA = p_typXSDCA_c(i).RPCRYNUMCA         '�e��ۯ�ID    2005/11
                
                    '2006/03 HOLD
                    .HOLDBCA = p_typXSDCA_b(1).HOLDBCA
                    .HOLDCCA = p_typXSDCA_b(1).HOLDCCA
                    .HOLDKTCA = p_typXSDCA_b(1).HOLDKTCA
                End With
            Next
            
            '�o�^���R�[�h(c)���Ȃ��ꍇ
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
    
    
    '�s�Ǔ�����
    intFuryoLen = 0
    intFuryoWei = 0
    
    recCnt = UBound(p_typXSD4)
    ReDim typXSDC4upd(recCnt)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  �@CC400�̉~������d�ʍ폜������g�ݍ���                            �@�@   '
'    �������A�{���̕s�ǔ�������͂��鎞�́A���̏����ł͑Ή��s�@�@�@�@�@�@�@�@ '
'�@�@��ʏ�ŕs�ǈʒu����͂��K�v                                            '                                                    '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    �����P�T�N�T���T���@�@�@�@�@�@�@�@�@�_�@�O�Y
If strNowCd = "CC400" Then
   recCnt = UBound(p_typXSDCA_b)
   ReDim typXSDC4upd(recCnt)
    For i = 1 To recCnt
         With typXSDC4upd(i)
            .XTALC4 = p_typXSDCA_b(i).CRYNUMCA        '��ۯ�ID
            .INPOSC4 = p_typXSDCA_b(i).INPOSCA        '�������J�n�ʒu
            .FCODEC4 = "088"                          '���탍�X
            .MACOC4 = intSyoriKaisu                   '���ݏ�����
            .HINBC4 = p_typXSDCA_b(i).HINBCA          '�i��
            .PUCUTLC4 = 0                             '�s�ǒ���
'            .PUCUTWC4 = p_typXSDCA_b(i).GNWCA - WeightOfCylinder(dblDiameter, CInt(p_typXSDCA_b(i).GNLCA))
            
            '' �s�Ǐd�ʓo�^�ύX�@04/09/30 ooba START ===========================================>
            If PutWtFlg = 1 Then
                .PUCUTWC4 = p_typXSDCA_b(i).SUMITWCA - WeightOfCylinder(dblDiameter, CInt(p_typXSDCA_b(i).SUMITLCA))
            Else
                .PUCUTWC4 = p_typXSDCA_b(i).GNWCA - WeightOfCylinder(dblDiameter, CInt(p_typXSDCA_b(i).GNLCA))
            End If
            '' �s�Ǐd�ʓo�^�ύX�@04/09/30 ooba END =============================================>

        End With
        intFuryoLen = intFuryoLen + typXSDC4upd(i).PUCUTLC4
        intFuryoWei = intFuryoWei + typXSDC4upd(i).PUCUTWC4
    Next i
Else
    For i = 1 To recCnt
        With typXSDC4upd(i)
            .XTALC4 = p_typXSD4(i).XTALC4             '��ۯ�ID
            .INPOSC4 = p_typXSD4(i).INPOSC4           '�������J�n�ʒu
            '.MACOC4 = intSyoriKaisu + (i - 1)         '���ݏ�����(2ں��ޖڂ���{�P)
            .MACOC4 = intSyoriKaisu                   '���ݏ�����
            .HINBC4 = p_typXSD4(i).HINBC4             '�i��
            .PUCUTLC4 = p_typXSD4(i).PUCUTLC4         '�s�ǒ���
            .PUCUTWC4 = WeightOfCylinder(dblDiameter, CInt(p_typXSD4(i).PUCUTLC4))
        End With
        intFuryoLen = intFuryoLen + typXSDC4upd(i).PUCUTLC4
        intFuryoWei = intFuryoWei + typXSDC4upd(i).PUCUTWC4
  
    Next i
End If

        
''e.���f����(�o�^�l��o�^�\���̂ɃZ�b�g����)
''f.�o�^����

    If regFLG = "Y" Then
        '��ۯ����������߂�
        intBlockLen = 0
        If CC300Flg = True Then '���ݍH��CC300�̏ꍇ
            recCnt = UBound(p_typXSDCA_c)
            For i = 1 To recCnt
                intBlockLen = intBlockLen + typXSDCAupd_c(i).GNLCA
            Next
        Else
            intBlockLen = p_typXSDC2_c.GNLC2
        End If
        
        If intBlockLen - intFuryoLen > 0 Then   'c-d
        ''�Z�b�g�p�^�[���T-�@
            If SetPattern1(typXSDC2upd_b, typXSDCAupd_b, typXSDC2upd_c, typXSDCAupd_c, typXSDC4upd, _
                                                p_strErrMsg) = FUNCTION_RETURN_FAILURE Then
                ExecBscProcess = FUNCTION_RETURN_FAILURE
                p_strErrMsg = p_strErrMsg
                GoTo proc_exit
            End If
        Else
        ''�Z�b�g�p�^�[���U-�@ (�����b�g����)
            If SetPattern2(typXSDC2upd_b, typXSDCAupd_b, typXSDC2upd_c, typXSDCAupd_c, typXSDC4upd, _
                                                p_strErrMsg) = FUNCTION_RETURN_FAILURE Then
                ExecBscProcess = FUNCTION_RETURN_FAILURE
                p_strErrMsg = p_strErrMsg
                GoTo proc_exit
            End If
        End If
    Else
        
        If p_typXSDC2_b.GNLC2 - intFuryoLen > 0 Then   'b-d
        ''�Z�b�g�p�^�[���T-�A
            If SetPattern1(typXSDC2upd_b, typXSDCAupd_b, typXSDC2upd_c, typXSDCAupd_c, typXSDC4upd, _
                                                p_strErrMsg) = FUNCTION_RETURN_FAILURE Then
                ExecBscProcess = FUNCTION_RETURN_FAILURE
                p_strErrMsg = p_strErrMsg
                GoTo proc_exit
            End If
        Else
        ''�Z�b�g�p�^�[���U-�A (�����b�g����)
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
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    gErr.HandleError
    ExecBscProcess = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v     :�Z�b�g�p�^�[���T�̏������s��(�ʏ폈��)
'���Ұ�   :�ϐ���           ,IO  ,�^                   ,����
'         :p_Block_b        ,I   ,typ_XSDC2_Update     ,��������(��ۯ�)�O�H�����я��
'         :p_Hinban_b()     ,I   ,typ_XSDCA_Update     ,��������(�i��)�O�H�����я��
'         :p_Block_c        ,I   ,typ_XSDC2_Update     ,��������(��ۯ�)�o�^���
'         :p_Hinban_c()     ,I   ,typ_XSDCA_Update     ,��������(�i��)�o�^���
'         :p_Furyo()        ,I   ,typ_XSDC4_Update     ,�s�Ǔ���o�^���
'         :p_Error          ,O   ,String               ,�װү����
'         :�߂�l           ,O    ,FUNCTION_RETURN      ,�V�c�a�ւ̏����݂̐���
'����     :
Private Function SetPattern1(p_Block_b As typ_XSDC2_Update, p_Hinban_b() As typ_XSDCA_Update, _
                        p_Block_c As typ_XSDC2_Update, p_Hinban_c() As typ_XSDCA_Update, _
                        p_Furyo() As typ_XSDC4_Update, p_Error As String) As FUNCTION_RETURN
                        
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcXSDC_Basic.bas -- Function SetPattern1"
                        
    Dim i As Long
    Dim j As Long
    Dim n As Long
    
    Dim strWhere As String                   '�X�V����WHERE��
    Dim RET As FUNCTION_RETURN               '�߂�l
    Dim strErrMsg As String                  '�G���[���b�Z�[�W
    Dim recCnt As Long                       '���R�[�h��
    Dim recCnt2 As Long                      '���R�[�h��
    Dim recCnt3 As Long                      '���R�[�h��
    Dim AccFlg As Boolean                    'bں����׸�(b��c����v������������)
    Dim XSDCA_c_flg() As typ_XSDCA_c_flg     'cں����׸�(�o�^/�X�V�����ƑΉ�����s�Ǐ��)
    Dim sDbName As String                    'ð��ٖ�
    Dim typKotei() As typ_XSDC3_Update       '�H�����ѓo�^���
    Dim strMotHin As String                  '�U�֕i��(��)
    Dim typSXL() As typ_XSDCB_Update         '��������(�i��)���
    Dim intLen As Integer                    'SXL����
    Dim intSyoriKaisu As Integer    '���ݏ�����
    Dim rs2 As OraDynaset           '���R�[�h�Z�b�g
    Dim sql As String
    Dim fullHinban As tFullHinban            '�ؽ�ٶ�۸ފi�グ���̍ŐV�i�ԁ@2003/11/10 ooba
    Dim p_Block_b_bar As typ_XSDC2_Update    'Bar�o�חp��������(��ۯ�)�o�^���@04/09/27 ooba
    Dim p_Hinban_b_bar() As typ_XSDCA_Update 'Bar�o�חp��������(�i��)�o�^���@04/09/27 ooba
    Dim typKotei_bar() As typ_XSDC3_Update   'Bar�o�חp�H�����ѓo�^���@04/09/27 ooba
    
    If regFLG = "Y" Then
''���Z�b�g�p�^�[���T-�@
        
        If CC300Flg = False Then   '���H��CC300����ۯ��̓o�^���s��Ȃ�
    
    '�ᕪ������(��ۯ�)-XSDC2��
        
            '�s�ǐ�������Ό��Z����
            p_Block_c.GNLC2 = p_Block_c.GNLC2 - intFuryoLen
            p_Block_c.GNWC2 = WeightOfCylinder(dblDiameter, CInt(p_Block_c.GNLC2))
    
            If p_Block_b.CRYNUMC2 = p_Block_c.CRYNUMC2 Then  'b=c
            '�o�^ں��ޏ��(c)��UPDATE
                sDbName = "(XSDC2)"
                With p_Block_c
                    .KCNTC2 = p_Block_b.KCNTC2 + 1    '�H���A�Ԃ��{�P���ăZ�b�g
'2002/10/16 �R�����g--------------------------------------------------------------------------------��1-�B
'                    .GNMACOC2 = p_Block_b.GNMACOC2    '���ݏ�����
'                    .NEWKNTC2 = p_Block_b.NEWKNTC2    '�ŏI�ʉߍH��
'                    .NEMACOC2 = p_Block_b.NEMACOC2    '�ŏI�ʉߏ�����
'2002/10/16 �R�����g--------------------------------------------------------------------------------��1-�B
                End With
                strWhere = "WHERE CRYNUMC2 = '" & p_Block_b.CRYNUMC2 & "'"
                If UpdateXSDC2(p_Block_c, strWhere) = FUNCTION_RETURN_FAILURE Then
                    SetPattern1 = FUNCTION_RETURN_FAILURE
                    p_Error = GetMsgStr("EAPLY") & sDbName
                    GoTo proc_exit
                End If
                
            Else                                            'b<>c
            
            'b��ں��ނ𐶎��׸�=1��UPDATE
                If Left(p_Block_b.CRYNUMC2, 1) <> vbNullChar Then '(��������ꍇ)
                    p_Block_b.LIVKC2 = "1"
                    p_Block_b.KCNTC2 = p_Block_b.KCNTC2 + 1    '�H���A�Ԃ��{�P���ăZ�b�g
                    strWhere = "WHERE CRYNUMC2 = '" & p_Block_b.CRYNUMC2 & "'"
                    If UpdateXSDC2(p_Block_b, strWhere) = FUNCTION_RETURN_FAILURE Then
                        SetPattern1 = FUNCTION_RETURN_FAILURE
                        p_Error = GetMsgStr("EAPLY") & sDbName
                        GoTo proc_exit
                    End If
                End If
                
            '�o�^ں��ޏ��(c)��INSERT
                p_Block_c.KCNTC2 = 1    '�H���A�ԂɂP���Z�b�g
                If strNowCd = "CC310" Then
                    p_Block_c.KCNTC2 = 2    '�H���A�ԂɂQ���Z�b�g(��ۯ��̑O�H�����Ȃ�)
                End If
                'p_Block_c.GNMACOCA = 1         '���ݏ�����
                If CreateXSDC2(p_Block_c, strErrMsg) = FUNCTION_RETURN_FAILURE Then
                    SetPattern1 = FUNCTION_RETURN_FAILURE
                    p_Error = strErrMsg
                    GoTo proc_exit
                End If
            End If
            
        End If
        
        
    '�ᕪ������(�i��)-XSDCA��
        sDbName = "(XSDCA)"
        
        recCnt = UBound(p_Hinban_b)
        recCnt2 = UBound(p_Hinban_c)
        recCnt3 = UBound(p_Furyo)
        
        ReDim XSDCA_c_flg(recCnt2)
        ReDim typKotei(recCnt2)
        
        '�s�ǐ�������Ό��Z����
        For i = 1 To recCnt2
            For j = 1 To recCnt3
                If p_Hinban_c(i).CRYNUMCA = p_Furyo(j).XTALC4 _
                            And p_Hinban_c(i).HINBCA = p_Furyo(j).HINBC4 _
                            And p_Hinban_c(i).INPOSCA = p_Furyo(j).INPOSC4 Then
                    '�����Z�b�g
                    p_Hinban_c(i).GNLCA = p_Hinban_c(i).GNLCA - p_Furyo(j).PUCUTLC4
                    '�d�ʃZ�b�g
'                   p_Hinban_c(i).GNWCA = WeightOfCylinder(dblDiameter, CInt(p_Hinban_c(i).GNLCA))
                    p_Hinban_c(i).GNWCA = p_Hinban_c(i).GNWCA - p_Furyo(j).PUCUTWC4


                    XSDCA_c_flg(i).Furyo = p_Furyo(j).PUCUTLC4   '�s�ǒ������
                    XSDCA_c_flg(i).Index_F = j                   'index���

                    Exit For
                Else
                    XSDCA_c_flg(i).Index_F = -1                   'index���
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
                    
                    '�o�^ں��ޏ��(c)��UPDATE
                    With p_Hinban_c(j)
                        .KCKNTCA = p_Block_c.KCNTC2           '��ۯ��̍H���A�Ԃ��Z�b�g
'2002/10/16 �R�����g--------------------------------------------------------------------------------��1-�C
'                        .GNMACOCA = p_Hinban_b(i).GNMACOCA    '���ݏ�����
'                        .NEWKNTCA = p_Hinban_b(i).NEWKNTCA    '�ŏI�ʉߍH��
'                        .NEMACOCA = p_Hinban_b(i).NEMACOCA    '�ŏI�ʉߏ�����
'2002/10/16 �R�����g--------------------------------------------------------------------------------��1-�C
                    End With
                    
                    '�s�Ǔ���ں��ނɕi�ԏ����Z�b�g
                    If XSDCA_c_flg(j).Index_F > 0 Then
                        With p_Furyo(XSDCA_c_flg(j).Index_F)
                            .KCKNTC4 = p_Hinban_c(j).KCKNTCA        '�H���A��
                            .REVNUMC4 = p_Hinban_c(j).REVNUMCA      '���i�ԍ������ԍ�
                            .FACTORYC4 = p_Hinban_c(j).FACTORYCA    '�H��
                            .OPEC4 = p_Hinban_c(j).OPECA            '���Ə���
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
                    
                '��H�����ѓo�^��
                    Call SetXSDC3(typKotei(n), p_Hinban_c(j), XSDCA_c_flg(j))
                    If CreateXSDC3(typKotei(n), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                        SetPattern1 = FUNCTION_RETURN_FAILURE
                        p_Error = strErrMsg
                        GoTo proc_exit
                    End If
                    
                    n = n + 1
                    AccFlg = True
                    XSDCA_c_flg(j).Entry = True   '�X�V�׸�(��v����ں��ރA��)
                    Exit For
'                Else
'                    XSDCA_c_flg(j).Entry = False   '�X�V�׸�(��v����ں��ރi�V)
                End If
            Next
            
            '��v���Ȃ�ں���(b)�̍X�V  b<>c
            If Left(p_Hinban_b(i).CRYNUMCA, 1) <> vbNullChar Then '��������ꍇ
                If AccFlg = False Then
                    
                    'b��ں��ނ𐶎��׸�=1��UPDATE
                    p_Hinban_b(i).LIVKCA = "1"
                    p_Hinban_b(i).KCKNTCA = p_Block_b.KCNTC2 + 1    '�H���A��(�O�H��)���{�P���ăZ�b�g
                    strWhere = "WHERE CRYNUMCA = '" & p_Hinban_b(i).CRYNUMCA
                    strWhere = strWhere & "' AND HINBCA = '" & p_Hinban_b(i).HINBCA
                    strWhere = strWhere & "' AND INPOSCA = " & p_Hinban_b(i).INPOSCA
                    
                    If UpdateXSDCA(p_Hinban_b(i), strWhere) = FUNCTION_RETURN_FAILURE Then
                        SetPattern1 = FUNCTION_RETURN_FAILURE
                        p_Error = GetMsgStr("EAPLY") & sDbName
                        GoTo proc_exit
                    End If
                        
'                    '�H�����ѓo�^(���̃��R�[�h)
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
        
        ReDim p_Hinban_b_bar(recCnt2)  'CC300����0�p��������(�i��)�o�^���@04/12/17 ooba
        ReDim typKotei_bar(recCnt2)    'CC300����0�p�H�����ѓo�^���@04/12/17 ooba
        
        '��v���Ȃ�ں���(c)�̓o�^
        For i = 1 To recCnt2
             If XSDCA_c_flg(i).Entry = False Then
                '�H���A�Ԏ擾
                '��2002/09/05 M.TOMITA CB410�������ǉ�
'                If CC300Flg = True Then
                If CC300Flg = True Or CB410Flg = True Then
                '��2002/09/05 M.TOMITA CB410�������ǉ�
                    p_Hinban_c(i).KCKNTCA = 1                  '�H���A�ԂɂP���Z�b�g
                    
                ElseIf recCnt <> 0 And p_Hinban_b(1).CRYNUMCA = p_Hinban_c(1).CRYNUMCA Then
                    p_Hinban_c(i).KCKNTCA = p_Block_c.KCNTC2   '��ۯ��̍H���A�Ԃ��Z�b�g
                    
                Else
                    p_Hinban_c(i).KCKNTCA = 1                  '�H���A�ԂɂP���Z�b�g
                End If
                
                '�s�Ǔ���ں��ނɕi�ԏ����Z�b�g
                If XSDCA_c_flg(i).Index_F > 0 Then
                    With p_Furyo(XSDCA_c_flg(i).Index_F)
                        .KCKNTC4 = p_Hinban_c(i).KCKNTCA       '�H���A��
                        .REVNUMC4 = p_Hinban_c(i).REVNUMCA     '���i�ԍ������ԍ�
                        .FACTORYC4 = p_Hinban_c(i).FACTORYCA   '�H��
                        .OPEC4 = p_Hinban_c(i).OPECA           '���Ə���
                        .WKKTC4 = strNowCd
                    End With
                End If
                
'2002/08/29----------------------------------------------------
'                '�o�^
'                If CreateXSDCA(p_Hinban_c(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
'                    SetPattern1 = FUNCTION_RETURN_FAILURE
'                    p_Error = strErrMsg
'                    GoTo proc_exit
'                End If
                
                With p_Hinban_c(i)
                    If CheckUniqueRecord(.CRYNUMCA, .HINBCA, CInt(.INPOSCA)) = True Then
                        '�o�^
                        If CreateXSDCA(p_Hinban_c(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                            SetPattern1 = FUNCTION_RETURN_FAILURE
                            p_Error = strErrMsg
                            GoTo proc_exit
                        End If
                    Else
                        '�X�V
                        .LIVKCA = "0"      '�����t���O��"0"���Z�b�g
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
                
                
                '��H�����ѓo�^��
                Call SetXSDC3(typKotei(n), p_Hinban_c(i), XSDCA_c_flg(i))
                If CreateXSDC3(typKotei(n), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                    SetPattern1 = FUNCTION_RETURN_FAILURE
                    p_Error = strErrMsg
                    GoTo proc_exit
                End If
                
                '' CC310�̏d��0���э쐬�@04/12/22 ooba START ================================>
                
                '�H��CC300���ذ����0�̏ꍇ
                If CC300Flg = True And iFreeLen = 0 Then
                
                    p_Hinban_b_bar(i) = p_Hinban_c(i)
                    typKotei_bar(n) = typKotei(n)
                    
                    With p_Hinban_b_bar(i)
                        .KCKNTCA = CInt(p_Hinban_c(i).KCKNTCA) + 1  '�H���A��
'                        .NEKKNTCA = p_Hinban_c(i).GNKKNTCA          '�ŏI�ʉߊǗ��H��
'                        .NEWKNTCA = p_Hinban_c(i).GNWKNTCA          '�ŏI�ʉߍH��
'                        .NEWKKBCA = p_Hinban_c(i).GNWKKBCA          '�ŏI�ʉߍ�Ƌ敪
'                        .NEMACOCA = p_Hinban_c(i).GNMACOCA          '�ŏI�ʉߏ�����
                        .GNWKNTCA = "CB210"                         '���ݍH��
                        .GNMACOCA = 1                               '���ݏ�����
                        .LSTATBCA = "R"                             '�ŏI��ԋ敪(����)
                        .RSTATBCA = "M"                             '������ԋ敪(���Ď���҂�)
                        .LDFRBCA = "1"                              '�i���敪(����)
                        .LIVKCA = "1"                               '�����敪(��ۯ�)
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
                        .KCNTC3 = p_Hinban_b_bar(i).KCKNTCA         '�H���A��
                        .KNKTC3 = Space(5)                          '�Ǘ��H��
                        .WKKTC3 = Space(5)                          '�H��
                        .WKKBC3 = p_Hinban_b_bar(i).NEWKKBCA        '��Ƌ敪
                        .MACOC3 = p_Hinban_b_bar(i).NEMACOCA        '������
                        .FRKNKTC3 = p_Hinban_c(i).NEKKNTCA          '(���)�Ǘ��H��
                        .FRWKKTC3 = p_Hinban_c(i).NEWKNTCA          '(���)�H��
                        .FRWKKBC3 = p_Hinban_c(i).NEWKKBCA          '(���)��Ƌ敪
                        .FRMACOC3 = p_Hinban_c(i).NEMACOCA          '(���)������
                        .TOWNKTC3 = p_Hinban_b_bar(i).GNKKNTCA      '(���o)�Ǘ��H��
                        .TOWKKTC3 = p_Hinban_b_bar(i).GNWKNTCA      '(���o)�H��
                        .TOMACOC3 = p_Hinban_b_bar(i).GNMACOCA      '(���o)������
                        .FRLC3 = p_Hinban_c(i).GNLCA                '�������
                        .FRWC3 = p_Hinban_c(i).GNWCA                '����d��
                        .FULC3 = p_Hinban_c(i).GNLCA                '�s�ǒ���
                        .FUWC3 = p_Hinban_c(i).GNWCA                '�s�Ǐd��
                        .TOLC3 = "0"                                '���o����
                        .TOWC3 = "0"                                '���o�d��
                        .SUMITLC3 = "0"                             'SUMIT����
                        .SUMITWC3 = "0"                             'SUMIT�d��
                    End With
                    
                    If CreateXSDC3(typKotei_bar(n), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                        SetPattern1 = FUNCTION_RETURN_FAILURE
                        p_Error = strErrMsg
                        GoTo proc_exit
                    End If
                End If
                '' CC310�̏d��0���э쐬�@04/12/22 ooba END ==================================>
                    
                n = n + 1
            End If
        Next
        
    '��s�Ǔ���o�^��
        For i = 1 To recCnt3
            If p_Furyo(i).KCKNTC4 = "" Then
            '��v����i�Ԃ���������ں��ނ̏����Z�b�g
                With p_Furyo(i)
                    If p_Block_b.CRYNUMC2 = p_Block_c.CRYNUMC2 Then  'b=c (��ۯ�)
                        .KCKNTC4 = p_Block_b.KCNTC2 + 1    '�H���A�Ԃ��{�P���ăZ�b�g
                    Else
                        .KCKNTC4 = 1                       '�H���A�ԂɂP���Z�b�g
                    End If
                    .WKKTC4 = strNowCd                     '�H��
                End With
            End If
            
            '�o�^
            If p_Furyo(i).PUCUTLC4 <> 0 Then
                If CreateXSDC4(p_Furyo(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                    SetPattern1 = FUNCTION_RETURN_FAILURE
                    p_Error = strErrMsg
                    GoTo proc_exit
                End If
            End If
        Next
        
    Else
    
''���Z�b�g�p�^�[���T-�A
    
    '�ᕪ������(��ۯ�)-XSDC2��
        sDbName = "(XSDC2)"
        
        '�����A�d�ʃZ�b�g
'        p_Block_b.GNLC2 = p_Block_b.GNLC2 - intFuryoLen
'        p_Block_b.GNWC2 = p_Block_b.GNWC2 - intFuryoWei
''       p_Block_b.GNWC2 = WeightOfCylinder(dblDiameter, CInt(p_Block_b.GNLC2))
        
        '' �����^�d�ʓo�^�ύX�@04/09/30 ooba START ================================>
        If PutWtFlg = 1 Then
            p_Block_b.SUMITLC2 = p_Block_b.SUMITLC2 - intFuryoLen
            p_Block_b.SUMITWC2 = p_Block_b.SUMITWC2 - intFuryoWei
        Else
            p_Block_b.GNLC2 = p_Block_b.GNLC2 - intFuryoLen
            p_Block_b.GNWC2 = p_Block_b.GNWC2 - intFuryoWei
        End If
        '' �����^�d�ʓo�^�ύX�@04/09/30 ooba END ==================================>
        
        '�����ŏI���o�̏ꍇ(�ŏI��ԋ敪���)
        With p_Block_b
            If strNowCd = "CC700" Then
                If strNxtCd = "CC705" Then
                    .LSTATBC2 = "B"   '�ŏI��ԋ敪(BAR�o��)
                    .KANKC2 = "2"     '�����敪(�I��)
                    .LIVKC2 = "1"
                Else
                    .LSTATBC2 = "W"   '�ŏI��ԋ敪(WF�o��)
                End If
            ElseIf strNowCd = "CB320" Then '�i�グ���ɗ�����ԋ敪��ʏ�ɖ߂� 2002/12/17 tuku
                    .RSTATBC2 = "T"
            End If
        End With
        '�H���R�[�h���X�V����
        With p_Block_b
            .GNWKNTC2 = strNxtCd         ' ���ݍH��
            .NEWKNTC2 = strNowCd         ' �ŏI�ʉߍH��
        End With
        
        '�X�V
        p_Block_b.KCNTC2 = p_Block_b.KCNTC2 + 1    '�H���A�Ԃ��{�P���ăZ�b�g
        
        '�����񐔎擾���W�b�N�ύX 2002/11/21 tuku  START
        intSyoriKaisu = GetGNMACOC(p_Block_b.CRYNUMC2, strNxtCd) '��������(��ۯ�)�O�H��������ۯ�ID���L�[�Ɏ擾
        If strNxtCd = strNowCd Then         ' ���H���d�|��Ή��@2002/11/25 tuku
            intSyoriKaisu = intSyoriKaisu + 1
        End If
        p_Block_b.GNMACOC2 = intSyoriKaisu
        p_Block_b.NEMACOC2 = GetNEMACOC2(p_Block_b.CRYNUMC2)
        '�����񐔎擾���W�b�N�ύX 2002/11/21 tuku  END
        
        strWhere = "WHERE CRYNUMC2 = '" & p_Block_b.CRYNUMC2 & "'"

        If UpdateXSDC2(p_Block_b, strWhere) = FUNCTION_RETURN_FAILURE Then
            SetPattern1 = FUNCTION_RETURN_FAILURE
            p_Error = GetMsgStr("EAPLY") & sDbName
            GoTo proc_exit
        End If
        
        '' CC705(Bar�o��)���э쐬�@04/09/27 ooba START ================================>
        If strNowCd = "CC700" And strNxtCd = "CC705" Then
        
            p_Block_b_bar = p_Block_b
            
            With p_Block_b_bar
                .KCNTC2 = CInt(p_Block_b.KCNTC2) + 1        '�H���A��
                .NEKKNTC2 = p_Block_b.GNKKNTC2              '�ŏI�ʉߊǗ��H��
                .NEWKNTC2 = p_Block_b.GNWKNTC2              '�ŏI�ʉߍH��
                .NEWKKBC2 = p_Block_b.GNWKKBC2              '�ŏI�ʉߍ�Ƌ敪
                .NEMACOC2 = p_Block_b.GNMACOC2              '�ŏI�ʉߏ�����
                .GNWKNTC2 = Space(5)                        '���ݍH��(��߰�)
                .GNMACOC2 = 1                               '���ݏ�����
            End With
            
            If UpdateXSDC2(p_Block_b_bar, strWhere) = FUNCTION_RETURN_FAILURE Then
                SetPattern1 = FUNCTION_RETURN_FAILURE
                p_Error = GetMsgStr("EAPLY") & sDbName
                GoTo proc_exit
            End If
        End If
        '' CC705(Bar�o��)���э쐬�@04/09/27 ooba END ==================================>
        
            
    '�ᕪ������(�i��)-XSDCA��
        sDbName = "(XSDCA)"
        
        recCnt = UBound(p_Hinban_b)   '��������(�i��)-b
        ReDim typKotei(recCnt)
        ReDim XSDCA_c_flg(recCnt)
        ReDim p_Hinban_b_bar(recCnt)  'Bar�o�חp��������(�i��)�o�^���@04/09/27 ooba
        ReDim typKotei_bar(recCnt)    'Bar�o�חp�H�����ѓo�^���@04/09/27 ooba
       
        recCnt3 = UBound(p_Furyo)     '�s�ǎ���-d
        
        '�����A�d�ʃZ�b�g
        For i = 1 To recCnt
            For j = 1 To recCnt3
                If p_Hinban_b(i).CRYNUMCA = p_Furyo(j).XTALC4 _
                            And p_Hinban_b(i).HINBCA = p_Furyo(j).HINBC4 _
                            And p_Hinban_b(i).INPOSCA = p_Furyo(j).INPOSC4 Then
                            
                    '�s�ǒ������}�C�i�X
'                    p_Hinban_b(i).GNLCA = p_Hinban_b(i).GNLCA - p_Furyo(j).PUCUTLC4
''                   p_Hinban_b(i).GNWCA = WeightOfCylinder(dblDiameter, CInt(p_Hinban_b(i).GNLCA))
'                    p_Hinban_b(i).GNWCA = p_Hinban_b(i).GNWCA - p_Furyo(j).PUCUTWC4
                    
                    '' �����^�d�ʓo�^�ύX�@04/09/30 ooba START ================================>
                    If PutWtFlg = 1 Then
                        p_Hinban_b(i).SUMITLCA = p_Hinban_b(i).SUMITLCA - p_Furyo(j).PUCUTLC4
                        p_Hinban_b(i).SUMITWCA = p_Hinban_b(i).SUMITWCA - p_Furyo(j).PUCUTWC4
                    Else
                        p_Hinban_b(i).GNLCA = p_Hinban_b(i).GNLCA - p_Furyo(j).PUCUTLC4
                        p_Hinban_b(i).GNWCA = p_Hinban_b(i).GNWCA - p_Furyo(j).PUCUTWC4
                    End If
                    '' �����^�d�ʓo�^�ύX�@04/09/30 ooba END ==================================>
                    
                     '�s�ǒ������
                    XSDCA_c_flg(i).Furyo = p_Furyo(j).PUCUTLC4
                    XSDCA_c_flg(i).FuryoW = p_Furyo(j).PUCUTWC4
                    'index���
                    XSDCA_c_flg(i).Index_F = j
                    Exit For
                Else
                    XSDCA_c_flg(i).Index_F = -1
                End If
            Next
        Next
        
        n = 0
        For i = 1 To recCnt
            '��ۯ��̍H���A�Ԃ��Z�b�g
            p_Hinban_b(i).KCKNTCA = p_Block_b.KCNTC2
            
            '�s�Ǔ���ں��ނɕi�ԏ����Z�b�g
            If XSDCA_c_flg(i).Index_F > 0 Then
                With p_Furyo(XSDCA_c_flg(i).Index_F)
                    .KCKNTC4 = p_Hinban_b(i).KCKNTCA      '�H���A��
                    .REVNUMC4 = p_Hinban_b(i).REVNUMCA    '���i�ԍ������ԍ�
                    .FACTORYC4 = p_Hinban_b(i).FACTORYCA  '�H��
                    .OPEC4 = p_Hinban_b(i).OPECA          '���Ə���
                    .WKKTC4 = strNowCd                    '�H��
                End With
            End If
            
            '�X�V
            strWhere = "WHERE CRYNUMCA = '" & p_Hinban_b(i).CRYNUMCA
            strWhere = strWhere & "' AND HINBCA = '" & p_Hinban_b(i).HINBCA
            strWhere = strWhere & "' AND INPOSCA = " & p_Hinban_b(i).INPOSCA
            
            '�ؽ�ٶ�۸ފi��̏ꍇ
            If strNowCd = "CB320" And strChgHin <> "" Then
                With p_Hinban_b(i)
                    .LIVKCA = "1"          '�����׸�1���
                End With
            '�����ŏI���o�̏ꍇ(�ŏI��ԋ敪���)
            Else
                If strNowCd = "CC700" Then
                    With p_Hinban_b(i)
                        If strNxtCd = "CC705" Then
                            .LSTATBCA = "B"   '�ŏI��ԋ敪(BAR�o��)
                            .KANKCA = "2"     '�����敪(�I��)
                            .LIVKCA = "1"     '�����׸�1���
                             typKotei(n).PAYCLASSC3 = "1"
                        Else
                            .LSTATBCA = "W"   '�ŏI��ԋ敪(WF�o��)
                             typKotei(n).PAYCLASSC3 = "0"
                        End If
                    End With
                End If
                
                With p_Hinban_b(i)
                    .GNWKNTCA = strNxtCd         ' ���ݍH��
                    .NEWKNTCA = strNowCd         ' �ŏI�ʉߍH��
                End With
                
                '�����񐔎擾���W�b�N�ύX 2002/11/21 tuku  START
                intSyoriKaisu = GetGNMACOC(p_Block_b.CRYNUMC2, strNxtCd) '��������(��ۯ�)�O�H��������ۯ�ID���L�[�Ɏ擾
                If strNxtCd = strNowCd Then                             ' ���H���d�|��Ή��@2002/11/25 tuku
                    intSyoriKaisu = intSyoriKaisu + 1
                End If
                p_Hinban_b(i).GNMACOCA = intSyoriKaisu
                p_Hinban_b(i).NEMACOCA = GetNEMACOC(p_Hinban_b(i).CRYNUMCA, CInt(p_Hinban_b(i).INPOSCA))
                '�����񐔎擾���W�b�N�ύX 2002/11/21 tuku  END
            End If
        
            If UpdateXSDCA(p_Hinban_b(i), strWhere) = FUNCTION_RETURN_FAILURE Then
                SetPattern1 = FUNCTION_RETURN_FAILURE
                p_Error = GetMsgStr("EAPLY") & sDbName
                GoTo proc_exit
            End If
            
            '' CC705(Bar�o��)���э쐬�@04/09/27 ooba START ================================>
            If strNowCd = "CC700" And strNxtCd = "CC705" Then
            
                p_Hinban_b_bar(i) = p_Hinban_b(i)
                
                With p_Hinban_b_bar(i)
                    .KCKNTCA = CInt(p_Hinban_b(i).KCKNTCA) + 1  '�H���A��
                    .NEKKNTCA = p_Hinban_b(i).GNKKNTCA          '�ŏI�ʉߊǗ��H��
                    .NEWKNTCA = p_Hinban_b(i).GNWKNTCA          '�ŏI�ʉߍH��
                    .NEWKKBCA = p_Hinban_b(i).GNWKKBCA          '�ŏI�ʉߍ�Ƌ敪
                    .NEMACOCA = p_Hinban_b(i).GNMACOCA          '�ŏI�ʉߏ�����
                    .GNWKNTCA = Space(5)                        '���ݍH��(��߰�)
                    .GNMACOCA = 1                               '���ݏ�����
                End With
                
                If UpdateXSDCA(p_Hinban_b_bar(i), strWhere) = FUNCTION_RETURN_FAILURE Then
                    SetPattern1 = FUNCTION_RETURN_FAILURE
                    p_Error = GetMsgStr("EAPLY") & sDbName
                    GoTo proc_exit
                End If
            End If
            '' CC705(Bar�o��)���э쐬�@04/09/27 ooba END ==================================>
        
        
'2002/08/29----------------------------------------------------
            '�ؽ�ٶ�۸ފi��̏ꍇ �i�ԕύX
            If strNowCd = "CB320" And strChgHin <> "" Then
                With p_Hinban_b(i)
                    strMotHin = .HINBCA & .REVNUMCA & .FACTORYCA & .OPECA
'''                    .HINBCA = strChgHin    'changeHinban���Z�b�g

                    ''�ŐV�i�Ԃ��擾����悤�ɕύX�@2003/11/10 ooba�@START
                    If GetLastHinban(strChgHin, fullHinban) = FUNCTION_RETURN_FAILURE Then
                        SetPattern1 = FUNCTION_RETURN_FAILURE
                        p_Error = GetMsgStr(EHIN0)
                        GoTo proc_exit
                    End If
                    .HINBCA = fullHinban.hinban
                    .REVNUMCA = fullHinban.mnorevno
                    .FACTORYCA = fullHinban.factory
                    .OPECA = fullHinban.opecond
                    ''�ŐV�i�Ԃ��擾����悤�ɕύX�@2003/11/10 ooba�@END
                    
                    .LIVKCA = "0"          '�����׸�0���
'2002/10/16 �ǉ�-------------------------------------------------------------------��3-�B
                    .GNWKNTCA = strNxtCd         ' ���ݍH��
                    .NEWKNTCA = strNowCd         ' �ŏI�ʉߍH��
'2002/10/16 �ǉ�-------------------------------------------------------------------��3-�B
                    .RSTATBCA = "T" '�i�グ���ɗ�����ԋ敪��ʏ�ɖ߂� 2002/12/17 tuku

                    '�����񐔎擾���W�b�N�ύX 2002/11/21 tuku  START
                    p_Hinban_b(i).GNMACOCA = GetGNMACOC(p_Block_b.CRYNUMC2, strNxtCd) '��������(��ۯ�)�O�H��������ۯ�ID���L�[�Ɏ擾
                    p_Hinban_b(i).NEMACOCA = GetNEMACOC(p_Hinban_b(i).CRYNUMCA, CInt(p_Hinban_b(i).INPOSCA))
                    '�����񐔎擾���W�b�N�ύX 2002/11/21 tuku  END
                    
                    If CheckUniqueRecord(.CRYNUMCA, .HINBCA, CInt(.INPOSCA)) = True Then
                        '�o�^
                        If CreateXSDCA(p_Hinban_b(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                            SetPattern1 = FUNCTION_RETURN_FAILURE
                            p_Error = strErrMsg
                            GoTo proc_exit
                        End If
                    Else
                        '�X�V
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

            
    '��H�����ѓo�^-XSDC3��
            Call SetXSDC3(typKotei(n), p_Hinban_b(i), XSDCA_c_flg(i))
            If strNowCd = "CC400" Then
'               typKotei(n).FRLC3 = p_Hinban_b(i).GNLCA
'               typKotei(n).FRWC3 = p_Hinban_b(i).GNWCA
'               typKotei(n).FUWC3 = XSDCA_c_flg(i).FuryoW
                '' �s�Ǐd�ʓo�^�ύX�@04/09/30 ooba START =================================>
                typKotei(n).LOSLC3 = XSDCA_c_flg(i).FuryoW          '���X�d��
                If PutWtFlg <> 1 Then
                    typKotei(n).FUWC3 = XSDCA_c_flg(i).FuryoW       '�s�Ǐd��
                End If
                '' �s�Ǐd�ʓo�^�ύX�@04/09/30 ooba END ===================================>
            End If
            If strNowCd = "CB320" And strChgHin <> "" Then
                typKotei(n).MOTHINC3 = strMotHin   '���i�Ԃ��Z�b�g
            End If
            If CreateXSDC3(typKotei(n), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                SetPattern1 = FUNCTION_RETURN_FAILURE
                p_Error = strErrMsg
                GoTo proc_exit
            End If
            
            '' CC705(Bar�o��)���э쐬�@04/09/27 ooba START ================================>
            If strNowCd = "CC700" And strNxtCd = "CC705" Then
            
                typKotei_bar(n) = typKotei(n)
                
                With typKotei_bar(n)
                    .KCNTC3 = p_Hinban_b_bar(i).KCKNTCA         '�H���A��
                    .KNKTC3 = p_Hinban_b_bar(i).NEKKNTCA        '�Ǘ��H��
                    .WKKTC3 = p_Hinban_b_bar(i).NEWKNTCA        '�H��
                    .WKKBC3 = p_Hinban_b_bar(i).NEWKKBCA        '��Ƌ敪
                    .MACOC3 = p_Hinban_b_bar(i).NEMACOCA        '������
                    .FRKNKTC3 = p_Hinban_b(i).NEKKNTCA          '(���)�Ǘ��H��
                    .FRWKKTC3 = p_Hinban_b(i).NEWKNTCA          '(���)�H��
                    .FRWKKBC3 = p_Hinban_b(i).NEWKKBCA          '(���)��Ƌ敪
                    .FRMACOC3 = p_Hinban_b(i).NEMACOCA          '(���)������
                    .TOWNKTC3 = p_Hinban_b_bar(i).GNKKNTCA      '(���o)�Ǘ��H��
                    .TOWKKTC3 = p_Hinban_b_bar(i).GNWKNTCA      '(���o)�H��
                    .TOMACOC3 = p_Hinban_b_bar(i).GNMACOCA      '(���o)������
                    .FRLC3 = p_Hinban_b(i).GNLCA                '�������
                    .FRWC3 = p_Hinban_b(i).GNWCA                '����d��
                    .FULC3 = 0                                  '�s�ǒ���
                    .FUWC3 = 0                                  '�s�Ǐd��
                    .TOLC3 = p_Hinban_b_bar(i).GNLCA            '���o����
                    .TOWC3 = p_Hinban_b_bar(i).GNWCA            '���o�d��
                End With
                
                If CreateXSDC3(typKotei_bar(n), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                    SetPattern1 = FUNCTION_RETURN_FAILURE
                    p_Error = strErrMsg
                    GoTo proc_exit
                End If
            End If
            '' CC705(Bar�o��)���э쐬�@04/09/27 ooba END ==================================>
            
            n = n + 1
        Next
        
        
    '��s�ǎ���-XSDC4��
        For i = 1 To recCnt3
            If p_Furyo(i).KCKNTC4 = "" Then
            '��v����i�Ԃ���������ں��ނ̏����Z�b�g
                With p_Furyo(i)
                    .KCKNTC4 = p_Block_b.KCNTC2        '�H���A��
                    .WKKTC4 = strNowCd                 '�H��
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
        
        
    '�ᕪ������(SXL)�o�^-XSDCB��      ����������(�i��)��SXL�P�ʂɏW��
        If SXLflg = True Then
            sDbName = "(XSDCB)"
            Call MakeSXLinfo(p_Hinban_b, typSXL)  '��������(SXL)���쐬
            recCnt = UBound(typSXL)

            For i = 1 To recCnt
                If typSXL(i).SXLIDCB <> "" Then
                    If typSXL(i).LSTCCB = "W" Then
                       typSXL(i).LSTCCB = "T"   '�ŏI��ԋ敪(�ʏ�)���
                    End If
                    'Sumit�A�g�����ύX�ɂ��c�a�ǉ��@�Z�b�g���ݍH���A�ŏI�ʉߍH��
                    If strNowCd = "CC700" Then
                        typSXL(i).NEWKNTCB = "CC700"
                        typSXL(i).GNWKNTCB = "CW750"
                    End If
'�ύX SystamBrain 2003/10/09 ---------------------------------------------------> START
''''                    If strNowCd = "CC710" Then
''''                        typSXL(i).NEWKNTCB = "CC710"
''''                        typSXL(i).GNWKNTCB = "CST02"
''''                    End If
                    If strNowCd = "CC720" Then
                        typSXL(i).NEWKNTCB = "CC720"
                        typSXL(i).GNWKNTCB = "CST02"
                    End If
'�ύX SystamBrain 2003/10/09 ---------------------------------------------------> END
                    '�c�a�ǉ��@�@�@�@�@�@�@�@�_�@�O�Y�@�@�����P�T�N�T���P��
                    If CheckSXLrecord(typSXL(i).SXLIDCB, intLen) = 0 Then
                    'ں��ނ��Ȃ��ꍇ(�o�^)
                    
                        If CreateXSDCB(typSXL(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                            SetPattern1 = FUNCTION_RETURN_FAILURE
                            p_Error = strErrMsg
                            GoTo proc_exit
                        End If
                        
                    Else
                    'ں��ނ�����ꍇ(�X�V)
                    
'                        '�����w�����́A�����ŏI���o���͂̏ꍇ�A�������v���X
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
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    SetPattern1 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v     :��������(SXL)�����Z�b�g����
'���Ұ�   :�ϐ���           ,IO  ,�^                   ,����
'         :p_Hinban_sxl()   ,I   ,typ_XSDCA_Update     ,��������(�i��)���
'         :recSXL()         ,O   ,typ_XSDCB_Update     ,��������(SXL)���
'����     :��������(�i��)��SXL�P�ʂɏW�񂷂�
Private Sub MakeSXLinfo(p_Hinban_sxl() As typ_XSDCA_Update, recSXL() As typ_XSDCB_Update)
    
    '�G���[�n���h���̐ݒ�
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
    Call SetXSDCB(p_Hinban_sxl(1), recSXL(1))  '1ں��ޖڂ��Z�b�g
    
    n = 1
    For i = 2 To recCnt
        For j = 1 To recCnt
            If p_Hinban_sxl(i).SXLIDCA = recSXL(j).SXLIDCB Then
            
            '��čς݂�ں��ނ�SXLID���ꏏ�Ȃ璷�������Z����
                intLength = CInt(p_Hinban_sxl(i).GNLCA) + CInt(recSXL(j).LENCB)
                
                If CInt(p_Hinban_sxl(i).INPOSCA) < CInt(recSXL(j).INPOSCB) Then
                '�J�n�ʒu�̏����������Z�b�g
                    Call SetXSDCB(p_Hinban_sxl(i), recSXL(j))
                    recSXL(j).LENCB = intLength
                Else
                    recSXL(j).LENCB = intLength
                End If
                setSXLflg = True
                Exit For
            End If
        Next
        '��v���Ȃ�ں��ނ��Z�b�g
        If setSXLflg = False Then
            n = n + 1
            
            Call SetXSDCB(p_Hinban_sxl(i), recSXL(n))
        Else
            setSXLflg = False
        End If
    Next
    
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit

    
End Sub


'�T�v     :��������(SXL)�����Z�b�g����
'���Ұ�   :�ϐ���           ,IO  ,�^                   ,����
'         :p_Hinban_sxl     ,I   ,typ_XSDCA_Update     ,��������(�i��)���
'         :recSXL_rtrn      ,O   ,typ_XSDCB_Update     ,��������(SXL)���
'����     :��������(�i��)ð��ٍ��ڂɒl��Ă���

Private Sub SetXSDCB(p_Hinban_sxl As typ_XSDCA_Update, recSXL_rtrn As typ_XSDCB_Update)
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcXSDC_Basic.bas -- Function SetXSDCB"
                        
    
    Dim i As Long
    
    
    '��������(SXL)
    With recSXL_rtrn
        .SXLIDCB = p_Hinban_sxl.SXLIDCA          ' SXLID
        '.KCNTCB = p_Hinban_sxl.KCKNTCA           ' �H���A��(��ۯ����󂯌p�����ꍇ)
        .KCNTCB = GetKCNTCB(.SXLIDCB)            ' �H���A��(SXL����U��ꍇ)
        .XTALCB = p_Hinban_sxl.XTALCA            ' �����ԍ�
        .INPOSCB = p_Hinban_sxl.INPOSCA          ' �������J�n�ʒu
        .LENCB = p_Hinban_sxl.GNLCA              ' ����
        .HINBCB = p_Hinban_sxl.HINBCA            ' �i��
        .REVNUMCB = p_Hinban_sxl.REVNUMCA        ' �d�b�ԍ������ԍ�
        .FACTORYCB = p_Hinban_sxl.FACTORYCA      ' �H��
        .OPECB = p_Hinban_sxl.OPECA              ' ���Ə���
        .MAICB = p_Hinban_sxl.GNMCA              ' ������
        .MOTHINCB = p_Hinban_sxl.HINBCA          ' ���ݕi�ԁ@�@�@�ǉ��@�_�@�����P�T�N�T���P��
'        .WSRMAICB =                             ' WS��㖇��
'        .WSNMAICB =                             ' WS��򌇗�����
'        .WFCMAICB =                             ' �������
'        .SXLRMAICB =                            ' SXL�w��(�Ǖi)
'        .SXLNMAICB =                            ' SXL�w��(�s��)
'        .WFCNMAICB =                            ' WFC����������
'        .SXLEMAICB =                            ' SXL�m�薇��
'        .SRMAICB =                              ' �T���v�����w��(�Ǖi)
'        .SNMAICB =                              ' �T���v�����w��(�s��)
'        .STMAICB =                              ' �T���v������
'        .FURIMAICB =                            ' �U�֖���
'        .XTWORKCB =                             ' �����H��
'        .WFWORKCB =                             ' �E�F�[�n����
'        .FURYCCB =                              ' �s�Ǘ��R
        .LSTCCB = p_Hinban_sxl.LSTATBCA          ' �ŏI��ԋ敪
'        .LUFRCCB =                              ' �i��R�[�h
'        .LUFRBCB =                              ' �i��敪
'        .LDERCCB =                              ' �i���R�[�h
        .LDFRBCB = p_Hinban_sxl.LDFRBCA          ' �i���敪
'        .HOLDCCB =                              ' �z�[���h�R�[�h
        .HOLDBCB = p_Hinban_sxl.HOLDBCA          ' �z�[���h�敪
'        .EXKUBCB =                              ' ��O�敪
'        .HENPKCB =                              ' �ԕi�敪
        .LIVKCB = p_Hinban_sxl.LIVKCA            ' �����敪
        .KANKCB = p_Hinban_sxl.KANKCA            ' �����敪
'        .NFCB =                                 ' ���ɋ敪
'        .SAKJCB =                               ' �폜�敪
'        .TDAYCB =                               ' �o�^���t
'        .KDAYCB =                               ' �X�V���t
'        .SUMITCB =                              ' SUMIT���M�t���O
'        .SNDKCB =                               ' �ԕi�敪
'        .SNDAYCB =                              ' ���M���t
'        .
    
    End With
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit

    
End Sub

'�T�v     :�H�����я����Z�b�g����
'���Ұ�   :�ϐ���           ,IO  ,�^                   ,����
'         :p_Koteij         ,O   ,typ_XSDC3_Update     ,�H�����ѓo�^���
'         :p_Hinban         ,I   ,typ_XSDCA_Update     ,��������(�i��)���
'         :hinbanflg        ,I   ,typ_XSDCA_c_flg      ,��������(�i��)���
'����     :
Private Sub SetXSDC3(p_Koteij As typ_XSDC3_Update, p_Hinban As typ_XSDCA_Update, _
                                  hinbanflg As typ_XSDCA_c_flg)

    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcXSDC_Basic.bas -- Function SetXSDC3"
                        
    
    Dim i As Long
    
    
    '�H�����ѓo�^
    With p_Koteij
        .CRYNUMC3 = p_Hinban.CRYNUMCA                                     '��ۯ�ID
        .INPOSC3 = p_Hinban.INPOSCA                                       '�������J�n�ʒu
        .KCNTC3 = p_Hinban.KCKNTCA                                        '�H���A��
        .HINBC3 = p_Hinban.HINBCA                                         '�i��
        .REVNUMC3 = p_Hinban.REVNUMCA                                     '���i�ԍ������ԍ�
        .FACTORYC3 = p_Hinban.FACTORYCA                                   '�H��
        .OPEC3 = p_Hinban.OPECA                                           '���Ə���
        .LENC3 = p_Hinban.GNLCA                                           '����
        .XTALC3 = p_Hinban.XTALCA                                         '�����ԍ�
        .SXLIDC3 = p_Hinban.SXLIDCA                                       'SXLID
        '.WKKTC3 = p_Hinban.GNWKNTCA                                       '�H��
        .WKKTC3 = p_Hinban.NEWKNTCA      '�H��(�ŏI�ʉߍH�����Z�b�g)�O�H��                  '2002/08/29
        '.MACOC3 = p_Hinban.GNMACOCA                                       '������
        .MACOC3 = p_Hinban.NEMACOCA      '������(�ŏI�ʉߏ�����)�O�H��                  '2002/08/29
        .FRWKKTC3 = msL2Wkkt             '(���)�H��(�O�H���̍ŏI�ʉߍH��)�O�O�H����           '2002/08/30
        .FRMACOC3 = msL2Maco             '(���)������(�O�H���̍ŏI�ʉߏ�����)�O�O�H����   '2002/08/30
        .TOWKKTC3 = p_Hinban.GNWKNTCA    '(���o)�H��(���ݍH��)���H��                       '2002/08/29
        .TOMACOC3 = p_Hinban.GNMACOCA    '(���o)������(���ݏ�����)���H��               '2002/08/29
        If hinbanflg.Furyo <> 0 Then
            .FULC3 = hinbanflg.Furyo                                      '�s�ǒ���
            .FUWC3 = hinbanflg.FuryoW                                     '�s�Ǐd��
'           .FUWC3 = WeightOfCylinder(dblDiameter, CInt(hinbanflg.Furyo)) '�s�Ǐd��
            '.FUMC3 =                                                     '�s�ǖ���
        End If
        .TOLC3 = p_Hinban.GNLCA                                           '���o����
        .TOWC3 = p_Hinban.GNWCA                                           '���o�d��
'       .TOWC3 = WeightOfCylinder(dblDiameter, CInt(p_Hinban.GNLCA))      '���o�d��
        '.TOMC3 =                                                         '���o����
        '' SUMIT�����^�d�ʓo�^�ǉ��@04/09/30 ooba
        If PutWtFlg > 0 Then
            .SUMITLC3 = p_Hinban.SUMITLCA                                 'SUMIT����
            .SUMITWC3 = p_Hinban.SUMITWCA                                 'SUMIT�d��
        End If
        '2003.06.11 (SPK)Y.katabami�@��\�i�ԁ��V�K�Đ؋敪���ǉ�
        .CUTCNTC3 = p_Hinban.CUTCNTCA             '�V�K�Đ؋敪
        .HINBFLGC3 = p_Hinban.HINBFLGCA           '��\�i�ԃt���O
        ''TEST 2005/11
        .RPCRYNUMC3 = p_Hinban.RPCRYNUMCA
    End With
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit

    
End Sub

'�T�v     :�Z�b�g�p�^�[���U�̏������s��(�����b�g����)
'���Ұ�   :�ϐ���           ,IO  ,�^                   ,����
'         :p_Block_b        ,I   ,typ_XSDC2_Update     ,��������(��ۯ�)�O�H�����я��
'         :p_Hinban_b()     ,I   ,typ_XSDCA_Update     ,��������(�i��)�O�H�����я��
'         :p_Block_c        ,I   ,typ_XSDC2_Update     ,��������(��ۯ�)�o�^���
'         :p_Hinban_c()     ,I   ,typ_XSDCA_Update     ,��������(�i��)�o�^���
'         :p_Furyo()        ,I   ,typ_XSDC4_Update     ,�s�Ǔ���o�^���
'         :p_Error          ,O   ,String               ,�װү����
'         :�߂�l           ,O    ,FUNCTION_RETURN      ,�V�c�a�ւ̏����݂̐���
'����     :
Private Function SetPattern2(p_Block_b As typ_XSDC2_Update, p_Hinban_b() As typ_XSDCA_Update, _
                    p_Block_c As typ_XSDC2_Update, p_Hinban_c() As typ_XSDCA_Update, _
                    p_Furyo() As typ_XSDC4_Update, p_Error As String) As FUNCTION_RETURN
                        
                        
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcXSDC_Basic.bas -- Function SetPattern2"
                        
    Dim recCnt As Long    '���R�[�h��
    Dim recCnt2 As Long   '���R�[�h��
    Dim recCnt3 As Long   '���R�[�h��
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim sDbName As String
    Dim strWhere As String
    Dim strErrMsg As String
    Dim SetRec2 As typ_XSDC2_Update
    Dim SetRecA() As typ_XSDCA_Update
    Dim SetRecCB320 As typ_XSDCA_Update 'CB320 G�i��ں��ލ쐬�p
    Dim XSDCA_c_flg() As typ_XSDCA_c_flg
    Dim typKotei() As typ_XSDC3_Update
    Dim HINFLG As typ_XSDCA_c_flg
    Dim strMotHin As String
    Dim AccFlg As Boolean         '�s�Ǔo�^�׸�(�s�Ǔ���ɓo�^������������)
    Dim typSXL() As typ_XSDCB_Update
    Dim intLen As Integer         'SXL����


'�Ώ�ں��ނ��Z�b�g
    If regFLG = "Y" Then
'        recCnt = UBound(p_Hinban_b)
        recCnt = UBound(p_Hinban_c)
'        ReDim p_Hinban_c(recCnt2)
        ReDim typKotei(recCnt)
        ReDim XSDCA_c_flg(recCnt)

        
    'c ��ں��ނ��
        SetRec2 = p_Block_c             '��ۯ�ں���
        
        ReDim SetRecA(recCnt)
        For i = 1 To recCnt
            SetRecA(i) = p_Hinban_c(i)  '�i��ں���
        Next
    Else
        recCnt = UBound(p_Hinban_b)
        ReDim typKotei(recCnt)
        ReDim XSDCA_c_flg(0)
        
    'b ��ں��ނ��
        SetRec2 = p_Block_b             '��ۯ�ں���

        ReDim SetRecA(recCnt)
        For i = 1 To recCnt
            SetRecA(i) = p_Hinban_b(i)  '�i��ں���
        Next
    End If
    
    '�s�ǐ��J�E���g
    recCnt3 = UBound(p_Furyo)



'�EnextCode �� CB210, CB320, '     ' �ȊO�̏ꍇ
    If strNxtCd <> "CB210" And strNxtCd <> "CB320" And strNxtCd <> "     " Then
    
        strNxtCd = "CB210"
        
'        '���ݍH���ύX
'        SetRec2.GNWKNTC2 = strNxtCd
'        For i = 1 To recCnt
'            SetRecA(i).GNWKNTCA = strNxtCd
'        Next
    End If
    

    
'�E�敪�Z�b�g
    Select Case strNxtCd
        Case "CB210"
            '��������(��ۯ�)-XSDC2
            With SetRec2
                .LSTATBC2 = "R"   '�ŏI��ԋ敪(����)
                .LDFRBC2 = "1"    '�i���敪(����)
                .RSTATBC2 = "M"   '������ԋ敪(���Ď���҂�)
            End With
            '��������(�i��)-XSDCA
            For i = 1 To recCnt
                With SetRecA(i)
                    .LSTATBCA = "R"   '�ŏI��ԋ敪(����)
                    .LDFRBCA = "1"    '�i���敪(����)
                    .RSTATBCA = "M"   '������ԋ敪(���Ď���҂�)
                End With
            Next
        Case "CB320"
            '��������(��ۯ�)-XSDC2
            SetRec2.RSTATBC2 = "G"   '������ԋ敪(�ؽ�ٶ�۸�)
            
            '��������(�i��)-XSDCA
            For i = 1 To recCnt
                SetRecA(i).RSTATBCA = "G"   '������ԋ敪(�ؽ�ٶ�۸�)
            Next
        Case "     "
            '��������(��ۯ�)-XSDC2
            With SetRec2
                .LSTATBC2 = "H"   '�ŏI��ԋ敪(�p��)
                .LDFRBC2 = "2"    '�i���敪(ʲ�)
            End With
            
            '��������(�i��)-XSDCA
            For i = 1 To recCnt
                With SetRecA(i)
                    .LSTATBCA = "H"   '�ŏI��ԋ敪(�p��)
                    .LDFRBCA = "2"    '�i���敪(ʲ�)
                End With
            Next
    End Select
           
    '�����b�g���������H����ύX����悤�ɕύX�@2002/12/17 tuku  START
    '���ݍH���ύX
    SetRec2.GNWKNTC2 = strNxtCd
    For i = 1 To recCnt
        SetRecA(i).GNWKNTCA = strNxtCd
    Next
    '�ŏI�ʉߍH���ύX
    SetRec2.NEWKNTC2 = strNowCd
    For i = 1 To recCnt
        SetRecA(i).NEWKNTCA = strNowCd
    Next
    '�����b�g���������H����ύX����悤�ɕύX�@2002/12/17 tuku END
        
    If CC300Flg = False Then   '���H��CC300����ۯ��̓o�^���s��Ȃ�
    
'�ᕪ������(��ۯ�)-XSDC2��
        sDbName = "(XSDC2)"
        With SetRec2
    '        If regFLG = "Y" Then
    '            .GNLC2 = p_Block_c.GNLC2   '�o�^����(c)
    '        Else
    '            .GNLC2 = p_Block_b.GNLC2   '�O�H������(b)
    '        End If
            
            '�����敪�Z�b�g
            If strNxtCd = "CB320" Then
                .LIVKC2 = "0"
'2002/10/16 �ǉ�-------------------------------------------------------------------��
                .GNWKNTC2 = strNxtCd         ' ���ݍH��
                .NEWKNTC2 = strNowCd         ' �ŏI�ʉߍH��
'2002/10/16 �ǉ�-------------------------------------------------------------------��
            Else
                .LIVKC2 = "1"
            End If
        End With
            
        If regFLG = "Y" Then
        ''���Z�b�g�p�^�[���U-�@ (�o�^���(c))
            If p_Block_b.CRYNUMC2 = p_Block_c.CRYNUMC2 Then   'b=c
            '�X�V(�o�^���(c))
                With SetRec2
                    .KCNTC2 = p_Block_b.KCNTC2 + 1    '�H���A�Ԃ��{�P���ăZ�b�g
                    .GNMACOC2 = p_Block_b.GNMACOC2    '���ݏ�����
                    .NEWKNTC2 = p_Block_b.NEWKNTC2    '�ŏI�ʉߍH��
                    .NEMACOC2 = p_Block_b.NEMACOC2    '�ŏI�ʉߏ�����
                End With
                strWhere = "WHERE CRYNUMC2 = '" & p_Block_b.CRYNUMC2 & "'"
                If UpdateXSDC2(SetRec2, strWhere) = FUNCTION_RETURN_FAILURE Then
                    SetPattern2 = FUNCTION_RETURN_FAILURE
                    p_Error = GetMsgStr("EAPLY") & sDbName
                    GoTo proc_exit
                End If
            Else
            '�o�^(�o�^���(c))
                SetRec2.KCNTC2 = 1    '�H���A�ԂɂP���Z�b�g
                If CreateXSDC2(SetRec2, strErrMsg) = FUNCTION_RETURN_FAILURE Then
                    SetPattern2 = FUNCTION_RETURN_FAILURE
                    p_Error = strErrMsg
                    GoTo proc_exit
                End If
            End If
        Else
        
        ''���Z�b�g�p�^�[���U-�A (�O�H�����(b))
            '�X�V(�O�H�����(b))
            If Left(p_Block_b.CRYNUMC2, 1) <> vbNullChar Then '(��������ꍇ)
                SetRec2.KCNTC2 = p_Block_b.KCNTC2 + 1    '�H���A�Ԃ��{�P���ăZ�b�g
                strWhere = "WHERE CRYNUMC2 = '" & p_Block_b.CRYNUMC2 & "'"
                If UpdateXSDC2(SetRec2, strWhere) = FUNCTION_RETURN_FAILURE Then
                    SetPattern2 = FUNCTION_RETURN_FAILURE
                    p_Error = GetMsgStr("EAPLY") & sDbName
                    GoTo proc_exit
                End If
            End If
        End If
    
    End If
    
    
'�ᕪ������(�i��)-XSDCA��
    sDbName = "(XSDCA)"
    For i = 1 To recCnt
        With SetRecA(i)
'            If regFLG = "Y" Then
'                .GNLCA = p_Hinban_c(i).GNLCA   '�o�^����(c)
'            Else
'                .GNLCA = p_Hinban_b(i).GNLCA   '�O�H������(b)
'            End If

            '�����敪�Z�b�g
            .LIVKCA = "1"
        End With
    Next
    
    n = 0

    If regFLG = "Y" Then
    ''���Z�b�g�p�^�[���U-�@ (�o�^���(c))
    
        recCnt = UBound(p_Hinban_b)
        recCnt2 = UBound(p_Hinban_c)
        
'        '�s�ǐ�������΃Z�b�g����
'        For i = 1 To recCnt2
'            For j = 1 To reccnt3
'                If p_Hinban_c(i).CRYNUMCA = p_Furyo(j).XTALC4 _
'                            And p_Hinban_c(i).HINBCA = p_Furyo(j).HINBC4 _
'                            And p_Hinban_c(i).INPOSCA = p_Furyo(j).INPOSC4 Then
'
'                    XSDCA_c_flg(i).Furyo = p_Furyo(j).PUCUTLC4   '�s�ǒ������
'                    XSDCA_c_flg(i).Index_F = j                   'index���
'
'                Else
'                    XSDCA_c_flg(i).Index_F = -1                   'index���
'                End If
'            Next
'        Next
    
        For i = 1 To recCnt
            For j = 1 To recCnt2
                If p_Hinban_b(i).CRYNUMCA = p_Hinban_c(j).CRYNUMCA _
                                    And p_Hinban_b(i).HINBCA = p_Hinban_c(j).HINBCA _
                                    And p_Hinban_b(i).INPOSCA = p_Hinban_c(j).INPOSCA Then  'b=c
                    
                    '�o�^ں��ޏ��(c)��UPDATE
                    With SetRecA(j)
                        .KCKNTCA = SetRec2.KCNTC2      '��ۯ��̍H���A�ԃZ�b�g
                        .GNMACOCA = p_Hinban_b(i).GNMACOCA    '���ݏ�����
                        .NEWKNTCA = p_Hinban_b(i).NEWKNTCA    '�ŏI�ʉߍH��
                        .NEMACOCA = p_Hinban_b(i).NEMACOCA    '�ŏI�ʉߏ�����
                    End With
                    
                    strWhere = "WHERE CRYNUMCA = '" & p_Hinban_b(i).CRYNUMCA
                    strWhere = strWhere & "' AND HINBCA = '" & p_Hinban_b(i).HINBCA
                    strWhere = strWhere & "' AND INPOSCA = " & p_Hinban_b(i).INPOSCA
                    If UpdateXSDCA(SetRecA(j), strWhere) = FUNCTION_RETURN_FAILURE Then
                        SetPattern2 = FUNCTION_RETURN_FAILURE
                        p_Error = GetMsgStr("EAPLY") & sDbName
                        GoTo proc_exit
                    End If
                    
                    '�H������-XSDC3�o�^
                    Call SetXSDC3(typKotei(n), SetRecA(j), XSDCA_c_flg(0))
                    With typKotei(n)
                        .LENC3 = "0"     '����
                        .TOLC3 = .LENC3  '���o����
                        .TOWC3 = "0"     '���o�d��
                        .FULC3 = SetRecA(j).GNLCA      '�s�ǒ���
                        .FUWC3 = SetRecA(j).GNWCA      '�d��
                    End With
                    If CreateXSDC3(typKotei(n), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                        SetPattern2 = FUNCTION_RETURN_FAILURE
                        p_Error = strErrMsg
                        GoTo proc_exit
                    End If
                    n = n + 1
                    XSDCA_c_flg(j).Entry = True   '�X�V�׸�(��v����ں��ރA��)
                    Exit For
                End If
'                XSDCA_c_flg(j).Entry = False   '�X�V�׸�(��v����ں��ރi�V)
            Next
        Next
        
        '��v���Ȃ�ں���(c)�̓o�^
        For i = 1 To recCnt2
             If XSDCA_c_flg(i).Entry = False Then
             
                '�H���A�ԃZ�b�g
                If recCnt <> 0 And p_Hinban_b(1).CRYNUMCA = p_Hinban_c(1).CRYNUMCA Then
                    SetRecA(i).KCKNTCA = SetRec2.KCNTC2   '��ۯ��̍H���A�Ԃ��Z�b�g
                Else
                    SetRecA(i).KCKNTCA = 1                '�H���A�ԂɂP���Z�b�g
                End If
             
             
                If CreateXSDCA(SetRecA(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                    SetPattern2 = FUNCTION_RETURN_FAILURE
                    p_Error = strErrMsg
                    GoTo proc_exit
                End If
                
                    
                '��H������-XSDC3�o�^��
                Call SetXSDC3(typKotei(n), SetRecA(i), XSDCA_c_flg(0))
                With typKotei(n)
                    .LENC3 = "0"                   '����
                    .TOLC3 = .LENC3                '���o����
                    .TOWC3 = "0"                   '���o�d��
                    .FULC3 = SetRecA(i).GNLCA      '�s�ǒ���
                    .FUWC3 = SetRecA(i).GNWCA      '�d��
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
    
    ''���Z�b�g�p�^�[���U-�A (�O�H�����(b))
    
'        '�s�ǐ�������΃Z�b�g����
'        For i = 1 To recCnt
'            For j = 1 To reccnt3
'                If p_Hinban_b(i).CRYNUMCA = p_Furyo(j).XTALC4 _
'                            And p_Hinban_b(i).HINBCA = p_Furyo(j).HINBC4 _
'                            And p_Hinban_b(i).INPOSCA = p_Furyo(j).INPOSC4 Then
'
'                    XSDCA_c_flg(i).Furyo = p_Furyo(j).PUCUTLC4   '�s�ǒ������
'                    XSDCA_c_flg(i).Index_F = j                   'index���
'
'                Else
'                    XSDCA_c_flg(i).Index_F = -1                   'index���
'                End If
'            Next
'        Next
    
        For i = 1 To recCnt
            '�X�V
            SetRecA(i).KCKNTCA = SetRec2.KCNTC2      '��ۯ��̍H���A�ԃZ�b�g
            strWhere = "WHERE CRYNUMCA = '" & p_Hinban_b(i).CRYNUMCA
            strWhere = strWhere & "' AND HINBCA = '" & p_Hinban_b(i).HINBCA
            strWhere = strWhere & "' AND INPOSCA = " & p_Hinban_b(i).INPOSCA
            If UpdateXSDCA(SetRecA(i), strWhere) = FUNCTION_RETURN_FAILURE Then
                SetPattern2 = FUNCTION_RETURN_FAILURE
                p_Error = GetMsgStr("EAPLY") & sDbName
                GoTo proc_exit
            End If
            
            '��H������-XSDC3�o�^��
            Call SetXSDC3(typKotei(n), SetRecA(i), XSDCA_c_flg(0))
            With typKotei(n)
                .LENC3 = "0"                  '����
                .TOLC3 = .LENC3               '���o����
                .TOWC3 = "0"                  '���o�d��
                .FULC3 = SetRecA(i).GNLCA     '�s�ǒ���
                .FUWC3 = SetRecA(i).GNWCA     '�d��
'2002/10/17----------------------------------------------------
                .WKKTC3 = SetRecA(i).NEWKNTCA                               '�H�� �@2002/12/17 tuku
                .MACOC3 = SetRecA(i).NEMACOCA                               '������ 2002/12/17 tuku
                .FRWKKTC3 = msL2Wkkt                                        '(���)�H��
                .FRMACOC3 = msL2Maco                                        '(���)������
                .TOWKKTC3 = strNxtCd                                        '(���o)�H��(���ݍH��)���H��
                If Left(SetRec2.CRYNUMC2, 1) <> vbNullChar Then
                     .TOMACOC3 = GetGNMACOC(SetRec2.CRYNUMC2, strNxtCd)     '(���o)������
                Else
                     .TOMACOC3 = 1                                          '(���o)������
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
    
    
    '���H����CB320�̏ꍇ(�V�Kں��ލ쐬)
    If strNxtCd = "CB320" Then
        SetRecCB320 = SetRecA(1)
        With SetRecCB320
            .GNLCA = SetRec2.GNLC2                                 '����(��ۯ�����)���
            .GNWCA = SetRec2.GNWC2
            .INPOSCA = SetRec2.INPOSC2
            .HINBCA = "G"                                          '�i�Ծ��
            .LIVKCA = "0"                                          '�����敪 0
            .KCKNTCA = SetRec2.KCNTC2                              '�H���A��(��ۯ�)���
            .CHGCA = 0                                             '�`���[�W�ʁ@0
            .KEIDAYCA = ""                                         '�v����t�@���Z�b�g
'2002/10/16 �ǉ�-------------------------------------------------------------------��
            .GNWKNTCA = strNxtCd         ' ���ݍH��
            .NEWKNTCA = strNowCd         ' �ŏI�ʉߍH��
'2002/10/16 �ǉ�-------------------------------------------------------------------��
        End With
        
'2002/08/29----------------------------------------------------
'        If CreateXSDCA(SetRecCB320, strErrMsg) = FUNCTION_RETURN_FAILURE Then
'            SetPattern2 = FUNCTION_RETURN_FAILURE
'            p_Error = strErrMsg
'            GoTo proc_exit
'        End If
                
        With SetRecCB320
            If CheckUniqueRecord(.CRYNUMCA, .HINBCA, CInt(.INPOSCA)) = True Then
                '�o�^
                If CreateXSDCA(SetRecCB320, strErrMsg) = FUNCTION_RETURN_FAILURE Then
                    SetPattern2 = FUNCTION_RETURN_FAILURE
                    p_Error = strErrMsg
                    GoTo proc_exit
                End If
            Else
                '�X�V
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
'                strMotHin = .HINBCA & .REVNUMCA & .FACTORYCA & .OPECA  '���i�Ԏ擾
'                '.HINBCA = "G"                                          '�i�Ծ��
'                .KCKNTCA = SetRec2.KCNTC2                              '�H���A��(��ۯ�)���
'            End With
'
'            '��H������-XSDC3�o�^��
'            Call SetXSDC3(typKotei(n), SetRecA(i), XSDCA_c_flg(0))
'            With typKotei(n)
'                .HINBC3 = "G"                                           '�i�Ծ��
'                .LENC3 = SetRecA(i).GNLCA                               '����(�i�Ԓ���)���
'                .TOLC3 = .LENC3                                         '���o����
'                .TOWC3 = WeightOfCylinder(dblDiameter, CInt(.TOLC3))    '���o�d��
'                '.FULC3 = SetRecA(i).GNLCA                               '�s�ǒ���
'                '.FUWC3 = SetRecA(i).GNWCA                               '�d��
'                .MOTHINC3 = strMotHin                                   '���i��
'2002/10/17----------------------------------------------------
'                .WKKTC3 = SetRecA(i).GNWKNTCA                               '�H��
'                .MACOC3 = SetRecA(i).GNMACOCA                               '������
'                .FRWKKTC3 = msL2Wkkt                                        '(���)�H��
'                .FRMACOC3 = msL2Maco                                        '(���)������
'                .TOWKKTC3 = strNxtCd                                        '(���o)�H��(���ݍH��)���H��
'                .TOMACOC3 = GetGNMACOC(SetRec2.CRYNUMC2, strNxtCd)          '(���o)������
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
        
'    '�E�H������-XSDC3
'    For i = 0 To recCnt
'        typXSDC3upd(i).LENC3 = "0"   '����
'    Next
    
    '��s�ǎ���-XSDC4��
    AccFlg = False
    For i = 1 To recCnt3
        For j = 1 To recCnt
            If p_Furyo(i).XTALC4 = SetRecA(j).CRYNUMCA And _
                                            p_Furyo(i).HINBC4 = SetRecA(j).HINBCA Then
'        With p_Furyo(i)
'            If regFLG = "Y" Then
'                .PUCUTLC4 = p_Hinban_c(i).GNLCA   '�o�^����(c)
'            Else
'                .PUCUTLC4 = p_Hinban_b(i).GNLCA   '�O�H������(b)
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
            '��v����i�Ԃ���������ں��ނ̏����Z�b�g
                With p_Furyo(i)
                    If p_Block_b.CRYNUMC2 = p_Block_c.CRYNUMC2 Then  'b=c (��ۯ�)
                        .KCKNTC4 = p_Block_b.KCNTC2 + 1    '�H���A�Ԃ��{�P���ăZ�b�g
                    Else
                        .KCKNTC4 = 1                       '�H���A�ԂɂP���Z�b�g
                    End If
                    .WKKTC4 = strNowCd                     '�H��
                End With
            End If
            
            '�o�^
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
    
    
    '�ᕪ������(SXL)�o�^-XSDCB��      ����������(�i��)��SXL�P�ʂɏW��
    If SXLflg = True Then
        sDbName = "(XSDCB)"
        Call MakeSXLinfo(SetRecA, typSXL)  '��������(SXL)���쐬
        recCnt = UBound(typSXL)
'        If strNowCd = "CC710" Then     '�����w�����͂̏ꍇ
'            For i = 1 To recCnt
'                '�o�^
'                If typSXL(i).SXLIDCB <> "" Then
'                    If CreateXSDCB(typSXL(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
'                        SetPattern2 = FUNCTION_RETURN_FAILURE
'                        p_Error = strErrMsg
'                        GoTo proc_exit
'                    End If
'                End If
'            Next
'        Else                           'WF�������o�A�������ύX�̏ꍇ
'            For i = 1 To recCnt
'                '�X�V
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
                'ں��ނ��Ȃ��ꍇ(�o�^)
                
                    If CreateXSDCB(typSXL(i), strErrMsg) = FUNCTION_RETURN_FAILURE Then
                        SetPattern2 = FUNCTION_RETURN_FAILURE
                        p_Error = strErrMsg
                        GoTo proc_exit
                    End If
                    
                Else
                'ں��ނ�����ꍇ(�X�V)
                
                    '�����w�����́A�����ŏI���o���͂̏ꍇ�A�������v���X
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
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    SetPattern2 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�V���R���~���̏d�ʂ����߂�
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :dblDiameter   ,I  ,Double    ,���a(mm)
'          :dblHeight     ,I  ,Double    ,����(mm)
'          :�߂�l        ,O  ,Double    ,�d��(g)
'����      :
'����      :2001/06/29 �쐬  �쑺
'          :2002/08/13 Y.Ohno s_cmmc001z ���R�s�[
Private Function WeightOfCylinder(ByVal dblDiameter As Double, ByVal dblHeight As Double) As Double
Dim dblRadius As Double

    dblRadius = dblDiameter / 2#
    WeightOfCylinder = Int(HIJU_SILICONE * cdblPI * (dblRadius ^ 2) * dblHeight)
End Function






'�T�v      :�u���b�N���̍Ń{�g���i�Ԃ��擾����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :Pi_CRYNUM       ,I  ,String     ,�Ώۂ̃u���b�NID
'          :tmpXSDCA        ,O  ,typ_XSDCA  ,�Ń{�g���i��
'����      :�����ɂ́A�u���b�NID�E�����ԍ����Z�b�g���܂�
'����      :2002/08/05 M.Tomita
Public Function GetBottomHinban(Pi_CRYNUM As String, tmpXSDCA As typ_XSDCA) As FUNCTION_RETURN


    '�ϐ��̒�`
    Dim fndXSDCA()  As typ_XSDCA
    Dim strWhere    As String
    Dim strOrder    As String
    Dim sql         As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcXSDC_Basic.bas -- Function GetBottomHinban" '2002/08/05���_�ł�bas���ۗ� M.TOMITA
    
    GetBottomHinban = FUNCTION_RETURN_SUCCESS
    
    'WHERE������
    strWhere = "WHERE CRYNUMCA = '" & Pi_CRYNUM & "' AND LIVKCA = '0'"
    'OrderBy
    strOrder = "ORDER BY INPOSCA "
    '�Y���f�[�^�̎擾
''    If DBDRV_GetXSDCA(fndXSDCA(), strWhere) = FUNCTION_RETURN_SUCCESS And UBound(fndXSDCA) > 0 Then '2002/08/22 �C�� in FFC����
    If DBDRV_GetXSDCA(fndXSDCA(), strWhere, strOrder) = FUNCTION_RETURN_SUCCESS And UBound(fndXSDCA) > 0 Then
        '�l�Z�b�g
        tmpXSDCA = fndXSDCA(UBound(fndXSDCA))
        GetBottomHinban = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    GetBottomHinban = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�_���i�Ԃ̑��݃`�F�b�N������
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :Pi_CRYNUM     ,I  ,String      ,�����ԍ�
'          :�߂�l         ,O�_���i�Ԃ����݂���ꍇ"0"
'                         �@ �_���i�Ԃ����݂��Ȃ��ꍇ"-1"
'����      :�����ɂ͌����ԍ����Z�b�g���܂�
'����      :2002/08/05 M.Tomita
Public Function Nerai_Hinban_Existence_check(Pi_CRYNUM As String) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset
Dim RET As String
Dim w_Nerai_Hinban As String


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcXSDC_Basic.bas -- Function Nerai_Hinban_Existence_check"

    Nerai_Hinban_Existence_check = FUNCTION_RETURN_SUCCESS

    '***�_���i�Ԃ��擾
    sql = ""
    sql = sql & "select PUHINBC1 from XSDC1 "
    sql = sql & "where XTALC1 = '" & Pi_CRYNUM & "'"

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP) '�f�[�^�𒊏o����
    
    If rs.RecordCount = 0 Then '���R�[�h���Ȃ��ꍇ�͐���I��
        rs.Close
        Nerai_Hinban_Existence_check = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    Else
        w_Nerai_Hinban = rs.Fields("PUHINBC1") '*�_���i��
    End If

    '***��������(�i��)�ɏ�L�ŋ��߂�"�_���i��"���܂ރu���b�N�����邩�`�F�b�N����B
    sql = ""
    sql = sql & "select CRYNUMCA from XSDCA "
    sql = sql & "where XTALCA = '" & Pi_CRYNUM & "' and " & _
                      "HINBICA = '" & w_Nerai_Hinban & "' "
    sql = sql & "and KANKCA = '0'"
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP) '�f�[�^�𒊏o����
    
    If rs.RecordCount = 0 Then '���R�[�h���Ȃ��ꍇ�͐���I��
        rs.Close
        Nerai_Hinban_Existence_check = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    
    Nerai_Hinban_Existence_check = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function
proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    Nerai_Hinban_Existence_check = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'�T�v      :���H�敪(���������u���b�N)
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :tblXSDC2      ,O  ,typ_XSDC2    ,���������i�u���b�N�j
'          :tblXSDC2_b    ,I  ,typ_XSDC2    ,���������i�u���b�N�j�O�H��
'          :tblXSDCA      ,I  ,typ_XSDCA    ,���������i�i�ԁj���H��
'          :�߂�l         ,FUNCTION_RETURN  ,����
'
'����      :��ۯ�ID����A���H�敪�̎擾
'����      :2002/08/05 M.Tomita
Public Function Get_ProcDivide(tblXSDC2 As typ_XSDC2, tblXSDC2_b As typ_XSDC2, _
                                    tblXSDCA() As typ_XSDCA) As FUNCTION_RETURN
    Dim sql As String
    Dim tblXSDC1() As typ_XSDC1
    Dim i As Integer
    Dim j As Integer
    
    '�߂�l�̏����ݒ�
    Get_ProcDivide = FUNCTION_RETURN_FAILURE
    
    '***�_���i�Ԃ��擾
    sql = "where XTALC1 = '" & tblXSDC2.XTALC2 & "'"

    If DBDRV_GetXSDC1(tblXSDC1(), sql) = FUNCTION_RETURN_FAILURE Then Exit Function
    '���R�[�h���Ȃ���΃v���V�[�W�����甲����
    If UBound(tblXSDC1) = 0 Then Exit Function
    
    '���������i�i�ԁj���H���ɑ_���i�Ԃ����݂��邩�`�F�b�N
    For i = 1 To UBound(tblXSDCA)
        '�_���i�Ԃ�����΁A���H�敪�P���Z�b�g���I��
        If tblXSDCA(i).HINBCA = tblXSDC1(1).PUHINBC1 Then
            tblXSDC2.KAKOUBC2 = "1"
            For j = 1 To UBound(tblXSDCA)
                Get_ProcDivide = FUNCTION_RETURN_SUCCESS
                tblXSDCA(j).KAKOUBCA = "1"
            Next j
            Exit Function
        End If
    Next i
    
    '���H���ɑ_���i�Ԃ��Ȃ��ꍇ�A�O�H���Ƃ̔�r
    '�O�H���̉��H�敪��1�̏ꍇ
    If tblXSDC2_b.KAKOUBC2 = "1" Then
        '���H�敪2���Z�b�g
        tblXSDC2.KAKOUBC2 = "2"
        For j = 1 To UBound(tblXSDCA)
            tblXSDCA(j).KAKOUBCA = "2"
        Next j
    '����ȊO
    Else
        '�O�H���̉��H�敪���Z�b�g
        tblXSDC2.KAKOUBC2 = tblXSDC2_b.KAKOUBC2
        For j = 1 To UBound(tblXSDCA)
            tblXSDCA(j).KAKOUBCA = tblXSDC2_b.KAKOUBC2
        Next j
    End If
    
    Get_ProcDivide = FUNCTION_RETURN_SUCCESS
    
    
End Function
'
''�T�v      :���H�敪(���������i��)
''���Ұ�    :�ϐ���        ,IO ,�^           ,����
''          :tblXSDCA      ,I  ,typ_XSDCA    ,���������i�i�ԁj���H��
''          :�߂�l         ,���H�敪
''
''����      :��ۯ�ID����A���H�敪�̎擾
''����      :2002/08/05 M.Tomita
'Public Function Get_ProcDivide_Hin(tblXSDCA As typ_XSDCA) As Integer
'    Dim sql As String
'    Dim tblXSDC1() As typ_XSDC1
'    Dim tblXSDCA_b() As typ_XSDCA
'    Dim i As Integer
'    Dim exitFlg As Boolean
'
'    '�����l
'    Get_ProcDivide_Hin = 0
'
'    '***�_���i�Ԃ��擾
'    sql = "where XTALC1 = '" & tblXSDCA.XTALCA & "'"
'
'    If DBDRV_GetXSDC1(tblXSDC1(), sql) = FUNCTION_RETURN_FAILURE Then Exit Function
'    '���R�[�h���Ȃ���΃v���V�[�W�����甲����
'    If UBound(tblXSDC1) = 0 Then Exit Function
'
'    '���������i�i�ԁj���H���ɑ_���i�Ԃ����݂��邩�`�F�b�N
'    '����΁A���H�敪�P���Z�b�g���ăv���V�[�W�����甲����
'    If tblXSDCA.HINBCA = tblXSDC1(1).PUHINBC1 Then
'        Get_ProcDivide_Hin = 1
'        Exit Function
'    End If
'
'    '�O�H���ɑ_���i�Ԃ�����΁A���H�敪�ɂQ���Z�b�g
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
'�T�v      :���a�̎Z�o
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :crynum        ,I  ,String           ,�����ԍ�Or�u���b�NID
'          :diameter      ,O  ,Double           ,���a
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,����
'
'����      :�����ԍ��E��ۯ�ID����A���a���Z�o
'����      :2002/08/09 H.FURUYA
Public Function GetDiameter(CRYNUM As String, DIAMETER As Double) As FUNCTION_RETURN
    Dim JudgKakou As Judg_Kakou
    Dim sumDiameter As Double
    Dim rs As OraDynaset
    Dim sql As String
    
    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcXSDC_Basic -- Function GetDiameter"
    GetDiameter = FUNCTION_RETURN_FAILURE
    
    
    '���H���т̎擾
'    If scmzc_getKakouJiltuseki(CRYNUM, JudgKakou) = FUNCTION_RETURN_SUCCESS And _
'            JudgKakou.TOP(1) <> "-1" Then
    If scmzc_getKakouJiltuseki(CRYNUM, JudgKakou) = FUNCTION_RETURN_SUCCESS And _
            CInt(JudgKakou.TOP(1)) <> -1 Then
            
        '���ς��Z�b�g���v���V�[�W�����甲����
        sumDiameter = JudgKakou.TOP(1) + JudgKakou.TOP(2) + JudgKakou.TAIL(1) + JudgKakou.TAIL(2)
        DIAMETER = sumDiameter / 4#
        '�߂�l�ɐ������Z�b�g
        GetDiameter = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
        
    '�擾�Ɏ��s�����ꍇ�AH�O�O�S�̃f�[�^���擾����
    sql = "SELECT DM1, DM2, DM3 FROM TBCMH004 " & _
          "WHERE CRYNUM ='" & CRYNUM & "'"

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    '���R�[�h���Ȃ���Ύ��s���Z�b�g���ăv���V�[�W�����甲����
    If rs.RecordCount = 0 Then
        rs.Close
        GetDiameter = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    '����������A���ς��Z�b�g
    sumDiameter = CDbl(rs("DM1")) + CDbl(rs("DM2")) + CDbl(rs("DM3"))
    DIAMETER = sumDiameter / 3#
    
    '�߂�l�ɐ������Z�b�g
    GetDiameter = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    GetDiameter = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :���H���т̎擾�h���C�o
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :BLOCKID        ,I   ,String            ,�����ԍ�or�u���b�NID
'          :Jiltuseki      ,O   ,Judg_Kakou        ,���H����
'      �@�@:�߂�l          , O  , FUNCTION_RETURN�@, �ǂݍ��݂̐���
'����      :
'����      :2002/04/17 ���� �M�� �쐬
Private Function scmzc_getKakouJiltuseki(BLOCKID As String, Jiltuseki As Judg_Kakou) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String
    Dim recCnt As Integer
    Dim c0 As Integer
    Dim AGRFlag As Boolean
    Dim ans As String
    Dim tINGOTPOS As Integer
    
    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcXSDC_Basic -- Function scmzc_getKakouJiltuseki"
    scmzc_getKakouJiltuseki = FUNCTION_RETURN_FAILURE
    
    '�Ώۃu���b�N�̉��H���т̏�����
    For c0 = 1 To 2
        Jiltuseki.TAIL(c0) = -1
        Jiltuseki.TOP(c0) = -1
        Jiltuseki.DPTH(c0) = -1
        Jiltuseki.WIDH(c0) = -1
    Next
    Jiltuseki.pos = ""
'2003/10/18 �폜 SystemBrain -------------------------------------------��
'    If Left(BLOCKID, 1) = "8" Then
'        '�w���P�����̏ꍇ
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
'2003/10/18 �폜 SystemBrain -------------------------------------------��
        '�����グ�����̏ꍇ
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
'    End If                         '2003/10/18 �폜 SystemBrain

    scmzc_getKakouJiltuseki = FUNCTION_RETURN_SUCCESS
proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    scmzc_getKakouJiltuseki = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'�T�v      :�`���[�W�ʂ̐ݒ�̗L�����`�F�b�N(�����N���X�Ȃ��j
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :blkID         ,I  ,String           ,�u���b�NID
'          :charge        ,O  ,LOng             ,�`���[�W��
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,����
'
'����      :
'����      :2002/08/15 H.FURUYA
Public Function chkCharge_2(blkID As String, CHARGE As Long, pHinban As String) As FUNCTION_RETURN
    Dim sqlWhere        As String                   'WHERE������
    Dim tblXSDC1()      As typ_XSDC1                '�������グ�e�[�u��
    Dim tblXSDC1_Up     As typ_XSDC1_Update         '�������グ�e�[�u��(�X�V�p�j
    Dim tblXSDCA()      As typ_XSDCA                '���������i�ԃe�[�u��
    Dim i               As Integer
    Dim nowtime         As Date                     '���ް�����@05/08/31 ooba
    
   
    chkCharge_2 = FUNCTION_RETURN_FAILURE
    
    '�`���[�W�ʂ̏�����
    CHARGE = -1
    
    '���グ���т̎擾
    'WHERE����
'2002/09/04
'    sqlWhere = "WHERE XTALC1 = '" & Left(blkID, 8) & "0000" & "' AND KAKOUBC1 ='0'"
    sqlWhere = "WHERE XTALC1 = '" & Left(blkID, 9) & "000" & "' AND KAKOUBC1 ='0'"
    
    '���R�[�h�Z�b�g�̎擾(���s������v���V�[�W�����甲����j
    If DBDRV_GetXSDC1(tblXSDC1, sqlWhere) = FUNCTION_RETURN_FAILURE Then Exit Function
    '�f�[�^���Ȃ���ΐ������Z�b�g���ăv���V�[�W�����甲����
    If UBound(tblXSDC1) = 0 Then
        chkCharge_2 = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If
    
    '�O�H���̌����ɂ˂炢�i�Ԃ����݂��邩�`�F�b�N
    'WHERE����
'2002/09/04
    'sqlWhere = "WHERE XTALCA = '" & Left(blkID, 8) & "0000" & "' AND " & _

    sqlWhere = "WHERE XTALCA = '" & Left(blkID, 9) & "000" & "' AND " & _
                     "HINBCA ='" & tblXSDC1(1).PUHINBC1 & "' AND " & _
                     "CRYNUMCA !='" & blkID & " ' AND " & _
                     "LIVKCA = '0'"
                 
    '���R�[�h�Z�b�g�̎擾(���s������v���V�[�W�����甲����j
    If DBDRV_GetXSDCA(tblXSDCA, sqlWhere) = FUNCTION_RETURN_FAILURE Then Exit Function
    '�f�[�^������ΐ������Z�b�g���ăv���V�[�W�����甲����
    If UBound(tblXSDCA) > 0 Then
        chkCharge_2 = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If
                  
    '�X�V�p�e�[�u���Ƀf�[�^���Z�b�g
    tblXSDC1_Up.KAKOUBC1 = "1"      '���H�敪
'    tblXSDC1_Up.KEIDAYC1 = Now     '�v����t
    nowtime = getSvrTime()          '���ް�����擾�@05/08/31 ooba
    tblXSDC1_Up.KEIDAYC1 = nowtime  '�v����t�@05/08/31 ooba
    
    'WHERE����
'2002/09/04
    'sqlWhere = "WHERE XTALC1 = '" & Left(blkID, 8) & "0000" & "' AND KAKOUBC1 ='0'"
    sqlWhere = "WHERE XTALC1 = '" & Left(blkID, 9) & "000" & "' AND KAKOUBC1 ='0'"
    
    '�f�[�^�̍X�V(���s������v���V�[�W�����甲����j
    If UpdateXSDC1(tblXSDC1_Up, sqlWhere) = FUNCTION_RETURN_FAILURE Then Exit Function
    
    '�߂�l�Ƀ`���[�W�ʂƐ������Z�b�g
    CHARGE = tblXSDC1(1).PUCHAGC1
    pHinban = tblXSDC1(1).PUHINBC1 '2002/11/19 tuku
    chkCharge_2 = FUNCTION_RETURN_SUCCESS
    


End Function





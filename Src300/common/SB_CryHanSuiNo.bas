Attribute VB_Name = "SB_CryHanSuiNo"
Option Explicit

'�����f�[�^�\����
Public Type typ_SeekData
    BLOCKID As String * 12      ' �u���b�N�h�c
    TBKBN   As String * 1       ' T/B�敪(T:Top, B:Bot)
    INPOS   As Integer          ' �������ʒu
    HINB    As tFullHinban      ' �i��(�\����)
    IND     As String           ' ���FLG(0:������, 1:�ʏ�, 2:���f, 3:����)
End Type


'------------------------------------------------
' �������f/����`�F�b�N(���тȂ�)���ʊ֐�
'------------------------------------------------

'�T�v      :�w�肳�ꂽ�]�����ڇ��ɂ��A���f�����肩�𔻒f���A�������f�`�F�b�N(���тȂ�)�A�܂��́A��������`�F�b�N(���тȂ�)���Ăяo���B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :sBlockid      ,I  ,String       :��ۯ�ID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :�����ԍ�
'          :tFullHin      ,I  ,tFullHinban  :�i��(�\����)
'          :iSmplPos      ,I  ,Integer      :�V�T���v���ʒu(mm)
'          :iItemNo       ,I  ,Integer      :�]�����ڇ� =  1 RS     �� ����ΐ͌v�Z
'                                                       =  2 Oi     �� �����1
'                                                       =  3 BMD1   �� �����1
'                                                       =  4 BMD2   �� �����1
'                                                       =  5 BMD3   �� �����1
'                                                       =  6 OSF1   �� �����1
'                                                       =  7 OSF2   �� �����1
'                                                       =  8 OSF3   �� �����1
'                                                       =  9 OSF4   �� �����1
'                                                       = 10 CS     �� �����2(����l,�����l��0����(0<)�̏ꍇ,�����1)
'                                                       = 11 GD     �� �����1
'                                                       = 12 LT     �� �����3
'                                                       = 13 EPD    �� �����2
'          :tSeekData()   ,I  ,typ_SeekData :�����f�[�^�\���̔z��
'          :iHanSuiKBN    ,O  ,Integer      :���f/����敪(0:���f,1:����)
'          :iGetDataNo1   ,O  ,Integer      :���茳�̔z��ԍ��P
'          :iGetDataNo2   ,O  ,Integer      :���茳�̔z��ԍ��Q(���f�����g�p)
'          :�߂�l        ,O  ,Integer      :�`�F�b�N���� = 0 : ����I��(���f/����OK)
'                                                           1 : ����I��(���f/����NG)
'                                                          -1 : ���͈����l�G���[
'                                                          -2 : ��L�ȊO�̃G���[
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funChkSxlHanSuiNo(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                                  iItemNo As Integer, tSeekData() As typ_SeekData, iHanSuiKBN As Integer, _
                                  iGetDataNo1 As Integer, iGetDataNo2 As Integer) As Integer
    Dim retCode As Integer
    
    '�p�����[�^�`�F�b�N
    If (Len(sBlockId) <> 12) Then GoTo ChkSxlHanSuiNoParameterErr
    If (Len(sCryNum) <> 12) Then GoTo ChkSxlHanSuiNoParameterErr
    If UBound(tSeekData) = 0 Then GoTo ChkSxlHanSuiNoParameterErr
    
    '�w�肳�ꂽ�]�����ڇ��ɂ��A���f�����肩�𔻒f���A�������f�`�F�b�N�A�܂��́A��������`�F�b�N���Ăяo���B
    Select Case iItemNo
    Case 1              'RS(���R)
        retCode = funChkSxlSuiteiNo(sBlockId, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, tSeekData(), iGetDataNo1, iGetDataNo2)
        iHanSuiKBN = 1
'    Case 2              'Oi(�_�f�Z�x)
'    Case 3              'BMD1
'    Case 4              'BMD2
'    Case 5              'BMD3
'    Case 6              'OSF1
'    Case 7              'OSF2
'    Case 8              'OSF3
'    Case 9              'OSF4
'    Case 10             'CS(�Y�f�Z�x)
'    Case 11             'GD
'    Case 12             'LT(ײ����)
'    Case 13             'EPD
    Case Else
        GoTo ChkSxlHanSuiNoParameterErr
    End Select
    
    '���ʊ֐��̃`�F�b�N���ʂ𓖊֐��̌��ʂƂ��āA�Ăяo�����֕Ԃ��B
    funChkSxlHanSuiNo = retCode
    Exit Function

ChkSxlHanSuiNoParameterErr:
    funChkSxlHanSuiNo = -1
    Exit Function

ChkSxlHanSuiNoSonotaErr:
    funChkSxlHanSuiNo = -2
End Function

'------------------------------------------------
' ��������`�F�b�N(���тȂ�)
'------------------------------------------------

'�T�v      :�w�肳�ꂽ��񂩂�A��������`�F�b�N(���тȂ�)���s�Ȃ����ʂ�Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :sBlockid      ,I  ,String       :��ۯ�ID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :�����ԍ�
'          :tFullHin      ,I  ,tFullHinban  :�i��(�\����)
'          :iSmplPos      ,I  ,Integer      :�V�T���v���ʒu(mm)
'          :iItemNo       ,I  ,Integer      :�]�����ڇ� =  1 RS     ���Ώ�
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
'          :tSeekData()   ,I  ,typ_SeekData :�����f�[�^�\���̔z��
'          :iGetDataNo1   ,O  ,Integer      :���茳�̔z��ԍ��P
'          :iGetDataNo2   ,O  ,Integer      :���茳�̔z��ԍ��Q
'          :�߂�l        ,O  ,Integer      :�`�F�b�N���� = 0 : ����I��(����OK)
'                                                           1 : ����I��(����NG)
'                                                          -1 : ���͈����l�G���[
'                                                          -2 : ��L�ȊO�̃G���[
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funChkSxlSuiteiNo(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                                  iItemNo As Integer, tSeekData() As typ_SeekData, iGetDataNo1 As Integer, iGetDataNo2 As Integer) As Integer
'    Dim tSiyou          As type_DBDRV_scmzc_fcmkc001c_Siyou
    Dim wGetDataNoTop   As Integer
    Dim wGetDataNoBot   As Integer
    
    '�p�����[�^�`�F�b�N
    If (Len(sBlockId) <> 12) Then GoTo ChkSxlSuiteiNoParameterErr
    If (Len(sCryNum) <> 12) Then GoTo ChkSxlSuiteiNoParameterErr
    If UBound(tSeekData) = 0 Then GoTo ChkSxlSuiteiNoParameterErr
    
    '�w�肳�ꂽ�]�����ڇ����ɕK�v�ȕi�Ԏd�l�l���擾����B�i�w�肳�ꂽ�]�����ڇ��ɂ��A�������������B�j
    Select Case iItemNo
    Case 1              'RS(���R)
'        If funGet_TBCME018(tFullHin, tSiyou) = FUNCTION_RETURN_FAILURE Then GoTo ChkSxlSuiteiNoNG
    Case 2 To 13        'Oi(�_�f�Z�x),BMD1,BMD2,BMD3,OSF1,OSF2,OSF3,OSF4,CS(�Y�f�Z�x),GD,LT(ײ����),EPD
        GoTo ChkSxlSuiteiNoNG
    Case Else
        GoTo ChkSxlSuiteiNoParameterErr
    End Select

    '�������茳�̔z��ԍ��̎擾�i�������茳�̔z��ԍ����擾�ł����琄��n�j�Ƃ���B�j
    If funGetSuiteiNo(sBlockId, sTB, sCryNum, tFullHin, iSmplPos, iItemNo, tSeekData(), _
                                            wGetDataNoTop, wGetDataNoBot) <> 0 Then GoTo ChkSxlSuiteiNoNG

    '�擾�z��ԍ��Ɩ߂�l�̐ݒ�
    iGetDataNo1 = wGetDataNoTop
    iGetDataNo2 = wGetDataNoBot
    
    funChkSxlSuiteiNo = 0
    Exit Function

ChkSxlSuiteiNoNG:
    funChkSxlSuiteiNo = 1
    Exit Function

ChkSxlSuiteiNoParameterErr:
    funChkSxlSuiteiNo = -1
    Exit Function

ChkSxlSuiteiNoSonotaErr:
    funChkSxlSuiteiNo = -2
End Function

'------------------------------------------------
' ��������l�擾(���тȂ�)
'------------------------------------------------

'�T�v      :�w�肳�ꂽ�V����وʒu��񂩂�A�������茳�ʒu�P�ƌ������茳�ʒu�Q�������f�[�^�\���̔z���茟�����A���ꂼ��̔z��ԍ���Ԃ��B
'           ���茳�ʒu�P�Ɛ��茳�ʒu�Q����������ꍇ�A�V�T���v���ʒu����ł��߂�TOP�ް��ƍł��߂�BOT�ް��̈ʒu���ΏۂƂȂ�B
'           �V����وʒu�̕i�Ԏd�l��TOP�^BOT�ʒu�̕i�Ԏd�l���A���ꂼ��u3�_����v�u5�_����v�ł�����݂��l�����邪�A��������݂ɂ���Đ���ۂ𔻒f����B
'           ����ۂ̔��f�Ƃ��āAXSDC1��SUIFLGC1�̒l(0:���苖��,1:����֎~)���l������B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :sBlockid      ,I  ,String       :��ۯ�ID
'          :sTB           ,I  ,String       :T:Top,B:Bot
'          :sCryNum       ,I  ,String       :�����ԍ�
'          :tFullHin      ,I  ,tFullHinban  :�i��(�\����)
'          :iSmplPos      ,I  ,Integer      :�V�T���v���ʒu(mm)
'          :iItemNo       ,I  ,Integer      :�]�����ڇ� =  1 RS     ���Ώ�
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
'          :tSeekData()   ,I  ,typ_SeekData :�����f�[�^�\���̔z��
'          :iGetDataNo1   ,O  ,Integer      :���茳�̔z��ԍ��P
'          :iGetDataNo2   ,O  ,Integer      :���茳�̔z��ԍ��Q
'          :�߂�l        ,O  ,Integer      :�擾���� = 0 : ����I��
'                                                       1 : ����I��(�Y���T���v���Ȃ�)
'                                                      -1 : �ُ�I��
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Public Function funGetSuiteiNo(sBlockId As String, sTB As String, sCryNum As String, tFullHin As tFullHinban, iSmplPos As Integer, _
                               iItemNo As Integer, tSeekData() As typ_SeekData, iGetDataNo1 As Integer, iGetDataNo2 As Integer) As Integer
    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    Dim wFind       As Boolean
    
    Dim getNewSpec  As String       '�V����وʒu���R�d�l�l
    
    Dim getTopNo    As Integer      'TOP�ʒu�z��ԍ�
    Dim getTopBlkID As String       'TOP�ʒu��ۯ�ID
    Dim getTopTB    As String       'TOP�ʒuT/B�敪
    Dim getTopHin   As tFullHinban  'TOP�ʒu�i��
    Dim getTopSpec  As String       'TOP�ʒu���R�d�l�l
    Dim getTopPtrn  As String       'TOP�ʒu����ݺ���
    
    Dim getBotNo    As Integer      'BOT�ʒu�z��ԍ�
    Dim getBotBlkID As String       'BOT�ʒu��ۯ�ID
    Dim getBotTB    As String       'BOT�ʒuT/B�敪
    Dim getBotHin   As tFullHinban  'BOT�ʒu�i��
    Dim getBotSpec  As String       'BOT�ʒu���R�d�l�l
    Dim getBotPtrn  As String       'BOT�ʒu����ݺ���
    
    '�p�����[�^�`�F�b�N
    If (Len(sBlockId) <> 12) Then GoTo GetSuiteiNoParameterErr
    If (Len(sCryNum) <> 12) Then GoTo GetSuiteiNoParameterErr
    If UBound(tSeekData) = 0 Then GoTo GetSuiteiNoParameterErr
    
    '����ۂ̔��f �� XSDC1��SUIFLGC1�̒l(0:���苖��,1:����֎~)
''    sql = "select SUIFLGC1 from XSDC1 where XTALC1 = '" & sCryNum & "'"       2003/10/14
    sql = "select SUIFLG from XSDC1 where XTALC1 = '" & sCryNum & "'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
''    If rs.EOF Or rs("SUIFLGC1") <> "0" Then           2003/10/14
    If rs.EOF Or rs("SUIFLG") <> "0" Then
        Set rs = Nothing
        GoTo GetSuiteiNoEmpty
    End If
    Set rs = Nothing
    
    '�w�肳�ꂽ�V�T���v���ʒu���Ō����f�[�^�\���̔z��(����)���������A���茳�̔z��ԍ������肷��B
    '�ᐄ�茳�z��ԍ��P(TOP�ʒu)�̎擾��
    wFind = False
'    For getTopNo = UBound(tSeekData) To 0 Step -1
'        If (tSeekData(getTopNo).TBKBN = "T") And (tSeekData(getTopNo).INPOS < iSmplPos) Then
    For getTopNo = 0 To UBound(tSeekData)
        If (tSeekData(getTopNo).TBKBN = "T") And (tSeekData(getTopNo).INPOS < iSmplPos) And (tSeekData(getTopNo).IND = "1") Then
            wFind = True
            Exit For
        End If
    Next getTopNo
'    getTopNo = 0                            '���茳TOP�́A�擪�ʒu�Ƃ���
'    wFind = True
    
    If wFind = False Then GoTo GetSuiteiNoEmpty

    'TOP�ʒu�f�[�^�̐ݒ�
    With tSeekData(getTopNo)
        getTopBlkID = .BLOCKID                  'TOP�ʒu��ۯ�ID
        getTopTB = .TBKBN                       'TOP�ʒuT/B�敪
        getTopHin.hinban = .HINB.hinban         'TOP�ʒu�i��
        getTopHin.mnorevno = .HINB.mnorevno     'TOP�ʒu���i�ԍ������ԍ�
        getTopHin.factory = .HINB.factory       'TOP�ʒu�H��
        getTopHin.opecond = .HINB.opecond       'TOP�ʒu���Ə���
    End With

    '�ᐄ�茳�z��ԍ��Q(BOT�ʒu)�̎擾��
    wFind = False
'    For getBotNo = 0 To UBound(tSeekData)
'        If (tSeekData(getBotNo).TBKBN = "B") And (tSeekData(getBotNo).INPOS > iSmplPos) Then
    For getBotNo = UBound(tSeekData) To 0 Step -1
        If (tSeekData(getBotNo).TBKBN = "B") And (tSeekData(getBotNo).INPOS > iSmplPos) And (tSeekData(getBotNo).IND = "1") Then
            wFind = True
            Exit For
        End If
    Next getBotNo
'    getBotNo = UBound(tSeekData)            '���茳BOT�́A�ŏI�ʒu�Ƃ���
'    wFind = True
    
    If wFind = False Then GoTo GetSuiteiNoEmpty

    'BOT�ʒu�f�[�^�̐ݒ�
    With tSeekData(getBotNo)
        getBotBlkID = .BLOCKID                  'Bot�ʒu��ۯ�ID
        getBotTB = .TBKBN                       'Bot�ʒuT/B�敪
        getBotHin.hinban = .HINB.hinban         'Bot�ʒu�i��
        getBotHin.mnorevno = .HINB.mnorevno     'Bot�ʒu���i�ԍ������ԍ�
        getBotHin.factory = .HINB.factory       'Bot�ʒu�H��
        getBotHin.opecond = .HINB.opecond       'Bot�ʒu���Ə���
    End With
    
    '�e�i�Ԃ̔��R�d�l�l�擾
    '��w�肳�ꂽ�V�T���v���ʒu��
    getNewSpec = funGetSuiSpecRS(tFullHin)
    If getNewSpec = " " Then GoTo GetSuiteiNoEmpty
    
    '�ᐄ�茳�T���v���h�c�P(TOP�ʒu)��
    getTopSpec = funGetSuiSpecRS(getTopHin)
    If getTopSpec = " " Then GoTo GetSuiteiNoEmpty
    
    '�ᐄ�茳�T���v���h�c�Q(BOT�ʒu)��
    getBotSpec = funGetSuiSpecRS(getBotHin)
    If getBotSpec = " " Then GoTo GetSuiteiNoEmpty

    '�R�[�hDB�擾�֐����Ăяo����R�[�h�e�[�u��������R����p�^�[���R�[�h���擾����
    '�ᐄ�茳�T���v���h�c�P �� �V�T���v���ʒu��
    If funCodeDBGet("SB", "ST", getTopSpec, 1, getNewSpec, getTopPtrn) <> 0 Then GoTo GetSuiteiNoParameterErr
    If getTopPtrn <> "A" And getTopPtrn <> "B" Then GoTo GetSuiteiNoEmpty
    
    '�ᐄ�茳�T���v���h�c�Q �� �V�T���v���ʒu��
    If funCodeDBGet("SB", "ST", getBotSpec, 1, getNewSpec, getBotPtrn) <> 0 Then GoTo GetSuiteiNoParameterErr
    If getBotPtrn <> "A" And getTopPtrn <> "B" Then GoTo GetSuiteiNoEmpty
    
    '�Ăяo�����ւ̌��ʒʒm
    iGetDataNo1 = getTopNo          '���茳�̔z��ԍ��P
    iGetDataNo2 = getBotNo          '���茳�̔z��ԍ��Q
    
    funGetSuiteiNo = 0
    Exit Function

GetSuiteiNoEmpty:
    funGetSuiteiNo = 1
    Exit Function

GetSuiteiNoParameterErr:
    funGetSuiteiNo = -1
End Function

'------------------------------------------------
' �������� ���R�d�l�l�擾�֐�
'------------------------------------------------

'�T�v      :�w�肳�ꂽ�i�Ԃ���ATBCME018���������A���R�d�l�l(�iSX���R����ʒu_��)���擾���Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :tFullHin      ,I  ,tFullHinban  :�i��(�\����)
'          :�߂�l        ,O  ,Sting        :���R�d�l�l(�iSX���R����ʒu_��)
'                                            (�擾�ł��Ȃ��ꍇ�́A�󔒂�Ԃ�)
'����      :
'����      :2003/09/05 �V�K�쐬�@�V�X�e���u���C��

Private Function funGetSuiSpecRS(tFullHin As tFullHinban) As String
    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    
    '�w�肳�ꂽ�i�Ԃ���TBCME018��HSXRSPOI(�iSX���R����ʒu_��)����������B
    sql = "select HSXRSPOI from TBCME018 "
    sql = sql & "where HINBAN = '" & tFullHin.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tFullHin.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tFullHin.factory & "' and "
    sql = sql & "      OPECOND = '" & tFullHin.opecond & "'"
    
    'SQL���̎��s
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Then
        funGetSuiSpecRS = " "
        Set rs = Nothing
        Exit Function
    End If
    
    'TOP�ʒu�f�[�^�̐ݒ�
    funGetSuiSpecRS = rs("HSXRSPOI")
    Set rs = Nothing

End Function

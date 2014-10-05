Attribute VB_Name = "s_cmzclabel"
'**************************************************
'   �����V�X�e���^�o�[�R�[�h���x���@�\
'
'   �v���O������    : ���x�����s�֐�
'   �t�@�C�����@    : s_cmzclabel.bas
'   �쐬�ҁ@�@�@    : JCE
'   �쐬���@�@�@    : 2001/08/03
'
'**************************************************

Option Explicit

'===============================================================================
'   ���e�����l��`
'===============================================================================
'ORACLE Object for OLE �萔
Private Const ORADYN_DEFAULT = &H0&          '�_�C�i�Z�b�g�̏����ݒ�p�����[�^

'���x�����
Private Const cLBL_INGOT   As String = "01"  '�C���S�b�g���x��
Private Const cLBL_TOPBTM  As String = "02"  '�g�b�v�E�{�g�����x��"
Private Const cLBL_NWMTRL  As String = "03"  '�V�������x��"
Private Const cLBL_SBMTRL  As String = "04"  '�����������x���i���O�j"
Private Const cLBL_SAMTRL  As String = "05"  '�����������x���i����j"
Private Const cLBL_CRYCAT  As String = "06"  '�N���X�^���J�^���O���x��"
Private Const cLBL_BLOCK   As String = "07"  '�u���b�N���x��"
Private Const cLBL_KARI    As String = "15"  '��������ۯ�����"
'Add Start 2011/04/15 SMPK Nakamura FRS�V�X�e�����Ή�
Private Const cLBL_FRS     As String = "16"  'FRS���胉�x��
'Add End 2011/04/15 SMPK Nakamura FRS�V�X�e�����Ή�

'===============================================================================
'   �ϐ���`
'===============================================================================
'���x���v�����^�v���e�[�u������
Private mdtmQueDate As String               '�L���[���t
Private mstrReqKind As String               '����v���敪
Private mstrPrintKind As String             '������
Private mstrEndFlg As String                '�����敪
Private mstrStatus As String                '�I���X�e�[�^�X
Private mstrBlockIDUmu As String            '�u���b�N�h�c�L���敪
Private mstrProcCode As String              '�H���R�[�h
Private mstrEtcPrKind As String             '���̑����x�����
Private mstrCryNum As String                '�����ԍ�
Private mintIngotPos As Integer             '�������ʒu
Private mintSmplNo As Long                  '�T���v���m���D Integer��Long 6���Ή� 2007/05/28 SETsw kubota
Private mstrMtrlNum As String               '�����ԍ�
Private mstrSmtrlNum As String              '���������ԍ�
Private mstrBlockID As String               '�u���b�N�h�c
Private mstrHinban As String                '�i��
Private mintRevNum As Integer               '���i�ԍ������ԍ�
Private mstrFactry As String                '�H��
Private mstrOpecond As String               '���Ə���
Private mstrCryindrs As String              '���������w���iRs�j
Private mstrCryIndoi As String              '���������w���iOi�j
Private mstrCryIndb1 As String              '���������w���iB1�j
Private mstrCryIndb2 As String              '���������w���iB2�j
Private mstrCryIndb3 As String              '���������w���iB3�j
Private mstrCryIndl1 As String              '���������w���iL1�j
Private mstrCryIndl2 As String              '���������w���iL2�j
Private mstrCryIndl3 As String              '���������w���iL3�j
Private mstrCryIndl4 As String              '���������w���iL4�j
Private mstrCryIndcs As String              '���������w���iCs�j
Private mstrCryIndgd As String              '���������w���iGD�j
Private mstrCryIndt As String               '���������w���iT�j
Private mstrCryIndep As String              '���������w���iEPD�j
Private mstrCryIndx As String               '���������w���iX�j  '2009/07/31�ǉ� SETsw kubota
Private mstrCryIndco3 As String             '���������w���iCO3�j'2010/12/14�ǉ� SETsw kubota
Private mstrCryIndc As String               '���������w���iC�j  '2010/12/14�ǉ� SETsw kubota
Private mstrCryIndcj As String              '���������w���iCJ�j '2010/12/14�ǉ� SETsw kubota
Private mstrCryIndcjlt As String            '���������w���iCJLT�j 'Add 2011/02/02 SMPK Miyata
Private mstrCryIndcj2 As String             '���������w���iCJ2�j'2010/12/14�ǉ� SETsw kubota
Private mstrStaffID As String               '�v���S���҂h�c
Private mstrMachine As String               '�v���}�V����
Private mdtmRegDate As Date                 '�o�^���t
Private mdtmUpdDate As Date                 '�X�V���t

'===============================================================================
'   Windows API ��`
'===============================================================================
' �R���s���[�^�����擾
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'*****************************************************
' �֐����@ : �T���v�����x���o�͊֐�
' �ړI���� : �H���Ǘ���ʂ���T���v�����x�����o�͂���B
'
' �����@�@ :
'     strCryNum(i)   : �����ԍ�
'     intIngotPos(i) : �T���v���ʒu
'     intSmplNo(i)   : �T���v��No   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'     strProcCode(i) : �H���R�[�h
'     strStaffID(i)  : �Ј�ID
'     strBlockID(i)  : �u���b�NID  2009/09/16 SETsw kubota
'
' �߂�l�@ :  0 : ����I��
' �@�@�@�@�@ -1 : �ُ�I��
'*****************************************************
Public Function samlabel(StrCryNum As String, _
                          intIngotPos As Integer, _
                          intSmplNo As Long, _
                          StrProcCode As String, _
                          StrStaffId As String _
                        , Optional ByVal StrBlockId As String = "" _
                          ) As Integer

    Dim objRec    As Object
    Dim strMsg    As String
    
    samlabel = -1
    
    strMsg = "�����ԍ�=" & StrCryNum & Chr(13) & _
             "�T���v���ʒu=" & CStr(intIngotPos) & Chr(13) & _
             "�T���v��No=" & CStr(intSmplNo) & Chr(13) & _
             "�H���R�[�h=" & StrProcCode & Chr(13) & _
             "�Ј�ID=" & StrStaffId

'    MsgBox (strMsg)
    
    ' �����ݒ���e�m�F
    If StrCryNum = "" Then Exit Function
    If IsNull(intIngotPos) Then Exit Function
    If IsNull(intSmplNo) Then Exit Function
    If StrProcCode = "" Then Exit Function
    If StrStaffId = "" Then Exit Function
    If StrBlockId = "" Then Exit Function
    
    '�������Ǘ��e�[�u�����擾
    'If Not Select_VECME017(StrCryNum, intIngotPos, intSmplNo, objRec) Then Exit Function
    If Not Select_VECME017(StrBlockId, intIngotPos, intSmplNo, objRec) Then Exit Function

' �ǉ�  2003/10/08 SystemBrain ===================> START
    If objRec.RecordCount = 0 Then
        samlabel = 0
        Exit Function
    End If
' �ǉ�  2003/10/08 SystemBrain ===================> START

    '���x���v�����^�v�����ݒ�
    mdtmQueDate = Format(Now, "yyyymmddhhmmss")         '�L���[���t
    mstrReqKind = "0"                                   '����v���敪�@ 0:�H��������̏o��
    mstrPrintKind = "0"                                 '������ �@�@�@0:�T���v�����x��
    mstrEndFlg = "0"                                    '�����敪�@�@�@ 0:����҂�
    mstrStatus = "0"                                    '�I���X�e�[�^�X 0:����
    mstrBlockIDUmu = "0"                                '�u���b�NID�L���敪(�T���v�����x���ł͖��g�p�̈�'0'��ݒ�)
    mstrProcCode = Trim(StrProcCode)                    '�H���R�[�h
    mstrEtcPrKind = "00"                                '���̑����x�����(�T���v�����x���ł͖��g�p�̈�'00'��ݒ�)
    mstrCryNum = Trim(StrCryNum)                        '�����ԍ�
    mintIngotPos = intIngotPos                          '�������ʒu
    mintSmplNo = intSmplNo                              '�T���v��No
    mstrMtrlNum = "0"                                   '�����ԍ�(�T���v�����x���ł͖��g�p�̈�'0'��ݒ�)
    mstrSmtrlNum = "0"                                  '���������ԍ�(�T���v�����x���ł͖��g�p�̈�'0'��ݒ�)
    mstrBlockID = "0"                                   '�u���b�NID(�T���v�����x���ł͖��g�p�̈�'0'��ݒ�)

    If Not objRec.EOF Then
' �C��  2003/09/26 SystemBrain ===================> START
'        mstrHinban = Trim(objRec.Fields!hinban)         '�i��
'        mintRevNum = objRec.Fields!REVNUM               '���i�ԍ������ԍ�
'        mstrFactry = Trim(objRec.Fields!factory)        '�H��
'        mstrOpecond = Trim(objRec.Fields!opecond)       '���Ə���
'        mstrCryindrs = Trim(objRec.Fields!CRYINDRS)     '���������w���iRs�j
'        mstrCryIndoi = Trim(objRec.Fields!CRYINDOI)     '���������w���iOi�j
'        mstrCryIndb1 = Trim(objRec.Fields!CRYINDB1)     '���������w���iB1�j
'        mstrCryIndb2 = Trim(objRec.Fields!CRYINDB2)     '���������w���iB2�j
'        mstrCryIndb3 = Trim(objRec.Fields!CRYINDB3)     '���������w���iB3�j
'        mstrCryIndl1 = Trim(objRec.Fields!CRYINDL1)     '���������w���iL1�j
'        mstrCryIndl2 = Trim(objRec.Fields!CRYINDL2)     '���������w���iL2�j
'        mstrCryIndl3 = Trim(objRec.Fields!CRYINDL3)     '���������w���iL3�j
'        mstrCryIndl4 = Trim(objRec.Fields!CRYINDL4)     '���������w���iL4�j
'        mstrCryIndcs = Trim(objRec.Fields!CRYINDCS)     '���������w���iCs�j
'        mstrCryIndgd = Trim(objRec.Fields!CRYINDGD)     '���������w���iGD�j
'        mstrCryIndt = Trim(objRec.Fields!CRYINDT)       '���������w���iT�j
'        mstrCryIndep = Trim(objRec.Fields!CRYINDEP)     '���������w���iEPD�j
        mstrHinban = Trim(objRec.Fields!HINBCS)         '�i��
        mintRevNum = objRec.Fields!REVNUMCS             '���i�ԍ������ԍ�
        mstrFactry = Trim(objRec.Fields!FACTORYCS)      '�H��
        mstrOpecond = Trim(objRec.Fields!OPECS)         '���Ə���
        mstrCryindrs = Trim(objRec.Fields!CRYINDRSCS)   '���������w���iRs�j
        mstrCryIndoi = Trim(objRec.Fields!CRYINDOICS)   '���������w���iOi�j
        mstrCryIndb1 = Trim(objRec.Fields!CRYINDB1CS)   '���������w���iB1�j
        mstrCryIndb2 = Trim(objRec.Fields!CRYINDB2CS)   '���������w���iB2�j
        mstrCryIndb3 = Trim(objRec.Fields!CRYINDB3CS)   '���������w���iB3�j
        mstrCryIndl1 = Trim(objRec.Fields!CRYINDL1CS)   '���������w���iL1�j
        mstrCryIndl2 = Trim(objRec.Fields!CRYINDL2CS)   '���������w���iL2�j
        mstrCryIndl3 = Trim(objRec.Fields!CRYINDL3CS)   '���������w���iL3�j
        mstrCryIndl4 = Trim(objRec.Fields!CRYINDL4CS)   '���������w���iL4�j
        mstrCryIndcs = Trim(objRec.Fields!CRYINDCSCS)   '���������w���iCs�j
        mstrCryIndgd = Trim(objRec.Fields!CRYINDGDCS)   '���������w���iGD�j
        mstrCryIndt = Trim(objRec.Fields!CRYINDTCS)     '���������w���iT�j
        mstrCryIndep = Trim(objRec.Fields!CRYINDEPCS)   '���������w���iEPD�j
        mstrCryIndx = Trim(objRec.Fields!CRYINDXCS)     '���������w���iX�j  '2009/07/31�ǉ� SETsw kubota
' �C��  2003/09/26 SystemBrain ===================> END
        'Add Start 2011/02/02 SMPK Miyata
        mstrCryIndc = Trim(objRec.Fields!CRYINDCCS)         '���������w���iC�j
        mstrCryIndcj = Trim(objRec.Fields!CRYINDCJCS)       '���������w���iCJ�j
        mstrCryIndcjlt = Trim(objRec.Fields!CRYINDCJLTCS)   '���������w���iCJLT�j
        mstrCryIndcj2 = Trim(objRec.Fields!CRYINDCJ2CS)     '���������w���iCJ2�j
        'Add End   2011/02/02 SMPK Miyata

        ''C,CJ,CJ2�ǉ��Ή� 2010/12/14 SETsw kubota
        mstrCryIndco3 = Trim(objRec.Fields!CRYINDL4CS)  '���������w���iCO3�j
        'Del Start 2011/02/02 SMPK Miyata
        'mstrCryIndc = "0"
        'mstrCryIndcj = "0"
        'mstrCryIndcj2 = "0"
        'If mstrCryIndco3 = "1" Then
        '    '���������w���iC�j
        '    If objRec.Fields!HSXCHS = "H" _
        '    Or objRec.Fields!HSXCHS = "S" Then
        '        mstrCryIndc = "1"
        '    End If
        '    '���������w���iCJ�j
        '    If objRec.Fields!HSXCJHS = "H" _
        '    Or objRec.Fields!HSXCJHS = "S" Then
        '        mstrCryIndcj = "1"
        '    End If
        '    '���������w���iCJ2�j
        '    If objRec.Fields!HSXCJ2HS = "H" _
        '    Or objRec.Fields!HSXCJ2HS = "S" Then
        '        mstrCryIndcj2 = "1"
        '    End If
        'End If
        'Del End   2011/02/02 SMPK Miyata
    
    End If
    objRec.Close

    mstrStaffID = StrStaffId                            '�v���S����ID
    mstrMachine = StrConv(GetComputerName, vbUpperCase) '�v���}�V����
    
    '���x���v�����^�v���e�[�u�����ǉ�
    If Not Insert_TBCMC001() Then Exit Function
    
    samlabel = 0
    
End Function


'*****************************************************
' �֐����@ : ���̑����x���o�͊֐�
' �ړI���� : �H���Ǘ���ʂ��炻�̑����x�����o�͂���B
'
' �����@�@ :
'     strEtcPrKind(i) : ���̑����[�敪
'                       01: �C���S�b�g���x��
'                       02:�g�b�v�E�{�g�����x��
'                       03:�V�������x��
'                       04:�����������x��(���ޑO)
'                       05:�����������x��(����)
'                       06:�N���X�^���J�^���O���x��
'                       07:�u���b�N���x��
'��������ۯ����ٔ��s�����ǉ��˗��@yakimura 2002.12.12 start
'                       15:��������ۯ�����
'��������ۯ����ٔ��s�����ǉ��˗��@yakimura 2002.12.12 start
'     strKey(i)       : �L�[����
'                       (�����ԍ��^����No.�^�u���b�NID�^��������No.)
'     strBlockID(i)   : �u���b�NID�L���敪
'     strProcCode(i)  : �H���R�[�h
'     strStaffID(i)   : �Ј�ID
'
' �߂�l�@ :  0 : ����I��
' �@�@�@�@�@ -1 : �ُ�I��
'*****************************************************
Public Function etclabel(StrEtcPrKind As String, _
                          strKey As String, _
                          StrBlockId As String, _
                          StrProcCode As String, _
                          StrStaffId As String) As Integer
    Dim strMsg    As String
    
    etclabel = -1
    
    strMsg = "���̑����[�敪=" & StrEtcPrKind & Chr(13) & _
             "�L�[����=" & strKey & Chr(13) & _
             "�u���b�NID�L���敪=" & StrBlockId & Chr(13) & _
             "�H���R�[�h=" & StrProcCode & Chr(13) & _
             "�Ј�ID=" & StrStaffId

'    MsgBox (strMsg)

    ' �����m�F
    If StrEtcPrKind = "" Then Exit Function
    If strKey = "" Then Exit Function
    If StrBlockId = "" Then Exit Function
    If StrProcCode = "" Then Exit Function
    If StrStaffId = "" Then Exit Function
    
    '���x���v�����^�v�����ݒ�
    mdtmQueDate = Format(Now, "yyyymmddhhmmss")         '�L���[���t
    mstrReqKind = "0"                                   '����v���敪�@ 0:�H��������̏o��
    mstrPrintKind = "1"                                 '������ �@�@�@1:���̑����x��
    mstrEndFlg = "0"                                    '�����敪�@�@�@ 0:����҂�
    mstrStatus = "0"                                    '�I���X�e�[�^�X 0:����
    mstrProcCode = Trim(StrProcCode)                    '�H���R�[�h
    mstrEtcPrKind = Trim(StrEtcPrKind)                  '���̑����x�����
                                                        '  01:�C���S�b�g���x��
                                                        '  02:�g�b�v�E�{�g�����x��
                                                        '  03:�V�������x��
                                                        '  04:�����������x��(���ޑO)
                                                        '  05:�����������x��(����)
                                                        '  06:�N���X�^���J�^���O���x��
                                                        '  07:�u���b�N���x��
    
    '�C���S�b�g���x��, �g�b�v�E�{�g�����x��
    If mstrEtcPrKind = cLBL_INGOT Or mstrEtcPrKind = cLBL_TOPBTM Then
        mstrCryNum = Trim(strKey)                       '�����ԍ�
        mstrMtrlNum = "0"                               '�����ԍ�(���g�p�̈�'0'��ݒ�)
        mstrSmtrlNum = "0"                              '���������ԍ�(���g�p�̈�'0'��ݒ�)
        mstrBlockID = "0"                               '�u���b�NID(���g�p�̈�'0'��ݒ�)
        mstrBlockIDUmu = "0"                            '�u���b�NID�L���敪(���g�p�̈�'0'��ݒ�)
    '�V�������x��
    ElseIf mstrEtcPrKind = cLBL_NWMTRL Then
        mstrCryNum = "0"                                '�����ԍ�(���g�p�̈�'0'��ݒ�)
        mstrMtrlNum = Trim(strKey)                      '�����ԍ�
        mstrSmtrlNum = "0"                              '���������ԍ�(���g�p�̈�'0'��ݒ�)
        mstrBlockID = "0"                               '�u���b�NID(���g�p�̈�'0'��ݒ�)
        mstrBlockIDUmu = "0"                            '�u���b�NID�L���敪(���g�p�̈�'0'��ݒ�)
    '�����������x��(���ޑO), �N���X�^���J�^���O���x��
    ElseIf mstrEtcPrKind = cLBL_SBMTRL Or mstrEtcPrKind = cLBL_CRYCAT Then
        mstrCryNum = "0"                                '�����ԍ�(���g�p�̈�'0'��ݒ�)
        mstrMtrlNum = "0"                               '�����ԍ�(���g�p�̈�'0'��ݒ�)
        mstrSmtrlNum = "0"                              '���������ԍ�(���g�p�̈�'0'��ݒ�)
        mstrBlockID = Trim(strKey)                      '�u���b�NID
        mstrBlockIDUmu = Trim(StrBlockId)               '�u���b�NID�L���敪  0:�u���b�NID�L�� 1:�u���b�NID����
    '�����������x��(����)
    ElseIf mstrEtcPrKind = cLBL_SAMTRL Then
        mstrCryNum = "0"                                '�����ԍ�(���g�p�̈�'0'��ݒ�)
        mstrMtrlNum = "0"                               '�����ԍ�(���g�p�̈�'0'��ݒ�)
        mstrSmtrlNum = Trim(strKey)                     '���������ԍ�
        mstrBlockID = "0"                               '�u���b�NID(���g�p�̈�'0'��ݒ�)
        mstrBlockIDUmu = "0"                            '�u���b�NID�L���敪(���g�p�̈�'0'��ݒ�)
    '�u���b�N���x��
    ElseIf mstrEtcPrKind = cLBL_BLOCK Then
        mstrCryNum = "0"                                '�����ԍ�(���g�p�̈�'0'��ݒ�)
        mstrMtrlNum = "0"                               '�����ԍ�(���g�p�̈�'0'��ݒ�)
        mstrSmtrlNum = "0"                              '���������ԍ�(���g�p�̈�'0'��ݒ�)
        mstrBlockID = Trim(strKey)                      '�u���b�NID
        mstrBlockIDUmu = "0"                            '�u���b�NID�L���敪(���g�p�̈�'0'��ݒ�)
'Add Start 2011/04/15 SMPK Nakamura FRS�V�X�e�����Ή�
    'FRS���胉�x��
    ElseIf mstrEtcPrKind = cLBL_FRS Then
        mstrCryNum = "0"                                '�����ԍ�(���g�p�̈�'0'��ݒ�)
        mstrMtrlNum = "0"                               '�����ԍ�(���g�p�̈�'0'��ݒ�)
        mstrSmtrlNum = "0"                              '���������ԍ�(���g�p�̈�'0'��ݒ�)
        mstrBlockID = Trim(strKey)                      '�u���b�NID
        mstrBlockIDUmu = "0"                            '�u���b�NID�L���敪(���g�p�̈�'0'��ݒ�)
'Add End 2011/04/15 SMPK Nakamura FRS�V�X�e�����Ή�
    End If
    
    mintIngotPos = 0                                    '�������ʒu(���̑����x���ł͖��g�p�̈�'0'��ݒ�)
    mintSmplNo = 0                                      '�T���v��No(���̑����x���ł͖��g�p�̈�'0'��ݒ�)
    mstrHinban = "0"                                    '�i��
    mintRevNum = 0                                      '���i�ԍ������ԍ�
    mstrFactry = "0"                                    '�H��
    mstrOpecond = "0"                                   '���Ə���
    mstrCryindrs = "0"                                  '���������w���iRs�j
    mstrCryIndoi = "0"                                  '���������w���iOi�j
    mstrCryIndb1 = "0"                                  '���������w���iB1�j
    mstrCryIndb2 = "0"                                  '���������w���iB2�j
    mstrCryIndb3 = "0"                                  '���������w���iB3�j
    mstrCryIndl1 = "0"                                  '���������w���iL1�j
    mstrCryIndl2 = "0"                                  '���������w���iL2�j
    mstrCryIndl3 = "0"                                  '���������w���iL3�j
    mstrCryIndl4 = "0"                                  '���������w���iL4�j
    mstrCryIndcs = "0"                                  '���������w���iCs�j
    mstrCryIndgd = "0"                                  '���������w���iGD�j
    mstrCryIndt = "0"                                   '���������w���iT�j
    mstrCryIndep = "0"                                  '���������w���iEPD�j
    mstrCryIndx = "0"                                   '���������w���iX�j
    mstrCryIndco3 = "0"                                 '���������w���iCO3�j
    mstrCryIndc = "0"                                   '���������w���iC�j
    mstrCryIndcj = "0"                                  '���������w���iCJ�j
    mstrCryIndcj2 = "0"                                 '���������w���iCJ2�j
    
    mstrStaffID = StrStaffId                            '�v���S����ID
    mstrMachine = StrConv(GetComputerName, vbUpperCase) '�v���}�V����
    
'��������ۯ����ٔ��s�����ǉ��˗��@yakimura 2002.12.12 start
    '�u���b�N���x���@�C�@��������ۯ�����
    If mstrEtcPrKind = cLBL_KARI Then
        mstrCryNum = "0"                                '�����ԍ�(���g�p�̈�'0'��ݒ�)
        mstrMtrlNum = "0"                               '�����ԍ�(���g�p�̈�'0'��ݒ�)
        mstrSmtrlNum = "0"                              '���������ԍ�(���g�p�̈�'0'��ݒ�)
        mstrBlockID = Trim(StrBlockId)                  '�u���b�NID
        mstrHinban = Trim(strKey)                       '�i��
        mstrBlockIDUmu = "0"                            '�u���b�NID�L���敪(���g�p�̈�'0'��ݒ�)
    End If
'��������ۯ����ٔ��s�����ǉ��˗��@yakimura 2002.12.12 end
    
    '���x���v�����^�v���e�[�u�����ǉ�
    If Not Insert_TBCMC001() Then Exit Function
    
    etclabel = 0

End Function

'*****************************************************
' �֐����@ : �����T���v�����Ǘ��e�[�u�����擾�֐�
' �ړI���� : �����T���v�����Ǘ��e�[�u����������擾����B
'
' �����@�@ :
'     strCryNum(i)   : �����ԍ�
'     intIngotPos(i) : �T���v���ʒu
'     intSmplNo(i)   : �T���v��No   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'     objRec(o)      : ��������
'
' �߂�l�@ : True  : ����I��
' �@�@�@     False : �ُ�I��
'*****************************************************
Private Function Select_VECME017(StrCryNum As String, _
                                intIngotPos As Integer, _
                                intSmplNo As Long, _
                                objRec As Object) As Boolean
    Dim strSQL     As String
    Dim strErrMsg  As String   '�G���[���̃��b�Z�[�W

    Select_VECME017 = False

    'SQL�쐬
' �C��  2003/09/26 SystemBrain ===================> START
'    strSQL = "SELECT * FROM VECME017 "
'    strSQL = strSQL & "WHERE CRYNUM = '" & strCryNum & "' "
'    strSQL = strSQL & "AND INGOTPOS = " & intIngotPos & " "
'    strSQL = strSQL & "AND SMPLNO = " & intSmplNo & " "
'    strSQL = strSQL & "AND KTKBN = '0'"
    
    strSQL = "SELECT"
    strSQL = strSQL & " XTALCS"           ' 0:�����ԍ�
    strSQL = strSQL & ",INPOSCS"          ' 1:�������ʒu
    strSQL = strSQL & ",REPSMPLIDCS"      ' 2:��\�T���v��ID
    strSQL = strSQL & ",HINBCS"           ' 3:�i��
    strSQL = strSQL & ",REVNUMCS"         ' 4:���i�ԍ������ԍ�
    strSQL = strSQL & ",FACTORYCS"        ' 5:�H��
    strSQL = strSQL & ",OPECS"            ' 6:���Ə���
    strSQL = strSQL & ",CRYINDRSCS"       ' 7:���FLG(Rs)
    strSQL = strSQL & ",CRYINDOICS"       ' 8:���FLG(Oi)
    strSQL = strSQL & ",CRYINDCSCS"       ' 9:���FLG(Cs)
    strSQL = strSQL & ",CRYINDB1CS"       '10:���FLG(B1)
    strSQL = strSQL & ",CRYINDB2CS"       '11:���FLG(B2)
    strSQL = strSQL & ",CRYINDB3CS"       '12:���FLG(B3)
    strSQL = strSQL & ",CRYINDL1CS"       '13:���FLG(L1)
    strSQL = strSQL & ",CRYINDL2CS"       '14:���FLG(L2)
    strSQL = strSQL & ",CRYINDL3CS"       '15:���FLG(L3)
    strSQL = strSQL & ",CRYINDL4CS"       '16:���FLG(L4)
    strSQL = strSQL & ",CRYINDGDCS"       '17:���FLG(GD)
    strSQL = strSQL & ",CRYINDTCS"        '18:���FLG(T)
    strSQL = strSQL & ",CRYINDEPCS"       '19:���FLG(EPD)
    strSQL = strSQL & ",NVL(CRYINDXCS,'0') CRYINDXCS"     '20:���FLG(X��)
    'Add Start 2011/02/02 SMPK Miyata
    strSQL = strSQL & ",NVL(CRYINDCCS,'0') CRYINDCCS"       '21:���FLG(C)
    strSQL = strSQL & ",NVL(CRYINDCJCS,'0') CRYINDCJCS"     '22:���FLG(CJ)
    strSQL = strSQL & ",NVL(CRYINDCJLTCS,'0') CRYINDCJLTCS" '23:���FLG(CJLT)
    strSQL = strSQL & ",NVL(CRYINDCJ2CS,'0') CRYINDCJ2CS"   '24:���FLG(CJ2)
    'Add End   2011/02/02 SMPK Miyata

    'Del Start 2011/02/02 SMPK Miyata
    ''C,CJ,CJ2�Ή� 2010/12/14 SETsw kubota
    'strSQL = strSQL & ",NVL(E020.HSXCHS , ' ') HSXCHS"          '�iSXL/C�ۏؕ��@�Q��
    'strSQL = strSQL & ",NVL(E020.HSXCJHS , ' ') HSXCJHS"        '�iSXL/CJ�ۏؕ��@�Q��
    'strSQL = strSQL & ",NVL(E020.HSXCJ2HS , ' ') HSXCJ2HS"      '�iSXL/CJ2�ۏؕ��@�Q��
    'Del End   2011/02/02 SMPK Miyata
    
    strSQL = strSQL & "  FROM XSDCS "
    strSQL = strSQL & "     , TBCME020 E020 "      '2010/12/14 SETsw kubota
    
''' 09/03/02 FAE)akiyama start
''    strSQL = strSQL & "WHERE XTALCS = '" & StrCryNum & "' "
'    strSQL = strSQL & "WHERE CRYNUMCS LIKE '" & Left(StrCryNum, 9) & "%' "
''' 09/03/02 FAE)akiyama end
    strSQL = strSQL & " WHERE CRYNUMCS = '" & StrCryNum & "' "
    strSQL = strSQL & " AND INPOSCS = " & intIngotPos & " "
    strSQL = strSQL & " AND REPSMPLIDCS = " & intSmplNo & " "
    strSQL = strSQL & " AND (CRYINDRSCS = '1'"
    strSQL = strSQL & " OR CRYINDOICS = '1'"
    strSQL = strSQL & " OR CRYINDCSCS = '1'"
    strSQL = strSQL & " OR CRYINDB1CS = '1'"
    strSQL = strSQL & " OR CRYINDB2CS = '1'"
    strSQL = strSQL & " OR CRYINDB3CS = '1'"
    strSQL = strSQL & " OR CRYINDL1CS = '1'"
    strSQL = strSQL & " OR CRYINDL2CS = '1'"
    strSQL = strSQL & " OR CRYINDL3CS = '1'"
    strSQL = strSQL & " OR CRYINDL4CS = '1'"
    strSQL = strSQL & " OR CRYINDGDCS = '1'"
    strSQL = strSQL & " OR CRYINDTCS = '1'"
    strSQL = strSQL & " OR CRYINDEPCS = '1'"
    strSQL = strSQL & " OR CRYINDXCS = '1'"
    'Add Start 2011/02/02 SMPK Miyata
    strSQL = strSQL & " OR CRYINDCCS = '1'"
    strSQL = strSQL & " OR CRYINDCJCS = '1'"
    strSQL = strSQL & " OR CRYINDCJLTCS = '1'"
    strSQL = strSQL & " OR CRYINDCJ2CS = '1'"
    'Add End   2011/02/02 SMPK Miyata
    strSQL = strSQL & " )"
'    strSQL = strSQL & "AND BLKKTFLAGCS = '0'"      '��ۯ��m���׸ނ͌Ăь��Ŕ��f���� 2009/09/16 SETsw kubota
' �C��  2003/09/26 SystemBrain ===================> END
    
    'C,CJ,CJ2�Ή� 2010/12/14 SETsw kubota
    strSQL = strSQL & " AND HINBCS = E020.HINBAN(+)"
    strSQL = strSQL & " AND REVNUMCS = E020.MNOREVNO(+)"
    strSQL = strSQL & " AND FACTORYCS = E020.FACTORY(+)"
    strSQL = strSQL & " AND OPECS = E020.OPECOND(+)"

Debug.Print strSQL
    'DB����������
    'SQL���s
    If Not Fun_Ora_Select((strSQL), objRec, True) Then
        Exit Function
    End If
    
' �폜(�Y���ް��Ȃ��ł�װ�Ƃ��Ȃ��Ŗ������Ƃ���) 2003/10/08 SystemBrain ===================> START
'''''    If objRec.RecordCount <= 0 Then
'''''        '�Y���f�[�^�Ȃ�
'''''        strErrMsg = GetMsgStr("ECRY7")
'''''        MsgBox strErrMsg, vbCritical
'''''        Exit Function
'''''    End If
' �폜(�Y���ް��Ȃ��ł�װ�Ƃ��Ȃ��Ŗ������Ƃ���) 2003/10/08 SystemBrain ===================> END
    
    Select_VECME017 = True

End Function

'*****************************************************
' �֐����@ : ���x���v�����^�v���e�[�u�����ǉ��֐�
' �ړI���� : ���x���v�����^�v���e�[�u���ɏ���ǉ�����B
'
' �����@�@ : �Ȃ�
'
' �߂�l�@ : True  : ����I��
' �@�@�@     False : �ُ�I��
'*****************************************************
Private Function Insert_TBCMC001() As Boolean
    Dim strSQL     As String
    Dim strSQL2    As String
    Dim strErrMsg  As String   '�G���[���̃��b�Z�[�W
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    Dim labelCls    As c_cmzcLabel      '�N���X���W���[���p�ϐ�
    Set labelCls = New c_cmzcLabel      '�N���X�I�u�W�F�N�g����
    '<<<<< -------------------------------------------------------END
    
    Insert_TBCMC001 = False

    '���[���o�b�N
    Call Fun_Ora_Rollback(False)

    '�g�����U�N�V�����J�n
    Fun_Ora_BeginTransaction
    
    strSQL = "INSERT INTO TBCMC001 (": strSQL2 = " VALUES (":
    strSQL = strSQL & " QUEDATE":      strSQL2 = strSQL2 & " TO_DATE('" & mdtmQueDate & "','YYYYMMDDHH24MISS')"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrQueDate = mdtmQueDate                        '�L���[���t���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",REQKIND":      strSQL2 = strSQL2 & ",'" & mstrReqKind & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrReqKind = mstrReqKind                        '����v���敪���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",PRINTKIND":    strSQL2 = strSQL2 & ",'" & mstrPrintKind & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrPrintKind = mstrPrintKind                    '�����ނ��v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",ENDFLG":       strSQL2 = strSQL2 & ",'" & mstrEndFlg & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrEndFlg = mstrEndFlg                          '�����敪���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",STATUS":       strSQL2 = strSQL2 & ",'" & mstrStatus & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrStatus = mstrStatus                          '�I���X�e�[�^�X���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",BLOCKIDUMU":   strSQL2 = strSQL2 & ",'" & mstrBlockIDUmu & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrBlockIdUmu = mstrBlockIDUmu                   '�u���b�NID�L���敪���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",PROCCODE":     strSQL2 = strSQL2 & ",'" & mstrProcCode & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrProcCode = mstrProcCode                       '�H���R�[�h���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",ETCPRKIND":    strSQL2 = strSQL2 & ",'" & mstrEtcPrKind & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrEtcPrKind = mstrEtcPrKind                     '���̑����x����ނ��v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYNUM":       strSQL2 = strSQL2 & ",'" & mstrCryNum & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryNum = mstrCryNum                           '�����ԍ����v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",INGOTPOS":     strSQL2 = strSQL2 & "," & mintIngotPos
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.NumIngotPos = mintIngotPos                        '�������ʒu���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",SMPLNO":       strSQL2 = strSQL2 & "," & mintSmplNo
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.NumSmplNo = mintSmplNo                            '�T���v�������v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",MTRLNUM":      strSQL2 = strSQL2 & ",'" & mstrMtrlNum & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrMtrlNum = mstrMtrlNum                          '�����ԍ����v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",SMTRLNUM":     strSQL2 = strSQL2 & ",'" & mstrSmtrlNum & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrSmtrlNum = mstrSmtrlNum                        '���������ԍ����v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",BLOCKID":      strSQL2 = strSQL2 & ",'" & mstrBlockID & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrBlockId = mstrBlockID                          '�u���b�NID���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",HINBAN":       strSQL2 = strSQL2 & ",'" & mstrHinban & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrHinban = mstrHinban                            '�i�Ԃ��v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",REVNUM":       strSQL2 = strSQL2 & "," & mintRevNum
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.NumRevNum = mintRevNum                            '���i�ԍ������ԍ����v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",FACTORY":      strSQL2 = strSQL2 & ",'" & mstrFactry & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrFactory = mstrFactry                           '�H����v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",OPECOND":      strSQL2 = strSQL2 & ",'" & mstrOpecond & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrOpeCond = mstrOpecond                          '���Ə������v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
' �C��  2003/09/26 SystemBrain ===================> START
    strSQL = strSQL & ",CRYINDRS":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryindrs = "1", "1", "0") & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndRs = IIf(mstrCryindrs = "1", "1", "0")   '���������w��(Rs)���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDOI":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndoi = "1", "1", "0") & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndOi = IIf(mstrCryIndoi = "1", "1", "0")   '���������w��(Oi)���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDB1":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndb1 = "1", "1", "0") & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndB1 = IIf(mstrCryIndb1 = "1", "1", "0")   '���������w��(B1)���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDB2":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndb2 = "1", "1", "0") & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndB2 = IIf(mstrCryIndb2 = "1", "1", "0")   '���������w��(B2)���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDB3":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndb3 = "1", "1", "0") & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndB3 = IIf(mstrCryIndb3 = "1", "1", "0")   '���������w��(B3)���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDL1":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndl1 = "1", "1", "0") & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndL1 = IIf(mstrCryIndl1 = "1", "1", "0")   '���������w��(L1)���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDL2":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndl2 = "1", "1", "0") & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndL2 = IIf(mstrCryIndl2 = "1", "1", "0")   '���������w��(L2)���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDL3":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndl3 = "1", "1", "0") & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndL3 = IIf(mstrCryIndl3 = "1", "1", "0")   '���������w��(L3)���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDL4":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndl4 = "1", "1", "0") & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
'L4��CO3�ɕύX 2010/12/14 SETsw kubota
'    labelCls.StrCryIndL4 = IIf(mstrCryIndl4 = "1", "1", "0")   '���������w��(L4)���v���p�e�B�ɃZ�b�g
    labelCls.StrCryIndL4 = "0"
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDCS":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndcs = "1", "1", "0") & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndCs = IIf(mstrCryIndcs = "1", "1", "0")   '���������w��(Cs)���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDGD":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndgd = "1", "1", "0") & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndGD = IIf(mstrCryIndgd = "1", "1", "0")   '���������w��(GD)���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDT":      strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndt = "1", "1", "0") & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndT = IIf(mstrCryIndt = "1", "1", "0")     '���������w��(T)���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDEP":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndep = "1", "1", "0") & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndEP = IIf(mstrCryIndep = "1", "1", "0")   '���������w��(EPD)���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
' �C��  2003/09/26 SystemBrain ===================> END

    strSQL = strSQL & ",STAFFID":      strSQL2 = strSQL2 & ",'" & mstrStaffID & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrStaffId = mstrStaffID                          '�v���S����ID���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",MACHINE":      strSQL2 = strSQL2 & ",'" & mstrMachine & "'"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrMachine = mstrMachine                          '�v���}�V�������v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",REGDATE":      strSQL2 = strSQL2 & ",SYSDATE"
    
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    If Not GetSysdate() Then GoTo proc_exit  '�V�X�e�����t���擾
    
    labelCls.StrRegDate = Format(gsSysdate, "yyyymmddhhmmss")  '�o�^���t���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",UPDDATE":      strSQL2 = strSQL2 & ",SYSDATE"
    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    labelCls.StrUpdDate = Format(gsSysdate, "yyyymmddhhmmss")  '�X�V���t���v���p�e�B�ɃZ�b�g
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ")":             strSQL2 = strSQL2 & ")"

    '���������w��(X)    2009/07/31�ǉ� SETsw kubota
    labelCls.StrCryIndX = IIf(mstrCryIndx = "1", "1", "0")

    '���������w��(C,CJ,CJ2)     2010/12/14�ǉ� SETsw kubota
    labelCls.StrCryIndCO3 = IIf(mstrCryIndco3 = "1", "1", "0")  '���������w��(CO3)
    labelCls.StrCryIndC = IIf(mstrCryIndc = "1", "1", "0")      '���������w��(C)
    labelCls.StrCryIndCJ = IIf(mstrCryIndcj = "1", "1", "0")    '���������w��(CJ)
    labelCls.StrCryIndCJ2 = IIf(mstrCryIndcj2 = "1", "1", "0")  '���������w��(CJ2)

    '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
    '�V���v���Z�X�̃`�F�b�N
    labelCls.Label_Process_Check
    
    '�V�v���Z�X�̏ꍇ
    If labelCls.ProcKubun = True Then
       'KODZ6�ɓo�^
        If labelCls.Label_Facade = False Then
            GoTo proc_exit
        End If
    Else
        '���v���Z�X�̏ꍇ�A�����̏���(TBCMC001�ɓo�^)
        If Not Fun_Ora_Execute(strSQL & strSQL2) Then
            '���[���o�b�N
            If Not Fun_Ora_Rollback(True) Then GoTo proc_exit
            GoTo proc_exit
        End If
    End If
    
    If Not Fun_Ora_Commit() Then GoTo proc_exit
    '<<<<< -------------------------------------------------------END
        
'    If Not Fun_Ora_Execute(strSQL & strSQL2) Then
'        '���[���o�b�N
'        If Not Fun_Ora_Rollback(True) Then Exit Function
'        Exit Function
'    End If
'
'    If Not Fun_Ora_Commit() Then Exit Function
        
    '1�b�ԑ҂�
    Sleep (1000)

    Insert_TBCMC001 = True
    
'>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
proc_exit:
    Set labelCls = Nothing              '�N���X�I�u�W�F�N�g���
'<<<<< -------------------------------------------------------END

End Function

'Oracle�֘A
'===============================================================================
'�֐���     :SELECT�������s���� Fun_Ora_Select
'�@�\       :�ڑ���ɑ΂��āASELECT�������s���܂��B
'-------------------------------------------------------------------------------
'       ���t        ��      �S����      �R�����g
'�쐬   2001/08/01  1.00    JCE
'�X�V
'-------------------------------------------------------------------------------
'�����@     :sPrmSql     �i���s�r�p�k�j
'�@�@�@     :bPrmOutMsg  �i���b�Z�[�W�o�̓t���O(�f�t�H���g:True){True:�o�͂���,False:�o�͂��Ȃ�}�j
'�߂�l     :���� True�A�G���[ False
'�@�@�@     :sPrmRslt    �i�c�a�擾���e�j
'===============================================================================
Private Function Fun_Ora_Select(sPrmSql As String, sPrmRslt As OraDynaset, _
                            Optional bPrmOutMsg As Boolean = True) As Boolean

    Dim lSv_Mouse As Long     '�}�E�X�|�C���^�[�i�[
    Dim strErrMsg As String   '�G���[���̃��b�Z�[�W

    Fun_Ora_Select = False
    
    '�}�E�X�|�C���^��Ԃ̕ۑ�
    lSv_Mouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrSub
    
    '���ʃZ�b�g�̍쐬
    Set sPrmRslt = OraDB.DBCreateDynaset(sPrmSql, ORADYN_DEFAULT)

    Fun_Ora_Select = True
    Screen.MousePointer = lSv_Mouse     '�}�E�X�|�C���^��Ԃ̕���

Exit Function

'�c�a�֘A�G���[�i�������f�j
ErrSub:
    '�G���[���b�Z�[�W����
    strErrMsg = ""
    strErrMsg = strErrMsg & "ErrorCode:" & Err.Number
    strErrMsg = strErrMsg & "  Message:" & Err.Description
    '�G���[���b�Z�[�W�o�́i�w�莞�̂ݏo�́j
    If bPrmOutMsg Then
        MsgBox "�G���[���������܂���" & vbCrLf & strErrMsg, vbCritical
    End If
    Screen.MousePointer = lSv_Mouse     '�}�E�X�|�C���^��Ԃ̕���
End Function

'===============================================================================
'�֐���     :����SQL���s���� Fun_Ora_Execute
'�@�\       :�ڑ���ɑ΂��āASELECT�ȊO��SQL�����s���܂��B
'-------------------------------------------------------------------------------
'       ���t        ��      �S����      �R�����g
'�쐬   2001/08/01  1.00    JCE
'�X�V
'-------------------------------------------------------------------------------
'�����@     :sPrmSql     �i���s�r�p�k�j
'�@�@�@     :bPrmOutMsg  �i���b�Z�[�W�o�̓t���O(�f�t�H���g:True){True:�o�͂���,False:�o�͂��Ȃ�}�j
'�߂�l     :���� True�A�G���[ False
'�@�@�@     :lPrmCnt     �i���������j
'===============================================================================
Private Function Fun_Ora_Execute(sPrmSql As String, Optional lPrmCnt As Long = 0, _
                            Optional bPrmOutMsg As Boolean = True) As Boolean

    Dim lSv_Mouse As Long     '�}�E�X�|�C���^�[�i�[
    Dim strErrMsg As String   '�G���[���̃��b�Z�[�W
    
    Fun_Ora_Execute = False
    
    '�}�E�X�|�C���^��Ԃ̕ۑ�
    lSv_Mouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrSub

    lPrmCnt = 0     '�X�V������������

    '�r�p�k�̔��s
    sPrmSql = Fun_Empty_To_Null(sPrmSql)   'SQL�����́u,'',�v��u,,�v���A�u,Null,�v�ɒu�������܂�
    lPrmCnt = OraDB.DbExecuteSQL(sPrmSql)

    Fun_Ora_Execute = True
    Screen.MousePointer = lSv_Mouse     '�}�E�X�|�C���^��Ԃ̕���

Exit Function

'�c�a�֘A�G���[�i�������f�j
ErrSub:
    '�G���[���b�Z�[�W����
    strErrMsg = ""
    strErrMsg = strErrMsg & "ErrorCode:" & Err.Number
    strErrMsg = strErrMsg & "  Message:" & Err.Description
    '�G���[���b�Z�[�W�o�́i�w�莞�̂ݏo�́j
    If bPrmOutMsg Then
        MsgBox "�G���[���������܂���" & vbCrLf & strErrMsg, vbCritical
    End If
    Screen.MousePointer = lSv_Mouse     '�}�E�X�|�C���^��Ԃ̕���
End Function

'===============================================================================
'�֐���     :�g�����U�N�V�����J�n���� Fun_Ora_BeginTransaction
'�X�V
'-------------------------------------------------------------------------------
'       ���t        ��      �S����      �R�����g
'�쐬   2001/08/01  1.00    JCE
'�X�V
'-------------------------------------------------------------------------------
'�����@     :bPrmOutMsg  �i���b�Z�[�W�o�̓t���O(�f�t�H���g:True){True:�o�͂���,False:�o�͂��Ȃ�}�j
'
'�߂�l     :���� True�A�G���[ False
'===============================================================================
Private Function Fun_Ora_BeginTransaction(Optional bPrmOutMsg As Boolean = True) As Boolean

    Dim lSv_Mouse As Long     '�}�E�X�|�C���^�[�i�[
    Dim strErrMsg As String   '�G���[���̃��b�Z�[�W

    Fun_Ora_BeginTransaction = False
    
    '�}�E�X�|�C���^��Ԃ̕ۑ�
    lSv_Mouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrSub

    '�r�p�k�̔��s
    Call OraSess.DbBeginTrans

    Fun_Ora_BeginTransaction = True
    Screen.MousePointer = lSv_Mouse     '�}�E�X�|�C���^��Ԃ̕���

Exit Function

'�c�a�֘A�G���[�i�������f�j
ErrSub:
    '�G���[���b�Z�[�W����
    strErrMsg = ""
    strErrMsg = strErrMsg & "ErrorCode:" & Err.Number
    strErrMsg = strErrMsg & "  Message:" & Err.Description
    '�G���[���b�Z�[�W�o�́i�w�莞�̂ݏo�́j
    If bPrmOutMsg Then
        MsgBox "�G���[���������܂���" & vbCrLf & strErrMsg, vbCritical
    End If
    Screen.MousePointer = lSv_Mouse     '�}�E�X�|�C���^��Ԃ̕���
End Function

'===============================================================================
'�֐���     :�g�����U�N�V��������I������ Fun_Ora_Commit
'�@�\       :lPrmConNo�̃g�����U�N�V�������R�~�b�g���܂��B
'-------------------------------------------------------------------------------
'       ���t        ��      �S����      �R�����g
'�쐬   2001/08/01  1.00    JCE
'�X�V
'-------------------------------------------------------------------------------
'�����@     :bPrmOutMsg  �i���b�Z�[�W�o�̓t���O(�f�t�H���g:True){True:�o�͂���,False:�o�͂��Ȃ�}�j
'
'�߂�l     :���� True�A�G���[ False
'===============================================================================
Private Function Fun_Ora_Commit(Optional bPrmOutMsg As Boolean = True) As Boolean

    Dim lSv_Mouse As Long     '�}�E�X�|�C���^�[�i�[
    Dim strErrMsg As String   '�G���[���̃��b�Z�[�W

    Fun_Ora_Commit = False
    
    '�}�E�X�|�C���^��Ԃ̕ۑ�
    lSv_Mouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrSub
    
    '�R�~�b�g�̔��s
    Call OraSess.DbCommitTrans

    Fun_Ora_Commit = True
    Screen.MousePointer = lSv_Mouse     '�}�E�X�|�C���^��Ԃ̕���

Exit Function

'�c�a�֘A�G���[�i�������f�j
ErrSub:
    '�G���[���b�Z�[�W����
    strErrMsg = ""
    strErrMsg = strErrMsg & "ErrorCode:" & Err.Number
    strErrMsg = strErrMsg & "  Message:" & Err.Description
    '�G���[���b�Z�[�W�o�́i�w�莞�̂ݏo�́j
    If bPrmOutMsg Then
        MsgBox "�G���[���������܂���" & vbCrLf & strErrMsg, vbCritical
    End If
    Screen.MousePointer = lSv_Mouse     '�}�E�X�|�C���^��Ԃ̕���
End Function

'===============================================================================
'�֐���     :�g�����U�N�V�����ُ�I������ Fun_Ora_Rollback
'�@�\       :lPrmConNo�̃g�����U�N�V���������[���o�b�N���܂��B
'-------------------------------------------------------------------------------
'       ���t        ��      �S����      �R�����g
'�쐬   2001/08/01  1.00    JCE
'�X�V
'-------------------------------------------------------------------------------
'�����@     :bPrmOutMsg  �i���b�Z�[�W�o�̓t���O(�f�t�H���g:True){True:�o�͂���,False:�o�͂��Ȃ�}�j
'
'�߂�l     :���� True�A�G���[ False
'===============================================================================
Private Function Fun_Ora_Rollback(Optional bPrmOutMsg As Boolean = True) As Boolean

    Dim lSv_Mouse As Long     '�}�E�X�|�C���^�[�i�[
    Dim strErrMsg As String   '�G���[���̃��b�Z�[�W
    
    Fun_Ora_Rollback = False
    
    '�}�E�X�|�C���^��Ԃ̕ۑ�
    lSv_Mouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrSub
    
    '���[���o�b�N�̔��s
    Call OraSess.DbRollback

    Fun_Ora_Rollback = True
    Screen.MousePointer = lSv_Mouse     '�}�E�X�|�C���^��Ԃ̕���

Exit Function

'�c�a�֘A�G���[�i�������f�j
ErrSub:
    '�G���[���b�Z�[�W����
    strErrMsg = ""
    strErrMsg = strErrMsg & "ErrorCode:" & Err.Number
    strErrMsg = strErrMsg & "  Message:" & Err.Description
    '�G���[���b�Z�[�W�o�́i�w�莞�̂ݏo�́j
    If bPrmOutMsg Then
        MsgBox "�G���[���������܂���" & vbCrLf & strErrMsg, vbCritical
    End If
    Screen.MousePointer = lSv_Mouse     '�}�E�X�|�C���^��Ԃ̕���
End Function

'�֐���     :Null�u�� Fun_Empty_To_Null
'�@�\       :�J���}��؂�̊e���ڂŁA���l���ځA�������ڂ���ł�����̂��A
'            'null'�ɒu�������܂��B
Private Function Fun_Empty_To_Null(sPrmStr As String) As String
    Dim lCnt As Long
    Dim lWkPos As Long
    Dim sWkStr As String
    
    '�󕶎�
    Do Until InStr(sPrmStr, "''") = 0
        lWkPos = InStr(sPrmStr, "''")
        sWkStr = Left$(sPrmStr, lWkPos - 1) & "null" & Mid$(sPrmStr, lWkPos + 2)
        sPrmStr = sWkStr
    Loop
    
    '�󐔒l
    Do Until InStr(sPrmStr, ",,") = 0
        lWkPos = InStr(sPrmStr, ",,")
        sWkStr = Left$(sPrmStr, lWkPos - 1) & ",null," & Mid$(sPrmStr, lWkPos + 2)
        sPrmStr = sWkStr
    Loop

    Fun_Empty_To_Null = sPrmStr

End Function

'�R���s���[�^�����擾
Private Function GetComputerName() As String
    '�R���s���[�^�����擾
    Dim sBuffer As String * 255
    
    '�o�b�t�@�̃N���A
    sBuffer = Space(255)
    
    'API�R�[��
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        '�擾�ł����ꍇ
        GetComputerName = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        '�擾�ł��Ȃ������ꍇ
        GetComputerName = "ERROR"
    End If
    
End Function

'>>>>>>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -----------------START
'>>>>>>>>>> mdlCommon.bas��Ins_TBCMC001_New�֐���s_cmzclabel.bas�Ɉړ� ------
    ' @(f)
    '
    ' �@�\    : ���x�����s�pӼޭ��
    '
    ' �Ԃ�l  : True:���� False:���s
    '
    ' ������  : after:���H���R�[�h
    '           cyoku:���敪
    '
    ' �@�\����:�@����������┭�����A�����������x���𔭍s����
    '
    ' ���l    :
    '       �����F
    '           sProcCode   �H������
    '           sEtcPrKind  ���̑����َ��
    '           sStaffID    �v���S����
    '           sPrKey01    ���[���ް�1
    '           sSysdate    ������t
    '           sRegDate    �o�^���t�@���o�^���t��PK�ׁ̈A���̏����ŕ������o�^����ꍇ�A
    '                                   �Ăяo������1�b���炷���̐��䂪�K�v
    '       �g�p��۸��сF
    '           cmbc008     �ؽ�ٶ�۸ތ����i�グ
    '           cmbc030     ������������
    '           cmbc018     �ؒf�E����َw���Ɖ�
    '
    Public Function Ins_TBCMC001_New(sProcCode As String, sEtcPrKind As String, sStaffID As String, sPrKey01 As String, sSysDate As String) As Boolean
        Dim sSql      As String       'SQL���i�[
        Dim iRet    As Integer        '�f�[�^�ǉ���
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        Dim labelCls    As c_cmzcLabel      '�N���X���W���[���p�ϐ�
        Set labelCls = New c_cmzcLabel      '�N���X�I�u�W�F�N�g����
        '<<<<< -------------------------------------------------------END
                
    '�G���[�n���h��
    On Error GoTo ErrHand

        '�߂�l�ݒ�
        Ins_TBCMC001_New = False
            
        '���߭�����ݒ�
        gsCompName = GetCompName
        
        '�o�^�p��ذ�ݒ�
        sSql = ""
        sSql = sSql & "insert into tbcmc001(" & vbLf    ''
        sSql = sSql & "                 quedate                     " & vbLf    ''�L���[���t
        sSql = sSql & "                 ,reqkind                    " & vbLf    ''����v���敪
        sSql = sSql & "                 ,printkind                  " & vbLf    ''������
        sSql = sSql & "                 ,endflg                     " & vbLf    ''�����敪
        sSql = sSql & "                 ,status                     " & vbLf    ''�I���X�e�[�^�X
        sSql = sSql & "                 ,blockidumu                 " & vbLf    ''�u���b�NID�L���敪
        sSql = sSql & "                 ,proccode                   " & vbLf    ''�H���R�[�h
        sSql = sSql & "                 ,etcprkind                  " & vbLf    ''���̑����x�����
        sSql = sSql & "                 ,crynum                     " & vbLf    ''�����ԍ�
        sSql = sSql & "                 ,ingotpos                   " & vbLf    ''�������ʒu
        sSql = sSql & "                 ,smplno                     " & vbLf    ''�T���v��No
        sSql = sSql & "                 ,mtrlnum                    " & vbLf    ''�����ԍ�
        sSql = sSql & "                 ,smtrlnum                   " & vbLf    ''���������ԍ�
        sSql = sSql & "                 ,blockid                    " & vbLf    ''�u���b�NID
        sSql = sSql & "                 ,hinban                     " & vbLf    ''�i��
        sSql = sSql & "                 ,revnum                     " & vbLf    ''���i�ԍ�����ԍ�
        sSql = sSql & "                 ,factory                    " & vbLf    ''�H��
        sSql = sSql & "                 ,opecond                    " & vbLf    ''���Ə���
        sSql = sSql & "                 ,cryindrs                   " & vbLf    ''���������w��(Rs)
        sSql = sSql & "                 ,cryindoi                   " & vbLf    ''���������w��(Oi)
        sSql = sSql & "                 ,cryindb1                   " & vbLf    ''���������w��(B1)
        sSql = sSql & "                 ,cryindb2                   " & vbLf    ''���������w��(B2)
        sSql = sSql & "                 ,cryindb3                   " & vbLf    ''���������w��(B3)
        sSql = sSql & "                 ,cryindl1                   " & vbLf    ''���������w��(L1)
        sSql = sSql & "                 ,cryindl2                   " & vbLf    ''���������w��(L2)
        sSql = sSql & "                 ,cryindl3                   " & vbLf    ''���������w��(L3)
        sSql = sSql & "                 ,cryindl4                   " & vbLf    ''���������w��(L4)
        sSql = sSql & "                 ,cryindcs                   " & vbLf    ''���������w��(Cs)
        sSql = sSql & "                 ,cryindgd                   " & vbLf    ''���������w��(Gd)
        sSql = sSql & "                 ,cryindt                    " & vbLf    ''���������w��(T)
        sSql = sSql & "                 ,cryindep                   " & vbLf    ''���������w��(EPD)
        sSql = sSql & "                 ,staffid                    " & vbLf    ''�v���S����
        sSql = sSql & "                 ,machine                    " & vbLf    ''�v���}�V����
        sSql = sSql & "                 ,regdate                    " & vbLf    ''�o�^���t
        sSql = sSql & "                 ,upddate                    " & vbLf    ''�X�V���t
        sSql = sSql & "                 ,prkey01                     " & vbLf    ''���[�L�[�f�[�^�P
        sSql = sSql & "     )                                       " & vbLf
        sSql = sSql & "values(                                      " & vbLf
        sSql = sSql & "                 to_date('" & sSysDate & "','yyyy/mm/dd hh24:mi:ss')       " & vbLf    ''�L���[���t
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrQueDate = Format(sSysDate, "yyyymmddhhmmss")  '�L���[���t���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'0'                        " & vbLf    ''����v���敪
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrReqKind = "0"                                 '����v���敪���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'1'                        " & vbLf    ''������
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrPrintKind = "1"                               '�����ނ��v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'0'                        " & vbLf    ''�����敪
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrEndFlg = "0"                                  '�����敪���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'0'                        " & vbLf    ''�I���X�e�[�^�X
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrStatus = "0"                                  '�I���X�e�[�^�X���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'0'                        " & vbLf    ''�u���b�NID�L���敪
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrBlockIdUmu = "0"                              '�u���b�NID�L���敪���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'" & sProcCode & "'        " & vbLf    ''�H���R�[�h
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrProcCode = sProcCode                          '�H���R�[�h���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'" & sEtcPrKind & "'       " & vbLf    ''���̑����x�����
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrEtcPrKind = sEtcPrKind                        '���̑����x����ނ��v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''�����ԍ�
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryNum = Null                                 '�����ԍ����v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''�������ʒu
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.NumIngotPos = Null                               '�������ʒu���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''�T���v��No
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.NumSmplNo = Null                                 '�T���v�������v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''�����ԍ�
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrMtrlNum = Null                                '�����ԍ����v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''���������ԍ�
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrSmtrlNum = Null                               '���������ԍ����v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''�u���b�NID
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrBlockId = Null                                '�u���b�NID���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''�i��
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrHinban = Null                                 '�i�Ԃ��v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''���i�ԍ�����ԍ�
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.NumRevNum = Null                                 '���i�ԍ������ԍ����v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''�H��
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrFactory = Null                                '�H����v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''���Ə���
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrOpeCond = Null                                '���Ə������v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(Rs)
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndRs = Null                               '���������w��(Rs)���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(Oi)
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndOi = Null                               '���������w��(Oi)���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(B1)
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndB1 = Null                               '���������w��(B1)���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(B2)
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndB2 = Null                               '���������w��(B2)���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(B3)
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndB3 = Null                               '���������w��(B3)���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(L1)
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndL1 = Null                               '���������w��(L1)���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(L2)
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndL2 = Null                               '���������w��(L2)���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(L3)
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndL3 = Null                               '���������w��(L3)���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(L4)
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndL4 = Null                               '���������w��(L4)���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(Cs)
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndCs = Null                               '���������w��(Cs)���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(Gd)
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndGD = Null                               '���������w��(GD)���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(T)
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndT = Null                                '���������w��(T)���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''���������w��(Epd)
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndEP = Null                               '���������w��(EPD)���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'" & sStaffID & "'         " & vbLf    ''�v���S���Җ�
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrStaffId = sStaffID                            '�v���S����ID���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'" & gsCompName & "'       " & vbLf    ''
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrMachine = gsCompName                          '�v���}�V�������v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,SYSDATE                    " & vbLf    ''�o�^���t
        
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        If Not GetSysdate() Then GoTo proc_exit  '�V�X�e�����t���擾
        
        labelCls.StrRegDate = Format(gsSysdate, "yyyymmddhhmmss") '�o�^���t���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,SYSDATE                    " & vbLf    ''�X�V���t
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.StrUpdDate = Format(gsSysdate, "yyyymmddhhmmss") '�X�V���t���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'" & sPrKey01 & "'         " & vbLf    ''���[�L�[�f�[�^�P
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        labelCls.strPrKey01 = sPrKey01                            '���[�L�[�f�[�^1���v���p�e�B�ɃZ�b�g
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                             )               " & vbLf

        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        '�V���v���Z�X�̃`�F�b�N
        labelCls.Label_Process_Check
        
        '�V�v���Z�X�̏ꍇ
        If labelCls.ProcKubun = True Then
            'KODZ6�ɓo�^
            If labelCls.Label_Facade = False Then
                Call MsgOut(100, sSql, ERR_DISP_LOG, "KODZ6")
                GoTo proc_exit
            End If
        Else
            '���v���Z�X�̏ꍇ������̏��� (TBCMC001�ɓo�^)
            iRet = SqlExec2(sSql)
            If iRet < 0 Then
                Call MsgOut(100, sSql, ERR_DISP_LOG, "TBCMC001")
                GoTo proc_exit
            End If
        End If
        '<<<<< -------------------------------------------------------END
        
        '���s
'        iRet = SqlExec2(sSql)
'
'        If iRet < 0 Then
'            Call MsgOut(100, sSql, ERR_DISP_LOG, "TBCMC001")
'            Exit Function
'        End If
        
        '�߂�l�ݒ�
        Ins_TBCMC001_New = True
        
'>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
proc_exit:
        Set labelCls = Nothing              '�N���X�I�u�W�F�N�g���
        Exit Function
'<<<<< -------------------------------------------------------END
           
    '�G���[��
ErrHand:
        ''�װ
        Call MsgOut(100, "", ERR_DISP_LOG, "TBCMC001")
        '>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -------START
        Resume proc_exit
        '<<<<< -------------------------------------------------------END
    End Function
'>>>>>>>>>> ���x��������������Ή� 2008/11/07 SETsw kakeida -----------------END

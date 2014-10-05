Attribute VB_Name = "s_cmbc036_SQL"
    Option Explicit

' �����ύX�w��

' �u���b�N���
Public Type typ_BlkInf3
    BLOCKID     As String * 12          ' �u���b�NID
    LENGTH      As Integer              ' ����
    REALLEN     As Integer              ' ������
    NOWPROC     As String * 5           ' ���ݍH��
    DELFLG      As String * 1           ' �폜�敪
    COF         As type_Coefficient     ' �ΐ͌W���v�Z
    '--- �쑺�ǉ��FDisp()�ł͏����l�Ƃ��ău���b�N�̂O�`��������ݒ�
    TOPPOS              As Integer          ' �u���b�N�̍ŏ��̃E�F�n�ʒu
    BOTPOS              As Integer          ' �u���b�N�̍Ō�̃E�F�n�ʒu
End Type

' �����E�F�n�[���
Public Type typ_LackWaf
    BLOCKID             As String * 12      ' �u���b�NID
    WAFERNO             As Integer          ' �E�F�n�[�A��
    WAFERTO             As Integer          ' �E�F�n�[�A��(to)
    TOP_POS             As Double           ' �E�F�n�[�J�n�ʒu'2002/02/27 S.Sano
    TAIL_POS            As Double           ' �E�F�n�[�I���ʒu'2002/02/27 S.Sano
    REJCAT              As String * 1       ' �������R
    ALLSCRAP            As String * 1       ' �S���X�N���b�v
End Type

'2002/09/11 ADD hitec)N.MATSUMOTO Start
Private HSXCTCEN        As Double           ' �i�r�w�����ʌX�c���S
Private HSXCYCEN        As Double           ' �i�r�w�����ʌX�����S
'WF�����v�Z�p�̃p�����[�^
Private SEEDDEG         As Integer          ' SEED�X��
Private Loss0           As Integer          ' �X����0�x�̂Ƃ��̌X�����X
Private Loss4           As Integer          ' �X����4�x�̂Ƃ��̌X�����X
Private Mlt4            As Double           ' �X����4�x�̎��̌W��
Private Pitch           As Double           ' ���C���\�[���C�����[���s�b�`

'�u���b�N�Ǘ�
Public Type typ_cmkc001f_Block
    'E040 �u���b�N�Ǘ�
    INGOTPOS            As Integer          ' �������J�n�ʒu
    LENGTH              As Integer          ' ����
    REALLEN             As Integer          ' ������
    KRPROCCD            As String * 5       ' ���݊Ǘ��H��
    NOWPROC             As String * 5       ' ���ݍH��
    LPKRPROCCD          As String * 5       ' �ŏI�ʉߊǗ��H��
    LASTPASS            As String * 5       ' �ŏI�ʉߍH��
    DELCLS              As String * 1       ' �폜�敪
    RSTATCLS            As String * 1       ' ������ԋ敪
    LSTATCLS            As String * 1       ' �ŏI��ԋ敪 */
    'E037 �������Ǘ�
    SEED                As String           'SEED
End Type

'�d�l�擾�p
Public Type typ_cmkc001f_Disp
    '�i�ԊǗ�
    hinban              As String * 8       ' �i��
    INGOTPOS            As Integer          ' �������J�n�ʒu
    REVNUM              As Integer          ' ���i�ԍ������ԍ�
    factory             As String * 1       ' �H��
    opecond             As String * 1       ' ���Ə���
    LENGTH              As Integer          ' ����
    '���i�d�lSXL�f�[�^
    HSXD1CEN            As Double           ' �i�r�w���a�P���S
    HSXRMIN             As Double           ' �i�r�w���R����
    HSXRMAX             As Double           ' �i�r�w���R���
    HSXRMBNP            As Double           ' �i�r�w���R�ʓ����z
    HSXRHWYS            As String * 1       ' �i�r�w���R�ۏؕ��@�Q��
    HSXONMIN            As Double           ' �i�r�w�_�f�Z�x����
    HSXONMAX            As Double           ' �i�r�w�_�f�Z�x���
    HSXONMBP            As Double           ' �i�r�w�_�f�Z�x�ʓ����z
    HSXONHWS            As String * 1       ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
    HSXCNMIN            As Double           ' �i�r�w�Y�f�Z�x����
    HSXCNMAX            As Double           ' �i�r�w�Y�f�Z�x���
    HSXCNHWS            As String * 1       ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
    HSXTMMAX            As Double           ' �i�r�w�]�ʖ��x���         ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
    HSXBMnAN(1 To 3)    As Double           ' �i�r�w�a�l�cn ���ω���
    HSXBMnAX(1 To 3)    As Double           ' �i�r�w�a�l�cn ���Ϗ��
    HSXBMnHS(1 To 3)    As String * 1       ' �i�r�w�a�l�cn �ۏؕ��@�Q��
    HSXOFnAX(1 To 4)    As Double           ' �i�r�w�n�r�en���Ϗ��
    HSXOFnMX(1 To 4)    As Double           ' �i�r�w�n�r�en���
    HSXOFnHS(1 To 4)    As String * 1       ' �i�r�w�n�r�en �ۏؕ��@�Q��
    HSXDENMX            As Integer          ' �i�r�w�c�������
    HSXDENMN            As Integer          ' �i�r�w�c��������
    HSXDENHS            As String * 1       ' �i�r�w�c�����ۏؕ��@�Q��
    HSXDVDMX            As Integer          ' �i�r�w�c�u�c�Q���
    HSXDVDMN            As Integer          ' �i�r�w�c�u�c�Q����
    HSXDVDHS            As String * 1       ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
    HSXLDLMX            As Integer          ' �i�r�w�k�^�c�k���
    HSXLDLMN            As Integer          ' �i�r�w�k�^�c�k����
    HSXLDLHS            As String * 1       ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
    HSXLTMIN            As Integer          ' �i�r�w�k�^�C������
    HSXLTMAX            As Integer          ' �i�r�w�k�^�C�����
    HSXLTHWS            As String * 1       ' �i�r�w�k�^�C���ۏؕ��@�Q��
    HSXDPDIR            As String * 2       ' �i�r�w�a�ʒu����
    HSXDPDRC            As String * 1       ' �i�r�w�a�ʒu����
    HSXDWMIN            As Double           ' �i�r�w�a�Љ���
    HSXDWMAX            As Double           ' �i�r�w�a�Џ��
    HSXDDMIN            As Double           ' �i�r�w�a�[����
    HSXDDMAX            As Double           ' �i�r�w�a�[���
    HSXD1MIN            As Double           ' �i�r�w���a�P����
    HSXD1MAX            As Double           ' �i�r�w���a�P���
    HSXCTCEN            As Double           ' �i�r�w�����ʌX�c���S
    HSXCYCEN            As Double           ' �i�r�w�����ʌX�����S
    EPDUP               As Integer          ' ���������Ǘ� EPD�@���
End Type

'=================================
'2003/02/28 ADD HITEC)okazaki start

Public Type type_DBDRV_Nukisi
    LOTID               As String * 12      ' �u���b�NID
    SXLID               As String * 13      ' SXLID
    MinMax              As Integer          ' 0:MIN 1:MAX
    BLOCKSEQ            As String * 3       ' �u���b�N���A��
    WFSTA               As String * 1       ' WF���
    hinban              As String * 8       ' �i��
    RTOP_POS            As Double           ' �_���u���b�N���ʒu
    RITOP_POS           As Double           ' �_���������ʒu
    SMPLEID             As String * 16      ' �����ʒu
    SHAFLAG             As String * 1       ' �T���v���t���O
    INDTM               As Date
    BASKETID            As String * 6
    SLOTNO              As Integer
    CURRWPCS            As Integer
    EXISTFLG            As String * 1
    TOP_POS             As Integer
    REJCAT              As String * 1
    TXID                As String * 6
    REGDATE             As Date
    SUMMITSENDFLAG      As String * 1
    SENDFLAG            As String * 1
    SENDDATE            As Date
    HREJCODE            As String * 4
    UPDPROC             As String * 5
    UPDDATE             As Date
    REVNUM              As Integer
    factory             As String * 1
    opecond             As String * 1
    KANKBN              As String * 1
    NREJCODE            As String * 6
    RINGOTPOS           As Double
End Type

Public Type type_DBDRV_LOTSXL
    LOTID               As String * 12      ' �u���b�NID
    SXLID               As String * 13      ' SXLID
    ''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
    INGOTPOS            As Integer          '�������ʒu
    ''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
End Type
'2003/02/28 Add HITEC)okazaki end

'2003/02/28 Hitec)okazaki add start
Public tExamine()       As type_DBDRV_Nukisi    '��ʕ\����
Public tGetExamine()    As type_DBDRV_Nukisi    '���s��
                                                '�E�F�n�[�Z���^�[���ɏ��e�[�u��
Public tKeturaku()      As typ_TBCMY012
'�\������
Public Const mSprChg_0  As Integer = 0          '�S��
Public Const mSprChg_1  As Integer = 1          '�Ǖi
Public Const mSprChg_2  As Integer = 2          '�T���v��
Public Const mSprChg_3  As Integer = 3          '�s��

Public tSXLID() As type_DBDRV_LOTSXL
'2003/02/28 Hitec)okazaki add end

'add start 2003/03/15 hitec)matsumoto ---------------

''WFϯ�ߊǗ�ð��ٍ\����
'Public Type typeWFmap
'    LOTID       As String       '�u���b�NID
'    BLOCKSEQ    As Integer      '�u���b�N���A��
'    INDTM       As Variant      '�v�e�Z���^�[���ɓ���
'    BASKETID    As String       '�o�X�P�b�gID
'    SLOTNO      As Integer      '�X���b�gNO
'    CURRWPCS    As Integer      '�v�e����
'    EXISTFLG    As String       '���݃t���O
'    TOP_POS     As Integer      '�u���b�N�̂s�n�o����̈ʒu
'    REJCAT      As String       '�������R
'    TXID        As String       '�g�����U�N�V����ID
'    REGDATE     As Variant      '�o�^���t
'    SUMMITSENDFLG   As String   'SUMIT���M�t���O
'    SENDFLG     As String       '���M�t���O
'    SENDDATE    As Variant      '���M����
'    WFSTA       As String       'WF���
'    HREJCODE    As String       '�s�Ǘ��R�R�[�h
'    UPDPROC     As String       '�X�V�H��
'    UPDDATE     As Variant      '�X�V����
'    SXLID       As String       'SXLID
'    hinban      As String       '�i��
'    REVNUM      As Integer      '���i�ԍ������ԍ�
'    factory     As String       '�H��
'    opecond     As String       '���Ə���
'    KANKBN      As String       '�����敪
'    SMPLEID     As String       '�����ʒu
'    NREJCODE    As String       '�����ԓ����R�R�[�h
'    SMPLEFLG    As String       '�T���v���t���O
'    RTOP_POS    As Double       '�_���u���b�N���ʒu
'    RITOP_POS   As Double       '�_���������ʒu
'End Type

'Public gtWFmap() As typeWFmap
Public bWfmapView As Boolean

''WFϯ�ߑΉ���ʃf�[�^�i�[�\����
'Public Type typeSprWFmap
'    LOTID       As Variant      '�u���b�NID
'    hinban      As Variant      '�i��
'    REVNUM      As Variant      '���i�ԍ������ԍ�
'    factory     As Variant      '�H��
'    opecond     As Variant      '���Ə���
'    HINUP       As tFullHinban  ' ��i��
'    HINDN       As tFullHinban  ' ���i��
'    blockp      As Variant      '�u���b�NP
'''''BLOCKP_T    As Variant      '�u���b�NP�i��j
'''''BLOCKP_B    As Variant      '�u���b�NP�i���j
'    KESSYOUP    As Variant      '����P
'''''KESSYOUP_T  As Variant      '����P�i��j
'''''KESSYOUP_B  As Variant      '����P�i���j
'    BLOCKSEQ    As Integer      '�}�b�v�ʒu
'''''BLOCKSEQ_T  As Integer      '�}�b�v�ʒu�i��j
'''''BLOCKSEQ_B  As Integer      '�}�b�v�ʒu�i���j
'    wfnum       As Integer      '�v�e����
'    WFSTA       As Variant      'WF���
'''''WFSTA_T     As Variant      'WF��ԁi��j
'''''WFSTA_B     As Variant      'WF��ԁi���j
'    REJCODE     As Integer      '�s�ǋ敪
'    SAMPLEID    As Variant      '�T���v��ID
'''''SAMPLEID_T  As Variant      '�T���v��ID�i��j
'''''SAMPLEID_B  As Variant      '�T���v��ID�i���j
'    WFSMP_Rs    As Variant      '�������ځiRs�j
'    WFSMP_Oi    As Variant      '�������ځiOi�j
'    WFSMP_B1    As Variant      '�������ځiB1�j
'    WFSMP_B2    As Variant      '�������ځiB2�j
'    WFSMP_B3    As Variant      '�������ځiB3�j
'    WFSMP_L1    As Variant      '�������ځiL1�j
'    WFSMP_L2    As Variant      '�������ځiL2�j
'    WFSMP_L3    As Variant      '�������ځiL3�j
'    WFSMP_L4    As Variant      '�������ځiL4�j
'    WFSMP_DS    As Variant      '�������ځiDS�j
'    WFSMP_DZ    As Variant      '�������ځiDZ�j
'    WFSMP_SP    As Variant      '�������ځiSP�j
'    WFSMP_D1    As Variant      '�������ځiD1�j
'    WFSMP_D2    As Variant      '�������ځiD2�j
'    WFSMP_D3    As Variant      '�������ځiD3�j
'    SHAFLAG     As Variant      '�T���v���t���O
'''''SHAFLAG_T   As Variant      '�T���v���t���O�i��j
'''''SHAFLAG_B   As Variant      '�T���v���t���O�i���j
'    ADD_FLG     As String       '0�F���������s�C1�F�ǉ������s
'End Type
'Public gtSprWfMap() As typeSprWFmap

''WF���
'Public Const gsWF_STA_0  As String = "0"      '�ʏ�
'Public Const gsWF_STA_1  As String = "1"      '���L
''Public Const gsWF_STA_2 As String = "2"      '�w���҂�
''Public Const gsWF_STA_3 As String = "3"      '�w��OK
'Public Const gsWF_STA_4  As String = "4"      '����
''Public Const gsWF_STA_5 As String = "5"      '����
'
''�T���v���t���O
'Public Const gsWF_SMPL_0 As String = "0"      '����
'Public Const gsWF_SMPL_1 As String = "1"      '�w���҂�
'Public Const gsWF_SMPL_2 As String = "2"      '�w��OK
'Public Const gsWF_SMPL_3 As String = "3"      '�w��NG
'Public Const gsWF_SMPL_4 As String = "4"      '����
'
''WF��ԁi��ʕ\���j
'Public Const gsWF_STA_NORMAL      As String = "�ʏ�"       '�ʏ�
'Public Const gsWF_STA_STA_K       As String = "����"       '����
'Public Const gsWF_STA_SIJI        As String = "�w���҂�"   '�w���҂�
'Public Const gsWF_STA_SIJI_OK     As String = "�w��OK"     '�w��OK
'Public Const gsWF_STA_SIJI_NG     As String = "�w��NG"     '�w��NG
'Public Const gsWF_STA_SIJI_KEKKA  As String = "����"       '����
'�T���v���t���O�i��ʕ\���j
'Public Const gsWF_SMPL_JOINT    As String = "���L"  '���L

'add end 2003/03/15 hitec)matsumoto ---------------------

'add 2003/03/25 hitec)matsumoto ��۰��ي֐��Ƃ��Ďg�������̂ŁAf_cmbc039_3.frm���ړ�----------------
Public SIngotP      As Integer              ' �C���S�b�g�㑤�ʒu
Public EIngotP      As Integer              ' �C���S�b�g�����ʒu
'add 2003/03/25 hitec)matsumoto ------------------------------
Public tmpSXLMng()  As typ_TBCME042

Public sWrpLOTID()      As String           '��ۯ�ID(Warp���ѕR�t���p)�@05/12/26 ooba
Public iWrpBLOCKSEQ()   As Integer          '��ۯ����A��(Warp���ѕR�t���p)�@05/12/26 ooba
Public pWafSmp_wk()     As typ_XSDCW        '����يǗ�(�����ް��ޔ�p)�@08/02/04 ooba

Public CngSmpID_UD()    As String



'2002/09/11 ADD hitec)N.MATSUMOTO End


'�T�v      :�����ύX�w���p �u���b�N�h�c���͎��c�a�h���C�o
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:sBlockID�@�@�@,I  ,String         �@,�u���b�NID
'      �@�@:bKounyu �@�@�@,I  ,Boolean        �@,�w���P�����t���O
'�@�@      :pCryInf �@�@�@,O  ,typ_TBCME037   �@,�������
'�@�@      :pHinDsn �@�@�@,O  ,typ_TBCME039   �@,�i�Ԑ݌v
'�@�@      :pHinMng �@�@�@,O  ,typ_TBCME041   �@,�i�ԊǗ�
'      �@�@:pSXLMng �@�@�@,O  ,typ_TBCME042   �@,SXL�Ǘ�
'      �@�@:pWafSmp �@�@�@,O  ,typ_XSDCW   �@   ,�V�T���v���Ǘ��iSXL�j
'�@�@      :pBlkInf �@�@�@,O  ,typ_BlkInf3    �@,�u���b�N���
'�@�@      :pHinSpec�@�@�@,O  ,typ_HinSpec    �@,���i�d�l
'�@�@      :pLackWaf�@�@�@,O  ,typ_LackWaf    �@,�����E�F�n�[���
'�@�@      :pBlkID  �@�@�@,O  ,String         �@,���o�P�ʃu���b�NID
'      �@�@:dNeraiRes �@�@,O  ,Double         �@,�˂炢�i�Ԃ̔��R����l�iP+�̔��f�p�j
'      �@�@:sErrMsg �@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
'����      :
'����      :2001/07/10 ���� �쐬
Public Function DBDRV_scmzc_fcmkc001k_Disp(ByVal sBlockId As String, bKounyu As Boolean, _
                                           pCryInf As typ_TBCME037, pHinDsn() As typ_TBCME039, _
                                           pHinMng() As typ_TBCME041, pSXLMng() As typ_TBCME042, _
                                           pWafSmp() As typ_XSDCW, pBlkInf() As typ_BlkInf3, _
                                           pHinSpec() As typ_HinSpec, pLackWaf() As typ_LackWaf, _
                                           pBlkID() As String, dNeraiRes As Double, sErrMsg As String) As FUNCTION_RETURN

    Dim tmpCryInf() As typ_TBCME037
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim sDbName     As String
    Dim sCryNum     As String
    Dim sHin        As String
    Dim sBLK        As String
    Dim dMenseki    As Double
    Dim dTopWght    As Double
    Dim dCharge     As Double
    Dim dMeas(4)    As Double
    Dim bFlag       As Boolean
    Dim recCnt      As Long
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim REJCAT      As String

    ''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
    Dim tmpHinMng() As typ_TBCME041     '�i�ԏ��X�V�p
'    Dim ltXSDCA()   As typ_XSDCA        'XSDCW�f�[�^�⊮�p�f�[�^
    Dim tmpWafSmp() As typ_XSDCW        'XSDCW�f�[�^�⊮�p�f�[�^
    ''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001k_SQL.bas -- Function DBDRV_scmzc_fcmkc001k_Disp"
    sErrMsg = ""

    '' �u���b�N�Ǘ��̎擾
    sDbName = "E040"
    sCryNum = Left(sBlockId, 9) & "000"
    sql = "select INGOTPOS, LENGTH, REALLEN, BLOCKID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, LSTATCLS"
    sql = sql & " from  TBCME040"
    sql = sql & " where CRYNUM   ='" & sCryNum & "'"
    sql = sql & "   and INGOTPOS>= 0"
    sql = sql & "   and LENGTH  >  0"
    sql = sql & " order by INGOTPOS"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    bFlag = False
    ReDim pBlkInf(recCnt)
    For i = 1 To recCnt
        With pBlkInf(i)
            .COF.TOPSMPLPOS = rs("INGOTPOS")
            .LENGTH = rs("LENGTH")
            .REALLEN = rs("REALLEN")
            .BLOCKID = rs("BLOCKID")
            .NOWPROC = rs("NOWPROC")
            .COF.BOTSMPLPOS = .COF.TOPSMPLPOS + .LENGTH
            .DELFLG = "0"
            .TOPPOS = 0
            .BOTPOS = .REALLEN
            If .BLOCKID = sBlockId Then
                '' �H���`�F�b�N
                If rs("LSTATCLS") <> "W" Then
                    sErrMsg = GetMsgStr("EPRC2")
                    DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
                bFlag = True
            End If
        End With
        rs.MoveNext
    Next i
    rs.Close

    '' �u���b�NID���݃`�F�b�N
    If bFlag = False Then
        sErrMsg = GetMsgStr("EBLK0")
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If


    '' �������̎擾(s_cmzcTBCME037_SQL.bas ���K�v)
    sDbName = "E037"
    sql = " where CRYNUM='" & sCryNum & "'"
    If DBDRV_GetTBCME037(tmpCryInf(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpCryInf) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pCryInf = tmpCryInf(1)


    '' �i�Ԑ݌v�̎擾(s_cmzcTBCME039_SQL.bas ���K�v)
    sDbName = "E039"
    '2004.09.08 Y.K �R�t���ύX
'    sql = " where substr(CRYNUM,1,7)='" & Left(sCryNum, 7) & "' and LENGTH>0 order by INGOTPOS"
    sql = " where substr(CRYNUM,1,9)='" & Left(sCryNum, 7) & "0" & Mid(sCryNum, 9, 1) & "' and LENGTH>0 order by INGOTPOS"
    If DBDRV_GetTBCME039(pHinDsn(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If


    '' �i�ԊǗ��̎擾(s_cmzcTBCME041_SQL.bas ���K�v)
    sDbName = "E041"
    sql = " where CRYNUM='" & sCryNum & "' order by INGOTPOS"
    If DBDRV_GetTBCME041(pHinMng(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(pHinMng) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If


    '' SXL�Ǘ��̎擾(s_cmzcTBCME042_SQL.bas ���K�v)
    sDbName = "E042"
    ''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
    sql = " where XTALCB='" & sCryNum & "' order by INPOSCB"
    ''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
    If DBDRV_GetTBCME042(pSXLMng(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(pSXLMng) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If


    '' WF�T���v���Ǘ��̎擾(s_cmzcTBCME044_SQL.bas ���K�v)
    '�@�@���V�T���v���Ǘ��ɕύX-------2003/09/18
    ' �������敪������K�v�L��
'   sDbName = "E044"
    sDbName = "XSDCW"
    sql = " where XTALCW='" & sCryNum & "'" _
        & "   and LIVKCW='0'" _
        & " order by INPOSCW"
'    If DBDRV_GetTBCME044(pWafSmp(), sql) = FUNCTION_RETURN_FAILURE Then
    'XSDCB�̏����ް��Ɏ擾�@08/02/04 ooba
    If DBDRV_GetXSDCW(sCryNum, pWafSmp()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' �֘A��ۯ����擾�@08/10/28 ooba
    If sKanrenFlg = "1" Then
        sDbName = "Y023"
        sql = "SELECT "
        sql = sql & "BLOCKID, "
        sql = sql & "PROCCAT "
        sql = sql & "FROM TBCMY023 "
        sql = sql & "WHERE CRYNUM LIKE '" & Mid(sBlockId, 1, 9) & "%' "
        sql = sql & "AND TRANCNT = ( "
        sql = sql & "    SELECT "
        sql = sql & "    MAX(TRANCNT) "
        sql = sql & "    FROM TBCMY023 "
        sql = sql & "    WHERE CRYNUM LIKE '" & Mid(sBlockId, 1, 9) & "%' "
        sql = sql & "    AND BLOCKID = '" & sBlockId & "' "
        sql = sql & ") "
        sql = sql & "ORDER BY BLOCKID "
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        recCnt = rs.RecordCount
        If recCnt <= 0 Then
            rs.Close
            sErrMsg = GetMsgStr("EGET2", sDbName)
            DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        '�֘A��ۯ��łȂ��ꍇ
        If rs.Fields("PROCCAT") = "D" Then
            ReDim pBlkID(1)
            pBlkID(1) = sBlockId
        Else
            ReDim pBlkID(recCnt)
            '��ۯ�ID���
            For i = 1 To recCnt
                pBlkID(i) = rs("BLOCKID")
                rs.MoveNext
            Next i
        End If
        rs.Close
    '�֘A��ۯ��łȂ��ꍇ
    Else
        ReDim pBlkID(1)
        pBlkID(1) = sBlockId
    End If
    
'''    '' �u���b�N�V�K���̎擾
'''    sDbName = "Y001"
'''    sql = "select SBLOCKID from TBCMY001 where BLOCKID='" & sBlockID & "'"
'''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''    If rs.RecordCount <= 0 Then
'''        rs.Close
'''        sErrMsg = GetMsgStr("EGET2", sDbName)
'''        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
'''        GoTo proc_exit
'''    End If
'''    sBLK = rs("SBLOCKID")
'''    rs.Close
'''
'''
'''    sql = "select BLOCKID from TBCMY001"
'''    sql = sql & " where SBLOCKID='" & sBLK & "'"
'''    sql = sql & " order by SBLOCKID, BLOCKORDER"
'''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''    recCnt = rs.RecordCount
'''    If recCnt <= 0 Then
'''        rs.Close
'''        sErrMsg = GetMsgStr("EGET2", sDbName)
'''        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
'''        GoTo proc_exit
'''    End If
'''
'''    ReDim pBlkID(recCnt)
'''    For i = 1 To recCnt
'''        pBlkID(i) = rs("BLOCKID")
'''        rs.MoveNext
'''    Next i
'''    rs.Close

    ''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
    ''�i�ԊǗ��̎擾(XSDCA,XSDCB:�w��u���b�N�̂�)
    sDbName = "E041update"
    If DBDRV_GetTBCME041_Clone(tmpHinMng(), sCryNum, pBlkID) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(pHinMng) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '�����S�̂̕i�ԃf�[�^�Ƀu���b�N�w��̃f�[�^����������
    Call s_cmbc036_2_F_SynHinban(pHinMng, tmpHinMng)

    '' XSDCW�⊮�p�f�[�^��XSDCA����쐬��
    '' �擾����XSDCA�f�[�^������XSDCW�̎擾�f�[�^��⊮����
    sDbName = "XSDCA"
    ReDim tmpWafSmp(UBound(pWafSmp))
    ReDim pWafSmp_wk(UBound(pWafSmp))       '08/02/04 ooba
    For i = 0 To UBound(pWafSmp)
        tmpWafSmp(i) = pWafSmp(i)
        pWafSmp_wk(i) = pWafSmp(i)          '�����ް��ޔ��@08/02/04 ooba
    Next i

    '���ǉ� 2006/03/20 ��Q�Ή� SMP�ΐ� WF���ύX����Ă��Ȃ��ꍇ�A�������\������Ȃ���Q�ɑΉ�
    ReDim tSXLID(0)
    '���ǉ� 2006/03/20 ��Q�Ή� SMP�ΐ� WF���ύX����Ă��Ȃ��ꍇ�A�������\������Ȃ���Q�ɑΉ�

    If DBDRV_GetXSDCWUpdate(tmpWafSmp(), sCryNum, pBlkID()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    '�⊮�Ȃ��̏ꍇ�AtmpWafSmp�͏����������̂ŁA����pWafSmp�����̂܂܎g�p����
    If UBound(tmpWafSmp) <> 0 Then
        ReDim pWafSmp(UBound(tmpWafSmp))
        For i = 0 To UBound(pWafSmp)
            pWafSmp(i) = tmpWafSmp(i)
        Next i
    End If

    ''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�

    '' ���グ�I�����т̎擾
    '2001/07/23 S.Sano Start
    If Not bKounyu Then
    '2001/07/23 S.Sano End
        sDbName = "H004"
        sql = "select (DM1+DM2+DM3)/3.0 as DM, WGHTTOP, CHARGE from TBCMH004 where CRYNUM='" & sCryNum & "'"
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            rs.Close
            sErrMsg = GetMsgStr("EGET2", sDbName)
            DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        dMenseki = AreaOfCircle(rs("DM"))
        dTopWght = rs("WGHTTOP")
        dCharge = rs("CHARGE")
        rs.Close
    '2001/07/23 S.Sano Start
    End If
    '2001/07/23 S.Sano End


    '' ������R���т̎擾
    sDbName = "J002"
    recCnt = UBound(pBlkInf)
    For i = 1 To recCnt
        With pBlkInf(i)
            .COF.DUNMENSEKI = dMenseki      ' �f�ʐ�
            .COF.CHARGEWEIGHT = dCharge     ' �`���[�W��
            .COF.TOPWEIGHT = dTopWght       ' �g�b�v�d��

            '' �g�b�v�����R�����l�̎擾
            sql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 from TBCMJ002"
            sql = sql & " where CRYNUM  ='" & sCryNum & "'"
            sql = sql & "   and POSITION= " & .COF.TOPSMPLPOS & " and SMPKBN='T'"
            sql = sql & "   and TRANCNT = ANY(select MAX(TRANCNT) from TBCMJ002 where CRYNUM='" & sCryNum & "' and POSITION=" & .COF.TOPSMPLPOS & " and SMPKBN='T')"
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount > 0 Then
                dMeas(0) = rs("MEAS1")
                dMeas(1) = rs("MEAS2")
                dMeas(2) = rs("MEAS3")
                dMeas(3) = rs("MEAS4")
                dMeas(4) = rs("MEAS5")
                .COF.TOPRES = JudgCenter(dMeas())
            Else
                .COF.TOPRES = -9999
            End If
            rs.Close

            '' �{�g�������R�����l�̎擾
            sql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 from TBCMJ002"
            sql = sql & " where CRYNUM  ='" & sCryNum & "'"
            sql = sql & "   and POSITION= " & .COF.BOTSMPLPOS & " and SMPKBN='B'"
            sql = sql & "   and TRANCNT = ANY(select MAX(TRANCNT) from TBCMJ002 where CRYNUM='" & sCryNum & "' and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='B')"
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                rs.Close
                sql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 from TBCMJ002"
                sql = sql & " where CRYNUM  ='" & sCryNum & "'"
                sql = sql & "   and POSITION= " & .COF.BOTSMPLPOS & " and SMPKBN='T'"
                sql = sql & "   and TRANCNT = ANY(select MAX(TRANCNT) from TBCMJ002 where CRYNUM='" & sCryNum & "' and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='T')"
                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            End If
            If rs.RecordCount > 0 Then
                dMeas(0) = rs("MEAS1")
                dMeas(1) = rs("MEAS2")
                dMeas(2) = rs("MEAS3")
                dMeas(3) = rs("MEAS4")
                dMeas(4) = rs("MEAS5")
                .COF.BOTRES = JudgCenter(dMeas())
            Else
                .COF.BOTRES = -9999
            End If
            rs.Close
        End With
    Next i


    '' ���i�d�l�̎擾
    sDbName = "VE004"
    recCnt = UBound(pHinMng)
    ReDim pHinSpec(recCnt)
    k = 0
    For i = 1 To recCnt
        With pHinMng(i)
            sHin = RTrim$(.hinban)
            If sHin <> "" And sHin <> "G" And sHin <> "Z" Then
                For j = 1 To k
                    If pHinSpec(j).hin.hinban = .hinban Then
                        pHinSpec(j).LENGTH = pHinSpec(j).LENGTH + .LENGTH
                        Exit For
                    End If
                Next j
                If j > k Then
                    k = k + 1
                    pHinSpec(k).INGOTPOS = .INGOTPOS
                    pHinSpec(k).hin.hinban = .hinban
                    pHinSpec(k).hin.mnorevno = .REVNUM
                    pHinSpec(k).hin.factory = .factory
                    pHinSpec(k).hin.opecond = .opecond
                    pHinSpec(k).LENGTH = .LENGTH

                    ''�c���_�f�d�l�`�F�b�N�@03/12/11 ooba START ==============================>
                    iChkAoi = ChkAoiSiyou(pHinSpec(k).hin)
                    If iChkAoi < 0 Then
                        sErrMsg = "�c���_�f(AOi)�d�l�G���["
                        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                    ''�c���_�f�d�l�`�F�b�N�@03/12/11 ooba END ================================>

                    If DBDRV_scmzc_fcmkc001g_GetSpec(pHinSpec(k)) = FUNCTION_RETURN_FAILURE Then
                        sErrMsg = GetMsgStr("EGET") & sDbName
                        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                End If
            End If
        End With
    Next i
    ReDim Preserve pHinSpec(k)



ReDim pLackWaf(0)   '2003/05/05 hitec)okazaki
    '' �����E�F�n�[���̎擾
#If False Then
    sDbName = "VW002"
    sql = "select distinct BLOCKID, WAFERNO, TOP_POS, TAIL_POS"
    sql = sql & " from VECMW002 where CRYNUM='" & sCryNum & "' order by BLOCKID, WAFERNO"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    ReDim pLackWaf(recCnt)
    For i = 1 To recCnt
        With pLackWaf(i)
            .BLOCKID = rs("BLOCKID")    ' �u���b�NID
            .WAFERNO = rs("WAFERNO")    ' �E�F�n�[�A��
            .TOP_POS = rs("TOP_POS")    ' �E�F�n�[�J�n�ʒu
            .TAIL_POS = rs("TAIL_POS")  ' �E�F�n�[�I���ʒu
        End With
        rs.MoveNext
    Next i
    rs.Close
#Else
''    sDbName = "VW004"
''    sql = "select distinct LOTID as BLOCKID, REJCAT, REJWFFROM as WAFERNO, REJWFTO as WAFERTO, REJFROM as TOP_POS, REJTO as TAIL_POS, ALLSCRAP"
''    sql = sql & " from VECMW004"
''    sql = sql & " where (LOTID like '" & Left$(sCryNum, 9) & "%') and (REJCAT<>'C')"
''    sql = sql & " order by LOTID, WAFERNO "

    '�ޭ��Q�ƒ�~�@06/02/06 ooba START ====================================================>
    sDbName = "Y012"
    sql = "select distinct LOTID as BLOCKID, REJCAT, REJWFFROM as WAFERNO, REJWFTO as WAFERTO, REJFROM as TOP_POS, REJTO as TAIL_POS, ALLSCRAP from "
    sql = sql & "("
    sql = sql & "select "
    sql = sql & "  C.CRYNUM,"
    sql = sql & "  XXX.LOTID,"
    sql = sql & "  REJCAT,"
    sql = sql & "  ALLSCRAP,"
    sql = sql & "  case when (XXX.REJFROM<=B.WFFROM) then 0 else XXX.REJFROM end as REJFROM,"
    sql = sql & "  case when (XXX.REJTO>=B.WFTO) then C.LENGTH else XXX.REJTO end as REJTO,"
    sql = sql & "  REJWFFROM,"
    sql = sql & "  REJWFTO"
    sql = sql & " from "
    sql = sql & "("
    sql = sql & "select "
    sql = sql & "    LOTID,"
    sql = sql & "    REJCAT,"
    sql = sql & "    ALLSCRAP,"
    sql = sql & "    0 as REJFROM,"
    sql = sql & "    LENGTH as REJTO,"
    sql = sql & "    0 as REJWFFROM,"
    sql = sql & "    0 as REJWFTO"
    sql = sql & "  from TBCMY012 A,"
    sql = sql & "    TBCME040 B"
    sql = sql & "  where (A.LOTID=B.BLOCKID)"
    sql = sql & "    and (A.ALLSCRAP='Y')"
    sql = sql & " and a.lotid like '" & Left$(sCryNum, 9) & "%'"
    sql = sql & " union all"
    sql = sql & "  select distinct"
    sql = sql & "    LOTID,"
    sql = sql & "    REJCAT,"
    sql = sql & "    ALLSCRAP,"
    sql = sql & "    LENFROM,"
    sql = sql & "    LENTO,"
    sql = sql & "    0 as REJWFFROM,"
    sql = sql & "    0 as REJWFTO"
    sql = sql & "  from TBCMY012"
'    sql = sql & "  where (REJCAT='A') and (ALLSCRAP='N')"
    sql = sql & "  where (REJCAT in ('A','E')) and (ALLSCRAP='N')"      '��ۯ���Ԃł̈ꕔ���ʑΉ� 09/02/27 ooba
    sql = sql & " and lotid like '" & Left$(sCryNum, 9) & "%'"
    sql = sql & " union all"
    sql = sql & "  select distinct"
    sql = sql & "    A.LOTID,"
    sql = sql & "    A.REJCAT,"
    sql = sql & "    A.ALLSCRAP,"
    sql = sql & "    A.TOP_POS/10.0 as REJFROM,"
    sql = sql & "    A.TOP_POS/10.0 as REJTO,"
    sql = sql & "    A.BLOCKSEQ as REJWFFROM,"
    sql = sql & "    A.BLOCKSEQ as REJWFTO"
    sql = sql & "  from TBCMY012 A"
    sql = sql & "  where (A.REJCAT='B') and (ALLSCRAP='N')"
    sql = sql & " and a.lotid like '" & Left$(sCryNum, 9) & "%'"
    sql = sql & " union all"
    sql = sql & "  select distinct"
    sql = sql & "    A.LOTID,"
    sql = sql & "    A.REJCAT,"
    sql = sql & "    A.ALLSCRAP,"
    sql = sql & "    B.TOP_POS/10.0 as REJFROM,"
    sql = sql & "    C.TOP_POS/10.0 as REJTO,"
    sql = sql & "    A.BLOCKSEQ as REJWFFROM,"
    sql = sql & "    A.BLOCKSEQ+A.REJPCS-1 as REJWFTO"
    sql = sql & "  from TBCMY012 A,"
    sql = sql & "    TBCMY011 B,"
    sql = sql & "    TBCMY011 C"
    sql = sql & "  where (A.REJCAT='C')"
    sql = sql & "    and (A.LOTID=B.LOTID) and (A.BLOCKSEQ=B.BLOCKSEQ)"
    sql = sql & "    and (A.LOTID=C.LOTID) and (A.BLOCKSEQ+A.REJPCS-1=C.BLOCKSEQ)"
    sql = sql & " and a.lotid like '" & Left$(sCryNum, 9) & "%'"
    sql = sql & " order by LOTID,REJFROM"
    sql = sql & ") XXX,"
    sql = sql & "  (select LOTID, min(TOP_POS)/10.0 as WFFROM, max(TOP_POS)/10.0 as WFTO from TBCMY011 "
    sql = sql & " where lotid like '" & Left$(sCryNum, 9) & "%'"
    sql = sql & " group by LOTID) B,"
    sql = sql & "  TBCME040 C"
    sql = sql & " where (XXX.LOTID=B.LOTID)"
    sql = sql & "  and (XXX.LOTID=C.BLOCKID)"
    sql = sql & "  and (XXX.ALLSCRAP='N')"
    sql = sql & " union all"
    sql = sql & " select distinct"
    sql = sql & "  C.CRYNUM,"
    sql = sql & "  XXX.LOTID,"
    sql = sql & "  REJCAT,"
    sql = sql & "  ALLSCRAP,"
    sql = sql & "  0 as REJFROM,"
    sql = sql & "  C.LENGTH as REJTO,"
    sql = sql & "  REJWFFROM,"
    sql = sql & "  REJWFTO"
    sql = sql & " from "
    sql = sql & "("
    sql = sql & "select "
    sql = sql & "    LOTID,"
    sql = sql & "    REJCAT,"
    sql = sql & "    ALLSCRAP,"
    sql = sql & "    0 as REJFROM,"
    sql = sql & "    LENGTH as REJTO,"
    sql = sql & "    0 as REJWFFROM,"
    sql = sql & "    0 as REJWFTO"
    sql = sql & "  from TBCMY012 A,"
    sql = sql & "    TBCME040 B"
    sql = sql & "  where (A.LOTID=B.BLOCKID)"
    sql = sql & "    and (A.ALLSCRAP='Y')"
    sql = sql & " and a.lotid like '" & Left$(sCryNum, 9) & "%'"
    sql = sql & " union all"
    sql = sql & "  select distinct"
    sql = sql & "    LOTID,"
    sql = sql & "    REJCAT,"
    sql = sql & "    ALLSCRAP,"
    sql = sql & "    LENFROM,"
    sql = sql & "    LENTO,"
    sql = sql & "    0 as REJWFFROM,"
    sql = sql & "    0 as REJWFTO"
    sql = sql & "  from TBCMY012"
'    sql = sql & "  where (REJCAT='A') and (ALLSCRAP='N')"
    sql = sql & "  where (REJCAT in ('A','E')) and (ALLSCRAP='N')"      '��ۯ���Ԃł̈ꕔ���ʑΉ� 09/02/27 ooba
    sql = sql & " and lotid like '" & Left$(sCryNum, 9) & "%'"
    sql = sql & " union all"
    sql = sql & "  select distinct"
    sql = sql & "    A.LOTID,"
    sql = sql & "    A.REJCAT,"
    sql = sql & "    A.ALLSCRAP,"
    sql = sql & "    A.TOP_POS/10.0 as REJFROM,"
    sql = sql & "    A.TOP_POS/10.0 as REJTO,"
    sql = sql & "    A.BLOCKSEQ as REJWFFROM,"
    sql = sql & "    A.BLOCKSEQ as REJWFTO"
    sql = sql & "  from TBCMY012 A"
    sql = sql & "  where (A.REJCAT='B') and (ALLSCRAP='N')"
    sql = sql & " and a.lotid like '" & Left$(sCryNum, 9) & "%'"
    sql = sql & " union all"
    sql = sql & "  select distinct"
    sql = sql & "    A.LOTID,"
    sql = sql & "    A.REJCAT,"
    sql = sql & "    A.ALLSCRAP,"
    sql = sql & "    B.TOP_POS/10.0 as REJFROM,"
    sql = sql & "    C.TOP_POS/10.0 as REJTO,"
    sql = sql & "    A.BLOCKSEQ as REJWFFROM,"
    sql = sql & "    A.BLOCKSEQ+A.REJPCS-1 as REJWFTO"
    sql = sql & "  from TBCMY012 A,"
    sql = sql & "    TBCMY011 B,"
    sql = sql & "    TBCMY011 C"
    sql = sql & "  where (A.REJCAT='C')"
    sql = sql & "    and (A.LOTID=B.LOTID) and (A.BLOCKSEQ=B.BLOCKSEQ)"
    sql = sql & "    and (A.LOTID=C.LOTID) and (A.BLOCKSEQ+A.REJPCS-1=C.BLOCKSEQ)"
    sql = sql & " and a.lotid like '" & Left$(sCryNum, 9) & "%'"
    sql = sql & " order by LOTID,REJFROM"
    sql = sql & ") XXX,"
    sql = sql & "  TBCME040 C"
    sql = sql & " where (XXX.LOTID=C.BLOCKID)"
    sql = sql & "  and (XXX.ALLSCRAP='Y')"
    sql = sql & ")"
    sql = sql & " where (REJCAT<>'C')"
    sql = sql & " order by LOTID, WAFERNO "
    '�ޭ��Q�ƒ�~�@06/02/06 ooba END ======================================================>

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    ReDim pLackWaf(recCnt)
    k = 0
    For i = 1 To recCnt

'2002/08/22
'        If (REJCAT = rs("REJCAT")) _
'          And rs("ALLSCRAP") = "N" _
'          And (pLackWaf(k).BLOCKID = rs("BLOCKID")) _
'          And (pLackWaf(k).WAFERTO + 1 = rs("WAFERNO")) Then
        If (REJCAT = rs("REJCAT")) _
          And rs("ALLSCRAP") = "N" _
          And (pLackWaf(k).ALLSCRAP = rs("ALLSCRAP")) _
          And (pLackWaf(k).BLOCKID = rs("BLOCKID")) _
          And (pLackWaf(k).WAFERTO + 1 = rs("WAFERNO")) Then

            With pLackWaf(k)
                .WAFERTO = rs("WAFERTO")    ' �E�F�n�[�A��(to)
                .TAIL_POS = rs("TAIL_POS")  ' �E�F�n�[�I���ʒu
            End With
        Else
            k = k + 1
            With pLackWaf(k)
                .BLOCKID = rs("BLOCKID")    ' �u���b�NID
                .WAFERNO = rs("WAFERNO")    ' �E�F�n�[�A��
                .WAFERTO = rs("WAFERTO")    ' �E�F�n�[�A��(to)
                .TOP_POS = rs("TOP_POS")    ' �E�F�n�[�J�n�ʒu
                .TAIL_POS = rs("TAIL_POS")  ' �E�F�n�[�I���ʒu
                .ALLSCRAP = rs("ALLSCRAP")  ' �S���X�N���b�v
                .REJCAT = rs("REJCAT")      ' �������R
            End With
        End If
        REJCAT = rs("REJCAT")
        rs.MoveNext
    Next i
    rs.Close
    ReDim Preserve pLackWaf(k)
#End If


    '' �˂炢�i�Ԃ̔��R����l���擾
    sql = "select HSXRMAX"
    sql = sql & " from TBCME037 E37, TBCME018 E18"
    sql = sql & " where (E37.CRYNUM  ='" & Left$(sBlockId, 9) & "000')"
    sql = sql & "   and (E37.RPHINBAN=E18.HINBAN)  and (E37.RPREVNUM=E18.MNOREVNO)"
    sql = sql & "   and (E37.RPFACT  =E18.FACTORY) and (E37.RPOPCOND=E18.OPECOND)"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount > 0 Then
        dNeraiRes = rs("HSXRMAX")
    Else
        dNeraiRes = 0#      '�����܂ł͂��Ȃ��͂�
    End If
    rs.Close


    DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("EGET2", sDbName)
    DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :�֘A��ۯ����擾
'���Ұ��@�@:�ϐ���         ,IO ,�^               ,����
'      �@�@:sBlockID       ,I  ,String         �@,��ۯ�ID
'      �@�@:tKanrenDisp()  ,I  ,typ_KanrenDisp   ,�֘A��ۯ��ꗗ
'      �@�@:�߂�l         ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :
'����      :08/01/23 ooba
Public Function DBDRV_scmzc_fcmkc001k_Disp2(sBlockId As String, _
                                            tKanrenDisp() As typ_KanrenDisp) As FUNCTION_RETURN

    Dim i, j        As Integer
    Dim iBlkCnt     As Integer      '��ۯ���
    Dim iHinCnt     As Integer      '�i�Ԑ�
    Dim skanblock() As String       '�֘A��ۯ��@08/10/28 ooba
    Dim rs          As OraDynaset
    Dim sql         As String
    
    '�װ����ׂ̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function DBDRV_scmzc_fcmkc001k_Disp2"
    
    DBDRV_scmzc_fcmkc001k_Disp2 = FUNCTION_RETURN_FAILURE
    
    '�֘A��ۯ��R�ؕR�tð��ق��֘A��ۯ��擾�@08/10/28 ooba
    sql = "SELECT "
    sql = sql & "BLOCKID, "
    sql = sql & "PROCCAT "
    sql = sql & "FROM TBCMY023 "
    sql = sql & "WHERE CRYNUM LIKE '" & Mid(sBlockId, 1, 9) & "%' "
    sql = sql & "AND TRANCNT = ( "
    sql = sql & "    SELECT "
    sql = sql & "    MAX(TRANCNT) "
    sql = sql & "    FROM TBCMY023 "
    sql = sql & "    WHERE CRYNUM LIKE '" & Mid(sBlockId, 1, 9) & "%' "
    sql = sql & "    AND BLOCKID = '" & sBlockId & "' "
    sql = sql & ") "
    sql = sql & "ORDER BY BLOCKID "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    j = rs.RecordCount
    
    '�ް�����
    If j <= 0 Then
        GoTo proc_exit
    End If
    
    '�֘A��ۯ��łȂ��ꍇ
    If rs.Fields("PROCCAT") = "D" Then
        GoTo proc_exit
    End If
        
    ReDim skanblock(j)
    '��ۯ�ID���
    For i = 1 To j
        skanblock(i) = rs.Fields("BLOCKID")
        rs.MoveNext
    Next i
    rs.Close
    
    
    '�C �B�Ŏ擾������ۯ�ID�������Ɋ֘A��ۯ������擾(XSDCA����)
    sql = "SELECT "
    sql = sql & "PLANTCATCA, "                          '����
    sql = sql & "XTALCA, "                              '�����ԍ�
    sql = sql & "CRYNUMCA, "                            '��ۯ�ID
    sql = sql & "HINBCA || "
    sql = sql & "TO_CHAR(NVL(REVNUMCA,0),'FM00') || "
    sql = sql & "FACTORYCA || "
    sql = sql & "OPECA AS HINBAN, "                     '�i��(12��)
    sql = sql & "SXLIDCA, "                             'SXLID
    sql = sql & "SXLIDCB, "                             'SXLID(XSDCB)�@08/07/10 ooba
    sql = sql & "GNWKNTCA, "                            '���ݍH��(XSDCA)
    sql = sql & "GNWKNTCB, "                            '���ݍH��(XSDCB)
    sql = sql & "INPOSCA, "                             '�������J�n�ʒu
    sql = sql & "INPOSCS, "                             '��ۯ��I���ʒu�@08/07/10 ooba
    sql = sql & "GNLCA, "                               '���ݒ���
    sql = sql & "GNMCA, "                               '���ݖ���
    sql = sql & "WFHOLDFLGCA, "                         'WFΰ��ދ敪
    sql = sql & "KDAYCA "                               '�X�V���t
    sql = sql & "FROM XSDCA,XSDCB,XSDCS "
    sql = sql & "WHERE CRYNUMCA IN ( "
    
    '�擾�����ύX�@08/10/28 ooba
    For i = 1 To UBound(skanblock)
        sql = sql & "'" & skanblock(i) & "' "
        If i <> UBound(skanblock) Then sql = sql & ","
    Next i
    
'''    '�B �A�Ŏ擾����SXLID���܂���ۯ�ID���擾(XSDCA����)
'''    sql = sql & "    SELECT "
'''    sql = sql & "    CRYNUMCA "
'''    sql = sql & "    FROM XSDCA "
'''    sql = sql & "    WHERE SXLIDCA IN ( "
'''    '�A �@�Ŏ擾����SXLID�̒��Ŋ֘A��ۯ���SXLID���擾(XSDCB����)
'''    sql = sql & "        SELECT "
'''    sql = sql & "        SXLIDCB "
'''    sql = sql & "        FROM XSDCB "
'''    sql = sql & "        WHERE SXLIDCB IN ( "
'''    '�@ �I��������ۯ�ID��������SXLID���擾(XSDCA����)
'''    sql = sql & "            SELECT "
'''    sql = sql & "            SXLIDCA "
'''    sql = sql & "            FROM XSDCA "
'''    sql = sql & "            WHERE CRYNUMCA = '" & sBlockID & "' "
'''    sql = sql & "            AND LIVKCA = '0' "
'''    sql = sql & "        ) "
'''    sql = sql & "        AND LIVKCB = '0' "
'''    sql = sql & "        AND KBLKFLGCB = '1' "
'''    sql = sql & "    ) "
    sql = sql & ") "
    sql = sql & "AND SXLIDCA LIKE '" & Mid(sBlockId, 1, 9) & "%' "      '08/10/28 ooba
    sql = sql & "AND (LIVKCA = '0' OR "
'    sql = sql & "     (LIVKCA = '1' AND LSTATBCA = 'H' AND LUFRBCA = 'H')) "        '�S���p���ް��@08/07/10 ooba
    sql = sql & "     (LIVKCA = '1' AND LSTATBCA in ('H','M','E') AND LUFRBCA = 'H')) "     '�����p�p,���ʏ����ǉ� 09/02/22 ooba
    sql = sql & "AND TBKBNCS = 'B' "
    sql = sql & "AND XSDCA.SXLIDCA = XSDCB.SXLIDCB(+) "
    sql = sql & "AND XSDCA.CRYNUMCA = XSDCS.CRYNUMCS "
    sql = sql & "ORDER BY CRYNUMCA, INPOSCA "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    '�ް�����
    If rs.RecordCount <= 0 Then
        GoTo proc_exit
    End If
    
    iBlkCnt = 0
    iHinCnt = 0
    
    ReDim tKanrenDisp(0)
    
    For i = 1 To rs.RecordCount
        '1�Ԗ��ް��܂��͑O�ް�����ۯ�ID���قȂ�ꍇ�A�s�ǉ�
        If iBlkCnt = 0 Or tKanrenDisp(iBlkCnt).BLOCKID <> rs.Fields("CRYNUMCA") Then
            iBlkCnt = iBlkCnt + 1       '��ۯ����{1
            iHinCnt = 1                 '�i�Ԑ�=1
            ReDim Preserve tKanrenDisp(iBlkCnt)
            
            With tKanrenDisp(iBlkCnt)
                '--����
                If IsNull(rs.Fields("PLANTCATCA")) = False Then
                    For j = 1 To UBound(s_MukesakiBase)
                        If s_MukesakiBase(j).sMukeCode = rs.Fields("PLANTCATCA") Then
                           .MUKESAKI = s_MukesakiBase(j).sMukeName
                        End If
                    Next j
                End If
                '--�����ԍ�
                If IsNull(rs.Fields("XTALCA")) = False Then .CRYNUM = rs.Fields("XTALCA")
                '--��ۯ�ID
                If IsNull(rs.Fields("CRYNUMCA")) = False Then .BLOCKID = rs.Fields("CRYNUMCA")
                '--�i�Ԑ�
                .HINCNT = iHinCnt
                '--�i��
                If IsNull(rs.Fields("HINBAN")) = False Then .hinban(iHinCnt) = rs.Fields("HINBAN")
                '--SXLID
                If IsNull(rs.Fields("SXLIDCA")) = False Then .SXLID(iHinCnt) = rs.Fields("SXLIDCA")
                '--SXLID(XSDCB)�@08/07/10 ooba
                If IsNull(rs.Fields("SXLIDCB")) = False Then .SXLID_CB(iHinCnt) = rs.Fields("SXLIDCB")
                '--SXLID(�X�V�p)
                .SXLID_NEW = ""
                '--�d�|�H��
                If rs.Fields("GNWKNTCA") = PROCD_WFC_SOUGOUHANTEI Then
                    'CW750��CW740�ϊ�
                    .Koutei = PROCD_NUKISI_HENKOU
                Else
                    If IsNull(rs.Fields("GNWKNTCA")) = False Then .Koutei = rs.Fields("GNWKNTCA")
                End If
                '--�������J�n�ʒu
                If IsNull(rs.Fields("INPOSCA")) = False Then .INGOTPOS(iHinCnt) = rs.Fields("INPOSCA")
                '--��ۯ��I���ʒu�@08/07/10 ooba
                If IsNull(rs.Fields("INPOSCS")) = False Then .BLKEPOS = rs.Fields("INPOSCS")
                '--����
                If IsNull(rs.Fields("GNLCA")) = False Then .LENGTH(iHinCnt) = rs.Fields("GNLCA")
                '--����
                If IsNull(rs.Fields("GNMCA")) = False Then .MAISU(iHinCnt) = rs.Fields("GNMCA")
                '--WFΰ��ދ敪
                If IsNull(rs.Fields("WFHOLDFLGCA")) = False Then .HOLD = rs.Fields("WFHOLDFLGCA")
                '--���t
                If IsNull(rs.Fields("KDAYCA")) = False Then .KDATE = rs.Fields("KDAYCA")
            End With
        '��ۯ�ID����v����ꍇ�A�i���ް��ǉ�
        Else
            iHinCnt = iHinCnt + 1       '�i�Ԑ��{1
            '�i�Ԑ��װ
            If iHinCnt > 5 Then
                GoTo proc_exit
            End If
            
            With tKanrenDisp(iBlkCnt)
                '--�i�Ԑ�
                .HINCNT = iHinCnt
                '--�i��
                If IsNull(rs.Fields("HINBAN")) = False Then .hinban(iHinCnt) = rs.Fields("HINBAN")
                '--SXLID
                If IsNull(rs.Fields("SXLIDCA")) = False Then .SXLID(iHinCnt) = rs.Fields("SXLIDCA")
                '--SXLID(XSDCB)�@08/07/10 ooba
                If IsNull(rs.Fields("SXLIDCB")) = False Then .SXLID_CB(iHinCnt) = rs.Fields("SXLIDCB")
                '--�������J�n�ʒu
                If IsNull(rs.Fields("INPOSCA")) = False Then .INGOTPOS(iHinCnt) = rs.Fields("INPOSCA")
                '--����
                If IsNull(rs.Fields("GNLCA")) = False Then .LENGTH(iHinCnt) = rs.Fields("GNLCA")
                '--����
                If IsNull(rs.Fields("GNMCA")) = False Then .MAISU(iHinCnt) = rs.Fields("GNMCA")
            End With
        End If
        rs.MoveNext
    Next i
    
    rs.Close
    
    DBDRV_scmzc_fcmkc001k_Disp2 = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�װ�����
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function

'Add Start 2010/07/09 SMPK Nakamura
'�T�v      :�֘A��ۯ����擾
'���Ұ��@�@:�ϐ���         ,IO ,�^               ,����
'      �@�@:sBlockID       ,I  ,String         �@,��ۯ�ID
'      �@�@:tKanrenDisp()  ,I  ,typ_KanrenDisp   ,�֘A��ۯ��ꗗ
'      �@�@:�߂�l         ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :
'����      :2010/07/09 SMPK Nakamura
Public Function DBDRV_scmzc_fcmkc001k_Disp3(ByVal sBlockId As String, _
                                            ByRef tKanrenList() As typ_KanrenList) As FUNCTION_RETURN

    Dim i           As Integer
    Dim iBlkCnt     As Integer      '��ۯ���
    Dim rs          As OraDynaset
    Dim sql         As String
    
    '�װ����ׂ̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function DBDRV_scmzc_fcmkc001k_Disp3"
    
    DBDRV_scmzc_fcmkc001k_Disp3 = FUNCTION_RETURN_FAILURE
    
    '�֘A��ۯ��R�ؕR�tð��ق��֘A��ۯ��擾
    sql = "SELECT "
    sql = sql & "BLOCKID, "
    sql = sql & "PROCCAT, "
    sql = sql & "DECODE( GNWKNTCA, "
    sql = sql & "        '" & PROCD_NUKISI_HENKOU & "', 0,"
    sql = sql & "        '" & PROCD_WFC_SOUGOUHANTEI & "', 0,"
    sql = sql & "        1) as WAITFLG "
    sql = sql & "FROM TBCMY023, XSDCA, XSDCS "
    sql = sql & "WHERE CRYNUM LIKE '" & Mid(sBlockId, 1, 9) & "%' "
    sql = sql & "AND TRANCNT = ( "
    sql = sql & "    SELECT "
    sql = sql & "    MAX(TRANCNT) "
    sql = sql & "    FROM TBCMY023 "
    sql = sql & "    WHERE CRYNUM LIKE '" & Mid(sBlockId, 1, 9) & "%' "
    sql = sql & "    AND BLOCKID = '" & sBlockId & "' "
    sql = sql & ") "
    sql = sql & "AND BLOCKID = CRYNUMCA "
    sql = sql & "AND (LIVKCA = '0' OR "
    sql = sql & "     (LIVKCA = '1' AND LSTATBCA in ('H','M','E') AND LUFRBCA = 'H')) " '�����p�p,���ʏ����ǉ�
    sql = sql & "AND SXLIDCA LIKE '" & Mid(sBlockId, 1, 9) & "%' "
    sql = sql & "AND TBKBNCS = 'B' "
    sql = sql & "AND XSDCA.CRYNUMCA = XSDCS.CRYNUMCS "
    sql = sql & "ORDER BY WAITFLG DESC, BLOCKID "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    '�ް�����
    If rs.RecordCount <= 0 Then
        GoTo proc_exit
    End If
    
    iBlkCnt = 0
    ReDim tKanrenList(0)
    '��ۯ�ID���
    For i = 1 To rs.RecordCount
        '�֘A��ۯ��łȂ��ꍇ
        If rs.Fields("PROCCAT") = "D" Then
            GoTo proc_exit
        End If
        If iBlkCnt = 0 Or tKanrenList(iBlkCnt).BLOCKID <> rs.Fields("BLOCKID") Then
            iBlkCnt = iBlkCnt + 1       '��ۯ����{1
            
            ReDim Preserve tKanrenList(iBlkCnt)

            tKanrenList(iBlkCnt).BLOCKID = rs.Fields("BLOCKID")
            tKanrenList(iBlkCnt).WAIT = rs.Fields("WAITFLG")
        End If
        rs.MoveNext
    Next i
    rs.Close
        
    DBDRV_scmzc_fcmkc001k_Disp3 = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�װ�����
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function
'Add End 2010/07/09 SMPK Nakamura

'�T�v      :�����ύX�w���p ���s���c�a�h���C�o
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:sStaffID�@�@�@,I  ,String         �@,�Ј�ID
'�@�@      :pBlkInf �@�@�@,I  ,typ_BlkInf3    �@,�u���b�N���
'      �@�@:pLackMap�@�@�@,I  ,typ_LackMap    �@,�����E�F�n�[
'      �@�@:pSXLMng �@�@�@,I  ,typ_TBCME042   �@,SXL�Ǘ�
'      �@�@:pWafSmp �@�@�@,I  ,typ_XSDCW   �@   ,�V�T���v���Ǘ��iSXL�j
'      �@�@:pMesInd �@�@�@,I  ,typ_TBCMY003   �@,����]�����@�w��
'      �@�@:pTrnScr �@�@�@,I  ,typ_TBCMW006   �@,�U�֔p������
'      �@�@:pSXLDcd �@�@�@,I  ,typ_TBCMY007   �@,SXL�m��w��
'      �@�@:pEpMesInd �@  ,I  ,typ_TBCMY020   �@,EP����]���w��
'      �@�@:sKanrenB �@   ,I  ,String         �@,�֘A��ۯ��@07/08/06 ooba
'      �@�@:sErrMsg �@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :
'����      :2001/07/11 ���{ �쐬
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
Public Function DBDRV_scmzc_fcmkc001k_Exec(sStaffID As String, pBlkInf() As typ_BlkInf3, _
                                           pLackMap() As typ_LackMap, pSXLMng() As typ_TBCME042, _
                                           pWafSmp() As typ_XSDCW, pMesInd() As typ_TBCMY003, _
                                           pTrnScr() As typ_TBCMW006, pSXLDcd() As typ_TBCMY007, pEpMesInd() As typ_TBCMY020, sKanrenB() As String, sErrMsg As String) As FUNCTION_RETURN
''Public Function DBDRV_scmzc_fcmkc001k_Exec(sStaffID As String, pBlkInf() As typ_BlkInf3, _
''                                           pLackMap() As typ_LackMap, pSXLMng() As typ_TBCME042, _
''                                           pWafSmp() As typ_XSDCW, pMesInd() As typ_TBCMY003, _
''                                           pTrnScr() As typ_TBCMW006, pSXLDcd() As typ_TBCMY007, sErrMsg As String) As FUNCTION_RETURN
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
    Dim sql     As String
    Dim sDbName As String
    Dim sCryNum As String
    Dim Blks    As String
    Dim sTmpSxl() As String     '�d�|�H���������pSXLID�@06/03/14 ooba
    Dim recCnt  As Long
    Dim i       As Long
    Dim dynOra  As OraDynaset

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001k_SQL.bas -- Function DBDRV_scmzc_fcmkc001k_Exec"
    sErrMsg = ""

    '' WriteDBLog " ", "Start"

    DBDRV_scmzc_fcmkc001k_Exec = FUNCTION_RETURN_FAILURE

    '�d�|�H���ă`�F�b�N�@�\�ǉ��@06/03/14 ooba START ========================================>
    sDbName = "XSDCA"
    sql = "SELECT SXLIDCA "
    sql = sql & " FROM XSDCA"
    sql = sql & " WHERE CRYNUMCA LIKE '" & Left(pBlkInf(1).BLOCKID, 9) & "%'"
    sql = sql & "   AND (INPOSCA>=" & SIngotP
    sql = sql & "   AND  INPOSCA< " & EIngotP & ")"
    sql = sql & "   AND LIVKCA = '0' "
    sql = sql & "GROUP BY SXLIDCA"
    Set dynOra = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    ReDim sTmpSxl(0)
    If dynOra.RecordCount > 0 Then
        For i = 1 To dynOra.RecordCount
            If Not IsNull(dynOra.Fields("SXLIDCA")) Then
                ReDim Preserve sTmpSxl(i)
                sTmpSxl(i) = dynOra.Fields("SXLIDCA")
            End If
            dynOra.MoveNext
        Next i
    End If
    dynOra.Close
    If DBDRV_CheckCodeXSDCB(sTmpSxl, PROCD_NUKISI_HENKOU, sErrMsg) = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmkc001k_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    '�d�|�H���ă`�F�b�N�@�\�ǉ��@06/03/14 ooba END ==========================================>

    '' SXL�Ǘ��̑}��(s_cmzcDBdriverCOM_SQL.bas ���K�v)
    recCnt = UBound(pSXLMng)
    If recCnt > 0 Then
        ''���ύXSTART SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '06/1/5 SMP�ΐ�
        '' XSDCB�ɕK�v�ȃf�[�^�����݂���\�����l���ADelete��Insert�͂�߂�
        sDbName = "XSDCB"
        If DBDRV_SXL_INS_CB(pSXLMng()) = FUNCTION_RETURN_FAILURE Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            GoTo proc_exit
        End If
        '�������W�b�N�iSXL�Ǘ��iE042�j��XSDCB�@�\�ڍs�j
'        sDbName = "E042"
'        sql = "delete from  TBCME042"
'        sql = sql & " where CRYNUM   ='" & pSXLMng(1).CRYNUM & "'"
''       sql = sql & "   and INGOTPOS>= " & pSXLMng(1).IngotPos
''       sql = sql & "   and INGOTPOS<  " & pSXLMng(recCnt).IngotPos + pSXLMng(recCnt).LENGTH
'        sql = sql & "   and INGOTPOS>= " & SIngotP
'        sql = sql & "   and INGOTPOS<  " & EIngotP
'        WriteDBLog sql, sDbName
'        Call OraDB.ExecuteSQL(sql)
'        If DBDRV_SXL_INS(pSXLMng()) = FUNCTION_RETURN_FAILURE Then
'            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
'            GoTo proc_exit
'        End If
        ''���ύXSTART SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '06/1/5 SMP�ΐ�
    End If

    '' WF�T���v���Ǘ��̑}��(s_cmzcDBdriverCOM_SQL.bas ���K�v)
    If recCnt > 0 Then
        sDbName = "XSDCW"
'        sDbName = "E044"
'        '�͈͊J�n�ʒu��WF�T���v�����폜
'        sql = "delete from XSDCW"
'        sql = sql & " where XTALCW='" & pSXLMng(1).CRYNUM & "'"
''       sql = sql & " and INGOTPOS=" & pSXLMng(1).IngotPos
'        sql = sql & " and INPOSCW=" & SIngotP
'        sql = sql & " and SMPKBNCW in ('T', 'D')"
'        WriteDBLog sql, sDbName
'        Call OraDB.ExecuteSQL(sql)
'        '�͈͂Ɋ��S�Ɋ܂܂��WF�T���v�����폜
'        sql = "delete from XSDCW"
'        sql = sql & " where XTALCW='" & pSXLMng(1).CRYNUM & "'"
''       sql = sql & " and INGOTPOS>" & pSXLMng(1).IngotPos
''       sql = sql & " and INGOTPOS<" & pSXLMng(recCnt).IngotPos + pSXLMng(recCnt).LENGTH
'        sql = sql & " and INPOSCW>" & SIngotP
'        sql = sql & " and INPOSCW<" & EIngotP
'        WriteDBLog sql, sDbName
'        Call OraDB.ExecuteSQL(sql)
'        '�͈͏I���ʒu��WF�T���v�����폜
'        sql = "delete from XSDCW"
'        sql = sql & " where XTALCW='" & pSXLMng(1).CRYNUM & "'"
''       sql = sql & " and INGOTPOS=" & pSXLMng(recCnt).IngotPos + pSXLMng(recCnt).LENGTH
'        sql = sql & " and INPOSCW=" & EIngotP
'        sql = sql & " and SMPKBNCW in ('B', 'U')"
'        WriteDBLog sql, sDbName
'        Call OraDB.ExecuteSQL(sql)

        '�V�T���v���Ǘ��Ƀf�[�^�����邩
        For i = 1 To UBound(pWafSmp)
            sql = "SELECT count(*) "
            sql = sql & "FROM  XSDCW "
            sql = sql & "WHERE SXLIDCW ='" & pWafSmp(i).SXLIDCW & "'"
        '   sql = sql & "  and SMPKBNCW='" & pWafSmp(i).SMPKBNCW & "'"
            sql = sql & "  and TBKBNCW ='" & pWafSmp(i).TBKBNCW & "'"

            Set dynOra = OraDB.DBCreateDynaset(sql, 0&)
        '   If 0 < OraDB.ExecuteSQL(sql) Then
            If dynOra.Fields(0) <> 0 Then
                '�f�[�^�������Update
                If DBDRV_WfSmp_UPD(pWafSmp(), i) = FUNCTION_RETURN_FAILURE Then
                    sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                    GoTo proc_exit
                End If
            Else
                '�Ȃ����Insert
                If DBDRV_WfSmp_INS(pWafSmp(), i) = FUNCTION_RETURN_FAILURE Then
                    sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                    GoTo proc_exit
                End If
            End If
        Next i
    End If


    '' �����ύX�w�����т̑}��
    sDbName = "W003"
    recCnt = UBound(pBlkInf)
    For i = 1 To recCnt
        With pBlkInf(i)
            sCryNum = Left(.BLOCKID, 9) & "000"
            sql = "insert into TBCMW003 "
            sql = sql & "(CRYNUM, INGOTPOS, TRANCNT, CRYLEN, KRPROCCD, PROCCODE, BLOCKID, DELFLG, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE)"

            sql = sql & " select '"
            sql = sql & sCryNum & "', "
            sql = sql & .COF.TOPSMPLPOS & ", "
            sql = sql & "nvl(max(TRANCNT),0)+1, "
            sql = sql & .REALLEN & ", '"
            sql = sql & MGPRCD_NUKISI_HENKOU & "', '"
            sql = sql & PROCD_NUKISI_HENKOU & "', '"
            sql = sql & .BLOCKID & "', '"
            sql = sql & .DELFLG & "', '"
            sql = sql & sStaffID & "', "
            sql = sql & "sysdate, '"
            sql = sql & sStaffID & "', "
            sql = sql & "sysdate, "
            sql = sql & "'0', "
            sql = sql & "sysdate"
            sql = sql & " from  TBCMW003"
            sql = sql & " where CRYNUM  ='" & sCryNum & "'"
            sql = sql & "   and INGOTPOS= " & .COF.TOPSMPLPOS

            '' WriteDBLog sql, sDbName

            Debug.Print sql

            If OraDB.ExecuteSQL(sql) <= 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                GoTo proc_exit
            End If
        End With
    Next i

    '' ����]�����@�w���̑}��(s_cmzcDBdriverCOM_SQL.bas ���K�v)
    sDbName = "Y003"
    If DBDRV_SokuSizi_Ins(pMesInd()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        GoTo proc_exit
    End If

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    '' �G�s����]���w�����̑}��(s_cmzcDBdriverCOM_SQL.bas ���K�v)
    sDbName = "Y020"
    If DBDRV_SokuSizi_EP_Ins(pEpMesInd()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        GoTo proc_exit
    End If
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

    '' �U�֔p�����т̑}��(s_cmzcDBdriverCOM_SQL.bas ���K�v)
    sDbName = "W006"
    recCnt = UBound(pTrnScr)
    For i = 1 To recCnt
        If DBDRV_Furikae_Ins(pTrnScr(i)) = FUNCTION_RETURN_FAILURE Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            GoTo proc_exit
        End If
    Next i

    '' SXL�m��w���̑}��
    sDbName = "Y007"
    recCnt = UBound(pSXLDcd)
    For i = 1 To recCnt
        With pSXLDcd(i)
            sql = "insert into TBCMY007 "
            ' 2007/09/03 SPK Tsutsumi Add Start
'            sql = sql & "(SXL_ID, SAMPLE_FROM, SAMPLE_TO, BLOCKID, HINBAN, KUBUN, TXID, REGDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE)"
            sql = sql & "(SXL_ID, SAMPLE_FROM, SAMPLE_TO, BLOCKID, HINBAN, KUBUN, TXID, REGDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE,PLANTCAT)"
            ' 2007/09/03 SPK Tsutsumi Add Start
            sql = sql & " values ('"
            sql = sql & .SXL_ID & "', '"        ' SXL-ID
            sql = sql & .SAMPLE_FROM & "', '"   ' �T���v��ID (From)
            sql = sql & .SAMPLE_TO & "', '"     ' �T���v��ID (To)
            sql = sql & .BLOCKID & "', '"       ' �u���b�N�h�c
            sql = sql & .hinban & "', "         ' �m��i��
            sql = sql & "'S ', "                ' �敪�R�[�h
            sql = sql & "'TX853I', "            ' �g�����U�N�V����ID
            sql = sql & "sysdate, "             ' �o�^���t
            sql = sql & "'0', "                 ' SUMMIT���M�t���O
            sql = sql & "'0', "                 ' ���M�t���O

            ' 2007/09/03 SPK Tsutsumi Add Start
            sql = sql & "sysdate,"              ' ���M���t
            sql = sql & "'" & sCmbMukesaki & "'"  ' ����
            ' 2007/09/03 SPK Tsutsumi Add End

            '' WriteDBLog sql, sDbName

            Debug.Print sql

            If OraDB.ExecuteSQL(sql) <= 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                DBDRV_scmzc_fcmkc001k_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        End With
    Next i

'2003/01/09 ooba �`�F�b�N�t���O�����ύX
    '' �������̍X�V
    sDbName = "Y012"
    Dim m As Integer
    Dim j As Long
    m = UBound(pBlkInf)
    recCnt = UBound(pLackMap)
    For i = 1 To m
        For j = 1 To recCnt
            If pBlkInf(i).BLOCKID = pLackMap(j).BLOCKID Then
                sql = "update TBCMY012 set CHKFLG='1' where LOTID='" & pLackMap(j).BLOCKID & "'"
                '' WriteDBLog sql, sDbName

                Debug.Print sql

                If OraDB.ExecuteSQL(sql) < 0 Then
                    sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                    GoTo proc_exit
                End If
            End If
        Next j
    Next i

    'XSDCB��ΰ��ދ敪(WF)�X�V�@04/06/30 ooba
    Dim SxlCnt As Integer
    sDbName = "XSDCB"
    For i = 1 To UBound(pSXLMng)
        sql = "select count(*) from XSDCA "
        sql = sql & "where LIVKCA = '0' "
        sql = sql & "and SXLIDCA = '" & pSXLMng(i).SXLID & "' "
        Set dynOra = OraDB.DBCreateDynaset(sql, 0&)
        SxlCnt = dynOra.Fields(0)
        'SXL�ް������݂���ꍇ
        If SxlCnt > 0 Then
            sql = "select count(*) from XSDCA, XSDCB "
            sql = sql & "where LIVKCA = '0' "
            sql = sql & "and LIVKCB = '0' "
            sql = sql & "and (WFHOLDFLGCA != '1' "
            sql = sql & "or WFHOLDFLGCA is NULL) "
            sql = sql & "and WFHOLDFLGCB = '1' "
            sql = sql & "and SXLIDCA = SXLIDCB "
            sql = sql & "and SXLIDCA = '" & pSXLMng(i).SXLID & "' "
            Set dynOra = OraDB.DBCreateDynaset(sql, 0&)
            'XSDCB��ΰ��ދ敪(WF)���u1�v��XSDCA��ΰ��ދ敪(WF)�����ׂāu1�v�ȊO�̏ꍇ
            If dynOra.Fields(0) = SxlCnt Then
                'XSDCB��ΰ��ދ敪(WF)���u0�v�ɍX�V
                sql = "update XSDCB set WFHOLDFLGCB = '0' where SXLIDCB = '" & pSXLMng(i).SXLID & "' "
                '' WriteDBLog sql, sDbName
                Debug.Print sql

                If OraDB.ExecuteSQL(sql) < 0 Then
                    sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                    DBDRV_scmzc_fcmkc001k_Exec = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            End If
        End If
    Next i

    '�֘A�u���b�N���o�^��~�@08/01/23 ooba
''    '�֘A��ۯ����o�^�@07/08/06 ooba START =====================================>
''    If UBound(sKanrenB) > 1 Then
''        sDbName = "Y023"
''        If DBDRV_KanrenBlk(left(pBlkInf(1).BLOCKID, 9) & "000", sKanrenB(), _
''                            SIngotP, EIngotP) = FUNCTION_RETURN_FAILURE Then
''
''            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
''            GoTo proc_exit
''        End If
''    End If
''    '�֘A��ۯ����o�^�@07/08/06 ooba END =======================================>

'    '' �������̍X�V
'    sDBName = "Y012"
'    recCnt = UBound(pLackMap)
'    For i = 1 To recCnt
'        sql = "update TBCMY012 set CHKFLG='1'"
'        sql = sql & " where LOTID='" & pLackMap(i).BLOCKID & "'"
'        WriteDBLog sql, sDBName
'        If OraDB.ExecuteSQL(sql) < 0 Then
'            sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
'            GoTo proc_exit
'        End If
'    Next i
'���M�ς̑���]���w�����A�V����TRANCNT�ō쐬�����B
'���̂��߁A�������R�[�h�̍đ��͕s�v
'    '' �������܂ރu���b�N�ɂ��鑪��]���w���̍đ��i�������̃T���v��ID�������j
'    sDBName = "Y003-2"
'    Blks = vbNullString
'    If UBound(pBlkInf) > 0 Then '�K������͂������O�̂���
'        For i = 1 To UBound(pBlkInf)
'            Blks = Blks & "'" & pBlkInf(i).BLOCKID & "',"
'        Next i
'        Blks = Left$(Blks, Len(Blks) - 1)
'        sql = "update TBCMY003 set SENDFLAG='0'"
'        sql = sql & " where  (substr(SAMPLEID,1,12) in (" & Blks & "))"
'        sql = sql & " and    (substr(SAMPLEID,1,12)"
'        sql = sql & " in     (select distinct LOTID from TBCMY012"
'        sql = sql & " where  ((REJCAT='A') or (REJCAT='B')) and (ALLSCRAP<>'Y')))"
'        sql = sql & " and    (substr(SAMPLEID,1,15)"
'        sql = sql & " not in (select distinct LOTID || to_char(TOP_POS,'FM000')"
'        sql = sql & " as LOTPOS from TBCMY012"
'        sql = sql & " where ((REJCAT='A') or (REJCAT='B')) and (ALLSCRAP<>'Y')))"
'        WriteDBLog sql, sDBName
'        If OraDB.ExecuteSQL(sql) < 0 Then
'            sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
'            GoTo proc_exit
'        End If
'    End If

    DBDRV_scmzc_fcmkc001k_Exec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    '' WriteDBLog " ", "End"
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
    DBDRV_scmzc_fcmkc001k_Exec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

Public Function WFExistCheck(tblLackMap() As typ_LackMap, sBLK As String, iPos As Integer, sDirection As String, bAns As Integer) As FUNCTION_RETURN
    Dim iBseq   As Integer
    Dim sql     As String
    Dim rs      As OraDynaset
    Dim c0      As Integer


    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001k_SQL.bas -- Function WFExistCheck"

    WFExistCheck = FUNCTION_RETURN_FAILURE

    bAns = 1    '�����l�F�v�e�����݂���

    sql = "select BLOCKSEQ"
    sql = sql & " from  TBCMY011"
    sql = sql & " where (LOTID='" & sBLK & "')"

    If (sDirection = "T") Or (sDirection = "D") Then
        sql = sql & " and (TOP_POS=ANY(select min(TOP_POS) from TBCMY011 where (LOTID='" & sBLK & "') and (TOP_POS>=" & iPos * 10 & ")))"
    Else
        sql = sql & " and (TOP_POS=ANY(select max(TOP_POS) from TBCMY011 where (LOTID='" & sBLK & "') and (TOP_POS<=" & iPos * 10 & ")))"
    End If

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount > 0 Then
        iBseq = rs("BLOCKSEQ")
        rs.Close
    Else
        bAns = 0    '�����ɂv�e�͂Ȃ�
        iBseq = -99
        rs.Close
    End If

    '' �u���b�N�o�̌����`�F�b�N
    For c0 = 1 To UBound(tblLackMap)
        With tblLackMap(c0)
'            If (.BLOCKID = sBLK And .REJCAT = "A" And .LACKCNTS <= iBseq And .LACKCNTE >= iBseq) _
'            Or (.BLOCKID = sBLK And .LACKCNTS < 0) Then
            '��ۯ���Ԃł̈ꕔ���ʑΉ� 09/02/27 ooba
            If (.BLOCKID = sBLK And (.REJCAT = "A" Or .REJCAT = "E") And .LACKCNTS <= iBseq And .LACKCNTE >= iBseq) _
            Or (.BLOCKID = sBLK And .LACKCNTS < 0) Then
                bAns = -1   '�����͌������Ă���
                Exit For
            End If
        End With
    Next


    WFExistCheck = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    WFExistCheck = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'(2002/07 s_cmzcF_cmkc001g_SQL.bas���R�s�[)
'�T�v      :�����w���p ���i�d�l��p�c�a�h���C�o
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:pHinSpec�@�@�@,IO ,typ_HinSpec    �@,���i�d�l
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
Public Function DBDRV_scmzc_fcmkc001g_GetSpec(pHinSpec As typ_HinSpec) As FUNCTION_RETURN

    Dim rs      As OraDynaset
    Dim sql     As String
    Dim sOT1    As String           '03/05/23 �㓡
    Dim sOT2    As String
    Dim rtn     As FUNCTION_RETURN
    Dim sMAI1    As String           '04/06/28
    Dim sMAI2    As String           '04/06/28

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001g_SQL.bas -- Function DBDRV_scmzc_fcmkc001g_GetSpec"

    '' ���i�d�l�̎擾
    With pHinSpec
        sql = "select E021HWFRMIN,  E021HWFRMAX,  E021HWFRHWYS, E024HWFMKHWS, E025HWFONHWS, E025HWFOS1HS,"
        sql = sql & " E025HWFOS2HS, E025HWFOS3HS, E026HWFDSOHS, E028HWFSPVHS, E028HWFDLHWS, E029HWFOF1HS,"
        sql = sql & " E029HWFOF2HS, E029HWFOF3HS, E029HWFOF4HS, E029HWFBM1HS, E029HWFBM2HS, E029HWFBM3HS "
        sql = sql & " from  VECME004"
        sql = sql & " where E018HINBAN  ='" & .hin.hinban & "'"
        sql = sql & "   and E018MNOREVNO= " & .hin.mnorevno
        sql = sql & "   and E018FACTORY ='" & .hin.factory & "'"
        sql = sql & "   and E018OPECOND ='" & .hin.opecond & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        .HWFRMIN = fncNullCheck(rs("E021HWFRMIN"))
        .HWFRMAX = fncNullCheck(rs("E021HWFRMAX"))
        .HWFRHWYS = rs("E021HWFRHWYS")
        .HWFMKHWS = rs("E024HWFMKHWS")
        .HWFONHWS = rs("E025HWFONHWS")
        .HWFOS1HS = rs("E025HWFOS1HS")
        .HWFOS2HS = rs("E025HWFOS2HS")
        .HWFOS3HS = rs("E025HWFOS3HS")
        .HWFDSOHS = rs("E026HWFDSOHS")
        .HWFSPVHS = rs("E028HWFSPVHS")
        .HWFDLHWS = rs("E028HWFDLHWS")
        .HWFOF1HS = rs("E029HWFOF1HS")
        .HWFOF2HS = rs("E029HWFOF2HS")
        .HWFOF3HS = rs("E029HWFOF3HS")
        .HWFOF4HS = rs("E029HWFOF4HS")
        .HWFBM1HS = rs("E029HWFBM1HS")
        .HWFBM2HS = rs("E029HWFBM2HS")
        .HWFBM3HS = rs("E029HWFBM3HS")
        'rtn = scmzc_getE036(pHinSpec.HIN, sOT1, sOT2)   '03/05/23
         'rtn = scmzc_getE036(pHinSpec.HIN, sOT1, sOT2)   '04/07/12 koyama update
        rtn = scmzc_getE036(pHinSpec.hin, sOT1, sOT2, sMAI1, sMAI2)   ''04/07/12 koyama update
        If rtn = FUNCTION_RETURN_FAILURE Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        .HWFOTHER1 = sOT1 '### 03/05/23
        .HWFOTHER2 = sOT2
        .HWFOTHER1MAI = sMAI1  '04/06/28
        .HWFOTHER2MAI = sMAI2  '04/06/28
        rs.Close

        ''�c���_�f�d�l�擾�@03/12/11 ooba START ==============================>
        sql = "select HWFZOHWS from TBCME025 "
        sql = sql & "where HINBAN  ='" & .hin.hinban & "' "
        sql = sql & "and MNOREVNO= " & .hin.mnorevno & " "
        sql = sql & "and FACTORY ='" & .hin.factory & "' "
        sql = sql & "and OPECOND ='" & .hin.opecond & "' "

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        If IsNull(rs("HWFZOHWS")) = False Then .HWFZOHWS = rs("HWFZOHWS") Else .HWFZOHWS = " "  '�iWF�c���_�f�ۏؕ��@_��
        rs.Close
        ''�c���_�f�d�l�擾�@03/12/11 ooba END ================================>

        '' GD�d�l�擾�@05/01/31 ooba START ================================================>
        sql = "select "
        sql = sql & "HWFDENHS, "
        sql = sql & "HWFLDLHS, "
        sql = sql & "HWFDVDHS "
        sql = sql & "from TBCME026 "
        sql = sql & "where HINBAN  ='" & .hin.hinban & "' "
        sql = sql & "and MNOREVNO= " & .hin.mnorevno & " "
        sql = sql & "and FACTORY ='" & .hin.factory & "' "
        sql = sql & "and OPECOND ='" & .hin.opecond & "' "

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        If IsNull(rs("HWFDENHS")) = False Then .HWFDENHS = rs("HWFDENHS") Else .HWFDENHS = " "  '�iWFDen�ۏؕ��@_��
        If IsNull(rs("HWFLDLHS")) = False Then .HWFLDLHS = rs("HWFLDLHS") Else .HWFLDLHS = " "  '�iWFL/DL�ۏؕ��@_��
        If IsNull(rs("HWFDVDHS")) = False Then .HWFDVDHS = rs("HWFDVDHS") Else .HWFDVDHS = " "  '�iWFDVD2�ۏؕ��@_��

        rs.Close
        '' GD�d�l�擾�@05/01/31 ooba END ==================================================>

        '' SPVNr�Z�x�d�l�擾�@06/06/08 ooba START ===========================>
        sql = "select "
        sql = sql & "HWFNRHS "          '�iWFSPVNR�ۏؕ��@_��
        '��--- 2010/01/20 SIRD�Ή� SPK habuki ADD START ---���jHWFOF4HS�̴ر�����g�p�̂��߁ATBCME048.HWFSIRDHS�ōė��p
        sql = sql & ",HWFSIRDHS "       '����]�ʕۏؕ��@�Q��
        '��--- 2010/01/20 SIRD�Ή� SPK habuki ADD END   ---���jHWFOF4HS�̴ر�����g�p�̂��߁ATBCME048.HWFSIRDHS�ōė��p
        sql = sql & "from TBCME048 "
        sql = sql & "where HINBAN = '" & .hin.hinban & "' "
        sql = sql & "and MNOREVNO = " & .hin.mnorevno & " "
        sql = sql & "and FACTORY = '" & .hin.factory & "' "
        sql = sql & "and OPECOND = '" & .hin.opecond & "' "

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        If IsNull(rs("HWFNRHS")) = False Then .HWFNRHS = rs("HWFNRHS") Else .HWFNRHS = " "
        '��--- 2010/01/20 SIRD�Ή� SPK habuki ADD START ---���jHWFOF4HS�̴ر�����g�p�̂��߁ATBCME048.HWFSIRDHS�ōė��p
        If IsNull(rs("HWFSIRDHS")) = False Then .HWFOF4HS = rs("HWFSIRDHS") Else .HWFOF4HS = " "
        '��--- 2010/01/20 SIRD�Ή� SPK habuki ADD END   ---���jHWFOF4HS�̴ر�����g�p�̂��߁ATBCME048.HWFSIRDHS�ōė��p

        rs.Close
        '' SPVNr�Z�x�d�l�擾�@06/06/08 ooba START ===========================>

        '' WF�J�b�g�P�ʎ擾�@05/04/12 ffc)tanabe START =====================================>
        sql = "select "
        sql = sql & "TO_CHAR(WFCUTT) as WFCUTT "
        sql = sql & "from TBCME036 "
        sql = sql & "where HINBAN  ='" & .hin.hinban & "' "
        sql = sql & "and MNOREVNO= " & .hin.mnorevno & " "
        sql = sql & "and FACTORY ='" & .hin.factory & "' "
        sql = sql & "and OPECOND ='" & .hin.opecond & "' "

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        If IsNull(rs("WFCUTT")) = False Then .WFCUTUNIT = rs("WFCUTT")   'WF�J�b�g�P��

        rs.Close
        '' WF�J�b�g�P�ʎ擾�@05/04/12 ffc)tanabe END =======================================>

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
        '' �G�s�d�l�擾(OSF�ABND)
        sql = "select "
        sql = sql & "HEPOF1HS, "        '�iEPOSF1�ۏؕ��@_��
        sql = sql & "HEPOF2HS, "        '�iEPOSF2�ۏؕ��@_��
        sql = sql & "HEPOF3HS, "        '�iEPOSF3�ۏؕ��@_��
        sql = sql & "HEPBM1HS, "        '�iEPBMD1�ۏؕ��@_��
        sql = sql & "HEPBM2HS, "        '�iEPBMD2�ۏؕ��@_��
        sql = sql & "HEPBM3HS "         '�iEPBMD3�ۏؕ��@_��
        sql = sql & "from TBCME050 "    '���i�d�l�G�s�f�[�^�P
        sql = sql & "where HINBAN = '" & .hin.hinban & "' "
        sql = sql & "and MNOREVNO = " & .hin.mnorevno & " "
        sql = sql & "and FACTORY = '" & .hin.factory & "' "
        sql = sql & "and OPECOND = '" & .hin.opecond & "' "

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        If IsNull(rs("HEPOF1HS")) = False Then .HEPOF1HS = rs("HEPOF1HS") Else .HEPOF1HS = " "   '�iEPOSF1�ۏؕ��@_��
        If IsNull(rs("HEPOF2HS")) = False Then .HEPOF2HS = rs("HEPOF2HS") Else .HEPOF2HS = " "   '�iEPOSF2�ۏؕ��@_��
        If IsNull(rs("HEPOF3HS")) = False Then .HEPOF3HS = rs("HEPOF3HS") Else .HEPOF3HS = " "   '�iEPOSF3�ۏؕ��@_��
        If IsNull(rs("HEPBM1HS")) = False Then .HEPBM1HS = rs("HEPBM1HS") Else .HEPBM1HS = " "   '�iEPBMD1�ۏؕ��@_��
        If IsNull(rs("HEPBM2HS")) = False Then .HEPBM2HS = rs("HEPBM2HS") Else .HEPBM2HS = " "   '�iEPBMD2�ۏؕ��@_��
        If IsNull(rs("HEPBM3HS")) = False Then .HEPBM3HS = rs("HEPBM3HS") Else .HEPBM3HS = " "   '�iEPBMD3�ۏؕ��@_��
        rs.Close
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

    End With

    DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME037�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME037 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcF_TBCME037_SQL.bas���ړ�)
Public Function DBDRV_GetTBCME037(records() As typ_TBCME037, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql     As String       'SQL�S��
Dim sqlBase As String       'SQL��{��(WHERE�߂̑O�܂�)
Dim rs      As OraDynaset   'RecordSet
Dim recCnt  As Long         '���R�[�h��
Dim i       As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, DELCLS, KRPROCCD, PROCCD, LPKRPROCCD, LASTPASS, RPHINBAN, RPREVNUM, RPFACT, RPOPCOND, PRODCOND," & _
              " PGID, UPLENGTH, TOPLENG, BODYLENG, BOTLENG, FREELENG, DIAMETER, CHARGE, SEED, ADDDPCLS, ADDDPPOS, ADDDPVAL," & _
              " REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME037"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME037 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .DELCLS = rs("DELCLS")           ' �폜�敪
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCD = rs("PROCCD")           ' �H���R�[�h
            .LPKRPROCCD = rs("LPKRPROCCD")   ' �ŏI�ʉߊǗ��H��
            .LASTPASS = rs("LASTPASS")       ' �ŏI�ʉߍH��
            .RPHINBAN = rs("RPHINBAN")       ' �˂炢�i��
            .RPREVNUM = rs("RPREVNUM")       ' �˂炢�i�Ԑ��i�ԍ������ԍ�
            .RPFACT = rs("RPFACT")           ' �˂炢�i�ԍH��
            .RPOPCOND = rs("RPOPCOND")       ' �˂炢�i�ԑ��Ə���
            .PRODCOND = rs("PRODCOND")       ' �������
            .PGID = rs("PGID")               ' �o�f�|�h�c
            .UPLENGTH = rs("UPLENGTH")       ' ���グ����
            .TOPLENG = rs("TOPLENG")         ' �s�n�o����
            .BODYLENG = rs("BODYLENG")       ' ��������
            .BOTLENG = rs("BOTLENG")         ' �a�n�s����
            .FREELENG = rs("FREELENG")       ' �t���[��
            .DIAMETER = rs("DIAMETER")       ' ���a
            .CHARGE = rs("CHARGE")           ' �`���[�W��
            .SEED = rs("SEED")               ' �V�[�h
            .ADDDPCLS = rs("ADDDPCLS")       ' �ǉ��h�[�v���
            .ADDDPPOS = rs("ADDDPPOS")       ' �ǉ��h�[�v�ʒu
            .ADDDPVAL = rs("ADDDPVAL")       ' �ǉ��h�[�v��
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME037 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME039�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME039 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcF_TBCME039_SQL.bas���ړ�)
Public Function DBDRV_GetTBCME039(records() As typ_TBCME039, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql     As String       'SQL�S��
Dim sqlBase As String       'SQL��{��(WHERE�߂̑O�܂�)
Dim rs      As OraDynaset   'RecordSet
Dim recCnt  As Long         '���R�[�h��
Dim i       As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, INGOTPOS, HINBAN, REVNUM, FACT, OPCOND, LENGTH, USECLASS, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME039"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME039 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .INGOTPOS = rs("INGOTPOS")       ' �������J�n�ʒu
            .hinban = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' �����ԍ�
            .FACT = rs("FACT")               ' �H��
            .OPCOND = rs("OPCOND")           ' ���Ə���
            .LENGTH = rs("LENGTH")           ' ����
            .USECLASS = rs("USECLASS")       ' �g�p�敪
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME039 = FUNCTION_RETURN_SUCCESS
End Function



'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME041�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME041 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺  (2002/07 s_cmzcTBCME041_SQL.bas���ړ�)
Public Function DBDRV_GetTBCME041(records() As typ_TBCME041, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql     As String       'SQL�S��
Dim sqlBase As String       'SQL��{��(WHERE�߂̑O�܂�)
Dim rs      As OraDynaset   'RecordSet
Dim recCnt  As Long         '���R�[�h��
Dim i       As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, INGOTPOS, HINBAN, REVNUM, FACTORY, OPECOND, LENGTH, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME041"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME041 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .INGOTPOS = rs("INGOTPOS")       ' �������J�n�ʒu
            .hinban = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .factory = rs("FACTORY")         ' �H��
            .opecond = rs("OPECOND")         ' ���Ə���
            .LENGTH = rs("LENGTH")           ' ����
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME041 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME042�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME042 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺  (2002/07 s_cmzcTBCME042_SQL.bas���ړ�)
Public Function DBDRV_GetTBCME042(records() As typ_TBCME042, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql     As String       'SQL�S��
Dim sqlBase As String       'SQL��{��(WHERE�߂̑O�܂�)
Dim rs      As OraDynaset   'RecordSet
Dim recCnt  As Long         '���R�[�h��
Dim i       As Long

    ''SQL��g�ݗ��Ă�
    ''���ύXSTART SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
'    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, SXLID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, DELCLS, LSTATCLS, HOLDCLS," & _
'              " HINBAN, REVNUM, FACTORY, OPECOND, BDCAUS, COUNT, REGDATE, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE, " & _
'              " PASSFLAG "   '02/04/16 Yam
'    sqlBase = sqlBase & "From TBCME042"
'    sql = sqlBase
    sql = ""
    sql = sql & " SELECT"
    sql = sql & "   XTALCB CRYNUM"      '�����ԍ�
    sql = sql & "  ,INPOSCB INGOTPOS"   '�������J�n�ʒu
    sql = sql & "  ,RLENCB LENGTH"      '���_����
    sql = sql & "  ,SXLIDCB SXLID"      'SXLID
    sql = sql & "  ,'     ' KRPROCCD"   '�Ǘ��H��(���s�u�����N)
    sql = sql & "  ,GNWKNTCB NOWPROC"   '���ݍH��
    sql = sql & "  ,'     ' LPKRPROCCD" '�ŏI�ʉߊǗ��H��(���s�u�����N)
    sql = sql & "  ,NEWKNTCB LASTPASS"  '�ŏI�ʉߍH��
    sql = sql & "  ,LIVKCB DELCLS"      '�����敪
    sql = sql & "  ,LSTCCB LSTATCLS"    '�ŏI��ԋ敪
    sql = sql & "  ,SHOLDCLSCB HOLDCLS" '�z�[���h�敪
    sql = sql & "  ,HINBCB HINBAN"      '�i��
    sql = sql & "  ,REVNUMCB REVNUM"    '���i�ԍ������ԍ�
    sql = sql & "  ,FACTORYCB FACTORY"  '�H��
    sql = sql & "  ,OPECB OPECOND"      '���Ə���
    sql = sql & "  ,FURYCCB BDCAUS"     '�s�Ǘ��R
    sql = sql & "  ,MAICB COUNT"        '����
    sql = sql & "  ,TDAYCB REGDATE"     '�o�^���t
    sql = sql & "  ,KDAYCB UPDDATE"     '�X�V���t
    sql = sql & "  ,' ' SUMMITSENDFLAG" 'SUMMIT���M�t���O(���g�p)
    sql = sql & "  ,SNDKCB SENDFLAG"    '���M�t���O
    sql = sql & "  ,SNDAYCB SENDDATE"   '���M���t
    sql = sql & "  ,' ' PASSFLAG"       'PASSFLAG(���g�p)
    sql = sql & " FROM"
    sql = sql & "   XSDCB"
    ''���ύXSTART SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME042 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            ''���ύXSTART SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
'            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
'            .INGOTPOS = rs("INGOTPOS")       ' �������J�n�ʒu
'            .LENGTH = rs("LENGTH")           ' ����
'            .SXLID = rs("SXLID")             ' SXLID
'            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H��
'            .NOWPROC = rs("NOWPROC")         ' ���ݍH��
'            .LPKRPROCCD = rs("LPKRPROCCD")   ' �ŏI�ʉߊǗ��H��
'            .LASTPASS = rs("LASTPASS")       ' �ŏI�ʉߍH��
'            .DELCLS = rs("DELCLS")           ' �폜�敪
'            .LSTATCLS = rs("LSTATCLS")       ' �ŏI��ԋ敪
'            .HOLDCLS = rs("HOLDCLS")         ' �z�[���h�敪
'            .hinban = rs("HINBAN")           ' �i��
'            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
'            .factory = rs("FACTORY")         ' �H��
'            .opecond = rs("OPECOND")         ' ���Ə���
'            .BDCAUS = rs("BDCAUS")           ' �s�Ǘ��R
'            .COUNT = rs("COUNT")             ' ����
'            .REGDATE = rs("REGDATE")         ' �o�^���t
'            .UPDDATE = rs("UPDDATE")         ' �X�V���t
'            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
'            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
'            .SENDDATE = rs("SENDDATE")       ' ���M���t
'            .PASSFLAG = " "   ' �ʉ߃t���O�̃X�y�[�X�N���A '02/04/16 Yam
'            If rs("PASSFLAG") = "1" Then
'                .PASSFLAG = rs("PASSFLAG")   ' �ʉ߃t���O '02/04/05 Yam
'            End If
            If IsNull(rs("CRYNUM")) = False Then .CRYNUM = rs("CRYNUM")     ' �����ԍ�
            If IsNull(rs("INGOTPOS")) = False Then .INGOTPOS = rs("INGOTPOS")     ' �������J�n�ʒu
            If IsNull(rs("LENGTH")) = False Then .LENGTH = rs("LENGTH")     ' ����
            If IsNull(rs("SXLID")) = False Then .SXLID = rs("SXLID")     ' SXLID
            If IsNull(rs("KRPROCCD")) = False Then .KRPROCCD = rs("KRPROCCD")     ' �Ǘ��H��
            If IsNull(rs("NOWPROC")) = False Then .NOWPROC = rs("NOWPROC")     ' ���ݍH��
            If IsNull(rs("LPKRPROCCD")) = False Then .LPKRPROCCD = rs("LPKRPROCCD")     ' �ŏI�ʉߊǗ��H��
            If IsNull(rs("LASTPASS")) = False Then .LASTPASS = rs("LASTPASS")     ' �ŏI�ʉߍH��
            If IsNull(rs("DELCLS")) = False Then .DELCLS = rs("DELCLS")     ' �폜�敪
            If IsNull(rs("LSTATCLS")) = False Then .LSTATCLS = rs("LSTATCLS")     ' �ŏI��ԋ敪
            If IsNull(rs("HOLDCLS")) = False Then .HOLDCLS = rs("HOLDCLS")     ' �z�[���h�敪
            If IsNull(rs("HINBAN")) = False Then .hinban = rs("HINBAN")     ' �i��
            If IsNull(rs("REVNUM")) = False Then .REVNUM = rs("REVNUM")     ' ���i�ԍ������ԍ�
            If IsNull(rs("FACTORY")) = False Then .factory = rs("FACTORY")     ' �H��
            If IsNull(rs("OPECOND")) = False Then .opecond = rs("OPECOND")     ' ���Ə���
            If IsNull(rs("BDCAUS")) = False Then .BDCAUS = rs("BDCAUS")     ' �s�Ǘ��R
            If IsNull(rs("COUNT")) = False Then .Count = rs("COUNT")     ' ����
            If IsNull(rs("REGDATE")) = False Then .REGDATE = rs("REGDATE")     ' �o�^���t
            If IsNull(rs("UPDDATE")) = False Then .UPDDATE = rs("UPDDATE")     ' �X�V���t
            If IsNull(rs("SUMMITSENDFLAG")) = False Then .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")     '
            If IsNull(rs("SENDFLAG")) = False Then .SENDFLAG = rs("SENDFLAG")     ' ���M�t���O
            If IsNull(rs("SENDDATE")) = False Then .SENDDATE = rs("SENDDATE")     ' ���M���t
            .PASSFLAG = " "   ' �ʉ߃t���O�̃X�y�[�X�N���A '02/04/16 Yam
            If rs("PASSFLAG") = "1" Then
                If IsNull(rs("PASSFLAG")) = False Then .PASSFLAG = rs("PASSFLAG")     ' �ʉ߃t���O
            End If
            ''���ύXSTART SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME042 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uXSDCW�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_XSDCW    ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺  (2002/07 s_cmzcTBCME044_SQL.bas���ړ�)
Public Function DBDRV_GetTBCME044(records() As typ_XSDCW, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql     As String       'SQL�S��
Dim sqlBase As String       'SQL��{��(WHERE�߂̑O�܂�)
Dim rs      As OraDynaset   'RecordSet
Dim recCnt  As Long         '���R�[�h��
Dim i       As Long

    ''SQL��g�ݗ��Ă�
    'GD�ǉ��@05/01/31 ooba
    '2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
    sqlBase = "Select SXLIDCW, SMPKBNCW, TBKBNCW, REPSMPLIDCW, XTALCW,INPOSCW ,HINBCW, REVNUMCW, FACTORYCW, OPECW, KTKBNCW, " & _
              " SMCRYNUMCW, WFSMPLIDRSCW, NVL(WFSMPLIDRS1CW,'0') as RS1, NVL(WFSMPLIDRS2CW,'0') as RS2, WFINDRSCW, WFRESRS1CW, WFRESRS2CW, WFSMPLIDOICW, WFINDOICW, " & _
              " WFRESOICW, WFSMPLIDB1CW, WFINDB1CW, WFRESB1CW, WFSMPLIDB2CW, WFINDB2CW, WFRESB2CW, WFSMPLIDB3CW, WFINDB3CW, " & _
              " WFRESB3CW, WFSMPLIDL1CW, WFINDL1CW, WFRESL1CW, WFSMPLIDL2CW, WFINDL2CW, WFRESL2CW, WFSMPLIDL3CW, WFINDL3CW, WFRESL3CW, " & _
              " WFSMPLIDL4CW, WFINDL4CW, WFRESL4CW, WFSMPLIDDSCW, WFINDDSCW, WFRESDSCW, WFSMPLIDDZCW, WFINDDZCW, WFRESDZCW, " & _
              " WFSMPLIDSPCW, WFINDSPCW, WFRESSPCW, WFSMPLIDDO1CW, WFINDDO1CW, WFRESDO1CW, WFSMPLIDDO2CW, WFINDDO2CW, WFRESDO2CW, " & _
              " WFSMPLIDDO3CW, WFINDDO3CW, WFRESDO3CW, WFSMPLIDOT1CW, NVL(WFINDOT1CW,'0') as DOT1, NVL(WFRESOT1CW,'0') as SOT1, " & _
              " WFSMPLIDOT2CW, NVL(WFINDOT2CW,'0') as DOT2, NVL(WFRESOT2CW,'0') as SOT2, NVL(WFSMPLIDAOICW,'0') as sAOI, NVL(WFINDAOICW,'0') as iAOI, NVL(WFRESAOICW,'0') as rAOI, NVL(SMPLNUMCW,'0') sNUM, " & _
              " NVL(SMPLPATCW,'0') as PAT, NVL(TSTAFFCW,'0') as STF, TDAYCW, NVL(KSTAFFCW,'0') as kSTF, KDAYCW, NVL(SNDKCW,'0') as SND, NVL(SNDDAYCW,'2003/09/18') as sDAY, " & _
              " WFSMPLIDGDCW, WFINDGDCW, WFRESGDCW, WFHSGDCW, " & _
              " EPSMPLIDB1CW, NVL(EPINDB1CW,'0') as EPINDB1CW, NVL(EPRESB1CW,'0') as EPRESB1CW," & _
              " EPSMPLIDB2CW, NVL(EPINDB2CW,'0') as EPINDB2CW, NVL(EPRESB2CW,'0') as EPRESB2CW," & _
              " EPSMPLIDB3CW, NVL(EPINDB3CW,'0') as EPINDB3CW, NVL(EPRESB3CW,'0') as EPRESB3CW," & _
              " EPSMPLIDL1CW, NVL(EPINDL1CW,'0') as EPINDL1CW, NVL(EPRESL1CW,'0') as EPRESL1CW," & _
              " EPSMPLIDL2CW, NVL(EPINDL2CW,'0') as EPINDL2CW, NVL(EPRESL2CW,'0') as EPRESL2CW," & _
              " EPSMPLIDL3CW, NVL(EPINDL3CW,'0') as EPINDL3CW, NVL(EPRESL3CW,'0') as EPRESL3CW "
    sqlBase = sqlBase & "From XSDCW"
    sql = sqlBase
'    sql = sql & "WHERE XTALCW =" & sqlOrder & " ORDER BY INPOSCW"
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME044 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
'        With records(i)
'            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
'            .INGOTPOS = rs("INGOTPOS")       ' �������ʒu
'            .SMPKBN = rs("SMPKBN")           ' �T���v���敪
'            .SMPLID = rs("SMPLID")           ' �T���v��ID
'            .hinban = rs("HINBAN")           ' �i��
'            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
'            .factory = rs("FACTORY")         ' �H��
'            .opecond = rs("OPECOND")         ' ���Ə���
'            .KTKBN = rs("KTKBN")             ' �m��敪
'            .WFINDRS = rs("WFINDRS")         ' WF�����w���iRs)
'            .WFINDOI = rs("WFINDOI")         ' WF�����w���iOi)
'            .WFINDB1 = rs("WFINDB1")         ' WF�����w���iB1)
'            .WFINDB2 = rs("WFINDB2")         ' WF�����w���iB2�j
'            .WFINDB3 = rs("WFINDB3")         ' WF�����w���iB3)
'            .WFINDL1 = rs("WFINDL1")         ' WF�����w���iL1)
'            .WFINDL2 = rs("WFINDL2")         ' WF�����w���iL2)
'            .WFINDL3 = rs("WFINDL3")         ' WF�����w���iL3)
'            .WFINDL4 = rs("WFINDL4")         ' WF�����w���iL4)
'            .WFINDDS = rs("WFINDDS")         ' WF�����w���iDS)
'            .WFINDDZ = rs("WFINDDZ")         ' WF�����w���iDZ)
'            .WFINDSP = rs("WFINDSP")         ' WF�����w���iSP)
'            .WFINDDO1 = rs("WFINDDO1")       ' WF�����w���iDO1)
'            .WFINDDO2 = rs("WFINDDO2")       ' WF�����w���iDO2)
'            .WFINDDO3 = rs("WFINDDO3")       ' WF�����w���iDO3)
'            '#####################################################03/05/23 �㓡
'            .WFINDOT1 = rs("DOT1")       ' WF�����w���iOT1)
'            .WFINDOT2 = rs("DOT2")       ' WF�����w���iOT2)
'            '#####################################################03/05/23
'            .WFRESRS = rs("WFRESRS")         ' WF�������сiRs)
'            .WFRESOI = rs("WFRESOI")         ' WF�������сiOi)
'            .WFRESB1 = rs("WFRESB1")         ' WF�������сiB1)
'            .WFRESB2 = rs("WFRESB2")         ' WF�������сiB2�j
'            .WFRESB3 = rs("WFRESB3")         ' WF�������сiB3)
'            .WFRESL1 = rs("WFRESL1")         ' WF�������сiL1)
'            .WFRESL2 = rs("WFRESL2")         ' WF�������сiL2)
'            .WFRESL3 = rs("WFRESL3")         ' WF�������сiL3)
'            .WFRESL4 = rs("WFRESL4")         ' WF�������сiL4)
'            .WFRESDS = rs("WFRESDS")         ' WF�������сiDS)
'            .WFRESDZ = rs("WFRESDZ")         ' WF�������сiDZ)
'            .WFRESSP = rs("WFRESSP")         ' WF�������сiSP)
'            .WFRESDO1 = rs("WFRESDO1")       ' WF�������сiDO1)
'            .WFRESDO2 = rs("WFRESDO2")       ' WF�������сiDO2)
'            .WFRESDO3 = rs("WFRESDO3")       ' WF�������сiDO3)
'            '#####################################################03/05/23 �㓡
'            .WFRESOT1 = rs("SOT1")       ' WF�������сiOT1)
'            .WFRESOT2 = rs("SOT2")       ' WF�������сiOT2)
'            '#####################################################03/05/23 �㓡
'            .REGDATE = rs("REGDATE")         ' �o�^���t
'            .UPDDATE = rs("UPDDATE")         ' �X�V���t
'            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
'            .SENDDATE = rs("SENDDATE")       ' ���M���t
'        End With

        With records(i)
'''''                    .SXLIDCW = rs!SXLIDCW
'''''                    .SMPKBNCW = rs!SMPKBNCW
'''''                    .TBKBNCW = rs!TBKBNCW
'''''                    .REPSMPLIDCW = rs!REPSMPLIDCW
'''''                    .XTALCW = rs!XTALCW
'''''                    .INPOSCW = rs!INPOSCW
'''''                    .HINBCW = rs!HINBCW
'''''                    .REVNUMCW = rs!REVNUMCW
'''''                    .FACTORYCW = rs!FACTORYCW
'''''                    .OPECW = rs!OPECW
'''''                    .KTKBNCW = rs!KTKBNCW
'''''                    .SMCRYNUMCW = rs!SMCRYNUMCW
'''''                    .WFSMPLIDRSCW = rs!WFSMPLIDRSCW
'''''                    .WFSMPLIDRS1CW = rs!RS1
'''''                    .WFSMPLIDRS2CW = rs!rs2
'''''                    .WFINDRSCW = rs!WFINDRSCW
'''''                    .WFRESRS1CW = rs!WFRESRS1CW
'''''                    .WFSMPLIDOICW = rs!WFSMPLIDOICW
'''''                    .WFINDOICW = rs!WFINDOICW
'''''                    .WFRESOICW = rs!WFRESOICW
'''''                    .WFSMPLIDB1CW = rs!WFSMPLIDB1CW
'''''                    .WFINDB1CW = rs!WFINDB1CW
'''''                    .WFRESB1CW = rs!WFRESB1CW
'''''                    .WFSMPLIDB2CW = rs!WFSMPLIDB2CW
'''''                    .WFINDB2CW = rs!WFINDB2CW
'''''                    .WFRESB2CW = rs!WFRESB2CW
'''''                    .WFSMPLIDB3CW = rs!WFSMPLIDB3CW
'''''                    .WFINDB3CW = rs!WFINDB3CW
'''''                    .WFRESB3CW = rs!WFRESB3CW
'''''                    .WFSMPLIDL1CW = rs!WFSMPLIDL1CW
'''''                    .WFINDL1CW = rs!WFINDL1CW
'''''                    .WFRESL1CW = rs!WFRESL1CW
'''''                    .WFSMPLIDL2CW = rs!WFSMPLIDL2CW
'''''                    .WFINDL2CW = rs!WFINDL2CW
'''''                    .WFRESL2CW = rs!WFRESL2CW
'''''                    .WFSMPLIDL3CW = rs!WFSMPLIDL3CW
'''''                    .WFINDL3CW = rs!WFINDL3CW
'''''                    .WFRESL3CW = rs!WFRESL3CW
'''''                    .WFSMPLIDL4CW = rs!WFSMPLIDL4CW
'''''                    .WFINDL4CW = rs!WFINDL4CW
'''''                    .WFRESL4CW = rs!WFRESL4CW
'''''                    .WFSMPLIDDSCW = rs!WFSMPLIDDSCW
'''''                    .WFINDDSCW = rs!WFINDDSCW
'''''                    .WFRESDSCW = rs!WFRESDSCW
'''''                    .WFSMPLIDDZCW = rs!WFSMPLIDDZCW
'''''                    .WFINDDZCW = rs!WFINDDZCW
'''''                    .WFRESDZCW = rs!WFRESDZCW
'''''                    .WFSMPLIDSPCW = rs!WFSMPLIDSPCW
'''''                    .WFINDSPCW = rs!WFINDSPCW
'''''                    .WFRESSPCW = rs!WFRESSPCW
'''''                    .WFSMPLIDDO1CW = rs!WFSMPLIDDO1CW
'''''                    .WFINDDO1CW = rs!WFINDDO1CW
'''''                    .WFRESDO1CW = rs!WFRESDO1CW
'''''                    .WFSMPLIDDO2CW = rs!WFSMPLIDDO2CW
'''''                    .WFINDDO2CW = rs!WFINDDO2CW
'''''                    .WFRESDO2CW = rs!WFRESDO2CW
'''''                    .WFSMPLIDDO3CW = rs!WFSMPLIDDO3CW
'''''                    .WFINDDO3CW = rs!WFINDDO3CW
'''''                    .WFRESDO3CW = rs!WFRESDO3CW
'''''                    .WFSMPLIDOT1CW = rs!WFSMPLIDOT1CW
'''''                    .WFINDOT1CW = rs!DOT1
'''''                    .WFRESOT1CW = rs!sOT1
'''''                    .WFSMPLIDOT2CW = rs!WFSMPLIDOT2CW
'''''                    .WFINDOT2CW = rs!DOT2
'''''                    .WFRESOT2CW = rs!sOT2
'''''''                    tHin.hinban = .hinban
'''''''                    tHin.factory = .factory
'''''''                    tHin.mnorevno = .REVNUM
'''''''                    tHin.opecond = .opecond
''''''                    rtn = scmzc_getE036(tHin, sOT1, sOT2)
''''''                    If rtn = FUNCTION_RETURN_FAILURE Then
''''''                        rs.Close
''''''                        DVDRV_KENSA_KOUMOKU = FUNCTION_RETURN_FAILURE
''''''                        GoTo proc_exit
''''''                    End If
''''''                    If sOT1 = "1" Then
''''''                        .WFINDOT1CW = rs!DOT1 '03/05/26
''''''                    Else
''''''                        .WFINDOT1CW = 0 '03/05/26
''''''                    End If
''''''                    If sOT2 = "1" Then
''''''                        .WFINDOT2CW = rs!DOT2 '03/05/26
''''''                    Else
''''''                        .WFINDOT2CW = 0 '03/05/26
''''''                    End If
'''''                    .WFSMPLIDAOICW = rs!sAOI
'''''                    .WFINDAOICW = rs!iAOI
'''''                    .WFRESAOICW = rs!rAOI
'''''                    .SMPLNUMCW = rs!sNUM
'''''                    .SMPLPATCW = rs!PAT
'''''                    .TSTAFFCW = rs!STF
'''''                    .TDAYCW = rs!TDAYCW
'''''                    .KSTAFFCW = rs!kSTF
'''''                    .KDAYCW = rs!KDAYCW
'''''                    .SNDKCW = rs!SND
'''''                    .SNDDAYCW = rs!sDAY

                    If IsNull(rs!SXLIDCW) = False Then .SXLIDCW = rs!SXLIDCW
                    If IsNull(rs!SMPKBNCW) = False Then .SMPKBNCW = rs!SMPKBNCW
                    If IsNull(rs!TBKBNCW) = False Then .TBKBNCW = rs!TBKBNCW
                    If IsNull(rs!REPSMPLIDCW) = False Then .REPSMPLIDCW = rs!REPSMPLIDCW
                    If IsNull(rs!XTALCW) = False Then .XTALCW = rs!XTALCW
                    If IsNull(rs!INPOSCW) = False Then .INPOSCW = rs!INPOSCW
                    If IsNull(rs!HINBCW) = False Then .HINBCW = rs!HINBCW
                    If IsNull(rs!REVNUMCW) = False Then .REVNUMCW = rs!REVNUMCW
                    If IsNull(rs!FACTORYCW) = False Then .FACTORYCW = rs!FACTORYCW
                    If IsNull(rs!OPECW) = False Then .OPECW = rs!OPECW
                    If IsNull(rs!KTKBNCW) = False Then .KTKBNCW = rs!KTKBNCW
                    If IsNull(rs!SMCRYNUMCW) = False Then .SMCRYNUMCW = rs!SMCRYNUMCW
                    If IsNull(rs!WFSMPLIDRSCW) = False Then .WFSMPLIDRSCW = rs!WFSMPLIDRSCW
                    If IsNull(rs!rs1) = False Then .WFSMPLIDRS1CW = rs!rs1
                    If IsNull(rs!rs2) = False Then .WFSMPLIDRS2CW = rs!rs2
                    If IsNull(rs!WFINDRSCW) = False Then .WFINDRSCW = rs!WFINDRSCW
                    If IsNull(rs!WFRESRS1CW) = False Then .WFRESRS1CW = rs!WFRESRS1CW
                    If IsNull(rs!WFSMPLIDOICW) = False Then .WFSMPLIDOICW = rs!WFSMPLIDOICW
                    If IsNull(rs!WFINDOICW) = False Then .WFINDOICW = rs!WFINDOICW
                    If IsNull(rs!WFRESOICW) = False Then .WFRESOICW = rs!WFRESOICW
                    If IsNull(rs!WFSMPLIDB1CW) = False Then .WFSMPLIDB1CW = rs!WFSMPLIDB1CW
                    If IsNull(rs!WFINDB1CW) = False Then .WFINDB1CW = rs!WFINDB1CW
                    If IsNull(rs!WFRESB1CW) = False Then .WFRESB1CW = rs!WFRESB1CW
                    If IsNull(rs!WFSMPLIDB2CW) = False Then .WFSMPLIDB2CW = rs!WFSMPLIDB2CW
                    If IsNull(rs!WFINDB2CW) = False Then .WFINDB2CW = rs!WFINDB2CW
                    If IsNull(rs!WFRESB2CW) = False Then .WFRESB2CW = rs!WFRESB2CW
                    If IsNull(rs!WFSMPLIDB3CW) = False Then .WFSMPLIDB3CW = rs!WFSMPLIDB3CW
                    If IsNull(rs!WFINDB3CW) = False Then .WFINDB3CW = rs!WFINDB3CW
                    If IsNull(rs!WFRESB3CW) = False Then .WFRESB3CW = rs!WFRESB3CW
                    If IsNull(rs!WFSMPLIDL1CW) = False Then .WFSMPLIDL1CW = rs!WFSMPLIDL1CW
                    If IsNull(rs!WFINDL1CW) = False Then .WFINDL1CW = rs!WFINDL1CW
                    If IsNull(rs!WFRESL1CW) = False Then .WFRESL1CW = rs!WFRESL1CW
                    If IsNull(rs!WFSMPLIDL2CW) = False Then .WFSMPLIDL2CW = rs!WFSMPLIDL2CW
                    If IsNull(rs!WFINDL2CW) = False Then .WFINDL2CW = rs!WFINDL2CW
                    If IsNull(rs!WFRESL2CW) = False Then .WFRESL2CW = rs!WFRESL2CW
                    If IsNull(rs!WFSMPLIDL3CW) = False Then .WFSMPLIDL3CW = rs!WFSMPLIDL3CW
                    If IsNull(rs!WFINDL3CW) = False Then .WFINDL3CW = rs!WFINDL3CW
                    If IsNull(rs!WFRESL3CW) = False Then .WFRESL3CW = rs!WFRESL3CW
                    If IsNull(rs!WFSMPLIDL4CW) = False Then .WFSMPLIDL4CW = rs!WFSMPLIDL4CW
                    If IsNull(rs!WFINDL4CW) = False Then .WFINDL4CW = rs!WFINDL4CW
                    If IsNull(rs!WFRESL4CW) = False Then .WFRESL4CW = rs!WFRESL4CW
                    If IsNull(rs!WFSMPLIDDSCW) = False Then .WFSMPLIDDSCW = rs!WFSMPLIDDSCW
                    If IsNull(rs!WFINDDSCW) = False Then .WFINDDSCW = rs!WFINDDSCW
                    If IsNull(rs!WFRESDSCW) = False Then .WFRESDSCW = rs!WFRESDSCW
                    If IsNull(rs!WFSMPLIDDZCW) = False Then .WFSMPLIDDZCW = rs!WFSMPLIDDZCW
                    If IsNull(rs!WFINDDZCW) = False Then .WFINDDZCW = rs!WFINDDZCW
                    If IsNull(rs!WFRESDZCW) = False Then .WFRESDZCW = rs!WFRESDZCW
                    If IsNull(rs!WFSMPLIDSPCW) = False Then .WFSMPLIDSPCW = rs!WFSMPLIDSPCW
                    If IsNull(rs!WFINDSPCW) = False Then .WFINDSPCW = rs!WFINDSPCW
                    If IsNull(rs!WFRESSPCW) = False Then .WFRESSPCW = rs!WFRESSPCW
                    If IsNull(rs!WFSMPLIDDO1CW) = False Then .WFSMPLIDDO1CW = rs!WFSMPLIDDO1CW
                    If IsNull(rs!WFINDDO1CW) = False Then .WFINDDO1CW = rs!WFINDDO1CW
                    If IsNull(rs!WFRESDO1CW) = False Then .WFRESDO1CW = rs!WFRESDO1CW
                    If IsNull(rs!WFSMPLIDDO2CW) = False Then .WFSMPLIDDO2CW = rs!WFSMPLIDDO2CW
                    If IsNull(rs!WFINDDO2CW) = False Then .WFINDDO2CW = rs!WFINDDO2CW
                    If IsNull(rs!WFRESDO2CW) = False Then .WFRESDO2CW = rs!WFRESDO2CW
                    If IsNull(rs!WFSMPLIDDO3CW) = False Then .WFSMPLIDDO3CW = rs!WFSMPLIDDO3CW
                    If IsNull(rs!WFINDDO3CW) = False Then .WFINDDO3CW = rs!WFINDDO3CW
                    If IsNull(rs!WFRESDO3CW) = False Then .WFRESDO3CW = rs!WFRESDO3CW
                    If IsNull(rs!WFSMPLIDOT1CW) = False Then .WFSMPLIDOT1CW = rs!WFSMPLIDOT1CW
                    If IsNull(rs!DOT1) = False Then .WFINDOT1CW = rs!DOT1
                    If IsNull(rs!sOT1) = False Then .WFRESOT1CW = rs!sOT1
                    If IsNull(rs!WFSMPLIDOT2CW) = False Then .WFSMPLIDOT2CW = rs!WFSMPLIDOT2CW
                    If IsNull(rs!DOT2) = False Then .WFINDOT2CW = rs!DOT2
                    If IsNull(rs!sOT2) = False Then .WFRESOT2CW = rs!sOT2

                    If IsNull(rs!sAOI) = False Then .WFSMPLIDAOICW = rs!sAOI
                    If IsNull(rs!iAOI) = False Then .WFINDAOICW = rs!iAOI
                    If IsNull(rs!rAOI) = False Then .WFRESAOICW = rs!rAOI
                    If IsNull(rs!sNum) = False Then .SMPLNUMCW = rs!sNum
                    If IsNull(rs!PAT) = False Then .SMPLPATCW = rs!PAT
                    If IsNull(rs!STF) = False Then .TSTAFFCW = rs!STF
                    If IsNull(rs!TDAYCW) = False Then .TDAYCW = rs!TDAYCW
                    If IsNull(rs!kSTF) = False Then .KSTAFFCW = rs!kSTF
                    If IsNull(rs!KDAYCW) = False Then .KDAYCW = rs!KDAYCW
                    If IsNull(rs!SND) = False Then .SNDKCW = rs!SND
                    If IsNull(rs!sDay) = False Then .SNDDAYCW = rs!sDay

                    '' GD�ǉ��@05/01/31 ooba START ===========================================>
                    If IsNull(rs!WFSMPLIDGDCW) = False Then .WFSMPLIDGDCW = rs!WFSMPLIDGDCW
                    If IsNull(rs!WFINDGDCW) = False Then .WFINDGDCW = rs!WFINDGDCW
                    If IsNull(rs!WFRESGDCW) = False Then .WFRESGDCW = rs!WFRESGDCW
                    If IsNull(rs!WFHSGDCW) = False Then .WFHSGDCW = rs!WFHSGDCW
                    '' GD�ǉ��@05/01/31 ooba END =============================================>

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                    If IsNull(rs!EPSMPLIDB1CW) = False Then .EPSMPLIDB1CW = rs!EPSMPLIDB1CW
                    If IsNull(rs!EPINDB1CW) = False Then .EPINDB1CW = rs!EPINDB1CW
                    If IsNull(rs!EPRESB1CW) = False Then .EPRESB1CW = rs!EPRESB1CW
                    If IsNull(rs!EPSMPLIDB2CW) = False Then .EPSMPLIDB2CW = rs!EPSMPLIDB2CW
                    If IsNull(rs!EPINDB2CW) = False Then .EPINDB2CW = rs!EPINDB2CW
                    If IsNull(rs!EPRESB2CW) = False Then .EPRESB2CW = rs!EPRESB2CW
                    If IsNull(rs!EPSMPLIDB3CW) = False Then .EPSMPLIDB3CW = rs!EPSMPLIDB3CW
                    If IsNull(rs!EPINDB3CW) = False Then .EPINDB3CW = rs!EPINDB3CW
                    If IsNull(rs!EPRESB3CW) = False Then .EPRESB3CW = rs!EPRESB3CW
                    If IsNull(rs!EPSMPLIDL1CW) = False Then .EPSMPLIDL1CW = rs!EPSMPLIDL1CW
                    If IsNull(rs!EPINDL1CW) = False Then .EPINDL1CW = rs!EPINDL1CW
                    If IsNull(rs!EPRESL1CW) = False Then .EPRESL1CW = rs!EPRESL1CW
                    If IsNull(rs!EPSMPLIDL2CW) = False Then .EPSMPLIDL2CW = rs!EPSMPLIDL2CW
                    If IsNull(rs!EPINDL2CW) = False Then .EPINDL2CW = rs!EPINDL2CW
                    If IsNull(rs!EPRESL2CW) = False Then .EPRESL2CW = rs!EPRESL2CW
                    If IsNull(rs!EPSMPLIDL3CW) = False Then .EPSMPLIDL3CW = rs!EPSMPLIDL3CW
                    If IsNull(rs!EPINDL3CW) = False Then .EPINDL3CW = rs!EPINDL3CW
                    If IsNull(rs!EPRESL3CW) = False Then .EPRESL3CW = rs!EPRESL3CW
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

                End With

        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME044 = FUNCTION_RETURN_SUCCESS
End Function

'�T�v      :XSDCW����ް��擾
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :sCryNum       ,I  ,String       ,�����ԍ�
'          :records()     ,O  ,typ_XSDCW    ,���o���R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :08/02/04 ooba
Public Function DBDRV_GetXSDCW(sCryNum As String, records() As typ_XSDCW) As FUNCTION_RETURN

    Dim sql     As String       'SQL�S��
    Dim rs      As OraDynaset   'RecordSet
    Dim recCnt  As Long         'ں��ސ�
    Dim i       As Long

    sql = "SELECT "
    sql = sql & "SXLIDCB, "
    sql = sql & "SXLIDCW, "
    sql = sql & "NVL(SMPKBNCW,'T') as SMPKBNCW, "
    sql = sql & "NVL(TBKBNCW,'T') as TBKBN, "
    sql = sql & "NVL(REPSMPLIDCW,' ') as REPSMPLIDCW, "
    sql = sql & "NVL(XTALCB,' ') as XTALCB, "
    sql = sql & "NVL(INGOTPOS,0) as INGOTPOS, "
    sql = sql & "NVL(INPOSCW,0), "
    sql = sql & "NVL(HINBCB,' ') as HINBCB, "
    sql = sql & "NVL(REVNUMCB,0) as REVNUMCB, "
    sql = sql & "NVL(FACTORYCB,' ') as FACTORYCB, "
    sql = sql & "NVL(OPECB,' ') as OPECB, "
    sql = sql & "NVL(HINBCW,' ') as HINBCW, "
    sql = sql & "NVL(REVNUMCW,0) as REVNUMCW, "
    sql = sql & "NVL(FACTORYCW,' ') as FACTORYCW, "
    sql = sql & "NVL(OPECW,' ') as OPECW, "
    sql = sql & "NVL(KTKBNCW,' ') as KTKBNCW, "
    sql = sql & "NVL(SMCRYNUMCW,' ') as SMCRYNUMCW, "
    sql = sql & "NVL(WFSMPLIDRSCW,' ') as WFSMPLIDRSCW, "
    sql = sql & "NVL(WFSMPLIDRS1CW,' ') as WFSMPLIDRS1CW, "
    sql = sql & "NVL(WFSMPLIDRS2CW,' ') as WFSMPLIDRS2CW, "
    sql = sql & "NVL(WFINDRSCW,'0') as WFINDRSCW, "
    sql = sql & "NVL(WFRESRS1CW,'0') as WFRESRS1CW, "
    sql = sql & "NVL(WFRESRS2CW,'0') as WFRESRS2CW, "
    sql = sql & "NVL(WFSMPLIDOICW,' ') as WFSMPLIDOICW, "
    sql = sql & "NVL(WFINDOICW,'0') as WFINDOICW, "
    sql = sql & "NVL(WFRESOICW,'0') as WFRESOICW, "
    sql = sql & "NVL(WFSMPLIDB1CW,' ') as WFSMPLIDB1CW, "
    sql = sql & "NVL(WFINDB1CW,'0') as WFINDB1CW, "
    sql = sql & "NVL(WFRESB1CW,'0') as WFRESB1CW, "
    sql = sql & "NVL(WFSMPLIDB2CW,' ') as WFSMPLIDB2CW, "
    sql = sql & "NVL(WFINDB2CW,'0') as WFINDB2CW, "
    sql = sql & "NVL(WFRESB2CW,'0') as WFRESB2CW, "
    sql = sql & "NVL(WFSMPLIDB3CW,' ') as WFSMPLIDB3CW, "
    sql = sql & "NVL(WFINDB3CW,'0') as WFINDB3CW, "
    sql = sql & "NVL(WFRESB3CW,'0') as WFRESB3CW, "
    sql = sql & "NVL(WFSMPLIDL1CW,' ') as WFSMPLIDL1CW, "
    sql = sql & "NVL(WFINDL1CW,'0') as WFINDL1CW, "
    sql = sql & "NVL(WFRESL1CW,'0') as WFRESL1CW, "
    sql = sql & "NVL(WFSMPLIDL2CW,' ') as WFSMPLIDL2CW, "
    sql = sql & "NVL(WFINDL2CW,'0') as WFINDL2CW, "
    sql = sql & "NVL(WFRESL2CW,'0') as WFRESL2CW, "
    sql = sql & "NVL(WFSMPLIDL3CW,' ') as WFSMPLIDL3CW, "
    sql = sql & "NVL(WFINDL3CW,'0') as WFINDL3CW, "
    sql = sql & "NVL(WFRESL3CW,'0') as WFRESL3CW, "
    sql = sql & "NVL(WFSMPLIDL4CW,' ') as WFSMPLIDL4CW, "
    sql = sql & "NVL(WFINDL4CW,'0') as WFINDL4CW, "
    sql = sql & "NVL(WFRESL4CW,'0') as WFRESL4CW, "
    sql = sql & "NVL(WFSMPLIDDSCW,' ') as WFSMPLIDDSCW, "
    sql = sql & "NVL(WFINDDSCW,'0') as WFINDDSCW, "
    sql = sql & "NVL(WFRESDSCW,'0') as WFRESDSCW, "
    sql = sql & "NVL(WFSMPLIDDZCW,' ') as WFSMPLIDDZCW, "
    sql = sql & "NVL(WFINDDZCW,'0') as WFINDDZCW, "
    sql = sql & "NVL(WFRESDZCW,'0') as WFRESDZCW, "
    sql = sql & "NVL(WFSMPLIDSPCW,' ') as WFSMPLIDSPCW, "
    sql = sql & "NVL(WFINDSPCW,'0') as WFINDSPCW, "
    sql = sql & "NVL(WFRESSPCW,'0') as WFRESSPCW, "
    sql = sql & "NVL(WFSMPLIDDO1CW,' ') as WFSMPLIDDO1CW, "
    sql = sql & "NVL(WFINDDO1CW,'0') as WFINDDO1CW, "
    sql = sql & "NVL(WFRESDO1CW,'0') as WFRESDO1CW, "
    sql = sql & "NVL(WFSMPLIDDO2CW,' ') as WFSMPLIDDO2CW, "
    sql = sql & "NVL(WFINDDO2CW,'0') as WFINDDO2CW, "
    sql = sql & "NVL(WFRESDO2CW,'0') as WFRESDO2CW, "
    sql = sql & "NVL(WFSMPLIDDO3CW,' ') as WFSMPLIDDO3CW, "
    sql = sql & "NVL(WFINDDO3CW,'0') as WFINDDO3CW, "
    sql = sql & "NVL(WFRESDO3CW,'0') as WFRESDO3CW, "
    sql = sql & "NVL(WFSMPLIDOT1CW,' ') as WFSMPLIDOT1CW, "
    sql = sql & "NVL(WFINDOT1CW,'0') as WFINDOT1CW, "
    sql = sql & "NVL(WFRESOT1CW,'0') as WFRESOT1CW, "
    sql = sql & "NVL(WFSMPLIDOT2CW,' ') as WFSMPLIDOT2CW, "
    sql = sql & "NVL(WFINDOT2CW,'0') as WFINDOT2CW, "
    sql = sql & "NVL(WFRESOT2CW,'0') as WFRESOT2CW, "
    sql = sql & "NVL(WFSMPLIDAOICW,' ') as WFSMPLIDAOICW, "
    sql = sql & "NVL(WFINDAOICW,'0') as WFINDAOICW, "
    sql = sql & "NVL(WFRESAOICW,'0') as WFRESAOICW, "
    sql = sql & "NVL(SMPLNUMCW,0) as SMPLNUMCW, "
    sql = sql & "NVL(SMPLPATCW,' ') as SMPLPATCW, "
    sql = sql & "NVL(LIVKCW,'0') as LIVKCW, "
    sql = sql & "NVL(WFSMPLIDGDCW,' ') as WFSMPLIDGDCW, "
    sql = sql & "NVL(WFINDGDCW,'0') as WFINDGDCW, "
    sql = sql & "NVL(WFRESGDCW,'0') as WFRESGDCW, "
    sql = sql & "NVL(WFHSGDCW,'0') as WFHSGDCW, "
    sql = sql & "NVL(EPSMPLIDB1CW,' ') as EPSMPLIDB1CW, "
    sql = sql & "NVL(EPINDB1CW,'0') as EPINDB1CW, "
    sql = sql & "NVL(EPRESB1CW,'0') as EPRESB1CW, "
    sql = sql & "NVL(EPSMPLIDB2CW,' ') as EPSMPLIDB2CW, "
    sql = sql & "NVL(EPINDB2CW,'0') as EPINDB2CW, "
    sql = sql & "NVL(EPRESB2CW,'0') as EPRESB2CW, "
    sql = sql & "NVL(EPSMPLIDB3CW,' ') as EPSMPLIDB3CW, "
    sql = sql & "NVL(EPINDB3CW,'0') as EPINDB3CW, "
    sql = sql & "NVL(EPRESB3CW,'0') as EPRESB3CW, "
    sql = sql & "NVL(EPSMPLIDL1CW,' ') as EPSMPLIDL1CW, "
    sql = sql & "NVL(EPINDL1CW,'0') as EPINDL1CW, "
    sql = sql & "NVL(EPRESL1CW,'0') as EPRESL1CW, "
    sql = sql & "NVL(EPSMPLIDL2CW,' ') as EPSMPLIDL2CW, "
    sql = sql & "NVL(EPINDL2CW,'0') as EPINDL2CW, "
    sql = sql & "NVL(EPRESL2CW,'0') as EPRESL2CW, "
    sql = sql & "NVL(EPSMPLIDL3CW,' ') as EPSMPLIDL3CW, "
    sql = sql & "NVL(EPINDL3CW,'0') as EPINDL3CW, "
    sql = sql & "NVL(EPRESL3CW,'0') as EPRESL3CW "
    sql = sql & "FROM "
    sql = sql & "    (SELECT SXLIDCB, "
    sql = sql & "     XTALCB, "
    sql = sql & "     INPOSCB as INGOTPOS, "
    sql = sql & "     HINBCB, "
    sql = sql & "     REVNUMCB, "
    sql = sql & "     FACTORYCB, "
    sql = sql & "     OPECB "
    sql = sql & "     FROM XSDCB "
    sql = sql & "     WHERE XTALCB = '" & sCryNum & "' "
    sql = sql & "     AND LIVKCB = '0' "
    sql = sql & "    ), "
    sql = sql & "    (SELECT SXLIDCW, "
    sql = sql & "     SMPKBNCW, "
    sql = sql & "     TBKBNCW, "
    sql = sql & "     REPSMPLIDCW, "
    sql = sql & "     INPOSCW, "
    sql = sql & "     HINBCW, "
    sql = sql & "     REVNUMCW, "
    sql = sql & "     FACTORYCW, "
    sql = sql & "     OPECW, "
    sql = sql & "     KTKBNCW, "
    sql = sql & "     SMCRYNUMCW, "
    sql = sql & "     WFSMPLIDRSCW, "
    sql = sql & "     WFSMPLIDRS1CW, "
    sql = sql & "     WFSMPLIDRS2CW, "
    sql = sql & "     WFINDRSCW, "
    sql = sql & "     WFRESRS1CW, "
    sql = sql & "     WFRESRS2CW, "
    sql = sql & "     WFSMPLIDOICW, "
    sql = sql & "     WFINDOICW, "
    sql = sql & "     WFRESOICW, "
    sql = sql & "     WFSMPLIDB1CW, "
    sql = sql & "     WFINDB1CW, "
    sql = sql & "     WFRESB1CW, "
    sql = sql & "     WFSMPLIDB2CW, "
    sql = sql & "     WFINDB2CW, "
    sql = sql & "     WFRESB2CW, "
    sql = sql & "     WFSMPLIDB3CW, "
    sql = sql & "     WFINDB3CW, "
    sql = sql & "     WFRESB3CW, "
    sql = sql & "     WFSMPLIDL1CW, "
    sql = sql & "     WFINDL1CW, "
    sql = sql & "     WFRESL1CW, "
    sql = sql & "     WFSMPLIDL2CW, "
    sql = sql & "     WFINDL2CW, "
    sql = sql & "     WFRESL2CW, "
    sql = sql & "     WFSMPLIDL3CW, "
    sql = sql & "     WFINDL3CW, "
    sql = sql & "     WFRESL3CW, "
    sql = sql & "     WFSMPLIDL4CW, "
    sql = sql & "     WFINDL4CW, "
    sql = sql & "     WFRESL4CW, "
    sql = sql & "     WFSMPLIDDSCW, "
    sql = sql & "     WFINDDSCW, "
    sql = sql & "     WFRESDSCW, "
    sql = sql & "     WFSMPLIDDZCW, "
    sql = sql & "     WFINDDZCW, "
    sql = sql & "     WFRESDZCW, "
    sql = sql & "     WFSMPLIDSPCW, "
    sql = sql & "     WFINDSPCW, "
    sql = sql & "     WFRESSPCW, "
    sql = sql & "     WFSMPLIDDO1CW, "
    sql = sql & "     WFINDDO1CW, "
    sql = sql & "     WFRESDO1CW, "
    sql = sql & "     WFSMPLIDDO2CW, "
    sql = sql & "     WFINDDO2CW, "
    sql = sql & "     WFRESDO2CW, "
    sql = sql & "     WFSMPLIDDO3CW, "
    sql = sql & "     WFINDDO3CW, "
    sql = sql & "     WFRESDO3CW, "
    sql = sql & "     WFSMPLIDOT1CW, "
    sql = sql & "     WFINDOT1CW, "
    sql = sql & "     WFRESOT1CW, "
    sql = sql & "     WFSMPLIDOT2CW, "
    sql = sql & "     WFINDOT2CW, "
    sql = sql & "     WFRESOT2CW, "
    sql = sql & "     WFSMPLIDAOICW, "
    sql = sql & "     WFINDAOICW, "
    sql = sql & "     WFRESAOICW, "
    sql = sql & "     SMPLNUMCW, "
    sql = sql & "     SMPLPATCW, "
    sql = sql & "     LIVKCW, "
    sql = sql & "     WFSMPLIDGDCW, "
    sql = sql & "     WFINDGDCW, "
    sql = sql & "     WFRESGDCW, "
    sql = sql & "     WFHSGDCW, "
    sql = sql & "     EPSMPLIDB1CW, "
    sql = sql & "     EPINDB1CW, "
    sql = sql & "     EPRESB1CW, "
    sql = sql & "     EPSMPLIDB2CW, "
    sql = sql & "     EPINDB2CW, "
    sql = sql & "     EPRESB2CW, "
    sql = sql & "     EPSMPLIDB3CW, "
    sql = sql & "     EPINDB3CW, "
    sql = sql & "     EPRESB3CW, "
    sql = sql & "     EPSMPLIDL1CW, "
    sql = sql & "     EPINDL1CW, "
    sql = sql & "     EPRESL1CW, "
    sql = sql & "     EPSMPLIDL2CW, "
    sql = sql & "     EPINDL2CW, "
    sql = sql & "     EPRESL2CW, "
    sql = sql & "     EPSMPLIDL3CW, "
    sql = sql & "     EPINDL3CW, "
    sql = sql & "     EPRESL3CW "
    sql = sql & "     FROM XSDCW "
    sql = sql & "     WHERE XTALCW = '" & sCryNum & "' "
    sql = sql & "     AND TBKBNCW = 'T' "
    sql = sql & "    ) "
'    sql = sql & "WHERE INGOTPOS = INPOSCW(+) "
    sql = sql & "WHERE SXLIDCB = SXLIDCW(+) "           '08/07/10 ooba
    sql = sql & "AND NVL(LIVKCW,'0') = '0' "
    
    sql = sql & "UNION ALL "
    
    sql = sql & "SELECT "
    sql = sql & "SXLIDCB, "
    sql = sql & "SXLIDCW, "
    sql = sql & "NVL(SMPKBNCW,'B') as SMPKBNCW, "
    sql = sql & "NVL(TBKBNCW,'B') as TBKBN, "
    sql = sql & "NVL(REPSMPLIDCW,' ') as REPSMPLIDCW, "
    sql = sql & "NVL(XTALCB,' ') as XTALCB, "
    sql = sql & "NVL(INGOTPOS,0) as INGOTPOS, "
    sql = sql & "NVL(INPOSCW,0), "
    sql = sql & "NVL(HINBCB,' ') as HINBCB, "
    sql = sql & "NVL(REVNUMCB,0) as REVNUMCB, "
    sql = sql & "NVL(FACTORYCB,' ') as FACTORYCB, "
    sql = sql & "NVL(OPECB,' ') as OPECB, "
    sql = sql & "NVL(HINBCW,' ') as HINBCW, "
    sql = sql & "NVL(REVNUMCW,0) as REVNUMCW, "
    sql = sql & "NVL(FACTORYCW,' ') as FACTORYCW, "
    sql = sql & "NVL(OPECW,' ') as OPECW, "
    sql = sql & "NVL(KTKBNCW,' ') as KTKBNCW, "
    sql = sql & "NVL(SMCRYNUMCW,' ') as SMCRYNUMCW, "
    sql = sql & "NVL(WFSMPLIDRSCW,' ') as WFSMPLIDRSCW, "
    sql = sql & "NVL(WFSMPLIDRS1CW,' ') as WFSMPLIDRS1CW, "
    sql = sql & "NVL(WFSMPLIDRS2CW,' ') as WFSMPLIDRS2CW, "
    sql = sql & "NVL(WFINDRSCW,'0') as WFINDRSCW, "
    sql = sql & "NVL(WFRESRS1CW,'0') as WFRESRS1CW, "
    sql = sql & "NVL(WFRESRS2CW,'0') as WFRESRS2CW, "
    sql = sql & "NVL(WFSMPLIDOICW,' ') as WFSMPLIDOICW, "
    sql = sql & "NVL(WFINDOICW,'0') as WFINDOICW, "
    sql = sql & "NVL(WFRESOICW,'0') as WFRESOICW, "
    sql = sql & "NVL(WFSMPLIDB1CW,' ') as WFSMPLIDB1CW, "
    sql = sql & "NVL(WFINDB1CW,'0') as WFINDB1CW, "
    sql = sql & "NVL(WFRESB1CW,'0') as WFRESB1CW, "
    sql = sql & "NVL(WFSMPLIDB2CW,' ') as WFSMPLIDB2CW, "
    sql = sql & "NVL(WFINDB2CW,'0') as WFINDB2CW, "
    sql = sql & "NVL(WFRESB2CW,'0') as WFRESB2CW, "
    sql = sql & "NVL(WFSMPLIDB3CW,' ') as WFSMPLIDB3CW, "
    sql = sql & "NVL(WFINDB3CW,'0') as WFINDB3CW, "
    sql = sql & "NVL(WFRESB3CW,'0') as WFRESB3CW, "
    sql = sql & "NVL(WFSMPLIDL1CW,' ') as WFSMPLIDL1CW, "
    sql = sql & "NVL(WFINDL1CW,'0') as WFINDL1CW, "
    sql = sql & "NVL(WFRESL1CW,'0') as WFRESL1CW, "
    sql = sql & "NVL(WFSMPLIDL2CW,' ') as WFSMPLIDL2CW, "
    sql = sql & "NVL(WFINDL2CW,'0') as WFINDL2CW, "
    sql = sql & "NVL(WFRESL2CW,'0') as WFRESL2CW, "
    sql = sql & "NVL(WFSMPLIDL3CW,' ') as WFSMPLIDL3CW, "
    sql = sql & "NVL(WFINDL3CW,'0') as WFINDL3CW, "
    sql = sql & "NVL(WFRESL3CW,'0') as WFRESL3CW, "
    sql = sql & "NVL(WFSMPLIDL4CW,' ') as WFSMPLIDL4CW, "
    sql = sql & "NVL(WFINDL4CW,'0') as WFINDL4CW, "
    sql = sql & "NVL(WFRESL4CW,'0') as WFRESL4CW, "
    sql = sql & "NVL(WFSMPLIDDSCW,' ') as WFSMPLIDDSCW, "
    sql = sql & "NVL(WFINDDSCW,'0') as WFINDDSCW, "
    sql = sql & "NVL(WFRESDSCW,'0') as WFRESDSCW, "
    sql = sql & "NVL(WFSMPLIDDZCW,' ') as WFSMPLIDDZCW, "
    sql = sql & "NVL(WFINDDZCW,'0') as WFINDDZCW, "
    sql = sql & "NVL(WFRESDZCW,'0') as WFRESDZCW, "
    sql = sql & "NVL(WFSMPLIDSPCW,' ') as WFSMPLIDSPCW, "
    sql = sql & "NVL(WFINDSPCW,'0') as WFINDSPCW, "
    sql = sql & "NVL(WFRESSPCW,'0') as WFRESSPCW, "
    sql = sql & "NVL(WFSMPLIDDO1CW,' ') as WFSMPLIDDO1CW, "
    sql = sql & "NVL(WFINDDO1CW,'0') as WFINDDO1CW, "
    sql = sql & "NVL(WFRESDO1CW,'0') as WFRESDO1CW, "
    sql = sql & "NVL(WFSMPLIDDO2CW,' ') as WFSMPLIDDO2CW, "
    sql = sql & "NVL(WFINDDO2CW,'0') as WFINDDO2CW, "
    sql = sql & "NVL(WFRESDO2CW,'0') as WFRESDO2CW, "
    sql = sql & "NVL(WFSMPLIDDO3CW,' ') as WFSMPLIDDO3CW, "
    sql = sql & "NVL(WFINDDO3CW,'0') as WFINDDO3CW, "
    sql = sql & "NVL(WFRESDO3CW,'0') as WFRESDO3CW, "
    sql = sql & "NVL(WFSMPLIDOT1CW,' ') as WFSMPLIDOT1CW, "
    sql = sql & "NVL(WFINDOT1CW,'0') as WFINDOT1CW, "
    sql = sql & "NVL(WFRESOT1CW,'0') as WFRESOT1CW, "
    sql = sql & "NVL(WFSMPLIDOT2CW,' ') as WFSMPLIDOT2CW, "
    sql = sql & "NVL(WFINDOT2CW,'0') as WFINDOT2CW, "
    sql = sql & "NVL(WFRESOT2CW,'0') as WFRESOT2CW, "
    sql = sql & "NVL(WFSMPLIDAOICW,' ') as WFSMPLIDAOICW, "
    sql = sql & "NVL(WFINDAOICW,'0') as WFINDAOICW, "
    sql = sql & "NVL(WFRESAOICW,'0') as WFRESAOICW, "
    sql = sql & "NVL(SMPLNUMCW,0) as SMPLNUMCW, "
    sql = sql & "NVL(SMPLPATCW,' ') as SMPLPATCW, "
    sql = sql & "NVL(LIVKCW,'0') as LIVKCW, "
    sql = sql & "NVL(WFSMPLIDGDCW,' ') as WFSMPLIDGDCW, "
    sql = sql & "NVL(WFINDGDCW,'0') as WFINDGDCW, "
    sql = sql & "NVL(WFRESGDCW,'0') as WFRESGDCW, "
    sql = sql & "NVL(WFHSGDCW,'0') as WFHSGDCW, "
    sql = sql & "NVL(EPSMPLIDB1CW,' ') as EPSMPLIDB1CW, "
    sql = sql & "NVL(EPINDB1CW,'0') as EPINDB1CW, "
    sql = sql & "NVL(EPRESB1CW,'0') as EPRESB1CW, "
    sql = sql & "NVL(EPSMPLIDB2CW,' ') as EPSMPLIDB2CW, "
    sql = sql & "NVL(EPINDB2CW,'0') as EPINDB2CW, "
    sql = sql & "NVL(EPRESB2CW,'0') as EPRESB2CW, "
    sql = sql & "NVL(EPSMPLIDB3CW,' ') as EPSMPLIDB3CW, "
    sql = sql & "NVL(EPINDB3CW,'0') as EPINDB3CW, "
    sql = sql & "NVL(EPRESB3CW,'0') as EPRESB3CW, "
    sql = sql & "NVL(EPSMPLIDL1CW,' ') as EPSMPLIDL1CW, "
    sql = sql & "NVL(EPINDL1CW,'0') as EPINDL1CW, "
    sql = sql & "NVL(EPRESL1CW,'0') as EPRESL1CW, "
    sql = sql & "NVL(EPSMPLIDL2CW,' ') as EPSMPLIDL2CW, "
    sql = sql & "NVL(EPINDL2CW,'0') as EPINDL2CW, "
    sql = sql & "NVL(EPRESL2CW,'0') as EPRESL2CW, "
    sql = sql & "NVL(EPSMPLIDL3CW,' ') as EPSMPLIDL3CW, "
    sql = sql & "NVL(EPINDL3CW,'0') as EPINDL3CW, "
    sql = sql & "NVL(EPRESL3CW,'0') as EPRESL3CW "
    sql = sql & "FROM "
    sql = sql & "    (SELECT SXLIDCB, "
    sql = sql & "     XTALCB, "
    sql = sql & "     (INPOSCB+RLENCB) as INGOTPOS, "
    sql = sql & "     HINBCB, "
    sql = sql & "     REVNUMCB, "
    sql = sql & "     FACTORYCB, "
    sql = sql & "     OPECB "
    sql = sql & "     FROM XSDCB "
    sql = sql & "     WHERE XTALCB = '" & sCryNum & "' "
    sql = sql & "     AND LIVKCB = '0' "
    sql = sql & "    ), "
    sql = sql & "    (SELECT SXLIDCW, "
    sql = sql & "     SMPKBNCW, "
    sql = sql & "     TBKBNCW, "
    sql = sql & "     REPSMPLIDCW, "
    sql = sql & "     INPOSCW, "
    sql = sql & "     HINBCW, "
    sql = sql & "     REVNUMCW, "
    sql = sql & "     FACTORYCW, "
    sql = sql & "     OPECW, "
    sql = sql & "     KTKBNCW, "
    sql = sql & "     SMCRYNUMCW, "
    sql = sql & "     WFSMPLIDRSCW, "
    sql = sql & "     WFSMPLIDRS1CW, "
    sql = sql & "     WFSMPLIDRS2CW, "
    sql = sql & "     WFINDRSCW, "
    sql = sql & "     WFRESRS1CW, "
    sql = sql & "     WFRESRS2CW, "
    sql = sql & "     WFSMPLIDOICW, "
    sql = sql & "     WFINDOICW, "
    sql = sql & "     WFRESOICW, "
    sql = sql & "     WFSMPLIDB1CW, "
    sql = sql & "     WFINDB1CW, "
    sql = sql & "     WFRESB1CW, "
    sql = sql & "     WFSMPLIDB2CW, "
    sql = sql & "     WFINDB2CW, "
    sql = sql & "     WFRESB2CW, "
    sql = sql & "     WFSMPLIDB3CW, "
    sql = sql & "     WFINDB3CW, "
    sql = sql & "     WFRESB3CW, "
    sql = sql & "     WFSMPLIDL1CW, "
    sql = sql & "     WFINDL1CW, "
    sql = sql & "     WFRESL1CW, "
    sql = sql & "     WFSMPLIDL2CW, "
    sql = sql & "     WFINDL2CW, "
    sql = sql & "     WFRESL2CW, "
    sql = sql & "     WFSMPLIDL3CW, "
    sql = sql & "     WFINDL3CW, "
    sql = sql & "     WFRESL3CW, "
    sql = sql & "     WFSMPLIDL4CW, "
    sql = sql & "     WFINDL4CW, "
    sql = sql & "     WFRESL4CW, "
    sql = sql & "     WFSMPLIDDSCW, "
    sql = sql & "     WFINDDSCW, "
    sql = sql & "     WFRESDSCW, "
    sql = sql & "     WFSMPLIDDZCW, "
    sql = sql & "     WFINDDZCW, "
    sql = sql & "     WFRESDZCW, "
    sql = sql & "     WFSMPLIDSPCW, "
    sql = sql & "     WFINDSPCW, "
    sql = sql & "     WFRESSPCW, "
    sql = sql & "     WFSMPLIDDO1CW, "
    sql = sql & "     WFINDDO1CW, "
    sql = sql & "     WFRESDO1CW, "
    sql = sql & "     WFSMPLIDDO2CW, "
    sql = sql & "     WFINDDO2CW, "
    sql = sql & "     WFRESDO2CW, "
    sql = sql & "     WFSMPLIDDO3CW, "
    sql = sql & "     WFINDDO3CW, "
    sql = sql & "     WFRESDO3CW, "
    sql = sql & "     WFSMPLIDOT1CW, "
    sql = sql & "     WFINDOT1CW, "
    sql = sql & "     WFRESOT1CW, "
    sql = sql & "     WFSMPLIDOT2CW, "
    sql = sql & "     WFINDOT2CW, "
    sql = sql & "     WFRESOT2CW, "
    sql = sql & "     WFSMPLIDAOICW, "
    sql = sql & "     WFINDAOICW, "
    sql = sql & "     WFRESAOICW, "
    sql = sql & "     SMPLNUMCW, "
    sql = sql & "     SMPLPATCW, "
    sql = sql & "     LIVKCW, "
    sql = sql & "     WFSMPLIDGDCW, "
    sql = sql & "     WFINDGDCW, "
    sql = sql & "     WFRESGDCW, "
    sql = sql & "     WFHSGDCW, "
    sql = sql & "     EPSMPLIDB1CW, "
    sql = sql & "     EPINDB1CW, "
    sql = sql & "     EPRESB1CW, "
    sql = sql & "     EPSMPLIDB2CW, "
    sql = sql & "     EPINDB2CW, "
    sql = sql & "     EPRESB2CW, "
    sql = sql & "     EPSMPLIDB3CW, "
    sql = sql & "     EPINDB3CW, "
    sql = sql & "     EPRESB3CW, "
    sql = sql & "     EPSMPLIDL1CW, "
    sql = sql & "     EPINDL1CW, "
    sql = sql & "     EPRESL1CW, "
    sql = sql & "     EPSMPLIDL2CW, "
    sql = sql & "     EPINDL2CW, "
    sql = sql & "     EPRESL2CW, "
    sql = sql & "     EPSMPLIDL3CW, "
    sql = sql & "     EPINDL3CW, "
    sql = sql & "     EPRESL3CW "
    sql = sql & "     FROM XSDCW "
    sql = sql & "     WHERE XTALCW = '" & sCryNum & "' "
    sql = sql & "     AND TBKBNCW = 'B' "
    sql = sql & "    ) "
'    sql = sql & "WHERE INGOTPOS = INPOSCW(+) "
    sql = sql & "WHERE SXLIDCB = SXLIDCW(+) "           '08/07/10 ooba
    sql = sql & "AND NVL(LIVKCW,'0') = '0' "
    sql = sql & "ORDER BY INGOTPOS, TBKBN "
    
    '�ް��𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Or rs.RecordCount Mod 2 <> 0 Then
        ReDim records(0)
        DBDRV_GetXSDCW = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    '���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .SXLIDCW = rs("SXLIDCB")                'SXLID
            .SMPKBNCW = rs("SMPKBNCW")              '�T���v���敪
            .TBKBNCW = rs("TBKBN")                  'T/B�敪
            .REPSMPLIDCW = rs("REPSMPLIDCW")        '��\�T���v��ID
            .XTALCW = rs("XTALCB")                  '�����ԍ�
            .INPOSCW = rs("INGOTPOS")               '�������ʒu
            'XSDCW�̕i�Ԃ�D��
            If i Mod 2 = 0 And .SXLIDCW = records(i - 1).SXLIDCW And _
                Trim(rs("HINBCW")) <> "" And rs("HINBCW") <> records(i - 1).HINBCW Then
                
                records(i - 1).HINBCW = rs("HINBCW")        '�i��
                records(i - 1).REVNUMCW = rs("REVNUMCW")    '���i�ԍ������ԍ�
                records(i - 1).FACTORYCW = rs("FACTORYCW")  '�H��
                records(i - 1).OPECW = rs("OPECW")          '���Ə���
                .HINBCW = rs("HINBCW")              '�i��
                .REVNUMCW = rs("REVNUMCW")          '���i�ԍ������ԍ�
                .FACTORYCW = rs("FACTORYCW")        '�H��
                .OPECW = rs("OPECW")                '���Ə���
            Else
                .HINBCW = rs("HINBCB")              '�i��
                .REVNUMCW = rs("REVNUMCB")          '���i�ԍ������ԍ�
                .FACTORYCW = rs("FACTORYCB")        '�H��
                .OPECW = rs("OPECB")                '���Ə���
            End If
            .KTKBNCW = rs("KTKBNCW")                '�m��敪
            .SMCRYNUMCW = rs("SMCRYNUMCW")          '�T���v���u���b�NID
            .WFSMPLIDRSCW = rs("WFSMPLIDRSCW")      '�T���v��ID(Rs)
            .WFSMPLIDRS1CW = rs("WFSMPLIDRS1CW")    '����T���v��ID1(Rs)
            .WFSMPLIDRS2CW = rs("WFSMPLIDRS2CW")    '����T���v��ID2(Rs)
            .WFINDRSCW = rs("WFINDRSCW")            '���FLG�iRs)
            .WFRESRS1CW = rs("WFRESRS1CW")          '����FLG1�iRs)
            .WFRESRS2CW = rs("WFRESRS2CW")          '����FLG2�iRs)
            .WFSMPLIDOICW = rs("WFSMPLIDOICW")      '�T���v��ID�iOi)
            .WFINDOICW = rs("WFINDOICW")            '���FLG�iOi)
            .WFRESOICW = rs("WFRESOICW")            '����FLG�iOi)
            .WFSMPLIDB1CW = rs("WFSMPLIDB1CW")      '�T���v��ID�iB1)
            .WFINDB1CW = rs("WFINDB1CW")            '���FLG�iB1)
            .WFRESB1CW = rs("WFRESB1CW")            '����FLG�iB1)
            .WFSMPLIDB2CW = rs("WFSMPLIDB2CW")      '�T���v��ID�iB2�j
            .WFINDB2CW = rs("WFINDB2CW")            '���FLG�iB2�j
            .WFRESB2CW = rs("WFRESB2CW")            '����FLG�iB2�j
            .WFSMPLIDB3CW = rs("WFSMPLIDB3CW")      '�T���v��ID�iB3)
            .WFINDB3CW = rs("WFINDB3CW")            '���FLG�iB3)
            .WFRESB3CW = rs("WFRESB3CW")            '����FLG�iB3)
            .WFSMPLIDL1CW = rs("WFSMPLIDL1CW")      '�T���v��ID�iL1)
            .WFINDL1CW = rs("WFINDL1CW")            '���FLG�iL1)
            .WFRESL1CW = rs("WFRESL1CW")            '����FLG�iL1)
            .WFSMPLIDL2CW = rs("WFSMPLIDL2CW")      '�T���v��ID�iL2)
            .WFINDL2CW = rs("WFINDL2CW")            '���FLG�iL2)
            .WFRESL2CW = rs("WFRESL2CW")            '����FLG�iL2)
            .WFSMPLIDL3CW = rs("WFSMPLIDL3CW")      '�T���v��ID�iL3)
            .WFINDL3CW = rs("WFINDL3CW")            '���FLG�iL3)
            .WFRESL3CW = rs("WFRESL3CW")            '����FLG�iL3)
            .WFSMPLIDL4CW = rs("WFSMPLIDL4CW")      '�T���v��ID�iL4)
            .WFINDL4CW = rs("WFINDL4CW")            '���FLG�iL4)
            .WFRESL4CW = rs("WFRESL4CW")            '����FLG�iL4)
            .WFSMPLIDDSCW = rs("WFSMPLIDDSCW")      '�T���v��ID�iDS)
            .WFINDDSCW = rs("WFINDDSCW")            '���FLG�iDS)
            .WFRESDSCW = rs("WFRESDSCW")            '����FLG�iDS)
            .WFSMPLIDDZCW = rs("WFSMPLIDDZCW")      '�T���v��ID�iDZ)
            .WFINDDZCW = rs("WFINDDZCW")            '���FLG�iDZ)
            .WFRESDZCW = rs("WFRESDZCW")            '����FLG�iDZ)
            .WFSMPLIDSPCW = rs("WFSMPLIDSPCW")      '�T���v��ID�iSP)
            .WFINDSPCW = rs("WFINDSPCW")            '���FLG�iSP)
            .WFRESSPCW = rs("WFRESSPCW")            '����FLG�iSP)
            .WFSMPLIDDO1CW = rs("WFSMPLIDDO1CW")    '�T���v��ID�iDO1)
            .WFINDDO1CW = rs("WFINDDO1CW")          '���FLG�iDO1)
            .WFRESDO1CW = rs("WFRESDO1CW")          '����FLG�iDO1)
            .WFSMPLIDDO2CW = rs("WFSMPLIDDO2CW")    '�T���v��ID�iDO2)
            .WFINDDO2CW = rs("WFINDDO2CW")          '���FLG�iDO2)
            .WFRESDO2CW = rs("WFRESDO2CW")          '����FLG�iDO2)
            .WFSMPLIDDO3CW = rs("WFSMPLIDDO3CW")    '�T���v��ID�iDO3)
            .WFINDDO3CW = rs("WFINDDO3CW")          '���FLG�iDO3)
            .WFRESDO3CW = rs("WFRESDO3CW")          '����FLG�iDO3)
            .WFSMPLIDOT1CW = rs("WFSMPLIDOT1CW")    '�T���v��ID�iOT1)
            .WFINDOT1CW = rs("WFINDOT1CW")          '���FLG�iOT1)
            .WFRESOT1CW = rs("WFRESOT1CW")          '����FLG�iOT1)
            .WFSMPLIDOT2CW = rs("WFSMPLIDOT2CW")    '�T���v��ID�iOT2)
            .WFINDOT2CW = rs("WFINDOT2CW")          '���FLG�iOT2)
            .WFRESOT2CW = rs("WFRESOT2CW")          '����FLG�iOT2)
            .WFSMPLIDAOICW = rs("WFSMPLIDAOICW")    '�T���v��ID�iAOi)
            .WFINDAOICW = rs("WFINDAOICW")          '���FLG�iAOi)
            .WFRESAOICW = rs("WFRESAOICW")          '����FLG�iAOi)
            .SMPLNUMCW = rs("SMPLNUMCW")            '�T���v������
            .SMPLPATCW = rs("SMPLPATCW")            '�T���v���p�^�[��
            .LIVKCW = rs("LIVKCW")                  '�����敪
            .WFSMPLIDGDCW = rs("WFSMPLIDGDCW")      '�T���v��ID�iGD)
            .WFINDGDCW = rs("WFINDGDCW")            '���FLG�iGD)
            .WFRESGDCW = rs("WFRESGDCW")            '����FLG�iGD)
            .WFHSGDCW = rs("WFHSGDCW")              '�ۏ�FLG�iGD)
            .EPSMPLIDB1CW = rs("EPSMPLIDB1CW")      '�T���v��ID�iB1E)
            .EPINDB1CW = rs("EPINDB1CW")            '���FLG�iB1E)
            .EPRESB1CW = rs("EPRESB1CW")            '����FLG�iB1E)
            .EPSMPLIDB2CW = rs("EPSMPLIDB2CW")      '�T���v��ID�iB2E�j
            .EPINDB2CW = rs("EPINDB2CW")            '���FLG�iB2E�j
            .EPRESB2CW = rs("EPRESB2CW")            '����FLG�iB2E�j
            .EPSMPLIDB3CW = rs("EPSMPLIDB3CW")      '�T���v��ID�iBE3)
            .EPINDB3CW = rs("EPINDB3CW")            '���FLG�iB3E)
            .EPRESB3CW = rs("EPRESB3CW")            '����FLG�iB3E)
            .EPSMPLIDL1CW = rs("EPSMPLIDL1CW")      '�T���v��ID�iL1E)
            .EPINDL1CW = rs("EPINDL1CW")            '���FLG�iL1E)
            .EPRESL1CW = rs("EPRESL1CW")            '����FLG�iL1E)
            .EPSMPLIDL2CW = rs("EPSMPLIDL2CW")      '�T���v��ID�iL2E)
            .EPINDL2CW = rs("EPINDL2CW")            '���FLG�iL2E)
            .EPRESL2CW = rs("EPRESL2CW")            '����FLG�iL2E)
            .EPSMPLIDL3CW = rs("EPSMPLIDL3CW")      '�T���v��ID�iL3E)
            .EPINDL3CW = rs("EPINDL3CW")            '���FLG�iL3E)
            .EPRESL3CW = rs("EPRESL3CW")            '����FLG�iL3E)
            
            '�������ID�o�^
            If Trim(.REPSMPLIDCW) = "" Then
                .REPSMPLIDCW = Mid(.SXLIDCW, 1, 10) & Format(CStr(.INPOSCW), "000") & .SMPKBNCW
            End If
        End With
        rs.MoveNext
    Next i
    rs.Close
    
    DBDRV_GetXSDCW = FUNCTION_RETURN_SUCCESS
    
End Function

'�T�v      :�����҂��ꗗ �\���p�c�a�h���C�o
'�p�����[�^�@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:pLackBlk�@�@�@,O  ,typ_LackBlk    �@,�����҂��u���b�N�ꗗ
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
'����      :
'����      :2001/07/10 ���{ �쐬
'          :2003/05/27 �}�@�����҂��ꗗ�擾�p�h���C�o�ɕύX
Public Function DBDRV_scmzc_fcmkc001j_Disp(pLackBlk() As typ_LackBlk) As FUNCTION_RETURN

'    Dim rs As OraDynaset
'    Dim sql As String
'    Dim recCnt As Integer
'    Dim i As Long
    Dim rs, rs2, rs3                As OraDynaset
    Dim sql, sql2, sql3             As String
    Dim recCnt, rec2Cnt, rec3Cnt    As Integer
    Dim i, j, k, cnt                As Long
    Dim SXLID, blkID                As String

''    ReDim wBLKID(0) As String
    Dim ChkCnt                      As Long
    Dim BLKIDFlg                    As Boolean
    Dim OldBlk                      As String
    Dim iCnt                        As Integer

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001j_SQL.bas -- Function DBDRV_scmzc_fcmkc001j_Disp"

    ''�@�����҂��ꗗ�̎擾���@�ύX�@2003/08/27 Mori =================> START
    ''
    ''�@�����҂��ꗗ�̎擾���@�ύX�@2003/05/27 tuku =================> START
    '  SXL�Ǘ��iXSDCB�j�̌��ݍH����CW740�̃��b�g�����ׂĕ\��������B
    sql = vbNullString
    ' SXL�Ǘ�<XSDCB>�̌��ݍH����CW740�̃��b�g���擾
    '                   �A���A�T���v���̊m��敪��'9'�̕��͏���
    '                   ���̏ꍇ�ł�XSDCB�̃z�[���h�敪<HOLDBCB>���D�悳���i'9�Ȃ�Ε\���j
'    sql = sql & " select ca.lot,e040.ingotpos,e040.REALLEN,ca.day"
'    sql = sql & " from      "
'    sql = sql & " (select distinct(crynumca) lot,max(inposca) pos,max(kdayca) day "
'    sql = sql & "  from xsdca       "
'    sql = sql & "  where sxlidca in "
'    sql = sql & "    (select sxlidcb"
'    sql = sql & "     from xsdcb cb,"
'    sql = sql & "     vecme011 ve011"
'    sql = sql & "     where gnwkntcb = 'CW740'"
'    sql = sql & "     and cb.sxlidcb = ve011.e042sxlid"
'    sql = sql & "     and (   ve011.E044KTKBN != '9' "
'    sql = sql & "          or cb.holdbcb = '9')"
'    sql = sql & "    )  "
'    sql = sql & " and livkca = '0'      "
'    sql = sql & " group by crynumca) ca ,       "
'    sql = sql & " tbcme040 e040"
'    sql = sql & " where e040.blockid = ca.lot"
'    sql = sql & " order by ca.lot           "


'''''    sql = sql & " select ca.lot,e040.ingotpos,e040.REALLEN,ca.day"
'''''
'''''    sql = sql & ",c1.puptnc1"   '��������ݒǉ��Ή�(2004/12/08) kubota
'''''
'''''    sql = sql & " from      "
'''''    sql = sql & " (select distinct(crynumca) lot,max(inposca) pos,max(kdayca) day "
'''''    sql = sql & "  from  xsdca      "
'''''    sql = sql & "  where sxlidca in "
'''''    sql = sql & "    (select sxlidcb"
'''''    sql = sql & "     from xsdcb cb, xsdcw cw "
'''''    sql = sql & "     where gnwkntcb  = 'CW740'"
'''''    sql = sql & "       and cb.sxlidcb=  cw.sxlidcw"
'''''    sql = sql & "       and (cw.ktkbncw!='9' "
'''''    sql = sql & "         or cb.holdbcb ='9')"
'''''    sql = sql & "       and cw.LIVKCW = '0'"
'''''    sql = sql & "       and cb.LIVKCB = '0'"
'''''    sql = sql & "    )"
'''''    sql = sql & " and livkca = '0'"
'''''    sql = sql & " group by crynumca) ca ,"
'''''    sql = sql & " tbcme040 e040"
'''''
'''''    sql = sql & ",xsdc1 c1,xsdc2 c2"    '��������ݒǉ��Ή�(2004/12/08) kubota
'''''
'''''    sql = sql & " where e040.blockid = ca.lot"
'''''
'''''    '��������ݒǉ��Ή�(2004/12/08) kubota
'''''    sql = sql & "   and c2.crynumc2  = ca.lot"
'''''    sql = sql & "   and c2.xtalc2    = c1.xtalc1"
'''''
'''''    sql = sql & " order by ca.lot"


''    ' �֘A����SXLID����ۯ�ID�����ׂĎ擾����SQL���ɕύX 2005/03/18 ffc)tanabe
''    sql = sql & " select ca.lot,e040.ingotpos,e040.REALLEN,ca.day,c1.puptnc1 from "
''    sql = sql & "   (select distinct(crynumca) lot,max(inposca) pos,max(kdayca) day from xsdca"
''    sql = sql & "     where crynumca in"
''    sql = sql & "       (select blockid from tbcmy001 where  sblockid in"
''    sql = sql & "         (select distinct sblockid from tbcmy001 where blockid in"
''    sql = sql & "           (select crynumca from xsdca where sxlidca in"
''    sql = sql & "             (select distinct sxlidcb from xsdcb cb, xsdcw cw"
''    sql = sql & "                where gnwkntcb  = 'CW740'"
''    sql = sql & "                and cb.sxlidcb= cw.sxlidcw"
'''    sql = sql & "                and (cw.ktkbncw!='9' or cb.holdbcb ='9')"  '�����ĉ��@06/01/12 ooba
''    sql = sql & "                and cw.LIVKCW = '0' and cb.LIVKCB = '0'"
''    sql = sql & "             )"
''    sql = sql & "             and livkca = '0' group by crynumca"
''    sql = sql & "           )"
''    sql = sql & "         )"
''    sql = sql & "       )"
''    sql = sql & "     and livkca = '0' group by crynumca"
''    sql = sql & "   ) ca ,"
''    sql = sql & " tbcme040 e040,xsdc1 c1,xsdc2 c2 where e040.blockid = ca.lot"
''    sql = sql & " and c2.crynumc2  = ca.lot"
''    sql = sql & " and c2.xtalc2    = c1.xtalc1"
''    sql = sql & " order by ca.lot"

''    '�҂��ꗗ�擾SQL�ύX�@06/02/06 ooba START ===============================================>
''    sql = sql & "SELECT DISTINCT "
''    sql = sql & "CA_GP.CRYNUMCA, "
''    sql = sql & "E40.INGOTPOS, "
''    sql = sql & "E40.REALLEN, "
''    sql = sql & "C1.PUPTNC1, "
''    sql = sql & "CA_GP.MHOLDBCA, "
''    sql = sql & "CA_GP.MWFHOLDFLGCA, "
''    sql = sql & "C2.WFHUFLG, "
''    sql = sql & "CA_GP.MKDAYCA "
''    sql = sql & ", C2.PLANTCATC2 "  ' ���� 2007/09/03 SPK Tsutsumi Add
''    sql = sql & "FROM "
''    sql = sql & "   (SELECT CRYNUMCA, MAX(HOLDBCA) MHOLDBCA, "
''    sql = sql & "    MAX(WFHOLDFLGCA) MWFHOLDFLGCA, MAX(KDAYCA) MKDAYCA "
''    sql = sql & "    FROM XSDCA WHERE LIVKCA = '0' GROUP BY CRYNUMCA "
''    sql = sql & "   ) CA_GP, "
''    sql = sql & "   (SELECT CRYNUMCA, SXLIDCA FROM XSDCA WHERE LIVKCA = '0' "
''    sql = sql & "   ) CA_AL, "
''' 2007/09/03 SPK Tsutsumi Add Start
''    sql = sql & "   (SELECT CRYNUMC2, XTALC2, WFHUFLG, PLANTCATC2 FROM XSDC2 WHERE LIVKC2 = '0' "
'''    sql = sql & "   (SELECT CRYNUMC2, XTALC2, WFHUFLG FROM XSDC2 WHERE LIVKC2 = '0' "
''' 2007/09/03 SPK Tsutsumi Add End
''sql = sql & "   ) C2, "
''    sql = sql & "   (SELECT SXLIDCB FROM XSDCB WHERE LIVKCB = '0' AND GNWKNTCB = 'CW740' "
''    sql = sql & "   ) CB, "
''    sql = sql & "   (SELECT XTALC1, PUPTNC1 FROM XSDC1 "
''    sql = sql & "   ) C1, "
''    sql = sql & "   (SELECT BLOCKID, INGOTPOS, REALLEN FROM TBCME040 "
''    sql = sql & "   ) E40 "
''    sql = sql & "WHERE C1.XTALC1 = C2.XTALC2 "
''    sql = sql & "AND C2.CRYNUMC2 = CA_GP.CRYNUMCA "
''    sql = sql & "AND C2.CRYNUMC2 = E40.BLOCKID "
''    sql = sql & "AND CA_GP.CRYNUMCA = CA_AL.CRYNUMCA "
''    sql = sql & "AND CA_AL.SXLIDCA = CB.SXLIDCB "
''
''    ' ���� 2007/09/03 SPK Tsutsumi Add Start
''    If sCmbMukesaki <> "ALL" Then
''        sql = sql & "   AND C2.PLANTCATC2      = '" & sCmbMukesaki & "'"
''    End If
''    ' 2007/09/03 SPK Tsutsumi Add End
''
''    sql = sql & "ORDER BY CA_GP.CRYNUMCA"
''    '�҂��ꗗ�擾SQL�ύX�@06/02/06 ooba END =================================================>

    '�҂��ꗗ�擾SQL�ύX�@08/01/31 ooba START ===============================================>
    sql = vbNullString
    sql = sql & "SELECT "
    sql = sql & "CA_GP.CRYNUMCA, "
    'Chg Start 2010/07/08 SMPK Nakamura
'    sql = sql & "E40.INGOTPOS, "
'    sql = sql & "E40.REALLEN, "
    sql = sql & "C2.INPOSC2, "
    sql = sql & "C2.GNLC2, "
    'Chg End 2010/07/08 SMPK Nakamura
    sql = sql & "C1.PUPTNC1, "
    sql = sql & "CA_GP.MHOLDBCA, "
    sql = sql & "CA_GP.MWFHOLDFLGCA, "
    sql = sql & "C2.WFHUFLG, "
    sql = sql & "CA_GP.MKDAYCA, "
    sql = sql & "C2.PLANTCATC2, "
    sql = sql & "MAX(CB.GNWKNTCB) MGNWKNTCB, "
    sql = sql & "MAX(CB.KBLKFLGCB) MKBLKFLGCB "
    ' ������~���ڒǉ� add SETkimizuka Start  09/03/18
    ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/29
    'sql = sql & " , NVL(TO_CHAR(Y4.AGRSTATUS),' ') as AGRSTATUSY4 "
    'sql = sql & " , NVL(TO_CHAR(Y4.STOP),'0') as STOP "
    'sql = sql & " , NVL(Y4.CAUSE,' ') as CAUSEY4 "
    'sql = sql & " , NVL(Y4.PRINTKIND || Y4.PRINTNO,' ') as PRINTNOY4 "
    sql = sql & " , NVL(TO_CHAR(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)),' ') as AGRSTATUS "
    sql = sql & " , NVL(TO_CHAR(Y4.STOPY4),' ') as STOP "
    sql = sql & " , DECODE(TRIM(Y4.CAUSEY4),NULL,' ',TRIM(Y4.CAUSEY4) || ':' || A9.NAMEJA9) as CAUSE "
    sql = sql & " , NVL(Y4.PRINTKINDY4 || Y4.PRINTNOY4,' ') as PRINTNOY4 "
    sql = sql & " , NVL(Y4.WKKTY4,'0') as WKKTY4 "
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/06/29
    ' ������~���ڒǉ� add SETkimizuka End    09/03/18
    'Add Start 2010/07/08 SMPK Nakamura
    sql = sql & " , CA_GP.HIN_CNT as HIN_CNT "
    sql = sql & " , NVL(CB.CW740STSCB,' ') as CW740STS "
    sql = sql & " , MAX(CA_AL.HINBCA) as HINBAN "
    'Add End 2010/07/08 SMPK Nakamura
    sql = sql & "FROM "
    sql = sql & "   (SELECT CRYNUMCA, MAX(HOLDBCA) MHOLDBCA, "
    sql = sql & "    MAX(WFHOLDFLGCA) MWFHOLDFLGCA, MAX(KDAYCA) MKDAYCA "
    'Add Start 2010/07/08 SMPK Nakamura
    sql = sql & "    , COUNT(CRYNUMCA) HIN_CNT "
    'Add End 2010/07/08 SMPK Nakamura
    sql = sql & "    FROM XSDCA WHERE LIVKCA = '0' "
    sql = sql & "    GROUP BY CRYNUMCA "
    sql = sql & "   ) CA_GP, "
    sql = sql & "   (SELECT CRYNUMCA, SXLIDCA "
    'Add Start 2010/07/08 SMPK Nakamura
    sql = sql & "    , HINBCA "
    'Add End 2010/07/08 SMPK Nakamura
    sql = sql & "    FROM XSDCA "
    sql = sql & "    WHERE LIVKCA = '0' "
    sql = sql & "   ) CA_AL, "
    sql = sql & "   (SELECT CRYNUMC2, XTALC2, WFHUFLG, PLANTCATC2, GNWKNTC2, KBLKFLGC2 "
    'Add Start 2010/07/08 SMPK Nakamura
    sql = sql & "    , INPOSC2, GNLC2 "
    'Add End 2010/07/08 SMPK Nakamura
    sql = sql & "    FROM XSDC2 "
    'Chg Start 2010/07/08 SMPK Nakamura
'    sql = sql & "    WHERE LIVKC2 = '0' AND GNWKNTC2 = 'CW750' "
    sql = sql & "    WHERE LIVKC2 = '0' AND (GNWKNTC2 = 'CW740' or GNWKNTC2 = 'CW750') "
    'Chg End 2010/07/08 SMPK Nakamura
    sql = sql & "   ) C2, "
    sql = sql & "   (SELECT SXLIDCB, GNWKNTCB, DECODE(KBLKFLGCB,'1',KBLKFLGCB,'0') KBLKFLGCB "
    'Add Start 2010/07/08 SMPK Nakamura
    sql = sql & "    , CW740STSCB "
    'Add End 2010/07/08 SMPK Nakamura
    sql = sql & "    FROM XSDCB "
    sql = sql & "    WHERE LIVKCB = '0' AND GNWKNTCB IN ('CST02', 'CW740') "
    sql = sql & "   ) CB, "
    sql = sql & "   (SELECT XTALC1, PUPTNC1 "
    sql = sql & "    FROM XSDC1 "
    sql = sql & "   ) C1, "
    'Chg Start 2010/07/08 SMPK Nakamura
'    sql = sql & "   (SELECT BLOCKID, INGOTPOS, REALLEN "
'    sql = sql & "    FROM TBCME040 "
'    sql = sql & "   ) E40 "
'    ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/29
'    sql = sql & "    ,XODY3 Y3,XODY4 Y4,KODA9 A9 "
    sql = sql & "    XODY3 Y3,XODY4 Y4,KODA9 A9 "
    'Chg End 2010/07/08 SMPK Nakamura
    ' ������~���ڒǉ� add SETkimizuka Start  09/03/18
    'sql = sql & "    ,(SELECT XTALNOY3 as XTALNO ,MIN(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)) as AGRSTATUS  "
    'sql = sql & "      ,MAX(STOPY4) as STOP,DECODE(TRIM(CAUSEY4),'',TRIM(CAUSEY4),TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSE ,Y5.PRINTNO,Y5.PRINTKIND "
    'sql = sql & "      FROM XODY3  "
    'sql = sql & "           LEFT OUTER JOIN XODY4 on ( XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND LIVKY4 = '0' AND STOPY4 <> '2'  AND WKKTY4 in " & CreateWkktSQL(WATCH_PROCCD) & ") "
    'sql = sql & "           LEFT OUTER JOIN KODA9 on ( SYSCA9 = 'X' AND SHUCA9 = '30' AND CAUSEY4 = CODEA9 ) "
    'sql = sql & "           LEFT OUTER JOIN (SELECT XTALNOY4 as XTALNO,SXLIDY4 as SXLID,PRINTNOY5 as PRINTNO,PRINTKINDY5 as PRINTKIND "
    'sql = sql & "                FROM XODY3,XODY4,XODY5 "
    'sql = sql & "              WHERE XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND LIVKY3 = '0' "
    'sql = sql & "                AND PRINTKINDY4 = PRINTKINDY5 AND PRINTNOY4 = PRINTNOY5  "
    'sql = sql & "                AND HKBNY5 ='0' GROUP BY XTALNOY4,SXLIDY4,PRINTNOY5,PRINTKINDY5) Y5 ON (XTALNOY3 = XTALNO AND SXLIDY3 = SXLID ) "
    'sql = sql & "      WHERE  "
    'sql = sql & "       LIVKY3    = '0' "
    'sql = sql & "       GROUP BY XTALNOY3,AGRSTATUSY4,CAUSEY4,Y5.PRINTNO,Y5.PRINTKIND,NAMEJA9) Y4 "
    ' ������~���ڒǉ� add SETkimizuka End  09/03/18
    ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/29
    sql = sql & "WHERE C1.XTALC1 = C2.XTALC2 "
    sql = sql & "AND C2.CRYNUMC2 = CA_GP.CRYNUMCA "
'    sql = sql & "AND C2.CRYNUMC2 = E40.BLOCKID "       2010/07/08 SMPK Nakamura
    sql = sql & "AND CA_GP.CRYNUMCA = CA_AL.CRYNUMCA "
    sql = sql & "AND CA_AL.SXLIDCA = CB.SXLIDCB "
    sql = sql & "AND ((CB.GNWKNTCB = 'CW740') OR "
    'Chg Start 2010/07/08 SMPK Nakamura
'    sql = sql & "     (CB.GNWKNTCB = 'CST02' AND C2.GNWKNTC2 = 'CW750' AND C2.KBLKFLGC2 = '1')) "
    sql = sql & "     (CB.GNWKNTCB = 'CST02' AND (C2.GNWKNTC2 = 'CW740' or C2.GNWKNTC2 = 'CW750') AND C2.KBLKFLGC2 = '1')) "
    'Chg End 2010/07/08 SMPK Nakamura
    If sCmbMukesaki <> "ALL" Then
        sql = sql & "   AND C2.PLANTCATC2      = '" & sCmbMukesaki & "'"
    End If
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/06/26
    'sql = sql & "   AND CA_GP.CRYNUMCA     = Y4.XTALNO(+) "            'add 09/03/18 SETkimizuka
    sql = sql & " AND CA_GP.CRYNUMCA = Y3.XTALNOY3(+) "
    sql = sql & " AND Y3.LIVKY3(+) = '0' "
    sql = sql & " AND Y4.LIVKY4(+) = '0' "
    sql = sql & " AND Y3.XTALNOY3 = Y4.XTALNOY4(+) "
    sql = sql & " AND Y3.RCNTY3 = Y4.RCNTY4(+) "
    sql = sql & " AND A9.SYSCA9(+) = 'X' AND A9.SHUCA9(+) = '30' AND Y4.CAUSEY4 = A9.CODEA9(+) "
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/06/26
    sql = sql & " GROUP BY "
    sql = sql & "CA_GP.CRYNUMCA, "
    'Chg Start 2010/07/08 SMPK Nakamura
'    sql = sql & "E40.INGOTPOS, "
'    sql = sql & "E40.REALLEN, "
    sql = sql & "C2.INPOSC2, "
    sql = sql & "C2.GNLC2, "
    'Chg End 2010/07/08 SMPK Nakamura
    sql = sql & "C1.PUPTNC1, "
    sql = sql & "CA_GP.MHOLDBCA, "
    sql = sql & "CA_GP.MWFHOLDFLGCA, "
    sql = sql & "C2.WFHUFLG, "
    sql = sql & "CA_GP.MKDAYCA, "
    sql = sql & "C2.PLANTCATC2 "
    ' ������~���ڒǉ� add SETkimizuka Start  09/03/18
    sql = sql & ",Y4.AGRSTATUSY4 "
    sql = sql & ",Y4.STOPY4 "
    sql = sql & ",Y4.CAUSEY4 "
    sql = sql & ",Y4.PRINTKINDY4 "
    sql = sql & ",Y4.PRINTNOY4 "
    sql = sql & ",Y4.WKKTY4 "   'add 09/06/29 SETkimizuka
    sql = sql & ",A9.NAMEJA9 "  'add 09/06/29 SETkimizuka
    'Add Start 2010/07/08 SMPK Nakamura
    sql = sql & ",CA_GP.HIN_CNT "
    sql = sql & ",CB.CW740STSCB "
    'Add End 2010/07/08 SMPK Nakamura
    ' ������~���ڒǉ� add SETkimizuka End    09/03/18
    sql = sql & "ORDER BY CA_GP.CRYNUMCA"
    '�҂��ꗗ�擾SQL�ύX�@08/01/31 ooba END =================================================>
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount

    ' ������~���ڒǉ��ɔ����擾���@�ύX add SETkimizuka Start  09/03/18
    iCnt = 0
    '�������ʂ��i�[����i�u���b�NID,�������J�n�ʒu,�u���b�N����,���ɓ���)
    For i = 1 To recCnt
        If OldBlk <> rs("CRYNUMCA") Then
            iCnt = iCnt + 1
            ReDim Preserve pLackBlk(iCnt)
            With pLackBlk(iCnt)
                'Chg Start 2010/07/08 SMPK Nakamura
'                .INGOTPOS = rs("INGOTPOS")  ' �������J�n�ʒu
'                .REALLEN = rs("REALLEN")    ' ������
                .INGOTPOS = rs("INPOSC2")  ' �������J�n�ʒu
                .REALLEN = rs("GNLC2")    ' ������
                'Chg End 2010/07/08 SMPK Nakamura
                .PUPTN = rs("PUPTNC1")
        
                .BLOCKID = rs("CRYNUMCA")       '��ۯ�ID
                'ΰ��ދ敪
                If IsNull(rs("MHOLDBCA")) Then .HOLDFLG = " " Else .HOLDFLG = rs("MHOLDBCA")
                'WFΰ��ދ敪
                If IsNull(rs("MWFHOLDFLGCA")) Then .WFHOLDFLG = " " Else .WFHOLDFLG = rs("MWFHOLDFLGCA")
                'WF�U��FLG
                If IsNull(rs("WFHUFLG")) Then .WFHUFLG = " " Else .WFHUFLG = rs("WFHUFLG")
                .REJDTTM = rs("MKDAYCA")        '�X�V���t
        
                If IsNull(rs("PLANTCATC2")) = False Then
                    For j = 0 To UBound(s_MukesakiBase)
                        If s_MukesakiBase(j).sMukeCode = rs("PLANTCATC2") Then
                           .MUKESAKI = s_MukesakiBase(j).sMukeName
                        End If
                    Next j
                End If
                .Koutei = rs("MGNWKNTCB")       '�H��(XSDCB)�@08/01/31 ooba
                .KANREN = rs("MKBLKFLGCB")      '�֘A��ۯ��L���@08/01/31 ooba
                
                ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/26
                '.STOP = rs("STOP")                   '��~�敪
                '.AGRSTATUS = rs("AGRSTATUSY4")       '���F�m�F�敪
                'If Trim(rs("CAUSEY4")) <> "" Then
                '    .CAUSE = rs("CAUSEY4") & vbTab       '��~���R
                'End If
                If rs("STOP") <> "2" And (rs("WKKTY4") = "CW740" Or rs("WKKTY4") = "CW000") Then
                    .STOP = rs("STOP")                   '��~�敪
                    .AGRSTATUS = rs("AGRSTATUS")       '���F�m�F�敪
                    If Trim(rs("CAUSE")) <> "" Then
                        .CAUSE = rs("CAUSE") & vbTab       '��~���R
                    End If
                End If
                ' �����Ď�SQL�C�� upd SETkimizuka End  09/06/26
                If Trim(rs("PRINTNOY4")) <> "" Then
                    .PRINTNO = rs("PRINTNOY4") & vbTab       '��s�]��
                End If
                'Add Start 2010/07/08 SMPK Nakamura
                '�i�Ԑ�
                .HINCNT = rs("HIN_CNT")
                '�i��
                .hinban = rs("HINBAN")
                'CW740�X�e�[�^�X
                .CW740STS = rs("CW740STS")
                'Add End 2010/07/08 SMPK Nakamura
            End With
        Else
            With pLackBlk(iCnt)
                ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/26
                'If Trim(rs("CAUSEY4")) <> "" And InStr(.CAUSE, rs("CAUSEY4")) = 0 Then
                '    .CAUSE = .CAUSE & rs("CAUSEY4") & vbTab        '��~�敪
                'End If
                If rs("STOP") <> "2" And (rs("WKKTY4") = "CW740" Or rs("WKKTY4") = "CW000") Then
                    If Trim(.AGRSTATUS) = "" Or (rs("AGRSTATUS") < .AGRSTATUS And Trim(rs("AGRSTATUS")) <> "") Then
                         .AGRSTATUS = rs("AGRSTATUS")
                         .STOP = rs("STOP")
                    End If
                    If Trim(rs("CAUSE")) <> "" And InStr(.CAUSE, rs("CAUSE")) = 0 Then
                        .CAUSE = .CAUSE & rs("CAUSE") & vbTab        '��~�敪
                    End If
                End If
                If Trim(rs("PRINTNOY4")) <> "" And InStr(.PRINTNO, rs("PRINTNOY4")) = 0 Then
                    .PRINTNO = .PRINTNO & rs("PRINTNOY4") & vbTab        '��s�]��
                End If
            End With
            
        End If
        
        OldBlk = rs("CRYNUMCA")
        rs.MoveNext
    Next i
    ' ������~���ڒǉ��ɔ����擾���@�ύX add SETkimizuka End  09/03/18

    ' ������~���ڒǉ��ɔ����擾���@�ύX del SETkimizuka Start  09/03/18
'    For i = 1 To recCnt
'        iCnt = iCnt + 1
'        ReDim Preserve pLackBlk(i)
'        With pLackBlk(i)
''        .BLOCKID = rs("LOT")        ' �u���b�NID
'        .IngotPos = rs("INGOTPOS")  ' �������J�n�ʒu
'        .REALLEN = rs("REALLEN")    ' ������
''        .REJDTTM = rs("DAY")        ' ���ɓ�
'        .PUPTN = rs("PUPTNC1")
'
'        '06/02/06 ooba START ===============================================================>
'        .BLOCKID = rs("CRYNUMCA")       '��ۯ�ID
'        'ΰ��ދ敪
'        If IsNull(rs("MHOLDBCA")) Then .HOLDFLG = " " Else .HOLDFLG = rs("MHOLDBCA")
'        'WFΰ��ދ敪
'        If IsNull(rs("MWFHOLDFLGCA")) Then .WFHOLDFLG = " " Else .WFHOLDFLG = rs("MWFHOLDFLGCA")
'        'WF�U��FLG
'        If IsNull(rs("WFHUFLG")) Then .WFHUFLG = " " Else .WFHUFLG = rs("WFHUFLG")
'        .REJDTTM = rs("MKDAYCA")        '�X�V���t
'        '06/02/06 ooba END =================================================================>
'
'        ' 2007/09/03 SPK Tsutsumi Add Start
'        If IsNull(rs("PLANTCATC2")) = False Then
'            For j = 0 To UBound(s_MukesakiBase)
'                If s_MukesakiBase(j).sMukeCode = rs("PLANTCATC2") Then
'                   .MUKESAKI = s_MukesakiBase(j).sMukeName
'                End If
'            Next j
'        End If
'        ' 2007/09/03 SPK Tsutsumi Add End
'        .Koutei = rs("MGNWKNTCB")       '�H��(XSDCB)�@08/01/31 ooba
'        .KANREN = rs("MKBLKFLGCB")      '�֘A��ۯ��L���@08/01/31 ooba
'
'        End With
'
'        OldBlk = rs("CRYNUMCA")  'add 09/03/18 SETkimizuka
'        rs.MoveNext
'    Next i
    ' ������~���ڒǉ��ɔ����擾���@�ύX del SETkimizuka End  09/03/18
    rs.Close


    ''�@�����҂��ꗗ�̎擾���@�ύX�@2003/05/27 tuku =================> END
    ''�@�����҂��ꗗ�̎擾���@�ύX�@2003/08/27 Mori =================> END

''    '' ΰ��ދ敪�擾�ǉ��@05/01/31 ooba START ===================================>
''    If recCnt > 0 Then      '05/04/19 ooba
''        For i = 1 To UBound(pLackBlk)
''            sql2 = "SELECT "
''            sql2 = sql2 & "HOLDBCA, "
''            sql2 = sql2 & "WFHOLDFLGCA "
''            sql2 = sql2 & "FROM XSDCA "
''            sql2 = sql2 & "WHERE "
''            sql2 = sql2 & "CRYNUMCA = '" & pLackBlk(i).BLOCKID & "' "
''            sql2 = sql2 & "AND LIVKCA = '0' "
''
''            Set rs2 = OraDB.DBCreateDynaset(sql2, ORADYN_NO_BLANKSTRIP)
''            rec2Cnt = rs2.RecordCount
''            If rec2Cnt > 0 Then
''                For j = 1 To rec2Cnt
''                    If IsNull(rs2("HOLDBCA")) = False Then
''                        pLackBlk(i).HOLDFLG = rs2("HOLDBCA")        'ΰ��ދ敪
''                    Else
''                        pLackBlk(i).HOLDFLG = ""
''                    End If
''                    If IsNull(rs2("WFHOLDFLGCA")) = False Then
''                        pLackBlk(i).WFHOLDFLG = rs2("WFHOLDFLGCA")  'WFΰ��ދ敪
''                    Else
''                        pLackBlk(i).WFHOLDFLG = ""
''                    End If
''
''                    If pLackBlk(i).HOLDFLG <> "0" Or pLackBlk(i).HOLDFLG <> " " Then
''                        Exit For
''                    End If
''                    rs2.MoveNext
''                Next
''            End If
''            rs2.Close
''        Next
''    End If
''    '' ΰ��ދ敪�擾�ǉ��@05/01/31 ooba END =====================================>

    DBDRV_scmzc_fcmkc001j_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmkc001j_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function






'��{�����p�����[�^�쐬 2002/09/10 ADD hitec)N.MATSUMOTO Start
'�����FfrmFormID=������ʂ̔���i1:WF�Z���^��������@2:�Ĕ����j
Public Function MakeParameter(ByVal StrCryNum As String) As FUNCTION_RETURN

    Dim lng                 As Long
    Dim dat                 As Variant
    Dim lRowCnt             As Long
    Dim rsMain              As OraDynaset
    Dim sql                 As String
    Dim intCnt              As Integer
    Dim errTbl              As String
    Dim sErrMsg             As String
    Dim lngBeginIngotpos    As Long
    Dim lngEndIngotpos      As Long
    Dim strIngotpos         As String
    Dim varIngotpos         As Variant
    Dim i                   As Integer  'add 2003/05/17 hitec)matsumoto

    With f_cmbc036_2.sprExamine
    '�i�Ԃ�1��ǉ��������Ƃɂ���̕ύX-------start iida 2003/09/03
''''        .GetText 3, 1, varIngotpos      'upd 2003/04/07 hitec)matsumoto ��ʃ��C�A�E�g�ύX�ɔ����C��
        .GetText 5, 1, varIngotpos
''''        lngBeginIngotpos = CInt(Trim(varIngotpos))
        lngBeginIngotpos = SIngotP  'upd 2003/05/16 hitec)matsumoto
''''        .GetText 3, .MaxRows, varIngotpos   'upd 2003/04/07 hitec)matsumoto ��ʃ��C�A�E�g�ύX�ɔ����C��
        .GetText 5, .MaxRows, varIngotpos
''''        lngEndIngotpos = CInt(Trim(varIngotpos))
        lngEndIngotpos = EIngotP    'upd 2003/05/16 hitec)matsumoto
    End With
    '�i�Ԃ�1��ǉ��������Ƃɂ���̕ύX-------end iida 2003/09/03
    '�\���̍쐬
    If cmbc036_2_CreateTable(StrCryNum, lngBeginIngotpos, lngEndIngotpos, sErrMsg) = FUNCTION_RETURN_FAILURE Then
        MakeParameter = FUNCTION_RETURN_FAILURE
        f_cmbc036_2.lblMsg.Caption = sErrMsg
        Exit Function
    End If
    MakeParameter = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

End Function
'2002/09/10 ADD hitec)N.MATSUMOTO End


'�\���̍쐬�����@2002/09/10 ADD hitec)N.MATSUMOTO Start
Public Function cmbc036_2_CreateTable(ByVal StrCryNum As String, ByVal lngBeginIngotpos As Long, ByVal lngEndIngotpos As Long, ByRef strErrMsg As String) As FUNCTION_RETURN

    Dim rs              As OraDynaset
    Dim errTbl          As String
    Dim StrBlockId()    As String
    Dim strDBName       As String
    Dim bNoData         As Boolean
    Dim intLoopCnt      As Integer
    Dim sql             As String
    Dim strCryNum9      As String

    bNoData = False

    giInpos = 9000  'add 2003/04/16 hitec)matsumoto �݌Ɍ��A�U�֏��̈ʒu��������

    '�u���b�N�Ǘ�����u���b�N�h�c���擾
    'upd start 2003/03/28 hitec)matsumoto ----------------------
''''    sql = "SELECT * from TBCME040 "
''''    sql = sql & " WHERE CRYNUM='" & strCryNum & "'"
''''    sql = sql & "   AND INGOTPOS>=" & lngBeginIngotpos & " AND (INGOTPOS + LENGTH) <=" & lngEndIngotpos
''''
''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''''    If rs.RecordCount = 0 Then
''''        rs.Close
''''        cmbc036_2_CreateTable = FUNCTION_RETURN_FAILURE
''''        GoTo proc_exit
''''    End If
    '2003/04/27 hitec)okazaki �ύX
'''''    sql = "SELECT DISTINCT(CRYNUMCA) "
'''''    sql = sql & " FROM XSDCA"
'''''    sql = sql & " WHERE CRYNUMCA = '" & strCryNum & "'"
    strCryNum9 = Left(StrCryNum, 9)
    sql = "SELECT CRYNUMCA "
    sql = sql & " FROM XSDCA"
    sql = sql & " WHERE SUBSTR(CRYNUMCA,1,9) = '" & strCryNum9 & "'"
    sql = sql & "   AND (INPOSCA>=" & lngBeginIngotpos
    sql = sql & "   AND  INPOSCA< " & lngEndIngotpos & ")"
    sql = sql & "   AND LIVKCA = '0' "
    sql = sql & "GROUP BY CRYNUMCA"
''''    strCryNum9 = Left(strCryNum, 9)
''''
''''    sql = "SELECT DISTINCT(CRYNUMCA) "
''''    sql = sql & " FROM XSDCA"
''''    sql = sql & " WHERE SUBSTR(CRYNUMCA,1,9) = '" & strCryNum9 & "'"
''''    sql = sql & " AND ( INPOSCA >= " & lngBeginIngotpos & ""
''''    sql = sql & " AND  INPOSCA < " & lngEndIngotpos & ")"
''''    sql = sql & " AND  LIVKCA = '0' "
    '�ύXend
    Debug.Print sql
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    '�u���b�NID���擾
    intLoopCnt = 0
    Do While Not rs.EOF
        ReDim Preserve StrBlockId(intLoopCnt) As String
        If IsNull(rs("CRYNUMCA")) = True Then
            StrBlockId(intLoopCnt) = ""
        Else
            StrBlockId(intLoopCnt) = rs("CRYNUMCA")            '�u���b�NID
        End If
        Debug.Print "cmbc036_2_CreateTable ���[�v�J�n strBlockID(" & intLoopCnt & ") =" & StrBlockId(intLoopCnt)

        '��{���\����
        With Kihon
            .StaffID = Trim(f_cmbc036_2.txtStaffID.Text)
            .NEWPROC = PROCD_WFC_SOUGOUHANTEI
            .NOWPROC = PROCD_NUKISI_HENKOU
            .DIAMETER = 0      '--------------�ۗ�
            .ALLSCRAP = "N" '�S���X�N���b�v
        End With

        '���������i�u���b�N�j����O�H�����ю擾
        strDBName = "XSDC2"
        If cmbc036_2_CreateXSDC2(StrBlockId(intLoopCnt), bNoData) = FUNCTION_RETURN_FAILURE Then
            If bNoData = True Then
                cmbc036_2_CreateTable = FUNCTION_RETURN_SUCCESS '�����͍s��Ȃ����A����ŕԂ�
                Debug.Print "cmbc036_2_CreateXSDC2(" & StrBlockId(intLoopCnt) & "," & bNoData & "):XSDC2�O�H�����і���"
                Exit Function
            Else
                cmbc036_2_CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EAPLY") & strDBName
                Debug.Print "cmbc036_2_CreateXSDC2(" & StrBlockId(intLoopCnt) & "," & bNoData & "):XSDC2�O�H�����ѓǍ��݃G���["
                Exit Function
            End If
        End If

        '���������i�i�ԁj����O�H�����ю擾
        strDBName = "XSDCA"
        If cmbc036_2_CreateXSDCA(StrBlockId(intLoopCnt), bNoData) = FUNCTION_RETURN_FAILURE Then
            If bNoData = True Then
                cmbc036_2_CreateTable = FUNCTION_RETURN_SUCCESS '�����͍s��Ȃ����A����ŕԂ�
                Debug.Print "XSDCA�F�O�H�����і���"
                Exit Function
            Else
                cmbc036_2_CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EAPLY") & strDBName
                Debug.Print "XSDCA�F�O�H�����ѓǍ��݃G���["
                Exit Function
            End If
        End If

        '���ݍH�����э쐬
        strErrMsg = GetMsgStr("EAPLY")
        If cmbc036_2_CreateNowProc(StrBlockId(intLoopCnt), lngBeginIngotpos, lngEndIngotpos, strErrMsg) = FUNCTION_RETURN_FAILURE Then
            cmbc036_2_CreateTable = FUNCTION_RETURN_FAILURE
'            strErrMsg = GetMsgStr("EAPLY")
            Debug.Print "XSDC2,XSDCA�F���ݍH�����э쐬�G���["
            Exit Function
        End If
        strErrMsg = ""

        '��{����
        If KihonProc = FUNCTION_RETURN_FAILURE Then
            cmbc036_2_CreateTable = FUNCTION_RETURN_FAILURE
            strErrMsg = GetMsgStr("EAPLY")
            Debug.Print "��{�����ُ�I��"
            Exit Function
        End If
        intLoopCnt = intLoopCnt + 1
        rs.MoveNext
    Loop
    rs.Close

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

End Function
'2002/09/10 ADD hitec)N.MATSUMOTO End


'���������i�u���b�N�j�O�H�����ю擾���\���̍쐬 2002/09/10 ADD hitec)N.MATSUMOTO
Public Function cmbc036_2_CreateXSDC2(ByVal StrBlockId As String, ByRef bNoData As Boolean) As FUNCTION_RETURN

    Dim rs          As OraDynaset
    Dim sql         As String
    Dim intProcNo   As Integer

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    intProcNo = 0
    bNoData = False


    sql = "SELECT * from XSDC2 "
    sql = sql & " WHERE CRYNUMC2 ='" & StrBlockId & "'"
    sql = sql & "   AND LIVKC2= '0'"   '�����敪

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        bNoData = True
        cmbc036_2_CreateXSDC2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    rs.MoveFirst
    If rs.EOF = False Then
        With BlkOld
            If IsNull(rs.Fields("CRYNUMC2")) = False Then .CRYNUMC2 = rs.Fields("CRYNUMC2")
            If IsNull(rs.Fields("KCNTC2")) = False Then .KCNTC2 = rs.Fields("KCNTC2")       '�H���A��
            If IsNull(rs.Fields("XTALC2")) = False Then .XTALC2 = rs.Fields("XTALC2")
            If IsNull(rs.Fields("INPOSC2")) = False Then .INPOSC2 = rs.Fields("INPOSC2")
            If IsNull(rs.Fields("NEKKNTC2")) = False Then .NEKKNTC2 = rs.Fields("NEKKNTC2")
            If IsNull(rs.Fields("NEWKNTC2")) = False Then .NEWKNTC2 = rs.Fields("NEWKNTC2")
            If IsNull(rs.Fields("NEWKKBC2")) = False Then .NEWKKBC2 = rs.Fields("NEWKKBC2")
            If IsNull(rs.Fields("NEMACOC2")) = False Then .NEMACOC2 = rs.Fields("NEMACOC2")
            If IsNull(rs.Fields("GNKKNTC2")) = False Then .GNKKNTC2 = rs.Fields("GNKKNTC2")
            If IsNull(rs.Fields("GNWKNTC2")) = False Then .GNWKNTC2 = rs.Fields("GNWKNTC2")
            If IsNull(rs.Fields("GNWKKBC2")) = False Then .GNWKKBC2 = rs.Fields("GNWKKBC2")
            If IsNull(rs.Fields("GNMACOC2")) = False Then .GNMACOC2 = rs.Fields("GNMACOC2")
            If IsNull(rs.Fields("GNDAYC2")) = False Then .GNDAYC2 = rs.Fields("GNDAYC2")
            If IsNull(rs.Fields("GNLC2")) = False Then .GNLC2 = rs.Fields("GNLC2")          '���ݒ���
            If IsNull(rs.Fields("GNWC2")) = False Then .GNWC2 = rs.Fields("GNWC2")          '���ݏd��
            If IsNull(rs.Fields("GNMC2")) = False Then .GNMC2 = rs.Fields("GNMC2")          '���ݖ���
            If IsNull(rs.Fields("SUMITLC2")) = False Then .SUMITLC2 = rs.Fields("SUMITLC2")
            If IsNull(rs.Fields("SUMITWC2")) = False Then .SUMITWC2 = rs.Fields("SUMITWC2")
            If IsNull(rs.Fields("SUMITMC2")) = False Then .SUMITMC2 = rs.Fields("SUMITMC2")
            If IsNull(rs.Fields("CHGC2")) = False Then .CHGC2 = rs.Fields("CHGC2")
            If IsNull(rs.Fields("KAKOUBC2")) = False Then .KAKOUBC2 = rs.Fields("KAKOUBC2")
            If IsNull(rs.Fields("KEIDAYC2")) = False Then .KEIDAYC2 = rs.Fields("KEIDAYC2")
            If IsNull(rs.Fields("GNTKUBC2")) = False Then .GNTKUBC2 = rs.Fields("GNTKUBC2")
            If IsNull(rs.Fields("GNTNOC2")) = False Then .GNTNOC2 = rs.Fields("GNTNOC2")
            If IsNull(rs.Fields("XTWORKC2")) = False Then .XTWORKC2 = rs.Fields("XTWORKC2")
            If IsNull(rs.Fields("WFWORKC2")) = False Then .WFWORKC2 = rs.Fields("WFWORKC2")
            If IsNull(rs.Fields("LSTATBC2")) = False Then .LSTATBC2 = rs.Fields("LSTATBC2")
            If IsNull(rs.Fields("RSTATBC2")) = False Then .RSTATBC2 = rs.Fields("RSTATBC2")
            If IsNull(rs.Fields("LUFRCC2")) = False Then .LUFRCC2 = rs.Fields("LUFRCC2")
            If IsNull(rs.Fields("LUFRBC2")) = False Then .LUFRBC2 = rs.Fields("LUFRBC2")
            If IsNull(rs.Fields("LDFRCC2")) = False Then .LDFRCC2 = rs.Fields("LDFRCC2")
            If IsNull(rs.Fields("LDFRBC2")) = False Then .LDFRBC2 = rs.Fields("LDFRBC2")
            If IsNull(rs.Fields("HOLDCC2")) = False Then .HOLDCC2 = rs.Fields("HOLDCC2")
            If IsNull(rs.Fields("HOLDBC2")) = False Then .HOLDBC2 = rs.Fields("HOLDBC2")
            If IsNull(rs.Fields("EXKUBC2")) = False Then .EXKUBC2 = rs.Fields("EXKUBC2")
            If IsNull(rs.Fields("HENPKC2")) = False Then .HENPKC2 = rs.Fields("HENPKC2")
            If IsNull(rs.Fields("LIVKC2")) = False Then .LIVKC2 = rs.Fields("LIVKC2")
            If IsNull(rs.Fields("KANKC2")) = False Then .KANKC2 = rs.Fields("KANKC2")
            If IsNull(rs.Fields("NFC2")) = False Then .NFC2 = rs.Fields("NFC2")
            If IsNull(rs.Fields("SAKJC2")) = False Then .SAKJC2 = rs.Fields("SAKJC2")
            If IsNull(rs.Fields("TDAYC2")) = False Then .TDAYC2 = rs.Fields("TDAYC2")
            If IsNull(rs.Fields("KDAYC2")) = False Then .KDAYC2 = rs.Fields("KDAYC2")
            If IsNull(rs.Fields("SUMITBC2")) = False Then .SUMITBC2 = rs.Fields("SUMITBC2")
            If IsNull(rs.Fields("SNDKC2")) = False Then .SNDKC2 = rs.Fields("SNDKC2")
            If IsNull(rs.Fields("SNDDAYC2")) = False Then .SNDDAYC2 = rs.Fields("SNDDAYC2")
            If IsNull(rs.Fields("PLANTCATC2")) = False Then .PLANTCATC2 = rs.Fields("PLANTCATC2") ' 2007/09/04 SPK Tsutsumi Add
        End With
    End If

    rs.Close
    cmbc036_2_CreateXSDC2 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    cmbc036_2_CreateXSDC2 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2002/09/10 ADD hitec)N.MATSUMOTO

'���ݍH���\���̍쐬 2002/09/10 ADD hitec)N.MATSUMOTO
Public Function cmbc036_2_CreateNowProc(ByVal StrBlockId As String, ByVal lngBeginIngotpos As Long, ByVal lngEndIngotpos As Long, _
                                        ByRef strErrMsg As String) As FUNCTION_RETURN

    Dim rs              As OraDynaset
    Dim sql             As String
    Dim intProcNo       As Integer
    Dim intHinOldCnt    As Integer
    Dim intLengthCnt    As Integer
    Dim intLoopCnt      As Integer
    Dim dblDiameter     As Double
    Dim intNum          As Integer
    Dim StrCryNum       As String
    Dim strLstatcls     As String
    Dim intBlkLength    As Integer  '�u���b�N�Ǘ��f�[�^�̒���
    Dim intBlkIngotPos  As Integer  '�u���b�N�Ǘ��f�[�^�̈ʒu
    Dim intSxlLength    As Integer  '�V���O���Ǘ��f�[�^�̒���
    Dim intSxlIngotPos  As Integer  '�V���O���Ǘ��f�[�^�̈ʒu
    Dim bFlg            As Boolean
    Dim sp              As Integer  '��������p
    Dim ep              As Integer  '��������p
    Dim sbp             As Integer  '��������p
    Dim ebp             As Integer  '��������p
    Dim intLength       As Integer  '����
    Dim intIngotPos     As Integer  '�ʒu
    Dim rs2             As OraDynaset   'add 2003/04/15 hitec)matsumoto
    Dim iWFcnt          As Integer
    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    intProcNo = 0

    intBlkLength = 0
    intBlkIngotPos = 0
    intSxlLength = 0
    intSxlIngotPos = 0
    StrCryNum = ""

    '�u���b�N�Ǘ����璷�����擾
    sql = "SELECT * from TBCME040 "
    sql = sql & " WHERE BLOCKID='" & StrBlockId & "'"
''''    sql = sql & "   AND INGOTPOS=0"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        cmbc036_2_CreateNowProc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    intLoopCnt = 0
    If rs.EOF = False Then
        If IsNull(rs("CRYNUM")) = False Then StrCryNum = rs("CRYNUM")               '�����ԍ�
        If IsNull(rs("LENGTH")) = False Then intBlkLength = rs("LENGTH")            '����
        If IsNull(rs("INGOTPOS")) = False Then intBlkIngotPos = rs("INGOTPOS")      '�ʒu
    End If

    rs.Close

    '�u���b�N�Ǘ��Ŏ擾�������������ƂɃV���O���Ǘ�����f�[�^���擾
    'upd start 2003/04/15 hitec)matsumoto �S���X�N���b�v��TBCMY011�Ŕ��f����悤�ɏC��---------
''''    sql = "SELECT * from TBCME042 "
''''    sql = sql & " WHERE CRYNUM='" & strCryNum & "'"
''''    '�����[�v���Ŕ���
''''    sql = sql & "   AND ((INGOTPOS >=" & lngBeginIngotpos & ") And (INGOTPOS + LENGTH <= " & lngEndIngotpos & "))"
''''    sql = sql & "   AND LSTATCLS<>'H'"
    sql = "SELECT LOTID from TBCMY011 "
    sql = sql & " WHERE LOTID='" & StrBlockId & "'"     '2003/04/03 hitec)matsumoto �S���X�N���b�v="Y"�̓u���b�N�P�ʂȂ̂ŁA�V���O���͈͂Ŏ��Ȃ�
''''    sql = sql & "   AND ((BLOCKSEQ >=" & lngWfBeginSeq & ") And (BLOCKSEQ <= " & lngWfEndSeq & "))"
    sql = sql & "   AND TO_NUMBER(WFSTA) <= 1"
    'upd end   2003/04/15 hitec)matsumoto �S���X�N���b�v��TBCMY011�Ŕ��f����悤�ɏC��---------

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then   '�Y���f�[�^0���̏ꍇ�A�S���X�N���b�v�̏���
        '�O�H�����т��A���ݍH�����тɃR�s�[
        BlkNow = BlkOld
        BlkNow.GNLC2 = "0"
        BlkNow.GNWC2 = "0"
        BlkNow.GNMC2 = "0"
        BlkNow.GNWKNTC2 = Kihon.NEWPROC
        BlkNow.NEWKNTC2 = Kihon.NOWPROC
        For intHinOldCnt = 0 To Kihon.CNTHINOLD - 1
            ReDim Preserve HinNow(intHinOldCnt) As typ_XSDCA_Update
            HinNow(intHinOldCnt) = HinOld(intHinOldCnt)
            HinNow(intHinOldCnt).GNLCA = "0"    '�S���X�N���b�v=������0
            HinNow(intHinOldCnt).GNWCA = "0"    '�d�� = 0
            HinNow(intHinOldCnt).GNMCA = "0"    '���� = 0
            HinNow(intHinOldCnt).GNWKNTCA = Kihon.NEWPROC
            HinNow(intHinOldCnt).NEWKNTCA = Kihon.NOWPROC
        Next
        Kihon.CNTHINNOW = 1
        Kihon.ALLSCRAP = "Y"

        '�O�H���̒����ƌ��ݍH���̒���������ׁA�s�ǂ����݂��邩����
'        If CInt(BlkNow.GNLC2) = CInt(BlkOld.GNLC2) Then '�s�ǂȂ�
        If CInt(BlkNow.GNMC2) = CInt(BlkOld.GNMC2) Then '�s�ǂȂ�
            '��{���\����
            With Kihon
                .FURYOUMU = "N"
            End With
        Else                                            '�s�ǂ���
''            '��{���\����
''            With Kihon
''                .FURYOUMU = "Y"
''            End With
''            '�s�Ǎ\���̂��쐬
''            With Furyou
''                .XTALC4 = BlkNow.CRYNUMC2   '�u���b�NID
''                .INPOSC4 = BlkNow.INPOSC2   '�������J�n�ʒu
''                .KCKNTC4 = BlkNow.KCNTC2    '�H���A��
''                .HINBC4 = "Z"               '�i��
''    '            .REVNUMC4                   '���i�ԍ������ԍ�
''    '            .FACTORYC4                  '�H��
''    '            .OPEC4                      '���Ə���
''                .WKKTC4 = PROCD_NUKISI_HENKOU
''                .PUCUTLC4 = CLng(BlkOld.GNLC2) - CLng(BlkNow.GNLC2) '�s�ǒ���(�O�H��-���ݍH���i�Ǖi�j)
''                '�s�Ǐd��
''                If GetDiameter(.XTALC4, dblDiameter) = FUNCTION_RETURN_FAILURE Then  '���a�����߂�
''                    dblDiameter = 0
''    ''''                GoTo proc_wxit
''                End If
''                '�擾�������a�����ɏd�ʂ����߂�
''                .PUCUTWC4 = CStr(CLng(WeightOfCylinder(dblDiameter, CDbl(.PUCUTLC4))))
''                '�s�ǖ���
''                If WfCount(.XTALC4, CLng(.PUCUTLC4), intNum) = FUNCTION_RETURN_FAILURE Then
''                    .PUCUTMC4 = 0
''    ''''                GoTo proc_wxit
''                Else
''                    .PUCUTMC4 = intNum
''                End If
''
''                .SUMITBC3 = "0"
''            End With
                rs.Close
                strErrMsg = GetMsgStr("EWFM5", "�O�H��=" & BlkOld.GNMC2 & "�F���ݍH��=" & BlkNow.GNMC2) '03/06/06 �㓡
'                lblMsg.Caption = "WF�����s��v�G���["
                cmbc036_2_CreateNowProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
        End If
        rs.Close
        cmbc036_2_CreateNowProc = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If

    '�O�H���̍\���̂����ݍH���̍\���̂փR�s�[
    BlkNow = BlkOld
    '�H���A�ԂɁ{�P����
    With BlkNow
        If BlkNow.KCNTC2 = "" Then BlkNow.KCNTC2 = "0"
        .KCNTC2 = CInt(.KCNTC2) + 1         '�H���A��
        .NEWKNTC2 = Kihon.NOWPROC           '�O�H��
        .GNWKNTC2 = Kihon.NEWPROC           '���ݍH��
        .SUMITLC2 = "0"                     'SUMMIT����
        .SUMITMC2 = "0"                     'SUMMIT����
        .SUMITWC2 = "0"                     'SUMMIT�d��
        .SUMITBC2 = "0"
    End With

    ''���ύXSTART SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '06/1/5 SMP�ΐ�
    sql = ""
    sql = sql & " SELECT"
    sql = sql & "   XTALCB"
    sql = sql & "  ,RLENCB"
    sql = sql & "  ,INPOSCB"
    sql = sql & "  ,HINBCB"
    sql = sql & "  ,REVNUMCB"
    sql = sql & "  ,FACTORYCB"
    sql = sql & "  ,OPECB"
    sql = sql & "  ,SXLIDCB"
    sql = sql & "  ,PLANTCATCB" ' 2007/09/05 SPK Tsutsumi Add
    sql = sql & " FROM"
    sql = sql & "   XSDCB"
    sql = sql & " WHERE XTALCB='" & StrCryNum & "'"
    '�����[�v���Ŕ���
    sql = sql & "   AND ((INPOSCB >=" & lngBeginIngotpos & ")"
    sql = sql & "   And (INPOSCB + RLENCB <= " & lngEndIngotpos & "))"
    sql = sql & "   AND LSTCCB <> 'H'"
    '�������W�b�N(SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs)
'    sql = "SELECT * from TBCME042 "
'    sql = sql & " WHERE CRYNUM='" & strCryNum & "'"
'    '�����[�v���Ŕ���
'    sql = sql & "   AND ((INGOTPOS >=" & lngBeginIngotpos & ") And (INGOTPOS + LENGTH <= " & lngEndIngotpos & "))"
'''''    sql = sql & "   AND ((INGOTPOS >=" & lngBeginIngotpos & ") And (INGOTPOS  <= " & lngEndIngotpos & "))"
'    sql = sql & "   AND LSTATCLS<>'H'"
    ''���ύXSTART SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '06/1/5 SMP�ΐ�


    Debug.Print sql
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)


    intLoopCnt = 0
''''    BlkNow.GNLC2 = 0    '���ݍH���i�u���b�N�j�̒������N���A���Ă���
    BlkNow.GNMC2 = 0
    Do While Not rs.EOF
        ReDim Preserve HinNow(intLoopCnt) As typ_XSDCA_Update
        '�O�H���̍\���̂����ݍH���̍\���̂փR�s�[
''''        HinNow(intLoopCnt) = HinOld(intHinOldCnt)

        ''���ύXSTART SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '06/1/5 SMP�ΐ�
        If IsNull(rs("XTALCB")) = False Then StrCryNum = rs("XTALCB")               '�����ԍ�
        If IsNull(rs("RLENCB")) = False Then intSxlLength = rs("RLENCB")            '����
        If IsNull(rs("INPOSCB")) = False Then intSxlIngotPos = rs("INPOSCB")        '�ʒu
'        If IsNull(rs("CRYNUM")) = False Then strCryNum = rs("CRYNUM")               '�����ԍ�
'        If IsNull(rs("LENGTH")) = False Then intSxlLength = rs("LENGTH")            '����
'        If IsNull(rs("INGOTPOS")) = False Then intSxlIngotPos = rs("INGOTPOS")      '�ʒu
        ''���ύXSTART SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '06/1/5 SMP�ΐ�

        '-- �u���b�N�ƃV���O���̈ʒu�֌W�𔻒肵�A�������Z�o --------
        sp = intSxlIngotPos         '�V���O���J�n�ʒu
        ep = sp + intSxlLength      '�V���O���I�[�ʒu
        sbp = intBlkIngotPos        '�u���b�N�J�n�ʒu
        ebp = sbp + intBlkLength    '�u���b�N�I�[�ʒu

        '' �u���b�N��SXL�̒��Ɋ��S�Ɋ܂܂�Ă���ꍇ ---------
        If sp <= sbp And ep >= ebp Then

            intLength = intBlkLength                    '�u���b�N�Ǘ��̒������g�p
            intIngotPos = intBlkIngotPos

        '' �u���b�N��SXL�̊J�n�ʒu����ɂ���A���I�[�ʒu���������ꍇ ---------
        ElseIf sp >= sbp And ep <= ebp Then

            intLength = intSxlLength                  '�V���O���Ǘ��̒������g�p
            intIngotPos = intSxlIngotPos

        '' �u���b�N���ꕔSXL�ɂ������Ă���ꍇ
        '' (�u���b�N���㑤�B�������u���b�N�̏I�[��SXL�̊J�n�ʒu����v���Ȃ�����) ------------
        ElseIf sp > sbp And sp < ebp And sp <> ebp Then

            intLength = ebp - sp                        '�u���b�N�̏I�[�ʒu - �V���O���̊J�n�ʒu
            intIngotPos = intSxlIngotPos

        '' �u���b�N���ꕔSXL�ɂ������Ă���ꍇ
        '' (�u���b�N�������B������SXL�̏I�[�ƃu���b�N�̊J�n�ʒu����v���Ȃ�����) ----------
        ElseIf sp < sbp And ep > sbp And ep <> sbp Then

            intLength = ep - sbp                        '�V���O���̏I�[�ʒu - �u���b�N�̊J�n�ʒu
            intIngotPos = intBlkIngotPos

        Else

''''            intLength = 0
''''            intIngotPos = intBlkIngotPos
            GoTo LoopNext

        End If
        '----------------------------------------------------

        '���ݍH���ҏW
        With HinNow(intLoopCnt)
            ''���ύXSTART SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '06/1/5 SMP�ΐ�
            If IsNull(rs("XTALCB")) = False Then .XTALCA = rs("XTALCB")
            .CRYNUMCA = StrBlockId         '�u���b�NID
            If IsNull(rs("HINBCB")) = False Then .HINBCA = rs("HINBCB")             '�i��
            If IsNull(rs("REVNUMCB")) = False Then .REVNUMCA = rs("REVNUMCB")       '���i�ԍ������ԍ�
            If IsNull(rs("FACTORYCB")) = False Then .FACTORYCA = rs("FACTORYCB")    '�H��
            If IsNull(rs("OPECB")) = False Then .OPECA = rs("OPECB")        '���Ə���

'            If IsNull(rs("CRYNUM")) = False Then .XTALCA = rs("CRYNUM")
'            .CRYNUMCA = strBlockID         '�u���b�NID
'            If IsNull(rs("HINBAN")) = False Then .HINBCA = rs("HINBAN")         '�i��
'            If IsNull(rs("REVNUM")) = False Then .REVNUMCA = rs("REVNUM")       '���i�ԍ������ԍ�
'            If IsNull(rs("FACTORY")) = False Then .FACTORYCA = rs("FACTORY")    '�H��
'            If IsNull(rs("OPECOND")) = False Then .OPECA = rs("OPECOND")        '���Ə���
            ''���ύXSTART SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '06/1/5 SMP�ΐ�
            .INPOSCA = intIngotPos    '�������J�n�ʒu
            .GNLCA = intLength          '����
'            BlkNow.GNLC2 = CStr(CLng(BlkNow.GNLC2) + CLng(HinNow(intLoopCnt).GNLCA))  '����
            ''���ύXSTART SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '06/1/5 SMP�ΐ�
            If IsNull(rs("SXLIDCB")) = False Then .SXLIDCA = rs("SXLIDCB")          '�V���O��ID
'            If IsNull(rs("SXLID")) = False Then .SXLIDCA = rs("SXLID")          '�V���O��ID
            ''���ύXSTART SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '06/1/5 SMP�ΐ�

            If IsNull(rs("PLANTCATCB")) = False Then .PLANTCATCA = rs("PLANTCATCB") '���� 2007/09/04 SPK Tsutsumi Add

            .SUMITBCA = 0
            .SUMITLCA = 0
            .SUMITMCA = 0
            .SUMITWCA = 0
            .NEWKNTCA = Kihon.NOWPROC   '�O�H��
            .GNWKNTCA = Kihon.NEWPROC   '���ݍH��
            .KCKNTCA = BlkNow.KCNTC2    '�H���A��
            .NEMACOCA = BlkNow.NEMACOC2 '�ŏI�ʉߏ�����
            .GNMACOCA = BlkNow.GNMACOC2 '���ݏ�����
''''        .XTALCA = strCryNum         '�����ԍ�
            '���ݏd�ʂ����߂�
            If GetDiameter(StrBlockId, dblDiameter) = FUNCTION_RETURN_FAILURE Then  '���a�����߂�
                dblDiameter = 0
''''                GoTo proc_wxit
            End If
            '��{���̒��a�Z�b�g
            Kihon.DIAMETER = dblDiameter

            '�擾�������a�����ɏd�ʂ����߂�
            .GNWCA = CStr(CLng(WeightOfCylinder(dblDiameter, CDbl(.GNLCA))))

            'add start hitec)matsumoto WF�}�b�vð��ق�薇���擾
            sql = "SELECT LOTID from TBCMY011 "
            sql = sql & " WHERE MSXLID='" & .SXLIDCA & "'"
            sql = sql & " AND LOTID='" & .CRYNUMCA & "'"
            sql = sql & " AND TO_NUMBER(WFSTA) <= 1"

            Debug.Print sql
            Set rs2 = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            iWFcnt = 0
            Do While Not rs2.EOF
                iWFcnt = iWFcnt + 1
                rs2.MoveNext
            Loop
            rs2.Close
            Debug.Print .SXLIDCA & " = " & iWFcnt & "��"
            HinNow(intLoopCnt).GNMCA = iWFcnt   'add 2003/03/29 hitec)matsumoto ���WF�}�b�v�e�[�u�����疇���J�E���g�擾���Ă���̂ŁA�����Ǖi�����Ƃ���
            BlkNow.GNMC2 = BlkNow.GNMC2 + iWFcnt
            Debug.Print "BlkNow.GNMC2 = " & BlkNow.GNMC2 & "��"
            .SUMITLCA = .GNLCA   '' 03/05/13 �㓡
            .SUMITMCA = .GNMCA
            .SUMITWCA = .GNWCA
''''            '���ݖ��������߂�
''''            If WfCount(strBlockID, CLng(.GNLCA), intNum) = FUNCTION_RETURN_FAILURE Then
''''                .GNMCA = 0
''''''''                GoTo proc_wxit
''''            Else
''''                .GNMCA = intNum
''''            End If
        End With

        With BlkNow
            '���ݏd�ʂ����߂�
            If GetDiameter(StrBlockId, dblDiameter) = FUNCTION_RETURN_FAILURE Then  '���a�����߂�
                dblDiameter = 0
    ''''                GoTo proc_wxit
            End If
            '��{���̒��a�Z�b�g
            Kihon.DIAMETER = dblDiameter
            '�擾�������a�����ɏd�ʂ����߂�
'            .GNWC2 = CStr(CLng(WeightOfCylinder(dblDiameter, CDbl(.GNLC2))))
            '���ݖ��������߂�
'            If WfCount(strBlockID, CLng(.GNLC2), intNum) = FUNCTION_RETURN_FAILURE Then
'                .GNMC2 = 0
''''                GoTo proc_wxit
'            Else
'                .GNMC2 = intNum
'            End If

        End With
        intLoopCnt = intLoopCnt + 1
        '�Ǖi�����Z�b�g
        With Kihon
            .CNTHINNOW = intLoopCnt
        End With

LoopNext:

        rs.MoveNext
    Loop



    rs.Close

    Debug.Print " WF�}�b�vð��ق�薇���擾 " & BlkNow.GNMC2 & "�� : �O�H��" & BlkOld.GNMC2 & "��"

    '�O�H���̒����ƌ��ݍH���̒���������ׁA�s�ǂ����݂��邩����
'    If CInt(BlkNow.GNLC2) = CInt(BlkOld.GNLC2) Then '�s�ǂȂ�
    If CInt(BlkNow.GNMC2) = CInt(BlkOld.GNMC2) Then '�s�ǂȂ�
        '��{���\����
        With Kihon
            .FURYOUMU = "N"
        End With
    Else                                            '�s�ǂ���
                strErrMsg = GetMsgStr("EWFM5", "�O�H��=" & BlkOld.GNMC2 & "�F���ݍH��=" & BlkNow.GNMC2) '03/06/06 �㓡
'                lblMsg.Caption = "WF�����s��v�G���["
                cmbc036_2_CreateNowProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit


    End If
    cmbc036_2_CreateNowProc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    cmbc036_2_CreateNowProc = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2002/09/10 ADD hitec)N.MATSUMOTO


'���������i�i�ԁj�O�H�����ю擾���\���̍쐬 2002/09/10 ADD hitec)N.MATSUMOTO
Public Function cmbc036_2_CreateXSDCA(ByVal StrBlockId As String, ByRef bNoData As Boolean) As FUNCTION_RETURN

    Dim iLoopCnt    As Integer
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim intProcNo   As Integer

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    intProcNo = 0

    '�u���b�NID�𓾂�
    sql = "SELECT * from XSDCA"
    sql = sql & " WHERE CRYNUMCA='" & StrBlockId & "'"
    sql = sql & "   AND LIVKCA= '0'"   '�����敪

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        cmbc036_2_CreateXSDCA = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    rs.MoveFirst
    iLoopCnt = 0
    'add start 2003/05/14 hitec)matsumoto ----------------------------------------
    BlkOld.GNLC2 = 0
    BlkOld.GNWC2 = 0
    BlkOld.GNMC2 = 0
    'add end   2003/05/14 hitec)matsumoto ----------------------------------------
    Do While Not rs.EOF
        ReDim Preserve HinOld(iLoopCnt)
        ReDim Preserve HinNow(iLoopCnt)
        With HinOld(iLoopCnt)
            If IsNull(rs.Fields("CRYNUMCA")) = False Then .CRYNUMCA = rs.Fields("CRYNUMCA")
            If IsNull(rs.Fields("HINBCA")) = False Then .HINBCA = rs.Fields("HINBCA")
            If IsNull(rs.Fields("INPOSCA")) = False Then .INPOSCA = rs.Fields("INPOSCA")
            If IsNull(rs.Fields("REVNUMCA")) = False Then .REVNUMCA = rs.Fields("REVNUMCA")
            If IsNull(rs.Fields("FACTORYCA")) = False Then .FACTORYCA = rs.Fields("FACTORYCA")
            If IsNull(rs.Fields("OPECA")) = False Then .OPECA = rs.Fields("OPECA")
            If IsNull(rs.Fields("KCKNTCA")) = False Then .KCKNTCA = rs.Fields("KCKNTCA")
            If IsNull(rs.Fields("SXLIDCA")) = False Then .SXLIDCA = rs.Fields("SXLIDCA")
            If IsNull(rs.Fields("XTALCA")) = False Then .XTALCA = rs.Fields("XTALCA")
            If IsNull(rs.Fields("NEKKNTCA")) = False Then .NEKKNTCA = rs.Fields("NEKKNTCA")
            If IsNull(rs.Fields("NEWKNTCA")) = False Then .NEWKNTCA = rs.Fields("NEWKNTCA")
            If IsNull(rs.Fields("NEWKKBCA")) = False Then .NEWKKBCA = rs.Fields("NEWKKBCA")
            If IsNull(rs.Fields("NEMACOCA")) = False Then .NEMACOCA = rs.Fields("NEMACOCA")
            If IsNull(rs.Fields("GNKKNTCA")) = False Then .GNKKNTCA = rs.Fields("GNKKNTCA")
            If IsNull(rs.Fields("GNWKNTCA")) = False Then .GNWKNTCA = rs.Fields("GNWKNTCA")
            If IsNull(rs.Fields("GNWKKBCA")) = False Then .GNWKKBCA = rs.Fields("GNWKKBCA")
            If IsNull(rs.Fields("GNMACOCA")) = False Then .GNMACOCA = rs.Fields("GNMACOCA")
            If IsNull(rs.Fields("GNDAYCA")) = False Then .GNDAYCA = rs.Fields("GNDAYCA")
            If IsNull(rs.Fields("GNLCA")) = False Then .GNLCA = rs.Fields("GNLCA")
            If IsNull(rs.Fields("GNWCA")) = False Then .GNWCA = rs.Fields("GNWCA")
            If IsNull(rs.Fields("GNMCA")) = False Then .GNMCA = rs.Fields("GNMCA")
            'add start 2003/05/14 hitec)matsumoto ----------------------------------------
            BlkOld.GNLC2 = CLng(BlkOld.GNLC2) + CLng(.GNLCA)
            BlkOld.GNWC2 = CLng(BlkOld.GNWC2) + CLng(.GNWCA)
            BlkOld.GNMC2 = CLng(BlkOld.GNMC2) + CLng(.GNMCA)
            'add end   2003/05/14 hitec)matsumoto ----------------------------------------
            If IsNull(rs.Fields("SUMITLCA")) = False Then .SUMITLCA = rs.Fields("SUMITLCA")
            If IsNull(rs.Fields("SUMITWCA")) = False Then .SUMITWCA = rs.Fields("SUMITWCA")
            If IsNull(rs.Fields("SUMITMCA")) = False Then .SUMITMCA = rs.Fields("SUMITMCA")
            If IsNull(rs.Fields("CHGCA")) = False Then .CHGCA = rs.Fields("CHGCA")
            If IsNull(rs.Fields("KAKOUBCA")) = False Then .KAKOUBCA = rs.Fields("KAKOUBCA")
            If IsNull(rs.Fields("KEIDAYCA")) = False Then .KEIDAYCA = rs.Fields("KEIDAYCA")
            If IsNull(rs.Fields("GNTKUBCA")) = False Then .GNTKUBCA = rs.Fields("GNTKUBCA")
            If IsNull(rs.Fields("GNTNOCA")) = False Then .GNTNOCA = rs.Fields("GNTNOCA")
            If IsNull(rs.Fields("XTWORKCA")) = False Then .XTWORKCA = rs.Fields("XTWORKCA")
            If IsNull(rs.Fields("WFWORKCA")) = False Then .WFWORKCA = rs.Fields("WFWORKCA")
            If IsNull(rs.Fields("LSTATBCA")) = False Then .LSTATBCA = rs.Fields("LSTATBCA")
            If IsNull(rs.Fields("RSTATBCA")) = False Then .RSTATBCA = rs.Fields("RSTATBCA")
            If IsNull(rs.Fields("LUFRCCA")) = False Then .LUFRCCA = rs.Fields("LUFRCCA")
            If IsNull(rs.Fields("LUFRBCA")) = False Then .LUFRBCA = rs.Fields("LUFRBCA")
            If IsNull(rs.Fields("LDFRCCA")) = False Then .LDFRCCA = rs.Fields("LDFRCCA")
            If IsNull(rs.Fields("LDFRBCA")) = False Then .LDFRBCA = rs.Fields("LDFRBCA")
            If IsNull(rs.Fields("HOLDCCA")) = False Then .HOLDCCA = rs.Fields("HOLDCCA")
            If IsNull(rs.Fields("HOLDBCA")) = False Then .HOLDBCA = rs.Fields("HOLDBCA")
            If IsNull(rs.Fields("EXKUBCA")) = False Then .EXKUBCA = rs.Fields("EXKUBCA")
            If IsNull(rs.Fields("HENPKCA")) = False Then .HENPKCA = rs.Fields("HENPKCA")
            If IsNull(rs.Fields("LIVKCA")) = False Then .LIVKCA = rs.Fields("LIVKCA")
            If IsNull(rs.Fields("KANKCA")) = False Then .KANKCA = rs.Fields("KANKCA")
            If IsNull(rs.Fields("NFCA")) = False Then .NFCA = rs.Fields("NFCA")
            If IsNull(rs.Fields("SAKJCA")) = False Then .SAKJCA = rs.Fields("SAKJCA")
            If IsNull(rs.Fields("TDAYCA")) = False Then .TDAYCA = rs.Fields("TDAYCA")
            If IsNull(rs.Fields("KDAYCA")) = False Then .KDAYCA = rs.Fields("KDAYCA")
            If IsNull(rs.Fields("SUMITBCA")) = False Then .SUMITBCA = rs.Fields("SUMITBCA")
            If IsNull(rs.Fields("SNDKCA")) = False Then .SNDKCA = rs.Fields("SNDKCA")
            If IsNull(rs.Fields("SNDDAYCA")) = False Then .SNDDAYCA = rs.Fields("SNDDAYCA")
            If IsNull(rs.Fields("PLANTCATCA")) = False Then .PLANTCATCA = rs.Fields("PLANTCATCA") ' 2007/09/04 SPK Tsutsumi Add
        End With
        '�Ǖi�����Z�b�g
        With Kihon
            .CNTHINOLD = iLoopCnt + 1
        End With
        iLoopCnt = iLoopCnt + 1
        rs.MoveNext
    Loop

    rs.Close
    cmbc036_2_CreateXSDCA = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    cmbc036_2_CreateXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2002/09/10 ADD hitec)N.MATSUMOTO
'�T�v    :�����w�� �u���b�N�h�c(SXL�����̃u���b�N�Ɍׂ�ꍇ�j���擾
'���Ұ�  :�ϐ���       ,IO   ,�^                                    ,����
'        :tBL_SXLID    ,IO   ,type_DBDRV_scmzc_fcmlc001d_LOTSXL     ,�u���b�N�h�c�A�r�w�k�h�c�\����
'        :��ؒl        ,O    ,FUNCTION_RETURN                       ,�ǂݍ��ݐ���
'����    :SXLID���u���b�N�h�c(SXL�����̃u���b�N�Ɍׂ�ꍇ�j���擾����
'����    :2003/2/25 Hitec)okazaki
Public Function DBDRV_BLOCKIDGET() As FUNCTION_RETURN
    Dim sql             As String
    Dim rs              As OraDynaset
    Dim i               As Long
    Dim inCnt           As Long
    Dim sDbName         As String
    Dim itUCount        As Integer
    Dim sBkBlkId        As String
    Dim sBkSxlId()      As String
    Dim iSxlCnt         As Integer
    Dim iLoopCnt        As Integer
    Dim iLoop           As Integer  '2003/05/28 HITEC)okazaki add
    Dim iLoop2          As Integer  '2003/05/28 HITEC)okazaki add
    Dim iLoop3          As Integer  '2003/05/28 HITEC)okazaki add
    Dim wkSXLID()       As type_DBDRV_LOTSXL
    Dim bCheckFlg       As Boolean
    Dim sBeforLotid     As String
    Dim iFLG            As Integer

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_BLOCKIDGET"
    DBDRV_BLOCKIDGET = FUNCTION_RETURN_FAILURE
    sDbName = "(V001)"
    ReDim wkSXLID(1)
    wkSXLID(1) = tSXLID(1)
    '���[�v�J�n
    For iLoop2 = 1 To 10        '�i�v���[�v�h�~�̂��ߍő�P�O��Ŕ�����
        ReDim sBkSxlId(1)
        iSxlCnt = 1
        '======================================================================
        '�O��̃��[�v�Ŏ擾�����u���b�N�S�Ă��SXL���擾�i�d�����ƂȂ��Ă���j
        '======================================================================
        For iLoopCnt = 1 To UBound(wkSXLID)
            iFLG = 0
            If iLoopCnt = 1 Then
                iFLG = 1
            ElseIf wkSXLID(iLoopCnt).LOTID <> wkSXLID(iLoopCnt - 1).LOTID Then
                iFLG = 1
            End If
            If iFLG <> 0 Then
                ' SXLID�̎擾
                sql = "select"
                sql = sql & " SXLIDCA"
                sql = sql & " from XSDCA "
                sql = sql & " where CRYNUMCA ='" & wkSXLID(iLoopCnt).LOTID & "'"
                sql = sql & "   and LIVKCA = '0'"
                sql = sql & " ORDER BY SXLIDCA"

                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

                '''���o���R�[�h�����݂Ȃ�ΊY��
                If Not rs.EOF Then
                    rs.MoveFirst
                    Do While Not rs.EOF
                        ReDim Preserve sBkSxlId(iSxlCnt)
                        sBkSxlId(iSxlCnt) = rs.Fields("SXLIDCA")
                        iSxlCnt = iSxlCnt + 1
                        rs.MoveNext
                    Loop
                End If
                rs.Close
            End If
        Next iLoopCnt
        If iSxlCnt = 1 Then
            Exit Function
        End If
        itUCount = UBound(tSXLID)
        '=============================================
        '�擾����SXL���SXL�EBLOCK�̑g�ݍ��킹�擾
        '=============================================
        For iLoopCnt = 1 To iSxlCnt - 1
            sql = "select"
            sql = sql & " CRYNUMCA,SXLIDCA"
            sql = sql & " from XSDCA "
            sql = sql & " where SXLIDCA ='" & sBkSxlId(iLoopCnt) & "'"
            sql = sql & "   and LIVKCA = '0'"
        '        sql = sql & " AND NOR Lotid ='" & tBL_SXLID(i).lotid & "'"
            sql = sql & " ORDER BY CRYNUMCA,SXLIDCA"

            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

            '''���o���R�[�h�����݂Ȃ�ΊY��
            If Not rs.EOF Then
                '�f���o���R�[�h�����ׂĎ擾�i���[�v�j
                Do While Not rs.EOF
                    '�f�z��ɂ��̑g�ݍ��킹��ǉ�����
                    If itUCount = 1 Then  '1�i�܂����b�g�����Ă��Ȃ���ԁj
                        With tSXLID(itUCount)
                            .SXLID = rs.Fields("SXLIDCA")         'SXLIDCA
                            .LOTID = rs.Fields("CRYNUMCA")
                        End With
                        itUCount = itUCount + 1

                    Else    '�Ώۃ��b�g��������������
                        bCheckFlg = False
                        For iLoop3 = 1 To UBound(tSXLID)
                            If tSXLID(iLoop3).SXLID = rs.Fields("SXLIDCA") And _
                               tSXLID(iLoop3).LOTID = rs.Fields("CRYNUMCA") Then
                               bCheckFlg = True
                               Exit For
                            End If
                        Next iLoop3
                        If bCheckFlg = False Then
                            ReDim Preserve tSXLID(itUCount)  '�z��̍Ē�`
                            With tSXLID(itUCount)
                                .SXLID = rs.Fields("SXLIDCA")         'SXLIDCA
                                .LOTID = rs.Fields("CRYNUMCA")
                            End With
                            itUCount = itUCount + 1
                        End If
                    End If
                    rs.MoveNext
                Loop
                rs.Close
            End If
        Next iLoopCnt

        '==================================================
        '�g�ݍ��킹���O��̃��[�v�Ɠ����Ȃ�I��
        '==================================================
        If UBound(tSXLID) = UBound(wkSXLID) Then
            Exit For
        End If
        ReDim wkSXLID(UBound(tSXLID))
        wkSXLID = tSXLID
    Next iLoop2


    ' �z����̃\�[�g���K�v
    iSxlCnt = UBound(tSXLID)
    itUCount = 1

    sql = "select"
    ''���C��START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
    ''  �u���b�N�ASXL�ɑ΂��ĕ����i�Ԃ��L�蓾��悤�ɂȂ����׏C��
    '�@�@�������ʒu���擾
'    sql = sql & " CRYNUMCA,SXLIDCA"
    sql = sql & " distinct CRYNUMCA,SXLIDCA,min(INPOSCA) INPOSCA"
    ''���C��START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
    sql = sql & " from  XSDCA "
    sql = sql & " where SXLIDCA in ("
    For iLoopCnt = 1 To iSxlCnt
        sql = sql & "'" & tSXLID(iLoopCnt).SXLID & "'"
        If iLoopCnt <> iSxlCnt Then sql = sql & ","
    Next
    sql = sql & " )"
    sql = sql & "   and LIVKCA ='0'"
    ''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
    sql = sql & " GROUP BY CRYNUMCA,SXLIDCA"
    ''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
    sql = sql & " ORDER BY CRYNUMCA,SXLIDCA"

    ReDim tSXLID(0)  '�z��̍Ē�`

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If Not rs.EOF Then
        '�f���o���R�[�h�����ׂĎ擾�i���[�v�j
        Do While Not rs.EOF
            ReDim Preserve tSXLID(itUCount)         '�z��̍Ē�`
            With tSXLID(itUCount)
                .SXLID = rs.Fields("SXLIDCA")       'SXLIDCA
                .LOTID = rs.Fields("CRYNUMCA")
                ''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
                .INGOTPOS = rs.Fields("INPOSCA")
                ''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
            End With
            itUCount = itUCount + 1
            rs.MoveNext
        Loop
        rs.Close
    End If

    DBDRV_BLOCKIDGET = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_BLOCKIDGET = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function



'�T�v    :�����w�� MIN,MAX�l���擾
'���Ұ�  :�ϐ���       ,IO   ,�^                                    ,����
'        :tBL_SXLID    ,IO   ,type_DBDRV_scmzc_fcmlc001d_LOTSXL     ,�u���b�N�h�c�A�r�w�k�h�c�\����
'        :��ؒl        ,O    ,FUNCTION_RETURN                       ,�ǂݍ��ݐ���
'����    :SXLID,BLOCKID���ő�A�ŏ��i�u���b�N�o�Ŕ���j�̃f�[�^���擾����
'����    :2003/2/25 Hitec)okazaki
Public Function DBDRV_MIN_MAX_SEQGET(ByRef iWfNum As Integer) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim i           As Long
    Dim inCnt       As Long
    Dim sDbName     As String
    Dim itUCount    As Integer
    Dim dblWFLen    As Double  '2003/04/25 hitec)okazaki
    Dim iRtn        As FUNCTION_RETURN
    Dim eps         As Double
    Dim sSmpKbn     As String
    Dim dblBlP      As Double
    Dim j, m, n     As Integer      '05/12/26 ooba

    ''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
    Dim lsBackHinban    As String
    ''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_MIN_MAX_SEQGET"

    eps = 0.000001        '�Â̐ݒ�
    iWfNum = 0
    itUCount = 1
    sDbName = "(Y011)"

    ReDim sWrpLOTID(0)      '05/12/26 ooba
    ReDim iWrpBLOCKSEQ(0)   '05/12/26 ooba

    'i = 0
    '�Ώ�SXL�͈�Ȃ̂Ń��[�v�Ȃ��i�V���O���P�ʂ̉�ʂȂ̂Łj
    '�f���[�v�J�n
    For i = 1 To UBound(tSXLID)
        ' SXLID�̎擾
        sql = "select "
        sql = sql & "LOTID,"                ' �u���b�NID"
        sql = sql & "MSXLID,"               ' SXLID"   'upd hitec)matsumoto �J�������ύX
        sql = sql & "blockseq,"             ' �u���b�N���A��"
        sql = sql & "WFSTA,"                ' WF���"
        sql = sql & "MHINBAN,"              ' �i��" 'upd hitec)matsumoto �J�������ύX
        sql = sql & "RTOP_POS,"             ' �_���u���b�N���ʒu"
        sql = sql & "RITOP_POS,"            ' �_���������ʒu"
        sql = sql & "MSMPLEID,"             ' �����ʒu"    'upd hitec)matsumoto �J�������ύX
        sql = sql & "SHAFLAG,"              ' �T���v���t���O"
        sql = sql & "INDTM,"
        sql = sql & "BASKETID,"
        sql = sql & "SLOTNO,"
        sql = sql & "CURRWPCS,"
        sql = sql & "EXISTFLG,"
        sql = sql & "TOP_POS,"
        sql = sql & "REJCAT,"
        sql = sql & "TXID,"
        sql = sql & "REGDATE,"
        sql = sql & "SUMMITSENDFLAG,"
        sql = sql & "SENDFLAG,"
        sql = sql & "SENDDATE,"
        sql = sql & "HREJCODE,"
        sql = sql & "UPDPROC,"
        sql = sql & "UPDDATE,"
        sql = sql & "MREVNUM,"
        sql = sql & "MFACTORY,"
        sql = sql & "MOPECOND,"
        sql = sql & "kankbn,"
        sql = sql & "NREJCODE"
        sql = sql & " from TBCMY011 "
        sql = sql & " where MSXLID='" & tSXLID(i).SXLID & "'"  'upd hitec)matsumoto �J�������ύX
        sql = sql & "   AND Lotid ='" & tSXLID(i).LOTID & "'"
        sql = sql & " ORDER BY Lotid,blockseq ASC"

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        '05/12/26 ooba START ==================================>
        m = UBound(sWrpLOTID)
        n = rs.RecordCount
        j = 0
        ReDim Preserve sWrpLOTID(m + n)         '��ۯ�ID
        ReDim Preserve iWrpBLOCKSEQ(m + n)      '��ۯ����A��
        '05/12/26 ooba END ====================================>

        '''���o���R�[�h�����݂Ȃ�ΊY��
        iWfNum = 0
        Do While Not rs.EOF
            ''���폜START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
            '1SXL-1�i�Ԃł͖����Ȃ����ׁA�v�Z���@��ς���
'            If CInt(rs.Fields("WFSTA")) <= 1 Then
'                iWfNum = iWfNum + 1
'            End If
            ''���폜START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
            '05/12/26 ooba START ==============================>
            j = j + 1
            '��ۯ�ID
            If IsNull(rs("LOTID")) Then
                sWrpLOTID(m + j) = ""
            Else
                sWrpLOTID(m + j) = rs("LOTID")
            End If
            '��ۯ����A��
            If IsNull(rs("BLOCKSEQ")) Then
                iWrpBLOCKSEQ(m + j) = 0
            Else
                iWrpBLOCKSEQ(m + j) = rs("BLOCKSEQ")
            End If
            '05/12/26 ooba END ================================>
            rs.MoveNext
        Loop
        If rs.RecordCount = 0 Then
            DBDRV_MIN_MAX_SEQGET = FUNCTION_RETURN_FAILURE
            f_cmbc036_2.lblMsg.Caption = GetMsgStr("EWFM6", "Y011") '03/06/06 �㓡
            rs.Close
            Exit Function
        End If

        rs.MoveFirst    '�擪ں��ނɈړ�
        Do While Not rs.EOF
            ReDim Preserve tExamine(itUCount)   '�z��̍Ē�`
            With tExamine(itUCount)
                If IsNull(rs!LOTID) = True Then
                    .LOTID = vbNullString           ' �u���b�NID
                 Else
                    .LOTID = rs!LOTID
                End If
                If IsNull(rs!MSXLID) = True Then
                    .SXLID = vbNullString
                Else
                    .SXLID = rs!MSXLID               ' SXLID
                End If
                .MinMax = 0                         ' 0:MIN 1:MAX
                If IsNull(rs!BLOCKSEQ) = True Then
                    .BLOCKSEQ = vbNullString
                Else
                    .BLOCKSEQ = rs!BLOCKSEQ         ' �u���b�N���A��
                End If
                If IsNull(rs!WFSTA) = True Then
                    .WFSTA = vbNullString
                Else
                    .WFSTA = rs!WFSTA               ' WF���
                End If
                If IsNull(rs!mhinban) = True Then
                    .hinban = vbNullString
                Else
                    .hinban = rs!mhinban             ' �i��
                End If
                If IsNull(rs!RTOP_POS) = True Then
                    .RTOP_POS = 0
                Else
                    'WF�ꖇ�̒����擾                                   '2003/04/25 hitec)okazaki
                    iRtn = DBDRV_WFLENGET(tSXLID(i).LOTID, dblWFLen)
                    '�u���b�N�擪�̕\���ʒu��WF�ꖇ�̒���������������   '2003/04/25 hitec)okazaki
''''                .RTOP_POS = Fix(CDbl(rs.Fields("RTOP_POS")) - dblWFLen + 0.9 + eps)         ' �_���u���b�N���ʒu  'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                    .RTOP_POS = Fix(CDbl(rs.Fields("RTOP_POS")) - dblWFLen + 0.99999)           ' �_���u���b�N���ʒu  'upd 2003/08/06 hitec)matsumoto
                End If
                If IsNull(rs!RITOP_POS) = True Then
                    .RITOP_POS = 0
                    .RINGOTPOS = 0
                Else
                    '�u���b�N�擪�̕\���ʒu��WF�ꖇ�̒���������������   '2003/04/25 hitec)okazaki
''''                .RITOP_POS = Fix(CDbl(rs.Fields("RITOP_POS")) - dblWFLen + 0.9 + eps)       ' �_���������ʒu    'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                    .RITOP_POS = Fix(CDbl(rs.Fields("RITOP_POS")) - dblWFLen + 0.99999)         ' �_���������ʒu    'upd 2003/08/06 hitec)matsumoto
                    .RINGOTPOS = CDbl(rs.Fields("RITOP_POS"))     ' 2003/04/30 hitec)okazaki �\�[�g�̋t�]��h�����ߒǉ�
                End If
                If IsNull(rs!MSMPLEID) = True Then
                    .SMPLEID = vbNullString
                Else
                    .SMPLEID = rs!MSMPLEID           ' �����ʒu
                End If
                If IsNull(rs!SHAFLAG) = True Then
                    .SHAFLAG = vbNullString
                Else
                    .SHAFLAG = rs!SHAFLAG           ' �T���v���t���O
                    If Trim(.SHAFLAG) = "1" Then
                        If Trim(.SMPLEID) = vbNullString Then   'add 2003/06/24 hitec)matsumoto �T���v���t���O��
                            DBDRV_MIN_MAX_SEQGET = FUNCTION_RETURN_FAILURE
                            f_cmbc036_2.lblMsg.Caption = GetMsgStr("ENSP4", "Y011")
                            rs.Close
                            Exit Function
                        End If
                    End If
                End If
                If IsNull(rs!INDTM) = True Then
                    .INDTM = vbNullString
                Else
                    .INDTM = rs!INDTM               ' �E�F�n�[�Z���^�[���ɓ���
                End If
                If IsNull(rs!BASKETID) = True Then
                    .BASKETID = vbNullString
                Else
                    .BASKETID = rs!BASKETID         ' �o�X�P�b�gID
                End If
                If IsNull(rs!SLOTNO) = True Then
                    .SLOTNO = vbNullString
                Else
                    .SLOTNO = rs!SLOTNO             ' �X���b�gNO
                End If
                ''���C��START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
                '1SXL-1�i�Ԃł͖����Ȃ����ׁA�v�Z���@��ς���
                iWfNum = 0
                If IsNull(rs!CURRWPCS) = True Then
                    .CURRWPCS = 0
                Else
                    If CInt(rs.Fields("WFSTA")) <= 1 Then
                        iWfNum = iWfNum + 1
                    End If
'                    .CURRWPCS = iWfNum              ' �E�F�n�[����
                End If

                '�������W�b�N
'                If IsNull(rs!CURRWPCS) = True Then
'                    .CURRWPCS = 0
'                Else
'                    .CURRWPCS = iWfNum              ' �E�F�n�[����
'                End If
                ''���C��START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
                If IsNull(rs!EXISTFLG) = True Then
                    .EXISTFLG = vbNullString
                Else
                    .EXISTFLG = rs!EXISTFLG         ' ���݃t���O
                End If

                If IsNull(rs!TOP_POS) = True Then   ' �u���b�N��TOP����̈ʒu
                    .TOP_POS = 0                    ' (��ʕ\���̂����h�����߁A�T���v���敪�ɂ��؂�グ
                Else                                '  �؎̂ď����ǉ� 2003/05/05)
                    dblBlP = CDbl(rs!TOP_POS)
                    .TOP_POS = Int(dblBlP / 10)             '�؎̂�
                End If
                If IsNull(rs!REJCAT) = True Then
                    .REJCAT = vbNullString
                Else
                    .REJCAT = rs!REJCAT             ' �������R
                End If
                If IsNull(rs!TXID) = True Then
                    .TXID = vbNullString
                Else
                    .TXID = rs!TXID                 ' �g�����U�N�V����ID
                End If
                If IsNull(rs!REGDATE) = True Then
                    .REGDATE = vbNullString
                Else
                    .REGDATE = rs!REGDATE           ' �o�^���t
                End If
                If IsNull(rs!SUMMITSENDFLAG) = True Then
                    .SUMMITSENDFLAG = vbNullString
                Else
                    .SUMMITSENDFLAG = rs!SUMMITSENDFLAG ' SUMIT���M�t���O
                End If
                If IsNull(rs!SENDFLAG) = True Then
                    .SENDFLAG = vbNullString
                Else
                    .SENDFLAG = rs!SENDFLAG         ' ���M�t���O
                End If
                If IsNull(rs!SENDDATE) = True Then
                    .SENDDATE = vbNullString
                Else
                    .SENDDATE = rs!SENDDATE         ' ���M���t
                End If
                If IsNull(rs!HREJCODE) = True Then
                    .HREJCODE = vbNullString
                Else
                    .HREJCODE = rs!HREJCODE         ' �s�Ǘ��R�R�[�h
                End If
                If IsNull(rs!UPDPROC) = True Then
                    .UPDPROC = vbNullString
                Else
                    .UPDPROC = rs!UPDPROC           ' �X�V�H��
                End If
                If IsNull(rs!UPDDATE) = True Then
                    .UPDDATE = vbNullString
                Else
                    .UPDDATE = rs!UPDDATE           ' �X�V���t
                End If
                If IsNull(rs!MREVNUM) = True Then
                    .REVNUM = 0
                Else
                    .REVNUM = rs!MREVNUM             ' ���i�ԍ������ԍ�
                End If
                If IsNull(rs!Mfactory) = True Then
                    .factory = vbNullString
                Else
                    .factory = rs!Mfactory           ' �H��
                End If
                If IsNull(rs!Mopecond) = True Then
                    .opecond = vbNullString
                Else
                    .opecond = rs!Mopecond           ' ���Ə���
                End If
                If IsNull(rs!KANKBN) = True Then
                    .KANKBN = vbNullString
                Else
                    .KANKBN = rs!KANKBN             ' �����敪
                End If
                If IsNull(rs!NREJCODE) = True Then
                    .NREJCODE = vbNullString
                Else
                    .NREJCODE = rs!NREJCODE         ' �����ԓ����R�R�[�h
                End If
            End With

            ''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
            'SXL�̕ς��ڂ����łȂ��A�i�Ԃ̕ς��ڂł��T���v�������擾����
            lsBackHinban = Trim(rs!mhinban)
            Do While (1)
                rs.MoveNext
                ''�f�[�^�̏I���̏ꍇ�A��߂��ă��[�v�I��
                If rs.EOF Then
                    rs.MovePrevious
                    tExamine(itUCount).CURRWPCS = iWfNum              ' �E�F�n�[����
                    Exit Do
                End If
                ''�i�Ԃ��ς������A��߂��ă��[�v�I��
                If IsNull(rs!mhinban) = False Then
                    If lsBackHinban <> Trim(rs!mhinban) Then
                        rs.MovePrevious
                        tExamine(itUCount).CURRWPCS = iWfNum              ' �E�F�n�[����
                        Exit Do
                    End If
                End If
                '�E�F�[�n�����J�E���g�A�b�v
                If CInt(rs.Fields("WFSTA")) <= 1 Then
                    iWfNum = iWfNum + 1
                End If
                lsBackHinban = Trim(rs!mhinban)
            Loop
            ''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�


            '�ŏIں���
'            rs.MoveLast                             '�ŏIں��ނɈړ�
            itUCount = itUCount + 1
            ReDim Preserve tExamine(itUCount)    '�z��̍Ē�`
            With tExamine(itUCount)
                If IsNull(rs!LOTID) = True Then
                    .LOTID = vbNullString           ' �u���b�NID
                Else
                    .LOTID = rs!LOTID
                End If
                If IsNull(rs!MSXLID) = True Then
                    .SXLID = vbNullString
                Else
                    .SXLID = rs!MSXLID               ' SXLID
                End If
                .MinMax = 1                          ' 0:MIN 1:MAX
                If IsNull(rs!BLOCKSEQ) = True Then
                    .BLOCKSEQ = vbNullString
                Else
                    .BLOCKSEQ = rs!BLOCKSEQ         ' �u���b�N���A��
                End If
                If IsNull(rs!WFSTA) = True Then
                    .WFSTA = vbNullString
                Else
                    .WFSTA = rs!WFSTA               ' WF���
                End If
                If IsNull(rs!mhinban) = True Then
                    .hinban = vbNullString
                Else
                    .hinban = rs!mhinban             ' �i��
                End If
                If IsNull(rs!RTOP_POS) = True Then
                    .RTOP_POS = 0
                Else
'                   .RTOP_POS = Int(CDbl(rs.Fields("RTOP_POS")))                    ' �_���u���b�N���ʒu
                    .RTOP_POS = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + eps)        ' �_���u���b�N���ʒu   'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                End If
                If IsNull(rs!RITOP_POS) = True Then
                    .RITOP_POS = 0
                    .RINGOTPOS = 0
                Else
'                   .RITOP_POS = Int(CDbl(rs.Fields("RITOP_POS")))                  ' �_���������ʒu
                    .RITOP_POS = Int(CDbl(rs.Fields("RITOP_POS")) + 0.9 + eps)      ' �_���������ʒu 'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                    .RINGOTPOS = CDbl(rs.Fields("RITOP_POS"))     ' 2003/04/30 hitec)okazaki �\�[�g�̋t�]��h�����ߒǉ�
                End If
                If IsNull(rs!MSMPLEID) = True Then
                    .SMPLEID = vbNullString
                Else
                    .SMPLEID = rs!MSMPLEID           ' �����ʒu
                End If
                If IsNull(rs!SHAFLAG) = True Then
                    .SHAFLAG = vbNullString
                Else
                    .SHAFLAG = rs!SHAFLAG           ' �T���v���t���O
                    If Trim(.SHAFLAG) = "1" Then
                        If Trim(.SMPLEID) = vbNullString Then   'add 2003/06/24 hitec)matsumoto �T���v���t���O��
                            DBDRV_MIN_MAX_SEQGET = FUNCTION_RETURN_FAILURE
                            f_cmbc036_2.lblMsg.Caption = GetMsgStr("ENSP4", "Y011")
                            rs.Close
                            Exit Function
                        End If
                    End If
                End If
                If IsNull(rs!INDTM) = True Then
                    .INDTM = vbNullString
                Else
                    .INDTM = rs!INDTM               ' �E�F�n�[�Z���^�[���ɓ���
                End If
                If IsNull(rs!BASKETID) = True Then
                    .BASKETID = vbNullString
                Else
                    .BASKETID = rs!BASKETID         ' �o�X�P�b�gID
                End If
                If IsNull(rs!SLOTNO) = True Then
                    .SLOTNO = vbNullString
                Else
                    .SLOTNO = rs!SLOTNO             ' �X���b�gNO
                End If
                If IsNull(rs!CURRWPCS) = True Then
                    .CURRWPCS = vbNullString
                Else
                    .CURRWPCS = rs!CURRWPCS         ' �E�F�n�[����
                End If
                If IsNull(rs!EXISTFLG) = True Then
                    .EXISTFLG = vbNullString
                Else
                    .EXISTFLG = rs!EXISTFLG         ' ���݃t���O
                End If
                If IsNull(rs!TOP_POS) = True Then   ' �u���b�N��TOP����̈ʒu
                    .TOP_POS = 0                    ' (��ʕ\���̂����h�����߁A�T���v���敪�ɂ��؂�グ
                Else                                '  �؎̂ď����ǉ� 2003/05/05)
                    dblBlP = CDbl(rs!TOP_POS)
                     .TOP_POS = Int((dblBlP / 10) + 0.9)     '�؂�グ
                End If

                If IsNull(rs!REJCAT) = True Then
                    .REJCAT = vbNullString
                Else
                    .REJCAT = rs!REJCAT             ' �������R
                End If
                If IsNull(rs!TXID) = True Then
                    .TXID = vbNullString
                Else
                    .TXID = rs!TXID                 ' �g�����U�N�V����ID
                End If
                If IsNull(rs!REGDATE) = True Then
                    .REGDATE = vbNullString
                Else
                    .REGDATE = rs!REGDATE           ' �o�^���t
                End If
                If IsNull(rs!SUMMITSENDFLAG) = True Then
                    .SUMMITSENDFLAG = vbNullString
                Else
                    .SUMMITSENDFLAG = rs!SUMMITSENDFLAG ' SUMIT���M�t���O
                End If
                If IsNull(rs!SENDFLAG) = True Then
                    .SENDFLAG = vbNullString
                Else
                    .SENDFLAG = rs!SENDFLAG         ' ���M�t���O
                End If
                If IsNull(rs!SENDDATE) = True Then
                    .SENDDATE = vbNullString
                Else
                    .SENDDATE = rs!SENDDATE         ' ���M���t
                End If
                If IsNull(rs!HREJCODE) = True Then
                    .HREJCODE = vbNullString
                Else
                    .HREJCODE = rs!HREJCODE         ' �s�Ǘ��R�R�[�h
                End If
                If IsNull(rs!UPDPROC) = True Then
                    .UPDPROC = vbNullString
                Else
                    .UPDPROC = rs!UPDPROC           ' �X�V�H��
                End If
                If IsNull(rs!UPDDATE) = True Then
                    .UPDDATE = vbNullString
                Else
                    .UPDDATE = rs!UPDDATE           ' �X�V���t
                End If
                If IsNull(rs!MREVNUM) = True Then
                    .REVNUM = 0
                Else
                    .REVNUM = rs!MREVNUM             ' ���i�ԍ������ԍ�
                End If
                If IsNull(rs!Mfactory) = True Then
                    .factory = vbNullString
                Else
                    .factory = rs!Mfactory           ' �H��
                End If
                If IsNull(rs!Mopecond) = True Then
                    .opecond = vbNullString
                Else
                    .opecond = rs!Mopecond           ' ���Ə���
                End If
                If IsNull(rs!KANKBN) = True Then
                    .KANKBN = vbNullString
                Else
                    .KANKBN = rs!KANKBN             ' �����敪
                End If
                If IsNull(rs!NREJCODE) = True Then
                    .NREJCODE = vbNullString
                Else
                    .NREJCODE = rs!NREJCODE         ' �����ԓ����R�R�[�h
                End If
            End With ''''        End If
            itUCount = itUCount + 1
            rs.MoveNext
        Loop
    Next
    '�f���[�v�I��

    DBDRV_MIN_MAX_SEQGET = FUNCTION_RETURN_SUCCESS

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

'�T�v    :�����w���@�������ڂ��擾
'���Ұ�  :�ϐ���       ,IO   ,�^                                    ,����
'        :tBL_SXLID    ,IO   ,type_DBDRV_LOTSXL                     ,�u���b�N�h�c�A�r�w�k�h�c�\����
'        :��ؒl        ,O    ,FUNCTION_RETURN                       ,�ǂݍ��ݐ���
'����    :SXLID,BLOCKID���ő�A�ŏ��i�u���b�N�o�Ŕ���j�̃f�[�^���擾����
'����    :2003/2/25 Hitec)okazaki
Public Function DVDRV_KENSA_KOUMOKU(tKensa() As typ_XSDCW _
                                            ) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim i           As Long
''''Dim inCnt       As Long
    Dim sDbName     As String
    Dim itUCount    As Integer
''''Dim tHIN        As tFullHinban
''''Dim sOT1        As String
''''Dim sOT2        As String
''''Dim rtn         As FUNCTION_RETURN

    Dim iIdx        As Integer
    Dim iCnt        As Integer
    Dim iChk        As Integer


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc039_SQL.bas -- Function DVDRV_KENSA_KOUMOKU"

    sDbName = "(V001)"
    'i = 0

''  --TEST-- Y011����擾�����T���v���h�c���g�p����Ƌ��L�����łǂ�������т�����Ă��܂��̂ŕύX
''''itUCount = UBound(tExamine)
''''ReDim tKensa(itUCount)                      '�̈�Ē�`
    itUCount = UBound(tSXLID)
    ReDim tKensa(itUCount * 2)                  '�̈�Ē�`
    iIdx = 0

    ''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�
    '��Ŕ�r�ׂ̈Ɏg�p����̂ŏ���������
    For iIdx = 0 To itUCount * 2
        tKensa(iIdx).SXLIDCW = ""
    Next iIdx
    iIdx = 0
    ''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP�ΐ�

    '�f���[�v�J�n
    For i = 1 To itUCount
''''    If Trim(tExamine(i).SMPLEID) <> "" Then
        If Trim(tSXLID(i).SXLID) <> "" Then
            ' SXLID�̎擾
'            sql = "select "
'            sql = sql & "CRYNUM,"              '�����ԍ�
'            sql = sql & "INGOTPOS,"            '�������ʒu
'            sql = sql & "SMPKBN,"              '�T���v���敪
'            sql = sql & "SMPLID,"              '�T���v��ID
'            sql = sql & "HINBAN,"              '�i��
'            sql = sql & "REVNUM,"              '���i�ԍ������ԍ�
'            sql = sql & "FACTORY,"             '�H��
'            sql = sql & "OPECOND,"             '���Ə���
'            sql = sql & "KTKBN,"               '�m��敪
'            sql = sql & "WFINDRS,"             'WF�����w���iRs)
'            sql = sql & "WFINDOI,"             'WF�����w���iOi)
'            sql = sql & "WFINDB1,"             'WF�����w���iB1)
'            sql = sql & "WFINDB2,"             'WF�����w���iB2�j
'            sql = sql & "WFINDB3,"             'WF�����w���iB3)
'            sql = sql & "WFINDL1,"             'WF�����w���iL1)
'            sql = sql & "WFINDL2,"             'WF�����w���iL2)
'            sql = sql & "WFINDL3,"             'WF�����w���iL3)
'            sql = sql & "WFINDL4,"             'WF�����w���iL4)
'            sql = sql & "WFINDDS,"             'WF�����w���iDS)
'            sql = sql & "WFINDDZ,"             'WF�����w���iDZ)
'            sql = sql & "WFINDSP,"             'WF�����w���iSP)
'            sql = sql & "WFINDDO1,"            'WF�����w���iDO1)
'            sql = sql & "WFINDDO2,"            'WF�����w���iDO2)
'            sql = sql & "WFINDDO3,"            'WF�����w���iDO3)
'            'add start 2003/05/23 hitec)�㓡 -------------------------
'            sql = sql & "NVL(WFINDOT1,'0') as DOT1,"            ' WF�����w���iOT1)
'            sql = sql & "NVL(WFINDOT2,'0') as DOT2,"           ' WF�����w���iOT2)
'            'add end   2003/05/23 hitec)�㓡 -------------------------
'            sql = sql & "WFRESRS,"             'WF�������сiRs)
'            sql = sql & "WFRESOI,"             'WF�������сiOi)
'            sql = sql & "WFRESB1,"             'WF�������сiB1)
'            sql = sql & "WFRESB2,"             'WF�������сiB2�j
'            sql = sql & "WFRESB3,"             'WF�������сiB3)
'            sql = sql & "WFRESL1,"             'WF�������сiL1)
'            sql = sql & "WFRESL2,"             'WF�������сiL2)
'            sql = sql & "WFRESL3,"             'WF�������сiL3)
'            sql = sql & "WFRESL4,"             'WF�������сiL4)
'            sql = sql & "WFRESDS,"             'WF�������сiDS)
'            sql = sql & "WFRESDZ,"             'WF�������сiDZ)
'            sql = sql & "WFRESSP,"             'WF�������сiSP)
'            sql = sql & "WFRESDO1,"            'WF�������сiDO1)
'            sql = sql & "WFRESDO2,"            'WF�������сiDO2)
'            sql = sql & "WFRESDO3,"            'WF�������сiDO3)
'            'add start 2003/05/23 hitec)�㓡 -------------------------
'            sql = sql & "NVL(WFRESOT1,'0') as SOT1,"            ' WF�������сiOT1)
'            sql = sql & "NVL(WFRESOT2,'0') as SOT2,"            ' WF�������сiOT2)
'            'add end   2003/05/23 hitec)�㓡 -------------------------
'            sql = sql & "REGDATE,"             '�o�^���t
'            sql = sql & "UPDDATE,"             '�X�V���t
'            sql = sql & "SENDFLAG,"            '���M�t���O
'            sql = sql & "SENDDATE"             '���M���t
'
'            sql = sql & " from XSDCW "
'            sql = sql & " where SMPLID ='" & tExamine(i).SMPLEID & "'"

            sql = "select "
            sql = sql & "SXLIDCW,"
            sql = sql & "SMPKBNCW,"
            sql = sql & "TBKBNCW,"
            sql = sql & "REPSMPLIDCW,"
            sql = sql & "XTALCW,"
            sql = sql & "INPOSCW,"
            sql = sql & "HINBCW,"
            sql = sql & "REVNUMCW,"
            sql = sql & "FACTORYCW,"
            sql = sql & "OPECW,"
            sql = sql & "KTKBNCW,"
            sql = sql & "SMCRYNUMCW,"
            sql = sql & "WFSMPLIDRSCW,"
            sql = sql & "NVL(WFSMPLIDRS1CW,'0') as RS1,"
            sql = sql & "NVL(WFSMPLIDRS2CW,'0') as RS2,"
            sql = sql & "WFINDRSCW,"
            sql = sql & "WFRESRS1CW,"
            sql = sql & "WFRESRS2CW,"
            sql = sql & "WFSMPLIDOICW,"
            sql = sql & "WFINDOICW,"
            sql = sql & "WFRESOICW,"
            sql = sql & "WFSMPLIDB1CW,"
            sql = sql & "WFINDB1CW,"
            sql = sql & "WFRESB1CW,"
            sql = sql & "WFSMPLIDB2CW,"
            sql = sql & "WFINDB2CW,"
            sql = sql & "WFRESB2CW,"
            sql = sql & "WFSMPLIDB3CW,"
            sql = sql & "WFINDB3CW,"
            sql = sql & "WFRESB3CW,"
            sql = sql & "WFSMPLIDL1CW,"
            sql = sql & "WFINDL1CW,"
            sql = sql & "WFRESL1CW,"
            sql = sql & "WFSMPLIDL2CW,"
            sql = sql & "WFINDL2CW,"
            sql = sql & "WFRESL2CW,"
            sql = sql & "WFSMPLIDL3CW,"
            sql = sql & "WFINDL3CW,"
            sql = sql & "WFRESL3CW,"
            sql = sql & "WFSMPLIDL4CW,"
            sql = sql & "WFINDL4CW,"
            sql = sql & "WFRESL4CW,"
            sql = sql & "WFSMPLIDDSCW,"
            sql = sql & "WFINDDSCW,"
            sql = sql & "WFRESDSCW,"
            sql = sql & "WFSMPLIDDZCW,"
            sql = sql & "WFINDDZCW,"
            sql = sql & "WFRESDZCW,"
            sql = sql & "WFSMPLIDSPCW,"
            sql = sql & "WFINDSPCW,"
            sql = sql & "WFRESSPCW,"
            sql = sql & "WFSMPLIDDO1CW,"
            sql = sql & "WFINDDO1CW,"
            sql = sql & "WFRESDO1CW,"
            sql = sql & "WFSMPLIDDO2CW,"
            sql = sql & "WFINDDO2CW,"
            sql = sql & "WFRESDO2CW,"
            sql = sql & "WFSMPLIDDO3CW,"
            sql = sql & "WFINDDO3CW,"
            sql = sql & "WFRESDO3CW,"
            sql = sql & "WFSMPLIDOT1CW,"
            sql = sql & "WFSMPLIDOT2CW,"
            'add start 2003/05/23 hitec)�㓡 -------------------------
            sql = sql & "NVL(WFINDOT1CW,   '0')     as DOT1,"           ' WF�����w���iOT1)
            sql = sql & "NVL(WFINDOT2CW,   '0')     as DOT2,"           ' WF�����w���iOT2)
            'add end   2003/05/23 hitec)�㓡 -------------------------
            'add start 2003/05/23 hitec)�㓡 -------------------------
            sql = sql & "NVL(WFRESOT1CW,   '0')     as SOT1,"           ' WF�������сiOT1)
            sql = sql & "NVL(WFRESOT2CW,   '0')     as SOT2,"           ' WF�������сiOT2)
            'add end   2003/05/23 hitec)�㓡 -------------------------
            sql = sql & "NVL(WFSMPLIDAOICW,'0')     as sAOI,"
            sql = sql & "NVL(WFINDAOICW,   '0')     as iAOI,"
            sql = sql & "NVL(WFRESAOICW,   '0')     as rAOI,"
            sql = sql & "NVL(SMPLNUMCW,    '0')     as sNUM,"
            sql = sql & "NVL(SMPLPATCW,    '0')     as PAT,"
            sql = sql & "NVL(TSTAFFCW,     '0')     as STF,"
            sql = sql & "TDAYCW,"
            sql = sql & "NVL(KSTAFFCW,     '0')     as kSTF,"
            sql = sql & "KDAYCW,"
            sql = sql & "NVL(SNDKCW,       '0')     as SND,"
            sql = sql & "NVL(SNDDAYCW,'2003/09/18') as sDAY,"

            '' GD�ǉ��@05/01/31 ooba START =====================================>
            sql = sql & "NVL(WFSMPLIDGDCW,'0')     as sGD,"
            sql = sql & "NVL(WFINDGDCW,   '0')     as iGD,"
            sql = sql & "NVL(WFRESGDCW,   '0')     as rGD,"
            sql = sql & "NVL(WFHSGDCW,   '0')      as hGD"
            '' GD�ǉ��@05/01/31 ooba END =======================================>

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
            sql = sql & ",NVL(EPSMPLIDB1CW,   '0')  as EPSMPLIDB1CW,"
            sql = sql & "NVL(EPINDB1CW,   '0')      as EPINDB1CW,"
            sql = sql & "NVL(EPRESB1CW,   '0')      as EPRESB1CW,"
            sql = sql & "NVL(EPSMPLIDB2CW,   '0')   as EPSMPLIDB2CW,"
            sql = sql & "NVL(EPINDB2CW,   '0')      as EPINDB2CW,"
            sql = sql & "NVL(EPRESB2CW,   '0')      as EPRESB2CW,"
            sql = sql & "NVL(EPSMPLIDB3CW,   '0')   as EPSMPLIDB3CW,"
            sql = sql & "NVL(EPINDB3CW,   '0')      as EPINDB3CW,"
            sql = sql & "NVL(EPRESB3CW,   '0')      as EPRESB3CW,"
            sql = sql & "NVL(EPSMPLIDL1CW,   '0')   as EPSMPLIDL1CW,"
            sql = sql & "NVL(EPINDL1CW,   '0')      as EPINDL1CW,"
            sql = sql & "NVL(EPRESL1CW,   '0')      as EPRESL1CW,"
            sql = sql & "NVL(EPSMPLIDL2CW,   '0')   as EPSMPLIDL2CW,"
            sql = sql & "NVL(EPINDL2CW,   '0')      as EPINDL2CW,"
            sql = sql & "NVL(EPRESL2CW,   '0')      as EPRESL2CW,"
            sql = sql & "NVL(EPSMPLIDL3CW,   '0')   as EPSMPLIDL3CW,"
            sql = sql & "NVL(EPINDL3CW,   '0')      as EPINDL3CW,"
            sql = sql & "NVL(EPRESL3CW,   '0')      as EPRESL3CW"
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

            sql = sql & " from  XSDCW "
            sql = sql & " where SXLIDCW ='" & tSXLID(i).SXLID & "'"
            sql = sql & "   and LIVKCW  ='0'"                           ' �����敪�͕K���m�F���鎖
            sql = sql & " order by INPOSCW"
'''''       sql = sql & " where REPSMPLIDCW ='" & tExamine(i).SMPLEID & "'"

            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            '''���o���R�[�h�����݂Ȃ�ΊY��
            If Not rs.EOF Then
'                With tKensa(i)
'                   .CRYNUM = rs!CRYNUM
'                    .INGOTPOS = rs!INGOTPOS
'                    .SMPKBN = rs!SMPKBN
'                    .SMPLID = rs!SMPLID
'                    .hinban = rs!hinban
'                    .REVNUM = rs!REVNUM
'                    .factory = rs!factory
'                    .opecond = rs!opecond
'                    .KTKBN = rs!KTKBN
'                    .WFINDRS = rs!WFINDRS
'                    .WFINDOI = rs!WFINDOI
'                    .WFINDB1 = rs!WFINDB1
'                    .WFINDB2 = rs!WFINDB2
'                    .WFINDB3 = rs!WFINDB3
'                    .WFINDL1 = rs!WFINDL1
'                    .WFINDL2 = rs!WFINDL2
'                    .WFINDL3 = rs!WFINDL3
'                    .WFINDL4 = rs!WFINDL4
'                    .WFINDDS = rs!WFINDDS
'                    .WFINDDZ = rs!WFINDDZ
'                    .WFINDSP = rs!WFINDSP
'                    .WFINDDO1 = rs!WFINDDO1
'                    .WFINDDO2 = rs!WFINDDO2
'                    .WFINDDO3 = rs!WFINDDO3
'                    tHin.hinban = .hinban
'                    tHin.factory = .factory
'                    tHin.mnorevno = .REVNUM
'                    tHin.opecond = .opecond
'                    rtn = scmzc_getE036(tHin, sOT1, sOT2)
'                    If rtn = FUNCTION_RETURN_FAILURE Then
'                        rs.Close
'                        DVDRV_KENSA_KOUMOKU = FUNCTION_RETURN_FAILURE
'                        GoTo proc_exit
'                    End If
'                    If sOT1 = "1" Then
'                        .WFINDOT1 = rs!DOT1 '03/05/23
'                    Else
'                        .WFINDOT1 = 0 '03/05/23
'                    End If
'                    If sOT2 = "1" Then
'                        .WFINDOT2 = rs!DOT2 '03/05/23
'                    Else
'                        .WFINDOT2 = 0 '03/05/23
'                    End If
'                    .WFRESRS = rs!WFRESRS
'                    .WFRESOI = rs!WFRESOI
'                    .WFRESB1 = rs!WFRESB1
'                    .WFRESB2 = rs!WFRESB2
'                    .WFRESB3 = rs!WFRESB3
'                    .WFRESL1 = rs!WFRESL1
'                    .WFRESL2 = rs!WFRESL2
'                    .WFRESL3 = rs!WFRESL3
'                    .WFRESL4 = rs!WFRESL4
'                    .WFRESDS = rs!WFRESDS
'                    .WFRESDZ = rs!WFRESDZ
'                    .WFRESSP = rs!WFRESSP
'                    .WFRESDO1 = rs!WFRESDO1
'                    .WFRESDO2 = rs!WFRESDO2
'                    .WFRESDO3 = rs!WFRESDO3
'                    .WFRESOT1 = rs!sOT1 '03/05/23
'                    .WFRESOT2 = rs!sOT2 '03/05/23
'                    .REGDATE = rs!REGDATE
'                    .UPDDATE = rs!UPDDATE
'                    .SENDFLAG = rs!SENDFLAG
'                    .SENDDATE = rs!SENDDATE
'                End With

                iCnt = 0
                Do While Not rs.EOF
                    iIdx = iIdx + 1
                    iCnt = iCnt + 1
                    ' �R���ڈȍ~�����݂���ꍇ�G���[
                    If iCnt > 2 Then
                        Exit Do
                    End If

                    If rs!TBKBNCW = "T" Then
                        For iChk = 1 To iIdx - 1
                            If tKensa(iChk).SXLIDCW = rs!SXLIDCW And tKensa(iChk).TBKBNCW = rs!TBKBNCW Then
                                Exit For
                            End If
                        Next
                    Else
                        iChk = iIdx
                    End If

                    If iChk = iIdx Then
                        With tKensa(iIdx)
                            .SXLIDCW = rs!SXLIDCW
                            .SMPKBNCW = rs!SMPKBNCW
                            .TBKBNCW = rs!TBKBNCW
                            .REPSMPLIDCW = rs!REPSMPLIDCW
                            .XTALCW = rs!XTALCW
                            .INPOSCW = rs!INPOSCW
                            .HINBCW = rs!HINBCW
                            .REVNUMCW = rs!REVNUMCW
                            .FACTORYCW = rs!FACTORYCW
                            .OPECW = rs!OPECW
                            .KTKBNCW = rs!KTKBNCW
                            .SMCRYNUMCW = rs!SMCRYNUMCW
                            .WFSMPLIDRSCW = rs!WFSMPLIDRSCW
                            .WFSMPLIDRS1CW = rs!rs1
                            .WFSMPLIDRS2CW = rs!rs2
                            .WFINDRSCW = rs!WFINDRSCW
                            .WFRESRS1CW = rs!WFRESRS1CW
                            .WFSMPLIDOICW = rs!WFSMPLIDOICW
                            .WFINDOICW = rs!WFINDOICW
                            .WFRESOICW = rs!WFRESOICW
                            .WFSMPLIDB1CW = rs!WFSMPLIDB1CW
                            .WFINDB1CW = rs!WFINDB1CW
                            .WFRESB1CW = rs!WFRESB1CW
                            .WFSMPLIDB2CW = rs!WFSMPLIDB2CW
                            .WFINDB2CW = rs!WFINDB2CW
                            .WFRESB2CW = rs!WFRESB2CW
                            .WFSMPLIDB3CW = rs!WFSMPLIDB3CW
                            .WFINDB3CW = rs!WFINDB3CW
                            .WFRESB3CW = rs!WFRESB3CW
                            .WFSMPLIDL1CW = rs!WFSMPLIDL1CW
                            .WFINDL1CW = rs!WFINDL1CW
                            .WFRESL1CW = rs!WFRESL1CW
                            .WFSMPLIDL2CW = rs!WFSMPLIDL2CW
                            .WFINDL2CW = rs!WFINDL2CW
                            .WFRESL2CW = rs!WFRESL2CW
                            .WFSMPLIDL3CW = rs!WFSMPLIDL3CW
                            .WFINDL3CW = rs!WFINDL3CW
                            .WFRESL3CW = rs!WFRESL3CW
                            .WFSMPLIDL4CW = rs!WFSMPLIDL4CW
                            .WFINDL4CW = rs!WFINDL4CW
                            .WFRESL4CW = rs!WFRESL4CW
                            .WFSMPLIDDSCW = rs!WFSMPLIDDSCW
                            .WFINDDSCW = rs!WFINDDSCW
                            .WFRESDSCW = rs!WFRESDSCW
                            .WFSMPLIDDZCW = rs!WFSMPLIDDZCW
                            .WFINDDZCW = rs!WFINDDZCW
                            .WFRESDZCW = rs!WFRESDZCW
                            .WFSMPLIDSPCW = rs!WFSMPLIDSPCW
                            .WFINDSPCW = rs!WFINDSPCW
                            .WFRESSPCW = rs!WFRESSPCW
                            .WFSMPLIDDO1CW = rs!WFSMPLIDDO1CW
                            .WFINDDO1CW = rs!WFINDDO1CW
                            .WFRESDO1CW = rs!WFRESDO1CW
                            .WFSMPLIDDO2CW = rs!WFSMPLIDDO2CW
                            .WFINDDO2CW = rs!WFINDDO2CW
                            .WFRESDO2CW = rs!WFRESDO2CW
                            .WFSMPLIDDO3CW = rs!WFSMPLIDDO3CW
                            .WFINDDO3CW = rs!WFINDDO3CW
                            .WFRESDO3CW = rs!WFRESDO3CW
                            .WFSMPLIDOT1CW = rs!WFSMPLIDOT1CW
                            .WFINDOT1CW = rs!DOT1
                            .WFRESOT1CW = rs!sOT1
                            .WFSMPLIDOT2CW = rs!WFSMPLIDOT2CW
                            .WFINDOT2CW = rs!DOT2
                            .WFRESOT2CW = rs!sOT2
                            .WFSMPLIDAOICW = rs!sAOI
                            .WFINDAOICW = rs!iAOI
                            .WFRESAOICW = rs!rAOI
                            .SMPLNUMCW = rs!sNum
                            .SMPLPATCW = rs!PAT
                            .TSTAFFCW = rs!STF
                            .TDAYCW = rs!TDAYCW
                            .KSTAFFCW = rs!kSTF
                            .KDAYCW = rs!KDAYCW
                            .SNDKCW = rs!SND
                            .SNDDAYCW = rs!sDay
                            .WFSMPLIDGDCW = rs!sGD      '�����ID(GD)    '05/01/31 ooba START ====>
                            .WFINDGDCW = rs!iGD         '���FLG(GD)
                            .WFRESGDCW = rs!rGD         '����FLG(GD)
                            .WFHSGDCW = rs!hGD          '�ۏ�FLG(GD)    '05/01/31 ooba END ======>
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                            .EPSMPLIDB1CW = rs!EPSMPLIDB1CW
                            .EPINDB1CW = rs!EPINDB1CW
                            .EPRESB1CW = rs!EPRESB1CW
                            .EPSMPLIDB2CW = rs!EPSMPLIDB2CW
                            .EPINDB2CW = rs!EPINDB2CW
                            .EPRESB2CW = rs!EPRESB2CW
                            .EPSMPLIDB3CW = rs!EPSMPLIDB3CW
                            .EPINDB3CW = rs!EPINDB3CW
                            .EPRESB3CW = rs!EPRESB3CW
                            .EPSMPLIDL1CW = rs!EPSMPLIDL1CW
                            .EPINDL1CW = rs!EPINDL1CW
                            .EPRESL1CW = rs!EPRESL1CW
                            .EPSMPLIDL2CW = rs!EPSMPLIDL2CW
                            .EPINDL2CW = rs!EPINDL2CW
                            .EPRESL2CW = rs!EPRESL2CW
                            .EPSMPLIDL3CW = rs!EPSMPLIDL3CW
                            .EPINDL3CW = rs!EPINDL3CW
                            .EPRESL3CW = rs!EPRESL3CW
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
                        End With
                    End If


                    If rs!TBKBNCW = "B" Then
                        For iChk = iIdx - 1 To 1 Step -1
                            If tKensa(iChk).SXLIDCW = rs!SXLIDCW And tKensa(iChk).TBKBNCW = rs!TBKBNCW Then
                                Exit For
                            End If
                        Next
                        If iChk > 0 Then
                            tKensa(iChk) = tKensa(0)
                        End If
                    End If


                    rs.MoveNext
                Loop
                rs.Close

                ' �擾�������Q���łȂ��ꍇ�G���[
                If iCnt <> 2 Then
                    f_cmbc036_2.lblMsg.Caption = GetMsgStr("ENSP2")
                    DVDRV_KENSA_KOUMOKU = FUNCTION_RETURN_FAILURE
                    Exit Function
                End If
           Else
                rs.Close
                f_cmbc036_2.lblMsg.Caption = GetMsgStr("ENSP2")    '03/06/06 �㓡
                DVDRV_KENSA_KOUMOKU = FUNCTION_RETURN_FAILURE
                Exit Function
            End If

        End If
    Next i
    '�f���[�v�I��

    DVDRV_KENSA_KOUMOKU = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    GoTo proc_exit
End Function

'�T�v    :�����w���@�������ڂ��擾
'���Ұ�  :�ϐ���       ,IO   ,�^                 ,����
'        :tWafSmp      ,I    ,typ_XSDCW          ,����يǗ��\����(�����S��)
'        :tKensa       ,O    ,typ_XSDCW          ,����يǗ��\����
'        :��ؒl        ,O    ,FUNCTION_RETURN    ,�ǂݍ��ݐ���
'����    :
'����    :08/02/04 ooba
Public Function DVDRV_KENSA_KOUMOKU_LOCAL(tWafSmp() As typ_XSDCW, tKensa() As typ_XSDCW) As FUNCTION_RETURN

    Dim i, j        As Integer
    Dim sql         As String
    Dim recCnt      As Integer
    Dim iChk        As Integer
    Dim bTflg       As Boolean
    Dim bBflg       As Boolean
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function DVDRV_KENSA_KOUMOKU_LOCAL"
    
    ReDim tKensa(UBound(tSXLID) * 2)
    recCnt = 0
    
    '������
    For i = 0 To UBound(tKensa)
        tKensa(i).SXLIDCW = ""
    Next i
    
    For i = 1 To UBound(tSXLID)
        bTflg = False
        bBflg = False
        'TOP
        For iChk = 1 To UBound(tSXLID)
            If tSXLID(iChk).SXLID = tSXLID(i).SXLID Then Exit For
        Next iChk
        If iChk = i Then
            For j = 1 To UBound(tWafSmp)
                If tSXLID(i).SXLID = tWafSmp(j).SXLIDCW And tWafSmp(j).TBKBNCW = "T" Then
                    recCnt = recCnt + 1
                    tKensa(recCnt) = tWafSmp(j)
                    bTflg = True
                    Exit For
                End If
            Next j
        Else
            recCnt = recCnt + 1
            tKensa(recCnt) = tKensa(0)
            bTflg = True
        End If
        
        'BOT
        For iChk = UBound(tSXLID) To 1 Step -1
            If tSXLID(iChk).SXLID = tSXLID(i).SXLID Then Exit For
        Next iChk
        If iChk = i Then
            For j = 1 To UBound(tWafSmp)
                If tSXLID(i).SXLID = tWafSmp(j).SXLIDCW And tWafSmp(j).TBKBNCW = "B" Then
                    recCnt = recCnt + 1
                    tKensa(recCnt) = tWafSmp(j)
                    bBflg = True
                    Exit For
                End If
            Next j
        Else
            recCnt = recCnt + 1
            tKensa(recCnt) = tKensa(0)
            bBflg = True
        End If
        
        '��������
        If bTflg = False Or bBflg = False Then
            f_cmbc036_2.lblMsg.Caption = GetMsgStr("ENSP2")
            DVDRV_KENSA_KOUMOKU_LOCAL = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
    Next i
    
    DVDRV_KENSA_KOUMOKU_LOCAL = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    GoTo proc_exit
    
End Function

'�T�v    :�����w�� �������̃u���b�N�������J�n�ʒu���擾
'���Ұ�  :�ϐ���       ,IO   ,�^                                    ,����
'        :��ؒl        ,O    ,FUNCTION_RETURN                       ,�ǂݍ��ݐ���
'����    :SXLID,BLOCKID���ő�A�ŏ��i�u���b�N�o�Ŕ���j�̃f�[�^���擾����
'����    :2003/3/05 Hitec)okazaki
Public Function DVDRV_KETURAKU_Ingotget(ByVal sLotid As String, iIngotpos As Integer) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
''''Dim i           As Long
''''Dim inCnt       As Long
    Dim sDbName     As String
''''Dim itUCount    As Integer

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc039_SQL.bas -- Function DVDRV_KETURAKU_Ingotget"

    sDbName = "(V001)"

    sql = "select INGOTPOS"            ' �������J�n�ʒu
    sql = sql & " from  TBCME040"
    sql = sql & " where BLOCKID = '" & sLotid & "'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    '''���o���R�[�h�����݂Ȃ�ΊY��
    If Not rs.EOF Then
        iIngotpos = Int(CDbl(rs!INGOTPOS))
    End If
    rs.Close

    DVDRV_KETURAKU_Ingotget = FUNCTION_RETURN_SUCCESS

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



'2003/02/28 hitec)okazaki ADD end
'********************************************************************************



'�T�v    :�����w�� ���͂����u���b�N�o����A�Y���v�e������
'���Ұ�  :�ϐ���       ,IO   ,�^                                    ,����
'        :iBlkP    �@�@,O    ,integer                               ,�u���b�NP
'        :��ؒl        ,O    ,FUNCTION_RETURN                       ,�ǂݍ��ݐ���
'����    :�����w�� ���͂����u���b�N�o����A�Y���v�e������
'����    :2003/2/25 Hitec)matsumoto
Public Function DBDRV_GET_WFMAP(ByVal sBlkId As String, ByVal iBlkP As Integer, _
                                ByRef sBlkP As Variant, ByRef sKessyoP As Variant, _
                                ByRef sBlkSeq As Variant, ByRef sBlkSeq2 As Variant, ByRef sSmpId1 As Variant, _
                                ByRef sSmpId2 As Variant, ByRef iNextBlkP As Integer, _
                                ByRef vWfNum As Variant, iKbnFlg As Integer) _
                                    As FUNCTION_RETURN

    Dim sql         As String
    Dim rs          As OraDynaset
    Dim i, j        As Long
    Dim inCnt       As Long
    Dim sDbName     As String
    Dim iLoopCnt    As Integer
    Dim dChkBlkP    As Double
'   Dim dChkBlkP    As Double
    Dim iTopPos     As Integer
    Dim sAddSmpId1  As String
    Dim sAddSmpId2  As String
    Dim iBlkflg     As Integer
    Dim vBlkId      As Variant
    Dim sSXLID      As String
    Dim dblWFLen    As Double
    Dim eps         As Double

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_GET_WFMAP"

    sDbName = "(Y011)"
    i = 0
    eps = 0.000001

    sql = "select "
    sql = sql & "LOTID,"                ' �u���b�NID"
''''sql = sql & "SXLID,"                ' SXLID"
    sql = sql & "MSXLID,"               ' SXLID"
    sql = sql & "blockseq,"             ' �u���b�N���A��"
    sql = sql & "WFSTA,"                ' WF���"
    sql = sql & "RTOP_POS,"             ' �_���u���b�N���ʒu"
    sql = sql & "RITOP_POS,"            ' �_���������ʒu"
    sql = sql & "MSMPLEID,"             ' �����ʒu"
    sql = sql & "SHAFLAG,"              ' �T���v���t���O"
    sql = sql & "TOP_POS"               ' �u���b�N���ʒu
    sql = sql & " from TBCMY011 "
''''sql = sql & " where SXLID ='" & sSxlId & "'"
    sql = sql & " where LOTID ='" & sBlkId & "'"
    sql = sql & "   AND TO_NUMBER(WFSTA) <= 1"
    sql = sql & " ORDER BY BLOCKSEQ ASC"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    iLoopCnt = 0
    vWfNum = 0
    Do While Not rs.EOF
        If IsNull(rs.Fields("RTOP_POS")) = True Then
            dChkBlkP = 0
        Else
            dChkBlkP = CDbl(rs.Fields("RTOP_POS"))
        End If
        If (iBlkP < dChkBlkP) And (dChkBlkP <= iNextBlkP) Then
            vWfNum = CInt(vWfNum) + 1
        End If
        rs.MoveNext
    Loop
    rs.Close

    sql = "select "
    sql = sql & "LOTID,"                ' �u���b�NID"
    sql = sql & "MSXLID,"               ' SXLID"
    sql = sql & "blockseq,"             ' �u���b�N���A��"
    sql = sql & "WFSTA,"                ' WF���"
    sql = sql & "RTOP_POS,"             ' �_���u���b�N���ʒu"
    sql = sql & "RITOP_POS,"            ' �_���������ʒu"
    sql = sql & "MSMPLEID,"             ' �����ʒu"
    sql = sql & "SHAFLAG,"              ' �T���v���t���O"
    sql = sql & "TOP_POS"               ' �u���b�N���ʒu
    sql = sql & " from TBCMY011 "
''''sql = sql & " where SXLID ='" & sSxlId & "'"
    sql = sql & " where LOTID ='" & sBlkId & "'"
''''sql = sql & "   AND TO_NUMBER(WFSTA) <= 1"
    sql = sql & " ORDER BY BLOCKSEQ ASC"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    iLoopCnt = 0
    rs.MoveFirst
    Do While Not rs.EOF
        Select Case Right(sSmpId1, 1)
            Case "T"
                If iKbnFlg = 0 Then     '�O�u���b�N�̈ʒu�A�Ǝ��u���b�N��T
                    If IsNull(rs.Fields("RTOP_POS")) = False Then
                        dChkBlkP = CDbl(rs.Fields("RTOP_POS"))
                    End If
                    If dChkBlkP > iBlkP Or dChkBlkP = iBlkP Then
                        If dChkBlkP > iBlkP Then
                            rs.MovePrevious
                        End If
                            'WF�̌����𔻒�iCW740�ł͂قڂ��肦�Ȃ�)
                            If rs.Fields("WFSTA") = "4" Then
                                rs.Close
                                DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                                Exit Function
                            End If

                            If IsNull(rs.Fields("RTOP_POS")) = False Then
    ''''                            sBlkP = Int(CDbl(rs.Fields("RTOP_POS")))
                                sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + eps)    'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                            End If
                            If IsNull(rs.Fields("RITOP_POS")) = False Then
    ''''                            sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")))
                                sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + 0.9 + eps) 'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                            End If
                            If IsNull(rs.Fields("BLOCKSEQ")) = False Then
                                sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                            End If
            ''''                    sBlkId = CStr(rs.Fields("LOTID"))
        '                    iTopPos = Int(CInt(rs.Fields("TOP_POS")) + 0.9) '�؂�グ
        '                    sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "B"
                        rs.Close
                        'XXX-XXXT
                        '���݂̃u���b�NID�̎��̃u���b�NID���擾
                        With f_cmbc036_2.sprExamine
                            iBlkflg = 0
                            For i = 1 To .MaxRows
                                .GetText 1, i, vBlkId
                                If vBlkId <> "" And Len(vBlkId) <> 1 Then
                                    If iBlkflg = 1 Then
                                        sBlkId = Left(sBlkId, 9) & CStr(vBlkId) '����BLID�擾
                                        Exit For
                                    ElseIf Right(sBlkId, 3) = vBlkId Then
                                        iBlkflg = 1
                                    End If
                                End If
                            Next i
                        End With


                        sql = "select "
                        sql = sql & "LOTID,"                ' �u���b�NID"
                        sql = sql & "MSXLID,"               ' SXLID"
                        sql = sql & "blockseq,"             ' �u���b�N���A��"
                        sql = sql & "WFSTA,"                ' WF���"
                        sql = sql & "RTOP_POS,"             ' �_���u���b�N���ʒu"
                        sql = sql & "RITOP_POS,"            ' �_���������ʒu"
                        sql = sql & "MSMPLEID,"             ' �����ʒu"
                        sql = sql & "SHAFLAG,"              ' �T���v���t���O"
                        sql = sql & "TOP_POS"               ' �u���b�N���ʒu
                        sql = sql & " from TBCMY011 "
                        sql = sql & " where LOTID ='" & sBlkId & "'"
''                      sql = sql & "   AND TO_NUMBER(WFSTA) <= 1"
                        sql = sql & " ORDER BY BLOCKSEQ ASC"

                        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

                        iLoopCnt = 0
                        rs.MoveFirst
                        'WF�̌����𔻒�iCW740�ł͂قڂ��肦�Ȃ�)
                        If rs.Fields("WFSTA") = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If

                        If IsNull(rs.Fields("RTOP_POS")) = False Then
                            If DBDRV_WFLENGET(sBlkId, dblWFLen) = FUNCTION_RETURN_SUCCESS Then
                                iNextBlkP = Int(CDbl(rs.Fields("RTOP_POS")) - dblWFLen + eps)   'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                            Else
                                iNextBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + eps)  'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                            End If
                        End If
        '                        If IsNull(rs.Fields("RITOP_POS")) = False Then
        '                            sNextIngotP = rs.Fields("RITOP_POS")
        '                        End If

                         sBlkSeq2 = CStr(rs.Fields("BLOCKSEQ"))

        '                rs.MoveFirst
        '                If IsNull(rs.Fields("RTOP_POS")) = False Then
        '                    sBlkP = rs.Fields("RTOP_POS")
        '                End If
        ''                If IsNull(rs.Fields("RITOP_POS")) = False Then
        ''                    sKessyoP = rs.Fields("RITOP_POS")
        ''                End If
        '                sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
        ''''                    sBlkId = CStr(rs.Fields("LOTID"))
                        iTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + eps)  '�؂�̂� 'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                        sAddSmpId2 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "T"
                        Exit Do
                    End If

                 Else   '���̃u���b�N��T(�ȑO�̂܂܂̃��W�b�N�j

                    rs.MoveFirst
                    'WF�̌����𔻒�iCW740�ł͂قڂ��肦�Ȃ�)
                    If rs.Fields("WFSTA") = "4" Then
                        rs.Close
                        DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                        Exit Function
                    End If

                    If IsNull(rs.Fields("RTOP_POS")) = False Then
                        sBlkP = rs.Fields("RTOP_POS")
                    End If
                    If IsNull(rs.Fields("RITOP_POS")) = False Then
                        sKessyoP = rs.Fields("RITOP_POS")
                    End If
                    sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
    ''''                    sBlkId = CStr(rs.Fields("LOTID"))
                    iTopPos = Int(CInt(rs.Fields("TOP_POS")) / 10 + eps) '�؂�̂�  'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                    sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "T"
                    Exit Do
                    End If
            Case "U"
                If IsNull(rs.Fields("RTOP_POS")) = False Then
                    dChkBlkP = CDbl(rs.Fields("RTOP_POS"))
                End If
'===ADD okazaki 2003/04/18
'�ʃu���b�N���͂��ރT���v���ւ̑Ή�
                If dChkBlkP > CInt(sBlkP) Or dChkBlkP = CInt(sBlkP) Then
                    If dChkBlkP > CInt(sBlkP) Then
                        rs.MovePrevious

                        If IsNull(rs.Fields("BLOCKSEQ")) = True Then    'add 2003/04/28 hitec)matsumoto  NULL�̏ꍇ�i�Y��WF�����j�́A���Ɍ�������
                            Do
                                rs.MoveNext
                                If IsNull(rs.Fields("RTOP_POS")) = False Then
                                    Exit Do
                                End If
                            Loop
                        End If
                        'WF�̌����𔻒�iCW740�ł͂قڂ��肦�Ȃ�)
                        If rs.Fields("WFSTA") = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If
''''                        sBlkP = Int(CDbl(rs.Fields("RTOP_POS")))
''''                        sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")))
                        sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + eps)    'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                        sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + 0.9 + eps)   'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                        sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
    ''''                        sBlkId = CStr(rs.Fields("LOTID"))
                        iTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + 0.9 + eps) '�؂�グ    'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                        sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "U"
                        rs.MoveNext
    '                    If sSmpId2 <> vbNullString Then 'D�̃T���v�����쐬
    '                        iTopPos = Int(CInt(rs.Fields("TOP_POS")) / 10)  '�؂�̂�
    '                        sAddSmpId2 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "D"
    '                    End If
    '                    sBlkSeq2 = CStr(rs.Fields("BLOCKSEQ"))
    '                    Exit Do
                    ElseIf dChkBlkP = CInt(sBlkP) Then
''''                    'WF�̌����𔻒�iCW740�ł͂قڂ��肦�Ȃ�)
                        If rs.Fields("WFSTA") = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If
                        sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + eps)  'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
''''                        sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")))
                        sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + eps)    'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                        sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + 0.9 + eps)    'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                        sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
    ''''                        sBlkId = CStr(rs.Fields("LOTID"))
                        iTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + 0.9 + eps) '�؂�グ    'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                        sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "U"
                        rs.MoveNext
                    End If

                    If Not rs.EOF Then
                        'WF�̌����𔻒�iCW740�ł͂قڂ��肦�Ȃ�)
                        If rs.Fields("WFSTA") = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If
                        If sSmpId2 <> vbNullString Then 'D�̃T���v�����쐬
                            '0�ȊO��0.1mm�����Đ؎̂�(WF����:D�͊Y���ʒu���܂܂��ɉ����������) 08/11/06 ooba
                            If rs.Fields("TOP_POS") > 0 Then
                                iTopPos = Int(CDbl(rs.Fields("TOP_POS") - 1) / 10 + eps)    '0.1mm�����Đ؎̂�
                            Else
                                iTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + eps)  '�؂�̂� 'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                            End If
                            sAddSmpId2 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "D"
                        End If
                        sBlkSeq2 = CStr(rs.Fields("BLOCKSEQ"))
                        Exit Do
                    Else
                        '���݂̃u���b�NID�̎��̃u���b�NID���擾
                        With f_cmbc036_2.sprExamine
                            iBlkflg = 0
                            For i = 1 To .MaxRows
                                .GetText 1, i, vBlkId
                                If vBlkId <> "" And Len(vBlkId) <> 1 Then
                                    If iBlkflg = 1 Then
                                        sBlkId = Left(sBlkId, 9) & CStr(vBlkId) '����BLID�擾
                                        Exit For
                                    ElseIf Right(sBlkId, 3) = vBlkId Then
                                        iBlkflg = 1
                                    End If
                                End If
                            Next i
                        End With
                        rs.Close

                        sql = "select "
                        sql = sql & "LOTID,"                ' �u���b�NID"
                        sql = sql & "MSXLID,"               ' SXLID"
                        sql = sql & "blockseq,"             ' �u���b�N���A��"
                        sql = sql & "WFSTA,"                ' WF���"
                        sql = sql & "RTOP_POS,"             ' �_���u���b�N���ʒu"
                        sql = sql & "RITOP_POS,"            ' �_���������ʒu"
                        sql = sql & "MSMPLEID,"             ' �����ʒu"
                        sql = sql & "SHAFLAG,"              ' �T���v���t���O"
                        sql = sql & "TOP_POS"               ' �u���b�N���ʒu
                        sql = sql & " from TBCMY011 "
                        sql = sql & " where LOTID ='" & sBlkId & "'"
'                       sql = sql & "   AND TO_NUMBER(WFSTA) <= 1"
                        sql = sql & " ORDER BY BLOCKSEQ ASC"

                        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

                        iLoopCnt = 0
                        rs.MoveFirst
                        'WF�̌����𔻒�iCW740�ł͂قڂ��肦�Ȃ�)
                        If rs.Fields("WFSTA") = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If
                        If IsNull(rs.Fields("RTOP_POS")) = False Then
                            iNextBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + eps)  'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                        End If
'                        If IsNull(rs.Fields("RITOP_POS")) = False Then
'                            sNextIngotP = rs.Fields("RITOP_POS")
'                        End If
                        If sSmpId2 <> vbNullString Then 'D�̃T���v�����쐬
                            '0�ȊO��0.1mm�����Đ؎̂�(WF����:D�͊Y���ʒu���܂܂��ɉ����������) 08/11/06 ooba
                            If rs.Fields("TOP_POS") > 0 Then
                                iNextBlkP = Int(CDbl(rs.Fields("TOP_POS") - 1) / 10 + eps)  '0.1mm�����Đ؎̂�
                            Else
                                iNextBlkP = Int(CDbl(rs.Fields("TOP_POS")) / 10 + eps)  '�؂�̂�   'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                            End If
                            sAddSmpId2 = Mid(sBlkId, 1, 12) & Format(CStr(iNextBlkP), "000") & "D"
                        End If
                        sBlkSeq2 = CStr(rs.Fields("BLOCKSEQ"))
                        Exit Do
                    End If
                End If

            Case "D"
                If IsNull(rs.Fields("RTOP_POS")) = False Then
                    dChkBlkP = CDbl(rs.Fields("RTOP_POS"))
                End If
''''                    If iChkBlkP < iBlkP Then
''''                        sBlkP = Int(CDbl(rs.Fields("RTOP_POS")))
''''                        sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")))
''''                        sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
''''                        sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
''''                        sBlkId = CStr(rs.Fields("LOTID"))
''''                        iTopPos = Int(CInt(rs.Fields("TOP_POS")) / 10)  '�؂�̂�
''''                        sSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & sSmpId1
''''                        rs.MoveNext
''''                        sBlkSeq2 = CStr(rs.Fields("BLOCKSEQ"))
''''                        Exit Do
''''                    End If
                If dChkBlkP > iBlkP Then
'                    sBlkP = Int(CDbl(rs.Fields("RTOP_POS")))
'                    sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")))
                    'WF�̌����𔻒�iCW740�ł͂قڂ��肦�Ȃ�)
                    If rs.Fields("WFSTA") = "4" Then
                        rs.Close
                        DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                        Exit Function
                    End If
                    sBlkSeq2 = CStr(rs.Fields("BLOCKSEQ"))
                    sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + eps)  'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
 ''''                        sBlkId = CStr(rs.Fields("LOTID"))
                    '0�ȊO��0.1mm�����Đ؎̂�(WF����:D�͊Y���ʒu���܂܂��ɉ����������) 08/11/06 ooba
                    If rs.Fields("TOP_POS") > 0 Then
                        iTopPos = Int(CDbl(rs.Fields("TOP_POS") - 1) / 10 + eps)    '0.1mm�����Đ؎̂�
                    Else
                        iTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + eps)                        '�؂�̂� 'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                    End If
                    sAddSmpId2 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "D"        'D�Ȃ̂�sAddSmpId2�ɓ����
                    rs.MovePrevious

                    'U�̃T���v���쐬�̏C��(������ۯ��ɑΉ�) 2005/04/21 ffc)tanabe =============================> START
                    If Not rs.BOF Then
                        'WF�̌����𔻒�iCW740�ł͂قڂ��肦�Ȃ�)
                        If rs.Fields("WFSTA") = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If
                        If sSmpId2 <> vbNullString Then 'U�̃T���v�����쐬
                            iTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + 0.9 + eps)              '�؂�グ   'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                            sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "U"    '"U"�Ȃ̂�sAddSmpId1�ɓ����
                        End If
    '                    sBlkP = Int(CDbl(rs.Fields("RTOP_POS")))
    '                    sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")))
                        If IsNull(rs.Fields("RTOP_POS")) = False Then
                            sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + eps)    'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                        End If
                        If IsNull(rs.Fields("BLOCKSEQ")) = False Then
                            sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                        End If
                        Exit Do
                    Else
                         
                         '���݂̃u���b�NID��Rows���擾     ## 2008.02.08
                         With f_cmbc036_2.sprExamine
                            iBlkflg = 0
                            For j = .MaxRows To 1 Step -1
                                .GetText 1, j, vBlkId
                                If vBlkId <> "" And Len(vBlkId) <> 1 Then
                                    If Right(sBlkId, 3) = vBlkId Then
                                        Exit For
                                    End If
                                End If
                            Next j
                        End With
                    
                        '���݂̃u���b�NID�̑O�̃u���b�NID���擾
                        With f_cmbc036_2.sprExamine
                            iBlkflg = 0
                            'For i = .MaxRows To 1 Step -1    '## 2008.02.08
                            For i = j To 1 Step -1
                                .GetText 1, i, vBlkId
                                If vBlkId <> "" And Len(vBlkId) <> 1 Then
                                    If Right(sBlkId, 3) <> vBlkId Then
                                        sBlkId = Left(sBlkId, 9) & CStr(vBlkId) '�O��BLID�擾
                                        Exit For
                                    End If
                                End If
                            Next i
                        End With
                        rs.Close

                        sql = "select "
                        sql = sql & "LOTID,"                ' �u���b�NID"
                        sql = sql & "MSXLID,"               ' SXLID"
                        sql = sql & "blockseq,"             ' �u���b�N���A��"
                        sql = sql & "WFSTA,"                ' WF���"
                        sql = sql & "RTOP_POS,"             ' �_���u���b�N���ʒu"
                        sql = sql & "RITOP_POS,"            ' �_���������ʒu"
                        sql = sql & "MSMPLEID,"             ' �����ʒu"
                        sql = sql & "SHAFLAG,"              ' �T���v���t���O"
                        sql = sql & "TOP_POS"               ' �u���b�N���ʒu
                        sql = sql & " from TBCMY011 "
                        sql = sql & " where LOTID ='" & sBlkId & "'"
                        sql = sql & " ORDER BY BLOCKSEQ DESC"

                        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

                        iLoopCnt = 0
                        rs.MoveFirst
                        'WF�̌����𔻒�iCW740�ł͂قڂ��肦�Ȃ�)
                        If rs.Fields("WFSTA") = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If
                        If sSmpId2 <> vbNullString Then 'U�̃T���v�����쐬
                            iTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + 0.9 + eps)              '�؂�グ
                            sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "U"
                        End If
                        If IsNull(rs.Fields("RTOP_POS")) = False Then
                            sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + eps)
                        End If
                        If IsNull(rs.Fields("BLOCKSEQ")) = False Then
                            sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                        End If
                        Exit Do
                    End If
                    'U�̃T���v���쐬�̏C��(������ۯ��ɑΉ�) 2005/04/21 ffc)tanabe =============================> END
                End If
            Case "B"
'                rs.MoveLast
                    If IsNull(rs.Fields("RTOP_POS")) = False Then
                        dChkBlkP = CDbl(rs.Fields("RTOP_POS"))
                    End If
                    If dChkBlkP > iBlkP Or dChkBlkP = iBlkP Then
                        If dChkBlkP > iBlkP Then
                            rs.MovePrevious
                        End If
                        'WF�̌����𔻒�iCW740�ł͂قڂ��肦�Ȃ�)
                        If rs.Fields("WFSTA") = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If

                        If IsNull(rs.Fields("RTOP_POS")) = False Then
''''                            sBlkP = Int(CDbl(rs.Fields("RTOP_POS")))
                            sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + eps)    'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                        End If
                        If IsNull(rs.Fields("RITOP_POS")) = False Then
''''                            sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")))
                            sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + 0.9 + eps)    'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                        End If
                        sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                        sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
        ''''                    sBlkId = CStr(rs.Fields("LOTID"))
                        iTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + 0.9 + eps) '�؂�グ    'add 2003/06/13 hitec)matsumoto [+ eps]�ǉ�
                        sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "B"
                        Exit Do
                    End If

        End Select
        rs.MoveNext
    Loop
    sSmpId1 = sAddSmpId1
    sSmpId2 = sAddSmpId2
    rs.Close

    DBDRV_GET_WFMAP = FUNCTION_RETURN_SUCCESS

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




'�T�v    :WF�}�b�v�e�[�u���X�V
'���Ұ�  :�ϐ���       ,IO   ,�^                                    ,����
'        :��ؒl        ,O    ,FUNCTION_RETURN                       ,�ǂݍ��ݐ���
'����    :WF�}�b�v�e�[�u��(TBCMY011)���X�V����
'����    :2003/3/25 Hitec)matsumoto
Public Function DBDRV_UPD_WFMap() As FUNCTION_RETURN

    Dim sql             As String
    Dim rs              As OraDynaset
    Dim i               As Long
    Dim iLoopCnt        As Long
    Dim sDbName         As String
    Dim itUCount        As Integer
''''Dim nowtime         As Date
    Dim vGetMaxSeq      As Variant
    Dim sGetSxlId       As String
    Dim vGetSXLID1      As Variant
    Dim vGetSXLID2      As Variant
    Dim NowIngotPos     As Integer
    Dim iGetSmplLoop    As Integer
    Dim iFromBlkSeq     As Integer
    Dim iToBlkSeq       As Integer
    Dim iNextLoopCnt    As Integer
    Dim vGetSample      As Variant
    Dim iBlkflg         As Integer

    Dim sLotid          As String
    Dim iFromIngotPos   As Integer
    Dim iToIngotPos     As Integer
    Dim vGetHinban      As Variant
    Dim m               As Integer
    Dim k               As Integer      '2003/05/18 add
    Dim sOldSXLID       As String       '2003/05/18 add
    Dim sOldIngotP      As String       '2003/05/18 add
    Dim vGetBlockSEQ_S  As Variant      '2003/05/29
    Dim vGetBlockSEQ_E  As Variant      '2003/05/29

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_UPD_WFMap"

    sDbName = "(Y011)"
'�i�Ԃ�1��ǉ��������Ƃɂ���̕ύX-------start iida 2003/09/03
    '2003/05/01
    With f_cmbc036_2.sprExamine
        m = .MaxRows
        'SXL�̍X�V
        For iLoopCnt = 1 To m Step 2

            '�T���v���s��SXL�𔻒肷��
            .row = iLoopCnt
            .col = 10
            If Len(Trim(.Text)) > 0 Then     '�T���v���s�̏ꍇ
                .GetText 5, iLoopCnt, gtSprWfMap(iLoopCnt).KESSYOUP 'add 2003/05/17 hitec)matsumoto �����ʒu����ʂ���擾
                sGetSxlId = Mid(gtSprWfMap(iLoopCnt).LOTID, 1, 10) & GetWafPos(CInt(gtSprWfMap(iLoopCnt).KESSYOUP))
                If iLoopCnt = 1 Then    '�擪
                    sGetSxlId = Mid(gtSprWfMap(iLoopCnt).LOTID, 1, 10) & GetWafPos(CInt(SIngotP))
                End If

                '#######2003/05/18 okazaki
                If Get_OLDSXLID(CInt(gtSprWfMap(iLoopCnt).KESSYOUP), sOldSXLID, sOldIngotP) = FUNCTION_RETURN_SUCCESS Then

                    sGetSxlId = sOldSXLID
                    If iLoopCnt = 1 Then
                        gtSprWfMap(iLoopCnt).KESSYOUP = SIngotP   ' 2003/05/18 okazaki
                    Else
                        gtSprWfMap(iLoopCnt).KESSYOUP = sOldIngotP  '�\��
                        For k = 0 To UBound(tmpSXLMng)
                            If sGetSxlId = tmpSXLMng(k).SXLID Then
                                gtSprWfMap(iLoopCnt).KESSYOUP = tmpSXLMng(k).INGOTPOS
                                Exit For
                            End If
                        Next k
                    End If
                End If
            End If


            .GetText 2, iLoopCnt, vGetHinban
            If vGetHinban <> "Z" Then
                '2003/05/29 hitec)okazaki �u���b�NSEQ����ʂ���擾�ɕύX
                .GetText 6, iLoopCnt, vGetBlockSEQ_S
                .GetText 6, iLoopCnt + 1, vGetBlockSEQ_E
                iFromBlkSeq = CInt(vGetBlockSEQ_S)                                  '�u���b�NSEQ���擾
                iToBlkSeq = CInt(vGetBlockSEQ_E)                                    '�u���b�NSEQ���擾
                '2003/05/29 end

                sql = "UPDATE TBCMY011 SET"
                sql = sql & " mhinban = '" & gtSprWfMap(iLoopCnt).hinban & "'"      ' �i��"
'''''                nowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")

                If iLoopCnt = 1 Then
                    NowIngotPos = SIngotP
                Else
                    NowIngotPos = gtSprWfMap(iLoopCnt).KESSYOUP
                End If

                sql = sql & ",MSXLID = '" & sGetSxlId & "'"
                sql = sql & ",UPDPROC= 'CW740'"                                     ' �X�V�H��
                sql = sql & ",UPDDATE=  sysdate"                                    ' �X�V����

                '���i����tblWafInd����擾���� 2003/05/28 okazaki start
                For i = 1 To UBound(tblWafInd)
                    If gtSprWfMap(iLoopCnt).hinban = tblWafInd(i).HINDN.hinban Then
                        Exit For
                    End If
                Next i
                sql = sql & ",MREVNUM =  " & tblWafInd(i).HINDN.mnorevno            ' ���i�ԍ������ԍ�
                sql = sql & ",MFACTORY= '" & tblWafInd(i).HINDN.factory & "'"       ' �H��
                sql = sql & ",MOPECOND= '" & tblWafInd(i).HINDN.opecond & "'"       ' ���Ə���

                '2003/05/28 end
                sql = sql & " WHERE LOTID ='" & gtSprWfMap(iLoopCnt).LOTID & "'"    ' �u���b�NID"
                If (iFromBlkSeq <= iToBlkSeq) Then
                    sql = sql & " AND ((BLOCKSEQ >= " & iFromBlkSeq & ")"           ' �u���b�N���A��"
                    sql = sql & " AND  (BLOCKSEQ <= " & iToBlkSeq & "  ))"
                Else
                    sql = sql & " AND  (BLOCKSEQ >= " & iFromBlkSeq & ")"           ' �u���b�N���A��"
                End If

                '' WriteDBLog sql
                Debug.Print sql
                If 0 >= OraDB.ExecuteSQL(sql) Then
                    DBDRV_UPD_WFMap = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If


                '�͈͓��̌��� (WFSTA=�S)�ȊO�̃T���v���������ׂăN���A
                sql = "UPDATE TBCMY011 SET"
                sql = sql & " SHAFLAG = '0'"                                        ' �T���v���t���O"
                sql = sql & ",WFSTA   = '0'"                                        ' WF���
                sql = sql & ",MSMPLEID= NULL"                                       ' �����ʒu"
                sql = sql & ",UPDDATE = sysdate"                                    ' �X�V����

                sql = sql & " WHERE LOTID ='" & gtSprWfMap(iLoopCnt).LOTID & "'"    ' �u���b�NID"
                If (iFromBlkSeq <= iToBlkSeq) Then
                    sql = sql & " AND ((BLOCKSEQ >= " & iFromBlkSeq & ")"           ' �u���b�N���A��"
                    sql = sql & " AND  (BLOCKSEQ <= " & iToBlkSeq & "  ))"
                Else
                    sql = sql & " AND  (BLOCKSEQ >= " & iFromBlkSeq & ")"           ' �u���b�N���A��"
                End If
                sql = sql & " AND WFSTA <> '4'"

                '' WriteDBLog sql
                Debug.Print sql
                If 0 >= OraDB.ExecuteSQL(sql) Then
                        '�X�V�Y����0�̏ꍇ�����s
                End If
            End If
        Next iLoopCnt


        '�T���v���̍X�V�i��ʂ���u���L�v�𔻒肷��j
        For iLoopCnt = 1 To UBound(gtSprWfMap())
            .GetText 10, iLoopCnt, vGetSample
            If (vGetSample <> vbNullString) Then

                sql = "UPDATE TBCMY011 SET"
                If vGetSample = gsWF_SMPL_JOINT Then
'                    .GetText 30, iLoopCnt, vGetSample
                    ''�c���_�f�������ڒǉ��ɂ��ύX�@03/12/09 ooba
'                    .GetText 31, iLoopCnt, vGetSample
                    'GD�ǉ��ɂ��ύX�@05/01/31 ooba
'                    .GetText 32, iLoopCnt, vGetSample
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
                    .GetText 38, iLoopCnt, vGetSample
                    Call Cnv_GetSample(vGetSample)

                    sql = sql & " MSMPLEID= '" & vGetSample & "'"                   ' �����ʒu"
                    sql = sql & ",SHAFLAG = '1'"                                    ' �T���v���t���O"
                    sql = sql & ",WFSTA   = '1'"                                    ' WF��ԃT���v��  'del 2003/05/03 hitec)matsumoto
                Else
'                    .GetText 30, iLoopCnt, vGetSample                               ' 03/05/28
                    ''�c���_�f�������ڒǉ��ɂ��ύX�@03/12/09 ooba
'                    .GetText 31, iLoopCnt, vGetSample
                    'GD�ǉ��ɂ��ύX�@05/01/31 ooba
'                    .GetText 32, iLoopCnt, vGetSample
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
                    .GetText 38, iLoopCnt, vGetSample
                    Call Cnv_GetSample(vGetSample)

                    sql = sql & " MSMPLEID= '" & vGetSample & "'"                   ' �����ʒu"  03/05/28
                    sql = sql & ",SHAFLAG = '1'"                                    ' �T���v���t���O"
                    sql = sql & ",WFSTA   = '0'"                                    ' WF��ԃT���v��  'del 2003/05/03 hitec)matsumoto
                End If

'''''                nowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")

                sql = sql & ",UPDPROC = 'CW740'"                                    ' �X�V�H��
                sql = sql & ",UPDDATE = sysdate"                                    ' �X�V����
                sql = sql & " WHERE LOTID   = '" & gtSprWfMap(iLoopCnt).LOTID & "'" ' �u���b�NID"
                sql = sql & "   AND BLOCKSEQ=  " & gtSprWfMap(iLoopCnt).BLOCKSEQ    ' �u���b�N���A��"
                sql = sql & "   AND WFSTA   <>'4'"
                '' WriteDBLog sql
                Debug.Print sql
                If 0 >= OraDB.ExecuteSQL(sql) Then

                End If

            End If
        Next
    End With
    '�i�Ԃ�1��ǉ��������Ƃɂ���̕ύX-------end iida 2003/09/03
    DBDRV_UPD_WFMap = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_UPD_WFMap = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function


Public Sub Cnv_GetSample(ByRef vGetSample As Variant)
    Dim i   As Integer
    Dim kbn As String

    For i = 1 To UBound(CngSmpID_UD)
        If CngSmpID_UD(i) = vGetSample Then
           kbn = Cnv_Smp_KB(Right(vGetSample, 1))
           vGetSample = Left(vGetSample, Len(vGetSample) - 1) + kbn
           Exit Sub
        End If
    Next
End Sub

Public Function Cnv_Smp_KB(SmpKb As String) As String
    If SmpKb = "U" Then
        Cnv_Smp_KB = "B"
        Exit Function
    End If

    If SmpKb = "D" Then
        Cnv_Smp_KB = "T"
        Exit Function
    End If
End Function


'�T�v    :�Y���u���b�N��WF�P���̒����i�v�Z���j���擾
'���Ұ�  :�ϐ���       ,IO   ,�^                    ,����
'        :BLOCKID       ,I   ,STRING                ,�u���b�N�h�c
'        :dblWFLen      ,O   ,DOUBLE        �@�@    ,WF1���̌v�Z����
'        :��ؒl         ,O   ,FUNCTION_RETURN       ,�ǂݍ��ݐ���
'����    :�Y���u���b�N��WF�P���̒����i�v�Z���j���擾
'����    :2003/4/25 Hitec)okazaki
Public Function DBDRV_WFLENGET(ByVal StrBlockId As String, ByRef dblWFLen As Double) As FUNCTION_RETURN

    Dim strSQL      As String
    Dim iRealLen    As Integer
    Dim iWFcnt      As Integer
    Dim rs          As OraDynaset
    Dim iKetuFrom   As Integer
    Dim iKetuTo     As Integer
    Dim iKetuLen    As Integer
    Dim sDbName     As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function DBDRV_WFLENGET"

    '�������AWF�����擾
    sDbName = "(Y011)"

    strSQL = "select e40.blockid,e40.reallen,y11.cnt"
    strSQL = strSQL & " from tbcme040 e40,"
    strSQL = strSQL & " xsdca xa,"
    strSQL = strSQL & " (select lotid,count(lotid) cnt"
    strSQL = strSQL & "  from   tbcmy011"
    strSQL = strSQL & "  where  lotid ='" & StrBlockId & "'"
    strSQL = strSQL & "  group by lotid  ) y11"
    strSQL = strSQL & " where e40.blockid =  xa.CRYNUMCA"
    strSQL = strSQL & "   and y11.lotid   =  xa.CRYNUMCA"
    strSQL = strSQL & "   and y11.lotid   = '" & StrBlockId & "'"

    Set rs = OraDB.DBCreateDynaset(strSQL, ORADYN_NO_BLANKSTRIP)
    If Not rs.EOF Then
           iRealLen = CInt(rs!REALLEN)
           iWFcnt = CInt(rs!cnt)
    Else
        rs.Close
        DBDRV_WFLENGET = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    rs.Close

    '���������擾
    sDbName = "(Y012)"
    strSQL = "SELECT DISTINCT LENFROM,LENTO FROM TBCMY012"
    strSQL = strSQL & " Where "
    strSQL = strSQL & " LOTID   = '" & StrBlockId & "'"

    Set rs = OraDB.DBCreateDynaset(strSQL, ORADYN_NO_BLANKSTRIP)
    iKetuLen = 0
    Do While Not rs.EOF
        If (IsNull(rs.Fields("LENFROM")) = True) Or rs.Fields("LENFROM") = -1 Or _
            (IsNull(rs.Fields("LENTO")) = True) Or rs.Fields("LENTO") = -1 Then
        Else
            iKetuFrom = CInt(rs.Fields("LENFROM"))
            iKetuTo = CInt(rs.Fields("LENTO"))
            iKetuLen = iKetuLen + iKetuTo - iKetuFrom
        End If
        rs.MoveNext
    Loop
    rs.Close

    'WF�����v�Z
    dblWFLen = (iRealLen - iKetuLen) / iWFcnt

    '�����_�Q���ڂ��l�̌ܓ�
''''    dblWFLen = Int((dblWFLen + 0.05) * 10) / 10 'del 2003/08/06 hitec)matusmoto

    DBDRV_WFLENGET = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print strSQL
    DBDRV_WFLENGET = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit


End Function

'�T�v      :SXL�h�c�̎擾&����
'���Ұ��@�@:�ϐ����@�@�@�@,IO ,�^       ,����
'�@�@      :iIngotPos     ,I  ,Integer�@,�����ʒu
'          :sSXLID        ,O  ,STRING   ,SXLID
'�@�@      :�߂�l�@�@�@�@,O  ,�@       ,�I���̗L��
'����      :��ʂ��AA����������SXL�擪�̕i�Ԃɑ΂�����SXLID���擾
'����      :2003/05/01   hitec)okazaki

Public Function Get_OLDSXLID(iIngotpos As Integer, sSXLID As String, sOldIngotP As String) As FUNCTION_RETURN

    Dim i           As Integer
    Dim iRowIngotP  As Integer
    Dim vGetIngotP  As Variant
    Dim vGetHinban  As Variant
    Dim vGetSXLID1  As Variant
    Dim vGetSXLID2  As Variant
    Dim vWFcnt      As Variant
    Dim iSumWFcnt   As Integer
    Dim j           As Integer
    Dim vGetSampl   As Variant
    Dim vGetSampl2  As Variant

    Dim idx2        As Integer

    Get_OLDSXLID = FUNCTION_RETURN_FAILURE

    '�i�Ԃ�1��ǉ��������Ƃɗ�̕ύX-------start iida 2003/09/03
    With f_cmbc036_2.sprExamine

        For i = 1 To .MaxRows - 1 Step 2
            If i > .MaxRows Then
                Exit Function
            End If
            .GetText 5, i, vGetIngotP
            If iIngotpos = CInt(vGetIngotP) Then
                iRowIngotP = i      '��ʂ̊Y���ʒu���擾
                Exit For
            End If
        Next i

        idx2 = Get_YukouRow(iRowIngotP, "D")
        iRowIngotP = idx2

'        .GetText 36, iRowIngotP, vGetSXLID1                 '��ʂ̊Y���ʒu�̌�SXLID���擾
        ''�c���_�f�������ڒǉ��ɂ��ύX�@03/12/09 ooba
'        .GetText 37, iRowIngotP, vGetSXLID1
        'GD�ǉ��ɂ��ύX�@05/01/31 ooba
'        .GetText 38, iRowIngotP, vGetSXLID1
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
        .GetText 44, iRowIngotP, vGetSXLID1
        iSumWFcnt = 0
        For i = 2 To .MaxRows Step 2
            If iRowIngotP - i < 1 Then                      '���̕i�Ԃ���SXL�̐擪�̏ꍇ�iA���������j
                .GetText 5, 1, vGetIngotP
                sOldIngotP = CStr(vGetIngotP)
                sSXLID = CStr(vGetSXLID1)

                Get_OLDSXLID = FUNCTION_RETURN_SUCCESS
                Exit For
            End If
            .GetText 2, iRowIngotP - i, vGetHinban
            If vGetHinban <> "Z" Then                       'A�����łȂ��ŏ��̑O�̕i��
'                .GetText 36, iRowIngotP - i, vGetSXLID2
                ''�c���_�f�������ڒǉ��ɂ��ύX�@03/12/09 ooba
'                .GetText 37, iRowIngotP - i, vGetSXLID2
                'GD�ǉ��ɂ��ύX�@05/01/31 ooba
'                .GetText 38, iRowIngotP - i, vGetSXLID2
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
                .GetText 44, iRowIngotP - i, vGetSXLID2
                If vGetSXLID2 <> vGetSXLID1 Then            '���̕i�Ԃ���SXL�̐擪�̏ꍇ�iA���������j
                    .GetText 5, iRowIngotP - i + 2, vGetIngotP
                    sOldIngotP = CStr(vGetIngotP)
                    sSXLID = CStr(vGetSXLID1)
                    Get_OLDSXLID = FUNCTION_RETURN_SUCCESS
                End If

                Exit For
            End If
        Next i

    End With
    '�i�Ԃ�1��ǉ��������Ƃɗ�̕ύX-------end iida 2003/09/03

End Function

'###############################################2003/05/19 okazaki

'�T�v      :A�������΂����L���ȍs���擾����(�u���b�N�͍l���Ȃ��j
'���Ұ��@�@:�ϐ����@�@�@�@,IO ,�^       ,����
'�@�@      :iNowRow      ,I  ,Integer�@,�����ʒu
'          :sUD          ,I  ,STRING   ,�����i�ォ�����j"U":��@"D":��
'�@�@      :�߂�l�@�@�@�@,O  ,INTEGER�@,�L���s
'����      :��ʂ��AA�����������i�Ԃ̂���s�ԍ�(Spread)���擾
'����      :2003/05/19   hitec)okazaki

Public Function Get_YukouRow(iNowRow As Integer, ByRef sUD As String) As Integer

    Dim vGetHinban  As Variant
    Dim iCount      As Integer

    On Error Resume Next


    Get_YukouRow = iNowRow
    With f_cmbc036_2.sprExamine
        '�p�����[�^�`�F�b�N
        If iNowRow < 1 Or iNowRow > .MaxRows Then
            Exit Function
        End If

        If sUD <> "U" And sUD <> "D" Then
            Exit Function
        End If


        '��Ɍ���
        If sUD = "U" Then
            For iCount = iNowRow To 1 Step -1
                If iCount Mod 2 = 1 Then
                    .GetText 2, iCount, vGetHinban
                    If vGetHinban <> "Z" Then
                        Get_YukouRow = iCount
                        Exit For
                    End If
                End If
            Next iCount

        '���Ɍ���
        ElseIf sUD = "D" Then
            For iCount = iNowRow To .MaxRows - 1 Step 1
                If iCount Mod 2 = 1 Then
                    .GetText 2, iCount, vGetHinban
                    If vGetHinban <> "Z" Then
                        Get_YukouRow = iCount
                        Exit For
                    End If
                End If
            Next iCount
        End If
    End With

End Function
'###############################################2003/05/19 okazaki
'---------------------------------------------------------------
'
' �@�\�@�@  : SXL�`�F�b�N�{�b�N�X�ڍׂ̕\��
'
' �Ԃ�l�@  : �Ȃ�
'
' �����@    : iIndex�@�P�F�\�� / �O�F��\��
'
'
' �@�\����  : SXL�`�F�b�N�{�b�N�X�\���̎��A�ڍׂ�\������
'
' ���l�@�@  : 03/05/31  �㓡
'
'---------------------------------------------------------------
Public Sub Pic_Disp(iIndex As Integer)
    Dim iCnt    As Integer

    With f_cmbc036_2
        If iIndex = 0 Then
            For iCnt = 0 To 2
                .lbl_check(iCnt).Visible = False
            Next
            .pic_check(0).Visible = False
            .pic_check(1).Visible = False

        ElseIf iIndex = 1 Then
            For iCnt = 0 To 2
                .lbl_check(iCnt).Visible = True
            Next
            .pic_check(0).Visible = True
            .pic_check(1).Visible = True
        End If
    End With
End Sub


'�T�v      :WF�T���v���Ǘ��̑}��
'���Ұ��@�@:�ϐ���      ,IO ,�^                 ,����
'      �@�@:WFSMP �@�@�@,I  ,typ_XSDCW   �@     ,�V�T���v���Ǘ��iSXL�j
'      �@�@:�߂�l      ,O  ,FUNCTION_RETURN�@  ,�������݂̐���
'����      :DBDRV_WfSmp_UpdIns�Ɉڍs����\��
'����      :2001/07/12  �쐬 ���{
Public Function DBDRV_WfSmp_INS(WFSMP() As typ_XSDCW, i As Long) As FUNCTION_RETURN

    Dim sql As String
'    Dim i As Long '2003/09/22�R�����g�ɂ���
    Dim sDbName As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_WfSmp_INS"

    DBDRV_WfSmp_INS = FUNCTION_RETURN_SUCCESS

    sDbName = "XSDCW"
'    For i = 1 To UBound(WFSMP)�@'2003/09/22 �R�����g�ɂ���
        With WFSMP(i)

                sql = "insert into XSDCW ("
                sql = sql & "SXLIDCW, "         ' SXLID
                sql = sql & "SMPKBNCW, "        ' �T���v���敪
                sql = sql & "TBKBNCW, "         ' T/B�敪
                sql = sql & "REPSMPLIDCW, "     ' �T���v��ID
                sql = sql & "XTALCW, "          ' �����ԍ�
                sql = sql & "INPOSCW, "         ' �������ʒu
                sql = sql & "HINBCW, "          ' �i��
                sql = sql & "REVNUMCW, "        ' ���i�ԍ������ԍ�
                sql = sql & "FACTORYCW, "       ' �H��
                sql = sql & "OPECW, "           ' ���Ə���
                sql = sql & "KTKBNCW, "         ' �m��敪
                sql = sql & "SMCRYNUMCW, "      ' �T���v���u���b�NID
                sql = sql & "WFSMPLIDRSCW, "    ' �T���v��ID(Rs)
                sql = sql & "WFSMPLIDRS1CW, "   ' ����T���v��ID1�iRs�j
                sql = sql & "WFSMPLIDRS2CW, "   ' ����T���v��ID2�iRs�j
                sql = sql & "WFINDRSCW, "       ' ���FLG�iRs)
                sql = sql & "WFRESRS1CW, "      ' ����FLG1�iRs)
                sql = sql & "WFRESRS2CW, "      ' ����FLG2�iRs)
                sql = sql & "WFSMPLIDOICW, "    ' �T���v��ID�iOi�j
                sql = sql & "WFINDOICW, "       ' ���FLG�iOi)
                sql = sql & "WFRESOICW, "       ' ����FLG�iOi)
                sql = sql & "WFSMPLIDB1CW, "    ' �T���v��ID�iB1�j
                sql = sql & "WFINDB1CW, "       ' ���FLG�iB1)
                sql = sql & "WFRESB1CW, "       ' ����FLG�iB1)
                sql = sql & "WFSMPLIDB2CW, "    ' �T���v��ID�iB2�j
                sql = sql & "WFINDB2CW, "       ' ���FLG�iB2)
                sql = sql & "WFRESB2CW, "       ' ����FLG�iB2)
                sql = sql & "WFSMPLIDB3CW, "    ' �T���v��ID�iB3�j
                sql = sql & "WFINDB3CW, "       ' ���FLG�iB3)
                sql = sql & "WFRESB3CW, "       ' ����FLG�iB3)
                sql = sql & "WFSMPLIDL1CW, "    ' �T���v��ID�iL1�j
                sql = sql & "WFINDL1CW, "       ' ���FLG�iL1)
                sql = sql & "WFRESL1CW, "       ' ����FLG�iL1)
                sql = sql & "WFSMPLIDL2CW, "    ' �T���v��ID�iL2�j
                sql = sql & "WFINDL2CW, "       ' ���FLG�iL2)
                sql = sql & "WFRESL2CW, "       ' ����FLG�iL2)
                sql = sql & "WFSMPLIDL3CW, "    ' �T���v��ID�iL3�j
                sql = sql & "WFINDL3CW, "       ' ���FLG�iL3)
                sql = sql & "WFRESL3CW, "       ' ����FLG�iL3)
                sql = sql & "WFSMPLIDL4CW, "    ' �T���v��ID�iL4�j
                sql = sql & "WFINDL4CW, "       ' ���FLG�iL4)
                sql = sql & "WFRESL4CW, "       ' ����FLG�iL4)
                sql = sql & "WFSMPLIDDSCW, "    ' �T���v��ID�iDS�j
                sql = sql & "WFINDDSCW, "       ' ���FLG�iDS)
                sql = sql & "WFRESDSCW, "       ' ����FLG�iDS)
                sql = sql & "WFSMPLIDDZCW, "    ' �T���v��ID�iDZ�j
                sql = sql & "WFINDDZCW, "       ' ���FLG�iDZ)
                sql = sql & "WFRESDZCW, "       ' ����FLG�iDZ)
                sql = sql & "WFSMPLIDSPCW, "    ' �T���v��ID�iSP�j
                sql = sql & "WFINDSPCW, "       ' ���FLG�iSP)
                sql = sql & "WFRESSPCW, "       ' ����FLG�iSP)
                sql = sql & "WFSMPLIDDO1CW,"    ' �T���v��ID�iDO1�j
                sql = sql & "WFINDDO1CW, "      ' ���FLG�iDO1)
                sql = sql & "WFRESDO1CW, "      ' ����FLG�iDO1)
                sql = sql & "WFSMPLIDDO2CW, "   ' �T���v��ID�iDO2�j
                sql = sql & "WFINDDO2CW, "      ' ���FLG�iDO2)
                sql = sql & "WFRESDO2CW, "      ' ����FLG�iDO2)
                sql = sql & "WFSMPLIDDO3CW, "   ' �T���v��ID�iDO3�j
                sql = sql & "WFINDDO3CW, "      ' ���FLG�iDO3)
                sql = sql & "WFRESDO3CW, "      ' ����FLG�iDO3)
                sql = sql & "WFSMPLIDOT1CW, "   ' �T���v��ID�iOT1�j
               'add start 2003/05/21 hitec)matsumoto -------------------------
                sql = sql & "WFINDOT1CW, "      ' ���FLG�iOT1)
                sql = sql & "WFRESOT1CW, "      ' ����FLG�iOT1)
                sql = sql & "WFSMPLIDOT2CW, "   ' �T���v��ID�iOT2�j
                sql = sql & "WFINDOT2CW, "      ' ���FLG�iOT2)
                sql = sql & "WFRESOT2CW, "      ' ����FLG�iOT2)
               'add end   2003/05/21 hitec)matsumoto -------------------------
                sql = sql & "WFSMPLIDAOICW, "   ' �T���v��ID�iAOi�j
                sql = sql & "WFINDAOICW, "      ' ���FLG�iAOi�j
                sql = sql & "WFRESAOICW, "      ' ����FLG�iAOi�j
                '' GD�ǉ��@05/01/31 ooba START =====================================>
                sql = sql & "WFSMPLIDGDCW, "    ' �T���v��ID (GD)
                sql = sql & "WFINDGDCW, "       ' ���FLG (GD)
                sql = sql & "WFRESGDCW, "       ' ����FLG (GD)
                sql = sql & "WFHSGDCW, "        ' �ۏ�FLG (GD)
                '' GD�ǉ��@05/01/31 ooba END =======================================>
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                sql = sql & "EPSMPLIDB1CW, "    ' �T���v��ID (BMD1E)
                sql = sql & "EPINDB1CW, "       ' ���FLG (BMD1E)
                sql = sql & "EPRESB1CW, "       ' ����FLG (BMD1E)
                sql = sql & "EPSMPLIDB2CW, "    ' �T���v��ID (BMD2E)
                sql = sql & "EPINDB2CW, "       ' ���FLG (BMD2E)
                sql = sql & "EPRESB2CW, "       ' ����FLG (BMD2E)
                sql = sql & "EPSMPLIDB3CW, "    ' �T���v��ID (BMD3E)
                sql = sql & "EPINDB3CW, "       ' ���FLG (BMD3E)
                sql = sql & "EPRESB3CW, "       ' ����FLG (BMD3E)
                sql = sql & "EPSMPLIDL1CW, "    ' �T���v��ID (OSF1E)
                sql = sql & "EPINDL1CW, "       ' ���FLG (OSF1E)
                sql = sql & "EPRESL1CW, "       ' ����FLG (OSF1E)
                sql = sql & "EPSMPLIDL2CW, "    ' �T���v��ID (OSF2E)
                sql = sql & "EPINDL2CW, "       ' ���FLG (OSF2E)
                sql = sql & "EPRESL2CW, "       ' ����FLG (OSF2E)
                sql = sql & "EPSMPLIDL3CW, "    ' �T���v��ID (OSF3E)
                sql = sql & "EPINDL3CW, "       ' ���FLG (OSF3E)
                sql = sql & "EPRESL3CW, "       ' ����FLG (OSF3E)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
                sql = sql & "SMPLNUMCW, "       ' �T���v������
                sql = sql & "SMPLPATCW, "       ' �T���v���p�^�[��
                sql = sql & "NUKISIFLGCW, "     ' �����w���ʉ߃t���O 09/05/26 ooba
                sql = sql & "TSTAFFCW,"         ' �o�^�Ј�ID
                sql = sql & "TDAYCW, "          ' �o�^���t
                sql = sql & "KSTAFFCW, "        ' �X�V�Ј�ID
                sql = sql & "KDAYCW, "          ' �X�V���t
                sql = sql & "SNDKCW, "          ' ���M�t���O
                sql = sql & "SNDDAYCW, "        ' ���M���t
                sql = sql & "LIVKCW)"           ' �����敪

                sql = sql & " values ('"
                sql = sql & .SXLIDCW & "', '"           ' SXLID
                sql = sql & .SMPKBNCW & "', '"          ' �T���v���敪
                sql = sql & .TBKBNCW & "', '"           ' T/B�敪
                sql = sql & .REPSMPLIDCW & "', '"       ' �T���v��ID
                sql = sql & .XTALCW & "', "             ' �����ԍ�
                sql = sql & .INPOSCW & ", '"            ' �������ʒu
                sql = sql & .HINBCW & "', "             ' �i��
                sql = sql & .REVNUMCW & ", '"           ' ���i�ԍ������ԍ�
                sql = sql & .FACTORYCW & "', '"         ' �H��
                sql = sql & .OPECW & "', '"             ' ���Ə���
                sql = sql & .KTKBNCW & "', '"           ' �m��敪
                sql = sql & .SMCRYNUMCW & "', '"        ' �T���v���u���b�NID
                sql = sql & .WFSMPLIDRSCW & "', "       ' �T���v��ID�iRs�j
'               sql = sql & .WFSMPLIDRS1CW & "',"       ' ����T���v��ID1�iRs�j
'               sql = sql & .WFSMPLIDRS2CW & "', "      ' ����T���v��ID2�iRs�j
                sql = sql & "Null, "                    ' ����T���v��ID1�iRs�j
                sql = sql & "Null, '"                   ' ����T���v��ID2�iRs�j
''              sql = sql & .WFINDRSCW & "', "          ' ���FLG�iRs)
''              sql = sql & "Null, "                    ' ����FLG1�iRs)
                sql = sql & .WFINDRSCW & "', '"         ' ���FLG�iRs)
                sql = sql & .WFRESRS1CW & "', "         ' ����FLG1�iRs)
                sql = sql & "Null, '"                   ' ����FLG2�iRs)
                sql = sql & .WFSMPLIDOICW & "', '"      ' �T���v��ID�iOi�j
                sql = sql & .WFINDOICW & "', '"         ' ���FLG�iOi)
                sql = sql & .WFRESOICW & "', '"         ' ����FLG�iOi)
                sql = sql & .WFSMPLIDB1CW & "', '"      ' �T���v��ID�iB1�j
                sql = sql & .WFINDB1CW & "', '"         ' ���FLG�iB1)
                sql = sql & .WFRESB1CW & "', '"         ' ����FLG�iB1)
                sql = sql & .WFSMPLIDB2CW & "', '"      ' �T���v��ID�iB2�j
                sql = sql & .WFINDB2CW & "', '"         ' ���FLG�iB2)
                sql = sql & .WFRESB2CW & "', '"         ' ����FLG�iB2)
                sql = sql & .WFSMPLIDB3CW & "', '"      ' �T���v��ID�iB3�j
                sql = sql & .WFINDB3CW & "', '"         ' ���FLG�iB3)
                sql = sql & .WFRESB3CW & "', '"         ' ����FLG�iB3)
                sql = sql & .WFSMPLIDL1CW & "', '"      ' �T���v��ID�iL1�j
                sql = sql & .WFINDL1CW & "', '"         ' ���FLG�iL1)
                sql = sql & .WFRESL1CW & "', '"         ' ����FLG�iL1)
                sql = sql & .WFSMPLIDL2CW & "', '"      ' �T���v��ID�iL2�j
                sql = sql & .WFINDL2CW & "', '"         ' ���FLG�iL2)
                sql = sql & .WFRESL2CW & "', '"         ' ����FLG�iL2)
                sql = sql & .WFSMPLIDL3CW & "', '"      ' �T���v��ID�iL3�j
                sql = sql & .WFINDL3CW & "', '"         ' ���FLG�iL3)
                sql = sql & .WFRESL3CW & "', '"         ' ����FLG�iL3)
                sql = sql & .WFSMPLIDL4CW & "', '"      ' �T���v��ID�iL4�j
                sql = sql & .WFINDL4CW & "', '"         ' ���FLG�iL4)
                sql = sql & .WFRESL4CW & "', '"         ' ����FLG�iL4)
                sql = sql & .WFSMPLIDDSCW & "', '"      ' �T���v��ID�iDS�j
                sql = sql & .WFINDDSCW & "', '"         ' ���FLG�iDS)
                sql = sql & .WFRESDSCW & "', '"         ' ����FLG�iDS)
                sql = sql & .WFSMPLIDDZCW & "', '"      ' �T���v��ID�iDZ�j
                sql = sql & .WFINDDZCW & "', '"         ' ���FLG�iDZ)
                sql = sql & .WFRESDZCW & "', '"         ' ����FLG�iDZ)
                sql = sql & .WFSMPLIDSPCW & "', '"      ' �T���v��ID�iSP�j
                sql = sql & .WFINDSPCW & "', '"         ' ���FLG�iSP)
                sql = sql & .WFRESSPCW & "', '"         ' ����FLG�iSP)
                sql = sql & .WFSMPLIDDO1CW & "', '"     ' �T���v��ID�iDO1�j
                sql = sql & .WFINDDO1CW & "', '"        ' ���FLG�iDO1)
                sql = sql & .WFRESDO1CW & "', '"        ' ����FLG�iDO1)
                sql = sql & .WFSMPLIDDO2CW & "', '"     ' �T���v��ID�iDO2�j
                sql = sql & .WFINDDO2CW & "', '"        ' ���FLG�iDO2)
                sql = sql & .WFRESDO2CW & "', '"        ' ����FLG�iDO2)
                sql = sql & .WFSMPLIDDO3CW & "', '"     ' �T���v��ID�iDO3�j
                sql = sql & .WFINDDO3CW & "', '"        ' ���FLG�iDO3)
                sql = sql & .WFRESDO3CW & "', '"        ' ����FLG�iDO3)
                sql = sql & .WFSMPLIDOT1CW & "', '"     ' �T���v��ID�iOT1�j
                sql = sql & .WFINDOT1CW & "', '"        ' ���FLG�iOT1)
                sql = sql & .WFRESOT1CW & "', '"        ' ����FLG�iOT1)
                sql = sql & .WFSMPLIDOT2CW & "', '"     ' �T���v��ID�iOT2�j
                sql = sql & .WFINDOT2CW & "', '"        ' ���FLG�iOT2)
                sql = sql & .WFRESOT2CW & "', '"        ' ����FLG�iOT2)
'                sql = sql & "NULL, "                    ' �T���v��ID�iAOi�j
'                sql = sql & "NULL, "                    ' ���FLG�iAOi�j
'                sql = sql & "NULL, "                    ' ����FLG�iAOi�j
                ''���ĉ����|�c���_�f�f�[�^�o�^�ǉ��@03/12/11 ooba
                sql = sql & .WFSMPLIDAOICW & "', '"     ' �T���v��ID�iAOi�j
                sql = sql & .WFINDAOICW & "', '"        ' ���FLG�iAOi�j
                sql = sql & .WFRESAOICW & "', '"        ' ����FLG�iAOi�j
                '' GD�ǉ��@05/01/31 ooba START =====================================>
                sql = sql & .WFSMPLIDGDCW & "', '"      ' �T���v��ID (GD)
                sql = sql & .WFINDGDCW & "', '"         ' ���FLG (GD)
                sql = sql & .WFRESGDCW & "', '"         ' ����FLG (GD)
                sql = sql & .WFHSGDCW & "', "           ' �ۏ�FLG (GD)
                '' GD�ǉ��@05/01/31 ooba END =======================================>
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                sql = sql & "'" & .EPSMPLIDB1CW & "', '"      ' �T���v��ID (BMD1E)
                sql = sql & .EPINDB1CW & "', '"         ' ���FLG (BMD1E)
                sql = sql & .EPRESB1CW & "', '"         ' ����FLG (BMD1E)
                sql = sql & .EPSMPLIDB2CW & "', '"      ' �T���v��ID (BMD2E)
                sql = sql & .EPINDB2CW & "', '"         ' ���FLG (BMD2E)
                sql = sql & .EPRESB2CW & "', '"         ' ����FLG (BMD2E)
                sql = sql & .EPSMPLIDB3CW & "', '"      ' �T���v��ID (BMD3E)
                sql = sql & .EPINDB3CW & "', '"         ' ���FLG (BMD3E)
                sql = sql & .EPRESB3CW & "', '"         ' ����FLG (BMD3E)
                sql = sql & .EPSMPLIDL1CW & "', '"      ' �T���v��ID (OSF1E)
                sql = sql & .EPINDL1CW & "', '"         ' ���FLG (OSF1E)
                sql = sql & .EPRESL1CW & "', '"         ' ����FLG (OSF1E)
                sql = sql & .EPSMPLIDL2CW & "', '"      ' �T���v��ID (OSF2E)
                sql = sql & .EPINDL2CW & "', '"         ' ���FLG (OSF2E)
                sql = sql & .EPRESL2CW & "', '"         ' ����FLG (OSF2E)
                sql = sql & .EPSMPLIDL3CW & "', '"      ' �T���v��ID (OSF3E)
                sql = sql & .EPINDL3CW & "', '"         ' ���FLG (OSF3E)
                sql = sql & .EPRESL3CW & "',  "         ' ����FLG (OSF3E)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
'''''           sql = sql & "NULL,"                     ' �T���v������
'''''           sql = sql & .SMPLPATCW & "', '"         ' �T���v���p�^�[��
                sql = sql & "NULL, "                    ' �T���v������
                sql = sql & "NULL, "                    ' �T���v���p�^�[��
                sql = sql & "'1', '"                    ' �����w���ʉ߃t���O 09/05/26 ooba
                sql = sql & .TSTAFFCW & "', "           ' �o�^�Ј�ID
                sql = sql & "sysdate, '"                ' �o�^���t
                sql = sql & .KSTAFFCW & "', "           ' �X�V�Ј�ID
                sql = sql & "sysdate, "                 ' �X�V���t
                sql = sql & "'0', "                     ' ���M�t���O
                sql = sql & "sysdate, "                 ' ���M���t
                sql = sql & "'0')"                      ' �����敪

                '' WriteDBLog sql, sDbName
        End With
        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_WfSmp_INS = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
'    Next i

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_WfSmp_INS = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'�T�v      :WF�T���v���Ǘ��̍X�V
'���Ұ��@�@:�ϐ���      ,IO ,�^                 ,����
'      �@�@:WFSMP �@�@�@,I  ,typ_XSDCW   �@     ,�V�T���v���Ǘ��iSXL�j
'      �@�@:�߂�l      ,O  ,FUNCTION_RETURN�@  ,�������݂̐���
'����      :�V�T���v���Ǘ��̃f�[�^���X�V����
'����      :2003/09/22  �쐬 �ѓc
Public Function DBDRV_WfSmp_UPD(WFSMP() As typ_XSDCW, i As Long) As FUNCTION_RETURN

    Dim sql As String
'    Dim i As Long
    Dim sDbName As String


    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_WfSmp_UPD"

    DBDRV_WfSmp_UPD = FUNCTION_RETURN_SUCCESS

    sDbName = "XSDCW"
'    For i = 1 To UBound(WFSMP)
       With WFSMP(i)
                sql = "UPDATE XSDCW "
                sql = sql & "SET "
'               sql = sql & "SXLIDCW      ='" & .SXLIDCW & "',"         ' SXLID"
                sql = sql & "SMPKBNCW     ='" & .SMPKBNCW & "',"        ' �T���v���敪"
'               sql = sql & "TBKBNCW      ='" & .TBKBNCW & "',"         ' T/B�敪"
                sql = sql & "REPSMPLIDCW  ='" & .REPSMPLIDCW & "',"     ' �T���v��ID"
                sql = sql & "XTALCW       ='" & .XTALCW & "',"          ' �����ԍ�"
                sql = sql & "INPOSCW      ='" & .INPOSCW & "',"         ' �������ʒu"
                sql = sql & "HINBCW       ='" & .HINBCW & "',"          ' �i��
                sql = sql & "REVNUMCW     ='" & .REVNUMCW & "',"        ' ���i�ԍ������ԍ�
                sql = sql & "FACTORYCW    ='" & .FACTORYCW & "',"       ' �H��
                sql = sql & "OPECW        ='" & .OPECW & "',"           ' ���Ə���
                sql = sql & "KTKBNCW      ='" & .KTKBNCW & "',"         ' �m��敪
                sql = sql & "SMCRYNUMCW   ='" & .SMCRYNUMCW & "',"      ' �T���v���u���b�NID
                sql = sql & "WFSMPLIDRSCW ='" & .WFSMPLIDRSCW & "',"    ' �T���v��ID(Rs)
                sql = sql & "WFSMPLIDRS1CW= NULL,"                      ' ����T���v��ID1�iRs�j
                sql = sql & "WFSMPLIDRS2CW= NULL,"                      ' ����T���v��ID2�iRs�j
                sql = sql & "WFINDRSCW    ='" & .WFINDRSCW & "',"       ' ���FLG�iRs)
'''''           sql = sql & "WFRESRS1CW   = NULL,"                      ' ����FLG1�iRs)
'''''           sql = sql & "WFRESRS2CW   = NULL,"                      ' ����FLG2�iRs)
                sql = sql & "WFRESRS1CW   ='" & .WFRESRS1CW & "',"      ' ����FLG1�iRs)
                sql = sql & "WFSMPLIDOICW ='" & .WFSMPLIDOICW & "',"    ' �T���v��ID�iOi�j
                sql = sql & "WFINDOICW    ='" & .WFINDOICW & "',"       ' ���FLG�iOi)
                sql = sql & "WFRESOICW    ='" & .WFRESOICW & "',"       ' ����FLG�iOi)
                sql = sql & "WFSMPLIDB1CW ='" & .WFSMPLIDB1CW & "',"    ' �T���v��ID�iB1�j
                sql = sql & "WFINDB1CW    ='" & .WFINDB1CW & "',"       ' ���FLG�iB1)
                sql = sql & "WFRESB1CW    ='" & .WFRESB1CW & "',"       ' ����FLG�iB1)
                sql = sql & "WFSMPLIDB2CW ='" & .WFSMPLIDB2CW & "',"    ' �T���v��ID�iB2�j
                sql = sql & "WFINDB2CW    ='" & .WFINDB2CW & "',"       ' ���FLG�iB2)
                sql = sql & "WFRESB2CW    ='" & .WFRESB2CW & "',"       ' ����FLG�iB2)
                sql = sql & "WFSMPLIDB3CW ='" & .WFSMPLIDB3CW & "',"    ' �T���v��ID�iB3�j
                sql = sql & "WFINDB3CW    ='" & .WFINDB3CW & "',"       ' ���FLG�iB3)
                sql = sql & "WFRESB3CW    ='" & .WFRESB3CW & "',"       ' ����FLG�iB3)
                sql = sql & "WFSMPLIDL1CW ='" & .WFSMPLIDL1CW & "',"    ' �T���v��ID�iL1�j
                sql = sql & "WFINDL1CW    ='" & .WFINDL1CW & "',"       ' ���FLG�iL1)
                sql = sql & "WFRESL1CW    ='" & .WFRESL1CW & "',"       ' ����FLG�iL1)
                sql = sql & "WFSMPLIDL2CW ='" & .WFSMPLIDL2CW & "',"    ' �T���v��ID�iL2�j
                sql = sql & "WFINDL2CW    ='" & .WFINDL2CW & "',"       ' ���FLG�iL2)
                sql = sql & "WFRESL2CW    ='" & .WFRESL2CW & "',"       ' ����FLG�iL2)
                sql = sql & "WFSMPLIDL3CW ='" & .WFSMPLIDL3CW & "',"    ' �T���v��ID�iL3�j
                sql = sql & "WFINDL3CW    ='" & .WFINDL3CW & "',"       ' ���FLG�iL3)
                sql = sql & "WFRESL3CW    ='" & .WFRESL3CW & "',"       ' ����FLG�iL3)
                sql = sql & "WFSMPLIDL4CW ='" & .WFSMPLIDL4CW & "',"    ' �T���v��ID�iL4�j
                sql = sql & "WFINDL4CW    ='" & .WFINDL4CW & "',"       ' ���FLG�iL4)
                sql = sql & "WFRESL4CW    ='" & .WFRESL4CW & "',"       ' ����FLG�iL4)
                sql = sql & "WFSMPLIDDSCW ='" & .WFSMPLIDDSCW & "',"    ' �T���v��ID�iDS�j
                sql = sql & "WFINDDSCW    ='" & .WFINDDSCW & "',"       ' ���FLG�iDS)
                sql = sql & "WFRESDSCW    ='" & .WFRESDSCW & "',"       ' ����FLG�iDS)
                sql = sql & "WFSMPLIDDZCW ='" & .WFSMPLIDDZCW & "',"    ' �T���v��ID�iDZ�j
                sql = sql & "WFINDDZCW    ='" & .WFINDDZCW & "',"       ' ���FLG�iDZ)
                sql = sql & "WFRESDZCW    ='" & .WFRESDZCW & "',"       ' ����FLG�iDZ)
                sql = sql & "WFSMPLIDSPCW ='" & .WFSMPLIDSPCW & "',"    ' �T���v��ID�iSP�j
                sql = sql & "WFINDSPCW    ='" & .WFINDSPCW & "',"       ' ���FLG�iSP)
                sql = sql & "WFRESSPCW    ='" & .WFRESSPCW & "',"       ' ����FLG�iSP)
                sql = sql & "WFSMPLIDDO1CW='" & .WFSMPLIDDO1CW & "',"   ' �T���v��ID�iDO1�j
                sql = sql & "WFINDDO1CW   ='" & .WFINDDO1CW & "',"      ' ���FLG�iDO1)
                sql = sql & "WFRESDO1CW   ='" & .WFRESDO1CW & "',"      ' ����FLG�iDO1)
                sql = sql & "WFSMPLIDDO2CW='" & .WFSMPLIDDO2CW & "',"   ' �T���v��ID�iDO2�j
                sql = sql & "WFINDDO2CW   ='" & .WFINDDO2CW & "',"      ' ���FLG�iDO2)
                sql = sql & "WFRESDO2CW   ='" & .WFRESDO2CW & "',"      ' ����FLG�iDO2)
                sql = sql & "WFSMPLIDDO3CW='" & .WFSMPLIDDO3CW & "',"   ' �T���v��ID�iDO3�j
                sql = sql & "WFINDDO3CW   ='" & .WFINDDO3CW & "',"      ' ���FLG�iDO3)
                sql = sql & "WFRESDO3CW   ='" & .WFRESDO3CW & "',"      ' ����FLG�iDO3)
                sql = sql & "WFSMPLIDOT1CW='" & .WFSMPLIDOT1CW & "',"   ' �T���v��ID�iOT1�j
                sql = sql & "WFSMPLIDOT2CW='" & .WFSMPLIDOT2CW & "',"   ' �T���v��ID�iOT2�j
               'add start 2003/05/21 hitec)matsumoto -------------------------
                sql = sql & "WFINDOT1CW   ='" & .WFINDOT1CW & "',"     ' ���FLG�iOT1)
                sql = sql & "WFRESOT1CW   ='" & .WFRESOT1CW & "',"     ' ����FLG�iOT1)
                sql = sql & "WFINDOT2CW   ='" & .WFINDOT2CW & "',"     ' ���FLG�iOT2)
                sql = sql & "WFRESOT2CW   ='" & .WFRESOT2CW & "',"     ' ����FLG�iOT2)
               'add end   2003/05/21 hitec)matsumoto -------------------------
'                sql = sql & "WFSMPLIDAOICW= NULL,"                      ' �T���v��ID�iAOi�j
'                sql = sql & "WFINDAOICW   = NULL,"                      ' ���FLG�iAOi�j
'                sql = sql & "WFRESAOICW   = NULL,"                      ' ����FLG�iAOi�j

                ''�c���_�f�f�[�^�o�^�ǉ��@03/12/11 ooba START ===============================>
                sql = sql & "WFSMPLIDAOICW='" & .WFSMPLIDAOICW & "',"   ' �T���v��ID�iAOi�j
                sql = sql & "WFINDAOICW   ='" & .WFINDAOICW & "',"      ' ���FLG�iAOi�j
                sql = sql & "WFRESAOICW   ='" & .WFRESAOICW & "',"      ' ����FLG�iAOi�j
                ''�c���_�f�f�[�^�o�^�ǉ��@03/12/11 ooba END =================================>

                '' GD�ǉ��@05/01/31 ooba START =============================================>
                sql = sql & "WFSMPLIDGDCW ='" & .WFSMPLIDGDCW & "',"    ' �T���v��ID (GD)
                sql = sql & "WFINDGDCW    ='" & .WFINDGDCW & "', "      ' ���FLG (GD)
                sql = sql & "WFRESGDCW    ='" & .WFRESGDCW & "', "      ' ����FLG (GD)
                sql = sql & "WFHSGDCW     ='" & .WFHSGDCW & "', "       ' �ۏ�FLG (GD)
                '' GD�ǉ��@05/01/31 ooba END ===============================================>

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                sql = sql & "EPSMPLIDL1CW = '" & .EPSMPLIDL1CW & "', "  ' �T���v��ID (OSF1E)
                sql = sql & "EPINDL1CW = '" & .EPINDL1CW & "', "        ' ���FLG (OSF1E)
                sql = sql & "EPRESL1CW = '" & .EPRESL1CW & "', "        ' ����FLG (OSF1E)
                sql = sql & "EPSMPLIDL2CW = '" & .EPSMPLIDL2CW & "', "  ' �T���v��ID (OSF2E)
                sql = sql & "EPINDL2CW = '" & .EPINDL2CW & "', "        ' ���FLG (OSF2E)
                sql = sql & "EPRESL2CW = '" & .EPRESL2CW & "', "        ' ����FLG (OSF2E)
                sql = sql & "EPSMPLIDL3CW = '" & .EPSMPLIDL3CW & "', "  ' �T���v��ID (OSF3E)
                sql = sql & "EPINDL3CW = '" & .EPINDL3CW & "', "        ' ���FLG (OSF3E)
                sql = sql & "EPRESL3CW = '" & .EPRESL3CW & "', "        ' ����FLG (OSF3E)
                sql = sql & "EPSMPLIDB1CW = '" & .EPSMPLIDB1CW & "', "  ' �T���v��ID (BMD1E)
                sql = sql & "EPINDB1CW = '" & .EPINDB1CW & "', "        ' ���FLG (BMD1E)
                sql = sql & "EPRESB1CW = '" & .EPRESB1CW & "', "        ' ����FLG (BMD1E)
                sql = sql & "EPSMPLIDB2CW = '" & .EPSMPLIDB2CW & "', "  ' �T���v��ID (BMD2E)
                sql = sql & "EPINDB2CW = '" & .EPINDB2CW & "', "        ' ���FLG (BMD2E)
                sql = sql & "EPRESB2CW = '" & .EPRESB2CW & "', "        ' ����FLG (BMD2E)
                sql = sql & "EPSMPLIDB3CW = '" & .EPSMPLIDB3CW & "', "  ' �T���v��ID (BMD3E)
                sql = sql & "EPINDB3CW = '" & .EPINDB3CW & "', "        ' ���FLG (BMD3E)
                sql = sql & "EPRESB3CW = '" & .EPRESB3CW & "', "        ' ����FLG (BMD3E)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

                sql = sql & "SMPLNUMCW    = NULL,"                      ' �T���v������
                sql = sql & "SMPLPATCW    = NULL,"                      ' �T���v���p�^�[��
                sql = sql & "NUKISIFLGCW  = '1',"                       ' �����w���ʉ߃t���O 09/05/26 ooba
'               sql = sql & "TSTAFFCW     ='" & .TSTAFFCW & "' "        ' �o�^�Ј�ID
'''''           sql = sql & "TDAYCW       = sysdate,"                   ' �o�^���t
                sql = sql & "KSTAFFCW     ='" & .KSTAFFCW & "',"        ' �X�V�Ј�ID
                sql = sql & "KDAYCW       = sysdate, "                  ' �X�V���t"
                sql = sql & "SNDKCW       ='0',"                        ' ���M�t���O"
                sql = sql & "SNDDAYCW     = sysdate "                   ' ���M���t"

                sql = sql & "WHERE "
                sql = sql & "SXLIDCW ='" & .SXLIDCW & "'and "           ' SXLID"
'               sql = sql & "SMPKBNCW='" & .SMPKBNCW & "'"              ' �T���v���敪"
                sql = sql & "TBKBNCW ='" & .TBKBNCW & "'"               ' TB�敪"

                '' WriteDBLog sql, sDbName
        End With
        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_WfSmp_UPD = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
'    Next i

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_WfSmp_UPD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'�z�[���h���b�g��������     '04/06/29 ooba �쐬
Public Function HoldLot_Get740(xtal As String, HOLDBCA As String, WFHOLDFLGCA As String) As FUNCTION_RETURN
    Dim sql As String
    Dim rs As OraDynaset
    Dim blkcnt As Integer
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function HoldLot_Get740"

    HoldLot_Get740 = FUNCTION_RETURN_SUCCESS

    sql = "select CRYNUMCA, HOLDBCA, WFHOLDFLGCA "
    sql = sql & "from XSDCA "
    sql = sql & "where LIVKCA = '0' "
    sql = sql & "and CRYNUMCA in ( "
    sql = sql & "     select BLOCKID "
    sql = sql & "     from TBCMY001 "
    sql = sql & "     where SBLOCKID in ( "
    sql = sql & "             select SBLOCKID "
    sql = sql & "             from TBCMY001 "
    sql = sql & "             where BLOCKID = '" & xtal & "' "
    sql = sql & "             ) "
    sql = sql & ") "
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)

    If rs.EOF = False Then
        For blkcnt = 1 To rs.RecordCount
            If rs("HOLDBCA") = "1" Then
                HOLDBCA = rs("HOLDBCA")
                Exit For
            Else
                HOLDBCA = " "
            End If
            If rs("WFHOLDFLGCA") = "1" Then
                WFHOLDFLGCA = rs("WFHOLDFLGCA")
                Exit For
            Else
                WFHOLDFLGCA = " "
            End If
            rs.MoveNext
        Next blkcnt
    Else
        HoldLot_Get740 = FUNCTION_RETURN_FAILURE
    End If
    rs.Close
    Set rs = Nothing

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    HoldLot_Get740 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'��(f)
'
'�@�\       :�i�ԊǗ� - �\���pTBCME041�f�[�^�擾
'
'�Ԃ�l     :0 - ����I��
'           :1 - �ُ�I��
'
'������     :records()  - ���o���R�[�h
'           :sCryNum    - �����ԍ�
'           :sBlockId   - �u���b�NID
'
'�@�\����   :�\���p�i�ԃf�[�^���擾����
'
'����       :2005/12/26�@SMP)�ΐ�@�쐬
'
'���l       :SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs
'           WF���ύX�ł́ATBCME041���X�V����Ȃ��̂ŁAXSDCA��XSDCB���g�p���ăf�[�^���쐬����
Private Function DBDRV_GetTBCME041_Clone(records() As typ_TBCME041, _
                                        sCryNum As String, _
                                        sBlockId() As String) As FUNCTION_RETURN
    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    Dim i           As Long
    Dim lsSXL()     As String
    Dim llSXLTop    As Long         'SXL�̌������J�n�ʒu
    Dim llLastCBLen As Long         '�ŏISXL�̒���
    Dim tmpXSDCA    As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function DBDRV_GetTBCME041_Clone"

    tmpXSDCA = "   AND a.CRYNUMCA IN ("
    For i = 1 To UBound(sBlockId)
        tmpXSDCA = tmpXSDCA & "'" & sBlockId(i) & "',"
    Next i
    tmpXSDCA = Mid(tmpXSDCA, 1, Len(tmpXSDCA) - 1)
    tmpXSDCA = tmpXSDCA & ") "

    ''SQL��g�ݗ��Ă�
    sql = ""
    sql = sql & " SELECT"
    sql = sql & "   a.XTALCA"
    sql = sql & "  ,a.HINBCA"
    sql = sql & "  ,a.INPOSCA"
    sql = sql & "  ,a.REVNUMCA"
    sql = sql & "  ,a.FACTORYCA"
    sql = sql & "  ,a.OPECA"
    sql = sql & "  ,b.RLENCB"
    sql = sql & "  ,a.SXLIDCA"
    sql = sql & "  ,NVL(b.INPOSCB,0) INPOSCB"
    sql = sql & " FROM"
    sql = sql & "   XSDCA A"
    sql = sql & "  ,XSDCB B"
    sql = sql & "  ,XSDC2 C"
    sql = sql & " WHERE a.SXLIDCA = b.SXLIDCB"
    sql = sql & "   AND a.CRYNUMCA  = c.CRYNUMC2"
'    sql = sql & "   AND c.WFHUFLG  = '1'"
    sql = sql & "   AND a.LIVKCA  = '0'"
    sql = sql & "   AND a.XTALCA = '" & Trim(sCryNum) & "'"
    sql = sql & tmpXSDCA
    sql = sql & " ORDER BY"
    sql = sql & "   a.SXLIDCA"
    sql = sql & "  ,a.INPOSCA"

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    ReDim records(0) As typ_TBCME041
    ReDim lsSXL(0) As String
    i = 0
    llSXLTop = 0
    llLastCBLen = 0

    ''���o���ʂ��i�[����
    Do Until rs.EOF '�f�[�^���Ȃ��Ȃ�܂Ŏ擾
        i = i + 1
        ReDim Preserve records(i) As typ_TBCME041
        ReDim Preserve lsSXL(i) As String

        With records(i)
            .CRYNUM = rs("XTALCA")          ' �����ԍ�
            .INGOTPOS = rs("INPOSCA")       ' �������J�n�ʒu
            .hinban = rs("HINBCA")          ' �i��
            .REVNUM = rs("REVNUMCA")        ' ���i�ԍ������ԍ�
            .factory = rs("FACTORYCA")      ' �H��
            .opecond = rs("OPECA")          ' ���Ə���
'            .LENGTH = rs("RLENCB")          ' XSDCB��SXL����
            lsSXL(i) = rs("SXLIDCA")        ' SXLID

            '�������Čv�Z����
            records(i - 1).LENGTH = .INGOTPOS - records(i - 1).INGOTPOS
            '�ŏISXL�̌�����TOP�ʒu��ێ�
            If lsSXL(i) <> lsSXL(i - 1) Then
                llSXLTop = rs("INPOSCB")
            End If

            '�u���b�N�̕ς��ڂŒ������v�Z����
            If records(i).CRYNUM <> records(i - 1).CRYNUM And i <> 1 Then
                records(UBound(records)).LENGTH = (llSXLTop + llLastCBLen) - records(UBound(records)).INGOTPOS
            End If

            '�ŏISXLID��XSCB.RLENB��ێ�����
            llLastCBLen = rs("RLENCB")

        End With
        rs.MoveNext
    Loop

    '�u���b�N�̍Ō�̕i�Ԃ̒����� (SXL��INGOTPOS + XSDCB.RLENB)-INGOTPOS
    records(UBound(records)).LENGTH = (llSXLTop + llLastCBLen) - records(UBound(records)).INGOTPOS

    rs.Close

    '�f�[�^�����̏ꍇ�G���[
    If i = 0 Then
        ReDim records(0) As typ_TBCME041
        DBDRV_GetTBCME041_Clone = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    DBDRV_GetTBCME041_Clone = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_GetTBCME041_Clone = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit

End Function

'��(s)
'
'�@�\       :�i�ԍ\���̃f�[�^����
'
'�Ԃ�l     :�Ȃ�
'
'������     :tMotoHinMng    :�S�̂̕i�ԃf�[�^
'            tUpdateHinMng  :��������i�ԃf�[�^
'
'�@�\����   :tMotoHinMng�̃f�[�^�̂����AtUpdateHinMng�̕�����tUpdateHinMng�ɒu������B
'
'����       :2005/12/26�@SMP)�ΐ�@�쐬
'
'���l       :SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs
'           tUpdateHinMng��1�u���b�N���̕i�ԃf�[�^�Ƃ���B
'
Private Sub s_cmbc036_2_F_SynHinban(tMotoHinMng() As typ_TBCME041, tUpdateHinMng() As typ_TBCME041)
    Dim tHinMng()   As typ_TBCME041         '�i�ԊǗ��e�[�u�����[�N
    Dim j           As Long
    Dim k           As Long
    Dim i           As Long
    Dim sHinban1    As String               '�ꗗ��̕i�Ԃ���1
    Dim sHinban2    As String               '�ꗗ��̕i�Ԃ���2
    Dim sRev        As String               '�i�Ԃ���2����擾�������i�ԍ������ԍ�
    Dim sFac        As String               '�i�Ԃ���2����擾�����H��
    Dim sOPE        As String               '�i�Ԃ���2����擾�������Ə���
    Dim sCrystalNo  As String               '�����ԍ�
    Dim lCrystalPos As Integer              '�������ʒu
    Dim lBlockTP    As Long                 '�������ʒu�i�u���b�NTOP�j
    Dim lBlockBP    As Long                 '�������ʒu�i�u���b�NBottom�j
    Dim llRow       As Long
    Dim lLen        As Long                 '����
    Dim llUpdateFlg As Long                 '�u�����ς݃t���O�i�֌W�u���b�N������������:1 ���ĂȂ�:0�j

    ''�i�ԊǗ��e�[�u���̃f�[�^�\���̍X�V
    ReDim tHinMng(0) As typ_TBCME041
    i = 1
    j = 1
    ''�����ԍ��擾
    sCrystalNo = tUpdateHinMng(1).CRYNUM
    ''�������ʒu�i�u���b�NTOP�j�擾
    lBlockTP = tUpdateHinMng(1).INGOTPOS
    ''�������ʒu�i�u���b�NBottom�j�擾
    lBlockBP = tUpdateHinMng(UBound(tUpdateHinMng)).INGOTPOS + tUpdateHinMng(UBound(tUpdateHinMng)).LENGTH
    ''���i�Ԃ̔z��0�̌������J�n�ʒu��������
    tMotoHinMng(0).INGOTPOS = 0
    '�u�����ς݃t���O������
    llUpdateFlg = 0

    For i = 1 To UBound(tMotoHinMng)
'        '���i�Ԃ̌������ʒu���֌W�u���b�N�̍�TOP�ʒu��菬�����A�܂��́A
'        '���i�Ԃ̌������ʒu���֌W�u���b�N�̍�Bottom�ʒu�ȏォ�A
'        '  ��O�̌��i�Ԃ̌������ʒu���֌W�u���b�N�̍�Bottom�ʒu���傫���ꍇ
'        If tMotoHinMng(i).INGOTPOS < lBlockTP Or _
'         (tMotoHinMng(i).INGOTPOS > lBlockBP And tMotoHinMng(i - 1).INGOTPOS >= lBlockBP) Then

        ''���i�Ԃ̌������ʒu���֌W�u���b�N�̍�TOP�ʒu��菬�����A�܂���
        ''���i�Ԃ̌������ʒu���֌W�u���b�N�̍�BOTTOM�ʒu���傫�����A�֌W�u���b�N�𔽉f�ς݂̏ꍇ
        If tMotoHinMng(i).INGOTPOS < lBlockTP Or _
          (tMotoHinMng(i).INGOTPOS >= lBlockBP And llUpdateFlg = 1) Then
            ''�������ʒu���u���b�N�͈̔͊O
            ReDim Preserve tHinMng(j) As typ_TBCME041
            tHinMng(j) = tMotoHinMng(i)
            j = j + 1

        Else
            ''�������ʒu���u���b�N�͈͓̔�
            '�u�����t���O�Z�b�g
            llUpdateFlg = 1
            ' ���̕i�Ԃ̒������ς��ꍇ�A��������
            ' (�֌W�u���b�N�̌������ʒu - ���̕i�Ԃ̌������ʒu)
            If tMotoHinMng(i).INGOTPOS <> lBlockTP Then
                tHinMng(j - 1).LENGTH = lBlockTP - tHinMng(j - 1).INGOTPOS
            End If

            For llRow = 1 To UBound(tUpdateHinMng)
                '�i�Ԏ擾
                sHinban1 = tUpdateHinMng(llRow).hinban
                '���i�ԍ������ԍ�
                sRev = tUpdateHinMng(llRow).REVNUM
                '�H��
                sFac = tUpdateHinMng(llRow).factory
                '���Ə���
                sOPE = tUpdateHinMng(llRow).opecond
                '�������ʒu�擾
                lCrystalPos = CLng(tUpdateHinMng(llRow).INGOTPOS)
                '����
                lLen = CLng(tUpdateHinMng(llRow).LENGTH)

                '���[�N�̈�ɐݒ�
                ReDim Preserve tHinMng(j) As typ_TBCME041

                tHinMng(j).CRYNUM = sCrystalNo
                tHinMng(j).INGOTPOS = lCrystalPos
                tHinMng(j).hinban = sHinban1
                tHinMng(j).REVNUM = CInt(sRev)
                tHinMng(j).factory = sFac
                tHinMng(j).opecond = sOPE
                tHinMng(j).LENGTH = CInt(lLen)

                j = j + 1

            Next llRow
            ''�u���b�N�͈͂𔲂���܂Ői�߂�
            Do While (1)
'                If tMotoHinMng(i).INGOTPOS < lBlockTP Or _
'                (tMotoHinMng(i).INGOTPOS > lBlockBP And tMotoHinMng(i - 1).INGOTPOS >= lBlockBP) Then

                ''�������J�n�ʒu���֌W�u���b�N��Bottom�ʒu���傫���i�Ԃ�T��
                If tMotoHinMng(i).INGOTPOS >= lBlockBP Then
                    '���Ƃ��Ƃ̕i�Ԃ̒������֌W�u���b�N��Bottom�ʒu��蒷���ꍇ�A�i�Ԃ���t������
                    '���̂Ƃ��A�����ƌ������J�n�ʒu�𒲐�����
                    If tMotoHinMng(i).INGOTPOS <> lBlockBP Then
                        ReDim Preserve tHinMng(j) As typ_TBCME041
                        tHinMng(j) = tMotoHinMng(i - 1)
                        tHinMng(j).LENGTH = tMotoHinMng(i - 1).LENGTH - (lBlockBP - tMotoHinMng(i - 1).INGOTPOS)
                        tHinMng(j).INGOTPOS = lBlockBP
                        j = j + 1
                    End If

                    '���[�v�̃J�E���g�A�b�v�ŃJ�E���^���i�ނ̂ŁA��߂�
                    i = i - 1
                    Exit Do
                End If
                i = i + 1
                '�f�[�^�������Ȃ�����I��
                If i > UBound(tMotoHinMng) Then
                    Exit Do
                End If
            Loop
        End If
    Next i

    ''�ŏI�i�Ԍ�ɒǉ������ꍇ���l��
    If tMotoHinMng(UBound(tMotoHinMng)).INGOTPOS < lBlockTP Then
        For llRow = 1 To UBound(tUpdateHinMng)
            '�i�Ԏ擾
            sHinban1 = tUpdateHinMng(llRow).hinban
            '���i�ԍ������ԍ�
            sRev = tUpdateHinMng(llRow).REVNUM
            '�H��
            sFac = tUpdateHinMng(llRow).factory
            '���Ə���
            sOPE = tUpdateHinMng(llRow).opecond
            '�������ʒu�擾
            lCrystalPos = CLng(tUpdateHinMng(llRow).INGOTPOS)
            '����
            lLen = CLng(tUpdateHinMng(llRow).LENGTH)

            '���[�N�̈�ɐݒ�
            ReDim Preserve tHinMng(j) As typ_TBCME041

            tHinMng(j).CRYNUM = sCrystalNo
            tHinMng(j).INGOTPOS = lCrystalPos
            tHinMng(j).hinban = sHinban1
            tHinMng(j).REVNUM = CInt(sRev)
            tHinMng(j).factory = sFac
            tHinMng(j).opecond = sOPE
            tHinMng(j).LENGTH = CInt(lLen)

            j = j + 1

        Next llRow
    End If

    ''tMotoHinMng�ɐݒ�
    ReDim tMotoHinMng(UBound(tHinMng)) As typ_TBCME041
    For i = 1 To UBound(tHinMng)
        tMotoHinMng(i) = tHinMng(i)
    Next i

End Sub

'��(f)
'
'�@�\       :�V�T���v���Ǘ�(SXL)�f�[�^�⊮
'
'�Ԃ�l     :0 - ����I��
'           :1 - �ُ�I��
'
'������     :tXSDCW()   - �擾XSDCW�f�[�^
'           :sCryNum    - �����ԍ�
'           :sBlockId   - �֌W�u���b�NID
'
'�@�\����   :XSDCW�ɕ⊮����f�[�^��XSDCA����擾���A�����œn���ꂽXSDCW�̃f�[�^�̏C�����s��
'
'����       :2005/12/26�@SMP)�ΐ�@�쐬
'
'���l       :SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs
'           WF���ύX�ł́AXSDCW���X�V����Ȃ��̂ŁAXSDCA��XSDC2���g�p���ăf�[�^���쐬����
Public Function DBDRV_GetXSDCWUpdate(tXSDCW() As typ_XSDCW, _
                                      sCryNum As String, _
                                      sBlockId() As String) As FUNCTION_RETURN
    Dim sql         As String       'SQL�S��
    Dim rs          As OraDynaset   'RecordSet
    Dim i           As Long
    Dim j           As Long
    Dim lsWhere     As String
    Dim lsHinban    As String       '�⊮���邩�`�F�b�N�p�̕i��
    Dim tmpXSDCA()  As typ_XSDCA    '�֌W�u���b�NID����XSDCA�f�[�^
    Dim tmpXSDCA2() As typ_XSDCA    '�֌W�u���b�NID����XSDCA�f�[�^
    Dim tmpXSDCW()  As typ_XSDCW    'XSDCW���[�N�̈�
    Dim liUpdateFLG As Integer      '�X�V�t���O
    Dim liUpFLG2    As Integer      '�X�V�t���O
    Dim liEndSxlFLG As Integer      '�ŏISXL�t���O

    Dim lsHinWork   As String
    Dim lsSXLWork   As String
    Dim lsSendFlgWork   As String
    Dim llCnt       As Long

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function DBDRV_GetXSDCWUpdate"

    '���폜 2006/03/20 ��Q�Ή� SMP�ΐ� WF���ύX����Ă��Ȃ��ꍇ�A�������\������Ȃ���Q�ɑΉ�
'    ''������
'    ReDim tSXLID(0)
    '���폜 2006/03/20 ��Q�Ή� SMP�ΐ� WF���ύX����Ă��Ȃ��ꍇ�A�������\������Ȃ���Q�ɑΉ�

    ''SQL��g�ݗ��Ă�

    '�u���b�N�̏������ʂō쐬
    lsWhere = "   AND a.CRYNUMCA IN ("
    For i = 1 To UBound(sBlockId)
        lsWhere = lsWhere & "'" & sBlockId(i) & "',"
    Next i
    lsWhere = Mid(lsWhere, 1, Len(lsWhere) - 1)
    lsWhere = lsWhere & ") "

    sql = ""
    sql = sql & " SELECT"
    sql = sql & "   a.XTALCA"       '�����ԍ�
    sql = sql & "  ,a.HINBCA"       '�i��
    sql = sql & "  ,a.INPOSCA"      '�������J�n�ʒu
    sql = sql & "  ,a.REVNUMCA"     '���i�ԍ������ԍ�
    sql = sql & "  ,a.FACTORYCA"    '�H��
    sql = sql & "  ,a.OPECA"        '���Ə���
    sql = sql & "  ,a.GNLCA"        '���ݒ���
    sql = sql & "  ,a.SXLIDCA"      'SXLID
    sql = sql & "  ,a.CRYNUMCA"     '�u���b�NID
    sql = sql & "  ,NVL(b.WFHUFLG,' ') WFHUFLG"      'WF�U��FLG
    sql = sql & " FROM"
    sql = sql & "   XSDCA A"
    sql = sql & "  ,XSDC2 B"
    sql = sql & " WHERE a.LIVKCA  = '0'"
    sql = sql & "   AND a.CRYNUMCA = b.CRYNUMC2"
    sql = sql & "   AND a.XTALCA = '" & Trim(sCryNum) & "'"
    sql = sql & lsWhere
    sql = sql & " ORDER BY"
    sql = sql & "   a.XTALCA"
    sql = sql & "  ,a.INPOSCA"

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    ReDim tmpXSDCA(0) As typ_XSDCA
    ReDim tmpXSDCW(0) As typ_XSDCW
    '���폜 2006/03/20 ��Q�Ή� SMP�ΐ� WF���ύX����Ă��Ȃ��ꍇ�A�������\������Ȃ���Q�ɑΉ�
'    ReDim tSXLID(0)
    '���폜 2006/03/20 ��Q�Ή� SMP�ΐ� WF���ύX����Ă��Ȃ��ꍇ�A�������\������Ȃ���Q�ɑΉ�
    i = 0
    liUpdateFLG = 0
    tmpXSDCA(0).SNDKCA = 0
    ''���o���ʂ��i�[����
    Do Until rs.EOF '�f�[�^���Ȃ��Ȃ�܂Ŏ擾
'        i = i + 1
'        ReDim Preserve tmpXSDCA(i) As typ_XSDCA

        lsHinWork = rs("HINBCA")        ' �i��
        lsSendFlgWork = rs("WFHUFLG")   ' WF�U��FLG
        lsSXLWork = rs("SXLIDCA")       ' SXLID

        '1�O�̃��R�[�h�̕i�ԁASXLID�������ŁA�Ƃ��ɐU�փt���O�������Ă���ꍇ�A���̃��R�[�h�͎擾���Ȃ�
'        If tmpXSDCA(i).HINBCA = lsHinWork _
'          And tmpXSDCA(i).SNDKCA = "1" _
'          And lsSendFlgWork = "1" _
'          And tmpXSDCA(i).SXLIDCA = lsSXLWork Then
'
'        Else

            i = i + 1
            ReDim Preserve tmpXSDCA(i) As typ_XSDCA
            With tmpXSDCA(i)

                .XTALCA = rs("XTALCA")          ' �����ԍ�
                .INPOSCA = rs("INPOSCA")        ' �������J�n�ʒu
                .HINBCA = rs("HINBCA")          ' �i��
                .REVNUMCA = rs("REVNUMCA")      ' ���i�ԍ������ԍ�
                .FACTORYCA = rs("FACTORYCA")    ' �H��
                .OPECA = rs("OPECA")            ' ���Ə���
                .GNLCA = rs("GNLCA")            ' ����
                .SXLIDCA = rs("SXLIDCA")        ' SXLID
                .CRYNUMCA = rs("CRYNUMCA")      ' SXLID
                'WF�U��FLG��CA�ɍ��ڂ������̂ŁA�ς��ɑ��M�t���O�ɓ����
                .SNDKCA = rs("WFHUFLG")         ' WF�U��FLG
                If .SNDKCA = "1" Then
                    liUpdateFLG = 1
                End If

            End With
'        End If

        rs.MoveNext
    Loop

    rs.Close

    '�f�[�^�����̏ꍇ�G���[
    If i = 0 Then
'        ReDim records(0) As typ_XSDCA
        DBDRV_GetXSDCWUpdate = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''�֌W�u���b�N��WF�U��FLG�����ׂė����Ă��Ȃ��ꍇ�AWF���ύX�ōX�V����Ă��Ȃ��̂ŁACW�̕⊮�͂��Ȃ�
    If liUpdateFLG = 0 Then
        ReDim tXSDCW(0)
        DBDRV_GetXSDCWUpdate = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If

    '���ǉ� 2006/03/20 ��Q�Ή� SMP�ΐ� WF���ύX����Ă��Ȃ��ꍇ�A�������\������Ȃ���Q�ɑΉ�
    ''������
    ReDim tSXLID(0)
    '���ǉ� 2006/03/20 ��Q�Ή� SMP�ΐ� WF���ύX����Ă��Ȃ��ꍇ�A�������\������Ȃ���Q�ɑΉ�

    '' �����œn���ꂽXSDCW�̃f�[�^�ɑ΂��āAWF���ύX�ŕύX��������⊮���Ă��
    i = 0
    j = 1
    liUpdateFLG = 0
    'XSDCW��Top��Bottom�ŃZ�b�g�Ȃ̂�2�Ði�߂�
    For i = 1 To UBound(tXSDCW) Step 2
        liUpFLG2 = 0
        liUpdateFLG = 0
        '1SXL����XSDCW�̕i�ԂƈႤ�i�Ԃ����݂����ꍇ = WF���ύX���ꂽ�ꍇ�A�⊮����B
        lsHinban = tXSDCW(i).HINBCW
        For j = 1 To UBound(tmpXSDCA)
            If tmpXSDCA(j).SXLIDCA = tXSDCW(i).SXLIDCW Then
                If tmpXSDCA(j).HINBCA <> lsHinban Then
                    liUpFLG2 = 1

                    For llCnt = 1 To UBound(tmpXSDCA)
                        ReDim Preserve tSXLID(llCnt)
                        tSXLID(llCnt).LOTID = tmpXSDCA(llCnt).CRYNUMCA
                        tSXLID(llCnt).SXLID = tmpXSDCA(llCnt).SXLIDCA
                        tSXLID(llCnt).INGOTPOS = tmpXSDCA(llCnt).INPOSCA
                    Next llCnt

                    Exit For
                End If
            End If
        Next j

        If liUpFLG2 = 1 Then
            ReDim tmpXSDCA2(0)
            For j = 1 To UBound(tmpXSDCA)
                '' ��O�̕i�Ԃ��Ⴄ�A���́@�i�Ԃ������ł��A�ǂ��炩���U�ւ����Ă��Ȃ��ꍇ
                    'SNDKCA=0��SNDKCA="0"�ɕύX�@06/06/15 ooba
                    'And (tmpXSDCA(j).SNDKCA = 0 Or tmpXSDCA(j - 1).SNDKCA = 0)
                If (tmpXSDCA(j).HINBCA <> tmpXSDCA(j - 1).HINBCA) _
                Or (tmpXSDCA(j).HINBCA = tmpXSDCA(j - 1).HINBCA _
                    And (tmpXSDCA(j).SNDKCA = "0" Or tmpXSDCA(j - 1).SNDKCA = "0") _
                   ) Then
                    ReDim Preserve tmpXSDCA2(UBound(tmpXSDCA2) + 1)
                    tmpXSDCA2(UBound(tmpXSDCA2)) = tmpXSDCA(j)

                End If
            Next j
            ReDim tmpXSDCA(UBound(tmpXSDCA2))
            For j = 1 To UBound(tmpXSDCA2)
                tmpXSDCA(j) = tmpXSDCA2(j)
            Next j


            liUpdateFLG = 0
            For j = 1 To UBound(tmpXSDCA)
                ''XSDCA�̃g�b�v�̈ʒu���AXSDCW�̃g�b�v�ȏ�ABottom��菬�����ꍇ�A
                ''���i�Ԃ��ς�����ꍇ�ɕ⊮����
                If tXSDCW(i).INPOSCW <= tmpXSDCA(j).INPOSCA _
                 And tXSDCW(i + 1).INPOSCW > tmpXSDCA(j).INPOSCA _
                 And tmpXSDCA(j).HINBCA <> tmpXSDCA(j - 1).HINBCA Then

'                ''XSDCA�̃g�b�v�̈ʒu���AXSDCW�̃g�b�v�ȏ�ABottom��菬�����ꍇ�A
'                ''���i�Ԃ��ς�����ꍇ�A���͕i�Ԃ������ŁA�����U�ւ����Ă���ɕ⊮����
'                If tXSDCW(i).INPOSCW <= tmpXSDCA(j).INPOSCA _
'                 And tXSDCW(i + 1).INPOSCW > tmpXSDCA(j).INPOSCA _
'                 And ( _
'                           (tmpXSDCA(j).HINBCA = tmpXSDCA(j - 1).HINBCA _
'                              And ((CLng(tmpXSDCA(j).SNDKCA) + CLng(tmpXSDCA(j - 1).SNDKCA)) <> 2)) _
'                       Or _
'                           (tmpXSDCA(j).HINBCA <> tmpXSDCA(j - 1).HINBCA) _
'                     ) Then

                    ReDim Preserve tmpXSDCW(UBound(tmpXSDCW) + 2) As typ_XSDCW

                    'SXL���ł̕i�Ԃ̐����J�E���g
                    liUpdateFLG = liUpdateFLG + 1

                    '' TOP�ʒu�ݒ� ------------------------------------------------------------------------------------
                    '��XSDCW��TOP�̈ʒu�́AXSDCW�̃f�[�^���g�p����
                    If liUpdateFLG = 1 Then
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SXLIDCW = tXSDCW(i).SXLIDCW              'SXLID
                        tmpXSDCA(j).SXLIDCA = tXSDCW(i).SXLIDCW 'SXLID�ۑ�
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SMPKBNCW = tXSDCW(i).SMPKBNCW            '�T���v���敪
                        tmpXSDCW(UBound(tmpXSDCW) - 1).TBKBNCW = tXSDCW(i).TBKBNCW              'T/B�敪
                        tmpXSDCW(UBound(tmpXSDCW) - 1).REPSMPLIDCW = tXSDCW(i).REPSMPLIDCW      '��\�T���v��ID
                        tmpXSDCW(UBound(tmpXSDCW) - 1).XTALCW = tXSDCW(i).XTALCW                '�����ԍ�
                        tmpXSDCW(UBound(tmpXSDCW) - 1).INPOSCW = tXSDCW(i).INPOSCW              '�������ʒu
                        tmpXSDCW(UBound(tmpXSDCW) - 1).HINBCW = tmpXSDCA(j).HINBCA              '�i��(XSDCA)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).REVNUMCW = tmpXSDCA(j).REVNUMCA          '���i�ԍ������ԍ�(XSDCA)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).FACTORYCW = tmpXSDCA(j).FACTORYCA        '�H��(XSDCA)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).OPECW = tmpXSDCA(j).OPECA                '���Ə���(XSDCA)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).KTKBNCW = tXSDCW(i).KTKBNCW              '�m��敪
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SMCRYNUMCW = tXSDCW(i).SMCRYNUMCW        '�T���v���u���b�NID
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDRSCW = tXSDCW(i).WFSMPLIDRSCW    '�T���v��ID(Rs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDRS1CW = tXSDCW(i).WFSMPLIDRS1CW  '����T���v��ID1(Rs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDRS2CW = tXSDCW(i).WFSMPLIDRS2CW  '����T���v��ID2(Rs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDRSCW = tXSDCW(i).WFINDRSCW          '���FLG�iRs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESRS1CW = tXSDCW(i).WFRESRS1CW        '����FLG1�iRs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESRS2CW = tXSDCW(i).WFRESRS2CW        '����FLG2�iRs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDOICW = tXSDCW(i).WFSMPLIDOICW    '�T���v��ID�iOi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDOICW = tXSDCW(i).WFINDOICW          '���FLG�iOi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESOICW = tXSDCW(i).WFRESOICW          '����FLG�iOi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDB1CW = tXSDCW(i).WFSMPLIDB1CW    '�T���v��ID�iB1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDB1CW = tXSDCW(i).WFINDB1CW          '���FLG�iB1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESB1CW = tXSDCW(i).WFRESB1CW          '����FLG�iB1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDB2CW = tXSDCW(i).WFSMPLIDB2CW    '�T���v��ID�iB2�j
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDB2CW = tXSDCW(i).WFINDB2CW          '���FLG�iB2�j
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESB2CW = tXSDCW(i).WFRESB2CW          '����FLG�iB2�j
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDB3CW = tXSDCW(i).WFSMPLIDB3CW    '�T���v��ID�iB3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDB3CW = tXSDCW(i).WFINDB3CW          '���FLG�iB3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESB3CW = tXSDCW(i).WFRESB3CW          '����FLG�iB3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDL1CW = tXSDCW(i).WFSMPLIDL1CW    '�T���v��ID�iL1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDL1CW = tXSDCW(i).WFINDL1CW          '���FLG�iL1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESL1CW = tXSDCW(i).WFRESL1CW          '����FLG�iL1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDL2CW = tXSDCW(i).WFSMPLIDL2CW    '�T���v��ID�iL2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDL2CW = tXSDCW(i).WFINDL2CW          '���FLG�iL2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESL2CW = tXSDCW(i).WFRESL2CW          '����FLG�iL2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDL3CW = tXSDCW(i).WFSMPLIDL3CW    '�T���v��ID�iL3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDL3CW = tXSDCW(i).WFINDL3CW          '���FLG�iL3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESL3CW = tXSDCW(i).WFRESL3CW          '����FLG�iL3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDL4CW = tXSDCW(i).WFSMPLIDL4CW    '�T���v��ID�iL4)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDL4CW = tXSDCW(i).WFINDL4CW          '���FLG�iL4)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESL4CW = tXSDCW(i).WFRESL4CW          '����FLG�iL4)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDSCW = tXSDCW(i).WFSMPLIDDSCW    '�T���v��ID�iDS)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDSCW = tXSDCW(i).WFINDDSCW          '���FLG�iDS)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDSCW = tXSDCW(i).WFRESDSCW          '����FLG�iDS)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDZCW = tXSDCW(i).WFSMPLIDDZCW    '�T���v��ID�iDZ)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDZCW = tXSDCW(i).WFINDDZCW          '���FLG�iDZ)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDZCW = tXSDCW(i).WFRESDZCW          '����FLG�iDZ)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDSPCW = tXSDCW(i).WFSMPLIDSPCW    '�T���v��ID�iSP)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDSPCW = tXSDCW(i).WFINDSPCW          '���FLG�iSP)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESSPCW = tXSDCW(i).WFRESSPCW          '����FLG�iSP)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDO1CW = tXSDCW(i).WFSMPLIDDO1CW  '�T���v��ID�iDO1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDO1CW = tXSDCW(i).WFINDDO1CW        '���FLG�iDO1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDO1CW = tXSDCW(i).WFRESDO1CW        '����FLG�iDO1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDO2CW = tXSDCW(i).WFSMPLIDDO2CW  '�T���v��ID�iDO2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDO2CW = tXSDCW(i).WFINDDO2CW        '���FLG�iDO2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDO2CW = tXSDCW(i).WFRESDO2CW        '����FLG�iDO2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDO3CW = tXSDCW(i).WFSMPLIDDO3CW  '�T���v��ID�iDO3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDO3CW = tXSDCW(i).WFINDDO3CW        '���FLG�iDO3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDO3CW = tXSDCW(i).WFRESDO3CW        '����FLG�iDO3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDOT1CW = tXSDCW(i).WFSMPLIDOT1CW  '�T���v��ID�iOT1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDOT1CW = tXSDCW(i).WFINDOT1CW        '���FLG�iOT1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESOT1CW = tXSDCW(i).WFRESOT1CW        '����FLG�iOT1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDOT2CW = tXSDCW(i).WFSMPLIDOT2CW  '�T���v��ID�iOT2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDOT2CW = tXSDCW(i).WFINDOT2CW        '���FLG�iOT2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESOT2CW = tXSDCW(i).WFRESOT2CW        '����FLG�iOT2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDAOICW = tXSDCW(i).WFSMPLIDAOICW  '�T���v��ID�iAOi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDAOICW = tXSDCW(i).WFINDAOICW        '���FLG�iAOi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESAOICW = tXSDCW(i).WFRESAOICW        '����FLG�iAOi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SMPLNUMCW = tXSDCW(i).SMPLNUMCW          '�T���v������
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SMPLPATCW = tXSDCW(i).SMPLPATCW          '�T���v���p�^�[��
                        tmpXSDCW(UBound(tmpXSDCW) - 1).LIVKCW = tXSDCW(i).LIVKCW                '�����敪
                        tmpXSDCW(UBound(tmpXSDCW) - 1).TSTAFFCW = tXSDCW(i).TSTAFFCW            '�o�^�Ј�ID
                        tmpXSDCW(UBound(tmpXSDCW) - 1).TDAYCW = tXSDCW(i).TDAYCW                '�o�^���t
                        tmpXSDCW(UBound(tmpXSDCW) - 1).KSTAFFCW = tXSDCW(i).KSTAFFCW            '�X�V�Ј�ID
                        tmpXSDCW(UBound(tmpXSDCW) - 1).KDAYCW = tXSDCW(i).KDAYCW                '�X�V���t
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SNDKCW = tXSDCW(i).SNDKCW                '���M�t���O
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SNDDAYCW = tXSDCW(i).SNDDAYCW            '���M���t
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDGDCW = tXSDCW(i).WFSMPLIDGDCW    '�T���v��ID�iGD)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDGDCW = tXSDCW(i).WFINDGDCW          '���FLG�iGD)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESGDCW = tXSDCW(i).WFRESGDCW          '����FLG�iGD)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFHSGDCW = tXSDCW(i).WFHSGDCW            '�ۏ�FLG�iGD)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDB1CW = tXSDCW(i).EPSMPLIDB1CW    '�T���v��ID(BMD1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDB1CW = tXSDCW(i).EPINDB1CW          '���FLG(BMD1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESB1CW = tXSDCW(i).EPRESB1CW          '����FLG(BMD1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDB2CW = tXSDCW(i).EPSMPLIDB2CW    '�T���v��ID(BMD2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDB2CW = tXSDCW(i).EPINDB2CW          '���FLG(BMD2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESB2CW = tXSDCW(i).EPRESB2CW          '����FLG(BMD2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDB3CW = tXSDCW(i).EPSMPLIDB3CW    '�T���v��ID(BMD3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDB3CW = tXSDCW(i).EPINDB3CW          '���FLG(BMD3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESB3CW = tXSDCW(i).EPRESB3CW          '����FLG(BMD3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDL1CW = tXSDCW(i).EPSMPLIDL1CW    '�T���v��ID(OSF1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDL1CW = tXSDCW(i).EPINDL1CW          '���FLG(OSF1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESL1CW = tXSDCW(i).EPRESL1CW          '����FLG(OSF1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDL2CW = tXSDCW(i).EPSMPLIDL2CW    '�T���v��ID(OSF2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDL2CW = tXSDCW(i).EPINDL2CW          '���FLG(OSF2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESL2CW = tXSDCW(i).EPRESL2CW          '����FLG(OSF2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDL3CW = tXSDCW(i).EPSMPLIDL3CW    '�T���v��ID(OSF3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDL3CW = tXSDCW(i).EPINDL3CW          '���FLG(OSF3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESL3CW = tXSDCW(i).EPRESL3CW          '����FLG(OSF3)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
                    Else
                        '' WF���ύX�ŕ�������đ��������̃��R�[�h��⊮����
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SMPKBNCW = tXSDCW(i).SMPKBNCW            '�T���v���敪
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SMPKBNCW = "D"                           '�T���v���敪

                        tmpXSDCW(UBound(tmpXSDCW) - 1).TBKBNCW = tXSDCW(i).TBKBNCW              'T/B�敪
                        tmpXSDCW(UBound(tmpXSDCW) - 1).XTALCW = tXSDCW(i).XTALCW                '�����ԍ�
                        tmpXSDCW(UBound(tmpXSDCW) - 1).HINBCW = tmpXSDCA(j).HINBCA              '�i��(XSDCA)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).REVNUMCW = tmpXSDCA(j).REVNUMCA          '���i�ԍ������ԍ�(XSDCA)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).FACTORYCW = tmpXSDCA(j).FACTORYCA        '�H��(XSDCA)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).OPECW = tmpXSDCA(j).OPECA                '���Ə���(XSDCA)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).KTKBNCW = tXSDCW(i).KTKBNCW              '�m��敪

                        tmpXSDCW(UBound(tmpXSDCW) - 1).REPSMPLIDCW = "                "         '��\�T���v��ID
                        tmpXSDCW(UBound(tmpXSDCW) - 1).INPOSCW = tmpXSDCA(j).INPOSCA            '�������ʒu
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SMCRYNUMCW = fncBsmpID(tmpXSDCW, UBound(tmpXSDCW) - 1)     '�T���v���u���b�NID

                        tmpXSDCW(UBound(tmpXSDCW) - 1).SXLIDCW = Left(tmpXSDCA(j).CRYNUMCA, 10) & GetWafPos(tmpXSDCA(j).INPOSCA) 'SXLID
                        tmpXSDCA(j).SXLIDCA = tmpXSDCW(UBound(tmpXSDCW) - 1).SXLIDCW 'SXLID�ۑ�
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDRSCW = ""                       '�T���v��ID(Rs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDRS1CW = "0"                     '����T���v��ID1(Rs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDRS2CW = "0"                     '����T���v��ID2(Rs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDRSCW = "0"                         '���FLG�iRs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESRS1CW = "0"                        '����FLG1�iRs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESRS2CW = "0"                        '����FLG2�iRs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDOICW = ""                       '�T���v��ID�iOi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDOICW = "0"                         '���FLG�iOi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESOICW = "0"                         '����FLG�iOi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDB1CW = ""                       '�T���v��ID�iB1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDB1CW = "0"                         '���FLG�iB1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESB1CW = "0"                         '����FLG�iB1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDB2CW = ""                       '�T���v��ID�iB2�j
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDB2CW = "0"                         '���FLG�iB2�j
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESB2CW = "0"                         '����FLG�iB2�j
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDB3CW = ""                       '�T���v��ID�iB3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDB3CW = "0"                         '���FLG�iB3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESB3CW = "0"                         '����FLG�iB3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDL1CW = ""                       '�T���v��ID�iL1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDL1CW = "0"                         '���FLG�iL1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESL1CW = "0"                         '����FLG�iL1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDL2CW = ""                       '�T���v��ID�iL2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDL2CW = "0"                         '���FLG�iL2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESL2CW = "0"                         '����FLG�iL2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDL3CW = ""                       '�T���v��ID�iL3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDL3CW = "0"                         '���FLG�iL3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESL3CW = "0"                         '����FLG�iL3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDL4CW = ""                       '�T���v��ID�iL4)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDL4CW = "0"                         '���FLG�iL4)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESL4CW = "0"                         '����FLG�iL4)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDSCW = ""                       '�T���v��ID�iDS)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDSCW = "0"                         '���FLG�iDS)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDSCW = "0"                         '����FLG�iDS)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDZCW = ""                       '�T���v��ID�iDZ)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDZCW = "0"                          '���FLG�iDZ)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDZCW = "0"                          '����FLG�iDZ)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDSPCW = ""                        '�T���v��ID�iSP)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDSPCW = "0"                          '���FLG�iSP)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESSPCW = "0"                          '����FLG�iSP)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDO1CW = ""                       '�T���v��ID�iDO1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDO1CW = "0"                         '���FLG�iDO1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDO1CW = "0"                         '����FLG�iDO1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDO2CW = ""                       '�T���v��ID�iDO2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDO2CW = "0"                         '���FLG�iDO2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDO2CW = "0"                         '����FLG�iDO2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDO3CW = ""                       '�T���v��ID�iDO3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDO3CW = "0"                         '���FLG�iDO3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDO3CW = "0"                         '����FLG�iDO3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDOT1CW = ""                       '�T���v��ID�iOT1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDOT1CW = "0"                         '���FLG�iOT1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESOT1CW = "0"                         '����FLG�iOT1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDOT2CW = ""                       '�T���v��ID�iOT2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDOT2CW = "0"                         '���FLG�iOT2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESOT2CW = "0"                         '����FLG�iOT2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDAOICW = ""                       '�T���v��ID�iAOi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDAOICW = "0"                         '���FLG�iAOi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESAOICW = "0"                         '����FLG�iAOi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SMPLNUMCW = "0"                           '�T���v������
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SMPLPATCW = ""                           '�T���v���p�^�[��
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDGDCW = ""                        '�T���v��ID�iGD)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDGDCW = "0"                          '���FLG�iGD)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESGDCW = "0"                          '����FLG�iGD)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFHSGDCW = "0"                           '�ۏ�FLG�iGD)

                        tmpXSDCW(UBound(tmpXSDCW) - 1).LIVKCW = tXSDCW(i).LIVKCW                '�����敪
                        tmpXSDCW(UBound(tmpXSDCW) - 1).TSTAFFCW = tXSDCW(i).TSTAFFCW            '�o�^�Ј�ID
                        tmpXSDCW(UBound(tmpXSDCW) - 1).TDAYCW = tXSDCW(i).TDAYCW                '�o�^���t
                        tmpXSDCW(UBound(tmpXSDCW) - 1).KSTAFFCW = tXSDCW(i).KSTAFFCW            '�X�V�Ј�ID
                        tmpXSDCW(UBound(tmpXSDCW) - 1).KDAYCW = tXSDCW(i).KDAYCW                '�X�V���t
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SNDKCW = tXSDCW(i).SNDKCW                '���M�t���O
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SNDDAYCW = tXSDCW(i).SNDDAYCW            '���M���t

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDB1CW = ""                        '�T���v��ID(BMD1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDB1CW = "0"                          '���FLG(BMD1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESB1CW = "0"                          '����FLG(BMD1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDB2CW = ""                        '�T���v��ID(BMD2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDB2CW = "0"                          '���FLG(BMD2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESB2CW = "0"                          '����FLG(BMD2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDB3CW = ""                        '�T���v��ID(BMD3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDB3CW = "0"                          '���FLG(BMD3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESB3CW = "0"                          '����FLG(BMD3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDL1CW = ""                        '�T���v��ID(OSF1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDL1CW = "0"                          '���FLG(OSF1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESL1CW = "0"                          '����FLG(OSF1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDL2CW = ""                        '�T���v��ID(OSF2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDL2CW = "0"                          '���FLG(OSF2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESL2CW = "0"                          '����FLG(OSF2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDL3CW = ""                        '�T���v��ID(OSF3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDL3CW = "0"                          '���FLG(OSF3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESL3CW = "0"                          '����FLG(OSF3)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

                    End If



                    '' Bottom�ʒu�ݒ� --------------------------------------------------------------------------------------
                    tmpXSDCW(UBound(tmpXSDCW)).SXLIDCW = tmpXSDCW(UBound(tmpXSDCW) - 1).SXLIDCW 'SXLID
'                    tmpXSDCW(UBound(tmpXSDCW)).SMPKBNCW = tXSDCW(i + 1).SMPKBNCW            '�T���v���敪
                    tmpXSDCW(UBound(tmpXSDCW)).TBKBNCW = tXSDCW(i + 1).TBKBNCW              'T/B�敪
                    tmpXSDCW(UBound(tmpXSDCW)).XTALCW = tXSDCW(i).XTALCW                    '�����ԍ�
                    tmpXSDCW(UBound(tmpXSDCW)).HINBCW = tmpXSDCA(j).HINBCA                  '�i��(XSDCA)
                    tmpXSDCW(UBound(tmpXSDCW)).REVNUMCW = tmpXSDCA(j).REVNUMCA              '���i�ԍ������ԍ�(XSDCA)
                    tmpXSDCW(UBound(tmpXSDCW)).FACTORYCW = tmpXSDCA(j).FACTORYCA            '�H��(XSDCA)
                    tmpXSDCW(UBound(tmpXSDCW)).OPECW = tmpXSDCA(j).OPECA                    '���Ə���(XSDCA)
                    tmpXSDCW(UBound(tmpXSDCW)).KTKBNCW = tXSDCW(i + 1).KTKBNCW              '�m��敪

                    ''�t���O������
                    liEndSxlFLG = 0
                    If j + 1 <= UBound(tmpXSDCA) Then
                        If tXSDCW(i).INPOSCW <= tmpXSDCA(j + 1).INPOSCA _
                            And tXSDCW(i + 1).INPOSCW > tmpXSDCA(j + 1).INPOSCA Then
                            '' SXL�̒��i�ǉ��f�[�^���j�̏ꍇ�t���O�𗧂Ă�
                            liEndSxlFLG = 1
                            '' XSCA�̃f�[�^�����̃��R�[�h������ꍇ
                            tmpXSDCW(UBound(tmpXSDCW)).REPSMPLIDCW = "                "         '��\�T���v��ID
                            tmpXSDCW(UBound(tmpXSDCW)).INPOSCW = tmpXSDCA(j + 1).INPOSCA        '�������ʒu
                            tmpXSDCW(UBound(tmpXSDCW)).SMCRYNUMCW = fncBsmpID(tmpXSDCW, UBound(tmpXSDCW))        '�T���v���u���b�NID
                            tmpXSDCW(UBound(tmpXSDCW)).SMPKBNCW = "U"                           '�T���v���敪

                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDRSCW = ""                        '�T���v��ID(Rs)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDRS1CW = "0"                      '����T���v��ID1(Rs)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDRS2CW = "0"                      '����T���v��ID2(Rs)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDRSCW = "0"                          '���FLG�iRs)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESRS1CW = "0"                         '����FLG1�iRs)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESRS2CW = "0"                         '����FLG2�iRs)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDOICW = ""                        '�T���v��ID�iOi)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDOICW = "0"                          '���FLG�iOi)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESOICW = "0"                          '����FLG�iOi)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDB1CW = ""                        '�T���v��ID�iB1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDB1CW = "0"                          '���FLG�iB1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESB1CW = "0"                          '����FLG�iB1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDB2CW = ""                        '�T���v��ID�iB2�j
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDB2CW = "0"                          '���FLG�iB2�j
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESB2CW = "0"                          '����FLG�iB2�j
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDB3CW = ""                        '�T���v��ID�iB3)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDB3CW = "0"                          '���FLG�iB3)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESB3CW = "0"                          '����FLG�iB3)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDL1CW = ""                        '�T���v��ID�iL1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDL1CW = "0"                          '���FLG�iL1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESL1CW = "0"                          '����FLG�iL1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDL2CW = ""                        '�T���v��ID�iL2)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDL2CW = "0"                          '���FLG�iL2)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESL2CW = "0"                          '����FLG�iL2)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDL3CW = ""                        '�T���v��ID�iL3)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDL3CW = "0"                          '���FLG�iL3)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESL3CW = "0"                          '����FLG�iL3)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDL4CW = ""                        '�T���v��ID�iL4)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDL4CW = "0"                          '���FLG�iL4)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESL4CW = "0"                          '����FLG�iL4)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDSCW = ""                        '�T���v��ID�iDS)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDDSCW = "0"                          '���FLG�iDS)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESDSCW = "0"                          '����FLG�iDS)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDZCW = ""                        '�T���v��ID�iDZ)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDDZCW = "0"          '���FLG�iDZ)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESDZCW = "0"          '����FLG�iDZ)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDSPCW = ""        '�T���v��ID�iSP)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDSPCW = "0"          '���FLG�iSP)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESSPCW = "0"          '����FLG�iSP)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDO1CW = ""       '�T���v��ID�iDO1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDDO1CW = "0"         '���FLG�iDO1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESDO1CW = "0"         '����FLG�iDO1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDO2CW = ""       '�T���v��ID�iDO2)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDDO2CW = "0"         '���FLG�iDO2)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESDO2CW = "0"         '����FLG�iDO2)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDO3CW = ""       '�T���v��ID�iDO3)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDDO3CW = "0"         '���FLG�iDO3)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESDO3CW = "0"         '����FLG�iDO3)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDOT1CW = ""       '�T���v��ID�iOT1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDOT1CW = "0"         '���FLG�iOT1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESOT1CW = "0"         '����FLG�iOT1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDOT2CW = ""       '�T���v��ID�iOT2)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDOT2CW = "0"         '���FLG�iOT2)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESOT2CW = "0"         '����FLG�iOT2)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDAOICW = ""       '�T���v��ID�iAOi)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDAOICW = "0"         '���FLG�iAOi)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESAOICW = "0"         '����FLG�iAOi)
                            tmpXSDCW(UBound(tmpXSDCW)).SMPLNUMCW = "0"           '�T���v������
                            tmpXSDCW(UBound(tmpXSDCW)).SMPLPATCW = ""           '�T���v���p�^�[��
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDGDCW = ""        '�T���v��ID�iGD)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDGDCW = "0"          '���FLG�iGD)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESGDCW = "0"          '����FLG�iGD)
                            tmpXSDCW(UBound(tmpXSDCW)).WFHSGDCW = "0"           '�ۏ�FLG�iGD)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                            tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDB1CW = ""                    '�T���v��ID(BMD1)
                            tmpXSDCW(UBound(tmpXSDCW)).EPINDB1CW = "0"                      '���FLG(BMD1)
                            tmpXSDCW(UBound(tmpXSDCW)).EPRESB1CW = "0"                      '����FLG(BMD1)
                            tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDB2CW = ""                    '�T���v��ID(BMD2)
                            tmpXSDCW(UBound(tmpXSDCW)).EPINDB2CW = "0"                      '���FLG(BMD2)
                            tmpXSDCW(UBound(tmpXSDCW)).EPRESB2CW = "0"                      '����FLG(BMD2)
                            tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDB3CW = ""                    '�T���v��ID(BMD3)
                            tmpXSDCW(UBound(tmpXSDCW)).EPINDB3CW = "0"                      '���FLG(BMD3)
                            tmpXSDCW(UBound(tmpXSDCW)).EPRESB3CW = "0"                      '����FLG(BMD3)
                            tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDL1CW = ""                    '�T���v��ID(OSF1)
                            tmpXSDCW(UBound(tmpXSDCW)).EPINDL1CW = "0"                      '���FLG(OSF1)
                            tmpXSDCW(UBound(tmpXSDCW)).EPRESL1CW = "0"                      '����FLG(OSF1)
                            tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDL2CW = ""                    '�T���v��ID(OSF2)
                            tmpXSDCW(UBound(tmpXSDCW)).EPINDL2CW = "0"                      '���FLG(OSF2)
                            tmpXSDCW(UBound(tmpXSDCW)).EPRESL2CW = "0"                      '����FLG(OSF2)
                            tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDL3CW = ""                    '�T���v��ID(OSF3)
                            tmpXSDCW(UBound(tmpXSDCW)).EPINDL3CW = "0"                      '���FLG(OSF3)
                            tmpXSDCW(UBound(tmpXSDCW)).EPRESL3CW = "0"                      '����FLG(OSF3)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
                        End If
                    End If
                    ''
                    If liEndSxlFLG = 0 Then
                        '' XSCA�̃f�[�^������ōŌ�̏ꍇ
                        tmpXSDCW(UBound(tmpXSDCW)).REPSMPLIDCW = tXSDCW(i + 1).REPSMPLIDCW         '��\�T���v��ID
                        tmpXSDCW(UBound(tmpXSDCW)).INPOSCW = tXSDCW(i + 1).INPOSCW              '�������ʒu
                        tmpXSDCW(UBound(tmpXSDCW)).SMCRYNUMCW = tXSDCW(i + 1).SMCRYNUMCW        '�T���v���u���b�NID
                        tmpXSDCW(UBound(tmpXSDCW)).SMPKBNCW = tXSDCW(i + 1).SMPKBNCW            '�T���v���敪

                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDRSCW = tXSDCW(i + 1).WFSMPLIDRSCW   '�T���v��ID(Rs)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDRS1CW = tXSDCW(i + 1).WFSMPLIDRS1CW '����T���v��ID1(Rs)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDRS2CW = tXSDCW(i + 1).WFSMPLIDRS2CW '����T���v��ID2(Rs)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDRSCW = tXSDCW(i + 1).WFINDRSCW         '���FLG�iRs)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESRS1CW = tXSDCW(i + 1).WFRESRS1CW       '����FLG1�iRs)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESRS2CW = tXSDCW(i + 1).WFRESRS2CW       '����FLG2�iRs)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDOICW = tXSDCW(i + 1).WFSMPLIDOICW   '�T���v��ID�iOi)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDOICW = tXSDCW(i + 1).WFINDOICW         '���FLG�iOi)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESOICW = tXSDCW(i + 1).WFRESOICW         '����FLG�iOi)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDB1CW = tXSDCW(i + 1).WFSMPLIDB1CW   '�T���v��ID�iB1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDB1CW = tXSDCW(i + 1).WFINDB1CW         '���FLG�iB1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESB1CW = tXSDCW(i + 1).WFRESB1CW         '����FLG�iB1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDB2CW = tXSDCW(i + 1).WFSMPLIDB2CW   '�T���v��ID�iB2�j
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDB2CW = tXSDCW(i + 1).WFINDB2CW         '���FLG�iB2�j
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESB2CW = tXSDCW(i + 1).WFRESB2CW         '����FLG�iB2�j
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDB3CW = tXSDCW(i + 1).WFSMPLIDB3CW   '�T���v��ID�iB3)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDB3CW = tXSDCW(i + 1).WFINDB3CW         '���FLG�iB3)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESB3CW = tXSDCW(i + 1).WFRESB3CW         '����FLG�iB3)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDL1CW = tXSDCW(i + 1).WFSMPLIDL1CW   '�T���v��ID�iL1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDL1CW = tXSDCW(i + 1).WFINDL1CW         '���FLG�iL1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESL1CW = tXSDCW(i + 1).WFRESL1CW         '����FLG�iL1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDL2CW = tXSDCW(i + 1).WFSMPLIDL2CW   '�T���v��ID�iL2)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDL2CW = tXSDCW(i + 1).WFINDL2CW         '���FLG�iL2)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESL2CW = tXSDCW(i + 1).WFRESL2CW         '����FLG�iL2)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDL3CW = tXSDCW(i + 1).WFSMPLIDL3CW   '�T���v��ID�iL3)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDL3CW = tXSDCW(i + 1).WFINDL3CW         '���FLG�iL3)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESL3CW = tXSDCW(i + 1).WFRESL3CW         '����FLG�iL3)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDL4CW = tXSDCW(i + 1).WFSMPLIDL4CW   '�T���v��ID�iL4)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDL4CW = tXSDCW(i + 1).WFINDL4CW         '���FLG�iL4)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESL4CW = tXSDCW(i + 1).WFRESL4CW         '����FLG�iL4)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDSCW = tXSDCW(i + 1).WFSMPLIDDSCW   '�T���v��ID�iDS)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDDSCW = tXSDCW(i + 1).WFINDDSCW         '���FLG�iDS)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESDSCW = tXSDCW(i + 1).WFRESDSCW         '����FLG�iDS)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDZCW = tXSDCW(i + 1).WFSMPLIDDZCW   '�T���v��ID�iDZ)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDDZCW = tXSDCW(i + 1).WFINDDZCW         '���FLG�iDZ)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESDZCW = tXSDCW(i + 1).WFRESDZCW         '����FLG�iDZ)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDSPCW = tXSDCW(i + 1).WFSMPLIDSPCW   '�T���v��ID�iSP)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDSPCW = tXSDCW(i + 1).WFINDSPCW         '���FLG�iSP)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESSPCW = tXSDCW(i + 1).WFRESSPCW         '����FLG�iSP)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDO1CW = tXSDCW(i + 1).WFSMPLIDDO1CW '�T���v��ID�iDO1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDDO1CW = tXSDCW(i + 1).WFINDDO1CW       '���FLG�iDO1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESDO1CW = tXSDCW(i + 1).WFRESDO1CW       '����FLG�iDO1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDO2CW = tXSDCW(i + 1).WFSMPLIDDO2CW '�T���v��ID�iDO2)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDDO2CW = tXSDCW(i + 1).WFINDDO2CW       '���FLG�iDO2)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESDO2CW = tXSDCW(i + 1).WFRESDO2CW       '����FLG�iDO2)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDO3CW = tXSDCW(i + 1).WFSMPLIDDO3CW '�T���v��ID�iDO3)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDDO3CW = tXSDCW(i + 1).WFINDDO3CW       '���FLG�iDO3)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESDO3CW = tXSDCW(i + 1).WFRESDO3CW       '����FLG�iDO3)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDOT1CW = tXSDCW(i + 1).WFSMPLIDOT1CW '�T���v��ID�iOT1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDOT1CW = tXSDCW(i + 1).WFINDOT1CW       '���FLG�iOT1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESOT1CW = tXSDCW(i + 1).WFRESOT1CW       '����FLG�iOT1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDOT2CW = tXSDCW(i + 1).WFSMPLIDOT2CW '�T���v��ID�iOT2)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDOT2CW = tXSDCW(i + 1).WFINDOT2CW       '���FLG�iOT2)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESOT2CW = tXSDCW(i + 1).WFRESOT2CW       '����FLG�iOT2)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDAOICW = tXSDCW(i + 1).WFSMPLIDAOICW '�T���v��ID�iAOi)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDAOICW = tXSDCW(i + 1).WFINDAOICW       '���FLG�iAOi)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESAOICW = tXSDCW(i + 1).WFRESAOICW       '����FLG�iAOi)
                        tmpXSDCW(UBound(tmpXSDCW)).SMPLNUMCW = tXSDCW(i + 1).SMPLNUMCW         '�T���v������
                        tmpXSDCW(UBound(tmpXSDCW)).SMPLPATCW = tXSDCW(i + 1).SMPLPATCW         '�T���v���p�^�[��
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDGDCW = tXSDCW(i + 1).WFSMPLIDGDCW   '�T���v��ID�iGD)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDGDCW = tXSDCW(i + 1).WFINDGDCW         '���FLG�iGD)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESGDCW = tXSDCW(i + 1).WFRESGDCW         '����FLG�iGD)
                        tmpXSDCW(UBound(tmpXSDCW)).WFHSGDCW = tXSDCW(i + 1).WFHSGDCW           '�ۏ�FLG�iGD)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                        tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDB1CW = tXSDCW(i + 1).EPSMPLIDB1CW    '�T���v��ID(BMD1)
                        tmpXSDCW(UBound(tmpXSDCW)).EPINDB1CW = tXSDCW(i + 1).EPINDB1CW          '���FLG(BMD1)
                        tmpXSDCW(UBound(tmpXSDCW)).EPRESB1CW = tXSDCW(i + 1).EPRESB1CW          '����FLG(BMD1)
                        tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDB2CW = tXSDCW(i + 1).EPSMPLIDB2CW    '�T���v��ID(BMD2)
                        tmpXSDCW(UBound(tmpXSDCW)).EPINDB2CW = tXSDCW(i + 1).EPINDB2CW          '���FLG(BMD2)
                        tmpXSDCW(UBound(tmpXSDCW)).EPRESB2CW = tXSDCW(i + 1).EPRESB2CW          '����FLG(BMD2)
                        tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDB3CW = tXSDCW(i + 1).EPSMPLIDB3CW    '�T���v��ID(BMD3)
                        tmpXSDCW(UBound(tmpXSDCW)).EPINDB3CW = tXSDCW(i + 1).EPINDB3CW          '���FLG(BMD3)
                        tmpXSDCW(UBound(tmpXSDCW)).EPRESB3CW = tXSDCW(i + 1).EPRESB3CW          '����FLG(BMD3)
                        tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDL1CW = tXSDCW(i + 1).EPSMPLIDL1CW    '�T���v��ID(OSF1)
                        tmpXSDCW(UBound(tmpXSDCW)).EPINDL1CW = tXSDCW(i + 1).EPINDL1CW          '���FLG(OSF1)
                        tmpXSDCW(UBound(tmpXSDCW)).EPRESL1CW = tXSDCW(i + 1).EPRESL1CW          '����FLG(OSF1)
                        tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDL2CW = tXSDCW(i + 1).EPSMPLIDL2CW    '�T���v��ID(OSF2)
                        tmpXSDCW(UBound(tmpXSDCW)).EPINDL2CW = tXSDCW(i + 1).EPINDL2CW          '���FLG(OSF2)
                        tmpXSDCW(UBound(tmpXSDCW)).EPRESL2CW = tXSDCW(i + 1).EPRESL2CW          '����FLG(OSF2)
                        tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDL3CW = tXSDCW(i + 1).EPSMPLIDL3CW    '�T���v��ID(OSF3)
                        tmpXSDCW(UBound(tmpXSDCW)).EPINDL3CW = tXSDCW(i + 1).EPINDL3CW          '���FLG(OSF3)
                        tmpXSDCW(UBound(tmpXSDCW)).EPRESL3CW = tXSDCW(i + 1).EPRESL3CW          '����FLG(OSF3)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

                    End If
                    tmpXSDCW(UBound(tmpXSDCW)).LIVKCW = tXSDCW(i).LIVKCW                '�����敪
                    tmpXSDCW(UBound(tmpXSDCW)).TSTAFFCW = tXSDCW(i).TSTAFFCW            '�o�^�Ј�ID
                    tmpXSDCW(UBound(tmpXSDCW)).TDAYCW = tXSDCW(i).TDAYCW                '�o�^���t
                    tmpXSDCW(UBound(tmpXSDCW)).KSTAFFCW = tXSDCW(i).KSTAFFCW            '�X�V�Ј�ID
                    tmpXSDCW(UBound(tmpXSDCW)).KDAYCW = tXSDCW(i).KDAYCW                '�X�V���t
                    tmpXSDCW(UBound(tmpXSDCW)).SNDKCW = tXSDCW(i).SNDKCW                '���M�t���O
                    tmpXSDCW(UBound(tmpXSDCW)).SNDDAYCW = tXSDCW(i).SNDDAYCW            '���M���t

                End If
            Next j
'            For j = 1 To UBound(tmpXSDCA)
'                ReDim Preserve tSXLID(j)
'                tSXLID(j).LOTID = tmpXSDCA(j).CRYNUMCA
'                tSXLID(j).SXLID = tmpXSDCA(j).SXLIDCA
'                tSXLID(j).IngotPos = tmpXSDCA(j).INPOSCA
'            Next j
        Else
            '�⊮����Ă��Ȃ��ꍇ�A���̂܂܂̃f�[�^���g�p
            If liUpdateFLG = 0 Then
                ReDim Preserve tmpXSDCW(UBound(tmpXSDCW) + 2) As typ_XSDCW
                'TOP�����̐ݒ�
                tmpXSDCW(UBound(tmpXSDCW) - 1) = tXSDCW(i)
                'Bottom�����̐ݒ�
                tmpXSDCW(UBound(tmpXSDCW)) = tXSDCW(i + 1)
            End If
        End If
    Next i

    ''���[�N�̈�̃f�[�^�𔽉f������
    ReDim tXSDCW(UBound(tmpXSDCW)) As typ_XSDCW
    For i = 0 To UBound(tmpXSDCW)
        tXSDCW(i) = tmpXSDCW(i)
    Next i

    DBDRV_GetXSDCWUpdate = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_GetXSDCWUpdate = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit

End Function

'�T�v      :SXL�Ǘ��̑}��
'���Ұ��@�@:�ϐ���      ,IO ,�^               ,����
'      �@�@:SXL   �@�@�@,I  ,typ_TBCME042   �@,SXL�Ǘ�
'      �@�@:�߂�l      ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :DBDRV_SXL_UpdIns�Ɉڍs����\��
'����      :2001/07/12  �쐬 ���{
'           2006/01/20 SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs SMP�ΐ�
'           s_cmzcDBdriverCOM_SQL.DBDRV_SXL_INS���ڐA

Private Function DBDRV_SXL_INS_CB(SXL() As typ_TBCME042) As FUNCTION_RETURN

    Dim sql             As String
    Dim i               As Long
    Dim rs              As OraDynaset   'RecordSet
    Dim liRecCnt        As Long
    Dim lsMotoHinban    As String      '���i��
    Dim iLoopBkHinGet   As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_SXL_INS_CB"

    DBDRV_SXL_INS_CB = FUNCTION_RETURN_SUCCESS

    For i = 1 To UBound(SXL)
        If SXL(i).LENGTH > 0 Then
            'E042�̎��́A�����ԍ��ƌ������J�n�ʒu�Ō��Ă������A
            'XSDCB�ɕύX�ɔ���SXLID�Ō�������悤�ɕς���
            sql = "select count(XTALCB) cnt from XSDCB where SXLIDCB='" & SXL(i).SXLID & "'"
            ''�f�[�^�𒊏o����
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
            liRecCnt = CLng(rs("CNT"))
            rs.Close

            '���i�Ԏ擾
            lsMotoHinban = ""
            For iLoopBkHinGet = 0 To Kihon.CNTHINOLD - 1
                If (CInt(HinOld(iLoopBkHinGet).INPOSCA) <= CInt(SXL(0).INGOTPOS)) And (CInt(SXL(0).INGOTPOS) <= CInt(HinOld(iLoopBkHinGet).INPOSCA) + CInt(HinOld(iLoopBkHinGet).GNLCA)) Then
                     lsMotoHinban = HinOld(iLoopBkHinGet).HINBCA
                     Exit For
                End If
            Next
            If lsMotoHinban = "" Then '�����Y��HINOLD�����������玩���̕i�Ԃ����i�ԂƂ���
                lsMotoHinban = SXL(0).hinban
            End If

            '�f�[�^���Ȃ��ꍇ��Insert�A�������ꍇ��Update�ɂ���
            With SXL(i)

                If liRecCnt = 0 Then
                    sql = ""
                    sql = sql & " INSERT INTO XSDCB"
                    sql = sql & " ("
                    sql = sql & "   XTALCB"                         ' �����ԍ�
                    sql = sql & "  ,INPOSCB"                        ' �������J�n�ʒu
                    sql = sql & "  ,RLENCB"                         ' ����
                    sql = sql & "  ,SXLIDCB"                        ' SXLID
                    sql = sql & "  ,GNWKNTCB"                       ' ���ݍH��
                    sql = sql & "  ,NEWKNTCB"                       ' �ŏI�ʉߍH��
                    sql = sql & "  ,LIVKCB"                         ' �폜�敪
                    sql = sql & "  ,LSTCCB"                         ' �ŏI��ԋ敪
                    sql = sql & "  ,SHOLDCLSCB"                     ' �z�[���h�敪
                    sql = sql & "  ,HINBCB"                         ' �i��
                    sql = sql & "  ,REVNUMCB"                       ' ���i�ԍ������ԍ�
                    sql = sql & "  ,FACTORYCB"                      ' �H��
                    sql = sql & "  ,OPECB"                          ' ���Ə���
                    sql = sql & "  ,FURYCCB"                        ' �s�Ǘ��R
                    sql = sql & "  ,MAICB"                          ' ����
                    sql = sql & "  ,TDAYCB"                         ' �o�^���t
                    sql = sql & "  ,KDAYCB"                         ' �X�V���t
                    sql = sql & "  ,SNDKCB"                         ' ���M�t���O
                    sql = sql & "  ,WSRMAICB"                       ' WS��㖇��
                    sql = sql & "  ,WSNMAICB"                       ' WS��򌇗�����
                    sql = sql & "  ,WFCMAICB"                       ' �������
                    sql = sql & "  ,SXLRMAICB"                      ' SXL�w��(�Ǖi)
                    sql = sql & "  ,WFCNMAICB"                      ' WFC����������
                    sql = sql & "  ,SXLEMAICB"                      ' SXL�m�薇��
                    sql = sql & "  ,SRMAICB"                        ' �T���v�����w������
                    sql = sql & "  ,SNMAICB"                        ' �T���v�����w���s�ǖ���
                    sql = sql & "  ,STMAICB"                        ' �T���v������
                    sql = sql & "  ,FURIMAICB"                      ' �U�֖���
                    sql = sql & "  ,XTWORKCB"                       ' �����H��
                    sql = sql & "  ,WFWORKCB"                       ' �E�F�[�n����
                    sql = sql & "  ,LUFRCCB"                        ' �i��R�[�h
                    sql = sql & "  ,LUFRBCB"                        ' �i��敪
                    sql = sql & "  ,LDERCCB"                        ' �i���R�[�h
                    sql = sql & "  ,HOLDCCB"                        ' �z�[���h�R�[�h
                    sql = sql & "  ,HOLDBCB"                        ' �z�[���h�敪
                    sql = sql & "  ,EXKUBCB"                        ' ��O�敪
                    sql = sql & "  ,HENPKCB"                        ' �ԕi�敪
                    sql = sql & "  ,KANKCB"                         ' �����敪
                    sql = sql & "  ,NFCB"                           ' ���ɋ敪
                    sql = sql & "  ,SAKJCB"                         ' �폜�敪
                    sql = sql & "  ,SUMITCB"                        ' SUMIT���M�t���O
                    sql = sql & " )"
                    sql = sql & " VALUES"
                    sql = sql & " ("
                    sql = sql & "   '" & .CRYNUM & "'"              ' �����ԍ�
                    sql = sql & "  ," & .INGOTPOS & ""              ' �������J�n�ʒu
                    sql = sql & "  ," & .LENGTH & ""                ' ����
                    sql = sql & "  ,'" & .SXLID & "'"               ' SXLID
                    sql = sql & "  ,'" & .NOWPROC & "'"             ' ���ݍH��
                    sql = sql & "  ,'" & .LPKRPROCCD & "'"          ' �ŏI�ʉߍH��
                    sql = sql & "  ,'" & .DELCLS & "'"              ' �폜�敪
                    sql = sql & "  ,'" & .LSTATCLS & "'"            ' �ŏI��ԋ敪
                    sql = sql & "  ,'" & .HOLDCLS & "'"             ' �z�[���h�敪
                    sql = sql & "  ,'" & .hinban & "'"              ' �i��
                    sql = sql & "  ," & .REVNUM & ""                ' ���i�ԍ������ԍ�
                    sql = sql & "  ,'" & .factory & "'"             ' �H��
                    sql = sql & "  ,'" & .opecond & "'"             ' ���Ə���
                    sql = sql & "  ,'" & .BDCAUS & "'"              ' �s�Ǘ��R
                    sql = sql & "  ," & .Count & ""                 ' ����
                    sql = sql & "  ,sysdate"                        ' �o�^���t
                    sql = sql & "  ,sysdate"                        ' �X�V���t
                    sql = sql & "  ,'0'"                            ' ���M�t���O
                    sql = sql & "  ,'0'"                            ' WS��㖇��
                    sql = sql & "  ,'0'"                            ' WS��򌇗�����
                    sql = sql & "  ,'0'"                            ' �������
                    sql = sql & "  ,'0'"                            ' SXL�w��(�Ǖi)
                    sql = sql & "  ,'0'"                            ' WFC����������
                    sql = sql & "  ,'0'"                            ' SXL�m�薇��
                    sql = sql & "  ,'0'"                            ' �T���v�����w������
                    sql = sql & "  ,'0'"                            ' �T���v�����w���s�ǖ���
                    sql = sql & "  ,'0'"                            ' �T���v������
                    sql = sql & "  ,'0'"                            ' �U�֖���
                    sql = sql & "  ,'42'"                           ' �����H��
                    sql = sql & "  ,'  '"                           ' �E�F�[�n����
                    sql = sql & "  ,'   '"                          ' �i��R�[�h
                    sql = sql & "  ,' '"                            ' �i��敪
                    sql = sql & "  ,'   '"                          ' �i���R�[�h
                    sql = sql & "  ,'   '"                          ' �z�[���h�R�[�h
                    sql = sql & "  ,'0'"                            ' �z�[���h�敪
                    sql = sql & "  ,' '"                            ' ��O�敪
                    sql = sql & "  ,' '"                            ' �ԕi�敪
                    sql = sql & "  ,'0'"                            ' �����敪
                    sql = sql & "  ,'0'"                            ' ���ɋ敪
                    sql = sql & "  ,'0'"                            ' �폜�敪
                    sql = sql & "  ,'0'"                            ' SUMIT���M�t���O
                    sql = sql & " )"
                Else
                    sql = ""
                    sql = sql & " UPDATE XSDCB"
                    sql = sql & " SET XTALCB   = '" & .CRYNUM & "'"
                    sql = sql & "  ,INPOSCB    = " & .INGOTPOS & ""
                    sql = sql & "  ,RLENCB     = " & .LENGTH & ""
                    sql = sql & "  ,SXLIDCB    = '" & .SXLID & "'"
                    sql = sql & "  ,GNWKNTCB   = '" & .NOWPROC & "'"
                    sql = sql & "  ,NEWKNTCB   = '" & .LPKRPROCCD & "'"
                    sql = sql & "  ,LIVKCB     = '" & .DELCLS & "'"
                    sql = sql & "  ,LSTCCB     = '" & .LSTATCLS & "'"
                    sql = sql & "  ,SHOLDCLSCB = '" & .HOLDCLS & "'"
                    sql = sql & "  ,HINBCB     = '" & .hinban & "'"
                    sql = sql & "  ,REVNUMCB   = " & .REVNUM & ""
                    sql = sql & "  ,FACTORYCB  = '" & .factory & "'"
                    sql = sql & "  ,OPECB      = '" & .opecond & "'"
                    sql = sql & "  ,FURYCCB    = '" & .BDCAUS & "'"
                    sql = sql & "  ,MAICB      = " & .Count & ""
                    sql = sql & "  ,TDAYCB     = sysdate"
                    sql = sql & "  ,KDAYCB     = sysdate"
                    sql = sql & "  ,SNDKCB     = '0'"
                    sql = sql & "  ,SNDAYCB    = sysdate"
                    sql = sql & " where SXLIDCB='" & SXL(i).SXLID & "'"
                End If
            End With
            '' WriteDBLog sql
            If OraDB.ExecuteSQL(sql) <= 0 Then
                DBDRV_SXL_INS_CB = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        End If
    Next i

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_SXL_INS_CB = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'��ۯ�����يǗ�(XSDCS)������ۯ��I���ʒu���擾�@08/07/10 ooba
Public Function GetCSpos(sBlkId As String, iPos As Integer) As Integer
    Dim sql         As String
    Dim rs          As OraDynaset
    
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function GetCSpos"
    
    GetCSpos = iPos
    
    sql = "select INPOSCS from XSDCS "
    sql = sql & "where CRYNUMCS = '" & sBlkId & "' "
    sql = sql & "and TBKBNCS = 'B' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    If rs.RecordCount <> 1 Then
        GoTo proc_exit
    End If

    If IsNull(rs("INPOSCS")) = False Then GetCSpos = rs("INPOSCS")
    rs.Close
    
proc_exit:
    gErr.Pop
    Exit Function

proc_err:
    Resume proc_exit
    
End Function

'�T�v      :WFϯ��(TBCMY011)�o�^(�֘A��ۯ��ύX��)
'���Ұ��@�@:�ϐ���       ,IO ,�^                ,����
'      �@�@:sSxlid       ,I  ,String            ,SXLID
'      �@�@:sqlWhere     ,I  ,String            ,������
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@ ,�������݂̐���
'����      :
'����      :08/01/28 ooba
Public Function DBDRV_KanrenBlkMap(sSXLID As String, sqlWhere As String) As FUNCTION_RETURN
    
    Dim sql     As String
    
    '�װ����ׂ̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function DBDRV_KanrenBlkMap"

    DBDRV_KanrenBlkMap = FUNCTION_RETURN_FAILURE
    
    '�������Ȃ�
    If Trim(sqlWhere) = "" Then GoTo proc_exit
    
    'SXLID�X�V
    sql = "UPDATE TBCMY011 "
    sql = sql & "SET MSXLID = '" & sSXLID & "' "
    sql = sql & sqlWhere
    
    If OraDB.ExecuteSQL(sql) < 0 Then
        '�X�V���s
        GoTo proc_exit
    End If
    
    DBDRV_KanrenBlkMap = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �װ�����
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function

'�T�v      :�֘A��ۯ��R�t�R��(TBCMY023)�o�^
'���Ұ��@�@:�ϐ���           ,IO ,�^                ,����
'      �@�@:tKanrenDisp()    ,I  ,typ_KanrenDisp    ,�֘A��ۯ��ꗗ
'      �@�@:�߂�l           ,O  ,FUNCTION_RETURN�@ ,�������݂̐���
'����      :
'����      :08/01/28 ooba
Public Function DBDRV_KanrenBlk(tKanrenDisp() As typ_KanrenDisp) As FUNCTION_RETURN

    Dim sql             As String
    Dim i               As Long
    Dim rs              As OraDynaset
    Dim lRecCnt         As Long             'ں��ސ�
    Dim iTrnCnt         As Integer          '������
    Dim bSaveFlg        As Boolean          '�o�^�L��
    Dim KanrenData()    As typ_TBCMY023     '�֘A��ۯ��R�t�R���ް�
    
    '�װ����ׂ̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function DBDRV_KanrenBlk"

    DBDRV_KanrenBlk = FUNCTION_RETURN_FAILURE

    '�����񐔎擾
    sql = "SELECT NVL(MAX(TRANCNT),0) MAXCNT FROM TBCMY023"
    sql = sql & " WHERE CRYNUM = '" & tKanrenDisp(1).CRYNUM & "'"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    If rs.RecordCount = 0 Then
        iTrnCnt = 1
    Else
        iTrnCnt = rs("MAXCNT") + 1          '������(�ő�) + 1
    End If
    rs.Close

    lRecCnt = 0             '�o�^ں��ސ�
    
    '�֘A��ۯ��R���ް����
    For i = 1 To UBound(tKanrenDisp)
        lRecCnt = lRecCnt + 1
        ReDim Preserve KanrenData(lRecCnt)
        With KanrenData(lRecCnt)
            .CRYNUM = tKanrenDisp(i).CRYNUM     '�����ԍ�
            .TRANCNT = iTrnCnt                  '������
            .BLOCKID = tKanrenDisp(i).BLOCKID   '��ۯ�ID
            .PROCCAT = "D"                      '�����敪(D:�R��)
            .TXID = "TX879I"                    '��ݻ޸���ID
        End With
    Next i
    
    '�֘A��ۯ��R�t�ް����
    For i = 1 To UBound(tKanrenDisp)
        bSaveFlg = False
        '�֘A��ۯ��̐擪��ۯ�(�����O)
        If i = 1 Then
            If tKanrenDisp(i).KANREN = 0 Then
                iTrnCnt = iTrnCnt + 1       '������+1
                bSaveFlg = True
            End If
        Else
            '�O����ۯ��Ɗ֘A��ۯ�
            If tKanrenDisp(i - 1).KANREN = 0 Then bSaveFlg = True
            '�֘A��ۯ��̐擪��ۯ�(������)
            If tKanrenDisp(i - 1).KANREN = 1 And tKanrenDisp(i).KANREN = 0 Then
                iTrnCnt = iTrnCnt + 1       '������+1
                bSaveFlg = True
            End If
        End If
        
        If bSaveFlg Then
            lRecCnt = lRecCnt + 1
            ReDim Preserve KanrenData(lRecCnt)
            With KanrenData(lRecCnt)
                .CRYNUM = tKanrenDisp(i).CRYNUM     '�����ԍ�
                .TRANCNT = iTrnCnt                  '������
                .BLOCKID = tKanrenDisp(i).BLOCKID   '��ۯ�ID
                .PROCCAT = "C"                      '�����敪(C:�t�ւ�)
                .TXID = "TX879I"                    '��ݻ޸���ID
            End With
        End If
    Next i
    
    '�֘A��ۯ��R�t�R��(TBCMY023)�ɓo�^
    For i = 1 To UBound(KanrenData)
        With KanrenData(i)
            sql = "INSERT INTO TBCMY023"
            sql = sql & " (CRYNUM,"
            sql = sql & " TRANCNT,"
            sql = sql & " BLOCKID,"
            sql = sql & " PROCCAT,"
            sql = sql & " TXID,"
            sql = sql & " REGDATE,"
            sql = sql & " SUMITFLAG,"
            sql = sql & " SUMITSND,"
            sql = sql & " SSENDNO,"
            sql = sql & " SENDFLAG,"
            sql = sql & " SENDDATE,"
            sql = sql & " PLANTCAT)"
            sql = sql & " VALUES"
            sql = sql & " ('" & .CRYNUM & "',"      '�����ԍ�
            sql = sql & .TRANCNT & ","              '������
            sql = sql & " '" & .BLOCKID & "',"      '��ۯ�ID
            sql = sql & " '" & .PROCCAT & "',"      '�����敪
            sql = sql & " '" & .TXID & "',"         '��ݻ޸���ID
            sql = sql & " SYSDATE,"                 '�o�^���t
            sql = sql & " '0',"                     'SUMIT���M�׸�
            sql = sql & " NULL,"                    'SUMIT���M���t
            sql = sql & " NULL,"                    '���M���A��
            sql = sql & " '0',"                     '���M�׸�
            sql = sql & " NULL,"                    '���M���t
            sql = sql & " '" & sCmbMukesaki & "')"  '����
        End With

        '�o�^���s
        If OraDB.ExecuteSQL(sql) <= 0 Then
            GoTo proc_exit
        End If
    Next i
    
    DBDRV_KanrenBlk = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �װ�����
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function

'�T�v      :�֘A��ۯ��R�t�R��(TBCMY023)�o�^
'���Ұ��@�@:�ϐ���      ,IO ,�^               ,����
'      �@�@:sCrynum     ,I  ,String         �@,�����ԍ�
'      �@�@:sKblockid() ,I  ,String         �@,�֘A��ۯ�
'      �@�@:iSpos       ,I  ,Integer        �@,�������J�n�ʒu
'      �@�@:iEpos       ,I  ,Integer        �@,�������I���ʒu
'      �@�@:�߂�l      ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :
'����      :07/08/06 ooba
Public Function DBDRV_KanrenBlk_BK(sCryNum As String, sKblockid() As String, _
                                iSpos As Integer, iEpos As Integer) As FUNCTION_RETURN

    Dim sql             As String
    Dim i, j            As Long
    Dim rs              As OraDynaset
    Dim lRecCnt         As Long             'ں��ސ�
    Dim sLotid          As String           '��ۯ�ID(WFϯ��)
    Dim sSXLID          As String           'SXLID(WFϯ��)
    Dim KanrenData()    As typ_TBCMY023     '�֘A��ۯ��R�t�R���ް�
    Dim bCutFlg         As Boolean          '�֘A��ۯ��R�؂��׸�
    Dim iTrnCnt         As Integer          '������


    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function DBDRV_KanrenBlk_BK"

    DBDRV_KanrenBlk_BK = FUNCTION_RETURN_FAILURE

    '�����񐔎擾
    sql = "SELECT NVL(MAX(TRANCNT),0) MAXCNT FROM TBCMY023"
    sql = sql & " WHERE CRYNUM = '" & sCryNum & "'"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    If rs.RecordCount = 0 Then
        iTrnCnt = 1
    Else
        iTrnCnt = rs("MAXCNT") + 1          '������(�ő�) + 1
    End If
    rs.Close


    lRecCnt = 0             '�o�^ں��ސ�
    bCutFlg = False         '�֘A��ۯ��R�؂��׸�(False:�R�؂薳)

    '�֘A��ۯ��R���ް����
    For i = 1 To UBound(sKblockid)
        lRecCnt = lRecCnt + 1
        ReDim Preserve KanrenData(lRecCnt)
        With KanrenData(lRecCnt)
            .CRYNUM = sCryNum               '�����ԍ�
            .TRANCNT = iTrnCnt              '������
            .BLOCKID = sKblockid(i)         '��ۯ�ID
            .PROCCAT = "D"                  '�����敪(D:�R��)
            .TXID = "TX879I"                '��ݻ޸���ID
        End With
    Next i


    'WFϯ�߂����ۯ�ID,SXLID���擾
    sql = "SELECT LOTID, MSXLID FROM TBCMY011"
    sql = sql & " WHERE LOTID LIKE '" & Left(sCryNum, 9) & "%'"
    sql = sql & " AND (WFSTA = '0' OR WFSTA = '1')"
    sql = sql & " AND RITOP_POS > " & iSpos
    sql = sql & " AND RITOP_POS <= " & iEpos
    sql = sql & " AND MSXLID IS NOT NULL"
    sql = sql & " GROUP BY LOTID, MSXLID"
    sql = sql & " ORDER BY LOTID, MAX(BLOCKSEQ)"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    '�ް��Ȃ�
    If rs.RecordCount = 0 Then
        rs.Close
        GoTo proc_exit
    End If

    '�֘A��ۯ��R�t�ް����
    For i = 1 To rs.RecordCount
        If i > 1 Then
            '����ۯ��œ���SXL(�֘A��ۯ���)
            If sLotid <> rs("LOTID") And sSXLID = rs("MSXLID") Then
                '�֘A��ۯ�(��)
                If KanrenData(lRecCnt).BLOCKID <> sLotid Then
                    iTrnCnt = iTrnCnt + 1       '������
                    lRecCnt = lRecCnt + 1
                    ReDim Preserve KanrenData(lRecCnt)
                    With KanrenData(lRecCnt)
                        .CRYNUM = sCryNum               '�����ԍ�
                        .TRANCNT = iTrnCnt              '������
                        .BLOCKID = sLotid               '��ۯ�ID
                        .PROCCAT = "C"                  '�����敪(C:�t�ւ�)
                        .TXID = "TX879I"                '��ݻ޸���ID
                    End With
                End If
                '�֘A��ۯ�(��)
                lRecCnt = lRecCnt + 1
                ReDim Preserve KanrenData(lRecCnt)
                With KanrenData(lRecCnt)
                    .CRYNUM = sCryNum                   '�����ԍ�
                    .TRANCNT = iTrnCnt                  '������
                    .BLOCKID = rs("LOTID")              '��ۯ�ID
                    .PROCCAT = "C"                      '�����敪(C:�t�ւ�)
                    .TXID = "TX879I"                    '��ݻ޸���ID
                End With

            '����ۯ��ŕ�SXL(�֘A��ۯ��~)
            ElseIf sLotid <> rs("LOTID") And sSXLID <> rs("MSXLID") Then
                bCutFlg = True          '�֘A��ۯ��R�؂��׸�(True:�R�؂�L)
            End If
        End If
        sLotid = rs("LOTID")        '��ۯ�ID
        sSXLID = rs("MSXLID")       'SXLID
        rs.MoveNext
    Next i
    rs.Close


    '�֘A��ۯ��R�؂肪���������ꍇ�A�֘A��ۯ��R�t�R��(TBCMY023)�ɓo�^
    If bCutFlg Then
        For i = 1 To UBound(KanrenData)
            With KanrenData(i)
                sql = "INSERT INTO TBCMY023"
                sql = sql & " (CRYNUM,"
                sql = sql & " TRANCNT,"
                sql = sql & " BLOCKID,"
                sql = sql & " PROCCAT,"
                sql = sql & " TXID,"
                sql = sql & " REGDATE,"
                sql = sql & " SUMITFLAG,"               '07/12/21 ooba
                sql = sql & " SUMITSND,"                '07/12/21 ooba
                sql = sql & " SSENDNO,"                 '07/12/21 ooba
                sql = sql & " SENDFLAG,"

                ' 2007/09/03 SPK Tsutsumi Add Start
                sql = sql & " SENDDATE,"
                sql = sql & " PLANTCAT)"
'                sql = sql & " SENDDATE)"
                ' 2007/09/03 SPK Tsutsumi Add End

                sql = sql & " VALUES"
                sql = sql & " ('" & .CRYNUM & "',"      '�����ԍ�
                sql = sql & .TRANCNT & ","              '������
                sql = sql & " '" & .BLOCKID & "',"      '��ۯ�ID
                sql = sql & " '" & .PROCCAT & "',"      '�����敪
                sql = sql & " '" & .TXID & "',"         '��ݻ޸���ID
                sql = sql & " SYSDATE,"                 '�o�^���t
                sql = sql & " '0',"                     'SUMIT���M�׸�  07/12/21 ooba
                sql = sql & " NULL,"                    'SUMIT���M���t  07/12/21 ooba
                sql = sql & " NULL,"                    '���M���A��  07/12/21 ooba
                sql = sql & " '0',"                     '���M�׸�

                ' 2007/09/03 SPK Tsutsumi Add Start
                sql = sql & " NULL,"                    '���M���t
                sql = sql & " '" & sCmbMukesaki & "')"  '����
'                sql = sql & " NULL)"                    '���M���t
                ' 2007/09/03 SPK Tsutsumi Add End
            End With

            '�o�^���s
            If OraDB.ExecuteSQL(sql) <= 0 Then
                GoTo proc_exit
            End If
        Next i
    End If

    DBDRV_KanrenBlk_BK = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit

End Function

'��--- 2010/01/20 SIRD�Ή� SPK habuki ADD START
'---------------------------------------------------------------------------
'�T�v      :�����ԍ���V�����TBCMJ022���������ASIRD��������Ԃ�
'---------------------------------------------------------------------------
'���Ұ�    :�ϐ���      ,IO     ,�^                     ,����
'          :pCRYNUM     ,I  �@�@,String                 ,�����ԍ�
'          :pflgSird    ,O  �@�@,Boolean                ,SIRD�������̗L��(True:�L�AFalse�F��)
'          :pSMPLID     ,O  �@�@,String                 ,SIRD�������̑�\�����ID
'          :�߂�l      ,O      ,Boolean                ,[True:OK�^False:NG]
'---------------------------------------------------------------------------
Public Function fncGetSirdSample(ByVal pCRYNUM As String, ByRef pflgSird As Boolean, ByRef pSMPLID As String) As Boolean

    Dim sql         As String           'SQL�S��
    Dim rs          As OraDynaset       'RecordSet

    '--�װ����ׂ̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function fncGetSirdSample"
    
    '--������
    fncGetSirdSample = False: pflgSird = False: pSMPLID = ""
    Set rs = Nothing      'Oracle RecordSet Free

    '--SQL������
    sql = "select SMPLNO  "
    sql = sql & "from TBCMJ022 " & vbCrLf
    sql = sql & "where" & vbCrLf
    sql = sql & "     substr(CRYNUM,1,7) = '" & left(pCRYNUM, 7) & "'" & vbCrLf     '�����ԍ�(��7��)
    sql = sql & " and TRANCNT   = 0" & vbCrLf                                       '������

    '--�ް��𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY Or ORADYN_NOCACHE)
    If rs Is Nothing Then
        GoTo proc_exit
    End If
    
    '--���o���ʎQ��
    If Not (rs.EOF) Then
        '<< �ް��L�� >>
        rs.MoveFirst
        pflgSird = True                 '[SIRD�������L��]
        pSMPLID = rs("SMPLNO")          '[��\�����ID]
    End If
    
    fncGetSirdSample = True

proc_exit:
    '<< �I�� >>
    'Oracle RecordSet Free
    If Not (rs Is Nothing) Then
        rs.Close
    End If
    Set rs = Nothing
    
    gErr.Pop
    Exit Function

proc_err:
    '<< �װ����� >>
    'Oracle RecordSet Free
    If Not (rs Is Nothing) Then
        rs.Close
    End If
    Set rs = Nothing

    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    
    gErr.HandleError
    Resume proc_exit

End Function
'��--- 2010/01/20 SIRD�Ή� SPK habuki ADD END

'Add Start 2010/07/08 SMPK Nakamura
'---------------------------------------------------------------------------
'�T�v      :�u���b�NID���HINBCA���������A�i�ԏ���Ԃ�
'---------------------------------------------------------------------------
'���Ұ�    :�ϐ���      ,IO     ,�^                     ,����
'          :pCRYNUM     ,I  �@�@,String                 ,�u���b�NID
'          :psHinban    ,O  �@�@,String                 ,�i��
'          :�߂�l      ,O      ,Boolean                ,[True:OK�^False:NG]
'---------------------------------------------------------------------------
Public Function fncGetMultiHinban(ByVal pCRYNUM As String, ByRef psHinban As String) As Boolean

    Dim sql         As String           'SQL�S��
    Dim rs          As OraDynaset       'RecordSet
    Dim i           As Integer          '���[�v�J�E���g
    Dim sBlockId()    As String

    '--�װ����ׂ̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function fncGetMultiHinban"
    
    sBlockId = Split(pCRYNUM, Chr(9))
    
    '--������
    fncGetMultiHinban = False
    Set rs = Nothing      'Oracle RecordSet Free

    '--SQL������
    sql = "select distinct HINBCA from ("
    sql = sql & "select HINBCA "
    sql = sql & "from XSDCA " & vbCrLf
    sql = sql & "where" & vbCrLf
    sql = sql & "     CRYNUMCA in ( "
    If UBound(sBlockId) > 1 Then
        For i = 0 To UBound(sBlockId) - 1
            If InStr(sBlockId(i), "Wait") > 0 Then
                sql = sql & "'" & Trim(Mid(sBlockId(i), 1, InStr(sBlockId(i), "Wait") - 1)) & "' " & vbCrLf
            Else
                sql = sql & "'" & sBlockId(i) & "' " & vbCrLf
            End If
            If i <> UBound(sBlockId) - 1 Then sql = sql & ","
        Next i
    Else
        sql = sql & "'" & pCRYNUM & "' " & vbCrLf
    End If
    sql = sql & ") "
    sql = sql & "     and LIVKCA = '0' " & vbCrLf                '�����t���O
    sql = sql & "order by INPOSCA)" & vbCrLf

    '--�ް��𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Or rs.RecordCount = 0 Then
        GoTo proc_exit
    End If
    
    '--���o���ʎQ��
    psHinban = ""
    For i = 1 To rs.RecordCount
        psHinban = psHinban & rs("HINBCA") & vbTab  '�i��
        rs.MoveNext
    Next i
    
    fncGetMultiHinban = True

proc_exit:
    '<< �I�� >>
    'Oracle RecordSet Free
    If Not (rs Is Nothing) Then
        rs.Close
    End If
    Set rs = Nothing
    
    gErr.Pop
    Exit Function

proc_err:
    '<< �װ����� >>
    'Oracle RecordSet Free
    If Not (rs Is Nothing) Then
        rs.Close
    End If
    Set rs = Nothing

    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    
    gErr.HandleError
    Resume proc_exit

End Function
'Add End 2010/07/08 SMPK Nakamura

Attribute VB_Name = "s_cmbc015_SQL"
Option Explicit

'================================================
'�v���W�F�N�gcmbc015�pSQLbas
'2002/08 s_cmzcF_cmmc001a_SQL.bas���ړ�
'================================================

'(2002/07 DBDRV_GetTBCME018���ړ�)
'�t�B�[���h�������p
Dim fldNames() As String    '��rs�Ɋ܂܂��t�B�[���h���ێ��z��
Dim fldCnt As Integer       '��rs�Ɋ܂܂��t�B�[���h��


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME018�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME018 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCME018_SQL.bas���ړ�)
Public Function DBDRV_GetTBCME018(records() As typ_TBCME018, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, MCNO, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME018"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME018 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
'NULL�Ή� ----- START ----- 2003/12/10
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .HINBAN = rs("HINBAN")           ' �i��
            .mnorevno = rs("MNOREVNO")       ' ���i�ԍ������ԍ�
            .factory = rs("FACTORY")         ' �H��
            .opecond = rs("OPECOND")         ' ���Ə���
            .HMGSTRRNO = rs("HMGSTRRNO")     ' �i�Ǘ��d�l�o�^�˗��ԍ�
            .HMGSTFNO = rs("HMGSTFNO")       ' �i�Ǘ��Ј��m��
            .HMGSXSNO = rs("HMGSXSNO")       ' �i�Ǘ��r�w���i�ԍ�
            .HMGSXSNE = fncNullCheck(rs("HMGSXSNE"))       ' �i�Ǘ��r�w���i�ԍ��}��
            .CONFLAG = rs("CONFLAG")         ' �m�F�t���O
            .REINFLAG = rs("REINFLAG")       ' �ĕt�^�t���O
            .HSXTRWKB = rs("HSXTRWKB")       ' �i�r�w�����ۋ敪
            .HSXTYPE = rs("HSXTYPE")         ' �i�r�w�^�C�v
            .KSXTYPKW = rs("KSXTYPKW")       ' �i�r�w�^�C�v�������@
            .HSXDOP = rs("HSXDOP")           ' �i�r�w�h�[�p���g
            .HSXRMIN = fncNullCheck(rs("HSXRMIN"))         ' �i�r�w���R����
            .HSXRMAX = fncNullCheck(rs("HSXRMAX"))         ' �i�r�w���R���
            .HSXRSPOH = rs("HSXRSPOH")       ' �i�r�w���R����ʒu�Q��
            .HSXRSPOT = rs("HSXRSPOT")       ' �i�r�w���R����ʒu�Q�_
            .HSXRSPOI = rs("HSXRSPOI")       ' �i�r�w���R����ʒu�Q��
            .HSXRHWYT = rs("HSXRHWYT")       ' �i�r�w���R�ۏؕ��@�Q��
            .HSXRHWYS = rs("HSXRHWYS")       ' �i�r�w���R�ۏؕ��@�Q��
            .HSXRKWAY = rs("HSXRKWAY")       ' �i�r�w���R�������@
            .HSXRKHNM = rs("HSXRKHNM")       ' �i�r�w���R�����p�x�Q��
            .HSXRKHNI = rs("HSXRKHNI")       ' �i�r�w���R�����p�x�Q��
            .HSXRKHNH = rs("HSXRKHNH")       ' �i�r�w���R�����p�x�Q��
            .HSXRKHNS = rs("HSXRKHNS")       ' �i�r�w���R�����p�x�Q��
            .HSXRMCAL = rs("HSXRMCAL")       ' �i�r�w���R�ʓ��v�Z
            .HSXRMBNP = fncNullCheck(rs("HSXRMBNP"))       ' �i�r�w���R�ʓ����z
            .HSXRMCL2 = rs("HSXRMCL2")       ' �i�r�w���R�ʓ��v�Z�Q
            .HSXRMBP2 = fncNullCheck(rs("HSXRMBP2"))       ' �i�r�w���R�ʓ����z�Q
            .HSXRSDEV = fncNullCheck(rs("HSXRSDEV"))       ' �i�r�w���R�W���΍�
            .HSXRAMIN = fncNullCheck(rs("HSXRAMIN"))       ' �i�r�w���R���ω���
            .HSXRAMAX = fncNullCheck(rs("HSXRAMAX"))       ' �i�r�w���R���Ϗ��
            .HSXFORM = rs("HSXFORM")         ' �i�r�w�`��
            .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))       ' �i�r�w���a�P���S
            .HSXD1MIN = fncNullCheck(rs("HSXD1MIN"))       ' �i�r�w���a�P����
            .HSXD1MAX = fncNullCheck(rs("HSXD1MAX"))       ' �i�r�w���a�P���
            .HSXD2CEN = fncNullCheck(rs("HSXD2CEN"))       ' �i�r�w���a�Q���S
            .HSXD2MIN = fncNullCheck(rs("HSXD2MIN"))       ' �i�r�w���a�Q����
            .HSXD2MAX = fncNullCheck(rs("HSXD2MAX"))       ' �i�r�w���a�Q���
            .HSXCDIR = rs("HSXCDIR")         ' �i�r�w�����ʕ���
            .HSXCSCEN = fncNullCheck(rs("HSXCSCEN"))       ' �i�r�w�����ʌX���S
            .HSXCSMIN = fncNullCheck(rs("HSXCSMIN"))       ' �i�r�w�����ʌX����
            .HSXCSMAX = fncNullCheck(rs("HSXCSMAX"))       ' �i�r�w�����ʌX���
            .HSXCKWAY = rs("HSXCKWAY")       ' �i�r�w�����ʌ������@
            .HSXCKHNM = rs("HSXCKHNM")       ' �i�r�w�����ʌ����p�x�Q��
            .HSXCKHNI = rs("HSXCKHNI")       ' �i�r�w�����ʌ����p�x�Q��
            .HSXCKHNH = rs("HSXCKHNH")       ' �i�r�w�����ʌ����p�x�Q��
            .HSXCKHNS = rs("HSXCKHNS")       ' �i�r�w�����ʌ����p�x�Q��
            .HSXCSDIR = rs("HSXCSDIR")       ' �i�r�w�����ʌX����
            .HSXCSDIS = rs("HSXCSDIS")       ' �i�r�w�����ʌX���ʎw��
            .HSXCTDIR = rs("HSXCTDIR")       ' �i�r�w�����ʌX�c����
            .HSXCTCEN = fncNullCheck(rs("HSXCTCEN"))       ' �i�r�w�����ʌX�c���S
            .HSXCTMIN = fncNullCheck(rs("HSXCTMIN"))       ' �i�r�w�����ʌX�c����
            .HSXCTMAX = fncNullCheck(rs("HSXCTMAX"))       ' �i�r�w�����ʌX�c���
            .HSXCYDIR = rs("HSXCYDIR")       ' �i�r�w�����ʌX������
            .HSXCYCEN = fncNullCheck(rs("HSXCYCEN"))       ' �i�r�w�����ʌX�����S
            .HSXCYMIN = fncNullCheck(rs("HSXCYMIN"))       ' �i�r�w�����ʌX������
            .HSXCYMAX = fncNullCheck(rs("HSXCYMAX"))       ' �i�r�w�����ʌX�����
            .HSXOF1PD = rs("HSXOF1PD")       ' �i�r�w�n�e�P�ʒu����
            .HSXOF1PN = fncNullCheck(rs("HSXOF1PN"))       ' �i�r�w�n�e�P�ʒu����
            .HSXOF1PX = fncNullCheck(rs("HSXOF1PX"))       ' �i�r�w�n�e�P�ʒu���
            .HSXOF1PW = rs("HSXOF1PW")       ' �i�r�w�n�e�P�ʒu�������@
            .HSXOF1LC = fncNullCheck(rs("HSXOF1LC"))       ' �i�r�w�n�e�P�����S
            .HSXOF1LN = fncNullCheck(rs("HSXOF1LN"))       ' �i�r�w�n�e�P������
            .HSXOF1LX = fncNullCheck(rs("HSXOF1LX"))       ' �i�r�w�n�e�P�����
            .HSXOF1DC = fncNullCheck(rs("HSXOF1DC"))       ' �i�r�w�n�e�P���a���S
            .HSXOF1DN = fncNullCheck(rs("HSXOF1DN"))       ' �i�r�w�n�e�P���a����
            .HSXOF1DX = fncNullCheck(rs("HSXOF1DX"))       ' �i�r�w�n�e�P���a���
            .HSXDFORM = rs("HSXDFORM")       ' �i�r�w�a�`��
            .HSXDPDRC = rs("HSXDPDRC")       ' �i�r�w�a�ʒu����
            .HSXDPACN = fncNullCheck(rs("HSXDPACN"))       ' �i�r�w�a�ʒu�p�x���S
            .HSXDPAMN = fncNullCheck(rs("HSXDPAMN"))       ' �i�r�w�a�ʒu�p�x����
            .HSXDPAMX = fncNullCheck(rs("HSXDPAMX"))       ' �i�r�w�a�ʒu�p�x���
            .HSXDPKWY = rs("HSXDPKWY")       ' �i�r�w�a�ʒu�������@
            .HSXDPDIR = rs("HSXDPDIR")       ' �i�r�w�a�ʒu����
            .HSXDPMIN = fncNullCheck(rs("HSXDPMIN"))       ' �i�r�w�a�ʒu����
            .HSXDPMAX = fncNullCheck(rs("HSXDPMAX"))       ' �i�r�w�a�ʒu���
            .HSXDWCEN = fncNullCheck(rs("HSXDWCEN"))       ' �i�r�w�a�В��S
            .HSXDWMIN = fncNullCheck(rs("HSXDWMIN"))       ' �i�r�w�a�Љ���
            .HSXDWMAX = fncNullCheck(rs("HSXDWMAX"))       ' �i�r�w�a�Џ��
            .HSXDDCEN = fncNullCheck(rs("HSXDDCEN"))       ' �i�r�w�a�[���S
            .HSXDDMIN = fncNullCheck(rs("HSXDDMIN"))       ' �i�r�w�a�[����
            .HSXDDMAX = fncNullCheck(rs("HSXDDMAX"))       ' �i�r�w�a�[���
            .HSXDACEN = fncNullCheck(rs("HSXDACEN"))       ' �i�r�w�a�p�x���S
            .HSXDAMIN = fncNullCheck(rs("HSXDAMIN"))       ' �i�r�w�a�p�x����
            .HSXDAMAX = fncNullCheck(rs("HSXDAMAX"))       ' �i�r�w�a�p�x���
            .MCNO = rs("MCNO")               ' �������Ɠ��������
            .IFKBN = rs("IFKBN")             ' �h�^�e�敪
            .SYORIKBN = rs("SYORIKBN")       ' �����敪
            .SPECRRNO = rs("SPECRRNO")       ' �d�l�o�^�˗��ԍ�
            .SXLMCNO = rs("SXLMCNO")         ' �r�w�k��������ԍ�
            .WFMCNO = rs("WFMCNO")           ' �v�e��������ԍ�
            .StaffID = rs("STAFFID")         ' �Ј�ID
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close
'NULL�Ή� -----  END  ----- 2003/12/10

    DBDRV_GetTBCME018 = FUNCTION_RETURN_SUCCESS
End Function




'�T�v      :�����̃t�B�[���h����fldNames()�z��Ɋ܂܂�Ă��邩�ǂ����̔���B
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :fldName       ,I  ,typ_TBCME018 ,���o���R�[�h
'          :�߂�l        ,O  ,Boolean      ,True:�݂�^False�F����
'����      :
'����      :2001/06/27�쐬�@�쑺  (2002/07 DBDRV_GetTBCME018���ړ�)

Private Function fldNameExist(fldName As String) As Boolean
    Dim sql         As String           'SQL�S��
    Dim i As Integer                    'ٰ�߶���


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc015_SQL.bas -- Function fldNameExist"

    fldNameExist = False                '�װ�ð���i�����l�j���
    
    For i = 1 To fldCnt                 '̨���ސ���ٰ��
        If fldName = fldNames(i) Then   '������̨���ޖ��ƈ�v������̂��������ꍇ
            fldNameExist = True         '����ð�����
            Exit For                    'ٰ�߂𔲂���
        End If
    Next
    

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
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

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
            .Crynum = rs("CRYNUM")           ' �����ԍ�
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

'�T�v      :�e�[�u���uTBCMH004�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMH004 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCMH004_SQL.bas���ړ�)
Public Function DBDRV_GetTBCMH004(records() As typ_TBCMH004, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, KRPROCCD, PROCCODE, LENGTOP, LENGTKDO, LENGTAIL, LENGFREE, DM1, DM2, DM3, WGHTTOP, WGHTTKDO," & _
              " WGHTTAIL, WGHTFREE, WGTOPCUT, UPWEIGHT, CHARGE, SEED, STATCLS, JDGECODE, PWTIME, ADDDPPOS, ADDDPCLS, ADDDPVAL," & _
              " ADDDPNAM, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMH004"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMH004 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .Crynum = rs("CRYNUM")           ' �����ԍ�
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
            .LENGTOP = rs("LENGTOP")         ' �����iTOP�j
            .LENGTKDO = rs("LENGTKDO")       ' �����i�����j
            .LENGTAIL = rs("LENGTAIL")       ' �����iTAIL�j
            .LENGFREE = rs("LENGFREE")       ' �t���[����
            .DM1 = rs("DM1")                 ' �������a�P
            .DM2 = rs("DM2")                 ' �������a�Q
            .DM3 = rs("DM3")                 ' �������a�R
            .WGHTTOP = rs("WGHTTOP")         ' �d�ʁiTOP�j
            .WGHTTKDO = rs("WGHTTKDO")       ' �d�ʁi�����j
            .WGHTTAIL = rs("WGHTTAIL")       ' �d�ʁiTAIL)
            .WGHTFREE = rs("WGHTFREE")       ' �d�ʁi�t���[�����j
            .WGTOPCUT = rs("WGTOPCUT")       ' �g�b�v�J�b�g�d��
            .UPWEIGHT = rs("UPWEIGHT")       ' ���グ�d��
            .CHARGE = rs("CHARGE")           ' �`���[�W��
            .SEED = rs("SEED")               ' �V�[�h
            .STATCLS = rs("STATCLS")         ' BOT�󋵋敪
            .JDGECODE = rs("JDGECODE")       ' ����R�[�h
            .PWTIME = rs("PWTIME")           ' �p���[����
            .ADDDPPOS = rs("ADDDPPOS")       ' �ǉ��h�[�v�ʒu
            .ADDDPCLS = rs("ADDDPCLS")       ' �ǉ��h�[�p���g���
            .ADDDPVAL = rs("ADDDPVAL")       ' �ǉ��h�[�v��
            .ADDDPNAM = rs("ADDDPNAM")       ' �ǉ��h�[�v��
            .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMH004 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMJ002�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMJ002 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺  (2002/07 s_cmzcTBCMJ002_SQL.bas���ړ�)
Public Function DBDRV_GetTBCMJ002(records() As typ_TBCMJ002, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, KRPROCCD, PROCCODE, HINBAN, REVNUM, FACTORY," & _
              " OPECOND, GOUKI, TYPE, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, EFEHS, RRG, JUDGDATA, TSTAFFID, REGDATE, KSTAFFID," & _
              " UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMJ002"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ002 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .Crynum = rs("CRYNUM")           ' �����ԍ�
            .POSITION = rs("POSITION")       ' �ʒu
            .SMPKBN = rs("SMPKBN")           ' �T���v���敪
            .TRANCOND = rs("TRANCOND")       ' ��������
            .TRANCNT = rs("TRANCNT")         ' ������
            .SMPLNO = rs("SMPLNO")           ' �T���v���m��
            .SMPLUMU = rs("SMPLUMU")         ' �T���v���L��
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
            .HINBAN = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .factory = rs("FACTORY")         ' �H��
            .opecond = rs("OPECOND")         ' ���Ə���
            .GOUKI = rs("GOUKI")             ' ���@
            .TYPE = rs("TYPE")               ' �^�C�v
            .MEAS1 = rs("MEAS1")             ' ����l�P
            .MEAS2 = rs("MEAS2")             ' ����l�Q
            .MEAS3 = rs("MEAS3")             ' ����l�R
            .MEAS4 = rs("MEAS4")             ' ����l�S
            .MEAS5 = rs("MEAS5")             ' ����l�T
            .EFEHS = rs("EFEHS")             ' �����ΐ�
            .RRG = rs("RRG")                 ' �q�q�f
            .JudgData = rs("JUDGDATA")       ' �����Ώےl
            .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMJ002 = FUNCTION_RETURN_SUCCESS
End Function



Attribute VB_Name = "SB_GetSiyou"
Option Explicit

'------------------------------------------------
' TBCME018�f�[�^�擾(����p)
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME018�v����w��i�Ԃ̃��R�[�h�𒊏o����B
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :tHin          ,I  ,tFullHinban                          :�i��
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :���o���R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN                      :���o�̐���
'����      :�����L�[�́A�HINBAN�+�uMNOREVNO�v+�uFACTORY�v+�uOPECOND�v�̕�����Ƃ���
'����      :2003/09/10 �V�K�쐬�@�V�X�e���u���C��

Public Function funGet_TBCME018(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmkc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL�S��
    Dim rs          As OraDynaset       'RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME018"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sql = sql & "HSXTYPE, HSXD1CEN, HSXCDIR, HSXRMIN, HSXRMAX, HSXRAMIN, HSXRAMAX, "
    sql = sql & "HSXRMCAL, HSXRMBNP, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS "
    sql = sql & "from TBCME018 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME018 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''���o���ʂ��i�[����
     With tGetRec
        .hin.hinban = rs("HINBAN")          ' �i��
        .hin.mnorevno = rs("MNOREVNO")      ' ���i�ԍ������ԍ�
        .hin.factory = rs("FACTORY")        ' �H��
        .hin.opecond = rs("OPECOND")        ' ���Ə���
        
        .HSXTYPE = rs("HSXTYPE")                    ' �i�r�w�^�C�v
        .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))    ' �i�r�w���a�P���S          2003/12/12 SystemBrain Null�Ή�
        .HSXCDIR = rs("HSXCDIR")                    ' �i�r�w�����ʕ���
        .HSXRMIN = fncNullCheck(rs("HSXRMIN"))      ' �i�r�w���R����          2003/12/12 SystemBrain Null�Ή�
        .HSXRMAX = fncNullCheck(rs("HSXRMAX"))      ' �i�r�w���R���          2003/12/12 SystemBrain Null�Ή�
        .HSXRAMIN = fncNullCheck(rs("HSXRAMIN"))    ' �i�r�w���R���ω���      2003/12/12 SystemBrain Null�Ή�
        .HSXRAMAX = fncNullCheck(rs("HSXRAMAX"))    ' �i�r�w���R���Ϗ��      2003/12/12 SystemBrain Null�Ή�
        .HSXRMCAL = rs("HSXRMCAL")                  ' �i�r�w���R�ʓ��v�Z
        .HSXRMBNP = fncNullCheck(rs("HSXRMBNP"))    ' �i�r�w���R�ʓ����z      2003/12/12 SystemBrain Null�Ή�
        .HSXRSPOH = rs("HSXRSPOH")                  ' �i�r�w���R����ʒu�Q��
        .HSXRSPOT = rs("HSXRSPOT")                  ' �i�r�w���R����ʒu�Q�_
        .HSXRSPOI = rs("HSXRSPOI")                  ' �i�r�w���R����ʒu�Q��
        .HSXRHWYT = rs("HSXRHWYT")                  ' �i�r�w���R�ۏؕ��@�Q��
        .HSXRHWYS = rs("HSXRHWYS")                  ' �i�r�w���R�ۏؕ��@�Q��
    End With
    Set rs = Nothing

    funGet_TBCME018 = FUNCTION_RETURN_SUCCESS
  

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
' TBCME019�f�[�^�擾(����p)
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME019�v����w��i�Ԃ̃��R�[�h�𒊏o����B
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :tHin          ,I  ,tFullHinban                          :�i��
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :���o���R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN                      :���o�̐���
'����      :�����L�[�́A�HINBAN�+�uMNOREVNO�v+�uFACTORY�v+�uOPECOND�v�̕�����Ƃ���
'����      :2003/09/10 �V�K�쐬�@�V�X�e���u���C��

Public Function funGet_TBCME019(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmkc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL�S��
    Dim rs          As OraDynaset       'RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME019"

    'HSXCNKHI�ǉ� 09/01/08 ooba
    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sql = sql & "HSXONMIN, HSXONMAX, HSXONAMN, HSXONAMX, HSXONMCL, HSXONMBP, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, "
    sql = sql & "HSXCNMIN, HSXCNMAX, HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKHI, "
    sql = sql & "HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT, HSXLTHWS "
    sql = sql & "from TBCME019 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME019 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''���o���ʂ��i�[����
     With tGetRec
        .hin.hinban = rs("HINBAN")         ' �i��
        .hin.mnorevno = rs("MNOREVNO")     ' ���i�ԍ������ԍ�
        .hin.factory = rs("FACTORY")       ' �H��
        .hin.opecond = rs("OPECOND")       ' ���Ə���
        
        .HSXONMIN = fncNullCheck(rs("HSXONMIN"))        ' �i�r�w�_�f�Z�x����            2003/12/12 SystemBrain Null�Ή�
        .HSXONMAX = fncNullCheck(rs("HSXONMAX"))        ' �i�r�w�_�f�Z�x���            2003/12/12 SystemBrain Null�Ή�
        .HSXONAMN = fncNullCheck(rs("HSXONAMN"))        ' �i�r�w�_�f�Z�x���ω���        2003/12/12 SystemBrain Null�Ή�
        .HSXONAMX = fncNullCheck(rs("HSXONAMX"))        ' �i�r�w�_�f�Z�x���Ϗ��        2003/12/12 SystemBrain Null�Ή�
        .HSXONMCL = rs("HSXONMCL")                      ' �i�r�w�_�f�Z�x�ʓ��v�Z
        .HSXONMBP = fncNullCheck(rs("HSXONMBP"))        ' �i�r�w�_�f�Z�x�ʓ����z        2003/12/12 SystemBrain Null�Ή�
        .HSXONSPH = rs("HSXONSPH")                      ' �i�r�w�_�f�Z�x����ʒu�Q��
        .HSXONSPT = rs("HSXONSPT")                      ' �i�r�w�_�f�Z�x����ʒu�Q�_
        .HSXONSPI = rs("HSXONSPI")                      ' �i�r�w�_�f�Z�x����ʒu�Q��
        .HSXONHWT = rs("HSXONHWT")                      ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
        .HSXONHWS = rs("HSXONHWS")                      ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
        
        .HSXCNMIN = fncNullCheck(rs("HSXCNMIN"))        ' �i�r�w�Y�f�Z�x����            2003/12/12 SystemBrain Null�Ή�
        .HSXCNMAX = fncNullCheck(rs("HSXCNMAX"))        ' �i�r�w�Y�f�Z�x���            2003/12/12 SystemBrain Null�Ή�
        .HSXCNSPH = rs("HSXCNSPH")                      ' �i�r�w�Y�f�Z�x����ʒu�Q��
        .HSXCNSPT = rs("HSXCNSPT")                      ' �i�r�w�Y�f�Z�x����ʒu�Q�_
        .HSXCNSPI = rs("HSXCNSPI")                      ' �i�r�w�Y�f�Z�x����ʒu�Q��
        .HSXCNHWT = rs("HSXCNHWT")                      ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
        .HSXCNHWS = rs("HSXCNHWS")                      ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
        .HSXCNKHI = rs("HSXCNKHI")                      ' �i�r�w�Y�f�Z�x�����p�x�Q�� 09/01/08 ooba
        
        .HSXLTMIN = fncNullCheck(rs("HSXLTMIN"))        ' �i�r�w�k�^�C������            2003/12/12 SystemBrain Null�Ή�
        .HSXLTMAX = fncNullCheck(rs("HSXLTMAX"))        ' �i�r�w�k�^�C�����            2003/12/12 SystemBrain Null�Ή�
        .HSXLTSPH = rs("HSXLTSPH")                      ' �i�r�w�k�^�C������ʒu�Q��
        .HSXLTSPT = rs("HSXLTSPT")                      ' �i�r�w�k�^�C������ʒu�Q�_
        .HSXLTSPI = rs("HSXLTSPI")                      ' �i�r�w�k�^�C������ʒu�Q��
        .HSXLTHWT = rs("HSXLTHWT")                      ' �i�r�w�k�^�C���ۏؕ��@�Q��
        .HSXLTHWS = rs("HSXLTHWS")                      ' �i�r�w�k�^�C���ۏؕ��@�Q��
    End With
    Set rs = Nothing

    funGet_TBCME019 = FUNCTION_RETURN_SUCCESS
  

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
' TBCME020�f�[�^�擾(����p)
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME020�v����w��i�Ԃ̃��R�[�h�𒊏o����B
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :tHin          ,I  ,tFullHinban                          :�i��
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :���o���R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN                      :���o�̐���
'����      :�����L�[�́A�HINBAN�+�uMNOREVNO�v+�uFACTORY�v+�uOPECOND�v�̕�����Ƃ���
'����      :2003/09/10 �V�K�쐬�@�V�X�e���u���C��

Public Function funGet_TBCME020(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmkc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL�S��
    Dim rs          As OraDynaset       'RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME020"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sql = sql & "HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1NS, "
    sql = sql & "HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS, HSXBM2NS, "
    sql = sql & "HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST, HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3NS, "
    sql = sql & "HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1NS, "
    sql = sql & "HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2NS, "
    sql = sql & "HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR, HSXOF3HT, HSXOF3HS, HSXOF3NS, "
    sql = sql & "HSXOF4AX, HSXOF4MX, HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4NS, "
    sql = sql & "HSXDENMX, HSXDENMN, HSXDENHT, HSXDENHS, HSXDENKU, "
    sql = sql & "HSXLDLMX, HSXLDLMN, HSXLDLHT, HSXLDLHS, HSXLDLKU, "
    sql = sql & "HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXDVDKU, "
    sql = sql & "HSXOSF1PTK, HSXOSF2PTK, HSXOSF3PTK, HSXOSF4PTK, HSXBMD1MBP, HSXBMD2MBP, HSXBMD3MBP "
    sql = sql & ", HSXGDPTK "   '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
    'Add Start 2011/02/01 SMPK Miyata
    sql = sql & ", HSXCPK, HSXCSZ, HSXCHT, HSXCHS "
    sql = sql & ", HSXCJPK, HSXCJNS, HSXCJHT, HSXCJHS "
    sql = sql & ", HSXCJLTPK, HSXCJLTNS, HSXCJLTHT, HSXCJLTHS "
    sql = sql & ", HSXCJ2PK, HSXCJ2NS, HSXCJ2HT, HSXCJ2HS "
    'Add End   2011/02/01 SMPK Miyata
    sql = sql & "from TBCME020 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME020 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
     
    ''���o���ʂ��i�[����
    With tGetRec
        .hin.hinban = rs("HINBAN")       ' �i��
        .hin.mnorevno = rs("MNOREVNO")   ' ���i�ԍ������ԍ�
        .hin.factory = rs("FACTORY")     ' �H��
        .hin.opecond = rs("OPECOND")     ' ���Ə���
        
        .HSXBM1AN = fncNullCheck(rs("HSXBM1AN"))        ' �i�r�w�a�l�c1���ω���     2003/12/12 SystemBrain Null�Ή�
        .HSXBM1AX = fncNullCheck(rs("HSXBM1AX"))        ' �i�r�w�a�l�c1���Ϗ��     2003/12/12 SystemBrain Null�Ή�
        .HSXBM1SH = rs("HSXBM1SH")                      ' �i�r�w�a�l�c1����ʒu�Q��
        .HSXBM1ST = rs("HSXBM1ST")                      ' �i�r�w�a�l�c1����ʒu�Q�_
        .HSXBM1SR = rs("HSXBM1SR")                      ' �i�r�w�a�l�c1����ʒu�Q��
        .HSXBM1HT = rs("HSXBM1HT")                      ' �i�r�w�a�l�c1�ۏؕ��@�Q��
        .HSXBM1HS = rs("HSXBM1HS")                      ' �i�r�w�a�l�c1�ۏؕ��@�Q��
        .HSXBM1NS = rs("HSXBM1NS")                      ' �i�r�w�a�l�c1�M�����@
        .HSXBM2AN = fncNullCheck(rs("HSXBM2AN"))        ' �i�r�w�a�l�c2���ω���     2003/12/12 SystemBrain Null�Ή�
        .HSXBM2AX = fncNullCheck(rs("HSXBM2AX"))        ' �i�r�w�a�l�c2���Ϗ��     2003/12/12 SystemBrain Null�Ή�
        .HSXBM2SH = rs("HSXBM2SH")                      ' �i�r�w�a�l�c2����ʒu�Q��
        .HSXBM2ST = rs("HSXBM2ST")                      ' �i�r�w�a�l�c2����ʒu�Q�_
        .HSXBM2SR = rs("HSXBM2SR")                      ' �i�r�w�a�l�c2����ʒu�Q��
        .HSXBM2HT = rs("HSXBM2HT")                      ' �i�r�w�a�l�c2�ۏؕ��@�Q��
        .HSXBM2HS = rs("HSXBM2HS")                      ' �i�r�w�a�l�c2�ۏؕ��@�Q��
        .HSXBM2NS = rs("HSXBM2NS")                      ' �i�r�w�a�l�c2�M�����@
        .HSXBM3AN = fncNullCheck(rs("HSXBM3AN"))        ' �i�r�w�a�l�c3���ω���     2003/12/12 SystemBrain Null�Ή�
        .HSXBM3AX = fncNullCheck(rs("HSXBM3AX"))        ' �i�r�w�a�l�c3���Ϗ��     2003/12/12 SystemBrain Null�Ή�
        .HSXBM3SH = rs("HSXBM3SH")                      ' �i�r�w�a�l�c3����ʒu�Q��
        .HSXBM3ST = rs("HSXBM3ST")                      ' �i�r�w�a�l�c3����ʒu�Q�_
        .HSXBM3SR = rs("HSXBM3SR")                      ' �i�r�w�a�l�c3����ʒu�Q��
        .HSXBM3HT = rs("HSXBM3HT")                      ' �i�r�w�a�l�c3�ۏؕ��@�Q��
        .HSXBM3HS = rs("HSXBM3HS")                      ' �i�r�w�a�l�c3�ۏؕ��@�Q��
        .HSXBM3NS = rs("HSXBM3NS")                      ' �i�r�w�a�l�c3�M�����@
        
        .HSXOF1AX = fncNullCheck(rs("HSXOF1AX"))        ' �i�r�w�n�r�e1���Ϗ��     2003/12/12 SystemBrain Null�Ή�
        .HSXOF1MX = fncNullCheck(rs("HSXOF1MX"))        ' �i�r�w�n�r�e1���         2003/12/12 SystemBrain Null�Ή�
        .HSXOF1SH = rs("HSXOF1SH")                      ' �i�r�w�n�r�e1����ʒu�Q��
        .HSXOF1ST = rs("HSXOF1ST")                      ' �i�r�w�n�r�e1����ʒu�Q�_
        .HSXOF1SR = rs("HSXOF1SR")                      ' �i�r�w�n�r�e1����ʒu�Q��
        .HSXOF1HT = rs("HSXOF1HT")                      ' �i�r�w�n�r�e1�ۏؕ��@�Q��
        .HSXOF1HS = rs("HSXOF1HS")                      ' �i�r�w�n�r�e1�ۏؕ��@�Q��
        .HSXOF1NS = rs("HSXOF1NS")                      ' �i�r�w�n�r�e1�M�����@
        .HSXOF2AX = fncNullCheck(rs("HSXOF2AX"))        ' �i�r�w�n�r�e2���Ϗ��     2003/12/12 SystemBrain Null�Ή�
        .HSXOF2MX = fncNullCheck(rs("HSXOF2MX"))        ' �i�r�w�n�r�e2���         2003/12/12 SystemBrain Null�Ή�
        .HSXOF2SH = rs("HSXOF2SH")                      ' �i�r�w�n�r�e2����ʒu�Q��
        .HSXOF2ST = rs("HSXOF2ST")                      ' �i�r�w�n�r�e2����ʒu�Q�_
        .HSXOF2SR = rs("HSXOF2SR")                      ' �i�r�w�n�r�e2����ʒu�Q��
        .HSXOF2HT = rs("HSXOF2HT")                      ' �i�r�w�n�r�e2�ۏؕ��@�Q��
        .HSXOF2HS = rs("HSXOF2HS")                      ' �i�r�w�n�r�e2�ۏؕ��@�Q��
        .HSXOF2NS = rs("HSXOF2NS")                      ' �i�r�w�n�r�e2�M�����@
        .HSXOF3AX = fncNullCheck(rs("HSXOF3AX"))        ' �i�r�w�n�r�e3���Ϗ��     2003/12/12 SystemBrain Null�Ή�
        .HSXOF3MX = fncNullCheck(rs("HSXOF3MX"))        ' �i�r�w�n�r�e3���         2003/12/12 SystemBrain Null�Ή�
        .HSXOF3SH = rs("HSXOF3SH")                      ' �i�r�w�n�r�e3����ʒu�Q��
        .HSXOF3ST = rs("HSXOF3ST")                      ' �i�r�w�n�r�e3����ʒu�Q�_
        .HSXOF3SR = rs("HSXOF3SR")                      ' �i�r�w�n�r�e3����ʒu�Q��
        .HSXOF3HT = rs("HSXOF3HT")                      ' �i�r�w�n�r�e3�ۏؕ��@�Q��
        .HSXOF3HS = rs("HSXOF3HS")                      ' �i�r�w�n�r�e3�ۏؕ��@�Q��
        .HSXOF3NS = rs("HSXOF3NS")                      ' �i�r�w�n�r�e3�M�����@
        .HSXOF4AX = fncNullCheck(rs("HSXOF4AX"))        ' �i�r�w�n�r�e4���Ϗ��     2003/12/12 SystemBrain Null�Ή�
        .HSXOF4MX = fncNullCheck(rs("HSXOF4MX"))        ' �i�r�w�n�r�e4���         2003/12/12 SystemBrain Null�Ή�
        .HSXOF4SH = rs("HSXOF4SH")                      ' �i�r�w�n�r�e4����ʒu�Q��
        .HSXOF4ST = rs("HSXOF4ST")                      ' �i�r�w�n�r�e4����ʒu�Q�_
        .HSXOF4SR = rs("HSXOF4SR")                      ' �i�r�w�n�r�e4����ʒu�Q��
        .HSXOF4HT = rs("HSXOF4HT")                      ' �i�r�w�n�r�e4�ۏؕ��@�Q��
        .HSXOF4HS = rs("HSXOF4HS")                      ' �i�r�w�n�r�e4�ۏؕ��@�Q��
        .HSXOF4NS = rs("HSXOF4NS")                      ' �i�r�w�n�r�e4�M�����@
        
        .HSXDENKU = rs("HSXDENKU")                      ' �i�r�w�c���������L��
        .HSXDENMX = fncNullCheck(rs("HSXDENMX"))        ' �i�r�w�c�������          2003/12/12 SystemBrain Null�Ή�
        .HSXDENMN = fncNullCheck(rs("HSXDENMN"))        ' �i�r�w�c��������          2003/12/12 SystemBrain Null�Ή�
        .HSXDENHT = rs("HSXDENHT")                      ' �i�r�w�c�����ۏؕ��@�Q��
        .HSXDENHS = rs("HSXDENHS")                      ' �i�r�w�c�����ۏؕ��@�Q��
        .HSXDVDKU = rs("HSXDVDKU")                      ' �i�r�w�c�u�c�Q�����L��
        .HSXDVDMX = fncNullCheck(rs("HSXDVDMXN"))       ' �i�r�w�c�u�c�Q���        2003/12/12 SystemBrain Null�Ή�
        .HSXDVDMN = fncNullCheck(rs("HSXDVDMNN"))       ' �i�r�w�c�u�c�Q����        2003/12/12 SystemBrain Null�Ή�
        .HSXDVDHT = rs("HSXDVDHT")                      ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
        .HSXDVDHS = rs("HSXDVDHS")                      ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
        .HSXLDLKU = rs("HSXLDLKU")                      ' �i�r�w�k�^�c�k�����L��
        .HSXLDLMX = fncNullCheck(rs("HSXLDLMX"))        ' �i�r�w�k�^�c�k���        2003/12/12 SystemBrain Null�Ή�
        .HSXLDLMN = fncNullCheck(rs("HSXLDLMN"))        ' �i�r�w�k�^�c�k����        2003/12/12 SystemBrain Null�Ή�
        .HSXLDLHT = rs("HSXLDLHT")                      ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
        .HSXLDLHS = rs("HSXLDLHS")                      ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
        
        If Not IsNull(rs("HSXOSF1PTK")) Then .HSXOSF1PTK = rs("HSXOSF1PTK")     ' �i�r�w�n�r�e�P�p�^���敪
        If Not IsNull(rs("HSXOSF2PTK")) Then .HSXOSF2PTK = rs("HSXOSF2PTK")     ' �i�r�w�n�r�e�Q�p�^���敪
        If Not IsNull(rs("HSXOSF3PTK")) Then .HSXOSF3PTK = rs("HSXOSF3PTK")     ' �i�r�w�n�r�e�R�p�^���敪
        If Not IsNull(rs("HSXOSF4PTK")) Then .HSXOSF4PTK = rs("HSXOSF4PTK")     ' �i�r�w�n�r�e�S�p�^���敪
'        If Not IsNull(rs("HSXBMD1MBP")) Then .HSXBMD1MBP = rs("HSXBMD1MBP")     ' �i�r�w�a�l�c�P�ʓ����z
'        If Not IsNull(rs("HSXBMD2MBP")) Then .HSXBMD2MBP = rs("HSXBMD2MBP")     ' �i�r�w�a�l�c�Q�ʓ����z
'        If Not IsNull(rs("HSXBMD3MBP")) Then .HSXBMD3MBP = rs("HSXBMD3MBP")     ' �i�r�w�a�l�c�R�ʓ����z
        .HSXBMD1MBP = fncNullCheck(rs("HSXBMD1MBP"))                            ' �i�r�w�a�l�c�P�ʓ����z    2003/12/12 SystemBrain Null�Ή�
        .HSXBMD2MBP = fncNullCheck(rs("HSXBMD2MBP"))                            ' �i�r�w�a�l�c�Q�ʓ����z    2003/12/12 SystemBrain Null�Ή�
        .HSXBMD3MBP = fncNullCheck(rs("HSXBMD3MBP"))                            ' �i�r�w�a�l�c�R�ʓ����z    2003/12/12 SystemBrain Null�Ή�
        
        If Not IsNull(rs("HSXGDPTK")) Then .HSXGDPTK = rs("HSXGDPTK") Else .HSXGDPTK = " "  ' �i�r�w�f�c�p�^���敪  '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
    
        'Add Start 2011/02/01 SMPK Miyata
        If Not IsNull(rs("HSXCPK")) Then .HSXCPK = rs("HSXCPK")         ' �i�r�w�b�p�^�[���敪
        If Not IsNull(rs("HSXCSZ")) Then .HSXCSZ = rs("HSXCSZ")         ' �i�r�w�b�������
        If Not IsNull(rs("HSXCHT")) Then .HSXCHT = rs("HSXCHT")         ' �i�r�w�b�ۏؕ��@�Q��
        If Not IsNull(rs("HSXCHS")) Then .HSXCHS = rs("HSXCHS")         ' �i�r�w�b�ۏؕ��@�Q��
        If Not IsNull(rs("HSXCJPK")) Then .HSXCJPK = rs("HSXCJPK")      ' �i�r�w�b�i�p�^�[���敪
        If Not IsNull(rs("HSXCJNS")) Then .HSXCJNS = rs("HSXCJNS")      ' �i�r�w�b�i�M�����@
        If Not IsNull(rs("HSXCJHT")) Then .HSXCJHT = rs("HSXCJHT")      ' �i�r�w�b�i�ۏؕ��@�Q��
        If Not IsNull(rs("HSXCJHS")) Then .HSXCJHS = rs("HSXCJHS")      ' �i�r�w�b�i�ۏؕ��@�Q��
        If Not IsNull(rs("HSXCJLTPK")) Then .HSXCJLTPK = rs("HSXCJLTPK")  ' �i�r�w�b�i�k�s�p�^�[���敪
        If Not IsNull(rs("HSXCJLTNS")) Then .HSXCJLTNS = rs("HSXCJLTNS")  ' �i�r�w�b�i�k�s�M�����@
        If Not IsNull(rs("HSXCJLTHT")) Then .HSXCJLTHT = rs("HSXCJLTHT")  ' �i�r�w�b�i�k�s�ۏؕ��@�Q��
        If Not IsNull(rs("HSXCJLTHS")) Then .HSXCJLTHS = rs("HSXCJLTHS")  ' �i�r�w�b�i�k�s�ۏؕ��@�Q��
        If Not IsNull(rs("HSXCJ2PK")) Then .HSXCJ2PK = rs("HSXCJ2PK")   ' �i�r�w�b�i�Q�p�^�[���敪
        If Not IsNull(rs("HSXCJ2NS")) Then .HSXCJ2NS = rs("HSXCJ2NS")   ' �i�r�w�b�i�Q�M�����@
        If Not IsNull(rs("HSXCJ2HT")) Then .HSXCJ2HT = rs("HSXCJ2HT")   ' �i�r�w�b�i�Q�ۏؕ��@�Q��
        If Not IsNull(rs("HSXCJ2HS")) Then .HSXCJ2HS = rs("HSXCJ2HS")   ' �i�r�w�b�i�Q�ۏؕ��@�Q��
        'Add End   2011/02/01 SMPK Miyata
    
    End With
    Set rs = Nothing

    funGet_TBCME020 = FUNCTION_RETURN_SUCCESS

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
' TBCME036�f�[�^�擾(����p)
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME036�v����w��i�Ԃ̃��R�[�h�𒊏o����B
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :tHin          ,I  ,tFullHinban                          :�i��
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmkc001c_Siyou     :���o���R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN                      :���o�̐���
'����      :�����L�[�́A�HINBAN�+�uMNOREVNO�v+�uFACTORY�v+�uOPECOND�v�̕�����Ƃ���
'����      :2003/09/10 �V�K�쐬�@�V�X�e���u���C��

Public Function funGet_TBCME036(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmkc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL�S��
    Dim rs          As OraDynaset       'RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME036"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga STRAT ---
'*** UPDATE �� Y.SIMIZU 2005/10/12 GDײݒǉ�
'    sql = sql & "EPDUP, TOPREG, TAILREG, BTMSPRT, BLOCKHFLAG "
'    sql = sql & "EPDUP, TOPREG, TAILREG, BTMSPRT, BLOCKHFLAG,HSXGDLINE "
'*** UPDATE �� Y.SIMIZU 2005/10/12 GDײݒǉ�
    sql = sql & "EPDUP, TOPREG, TAILREG, BTMSPRT, BLOCKHFLAG, HSXGDLINE, COSF3FLAG "
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga END   ---
'--------------- 2008/08/25 INSERT START  By Systech ---------------
    sql = sql & ",NVL(HSXDKTMP,' ') HSXDKTMP "
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    sql = sql & ",HSXLDLRMN, HSXLDLRMX, HWFLDLRMN, HWFLDLRMX, HSXOF1ARPTK, HSXOFARMIN, HSXOFARMAX, HSXOFARMHMX "
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    'Add Start 2011/02/01 SMPK Miyata
    sql = sql & ",HSXCJLTBND "
    'Add End   2011/02/01 SMPK Miyata

    sql = sql & "from TBCME036 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME036 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''���o���ʂ��i�[����
     With tGetRec
        .hin.hinban = rs("HINBAN")          ' �i��
        .hin.mnorevno = rs("MNOREVNO")      ' ���i�ԍ������ԍ�
        .hin.factory = rs("FACTORY")        ' �H��
        .hin.opecond = rs("OPECOND")        ' ���Ə���
        
'        If Not IsNull(rs("EPDUP")) Then .EPDUP = rs("EPDUP")                    ' EPD���
'        If Not IsNull(rs("TOPREG")) Then .TOPREG = rs("TOPREG")                 ' TOP�K��
'        If Not IsNull(rs("TAILREG")) Then .TAILREG = rs("TAILREG")              ' TAIL�K��
'        If Not IsNull(rs("BTMSPRT")) Then .BTMSPRT = rs("BTMSPRT")              ' �{�g���͏o�K��
        .EPDUP = fncNullCheck(rs("EPDUP"))                                      ' EPD���                   2003/12/12 SystemBrain Null�Ή�
        .TOPREG = fncNullCheck(rs("TOPREG"))                                    ' TOP�K��                   2003/12/12 SystemBrain Null�Ή�
        .TAILREG = fncNullCheck(rs("TAILREG"))                                  ' TAIL�K��                  2003/12/12 SystemBrain Null�Ή�
        .BTMSPRT = fncNullCheck(rs("BTMSPRT"))                                  ' �{�g���͏o�K��            2003/12/12 SystemBrain Null�Ή�
        If Not IsNull(rs("BLOCKHFLAG")) Then .BLOCKHFLAG = rs("BLOCKHFLAG")     ' �u���b�N�P�ʕۏؕi�ԃt���O
    '*** UPDATE �� Y.SIMIZU 2005/10/12 GDײݒǉ�
        .HSXGDLINE = fncNullCheck(rs("HSXGDLINE"))
    '*** UPDATE �� Y.SIMIZU 2005/10/12 GDײݒǉ�
    
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga STRAT ---
        If IsNull(rs("COSF3FLAG")) = False Then .COSF3FLAG = rs("COSF3FLAG") Else .COSF3FLAG = " "            'C-OSF3�׸�
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga END   ---

'--------------- 2008/08/25 INSERT START  By Systech ---------------
        .HSXDKTMP = rs("HSXDKTMP")
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
        .HSXLDLRMN = fncNullCheck(rs("HSXLDLRMN"))      ' �iSXL/DL�A��0����
        .HSXLDLRMX = fncNullCheck(rs("HSXLDLRMX"))      ' �iSXL/DL�A��0���
        .HWFLDLRMN = fncNullCheck(rs("HWFLDLRMN"))      ' �iWFL/DL�A��0����
        .HWFLDLRMX = fncNullCheck(rs("HWFLDLRMX"))      ' �iWFL/DL�A��0���
        If IsNull(rs("HSXOF1ARPTK")) = False Then .HSXOF1ARPTK = rs("HSXOF1ARPTK") Else .HSXOF1ARPTK = " "  ' �iSXOSF1(ArAN)�p�^���敪
        .HSXOFARMIN = fncNullCheck(rs("HSXOFARMIN"))    ' �iSXOSF(ArAN)����
        .HSXOFARMAX = fncNullCheck(rs("HSXOFARMAX"))    ' �iSXOSF(ArAN)���
        .HSXOFARMHMX = fncNullCheck(rs("HSXOFARMHMX"))  ' �iSXOSF(ArAN)�ʓ�����
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
        'Add Start 2011/02/01 SMPK Miyata
        .HSXCJLTBND = fncNullCheck(rs("HSXCJLTBND"))    ' �iSXL/CJLT�o���h�� Number(3,0)
        'Add End   2011/02/01 SMPK Miyata

    End With
    Set rs = Nothing

    funGet_TBCME036 = FUNCTION_RETURN_SUCCESS
  

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
' TBCME036�f�[�^�擾(����p)
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME036�v����w��i�Ԃ̃��R�[�h�𒊏o����B
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :tHin          ,I  ,tFullHinban                          :�i��
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :���o���R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN                      :���o�̐���
'����      :�����L�[�́A�HINBAN�+�uMNOREVNO�v+�uFACTORY�v+�uOPECOND�v�̕�����Ƃ���
'����      :2005/10/12 Y.SIMIZU

Public Function funGet_TBCME036_2(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL�S��
    Dim rs          As OraDynaset       'RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME036_2"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sql = sql & "HWFGDLINE "
    sql = sql & "from TBCME036 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME036_2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''���o���ʂ��i�[����
     With tGetRec
        .HWFGDLINE = fncNullCheck(rs("HWFGDLINE"))                                      ' EPD���                   2003/12/12 SystemBrain Null�Ή�
    End With
    Set rs = Nothing

    funGet_TBCME036_2 = FUNCTION_RETURN_SUCCESS
  

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
' TBCME021�f�[�^�擾(����p)
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME021�v����w��i�Ԃ̃��R�[�h�𒊏o����B
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :tHin          ,I  ,tFullHinban                          :�i��
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :���o���R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN                      :���o�̐���
'����      :�����L�[�́A�HINBAN�+�uMNOREVNO�v+�uFACTORY�v+�uOPECOND�v�̕�����Ƃ���
'����      :2003/09/10 �V�K�쐬�@�V�X�e���u���C��

Public Function funGet_TBCME021(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL�S��
    Dim rs          As OraDynaset       'RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME021"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sql = sql & "HWFTYPE, HWFRMIN, HWFRMAX, HWFRSPOH, HWFRSPOT, HWFRSPOI, "
    sql = sql & "HWFRHWYT, HWFRHWYS, HWFRMCAL, HWFRAMIN, HWFRAMAX, HWFRMBNP "
    sql = sql & "from TBCME021 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME021 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''���o���ʂ��i�[����
     With tGetRec
'        .HIN.hinban = rs("HINBAN")          ' �i��
'        .HIN.mnorevno = rs("MNOREVNO")      ' ���i�ԍ������ԍ�
'        .HIN.factory = rs("FACTORY")        ' �H��
'        .HIN.opecond = rs("OPECOND")        ' ���Ə���
        
        .HWFTYPE = rs("HWFTYPE")                        ' �i�v�e�^�C�v
        .HWFRMIN = fncNullCheck(rs("HWFRMIN"))          ' �i�v�e���R����          2003/12/12 SystemBrain Null�Ή�
        .HWFRMAX = fncNullCheck(rs("HWFRMAX"))          ' �i�v�e���R���          2003/12/12 SystemBrain Null�Ή�
        .HWFRSPOH = rs("HWFRSPOH")                      ' �i�v�e���R����ʒu�Q��
        .HWFRSPOT = rs("HWFRSPOT")                      ' �i�v�e���R����ʒu�Q�_
        .HWFRSPOI = rs("HWFRSPOI")                      ' �i�v�e���R����ʒu�Q��
        .HWFRHWYT = rs("HWFRHWYT")                      ' �i�v�e���R�ۏؕ��@�Q��
        .HWFRHWYS = rs("HWFRHWYS")                      ' �i�v�e���R�ۏؕ��@�Q��
        .HWFRMCAL = rs("HWFRMCAL")                      ' �i�v�e���R�ʓ��v�Z
        .HWFRAMIN = fncNullCheck(rs("HWFRAMIN"))        ' �i�v�e���R���ω���      2003/12/12 SystemBrain Null�Ή�
        .HWFRAMAX = fncNullCheck(rs("HWFRAMAX"))        ' �i�v�e���R���Ϗ��      2003/12/12 SystemBrain Null�Ή�
        .HWFRMBNP = fncNullCheck(rs("HWFRMBNP"))        ' �i�v�e���R�ʓ����z      2003/12/12 SystemBrain Null�Ή�
    End With
    Set rs = Nothing

    funGet_TBCME021 = FUNCTION_RETURN_SUCCESS
  

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
' TBCME024�f�[�^�擾(����p)
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME024�v����w��i�Ԃ̃��R�[�h�𒊏o����B
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :tHin          ,I  ,tFullHinban                          :�i��
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :���o���R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN                      :���o�̐���
'����      :�����L�[�́A�HINBAN�+�uMNOREVNO�v+�uFACTORY�v+�uOPECOND�v�̕�����Ƃ���
'����      :2003/09/10 �V�K�쐬�@�V�X�e���u���C��

Public Function funGet_TBCME024(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL�S��
    Dim rs          As OraDynaset       'RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME024"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
    sql = sql & "HWFMKMIN, HWFMKMAX, HWFMKSPH, HWFMKSPT, HWFMKSPR, HWFMKHWT, HWFMKHWS "
    sql = sql & "from TBCME024 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME024 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''���o���ʂ��i�[����
     With tGetRec
'        .HIN.hinban = rs("HINBAN")          ' �i��
'        .HIN.mnorevno = rs("MNOREVNO")      ' ���i�ԍ������ԍ�
'        .HIN.factory = rs("FACTORY")        ' �H��
'        .HIN.opecond = rs("OPECOND")        ' ���Ə���
        
        .HWFMKMIN = fncNullCheck(rs("HWFMKMIN"))        ' �i�v�e�����בw����            2003/12/12 SystemBrain Null�Ή�
        .HWFMKMAX = fncNullCheck(rs("HWFMKMAX"))        ' �i�v�e�����בw���            2003/12/12 SystemBrain Null�Ή�
        .HWFMKSPH = rs("HWFMKSPH")                      ' �i�v�e�����בw����ʒu�Q��
        .HWFMKSPT = rs("HWFMKSPT")                      ' �i�v�e�����בw����ʒu�Q�_
        .HWFMKSPR = rs("HWFMKSPR")                      ' �i�v�e�����בw����ʒu�Q��
        .HWFMKHWT = rs("HWFMKHWT")                      ' �i�v�e�����בw�ۏؕ��@�Q��
        .HWFMKHWS = rs("HWFMKHWS")                      ' �i�v�e�����בw�ۏؕ��@�Q��
    End With
    Set rs = Nothing

    funGet_TBCME024 = FUNCTION_RETURN_SUCCESS
  

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
' TBCME025�f�[�^�擾
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME025�v����w��i�Ԃ̃��R�[�h�𒊏o����B
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :tHin          ,I  ,tFullHinban                          :�i��
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :���o���R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN                      :���o�̐���
'����      :�����L�[�́A�HINBAN�+�uMNOREVNO�v+�uFACTORY�v+�uOPECOND�v�̕�����Ƃ���
'����      :2003/09/10 �V�K�쐬�@�V�X�e���u���C��

Public Function funGet_TBCME025(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL�S��
    Dim rs          As OraDynaset       'RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME025"

    sql = "select E025.HINBAN, E025.MNOREVNO, E025.FACTORY, E025.OPECOND, "
    sql = sql & "E025.HWFONMIN, E025.HWFONMAX, E025.HWFONSPH, E025.HWFONSPT, E025.HWFONSPI, E025.HWFONHWT, E025.HWFONHWS, "
    sql = sql & "HSXONSPT, HSXONSPI, "
    sql = sql & "E025.HWFONMCL, E025.HWFONMBP, E025.HWFONAMN, E025.HWFONAMX, "
    sql = sql & "E025.HWFOS1MN, E025.HWFOS1MX, E025.HWFOS1SH, E025.HWFOS1ST, E025.HWFOS1SI, E025.HWFOS1HT, E025.HWFOS1HS, E025.HWFOS1NS, "
    sql = sql & "E025.HWFOS2MN, E025.HWFOS2MX, E025.HWFOS2SH, E025.HWFOS2ST, E025.HWFOS2SI, E025.HWFOS2HT, E025.HWFOS2HS, E025.HWFOS2NS, "
    sql = sql & "E025.HWFOS3MN, E025.HWFOS3MX, E025.HWFOS3SH, E025.HWFOS3ST, E025.HWFOS3SI, E025.HWFOS3HT, E025.HWFOS3HS, E025.HWFOS3NS, "
    ''�c���_�f�d�l�擾�ǉ��@03/12/09 ooba
    sql = sql & "E025.HWFZOMIN, E025.HWFZOMAX, E025.HWFZOSPH, E025.HWFZOSPT, E025.HWFZOSPI, E025.HWFZOHWT, E025.HWFZOHWS, E025.HWFZONSW, "
    sql = sql & "E025.HWFANTNP, E025.HWFANTIM "
    sql = sql & "from TBCME025 E025, TBCME019 E019 "
    sql = sql & "Where E025.HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      E025.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E025.FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      E025.OPECOND = '" & tHIN.opecond & "' and "
    sql = sql & "      E019.HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      E019.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E019.FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      E019.OPECOND = '" & tHIN.opecond & "'"
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME025 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''���o���ʂ��i�[����
     With tGetRec
'        .HIN.hinban = rs("HINBAN")          ' �i��
'        .HIN.mnorevno = rs("MNOREVNO")      ' ���i�ԍ������ԍ�
'        .HIN.factory = rs("FACTORY")        ' �H��
'        .HIN.opecond = rs("OPECOND")        ' ���Ə���
        
        .HWFONMIN = fncNullCheck(rs("HWFONMIN"))        ' �i�v�e�_�f�Z�x����            2003/12/12 SystemBrain Null�Ή�
        .HWFONMAX = fncNullCheck(rs("HWFONMAX"))        ' �i�v�e�_�f�Z�x���            2003/12/12 SystemBrain Null�Ή�
        .HWFONSPH = rs("HWFONSPH")                      ' �i�v�e�_�f�Z�x����ʒu�Q��
'        .HWFONSPT = rs("HWFONSPT")                      ' �i�v�e�_�f�Z�x����ʒu�Q�_
'        .HWFONSPI = rs("HWFONSPI")                      ' �i�v�e�_�f�Z�x����ʒu�Q��
        .HWFONSPT = rs("HSXONSPT")                      ' �i�r�w�_�f�Z�x����ʒu�Q�_
        .HWFONSPI = rs("HSXONSPI")                      ' �i�r�w�_�f�Z�x����ʒu�Q��
        .HWFONHWT = rs("HWFONHWT")                      ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��
        .HWFONHWS = rs("HWFONHWS")                      ' �i�v�e�_�f�Z�x�ۏؕ��@�Q��
        .HWFONMCL = rs("HWFONMCL")                      ' �i�v�e�_�f�Z�x�ʓ��v�Z
        .HWFONMBP = fncNullCheck(rs("HWFONMBP"))        ' �i�v�e�_�f�Z�x�ʓ����z        2003/12/12 SystemBrain Null�Ή�
        .HWFONAMN = fncNullCheck(rs("HWFONAMN"))        ' �i�v�e�_�f�Z�x���ω���        2003/12/12 SystemBrain Null�Ή�
        .HWFONAMX = fncNullCheck(rs("HWFONAMX"))        ' �i�v�e�_�f�Z�x���Ϗ��        2003/12/12 SystemBrain Null�Ή�
        
        .HWFOS1MN = fncNullCheck(rs("HWFOS1MN"))        ' �i�v�e�_�f�͏o�P����          2003/12/12 SystemBrain Null�Ή�
        .HWFOS1MX = fncNullCheck(rs("HWFOS1MX"))        ' �i�v�e�_�f�͏o�P���          2003/12/12 SystemBrain Null�Ή�
        .HWFOS1SH = rs("HWFOS1SH")                      ' �i�v�e�_�f�͏o�P����ʒu�Q��
        .HWFOS1ST = rs("HWFOS1ST")                      ' �i�v�e�_�f�͏o�P����ʒu�Q�_
        .HWFOS1SI = rs("HWFOS1SI")                      ' �i�v�e�_�f�͏o�P����ʒu�Q��
        .HWFOS1HT = rs("HWFOS1HT")                      ' �i�v�e�_�f�͏o�P�ۏؕ��@�Q��
        .HWFOS1HS = rs("HWFOS1HS")                      ' �i�v�e�_�f�͏o�P�ۏؕ��@�Q��
        .HWFOS1NS = rs("HWFOS1NS")                      ' �i�v�e�_�f�͏o�P�M�����@
        
        .HWFOS2MN = fncNullCheck(rs("HWFOS2MN"))        ' �i�v�e�_�f�͏o�Q����          2003/12/12 SystemBrain Null�Ή�
        .HWFOS2MX = fncNullCheck(rs("HWFOS2MX"))        ' �i�v�e�_�f�͏o�Q���          2003/12/12 SystemBrain Null�Ή�
        .HWFOS2SH = rs("HWFOS2SH")                      ' �i�v�e�_�f�͏o�Q����ʒu�Q��
        .HWFOS2ST = rs("HWFOS2ST")                      ' �i�v�e�_�f�͏o�Q����ʒu�Q�_
        .HWFOS2SI = rs("HWFOS2SI")                      ' �i�v�e�_�f�͏o�Q����ʒu�Q��
        .HWFOS2HT = rs("HWFOS2HT")                      ' �i�v�e�_�f�͏o�Q�ۏؕ��@�Q��
        .HWFOS2HS = rs("HWFOS2HS")                      ' �i�v�e�_�f�͏o�Q�ۏؕ��@�Q��
        .HWFOS2NS = rs("HWFOS2NS")                      ' �i�v�e�_�f�͏o�Q�M�����@
        
        .HWFOS3MN = fncNullCheck(rs("HWFOS3MN"))        ' �i�v�e�_�f�͏o�R����          2003/12/12 SystemBrain Null�Ή�
        .HWFOS3MX = fncNullCheck(rs("HWFOS3MX"))        ' �i�v�e�_�f�͏o�R���          2003/12/12 SystemBrain Null�Ή�
        .HWFOS3SH = rs("HWFOS3SH")                      ' �i�v�e�_�f�͏o�R����ʒu�Q��
        .HWFOS3ST = rs("HWFOS3ST")                      ' �i�v�e�_�f�͏o�R����ʒu�Q�_
        .HWFOS3SI = rs("HWFOS3SI")                      ' �i�v�e�_�f�͏o�R����ʒu�Q��
        .HWFOS3HT = rs("HWFOS3HT")                      ' �i�v�e�_�f�͏o�R�ۏؕ��@�Q��
        .HWFOS3HS = rs("HWFOS3HS")                      ' �i�v�e�_�f�͏o�R�ۏؕ��@�Q��
        .HWFOS3NS = rs("HWFOS3NS")                      ' �i�v�e�_�f�͏o�R�M�����@
        
        ''�c���_�f�d�l�擾�ǉ��@03/12/09 ooba START ==============================>
'''        If IsNull(rs("HWFZOMIN")) = False Then .HWFZOMIN = rs("HWFZOMIN") ' �i�v�e�c���_�f����
'''        If IsNull(rs("HWFZOMAX")) = False Then .HWFZOMAX = rs("HWFZOMAX") ' �i�v�e�c���_�f���
'''        .HWFZOSPH = rs("HWFZOSPH")                  ' �i�v�e�c���_�f����ʒu�Q��
'''        .HWFZOSPT = rs("HWFZOSPT")                  ' �i�v�e�c���_�f����ʒu�Q�_
'''        .HWFZOSPI = rs("HWFZOSPI")                  ' �i�v�e�c���_�f����ʒu�Q��
'''        .HWFZOHWT = rs("HWFZOHWT")                  ' �i�v�e�c���_�f�ۏؕ��@�Q��
'''        .HWFZOHWS = rs("HWFZOHWS")                  ' �i�v�e�c���_�f�ۏؕ��@�Q��
'''        .HWFZONSW = rs("HWFZONSW")                  ' �i�v�e�c���_�f�M�����@

        .HWFZOMIN = fncNullCheck(rs("HWFZOMIN"))    ' �i�v�e�c���_�f����
        .HWFZOMAX = fncNullCheck(rs("HWFZOMAX"))    ' �i�v�e�c���_�f���
        If IsNull(rs("HWFZOSPH")) = False Then .HWFZOSPH = rs("HWFZOSPH") ' �i�v�e�c���_�f����ʒu�Q��
        If IsNull(rs("HWFZOSPT")) = False Then .HWFZOSPT = rs("HWFZOSPT") ' �i�v�e�c���_�f����ʒu�Q�_
        If IsNull(rs("HWFZOSPI")) = False Then .HWFZOSPI = rs("HWFZOSPI") ' �i�v�e�c���_�f����ʒu�Q��
        If IsNull(rs("HWFZOHWT")) = False Then .HWFZOHWT = rs("HWFZOHWT") ' �i�v�e�c���_�f�ۏؕ��@�Q��
        If IsNull(rs("HWFZOHWS")) = False Then .HWFZOHWS = rs("HWFZOHWS") ' �i�v�e�c���_�f�ۏؕ��@�Q��
        If IsNull(rs("HWFZONSW")) = False Then .HWFZONSW = rs("HWFZONSW") ' �i�v�e�c���_�f�M�����@
        ''�c���_�f�d�l�擾�ǉ��@03/12/09 ooba END ================================>
        
        .HWFANTIM = fncNullCheck(rs("HWFANTIM"))        ' �i�v�e�`�m����                2003/12/12 SystemBrain Null�Ή�
        .HWFANTNP = fncNullCheck(rs("HWFANTNP"))        ' �i�v�e�`�m���x                2003/12/12 SystemBrain Null�Ή�
    End With
    Set rs = Nothing

    funGet_TBCME025 = FUNCTION_RETURN_SUCCESS
  

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
' TBCME026�f�[�^�擾
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME026�v����w��i�Ԃ̃��R�[�h�𒊏o����B
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :tHin          ,I  ,tFullHinban                          :�i��
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :���o���R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN                      :���o�̐���
'����      :�����L�[�́A�HINBAN�+�uMNOREVNO�v+�uFACTORY�v+�uOPECOND�v�̕�����Ƃ���
'����      :2003/09/10 �V�K�쐬�@�V�X�e���u���C��

Public Function funGet_TBCME026(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL�S��
    Dim rs          As OraDynaset       'RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME026"

    'DSOD����݋敪�擾�ǉ��@04/08/09
    'GD�d�l�擾�ǉ��@05/01/26
''    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
''    sql = sql & "HWFDSOPTK, "
''    sql = sql & "HWFDENKU, HWFDENMX, HWFDENMN, HWFDENHT, HWFDENHS, "
''    sql = sql & "HWFDVDKU, HWFDVDMXN, HWFDVDMNN, HWFDVDHT, HWFDVDHS, "
''    sql = sql & "HWFLDLKU, HWFLDLMX, HWFLDLMN, HWFLDLHT, HWFLDLHS, "
''    sql = sql & "HWFGDSPH, HWFGDSPT, HWFGDSPR, "
''    sql = sql & "HWFDSOMX, HWFDSOMN, HWFDSOAX, HWFDSOAN, HWFDSOHT, HWFDSOHS "
''    sql = sql & "from TBCME026 "
''    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
''    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
''    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
''    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    'DK�ưى��x�ǉ��@06/12/22 ooba
    sql = "select E025.HINBAN, E025.MNOREVNO, E025.FACTORY, E025.OPECOND, "
    sql = sql & "E026.HWFDENKU, E026.HWFDENMX, E026.HWFDENMN, E026.HWFDENHT, E026.HWFDENHS, "
    sql = sql & "E026.HWFDVDKU, E026.HWFDVDMXN, E026.HWFDVDMNN, E026.HWFDVDHT, E026.HWFDVDHS, "
    sql = sql & "E026.HWFLDLKU, E026.HWFLDLMX, E026.HWFLDLMN, E026.HWFLDLHT, E026.HWFLDLHS, "
    sql = sql & "E026.HWFGDSPH, E026.HWFGDSPT, E026.HWFGDSPR, "
    sql = sql & "E026.HWFDSOMX, E026.HWFDSOMN, E026.HWFDSOAX, E026.HWFDSOAN, E026.HWFDSOHT, "
    sql = sql & "E026.HWFDSOHS, E026.HWFDSOPTK, E025.HWFANTNP "
    sql = sql & ",E026.HWFGDPTK "    '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
    sql = sql & "from TBCME025 E025, TBCME026 E026 "
    sql = sql & "Where E025.HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      E025.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E025.FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      E025.OPECOND = '" & tHIN.opecond & "' and "
    sql = sql & "      E026.HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      E026.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E026.FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      E026.OPECOND = '" & tHIN.opecond & "'"
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME026 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''���o���ʂ��i�[����
     With tGetRec
'        .HIN.hinban = rs("HINBAN")          ' �i��
'        .HIN.mnorevno = rs("MNOREVNO")      ' ���i�ԍ������ԍ�
'        .HIN.factory = rs("FACTORY")        ' �H��
'        .HIN.opecond = rs("OPECOND")        ' ���Ə���
        
        .HWFDSOMX = fncNullCheck(rs("HWFDSOMX"))        ' �i�v�e�c�r�n�c���            2003/12/12 SystemBrain Null�Ή�
        .HWFDSOMN = fncNullCheck(rs("HWFDSOMN"))        ' �i�v�e�c�r�n�c����            2003/12/12 SystemBrain Null�Ή�
        .HWFDSOAX = fncNullCheck(rs("HWFDSOAX"))        ' �i�v�e�c�r�n�c�̈���        2003/12/12 SystemBrain Null�Ή�
        .HWFDSOAN = fncNullCheck(rs("HWFDSOAN"))        ' �i�v�e�c�r�n�c�̈扺��        2003/12/12 SystemBrain Null�Ή�
        .HWFDSOHT = rs("HWFDSOHT")                      ' �i�v�e�c�r�n�c�ۏؕ��@�Q��
        .HWFDSOHS = rs("HWFDSOHS")                      ' �i�v�e�c�r�n�c�ۏؕ��@�Q��
        If IsNull(rs("HWFDSOPTK")) = False Then .HWFDSOPTK = rs("HWFDSOPTK") Else .HWFDSOPTK = " "          '�p�^�[���敪�@04/08/09 ooba
        
        ''GD�d�l�擾�ǉ��@05/01/26 ooba START ========================================>
        .HWFDENKU = rs("HWFDENKU")                      ' �i�v�e�c���������L��
        .HWFDENMX = fncNullCheck(rs("HWFDENMX"))        ' �i�v�e�c�������
        .HWFDENMN = fncNullCheck(rs("HWFDENMN"))        ' �i�v�e�c��������
        .HWFDENHT = rs("HWFDENHT")                      ' �i�v�e�c�����ۏؕ��@�Q��
        .HWFDENHS = rs("HWFDENHS")                      ' �i�v�e�c�����ۏؕ��@�Q��
        .HWFDVDKU = rs("HWFDVDKU")                      ' �i�v�e�c�u�c�Q�����L��
        .HWFDVDMXN = fncNullCheck(rs("HWFDVDMXN"))      ' �i�v�e�c�u�c�Q���
        .HWFDVDMNN = fncNullCheck(rs("HWFDVDMNN"))      ' �i�v�e�c�u�c�Q����
        .HWFDVDHT = rs("HWFDVDHT")                      ' �i�v�e�c�u�c�Q�ۏؕ��@�Q��
        .HWFDVDHS = rs("HWFDVDHS")                      ' �i�v�e�c�u�c�Q�ۏؕ��@�Q��
        .HWFLDLKU = rs("HWFLDLKU")                      ' �i�v�e�k�^�c�k�����L��
        .HWFLDLMX = fncNullCheck(rs("HWFLDLMX"))        ' �i�v�e�k�^�c�k���
        .HWFLDLMN = fncNullCheck(rs("HWFLDLMN"))        ' �i�v�e�k�^�c�k����
        .HWFLDLHT = rs("HWFLDLHT")                      ' �i�v�e�k�^�c�k�ۏؕ��@�Q��
        .HWFLDLHS = rs("HWFLDLHS")                      ' �i�v�e�k�^�c�k�ۏؕ��@�Q��
        .HWFGDSPH = rs("HWFGDSPH")                      ' �i�v�e�f�c����ʒu�Q��
        .HWFGDSPT = rs("HWFGDSPT")                      ' �i�v�e�f�c����ʒu�Q�_
        .HWFGDSPR = rs("HWFGDSPR")                      ' �i�v�e�f�c����ʒu�Q��
        ''GD�d�l�擾�ǉ��@05/01/26 ooba END ==========================================>
        
        If Not IsNull(rs("HWFANTNP")) Then .HWFANTNP = rs("HWFANTNP")   ' �i�v�e�`�m���x�@06/12/22 ooba
        
        If Not IsNull(rs("HWFGDPTK")) Then .HWFGDPTK = rs("HWFGDPTK") Else .HWFGDPTK = " "  ' �i�v�e�f�c�p�^���敪  '' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech
    End With
    Set rs = Nothing

    funGet_TBCME026 = FUNCTION_RETURN_SUCCESS
  

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
' TBCME028�f�[�^�擾
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME028�v����w��i�Ԃ̃��R�[�h�𒊏o����B
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :tHin          ,I  ,tFullHinban                          :�i��
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :���o���R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN                      :���o�̐���
'����      :�����L�[�́A�HINBAN�+�uMNOREVNO�v+�uFACTORY�v+�uOPECOND�v�̕�����Ƃ���
'����      :2003/09/10 �V�K�쐬�@�V�X�e���u���C��

Public Function funGet_TBCME028(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL�S��
    Dim rs          As OraDynaset       'RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME028"

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "

''Upd start 2005/06/28 (TCS)T.terauchi  SPV9�_�Ή�
''    sql = sql & "HWFSPVMX, HWFSPVSH, HWFSPVST, HWFSPVSI, HWFSPVHT, HWFSPVHS, "
    sql = sql & "HWFSPVMX, HWFSPVMXN, HWFSPVSH, HWFSPVST, HWFSPVSI, HWFSPVHT, HWFSPVHS, "
    sql = sql & "HWFSPVKN, HWFDLKHN, "
''Upd end   2005/06/28 (TCS)T.Terauchi  SPV9�_�Ή�

'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
    sql = sql & "HWFSPVAMN, "
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
    
    sql = sql & "HWFDLMIN, HWFDLMAX, HWFDLSPH, HWFDLSPT, HWFDLSPI, HWFDLHWT, HWFDLHWS "
    sql = sql & "from TBCME028 "
    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME028 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''���o���ʂ��i�[����
     With tGetRec
'        .HIN.hinban = rs("HINBAN")          ' �i��
'        .HIN.mnorevno = rs("MNOREVNO")      ' ���i�ԍ������ԍ�
'        .HIN.factory = rs("FACTORY")        ' �H��
'        .HIN.opecond = rs("OPECOND")        ' ���Ə���
        
    ''Upd start 2005/06/28 (TCS)T.Terauchi  SPV9�_�Ή�
    ''    .HWFSPVMX = fncNullCheck(rs("HWFSPVMX"))        ' �i�v�e�r�o�u�e�d���          2003/12/12 SystemBrain Null�Ή�
        .HWFSPVMX = fncNullCheck(rs("HWFSPVMXN"))       ' �i�v�e�r�o�u�e�d���
        .HWFSPVKN = rs("HWFSPVKN")                      ' �i�v�e�r�o�u�e�d�����p�x�Q��
        .HWFDLKHN = rs("HWFDLKHN")                      ' �i�v�e�g�U�������p�x�Q��
    ''Upd end   2005/06/28 (TCS)T.Terauchi  SPV9�_�Ή�
    
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
        .HWFSPVAM = fncNullCheck(rs("HWFSPVAMN"))       ' �i�v�e�r�o�u�e�d����
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
        
        
        .HWFSPVSH = rs("HWFSPVSH")                      ' �i�v�e�r�o�u�e�d����ʒu�Q��
        .HWFSPVST = rs("HWFSPVST")                      ' �i�v�e�r�o�u�e�d����ʒu�Q�_
        .HWFSPVSI = rs("HWFSPVSI")                      ' �i�v�e�r�o�u�e�d����ʒu�Q��
        .HWFSPVHT = rs("HWFSPVHT")                      ' �i�v�e�r�o�u�e�d�ۏؕ��@�Q��
        .HWFSPVHS = rs("HWFSPVHS")                      ' �i�v�e�r�o�u�e�d�ۏؕ��@�Q��
        
        .HWFDLMIN = fncNullCheck(rs("HWFDLMIN"))        ' �i�v�e�g�U������              2003/12/12 SystemBrain Null�Ή�
        .HWFDLMAX = fncNullCheck(rs("HWFDLMAX"))        ' �i�v�e�g�U�����              2003/12/12 SystemBrain Null�Ή�
        .HWFDLSPH = rs("HWFDLSPH")                      ' �i�v�e�g�U������ʒu�Q��
        .HWFDLSPT = rs("HWFDLSPT")                      ' �i�v�e�g�U������ʒu�Q�_
        .HWFDLSPI = rs("HWFDLSPI")                      ' �i�v�e�g�U������ʒu�Q��
        .HWFDLHWT = rs("HWFDLHWT")                      ' �i�v�e�g�U���ۏؕ��@�Q��
        .HWFDLHWS = rs("HWFDLHWS")                      ' �i�v�e�g�U���ۏؕ��@�Q��
    End With
    Set rs = Nothing

    funGet_TBCME028 = FUNCTION_RETURN_SUCCESS
  

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
' TBCME029�f�[�^�擾
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME029�v����w��i�Ԃ̃��R�[�h�𒊏o����B
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :tHin          ,I  ,tFullHinban                          :�i��
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :���o���R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN                      :���o�̐���
'����      :�����L�[�́A�HINBAN�+�uMNOREVNO�v+�uFACTORY�v+�uOPECOND�v�̕�����Ƃ���
'����      :2003/09/10 �V�K�쐬�@�V�X�e���u���C��

Public Function funGet_TBCME029(tHIN As tFullHinban, tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou) As FUNCTION_RETURN
    Dim sql         As String           'SQL�S��
    Dim rs          As OraDynaset       'RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_Com.bas -- Function funGet_TBCME029"

'���ύX �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    'AN���x�`�F�b�N�ׂ̈�TBCME025����i�v�e�`�m���x���擾����
'    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, "
'    sql = sql & "HWFOF1AX, HWFOF1MX, HWFOF1SH, HWFOF1ST, HWFOF1SR, HWFOF1HT, HWFOF1HS, HWFOF1NS, HWFOF1ET, HWFOF1SZ, "
'    sql = sql & "HWFOF2AX, HWFOF2MX, HWFOF2SH, HWFOF2ST, HWFOF2SR, HWFOF2HT, HWFOF2HS, HWFOF2NS, HWFOF2ET, HWFOF2SZ, "
'    sql = sql & "HWFOF3AX, HWFOF3MX, HWFOF3SH, HWFOF3ST, HWFOF3SR, HWFOF3HT, HWFOF3HS, HWFOF3NS, HWFOF3ET, HWFOF3SZ, "
'    sql = sql & "HWFOF4AX, HWFOF4MX, HWFOF4SH, HWFOF4ST, HWFOF4SR, HWFOF4HT, HWFOF4HS, HWFOF4NS, HWFOF4ET, HWFOF4SZ, "
'    sql = sql & "HWFBM1AN, HWFBM1AX, HWFBM1SH, HWFBM1ST, HWFBM1SR, HWFBM1HT, HWFBM1HS, HWFBM1NS, HWFBM1ET, HWFBM1SZ, "
'    sql = sql & "HWFBM2AN, HWFBM2AX, HWFBM2SH, HWFBM2ST, HWFBM2SR, HWFBM2HT, HWFBM2HS, HWFBM2NS, HWFBM2ET, HWFBM2SZ, "
'    sql = sql & "HWFBM3AN, HWFBM3AX, HWFBM3SH, HWFBM3ST, HWFBM3SR, HWFBM3HT, HWFBM3HS, HWFBM3NS, HWFBM3ET, HWFBM3SZ, "
'    sql = sql & "HWFOSF1PTK, HWFOSF2PTK, HWFOSF3PTK, HWFOSF4PTK, "
'    sql = sql & "HWFBM1MBP, HWFBM2MBP, HWFBM3MBP, HWFBM1MCL, HWFBM2MCL, HWFBM3MCL "
'    sql = sql & "from TBCME029 "
'    sql = sql & "Where HINBAN = '" & tHIN.hinban & "' and "
'    sql = sql & "      MNOREVNO = " & tHIN.mnorevno & " and "
'    sql = sql & "      FACTORY = '" & tHIN.factory & "' and "
'    sql = sql & "      OPECOND = '" & tHIN.opecond & "'"
    sql = "select E029.HINBAN, E029.MNOREVNO, E029.FACTORY, E029.OPECOND, "
    sql = sql & "E029.HWFOF1AX, E029.HWFOF1MX, E029.HWFOF1SH, E029.HWFOF1ST, E029.HWFOF1SR, E029.HWFOF1HT, E029.HWFOF1HS, E029.HWFOF1NS, E029.HWFOF1ET, HWFOF1SZ, "
    sql = sql & "E029.HWFOF2AX, E029.HWFOF2MX, E029.HWFOF2SH, E029.HWFOF2ST, E029.HWFOF2SR, E029.HWFOF2HT, E029.HWFOF2HS, E029.HWFOF2NS, E029.HWFOF2ET, HWFOF2SZ, "
    sql = sql & "E029.HWFOF3AX, E029.HWFOF3MX, E029.HWFOF3SH, E029.HWFOF3ST, E029.HWFOF3SR, E029.HWFOF3HT, E029.HWFOF3HS, E029.HWFOF3NS, E029.HWFOF3ET, HWFOF3SZ, "
    sql = sql & "E029.HWFOF4AX, E029.HWFOF4MX, E029.HWFOF4SH, E029.HWFOF4ST, E029.HWFOF4SR, E029.HWFOF4HT, E029.HWFOF4HS, E029.HWFOF4NS, E029.HWFOF4ET, HWFOF4SZ, "
    sql = sql & "E029.HWFBM1AN, E029.HWFBM1AX, E029.HWFBM1SH, E029.HWFBM1ST, E029.HWFBM1SR, E029.HWFBM1HT, E029.HWFBM1HS, E029.HWFBM1NS, E029.HWFBM1ET, HWFBM1SZ, "
    sql = sql & "E029.HWFBM2AN, E029.HWFBM2AX, E029.HWFBM2SH, E029.HWFBM2ST, E029.HWFBM2SR, E029.HWFBM2HT, E029.HWFBM2HS, E029.HWFBM2NS, E029.HWFBM2ET, HWFBM2SZ, "
    sql = sql & "E029.HWFBM3AN, E029.HWFBM3AX, E029.HWFBM3SH, E029.HWFBM3ST, E029.HWFBM3SR, E029.HWFBM3HT, E029.HWFBM3HS, E029.HWFBM3NS, E029.HWFBM3ET, HWFBM3SZ, "
    sql = sql & "E029.HWFOSF1PTK, E029.HWFOSF2PTK, E029.HWFOSF3PTK, E029.HWFOSF4PTK, "
    sql = sql & "E029.HWFBM1MBP, E029.HWFBM2MBP, E029.HWFBM3MBP, E029.HWFBM1MCL, E029.HWFBM2MCL, E029.HWFBM3MCL, "
    sql = sql & "E025.HWFANTNP "
'��--- 2010/01/20 SIRD�Ή� SPK habuki ADD START(OSF4->SIRD)
    sql = sql & ",E048.HWFSIRDMX "          '����]�ʏ��
    sql = sql & ",E048.HWFSIRDHT "          '����]�ʕۏؕ��@�Q��
    sql = sql & ",E048.HWFSIRDHS "          '����]�ʕۏؕ��@�Q��
    sql = sql & ",E048.HWFSIRDSZ "          '����]�ʑ������
'��--- 2010/01/20 SIRD�Ή� SPK habuki ADD  END (OSF4->SIRD)
    sql = sql & "from TBCME029 E029 "
    sql = sql & "    ,TBCME025 E025 "
'��--- 2010/01/20 SIRD�Ή� SPK habuki ADD START(OSF4->SIRD)
    sql = sql & "    ,TBCME048 E048 "
'��--- 2010/01/20 SIRD�Ή� SPK habuki ADD  END (OSF4->SIRD)
    sql = sql & "Where E029.HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      E029.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E029.FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      E029.OPECOND = '" & tHIN.opecond & "' and "
    sql = sql & "      E025.HINBAN = '" & tHIN.hinban & "' and "
    sql = sql & "      E025.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E025.FACTORY = '" & tHIN.factory & "' and "
    sql = sql & "      E025.OPECOND = '" & tHIN.opecond & "'"
'���ύX �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
'��--- 2010/01/20 SIRD�Ή� SPK habuki ADD START(OSF4->SIRD)
    sql = sql & "  and E048.HINBAN = '" & tHIN.HINBAN & "' and "
    sql = sql & "      E048.MNOREVNO = " & tHIN.mnorevno & " and "
    sql = sql & "      E048.FACTORY = '" & tHIN.FACTORY & "' and "
    sql = sql & "      E048.OPECOND = '" & tHIN.OPECOND & "'"
'��--- 2010/01/20 SIRD�Ή� SPK habuki ADD  END (OSF4->SIRD)
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        funGet_TBCME029 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''���o���ʂ��i�[����
     With tGetRec
'        .HIN.hinban = rs("HINBAN")          ' �i��
'        .HIN.mnorevno = rs("MNOREVNO")      ' ���i�ԍ������ԍ�
'        .HIN.factory = rs("FACTORY")        ' �H��
'        .HIN.opecond = rs("OPECOND")        ' ���Ə���
        
        .HWFOF1AX = fncNullCheck(rs("HWFOF1AX"))        ' �i�v�e�n�r�e�P���Ϗ��        2003/12/12 SystemBrain Null�Ή�
        .HWFOF1MX = fncNullCheck(rs("HWFOF1MX"))        ' �i�v�e�n�r�e�P���            2003/12/12 SystemBrain Null�Ή�
        .HWFOF1SH = rs("HWFOF1SH")                      ' �i�v�e�n�r�e�P����ʒu�Q��
        .HWFOF1ST = rs("HWFOF1ST")                      ' �i�v�e�n�r�e�P����ʒu�Q�_
        .HWFOF1SR = rs("HWFOF1SR")                      ' �i�v�e�n�r�e�P����ʒu�Q��
        .HWFOF1HT = rs("HWFOF1HT")                      ' �i�v�e�n�r�e�P�ۏؕ��@�Q��
        .HWFOF1HS = rs("HWFOF1HS")                      ' �i�v�e�n�r�e�P�ۏؕ��@�Q��
        .HWFOF1NS = rs("HWFOF1NS")                      ' �i�v�e�n�r�e�P�M�����@
        .HWFOF1ET = fncNullCheck(rs("HWFOF1ET"))        ' �i�v�e�n�r�e�P�I���d�s��      2003/12/12 SystemBrain Null�Ή�
        .HWFOF1SZ = rs("HWFOF1SZ")                      ' �i�v�e�n�r�e�P�������
        .HWFOF2AX = fncNullCheck(rs("HWFOF2AX"))        ' �i�v�e�n�r�e�Q���Ϗ��        2003/12/12 SystemBrain Null�Ή�
        .HWFOF2MX = fncNullCheck(rs("HWFOF2MX"))        ' �i�v�e�n�r�e�Q���            2003/12/12 SystemBrain Null�Ή�
        .HWFOF2SH = rs("HWFOF2SH")                      ' �i�v�e�n�r�e�Q����ʒu�Q��
        .HWFOF2ST = rs("HWFOF2ST")                      ' �i�v�e�n�r�e�Q����ʒu�Q�_
        .HWFOF2SR = rs("HWFOF2SR")                      ' �i�v�e�n�r�e�Q����ʒu�Q��
        .HWFOF2HT = rs("HWFOF2HT")                      ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��
        .HWFOF2HS = rs("HWFOF2HS")                      ' �i�v�e�n�r�e�Q�ۏؕ��@�Q��
        .HWFOF2NS = rs("HWFOF2NS")                      ' �i�v�e�n�r�e�Q�M�����@
        .HWFOF2ET = fncNullCheck(rs("HWFOF2ET"))        ' �i�v�e�n�r�e�Q�I���d�s��      2003/12/12 SystemBrain Null�Ή�
        .HWFOF2SZ = rs("HWFOF2SZ")                      ' �i�v�e�n�r�e�Q�������
        .HWFOF3AX = fncNullCheck(rs("HWFOF3AX"))        ' �i�v�e�n�r�e�R���Ϗ��        2003/12/12 SystemBrain Null�Ή�
        .HWFOF3MX = fncNullCheck(rs("HWFOF3MX"))        ' �i�v�e�n�r�e�R���            2003/12/12 SystemBrain Null�Ή�
        .HWFOF3SH = rs("HWFOF3SH")                      ' �i�v�e�n�r�e�R����ʒu�Q��
        .HWFOF3ST = rs("HWFOF3ST")                      ' �i�v�e�n�r�e�R����ʒu�Q�_
        .HWFOF3SR = rs("HWFOF3SR")                      ' �i�v�e�n�r�e�R����ʒu�Q��
        .HWFOF3HT = rs("HWFOF3HT")                      ' �i�v�e�n�r�e�R�ۏؕ��@�Q��
        .HWFOF3HS = rs("HWFOF3HS")                      ' �i�v�e�n�r�e�R�ۏؕ��@�Q��
        .HWFOF3NS = rs("HWFOF3NS")                      ' �i�v�e�n�r�e�R�M�����@
        .HWFOF3ET = fncNullCheck(rs("HWFOF3ET"))        ' �i�v�e�n�r�e�R�I���d�s��      2003/12/12 SystemBrain Null�Ή�
        .HWFOF3SZ = rs("HWFOF3SZ")                      ' �i�v�e�n�r�e�R�������
        .HWFOF4AX = fncNullCheck(rs("HWFOF4AX"))        ' �i�v�e�n�r�e�S���Ϗ��        2003/12/12 SystemBrain Null�Ή�
        .HWFOF4MX = fncNullCheck(rs("HWFOF4MX"))        ' �i�v�e�n�r�e�S���            2003/12/12 SystemBrain Null�Ή�
        .HWFOF4SH = rs("HWFOF4SH")                      ' �i�v�e�n�r�e�S����ʒu�Q��
        .HWFOF4ST = rs("HWFOF4ST")                      ' �i�v�e�n�r�e�S����ʒu�Q�_
        .HWFOF4SR = rs("HWFOF4SR")                      ' �i�v�e�n�r�e�S����ʒu�Q��
        .HWFOF4HT = rs("HWFOF4HT")                      ' �i�v�e�n�r�e�S�ۏؕ��@�Q��
        .HWFOF4HS = rs("HWFOF4HS")                      ' �i�v�e�n�r�e�S�ۏؕ��@�Q��
        .HWFOF4NS = rs("HWFOF4NS")                      ' �i�v�e�n�r�e�S�M�����@
        .HWFOF4ET = fncNullCheck(rs("HWFOF4ET"))        ' �i�v�e�n�r�e�S�I���d�s��      2003/12/12 SystemBrain Null�Ή�
        .HWFOF4SZ = rs("HWFOF4SZ")                      ' �i�v�e�n�r�e�S�������
'��--- 2010/01/20 SIRD�Ή� SPK habuki ADD START(OSF4->SIRD)
        If IsNull(rs("HWFSIRDMX")) = False Then .HWFOF4MX = rs("HWFSIRDMX") Else .HWFOF4MX = "0"        ' ����]�ʏ��
        If IsNull(rs("HWFSIRDHT")) = False Then .HWFOF4HT = rs("HWFSIRDHT") Else .HWFOF4HT = " "        ' ����]�ʕۏؕ��@�Q��
        If IsNull(rs("HWFSIRDHS")) = False Then .HWFOF4HS = rs("HWFSIRDHS") Else .HWFOF4HS = " "        ' ����]�ʕۏؕ��@�Q��
        If IsNull(rs("HWFSIRDSZ")) = False Then .HWFOF4SZ = rs("HWFSIRDSZ") Else .HWFOF4SZ = " "        ' ����]�ʑ������
'��--- 2010/01/20 SIRD�Ή� SPK habuki ADD  END (OSF4->SIRD)
        
        .HWFBM1AN = fncNullCheck(rs("HWFBM1AN"))        ' �i�v�e�a�l�c�P���ω���        2003/12/12 SystemBrain Null�Ή�
        .HWFBM1AX = fncNullCheck(rs("HWFBM1AX"))        ' �i�v�e�a�l�c�P���Ϗ��        2003/12/12 SystemBrain Null�Ή�
        .HWFBM1SH = rs("HWFBM1SH")                      ' �i�v�e�a�l�c�P����ʒu�Q��
        .HWFBM1ST = rs("HWFBM1ST")                      ' �i�v�e�a�l�c�P����ʒu�Q�_
        .HWFBM1SR = rs("HWFBM1SR")                      ' �i�v�e�a�l�c�P����ʒu�Q��
        .HWFBM1HT = rs("HWFBM1HT")                      ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
        .HWFBM1HS = rs("HWFBM1HS")                      ' �i�v�e�a�l�c�P�ۏؕ��@�Q��
        .HWFBM1NS = rs("HWFBM1NS")                      ' �i�v�e�a�l�c�P�M�����@
        .HWFBM1ET = fncNullCheck(rs("HWFBM1ET"))        ' �i�v�e�a�l�c�P�I���d�s��      2003/12/12 SystemBrain Null�Ή�
        .HWFBM1SZ = rs("HWFBM1SZ")                      ' �i�v�e�a�l�c�P�������
        .HWFBM2AN = fncNullCheck(rs("HWFBM2AN"))        ' �i�v�e�a�l�c�Q���ω���        2003/12/12 SystemBrain Null�Ή�
        .HWFBM2AX = fncNullCheck(rs("HWFBM2AX"))        ' �i�v�e�a�l�c�Q���Ϗ��        2003/12/12 SystemBrain Null�Ή�
        .HWFBM2SH = rs("HWFBM2SH")                      ' �i�v�e�a�l�c�Q����ʒu�Q��
        .HWFBM2ST = rs("HWFBM2ST")                      ' �i�v�e�a�l�c�Q����ʒu�Q�_
        .HWFBM2SR = rs("HWFBM2SR")                      ' �i�v�e�a�l�c�Q����ʒu�Q��
        .HWFBM2HT = rs("HWFBM2HT")                      ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��
        .HWFBM2HS = rs("HWFBM2HS")                      ' �i�v�e�a�l�c�Q�ۏؕ��@�Q��
        .HWFBM2NS = rs("HWFBM2NS")                      ' �i�v�e�a�l�c�Q�M�����@
        .HWFBM2ET = fncNullCheck(rs("HWFBM2ET"))        ' �i�v�e�a�l�c�Q�I���d�s��      2003/12/12 SystemBrain Null�Ή�
        .HWFBM2SZ = rs("HWFBM2SZ")                      ' �i�v�e�a�l�c�Q�������
        .HWFBM3AN = fncNullCheck(rs("HWFBM3AN"))        ' �i�v�e�a�l�c�R���ω���        2003/12/12 SystemBrain Null�Ή�
        .HWFBM3AX = fncNullCheck(rs("HWFBM3AX"))        ' �i�v�e�a�l�c�R���Ϗ��        2003/12/12 SystemBrain Null�Ή�
        .HWFBM3SH = rs("HWFBM3SH")                      ' �i�v�e�a�l�c�R����ʒu�Q��
        .HWFBM3ST = rs("HWFBM3ST")                      ' �i�v�e�a�l�c�R����ʒu�Q�_
        .HWFBM3SR = rs("HWFBM3SR")                      ' �i�v�e�a�l�c�R����ʒu�Q��
        .HWFBM3HT = rs("HWFBM3HT")                      ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
        .HWFBM3HS = rs("HWFBM3HS")                      ' �i�v�e�a�l�c�R�ۏؕ��@�Q��
        .HWFBM3NS = rs("HWFBM3NS")                      ' �i�v�e�a�l�c�R�M�����@
        .HWFBM3ET = fncNullCheck(rs("HWFBM3ET"))        ' �i�v�e�a�l�c�R�I���d�s��      2003/12/12 SystemBrain Null�Ή�
        .HWFBM3SZ = rs("HWFBM3SZ")                      ' �i�v�e�a�l�c�R�������
        
        If Not IsNull(rs("HWFOSF1PTK")) Then .HWFOSF1PTK = rs("HWFOSF1PTK")   ' �i�v�e�n�r�e�P�p�^���敪
        If Not IsNull(rs("HWFOSF2PTK")) Then .HWFOSF2PTK = rs("HWFOSF2PTK")   ' �i�v�e�n�r�e�Q�p�^���敪
        If Not IsNull(rs("HWFOSF3PTK")) Then .HWFOSF3PTK = rs("HWFOSF3PTK")   ' �i�v�e�n�r�e�R�p�^���敪
        If Not IsNull(rs("HWFOSF4PTK")) Then .HWFOSF4PTK = rs("HWFOSF4PTK")   ' �i�v�e�n�r�e�S�p�^���敪
        
'        If Not IsNull(rs("HWFBM1MBP")) Then .HWFBM1MBP = rs("HWFBM1MBP")      ' �i�v�e�a�l�c�P�ʓ����z
'        If Not IsNull(rs("HWFBM2MBP")) Then .HWFBM2MBP = rs("HWFBM2MBP")      ' �i�v�e�a�l�c�Q�ʓ����z
'        If Not IsNull(rs("HWFBM3MBP")) Then .HWFBM3MBP = rs("HWFBM3MBP")      ' �i�v�e�a�l�c�R�ʓ����z
        .HWFBM1MBP = fncNullCheck(rs("HWFBM1MBP"))      ' �i�v�e�a�l�c�P�ʓ����z        2003/12/12 SystemBrain Null�Ή�
        .HWFBM2MBP = fncNullCheck(rs("HWFBM2MBP"))      ' �i�v�e�a�l�c�Q�ʓ����z        2003/12/12 SystemBrain Null�Ή�
        .HWFBM3MBP = fncNullCheck(rs("HWFBM3MBP"))      ' �i�v�e�a�l�c�R�ʓ����z        2003/12/12 SystemBrain Null�Ή�
        If Not IsNull(rs("HWFBM1MCL")) Then .HWFBM1MCL = rs("HWFBM1MCL")      ' �i�v�e�a�l�c�P�ʓ��v�Z
        If Not IsNull(rs("HWFBM2MCL")) Then .HWFBM2MCL = rs("HWFBM2MCL")      ' �i�v�e�a�l�c�Q�ʓ��v�Z
        If Not IsNull(rs("HWFBM3MCL")) Then .HWFBM3MCL = rs("HWFBM3MCL")      ' �i�v�e�a�l�c�R�ʓ��v�Z
    
    '���ύX �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        'AN���x�`�F�b�N�ׂ̈�TBCME025����i�v�e�`�m���x���擾����
        If Not IsNull(rs("HWFANTNP")) Then .HWFANTNP = rs("HWFANTNP")       ' �i�v�e�`�m���x
    '���ύX �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
    
    End With
    Set rs = Nothing

    funGet_TBCME029 = FUNCTION_RETURN_SUCCESS
  

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

'><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><<><><><
'�T�v      :��R����l�����߂�B
'���Ұ�    :�ϐ���         ,IO ,�^        ,����
'          :CRYNUM        ,I  ,String    ,�����ԍ�
'          :TopRs         ,I  ,Double    ,TOP�����茳��R����
'          :TopPos        ,I  ,Double    ,TOP�����茳�ʒu
'          :BotRs         ,I  ,Double    ,TOP�����茳��R����
'          :BotPos        ,I  ,Double    ,TOP�����茳�ʒu
'          :SuiPos        ,I  ,Double    ,����ʒu
'          :Suitei  �@    ,O  ,Double    ,����l
'          :�߂�l        ,O  ,FUNCTION_RETURN,
'����      :�����ԍ��ATOP/BOT�̒�R���ђl�A�ʒu����R������s���B
'����      :2003/9/4 �쐬  �}
Public Function new_ResSuitei(CRYNUM, TopRs, TOPPOS, BotRs, BOTPOS, SuiPos, Suitei As Double) As FUNCTION_RETURN
Dim cc As type_Coefficient  '���s�ΐ͌v�Z�p�\����
Dim rp As type_ResPosCal    '����v�Z�p�\����
Dim Jikouhen As Double  '���s�ΐ�
Dim wgtCharge As Long   '�`���[�W��
Dim wgtTop As Double    '�g�b�v�d�ʎ��ђl
Dim wgtTopCut As Double '�g�b�v�J�b�g�d�ʎ��ђl
Dim DM As Double        '���a�P�`�R�̕���
    
    new_ResSuitei = FUNCTION_RETURN_FAILURE
    
    ''���s�ΐ͗p�p�����[�^�擾 �}���`����Ή� �Q�Ɗ֐��ύX 2008/04/23 SETsw Nakada
    If GetCoeffParams_new(CRYNUM, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then
'    If GetCoeffParams(CRYNUM, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then
        Debug.Print "�ΐ͌v�Z�p�p�����[�^�̎擾�Ɏ��s����"
    End If
    
    ''�u���b�N�̎��s�ΐ͂����߂�
    cc.DUNMENSEKI = AreaOfCircle(DM)    '�f�ʐ�
    cc.CHARGEWEIGHT = wgtCharge         '�`���[�W��
    cc.TOPWEIGHT = wgtTop + wgtTopCut   '�g�b�v�d��
    cc.TOPSMPLPOS = TOPPOS
    cc.BOTSMPLPOS = BOTPOS
    cc.TOPRES = TopRs
    cc.BOTRES = BotRs
    
    Jikouhen = CoefficientCalculation(cc) '���s�ΐ͌v�Z
    
    
    ''�����R�l�����߂�
    If Jikouhen <> -9999 Then
        rp.COEFFICIENT = Jikouhen           '���s�ΐ�
        rp.DUNMENSEKI = cc.DUNMENSEKI       '�f�ʐ�
        rp.CHARGEWEIGHT = cc.CHARGEWEIGHT   '�`���[�W��
        rp.TOPWEIGHT = cc.TOPWEIGHT         '�g�b�v�d��
        rp.TOPSMPLPOS = TOPPOS
        rp.TOPRES = TopRs
        rp.target = SuiPos
        
        Suitei = ResCalculation(rp)         '����v�Z
    Else
        new_ResSuitei = FUNCTION_RETURN_FAILURE
        Exit Function
    End If
    
    new_ResSuitei = FUNCTION_RETURN_SUCCESS

End Function
'
''><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><<><><><
''�T�v      :�ΐ͌v�Z�ɕK�v�Ȋe���v�d�ʎ��т��擾����
''���Ұ�    :�ϐ���        ,IO ,�^        ,����
''          :CRYNUM        ,I  ,String    ,�����ԍ�
''          :wgtCharge     ,O  ,Long      ,�F���ʁi����`���[�W�ʁ|�O��܂ł̈��グ�d�ʁ|�O��܂ł�į�߶�ďd�ʁj
''          :wgtTop        ,O  ,Double    ,�g�b�v�d�ʎ��ђl
''          :wgtTopCut     ,O  ,Double    ,�g�b�v�J�b�g�d�ʎ��ђl
''          :DM            ,O  ,Double    ,���a�P�`�R�̕���
''          :�߂�l        ,O  ,FUNCTION_RETURN,
''����      :�P�{�����A�c�ʈ����ɂ��킹�Ď��уf�[�^���擾����
''����      :2001/8/29 �쐬  �쑺
'Public Function GetCoeffParams(ByVal CRYNUM$, wgtCharge As Long, wgtTop As Double, wgtTopCut As Double, DM As Double) As FUNCTION_RETURN
'Dim sql As String
'Dim rs As OraDynaset
'
'    On Error GoTo Err
'    GetCoeffParams = FUNCTION_RETURN_FAILURE
'    wgtCharge = 0
'    wgtTop = 0#
'    wgtTopCut = 0#
'    DM = 0#
'
'    sql = "select decode(RONAI,null,CHARGE,RONAI) as RONAI, WGHTTOP, WGTOPCUT, (DM1+DM2+DM3)/3.0 as DM " & _
'          "from TBCMH004 H004, " & _
'          "  (select sum(CHARGE) - sum(UPWEIGHT) - sum(WGTOPCUT) as RONAI" & _
'          "   From TBCMH004" & _
'          "   where (CRYNUM<'" & CRYNUM & "')" & _
'          "    and  (substr(CRYNUM,1,7)='" & Left$(CRYNUM, 7) & "')" & _
'          "  ) SUMDATA " & _
'          "where (CRYNUM='" & CRYNUM & "')"
'
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'    If rs.RecordCount > 0 Then
'        wgtCharge = rs("RONAI")
'        wgtTop = rs("WGHTTOP")
'        wgtTopCut = rs("WGTOPCUT")
'        DM = rs("DM")
'    End If
'    rs.Close
'
'    GetCoeffParams = FUNCTION_RETURN_SUCCESS
'
'proc_exit:
'    On Error GoTo 0
'    Exit Function
'
'Err:
'    Resume proc_exit
'End Function
'
''><><><><><><><><><><><><><><><><><><><><><>><><><><><><><><><><><><><><><><><><><><><><
''�T�v      :�ʒu�ɑ΂����R�l�𐄒肷��B
''���Ұ�    :�ϐ���        ,IO ,�^             ,����
''          :d             ,IO ,type_ResPosCal ,����v�Z�\����
''          :�߂�l        ,O  ,Double         ,�����R�l
''����      :
''����      :2001/06/23�@���� �M�Ɓ@�쐬
'Public Function ResCalculation(d As type_ResPosCal) As Double
'    Dim GS As Double
'    Dim Ro As Double
'    Dim Gx As Double
'
'    On Error GoTo Err
'    GS = (d.DUNMENSEKI * HIJU_SILICONE * d.TOPSMPLPOS) / (d.CHARGEWEIGHT - d.TOPWEIGHT)
'    Ro = d.TOPRES * (1 - GS) ^ (d.COEFFICIENT - 1)
'    Gx = d.DUNMENSEKI * d.target * HIJU_SILICONE / (d.CHARGEWEIGHT - d.TOPWEIGHT)
'
'    ResCalculation = Ro / (1 - Gx) ^ (d.COEFFICIENT - 1)
'    On Error GoTo 0
'    Exit Function
'Err:
'    On Error GoTo 0
'    ResCalculation = -9999
'End Function

'------------------------------------------------
' TBCME050�f�[�^�擾
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME050�v����w��i�Ԃ̃��R�[�h�𒊏o����B
'���Ұ�    :�ϐ���        ,IO ,�^                                   :����
'          :tHin          ,I  ,tFullHinban                          :�i��
'          :tGetRec       ,I  ,type_DBDRV_scmzc_fcmlc001c_Siyou     :���o���R�[�h
'    �@�@  :sErrMsg �@�@  ,O  ,String     �@�@�@�@�@�@�@�@�@�@�@    :�G���[���b�Z�[�W
'          :�߂�l        ,O  ,FUNCTION_RETURN                      :���o�̐���
'����      :
'����      :2006/08/15 �V�K�쐬 �G�s��s�]���ǉ��Ή� SMP)kondoh

Public Function funGet_TBCME050(tHIN As tFullHinban, _
                                tGetRec As type_DBDRV_scmzc_fcmlc001c_Siyou, _
                                Optional sErrMsg As String = vbNullString) As FUNCTION_RETURN

    Dim sql         As String           'SQL�S��
    Dim rs          As OraDynaset       'RecordSet
    Dim sDBName     As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_GetSiyou.bas -- Function funGet_TBCME050"

    sDBName = "E050"
    '�iEPBMD3���ω���(�O��),���Ϗ��(�O��)�ǉ��@09/05/07 ooba
    sql = "SELECT hinban, mnorevno, factory, opecond, hepantnp"
    sql = sql & " ,hepof1ax ,hepof1mx ,hepof1et ,hepof1ns ,hepof1sz ,hepof1sh ,hepof1st ,hepof1sr ,hepof1ht ,hepof1hs "
    sql = sql & " ,hepof1km ,hepof1kn ,hepof1kh ,hepof1ku ,heposf1ptk"
    sql = sql & " ,hepof2ax ,hepof2mx ,hepof2et ,hepof2ns ,hepof2sz ,hepof2sh ,hepof2st ,hepof2sr ,hepof2ht ,hepof2hs"
    sql = sql & " ,hepof2km ,hepof2kn ,hepof2kh ,hepof2ku ,heposf2ptk"
    sql = sql & " ,hepof3ax ,hepof3mx ,hepof3et ,hepof3ns ,hepof3sz ,hepof3sh ,hepof3st ,hepof3sr ,hepof3ht ,hepof3hs"
    sql = sql & " ,hepof3km ,hepof3kn ,hepof3kh ,hepof3ku ,heposf3ptk"
    sql = sql & " ,hepbm1an ,hepbm1ax ,hepbm1et ,hepbm1ns ,hepbm1sz ,hepbm1sh ,hepbm1st ,hepbm1sr ,hepbm1ht ,hepbm1hs"
    sql = sql & " ,hepbm1km ,hepbm1kn ,hepbm1kh ,hepbm1ku ,hepbm1mbp ,hepbm1mcl"
    sql = sql & " ,hepbm2an ,hepbm2ax ,hepbm2et ,hepbm2ns ,hepbm2sz ,hepbm2sh ,hepbm2st ,hepbm2sr ,hepbm2ht ,hepbm2hs"
    sql = sql & " ,hepbm2km ,hepbm2kn ,hepbm2kh ,hepbm2ku ,hepbm2mbp ,hepbm2mcl"
    sql = sql & " ,hepbm3an ,hepbm3ax ,hepbm3gsan ,hepbm3gsax ,hepbm3et ,hepbm3ns ,hepbm3sz ,hepbm3sh ,hepbm3st ,hepbm3sr ,hepbm3ht ,hepbm3hs"
    sql = sql & " ,hepbm3km ,hepbm3kn ,hepbm3kh ,hepbm3ku ,hepbm3mbp ,hepbm3mcl"
    sql = sql & " FROM tbcme050 "
    sql = sql & " WHERE hinban = '" & tHIN.hinban & "' and "
    sql = sql & "      mnorevno = " & tHIN.mnorevno & " and "
    sql = sql & "      factory = '" & tHIN.factory & "' and "
    sql = sql & "      opecond = '" & tHIN.opecond & "'"

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 1 Then
        Set rs = Nothing
        sErrMsg = GetMsgStr("EGET2", sDBName)
        funGet_TBCME050 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

     ''���o���ʂ��i�[����
     With tGetRec
        .HEPANTNP = fncNullCheck(rs("HEPANTNP"))                            ' �iEPAN���x
        .HEPOF1AX = fncNullCheck(rs("HEPOF1AX"))                            ' �iEPOSF1���Ϗ��
        .HEPOF1MX = fncNullCheck(rs("HEPOF1MX"))                            ' �iEPOSF1���
        .HEPOF1ET = fncNullCheck(rs("HEPOF1ET"))                            ' �iEPOSF1�I��ET��
        .HEPOF1NS = IIf(IsNull(rs("HEPOF1NS")), "", rs("HEPOF1NS"))         ' �iEPOSF1�M�����@
        .HEPOF1SZ = IIf(IsNull(rs("HEPOF1SZ")), "", rs("HEPOF1SZ"))         ' �iEPOSF1�������
        .HEPOF1SH = IIf(IsNull(rs("HEPOF1SH")), "", rs("HEPOF1SH"))         ' �iEPOSF1����ʒu_��
        .HEPOF1ST = IIf(IsNull(rs("HEPOF1ST")), "", rs("HEPOF1ST"))         ' �iEPOSF1����ʒu_�_
        .HEPOF1SR = IIf(IsNull(rs("HEPOF1SR")), "", rs("HEPOF1SR"))         ' �iEPOSF1����ʒu_��
        .HEPOF1HT = IIf(IsNull(rs("HEPOF1HT")), "", rs("HEPOF1HT"))         ' �iEPOSF1�ۏؕ��@_��
        .HEPOF1HS = IIf(IsNull(rs("HEPOF1HS")), "", rs("HEPOF1HS"))         ' �iEPOSF1�ۏؕ��@_��
        .HEPOF1KM = IIf(IsNull(rs("HEPOF1KM")), "", rs("HEPOF1KM"))         ' �iEPOSF1�����p�x_��
        .HEPOF1KN = IIf(IsNull(rs("HEPOF1KN")), "", rs("HEPOF1KN"))         ' �iEPOSF1�����p�x_��
        .HEPOF1KH = IIf(IsNull(rs("HEPOF1KH")), "", rs("HEPOF1KH"))         ' �iEPOSF1�����p�x_��
        .HEPOF1KU = IIf(IsNull(rs("HEPOF1KU")), "", rs("HEPOF1KU"))         ' �iEPOSF1�����p�x_�
        .HEPOSF1PTK = IIf(IsNull(rs("HEPOSF1PTK")), "", rs("HEPOSF1PTK"))   ' �iEPOSF1���݋敪
        .HEPOF2AX = fncNullCheck(rs("HEPOF2AX"))                            ' �iEPOSF2���Ϗ��
        .HEPOF2MX = fncNullCheck(rs("HEPOF2MX"))                            ' �iEPOSF2���
        .HEPOF2ET = fncNullCheck(rs("HEPOF2ET"))                            ' �iEPOSF2�I��ET��
        .HEPOF2NS = IIf(IsNull(rs("HEPOF2NS")), "", rs("HEPOF2NS"))         ' �iEPOSF2�M�����@
        .HEPOF2SZ = IIf(IsNull(rs("HEPOF2SZ")), "", rs("HEPOF2SZ"))         ' �iEPOSF2�������
        .HEPOF2SH = IIf(IsNull(rs("HEPOF2SH")), "", rs("HEPOF2SH"))         ' �iEPOSF2����ʒu_��
        .HEPOF2ST = IIf(IsNull(rs("HEPOF2ST")), "", rs("HEPOF2ST"))         ' �iEPOSF2����ʒu_�_
        .HEPOF2SR = IIf(IsNull(rs("HEPOF2SR")), "", rs("HEPOF2SR"))         ' �iEPOSF2����ʒu_��
        .HEPOF2HT = IIf(IsNull(rs("HEPOF2HT")), "", rs("HEPOF2HT"))         ' �iEPOSF2�ۏؕ��@_��
        .HEPOF2HS = IIf(IsNull(rs("HEPOF2HS")), "", rs("HEPOF2HS"))         ' �iEPOSF2�ۏؕ��@_��
        .HEPOF2KM = IIf(IsNull(rs("HEPOF2KM")), "", rs("HEPOF2KM"))         ' �iEPOSF2�����p�x_��
        .HEPOF2KN = IIf(IsNull(rs("HEPOF2KN")), "", rs("HEPOF2KN"))         ' �iEPOSF2�����p�x_��
        .HEPOF2KH = IIf(IsNull(rs("HEPOF2KH")), "", rs("HEPOF2KH"))         ' �iEPOSF2�����p�x_��
        .HEPOF2KU = IIf(IsNull(rs("HEPOF2KU")), "", rs("HEPOF2KU"))         ' �iEPOSF2�����p�x_�
        .HEPOSF2PTK = IIf(IsNull(rs("HEPOSF2PTK")), "", rs("HEPOSF2PTK"))   ' �iEPOSF2���݋敪
        .HEPOF3AX = fncNullCheck(rs("HEPOF3AX"))                            ' �iEPOSF3���Ϗ��
        .HEPOF3MX = fncNullCheck(rs("HEPOF3MX"))                            ' �iEPOSF3���
        .HEPOF3ET = fncNullCheck(rs("HEPOF3ET"))                            ' �iEPOSF3�I��ET��
        .HEPOF3NS = IIf(IsNull(rs("HEPOF3NS")), "", rs("HEPOF3NS"))         ' �iEPOSF3�M�����@
        .HEPOF3SZ = IIf(IsNull(rs("HEPOF3SZ")), "", rs("HEPOF3SZ"))         ' �iEPOSF3�������
        .HEPOF3SH = IIf(IsNull(rs("HEPOF3SH")), "", rs("HEPOF3SH"))         ' �iEPOSF3����ʒu_��
        .HEPOF3ST = IIf(IsNull(rs("HEPOF3ST")), "", rs("HEPOF3ST"))         ' �iEPOSF3����ʒu_�_
        .HEPOF3SR = IIf(IsNull(rs("HEPOF3SR")), "", rs("HEPOF3SR"))         ' �iEPOSF3����ʒu_��
        .HEPOF3HT = IIf(IsNull(rs("HEPOF3HT")), "", rs("HEPOF3HT"))         ' �iEPOSF3�ۏؕ��@_��
        .HEPOF3HS = IIf(IsNull(rs("HEPOF3HS")), "", rs("HEPOF3HS"))         ' �iEPOSF3�ۏؕ��@_��
        .HEPOF3KM = IIf(IsNull(rs("HEPOF3KM")), "", rs("HEPOF3KM"))         ' �iEPOSF3�����p�x_��
        .HEPOF3KN = IIf(IsNull(rs("HEPOF3KN")), "", rs("HEPOF3KN"))         ' �iEPOSF3�����p�x_��
        .HEPOF3KH = IIf(IsNull(rs("HEPOF3KH")), "", rs("HEPOF3KH"))         ' �iEPOSF3�����p�x_��
        .HEPOF3KU = IIf(IsNull(rs("HEPOF3KU")), "", rs("HEPOF3KU"))         ' �iEPOSF3�����p�x_�
        .HEPOSF3PTK = IIf(IsNull(rs("HEPOSF3PTK")), "", rs("HEPOSF3PTK"))   ' �iEPOSF3���݋敪
        .HEPBM1AN = fncNullCheck(rs("HEPBM1AN"))                            ' �iEPBMD1���ω���
        .HEPBM1AX = fncNullCheck(rs("HEPBM1AX"))                            ' �iEPBMD1���Ϗ��
        .HEPBM1ET = fncNullCheck(rs("HEPBM1ET"))                            ' �iEPBMD1�I��ET��
        .HEPBM1NS = IIf(IsNull(rs("HEPBM1NS")), "", rs("HEPBM1NS"))         ' �iEPBMD1�M�����@
        .HEPBM1SZ = IIf(IsNull(rs("HEPBM1SZ")), "", rs("HEPBM1SZ"))         ' �iEPBMD1�������
        .HEPBM1SH = IIf(IsNull(rs("HEPBM1SH")), "", rs("HEPBM1SH"))         ' �iEPBMD1����ʒu_��
        .HEPBM1ST = IIf(IsNull(rs("HEPBM1ST")), "", rs("HEPBM1ST"))         ' �iEPBMD1����ʒu_�_
        .HEPBM1SR = IIf(IsNull(rs("HEPBM1SR")), "", rs("HEPBM1SR"))         ' �iEPBMD1����ʒu_��
        .HEPBM1HT = IIf(IsNull(rs("HEPBM1HT")), "", rs("HEPBM1HT"))         ' �iEPBMD1�ۏؕ��@_��
        .HEPBM1HS = IIf(IsNull(rs("HEPBM1HS")), "", rs("HEPBM1HS"))         ' �iEPBMD1�ۏؕ��@_��
        .HEPBM1KM = IIf(IsNull(rs("HEPBM1KM")), "", rs("HEPBM1KM"))         ' �iEPBMD1�����p�x_��
        .HEPBM1KN = IIf(IsNull(rs("HEPBM1KN")), "", rs("HEPBM1KN"))         ' �iEPBMD1�����p�x_��
        .HEPBM1KH = IIf(IsNull(rs("HEPBM1KH")), "", rs("HEPBM1KH"))         ' �iEPBMD1�����p�x_��
        .HEPBM1KU = IIf(IsNull(rs("HEPBM1KU")), "", rs("HEPBM1KU"))         ' �iEPBMD1�����p�x_�
        .HEPBM1MBP = fncNullCheck(rs("HEPBM1MBP"))                          ' �iEPBMD1�ʓ����z
        .HEPBM1MCL = IIf(IsNull(rs("HEPBM1MCL")), "", rs("HEPBM1MCL"))      ' �iEPBMD1�ʓ��v�Z
        .HEPBM2AN = fncNullCheck(rs("HEPBM2AN"))                            ' �iEPBMD2���ω���
        .HEPBM2AX = fncNullCheck(rs("HEPBM2AX"))                            ' �iEPBMD2���Ϗ��
        .HEPBM2ET = fncNullCheck(rs("HEPBM2ET"))                            ' �iEPBMD2�I��ET��
        .HEPBM2NS = IIf(IsNull(rs("HEPBM2NS")), "", rs("HEPBM2NS"))         ' �iEPBMD2�M�����@
        .HEPBM2SZ = IIf(IsNull(rs("HEPBM2SZ")), "", rs("HEPBM2SZ"))         ' �iEPBMD2�������
        .HEPBM2SH = IIf(IsNull(rs("HEPBM2SH")), "", rs("HEPBM2SH"))         ' �iEPBMD2����ʒu_��
        .HEPBM2ST = IIf(IsNull(rs("HEPBM2ST")), "", rs("HEPBM2ST"))         ' �iEPBMD2����ʒu_�_
        .HEPBM2SR = IIf(IsNull(rs("HEPBM2SR")), "", rs("HEPBM2SR"))         ' �iEPBMD2����ʒu_��
        .HEPBM2HT = IIf(IsNull(rs("HEPBM2HT")), "", rs("HEPBM2HT"))         ' �iEPBMD2�ۏؕ��@_��
        .HEPBM2HS = IIf(IsNull(rs("HEPBM2HS")), "", rs("HEPBM2HS"))         ' �iEPBMD2�ۏؕ��@_��
        .HEPBM2KM = IIf(IsNull(rs("HEPBM2KM")), "", rs("HEPBM2KM"))         ' �iEPBMD2�����p�x_��
        .HEPBM2KN = IIf(IsNull(rs("HEPBM2KN")), "", rs("HEPBM2KN"))         ' �iEPBMD2�����p�x_��
        .HEPBM2KH = IIf(IsNull(rs("HEPBM2KH")), "", rs("HEPBM2KH"))         ' �iEPBMD2�����p�x_��
        .HEPBM2KU = IIf(IsNull(rs("HEPBM2KU")), "", rs("HEPBM2KU"))         ' �iEPBMD2�����p�x_�
        .HEPBM2MBP = fncNullCheck(rs("HEPBM2MBP"))                          ' �iEPBMD2�ʓ����z
        .HEPBM2MCL = IIf(IsNull(rs("HEPBM2MCL")), "", rs("HEPBM2MCL"))      ' �iEPBMD2�ʓ��v�Z
        .HEPBM3AN = fncNullCheck(rs("HEPBM3AN"))                            ' �iEPBMD3���ω���
        .HEPBM3AX = fncNullCheck(rs("HEPBM3AX"))                            ' �iEPBMD3���Ϗ��
        .HEPBM3GSAN = fncNullCheck(rs("HEPBM3GSAN"))                        ' �iEPBMD3���ω���(�O��)�@09/05/07 ooba
        .HEPBM3GSAX = fncNullCheck(rs("HEPBM3GSAX"))                        ' �iEPBMD3���Ϗ��(�O��)�@09/05/07 ooba
        .HEPBM3ET = fncNullCheck(rs("HEPBM3ET"))                            ' �iEPBMD3�I��ET��
        .HEPBM3NS = IIf(IsNull(rs("HEPBM3NS")), "", rs("HEPBM3NS"))         ' �iEPBMD3�M�����@
        .HEPBM3SZ = IIf(IsNull(rs("HEPBM3SZ")), "", rs("HEPBM3SZ"))         ' �iEPBMD3�������
        .HEPBM3SH = IIf(IsNull(rs("HEPBM3SH")), "", rs("HEPBM3SH"))         ' �iEPBMD3����ʒu_��
        .HEPBM3ST = IIf(IsNull(rs("HEPBM3ST")), "", rs("HEPBM3ST"))         ' �iEPBMD3����ʒu_�_
        .HEPBM3SR = IIf(IsNull(rs("HEPBM3SR")), "", rs("HEPBM3SR"))         ' �iEPBMD3����ʒu_��
        .HEPBM3HT = IIf(IsNull(rs("HEPBM3HT")), "", rs("HEPBM3HT"))         ' �iEPBMD3�ۏؕ��@_��
        .HEPBM3HS = IIf(IsNull(rs("HEPBM3HS")), "", rs("HEPBM3HS"))         ' �iEPBMD3�ۏؕ��@_��
        .HEPBM3KM = IIf(IsNull(rs("HEPBM3KM")), "", rs("HEPBM3KM"))         ' �iEPBMD3�����p�x_��
        .HEPBM3KN = IIf(IsNull(rs("HEPBM3KN")), "", rs("HEPBM3KN"))         ' �iEPBMD3�����p�x_��
        .HEPBM3KH = IIf(IsNull(rs("HEPBM3KH")), "", rs("HEPBM3KH"))         ' �iEPBMD3�����p�x_��
        .HEPBM3KU = IIf(IsNull(rs("HEPBM3KU")), "", rs("HEPBM3KU"))         ' �iEPBMD3�����p�x_�
        .HEPBM3MBP = fncNullCheck(rs("HEPBM3MBP"))                          ' �iEPBMD3�ʓ����z
        .HEPBM3MCL = IIf(IsNull(rs("HEPBM3MCL")), "", rs("HEPBM3MCL"))      ' �iEPBMD3�ʓ��v�Z
    End With
    Set rs = Nothing

    funGet_TBCME050 = FUNCTION_RETURN_SUCCESS
  
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

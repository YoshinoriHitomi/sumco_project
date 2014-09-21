Attribute VB_Name = "s_cmzcDBdriverCOM_SQL"
Option Explicit

' DB�h���C�o���ʊ֐�

'�T�v      :���グ�I�����сA�R�[�h�}�X�^�[����V�[�h�X�����擾
'���Ұ��@�@:�ϐ���      ,IO ,�^               ,����
'      �@�@:CRYNUM�@�@�@,I  ,String         �@,�����ԍ�
'      �@�@:SEED  �@�@�@,I  ,Integer        �@,�V�[�h�X��
'      �@�@:�߂�l      ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
'����      :
'����      :2001/07/12  �쐬 ���{
Public Function DBDRV_getSEED(ByVal CRYNUM As String, SEED As Integer) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_getSEED"

    sql = "select INFO3"
    sql = sql & " from TBCME037 H, TBCMB005 CM"
    sql = sql & " where H.CRYNUM='" & CRYNUM & "'"
    sql = sql & " and rtrim(CM.CODE,' ')=substr(H.SEED,1,1)"
    sql = sql & " and SYSCLASS='SC'"
    sql = sql & " and CLASS='28'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <= 0 Then
        DBDRV_getSEED = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    SEED = val(rs("INFO3"))
    rs.Close

    DBDRV_getSEED = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_getSEED = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :�������̑}��
'���Ұ��@�@:�ϐ���      ,IO ,�^               ,����
'      �@�@:CryInf�@�@�@,I  ,typ_TBCME037   �@,�������
'      �@�@:�߂�l      ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :
'����      :2001/07/12  �쐬 ���{
Public Function DBDRV_CryInf_Ins(CryInf As typ_TBCME037) As FUNCTION_RETURN

    Dim sql As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_CryInf_Ins"

    '' �������̑}��
    With CryInf
        sql = "insert into TBCME037 ("
        sql = sql & "CRYNUM, "              ' �����ԍ�
        sql = sql & "DELCLS, "              ' �폜�敪
        sql = sql & "KRPROCCD, "            ' �Ǘ��H���R�[�h
        sql = sql & "PROCCD, "              ' �H���R�[�h
        sql = sql & "LPKRPROCCD, "          ' �ŏI�ʉߊǗ��H��
        sql = sql & "LASTPASS, "            ' �ŏI�ʉߍH��
        sql = sql & "RPHINBAN, "            ' �˂炢�i��
        sql = sql & "RPREVNUM, "            ' �˂炢�i�Ԑ��i�ԍ������ԍ�
        sql = sql & "RPFACT, "              ' �˂炢�i�ԍH��
        sql = sql & "RPOPCOND, "            ' �˂炢�i�ԑ��Ə���
        sql = sql & "PRODCOND, "            ' �������
        sql = sql & "PGID, "                ' �o�f�|�h�c
        sql = sql & "UPLENGTH, "            ' ���グ����
        sql = sql & "TOPLENG, "             ' �s�n�o����
        sql = sql & "BODYLENG, "            ' ��������
        sql = sql & "BOTLENG, "             ' �a�n�s����
        sql = sql & "FREELENG, "            ' �t���[��
        sql = sql & "DIAMETER, "            ' ���a
        sql = sql & "CHARGE, "              ' �`���[�W��
        sql = sql & "SEED, "                ' �V�[�h
        sql = sql & "ADDDPCLS, "            ' �ǉ��h�[�v���
        sql = sql & "ADDDPPOS, "            ' �ǉ��h�[�v�ʒu
        sql = sql & "ADDDPVAL, "            ' �ǉ��h�[�v��
        sql = sql & "REGDATE, "             ' �o�^���t
        sql = sql & "UPDDATE, "             ' �X�V���t
        sql = sql & "SENDFLAG, "            ' ���M�t���O
        sql = sql & "SENDDATE)"             ' ���M���t
        sql = sql & " values ('"
        sql = sql & .CRYNUM & "', '"        ' �����ԍ�
        sql = sql & .DELCLS & "', '"        ' �폜�敪
        sql = sql & .KRPROCCD & "', '"      ' �Ǘ��H���R�[�h
        sql = sql & .PROCCD & "', '"        ' �H���R�[�h
        sql = sql & .LPKRPROCCD & "', '"    ' �ŏI�ʉߊǗ��H��
        sql = sql & .LASTPASS & "', '"      ' �ŏI�ʉߍH��
        sql = sql & .RPHINBAN & "', "       ' �˂炢�i��
        sql = sql & .RPREVNUM & ", '"       ' �˂炢�i�Ԑ��i�ԍ������ԍ�
        sql = sql & .RPFACT & "', '"        ' �˂炢�i�ԍH��
        sql = sql & .RPOPCOND & "', '"      ' �˂炢�i�ԑ��Ə���
        sql = sql & .PRODCOND & "', '"      ' �������
        sql = sql & .PGID & "', "           ' �o�f�|�h�c
        sql = sql & .UPLENGTH & ", "        ' ���グ����
        sql = sql & .TOPLENG & ", "         ' �s�n�o����
        sql = sql & .BODYLENG & ", "        ' ��������
        sql = sql & .BOTLENG & ", "         ' �a�n�s����
        sql = sql & .FREELENG & ", "        ' �t���[��
        sql = sql & .DIAMETER & ", "        ' ���a
        sql = sql & .CHARGE & ", '"         ' �`���[�W��
        sql = sql & .SEED & "', '"          ' �V�[�h
        sql = sql & .ADDDPCLS & "', "       ' �ǉ��h�[�v���
        sql = sql & .ADDDPPOS & ", "        ' �ǉ��h�[�v�ʒu
        sql = sql & .ADDDPVAL & ", "        ' �ǉ��h�[�v��
        sql = sql & "sysdate, "             ' �o�^���t
        sql = sql & "sysdate, "             ' �X�V���t
        sql = sql & "'0', "                 ' ���M�t���O
        sql = sql & "sysdate)"              ' ���M���t
    End With
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_CryInf_Ins = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_CryInf_Ins = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_CryInf_Ins = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'�T�v      :�������̍X�V
'���Ұ��@�@:�ϐ���      ,IO ,�^               ,����
'      �@�@:CryInf�@�@�@,I  ,typ_TBCME037   �@,�������
'      �@�@:�߂�l      ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :
'-------�g�p���Ȃ��ق����悢�i���{�j---------
Public Function DBDRV_CryInf_Upd(CryInf As typ_TBCME037) As FUNCTION_RETURN

    Dim sql As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_CryInf_Upd"

    '' �������̍X�V
    With CryInf
        sql = "update TBCME037 set "
        sql = sql & "CRYNUM='" & .CRYNUM & "', "            ' �����ԍ�
        sql = sql & "DELCLS='" & .DELCLS & "', "            ' �폜�敪
        sql = sql & "KRPROCCD='" & .KRPROCCD & "', "        ' �Ǘ��H���R�[�h
        sql = sql & "PROCCD='" & .PROCCD & "', "            ' �H���R�[�h
        sql = sql & "LPKRPROCCD='" & .LPKRPROCCD & "', "    ' �ŏI�ʉߊǗ��H��
        sql = sql & "LASTPASS='" & .LASTPASS & "', "        ' �ŏI�ʉߍH��
        sql = sql & "RPHINBAN='" & .RPHINBAN & "', "        ' �˂炢�i��
        sql = sql & "RPREVNUM=" & .RPREVNUM & ", "          ' �˂炢�i�Ԑ��i�ԍ������ԍ�
        sql = sql & "RPFACT='" & .RPFACT & "', "            ' �˂炢�i�ԍH��
        sql = sql & "RPOPCOND='" & .RPOPCOND & "', "        ' �˂炢�i�ԑ��Ə���
        sql = sql & "PRODCOND='" & .PRODCOND & "', "        ' �������
        sql = sql & "PGID='" & .PGID & "', "                ' �o�f�|�h�c
        sql = sql & "UPLENGTH=" & .UPLENGTH & ", "          ' ���グ����
        sql = sql & "TOPLENG=" & .TOPLENG & ", "            ' �s�n�o����
        sql = sql & "BODYLENG=" & .BODYLENG & ", "          ' ��������
        sql = sql & "BOTLENG=" & .BOTLENG & ", "            ' �a�n�s����
        sql = sql & "FREELENG=" & .FREELENG & ", "          ' �t���[��
        sql = sql & "DIAMETER=" & .DIAMETER & ", "          ' ���a
        sql = sql & "CHARGE=" & .CHARGE & ", "              ' �`���[�W��
        sql = sql & "SEED='" & .SEED & "', "                ' �V�[�h
        sql = sql & "ADDDPCLS='" & .ADDDPCLS & "', "        ' �ǉ��h�[�v���
        sql = sql & "ADDDPPOS=" & .ADDDPPOS & ", "          ' �ǉ��h�[�v�ʒu
        sql = sql & "ADDDPVAL=" & .ADDDPVAL & ", "          ' �ǉ��h�[�v��
        sql = sql & "UPDDATE=sysdate, "                     ' �X�V���t
        sql = sql & "SENDFLAG='0', "                        ' ���M�t���O
        sql = sql & "SENDDATE=sysdate"                      ' ���M���t
        sql = sql & " where CRYNUM='" & .CRYNUM & "'"
    End With
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_CryInf_Upd = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_CryInf_Upd = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_CryInf_Upd = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'�T�v      :�u���b�N�Ǘ��̑}���^�X�V
'���Ұ��@�@:�ϐ���           ,IO ,�^               ,����
'      �@�@:BlockMngOld�@�@�@,I  ,typ_TBCME040   �@,�u���b�N�Ǘ��i���j
'      �@�@:BlockMngNew�@�@�@,I  ,typ_TBCME040   �@,�u���b�N�Ǘ��i�V�j
'      �@�@:�߂�l           ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :�Â����R�[�h���݂čX�V���}�����𔻕ʂ���
'����      :2001/07/12  �쐬 ���{
Public Function DBDRV_BlockMng_UpdIns(BlockMngOld() As typ_TBCME040, BlockMngNew() As typ_TBCME040) As FUNCTION_RETURN

    Dim sql As String
    Dim lFlg As Boolean
    Dim i As Long
    Dim j As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_BlockMng_UpdIns"

    DBDRV_BlockMng_UpdIns = FUNCTION_RETURN_SUCCESS

    For i = 1 To UBound(BlockMngNew)
        With BlockMngNew(i)
            lFlg = False
            For j = 1 To UBound(BlockMngOld)
                If BlockMngOld(j).CRYNUM = .CRYNUM And _
                   BlockMngOld(j).INGOTPOS = .INGOTPOS Then
                    '' �u���b�N�Ǘ��e�[�u���̍X�V
                    sql = "update TBCME040 set "
                    sql = sql & "CRYNUM='" & .CRYNUM & "', "                    ' �����ԍ�
                    sql = sql & "INGOTPOS=" & .INGOTPOS & ", "                  ' �������J�n�ʒu
                    sql = sql & "LENGTH=" & .Length & ", "                      ' ����
                    sql = sql & "REALLEN=" & .REALLEN & ", "                    ' ������
                    sql = sql & "BLOCKID='" & .BLOCKID & "', "                  ' �u���b�NID
                    sql = sql & "KRPROCCD='" & .KRPROCCD & "', "                ' ���݊Ǘ��H��
                    sql = sql & "NOWPROC='" & .NOWPROC & "', "                  ' ���ݍH��
                    sql = sql & "LPKRPROCCD='" & .LPKRPROCCD & "', "            ' �ŏI�ʉߊǗ��H��
                    sql = sql & "LASTPASS='" & .LASTPASS & "', "                ' �ŏI�ʉߍH��
                    sql = sql & "DELCLS='" & .DELCLS & "', "                    ' �폜�敪
                    sql = sql & "LSTATCLS='" & .LSTATCLS & "', "                ' �ŏI��ԋ敪
                    sql = sql & "RSTATCLS='" & .RSTATCLS & "', "                ' ������ԋ敪
                    sql = sql & "HOLDCLS='" & .HOLDCLS & "', "                  ' �z�[���h�敪
                    sql = sql & "BDCAUS='" & .BDCAUS & "', "                    ' �s�Ǘ��R
                    sql = sql & "UPDDATE=sysdate, "                             ' �X�V���t
                    sql = sql & "SUMMITSENDFLAG='" & .SUMMITSENDFLAG & "', "    ' SUMMIT���M�t���O
                    sql = sql & "SENDFLAG='0', "                                ' ���M�t���O
                    sql = sql & "SENDDATE=sysdate "                             ' ���M���t
                    sql = sql & "where CRYNUM='" & .CRYNUM & "' "
                    sql = sql & "and INGOTPOS=" & .INGOTPOS
                    '' WriteDBLog sql
                    If OraDB.ExecuteSQL(sql) <= 0 Then
                        DBDRV_BlockMng_UpdIns = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                    lFlg = True
                    Exit For
                End If
            Next j

            If lFlg <> True Then
                '' �u���b�N�Ǘ��e�[�u���̑}��
                sql = "insert into TBCME040 ("
                sql = sql & "CRYNUM, "              ' �����ԍ�
                sql = sql & "INGOTPOS, "            ' �������J�n�ʒu
                sql = sql & "LENGTH, "              ' ����
                sql = sql & "REALLEN, "             ' ������
                sql = sql & "BLOCKID, "             ' �u���b�NID
                sql = sql & "KRPROCCD, "            ' ���݊Ǘ��H��
                sql = sql & "NOWPROC, "             ' ���ݍH��
                sql = sql & "LPKRPROCCD, "          ' �ŏI�ʉߊǗ��H��
                sql = sql & "LASTPASS, "            ' �ŏI�ʉߍH��
                sql = sql & "DELCLS, "              ' �폜�敪
                sql = sql & "LSTATCLS, "            ' �ŏI��ԋ敪
                sql = sql & "RSTATCLS, "            ' ������ԋ敪
                sql = sql & "HOLDCLS, "             ' �z�[���h�敪
                sql = sql & "BDCAUS, "              ' �s�Ǘ��R
                sql = sql & "REGDATE, "             ' �o�^���t
                sql = sql & "UPDDATE, "             ' �X�V���t
                sql = sql & "SUMMITSENDFLAG, "      ' SUMMIT���M�t���O
                sql = sql & "SENDFLAG, "            ' ���M�t���O
                sql = sql & "SENDDATE)"             ' ���M���t
                sql = sql & " values ('"
                sql = sql & .CRYNUM & "', "         ' �����ԍ�
                sql = sql & .INGOTPOS & ", "        ' �������J�n�ʒu
                sql = sql & .Length & ", "          ' ����
                sql = sql & .REALLEN & ", '"        ' ������
                sql = sql & .BLOCKID & "', '"       ' �u���b�NID
                sql = sql & .KRPROCCD & "', '"      ' ���݊Ǘ��H��
                sql = sql & .NOWPROC & "', '"       ' ���ݍH��
                sql = sql & .LPKRPROCCD & "', '"    ' �ŏI�ʉߊǗ��H��
                sql = sql & .LASTPASS & "', '"      ' �ŏI�ʉߍH��
                sql = sql & .DELCLS & "', '"        ' �폜�敪
                sql = sql & .LSTATCLS & "', '"      ' �ŏI��ԋ敪
                sql = sql & .RSTATCLS & "', '"      ' ������ԋ敪
                sql = sql & .HOLDCLS & "', '"       ' �z�[���h�敪
                sql = sql & .BDCAUS & "', "         ' �s�Ǘ��R
                sql = sql & "sysdate, "             ' �o�^���t
                sql = sql & "sysdate, '"            ' �X�V���t
                sql = sql & .SUMMITSENDFLAG & "', " ' SUMMIT���M�t���O
                sql = sql & "'0', "                 ' ���M�t���O
                sql = sql & "sysdate)"              ' ���M���t
                '' WriteDBLog sql
                If OraDB.ExecuteSQL(sql) <= 0 Then
                    DBDRV_BlockMng_UpdIns = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            End If
        End With
    Next i

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_BlockMng_UpdIns = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'�T�v      :�u���b�N�Ǘ��̑}��
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:BlockMng�@�@�@,I  ,typ_TBCME040   �@,�u���b�N�Ǘ�
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :
'����      :2001/07/12  �쐬 ���{
Public Function DBDRV_BlockMng_Ins(BlockMng As typ_TBCME040) As FUNCTION_RETURN

    Dim sql As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_BlockMng_Ins"

    With BlockMng
        sql = "insert into TBCME040 ("
        sql = sql & "CRYNUM, "              ' �����ԍ�
        sql = sql & "INGOTPOS, "            ' �������J�n�ʒu
        sql = sql & "LENGTH, "              ' ����
        sql = sql & "REALLEN, "             ' ������
        sql = sql & "BLOCKID, "             ' �u���b�NID
        sql = sql & "KRPROCCD, "            ' ���݊Ǘ��H��
        sql = sql & "NOWPROC, "             ' ���ݍH��
        sql = sql & "LPKRPROCCD, "          ' �ŏI�ʉߊǗ��H��
        sql = sql & "LASTPASS, "            ' �ŏI�ʉߍH��
        sql = sql & "DELCLS, "              ' �폜�敪
        sql = sql & "LSTATCLS, "            ' �ŏI��ԋ敪
        sql = sql & "RSTATCLS, "            ' ������ԋ敪
        sql = sql & "HOLDCLS, "             ' �z�[���h�敪
        sql = sql & "BDCAUS, "              ' �s�Ǘ��R
        sql = sql & "REGDATE, "             ' �o�^���t
        sql = sql & "UPDDATE, "             ' �X�V���t
        sql = sql & "SUMMITSENDFLAG, "      ' SUMMIT���M�t���O
        sql = sql & "SENDFLAG, "            ' ���M�t���O
        sql = sql & "SENDDATE)"             ' ���M���t
        sql = sql & " values ('"
        sql = sql & .CRYNUM & "', "         ' �����ԍ�
        sql = sql & .INGOTPOS & ", "        ' �������J�n�ʒu
        sql = sql & .Length & ", "          ' ����
        sql = sql & .REALLEN & ", '"        ' ������
        sql = sql & .BLOCKID & "', '"       ' �u���b�NID
        sql = sql & .KRPROCCD & "', '"      ' ���݊Ǘ��H��
        sql = sql & .NOWPROC & "', '"       ' ���ݍH��
        sql = sql & .LPKRPROCCD & "', '"    ' �ŏI�ʉߊǗ��H��
        sql = sql & .LASTPASS & "', '"      ' �ŏI�ʉߍH��
        sql = sql & .DELCLS & "', '"        ' �폜�敪
        sql = sql & .LSTATCLS & "', '"      ' �ŏI��ԋ敪
        sql = sql & .RSTATCLS & "', '"      ' ������ԋ敪
        sql = sql & .HOLDCLS & "', '"       ' �z�[���h�敪
        sql = sql & .BDCAUS & "', "         ' �s�Ǘ��R
        sql = sql & "sysdate, "             ' �o�^���t
        sql = sql & "sysdate, "             ' �X�V���t
        sql = sql & "'0', "                 ' SUMMIT���M�t���O
        sql = sql & "'0', "                 ' ���M�t���O
        sql = sql & "sysdate)"              ' ���M���t
    End With
    '' WriteDBLog sql
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_BlockMng_Ins = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_BlockMng_Ins = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_BlockMng_Ins = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'�T�v      :�u���b�N�Ǘ��̍X�V
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:BlockMng�@�@�@,I  ,typ_TBCME040   �@,�u���b�N�Ǘ�
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :
'-------�g�p���Ȃ��ق����悢�i���{�j---------
Public Function DBDRV_BlockMng_Upd(BlockMng As typ_TBCME040) As FUNCTION_RETURN

    Dim sql As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_BlockMng_Upd"

    '' �u���b�N�Ǘ��e�[�u���̍X�V
    With BlockMng
        sql = "update TBCME040 set "
        sql = sql & "LENGTH=" & .Length & ", "              ' ����
        sql = sql & "REALLEN=" & .REALLEN & ", "            ' ������
        sql = sql & "BLOCKID='" & .BLOCKID & "', "          ' �u���b�NID
        sql = sql & "KRPROCCD='" & .KRPROCCD & "', "        ' ���݊Ǘ��H��
        sql = sql & "NOWPROC='" & .NOWPROC & "', "          ' ���ݍH��
        sql = sql & "LPKRPROCCD='" & .LPKRPROCCD & "', "    ' �ŏI�ʉߊǗ��H��
        sql = sql & "LASTPASS='" & .LASTPASS & "', "        ' �ŏI�ʉߍH��
        sql = sql & "DELCLS='" & .DELCLS & "',"             ' �폜�敪
        sql = sql & "LSTATCLS='" & .LSTATCLS & "', "        ' �ŏI��ԋ敪
        sql = sql & "RSTATCLS='" & .RSTATCLS & "', "        ' ������ԋ敪
        sql = sql & "HOLDCLS='" & .HOLDCLS & "', "          ' �z�[���h�敪
        sql = sql & "BDCAUS='" & .BDCAUS & "', "            ' �s�Ǘ��R
        sql = sql & "UPDDATE=sysdate, "                     ' �X�V���t
        sql = sql & "SUMMITSENDFLAG='0', "                  ' SUMMIT���M�t���O
        sql = sql & "SENDFLAG='0' "                        ' ���M�t���O
        sql = sql & "where CRYNUM='" & .CRYNUM & "' "
        sql = sql & "and INGOTPOS=" & .INGOTPOS
    End With
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_BlockMng_Upd = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_BlockMng_Upd = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_BlockMng_Upd = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'�T�v      :�i�ԊǗ��̑}���^�X�V
'���Ұ��@�@:�ϐ���         ,IO ,�^               ,����
'      �@�@:HinbanOld�@�@�@,I  ,typ_TBCME041   �@,�i�ԊǗ��i���j
'      �@�@:HinbanNew�@�@�@,I  ,typ_TBCME041   �@,�i�ԊǗ��i�V�j
'      �@�@:�߂�l         ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :�Â����R�[�h���݂čX�V���}�����𔻕ʂ���
'����      :2001/07/12  �쐬 ���{
'      �@�@:2001/11/06  �C�� �쑺
Public Function DBDRV_Hinban_UpdIns(HinbanOld() As typ_TBCME041, HinbanNew() As typ_TBCME041) As FUNCTION_RETURN

    Dim sql As String
    Dim lFlg As Boolean
    Dim i As Long
    Dim j As Long
    Dim nOld As Long
    Dim HIN As tFullHinban

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_Hinban_UpdIns"

    DBDRV_Hinban_UpdIns = FUNCTION_RETURN_SUCCESS

    nOld = UBound(HinbanOld)
    For i = 1 To UBound(HinbanNew)
        With HinbanNew(i)
            For j = 1 To nOld
                If (.INGOTPOS = HinbanOld(j).INGOTPOS) _
                  And (.Length = HinbanOld(j).Length) _
                  And (.hinban = HinbanOld(j).hinban) _
                  And (.REVNUM = HinbanOld(j).REGDATE) _
                  And (.factory = HinbanOld(j).factory) _
                  And (.opecond = HinbanOld(j).opecond) Then
                    '�S�������e�̃��R�[�h�����łɗL��
                    Exit For
                End If
           Next
           If j > nOld Then '�����e�̃��R�[�h�͂Ȃ�����
                '�����ύX������΁A���͈̔͂̕i�Ԃ�u��������
                HIN.hinban = .hinban
                HIN.mnorevno = .REVNUM
                HIN.factory = .factory
                HIN.opecond = .opecond
                If ChangeAreaHinban(.CRYNUM, .INGOTPOS, .Length, HIN) = FUNCTION_RETURN_FAILURE Then
                    DBDRV_Hinban_UpdIns = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            End If
        End With
    Next

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_Hinban_UpdIns = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'�T�v      :�i�ԊǗ��̑}��
'���Ұ��@�@:�ϐ���         ,IO ,�^               ,����
'      �@�@:HinbanNew�@�@�@,I  ,typ_TBCME041   �@,�i�ԊǗ�
'      �@�@:�߂�l         ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :
'����      :2001/07/12  �쐬 ���{
Public Function DBDRV_Hinban_Ins(HinbanNew() As typ_TBCME041) As FUNCTION_RETURN

    Dim sql As String
    Dim lFlg As Boolean
    Dim i As Long
    Dim j As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_Hinban_Ins"

    DBDRV_Hinban_Ins = FUNCTION_RETURN_SUCCESS

    For i = 1 To UBound(HinbanNew)
        With HinbanNew(i)
            sql = "insert into TBCME041 ("
            sql = sql & "CRYNUM, "          ' �����ԍ�
            sql = sql & "INGOTPOS, "        ' �������J�n�ʒu
            sql = sql & "HINBAN, "          ' �i��
            sql = sql & "REVNUM, "          ' ���i�ԍ������ԍ�
            sql = sql & "FACTORY, "         ' �H��
            sql = sql & "OPECOND, "         ' ���Ə���
            sql = sql & "LENGTH, "          ' ����
            sql = sql & "REGDATE, "         ' �o�^���t
            sql = sql & "UPDDATE, "         ' �X�V���t
            sql = sql & "SENDFLAG, "        ' ���M�t���O
            sql = sql & "SENDDATE)"         ' ���M���t
            sql = sql & " values ('"
            sql = sql & .CRYNUM & "', "     ' �����ԍ�
            sql = sql & .INGOTPOS & ", '"   ' �������J�n�ʒu
            sql = sql & .hinban & "', "     ' �i��
            sql = sql & .REVNUM & ", '"     ' ���i�ԍ������ԍ�
            sql = sql & .factory & "', '"   ' �H��
            sql = sql & .opecond & "', "    ' ���Ə���
            sql = sql & .Length & ", "      ' ����
            sql = sql & "sysdate, "         ' �o�^���t
            sql = sql & "sysdate, "         ' �X�V���t
            sql = sql & "'0', "             ' ���M�t���O
            sql = sql & "sysdate)"          ' ���M���t
        End With
        '' WriteDBLog sql
        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_Hinban_Ins = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
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
    DBDRV_Hinban_Ins = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'�T�v      :�����T���v���Ǘ��̑}���^�X�V
'���Ұ��@�@:�ϐ���         ,IO ,�^                  ,����
'      �@�@:CrySmpOld�@�@�@,I  ,typ_XSDCS   �@      ,�V�T���v���Ǘ��i�u���b�N�j�i���j
'      �@�@:CrySmpNew�@�@�@,I  ,typ_XSDCS   �@      ,�V�T���v���Ǘ��i�u���b�N�j�i�V�j
'      �@�@:�߂�l         ,O  ,FUNCTION_RETURN�@   ,�������݂̐���
'����      :�Â����R�[�h���݂čX�V���}�����𔻕ʂ���
'����      :2001/07/12  �쐬 ���{
Public Function DBDRV_CrySmp_UpdIns(CrySmpOld() As typ_XSDCS, CrySmpNew() As typ_XSDCS) As FUNCTION_RETURN

    Dim sql As String
    Dim lFlg As Boolean
    Dim i As Long
    Dim j As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_CrySmp_UpdIns"

    DBDRV_CrySmp_UpdIns = FUNCTION_RETURN_SUCCESS

    For i = 1 To UBound(CrySmpNew)
        With CrySmpNew(i)
            lFlg = False
            For j = 1 To UBound(CrySmpOld)
                If CrySmpOld(j).XTALCS = .XTALCS And _
                   CrySmpOld(j).INPOSCS = .INPOSCS And _
                   CrySmpOld(j).SMPKBNCS = .SMPKBNCS Then
'                    sql = "update TBCME043 set "
'                    sql = sql & "HINBAN='" & .HINBAN & "', "        ' �i��
'                    sql = sql & "REVNUM=" & .REVNUM & ", "          ' ���i�ԍ������ԍ�
'                    sql = sql & "FACTORY='" & .factory & "', "      ' �H��
'                    sql = sql & "OPECOND='" & .opecond & "', "      ' ���Ə���
'                    sql = sql & "KTKBN='" & .KTKBN & "', "          ' �m��敪
'                    sql = sql & "SMPLNO='" & Abs(.SMPLNO) & "', "   ' �T���v���m��
'                    sql = sql & "CRYINDRS='" & .CRYINDRS & "', "    ' ���������w���iRs)
'                    sql = sql & "CRYINDOI='" & .CRYINDOI & "', "    ' ���������w���iOi)
'                    sql = sql & "CRYINDB1='" & .CRYINDB1 & "', "    ' ���������w���iB1)
'                    sql = sql & "CRYINDB2='" & .CRYINDB2 & "', "    ' ���������w���iB2)
'                    sql = sql & "CRYINDB3='" & .CRYINDB3 & "', "    ' ���������w���iB3)
'                    sql = sql & "CRYINDL1='" & .CRYINDL1 & "', "    ' ���������w���iL1)
'                    sql = sql & "CRYINDL2='" & .CRYINDL2 & "', "    ' ���������w���iL2)
'                    sql = sql & "CRYINDL3='" & .CRYINDL3 & "', "    ' ���������w���iL3)
'                    sql = sql & "CRYINDL4='" & .CRYINDL4 & "', "    ' ���������w���iL4)
'                    sql = sql & "CRYINDCS='" & .CRYINDCS & "', "    ' ���������w���iCs)
'                    sql = sql & "CRYINDGD='" & .CRYINDGD & "', "    ' ���������w���iGD)
'                    sql = sql & "CRYINDT='" & .CRYINDT & "', "      ' ���������w���iT)
'                    sql = sql & "CRYINDEP='" & .CRYINDEP & "', "    ' ���������w���iEPD)
'                    sql = sql & "CRYRESRS='" & .CRYRESRS & "', "    ' �����������сiRs)
'                    sql = sql & "CRYRESOI='" & .CRYRESOI & "', "    ' �����������сiOi)
'                    sql = sql & "CRYRESB1='" & .CRYRESB1 & "', "    ' �����������сiB1)
'                    sql = sql & "CRYRESB2='" & .CRYRESB2 & "', "    ' �����������сiB2)
'                    sql = sql & "CRYRESB3='" & .CRYRESB3 & "', "    ' �����������сiB3)
'                    sql = sql & "CRYRESL1='" & .CRYRESL1 & "', "    ' �����������сiL1)
'                    sql = sql & "CRYRESL2='" & .CRYRESL2 & "', "    ' �����������сiL2)
'                    sql = sql & "CRYRESL3='" & .CRYRESL3 & "', "    ' �����������сiL3)
'                    sql = sql & "CRYRESL4='" & .CRYRESL4 & "', "    ' �����������сiL4)
'                    sql = sql & "CRYRESCS='" & .CRYRESCS & "', "    ' �����������сiCs)
'                    sql = sql & "CRYRESGD='" & .CRYRESGD & "', "    ' �����������сiGD)
'                    sql = sql & "CRYREST='" & .CRYREST & "', "      ' �����������сiT)
'                    sql = sql & "CRYRESEP='" & .CRYRESEP & "', "    ' �����������сiEPD)
'                    sql = sql & "SMPLNUM=" & .SMPLNUM & ", "        ' �T���v������
'                    sql = sql & "SMPLPAT='" & .SMPLPAT & "', "      ' �T���v���p�^�[��
'                    sql = sql & "UPDDATE=sysdate, "                 ' �X�V���t
'                    sql = sql & "SENDFLAG='0' "                     ' ���M�t���O
'                    sql = sql & " where CRYNUM='" & .CRYNUM & "'"
'                    sql = sql & " and INGOTPOS=" & .INGOTPOS
'                    sql = sql & " and SMPKBN='" & .SMPKBN & "'"
                    sql = "update XSDCS set "
                    sql = sql & "CRYNUMCS='" & .CRYNUMCS & "', "            ' �u���b�NID
                    sql = sql & "TBKBNCS='" & .TBKBNCS & "', "              ' T/B�敪
                    sql = sql & "REPSMPLIDCS='" & Abs(.REPSMPLIDCS) & "', " ' ��\�T���v��ID
                    sql = sql & "HINBCS='" & .HINBCS & "', "                ' �i��
                    sql = sql & "REVNUMCS=" & .REVNUMCS & ", "              ' ���i�ԍ������ԍ�
                    sql = sql & "FACTORYCS='" & .FACTORYCS & "', "          ' �H��
                    sql = sql & "OPECS='" & .OPECS & "', "                  ' ���Ə���
                    sql = sql & "KTKBNCS='" & .KTKBNCS & "', "              ' �m��敪
                    sql = sql & "BLKKTFLAGCS='" & .BLKKTFLAGCS & "', "      ' �u���b�N�m��t���O
                    sql = sql & "CRYSMPLIDRSCS=" & .CRYSMPLIDRSCS & ", "    ' �T���v��ID(Rs)
                    sql = sql & "CRYSMPLIDRS1CS=" & .CRYSMPLIDRS1CS & ", "  ' ����T���v��ID1�iRs�j
                    sql = sql & "CRYSMPLIDRS2CS=" & .CRYSMPLIDRS2CS & ", "  ' ����T���v��ID2�iRs�j
                    sql = sql & "CRYSMPLIDOICS=" & .CRYSMPLIDOICS & ", "    ' �T���v��ID(Oi)
                    sql = sql & "CRYSMPLIDB1CS=" & .CRYSMPLIDB1CS & ", "    ' �T���v��ID(B1)
                    sql = sql & "CRYSMPLIDB2CS=" & .CRYSMPLIDB2CS & ", "    ' �T���v��ID(B2)
                    sql = sql & "CRYSMPLIDB3CS=" & .CRYSMPLIDB3CS & ", "    ' �T���v��ID(B3)
                    sql = sql & "CRYSMPLIDL1CS=" & .CRYSMPLIDL1CS & ", "    ' �T���v��ID(L1)
                    sql = sql & "CRYSMPLIDL2CS=" & .CRYSMPLIDL2CS & ", "    ' �T���v��ID(L2)
                    sql = sql & "CRYSMPLIDL3CS=" & .CRYSMPLIDL3CS & ", "    ' �T���v��ID(L3)
                    sql = sql & "CRYSMPLIDL4CS=" & .CRYSMPLIDL4CS & ", "    ' �T���v��ID(L4)
                    sql = sql & "CRYSMPLIDCSCS=" & .CRYSMPLIDCSCS & ", "    ' �T���v��ID(Cs)
                    sql = sql & "CRYSMPLIDGDCS=" & .CRYSMPLIDGDCS & ", "    ' �T���v��ID(GD)
                    sql = sql & "CRYSMPLIDTCS=" & .CRYSMPLIDTCS & ", "      ' �T���v��ID(T)
                    sql = sql & "CRYSMPLIDEPCS=" & .CRYSMPLIDEPCS & ", "    ' �T���v��ID(EPD)
                    'Cng Start 2011/03/31 SMPK Y.Hitomi
'                    'Add Start 2010/12/13 SMPK Miyata
'                    sql = sql & "CRYSMPLIDCCS=" & .CRYSMPLIDCCS & ", "      ' �T���v��ID(C)
'                    sql = sql & "CRYSMPLIDCJCS=" & .CRYSMPLIDCJCS & ", "    ' �T���v��ID(CJ)
'                    sql = sql & "CRYSMPLIDCJLTCS=" & .CRYSMPLIDCJLTCS & ", " ' �T���v��ID(CJLT)
'                    sql = sql & "CRYSMPLIDCJ2CS=" & .CRYSMPLIDCJ2CS & ", "  ' �T���v��ID(CJ2)
'                    'Add End   2010/12/13 SMPK Miyata
                    If .CRYSMPLIDCCS <> 0 Then
                        sql = sql & "CRYSMPLIDCCS=" & .CRYSMPLIDCCS & ", "      ' �T���v��ID(C)
                    End If
                    If .CRYSMPLIDCJCS <> 0 Then
                        sql = sql & "CRYSMPLIDCJCS=" & .CRYSMPLIDCJCS & ", "    ' �T���v��ID(CJ)
                    End If
                    If .CRYSMPLIDCJLTCS <> 0 Then
                        sql = sql & "CRYSMPLIDCJLTCS=" & .CRYSMPLIDCJLTCS & ", " ' �T���v��ID(CJLT)
                    End If
                    If .CRYSMPLIDCJ2CS <> 0 Then
                        sql = sql & "CRYSMPLIDCJ2CS=" & .CRYSMPLIDCJ2CS & ", "  ' �T���v��ID(CJ2)
                    End If
                    'Add End   2011/03/31 SMPK Y.Hitomi
                    sql = sql & "CRYINDRSCS='" & .CRYINDRSCS & "', "        ' ���FLG�iRs)
                    sql = sql & "CRYINDOICS='" & .CRYINDOICS & "', "        ' ���FLG�iOi)
                    sql = sql & "CRYINDB1CS='" & .CRYINDB1CS & "', "        ' ���FLG�iB1)
                    sql = sql & "CRYINDB2CS='" & .CRYINDB2CS & "', "        ' ���FLG�iB2)
                    sql = sql & "CRYINDB3CS='" & .CRYINDB3CS & "', "        ' ���FLG�iB3)
                    sql = sql & "CRYINDL1CS='" & .CRYINDL1CS & "', "        ' ���FLG�iL1)
                    sql = sql & "CRYINDL2CS='" & .CRYINDL2CS & "', "        ' ���FLG�iL2)
                    sql = sql & "CRYINDL3CS='" & .CRYINDL3CS & "', "        ' ���FLG�iL3)
                    sql = sql & "CRYINDL4CS='" & .CRYINDL4CS & "', "        ' ���FLG�iL4)
                    sql = sql & "CRYINDCSCS='" & .CRYINDCSCS & "', "        ' ���FLG�iCs)
                    sql = sql & "CRYINDGDCS='" & .CRYINDGDCS & "', "        ' ���FLG�iGD)
                    sql = sql & "CRYINDTCS='" & .CRYINDTCS & "', "          ' ���FLG�iT)
                    sql = sql & "CRYINDEPCS='" & .CRYINDEPCS & "', "        ' ���FLG�iEPD)
                    'Cng Start 2011/03/31 SMPK Y.Hitomi
'                    'Add Start 2010/12/13 SMPK Miyata
'                    sql = sql & "CRYINDCCS='" & .CRYINDCCS & "', "          ' ���FLG�iC)
'                    sql = sql & "CRYINDCJCS='" & .CRYINDCJCS & "', "        ' ���FLG�iCJ)
'                    sql = sql & "CRYINDCJLTCS='" & .CRYINDCJLTCS & "', "    ' ���FLG�iCJLT)
'                    sql = sql & "CRYINDCJ2CS='" & .CRYINDCJ2CS & "', "      ' ���FLG�iCJ2)
'                    'Add End   2010/12/13 SMPK Miyata
                    If .CRYINDCCS <> "" And left(.CRYINDCCS, 1) <> vbNullChar Then
                        sql = sql & "CRYINDCCS='" & .CRYINDCCS & "', "          ' ���FLG�iC)
                    End If
                    If .CRYINDCJCS <> "" And left(.CRYINDCJCS, 1) <> vbNullChar Then
                        sql = sql & "CRYINDCJCS='" & .CRYINDCJCS & "', "        ' ���FLG�iCJ)
                    End If
                    If .CRYINDCJLTCS <> "" And left(.CRYINDCJLTCS, 1) <> vbNullChar Then
                        sql = sql & "CRYINDCJLTCS='" & .CRYINDCJLTCS & "', "    ' ���FLG�iCJLT)
                    End If
                    If .CRYINDCJ2CS <> "" And left(.CRYINDCJ2CS, 1) <> vbNullChar Then
                        sql = sql & "CRYINDCJ2CS='" & .CRYINDCJ2CS & "', "      ' ���FLG�iCJ2)
                    End If
                    'Cng End 2011/03/31 SMPK Y.Hitomi
                    sql = sql & "CRYRESRS1CS='" & .CRYRESRS1CS & "', "      ' ����FLG1�iRs)
                    sql = sql & "CRYRESRS2CS='" & .CRYRESRS2CS & "', "      ' ����FLG2�iRs)
                    sql = sql & "CRYRESOICS='" & .CRYRESOICS & "', "        ' ����FLG�iOi)
                    sql = sql & "CRYRESB1CS='" & .CRYRESB1CS & "', "        ' ����FLG�iB1)
                    sql = sql & "CRYRESB2CS='" & .CRYRESB2CS & "', "        ' ����FLG�iB2)
                    sql = sql & "CRYRESB3CS='" & .CRYRESB3CS & "', "        ' ����FLG�iB3)
                    sql = sql & "CRYRESL1CS='" & .CRYRESL1CS & "', "        ' ����FLG�iL1)
                    sql = sql & "CRYRESL2CS='" & .CRYRESL2CS & "', "        ' ����FLG�iL2)
                    sql = sql & "CRYRESL3CS='" & .CRYRESL3CS & "', "        ' ����FLG�iL3)
                    sql = sql & "CRYRESL4CS='" & .CRYRESL4CS & "', "        ' ����FLG�iL4)
                    sql = sql & "CRYRESCSCS='" & .CRYRESCSCS & "', "        ' ����FLG�iCs)
                    sql = sql & "CRYRESGDCS='" & .CRYRESGDCS & "', "        ' ����FLG�iGD)
                    sql = sql & "CRYRESTCS='" & .CRYRESTCS & "', "          ' ����FLG�iT)
                    sql = sql & "CRYRESEPCS='" & .CRYRESEPCS & "', "        ' ����FLG�iEPD)
'                    'Add Start 2010/12/13 SMPK Miyata
'                    sql = sql & "CRYRESCCS='" & .CRYRESCCS & "', "          ' ����FLG�iC)
'                    sql = sql & "CRYRESCJCS='" & .CRYRESCJCS & "', "        ' ����FLG�iCJ)
'                    sql = sql & "CRYRESCJLTCS='" & .CRYRESCJLTCS & "', "    ' ����FLG�iCJLT)
'                    sql = sql & "CRYRESCJ2CS='" & .CRYRESCJ2CS & "', "      ' ����FLG�iCJ2)
'                    'Add End   2010/12/13 SMPK Miyata
                    'Cng Start 2011/03/31 SMPK Y.Hitomi
                    If .CRYRESCCS <> "" And left(.CRYRESCCS, 1) <> vbNullChar Then
                        sql = sql & "CRYRESCCS='" & .CRYRESCCS & "', "          ' ����FLG�iC)
                    End If
                    If .CRYRESCJCS <> "" And left(.CRYRESCJCS, 1) <> vbNullChar Then
                        sql = sql & "CRYRESCJCS='" & .CRYRESCJCS & "', "        ' ����FLG�iCJ)
                    End If
                    If .CRYRESCJLTCS <> "" And left(.CRYRESCJLTCS, 1) <> vbNullChar Then
                        sql = sql & "CRYRESCJLTCS='" & .CRYRESCJLTCS & "', "    ' ����FLG�iCJLT)
                    End If
                    If .CRYRESCJ2CS <> "" And left(.CRYRESCJ2CS, 1) <> vbNullChar Then
                        sql = sql & "CRYRESCJ2CS='" & .CRYRESCJ2CS & "', "      ' ����FLG�iCJ2)
                    End If
                    'Add End   Cng Start 2011/03/31 SMPK Y.Hitomi
                    
                    sql = sql & "SMPLNUMCS=" & .SMPLNUMCS & ", "            ' �T���v������
                    sql = sql & "SMPLPATCS='" & .SMPLPATCS & "', "          ' �T���v���p�^�[��
                    sql = sql & "LIVKCS='" & .LIVKCS & "', "                ' �����敪
                    sql = sql & "KSTAFFCS='" & .KSTAFFCS & "', "            ' �X�V�Ј�ID
                    sql = sql & "KDAYCS=sysdate, "                          ' �X�V���t
                    sql = sql & "SNDKCS='0' "                               ' ���M�t���O
                    '>>>>> X������ǉ��Ή� 2009/07/28 SETsw kubota ---------------
                    If .CRYSMPLIDXCS <> 0 Then
                        sql = sql & ",CRYSMPLIDXCS=" & .CRYSMPLIDXCS        ' �T���v��ID(X��)
                    End If
                    If .CRYINDXCS <> "" And left(.CRYINDXCS, 1) <> vbNullChar Then
                        sql = sql & ",CRYINDXCS='" & .CRYINDXCS & "'"       ' ���FLG�iX��)
                    End If
                    If .CRYRESXCS <> "" And left(.CRYRESXCS, 1) <> vbNullChar Then
                        sql = sql & ",CRYRESXCS='" & .CRYRESXCS & "'"       ' ����FLG�iX��)
                    End If
                    '<<<<< X������ǉ��Ή� 2009/07/28 SETsw kubota ---------------
                    '>>>>> ��R�_���ʒu�Ή� 2009/11/06 SETsw kubota ---------------
                    If .QCKBNCS <> "" And left(.QCKBNCS, 1) <> vbNullChar Then
                        sql = sql & ",QCKBNCS='" & .QCKBNCS & "'"           ' (��R�_���ʒu)�Ǘ��敪
                    End If
                    '<<<<< ��R�_���ʒu�Ή� 2009/11/06 SETsw kubota ---------------
                    
                    '05/10/17 ooba START =====================================================>
                    '�e�u���b�NID
                    If left(.RPCRYNUMCS, 1) <> vbNullChar And Trim(.RPCRYNUMCS) <> "" Then
                        sql = sql & ",RPCRYNUMCS='" & .RPCRYNUMCS & "' "
                    End If
                    '�ؒf�t���O
                    If left(.CUTFLGCS, 1) <> vbNullChar And Trim(.CUTFLGCS) <> "" Then
                        sql = sql & ",CUTFLGCS='" & .CUTFLGCS & "' "
                    Else
                        sql = sql & ",CUTFLGCS=NULL "
                    End If
                    '05/10/17 ooba END =======================================================>
'' 09/03/02 FAE)akiyama start
'                    sql = sql & " where XTALCS='" & .XTALCS & "'"
                    sql = sql & " where CRYNUMCS LIKE '" & left(.XTALCS, 9) & "%'"
'' 09/03/02 FAE)akiyama end
                    sql = sql & " and INPOSCS=" & .INPOSCS
                    sql = sql & " and SMPKBNCS='" & .SMPKBNCS & "'"

                    '' WriteDBLog sql
                    If OraDB.ExecuteSQL(sql) <= 0 Then
                        DBDRV_CrySmp_UpdIns = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                    lFlg = True
                    Exit For
                End If
            Next j

            If lFlg <> True Then
'                sql = "insert into TBCME043 ("
'                sql = sql & "CRYNUM, "          ' �����ԍ�
'                sql = sql & "INGOTPOS, "        ' �������ʒu
'                sql = sql & "SMPKBN, "          ' �T���v���敪
'                sql = sql & "SMPLNO, "          ' �T���v��No
'                sql = sql & "HINBAN, "          ' �i��
'                sql = sql & "REVNUM, "          ' ���i�ԍ������ԍ�
'                sql = sql & "FACTORY, "         ' �H��
'                sql = sql & "OPECOND, "         ' ���Ə���
'                sql = sql & "KTKBN, "           ' �m��敪
'                sql = sql & "CRYINDRS, "        ' ���������w���iRs)
'                sql = sql & "CRYINDOI, "        ' ���������w���iOi)
'                sql = sql & "CRYINDB1, "        ' ���������w���iB1)
'                sql = sql & "CRYINDB2, "        ' ���������w���iB2)
'                sql = sql & "CRYINDB3, "        ' ���������w���iB3)
'                sql = sql & "CRYINDL1, "        ' ���������w���iL1)
'                sql = sql & "CRYINDL2, "        ' ���������w���iL2)
'                sql = sql & "CRYINDL3, "        ' ���������w���iL3)
'                sql = sql & "CRYINDL4, "        ' ���������w���iL4)
'                sql = sql & "CRYINDCS, "        ' ���������w���iCs)
'                sql = sql & "CRYINDGD, "        ' ���������w���iGD)
'                sql = sql & "CRYINDT, "         ' ���������w���iT)
'                sql = sql & "CRYINDEP, "        ' ���������w���iEPD)
'                sql = sql & "CRYRESRS, "        ' �����������сiRs)
'                sql = sql & "CRYRESOI, "        ' �����������сiOi)
'                sql = sql & "CRYRESB1, "        ' �����������сiB1)
'                sql = sql & "CRYRESB2, "        ' �����������сiB2)
'                sql = sql & "CRYRESB3, "        ' �����������сiB3)
'                sql = sql & "CRYRESL1, "        ' �����������сiL1)
'                sql = sql & "CRYRESL2, "        ' �����������сiL2)
'                sql = sql & "CRYRESL3, "        ' �����������сiL3)
'                sql = sql & "CRYRESL4, "        ' �����������сiL4)
'                sql = sql & "CRYRESCS, "        ' �����������сiCs)
'                sql = sql & "CRYRESGD, "        ' �����������сiGD)
'                sql = sql & "CRYREST, "         ' �����������сiT)
'                sql = sql & "CRYRESEP, "        ' �����������сiEPD)
'                sql = sql & "SMPLNUM, "         ' �T���v������
'                sql = sql & "SMPLPAT, "         ' �T���v���p�^�[��
'                sql = sql & "REGDATE, "         ' �o�^���t
'                sql = sql & "UPDDATE, "         ' �X�V���t
'                sql = sql & "SENDFLAG, "        ' ���M�t���O
'                sql = sql & "SENDDATE)"         ' ���M���t
'                sql = sql & " values ('"
'                sql = sql & .CRYNUM & "', "
'                sql = sql & .INGOTPOS & ", '"   ' �������ʒu
'                sql = sql & .SMPKBN & "', "     ' �T���v���敪
'                sql = sql & Abs(.SMPLNO) & ", '"     ' �T���v��No
'                sql = sql & .HINBAN & "', "     ' �i��
'                sql = sql & .REVNUM & ", '"     ' ���i�ԍ������ԍ�
'                sql = sql & .factory & "', '"   ' �H��
'                sql = sql & .opecond & "', '"   ' ���Ə���
'                sql = sql & .KTKBN & "', '"     ' �m��敪
'                sql = sql & .CRYINDRS & "', '"  ' ���������w���iRs)
'                sql = sql & .CRYINDOI & "', '"  ' ���������w���iOi)
'                sql = sql & .CRYINDB1 & "', '"  ' ���������w���iB1)
'                sql = sql & .CRYINDB2 & "', '"  ' ���������w���iB2)
'                sql = sql & .CRYINDB3 & "', '"  ' ���������w���iB3)
'                sql = sql & .CRYINDL1 & "', '"  ' ���������w���iL1)
'                sql = sql & .CRYINDL2 & "', '"  ' ���������w���iL2)
'                sql = sql & .CRYINDL3 & "', '"  ' ���������w���iL3)
'                sql = sql & .CRYINDL4 & "', '"  ' ���������w���iL4)
'                sql = sql & .CRYINDCS & "', '"  ' ���������w���iCs)
'                sql = sql & .CRYINDGD & "', '"  ' ���������w���iGD)
'                sql = sql & .CRYINDT & "', '"   ' ���������w���iT)
'                sql = sql & .CRYINDEP & "', '"  ' ���������w���iEPD)
'                sql = sql & .CRYRESRS & "', '"  ' �����������сiRs)
'                sql = sql & .CRYRESOI & "', '"  ' �����������сiOi)
'                sql = sql & .CRYRESB1 & "', '"  ' �����������сiB1)
'                sql = sql & .CRYRESB2 & "', '"  ' �����������сiB2)
'                sql = sql & .CRYRESB3 & "', '"  ' �����������сiB3)
'                sql = sql & .CRYRESL1 & "', '"  ' �����������сiL1)
'                sql = sql & .CRYRESL2 & "', '"  ' �����������сiL2)
'                sql = sql & .CRYRESL3 & "', '"  ' �����������сiL3)
'                sql = sql & .CRYRESL4 & "', '"  ' �����������сiL4)
'                sql = sql & .CRYRESCS & "', '"  ' �����������сiCs)
'                sql = sql & .CRYRESGD & "', '"  ' �����������сiGD)
'                sql = sql & .CRYREST & "', '"   ' �����������сiT)
'                sql = sql & .CRYRESEP & "', "   ' �����������сiEPD)
'                sql = sql & .SMPLNUM & ", "     ' �T���v������
'                sql = sql & "' ', "             ' �T���v���p�^�[��
'                sql = sql & "sysdate, "
'                sql = sql & "sysdate, "
'                sql = sql & "'0', "
'                sql = sql & "sysdate)"
                sql = "insert into XSDCS ("
                sql = sql & "CRYNUMCS,"         '�u���b�NID
                sql = sql & "SMPKBNCS,"         '�T���v���敪
                sql = sql & "TBKBNCS,"          'T/B�敪
                sql = sql & "REPSMPLIDCS,"      '��\�T���v��ID
                sql = sql & "XTALCS,"           '�����ԍ�
                sql = sql & "INPOSCS,"          '�������ʒu
                sql = sql & "HINBCS,"           '�i��
                sql = sql & "REVNUMCS,"         '���i�ԍ������ԍ�
                sql = sql & "FACTORYCS,"        '�H��
                sql = sql & "OPECS,"            '���Ɣԍ�
                sql = sql & "KTKBNCS,"          '�m��敪
                sql = sql & "BLKKTFLAGCS,"      '�u���b�N�m��t���O
                sql = sql & "CRYSMPLIDRSCS,"    '�T���v��ID(Rs)
                sql = sql & "CRYSMPLIDRS1CS,"   '����T���v��ID1�iRs�j
                sql = sql & "CRYSMPLIDRS2CS,"   '����T���v��ID2�iRs�j
                sql = sql & "CRYINDRSCS,"       '���FLG(Rs)
                sql = sql & "CRYRESRS1CS,"      '����FLG1(Rs)
                sql = sql & "CRYRESRS2CS,"      '����FLG2(Rs)
                sql = sql & "CRYSMPLIDOICS,"    '�T���v��ID�iOi�j
                sql = sql & "CRYINDOICS,"       '���FLG�iOi�j
                sql = sql & "CRYRESOICS,"       '����FLG�iOi�j
                sql = sql & "CRYSMPLIDB1CS,"    '�T���v��ID�iB1�j
                sql = sql & "CRYINDB1CS,"       '���FLG�iB1�j
                sql = sql & "CRYRESB1CS,"       '����FLG�iB1�j
                sql = sql & "CRYSMPLIDB2CS,"    '�T���v��ID�iB2�j
                sql = sql & "CRYINDB2CS,"       '���FLG�iB2�j
                sql = sql & "CRYRESB2CS,"       '����FLG�iB2�j
                sql = sql & "CRYSMPLIDB3CS,"    '�T���v��ID�iB3�j
                sql = sql & "CRYINDB3CS,"       '���FLG�iB3�j
                sql = sql & "CRYRESB3CS,"       '����FLG�iB3�j
                sql = sql & "CRYSMPLIDL1CS,"    '�T���v��ID�iL1�j
                sql = sql & "CRYINDL1CS,"       '���FLG�iL1�j
                sql = sql & "CRYRESL1CS,"       '����FLG�iL1�j
                sql = sql & "CRYSMPLIDL2CS,"    '�T���v��ID�iL2�j
                sql = sql & "CRYINDL2CS,"       '���FLG�iL2�j
                sql = sql & "CRYRESL2CS,"       '����FLG�iL2�j
                sql = sql & "CRYSMPLIDL3CS,"    '�T���v��ID�iL3�j
                sql = sql & "CRYINDL3CS,"       '���FLG�iL3�j
                sql = sql & "CRYRESL3CS,"       '����FLG�iL3�j
                sql = sql & "CRYSMPLIDL4CS,"    '�T���v��ID�iL4�j
                sql = sql & "CRYINDL4CS,"       '���FLG�iL4�j
                sql = sql & "CRYRESL4CS,"       '����FLG�iL4�j
                sql = sql & "CRYSMPLIDCSCS,"    '�T���v��ID�iCS�j
                sql = sql & "CRYINDCSCS,"       '���FLG�iCS�j
                sql = sql & "CRYRESCSCS,"       '����FLG�iCS�j
                sql = sql & "CRYSMPLIDGDCS,"    '�T���v��ID�iGD�j
                sql = sql & "CRYINDGDCS,"       '���FLG�iGD�j
                sql = sql & "CRYRESGDCS,"       '����FLG�iGD�j
                sql = sql & "CRYSMPLIDTCS,"     '�T���v��ID�iT�j
                sql = sql & "CRYINDTCS,"        '���FLG�iT�j
                sql = sql & "CRYRESTCS,"        '����FLG�iT�j
                sql = sql & "CRYSMPLIDEPCS,"    '�T���v��ID�iEPD�j
                sql = sql & "CRYINDEPCS,"       '���FLG�iEPD�j
                sql = sql & "CRYRESEPCS,"       '����FLG�iEPD�j
                sql = sql & "CRYSMPLIDXCS,"     '�T���v��ID�iX���j  'X������ 2009/07/27�ǉ� SETsw kubota
                sql = sql & "CRYINDXCS,"        '���FLG�iX���j
                sql = sql & "CRYRESXCS,"        '����FLG�iX���j
                'Add Start 2010/12/13 SMPK Miyata
                sql = sql & "CRYSMPLIDCCS,"     '�T���v��ID�iC�j
                sql = sql & "CRYINDCCS,"        '���FLG�iC�j
                sql = sql & "CRYRESCCS,"        '����FLG�iC�j
                sql = sql & "CRYSMPLIDCJCS,"    '�T���v��ID�iCJ�j
                sql = sql & "CRYINDCJCS,"       '���FLG�iCJ�j
                sql = sql & "CRYRESCJCS,"       '����FLG�iCJ�j
                sql = sql & "CRYSMPLIDCJLTCS,"  '�T���v��ID�iCJLT�j
                sql = sql & "CRYINDCJLTCS,"     '���FLG�iCJLT�j
                sql = sql & "CRYRESCJLTCS,"     '����FLG�iCJLT�j
                sql = sql & "CRYSMPLIDCJ2CS,"   '�T���v��ID�iCJ2�j
                sql = sql & "CRYINDCJ2CS,"      '���FLG�iCJ2�j
                sql = sql & "CRYRESCJ2CS,"      '����FLG�iCJ2�j
                'Add End   2010/12/13 SMPK Miyata
                sql = sql & "SMPLNUMCS,"        '�T���v������
                sql = sql & "SMPLPATCS,"        '�T���v���p�^�[��
                sql = sql & "LIVKCS,"           '�����敪
                sql = sql & "TSTAFFCS,"         '�o�^�Ј�ID
                sql = sql & "TDAYCS,"           '�o�^���t
                sql = sql & "KSTAFFCS,"         '�X�V�Ј�ID
                sql = sql & "KDAYCS,"           '�X�V���t
                sql = sql & "SNDKCS,"           '���M�t���O
'                sql = sql & "SNDDAYCS)"         '���M���t
                '05/10/17 ooba START =====================================================>
                sql = sql & "SNDDAYCS"          '���M���t
                '�e�u���b�NID
                If left(.RPCRYNUMCS, 1) <> vbNullChar And Trim(.RPCRYNUMCS) <> "" Then
                    sql = sql & ",RPCRYNUMCS"
                End If
                '�ؒf�t���O
                sql = sql & ",CUTFLGCS"
                sql = sql & ",QCKBNCS"          '�Ǘ��敪       2009/11/06�ǉ� SETsw kubota
                sql = sql & ")"
                '05/10/17 ooba END =======================================================>
                sql = sql & " values ('"
                sql = sql & .CRYNUMCS & "', '"          '�u���b�NID
                sql = sql & .SMPKBNCS & "', '"          '�T���v���敪
                sql = sql & .TBKBNCS & "', "            'T/B�敪
                sql = sql & .REPSMPLIDCS & ", '"        '��\�T���v��ID
                sql = sql & .XTALCS & "', "             '�����ԍ�
                sql = sql & .INPOSCS & ", '"            '�������ʒu
                sql = sql & .HINBCS & "', "             '�i��
                sql = sql & .REVNUMCS & ", '"           '���i�ԍ������ԍ�
                sql = sql & .FACTORYCS & "', '"         '�H��
                sql = sql & .OPECS & "', '"             '���Ə���
                sql = sql & .KTKBNCS & "', '"           '�m��敪
                sql = sql & .BLKKTFLAGCS & "', "        '�u���b�N�m��t���O
                sql = sql & .CRYSMPLIDRSCS & ", "       '�T���v��ID�iRs�j
                sql = sql & .CRYSMPLIDRS1CS & ", "      '����T���v��ID1�iRs�j
                sql = sql & .CRYSMPLIDRS2CS & ", '"     '����T���v��ID2�iRs�j
                sql = sql & .CRYINDRSCS & "', '"        '���FLG�iRs�j
                sql = sql & .CRYRESRS1CS & "', '"       '����FLG1�iRs�j
                sql = sql & .CRYRESRS2CS & "', "        '����FLG2�iRs�j
                sql = sql & .CRYSMPLIDOICS & ", '"      '�T���v��ID�iOi�j
                sql = sql & .CRYINDOICS & "', '"        '���FLG�iOi�j
                sql = sql & .CRYRESOICS & "', "         '����FLG�iOi�j
                sql = sql & .CRYSMPLIDB1CS & ", '"      '�T���v��ID�iB1�j
                sql = sql & .CRYINDB1CS & "', '"        '���FLG�iB1�j
                sql = sql & .CRYRESB1CS & "', "         '����FLG�iB1�j
                sql = sql & .CRYSMPLIDB2CS & ", '"      '�T���v��ID�iB2�j
                sql = sql & .CRYINDB2CS & "', '"        '���FLG�iB2�j
                sql = sql & .CRYRESB2CS & "', "         '����FLG�iB2�j
                sql = sql & .CRYSMPLIDB3CS & ", '"      '�T���v��ID�iB3�j
                sql = sql & .CRYINDB3CS & "', '"        '���FLG�iB3�j
                sql = sql & .CRYRESB3CS & "', "         '����FLG�iB3�j
                sql = sql & .CRYSMPLIDL1CS & ", '"      '�T���v��ID�iL1�j
                sql = sql & .CRYINDL1CS & "', '"        '���FLG�iL1�j
                sql = sql & .CRYRESL1CS & "', "         '����FLG�iL1�j
                sql = sql & .CRYSMPLIDL2CS & ", '"      '�T���v��ID�iL2�j
                sql = sql & .CRYINDL2CS & "', '"        '���FLG�iL2�j
                sql = sql & .CRYRESL2CS & "', "         '����FLG�iL2�j
                sql = sql & .CRYSMPLIDL3CS & ", '"      '�T���v��ID�iL3�j
                sql = sql & .CRYINDL3CS & "', '"        '���FLG�iL3�j
                sql = sql & .CRYRESL3CS & "', "         '����FLG�iL3�j
                sql = sql & .CRYSMPLIDL4CS & ", '"      '�T���v��ID�iL4�j
                sql = sql & .CRYINDL4CS & "', '"        '���FLG�iL4�j
                sql = sql & .CRYRESL4CS & "', "         '����FLG�iL4�j
                sql = sql & .CRYSMPLIDCSCS & ", '"      '�T���v��ID�iCS�j
                sql = sql & .CRYINDCSCS & "', '"        '���FLG�iCS�j
                sql = sql & .CRYRESCSCS & "', "         '����FLG�iCS�j
                sql = sql & .CRYSMPLIDGDCS & ", '"      '�T���v��ID�iGD�j
                sql = sql & .CRYINDGDCS & "', '"        '���FLG�iGD�j
                sql = sql & .CRYRESGDCS & "', "         '����FLG�iGD�j
                sql = sql & .CRYSMPLIDTCS & ", '"       '�T���v��ID�iT�j
                sql = sql & .CRYINDTCS & "', '"         '���FLG�iT�j
                sql = sql & .CRYRESTCS & "', "          '����FLG�iT�j
                sql = sql & .CRYSMPLIDEPCS & ", '"      '�T���v��ID�iEPD�j
                sql = sql & .CRYINDEPCS & "', '"        '���FLG�iEPD�j
                sql = sql & .CRYRESEPCS & "', "         '����FLG�iEPD�j
                
                '>>>>> X������ǉ��Ή� 2009/07/28 SETsw kubota ---------------
                sql = sql & .CRYSMPLIDXCS               '�T���v��ID�iX���j
                '���FLG�iX���j
                If .CRYINDXCS <> "" And left(.CRYINDXCS, 1) <> vbNullChar Then
                    sql = sql & ",'" & .CRYINDXCS & "'"
                Else
                    sql = sql & ",'0'"
                End If
                '����FLG�iX���j
                If .CRYRESXCS <> "" And left(.CRYRESXCS, 1) <> vbNullChar Then
                    sql = sql & ",'" & .CRYRESXCS & "'"
                Else
                    sql = sql & ",'0'"
                End If
                sql = sql & ", "
                '<<<<< X������ǉ��Ή� 2009/07/28 SETsw kubota ---------------

                'Add Start 2010/12/13 SMPK Miyata
                sql = sql & .CRYSMPLIDCCS & ", '"       '�T���v��ID�iC�j
                sql = sql & .CRYINDCCS & "', '"         '���FLG�iC�j
                sql = sql & .CRYRESCCS & "', "          '����FLG�iC�j
                sql = sql & .CRYSMPLIDCJCS & ", '"      '�T���v��ID�iCJ�j
                sql = sql & .CRYINDCJCS & "', '"        '���FLG�iCJ�j
                sql = sql & .CRYRESCJCS & "', "         '����FLG�iCJ�j
                sql = sql & .CRYSMPLIDCJLTCS & ", '"    '�T���v��ID�iCJLT�j
                sql = sql & .CRYINDCJLTCS & "', '"      '���FLG�iCJLT�j
                sql = sql & .CRYRESCJLTCS & "', "       '����FLG�iCJLT�j
                sql = sql & .CRYSMPLIDCJ2CS & ", '"     '�T���v��ID�iCJ2�j
                sql = sql & .CRYINDCJ2CS & "', '"       '���FLG�iCJ2�j
                sql = sql & .CRYRESCJ2CS & "', "        '����FLG�iCJ2�j
                'Add End   2010/12/13 SMPK Miyata

                sql = sql & .SMPLNUMCS & ", "           '�T���v������
                sql = sql & "' ', '"                    '�T���v���p�^�[��
                sql = sql & .LIVKCS & "', '"            '�����敪
                sql = sql & .TSTAFFCS & "', "           '�o�^�Ј�ID
                sql = sql & "sysdate, '"                '�o�^���t
                sql = sql & .KSTAFFCS & "', "           '�X�V�Ј�ID
                sql = sql & "sysdate, "                 '�X�V���t
                sql = sql & "'0', "                     '���M�t���O
'                sql = sql & "sysdate)"                  '���M���t
                '05/10/17 ooba START =====================================================>
                sql = sql & "sysdate"                   '���M���t
                '�e�u���b�NID
                If left(.RPCRYNUMCS, 1) <> vbNullChar And Trim(.RPCRYNUMCS) <> "" Then
                    sql = sql & ", '" & .RPCRYNUMCS & "'"
                End If
                '�ؒf�t���O
                If left(.CUTFLGCS, 1) <> vbNullChar And Trim(.CUTFLGCS) <> "" Then
                    sql = sql & ", '" & .CUTFLGCS & "'"
                Else
                    sql = sql & ", NULL"
                End If
                '05/10/17 ooba END =======================================================>
                
                '>>>>> ��R�_���ʒu�Ή� 2009/11/06 SETsw kubota ---------------
                If .QCKBNCS <> "" And left(.QCKBNCS, 1) <> vbNullChar Then
                    sql = sql & ",'" & .QCKBNCS & "'"
                Else
                    sql = sql & ",NULL"
                End If
                '<<<<< ��R�_���ʒu�Ή� 2009/11/06 SETsw kubota ---------------
                
                sql = sql & ")"
                
                
                '' WriteDBLog sql
                If OraDB.ExecuteSQL(sql) <= 0 Then
                    DBDRV_CrySmp_UpdIns = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            End If
        End With
    Next i

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_CrySmp_UpdIns = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'�T�v      :�����T���v���Ǘ��̑}��
'���Ұ��@�@:�ϐ���         ,IO ,�^                  ,����
'      �@�@:CrySmpNew�@�@�@,I  ,typ_XSDCS   �@      ,�V�T���v���Ǘ��i�u���b�N�j
'      �@�@:�߂�l         ,O  ,FUNCTION_RETURN�@   ,�������݂̐���
'����      :
'����      :2001/07/12  �쐬 ���{
Public Function DBDRV_CrySmp_Ins(CrySmpNew() As typ_XSDCS) As FUNCTION_RETURN

    Dim sql As String
    Dim lFlg As Boolean
    Dim i As Long
    Dim j As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_CrySmp_Ins"

    DBDRV_CrySmp_Ins = FUNCTION_RETURN_SUCCESS

    For i = 1 To UBound(CrySmpNew)
        With CrySmpNew(i)
'            sql = "insert into TBCME043 ("
'            sql = sql & "CRYNUM, "          ' �����ԍ�
'            sql = sql & "INGOTPOS, "        ' �������ʒu
'            sql = sql & "SMPKBN, "          ' �T���v���敪
'            sql = sql & "SMPLNO, "          ' �T���v��No
'            sql = sql & "HINBAN, "          ' �i��
'            sql = sql & "REVNUM, "          ' ���i�ԍ������ԍ�
'            sql = sql & "FACTORY, "         ' �H��
'            sql = sql & "OPECOND, "         ' ���Ə���
'            sql = sql & "KTKBN, "           ' �m��敪
'            sql = sql & "CRYINDRS, "        ' ���������w���iRs)
'            sql = sql & "CRYINDOI, "        ' ���������w���iOi)
'            sql = sql & "CRYINDB1, "        ' ���������w���iB1)
'            sql = sql & "CRYINDB2, "        ' ���������w���iB2)
'            sql = sql & "CRYINDB3, "        ' ���������w���iB3)
'            sql = sql & "CRYINDL1, "        ' ���������w���iL1)
'            sql = sql & "CRYINDL2, "        ' ���������w���iL2)
'            sql = sql & "CRYINDL3, "        ' ���������w���iL3)
'            sql = sql & "CRYINDL4, "        ' ���������w���iL4)
'            sql = sql & "CRYINDCS, "        ' ���������w���iCs)
'            sql = sql & "CRYINDGD, "        ' ���������w���iGD)
'            sql = sql & "CRYINDT, "         ' ���������w���iT)
'            sql = sql & "CRYINDEP, "        ' ���������w���iEPD)
'            sql = sql & "CRYRESRS, "        ' �����������сiRs)
'            sql = sql & "CRYRESOI, "        ' �����������сiOi)
'            sql = sql & "CRYRESB1, "        ' �����������сiB1)
'            sql = sql & "CRYRESB2, "        ' �����������сiB2)
'            sql = sql & "CRYRESB3, "        ' �����������сiB3)
'            sql = sql & "CRYRESL1, "        ' �����������сiL1)
'            sql = sql & "CRYRESL2, "        ' �����������сiL2)
'            sql = sql & "CRYRESL3, "        ' �����������сiL3)
'            sql = sql & "CRYRESL4, "        ' �����������сiL4)
'            sql = sql & "CRYRESCS, "        ' �����������сiCs)
'            sql = sql & "CRYRESGD, "        ' �����������сiGD)
'            sql = sql & "CRYREST, "         ' �����������сiT)
'            sql = sql & "CRYRESEP, "        ' �����������сiEPD)
'            sql = sql & "SMPLNUM, "         ' �T���v������
'            sql = sql & "SMPLPAT, "         ' �T���v���p�^�[��
'            sql = sql & "REGDATE, "         ' �o�^���t
'            sql = sql & "UPDDATE, "         ' �X�V���t
'            sql = sql & "SENDFLAG, "        ' ���M�t���O
'            sql = sql & "SENDDATE)"         ' ���M���t
'            sql = sql & " values ('"
'            sql = sql & .CRYNUM & "', "
'            sql = sql & .INGOTPOS & ", '"   ' �������ʒu
'            sql = sql & .SMPKBN & "', "     ' �T���v���敪
'            sql = sql & .SMPLNO & ", '"     ' �T���v��No
'            sql = sql & .HINBAN & "', "     ' �i��
'            sql = sql & .REVNUM & ", '"     ' ���i�ԍ������ԍ�
'            sql = sql & .factory & "', '"   ' �H��
'            sql = sql & .opecond & "', '"   ' ���Ə���
'            sql = sql & .KTKBN & "', '"     ' �m��敪
'            sql = sql & .CRYINDRS & "', '"  ' ���������w���iRs)
'            sql = sql & .CRYINDOI & "', '"  ' ���������w���iOi)
'            sql = sql & .CRYINDB1 & "', '"  ' ���������w���iB1)
'            sql = sql & .CRYINDB2 & "', '"  ' ���������w���iB2)
'            sql = sql & .CRYINDB3 & "', '"  ' ���������w���iB3)
'            sql = sql & .CRYINDL1 & "', '"  ' ���������w���iL1)
'            sql = sql & .CRYINDL2 & "', '"  ' ���������w���iL2)
'            sql = sql & .CRYINDL3 & "', '"  ' ���������w���iL3)
'            sql = sql & .CRYINDL4 & "', '"  ' ���������w���iL4)
'            sql = sql & .CRYINDCS & "', '"  ' ���������w���iCs)
'            sql = sql & .CRYINDGD & "', '"  ' ���������w���iGD)
'            sql = sql & .CRYINDT & "', '"   ' ���������w���iT)
'            sql = sql & .CRYINDEP & "', '"  ' ���������w���iEPD)
'            sql = sql & .CRYRESRS & "', '"  ' �����������сiRs)
'            sql = sql & .CRYRESOI & "', '"  ' �����������сiOi)
'            sql = sql & .CRYRESB1 & "', '"  ' �����������сiB1)
'            sql = sql & .CRYRESB2 & "', '"  ' �����������сiB2)
'            sql = sql & .CRYRESB3 & "', '"  ' �����������сiB3)
'            sql = sql & .CRYRESL1 & "', '"  ' �����������сiL1)
'            sql = sql & .CRYRESL2 & "', '"  ' �����������сiL2)
'            sql = sql & .CRYRESL3 & "', '"  ' �����������сiL3)
'            sql = sql & .CRYRESL4 & "', '"  ' �����������сiL4)
'            sql = sql & .CRYRESCS & "', '"  ' �����������сiCs)
'            sql = sql & .CRYRESGD & "', '"  ' �����������сiGD)
'            sql = sql & .CRYREST & "', '"   ' �����������сiT)
'            sql = sql & .CRYRESEP & "', "   ' �����������сiEPD)
'            sql = sql & .SMPLNUM & ", "     ' �T���v������
'            sql = sql & "' ', "             ' �T���v���p�^�[��
'            sql = sql & "sysdate, "
'            sql = sql & "sysdate, "
'            sql = sql & "'0', "
'            sql = sql & "sysdate)"
            sql = "insert into XSDCS ("
            sql = sql & "CRYNUMCS,"         '�u���b�NID
            sql = sql & "SMPKBNCS,"         '�T���v���敪
            sql = sql & "TBKBNCS,"          'T/B�敪
            sql = sql & "REPSMPLIDCS,"      '��\�T���v��ID
            sql = sql & "XTALCS,"           '�����ԍ�
            sql = sql & "INPOSCS,"          '�������ʒu
            sql = sql & "HINBCS,"           '�i��
            sql = sql & "REVNUMCS,"         '���i�ԍ������ԍ�
            sql = sql & "FACTORYCS,"        '�H��
            sql = sql & "OPECS,"            '���Ɣԍ�
            sql = sql & "KTKBNCS,"          '�m��敪
            sql = sql & "BLKKTFLAGCS,"      '�u���b�N�m��t���O
            sql = sql & "CRYSMPLIDRSCS,"    '�T���v��ID(Rs)
            sql = sql & "CRYSMPLIDRS1CS,"   '����T���v��ID1�iRs�j
            sql = sql & "CRYSMPLIDRS2CS,"   '����T���v��ID2�iRs�j
            sql = sql & "CRYINDRSCS,"       '���FLG(Rs)
            sql = sql & "CRYRESRS1CS,"      '����FLG1(Rs)
            sql = sql & "CRYRESRS2CS,"      '����FLG2(Rs)
            sql = sql & "CRYSMPLIDOICS,"    '�T���v��ID�iOi�j
            sql = sql & "CRYINDOICS,"       '���FLG�iOi�j
            sql = sql & "CRYRESOICS,"       '����FLG�iOi�j
            sql = sql & "CRYSMPLIDB1CS,"    '�T���v��ID�iB1�j
            sql = sql & "CRYINDB1CS,"       '���FLG�iB1�j
            sql = sql & "CRYRESB1CS,"       '����FLG�iB1�j
            sql = sql & "CRYSMPLIDB2CS,"    '�T���v��ID�iB2�j
            sql = sql & "CRYINDB2CS,"       '���FLG�iB2�j
            sql = sql & "CRYRESB2CS,"       '����FLG�iB2�j
            sql = sql & "CRYSMPLIDB3CS,"    '�T���v��ID�iB3�j
            sql = sql & "CRYINDB3CS,"       '���FLG�iB3�j
            sql = sql & "CRYRESB3CS,"       '����FLG�iB3�j
            sql = sql & "CRYSMPLIDL1CS,"    '�T���v��ID�iL1�j
            sql = sql & "CRYINDL1CS,"       '���FLG�iL1�j
            sql = sql & "CRYRESL1CS,"       '����FLG�iL1�j
            sql = sql & "CRYSMPLIDL2CS,"    '�T���v��ID�iL2�j
            sql = sql & "CRYINDL2CS,"       '���FLG�iL2�j
            sql = sql & "CRYRESL2CS,"       '����FLG�iL2�j
            sql = sql & "CRYSMPLIDL3CS,"    '�T���v��ID�iL3�j
            sql = sql & "CRYINDL3CS,"       '���FLG�iL3�j
            sql = sql & "CRYRESL3CS,"       '����FLG�iL3�j
            sql = sql & "CRYSMPLIDL4CS,"    '�T���v��ID�iL4�j
            sql = sql & "CRYINDL4CS,"       '���FLG�iL4�j
            sql = sql & "CRYRESL4CS,"       '����FLG�iL4�j
            sql = sql & "CRYSMPLIDCSCS,"    '�T���v��ID�iCS�j
            sql = sql & "CRYINDCSCS,"       '���FLG�iCS�j
            sql = sql & "CRYRESCSCS,"       '����FLG�iCS�j
            sql = sql & "CRYSMPLIDGDCS,"    '�T���v��ID�iGD�j
            sql = sql & "CRYINDGDCS,"       '���FLG�iGD�j
            sql = sql & "CRYRESGDCS,"       '����FLG�iGD�j
            sql = sql & "CRYSMPLIDTCS,"     '�T���v��ID�iT�j
            sql = sql & "CRYINDTCS,"        '���FLG�iT�j
            sql = sql & "CRYRESTCS,"        '����FLG�iT�j
            sql = sql & "CRYSMPLIDEPCS,"    '�T���v��ID�iEPD�j
            sql = sql & "CRYINDEPCS,"       '���FLG�iEPD�j
            sql = sql & "CRYRESEPCS,"       '����FLG�iEPD�j
            sql = sql & "CRYSMPLIDXCS,"     '�T���v��ID�iX���j  'X������ 2009/07/27�ǉ� SETsw kubota
            sql = sql & "CRYINDXCS,"        '���FLG�iX���j
            sql = sql & "CRYRESXCS,"        '����FLG�iX���j
            sql = sql & "SMPLNUMCS,"        '�T���v������
            sql = sql & "SMPLPATCS,"        '�T���v���p�^�[��
            sql = sql & "TSTAFFCS,"         '�o�^�Ј�ID
            sql = sql & "TDAYCS,"           '�o�^���t
            sql = sql & "KSTAFFCS,"         '�X�V�Ј�ID
            sql = sql & "KDAYCS,"           '�X�V���t
            sql = sql & "SNDKCS,"           '���M�t���O
            sql = sql & "SNDDAYCS)"         '���M���t
            sql = sql & " values ('"
            sql = sql & .CRYNUMCS & "', '"          '�u���b�NID
            sql = sql & .SMPKBNCS & "', '"          '�T���v���敪
            sql = sql & .TBKBNCS & "', "            'T/B�敪
            sql = sql & .REPSMPLIDCS & ", '"        '��\�T���v��ID
            sql = sql & .XTALCS & "', "             '�����ԍ�
            sql = sql & .INPOSCS & ", '"            '�������ʒu
            sql = sql & .HINBCS & "', "             '�i��
            sql = sql & .REVNUMCS & ", '"           '���i�ԍ������ԍ�
            sql = sql & .FACTORYCS & "', '"         '�H��
            sql = sql & .OPECS & "', '"             '���Ə���
            sql = sql & .KTKBNCS & "', '"           '�m��敪
            sql = sql & .BLKKTFLAGCS & "', "        '�u���b�N�m��t���O
            sql = sql & .CRYSMPLIDRSCS & ", "       '�T���v��ID�iRs�j
            sql = sql & .CRYSMPLIDRS1CS & ", "      '����T���v��ID1�iRs�j
            sql = sql & .CRYSMPLIDRS2CS & ", '"     '����T���v��ID2�iRs�j
            sql = sql & .CRYINDRSCS & "', '"        '���FLG�iRs�j
            sql = sql & .CRYRESRS1CS & "', '"       '����FLG1�iRs�j
            sql = sql & .CRYRESRS2CS & "', "        '����FLG2�iRs�j
            sql = sql & .CRYSMPLIDOICS & ", '"      '�T���v��ID�iOi�j
            sql = sql & .CRYINDOICS & "', '"        '���FLG�iOi�j
            sql = sql & .CRYRESOICS & "', "         '����FLG�iOi�j
            sql = sql & .CRYSMPLIDB1CS & ", '"      '�T���v��ID�iB1�j
            sql = sql & .CRYINDB1CS & "', '"        '���FLG�iB1�j
            sql = sql & .CRYRESB1CS & "', "         '����FLG�iB1�j
            sql = sql & .CRYSMPLIDB2CS & ", '"      '�T���v��ID�iB2�j
            sql = sql & .CRYINDB2CS & "', '"        '���FLG�iB2�j
            sql = sql & .CRYRESB2CS & "', "         '����FLG�iB2�j
            sql = sql & .CRYSMPLIDB3CS & ", '"      '�T���v��ID�iB3�j
            sql = sql & .CRYINDB3CS & "', '"        '���FLG�iB3�j
            sql = sql & .CRYRESB3CS & "', "         '����FLG�iB3�j
            sql = sql & .CRYSMPLIDL1CS & ", '"      '�T���v��ID�iL1�j
            sql = sql & .CRYINDL1CS & "', '"        '���FLG�iL1�j
            sql = sql & .CRYRESL1CS & "', "         '����FLG�iL1�j
            sql = sql & .CRYSMPLIDL2CS & ", '"      '�T���v��ID�iL2�j
            sql = sql & .CRYINDL2CS & "', '"        '���FLG�iL2�j
            sql = sql & .CRYRESL2CS & "', "         '����FLG�iL2�j
            sql = sql & .CRYSMPLIDL3CS & ", '"      '�T���v��ID�iL3�j
            sql = sql & .CRYINDL3CS & "', '"        '���FLG�iL3�j
            sql = sql & .CRYRESL3CS & "', "         '����FLG�iL3�j
            sql = sql & .CRYSMPLIDL4CS & ", '"      '�T���v��ID�iL4�j
            sql = sql & .CRYINDL4CS & "', '"        '���FLG�iL4�j
            sql = sql & .CRYRESL4CS & "', "         '����FLG�iL4�j
            sql = sql & .CRYSMPLIDCSCS & ", '"      '�T���v��ID�iCS�j
            sql = sql & .CRYINDCSCS & "', '"        '���FLG�iCS�j
            sql = sql & .CRYRESCSCS & "', "         '����FLG�iCS�j
            sql = sql & .CRYSMPLIDGDCS & ", '"      '�T���v��ID�iGD�j
            sql = sql & .CRYINDGDCS & "', '"        '���FLG�iGD�j
            sql = sql & .CRYRESGDCS & "', "         '����FLG�iGD�j
            sql = sql & .CRYSMPLIDTCS & ", '"       '�T���v��ID�iT�j
            sql = sql & .CRYINDTCS & "', '"         '���FLG�iT�j
            sql = sql & .CRYRESTCS & "', "          '����FLG�iT�j
            sql = sql & .CRYSMPLIDEPCS & ", '"      '�T���v��ID�iEPD�j
            sql = sql & .CRYINDEPCS & "', '"        '���FLG�iEPD�j
            sql = sql & .CRYRESEPCS & "', "         '����FLG�iEPD�j
            
            '>>>>> X������ǉ��Ή� 2009/07/28 SETsw kubota ---------------
            sql = sql & .CRYSMPLIDXCS               '�T���v��ID�iX���j
            '���FLG�iX���j
            If .CRYINDXCS <> "" And left(.CRYINDXCS, 1) <> vbNullChar Then
                sql = sql & ",'" & .CRYINDXCS & "'"
            Else
                sql = sql & ",'0'"
            End If
            '����FLG�iX���j
            If .CRYRESXCS <> "" And left(.CRYRESXCS, 1) <> vbNullChar Then
                sql = sql & ",'" & .CRYRESXCS & "'"
            Else
                sql = sql & ",'0'"
            End If
            sql = sql & ", "
            '<<<<< X������ǉ��Ή� 2009/07/28 SETsw kubota ---------------
            
            sql = sql & .SMPLNUMCS & ", "           '�T���v������
            sql = sql & "' ', '"                    '�T���v���p�^�[��
            sql = sql & .TSTAFFCS & "', "           '�o�^�Ј�ID
            sql = sql & "sysdate, '"                '�o�^���t
            sql = sql & .KSTAFFCS & "', "           '�X�V�Ј�ID
            sql = sql & "sysdate, "                 '�X�V���t
            sql = sql & "'0', "                     '���M�t���O
            sql = sql & "sysdate)"                  '���M���t
        End With
        '' WriteDBLog sql
        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_CrySmp_Ins = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
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
    DBDRV_CrySmp_Ins = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'�T�v      :WF�T���v���Ǘ��̑}���^�X�V
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:WfSmpOld�@�@�@,I  ,typ_XSDCW   �@   ,�V�T���v���Ǘ��iSXL�j�i���j
'      �@�@:WfSmpNew�@�@�@,I  ,typ_XSDCW   �@   ,�V�T���v���Ǘ��iSXL�j�i�V�j
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :�Â����R�[�h���݂čX�V���}�����𔻕ʂ���
'����      :2001/07/12  �쐬 ���{
Public Function DBDRV_WfSmp_UpdIns(WfSmpOld() As typ_XSDCW, WfSmpNew() As typ_XSDCW) As FUNCTION_RETURN

    Dim sql As String
    Dim lFlg As Boolean
    Dim i As Long
    Dim j As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_WfSmp_UpdIns"

    DBDRV_WfSmp_UpdIns = FUNCTION_RETURN_SUCCESS

    For i = 1 To UBound(WfSmpNew)
        With WfSmpNew(i)
            lFlg = False
            For j = 1 To UBound(WfSmpOld)
                If WfSmpOld(j).XTALCW = .XTALCW And _
                   WfSmpOld(j).INPOSCW = .INPOSCW And _
                   WfSmpOld(j).SMPKBNCW = .SMPKBNCW Then
'                    sql = "update TBCME044 set "
'                    sql = sql & "SMPLID='" & .SMPLID & "', "        ' �T���v��ID
'                    sql = sql & "HINBAN='" & .HINBAN & "', "        ' �i��
'                    sql = sql & "REVNUM=" & .REVNUM & ", "          ' ���i�ԍ������ԍ�
'                    sql = sql & "FACTORY='" & .factory & "', "      ' �H��
'                    sql = sql & "OPECOND='" & .opecond & "', "      ' ���Ə���
'                    sql = sql & "KTKBN='" & .KTKBN & "', "          ' �m��敪
'                    sql = sql & "WFINDRS='" & .WFINDRS & "', "      ' WF�����w���iRs)
'                    sql = sql & "WFINDOI='" & .WFINDOI & "', "      ' WF�����w���iOi)
'                    sql = sql & "WFINDB1='" & .WFINDB1 & "', "      ' WF�����w���iB1)
'                    sql = sql & "WFINDB2='" & .WFINDB2 & "', "      ' WF�����w���iB2)
'                    sql = sql & "WFINDB3='" & .WFINDB3 & "', "      ' WF�����w���iB3)
'                    sql = sql & "WFINDL1='" & .WFINDL1 & "', "      ' WF�����w���iL1)
'                    sql = sql & "WFINDL2='" & .WFINDL2 & "', "      ' WF�����w���iL2)
'                    sql = sql & "WFINDL3='" & .WFINDL3 & "', "      ' WF�����w���iL3)
'                    sql = sql & "WFINDL4='" & .WFINDL4 & "', "      ' WF�����w���iL4)
'                    sql = sql & "WFINDDS='" & .WFINDDS & "', "      ' WF�����w���iDS)
'                    sql = sql & "WFINDDZ='" & .WFINDDZ & "', "      ' WF�����w���iDZ)
'                    sql = sql & "WFINDSP='" & .WFINDSP & "', "      ' WF�����w���iSP)
'                    sql = sql & "WFINDDO1='" & .WFINDDO1 & "', "    ' WF�����w���iDO1)
'                    sql = sql & "WFINDDO2='" & .WFINDDO2 & "', "    ' WF�����w���iDO2)
'                    sql = sql & "WFINDDO3='" & .WFINDDO3 & "', "    ' WF�����w���iDO3)
'                    'add start 2003/05/21 hitec)matsumoto -------------------------
'                    sql = sql & "WFINDOT1='" & .WFINDOT1 & "', "    ' WF�����w���iOT1)
'                    sql = sql & "WFINDOT2='" & .WFINDOT2 & "', "    ' WF�����w���iOT2)
'                    'add end   2003/05/21 hitec)matsumoto -------------------------
'                    sql = sql & "WFRESRS='" & .WFRESRS & "', "      ' WF�������сiRs)
'                    sql = sql & "WFRESOI='" & .WFRESOI & "', "      ' WF�������сiOi)
'                    sql = sql & "WFRESB1='" & .WFRESB1 & "', "      ' WF�������сiB1)
'                    sql = sql & "WFRESB2='" & .WFRESB2 & "', "      ' WF�������сiB2)
'                    sql = sql & "WFRESB3='" & .WFRESB3 & "', "      ' WF�������сiB3)
'                    sql = sql & "WFRESL1='" & .WFRESL1 & "', "      ' WF�������сiL1)
'                    sql = sql & "WFRESL2='" & .WFRESL2 & "', "      ' WF�������сiL2)
'                    sql = sql & "WFRESL3='" & .WFRESL3 & "', "      ' WF�������сiL3)
'                    sql = sql & "WFRESL4='" & .WFRESL4 & "', "      ' WF�������сiL4)
'                    sql = sql & "WFRESDS='" & .WFRESDS & "', "      ' WF�������сiDS)
'                    sql = sql & "WFRESDZ='" & .WFRESDZ & "', "      ' WF�������сiDZ)
'                    sql = sql & "WFRESSP='" & .WFRESSP & "', "      ' WF�������сiSP)
'                    sql = sql & "WFRESDO1='" & .WFRESDO1 & "', "    ' WF�������сiDO1)
'                    sql = sql & "WFRESDO2='" & .WFRESDO2 & "', "    ' WF�������сiDO2)
'                    sql = sql & "WFRESDO3='" & .WFRESDO3 & "', "    ' WF�������сiDO3)
'                    'add start 2003/05/21 hitec)matsumoto -------------------------
'                    sql = sql & "WFRESOT1='" & .WFRESOT1 & "', "    ' WF�����w���iOT1)
'                    sql = sql & "WFRESOT2='" & .WFRESOT2 & "', "    ' WF�����w���iOT2)
'                    'add end   2003/05/21 hitec)matsumoto -------------------------
'                    sql = sql & "UPDDATE=sysdate, "
'                    sql = sql & "SENDFLAG='0'"
'                    sql = sql & " where CRYNUM='" & .CRYNUM & "'"
'                    sql = sql & " and INGOTPOS=" & .INGOTPOS
'                    sql = sql & " and SMPKBN='" & .SMPKBN & "'"

                    sql = "update XSDCW set "
                    sql = sql & "SXLIDCW='" & .SXLIDCW & "', "          ' SXLID
                    sql = sql & "REPSMPLIDCW='" & .REPSMPLIDCW & "', "  ' �T���v��ID
                    sql = sql & "HINBCW='" & .HINBCW & "', "            ' �i��
                    sql = sql & "REVNUMCW=" & .REVNUMCW & ", "          ' ���i�ԍ������ԍ�
                    sql = sql & "FACTORYCW='" & .FACTORYCW & "', "      ' �H��
                    sql = sql & "OPECW='" & .OPECW & "', "              ' ���Ə���
                    sql = sql & "KTKBNCW='" & .KTKBNCW & "', "          ' �m��敪
                    sql = sql & "WFINDRSCW='" & .WFINDRSCW & "', "      ' ���FLG�iRs)
                    sql = sql & "WFRESRS1CW='" & .WFRESRS1CW & "', "    ' ����FLG1�iRs)
                    sql = sql & "WFINDOICW='" & .WFINDOICW & "', "      ' ���FLG�iOi)
                    sql = sql & "WFRESOICW='" & .WFRESOICW & "', "      ' ����FLG�iOi)
                    sql = sql & "WFINDB1CW='" & .WFINDB1CW & "', "      ' ���FLG�iB1)
                    sql = sql & "WFRESB1CW='" & .WFRESB1CW & "', "      ' ����FLG�iB1)
                    sql = sql & "WFINDB2CW='" & .WFINDB2CW & "', "      ' ���FLG�iB2)
                    sql = sql & "WFRESB2CW='" & .WFRESB2CW & "', "      ' ����FLG�iB2)
                    sql = sql & "WFINDB3CW='" & .WFINDB3CW & "', "      ' ���FLG�iB3)
                    sql = sql & "WFRESB3CW='" & .WFRESB3CW & "', "      ' ����FLG�iB3)
                    sql = sql & "WFINDL1CW='" & .WFINDL1CW & "', "      ' ���FLG�iL1)
                    sql = sql & "WFRESL1CW='" & .WFRESL1CW & "', "      ' ����FLG�iL1)
                    sql = sql & "WFINDL2CW='" & .WFINDL2CW & "', "      ' ���FLG�iL2)
                    sql = sql & "WFRESL2CW='" & .WFRESL2CW & "', "      ' ����FLG�iL2)
                    sql = sql & "WFINDL3CW='" & .WFINDL3CW & "', "      ' ���FLG�iL3)
                    sql = sql & "WFRESL3CW='" & .WFRESL3CW & "', "      ' ����FLG�iL3)
                    sql = sql & "WFINDL4CW='" & .WFINDL4CW & "', "      ' ���FLG�iL4)
                    sql = sql & "WFRESL4CW='" & .WFRESL4CW & "', "      ' ����FLG�iL4)
                    sql = sql & "WFINDDSCW='" & .WFINDDSCW & "', "      ' ���FLG�iDS)
                    sql = sql & "WFRESDSCW='" & .WFRESDSCW & "', "      ' ����FLG�iDS)
                    sql = sql & "WFINDDZCW='" & .WFINDDZCW & "', "      ' ���FLG�iDZ)
                    sql = sql & "WFRESDZCW='" & .WFRESDZCW & "', "      ' ����FLG�iDZ)
                    sql = sql & "WFINDSPCW='" & .WFINDSPCW & "', "      ' ���FLG�iSP)
                    sql = sql & "WFRESSPCW='" & .WFRESSPCW & "', "      ' ����FLG�iSP)
                    sql = sql & "WFINDDO1CW='" & .WFINDDO1CW & "', "    ' ���FLG�iDO1)
                    sql = sql & "WFRESDO1CW='" & .WFRESDO1CW & "', "    ' ����FLG�iDO1)
                    sql = sql & "WFINDDO2CW='" & .WFINDDO2CW & "', "    ' ���FLG�iDO2)
                    sql = sql & "WFRESDO2CW='" & .WFRESDO2CW & "', "    ' ����FLG�iDO2)
                    sql = sql & "WFINDDO3CW='" & .WFINDDO3CW & "', "    ' ���FLG�iDO3)
                    sql = sql & "WFRESDO3CW='" & .WFRESDO3CW & "', "    ' ����FLG�iDO3)
                    'add start 2003/05/21 hitec)matsumoto -------------------------
                    sql = sql & "WFINDOT1CW='" & .WFINDOT1CW & "', "    ' ���FLG�iOT1)
                    sql = sql & "WFRESOT1CW='" & .WFRESOT1CW & "', "    ' ����FLG�iOT1)
                    sql = sql & "WFINDOT2CW='" & .WFINDOT2CW & "', "    ' ���FLG�iOT2)
                    sql = sql & "WFRESOT2CW='" & .WFRESOT2CW & "', "    ' ����FLG�iOT2)
                    'add end   2003/05/21 hitec)matsumoto -------------------------
                    '' �c���_�f�ǉ��@03/12/05 ooba START ===============================>
                    sql = sql & "WFINDAOICW='" & .WFINDAOICW & "', "    ' ���FLG (AOI)
                    sql = sql & "WFRESAOICW='" & .WFRESAOICW & "', "    ' ����FLG (AOI)
                    '' �c���_�f�ǉ��@03/12/05 ooba END =================================>
                    '' GD�ǉ��@05/01/17 ooba START =====================================>
                    sql = sql & "WFINDGDCW='" & .WFINDGDCW & "', "    ' ���FLG (GD)
                    sql = sql & "WFRESGDCW='" & .WFRESGDCW & "', "    ' ����FLG (GD)
                    sql = sql & "WFHSGDCW='" & .WFHSGDCW & "', "      ' �ۏ�FLG (GD)
                    '' GD�ǉ��@05/01/17 ooba END =======================================>
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                    sql = sql & "EPINDL1CW = " & .EPINDL1CW & "', "     ' ���FLG (OSF1E)
                    sql = sql & "EPRESL1CW = " & .EPRESL1CW & "', "     ' ����FLG (OSF1E)
                    sql = sql & "EPINDL2CW = " & .EPINDL2CW & "', "     ' ���FLG (OSF2E)
                    sql = sql & "EPRESL2CW = " & .EPRESL2CW & "', "     ' ����FLG (OSF2E)
                    sql = sql & "EPINDL3CW = " & .EPINDL3CW & "', "     ' ���FLG (OSF3E)
                    sql = sql & "EPRESL3CW = " & .EPRESL3CW & "', "     ' ����FLG (OSF3E)
                    sql = sql & "EPINDB1CW = " & .EPINDB1CW & "', "     ' ���FLG (BMD1E)
                    sql = sql & "EPRESB1CW = " & .EPRESB1CW & "', "     ' ����FLG (BMD1E)
                    sql = sql & "EPINDB2CW = " & .EPINDB2CW & "', "     ' ���FLG (BMD2E)
                    sql = sql & "EPRESB2CW = " & .EPRESB2CW & "', "     ' ����FLG (BMD2E)
                    sql = sql & "EPINDB3CW = " & .EPINDB3CW & "', "     ' ���FLG (BMD3E)
                    sql = sql & "EPRESB3CW = " & .EPRESB3CW & "', "     ' ����FLG (BMD3E)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
                    sql = sql & "KDAYCW=sysdate, "
                    sql = sql & "SNDKCW='0'"
                    sql = sql & " where XTALCW='" & .XTALCW & "'"
                    sql = sql & " and INPOSCW=" & .INPOSCW
                    sql = sql & " and SMPKBNCW='" & .SMPKBNCW & "'"

                    '' WriteDBLog sql
                    If OraDB.ExecuteSQL(sql) <= 0 Then
                        DBDRV_WfSmp_UpdIns = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                    lFlg = True
                    Exit For
                End If
            Next j

            If lFlg <> True Then
'                sql = "insert into TBCME044 ("
'                sql = sql & "CRYNUM, "          ' �����ԍ�
'                sql = sql & "INGOTPOS, "        ' �������ʒu
'                sql = sql & "SMPKBN, "          ' �T���v���敪
'                sql = sql & "SMPLID, "          ' �T���v��ID
'                sql = sql & "HINBAN, "          ' �i��
'                sql = sql & "REVNUM, "          ' ���i�ԍ������ԍ�
'                sql = sql & "FACTORY, "         ' �H��
'                sql = sql & "OPECOND, "         ' ���Ə���
'                sql = sql & "KTKBN, "           ' �m��敪
'                sql = sql & "WFINDRS, "         ' WF�����w���iRs)
'                sql = sql & "WFINDOI, "         ' WF�����w���iOi)
'                sql = sql & "WFINDB1, "         ' WF�����w���iB1)
'                sql = sql & "WFINDB2, "         ' WF�����w���iB2)
'                sql = sql & "WFINDB3, "         ' WF�����w���iB3)
'                sql = sql & "WFINDL1, "         ' WF�����w���iL1)
'                sql = sql & "WFINDL2, "         ' WF�����w���iL2)
'                sql = sql & "WFINDL3, "         ' WF�����w���iL3)
'                sql = sql & "WFINDL4, "         ' WF�����w���iL4)
'                sql = sql & "WFINDDS, "         ' WF�����w���iDS)
'                sql = sql & "WFINDDZ, "         ' WF�����w���iDZ)
'                sql = sql & "WFINDSP, "         ' WF�����w���iSP)
'                sql = sql & "WFINDDO1, "        ' WF�����w���iDO1)
'                sql = sql & "WFINDDO2, "        ' WF�����w���iDO2)
'                sql = sql & "WFINDDO3, "        ' WF�����w���iDO3)
'               'add start 2003/05/21 hitec)matsumoto -------------------------
'                sql = sql & "WFINDOT1, "        ' WF�����w���iOT1)
'                sql = sql & "WFINDOT2, "        ' WF�����w���iOT2)
'               'add end   2003/05/21 hitec)matsumoto -------------------------
'                sql = sql & "WFRESRS, "         ' WF�������сiRs)
'                sql = sql & "WFRESOI, "         ' WF�������сiOi)
'                sql = sql & "WFRESB1, "         ' WF�������сiB1)
'                sql = sql & "WFRESB2, "         ' WF�������сiB2)
'                sql = sql & "WFRESB3, "         ' WF�������сiB3)
'                sql = sql & "WFRESL1, "         ' WF�������сiL1)
'                sql = sql & "WFRESL2, "         ' WF�������сiL2)
'                sql = sql & "WFRESL3, "         ' WF�������сiL3)
'                sql = sql & "WFRESL4, "         ' WF�������сiL4)
'                sql = sql & "WFRESDS, "         ' WF�������сiDS)
'                sql = sql & "WFRESDZ, "         ' WF�������сiDZ)
'                sql = sql & "WFRESSP, "         ' WF�������сiSP)
'                sql = sql & "WFRESDO1, "        ' WF�������сiDO1)
'                sql = sql & "WFRESDO2, "        ' WF�������сiDO2)
'                sql = sql & "WFRESDO3, "        ' WF�������сiDO3)
'               'add start 2003/05/21 hitec)matsumoto -------------------------
'                sql = sql & "WFRESOT1, "        ' WF�������сiOT1)
'                sql = sql & "WFRESOT2, "        ' WF�������сiOT2)
'               'add end   2003/05/21 hitec)matsumoto -------------------------
'                sql = sql & "REGDATE, "         ' �X�V���t
'                sql = sql & "UPDDATE, "         ' �X�V���t
'                sql = sql & "SENDFLAG, "        ' ���M�t���O
'                sql = sql & "SENDDATE)"         ' ���M���t
'                sql = sql & " values ('"
'                sql = sql & .XTALCW & "', "     ' �����ԍ�
'                sql = sql & .INPOSCW & ", '"   ' �������ʒu
'                sql = sql & .SMPKBNCW & "', '"    ' �T���v���敪
'                sql = sql & .REPSMPLIDCW & "', '"    ' �T���v��ID
'                sql = sql & .HINBCW & "', "     ' �i��
'                sql = sql & .REVNUMCW & ", '"     ' ���i�ԍ������ԍ�
'                sql = sql & .FACTORYCW & "', '"   ' �H��
'                sql = sql & .OPECW & "', '"   ' ���Ə���
'                sql = sql & .KTKBNCW & "', '"     ' �m��敪
'                sql = sql & .WFINDRSCW & "', '"   ' WF�����w���iRs)
'                sql = sql & .WFINDOICW & "', '"   ' WF�����w���iOi)
'                sql = sql & .WFINDB1CW & "', '"   ' WF�����w���iB1)
'                sql = sql & .WFINDB2CW & "', '"   ' WF�����w���iB2)
'                sql = sql & .WFINDB3CW & "', '"   ' WF�����w���iB3)
'                sql = sql & .WFINDL1CW & "', '"   ' WF�����w���iL1)
'                sql = sql & .WFINDL2CW & "', '"   ' WF�����w���iL2)
'                sql = sql & .WFINDL3CW & "', '"   ' WF�����w���iL3)
'                sql = sql & .WFINDL4CW & "', '"   ' WF�����w���iL4)
'                sql = sql & .WFINDDSCW & "', '"   ' WF�����w���iDS)
'                sql = sql & .WFINDDZCW & "', '"   ' WF�����w���iDZ)
'                sql = sql & .WFINDSPCW & "', '"   ' WF�����w���iSP)
'                sql = sql & .WFINDDO1CW & "', '"  ' WF�����w���iDO1)
'                sql = sql & .WFINDDO2CW & "', '"  ' WF�����w���iDO2)
'                sql = sql & .WFINDDO3CW & "', '"  ' WF�����w���iDO3)
'                'add start 2003/05/21 hitec)matsumoto -------------------------
'                sql = sql & .WFINDOT1CW & "', '"  ' WF�����w���iOT1)
'                sql = sql & .WFINDOT2CW & "', '"  ' WF�����w���iOT2)
'                'add end   2003/05/21 hitec)matsumoto -------------------------
'                sql = sql & .WFRESRS1CW & "', '"   ' WF�������сiRs)
'                sql = sql & .WFRESOICW & "', '"   ' WF�������сiOi)
'                sql = sql & .WFRESB1CW & "', '"   ' WF�������сiB1)
'                sql = sql & .WFRESB2CW & "', '"   ' WF�������сiB2)
'                sql = sql & .WFRESB3CW & "', '"   ' WF�������сiB3)
'                sql = sql & .WFRESL1CW & "', '"   ' WF�������сiL1)
'                sql = sql & .WFRESL2CW & "', '"   ' WF�������сiL2)
'                sql = sql & .WFRESL3CW & "', '"   ' WF�������сiL3)
'                sql = sql & .WFRESL4CW & "', '"   ' WF�������сiL4)
'                sql = sql & .WFRESDSCW & "', '"   ' WF�������сiDS)
'                sql = sql & .WFRESDZCW & "', '"   ' WF�������сiDZ)
'                sql = sql & .WFRESSPCW & "', '"   ' WF�������сiSP)
'                sql = sql & .WFRESDO1CW & "', '"  ' WF�������сiDO1)
'                sql = sql & .WFRESDO2CW & "', '"  ' WF�������сiDO2)
'                sql = sql & .WFRESDO3CW & "', '"   ' WF�������сiDO3)
'                'add start 2003/05/21 hitec)matsumoto -------------------------
'                sql = sql & .WFRESOT1CW & "', '"  ' WF�������сiOT1)
'                sql = sql & .WFRESOT2CW & "',"  ' WF�������сiOT2)
'                'add end   2003/05/21 hitec)matsumoto -------------------------
'                sql = sql & "sysdate, "         ' �o�^���t
'                sql = sql & "sysdate, "         ' �X�V���t
'                sql = sql & "'0', "             ' ���M�t���O
'                sql = sql & "sysdate)"          ' ���M���t


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
                sql = sql & "WFSMPLIDOT2CW, "   ' �T���v��ID�iOT2�j
               'add start 2003/05/21 hitec)matsumoto -------------------------
                sql = sql & "WFINDOT1CW, "      ' ���FLG�iOT1)
                sql = sql & "WFRESOT1CW, "      ' ����FLG�iOT1)
                sql = sql & "WFINDOT2CW, "      ' ���FLG�iOT2)
                sql = sql & "WFRESOT2CW, "      ' ����FLG�iOT2)
               'add end   2003/05/21 hitec)matsumoto -------------------------
                sql = sql & "WFSMPLIDAOICW, "   ' �T���v��ID�iAOi�j
                sql = sql & "WFINDAOICW, "      ' ���FLG�iAOi�j
                sql = sql & "WFRESAOICW, "      ' ����FLG�iAOi�j
                sql = sql & "SMPLNUMCW, "       ' �T���v������
                sql = sql & "SMPLPATCW, "       ' �T���v���p�^�[��
                sql = sql & "LIVKCW,"           ' �����敪
                sql = sql & "TSTAFFCW,"         ' �o�^�Ј�ID
                sql = sql & "TDAYCW, "          ' �o�^���t
                sql = sql & "KSTAFFCW, "        ' �X�V�Ј�ID
                sql = sql & "KDAYCW, "          ' �X�V���t
                sql = sql & "SNDKCW, "          ' ���M�t���O
'                sql = sql & "SNDDAYCW)"         ' ���M���t
                '' GD�ǉ��@05/01/17 ooba START =====================================>
                sql = sql & "SNDDAYCW, "        ' ���M���t
                sql = sql & "WFSMPLIDGDCW, "    ' �T���v��ID (GD)
                sql = sql & "WFINDGDCW, "       ' ���FLG (GD)
                sql = sql & "WFRESGDCW, "       ' ����FLG (GD)
                sql = sql & "WFHSGDCW"         ' �ۏ�FLG (GD)
                '' GD�ǉ��@05/01/17 ooba END =======================================>
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                sql = sql & ", EPSMPLIDB1CW, "  ' �T���v��ID (BMD1E)
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
                sql = sql & "EPRESL3CW"         ' ����FLG (OSF3E)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
                sql = sql & " )"
                sql = sql & " values ('"
                sql = sql & .SXLIDCW & "', '"       ' SXLID
                sql = sql & .SMPKBNCW & "', '"      ' �T���v���敪
                sql = sql & .TBKBNCW & "', '"       ' T/B�敪
                sql = sql & .REPSMPLIDCW & "', '"   ' �T���v��ID
                sql = sql & .XTALCW & "', "         ' �����ԍ�
                sql = sql & .INPOSCW & ", '"        ' �������ʒu
                sql = sql & .HINBCW & "', "         ' �i��
                sql = sql & .REVNUMCW & ", '"       ' ���i�ԍ������ԍ�
                sql = sql & .FACTORYCW & "', '"     ' �H��
                sql = sql & .OPECW & "', '"         ' ���Ə���
                sql = sql & .KTKBNCW & "', '"       ' �m��敪
                sql = sql & .SMCRYNUMCW & "', '"    ' �T���v���u���b�NID
                sql = sql & .WFSMPLIDRSCW & "', '"  ' �T���v��ID�iRs�j
                sql = sql & .WFSMPLIDRS1CW & "', '" ' ����T���v��ID1�iRs�j
                sql = sql & .WFSMPLIDRS2CW & "', '" ' ����T���v��ID2�iRs�j
                sql = sql & .WFINDRSCW & "', '"     ' ���FLG�iRs)
                sql = sql & .WFRESRS1CW & "', '"    ' ����FLG1�iRs)
                sql = sql & .WFRESRS2CW & "', '"    ' ����FLG2�iRs)
                sql = sql & .WFSMPLIDOICW & "', '"  ' �T���v��ID�iOi�j
                sql = sql & .WFINDOICW & "', '"     ' ���FLG�iOi)
                sql = sql & .WFRESOICW & "', '"     ' ����FLG�iOi)
                sql = sql & .WFSMPLIDB1CW & "', '"  ' �T���v��ID�iB1�j
                sql = sql & .WFINDB1CW & "', '"     ' ���FLG�iB1)
                sql = sql & .WFRESB1CW & "', '"     ' ����FLG�iB1)
                sql = sql & .WFSMPLIDB2CW & "', '"  ' �T���v��ID�iB2�j
                sql = sql & .WFINDB2CW & "', '"     ' ���FLG�iB2)
                sql = sql & .WFRESB2CW & "', '"     ' ����FLG�iB2)
                sql = sql & .WFSMPLIDB3CW & "', '"  ' �T���v��ID�iB3�j
                sql = sql & .WFINDB3CW & "', '"     ' ���FLG�iB3)
                sql = sql & .WFRESB3CW & "', '"     ' ����FLG�iB3)
                sql = sql & .WFSMPLIDL1CW & "', '"  ' �T���v��ID�iL1�j
                sql = sql & .WFINDL1CW & "', '"     ' ���FLG�iL1)
                sql = sql & .WFRESL1CW & "', '"     ' ����FLG�iL1)
                sql = sql & .WFSMPLIDL2CW & "', '"  ' �T���v��ID�iL2�j
                sql = sql & .WFINDL2CW & "', '"     ' ���FLG�iL2)
                sql = sql & .WFRESL2CW & "', '"     ' ����FLG�iL2)
                sql = sql & .WFSMPLIDL3CW & "', '"  ' �T���v��ID�iL3�j
                sql = sql & .WFINDL3CW & "', '"     ' ���FLG�iL3)
                sql = sql & .WFRESL3CW & "', '"     ' ����FLG�iL3)
                sql = sql & .WFSMPLIDL4CW & "', '"  ' �T���v��ID�iL4�j
                sql = sql & .WFINDL4CW & "', '"     ' ���FLG�iL4)
                sql = sql & .WFRESL4CW & "', '"     ' ����FLG�iL4)
                sql = sql & .WFSMPLIDDSCW & "', '"  ' �T���v��ID�iDS�j
                sql = sql & .WFINDDSCW & "', '"     ' ���FLG�iDS)
                sql = sql & .WFRESDSCW & "', '"     ' ����FLG�iDS)
                sql = sql & .WFSMPLIDDZCW & "', '"  ' �T���v��ID�iDZ�j
                sql = sql & .WFINDDZCW & "', '"     ' ���FLG�iDZ)
                sql = sql & .WFRESDZCW & "', '"     ' ����FLG�iDZ)
                sql = sql & .WFSMPLIDSPCW & "', '"  ' �T���v��ID�iSP�j
                sql = sql & .WFINDSPCW & "', '"     ' ���FLG�iSP)
                sql = sql & .WFRESSPCW & "', '"     ' ����FLG�iSP)
                sql = sql & .WFSMPLIDDO1CW & "', '" ' �T���v��ID�iDO1�j
                sql = sql & .WFINDDO1CW & "', '"    ' ���FLG�iDO1)
                sql = sql & .WFRESDO1CW & "', '"    ' ����FLG�iDO1)
                sql = sql & .WFSMPLIDDO2CW & "', '" ' �T���v��ID�iDO2�j
                sql = sql & .WFINDDO2CW & "', '"    ' ���FLG�iDO2)
                sql = sql & .WFRESDO2CW & "', '"    ' ����FLG�iDO2)
                sql = sql & .WFSMPLIDDO3CW & "', '" ' �T���v��ID�iDO3�j
                sql = sql & .WFINDDO3CW & "', '"    ' ���FLG�iDO3)
                sql = sql & .WFRESDO3CW & "', '"    ' ����FLG�iDO3)
                sql = sql & .WFSMPLIDOT1CW & "', '" ' �T���v��ID�iOT1�j
                sql = sql & .WFSMPLIDOT2CW & "', '" ' �T���v��ID�iOT2�j
                'add start 2003/05/21 hitec)matsumoto -------------------------
                sql = sql & .WFINDOT1CW & "', '"    ' ���FLG�iOT1)
                sql = sql & .WFRESOT1CW & "', '"    ' ����FLG�iOT1)
                sql = sql & .WFINDOT2CW & "', '"    ' ���FLG�iOT2)
                sql = sql & .WFRESOT2CW & "', '"    ' ����FLG�iOT2)
                'add end   2003/05/21 hitec)matsumoto -------------------------
                sql = sql & .WFSMPLIDAOICW & "', '" ' �T���v��ID�iAOi�j
                sql = sql & .WFINDAOICW & "', '"    ' ���FLG�iAOi�j
                sql = sql & .WFRESAOICW & "', "     ' ����FLG�iAOi�j
                sql = sql & .SMPLNUMCW & ", '"      ' �T���v������
                sql = sql & .SMPLPATCW & "', '"     ' �T���v���p�^�[��
                sql = sql & .LIVKCW & "', '"        ' �����敪
                sql = sql & .TSTAFFCW & "', "       ' �o�^�Ј�ID
                sql = sql & "sysdate, '"            ' �o�^���t
                sql = sql & .KSTAFFCW & "', "       ' �X�V�Ј�ID
                sql = sql & "sysdate, "             ' �X�V���t
                sql = sql & "'0', "                 ' ���M�t���O
'                sql = sql & "sysdate)"              ' ���M���t
                '' GD�ǉ��@05/01/17 ooba START =====================================>
                sql = sql & "sysdate, '"            ' ���M���t
                sql = sql & .WFSMPLIDGDCW & "', '"  ' �T���v��ID (GD)
                sql = sql & .WFINDGDCW & "', '"     ' ���FLG (GD)
                sql = sql & .WFRESGDCW & "', '"     ' ����FLG (GD)
                sql = sql & .WFHSGDCW & "', '"      ' �ۏ�FLG (GD)
                '' GD�ǉ��@05/01/17 ooba END =======================================>
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                sql = sql & .EPSMPLIDB1CW & "', '"  ' �T���v��ID (BMD1E)
                sql = sql & .EPINDB1CW & "', '"     ' ���FLG (BMD1E)
                sql = sql & .EPRESB1CW & "', '"     ' ����FLG (BMD1E)
                sql = sql & .EPSMPLIDB2CW & "', '"  ' �T���v��ID (BMD2E)
                sql = sql & .EPINDB2CW & "', '"     ' ���FLG (BMD2E)
                sql = sql & .EPRESB2CW & "', '"     ' ����FLG (BMD2E)
                sql = sql & .EPSMPLIDB3CW & "', '"  ' �T���v��ID (BMD3E)
                sql = sql & .EPINDB3CW & "', '"     ' ���FLG (BMD3E)
                sql = sql & .EPRESB3CW & "', '"       ' ����FLG (BMD3E)
                sql = sql & .EPSMPLIDL1CW & "', '"  ' �T���v��ID (OSF1E)
                sql = sql & .EPINDL1CW & "', '"     ' ���FLG (OSF1E)
                sql = sql & .EPRESL1CW & "', '"     ' ����FLG (OSF1E)
                sql = sql & .EPSMPLIDL2CW & "', '"  ' �T���v��ID (OSF2E)
                sql = sql & .EPINDL2CW & "', '"     ' ���FLG (OSF2E)
                sql = sql & .EPRESL2CW & "', '"     ' ����FLG (OSF2E)
                sql = sql & .EPSMPLIDL3CW & "', '"  ' �T���v��ID (OSF3E)
                sql = sql & .EPINDL3CW & "', '"     ' ���FLG (OSF3E)
                sql = sql & .EPRESL3CW & "')"       ' ����FLG (OSF3E)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

                '' WriteDBLog sql
                If OraDB.ExecuteSQL(sql) <= 0 Then
                    DBDRV_WfSmp_UpdIns = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            End If
        End With
    Next i

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_WfSmp_UpdIns = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function


'�T�v      :WF�T���v���Ǘ��̑}��
'���Ұ��@�@:�ϐ���      ,IO ,�^                 ,����
'      �@�@:WFSMP �@�@�@,I  ,typ_XSDCW   �@     ,�V�T���v���Ǘ��iSXL�j
'      �@�@:�߂�l      ,O  ,FUNCTION_RETURN�@  ,�������݂̐���
'����      :DBDRV_WfSmp_UpdIns�Ɉڍs����\��
'����      :2001/07/12  �쐬 ���{
Public Function DBDRV_WfSmp_INS(WFSMP() As typ_XSDCW) As FUNCTION_RETURN

    Dim sql     As String
    Dim i       As Long
    Dim sDbName As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_WfSmp_INS"

    DBDRV_WfSmp_INS = FUNCTION_RETURN_SUCCESS

    sDbName = "XSDCW"
    For i = 1 To UBound(WFSMP)
        With WFSMP(i)
'            sql = "insert into TBCME044 ("
'            sql = sql & "CRYNUM, "          ' �����ԍ�
'            sql = sql & "INGOTPOS, "        ' �������ʒu
'            sql = sql & "SMPKBN, "          ' �T���v���敪
'            sql = sql & "SMPLID, "          ' �T���v��ID
'            sql = sql & "HINBAN, "          ' �i��
'            sql = sql & "REVNUM, "          ' ���i�ԍ������ԍ�
'            sql = sql & "FACTORY, "         ' �H��
'            sql = sql & "OPECOND, "         ' ���Ə���
'            sql = sql & "KTKBN, "           ' �m��敪
'            sql = sql & "WFINDRS, "         ' WF�����w���iRs)
'            sql = sql & "WFINDOI, "         ' WF�����w���iOi)
'            sql = sql & "WFINDB1, "         ' WF�����w���iB1)
'            sql = sql & "WFINDB2, "         ' WF�����w���iB2)
'            sql = sql & "WFINDB3, "         ' WF�����w���iB3)
'            sql = sql & "WFINDL1, "         ' WF�����w���iL1)
'            sql = sql & "WFINDL2, "         ' WF�����w���iL2)
'            sql = sql & "WFINDL3, "         ' WF�����w���iL3)
'            sql = sql & "WFINDL4, "         ' WF�����w���iL4)
'            sql = sql & "WFINDDS, "         ' WF�����w���iDS)
'            sql = sql & "WFINDDZ, "         ' WF�����w���iDZ)
'            sql = sql & "WFINDSP, "         ' WF�����w���iSP)
'            sql = sql & "WFINDDO1, "        ' WF�����w���iDO1)
'            sql = sql & "WFINDDO2, "        ' WF�����w���iDO2)
'            sql = sql & "WFINDDO3, "        ' WF�����w���iDO3)
'            'add start 2003/05/21 hitec)matsumoto -------------------------
'            sql = sql & "WFINDOT1, "        ' WF�����w���iOT1)
'            sql = sql & "WFINDOT2, "        ' WF�����w���iOT2)
'            'add end   2003/05/21 hitec)matsumoto -------------------------
'            sql = sql & "WFRESRS, "         ' WF�������сiRs)
'            sql = sql & "WFRESOI, "         ' WF�������сiOi)
'            sql = sql & "WFRESB1, "         ' WF�������сiB1)
'            sql = sql & "WFRESB2, "         ' WF�������сiB2)
'            sql = sql & "WFRESB3, "         ' WF�������сiB3)
'            sql = sql & "WFRESL1, "         ' WF�������сiL1)
'            sql = sql & "WFRESL2, "         ' WF�������сiL2)
'            sql = sql & "WFRESL3, "         ' WF�������сiL3)
'            sql = sql & "WFRESL4, "         ' WF�������сiL4)
'            sql = sql & "WFRESDS, "         ' WF�������сiDS)
'            sql = sql & "WFRESDZ, "         ' WF�������сiDZ)
'            sql = sql & "WFRESSP, "         ' WF�������сiSP)
'            sql = sql & "WFRESDO1, "        ' WF�������сiDO1)
'            sql = sql & "WFRESDO2, "        ' WF�������сiDO2)
'            sql = sql & "WFRESDO3, "        ' WF�������сiDO3)
'            'add start 2003/05/21 hitec)matsumoto -------------------------
'            sql = sql & "WFRESOT1, "        ' WF�������сiOT1)
'            sql = sql & "WFRESOT2, "        ' WF�������сiOT2)
'            'add end   2003/05/21 hitec)matsumoto -------------------------
'            sql = sql & "REGDATE, "         ' �X�V���t
'            sql = sql & "UPDDATE, "         ' �X�V���t
'            sql = sql & "SENDFLAG, "        ' ���M�t���O
'            sql = sql & "SENDDATE)"         ' ���M���t
'            sql = sql & " values ('"
'            sql = sql & .CRYNUM & "', "     ' �����ԍ�
'            sql = sql & .INGOTPOS & ", '"   ' �������ʒu
'            sql = sql & .SMPKBN & "', '"    ' �T���v���敪
'            sql = sql & .SMPLID & "', '"    ' �T���v��ID
'            sql = sql & .HINBAN & "', "     ' �i��
'            sql = sql & .REVNUM & ", '"     ' ���i�ԍ������ԍ�
'            sql = sql & .factory & "', '"   ' �H��
'            sql = sql & .opecond & "', '"   ' ���Ə���
'            sql = sql & .KTKBN & "', '"     ' �m��敪
'            sql = sql & .WFINDRS & "', '"   ' WF�����w���iRs)
'            sql = sql & .WFINDOI & "', '"   ' WF�����w���iOi)
'            sql = sql & .WFINDB1 & "', '"   ' WF�����w���iB1)
'            sql = sql & .WFINDB2 & "', '"   ' WF�����w���iB2)
'            sql = sql & .WFINDB3 & "', '"   ' WF�����w���iB3)
'            sql = sql & .WFINDL1 & "', '"   ' WF�����w���iL1)
'            sql = sql & .WFINDL2 & "', '"   ' WF�����w���iL2)
'            sql = sql & .WFINDL3 & "', '"   ' WF�����w���iL3)
'            sql = sql & .WFINDL4 & "', '"   ' WF�����w���iL4)
'            sql = sql & .WFINDDS & "', '"   ' WF�����w���iDS)
'            sql = sql & .WFINDDZ & "', '"   ' WF�����w���iDZ)
'            sql = sql & .WFINDSP & "', '"   ' WF�����w���iSP)
'            sql = sql & .WFINDDO1 & "', '"  ' WF�����w���iDO1)
'            sql = sql & .WFINDDO2 & "', '"  ' WF�����w���iDO2)
'            sql = sql & .WFINDDO3 & "', '"  ' WF�����w���iDO3)
'            'add start 2003/05/21 hitec)matsumoto -------------------------
'            sql = sql & .WFINDOT1 & "', '"  ' WF�����w���iOT1)
'            sql = sql & .WFINDOT2 & "', '"  ' WF�����w���iOT2)
'            'add end   2003/05/21 hitec)matsumoto -------------------------
'            sql = sql & .WFRESRS & "', '"   ' WF�������сiRs)
'            sql = sql & .WFRESOI & "', '"   ' WF�������сiOi)
'            sql = sql & .WFRESB1 & "', '"   ' WF�������сiB1)
'            sql = sql & .WFRESB2 & "', '"   ' WF�������сiB2)
'            sql = sql & .WFRESB3 & "', '"   ' WF�������сiB3)
'            sql = sql & .WFRESL1 & "', '"   ' WF�������сiL1)
'            sql = sql & .WFRESL2 & "', '"   ' WF�������сiL2)
'            sql = sql & .WFRESL3 & "', '"   ' WF�������сiL3)
'            sql = sql & .WFRESL4 & "', '"   ' WF�������сiL4)
'            sql = sql & .WFRESDS & "', '"   ' WF�������сiDS)
'            sql = sql & .WFRESDZ & "', '"   ' WF�������сiDZ)
'            sql = sql & .WFRESSP & "', '"   ' WF�������сiSP)
'            sql = sql & .WFRESDO1 & "', '"  ' WF�������сiDO1)
'            sql = sql & .WFRESDO2 & "', '"  ' WF�������сiDO2)
'            sql = sql & .WFRESDO3 & "', '"   ' WF�������сiDO3)
'            'add start 2003/05/21 hitec)matsumoto -------------------------
'            sql = sql & .WFRESOT1 & "', '"  ' WF�������сiOT1)
'            sql = sql & .WFRESOT2 & "',"  ' WF�������сiOT2)
'            'add end   2003/05/21 hitec)matsumoto -------------------------
'            sql = sql & "sysdate, "         ' �o�^���t
'            sql = sql & "sysdate, "         ' �X�V���t
'            sql = sql & "'0', "             ' ���M�t���O
'            sql = sql & "sysdate)"          ' ���M���t


                sql = "insert into XSDCW ("
                sql = sql & "SXLIDCW, "             ' SXLID
                sql = sql & "SMPKBNCW, "            ' �T���v���敪
                sql = sql & "TBKBNCW, "             ' T/B�敪
                sql = sql & "REPSMPLIDCW, "         ' �T���v��ID
                sql = sql & "XTALCW, "              ' �����ԍ�
                sql = sql & "INPOSCW, "             ' �������ʒu
                sql = sql & "HINBCW, "              ' �i��
                sql = sql & "REVNUMCW, "            ' ���i�ԍ������ԍ�
                sql = sql & "FACTORYCW, "           ' �H��
                sql = sql & "OPECW, "               ' ���Ə���
                sql = sql & "KTKBNCW, "             ' �m��敪
                sql = sql & "SMCRYNUMCW, "          ' �T���v���u���b�NID
                sql = sql & "WFSMPLIDRSCW, "        ' �T���v��ID(Rs)
                sql = sql & "WFSMPLIDRS1CW, "       ' ����T���v��ID1�iRs�j
                sql = sql & "WFSMPLIDRS2CW, "       ' ����T���v��ID2�iRs�j
                sql = sql & "WFINDRSCW, "           ' ���FLG�iRs)
                sql = sql & "WFRESRS1CW, "          ' ����FLG1�iRs)
                sql = sql & "WFRESRS2CW, "          ' ����FLG2�iRs)
                sql = sql & "WFSMPLIDOICW, "        ' �T���v��ID�iOi�j
                sql = sql & "WFINDOICW, "           ' ���FLG�iOi)
                sql = sql & "WFRESOICW, "           ' ����FLG�iOi)
                sql = sql & "WFSMPLIDB1CW, "        ' �T���v��ID�iB1�j
                sql = sql & "WFINDB1CW, "           ' ���FLG�iB1)
                sql = sql & "WFRESB1CW, "           ' ����FLG�iB1)
                sql = sql & "WFSMPLIDB2CW, "        ' �T���v��ID�iB2�j
                sql = sql & "WFINDB2CW, "           ' ���FLG�iB2)
                sql = sql & "WFRESB2CW, "           ' ����FLG�iB2)
                sql = sql & "WFSMPLIDB3CW, "        ' �T���v��ID�iB3�j
                sql = sql & "WFINDB3CW, "           ' ���FLG�iB3)
                sql = sql & "WFRESB3CW, "           ' ����FLG�iB3)
                sql = sql & "WFSMPLIDL1CW, "        ' �T���v��ID�iL1�j
                sql = sql & "WFINDL1CW, "           ' ���FLG�iL1)
                sql = sql & "WFRESL1CW, "           ' ����FLG�iL1)
                sql = sql & "WFSMPLIDL2CW, "        ' �T���v��ID�iL2�j
                sql = sql & "WFINDL2CW, "           ' ���FLG�iL2)
                sql = sql & "WFRESL2CW, "           ' ����FLG�iL2)
                sql = sql & "WFSMPLIDL3CW, "        ' �T���v��ID�iL3�j
                sql = sql & "WFINDL3CW, "           ' ���FLG�iL3)
                sql = sql & "WFRESL3CW, "           ' ����FLG�iL3)
                sql = sql & "WFSMPLIDL4CW, "        ' �T���v��ID�iL4�j
                sql = sql & "WFINDL4CW, "           ' ���FLG�iL4)
                sql = sql & "WFRESL4CW, "           ' ����FLG�iL4)
                sql = sql & "WFSMPLIDDSCW, "        ' �T���v��ID�iDS�j
                sql = sql & "WFINDDSCW, "           ' ���FLG�iDS)
                sql = sql & "WFRESDSCW, "           ' ����FLG�iDS)
                sql = sql & "WFSMPLIDDZCW, "        ' �T���v��ID�iDZ�j
                sql = sql & "WFINDDZCW, "           ' ���FLG�iDZ)
                sql = sql & "WFRESDZCW, "           ' ����FLG�iDZ)
                sql = sql & "WFSMPLIDSPCW, "        ' �T���v��ID�iSP�j
                sql = sql & "WFINDSPCW, "           ' ���FLG�iSP)
                sql = sql & "WFRESSPCW, "           ' ����FLG�iSP)
                sql = sql & "WFSMPLIDDO1CW,"        ' �T���v��ID�iDO1�j
                sql = sql & "WFINDDO1CW, "          ' ���FLG�iDO1)
                sql = sql & "WFRESDO1CW, "          ' ����FLG�iDO1)
                sql = sql & "WFSMPLIDDO2CW, "       ' �T���v��ID�iDO2�j
                sql = sql & "WFINDDO2CW, "          ' ���FLG�iDO2)
                sql = sql & "WFRESDO2CW, "          ' ����FLG�iDO2)
                sql = sql & "WFSMPLIDDO3CW, "       ' �T���v��ID�iDO3�j
                sql = sql & "WFINDDO3CW, "          ' ���FLG�iDO3)
                sql = sql & "WFRESDO3CW, "          ' ����FLG�iDO3)
                sql = sql & "WFSMPLIDOT1CW, "       ' �T���v��ID�iOT1�j
               'add start 2003/05/21 hitec)matsumoto -------------------------
                sql = sql & "WFINDOT1CW, "          ' ���FLG�iOT1)
                sql = sql & "WFRESOT1CW, "          ' ����FLG�iOT1)
                sql = sql & "WFSMPLIDOT2CW, "       ' �T���v��ID�iOT2�j
                sql = sql & "WFINDOT2CW, "          ' ���FLG�iOT2)
                sql = sql & "WFRESOT2CW, "          ' ����FLG�iOT2)
               'add end   2003/05/21 hitec)matsumoto -------------------------
                sql = sql & "WFSMPLIDAOICW, "       ' �T���v��ID�iAOi�j
                sql = sql & "WFINDAOICW, "          ' ���FLG�iAOi�j
                sql = sql & "WFRESAOICW, "          ' ����FLG�iAOi�j
                sql = sql & "SMPLNUMCW, "           ' �T���v������
                sql = sql & "SMPLPATCW, "           ' �T���v���p�^�[��
                sql = sql & "TSTAFFCW,"             ' �o�^�Ј�ID
                sql = sql & "TDAYCW, "              ' �o�^���t
                sql = sql & "KSTAFFCW, "            ' �X�V�Ј�ID
                sql = sql & "KDAYCW, "              ' �X�V���t
                sql = sql & "SNDKCW, "              ' ���M�t���O
                sql = sql & "SNDDAYCW,"             ' ���M���t
'                sql = sql & "LIVKCW)"               ' �����敪
                '' GD�ǉ��@05/01/17 ooba START =====================================>
                sql = sql & "LIVKCW, "              ' �����敪
                sql = sql & "WFSMPLIDGDCW, "        ' �T���v��ID (GD)
                sql = sql & "WFINDGDCW, "           ' ���FLG (GD)
                sql = sql & "WFRESGDCW, "           ' ����FLG (GD)
'                sql = sql & "WFHSGDCW)"             ' �ۏ�FLG (GD) 2006/08/15 Del �G�s��s�]���ǉ��Ή� SMP)kondoh
                sql = sql & "WFHSGDCW, "            ' �ۏ�FLG (GD)  2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
                '' GD�ǉ��@05/01/17 ooba END =======================================>
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                sql = sql & "EPSMPLIDB1CW, "        ' �T���v��ID (BMD1E)
                sql = sql & "EPINDB1CW, "           ' ���FLG (BMD1E)
                sql = sql & "EPRESB1CW, "           ' ����FLG (BMD1E)
                sql = sql & "EPSMPLIDB2CW, "        ' �T���v��ID (BMD2E)
                sql = sql & "EPINDB2CW, "           ' ���FLG (BMD2E)
                sql = sql & "EPRESB2CW, "           ' ����FLG (BMD2E)
                sql = sql & "EPSMPLIDB3CW, "        ' �T���v��ID (BMD3E)
                sql = sql & "EPINDB3CW, "           ' ���FLG (BMD3E)
                sql = sql & "EPRESB3CW, "           ' ����FLG (BMD3E)
                sql = sql & "EPSMPLIDL1CW, "        ' �T���v��ID (OSF1E)
                sql = sql & "EPINDL1CW, "           ' ���FLG (OSF1E)
                sql = sql & "EPRESL1CW, "           ' ����FLG (OSF1E)
                sql = sql & "EPSMPLIDL2CW, "        ' �T���v��ID (OSF2E)
                sql = sql & "EPINDL2CW, "           ' ���FLG (OSF2E)
                sql = sql & "EPRESL2CW, "           ' ����FLG (OSF2E)
                sql = sql & "EPSMPLIDL3CW, "        ' �T���v��ID (OSF3E)
                sql = sql & "EPINDL3CW, "           ' ���FLG (OSF3E)
                sql = sql & "EPRESL3CW"             ' ����FLG (OSF3E)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
                sql = sql & ")"                     ' 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
                sql = sql & " values ('"
                sql = sql & .SXLIDCW & "', '"       ' SXLID
                sql = sql & .SMPKBNCW & "', '"      ' �T���v���敪
                sql = sql & .TBKBNCW & "', '"       ' T/B�敪
                sql = sql & .REPSMPLIDCW & "', '"   ' �T���v��ID
                sql = sql & .XTALCW & "', "         ' �����ԍ�
                sql = sql & .INPOSCW & ", '"        ' �������ʒu
                sql = sql & .HINBCW & "', "         ' �i��
                sql = sql & .REVNUMCW & ", '"       ' ���i�ԍ������ԍ�
                sql = sql & .FACTORYCW & "', '"     ' �H��
                sql = sql & .OPECW & "', '"         ' ���Ə���
                sql = sql & .KTKBNCW & "', '"       ' �m��敪
                sql = sql & .SMCRYNUMCW & "', '"    ' �T���v���u���b�NID
                sql = sql & .WFSMPLIDRSCW & "', '"  ' �T���v��ID�iRs�j
                sql = sql & .WFSMPLIDRS1CW & "', '" ' ����T���v��ID1�iRs�j
'               sql = sql & "Null, "                ' ����T���v��ID1�iRs�j
                sql = sql & .WFSMPLIDRS2CW & "', '" ' ����T���v��ID2�iRs�j
'               sql = sql & "Null, '"               ' ����T���v��ID2�iRs�j
                sql = sql & .WFINDRSCW & "', '"     ' ���FLG�iRs)
                sql = sql & .WFRESRS1CW & "', "     ' ����FLG1�iRs)
                sql = sql & "Null, '"               ' ����FLG2�iRs)
                sql = sql & .WFSMPLIDOICW & "', '"  ' �T���v��ID�iOi�j
                sql = sql & .WFINDOICW & "', '"     ' ���FLG�iOi)
                sql = sql & .WFRESOICW & "', '"     ' ����FLG�iOi)
                sql = sql & .WFSMPLIDB1CW & "', '"  ' �T���v��ID�iB1�j
                sql = sql & .WFINDB1CW & "', '"     ' ���FLG�iB1)
                sql = sql & .WFRESB1CW & "', '"     ' ����FLG�iB1)
                sql = sql & .WFSMPLIDB2CW & "', '"  ' �T���v��ID�iB2�j
                sql = sql & .WFINDB2CW & "', '"     ' ���FLG�iB2)
                sql = sql & .WFRESB2CW & "', '"     ' ����FLG�iB2)
                sql = sql & .WFSMPLIDB3CW & "', '"  ' �T���v��ID�iB3�j
                sql = sql & .WFINDB3CW & "', '"     ' ���FLG�iB3)
                sql = sql & .WFRESB3CW & "', '"     ' ����FLG�iB3)
                sql = sql & .WFSMPLIDL1CW & "', '"  ' �T���v��ID�iL1�j
                sql = sql & .WFINDL1CW & "', '"     ' ���FLG�iL1)
                sql = sql & .WFRESL1CW & "', '"     ' ����FLG�iL1)
                sql = sql & .WFSMPLIDL2CW & "', '"  ' �T���v��ID�iL2�j
                sql = sql & .WFINDL2CW & "', '"     ' ���FLG�iL2)
                sql = sql & .WFRESL2CW & "', '"     ' ����FLG�iL2)
                sql = sql & .WFSMPLIDL3CW & "', '"  ' �T���v��ID�iL3�j
                sql = sql & .WFINDL3CW & "', '"     ' ���FLG�iL3)
                sql = sql & .WFRESL3CW & "', '"     ' ����FLG�iL3)
                sql = sql & .WFSMPLIDL4CW & "', '"  ' �T���v��ID�iL4�j
                sql = sql & .WFINDL4CW & "', '"     ' ���FLG�iL4)
                sql = sql & .WFRESL4CW & "', '"     ' ����FLG�iL4)
                sql = sql & .WFSMPLIDDSCW & "', '"  ' �T���v��ID�iDS�j
                sql = sql & .WFINDDSCW & "', '"     ' ���FLG�iDS)
                sql = sql & .WFRESDSCW & "', '"     ' ����FLG�iDS)
                sql = sql & .WFSMPLIDDZCW & "', '"  ' �T���v��ID�iDZ�j
                sql = sql & .WFINDDZCW & "', '"     ' ���FLG�iDZ)
                sql = sql & .WFRESDZCW & "', '"     ' ����FLG�iDZ)
                sql = sql & .WFSMPLIDSPCW & "', '"  ' �T���v��ID�iSP�j
                sql = sql & .WFINDSPCW & "', '"     ' ���FLG�iSP)
                sql = sql & .WFRESSPCW & "', '"     ' ����FLG�iSP)
                sql = sql & .WFSMPLIDDO1CW & "', '" ' �T���v��ID�iDO1�j
                sql = sql & .WFINDDO1CW & "', '"    ' ���FLG�iDO1)
                sql = sql & .WFRESDO1CW & "', '"    ' ����FLG�iDO1)
                sql = sql & .WFSMPLIDDO2CW & "', '" ' �T���v��ID�iDO2�j
                sql = sql & .WFINDDO2CW & "', '"    ' ���FLG�iDO2)
                sql = sql & .WFRESDO2CW & "', '"    ' ����FLG�iDO2)
                sql = sql & .WFSMPLIDDO3CW & "', '" ' �T���v��ID�iDO3�j
                sql = sql & .WFINDDO3CW & "', '"    ' ���FLG�iDO3)
                sql = sql & .WFRESDO3CW & "', '"    ' ����FLG�iDO3)
                sql = sql & .WFSMPLIDOT1CW & "', '" ' �T���v��ID�iOT1�j
                sql = sql & .WFINDOT1CW & "', '"    ' ���FLG�iOT1)
                sql = sql & .WFRESOT1CW & "', '"    ' ����FLG�iOT1)
                sql = sql & .WFSMPLIDOT2CW & "', '" ' �T���v��ID�iOT2�j
                sql = sql & .WFINDOT2CW & "', '"    ' ���FLG�iOT2)
                sql = sql & .WFRESOT2CW & "', '"    ' ����FLG�iOT2)
                sql = sql & .WFSMPLIDAOICW & "', '" ' �T���v��ID�iAOi�j
''              sql = sql & "NULL, "                ' �T���v��ID�iAOi�j
''              sql = sql & "NULL, "                ' ���FLG�iAOi�j
                sql = sql & .WFINDAOICW & "', '"    ' ���FLG�iAOi�j
                sql = sql & .WFRESAOICW & "', "     ' ����FLG�iAOi�j
''              sql = sql & "NULL, "                ' ����FLG�iAOi�j
                sql = sql & "NULL, '"               ' �T���v������
                sql = sql & .SMPLPATCW & "', '"     ' �T���v���p�^�[��
''              sql = sql & "NULL, "                ' �T���v������
''              sql = sql & "NULL, '"               ' �T���v���p�^�[��
                sql = sql & .TSTAFFCW & "', "       ' �o�^�Ј�ID
                sql = sql & "sysdate, '"            ' �o�^���t
                sql = sql & .KSTAFFCW & "', "       ' �X�V�Ј�ID
                sql = sql & "sysdate, "             ' �X�V���t
                sql = sql & "'0', "                 ' ���M�t���O
                sql = sql & "sysdate,"              ' ���M���t
'                sql = sql & "'0')"                  ' �����敪
                '' GD�ǉ��@05/01/17 ooba START =====================================>
                sql = sql & "'0', '"                ' �����敪
                sql = sql & .WFSMPLIDGDCW & "', '"  ' �T���v��ID (GD)
                sql = sql & .WFINDGDCW & "', '"     ' ���FLG (GD)
                sql = sql & .WFRESGDCW & "', '"     ' ����FLG (GD)
'                sql = sql & .WFHSGDCW & "')"        ' �ۏ�FLG (GD)  2006/08/15 Del �G�s��s�]���ǉ��Ή� SMP)kondoh
                sql = sql & .WFHSGDCW & "', '"      ' �ۏ�FLG (GD)  2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
                '' GD�ǉ��@05/01/17 ooba END =======================================>
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                sql = sql & .EPSMPLIDB1CW & "', '"  ' �T���v��ID (BMD1E)
                sql = sql & .EPINDB1CW & "', '"     ' ���FLG (BMD1E)
                sql = sql & .EPRESB1CW & "', '"     ' ����FLG (BMD1E)
                sql = sql & .EPSMPLIDB2CW & "', '"  ' �T���v��ID (BMD2E)
                sql = sql & .EPINDB2CW & "', '"     ' ���FLG (BMD2E)
                sql = sql & .EPRESB2CW & "', '"     ' ����FLG (BMD2E)
                sql = sql & .EPSMPLIDB3CW & "', '"  ' �T���v��ID (BMD3E)
                sql = sql & .EPINDB3CW & "', '"     ' ���FLG (BMD3E)
                sql = sql & .EPRESB3CW & "', '"     ' ����FLG (BMD3E)
                sql = sql & .EPSMPLIDL1CW & "', '"  ' �T���v��ID (OSF1E)
                sql = sql & .EPINDL1CW & "', '"     ' ���FLG (OSF1E)
                sql = sql & .EPRESL1CW & "', '"     ' ����FLG (OSF1E)
                sql = sql & .EPSMPLIDL2CW & "', '"  ' �T���v��ID (OSF2E)
                sql = sql & .EPINDL2CW & "', '"     ' ���FLG (OSF2E)
                sql = sql & .EPRESL2CW & "', '"     ' ����FLG (OSF2E)
                sql = sql & .EPSMPLIDL3CW & "', '"  ' �T���v��ID (OSF3E)
                sql = sql & .EPINDL3CW & "', '"     ' ���FLG (OSF3E)
                sql = sql & .EPRESL3CW & "')"       ' ����FLG (OSF3E)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
        
                '' WriteDBLog sql, sDBName
        End With
        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_WfSmp_INS = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
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
    DBDRV_WfSmp_INS = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function




'�T�v      :SXL�Ǘ��̑}���^�X�V
'���Ұ��@�@:�ϐ���      ,IO ,�^               ,����
'      �@�@:SXLOld�@�@�@,I  ,typ_TBCME042   �@,SXL�Ǘ��i���j
'      �@�@:SXLNew�@�@�@,I  ,typ_TBCME042   �@,SXL�Ǘ��i�V�j
'      �@�@:�߂�l      ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :�Â����R�[�h���݂čX�V���}�����𔻕ʂ���
'����      :2001/07/12  �쐬 ���{
Public Function DBDRV_SXL_UpdIns(SXLOld() As typ_TBCME042, SXLNew() As typ_TBCME042) As FUNCTION_RETURN

    Dim sql As String
    Dim lFlg As Boolean
    Dim i As Long
    Dim j As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_SXL_UpdIns"

    DBDRV_SXL_UpdIns = FUNCTION_RETURN_SUCCESS

    For i = 1 To UBound(SXLNew)
        With SXLNew(i)
            lFlg = False
            For j = 1 To UBound(SXLOld)
                If SXLOld(j).CRYNUM = .CRYNUM And _
                   SXLOld(j).INGOTPOS = .INGOTPOS Then
                    sql = "update TBCME042 set "
                    sql = sql & "CRYNUM='" & .CRYNUM & "', "            ' �����ԍ�
                    sql = sql & "INGOTPOS=" & .INGOTPOS & ", "          ' �������J�n�ʒu
                    sql = sql & "LENGTH=" & .Length & ", "              ' ����
                    sql = sql & "SXLID='" & .SXLID & "', "              ' SXLID
                    sql = sql & "KRPROCCD='" & .KRPROCCD & "', "        ' �Ǘ��H��
                    sql = sql & "NOWPROC='" & .NOWPROC & "', "          ' ���ݍH��
                    sql = sql & "LPKRPROCCD='" & .LPKRPROCCD & "', "    ' �ŏI�ʉߊǗ��H��
                    sql = sql & "LASTPASS='" & .LASTPASS & "', "        ' �ŏI�ʉߍH��
                    sql = sql & "DELCLS='" & .DELCLS & "', "            ' �폜�敪
                    sql = sql & "LSTATCLS='" & .LSTATCLS & "', "        ' �ŏI��ԋ敪
                    sql = sql & "HOLDCLS='" & .HOLDCLS & "', "          ' �z�[���h�敪
                    sql = sql & "HINBAN='" & .hinban & "', "            ' �i��
                    sql = sql & "REVNUM=" & .REVNUM & ", "              ' ���i�ԍ������ԍ�
                    sql = sql & "FACTORY='" & .factory & "', "          ' �H��
                    sql = sql & "OPECOND='" & .opecond & "', "          ' ���Ə���
                    sql = sql & "BDCAUS='" & .BDCAUS & "', "            ' �s�Ǘ��R
                    sql = sql & "COUNT=" & .COUNT & ", "                ' ����
                    sql = sql & "UPDDATE=sysdate, "                     ' �X�V���t
                    sql = sql & "SUMMITSENDFLAG='0', "                  ' SUMMIT���M�t���O
                    sql = sql & "SENDFLAG='0'"                          ' ���M�t���O
                    sql = sql & " where CRYNUM='" & .CRYNUM & "'"
                    sql = sql & " and INGOTPOS=" & .INGOTPOS
                    '' WriteDBLog sql
                    If OraDB.ExecuteSQL(sql) <= 0 Then
                        GoTo proc_err
                    End If
                    lFlg = True
                    Exit For
                End If
            Next j

            If lFlg <> True Then
                sql = "insert into TBCME042 ("
                sql = sql & "CRYNUM, "              ' �����ԍ�
                sql = sql & "INGOTPOS, "            ' �������J�n�ʒu
                sql = sql & "LENGTH, "              ' ����
                sql = sql & "SXLID, "               ' SXLID
                sql = sql & "KRPROCCD, "            ' �Ǘ��H��
                sql = sql & "NOWPROC, "             ' ���ݍH��
                sql = sql & "LPKRPROCCD, "          ' �ŏI�ʉߊǗ��H��
                sql = sql & "LASTPASS, "            ' �ŏI�ʉߍH��
                sql = sql & "DELCLS, "              ' �폜�敪
                sql = sql & "LSTATCLS, "            ' �ŏI��ԋ敪
                sql = sql & "HOLDCLS, "             ' �z�[���h�敪
                sql = sql & "HINBAN, "              ' �i��
                sql = sql & "REVNUM, "              ' ���i�ԍ������ԍ�
                sql = sql & "FACTORY, "             ' �H��
                sql = sql & "OPECOND, "             ' ���Ə���
                sql = sql & "BDCAUS, "              ' �s�Ǘ��R
                sql = sql & "COUNT, "               ' ����
                sql = sql & "REGDATE, "             ' �o�^���t
                sql = sql & "UPDDATE, "             ' �X�V���t
                sql = sql & "SUMMITSENDFLAG, "      ' SUMMIT���M�t���O
                sql = sql & "SENDFLAG, "            ' ���M�t���O
                sql = sql & "SENDDATE)"             ' ���M���t
                sql = sql & " values ('"
                sql = sql & .CRYNUM & "', "         ' �����ԍ�
                sql = sql & .INGOTPOS & ", "        ' �������J�n�ʒu
                sql = sql & .Length & ", '"         ' ����
                sql = sql & .SXLID & "', '"         ' SXLID
                sql = sql & .KRPROCCD & "', '"      ' �Ǘ��H��
                sql = sql & .NOWPROC & "', '"       ' ���ݍH��
                sql = sql & .LPKRPROCCD & "', '"    ' �ŏI�ʉߊǗ��H��
                sql = sql & .LASTPASS & "', '"      ' �ŏI�ʉߍH��
                sql = sql & .DELCLS & "', '"        ' �폜�敪
                sql = sql & .LSTATCLS & "', '"      ' �ŏI��ԋ敪
                sql = sql & .HOLDCLS & "', '"       ' �z�[���h�敪
                sql = sql & .hinban & "', "         ' �i��
                sql = sql & .REVNUM & ", '"         ' ���i�ԍ������ԍ�
                sql = sql & .factory & "', '"       ' �H��
                sql = sql & .opecond & "', '"       ' ���Ə���
                sql = sql & .BDCAUS & "', "         ' �s�Ǘ��R
                sql = sql & .COUNT & ", "           ' ����
                sql = sql & "sysdate, "             ' �o�^���t
                sql = sql & "sysdate, "             ' �X�V���t
                sql = sql & "'0', "                 ' SUMMIT���M�t���O
                sql = sql & "'0', "                 ' ���M�t���O
                sql = sql & "sysdate)"              ' ���M���t
                '' WriteDBLog sql
                If OraDB.ExecuteSQL(sql) <= 0 Then
                    DBDRV_SXL_UpdIns = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            End If
        End With
    Next i

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_SXL_UpdIns = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'�T�v      :SXL�Ǘ��̑}��
'���Ұ��@�@:�ϐ���      ,IO ,�^               ,����
'      �@�@:SXL   �@�@�@,I  ,typ_TBCME042   �@,SXL�Ǘ�
'      �@�@:�߂�l      ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :DBDRV_SXL_UpdIns�Ɉڍs����\��
'����      :2001/07/12  �쐬 ���{
Public Function DBDRV_SXL_INS(SXL() As typ_TBCME042) As FUNCTION_RETURN

    Dim sql As String
    Dim i As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_SXL_INS"

    DBDRV_SXL_INS = FUNCTION_RETURN_SUCCESS

    For i = 1 To UBound(SXL)
        If SXL(i).Length > 0 Then
            sql = "delete from TBCME042 where CRYNUM='" & SXL(i).CRYNUM & "' and INGOTPOS=" & SXL(i).INGOTPOS
            OraDB.ExecuteSQL sql
            With SXL(i)
                sql = "insert into TBCME042 ("
                sql = sql & "CRYNUM, "              ' �����ԍ�
                sql = sql & "INGOTPOS, "            ' �������J�n�ʒu
                sql = sql & "LENGTH, "              ' ����
                sql = sql & "SXLID, "               ' SXLID
                sql = sql & "KRPROCCD, "            ' �Ǘ��H��
                sql = sql & "NOWPROC, "             ' ���ݍH��
                sql = sql & "LPKRPROCCD, "          ' �ŏI�ʉߊǗ��H��
                sql = sql & "LASTPASS, "            ' �ŏI�ʉߍH��
                sql = sql & "DELCLS, "              ' �폜�敪
                sql = sql & "LSTATCLS, "            ' �ŏI��ԋ敪
                sql = sql & "HOLDCLS, "             ' �z�[���h�敪
                sql = sql & "HINBAN, "              ' �i��
                sql = sql & "REVNUM, "              ' ���i�ԍ������ԍ�
                sql = sql & "FACTORY, "             ' �H��
                sql = sql & "OPECOND, "             ' ���Ə���
                sql = sql & "BDCAUS, "              ' �s�Ǘ��R
                sql = sql & "COUNT, "               ' ����
                sql = sql & "REGDATE, "             ' �o�^���t
                sql = sql & "UPDDATE, "             ' �X�V���t
                sql = sql & "SUMMITSENDFLAG, "      ' SUMMIT���M�t���O
                sql = sql & "SENDFLAG, "            ' ���M�t���O
                sql = sql & "SENDDATE)"             ' ���M���t
                sql = sql & " values ('"
                sql = sql & .CRYNUM & "', "         ' �����ԍ�
                sql = sql & .INGOTPOS & ", "        ' �������J�n�ʒu
                sql = sql & .Length & ", '"         ' ����
                sql = sql & .SXLID & "', '"         ' SXLID
                sql = sql & .KRPROCCD & "', '"      ' �Ǘ��H��
                sql = sql & .NOWPROC & "', '"       ' ���ݍH��
                sql = sql & .LPKRPROCCD & "', '"    ' �ŏI�ʉߊǗ��H��
                sql = sql & .LASTPASS & "', '"      ' �ŏI�ʉߍH��
                sql = sql & .DELCLS & "', '"        ' �폜�敪
                sql = sql & .LSTATCLS & "', '"      ' �ŏI��ԋ敪
                sql = sql & .HOLDCLS & "', '"       ' �z�[���h�敪
                sql = sql & .hinban & "', "         ' �i��
                sql = sql & .REVNUM & ", '"         ' ���i�ԍ������ԍ�
                sql = sql & .factory & "', '"       ' �H��
                sql = sql & .opecond & "', '"       ' ���Ə���
                sql = sql & .BDCAUS & "', "         ' �s�Ǘ��R
                sql = sql & .COUNT & ", "           ' ����
                sql = sql & "sysdate, "             ' �o�^���t
                sql = sql & "sysdate, "             ' �X�V���t
                sql = sql & "'0', "                 ' SUMMIT���M�t���O
                sql = sql & "'0', "                 ' ���M�t���O
                sql = sql & "sysdate)"              ' ���M���t
            End With
            '' WriteDBLog sql
            If OraDB.ExecuteSQL(sql) <= 0 Then
                DBDRV_SXL_INS = FUNCTION_RETURN_FAILURE
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
    DBDRV_SXL_INS = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'�T�v      :����]�����@�w���̑}��
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:SokuSizi�@�@�@,I  ,typ_TBCMY003   �@,����]�����@�w��
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :
'����      :2001/07/12  �쐬 ���{
Public Function DBDRV_SokuSizi_Ins(SokuSizi() As typ_TBCMY003) As FUNCTION_RETURN

    Dim sql As String
    Dim i As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_SokuSizi_Ins"

    '' ����]�����@�w���̑}��
    For i = 1 To UBound(SokuSizi)
        With SokuSizi(i)
            sql = " insert into TBCMY003 ("
            sql = sql & "SAMPLEID, "                ' �T���v��ID
            sql = sql & "OSITEM, "                  ' �]������
            sql = sql & "TRANCNT, "                 ' ������
            sql = sql & "SAMPLEKB, "                ' �T���v���敪
            sql = sql & "MAISU, "                   ' �]������
            sql = sql & "SPEC, "                    ' �K�i�l
            sql = sql & "NETSU, "                   ' �M��������
            sql = sql & "ET, "                      ' �G�b�`���O����
            sql = sql & "MES, "                     ' �v�����@
            sql = sql & "DKAN, "                    ' �c�j�A�j�[������
            sql = sql & "TXID, "                    ' �g�����U�N�V����ID
            sql = sql & "REGDATE, "                 ' �o�^���t
            sql = sql & "SENDFLAG, "                ' ���M�t���O
            sql = sql & "SENDDATE, "                ' ���M���t
            sql = sql & "PLANTCAT, "                ' ���� 2007/08/31 SPK Tsutsumi Add
            '06/06/08 ooba START =======================================================>
            sql = sql & "FEPUA, "                   ' SPV_Fe_PUA�l
            sql = sql & "FEPUAPCT, "                ' SPV_Fe_PUA���l
            sql = sql & "FESTD, "                   ' SPV_Fe_STD
            sql = sql & "DIFFPUA, "                 ' SPV_�g�U��_PUA�l
            sql = sql & "DIFFPUAPCT, "              ' SPV_�g�U��_PUA���l
            sql = sql & "NRPUA, "                   ' SPV_NR_PUA�l
            sql = sql & "NRPUAPCT, "                ' SPV_NR_PUA%�l
            sql = sql & "NRSTD) "                   ' SPV_NR_STD
            '06/06/08 ooba END =========================================================>
            sql = sql & " select '"
            sql = sql & .SAMPLEID & "', '"          ' �T���v��ID
            sql = sql & .OSITEM & "', "             ' �]������
            sql = sql & "nvl(max(TRANCNT),0)+1, '"  ' ������
            sql = sql & .SAMPLEKB & "', '"           ' �T���v���敪
            'sql = sql & "'1', '"                    ' �]������   2004/06/23
            sql = sql & .MAISU & "', '"        ' �]������
            sql = sql & .Spec & "', '"              ' �K�i�l
            sql = sql & .NETSU & "', '"             ' �M��������
            sql = sql & .ET & "', '"                ' �G�b�`���O����
            sql = sql & .MES & "', '"               ' �v�����@
            sql = sql & .DKAN & "', "               ' �c�j�A�j�[������
            sql = sql & "'TX851I', "                ' �g�����U�N�V����ID
            sql = sql & "sysdate, "                 ' �o�^���t
            sql = sql & "'" & .SENDFLAG & "', "     ' ���M�t���O
            sql = sql & "sysdate, "                 ' ���M���t
            sql = sql & "'" & .MUKESAKI & "', "     ' ���� 2007/08/31 SPK Tsutsumi Add
            '06/06/08 ooba START =======================================================>
            If IsNumeric(.FEPUA) Then sql = sql & CDbl(.FEPUA) & ", " Else sql = sql & "NULL, "
            If IsNumeric(.FEPUAPCT) Then sql = sql & CDbl(.FEPUAPCT) & ", " Else sql = sql & "NULL, "
            If IsNumeric(.FESTD) Then sql = sql & CDbl(.FESTD) & ", " Else sql = sql & "NULL, "
            If IsNumeric(.DIFFPUA) Then sql = sql & CDbl(.DIFFPUA) & ", " Else sql = sql & "NULL, "
            If IsNumeric(.DIFFPUAPCT) Then sql = sql & CDbl(.DIFFPUAPCT) & ", " Else sql = sql & "NULL, "
            If IsNumeric(.NRPUA) Then sql = sql & CDbl(.NRPUA) & ", " Else sql = sql & "NULL, "
            If IsNumeric(.NRPUAPCT) Then sql = sql & CDbl(.NRPUAPCT) & ", " Else sql = sql & "NULL, "
            If IsNumeric(.NRSTD) Then sql = sql & CDbl(.NRSTD) Else sql = sql & "NULL "
            '06/06/08 ooba END =========================================================>
            sql = sql & " from TBCMY003"
            sql = sql & " where SAMPLEID='" & .SAMPLEID & "'"
            sql = sql & " and OSITEM='" & .OSITEM & "'"
            sql = sql & " and SPEC='" & .Spec & "'"
        End With
        '' WriteDBLog sql
        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_SokuSizi_Ins = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    DBDRV_SokuSizi_Ins = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_SokuSizi_Ins = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'�T�v      :�]�p���т̑}��
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:Tenyou�@ �@�@,I  ,typ_TBCMJ013   �@,�]�p����
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :
'����      :2002/01/15  �쐬 S.Sano
Public Function DBDRV_Tenyou_Ins(Tenyou As typ_TBCMJ013) As FUNCTION_RETURN

    Dim sql As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_Tenyou_Ins"

    '' �U�֔p�����т̑}��
    With Tenyou
        sql = "insert into TBCMJ013 ("
        sql = sql & "CRYNUM, "                  ' �����ԍ�
        sql = sql & "INGOTPOS, "                ' �C���S�b�g�ʒu
        sql = sql & "TRANCNT, "                 ' ������
        sql = sql & "LENGTH, "                  ' ����
        sql = sql & "KRPROCCD, "                ' �Ǘ��H���R�[�h
        sql = sql & "PROCCODE, "                ' �H���R�[�h
        sql = sql & "DUNWNUM, "                 ' �]�p��i��
        sql = sql & "DUNWREV, "                 ' �]�p��i�� ���i�ԍ������ԍ�
        sql = sql & "DUNWFACT, "                ' �]�p��i�� �H��
        sql = sql & "DUNWOPCD, "                ' �]�p��i�� ���Ə���
        sql = sql & "DUOGNUM, "                 ' �]�p���i��
        sql = sql & "DUOGREV, "                 ' �]�p���i�� ���i�ԍ������ԍ�
        sql = sql & "DUOGFACT, "                ' �]�p���i�� �H��
        sql = sql & "DUOGOPCD, "                ' �]�p���i�� ���Ə���
        sql = sql & "TSTAFFID, "                ' �o�^�Ј�ID
        sql = sql & "REGDATE, "                 ' �o�^���t
        sql = sql & "KSTAFFID, "                ' �X�V�Ј�ID
        sql = sql & "UPDDATE, "                 ' �X�V���t
        sql = sql & "SENDFLAG)"                 ' ���M�t���O
        sql = sql & " select '"
        sql = sql & .CRYNUM & "', "             ' �����ԍ�
        sql = sql & .INGOTPOS & ", "            ' �C���S�b�g�ʒu
        sql = sql & "nvl(max(TRANCNT),0)+1, "   ' ������
        sql = sql & .Length & ", '"             ' ����
        sql = sql & .KRPROCCD & "', '"          ' �Ǘ��H���R�[�h
        sql = sql & .PROCCODE & "', '"          ' �H���R�[�h
        sql = sql & .DUNWNUM & "', "            ' �]�p��i��
        sql = sql & .DUNWREV & ", '"            ' �]�p��i�� ���i�ԍ������ԍ�
        sql = sql & .DUNWFACT & "', '"          ' �]�p��i�� �H��
        sql = sql & .DUNWOPCD & "', '"          ' �]�p��i�� ���Ə���
        sql = sql & .DUOGNUM & "', "            ' �]�p���i��
        sql = sql & .DUOGREV & ", '"            ' �]�p���i�� ���i�ԍ������ԍ�
        sql = sql & .DUOGFACT & "', '"          ' �]�p���i�� �H��
        sql = sql & .DUOGOPCD & "', '"          ' �]�p���i�� ���Ə���
        sql = sql & .TSTAFFID & "', "           ' �o�^�Ј�ID
        sql = sql & "sysdate, '"                ' �o�^���t
        sql = sql & .KSTAFFID & "', "           ' �X�V�Ј�ID
        sql = sql & "sysdate, "                 ' �X�V���t
        sql = sql & "'0'"                       ' ���M�t���O
        sql = sql & " from TBCMJ013"
        sql = sql & " where CRYNUM='" & .CRYNUM & "' and INGOTPOS=" & .INGOTPOS
    End With
    '' WriteDBLog sql
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_Tenyou_Ins = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_Tenyou_Ins = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_Tenyou_Ins = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'�T�v      :�U�֔p�����т̑}��
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:Hurikae�@�@�@,I  ,typ_TBCMW006   �@,�U�֔p������
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :
'����      :2001/07/12  �쐬 ���{
Public Function DBDRV_Furikae_Ins(Hurikae As typ_TBCMW006) As FUNCTION_RETURN

    Dim sql As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_Furikae_Ins"

    '' �U�֔p�����т̑}��
    With Hurikae
        sql = "insert into TBCMW006 ("
        sql = sql & "CRYNUM, "                  ' �����ԍ�
        sql = sql & "INGOTPOS, "                ' �C���S�b�g�ʒu
        sql = sql & "TRANCNT, "                 ' ������
        sql = sql & "CRYLEN, "                  ' ����
        sql = sql & "KRPROCCD, "                ' �Ǘ��H���R�[�h
        sql = sql & "PROCCODE, "                ' �H���R�[�h
        sql = sql & "TRANCLS, "                 ' �����敪
        sql = sql & "DUNWNUM, "                 ' �]�p��i��
        sql = sql & "DUNWREV, "                 ' �]�p��i�� ���i�ԍ������ԍ�
        sql = sql & "DUNWFACT, "                ' �]�p��i�� �H��
        sql = sql & "DUNWOPCD, "                ' �]�p��i�� ���Ə���
        sql = sql & "DUOGNUM, "                 ' �]�p���i��
        sql = sql & "DUOGREV, "                 ' �]�p���i�� ���i�ԍ������ԍ�
        sql = sql & "DUOGFACT, "                ' �]�p���i�� �H��
        sql = sql & "DUOGOPCD, "                ' �]�p���i�� ���Ə���
        sql = sql & "TSTAFFID, "                ' �o�^�Ј�ID
        sql = sql & "REGDATE, "                 ' �o�^���t
        sql = sql & "KSTAFFID, "                ' �X�V�Ј�ID
        sql = sql & "UPDDATE, "                 ' �X�V���t
        ' 2007/09/03 SPK Tsutsumi Add Start
        sql = sql & "SENDFLAG, "                ' ���M�t���O
'        sql = sql & "SENDFLAG) "                ' ���M�t���O
        sql = sql & "PLANTCAT) "                ' ����
        ' 2007/09/03 SPK Tsutsumi Add End
        sql = sql & " select '"
        sql = sql & .CRYNUM & "', "             ' �����ԍ�
        sql = sql & .INGOTPOS & ", "            ' �C���S�b�g�ʒu
        sql = sql & "nvl(max(TRANCNT),0)+1, "   ' ������
        sql = sql & .CRYLEN & ", '"             ' ����
        sql = sql & .KRPROCCD & "', '"          ' �Ǘ��H���R�[�h
        sql = sql & .PROCCODE & "', '"          ' �H���R�[�h
        sql = sql & .TRANCLS & "', '"           ' �����敪
        sql = sql & .DUNWNUM & "', "            ' �]�p��i��
        sql = sql & .DUNWREV & ", '"            ' �]�p��i�� ���i�ԍ������ԍ�
        sql = sql & .DUNWFACT & "', '"          ' �]�p��i�� �H��
        sql = sql & .DUNWOPCD & "', '"          ' �]�p��i�� ���Ə���
        sql = sql & .DUOGNUM & "', "            ' �]�p���i��
        sql = sql & .DUOGREV & ", '"            ' �]�p���i�� ���i�ԍ������ԍ�
        sql = sql & .DUOGFACT & "', '"          ' �]�p���i�� �H��
        sql = sql & .DUOGOPCD & "', '"          ' �]�p���i�� ���Ə���
        sql = sql & .TSTAFFID & "', "           ' �o�^�Ј�ID
        sql = sql & "sysdate, '"                ' �o�^���t
        sql = sql & .KSTAFFID & "', "           ' �X�V�Ј�ID
        sql = sql & "sysdate, "                 ' �X�V���t
        ' 2007/09/03 SPK Tsutsumi Add Start
        sql = sql & "'0',"                      ' ���M�t���O
        sql = sql & "'" & .MUKESAKI & "'"    ' ����
'        sql = sql & "'0' "                      ' ���M�t���O
        ' 2007/09/03 SPK Tsutsumi Add End
        sql = sql & " from TBCMW006 "
        sql = sql & " where CRYNUM='" & .CRYNUM & "' and INGOTPOS=" & .INGOTPOS
    End With
    '' WriteDBLog sql
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_Furikae_Ins = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_Furikae_Ins = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_Furikae_Ins = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'�N���X�^���J�^���O������тւ̑}��
Public Function DBDRV_Catalog_Ins(CryCatalog As typ_TBCMG007) As FUNCTION_RETURN

    Dim sql As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_Catalog_Ins"

    DBDRV_Catalog_Ins = FUNCTION_RETURN_SUCCESS

    ' �N���X�^���J�^���O������тւ̑}��
    sql = "insert into TBCMG007 ( "
    sql = sql & "CRYNUM, "            ' �����ԍ��i�u���b�NID�j
    sql = sql & "TRANCNT, "           ' ������
    sql = sql & "KRPROCCD, "          ' �Ǘ��H���R�[�h
    sql = sql & "PROCCODE, "          ' �H���R�[�h
    sql = sql & "BDCODE, "            ' �s�Ǘ��R�R�[�h
    sql = sql & "PALTNUM, "           ' �p���b�g�ԍ�
    sql = sql & "TSTAFFID, "          ' �o�^�Ј�ID
    sql = sql & "REGDATE, "           ' �o�^���t
    sql = sql & "KSTAFFID, "          ' �X�V�Ј�ID
    sql = sql & "UPDDATE, "           ' �X�V���t
    sql = sql & "SENDFLAG, "          ' ���M�t���O
    sql = sql & "SENDDATE) "          ' ���M���t

    With CryCatalog
        sql = sql & "Select "
        sql = sql & " '" & .CRYNUM & "', "          ' �����ԍ�
        sql = sql & "nvl(max(TRANCNT),0)+1, "       ' ������
        sql = sql & " '" & .KRPROCCD & "', "        ' �Ǘ��H���R�[�h
        sql = sql & " '" & .PROCCODE & "', "        ' �H���R�[�h
        sql = sql & " '" & .BDCODE & "', "          ' �s�Ǘ��R�R�[�h
        sql = sql & " '" & .PALTNUM & "', "         ' �p���b�g�ԍ�
        sql = sql & " '" & .TSTAFFID & "', "        ' �o�^�Ј�ID
        sql = sql & "sysdate, "                     ' �o�^���t
        sql = sql & " '" & .TSTAFFID & "', "        ' �X�V�Ј�ID
        sql = sql & "sysdate, "                     ' �X�V���t
        sql = sql & "'0', "                         ' ���M�t���O
        sql = sql & "sysdate "                      ' ���M���t
        sql = sql & "From TBCMG007 " & _
              "Where (CRYNUM='" & .CRYNUM & "')"
    End With

    '' WriteDBLog sql
    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_Catalog_Ins = FUNCTION_RETURN_FAILURE
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_Catalog_Ins = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�����ԍ�����AGR or MGR���擾
'���Ұ��@�@:�ϐ���      ,IO ,�^               ,����
'      �@�@:CRYNUM�@�@  ,I  ,String         �@,�����ԍ�
'      �@�@:ans  �@�@�@ ,I  ,String         �@,A or M or ""
'      �@�@:�߂�l      ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
'����      :
'����      :2001/07/12  �쐬 ���{
Public Function DBDRV_get_xGR(CRYNUM As String, Ans As String) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_get_xGR"
    DBDRV_get_xGR = FUNCTION_RETURN_FAILURE

    sql = "select PRCMCN from TBCMI001 "
    sql = sql & "where CRYNUM = '" & left(CRYNUM, 9) & "000" & "' and "
    sql = sql & "TRANCNT = any(select max(TRANCNT) from TBCMI001 where CRYNUM = '" & left(CRYNUM, 9) & "000" & "')"

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        DBDRV_get_xGR = FUNCTION_RETURN_SUCCESS
        Ans = ""
        rs.Close
        GoTo proc_exit
    End If
    Ans = rs("PRCMCN")
    rs.Close
    
    
    DBDRV_get_xGR = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_get_xGR = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�w��̃u���b�N�����݂��邩�ǂ����̃`�F�b�N����
'���Ұ��@�@:�ϐ���      ,IO ,�^               ,����
'      �@�@:BLOCKID�@�@ ,I  ,String         �@,�����ԍ�
'      �@�@:ans  �@�@�@ ,I  ,Boolean        �@,�L��(True)����(False)
'      �@�@:�߂�l      ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
'����      :
'����      :2001/07/12  �쐬 ���{
Public Function DBDRV_BlockIDCheck(BLOCKID As String, Ans As Boolean) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_BlockIDCheck"
    DBDRV_BlockIDCheck = FUNCTION_RETURN_FAILURE

    sql = "select BLOCKID from TBCME040 "
    sql = sql & "where CRYNUM = '" & left(BLOCKID, 9) & "000" & "' and "
    sql = sql & "BLOCKID = '" & BLOCKID & "'"

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Ans = False
    Else
        Ans = True
    End If
    rs.Close
    DBDRV_BlockIDCheck = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_BlockIDCheck = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function
        
'�T�v      :�w���P�����̃V�[�h�X�������߂�
'���Ұ��@�@:�ϐ���      ,IO ,�^               ,����
'      �@�@:BLOCKID�@�@ ,I  ,String         �@,�����ԍ�
'      �@�@:ans  �@�@�@ ,I  ,Boolean        �@,�L��(True)����(False)
'      �@�@:�߂�l      ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
'����      :
'����      :2001/07/12  �쐬 ���{
Public Function DBDRV_getSEEDDEG(BLOCKID As String, Ans As Integer) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_getSEEDDEG"
    DBDRV_getSEEDDEG = FUNCTION_RETURN_FAILURE
        
    '�u���b�N�V�K���A�V�[�h�X���̋��ߕ���ύX
    sql = "select SEEDDEG from TBCMG002 "
    sql = sql & "where TRANCNT=ANY(select MAX(TRANCNT) from TBCMG002 Where CRYNUM='" & BLOCKID & "') and "
    sql = sql & "CRYNUM='" & BLOCKID & "'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <= 0 Then
        rs.Close
        GoTo proc_exit
    End If
    If rs("SEEDDEG") = 4 Then
        Ans = 4
    Else
        Ans = 0
    End If
    rs.Close
    
    DBDRV_getSEEDDEG = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_getSEEDDEG = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�u���b�N���ōł�LT�d�l���������i�Ԃ����߂�
'���Ұ�    :�ϐ���        ,IO ,�^          ,����
'          :Crynum        ,I  ,String      ,�����ԍ�
'          :Ingotpos      ,I  ,Integer     ,�u���b�N�̏I���ʒu
'          :hin           ,O  ,tFullHinban ,�u���b�N���ōł�LT�d�l�̌������i��
'          :LTSPI         ,O  ,String      ,�u���b�N���ōł�LT�d�l�̌������i�Ԃ�LT����ʒu�R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN,�Ǎ��̐���
'����      :�u�ł�LT�d�l���������i�ԁv���Ȃ���΁Ahin.HINBAN='        ', LTSPI=VbNullString
'����      :2002/4/23 �쑺 �쐬
Public Function DBDRV_getLtHinbanInBlock(CRYNUM As String, INGOTPOS As Integer, HIN As tFullHinban, LTSPI As String) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset
Dim recCnt As Integer
Dim BlkFrom As Integer
Dim BlkTo As Integer

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_getLtHinbanInBlock"
    DBDRV_getLtHinbanInBlock = FUNCTION_RETURN_FAILURE
    
    '������
    HIN.hinban = vbNullString
    HIN.mnorevno = 0
    HIN.factory = vbNullString
    HIN.factory = vbNullString
    LTSPI = vbNullString
    
    '�u���b�N�͈̔͂����߂�
    BlkTo = INGOTPOS
    sql = "select INGOTPOS from TBCME040 where CRYNUM='" & CRYNUM & "' and INGOTPOS+LENGTH=" & BlkTo
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <= 0 Then     '�u���b�N�I�[�łȂ����߁A�Ώەi�ԂȂ�
        rs.Close
        DBDRV_getLtHinbanInBlock = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    BlkFrom = rs("INGOTPOS")
    rs.Close
    Set rs = Nothing
    
    '�u���b�N���̕i�ԂŁA�ł�LT�d�l�����������̂����߂�
    sql = "select SIYO.HINBAN, SIYO.MNOREVNO, SIYO.FACTORY, SIYO.OPECOND, SIYO.HSXLTHWS, SIYO.HSXLTSPI "
    sql = sql & "from TBCME041 HIN, TBCME019 SIYO "
    sql = sql & "where HIN.CRYNUM='" & CRYNUM & "'"
    sql = sql & "  and HIN.INGOTPOS<" & BlkTo & " and HIN.INGOTPOS+HIN.LENGTH>" & BlkFrom
    sql = sql & "  and SIYO.HINBAN=HIN.HINBAN and SIYO.MNOREVNO=HIN.REVNUM and SIYO.FACTORY=HIN.FACTORY and SIYO.OPECOND=HIN.OPECOND"
    sql = sql & "  and SIYO.HSXLTHWS in ('H','S') "
    sql = sql & "order by SIYO.HSXLTSPI, HIN.INGOTPOS desc"
    sql = "select * from (" & sql & ") where rownum=1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <= 0 Then     '�Ώەi�ԂȂ�(LT���ۏ�/�Q�l�̕i�Ԃ��Ȃ��Ǝv����)
        rs.Close
        DBDRV_getLtHinbanInBlock = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    HIN.hinban = rs("HINBAN")
    HIN.mnorevno = rs("MNOREVNO")
    HIN.factory = rs("FACTORY")
    HIN.opecond = rs("OPECOND")
    LTSPI = rs("HSXLTSPI")
    rs.Close
    Set rs = Nothing
    
    DBDRV_getLtHinbanInBlock = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_getLtHinbanInBlock = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :����وʒu�̏�i�ԂƉ��i�Ԃ��قȂ�ꍇ�A�ł�GD�d�l���������i�Ԃ����߂�
'���Ұ�    :�ϐ���      ,IO ,�^          ,����
'          :tblCrySmp   ,I  ,typ_XSDCS   ,�����
'          :HIN         ,I  ,tFullHinban ,�ł�GD�d�l���������i��
'          :�߂�l        ,O  ,FUNCTION_RETURN,�Ǎ��̐���
'����      :�u�ł�GD�d�l���������i�ԁv���Ȃ���΁Ahin.HINBAN='        '
'����      :2005/10/05 Y.SIMIZU
Public Function DBDRV_getGDHinbanInBlock(tblCrySmp As typ_XSDCS, HIN As tFullHinban) As FUNCTION_RETURN
    Dim sql     As String
    Dim rs      As OraDynaset
    Dim sGDLine As Single

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_getGDHinbanInBlock"
    DBDRV_getGDHinbanInBlock = FUNCTION_RETURN_FAILURE
    
    '������
    HIN.hinban = vbNullString
    HIN.mnorevno = 0
    HIN.factory = vbNullString
    HIN.factory = vbNullString
    
    '���ʒu�̕i�Ԃ�GDײݐ����擾����
    sql = "SELECT  DISTINCT HINBCS,REVNUMCS,FACTORYCS,OPECS,HSXGDLINE,HWFGDLINE "
    sql = sql & "FROM   XSDCS T1,TBCME036 T2 "
    sql = sql & "WHERE  T1.XTALCS = '" & tblCrySmp.XTALCS & "' "
    sql = sql & "AND    T1.INPOSCS = " & tblCrySmp.INPOSCS & " "
    sql = sql & "AND    T1.LIVKCS <> '1' "
    sql = sql & "AND    T1.HINBCS = T2.HINBAN "
    sql = sql & "AND    T1. REVNUMCS = T2.MNOREVNO "
    sql = sql & "AND    T1. FACTORYCS = T2.FACTORY "
    sql = sql & "AND    T1. OPECS = T2.OPECOND "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount <= 0 Then     '�Ώەi�ԂȂ�
        rs.Close
        DBDRV_getGDHinbanInBlock = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    
    '���i�Ԃ�GDײݐ����擾
    Do Until rs.EOF
        With tblCrySmp
            '���i�Ԃ̏ꍇ
            If rs("HINBCS") = .HINBCS And rs("REVNUMCS") = .REVNUMCS And rs("FACTORYCS") = .FACTORYCS And rs("OPECS") = .OPECS Then
                HIN.hinban = rs("HINBCS")
                HIN.mnorevno = rs("REVNUMCS")
                HIN.factory = rs("FACTORYCS")
                HIN.opecond = rs("OPECS")
                sGDLine = fncNullCheck(rs("HSXGDLINE"))
            End If
        End With
        rs.MoveNext
    Loop
    
    rs.MoveFirst
    
    '���i�Ԃ�ײݐ�����ײݐ��������i�Ԃ��擾
    Do Until rs.EOF
        '���i�Ԃ�ײݐ�����ײݐ��������ꍇ
        If fncNullCheck(rs("HSXGDLINE")) > sGDLine Then
            HIN.hinban = rs("HINBCS")
            HIN.mnorevno = rs("REVNUMCS")
            HIN.factory = rs("FACTORYCS")
            HIN.opecond = rs("OPECS")
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    
    Set rs = Nothing
    
    DBDRV_getGDHinbanInBlock = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_getGDHinbanInBlock = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�G�s����]�����@�w���̑}��
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:SokuSizi�@�@�@,I  ,typ_TBCMY020   �@,�G�s����]�����@�w��
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :
'����      :2006/08/15  �쐬 SMP)kondoh
Public Function DBDRV_SokuSizi_EP_Ins(SokuSizi() As typ_TBCMY020) As FUNCTION_RETURN

    Dim sql As String
    Dim i As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_SokuSizi_EP_Ins"

    '' ����]�����@�w���̑}��
    For i = 1 To UBound(SokuSizi)
        With SokuSizi(i)
            sql = " insert into TBCMY020 ("
            sql = sql & "SAMPLEID, "                ' �T���v��ID
            sql = sql & "OSITEM, "                  ' �]������
            sql = sql & "TRANCNT, "                 ' ������
            sql = sql & "SAMPLEKB, "                ' �T���v���敪
            sql = sql & "MAISU, "                   ' �]������
            sql = sql & "SPEC, "                    ' �K�i�l
            sql = sql & "NETSU, "                   ' �M��������
            sql = sql & "ET, "                      ' �G�b�`���O����
            sql = sql & "MES, "                     ' �v�����@
            sql = sql & "DKAN, "                    ' �c�j�A�j�[������
            sql = sql & "TXID, "                    ' �g�����U�N�V����ID
            sql = sql & "REGDATE, "                 ' �o�^���t
            sql = sql & "SENDFLAG, "                ' ���M�t���O
            
            ' 2007/08/31 SPK Tsutsumi Add Start
            sql = sql & "SENDDATE, "                ' ���M���t
            sql = sql & "PLANTCAT) "                ' ���� 2007/08/31 SPK Tsutsumi Add
            'sql = sql & "SENDDATE) "                ' ���M���t
            ' 2007/08/31 SPK Tsutsumi Add End

            sql = sql & " select '"
            sql = sql & .SAMPLEID & "', '"          ' �T���v��ID
            sql = sql & .OSITEM & "', "             ' �]������
            sql = sql & "nvl(max(TRANCNT),0)+1, '"  ' ������
            sql = sql & .SAMPLEKB & "', '"           ' �T���v���敪
            sql = sql & .MAISU & "', '"             ' �]������
            sql = sql & .Spec & "', '"              ' �K�i�l
            sql = sql & .NETSU & "', '"             ' �M��������
            sql = sql & .ET & "', '"                 ' �G�b�`���O����
            sql = sql & .MES & "', '"               ' �v�����@
            sql = sql & .DKAN & "', "               ' �c�j�A�j�[������
            sql = sql & "'TX871I', "                ' �g�����U�N�V����ID
            sql = sql & "sysdate, "                 ' �o�^���t
            sql = sql & "'" & .SENDFLAG & "', "     ' ���M�t���O
            
            ' 2007/08/31 SPK Tsutsumi Add Start
'            sql = sql & "sysdate "                 ' ���M���t
            sql = sql & "sysdate, "                 ' ���M���t
            sql = sql & "'" & .MUKESAKI & "' "   ' ����
            ' 2007/08/31 SPK Tsutsumi Add End
            
            sql = sql & " from TBCMY020"
            sql = sql & " where SAMPLEID='" & .SAMPLEID & "'"
            sql = sql & " and OSITEM='" & .OSITEM & "'"
            sql = sql & " and SPEC='" & .Spec & "'"
        End With
        '' WriteDBLog sql
        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_SokuSizi_EP_Ins = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    DBDRV_SokuSizi_EP_Ins = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_SokuSizi_EP_Ins = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

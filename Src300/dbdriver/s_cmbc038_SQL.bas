Attribute VB_Name = "s_cmbc038_SQL"
Option Explicit

'�u���b�N���x�����o��  4/16 Yam�쐬

' �u���b�N�ꗗ
Public Type typ_BlkLbl
    BLOCKID As String * 12      ' �u���b�NID
    HIN(5) As tFullHinban       ' �i��
    WFINDDATE As String * 10    ' �ŏI�������t
    CRYNUM As String * 12       ' �����ԍ�
    INGOTPOS As Integer         ' �C���S�b�g���ʒu
    LENGTH As Integer           ' �u���b�N����
    REALLEN As Integer          ' �u���b�N������
    HINLEN(5) As Integer        ' �i�Ԓ���
    DIAMETER As Integer         ' ���a
    SBLOCKID As String * 12     ' �擪�u���b�NID
    BLOCKORDER As Integer       ' �u���b�N����
    HOLDCLS As String * 1       ' �z�[���h���  --- 2001/09/19 kuramoto �ǉ� ---
    PASSFLAG As String * 1      ' �ʉ߃t���O�@�@--- 200/04/16 Yam
End Type


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME040�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME040 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺
Public Function DBDRV_GetTBCME040(records() As typ_TBCME040, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, REALLEN, BLOCKID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, DELCLS, LSTATCLS," & _
              " RSTATCLS, HOLDCLS, BDCAUS, REGDATE, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE," & _
              " PASSFLAG "   '02/07/05 hama
    
    sqlBase = sqlBase & "From TBCME040"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME040 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .INGOTPOS = rs("INGOTPOS")       ' �������J�n�ʒu
            .LENGTH = rs("LENGTH")           ' ����
            .REALLEN = rs("REALLEN")         ' ������
            .BLOCKID = rs("BLOCKID")         ' �u���b�NID
            .KRPROCCD = rs("KRPROCCD")       ' ���݊Ǘ��H��
            .NOWPROC = rs("NOWPROC")         ' ���ݍH��
            .LPKRPROCCD = rs("LPKRPROCCD")   ' �ŏI�ʉߊǗ��H��
            .LASTPASS = rs("LASTPASS")       ' �ŏI�ʉߍH��
            .DELCLS = rs("DELCLS")           ' �폜�敪
            .LSTATCLS = rs("LSTATCLS")       ' �ŏI��ԋ敪
            .RSTATCLS = rs("RSTATCLS")       ' ������ԋ敪
            .HOLDCLS = rs("HOLDCLS")         ' �z�[���h�敪
            .BDCAUS = rs("BDCAUS")           ' �s�Ǘ��R
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
            .PASSFLAG = " "   ' �ʉ߃t���O�̃X�y�[�X�N���A '02/07/05 hama
             If rs("PASSFLAG") = "1" Then
                .PASSFLAG = rs("PASSFLAG")   ' �ʉ߃t���O '02/07/05 hama
            End If

        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME040 = FUNCTION_RETURN_SUCCESS
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
'����      :2001/08/24�쐬�@�쑺
Public Function DBDRV_GetTBCME042(records() As typ_TBCME042, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, SXLID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, DELCLS, LSTATCLS, HOLDCLS," & _
              " HINBAN, REVNUM, FACTORY, OPECOND, BDCAUS, COUNT, REGDATE, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE, " & _
              " PASSFLAG "   '02/04/16 Yam
    sqlBase = sqlBase & "From TBCME042"
    sql = sqlBase
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
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .INGOTPOS = rs("INGOTPOS")       ' �������J�n�ʒu
            .LENGTH = rs("LENGTH")           ' ����
            .SXLID = rs("SXLID")             ' SXLID
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H��
            .NOWPROC = rs("NOWPROC")         ' ���ݍH��
            .LPKRPROCCD = rs("LPKRPROCCD")   ' �ŏI�ʉߊǗ��H��
            .LASTPASS = rs("LASTPASS")       ' �ŏI�ʉߍH��
            .DELCLS = rs("DELCLS")           ' �폜�敪
            .LSTATCLS = rs("LSTATCLS")       ' �ŏI��ԋ敪
            .HOLDCLS = rs("HOLDCLS")         ' �z�[���h�敪
            .HINBAN = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .FACTORY = rs("FACTORY")         ' �H��
            .OPECOND = rs("OPECOND")         ' ���Ə���
            .BDCAUS = rs("BDCAUS")           ' �s�Ǘ��R
            .COUNT = rs("COUNT")             ' ����
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
            .PASSFLAG = " "   ' �ʉ߃t���O�̃X�y�[�X�N���A '02/04/16 Yam
            If rs("PASSFLAG") = "1" Then
                .PASSFLAG = rs("PASSFLAG")   ' �ʉ߃t���O '02/04/05 Yam
            End If
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME042 = FUNCTION_RETURN_SUCCESS
End Function


'�T�v      :H��N�ɉ���
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:sStaffID�@�@�@,I  ,String         �@,�Ј�ID
'      �@�@:pBlkMap �@�@�@,I  ,typ_BlkLbl     �@,�u���b�N�ꗗ
'      �@�@:sErrMsg �@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :2002/4/4 Yam  SXTL�̒ʉ�FLG�ɂP�����Ă�
Public Function DBDRV_s_cmbc038_Exec(ByVal sStaffID As String, pBlkMap() As typ_BlkLbl, sErrMsg As String) As FUNCTION_RETURN

    Dim sql As String
    Dim sDBName As String
    Dim recCnt As Long
    Dim iPos As Integer
    Dim i As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc038_SQL.bas -- Function DBDRV_s_cmbc038_Exec"
    sErrMsg = ""

    recCnt = UBound(pBlkMap)
    For i = 1 To recCnt
        With pBlkMap(i)
            '' SXL�Ǘ��̍X�V
            sDBName = "E042"
            iPos = .INGOTPOS + .LENGTH
            sql = "update TBCME040 set "
            sql = sql & "PASSFLAG='1' "
'            sql = sql & "PASSFLAG='1', "
'            sql = sql & "UPDDATE=sysdate "
            sql = sql & " where CRYNUM='" & .CRYNUM & "'"
            sql = sql & " and INGOTPOS>=" & .INGOTPOS
            sql = sql & " and INGOTPOS<" & iPos
            If OraDB.ExecuteSQL(sql) < 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
                DBDRV_s_cmbc038_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        End With
    Next i

    DBDRV_s_cmbc038_Exec = FUNCTION_RETURN_SUCCESS

proc_exit:    '' �I��
    gErr.Pop
    Exit Function

proc_err: '' �G���[�n���h��
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
    DBDRV_s_cmbc038_Exec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function



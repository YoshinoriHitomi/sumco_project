Attribute VB_Name = "s_cmzcF_cmmc001db_SQL"

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
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
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




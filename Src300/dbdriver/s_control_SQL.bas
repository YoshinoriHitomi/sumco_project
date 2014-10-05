Attribute VB_Name = "s_control_SQL"

'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMB005�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMB005 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCMB005_SQL.bas���ړ�)
Public Function DBDRV_GetTBCMB005(records() As typ_TBCMB005, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select SYSCLASS, CLASS, CODE, INFO1, INFO2, INFO3, INFO4, INFO5, INFO6, INFO7, INFO8, INFO9, NOTE, TSTAFFID," & _
              " REGDATE, KSTAFFID, UPDDATE "
    sqlBase = sqlBase & "From TBCMB005"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB005 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .SYSCLASS = rs("SYSCLASS")       ' �V�X�e���敪
            .Class = rs("CLASS")             ' �敪
            .CODE = rs("CODE")               ' �R�[�h
            .INFO1 = rs("INFO1")             ' ���P
            .INFO2 = rs("INFO2")             ' ���Q
            .INFO3 = rs("INFO3")             ' ���R
            .INFO4 = rs("INFO4")             ' ���S
            .INFO5 = rs("INFO5")             ' ���T
            .INFO6 = rs("INFO6")             ' ���U
            .INFO7 = rs("INFO7")             ' ���V
            .INFO8 = rs("INFO8")             ' ���W
            .INFO9 = rs("INFO9")             ' ���X
            .NOTE = rs("NOTE")               ' ���l
            .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMB005 = FUNCTION_RETURN_SUCCESS
End Function



Attribute VB_Name = "s_cmbc026_SQL"
Option Explicit
'                                     2001/06/20
'================================================
' DB�A�N�Z�X�֐�
' ��`���e: TBCMJ006 (�f�c����)
' �Q�Ɓ@�@: 060211_��������
'================================================

'------------------------------------------------
' ���[�U��`�^�̐錾
'------------------------------------------------
Public Type typ_cmjc001g_Disp
'    CRYNUM As String * 12           ' �����ԍ�
    POSITION As Integer             ' �ʒu
    SMPKBN As String * 1            ' �T���v���敪
    TRANCOND As String * 1          ' ��������
'    TRANCNT As Integer              ' ������
    SMPLNO As Long                  ' �T���v���m��  Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    SMPLUMU As String * 1           ' �T���v���L��
    hinban As String * 8            ' �i��
    REVNUM As Integer               ' ���i�ԍ������ԍ�
    Factory As String * 1           ' �H��
    OpeCond As String * 1           ' ���Ə���
    KRPROCCD As String * 5          ' �Ǘ��H���R�[�h
    PROCCODE As String * 5          ' �H���R�[�h
    GOUKI As String * 3             ' ���@
    MSRSDEN As Integer              ' ���茋�� Den
    MSRSLDL As Integer              ' ���茋�� L/DL
    MSRSDVD2 As Integer             ' ���茋�� DVD2
    MSLDL(1 To 15, 1 To 5) As Integer ' ����l LDL (����ʒu, n�Ԗ�)
    MSDEN(1 To 15, 1 To 5) As Integer ' ����l DEN (����ʒu, n�Ԗ�)
    MSDVD(1 To 5, 1) As Integer        ' ����l DVD (����ʒu, n�Ԗ�) 2002/7/4 tuku
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    MSZEROMN As Integer             ' L/DL0�A�����ŏ��l
    MSZEROMX As Integer             ' L/DL0�A�����ő�l
    PTNJUDGRES As String * 1        ' �p�^�[�����茋��
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
 '   TSTAFFID As String * 8          ' �o�^�Ј�ID
 '   REGDATE As Date                 ' �o�^���t
 '   KSTAFFID As String * 8          ' �X�V�Ј�ID
 '   UPDDATE As Date                 ' �X�V���t
 '   SENDFLAG As String * 1          ' ���M�t���O
 '   SENDDATE As Date                ' ���M���t
End Type



'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMJ006�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_cmjc001g_Disp ,���o���R�[�h
'          :SPLNUMs()     ,I  ,Integer      ,���o�����z��(�T���v��No)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/06/20�쐬�@����
Public Function DBDRV_Getcmjc001g_Disp(records() As typ_cmjc001g_Disp, SMPLNUMs() As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim sqlWhere As String  'SQL��WHERE����
Dim sqlGroup As String  'SQL��GROUP����
Dim sqlOrder As String  'SQL��Order����
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long
Dim POS As Integer
Dim n As Integer

    DBDRV_Getcmjc001g_Disp = FUNCTION_RETURN_FAILURE

    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001g_SQL.bas -- Function DBDRV_Getcmjc001g_Disp"

    sqlBase = "Select POSITION, SMPKBN, TRANCOND, Max(TRANCNT), SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MSRSDEN, MSRSLDL, MSRSDVD2, "
    For POS = 1 To 15
        For n = 1 To 5
            sqlBase = sqlBase & "MS" & POS & "LDL" & n & ", "
        Next
        For n = 1 To 5
            sqlBase = sqlBase & "MS" & POS & "DEN" & n
            If POS = 15 And n = 5 Then
                Exit For
            Else
                sqlBase = sqlBase & ", "
            End If
        Next
        If POS = 15 Then
            sqlBase = sqlBase & " """
        End If
    Next
    sqlBase = sqlBase & "From TBCMJ006"
        
    ''���o����(�����NO)�̎��o��
    sqlWhere = "Where SMPLNO in ("
    For i = 1 To UBound(SMPLNUMs)
        sqlWhere = sqlWhere & "'" & SMPLNUMs(i) & "'"
        If i < UBound(SMPLNUMs) Then
            sqlWhere = sqlWhere & ", "
        End If
    Next
    sqlWhere = sqlWhere & ") "
    sqlGroup = "GROUP BY CRYNUM, POSITION, SMPKBN, TRANCOND "
    sqlOrder = "ORDER BY POSITION"
    sql = sqlBase & sqlWhere & sqlGroup & sqlOrder

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_Getcmjc001g_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .POSITION = rs("POSITION")       ' �ʒu
            .SMPKBN = rs("SMPKBN")           ' �T���v���敪
            .TRANCOND = rs("TRANCOND")       ' ��������
            .SMPLNO = rs("SMPLNO")           ' �T���v���m��
            .SMPLUMU = rs("SMPLUMU")         ' �T���v���L��
            .hinban = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .Factory = rs("FACTORY")         ' �H��
            .OpeCond = rs("OPECOND")         ' ���Ə���
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
            .GOUKI = rs("GOUKI")             ' ���@
            .MSRSDEN = rs("MSRSDEN")         ' ���茋�� Den
            .MSRSLDL = rs("MSRSLDL")         ' ���茋�� L/DL
            .MSRSDVD2 = rs("MSRSDVD2")       ' ���茋�� DVD2
            
            For POS = 1 To 15
                For n = 1 To 5
                    .MSLDL(POS, n) = rs("MS" & Format$(POS, "00") & "LDL" & n) ' ����l(pos) L/DL(n)
                    .MSDEN(POS, n) = rs("MS" & Format$(POS, "00") & "DEN" & n) ' ����l(pos) Den(n)
                Next
            Next

        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_Getcmjc001g_Disp = FUNCTION_RETURN_SUCCESS

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

'�T�v      :�����œn���ꂽ���R�[�h��TBCMJ006�ɒǉ�����
'���Ұ�    :�ϐ���        ,IO ,�^            ,����
'          :record        ,I  ,typ_cmjc001g_Disp ,���o���R�[�h
'          :CRYNUM        ,I  ,String       ,�����ԍ�
'          :TSTAFFID      ,I  ,String       ,�o�^�Ј�ID
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :�����񐔂̓e�[�u����̍ő�l+1�Ƃ���B
'����      :2001/06/22(Fri)�쐬�@����

Public Function DBDRV_Getcmjc001g_Exec(record As typ_cmjc001g_Disp, CRYNUM$, TSTAFFID$) As FUNCTION_RETURN

Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL�x�[�X����
Dim sqlWhere As String  'SQLWhere����
Dim sqlGroup As String  'SQLGroup����
Dim POS As Integer
Dim n As Integer

'    CRYNUM             �����ԍ��@�ˈ���
'    TRANCNT         �@ �����񐔁@�ˍő�
'   TSTAFFID            �o�^�Ј�ID�@�ˈ���
 '   REGDATE �@�@�@     �o�^���t�@��SYSDATE
 '   KSTAFFID           �X�V�Ј�ID�@��" "
 '   UPDDATE            �X�V���t�@��SYSDATE
 '   SENDFLAG           ���M�t���O�@��"0"
 '   SENDDATE           ���M���t�@��SYSDATE
     
    DBDRV_Getcmjc001g_Exec = FUNCTION_RETURN_FAILURE

    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001g_SQL.bas -- Function DBDRV_Getcmjc001g_Exec"

    sqlBase = "Insert into TBCMJ006 (CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MSRSDEN, MSRSLDL, MSRSDVD2, "
            For POS = 1 To 15
                For n = 1 To 5
                    sqlBase = sqlBase & "MS" & Format(POS, "00") & "LDL" & n & ", "
                Next
                For n = 1 To 5
                    sqlBase = sqlBase & "MS" & Format(POS, "00") & "DEN" & n & ", "
                Next
            Next
            For POS = 1 To 5
                sqlBase = sqlBase & "MS" & Format(POS, "00") & "DVD2" & ", "
            Next
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    sqlBase = sqlBase & "MSZEROMN, MSZEROMX, PTNJUDGRES, "
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    sqlBase = sqlBase & "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE)"
    sqlBase = sqlBase & " select '" & CRYNUM & "', " & record.POSITION & ", '" & record.SMPKBN & "', '" & record.TRANCOND & "', nvl(MAX(TRANCNT),0) + 1, " & _
               record.SMPLNO & ", '" & record.SMPLUMU & "', '" & record.hinban & "', " & record.REVNUM & ", '" & record.Factory & "', '" & _
               record.OpeCond & "', '" & record.KRPROCCD & "', '" & record.PROCCODE & "', '" & record.GOUKI & "', " & record.MSRSDEN & ", " & _
               record.MSRSLDL & ", " & record.MSRSDVD2 & ", "
            For POS = 1 To 15
                For n = 1 To 5
                    sqlBase = sqlBase & record.MSLDL(POS, n) & ", "
                Next
                For n = 1 To 5
                    sqlBase = sqlBase & record.MSDEN(POS, n) & ", "
                Next
            Next
            
            For POS = 1 To 5 'DVD2���ړ��̓J�����ǉ��@2002/7/5 tuku
                    sqlBase = sqlBase & record.MSDVD(POS, 1) & ", "
            Next

'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    sqlBase = sqlBase & record.MSZEROMN & "," & record.MSZEROMX & ",'" & record.PTNJUDGRES & "',"
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    sqlBase = sqlBase & "'" & TSTAFFID & "', SYSDATE, ' ', SYSDATE, '0', SYSDATE from TBCMJ006 "
              
  
    sqlWhere = "where (CRYNUM='" & CRYNUM & "') and (POSITION=" & record.POSITION & ") and (SMPKBN='" & record.SMPKBN & "') and (TRANCOND='" & record.TRANCOND & "') "
'    sqlGroup = "group by CRYNUM, POSITION, SMPKBN, TRANCOND"
    sql = sqlBase & sqlWhere & sqlGroup
            
  ''SQL�̎��s
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001g_Exec = FUNCTION_RETURN_SUCCESS
    

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





'�T�v      :�e�[�u���uTBCMJ006�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMJ006 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺
Public Function DBDRV_GetTBCMJ006(records() As typ_TBCMJ006, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, HINBAN, REVNUM, FACTORY, OPECOND, KRPROCCD," & _
              " PROCCODE, GOUKI, MSRSDEN, MSRSLDL, MSRSDVD2, MS01LDL1, MS01LDL2, MS01LDL3, MS01LDL4, MS01LDL5, MS01DEN1, MS01DEN2," & _
              " MS01DEN3, MS01DEN4, MS01DEN5, MS02LDL1, MS02LDL2, MS02LDL3, MS02LDL4, MS02LDL5, MS02DEN1, MS02DEN2, MS02DEN3," & _
              " MS02DEN4, MS02DEN5, MS03LDL1, MS03LDL2, MS03LDL3, MS03LDL4, MS03LDL5, MS03DEN1, MS03DEN2, MS03DEN3, MS03DEN4," & _
              " MS03DEN5, MS04LDL1, MS04LDL2, MS04LDL3, MS04LDL4, MS04LDL5, MS04DEN1, MS04DEN2, MS04DEN3, MS04DEN4, MS04DEN5," & _
              " MS05LDL1, MS05LDL2, MS05LDL3, MS05LDL4, MS05LDL5, MS05DEN1, MS05DEN2, MS05DEN3, MS05DEN4, MS05DEN5, MS06LDL1," & _
              " MS06LDL2, MS06LDL3, MS06LDL4, MS06LDL5, MS06DEN1, MS06DEN2, MS06DEN3, MS06DEN4, MS06DEN5, MS07LDL1, MS07LDL2," & _
              " MS07LDL3, MS07LDL4, MS07LDL5, MS07DEN1, MS07DEN2, MS07DEN3, MS07DEN4, MS07DEN5, MS08LDL1, MS08LDL2, MS08LDL3," & _
              " MS08LDL4, MS08LDL5, MS08DEN1, MS08DEN2, MS08DEN3, MS08DEN4, MS08DEN5, MS09LDL1, MS09LDL2, MS09LDL3, MS09LDL4," & _
              " MS09LDL5, MS09DEN1, MS09DEN2, MS09DEN3, MS09DEN4, MS09DEN5, MS10LDL1, MS10LDL2, MS10LDL3, MS10LDL4, MS10LDL5," & _
              " MS10DEN1, MS10DEN2, MS10DEN3, MS10DEN4, MS10DEN5, MS11LDL1, MS11LDL2, MS11LDL3, MS11LDL4, MS11LDL5, MS11DEN1," & _
              " MS11DEN2, MS11DEN3, MS11DEN4, MS11DEN5, MS12LDL1, MS12LDL2, MS12LDL3, MS12LDL4, MS12LDL5, MS12DEN1, MS12DEN2," & _
              " MS12DEN3, MS12DEN4, MS12DEN5, MS13LDL1, MS13LDL2, MS13LDL3, MS13LDL4, MS13LDL5, MS13DEN1, MS13DEN2, MS13DEN3," & _
              " MS13DEN4, MS13DEN5, MS14LDL1, MS14LDL2, MS14LDL3, MS14LDL4, MS14LDL5, MS14DEN1, MS14DEN2, MS14DEN3, MS14DEN4," & _
              " MS14DEN5, MS15LDL1, MS15LDL2, MS15LDL3, MS15LDL4, MS15LDL5, MS15DEN1, MS15DEN2, MS15DEN3, MS15DEN4, MS15DEN5," & _
              " MS01DVD2,  MS02DVD2 , MS03DVD2 , MS04DVD2 , MS05DVD2 , TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
    sqlBase = sqlBase & ",MSZEROMN, MSZEROMX, PTNJUDGRES "
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    sqlBase = sqlBase & "From TBCMJ006"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ006 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .POSITION = rs("POSITION")       ' �ʒu
            .SMPKBN = rs("SMPKBN")           ' �T���v���敪
            .TRANCOND = rs("TRANCOND")       ' ��������
            .TRANCNT = rs("TRANCNT")         ' ������
            .SMPLNO = rs("SMPLNO")           ' �T���v���m��
            .SMPLUMU = rs("SMPLUMU")         ' �T���v���L��
            .hinban = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .Factory = rs("FACTORY")         ' �H��
            .OpeCond = rs("OPECOND")         ' ���Ə���
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
            .GOUKI = rs("GOUKI")             ' ���@
            .MSRSDEN = rs("MSRSDEN")         ' ���茋�� Den
            .MSRSLDL = rs("MSRSLDL")         ' ���茋�� L/DL
            .MSRSDVD2 = rs("MSRSDVD2")       ' ���茋�� DVD2
            .MS01LDL1 = rs("MS01LDL1")       ' ����l01 L/DL1
            .MS01LDL2 = rs("MS01LDL2")       ' ����l01 L/DL2
            .MS01LDL3 = rs("MS01LDL3")       ' ����l01 L/DL3
            .MS01LDL4 = rs("MS01LDL4")       ' ����l01 L/DL4
            .MS01LDL5 = rs("MS01LDL5")       ' ����l01 L/DL5
            .MS01DEN1 = rs("MS01DEN1")       ' ����l01 Den1
            .MS01DEN2 = rs("MS01DEN2")       ' ����l01 Den2
            .MS01DEN3 = rs("MS01DEN3")       ' ����l01 Den3
            .MS01DEN4 = rs("MS01DEN4")       ' ����l01 Den4
            .MS01DEN5 = rs("MS01DEN5")       ' ����l01 Den5
            .MS02LDL1 = rs("MS02LDL1")       ' ����l02 L/DL1
            .MS02LDL2 = rs("MS02LDL2")       ' ����l02 L/DL2
            .MS02LDL3 = rs("MS02LDL3")       ' ����l02 L/DL3
            .MS02LDL4 = rs("MS02LDL4")       ' ����l02 L/DL4
            .MS02LDL5 = rs("MS02LDL5")       ' ����l02 L/DL5
            .MS02DEN1 = rs("MS02DEN1")       ' ����l02 Den1
            .MS02DEN2 = rs("MS02DEN2")       ' ����l02 Den2
            .MS02DEN3 = rs("MS02DEN3")       ' ����l02 Den3
            .MS02DEN4 = rs("MS02DEN4")       ' ����l02 Den4
            .MS02DEN5 = rs("MS02DEN5")       ' ����l02 Den5
            .MS03LDL1 = rs("MS03LDL1")       ' ����l03 L/DL1
            .MS03LDL2 = rs("MS03LDL2")       ' ����l03 L/DL2
            .MS03LDL3 = rs("MS03LDL3")       ' ����l03 L/DL3
            .MS03LDL4 = rs("MS03LDL4")       ' ����l03 L/DL4
            .MS03LDL5 = rs("MS03LDL5")       ' ����l03 L/DL5
            .MS03DEN1 = rs("MS03DEN1")       ' ����l03 Den1
            .MS03DEN2 = rs("MS03DEN2")       ' ����l03 Den2
            .MS03DEN3 = rs("MS03DEN3")       ' ����l03 Den3
            .MS03DEN4 = rs("MS03DEN4")       ' ����l03 Den4
            .MS03DEN5 = rs("MS03DEN5")       ' ����l03 Den5
            .MS04LDL1 = rs("MS04LDL1")       ' ����l04 L/DL1
            .MS04LDL2 = rs("MS04LDL2")       ' ����l04 L/DL2
            .MS04LDL3 = rs("MS04LDL3")       ' ����l04 L/DL3
            .MS04LDL4 = rs("MS04LDL4")       ' ����l04 L/DL4
            .MS04LDL5 = rs("MS04LDL5")       ' ����l04 L/DL5
            .MS04DEN1 = rs("MS04DEN1")       ' ����l04 Den1
            .MS04DEN2 = rs("MS04DEN2")       ' ����l04 Den2
            .MS04DEN3 = rs("MS04DEN3")       ' ����l04 Den3
            .MS04DEN4 = rs("MS04DEN4")       ' ����l04 Den4
            .MS04DEN5 = rs("MS04DEN5")       ' ����l04 Den5
            .MS05LDL1 = rs("MS05LDL1")       ' ����l05 L/DL1
            .MS05LDL2 = rs("MS05LDL2")       ' ����l05 L/DL2
            .MS05LDL3 = rs("MS05LDL3")       ' ����l05 L/DL3
            .MS05LDL4 = rs("MS05LDL4")       ' ����l05 L/DL4
            .MS05LDL5 = rs("MS05LDL5")       ' ����l05 L/DL5
            .MS05DEN1 = rs("MS05DEN1")       ' ����l05 Den1
            .MS05DEN2 = rs("MS05DEN2")       ' ����l05 Den2
            .MS05DEN3 = rs("MS05DEN3")       ' ����l05 Den3
            .MS05DEN4 = rs("MS05DEN4")       ' ����l05 Den4
            .MS05DEN5 = rs("MS05DEN5")       ' ����l05 Den5
            .MS06LDL1 = rs("MS06LDL1")       ' ����l06 L/DL1
            .MS06LDL2 = rs("MS06LDL2")       ' ����l06 L/DL2
            .MS06LDL3 = rs("MS06LDL3")       ' ����l06 L/DL3
            .MS06LDL4 = rs("MS06LDL4")       ' ����l06 L/DL4
            .MS06LDL5 = rs("MS06LDL5")       ' ����l06 L/DL5
            .MS06DEN1 = rs("MS06DEN1")       ' ����l06 Den1
            .MS06DEN2 = rs("MS06DEN2")       ' ����l06 Den2
            .MS06DEN3 = rs("MS06DEN3")       ' ����l06 Den3
            .MS06DEN4 = rs("MS06DEN4")       ' ����l06 Den4
            .MS06DEN5 = rs("MS06DEN5")       ' ����l06 Den5
            .MS07LDL1 = rs("MS07LDL1")       ' ����l07 L/DL1
            .MS07LDL2 = rs("MS07LDL2")       ' ����l07 L/DL2
            .MS07LDL3 = rs("MS07LDL3")       ' ����l07 L/DL3
            .MS07LDL4 = rs("MS07LDL4")       ' ����l07 L/DL4
            .MS07LDL5 = rs("MS07LDL5")       ' ����l07 L/DL5
            .MS07DEN1 = rs("MS07DEN1")       ' ����l07 Den1
            .MS07DEN2 = rs("MS07DEN2")       ' ����l07 Den2
            .MS07DEN3 = rs("MS07DEN3")       ' ����l07 Den3
            .MS07DEN4 = rs("MS07DEN4")       ' ����l07 Den4
            .MS07DEN5 = rs("MS07DEN5")       ' ����l07 Den5
            .MS08LDL1 = rs("MS08LDL1")       ' ����l08 L/DL1
            .MS08LDL2 = rs("MS08LDL2")       ' ����l08 L/DL2
            .MS08LDL3 = rs("MS08LDL3")       ' ����l08 L/DL3
            .MS08LDL4 = rs("MS08LDL4")       ' ����l08 L/DL4
            .MS08LDL5 = rs("MS08LDL5")       ' ����l08 L/DL5
            .MS08DEN1 = rs("MS08DEN1")       ' ����l08 Den1
            .MS08DEN2 = rs("MS08DEN2")       ' ����l08 Den2
            .MS08DEN3 = rs("MS08DEN3")       ' ����l08 Den3
            .MS08DEN4 = rs("MS08DEN4")       ' ����l08 Den4
            .MS08DEN5 = rs("MS08DEN5")       ' ����l08 Den5
            .MS09LDL1 = rs("MS09LDL1")       ' ����l09 L/DL1
            .MS09LDL2 = rs("MS09LDL2")       ' ����l09 L/DL2
            .MS09LDL3 = rs("MS09LDL3")       ' ����l09 L/DL3
            .MS09LDL4 = rs("MS09LDL4")       ' ����l09 L/DL4
            .MS09LDL5 = rs("MS09LDL5")       ' ����l09 L/DL5
            .MS09DEN1 = rs("MS09DEN1")       ' ����l09 Den1
            .MS09DEN2 = rs("MS09DEN2")       ' ����l09 Den2
            .MS09DEN3 = rs("MS09DEN3")       ' ����l09 Den3
            .MS09DEN4 = rs("MS09DEN4")       ' ����l09 Den4
            .MS09DEN5 = rs("MS09DEN5")       ' ����l09 Den5
            .MS10LDL1 = rs("MS10LDL1")       ' ����l10 L/DL1
            .MS10LDL2 = rs("MS10LDL2")       ' ����l10 L/DL2
            .MS10LDL3 = rs("MS10LDL3")       ' ����l10 L/DL3
            .MS10LDL4 = rs("MS10LDL4")       ' ����l10 L/DL4
            .MS10LDL5 = rs("MS10LDL5")       ' ����l10 L/DL5
            .MS10DEN1 = rs("MS10DEN1")       ' ����l10 Den1
            .MS10DEN2 = rs("MS10DEN2")       ' ����l10 Den2
            .MS10DEN3 = rs("MS10DEN3")       ' ����l10 Den3
            .MS10DEN4 = rs("MS10DEN4")       ' ����l10 Den4
            .MS10DEN5 = rs("MS10DEN5")       ' ����l10 Den5
            .MS11LDL1 = rs("MS11LDL1")       ' ����l11 L/DL1
            .MS11LDL2 = rs("MS11LDL2")       ' ����l11 L/DL2
            .MS11LDL3 = rs("MS11LDL3")       ' ����l11 L/DL3
            .MS11LDL4 = rs("MS11LDL4")       ' ����l11 L/DL4
            .MS11LDL5 = rs("MS11LDL5")       ' ����l11 L/DL5
            .MS11DEN1 = rs("MS11DEN1")       ' ����l11 Den1
            .MS11DEN2 = rs("MS11DEN2")       ' ����l11 Den2
            .MS11DEN3 = rs("MS11DEN3")       ' ����l11 Den3
            .MS11DEN4 = rs("MS11DEN4")       ' ����l11 Den4
            .MS11DEN5 = rs("MS11DEN5")       ' ����l11 Den5
            .MS12LDL1 = rs("MS12LDL1")       ' ����l12 L/DL1
            .MS12LDL2 = rs("MS12LDL2")       ' ����l12 L/DL2
            .MS12LDL3 = rs("MS12LDL3")       ' ����l12 L/DL3
            .MS12LDL4 = rs("MS12LDL4")       ' ����l12 L/DL4
            .MS12LDL5 = rs("MS12LDL5")       ' ����l12 L/DL5
            .MS12DEN1 = rs("MS12DEN1")       ' ����l12 Den1
            .MS12DEN2 = rs("MS12DEN2")       ' ����l12 Den2
            .MS12DEN3 = rs("MS12DEN3")       ' ����l12 Den3
            .MS12DEN4 = rs("MS12DEN4")       ' ����l12 Den4
            .MS12DEN5 = rs("MS12DEN5")       ' ����l12 Den5
            .MS13LDL1 = rs("MS13LDL1")       ' ����l13 L/DL1
            .MS13LDL2 = rs("MS13LDL2")       ' ����l13 L/DL2
            .MS13LDL3 = rs("MS13LDL3")       ' ����l13 L/DL3
            .MS13LDL4 = rs("MS13LDL4")       ' ����l13 L/DL4
            .MS13LDL5 = rs("MS13LDL5")       ' ����l13 L/DL5
            .MS13DEN1 = rs("MS13DEN1")       ' ����l13 Den1
            .MS13DEN2 = rs("MS13DEN2")       ' ����l13 Den2
            .MS13DEN3 = rs("MS13DEN3")       ' ����l13 Den3
            .MS13DEN4 = rs("MS13DEN4")       ' ����l13 Den4
            .MS13DEN5 = rs("MS13DEN5")       ' ����l13 Den5
            .MS14LDL1 = rs("MS14LDL1")       ' ����l14 L/DL1
            .MS14LDL2 = rs("MS14LDL2")       ' ����l14 L/DL2
            .MS14LDL3 = rs("MS14LDL3")       ' ����l14 L/DL3
            .MS14LDL4 = rs("MS14LDL4")       ' ����l14 L/DL4
            .MS14LDL5 = rs("MS14LDL5")       ' ����l14 L/DL5
            .MS14DEN1 = rs("MS14DEN1")       ' ����l14 Den1
            .MS14DEN2 = rs("MS14DEN2")       ' ����l14 Den2
            .MS14DEN3 = rs("MS14DEN3")       ' ����l14 Den3
            .MS14DEN4 = rs("MS14DEN4")       ' ����l14 Den4
            .MS14DEN5 = rs("MS14DEN5")       ' ����l14 Den5
            .MS15LDL1 = rs("MS15LDL1")       ' ����l15 L/DL1
            .MS15LDL2 = rs("MS15LDL2")       ' ����l15 L/DL2
            .MS15LDL3 = rs("MS15LDL3")       ' ����l15 L/DL3
            .MS15LDL4 = rs("MS15LDL4")       ' ����l15 L/DL4
            .MS15LDL5 = rs("MS15LDL5")       ' ����l15 L/DL5
            .MS15DEN1 = rs("MS15DEN1")       ' ����l15 Den1
            .MS15DEN2 = rs("MS15DEN2")       ' ����l15 Den2
            .MS15DEN3 = rs("MS15DEN3")       ' ����l15 Den3
            .MS15DEN4 = rs("MS15DEN4")       ' ����l15 Den4
            .MS15DEN5 = rs("MS15DEN5")       ' ����l15 Den5
            'NULL �`�F�b�N
            If IsNull(rs("MS01DVD2")) = False Then
                .MS01DVD2 = rs("MS01DVD2")       ' ����l01 DVD2   2002/7/02 tuku
            Else
                .MS01DVD2 = DEF_PARAM_VALUE
            End If
            If IsNull(rs("MS02DVD2")) = False Then
                .MS02DVD2 = rs("MS02DVD2")       ' ����l01 DVD2   2002/7/02 tuku
            Else
                .MS02DVD2 = DEF_PARAM_VALUE
            End If
            If IsNull(rs("MS03DVD2")) = False Then
                .MS03DVD2 = rs("MS03DVD2")       ' ����l01 DVD2   2002/7/02 tuku
            Else
                .MS03DVD2 = DEF_PARAM_VALUE
            End If
            If IsNull(rs("MS04DVD2")) = False Then
                .MS04DVD2 = rs("MS04DVD2")       ' ����l01 DVD2   2002/7/02 tuku
            Else
                .MS04DVD2 = DEF_PARAM_VALUE
            End If
            If IsNull(rs("MS05DVD2")) = False Then
                .MS05DVD2 = rs("MS05DVD2")       ' ����l01 DVD2   2002/7/02 tuku
            Else
                .MS05DVD2 = DEF_PARAM_VALUE
            End If
            .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
            If IsNull(rs("MSZEROMN")) = False Then
                .MSZEROMN = rs("MSZEROMN")
            Else
                .MSZEROMN = DEF_PARAM_VALUE
            End If
            If IsNull(rs("MSZEROMX")) = False Then
                .MSZEROMX = rs("MSZEROMX")
            Else
                .MSZEROMX = DEF_PARAM_VALUE
            End If
            If IsNull(rs("PTNJUDGRES")) = False Then
                .PTNJUDGRES = rs("PTNJUDGRES")
            Else
                .PTNJUDGRES = " "
            End If
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMJ006 = FUNCTION_RETURN_SUCCESS
End Function


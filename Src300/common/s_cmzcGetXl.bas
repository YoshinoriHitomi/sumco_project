Attribute VB_Name = "s_cmzcGetXl"
Option Explicit


Public Function GetXl(CRYNUM$, FormName$) As c_cmzcXl
Dim sqlWhere$
Dim sqlWherePlan$
Dim sqlWhereBlk$
Dim Xl As c_cmzcXl
Dim RET As FUNCTION_RETURN

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetXl"

    ''���ʂ�WHERE������
    sqlWhere = " Where (CRYNUM='" & CRYNUM & "')"
    '���`���[�W�A�ԑΉ��i�w��No�X���ύX�j Y.K 2004/09.03 �V���{�X����
    sqlWherePlan = " Where (SUBSTR(CRYNUM,1,7)='" & Mid(CRYNUM, 1, 7) & "') and (SUBSTR(CRYNUM,9,1)='" & Mid(CRYNUM, 9, 1) & "')"
'    sqlWherePlan = " Where (SUBSTR(CRYNUM,1,7)='" & Mid(CRYNUM, 1, 7) & "') "
    
    sqlWhereBlk = " Where (CRYNUM='" & CRYNUM & "') and (INGOTPOS>=0)"
    
    If FormName = "f_cmbc009_3" Then
        Set Xl = New c_cmzcXl
    Else
        ''���������擾����
        RET = GetTBCME037(Xl, sqlWhere)
        If RET = FUNCTION_RETURN_FAILURE Then
            ''�������̓Ǎ��Ɏ��s����
            Set GetXl = Nothing
            GoTo proc_exit
        End If
    End If
        
    ''��ʖ��ɂ��A�e�f�[�^��ǂݍ���
    Select Case FormName
      Case "f_cmbc009_3"     ' �u���b�N�g����
            RET = GetTBCME038(Xl.BlkPlans, sqlWherePlan)
            RET = GetTBCME039(Xl.HinPlans, sqlWherePlan)
      Case "f_cmbc016_1"     ' ���H���o��
            RET = GetTBCME039(Xl.HinPlans, sqlWherePlan)
            RET = GetTBCME040(Xl.Blks, sqlWhereBlk)
            RET = GetTBCME041(Xl.Hins, sqlWhere)
            RET = GetTBCME045(Xl.Cuts, CRYNUM)
      Case "f_cmbc018_2"     ' �ؒf
            RET = GetTBCME039(Xl.HinPlans, sqlWherePlan)
            RET = GetTBCME040(Xl.Blks, sqlWhereBlk)
            RET = GetTBCME041(Xl.Hins, sqlWhere)
            RET = GetTBCME045(Xl.Cuts, CRYNUM)
'      Case "f_cmbc030_1"     ' �҂��ꗗ
'            RET = GetTBCME040(xl.Blks, sqlWhereBlk)
'            RET = GetTBCME041(xl.Hins, sqlWhere)
'            ret = GetTBCME044(xl.WfSmps, sqlWhere)
      Case "f_cmbc032_1"     ' �҂��ꗗ
''���X�V START SPT�p���э쐬���@�ύX 2006/06/30 SMP-OKAMOTO
            RET = GetBlockData_2(Xl.Blks, CRYNUM)
'            RET = GetTBCME040(xl.Blks, sqlWhereBlk)
''���X�V END   SPT�p���э쐬���@�ύX 2006/06/30 SMP-OKAMOTO
            RET = GetTBCME041(Xl.Hins, sqlWhere)
      Case "f_cmbc033_1"     ' �҂��ꗗ
''���X�V START SPT�p���э쐬���@�ύX 2006/06/30 SMP-OKAMOTO
            RET = GetBlockData_2(Xl.Blks, CRYNUM)
'            RET = GetTBCME040(xl.Blks, sqlWhereBlk)
''���X�V END   SPT�p���э쐬���@�ύX 2006/06/30 SMP-OKAMOTO
            RET = GetTBCME041(Xl.Hins, sqlWhere)

      Case "f_cmbc030_1"     ' �҂��ꗗ
'            RET = GetTBCME040(xl.Blks, sqlWhereBlk)
''���X�V START SPT�p���э쐬���@�ύX 2006/05/12 SMP-OKAMOTO
            RET = GetBlockData_2(Xl.Blks, CRYNUM)
'            RET = GetBlockData(xl.Blks, CRYNUM)     '��ۯ��Ǘ�(TBCME040)�Q�ƒ�~�@05/10/05 ooba
''���X�V END   SPT�p���э쐬���@�ύX 2006/05/12 SMP-OKAMOTO
            RET = GetTBCME041(Xl.Hins, sqlWhere)
'            ret = GetTBCME044(xl.WfSmps, sqlWhere)
      Case "f_cmbc031_1"   ' �i2002/07�@���g�p�@��), "f_cmkc001e"    ' �Đؒf�w��
'            RET = GetTBCME040(xl.Blks, sqlWhereBlk)
''���X�V START SPT�p���э쐬���@�ύX 2006/05/12 SMP-OKAMOTO
            RET = GetBlockData_2(Xl.Blks, CRYNUM)
'            RET = GetBlockData(xl.Blks, CRYNUM)     '��ۯ��Ǘ�(TBCME040)�Q�ƒ�~�@05/10/05 ooba
''���X�V END   SPT�p���э쐬���@�ύX 2006/05/12 SMP-OKAMOTO
            RET = GetTBCME041(Xl.Hins, sqlWhere)
'            ret = GetTBCME045(xl.cuts, crynum)
      Case "f_cmbc033_2"     ' �����w��
            RET = GetTBCME040(Xl.Blks, sqlWhereBlk)
            RET = GetTBCME041(Xl.Hins, sqlWhere)
''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{
            sqlWhere = " Where (a.XTALCB='" & CRYNUM & "')"
''���ǉ�END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{
            RET = GetTBCME042(Xl.Sxls, sqlWhere)
            sqlWhere = " Where (XTALCW='" & CRYNUM & "')"
            RET = GetTBCME044(Xl.WfSmps, sqlWhere)
      Case "f_cmbc035_1"     ' �������ύX
            RET = GetTBCME040(Xl.Blks, sqlWhereBlk)
            
''���폜START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{
'            RET = GetTBCME042(xl.Sxls, sqlWhere)
'            RET = GetTBCME044(xl.WfSmps, sqlWhere)
''���폜END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{
''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{
            RET = GetTBCME044(Xl.WfSmps, sqlWhere)
            sqlWhere = " Where (a.XTALCB='" & CRYNUM & "')"
            RET = GetTBCME042(Xl.Sxls, sqlWhere)
''���ǉ�END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{

      Case "f_cmbc036_2"     ' �����w���ύX
            RET = GetTBCME040(Xl.Blks, sqlWhereBlk)
            
''���폜START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{
'            RET = GetTBCME042(xl.Sxls, sqlWhere)
'            RET = GetTBCME044(xl.WfSmps, sqlWhere)
''���폜END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{
''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{
            RET = GetTBCME044(Xl.WfSmps, sqlWhere)
            sqlWhere = " Where (a.XTALCB='" & CRYNUM & "')"
            RET = GetTBCME042(Xl.Sxls, sqlWhere)
''���ǉ�END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{

            RET = GetReject(Xl.Rejs, CRYNUM)
      Case "f_cmbc039_3"     ' �Ĕ���
            RET = GetTBCME040(Xl.Blks, sqlWhereBlk)
            RET = GetTBCME041(Xl.Hins, sqlWhere)
''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{
            sqlWhere = " Where (a.XTALCB='" & CRYNUM & "')"
''���ǉ�END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{
            RET = GetTBCME042(Xl.Sxls, sqlWhere)
            sqlWhere = " where XTALCW = '" & CRYNUM & "' "
            RET = GetTBCME044(Xl.WfSmps, sqlWhere)
            RET = GetReject(Xl.Rejs, CRYNUM)
      Case "block"          ' �u���b�N���̂�
            RET = GetTBCME040(Xl.Blks, sqlWhereBlk)
      Case "hinban"          ' �i�ԏ��̂�
            RET = GetTBCME041(Xl.Hins, sqlWhere)
      Case "All"            ' �f�o�b�O�p �S���
            RET = GetTBCME038(Xl.BlkPlans, sqlWherePlan)
            RET = GetTBCME039(Xl.HinPlans, sqlWherePlan)
            RET = GetTBCME040(Xl.Blks, sqlWhereBlk)
            RET = GetTBCME041(Xl.Hins, sqlWhere)
            
''���폜START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{
'            RET = GetTBCME042(xl.Sxls, sqlWhere)
'            RET = GetTBCME043(xl.XlSmps, sqlWhere)
'            RET = GetTBCME044(xl.WfSmps, sqlWhere)
''���폜END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{
''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{
            RET = GetTBCME043(Xl.XlSmps, sqlWhere)
            RET = GetTBCME044(Xl.WfSmps, sqlWhere)
            sqlWhere = " Where (a.XTALCB='" & CRYNUM & "')"
            RET = GetTBCME042(Xl.Sxls, sqlWhere)
''���ǉ�END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{

            RET = GetTBCME045(Xl.Cuts, CRYNUM)
            RET = GetReject(Xl.Rejs, CRYNUM)
      Case Else
            Debug.Print "GetXl() : FormName ���z��O"
            Set Xl = Nothing
    End Select
    
    Set GetXl = Xl

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�e�[�u���uTBCME037�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME037 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/06/29�쐬�@�쑺
Public Function GetTBCME037(Xl As c_cmzcXl, Optional sqlWhere$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetTBCME037"

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, DELCLS, KRPROCCD, PROCCD, LPKRPROCCD, LASTPASS, RPHINBAN, RPREVNUM, RPFACT, RPOPCOND, PRODCOND," & _
              " PGID, UPLENGTH, TOPLENG, BODYLENG, BOTLENG, FREELENG, DIAMETER, CHARGE, SEED, ADDDPCLS, ADDDPPOS, ADDDPVAL," & _
              " REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME037"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
        sql = sql & sqlWhere
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetTBCME037 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    If rs.RecordCount > 0 Then
        Set Xl = New c_cmzcXl
        With Xl
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
        End With
    Else
        GetTBCME037 = FUNCTION_RETURN_FAILURE
        Set Xl = Nothing
        GoTo proc_exit
    End If
    rs.Close

    GetTBCME037 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�e�[�u���uTBCME038�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :col           ,O  ,c_cmczBlkPlans,�u���b�N�݌v�R���N�V����
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/07/01�쐬�@�쑺
Private Function GetTBCME038(col As c_cmzcBlkPlans, Optional sqlWhere$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long
Dim target As c_cmzcBlkPlan

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetTBCME038"

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, USECLASS, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME038"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
        sql = sql & sqlWhere
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetTBCME038 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcBlkPlan
        With target
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .INGOTPOS = rs("INGOTPOS")       ' �������J�n�ʒu
            .Length = rs("LENGTH")           ' ����
            .USECLASS = rs("USECLASS")       ' �g�p�敪
        End With
        col.Add target
        rs.MoveNext
    Next
    rs.Close

    GetTBCME038 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�e�[�u���uTBCME039�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME039 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/06/29�쐬�@�쑺
Private Function GetTBCME039(col As c_cmzcHinPlans, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long
Dim target As c_cmzcHinPlan

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetTBCME039"

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, INGOTPOS, HINBAN, REVNUM, FACT, OPCOND, LENGTH, USECLASS, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME039"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
        sql = sql & sqlWhere
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetTBCME039 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcHinPlan
        With target
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .INGOTPOS = rs("INGOTPOS")       ' �������J�n�ʒu
            .hinban = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' �����ԍ�
            .FACT = rs("FACT")               ' �H��
            .OPCOND = rs("OPCOND")           ' ���Ə���
            .Length = rs("LENGTH")           ' ����
            .USECLASS = rs("USECLASS")       ' �g�p�敪
        End With
        col.Add target
        Set target = Nothing
        rs.MoveNext
    Next
    rs.Close

    GetTBCME039 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�e�[�u���uTBCME040�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME040 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/06/29�쐬�@�쑺
Private Function GetTBCME040(col As c_cmzcBlks, Optional sqlWhere$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long
Dim target As c_cmzcBlk

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetTBCME040"

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, REALLEN, BLOCKID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, DELCLS, LSTATCLS," & _
              " RSTATCLS, HOLDCLS, BDCAUS, REGDATE, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME040"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
        sql = sql & sqlWhere & " and (LENGTH>0)"
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetTBCME040 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcBlk
        With target
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .INGOTPOS = rs("INGOTPOS")       ' �������J�n�ʒu
            .Length = rs("LENGTH")           ' ����
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
        End With
        col.Add target
        rs.MoveNext
    Next
    rs.Close

    GetTBCME040 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�e�[�u���uXSDC2�v�uXSDCS�v�uXSDC4�v����u���b�N�����擾����
'���Ұ�    :�ϐ���        ,IO  ,�^                ,����
'          :col           ,O   ,c_cmzcBlks        ,���o���R�[�h
'          :CRYNUM        ,I   ,String            ,�����ԍ�
'          :�߂�l        ,O   ,FUNCTION_RETURN   ,���o�̐���
'����      :
'����      :2005/10/05 ooba
Private Function GetBlockData(col As c_cmzcBlks, Optional CRYNUM$) As FUNCTION_RETURN

    Dim sql As String       'SQL�S��
    Dim sql2 As String      'SQL�S��
    Dim rs As OraDynaset    'RecordSet
    Dim rs2 As OraDynaset   'RecordSet
    Dim recCnt As Long      '���R�[�h��
    Dim i As Long
    Dim target As c_cmzcBlk


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetBlockData"

    ''SQL��g�ݗ��Ă�
    sql = "select "
    sql = sql & "CSTOP.XTALCS, "                                            '�����ԍ�
    sql = sql & "CSTOP.INPOSCS, "                                           '�������J�n�ʒu
    sql = sql & "CSBOT.INPOSCS - CSTOP.INPOSCS as LENGTH, "                 '����
    sql = sql & "nvl(GNLC2,CSBOT.INPOSCS - CSTOP.INPOSCS) as REALLEN, "     '������
    sql = sql & "CSTOP.CRYNUMCS, "                                          '�u���b�NID
    sql = sql & "nvl(GNKKNTC2,' ') as KRPROCCD, "                           '���݊Ǘ��H��
    sql = sql & "nvl(GNWKNTC2,' ') as NOWPROC, "                            '���ݍH��
    sql = sql & "nvl(NEKKNTC2,' ') as LPKRPROCCD, "                         '�ŏI�ʉߊǗ��H��
    sql = sql & "nvl(NEWKNTC2,' ') as LASTPASS, "                           '�ŏI�ʉߍH��
    sql = sql & "nvl(SAKJC2,'0') as DELCLS, "                               '�폜�敪
    sql = sql & "nvl(LSTATBC2,'T') as LSTATCLS, "                           '�ŏI��ԋ敪
    sql = sql & "nvl(RSTATBC2,'T') as RSTATCLS, "                           '������ԋ敪
    sql = sql & "nvl(HOLDBC2,'0') as HOLDCLS, "                             '�z�[���h�敪
    sql = sql & "BDCAUSC2 as BDCAUS, "                                      '�s�Ǘ��R
    sql = sql & "C4.KNKTC4, "                                               '�ŏI�ʉߊǗ��H��(XSDC4)
    sql = sql & "C4.WKKTC4, "                                               '�ŏI�ʉߍH��(XSDC4)
    sql = sql & "C4.FCODEC4 "                                               '�s�Ǘ��R(XSDC4)
    sql = sql & "from XSDC2, "
    sql = sql & "     (select "
    sql = sql & "      CRYNUMCS, "
    sql = sql & "      XTALCS, "
    sql = sql & "      INPOSCS "
    sql = sql & "      from XSDCS "
    sql = sql & "      where "
    sql = sql & "      TBKBNCS = 'T' "
    sql = sql & "      and substr(CRYNUMCS,10,3) not in ('TOP','BOT') "
    sql = sql & "      and XTALCS = '" & CRYNUM & "' "
    sql = sql & "     ) CSTOP, "
    sql = sql & "     (select "
    sql = sql & "      CRYNUMCS, "
    sql = sql & "      XTALCS, "
    sql = sql & "      INPOSCS "
    sql = sql & "      from XSDCS "
    sql = sql & "      where "
    sql = sql & "      TBKBNCS = 'B' "
    sql = sql & "      and substr(CRYNUMCS,10,3) not in ('TOP','BOT') "
    sql = sql & "      and XTALCS = '" & CRYNUM & "' "
    sql = sql & "     ) CSBOT, "
    sql = sql & "     (select "
    sql = sql & "      XTALC4, "
    sql = sql & "      INPOSC4, "
    sql = sql & "      KNKTC4, "
    sql = sql & "      WKKTC4, "
    sql = sql & "      FCODEC4 "
    sql = sql & "      from XSDC4 TMP4 "
    sql = sql & "      where "
    sql = sql & "      XTALC4 like '" & Mid(CRYNUM, 1, 9) & "%' "
    sql = sql & "      and (KCKNTC4, KDAYC4) = ("
    sql = sql & "                     select MAX(KCKNTC4), MAX(KDAYC4) "
    sql = sql & "                     from XSDC4 "
    sql = sql & "                     where XTALC4 like '" & Mid(CRYNUM, 1, 9) & "%' "
    sql = sql & "                     and INPOSC4 = TMP4.INPOSC4) "
    sql = sql & "     ) C4 "
    sql = sql & "where "
    sql = sql & "CSTOP.CRYNUMCS = CRYNUMC2(+) "
    sql = sql & "and CSTOP.CRYNUMCS = CSBOT.CRYNUMCS "
    sql = sql & "and CSTOP.INPOSCS = C4.INPOSC4(+) "
    sql = sql & "and (LIVKC2 is null or LIVKC2 = '0' "
    sql = sql & "     or LSTATBC2 in ('R', 'H', 'B') or KANKC2 = '2') "
    sql = sql & "order by CSTOP.INPOSCS "
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetBlockData = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcBlk
        With target
            .CRYNUM = rs("XTALCS")              ' �����ԍ�
            .INGOTPOS = rs("INPOSCS")           ' �������J�n�ʒu
            .Length = rs("LENGTH")              ' ����
            .REALLEN = rs("REALLEN")            ' ������
            .BLOCKID = rs("CRYNUMCS")           ' �u���b�NID
            .KRPROCCD = rs("KRPROCCD")          ' ���݊Ǘ��H��
            .NOWPROC = rs("NOWPROC")            ' ���ݍH��
            .LPKRPROCCD = rs("LPKRPROCCD")      ' �ŏI�ʉߊǗ��H��
            .LASTPASS = rs("LASTPASS")          ' �ŏI�ʉߍH��
            .DELCLS = rs("DELCLS")              ' �폜�敪
            .LSTATCLS = rs("LSTATCLS")          ' �ŏI��ԋ敪
            .RSTATCLS = rs("RSTATCLS")          ' ������ԋ敪
            .HOLDCLS = rs("HOLDCLS")            ' �z�[���h�敪
            If InStr(.BLOCKID, "$") <> 0 Then
                .KRPROCCD = MGPRCD_RIMERUTO_UKEIRE          ' ���݊Ǘ��H��
                .NOWPROC = PROCD_RIMERUTO_UKEIRE            ' ���ݍH��
                .RSTATCLS = "M"                             ' ������ԋ敪
                ' �ŏI�ʉߊǗ��H��
                If IsNull(rs("KNKTC4")) Then .LPKRPROCCD = "" Else .LPKRPROCCD = rs("KNKTC4")
                ' �ŏI�ʉߍH��
                If IsNull(rs("WKKTC4")) Then .LASTPASS = "" Else .LASTPASS = rs("WKKTC4")
                ' �s�Ǘ��R
                If IsNull(rs("FCODEC4")) Then .BDCAUS = "0" Else .BDCAUS = rs("FCODEC4")
            Else
                ' �s�Ǘ��R
                If IsNull(rs("BDCAUS")) Then .BDCAUS = "0" Else .BDCAUS = rs("BDCAUS")
            End If
            If Trim(.NOWPROC) = "" Then .DELCLS = "1"
            
'''            '�s�Ǘ��R��XSDC4����擾
'''            If InStr(.BLOCKID, "$") <> 0 Then
'''                .KRPROCCD = MGPRCD_RIMERUTO_UKEIRE      ' ���݊Ǘ��H��
'''                .NOWPROC = PROCD_RIMERUTO_UKEIRE        ' ���ݍH��
'''                .RSTATCLS = "M"                         ' ������ԋ敪
'''
'''                sql2 = "select KNKTC4, WKKTC4, FCODEC4 from XSDC4 "
''''                sql2 = sql2 & "where substr(XTALC4, 1 ,10) = '" & Mid(.BLOCKID, 1, 10) & "' "
'''                sql2 = sql2 & "where XTALC4 like '" & Mid(.BLOCKID, 1, 10) & "%' "  '05/12/26
'''                sql2 = sql2 & "and INPOSC4 = " & .INGOTPOS & " "
'''                sql2 = sql2 & "order by KCKNTC4 desc "
'''
'''                Set rs2 = OraDB.DBCreateDynaset(sql2, ORADYN_NO_BLANKSTRIP)
'''
'''                If rs2 Is Nothing Then
'''                    GetBlockData = FUNCTION_RETURN_FAILURE
'''                    GoTo proc_exit
'''                End If
'''
'''                If rs2.RecordCount > 0 Then
'''                    ' �ŏI�ʉߊǗ��H��
'''                    If IsNull(rs2("KNKTC4")) Then .LPKRPROCCD = "" Else .LPKRPROCCD = rs2("KNKTC4")
'''                    ' �ŏI�ʉߍH��
'''                    If IsNull(rs2("WKKTC4")) Then .LASTPASS = "" Else .LASTPASS = rs2("WKKTC4")
'''                    ' �s�Ǘ��R
'''                    If IsNull(rs2("FCODEC4")) Then .BDCAUS = "0" Else .BDCAUS = rs2("FCODEC4")
'''                Else
'''                    .BDCAUS = "0"                       ' �s�Ǘ��R
'''                End If
'''                rs2.Close
'''            Else
'''                ' �s�Ǘ��R
'''                If IsNull(rs("BDCAUS")) Then .BDCAUS = "0" Else .BDCAUS = rs("BDCAUS")
'''            End If
        End With
        col.Add target
        rs.MoveNext
    Next
    rs.Close

    GetBlockData = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�e�[�u���uTBCME041�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME041 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/06/29�쐬�@�쑺
Private Function GetTBCME041(col As c_cmzcHins, Optional sqlWhere$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long
Dim target As c_cmzcHin

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetTBCME041"

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, INGOTPOS, HINBAN, REVNUM, FACTORY, OPECOND, LENGTH, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME041"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
        sql = sql & sqlWhere
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetTBCME041 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcHin
        With target
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .INGOTPOS = rs("INGOTPOS")       ' �������J�n�ʒu
            .hinban = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .Factory = rs("FACTORY")         ' �H��
            .OpeCond = rs("OPECOND")         ' ���Ə���
            .Length = rs("LENGTH")           ' ����
        End With
        col.Add target
        Set target = Nothing
        rs.MoveNext
    Next
    rs.Close

    GetTBCME041 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�e�[�u���uTBCME042�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME042 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/06/29�쐬�@�쑺
Private Function GetTBCME042(col As c_cmzcSxls, Optional sqlWhere$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long
Dim target As c_cmzcSxl

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetTBCME042"

    ''SQL��g�ݗ��Ă�
''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{
    sqlBase = ""
    sqlBase = sqlBase & " SELECT"
    sqlBase = sqlBase & "  a.xtalcb as CRYNUM"        ''�����ԍ�
    sqlBase = sqlBase & " ,a.inposcb as INGOTPOS"     ''�������J�n�ʒu
    sqlBase = sqlBase & " ,a.rlencb as LENGTH"        ''���_����
    sqlBase = sqlBase & " ,a.sxlidcb as SXLID"        ''SXLID
    sqlBase = sqlBase & " ,' ' as KRPROCCD"           ''�Ǘ��H��(���ݸ)
    sqlBase = sqlBase & " ,a.gnwkntcb as NOWPROC"     ''���ݍH��
    sqlBase = sqlBase & " ,' ' as LPKRPROCCD"         ''�ŏI�ʉߊǗ��H��(���ݸ)
    sqlBase = sqlBase & " ,a.newkntcb as LASTPASS"    ''�ŏI�ʉߍH��
    sqlBase = sqlBase & " ,a.livkcb as DELCLS"        ''�����敪
    sqlBase = sqlBase & " ,a.lstccb as LSTATCLS"      ''�ŏI��ԋ敪
    sqlBase = sqlBase & " ,a.sholdclscb as HOLDCLS"   ''�z�[���h�敪
    sqlBase = sqlBase & " ,a.hinbcb as HINBAN"        ''�i��
    sqlBase = sqlBase & " ,a.revnumcb as REVNUM"      ''���i�ԍ������ԍ�
    sqlBase = sqlBase & " ,a.factorycb as FACTORY"    ''�H��
    sqlBase = sqlBase & " ,a.opecb as OPECOND"        ''���Ə���
    sqlBase = sqlBase & " ,a.furyccb as BDCAUS"       ''�s�Ǘ��R
    sqlBase = sqlBase & " ,a.maicb as COUNT"          ''������
    sqlBase = sqlBase & " ,a.tdaycb as REGDATE"       ''�o�^���t
    sqlBase = sqlBase & " ,a.kdaycb as UPDDATE"       ''�X�V���t
    sqlBase = sqlBase & " ,' ' as SUMMITSENDFLAG"     ''SUMMIT���M�t���O(���ݸ)
    sqlBase = sqlBase & " ,a.sndkcb as SENDFLAG"      ''���MFLG
    sqlBase = sqlBase & " ,a.sndaycb as SENDDATE"     ''���M���t
    sqlBase = sqlBase & " FROM"
    sqlBase = sqlBase & "  xsdcb a"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
        sql = sql & sqlWhere
'        sql = sql & " AND a.livkcb = '0'"
'    Else
'        sql = sql & " WHERE a.livkcb = '0'"
    End If
''���ǉ�END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{
''���폜START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{
'    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, SXLID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, DELCLS, LSTATCLS, HOLDCLS," & _
'              " HINBAN, REVNUM, FACTORY, OPECOND, BDCAUS, COUNT, REGDATE, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE "
'    sqlBase = sqlBase & "From TBCME042"
'    sql = sqlBase
'    If (sqlWhere <> vbNullString) Then
'        sql = sql & sqlWhere
'    End If
''���폜END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        GetTBCME042 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcSxl
        With target
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .INGOTPOS = rs("INGOTPOS")       ' �������J�n�ʒu
            If IsNull(rs("LENGTH")) = False Then .Length = rs("LENGTH")         ' ����
            .SXLID = rs("SXLID")             ' SXLID
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H��
            .NOWPROC = rs("NOWPROC")         ' ���ݍH��
            .LPKRPROCCD = rs("LPKRPROCCD")   ' �ŏI�ʉߊǗ��H��
            .LASTPASS = rs("LASTPASS")       ' �ŏI�ʉߍH��
            .DELCLS = rs("DELCLS")           ' �폜�敪
            .LSTATCLS = rs("LSTATCLS")       ' �ŏI��ԋ敪
            If IsNull(rs("HOLDCLS")) = False Then .HOLDCLS = rs("HOLDCLS")      ' �z�[���h�敪
            .hinban = rs("HINBAN")           ' �i��
''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{
            If .LSTATCLS = "H" Then
                .hinban = "Z"
            End If
''���ǉ�END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/19 SMP���{
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .Factory = rs("FACTORY")         ' �H��
            .OpeCond = rs("OPECOND")         ' ���Ə���
            .BDCAUS = rs("BDCAUS")           ' �s�Ǘ��R
            .COUNT = rs("COUNT")             ' ����
        End With
        col.Add target
        Set target = Nothing
        rs.MoveNext
    Next
    rs.Close

    GetTBCME042 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function



'�T�v      :�e�[�u���uXSDCS�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_XSDCS    ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/06/29�쐬�@�쑺
Private Function GetTBCME043(col As c_cmzcXlSmps, Optional sqlWhere$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long
Dim target As c_cmzcXlSmp
Dim j As Integer

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetTBCME043"

    ''SQL��g�ݗ��Ă�
'    sqlBase = "Select CRYNUM, INGOTPOS, SMPKBN, SMPLNO, HINBAN, REVNUM, FACTORY, OPECOND, KTKBN, CRYINDRS, CRYINDOI, CRYINDB1," & _
'              " CRYINDB2, CRYINDB3, CRYINDL1, CRYINDL2, CRYINDL3, CRYINDL4, CRYINDCS, CRYINDGD, CRYINDT, CRYINDEP, CRYRESRS," & _
'              " CRYRESOI, CRYRESB1, CRYRESB2, CRYRESB3, CRYRESL1, CRYRESL2, CRYRESL3, CRYRESL4, CRYRESCS, CRYRESGD, CRYREST," & _
'              " CRYRESEP, SMPLNUM, SMPLPAT, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
'    sqlBase = sqlBase & "From TBCME043"
    sqlBase = "Select CRYNUMCS, SMPKBNCS, TBKBNCS, REPSMPLIDCS, XTALCS, INPOSCS, HINBCS, REVNUMCS, FACTORYCS, OPECS, KTKBNCS, " & _
              " BLKKTFLAGCS, CRYSMPLIDRSCS, CRYSMPLIDRS1CS, CRYSMPLIDRS2CS, CRYINDRSCS, CRYRESRS1CS, CRYRESRS2CS,CRYSMPLIDOICS, " & _
              " CRYINDOICS, CRYRESOICS, CRYSMPLIDB1CS, CRYINDB1CS, CRYRESB1CS, CRYSMPLIDB2, CRYINDB2CS, CRYRESB2CS, CRYSMPLIDB3CS, " & _
              " CRYINDB3CS, CRYRESB3CS, CRYSMPLIDL1CS, CRYINDL1CS, CRYRESL1CS, CRYSMPLIDL2CS,  CRYINDL2CS, CRYRESL2CS, CRYSMPLIDL3CS, " & _
              " CRYINDL3CS, CRYRESL3CS, CRYSMPLIDL4CS, CRYINDL4CS, CRYRESL4CS, CRYSMPLIDCSCS, CRYINDCSCS, CRYRESCSCS, CRYSMPLIDGDCS, " & _
              " CRYINDGDCS, CRYRESGDCS, CRYSMPLIDTCS, CRYINDTCS,  CRYRESTCS, CRYSMPLIDEPCS, CRYINDEPCS,CRYRESEPCS, SMPLNUMCS, " & _
              " SMPLPATCS, TSTAFFCS, TDAYCS, KSTAFFCS, KDAYCS, SNDKCS, SNDDAYCS "
    sqlBase = sqlBase & "From XSDCS"

    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
        sql = sql & sqlWhere
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetTBCME043 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcXlSmp
        With target
            .CRYNUM = rs("XTALCS")           ' �����ԍ�
            .INGOTPOS = rs("INPOSCS")       ' �������ʒu
            .SMPKBN = rs("SMPKBNCS")           ' �T���v���敪
            .SMPLNO = rs("REPSMPLIDCS")           ' �T���v��No
            .hinban = rs("HINBCS")           ' �i��
            .REVNUM = rs("REVNUMCS")           ' ���i�ԍ������ԍ�
            .Factory = rs("FACTORYCS")         ' �H��
            .OpeCond = rs("OPECS")         ' ���Ə���
            .KTKBN = rs("KTKBNCS")             ' �m��敪
            .CRYINDRS = rs("CRYINDRSCS")       ' ���������w���iRs)
            .CRYINDOI = rs("CRYINDOICS")       ' ���������w���iOi)
            .CRYINDB1 = rs("CRYINDB1CS")       ' ���������w���iB1)
            .CRYINDB2 = rs("CRYINDB2CS")       ' ���������w���iB2�j
            .CRYINDB3 = rs("CRYINDB3CS")       ' ���������w���iB3)
            .CRYINDL1 = rs("CRYINDL1CS")       ' ���������w���iL1)
            .CRYINDL2 = rs("CRYINDL2CS")       ' ���������w���iL2)
            .CRYINDL3 = rs("CRYINDL3CS")       ' ���������w���iL3)
            .CRYINDL4 = rs("CRYINDL4CS")       ' ���������w���iL4)
            .CRYINDCS = rs("CRYINDCSCS")       ' ���������w���iCs)
            .CRYINDGD = rs("CRYINDGDCS")       ' ���������w���iGD)
            .CRYINDT = rs("CRYINDTCS")         ' ���������w���iT)
            .CRYINDEP = rs("CRYINDEPCS")       ' ���������w���iEPD)
            .CRYRESRS = rs("CRYRESRSCS")       ' �����������сiRs)
            .CRYRESOI = rs("CRYRESOICS")       ' �����������сiOi)
            .CRYRESB1 = rs("CRYRESB1CS")       ' �����������сiB1)
            .CRYRESB2 = rs("CRYRESB2CS")       ' �����������сiB2�j
            .CRYRESB3 = rs("CRYRESB3CS")       ' �����������сiB3)
            .CRYRESL1 = rs("CRYRESL1CS")       ' �����������сiL1)
            .CRYRESL2 = rs("CRYRESL2CS")       ' �����������сiL2)
            .CRYRESL3 = rs("CRYRESL3CS")       ' �����������сiL3)
            .CRYRESL4 = rs("CRYRESL4CS")       ' �����������сiL4)
            .CRYRESCS = rs("CRYRESCSCS")       ' �����������сiCs)
            .CRYRESGD = rs("CRYRESGDCS")       ' �����������сiGD)
            .CRYREST = rs("CRYRESTCS")         ' �����������сiT)
            .CRYRESEP = rs("CRYRESEPCS")       ' �����������сiEPD)
            .SMPLNUM = rs("SMPLNUMCS")         ' �T���v������
            .SMPLPAT = rs("SMPLPATCS")         ' �T���v���p�^�[��
        End With
        col.Add target
        Set target = Nothing
        rs.MoveNext
    Next
    rs.Close

    GetTBCME043 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�e�[�u���uXSDCW�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_XSDCW    ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/06/29�쐬�@�쑺
Private Function GetTBCME044(col As c_cmzcWfSmps, Optional sqlWhere$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long
Dim target As c_cmzcWfSmp

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetTBCME044"

    ''SQL��g�ݗ��Ă�
'    sqlBase = "Select CRYNUM, INGOTPOS, SMPKBN, SMPLID, HINBAN, REVNUM, FACTORY, OPECOND, KTKBN, WFINDRS, WFINDOI, WFINDB1," & _
'              " WFINDB2, WFINDB3, WFINDL1, WFINDL2, WFINDL3, WFINDL4, WFINDDS, WFINDDZ, WFINDSP, WFINDDO1, WFINDDO2, WFINDDO3," & _
'              " WFRESRS, WFRESOI, WFRESB1, WFRESB2, WFRESB3, WFRESL1, WFRESL2, WFRESL3, WFRESL4, WFRESDS, WFRESDZ, WFRESSP," & _
'              " WFRESDO1, WFRESDO2, WFRESDO3, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
'    sqlBase = sqlBase & "From TBCME044"

    'GD���ڒǉ��@05/01/17 ooba
    '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)kondoh
    sqlBase = "Select SXLIDCW, SMPKBNCW, TBKBNCW, REVNUMCW, XTALCW, INPOSCW, REPSMPLIDCW, HINBCW, FACTORYCW, OPECW, KTKBNCW, " & _
              " SMCRYNUMCW, WFSMPLIDRSCW, WFSMPLIDRS1CW, WFSMPLIDRS2CW, WFINDRSCW, WFRESRS1CW, WFRESRS2CW, WFSMPLIDOICW, WFINDOICW, " & _
              " WFRESOICW, WFSMPLIDB1CW, WFINDB1CW, WFRESB1CW, WFSMPLIDB2CW, WFINDB2CW, WFRESB2CW, WFSMPLIDB3CW, WFINDB3CW, " & _
              " WFRESB3CW, WFSMPLIDL1CW, WFINDL1CW, WFRESL1CW, WFSMPLIDL2CW, WFINDL2CW, WFRESL2CW, WFSMPLIDL3CW, WFINDL3CW, WFRESL3CW," & _
              " WFSMPLIDL4CW, WFINDL4CW, WFRESL4CW, WFSMPLIDDSCW, WFINDDSCW, WFRESDSCW, WFSMPLIDDZCW, WFINDDZCW, WFRESDZCW, " & _
              " WFSMPLIDSPCW, WFINDSPCW, WFRESSPCW, WFSMPLIDDO1CW, WFINDDO1CW, WFRESDO1CW, WFSMPLIDDO2CW, WFINDDO2CW, WFRESDO2CW, " & _
              " WFSMPLIDDO3CW, WFINDDO3CW, WFRESDO3CW, WFSMPLIDAOICW, WFINDAOICW, WFRESAOICW, SMPLNUMCW, SMPLPATCW, TSTAFFCW, TDAYCW, " & _
              " KSTAFFCW, KDAYCW, SNDKCW, SNDDAYCW, WFSMPLIDGDCW, WFINDGDCW, WFRESGDCW, WFHSGDCW " & _
              " ,EPSMPLIDB1CW, EPINDB1CW, EPRESB1CW, EPSMPLIDB2CW, EPINDB2CW, EPRESB2CW, EPSMPLIDB3CW, EPINDB3CW, EPRESB3CW, " & _
              " EPSMPLIDL1CW, EPINDL1CW, EPRESL1CW, EPSMPLIDL2CW, EPINDL2CW, EPRESL2CW, EPSMPLIDL3CW, EPINDL3CW, EPRESL3CW "

    sqlBase = sqlBase & "From XSDCW"

    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
'�m��敪���f�X�f�������珜���@�_�@�����P�T�N�T���R�O��
'       sql = sql & sqlWhere
''      sql = sql & sqlWhere & " and SMPKBN != '9'"
        sql = sql & sqlWhere & " and SMPKBNCW != '9'"
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetTBCME044 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcWfSmp
        With target
            If IsNull(rs("XTALCW")) = False Then .CRYNUM = rs("XTALCW")             ' �����ԍ�
            If IsNull(rs("INPOSCW")) = False Then .INGOTPOS = rs("INPOSCW")         ' �������ʒu
            If IsNull(rs("SMPKBNCW")) = False Then .SMPKBN = rs("SMPKBNCW")         ' �T���v���敪
            If IsNull(rs("REPSMPLIDCW")) = False Then .SMPLID = rs("REPSMPLIDCW")   ' �T���v��ID
            If IsNull(rs("HINBCW")) = False Then .hinban = rs("HINBCW")             ' �i��
            If IsNull(rs("REVNUMCW")) = False Then .REVNUM = rs("REVNUMCW")         ' ���i�ԍ������ԍ�
            If IsNull(rs("FACTORYCW")) = False Then .Factory = rs("FACTORYCW")      ' �H��
            If IsNull(rs("OPECW")) = False Then .OpeCond = rs("OPECW")              ' ���Ə���
            If IsNull(rs("KTKBNCW")) = False Then .KTKBN = rs("KTKBNCW")            ' �m��敪
            If IsNull(rs("WFINDRSCW")) = False Then .WFINDRS = rs("WFINDRSCW")      ' WF�����w���iRs)
            If IsNull(rs("WFINDOICW")) = False Then .WFINDOI = rs("WFINDOICW")      ' WF�����w���iOi)
            If IsNull(rs("WFINDB1CW")) = False Then .WFINDB1 = rs("WFINDB1CW")      ' WF�����w���iB1)
            If IsNull(rs("WFINDB2CW")) = False Then .WFINDB2 = rs("WFINDB2CW")      ' WF�����w���iB2�j
            If IsNull(rs("WFINDB3CW")) = False Then .WFINDB3 = rs("WFINDB3CW")      ' WF�����w���iB3)
            If IsNull(rs("WFINDL1CW")) = False Then .WFINDL1 = rs("WFINDL1CW")      ' WF�����w���iL1)
            If IsNull(rs("WFINDL2CW")) = False Then .WFINDL2 = rs("WFINDL2CW")      ' WF�����w���iL2)
            If IsNull(rs("WFINDL3CW")) = False Then .WFINDL3 = rs("WFINDL3CW")      ' WF�����w���iL3)
            If IsNull(rs("WFINDL4CW")) = False Then .WFINDL4 = rs("WFINDL4CW")      ' WF�����w���iL4)
            If IsNull(rs("WFINDDSCW")) = False Then .WFINDDS = rs("WFINDDSCW")      ' WF�����w���iDS)
            If IsNull(rs("WFINDDZCW")) = False Then .WFINDDZ = rs("WFINDDZCW")      ' WF�����w���iDZ)
            If IsNull(rs("WFINDSPCW")) = False Then .WFINDSP = rs("WFINDSPCW")      ' WF�����w���iSP)
            If IsNull(rs("WFINDDO1CW")) = False Then .WFINDDO1 = rs("WFINDDO1CW")   ' WF�����w���iDO1)
            If IsNull(rs("WFINDDO2CW")) = False Then .WFINDDO2 = rs("WFINDDO2CW")   ' WF�����w���iDO2)
            If IsNull(rs("WFINDDO3CW")) = False Then .WFINDDO3 = rs("WFINDDO3CW")   ' WF�����w���iDO3)
            If IsNull(rs("WFINDAOICW")) = False Then .WFINDAOI = rs("WFINDAOICW")   ' WF�����w�� (AO)�@�ǉ��@03/12/05 ooba
            If IsNull(rs("WFINDGDCW")) = False Then .WFINDGD = rs("WFINDGDCW")      ' WF�����w�� (GD)�@�ǉ��@05/01/17 ooba
            If IsNull(rs("WFRESRS1CW")) = False Then .WFRESRS = rs("WFRESRS1CW")    ' WF�������сiRs)
            If IsNull(rs("WFRESOICW")) = False Then .WFRESOI = rs("WFRESOICW")      ' WF�������сiOi)
            If IsNull(rs("WFRESB1CW")) = False Then .WFRESB1 = rs("WFRESB1CW")      ' WF�������сiB1)
            If IsNull(rs("WFRESB2CW")) = False Then .WFRESB2 = rs("WFRESB2CW")      ' WF�������сiB2�j
            If IsNull(rs("WFRESB3CW")) = False Then .WFRESB3 = rs("WFRESB3CW")      ' WF�������сiB3)
            If IsNull(rs("WFRESL1CW")) = False Then .WFRESL1 = rs("WFRESL1CW")      ' WF�������сiL1)
            If IsNull(rs("WFRESL2CW")) = False Then .WFRESL2 = rs("WFRESL2CW")      ' WF�������сiL2)
            If IsNull(rs("WFRESL3CW")) = False Then .WFRESL3 = rs("WFRESL3CW")      ' WF�������сiL3)
            If IsNull(rs("WFRESL4CW")) = False Then .WFRESL4 = rs("WFRESL4CW")      ' WF�������сiL4)
            If IsNull(rs("WFRESDSCW")) = False Then .WFRESDS = rs("WFRESDSCW")      ' WF�������сiDS)
            If IsNull(rs("WFRESDZCW")) = False Then .WFRESDZ = rs("WFRESDZCW")      ' WF�������сiDZ)
            If IsNull(rs("WFRESSPCW")) = False Then .WFRESSP = rs("WFRESSPCW")      ' WF�������сiSP)
            If IsNull(rs("WFRESDO1CW")) = False Then .WFRESDO1 = rs("WFRESDO1CW")   ' WF�������сiDO1)
            If IsNull(rs("WFRESDO2CW")) = False Then .WFRESDO2 = rs("WFRESDO2CW")   ' WF�������сiDO2)
            If IsNull(rs("WFRESDO3CW")) = False Then .WFRESDO3 = rs("WFRESDO3CW")   ' WF�������сiDO3)
            If IsNull(rs("WFRESAOICW")) = False Then .WFRESAOI = rs("WFRESAOICW")   ' WF�������� (AO)�@�ǉ��@03/12/05 ooba
            If IsNull(rs("WFRESGDCW")) = False Then .WFRESGD = rs("WFRESGDCW")      ' WF�������� (GD)�@�ǉ��@05/01/17 ooba
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
            If IsNull(rs("EPINDB1CW")) = False Then .EPINDB1 = rs("EPINDB1CW")      ' WF�����w�� (BMD1E)
            If IsNull(rs("EPRESB1CW")) = False Then .EPRESB1 = rs("EPRESB1CW")      ' WF�������� (BMD1E)
            If IsNull(rs("EPINDB2CW")) = False Then .EPINDB2 = rs("EPINDB2CW")      ' WF�����w�� (BMD2E)
            If IsNull(rs("EPRESB2CW")) = False Then .EPRESB2 = rs("EPRESB2CW")      ' WF�������� (BMD2E)
            If IsNull(rs("EPINDB3CW")) = False Then .EPINDB3 = rs("EPINDB3CW")      ' WF�����w�� (BMD3E)
            If IsNull(rs("EPRESB3CW")) = False Then .EPRESB3 = rs("EPRESB3CW")      ' WF�������� (BMD3E)
            If IsNull(rs("EPINDL1CW")) = False Then .EPINDL1 = rs("EPINDL1CW")      ' WF�����w�� (OSF1E)
            If IsNull(rs("EPRESL1CW")) = False Then .EPRESL1 = rs("EPRESL1CW")      ' WF�������� (OSF1E)
            If IsNull(rs("EPINDL2CW")) = False Then .EPINDL2 = rs("EPINDL2CW")      ' WF�����w�� (OSF2E)
            If IsNull(rs("EPRESL2CW")) = False Then .EPRESL2 = rs("EPRESL2CW")      ' WF�������� (OSF2E)
            If IsNull(rs("EPINDL3CW")) = False Then .EPINDL3 = rs("EPINDL3CW")      ' WF�����w�� (OSF3E)
            If IsNull(rs("EPRESL3CW")) = False Then .EPRESL3 = rs("EPRESL3CW")      ' WF�������� (OSF3E)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
        End With
        col.Add target
        Set target = Nothing
        rs.MoveNext
    Next
    rs.Close

    GetTBCME044 = FUNCTION_RETURN_SUCCESS

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


'�T�v      :�e�[�u���uTBCME045�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME045 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/06/29�쐬�@�쑺
Private Function GetTBCME045(col As c_cmzcCuts, CRYNUM$) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long
Dim target As c_cmzcCut

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetTBCME045"

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, INGOTPOS, TRANCNT, LENGTH, KRPROCCD, PROCCODE, HINBAN, REVNUM, FACTORY, OPECOND, BDCAUS," & _
              " STATCLS, BLOCKID, CRYINDRS, CRYINDOI, CRYINDB1, CRYINDB2, CRYINDB3, CRYINDL1, CRYINDL2, CRYINDL3, CRYINDL4," & _
              " CRYINDCS, CRYINDGD, CRYINDT, CRYINDEP, PRIORITY "
    sqlBase = sqlBase & "From TBCME045 IT " & _
              "Where (CRYNUM='" & CRYNUM & "') And (STATCLS<>'1') " & _
              "  and (TRANCNT=(select MAX(TRANCNT) from TBCME045 where (CRYNUM='" & CRYNUM & "') and (STATCLS<>'1')))"
    sql = sqlBase

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetTBCME045 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcCut
        With target
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .INGOTPOS = rs("INGOTPOS")       ' �������J�n�ʒu
            .TRANCNT = rs("TRANCNT")         ' ������
            .Length = rs("LENGTH")           ' ����
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCODE = rs("PROCCODE")       ' �H���R�[�h
            .hinban = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' �i�Ԑ��i�ԍ������ԍ�
            .Factory = rs("FACTORY")         ' �i�ԍH��
            .OpeCond = rs("OPECOND")         ' �i�ԑ��Ə���
            .BDCAUS = rs("BDCAUS")           ' �敪�R�[�h
            .STATCLS = rs("STATCLS")         ' ��ԋ敪
            .BLOCKID = rs("BLOCKID")         ' �u���b�NID
            .CRYINDRS = rs("CRYINDRS")       ' ���������w���iRs)
            .CRYINDOI = rs("CRYINDOI")       ' ���������w���iOi)
            .CRYINDB1 = rs("CRYINDB1")       ' ���������w���iB1)
            .CRYINDB2 = rs("CRYINDB2")       ' ���������w���iB2�j
            .CRYINDB3 = rs("CRYINDB3")       ' ���������w���iB3)
            .CRYINDL1 = rs("CRYINDL1")       ' ���������w���iL1)
            .CRYINDL2 = rs("CRYINDL2")       ' ���������w���iL2)
            .CRYINDL3 = rs("CRYINDL3")       ' ���������w���iL3)
            .CRYINDL4 = rs("CRYINDL4")       ' ���������w���iL4)
            .CRYINDCS = rs("CRYINDCS")       ' ���������w���iCs)
            .CRYINDGD = rs("CRYINDGD")       ' ���������w���iGD)
            .CRYINDT = rs("CRYINDT")         ' ���������w���iT)
            .CRYINDEP = rs("CRYINDEP")       ' ���������w���iEPD)
            .PRIORITY = rs("PRIORITY")       ' �D��x
        End With
        col.Add target
        Set target = Nothing
        rs.MoveNext
    Next
    rs.Close

    GetTBCME045 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

Private Function GetReject(col As c_cmzcRejs, CRYNUM$) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long
Dim target As c_cmzcRej

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetReject"
    
    GetReject = FUNCTION_RETURN_FAILURE
    
    ''SQL��g�ݗ��Ă�
''    sql = "select LOTID, ALLSCRAP, REJFROM, REJTO " & _
          "from VECMW004 " & _
          "where (LOTID like '" & Left$(CRYNUM, 9) & "%') and (REJCAT<>'C') " & _
          "order by LOTID, REJFROM"

    '�ޭ��Q�ƒ�~�@06/02/06 ooba START ====================================================>
    sql = "select LOTID, ALLSCRAP, REJFROM, REJTO from "
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
    sql = sql & " and a.lotid like '" & left$(CRYNUM, 9) & "%'"
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
    sql = sql & " and lotid like '" & left$(CRYNUM, 9) & "%'"
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
    sql = sql & " and a.lotid like '" & left$(CRYNUM, 9) & "%'"
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
    sql = sql & " and a.lotid like '" & left$(CRYNUM, 9) & "%'"
    sql = sql & " order by LOTID,REJFROM"
    sql = sql & ") XXX,"
    sql = sql & "  (select LOTID, min(TOP_POS)/10.0 as WFFROM, max(TOP_POS)/10.0 as WFTO from TBCMY011 "
    sql = sql & " where lotid like '" & left$(CRYNUM, 9) & "%'"
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
    sql = sql & " and a.lotid like '" & left$(CRYNUM, 9) & "%'"
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
    sql = sql & " and lotid like '" & left$(CRYNUM, 9) & "%'"
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
    sql = sql & " and a.lotid like '" & left$(CRYNUM, 9) & "%'"
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
    sql = sql & " and a.lotid like '" & left$(CRYNUM, 9) & "%'"
    sql = sql & " order by LOTID,REJFROM"
    sql = sql & ") XXX,"
    sql = sql & "  TBCME040 C"
    sql = sql & " where (XXX.LOTID=C.BLOCKID)"
    sql = sql & "  and (XXX.ALLSCRAP='Y')"
    sql = sql & ")"
    sql = sql & " where (REJCAT<>'C')"
    sql = sql & " order by LOTID, REJFROM "
    '�ޭ��Q�ƒ�~�@06/02/06 ooba END ======================================================>
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcRej
        With target
            .LOTID = rs("LOTID")                  ' �u���b�NID
            .ALLSCRAP = rs("ALLSCRAP")            ' �S���X�N���b�v
            .LENFROM = rs("REJFROM")              ' �����@FROM
            .LENTO = rs("REJTO")                  ' �����@TO
        End With
        col.Add target
        Set target = Nothing
        rs.MoveNext
    Next
    rs.Close
    
    GetReject = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

''���ǉ� START SPT�p���э쐬���@�ύX 2006/05/12 SMP-OKAMOTO
'�T�v      :�e�[�u���uXSDC2�v�uXSDCS�v�uXSDC4�v����u���b�N�����擾����
'���Ұ�    :�ϐ���        ,IO  ,�^                ,����
'          :col           ,O   ,c_cmzcBlks        ,���o���R�[�h
'          :CRYNUM        ,I   ,String            ,�����ԍ�
'          :�߂�l        ,O   ,FUNCTION_RETURN   ,���o�̐���
'����      :
'����      :2005/10/05 ooba
Private Function GetBlockData_2(col As c_cmzcBlks, Optional CRYNUM$) As FUNCTION_RETURN

    Dim sql As String       'SQL�S��
    Dim sql2 As String      'SQL�S��
    Dim rs As OraDynaset    'RecordSet
    Dim rs2 As OraDynaset   'RecordSet
    Dim recCnt As Long      '���R�[�h��
    Dim i As Long
    Dim target As c_cmzcBlk

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcGetXl.bas -- Function GetBlockData_2"
    
    ''SQL��g�ݗ��Ă�
    ''�܂��͐ؒf�w���Ȃ��̃u���b�N���擾
    sql = "select "
    sql = sql & "CSTOP.XTALCS, "                                            '�����ԍ�
    sql = sql & "CSTOP.INPOSCS�@INPOSCS, "                                  '�������J�n�ʒu
    sql = sql & "CSBOT.INPOSCS - CSTOP.INPOSCS as LENGTH, "                 '����
    sql = sql & "nvl(GNLC2,CSBOT.INPOSCS - CSTOP.INPOSCS) as REALLEN, "     '������
    sql = sql & "CSTOP.CRYNUMCS, "                                          '�u���b�NID
    sql = sql & "nvl(GNKKNTC2,' ') as KRPROCCD, "                           '���݊Ǘ��H��
    sql = sql & "nvl(GNWKNTC2,' ') as NOWPROC, "                            '���ݍH��
    sql = sql & "nvl(NEKKNTC2,' ') as LPKRPROCCD, "                         '�ŏI�ʉߊǗ��H��
    sql = sql & "nvl(NEWKNTC2,' ') as LASTPASS, "                           '�ŏI�ʉߍH��
    sql = sql & "nvl(SAKJC2,'0') as DELCLS, "                               '�폜�敪
    sql = sql & "nvl(LSTATBC2,'T') as LSTATCLS, "                           '�ŏI��ԋ敪
    sql = sql & "nvl(RSTATBC2,'T') as RSTATCLS, "                           '������ԋ敪
    sql = sql & "nvl(HOLDBC2,'0') as HOLDCLS, "                             '�z�[���h�敪
    sql = sql & "BDCAUSC2 as BDCAUS, "                                      '�s�Ǘ��R
    sql = sql & "C4.KNKTC4, "                                               '�ŏI�ʉߊǗ��H��(XSDC4)
    sql = sql & "C4.WKKTC4, "                                               '�ŏI�ʉߍH��(XSDC4)
    sql = sql & "C4.FCODEC4 "                                               '�s�Ǘ��R(XSDC4)
    sql = sql & "from XSDC2, "
    sql = sql & "     (select "
    sql = sql & "      CRYNUMCS, "
    sql = sql & "      XTALCS, "
    sql = sql & "      INPOSCS, "
    sql = sql & "      CUTFLGCS "
    sql = sql & "      from XSDCS "
    sql = sql & "      where "
    sql = sql & "      TBKBNCS = 'T' "
    sql = sql & "      and substr(CRYNUMCS,10,3) not in ('TOP','BOT') "
    sql = sql & "      and XTALCS = '" & CRYNUM & "' "
    sql = sql & "     ) CSTOP, "
    sql = sql & "     (select "
    sql = sql & "      CRYNUMCS, "
    sql = sql & "      XTALCS, "
    sql = sql & "      INPOSCS, "
    sql = sql & "      CUTFLGCS "
    sql = sql & "      from XSDCS "
    sql = sql & "      where "
    sql = sql & "      TBKBNCS = 'B' "
    sql = sql & "      and substr(CRYNUMCS,10,3) not in ('TOP','BOT') "
    sql = sql & "      and XTALCS = '" & CRYNUM & "' "
    sql = sql & "     ) CSBOT, "
    sql = sql & "     (select "
    sql = sql & "      XTALC4, "
    sql = sql & "      INPOSC4, "
    sql = sql & "      KNKTC4, "
    sql = sql & "      WKKTC4, "
    sql = sql & "      FCODEC4 "
    sql = sql & "      from XSDC4 TMP4 "
    sql = sql & "      where "
    sql = sql & "      XTALC4 like '" & Mid(CRYNUM, 1, 9) & "%' "
    sql = sql & "      and (KCKNTC4, KDAYC4) = ("
    sql = sql & "                     select MAX(KCKNTC4), MAX(KDAYC4) "
    sql = sql & "                     from XSDC4 "
    sql = sql & "                     where XTALC4 like '" & Mid(CRYNUM, 1, 9) & "%' "
    sql = sql & "                     and INPOSC4 = TMP4.INPOSC4) "
    sql = sql & "     ) C4 "
    sql = sql & "where "
    sql = sql & "CSTOP.CRYNUMCS = CRYNUMC2(+) "
    sql = sql & "and CSTOP.CUTFLGCS is null "
    sql = sql & "and CSTOP.CRYNUMCS = CSBOT.CRYNUMCS "
    sql = sql & "and CSTOP.INPOSCS = C4.INPOSC4(+) "
    sql = sql & "and (LIVKC2 is null or LIVKC2 = '0' "
    sql = sql & "     or LSTATBC2 in ('R', 'H', 'B') or KANKC2 = '2') "
    sql = sql & " UNION ("
    ''���ɐؒf�w������̃u���b�N���擾
    sql = sql & "select "
    sql = sql & "CSTOP.XTALCS, "                                            '�����ԍ�
    sql = sql & "CSTOP.INPOSCS�@INPOSCS, "                                  '�������J�n�ʒu
    sql = sql & "CSBOT.INPOSCS - CSTOP.INPOSCS as LENGTH, "                 '����
    sql = sql & "nvl(GNLC2,CSBOT.INPOSCS - CSTOP.INPOSCS) as REALLEN, "     '������
    sql = sql & "CSTOP.CRYNUMCS, "                                          '�u���b�NID
    sql = sql & "nvl(GNKKNTC2,' ') as KRPROCCD, "                           '���݊Ǘ��H��
    sql = sql & "nvl(GNWKNTC2,' ') as NOWPROC, "                            '���ݍH��
    sql = sql & "nvl(NEKKNTC2,' ') as LPKRPROCCD, "                         '�ŏI�ʉߊǗ��H��
    sql = sql & "nvl(NEWKNTC2,' ') as LASTPASS, "                           '�ŏI�ʉߍH��
    sql = sql & "nvl(SAKJC2,'0') as DELCLS, "                               '�폜�敪
    sql = sql & "nvl(LSTATBC2,'T') as LSTATCLS, "                           '�ŏI��ԋ敪
    sql = sql & "nvl(RSTATBC2,'T') as RSTATCLS, "                           '������ԋ敪
    sql = sql & "nvl(HOLDBC2,'0') as HOLDCLS, "                             '�z�[���h�敪
    sql = sql & "BDCAUSC2 as BDCAUS, "                                      '�s�Ǘ��R
    sql = sql & "C4.KNKTC4, "                                               '�ŏI�ʉߊǗ��H��(XSDC4)
    sql = sql & "C4.WKKTC4, "                                               '�ŏI�ʉߍH��(XSDC4)
    sql = sql & "C4.FCODEC4 "                                               '�s�Ǘ��R(XSDC4)
    sql = sql & "from XSDC2, "
    sql = sql & "     (select "
    sql = sql & "      CRYNUMCS, "
    sql = sql & "      XTALCS, "
    sql = sql & "      INPOSCS, "
    sql = sql & "      RPCRYNUMCS, "
    sql = sql & "      CUTFLGCS "
    sql = sql & "      from XSDCS "
    sql = sql & "      where "
    sql = sql & "      TBKBNCS = 'T' "
    sql = sql & "      and substr(CRYNUMCS,10,3) not in ('TOP','BOT') "
    sql = sql & "      and XTALCS = '" & CRYNUM & "' "
    sql = sql & "     ) CSTOP, "
    sql = sql & "     (select "
    sql = sql & "      CRYNUMCS, "
    sql = sql & "      XTALCS, "
    sql = sql & "      INPOSCS, "
    sql = sql & "      RPCRYNUMCS, "
    sql = sql & "      CUTFLGCS "
    sql = sql & "      from XSDCS "
    sql = sql & "      where "
    sql = sql & "      TBKBNCS = 'B' "
    sql = sql & "      and substr(CRYNUMCS,10,3) not in ('TOP','BOT') "
    sql = sql & "      and XTALCS = '" & CRYNUM & "' "
    sql = sql & "     ) CSBOT, "
    sql = sql & "     (select "
    sql = sql & "      XTALC4, "
    sql = sql & "      INPOSC4, "
    sql = sql & "      KNKTC4, "
    sql = sql & "      WKKTC4, "
    sql = sql & "      FCODEC4 "
    sql = sql & "      from XSDC4 TMP4 "
    sql = sql & "      where "
    sql = sql & "      XTALC4 like '" & Mid(CRYNUM, 1, 9) & "%' "
    sql = sql & "      and (KCKNTC4, KDAYC4) = ("
    sql = sql & "                     select MAX(KCKNTC4), MAX(KDAYC4) "
    sql = sql & "                     from XSDC4 "
    sql = sql & "                     where XTALC4 like '" & Mid(CRYNUM, 1, 9) & "%' "
    sql = sql & "                     and INPOSC4 = TMP4.INPOSC4) "
    sql = sql & "     ) C4 "
    sql = sql & "where "
    sql = sql & "CSTOP.RPCRYNUMCS = CRYNUMC2(+) "
    sql = sql & "and CSTOP.CUTFLGCS = '1' "
    sql = sql & "and CSTOP.CRYNUMCS = CSBOT.CRYNUMCS "
    sql = sql & "and CSTOP.INPOSCS = C4.INPOSC4(+) "
    sql = sql & "and (LIVKC2 is null or LIVKC2 = '0' "
    sql = sql & "     or LSTATBC2 in ('R', 'H', 'B') or KANKC2 = '2') "
    sql = sql & " ) "
    sql = sql & "order by INPOSCS "
    
    
    
''���폜 START SPT�p���э쐬���@�ύX IT��Q 2006/06/14 SMP-OKAMOTO
'    sql = "select "
'    sql = sql & "DISTINCT "
'    sql = sql & "CSTOP.XTALCS, "                                            '�����ԍ�
'    sql = sql & "CSTOP.INPOSCS, "                                           '�������J�n�ʒu
'    sql = sql & "CSBOT.INPOSCS - CSTOP.INPOSCS as LENGTH, "                 '����
'    sql = sql & "nvl(GNLC2,CSBOT.INPOSCS - CSTOP.INPOSCS) as REALLEN, "     '������
'    sql = sql & "CSTOP.CRYNUMCS, "                                          '�u���b�NID
'    sql = sql & "nvl(GNKKNTC2,' ') as KRPROCCD, "                           '���݊Ǘ��H��
'    sql = sql & "nvl(GNWKNTC2,' ') as NOWPROC, "                            '���ݍH��
'    sql = sql & "nvl(NEKKNTC2,' ') as LPKRPROCCD, "                         '�ŏI�ʉߊǗ��H��
'    sql = sql & "nvl(NEWKNTC2,' ') as LASTPASS, "                           '�ŏI�ʉߍH��
'    sql = sql & "nvl(SAKJC2,'0') as DELCLS, "                               '�폜�敪
'    sql = sql & "nvl(LSTATBC2,'T') as LSTATCLS, "                           '�ŏI��ԋ敪
'    sql = sql & "nvl(RSTATBC2,'T') as RSTATCLS, "                           '������ԋ敪
'    sql = sql & "nvl(HOLDBC2,'0') as HOLDCLS, "                             '�z�[���h�敪
'    sql = sql & "BDCAUSC2 as BDCAUS, "                                      '�s�Ǘ��R
'    sql = sql & "C4.KNKTC4, "                                               '�ŏI�ʉߊǗ��H��(XSDC4)
'    sql = sql & "C4.WKKTC4, "                                               '�ŏI�ʉߍH��(XSDC4)
'    sql = sql & "C4.FCODEC4 "                                               '�s�Ǘ��R(XSDC4)
'    sql = sql & "from XSDC2, "
'    sql = sql & "     (select "
'    sql = sql & "      CRYNUMCS, "
'    sql = sql & "      RPCRYNUMCS, " 'add
'    sql = sql & "      XTALCS, "
'    sql = sql & "      INPOSCS "
'    sql = sql & "      from XSDCS "
'    sql = sql & "      where "
'    sql = sql & "      TBKBNCS = 'T' "
'    sql = sql & "      and substr(CRYNUMCS,10,3) not in ('TOP','BOT') "
'    sql = sql & "      and XTALCS = '" & CRYNUM & "' "
'    sql = sql & "     ) CSTOP, "
'    sql = sql & "     (select "
'    sql = sql & "      CRYNUMCS, "
'    sql = sql & "      RPCRYNUMCS, " 'add
'    sql = sql & "      XTALCS, "
'    sql = sql & "      INPOSCS "
'    sql = sql & "      from XSDCS "
'    sql = sql & "      where "
'    sql = sql & "      TBKBNCS = 'B' "
'    sql = sql & "      and substr(CRYNUMCS,10,3) not in ('TOP','BOT') "
'    sql = sql & "      and XTALCS = '" & CRYNUM & "' "
'    sql = sql & "     ) CSBOT, "
'    sql = sql & "     (select "
'    sql = sql & "      XTALC4, "
'    sql = sql & "      INPOSC4, "
'    sql = sql & "      KNKTC4, "
'    sql = sql & "      WKKTC4, "
'    sql = sql & "      FCODEC4 "
'    sql = sql & "      from XSDC4 TMP4 "
'    sql = sql & "      where "
'    sql = sql & "      XTALC4 like '" & Mid(CRYNUM, 1, 9) & "%' "
'    sql = sql & "      and (KCKNTC4, KDAYC4) = ("
'    sql = sql & "                     select MAX(KCKNTC4), MAX(KDAYC4) "
'    sql = sql & "                     from XSDC4 "
'    sql = sql & "                     where XTALC4 like '" & Mid(CRYNUM, 1, 9) & "%' "
'    sql = sql & "                     and INPOSC4 = TMP4.INPOSC4) "
'    sql = sql & "     ) C4 "
'    sql = sql & " ,XSDCZ "
'    sql = sql & "where "
'    sql = sql & "CSTOP.CRYNUMCS = CRYNUMCZ(+) "
'    sql = sql & "and RPCRYNUMCZ = CRYNUMC2 "
''    sql = sql & "and CSTOP.CRYNUMCS = CRYNUMC2(+) "
'    sql = sql & "and CSTOP.CRYNUMCS = CSBOT.CRYNUMCS "
'    sql = sql & "and CSTOP.INPOSCS = C4.INPOSC4(+) "
'    sql = sql & "and (LIVKC2 is null or LIVKC2 = '0' "
'    sql = sql & "     or LSTATBC2 in ('R', 'H', 'B') or KANKC2 = '2') "
'    sql = sql & "order by CSTOP.INPOSCS "
''���폜 END   SPT�p���э쐬���@�ύX IT��Q 2006/06/14 SMP-OKAMOTO
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        GetBlockData_2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    For i = 1 To recCnt
        Set target = New c_cmzcBlk
        With target
            .CRYNUM = rs("XTALCS")              ' �����ԍ�
            .INGOTPOS = rs("INPOSCS")           ' �������J�n�ʒu
            .Length = rs("LENGTH")              ' ����
            .REALLEN = rs("REALLEN")            ' ������
            .BLOCKID = rs("CRYNUMCS")           ' �u���b�NID
            .KRPROCCD = rs("KRPROCCD")          ' ���݊Ǘ��H��
            .NOWPROC = rs("NOWPROC")            ' ���ݍH��
            .LPKRPROCCD = rs("LPKRPROCCD")      ' �ŏI�ʉߊǗ��H��
            .LASTPASS = rs("LASTPASS")          ' �ŏI�ʉߍH��
            .DELCLS = rs("DELCLS")              ' �폜�敪
            .LSTATCLS = rs("LSTATCLS")          ' �ŏI��ԋ敪
            .RSTATCLS = rs("RSTATCLS")          ' ������ԋ敪
            .HOLDCLS = rs("HOLDCLS")            ' �z�[���h�敪
            If InStr(.BLOCKID, "$") <> 0 Then
                .KRPROCCD = MGPRCD_RIMERUTO_UKEIRE          ' ���݊Ǘ��H��
                .NOWPROC = PROCD_RIMERUTO_UKEIRE            ' ���ݍH��
                .RSTATCLS = "M"                             ' ������ԋ敪
                ' �ŏI�ʉߊǗ��H��
                If IsNull(rs("KNKTC4")) Then .LPKRPROCCD = "" Else .LPKRPROCCD = rs("KNKTC4")
                ' �ŏI�ʉߍH��
                If IsNull(rs("WKKTC4")) Then .LASTPASS = "" Else .LASTPASS = rs("WKKTC4")
                ' �s�Ǘ��R
                If IsNull(rs("FCODEC4")) Then .BDCAUS = "0" Else .BDCAUS = rs("FCODEC4")
            Else
                ' �s�Ǘ��R
                If IsNull(rs("BDCAUS")) Then .BDCAUS = "0" Else .BDCAUS = rs("BDCAUS")
            End If
            If Trim(.NOWPROC) = "" Then .DELCLS = "1"
            
        End With
        col.Add target
        rs.MoveNext
    Next
    rs.Close

    GetBlockData_2 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

''���ǉ� END   SPT�p���э쐬���@�ύX 2006/05/12 SMP-OKAMOTO

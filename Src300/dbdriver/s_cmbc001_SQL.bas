Attribute VB_Name = "s_cmbc001_SQL"
' TBCME017 (���i�d�l�Ǘ�)���
Public Type s_cmzcF_cmfc001b_Disp
    '���i�d�l�Ǘ�
    Hinban12 As String * 12         ' �i��
    HMGSTRRNO As String * 9         ' �i�Ǘ��d�l�o�^�˗��ԍ�
    REGDATE As Date                 ' �o�^���t
    SENDDATE As Date                ' ���M���t
    SENDFLAG As String              ' �t���O
    TOUROKU As Date
End Type

#If False Then '---------------- �Q�l
' TBCME017 (���i�d�l�Ǘ�)���
Public Type s_cmzcF_cmfc001a_Disp
    '���i�d�l�Ǘ�
    hinban As String * 8            ' �i��
    MNOREVNO As Integer             ' ���i�ԍ������ԍ�
    FACTORY As String * 1           ' �H��
    OPECOND As String * 1           ' ���Ə���
    HMGSTRRNO As String * 9         ' �i�Ǘ��d�l�o�^�˗��ԍ�
    REGDATE As Date                 ' �o�^���t
End Type
#End If '----------------------- �����܂�

'                                     2001/06/11
'================================================
' DB�A�N�Z�X�֐�
' ��`���e: TBCMB011 (PG-ID�Ǘ�)
' �Q�Ɓ@�@: 060200_�S�e�[�u��
'================================================
#If False Then
'------------------------------------------------
' ���[�U��`�^�̐錾
'------------------------------------------------
Public Type typ_cmbc001d_Disp
    PGID As String * 10             ' PG-ID
    HZPART As String * 4            ' HZ�p�[�c
    HZPTRN As String * 2            ' HZ�p�^�[��
    SPACER As String * 5            ' �X�y�[�T
    UPRING As String * 5            ' �A�b�p�[�����O
    CHARGE As Long                  ' �`���[�W��
    RTBPOS As Integer               ' ���c�{�ʒu
    RTBSIZE As String * 2           ' ���c�{�T�C�Y
    GAP As Integer                  ' �M���b�v
    UPDM As Integer                 ' ���㒼�a
    UPLENGTH As Integer             ' ���㒷�i�S���j
    UPRC As Integer                 ' ����iRC�j
    RFRNEED As String * 1           ' ���t���N�^�v��
    UPSPIN As Double                ' �㎲��]��
    DOWNSPIN As Double              ' ������]��
    ROPRESS As String * 8           ' �F����
    ARUGON As String * 7            ' �A���S����
    AIMOIMIN As Double              ' �˂炢Oi�iMIN)
    AIMOIMAX As Double              ' �˂炢Oi�iMAX)
    HCCLASS As String * 7           ' HC���
    HC As String * 3                ' HC
    AVEUPSPD As Double              ' ���ψ��㑬�x
    UPCNTL As String * 1            ' ���㐧��
    BTMSHAPE As String * 1          ' �{�g���`��
    MAGSTR As Long                  ' ���ꋭ�x
    MAGPOS As Long                  ' ����ʒu
    CONDGRT As String * 10          ' �����ۏؓo�^
    MODEL As String * 4             ' �@��
    UPMETHOD As String * 1          ' ������@
    UPCLASS As String * 2           ' ����敪
    UPNUM As String * 1             ' ����{��
    OPETIME As Long                 ' �^�]����
    WTRCOOL As String * 1           ' ����Ǘv��
    PGID2 As String * 8             ' PG-ID�i��{���j
    RCPT1 As String * 3             ' �Ή����V�sNo�iT1)
    RCPT2 As String * 3             ' �Ή����V�sNo�iT2)
    RCPT3 As String * 3             ' �Ή����V�sNo�iT3)
    RCPT4 As String * 3             ' �Ή����V�sNo�iT4)
    RCPT5 As String * 3             ' �Ή����V�sNo�iT5)
    CNTL1 As String * 1             ' ���䍀�ځi1�j
    CNTL2 As String * 1             ' ���䍀�ځi2�j
    CNTL3 As String * 1             ' ���䍀�ځi3�j
    CNTL4 As String * 1             ' ���䍀�ځi4�j
    CNTL5 As String * 1             ' ���䍀�ځi5�j
    CNTL6 As String * 1             ' ���䍀�ځi6�j
    CNTL7 As String * 1             ' ���䍀�ځi7�j
    CNTL8 As String * 1             ' ���䍀�ځi8�j
    CNTL9 As String * 1             ' ���䍀�ځi9�j
    CNTL10 As String * 1            ' ���䍀�ځi10�j
    CNTL11 As String * 1            ' ���䍀�ځi11�j
    CNTL12 As String * 1            ' ���䍀�ځi12�j
    CNTL13 As String * 1            ' ���䍀�ځi13�j
    CNTL14 As String * 1            ' ���䍀�ځi14�j
    CNTL15 As String * 1            ' ���䍀�ځi15�j
    RUNCOND1 As String              ' �^�]�����P
    RUNCOND2 As String              ' �^�]�����Q
'    TSTAFFID As String * 5          ' �o�^�Ј�ID
'    REGDATE As Date                 ' �o�^���t
'    KSTAFFID As String * 8          ' �X�V�Ј�ID
'    UPDDATE As Date                 ' �X�V���t
'    SENDFLAG As String * 1          ' ���M�t���O
'    SENDDATE As Date                ' ���M���t
End Type
#End If



'�T�v      :����w���ԍ��̘A�ԕ��ɒl��������
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :sijiNo        ,I  ,String    ,���̈���w���ԍ�
'          :addVal        ,I  ,Integer   ,���Z�l(�}�C�i�X����)
'          :�߂�l        ,O  ,String    ,���Z��̈���w���ԍ�
'����      :
'����      :2001/07/09 �쐬  �쑺 (2002/07 s_cmzcF_cmhc001d_SQL.bas���ړ�)
Public Function SijiNoAdd(sijiNo$, addVal%) As String
Dim seq As Integer
Dim newNo As String


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmhc001d_SQL.bas -- Function SijiNoAdd"

    seq = val(Mid$(sijiNo, 5, 3))
    SijiNoAdd = Left$(sijiNo, 4) & Format$(seq + addVal, "000") & Mid$(sijiNo, 8)

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'���݂̎d�l�o�^�ł͖��g�p
Public Function DBDRV_s_cmzcF_cmfc001b_Disp(records() As s_cmzcF_cmfc001b_Disp) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001b_SQL.bas -- Function DBDRV_s_cmzcF_cmfc001b_Disp"
    
    ''���i�d�l�Ǘ���������SXL����������Ȃ����R�[�h�擾�i�i�ԁA�d�l�o�^�˗��ԍ��A�o�^���t�j
    ''�������A��������t�^����ɂ��郌�R�[�h�͏���
    'sql = "select hinban||ltrim(to_char(mnorevno,'00'))||factory||opecond as hinban12, HMGSTRRNO, REGDATE " & _
          "From tbcme018 " & _
          "where (opecond='1') and " & _
          "(hinban||mnorevno||factory) not in (select hinban||mnorevno||factory from tbcme030) and " & _
          "(hinban||mnorevno||factory) not in (select hinban||mnorevno||factory from tbcme031)"
    'sql = "select A.hinban||ltrim(to_char(A.mnorevno,'00'))||A.factory||A.opecond as hinban12, A.HMGSTRRNO, A.REGDATE , B.UPDDATE as TOUROKU " & _
          "From tbcme018 A , tbcme036 B " & _
          "where (A.opecond='1') and (B.opecond='1') and " & _
          "(A.hinban||A.mnorevno||A.factory) = (B.hinban||B.mnorevno||B.factory) and " & _
          "(A.hinban||A.mnorevno||A.factory) not in (select D.hinban||D.mnorevno||D.factory from tbcme030 D) and " & _
          "(A.hinban||A.mnorevno||A.factory) not in (select C.hinban||C.mnorevno||C.factory from tbcme031 C)"
    sql = "select A.hinban||ltrim(to_char(A.mnorevno,'00'))||A.factory||A.opecond as hinban12, A.HMGSTRRNO, A.REGDATE , B.UPDDATE as TOUROKU " & _
          "From tbcme018 A , tbcme036 B " & _
          "where (A.opecond='1')  and (B.opecond(+) = '1') and " & _
          "(A.hinban||A.mnorevno||A.factory) = (B.hinban(+)||B.mnorevno(+)||B.factory(+)) and " & _
          "(A.hinban||A.mnorevno||A.factory) not in (select D.hinban||D.mnorevno||D.factory from tbcme030 D) and " & _
          "(A.hinban||A.mnorevno||A.factory) not in (select C.hinban||C.mnorevno||C.factory from tbcme031 C)"

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_s_cmzcF_cmfc001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .Hinban12 = rs("HINBAN12")        ' �i��
            .HMGSTRRNO = rs("HMGSTRRNO")    ' �i�Ǘ��d�l�o�^�˗��ԍ�
            .REGDATE = rs("REGDATE")        ' �o�^���t
            If IsNull(rs("TOUROKU")) = False Then
                .TOUROKU = rs("TOUROKU")        ' �o�^���t
            Else
                .SENDFLAG = "X"
            End If
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_s_cmzcF_cmfc001b_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'���̉�ʂł�Exec()�͂���Ȃ�
'Public Function DBDRV_s_cmzcF_cmgc001d_Exec(s_cmzcF_cmfc001a_Disp As type_DBDRV_s_cmzcF_cmgc001d_Exec) As FUNCTION_RETURN
'    s_cmzcF_cmgc001c_Exec = FUNCTION_RETURN_SUCCESS
'
'    '�������g��򕥏o���уe�[�u���Ɍ����ԍ�()�A�Ǘ��H���R�[�h()�A�H���R�[�h()�A������d��()�A���X�d�ʁA�Ј��h�c���C���T�[�g
'
'End Function



'�T�v      :SQL�����ɁA�f�[�^���擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :TbName        ,I  ,String    ,�e�[�u����
'          :sql           ,I  ,String    ,SQL
'          :rec           ,O  ,c_cmzcrec ,�擾�f�[�^�i�[��
'          :�߂�l        ,O  ,FUNCTION_RETURN,�����̐���
'����      :
'����      :2001/06/08 �쐬  �쑺
Private Function DispSXL_GetData(TbName$, sql$, rec As c_cmzcrec) As FUNCTION_RETURN
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001c_SQL.bas -- Function DispSXL_GetData"
    
    '' �f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If (rs Is Nothing) Or (rs.RecordCount = 0) Then
        DispSXL_GetData = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    '' ���o���ʂ��i�[����
    Set rec = New c_cmzcrec
    rec.CopyFromRs TbName, rs
    
    rs.Close
    DispSXL_GetData = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'���݂̎d�l�o�^�ł͖��g�p
'�T�v      :���i�d�l���͉�ʗp SXL�^�u�ɕ\��������e�𓾂�
'���Ұ�    :�ϐ���         ,IO ,�^          ,����
'          :targetHinban   ,   ,tFullHinban ,�i�ԏ��
'          :SxlKokyaku_1 ,   ,c_cmzcrec   ,�ڋq�d�lSXL 1 �̓��e
'          :SxlKokyaku_2 ,   ,c_cmzcrec   ,�ڋq�d�lSXL 2 �̓��e
'          :SxlKokyaku_3 ,   ,c_cmzcrec   ,�ڋq�d�lSXL 3 �̓��e
'          :Sxl_1        ,   ,c_cmzcrec   ,���i�d�lSXL_1 �̓��e
'          :Sxl_2        ,   ,c_cmzcrec   ,���i�d�lSXL_2 �̓��e
'          :Sxl_3        ,   ,c_cmzcrec   ,���i�d�lSXL_3 �̓��e
'          :WfKokyaku_2  ,   ,c_cmzcrec   ,�ڋq�d�lWF 2 �̓��e
'          :WfKokyaku_8  ,   ,c_cmzcrec   ,�ڋq�d�lWF 8 �̓��e
'          :SxlUchigawa  ,   ,c_cmzcrec   ,���� �̓��e
'          :�߂�l         ,O  ,FUNCTION_RETURN,�����̐���
'����      :�e�o�̓p�����[�^�̔z��́A(1)�Ɏd�l�f�[�^ (2)�Ɏw�葀�Ə����̃f�[�^ ������
'          :�ďo���ŕi�Ԃ�12�����͉\�Ȃ��߁A�Y���i�Ԃ����݂��Ȃ��ꍇ�����肤��
'����      :2001/06/08 �쐬  �쑺
Public Function DBDRV_s_cmzcF_cmfc001c_DispSXL(targetHinban As tFullHinban, SxlKokyaku_1 As c_cmzcrec, SxlKokyaku_2 As c_cmzcrec, SxlKokyaku_3 As c_cmzcrec, Sxl_1 As c_cmzcrec, Sxl_2 As c_cmzcrec, Sxl_3 As c_cmzcrec, WfKokyaku_2 As c_cmzcrec, WfKokyaku_8 As c_cmzcrec, Sxluchigawa As c_cmzcrec) As FUNCTION_RETURN
Dim sql As String
Dim sqlBase As String       'SQL�̊�{��
Dim sqlWhere As String      'Where��ȍ~
Dim TbName As String        '�e�[�u����
Dim i As Integer
Dim HWFMKnSI(4) As Double   'LPD�T�C�Y(0:���� 1�`4:���)
Dim HWFMKnMX(4) As Integer  'LPD���(0:���� 1�`4:���)
Dim rs As OraDynaset

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001c_SQL.bas -- Function DBDRV_s_cmzcF_cmfc001c_DispSXL"
        
    DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
    
    ''SQL�̋��ʕ�������������
    With targetHinban
        'WHERE��
        sqlWhere = " Where (HINBAN='" & .hinban & "') AND (MNOREVNO=" & .MNOREVNO & ") " & _
              "AND (FACTORY='" & .FACTORY & "') AND (OPECOND='" & .OPECOND & "')"
    End With
    
    ''�o�̓f�[�^������������
    Set SxlKokyaku_1 = New c_cmzcrec
    Set SxlKokyaku_2 = New c_cmzcrec
    Set SxlKokyaku_3 = New c_cmzcrec
    Set Sxl_1 = New c_cmzcrec
    Set Sxl_2 = New c_cmzcrec
    Set Sxl_3 = New c_cmzcrec
    Set WfKokyaku_2 = New c_cmzcrec
    Set WfKokyaku_8 = New c_cmzcrec
    Set Sxluchigawa = New c_cmzcrec
    
    ''�����i�Ԃ̃`�F�b�N�i���i�d�lSXL1�ɓo�^����Ă��邱�ƁA��������t�^����ɓo�^����Ă��Ȃ����Ɓj
    With targetHinban
        sql = "select A.HINBAN from TBCME018 A, TBCME031 B " & _
              "where (A.HINBAN='" & .hinban & "') and (A.MNOREVNO=" & .MNOREVNO & ") and (A.FACTORY='" & .FACTORY & "')" & _
              " and (A.HINBAN=B.HINBAN(+)) and (A.MNOREVNO=B.MNOREVNO(+)) and (A.FACTORY=B.FACTORY(+))" & _
              " and (B.HINBAN is null)"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If (rs Is Nothing) Or (rs.RecordCount = 0) Then
            GoTo proc_exit
        End If
    End With
    
    ''1. �ڋqSXL�d�l_1 (TBCME005) �̓��e���擾����
    ''1-1.SQL��g�ݗ��Ă�(�w��̑��Ə����܂ł̃��R�[�h:���Ə����t��)
    TbName = "TBCME005"
    sqlBase = "Select KSXTYPKB, KSXRUNIT, KSXRKKBN, KSXD1KBN, KSXD2KBN," & _
                " KSXDFKBN, KSXDPKBN, KSXDWKBN, KSXDDKBN, KSXDAKBN "
    sqlBase = sqlBase & "From " & TbName
    ''1-2.�f�[�^�𒊏o�E�i�[����
    '�d�l���R�[�h
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, SxlKokyaku_1) = FUNCTION_RETURN_FAILURE Then
        DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
'        Exit Function
    End If
    
    ''2. �ڋqSXL�d�l_2 (TBCME006) �̓��e���擾����
    ''2-1.SQL��g�ݗ��Ă�(�w��̑��Ə����܂ł̃��R�[�h:���Ə����t��)
    TbName = "TBCME006"
    sqlBase = "Select KSXTMKBN, KSXLTUNT, KSXLTKBN, KSXCNIND, KSXCNUNT," & _
                " KSXCNKBN, KSXONIND, KSXONUNT, KSXONKBN "
    sqlBase = sqlBase & "From " & TbName
    ''2-2.�f�[�^�𒊏o�E�i�[����
    '�d�l���R�[�h
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, SxlKokyaku_2) = FUNCTION_RETURN_FAILURE Then
        DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
'        Exit Function
    End If
    
    ''3. �ڋqSXL�d�l_3 (TBCME007) �̓��e���擾����
    ''3-1.SQL��g�ݗ��Ă�(�w��̑��Ə����܂ł̃��R�[�h:���Ə����t��)
    TbName = "TBCME007"
    sqlBase = "Select KSXOF1KBN, KSXOF1FGS, KSXOF1SO1, KSXOF1ST1, KSXOF2KB, KSXOF2GS, " & _
                "KSXOF2O1, KSXOF2ST, KSXBMKBN, KSXBMFGS, KSXBM2KB, KSXBM2GS "
    sqlBase = sqlBase & "From " & TbName
    ''3-2.�f�[�^�𒊏o�E�i�[����
    '�d�l���R�[�h
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, SxlKokyaku_3) = FUNCTION_RETURN_FAILURE Then
        DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
'        Exit Function
    End If
    
    ''4. ���iSXL�d�l_1 (TBCME018) �̓��e���擾����
    ''4-1.SQL��g�ݗ��Ă�(�w��̑��Ə����܂ł̃��R�[�h:���Ə����t��)
    TbName = "TBCME018"
    'sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND," & _
                " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX," & _
                " HSXRSPOH||HSXRSPOT||HSXRSPOI as HSXSPO," & _
                " HSXRHWYT||HSXRHWYS as HSXRHWY," & _
                " HSXRKWAY," & _
                " HSXRKHNM||HSXRKHNI||HSXRKHNH||HSXRKHNS as HSXRKHN," & _
                " HSXRMCAL, HSXRMBNP, HSXRMCL2," & _
                " HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM, HSXD1CEN," & _
                " HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR," & _
                " HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
                " HSXCKHNM||HSXCKHNI||HSXCKHNH||HSXCKHNS as HSXCKHN," & _
                " HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN," & _
                " HSXCTMIN, HSXCTMAX, HSXCYDIR, HSXCYCEN, HSXCYMIN, HSXCYMAX," & _
                " HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY," & _
                " HSXDPDIR, HSXDPMIN, HSXDPMAX, HSXDWCEN, HSXDWMIN, HSXDWMAX," & _
                " HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX," & _
                " SPECRRNO, SXLMCNO, WFMCNO "
    sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND," & _
                " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH," & _
                " HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
                " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2," & _
                " HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM, HSXD1CEN," & _
                " HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR," & _
                " HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY, HSXCKHNM, HSXCKHNI," & _
                " HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN," & _
                " HSXCTMIN, HSXCTMAX, HSXCYDIR, HSXCYCEN, HSXCYMIN, HSXCYMAX," & _
                " HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY," & _
                " HSXDPDIR, HSXDPMIN, HSXDPMAX, HSXDWCEN, HSXDWMIN, HSXDWMAX," & _
                " HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX," & _
                " SPECRRNO, SXLMCNO, WFMCNO, MCNO, SSTAFFID, SYNDATE "
    sqlBase = sqlBase & "From " & TbName
    ''4-2.�f�[�^�𒊏o�E�i�[����
    '�d�l���R�[�h
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Sxl_1) = FUNCTION_RETURN_FAILURE Then
        DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
'        Exit Function
    End If
    
    ''5. ���iSXL�d�l_2 (TBCME019) �̓��e���擾����
    ''5-1.SQL��g�ݗ��Ă�(�w��̑��Ə����܂ł̃��R�[�h:���Ə����t��)
    TbName = "TBCME019"
    'sqlBase = "Select HSXTMMAX," & _
                " HSXTMSPH||HSXTMSPT||HSXTMSPR as HSXTMSP," & _
                " HSXTMKHM||HSXTMKHI||HSXTMKHH||HSXTMKHS as HSXTMKH," & _
                " HSXLTMIN, HSXLTMAX," & _
                " HSXLTSPH||HSXLTSPT||HSXLTSPI as HSXLTSP," & _
                " HSXLTHWT||HSXLTHWS as HSXLTHW," & _
                " HSXLTNSW," & _
                " HSXLTKHM||HSXLTKHI||HSXLTKHH||HSXLTKHS as HSXLTKH," & _
                " HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
                " HSXCNSPH||HSXCNSPT||HSXCNSPI as HSXCNSP," & _
                " HSXCNHWT||HSXCNHWS as HSXCNHW," & _
                " HSXCNKWY," & _
                " HSXCNKHM||HSXCNKHI||HSXCNKHH||HSXCNKHS as HSXCNKH," & _
                " HSXONMIN, HSXONMAX,"
    'sqlBase = sqlBase & " HSXONSPH||HSXONSPT||HSXONSPI as HSXONSP," & _
                " HSXONHWT||HSXONHWS as HSXONHW," & _
                " HSXONKWY," & _
                " HSXONKHM||HSXONKHI||HSXONKHH||HSXONKHS as HSXONKH," & _
                " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS," & _
                " HSXOS1SH||HSXOS1ST||HSXOS1SI as HSXOS1S," & _
                " HSXOS1HT||HSXOS1HS as HSXOS1H," & _
                " HSXOS1HM||HSXOS1KI||HSXOS1KH||HSXOS1KS as HSXOS1K," & _
                " HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
                " HSXOS2SH||HSXOS2ST||HSXOS2SI as HSXOS2S," & _
                " HSXOS2HT||HSXOS2HS as HSXOS2H," & _
                " HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU as HSXOS2K, HSXTMMAXN "
    sqlBase = "Select HSXTMMAX, HSXTMSPH, HSXTMSPT, HSXTMSPR, HSXTMKHM," & _
                " HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH," & _
                " HSXLTSPT, HSXLTSPI, HSXLTHWT, HSXLTHWS, HSXLTNSW, HSXLTKHM," & _
                " HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN," & _
                " HSXCNMAX, HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS," & _
                " HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
                " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS," & _
                " HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS, HSXONMBP," & _
                " HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX," & _
                " HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH, HSXOS1ST, HSXOS1SI," & _
                " HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS," & _
                " HSXOS2MN, HSXOS2MX, HSXOS2NS, HSXOS2SH, HSXOS2ST, HSXOS2SI," & _
                " HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, HSXTMMAXN "
    sqlBase = sqlBase & "From " & TbName
    ''5-2.�f�[�^�𒊏o�E�i�[����
    '�d�l���R�[�h
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Sxl_2) = FUNCTION_RETURN_FAILURE Then
        DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
'        Exit Function
    End If
    
    ''6. ���iSXL�d�l_3 (TBCME020) �̓��e���擾����
    ''6-1.SQL��g�ݗ��Ă�(�w��̑��Ə����܂ł̃��R�[�h:���Ə����t��)
    TbName = "TBCME020"
    'sqlBase = "Select HSXDENKU," & _
              " HSXDENHT||HSXDENHS as HSXDENH," & _
              " HSXDVDKU," & _
              " HSXDVDHT||HSXDVDHS as HSXDVDH," & _
              " HSXLDLKU," & _
              " HSXLDLHT||HSXLDLHS as HSXLDLH," & _
              " HSXGDSPH||HSXGDSPT||HSXGDSPR as HSXGDSP," & _
              " HSXGDSZY," & _
              " HSXGDKHM||HSXGDKHI||HSXGDKHH||HSXGDKHS as HSXGDKH," & _
              " HSXDSOHT||HSXDSOHS as HSXDSOH," & _
              " HSXDSOKM||HSXDSOKI||HSXDSOKH||HSXDSOKS as HSXDSOK," & _
              " HSXLIFTW, HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDPNI, HSXGSFIN, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SZ," & _
              " HSXOF1SH||HSXOF1ST||HSXOF1SR as HSXOF1S," & _
              " HSXOF1HT||HSXOF1HS as HSXOF1H," & _
              " HSXOF1NS, HSXOF1ET,"
    'sqlBase = sqlBase & " HSXOF1KM||HSXOF1KI||HSXOF1KH||HSXOF1KS as HSXOF1K," & _
              " HSXOF2AX, HSXOF2MX, HSXOF2SZ," & _
              " HSXOF2SH||HSXOF2ST||HSXOF2SR as HSXOF2S," & _
              " HSXOF2HT||HSXOF2HS as HSXOF2H," & _
              " HSXOF2NS, HSXOF2ET," & _
              " HSXOF2KM||HSXOF2KI||HSXOF2KH||HSXOF2KS as HSXOF2K," & _
              " HSXOF3AX, HSXOF3MX, HSXOF3SZ," & _
              " HSXOF3SH||HSXOF3ST||HSXOF3SR as HSXOF3S," & _
              " HSXOF3HT||HSXOF3HS as HSXOF3H," & _
              " HSXOF3NS, HSXOF3ET," & _
              " HSXOF3KM||HSXOF3KI||HSXOF3KH||HSXOF3KS as HSXOF3K," & _
              " HSXOF4AX, HSXOF4MX, HSXOF4SZ," & _
              " HSXOF4SH||HSXOF4ST||HSXOF4SR as HSXOF4S," & _
              " HSXOF4HT||HSXOF4HS as HSXOF4H," & _
              " HSXOF4NS, HSXOF4ET," & _
              " HSXOF4KM||HSXOF4KI||HSXOF4KH||HSXOF4KS as HSXOF4K,"
    'sqlBase = sqlBase & " HSXBM1SZ, HSXBM1AN, HSXBM1AX," & _
              " HSXBM1SH||HSXBM1ST||HSXBM1SR as HSXBM1S," & _
              " HSXBM1HT||HSXBM1HS as HSXBM1H," & _
              " HSXBM1NS,HSXBM1ET," & _
              " HSXBM1KM||HSXBM1KI||HSXBM1KH||HSXBM1KS as HSXBM1K," & _
              " HSXBM2SZ, HSXBM2AN, HSXBM2AX," & _
              " HSXBM2SH||HSXBM2ST||HSXBM2SR as HSXBM2S," & _
              " HSXBM2HT||HSXBM2HS as HSXBM2H," & _
              " HSXBM2NS,HSXBM2ET," & _
              " HSXBM2KM||HSXBM2KI||HSXBM2KH||HSXBM2KS as HSXBM2K," & _
              " HSXBM3SZ, HSXBM3AN, HSXBM3AX," & _
              " HSXBM3SH||HSXBM3ST||HSXBM3SR as HSXBM3S," & _
              " HSXBM3HT||HSXBM3HS as HSXBM3H," & _
              " HSXBM3NS,HSXBM3ET," & _
              " HSXBM3KM||HSXBM3KI||HSXBM3KH||HSXBM3KS as HSXBM3K," & _
              " HSXDVDMNN, HSXDVDMXN, HSXDSONS, HSXCDOPMN, HSXCDOPMX," & _
              " HSXOSF1PTK, HSXOSF2PTK, HSXOSF3PTK, HSXOSF4PTK, " & _
              " HSXBMD1MBP, HSXBMD1MCL, HSXBMD2MBP, HSXBMD2MCL, HSXBMD3MBP, HSXBMD3MCL "
    sqlBase = "Select HSXDENKU, HSXDENMX, HSXDENMN, HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMX, HSXDVDMN, HSXDVDHT, HSXDVDHS, HSXLDLKU," & _
                " HSXLDLMX, HSXLDLMN, HSXLDLHT, HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH," & _
                " HSXGDKHS, HSXDSOKE, HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS," & _
                " HSXLIFTW, HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
                " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH, HSXOF1KS," & _
                " HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ, HSXOF2KM, HSXOF2KI," & _
                " HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR, HSXOF3HT, HSXOF3HS, HSXOF3SZ," & _
                " HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX, HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT," & _
                " HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS, HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST," & _
                " HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI, HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX," & _
                " HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS, HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET," & _
                " HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST, HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS," & _
                " HSXBM3NS, HSXBM3ET, HSXNOTE, HSXRS1N, HSXRS1Y, HSXRS2N, HSXRS2Y, HSXRS3N, HSXRS3Y, HSXRS4N, HSXRS4Y, HSXRS5N, HSXRS5Y," & _
                " HSXRS6N, HSXRS6Y, HSXRS7N, HSXRS7Y, HSXRS8N, HSXRS8Y, HSXRS9N, HSXRS9Y, HSXRS10N, HSXRS10Y, " & _
                " HSXDVDMNN, HSXDVDMXN, HSXDSONS, HSXCDOPMN, HSXCDOPMX, HSXOSF1PTK, HSXOSF2PTK, HSXOSF3PTK, HSXOSF4PTK, " & _
                " HSXBMD1MBP, HSXBMD1MCL, HSXBMD2MBP, HSXBMD2MCL, HSXBMD3MBP, HSXBMD3MCL, HSXDSOPTK "
    sqlBase = sqlBase & "From " & TbName
    ''6-2.�f�[�^�𒊏o�E�i�[����
    '�d�l���R�[�h
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Sxl_3) = FUNCTION_RETURN_FAILURE Then
        DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
'        Exit Function
    End If
    
    ''7. �ڋq�d�lWF�ް�2 (TBCME009) �̓��e���擾����
    ''7-1.SQL��g�ݗ��Ă�(�w��̑��Ə����܂ł̃��R�[�h:���Ə����t��)
    TbName = "TBCME009"
    sqlBase = "Select KPRDFORM "
    sqlBase = sqlBase & "From " & TbName
    ''7-2.�f�[�^�𒊏o�E�i�[����
    '�d�l���R�[�h
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WfKokyaku_2) = FUNCTION_RETURN_FAILURE Then
        DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
'        Exit Function
    End If
    
    ''8. �ڋq�d�lWF�ް�8 (TBCME028) �̓��e���擾����
    ''8-1.SQL��g�ݗ��Ă�(�w��̑��Ə����܂ł̃��R�[�h:���Ə����t��)
    TbName = "TBCME028"
    sqlBase = "Select HWFMK1SI, HWFMK2SI, HWFMK3SI, HWFMK4SI, HWFMK1MX, HWFMK2MX, HWFMK3MX, HWFMK4MX "
    sqlBase = sqlBase & "From " & TbName
    ''8-2.�f�[�^�𒊏o�E�i�[����
    '�d�l���R�[�h
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, WfKokyaku_8) = FUNCTION_RETURN_FAILURE Then
        DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
'        Exit Function
    End If
    
    ''9. ���������Ǘ� (TBCME036) �̓��e���擾����
    ''9-1.SQL��g�ݗ��Ă�(�w��̑��Ə����܂ł̃��R�[�h:���Ə����t��)
    TbName = "TBCME036"
    sqlBase = "Select EPDSETCH, EPDUP, CUTUNIT "
    sqlBase = sqlBase & "From " & TbName
    ''9-2.�f�[�^�𒊏o�E�i�[����
    '�d�l���R�[�h
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, Sxluchigawa) = FUNCTION_RETURN_FAILURE Then
        DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_FAILURE
'        Exit Function
    End If
    
    ''10. LPD�T�C�Y�ELPD����̕\���l��ݒ肷��
    With WfKokyaku_8
        'HWFMK1SI�`HWFMK4SI, HWFMK1MX�`HWFMK4MX ���擾
        HWFMKnSI(1) = .Fields.GetValueOrDefault("HWFMK1SI", 0#)
        HWFMKnMX(1) = .Fields.GetValueOrDefault("HWFMK1MX", 0)
        HWFMKnSI(2) = .Fields.GetValueOrDefault("HWFMK2SI", 0#)
        HWFMKnMX(2) = .Fields.GetValueOrDefault("HWFMK2MX", 0)
        HWFMKnSI(3) = .Fields.GetValueOrDefault("HWFMK3SI", 0#)
        HWFMKnMX(3) = .Fields.GetValueOrDefault("HWFMK3MX", 0)
        HWFMKnSI(4) = .Fields.GetValueOrDefault("HWFMK4SI", 0#)
        HWFMKnMX(4) = .Fields.GetValueOrDefault("HWFMK4MX", 0)
    End With
    '��₩��i�荞��
    HWFMKnSI(0) = 9999#   '�[���傫�Ȓl
    For i = 1 To 4
        If HWFMKnSI(0) > HWFMKnSI(i) Then
            HWFMKnSI(0) = HWFMKnSI(i)
            HWFMKnMX(0) = HWFMKnMX(i)
        End If
    Next
    '����ꂽ���ʂ�WfKokyaku_8(1)�ɓo�^����
    WfKokyaku_8.Fields.Add "HWFMKnSI", HWFMKnSI(0), ORADB_DOUBLE, -1
    WfKokyaku_8.Fields.Add "HWFMKnMX", HWFMKnMX(0), ORADB_INTEGER, -1
    
    DBDRV_s_cmzcF_cmfc001c_DispSXL = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :��������̊T�v�𓾂�
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :jokenNo       ,I  ,String    ,��������ԍ�
'          :rec           ,O  ,c_cmzcrec ,�T�v
'          :�߂�l        ,O  ,FUNCTION_RETURN,�����̐���
'����      :
'����      :2001/06/08 �쐬  �쑺
Public Function DBDRV_s_cmzcF_cmfc001c_GetSJoken(ByVal jokenNo$, rec As c_cmzcrec) As FUNCTION_RETURN
Dim sql As String           'SQL
Dim TbName As String        '�e�[�u����

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001c_SQL.bas -- Function DBDRV_s_cmzcF_cmfc001c_GetSJoken"

    ''�o�̓f�[�^������������
    Set rec = New c_cmzcrec
    
    ''1. ������� (TBCMB012) �̓��e���擾����
    ''1-1.SQL��g�ݗ��Ă�
    TbName = "TBCMB012"
    sql = "Select MODEL, RTBSIZE, CHARGE, HZTYPE, UPSPDTYP, MAGTYPE" & _
          " From " & TbName & _
          " Where (rtrim(MKCONDNO)='" & jokenNo & "')"
    ''1-2.�f�[�^�𒊏o�E�i�[����
    DBDRV_s_cmzcF_cmfc001c_GetSJoken = DispSXL_GetData(TbName, sql, rec)

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :��������ԍ��ɑΉ�����PGID�̈ꗗ�𓾂�
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :jokenNo       ,I  ,String    ,��������ԍ�
'          :PGIDs()       ,O  ,String    ,�Ή�����PGID�̈ꗗ
'          :�߂�l        ,O  ,FUNCTION_RETURN,�����̐���
'����      :
'����      :2001/06/08 �쐬  �쑺
Public Function DBDRV_s_cmzcF_cmfc001c_GetPGID(ByVal jokenNo$, PGIDs() As String) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001c_SQL.bas -- Function DBDRV_s_cmzcF_cmfc001c_GetPGID"

    sql = "Select PGIDNO " & _
          "From TBCMB013 " & _
          "Where (rtrim(MKCONDNO)='" & jokenNo & "')"

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim PGIDs(0)
        DBDRV_s_cmzcF_cmfc001c_GetPGID = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim PGIDs(recCnt)
    For i = 1 To recCnt
        PGIDs(i) = rs("PGIDNO")     ' PG-IDNo
        rs.MoveNext
    Next
    rs.Close

    DBDRV_s_cmzcF_cmfc001c_GetPGID = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

Public Function DBDRV_s_cmzcF_cmfc001c_GetHikiage(targetHinban As tFullHinban, Hikiage As String) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001c_SQL.bas -- Function DBDRV_s_cmzcF_cmfc001c_GetHikiage"
    With targetHinban
    'sql = "Select SSXLIFTW " & _
          "From TBCME030 " & _
          " Where (HINBAN='" & .HINBAN & "') AND (MNOREVNO=" & .MNOREVNO & ") " & _
              "AND (FACTORY='" & .FACTORY & "') AND (OPECOND='1')"
    sql = "Select SSXLIFTW " & _
          "From TBCME036 " & _
          " Where (HINBAN='" & .hinban & "') AND (MNOREVNO=" & .MNOREVNO & ") " & _
              "AND (FACTORY='" & .FACTORY & "') AND (OPECOND='1')"
    End With
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        DBDRV_s_cmzcF_cmfc001c_GetHikiage = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    If IsNull(rs("SSXLIFTW")) Then
        Hikiage = ""
    Else
        Hikiage = rs("SSXLIFTW")     '    rs.Close
    End If
    DBDRV_s_cmzcF_cmfc001c_GetHikiage = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'���݂̎d�l�o�^�ł͖��g�p
'�T�v      :[���s]���f�[�^��������
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :IraiNo        ,I  ,String    ,�d�l�o�^�˗��ԍ�
'          :SXLMCNO       ,I  ,String    ,SXL�������(�d�l��)
'          :WFMCNO        ,I  ,String    ,WF�������(�d�l��)
'          :Hinban12      ,I  ,String    ,12���i��
'          :SJokenNo      ,I  ,String    ,��������ԍ�
'          :Hikiage       ,I  ,String    ,������@
'          :StaffID       ,I  ,String    ,�S���҃R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN,�����̐���
'����      :
'����      :2001/06/08 �쐬  �쑺
Public Function DBDRV_s_cmzcF_cmfc001c_Exec(ByVal IraiNo$, ByVal SXLMCNO$, ByVal WFMCNO$, ByVal Hinban12$, _
                                            ByVal SJokenNo$, ByVal Hikiage$, ByVal StaffID$) As FUNCTION_RETURN
Dim sql_top As String
Dim sql_sel As String
Dim sql As String
Dim fld As OraField
Dim rs As OraDynaset
Dim fullHinban As tFullHinban


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc001c_SQL.bas -- Function DBDRV_s_cmzcF_cmfc001c_Exec"

    If Len(Hinban12) <> 12 Then
        DBDRV_s_cmzcF_cmfc001c_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    With fullHinban
        .hinban = Left$(Hinban12, 8)
        .MNOREVNO = val(Mid$(Hinban12, 9, 2))
        .FACTORY = Mid$(Hinban12, 11)
        .OPECOND = Right$(Hinban12, 1)
    End With
    sql = "insert into TBCME030 " & _
          "(HINBAN, MNOREVNO, FACTORY, OPECOND, SSXLIFTW, IFKBN, SYORIKBN, SPECRRNO, SXLMCNO, WFMCNO, " & _
          "STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE) " & _
          "values ("
    sql = sql & "'" & fullHinban.hinban & "', "     ' �i��
    sql = sql & fullHinban.MNOREVNO & ", "          ' ���i�ԍ������ԍ�
    sql = sql & "'" & fullHinban.FACTORY & "', "    ' �H��
    sql = sql & "'" & fullHinban.OPECOND & "', "    ' ���Ə���
    sql = sql & "'" & Hikiage & "', "               ' ���r�w������@
    sql = sql & "' ', "                             ' �h�^�e�敪
    sql = sql & "' ', "                             ' �����敪
    sql = sql & "'" & IraiNo & "', "                ' �d�l�o�^�˗��ԍ�
    sql = sql & "'" & SXLMCNO & "', "               ' �r�w�k��������ԍ�
    sql = sql & "'" & WFMCNO & "', "                ' �v�e��������ԍ�
    sql = sql & "'" & StaffID & "', "               ' �Ј�ID
    sql = sql & "SYSDATE, "                         ' �o�^���t
    sql = sql & "SYSDATE, "                         ' �X�V���t
    sql = sql & "'0', "                             ' ���M�t���O
    sql = sql & "SYSDATE "                          ' ���M���t
    sql = sql & ")"
    If 0 >= OraDB.ExecuteSQL(sql) Then
        GoTo proc_err
    End If
    
    ''�i�ԃf�[�^�ɐ������No����������(���d�l�ł��郊�r�W�����P�������e���r�W�����Ɂj
    sql = "update TBCME018 set " & _
          "MCNO = '" & SJokenNo & "' " & _
          "where " & _
          "(HINBAN = '" & fullHinban.hinban & "') and " & _
          "(MNOREVNO = " & fullHinban.MNOREVNO & ") and " & _
          "(FACTORY = '" & fullHinban.FACTORY & "')"
    OraDB.ExecuteSQL sql
    If 0 >= OraDB.ExecuteSQL(sql) Then
        GoTo proc_err
    End If
    
    DBDRV_s_cmzcF_cmfc001c_Exec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "==== Error SQL ===="
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function



'------------------------------------------------
' DB�A�N�Z�X�֐��i���o�ҁj
'------------------------------------------------
'�T�v      :�e�[�u���uTBCMB011�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMB011 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/06/11�쐬�@����
Public Function DBDRV_cmbc001d_Disp(records() As typ_TBCMB011, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql     As String                   'SQL�S��
Dim sqlBase As String                   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs      As OraDynaset               'RecordSet
Dim recCnt  As Long                     '���R�[�h��
Dim i       As Long                     'ٰ�߶���

    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001d_SQL.bas -- Function DBDRV_cmbc001d_Disp"

    sqlBase = "Select PGID, HZPART, HZPTRN, SPACER, UPRING, CHARGE, RTBPOS, RTBSIZE, GAP, UPDM, UPLENGTH, UPRC, RFRNEED, UPSPIN," & _
              " DOWNSPIN, ROPRESS, ARUGON, AIMOIMIN, AIMOIMAX, HCCLASS, HC, AVEUPSPD, UPCNTL, BTMSHAPE, MAGSTR, MAGPOS, CONDGRT," & _
              " MODEL, UPMETHOD, UPCLASS, UPNUM, OPETIME, WTRCOOL, PGID2, RCPT1, RCPT2, RCPT3, RCPT4, RCPT5, CNTL1, CNTL2," & _
              " CNTL3, CNTL4, CNTL5, CNTL6, CNTL7, CNTL8, CNTL9, CNTL10, CNTL11, CNTL12, CNTL13, CNTL14, CNTL15, RUNCOND1," & _
              " RUNCOND2, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMB011 "
    sql = sqlBase
    If (sqlWhere <> vbNullString) Then
        sql = sql & sqlWhere
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_cmbc001d_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .PGID = rs("PGID")               ' PG-ID
            .HZPART = rs("HZPART")           ' HZ�p�[�c
            .HZPTRN = rs("HZPTRN")           ' HZ�p�^�[��
            .SPACER = rs("SPACER")           ' �X�y�[�T
            .UPRING = rs("UPRING")           ' �A�b�p�[�����O
            .CHARGE = rs("CHARGE")           ' �`���[�W��
            .RTBPOS = rs("RTBPOS")           ' ���c�{�ʒu
            .RTBSIZE = rs("RTBSIZE")         ' ���c�{�T�C�Y
            .GAP = rs("GAP")                 ' �M���b�v
            .UPDM = rs("UPDM")               ' ���㒼�a
            .UPLENGTH = rs("UPLENGTH")       ' ���㒷�i�S���j
            .UPRC = rs("UPRC")               ' ����iRC�j
            .RFRNEED = rs("RFRNEED")         ' ���t���N�^�v��
            .UPSPIN = rs("UPSPIN")           ' �㎲��]��
            .DOWNSPIN = rs("DOWNSPIN")       ' ������]��
            .ROPRESS = rs("ROPRESS")         ' �F����
            .ARUGON = rs("ARUGON")           ' �A���S����
            .AIMOIMIN = rs("AIMOIMIN")       ' �˂炢Oi�iMIN)
            .AIMOIMAX = rs("AIMOIMAX")       ' �˂炢Oi�iMAX)
            .HCCLASS = rs("HCCLASS")         ' HC���
            .HC = rs("HC")                   ' HC
            .AVEUPSPD = rs("AVEUPSPD")       ' ���ψ��㑬�x
            .UPCNTL = rs("UPCNTL")           ' ���㐧��
            .BTMSHAPE = rs("BTMSHAPE")       ' �{�g���`��
            .MAGSTR = rs("MAGSTR")           ' ���ꋭ�x
            .MAGPOS = rs("MAGPOS")           ' ����ʒu
            .CONDGRT = rs("CONDGRT")         ' �����ۏؓo�^
            .MODEL = rs("MODEL")             ' �@��
            .UPMETHOD = rs("UPMETHOD")       ' ������@
            .UPCLASS = rs("UPCLASS")         ' ����敪
            .UPNUM = rs("UPNUM")             ' ����{��
            .OPETIME = rs("OPETIME")         ' �^�]����
            .WTRCOOL = rs("WTRCOOL")         ' ����Ǘv��
            .PGID2 = rs("PGID2")             ' PG-ID�i��{���j
            .RCPT1 = rs("RCPT1")             ' �Ή����V�sNo�iT1)
            .RCPT2 = rs("RCPT2")             ' �Ή����V�sNo�iT2)
            .RCPT3 = rs("RCPT3")             ' �Ή����V�sNo�iT3)
            .RCPT4 = rs("RCPT4")             ' �Ή����V�sNo�iT4)
            .RCPT5 = rs("RCPT5")             ' �Ή����V�sNo�iT5)
            .CNTL1 = rs("CNTL1")             ' ���䍀�ځi1�j
            .CNTL2 = rs("CNTL2")             ' ���䍀�ځi2�j
            .CNTL3 = rs("CNTL3")             ' ���䍀�ځi3�j
            .CNTL4 = rs("CNTL4")             ' ���䍀�ځi4�j
            .CNTL5 = rs("CNTL5")             ' ���䍀�ځi5�j
            .CNTL6 = rs("CNTL6")             ' ���䍀�ځi6�j
            .CNTL7 = rs("CNTL7")             ' ���䍀�ځi7�j
            .CNTL8 = rs("CNTL8")             ' ���䍀�ځi8�j
            .CNTL9 = rs("CNTL9")             ' ���䍀�ځi9�j
            .CNTL10 = rs("CNTL10")           ' ���䍀�ځi10�j
            .CNTL11 = rs("CNTL11")           ' ���䍀�ځi11�j
            .CNTL12 = rs("CNTL12")           ' ���䍀�ځi12�j
            .CNTL13 = rs("CNTL13")           ' ���䍀�ځi13�j
            .CNTL14 = rs("CNTL14")           ' ���䍀�ځi14�j
            .CNTL15 = rs("CNTL15")           ' ���䍀�ځi15�j
            .RUNCOND1 = rs("RUNCOND1")       ' �^�]�����P
            .RUNCOND2 = rs("RUNCOND2")       ' �^�]�����Q
'            .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
'            .REGDATE = rs("REGDATE")         ' �o�^���t
'            .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
'            .UPDDATE = rs("UPDDATE")         ' �X�V���t
'            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
'            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_cmbc001d_Disp = FUNCTION_RETURN_SUCCESS

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

Public Function DBDRV_cmbc001d_Exec(records As typ_TBCMB011) As FUNCTION_RETURN
'------------------------------------------------
' DB�A�N�Z�X�֐��i�X�V�ҁj
'------------------------------------------------
'�T�v      :�e�[�u���uTBCMB011�v�̏����ɂ��������R�[�h�ɍX�V��������
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records�@     ,O  ,typ_TBCMB011 ,���o���R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :���o�ް��̌����E����������"�ς�"�Ƃ���
'����      :2001/06/19(TUE)�쐬�@����
Dim sql     As String                   'SQL�S��
Dim rs      As OraDynaset               'RecordSet
Dim UpdID   As String                   '�X�V�Ώ�PGID


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001d_SQL.bas -- Function DBDRV_cmbc001d_Exec"

    UpdID = records.PGID

'2001/09/05 S.Sano Start �X�V�������Z�b�g����Ă��Ȃ��B
'2001/09/05 S.Sano Start ���̃��[�h��sysdate�̃Z�b�g���@���s���B
'    sql = "SELECT * FROM TBCMB011 WHERE(PGID = '" & UpdID & "')"
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'
'    rs.Edit
'    With records
'        rs("HZPART") = StrNoNull(.HZPART)          ' HZ�߰�
'        rs("HZPTRN") = StrNoNull(.HZPTRN)          ' HZ�����
'        rs("SPACER") = StrNoNull(.SPACER)          ' ��߰�
'        rs("UPRING") = StrNoNull(.UPRING)          ' ���߰�ݸ�
'        rs("CHARGE") = .CHARGE          ' ����ޗ�
'        rs("RTBPOS") = .RTBPOS          ' ���ވʒu
'        rs("RTBSIZE") = StrNoNull(.RTBSIZE)        ' ���޻���
'        rs("GAP") = .GAP                ' �ެ���
'        rs("UPDM") = .UPDM              ' ���㒼�a
'        rs("UPLENGTH") = .UPLENGTH      ' ���㒷
'        rs("UPRC") = .UPRC              ' ����RC
'        rs("RFRNEED") = StrNoNull(.RFRNEED)        ' ��׸��v��
'        rs("UPSPIN") = StrNoNull(.UPSPIN)          ' �㎲��]��
'        rs("DOWNSPIN") = StrNoNull(.DOWNSPIN)      ' ������]��
'        rs("ROPRESS") = StrNoNull(.ROPRESS)        ' �F����
'        rs("ARUGON") = StrNoNull(.ARUGON)          ' �ٺ�ݗ�
'        rs("AIMOIMIN") = .AIMOIMIN      ' �˂炢iO(MIN)
'        rs("AIMOIMAX") = .AIMOIMAX      ' �˂炢iO(MAX)
'        rs("HCCLASS") = StrNoNull(.HCCLASS)        ' HC���
'        rs("HC") = StrNoNull(.HC)                  ' HC
'        rs("AVEUPSPD") = .AVEUPSPD      ' ���ψ��㑬�x
'        rs("UPCNTL") = StrNoNull(.UPCNTL)          ' ���㐧��
'        rs("BTMSHAPE") = StrNoNull(.BTMSHAPE)      ' ���ь`��
'        rs("MAGSTR") = .MAGSTR          ' ���ꋭ�x
'        rs("MAGPOS") = .MAGPOS          ' ����ʒu
'        rs("CONDGRT") = StrNoNull(.CONDGRT)        ' �����ۏؓo�^
'        rs("MODEL") = StrNoNull(.MODEL)            ' �@��
'        rs("UPMETHOD") = StrNoNull(.UPMETHOD)      ' ������@
'        rs("UPCLASS") = StrNoNull(.UPCLASS)        ' ����敪
'        rs("UPNUM") = StrNoNull(.UPNUM)            ' ����{��
'        rs("OPETIME") = .OPETIME        ' �^�]����
'        rs("WTRCOOL") = StrNoNull(.WTRCOOL)        ' ����Ǘv��
'        rs("PGID2") = StrNoNull(.PGID2)            ' PG-ID�i��{���j
'        rs("RCPT1") = StrNoNull(.RCPT1)            ' �Ή�ڼ��No�iT1�j
'        rs("RCPT2") = StrNoNull(.RCPT2)            ' �Ή�ڼ��No�iT2�j
'        rs("RCPT3") = StrNoNull(.RCPT3)            ' �Ή�ڼ��No�iT3�j
'        rs("RCPT4") = StrNoNull(.RCPT4)            ' �Ή�ڼ��No�iT4�j
'        rs("RCPT5") = StrNoNull(.RCPT5)            ' �Ή�ڼ��No�iT5�j
'        rs("CNTL1") = StrNoNull(.CNTL1)            ' ���䍀��(1)
'        rs("CNTL2") = StrNoNull(.CNTL2)            ' ���䍀��(2)
'        rs("CNTL3") = StrNoNull(.CNTL3)            ' ���䍀��(3)
'        rs("CNTL4") = StrNoNull(.CNTL4)            ' ���䍀��(4)
'        rs("CNTL5") = StrNoNull(.CNTL5)            ' ���䍀��(5)
'        rs("CNTL6") = StrNoNull(.CNTL6)            ' ���䍀��(6)
'        rs("CNTL7") = StrNoNull(.CNTL7)            ' ���䍀��(7)
'        rs("CNTL8") = StrNoNull(.CNTL8)            ' ���䍀��(8)
'        rs("CNTL9") = StrNoNull(.CNTL9)            ' ���䍀��(9)
'        rs("CNTL10") = StrNoNull(.CNTL10)          ' ���䍀��(10)
'        rs("CNTL11") = StrNoNull(.CNTL11)          ' ���䍀��(11)
'        rs("CNTL12") = StrNoNull(.CNTL12)          ' ���䍀��(12)
'        rs("CNTL13") = StrNoNull(.CNTL13)          ' ���䍀��(13)
'        rs("CNTL14") = StrNoNull(.CNTL14)          ' ���䍀��(14)
'        rs("CNTL15") = StrNoNull(.CNTL15)          ' ���䍀��(15)
'        rs("RUNCOND1") = StrNoNull(.RUNCOND1)      ' �^�]����1
'        rs("RUNCOND2") = StrNoNull(.RUNCOND2)      ' �^�]����2
'    End With
'    rs.Update
'
'    rs.Close
    
'2001/09/05 S.Sano Start
    With records
    sql = "update TBCMB011 set "
    sql = sql & "HZPART = '" & StrNoNull(.HZPART) & "', "       ' HZ�߰�
    sql = sql & "HZPTRN = '" & StrNoNull(.HZPTRN) & "', "       ' HZ�����
    sql = sql & "SPACER = '" & StrNoNull(.SPACER) & "', "       ' ��߰�
    sql = sql & "UPRING = '" & StrNoNull(.UPRING) & "', "       ' ���߰�ݸ�
    sql = sql & "CHARGE = " & .CHARGE & ", "                    ' ����ޗ�
    sql = sql & "RTBPOS = " & .RTBPOS & ", "                    ' ���ވʒu
    sql = sql & "RTBSIZE = '" & StrNoNull(.RTBSIZE) & "', "     ' ���޻���
    sql = sql & "GAP = " & .GAP & ", "                          ' �ެ���
    sql = sql & "UPDM = " & .UPDM & ", "                        ' ���㒼�a
    sql = sql & "UPLENGTH = " & .UPLENGTH & ", "                ' ���㒷
    sql = sql & "UPRC = " & .UPRC & ", "                        ' ����RC
    sql = sql & "RFRNEED = '" & StrNoNull(.RFRNEED) & "', "     ' ��׸��v��
    sql = sql & "UPSPIN = '" & StrNoNull(.UPSPIN) & "', "       ' �㎲��]��
    sql = sql & "DOWNSPIN = '" & StrNoNull(.DOWNSPIN) & "', "   ' ������]��
    sql = sql & "ROPRESS = '" & StrNoNull(.ROPRESS) & "', "     ' �F����
    sql = sql & "ARUGON = '" & StrNoNull(.ARUGON) & "', "       ' �ٺ�ݗ�
    sql = sql & "AIMOIMIN = " & .AIMOIMIN & ", "                ' �˂炢iO(MIN)
    sql = sql & "AIMOIMAX = " & .AIMOIMAX & ", "                ' �˂炢iO(MAX)
    sql = sql & "HCCLASS = '" & StrNoNull(.HCCLASS) & "', "     ' HC���
    sql = sql & "HC = '" & StrNoNull(.HC) & "', "               ' HC
    sql = sql & "AVEUPSPD = " & .AVEUPSPD & ", "                ' ���ψ��㑬�x
    sql = sql & "UPCNTL = '" & StrNoNull(.UPCNTL) & "', "       ' ���㐧��
    sql = sql & "BTMSHAPE = '" & StrNoNull(.BTMSHAPE) & "', "   ' ���ь`��
    sql = sql & "MAGSTR = " & .MAGSTR & ", "                    ' ���ꋭ�x
    sql = sql & "MAGPOS = " & .MAGPOS & ", "                    ' ����ʒu
    sql = sql & "CONDGRT = '" & StrNoNull(.CONDGRT) & "', "     ' �����ۏؓo�^
    sql = sql & "MODEL = '" & StrNoNull(.MODEL) & "', "         ' �@��
    sql = sql & "UPMETHOD = '" & StrNoNull(.UPMETHOD) & "', "   ' ������@
    sql = sql & "UPCLASS = '" & StrNoNull(.UPCLASS) & "', "     ' ����敪
    sql = sql & "UPNUM = '" & StrNoNull(.UPNUM) & "', "         ' ����{��
    sql = sql & "OPETIME = " & .OPETIME & ", "                  ' �^�]����
    sql = sql & "WTRCOOL = '" & StrNoNull(.WTRCOOL) & "', "     ' ����Ǘv��
    sql = sql & "PGID2 = '" & StrNoNull(.PGID2) & "', "         ' PG-ID�i��{���j
    sql = sql & "RCPT1 = '" & StrNoNull(.RCPT1) & "', "         ' �Ή�ڼ��No�iT1�j
    sql = sql & "RCPT2 = '" & StrNoNull(.RCPT2) & "', "         ' �Ή�ڼ��No�iT2�j
    sql = sql & "RCPT3 = '" & StrNoNull(.RCPT3) & "', "         ' �Ή�ڼ��No�iT3�j
    sql = sql & "RCPT4 = '" & StrNoNull(.RCPT4) & "', "         ' �Ή�ڼ��No�iT4�j
    sql = sql & "RCPT5 = '" & StrNoNull(.RCPT5) & "', "         ' �Ή�ڼ��No�iT5�j
    sql = sql & "CNTL1 = '" & StrNoNull(.CNTL1) & "', "         ' ���䍀��(1)
    sql = sql & "CNTL2 = '" & StrNoNull(.CNTL2) & "', "         ' ���䍀��(2)
    sql = sql & "CNTL3 = '" & StrNoNull(.CNTL3) & "', "         ' ���䍀��(3)
    sql = sql & "CNTL4 = '" & StrNoNull(.CNTL4) & "', "         ' ���䍀��(4)
    sql = sql & "CNTL5 = '" & StrNoNull(.CNTL5) & "', "         ' ���䍀��(5)
    sql = sql & "CNTL6 = '" & StrNoNull(.CNTL6) & "', "         ' ���䍀��(6)
    sql = sql & "CNTL7 = '" & StrNoNull(.CNTL7) & "', "         ' ���䍀��(7)
    sql = sql & "CNTL8 = '" & StrNoNull(.CNTL8) & "', "         ' ���䍀��(8)
    sql = sql & "CNTL9 = '" & StrNoNull(.CNTL9) & "', "         ' ���䍀��(9)
    sql = sql & "CNTL10 = '" & StrNoNull(.CNTL10) & "', "       ' ���䍀��(10)
    sql = sql & "CNTL11 = '" & StrNoNull(.CNTL11) & "', "       ' ���䍀��(11)
    sql = sql & "CNTL12 = '" & StrNoNull(.CNTL12) & "', "       ' ���䍀��(12)
    sql = sql & "CNTL13 = '" & StrNoNull(.CNTL13) & "', "       ' ���䍀��(13)
    sql = sql & "CNTL14 = '" & StrNoNull(.CNTL14) & "', "       ' ���䍀��(14)
    sql = sql & "CNTL15 = '" & StrNoNull(.CNTL15) & "', "       ' ���䍀��(15)
    sql = sql & "RUNCOND1 = '" & StrNoNull(.RUNCOND1) & "', "   ' �^�]����1
    sql = sql & "RUNCOND2 = '" & StrNoNull(.RUNCOND2) & "', "   ' �^�]����2
    sql = sql & "KSTAFFID = '" & .KSTAFFID & "', "              ' �X�V�Ј�ID
    sql = sql & "UPDDATE = sysdate "                            ' �X�V���t
    sql = sql & "where PGID = '" & UpdID & "'"
    End With
    
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_cmbc001d_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_cmbc001d_Exec = FUNCTION_RETURN_SUCCESS
'2001/09/05 S.Sano End

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

Private Function StrNoNull(s$) As String
    If Trim$(s) = vbNullString Then
        StrNoNull = " "
    Else
        StrNoNull = Trim$(s)
    End If
End Function

'------------------------------------------------
' DB�A�N�Z�X�֐��i�폜�ҁj
'------------------------------------------------
'�T�v      :�e�[�u���uTBCMB011�v�̏����ɂ��������R�[�h���폜
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :PGID�@        ,O  ,String       ,�폜PG-ID
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/10/05 �쐬�@���{
Public Function DBDRV_cmbc001d_Del(PGID As String) As FUNCTION_RETURN

    Dim sql     As String                   'SQL�S��

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001d_SQL.bas -- Function DBDRV_cmbc001d_Del"
    
    sql = "delete "
    sql = sql & "from "
    sql = sql & "TBCMB011 "
    sql = sql & "where "
    sql = sql & "trim(PGID)='" & Trim(PGID) & "'"
    
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_cmbc001d_Del = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_cmbc001d_Del = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_cmbc001d_Del = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


' ������������e�i���X

'�T�v      :������������e�i���X ��������X�V�^�}���p�c�a�h���C�o
'���Ұ��@�@:�ϐ���         ,IO ,�^               ,����
'      �@�@:sMkCondNo�@�@�@,I  ,String         �@,���������
'      �@�@:pMkOld   �@�@�@,I  ,typ_TBCMB012   �@,��������I���W�i��
'      �@�@:pMkNew   �@�@�@,I  ,typ_TBCMB012   �@,�������
'      �@�@:�߂�l         ,O  ,FUNCTION_RETURN�@,�������ݐ���
'����      :
'����      :2001/07/30 ���{ �쐬
Public Function DBDRV_scmzc_fcmbc001e_UpdInsMkCond(sMkCondNo As String, pMkOld() As typ_TBCMB012, pMkNew As typ_TBCMB012) As FUNCTION_RETURN

    Dim sql As String
    Dim bFlag As Boolean
    Dim i As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001e_SQL.bas -- Function DBDRV_scmzc_fcmbc001e_UpdInsMkCond"

    With pMkNew
        bFlag = False
        For i = 1 To UBound(pMkOld)
            If RTrim$(pMkOld(i).MKCONDNO) = RTrim$(sMkCondNo) Then
                bFlag = True
                Exit For
            End If
        Next i

        If bFlag = True Then
            '' ��������̍X�V
            sql = "update TBCMB012 set "
            sql = sql & "MKCONDNO='" & .MKCONDNO & "', "    ' �������No.
            sql = sql & "MODEL='" & .MODEL & "', "          ' �@��
            sql = sql & "RTBSIZE='" & .RTBSIZE & "', "      ' ���c�{�T�C�Y
            sql = sql & "CHARGE='" & .CHARGE & "', "        ' �`���[�W��
            sql = sql & "HZTYPE='" & .HZTYPE & "', "        ' HZ�^�C�v
            sql = sql & "UPSPDTYP='" & .UPSPDTYP & "', "    ' ���グ���x�^�C�v
            sql = sql & "MAGTYPE='" & .MAGTYPE & "', "      ' ����^�C�v
            sql = sql & "USECLS='0', "                      ' �g�p�敪
            sql = sql & "TSTAFFID='" & .TSTAFFID & "', "    ' �o�^�Ј�ID
            sql = sql & "REGDATE=sysdate, "                 ' �o�^���t
            sql = sql & "KSTAFFID='" & .KSTAFFID & "', "    ' �X�V�Ј�ID
            sql = sql & "UPDDATE=sysdate, "                 ' �X�V���t
            sql = sql & "SENDFLAG='0', "                    ' ���M�t���O
            sql = sql & "SENDDATE=sysdate"                  ' ���M����
            sql = sql & " where rtrim(MKCONDNO)='" & RTrim$(sMkCondNo) & "'"
        Else
            '' ��������̑}��
            sql = "insert into TBCMB012 ("
            sql = sql & "MKCONDNO, "        ' �������No.
            sql = sql & "MODEL, "           ' �@��
            sql = sql & "RTBSIZE, "         ' ���c�{�T�C�Y
            sql = sql & "CHARGE, "          ' �`���[�W��
            sql = sql & "HZTYPE, "          ' HZ�^�C�v
            sql = sql & "UPSPDTYP, "        ' ���グ���x�^�C�v
            sql = sql & "MAGTYPE, "         ' ����^�C�v
            sql = sql & "USECLS, "          ' �g�p�敪
            sql = sql & "TSTAFFID, "        ' �o�^�Ј�ID
            sql = sql & "REGDATE, "         ' �o�^���t
            sql = sql & "KSTAFFID, "        ' �X�V�Ј�ID
            sql = sql & "UPDDATE, "         ' �X�V���t
            sql = sql & "SENDFLAG, "        ' ���M�t���O
            sql = sql & "SENDDATE)"         ' ���M����
            sql = sql & " values ('"
            sql = sql & .MKCONDNO & "', '"  ' �������No.
            sql = sql & .MODEL & "', '"     ' �@��
            sql = sql & .RTBSIZE & "', '"   ' ���c�{�T�C�Y
            sql = sql & .CHARGE & "', '"    ' �`���[�W��
            sql = sql & .HZTYPE & "', '"    ' HZ�^�C�v
            sql = sql & .UPSPDTYP & "', '"  ' ���グ���x�^�C�v
            sql = sql & .MAGTYPE & "', "    ' ����^�C�v
            sql = sql & "'0', '"            ' �g�p�敪
            sql = sql & .TSTAFFID & "', "   ' �o�^�Ј�ID
            sql = sql & "sysdate, '"        ' �o�^���t
            sql = sql & .KSTAFFID & "', "   ' �X�V�Ј�ID
            sql = sql & "sysdate, "         ' �X�V���t
            sql = sql & "'0', "             ' ���M�t���O
            sql = sql & "sysdate)"          ' ���M����
        End If
    End With
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_scmzc_fcmbc001e_UpdInsMkCond = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmbc001e_UpdInsMkCond = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmbc001e_UpdInsMkCond = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :������������e�i���X ��������폜�p�c�a�h���C�o
'���Ұ��@�@:�ϐ���         ,IO ,�^               ,����
'      �@�@:sMkCondNo�@�@�@,I  ,String         �@,���������
'      �@�@:�߂�l         ,O  ,FUNCTION_RETURN�@,�������ݐ���
'����      :
'����      :2001/07/30 ���{ �쐬
Public Function DBDRV_scmzc_fcmbc001e_DelMkCond(sMkCondNo As String) As FUNCTION_RETURN

    Dim sql As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001e_SQL.bas -- Function DBDRV_scmzc_fcmbc001e_DelMkCond"

    '' ��������̍폜
    sql = "delete TBCMB012 where rtrim(MKCONDNO)='" & RTrim$(sMkCondNo) & "'"
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_scmzc_fcmbc001e_DelMkCond = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmbc001e_DelMkCond = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmbc001e_DelMkCond = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :������������e�i���X �������PG-ID�Ή��X�V�^�}���p�c�a�h���C�o
'���Ұ��@�@:�ϐ���         ,IO ,�^               ,����
'      �@�@:sMkCondNo�@�@�@,I  ,String         �@,���������
'      �@�@:pPGIDOld �@�@�@,I  ,typ_TBCMB013   �@,�������PG-ID�Ή��I���W�i��
'      �@�@:pPGIDNew �@�@�@,I  ,typ_TBCMB013   �@,�������PG-ID�Ή�
'      �@�@:�߂�l         ,O  ,FUNCTION_RETURN�@,�������ݐ���
'����      :
'����      :2001/07/30 ���{ �쐬
Public Function DBDRV_scmzc_fcmbc001e_UpdInsPGIDMng(sMkCondNo As String, pPGIDOld() As typ_TBCMB013, pPGIDNew() As typ_TBCMB013) As FUNCTION_RETURN

    Dim sql As String
    Dim bFlag As Boolean
    Dim i As Long
    Dim j As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001e_SQL.bas -- Function DBDRV_scmzc_fcmbc001e_InsPGIDMng"

    For i = 1 To UBound(pPGIDNew)
        With pPGIDNew(i)
            bFlag = False
            For j = 1 To UBound(pPGIDOld)
                If RTrim$(pPGIDOld(j).MKCONDNO) = RTrim$(sMkCondNo) And _
                   RTrim$(pPGIDOld(j).PGIDNO) = RTrim$(.PGIDNO) Then
                    bFlag = True
                    Exit For
                End If
            Next j

            If bFlag = True Then
                '' �������PG-ID�Ή��̍X�V
                sql = "update TBCMB013 set "
                sql = sql & "MKCONDNO='" & .MKCONDNO & "', "    ' �������No.
                sql = sql & "PGIDNO='" & .PGIDNO & "', "        ' PG-IDNo
                sql = sql & "TSTAFFID='" & .TSTAFFID & "', "    ' �o�^�Ј�ID
                sql = sql & "REGDATE=sysdate, "                 ' �o�^���t
                sql = sql & "KSTAFFID='" & .KSTAFFID & "', "    ' �X�V�Ј�ID
                sql = sql & "UPDDATE=sysdate, "                 ' �X�V���t
                sql = sql & "SENDFLAG='0', "                    ' ���M�t���O
                sql = sql & "SENDDATE=sysdate"                  ' ���M���t
                sql = sql & " where rtrim(MKCONDNO)='" & RTrim$(sMkCondNo) & "'"
                sql = sql & " and rtrim(PGIDNO)='" & RTrim$(.PGIDNO) & "'"
            Else
                '' �������PG-ID�Ή��̑}��
                sql = "insert into TBCMB013 ("
                sql = sql & "MKCONDNO, "        ' �������No.
                sql = sql & "PGIDNO, "          ' PG-IDNo
                sql = sql & "TSTAFFID, "        ' �o�^�Ј�ID
                sql = sql & "REGDATE, "         ' �o�^���t
                sql = sql & "KSTAFFID, "        ' �X�V�Ј�ID
                sql = sql & "UPDDATE, "         ' �X�V���t
                sql = sql & "SENDFLAG, "        ' ���M�t���O
                sql = sql & "SENDDATE)"         ' ���M���t
                sql = sql & " values ('"
                sql = sql & .MKCONDNO & "', '"  ' �������No.
                sql = sql & .PGIDNO & "', '"    ' PG-IDNo
                sql = sql & .TSTAFFID & "', "   ' �o�^�Ј�ID
                sql = sql & "sysdate, '"        ' �o�^���t
                sql = sql & .KSTAFFID & "', "   ' �X�V�Ј�ID
                sql = sql & "sysdate, "         ' �X�V���t
                sql = sql & "'0', "             ' ���M�t���O
                sql = sql & "sysdate)"          ' ���M���t
            End If
        End With
        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_scmzc_fcmbc001e_UpdInsPGIDMng = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    DBDRV_scmzc_fcmbc001e_UpdInsPGIDMng = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmbc001e_UpdInsPGIDMng = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :������������e�i���X �������PG-ID�Ή��폜�p�c�a�h���C�o
'���Ұ��@�@:�ϐ���         ,IO ,�^               ,����
'      �@�@:sMkCondNo�@�@�@,I  ,String         �@,���������
'      �@�@:sPGIDNo  �@�@�@,I  ,String         �@,PG-ID��
'      �@�@:�߂�l         ,O  ,FUNCTION_RETURN�@,�������ݐ���
'����      :
'����      :2001/07/30 ���{ �쐬
Public Function DBDRV_scmzc_fcmbc001e_DelPGIDMng(sMkCondNo As String, sPGIDNo As String) As FUNCTION_RETURN

    Dim sql As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc001e_SQL.bas -- Function DBDRV_scmzc_fcmbc001e_DelPGIDMng"

    '' �������PG-ID�Ή��̍폜
    sql = "delete TBCMB013 where rtrim(MKCONDNO)='" & RTrim$(sMkCondNo) & "'"
    If RTrim$(sPGIDNo) <> "" Then
        sql = sql & " and rtrim(PGIDNO)='" & RTrim$(sPGIDNo) & "'"
    End If
    If OraDB.ExecuteSQL(sql) < 0 Then
        DBDRV_scmzc_fcmbc001e_DelPGIDMng = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmbc001e_DelPGIDMng = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmbc001e_DelPGIDMng = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMB011�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMB011 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCMB011_SQL.bas���ړ�)
Public Function DBDRV_GetTBCMB011(records() As typ_TBCMB011, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select PGID, HZPART, HZPTRN, SPACER, UPRING, CHARGE, RTBPOS, RTBSIZE, GAP, UPDM, UPLENGTH, UPRC, RFRNEED, UPSPIN," & _
              " DOWNSPIN, ROPRESS, ARUGON, AIMOIMIN, AIMOIMAX, HCCLASS, HC, AVEUPSPD, UPCNTL, BTMSHAPE, MAGSTR, MAGPOS, CONDGRT," & _
              " MODEL, UPMETHOD, UPCLASS, UPNUM, OPETIME, WTRCOOL, PGID2, RCPT1, RCPT2, RCPT3, RCPT4, RCPT5, CNTL1, CNTL2," & _
              " CNTL3, CNTL4, CNTL5, CNTL6, CNTL7, CNTL8, CNTL9, CNTL10, CNTL11, CNTL12, CNTL13, CNTL14, CNTL15, RUNCOND1," & _
              " RUNCOND2, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMB011"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB011 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .PGID = rs("PGID")               ' PG-ID
            .HZPART = rs("HZPART")           ' HZ�p�[�c
            .HZPTRN = rs("HZPTRN")           ' HZ�p�^�[��
            .SPACER = rs("SPACER")           ' �X�y�[�T
            .UPRING = rs("UPRING")           ' �A�b�p�[�����O
            .CHARGE = rs("CHARGE")           ' �`���[�W��
            .RTBPOS = rs("RTBPOS")           ' ���c�{�ʒu
            .RTBSIZE = rs("RTBSIZE")         ' ���c�{�T�C�Y
            .GAP = rs("GAP")                 ' �M���b�v
            .UPDM = rs("UPDM")               ' ���㒼�a
            .UPLENGTH = rs("UPLENGTH")       ' ���㒷�i�S���j
            .UPRC = rs("UPRC")               ' ����iRC�j
            .RFRNEED = rs("RFRNEED")         ' ���t���N�^�v��
            .UPSPIN = rs("UPSPIN")           ' �㎲��]��
            .DOWNSPIN = rs("DOWNSPIN")       ' ������]��
            .ROPRESS = rs("ROPRESS")         ' �F����
            .ARUGON = rs("ARUGON")           ' �A���S����
            .AIMOIMIN = rs("AIMOIMIN")       ' �˂炢Oi�iMIN)
            .AIMOIMAX = rs("AIMOIMAX")       ' �˂炢Oi�iMAX)
            .HCCLASS = rs("HCCLASS")         ' HC���
            .HC = rs("HC")                   ' HC
            .AVEUPSPD = rs("AVEUPSPD")       ' ���ψ��㑬�x
            .UPCNTL = rs("UPCNTL")           ' ���㐧��
            .BTMSHAPE = rs("BTMSHAPE")       ' �{�g���`��
            .MAGSTR = rs("MAGSTR")           ' ���ꋭ�x
            .MAGPOS = rs("MAGPOS")           ' ����ʒu
            .CONDGRT = rs("CONDGRT")         ' �����ۏؓo�^
            .MODEL = rs("MODEL")             ' �@��
            .UPMETHOD = rs("UPMETHOD")       ' ������@
            .UPCLASS = rs("UPCLASS")         ' ����敪
            .UPNUM = rs("UPNUM")             ' ����{��
            .OPETIME = rs("OPETIME")         ' �^�]����
            .WTRCOOL = rs("WTRCOOL")         ' ����Ǘv��
            .PGID2 = rs("PGID2")             ' PG-ID�i��{���j
            .RCPT1 = rs("RCPT1")             ' �Ή����V�sNo�iT1)
            .RCPT2 = rs("RCPT2")             ' �Ή����V�sNo�iT2)
            .RCPT3 = rs("RCPT3")             ' �Ή����V�sNo�iT3)
            .RCPT4 = rs("RCPT4")             ' �Ή����V�sNo�iT4)
            .RCPT5 = rs("RCPT5")             ' �Ή����V�sNo�iT5)
            .CNTL1 = rs("CNTL1")             ' �������ځi1�j
            .CNTL2 = rs("CNTL2")             ' �������ځi2�j
            .CNTL3 = rs("CNTL3")             ' �������ځi3�j
            .CNTL4 = rs("CNTL4")             ' �������ځi4�j
            .CNTL5 = rs("CNTL5")             ' �������ځi5�j
            .CNTL6 = rs("CNTL6")             ' �������ځi6�j
            .CNTL7 = rs("CNTL7")             ' �������ځi7�j
            .CNTL8 = rs("CNTL8")             ' �������ځi8�j
            .CNTL9 = rs("CNTL9")             ' �������ځi9�j
            .CNTL10 = rs("CNTL10")           ' �������ځi10�j
            .CNTL11 = rs("CNTL11")           ' �������ځi11�j
            .CNTL12 = rs("CNTL12")           ' �������ځi12�j
            .CNTL13 = rs("CNTL13")           ' �������ځi13�j
            .CNTL14 = rs("CNTL14")           ' �������ځi14�j
            .CNTL15 = rs("CNTL15")           ' �������ځi15�j
            .RUNCOND1 = rs("RUNCOND1")       ' �^�]�����P
            .RUNCOND2 = rs("RUNCOND2")       ' �^�]�����Q
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

    DBDRV_GetTBCMB011 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMB012�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMB012 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCMB012_SQL.bas���ړ�)
Public Function DBDRV_GetTBCMB012(records() As typ_TBCMB012, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select MKCONDNO, MODEL, RTBSIZE, CHARGE, HZTYPE, UPSPDTYP, MAGTYPE, USECLS, TSTAFFID, REGDATE, KSTAFFID, UPDDATE," & _
              " SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMB012"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB012 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .MKCONDNO = rs("MKCONDNO")       ' �������No.
            .MODEL = rs("MODEL")             ' �@��
            .RTBSIZE = rs("RTBSIZE")         ' ���c�{�T�C�Y
            .CHARGE = rs("CHARGE")           ' �`���[�W��
            .HZTYPE = rs("HZTYPE")           ' HZ�^�C�v
            .UPSPDTYP = rs("UPSPDTYP")       ' ���グ���x�^�C�v
            .MAGTYPE = rs("MAGTYPE")         ' ����^�C�v
            .USECLS = rs("USECLS")           ' �g�p�敪
            .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M����
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMB012 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMB013�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMB013 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCMB013_SQL.bas���ړ�)
Public Function DBDRV_GetTBCMB013(records() As typ_TBCMB013, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select MKCONDNO, PGIDNO, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMB013"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB013 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .MKCONDNO = rs("MKCONDNO")       ' �������No.
            .PGIDNO = rs("PGIDNO")           ' PG-IDNo
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

    DBDRV_GetTBCMB013 = FUNCTION_RETURN_SUCCESS
End Function

Public Function DBDRV_Syounin_Disp(records() As s_cmzcF_cmfc001b_Disp) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "Syounin_SQL.bas -- Function DBDRV_Syounin_Disp"
    
    ''�����F���i�d�l�Ǘ��R�[�h�擾�i�i�ԁA�d�l�o�^�˗��ԍ��A�X�V���t�j
    ''�������A��������t�^����ɂ��郌�R�[�h�͏���
    'sql = "select hinban||ltrim(to_char(mnorevno,'00'))||factory||opecond as hinban12, HMGSTRRNO, REGDATE " & _
          "From tbcme018 " & _
          "where (opecond > '1') and (nvl(synflag, ' ') = '0') and " & _
          "(hinban||mnorevno||factory) not in (select hinban||mnorevno||factory from tbcme031) and " & _
          "(hinban||mnorevno||factory) in (select hinban||mnorevno||factory from tbcme032) "
    'sql = "select hinban||ltrim(to_char(mnorevno,'00'))||factory||opecond as hinban12, HMGSTRRNO, REGDATE " & _
          "From tbcme018 " & _
          "where (opecond > '1') and (nvl(synflag, ' ') = '0') and " & _
          "(hinban||mnorevno||factory) not in (select hinban||mnorevno||factory from tbcme031)  "
    'sql = "select A.hinban||ltrim(to_char(A.mnorevno,'00'))||A.factory||A.opecond as hinban12, A.HMGSTRRNO, A.REGDATE, B.SENDFLAG, B.SENDDATE " & _
          "From tbcme018 A , tbcme032 B " & _
          "where (A.opecond > '1') and (nvl(A.synflag, ' ') = '0') and " & _
          "(A.hinban||A.mnorevno||A.factory = B.hinban(+)||B.mnorevno(+)||B.factory(+) ) and " & _
          "(A.hinban||A.mnorevno||A.factory) not in (select hinban||mnorevno||factory from tbcme031) "
    'sql = "select A.hinban||ltrim(to_char(A.mnorevno,'00'))||A.factory||A.opecond as hinban12, A.HMGSTRRNO, A.REGDATE , B.SYNFLAG, B.UPDDATE as TOUROKU " & _
          "From tbcme018 A , tbcme036 B " & _
          "where (A.opecond='1')  and (B.opecond(+) = '1') and " & _
          "(A.hinban||A.mnorevno||A.factory) = (B.hinban(+)||B.mnorevno(+)||B.factory(+)) and " & _
          "(A.hinban||A.mnorevno||A.factory) not in (select D.hinban||D.mnorevno||D.factory from tbcme030 D) and " & _
          "(A.hinban||A.mnorevno||A.factory) not in (select C.hinban||C.mnorevno||C.factory from tbcme031 C)"
    sql = "select A.hinban||ltrim(to_char(A.mnorevno,'00'))||A.factory||A.opecond as hinban12, A.HMGSTRRNO, A.REGDATE , B.UPDDATE as TOUROKU , C.SENDFLAG, C.SENDDATE " & _
          "From tbcme018 A , tbcme036 B , tbcme032 C " & _
          "where (nvl(A.synflag, ' ') = '0') and A.opecond = B.opecond(+) and " & _
          "(A.hinban||A.mnorevno||A.factory) = (B.hinban(+)||B.mnorevno(+)||B.factory(+)) and " & _
          "(A.hinban||A.mnorevno||A.factory = C.hinban(+)||C.mnorevno(+)||C.factory(+) ) and " & _
          "(A.hinban||A.mnorevno||A.factory) not in (select hinban||mnorevno||factory from tbcme031 )"
          Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_Syounin_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .Hinban12 = rs("HINBAN12")       ' �i��
            .HMGSTRRNO = rs("HMGSTRRNO")   ' �i�Ǘ��d�l�o�^�˗��ԍ�
            .REGDATE = rs("REGDATE")       ' �o�^���t
            If IsNull(rs("TOUROKU")) = False Then .TOUROKU = rs("TOUROKU")
            If IsNull(rs("SENDDATE")) = False Then .SENDDATE = rs("SENDDATE")
            If IsNull(rs("SENDFLAG")) = False Then .SENDFLAG = rs("SENDFLAG")
        End With
        rs.MoveNext
    Next
    rs.Close


    DBDRV_Syounin_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :[���s]���f�[�^��������
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :IraiNo        ,I  ,String    ,�d�l�o�^�˗��ԍ�
'          :SXLMCNO       ,I  ,String    ,SXL�������(�d�l��)
'          :WFMCNO        ,I  ,String    ,WF�������(�d�l��)
'          :Hinban12      ,I  ,String    ,12���i��
'          :SJokenNo      ,I  ,String    ,��������ԍ�
'          :Hikiage       ,I  ,String    ,������@
'          :StaffID       ,I  ,String    ,�S���҃R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN,�����̐���
'����      :
'����      :2001/06/08 �쐬  �쑺
'Public Function DBDRV_s_cmzcF_cmfc003c_Exec(ByVal IraiNo$, ByVal Sxluchigawa As c_cmzcrec, Hinban12$, ByVal StaffID$, ByVal Snote$, ByVal Jnote$) As FUNCTION_RETURN
Public Function DBDRV_s_cmzcF_cmfc003c_Exec(ByVal IraiNo$, ByVal Sxluchigawa As c_cmzcrec, Hinban12$, ByVal StaffID$) As FUNCTION_RETURN
Dim sql_top As String
Dim sql_sel As String
Dim sql As String
Dim fld As OraField
Dim rs As OraDynaset
Dim fullHinban As tFullHinban


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmfc003c_SQL.bas -- Function DBDRV_s_cmzcF_cmfc003c_Exec"
    ''�g�����U�N�V�����J�n
    Debug.Print "BeginTrans ======="
    OraDB.BeginTrans
    
    If Len(Hinban12) <> 12 Then
        DBDRV_s_cmzcF_cmfc003c_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    With fullHinban
        .hinban = Left$(Hinban12, 8)
        .MNOREVNO = val(Mid$(Hinban12, 9, 2))
        .FACTORY = Mid$(Hinban12, 11)
        .OPECOND = Right$(Hinban12, 1)
    End With
    
    sql = "insert into TBCME030 " & _
          "(HINBAN, MNOREVNO, FACTORY, OPECOND, SSXLIFTW, IFKBN, SYORIKBN, SPECRRNO, SXLMCNO, WFMCNO, " & _
          "TOPREG, BTMSPRT, MCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE, QASENDFLAG) " & _
          "values ("
    sql = sql & "'" & fullHinban.hinban & "', "     ' �i��
    sql = sql & fullHinban.MNOREVNO & ", "          ' ���i�ԍ������ԍ�
    sql = sql & "'" & fullHinban.FACTORY & "', "    ' �H��
    'sql = sql & "'" & fullHinban.OPECOND & "', "    ' ���Ə���
    sql = sql & "'1', "    ' ���Ə���
    'sql = sql & "'" & Hikiage & "', "               ' ���r�w������@
    sql = sql & "'" & Sxluchigawa("SSXLIFTW") & "', "               ' ���r�w������@
    sql = sql & "' ', "                             ' �h�^�e�敪
    sql = sql & "' ', "                             ' �����敪
    sql = sql & "'" & IraiNo & "', "                ' �d�l�o�^�˗��ԍ�
    sql = sql & "'" & Sxluchigawa("SXLMCNO") & "', "    ' �r�w�k��������ԍ�
    sql = sql & "'" & Sxluchigawa("WFMCNO") & "', "     ' �v�e��������ԍ�
    sql = sql & "'" & Sxluchigawa("TOPREG") & "', "     ' TOP�K��          04/07/09
    sql = sql & "'" & Sxluchigawa("BTMSPRT") & "', "    ' �{�g���͏o�K��    04/07/09
    sql = sql & "'" & Sxluchigawa("MCNO") & "', "       ' �������    04/09/01
    sql = sql & "'" & StaffID & "', "               ' �Ј�ID
    sql = sql & "SYSDATE, "                         ' �o�^���t
    sql = sql & "SYSDATE, "                         ' �X�V���t
    sql = sql & "'0', "                             ' ���M�t���O
    sql = sql & "SYSDATE, "                         ' ���M���t
    sql = sql & "'0' "                              ' �i�����M�t���O
    sql = sql & ")"
    If 0 >= OraDB.ExecuteSQL(sql) Then
        GoTo proc_err
    End If
    ''�i�ԃf�[�^�ɐ������No����������(���d�l�ł��郊�r�W�����P�������e���r�W�����Ɂj<---�o�^�Ɉړ�
    'sql = "update TBCME018 set " & _
          "MCNO = '" & SJokenNo & "' " & _
          "where " & _
          "(HINBAN = '" & fullHinban.HINBAN & "') and " & _
          "(MNOREVNO = " & fullHinban.MNOREVNO & ") and " & _
          "(FACTORY = '" & fullHinban.FACTORY & "')"
    'OraDB.ExecuteSQL sql
    'If 0 >= OraDB.ExecuteSQL(sql) Then
    '    GoTo proc_err
    'End If
    
    ''���L�����̍X�V�͏��F�����ōs���̂ō폜
    'sql = "update TBCME036 set " & _
    '      "UPDDATE = sysdate ," & _
    '      "SNOTE = '" & Snote & "' ," & _
    '      "JNOTE = '" & Jnote & "' " & _
    '      "where " & _
    '      "(HINBAN = '" & fullHinban.HINBAN & "') and " & _
    '      "(MNOREVNO = " & fullHinban.MNOREVNO & ") and " & _
    '      "(FACTORY = '" & fullHinban.FACTORY & "') and " & _
    '      "(OPECOND = '" & fullHinban.OPECOND & "')"
    'If 0 >= OraDB.ExecuteSQL(sql) Then
    '    GoTo proc_err
    'End If
    
    DBDRV_s_cmzcF_cmfc003c_Exec = FUNCTION_RETURN_SUCCESS
    
    ''����I���Ȃ�R�~�b�g
    Debug.Print "CommitTrans ======="
    OraDB.CommitTrans

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "==== Error SQL ===="
    Debug.Print sql
    gErr.HandleError
    ''�G���[���̓��[���o�b�N
    Debug.Print "RollBack ======="
    OraDB.Rollback
    Resume proc_exit
End Function


'�T�v      :[���s]���f�[�^��������
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :IraiNo        ,I  ,String    ,�d�l�o�^�˗��ԍ�
'          :SXLMCNO       ,I  ,String    ,SXL�������(�d�l��)
'          :WFMCNO        ,I  ,String    ,WF�������(�d�l��)
'          :Hinban12      ,I  ,String    ,12���i��
'          :SJokenNo      ,I  ,String    ,��������ԍ�
'          :Hikiage       ,I  ,String    ,������@
'          :StaffID       ,I  ,String    ,�S���҃R�[�h
'          :�߂�l        ,O  ,FUNCTION_RETURN,�����̐���
'����      :
'����      :2001/06/08 �쐬  �쑺
Public Function DBDRV_TOUROKU_Exec(ByVal IraiNo$, ByVal Sxl_1 As c_cmzcrec, Sxluchigawa As c_cmzcrec, Hinban12$, _
                                            ByVal SJokenNo$, ByVal Hikiage$, ByVal StaffID$, ByVal Snote$, ByVal Jnote$) As FUNCTION_RETURN
Dim sql_top As String
Dim sql_sel As String
Dim sql As String
Dim fld As OraField
Dim rs As OraDynaset
Dim fullHinban As tFullHinban


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "TOUROKU_SQL.bas -- Function DBDRV_TOUROKU_Exec"
    ''�g�����U�N�V�����J�n
    Debug.Print "BeginTrans ======="
    OraDB.BeginTrans
    
    If Len(Hinban12) <> 12 Then
        DBDRV_TOUROKU_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    With fullHinban
        .hinban = Left$(Hinban12, 8)
        .MNOREVNO = val(Mid$(Hinban12, 9, 2))
        .FACTORY = Mid$(Hinban12, 11)
        .OPECOND = Right$(Hinban12, 1)
    End With
    ''�i�ԃf�[�^�ɐ������No����������
    sql = "update TBCME018 set " & _
          "UPDDATE = SYSDATE, " & _
          "MCNO = '" & SJokenNo & "' " & _
          "where " & _
          "(HINBAN = '" & fullHinban.hinban & "') and " & _
          "(MNOREVNO = " & fullHinban.MNOREVNO & ") and " & _
          "(FACTORY = '" & fullHinban.FACTORY & "')"
    Debug.Print "ExecuteSQL ==========="
    Debug.Print sql
    If 0 >= OraDB.ExecuteSQL(sql) Then
        GoTo proc_err
    End If
    With Sxluchigawa
        ''�S�Ẵ��r�W�����̐������No������������
        sql = "update TBCME036 set " & _
              "UPDDATE = SYSDATE ," & _
              "MCNO = '" & SJokenNo & "' ," & _
              "SSXLIFTW = '" & Hikiage & "' " & _
              "where " & _
              "(HINBAN = '" & fullHinban.hinban & "') and " & _
              "(MNOREVNO = " & fullHinban.MNOREVNO & ") and " & _
              "(FACTORY = '" & fullHinban.FACTORY & "') "
              '"(OPECOND = '" & fullHinban.OPECOND & "')"
        If 0 >= OraDB.ExecuteSQL(sql) Then
            .Fields("HINBAN") = fullHinban.hinban
            .Fields("MNOREVNO") = fullHinban.MNOREVNO
            .Fields("FACTORY") = fullHinban.FACTORY
            '.Fields("OPECOND") = "1"
            .Fields("OPECOND") = fullHinban.OPECOND
            .Fields("SPECRRNO") = IraiNo
            .Fields("SXLMCNO") = Sxl_1("SXLMCNO")
            .Fields("WFMCNO") = Sxl_1("WFMCNO")
            .Fields("SNOTE") = Snote
            .Fields("JNOTE") = Jnote
            .Fields("STAFFID") = StaffID
            .Fields("MCNO") = SJokenNo
            .Fields("SSXLIFTW") = Hikiage
            sql = .SqlInsert
            If 0 >= OraDB.ExecuteSQL(sql) Then
                    GoTo proc_err
            End If
        End If
        Debug.Print "ExecuteSQL ==========="
        Debug.Print sql
        ''���L�������X�V����
        'sql = "update TBCME036 set " & _
        '      "SNOTE = '" & Snote & "' ," & _
        '      "JNOTE = '" & Jnote & "' " & _
        '      "where " & _
        '      "(HINBAN = '" & fullHinban.HINBAN & "') and " & _
        '      "(MNOREVNO = " & fullHinban.MNOREVNO & ") and " & _
        '      "(FACTORY = '" & fullHinban.FACTORY & "') and " & _
        '      "(OPECOND = '" & fullHinban.OPECOND & "')"
        'Debug.Print "ExecuteSQL ==========="
        'Debug.Print sql
        'If 0 >= OraDB.ExecuteSQL(sql) Then
        '    GoTo proc_err
        'End If
        
    End With
    
    DBDRV_TOUROKU_Exec = FUNCTION_RETURN_SUCCESS
    
    ''����I���Ȃ�R�~�b�g
    Debug.Print "CommitTrans ======="
    OraDB.CommitTrans

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "==== Error SQL ===="
    Debug.Print sql
    gErr.HandleError
    ''�G���[���̓��[���o�b�N
    Debug.Print "RollBack ======="
    OraDB.Rollback
    Resume proc_exit
End Function

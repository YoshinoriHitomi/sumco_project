Attribute VB_Name = "s_cmbc001c_SQL"
Option Explicit

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
Public Function DBDRV_s_cmzcF_cmfc001c_DispSXL(targetHinban As tFullHinban, SxlKokyaku_1 As c_cmzcrec, SxlKokyaku_2 As c_cmzcrec, SxlKokyaku_3 As c_cmzcrec, Sxl_1 As c_cmzcrec, Sxl_2 As c_cmzcrec, Sxl_3 As c_cmzcrec, WfKokyaku_2 As c_cmzcrec, WfKokyaku_8 As c_cmzcrec, SxlUchigawa As c_cmzcrec) As FUNCTION_RETURN
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
    Set SxlUchigawa = New c_cmzcrec
    
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
    sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND," & _
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
    sqlBase = "Select HSXTMMAX," & _
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
    sqlBase = sqlBase & " HSXONSPH||HSXONSPT||HSXONSPI as HSXONSP," & _
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
                " HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU as HSXOS2K "
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
    sqlBase = "Select HSXDENKU," & _
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
    sqlBase = sqlBase & " HSXOF1KM||HSXOF1KI||HSXOF1KH||HSXOF1KS as HSXOF1K," & _
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
    sqlBase = sqlBase & " HSXBM1SZ, HSXBM1AN, HSXBM1AX," & _
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
              " HSXBM3KM||HSXBM3KI||HSXBM3KH||HSXBM3KS as HSXBM3K "
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
    If DispSXL_GetData(TbName, sqlBase & sqlWhere, SxlUchigawa) = FUNCTION_RETURN_FAILURE Then
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
Public Function DBDRV_s_cmzcF_cmfc001c_Exec(ByVal IraiNo$, ByVal SXLMCNO$, ByVal WFMCNO$, ByVal Hinban12$, ByVal SJokenNo$, ByVal Hikiage$, ByVal StaffID$) As FUNCTION_RETURN
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
        .MNOREVNO = Val(Mid$(Hinban12, 9, 2))
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


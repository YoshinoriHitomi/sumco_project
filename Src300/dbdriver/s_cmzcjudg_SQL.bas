Attribute VB_Name = "s_cmzcjudg_SQL"
Option Explicit

'����\����
Public Type TYPE_JUDG
    Guarantee   As Guarantee                ''�i���ۏ؏��\����
    SpecMin     As Double                   ''�iSXL***����
    SpecMax     As Double                   ''�iSXL***���
    JudgData    As Double                   ''�G���[�R�[�h
    Judg        As Boolean                  ''�I�v�V����������
End Type
'�����ۏؔ���\����
Public Type TYPE_CRYJUDG
    Den     As TYPE_JUDG                    ''Den
    Ldl     As TYPE_JUDG                    ''L/DL
    Dvd2    As TYPE_JUDG                    ''Dvd2
    Lt      As TYPE_JUDG                    ''���C�t�^�C��
    Cs      As TYPE_JUDG                    ''Cs
    JFDen   As String * 1                   ''�iSXDen�����L��
    JFLdl   As String * 1                   ''�iSXL/DL�����L��
    JFDvd2  As String * 1                   ''�iSXDVD2�����L��
End Type
Public Type TYPE_FRIKAE
    HIN As tFullHinban
    blkID As String
End Type



'�T�v      :���i�d�lSXL�f�[�^�̎擾�h���C�o
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'      �@�@:HIN       �@�@�@, O  , tFullHinban    �@, �t���i�ԍ\����
'          :Gd             ,I   ,C_GD              ,����GD����\����
'          :Cs             ,I   ,C_Cs              ,����Cs����\����
'          :Lt             ,I   ,C_LT              ,�������C�t�^�C������\����
'      �@�@:�߂�l          , O  , FUNCTION_RETURN�@, �ǂݍ��݂̐���
'����      :GD/LT/Cs�̔���ɕK�v�Ȏd�l�����擾����
'����      :2002/03/14 ���� �M�� �쐬
Public Function scmzc_getSXLGuarantee(HIN As tFullHinban, GD() As C_GD, Cs() As C_Cs, Lt() As C_LT) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcjudg_SQL.bas -- Function scmzc_getSXLGuarantee"
    scmzc_getSXLGuarantee = FUNCTION_RETURN_FAILURE

    '' ���i�d�l�̎擾
    sql = "select " & _
          "E019HSXLTMIN, E019HSXLTMAX, E019HSXLTSPH, E019HSXLTSPT, E019HSXLTSPI, E019HSXLTHWT, E019HSXLTHWS, E019HSXCNSPH, " & _
          "E019HSXCNSPT, E019HSXCNSPI, E019HSXCNHWT, E019HSXCNHWS, E019HSXCNMIN, E019HSXCNMAX, E020HSXDENKU, E020HSXDENMX, " & _
          "E020HSXDENMN, E020HSXDENHT, E020HSXDENHS, E020HSXDVDKU, E020HSXDVDMXN, E020HSXDVDMNN, E020HSXDVDHT, E020HSXDVDHS, " & _
          "E020HSXLDLKU, E020HSXLDLMX, E020HSXLDLMN, E020HSXLDLHT, E020HSXLDLHS " & _
          " from VECME001" & _
          " where E018HINBAN='" & HIN.hinban & "' and E018MNOREVNO=" & HIN.mnorevno & _
          " and E018FACTORY='" & HIN.factory & "' and E018OPECOND='" & HIN.opecond & "'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        GoTo proc_exit
    End If

    Lt(0).SpecLtMin = fncNullCheck(rs("E019HSXLTMIN"))
    Lt(0).SpecLtMax = fncNullCheck(rs("E019HSXLTMAX"))
    Lt(0).GuaranteeLt.cMeth = rs("E019HSXLTSPH")
    Lt(0).GuaranteeLt.cCount = rs("E019HSXLTSPT")
    Lt(0).GuaranteeLt.cPos = rs("E019HSXLTSPI")
    Lt(0).GuaranteeLt.cObj = rs("E019HSXLTHWT")
    Lt(0).GuaranteeLt.cJudg = rs("E019HSXLTHWS")
    Cs(0).SpecCsMin = fncNullCheck(rs("E019HSXCNMIN"))
    Cs(0).SpecCsMax = fncNullCheck(rs("E019HSXCNMAX"))
    Cs(0).GuaranteeCs.cMeth = rs("E019HSXCNSPH")
    Cs(0).GuaranteeCs.cCount = rs("E019HSXCNSPT")
    Cs(0).GuaranteeCs.cPos = rs("E019HSXCNSPI")
    Cs(0).GuaranteeCs.cObj = rs("E019HSXCNHWT")
    Cs(0).GuaranteeCs.cJudg = rs("E019HSXCNHWS")
    GD(0).JudgFlagDen = rs("E020HSXDENKU")
    GD(0).SpecDenMax = fncNullCheck(rs("E020HSXDENMX"))
    GD(0).SpecDenMin = fncNullCheck(rs("E020HSXDENMN"))
    GD(0).GuaranteeDen.cObj = rs("E020HSXDENHT")
    GD(0).GuaranteeDen.cJudg = rs("E020HSXDENHS")
    GD(0).JudgFlagDvd2 = rs("E020HSXDVDKU")
    GD(0).SpecDvd2Max = fncNullCheck(rs("E020HSXDVDMXN"))
    GD(0).SpecDvd2Min = fncNullCheck(rs("E020HSXDVDMNN"))
    GD(0).GuaranteeDvd2.cObj = rs("E020HSXDVDHT")
    GD(0).GuaranteeDvd2.cJudg = rs("E020HSXDVDHS")
    GD(0).JudgFlagLdl = rs("E020HSXLDLKU")
    GD(0).SpecLdlMax = fncNullCheck(rs("E020HSXLDLMX"))
    GD(0).SpecLdlMin = fncNullCheck(rs("E020HSXLDLMN"))
    GD(0).GuaranteeLdl.cObj = rs("E020HSXLDLHT")
    GD(0).GuaranteeLdl.cJudg = rs("E020HSXLDLHS")
    rs.Close
    'Tail���ɂ��������B
    Lt(1) = Lt(0)
    Cs(1) = Cs(0)
    GD(1) = GD(0)

    scmzc_getSXLGuarantee = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    scmzc_getSXLGuarantee = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :GD���уf�[�^�̎擾�h���C�o
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :crynum         ,I   ,String            ,�����ԍ�
'          :ingotpos       ,I   ,Integer           ,�Ώ۔͈͂̊J�n�ʒu
'          :length         ,I   ,Integer           ,�Ώ۔͈͂̒���
'          :Gd             ,I   ,C_GD              ,����GD����\����
'      �@�@:�߂�l          , O  , FUNCTION_RETURN�@, �ǂݍ��݂̐���
'����      :GD�̏㉺���т��擾����
'����      :2002/03/14 ���� �M�� �쐬
'          :2002/03/22 �쑺 �C��
Public Function scmzc_getSXLGD(CRYNUM$, INGOTPOS%, LENGTH%, GD() As C_GD) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String
    Dim c1 As Integer

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcjudg_SQL.bas -- Function scmzc_getSXLGD"
    scmzc_getSXLGD = FUNCTION_RETURN_FAILURE

    'Top�����я�����
    GD(0).Den = -1
    GD(0).Dvd2 = -1
    GD(0).Ldl = -1
    'Tail�����я�����
    GD(1).Den = -1
    GD(1).Dvd2 = -1
    GD(1).Ldl = -1
    
'     If Left(CRYNUM, 1) <> "8" Then                    '2003/10/18 �폜 SystemBrain
   
'' ���㌋�����ю擾�@2003/09/16 Motegi ==================================> START
        '���㌋���̎��ю擾(�������葪��lTBL���)
'        For c1 = 0 To 1
'            sql = vbNullString
'            sql = sql & "select * from ("
'            sql = sql & "  select SXLGD_MSRSDEN, SXLGD_MSRSLDL, SXLGD_MSRSDVD2"
'            sql = sql & "  from TBCMJ014"
'            sql = sql & "  where CRYNUM='" & CRYNUM & "'"
'            If c1 = 0 Then
'                sql = sql & "    and POSITION<=" & INGOTPOS
'                sql = sql & "  order by POSITION desc, SMPKBN desc"
'            Else
'                sql = sql & "    and POSITION>=" & INGOTPOS + LENGTH
'                sql = sql & "  order by POSITION, SMPKBN"
'            End If
'            sql = sql & ") where rownum=1"
'            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'            If rs.RecordCount > 0 Then
'                GD(c1).Den = rs("SXLGD_MSRSDEN")
'                GD(c1).Dvd2 = rs("SXLGD_MSRSDVD2")
'                GD(c1).Ldl = rs("SXLGD_MSRSLDL")
'            End If
'            rs.Close
'        Next
'' -----------------------------------------
        '���㌋���̎��ю擾(�������葪��lTBL���)
        For c1 = 0 To 1
            sql = vbNullString
            sql = sql & "select * from ("
            sql = sql & "  select MSRSDEN, MSRSLDL, MSRSDVD2"
            sql = sql & "  from TBCMJ006"
            sql = sql & "  where CRYNUM='" & CRYNUM & "'"
            If c1 = 0 Then
                sql = sql & "    and POSITION<=" & INGOTPOS
                sql = sql & "  order by POSITION desc, SMPKBN desc"
            Else
                sql = sql & "    and POSITION>=" & INGOTPOS + LENGTH
                sql = sql & "  order by POSITION, SMPKBN"
            End If
            sql = sql & ") where rownum=1"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount > 0 Then
                GD(c1).Den = rs("MSRSDEN")
                GD(c1).Dvd2 = rs("MSRSDVD2")
                GD(c1).Ldl = rs("MSRSLDL")
            End If
            rs.Close
        Next
'' ���㌋�����ю擾�@2003/09/16 Motegi ==================================> END
'2003/10/18 �폜 SystemBrain -------------------------------------------��
'    Else
'        '�w���P�����̎��ю擾
'        sql = vbNullString
'        sql = sql & "select * from ("
'        sql = sql & " select GD1TOP as DEN_0, GD1TAIL as DEN_1, DIA2TOP as DVD_0, DIA2TAIL as DVD_1"
'        sql = sql & " from TBCMG002 "
'        sql = sql & " where CRYNUM = (select BLOCKID from TBCME040 where CRYNUM='" & CRYNUM & "' )"
'        sql = sql & " order by TRANCNT desc"
'        sql = sql & ") where rownum=1"
'        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        If rs.RecordCount > 0 Then
'            GD(0).Den = rs("DEN_0")
'            GD(0).Dvd2 = rs("DVD_0")
'            GD(0).Ldl = 0
'            GD(1).Den = rs("DEN_1")
'            GD(1).Dvd2 = rs("DVD_1")
'            GD(1).Ldl = 0
'        End If
'        rs.Close
'    End If
'2003/10/18 �폜 SystemBrain -------------------------------------------��
    
    '���т��擾�ł��Ȃ��ꍇ���L�蓾��B
    '�G���[�Ƃ����ɏ����l�̂܂ܐ���I������B
    scmzc_getSXLGD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    scmzc_getSXLGD = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :Cs���уf�[�^�̎擾�h���C�o
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :crynum         ,I   ,String            ,�����ԍ�
'          :ingotpos       ,I   ,Integer           ,�Ώ۔͈͂̊J�n�ʒu
'          :length         ,I   ,Integer           ,�Ώ۔͈͂̒���
'          :Cs             ,I   ,C_Cs              ,����Cs����\����
'      �@�@:�߂�l          , O  , FUNCTION_RETURN�@, �ǂݍ��݂̐���
'����      :Cs�̏㉺���т��擾����
'����      :2002/03/14 ���� �M�� �쐬
'          :2002/03/22 �쑺 �C��
Public Function scmzc_getSXLCs(CRYNUM$, INGOTPOS%, LENGTH%, Cs() As C_Cs) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String
    Dim c1 As Integer

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcjudg_SQL.bas -- Function scmzc_getSXLCs"
    scmzc_getSXLCs = FUNCTION_RETURN_FAILURE

    'Top�����я�����
    Cs(0).Cs = -1
    'Tail�����я�����
    Cs(1).Cs = -1
    
    
'    If Left(CRYNUM, 1) <> "8" Then                 '2003/10/18 �폜 SystemBrain
        '���㌋���̎��ю擾
        If Cs(0).SpecCsMin > 0 Then
            'FromTo�d�l�̏ꍇ�́A�u���b�N��Top/Bot����l����������(���p�s��)
            'Top��
            sql = vbNullString
            sql = sql & "select J.POSITION, J.CSMEAS "
            sql = sql & "from TBCME040 B, TBCMJ004 J "
            sql = sql & "where B.CRYNUM='" & CRYNUM & "'"
            sql = sql & "  and B.INGOTPOS<=" & INGOTPOS
            sql = sql & "  and " & INGOTPOS & "<B.INGOTPOS+B.LENGTH"
            sql = sql & "  and J.CRYNUM=B.CRYNUM and J.POSITION=B.INGOTPOS "
            sql = sql & "order by TRANCNT desc"
            sql = "select * from (" & sql & ") where rownum=1"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount > 0 Then
                'Cs(0).Cs = rs("CSMEAS")
                If IsNull(rs("CSMEAS")) = False Then Cs(0).Cs = rs("CSMEAS") Else Cs(0).Cs = -1  ' OINULL�Ή��@2005/03/08 TUKU
            End If
            'Bot��
            sql = vbNullString
            sql = sql & "select J.POSITION, J.CSMEAS "
            sql = sql & "from TBCME040 B, TBCMJ004 J "
            sql = sql & "where B.CRYNUM='" & CRYNUM & "'"
            sql = sql & "  and B.INGOTPOS<" & INGOTPOS + LENGTH
            sql = sql & "  and " & INGOTPOS + LENGTH & "<=B.INGOTPOS+B.LENGTH"
            sql = sql & "  and J.CRYNUM=B.CRYNUM and J.POSITION=B.INGOTPOS+B.LENGTH "
            sql = sql & "order by TRANCNT desc"
            sql = "select * from (" & sql & ") where rownum=1"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount > 0 Then
                'Cs(1).Cs = rs("CSMEAS")
                If IsNull(rs("CSMEAS")) = False Then Cs(1).Cs = rs("CSMEAS") Else Cs(1).Cs = -1  ' OINULL�Ή��@2005/03/08 TUKU
            End If
        Else
            'FromTo�d�l�łȂ���΁A�Ȃ�ׂ��߂��������猟������
            sql = vbNullString
'            sql = sql & "select * from ("
'            sql = sql & "  select CSMEAS"
'            sql = sql & "  from TBCMJ004 J"
'            sql = sql & "  where CRYNUM='" & CRYNUM & "'"
'            sql = sql & "    and POSITION>=" & INGOTPOS + LENGTH
'            sql = sql & "    and POSITION<=(select min(INGOTPOS) from TBCME043 where CRYNUM=J.CRYNUM and INGOTPOS>=" & INGOTPOS + LENGTH & " and CRYINDCS in ('1','2','3','4'))"
'            sql = sql & "  order by POSITION, TRANCOND, SMPKBN, TRANCNT desc"
'            sql = sql & ") where rownum=1"
            sql = sql & "select * from ("
            sql = sql & "  select CSMEAS"
            sql = sql & "  from TBCMJ004 J"
            sql = sql & "  where CRYNUM='" & CRYNUM & "'"
            sql = sql & "    and POSITION>=" & INGOTPOS + LENGTH
            sql = sql & "    and POSITION<=(select min(INPOSCS) from XSDCS where XTALCS=J.CRYNUM and INPOSCS>=" & INGOTPOS + LENGTH & " and CRYINDCSCS in ('1','2','3','4'))"
            sql = sql & "  order by POSITION, TRANCOND, SMPKBN, TRANCNT desc"
            sql = sql & ") where rownum=1"

            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount > 0 Then
                'Cs(1).Cs = rs("CSMEAS")
                If IsNull(rs("CSMEAS")) = False Then Cs(1).Cs = rs("CSMEAS") Else Cs(1).Cs = -1  ' OINULL�Ή��@2005/03/08 TUKU
            End If
            rs.Close
        End If
'2003/10/18 �폜 SystemBrain -------------------------------------------��
'    Else
'        '�w���P�����̎��ю擾
'        sql = vbNullString
'        sql = sql & "select * from ("
'        sql = sql & " select CSTOP, CSTAIL"
'        sql = sql & " from TBCMG002 "
'        sql = sql & " where CRYNUM = (select BLOCKID from TBCME040 where CRYNUM='" & CRYNUM & "' )"
'        sql = sql & " order by TRANCNT desc"
'        sql = sql & ") where rownum=1"
'        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        If rs.RecordCount > 0 Then
'            Cs(0).Cs = rs("CSTOP")
'            Cs(1).Cs = rs("CSTAIL")
'        End If
'        rs.Close
'    End If
'2003/10/18 �폜 SystemBrain -------------------------------------------��
    
    '���т��擾�ł��Ȃ��ꍇ���L�蓾��B
    '�G���[�Ƃ����ɏ����l�̂܂ܐ���I������B
    scmzc_getSXLCs = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    scmzc_getSXLCs = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :Lt���уf�[�^�̎擾�h���C�o
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :crynum         ,I   ,String            ,�����ԍ�
'          :ingotpos       ,I   ,Integer           ,�Ώ۔͈͂̊J�n�ʒu
'          :length         ,I   ,Integer           ,�Ώ۔͈͂̒���
'          :Lt             ,I   ,C_LT              ,�������C�t�^�C������\����
'      �@�@:�߂�l          , O  , FUNCTION_RETURN�@, �ǂݍ��݂̐���
'����      :Lt�̏㉺���т��擾����
'����      :2002/03/14 ���� �M�� �쐬
'          :2002/03/22 �쑺 �C��
Public Function scmzc_getSXLLt(CRYNUM$, INGOTPOS%, LENGTH%, Lt() As C_LT) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String
    Dim LTSPI As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcjudg_SQL.bas -- Function scmzc_getSXLLt"
    scmzc_getSXLLt = FUNCTION_RETURN_FAILURE

    'Top�����я�����
    Lt(0).Lt = -1
    'Tail�����я�����
    Lt(1).Lt = -1
    
'    If Left(CRYNUM, 1) <> "8" Then                 '2003/10/18 �폜 SystemBrain
        '���㌋���̎��ю擾

        '���ʂ̑��݂���LT���т���A�w��̑���ʒu�ɂ��������ʂ���������
        '�D�揇�ʂ� �@�U�֐�Ɠ���̑���ʒu�ōł��߂����� �A�U�֐��茵��������ʒu�ōł��߂�����
        LTSPI = Lt(1).GuaranteeLt.cPos
        If Trim$(LTSPI) = vbNullString Then LTSPI = "ZZ"    '�U�֐��HSXLTSPI=' '�̂Ƃ��́A�ǂ̑���ʒu�̌��ʂł�OK�Ƃ���
        sql = "select * from ("
        sql = sql & " select CRYNUM, POSITION, TRANCOND, SMPKBN, J.HINBAN, CALCMEAS, HSXLTSPI"
        sql = sql & " from TBCMJ007 J, TBCME019 E19"
        sql = sql & " where CRYNUM='" & CRYNUM & "'"
        sql = sql & "   and POSITION>=" & INGOTPOS + LENGTH
        sql = sql & "   and SMPLUMU='0'"
        sql = sql & "   and TRANCNT=(select max(TRANCNT) from TBCMJ007 where CRYNUM=J.CRYNUM and POSITION=J.POSITION and SMPKBN=J.SMPKBN and TRANCOND=J.TRANCOND)"
        sql = sql & "   and E19.HINBAN=J.HINBAN and E19.MNOREVNO=J.REVNUM and E19.FACTORY=J.FACTORY and E19.OPECOND=J.OPECOND"
        sql = sql & "   and HSXLTSPI<='" & LTSPI & "'"
        sql = sql & " order by decode(HSXLTSPI,'" & LTSPI & "',0,1), POSITION, TRANCOND, SMPKBN"
        sql = sql & ") where rownum=1"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount > 0 Then
            Lt(1).Lt = rs("CALCMEAS")
        End If
        rs.Close
'2003/10/18 �폜 SystemBrain -------------------------------------------��
'    Else
'        '�w���P�����̎��ю擾
'        sql = vbNullString
'        sql = sql & "select * from ("
'        sql = sql & " select LTFTOP, LTFTAIL"
'        sql = sql & " from TBCMG002 "
'        sql = sql & " where CRYNUM = (select BLOCKID from TBCME040 where CRYNUM='" & CRYNUM & "' )"
'        sql = sql & " order by TRANCNT desc"
'        sql = sql & ") where rownum=1"
'        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        If rs.RecordCount > 0 Then
'            Lt(0).Lt = rs("LTFTOP")
'            Lt(1).Lt = rs("LTFTAIL")
'        End If
'        rs.Close
'    End If
'2003/10/18 �폜 SystemBrain -------------------------------------------��

    '���т��擾�ł��Ȃ��ꍇ���L�蓾��B
    '�G���[�Ƃ����ɏ����l�̂܂ܐ���I������B
    
    scmzc_getSXLLt = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    scmzc_getSXLLt = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'�T�v      :�S�i�Ԃ̉��H�d�l�f�[�^�̎擾�h���C�o
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :HIN()          ,I   ,tFullHinban       ,�i�ԃ��X�g
'          :Spec()         ,O   ,Judg_Kakou        ,���H�d�l
'      �@�@:�߂�l          , O  , FUNCTION_RETURN�@, �ǂݍ��݂̐���
'����      :
'����      :2002/04/17 ���� �M�� �쐬
Public Function scmzc_getKakouSpec(HIN() As tFullHinban, Spec() As Judg_Kakou) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String
    Dim recCnt As Integer
    Dim c0 As Integer
    
    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcjudg_SQL.bas -- Function scmzc_getKakouSpec"
    scmzc_getKakouSpec = FUNCTION_RETURN_FAILURE
    
    '���߂��S�i�Ԃ̉��H�d�l�����߂�
    sql = "select HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXDPDIR, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDWCEN, HSXDWMIN, HSXDWMAX from TBCME018 "
    sql = sql & "Where " & SQLMake_HINBAN(HIN())

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    ReDim Spec(recCnt)
    For c0 = 1 To recCnt
        Spec(c0).TOP(0) = fncNullCheck(rs("HSXD1CEN"))
        Spec(c0).TOP(1) = fncNullCheck(rs("HSXD1MIN"))
        Spec(c0).TOP(2) = fncNullCheck(rs("HSXD1MAX"))
        Spec(c0).TAIL(0) = fncNullCheck(rs("HSXD2CEN"))
        Spec(c0).TAIL(1) = fncNullCheck(rs("HSXD2MIN"))
        Spec(c0).TAIL(2) = fncNullCheck(rs("HSXD2MAX"))
        Spec(c0).DPTH(0) = fncNullCheck(rs("HSXDDCEN"))
        Spec(c0).DPTH(1) = fncNullCheck(rs("HSXDDMIN"))
        Spec(c0).DPTH(2) = fncNullCheck(rs("HSXDDMAX"))
        Spec(c0).WIDH(0) = fncNullCheck(rs("HSXDWCEN"))
        Spec(c0).WIDH(1) = fncNullCheck(rs("HSXDWMIN"))
        Spec(c0).WIDH(2) = fncNullCheck(rs("HSXDWMAX"))
        Spec(c0).pos = rs("HSXDPDIR")
        rs.MoveNext
    Next
    rs.Close

    scmzc_getKakouSpec = FUNCTION_RETURN_SUCCESS
proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    scmzc_getKakouSpec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :���H���т̎擾�h���C�o
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :BLOCKID        ,I   ,String            ,�����ԍ�or�u���b�NID
'          :Jiltuseki      ,O   ,Judg_Kakou        ,���H����
'      �@�@:�߂�l          , O  , FUNCTION_RETURN�@, �ǂݍ��݂̐���
'����      :
'����      :2002/04/17 ���� �M�� �쐬
Public Function scmzc_getKakouJiltuseki(BLOCKID As String, Jiltuseki As Judg_Kakou) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String
    Dim recCnt As Integer
    Dim c0 As Integer
    Dim AGRFlag As Boolean
    Dim Ans As String
    Dim tINGOTPOS As Integer
    
    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcjudg_SQL.bas -- Function scmzc_getKakouJiltuseki"
    scmzc_getKakouJiltuseki = FUNCTION_RETURN_FAILURE
    
    '�Ώۃu���b�N�̉��H���т̏�����
    For c0 = 1 To 2
        Jiltuseki.TAIL(c0) = -1
        Jiltuseki.TOP(c0) = -1
        Jiltuseki.DPTH(c0) = -1
        Jiltuseki.WIDH(c0) = -1
    Next
    Jiltuseki.pos = ""
'2003/10/18 �폜 SystemBrain -------------------------------------------��
'    If Left(BLOCKID, 1) = "8" Then
'        '�w���P�����̏ꍇ
'        sql = "select DMTOP1, DMTOP2, DMTAIL1, DMTAIL2, NCHPOS, NCHDPTH1, NCHDPTH2, NCHWID1, NCHWID2 from TBCMG002 "
'        sql = sql & "where CRYNUM = '" & BLOCKID & "' and "
'        sql = sql & "TRANCNT = any(select max(TRANCNT) from TBCMG002 where CRYNUM = '" & BLOCKID & "')"
'
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        recCnt = rs.RecordCount
'        If recCnt = 0 Then
'            rs.Close
'            scmzc_getKakouJiltuseki = FUNCTION_RETURN_SUCCESS
'            GoTo proc_exit
'        End If
'        Jiltuseki.TAIL(1) = rs("DMTAIL1")
'        Jiltuseki.TAIL(2) = rs("DMTAIL2")
'        Jiltuseki.Top(1) = rs("DMTOP1")
'        Jiltuseki.Top(2) = rs("DMTOP2")
'        Jiltuseki.DPTH(1) = rs("NCHDPTH1")
'        Jiltuseki.DPTH(2) = rs("NCHDPTH2")
'        Jiltuseki.WIDH(1) = rs("NCHWID1")
'        Jiltuseki.WIDH(2) = rs("NCHWID2")
'        Jiltuseki.pos = rs("NCHPOS")
'        rs.Close
'    Else
'2003/10/18 �폜 SystemBrain -------------------------------------------��
        '�����グ�����̏ꍇ
        sql = "select DMTOP1, DMTOP2, DMTAIL1, DMTAIL2, NCHPOS, NCHDPTH, NCHWIDTH from TBCMI002 "
        sql = sql & "where CRYNUM='" & Left(BLOCKID, 9) & "000" & "'"
'        sql = sql & " and (select INGOTPOS from TBCME040 where BLOCKID='" & BLOCKID & "') between INGOTPOS and INGOTPOS+LENGTH-1 "
        '��ۯ��Ǘ�(TBCME040)�Q�ƒ�~�@05/10/03 ooba
        sql = sql & " and (select INPOSC2 from XSDC2 where CRYNUMC2 = '" & BLOCKID & "') between INGOTPOS and INGOTPOS+LENGTH-1 "
        sql = sql & "order by INGOTPOS desc, TRANCNT desc"
        sql = "select * from (" & sql & ") where rownum=1"
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        recCnt = rs.RecordCount
        If recCnt = 0 Then
            rs.Close
            scmzc_getKakouJiltuseki = FUNCTION_RETURN_SUCCESS
            GoTo proc_exit
        End If
        Jiltuseki.TAIL(1) = rs("DMTAIL1")
        Jiltuseki.TAIL(2) = rs("DMTAIL2")
        Jiltuseki.TOP(1) = rs("DMTOP1")
        Jiltuseki.TOP(2) = rs("DMTOP2")
        Jiltuseki.DPTH(1) = rs("NCHDPTH")
        Jiltuseki.DPTH(2) = -1
        Jiltuseki.WIDH(1) = rs("NCHWIDTH")
        Jiltuseki.WIDH(2) = -1
        Jiltuseki.pos = rs("NCHPOS")
        rs.Close
'    End If                         '2003/10/18 �폜 SystemBrain

    scmzc_getKakouJiltuseki = FUNCTION_RETURN_SUCCESS
proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    scmzc_getKakouJiltuseki = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'�T�v      :�����֐��F�����Ώەi�Ԃ�񋓂���SQL��Ԃ�
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:HIN()  �@�@�@,I  ,tFullHinban    �@,�i�Ԉꗗ
'      �@�@:�߂�l       ,O  ,String         �@,SQL
'����      :
'����      :2002/04/17  ���� �M�� �쐬
Public Function SQLMake_HINBAN(HIN() As tFullHinban) As String
    Dim c0 As Integer
    Dim temp As String

    temp = ""
    For c0 = 1 To UBound(HIN())
        If (Trim(HIN(c0).hinban) <> "Z") Or (Trim(HIN(c0).hinban) <> "G") Or (Trim(HIN(c0).hinban) <> "") Or (Trim(HIN(c0).hinban) <> vbNullString) Then
            temp = temp & "(HINBAN='" & HIN(c0).hinban & "'"
            temp = temp & " and MNOREVNO=" & HIN(c0).mnorevno
            temp = temp & " and FACTORY='" & HIN(c0).factory & "'"
            temp = temp & " and OPECOND='" & HIN(c0).opecond & "') or "
        End If
    Next
    SQLMake_HINBAN = "(" & Left(temp, Len(temp) - 4) & ")"
End Function

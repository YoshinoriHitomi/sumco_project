Attribute VB_Name = "s_cmzcSXLPreJudge"
'�i�Ԃ̐U�փ`�F�b�N�͋��ʊ֐����g�p����̂ō폜����-------start iida 2003/09/03
'Option Explicit
'
''WF�d�l�擾�\���̒�`
'Public Type typ_Wfsiyou
'    HWFTYPE As String * 1             ' �i�v�e�^�C�v
'    HWFCDIR As String * 1             ' �i�v�e�����ʕ�
'    HWFCDOP As String * 1             ' �i�v�e�����h�[�v
'    HWFDOP As String * 1              ' �i�v�e�h�[�p���g
'End Type
'
''�T�v      :�����ȍ~�ł̕i�ԐU�֎��ɁA�^�C�v���̔�����s��
''�p�����[�^    :�ϐ���        ,IO ,�^          ,����
''          :crynum        ,I  ,String      ,�����ԍ�
''          :ingotpos      ,I  ,Integer     ,�Ώ۔͈͂̊J�n�ʒu
''          :length        ,I  ,Integer     ,�Ώ۔͈͂̒���
''          :hin           ,I  ,tFullHinban ,�U�֐�̕i��
''          :judge_ok      ,O  ,Boolean     ,���茋��
''          :itemNG        ,O  ,String      ,����NG�ƂȂ�������
''          :�߂�l        ,O  ,FUNCTION_RETURN, ����̍���
''          :                   FUNCTION_RETURN_SUCCESS: �U�։�
''          :                   FUNCTION_RETURN_FAILURE: �U�֕s��
''          :                                �������͎d�l�G���[
''����      :�^�C�v�E���ʁE�h�[�p���g �ɂ��Ĕ��肷��
''����      :2002/03/26 �} �쐬
''          :itemNG�G���[���e
''               TYPE :�^�C�v�G���[
''               CDIR :���ʃG���[
''               CDOP :�����h�[�v�G���[
''               DOP  :�h�[�p���g�[�G���[
''               E021   :DB�G���[(E021,E022,E023�d�l�擾)
''               E042   :DB�G���[(E042)
'Public Function SXLPreJudge(CRYNUM$, IngotPos%, LENGTH%, HIN As tFullHinban, judge_ok As Boolean, itemNG$) As FUNCTION_RETURN
'Dim dbIsMine As Boolean
'Dim rs As OraDynaset
'Dim sql As String
'Dim mHIN As tFullHinban               '�U�֑O�i�ԗp�\����
'
'Dim Wsi  As typ_Wfsiyou            'WF�d�l�擾�\����
'Dim mWsi As typ_Wfsiyou            'WF�d�l�擾�\����(�U�֑O�i�ԗp�j
'
'    '�G���[�n���h���̐ݒ�
'    On Error GoTo PROC_ERR
'    gErr.Push "SXLPreJudge.bas -- Function SXLPreJudge"
'
'    If OraDB Is Nothing Then
'        dbIsMine = True
'        OraDBOpen
'    End If
'
'    SXLPreJudge = FUNCTION_RETURN_FAILURE
'
'    '�U�֌�i�Ԃ�Z�AG�i�Ԃ̏ꍇ�́A��������OK
'    If Trim(HIN.HINBAN) = "Z" Or Trim(HIN.HINBAN) = "G" Then
'        judge_ok = True
'        itemNG = "OK"
'        SXLPreJudge = FUNCTION_RETURN_SUCCESS
'        GoTo PROC_EXIT
'    End If
'
'    '�U�֑O�i�Ԏ擾�iSXL�Ǘ����j
'    sql = "select "
'    sql = sql & " E042.HINBAN, "
'    sql = sql & " E042.REVNUM, "
'    sql = sql & " E042.FACTORY, "
'    sql = sql & " E042.OPECOND "
'    sql = sql & " from "
'    sql = sql & " TBCME042 E042 "
'    sql = sql & " where  "
'    sql = sql & " E042.CRYNUM ='" & CRYNUM & "' and "
'    sql = sql & " E042.INGOTPOS >= "
'    sql = sql & "   (select MAX(INGOTPOS) "
'    sql = sql & "   from TBCME042 "
'    sql = sql & "   where  "
'    sql = sql & "   CRYNUM ='" & CRYNUM & "' and "
'    sql = sql & "   INGOTPOS <= " & IngotPos
'    sql = sql & "   GROUP BY CRYNUM  ) and "
'    sql = sql & " E042.INGOTPOS <" & IngotPos + LENGTH
'
'    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'
'    If rs.RecordCount <> 1 Then
'    '������Ȃ�������AFUNCTION_RETURN_FAILURE��Ԃ��B
'        judge_ok = False
'        itemNG = "E042"
'        rs.Close
'        SXLPreJudge = FUNCTION_RETURN_FAILURE
'        GoTo PROC_EXIT
'    Else
'        With mHIN
'            .HINBAN = rs("HINBAN")
'            .mnorevno = rs("REVNUM")
'            .factory = rs("FACTORY")
'            .opecond = rs("OPECOND")
'        End With
'        judge_ok = True
'    End If
'    rs.Close
'
'Debug.Print "1 �ύX��i��(�擾)'" & HIN.HINBAN & "' '" & HIN.mnorevno & "' '" & HIN.factory & "' '" & HIN.opecond & "'"
'Debug.Print "2 �U�֑O�i��'" & mHIN.HINBAN & "' '" & mHIN.mnorevno & "' '" & mHIN.factory & "' '" & mHIN.opecond & "'"
'
'
''�U�֑O�i�Ԏd�l�擾
'    sql = "select "
'    sql = sql & " E021.HWFTYPE HWFTYPE, "
'    sql = sql & " E022.HWFCDIR HWFCDIR, "
'    sql = sql & " E021.HWFDOP HWFDOP, "
'    sql = sql & " E023.HWFCDOP HWFCDOP "
'    sql = sql & " from "
'    sql = sql & " TBCME021 E021, TBCME022 E022, TBCME023 E023 "
'    sql = sql & " where "
'    sql = sql & " E021.HINBAN='" & mHIN.HINBAN & "' and "
'    sql = sql & " E021.MNOREVNO=" & mHIN.mnorevno & " and "
'    sql = sql & " E021.FACTORY='" & mHIN.factory & "' and "
'    sql = sql & " E021.OPECOND='" & mHIN.opecond & "' and "
'    sql = sql & " E022.HINBAN='" & mHIN.HINBAN & "' and "
'    sql = sql & " E022.MNOREVNO=" & mHIN.mnorevno & " and "
'    sql = sql & " E022.FACTORY='" & mHIN.factory & "' and "
'    sql = sql & " E022.OPECOND='" & mHIN.opecond & "' and "
'    sql = sql & " E023.HINBAN='" & mHIN.HINBAN & "' and "
'    sql = sql & " E023.MNOREVNO=" & mHIN.mnorevno & " and "
'    sql = sql & " E023.FACTORY='" & mHIN.factory & "' and "
'    sql = sql & " E023.OPECOND='" & mHIN.opecond & "'"
'
'    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'
'    If rs.RecordCount <> 1 Then
'        '������Ȃ�������AFUNCTION_RETURN_FAILURE��Ԃ��B
'        judge_ok = False
'        itemNG = "E042"
'        rs.Close
'        SXLPreJudge = FUNCTION_RETURN_FAILURE
'        GoTo PROC_EXIT
'    Else
'        ''����������A�ŐV�̕i�ԏ����Z�b�g���� FUNCTION_RETURN_SUCCESS��Ԃ�
'        With mWsi
'            .HWFCDIR = rs("HWFCDIR")
'            .HWFCDOP = rs("HWFCDOP")
'            .HWFTYPE = rs("HWFTYPE")
'            .HWFDOP = rs("HWFDOP")
'        End With
'        judge_ok = True
'        rs.Close
'    End If
'
''�U�֌�i�Ԏd�l�擾
'    sql = "select "
'    sql = sql & " E021.HWFTYPE HWFTYPE, "
'    sql = sql & " E022.HWFCDIR HWFCDIR, "
'    sql = sql & " E021.HWFDOP HWFDOP, "
'    sql = sql & " E023.HWFCDOP HWFCDOP "
'    sql = sql & " from "
'    sql = sql & " TBCME021 E021, TBCME022 E022, TBCME023 E023 "
'    sql = sql & " where "
'    sql = sql & " E021.HINBAN='" & HIN.HINBAN & "' and "
'    sql = sql & " E021.MNOREVNO=" & HIN.mnorevno & " and "
'    sql = sql & " E021.FACTORY='" & HIN.factory & "' and "
'    sql = sql & " E021.OPECOND='" & HIN.opecond & "' and "
'    sql = sql & " E022.HINBAN='" & HIN.HINBAN & "' and "
'    sql = sql & " E022.MNOREVNO=" & HIN.mnorevno & " and "
'    sql = sql & " E022.FACTORY='" & HIN.factory & "' and "
'    sql = sql & " E022.OPECOND='" & HIN.opecond & "' and "
'    sql = sql & " E023.HINBAN='" & HIN.HINBAN & "' and "
'    sql = sql & " E023.MNOREVNO=" & HIN.mnorevno & " and "
'    sql = sql & " E023.FACTORY='" & HIN.factory & "' and "
'    sql = sql & " E023.OPECOND='" & HIN.opecond & "'"
'
'    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'
'    If rs.RecordCount <> 1 Then
'        '������Ȃ�������AFUNCTION_RETURN_FAILURE��Ԃ��B
'        judge_ok = False
'        itemNG = "E021"
'        rs.Close
'        SXLPreJudge = FUNCTION_RETURN_FAILURE
'        GoTo PROC_EXIT
'    Else
'        ''����������A�ŐV�̕i�ԏ����Z�b�g���� FUNCTION_RETURN_SUCCESS��Ԃ�
'        With Wsi
'            .HWFCDIR = rs("HWFCDIR")
'            .HWFCDOP = rs("HWFCDOP")
'            .HWFTYPE = rs("HWFTYPE")
'            .HWFDOP = rs("HWFDOP")
'        End With
'        rs.Close
'        judge_ok = True
'    End If
'
'Debug.Print "3 �U�֑O�d�l'" & mWsi.HWFTYPE & "' '" & mWsi.HWFCDIR & "' '" & mWsi.HWFCDOP & "' '" & mWsi.HWFDOP & "'"
'Debug.Print "4 �U�֌�d�l'" & Wsi.HWFTYPE & "' '" & Wsi.HWFCDIR & "' '" & Wsi.HWFCDOP & "' '" & Wsi.HWFDOP & "'"
'
''�E�d�l��r�iOR�j
'    If mWsi.HWFTYPE <> Wsi.HWFTYPE Then
''       SXLPreJudge = FUNCTION_RETURN_FAILURE
'        judge_ok = False
'        itemNG = "�^�C�v"
'        rs.Close
''        GoTo proc_exit
'    ElseIf mWsi.HWFCDIR <> Wsi.HWFCDIR Then
''        SXLPreJudge = FUNCTION_RETURN_FAILURE
'        judge_ok = False
'        itemNG = "����"
''        GoTo proc_exit
'    ElseIf mWsi.HWFCDOP <> Wsi.HWFCDOP Then
''        SXLPreJudge = FUNCTION_RETURN_FAILURE
'    If Tokusai = "1" Then Else judge_ok = False
'        itemNG = "�����h�[�v"
''        GoTo proc_exit
'    ElseIf mWsi.HWFDOP <> Wsi.HWFDOP Then
''        SXLPreJudge = FUNCTION_RETURN_FAILURE
'        judge_ok = False
'        itemNG = "�h�[�p���g"
''        GoTo proc_exit
'    Else
'        itemNG = "OK"
'        judge_ok = True
''        SXLPreJudge = FUNCTION_RETURN_SUCCESS
'    End If
'
'Debug.Print "5 ���� '" & judge_ok & "' ���� '" & itemNG & "'"
'
'    SXLPreJudge = FUNCTION_RETURN_SUCCESS
'
'PROC_EXIT:
'    '�I��
'    gErr.Pop
'    Exit Function
'
'PROC_ERR:
'    '�G���[�n���h��
'    Debug.Print "====== Error SQL ======"
'    Debug.Print sql
'    gErr.HandleError
'    Resume PROC_EXIT
'
'End Function
'�i�Ԃ̐U�փ`�F�b�N�͋��ʊ֐����g�p����̂ō폜����-------end iida 2003/09/03


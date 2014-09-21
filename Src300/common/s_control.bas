Attribute VB_Name = "s_control"
Option Explicit

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


''�ėp�R�[�h�擾

'���[�U��`�^(def_�����d�l3 ���)
' �ėp����Ͻ�
Public Type typ_GPCodeMaster
    'HINBAN As String * 8        ' �i��
    'MNOREVNO As Integer         ' ���i�ԍ������ԍ�
    'FACTORY As String * 1       ' �H��
    'OPECOND As String * 1       ' ���Ə���
    codeNo As String * 12       ' �R�[�h�m�n
    CODE As String * 5          ' �R�[�h
    codeCont As String          ' �R�[�h���e
    INDORDER As Long            ' �\����
    codename As String          ' �R�[�h����
    KUBUN As String             ' �敪
    READTIME As Double          ' ���[�h�^�C��
    'IFKBN As String * 4         ' �h�^�e�敪
    'SYORIKBN As String * 1      ' �����敪
    'SPECRRNO As String * 9      ' �d�l�o�^�˗��ԍ�
    'SXLMCNO As String * 12      ' �r�w�k��������ԍ�
    'WFMCNO As String * 12       ' �v�e��������ԍ�
    'STAFFID As String * 8       ' �Ј�ID
    'REGDATE As Date             ' �o�^���t
    'UPDDATE As Date             ' �X�V���t
    'SENDFLAG As String * 1      ' ���M�t���O
    'SENDDATE As Date            ' ���M���t
End Type

Const ERR_INVALID_MSGID = "���b�Z�[�W�����o�^�ł�"




'�T�v      :�ėp�R�[�h�}�X�^����A����̃R�[�h�̓��e�����������
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :CODENO        ,I  ,String           ,�R�[�hNO
'          :CODE          ,I  ,String           ,�R�[�h
'          :�߂�l        ,O  ,String           ,���e������
'����      :������Ȃ��ꍇ��VbNullString��Ԃ�
'����      :2001/06/07 �쐬  �쑺
Public Function GetGPCodeCont(ByVal codeNo As String, ByVal CODE As String) As String
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''�ėp�R�[�h�}�X�^����A����̃R�[�h�̓��e�����������
    sql = "select CODECONT from TBCME033 where (rtrim(CODENO)='" & Trim$(codeNo) & "') and (rtrim(CODE)='" & Trim$(CODE) & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''������Ȃ�������AVbNullString��Ԃ�
        GetGPCodeCont = vbNullString
    Else
        ''����������A�R�[�h���e�������Ԃ�
        GetGPCodeCont = rs("CODECONT")
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If
End Function


'�T�v      :�ėp�R�[�h�}�X�^����A����̃R�[�h������
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :CODENO        ,I  ,String           ,�R�[�hNO
'          :CODE          ,I  ,String           ,�R�[�h
'          :GPCode        ,O  ,typ_GPCodeMaster ,�Ή�����f�[�^
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,����/���s
'����      :
'����      :2001/06/04 �쐬  �쑺
Public Function GetGPCode(ByVal codeNo As String, ByVal CODE As String, GPCode As typ_GPCodeMaster) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''�ėp�R�[�h�}�X�^����A����̃R�[�h������
    sql = "select CODECONT, CODENAME, INDORDER, KUBUN, READTIME from TBCME033 " & _
          "where (rtrim(CODENO)='" & Trim$(codeNo) & "') and (rtrim(CODE)='" & Trim$(CODE) & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''������Ȃ�������A�f�[�^���e����������FUNCTION_RETURN_FAILURE��Ԃ�
        With GPCode
            .codeNo = vbNullString
            .CODE = vbNullString
            .codeCont = vbNullString
            .codename = vbNullString
            .INDORDER = 0
            .KUBUN = vbNullString
            .READTIME = 0
        End With
        GetGPCode = FUNCTION_RETURN_FAILURE
    Else
        ''������Ȃ�������A�f�[�^���e��ݒ肵��FUNCTION_RETURN_SUCCESS��Ԃ�
        With GPCode
            .codeNo = codeNo
            .CODE = CODE
            .codeCont = rs("CODECONT")
            .codename = rs("CODENAME")
            .INDORDER = rs("INDORDER")
            .KUBUN = rs("KUBUN")
            .READTIME = rs("READTIME")
        End With
        GetGPCode = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If
End Function


'�T�v      :�ėp�R�[�h�}�X�^����A�R�[�hNO�ɑΉ�����R�[�h�̈ꗗ�𓾂�
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :CODENO        ,I  ,String           ,�R�[�hNO
'          :GPCodeList()  ,O  ,typ_GPCodeMaster ,�Ή�����R�[�h�f�[�^�̈ꗗ
'          :�߂�l        ,O  ,Integer          ,����/���s
'����      :
'����      :2001/06/04 �쐬  �쑺
Public Function GetGPCodeList(ByVal codeNo As String, GPCodeList() As typ_GPCodeMaster) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim i As Integer
Dim recCnt As Integer

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''�ėp�R�[�h�}�X�^����A�R�[�hNO�ɑΉ�����R�[�h�̈ꗗ�𓾂�
    sql = "select CODE, CODECONT, CODENAME, INDORDER, KUBUN, READTIME from TBCME033 where (rtrim(CODENO)='" & Trim$(codeNo) & "') order by INDORDER"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.EOF Then
        ''������Ȃ�������A0���Ƃ���FUNCTION_RETURN_FAILURE��Ԃ�
        ReDim GPCodeList(0)
        GetGPCodeList = FUNCTION_RETURN_FAILURE
    Else
        ''����������A���̌������̃f�[�^���R�s�[����FUNCTION_RETURN_SUCCESS��Ԃ�
        recCnt = rs.RecordCount
        ReDim GPCodeList(recCnt)
        For i = 1 To recCnt
            With GPCodeList(i)
                .codeNo = codeNo
                .CODE = rs("CODE")
                .codeCont = rs("CODECONT")
                .codename = rs("CODENAME")
                .INDORDER = rs("INDORDER")
                .KUBUN = rs("KUBUN")
                .READTIME = rs("READTIME")
                rs.MoveNext
            End With
        Next
        GetGPCodeList = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If
End Function


'�T�v      :�R���{���ɕ\������u�R�[�h:�R�[�h���e�v�̕�����𐶐�����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :code          ,I  ,String    ,�R�[�h
'          :codeCont      ,I  ,String    ,�R�[�h���e
'          :�߂�l        ,O  ,String    ,���ʕ�����
'����      :
'����      :2001/06/07 �쐬  �쑺
Public Function GetGPCodeDspStr(CODE$, codeCont$) As String

    If (Trim$(CODE) = "SPACE") Or (Trim$(CODE) = vbNullString) Then
        ''�R�[�h���uSPACE�v�̏ꍇ�A���ʕ�����=" "�Ƃ���
        GetGPCodeDspStr = " "
    Else
        ''����ȊO�̏ꍇ�A��������Ȃ����킹��
        GetGPCodeDspStr = Trim$(CODE) & ":" & Trim$(codeCont)
    End If
End Function


'�T�v      :�R���{�{�b�N�X�ɔėp�}�X�^���̑I������ݒ肷��
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :GPCodeList()  ,I  ,typ_GPCodeMaster ,�I�����̃��X�g
'          :cmb           ,I  ,ComboBox         ,�ݒ��̃R���{�{�b�N�X
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,����/���s
'����      :
'����      :2001/06/04 �쐬  �쑺
Public Function SetGPCodeList2Combo(GPCodeList() As typ_GPCodeMaster, cmb As ComboBox) As FUNCTION_RETURN
Dim RET As FUNCTION_RETURN
Dim max As Integer
Dim i As Integer

    With cmb
        .Clear
        max = UBound(GPCodeList)
        For i = 1 To max
            .AddItem GetGPCodeDspStr(GPCodeList(i).CODE, GPCodeList(i).codeCont)
        Next
    End With
    SetGPCodeList2Combo = FUNCTION_RETURN_SUCCESS
End Function


'�T�v      :�R���{�{�b�N�X�ɔėp�}�X�^���̑I������ݒ肷��
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :CODENO        ,I  ,String           ,�R�[�hNO
'          :cmb           ,I  ,ComboBox         ,�ݒ��̃R���{�{�b�N�X
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,����/���s
'����      :
'����      :2001/06/28 �쐬  �쑺
Public Function SetGPCode2Combo(ByVal codeNo As String, cmb As ComboBox) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim i As Integer
Dim recCnt As Integer

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''�ėp�R�[�h�}�X�^����A�R�[�hNO�ɑΉ�����R�[�h�̈ꗗ�𓾂�
    sql = "select CODE, CODECONT from TBCME033 where (rtrim(CODENO)='" & Trim$(codeNo) & "') order by INDORDER"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.EOF Then
        ''������Ȃ�������A0���Ƃ���FUNCTION_RETURN_FAILURE��Ԃ�
        SetGPCode2Combo = FUNCTION_RETURN_FAILURE
    Else
        ''����������A���̌������̃f�[�^���R�s�[����FUNCTION_RETURN_SUCCESS��Ԃ�
        recCnt = rs.RecordCount
        cmb.Clear
        For i = 1 To recCnt
            cmb.AddItem GetGPCodeDspStr(rs("CODE"), rs("CODECONT"))
            rs.MoveNext
        Next
        SetGPCode2Combo = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If
End Function


''�������ƃR�[�h�擾
''�Q��:s_cmzcTBCMB005_SQL.bas (�R�[�h�}�X�^)

'���[�U��`�^


'�T�v      :�������Ɨp�R�[�h���������A�w��t�B�[���h�̓��e�𓾂�
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :SYSCLASS      ,I  ,String    ,�V�X�e���敪
'          :CLASS         ,I  ,String    ,�敪
'          :CODE          ,I  ,String    ,�R�[�h
'          :FieldName     ,I  ,String    ,�t�B�[���h��
'          :�߂�l        ,O  ,String    ,�t�B�[���h���e
'����      :
'����      :2001/06/14 �쐬  �쑺
Public Function GetCodeField(ByVal SYSCLASS$, ByVal Class$, ByVal CODE$, ByVal FieldName$) As String
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function GetCodeField"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''�R�[�h�}�X�^����A�w��̃t�B�[���h������
    sql = "SELECT " & FieldName & " from TBCMB005 WHERE (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "') and (CODE='" & CODE & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''������Ȃ�������AVbNullString��Ԃ�
        GetCodeField = vbNullString
    Else
        ''����������A�w��t�B�[���h�̓��e��Ԃ�
        GetCodeField = Trim$(rs(FieldName))
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�������Ɨp�R�[�h���������A�w��t�B�[���h�̓��e�𓾂�
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :SYSCLASS      ,I  ,String    ,�V�X�e���敪
'          :CLASS         ,I  ,String    ,�敪
'          :FieldName     ,I  ,String    ,�t�B�[���h��
'          :FieldData()   ,O  ,String    ,�t�B�[���h���e
'          :�߂�l        ,O  ,String    ,�t�B�[���h���e
'����      :
'����      :2001/07/26 �쐬  �쑺
Public Function GetCodeField2(ByVal SYSCLASS$, ByVal Class$, ByVal FieldName$, FieldData() As String) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim i As Integer

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function GetCodeField2"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''�R�[�h�}�X�^����A�w��̃t�B�[���h������
    sql = "SELECT " & FieldName & " from TBCMB005 WHERE (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    ReDim FieldData(rs.RecordCount)
    For i = 1 To rs.RecordCount
        FieldData(i) = rs(FieldName)
        rs.MoveNext
    Next
    GetCodeField2 = FUNCTION_RETURN_SUCCESS
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    GetCodeField2 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�������Ɨp�R�[�h����������
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :SYSCLASS      ,I  ,String       ,�V�X�e���敪
'          :CLASS         ,I  ,String       ,�敪
'          :CODE          ,I  ,String       ,�R�[�h
'          :CodeData      ,O  ,typ_TBCMB005 ,��������
'          :�߂�l        ,O  ,FUNCTION_RETURN,�����̐���
'����      :
'����      :2001/06/07 �쐬  �쑺
Public Function GetCode(ByVal SYSCLASS$, ByVal Class$, ByVal CODE$, CodeData As typ_TBCMB005) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim sqlWhere As String
Dim rec() As typ_TBCMB005
Dim RET As FUNCTION_RETURN



    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function GetCode"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''DB�֐��𗘗p���āA�R�[�h���e���擾����
    sqlWhere = "where (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "') and (CODE='" & CODE & "')"
    RET = DBDRV_GetTBCMB005(rec, sqlWhere)
    If (UBound(rec) = 0) Then
        GetCode = FUNCTION_RETURN_FAILURE
        CodeData = rec(0)
    Else
        GetCode = FUNCTION_RETURN_SUCCESS
        CodeData = rec(1)
    End If
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�������Ɨp�R�[�h�̈ꗗ�𓾂�
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :SYSCLASS      ,I  ,String       ,�V�X�e���敪
'          :CLASS         ,I  ,String       ,�敪
'          :CodeList()    ,O  ,typ_TBCMB005 ,�R�[�h���e�̈ꗗ
'          :�߂�l        ,O  ,FUNCTION_RETURN,�����̐���
'����      :
'����      :
Public Function GetCodeList(ByVal SYSCLASS$, ByVal Class$, CodeList() As typ_TBCMB005) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim sqlWhere As String
Dim sqlOrder As String
Dim RET As FUNCTION_RETURN



    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function GetCodeList"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''DB�֐��𗘗p���āA�R�[�h���e���擾����
    sqlWhere = " where (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "')"
    sqlOrder = " Order by INFO9,CODE"
    RET = DBDRV_GetTBCMB005(CodeList, sqlWhere, sqlOrder)
    GetCodeList = RET
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function
'�T�v      :�������Ɨp�R�[�h�̈ꗗ�𓾂�   2006/01
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :SYSCLASS      ,I  ,String       ,�V�X�e���敪
'          :CLASS         ,I  ,String       ,�敪
'          :INFO          ,I  ,String       ,�\������
'          :CodeList()    ,O  ,typ_TBCMB005 ,�R�[�h���e�̈ꗗ
'          :�߂�l        ,O  ,FUNCTION_RETURN,�����̐���
'����      :
'����      :
Public Function GetCodeListSC18(ByVal SYSCLASS$, ByVal Class$, ByVal INFO$, CodeList() As typ_TBCMB005) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim sqlWhere As String
Dim sqlOrder As String
Dim RET As FUNCTION_RETURN



    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function GetCodeListSC18"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''DB�֐��𗘗p���āA�R�[�h���e���擾����
    If INFO = "CM" Then
        sqlWhere = " where (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "')and (trim(INFO2) ='" & INFO & "')"
    ElseIf INFO = "�O�`" Then
        '08/12/23 ooba
        sqlWhere = " where (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "')and (trim(INFO4) ='" & INFO & "')"
    Else
        sqlWhere = " where (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "')and (trim(INFO3) ='" & INFO & "')"
    End If
    sqlOrder = " Order by INFO9,CODE"
    RET = DBDRV_GetTBCMB005(CodeList, sqlWhere, sqlOrder)
    GetCodeListSC18 = RET
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�Ј�ID����Ј����������߂�(TBCMB001���擾)
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :staffID       ,I  ,String    ,�Ј�ID
'          :�߂�l        ,O  ,String    ,�Ј�����
'����      :������Ȃ������ꍇ�́AVbNullString��Ԃ�
'����      :2001/06/07 �쐬  �쑺
Public Function GetStaffName(StaffID$) As String
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function GetStaffName"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''�Ј��}�X�^����A�Ј���������
    sql = "SELECT JFMLNAME, JFSTNAME from TBCMB001 WHERE (STAFFID='" & StaffID & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''������Ȃ�������AVbNullString��Ԃ�
        GetStaffName = vbNullString
    Else
        ''����������A������Ԃ�
        GetStaffName = Trim$(rs("JFMLNAME")) & " " & Trim$(rs("JFSTNAME"))
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�Ј�ID����Ј����������߂�(�e�[�u��KODA9���擾)
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :staffID       ,I  ,String    ,�Ј�ID
'          :�߂�l        ,O  ,String    ,�Ј�����
'����      :������Ȃ������ꍇ�́AVbNullString��Ԃ�
'           2009/09/04 SUMCO Akizuki
'                      CMBC052���Q�l�ɍ쐬

Public Function GetStaffName_KODA9(StaffID$) As String
    Dim dbIsMine As Boolean
    Dim rs As OraDynaset
    Dim sql As String


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cntrol.bas -- Function newGetStaffName"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''�Ј��}�X�^����A�Ј���������
    sql = ""
    sql = "select NAMEJA9 from KODA9 "
    sql = sql & " where SYSCA9='K' and SHUCA9='55' and CODEA9='" & StaffID & "'"
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    
    If rs.RecordCount = 0 Then
        ''������Ȃ�������AVbNullString��Ԃ�
        GetStaffName_KODA9 = vbNullString
    Else
        ''����������A������Ԃ�
        GetStaffName_KODA9 = Trim$(rs("NAMEJA9"))
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�Ј�ID����Ј����������߂�(200mm�ڑ����Acmcc100�Acmec053�p)
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :staffID       ,I  ,String    ,�Ј�ID
'          :�߂�l        ,O  ,String    ,�Ј�����
'����      :������Ȃ������ꍇ�́AVbNullString��Ԃ�
'����      :2008/07/07�@SET ���� �쐬
Public Function GetStaffName200(StaffID$) As String
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function GetStaffName200"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''�Ј��}�X�^����A�Ј���������
    sql = "SELECT NAMEJA9 from KODA9 WHERE CODEA9='" & StaffID & _
    "' and SYSCA9='K' and SHUCA9='55'"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''������Ȃ�������AVbNullString��Ԃ�
        GetStaffName200 = vbNullString
    Else
        ''����������A������Ԃ�
        GetStaffName200 = Trim$(rs("NAMEJA9"))
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�R���{�{�b�N�X�Ɍ������ƃ}�X�^���̑I������ݒ肷��
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :SYSCLASS      ,I  ,String    ,SYS�敪
'          :Class         ,I  ,String    ,�敪
'          :FieldName     ,I  ,String    ,�I�������̃t�B�[���h��
'          :cmb           ,O  ,ComboBox  ,�ݒ��̃R���{�{�b�N�X
'          :�߂�l        ,O  ,FUNCTION_RETURN,
'����      :
'����      :2001/06/28 �쐬  �쑺
Public Function SetCode2Combo(ByVal SYSCLASS$, ByVal Class$, ByVal FieldName$, cmb As ComboBox) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim max As Integer
Dim i As Integer


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function SetCode2Combo"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''�R�[�h�}�X�^����A�w��̃t�B�[���h������
    sql = "SELECT CODE, " & FieldName & " from TBCMB005" & _
          " WHERE (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "')" & _
          " Order by INFO9"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''������Ȃ�������ANG
        cmb.Clear
        SetCode2Combo = FUNCTION_RETURN_FAILURE
    Else
        ''����������A�R���{�{�b�N�X�ɑI������ݒ肷��
        With cmb
            .Clear
            max = rs.RecordCount
            For i = 1 To max
                .AddItem GetGPCodeDspStr(rs("CODE"), rs(FieldName))
                rs.MoveNext
            Next
        End With
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function
'�T�v      :�R���{�{�b�N�X�Ɍ������ƃ}�X�^���̑I������ݒ肷��
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :SYSCLASS      ,I  ,String    ,SYS�敪
'          :Class         ,I  ,String    ,�敪
'          :FieldName     ,I  ,String    ,�I�������̃t�B�[���h��
'          :cmb           ,O  ,ComboBox  ,�ݒ��̃R���{�{�b�N�X
'          :�߂�l        ,O  ,FUNCTION_RETURN,
'����      :
'����      :2001/06/28 �쐬  �쑺
Public Function SetCode2ComboSC18(ByVal SYSCLASS$, ByVal Class$, ByVal INFO$, cmb As ComboBox) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim max As Integer
Dim i As Integer


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function SetCode2ComboSC18"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''�R�[�h�}�X�^����A�w��̃t�B�[���h������
    sql = "SELECT CODE, INFO1 from TBCMB005" & _
          " WHERE (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "')"
    If INFO = "CM" Then
        sql = sql & " AND   (INFO2 ='" & INFO & "')"
    Else
        sql = sql & " AND   (INFO3 ='" & INFO & "')"
    End If
    sql = sql & " Order by INFO9"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''������Ȃ�������ANG
        cmb.Clear
        SetCode2ComboSC18 = FUNCTION_RETURN_FAILURE
    Else
        ''����������A�R���{�{�b�N�X�ɑI������ݒ肷��
        With cmb
            .Clear
            max = rs.RecordCount
            For i = 1 To max
                .AddItem GetGPCodeDspStr(rs("CODE"), rs("INFO1"))
                rs.MoveNext
            Next
        End With
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�R���{�{�b�N�X�Ɍ������ƃ}�X�^���̑I������ݒ肷��i":�R�[�h���e"�Ȃ��o�[�W�����j
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :SYSCLASS      ,I  ,String    ,SYS�敪
'          :Class         ,I  ,String    ,�敪
'          :FieldName     ,I  ,String    ,�I�������̃t�B�[���h��
'          :cmb           ,O  ,ComboBox  ,�ݒ��̃R���{�{�b�N�X
'          :�߂�l        ,O  ,FUNCTION_RETURN,
'����      :
'����      :2001/08/21 �쐬  ���{
Public Function SetCode2Combo2(ByVal SYSCLASS$, ByVal Class$, ByVal FieldName$, cmb As ComboBox) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim max As Integer
Dim i As Integer
Dim CODE As String


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function SetCode2Combo2"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''�R�[�h�}�X�^����A�w��̃t�B�[���h������
    sql = "SELECT CODE, " & FieldName & " from TBCMB005" & _
          " WHERE (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "')" & _
          " Order by INFO9"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''������Ȃ�������ANG
        cmb.Clear
        SetCode2Combo2 = FUNCTION_RETURN_FAILURE
    Else
        ''����������A�R���{�{�b�N�X�ɑI������ݒ肷��
        With cmb
            .Clear
            max = rs.RecordCount
            For i = 1 To max
                CODE = rs("CODE")
                '.AddItem GetGPCodeDspStr(rs("CODE"), rs(FieldName))
                If (Trim$(CODE) = "SPACE") Or (Trim$(CODE) = vbNullString) Then
                     ''�R�[�h���uSPACE�v�̏ꍇ�A���ʕ�����=" "�Ƃ���
                    CODE = " "
                Else
                    CODE = Trim$(CODE)
                End If
                .AddItem CODE
                rs.MoveNext
            Next
        End With
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :SPREAD�̃R���{�{�b�N�X�ɐݒ肷�邽�߁A�������ƃ}�X�^���̑I������������擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :SYSCLASS      ,I  ,String    ,SYS�敪
'          :Class         ,I  ,String    ,�敪
'          :FieldName     ,I  ,String    ,�I�������̃t�B�[���h��
'          :�߂�l        ,O  ,String    ,�I����������
'����      :
'����      :2001/06/28 �쐬  �쑺
Public Function GetSSComboStr(ByVal SYSCLASS$, ByVal Class$, ByVal FieldName$) As String
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim max As Integer
Dim i As Integer
Dim cmbStr As String


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function GetSSComboStr"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''�R�[�h�}�X�^����A�w��̃t�B�[���h������
    sql = "SELECT CODE, " & FieldName & " from TBCMB005" & _
          " WHERE (SYSCLASS='" & SYSCLASS & "') and (CLASS='" & Class & "')" & _
          " Order by INFO9"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''������Ȃ�������ANG
        GetSSComboStr = vbNullString
    Else
        ''����������A�I�����������ݒ肷��
        max = rs.RecordCount
        For i = 1 To max
            If cmbStr <> vbNullString Then
                cmbStr = cmbStr & vbTab
            End If
            cmbStr = cmbStr & GetGPCodeDspStr(rs("CODE"), rs(FieldName))
            rs.MoveNext
        Next
        GetSSComboStr = cmbStr
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function
'�T�v      :SPREAD�̃R���{�{�b�N�X�ɐݒ肷�邽�߁A�������ƃ}�X�^���̑I������������擾����(KODA9��)
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :SYSCLASS      ,I  ,String    ,SYS�敪
'          :Class         ,I  ,String    ,�敪
'          :FieldName     ,I  ,String    ,�I�������̃t�B�[���h��
'          :�߂�l        ,O  ,String    ,�I����������
'����      :
'����      :2001/06/28 �쐬  �쑺
Public Function GetSSComboStrA9(ByVal SYSCLASS$, ByVal Class$, ByVal FieldName$) As String
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim max As Integer
Dim i As Integer
Dim cmbStr As String


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc004b.bas -- Function GetSSComboStrA9"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''�R�[�h�}�X�^����A�w��̃t�B�[���h������
    sql = "SELECT CODEA9, " & FieldName & " from KODA9" & _
          " WHERE (SYSCA9='" & SYSCLASS & "') and (SHUCA9='" & Class & "')" & _
          " Order by CTR01A9"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''������Ȃ�������ANG
        GetSSComboStrA9 = vbNullString
    Else
        ''����������A�I�����������ݒ肷��
        max = rs.RecordCount
        For i = 1 To max
            If cmbStr <> vbNullString Then
                cmbStr = cmbStr & vbTab
            End If
            cmbStr = cmbStr & GetGPCodeDspStr(rs("CODEA9"), rs(FieldName))
            rs.MoveNext
        Next
        GetSSComboStrA9 = cmbStr
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function



'�T�v      :���b�Z�[�W��������擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :MsgID         ,I  ,String    ,���b�Z�[�WID
'          :params()      ,I  ,Variant   ,���ߍ��݃p�����[�^(�K�v�Ȑ�����)
'          :�߂�l        ,O  ,String    ,���b�Z�[�W������
'����      :�擾�ł��Ȃ������ꍇ�́A�Œ胁�b�Z�[�W(���b�Z�[�W��DB�ɂ���܂���)��Ԃ�
'����      :2001/06/07 �쐬  �쑺
Public Function GetMsgStr(ByVal MsgID$, ParamArray params() As Variant) As String
Dim dbIsMine As Boolean
Dim rs As OraDynaset    '���R�[�h�Z�b�g
Dim sql As String       'SQL������
Dim fmt As String       '����������


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc004c.bas -- Function GetMsgStr"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''���b�Z�[�W�}�X�^�[���珑����������擾����
    sql = "select FORMINFO from TBCMB003 where (MSGID='" & MsgID & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''�Y�����郁�b�Z�[�WID���Ȃ��Ƃ��́A�Œ胁�b�Z�[�W��Ԃ�
        GetMsgStr = ERR_INVALID_MSGID & "(" & MsgID & ")"
    Else
        ''�ʏ�́A�����Ƀp�����[�^�𖄂ߍ���ŕԂ�
        fmt = rs("FORMINFO")
        GetMsgStr = FmtStr(fmt, params)
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If


proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :���b�Z�[�W��������擾����(200mm�ڑ����Acmcc100�Acmec053�p)
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :MsgID         ,I  ,String    ,���b�Z�[�WID
'          :params()      ,I  ,Variant   ,���ߍ��݃p�����[�^(�K�v�Ȑ�����)
'          :�߂�l        ,O  ,String    ,���b�Z�[�W������
'����      :�擾�ł��Ȃ������ꍇ�́A�Œ胁�b�Z�[�W(���b�Z�[�W��DB�ɂ���܂���)��Ԃ�
'          :200mmDB�ڑ�����DBLINK�o�R��300mmDB�̃e�[�u�����Q�Ƃ���B
'����      :2008/07/09 �쐬  �쑺
Public Function GetMsgStr200(ByVal MsgID$, ParamArray params() As Variant) As String
Dim dbIsMine As Boolean
Dim rs As OraDynaset    '���R�[�h�Z�b�g
Dim sql As String       'SQL������
Dim fmt As String       '����������


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc004c.bas -- Function GetMsgStr"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''���b�Z�[�W�}�X�^�[���珑����������擾����
    sql = "select FORMINFO from TBCMB003@DBLINK300 where (MSGID='" & MsgID & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''�Y�����郁�b�Z�[�WID���Ȃ��Ƃ��́A�Œ胁�b�Z�[�W��Ԃ�
        GetMsgStr200 = ERR_INVALID_MSGID & "(" & MsgID & ")"
    Else
        ''�ʏ�́A�����Ƀp�����[�^�𖄂ߍ���ŕԂ�
        fmt = rs("FORMINFO")
        GetMsgStr200 = FmtStr(fmt, params)
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If


proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�w��̏����Ƀp�����[�^�𖄂ߍ��񂾕������Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :fmt           ,I  ,String    ,���ߍ��ݐ敶����(printf������)
'          :params()      ,I  ,Variant   ,���ߍ��݃p�����[�^(�ό�)
'          :�߂�l        ,O  ,String    ,���ߍ��݌��ʕ�����
'����      :
'����      :2001/06/06 �쐬  ����
Private Function FmtStr(ByVal fmt$, ParamArray params() As Variant) As String
Dim w_str       As String       '���ߍ��ݕ�����
Dim w_wrd       As String       '���Ұ�������
Dim i           As Integer      'ٰ�߶���
Dim n           As Integer      '���Ұ�����
Dim Str_Value() As String       '̫�ϯĂ�%����ɋ�؂����z��
Dim s_max       As Integer      '�Y���̍ő�l(̫�ϯĔz��)


    '������̫�ϯĕ������%����ɋ�؂��Ĕz��Ɋi�[

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc004c.bas -- Function FmtStr"

    Str_Value() = Split(fmt, "%")
    s_max = UBound(Str_Value)           '�Y���̍ő�l�擾
    
    n = 0                               '���Ұ��z��̓Y�����ď�����
    
    '������쐬
    For i = 0 To s_max                  '�%��ŋ�؂���̫�ϯĕ����񐔕�ٰ��
        '%�̎��̕����ɂ���ď����𕪊�
        Select Case Left(Str_Value(i), 1)
        Case "s"                        '������̏ꍇ
            If n > UBound(params(0)) Then
                w_wrd = vbNullString
            Else
                w_wrd = params(0)(n)        '���Ұ�������̎擾
            End If
            w_str = w_str & w_wrd & Mid(Str_Value(i), 2)
            n = n + 1                   '�������Ұ���
        Case ""                         '�%������̏ꍇ
            'If i = s_max Then           '���߂��Ō�̏ꍇ
                w_str = w_str & Str_Value(i) 'Mid(Str_Value(i), 2)
            'Else                        '���߂���ɑ����ꍇ
            '    w_str = w_str & "%"     '�%�����
            '    i = i + 1               '���̐߂͔�΂�
            'End If
        Case Else
            w_str = w_str & Str_Value(i)
        End Select
    Next
       
    FmtStr = w_str

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :8���i�ԂɑΉ�����ŐV�̕i�ԏ�����������
'���Ұ�    :�ϐ���        ,IO ,�^          ,����
'          :hinban        ,I  ,String      ,8���i��
'          :fullHinban    ,O  ,tFullHinban ,�i�ԏ��
'          :[chkUsable]   ,I  ,Boolean     ,�g�p�J�nTbl �� ���������Ǘ�Tbl ��K�{�Ƃ��邩
'          :�߂�l        ,O  ,FUNCTION_RETURN,�����̐���
'����      :�����8���i�Ԃ̒��ł́A�����ԍ����傫�����̂����V����
'          :�X�ɂ��̒��ł͑��Ə����ԍ����傫�����̂����V����
'����      :2001/06/07 �쐬  �쑺
Public Function GetLastHinban(hinban$, fullHinban As tFullHinban, Optional chkUsable As Boolean = True) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc010a.bas -- Function GetLastHinban"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''���i�d�lSXL�f�[�^1 ����A�w��i�Ԃ̃��R�[�h��V�������Ɏ��o��
    ''�ł��u�V�����v�f�[�^�́A�����ԍ����ő�̂��̂̒��ő��Ə����ԍ����ő�ł��郌�R�[�h�ł���
    ''(�������A�g�p�J�nTbl�ƌ��������Ǘ�Tbl�ɓo�^����Ă��Ȃ��i�Ԃ́A�܂����p�s�\�ł���)8/22�C��
    ''2006/09 TBCME036�J�n�t���O��'1'(�J�n�d������M)�̂��� <---- 2006/12/11 �C��
    If chkUsable Then
        'sql = "select E018.HINBAN, E018.MNOREVNO, E018.FACTORY, E018.OPECOND " & _
              "from TBCME018 E018, TBCME032 E032, TBCME036 E036 " & _
              "Where (E018.HINBAN = '" & hinban & "')" & _
              " and (E018.HINBAN=E032.HINBAN) and (E018.MNOREVNO=E032.MNOREVNO)" & _
              " and (E018.FACTORY=E032.FACTORY) " & _
              " and (E018.HINBAN=E036.HINBAN) and (E018.MNOREVNO=E036.MNOREVNO)" & _
              " and (E018.FACTORY=E036.FACTORY) and (E018.OPECOND=E036.OPECOND) " & _
              "order by MNOREVNO DESC, OPECOND DESC"
        'sql = "select E018.HINBAN, E018.MNOREVNO, E018.FACTORY, E018.OPECOND " & _
              "from TBCME018 E018 " & _
              "Where (E018.HINBAN = '" & hinban & "')" & _
              " and (E018.SYNFLAG IS NULL OR E018.SYNFLAG='1') " & _
              " and (E018.OPECOND <> '1') " & _
              "order by MNOREVNO DESC, OPECOND DESC"
        sql = "select E018.HINBAN, E018.MNOREVNO, E018.FACTORY, E018.OPECOND " & _
              "from TBCME018 E018 , TBCME036 E036 " & _
              "Where (E018.HINBAN = '" & hinban & "')" & _
              " and (E018.SYNFLAG IS NULL OR E018.SYNFLAG='1') " & _
              " and (E018.OPECOND <> '1') " & _
              " and (E018.HINBAN=E036.HINBAN) and (E018.MNOREVNO=E036.MNOREVNO)" & _
              " and (E018.FACTORY=E036.FACTORY) and (E018.OPECOND=E036.OPECOND) " & _
              " and (E036.KAISIFLG = '1') " & _
              "order by MNOREVNO DESC, OPECOND DESC"
    Else
        '������͎d�l����E����������͗p
        '��������t�^����ɓo�^����Ă���i�Ԃ͖����Ƃ���
        sql = "select A.HINBAN, A.MNOREVNO, A.FACTORY, A.OPECOND " & _
              "from TBCME018 A, TBCME031 B " & _
              "where (A.HINBAN = '" & hinban & "')" & _
              " and (A.HINBAN=B.HINBAN(+)) and (A.MNOREVNO=B.MNOREVNO(+)) and (A.FACTORY=B.FACTORY(+))" & _
              " and (B.HINBAN is null) " & _
              "order by MNOREVNO DESC, OPECOND DESC"
    End If
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
    'If rs.RecordCount = 0 Or rs("OPECOND") = "1" Then
        ''������Ȃ�������AFUNCTION_RETURN_FAILURE��Ԃ�
        With fullHinban
            .hinban = vbNullString
            .mnorevno = 0
            .factory = vbNullString
            .opecond = vbNullString
        End With
        GetLastHinban = FUNCTION_RETURN_FAILURE
    Else
        ''����������A�ŐV�̕i�ԏ����Z�b�g���� FUNCTION_RETURN_SUCCESS��Ԃ�
        With fullHinban
            .hinban = rs("HINBAN")
            .mnorevno = rs("MNOREVNO")
            .factory = rs("FACTORY")
            .opecond = rs("OPECOND")
        End With
        GetLastHinban = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :8���i�ԂɑΉ�����ŐV�̎d�l�i�ԏ�����������
'���Ұ�    :�ϐ���        ,IO ,�^          ,����
'          :hinban        ,I  ,String      ,8���i��
'          :fullHinban    ,O  ,tFullHinban ,�i�ԏ��
'          :�߂�l        ,O  ,FUNCTION_RETURN,�����̐���
'����      :�d�l�f�[�^�̑��Ə����͏�Ɂu1�v�ł���
'          :�����8���i�Ԃ̒��ł́A�����ԍ����傫�����̂����V����
'����      :2001/06/07 �쐬  �쑺
Public Function GetLastSpecHinban(hinban$, fullHinban As tFullHinban) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc010a.bas -- Function GetLastSpecHinban"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''���i�d�lSXL�f�[�^1 ����A�w��i�Ԃ̎d�l���R�[�h��V�������Ɏ��o��
    ''�ł��u�V�����v�f�[�^�́A�����ԍ����ő�̂��̂̒��ł��郌�R�[�h�ł���
    sql = "SELECT HINBAN, MNOREVNO, FACTORY, OPECOND " & _
          "From TBCME018 " & _
          "Where (HINBAN = '" & hinban & "') AND (OPECOND = '1')" & _
          "ORDER BY MNOREVNO DESC;"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        ''������Ȃ�������AFUNCTION_RETURN_FAILURE��Ԃ�
        With fullHinban
            .hinban = vbNullString
            .mnorevno = 0
            .factory = vbNullString
            .opecond = vbNullString
        End With
        GetLastSpecHinban = FUNCTION_RETURN_FAILURE
    Else
        ''����������A�ŐV�̕i�ԏ����Z�b�g���� FUNCTION_RETURN_SUCCESS��Ԃ�
        With fullHinban
            .hinban = rs("HINBAN")
            .mnorevno = rs("MNOREVNO")
            .factory = rs("FACTORY")
            .opecond = rs("OPECOND")
        End With
        GetLastSpecHinban = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�w��̌����ԍ��Ɋ܂܂��i�Ԃ̈ꗗ�𓾂�
'���Ұ�    :�ϐ���        ,IO ,�^          ,����
'          :cryno         ,I  ,String      ,�����ԍ�
'          :hinban()      ,O  ,tFullHinban ,�i�ԃ��X�g
'          :�߂�l        ,O  ,FUNCTION_RETURN,���o�̐���
'����      :
'����      :2001/06/27 �쐬  ����
Public Function GetXlHinban(cryno$, hinban() As tFullHinban) As FUNCTION_RETURN
Dim rs      As OraDynaset               '���oRecordDynaset
Dim rsCnt   As Integer                  'ں��޶���
Dim sql     As String                   'SQL��
Dim i       As Integer                  'ٰ�߶���

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc010a.bas -- Function GetXlHinban"

    'SQL���̍쐬
    sql = "Select CRYNUM, HINBAN, REVNUM, FACTORY, OPECOND from TBCME041 "
    sql = sql & "Where(CRYNUM = '" & cryno & "')"
    
    '�f�[�^�̒��o
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    '''���o���R�[�h�����݂��Ȃ��ꍇ
    If rs.EOF Then
        ReDim hinban(0)                     '�z��̏�����
        GetXlHinban = FUNCTION_RETURN_FAILURE   '�װ�ð��
        GoTo proc_exit
    End If
        
    rsCnt = rs.RecordCount                  'ں��ސ��̶��Ă����
    ReDim hinban(rsCnt - 1)                 '�z��̍Ē�`
    
    '�z��ɒl���Z�b�g
    rs.MoveFirst                            '�擪ں��ނɈړ�
    For i = 0 To rsCnt - 1                  'ں��ސ���ٰ��
        DoEvents
        With hinban(i)
            .hinban = rs!hinban             '�i��
            .mnorevno = rs!REVNUM           '���i�ԍ������ԍ�
            .factory = rs!factory           '�H��
            .opecond = rs!opecond           '���Ə���
        End With
        rs.MoveNext                         '��ں��ނɈړ�
    Next
    
    GetXlHinban = FUNCTION_RETURN_SUCCESS   '����ð��
 

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function



'�T�v      :�h�[�p���g�Z�x�}�X�^����h�[�p���g���̈ꗗ�𓾂�
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :DopeName()    ,O  ,String    ,�h�[�p���g��
'          :�߂�l        ,O  ,FUNCTION_RETURN,���o�̐���
'����      :
'����      :2001/08/08 �쐬  �쑺
'          :2011/05/09 �擾�c�a�ύX Kameda
Public Function GetDopeNames(DopeName() As String) As FUNCTION_RETURN
Dim rs      As OraDynaset               '���oRecordDynaset
Dim rsCnt   As Integer                  'ں��޶���
Dim sql     As String                   'SQL��
Dim i       As Integer                  'ٰ�߶���

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc010a.bas -- Function GetDopeNames"

    GetDopeNames = FUNCTION_RETURN_FAILURE
    
    'SQL���̍쐬
    'sql = "select DOPKIND from TBCMB009 order by DOPKIND"    2011/05/09 Kameda
    'SQL�ҏW
    sql = "SELECT  NVL(codea9   , ' ') DOPKIND "
    sql = sql & "  FROM koda9 "
    sql = sql & " WHERE sysca9 = 'X'"
    sql = sql & "   AND shuca9 = 'D0'"
    sql = sql & " ORDER BY codea9"
    '�f�[�^�̒��o
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    '''���o���R�[�h�����݂��Ȃ��ꍇ
    If rs.EOF Then
        ReDim DopeName(0)                     '�z��̏�����
        GoTo proc_exit
    End If
        
    rsCnt = rs.RecordCount                  'ں��ސ��̶��Ă����
    ReDim DopeName(1 To rsCnt)                 '�z��̍Ē�`
    For i = 1 To rsCnt
        DopeName(i) = rs("DOPKIND")
        
        If Len(DopeName(i)) < 7 Then         '7���ɂ��킹��
            DopeName(i) = DopeName(i) & Space(7 - Len(DopeName(i)))
        Else
            DopeName(i) = Left(DopeName(i), 7)
        End If
        
        rs.MoveNext
    Next
    rs.Close

    GetDopeNames = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�w��̈�̕i�Ԃ�����������i�i�ԊǗ�Tbl�Ώہj
'���Ұ�    :�ϐ���        ,IO ,�^          ,����
'          :CRYNUM        ,I  ,String      ,�����ԍ�
'          :ChgFrom       ,I  ,Integer     ,�̈�J�n�ʒu
'          :ChgLength     ,I  ,Integer     ,�̈�I���ʒu
'          :hin           ,I  ,tFullHinban ,����������i��
'          :�߂�l        ,O  ,FUNCTION_RETURN,
'����      :
'����      :2001/08/11 �쐬  �쑺
Public Function ChangeAreaHinban(CRYNUM$, ChgFrom%, ChgLength%, HIN As tFullHinban) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim ChgTo As Integer

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc010a.bas -- Function ChangeAreaHinban"

    ChangeAreaHinban = FUNCTION_RETURN_FAILURE
    ChgTo = ChgFrom + ChgLength
    
    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
Debug.Print "=== ChangeAreaHinban ==="
    ''�w��̈��S�Ċ܂ޕi�Ԃ�����Ύw��̈�̉����ƂȂ郌�R�[�h�𕪊��쐬����
    sql = "insert into TBCME041 (CRYNUM,INGOTPOS,HINBAN,REVNUM,FACTORY,OPECOND,LENGTH,REGDATE,UPDDATE,SENDFLAG,SENDDATE) " & _
          "select CRYNUM, " & ChgTo & ", HINBAN, REVNUM, FACTORY, OPECOND, INGOTPOS+LENGTH-" & ChgTo & ", REGDATE, UPDDATE, SENDFLAG, SENDDATE " & _
          "From TBCME041 " & _
          "where (CRYNUM='" & CRYNUM & "') and (INGOTPOS<" & ChgFrom & ") and (INGOTPOS+LENGTH>" & ChgTo & ")"
Debug.Print sql
    If OraDB.ExecuteSQL(sql) < 1 Then
        Debug.Print "�S����܂ޕi�Ԃ͂Ȃ�����"
    Else
    ''     WriteDBLog sql
    End If
    
    ''�w��̈�̊J�n�ʒu���܂ޕi�Ԃ�����΂���𒲐�����
    sql = "update TBCME041 set LENGTH=" & ChgFrom & "-INGOTPOS, UPDDATE=SYSDATE " & _
          "where (CRYNUM='" & CRYNUM & "') and (INGOTPOS<" & ChgFrom & ") and (INGOTPOS+LENGTH>" & ChgFrom & ")"
Debug.Print sql
    If OraDB.ExecuteSQL(sql) < 1 Then
        Debug.Print "�J�n�ʒu���܂ޕi�Ԃ͂Ȃ�����"
    Else
    ''    WriteDBLog sql
    End If
    
    ''�w��̈�̏I���ʒu���܂ޕi�Ԃ�����΂���𒲐�����
    sql = "update TBCME041 set INGOTPOS=" & ChgTo & ", LENGTH=INGOTPOS+LENGTH-" & ChgTo & ", UPDDATE=SYSDATE " & _
          "where (CRYNUM='" & CRYNUM & "') and (INGOTPOS<" & ChgTo & ") and (INGOTPOS+LENGTH>" & ChgTo & ")"
Debug.Print sql
    If OraDB.ExecuteSQL(sql) < 1 Then
        Debug.Print "�I���ʒu���܂ޕi�Ԃ͂Ȃ�����"
    Else
    ''    WriteDBLog sql
    End If
    
    ''�w��̈���ɑS�悪�܂܂��i�Ԃ��폜����(��v���郌�R�[�h���܂�)
    sql = "delete from TBCME041 where (CRYNUM='" & CRYNUM & "') and (INGOTPOS>=" & ChgFrom & ") and (INGOTPOS+LENGTH<=" & ChgTo & ")"
Debug.Print sql
    If OraDB.ExecuteSQL(sql) < 1 Then
        Debug.Print "�S�悪�܂܂��i�Ԃ͂Ȃ�����"
    Else
    ''    WriteDBLog sql
    End If
    
    ''�w��̈�̕i�Ԃ�ǉ�����
    With HIN
    ''    WriteDBLog sql
        sql = "insert into TBCME041 (CRYNUM,INGOTPOS,HINBAN,REVNUM,FACTORY,OPECOND,LENGTH,REGDATE,UPDDATE,SENDFLAG,SENDDATE) values " & _
              "('" & CRYNUM & "', " & ChgFrom & ", '" & .hinban & "', " & .mnorevno & ", '" & .factory & "', '" & .opecond & "'," & _
              ChgLength & ", SYSDATE, SYSDATE, '0', SYSDATE)"
    End With
Debug.Print sql
    If OraDB.ExecuteSQL(sql) < 1 Then
        Debug.Print "�i�Ԓǉ����s"
        GoTo proc_exit
    Else
    ''    WriteDBLog sql
    End If
    
    If dbIsMine Then
        OraDBClose
    End If
    
    ChangeAreaHinban = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�u���b�N���z�[���h��Ԃ��ǂ������ׂ�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :DopeName()    ,I  ,String    ,�u���b�NID
'          :�߂�l        ,O  ,Integer   ,0:�z�[���h��ԂłȂ� 1:�z�[���h��� -1:�ǂݍ��݃G���[
'����      :
'����      :2001/09/18 �쐬  ���{
Public Function CheckHoldBlock(BLOCKID As String) As Integer

    Dim rs      As OraDynaset               '���oRecordDynaset
    Dim sql     As String                   'SQL��

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc010a.bas -- Function CheckHoldBlock"

    CheckHoldBlock = 0
    
    'SQL���̍쐬
    sql = "select HOLDCLS from TBCME040 where BLOCKID='" & BLOCKID & "' "
     
    '�f�[�^�̒��o
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '''���o���R�[�h�����݂��Ȃ��ꍇ
    If rs.EOF Then
        CheckHoldBlock = -1
        GoTo proc_exit
    End If
            
    If rs("HOLDCLS") = 1 Then
        CheckHoldBlock = 1
    End If

    rs.Close

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    CheckHoldBlock = -1
    Resume proc_exit
End Function

'�T�v      :�����̌^�𒲂ׂ�
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :Crynum        ,I  ,String    ,�����ԍ�
'          :�߂�l        ,O  ,String    ,"P+","P-","N+","N-","Unknown" �̂����ꂩ
'����      :
'����      :2002/03/28 �쐬  �쑺
Public Function GetXlType(CRYNUM$) As String
Dim sql As String
Dim rs As OraDynaset               '���oRecordDynaset

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc010a.bas -- Function GetXlType"
    
    GetXlType = "Unknown"

    sql = "select"
    sql = sql & " case when E018.HSXTYPE='P' then"
    sql = sql & "   case when E018.HSXRMAX <="
    sql = sql & "             (select to_number(INFO1) from TBCMB005"
    sql = sql & "              where SYSCLASS='LG' and CLASS='02' and CODE='P+')"
    sql = sql & "        then 'P+' else 'P-' end"
    sql = sql & "   when E018.HSXTYPE='N' then"
    sql = sql & " case when E018.HSXRMAX <="
    sql = sql & "             (select to_number(INFO1) from TBCMB005"
    sql = sql & "              where SYSCLASS='LG' and CLASS='02' and CODE='N+')"
    sql = sql & "        then 'N+' else 'N-' end"
    sql = sql & " else 'Unknown'"
    sql = sql & " end as TYPE "
    sql = sql & "from TBCME037 XL, TBCME018 E018 "
    sql = sql & "where E018.HINBAN=XL.RPHINBAN"
    sql = sql & "  and E018.MNOREVNO=XL.RPREVNUM"
    sql = sql & "  and E018.FACTORY=XL.RPFACT"
    sql = sql & "  and E018.OPECOND=XL.RPOPCOND"
    sql = sql & "  and XL.CRYNUM='" & CRYNUM & "'"

    '�f�[�^�̒��o
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount > 0 Then
        GetXlType = rs("TYPE")
    End If
    rs.Close
    Set rs = Nothing

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


'==========================================
' �A�Ԏ擾�֐�
'==========================================


'�T�v      :����w��No���擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :gouki         ,I  ,String    ,���@ID
'          :�߂�l        ,O  ,String    ,�V����w��No�̘A�ԕ���
'����      :
'����      :2001/06/20 �쐬  �쑺
Public Function GetNewID_Siji(GOUKI$) As String
Dim key As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc010b.bas -- Function GetNewID_Siji"

    ''�Œ蕔�𓾂�(���@+�N�x)
    key = GOUKI & (year(oraGetSysdate()) Mod 10)
    GetNewID_Siji = key & GetNewSeq(SEQ_HIKIAGE_SIJI, key) & "00"

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�������g�����ԍ��̘A�ԕ������擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :�߂�l        ,O  ,String    ,�V�������g�����ԍ��̘A�ԕ���
'����      :
'����      :2001/06/20 �쐬  �쑺
Public Function GetNewID_RemeltGenryo() As String
Dim key As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc010b.bas -- Function GetNewID_RemeltGenryo"

    ''�Œ蕔�͂Ȃ����߁A�u_�v�Ƃ���
    key = "_"
    GetNewID_RemeltGenryo = GetNewSeq(SEQ_RMLT_GENRYO, key)

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�����ԍ��̘A�ԕ������擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :gouki         ,I  ,String    ,���@ID
'          :�߂�l        ,O  ,String    ,�V�����ԍ��̘A�ԕ���
'����      :
'����      :2001/06/20 �쐬  �쑺
Public Function GetNewID_CryNum(GOUKI$) As String
Dim key As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc010b.bas -- Function GetNewID_CryNum"

    ''�Œ蕔�𓾂�(���@+�N�x)
    key = GOUKI & (year(oraGetSysdate()) Mod 10)
    GetNewID_CryNum = GetNewSeq(SEQ_CRYNUM, key)

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�T���v��No���擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :�߂�l        ,O  ,String    ,�V�T���v��No
'����      :
'����      :2001/06/20 �쐬  �쑺
Public Function GetNewID_SampleNo() As String
Dim key As String
Dim sql As String
Dim rs As OraDynaset
Dim newID As String
Dim firstID As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc010b.bas -- Function GetNewID_SampleNo"
    
    GetNewID_SampleNo = 0

    ''�Œ蕔�͂Ȃ����߁A�u_�v�Ƃ���
    key = "_"
    newID = GetNewSeq(SEQ_SAMPLENO, key)
    
'>>>>> �T���v����6���Ή� 2007/05/25 SETsw kubota -------------
    newID = SAMPLENO_HEAD & Format$(newID, "00000")
'<<<<< �T���v����6���Ή� 2007/05/25 SETsw kubota -------------
    
    firstID = newID
    Do
        sql = "select REPSMPLIDCS from XSDCS where (REPSMPLIDCS='" & newID & "') and (KTKBNCS='0')"
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            '�Y���Ȃ��Ȃ�A���̔ԍ����g���Ă悢
            rs.Close
            Exit Do
        End If
        rs.Close
        
        newID = GetNewSeq(SEQ_SAMPLENO, key)

'>>>>> �T���v����6���Ή� 2007/05/25 SETsw kubota -------------
        newID = SAMPLENO_HEAD & Format$(newID, "00000")
'<<<<< �T���v����6���Ή� 2007/05/25 SETsw kubota -------------
        
        If newID = firstID Then
            '�܂��Ȃ��͂������A�P���S�Đ������T���v���������ꍇ
            newID = 0
            Exit Do
        End If
    Loop

    GetNewID_SampleNo = newID

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�V���ȘA�Ԃ��擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :SeqCode       ,I  ,String    ,�A�Ԏ�ʊǗ��R�[�h
'          :key           ,I  ,String    ,�A�Ԏ�ʃR�[�h
'          :�߂�l        ,O  ,String    ,�A�ԕ�����
'����      :
'����      :2001/06/20 �쐬  �쑺
Private Function GetNewSeq(SeqCode$, key$) As String
Dim rs As OraDynaset
Dim sql As String
Dim seq As Long
Dim keta As Integer
Dim clrWhen As String
Dim clrAt As Date
Dim sysNow As Date
Dim dbIsMine As Boolean

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc010b.bas -- Function GetNewSeq"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    
    ''�A�ԊǗ� ����A�w��i�Ԃ̎d�l���R�[�h��V�������Ɏ��o��
    sql = "select CONTNUM, MAXFIG, NUMUNIT, CLRDATE from TBCMB015 where (CNTMNGCD='" & SeqCode & "') and (CNTNUMCD='" & key & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        ''������Ȃ�������A1�Ԃ�o�^���ĕԂ�
        seq = 1                     '�V�ԍ� = 1
        rs.Close
        
        ''�����𓾂�
        sql = "select MAXFIG from tbcmb015 where (cntmngcd='" & SeqCode & "') and (cntnumcd='DEF')"
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Debug.Print "GetNewSeq: �A�ԊǗ�TBL�ɓo�^����Ă��Ȃ� (" & SeqCode & ")"
            keta = 3    '����l��3���Ƃ��Ă����B
        Else
            keta = rs("MAXFIG")         '����
        End If
        rs.Close
        
        ''�A�ԊǗ��e�[�u���Ɏw��L�[�̍s��ǉ�����
        sql = "insert into tbcmb015 (cntmngcd,cntnumcd,contnum,maxfig,numunit,numname,clrdate,regdate,upddate) " & _
              "(select cntmngcd, '" & key & "', 1, " & keta & ", numunit, numname, sysdate, sysdate, sysdate" & _
              " From TBCMB015 where (cntmngcd='" & SeqCode & "') and (cntnumcd='DEF'))"
        OraDB.ExecuteSQL sql
    Else
        ''����������A�A�Ԃ�1�グ��
        seq = rs("CONTNUM") + 1     '���݂̔ԍ�+1
        keta = rs("MAXFIG")         '����
        clrWhen = rs("NUMUNIT")     '�N���A�^�C�~���O
        clrAt = rs("CLRDATE")       '�O��N���A��������
        rs.Close
        
        ''�����I�[�o�[�ɂȂ�����A�Ԃ�1�ɖ߂�
        If Len(CStr(seq)) > keta Then
            seq = 1
        End If
        
        ''�N���A���������Ă�����A�A�Ԃ�1�ɖ߂�
        sysNow = oraGetSysdate()
        Select Case clrWhen
          Case "Y"      '�N���N���A
                If year(clrAt) <> year(sysNow) Then
                    seq = 1
                End If
          Case "M"      '�����N���A
                If (year(clrAt) <> year(sysNow)) Or (month(clrAt) <> month(sysNow)) Then
                    seq = 1
                End If
          Case "D"
                If (year(clrAt) <> year(sysNow)) Or (month(clrAt) <> month(sysNow)) Or (day(clrAt) <> day(sysNow)) Then
                    seq = 1
                End If
        End Select
        
        ''�A�ԊǗ��e�[�u�����X�V����
        sql = "update tbcmb015 set" & _
              " contnum=" & seq & _
              ",upddate=sysdate"
        If seq = 1 Then     '�N���A�����ꍇ
            sql = sql & ",clrdate=sysdate"
        End If
        sql = sql & " where (cntmngcd='" & SeqCode & "') and (cntnumcd='" & key & "')"
        OraDB.ExecuteSQL sql
    End If
    
    If dbIsMine Then
        OraDBClose
    End If
    
    GetNewSeq = Format$(seq, String(keta, "0"))

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "==== ERROR SQL ===="
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�t�H�[������ʒ����Ɉړ�
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :frmObj        ,I   ,Form      ,�t�H�[���I�u�W�F�N�g
'����      :
Public Sub CenterForm(frmObj As Form)
    '' �t�H�[������ʒ����Ɉړ�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub CenterForm"

    With frmObj
        If .WindowState <> 2 Then
            .Left = (Screen.Width - .Width) / 2
            .Top = (Screen.Height - .Height) / 2
        End If
    End With
    

proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub

'�T�v      :�R���g���[���I�u�W�F�N�g�Ɍ��ݎ������Z�b�g����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :ctrlObj       ,I   ,Control   ,�R���g���[���I�u�W�F�N�g
'����      :
Public Sub SetPresentTime(ctrlObj As Control)
    '' ���ݎ����̎擾�ƃZ�b�g

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SetPresentTime"

    ctrlObj.Caption = Format$(Now, "yyyy/mm/dd hh:nn")

proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub


'�T�v      :�t�H�[���\�������i����ʈړ������j
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :frmOwnerObj   ,I   ,Form      ,�I�[�i�[�t�H�[���I�u�W�F�N�g�i�Ăяo�����j
'          :frmShowObj    ,I   ,Form      ,�\���t�H�[���I�u�W�F�N�g
'����      :
Public Sub ShowFormProc(frmOwnerObj As Form, frmShowObj As Form)

    '' �}�E�X�J�[�\�����������ɕύX

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub ShowFormProc"

    Screen.MousePointer = vbHourglass
    '' �t�H�[���̕\��
    frmShowObj.Show
    '' �I�[�i�[��ʂ��B��
    frmOwnerObj.Hide
    '' �}�E�X�J�[�\������ɖ߂�
    Screen.MousePointer = vbDefault


proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub



'�T�v      :�t�H�[���N���[�Y����(�O��ʖߏ���)
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :frmPrevObj    ,I   ,Form      ,�O�\���t�H�[���I�u�W�F�N�g�i������ɕ\��������t�H�[���j
'          :frmCloseObj   ,I   ,Form      ,�N���[�Y�t�H�[���I�u�W�F�N�g�i����t�H�[���j
'����      :
Public Sub CloseFormProc(frmPrevObj As Form, frmCloseObj As Form)

    '' �}�E�X�J�[�\�����������ɕύX

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub CloseFormProc"

    Screen.MousePointer = vbHourglass
    '' �t�H�[�����N���[�Y����
    Unload frmCloseObj
    DoEvents
    '' �O��ʂ�\������
    frmPrevObj.Show
    DoEvents
    '' �}�E�X�J�[�\������ɖ߂�
    Screen.MousePointer = vbDefault
    DoEvents


proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub


'�T�v      :�����J�n����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :[strMsg]      ,I   ,String    ,�\�����b�Z�[�W������
'����      :���Ԃ̂����鏈�����s���Ƃ��̑O�������s��
'           EndProcess()�ƕ��p����B
Public Sub BeginProcess(Optional strMsg As String = "")

    '' ���b�Z�[�W������ꍇ�A���b�Z�[�W��\������

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub BeginProcess"

    If strMsg <> "" Then
        MsgBox strMsg, vbOKOnly + vbInformation
    End If

    '' �}�E�X�J�[�\�����������ɕύX
    Screen.MousePointer = vbHourglass
    DoEvents

proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub

'�T�v      :�����I������
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :[strMsg]      ,I   ,String    ,�\�����b�Z�[�W������
'����      :���Ԃ̂����鏈�����s������̌㏈�����s���B
'           BeginProcess()�ƕ��p����B
Public Sub EndProcess(Optional strMsg As String = "")

    '' ���b�Z�[�W������ꍇ�A���b�Z�[�W��\������

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub EndProcess"

    If strMsg <> "" Then
        MsgBox strMsg, vbOKOnly + vbInformation
    
    End If

    '' �}�E�X�J�[�\������ɖ߂�
    Screen.MousePointer = vbDefault
    DoEvents

proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub





'�T�v      :�X�v���b�h�R���g���[���̏���������
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :ctrlObj       ,I   ,vaSpread   ,�X�v���b�h�R���g���[���I�u�W�F�N�g
'          :[lMaxRows]    ,I   ,Long      ,�X�v���b�h�̏����\���s��
'����      :
Public Sub SpCtrlInit(ctrlObj As vaSpread, Optional lMaxRows As Long = -1)
    
    '' �X�v���b�h�̏����\���s�����Z�b�g

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SpCtrlInit"

    If lMaxRows >= 0 Then
        ''�@�����\���s���w�肪����ꍇ�A�X�v���b�h�̏����\���s�����Z�b�g
        ctrlObj.MaxRows = lMaxRows
    End If
    '' �Z���̃��b�N�𔽉f����
    ctrlObj.Protect = True


proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub


'�T�v      :�X�v���b�h�ɍs�ǉ�
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :ctrlObj       ,I   ,vaSpread   ,�X�v���b�h�R���g���[���I�u�W�F�N�g
'����      :
Public Sub SpCtrlInsertRow(ctrlObj As vaSpread)

    Dim lSmpPos As Long


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SpCtrlInsertRow"

    ctrlObj.MaxRows = ctrlObj.MaxRows + 1
    lSmpPos = ctrlObj.MaxRows
    
    With ctrlObj
        .row = lSmpPos
        .row2 = lSmpPos + 1
        .BlockMode = True
        .Action = ActionInsertRow
        .BlockMode = False
    End With

proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub




'�T�v      :�R���g���[���̏�Ԃ�ݒ肷��
'���Ұ�    :�ϐ���        ,IO ,�^            ,����
'          :ctrlObj       ,I   ,Control       ,�R���g���[���I�u�W�F�N�g
'          :ctrlState     ,I   ,enm_CtrlStateKind ,�R���g���[���̏�Ԏw��
'          :[bClear]      ,I   ,Boolean       ,�R���g���[���e�L�X�g���e�̃N���A�w���iTrue�F�N���A False�F�N���A���Ȃ��j
'����      :
Public Sub CtrlEnabled(ctrlObj As Control, ctrlState As enm_CtrlStateKind, Optional bClear As Boolean = False)


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub CtrlEnabled"

    On Error Resume Next
    
    If TypeOf ctrlObj Is Frame Then      '' �R���g���[�����t���[���̏ꍇ
        '' �R���g���[���w���Ԃ��`�F�b�N
        Select Case ctrlState
            Case CTRL_DISABLE             '' �ҏW�s�̏ꍇ
                ctrlObj.Enabled = False         '' �I�u�W�F�N�g���g�p�s�ɂ���
            Case Else                       '' ���̑��̏ꍇ
                ctrlObj.Enabled = True          '' �I�u�W�F�N�g���g�p�\�ɂ���
        End Select
    Else                                    '' �R���g���[�����t���[���ȊO�̏ꍇ
        '' �R���g���[���w���Ԃ��`�F�b�N
        Select Case ctrlState
            Case CTRL_DISABLE             '' �ҏW�s�̏ꍇ
                ctrlObj.BackColor = COLOR_DISABLE    '' �w�i�F��\�����ڐF�ɂ���
                ctrlObj.Locked = True                   '' ���b�N����
                ctrlObj.TabStop = False                 '' �^�u�X�g�b�v���Ȃ�
            Case CTRL_DISABLE_GRAY        '' �ҏW�s��(�O���[�F�\��)�̏ꍇ
                ctrlObj.BackColor = COLOR_GRAY '' �w�i�F���O���[�F�ɂ���
                ctrlObj.Locked = True                   '' ���b�N����
                ctrlObj.TabStop = False                 '' �^�u�X�g�b�v���Ȃ�
            Case CTRL_WARNING               '' �x���w���̏ꍇ
                ctrlObj.BackColor = COLOR_WARNING      '' �w�i�F���x���F�ɂ���
                ctrlObj.Locked = False                  '' ���b�N���Ȃ�
                ctrlObj.TabStop = True                  '' �^�u�X�g�b�v����
            Case CTRL_DISABLE_WARNING               '' �x���w���ҏW�s�̏ꍇ
                ctrlObj.BackColor = COLOR_WARNING      '' �w�i�F���x���F�ɂ���
                ctrlObj.Locked = True                  '' ���b�N����
                ctrlObj.TabStop = False                 '' �^�u�X�g�b�v���Ȃ�
            Case CTRL_SELECTED
                ctrlObj.BackColor = COLOR_SELECTED    '' �w�i�F��I��F�ɂ���
                ctrlObj.Locked = True                   '' ���b�N����
                ctrlObj.TabStop = False                 '' �^�u�X�g�b�v���Ȃ�
            Case CTRL_ENABLE_YELLOW
                ctrlObj.BackColor = COLOR_YELLOW        '' �w�i�F���C�G���[�ɂ���
                ctrlObj.Locked = False                   '' ���b�N���Ȃ�
                ctrlObj.TabStop = True                   '' �^�u�X�g�b�v����
            Case CTRL_DISABLE_SKY
                ctrlObj.BackColor = COLOR_SKY      '' �w�i�F���x���F�ɂ���
                ctrlObj.Locked = True                  '' ���b�N����
                ctrlObj.TabStop = False                 '' �^�u�X�g�b�v���Ȃ�
            Case Else                       '' ���̑��̏�Ԏw��̏ꍇ
                ctrlObj.BackColor = COLOR_ENABLE        '' �w�i�F���E�C���h�E�̔w�i�F�ɂ���
                ctrlObj.Locked = False                   '' ���b�N���Ȃ�
                ctrlObj.TabStop = True                   '' �^�u�X�g�b�v����
        End Select
    
        ''�e�L�X�g�N���A�`�F�b�N
        If bClear = True Then
            ctrlObj = ""
        End If
        
    End If
    
    '' �R���g���[����Ԏw�����x���w���̏ꍇ
    Select Case ctrlState
    Case CTRL_WARNING
        ctrlObj.SetFocus     '' �t�H�[�J�X���Z�b�g����
    End Select
    
    On Error GoTo 0


proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub


'�T�v      :�X�v���b�h�R���g���[���̃Z���̏�Ԃ�ݒ肷��i�P��Z���j
'���Ұ�    :�ϐ���        ,IO ,�^                ,����
'          :ctrlObj       ,   ,Control           ,
'          :Col           ,   ,Long              ,
'          :Row           ,   ,Long              ,
'          :ctrlState     ,   ,enm_CtrlStateKind ,
'          :[bClear]      ,   ,Boolean           ,
'����      :
Public Sub SpCtrlEnabled(ctrlObj As Control, ByVal col As Long, ByVal row As Long, ctrlState As enm_CtrlStateKind, Optional bClear As Boolean = False)

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SpCtrlEnabled"

    SpCtrlBlockEnabled ctrlObj, col, row, col, row, ctrlState, bClear

proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub


'�T�v      :�X�v���b�h�R���g���[���̃Z���̏�Ԃ�ݒ肷��
'���Ұ�    :�ϐ���        ,IO ,�^                ,����
'          :ctrlObj       ,I   ,vaSpread           ,�X�v���b�h�R���g���[���I�u�W�F�N�g
'          :Col           ,I   ,Long              ,��@Min�ʒu�i�͈͎w��j
'          :Row           ,I   ,Long              ,�s�@Min�ʒu�i�͈͎w��j
'          :Col2          ,I   ,Long              ,��@Max�ʒu�i�͈͎w��j
'          :Row2          ,I   ,Long              ,�s�@Max�ʒu�i�͈͎w��j
'          :ctrlState     ,I   ,enm_CtrlStateKind ,�R���g���[���̏�Ԏw��
'          :[bClear]      ,I   ,Boolean           ,�R���g���[���e�L�X�g���e�̃N���A�w���iTrue�F�N���A False�F�N���A���Ȃ��j
'����      :
Public Sub SpCtrlBlockEnabled(ctrlObj As vaSpread, ByVal col As Long, ByVal row As Long, ByVal col2 As Long, ByVal row2 As Long, ctrlState As enm_CtrlStateKind, Optional bClear As Boolean = False)


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SpCtrlBlockEnabled"

    On Error Resume Next

    ctrlObj.col = col
    ctrlObj.row = row
    ctrlObj.col2 = col2
    ctrlObj.row2 = row2
    
    '' �X�v���b�h�R���g���[���̃Z�����w�肵����ԂɃZ�b�g����B
    ctrlObj.BlockMode = True
    Select Case ctrlState
        Case CTRL_DISABLE         '' �ҏW�s�̏ꍇ
            ctrlObj.BackColor = COLOR_DISABLE           '' �w�i�F��\�����ڐF�ɂ���
            ctrlObj.Lock = True                         '' ���b�N����
        Case CTRL_DISABLE_GRAY    '' �ҏW�s��(�O���[�F�\��)�̏ꍇ
            ctrlObj.BackColor = COLOR_GRAY      '' �w�i�F���O���[�F�ɂ���
            ctrlObj.Lock = True                         '' ���b�N����
        Case CTRL_DISABLE_SKY    '' �ҏW�s��(����l)�̏ꍇ
            ctrlObj.BackColor = COLOR_SKY      '' �w�i�F�𐄒�F�ɂ���
            ctrlObj.Lock = True                         '' ���b�N����
        'Add Start 2010/08/04 SMPK Nakamura
        Case CTRL_ENABLE_SKY     '' �ҏW��(����l)�̏ꍇ
            ctrlObj.BackColor = COLOR_SKY      '' �w�i�F�𐄒�F�ɂ���
            ctrlObj.Lock = False                         '' ���b�N����
        'Add End 2010/08/04 SMPK Nakamura
        Case CTRL_WARNING           '' �x���w���̏ꍇ
            ctrlObj.BackColor = COLOR_WARNING           '' �w�i�F��ԐF�\���ɂ���
            ctrlObj.Lock = False                        '' ���b�N���Ȃ�
        Case CTRL_DISABLE_WARNING           '' �x���w���ҏW�s�̏ꍇ
            ctrlObj.BackColor = COLOR_WARNING           '' �w�i�F��ԐF�\���ɂ���
            ctrlObj.Lock = True                        '' ���b�N����
        Case CTRL_SELECTED
            ctrlObj.BackColor = COLOR_SELECTED          '' �w�i�F��I��F�ɂ���
            ctrlObj.Lock = True                         '' ���b�N����
        Case CTRL_ENABLE_GRAY    '' �ҏW��(�O���[�F�\��)�̏ꍇ
            ctrlObj.BackColor = COLOR_GRAY      '' �w�i�F���O���[�F�ɂ���
            ctrlObj.Lock = False                        '' ���b�N���Ȃ�
        Case CTRL_ENABLE_YELLOW
            ctrlObj.BackColor = COLOR_YELLOW        '' �w�i�F���C�G���[�ɂ���
            ctrlObj.Lock = False                   '' ���b�N���Ȃ�
        Case CTRL_DISABLE_YELLOW
            ctrlObj.BackColor = COLOR_YELLOW        '' �w�i�F���C�G���[�ɂ���
            ctrlObj.Lock = True                   '' ���b�N����
        '------ kuramoto �ǉ� 2001/09/25 ------
        Case CTRL_ENABLE_RED
            ctrlObj.BackColor = COLOR_RED        '' �w�i�F�����b�h�ɂ���
            ctrlObj.Lock = False                  '' ���b�N���Ȃ�
        Case CTRL_DISABLE_RED
            ctrlObj.BackColor = COLOR_RED        '' �w�i�F�����b�h�ɂ���
            ctrlObj.Lock = True                  '' ���b�N����
        '--------------------------------------
        Case Else                   '' ���̑��̏�Ԏw��̏ꍇ
            ctrlObj.BackColor = COLOR_ENABLE            '' �w�i�F�𔒐F�\���ɂ���
            ctrlObj.Lock = False                        '' ���b�N���Ȃ�
    End Select
    ctrlObj.BlockMode = False

    ''�e�L�X�g�N���A�`�F�b�N
    If bClear = True Then
        Dim iCol As Long
        Dim IRow As Long
        For IRow = row To row2
            For iCol = col To col2
                ctrlObj.SetText iCol, IRow, ""
            Next iCol
        Next IRow
    End If
    
    '' �Z���̃��b�N�𔽉f����
    ctrlObj.Protect = True
    
    
    '' �����ΏۃT���v�����A�N�e�B�u�Z���ɂ���
    Select Case ctrlState
        Case CTRL_WARNING           '' �x���w���̏ꍇ
        SpCtrlSetAction ctrlObj, col, row, col2, row2, ActionSelectBlock
    End Select
    
    On Error GoTo 0
    

proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub


'�T�v      :�X�v���b�h�ɃR���{�{�b�N�X��ݒ肷��
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :ctrlObj       ,I   ,vaSpread  ,�X�v���b�h�R���g���[���I�u�W�F�N�g
'          :Col           ,I   ,Long     ,��@Min�ʒu�i�͈͎w��j
'          :Row           ,I   ,Long     ,�s�@Min�ʒu�i�͈͎w��j
'          :Col2          ,I   ,Long     ,��@Max�ʒu�i�͈͎w��j
'          :Row2          ,I   ,Long     ,�s�@Max�ʒu�i�͈͎w��j
'          :strItem       ,I   ,String   ,�R���{�{�b�N�X�\�����ړ��e������iTab��؂�j
'          :[lParam]      ,I   ,Long     ,�R���{�{�b�N�X�����\�����ڎw��
'����      :
Public Sub SpCtrlSetCombo(ctrlObj As vaSpread, ByVal col As Long, ByVal row As Long, strItem As String, Optional lParam As Long = 0)

    Dim x As Long
    Dim Y As Long


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SpCtrlSetCombo"

    On Error Resume Next

    ctrlObj.col = col
    ctrlObj.row = row
    
    ctrlObj.CellType = CellTypeComboBox
    ctrlObj.TypeComboBoxList = strItem
    ctrlObj.TypeComboBoxCurSel = lParam
    
    On Error GoTo 0


proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub

'�T�v      :�w�肵������̃X�v���b�h�������s��
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :ctrlObj       ,I   ,vaSpread  ,�X�v���b�h�R���g���[���I�u�W�F�N�g
'          :Col           ,I   ,Long     ,��@Min�ʒu�i�͈͎w��j
'          :Row           ,I   ,Long     ,�s�@Min�ʒu�i�͈͎w��j
'          :Col2          ,I   ,Long     ,��@Max�ʒu�i�͈͎w��j
'          :Row2          ,I   ,Long     ,�s�@Max�ʒu�i�͈͎w��j
'          :iAction       ,I   ,Integer  ,�X�v���b�h��������w��
'����      :
Public Sub SpCtrlSetAction(ctrlObj As vaSpread, ByVal col As Long, ByVal row As Long, ByVal col2 As Long, ByVal row2 As Long, iAction As Integer)


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SpCtrlSetAction"

    On Error Resume Next

    '' �w�肵������̃X�v���b�h�������s��
    ctrlObj.col = col
    ctrlObj.row = row
    ctrlObj.col2 = col2
    ctrlObj.row2 = row2
    ctrlObj.BlockMode = True
    ctrlObj.Action = iAction
    ctrlObj.BlockMode = False
    
    On Error GoTo 0


proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub

'�T�v      :�w��Z���̃��b�N��Ԃ��擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :ctrlObj       ,I   ,Control  ,�X�v���b�h�R���g���[���I�u�W�F�N�g
'          :Col           ,I   ,Long     ,��@Min�ʒu�i�͈͎w��j
'          :Row           ,I   ,Long     ,�s�@Min�ʒu�i�͈͎w��j
'          :�߂�l        ,O  ,Boolean   ,True:���b�N     False:���b�N����Ă��Ȃ�
'����      :
Public Function SpCtrlIsLock(ctrlObj As vaSpread, ByVal col As Long, ByVal row As Long) As Boolean


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Function SpCtrlIsLock"

    ctrlObj.col = col
    ctrlObj.row = row

    SpCtrlIsLock = ctrlObj.Lock


proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�w��Z�����}�[�L���O����
'���Ұ�    :�ϐ���        ,IO ,�^                ,����
'          :ctrlObj       ,I   ,vaSpread           ,�X�v���b�h�R���g���[���I�u�W�F�N�g
'          :Col           ,I   ,Long              ,��@Min�ʒu�i�͈͎w��j
'          :Row           ,I   ,Long              ,�s�@Min�ʒu�i�͈͎w��j
'          :Col2          ,I   ,Long              ,��@Max�ʒu�i�͈͎w��j
'          :Row2          ,I   ,Long              ,�s�@Max�ʒu�i�͈͎w��j
'����      :
Public Sub SpCtrlSetMark(ctrlObj As vaSpread, ByVal col As Long, ByVal row As Long, ByVal col2 As Long, ByVal row2 As Long)
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SpCtrlSetMark"

    On Error Resume Next

    ctrlObj.col = col
    ctrlObj.row = row
    ctrlObj.col2 = col2
    ctrlObj.row2 = row2
    
    '' �X�v���b�h�R���g���[���̃Z�����w�肵����ԂɃZ�b�g����B
    ctrlObj.BlockMode = True
    ctrlObj.BackColor = &HFFFF80       '' ��F
    ctrlObj.BlockMode = False
   
    '' �Z���̃��b�N�𔽉f����
    ctrlObj.Protect = True
    
    On Error GoTo 0

proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub


'�T�v      :�w��Z���̃��b�N��ݒ�E��������
'���Ұ�    :�ϐ���        ,IO ,�^                ,����
'          :ctrlObj       ,I   ,vaSpread           ,�X�v���b�h�R���g���[���I�u�W�F�N�g
'          :Col           ,I   ,Long              ,��@Min�ʒu�i�͈͎w��j
'          :Row           ,I   ,Long              ,�s�@Min�ʒu�i�͈͎w��j
'          :Col2          ,I   ,Long              ,��@Max�ʒu�i�͈͎w��j
'          :Row2          ,I   ,Long              ,�s�@Max�ʒu�i�͈͎w��j
'          :[bLock]       ,I   ,Boolean           ,�R���g���[���̏�Ԏw��(True�F���b�N False�F���b�N���Ȃ�)
'          :[bClear]      ,I   ,Boolean           ,�R���g���[���e�L�X�g���e�̃N���A�w���iTrue�F�N���A False�F�N���A���Ȃ��j
'����      :
Public Sub SpCtrlSetLock(ctrlObj As vaSpread, ByVal col As Long, ByVal row As Long, ByVal col2 As Long, ByVal row2 As Long, Optional Block As Boolean = True, Optional bClear As Boolean = False)
    
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SpCtrlSetLock"

    On Error Resume Next

    ctrlObj.col = col
    ctrlObj.row = row
    ctrlObj.col2 = col2
    ctrlObj.row2 = row2
    
    '' �X�v���b�h�R���g���[���̃Z�����w�肵����ԂɃZ�b�g����B
    ctrlObj.BlockMode = True
    ctrlObj.Lock = Block
    ctrlObj.BlockMode = False

    ''�e�L�X�g�N���A�`�F�b�N
    If bClear = True Then
        Dim iCol As Long
        Dim IRow As Long
        For IRow = row To row2
            For iCol = col To col2
                ctrlObj.SetText iCol, IRow, ""
            Next iCol
        Next IRow
    End If
    
    '' �Z���̃��b�N�𔽉f����
    ctrlObj.Protect = True
    
    On Error GoTo 0

proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub

'�T�v      :�w��Z���̃t�H���g�𑾎��ɂ���
'���Ұ�    :�ϐ���        ,IO ,�^                ,����
'          :ctrlObj       ,I   ,vaSpread           ,�X�v���b�h�R���g���[���I�u�W�F�N�g
'          :Col           ,I   ,Long              ,��@Min�ʒu�i�͈͎w��j
'          :Row           ,I   ,Long              ,�s�@Min�ʒu�i�͈͎w��j
'          :Col2          ,I   ,Long              ,��@Max�ʒu�i�͈͎w��j
'          :Row2          ,I   ,Long              ,�s�@Max�ʒu�i�͈͎w��j
'          :[bState]      ,I   ,Boolean           ,������Ԏw���iTrue:�����w�肠�� False:�����w��Ȃ��j
'����      :
Public Sub SpCtrlFontBold(ctrlObj As vaSpread, ByVal col As Long, ByVal row As Long, ByVal col2 As Long, ByVal row2 As Long, Optional bState As Boolean = False)
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc020a.bas -- Sub SpCtrlFontBold"

    On Error Resume Next
    
    ctrlObj.col = col
    ctrlObj.row = row
    ctrlObj.col2 = col2
    ctrlObj.row2 = row2
    ctrlObj.BlockMode = True
    ctrlObj.FontBold = bState
    ctrlObj.BlockMode = False
    On Error GoTo 0

proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub



'�T�v      :�t�H�[���̑S�e�L�X�g�{�b�N�X�ɂ��āA.Text��RTrim����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :frm           ,I  ,Form      ,�Ώۃt�H�[��
'����      :
'����      :2001/08/24 �쐬  �쑺
Public Sub TrimAll(frm As Form)
Dim ctl As Control

    For Each ctl In frm.Controls
        If TypeName(ctl) = "TextBox" Then
            ctl.Text = RTrim$(ctl.Text)
        End If
    Next
End Sub


'�T�v      :�R���g���[���R�[�h���X�y�[�X�ɒu��������
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :s             ,I  ,String    ,��������
'����      :Debug�����SQL�ُ�̑Ώ��p
'����      :2001/09/26 �쐬  �쑺
Public Function toNormalStr(s$) As String
Dim i As Integer

    On Error Resume Next
    For i = 1 To Len(s)
        If Asc(Mid$(s, i, 1)) < &H20 Then
            Debug.Print "toNormalStr(""" & s & """) : " & i & "������=&H" & Asc(Mid$(s, i, 1))
            Mid$(s, i, 1) = " "
        End If
    Next
    toNormalStr = s
End Function


'�T�v      :��R�l�\���p�̏��������������߂�
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :rs            ,I  ,Double    ,��R�l
'          :[IgnoreZero]  ,I  ,Boolean   ,0���^����ꂽ�Ƃ���1��Ԃ��iFalse:5��Ԃ�)
'����      :��R�l�\�������𓝈ꂷ�邽�߁B1�`5���́A�����\�����������߂�
'����      :2002/1/15 �쐬  �쑺
Public Function GetLowerCol(ByVal rs As Double, Optional IgnoreZero As Boolean = False) As Integer
    rs = Abs(rs)
    If rs = 0 Then
        If IgnoreZero Then
            '0�𖳌����͂Ƃ݂Ȃ��ĂP��Ԃ�
            GetLowerCol = 1
        Else
            '0�𐔒l�Ƃ݂Ȃ��ĂT��Ԃ�
            GetLowerCol = 5
        End If
    ElseIf rs >= 10000 Then
        GetLowerCol = 1
    ElseIf rs >= 1000 Then
        GetLowerCol = 2
    ElseIf rs >= 100 Then
        GetLowerCol = 3
    ElseIf rs >= 10 Then
        GetLowerCol = 4
    ElseIf rs > 0 Then
        GetLowerCol = 5
    Else
        '�}�C�i�X�l�͓����Ă��Ȃ��͂�
        GetLowerCol = 0
    End If
End Function



'�T�v      :��R�̒l��\���p�ɕ����񉻂���(�L��6��+�����_)
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :rs_1          ,I  ,Double    ,��R�l1
'          :[rs_2]        ,I  ,Double    ,��R�l2
'����      :��R�l�\�������𓝈ꂷ�邽�߁B��R�l2������ƁA�͈͕������Ԃ�
'����      :2001/12/21 �쐬  �쑺
Public Function toRsStr(rs_1 As Double, Optional rs_2 As Double = -1#) As String
Dim s$, rsStr$

    If rs_1 >= 99999.9 Then
        s = "99999.9"
    Else
        s = Format$(rs_1, "0." & String(GetLowerCol(rs_1), "0"))
    End If
    rsStr = s
    
    If rs_2 >= 0 Then
        rsStr = rsStr & "-"
        
        If rs_2 >= 99999.9 Then
            s = "99999.9"
        Else
            s = Format$(rs_2, "0." & String(GetLowerCol(rs_2), "0"))
        End If
        rsStr = rsStr & s
    End If
    
    toRsStr = rsStr
End Function

'�T�v      :��R�̒l��\���p�ɕ����񉻂���(�w��̏����_�ȉ�����)
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :rs            ,I  ,Double    ,��R�l
'          :place         ,I  ,Integer   ,�����_�ȉ�����
'����      :��R�l�\�������𓝈ꂷ�邽�߁B<0�̂Ƃ��͋󕶎����Ԃ�
'����      :2002/1/16 �쐬  �쑺
'����      :2002/1/17 S.Sano
Public Function toRsStrByPlace(rs As Double, place As Integer) As String
Dim s$

    If rs < 0 Then
        s = vbNullString
'2002/01/17 S.Sano    ElseIf rs >= 99999.9 Then
'2002/01/17 S.Sano        s = "99999.9"
    Else
        s = Format$(rs, "0." & String(place, "0"))
        If val(s) >= 100000 Then
            s = "99999." & String(place, "9")
        End If
    End If
    toRsStrByPlace = s
End Function

Public Function toRsStr_nl(rs_1 As Double, Optional rs_2 As Double = -1#) As String '��R�̕\�� 2003/12/8
Dim s$, rsStr$
    
    If rs_1 < 0 Then  '-1(Null)�̂Ƃ�
        s = vbNullString
    Else
        If rs_1 >= 99999.9 Then
            s = "99999.9"
        Else
            s = Format$(rs_1, "0." & String(GetLowerCol(rs_1), "0"))
        End If
    End If
            rsStr = s
    
    If rs_2 >= 0 Then
        rsStr = rsStr & "-"
        
        If rs_2 >= 99999.9 Then
            s = "99999.9"
        Else
            s = Format$(rs_2, "0." & String(GetLowerCol(rs_2), "0"))
        End If
        rsStr = rsStr & s
    Else    'rs_2��-1(Null)�̂Ƃ��̏���
        rsStr = rsStr & "-"
        s = vbNullString
        rsStr = rsStr & s
    End If
    
    toRsStr_nl = rsStr
End Function

'�T�v      :�X�v���b�h�̒�R�����\�����d�l�ɍ��킹�Č��𑵂���B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :targetSpread  ,I  ,vaSpread  ,�ΏۃX�v���b�h
'          :col1    ,I  ,Long      ,�ΏۃJ����(From)
'          :[col2]  ,I  ,Long      ,�ΏۃJ����(To)
'          :[row1]  ,I  ,Long      ,�Ώۍs(From)
'          :[row2]  ,I  ,Long      ,�Ώۍs(To)
'����      :�X�v���b�h�̑Ώ۔͈͂ɂ��āA�S�ẴZ���ŏ����_�ȉ������𑵂���
'����      :2002/01/15 �쐬 S.Sano
'          :2002/01/16 �C�� �쑺
Public Sub RsSpreadSet(targetSpread As vaSpread, col1 As Long, Optional col2 As Long = 0, Optional row1 As Long = 0, Optional row2 As Long = 0)
Dim row As Long
Dim col As Long
Dim MaxLowerCol As Integer
Dim lowCol As Integer
Dim rs As Double
    
    MaxLowerCol = 0
    '����͈͂̐ݒ�
    If row1 = 0 Then
        row1 = 1
        If row2 = 0 Then row2 = targetSpread.MaxRows
    ElseIf row2 = 0 Then
        row2 = row1
    End If
    If col2 = 0 Then
        col2 = col1
    End If
    
    With targetSpread
        .ReDraw = False
        
        '�\�����ׂ������_�ȉ����������߂�
        For col = col1 To col2
            For row = row1 To row2
                .GetFloat col, row, rs
                lowCol = GetLowerCol(rs, True)
                If MaxLowerCol < lowCol Then
                    MaxLowerCol = lowCol
                End If
            Next
        Next
        
        '�����_�ȉ������𑵂���
        .BlockMode = True
        .col = col1
        .col2 = col2
        .row = row1
        .row2 = row2
        .CellType = CellTypeFloat
        .TypeFloatMax = 99999.99999
        .TypeFloatMin = 0#
        '.TypeFloatMax = Val("99999." & Left("99999", MaxLowerCol))
        '.TypeFloatMin = 0
        .TypeFloatDecimalPlaces = MaxLowerCol
        .BlockMode = False
        
        .ReDraw = True
    End With
End Sub

Public Sub WriteDBLog(ByVal sqlStr$, Optional ByVal memo$ = " ")
Dim dbIsMine    As Boolean
Dim sql         As String
Dim s           As String
Dim i           As Integer
Dim hostname    As String
Dim fncName     As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    fncName = gErr.fncName
    gErr.Push "s_cmzc004c.bas -- Function WriteDBLog"

    ' �{�������͂��̊֐����擾�s�ɑΉ�
    ' ���G���[�Ɋւ���Push�APop�̏����񐔕s��v�Ŕ�������\���L��
    If Trim(fncName) = "" Then
        fncName = "fncName is Nothing"
    ElseIf fncName = vbNullString Then
        fncName = "fncName is Null"
    End If

#If DBG Then
Dim fno As Integer
    fno = FreeFile
    Open App.Path & "\" & App.EXENAME & ".LOG" For Append As fno
    Print #fno, Now, fncName, memo
    Print #fno, "    " & sqlStr
    Close fno
#End If

    ''�^����ꂽSQL���̃V���O���N�H�[�g��u��������
    s = Replace(sqlStr, "'", "''")
    
    ''�z�X�g���𓾂�
    hostname = String(51, " ")
    GetComputerName hostname, 50
    If InStr(1, hostname, vbNullChar) Then
        hostname = Trim$(Left$(hostname, InStr(1, hostname, vbNullChar) - 1))
    End If

    ''SQL���쐬
    If sqlStr = vbNullString Then sqlStr = " "
    If memo = vbNullString Then memo = " "
    sql = "insert into TBCMC003 " & _
        "(L_DATE, SEQ, HOSTNAME, APPNAME, FNCNAME, SQL, MEMO) values ("
    sql = sql & "sysdate, "                 '�^�C���X�^���v
    sql = sql & "LOG_SEQ.NEXTVAL, "         'SEQ
    sql = sql & "'" & hostname & "', "      '�[����
    sql = sql & "'" & App.EXENAME & "', "   'APPNAME
    sql = sql & "'" & fncName & "', "       '�֐���
    sql = sql & "'" & s & "', "             'SQL
    sql = sql & "'" & memo & "' "           'memo
    sql = sql & ")"
    
    ''Log������
    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    OraDB.ExecuteSQL sql
    If dbIsMine Then
        OraDBClose
    End If
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub

'�T�v      :�w��̈�̕i�Ԃ�����������i�i�ԊǗ�Tbl�Ώہj
'���Ұ�    :�ϐ���        ,IO ,�^          ,����
'          :BLOCKID       ,I  ,String      ,�����ԍ�
'          :hin           ,I  ,tFullHinban ,����������i��
'          :�߂�l        ,O  ,FUNCTION_RETURN,
'����      :
'����      :2003/10/29 ��n
Public Function ChangeXSDCSHinban(BLOCKID$, HIN As tFullHinban) As FUNCTION_RETURN
Dim rs As OraDynaset
Dim sql As String

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc010a.bas -- Function ChangeAreaHinban"

    ChangeXSDCSHinban = FUNCTION_RETURN_FAILURE
          
    ''�w��̈�̊J�n�ʒu���܂ޕi�Ԃ�����΂���𒲐�����
    sql = "update XSDCS "
    sql = sql & "set HINBCS = '" & HIN.hinban & "',"
    sql = sql & "REVNUMCS = '" & HIN.mnorevno & "',"
    sql = sql & "FACTORYCS = '" & HIN.factory & "',"
'    sql = sql & "OPECS = " & HIN.OPECOND & ","
    sql = sql & "OPECS = '" & HIN.opecond & "',"    '�ݸ�ٸ��Ēǉ� 2009/11/16 SETsw Nakada
    sql = sql & "KDAYCS = sysdate,"
    sql = sql & "KSTAFFCS = '" & STAFFIDBUFF & "'"
    sql = sql & " WHERE LIVKCS = '0' AND "
    sql = sql & "CRYNUMCS = '" & BLOCKID$ & "'"
    
Debug.Print sql
    If OraDB.ExecuteSQL(sql) < 1 Then
        Debug.Print "�Y������u���b�N����������"
    Else
    ''    WriteDBLog sql
    End If
    
      
    ChangeXSDCSHinban = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�R���{�{�b�N�X�Ɍ������ƃ}�X�^���̑I������ݒ肷��
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :SYSCLASS      ,I  ,String    ,SYS�敪
'          :Class         ,I  ,String    ,�敪
'          :FieldName     ,I  ,String    ,�I�������̃t�B�[���h��
'          :cmb           ,O  ,ComboBox  ,�ݒ��̃R���{�{�b�N�X
'          :�߂�l        ,O  ,FUNCTION_RETURN,
'����      :
'����      :2005/06/02 KODA9
Public Function SetCodeComboA9(ByVal SYSCLASS$, ByVal Class$, ByVal FieldName$, cmb As ComboBox) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String
Dim max As Integer
Dim i As Integer

    ''�R�[�h�}�X�^����A�w��̃t�B�[���h������
    sql = "SELECT CODEA9 , " & FieldName & " from KODA9" & _
          " WHERE (SYSCA9='" & SYSCLASS & "') and (SHUCA9='" & Class & "')" & _
          " Order by CTR01A9"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''������Ȃ�������ANG
        cmb.Clear
        SetCodeComboA9 = FUNCTION_RETURN_FAILURE
    Else
        ''����������A�R���{�{�b�N�X�ɑI������ݒ肷��
        With cmb
            .Clear
            max = rs.RecordCount
            For i = 1 To max
                .AddItem GetGPCodeDspStr(rs("CODEA9"), rs(FieldName))
                rs.MoveNext
            Next
        End With
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

End Function
'�T�v      :�������Ɨp�R�[�h���������A�w��t�B�[���h�̓��e�𓾂�
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :SYSCLASS      ,I  ,String    ,�V�X�e���敪
'          :CLASS         ,I  ,String    ,�敪
'          :CODE          ,I  ,String    ,�R�[�h
'          :FieldName     ,I  ,String    ,�t�B�[���h��
'          :�߂�l        ,O  ,String    ,�t�B�[���h���e
'����      :
'����      :2005/06/02  KODA9
Public Function GetCodeFieldA9(ByVal SYSCLASS$, ByVal Class$, ByVal CODE$, ByVal FieldName$) As String
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String

    ''�R�[�h�}�X�^����A�w��̃t�B�[���h������
    sql = "SELECT " & FieldName & " from KODA9 WHERE (SYSCA9='" & SYSCLASS & "') and (SHUCA9='" & Class & "') and (CODEA9='" & CODE & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        ''������Ȃ�������AVbNullString��Ԃ�
        GetCodeFieldA9 = vbNullString
    Else
        ''����������A�w��t�B�[���h�̓��e��Ԃ�
        If IsNull(rs(FieldName)) = False Then
            GetCodeFieldA9 = Trim$(rs(FieldName))
        End If
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

End Function


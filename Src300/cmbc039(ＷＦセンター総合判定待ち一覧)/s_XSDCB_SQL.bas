Attribute VB_Name = "s_XSDCB_SQL"
'��������(SXL) (XSDCB) �����֐�

'��TEST�p

'***�e�[�u���uXSDCB�v�ւ̃f�[�^�A�N�Z�X�֐�***
'������ ���Ұ��ɒl��Ă��鎞�A�܂��S�ď��������邱��

Option Explicit

'����������(SXL)
Public Type typ_XSDCB
    SXLIDCB As String * 1        ' SXLID
    KCNTCB As Integer            ' �H���A��
    XTALCB As String * 12        ' �����ԍ�
    INPOSCB As Integer           ' �������J�n�ʒu
    LENCB As Integer             ' ����
    HINBCB As String * 8         ' �i��
    REVNUMCB As Integer          ' �d�b�ԍ������ԍ�
    FACTORYCB As String * 1      ' �H��
    OPECB As String * 1          ' ���Ə���
    MAICB As Integer             ' ������
    WSRMAICB As Integer          ' WS��㖇��
    WSNMAICB As Integer          ' WS��򌇗�����
    WFCMAICB As Integer          ' �������
    SXLRMAICB As Integer         ' SXL�w��(�Ǖi)
    SXLNMAICB As Integer         ' SXL�w��(�s��)
    WFCNMAICB As Integer         ' WFC����������
    SXLEMAICB As Integer         ' SXL�m�薇��
    SRMAICB As Integer           ' �T���v�����w��(�Ǖi)
    SNMAICB As Integer           ' �T���v�����w��(�s��)
    STMAICB As Integer           ' �T���v������
    FURIMAICB As Integer         ' �U�֖���
    XTWORKCB As String * 2       ' �����H��
    WFWORKCB As String * 2       ' �E�F�[�n����
    FURYCCB As String * 3        ' �s�Ǘ��R
    LSTCCB As String * 1         ' �̎��ԋ敪
    LUFRCCB As String * 3        ' �i��R�[�h
    LUFRBCB As String * 1        ' �i��敪
    LDERCCB As String * 3        ' �i���R�[�h
    LDFRBCB As String * 1        ' �i���敪
    HOLDCCB As String * 3        ' �z�[���h�R�[�h
    HOLDBCB As String * 1        ' �z�[���h�敪
    EXKUBCB As String * 1        ' ��O�敪
    HENPKCB As String * 1        ' �ԕi�敪
    LIVKCB As String * 1         ' �����敪
    KANKCB As String * 1         ' �����敪
    NFCB As String * 1           ' ���ɋ敪
    SAKJCB As String * 1         ' �폜�敪
    TDAYCB As Date               ' �o�^���t
    KDAYCB As Date               ' �X�V���t
    SUMITCB As String * 1        ' SUMIT���M�t���O
    SNDKCB As String * 1         ' �ԕi�敪
    SNDAYCB As Date              ' ���M���t
    NEWKNTCB As String           ' �ŏI�ʉߍH��
    GNWKNTCB As String           ' ���ݍH��
    MOTHINCB As String           ' ���i��
    RLENCB As Integer            ' ���_����
    SHOLDCLSCB As String         ' �z�[���h�敪(SXL�m��)
    PLANTCATCB As String         ' ����
    KBLKFLGCB As String * 1      ' �֘A��ۯ��׸ށ@08/01/31 ooba
End Type

'�X�V�p
Public Type typ_XSDCB_Update
    SXLIDCB As String            ' SXLID
    KCNTCB As String             ' �H���A��
    XTALCB As String             ' �����ԍ�
    INPOSCB As String            ' �������J�n�ʒu
    LENCB As String              ' ����
    HINBCB As String             ' �i��
    REVNUMCB As String           ' �d�b�ԍ������ԍ�
    FACTORYCB As String          ' �H��
    OPECB As String              ' ���Ə���
    MAICB As String              ' ������
    WSRMAICB As String           ' WS��㖇��
    WSNMAICB As String           ' WS��򌇗�����
    WFCMAICB As String           ' �������
    SXLRMAICB As String          ' SXL�w��(�Ǖi)
    SXLNMAICB As String          ' SXL�w��(�s��)
    WFCNMAICB As String          ' WFC����������
    SXLEMAICB As String          ' SXL�m�薇��
    SRMAICB As String            ' �T���v�����w��(�Ǖi)
    SNMAICB As String            ' �T���v�����w��(�s��)
    STMAICB As String            ' �T���v������
    FURIMAICB As String          ' �U�֖���
    XTWORKCB As String           ' �����H��
    WFWORKCB As String           ' �E�F�[�n����
    FURYCCB As String            ' �s�Ǘ��R
    LSTCCB As String             ' �̎��ԋ敪
    LUFRCCB As String            ' �i��R�[�h
    LUFRBCB As String            ' �i��敪
    LDERCCB As String            ' �i���R�[�h
    LDFRBCB As String            ' �i���敪
    HOLDCCB As String            ' �z�[���h�R�[�h
    HOLDBCB As String            ' �z�[���h�敪
    EXKUBCB As String            ' ��O�敪
    HENPKCB As String            ' �ԕi�敪
    LIVKCB As String             ' �����敪
    KANKCB As String             ' �����敪
    NFCB As String               ' ���ɋ敪
    SAKJCB As String             ' �폜�敪
    TDAYCB As String             ' �o�^���t
    KDAYCB As String             ' �X�V���t
    SUMITCB As String            ' SUMIT���M�t���O
    SNDKCB As String             ' �ԕi�敪
    SNDAYCB As String            ' ���M���t
    NEWKNTCB As String           ' �ŏI�ʉߍH��
    GNWKNTCB As String           ' ���ݍH��
    MOTHINCB As String           ' ���i��
    RLENCB As String             ' ���_����
    SHOLDCLSCB As String         ' �z�[���h�敪(SXL�m��)
    PLANTCATCB As String         ' ����
    KBLKFLGCB As String          ' �֘A��ۯ��׸ށ@08/01/31 ooba
End Type

'��SELECT��
'*******************************************************************************************
'*    �֐���        : DBDRV_GetXSDCB
'*
'*    �����T�v      : 1.�e�[�u���uXSDCB�v��������ɂ��������R�[�h�𒊏o����
'*
'*    �p�����[�^    : �ϐ���       ,IO   ,�^            ,����
'*                   records()     ,O    ,typ_XSDCB     ,���o���R�[�h
'*                   sqlWhere      ,I    ,String        ,���o����(SQL��Where��:�ȗ��\)
'*                   sqlOrder      ,I    ,String        ,���o����(SQL��Order by��:�ȗ��\)
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Public Function DBDRV_GetXSDCB(records() As typ_XSDCB, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN

    Dim sSql        As String       ' SQL�S��
    Dim sSqlBase    As String       ' SQL��{��(WHERE�߂̑O�܂�)
    Dim rs          As OraDynaset   ' RecordSet
    Dim intRecCnt   As Long         ' ���R�[�h��
    Dim i           As Long

    ' SQL��g�ݗ��Ă�
    sSqlBase = "Select * From XSDCB"
    sSql = sSqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sSql = sSql & " " & sqlWhere & " " & sqlOrder
    End If

    ' �f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetXSDCB = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ' ���o���ʂ��i�[����
    intRecCnt = rs.RecordCount
    ReDim records(intRecCnt)
    If intRecCnt = 0 Then
        Exit Function
    End If
    For i = 1 To intRecCnt
        With records(i)
            If IsNull(rs.Fields("SXLIDCB")) = False Then .SXLIDCB = rs.Fields("SXLIDCB")
            If IsNull(rs.Fields("KCNTCB")) = False Then .KCNTCB = rs.Fields("KCNTCB")
            If IsNull(rs.Fields("XTALCB")) = False Then .XTALCB = rs.Fields("XTALCB")
            If IsNull(rs.Fields("INPOSCB")) = False Then .INPOSCB = rs.Fields("INPOSCB")
            If IsNull(rs.Fields("LENCB")) = False Then .LENCB = rs.Fields("LENCB")
            If IsNull(rs.Fields("HINBCB")) = False Then .HINBCB = rs.Fields("HINBCB")
            If IsNull(rs.Fields("REVNUMCB")) = False Then .REVNUMCB = rs.Fields("REVNUMCB")
            If IsNull(rs.Fields("FACTORYCB")) = False Then .FACTORYCB = rs.Fields("FACTORYCB")
            If IsNull(rs.Fields("OPECB")) = False Then .OPECB = rs.Fields("OPECB")
            If IsNull(rs.Fields("MAICB")) = False Then .MAICB = rs.Fields("MAICB")
            If IsNull(rs.Fields("WSRMAICB")) = False Then .WSRMAICB = rs.Fields("WSRMAICB")
            If IsNull(rs.Fields("WSNMAICB")) = False Then .WSNMAICB = rs.Fields("WSNMAICB")
            If IsNull(rs.Fields("WFCMAICB")) = False Then .WFCMAICB = rs.Fields("WFCMAICB")
            If IsNull(rs.Fields("SXLRMAICB")) = False Then .SXLRMAICB = rs.Fields("SXLRMAICB")
            If IsNull(rs.Fields("SXLNMAICB")) = False Then .SXLNMAICB = rs.Fields("SXLNMAICB")
            If IsNull(rs.Fields("WFCNMAICB")) = False Then .WFCNMAICB = rs.Fields("WFCNMAICB")
            If IsNull(rs.Fields("SXLEMAICB")) = False Then .SXLEMAICB = rs.Fields("SXLEMAICB")
            If IsNull(rs.Fields("SRMAICB")) = False Then .SRMAICB = rs.Fields("SRMAICB")
            If IsNull(rs.Fields("SNMAICB")) = False Then .SNMAICB = rs.Fields("SNMAICB")
            If IsNull(rs.Fields("STMAICB")) = False Then .STMAICB = rs.Fields("STMAICB")
            If IsNull(rs.Fields("FURIMAICB")) = False Then .FURIMAICB = rs.Fields("FURIMAICB")
            If IsNull(rs.Fields("XTWORKCB")) = False Then .XTWORKCB = rs.Fields("XTWORKCB")
            If IsNull(rs.Fields("WFWORKCB")) = False Then .WFWORKCB = rs.Fields("WFWORKCB")
            If IsNull(rs.Fields("FURYCCB")) = False Then .FURYCCB = rs.Fields("FURYCCB")
            If IsNull(rs.Fields("LSTCCB")) = False Then .LSTCCB = rs.Fields("LSTCCB")
            If IsNull(rs.Fields("LUFRCCB")) = False Then .LUFRCCB = rs.Fields("LUFRCCB")
            If IsNull(rs.Fields("LUFRBCB")) = False Then .LUFRBCB = rs.Fields("LUFRBCB")
            If IsNull(rs.Fields("LDERCCB")) = False Then .LDERCCB = rs.Fields("LDERCCB")
            If IsNull(rs.Fields("LDFRBCB")) = False Then .LDFRBCB = rs.Fields("LDFRBCB")
            If IsNull(rs.Fields("HOLDCCB")) = False Then .HOLDCCB = rs.Fields("HOLDCCB")
            If IsNull(rs.Fields("HOLDBCB")) = False Then .HOLDBCB = rs.Fields("HOLDBCB")
            If IsNull(rs.Fields("EXKUBCB")) = False Then .EXKUBCB = rs.Fields("EXKUBCB")
            If IsNull(rs.Fields("HENPKCB")) = False Then .HENPKCB = rs.Fields("HENPKCB")
            If IsNull(rs.Fields("LIVKCB")) = False Then .LIVKCB = rs.Fields("LIVKCB")
            If IsNull(rs.Fields("KANKCB")) = False Then .KANKCB = rs.Fields("KANKCB")
            If IsNull(rs.Fields("NFCB")) = False Then .NFCB = rs.Fields("NFCB")
            If IsNull(rs.Fields("SAKJCB")) = False Then .SAKJCB = rs.Fields("SAKJCB")
            If IsNull(rs.Fields("TDAYCB")) = False Then .TDAYCB = rs.Fields("TDAYCB")
            If IsNull(rs.Fields("KDAYCB")) = False Then .KDAYCB = rs.Fields("KDAYCB")
            If IsNull(rs.Fields("SUMITCB")) = False Then .SUMITCB = rs.Fields("SUMITCB")
            If IsNull(rs.Fields("SNDKCB")) = False Then .SNDKCB = rs.Fields("SNDKCB")
            If IsNull(rs.Fields("SNDAYCB")) = False Then .SNDAYCB = rs.Fields("SNDAYCB")
            If IsNull(rs.Fields("NEWKNTCB")) = False Then .NEWKNTCB = rs.Fields("NEWKNTCB")
            If IsNull(rs.Fields("GNWKNTCB")) = False Then .GNWKNTCB = rs.Fields("GNWKNTCB")
            If IsNull(rs.Fields("MOTHINCB")) = False Then .MOTHINCB = rs.Fields("MOTHINCB")
            If IsNull(rs.Fields("RLENCB")) = False Then .RLENCB = rs.Fields("RLENCB")
            If IsNull(rs.Fields("SHOLDCLSCB")) = False Then .SHOLDCLSCB = rs.Fields("SHOLDCLSCB")
            If IsNull(rs.Fields("PLANTCATCB")) = False Then .PLANTCATCB = rs.Fields("PLANTCATCB")   ' ���� 2007/09/04 SPK Tsutsumi Add
            If IsNull(rs.Fields("KBLKFLGCB")) = False Then .KBLKFLGCB = rs.Fields("KBLKFLGCB")      '08/01/31 ooba
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetXSDCB = FUNCTION_RETURN_SUCCESS
End Function

'��UPDATE��
'*******************************************************************************************
'*    �֐���        : UpdateXSDCB
'*
'*    �����T�v      : 1.�e�[�u���uXSDCB�v���X�V���� ptrn1
'*                    (�X�V���ڂ��\���̂ɃZ�b�g���Ĉ����n��)
'*
'*    �p�����[�^    : �ϐ���       ,IO   ,�^            ,����
'*                   records()     ,O    ,typ_XSDCB     ,�X�V���R�[�h
'*                   sqlWhere      ,I    ,String        ,���o����(SQL��Where��:�ȗ��\)
'*                   sqlOrder      ,I    ,String        ,���o����(SQL��Order by��:�ȗ��\)
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Public Function UpdateXSDCB(records As typ_XSDCB_Update, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
    On Error GoTo proc_err
'    gErr.Push "s_XSDCB_SQL.bas -- Function CreateXSDCB"

    Dim sSql        As String       ' SQL�S��
    Dim sSqlBase    As String       ' SQL��{��(WHERE�߂̑O�܂�)
    Dim rs          As OraDynaset   ' RecordSet
    Dim intRecCnt   As Long         ' ���R�[�h��
    Dim i           As Long
    Dim dtmNowtime  As Date

    dtmNowtime = getSvrTime()       ' �T�[�o�[�̎��Ԃ��擾����悤�ɕύX 2003/6/4 tuku

'>>>>> Edit-->UPDATE�ɕύX�@2009/07/21�@SSS.Marushita
    
    With records
        
        ''SQL��g�ݗ��Ă�
        sSql = "UPDATE XSDCB SET" & vbLf
        
        ''�X�V���t
        sSql = sSql & " KDAYCB = TO_DATE('" & Format$(dtmNowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
        
        ''SXLID
        If .SXLIDCB <> "" And left(.SXLIDCB, 1) <> vbNullChar Then
            sSql = sSql & ",SXLIDCB = '" & .SXLIDCB & "'" & vbLf
        End If
        
        ''�H���A��
        If .KCNTCB <> "" Then
            sSql = sSql & ",KCNTCB = '" & CStr(CInt(.KCNTCB)) & "'" & vbLf
        End If
        
        ''�����ԍ�
        If .XTALCB <> "" And left(.XTALCB, 1) <> vbNullChar Then
            sSql = sSql & ",XTALCB = '" & .XTALCB & "'" & vbLf
        End If
        
        ''�������J�n�ʒu
        If .INPOSCB <> "" Then
            sSql = sSql & ",INPOSCB = '" & CStr(CInt(.INPOSCB)) & "'" & vbLf
        End If
        
        ''����
        If .LENCB <> "" Then
            sSql = sSql & ",LENCB = '" & CStr(CInt(.LENCB)) & "'" & vbLf
        End If
        
        ''�i��
        If .HINBCB <> "" And left(.HINBCB, 1) <> vbNullChar Then
            sSql = sSql & ",HINBCB = '" & .HINBCB & "'" & vbLf
        End If
        
        ''���i�ԍ������ԍ�
        If .REVNUMCB <> "" Then
            sSql = sSql & ",REVNUMCB = '" & CStr(CInt(.REVNUMCB)) & "'" & vbLf
        End If
        
        ''�H��
        If .FACTORYCB <> "" And left(.FACTORYCB, 1) <> vbNullChar Then
            sSql = sSql & ",FACTORYCB = '" & .FACTORYCB & "'" & vbLf
        End If
        
        ''���Ə���
        If .OPECB <> "" And left(.OPECB, 1) <> vbNullChar Then
            sSql = sSql & ",OPECB = '" & .OPECB & "'" & vbLf
        End If
        
        ''������
        If .MAICB <> "" Then
            sSql = sSql & ",MAICB = '" & CStr(CInt(.MAICB)) & "'" & vbLf
        End If
        
        ''WS��㖇��
        If .WSRMAICB <> "" Then
            sSql = sSql & ",WSRMAICB = '" & CStr(CInt(.WSRMAICB)) & "'" & vbLf
        End If
        
        ''WS��㌇������
        If .WSNMAICB <> "" Then
            sSql = sSql & ",WSNMAICB = '" & CStr(CInt(.WSNMAICB)) & "'" & vbLf
        End If
        
        ''WFC�������
        If .WFCMAICB <> "" Then
            sSql = sSql & ",WFCMAICB = '" & CStr(CInt(.WFCMAICB)) & "'" & vbLf
        End If
        
        ''SXL�w���i�Ǖi�j
        If .SXLRMAICB <> "" Then
            sSql = sSql & ",SXLRMAICB = '" & CStr(CInt(.SXLRMAICB)) & "'" & vbLf
        End If
        
        ''SXL�w���i�s�ǁj
        If .SXLNMAICB <> "" Then
            sSql = sSql & ",SXLNMAICB = '" & CStr(CInt(.SXLNMAICB)) & "'" & vbLf
        End If
        
        ''WFC����������
        If .WFCNMAICB <> "" Then
            sSql = sSql & ",WFCNMAICB = '" & CStr(CInt(.WFCNMAICB)) & "'" & vbLf
        End If
        
        ''SXL�m�薇��
        If .SXLEMAICB <> "" Then
            sSql = sSql & ",SXLEMAICB = '" & CStr(CInt(.SXLEMAICB)) & "'" & vbLf
        End If
        
        ''�T���v�����w���i�Ǖi�j
        If .SRMAICB <> "" Then
            sSql = sSql & ",SRMAICB = '" & CStr(CInt(.SRMAICB)) & "'" & vbLf
        End If
        
        ''�T���v�����w���i�s�ǁj
        If .SNMAICB <> "" Then
            sSql = sSql & ",SNMAICB = '" & CStr(CInt(.SNMAICB)) & "'" & vbLf
        End If
        
        ''�T���v������
        If .STMAICB <> "" Then
            sSql = sSql & ",STMAICB = '" & CStr(CInt(.STMAICB)) & "'" & vbLf
        End If
        
        ''�U�֖���
        If .FURIMAICB <> "" Then
            sSql = sSql & ",FURIMAICB = '" & CStr(CInt(.FURIMAICB)) & "'" & vbLf
        End If
        
        ''�����H��
        If .XTWORKCB <> "" And left(.XTWORKCB, 1) <> vbNullChar Then
            sSql = sSql & ",XTWORKCB = '" & .XTWORKCB & "'" & vbLf
        End If
        
        ''�E�F�[�n����
        If .WFWORKCB <> "" And left(.WFWORKCB, 1) <> vbNullChar Then
            sSql = sSql & ",WFWORKCB = '" & .WFWORKCB & "'" & vbLf
        End If
        
        ''�s�Ǘ��R
        If .FURYCCB <> "" And left(.FURYCCB, 1) <> vbNullChar Then
            sSql = sSql & ",FURYCCB = '" & .FURYCCB & "'" & vbLf
        End If
        
        ''�ŏI��ԋ敪
        If .LSTCCB <> "" And left(.LSTCCB, 1) <> vbNullChar Then
            sSql = sSql & ",LSTCCB = '" & .LSTCCB & "'" & vbLf
        End If
        
        ''�i��R�[�h
        If .LUFRCCB <> "" And left(.LUFRCCB, 1) <> vbNullChar Then
            sSql = sSql & ",LUFRCCB = '" & .LUFRCCB & "'" & vbLf
        End If
        
        ''�i��敪
        If .LUFRBCB <> "" And left(.LUFRBCB, 1) <> vbNullChar Then
            sSql = sSql & ",LUFRBCB = '" & .LUFRBCB & "'" & vbLf
        End If
        
        ''�i���R�[�h
        If .LDERCCB <> "" And left(.LDERCCB, 1) <> vbNullChar Then
            sSql = sSql & ",LDERCCB = '" & .LDERCCB & "'" & vbLf
        End If
        
        ''�i���敪
        If .LDFRBCB <> "" And left(.LDFRBCB, 1) <> vbNullChar Then
            sSql = sSql & ",LDFRBCB = '" & .LDFRBCB & "'" & vbLf
        End If
        
        ''�z�[���h�R�[�h
        If .HOLDCCB <> "" And left(.HOLDCCB, 1) <> vbNullChar Then
            sSql = sSql & ",HOLDCCB = '" & .HOLDCCB & "'" & vbLf
        End If
        
        ''�z�[���h�敪
        If .HOLDBCB <> "" And left(.HOLDBCB, 1) <> vbNullChar Then
            sSql = sSql & ",HOLDBCB = '" & .HOLDBCB & "'" & vbLf
        End If
        
        ''��O�敪
        If .EXKUBCB <> "" And left(.EXKUBCB, 1) <> vbNullChar Then
            sSql = sSql & ",EXKUBCB = '" & .EXKUBCB & "'" & vbLf
        End If
        
        ''�ԕi�敪
        If .HENPKCB <> "" And left(.HENPKCB, 1) <> vbNullChar Then
            sSql = sSql & ",HENPKCB = '" & .HENPKCB & "'" & vbLf
        End If
        
        ''�����敪
        If .LIVKCB <> "" And left(.LIVKCB, 1) <> vbNullChar Then
            sSql = sSql & ",LIVKCB = '" & .LIVKCB & "'" & vbLf
        End If
        
        ''�����敪
        If .KANKCB <> "" And left(.KANKCB, 1) <> vbNullChar Then
            sSql = sSql & ",KANKCB = '" & .KANKCB & "'" & vbLf
        End If
        
        ''���ɋ敪
        If .NFCB <> "" And left(.NFCB, 1) <> vbNullChar Then
            sSql = sSql & ",NFCB = '" & .NFCB & "'" & vbLf
        End If
        
        ''�폜�敪
        If .SAKJCB <> "" And left(.SAKJCB, 1) <> vbNullChar Then
            sSql = sSql & ",SAKJCB = '" & .SAKJCB & "'" & vbLf
        End If
        
        ''�o�^���t
        If .TDAYCB <> "" Then
            sSql = sSql & ",TDAYCB = TO_DATE('" & Format$(CDate(.TDAYCB), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If

        ''SUMIT���M�t���O
        If .SUMITCB <> "" And left(.SUMITCB, 1) <> vbNullChar Then
            sSql = sSql & ",SUMITCB = '" & .SUMITCB & "'" & vbLf
        End If
        
        ''���M�t���O
        If .SNDKCB <> "" And left(.SNDKCB, 1) <> vbNullChar Then
            sSql = sSql & ",SNDKCB = '" & .SNDKCB & "'" & vbLf
        End If
        
        ''���M���t
        If .SNDAYCB <> "" Then
            sSql = sSql & ",SNDAYCB = TO_DATE('" & Format$(CDate(.SNDAYCB), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''�ŏI�ʉߍH��
        If .NEWKNTCB <> "" And left(.NEWKNTCB, 1) <> vbNullChar Then
            sSql = sSql & ",NEWKNTCB = '" & .NEWKNTCB & "'" & vbLf
        End If
        
        ''���ݍH��
        If .GNWKNTCB <> "" And left(.GNWKNTCB, 1) <> vbNullChar Then
            sSql = sSql & ",GNWKNTCB = '" & .GNWKNTCB & "'" & vbLf
        End If
        
        ''�U�֕i��(���j
        If .MOTHINCB <> "" And left(.MOTHINCB, 1) <> vbNullChar Then
            sSql = sSql & ",MOTHINCB = '" & .MOTHINCB & "'" & vbLf
        End If

        ''���_����
        If .RLENCB <> "" And left(.RLENCB, 1) <> vbNullChar Then
            sSql = sSql & ",RLENCB = '" & .RLENCB & "'" & vbLf
        End If
        
        ''�z�[���h�敪(SXL�m��)
        If .SHOLDCLSCB <> "" And left(.SHOLDCLSCB, 1) <> vbNullChar Then
            sSql = sSql & ",SHOLDCLSCB = '" & .SHOLDCLSCB & "'" & vbLf
        End If

        ''����
        If .PLANTCATCB <> "" And left(.PLANTCATCB, 2) <> vbNullChar Then
            sSql = sSql & ",PLANTCATCB = '" & .PLANTCATCB & "'" & vbLf
        End If
            
        ''�֘A��ۯ��׸�
        If .KBLKFLGCB <> "" And left(.KBLKFLGCB, 1) <> vbNullChar Then
            sSql = sSql & ",KBLKFLGCB = '" & .KBLKFLGCB & "'" & vbLf
        End If
        
        sSql = sSql & " " & sqlWhere & vbLf
    
        'SQL�����s
        intRecCnt = OraDB.ExecuteSQL(sSql)
        
        '�Ԃ�l��1�ȊO�̓G���[
        If intRecCnt < 0 Then
            GoTo proc_err
        ElseIf intRecCnt = 0 Then
            '0���X�V�c�G���[(�����ʂ�)
            UpdateXSDCB = FUNCTION_RETURN_FAILURE
            Exit Function
        ElseIf intRecCnt > 1 Then
            '�������X�V�c�G���[(�����͕���SELECT�����ŏ��̈ꌏ�̂ݍX�V)
            UpdateXSDCB = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
    
    End With

'<<<<< Edit-->UPDATE�ɕύX�@2009/07/21�@SSS.Marushita
    
    UpdateXSDCB = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' �I��
'    gErr.Pop
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    UpdateXSDCB = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'��INSERT��  NULL�̏ꍇ�Achar�Ȃ�X�y�[�X�ANumber�Ȃ�NULL������
'*******************************************************************************************
'*    �֐���        : CreateXSDCB
'*
'*    �����T�v      : 1.�e�[�u���uXSDCB�v�Ƀ��R�[�h��}������
'*                      (NULL�̏ꍇ�Achar�Ȃ�X�y�[�X�ANumber�Ȃ�NULL������)
'*
'*    �p�����[�^    : �ϐ���      ,IO  ,�^                ,����
'*      �@         �@udtXSDCB �@�@  ,I   ,typ_XSDCB_Update  ,XSDCB�X�V�p�ް�
'*               �@�@sErrMsg�@�@�@,O   ,String         �@ ,�G���[���b�Z�[�W
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Public Function CreateXSDCB(udtXSDCB As typ_XSDCB_Update, sErrMsg As String) As FUNCTION_RETURN
    Dim sSql        As String
    Dim sDBName     As String
    Dim rs          As OraDynaset   ' RecordSet
    Dim lngRecCnt   As Long         ' ���R�[�h��
    Dim dtmNowtime  As Date

    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_XSDCB_SQL.bas -- Function CreateXSDCB"
    sErrMsg = ""
    sDBName = "XSDCB"
    dtmNowtime = getSvrTime()       ' �T�[�o�[�̎��Ԃ��擾����悤�ɕύX

'>>>>> AddNew-->INSERT�ɕύX�@2009/07/21�@SSS.Marushita
    
    With udtXSDCB
        sSql = "INSERT INTO XSDCB ("
        sSql = sSql & " SXLIDCB" & vbLf      ' 1:SXLID
        sSql = sSql & ",KCNTCB" & vbLf       ' 2:�H���A��
        sSql = sSql & ",XTALCB" & vbLf       ' 3:�����ԍ�
        sSql = sSql & ",INPOSCB" & vbLf      ' 4:�������J�n�ʒu
        sSql = sSql & ",LENCB" & vbLf        ' 5:����
        sSql = sSql & ",HINBCB" & vbLf       ' 6:�i��
        sSql = sSql & ",REVNUMCB" & vbLf     ' 7:���i�ԍ������ԍ�
        sSql = sSql & ",FACTORYCB" & vbLf    ' 8:�H��
        sSql = sSql & ",OPECB" & vbLf        ' 9:���Ə���
        sSql = sSql & ",MAICB" & vbLf        '10:������
        sSql = sSql & ",WSRMAICB" & vbLf     '11:WS��㖇��
        sSql = sSql & ",WSNMAICB" & vbLf     '12:WS��㌇������
        sSql = sSql & ",WFCMAICB" & vbLf     '13:WFC�������
        sSql = sSql & ",SXLRMAICB" & vbLf    '14:SXL�w���i�Ǖi�j
        sSql = sSql & ",SXLNMAICB" & vbLf    '15:SXL�w���i�s�ǁj
        sSql = sSql & ",WFCNMAICB" & vbLf    '16:WFC����������
        sSql = sSql & ",SXLEMAICB" & vbLf    '17:SXL�m�薇��
        sSql = sSql & ",SRMAICB" & vbLf      '18:�T���v�����w���i�Ǖi�j
        sSql = sSql & ",SNMAICB" & vbLf      '19:�T���v�����w���i�s�ǁj
        sSql = sSql & ",STMAICB" & vbLf      '20:�T���v������
        sSql = sSql & ",FURIMAICB" & vbLf    '21:�U�֖���
        sSql = sSql & ",XTWORKCB" & vbLf     '22:�����H��
        sSql = sSql & ",WFWORKCB" & vbLf     '23:�E�F�[�n����
        sSql = sSql & ",FURYCCB" & vbLf      '24:�s�Ǘ��R
        sSql = sSql & ",LSTCCB" & vbLf       '25:�ŏI��ԋ敪
        sSql = sSql & ",LUFRCCB" & vbLf      '26:�i��R�[�h
        sSql = sSql & ",LUFRBCB" & vbLf      '27:�i��敪
        sSql = sSql & ",LDERCCB" & vbLf      '28:�i���R�[�h
        sSql = sSql & ",LDFRBCB" & vbLf      '29:�i���敪
        sSql = sSql & ",HOLDCCB" & vbLf      '30:�z�[���h�R�[�h
        sSql = sSql & ",HOLDBCB" & vbLf      '31:�z�[���h�敪
        sSql = sSql & ",EXKUBCB" & vbLf      '32:��O�敪
        sSql = sSql & ",HENPKCB" & vbLf      '33:�ԕi�敪
        sSql = sSql & ",LIVKCB" & vbLf       '34:�����敪
        sSql = sSql & ",KANKCB" & vbLf       '35:�����敪
        sSql = sSql & ",NFCB" & vbLf         '36:���ɋ敪
        sSql = sSql & ",SAKJCB" & vbLf       '37:�폜�敪
        sSql = sSql & ",TDAYCB" & vbLf       '38:�o�^���t
        sSql = sSql & ",KDAYCB" & vbLf       '39:�X�V���t
        sSql = sSql & ",SUMITCB" & vbLf      '40:SUMIT���M�t���O
        sSql = sSql & ",SNDKCB" & vbLf       '41:���M�t���O
        sSql = sSql & ",SNDAYCB" & vbLf      '42:���M���t
        sSql = sSql & ",NEWKNTCB" & vbLf     '43:�ŏI�ʉߍH��
        sSql = sSql & ",GNWKNTCB" & vbLf     '44:���ݍH��
        sSql = sSql & ",MOTHINCB" & vbLf     '45:�U�֕i��(���j
        sSql = sSql & ",SHOLDCLSCB" & vbLf   '46:�z�[���h�敪
        sSql = sSql & ",RLENCB" & vbLf       '47:���_����
        sSql = sSql & ",KBLKFLGCB" & vbLf    '48:�֘A��ۯ��׸�
        sSql = sSql & ")"
        sSql = sSql & "VALUES ("

        ' 1:SXLID
        If .SXLIDCB <> "" Then
            sSql = sSql & " '" & .SXLIDCB & "'" & vbLf
        Else
            sSql = sSql & " '" & Space(13) & "'" & vbLf
        End If
               
        ' 2:�H���A��
        If .KCNTCB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.KCNTCB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        ' 3:�����ԍ�
        If .XTALCB <> "" Then
            sSql = sSql & ",'" & .XTALCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(12) & "'" & vbLf
        End If
        
        ' 4:�������J�n�ʒu
        If .INPOSCB <> "" Then
            sSql = sSql & ",'" & .INPOSCB & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If
        
        ' 5:����
        If .LENCB <> "" Then
            sSql = sSql & ",'" & .LENCB & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        ' 6:�i��
        If .HINBCB <> "" Then
            sSql = sSql & ",'" & .HINBCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(8) & "'" & vbLf
        End If
        
        ' 7:���i�ԍ������ԍ�
        If .REVNUMCB <> "" Then
            sSql = sSql & ",'" & .REVNUMCB & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        ' 8:�H��
        If .FACTORYCB <> "" Then
            sSql = sSql & ",'" & .FACTORYCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(1) & "'" & vbLf
        End If

        ' 9:���Ə���
        If .OPECB <> "" Then
            sSql = sSql & ",'" & .OPECB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(1) & "'" & vbLf
        End If

        '10:������
        If .MAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.MAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '11:WS��㖇��
        If .WSRMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.WSRMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '12:WS��㌇������
        If .WSNMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.WSNMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '13:WFC�������
        If .WFCMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.WFCMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '14:SXL�w���i�Ǖi�j
        If .SXLRMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.SXLRMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '15:SXL�w���i�s�ǁj
        If .SXLNMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.SXLNMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '16:WFC����������
        If .WFCNMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.WFCNMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '17:SXL�m�薇��
        If .SXLEMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.SXLEMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '18:�T���v�����w���i�Ǖi�j
        If .SRMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.SRMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '19:�T���v�����w���i�s�ǁj
        If .SNMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.SNMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '20:�T���v������
        If .STMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.STMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If

        '21:�U�֖���
        If .FURIMAICB <> "" Then
            sSql = sSql & ",'" & CStr(CInt(.FURIMAICB)) & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf  '���g�p(NULL)
        End If

        '22:�����H��
        sSql = sSql & ",'" & FACTORYCD & "'" & vbLf     '42 �Œ�

        '23:�E�F�[�n����
        If .WFWORKCB <> "" Then
            sSql = sSql & ",'" & .WFWORKCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(2) & "'" & vbLf '���g�p(��߰�)
        End If

        '24:�s�Ǘ��R
        If .FURYCCB <> "" Then
            sSql = sSql & ",'" & .FURYCCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(3) & "'" & vbLf '���g�p(��߰�)
        End If

        '25:�ŏI��ԋ敪
        If .LSTCCB <> "" Then
            sSql = sSql & ",'" & .LSTCCB & "'" & vbLf
        Else
            sSql = sSql & ",'T'" & vbLf       '�ʏ�
        End If

        '26:�i��R�[�h
        If .LUFRCCB <> "" Then
            sSql = sSql & ",'" & .LUFRCCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(3) & "'" & vbLf '���g�p(��߰�)
        End If

        '27:�i��敪
        If .LUFRBCB <> "" Then
            sSql = sSql & ",'" & .LUFRBCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(1) & "'" & vbLf '���g�p(��߰�)
        End If

        '28:�i���R�[�h
        If .LDERCCB <> "" Then
            sSql = sSql & ",'" & .LDERCCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(3) & "'" & vbLf '���g�p(��߰�)
        End If

        '29:�i���敪
        If .LDFRBCB <> "" Then
            sSql = sSql & ",'" & .LDFRBCB & "'" & vbLf
        Else
            sSql = sSql & ",'0'" & vbLf        '�ʏ�
        End If

        '30:�z�[���h�R�[�h
        If .HOLDCCB <> "" Then
            sSql = sSql & ",'" & .HOLDCCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(3) & "'" & vbLf '���g�p(��߰�)
        End If

        '31:�z�[���h�敪
        If .HOLDBCB <> "" Then
            sSql = sSql & ",'" & .HOLDBCB & "'" & vbLf
        Else
            sSql = sSql & ",'0'" & vbLf        '�ʏ�
        End If

        '32:��O�敪
        If .EXKUBCB <> "" Then
            sSql = sSql & ",'" & .EXKUBCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(1) & "'" & vbLf
        End If

        '33:�ԕi�敪
        If .HENPKCB <> "" Then
            sSql = sSql & ",'" & .HENPKCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(1) & "'" & vbLf '���g�p(��߰�)
        End If

        '34:�����敪
        If .LIVKCB <> "" Then
            sSql = sSql & ",'" & .LIVKCB & "'" & vbLf
        Else
            sSql = sSql & ",'0'" & vbLf        '�����b�g
        End If

        '35:�����敪
        If .KANKCB <> "" Then
            sSql = sSql & ",'" & .KANKCB & "'" & vbLf
        Else
            sSql = sSql & ",'0'" & vbLf        '�ʏ�
        End If

        '36:���ɋ敪
        If .NFCB <> "" Then
            sSql = sSql & ",'" & .NFCB & "'" & vbLf
        Else
            sSql = sSql & ",'0'" & vbLf        '�Œ�
        End If

        '37:�폜�敪
        If .SAKJCB <> "" Then
            sSql = sSql & ",'" & .SAKJCB & "'" & vbLf
        Else
            sSql = sSql & ",'0'" & vbLf         '�Œ�
        End If

        '38:�o�^���t
        sSql = sSql & ",TO_DATE('" & Format$(dtmNowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf

        '39:�X�V���t
        sSql = sSql & ",TO_DATE('" & Format$(dtmNowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf

        '40:SUMIT���M�t���O
        If .SUMITCB <> "" Then
            sSql = sSql & ",'" & .SUMITCB & "'" & vbLf
        Else
            sSql = sSql & ",'0'" & vbLf    '�����l(����g�p���Ȃ�)
        End If

        '41:���M�t���O
        If .SNDKCB <> "" Then
            sSql = sSql & ",'" & .SNDKCB & "'" & vbLf
        Else
            sSql = sSql & ",'0'" & vbLf    '�����l(����g�p���Ȃ�)
        End If

        '42:���M���t
        sSql = sSql & ",NULL" & vbLf

        '43:�ŏI�ʉߍH��
        If .NEWKNTCB <> "" Then
            sSql = sSql & ",'" & .NEWKNTCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(5) & "'" & vbLf
        End If
        
        '44:���ݍH��
        If .GNWKNTCB <> "" Then
            sSql = sSql & ",'" & .GNWKNTCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(5) & "'" & vbLf
        End If
        
        '45:�U�֕i��(���j
        If .MOTHINCB <> "" Then
            sSql = sSql & ",'" & .MOTHINCB & "'" & vbLf
        Else
            sSql = sSql & ",'" & Space(8) & "'" & vbLf
        End If

        '46:�z�[���h�敪
        If .SHOLDCLSCB <> "" Then
            sSql = sSql & ",'" & .SHOLDCLSCB & "'" & vbLf
        Else
            sSql = sSql & ",'0'" & vbLf
        End If
        
        '47:���_����
        If .RLENCB <> "" Then
            sSql = sSql & ",'" & .RLENCB & "'" & vbLf
        Else
            sSql = sSql & ",0" & vbLf
        End If
        
        '48:�֘A�u���b�N�t���O
        If .KBLKFLGCB <> "" And left(.KBLKFLGCB, 1) <> vbNullChar Then
            sSql = sSql & ",'" & .KBLKFLGCB & "'" & vbLf
        Else
            sSql = sSql & ",NULL" & vbLf
        End If

        sSql = sSql & ")" & vbLf
    
        'SQL�����s
        If OraDB.ExecuteSQL(sSql) < 1 Then
            GoTo proc_err
        End If

    End With

'<<<<< AddNew-->INSERT�ɕύX�@2009/07/21�@SSS.Marushita

    Debug.Print sSql

    CreateXSDCB = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' �I��
'    gErr.Pop
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    sErrMsg = GetMsgStr("ENG11", "DB", sDBName)
    CreateXSDCB = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************************
'*    �֐���        : GetKCNTCB
'*
'*    �����T�v      : 1.�H���A�Ԃ��擾����
'*
'*    �p�����[�^    : �ϐ���      ,IO  ,�^               ,����
'*      �@         �@sP_SXLID      ,I   ,String           ,SXLID
'*
'*    �߂�l        : �H���A��
'*
'*******************************************************************************************
Public Function GetKCNTCB(sP_SXLID As String) As Integer
    Dim sSql    As String
    Dim rs      As OraDynaset

    sSql = "SELECT KCNTCB FROM XSDCB WHERE SXLIDCB = '" & sP_SXLID & "'"

    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    If IsNull(rs.Fields("KCNTCB")) Then
        GetKCNTCB = 1
    Else
        GetKCNTCB = CInt(rs.Fields("KCNTCB")) + 1
    End If
End Function

'*******************************************************************************************
'*    �֐���        : CheckSXLrecord
'*
'*    �����T�v      : 1.�Y������ں��ޗL��������(����Β������擾����)
'*
'*    �p�����[�^    : �ϐ���      ,IO  ,�^               ,����
'*      �@         �@sP_SXLID      ,I   ,String           ,SXLID
'*�@�@�@�@�@�@�@�@�@ intP_Length     ,O   ,Integer          ,����
'*
'*    �߂�l        : ���R�[�h��
'*
'*******************************************************************************************
Public Function CheckSXLrecord(sP_SXLID As String, intP_Length As Integer) As Integer
    Dim sSql    As String
    Dim rs      As OraDynaset

    sSql = "SELECT LENCB FROM XSDCB WHERE SXLIDCB = '" & sP_SXLID & "'"

    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    If IsNull(rs.Fields("LENCB")) Then
        CheckSXLrecord = 0
    Else
        CheckSXLrecord = 1
        intP_Length = CInt(rs.Fields("LENCB"))
    End If
End Function

'*******************************************************************************************
'*    �֐���        : DBDRV_CheckCodeXSDCB
'*
'*    �����T�v      : 1.�e�[�u���uXSDCB�v�̍H�����`�F�b�N����(CW740/CW750/CW760/CW800)
'*                      (�w��f�[�^�̍H���������œn���ꂽ�H���Ɠ������`�F�b�N����)
'*
'*    �p�����[�^    : �ϐ���      ,IO  ,�^               ,����
'*                   sChkSXLID()  ,I   ,String           ,�`�F�b�N���R�[�h
'*                   sNowCode     ,I   ,String           ,�`�F�b�N�H��
'*                   sErrMsg      ,O   ,String           ,�G���[���b�Z�[�W
'*
'*    �߂�l        : ���R�[�h��
'*
'*******************************************************************************************
Public Function DBDRV_CheckCodeXSDCB(sChkSXLID() As String, sNowCode As String, sErrMsg As String) As FUNCTION_RETURN
    Dim sSql            As String             ' SQL�S��
    Dim rs              As OraDynaset         ' RecordSet
    Dim udtReadXSDCB()  As typ_XSDCB_Update   ' �擾�f�[�^
    Dim lngLoopCnt      As Long
    Dim i               As Long
    Dim j               As Long
    Dim sDBName         As String

    ' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
'    gErr.Push "s_XSDCB_SQL.bas -- Function DBDRV_CheckCodeXSDCB"
    sErrMsg = ""
    sDBName = "XSDCB"

    i = 0
    ReDim udtReadXSDCB(0)
    For lngLoopCnt = 1 To UBound(sChkSXLID)
        If lngLoopCnt = 1 Or _
           (lngLoopCnt > 1 And sChkSXLID(lngLoopCnt) <> sChkSXLID(lngLoopCnt - 1)) Then

            ' SQL��g�ݗ��Ă�
            sSql = ""
            sSql = sSql & " SELECT"
            sSql = sSql & "   SXLIDCB"
            sSql = sSql & "  ,GNWKNTCB"
            sSql = sSql & " FROM"
            sSql = sSql & "   XSDCB"
            sSql = sSql & " WHERE SXLIDCB = '" & sChkSXLID(lngLoopCnt) & "'"
            sSql = sSql & "   AND LIVKCB = '0' "

            ' �f�[�^�𒊏o����
            Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
            If rs Is Nothing Then
                DBDRV_CheckCodeXSDCB = FUNCTION_RETURN_FAILURE
                Exit Function
            End If

            i = i + 1
            ReDim Preserve udtReadXSDCB(i)

            ' ���o���ʂ��i�[����
            If IsNull(rs.Fields("SXLIDCB")) = False Then udtReadXSDCB(i).SXLIDCB = rs.Fields("SXLIDCB")
            If IsNull(rs.Fields("GNWKNTCB")) = False Then udtReadXSDCB(i).GNWKNTCB = rs.Fields("GNWKNTCB")
            rs.Close
        End If
    Next lngLoopCnt

    For j = 1 To i
        ' ���ݍH�����������`�F�b�N����(CST02�͏���)
        If Trim(udtReadXSDCB(j).GNWKNTCB) <> Trim(sNowCode) And _
           Trim(udtReadXSDCB(j).GNWKNTCB) <> "CST02" Then
            ' ���ݍH�����Ⴄ = SXL�����łɓ����Ă���ꍇ�A�G���[�I��
            sErrMsg = GetMsgStr("EBLK6")
            DBDRV_CheckCodeXSDCB = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
    Next j

    DBDRV_CheckCodeXSDCB = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' �I��
'    gErr.Pop
    Exit Function

proc_err:
    ' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    sErrMsg = GetMsgStr("ENG11", "DB", sDBName)
    DBDRV_CheckCodeXSDCB = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

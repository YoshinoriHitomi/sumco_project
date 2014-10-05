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
    'add start 2003/03/25 hitec)matsumoto ----
    NEWKNTCB As String           ' �ŏI�ʉߍH��
    GNWKNTCB As String           ' ���ݍH��
    MOTHINCB As String           ' ���i��
    'add end 2003/03/25 hitec)matsumoto ----
    PLANTCATCB As String         ' ���� 2007/08/30 SPK Tsutsumi Add
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
    'add start 2003/03/25 hitec)matsumoto ----
    NEWKNTCB As String           ' �ŏI�ʉߍH��
    GNWKNTCB As String           ' ���ݍH��
    MOTHINCB As String           ' ���i��
    'add end 2003/03/25 hitec)matsumoto ----
    PLANTCATCB As String         ' ���� 2007/08/30 SPK Tsutsumi Add
End Type

'��SELECT��

'�T�v      :�e�[�u���uXSDCB�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO   ,�^               ,����
'          :records()     ,O    ,typ_XSDCB     ,���o���R�[�h
'          :sqlWhere      ,I    ,String           ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I    ,String           ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O    ,FUNCTION_RETURN   ,���o�̐���
'����      :

Public Function DBDRV_GetXSDCB(records() As typ_XSDCB, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
    
    Dim sql As String       'SQL�S��
    Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      '���R�[�h��
    Dim i As Long


    ''SQL��g�ݗ��Ă�
    sqlBase = "Select * From XSDCB"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetXSDCB = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    If recCnt = 0 Then
        Exit Function
    End If
    For i = 1 To recCnt
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
            'add start 2003/03/25 hitec)matsumoto ------
            If IsNull(rs.Fields("NEWKNTCB")) = False Then .NEWKNTCB = rs.Fields("NEWKNTCB")
            If IsNull(rs.Fields("GNWKNTCB")) = False Then .GNWKNTCB = rs.Fields("GNWKNTCB")
            If IsNull(rs.Fields("MOTHINCB")) = False Then .MOTHINCB = rs.Fields("MOTHINCB")
            'add edn   2003/03/25 hitec)matsumoto ------
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetXSDCB = FUNCTION_RETURN_SUCCESS
End Function

'��UPDATE��

'���X�V���ڂ��\���̂ɃZ�b�g���Ĉ����n��

'�T�v      :�e�[�u���uXSDCB�v���X�V���� ptrn1
'���Ұ�    :�ϐ���        ,IO  ,�^               ,����
'          :records()     ,O   ,typ_XSDCB     ,�X�V���R�[�h
'          :sqlWhere      ,I   ,String           ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I   ,String           ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O   ,FUNCTION_RETURN  ,���o�̐���
'����      :

Public Function UpdateXSDCB(records As typ_XSDCB_Update, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
On Error GoTo proc_err
    gErr.Push "s_XSDCB_SQL.bas -- Function UpdateXSDCB"

    Dim sql As String       'SQL�S��
'    Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
'    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      '���R�[�h��
'    Dim i As Long
    Dim nowtime As Date
    Dim nowtime_sql As String   '�T�[�o����(SQL��)
    
    nowtime = getSvrTime()    '�T�[�o�[�̎��Ԃ��擾����悤�ɕύX 2003/6/4 tuku

'>>>>> .Edit��SQL(UPDATE)���ɕύX�@2009/06/18 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"

    With records
        
        ''SQL��g�ݗ��Ă�
        sql = "UPDATE XSDCB SET" & vbLf
        
        ''�X�V���t
        sql = sql & " KDAYCB = " & nowtime_sql & vbLf
        
        ''SXLID
        If .SXLIDCB <> "" And Left(.SXLIDCB, 1) <> vbNullChar Then
            sql = sql & ",SXLIDCB = '" & .SXLIDCB & "'" & vbLf
        End If
        
        ''�H���A��
        If .KCNTCB <> "" Then
            sql = sql & ",KCNTCB = '" & CStr(CInt(.KCNTCB)) & "'" & vbLf
        End If
        
        ''�����ԍ�
        If .XTALCB <> "" And Left(.XTALCB, 1) <> vbNullChar Then
            sql = sql & ",XTALCB = '" & .XTALCB & "'" & vbLf
        End If
        
        ''�������J�n�ʒu
        If .INPOSCB <> "" Then
            sql = sql & ",INPOSCB = '" & CStr(CInt(.INPOSCB)) & "'" & vbLf
        End If
        
        ''����
        If .LENCB <> "" Then
            sql = sql & ",LENCB = '" & CStr(CInt(.LENCB)) & "'" & vbLf
        End If
        
        ''�i��
        If .HINBCB <> "" And Left(.HINBCB, 1) <> vbNullChar Then
            sql = sql & ",HINBCB = '" & .HINBCB & "'" & vbLf
        End If
        
        ''���i�ԍ������ԍ�
        If .REVNUMCB <> "" Then
            sql = sql & ",REVNUMCB = '" & CStr(CInt(.REVNUMCB)) & "'" & vbLf
        End If
        
        ''�H��
        If .FACTORYCB <> "" And Left(.FACTORYCB, 1) <> vbNullChar Then
            sql = sql & ",FACTORYCB = '" & .FACTORYCB & "'" & vbLf
        End If
        
        ''���Ə���
        If .OPECB <> "" And Left(.OPECB, 1) <> vbNullChar Then
            sql = sql & ",OPECB = '" & .OPECB & "'" & vbLf
        End If
        
        ''������
        If .MAICB <> "" Then
            sql = sql & ",MAICB = '" & CStr(CInt(.MAICB)) & "'" & vbLf
        End If
        
        ''WS��㖇��
        If .WSRMAICB <> "" Then
            sql = sql & ",WSRMAICB = '" & CStr(CInt(.WSRMAICB)) & "'" & vbLf
        End If
        
        ''WS��㌇������
        If .WSNMAICB <> "" Then
            sql = sql & ",WSNMAICB = '" & CStr(CInt(.WSNMAICB)) & "'" & vbLf
        End If
        
        ''WFC�������
        If .WFCMAICB <> "" Then
            sql = sql & ",WFCMAICB = '" & CStr(CInt(.WFCMAICB)) & "'" & vbLf
        End If
        
        ''SXL�w���i�Ǖi�j
        If .SXLRMAICB <> "" Then
            sql = sql & ",SXLRMAICB = '" & CStr(CInt(.SXLRMAICB)) & "'" & vbLf
        End If
        
        ''SXL�w���i�s�ǁj
        If .SXLNMAICB <> "" Then
            sql = sql & ",SXLNMAICB = '" & CStr(CInt(.SXLNMAICB)) & "'" & vbLf
        End If
        
        ''WFC����������
        If .WFCNMAICB <> "" Then
            sql = sql & ",WFCNMAICB = '" & CStr(CInt(.WFCNMAICB)) & "'" & vbLf
        End If
        
        ''SXL�m�薇��
        If .SXLEMAICB <> "" Then
            sql = sql & ",SXLEMAICB = '" & CStr(CInt(.SXLEMAICB)) & "'" & vbLf
        End If
        
        ''�T���v�����w���i�Ǖi�j
        If .SRMAICB <> "" Then
            sql = sql & ",SRMAICB = '" & CStr(CInt(.SRMAICB)) & "'" & vbLf
        End If
        
        ''�T���v�����w���i�s�ǁj
        If .SNMAICB <> "" Then
            sql = sql & ",SNMAICB = '" & CStr(CInt(.SNMAICB)) & "'" & vbLf
        End If
        
        ''�T���v������
        If .STMAICB <> "" Then
            sql = sql & ",STMAICB = '" & CStr(CInt(.STMAICB)) & "'" & vbLf
        End If
        
        ''�U�֖���
        If .FURIMAICB <> "" Then
            sql = sql & ",FURIMAICB = '" & CStr(CInt(.FURIMAICB)) & "'" & vbLf
        End If
        
        ''�����H��
        If .XTWORKCB <> "" And Left(.XTWORKCB, 1) <> vbNullChar Then
            sql = sql & ",XTWORKCB = '" & .XTWORKCB & "'" & vbLf
        End If
        
        ''�E�F�[�n����
        If .WFWORKCB <> "" And Left(.WFWORKCB, 1) <> vbNullChar Then
            sql = sql & ",WFWORKCB = '" & .WFWORKCB & "'" & vbLf
        End If
        
        ''�s�Ǘ��R
        If .FURYCCB <> "" And Left(.FURYCCB, 1) <> vbNullChar Then
            sql = sql & ",FURYCCB = '" & .FURYCCB & "'" & vbLf
        End If
        
        ''�ŏI��ԋ敪
        If .LSTCCB <> "" And Left(.LSTCCB, 1) <> vbNullChar Then
            sql = sql & ",LSTCCB = '" & .LSTCCB & "'" & vbLf
        End If
        
        ''�i��R�[�h
        If .LUFRCCB <> "" And Left(.LUFRCCB, 1) <> vbNullChar Then
            sql = sql & ",LUFRCCB = '" & .LUFRCCB & "'" & vbLf
        End If
        
        ''�i��敪
        If .LUFRBCB <> "" And Left(.LUFRBCB, 1) <> vbNullChar Then
            sql = sql & ",LUFRBCB = '" & .LUFRBCB & "'" & vbLf
        End If
        
        ''�i���R�[�h
        If .LDERCCB <> "" And Left(.LDERCCB, 1) <> vbNullChar Then
            sql = sql & ",LDERCCB = '" & .LDERCCB & "'" & vbLf
        End If
        
        ''�i���敪
        If .LDFRBCB <> "" And Left(.LDFRBCB, 1) <> vbNullChar Then
            sql = sql & ",LDFRBCB = '" & .LDFRBCB & "'" & vbLf
        End If
        
        ''�z�[���h�R�[�h
        If .HOLDCCB <> "" And Left(.HOLDCCB, 1) <> vbNullChar Then
            sql = sql & ",HOLDCCB = '" & .HOLDCCB & "'" & vbLf
        End If
        
        ''�z�[���h�敪
        If .HOLDBCB <> "" And Left(.HOLDBCB, 1) <> vbNullChar Then
            sql = sql & ",HOLDBCB = '" & .HOLDBCB & "'" & vbLf
        End If
        
        ''��O�敪
        If .EXKUBCB <> "" And Left(.EXKUBCB, 1) <> vbNullChar Then
            sql = sql & ",EXKUBCB = '" & .EXKUBCB & "'" & vbLf
        End If
        
        ''�ԕi�敪
        If .HENPKCB <> "" And Left(.HENPKCB, 1) <> vbNullChar Then
            sql = sql & ",HENPKCB = '" & .HENPKCB & "'" & vbLf
        End If
        
        ''�����敪
        If .LIVKCB <> "" And Left(.LIVKCB, 1) <> vbNullChar Then
            sql = sql & ",LIVKCB = '" & .LIVKCB & "'" & vbLf
        End If
        
        ''�����敪
        If .KANKCB <> "" And Left(.KANKCB, 1) <> vbNullChar Then
            sql = sql & ",KANKCB = '" & .KANKCB & "'" & vbLf
        End If
        
        ''���ɋ敪
        If .NFCB <> "" And Left(.NFCB, 1) <> vbNullChar Then
            sql = sql & ",NFCB = '" & .NFCB & "'" & vbLf
        End If
        
        ''�폜�敪
        If .SAKJCB <> "" And Left(.SAKJCB, 1) <> vbNullChar Then
            sql = sql & ",SAKJCB = '" & .SAKJCB & "'" & vbLf
        End If
        
        ''�o�^���t
        If .TDAYCB <> "" Then
            sql = sql & ",TDAYCB = TO_DATE('" & Format$(CDate(.TDAYCB), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''SUMIT���M�t���O
        If .SUMITCB <> "" And Left(.SUMITCB, 1) <> vbNullChar Then
            sql = sql & ",SUMITCB = '" & .SUMITCB & "'" & vbLf
        End If
        
        ''���M�t���O
        If .SNDKCB <> "" And Left(.SNDKCB, 1) <> vbNullChar Then
            sql = sql & ",SNDKCB = '" & .SNDKCB & "'" & vbLf
        End If
        
        ''���M���t
        If .SNDAYCB <> "" Then
            sql = sql & ",SNDAYCB = TO_DATE('" & Format$(CDate(.SNDAYCB), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''�ŏI�ʉߍH��
        If .NEWKNTCB <> "" And Left(.NEWKNTCB, 1) <> vbNullChar Then
            sql = sql & ",NEWKNTCB = '" & .NEWKNTCB & "'" & vbLf
        End If
        
        ''���ݍH��
        If .GNWKNTCB <> "" And Left(.GNWKNTCB, 1) <> vbNullChar Then
            sql = sql & ",GNWKNTCB = '" & .GNWKNTCB & "'" & vbLf
        End If
        
        ''�U�֕i��(���j
        If .MOTHINCB <> "" And Left(.MOTHINCB, 1) <> vbNullChar Then
            sql = sql & ",MOTHINCB = '" & .MOTHINCB & "'" & vbLf
        End If
        
        ''���Ə��敪
        If .PLANTCATCB <> "" And Left(.PLANTCATCB, 2) <> vbNullChar Then
            sql = sql & ",PLANTCATCB = '" & .PLANTCATCB & "'" & vbLf
        End If

        sql = sql & " " & sqlWhere & vbLf
    
        'SQL�����s
        recCnt = OraDB.ExecuteSQL(sql)
        
        '�Ԃ�l��1�ȊO�̓G���[
        If recCnt < 0 Then
            GoTo proc_err
        ElseIf recCnt = 0 Then
            '0���X�V�c�G���[(�����ʂ�)
            UpdateXSDCB = FUNCTION_RETURN_FAILURE
            Exit Function
        ElseIf recCnt > 1 Then
            '�������X�V�c�G���[(�����͕���SELECT�����ŏ��̈ꌏ�̂ݍX�V)
            UpdateXSDCB = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
    
    End With
'<<<<< .Edit��SQL(UPDATE)���ɕύX�@2009/06/18 SETsw kubota ------------------

    UpdateXSDCB = FUNCTION_RETURN_SUCCESS


proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    UpdateXSDCB = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'��INSERT��  NULL�̏ꍇ�Achar�Ȃ�X�y�[�X�ANumber�Ȃ�NULL������

'�T�v      :�e�[�u���uXSDCB�v�Ƀ��R�[�h��}������
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:pXSDCB �@�@  ,I  ,typ_XSDCB_Update   ,XSDCB�X�V�p�ް�
'      �@�@:sErrMsg�@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�������݂̐���
Public Function CreateXSDCB(pXSDCB As typ_XSDCB_Update, sErrMsg As String) As FUNCTION_RETURN


    Dim sql As String
    Dim sDbName As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim recCnt As Long      '���R�[�h��
    Dim nowtime As Date
    Dim nowtime_sql As String   '�T�[�o����(SQL��)
    
    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_XSDCB_SQL.bas -- Function CreateXSDCB"
    sErrMsg = ""
    sDbName = "XSDCB"
    'nowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")
    nowtime = getSvrTime()    '�T�[�o�[�̎��Ԃ��擾����悤�ɕύX 2003/6/4 tuku

'>>>>> .AddNew��SQL(INSERT)���ɕύX�@2009/06/18 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
    
    With pXSDCB
        sql = "INSERT INTO XSDCB ("
        sql = sql & " SXLIDCB"      ' 1:SXLID
        sql = sql & ",KCNTCB"       ' 2:�H���A��
        sql = sql & ",XTALCB"       ' 3:�����ԍ�
        sql = sql & ",INPOSCB"      ' 4:�������J�n�ʒu
        sql = sql & ",LENCB"        ' 5:����
        sql = sql & ",HINBCB"       ' 6:�i��
        sql = sql & ",REVNUMCB"     ' 7:���i�ԍ������ԍ�
        sql = sql & ",FACTORYCB"    ' 8:�H��
        sql = sql & ",OPECB"        ' 9:���Ə���
        sql = sql & ",MAICB"        '10:������
        sql = sql & ",WSRMAICB"     '11:WS��㖇��
        sql = sql & ",WSNMAICB"     '12:WS��㌇������
        sql = sql & ",WFCMAICB"     '13:WFC�������
        sql = sql & ",SXLRMAICB"    '14:SXL�w���i�Ǖi�j
        sql = sql & ",SXLNMAICB"    '15:SXL�w���i�s�ǁj
        sql = sql & ",WFCNMAICB"    '16:WFC����������
        sql = sql & ",SXLEMAICB"    '17:SXL�m�薇��
        sql = sql & ",SRMAICB"      '18:�T���v�����w���i�Ǖi�j
        sql = sql & ",SNMAICB"      '19:�T���v�����w���i�s�ǁj
        sql = sql & ",STMAICB"      '20:�T���v������
        sql = sql & ",FURIMAICB"    '21:�U�֖���
        sql = sql & ",XTWORKCB"     '22:�����H��
        sql = sql & ",WFWORKCB"     '23:�E�F�[�n����
        sql = sql & ",FURYCCB"      '24:�s�Ǘ��R
        sql = sql & ",LSTCCB"       '25:�ŏI��ԋ敪
        sql = sql & ",LUFRCCB"      '26:�i��R�[�h
        sql = sql & ",LUFRBCB"      '27:�i��敪
        sql = sql & ",LDERCCB"      '28:�i���R�[�h
        sql = sql & ",LDFRBCB"      '29:�i���敪
        sql = sql & ",HOLDCCB"      '30:�z�[���h�R�[�h
        sql = sql & ",HOLDBCB"      '31:�z�[���h�敪
        sql = sql & ",EXKUBCB"      '32:��O�敪
        sql = sql & ",HENPKCB"      '33:�ԕi�敪
        sql = sql & ",LIVKCB"       '34:�����敪
        sql = sql & ",KANKCB"       '35:�����敪
        sql = sql & ",NFCB"         '36:���ɋ敪
        sql = sql & ",SAKJCB"       '37:�폜�敪
        sql = sql & ",TDAYCB"       '38:�o�^���t
        sql = sql & ",KDAYCB"       '39:�X�V���t
        sql = sql & ",SUMITCB"      '40:SUMIT���M�t���O
        sql = sql & ",SNDKCB"       '41:���M�t���O
        sql = sql & ",SNDAYCB"      '42:���M���t
        sql = sql & ",NEWKNTCB"     '43:�ŏI�ʉߍH��
        sql = sql & ",GNWKNTCB"     '44:���ݍH��
        sql = sql & ",MOTHINCB"     '45:�U�֕i��(���j
        sql = sql & ",PLANTCATCB"   '46:���Ə��敪
        sql = sql & ")"
        sql = sql & "VALUES ("

        ' 1:SXLID
        If .SXLIDCB <> "" Then
            sql = sql & " '" & .SXLIDCB & "'" & vbLf
        Else
            sql = sql & " '" & Space(13) & "'" & vbLf
        End If

        ' 2:�H���A��
        If .KCNTCB <> "" Then
            sql = sql & ",'" & CStr(CInt(.KCNTCB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 3:�����ԍ�
        If .XTALCB <> "" Then
            sql = sql & ",'" & .XTALCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(12) & "'" & vbLf
        End If
        
        ' 4:�������J�n�ʒu
        If .INPOSCB <> "" Then
            sql = sql & ",'" & .INPOSCB & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 5:����
        If .LENCB <> "" Then
            sql = sql & ",'" & .LENCB & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 6:�i��
        If .HINBCB <> "" Then
            sql = sql & ",'" & .HINBCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(8) & "'" & vbLf
        End If

        ' 7:���i�ԍ������ԍ�
        If .REVNUMCB <> "" Then
            sql = sql & ",'" & .REVNUMCB & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 8:�H��
        If .FACTORYCB <> "" Then
            sql = sql & ",'" & .FACTORYCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        ' 9:���Ə���
        If .OPECB <> "" Then
            sql = sql & ",'" & .OPECB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '10:������
        If .MAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.MAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '11:WS��㖇��
        If .WSRMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.WSRMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '12:WS��㌇������
        If .WSNMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.WSNMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '13:WFC�������
        If .WFCMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.WFCMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '14:SXL�w���i�Ǖi�j
        If .SXLRMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.SXLRMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '15:SXL�w���i�s�ǁj
        If .SXLNMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.SXLNMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '16:WFC����������
        If .WFCNMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.WFCNMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '17:SXL�m�薇��
        If .SXLEMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.SXLEMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '18:�T���v�����w���i�Ǖi�j
        If .SRMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.SRMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '19:�T���v�����w���i�s�ǁj
        If .SNMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.SNMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '20:�T���v������
        If .STMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.STMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '21:�U�֖���
        If .FURIMAICB <> "" Then
            sql = sql & ",'" & CStr(CInt(.FURIMAICB)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '22:�����H��
        sql = sql & ",'" & FACTORYCD & "'" & vbLf

        '23:�E�F�[�n����
        If .WFWORKCB <> "" Then
            sql = sql & ",'" & .WFWORKCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '24:�s�Ǘ��R
        If .FURYCCB <> "" Then
            sql = sql & ",'" & .FURYCCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '25:�ŏI��ԋ敪
        If .LSTCCB <> "" Then
            sql = sql & ",'" & .LSTCCB & "'" & vbLf
        Else
            sql = sql & ",'T'" & vbLf           '�ʏ�
        End If

        '26:�i��R�[�h
        If .LUFRCCB <> "" Then
            sql = sql & ",'" & .LUFRCCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '27:�i��敪
        If .LUFRBCB <> "" Then
            sql = sql & ",'" & .LUFRBCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '28:�i���R�[�h
        If .LDERCCB <> "" Then
            sql = sql & ",'" & .LDERCCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '29:�i���敪
        If .LDFRBCB <> "" Then
            sql = sql & ",'" & .LDFRBCB & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '30:�z�[���h�R�[�h
        If .HOLDCCB <> "" Then
            sql = sql & ",'" & .HOLDCCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '31:�z�[���h�敪
        If .HOLDBCB <> "" Then
            sql = sql & ",'" & .HOLDBCB & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '32:��O�敪
        If .EXKUBCB <> "" Then
            sql = sql & ",'" & .EXKUBCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '33:�ԕi�敪
        If .HENPKCB <> "" Then
            sql = sql & ",'" & .HENPKCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '34:�����敪
        If .LIVKCB <> "" Then
            sql = sql & ",'" & .LIVKCB & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '35:�����敪
        If .KANKCB <> "" Then
            sql = sql & ",'" & .KANKCB & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '36:���ɋ敪
        If .NFCB <> "" Then
            sql = sql & ",'" & .NFCB & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '37:�폜�敪
        If .SAKJCB <> "" Then
            sql = sql & ",'" & .SAKJCB & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '38:�o�^���t
        sql = sql & "," & nowtime_sql & vbLf
        
        '39:�X�V���t
        sql = sql & "," & nowtime_sql & vbLf

        '40:SUMIT���M�t���O
        If .SUMITCB <> "" Then
            sql = sql & ",'" & .SUMITCB & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '41:���M�t���O
        If .SNDKCB <> "" Then
            sql = sql & ",'" & .SNDKCB & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '42:���M���t
        sql = sql & ",NULL" & vbLf

        '43:�ŏI�ʉߍH��
        If .NEWKNTCB <> "" Then
            sql = sql & ",'" & .NEWKNTCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If
        
        '44:���ݍH��
        If .GNWKNTCB <> "" Then
            sql = sql & ",'" & .GNWKNTCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If
        
        '45:�U�֕i��(���j
        If .MOTHINCB <> "" Then
            sql = sql & ",'" & .MOTHINCB & "'" & vbLf
        Else
            sql = sql & ",'" & Space(8) & "'" & vbLf
        End If

        '46:���Ə��敪
        If .PLANTCATCB <> "" Then
            sql = sql & ",'" & .PLANTCATCB & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        sql = sql & ")" & vbLf
    
        'SQL�����s
        If OraDB.ExecuteSQL(sql) < 1 Then
            GoTo proc_err
        End If

    End With
'<<<<< .AddNew��SQL(INSERT)���ɕύX�@2009/06/18 SETsw kubota ------------------

    CreateXSDCB = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", "DB", sDbName)
    CreateXSDCB = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :�H���A�Ԃ��擾����
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:p_SXLID      ,I  ,String           ,SXLID
'      �@�@:�߂�l       ,O  ,Integer        �@,�H���A��
Public Function GetKCNTCB(p_SXLID As String) As Integer
    Dim sql As String
    Dim rs As OraDynaset
    
    sql = "SELECT KCNTCB FROM XSDCB WHERE SXLIDCB = '" & p_SXLID & "'"
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If IsNull(rs.Fields("KCNTCB")) Then
        GetKCNTCB = 1
    Else
        GetKCNTCB = CInt(rs.Fields("KCNTCB")) + 1
    End If
    
End Function


'�T�v      :�Y������ں��ޗL��������(����Β������擾����)
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:p_SXLID      ,I  ,String           ,SXLID
'      �@�@:p_Length     ,O  ,Integer          ,����
'      �@�@:�߂�l       ,O  ,Integer        �@,ں��ސ�
Public Function CheckSXLrecord(p_SXLID As String, p_Length As Integer) As Integer
    Dim sql As String
    Dim rs As OraDynaset
    
    sql = "SELECT LENCB FROM XSDCB WHERE SXLIDCB = '" & p_SXLID & "'"
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If IsNull(rs.Fields("LENCB")) Then
        CheckSXLrecord = 0
    Else
        CheckSXLrecord = 1
        p_Length = CInt(rs.Fields("LENCB"))
    End If
    
End Function



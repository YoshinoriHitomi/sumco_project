Attribute VB_Name = "s_XSDCZ_SQL"
'@'s_XSDCZ_SQL.bas              '( 06/04/14 ) SMP-OKAMOTO �V�K�ǉ�
''�ؒf�w�� (XSDCZ) �����֐�

''***�e�[�u���uXSDCZ�v�ւ̃f�[�^�A�N�Z�X�֐�***
'������ ���Ұ��ɒl��Ă��鎞�A�܂��S�ď��������邱��

Option Explicit

''XSDCZ�p�\����
''�z��̏ꍇ�A�f�[�^�i�[�̓C���f�b�N�X=1����n�߂邱�ƁB
Public Type typ_XSDCZ
    CRYNUMCZ    As String       ''�u���b�NID������ԍ�
    HINBCZ      As String       ''�i��
    INPOSCZ     As String       ''�������J�n�ʒu
    REVNUMCZ    As String       ''���i�ԍ������ԍ�
    FACTORYCZ   As String       ''�H��
    OPECZ       As String       ''���Ə���
    KCKNTCZ     As String       ''�H���A��
    SXLIDCZ     As String       ''SXLID
    XTALCZ      As String       ''�����ԍ�
    NEKKNTCZ    As String       ''�ŏI�ʉߊǗ��H��
    NEWKNTCZ    As String       ''�ŏI�ʉߍH��
    NEWKKBCZ    As String       ''�ŏI�ʉߍ�Ƌ敪
    NEMACOCZ    As String       ''�ŏI�ʉߏ�����
    GNKKNTCZ    As String       ''���݊Ǘ��H��
    GNWKNTCZ    As String       ''���ݍH��
    GNWKKBCZ    As String       ''���ݍ�Ƌ敪
    GNMACOCZ    As String       ''���ݏ�����
    GNDAYCZ     As String       ''���ݏ������t
    GNLCZ       As String       ''���ݒ���
    GNWCZ       As String       ''���ݏd��
    GNMCZ       As String       ''���ݖ���
    SUMITLCZ    As String       ''SUMMIT����
    SUMITWCZ    As String       ''SUMMIT�d��
    SUMITMCZ    As String       ''SUMMIT����
    CHGCZ       As String       ''�`���[�W��
    KAKOUBCZ    As String       ''���H�敪
    KEIDAYCZ    As String       ''�v����t
    GNTKUBCZ    As String       ''�I�敪
    GNTNOCZ     As String       ''�I�ԍ�
    XTWORKCZ    As String       ''�����H��
    WFWORKCZ    As String       ''�E�F�[�n����
    LSTATBCZ    As String       ''�ŏI��ԋ敪
    RSTATBCZ    As String       ''������ԋ敪
    LUFRCCZ     As String       ''�i�㺰��
    LUFRBCZ     As String       ''�i��敪
    LDFRCCZ     As String       ''�i������
    LDFRBCZ     As String       ''�i���敪
    HOLDCCZ     As String       ''ΰ��޺���
    HOLDBCZ     As String       ''�z�[���h�敪
    EXKUBCZ     As String       ''��O�敪
    HENPKCZ     As String       ''�ԕi�敪
    LIVKCZ      As String       ''�����敪
    KANKCZ      As String       ''�����敪
    NFCZ        As String       ''���ɋ敪
    SAKJCZ      As String       ''�폜�敪
    TDAYCZ      As String       ''�o�^���t
    KDAYCZ      As String       ''�X�V���t
    SUMITBCZ    As String       ''SUMMIT���M�t���O
    SNDKCZ      As String       ''���M�t���O
    SNDDAYCZ    As String       ''���M���t
    LBLFLGCZ    As String       ''���x���o�͊m�F�t���O
    CUTCNTCZ    As String * 1   ''�ؒf�����敪      '1:�Đ�
    HINBFLGCZ   As String * 1   ''��\�i�ԃt���O    '1�F��\�i��
    WFHOLDFLGCZ As String       ''�z�[���h�敪(WF)
    HOLDKTCZ    As String * 5   ''�z�[���h�H��
    RPCRYNUMCZ  As String * 12  ''�e�u���b�NID
    FCODECZ     As String       ''�s�ǃR�[�h
    SGNKCZ      As String       ''���������敪
    CUTKCZ      As String       ''�ؒf�敪
    STOPFLG     As String       ''�����ύX�t���O
    PLANTCATCZ  As String       ''����@2007/08/15 SPK Tsutsumi
End Type

'��SELECT��

'�T�v      :�e�[�u���uXSDCZ�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO   ,�^               ,����
'          :records()     ,O    ,typ_XSDCZ        ,���o���R�[�h
'          :lsSqlWhere    ,I    ,String           ,���o����(SQL��Where��:�ȗ��\)
'          :lsSqlOrder    ,I    ,String           ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O    ,FUNCTION_RETURN   ,���o�̐���
'����      :

Public Function DBDRV_GetXSDCZ(records() As typ_XSDCZ, _
                                Optional lsSqlWhere$ = vbNullString, _
                                Optional lsSqlOrder$ = vbNullString) As FUNCTION_RETURN
    
    Dim lsSql       As String           ''SQL�S��
    Dim lsSqlBase   As String           ''SQL��{��(WHERE�߂̑O�܂�)
    Dim rs          As OraDynaset       ''RecordSet
    Dim recCnt      As Long             ''���R�[�h��
    Dim i           As Long             ''�J�E���^

    ''SQL��g�ݗ��Ă�
    lsSqlBase = "Select * From XSDCZ"
    lsSql = lsSqlBase
    If (lsSqlWhere <> vbNullString) Or (lsSqlOrder <> vbNullString) Then
        lsSql = lsSql & " " & lsSqlWhere & " " & lsSqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(lsSql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetXSDCZ = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    If recCnt = 0 Then
        Exit Function
    End If
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            ''�u���b�NID������ԍ�
            If IsNull(rs.Fields("CRYNUMCZ")) = False Then .CRYNUMCZ = CStr(rs.Fields("CRYNUMCZ"))
            ''�i��
            If IsNull(rs.Fields("HINBCZ")) = False Then .HINBCZ = CStr(rs.Fields("HINBCZ"))
            ''�������J�n�ʒu
            If IsNull(rs.Fields("INPOSCZ")) = False Then .INPOSCZ = CStr(rs.Fields("INPOSCZ"))
            ''���i�ԍ������ԍ�
            If IsNull(rs.Fields("REVNUMCZ")) = False Then .REVNUMCZ = CStr(rs.Fields("REVNUMCZ"))
            ''�H��
            If IsNull(rs.Fields("FACTORYCZ")) = False Then .FACTORYCZ = CStr(rs.Fields("FACTORYCZ"))
            ''���Ə���
            If IsNull(rs.Fields("OPECZ")) = False Then .OPECZ = CStr(rs.Fields("OPECZ"))
            ''�H���A��
            If IsNull(rs.Fields("KCKNTCZ")) = False Then .KCKNTCZ = CStr(rs.Fields("KCKNTCZ"))
            ''SXLID
            If IsNull(rs.Fields("SXLIDCZ")) = False Then .SXLIDCZ = CStr(rs.Fields("SXLIDCZ"))
            ''�����ԍ�
            If IsNull(rs.Fields("XTALCZ")) = False Then .XTALCZ = CStr(rs.Fields("XTALCZ"))
            ''�ŏI�ʉߊǗ��H��
            If IsNull(rs.Fields("NEKKNTCZ")) = False Then .NEKKNTCZ = CStr(rs.Fields("NEKKNTCZ"))
            ''�ŏI�ʉߍH��
            If IsNull(rs.Fields("NEWKNTCZ")) = False Then .NEWKNTCZ = CStr(rs.Fields("NEWKNTCZ"))
            ''�ŏI�ʉߍ�Ƌ敪
            If IsNull(rs.Fields("NEWKKBCZ")) = False Then .NEWKKBCZ = CStr(rs.Fields("NEWKKBCZ"))
            ''�ŏI�ʉߏ�����
            If IsNull(rs.Fields("NEMACOCZ")) = False Then .NEMACOCZ = CStr(rs.Fields("NEMACOCZ"))
            ''���݊Ǘ��H��
            If IsNull(rs.Fields("GNKKNTCZ")) = False Then .GNKKNTCZ = CStr(rs.Fields("GNKKNTCZ"))
            ''���ݍH��
            If IsNull(rs.Fields("GNWKNTCZ")) = False Then .GNWKNTCZ = CStr(rs.Fields("GNWKNTCZ"))
            ''���ݍ�Ƌ敪
            If IsNull(rs.Fields("GNWKKBCZ")) = False Then .GNWKKBCZ = CStr(rs.Fields("GNWKKBCZ"))
            ''���ݏ�����
            If IsNull(rs.Fields("GNMACOCZ")) = False Then .GNMACOCZ = CStr(rs.Fields("GNMACOCZ"))
            ''���ݏ������t
            If IsNull(rs.Fields("GNDAYCZ")) = False Then .GNDAYCZ = Format(CStr(rs.Fields("GNDAYCZ")), "yyyy/mm/dd hh:mm")
            ''���ݒ���
            If IsNull(rs.Fields("GNLCZ")) = False Then .GNLCZ = CStr(rs.Fields("GNLCZ"))
            ''���ݏd��
            If IsNull(rs.Fields("GNWCZ")) = False Then .GNWCZ = CStr(rs.Fields("GNWCZ"))
            ''���ݖ���
            If IsNull(rs.Fields("GNMCZ")) = False Then .GNMCZ = CStr(rs.Fields("GNMCZ"))
            ''SUMMIT����
            If IsNull(rs.Fields("SUMITLCZ")) = False Then .SUMITLCZ = CStr(rs.Fields("SUMITLCZ"))
            ''SUMMIT�d��
            If IsNull(rs.Fields("SUMITWCZ")) = False Then .SUMITWCZ = CStr(rs.Fields("SUMITWCZ"))
            ''SUMMIT����
            If IsNull(rs.Fields("SUMITMCZ")) = False Then .SUMITMCZ = CStr(rs.Fields("SUMITMCZ"))
            ''�`���[�W��
            If IsNull(rs.Fields("CHGCZ")) = False Then .CHGCZ = CStr(rs.Fields("CHGCZ"))
            ''���H�敪
            If IsNull(rs.Fields("KAKOUBCZ")) = False Then .KAKOUBCZ = CStr(rs.Fields("KAKOUBCZ"))
            ''�v����t
            If IsNull(rs.Fields("KEIDAYCZ")) = False Then .KEIDAYCZ = CStr(rs.Fields("KEIDAYCZ"))
            ''�I�敪
            If IsNull(rs.Fields("GNTKUBCZ")) = False Then .GNTKUBCZ = CStr(rs.Fields("GNTKUBCZ"))
            ''�I�ԍ�
            If IsNull(rs.Fields("GNTNOCZ")) = False Then .GNTNOCZ = CStr(rs.Fields("GNTNOCZ"))
            ''�����H��
            If IsNull(rs.Fields("XTWORKCZ")) = False Then .XTWORKCZ = CStr(rs.Fields("XTWORKCZ"))
            ''�E�F�[�n����
            If IsNull(rs.Fields("WFWORKCZ")) = False Then .WFWORKCZ = CStr(rs.Fields("WFWORKCZ"))
            ''�ŏI��ԋ敪
            If IsNull(rs.Fields("LSTATBCZ")) = False Then .LSTATBCZ = CStr(rs.Fields("LSTATBCZ"))
            ''������ԋ敪
            If IsNull(rs.Fields("RSTATBCZ")) = False Then .RSTATBCZ = CStr(rs.Fields("RSTATBCZ"))
            ''�i�㺰��
            If IsNull(rs.Fields("LUFRCCZ")) = False Then .LUFRCCZ = CStr(rs.Fields("LUFRCCZ"))
            ''�i��敪
            If IsNull(rs.Fields("LUFRBCZ")) = False Then .LUFRBCZ = CStr(rs.Fields("LUFRBCZ"))
            ''�i�㺰��
            If IsNull(rs.Fields("LDFRCCZ")) = False Then .LDFRCCZ = CStr(rs.Fields("LDFRCCZ"))
            ''�i��敪
            If IsNull(rs.Fields("LDFRBCZ")) = False Then .LDFRBCZ = CStr(rs.Fields("LDFRBCZ"))
            ''ΰ��޺���
            If IsNull(rs.Fields("HOLDCCZ")) = False Then .HOLDCCZ = CStr(rs.Fields("HOLDCCZ"))
            ''�z�[���h�敪
            If IsNull(rs.Fields("HOLDBCZ")) = False Then .HOLDBCZ = CStr(rs.Fields("HOLDBCZ"))
            ''��O�敪
            If IsNull(rs.Fields("EXKUBCZ")) = False Then .EXKUBCZ = CStr(rs.Fields("EXKUBCZ"))
            ''�ԕi�敪
            If IsNull(rs.Fields("HENPKCZ")) = False Then .HENPKCZ = CStr(rs.Fields("HENPKCZ"))
            ''�����敪
            If IsNull(rs.Fields("LIVKCZ")) = False Then .LIVKCZ = CStr(rs.Fields("LIVKCZ"))
            ''�����敪
            If IsNull(rs.Fields("KANKCZ")) = False Then .KANKCZ = CStr(rs.Fields("KANKCZ"))
            ''���ɋ敪
            If IsNull(rs.Fields("NFCZ")) = False Then .NFCZ = CStr(rs.Fields("NFCZ"))
            ''�폜�敪
            If IsNull(rs.Fields("SAKJCZ")) = False Then .SAKJCZ = CStr(rs.Fields("SAKJCZ"))
            ''�o�^���t
            If IsNull(rs.Fields("TDAYCZ")) = False Then .TDAYCZ = Format(CStr(rs.Fields("TDAYCZ")), "yyyy/mm/dd hh:mm")
            ''�X�V���t
            If IsNull(rs.Fields("KDAYCZ")) = False Then .KDAYCZ = Format(CStr(rs.Fields("KDAYCZ")), "yyyy/mm/dd hh:mm")
            ''SUMMIT���M�t���O
            If IsNull(rs.Fields("SUMITBCZ")) = False Then .SUMITBCZ = CStr(rs.Fields("SUMITBCZ"))
            ''���M�t���O
            If IsNull(rs.Fields("SNDKCZ")) = False Then .SNDKCZ = CStr(rs.Fields("SNDKCZ"))
            ''���M���t
            If IsNull(rs.Fields("SNDDAYCZ")) = False Then .SNDDAYCZ = Format(CStr(rs.Fields("SNDDAYCZ")), "yyyy/mm/dd hh:mm")
            ''���x���o�͊m�F�t���O
            If IsNull(rs.Fields("LBLFLGCZ")) = False Then .LBLFLGCZ = CStr(rs.Fields("LBLFLGCZ"))
            ''�ؒf�����敪
            If IsNull(rs.Fields("CUTCNTCZ")) = False Then .CUTCNTCZ = CStr(rs.Fields("CUTCNTCZ"))
            ''��\�i��
            If IsNull(rs.Fields("HINBFLGCZ")) = False Then .HINBFLGCZ = CStr(rs.Fields("HINBFLGCZ"))
            ''�z�[���h�敪(WF)
            If IsNull(rs.Fields("WFHOLDFLGCZ")) = False Then .WFHOLDFLGCZ = CStr(rs.Fields("WFHOLDFLGCZ"))
            ''�z�[���h�H��
            If IsNull(rs.Fields("HOLDKTCZ")) = False Then .HOLDKTCZ = CStr(rs.Fields("HOLDKTCZ"))
            ''�e�u���b�NID
            If IsNull(rs.Fields("RPCRYNUMCZ")) = False Then .RPCRYNUMCZ = CStr(rs.Fields("RPCRYNUMCZ"))
            ''�s�ǃR�[�h
            If IsNull(rs.Fields("FCODECZ")) = False Then .FCODECZ = CStr(rs.Fields("FCODECZ"))
            ''���������敪
            If IsNull(rs.Fields("SGNKCZ")) = False Then .SGNKCZ = CStr(rs.Fields("SGNKCZ"))
            ''�ؒf�敪
            If IsNull(rs.Fields("CUTKCZ")) = False Then .CUTKCZ = CStr(rs.Fields("CUTKCZ"))
            ''���� 2007/08/17 SPK Tsutsumi Add Start
            If IsNull(rs.Fields("PLANTCATCZ")) = False Then .PLANTCATCZ = CStr(rs.Fields("PLANTCATCZ"))
            ''���� 2007/08/17 SPK Tsutsumi Add End
            End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetXSDCZ = FUNCTION_RETURN_SUCCESS
End Function

'��UPDATE��

'���X�V���ڂ��\���̂ɃZ�b�g���Ĉ����n��

'�T�v      :�e�[�u���uXSDCZ�v���X�V���� ptrn1
'���Ұ�    :�ϐ���        ,IO  ,�^               ,����
'          :records()     ,O   ,typ_XSDCZ        ,�X�V���R�[�h
'          :lsSqlWhere    ,I   ,String           ,���o����(SQL��Where��:�ȗ��\)
'          :lsSqlOrder    ,I   ,String           ,���o����(SQL��Order by��:�ȗ��\)
'          :lsUpdate      ,I   ,String           ,�X�V�ӏ��ݒ�(�ȗ��\)
'          :�߂�l        ,O   ,FUNCTION_RETURN  ,���o�̐���
'����      :

Public Function UpdateXSDCZ(records As typ_XSDCZ, _
                                Optional lsSqlWhere$ = vbNullString, _
                                Optional lsSqlOrder$ = vbNullString, _
                                Optional lsUpdate$ = vbNullString) As FUNCTION_RETURN
On Error GoTo proc_err
    gErr.Push "s_XSDCZ_SQL.bas -- Function UpdateXSDCZ"

    Dim lsSql       As String       ''SQL�S��
'    Dim lsSqlBase   As String       ''SQL��{��(WHERE�߂̑O�܂�)
'    Dim rs          As OraDynaset   ''RecordSet
    Dim recCnt      As Long         ''���R�[�h��
    Dim nowtime     As Date         ''�T�[�o����
    Dim nowtime_sql As String       ''�T�[�o����(SQL��)
    
    ''�T�[�o�[���Ԏ擾
    nowtime = getSvrTime()

'>>>>> .Edit��SQL(UPDATE)���ɕύX�@2009/06/16 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
    
    With records
        
        ''SQL��g�ݗ��Ă�
        lsSql = "UPDATE XSDCZ SET" & vbLf
        
        ''�X�V���t
        lsSql = lsSql & " KDAYCZ = " & nowtime_sql & vbLf
        
        ''�u���b�NID������ԍ�
        If .CRYNUMCZ <> "" And Left(.CRYNUMCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",CRYNUMCZ = '" & .CRYNUMCZ & "'" & vbLf
        End If
        ''�i��
        If .HINBCZ <> "" And Left(.HINBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",HINBCZ = '" & .HINBCZ & "'" & vbLf
        End If
        ''�������J�n�ʒu
        If .INPOSCZ <> "" Then
            lsSql = lsSql & ",INPOSCZ = '" & CStr(CInt(.INPOSCZ)) & "'" & vbLf
        End If
        ''���i�ԍ������ԍ�
        If .REVNUMCZ <> "" Then
            lsSql = lsSql & ",REVNUMCZ = '" & CStr(CInt(.REVNUMCZ)) & "'" & vbLf
        End If
        ''�H��
        If .FACTORYCZ <> "" And Left(.FACTORYCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",FACTORYCZ = '" & .FACTORYCZ & "'" & vbLf
        End If
        ''���Ə���
        If .OPECZ <> "" And Left(.OPECZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",OPECZ = '" & .OPECZ & "'" & vbLf
        End If
        ''���H���A�Ԃ͓o�^���Ȃ�
'        ''�H���A��
'        If .KCKNTCZ <> "" Then
'            lsSql = lsSql & ",KCKNTCZ = '" & CStr(CInt(.KCKNTCZ)) & "'" & vbLf
'        End If
        ''���H���A�Ԃ͓o�^���Ȃ�
        ''SXLID
        If .SXLIDCZ <> "" And Left(.SXLIDCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",SXLIDCZ = '" & .SXLIDCZ & "'" & vbLf
        End If
        ''�����ԍ�
        If .XTALCZ <> "" And Left(.XTALCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",XTALCZ = '" & .XTALCZ & "'" & vbLf
        End If
        ''�ŏI�ʉߊǗ��H��
        If .NEKKNTCZ <> "" And Left(.NEKKNTCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",NEKKNTCZ = '" & .NEKKNTCZ & "'" & vbLf
        End If
        ''�ŏI�ʉߍH��
        If .NEWKNTCZ <> "" And Left(.NEWKNTCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",NEWKNTCZ = '" & .NEWKNTCZ & "'" & vbLf
        End If
        ''�ŏI�ʉߍ�Ƌ敪
        If .NEWKKBCZ <> "" And Left(.NEWKKBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",NEWKKBCZ = '" & .NEWKKBCZ & "'" & vbLf
        End If
        ''�ŏI�ʉߏ�����
        If .NEMACOCZ <> "" Then
            lsSql = lsSql & ",NEMACOCZ = '" & CStr(CInt(.NEMACOCZ)) & "'" & vbLf
        End If
        ''���݊Ǘ��H��
        If .GNKKNTCZ <> "" And Left(.GNKKNTCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",GNKKNTCZ = '" & .GNKKNTCZ & "'" & vbLf
        End If
        ''���ݍH��
        If .GNWKNTCZ <> "" And Left(.GNWKNTCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",GNWKNTCZ = '" & .GNWKNTCZ & "'" & vbLf
        End If
        ''���ݍ�Ƌ敪
        If .GNWKKBCZ <> "" And Left(.GNWKKBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",GNWKKBCZ = '" & .GNWKKBCZ & "'" & vbLf
        End If
        ''���ݏ�����
        If .GNMACOCZ <> "" Then
            lsSql = lsSql & ",GNMACOCZ = '" & CStr(CInt(.GNMACOCZ)) & "'" & vbLf
        End If
        ''���ݏ������t
        If lsUpdate = "NEW" Then    ''XSDCA�͐V�K�o�^�������ꍇ
            lsSql = lsSql & ",GNDAYCZ = " & nowtime_sql & vbLf
        Else                        ''XSDCA���X�V�������ꍇ
            If .GNDAYCZ <> "" And Left(.GNDAYCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",GNDAYCZ = TO_DATE('" & Format$(CDate(.GNDAYCZ), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
            End If
        End If
        ''���ݒ���
        If .GNLCZ <> "" Then
            lsSql = lsSql & ",GNLCZ = '" & CStr(CInt(.GNLCZ)) & "'" & vbLf
        End If
        ''���ݏd��
        If .GNWCZ <> "" Then
            lsSql = lsSql & ",GNWCZ = '" & CStr(CLng(.GNWCZ)) & "'" & vbLf
        End If
        ''���ݖ���
        If .GNMCZ <> "" Then
            lsSql = lsSql & ",GNMCZ = '" & CStr(CInt(.GNMCZ)) & "'" & vbLf
        End If
        ''SUMMIT����
        If .SUMITLCZ <> "" Then
            lsSql = lsSql & ",SUMITLCZ = '" & CStr(CInt(.SUMITLCZ)) & "'" & vbLf
        End If
        ''SUMMIT�d��
        If .SUMITWCZ <> "" Then
            lsSql = lsSql & ",SUMITWCZ = '" & CStr(CLng(.SUMITWCZ)) & "'" & vbLf
        End If
        ''SUMMIT����
        If .SUMITMCZ <> "" Then
            lsSql = lsSql & ",SUMITMCZ = '" & CStr(CInt(.SUMITMCZ)) & "'" & vbLf
        End If
        ''�`���[�W��
        If .CHGCZ <> "" Then
            lsSql = lsSql & ",CHGCZ = '" & CStr(CLng(.CHGCZ)) & "'" & vbLf
        End If
        ''���H�敪
        If .KAKOUBCZ <> "" And Left(.KAKOUBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",KAKOUBCZ = '" & .KAKOUBCZ & "'" & vbLf
        End If
        ''�v����t
        If .KEIDAYCZ <> "" Then
            lsSql = lsSql & ",KEIDAYCZ = TO_DATE('" & Format$(CDate(.KEIDAYCZ), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        ''�I�敪
        If .GNTKUBCZ <> "" And Left(.GNTKUBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",GNTKUBCZ = '" & .GNTKUBCZ & "'" & vbLf
        End If
        ''�I�敪
        If .GNTNOCZ <> "" And Left(.GNTNOCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",GNTNOCZ = '" & .GNTNOCZ & "'" & vbLf
        End If
        ''�����H��
        If .XTWORKCZ <> "" And Left(.XTWORKCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",XTWORKCZ = '" & .XTWORKCZ & "'" & vbLf
        End If
        ''�E�F�[�n����
        If .WFWORKCZ <> "" And Left(.WFWORKCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",WFWORKCZ = '" & .WFWORKCZ & "'" & vbLf
        End If
        ''�ŏI��ԋ敪
        If .LSTATBCZ <> "" And Left(.LSTATBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",LSTATBCZ = '" & .LSTATBCZ & "'" & vbLf
        End If
        ''������ԋ敪
        If .RSTATBCZ <> "" And Left(.RSTATBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",RSTATBCZ = '" & .RSTATBCZ & "'" & vbLf
        End If
        ''�i�㺰��
        If .LUFRCCZ <> "" And Left(.LUFRCCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",LUFRCCZ = '" & .LUFRCCZ & "'" & vbLf
        End If
        ''�i��敪
        If .LUFRBCZ <> "" And Left(.LUFRBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",LUFRBCZ = '" & .LUFRBCZ & "'" & vbLf
        End If
        ''�i������
        If .LDFRCCZ <> "" And Left(.LDFRCCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",LDFRCCZ = '" & .LDFRCCZ & "'" & vbLf
        End If
        ''�i���敪
        If .LDFRBCZ <> "" And Left(.LDFRBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",LDFRBCZ = '" & .LDFRBCZ & "'" & vbLf
        End If
        ''ΰ��޺���
        If .HOLDCCZ <> "" And Left(.HOLDCCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",HOLDCCZ = '" & .HOLDCCZ & "'" & vbLf
        End If
        ''�z�[���h�敪
        If .HOLDBCZ <> "" And Left(.HOLDBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",HOLDBCZ = '" & .HOLDBCZ & "'" & vbLf
        End If
        ''��O�敪
        If .EXKUBCZ <> "" And Left(.EXKUBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",EXKUBCZ = '" & .EXKUBCZ & "'" & vbLf
        End If
        ''�ԕi�敪
        If .HENPKCZ <> "" And Left(.HENPKCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",HENPKCZ = '" & .HENPKCZ & "'" & vbLf
        End If
        ''�����敪
        If .LIVKCZ <> "" And Left(.LIVKCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",LIVKCZ = '" & .LIVKCZ & "'" & vbLf
        End If
        ''�����敪
        If .KANKCZ <> "" And Left(.KANKCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",KANKCZ = '" & .KANKCZ & "'" & vbLf
        End If
        ''���ɋ敪
        If .NFCZ <> "" And Left(.NFCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",NFCZ = '" & .NFCZ & "'" & vbLf
        End If
        ''�폜�敪
        If .SAKJCZ <> "" And Left(.SAKJCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",SAKJCZ = '" & .SAKJCZ & "'" & vbLf
        End If
        ''�o�^���t
        If lsUpdate = "NEW" Then    ''XSDCA�͐V�K�o�^�������ꍇ
            lsSql = lsSql & ",TDAYCZ = " & nowtime_sql & vbLf
        Else                        ''XSDCA���X�V�������ꍇ
            If .TDAYCZ <> "" And Left(.TDAYCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",TDAYCZ = TO_DATE('" & Format$(CDate(.TDAYCZ), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
            End If
        End If
        ''SUMMIT���M�t���O
        If .SUMITBCZ <> "" And Left(.SUMITBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",SUMITBCZ = '" & .SUMITBCZ & "'" & vbLf
        End If
        ''���M�t���O
        If .SNDKCZ <> "" And Left(.SNDKCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",SNDKCZ = '" & .SNDKCZ & "'" & vbLf
        End If
        ''���M���t
        If lsUpdate = "NEW" Then    ''XSDCA�͐V�K�o�^�������ꍇ
            lsSql = lsSql & ",SNDDAYCZ = " & nowtime_sql & vbLf
        Else                        ''XSDCA���X�V�������ꍇ
            If .SNDDAYCZ <> "" And Left(.SNDDAYCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",SNDDAYCZ = TO_DATE('" & Format$(CDate(.SNDDAYCZ), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
            End If
        End If
        ''���x���o�͊m�F�t���O
        If .LBLFLGCZ <> "" Then
            lsSql = lsSql & ",LBLFLGCZ = '" & .LBLFLGCZ & "'" & vbLf
        End If
        ''�ؒf�����敪
        If .CUTCNTCZ <> "" And Left(.CUTCNTCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",CUTCNTCZ = '" & .CUTCNTCZ & "'" & vbLf
        End If
        ''��\�i�ԃt���O
        If .HINBFLGCZ <> "" And Left(.HINBFLGCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",HINBFLGCZ = '" & .HINBFLGCZ & "'" & vbLf
        End If
        ''�z�[���h�敪(WF)
        If .WFHOLDFLGCZ <> "" And Left(.WFHOLDFLGCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",WFHOLDFLGCZ = '" & .WFHOLDFLGCZ & "'" & vbLf
        End If
        ''�z�[���h�H��
        If .HOLDKTCZ <> "" And Left(.HOLDKTCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",HOLDKTCZ = '" & .HOLDKTCZ & "'" & vbLf
        End If
        ''�e�u���b�NID
        If .RPCRYNUMCZ <> "" And Left(.RPCRYNUMCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",RPCRYNUMCZ = '" & .RPCRYNUMCZ & "'" & vbLf
        End If
        ''�s�ǃR�[�h
        If .FCODECZ <> "" And Left(.FCODECZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",FCODECZ = '" & .FCODECZ & "'" & vbLf
        End If
        ''���������敪
        If .SGNKCZ <> "" And Left(.SGNKCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",SGNKCZ = '" & .SGNKCZ & "'" & vbLf
        End If
        ''�ؒf�敪
        If .CUTKCZ <> "" And Left(.CUTKCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",CUTKCZ = '" & .CUTKCZ & "'" & vbLf
        End If
        ''����
        If .PLANTCATCZ <> "" And Left(.PLANTCATCZ, 2) <> vbNullChar And Trim(.HINBCZ) <> "Z" And Trim(.HINBCZ) <> "G" Then
            lsSql = lsSql & ",PLANTCATCZ = '" & .PLANTCATCZ & "'" & vbLf
        End If
    
        lsSql = lsSql & " " & lsSqlWhere & vbLf
    
        'SQL�����s
        recCnt = OraDB.ExecuteSQL(lsSql)
        
        '�Ԃ�l��1�ȊO�̓G���[
        If recCnt < 0 Then
            GoTo proc_err
        ElseIf recCnt = 0 Then
            '0���X�V�c�G���[(�����ʂ�)
            UpdateXSDCZ = FUNCTION_RETURN_FAILURE
            Exit Function
        ElseIf recCnt > 1 Then
            '�������X�V�c�G���[(�����͕���SELECT�����ŏ��̈ꌏ�̂ݍX�V)
            UpdateXSDCZ = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
    
    End With
'<<<<< .Edit��SQL(UPDATE)���ɕύX�@2009/06/16 SETsw kubota ------------------

    UpdateXSDCZ = FUNCTION_RETURN_SUCCESS
    

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print lsSql
    gErr.HandleError
    UpdateXSDCZ = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'��INSERT��

'�T�v      :�e�[�u���uXSDCZ�v�Ƀ��R�[�h��}������
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:pXSDCZ �@�@  ,I  ,typ_XSDCZ        ,XSDCZ�X�V�p�ް�
'      �@�@:sErrMsg�@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�������݂̐���
Public Function CreateXSDCZ(pXSDCZ() As typ_XSDCZ, sErrMsg As String) As FUNCTION_RETURN

    Dim lsSql       As String       ''SQL�S��
    Dim sDbName     As String       ''ð��ٖ�
'    Dim rs          As OraDynaset   ''RecordSet
    Dim nowtime     As Date         ''�T�[�o����
    Dim nowtime_sql As String       ''�T�[�o����(SQL��)
    Dim i           As Long         ''�J�E���^
    
    ''�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_XSDCZ_SQL.bas -- Function CreateXSDCZ"
    sErrMsg = ""
    sDbName = "XSDCZ"
    
    ''�z��Ƀf�[�^���Ȃ��ꍇ
    If UBound(pXSDCZ()) < 1 Then
        CreateXSDCZ = FUNCTION_RETURN_FAILURE
        Exit Function
    End If
    
    ''�z��1�Ԗڂ̐e�u���b�NID���L�[��XSDCZ���폜����
    lsSql = ""
    lsSql = lsSql & " DELETE XSDCZ"
    lsSql = lsSql & " WHERE RPCRYNUMCZ = '" & pXSDCZ(LBound(pXSDCZ()) + 1).RPCRYNUMCZ & "'"
'    lsSql = lsSql & " WHERE RPCRYNUMCZ = '" & pXSDCZ(LBound(pXSDCZ()) + 1).CRYNUMCZ & "'"
    ''SQL���s
    OraDB.ExecuteSQL lsSql
    ''LOG�o��
'    WriteDBLog lsSql        '���ā@07/06/20 ooba

    ''�T�[�o�[���Ԏ擾
    nowtime = getSvrTime()
    
'>>>>> .AddNew��SQL(INSERT)���ɕύX�@2009/06/16 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
    For i = LBound(pXSDCZ()) + 1 To UBound(pXSDCZ())
        With pXSDCZ(i)
            
            lsSql = "INSERT INTO XSDCZ ("
            lsSql = lsSql & " CRYNUMCZ"     ' 1:��ۯ�ID�E�����ԍ�
            lsSql = lsSql & ",HINBCZ"       ' 2:�i��
            lsSql = lsSql & ",INPOSCZ"      ' 3:�������J�n�ʒu
            lsSql = lsSql & ",REVNUMCZ"     ' 4:���i�ԍ������ԍ�
            lsSql = lsSql & ",FACTORYCZ"    ' 5:�H��
            lsSql = lsSql & ",OPECZ"        ' 6:���Ə���
            'lssql = lssql & ",KCKNTCZ"     ' 7:�H���A��   �H���A�Ԃ͓o�^���Ȃ�(�����̂܂�)
            lsSql = lsSql & ",SXLIDCZ"      ' 8:SXLID
            lsSql = lsSql & ",XTALCZ"       ' 9:�����ԍ�
            lsSql = lsSql & ",NEKKNTCZ"     '10:�ŏI�ʉߊǗ��H��
            lsSql = lsSql & ",NEWKNTCZ"     '11:�ŏI�ʉߍH��
            lsSql = lsSql & ",NEWKKBCZ"     '12:�ŏI�ʉߍ�Ƌ敪
            lsSql = lsSql & ",NEMACOCZ"     '13:�ŏI�ʉߏ�����
            lsSql = lsSql & ",GNKKNTCZ"     '14:���݊Ǘ��H��
            lsSql = lsSql & ",GNWKNTCZ"     '15:���ݍH��
            lsSql = lsSql & ",GNWKKBCZ"     '16:���ݍ�Ƌ敪
            lsSql = lsSql & ",GNMACOCZ"     '17:���ݏ�����
            lsSql = lsSql & ",GNDAYCZ"      '18:���ݏ������t
            lsSql = lsSql & ",GNLCZ"        '19:���ݒ���
            lsSql = lsSql & ",GNWCZ"        '20:���ݏd��
            lsSql = lsSql & ",GNMCZ"        '21:���ݖ���
            lsSql = lsSql & ",SUMITLCZ"     '22:SUMMIT����
            lsSql = lsSql & ",SUMITWCZ"     '23:SUMMIT�d��
            lsSql = lsSql & ",SUMITMCZ"     '24:SUMMIT����
            lsSql = lsSql & ",CHGCZ"        '25:����ޗ�
            lsSql = lsSql & ",KAKOUBCZ"     '26:���H�敪
            If .KEIDAYCZ <> "" Then
                lsSql = lsSql & ",KEIDAYCZ"     '27:�v����t
            End If
            lsSql = lsSql & ",GNTKUBCZ"     '28:�I�敪
            lsSql = lsSql & ",GNTNOCZ"      '29:�I�ԍ�
            lsSql = lsSql & ",XTWORKCZ"     '30:�����H��
            lsSql = lsSql & ",WFWORKCZ"     '31:���ʐ���
            lsSql = lsSql & ",LSTATBCZ"     '32:�ŏI��ԋ敪
            lsSql = lsSql & ",RSTATBCZ"     '33:������ԋ敪
            lsSql = lsSql & ",LUFRCCZ"      '34:�i�㺰��
            lsSql = lsSql & ",LUFRBCZ"      '35:�i��敪
            lsSql = lsSql & ",LDFRCCZ"      '36:�i������
            lsSql = lsSql & ",LDFRBCZ"      '37:�i���敪
            lsSql = lsSql & ",HOLDCCZ"      '38:ΰ��޺���
            lsSql = lsSql & ",HOLDBCZ"      '39:ΰ��ދ敪
            lsSql = lsSql & ",EXKUBCZ"      '40:��O�敪
            lsSql = lsSql & ",HENPKCZ"      '41:�ԕi�敪
            lsSql = lsSql & ",LIVKCZ"       '42:�����敪
            lsSql = lsSql & ",KANKCZ"       '43:�����敪
            lsSql = lsSql & ",NFCZ"         '44:���ɋ敪
            lsSql = lsSql & ",SAKJCZ"       '45:�폜�敪
            lsSql = lsSql & ",TDAYCZ"       '46:�o�^���t
            lsSql = lsSql & ",KDAYCZ"       '47:�X�V���t
            lsSql = lsSql & ",SUMITBCZ"     '48:SUMMIT���M�׸�
            lsSql = lsSql & ",SNDKCZ"       '49:���M�׸�
            lsSql = lsSql & ",SNDDAYCZ"     '50:���M���t
            lsSql = lsSql & ",LBLFLGCZ"     '51:���x���o�͊m�F�t���O
            lsSql = lsSql & ",CUTCNTCZ"     '52:�V�K�^�Đ؋敪
            lsSql = lsSql & ",HINBFLGCZ"    '53:��\�i�ԃt���O
            lsSql = lsSql & ",HOLDKTCZ"     '54:ΰ��ލH��
            lsSql = lsSql & ",RPCRYNUMCZ"   '55:�e��ۯ�ID
            lsSql = lsSql & ",FCODECZ"      '56:�s�Ǻ���
            lsSql = lsSql & ",SGNKCZ"       '57:���������敪
            lsSql = lsSql & ",CUTKCZ"       '58:�ؒf�敪
            lsSql = lsSql & ",PLANTCATCZ"   '59:����
            lsSql = lsSql & ")"
            lsSql = lsSql & "VALUES (" & vbLf
            
            ' 1:��ۯ�ID�E�����ԍ�
            If .CRYNUMCZ <> "" And Left(.CRYNUMCZ, 1) <> vbNullChar Then
                lsSql = lsSql & " '" & .CRYNUMCZ & "'" & vbLf
            Else
                lsSql = lsSql & " '" & Space(12) & "'" & vbLf
            End If
            
            ' 2:�i��
            If .HINBCZ <> "" And Left(.HINBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .HINBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(8) & "'" & vbLf
            End If
            
            ' 3:�������J�n�ʒu
            If .INPOSCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CInt(.INPOSCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            ' 4:���i�ԍ������ԍ�
            If .REVNUMCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CInt(.REVNUMCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            ' 5:�H��
            If .FACTORYCZ <> "" And Left(.FACTORYCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .FACTORYCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(1) & "'" & vbLf
            End If
            
            ' 6:���Ə���
            If .OPECZ <> "" And Left(.OPECZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .OPECZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(1) & "'" & vbLf
            End If
            
            ' 7:�H���A��   �H���A�Ԃ͓o�^���Ȃ�(�����̂܂�)
            'If .KCKNTCZ <> "" Then
            '    lsSql = lsSql & ",'" & CStr(CInt(.KCKNTCZ)) & "'" & vbLf
            'Else
            '    lsSql = lsSql & ",0" & vbLf
            'End If
            
            ' 8:SXLID
            If .SXLIDCZ <> "" And Left(.SXLIDCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .SXLIDCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(13) & "'" & vbLf
            End If
            
            ' 9:�����ԍ�
            If .XTALCZ <> "" And Left(.XTALCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .XTALCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(12) & "'" & vbLf
            End If
            
            '10:�ŏI�ʉߊǗ��H��
            If .NEKKNTCZ <> "" And Left(.NEKKNTCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .NEKKNTCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(5) & "'" & vbLf
            End If
            
            '11:�ŏI�ʉߍH��
            If .NEWKNTCZ <> "" And Left(.NEWKNTCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .NEWKNTCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(5) & "'" & vbLf
            End If
            
            '12:�ŏI�ʉߍ�Ƌ敪
            If .NEWKKBCZ <> "" And Left(.NEWKKBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .NEWKKBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(2) & "'" & vbLf
            End If
            
            '13:�ŏI�ʉߏ�����
            If .NEMACOCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CInt(.NEMACOCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            '14:���݊Ǘ��H��
            If .GNKKNTCZ <> "" And Left(.GNKKNTCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .GNKKNTCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(5) & "'" & vbLf
            End If
            
            '15:���ݍH��
            If .GNWKNTCZ <> "" And Left(.GNWKNTCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .GNWKNTCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(5) & "'" & vbLf
            End If
            
            '16:���ݍ�Ƌ敪
            If .GNWKKBCZ <> "" And Left(.GNWKKBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .GNWKKBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(2) & "'" & vbLf
            End If
            
            '17:���ݏ�����
            If .GNMACOCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CInt(.GNMACOCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            '18:���ݏ������t
            lsSql = lsSql & "," & nowtime_sql & vbLf
            
            '19:���ݒ���
            If .GNLCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CInt(.GNLCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            '20:���ݏd��
            If .GNWCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CLng(.GNWCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            '21:���ݖ���
            If .GNMCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CInt(.GNMCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            '22:SUMMIT����
            If .SUMITLCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CInt(.SUMITLCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            '23:SUMMIT�d��
            If .SUMITWCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CLng(.SUMITWCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            '24:SUMMIT����
            If .SUMITMCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CInt(.SUMITMCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            '25:����ޗ�
            If .CHGCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CLng(.CHGCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            '26:���H�敪
            If .KAKOUBCZ <> "" And Left(.KAKOUBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .KAKOUBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(1) & "'" & vbLf
            End If
            
            '27:�v����t
            If .KEIDAYCZ <> "" Then
                lsSql = lsSql & ",TO_DATE('" & Format$(CDate(.KEIDAYCZ), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
            End If
            
            '28:�I�敪
            If .GNTKUBCZ <> "" And Left(.GNTKUBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .GNTKUBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(3) & "'" & vbLf
            End If
            
            '29:�I�ԍ�
            If .GNTNOCZ <> "" And Left(.GNTNOCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .GNTNOCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(4) & "'" & vbLf
            End If
            
            '30:�����H��
            lsSql = lsSql & ",'" & FACTORYCD & "'" & vbLf
            
            '31:���ʐ���
            If .WFWORKCZ <> "" And Left(.WFWORKCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .WFWORKCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(2) & "'" & vbLf
            End If

            '32:�ŏI��ԋ敪
            If .LSTATBCZ <> "" And Left(.LSTATBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .LSTATBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'T'" & vbLf       '�ʏ�
            End If

            '33:������ԋ敪
            If .RSTATBCZ <> "" And Left(.RSTATBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .RSTATBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'T'" & vbLf       '�ʏ�
            End If

            '34:�i�㺰��
            If .LUFRCCZ <> "" And Left(.LUFRCCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .LUFRCCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(3) & "'" & vbLf
            End If

            '35:�i��敪
            If .LUFRBCZ <> "" And Left(.LUFRBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .LUFRBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(1) & "'" & vbLf
            End If

            '36:�i������
            If .LDFRCCZ <> "" And Left(.LDFRCCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .LDFRCCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(3) & "'" & vbLf
            End If

            '37:�i���敪
            If .LDFRBCZ <> "" And Left(.LDFRBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .LDFRBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'0'" & vbLf
            End If

            '38:ΰ��޺���
            If .HOLDCCZ <> "" And Left(.HOLDCCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .HOLDCCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(3) & "'" & vbLf
            End If

            '39:ΰ��ދ敪
            If .HOLDBCZ <> "" And Left(.HOLDBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .HOLDBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'0'" & vbLf
            End If

            '40:��O�敪
            If .EXKUBCZ <> "" And Left(.EXKUBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .EXKUBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(1) & "'" & vbLf
            End If

            '41:�ԕi�敪
            If .HENPKCZ <> "" And Left(.HENPKCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .HENPKCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(1) & "'" & vbLf
            End If

            '42:�����敪
            If .LIVKCZ <> "" And Left(.LIVKCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .LIVKCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'0'" & vbLf
            End If

            '43:�����敪
            If .KANKCZ <> "" And Left(.KANKCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .KANKCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'0'" & vbLf
            End If

            '44:���ɋ敪
            If .NFCZ <> "" And Left(.NFCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .NFCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'0'" & vbLf
            End If

            '45:�폜�敪
            If .SAKJCZ <> "" And Left(.SAKJCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .SAKJCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'0'" & vbLf
            End If

            '46:�o�^���t
            lsSql = lsSql & "," & nowtime_sql & vbLf

            '47:�X�V���t
            lsSql = lsSql & "," & nowtime_sql & vbLf

            '48:SUMMIT���M�׸�
            lsSql = lsSql & ",'0'" & vbLf

            '49:���M�׸�
            lsSql = lsSql & ",'0'" & vbLf

            '50:���M���t
            lsSql = lsSql & ",NULL" & vbLf

            '51:���x���o�͊m�F�t���O
            If .LBLFLGCZ <> "" And Left(.LBLFLGCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .LBLFLGCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(1) & "'" & vbLf
            End If

            '52:�V�K�^�Đ؋敪
            If .CUTCNTCZ <> "" And Left(.CUTCNTCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .CUTCNTCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",NULL" & vbLf
            End If

            '53:��\�i�ԃt���O
            If .HINBFLGCZ <> "" And Left(.HINBFLGCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .HINBFLGCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",NULL" & vbLf
            End If

            '54:ΰ��ލH��
            If .HOLDKTCZ <> "" And Left(.HOLDKTCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .HOLDKTCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(5) & "'" & vbLf
            End If

            '55:�e��ۯ�ID
            If .RPCRYNUMCZ <> "" And Left(.RPCRYNUMCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .RPCRYNUMCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(12) & "'" & vbLf
            End If

            '56:�s�Ǻ���
            If .FCODECZ <> "" And Left(.FCODECZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .FCODECZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(3) & "'" & vbLf
            End If

            '57:���������敪
            If .SGNKCZ <> "" And Left(.SGNKCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .SGNKCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",NULL" & vbLf
            End If

            '58:�ؒf�敪
            If .CUTKCZ <> "" And Left(.CUTKCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .CUTKCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",NULL" & vbLf
            End If

            '59:����
            If .PLANTCATCZ <> "" And Left(.PLANTCATCZ, 2) <> vbNullChar Then
                lsSql = lsSql & ",'" & .PLANTCATCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",NULL" & vbLf
            End If
            
            lsSql = lsSql & ")" & vbLf
        
            'SQL�����s
            If OraDB.ExecuteSQL(lsSql) < 1 Then
                GoTo proc_err
            End If

            ' del SIRD�Ή� SETkimizuka Start 2010/02/15
            '''XODY3�쐬 �����Ď��@�\�ǉ��ɔ����C��  add SETkimizuka Start 09/08/03
            'If Y3Flg = True Then
            '    If pXSDCZ(i).SGNKCZ = "0" Then
            '        Call CreateOrUpdateXODY3(pXSDCZ(i).CRYNUMCZ, pXSDCZ(i).SXLIDCZ, pXSDCZ(i).LIVKCZ, "", "1")
            '    End If
            'End If
            ''XODY3�쐬 �����Ď��@�\�ǉ��ɔ����C��  add SETkimizuka End 09/08/03
            ' del SIRD�Ή� SETkimizuka End 2010/02/15

        End With
    Next i
'<<<<< .AddNew��SQL(INSERT)���ɕύX�@2009/06/16 SETsw kubota ------------------

    CreateXSDCZ = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print lsSql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", "DB", sDbName)
    CreateXSDCZ = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :�ŏI�ʉߏ����񐔂��擾����
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:p_sCrynum    ,I  ,String           ,�u���b�NID
'      �@�@:p_iInpos     ,I  ,Integer          ,�J�n�ʒu
'      �@�@:�߂�l       ,O  ,Integer        �@,������
Public Function GetGNMACOCZ(p_sCrynum As String, p_iInpos As Integer) As Integer
    Dim sql As String
    Dim rs As OraDynaset
    
    
    sql = "SELECT GNMACOCZ FROM XSDCZ WHERE CRYNUMCZ = '" & p_sCrynum
    sql = sql & "' AND INPOSCZ = " & p_iInpos
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    If rs.RecordCount = 0 Then
        GetGNMACOCZ = 1
    Else
        GetGNMACOCZ = CInt(rs.Fields("GNMACOCZ"))
    End If

End Function

'�T�v      :�Y������ں��ޗL��������
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:p_BlockID    ,I  ,String           ,SXLID
'      �@�@:p_Hinban     ,O  ,String           ,����
'      �@�@:p_Inpos      ,O  ,Integer          ,����
'      �@�@:�߂�l       ,O  ,Boolean        �@,ں��ނȂ�(TRUE)/����(FALSE)
'�����@�@�@�F�i�ԐU�ցA�ؽ�يi��Ȃǎ�ں��ނƓ��i�Ԃւ̕ύX�ɑΉ�
'�����@�@�@�F2002/08/29 ohno
Public Function CheckUniqueRecordXSDCZ(p_BlockID As String, p_Hinban As String, p_Inpos As Integer) As Boolean
    Dim sql As String
    Dim rs As OraDynaset
    
    sql = "SELECT * FROM XSDCZ WHERE CRYNUMCZ = '" & p_BlockID
    sql = sql & "' AND HINBCZ = '" & p_Hinban
    sql = sql & "' AND INPOSCZ = " & p_Inpos
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    If rs.RecordCount = 0 Then
        CheckUniqueRecordXSDCZ = True
    Else
        CheckUniqueRecordXSDCZ = False
    End If
    
End Function

'��DELETE��

'�T�v      :�e�[�u���uXSDCZ�v�̃��R�[�h���폜����
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:sCRYNUMCZ �@ ,I  ,String           ,�u���b�NID
'      �@�@:sHINBCZ �@   ,I  ,String           ,�i��
'      �@�@:sINPOSCZ �@  ,I  ,String           ,�������J�n�ʒu
'      �@�@:sErrMsg�@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l       ,O  ,FUNCTION_RETURN�@,�������݂̐���
Public Function DeleteXSDCZ(sCRYNUMCZ As String, sHINBCZ As String, sINPOSCZ As String, _
                                sErrMsg As String) As FUNCTION_RETURN
    Dim lsSql       As String       ''SQL�S��
    Dim sDbName     As String       ''ð��ٖ�
    
    ''�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_XSDCZ_SQL.bas -- Function DeleteXSDCZ"
    sErrMsg = ""
    sDbName = "XSDCZ"
    
    DeleteXSDCZ = FUNCTION_RETURN_FAILURE
    
    ''�u���b�NID���L�[��XSDCZ���폜����
    lsSql = ""
    lsSql = lsSql & " DELETE XSDCZ"
    lsSql = lsSql & " WHERE CRYNUMCZ = '" & sCRYNUMCZ & "'"
    lsSql = lsSql & "   AND HINBCZ   = '" & sHINBCZ & "'"
    lsSql = lsSql & "   AND INPOSCZ  = '" & sINPOSCZ & "'"
    ''SQL���s
    OraDB.ExecuteSQL lsSql
    ''LOG�o��
'    WriteDBLog lsSql        '���ā@07/06/20 ooba

    DeleteXSDCZ = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print lsSql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", "DB", sDbName)
    DeleteXSDCZ = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


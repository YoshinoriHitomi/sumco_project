Attribute VB_Name = "SB_cmzcXSDCE"
Option Explicit
'                                     2003/09/05
'======================================================
' ���[�U��`�^�̐錾
'======================================================

' �U�֗���
Public Type typ_XSDCE
    CRYNUMCE As String * 12         ' �u���b�NID�E�����ԍ�
    INPOSCE As Integer              ' �������J�n�ʒu
    KCNTCE As Integer               ' �H���A��
    HINBCE As String * 8            ' �U�֐�i��
    REVNUMCE As Integer             ' ���i�ԍ������ԍ�(�U�֐�)
    FACTORYCE As String * 1         ' �H��(�U�֐�)
    OPECE As String * 1             ' ���Ə���(�U�֐�)
    MOTHINCE As String * 8          ' �U�֌��i��
    MREVNUMCE As Integer            ' ���i�ԍ������ԍ�(�U�֌�)
    MFACTORYCE As String * 1        ' �H��(�U�֌�)
    MOPECE As String * 1            ' ���Ə���(�U�֌�)
    SXLIDCE As String * 13          ' SXLID
    WKKTCE As String * 5            ' �H��
    KNKTCE As String * 5            ' �Ǘ��H��
    REPSMPLIDTCE As String * 16     ' ��\�T���v��ID(TOP)
    REPSMPLIDBCE As String * 16     ' ��\�T���v��ID(BOT)
    TOKNUMCE As String * 10         ' ���̔ԍ�
    TOKCAUSECE As String * 200      ' ���̗��R
    TOKCODECE As String * 2         ' ���̗��R�R�[�h
    ERRCAUSECE As String * 50       ' �G���[���R
    HULCE As Integer                ' �U�֒���
    HUWCE As Long                   ' �U�֏d��
    HUMCE As Integer                ' �U�֖���
    TSTAFFCE As String * 8          ' �o�^�Ј�ID
    TDAYCE As Date                  ' �o�^���t
    KSTAFFCE As String * 8          ' �X�V�Ј�ID
    KDAYCE As Date                  ' �X�V���t
    SNDKCE As String * 1            ' ���M�t���O
    SNDDAYCE As Date                ' ���M���t
End Type

'�X�V�p
Public Type typ_XSDCE_Update
    CRYNUMCE As String              ' �u���b�NID�E�����ԍ�
    INPOSCE As String               ' �������J�n�ʒu
    KCNTCE As String                ' �H���A��
    HINBCE As String                ' �U�֐�i��
    REVNUMCE As String              ' ���i�ԍ������ԍ�(�U�֐�)
    FACTORYCE As String             ' �H��(�U�֐�)
    OPECE As String                 ' ���Ə���(�U�֐�)
    MOTHINCE As String              ' �U�֌��i��
    MREVNUMCE As String             ' ���i�ԍ������ԍ�(�U�֌�)
    MFACTORYCE As String            ' �H��(�U�֌�)
    MOPECE As String                ' ���Ə���(�U�֌�)
    SXLIDCE As String               ' SXLID
    WKKTCE As String                ' �H��
    KNKTCE As String                ' �Ǘ��H��
    REPSMPLIDTCE As String          ' ��\�T���v��ID(TOP)
    REPSMPLIDBCE As String          ' ��\�T���v��ID(BOT)
    TOKNUMCE As String              ' ���̔ԍ�
    TOKCAUSECE As String            ' ���̗��R
    TOKCODECE As String             ' ���̗��R�R�[�h
    ERRCAUSECE As String            ' �G���[���R
    HULCE As String                 ' �U�֒���
    HUWCE As String                 ' �U�֏d��
    HUMCE As String                 ' �U�֖���
    TSTAFFCE As String              ' �o�^�Ј�ID
    TDAYCE As String                ' �o�^���t
    KSTAFFCE As String              ' �X�V�Ј�ID
    KDAYCE As String                ' �X�V���t
    SNDKCE As String                ' ���M�t���O
    SNDDAYCE As String              ' ���M���t
End Type

'�T�v      :�H���A�Ԃ��擾����
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:p_sCrynum    ,I  ,String           ,�u���b�NID�E�����ԍ�
'      �@�@:p_sGenKotei  ,I  ,String           ,���ݍH��
'      �@�@:�߂�l       ,O  ,Integer        �@,�H���A��
'����      :�����ԍ��ƍH������H���A�Ԃ��擾����
Public Function GetKCNTC3(p_sCrynum As String, p_sGenKotei As String) As Integer
    Dim sql As String
    Dim rs As OraDynaset
    
    sql = "SELECT MAX(KCNTC3) AS MAXKCNTC3 FROM XSDC3 "
    sql = sql & "WHERE CRYNUMC3 = '" & p_sCrynum & "' "
    sql = sql & "AND   WKKTC3   = '" & p_sGenKotei & "' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If IsNull(rs.Fields("MAXKCNTC3")) Then
        GetKCNTC3 = 0
    Else
        GetKCNTC3 = CInt(rs.Fields("MAXKCNTC3"))
    End If
End Function

'�T�v      :�H���A�Ԃ��擾����
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:p_sRpCrynum  ,I  ,String           ,�e�u���b�NID
'      �@�@:p_sGenKotei  ,I  ,String           ,���ݍH��
'      �@�@:�߂�l       ,O  ,Integer        �@,�H���A��
'����      :�e�u���b�NID�ƍH������H���A�Ԃ��擾����
'����      :05/09/16 ooba
Public Function GetKCNTC3_New(p_sRpCrynum As String, p_sGenKotei As String) As Integer
    Dim sql As String
    Dim rs As OraDynaset
    
    sql = "SELECT MAX(KCNTC3) AS MAXKCNTC3 FROM XSDC3 "
    sql = sql & "WHERE RPCRYNUMC3 = '" & p_sRpCrynum & "' "
    sql = sql & "AND   WKKTC3   = '" & p_sGenKotei & "' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If IsNull(rs.Fields("MAXKCNTC3")) Then
        GetKCNTC3_New = 0
    Else
        GetKCNTC3_New = CInt(rs.Fields("MAXKCNTC3"))
    End If
End Function

'�T�v      :�������J�n�ʒu���擾����
'���Ұ��@�@:�ϐ���       ,IO ,�^               ,����
'      �@�@:p_sCrynum    ,I  ,String           ,�u���b�NID�E�����ԍ�
'      �@�@:p_iInpos     ,I  ,Integer          ,�������J�n�ʒu
'      �@�@:p_sKcnt      ,I  ,Integer          ,�H���A��
'      �@�@:�߂�l       ,O  ,Integer        �@,�������J�n�ʒu
'����      :�����ԍ��ƍH���A�Ԃ��猋�����J�n�ʒu���擾����
Public Function GetINPOSC3(p_sCrynum As String, p_iInpos As Integer, p_iKcnt As Integer) As Integer
    Dim sql As String
    Dim rs As OraDynaset
    
    sql = "SELECT INPOSC3 FROM XSDC3 "
    sql = sql & "WHERE CRYNUMC3 = '" & p_sCrynum & "' "
    sql = sql & "AND   INPOSC3  < " & p_iInpos & " "
    sql = sql & "AND   KCNTC3   = " & p_iKcnt & " "
    sql = sql & "ORDER BY INPOSC3 desc "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If IsNull(rs.Fields("INPOSC3")) Then
        GetINPOSC3 = -1
    Else
        GetINPOSC3 = CInt(rs.Fields("INPOSC3"))
    End If
End Function

'��INSERT��  NULL�̏ꍇ�Achar�Ȃ�X�y�[�X�ANumber�Ȃ�NULL������

'�T�v      :�e�[�u���uXSDCE�v�Ƀ��R�[�h��}������
'���Ұ��@�@:�ϐ���       ,IO ,�^                ,����
'      �@�@:pXSDCE �@�@  ,I  ,typ_XSDCE_Update  ,XSDCE�X�V�p�ް�
'      �@�@:sErrMsg�@�@�@,O  ,String         �@ ,�G���[���b�Z�[�W
'      �@�@:�߂�l       ,O  ,Boolean        �@ ,True:OK False:NG
'����      :�����Ǘ��c�a�̓o�^���s��
Public Function CreateXSDCE(pXSDCE As typ_XSDCE_Update, sErrMsg As String) As Boolean
    
    Dim sql As String
    Dim sDbName As String
'    Dim rs As OraDynaset    'RecordSet
    Dim nowtime As Date
    Dim nowtime_sql As String       ''�T�[�o����(SQL��)
    
    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcXSDCE.bas -- Function CreateXSDCE"
    sErrMsg = ""
    sDbName = "XSDCE"
    nowtime = getSvrTime()    '�T�[�o�[�̎��Ԃ��擾����悤�ɕύX 2003/6/4 tuku
    
'>>>>> .AddNew��SQL(INSERT)���ɕύX�@2009/06/16 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
    With pXSDCE
        sql = "INSERT INTO XSDCE ( "
        sql = sql & " CRYNUMCE"         ' 1:��ۯ�ID�E�����ԍ�
        sql = sql & ",INPOSCE"          ' 2:�������J�n�ʒu
        sql = sql & ",KCNTCE"           ' 3:�H���A��
        sql = sql & ",HINBCE"           ' 4:�U�֐�i��
        sql = sql & ",REVNUMCE"         ' 5:���i�ԍ������ԍ�(�U�֐�)
        sql = sql & ",FACTORYCE"        ' 6:�H��(�U�֐�)
        sql = sql & ",OPECE"            ' 7:���Ə���(�U�֐�)
        sql = sql & ",MOTHINCE"         ' 8:�U�֌��i��
        sql = sql & ",MREVNUMCE"        ' 9:���i�ԍ������ԍ�(�U�֌�)
        sql = sql & ",MFACTORYCE"       '10:�H��(�U�֌�)
        sql = sql & ",MOPECE"           '11:���Ə���(�U�֐�)
        sql = sql & ",SXLIDCE"          '12:SXLID
        sql = sql & ",WKKTCE"           '13:�H��
        sql = sql & ",KNKTCE"           '14:�Ǘ��H��
        sql = sql & ",REPSMPLIDTCE"     '15:��\�T���v��ID(TOP)
        sql = sql & ",REPSMPLIDBCE"     '16:��\�T���v��ID(BOT)
        sql = sql & ",TOKNUMCE"         '17:���̔ԍ�
        sql = sql & ",TOKCAUSECE"       '18:���̗��R
        sql = sql & ",TOKCODECE"        '19:���̗��R�R�[�h
        sql = sql & ",ERRCAUSECE"       '20:�G���[���R
        sql = sql & ",HULCE"            '21:�U�֒���
        sql = sql & ",HUWCE"            '22:�U�֏d��
        sql = sql & ",HUMCE"            '23:�U�֖���
        sql = sql & ",TSTAFFCE"         '24:�o�^�Ј�ID
        sql = sql & ",TDAYCE"           '25:�o�^���t
        sql = sql & ",KSTAFFCE"         '26:�X�V�Ј�ID
        sql = sql & ",KDAYCE"           '27:�X�V���t
        sql = sql & ",SNDKCE"           '28:���M�t���O
        sql = sql & ",SNDDAYCE"         '29:���M���t
        sql = sql & ") "
        sql = sql & "VALUES ( "
        
        ' 1:��ۯ�ID�E�����ԍ�
        If .CRYNUMCE <> "" Then
            sql = sql & " '" & .CRYNUMCE & "'"
        Else
            sql = sql & " '" & Space(12) & "'"
        End If
        
        ' 2:�������J�n�ʒu
        If .INPOSCE <> "" Then
            sql = sql & ",'" & CStr(CInt(.INPOSCE)) & "'"
        Else
            sql = sql & ",0"
        End If
        
        ' 3:�H���A��
        If .KCNTCE <> "" Then
            sql = sql & ",'" & CStr(CInt(.KCNTCE)) & "'"
        Else
            sql = sql & ",0"
        End If
        
        ' 4:�U�֐�i��
        If .HINBCE <> "" Then
            sql = sql & ",'" & .HINBCE & "'"
        Else
            sql = sql & ",'" & Space(8) & "'"
        End If
        
        ' 5:���i�ԍ������ԍ�(�U�֐�)
        If .REVNUMCE <> "" Then
            sql = sql & ",'" & CStr(CInt(.REVNUMCE)) & "'"
        Else
            sql = sql & ",0"
        End If
        
        ' 6:�H��(�U�֐�)
        If .FACTORYCE <> "" Then
            sql = sql & ",'" & .FACTORYCE & "'"
        Else
            sql = sql & ",'" & Space(1) & "'"
        End If
        
        ' 7:���Ə���(�U�֐�)
        If .OPECE <> "" Then
            sql = sql & ",'" & .OPECE & "'"
        Else
            sql = sql & ",'" & Space(1) & "'"
        End If
        
        ' 8:�U�֌��i��
        If .MOTHINCE <> "" Then
            sql = sql & ",'" & .MOTHINCE & "'"
        Else
            sql = sql & ",'" & Space(8) & "'"
        End If
        
        ' 9:���i�ԍ������ԍ�(�U�֌�)
        If .MREVNUMCE <> "" Then
            sql = sql & ",'" & CStr(CInt(.MREVNUMCE)) & "'"
        Else
            sql = sql & ",'0'"
        End If
        
        '10:�H��(�U�֌�)
        If .MFACTORYCE <> "" Then
            sql = sql & ",'" & .MFACTORYCE & "'"
        Else
            sql = sql & ",'" & Space(1) & "'"
        End If
        
        '11:���Ə���(�U�֐�)
        If .MOPECE <> "" Then
            sql = sql & ",'" & .MOPECE & "'"
        Else
            sql = sql & ",'" & Space(1) & "'"
        End If
        
        '12:SXLID
        If .SXLIDCE <> "" And Left(.SXLIDCE, 1) <> vbNullChar Then
            sql = sql & ",'" & .SXLIDCE & "'"
        Else
            sql = sql & ",'" & Space(13) & "'"
        End If
        
        '13:�H��
        If .WKKTCE <> "" Then
            sql = sql & ",'" & .WKKTCE & "'"
        Else
            sql = sql & ",'" & Space(5) & "'"
        End If
        
        '14:�Ǘ��H��
        If .KNKTCE <> "" Then
            sql = sql & ",'" & .KNKTCE & "'"
        Else
            sql = sql & ",'" & Space(5) & "'"
        End If
        
        '15:��\�T���v��ID(TOP)
        If .REPSMPLIDTCE <> "" Then
            sql = sql & ",'" & .REPSMPLIDTCE & "'"
        Else
            sql = sql & ",'" & Space(16) & "'"
        End If
        
        '16:��\�T���v��ID(BOT)
        If .REPSMPLIDBCE <> "" Then
            sql = sql & ",'" & .REPSMPLIDBCE & "'"
        Else
            sql = sql & ",'" & Space(16) & "'"
        End If
        
        '17:���̔ԍ�
        If .TOKNUMCE <> "" Then
            sql = sql & ",'" & .TOKNUMCE & "'"
        Else
            sql = sql & ",'" & Space(10) & "'"
        End If
        
        '18:���̗��R
        If .TOKCAUSECE <> "" Then
            sql = sql & ",'" & .TOKCAUSECE & "'"
        Else
            sql = sql & ",NULL"
        End If
        
        '19:���̗��R�R�[�h
        If .TOKCODECE <> "" Then
            sql = sql & ",'" & .TOKCODECE & "'"
        Else
            sql = sql & ",'" & Space(2) & "'"
        End If
        
        '20:�G���[���R
        If .ERRCAUSECE <> "" Then
            sql = sql & ",'" & .ERRCAUSECE & "'"
        Else
            sql = sql & ",NULL"
        End If
        
        '21:�U�֒���
        If .HULCE <> "" Then
            sql = sql & ",'" & CStr(CInt(.HULCE)) & "'"
        Else
            sql = sql & ",0"
        End If
        
        '22:�U�֏d��
        If .HUWCE <> "" Then
            sql = sql & ",'" & CStr(CLng(.HUWCE)) & "'"
        Else
            sql = sql & ",0"
        End If
        
        '23:�U�֖���
        If .HUMCE <> "" Then
            sql = sql & ",'" & CStr(CInt(.HUMCE)) & "'"
        Else
            sql = sql & ",0"
        End If
        
        '24:�o�^�Ј�ID
        If .TSTAFFCE <> "" Then
            sql = sql & ",'" & .TSTAFFCE & "'"
        Else
            sql = sql & ",'" & Space(8) & "'"
        End If
        
        '25:�o�^���t
        sql = sql & "," & nowtime_sql
        
        '26:�X�V�Ј�ID
        If .KSTAFFCE <> "" Then
            sql = sql & ",'" & .KSTAFFCE & "'"
        Else
            sql = sql & ",'" & Space(8) & "'"
        End If
        
        '27:�X�V���t
        sql = sql & "," & nowtime_sql
        
        '28:���M�t���O
        If .SNDKCE <> "" Then
            sql = sql & ",'" & .SNDKCE & "'"
        Else
            sql = sql & ",'" & Space(1) & "'"
        End If
        
        '29:���M���t
        If .SNDDAYCE <> "" Then
            sql = sql & ",TO_DATE('" & Format$(CDate(.SNDDAYCE), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
        Else
            sql = sql & ",NULL"
        End If
        
        sql = sql & ")"
    
        'SQL�����s
        If OraDB.ExecuteSQL(sql) < 1 Then
            GoTo proc_err
        End If
    
    End With
'<<<<< .AddNew��SQL(INSERT)���ɕύX�@2009/06/16 SETsw kubota ------------------

    CreateXSDCE = True

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
    CreateXSDCE = False
    Resume proc_exit

End Function


Attribute VB_Name = "s_XSDCS_1_SQL"
Option Explicit
'***�e�[�u���uXSDCS_1�v�ւ̃f�[�^�A�N�Z�X�֐�***

Public Type typ_XSDCS_1
    CRYNUMCS1       As String       ' �u���b�NID
    XTALCS1         As String       ' �����ԍ�
    INPOSCS1        As String       ' �������ʒu
    HINBCS1         As String       ' �i��
    REVNUMCS1       As String       ' �i�Ԑ��i�ԍ������ԍ�
    FACTORYCS1      As String       ' �i�ԍH��
    OPECS1          As String       ' �i�ԑ��Ə���
    TRANCNTFRSCS1   As String       ' ������(FRS)
    CRYINDOIFRSCS1  As String       ' ���FLG(FRS)
    CRYRESOIFRSCS1  As String       ' ����FLG(FRS)
    RPCRYNUMCS1     As String       ' �e�u���b�NID
    LIVKCS1         As String       ' �����敪
    TSTAFFCS1       As String       ' �o�^�Ј�ID
    TDAYCS1         As String       ' �o�^���t
    KSTAFFCS1       As String       ' �X�V�Ј�ID
    KDAYCS1         As String       ' �X�V���t
    SNDKCS1         As String       ' ���M�t���O
    SNDDAYCS1       As String       ' ���M���t
    SNDKDWHCS1      As String       ' ���M�t���O(DWH)
    SDAYDWHCS1      As String       ' ���M���t(DWH)
    SNDKSPCCS1      As String       ' ���M�t���O(SPC)
    SDAYSPCCS1      As String       ' ���M���t(SPC)
    SAKJCS1         As String       ' �폜�敪
End Type

Private Const SQRT = "'"

'�T�v      :�e�[�u���uXSDCS_1�v��������ɂ��������R�[�h�𒊏o����
'���Ұ��@�@:�u���b�NID
'           XSDCS_1�@�@���o���ʂP�z��ڂ���i�O�z��ږ��g�p�j
'�߂�l    :���o�̐��� Boolean
'����      :��ۯ�ID��XSDC_1������
'����      :2011/02/28�@�쐬 SMPK H.Ohkubo
Public Function GetXSDCS_1(sBlock As String, typXSDCS1() As typ_XSDCS_1) As Boolean
    Dim objDS       As Object
    Dim sSQL        As String
    Dim recCnt      As Integer
    Dim lRecCnt     As Long
    Dim lDtCnt      As Long
    
    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_XSDCS_1_SQL.bas -- Function GetXSDCS_1"
    
    '' ���Y�����R�[�h�����擾���A�����ꍇ�A���e�[�u������
    If CheckUniqueRecordXSDCS_1(sBlock) = False Then
                
        '' �Y���f�[�^�����̏ꍇ�A��̉��e�[�u���𐶐�����
        ReDim Preserve typXSDCS1(1) As typ_XSDCS_1
        
        With typXSDCS1(1)
            .CRYNUMCS1 = sBlock                     '' �u���b�NID
            .XTALCS1 = left(sBlock, 9) & "000"      '' �����ԍ�
            .INPOSCS1 = ""                          '' �������ʒu
            .HINBCS1 = ""                           '' �i��
            .REVNUMCS1 = ""                         '' �i�Ԑ��i�ԍ������ԍ�
            .FACTORYCS1 = ""                        '' �i�ԍH��
            .OPECS1 = ""                            '' �i�ԑ��Ə���
            .TRANCNTFRSCS1 = "0"                    '' ������(FRS)
            .CRYINDOIFRSCS1 = "0"                   '' ���FLG
            .CRYRESOIFRSCS1 = "0"                   '' ����FLG
            .RPCRYNUMCS1 = left(sBlock, 9) & "000"  '' �e�u���b�NID
            .LIVKCS1 = "0"                          '' �����敪
            .TSTAFFCS1 = ""                         '' �o�^�Ј�ID
            .TDAYCS1 = ""                           '' �o�^���t
            .KSTAFFCS1 = ""                         '' �X�V�Ј�ID
            .KDAYCS1 = ""                           '' �X�V���t
            .SNDKCS1 = ""                           '' ���M�t���O
            .SNDDAYCS1 = ""                         '' ���M���t
            .SNDKDWHCS1 = ""                        '' ���M�t���O(DWH)
            .SDAYDWHCS1 = ""                        '' ���M���t(DWH)
            .SNDKSPCCS1 = ""                        '' ���M�t���O(SPC)
            .SDAYSPCCS1 = ""                        '' ���M���t(SPC)
            .SAKJCS1 = ""                           '' �폜�敪
        End With
        
        GetXSDCS_1 = True
        Exit Function
    End If
    
    '' ���Y�����R�[�h�f�[�^����
    sSQL = "select"
    sSQL = sSQL & "  CRYNUMCS1"           '' �u���b�NID
    sSQL = sSQL & ", XTALCS1"             '' �����ԍ�
    sSQL = sSQL & ", INPOSCS1"            '' �������ʒu
    sSQL = sSQL & ", HINBCS1"             '' �i��
    sSQL = sSQL & ", REVNUMCS1"           '' �i�Ԑ��i�ԍ������ԍ�
    sSQL = sSQL & ", FACTORYCS1"          '' �i�ԍH��
    sSQL = sSQL & ", OPECS1"              '' �i�ԑ��Ə���
    sSQL = sSQL & ", TRANCNTFRSCS1"       '' ������(FRS)
    sSQL = sSQL & ", CRYINDOIFRSCS1"      '' ���FLG(FRS)
    sSQL = sSQL & ", CRYRESOIFRSCS1"      '' ����FLG(FRS)
    sSQL = sSQL & ", RPCRYNUMCS1"         '' �e�u���b�NID
    sSQL = sSQL & ", LIVKCS1"             '' �����敪
    sSQL = sSQL & ", TSTAFFCS1"           '' �o�^�Ј�ID
    sSQL = sSQL & ", TDAYCS1"             '' �o�^���t
    sSQL = sSQL & ", KSTAFFCS1"           '' �X�V�Ј�ID
    sSQL = sSQL & ", KDAYCS1"             '' �X�V���t
    sSQL = sSQL & ", SNDKCS1"             '' ���M�t���O
    sSQL = sSQL & ", SNDDAYCS1"           '' ���M���t
    sSQL = sSQL & ", SNDKDWHCS1"          '' ���M�t���O(DWH)
    sSQL = sSQL & ", SDAYDWHCS1"          '' ���M���t(DWH)
    sSQL = sSQL & ", SNDKSPCCS1"          '' ���M�t���O(SPC)
    sSQL = sSQL & ", SDAYSPCCS1"          '' ���M���t(SPC)
    sSQL = sSQL & ", SAKJCS1"             '' �폜�敪
    sSQL = sSQL & " from XSDCS_1"
    sSQL = sSQL & " where CRYNUMCS1 like '" & Trim(sBlock) & "%'"
    
    ''�f�[�^�𒊏o����
#If SRC_200_FLG = 1 Then
    If DynSet(objDS, sSQL) = False Then
        Call MsgOut(100, sSQL, ERR_DISP_LOG)
        GetXSDCS_1 = False
        Exit Function
    End If
#Else
    If DynSet2(objDS, sSQL) = False Then
        Call MsgOut(100, sSQL, ERR_DISP_LOG)
        GetXSDCS_1 = False
        Exit Function
    End If
#End If
    
    ReDim typXSDCS1(0)
    lRecCnt = 0
    ''���o���ʂ��i�[����
    If objDS.EOF = False Then
        Do Until objDS.EOF '�f�[�^���Ȃ��Ȃ�܂Ŏ擾
            
            lRecCnt = lRecCnt + 1
            ReDim Preserve typXSDCS1(lRecCnt) As typ_XSDCS_1
            With typXSDCS1(lRecCnt)
                If IsNull(objDS.Fields("CRYNUMCS1")) = False Then .CRYNUMCS1 = objDS.Fields("CRYNUMCS1")                    '' �u���b�NID
                If IsNull(objDS.Fields("XTALCS1")) = False Then .XTALCS1 = objDS.Fields("XTALCS1")                          '' �����ԍ�
                If IsNull(objDS.Fields("INPOSCS1")) = False Then .INPOSCS1 = objDS.Fields("INPOSCS1")                       '' �������ʒu
                If IsNull(objDS.Fields("HINBCS1")) = False Then .HINBCS1 = objDS.Fields("HINBCS1")                          '' �i��
                If IsNull(objDS.Fields("REVNUMCS1")) = False Then .REVNUMCS1 = objDS.Fields("REVNUMCS1")                    '' �i�Ԑ��i�ԍ������ԍ�
                If IsNull(objDS.Fields("FACTORYCS1")) = False Then .FACTORYCS1 = objDS.Fields("FACTORYCS1")                 '' �i�ԍH��
                If IsNull(objDS.Fields("OPECS1")) = False Then .OPECS1 = objDS.Fields("OPECS1")                             '' �i�ԑ��Ə���
                If IsNull(objDS.Fields("TRANCNTFRSCS1")) = False Then .TRANCNTFRSCS1 = objDS.Fields("TRANCNTFRSCS1")        '' ������(FRS)
                If IsNull(objDS.Fields("CRYINDOIFRSCS1")) = False Then .CRYINDOIFRSCS1 = objDS.Fields("CRYINDOIFRSCS1")     '' ���FLG
                If IsNull(objDS.Fields("CRYRESOIFRSCS1")) = False Then .CRYRESOIFRSCS1 = objDS.Fields("CRYRESOIFRSCS1")     '' ����FLG
                If IsNull(objDS.Fields("RPCRYNUMCS1")) = False Then .RPCRYNUMCS1 = objDS.Fields("RPCRYNUMCS1")              '' �e�u���b�NID
                If IsNull(objDS.Fields("LIVKCS1")) = False Then .LIVKCS1 = objDS.Fields("LIVKCS1")                          '' �����敪
                If IsNull(objDS.Fields("TSTAFFCS1")) = False Then .TSTAFFCS1 = objDS.Fields("TSTAFFCS1")                    '' �o�^�Ј�ID
                If IsNull(objDS.Fields("TDAYCS1")) = False Then .TDAYCS1 = objDS.Fields("TDAYCS1")                          '' �o�^���t
                If IsNull(objDS.Fields("KSTAFFCS1")) = False Then .KSTAFFCS1 = objDS.Fields("KSTAFFCS1")                    '' �X�V�Ј�ID
                If IsNull(objDS.Fields("KDAYCS1")) = False Then .KDAYCS1 = objDS.Fields("KDAYCS1")                          '' �X�V���t
                If IsNull(objDS.Fields("SNDKCS1")) = False Then .SNDKCS1 = objDS.Fields("SNDKCS1")                          '' ���M�t���O
                If IsNull(objDS.Fields("SNDDAYCS1")) = False Then .SNDDAYCS1 = objDS.Fields("SNDDAYCS1")                    '' ���M���t
                If IsNull(objDS.Fields("SNDKDWHCS1")) = False Then .SNDKDWHCS1 = objDS.Fields("SNDKDWHCS1")                 '' ���M�t���O(DWH)
                If IsNull(objDS.Fields("SDAYDWHCS1")) = False Then .SDAYDWHCS1 = objDS.Fields("SDAYDWHCS1")                 '' ���M���t(DWH)
                If IsNull(objDS.Fields("SNDKSPCCS1")) = False Then .SNDKSPCCS1 = objDS.Fields("SNDKSPCCS1")                 '' ���M�t���O(SPC)
                If IsNull(objDS.Fields("SDAYSPCCS1")) = False Then .SDAYSPCCS1 = objDS.Fields("SDAYSPCCS1")                 '' ���M���t(SPC)
                If IsNull(objDS.Fields("SAKJCS1")) = False Then .SAKJCS1 = objDS.Fields("SAKJCS1")                          '' �폜�敪
            End With
            objDS.MoveNext
        Loop
    End If
    
    objDS.Close

    GetXSDCS_1 = True

proc_exit:
    '' �I��
    Exit Function

proc_err:
    Call MsgOut(100, "", ERR_DISP_LOG, "XSDCS_1")
    GetXSDCS_1 = False
    Resume proc_exit
End Function

'�T�v      :�Y������ں��ޗL��������
'���Ұ��@�@:�u���b�NID
'      �@�@:�߂�l       ,O  ,Boolean        �@,TRUE:�L/ FALSE:���i�ُ�j
'����      :��ۯ�ID��XSDC_1������
'����      :2011/02/28�@�쐬 SMPK H.Ohkubo
Public Function CheckUniqueRecordXSDCS_1(sBlock As String) As Boolean
    Dim objDS       As Object
    Dim sSQL        As String
    Dim lRecCnt     As Long
    
    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_XSDCS_1_SQL.bas -- Function CheckUniqueRecordXSDCS_1"
    
    sSQL = "select count(*) CNT"
    sSQL = sSQL & " from XSDCS_1"
    sSQL = sSQL & " where CRYNUMCS1 like '" & Trim(sBlock) & "%'"
    
    ''�f�[�^�𒊏o����
#If SRC_200_FLG = 1 Then
    If DynSet(objDS, sSQL) = False Then
        Call MsgOut(100, sSQL, ERR_DISP_LOG)
        CheckUniqueRecordXSDCS_1 = False
        Exit Function
    End If
#Else
    If DynSet2(objDS, sSQL) = False Then
        Call MsgOut(100, sSQL, ERR_DISP_LOG)
        CheckUniqueRecordXSDCS_1 = False
        Exit Function
    End If
#End If
    
    lRecCnt = 0

    ''���o���ʂ��i�[����
    If objDS.EOF = False Then
        lRecCnt = objDS.Fields("CNT")
    End If
    
    objDS.Close

    If lRecCnt > 0 Then
        CheckUniqueRecordXSDCS_1 = True
    Else
        CheckUniqueRecordXSDCS_1 = False
    End If

proc_exit:
    '' �I��
    Exit Function

proc_err:
    Call MsgOut(100, "", ERR_DISP_LOG, "XSDCS_1")
    CheckUniqueRecordXSDCS_1 = False
    Resume proc_exit
End Function

'�T�v      :�e�[�u���uXSDCS_1�v�Ƀ��R�[�h��}������
'���Ұ��@�@:�ϐ���       ,IO ,�^                ,����
'      �@�@:pXSDCS_1�@   ,I  ,typ_XSDCS_1       ,XSDCS_1�X�V�p�ް�
'      �@�@:sErrMsg�@�@�@,O  ,String         �@ ,�G���[���b�Z�[�W
'      �@�@:�߂�l       ,O  ,Boolean�@,�������݂̐���
Public Function InsertXSDCS_1(pXSDCS_1 As typ_XSDCS_1) As Boolean

    Dim sSQL    As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    
    '' ��INSERT
    sSQL = ""
    sSQL = sSQL & " INSERT INTO XSDCS_1"
    sSQL = sSQL & " ("
    sSQL = sSQL & "  CRYNUMCS1"           '' �u���b�NID
    sSQL = sSQL & ", XTALCS1"             '' �����ԍ�
    sSQL = sSQL & ", INPOSCS1"            '' �������ʒu
    sSQL = sSQL & ", HINBCS1"             '' �i��
    sSQL = sSQL & ", REVNUMCS1"           '' �i�Ԑ��i�ԍ������ԍ�
    sSQL = sSQL & ", FACTORYCS1"          '' �i�ԍH��
    sSQL = sSQL & ", OPECS1"              '' �i�ԑ��Ə���
    sSQL = sSQL & ", TRANCNTFRSCS1"       '' ������(FRS)
    sSQL = sSQL & ", CRYINDOIFRSCS1"      '' ���FLG(FRS)
    sSQL = sSQL & ", CRYRESOIFRSCS1"      '' ����FLG(FRS)
    sSQL = sSQL & ", RPCRYNUMCS1"         '' �e�u���b�NID
    sSQL = sSQL & ", LIVKCS1"             '' �����敪
    sSQL = sSQL & ", TSTAFFCS1"           '' �o�^�Ј�ID
    sSQL = sSQL & ", TDAYCS1"             '' �o�^���t
    sSQL = sSQL & " )VALUES"
    sSQL = sSQL & " ("
    
    With pXSDCS_1
        sSQL = sSQL & " " & Cnv2String2(.CRYNUMCS1) & ""                        '' �u���b�NID
        sSQL = sSQL & "," & Cnv2String2(.XTALCS1) & ""                          '' �����ԍ�
        sSQL = sSQL & "," & Cnv2Number(.INPOSCS1) & ""                          '' �������ʒu
        sSQL = sSQL & "," & Cnv2String2(.HINBCS1) & ""                          '' �i��
        sSQL = sSQL & "," & Cnv2Number(.REVNUMCS1) & ""                         '' �i�Ԑ��i�ԍ������ԍ�
        sSQL = sSQL & "," & Cnv2String2(.FACTORYCS1) & ""                       '' �i�ԍH��
        sSQL = sSQL & "," & Cnv2String2(.OPECS1) & ""                           '' �i�ԑ��Ə���
        sSQL = sSQL & "," & Cnv2Number(.TRANCNTFRSCS1) & ""                     '' ������(FRS)
        sSQL = sSQL & "," & Cnv2String2(.CRYINDOIFRSCS1) & ""                   '' ���FLG(FRS)
        sSQL = sSQL & "," & Cnv2String2(.CRYRESOIFRSCS1) & ""                   '' ����FLG(FRS)
        sSQL = sSQL & "," & Cnv2String2(.RPCRYNUMCS1) & ""                      '' �e�u���b�NID
        sSQL = sSQL & "," & Cnv2String2(.LIVKCS1) & ""                          '' �����敪
        sSQL = sSQL & "," & Cnv2String2(.TSTAFFCS1) & ""                        '' �o�^�Ј�ID
        sSQL = sSQL & ",to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')"   '' �o�^���t
        sSQL = sSQL & " )"
    End With
    
'    Debug.Print sSql
    
#If SRC_200_FLG = 1 Then
    If SqlExec(sSQL) <> 1 Then
        Call MsgOut(100, sSQL, ERR_DISP_LOG, "XSDCS_1")
        InsertXSDCS_1 = False
        Exit Function
    End If
#Else
    If SqlExec2(sSQL) <> 1 Then
        Call MsgOut(0, "DB�o�^���s" & sSQL, ERR_DISP_LOG, "CRYNUMCS1")
        InsertXSDCS_1 = False
        Exit Function
    End If
#End If
    
    InsertXSDCS_1 = True

proc_exit:
    '' �I��
    Exit Function

proc_err:
    Call MsgOut(0, "DB�o�^���s" & sSQL, ERR_DISP_LOG, "CRYNUMCS1")
    InsertXSDCS_1 = False
    Resume proc_exit

End Function

'���X�V���ڂ��\���̂ɃZ�b�g���Ĉ����n��

'�T�v      :�e�[�u���uXSDCS_1�v���X�V����
'���Ұ�    :�ϐ���        ,IO  ,�^              ,����
'          :sBlock        ,I   ,String          ,�u���b�NID
'          :records()     ,O   ,typ_XSDCS_1     ,�X�V���R�[�h
'          :�߂�l        ,O   ,Boolean         ,���o�̐���
'����      :
Public Function UpdateXSDCS_1(sBlock As String, records As typ_XSDCS_1) As Boolean
    Dim sql     As String       'SQL�S��
    
    On Error GoTo proc_err

    With records
        sql = ""
        sql = sql & "UPDATE XSDCS_1 "
        sql = sql & "SET "
        
        If .XTALCS1 <> "" And left(.XTALCS1, 1) <> vbNullChar Then sql = sql & "  XTALCS1 = " & Cnv2String(.XTALCS1)                                ' �����ԍ�
        If .INPOSCS1 <> "" And left(.INPOSCS1, 1) <> vbNullChar Then sql = sql & ", INPOSCS1 = " & Cnv2Number(.INPOSCS1)                            ' �������ʒu
        If .HINBCS1 <> "" And left(.HINBCS1, 1) <> vbNullChar Then sql = sql & ", HINBCS1 = " & Cnv2String(.HINBCS1)                                ' �i��
        If .REVNUMCS1 <> "" And left(.REVNUMCS1, 1) <> vbNullChar Then sql = sql & ", REVNUMCS1 = " & Cnv2Number(.REVNUMCS1)                        ' �i�Ԑ��i�ԍ������ԍ�
        If .FACTORYCS1 <> "" And left(.FACTORYCS1, 1) <> vbNullChar Then sql = sql & ", FACTORYCS1 = " & Cnv2String(.FACTORYCS1)                    ' �i�ԍH��
        If .OPECS1 <> "" And left(.OPECS1, 1) <> vbNullChar Then sql = sql & ", OPECS1 = " & Cnv2String(.OPECS1)                                    ' �i�ԑ��Ə���
        If .TRANCNTFRSCS1 <> "" And left(.TRANCNTFRSCS1, 1) <> vbNullChar Then sql = sql & ", TRANCNTFRSCS1 = " & Cnv2Number(.TRANCNTFRSCS1)        ' ������(FRS)
        If .CRYINDOIFRSCS1 <> "" And left(.CRYINDOIFRSCS1, 1) <> vbNullChar Then sql = sql & ", CRYINDOIFRSCS1 = " & Cnv2String(.CRYINDOIFRSCS1)    ' ���FLG(FRS)
        If .CRYRESOIFRSCS1 <> "" And left(.CRYRESOIFRSCS1, 1) <> vbNullChar Then sql = sql & ", CRYRESOIFRSCS1 = " & Cnv2String(.CRYRESOIFRSCS1)    ' ����FLG(FRS)
        
        If .RPCRYNUMCS1 <> sBlock Then
            If .RPCRYNUMCS1 <> "" And left(.RPCRYNUMCS1, 1) <> vbNullChar Then sql = sql & ", RPCRYNUMCS1 = " & Cnv2String(.RPCRYNUMCS1)            ' �e�u���b�NID
        End If
        
        If .LIVKCS1 <> "" And left(.LIVKCS1, 1) <> vbNullChar Then sql = sql & ", LIVKCS1 = " & Cnv2String(.LIVKCS1)                                ' �����敪
        If .KSTAFFCS1 <> "" And left(.KSTAFFCS1, 1) <> vbNullChar Then sql = sql & ", KSTAFFCS1 = " & Cnv2String(.KSTAFFCS1)                        ' �X�V�Ј�ID
        sql = sql & ", KDAYCS1 = to_date(" & Cnv2String(gsSysdate) & ",'yyyy/mm/dd hh24:mi:ss')"                                                    ' �X�V���t
        
        sql = sql & " where CRYNUMCS1 = '" & sBlock & "'"
    
    End With
'    Debug.Print sql

#If SRC_200_FLG = 1 Then
    If SqlExec(sql) < 0 Then
        Call MsgOut(100, sql, ERR_DISP_LOG, "XSDCS_1")
        UpdateXSDCS_1 = False
        Exit Function
    End If
#Else
    If SqlExec2(sql) < 0 Then
        Call MsgOut(100, sql, ERR_DISP_LOG, "XSDCS_1")
        UpdateXSDCS_1 = False
        Exit Function
    End If
#End If

    UpdateXSDCS_1 = True

proc_exit:
    '' �I��
    Exit Function

proc_err:
    Call MsgOut(0, "DB�X�V���s", ERR_DISP_LOG, "CRYNUMCS1")
    UpdateXSDCS_1 = False
    Resume proc_exit

End Function

'�T�v      :�e�[�u���uXSDCS_1�v�������b�g�ɂ���
'���Ұ�    :�ϐ���        ,IO  ,�^              ,����
'          :sBlock        ,I   ,String          ,�u���b�NID
'          :sStaff        ,I   ,String          ,�S����ID
'          :�߂�l        ,O   ,Boolean         ,���o�̐���
'����      :
Public Function UpdateXSDCS_1Delete(sBlock As String, sStaff As String) As Boolean
    Dim sql         As String       'SQL�S��
    
    On Error GoTo proc_err

    '' ���Y�����R�[�h�����擾���A�����ꍇ�A�X�V�����I��
    If CheckUniqueRecordXSDCS_1(sBlock) = False Then
        '' �X�V�����I��
        UpdateXSDCS_1Delete = True
        Exit Function
    End If
    
    '' ��UPDATE
    sql = ""
    sql = sql & "UPDATE XSDCS_1 "
    sql = sql & "SET "
    
    sql = sql & "  LIVKCS1 = '1'"                                                               ' �����敪
    sql = sql & ", KSTAFFCS1 = " & Cnv2String(sStaff)                                           ' �X�V�Ј�ID
    sql = sql & ", KDAYCS1 = to_date(" & Cnv2String(gsSysdate) & ",'yyyy/mm/dd hh24:mi:ss')"    ' �X�V���t
    sql = sql & " where CRYNUMCS1 = '" & sBlock & "'"

'    Debug.Print sql

#If SRC_200_FLG = 1 Then
    If SqlExec(sql) < 0 Then
        Call MsgOut(100, sql, ERR_DISP_LOG, "XSDCS_1")
        UpdateXSDCS_1Delete = False
        Exit Function
    End If
#Else
    If SqlExec2(sql) < 0 Then
        Call MsgOut(100, sql, ERR_DISP_LOG, "XSDCS_1")
        UpdateXSDCS_1Delete = False
        Exit Function
    End If
#End If

    UpdateXSDCS_1Delete = True

proc_exit:
    '' �I��
    Exit Function

proc_err:
    Call MsgOut(0, "DB�X�V���s", ERR_DISP_LOG, "CRYNUMCS1")
    UpdateXSDCS_1Delete = False
    Resume proc_exit

End Function

'�T�v      :�e�[�u���uXSDCS_1�v�o�^�E�X�V
'���Ұ�    :�ϐ���        ,IO   ,�^             ,����
'          :sBlockID      ,I    ,String         ,�u���b�NID
'          :sCut          ,I    ,String         ,CUT�ʒu
'          :sHin          ,I    ,String         ,�i��
'          :sRev          ,I    ,String         ,���r�W����
'          :sTranCnt      ,I    ,String         ,������
'          :sFRSKbn       ,I    ,String         ,FRS���FLG
'          :sResult       ,I    ,String         ,FRS����FLG
'          :sStaff        ,I    ,String         ,�S����ID
'          :sXtal         ,I    ,String         ,�����ԍ�
'          :sRpBlockId    ,I    ,String         ,�e�u���b�NID
'          :�߂�l        ,O    ,Boolean        ,���s�̐���
'����      :
Public Function CreateOrUpdateXSDCS_1(sBlockId As String _
                                    , sCut As String _
                                    , sHin As String _
                                    , sRev As String _
                                    , sTranCnt As String _
                                    , sFRSKbn As String _
                                    , sResult As String _
                                    , sStaff As String _
                                    , sXtal As String _
                                    , sRpBlockId As String _
                                    ) As Boolean
    Dim tXSDCS_1        As typ_XSDCS_1
    
    CreateOrUpdateXSDCS_1 = False
    
    Call LogInit
    Call GetSysdate
    
    tXSDCS_1.CRYNUMCS1 = sBlockId                           '' �u���b�NID
    tXSDCS_1.XTALCS1 = sXtal                                '' �����ԍ�
    tXSDCS_1.INPOSCS1 = sCut                                '' �������ʒu
    tXSDCS_1.HINBCS1 = sHin                                 '' �i��
    tXSDCS_1.REVNUMCS1 = left(sRev, 2)                      '' �i�Ԑ��i�ԍ������ԍ�
    tXSDCS_1.FACTORYCS1 = Mid(sRev, 3, 1)                   '' �i�ԍH��
    tXSDCS_1.OPECS1 = Right(sRev, 1)                        '' �i�ԑ��Ə���
    tXSDCS_1.TRANCNTFRSCS1 = sTranCnt                       '' ������(FRS)
    tXSDCS_1.CRYINDOIFRSCS1 = sFRSKbn                       '' ���FLG(FRS)
    tXSDCS_1.CRYRESOIFRSCS1 = sResult                       '' ����FLG(FRS)
    tXSDCS_1.RPCRYNUMCS1 = sRpBlockId                       '' �e�u���b�NID
    tXSDCS_1.LIVKCS1 = "0"                                  '' �����敪
    tXSDCS_1.TSTAFFCS1 = sStaff                             '' �o�^�Ј�ID
    
    '' �Ώۃu���b�NID�f�[�^���݃`�F�b�N
    If CheckUniqueRecordXSDCS_1(sBlockId) = False Then
        
        '' �Ώۃu���b�NID��XSDCS_1�����݂��Ȃ��ꍇ�A�ǉ��i�ؒf�L��j
        
        '' ��DB�o�^
        If InsertXSDCS_1(tXSDCS_1) = False Then
            Exit Function
        End If
        
'' Chg Start 2011/05/10 SMPK H.Ohkubo
        '' ���e�u���b�NID��XSDCS_1.LIVKCS1="1"
''        If UpdateXSDCS_1Delete(sRpBlockId, sStaff) = False Then
''            Exit Function
''        End If
        If sBlockId <> sRpBlockId Then
            If UpdateXSDCS_1Delete(sRpBlockId, sStaff) = False Then
                Exit Function
            End If
        End If
'' Chg End 2011/05/10 SMPK H.Ohkubo

    Else
        '' �Ώۃu���b�NID��XSDCS_1�����݂���ꍇ�A�X�V�i�ؒf�����j
        
        '' ��DB�X�V
        If UpdateXSDCS_1(sBlockId, tXSDCS_1) = False Then
            Exit Function
        End If
        
    End If
    
    CreateOrUpdateXSDCS_1 = True

End Function

' @(f)
' �@�\      : SQL������ϊ��֐�
'
' �Ԃ�l    : '<���͕�����>' or NULL
'
' ������    : �ϊ��Ώە�����
'
' �@�\����  : ������null�ł����"NULL"���A�����łȂ���΃V���O���R�[�e�[�V���������ďo�͂���
Private Function Cnv2String(vinput) As String
    If IsNull(vinput) Or vinput = "NULL" Then
        vinput = ""
    End If
    
    If vinput = "" Then
        Cnv2String = "NULL"
    Else
        Cnv2String = SQRT & vinput & SQRT
    End If
    
End Function

' @(f)
' �@�\      : SQL������ϊ��֐�
'
' �Ԃ�l    : '<���͕�����>' or NULL
'
' ������    : �ϊ��Ώە�����
'
' �@�\����  : ������null�ł����"NULL"���A�����łȂ���΃V���O���R�[�e�[�V���������ďo�͂���
Private Function Cnv2String2(vinput) As String
    If IsNull(vinput) Or vinput = "NULL" Then
        vinput = ""
    End If
    
    If vinput = "" Then
        Cnv2String2 = SQRT & " " & SQRT
    Else
        Cnv2String2 = SQRT & vinput & SQRT
    End If
    
End Function

' @(f)
' �@�\      : SQL���l�ϊ��֐�
'
' �Ԃ�l    : <���͐��l> or NULL
'
' ������    : �ϊ��Ώې��l
'
' �@�\����  : �n���ꂽ���l��NULL�ł����"NULL"�������łȂ���΂��̂܂܏o�͂���
Private Function Cnv2Number(vinput) As String
    If IsNull(vinput) Or vinput = "NULL" Then
        vinput = ""
    End If
    
    If vinput = "" Then
        Cnv2Number = "NULL"
    Else
        Cnv2Number = vinput
    End If
End Function

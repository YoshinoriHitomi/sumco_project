Attribute VB_Name = "s_cmbc053"
'�T�v      :�d�o�c���т�o�^����
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :tblEPD        ,I  ,typ_TBCMJ001     ,�d�o�c���уe�[�u��
'          :�߂�l        ,O  ,Integer          ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Public Function InsertTbl_EPD(tblEPD As typ_TBCMJ001) As Integer
    Dim iRet       As Integer
    Dim tblTarget  As typ_cmjc001i_Disp
    
    InsertTbl_EPD = FUNCTION_RETURN_FAILURE

    '' �f�[�^�`���̕ϊ�
    ConvDate_F_cmjc001i_a tblEPD, tblTarget, True
    ''�d�o�c���тɓo�^
    iRet = DBDRV_Getcmjc001i_Exec(tblTarget, tblEPD.CRYNUM, tblEPD.TSTAFFID)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function

    InsertTbl_EPD = FUNCTION_RETURN_SUCCESS
End Function

'�T�v      :�d�o�c���т��擾����
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :tblEPD        ,O   ,typ_TBCMJ001      ,�d�o�c���уe�[�u��
'          :strCryNum     ,I   ,String           ,�����ԍ�
'          :iSmpNo        ,I   ,Long             ,�T���v��No.   Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
'          :iIngotPos     ,I   ,Integer          ,�������ʒu
'          :strSmpKbn     ,I   ,String           ,�T���v���敪
'          :�߂�l        ,O   ,Integer          ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Public Function GetResultsEPD(tblEPD As typ_TBCMJ001, strCryNum As String, iSmpNo As Long, iIngotpos As Integer, strSmpKbn As String) As Integer
    Dim iRet        As Integer
    Dim tblGetEPD() As typ_TBCMJ001

    GetResultsEPD = FUNCTION_RETURN_FAILURE

    '' �d�o�c���т̎擾
    iRet = DBDRV_GetTBCMJ001(tblGetEPD, _
             "A where CRYNUM='" & strCryNum & "' and POSITION=" & iIngotpos & _
             " and TRANCNT=any(select max(TRANCNT) from TBCMJ001 where CRYNUM='" & strCryNum & "' and POSITION=" & iIngotpos & _
             " and SMPKBN=A.SMPKBN" & _
             ")", _
             " order by POSITION, SMPKBN" & IIf(strSmpKbn = "B", "", " desc"))
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGetEPD) = 0 Then Exit Function

    tblEPD = tblGetEPD(1)
    
    GetResultsEPD = FUNCTION_RETURN_SUCCESS

End Function
'Akizuki <<<<<XSDC1��֏����Ƃ��č쐬��>>>>>
'�T�v      :�����ԍ����i�ԊǗ��e�[�u�����擾����B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :tblHinban()   ,O   ,typ_TBCME041 ,�i�ԊǗ��e�[�u��
'          :strCryNum     ,I   ,String           ,�����ԍ�
'          :�߂�l        ,O  ,Integer          ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Public Function GetHinban_X(tblHinban() As typ_TBCME041, strCryNum As String) As Integer
    Dim iRet        As Integer
    Dim tblGet()    As typ_TBCME041
    Dim Index       As Integer
    Dim tblPlup     As typ_TBCMH004

    GetHinban = FUNCTION_RETURN_FAILURE

    '' �i�ԊǗ��e�[�u����������
    RemoveAll_HinbanManage tblHinban

    '' �i�ԊǗ��e�[�u���̎擾
    iRet = DBDRV_GetTBCME041(tblGet, "where CRYNUM='" & strCryNum & "' ", "order by INGOTPOS")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then
        If Len(strCryNum) <> 12 Then Exit Function
        '' �i�ԊǗ��e�[�u���O�̏ꍇ�A�������㎸�s���Ă�����̂Ƃ���
        '' �����́A�y�i�Ԃ̌����Ƃ���
        '' ���㒷���擾
        If GetPlupEndRslt(tblPlup, strCryNum) <> FUNCTION_RETURN_SUCCESS Then Exit Function
        ReDim tblGet(1)
        With tblGet(1)
            .CRYNUM = strCryNum
            .INGOTPOS = 0
            .hinban = "Z"
            .REVNUM = 0
            .Factory = vbNullString
            .OpeCond = vbNullString
            .Length = tblPlup.LENGFREE
        End With
    End If

    For Index = 1 To UBound(tblGet)
        If Add_HinbanManage(tblHinban, tblGet(Index)) <> FUNCTION_RETURN_SUCCESS Then
            Exit Function
        End If
    Next Index

    If UBound(tblHinban) <= 0 Then
        Exit Function
    End If

    GetHinban = FUNCTION_RETURN_SUCCESS

End Function

'�T�v      :�i�Ԃ�茋�������Ǘ����擾����B
'���Ұ�    :�ϐ���        ,IO ,�^                           ,����
'          :tblData        ,O  ,typ_TBCME036               ,���������Ǘ��e�[�u��
'          :tHinInf       ,I  ,tFullHinban                 ,�i��
'          :�߂�l        ,O  ,Integer                      ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Public Function GetSXLInsideSpecManager(tblData As typ_TBCME036, tHinInf As tFullHinban) As Integer
    Dim Index       As Long
    Dim iRet        As Integer
    Dim tblGet()    As typ_TBCME036
    
    GetSXLInsideSpecManager = FUNCTION_RETURN_FAILURE

    '' ���������Ǘ��̎擾
'    iRet = DBDRV_GetTBCME036(tblGet, _
'             "where HINBAN='" & tHinInf.hinban & "' and MNOREVNO=" & tHinInf.mnorevno & " and FACTORY='" & tHinInf.factory & "' and OPECOND='" & tHinInf.opecond & "'")
    '06/04/11 ooba
    iRet = DBDRV_GetTBCME036_cmbc028(tblGet, _
             "where HINBAN='" & tHinInf.hinban & "' and MNOREVNO=" & tHinInf.mnorevno & " and FACTORY='" & tHinInf.Factory & "' and OPECOND='" & tHinInf.OpeCond & "'")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function

    tblData = tblGet(1)
    
    GetSXLInsideSpecManager = FUNCTION_RETURN_SUCCESS

End Function
Public Function Add_EPDRslt(tblTarget() As typ_TBCMJ001, tblDat As typ_TBCMJ001, Optional Index As Long = -1) As Integer
    Dim tblIndex As Long
    
    Add_EPDRslt = FUNCTION_RETURN_FAILURE

    '' �f�[�^�̒ǉ��E�X�V�`�F�b�N
    If Index > -1 Then
        '' �f�[�^�X�V�̏ꍇ
        tblIndex = Index
        If Index > UBound(tblTarget) Then
            '' �X�V�f�[�^�ʒu�C���f�b�N�X�͈͂������̏ꍇ�A�G���[�I��
            Exit Function
        End If
    Else
        '' �f�[�^�ǉ��̏ꍇ
        '' �e�[�u���f�[�^�i�[�̈�g��
        ReDim Preserve tblTarget(UBound(tblTarget) + 1)
        '' �e�[�u���f�[�^�����擾
        tblIndex = UBound(tblTarget) - 1
    End If

    '' �f�[�^�ǉ�
    tblTarget(tblIndex) = tblDat

    Add_EPDRslt = FUNCTION_RETURN_SUCCESS
End Function

'�T�v      :�f�[�^�ϊ����s��
'���Ұ�    :�ϐ���        ,IO ,�^                ,����
'          :tblLeft       ,IO   ,typ_TBCMJ001      ,�e�[�u���f�[�^�P
'          :tblRight      ,IO   ,typ_cmjc001i_Disp ,�e�[�u���f�[�^�Q
'          :bFlg          ,I   ,Boolean           ,TRUE:�����P�f�[�^�������Q�f�[�^�ւ̕ϊ�  FALSE:�����P�f�[�^�������Q�f�[�^�ւ̕ϊ�
'����      :
Public Sub ConvDate_F_cmjc001i_a(tblLeft As typ_TBCMJ001, tblRight As typ_cmjc001i_Disp, bFlg As Boolean)
    If bFlg = True Then
        With tblRight
            .POSITION = tblLeft.POSITION
            .SMPKBN = tblLeft.SMPKBN
            .TRANCOND = tblLeft.TRANCOND
            .SMPLNO = tblLeft.SMPLNO
            .SMPLUMU = tblLeft.SMPLUMU
            .KRPROCCD = tblLeft.KRPROCCD
            .PROCCODE = tblLeft.PROCCODE
            .hinban = tblLeft.hinban
            .REVNUM = tblLeft.REVNUM
            .Factory = tblLeft.Factory
            .OpeCond = tblLeft.OpeCond
            .GOUKI = tblLeft.GOUKI
            .MEASURE = tblLeft.MEASURE
        End With
    Else
        With tblLeft
            .POSITION = tblRight.POSITION
            .SMPKBN = tblRight.SMPKBN
            .TRANCOND = tblRight.TRANCOND
            .SMPLNO = tblRight.SMPLNO
            .SMPLUMU = tblRight.SMPLUMU
            .KRPROCCD = tblRight.KRPROCCD
            .PROCCODE = tblRight.PROCCODE
            .hinban = tblRight.hinban
            .REVNUM = tblRight.REVNUM
            .Factory = tblRight.Factory
            .OpeCond = tblRight.OpeCond
            .GOUKI = tblRight.GOUKI
            .MEASURE = tblRight.MEASURE
        End With
    End If

End Sub

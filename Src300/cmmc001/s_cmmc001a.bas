Attribute VB_Name = "s_cmmc001a"
''
'' ��R�ΐ͌v�Z��ʕW�����W���[��
''



'�T�v      :���������擾����
'���Ұ�    :�ϐ���        ,IO ,�^             ,����
'          :tblTarget     ,I   ,typ_TBCME037  ,�������e�[�u��
'          :strCryNum     ,I   ,String        ,�����ԍ�
'          :�߂�l        ,O   ,FUNCTION_RETURN       ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :���������A�˂炢�i�Ԃ��擾����
Public Function GetRpHinban(tblTarget As typ_TBCME037, strCryNum As String) As FUNCTION_RETURN
    Dim iRet        As Integer
    Dim tblGet()    As typ_TBCME037
    
    GetRpHinban = FUNCTION_RETURN_FAILURE
    
    '' ���������擾����
    iRet = DBDRV_GetTBCME037(tblGet, "where CRYNUM='" & strCryNum & "'")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function

    tblTarget = tblGet(1)

    GetRpHinban = FUNCTION_RETURN_SUCCESS
End Function


'�T�v      :���グ�I�����т��擾����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :tblTarget     ,I   ,typ_TBCMH004 ,���グ�I�����уe�[�u��
'          :strCryNum     ,I   ,String       ,�����ԍ�
'          :�߂�l        ,O   ,FUNCTION_RETURN       ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Public Function GetPlupEndRslt(tblTarget As typ_TBCMH004, strCryNum As String) As FUNCTION_RETURN
    Dim iRet        As Integer
    Dim tblGet()    As typ_TBCMH004
    Dim strCry9     As String * 1
    Dim strCryWork  As String
    
    GetPlupEndRslt = FUNCTION_RETURN_FAILURE

    '' ���グ�I�����т̎擾
    iRet = DBDRV_GetTBCMH004(tblGet, "where CRYNUM='" & strCryNum & "'")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function
    tblTarget = tblGet(1)

    '' �����ԍ��X���ڂ��擾�i�c�ʈ��������̊m�F�j
    strCry9 = Mid(strCryNum, 9, 1)
    If (strCry9 <> "") And (InStr(REST_WT_CRYCODE, strCry9) <> 0) Then
        If strCry9 <> "A" Then
            strCryWork = Left(strCryNum, 8) + "A" + Right(strCryNum, 3)
            iRet = DBDRV_GetTBCMH004(tblGet, "where CRYNUM='" & strCryWork & "'")
            If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
            If UBound(tblGet) = 0 Then Exit Function
            tblTarget.CHARGE = tblGet(1).CHARGE
        End If
    End If

    GetPlupEndRslt = FUNCTION_RETURN_SUCCESS
End Function


'�T�v      :��R���т��擾����
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :tblRs()       ,O   ,typ_TBCMJ002     ,��R���уe�[�u���z��(1�`)
'          :strCryNum     ,I   ,String           ,�����ԍ�
'          :�߂�l        ,O   ,FUNCTION_RETURN   ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Public Function GetResultsRs(tblRs() As typ_TBCMJ002, strCryNum As String) As FUNCTION_RETURN
    Dim iRet            As Integer

    GetResultsRs = FUNCTION_RETURN_FAILURE

    '' �����ԍ�����R���т��擾����i�ʒu�Ń\�[�g�j
    iRet = DBDRV_GetTBCMJ002(tblRs, _
             " A where CRYNUM='" & strCryNum & "'" & _
             " and TRANCNT=any(select max(TRANCNT) from TBCMJ002" & _
             " where CRYNUM='" & strCryNum & "' and POSITION=A.POSITION)" & _
             " order by POSITION")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblRs) = 0 Then Exit Function

    GetResultsRs = FUNCTION_RETURN_SUCCESS

End Function


'�T�v      :�i�Ԃ�萻�i�d�l�r�w�k�f�[�^�P���擾�A�����āA�擾�������i�d�l�f�[�^��ǉ�����B
'���Ұ�    :�ϐ���        ,IO ,�^                   ,����
'          :tblTarget     ,O   ,typ_TBCME018        ,���i�d�l�r�w�k�f�[�^�P�e�[�u��
'          :tHinInf       ,I   ,tFullHinban         ,�i��
'          :�߂�l        ,O  ,Integer              ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Public Function GetSPSXLData1(tblTarget As typ_TBCME018, tHinInf As tFullHinban) As Integer
    Dim Index       As Long
    Dim iRet        As Integer
    Dim tblGet()    As typ_TBCME018
    
    GetSPSXLData1 = FUNCTION_RETURN_FAILURE

    '' ���i�d�l�r�w�k�f�[�^�P�̎擾
    iRet = DBDRV_GetTBCME018(tblGet, "where HINBAN='" & tHinInf.HINBAN & "' and MNOREVNO=" & tHinInf.mnorevno & " and FACTORY='" & tHinInf.factory & "' and OPECOND='" & tHinInf.opecond & "'")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function
    
    tblTarget = tblGet(1)
    
    GetSPSXLData1 = FUNCTION_RETURN_SUCCESS

End Function


'�T�v      :��R�̒l��\���p�ɕ����񉻂���(�w��̏����_�ȉ�����)
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :rs            ,I  ,Double    ,��R�l
'          :place         ,I  ,Integer   ,�����_�ȉ�����
'����      :��R�l�\�������𓝈ꂷ�邽�߁B<0�̂Ƃ��͋󕶎����Ԃ�
'����      :2002/1/16 �쐬  �쑺 (2002/07 s_cmzc020a.bas���ړ�)
'����      :2002/1/17 S.Sano
Public Function toRsStrByPlace(rs As Double, place As Integer) As String
Dim s$

    If rs < 0 Then
        s = vbNullString
'2002/01/17 S.Sano    ElseIf rs >= 99999.9 Then
'2002/01/17 S.Sano        s = "99999.9"
    Else
        s = Format$(rs, "0." & String(place, "0"))
        If Val(s) >= 100000 Then
            s = "99999." & String(place, "9")
        End If
    End If
    toRsStrByPlace = s
End Function

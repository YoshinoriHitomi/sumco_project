Attribute VB_Name = "s_MaeBatGet"

'=================================================================================
'�T�v      :�O�o�b�`�������擾����
'���Ұ��@�@:�ϐ���          , IO , �^              , ����
'          :CRYNUM          , I  ,String           , �����ԍ�
'      �@�@:�߂�l          , O  ,Double�@         , �O�o�b�`������
'����      :
'����      :2011/09/27 Marushita
Public Function GetMaeBatLen(ByVal CRYNUM As String, Optional ByVal iKata As Integer = 0) As Double
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    
    Dim dBatLen     As Double               '�o�b�`������
    Dim p_CRYNUM    As String               '�O�����ԍ��擾�p
    Dim p_wgtTop    As Double               '�O������擾�p(�OTop���)
    Dim p_DM        As Double               '�O������擾�p(�O���a����)
    Dim p_wgtTA     As Double               '�O������擾�p(�O�e�C���d��)
    Dim p_LENTK     As Long                 '�O������擾�p(�O���㒷)
    Dim p_wgtKata   As Double               '�O������擾�p(�O���d��)
    Dim n_wgtTop    As Double               '���݈�����擾�p(Top���)
    Dim n_DM        As Double               '���݈�����擾�p(���a����)
    Dim n_wgtTA     As Double               '���݈�����擾�p(�e�C���d��)
    Dim n_LENTK     As Long                 '���݈�����擾�p(���㒷)
    Dim n_wgtKata   As Double               '���݈�����擾�p(���d��)
    Dim iMaeBatFlg  As Integer              '�������v�ZFLG
    Dim N_CRYNUM    As String               '���݌����ԍ��Z�b�g�p
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_MaeBatLen.bas -- Function GetMaeBatLen"
    
    GetMaeBatLen = 0
        
    dBatLen = 0
    
    'CRYNUM��9���ڂ�"A"�̂Ƃ������𔲂���
    If Mid(CRYNUM, 9, 1) = "A" Then
        Exit Function
    Else
        '���݈���̌��������擾(Top��ʁA���a���ρA���㒷�A�e�C���d�ʁA���d��)
        If GetXSDC1Info(CRYNUM, n_wgtTop, n_DM, n_LENTK, n_wgtTA, n_wgtKata) = FUNCTION_RETURN_FAILURE Then
            Exit Function
        End If
    End If
    
    '���݌����ԍ��̏����Z�b�g
    N_CRYNUM = CRYNUM
        
    Do While iMaeBatFlg = 0
        '�O����̌����ԍ����擾(���݌����ԍ�)
        If GetPreCrynum(N_CRYNUM, p_CRYNUM) = FUNCTION_RETURN_FAILURE Then
            iMaeBatFlg = 1
        Else
            '�O����̌��������擾(�OTop��ʁA�O���a���ρA�O���㒷�A�O�e�C���d��)
            If GetXSDC1Info(p_CRYNUM, p_wgtTop, p_DM, p_LENTK, p_wgtTA, p_wgtKata) = FUNCTION_RETURN_FAILURE Then
                iMaeBatFlg = 1
            Else
                '�O�o�b�`�������̌v�Z
                '���d�ʍ���
                If iKata = 0 Then
                    '�O�o�b�`������ = �OTop��� / (�O�f�ʐ� * 0.00233) + �O���㒷 + ((�O�e�C���d�� + ���d��) / (�f�ʐ� * 0.00233))
                    dBatLen = dBatLen + (p_wgtTop / (AreaOfCircle(p_DM) * HIJU_SILICONE)) + p_LENTK + ((p_wgtTA + n_wgtKata) / (AreaOfCircle(n_DM) * HIJU_SILICONE))
                Else
                    '�O�o�b�`������ = �OTop��� / (�O�f�ʐ� * 0.00233) + �O���㒷 + ((�O�e�C���d�� + ���d��) / (�f�ʐ� * 0.00233))
                    dBatLen = dBatLen + (p_wgtTop / (AreaOfCircle(p_DM) * HIJU_SILICONE)) + p_LENTK + (p_wgtTA) / (AreaOfCircle(n_DM) * HIJU_SILICONE)
                    iKata = 0
                End If
                '���݌����ԍ��̃Z�b�g
                N_CRYNUM = p_CRYNUM
                '�O���d�ʁA�O�f�ʐς����݂ɃZ�b�g
                n_wgtKata = p_wgtKata
                n_DM = p_DM
            End If
        End If
    Loop
    
    GetMaeBatLen = dBatLen
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    GetMaeBatLen = 0
    gErr.HandleError
    Resume proc_exit
End Function

'=================================================================================
'�T�v      :�O�o�b�`�c�t�擾����
'���Ұ��@�@:�ϐ���          , IO , �^              , ����
'          :CRYNUM          , I  ,String           , �����ԍ�
'      �@�@:�߂�l          , O  ,Long  �@         , �O�o�b�`�c�t
'����      :
'����      :2011/09/27 Marushita
Public Function GetMaeBatZan(ByVal CRYNUM As String) As Double
    
    Dim sqlWhere As String
    
    Dim lZaneki     As Double               '�c�t��
    Dim p_CRYNUM    As String               '�O�����ԍ��擾�p
    Dim p_wgtTop    As Double               '�O������擾�p(�OBtTop���)
    Dim p_wgtTA     As Double               '�O������擾�p(�OBt�d����)
    Dim p_wgWGHTTK  As Long                 '�O������擾�p(�OBt�F��d��)
    Dim iMaeBatFlg  As Integer              '�������v�ZFLG
    Dim N_CRYNUM    As String               '���݌����ԍ��Z�b�g�p
    Dim tblXSDC1()  As typ_XSDC1            'XSDC1�f�[�^�擾�p
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_MaeBatGet.bas -- Function GetMaeBatZan"
    
    GetMaeBatZan = 0
        
    lZaneki = 0
    
    '���݌����ԍ��̏����Z�b�g
    N_CRYNUM = CRYNUM
        
    Do While iMaeBatFlg = 0
        '�O����A�Ԃ̌����ԍ����擾(���݌����ԍ�)
        If GetPreRenCrynum(N_CRYNUM, p_CRYNUM) = FUNCTION_RETURN_FAILURE Then
            iMaeBatFlg = 1
        Else
            'WHERE����
            sqlWhere = "WHERE XTALC1 = '" & p_CRYNUM & "'"
            '���R�[�h�Z�b�g�̎擾(���s������v���V�[�W�����甲����j
            If DBDRV_GetXSDC1(tblXSDC1, sqlWhere) = FUNCTION_RETURN_FAILURE Then
                iMaeBatFlg = 1
            Else
                '�O�c�t�̌v�Z
                '�O�c�t = �OBt�O�c�t + �OBt�d���� - �OBt�F��d�� - �OBtTop���
                lZaneki = lZaneki + tblXSDC1(1).PUCHAGC1 - tblXSDC1(1).PUWC1 - tblXSDC1(1).PUTCUTWC1
                'lZaneki = lZaneki + tblXSDC1(1).PUCHAGC1 - tblXSDC1(1).PUWC1 - tblXSDC1(1).WGHTTOC1
                '���݌����ԍ��̃Z�b�g
                N_CRYNUM = p_CRYNUM
            End If
        End If
    Loop
    
    GetMaeBatZan = lZaneki
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    GetMaeBatZan = 0
    gErr.HandleError
    Resume proc_exit
End Function

'=================================================================================
'�T�v       :�O����̌����ԍ����擾����B
'���Ұ�     :�ϐ���         ,IO ,�^                         ,����
'           :sCrynum        ,I  ,String             �@      ,���݌����ԍ�
'           :sP_Crynum      ,O  ,String             �@      ,�O����̌����ԍ�
'           :�߂�l         ,O  ,FUNCTION_RETURN            ,
'����       :
'����       :2011/09/27 Marushita
Public Function GetPreCrynum(ByVal sCrynum As String, ByRef sP_Crynum As String) As FUNCTION_RETURN
    
    On Error GoTo Err
    GetPreCrynum = FUNCTION_RETURN_FAILURE
                
    Dim sCrynum9 As String
    Dim iCrynum9 As Integer
            
    sP_Crynum = ""
    
    'sCrynum��9���ڂ��Z�b�g
    sCrynum9 = Mid(sCrynum, 9, 1)
    
    '"A","1"�̓G���[�ŕԂ�
    If sCrynum9 = "A" Or sCrynum9 = "1" Then
        Exit Function
    Else
        iCrynum9 = Asc(sCrynum9)
        sP_Crynum = Mid(sCrynum, 1, 8) & Chr(iCrynum9 - 1) & Mid(sCrynum, 10, 3)
    End If
    
    GetPreCrynum = FUNCTION_RETURN_SUCCESS
    
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    
End Function

'=================================================================================
'�T�v       :�w�茋���ԍ��̌��������擾����B
'���Ұ�     :�ϐ���         ,IO ,�^                         ,����
'           :sCRYNUM        ,I  ,String    ,�O���㌋���ԍ�
'           :dWgtTop        ,O  ,Double    ,�g�b�v�d�ʎ��ђl
'           :dDM            ,O  ,Double    ,���a�P�`�R�̕���
'           :lLentk         ,O  ,Long      ,���㒷
'           :dwgtTA         ,O  ,Double    ,�e�C���d�ʎ��ђl
'           :dwgtKata�@     ,O  ,Double    ,���d��
'           :�߂�l         ,O  ,FUNCTION_RETURN
'����       :
'����       :2011/10/19 Marushita�@�֐����E���d�ʍ��ڂ̒ǉ�
Public Function GetXSDC1Info(ByVal sCrynum As String, _
                            ByRef dwgtTop As Double, _
                            ByRef dDM As Double, _
                            ByRef lLenTK As Long, _
                            ByRef dwgtTA As Double, _
                            ByRef dwgtKata As Double) As FUNCTION_RETURN

    On Error GoTo proc_err
    
    Dim sql As String
    Dim rs As OraDynaset
        
    GetXSDC1Info = FUNCTION_RETURN_FAILURE
    
    sql = " SELECT WGHTTOC1, LENTKC1, WGHTTAC1, PUTCUTWC1, "
    sql = sql & " (DIA1C1 + DIA2C1 + DIA3C1) / 3.0 AS DM "
    sql = sql & " FROM XSDC1 "
    sql = sql & " WHERE XTALC1 = '" & sCrynum & "'"
        
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount > 0 Then
        dwgtTop = rs("PUTCUTWC1")           ''�d�ʁiTOP�jPUTCUTWC1�ɕύX�@2011/11/10
        'dwgtTop = rs("WGHTTOC1")           ''�d�ʁiTOP�j
        dDM = rs("DM")                     ''�������a(���ϒl)
        lLenTK = rs("LENTKC1")             ''���㒷
        dwgtTA = rs("WGHTTAC1")            ''�d�ʁiTAIL�j
        dwgtKata = rs("WGHTTOC1")         ''���d�� WGHTTOC1�ɕύX�@2011/11/10
        'dwgtKata = rs("PUTCUTWC1")         ''���d��
    End If
    rs.Close
    
    GetXSDC1Info = FUNCTION_RETURN_SUCCESS
    
    On Error GoTo 0
    Exit Function

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    GetXSDC1Info = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
    
End Function

'=================================================================================
'�T�v       :�O����A�Ԃ̌����ԍ����擾����B
'���Ұ�     :�ϐ���         ,IO ,�^                         ,����
'           :sCrynum        ,I  ,String             �@      ,���݌����ԍ�
'           :sP_Crynum      ,O  ,String             �@      ,�O����̌����ԍ�
'           :�߂�l         ,O  ,FUNCTION_RETURN            ,
'����       :
'����       :2011/10/14 Marushita
Public Function GetPreRenCrynum(ByVal sCrynum As String, ByRef sP_Crynum As String) As FUNCTION_RETURN
    
    On Error GoTo Err
    GetPreRenCrynum = FUNCTION_RETURN_FAILURE
                
    Dim sql As String
    Dim rs As OraDynaset
    
    Dim sUpGrp      As String               '����O���[�v�ԍ�
    Dim iRenban     As Integer              '�O���[�v���A��
            
    sUpGrp = ""
    sP_Crynum = ""
    iRenban = 0
    
    sql = " SELECT RENBAN,GROUPUPINDNO "
    sql = sql & " FROM XSDC1,TBCMH001 "
    sql = sql & " WHERE XTALC1 = '" & sCrynum & "'"
    sql = sql & " AND HISIJIC1 = UPINDNO "
    sql = sql & " AND LIVK <> '1' "
        
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount > 0 Then
        '�O���[�v�ԍ��E�O���[�v���A�Ԃ��擾
        sUpGrp = rs("GROUPUPINDNO")      ''�O���[�v�ԍ�
        iRenban = rs("RENBAN")           ''�A��
    End If
    rs.Close
        
    '�A�Ԃ��P�ȉ��̓G���[�ŕԂ�
    If iRenban <= 1 Then
        Exit Function
    Else
        iRenban = iRenban - 1
    End If
    
    sql = " SELECT XTALC1 "
    sql = sql & " FROM XSDC1,TBCMH001 "
    sql = sql & " WHERE GROUPUPINDNO = '" & sUpGrp & "'"
    sql = sql & " AND RENBAN = " & iRenban & " "
    sql = sql & " AND LIVK <> '1' "
    sql = sql & " AND HISIJIC1 = UPINDNO "
        
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount > 0 Then
        '�����ԍ����擾
        sP_Crynum = rs("XTALC1")           ''�����ԍ�
    End If
    rs.Close
    
    GetPreRenCrynum = FUNCTION_RETURN_SUCCESS
    
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    
End Function


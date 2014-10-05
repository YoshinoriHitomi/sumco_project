Attribute VB_Name = "mdlVBX9000"
'///////////////////////////////////////////////////
' @(S)
'       ���[�o�͈˗�����
'
' @(h)  mdlVBX9000.bas ver 0.1 ( 2000.04.20  )
'
'///////////////////////////////////////////////////
Option Explicit

Public Enum gPRINTCODE  ''  ���[�o�͎w�ߋ敪�^
    PULCRYSPRINT        ''  ����w�����o�͎w��
    BUNKATUPRINT        ''  ���H�����[�o�͎w��
End Enum
Type PRINTERINFODATA            ''  ���[�o�͏��e�[�u���Ǘ��f�[�^�i�[�^
    sPrinterName    As String   ''  �v�����^�[��
    sPrintDatName   As String   ''  �o�͒��[����
    sInfoDat01      As String   ''  ���[�o�̓L�[�f�[�^01
    sInfoDat02      As String   ''  ���[�o�̓L�[�f�[�^02
    sInfoDat03      As String   ''  ���[�o�̓L�[�f�[�^03
    sInfoDat04      As String   ''  ���[�o�̓L�[�f�[�^04
    sInfoDat05      As String   ''  ���[�o�̓L�[�f�[�^05
    lMaxCnt         As Long     ''  �ő僌�R�[�h
    lReadCnt        As Long     ''  ���[�h�|�C���^�[
    lWriteCnt       As Long     ''  ���C�g�|�C���^�[
End Type
Type KANRIKODEDAT
    sSysCode    As String       ''  �V�X�e���敪
    sShuCode    As String       ''  ��ʃR�[�h
    sCode       As String       ''  �ʃR�[�h
    sShortName  As String       ''  ����(�Z�k)
    sCodeName   As String       ''  ����
    sKanrenCode As String       ''  �֘A�R�[�h
    sDat01      As String       ''  �f�[�^�P
    sDat02      As String       ''  �f�[�^�Q
    sDat03      As String       ''  �f�[�^�R
    sDat04      As String       ''  �f�[�^�S
    sDat05      As String       ''  �f�[�^�T
    lCnt01      As Long         ''  �J�E���^�[�P
    lCnt02      As Long         ''  �J�E���^�[�Q
    lCnt03      As Long         ''  �J�E���^�[�R
    lCnt04      As Long         ''  �J�E���^�[�S
    lCnt05      As Long         ''  �J�E���^�[�T
End Type
Private Const mUNTENNISSI = "CRX1000"
Private Const mHIKIAGESIJISYO = "CRX1010"
Private Const mKAKOUKENSAHYO = "CRX1020"
Private Const mSAIKAKUDUKEUKAGAISYO = "CRX1030"

' @(f)
' �@�\      :   ���[�o�͈˗�����
'
' �Ԃ�l    :   True    -   ����
'               False   -   �ُ�
'
' ������    :   sWorkCode   -   �H��R�[�h
'               pPrintKbn   -   ���[�R�[�h
'                               �P�F�^�]����
'                               �Q�F����w����
'                               �R�F���H�����[
'                               �S�F�Ċi�t�f���w����
'
' �@�\����  :   ���[�o�͈˗������u�˗����e�[�u���v�ɏ�������
'
' ���l      :   sInfCode(1-4) - sPrintKbn(1)��    ���㌋���ԍ��A�����ԍ��A�H���A�ԁA���@
'               sInfCode(1-2) - sPrintKbn(2)��    ����w�����A���@
'               sInfCode(1)   - sPrintKbn(3)��    ���������ԍ�
'               sInfCode(1-3) - sPrintKbn(4)��    ���������ԍ��A�]�p���i�ԁA�]�p��i��
'
Public Function SetPrinterManager(sWorkCode As String, _
                                 pPrintKbn As gPRINTCODE, _
                                 Optional sInfCode1 As String = "", _
                                 Optional sInfCode2 As String = "", _
                                 Optional sInfCode3 As String = "", _
                                 Optional sInfCode4 As String = "", _
                                 Optional sInfCode5 As String = "") As Boolean
Dim KanriDat        As KANRIKODEDAT
Dim PrintDat        As PRINTERINFODATA
    
    
    SetPrinterManager = False
    
    PrintDat.sInfoDat01 = sInfCode1
    PrintDat.sInfoDat02 = sInfCode2
    PrintDat.sInfoDat03 = sInfCode3
    PrintDat.sInfoDat04 = sInfCode4
    PrintDat.sInfoDat05 = sInfCode5
    ''  �Ǘ��R�[�h�p�����[�^�[�ݒ�
    KanriDat.sSysCode = "X"
    KanriDat.sShuCode = "98"
    KanriDat.sCode = "X"
    ''  �Ǘ��R�[�h�e�[�u����������
    ''  ���[�R�[�h����v�����^�[���̂��擾����
    If SelectKanriTable(KanriDat) = False Then
        Exit Function
    End If
    ''  �v�����^�[���̂̎擾
    ''  �o�͎w�ߋ敪�ɂ�钠�[�R�[�h�̐ݒ�
    Select Case pPrintKbn
    Case PULCRYSPRINT        ''  ����w�����o�͎w��
        PrintDat.sPrintDatName = "r_cmdc001a"
        PrintDat.sPrinterName = KanriDat.sDat02
    Case BUNKATUPRINT        ''  ���H�����[�o�͎w��
        PrintDat.sPrintDatName = "r_cmdc001b"
        PrintDat.sPrinterName = KanriDat.sDat03
    End Select
    ''  �Ǘ��R�[�h�p�����[�^�[�ݒ�
    KanriDat.sSysCode = "X"
    KanriDat.sShuCode = "99"
    KanriDat.sCode = "1"
    ''  �Ǘ��R�[�h�e�[�u����������
    ''  ���[�R�[�h����Ǘ��f�[�^(�l�`�w���R�[�h���A�Ǎ��|�C���^�A�������݃|�C���^)���擾����
    If SelectKanriTable(KanriDat) = False Then
        Exit Function
    End If
    ''  �l�`�w���R�[�h�A�Ǎ��|�C���^�A�����|�C���^�擾
    PrintDat.lMaxCnt = Val(KanriDat.sDat01)
    PrintDat.lReadCnt = Val(KanriDat.sDat02)
    PrintDat.lWriteCnt = Val(KanriDat.sDat03)
    ''  �����|�C���^�[�C���N�������g
    If PrintDat.lMaxCnt < PrintDat.lWriteCnt + 1 Then
        PrintDat.lWriteCnt = 1
    Else
        PrintDat.lWriteCnt = PrintDat.lWriteCnt + 1
    End If
    ''  �Ǎ��|�C���^�[�Ɠ����ꍇ�̓G���[�Ƃ���
    If PrintDat.lReadCnt = PrintDat.lWriteCnt Then
        Call MsgOut(0, "���[�o�͖��߃G���[", ERR_LOG)
        Exit Function
    End If
    
    ''�R���s���[�^���擾   1/9 Yam
    If GetCompName() = "" Then
        ''�R���s���[�^���擾���s
        Call MsgOut(60, "", ERR_DISP_LOG)
        Debug.Print "�R���s���[�^���擾���s"
        Exit Function
    End If
    
    If UpdatePrinterInfo(PrintDat) = False Then
        Exit Function
    End If
    If UpdateKanriTable(2, CStr(PrintDat.lWriteCnt), "") = False Then
        Exit Function
    End If
    
    '2004.09.06 Y.Katabami
    '�w�������s�e�k�f�A�^�]�����Ĕ��s�e�k�f�̍X�V���s���B
    '����o�͗v�������e�k�f�� "2" ���s�ς݂Ƃ���
    If funcPrintOutFlgUpdata(pPrintKbn, sInfCode1) = False Then
        Exit Function
    End If
    
    SetPrinterManager = True
End Function

'�o�͓o�^��������TBCMH001�̊e�o�̓t���O���Q�i�o�͍ς݁j�ɕύX����
'2004.09.06 Y.Katabami
Function funcPrintOutFlgUpdata(pPrintKbn, sInfCode1) As Boolean
    Dim sSQL        As String
    Dim lRowCount   As Long
    Dim sSijiNo As String
    
    funcPrintOutFlgUpdata = False
        
    If (pPrintKbn = PULCRYSPRINT) Then        ''  ����w�����o�͎w��
        'sInfCode1�͎w���m���Ȃ̂ł��̂܂ܕR�t��
        sSQL = "UPDATE  TBCMH001  SET "
        sSQL = sSQL & " SIJISYOFLG = '2' "
        sSQL = sSQL & " WHERE UPINDNO = '" & sInfCode1 & "'"
        lRowCount = SqlExec2(sSQL)
        If lRowCount < 0 Then
            Call MsgOut(100, sSQL, ERR_LOG, "koda9")
            Exit Function
        End If
    ElseIf (pPrintKbn = BUNKATUPRINT) Then       ''  ���H�����[�o�͎w��
        'sInfCode1�̓`���[�W�m������w���m�����쐬���R�t��
        sSijiNo = Mid(sInfCode1, 1, 7) & "0" & Mid(sInfCode1, 9, 1)
        
        sSQL = "UPDATE  TBCMH001  SET "
        sSQL = sSQL & " UNTENFLG = '2' "
        sSQL = sSQL & " WHERE UPINDNO = '" & sSijiNo & "'"
        lRowCount = SqlExec2(sSQL)
        If lRowCount < 0 Then
            Call MsgOut(100, sSQL, ERR_LOG, "koda9")
            Exit Function
        End If
    End If
    
    funcPrintOutFlgUpdata = True

End Function

' @(f)
' �@�\      :   ���[�˗����X�V����
' �Ԃ�l    :   True    -   ����
'               False   -   �ُ�
' ������    :   PrintDat    -   ���[�o�͏��i�[�ϐ�
' �@�\����  :   ���[�o�͈˗������u�˗����e�[�u���v�ɏ�������
' ���l      :   �X�V����0���̏ꍇ�͓o�^���s��
'
Private Function UpdatePrinterInfo(PrintDat As PRINTERINFODATA) As Boolean
    Dim sSQL        As String
    Dim lRowCount   As Long
    UpdatePrinterInfo = False
    sSQL = "UPDATE  kodz5   "
    sSQL = sSQL & " SET crclientz5  =   '" & gsCompName & "'            , "
    sSQL = sSQL & "     crcodez5    =   '" & PrintDat.sPrintDatName & "', "
    sSQL = sSQL & "     crprintz5   =   NULL                            , "
    sSQL = sSQL & "     crymdz5     =   sysdate                         , "
    sSQL = sSQL & "     sdayz5     =   sysdate                         , "
    sSQL = sSQL & "     sndkz5     =   ' '                         , "
    If PrintDat.sInfoDat01 = "" Then
        sSQL = sSQL & " crkey01z5   = NULL                              , "
    Else
        sSQL = sSQL & " crkey01z5   = '" & PrintDat.sInfoDat01 & "'     , "
    End If
    If PrintDat.sInfoDat02 = "" Then
        sSQL = sSQL & " crkey02z5   = NULL                              , "
    Else
        sSQL = sSQL & " crkey02z5   = '" & PrintDat.sInfoDat02 & "'     , "
    End If
    If PrintDat.sInfoDat03 = "" Then
        sSQL = sSQL & " crkey03z5   = NULL                              , "
    Else
        sSQL = sSQL & " crkey03z5   = '" & PrintDat.sInfoDat03 & "'     , "
    End If
    If PrintDat.sInfoDat04 = "" Then
        sSQL = sSQL & " crkey04z5   = NULL                              , "
    Else
        sSQL = sSQL & " crkey04z5   = '" & PrintDat.sInfoDat04 & "'     , "
    End If
    If PrintDat.sInfoDat05 = "" Then
        sSQL = sSQL & " crkey05z5   = NULL                                "
    Else
        sSQL = sSQL & " crkey05z5   = '" & PrintDat.sInfoDat05 & "'       "
    End If
    sSQL = sSQL & "WHERE    crseqz5 = " & CStr(PrintDat.lWriteCnt)
    lRowCount = SqlExec2(sSQL)
    If lRowCount < 0 Then
        Call MsgOut(100, sSQL, ERR_LOG, "koda9")
        Exit Function
    ElseIf lRowCount = 0 Then
        sSQL = "INSERT INTO kodz5"
        sSQL = sSQL & "( crseqz5, crclientz5, crcodez5, croutz5, crprintz5, crymdz5,      "
        sSQL = sSQL & " sdayz5, sndkZ5, crkey01z5, crkey02z5, crkey03z5, crkey04z5, crkey05z5   )"
        
        sSQL = sSQL & " VALUES "
        sSQL = sSQL & "(" & CStr(PrintDat.lWriteCnt) & ","
        sSQL = sSQL & "'" & gsCompName & "'           ,"
        sSQL = sSQL & "'" & PrintDat.sPrintDatName & "', "
        sSQL = sSQL & " '1'                           ,"
        sSQL = sSQL & "NULL                           ,"
        sSQL = sSQL & "     sysdate,"
        sSQL = sSQL & "     sysdate,"
        sSQL = sSQL & "     ' ',"
        If PrintDat.sInfoDat01 = "" Then
            sSQL = sSQL & " NULL,"
        Else
            sSQL = sSQL & " '" & PrintDat.sInfoDat01 & "',"
        End If
        If PrintDat.sInfoDat02 = "" Then
            sSQL = sSQL & " NULL,"
        Else
            sSQL = sSQL & " '" & PrintDat.sInfoDat02 & "',"
        End If
        If PrintDat.sInfoDat03 = "" Then
            sSQL = sSQL & "  NULL,"
        Else
            sSQL = sSQL & "  '" & PrintDat.sInfoDat03 & "',"
        End If
        If PrintDat.sInfoDat04 = "" Then
            sSQL = sSQL & "  NULL,"
        Else
            sSQL = sSQL & "  '" & PrintDat.sInfoDat04 & "',"
        End If
        If PrintDat.sInfoDat05 = "" Then
            sSQL = sSQL & "  NULL)"
        Else
            sSQL = sSQL & "  '" & PrintDat.sInfoDat05 & "')"
        End If
        lRowCount = SqlExec2(sSQL)
        If lRowCount <> 1 Then
            Call MsgOut(100, sSQL, ERR_LOG, "kodz5")
            Exit Function
        End If
    End If
    UpdatePrinterInfo = True
End Function

' @(f)
' �@�\      :   �Ǘ��R�[�h�s�a�k�Ǐo������
' �Ԃ�l    :   TRUE    -   ����
'               FALSE   -   �ُ�
' ������    :   KanriDat    -   �Ǘ��R�[�h�f�[�^�i�[�f�[�^
'
' �@�\����  :   �Ǘ��R�[�h�e�[�u���̎w�肳�ꂽ�V�X�e���敪�A��ʃR�[�h�A�ʃR�[�h�Ńf�[�^����������
' ���l      :
' �C��      :   2000.04.27  �Ǘ��R�[�h�e�[�u���̒��[�����A�ԃ��R�[�h�͑��̒��[�o�̓��W���[�������������
'                           �Q�Ƃ���\��������̂ŏ����A�Ԃ��d�����Ȃ��悤�Ƀ��R�[�h�����b�N����
'
Public Function SelectKanriTable(KanriDat As KANRIKODEDAT) As Boolean
    Dim sSQL        As String
    Dim objDS       As Object
    SelectKanriTable = False
    sSQL = "SELECT  NVL(kcode01a9,' '), "
    sSQL = sSQL & " NVL(kcode02a9,' '), "
    sSQL = sSQL & " NVL(kcode03a9,' '), "
    sSQL = sSQL & " NVL(kcode04a9,' '), "
    sSQL = sSQL & " NVL(kcode05a9,' '), "
    sSQL = sSQL & " NVL(ctr01a9,0),     "
    sSQL = sSQL & " NVL(ctr02a9,0),     "
    sSQL = sSQL & " NVL(ctr03a9,0),     "
    sSQL = sSQL & " NVL(ctr04a9,0),     "
    sSQL = sSQL & " NVL(ctr05a9,0)      "
    sSQL = sSQL & "FROM koda9           "
    sSQL = sSQL & "WHERE    codea9  =   '" & KanriDat.sCode & "'"
    sSQL = sSQL & "  AND    shuca9  =   '" & KanriDat.sShuCode & "'"
    sSQL = sSQL & "  AND    sysca9  =   '" & KanriDat.sSysCode & "'"
    sSQL = sSQL & " FOR UPDATE           "   ''  �Ǘ��R�[�h�e�[�u�����R�[�h���b�N
    If DynSet2(objDS, sSQL) = False Then
        Call MsgOut(100, sSQL, ERR_LOG, "koda9")
        Exit Function
    End If
    If objDS.EOF = False Then
        Do Until objDS.EOF
            KanriDat.sDat01 = objDS(0)
            KanriDat.sDat02 = objDS(1)
            KanriDat.sDat03 = objDS(2)
            KanriDat.sDat04 = objDS(3)
            KanriDat.sDat05 = objDS(4)
            KanriDat.lCnt01 = objDS(5)
            KanriDat.lCnt02 = objDS(6)
            KanriDat.lCnt03 = objDS(7)
            KanriDat.lCnt04 = objDS(8)
            KanriDat.lCnt05 = objDS(9)
            objDS.MoveNext
        Loop
    Else
        Call MsgOut(55, sSQL, ERR_LOG, "koda9")
    End If
    SelectKanriTable = True
    Set objDS = Nothing
End Function

' @(f)
' �@�\      :   �Ǘ��R�[�h�s�a�k�X�V�iRW�|�C���^�X�V�j
' �Ԃ�l    :   1     :����
'               1�ȊO�F�ُ�
'
' ������    :   iUpFlg          -   �X�V�|�C���^�[�t���O
'               sWritePointer   -   �Vײ��߲��
'               sReadPointer    -   �Vذ���߲��
'
' �@�\����  : �˗����ð��قŎg�p���Ă��鏈���A�Ԃ��C���N�������g����B
'
' ���l      :   iUpFlg      �O�|��RW�|�C���^�X�V
'                           �P�|��R�|�C���^�̂ݍX�V
'                           �Q�|��W�|�C���^�̂ݍX�V
'
Public Function UpdateKanriTable(iUpFlg As Integer, sWritePointer As String, sReadPointer As String) As Boolean
    Dim sSQL        As String
    Dim lRowCount   As Long
    UpdateKanriTable = False
    sSQL = "UPDATE  koda9"
    sSQL = sSQL & " SET "
    Select Case iUpFlg
    Case 0  ''  RW�|�C���^�X�V
        sSQL = sSQL & " kcode02a9 = " & sReadPointer & ","
        sSQL = sSQL & " kcode03a9 = " & sWritePointer & ","
    Case 1  ''  R�|�C���^�̂ݍX�V
        sSQL = sSQL & " kcode02a9 = " & sReadPointer & ","
    Case 2  ''  W�|�C���^�̂ݍX�V
        sSQL = sSQL & " kcode03a9 = " & sWritePointer & ","
    End Select
    sSQL = sSQL & " sdaya9 =  sysdate,"
    sSQL = sSQL & " sndka9 =    ' '"
    sSQL = sSQL & " WHERE   codea9  =   '1' "
    sSQL = sSQL & "   AND   shuca9  =   '99'"
    sSQL = sSQL & "   AND   sysca9  =   'X' "
    'SQL���s
    lRowCount = SqlExec2(sSQL)
    Select Case lRowCount
    Case -1
        Call MsgOut(100, sSQL, ERR_LOG, "koda9")
        Exit Function
    Case 0
        Call MsgOut(72, "koda9", ERR_LOG)
        Exit Function
    Case 1
    Case Else
        Call MsgOut(0, "�������X�V����:" & sSQL, ERR_LOG)
        Exit Function
    End Select
    UpdateKanriTable = True
End Function

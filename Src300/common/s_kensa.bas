Attribute VB_Name = "s_kensa"
''
'' �\���́A�萔��`�i���������֘A�j
''

Option Explicit

'' �R���{�{�b�N�X�I�𕶎���
Public Const SP_COMBO_UMU1 = "�L" + vbTab + "��"                    '' �i�L�^���j
Public Const SP_COMBO_UMU2 = "�L" + vbTab + "��" + vbTab + "NDF"    '' �i�L�^���^�m�c�e�j
Public Const SP_COMBO_PN = "P" + vbTab + "N"                        '' �iP�^N�j
'Add Start 2011/01/05 SMPK Miyata
Public Const SP_COMBO_PTN1 = "0:None" + vbTab + "1:Ring" + vbTab + "2:Disk" + vbTab + "3:DiskRing"
Public Const SP_COMBO_PTN2 = "0:None" + vbTab + "5:PB-band" + vbTab + "6:P-band" + vbTab + "7:B-band"
'Add End   2011/01/05 SMPK Miyata


'' ���茋�ʕ�����
Public Const STR_JUDG_OK = "��"   '' ����OK
Public Const STR_JUDG_NG = "�~"   '' ����NG

'' �_�f�Y�f�������@�R�[�h��`
Public Const CODE_INSPECTWAY_FRIR = "CA"    '' FRIR
Public Const CODE_INSPECTWAY_FTIR = "CD"    '' FTIR
Public Const CODE_INSPECTWAY_SIMS = "CS"    '' SIMS
Public Const CODE_INSPECTWAY_GFA = "CG"     '' GFA

'' �f�t�H���g�l��`
Public Const DEF_PARAM_VALUE = -1                       '' �f�t�H���g�l
Public Const DEF_PARAM_DATE = "1901/01/01 01:00:01"     '' �f�t�H���g���t

'' �������͉\�����w���R�[�h
Public Const CODE_KENSA = "1234"


'' G�i�ԃf�t�H���g�l
'' ��R
Public Const DEFCODE_G_POS_RS = "Q3A"   '' ����ʒu�Q���_��
Public Const DEFCODE_G_GUA_RS = "3S"    '' �ۏؕ��@�Q�Ώ�
'' �b��
Public Const DEFCODE_G_POS_CS = "Q1Y"   '' ����ʒu�Q���_��
Public Const DEFCODE_G_GUA_CS = "1S"    '' �ۏؕ��@�Q�Ώ�
'' �n���i����ʒu�Q���A�Q�ʂ́A�˂炢�i�Ԃ̎d�l���g�p����j
Public Const DEFCODE_G_POS_OI = " 3 "   '' ����ʒu�Q���_��
Public Const DEFCODE_G_GUA_OI = "3S"    '' �ۏؕ��@�Q�Ώ�
'' ���C�t�^�C��
Public Const DEFCODE_G_POS_LT = "Q5J"   '' ����ʒu�Q���_��
Public Const DEFCODE_G_GUA_LT = "BS"    '' �ۏؕ��@�Q�Ώ�

'' Z�i�ԃf�t�H���g�l
'' ��R
Public Const DEFCODE_Z_POS_RS = "Q3A"   '' ����ʒu�Q���_��
Public Const DEFCODE_Z_GUA_RS = "3S"    '' �ۏؕ��@�Q�Ώ�
'' �b��
Public Const DEFCODE_Z_POS_CS = "Q1Y"   '' ����ʒu�Q���_��
Public Const DEFCODE_Z_GUA_CS = "1S"    '' �ۏؕ��@�Q�Ώ�


'' �M�����@���Ǘ��e�[�u��
Public Type typ_HeatInfo
    iHeatClass      As Integer  '' �M�������ށi�M���������̔ԍ��ɑΉ��B��FBMD1��1�j
    strHeatProc     As String   '' �M�������@�i�M�����@�R�[�h�j
End Type

'' �M���������R���{�{�b�N�X�Ǘ��e�[�u��
Public Type typ_cmbTInfo
    iHeatClass      As Integer  '' �M�������ށi�M���������̔ԍ��ɑΉ��B��FBMD1��1�j
    strHeatProc     As String   '' �M�������@�i�M�����@�R�[�h�j
End Type


'' ���[�U�f�[�^�Ǘ��e�[�u����`
Public tbl_HeatInfo() As typ_HeatInfo   '' �M�����@���Ǘ��e�[�u��
Public tbl_cmbTInfo() As typ_cmbTInfo   '' �M���������R���{�{�b�N�X�Ǘ��e�[�u��
''
'' �c�a�A�N�Z�X���W���[���i���������֘A�j
''

'' ���b�Z�[�W�R�[�h��`
Public Const MSG_NOTFOUND_STAFFID = "ESTAF" '' �S���҃R�[�h�G���[
'Public Const MSG_NOTFOUND_CRYNUM = "ECRY0"  '' �����ԍ��G���[
Public Const MSG_NOTFOUND_SMPLNO = "ENSMP"  '' �T���v��NO.�G���[
Public Const MSG_NOTFOUND_GOUKI_ = "ENGOK"  '' ���@�G���[
Public Const MSG_INPUT_STAFFID = "EISTF"    '' �S���҃R�[�h����͂��Ă��������B
Public Const MSG_INPUT_CRYNUM = "EICRY"     '' �����ԍ�����͂��Ă��������B
Public Const MSG_INPUT_SMPLNO = "EISMP"     '' �T���v��No.����͂��Ă��������B
Public Const MSG_INPUT_GOUKI = "EIGOK"      '' ���@����͂��Ă��������B
Public Const MSG_ERROR_PARAM = "EINPM"      '' ���͒l���s���ł��B
Public Const MSG_JUDG_ERROR = "EJUDG"       '' ����G���[
Public Const MSG_ENTRY_ERROR = "EETRY"      '' �o�^�G���[
Public Const MSG_SIGMACHECK_ERROR = "ESIGM" '' �V�O�}�`�F�b�N�G���[
Public Const MSG_R2CHECK_ERROR = "ER2CK"    '' ���֌W���`�F�b�N�G���[
Public Const MSG_CALCULATE_ERROR = "ECALC"  '' �v�Z�G���[
Public Const MSG_FTIRCHECK_ERROR = "EFTIR"  '' FTIR���Z�l�`�F�b�N�G���[
Public Const MSG_EFFECTTIME_ERROR = "EFTIM" '' FTIR���֎��L�����ԃG���[
'Public Const MSG_GETERROR_DBDATA = "EGET"   '' DB�f�[�^�擾�G���[
'Public Const MSG_DISPLAY_ERROR = "EDISP"    '' �\���G���[
Public Const MSG_ENTRY = "PPROK"            '' �o�^���b�Z�[�W
Public Const MSG_KTKBN = "ESMPK"            '' �m�胁�b�Z�[�W
Public Const MSG_INSPECT_ERROR = "EINSP"    '' �������@���Ή����b�Z�[�W

'' �����T���v���Ǘ��e�[�u���擾�E�X�V���[�h
Public Const MODE_GETSMPL_FTIR = 1          '' FTIR(Oi,Cs)
Public Const MODE_GETSMPL_GFA = 2           '' GFA(Oi)
Public Const MODE_GETSMPL_RS = 3            '' ��R
Public Const MODE_GETSMPL_BMD = 4           '' BMD
Public Const MODE_GETSMPL_OSF = 5           '' OSF
Public Const MODE_GETSMPL_GD = 6            '' GD
Public Const MODE_GETSMPL_LT = 7            '' ���C�t�^�C��
Public Const MODE_GETSMPL_EPD = 8           '' EPD
Public Const MODE_GETSMPL_X = 9             '' X��
Public Const MODE_GETSMPL_CUDECO = 10       '' Cu-deco(C,CJ,CJLT,CJ2)   Add 2010/12/17 SMPK Miyata

'' �����������
Public Enum chkKensaType
    CHK_OI         '' Oi
    CHK_CS         '' Cs
    CHK_RS         '' Rs
    CHK_B1         '' BMD1
    CHK_B2         '' BMD2
    CHK_B3         '' BMD3
    CHK_L1         '' OSF1
    CHK_L2         '' OSF2
    CHK_L3         '' OSF3
    CHK_L4         '' OSF4
    CHK_GD         '' GD
    CHK_LT         '' LT
    CHK_EP         '' EPD
    CHK_X          '' X��   2009/08 SUMCO Akizuki
    'Add Start 2010/12/17 SMPK Miyata
    CHK_C          '' C
    CHK_CJ         '' CJ
    CHK_CJLT       '' CJLT
    CHK_CJ2        '' CJ2
    'Add End   2010/12/17 SMPK Miyata
End Enum

'' �f�[�^�Ǘ��e�[�u����`�i�O���[�o���ϐ���`�j
Public tbl_PrSpSXLData1() As typ_TBCME018   '' ���i�d�l�r�w�k�f�[�^�P
Public tbl_PrSpSXLData2() As typ_TBCME019   '' ���i�d�l�r�w�k�f�[�^�Q
Public tbl_PrSpSXLData3() As typ_TBCME020   '' ���i�d�l�r�w�k�f�[�^�R
'*** UPDATE START Y.SIMIZU 2005/10/1 TBCME036ð����ް��i�[�p
Public tbl_PrSpSXLData4() As typ_TBCME036   '' ���i�d�l�r�w�k�f�[�^�S
'*** UPDATE END Y.SIMIZU 2005/10/1 TBCME036ð����ް��i�[�p
Public tbl_PrSpWFData1() As typ_TBCME021            '' ���i�d�l�v�e�f�[�^�P�@05/03/01 ooba START ==>
Public tbl_PrSpWFData2() As typ_TBCME022            '' ���i�d�l�v�e�f�[�^�Q
Public tbl_PrSpWFData6() As typ_TBCME026            '' ���i�d�l�v�e�f�[�^�U
Public tbl_PrSpWFData8() As typ_TBCME028            '' ���i�d�l�v�e�f�[�^�W  2005/06/15 ffc)tanabe

''Upd Start (TCS)T.Terauchi 2005/10/07
Public tbl_PrSpWFData36() As typ_TBCME036            '' ���i�d�l�v�e�f�[�^
''Upd End   (TCS)T.Terauchi 2005/10/07

Public tbl_CrystalSampleManage_Cw() As typ_XSDCW    '' �V����يǗ�(SXL)
Public tbl_HinbanCW() As tFullHinban                '' �V����يǗ�(SXL)�̕i��
Public tbl_WFGDRslt() As typ_TBCMJ015               '' �f�c����(WF)         05/03/01 ooba END ====>
Public tbl_WFSPVRslt() As typ_TBCMJ016              '' SPV����(WF)          2005/06/16 ffc)tanabe
Public tbl_SXLInsideSpecManager() As typ_TBCME036   '' ���������Ǘ�
Public tbl_PlupEndRslt() As typ_TBCMH004            '' ���グ�I������

Public tbl_GFADevInfo() As typ_TBCMB014             '' �f�e�`�Z�����
Public tbl_HinbanManage() As typ_TBCME041           '' �i�ԊǗ�
Public tbl_BlockManage() As typ_TBCME040            '' �u���b�N�Ǘ�
Public tbl_CrystalSampleManage() As typ_XSDCS       '' �����T���v���Ǘ�
Public tbl_CrystalSampleManage2() As typ_XSDCS      '' �����T���v���Ǘ�     Add 2010/12/17 SMPK Miyata
Public tbl_EPDRslt() As typ_TBCMJ001                '' �d�o�c����
Public tbl_CryRsRslt() As typ_TBCMJ002              '' ������R����
Public tbl_OiRslt() As typ_TBCMJ003                 '' �n������
Public tbl_CsRslt() As typ_TBCMJ004                 '' �b������
Public tbl_BMDRslt() As typ_TBCMJ008                '' �a�l�c����
Public tbl_OSFRslt() As typ_TBCMJ005                '' �n�r�e����
Public tbl_GDRslt() As typ_TBCMJ006                 '' �f�c����
Public tbl_LifeTime() As typ_TBCMJ007               '' ���C�t�^�C������
Public tbl_XRslt() As typ_TBCMJ021                  '' X���������          2009/08 SUMCO Akizuki
Public tbl_CuDecoRslt() As typ_TBCMJ023             '' Cu-deco����          Add 2010/12/17 SMPK Miyata


'' ���R�v�Z�ʒu�v�Z���
Public Type typ_CalcRsPosInf
    dChgWt      As Double   '' �d���ݏd��(�`���[�W��)
    dTopWT      As Double   '' �g�b�v�d��
    dArea       As Double   '' �f�ʐ�
    dHenseki    As Double   '' ���s�ΐ͒l
    dSmpPos     As Double   '' �T���v���ʒu
    dR0Ce       As Double   '' �T���v���ʒu0mm�ɂ�������R�����l�i���݂���ꍇ�ɐݒ�j
    dRx         As Double   '' �ΏۃT���v���̔��R�����l
End Type

'' Cs70%����l�v�Z���
Public Type typ_Cs70PInf
    dChgWt      As Double   '' �d���ݏd��(�`���[�W��)
    dTopWT      As Double   '' �g�b�v�d��
    dArea       As Double   '' �f�ʐ�
    dSmpPos     As Double   '' �T���v���ʒu
    dCs         As Double   '' Cs�Z�x�l
End Type

'�T�v      :�R���{�{�b�N�X�̏�ԕύX���s��
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :ctrlObj       ,I   ,Control   ,�R���g���[���I�u�W�F�N�g
'          :bFlag         ,I   ,Boolean   ,�R���g���[����Ԏw���iTrue�F�L���@False�F�����j
'          :[bClear]      ,I   ,Boolean   ,�R���{�{�b�N�X���e�N���A�w���iTrue�F�N���A�@False�F�N���A���Ȃ��j
'����      :
Public Sub EnableComboBoxCtrl(ctrlObj As Control, bFlag As Boolean, Optional bClear As Boolean = False)

    If bFlag = True Then
        ctrlObj.Enabled = True
        ctrlObj.BackColor = vbWindowBackground
    Else
        ctrlObj.Enabled = False
        ctrlObj.BackColor = vbButtonFace
    End If

    If bClear Then
        ctrlObj.Clear
    End If

End Sub


'�T�v      :�J���}��؂蕶���񂩂�C�ӂ̏ꏊ�̕������؂���
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :strTarget     ,I   ,String    ,�J���}��؂�̕�����
'          :iField        ,I   ,Integer   ,�擾�������ꏊ(1�`)
'          :�߂�l        ,O  ,String    ,�擾������B�擾�ł��Ȃ������ꍇvbNullString��Ԃ�
'����      :
Public Function GetStringField(strTarget As String, iField As Integer) As String
    Dim strWork     As String
    Dim strGet      As String
    Dim iPos        As Integer
    Dim Index       As Integer
    
    GetStringField = vbNullString

    If iField <= 0 Then Exit Function

    strWork = strTarget
    Index = 1
    Do
        iPos = InStr(strWork, ",")
        If iPos = 0 Then
            If Index < iField Then Exit Function
            strGet = strWork
            Exit Do
        End If
        strGet = Left(strWork, iPos - 1)
        strWork = Right(strWork, Len(strWork) - iPos)
        If Index = iField Then Exit Do
        Index = Index + 1
    Loop

    GetStringField = strGet

End Function

'�T�v      :�����_�؂�̂Ă��s��
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :param         ,I   ,Double    ,�Ώې��l
'          :iPoint        ,I   ,Integer   ,�����_�ȉ��؂�̂Đ�
'          :�߂�l        ,O  ,Double    ,�����_�ȉ��؎̂Č���
'����      :
Public Function CutDecimalPointParam(ByVal param As Double, iPoint As Integer) As Double
    Dim Index      As Integer
    Dim strParam   As String
    Dim iStrLen    As Integer
    Dim iTen       As Integer
    Dim strWork    As String
    Dim bFlag      As Boolean

    bFlag = False
    strParam = Str(param)
    iStrLen = Len(strParam)
    For Index = 1 To iStrLen
        If Mid(strParam, Index, 1) = "." Then
            bFlag = True
            iTen = Index
            Exit For
        End If
    Next Index
    If bFlag <> True Then CutDecimalPointParam = param: Exit Function
    
    strWork = Mid(strParam, iTen + 1, iPoint)
    CutDecimalPointParam = val(Left(strParam, iTen) + strWork)

End Function

'�T�v      :�����_�؂�グ���s��
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :param         ,I   ,Double    ,�Ώې��l
'          :iPoint        ,I   ,Integer   ,�����_�ȉ��؂�グ��
'          :�߂�l        ,O  ,Double    ,�����_�ȉ��؂�グ����
'����      :
Public Function UpDecimalPointParam(ByVal param As Double, iPoint As Integer) As Double
    Dim dWork As Double
    
    dWork = param - CutDecimalPointParam(param, iPoint)
    If dWork > 0 Then
        UpDecimalPointParam = CutDecimalPointParam(param, iPoint) + (10 ^ (-iPoint))
    Else
        UpDecimalPointParam = param
    End If

End Function

'�T�v      :�R���{�{�b�N�X�̏�ԕύX���s��
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :ctrlObj       ,I   ,Control   ,�R���g���[���I�u�W�F�N�g
'          :�߂�l        ,O  ,Integer   ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Public Function CheckIsInputText(ctrlObj As Control) As Integer
    Dim bDisable As Boolean
    
    '' ������
    bDisable = False
    
    CheckIsInputText = FUNCTION_RETURN_SUCCESS
    
    '' �\�����ڂł���ꍇ
    If ctrlObj.BackColor = COLOR_DISABLE And _
           ctrlObj.Locked = True And ctrlObj.TabStop = False Then
        bDisable = True
    End If
    
    '' ���̓`�F�b�N
    If ctrlObj.Text <> "" Then  '' ���͂���Ă���ꍇ
        If bDisable <> True Then
            CtrlEnabled ctrlObj, CTRL_ENABLE
        End If
        Exit Function
    Else                        '' ���͂���Ă��Ȃ��ꍇ
        If bDisable = True Then
            Exit Function
        End If
        CtrlEnabled ctrlObj, CTRL_WARNING
    End If
    
    CheckIsInputText = FUNCTION_RETURN_FAILURE
End Function


'�T�v      :�w��T���v��No.���m�肳��Ă��邩���ׂ�
'���Ұ�    :�ϐ���        ,IO ,�^             ,����
'          :tblCrySmp()   ,I  ,typ_XSDCS   ,�����T���v���Ǘ��e�[�u���z��
'          :iSmpNo        ,I  ,Long        ,�T���v��No.     Integer��Long 6���Ή� 2007/05/28 SETsw kubota
'          :�߂�l        ,O  ,Boolean        ,True:�m�肵�Ă���  False:�m�肵�Ă��Ȃ�
'����      :
Public Function CheckKTKBN(tblCrySmp() As typ_XSDCS, iSmpNo As Long) As Boolean
    Dim Index As Integer
    
    CheckKTKBN = False
    For Index = 0 To UBound(tblCrySmp) - 1
        If (tblCrySmp(Index).REPSMPLIDCS = iSmpNo) And (tblCrySmp(Index).KTKBNCS = "1") Then
            CheckKTKBN = True
            Exit Function
        End If
    Next Index

End Function
'�T�v      :����t���O�l��蔻�蕶������擾����
'���Ұ�    :�ϐ���       ,IO ,�^        ,����
'          :bJudg        ,I   ,Boolean  ,����t���O�l
'          :�߂�l        ,O  ,String   ,���蕶����
'����      :
Public Function GetJudgStr(bJudg As Boolean) As String

    If bJudg = True Then
        GetJudgStr = STR_JUDG_OK
    Else
        GetJudgStr = STR_JUDG_NG
    End If

End Function
'�T�v      :����t���O�l��蔻�蕶������擾����
'���Ұ�    :�ϐ���       ,IO ,�^        ,����
'          :strJudg      ,I   ,String  ,����t���O
'          :�߂�l        ,O  ,String   ,���蕶����
'����      :
Public Function GetResJudgStr(strJudg As String) As String

    Select Case strJudg
    Case "1"
        GetResJudgStr = STR_JUDG_OK
    Case "2"
        GetResJudgStr = STR_JUDG_NG
    Case Else
        GetResJudgStr = ""
    End Select

End Function
'�T�v      :����t���O�l��茋���������уR�[�h���쐬����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :bJudg         ,I   ,Boolean   ,����t���O�l
'          :�߂�l        ,O  ,String    ,�����������уR�[�h
'����      :
Public Function MakeCryResultCode(bJudg As Boolean) As String


''�@����FLG�́u0:������,1:����OK,2:����NG�v�ύX�ɔ����A�ďC���@2003/09/26 SystemBrain ==========================> START
''�@�����w���ύX�@2003/09/10 Motegi ==========================> START
    If bJudg = True Then
        MakeCryResultCode = "1"
    Else
        MakeCryResultCode = "2"
    End If
     
'     MakeCryResultCode = "1"
''�@�����w���ύX�@2003/09/10 Motegi ==========================> END

End Function
'�T�v      :����_�R�[�h��葪��_�����߂�
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :strMeasNum    ,I   ,String  ,����_�R�[�h�i1byte�j
'          :[iDefNum]     ,I   ,Integer  ,�f�t�H���g����_�i�Y���̑���_�R�[�h���Ȃ��ꍇ�ɂ��̒l��Ԃ��j
'          :�߂�l        ,O  ,Integer   ,����_
'����      :
Public Function GetMeasureNum(strMeasNum As String, Optional iDefNum As Integer = 0) As Integer
    Dim iNum    As Integer

    iNum = iDefNum
    If strMeasNum <> "" Then
        If Asc(strMeasNum) >= Asc("0") And Asc(strMeasNum) <= Asc("9") Then
            '' �O�`�X�̏ꍇ
            iNum = val(strMeasNum)
        ElseIf Asc(strMeasNum) >= Asc("A") And Asc(strMeasNum) <= Asc("K") Then
            '' �`�`�j�̏ꍇ
            iNum = 10 + Asc(strMeasNum) - Asc("A")
        ElseIf strMeasNum = "X" Then
            '' �w�̏ꍇ
            iNum = 20
        End If
    End If
    
    GetMeasureNum = iNum

End Function
'' ���@�R���{�{�b�N�X��������R�[�h���擾����
Public Function GetCmbCode(cmb As ComboBox) As String
Dim s As String
Dim POS As Integer

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "f_cmhc001j.frm -- Function GetCmbCode"

    s = cmb.Text
    POS = InStr(1, s, ":")
    If POS Then
        s = Left$(s, POS - 1)
    End If
    GetCmbCode = s

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function
'' �i�ԏ����N���A����
Public Sub ClearFullHinban(tHIN As tFullHinban)
    tHIN.hinban = ""
    tHIN.mnorevno = 0
    tHIN.factory = ""
    tHIN.opecond = ""
End Sub

'' �i�ԊǗ����i�ԏ����Z�b�g����
Public Sub SetFullHinban_TBCME041(tHIN As tFullHinban, tblHinban As typ_TBCME041)
    tHIN.hinban = tblHinban.hinban
    tHIN.mnorevno = tblHinban.REVNUM
    tHIN.factory = tblHinban.factory
    tHIN.opecond = tblHinban.opecond
End Sub

'�T�v      :�T���v���m��.��茋���ԍ����擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :strCryNum     ,O   ,String    ,�����ԍ�
'          :lSmpNo        ,I   ,Long      ,�T���v���m��.
'          :lSmpMode      ,I   ,Long      ,�T���v���Ǘ��e�[�u���擾���[�h
'          :�߂�l        ,O  ,Integer    ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
'Public Function GetCryNum(strCryNum As String, iSmpNo As Long) As Integer
Public Function GetCryNum(strCryNum As String, lSmpNo As Long, Optional ByVal lSmpMode As Long = 0) As Integer
    Dim iRet        As Integer
    Dim tblGet()    As typ_XSDCS
    
    Dim lCnt        As Long
    Dim lXtalNoCnt  As Long

    GetCryNum = FUNCTION_RETURN_FAILURE

    '' �����ԍ��̎擾
    'iRet = DBDRV_GetTBCME043(tblGet, "where REPSMPLIDCS=" & CStr(lSmpNo) & " and KTKBNCS='0'")
    iRet = DBDRV_GetTBCME043(tblGet, "where REPSMPLIDCS=" & CStr(lSmpNo) & " and KTKBNCS='0'", "order by TDAYCS desc")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then
        'iRet = DBDRV_GetTBCME043(tblGet, "where REPSMPLIDCS=" & CStr(lSmpNo))
        iRet = DBDRV_GetTBCME043(tblGet, "where REPSMPLIDCS=" & CStr(lSmpNo), "order by TDAYCS desc")
        If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
        If UBound(tblGet) = 0 Then Exit Function
'>>>>> �T���v��No.�`�F�b�N�̒ǉ� 2011/06/27 SETsw kubota -----------------------------------
        '�m��ς݂̏ꍇ�A��\�T���v��ID�����łȂ��A�e�T���v���̃T���v��ID�ƍ��v���邩���`�F�b�N
        lXtalNoCnt = 0
        For lCnt = 1 To UBound(tblGet)
            '�T���v��ID�����v�����烋�[�v���甲����
            If ChkMeasSmpl(tblGet(lCnt), lSmpNo, lSmpMode) = True Then
                lXtalNoCnt = lCnt       '���v�����s��Ԃ�
                Exit For
            End If
        Next lCnt
        If lXtalNoCnt = 0 Then
            '���v����f�[�^��������΃G���[
            Exit Function
        End If
    Else
        '���m��̃��R�[�h������΁A�����ʂ�
        lXtalNoCnt = 1
'<<<<< �T���v��No.�`�F�b�N�̒ǉ� 2011/06/27 SETsw kubota -----------------------------------
    End If
    
    'strCryNum = tblGet(1).XTALCS
    strCryNum = tblGet(lXtalNoCnt).XTALCS
    
    GetCryNum = FUNCTION_RETURN_SUCCESS

End Function

'�T�v      :�T���v���m��.���e����̃T���v��ID�J�����ɂ��邩�𔻒f
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :tXsdcs        ,I  ,typ_XSDCS ,XSDCS�f�[�^
'          :lSmpNo        ,I  ,Long      ,�T���v���m��.
'          :lSmpMode      ,I  ,Long      ,�T���v���Ǘ��e�[�u���擾���[�h
'          :�߂�l        ,O  ,Boolean   ,�T���v��ID����v�FTrue�@��v���Ȃ��FFalse
'����      :2011/06/28�ǉ� SETsw kubota
Public Function ChkMeasSmpl(ByRef tXsdcs As typ_XSDCS _
                          , ByVal lSmpNo As Long _
                          , ByVal lSmpMode As Long _
                          ) As Boolean
    
    Dim bSmpFlg     As Boolean
    bSmpFlg = False
    
    '�e�T���v���̃T���v��ID�ƈ�v���邩���`�F�b�N
    With tXsdcs
        Select Case lSmpMode
        Case MODE_GETSMPL_FTIR      '' FTIR(Oi,Cs)
            If lSmpNo = .CRYSMPLIDOICS _
            Or lSmpNo = .CRYSMPLIDCSCS Then
                bSmpFlg = True
            End If
        Case MODE_GETSMPL_GFA       '' GFA(Oi)
            If lSmpNo = .CRYSMPLIDOICS Then
                bSmpFlg = True
            End If
        Case MODE_GETSMPL_RS        '' ��R
            If lSmpNo = .CRYSMPLIDRSCS Then
                bSmpFlg = True
            End If
        Case MODE_GETSMPL_BMD       '' BMD
            If lSmpNo = .CRYSMPLIDB1CS _
            Or lSmpNo = .CRYSMPLIDB2CS _
            Or lSmpNo = .CRYSMPLIDB3CS Then
                bSmpFlg = True
            End If
        Case MODE_GETSMPL_OSF       '' OSF
            If lSmpNo = .CRYSMPLIDL1CS _
            Or lSmpNo = .CRYSMPLIDL2CS _
            Or lSmpNo = .CRYSMPLIDL3CS _
            Or lSmpNo = .CRYSMPLIDL4CS Then
                bSmpFlg = True
            End If
        Case MODE_GETSMPL_GD        '' GD
            If lSmpNo = .CRYSMPLIDGDCS Then
                bSmpFlg = True
            End If
        Case MODE_GETSMPL_LT        '' ���C�t�^�C��
            If lSmpNo = .CRYSMPLIDTCS Then
                bSmpFlg = True
            End If
        Case MODE_GETSMPL_EPD       '' EPD
            If lSmpNo = .CRYSMPLIDEPCS Then
                bSmpFlg = True
            End If
        Case MODE_GETSMPL_X         '' X��
            If lSmpNo = .CRYSMPLIDXCS Then
                bSmpFlg = True
            End If
        Case MODE_GETSMPL_CUDECO    '' Cu-deco(C,CJ,CJLT,CJ2)
            If lSmpNo = .CRYSMPLIDCCS _
            Or lSmpNo = .CRYSMPLIDCJCS _
            Or lSmpNo = .CRYSMPLIDCJLTCS _
            Or lSmpNo = .CRYSMPLIDCJ2CS Then
                bSmpFlg = True
            End If
        Case Else                   '' ���̑�
            '���̑��͂��̂܂�OK
            bSmpFlg = True
        End Select
    End With

    ChkMeasSmpl = bSmpFlg

End Function

'�T�v      :�Ј��R�[�h���Ј������擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :strName       ,O   ,String    ,�Ј���
'          :strID         ,I   ,String    ,�Ј��R�[�h
'          :�߂�l        ,O  ,Integer    ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Public Function GetStaffNameStr(strName As String, strID As String) As Integer

    GetStaffNameStr = FUNCTION_RETURN_FAILURE

    '' �Ј����̎擾
        '2009/09 Akizuki TBCMB001�Q�Ƃ���AKODA9�Q�Ƃ֕ύX
        'strName = GetStaffName(strID)
        strName = GetStaffName_KODA9(strID)
    
    If strName = vbNullString Then
        Exit Function
    End If

    GetStaffNameStr = FUNCTION_RETURN_SUCCESS

End Function
'�T�v      :�����ԍ����i�ԊǗ��e�[�u�����擾����B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :tblHinban()   ,O   ,typ_TBCME041 ,�i�ԊǗ��e�[�u��
'          :strCryNum     ,I   ,String           ,�����ԍ�
'          :�߂�l        ,O  ,Integer          ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Public Function GetHinban(tblHinban() As typ_TBCME041, strCryNum As String) As Integer
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
            .factory = vbNullString
            .opecond = vbNullString
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

'�T�v      :�����ԍ����V����يǗ�(SXL)�̕i�Ԃ��擾����B
'���Ұ�    :�ϐ���        ,IO  ,�^               ,����
'          :tblHinban()   ,O   ,tFullHinban      ,12���i��
'          :strCryNum     ,I   ,String           ,�����ԍ�
'          :�߂�l        ,O   ,Integer          ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
'����      :05/03/01 ooba
Public Function GetHinban_WF(tblHinban() As tFullHinban, strCryNum As String) As Integer

    Dim Index       As Integer
    Dim tblIndex    As Integer
    Dim iChk        As Integer
    Dim bChkFlg     As Boolean

    GetHinban_WF = FUNCTION_RETURN_FAILURE

    '' 12���i�ԍ\���̂�������
    ReDim tblHinban(0)

    '' �V����يǗ�(SXL)ð��ق����݂��Ȃ���Ώ����𔲂���
    If UBound(tbl_CrystalSampleManage_Cw) = 0 Then Exit Function
    
    For Index = 1 To UBound(tbl_CrystalSampleManage_Cw)
        If Trim(tbl_CrystalSampleManage_Cw(Index).HINBCW) <> "Z" And _
           Trim(tbl_CrystalSampleManage_Cw(Index).HINBCW) <> "G" And _
           Trim(tbl_CrystalSampleManage_Cw(Index).HINBCW) <> "" Then
            
            bChkFlg = False
            '' �����i�ԂƂ̏d������
            If UBound(tblHinban) > 0 Then
                For iChk = 1 To UBound(tblHinban)
                    '' �i�Ԃ���v���Ă�����ް���Ă��Ȃ�
                    If tbl_CrystalSampleManage_Cw(Index).HINBCW = tblHinban(iChk).hinban And _
                       tbl_CrystalSampleManage_Cw(Index).REVNUMCW = tblHinban(iChk).mnorevno And _
                       tbl_CrystalSampleManage_Cw(Index).FACTORYCW = tblHinban(iChk).factory And _
                       tbl_CrystalSampleManage_Cw(Index).OPECW = tblHinban(iChk).opecond Then
                       
                        bChkFlg = True
                        Exit For
                    End If
                Next
            End If
            
            If bChkFlg = False Then
                '' �i���ް��i�[�̈�g��
                tblIndex = UBound(tblHinban) + 1
                ReDim Preserve tblHinban(tblIndex)
                '' 12���i���ް��̎擾
                tblHinban(tblIndex).hinban = tbl_CrystalSampleManage_Cw(Index).HINBCW
                tblHinban(tblIndex).mnorevno = tbl_CrystalSampleManage_Cw(Index).REVNUMCW
                tblHinban(tblIndex).factory = tbl_CrystalSampleManage_Cw(Index).FACTORYCW
                tblHinban(tblIndex).opecond = tbl_CrystalSampleManage_Cw(Index).OPECW
            End If
        End If
    Next Index
    
    If UBound(tblHinban) = 0 Then
        If Len(strCryNum) <> 12 Then Exit Function
        
        ReDim tblHinban(1)
        tblHinban(1).hinban = "Z"
        tblHinban(1).mnorevno = 0
        tblHinban(1).factory = vbNullString
        tblHinban(1).opecond = vbNullString
    End If

    GetHinban_WF = FUNCTION_RETURN_SUCCESS

End Function

'�T�v      :�i�Ԃ�萻�i�d�l�r�w�k�f�[�^�P���擾�A�����āA�擾�������i�d�l�f�[�^��ǉ�����B
'���Ұ�    :�ϐ���        ,IO ,�^                   ,����
'          :tblSP()       ,O   ,typ_TBCME018        ,���i�d�l�r�w�k�f�[�^�P�e�[�u���z��
'          :tHinInf       ,I   ,tFullHinban         ,�i��
'          :ctrlFrm       ,I   ,Form                ,�t�H�[��ID
'          :�߂�l        ,O  ,Integer              ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Public Function GetSPSXLData1(tblSP() As typ_TBCME018, tHinInf As tFullHinban, ctrlFrm As Form) As Integer
    Dim Index       As Long
    Dim iRet        As Integer
    Dim tHinban(0)  As tFullHinban
    Dim tblGet()    As typ_TBCME018
    
    GetSPSXLData1 = FUNCTION_RETURN_FAILURE

    tHinban(0) = tHinInf
    
    ''Cng Start 2011/02/21 Y.Hitomi �����������t�]
    '' ���łɎ擾�i�Ԏd�l���L�����Ă���ꍇ�A�擾�f�[�^��ǉ����Ȃ�
    For Index = 0 To UBound(tblSP) - 1
        If (tblSP(Index).hinban = tHinInf.hinban) And (tblSP(Index).mnorevno = tHinInf.mnorevno) And _
           (tblSP(Index).factory = tHinInf.factory) And (tblSP(Index).opecond = tHinInf.opecond) Then
            GetSPSXLData1 = FUNCTION_RETURN_SUCCESS
            Exit Function
        End If
    Next Index
    
    '' ���i�d�l�r�w�k�f�[�^�P�̎擾
    iRet = DBDRV_GetTBCME018(tblGet, ctrlFrm.Name, tHinban)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function
    ''Cng End 2011/02/21 Y.Hitomi
    
    '' �擾�������i�d�l�r�w�k�f�[�^�̒ǉ�
    Add_PrSpSXLData1 tblSP, tblGet(1)
    
    GetSPSXLData1 = FUNCTION_RETURN_SUCCESS

End Function


'�T�v      :�i�Ԃ�萻�i�d�l�r�w�k�f�[�^�Q���擾�A�����āA�擾�������i�d�l�f�[�^��ǉ�����B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :tblSP()       ,O   ,typ_TBCME019     ,���i�d�l�r�w�k�f�[�^�Q�e�[�u���z��
'          :tHinInf       ,I   ,tFullHinban      ,�i��
'          :ctrlFrm       ,I   ,Form             ,�t�H�[��ID
'          :�߂�l        ,O  ,Integer          ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Public Function GetSPSXLData2(tblSP() As typ_TBCME019, tHinInf As tFullHinban, ctrlFrm As Form) As Integer
    Dim Index       As Long
    Dim iRet        As Integer
    Dim tHinban(0)  As tFullHinban
    Dim tblGet()    As typ_TBCME019

    GetSPSXLData2 = FUNCTION_RETURN_FAILURE

    tHinban(0) = tHinInf

    ''Cng Start 2011/02/21 Y.Hitomi �����������t�]
    '' ���łɎ擾�i�Ԏd�l���L�����Ă���ꍇ�A�擾�f�[�^��ǉ����Ȃ�
    For Index = 0 To UBound(tblSP) - 1
        If (tblSP(Index).hinban = tHinInf.hinban) And (tblSP(Index).mnorevno = tHinInf.mnorevno) And _
           (tblSP(Index).factory = tHinInf.factory) And (tblSP(Index).opecond = tHinInf.opecond) Then
            GetSPSXLData2 = FUNCTION_RETURN_SUCCESS
            Exit Function
        End If
    Next Index
    
        '' ���i�d�l�r�w�k�f�[�^�Q�̎擾
    iRet = DBDRV_GetTBCME019(tblGet, ctrlFrm.Name, tHinban)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function
    ''Cng End 2011/02/21 Y.Hitomi
    
    '' �擾�������i�d�l�r�w�k�f�[�^�̒ǉ�
    Add_PrSpSXLData2 tblSP, tblGet(1)
    
    GetSPSXLData2 = FUNCTION_RETURN_SUCCESS

End Function


'�T�v      :�i�Ԃ�萻�i�d�l�r�w�k�f�[�^�R���擾�A�����āA�擾�������i�d�l�f�[�^��ǉ�����B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :tblSP()       ,O   ,typ_TBCME020 ,���i�d�l�r�w�k�f�[�^�R�e�[�u���z��
'          :tHinInf       ,I   ,tFullHinban         ,�i��
'          :ctrlFrm       ,I   ,Form             ,�t�H�[��ID
'          :�߂�l        ,O  ,Integer           ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Public Function GetSPSXLData3(tblSP() As typ_TBCME020, tHinInf As tFullHinban, ctrlFrm As Form) As Integer
    Dim Index       As Long
    Dim iRet        As Integer
    Dim tHinban(0)  As tFullHinban
    Dim tblGet()    As typ_TBCME020
    
    GetSPSXLData3 = FUNCTION_RETURN_FAILURE

    tHinban(0) = tHinInf

    ''Cng Start 2011/02/21 Y.Hitomi �����������t�]
    '' ���łɎ擾�i�Ԏd�l���L�����Ă���ꍇ�A�擾�f�[�^��ǉ����Ȃ�
    For Index = 0 To UBound(tblSP) - 1
        If (tblSP(Index).hinban = tHinInf.hinban) And (tblSP(Index).mnorevno = tHinInf.mnorevno) And _
           (tblSP(Index).factory = tHinInf.factory) And (tblSP(Index).opecond = tHinInf.opecond) Then
            GetSPSXLData3 = FUNCTION_RETURN_SUCCESS
            Exit Function
        End If
    Next Index
    
    '' ���i�d�l�r�w�k�f�[�^�R�̎擾
    iRet = DBDRV_GetTBCME020(tblGet, ctrlFrm.Name, tHinban)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function
    ''Cng End 2011/02/21 Y.Hitomi
    
    '' �擾�������i�d�l�r�w�k�f�[�^�̒ǉ�
    Add_PrSpSXLData3 tblSP, tblGet(1)
    
    GetSPSXLData3 = FUNCTION_RETURN_SUCCESS

End Function

'*** UPDATE �� Y.SIMIZU 2005/10/1
'�T�v      :�i�Ԃ�萻�i�d�l�r�w�k�f�[�^�S���擾�A�����āA�擾�������i�d�l�f�[�^��ǉ�����B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :tblSP()       ,O   ,typ_TBCME036 ,���i�d�l�r�w�k�f�[�^�S�e�[�u���z��
'          :tHinInf       ,I   ,tFullHinban         ,�i��
'          :ctrlFrm       ,I   ,Form             ,�t�H�[��ID
'          :�߂�l        ,O  ,Integer           ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Public Function GetSPSXLData4(tblSP() As typ_TBCME036, tHinInf As tFullHinban, ctrlFrm As Form) As Integer
    Dim Index       As Long
    Dim iRet        As Integer
    Dim tHinban(0)  As tFullHinban
    Dim tblGet()    As typ_TBCME036
    
    GetSPSXLData4 = FUNCTION_RETURN_FAILURE

    tHinban(0) = tHinInf

    ''Cng Start 2011/02/21 Y.Hitomi �����������t�]
    '' ���łɎ擾�i�Ԏd�l���L�����Ă���ꍇ�A�擾�f�[�^��ǉ����Ȃ�
    For Index = 0 To UBound(tblSP) - 1
        If (tblSP(Index).hinban = tHinInf.hinban) And (tblSP(Index).mnorevno = tHinInf.mnorevno) And _
           (tblSP(Index).factory = tHinInf.factory) And (tblSP(Index).opecond = tHinInf.opecond) Then
            GetSPSXLData4 = FUNCTION_RETURN_SUCCESS
            Exit Function
        End If
    Next Index
    
    '' ���i�d�l�r�w�k�f�[�^�S�̎擾
    iRet = DBDRV_GetTBCME036(tblGet, ctrlFrm.Name, tHinban)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function
    ''Cng End 2011/02/21 Y.Hitomi

    
    '' �擾�������i�d�l�r�w�k�f�[�^�̒ǉ�
    Add_PrSpSXLData4 tblSP, tblGet(1)
    
    GetSPSXLData4 = FUNCTION_RETURN_SUCCESS

End Function
'*** UPDATE �� Y.SIMIZU 2005/10/1

'�T�v      :�i�Ԃ�萻�i�d�l�v�e�f�[�^�P���擾�A�����āA�擾�������i�d�l�f�[�^��ǉ�����B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :tblSP()       ,O   ,typ_TBCME021   ,���i�d�l�v�e�f�[�^�P�e�[�u���z��
'          :tHinInf       ,I   ,tFullHinban    ,12���i��
'          :ctrlFrm       ,I   ,Form           ,�t�H�[��ID
'          :�߂�l        ,O  ,Integer          ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
'����      :05/03/01 ooba
Public Function GetSPWFData1(tblSP() As typ_TBCME021, tHinInf As tFullHinban, ctrlFrm As Form) As Integer

    Dim Index       As Integer
    Dim iRet        As Integer
    Dim tHinban(0)  As tFullHinban
    Dim tblGet()    As typ_TBCME021
    
    GetSPWFData1 = FUNCTION_RETURN_FAILURE

    tHinban(0) = tHinInf

    '' ���i�d�l�v�e�f�[�^�P�̎擾
    iRet = DBDRV_GetTBCME021(tblGet, ctrlFrm.Name, tHinban)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function

    '' �擾�������i�d�l�v�e�f�[�^�̒ǉ�
    '' �e�[�u���f�[�^�i�[�̈�g��
    ReDim Preserve tblSP(UBound(tblSP) + 1)
    '' �e�[�u���f�[�^�����擾
    Index = UBound(tblSP)

    '' �f�[�^�ǉ�
    tblSP(Index) = tblGet(1)
    
    GetSPWFData1 = FUNCTION_RETURN_SUCCESS
    
End Function

'�T�v      :�i�Ԃ�萻�i�d�l�v�e�f�[�^�Q���擾�A�����āA�擾�������i�d�l�f�[�^��ǉ�����B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :tblSP()       ,O   ,typ_TBCME022   ,���i�d�l�v�e�f�[�^�Q�e�[�u���z��
'          :tHinInf       ,I   ,tFullHinban    ,12���i��
'          :ctrlFrm       ,I   ,Form           ,�t�H�[��ID
'          :�߂�l        ,O  ,Integer          ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
'����      :05/03/01 ooba
Public Function GetSPWFData2(tblSP() As typ_TBCME022, tHinInf As tFullHinban, ctrlFrm As Form) As Integer

    Dim Index       As Integer
    Dim iRet        As Integer
    Dim tHinban(0)  As tFullHinban
    Dim tblGet()    As typ_TBCME022
    
    GetSPWFData2 = FUNCTION_RETURN_FAILURE

    tHinban(0) = tHinInf

    '' ���i�d�l�v�e�f�[�^�Q�̎擾
    iRet = DBDRV_GetTBCME022(tblGet, ctrlFrm.Name, tHinban)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function

    '' �擾�������i�d�l�v�e�f�[�^�̒ǉ�
    '' �e�[�u���f�[�^�i�[�̈�g��
    ReDim Preserve tblSP(UBound(tblSP) + 1)
    '' �e�[�u���f�[�^�����擾
    Index = UBound(tblSP)

    '' �f�[�^�ǉ�
    tblSP(Index) = tblGet(1)
    
    GetSPWFData2 = FUNCTION_RETURN_SUCCESS
    
End Function

'�T�v      :�i�Ԃ�萻�i�d�l�v�e�f�[�^�U���擾�A�����āA�擾�������i�d�l�f�[�^��ǉ�����B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :tblSP()       ,O   ,typ_TBCME026   ,���i�d�l�v�e�f�[�^�U�e�[�u���z��
'          :tHinInf       ,I   ,tFullHinban    ,12���i��
'          :ctrlFrm       ,I   ,Form           ,�t�H�[��ID
'          :�߂�l        ,O  ,Integer          ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
'����      :05/03/01 ooba
Public Function GetSPWFData6(tblSP() As typ_TBCME026, tHinInf As tFullHinban, ctrlFrm As Form) As Integer

    Dim Index       As Integer
    Dim iRet        As Integer
    Dim tHinban(0)  As tFullHinban
    Dim tblGet()    As typ_TBCME026
    
    GetSPWFData6 = FUNCTION_RETURN_FAILURE

    tHinban(0) = tHinInf

    '' ���i�d�l�v�e�f�[�^�U�̎擾
    iRet = DBDRV_GetTBCME026(tblGet, ctrlFrm.Name, tHinban)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function

    '' �擾�������i�d�l�v�e�f�[�^�̒ǉ�
    '' �e�[�u���f�[�^�i�[�̈�g��
    ReDim Preserve tblSP(UBound(tblSP) + 1)
    '' �e�[�u���f�[�^�����擾
    Index = UBound(tblSP)

    '' �f�[�^�ǉ�
    tblSP(Index) = tblGet(1)
    
    GetSPWFData6 = FUNCTION_RETURN_SUCCESS
    
End Function

'�T�v      :�i�Ԃ�萻�i�d�l�v�e�f�[�^�W���擾�A�����āA�擾�������i�d�l�f�[�^��ǉ�����B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :tblSP()       ,O   ,typ_TBCME028   ,���i�d�l�v�e�f�[�^�W�e�[�u���z��
'          :tHinInf       ,I   ,tFullHinban    ,12���i��
'          :ctrlFrm       ,I   ,Form           ,�t�H�[��ID
'          :�߂�l        ,O  ,Integer          ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
'����      :�V�K�쐬 05/06/16 ffc)tanabe
Public Function GetSPWFData8(tblSP() As typ_TBCME028, tHinInf As tFullHinban, ctrlFrm As Form) As Integer

    Dim Index       As Integer
    Dim iRet        As Integer
    Dim tHinban(0)  As tFullHinban
    Dim tblGet()    As typ_TBCME028
    
    GetSPWFData8 = FUNCTION_RETURN_FAILURE

    tHinban(0) = tHinInf

    '' ���i�d�l�v�e�f�[�^�W�̎擾
    iRet = DBDRV_GetTBCME028(tblGet, ctrlFrm.Name, tHinban)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function

    '' �擾�������i�d�l�v�e�f�[�^�̒ǉ�
    '' �e�[�u���f�[�^�i�[�̈�g��
    ReDim Preserve tblSP(UBound(tblSP) + 1)
    '' �e�[�u���f�[�^�����擾
    Index = UBound(tblSP)

    '' �f�[�^�ǉ�
    tblSP(Index) = tblGet(1)
    
    GetSPWFData8 = FUNCTION_RETURN_SUCCESS
    
End Function

'�T�v      :�����T���v�������w��"3"�T���v���`�F�b�N
'���Ұ�    :�ϐ���        ,IO ,�^                       ,����
'          :tblCrySmp     ,I   ,typ_XSDCS               ,�V�T���v���Ǘ��i�u���b�N�j�e�[�u��
'          :kensa_typ     ,I   ,chkKensaType            ,�����w������
'          :[iChkOpt]     ,I   ,Integer                 ,�T���v���`�F�b�N���샂�[�h
'          :�߂�l        ,O  ,Boolean                  ,�T���v���`�F�b�N���샂�[�h�ɂ��߂�l�̈Ӗ����قȂ�
'����      :
'           [iChkOpt = 1,2�ȊO] �F  �������A���ʒu�̑��̃T���v���Ɍ����w��"3"���w������Ă��邩�m�F
'                   �߂�l�FTRUE    �w������
'                   �@�@�@�FFALSE   �w���Ȃ�
'           [iChkOpt = 1] �F        �������A���ʒu�̑��̃T���v���Ɍ����w��"3"���w������Ă��邩�m�F
'                                   �����āA�������̌����T���v���̃T���v���敪�̎�ނ�Ԃ�
'                   �߂�l�FTRUE    �T���v���敪"B"�ł���
'                   �@�@�@�FFALSE   �T���v���敪"T"�ł���
'           [iChkOpt = 2] �F        �������A���ʒu�̑��̃T���v���Ɍ����w��"3"���w������Ă��邩�m�F
'                                   �����āA�T���v���敪"B","T"�̃T���v���̊m��敪���m�F����
'                   �߂�l�FTRUE    �T���v���敪"B","T"�̃T���v���̂����ꂩ���m�肳��Ă���
'                   �@�@�@�FFALSE   �T���v���敪"B","T"�̃T���v���̂�������m�肳��Ă��Ȃ�
'
Public Function ChkCommonKensa(tblCrySmp As typ_XSDCS, Kensa_Typ As chkKensaType, Optional iChkOpt = 1) As Boolean
    Dim iRet        As Integer
    Dim tblGet()    As typ_XSDCS
    Dim bFind       As Boolean
    Dim sqlWhere    As String
    Dim KeyItem     As String
    ''�@�����T���v�������w��"3"�`�F�b�N�@2003/09/08 Motegi ===========================> START
'    Const keyComm = "3"
    '----------------------
    Const keyComm = "1"
    
    
    ChkCommonKensa = False

    '' �����w��"3"�̃`�F�b�N
'    bFind = False
'    Select Case Kensa_Typ
'        Case CHK_OI         '' Oi
'            If tblCrySmp.CRYINDOICS = keyComm Then bFind = True: KeyItem = "CRYINDOICS"
'        Case CHK_CS         '' Cs
'            If tblCrySmp.CRYINDCSCS = keyComm Then bFind = True: KeyItem = "CRYINDCSCS"
'        Case CHK_RS         '' Rs
'            If tblCrySmp.CRYINDRSCS = keyComm Then bFind = True: KeyItem = "CRYINDRSCS"
'        Case CHK_B1         '' BMD1
'            If tblCrySmp.CRYINDB1CS = keyComm Then bFind = True: KeyItem = "CRYINDB1CS"
'        Case CHK_B2         '' BMD2
'            If tblCrySmp.CRYINDB2CS = keyComm Then bFind = True: KeyItem = "CRYINDB2CS"
'        Case CHK_B3         '' BMD3
'            If tblCrySmp.CRYINDB3CS = keyComm Then bFind = True: KeyItem = "CRYINDB3CS"
'        Case CHK_L1         '' OSF1
'            If tblCrySmp.CRYINDL1CS = keyComm Then bFind = True: KeyItem = "CRYINDL1CS"
'        Case CHK_L2         '' OSF2
'            If tblCrySmp.CRYINDL2CS = keyComm Then bFind = True: KeyItem = "CRYINDL2CS"
'        Case CHK_L3         '' OSF3
'            If tblCrySmp.CRYINDL3CS = keyComm Then bFind = True: KeyItem = "CRYINDL3CS"
'        Case CHK_L4         '' OSF4
'            If tblCrySmp.CRYINDL4CS = keyComm Then bFind = True: KeyItem = "CRYINDL4CS"
'        Case CHK_GD         '' GD
'            If tblCrySmp.CRYINDGDCS = keyComm Then bFind = True: KeyItem = "CRYINDGDCS"
'        Case CHK_LT         '' LT
'            If tblCrySmp.CRYINDTCS = keyComm Then bFind = True: KeyItem = "CRYINDTCS"
'        Case CHK_EP         '' EPD
'            If tblCrySmp.CRYINDEPCS = keyComm Then bFind = True: KeyItem = "CRYINDEPCS"
'    End Select
'    If bFind <> True Then Exit Function

    Select Case Kensa_Typ
        Case CHK_OI         '' Oi
            KeyItem = "CRYINDOICS"
        Case CHK_CS         '' Cs
            KeyItem = "CRYINDCSCS"
        Case CHK_RS         '' Rs
            KeyItem = "CRYINDRSCS"
        Case CHK_B1         '' BMD1
            KeyItem = "CRYINDB1CS"
        Case CHK_B2         '' BMD2
            KeyItem = "CRYINDB2CS"
        Case CHK_B3         '' BMD3
            KeyItem = "CRYINDB3CS"
        Case CHK_L1         '' OSF1
            KeyItem = "CRYINDL1CS"
        Case CHK_L2         '' OSF2
            KeyItem = "CRYINDL2CS"
        Case CHK_L3         '' OSF3
            KeyItem = "CRYINDL3CS"
        Case CHK_L4         '' OSF4
            KeyItem = "CRYINDL4CS"
        Case CHK_GD         '' GD
            KeyItem = "CRYINDGDCS"
        Case CHK_LT         '' LT
            KeyItem = "CRYINDTCS"
        Case CHK_EP         '' EPD
            KeyItem = "CRYINDEPCS"
    End Select

    ''�@�����T���v�������w��"3"�`�F�b�N�@2003/09/08 Motegi ===========================> END

    '' SQL�����쐬
    sqlWhere = "where XTALCS='" & tblCrySmp.XTALCS & "' "
    sqlWhere = sqlWhere + "and INPOSCS=" & tblCrySmp.INPOSCS & " "
    sqlWhere = sqlWhere + "and " & KeyItem & "='" & keyComm & "'"

    '' �����T���v���Ǘ��e�[�u���̎擾
    iRet = DBDRV_GetTBCME043(tblGet, sqlWhere, "order by SMPKBNCS")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    
    If UBound(tblGet) <> 2 Then '' 2���ȊO�̏ꍇ
        Exit Function
    End If

    Select Case iChkOpt '' ���샂�[�h�̑I��
        Case 1
            ''�@�T���v���敪"T"�̏ꍇ�A�敪"T"�T���v����\�����������̂Ŗ߂�lFalse�ŏ����I��
            If Trim(tblCrySmp.SMPKBNCS) = "T" Then Exit Function
        Case 2 '' �m��敪�`�F�b�N
            If tblGet(1).KTKBNCS = "0" And tblGet(2).KTKBNCS = "0" Then Exit Function
    End Select

    ChkCommonKensa = True
End Function
'�T�v      :�����ԍ��Ɉ�v���錋���T���v���Ǘ����擾����
'���Ұ�    :�ϐ���        ,IO ,�^                       ,����
'          :tblSmpl()     ,O   ,typ_XSDCS               ,�V�T���v���Ǘ��i�u���b�N�j�e�[�u��
'          :strCryNum     ,I   ,String                  ,�����ԍ�
'          :iMode         ,I   ,Integer                 ,�����T���v���Ǘ��X�V���[�h
'          :�߂�l        ,O  ,Integer                  ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Public Function GetSmplManage(tblSmpl() As typ_XSDCS, strCryNum As String, iMode As Integer) As Integer
    
    Dim Index       As Long
    Dim iRet        As Integer
    Dim tblGet()    As typ_XSDCS
    Dim tblTgt()    As typ_XSDCS
    Dim bFind       As Boolean
    Dim HinbanMng() As typ_TBCME041
    Dim UpHin       As tFullHinban
    Dim downHin     As tFullHinban
    Dim sKensa      As String * 1
    Dim tHinInf     As tFullHinban
    
    GetSmplManage = FUNCTION_RETURN_FAILURE
    ReDim tblGet(0)
    ReDim tblTgt(0)

    '' �����T���v���Ǘ��e�[�u����������
    RemoveAll_CrystalSampleManage tblSmpl

    '' �����T���v���Ǘ��e�[�u���̎擾
    iRet = DBDRV_GetTBCME043(tblGet, "where XTALCS='" & strCryNum & "' ", "order by INPOSCS, SMPKBNCS")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function
    
    '' �ΏہA�����T���v���Ǘ��e�[�u���̔����o��
    For Index = 1 To UBound(tblGet)
        bFind = False
        '' �擾���[�h�`�F�b�N
'�V����يǗ��Ή��@2003/09/08 Motegi ========================================> �ύX�J�n
'        Select Case iMode
'            Case MODE_GETSMPL_FTIR      '' FTIR(Oi,Cs)
'                If InStr(CODE_KENSA, tblGet(Index).CRYINDOICS) <> 0 Or _
'                   InStr(CODE_KENSA, tblGet(Index).CRYINDCSCS) <> 0 Then
'                    '' ���ʌ������ڃ`�F�b�N
'                    If ChkCommonKensa(tblGet(Index), CHK_OI) And _
'                       ChkCommonKensa(tblGet(Index), CHK_CS) Then
'                        bFind = False
'                    Else
'                        If ChkCommonKensa(tblGet(Index), CHK_OI, 2) Or _
'                           ChkCommonKensa(tblGet(Index), CHK_CS, 2) Then tblGet(Index).KTKBNCS = "1"
''����̌������@���ݒ肳��Ă��邽�߁A�d�l�L���̃`�F�b�N�͕s�v
''                        If (GetReferHinban(tHinInf, tblGet(Index), CHK_OI) = FUNCTION_RETURN_SUCCESS) Or _
''                           (GetReferHinban(tHinInf, tblGet(Index), CHK_CS) = FUNCTION_RETURN_SUCCESS) Then
'                            bFind = True
''                        End If
'                    End If
'                End If
'            Case MODE_GETSMPL_GFA       '' GFA(Oi)
'                If InStr(CODE_KENSA, tblGet(Index).CRYINDOICS) <> 0 Then
'                    '' ���ʌ������ڃ`�F�b�N
'                    If ChkCommonKensa(tblGet(Index), CHK_OI) Then
'                        bFind = False
'                    Else
'                        If ChkCommonKensa(tblGet(Index), CHK_OI, 2) Then tblGet(Index).KTKBNCS = "1"
''                        If GetReferHinban(tHinInf, tblGet(Index), CHK_OI) = FUNCTION_RETURN_SUCCESS Then
'                            bFind = True
''                        End If
'                    End If
'                End If
'            Case MODE_GETSMPL_RS        '' ��R
'                If InStr(CODE_KENSA, tblGet(Index).CRYINDRSCS) <> 0 Then
'                    If ChkCommonKensa(tblGet(Index), CHK_RS) Then
'                        bFind = False
'                    Else
'                        If ChkCommonKensa(tblGet(Index), CHK_RS, 2) Then tblGet(Index).KTKBNCS = "1"
''                        If GetReferHinban(tHinInf, tblGet(Index), CHK_RS) = FUNCTION_RETURN_SUCCESS Then
'                            bFind = True
''                        End If
'                    End If
'            Case MODE_GETSMPL_BMD       '' BMD
'                If InStr(CODE_KENSA, tblGet(Index).CRYINDB1CS) <> 0 Or _
'                   InStr(CODE_KENSA, tblGet(Index).CRYINDB2CS) <> 0 Or _
'                   InStr(CODE_KENSA, tblGet(Index).CRYINDB3CS) <> 0 Then
'                    bFind = True
'                End If
'            Case MODE_GETSMPL_OSF       '' OSF
'                If InStr(CODE_KENSA, tblGet(Index).CRYINDL1CS) <> 0 Or _
'                   InStr(CODE_KENSA, tblGet(Index).CRYINDL2CS) <> 0 Or _
'                   InStr(CODE_KENSA, tblGet(Index).CRYINDL3CS) <> 0 Or _
'                   InStr(CODE_KENSA, tblGet(Index).CRYINDL4CS) <> 0 Then
'                    bFind = True
'                End If
'            Case MODE_GETSMPL_GD        '' GD
'                If InStr(CODE_KENSA, tblGet(Index).CRYINDGDCS) <> 0 Then
'                    If ChkCommonKensa(tblGet(Index), CHK_GD) Then
'                        bFind = False
'                    Else
'                        If ChkCommonKensa(tblGet(Index), CHK_GD, 2) Then tblGet(Index).KTKBNCS = "1"
'                        If GetReferHinban(tHinInf, tblGet(Index), CHK_GD) = FUNCTION_RETURN_SUCCESS Then
'                            bFind = True
'                        End If
'                    End If
'                End If
'            Case MODE_GETSMPL_LT        '' ���C�t�^�C��
'                If InStr(CODE_KENSA, tblGet(Index).CRYINDTCS) <> 0 Then
'                    If ChkCommonKensa(tblGet(Index), CHK_LT) Then
'                        bFind = False
'                    Else
'                        If ChkCommonKensa(tblGet(Index), CHK_LT, 2) Then tblGet(Index).KTKBNCS = "1"
'                        If GetReferHinban(tHinInf, tblGet(Index), CHK_LT) = FUNCTION_RETURN_SUCCESS Then
'                            bFind = True
'                        End If
'                    End If
'                End If
'            Case MODE_GETSMPL_EPD       '' EPD
'                If InStr(CODE_KENSA, tblGet(Index).CRYINDEPCS) <> 0 Then
'                    If ChkCommonKensa(tblGet(Index), CHK_EP) Then
'                        bFind = False
'                    Else
'                        If ChkCommonKensa(tblGet(Index), CHK_EP, 2) Then tblGet(Index).KTKBNCS = "1"
'                        If GetReferHinban(tHinInf, tblGet(Index), CHK_EP) = FUNCTION_RETURN_SUCCESS Then
'                            bFind = True
'                        End If
'                    End If
'                End If
'            Case Else                   '' ���̑�
'                Exit Function
'        End Select
'-------------------------------------
        Select Case iMode
            Case MODE_GETSMPL_FTIR      '' FTIR(Oi,Cs)
                If InStr(CODE_KENSA, tblGet(Index).CRYINDOICS) = 1 Or _
                   InStr(CODE_KENSA, tblGet(Index).CRYINDCSCS) = 1 Then
                    bFind = True
                End If
            Case MODE_GETSMPL_GFA       '' GFA(Oi)
                If InStr(CODE_KENSA, tblGet(Index).CRYINDOICS) = 1 Then
                    bFind = True
                End If
            Case MODE_GETSMPL_RS        '' ��R
                If InStr(CODE_KENSA, tblGet(Index).CRYINDRSCS) = 1 Then
                    bFind = True
                End If
            Case MODE_GETSMPL_BMD       '' BMD
                If InStr(CODE_KENSA, tblGet(Index).CRYINDB1CS) = 1 Or _
                   InStr(CODE_KENSA, tblGet(Index).CRYINDB2CS) = 1 Or _
                   InStr(CODE_KENSA, tblGet(Index).CRYINDB3CS) = 1 Then
                    bFind = True
                End If
            Case MODE_GETSMPL_OSF       '' OSF
                If InStr(CODE_KENSA, tblGet(Index).CRYINDL1CS) = 1 Or _
                   InStr(CODE_KENSA, tblGet(Index).CRYINDL2CS) = 1 Or _
                   InStr(CODE_KENSA, tblGet(Index).CRYINDL3CS) = 1 Or _
                   InStr(CODE_KENSA, tblGet(Index).CRYINDL4CS) = 1 Then
                    bFind = True
                End If
            Case MODE_GETSMPL_GD        '' GD
                If InStr(CODE_KENSA, tblGet(Index).CRYINDGDCS) = 1 Then
                    bFind = True
                End If
            Case MODE_GETSMPL_LT        '' ���C�t�^�C��
                If InStr(CODE_KENSA, tblGet(Index).CRYINDTCS) = 1 Then
                    bFind = True
                End If
            Case MODE_GETSMPL_EPD       '' EPD
                If InStr(CODE_KENSA, tblGet(Index).CRYINDEPCS) = 1 Then
                    bFind = True
                End If
            
            '2009/08 SUMCO Akizuki�@X��������ѓ��́@�쐬�ɔ����ǉ�
            '[1:�����w������]
            Case MODE_GETSMPL_X       '' X��
                If InStr(CODE_KENSA, tblGet(Index).CRYINDXCS) = 1 Then
                    bFind = True
                End If

            'Add Start 2010/12/17 SMPK Miyata
            Case MODE_GETSMPL_CUDECO    '' Cu-deco
                If InStr(CODE_KENSA, tblGet(Index).CRYINDCCS) = 1 Or _
                   InStr(CODE_KENSA, tblGet(Index).CRYINDCJCS) = 1 Or _
                   InStr(CODE_KENSA, tblGet(Index).CRYINDCJLTCS) = 1 Or _
                   InStr(CODE_KENSA, tblGet(Index).CRYINDCJ2CS) = 1 Then
                    bFind = True
                End If
            'Add End   2010/12/17 SMPK Miyata

            Case Else                   '' ���̑�
                Exit Function
        End Select
'�V����يǗ��Ή��@2003/09/08 Motegi ========================================> �ύX�I��
        
        '' ���������w��������ꍇ�A�����T���v���Ǘ��e�[�u���ɏo�͂���
        If bFind = True Then
            If Add_CrystalSampleManage(tblSmpl, tblGet(Index)) <> FUNCTION_RETURN_SUCCESS Then
                Exit Function
            End If
        End If
    Next Index
    
    GetSmplManage = FUNCTION_RETURN_SUCCESS

End Function

'�T�v      :�����T���v���Ǘ����X�V����
'���Ұ�    :�ϐ���        ,IO ,�^                       ,����
'          :tblCrySmpMan  ,I   ,typ_XSDCS               ,�V�T���v���Ǘ��i�u���b�N�j�e�[�u���X�V�p�����[�^
'          :strCryNum     ,I   ,String                  ,�����ԍ�
'          :iIngotPos     ,I   ,Integer                 ,�������ʒu
'          :strSmpKbn     ,I   ,String                  ,�T���v���敪
'          :iSmpNo        ,I   ,Long                    ,�T���v��No.    Integer��Long 6���Ή� SETsw kubota
'          :iMode         ,I   ,Integer                 ,�����T���v���Ǘ��X�V���[�h
'          :[iOption]     ,I   ,Integer                 ,�����T���v���Ǘ��X�V���[�h�I�v�V����
'          :�߂�l        ,O  ,Integer                  ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Public Function UpdateTbl_CrySmpManage(tblCrySmpMan As typ_XSDCS, strCryNum As String, iIngotpos As Integer, strSmpKbn As String, iSmpNo As Long, iMode As Integer, Optional iOption As Integer = 0) As Integer
    Dim iRet        As Integer
    Dim sqlWhere    As String
''�@�����w���ύX�@2003/09/10 Motegi ==========================> START
'    Dim tblGet()    As typ_XSDCS
'    Dim strKeySmpKbn As String
'    Dim bUpdateFlag As Boolean
    Dim sqlUpdate   As String
''�@�����w���ύX�@2003/09/10 Motegi ==========================> END
    
    UpdateTbl_CrySmpManage = FUNCTION_RETURN_FAILURE
''�@�����w���ύX�@2003/09/10 Motegi ==========================> START
'    bUpdateFlag = False
    
'    If Trim(strSmpKbn) = "B" Then
'        strKeySmpKbn = "T"
'    Else
'        strKeySmpKbn = "B"
'    End If
    
'    ReDim tblGet(0)
'    '' �����T���v���Ǘ��e�[�u���̎擾
'    sqlWhere = " where XTALCS='" & strCryNum & "' " & "and INPOSCS=" & iIngotPos & _
'               " and SMPKBNCS='" & strKeySmpKbn & "' "
'    iRet = DBDRV_GetTBCME043(tblGet, sqlWhere, "order by INPOSCS, SMPKBNCS")
'    If (iRet = FUNCTION_RETURN_SUCCESS) And (UBound(tblGet) > 0) Then
        '' �T���v����"3�i���ʁj"�̌����w�����ݒ肳��Ă���ꍇ�A
        '' �������ꂽ�����T���v���Ǘ��e�[�u���̌������т��X�V����
'        Select Case iMode
'        Case MODE_GETSMPL_FTIR      '' FTIR(Oi,Cs)
'            If (tblCrySmpMan.CRYINDOICS = tblGet(1).CRYINDOICS) And (tblCrySmpMan.CRYINDOICS = "3") And (iOption = 1) Then
'                tblGet(1).CRYRESOICS = tblCrySmpMan.CRYRESOICS
'                bUpdateFlag = True
'            End If
'            If (tblCrySmpMan.CRYINDCSCS = tblGet(1).CRYINDCSCS) And (tblCrySmpMan.CRYINDCSCS = "3") And (iOption = 2) Then
'                tblGet(1).CRYRESCSCS = tblCrySmpMan.CRYRESCSCS
'                bUpdateFlag = True
'            End If
'        Case MODE_GETSMPL_GFA       '' GFA(Oi)
'            If (tblCrySmpMan.CRYINDOICS = tblGet(1).CRYINDOICS) And (tblCrySmpMan.CRYINDOICS = "3") Then
'                tblGet(1).CRYRESOICS = tblCrySmpMan.CRYRESOICS
'                bUpdateFlag = True
'            End If
'        Case MODE_GETSMPL_RS        '' ��R
'            If (tblCrySmpMan.CRYINDRSCS = tblGet(1).CRYINDRSCS) And (tblCrySmpMan.CRYINDRSCS = "3") Then
'                tblGet(1).CRYRESRS1CS = tblCrySmpMan.CRYRESRS1CS
'                bUpdateFlag = True
'            End If
'        Case MODE_GETSMPL_BMD       '' BMD
'            If (tblCrySmpMan.CRYINDB1CS = tblGet(1).CRYINDB1CS) And (tblCrySmpMan.CRYINDB1CS = "3") And (iOption = 1) Then
'                tblGet(1).CRYRESB1CS = tblCrySmpMan.CRYRESB1CS
'                bUpdateFlag = True
'            End If
'            If (tblCrySmpMan.CRYINDB2CS = tblGet(1).CRYINDB2CS) And (tblCrySmpMan.CRYINDB2CS = "3") And (iOption = 2) Then
'                tblGet(1).CRYRESB2CS = tblCrySmpMan.CRYRESB2CS
'                bUpdateFlag = True
'            End If
'            If (tblCrySmpMan.CRYINDB3CS = tblGet(1).CRYINDB3CS) And (tblCrySmpMan.CRYINDB3CS = "3") And (iOption = 3) Then
'                tblGet(1).CRYRESB3CS = tblCrySmpMan.CRYRESB3CS
'                bUpdateFlag = True
'            End If
'        Case MODE_GETSMPL_OSF       '' OSF
'            If (tblCrySmpMan.CRYINDL1CS = tblGet(1).CRYINDL1CS) And (tblCrySmpMan.CRYINDL1CS = "3") And (iOption = 1) Then
'                tblGet(1).CRYRESL1CS = tblCrySmpMan.CRYRESL1CS
'                bUpdateFlag = True
'            End If
'            If (tblCrySmpMan.CRYINDL2CS = tblGet(1).CRYINDL2CS) And (tblCrySmpMan.CRYINDL2CS = "3") And (iOption = 2) Then
'                tblGet(1).CRYRESL2CS = tblCrySmpMan.CRYRESL2CS
'                bUpdateFlag = True
'            End If
'            If (tblCrySmpMan.CRYINDL3CS = tblGet(1).CRYINDL3CS) And (tblCrySmpMan.CRYINDL3CS = "3") And (iOption = 3) Then
'                tblGet(1).CRYRESL3CS = tblCrySmpMan.CRYRESL3CS
'                bUpdateFlag = True
'            End If
'            If (tblCrySmpMan.CRYINDL4CS = tblGet(1).CRYINDL4CS) And (tblCrySmpMan.CRYINDL4CS = "3") And (iOption = 4) Then
'                tblGet(1).CRYRESL4CS = tblCrySmpMan.CRYRESL4CS
'                bUpdateFlag = True
'            End If
'        Case MODE_GETSMPL_GD        '' GD
'            If (tblCrySmpMan.CRYINDGDCS = tblGet(1).CRYINDGDCS) And (tblCrySmpMan.CRYINDGDCS = "3") Then
'                tblGet(1).CRYINDGDCS = tblCrySmpMan.CRYINDGDCS
'                bUpdateFlag = True
'            End If
'        Case MODE_GETSMPL_LT        '' ���C�t�^�C��
'            If (tblCrySmpMan.CRYINDTCS = tblGet(1).CRYINDTCS) And (tblCrySmpMan.CRYINDTCS = "3") Then
'                tblGet(1).CRYINDTCS = tblCrySmpMan.CRYINDTCS
'                bUpdateFlag = True
'            End If
'        Case MODE_GETSMPL_EPD       '' EPD
'            If (tblCrySmpMan.CRYINDEPCS = tblGet(1).CRYINDEPCS) And (tblCrySmpMan.CRYINDEPCS = "3") Then
'                tblGet(1).CRYINDEPCS = tblCrySmpMan.CRYINDEPCS
'                bUpdateFlag = True
'            End If
'        End Select
'------------------------------
    With tblCrySmpMan
'2009/08�@SUMCO Akizuki �T���v���Ǘ��X�V�����ɁA�ύX���̔��f���Ȃ��������߁A�ǉ�
'>>>>>
        sqlUpdate = "update XSDCS set "
        sqlUpdate = sqlUpdate & "KSTAFFCS = '" & .KSTAFFCS & "' ,"          '�X�V�Ј�ID
        sqlUpdate = sqlUpdate & "KDAYCS = SYSDATE ,"                        '�X�V���t
        sqlUpdate = sqlUpdate & "SNDKDWHCS = '0' ,"                         '���M�t���O(DWH)
'<<<<<

        Select Case iMode
            Case MODE_GETSMPL_FTIR      '' FTIR(Oi,Cs)
                If iOption = 1 Then
                    sqlUpdate = sqlUpdate & "CRYRESOICS = '" & .CRYRESOICS & "' "           ' �����������сiOi)
                    sqlWhere = "CRYSMPLIDOICS = " & iSmpNo
                ElseIf iOption = 2 Then
                    sqlUpdate = sqlUpdate & "CRYRESCSCS = '" & .CRYRESCSCS & "' "           ' �����������сiCs)
                    sqlWhere = "CRYSMPLIDCSCS = " & iSmpNo
                End If
            
            Case MODE_GETSMPL_GFA       '' GFA(Oi)
                sqlUpdate = sqlUpdate & "CRYRESOICS = '" & .CRYRESOICS & "' "               ' �����������сiOi)
                sqlWhere = "CRYSMPLIDOICS = " & iSmpNo
            
            Case MODE_GETSMPL_RS        '' ��R
                sqlUpdate = sqlUpdate & "CRYRESRS1CS = '" & .CRYRESRS1CS & "' "             ' �����������сiRs)
                sqlWhere = "CRYSMPLIDRSCS = " & iSmpNo
            
            Case MODE_GETSMPL_BMD       '' BMD
                If iOption = 1 Then
                    sqlUpdate = sqlUpdate & "CRYRESB1CS = '" & .CRYRESB1CS & "' "           ' �����������сiBMD1)
                    sqlWhere = "CRYSMPLIDB1CS = " & iSmpNo
                ElseIf iOption = 2 Then
                    sqlUpdate = sqlUpdate & "CRYRESB2CS = '" & .CRYRESB2CS & "' "           ' �����������сiBMD2)
                    sqlWhere = "CRYSMPLIDB2CS = " & iSmpNo
                ElseIf iOption = 3 Then
                    sqlUpdate = sqlUpdate & "CRYRESB3CS = '" & .CRYRESB3CS & "' "           ' �����������сiBMD3)
                    sqlWhere = "CRYSMPLIDB3CS = " & iSmpNo
                End If
            
            Case MODE_GETSMPL_OSF       '' OSF
                If iOption = 1 Then
                    sqlUpdate = sqlUpdate & "CRYRESL1CS = '" & .CRYRESL1CS & "' "           ' �����������сiOSF1)
                    sqlWhere = "CRYSMPLIDL1CS = " & iSmpNo
                ElseIf iOption = 2 Then
                    sqlUpdate = sqlUpdate & "CRYRESL2CS = '" & .CRYRESL2CS & "' "           ' �����������сiOSF2)
                    sqlWhere = "CRYSMPLIDL2CS = " & iSmpNo
                ElseIf iOption = 3 Then
                    sqlUpdate = sqlUpdate & "CRYRESL3CS = '" & .CRYRESL3CS & "' "           ' �����������сiOSF3)
                    sqlWhere = "CRYSMPLIDL3CS = " & iSmpNo
                ElseIf iOption = 4 Then
                    sqlUpdate = sqlUpdate & "CRYRESL4CS = '" & .CRYRESL4CS & "' "           ' �����������сiOSF4)
                    sqlWhere = "CRYSMPLIDL4CS = " & iSmpNo
                End If
            
            Case MODE_GETSMPL_GD        '' GD
                sqlUpdate = sqlUpdate & "CRYRESGDCS = '" & .CRYRESGDCS & "' "               ' �����������сiGD)
                sqlWhere = "CRYSMPLIDGDCS = " & iSmpNo
            
            Case MODE_GETSMPL_LT        '' ���C�t�^�C��
                sqlUpdate = sqlUpdate & "CRYRESTCS = '" & .CRYRESTCS & "', "                 ' �����������сiLT)
                                        '' ���C�t�^�C��(10�����Z)
                sqlUpdate = sqlUpdate & "CRYREST10CS = '" & .CRYREST10CS & "' "                 ' �����������сiLT)
                
                sqlWhere = "CRYSMPLIDTCS = " & iSmpNo
            
            Case MODE_GETSMPL_EPD       '' EPD
                sqlUpdate = sqlUpdate & "CRYRESEPCS = '" & .CRYRESEPCS & "' "               ' �����������сiEPD)
                sqlWhere = "CRYSMPLIDEPCS = " & iSmpNo
            
            '2009/08 Akizuki
            Case MODE_GETSMPL_X         '' X��
                sqlUpdate = sqlUpdate & "CRYRESXCS = '" & .CRYRESXCS & "' "                 ' �����������сiX��)
                sqlWhere = "CRYSMPLIDXCS = " & iSmpNo
        
            'Add Start 2011/01/07 SMPK Miyata
            Case MODE_GETSMPL_CUDECO    '' Cu-deco
                If iOption = 1 Then
                    sqlUpdate = sqlUpdate & "CRYRESCCS = '" & .CRYRESCCS & "' "             ' �����������сiC)
                    sqlWhere = "CRYSMPLIDCCS = " & iSmpNo
                
                ElseIf iOption = 2 Then
                    sqlUpdate = sqlUpdate & "CRYRESCJCS = '" & .CRYRESCJCS & "' "           ' �����������сiCJ)
                    sqlWhere = "CRYSMPLIDCJCS = " & iSmpNo

                ElseIf iOption = 3 Then
                    sqlUpdate = sqlUpdate & "CRYRESCJLTCS = '" & .CRYRESCJLTCS & "' "       ' �����������сiCJLT)
                    sqlWhere = "CRYSMPLIDCJLTCS = " & iSmpNo

                ElseIf iOption = 4 Then
                    sqlUpdate = sqlUpdate & "CRYRESCJ2CS = '" & .CRYRESCJ2CS & "' "         ' �����������сiCJ2)
                    sqlWhere = "CRYSMPLIDCJ2CS = " & iSmpNo
                End If
            'Add End   2011/01/07 SMPK Miyata
        
        End Select
        
    End With
                
''�@�����w���ύX�@2003/09/10 Motegi ==========================> END

''�@�����w���ύX�@2003/09/10 Motegi ==========================> �폜START
        '' �����T���v���Ǘ��e�[�u���̍X�V
'        If bUpdateFlag = True Then
'            '' �X�V�����̍쐬
'            sqlWhere = " where XTALCS='" & strCryNum & "' " & " and INPOSCS=" & iIngotPos & _
'                       " and SMPKBNCS='" & strKeySmpKbn & "' " & " and REPSMPLIDCS=" & tblGet(1).REPSMPLIDCS
'            '' �����T���v���Ǘ��e�[�u���̍X�V
'            iRet = DBDRV_UpdateTBCME043(tblCrySmpMan, sqlWhere)
'            If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
'        End If
'    End If
''�@�����w���ύX�@2003/09/10 Motegi ==========================> �폜END

    '' �X�V�����̍쐬
''�@�����w���ύX�@2003/09/10 Motegi ==========================> START
    sqlUpdate = sqlUpdate & " where XTALCS='" & strCryNum & "' " & " and " & sqlWhere
    
    '' �����T���v���Ǘ��e�[�u���̍X�V
    iRet = DBDRV_UpdateXSDCS(sqlUpdate)
''�@�����w���ύX�@2003/09/10 Motegi ==========================> END
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function

    UpdateTbl_CrySmpManage = FUNCTION_RETURN_SUCCESS

End Function
'�T�v      :�d�l�����邩�𒲂ׂ�
'���Ұ�    :�ϐ���        ,IO ,�^                    ,����
'          :tHinInf       ,O  ,tFullHinban           ,�����i��
'          :kensa_Typ     ,I  ,chkKensaType          ,������ʂ��ΏۂƂ��錟��
'          :�߂�l        ,O  ,Boolean               ,True�F�d�l����@False�F�d�l�Ȃ�
'����      :
Public Function ChkSpExistence(tHinInf As tFullHinban, Kensa_Typ As chkKensaType) As Boolean
    Dim chkSyo      As String * 1
    Dim chkPoint    As String * 1
    Dim iPoint      As Integer
    Dim idx1        As Integer
    Dim idx2        As Integer
    Dim idx3        As Integer
    Dim bFind1      As Boolean
    Dim bFind2      As Boolean
    Dim bFind3      As Boolean

    ChkSpExistence = False
    bFind1 = False
    bFind2 = False
    bFind3 = False

    '' �i�Ԏ�ނ̓���
    If (Trim(tHinInf.hinban) = "") Or (Trim(tHinInf.hinban) = "Z") Then '' ��i�ԁA�y�i�Ԃ̏ꍇ
        '' �d�l�̌����w���𒲂ׂ�
        If (Kensa_Typ = CHK_RS) Or (Kensa_Typ = CHK_CS) Then
            ChkSpExistence = True
            Exit Function
        Else
            Exit Function
        End If
    ElseIf (Trim(tHinInf.hinban) = "G") Then  '' �f�i�Ԃ̏ꍇ
        '' �d�l�̌����w���𒲂ׂ�
        If (Kensa_Typ = CHK_RS) Or (Kensa_Typ = CHK_CS) Or (Kensa_Typ = CHK_OI) Or (Kensa_Typ = CHK_LT) Then
            ChkSpExistence = True
            Exit Function
        Else
            Exit Function
        End If
    Else        '' ���̑��A�i�Ԃ̏ꍇ
        '' �Ώەi�Ԃ̐��i�d�l�̒T��
        For idx1 = 0 To UBound(tbl_PrSpSXLData1) - 1
            If (tbl_PrSpSXLData1(idx1).hinban = tHinInf.hinban) And (tbl_PrSpSXLData1(idx1).mnorevno = tHinInf.mnorevno) And _
               (tbl_PrSpSXLData1(idx1).factory = tHinInf.factory) And (tbl_PrSpSXLData1(idx1).opecond = tHinInf.opecond) Then
                bFind1 = True
                Exit For
            End If
        Next idx1
        For idx2 = 0 To UBound(tbl_PrSpSXLData2) - 1
            If (tbl_PrSpSXLData2(idx2).hinban = tHinInf.hinban) And (tbl_PrSpSXLData2(idx2).mnorevno = tHinInf.mnorevno) And _
               (tbl_PrSpSXLData2(idx2).factory = tHinInf.factory) And (tbl_PrSpSXLData2(idx2).opecond = tHinInf.opecond) Then
                bFind2 = True
                Exit For
            End If
        Next idx2
        For idx3 = 0 To UBound(tbl_PrSpSXLData3) - 1
            If (tbl_PrSpSXLData3(idx3).hinban = tHinInf.hinban) And (tbl_PrSpSXLData3(idx3).mnorevno = tHinInf.mnorevno) And _
               (tbl_PrSpSXLData3(idx3).factory = tHinInf.factory) And (tbl_PrSpSXLData3(idx3).opecond = tHinInf.opecond) Then
                bFind3 = True
                Exit For
            End If
        Next idx3
        
        '' �d�l�̌����w���𒲂ׂ�
        Select Case Kensa_Typ
            Case CHK_OI         '' Oi
                If bFind2 = False Then Exit Function
                With tbl_PrSpSXLData2(idx2)
                    chkSyo = .HSXONHWS
                    chkPoint = .HSXONSPT
                End With
            Case CHK_CS         '' Cs
                If bFind2 = False Then Exit Function
                With tbl_PrSpSXLData2(idx2)
                    chkSyo = .HSXCNHWS
                    chkPoint = .HSXCNSPT
                End With
            Case CHK_RS         '' Rs
                If bFind1 = False Then Exit Function
                With tbl_PrSpSXLData1(idx1)
                    chkSyo = .HSXRHWYS
                    chkPoint = .HSXRSPOT
                End With
            Case CHK_B1         '' BMD1
                If bFind3 = False Then Exit Function
                With tbl_PrSpSXLData3(idx3)
                    chkSyo = .HSXBM1HS
                    chkPoint = .HSXBM1ST
                End With
            Case CHK_B2         '' BMD2
                If bFind3 = False Then Exit Function
                With tbl_PrSpSXLData3(idx3)
                    chkSyo = .HSXBM2HS
                    chkPoint = .HSXBM2ST
                End With
            Case CHK_B3         '' BMD3
                If bFind3 = False Then Exit Function
                With tbl_PrSpSXLData3(idx3)
                    chkSyo = .HSXBM3HS
                    chkPoint = .HSXBM3ST
                End With
            Case CHK_L1         '' OSF1
                If bFind3 = False Then Exit Function
                With tbl_PrSpSXLData3(idx3)
                    chkSyo = .HSXOF1HS
                    chkPoint = .HSXOF1ST
                End With
            Case CHK_L2         '' OSF2
                If bFind3 = False Then Exit Function
                With tbl_PrSpSXLData3(idx3)
                    chkSyo = .HSXOF2HS
                    chkPoint = .HSXOF2ST
                End With
            Case CHK_L3         '' OSF3
                If bFind3 = False Then Exit Function
                With tbl_PrSpSXLData3(idx3)
                    chkSyo = .HSXOF3HS
                    chkPoint = .HSXOF3ST
                End With
            Case CHK_L4         '' OSF4
                If bFind3 = False Then Exit Function
                With tbl_PrSpSXLData3(idx3)
                    chkSyo = .HSXOF4HS
                    chkPoint = .HSXOF4ST
                End With
            Case CHK_GD         '' GD
                If bFind3 = False Then Exit Function
                With tbl_PrSpSXLData3(idx3)
                    If (.HSXDENHS = "H" Or .HSXDENHS = "S") Or _
                       (.HSXDVDHS = "H" Or .HSXDVDHS = "S") Or _
                       (.HSXLDLHS = "H" Or .HSXLDLHS = "S") Then
                        chkSyo = "S"
                        chkPoint = .HSXGDSPT
                    Else
                        Exit Function
                    End If
                End With
            Case CHK_LT         '' LT
                If bFind2 = False Then Exit Function
                With tbl_PrSpSXLData2(idx2)
                    chkSyo = .HSXLTHWS
                    chkPoint = .HSXLTSPT
                End With
            Case CHK_EP         '' EPD
                '' EPD �͒ʏ�i�Ԃł���Ύd�l����
                ChkSpExistence = True
                Exit Function
            Case Else           '' ���̑�
                Exit Function
        End Select
    
        iPoint = GetMeasureNum(chkPoint)
        '' �ۏؕ��@�Q���A����_�𒲂ׂ�
        If (chkSyo = "H" Or chkSyo = "S") And (iPoint > 0) Then
            ChkSpExistence = True
        End If
    End If

End Function
'�T�v      :�^����ꂽ�i�Ԃ̑���_����Ԃ� �i��R��OI�̂ݑΉ��j
'���Ұ�    :�ϐ���        ,IO ,�^                    ,����
'          :KensaTyp    ,I  ,chkKensaType          ,�������
'          :tHinInf       ,I  ,tfullhinban           ,��i��
'          :�߂�l        ,O  ,integer       ,����_����Ԃ�
'����      :
Private Function GetMark(Kensa_Typ As chkKensaType, tHinInf As tFullHinban) As Integer
    Dim chkSyo      As String * 1
    Dim chkPoint    As String * 1
    Dim iPoint      As Integer
    Dim idx        As Integer
    Dim bFind      As Boolean
    
    bFind = False
    
    GetMark = -1
    
    '' �i�Ԏ�ނ̓���
    If (Trim(tHinInf.hinban) = "") Or (Trim(tHinInf.hinban) = "Z") Then '' ��i�ԁA�y�i�Ԃ̏ꍇ
        '' �d�l�̌����w���𒲂ׂ�
        If (Kensa_Typ = CHK_RS) Then
            GetMark = 3
            Exit Function
        ElseIf (Kensa_Typ = CHK_OI) Then
            GetMark = 1
            Exit Function
        Else
            GetMark = -1
            Exit Function
        End If
    ElseIf (Trim(tHinInf.hinban) = "G") Then  '' �f�i�Ԃ̏ꍇ
        '' �d�l�̌����w���𒲂ׂ�
        If (Kensa_Typ = CHK_RS) Then
            GetMark = 3
            Exit Function
        ElseIf (Kensa_Typ = CHK_OI) Then
            GetMark = 3
            Exit Function
        Else
            GetMark = -1
            Exit Function
        End If
    Else        '' ���̑��A�i�Ԃ̏ꍇ
        
        '' �d�l�̌����w���𒲂ׂ�
        Select Case Kensa_Typ
        Case CHK_OI         '' Oi
            For idx = 0 To UBound(tbl_PrSpSXLData2) - 1
                If (tbl_PrSpSXLData2(idx).hinban = tHinInf.hinban) And (tbl_PrSpSXLData2(idx).mnorevno = tHinInf.mnorevno) And _
                   (tbl_PrSpSXLData2(idx).factory = tHinInf.factory) And (tbl_PrSpSXLData2(idx).opecond = tHinInf.opecond) Then
                    bFind = True
                    Exit For
                End If
            Next idx
            
    
            If bFind = False Then Exit Function
            With tbl_PrSpSXLData2(idx)
                chkSyo = .HSXONHWS
                chkPoint = .HSXONSPT
            End With
        Case CHK_RS         '' Rs
            For idx = 0 To UBound(tbl_PrSpSXLData1) - 1
                If (tbl_PrSpSXLData1(idx).hinban = tHinInf.hinban) And (tbl_PrSpSXLData1(idx).mnorevno = tHinInf.mnorevno) And _
                   (tbl_PrSpSXLData1(idx).factory = tHinInf.factory) And (tbl_PrSpSXLData1(idx).opecond = tHinInf.opecond) Then
                    bFind = True
                    Exit For
                End If
            Next idx
        
            If bFind = False Then Exit Function
            With tbl_PrSpSXLData1(idx)
                chkSyo = .HSXRHWYS
                chkPoint = .HSXRSPOT
            End With
        End Select
    
        iPoint = GetMeasureNum(chkPoint)
        '' �ۏؕ��@�Q���A����_�𒲂ׂ�
'2002/02/14 S.Sano        If (chkSyo = "H") Then
            GetMark = iPoint
'2002/02/14 S.Sano        Else
'2002/02/14 S.Sano            GetMark = -1
'2002/02/14 S.Sano        End If
    End If
    
End Function
'�T�v      :����_�������������̕i�Ԃ�Ԃ�
'���Ұ�    :�ϐ���        ,IO ,�^                    ,����
'          :KensaTyp    ,I  ,chkKensaType          ,�������
'          :UpHin       ,I  ,tFullHinban           ,��i��
'          :DwHin       ,I  ,tFullHinban           ,���i��
'          :SmpKbn      ,I  ,string                ,�T���v���敪
'          :�߂�l        ,O  ,tFullHinban       ,�㉺�i�Ԃ̂ǂ��炩��Ԃ�
'����      :
Private Function GetManyMark(KensaTyp As chkKensaType, UpHin As tFullHinban, DwHin As tFullHinban, SMPKBN As String) As tFullHinban
    
    Dim UpMark As Integer, DwMark As Integer

    UpMark = GetMark(KensaTyp, UpHin)
    DwMark = GetMark(KensaTyp, DwHin)
    If UpMark = DwMark Then                ' ������������T���v���敪�̕���
        If Trim$(SMPKBN) = "T" Then
            GetManyMark = DwHin
        Else
            GetManyMark = UpHin
        End If
    ElseIf UpMark > DwMark Then
        GetManyMark = UpHin
    Else
        GetManyMark = DwHin
    End If
    
End Function
'�T�v      :�Q�ƕi�Ԃ��擾����
'���Ұ�    :�ϐ���        ,IO ,�^                   ,����
'          :tHinInf       ,O  ,tFullHinban          ,�Q�Ƃ��ׂ��i��
'          :tblCrySmp     ,I  ,typ_XSDCS            ,�����T���v���Ǘ��e�[�u��
'          :kensa_Typ     ,I  ,chkKensaType         ,������ʂ��ΏۂƂ��錟��
'          :�߂�l        ,O  ,FUNCTION_RETURN      ,���������F�Q�ƕi�Ԃ�����@�������s�F�Q�Ƃ��ׂ��i�Ԃ͂Ȃ�
'����      :
Public Function GetReferHinban(tHinInf As tFullHinban, tblCrySmp As typ_XSDCS, Kensa_Typ As chkKensaType) As FUNCTION_RETURN
    Dim UpHin As tFullHinban
    Dim downHin As tFullHinban
    Dim isBtm As Boolean
    Dim chkShiji As String * 1
    Dim LtHin   As tFullHinban
    Dim sLtspi  As String
''*** UPDATE �� Y.SIMIZU 2005/10/1 GDײݐ��i�[�p
    Dim GDhin   As tFullHinban
''*** UPDATE �� Y.SIMIZU 2005/10/1 GDײݐ��i�[�p
    
    GetReferHinban = FUNCTION_RETURN_FAILURE
    
    '' ������
    With tHinInf
        .hinban = vbNullString
        .mnorevno = 0
        .factory = vbNullString
        .opecond = vbNullString
    End With
    
    '' �����T���v���Ǘ������w���̎擾
    Select Case Kensa_Typ
        Case CHK_OI         '' Oi
            chkShiji = tblCrySmp.CRYINDOICS
        Case CHK_CS         '' Cs
            chkShiji = tblCrySmp.CRYINDCSCS
        Case CHK_RS         '' Rs
            chkShiji = tblCrySmp.CRYINDRSCS
        Case CHK_B1         '' BMD1
            chkShiji = tblCrySmp.CRYINDB1CS
        Case CHK_B2         '' BMD2
            chkShiji = tblCrySmp.CRYINDB2CS
        Case CHK_B3         '' BMD3
            chkShiji = tblCrySmp.CRYINDB3CS
        Case CHK_L1         '' OSF1
            chkShiji = tblCrySmp.CRYINDL1CS
        Case CHK_L2         '' OSF2
            chkShiji = tblCrySmp.CRYINDL2CS
        Case CHK_L3         '' OSF3
            chkShiji = tblCrySmp.CRYINDL3CS
        Case CHK_L4         '' OSF4
            chkShiji = tblCrySmp.CRYINDL4CS
        Case CHK_GD         '' GD
            chkShiji = tblCrySmp.CRYINDGDCS
        Case CHK_LT         '' LT
            chkShiji = tblCrySmp.CRYINDTCS
        Case CHK_EP         '' EPD
            chkShiji = tblCrySmp.CRYINDEPCS
        Case CHK_X          '' X��              '2009/08 SUMCO Akizuki
            chkShiji = tblCrySmp.CRYINDXCS      '2009/08 SUMCO Akizuki
        'Add Start 2010/12/17 SMPK Miyata
        Case CHK_C         '' C
            chkShiji = tblCrySmp.CRYINDCCS
        Case CHK_CJ        '' CJ
            chkShiji = tblCrySmp.CRYINDCJCS
        Case CHK_CJLT      '' CJLT
            chkShiji = tblCrySmp.CRYINDCJLTCS
        Case CHK_CJ2       '' CJ2
            chkShiji = tblCrySmp.CRYINDCJ2CS
        'Add End   2010/12/17 SMPK Miyata
        Case Else           '' ���̑�
            Exit Function
    End Select
    
    '���C�t�^�C�����ѓ��͂̎��͉��[�ɃT���v���ʒu���܂ރu���b�N�̒���
    '�ł�������LT����ʒu�����i�Ԃ��擾
    If Kensa_Typ = CHK_LT Then
        With tblCrySmp
            DBDRV_getLtHinbanInBlock .XTALCS, .INPOSCS, LtHin, sLtspi
            If LtHin.hinban <> "        " Then
                tHinInf = LtHin
                .HINBCS = LtHin.hinban
                .REVNUMCS = LtHin.mnorevno
                .FACTORYCS = LtHin.factory
                .OPECS = LtHin.opecond
                GetReferHinban = FUNCTION_RETURN_SUCCESS
            End If
        End With
        Exit Function
    End If
    
''*** UPDATE �� Y.SIMIZU 2005/10/1 �ł�������GDײݐ��i�Ԃ����i�Ԃ��擾
    If Kensa_Typ = CHK_GD Then
        With tblCrySmp
            DBDRV_getGDHinbanInBlock tblCrySmp, GDhin
            If GDhin.hinban <> "        " Then
                tHinInf = GDhin
                .HINBCS = GDhin.hinban
                .REVNUMCS = GDhin.mnorevno
                .FACTORYCS = GDhin.factory
                .OPECS = GDhin.opecond
                GetReferHinban = FUNCTION_RETURN_SUCCESS
            End If
        End With
        Exit Function
    End If
''*** UPDATE �� Y.SIMIZU 2005/10/1 �ł�������GDײݐ��i�Ԃ����i�Ԃ��擾

''���������w���ύX�@2003/09/10 Motegi ========================> START
    '' ��i�ԁA���i�Ԃ����߂�
'    GetUpDownHinban UpHin, downHin, tblCrySmp, tbl_HinbanManage

    '' ���������w���𒲂ׂ�
'    If (chkShiji = "1") Or (chkShiji = "4" And tblCrySmp.SMPKBNCS = "T") Then       '' �����������̏ꍇ
'        If ChkSpExistence(downHin, Kensa_Typ) Then  '' ���i�ԂɎd�l������ꍇ
'            tHinInf = downHin
'        Else
'            Exit Function
'        End If
'    ElseIf (chkShiji = "2") Or (chkShiji = "4" And tblCrySmp.SMPKBNCS = "B") Then   '' �����������̏ꍇ
'        If ChkSpExistence(UpHin, Kensa_Typ) Then  '' ��i�ԂɎd�l������ꍇ
'            tHinInf = UpHin
'        Else
'            Exit Function
'        End If
'    ElseIf (chkShiji = "3") Then    '' ���ʌ����̏ꍇ
'
'        ' ��R�AOI�@�Ɋւ��Ă͑���_���������������g�p����
'        If (Kensa_Typ = CHK_RS) Or (Kensa_Typ = CHK_OI) Then
'            tHinInf = GetManyMark(Kensa_Typ, UpHin, downHin, tblCrySmp.SMPKBNCS)
'        Else
'            If (tblCrySmp.SMPKBNCS = "T") Then      '' �T���v���敪"T"�̏ꍇ
'                If ChkSpExistence(downHin, Kensa_Typ) Then  '' ���i�ԂɎd�l������ꍇ
'                    tHinInf = downHin
'                ElseIf ChkSpExistence(UpHin, Kensa_Typ) Then  '' ��i�ԂɎd�l������ꍇ
'                    tHinInf = UpHin
'                Else
'                    Exit Function
'                End If
'            ElseIf (tblCrySmp.SMPKBNCS = "B") Then  '' �T���v���敪"B"�̏ꍇ
'                If ChkSpExistence(UpHin, Kensa_Typ) Then  '' ��i�ԂɎd�l������ꍇ
'                    tHinInf = UpHin
'                ElseIf ChkSpExistence(downHin, Kensa_Typ) Then  '' ���i�ԂɎd�l������ꍇ
'                    tHinInf = downHin
'                Else
'                    Exit Function
'                End If
'            Else
'                Exit Function
'            End If
'        End If
'    Else    '' ���̑��̌����w��
'        Exit Function
'    End If
'--------------------------------

    With tHinInf
        .hinban = tblCrySmp.HINBCS
        .mnorevno = tblCrySmp.REVNUMCS
        .factory = tblCrySmp.FACTORYCS
        .opecond = tblCrySmp.OPECS
    End With

''���������w���ύX�@2003/09/10 Motegi ========================> END

    GetReferHinban = FUNCTION_RETURN_SUCCESS

End Function

'�����T���v���Ǘ�TBL���猟���w���̒l�𓾂�
Private Function GetSijiFlg(iSmpNo%, strFldName$) As String
Dim sql$
Dim rs As OraDynaset

    sql = "select " & strFldName & " from XSDCS where REPSMPLIDCS=" & iSmpNo
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount > 0 Then
        GetSijiFlg = rs(strFldName)
    Else
        GetSijiFlg = vbNullString
    End If
    rs.Close
    Set rs = Nothing
End Function
Public Function Add_CryRsRslt(tblTarget() As typ_TBCMJ002, tblDat As typ_TBCMJ002, Optional Index As Long = -1) As Integer
    Dim tblIndex As Long
    
    Add_CryRsRslt = FUNCTION_RETURN_FAILURE

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

    Add_CryRsRslt = FUNCTION_RETURN_SUCCESS
End Function
'�T�v      :���̓p�����[�^�̍��v�l���擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :dParam()      ,I   ,Double    ,�p�����[�^�l�z��
'          :�߂�l        ,O  ,Double    ,���v�l
'����      :
Public Function GetSum(dParam() As Double) As Double
    Dim Index   As Integer
    Dim dWork   As Double

    On Error GoTo Err

    dWork = 0
    For Index = 0 To UBound(dParam)
        dWork = dWork + dParam(Index)
    Next Index

    GetSum = dWork
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    GetSum = 0
End Function


'�T�v      :���̓p�����[�^�̕��ϒl���擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :dParam()      ,I   ,Double    ,�p�����[�^�l�z��
'          :�߂�l        ,O  ,Double    ,���ϒl
'����      :
Public Function GetAve(dParam() As Double) As Double
    Dim dWork   As Double

    On Error GoTo Err

    GetAve = GetSum(dParam) / (UBound(dParam) + 1)

    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    GetAve = 0
End Function


'�T�v      :���̓p�����[�^�̍ő�l���擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :dParam()      ,I   ,Double    ,�p�����[�^�l�z��
'          :�߂�l        ,O  ,Double    ,�ő�l
'����      :
Public Function GetMax(dParam() As Double) As Double
    Dim Index   As Integer
    Dim dMax    As Double

    On Error GoTo Err

    If UBound(dParam) = 0 Then GetMax = dParam(0): Exit Function
    
    dMax = dParam(0)
    For Index = 1 To UBound(dParam)
        If dMax < dParam(Index) Then dMax = dParam(Index)
    Next Index

    GetMax = dMax

    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    GetMax = 0
End Function

'�T�v      :���̓p�����[�^�̍ŏ��l���擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :dParam()      ,I   ,Double    ,�p�����[�^�l�z��
'          :�߂�l        ,O  ,Double    ,�ŏ��l
'����      :
Public Function GetMin(dParam() As Double) As Double
    Dim Index   As Integer
    Dim dMin    As Double

    On Error GoTo Err

    If UBound(dParam) = 0 Then GetMin = dParam(0): Exit Function
    
    dMin = dParam(0)
    For Index = 1 To UBound(dParam)
        If dMin > dParam(Index) Then dMin = dParam(Index)
    Next Index

    GetMin = dMin

    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    GetMin = 0
End Function



'�T�v      :�������W�c�̕W�{�ł���ƌ��Ȃ��āA��W�c�ɑ΂���W���΍����擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :dParam()      ,I   ,Double   ,�p�����[�^�l�z��
'          :�߂�l        ,O  ,Double    ,�W���΍��l
'����      :
Public Function GetSTDEV(dParam() As Double) As Double
    Dim Index   As Integer
    Dim dNum    As Double
    Dim dCalc1  As Double
    Dim dCalc2  As Double

    On Error GoTo Err

    dNum = UBound(dParam) + 1

    dCalc1 = 0
    For Index = 0 To dNum - 1
        dCalc1 = dCalc1 + dParam(Index) ^ 2
    Next Index

    dCalc2 = GetSum(dParam) ^ 2

    GetSTDEV = ((dNum * dCalc1 - dCalc2) / (dNum * (dNum - 1))) ^ 0.5

    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    GetSTDEV = 0
End Function


'�T�v      :�ŏ��Q��@�ɂ��A�����̌X�����v�Z����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :dY()          ,I  ,Double    ,���ɂ킩���Ă��� y �̒l�̌n��̕ϐ��iy = mx + b�j
'          :dX()          ,I  ,Double    ,���ɂ킩���Ă��� x �̒l�̌n��̕ϐ��iy = mx + b�j
'          :�߂�l        ,O  ,Double    ,�X��
'����      :
Public Function CalculateSlope(dY() As Double, dX() As Double) As Double
    Dim Index   As Integer
    Dim dNum    As Double
    Dim dCalc1  As Double
    Dim dCalc2  As Double
    Dim dParam  As Double

    On Error GoTo Err

    dNum = UBound(dY) + 1
    
    '' ��X(i)Y(i) �v�Z
    dCalc1 = 0
    For Index = 0 To dNum - 1
        dCalc1 = dCalc1 + dX(Index) * dY(Index)
    Next Index
    '' ��X(i)^2 �v�Z
    dCalc2 = 0
    For Index = 0 To dNum - 1
        dCalc2 = dCalc2 + dX(Index) ^ 2
    Next Index

    dParam = ((GetSum(dX) ^ 2) - dNum * dCalc2)
    If dParam = 0 Then CalculateSlope = 0: Exit Function

    CalculateSlope = (GetSum(dX) * GetSum(dY) - dNum * dCalc1) / dParam

    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    CalculateSlope = 0
End Function


'�T�v      :�ŏ��Q��@�ɂ��A������Y�ؕЂ��v�Z����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :dY()          ,I  ,Double    ,���ɂ킩���Ă��� y �̒l�̌n��̕ϐ��iy = mx + b�j
'          :dX()          ,I  ,Double    ,���ɂ킩���Ă��� x �̒l�̌n��̕ϐ��iy = mx + b�j
'          :�߂�l        ,O  ,Double    ,Y�ؕ�
'����      :
Public Function CalculateYFragment(dY() As Double, dX() As Double) As Double
    Dim Index   As Integer
    Dim dNum    As Double

    On Error GoTo Err

    dNum = UBound(dY) + 1

    CalculateYFragment = (GetSum(dY) - GetSum(dX) * CalculateSlope(dY, dX)) / dNum

    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    CalculateYFragment = 0
End Function

'�T�v      :���֌W�����v�Z����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :dY()          ,I  ,Double    ,���ɂ킩���Ă��� y �̒l�̌n��̕ϐ��iy = mx + b�j
'          :dX()          ,I  ,Double    ,���ɂ킩���Ă��� x �̒l�̌n��̕ϐ��iy = mx + b�j
'          :�߂�l        ,O  ,Double    ,Y�ؕ�
'����      :
Public Function CalculateR2(dY() As Double, dX() As Double) As Double
    Dim Index   As Integer
    Dim dNum    As Double
    Dim dCalc1  As Double
    Dim dCalc2  As Double
    Dim dCalc3  As Double
    Dim dParam  As Double

    On Error GoTo Err

    dNum = UBound(dY) + 1

    dCalc1 = 0
    For Index = 0 To dNum - 1
        dCalc1 = dCalc1 + ((dX(Index) - GetAve(dX)) * (dY(Index) - GetAve(dY)))
    Next Index

    dCalc2 = 0
    For Index = 0 To dNum - 1
        dCalc2 = dCalc2 + (dX(Index) - GetAve(dX)) ^ 2
    Next Index

    dCalc3 = 0
    For Index = 0 To dNum - 1
        dCalc3 = dCalc3 + (dY(Index) - GetAve(dY)) ^ 2
    Next Index

    dParam = (dCalc2 / (dNum - 1)) ^ 0.5 * (dCalc3 / (dNum - 1)) ^ 0.5
    If dParam = 0 Then CalculateR2 = 0: Exit Function
    
    CalculateR2 = ((dCalc1 / (dNum - 1)) / dParam) ^ 2

    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    CalculateR2 = 0
End Function

Public Sub RemoveAll_PlupEndRslt(tblTarget() As typ_TBCMH004)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_PrSpSXLData1(tblTarget() As typ_TBCME018)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_PrSpSXLData2(tblTarget() As typ_TBCME019)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_PrSpSXLData3(tblTarget() As typ_TBCME020)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub

'*** UPDATE �� Y.SIMIZU 2005/10/1 TBCME036�\���̂��ް���������
Public Sub RemoveAll_PrSpSXLData4(tblTarget() As typ_TBCME036)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub
'*** UPDATE �� Y.SIMIZU 2005/10/1 TBCME036�\���̂��ް���������

Public Sub RemoveAll_SXLInsideSpecManager(tblTarget() As typ_TBCME036)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_GFADevInfo(tblTarget() As typ_TBCMB014)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_HinbanManage(tblTarget() As typ_TBCME041)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_BlockManage(tblTarget() As typ_TBCME040)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_CrystalSampleManage(tblTarget() As typ_XSDCS)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_EPDRslt(tblTarget() As typ_TBCMJ001)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_CryRsRslt(tblTarget() As typ_TBCMJ002)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_OiRslt(tblTarget() As typ_TBCMJ003)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_CsRslt(tblTarget() As typ_TBCMJ004)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_BMDrslt(tblTarget() As typ_TBCMJ008)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_OSFRslt(tblTarget() As typ_TBCMJ005)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_GDRslt(tblTarget() As typ_TBCMJ006)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub

Public Sub RemoveAll_LifeTime(tblTarget() As typ_TBCMJ007)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub

'2009/08 SUMCO Akizuki X��������э쐬�ɔ����ǉ�
Public Sub RemoveAll_XRslt(tblTarget() As typ_TBCMJ021)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub

'�ǉ� 2005/06/17 ffc)tanabe
Public Sub RemoveAll_CrystalSampleManage_Cw(tblTarget() As typ_XSDCW)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub

'Add Start 2010/12/17 SMPK Miyata
Public Sub RemoveAll_CuDecoRslt(tblTarget() As typ_TBCMJ023)
    '' �e�[�u���f�[�^�S�폜
    ReDim tblTarget(0)
End Sub
'Add End   2010/12/17 SMPK Miyata

Public Sub InitAllTable()

    '' ���ׂẴe�[�u���̃f�[�^������������
    RemoveAll_PrSpSXLData1 tbl_PrSpSXLData1
    RemoveAll_PrSpSXLData2 tbl_PrSpSXLData2
    RemoveAll_PrSpSXLData3 tbl_PrSpSXLData3
    RemoveAll_GFADevInfo tbl_GFADevInfo
    RemoveAll_HinbanManage tbl_HinbanManage
    RemoveAll_BlockManage tbl_BlockManage
    RemoveAll_CrystalSampleManage tbl_CrystalSampleManage
    RemoveAll_EPDRslt tbl_EPDRslt
    RemoveAll_CryRsRslt tbl_CryRsRslt
    RemoveAll_OiRslt tbl_OiRslt
    RemoveAll_CsRslt tbl_CsRslt
    RemoveAll_BMDrslt tbl_BMDRslt
    RemoveAll_OSFRslt tbl_OSFRslt
    RemoveAll_GDRslt tbl_GDRslt
    RemoveAll_LifeTime tbl_LifeTime
    RemoveAll_SXLInsideSpecManager tbl_SXLInsideSpecManager
    RemoveAll_PlupEndRslt tbl_PlupEndRslt
    RemoveAll_CrystalSampleManage_Cw tbl_CrystalSampleManage_Cw '�ǉ� 2005/06/17 ffc)tanabe
    RemoveAll_CuDecoRslt tbl_CuDecoRslt                         'Add 2010/12/17 SMPK Miyata
End Sub
'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCME041�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME041 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺
Public Function DBDRV_GetTBCME041(records() As typ_TBCME041, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, INGOTPOS, HINBAN, REVNUM, FACTORY, OPECOND, LENGTH, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME041"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME041 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .INGOTPOS = rs("INGOTPOS")       ' �������J�n�ʒu
            .hinban = rs("HINBAN")           ' �i��
            .REVNUM = rs("REVNUM")           ' ���i�ԍ������ԍ�
            .factory = rs("FACTORY")         ' �H��
            .opecond = rs("OPECOND")         ' ���Ə���
            .Length = rs("LENGTH")           ' ����
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME041 = FUNCTION_RETURN_SUCCESS
End Function
Public Function Add_PrSpSXLData2(tblTarget() As typ_TBCME019, tblDat As typ_TBCME019, Optional Index As Long = -1) As Integer
    Dim tblIndex As Long
    
    Add_PrSpSXLData2 = FUNCTION_RETURN_FAILURE

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

    Add_PrSpSXLData2 = FUNCTION_RETURN_SUCCESS
End Function
Public Function Add_PrSpSXLData1(tblTarget() As typ_TBCME018, tblDat As typ_TBCME018, Optional Index As Long = -1) As Integer
    Dim tblIndex As Long
    
    Add_PrSpSXLData1 = FUNCTION_RETURN_FAILURE

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

    Add_PrSpSXLData1 = FUNCTION_RETURN_SUCCESS
End Function

Public Function Add_PrSpSXLData3(tblTarget() As typ_TBCME020, tblDat As typ_TBCME020, Optional Index As Long = -1) As Integer
    Dim tblIndex As Long
    
    Add_PrSpSXLData3 = FUNCTION_RETURN_FAILURE

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

    Add_PrSpSXLData3 = FUNCTION_RETURN_SUCCESS
End Function

'*** UPDATE �� Y.SIMIZU 2005/10/1 �n���ꂽ�i�Ԏd�l�ް����\���̂ɒǉ�
Public Function Add_PrSpSXLData4(tblTarget() As typ_TBCME036, tblDat As typ_TBCME036, Optional Index As Long = -1) As Integer
    Dim tblIndex As Long
    
    Add_PrSpSXLData4 = FUNCTION_RETURN_FAILURE

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

    Add_PrSpSXLData4 = FUNCTION_RETURN_SUCCESS
End Function
'*** UPDATE �� Y.SIMIZU 2005/10/1 �n���ꂽ�i�Ԏd�l�ް����\���̂ɒǉ�

'' �T���v���Ǘ��̌������ʒu���猩����i�ԁA���i�Ԃ��擾����
Public Function GetUpDownHinban(tUpHin As tFullHinban, _
                                tDownHin As tFullHinban, _
                                tblCrySmp As typ_XSDCS, _
                                tblHinban() As typ_TBCME041) As FUNCTION_RETURN
    Dim Index       As Integer
    Dim iIngPos     As Integer
    Dim iPos2       As Integer
    Dim tblFHin()   As typ_TBCME041
    Dim iHin        As Integer

    GetUpDownHinban = FUNCTION_RETURN_SUCCESS
    
    ClearFullHinban tUpHin
    ClearFullHinban tDownHin
    
    iIngPos = tblCrySmp.INPOSCS
    iHin = 0
    ReDim tblFHin(iHin)
    With tblFHin(iHin)
        .hinban = vbNullString            ' �i��
        .factory = vbNullString           ' �H��
        .opecond = vbNullString           ' ���Ə���
    End With
    
    '�w��ʒu�ɐڂ���i�Ԃ����X�g�A�b�v����i�u�w��ʒu���܂ށv�ł͂Ȃ��j
    For Index = 0 To UBound(tblHinban) - 1
        iPos2 = tblHinban(Index).INGOTPOS + tblHinban(Index).Length
        If (iIngPos >= tblHinban(Index).INGOTPOS) And (iIngPos <= iPos2) Then
            ReDim Preserve tblFHin(iHin)
            tblFHin(iHin) = tblHinban(Index)
            iHin = iHin + 1
        End If
    Next Index

    If UBound(tblFHin) = 0 Then
        SetFullHinban_TBCME041 tUpHin, tblFHin(0)
        SetFullHinban_TBCME041 tDownHin, tblFHin(0)
        Exit Function
    Else
        For Index = 0 To UBound(tblFHin)
            If iIngPos = tblFHin(Index).INGOTPOS Then
                SetFullHinban_TBCME041 tDownHin, tblFHin(Index)
            Else
                SetFullHinban_TBCME041 tUpHin, tblFHin(Index)
            End If
        Next
    End If

    GetUpDownHinban = FUNCTION_RETURN_FAILURE
End Function

'�T�v      :���グ�I�����т��擾����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :tblTarget     ,I   ,typ_TBCMH004 ,���グ�I�����уe�[�u��
'          :strCryNum     ,I   ,String       ,�����ԍ�
'          :�߂�l        ,O   ,Integer       ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Public Function GetPlupEndRslt(tblTarget As typ_TBCMH004, strCryNum As String) As Integer
    Dim iRet         As Integer
    Dim tblGet()    As typ_TBCMH004
    
    GetPlupEndRslt = FUNCTION_RETURN_FAILURE

    '' ���グ�I�����т̎擾
    iRet = DBDRV_GetTBCMH004(tblGet, "where CRYNUM='" & strCryNum & "'")
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function

    tblTarget = tblGet(1)


    GetPlupEndRslt = FUNCTION_RETURN_SUCCESS
End Function
Public Function Add_HinbanManage(tblTarget() As typ_TBCME041, tblDat As typ_TBCME041, Optional Index As Long = -1) As Integer
    Dim tblIndex As Long
    
    Add_HinbanManage = FUNCTION_RETURN_FAILURE

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

    Add_HinbanManage = FUNCTION_RETURN_SUCCESS
End Function
Public Function Add_CrystalSampleManage(tblTarget() As typ_XSDCS, tblDat As typ_XSDCS, Optional Index As Long = -1) As Integer
    Dim tblIndex As Long
    
    Add_CrystalSampleManage = FUNCTION_RETURN_FAILURE

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

    Add_CrystalSampleManage = FUNCTION_RETURN_SUCCESS
End Function

''Upd Start (TCS)T.Terauchi 2005/10/07  GDײݐ��\���Ή�
'�T�v      :�i�Ԃ�萻�i�d�l�v�e�f�[�^���擾�A�����āA�擾�������i�d�l�f�[�^��ǉ�����B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :tblSP()       ,O   ,typ_TBCME036   ,���i�d�l�v�e�f�[�^�e�[�u���z��
'          :tHinInf       ,I   ,tFullHinban    ,12���i��
'          :ctrlFrm       ,I   ,Form           ,�t�H�[��ID
'          :�߂�l        ,O  ,Integer          ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
'����      :05/10/07    (TCS)T.Terauchi
Public Function GetSPWFData36(tblSP() As typ_TBCME036, tHinInf As tFullHinban, ctrlFrm As Form) As Integer

    Dim Index       As Integer
    Dim iRet        As Integer
    Dim tHinban(0)  As tFullHinban
    Dim tblGet()    As typ_TBCME036
    
    GetSPWFData36 = FUNCTION_RETURN_FAILURE

    tHinban(0) = tHinInf

    '' ���i�d�l�v�e�f�[�^�̎擾
    iRet = DBDRV_GetTBCME036(tblGet, ctrlFrm.Name, tHinban)
    If iRet <> FUNCTION_RETURN_SUCCESS Then Exit Function
    If UBound(tblGet) = 0 Then Exit Function

    '' �擾�������i�d�l�v�e�f�[�^�̒ǉ�
    '' �e�[�u���f�[�^�i�[�̈�g��
    ReDim Preserve tblSP(UBound(tblSP) + 1)
    '' �e�[�u���f�[�^�����擾
    Index = UBound(tblSP)

    '' �f�[�^�ǉ�
    tblSP(Index) = tblGet(1)
    
    GetSPWFData36 = FUNCTION_RETURN_SUCCESS
    
End Function


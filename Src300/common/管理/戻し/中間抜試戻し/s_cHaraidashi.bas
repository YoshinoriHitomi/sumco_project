Attribute VB_Name = "s_cHaraidashi"
Option Explicit
'===========================================
' �v�e���H�p���ʃe�[�u��
'===========================================

' �����w��
Public Type typ_WafInd
    BLOCKID As String * 12      ' �u���b�NID
    BlockPos As Integer         ' �u���b�N�o
    SAMPLEID    As Variant      ' add 2003/03/28 hitec)matsumoto �T���v��ID���擾
    SAMPLEID2   As Variant      ' add 2003/03/28 hitec)matsumoto �T���v��ID2���擾
    INGOTPOS As Integer         ' �����o
    BkIngotPos  As Integer      ' add 2003/03/28 hitec)matsumoto
    Length As Integer           ' ����
    HINUP As tFullHinban        ' ��i��
    HINDN As tFullHinban        ' ���i��
    SMP As typ_WFSample         ' ��������
    HINFLG As Boolean           ' �i�ԋ�؂�t���O
    SMPFLG As Boolean           ' WF�T���v����؂�t���O
    ERRDNFLG As Boolean         ' ���i�ԃG���[�t���O
    SMPLKBN1 As String * 1      ' �T���v���敪�P
    SMPLKBN2 As String * 1      ' �T���v���敪�Q
    HANEIFLG As Boolean         '���f�t���O-------2003/09/23 �ǉ� iida
End Type

' ���i�d�l
Public Type typ_HinSpec
    hin As tFullHinban          ' �i��
    INGOTPOS As Integer         ' �������J�n�ʒu
    Length As Integer           ' ����
    HWFRMIN As Double           ' ���R����
    HWFRMAX As Double           ' ���R���
    HWFRHWYS As String * 1      ' �����L��(Rs)
    HWFONHWS As String * 1      ' �����L��(Oi)
    HWFBM1HS As String * 1      ' �����L��(B1)
    HWFBM2HS As String * 1      ' �����L��(B2)
    HWFBM3HS As String * 1      ' �����L��(B3)
    HWFOF1HS As String * 1      ' �����L��(L1)
    HWFOF2HS As String * 1      ' �����L��(L2)
    HWFOF3HS As String * 1      ' �����L��(L3)
    HWFOF4HS As String * 1      ' �����L��(L4)
    HWFDSOHS As String * 1      ' �����L��(DS)
    HWFMKHWS As String * 1      ' �����L��(DZ)
    HWFSPVHS As String * 1      ' �����L��(SP/Fe�Z�x)
    HWFDLHWS As String * 1      ' �����L��(SP/�g�U��)
    HWFNRHS  As String * 1      ' �����L��(SP/Nr�Z�x)  06/06/08 ooba
    HWFOS1HS As String * 1      ' �����L��(D1)
    HWFOS2HS As String * 1      ' �����L��(D2)
    HWFOS3HS As String * 1      ' �����L��(D3)
    HWFOTHER1 As String * 1     ' �����L��(OT1) '03/05/23
    HWFOTHER2 As String * 1     ' �����L��(OT2) '03/05/23
    HWFZOHWS As String * 1      ' �����L��(AO)�@'�ǉ� 03/12/09 ooba
    HWFDENHS As String * 1      ' �����L��(GD/DEN)  '�ǉ��@05/01/25 ooba START ====>
    HWFLDLHS As String * 1      ' �����L��(GD/LDL)
    HWFDVDHS As String * 1      ' �����L��(GD/DVD2) '�ǉ��@05/01/25 ooba END ======>
    HWFOTHER1MAI As String * 1  ' ����(OT1) '04/06/25
    HWFOTHER2MAI As String * 1  ' ����(OT2) '04/06/25
    WFCUTUNIT As String * 4     ' WF�J�b�g�P�� '�ǉ� 2005/04/12 ffc)tanabe
' 06/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    HEPOF1HS As String * 1      ' �����L��(OSF1E)
    HEPOF2HS As String * 1      ' �����L��(OSF2E)
    HEPOF3HS As String * 1      ' �����L��(OSF3E)
    HEPBM1HS As String * 1      ' �����L��(BMD1E)
    HEPBM2HS As String * 1      ' �����L��(BMD2E)
    HEPBM3HS As String * 1      ' �����L��(BMD3E)
' 06/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
' 10/02/16 Add SIRD�Ή� Y.Hitomi
    HWFSIRDHS As String * 1     ' �����L��(SIRD)
End Type

' �����E�F�n�[
Public Type typ_LackMap
    BLOCKID As String * 12      ' �u���b�NID
    LACKPOSS As Double          ' �����ʒu(From)
    LACKPOSE As Double          ' �����ʒu(To)
    REJCAT As String * 1        ' �������R
    LACKCNTS As Integer         ' ��������(From)
    LACKCNTE As Integer         ' ��������(To)
End Type

Public tblHinSpec() As typ_HinSpec      ' ���i�d�l�e�[�u��
Public tblWafInd() As typ_WafInd        ' �����w���e�[�u��
Public tblNukishi() As typ_XSDCW        '�����f�[�^�\���̍쐬�p�@2003/10/02 iida
Public iNowBlkPos As Integer            ' ���ݕ\���u���b�N�ʒu
Public iNowBlkCnt As Integer            ' ���ݕ\���u���b�N�T���v����
Public tblLackMap() As typ_LackMap      ' �����E�F�n�[�e�[�u��
Public bDispLock As Boolean             ' ��ʃ��b�N�t���O

'�T�v      :�t�^�c���㉺�ɕ���
'����      :�t�^�c�T���v�����㉺�ɕ�������
'����      :2001/10/05�@��� �쐬
Public Sub SeparateUD()
    'Step3.2�ɂāA�@�\�p�~
End Sub

''�T�v      :�����T���v���̎擾
''���Ұ��@�@:�ϐ���      ,IO ,�^       ,����
''�@�@      :sSamp �@�@�@,IO ,String �@,�T���v��
''�@�@      :iMode �@�@�@,I  ,Integer�@,1:�㑤�T���v��, 2:�����T���v��
''����      :�����T���v�����擾����
''����      :2001/10/05�@��� �쐬
'Public Sub GetSampleUD(sSamp As String, iMode As Integer)
'
'    Select Case sSamp
'    Case "1"
'        sSamp = IIf(iMode = 1, "0", "1")
'    Case "2"
'        sSamp = IIf(iMode = 1, "2", "0")
'    Case "3", "4"
'        sSamp = IIf(iMode = 1, "2", "1")
'    End Select
'
'End Sub
'---------------------�폜------------------------------------------

'�T�v      :�t���i�Ԃ̎擾
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'�@�@      :pHIN    �@�@�@,IO ,tFullHinban    �@,�i�ԃe�[�u��
'�@�@      :�߂�l        ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
'����      :�W���i�Ԃ���t���i�Ԃ��擾����
'����      :2001/07/11�@��� �쐬
Public Function GetFullHinban(pHin As tFullHinban) As FUNCTION_RETURN

    Dim sHin As String
    Dim m As Integer
    Dim i As Integer

    sHin = Trim(pHin.hinban)
    If sHin = "" Or sHin = "G" Or sHin = "Z" Then
        pHin.mnorevno = 0
        pHin.FACTORY = ""
        pHin.OPECOND = ""
        GetFullHinban = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If
    m = UBound(tblHinSpec)
    For i = 1 To m
        If tblHinSpec(i).hin.hinban = pHin.hinban Then
            pHin = tblHinSpec(i).hin
            GetFullHinban = FUNCTION_RETURN_SUCCESS
            Exit Function
        End If
    Next i
    GetFullHinban = GetLastHinban(pHin.hinban, pHin)

End Function

'�T�v      :�T���v���̃g�b�v���^�{�g�����敪�̎擾
'���Ұ��@�@:�ϐ����@�@�@�@,IO ,�^       ,����
'�@�@      :sSample      ,I  ,String �@,�T���v��
'�@�@      :bTop         ,O  ,Boolean�@,�g�b�v���敪�̗L��
'�@�@      :bBot         ,O  ,Boolean�@,�{�g�����敪�̗L��
'����      :�T���v���敪�̗L����Ԃ�
'����      :2001/07/11�@��� �쐬
Public Sub GetSampleBT(ByVal sSample As String, bTop As Boolean, bBot As Boolean)

    Select Case sSample
    Case "1"
        bTop = True
    Case "2"
        bBot = True
    Case "4"
        bTop = True
        bBot = True
    End Select

End Sub

'�T�v      :�T���v���̃X�L�b�v
'���Ұ��@�@:�ϐ���       ,IO ,�^      ,����
'�@�@      :sSample�@�@�@,I  ,String�@,�T���v��
'�@�@      :sSkip  �@�@�@,I  ,String�@,�X�L�b�v����T���v��
'�@�@      :�߂�l       ,O  ,String�@,�X�L�b�v��̃T���v��
'����      :�w�肳�ꂽ�T���v���Ȃ�O�N���A����
'����      :2001/07/03�@��� �쐬
Public Function SkipSample(ByVal sSample As String, ByVal sSkip As String) As String

    If sSample = sSkip Then
        SkipSample = "0"
    Else
        SkipSample = sSample
    End If

End Function

'�T�v      :SXL ID�̎擾
'���Ұ��@�@:�ϐ���         ,IO ,�^       ,����
'�@�@      :sBlockID �@�@�@,I  ,String �@,�u���b�NID
'�@�@      :iIngotPos�@�@�@,I  ,Integer�@,�������J�n�ʒu
'�@�@      :�߂�l         ,O  ,String �@,SXL ID
'����      :SXL ID��Ԃ�
'����      :2001/07/11�@��� �쐬
Public Function GetSXLID(sBlockId As String, iIngotpos As Integer) As String

    GetSXLID = left(sBlockId, 10) & GetWafPos(iIngotpos)

End Function

'�T�v      :�����ʒu������̎擾
'���Ұ��@�@:�ϐ���         ,IO ,�^       ,����
'�@�@      :iIngotPos�@�@�@,I  ,Integer�@,�������J�n�ʒu
'�@�@      :�߂�l         ,O  ,String �@,�����ʒu������
'����      :�����ʒu�������Ԃ�
'����      :2001/07/11�@��� �쐬
Public Function GetWafPos(iIngotpos As Integer) As String

    Dim i As Integer
    Dim j As Integer

    If iIngotpos >= 1000 Then
        i = Int(iIngotpos / 100)
        j = iIngotpos Mod 100
        GetWafPos = Chr$(i - 10 + Asc("A")) & Format(j, "00")
    Else
        GetWafPos = Format(iIngotpos, "000")
    End If

End Function

'�T�v      :����]�����@�w���e�[�u���̍쐬
'���Ұ��@�@:�ϐ���       ,IO ,�^                ,����
'�@�@      :pSXLMng�@�@�@,I  ,typ_TBCME042   �@ ,SXL�Ǘ�
'�@�@      :pWafSmp�@�@�@,I  ,typ_XSDCW   �@    ,�V�T���v���Ǘ��iSXL�j
'�@�@      :pMesInd�@�@�@,O  ,typ_TBCMY003   �@ ,����]�����@�w��
'�@�@      :�߂�l       ,O  ,FUNCTION_RETURN�@ ,�ǂݍ��݂̐���
'����      :����]�����@�w���e�[�u�����쐬����
'����      :2001/07/23�@��� �쐬
'           2006/08/15�@SMP)kondoh �C��
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
'Public Function MakeMesIndTbl(pSXLMng() As typ_TBCME042, pWafSmp() As typ_XSDCW, pMesInd() As typ_TBCMY003) As FUNCTION_RETURN
Public Function MakeMesIndTbl(pSXLMng() As typ_TBCME042, pWafSmp() As typ_XSDCW, _
                        pMesInd() As typ_TBCMY003, pEpMesInd() As typ_TBCMY020) As FUNCTION_RETURN
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

    Dim tmpSpWFSamp() As typ_SpWFSamp
    Dim sHin As String
    Dim sDKAN As String
    Dim m As Integer
    Dim n As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim sGdSpec As String       '�K�i�l(GD)�@05/01/27 ooba
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    '' �G�s��s�]�����ڗp��DK�A�j�[������
    Dim sDKAN_EP        As String
    Dim l               As Integer
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

    '' ����]�����@�w���p�̐��i�d�l���擾
    j = 0
    m = UBound(pSXLMng)
    ReDim tmpSpWFSamp(m)
    For i = 1 To m
        sHin = RTrim$(pSXLMng(i).hinban)
        If sHin <> "" And sHin <> "G" And sHin <> "Z" Then
            j = j + 1
            tmpSpWFSamp(j).hin.hinban = pSXLMng(i).hinban
            tmpSpWFSamp(j).hin.mnorevno = pSXLMng(i).REVNUM
            tmpSpWFSamp(j).hin.FACTORY = pSXLMng(i).FACTORY
            tmpSpWFSamp(j).hin.OPECOND = pSXLMng(i).OPECOND
            If scmzc_getWF(tmpSpWFSamp(j)) = FUNCTION_RETURN_FAILURE Then
                MakeMesIndTbl = FUNCTION_RETURN_FAILURE
                Exit Function
            End If
        End If
    Next i
    ReDim Preserve tmpSpWFSamp(j)

    '' ����]�����@�w���e�[�u���̍쐬
    k = 0
    m = UBound(pWafSmp)
    n = UBound(tmpSpWFSamp)
    ReDim pMesInd(m * 18)   'OTH2���폜 �G�s��s�]���ǉ��Ή� 06/08/15 SMP)kondoh
'    ReDim pMesInd(m * 19)   'GD�ǉ��@05/01/18 ooba
'    ReDim pMesInd(m * 18)   '�c���_�f�ǉ��@03/12/05 ooba
'    ReDim pMesInd(m * 17)   '03/05/24
'    ReDim pMesInd(m * 15)

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    l = 0
    ReDim pEpMesInd(m * 7)  ' OTH2�AOSF1E�`OSF3E�ABMD1E�`BMD3E
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

    For i = 1 To m
        For j = 1 To n
            If tmpSpWFSamp(j).hin.hinban = pWafSmp(i).HINBCW Then
                Exit For
            End If
        Next j
        If j <= n Then
            With tmpSpWFSamp(j)
'                sDKAN = IIf(.HWFIGKBN = "3", "R ", "V ") & Format(.HWFANTNP, "@@@@") & Format(.HWFANTIM, "@@@@")
                'DK�ưُ����ύX(IG�敪"4"��"R",�K�X��ǉ�)�@04/07/29 ooba
                sDKAN = IIf(.HWFIGKBN = "3" Or .HWFIGKBN = "4", "R ", "V ") & Format(.HWFANTNP, "@@@@") & " " & .HWFANGZY
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                ' �G�s��s�]�����ڗp��DK�A�j�[������
                ' (1���ځF�iEPIG�敪,3�`6���ځF�iEPAN���x,8���ځF�iEP����AN�K�X����,10���ځF�iE1�����S�̐�����1�̈�)
                sDKAN_EP = IIf(.HEPIGKBN = "3" Or .HEPIGKBN = "4", "R", "V") & " " & _
                            IIf(.HEPANTNP >= 0, Format(.HEPANTNP, "@@@@"), Space(4)) & " " & _
                            .HEPANGZY & " " & _
                            IIf(.HEPACEN >= 0, Mid(Format(.HEPACEN, "000.00"), 3, 1), Space(1))
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
                '���@�ȉ��S�Ă̎w���̗L���� <>"0" �Ŕ��肵�Ă������A�����̂݁i="1"�j�Ŕ��肷��悤�ɕύX
                If pWafSmp(i).WFINDRSCW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "RES"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "RES"
                    pMesInd(k).NETSU = ""
                    pMesInd(k).ET = ""
                    pMesInd(k).MES = .HWFRSPOH & .HWFRSPOT & .HWFRSPOI
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDOICW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "OI"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "OI"
                    pMesInd(k).NETSU = ""
                    pMesInd(k).ET = ""
                    pMesInd(k).MES = .HWFONSPH & .HWFONSPT & .HWFONSPI
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDB1CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "BMD"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "BMD1"
                    pMesInd(k).NETSU = .HWFBM1NS
                    pMesInd(k).ET = .HWFBM1SZ & Format(.HWFBM1ET, "00")
                    pMesInd(k).MES = .HWFBM1SH & .HWFBM1ST & .HWFBM1SR
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDB2CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "BMD"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "BMD2"
                    pMesInd(k).NETSU = .HWFBM2NS
                    pMesInd(k).ET = .HWFBM2SZ & Format(.HWFBM2ET, "00")
                    pMesInd(k).MES = .HWFBM2SH & .HWFBM2ST & .HWFBM2SR
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDB3CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "BMD"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "BMD3"
                    pMesInd(k).NETSU = .HWFBM3NS
                    pMesInd(k).ET = .HWFBM3SZ & Format(.HWFBM3ET, "00")
                    pMesInd(k).MES = .HWFBM3SH & .HWFBM3ST & .HWFBM3SR
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDL1CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "OSF"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "OSF1"
                    pMesInd(k).NETSU = .HWFOF1NS
                    pMesInd(k).ET = .HWFOF1SZ & Format(.HWFOF1ET, "00")
                    pMesInd(k).MES = .HWFOF1SH & .HWFOF1ST & .HWFOF1SR
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDL2CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "OSF"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "OSF2"
                    pMesInd(k).NETSU = .HWFOF2NS
                    pMesInd(k).ET = .HWFOF2SZ & Format(.HWFOF2ET, "00")
                    pMesInd(k).MES = .HWFOF2SH & .HWFOF2ST & .HWFOF2SR
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDL3CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "OSF"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "OSF3"
                    pMesInd(k).NETSU = .HWFOF3NS
                    pMesInd(k).ET = .HWFOF3SZ & Format(.HWFOF3ET, "00")
                    pMesInd(k).MES = .HWFOF3SH & .HWFOF3ST & .HWFOF3SR
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP START(OSF4->SIRD)
'''                If pWafSmp(i).WFINDL4CW = "1" Then
'''                    k = k + 1
'''                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
'''                    pMesInd(k).OSITEM = "OSF"
'''                    pMesInd(k).SAMPLEKB = "A"
'''                    pMesInd(k).Spec = "OSF4"
'''                    pMesInd(k).NETSU = .HWFOF4NS
'''                    pMesInd(k).ET = .HWFOF4SZ & Format(.HWFOF4ET, "00")
'''                    pMesInd(k).MES = .HWFOF4SH & .HWFOF4ST & .HWFOF4SR
'''                    pMesInd(k).DKAN = sDKAN
'''                    pMesInd(k).MAISU = "1"
'''                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
'''                End If

                If pWafSmp(i).WFINDL4CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
'                    pMesInd(k).OSITEM = "SIRD"
                    pMesInd(k).OSITEM = "TENI"  '2010/05/19 REP Y.HItomi
                    pMesInd(k).SAMPLEKB = "A"
'                    pMesInd(k).Spec = "SIRD"
                    pMesInd(k).Spec = "TENI"    '2010/05/19 REP Y.HItomi
                    pMesInd(k).NETSU = ""
                    pMesInd(k).ET = .HWFSIRDSZ '''& Format(.HWFOF4ET, "00")
                    pMesInd(k).MES = ""
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP  END (OSF4->SIRD)
                If pWafSmp(i).WFINDDSCW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "DSOD"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "DSOD"
                    pMesInd(k).NETSU = "G0"
                    pMesInd(k).ET = ""
                    pMesInd(k).MES = ""
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDDZCW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "DZ"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "DZ"
                    pMesInd(k).NETSU = .HWFMKNSW
                    pMesInd(k).ET = .HWFMKSZY & Format(.HWFMKCET, "00")
                    pMesInd(k).MES = .HWFMKSPH & .HWFMKSPT & .HWFMKSPR
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDSPCW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "SPV"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "SPV"
                    pMesInd(k).NETSU = ""
                    pMesInd(k).ET = ""
'                    pMesInd(k).MES = .HWFSPVSH & .HWFSPVST & .HWFSPVSI
                    '05/10/13 ooba START ==============================================>
                    If .HWFSPVHS = "H" Or .HWFSPVHS = "S" Then
                        pMesInd(k).MES = .HWFSPVSH & .HWFSPVST & .HWFSPVSI
                    ElseIf .HWFDLHWS = "H" Or .HWFDLHWS = "S" Then
                        pMesInd(k).MES = .HWFDLSPH & .HWFDLSPT & .HWFDLSPI
                    Else    'Nr�Z�x�ǉ��@06/06/08 ooba
                        pMesInd(k).MES = .HWFNRSH & .HWFNRST & .HWFNRSI
                    End If
                    '05/10/13 ooba END ================================================>
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    '06/06/08 ooba START ==============================================>
                    pMesInd(k).FEPUA = .HWFSPVPUG           'SPV_Fe_PUA�l
                    pMesInd(k).FEPUAPCT = .HWFSPVPUR        'SPV_Fe_PUA���l
                    pMesInd(k).FESTD = .HWFSPVSTD           'SPV_Fe_STD
                    pMesInd(k).DIFFPUA = .HWFDLPUG          'SPV_�g�U��_PUA�l
                    pMesInd(k).DIFFPUAPCT = .HWFDLPUR       'SPV_�g�U��_PUA���l
                    pMesInd(k).NRPUA = .HWFNRPUG            'SPV_NR_PUA�l
                    pMesInd(k).NRPUAPCT = .HWFNRPUR         'SPV_NR_PUA%�l
                    pMesInd(k).NRSTD = .HWFNRSTD            'SPV_NR_STD
                    '06/06/08 ooba END ================================================>
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDDO1CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "DOI"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "DOI1"
                    pMesInd(k).NETSU = .HWFOS1NS
                    pMesInd(k).ET = ""
                    pMesInd(k).MES = .HWFOS1SH & .HWFOS1ST & .HWFOS1SI
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDDO2CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "DOI"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "DOI2"
                    pMesInd(k).NETSU = .HWFOS2NS
                    pMesInd(k).ET = ""
                    pMesInd(k).MES = .HWFOS2SH & .HWFOS2ST & .HWFOS2SI
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDDO3CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "DOI"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "DOI3"
                    pMesInd(k).NETSU = .HWFOS3NS
                    pMesInd(k).ET = ""
                    pMesInd(k).MES = .HWFOS3SH & .HWFOS3ST & .HWFOS3SI
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                '################################## Add,03/05/23 hitec)matsumoto ##########
                If pWafSmp(i).WFINDOT1CW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
''''                pMesInd(k).OSITEM = "OTH"
                    pMesInd(k).OSITEM = "OTH1"  'upd 2003/06/09 hitec)matsumoto �d�l�ύX
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "OTHER1"
'''                 pMesInd(k).NETSU = .HWFOS3NS
'''                 pMesInd(k).ET = ""
'''                 pMesInd(k).MES = .HWFOS3SH & .HWFOS3ST & .HWFOS3SI
                    pMesInd(k).NETSU = vbNullString
                    pMesInd(k).ET = vbNullString
                    pMesInd(k).MES = vbNullString
''''                pMesInd(k).DKAN = vbNullString  '03/05/22
                    pMesInd(k).DKAN = sDKAN 'upd 2003/06/10 hitec)matsumoto
                    pMesInd(k).MAISU = .HWOTHER1MAI
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                If pWafSmp(i).WFINDOT2CW = "1" Then
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
''                    k = k + 1
''                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
''''''                pMesInd(k).OSITEM = "OTH"
''                    pMesInd(k).OSITEM = "OTH2"  'upd 2003/06/09 hitec)matsumoto �d�l�ύX
''                    pMesInd(k).SAMPLEKB = "A"
''                    pMesInd(k).Spec = "OTHER2"
'''''                 pMesInd(k).NETSU = .HWFOS3NS
'''''                 pMesInd(k).ET = ""
'''''                 pMesInd(k).MES = .HWFOS3SH & .HWFOS3ST & .HWFOS3SI
''                    pMesInd(k).NETSU = vbNullString
''                    pMesInd(k).ET = vbNullString
''                    pMesInd(k).MES = vbNullString
''''''                pMesInd(k).DKAN = vbNullString  '03/05/22
''                    pMesInd(k).DKAN = sDKAN 'upd 2003/06/10 hitec)matsumoto
''                    pMesInd(k).MAISU = .HWOTHER2MAI
                    l = l + 1
                    pEpMesInd(l).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pEpMesInd(l).OSITEM = "OTH2"
                    pEpMesInd(l).SAMPLEKB = "A"
                    pEpMesInd(l).Spec = "OTHER2"
                    pEpMesInd(l).NETSU = vbNullString
                    pEpMesInd(l).ET = vbNullString
                    pEpMesInd(l).MES = vbNullString
                    pEpMesInd(l).DKAN = sDKAN_EP
                    pEpMesInd(l).MAISU = .HWOTHER2MAI
                    pEpMesInd(l).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
                End If
                '################################## End,03/05/23 hitec)matsumoto ##########
                
                '' �c���_�f�ǉ��@03/12/05 ooba START ===============================>
                If pWafSmp(i).WFINDAOICW = "1" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pMesInd(k).OSITEM = "AOI"
                    pMesInd(k).SAMPLEKB = "A"
                    pMesInd(k).Spec = "AOI"
                    pMesInd(k).NETSU = .HWFZONSW
                    pMesInd(k).ET = ""
                    pMesInd(k).MES = .HWFZOSPH & .HWFZOSPT & .HWFZOSPI
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                '' �c���_�f�ǉ��@03/12/05 ooba END =================================>
                
                '' GD�ǉ��@05/01/18 ooba START =====================================>
                If pWafSmp(i).WFINDGDCW = "1" And pWafSmp(i).WFHSGDCW = "0" Then
                    k = k + 1
                    pMesInd(k).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    
                ''Upd Start (TCS)T.Terauchi 2005/10/05  �����w��4.5ײݑΉ�
                ''    pMesInd(k).OSITEM = "GD"
                    
                    If Trim(.HWFGDLINE) = "3" Then
                        pMesInd(k).OSITEM = "GD"
                    ElseIf Trim(.HWFGDLINE) = "4.5" Then
                        pMesInd(k).OSITEM = "GD45"
                    ElseIf Trim(.HWFGDLINE) = "5" Then
                        pMesInd(k).OSITEM = "GD50"
                    Else
                        pMesInd(k).OSITEM = "GD"
                    End If
                ''Upd End   (TCS)T.Terauchi 2005/10/05  �����w��4.5ײݑΉ�
                    
                    pMesInd(k).SAMPLEKB = "A"
                    
                    '�K�i�l(SPEC) 1����:DVD2
                    If .HWFDVDHS = "H" Or .HWFDVDHS = "S" Then sGdSpec = "V" Else sGdSpec = Space(1)
                    sGdSpec = sGdSpec & Space(1)
                    '�K�i�l(SPEC) 3����:L/DL
                    If .HWFLDLHS = "H" Or .HWFLDLHS = "S" Then sGdSpec = sGdSpec & "L" Else sGdSpec = sGdSpec & Space(1)
                    sGdSpec = sGdSpec & Space(1)
                    '�K�i�l(SPEC) 5����:Den
                    If .HWFDENHS = "H" Or .HWFDENHS = "S" Then sGdSpec = sGdSpec & "D" Else sGdSpec = sGdSpec & Space(1)
                    
                    pMesInd(k).Spec = sGdSpec
                    pMesInd(k).NETSU = ""
                    pMesInd(k).ET = ""
'                    pMesInd(k).MES = ""
                    pMesInd(k).MES = .HWFGDSPH & .HWFGDSPT & .HWFGDZAR      '05/10/25 ooba
                    pMesInd(k).DKAN = sDKAN
                    pMesInd(k).MAISU = "1"
                    pMesInd(k).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                '' GD�ǉ��@05/01/18 ooba END =======================================>
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                '' OSF1E
                If pWafSmp(i).EPINDL1CW = "1" Then
                    l = l + 1
                    pEpMesInd(l).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pEpMesInd(l).OSITEM = "OSF"
                    pEpMesInd(l).SAMPLEKB = "A"
                    pEpMesInd(l).Spec = "OSF1"
                    pEpMesInd(l).NETSU = .HEPOF1NS
                    pEpMesInd(l).ET = .HEPOF1SZ & IIf(.HEPOF1ET >= 0, Format(.HEPOF1ET, "00"), Space(2))
                    pEpMesInd(l).MES = .HEPOF1SH & .HEPOF1ST & .HEPOF1SR
                    pEpMesInd(l).DKAN = sDKAN_EP
                    pEpMesInd(l).MAISU = "1"
                    pEpMesInd(l).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                '' OSF2E
                If pWafSmp(i).EPINDL2CW = "1" Then
                    l = l + 1
                    pEpMesInd(l).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pEpMesInd(l).OSITEM = "OSF"
                    pEpMesInd(l).SAMPLEKB = "A"
                    pEpMesInd(l).Spec = "OSF2"
                    pEpMesInd(l).NETSU = .HEPOF2NS
                    pEpMesInd(l).ET = .HEPOF2SZ & IIf(.HEPOF2ET >= 0, Format(.HEPOF2ET, "00"), Space(2))
                    pEpMesInd(l).MES = .HEPOF2SH & .HEPOF2ST & .HEPOF2SR
                    pEpMesInd(l).DKAN = sDKAN_EP
                    pEpMesInd(l).MAISU = "1"
                    pEpMesInd(l).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                '' OSF3E
                If pWafSmp(i).EPINDL3CW = "1" Then
                    l = l + 1
                    pEpMesInd(l).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pEpMesInd(l).OSITEM = "OSF"
                    pEpMesInd(l).SAMPLEKB = "A"
                    pEpMesInd(l).Spec = "OSF3"
                    pEpMesInd(l).NETSU = .HEPOF3NS
                    pEpMesInd(l).ET = .HEPOF3SZ & IIf(.HEPOF3ET >= 0, Format(.HEPOF3ET, "00"), Space(2))
                    pEpMesInd(l).MES = .HEPOF3SH & .HEPOF3ST & .HEPOF3SR
                    pEpMesInd(l).DKAN = sDKAN_EP
                    pEpMesInd(l).MAISU = "1"
                    pEpMesInd(l).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                '' BMD1E
                If pWafSmp(i).EPINDB1CW = "1" Then
                    l = l + 1
                    pEpMesInd(l).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pEpMesInd(l).OSITEM = "BMD"
                    pEpMesInd(l).SAMPLEKB = "A"
                    pEpMesInd(l).Spec = "BMD1"
                    pEpMesInd(l).NETSU = .HEPBM1NS
                    pEpMesInd(l).ET = .HEPBM1SZ & IIf(.HEPBM1ET >= 0, Format(.HEPBM1ET, "00"), Space(2))
                    pEpMesInd(l).MES = .HEPBM1SH & .HEPBM1ST & .HEPBM1SR
                    pEpMesInd(l).DKAN = sDKAN_EP
                    pEpMesInd(l).MAISU = "1"
                    pEpMesInd(l).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                '' BMD2E
                If pWafSmp(i).EPINDB2CW = "1" Then
                    l = l + 1
                    pEpMesInd(l).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pEpMesInd(l).OSITEM = "BMD"
                    pEpMesInd(l).SAMPLEKB = "A"
                    pEpMesInd(l).Spec = "BMD2"
                    pEpMesInd(l).NETSU = .HEPBM2NS
                    pEpMesInd(l).ET = .HEPBM2SZ & IIf(.HEPBM2ET >= 0, Format(.HEPBM2ET, "00"), Space(2))
                    pEpMesInd(l).MES = .HEPBM2SH & .HEPBM2ST & .HEPBM2SR
                    pEpMesInd(l).DKAN = sDKAN_EP
                    pEpMesInd(l).MAISU = "1"
                    pEpMesInd(l).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
                '' BMD3E
                If pWafSmp(i).EPINDB3CW = "1" Then
                    l = l + 1
                    pEpMesInd(l).SAMPLEID = pWafSmp(i).REPSMPLIDCW
                    pEpMesInd(l).OSITEM = "BMD"
                    pEpMesInd(l).SAMPLEKB = "A"
                    pEpMesInd(l).Spec = "BMD3"
                    pEpMesInd(l).NETSU = .HEPBM3NS
                    pEpMesInd(l).ET = .HEPBM3SZ & IIf(.HEPBM3ET >= 0, Format(.HEPBM3ET, "00"), Space(2))
                    pEpMesInd(l).MES = .HEPBM3SH & .HEPBM3ST & .HEPBM3SR
                    pEpMesInd(l).DKAN = sDKAN_EP
                    pEpMesInd(l).MAISU = "1"
                    pEpMesInd(l).MUKESAKI = sCmbMukesaki    ' 2007/09/13 SPK Tsutsumi Add
                End If
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
            End With
        End If
    Next i
    ReDim Preserve pMesInd(k)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    ReDim Preserve pEpMesInd(l)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

    MakeMesIndTbl = FUNCTION_RETURN_SUCCESS

End Function


Attribute VB_Name = "SB_Com"
Option Explicit

'WF����ٍ\����
Public Type typ_Wf_Smpl
    SXLIDCW     As String * 13      'SXL-ID
    TBKBNCW     As String * 1       'T/B�敪
    XTALCW      As String * 12      '�����ԍ�
    INPOSCW     As Integer          '�������ʒu
    HINBCW      As String * 8       '�i��
    REVNUMCW    As Integer          '���i�ԍ������ԍ�
    FACTORYCW   As String * 1       '�H��
    OPECW       As String * 1       '���Ə���
End Type

'��������ٍ\����
Public Type typ_Cry_Smpl
    CRYNUMCS    As String * 12      '��ۯ�ID
    SMPKBNCS    As String * 1       '����ً敪
    TBKBNCS     As String * 1       'T/B�敪
    REPSMPLIDCS As Long             '��\�����ID    Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    XTALCS      As String * 12      '�����ԍ�
    INPOSCS     As Integer          '�������ʒu
    HINBCS      As String * 8       '�i��
    REVNUMCS    As Integer          '���i�ԍ������ԍ�
    FACTORYCS   As String * 1       '�H��
    OPECS       As String * 1       '���Ə���
End Type

Public CrySampleID  As typ_CpyJisseki     ' �������ш��p���ް��@05/06/13 ooba

'�������ш��p���ް��@05/06/13 ooba
Public Type typ_CpyJisseki
    TsmplidGD   As String * 16          'TOP_�����ID(GD)
    TindGD      As String * 1           'TOP_���FLG(GD)
    BsmplidGD   As String * 16          'BOT_�����ID(GD)
    BindGD      As String * 1           'BOT_���FLG(GD)
End Type

'Warp/�����p�x����l�\���ް��@05/12/19 ooba
Public Type typ_WarpKakuData
    BLOCKID     As String * 12              '��ۯ�ID
    HIN         As tFullHinban              '�i��
    WAFID       As Double                   '��ʰID
    Min         As Double                   '�d�lMin�l
    max         As Double                   '�d�lMax�l
    MEASDATA    As Double                   '����l
    Judg        As Boolean                  '����(True:����OK,False:����NG)
    EXISTFLG    As Integer                  '�����׸�(1:���ް��L,0:���ް���,-1:WFϯ�ߕR�t����)
End Type

'WFϯ�ߏ�̕i���ް��@05/12/19 ooba
Public Type typ_MapHinData
    BLOCKID     As String * 12              '��ۯ�ID
    HIN         As tFullHinban              '�i��
    BLKSEQ_S    As Integer                  '��ۯ����A��(Start)
    BLKSEQ_E    As Integer                  '��ۯ����A��(End)
    WARPFLG     As Boolean                  'Warp�U�������׸�
    KAKUFLG     As Boolean                  '�����p�x�U�������׸�
    'Add Start 2011/04/25 SMPK Miyata
    XTALCS      As String * 12              '�����ԍ�
    INPOSCS_S   As Integer                  '�������ʒu(Start)
    INPOSCS_E   As Integer                  '�������ʒu(End)
    'Add End   2011/04/25 SMPK Miyata
End Type

'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga STRAT ---
Public Type typ_COSF3ID
    
    C_XTALC1    As String          ' �����ԍ�
    C_JDGEIDC1  As String          ' C-OSF3����ID
    C_SYNFLAGC5 As String          ' ���F�׸�
    C_YMKFLAGC5 As String          ' �폜�׸�
    C_strChkR   As String          ' ����p����݋敪
    C_strChkD   As String          ' ����p����݋敪
    C_POSC5     As String          ' ����وʒu
    C_DMAXC5    As String          ' D�̂ݏ��
    C_RMAXC5    As String          ' R�̂ݏ��
    C_DRDMAXC5  As String          ' D�������
    C_DRRMAXC5  As String          ' R�������
         
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga END ---
    
End Type

'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA�]���Ή�(Cu-deco) : Cu-deco(CJ, CJ2)�̕��ʕʋK�i�l�\����(�������xodc5_osf31)
Public Type typ_SB_com_xodb5_osf31_Cudeco
    
    JDGEIDC5            As String * 4       ' ����D(C-OSF3, Cu-Deco)
    POSC5               As Long             ' ����
    
    CJDMAXPIC5          As Integer          ' CJ Disk�̂݃p�^�[�� Pi�����
    CJRMAXPIC5          As Integer          ' CJ Ring�̂݃p�^�[�� Pi�����
    CJDRMAXPIC5         As Integer          ' CJ DiskRing�p�^�[�� Pi�����
    CJALLMAXDIC5        As Integer          ' CJ ����Disk���a���
    CJALLMINRINC5       As Integer          ' CJ ����Ring���a����
    CJALLMAXRIGC5       As Integer          ' CJ ����Ring�O�a���
    
    CJ2DMAXPIC5         As Integer          ' CJ2 Disk�̂݃p�^�[�� Pi������(MAX���������ł�)
    CJ2RMAXPIC5         As Integer          ' CJ2 Ring�̂݃p�^�[�� Pi������(MAX���������ł�)
    CJ2RMINRINC5        As Integer          ' CJ2 Ring�̂݃p�^�[�� Ring���a����
    CJ2RMAXRIGC5        As Integer          ' CJ2 Ring�̂݃p�^�[�� Ring�O�a���
    CJ2DRMAXPIC5        As Integer          ' CJ2 DiskRing�p�^�[�� Pi������(MAX���������ł�)
    CJ2DRMINRINC5       As Integer          ' CJ2 DiskRing�p�^�[�� Ring���a����
    CJ2DRMAXRIGC5       As Integer          ' CJ2 DiskRing�p�^�[�� Ring�O�a���

End Type
''Add End   2011/01/17 SMPK A.Nagamine

Public JudgKoutei           As String       '�H��(�������������p)�@08/04/15 ooba

'--------------- 2008/07/25 INSERT START  By Systech ---------------
Public gsTbcmy028ErrCode    As String           ' �U�փ`�F�b�N�G���[�R�[�h
'--------------- 2008/07/25 INSERT  END   By Systech ---------------

Public tWarpInitG() As typ_WarpKakuData     ' Warp�ް�(TBCMY018)      '05/12/18 ooba START ===>
Public tKakuInitG() As typ_WarpKakuData     ' �����p�x�ް�(TBCMY018)
Public tWarpMeasG() As typ_WarpKakuData     ' Warp�ް�(�\��/����p)
Public tKakuMeasG() As typ_WarpKakuData     ' �����p�x�ް�(�\��/����p)
Public tMapHinG     As typ_MapHinData       ' WFϯ�ߏ�̕i���ް�       '05/12/18 ooba END =====>

'------------------------------------------------
' �R�[�h�c�a�擾���ʊ֐�
'------------------------------------------------

'�T�v      :�w�肳�ꂽ���ڂ��L�[�ɁA�R�[�h�}�X�^�[(TBCMB005)����Y������f�[�^���擾����B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :sSysclass     ,I  ,String       :���ы敪('SB'�Œ�)
'          :sClass        ,I  ,String       :�敪
'          :sCode         ,I  ,String       :����
'          :iForm         ,I  ,Integer      :�擾�`��(0:50�޲��ް�, 1:1�޲��ް�)
'          :sSubCode      ,I  ,String       :��޺���(�擾�`��=1�̂ݗL��)
'          :sResult       ,O  ,String       :�擾�ް�
'          :�߂�l        ,O  ,Integer      :�擾�̐���(0:����擾, -1:�擾�װ)
'����      :
'����      :2003/09/04 �V�K�쐬�@�V�X�e���u���C��

Public Function funCodeDBGet(sSysclass As String, sClass As String, sCode As String, iForm As Integer, sSubCode As String, sResult As String) As Integer
    Dim sql As String       'SQL�S��
    Dim rs  As OraDynaset   'RecordSet

    '�p�����[�^�`�F�b�N
    If sSysclass = "" Or sSysclass = vbNullString Then GoTo CodeDBGetErr
    If sClass = "" Or sClass = vbNullString Then GoTo CodeDBGetErr
    If sCode = "" Or sCode = vbNullString Then GoTo CodeDBGetErr
    If iForm <> 0 And iForm <> 1 Then GoTo CodeDBGetErr
    If sSubCode = "" Or sSubCode = vbNullString Then GoTo CodeDBGetErr
    
    '�擾�`�� = 0(50�޲��ް�)�̏ꍇ
    If iForm = 0 Then
        sql = "select info1 from tbcmb005 where sysclass = '" & sSysclass & "' and class = '" & sClass & "' and code = '" & sCode & "'"
    
    '�擾�`�� = 1(1�޲��ް�)�̏ꍇ
    Else
        sql = "select substr(a1.info1, a2.info2, 1) as info1 from tbcmb005 a1, "
        sql = sql & "(select to_number(info2) as info2 from tbcmb005 "
        sql = sql & " where sysclass = '" & sSysclass & "' and class = '" & sClass & "' and code = '" & sSubCode & "') a2 "
        sql = sql & "where a1.sysclass = '" & sSysclass & "' and a1.class = '" & sClass & "' and a1.code = '" & sCode & "'"
    End If
    
    'SQL���̎��s
Debug.Print sql
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo CodeDBGetErr
    End If
    
    '�擾�f�[�^�Z�b�g
    sResult = rs("info1")
    Set rs = Nothing

    funCodeDBGet = 0
    
    Exit Function

CodeDBGetErr:
    funCodeDBGet = -1
    Set rs = Nothing
End Function

'�T�v      :�w�肳�ꂽ���ڂ��L�[�ɁA�}�g���b�N�X����OK/NG��Ԃ�
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :sSysclass     ,I  ,String       :���ы敪('SB'�Œ�)
'          :sClass        ,I  ,String       :�敪
'          :sCode1        ,I  ,String       :����1(�}�g���b�N�X�c��)
'          :sCode2        ,I  ,String       :����2(�}�g���b�N�X����)
'          :�߂�l        ,O  ,Integer      :�擾�̐���(1:OK(�����),0:NG(�����), -1:�擾�װ)
'����      :�R�[�hDB�ɓo�^����Ă���}�g���b�N�X��Code���擾���A
'           ���̃R�[�h�ɖ����l���w�肵���ꍇ�̓X�y�[�X�ɒu�������}�g���b�N�X����OK/NG���擾����
'����      :2006/02/10 �V�K�쐬�@SMP�ΐ�
Public Function funCodeDBGetMatrixReturn(sSysclass As String, sClass As String, sCode1 As String, sCode2 As String) As Integer
    Dim liRet           As Integer
    Dim sResult         As String       '�R�[�h�c�a�擾�֐��̎擾�ϐ�
    Dim lsCodeList()    As String       '�R�[�hDB��Code�ꗗ
    Dim llCnt           As Long
    Dim lsCode(1)       As String
    Dim liLoopCnt       As Integer
    
    funCodeDBGetMatrixReturn = -1
    
    lsCode(0) = Trim(sCode1)
    lsCode(1) = Trim(sCode2)
    
    '' �R�[�h�}�X�^�̃R�[�h�̈ꗗ���擾
    liRet = funCodeDBGetCodeList(sSysclass, sClass, lsCodeList)
    If liRet <> 0 Then
        funCodeDBGetMatrixReturn = -1
        Exit Function
    Else
        ''�R�[�h�}�X�^�ɓo�^����Ă��Ȃ��R�[�h�̓X�y�[�X�ɕϊ�����
        For liLoopCnt = 0 To 1
            liRet = 0
            For llCnt = 1 To UBound(lsCodeList)
                If Trim(lsCodeList(llCnt)) = Trim(lsCode(liLoopCnt)) Then
                    liRet = 1
                    Exit For
                End If
            Next llCnt
            If liRet = 0 Or Trim(lsCode(liLoopCnt)) = "" Then
                lsCode(liLoopCnt) = "     "
            End If
        Next liLoopCnt
        
        liRet = funCodeDBGet(sSysclass, sClass, lsCode(0), 1, lsCode(1), sResult)
        
        If liRet <> 0 Then
            funCodeDBGetMatrixReturn = -1
            Exit Function
        End If
        
        If sResult = 0 Then
            funCodeDBGetMatrixReturn = 0
            Exit Function
        End If
    End If
    
    funCodeDBGetMatrixReturn = 1
End Function



'�T�v      :�w�肳�ꂽ���ڂ��L�[�ɁA�R�[�h�}�X�^�[(TBCMB005)����CODE�̈ꗗ���擾����
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :sSysclass     ,I  ,String       :���ы敪('SB'�Œ�)
'          :sClass        ,I  ,String       :�敪
'          :sCode()       ,O  ,String       :���ނ̈ꗗ
'          :�߂�l        ,O  ,Integer      :�擾�̐���(0:����擾, -1:�擾�װ)
'����      :
'����      :2006/02/10 �V�K�쐬�@SMP�ΐ�

Public Function funCodeDBGetCodeList(sSysclass As String, sClass As String, sCode() As String) As Integer
    Dim sql As String       'SQL�S��
    Dim rs  As OraDynaset   'RecordSet

    '�p�����[�^�`�F�b�N
    If sSysclass = "" Or sSysclass = vbNullString Then GoTo CodeDBGetErr
    If sClass = "" Or sClass = vbNullString Then GoTo CodeDBGetErr
    
    '������
    ReDim sCode(0) As String
    
    sql = ""
    sql = sql & " SELECT"
    sql = sql & "   code"
    sql = sql & " FROM"
    sql = sql & "   tbcmb005"
    sql = sql & " WHERE sysclass = '" & sSysclass & "'"
    sql = sql & "   AND class    = '" & sClass & "'"
    
    'SQL���̎��s
Debug.Print sql
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    Do Until rs.EOF '�f�[�^���Ȃ��Ȃ�܂Ŏ擾
        ReDim Preserve sCode(UBound(sCode) + 1) As String
        '�擾�f�[�^�Z�b�g
        sCode(UBound(sCode)) = Trim(rs("code"))
        rs.MoveNext
    Loop
    
    '�f�[�^�����̏ꍇ
    If UBound(sCode) = 0 Then
        GoTo CodeDBGetErr
    End If
    
    Set rs = Nothing

    funCodeDBGetCodeList = 0
    
    Exit Function

CodeDBGetErr:
    funCodeDBGetCodeList = -1
    Set rs = Nothing
End Function

'------------------------------------------------
' ��������قƂv�e����ٕR�t�����ʊ֐�
'------------------------------------------------

'�T�v      :�w�肳�ꂽWF����ُ�񂩂�A�Ή����錋������ق��������Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^               :����
'          :tWfSmpl       ,I  ,typ_Wf_Smpl      :WF����ٍ\����
'          :tCrySmpl      ,O  ,typ_Cry_Smpl     :��������ٍ\����
'          :�߂�l        ,O  ,Integer          :�擾�̐���(0:����擾, -1:�Y����ۯ��Ȃ�)
'����      :
'����      :2003/09/04 �V�K�쐬�@�V�X�e���u���C��

Public Function funConSxl_Wf_Sampl(tWfSmpl As typ_Wf_Smpl, tCrySmpl As typ_Cry_Smpl) As Integer
    Dim sql     As String       'SQL�S��
    Dim rs      As OraDynaset   'RecordSet

    'SQL���̕ҏW
    '��ۯ�ID,�����敪�����ǉ� 08/08/11 ooba
    sql = "select CRYNUMCS, SMPKBNCS, TBKBNCS, REPSMPLIDCS, XTALCS, INPOSCS, HINBCS, REVNUMCS, FACTORYCS, OPECS from XSDCS "

    'TOP�ʒu(T/B�敪='T')�̌���
    If tWfSmpl.TBKBNCW = "T" Then
        sql = sql & "where crynumcs like '" & Mid(tWfSmpl.XTALCW, 1, 9) & "%' and "
        sql = sql & "      tbkbncs = '" & tWfSmpl.TBKBNCW & "' and "
        sql = sql & "      xtalcs = '" & tWfSmpl.XTALCW & "' and "
        sql = sql & "      livkcs = '0' and "
        sql = sql & "      inposcs = (select max(inposcs) from xsdcs "
        sql = sql & "                 where  crynumcs like '" & Mid(tWfSmpl.XTALCW, 1, 9) & "%' and "
        sql = sql & "                        tbkbncs = '" & tWfSmpl.TBKBNCW & "' and "
        sql = sql & "                        xtalcs = '" & tWfSmpl.XTALCW & "' and "
        sql = sql & "                        livkcs = '0' and "
        sql = sql & "                        inposcs <= '" & tWfSmpl.INPOSCW & "')"
    
    'BOT�ʒu(T/B�敪='B')�̌���
    ElseIf tWfSmpl.TBKBNCW = "B" Then
        sql = sql & "where crynumcs like '" & Mid(tWfSmpl.XTALCW, 1, 9) & "%' and "
        sql = sql & "      tbkbncs = '" & tWfSmpl.TBKBNCW & "' and "
        sql = sql & "      xtalcs = '" & tWfSmpl.XTALCW & "' and "
        sql = sql & "      livkcs = '0' and "
        sql = sql & "      inposcs = (select min(inposcs) from xsdcs "
        sql = sql & "                 where  crynumcs like '" & Mid(tWfSmpl.XTALCW, 1, 9) & "%' and "
        sql = sql & "                        tbkbncs = '" & tWfSmpl.TBKBNCW & "' and "
        sql = sql & "                        xtalcs = '" & tWfSmpl.XTALCW & "' and "
        sql = sql & "                        livkcs = '0' and "
        sql = sql & "                        inposcs >= '" & tWfSmpl.INPOSCW & "')"
    Else
        funConSxl_Wf_Sampl = -1
        Exit Function
    End If
    
    'SQL���̎��s
Debug.Print sql
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '�Y���f�[�^�Ȃ�
    If rs.EOF Or rs.RecordCount > 1 Then
        funConSxl_Wf_Sampl = -1
        Exit Function
    End If
    
    '�擾�f�[�^�Z�b�g
    With tCrySmpl
        .CRYNUMCS = rs("CRYNUMCS")          ' ��ۯ�ID
        .SMPKBNCS = rs("SMPKBNCS")          ' ����ً敪
        .TBKBNCS = rs("TBKBNCS")            ' T/B�敪
        .REPSMPLIDCS = rs("REPSMPLIDCS")    ' ��\�����ID
        .XTALCS = rs("XTALCS")              ' �����ԍ�
        .INPOSCS = rs("INPOSCS")            ' �������ʒu
        .HINBCS = rs("HINBCS")              ' �i��
        .REVNUMCS = rs("REVNUMCS")          ' ���i�ԍ������ԍ�
        .FACTORYCS = rs("FACTORYCS")        ' �H��
        .OPECS = rs("OPECS")                ' ���Ə���
    End With
    Set rs = Nothing

    funConSxl_Wf_Sampl = 0

End Function

'�T�v      :�H�����т���U�֎��̌������ʒu/����/�����ԍ����擾����
'���Ұ�    :�ϐ���          ,IO  ,�^                :����
'          :sLotid         ,I   ,String            :��ۯ�ID or SXL_ID
'          :iKcnt          ,I   ,Integer           :�H���A��
'          :iIngotpos       ,O   ,Integer          :�������ʒu
'          :iLength         ,O   ,Integer          :����
'          :sCrynum         ,O   ,String           :�����ԍ�
'          :�߂�l          ,O   ,FUNCTION_RETURN   :���o�̐���
'����      :
'����      :2003/11/07 ooba
Public Function GET_hurikaeC3(sLotid As String, iKcnt As Integer, iIngotPos As Integer, _
                                iLength As Integer, sCryNum As String) As FUNCTION_RETURN

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmec064.bas -- Function GET_hurikaeC3"
    GET_hurikaeC3 = FUNCTION_RETURN_FAILURE

    Dim sSql As String
    Dim rs As OraDynaset
    
    GET_hurikaeC3 = FUNCTION_RETURN_FAILURE
    
    sSql = "select min(INPOSC3), sum(LENC3), XTALC3 "
    sSql = sSql & "from XSDC3 "
    If Len(sLotid) = 12 Then
        sSql = sSql & "where CRYNUMC3 = '" & sLotid & "' "
    ElseIf Len(sLotid) = 13 Then
        sSql = sSql & "where SXLIDC3 = '" & sLotid & "' "
    Else
        GET_hurikaeC3 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    sSql = sSql & "and KCNTC3 = " & iKcnt
    sSql = sSql & "and substr(KNKTC3, 5, 1) = '3' "
    sSql = sSql & "group by XTALC3 "
    
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    
    If rs.RecordCount <> 1 Then
        Set rs = Nothing
        GET_hurikaeC3 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    Else
        If IsNull(rs("min(INPOSC3)")) = False Then iIngotPos = rs("min(INPOSC3)")
        If IsNull(rs("sum(LENC3)")) = False Then iLength = rs("sum(LENC3)")
        If IsNull(rs("XTALC3")) = False Then sCryNum = rs("XTALC3")
    End If

    Set rs = Nothing
    
    GET_hurikaeC3 = FUNCTION_RETURN_SUCCESS


proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    Resume proc_exit

End Function

'�T�v      :SXL�m��w��(TBCMY007)ð��قɾ�Ă���SXL�̔��R�ް��擾����݂����߂�B
'���Ұ�    :�ϐ���          ,IO  ,�^                :����
'          :HIN            ,I   ,tFullHinban       ,12���i��
'�@�@      :sPos  �@�@�@    ,I   ,String �@         ,SXL�ʒu(TOP/BOT)   04/04/15 ooba
'          :sPattern       ,O   ,String            ,���R�ް��擾�����
'                                                   �������A : WF�����ް��擾
'                                                   �������B : ���������ް��擾
'                                                   �������C : �擾�ް��Ȃ�
'          :�߂�l          ,O   ,FUNCTION_RETURN   ,���o�̐���
'����      :
'����      :04/02/12 ooba�@�쐬
Public Function SxlRsPattern(HIN As tFullHinban, sPos As String, sPattern As String) As FUNCTION_RETURN

    Dim HSXRHWYS As String      '�i�r�w���R�ۏؕ��@�Q��
    Dim HWFRHWYS As String      '�i�v�e���R�ۏؕ��@�Q��
    Dim HWFRKHNN As String      '�i�v�e���R�����p�x�Q���@04/04/15 ooba
    Dim sSql As String
    Dim rs As OraDynaset
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "SB_com.bas -- Function SxlRsPattern"
    SxlRsPattern = FUNCTION_RETURN_FAILURE
    
    sPattern = "C"
    
    If Trim(HIN.hinban) <> "" And Trim(HIN.hinban) <> "Z" Then
        '�Y���i�Ԃ̔��R(Rs)�d�l���擾
        sSql = "select HSXRHWYS, HWFRHWYS, HWFRKHNN "
        sSql = sSql & "from TBCME018, TBCME021 "
        sSql = sSql & "where TBCME018.HINBAN = TBCME021.HINBAN "
        sSql = sSql & "and TBCME018.MNOREVNO = TBCME021.MNOREVNO "
        sSql = sSql & "and TBCME018.FACTORY = TBCME021.FACTORY "
        sSql = sSql & "and TBCME018.OPECOND = TBCME021.OPECOND "
        sSql = sSql & "and TBCME018.HINBAN = '" & HIN.hinban & "' "
        sSql = sSql & "and TBCME018.MNOREVNO = " & HIN.mnorevno & " "
        sSql = sSql & "and TBCME018.FACTORY = '" & HIN.factory & "' "
        sSql = sSql & "and TBCME018.OPECOND = '" & HIN.opecond & "' "
        
        Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
        
        If rs.RecordCount <> 1 Then
            Set rs = Nothing
            GoTo proc_exit
        Else
            If IsNull(rs("HSXRHWYS")) = False Then HSXRHWYS = rs("HSXRHWYS")   '�i�r�w���R�ۏؕ��@�Q��
            If IsNull(rs("HWFRHWYS")) = False Then HWFRHWYS = rs("HWFRHWYS")   '�i�v�e���R�ۏؕ��@�Q��
            If IsNull(rs("HWFRKHNN")) = False Then HWFRKHNN = rs("HWFRKHNN")   '�i�v�e���R�����p�x�Q���@04/04/15 ooba
        End If
        
        Set rs = Nothing
    Else
        GoTo proc_exit
    End If
    
    '�ۏؕ��@�����ǉ��@04/04/15 ooba
'    If HWFRHWYS = "H" Then
    If HWFRHWYS = "H" And CheckKHN(HWFRKHNN, 1, sPos) Then
        'WF�d�l�wH�x�̏ꍇ
        sPattern = "A"
    ElseIf HWFRHWYS = "S" And CheckKHN(HWFRKHNN, 1, sPos) Then
        If HSXRHWYS = "H" Then
            'WF�d�l�wS�x�Ō����d�l�wH�x�̏ꍇ
            sPattern = "B"
        Else
            'WF�d�l�wS�x�Ō����d�l�wH�x�ȊO�̏ꍇ
            sPattern = "A"
        End If
    Else
        If HSXRHWYS = "H" Or HSXRHWYS = "S" Then
            'WF�d�l�Ȃ��Ō����d�l�wH�x�wS�x�̏ꍇ
            sPattern = "B"
        End If
    End If
    
    Set rs = Nothing
    SxlRsPattern = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    Resume proc_exit
    
End Function

'�T�v      :�ۏؕ��@(�����p�x�Q��)�̃`�F�b�N
'���Ұ��@�@:�ϐ���      ,IO ,�^       ,����
'�@�@      :sKHN  �@�@�@,I  ,String �@,�ۏؕ��@(�����p�x�Q��)
'�@�@      :iItemNo     ,I  ,Integer  ,�������ځ@1:Rs  2:Oi  3:OSF1  4:OSF2  5:OSF3  6:OSF4
'                                               7:BMD1  8:BMD2  9:BMD3  10:Doi1  11:Doi2
'                                               12:Doi3 13:Dsod  14:DZ  15:SPVFE  16:SPV�g
'                                               17:Aoi 18:GD 19:SPVNR
'�@�@      :sPos  �@�@�@,I  ,String �@,SXL�ʒu(TOP/BOT/MID)
'      �@�@:�߂�l      ,O  ,Boolean�@,�����̗L��
'����      :�ۏؕ��@(�����p�x�Q��)���`�F�b�N���Č����̗L����Ԃ�
'����      :04/04/07 ooba
'          :GD�ǉ��@05/01/20 ooba
'          :SPVNR�ǉ��@06/06/08 ooba
Public Function CheckKHN(sKHN As String, iItemNo As Integer, sPos As String) As Boolean
    Dim RET     As Integer
    Dim sChkPtn As String
    Dim sResult As String
    Dim iChk    As Integer
    
'    CheckKHN = False
    CheckKHN = True '04/05/26 ooba
    sChkPtn = ""
    sResult = ""
    If sPos <> "TOP" And sPos <> "BOT" Then Exit Function
    
    '����DB���ۏؕ��@�����̗L�������擾����
    RET = funCodeDBGet("SB", "HO", "PTN", 0, " ", sChkPtn)
    If RET <> 0 Then Exit Function
    
    If Mid(sChkPtn, iItemNo, 1) = "1" Then
        '����DB���ۏؕ��@��������݂��擾����
        RET = funCodeDBGet("SB", "HO", sPos, 0, " ", sResult)
        If RET <> 0 Then Exit Function
        
        '�擾����݂�茟���̗L���𔻒f����
        Select Case sKHN
        Case "3"    'TOP�ۏ�
            iChk = 1
        Case "4"    'BOT�ۏ�
            iChk = 2
        Case "6"    'T/B�ۏ�
            iChk = 3
        Case Else   '�Ȃ�(NULL,��߰�,346�ȊO)
            iChk = 4
        End Select
        
        If Mid(sResult, iChk, 1) = "1" Then CheckKHN = True Else CheckKHN = False
    Else
        CheckKHN = True
    End If
    
End Function

'�T�v      :�ۏؕ��@(�����p�x�Q��)�̃`�F�b�N(�G�s�p)
'���Ұ��@�@:�ϐ���      ,IO ,�^       ,����
'�@�@      :sKHN  �@�@�@,I  ,String �@,�ۏؕ��@(�����p�x�Q��)
'�@�@      :iItemNo     ,I  ,Integer  ,�������ځ@1:BMD1E  2:BMD2E  3:BMD3E  4:OSF1E  5:OSF2E  6:OSF3E
'�@�@      :sPos  �@�@�@,I  ,String �@,SXL�ʒu(TOP/BOT)
'      �@�@:�߂�l      ,O  ,Boolean�@,�����̗L��
'����      :�ۏؕ��@(�����p�x�Q��)���`�F�b�N���Č����̗L����Ԃ�
'����      :06/08/15 SMP)kondoh �V�K�쐬
Public Function CheckKHN_EP(sKHN As String, iItemNo As Integer, sPos As String) As Boolean
    Dim RET     As Integer
    Dim sChkPtn As String
    Dim sResult As String
    Dim iChk    As Integer
    
    CheckKHN_EP = True
    sChkPtn = ""
    sResult = ""
    If sPos <> "TOP" And sPos <> "BOT" Then Exit Function
    
    '����DB���ۏؕ��@�����̗L�������擾����
    RET = funCodeDBGet("SB", "HO", "PTNE", 0, " ", sChkPtn)
    If RET <> 0 Then Exit Function
    
    If Mid(sChkPtn, iItemNo, 1) = "1" Then
        '����DB���ۏؕ��@��������݂��擾����
        RET = funCodeDBGet("SB", "HO", sPos, 0, " ", sResult)
        If RET <> 0 Then Exit Function
        
        '�擾����݂�茟���̗L���𔻒f����
        Select Case sKHN
        Case "3"    'TOP�ۏ�
            iChk = 1
        Case "4"    'BOT�ۏ�
            iChk = 2
        Case "6"    'T/B�ۏ�
            iChk = 3
        Case Else   '�Ȃ�(NULL,��߰�,346�ȊO)
            iChk = 4
        End Select
        
        If Mid(sResult, iChk, 1) = "1" Then CheckKHN_EP = True Else CheckKHN_EP = False
    Else
        CheckKHN_EP = True
    End If
    
End Function

'�T�v      :��ۯ��P�ʕۏ��׸ނ̎擾
'���Ұ��@�@:�ϐ���      ,IO   ,�^                ,����
'�@�@      :HIN  �@�@ �@,I    ,tFullHinban �@    ,12���i��
'�@�@      :sBflg  �@�@ ,O    ,String �@         ,��ۯ��P�ʕۏ��׸�
'      �@�@:�߂�l      ,O    ,FUNCTION_RETURN   ,���o�̐���
'����      :
'����      :05/01/20 ooba
Public Function chkBlkTanFlg(HIN As tFullHinban, sBflg As String) As FUNCTION_RETURN

    Dim sSql As String
    Dim rs As OraDynaset
    
    chkBlkTanFlg = FUNCTION_RETURN_FAILURE
        
    sBflg = ""
    If Trim(HIN.hinban) = "Z" Or Trim(HIN.hinban) = "G" Or Trim(HIN.hinban) = "" Then
        Exit Function
    End If
    
    sSql = "SELECT BLOCKHFLAG"
    sSql = sSql & " FROM TBCME036"
    sSql = sSql & " WHERE"
    sSql = sSql & " HINBAN = '" & HIN.hinban & "'"
    sSql = sSql & " AND MNOREVNO = " & HIN.mnorevno
    sSql = sSql & " AND FACTORY = '" & HIN.factory & "'"
    sSql = sSql & " AND OPECOND = '" & HIN.opecond & "'"
    
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)

    If rs.RecordCount > 0 Then
        If IsNull(rs.Fields("BLOCKHFLAG")) = False Then sBflg = rs.Fields("BLOCKHFLAG")
        chkBlkTanFlg = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
End Function

'�T�v      :WF��ĒP�ʂ̎擾
'���Ұ��@�@:�ϐ���      ,IO   ,�^                ,����
'�@�@      :HIN  �@�@ �@,I    ,tFullHinban �@    ,12���i��
'�@�@      :iWCtani  �@ ,O    ,Integer �@        ,WF��ĒP��
'      �@�@:�߂�l      ,O    ,FUNCTION_RETURN   ,���o�̐���
'����      :
'����      :05/04/19 ooba
Public Function getWFCUTT(HIN As tFullHinban, iWCtani As Integer) As FUNCTION_RETURN

    Dim sSql As String
    Dim rs As OraDynaset
    
    getWFCUTT = FUNCTION_RETURN_FAILURE
        
    iWCtani = -1
    If Trim(HIN.hinban) = "Z" Or Trim(HIN.hinban) = "G" Or Trim(HIN.hinban) = "" Then
        Exit Function
    End If
    
    sSql = "SELECT WFCUTT"
    sSql = sSql & " FROM TBCME036"
    sSql = sSql & " WHERE"
    sSql = sSql & " HINBAN = '" & HIN.hinban & "'"
    sSql = sSql & " AND MNOREVNO = " & HIN.mnorevno
    sSql = sSql & " AND FACTORY = '" & HIN.factory & "'"
    sSql = sSql & " AND OPECOND = '" & HIN.opecond & "'"
    
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)

    If rs.RecordCount > 0 Then
        If IsNull(rs.Fields("WFCUTT")) = False Then iWCtani = rs.Fields("WFCUTT") Else iWCtani = -1
        getWFCUTT = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
End Function
'�T�v      :SIRD�]�����̎擾
'���Ұ��@�@:�ϐ���      ,IO   ,�^                ,����
'�@�@      :HIN  �@�@ �@,I    ,tFullHinban �@    ,12���i��
'�@�@      :sSIRDFLG �@ ,O    ,String �@   �@    ,SIRD�]���t���O
'      �@�@:�߂�l      ,O    ,FUNCTION_RETURN   ,���o�̐���
'����      :
'����      :2010/01/18 Y.Hitomi
Public Function getSDFlg(HIN As tFullHinban, sSirdFlg As String) As FUNCTION_RETURN

    Dim sSql As String
    Dim rs As OraDynaset
    
    getSDFlg = FUNCTION_RETURN_FAILURE
        
    sSirdFlg = " "
    If Trim(HIN.hinban) = "Z" Or Trim(HIN.hinban) = "G" Or Trim(HIN.hinban) = "" Then
        Exit Function
    End If
    
    sSql = "SELECT HWFSIRDHS"
    sSql = sSql & " FROM TBCME048"
    sSql = sSql & " WHERE"
    sSql = sSql & " HINBAN = '" & HIN.hinban & "'"
    sSql = sSql & " AND MNOREVNO = " & HIN.mnorevno
    sSql = sSql & " AND FACTORY = '" & HIN.factory & "'"
    sSql = sSql & " AND OPECOND = '" & HIN.opecond & "'"
    
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)

    If rs.RecordCount > 0 Then
        If IsNull(rs.Fields("HWFSIRDHS")) = False Then sSirdFlg = rs.Fields("HWFSIRDHS") Else sSirdFlg = " "
        getSDFlg = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
End Function

'�T�v      :�����������т̈��p������
'���Ұ��@�@:�ϐ���      ,IO   ,�^                ,����
'�@�@      :sBlockID  �@,I    ,String      �@    ,��ۯ�ID
'�@�@      :sTBkbn  �@�@,I    ,String      �@    ,TB�敪
'�@�@      :iInPos  �@�@,I    ,Integer�@�@ �@    ,�������ʒu
'�@�@      :HIN  �@�@ �@,I    ,tFullHinban �@    ,12���i��
'�@�@      :iItemNo   �@,I    ,Integer �@        ,�������ځ@(1:GD)
'�@�@      :sSampleID   ,O    ,String �@         ,���������ID
'�@�@      :sCryind     ,O    ,String �@         ,�������FLG
'      �@�@:�߂�l      ,O    ,FUNCTION_RETURN   ,���o�̐���
'����      :��ۯ�ID�^�i�ԁ^TB�敪�^�������ʒu�����ɻ���يǗ�(��ۯ�)���猋�������ID�^�������FLG���擾����
'����      :05/01/20 ooba
Public Function funBlkSmpDataGet(sBlockId As String, sTBkbn As String, _
                                    iInpos As Integer, HIN As tFullHinban, _
                                    iItemNo As Integer, sSampleid As String, _
                                    sCryind As String) As FUNCTION_RETURN
                                                          
    Dim sKensa As String        '�������ږ�
    Dim sSql As String
    Dim rs As OraDynaset
    
    funBlkSmpDataGet = FUNCTION_RETURN_FAILURE
        
    sSampleid = ""
    sCryind = ""
    
    '�������ڂ��
    Select Case iItemNo
    Case 1  'GD
        sKensa = "GD"
    Case Else
        Exit Function
    End Select
    
    '���������ID�̎擾
    sSql = "SELECT"
    sSql = sSql & " CRYSMPLID" & sKensa & "CS, "                    '�����ID
    sSql = sSql & " CRYIND" & sKensa & "CS"                         '���FLG
    sSql = sSql & " FROM XSDCS"
    sSql = sSql & " WHERE"
    sSql = sSql & " CRYNUMCS = '" & sBlockId & "'"                  '��ۯ�ID
    sSql = sSql & " AND TBKBNCS = '" & sTBkbn & "'"                 'T/B�敪
    If sTBkbn = "T" Then                                            '�������ʒu
        sSql = sSql & " AND INPOSCS <= " & iInpos
    ElseIf sTBkbn = "B" Then
        sSql = sSql & " AND INPOSCS >= " & iInpos
    End If
    If Trim(HIN.hinban) <> "" Then      'CW740/CW760�̏ꍇ�͕i�Ԃ̏��������@05/06/13 ooba
        sSql = sSql & " AND HINBCS = '" & HIN.hinban & "'"          '�i��
        sSql = sSql & " AND REVNUMCS = " & HIN.mnorevno             '���i�ԍ������ԍ�
        sSql = sSql & " AND FACTORYCS = '" & HIN.factory & "'"      '�H��
        sSql = sSql & " AND OPECS = '" & HIN.opecond & "'"          '���Ə���
    End If
    sSql = sSql & " AND CRYIND" & sKensa & "CS IN ('1', '2')"       '���FLG
'    sSql = sSql & " AND CRYRES" & sKensa & "CS = '1'"               '����FLG
    '�����ύX�@05/07/20 ooba
    sSql = sSql & " AND CRYRES" & sKensa & "CS IN ('1', '2')"       '����FLG
    
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    
    '���������ID�^�������FLG���
    If rs.RecordCount > 0 Then
        If IsNull(rs.Fields(0)) = False Then sSampleid = CStr(rs.Fields(0))
        If IsNull(rs.Fields(1)) = False Then sCryind = CStr(rs.Fields(1))
        funBlkSmpDataGet = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
End Function

'------------------------------------------------
' Null�������ʊ֐�
'------------------------------------------------

'�T�v      :�w�肳�ꂽ�l��Null�Ȃ�-1��Ԃ��ANull�ȊO�Ȃ�w�肳�ꂽ�l��Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :vrnN          ,I  ,Variant      :�w��l
'          :�߂�l        ,O  ,Double       :�w��l�A�܂��́A-1
'����      :
'����      :2003/12/08 �V�K�쐬�@�V�X�e���u���C��

Public Function fncNullCheck(vrnN As Variant) As Double 'Null�̃`�F�b�N������
    If IsNull(vrnN) = False Then
        fncNullCheck = vrnN 'NULL����Ȃ��Ƃ��͂��̂܂�
    Else
        fncNullCheck = -1  'NULL�̂Ƃ���-1������
    End If
End Function

'------------------------------------------------
' Null�Ή��\�����ʊ֐�
'------------------------------------------------

'�T�v      :�w�肳�ꂽ�l��-1(Null�l�̑���)�Ȃ�vbNullString��Ԃ��A-1�ȊO�Ȃ�w��l��Ԃ��B̫�ϯĎw�肪����ꍇ�ɂ́A�w��̫�ϯĂŕԂ��B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :data          ,I  ,Variant      :�w��l
'          :Formatstr     ,I  ,String       :̫�ϯČ`��(�ȗ���)
'          :�߂�l        ,O  ,Variant      :�߂�l
'����      :
'����      :2003/12/09 �V�K�쐬�@�V�X�e���u���C��

Public Function DBData2DispData_nl(data As Variant, Optional Formatstr As String) As Variant   'NULL�Ή��p 2003/12/9
    If data = -1 Then
'        DBData2DispData = ""
        DBData2DispData_nl = vbNullString
    Else
        If Formatstr = "" Then
            DBData2DispData_nl = data
        Else
            DBData2DispData_nl = Format(data, Formatstr)
        End If
    End If
End Function

'''��������s_cmzcjudg.bas �Ɉړ�            2003/12/18 tuku
'''�T�v      :����l��NULL�������ꍇ�͈͔̔�����s���B
'''���Ұ�    :�ϐ���        ,IO ,�^        ,����
'''          :JudgData      ,I  ,double    ,����l
'''          :SpecMin       ,I  ,double    ,�����l
'''          :SpecMax       ,I  ,double    ,����l
'''          :�߂�l        ,O  ,Boolean   ,JUDG_OK or JUDG_NG
'''����      :
'''����      :2003/12/11 �V�K�쐬 �V�X�e���u���C��
''Public Function RangeDecision_nl(JudgData As Double, SpecMin As Double, SpecMax As Double) As Boolean
''    RangeDecision_nl = False
''    If (JudgData >= SpecMin) Or (SpecMin = -1) Then
''        If (JudgData <= SpecMax) Or (SpecMax = -1) Then
''            RangeDecision_nl = True
''        End If
''    End If
'''    RangeDecision = ((JudgData >= SpecMin) And (JudgData <= SpecMax))
''End Function



'------------------------------------------------
' Null�Ή����ѓ��͔��苤�ʊ֐�
'------------------------------------------------

'�T�v      :�ۏؕ��@��'H'�܂���'S'�ŁA�d�l�l�z���-1�����݂����ꍇ�AFalse��Ԃ��B����ȊO�̏ꍇ�ATrue��Ԃ��B
'���Ұ�    :�ϐ���        ,IO ,�^           :����
'          :Hosyo         ,I  ,String       :�ۏؕ��@_�Ώ�
'          :Shiyo()       ,I  ,Double       :�d�l�l�z��
'          :�߂�l        ,O  ,Boolean      : True:OK, False:NG
'����      :
'����      :2003/12/11 �V�K�쐬�@�V�X�e���u���C��

Public Function fncJissekiHantei_nl(Hosyo As String, Shiyo() As Double) As Boolean
    Dim cnt As Integer
    
    fncJissekiHantei_nl = True
'    If Hosyo = "H" Or Hosyo = "S" Then '�ۏؕ��@S�̓`�F�b�N���Ȃ��@2003/12/19�@tuku
    If Hosyo = "H" Then
        For cnt = 1 To UBound(Shiyo)
            If Shiyo(cnt) = -1 Then
                fncJissekiHantei_nl = False
                Exit For
            End If
        Next
    End If
End Function

'�T�v      :�_���i�Ԃ��擾����B
'���Ұ��@�@:�ϐ���      ,IO   ,�^                ,����
'�@�@      :sBlockID  �@,I    ,String      �@    ,�����ԍ�
'�@�@      :HIN  �@�@ �@,O    ,tFullHinban �@    ,12���_���i��
'      �@�@:�߂�l      ,O    ,FUNCTION_RETURN   ,���o�̐���
'����      :
'����      :06/04/25 ooba
Public Function funNeraiHinGet(sCryNum As String, tHIN As tFullHinban) As FUNCTION_RETURN

    Dim sSql As String
    Dim rs As OraDynaset
    
    funNeraiHinGet = FUNCTION_RETURN_FAILURE
        
    tHIN.hinban = ""
    tHIN.mnorevno = 0
    tHIN.factory = ""
    tHIN.opecond = ""
    tHIN.Hinkubun = ""
    
    sSql = "SELECT PUHINBC1, PUREVNUMC1, PUFACTORYC1, PUOPEC1 "
    sSql = sSql & "FROM XSDC1 "
    sSql = sSql & "WHERE XTALC1 = '" & sCryNum & "' "
    
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)

    If rs.RecordCount = 1 Then
        If Not IsNull(rs.Fields("PUHINBC1")) Then tHIN.hinban = rs.Fields("PUHINBC1")
        If Not IsNull(rs.Fields("PUREVNUMC1")) Then tHIN.mnorevno = rs.Fields("PUREVNUMC1")
        If Not IsNull(rs.Fields("PUFACTORYC1")) Then tHIN.factory = rs.Fields("PUFACTORYC1")
        If Not IsNull(rs.Fields("PUOPEC1")) Then tHIN.opecond = rs.Fields("PUOPEC1")
        funNeraiHinGet = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close

End Function


'�T�v      :��������(XSDC1)���C�|OSF3����ID�̊l�����s��
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :XTALC1        ,IO ,String           ,�����ԍ�
'          :JDGEIDC1      ,O  ,String           ,C�|OSF3����ID
'����      :
'����      :2007/04/23 �쐬  ����
Public Function GetCOSF3ID(C_JDGEIDC1 As String, C_XTALC1 As String) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String

    '������
    GetCOSF3ID = FUNCTION_RETURN_FAILURE

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''��������(XSDC1)���C�|OSF3����ID�̊l�����s��
    sql = ""
    sql = "select XTALC1,JDGEIDC1 from XSDC1 where (trim(XTALC1)='" & Trim$(C_XTALC1) & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    
    'ں��ގ��̂����݂��Ȃ��ꍇ
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GetCOSF3ID = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    Else
        '�����ԍ�
        C_XTALC1 = rs("XTALC1")
        'C�|OSF3����ID��NULL�̏ꍇ
        If Trim(rs("JDGEIDC1")) = "" Or IsNull(rs("JDGEIDC1")) Then
            C_JDGEIDC1 = vbNullString
        Else
            C_JDGEIDC1 = rs("JDGEIDC1")
        End If
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

    '����I��
    GetCOSF3ID = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

End Function


'�T�v      :�������(XODC5_OSF30)��菳�F�׸ނ̊l�����s��
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :JDGEIDC5      ,IO ,String           ,����ID
'          :SYNFLAGC5     ,O  ,String           ,���F�׸�
'          :YNKFLAGC5     ,O  ,String           ,�폜�׸�
'����      :
'����      :2007/04/23 �쐬  ����
Public Function GetSYNFLAGC5(C_SYNFLAGC5 As String, C_YMKFLAGC5 As String, C_JDGEIDC1 As String) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String

    '������
    GetSYNFLAGC5 = FUNCTION_RETURN_FAILURE

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''��������(XSDC1)���C�|OSF3����ID�̊l�����s��
    sql = ""
    sql = "select SYNFLAGC5,YMKFLAGC5 from XODC5_OSF30 where (trim(JDGEIDC5)='" & Trim$(C_JDGEIDC1) & "') and YMKFLAGC5 = '0'"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    
    'ں��ގ��̂����݂��Ȃ��ꍇ
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GetSYNFLAGC5 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    Else
        '���F�׸ނ�NULL�̏ꍇ
        If Trim(rs("SYNFLAGC5")) = "" Or IsNull(rs("SYNFLAGC5")) Then
            C_SYNFLAGC5 = vbNullString
        Else
            C_SYNFLAGC5 = rs("SYNFLAGC5")
        End If
        '�폜�׸ނ�NULL�̏ꍇ
        If Trim(rs("YMKFLAGC5")) = "" Or IsNull(rs("YMKFLAGC5")) Then
            C_YMKFLAGC5 = vbNullString
        Else
            C_YMKFLAGC5 = rs("YMKFLAGC5")
        End If

    End If
    
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If
    
    '����I��
    GetSYNFLAGC5 = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

End Function

'�T�v      :�������(XODC5_OSF31)��蔻��f�[�^�̊l�����s��
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :strChkR       ,I  ,String           ,����݋敪
'          :strChkD       ,I  ,String           ,����݋敪
'          :JDGEIDC5      ,IO ,String           ,����ID
'          :POSC5         ,IO ,Long             ,����وʒu
'          :DMAXC5        ,O  ,Long             ,D�̂ݏ��
'          :RMAXC5        ,O  ,Long             ,R�̂ݏ��
'          :DRDMAXC5      ,O  ,Long             ,����D���
'          :DRRMAXC5      ,O  ,Long             ,����R���
'����      :
'����      :2007/04/23 �쐬  ����

Public Function GetCOSF3PTN(C_JDGEIDC5 As String, C_POSC5 As Long, C_strChkR As String, C_strChkD As String, C_RMAXC5 As String, C_DMAXC5 As String, C_DRRMAXC5 As String, C_DRDMAXC5 As String) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String


    '������
    GetCOSF3PTN = FUNCTION_RETURN_FAILURE

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    '�������(XODC5_OSF31)��蔻��f�[�^�̊l�����s��
    '���Ұ�������݋敪�ɂ���ď�������
    'R�݂̂̏ꍇ
    If C_strChkR = "R" And C_strChkD = "-" Then
        'SQL�ҏW
        sql = ""
        sql = "SELECT TO_CHAR(RMAXC5) as RMAXC5 FROM(select MIN(POSC5) as W_POSC5 from(select POSC5 from XODC5_OSF31 where (trim(JDGEIDC5)='" & Trim$(C_JDGEIDC5) & "') and POSC5 >='" & Trim$(C_POSC5) & "')  ),XODC5_OSF31 WHERE POSC5 = W_POSC5 AND (trim(JDGEIDC5)='" & Trim$(C_JDGEIDC5) & "')"
    'D�݂̂̏ꍇ
    ElseIf C_strChkR = "D" Then
        'SQL�ҏW
        sql = ""
        sql = "SELECT TO_CHAR(DMAXC5) as DMAXC5 FROM(select MIN(POSC5) as W_POSC5 from(select POSC5 from XODC5_OSF31 where (trim(JDGEIDC5)='" & Trim$(C_JDGEIDC5) & "') and POSC5 >='" & Trim$(C_POSC5) & "')  ),XODC5_OSF31 WHERE POSC5 = W_POSC5 AND (trim(JDGEIDC5)='" & Trim$(C_JDGEIDC5) & "')"
    'R&D�̏ꍇ
    ElseIf C_strChkR = "R" And C_strChkD = "D" Then
        'SQL�ҏW
        sql = ""
        sql = "SELECT TO_CHAR(DRRMAXC5) as DRRMAXC5,TO_CHAR(DRDMAXC5) as DRDMAXC5 FROM(select MIN(POSC5) as W_POSC5 from(select POSC5 from XODC5_OSF31 where (trim(JDGEIDC5)='" & Trim$(C_JDGEIDC5) & "') and POSC5 >='" & Trim$(C_POSC5) & "')  ),XODC5_OSF31 WHERE POSC5 = W_POSC5 AND (trim(JDGEIDC5)='" & Trim$(C_JDGEIDC5) & "')"
    End If
    
    'SQL���s
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
        
   '���R�[�h���̂����݂��Ȃ��ꍇ
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GetCOSF3PTN = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    Else
        '������ђl��NULL�̏ꍇ
        If C_strChkR = "R" And C_strChkD = "-" Then
            If Trim(rs("RMAXC5")) = "" Or IsNull(rs("RMAXC5")) Then
                C_RMAXC5 = vbNullString
            Else
                C_RMAXC5 = rs("RMAXC5")
            End If
        ElseIf C_strChkR = "D" Then
            If Trim(rs("DMAXC5")) = "" Or IsNull(rs("DMAXC5")) Then
                C_DMAXC5 = vbNullString
            Else
                C_DMAXC5 = rs("DMAXC5")
            End If
        ElseIf C_strChkR = "R" And C_strChkD = "D" Then
            If Trim(rs("DRRMAXC5")) = "" Or IsNull(rs("DRRMAXC5")) Then
                C_DRRMAXC5 = vbNullString
            Else
                C_DRRMAXC5 = rs("DRRMAXC5")
            End If
            If Trim(rs("DRDMAXC5")) = "" Or IsNull(rs("DRDMAXC5")) Then
                C_DRDMAXC5 = vbNullString
            Else
                C_DRDMAXC5 = rs("DRDMAXC5")
            End If
        End If
    End If
    
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If
    
    '����I��
    GetCOSF3PTN = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function
    
End Function


'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA�]���Ή�(Cu-deco) : Cu-deco(CJ, CJ2)�̕��ʕʋK�i�l�\����(�������xodc5_osf31)�擾�֐�
Public Function GetOsf31_CuDeco(pstrC_JDGEIDC5 As String, plngC_POSC5 As Long, ptyp_Ret As typ_SB_com_xodb5_osf31_Cudeco) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String


    '������
    GetOsf31_CuDeco = FUNCTION_RETURN_FAILURE

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    '�������(XODC5_OSF31)��蔻��f�[�^�̊l�����s��
    '���Ұ�������݋敪�ɂ���ď�������
    sql = " select CJDMAXPIC5, CJRMAXPIC5, CJDRMAXPIC5, CJALLMAXDIC5, CJALLMINRINC5"
    sql = sql & ", CJALLMAXRIGC5, CJ2DMAXPIC5, CJ2RMAXPIC5, CJ2RMINRINC5, CJ2RMAXRIGC5"
    sql = sql & ", CJ2DRMAXPIC5, CJ2DRMINRINC5, CJ2DRMAXRIGC5, JDGEIDC5, POSC5"

    sql = sql & " FROM (select MIN(POSC5) as W_POSC5 from (select POSC5 from XODC5_OSF31 where (trim(JDGEIDC5)='" & Trim$(pstrC_JDGEIDC5) & "') and POSC5 >='" & Trim$(plngC_POSC5) & "')  ), XODC5_OSF31"
    sql = sql & " WHERE POSC5 = W_POSC5 AND (trim(JDGEIDC5)='" & Trim$(pstrC_JDGEIDC5) & "')"
    
    'SQL���s
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
        
   '���R�[�h���̂����݂��Ȃ��ꍇ
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GetOsf31_CuDeco = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    Else
        ptyp_Ret.JDGEIDC5 = rs("JDGEIDC5")                                  ' ����ID
        ptyp_Ret.POSC5 = CInt(fncNullCheck(rs("POSC5")))                    ' ����
        
        '������ђl��NULL�̏ꍇ
        ptyp_Ret.CJDMAXPIC5 = CInt(fncNullCheck(rs("CJDMAXPIC5")))          ' CJ Disk�̂݃p�^�[�� Pi�����
        ptyp_Ret.CJRMAXPIC5 = CInt(fncNullCheck(rs("CJRMAXPIC5")))          ' CJ Ring�̂݃p�^�[�� Pi�����
        ptyp_Ret.CJDRMAXPIC5 = CInt(fncNullCheck(rs("CJDRMAXPIC5")))        ' CJ DiskRing�p�^�[�� Pi�����
        ptyp_Ret.CJALLMAXDIC5 = CInt(fncNullCheck(rs("CJALLMAXDIC5")))      ' CJ ����Disk���a���
        ptyp_Ret.CJALLMINRINC5 = CInt(fncNullCheck(rs("CJALLMINRINC5")))    ' CJ ����Ring���a����
        ptyp_Ret.CJALLMAXRIGC5 = CInt(fncNullCheck(rs("CJALLMAXRIGC5")))    ' CJ ����Ring�O�a���
        
        ptyp_Ret.CJ2DMAXPIC5 = CInt(fncNullCheck(rs("CJ2DMAXPIC5")))        ' CJ2 Disk�̂݃p�^�[�� Pi������(MAX���������ł�)
        ptyp_Ret.CJ2RMAXPIC5 = CInt(fncNullCheck(rs("CJ2RMAXPIC5")))        ' CJ2 Ring�̂݃p�^�[�� Pi������(MAX���������ł�)
        ptyp_Ret.CJ2RMINRINC5 = CInt(fncNullCheck(rs("CJ2RMINRINC5")))      ' CJ2 Ring�̂݃p�^�[�� Ring���a����
        ptyp_Ret.CJ2RMAXRIGC5 = CInt(fncNullCheck(rs("CJ2RMAXRIGC5")))      ' CJ2 Ring�̂݃p�^�[�� Ring�O�a���
        ptyp_Ret.CJ2DRMAXPIC5 = CInt(fncNullCheck(rs("CJ2DRMAXPIC5")))      ' CJ2 DiskRing�p�^�[�� Pi������(MAX���������ł�)
        ptyp_Ret.CJ2DRMINRINC5 = CInt(fncNullCheck(rs("CJ2DRMINRINC5")))    ' CJ2 DiskRing�p�^�[�� Ring���a����
        ptyp_Ret.CJ2DRMAXRIGC5 = CInt(fncNullCheck(rs("CJ2DRMAXRIGC5")))    ' CJ2 DiskRing�p�^�[�� Ring�O�a���
        
    End If
    
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If
    
    '����I��
    GetOsf31_CuDeco = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function
    
End Function
''Add End   2011/01/17 SMPK A.Nagamine



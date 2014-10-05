Attribute VB_Name = "s_cmbc037_SQL"
Option Explicit
'�w���P����


Public Type PURCHASE_CRYSTAL_SPECIFICATION
    Res(4)      As Double   '���R Top1�`5
    RRG         As Double   '���R Top RRG
    Oi(4)       As Double   'Oi Top 1�`5
    ORG         As Double   'Oi Top ORG
    Cs          As Double   'Cs Top
    LD1(1)      As Double   'LD-1 Top Max, LD-1 Top Ave
    LD2(1)      As Double   'LD-2 Top Max, LD-2 Top Ave
    BMD(1)      As Double   'BMD Top Max, LD-2 Top Ave
    GD(3)       As Double   'GD1 Top,GD2 Top, DIA1 Top, DIA2 Top
    Lt          As Double   'LifeTime From Top
    EPD         As Double   'EPD
End Type


Public Type PURCHASE_CRYSTAL
    DELETE      As String       '�������敪
    KRPROCCD    As String       '�Ǘ��H���R�[�h
    PROCCODE    As String       '�H���R�[�h
    TSTAFFID    As String       '�Ј�ID
    HCNO        As String       '����NO
    RBATCHNO    As String       '�F�p�b�`No
    blkID       As String       '�u���b�NID (IN)
    hinban      As String       '�i��
    DMTOP(1)    As Double       '���aTop1�`2
    DMTAIL(1)   As Double       '���aBot1�`2
    NCHDPTH(1)  As Double       '�m�b�`�[��1�`2
    NCHWIDTH(1) As Double       '�m�b�`��1�`2
    NCHPOS      As String * 2   '�m�b�`�ʒu
    SEEDDEG     As String * 1   '�V�[�h�X��
    UPLENGTH    As Double       '���㒷
    SXLPOS      As Double       'SXL�ʒu
    BlkLen      As Double       '�u���b�N����
    BLKWGHT     As Double       '�u���b�N�d��
    Spec(1)     As PURCHASE_CRYSTAL_SPECIFICATION
End Type
' ���i�d�l
Public Type typ_HinSpec1
    HIN As tFullHinban          ' �i��
    HSXTYPE As String * 1       ' �^�C�v
    HSXCDIR As String * 1       ' ����
    HSXD1CEN As Double          ' ���a
    HSXDOP As String * 1        ' �����h�[�v
    HSXDPDIR As String * 2      ' �m�b�`�ʒu
    HSXDDMIN As Double          ' �m�b�`�[���i�l�h�m�j
    HSXDDMAX As Double          ' �m�b�`�[���i�l�`�w�j
    HSXSDSLP As Integer         ' �V�[�h�X��
' ���o�K�����ڒǉ��Ή� yakimura 2002.12.01 start
    TOPREG As Integer           ' TOP�K��
    TAILREG As Double           ' TAIL�K��
    BTMSPRT As Integer          ' �{�g���͏o�K��
' ���o�K�����ڒǉ��Ή� yakimura 2002.12.01 end
End Type
' �ؒf�w��
Public Type typ_CutInd
    INGOTPOS As Integer         ' �J�b�g�ʒu
    TRANCNT As Integer          ' ������
    LENGTH As Integer           ' ����
    PROCCODE As String * 5      ' �H���R�[�h
    BDCAUS As String * 3        ' �敪
    HINUP As tFullHinban        ' ��i��
    HINDN As tFullHinban        ' ���i��
    BLOCKID As String * 12      ' �u���b�NID
    SMP As typ_SXLSample        ' ��������
    PALTNUM As String * 4       ' �p���b�g�ԍ�
    ERRUPFLG As Boolean         ' ��i�ԃG���[�t���O
    ERRDNFLG As Boolean         ' ���i�ԃG���[�t���O
    RECOMMEND(1 To 13) As String * 1    '�����ߌ���(Rs�`EPD)
End Type





'�u���b�NID���͎�
'�T�v    :�w���P���� �\���p�c�a�h���C�o
'���Ұ�  :�ϐ���       ,IO  ,�^                                    ,����
'        :record       ,IO  ,PURCHASE_CRYSTAL                      ,�w���P�����擾�p
'        :��ؒl        ,O   ,FUNCTION_RETURN                       ,�ǂݍ��ݐ���
'����    :
'����    :2001/06/18 ���{ �쐬
Public Function DBDRV_s_cmbc037_Disp(record As PURCHASE_CRYSTAL) As FUNCTION_RETURN
    
    
    Dim sql As String
    Dim rs As OraDynaset
    Dim cdc As OraFields

    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc037_SQL.bas -- Function DBDRV_s_cmbc037_Disp"

    DBDRV_s_cmbc037_Disp = FUNCTION_RETURN_SUCCESS

    sql = "select "
    sql = sql & "CRYNUM, "           ' �����ԍ�
    sql = sql & "TRANCNT, "          ' ������
    sql = sql & "KRPROCCD, "         ' �Ǘ��H���R�[�h
    sql = sql & "PROCCODE, "         ' �H���R�[�h
    sql = sql & "HINBAN, "           ' �i��
    sql = sql & "MNOREVNO, "         ' ���i�ԍ������ԍ�
    sql = sql & "FACTORY, "          ' �H��
    sql = sql & "OPECOND, "          ' ���Ə���
    sql = sql & "REPCCL, "           ' �������敪
    sql = sql & "RBATCHNO, "         ' �F�o�b�`�m��
    sql = sql & "DMTOP1, "           ' ���a�s�n�o�P
    sql = sql & "DMTOP2, "           ' ���a�s�n�o�Q
    sql = sql & "DMTAIL1, "          ' ���a�s�`�h�k�P
    sql = sql & "DMTAIL2, "          ' ���a�s�`�h�k�Q
    sql = sql & "NCHPOS, "           ' �m�b�`�ʒu
    sql = sql & "NCHWID1, "          ' �m�b�`�ЂP
    sql = sql & "NCHWID2, "          ' �m�b�`�ЂQ
    sql = sql & "SEEDDEG, "          ' �V�[�h�X��
    sql = sql & "NCHDPTH1, "         ' �m�b�`�[���P
    sql = sql & "NCHDPTH2, "         ' �m�b�`�[���Q
    sql = sql & "UPLENGTH, "         ' ���グ��
    sql = sql & "SXLPOS, "           ' �r�w�k�ʒu
    sql = sql & "BLKLEN, "           ' �u���b�N����
    sql = sql & "BLKWGHT, "          ' �u���b�N�d��
    sql = sql & "CMPTOP1, "          ' ���RTOP�@�P
    sql = sql & "CMPTOP2, "          ' ���RTOP�@�Q
    sql = sql & "CMPTOP3, "          ' ���RTOP�@�R
    sql = sql & "CMPTOP4, "          ' ���RTOP�@�S
    sql = sql & "CMPTOP5, "          ' ���RTOP�@�T
    sql = sql & "CMPTOPR, "          ' ���RTOP�@RRG
    sql = sql & "CMPTAIL1, "         ' ���RTAIL�@�P
    sql = sql & "CMPTAIL2, "         ' ���RTAIL�@�Q
    sql = sql & "CMPTAIL3, "         ' ���RTAIL�@�R
    sql = sql & "CMPTAIL4, "         ' ���RTAIL�@�S
    sql = sql & "CMPTAIL5, "         ' ���RTAIL�@�T
    sql = sql & "CMPTAILR, "         ' ���RTAIL�@RRG
    sql = sql & "OITOP1, "           ' Oi�@TOP�@�P
    sql = sql & "OITOP2, "           ' Oi�@TOP�@�Q
    sql = sql & "OITOP3, "           ' Oi�@TOP�@�R
    sql = sql & "OITOP4, "           ' Oi�@TOP�@�S
    sql = sql & "OITOP5, "           ' Oi�@TOP�@�T
    sql = sql & "OITOPR, "           ' Oi�@TOP�@ROG
    sql = sql & "OITAIL1, "          ' Oi�@TAIL�@�P
    sql = sql & "OITAIL2, "          ' Oi�@TAIL�@�Q
    sql = sql & "OITAIL3, "          ' Oi�@TAIL�@�R
    sql = sql & "OITAIL4, "          ' Oi�@TAIL�@�S
    sql = sql & "OITAIL5, "          ' Oi�@TAIL�@�T
    sql = sql & "OITAILR, "          ' Oi�@TAIL�@ROG
    sql = sql & "CSTOP, "            ' Cs�@TOP
    sql = sql & "CSTAIL, "           ' Cs�@TAIL
    sql = sql & "LD1TOPMX, "         ' LD-1�@TOP�@MAX
    sql = sql & "LD1TOPAV, "         ' LD-1�@TOP�@AVE
    sql = sql & "LD1TAILM, "         ' LD-1�@TAIL�@MAX
    sql = sql & "LD1TAILA, "         ' LD-1�@TAIL�@AVE
    sql = sql & "LD2TOPMM, "         ' LD-2�@TOP�@MAX
    sql = sql & "LD2TOPAV, "         ' LD-2�@TOP�@AVE
    sql = sql & "LD2TAILM, "         ' LD-2�@TAIL�@MAX
    sql = sql & "LD2TAILA, "         ' LD-2�@TAIL�@AVE
    sql = sql & "BMDTOPMX, "         ' BMD�@TOP�@MAX
    sql = sql & "BMDTOPAV, "         ' BMD�@TOP�@AVE
    sql = sql & "BMDTAILM, "         ' BMD�@TAIL�@MAX
    sql = sql & "BMDTAILA, "         ' BMD�@TAIL�@AVE
    sql = sql & "GD1TOP, "           ' GD1 TOP
    sql = sql & "GD1TAIL, "          ' GD1 TAIL
    sql = sql & "GD2TOP, "           ' GD2 TOP
    sql = sql & "GD2TAIL, "          ' GD2 TAIL
    sql = sql & "DIA1TOP, "          ' DIA1 TOP
    sql = sql & "DIA1TAIL, "         ' DIA1 TAIL
    sql = sql & "DIA2TOP, "          ' DIA2 TOP
    sql = sql & "DIA2TAIL, "         ' DIA2 TAIL
    sql = sql & "LTFTOP, "           ' LIFETIME from TOP
    sql = sql & "LTFTAIL, "          ' LIFETIME from TAIL
    sql = sql & "EPD, "              ' EPD
    sql = sql & "HCNO, "             ' ����No
    sql = sql & "TSTAFFID, "         ' �o�^�Ј�ID
    sql = sql & "REGDATE, "          ' �o�^���t
    sql = sql & "KSTAFFID, "         ' �X�V�Ј�ID
    sql = sql & "UPDDATE, "          ' �X�V���t
    sql = sql & "SENDFLAG, "         ' ���M�t���O
    sql = sql & "SENDDATE "          ' ���M���t
    sql = sql & " From TBCMG002 "
    sql = sql & " where TRANCNT=ANY(select MAX(TRANCNT) from TBCMG002 Where CRYNUM='" & record.blkID & "') and  CRYNUM='" & record.blkID & "'"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '���R�[�h0�����̓G���[
    If rs.RecordCount = 0 Then
        DBDRV_s_cmbc037_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    Set cdc = rs.Fields
    With record
        .DELETE = cdc("REPCCL").Value                   '�������敪
        .RBATCHNO = cdc("RBATCHNO").Value               ' �F�o�b�`�m��
        .hinban = cdc("HINBAN").Value & String(2 - Len(Trim(cdc("MNOREVNO").Value)), "0") & Trim(cdc("MNOREVNO").Value) ' �i��
        .DMTOP(0) = cdc("DMTOP1").Value                 ' ���a�s�n�o�P
        .DMTOP(1) = cdc("DMTOP2").Value                 ' ���a�s�n�o�Q
        .DMTAIL(0) = cdc("DMTAIL1").Value               ' ���a�s�`�h�k�P
        .DMTAIL(1) = cdc("DMTAIL2").Value               ' ���a�s�`�h�k�Q
        .NCHPOS = cdc("NCHPOS").Value                   ' �m�b�`�ʒu
        .NCHWIDTH(0) = cdc("NCHWID1").Value             ' �m�b�`�ЂP
        .NCHWIDTH(1) = cdc("NCHWID2").Value             ' �m�b�`�ЂQ
        .SEEDDEG = cdc("SEEDDEG").Value                 ' �V�[�h�X��
        .NCHDPTH(0) = cdc("NCHDPTH1").Value             ' �m�b�`�[���P
        .NCHDPTH(1) = cdc("NCHDPTH2").Value             ' �m�b�`�[���Q
        .UPLENGTH = cdc("UPLENGTH").Value               ' ���グ��
        .SXLPOS = cdc("SXLPOS").Value                   ' �r�w�k�ʒu
        .BlkLen = cdc("BLKLEN").Value                   ' �u���b�N����
        .BLKWGHT = cdc("BLKWGHT").Value                 ' �u���b�N�d��
        .Spec(0).Res(0) = cdc("CMPTOP1").Value          ' ���RTOP�@�P
        .Spec(0).Res(1) = cdc("CMPTOP2").Value          ' ���RTOP�@�Q
        .Spec(0).Res(2) = cdc("CMPTOP3").Value          ' ���RTOP�@�R
        .Spec(0).Res(3) = cdc("CMPTOP4").Value          ' ���RTOP�@�S
        .Spec(0).Res(4) = cdc("CMPTOP5").Value          ' ���RTOP�@�T
        .Spec(1).Res(0) = cdc("CMPTAIL1").Value         ' ���RTAIL�@�P
        .Spec(1).Res(1) = cdc("CMPTAIL2").Value         ' ���RTAIL�@�Q
        .Spec(1).Res(2) = cdc("CMPTAIL3").Value         ' ���RTAIL�@�R
        .Spec(1).Res(3) = cdc("CMPTAIL4").Value         ' ���RTAIL�@�S
        .Spec(1).Res(4) = cdc("CMPTAIL5").Value         ' ���RTAIL�@�T
        .Spec(0).RRG = cdc("CMPTOPR").Value             ' ���RTOP�@RRG
        .Spec(1).RRG = cdc("CMPTAILR").Value            ' ���RTAIL�@RRG
        .Spec(0).Oi(0) = cdc("OITOP1").Value            ' Oi�@TOP�@�P
        .Spec(0).Oi(1) = cdc("OITOP2").Value            ' Oi�@TOP�@�Q
        .Spec(0).Oi(2) = cdc("OITOP3").Value            ' Oi�@TOP�@�R
        .Spec(0).Oi(3) = cdc("OITOP4").Value            ' Oi�@TOP�@�S
        .Spec(0).Oi(4) = cdc("OITOP5").Value            ' Oi�@TOP�@�T
        .Spec(1).Oi(0) = cdc("OITAIL1").Value           ' Oi�@TAIL�@�P
        .Spec(1).Oi(1) = cdc("OITAIL2").Value           ' Oi�@TAIL�@�Q
        .Spec(1).Oi(2) = cdc("OITAIL3").Value           ' Oi�@TAIL�@�R
        .Spec(1).Oi(3) = cdc("OITAIL4").Value           ' Oi�@TAIL�@�S
        .Spec(1).Oi(4) = cdc("OITAIL5").Value           ' Oi�@TAIL�@�T
        .Spec(0).ORG = cdc("OITOPR").Value              ' Oi�@TOP�@ROG
        .Spec(1).ORG = cdc("OITAILR").Value             ' Oi�@TAIL�@ROG
        .Spec(0).Cs = cdc("CSTOP").Value                ' Cs�@TOP
        .Spec(1).Cs = cdc("CSTAIL").Value               ' Cs�@TAIL
        .Spec(0).LD1(0) = cdc("LD1TOPMX").Value         ' LD-1�@TOP�@MAX
        .Spec(0).LD1(1) = cdc("LD1TOPAV").Value         ' LD-1�@TOP�@AVE
        .Spec(1).LD1(0) = cdc("LD1TAILM").Value         ' LD-1�@TAIL�@MAX
        .Spec(1).LD1(1) = cdc("LD1TAILA").Value         ' LD-1�@TAIL�@AVE
        .Spec(0).LD2(0) = cdc("LD2TOPMM").Value         ' LD-2�@TOP�@MAX
        .Spec(0).LD2(1) = cdc("LD2TOPAV").Value         ' LD-2�@TOP�@AVE
        .Spec(1).LD2(0) = cdc("LD2TAILM").Value         ' LD-2�@TAIL�@MAX
        .Spec(1).LD2(1) = cdc("LD2TAILA").Value         ' LD-2�@TAIL�@AVE
        .Spec(0).BMD(0) = cdc("BMDTOPMX").Value         ' BMD�@TOP�@MAX
        .Spec(0).BMD(1) = cdc("BMDTOPAV").Value         ' BMD�@TOP�@AVE
        .Spec(1).BMD(0) = cdc("BMDTAILM").Value         ' BMD�@TAIL�@MAX
        .Spec(1).BMD(1) = cdc("BMDTAILA").Value         ' BMD�@TAIL�@AVE
        .Spec(0).GD(0) = cdc("GD1TOP").Value            ' GD1 TOP
        .Spec(0).GD(1) = cdc("GD2TOP").Value            ' GD1 TAIL
        .Spec(0).GD(2) = cdc("DIA1TOP").Value           ' GD2 TOP
        .Spec(0).GD(3) = cdc("DIA2TOP").Value           ' GD2 TAIL
        .Spec(1).GD(0) = cdc("GD1TAIL").Value           ' DIA1 TOP
        .Spec(1).GD(1) = cdc("GD2TAIL").Value           ' DIA1 TAIL
        .Spec(1).GD(2) = cdc("DIA1TAIL").Value          ' DIA2 TOP
        .Spec(1).GD(3) = cdc("DIA2TAIL").Value          ' DIA2 TAIL
        .Spec(0).Lt = cdc("LTFTOP").Value               ' LIFETIME from TOP
        .Spec(1).Lt = cdc("LTFTAIL").Value              ' LIFETIME from TAIL
        .Spec(1).EPD = cdc("EPD").Value                 ' EPD
        .HCNO = cdc("HCNO").Value                       ' ����No
    End With
    rs.Close
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_s_cmbc037_Disp = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'���s��
'�T�v    :�w���P���� �X�V�A�}���p�c�a�h���C�o
'���Ұ�  :�ϐ���       ,IO  ,�^                                ,����
'        :record       ,I   ,PURCHASE_CRYSTAL                 ,�w���P�����}���p
'        :sCmd         ,I   ,String                           ,�֐��ďo�R�}���h�@2003/10/31 ooba
'        :UpDateFlag   ,I   ,Boolean                          ,�X�V�}���t���O
'        :��ؒl        ,O   ,FUNCTION_RETURN                   ,�ǂݍ��ݐ���
'����    :
'����    :2001/06/18 ���{ �쐬
'�@�@    :2001/07/19 Sano ����
'        :������/�X�V�o�^�̏ꍇ[delete]����[insert]����悤�ɕύX�@2003/10/31 ooba

Public Function DBDRV_scmzc_fcmec001b_Exec(record As PURCHASE_CRYSTAL, _
                                            sCmd As String, _
                                            UpDateFlag As Boolean, _
                                            pCryOld() As typ_XSDCS, _
                                            pCrySmp() As typ_XSDCS _
                                            ) As FUNCTION_RETURN

    Dim sql As String
    Dim fullHinban As tFullHinban
    Dim CryInf As typ_TBCME037
    Dim BlockMng As typ_TBCME040
    Dim hinban As typ_TBCME041
    Dim recCnt As Integer
    Dim i As Long
    Dim sDbName As String
    Dim sDelSql As String    ''delete�pSQL�@2003/10/31 ooba
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc037_SQL.bas -- Function DBDRV_scmzc_fcmec001b_Exec"

    DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_SUCCESS

    '12���i�Ԃ����߂�
    If GetLastHinban(record.hinban, fullHinban) = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    '�w���P�����e�[�u���̑}���A�X�V TBCMG002
    
    If DBDRV_KCryTbl_Exec(record, UpDateFlag) = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    'TBCMI002�@������H����
    If UpDateFlag Then
        sDelSql = "delete from TBCMI002 "
        sDelSql = sDelSql & "where CRYNUM = '" & Left(record.blkID, 9) & "000' "
        
        If DBDRV_DeleteTable(sDelSql, "TBCMI002") = FUNCTION_RETURN_FAILURE Then
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    ''[execution]:���s�A[cancel]:������
    If sCmd = "execution" Then
        If DBDRV_TBCMI002_Exec(record, UpDateFlag) = FUNCTION_RETURN_FAILURE Then
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    
    
    '�������ւ̑}��
    With CryInf
        .CRYNUM = Left(record.blkID, 9) & "000"             ' �����ԍ�
        .DELCLS = record.DELETE                             ' �폜�敪
        .LPKRPROCCD = MGPRCD_KOUNYU_TAN_KESSYOU             ' �ŏI�ʉߊǗ��H��
        .LASTPASS = PROCD_KOUNYU_TAN_KESSYOU                ' �ŏI�ʉߍH��
        .KRPROCCD = MGPRCD_KESSYOU_SOUGOUHANTEI             ' �Ǘ��H���R�[�h
        .PROCCD = PROCD_KESSYOU_SOUGOUHANTEI                ' �H���R�[�h
        .RPHINBAN = fullHinban.hinban                       ' �˂炢�i��
        .RPREVNUM = fullHinban.mnorevno                     ' �˂炢�i�Ԑ��i�ԍ������ԍ�
        .RPFACT = fullHinban.factory                        ' �˂炢�i�ԍH��
        .RPOPCOND = fullHinban.opecond                      ' �˂炢�i�ԑ��Ə���
        .PRODCOND = ""                                      ' �������
        .PGID = ""                                          ' �o�f�|�h�c
        .UPLENGTH = record.UPLENGTH                         ' ���グ����
        .TOPLENG = 0                                        ' �s�n�o����
        .BODYLENG = record.UPLENGTH                         ' ��������
        .BOTLENG = 0                                        ' �a�n�s����
        .FREELENG = record.UPLENGTH                         ' �t���[��
        .DIAMETER = (record.DMTOP(0) + record.DMTOP(1)) / 2 ' ���a
        .CHARGE = 0                                         ' �`���[�W��
        .SEED = ""                                          ' �V�[�h
        .ADDDPCLS = ""                                      ' �ǉ��h�[�v���
        .ADDDPPOS = 0                                       ' �ǉ��h�[�v�ʒu
        .ADDDPVAL = 0                                       ' �ǉ��h�[�v��
'        .REGDATE                                            ' �o�^���t
'        .UPDDATE                                            ' �X�V���t
'        .SENDFLAG                                           ' ���M�t���O
'        .SENDDATE                                           ' ���M���t
    End With
    If UpDateFlag Then
'        If DBDRV_CryInf_Upd(CryInf) = FUNCTION_RETURN_FAILURE Then
'            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
'            GoTo proc_exit
'        End If
'    Else
        sDelSql = "delete from TBCME037 "
        sDelSql = sDelSql & "where CRYNUM = '" & CryInf.CRYNUM & "' "
        
        If DBDRV_DeleteTable(sDelSql, "TBCME037") = FUNCTION_RETURN_FAILURE Then
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    ''[execution]:���s�A[cancel]:������
    If sCmd = "execution" Then
        If DBDRV_CryInf_Ins(CryInf) = FUNCTION_RETURN_FAILURE Then
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    
    
    '�u���b�N�Ǘ��ւ̑}��
    With BlockMng
        .CRYNUM = Left(record.blkID, 9) & "000"             ' �����ԍ�
        .INGOTPOS = record.SXLPOS                           ' �������J�n�ʒu
        .REALLEN = record.BlkLen                            ' ������
        .LENGTH = record.BlkLen                             ' ����
        .BLOCKID = record.blkID                             ' �u���b�NID
        .KRPROCCD = MGPRCD_KESSYOU_SOUGOUHANTEI             ' ���݊Ǘ��H��
        .NOWPROC = PROCD_KESSYOU_SOUGOUHANTEI               ' ���ݍH��
        .LPKRPROCCD = MGPRCD_KOUNYU_TAN_KESSYOU             ' �ŏI�ʉߊǗ��H��
        .LASTPASS = PROCD_KOUNYU_TAN_KESSYOU                ' �ŏI�ʉߍH��
        .DELCLS = record.DELETE                             ' �폜�敪
        .LSTATCLS = "T"                                     ' �ŏI��ԋ敪
        .RSTATCLS = "T"                                     ' ������ԋ敪
        .HOLDCLS = "0"                                      ' �z�[���h�敪
        .BDCAUS = ""                                        ' �s�Ǘ��R
'        .REGDATE                                            ' �o�^���t
'        .UPDDATE                                            ' �X�V���t
        .SUMMITSENDFLAG = ""                                ' SUMMIT���M�t���O
'        .SENDFLAG                                           ' ���M�t���O
'        .SENDDATE                                           ' ���M���t
    End With
    If UpDateFlag Then
'        If DBDRV_BlockMng_Upd_SS(BlockMng) = FUNCTION_RETURN_FAILURE Then
'            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
'            GoTo proc_exit
'        End If
'    Else
        sDelSql = "delete from TBCME040 "
        sDelSql = sDelSql & "where CRYNUM = '" & BlockMng.CRYNUM & "' "
'        sDelSql = sDelSql & "and INGOTPOS = " & BlockMng.INGOTPOS
        
        If DBDRV_DeleteTable(sDelSql, "TBCME040") = FUNCTION_RETURN_FAILURE Then
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    ''[execution]:���s�A[cancel]:������
    If sCmd = "execution" Then
        If DBDRV_BlockMng_Ins(BlockMng) = FUNCTION_RETURN_FAILURE Then
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    With hinban
        .CRYNUM = BlockMng.CRYNUM       ' �����ԍ�
        .INGOTPOS = record.SXLPOS       ' �������J�n�ʒu
        .hinban = fullHinban.hinban     ' �i��
        .REVNUM = fullHinban.mnorevno   ' ���i�ԍ������ԍ�
        .factory = fullHinban.factory   ' �H��
        .opecond = fullHinban.opecond   ' ���Ə���
        .LENGTH = record.BlkLen         ' ����
    End With
    If UpDateFlag Then
'        If record.DELETE = "1" Then
''            sql = "delete from TBCME041 where CRYNUM = '" & HINBAN.CryNum & "' and INGOTPOS = " & HINBAN.IngotPos
''            If 0 >= OraDB.ExecuteSQL(sql) Then
''                DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
''                GoTo proc_exit
''            End If
'        Else
'            With hinban
'            sql = "update TBCME041 set "
'            sql = sql & "INGOTPOS=" & .INGOTPOS & ", "
'            sql = sql & "HINBAN='" & .hinban & "', "           ' �i��
'            sql = sql & "REVNUM='" & .REVNUM & "', "           ' ���i�ԍ������ԍ�
'            sql = sql & "FACTORY='" & .factory & "', "         ' �H��
'            sql = sql & "OPECOND='" & .opecond & "', "         ' ���Ə���
'            sql = sql & "LENGTH='" & .LENGTH & "', "           ' ����
'            sql = sql & " UPDDATE=sysdate, "
'            sql = sql & " SENDFLAG='0' "
'            sql = sql & " where CRYNUM='" & .CRYNUM & "' "
'            sql = sql & " and INGOTPOS=" & key.POSITION & " "
'            End With
'            If 0 >= OraDB.ExecuteSQL(sql) Then
'                DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
'                GoTo proc_exit
'            End If
'        End If
'    Else
        sDelSql = "delete from TBCME041 "
        sDelSql = sDelSql & "where CRYNUM = '" & hinban.CRYNUM & "' "
'        sDelSql = sDelSql & "and INGOTPOS = " & hinban.INGOTPOS
        
        If DBDRV_DeleteTable(sDelSql, "TBCME041") = FUNCTION_RETURN_FAILURE Then
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    
    ''[execution]:���s�A[cancel]:������
    If sCmd = "execution" Then
        '�i�ԊǗ��̑}��
        sql = "insert into TBCME041 ( "
        sql = sql & "CRYNUM, "            ' �����ԍ�
        sql = sql & "INGOTPOS, "          ' �������J�n�ʒu
        sql = sql & "HINBAN, "            ' �i��
        sql = sql & "REVNUM, "            ' ���i�ԍ������ԍ�
        sql = sql & "FACTORY, "           ' �H��
        sql = sql & "OPECOND, "           ' ���Ə���
        sql = sql & "LENGTH, "            ' ����
        sql = sql & "REGDATE, "           ' �o�^���t
        sql = sql & "UPDDATE, "           ' �X�V���t
        sql = sql & "SENDFLAG, "          ' ���M�t���O
        sql = sql & "SENDDATE  ) "          ' ���M���t
        With hinban
        sql = sql & "values ("
        sql = sql & " '" & .CRYNUM & "', "          ' �����ԍ�
        sql = sql & " " & .INGOTPOS & ", "          ' �������J�n�ʒu
        sql = sql & " '" & .hinban & "', "          ' �i��
        sql = sql & " " & .REVNUM & ", "            ' ���i�ԍ������ԍ�
        sql = sql & " '" & .factory & "', "         ' �H��
        sql = sql & " '" & .opecond & "', "         ' ���Ə���
        sql = sql & " " & .LENGTH & ", "            ' ����
        End With
        sql = sql & " sysdate, "                    ' �o�^���t
        sql = sql & " sysdate, "                    ' �X�V���t
        sql = sql & " '0', "                        ' ���M�t���O
        sql = sql & " sysdate ) "                   ' ���M���t
        If 0 >= OraDB.ExecuteSQL(sql) Then
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    
    'XSDC1 �ǉ�/�X�V
            
    'XSDC2 �ǉ�/�X�V
    
    'XSDC3 �ǉ�/�X�V
    
    'XSDCA �ǉ�/�X�V
    
    
    sDbName = "E043"
'    If record.DELETE = "1" Then
    If UpDateFlag Then
        '' �����T���v���Ǘ��̍폜
        If DBDRV_CrySmp_Del(pCryOld()) = FUNCTION_RETURN_FAILURE Then
            'sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    ''[execution]:���s�A[cancel]:������
    If sCmd = "execution" Then
        '' �T���v�����̎擾
        recCnt = UBound(pCrySmp)
        For i = 1 To recCnt
            If pCrySmp(i).REPSMPLIDCS = 0 Then
                pCrySmp(i).REPSMPLIDCS = GetNewID_SampleNo()
            End If
        Next i

        '' �����T���v���Ǘ��̑}���^�X�V
        If DBDRV_CrySmp_UpdIns037Only(pCryOld(), pCrySmp()) = FUNCTION_RETURN_FAILURE Then
            'sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

Public Function DBDRV_GetCryCheck(CRYNUM As String, CryFlag As Boolean) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc037_SQL.bas -- Function DBDRV_GetCryCheck"

    DBDRV_GetCryCheck = FUNCTION_RETURN_FAILURE
    
    sql = "select CRYNUM, DELCLS from TBCME037 where CRYNUM='" & CRYNUM & "'"
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    CryFlag = True
    If rs.RecordCount = 0 Then
        CryFlag = False
    Else
        If rs("DELCLS") = "1" Then
            CryFlag = False
        End If
    End If
    rs.Close
    
    DBDRV_GetCryCheck = FUNCTION_RETURN_SUCCESS
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_GetCryCheck = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

Public Function DBDRV_GetCryInBlk(CRYNUM As String, blkPos As Integer, BlkLen As Integer, OkNgFlag As Boolean) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc037_SQL.bas -- Function DBDRV_GetCryInBlk"

    DBDRV_GetCryInBlk = FUNCTION_RETURN_FAILURE
    
   '�w��͈͂ɓ����Ă���u���b�N �̌���

    sql = "select CRYNUM from TBCME040 where CRYNUM='" & CRYNUM & "' and ("
'    sql = "select count(CRYNUM) from TBCME040 where CRYNUM='" & CryNum & "' and ("
    sql = sql & "(INGOTPOS > " & blkPos & " and INGOTPOS < " & blkPos + BlkLen & ") "
    sql = sql & "or (INGOTPOS + LENGTH > " & blkPos & " and INGOTPOS + LENGTH <  " & blkPos + BlkLen & ") "
    sql = sql & "or (" & blkPos & " > INGOTPOS and " & blkPos + BlkLen & " < INGOTPOS + LENGTH ) "
'    sql = sql & "and (" & BlkPos + BlkLen & " > INGOTPOS and " & BlkPos & " < INGOTPOS + LENGTH)"
    sql = sql & ")"
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    OkNgFlag = False
    If rs.RecordCount = 0 Then
        rs.Close
        OkNgFlag = True
    End If
    rs.Close
    
    DBDRV_GetCryInBlk = FUNCTION_RETURN_SUCCESS
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_GetCryInBlk = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


Public Function DBDRV_KCryTbl_Exec(record As PURCHASE_CRYSTAL, UpDateFlag As Boolean) As FUNCTION_RETURN

    Dim sql As String
    Dim fullHinban As tFullHinban
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc037_SQL.bas -- Function DBDRV_KCryTbl_Exec"

    DBDRV_KCryTbl_Exec = FUNCTION_RETURN_SUCCESS

    '12���i�Ԃ����߂�
    If GetLastHinban(record.hinban, fullHinban) = FUNCTION_RETURN_FAILURE Then
        DBDRV_KCryTbl_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    'If Not UpDateFlag Then
        
        '�w���P�����e�[�u���֒l�̑}��
        sql = "insert into TBCMG002 ( "
        sql = sql & "CRYNUM, "           ' �����ԍ�
        sql = sql & "TRANCNT, "          ' ������
        sql = sql & "KRPROCCD, "         ' �Ǘ��H���R�[�h
        sql = sql & "PROCCODE, "         ' �H���R�[�h
        sql = sql & "HINBAN, "           ' �i��
        sql = sql & "MNOREVNO, "         ' ���i�ԍ������ԍ�
        sql = sql & "FACTORY, "          ' �H��
        sql = sql & "OPECOND, "          ' ���Ə���
        sql = sql & "REPCCL, "           ' �������敪
        sql = sql & "RBATCHNO, "         ' �F�o�b�`�m��
        sql = sql & "DMTOP1, "           ' ���a�s�n�o�P
        sql = sql & "DMTOP2, "           ' ���a�s�n�o�Q
        sql = sql & "DMTAIL1, "          ' ���a�s�`�h�k�P
        sql = sql & "DMTAIL2, "          ' ���a�s�`�h�k�Q
        sql = sql & "NCHPOS, "           ' �m�b�`�ʒu
        sql = sql & "NCHWID1, "          ' �m�b�`�ЂP
        sql = sql & "NCHWID2, "          ' �m�b�`�ЂQ
        sql = sql & "SEEDDEG, "          ' �V�[�h�X��
        sql = sql & "NCHDPTH1, "         ' �m�b�`�[���P
        sql = sql & "NCHDPTH2, "         ' �m�b�`�[���Q
        sql = sql & "UPLENGTH, "         ' ���グ��
        sql = sql & "SXLPOS, "           ' �r�w�k�ʒu
        sql = sql & "BLKLEN, "           ' �u���b�N����
        sql = sql & "BLKWGHT, "          ' �u���b�N�d��
        sql = sql & "CMPTOP1, "          ' ���RTOP�@�P
        sql = sql & "CMPTOP2, "          ' ���RTOP�@�Q
        sql = sql & "CMPTOP3, "          ' ���RTOP�@�R
        sql = sql & "CMPTOP4, "          ' ���RTOP�@�S
        sql = sql & "CMPTOP5, "          ' ���RTOP�@�T
        sql = sql & "CMPTOPR, "          ' ���RTOP�@RRG
        sql = sql & "CMPTAIL1, "         ' ���RTAIL�@�P
        sql = sql & "CMPTAIL2, "         ' ���RTAIL�@�Q
        sql = sql & "CMPTAIL3, "         ' ���RTAIL�@�R
        sql = sql & "CMPTAIL4, "         ' ���RTAIL�@�S
        sql = sql & "CMPTAIL5, "         ' ���RTAIL�@�T
        sql = sql & "CMPTAILR, "         ' ���RTAIL�@RRG
        sql = sql & "OITOP1, "           ' Oi�@TOP�@�P
        sql = sql & "OITOP2, "           ' Oi�@TOP�@�Q
        sql = sql & "OITOP3, "           ' Oi�@TOP�@�R
        sql = sql & "OITOP4, "           ' Oi�@TOP�@�S
        sql = sql & "OITOP5, "           ' Oi�@TOP�@�T
        sql = sql & "OITOPR, "           ' Oi�@TOP�@ROG
        sql = sql & "OITAIL1, "          ' Oi�@TAIL�@�P
        sql = sql & "OITAIL2, "          ' Oi�@TAIL�@�Q
        sql = sql & "OITAIL3, "          ' Oi�@TAIL�@�R
        sql = sql & "OITAIL4, "          ' Oi�@TAIL�@�S
        sql = sql & "OITAIL5, "          ' Oi�@TAIL�@�T
        sql = sql & "OITAILR, "          ' Oi�@TAIL�@ROG
        sql = sql & "CSTOP, "            ' Cs�@TOP
        sql = sql & "CSTAIL, "           ' Cs�@TAIL
        sql = sql & "LD1TOPMX, "         ' LD-1�@TOP�@MAX
        sql = sql & "LD1TOPAV, "         ' LD-1�@TOP�@AVE
        sql = sql & "LD1TAILM, "         ' LD-1�@TAIL�@MAX
        sql = sql & "LD1TAILA, "         ' LD-1�@TAIL�@AVE
        sql = sql & "LD2TOPMM, "         ' LD-2�@TOP�@MAX
        sql = sql & "LD2TOPAV, "         ' LD-2�@TOP�@AVE
        sql = sql & "LD2TAILM, "         ' LD-2�@TAIL�@MAX
        sql = sql & "LD2TAILA, "         ' LD-2�@TAIL�@AVE
        sql = sql & "BMDTOPMX, "         ' BMD�@TOP�@MAX
        sql = sql & "BMDTOPAV, "         ' BMD�@TOP�@AVE
        sql = sql & "BMDTAILM, "         ' BMD�@TAIL�@MAX
        sql = sql & "BMDTAILA, "         ' BMD�@TAIL�@AVE
        sql = sql & "GD1TOP, "           ' GD1 TOP
        sql = sql & "GD1TAIL, "          ' GD1 TAIL
        sql = sql & "GD2TOP, "           ' GD2 TOP
        sql = sql & "GD2TAIL, "          ' GD2 TAIL
        sql = sql & "DIA1TOP, "          ' DIA1 TOP
        sql = sql & "DIA1TAIL, "         ' DIA1 TAIL
        sql = sql & "DIA2TOP, "          ' DIA2 TOP
        sql = sql & "DIA2TAIL, "         ' DIA2 TAIL
        sql = sql & "LTFTOP, "           ' LIFETIME from TOP
        sql = sql & "LTFTAIL, "          ' LIFETIME from TAIL
        sql = sql & "EPD, "              ' EPD
        sql = sql & "HCNO, "             ' ����No
        sql = sql & "TSTAFFID, "         ' �o�^�Ј�ID
        sql = sql & "REGDATE, "          ' �o�^���t
        sql = sql & "KSTAFFID, "         ' �X�V�Ј�ID
        sql = sql & "UPDDATE, "          ' �X�V���t
        sql = sql & "SENDFLAG, "         ' ���M�t���O
        sql = sql & "SENDDATE ) "         ' ���M���t
        With record
            sql = sql & " select "
            sql = sql & " '" & .blkID & "', "             ' �����ԍ�
            sql = sql & "nvl(max(TRANCNT),0)+1, "                   ' ������
            sql = sql & " '" & .KRPROCCD & "', "              ' �Ǘ��H���R�[�h
            sql = sql & " '" & .PROCCODE & "', "              ' �H���R�[�h
            sql = sql & " '" & fullHinban.hinban & "', "            ' �i��
            sql = sql & fullHinban.mnorevno & ", "
            sql = sql & " '" & fullHinban.factory & "', "
            sql = sql & " '" & fullHinban.opecond & "', "
            sql = sql & " '" & .DELETE & "', "                '�������敪
            sql = sql & " '" & .RBATCHNO & "', "
            sql = sql & .DMTOP(0) & ", "
            sql = sql & .DMTOP(1) & ", "
            sql = sql & .DMTAIL(0) & ", "
            sql = sql & .DMTAIL(1) & ", "
            sql = sql & "'" & .NCHPOS & "', "
            sql = sql & .NCHWIDTH(0) & ", "
            sql = sql & "-1, "
            sql = sql & "'" & .SEEDDEG & "', "
            sql = sql & .NCHDPTH(0) & ", "
            sql = sql & "-1, "
            sql = sql & .UPLENGTH & ", "
            sql = sql & .SXLPOS & ", "
            sql = sql & .BlkLen & ", "
            sql = sql & .BLKWGHT & ", "
            sql = sql & .Spec(0).Res(0) & ", "
            sql = sql & .Spec(0).Res(1) & ", "
            sql = sql & .Spec(0).Res(2) & ", "
            sql = sql & .Spec(0).Res(3) & ", "
            sql = sql & .Spec(0).Res(4) & ", "
            sql = sql & .Spec(0).RRG & ", "
            sql = sql & .Spec(1).Res(0) & ", "
            sql = sql & .Spec(1).Res(1) & ", "
            sql = sql & .Spec(1).Res(2) & ", "
            sql = sql & .Spec(1).Res(3) & ", "
            sql = sql & .Spec(1).Res(4) & ", "
            sql = sql & .Spec(1).RRG & ", "
            sql = sql & .Spec(0).Oi(0) & ", "
            sql = sql & .Spec(0).Oi(1) & ", "
            sql = sql & .Spec(0).Oi(2) & ", "
            sql = sql & .Spec(0).Oi(3) & ", "
            sql = sql & .Spec(0).Oi(4) & ", "
            sql = sql & .Spec(0).ORG & ", "
            sql = sql & .Spec(1).Oi(0) & ", "
            sql = sql & .Spec(1).Oi(1) & ", "
            sql = sql & .Spec(1).Oi(2) & ", "
            sql = sql & .Spec(1).Oi(3) & ", "
            sql = sql & .Spec(1).Oi(4) & ", "
            sql = sql & .Spec(1).ORG & ", "
            sql = sql & .Spec(0).Cs & ", "
            sql = sql & .Spec(1).Cs & ", "
            sql = sql & .Spec(0).LD1(0) & ", "
            sql = sql & .Spec(0).LD1(1) & ", "
            sql = sql & .Spec(1).LD1(0) & ", "
            sql = sql & .Spec(1).LD1(1) & ", "
            sql = sql & .Spec(0).LD2(0) & ", "
            sql = sql & .Spec(0).LD2(1) & ", "
            sql = sql & .Spec(1).LD2(0) & ", "
            sql = sql & .Spec(1).LD2(1) & ", "
            sql = sql & .Spec(0).BMD(0) & ", "
            sql = sql & .Spec(0).BMD(1) & ", "
            sql = sql & .Spec(1).BMD(0) & ", "
            sql = sql & .Spec(1).BMD(1) & ", "
            sql = sql & .Spec(0).GD(0) & ", "
            sql = sql & .Spec(1).GD(0) & ", "
            sql = sql & .Spec(0).GD(1) & ", "
            sql = sql & .Spec(1).GD(1) & ", "
            sql = sql & .Spec(0).GD(2) & ", "
            sql = sql & .Spec(1).GD(2) & ", "
            sql = sql & .Spec(0).GD(3) & ", "
            sql = sql & .Spec(1).GD(3) & ", "
            sql = sql & .Spec(0).Lt & ", "
            sql = sql & .Spec(1).Lt & ", "
            sql = sql & .Spec(1).EPD & ", "
            sql = sql & " '" & .HCNO & "', "
            sql = sql & " '" & .TSTAFFID & "', "
            sql = sql & " sysdate, "
            sql = sql & " '" & .TSTAFFID & "', "
            sql = sql & " sysdate , "
            sql = sql & " '0', "
            sql = sql & " sysdate  "
            sql = sql & " from TBCMG002 "
            sql = sql & " where CRYNUM='" & .blkID & "' "
        End With
    'Else
        '�X�V��
        
    '    With record
    '        sql = "UPDATE TBCMG002 SET "
    '        sql = sql & "KRPROCCD='" & .KRPROCCD & "',"
    '        sql = sql & "PROCCODE='" & .PROCCODE & "',"
    '        sql = sql & "HINBAN='" & fullHinban.HINBAN & "',"
    '        sql = sql & "MNOREVNO=" & fullHinban.mnorevno & ","
    '        sql = sql & "FACTORY='" & fullHinban.factory & "',"
    '        sql = sql & "OPECOND='" & fullHinban.opecond & "',"
    '        sql = sql & "REPCCL='" & .DELETE & "',"
    '        sql = sql & "RBATCHNO='" & .RBATCHNO & "',"
    '        sql = sql & "DMTOP1=" & .DMTOP(0) & ","
    '        sql = sql & "DMTOP2=" & .DMTOP(1) & ","
    '        sql = sql & "DMTAIL1=" & .DMTAIL(0) & ","
    '        sql = sql & "DMTAIL2=" & .DMTAIL(1) & ","
    '        sql = sql & "NCHPOS='" & .NCHPOS & "',"
    '        sql = sql & "NCHDPTH1=" & .NCHWIDTH(0) & ","
    '        sql = sql & "NCHDPTH2=" & .NCHWIDTH(1) & ","
    '        sql = sql & "NCHWID1='" & .NCHDPTH(0) & "',"
    '        sql = sql & "NCHWID2=" & .NCHDPTH(1) & ","
    '        sql = sql & "SEEDDEG=" & .SEEDDEG & ","
    '        sql = sql & "UPLENGTH=" & .UPLENGTH & ","
    '        sql = sql & "SXLPOS=" & .SXLPOS & ","
    '        sql = sql & "BLKLEN=" & .BlkLen & ","
    '        sql = sql & "BLKWGHT=" & .BLKWGHT & ","
    '        sql = sql & "CMPTOP1=" & .Spec(0).Res(0) & ","
    '        sql = sql & "CMPTOP2=" & .Spec(0).Res(1) & ","
    '        sql = sql & "CMPTOP3=" & .Spec(0).Res(2) & ","
    '        sql = sql & "CMPTOP4=" & .Spec(0).Res(3) & ","
    '        sql = sql & "CMPTOP5=" & .Spec(0).Res(4) & ","
    '        sql = sql & "CMPTOPR=" & .Spec(0).RRG & ","
    '        sql = sql & "CMPTAIL1=" & .Spec(1).Res(0) & ","
    '        sql = sql & "CMPTAIL2=" & .Spec(1).Res(1) & ","
    '        sql = sql & "CMPTAIL3=" & .Spec(1).Res(2) & ","
    '        sql = sql & "CMPTAIL4=" & .Spec(1).Res(3) & ","
    '        sql = sql & "CMPTAIL5=" & .Spec(1).Res(4) & ","
    '        sql = sql & "CMPTAILR=" & .Spec(1).RRG & ","
    '        sql = sql & "OITOP1=" & .Spec(0).Oi(0) & ","
    '        sql = sql & "OITOP2=" & .Spec(0).Oi(1) & ","
    '        sql = sql & "OITOP3=" & .Spec(0).Oi(2) & ","
    '        sql = sql & "OITOP4=" & .Spec(0).Oi(3) & ","
    '        sql = sql & "OITOP5=" & .Spec(0).Oi(4) & ","
    '        sql = sql & "OITOPR=" & .Spec(0).ORG & ","
    '        sql = sql & "OITAIL1=" & .Spec(1).Oi(0) & ","
    '        sql = sql & "OITAIL2=" & .Spec(1).Oi(1) & ","
    '        sql = sql & "OITAIL3=" & .Spec(1).Oi(2) & ","
    '        sql = sql & "OITAIL4=" & .Spec(1).Oi(3) & ","
    '        sql = sql & "OITAIL5=" & .Spec(1).Oi(4) & ","
    '        sql = sql & "OITAILR=" & .Spec(1).ORG & ","
    '        sql = sql & "CSTOP=" & .Spec(0).Cs & ","
    '        sql = sql & "CSTAIL=" & .Spec(1).Cs & ","
    '        sql = sql & "LD1TOPMX=" & .Spec(0).LD1(0) & ","
    '        sql = sql & "LD1TOPAV=" & .Spec(0).LD1(1) & ","
    '        sql = sql & "LD1TAILM=" & .Spec(1).LD1(0) & ","
    '        sql = sql & "LD1TAILA=" & .Spec(1).LD1(1) & ","
    '        sql = sql & "LD2TOPMM=" & .Spec(0).LD2(0) & ","
    '        sql = sql & "LD2TOPAV=" & .Spec(0).LD2(1) & ","
    '        sql = sql & "LD2TAILM=" & .Spec(1).LD2(0) & ","
    '        sql = sql & "LD2TAILA=" & .Spec(1).LD2(1) & ","
    '        sql = sql & "BMDTOPMX=" & .Spec(0).BMD(0) & ","
    '        sql = sql & "BMDTOPAV=" & .Spec(0).BMD(1) & ","
    '        sql = sql & "BMDTAILM=" & .Spec(1).BMD(0) & ","
    '        sql = sql & "BMDTAILA=" & .Spec(1).BMD(1) & ","
    '        sql = sql & "GD1TOP=" & .Spec(0).GD(0) & ","
    '        sql = sql & "GD1TAIL=" & .Spec(1).GD(0) & ","
    '        sql = sql & "GD2TOP=" & .Spec(0).GD(1) & ","
    '        sql = sql & "GD2TAIL=" & .Spec(1).GD(1) & ","
    '        sql = sql & "DIA1TOP=" & .Spec(0).GD(2) & ","
    '        sql = sql & "DIA1TAIL=" & .Spec(1).GD(2) & ","
    '        sql = sql & "DIA2TOP=" & .Spec(0).GD(3) & ","
    '        sql = sql & "DIA2TAIL=" & .Spec(1).GD(3) & ","
    '        sql = sql & "LTFTOP=" & .Spec(0).Lt & ","
    '        sql = sql & "LTFTAIL=" & .Spec(1).Lt & ","
    '        sql = sql & "EPD=" & .Spec(1).EPD & ","
    '        sql = sql & "HCNO='" & .HCNO & "',"
    '        sql = sql & "KSTAFFID='" & .TSTAFFID & "',"
    '        sql = sql & "UPDDATE=sysdate,"
    '        sql = sql & "SENDFLAG='0',"
    '        sql = sql & "SENDDATE=sysdate "
    '        sql = sql & "WHERE CRYNUM='" & .blkID & "'"

    '    End With
    'End If
    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_KCryTbl_Exec = FUNCTION_RETURN_FAILURE
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_KCryTbl_Exec = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'�T�v      :�u���b�N�Ǘ��̍X�V
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:BlockMng�@�@�@,I  ,typ_TBCME040   �@,�u���b�N�Ǘ�
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�������݂̐���
'����      :
'-------�g�p���Ȃ��ق����悢�i���{�j---------
Public Function DBDRV_BlockMng_Upd_SS(BlockMng As typ_TBCME040) As FUNCTION_RETURN

    Dim sql As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc037_SQL.bas -- Function DBDRV_BlockMng_Upd"

    '' �u���b�N�Ǘ��e�[�u���̍X�V
    With BlockMng
        sql = "update TBCME040 set "
        sql = sql & "INGOTPOS=" & .INGOTPOS & ", "
        sql = sql & "LENGTH=" & .LENGTH & ", "              ' ����
        sql = sql & "REALLEN=" & .REALLEN & ", "            ' ������
        sql = sql & "BLOCKID='" & .BLOCKID & "', "          ' �u���b�NID
        sql = sql & "KRPROCCD='" & .KRPROCCD & "', "        ' ���݊Ǘ��H��
        sql = sql & "NOWPROC='" & .NOWPROC & "', "          ' ���ݍH��
        sql = sql & "LPKRPROCCD='" & .LPKRPROCCD & "', "    ' �ŏI�ʉߊǗ��H��
        sql = sql & "LASTPASS='" & .LASTPASS & "', "        ' �ŏI�ʉߍH��
        sql = sql & "DELCLS='" & .DELCLS & "',"             ' �폜�敪
        sql = sql & "LSTATCLS='" & .LSTATCLS & "', "        ' �ŏI��ԋ敪
        sql = sql & "RSTATCLS='" & .RSTATCLS & "', "        ' ������ԋ敪
        sql = sql & "HOLDCLS='" & .HOLDCLS & "', "          ' �z�[���h�敪
        sql = sql & "BDCAUS='" & .BDCAUS & "', "            ' �s�Ǘ��R
        sql = sql & "UPDDATE=sysdate, "                     ' �X�V���t
        sql = sql & "SENDFLAG='0' "                        ' ���M�t���O
        sql = sql & "where CRYNUM='" & .CRYNUM & "' "
        sql = sql & "and INGOTPOS=" & Key.POSITION
    End With
    Debug.Print sql
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_BlockMng_Upd_SS = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_BlockMng_Upd_SS = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_BlockMng_Upd_SS = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function



'�T�v      :�������H���o�p �����ԍ����͎��c�a�h���C�o
'���Ұ��@�@:�ϐ���        ,IO ,�^               ,����
'      �@�@:sCryNum �@�@�@,I  ,String         �@,�����ԍ�
'      �@�@:pCryInf �@�@�@,I  ,typ_TBCME037   �@,�������
'      �@�@:pHinDsn �@�@�@,O  ,typ_TBCME039   �@,�i�Ԑ݌v
'      �@�@:pPupEnd �@�@�@,O  ,typ_TBCMH004   �@,���グ�I������
'      �@�@:pHinSpec�@�@�@,O  ,typ_HinSpec1   �@,���i�d�l
'      �@�@:pCutInd �@�@�@,O  ,typ_CutInd     �@,�ؒf�w��
'      �@�@:sErrMsg �@�@�@,O  ,String         �@,�G���[���b�Z�[�W
'      �@�@:�߂�l        ,O  ,FUNCTION_RETURN�@,�ǂݍ��݂̐���
Public Function DBDRV_scmzc_fcmic001b_Disp(sCrynum As String, _
                                           pCryInf As typ_TBCME037, _
                                           pHinDsn() As typ_TBCME039, _
                                           pPupEnd As typ_TBCMH004, _
                                           pHinSpec() As typ_HinSpec1, _
                                           pCutInd() As typ_CutInd, _
                                           pCryOld() As typ_XSDCS, _
                                           sErrMsg As String, _
                                           fullHinban As tFullHinban) As FUNCTION_RETURN

    Dim tmpCryInf() As typ_TBCME037
    Dim tmpPupEnd() As typ_TBCMH004
    Dim rs As OraDynaset
    Dim sql As String
    Dim sDbName As String
    Dim sHin As String
    Dim recCnt As Long
    Dim i As Long
    Dim j As Long
    Dim ctcen As Double
    Dim cycen As Double
    Dim iLp2 As Integer
    
    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc016_SQL.bas -- Function DBDRV_scmzc_fcmic001b_Disp"
    sErrMsg = ""
    
        '�i�Ԃ�1�������Ȃ��̂ŌŒ肷��
    ReDim pHinDsn(1)
    pHinDsn(1).CRYNUM = sCrynum
    pHinDsn(1).INGOTPOS = 0 'Debug ���邱�� 0�ł������Ǝv������
    pHinDsn(1).hinban = fullHinban.hinban
    pHinDsn(1).REVNUM = fullHinban.mnorevno
    pHinDsn(1).FACT = fullHinban.factory
    pHinDsn(1).OPCOND = fullHinban.opecond
    'pHinDsn(1).LENGTH =   'LoadData���I��������_�œ����
    'pHinDsn(1).USECLASS
    'pHinDsn(1).REGDATE
    'pHinDsn(1).Update
    'pHinDsn(1).SENDFLAG
    'pHinDsn(1).SENDDATE

    
    '�_�~�[��1�����Ă������Ƃɂ���
    recCnt = 1
    
    '' ���i�d�l�̎擾
' ���o�K�����ڒǉ��Ή� yakimura 2002.12.01 start
    sDbName = "E018"
    j = 0
    ReDim pHinSpec(recCnt)
    For i = 1 To recCnt
        sHin = Trim(pHinDsn(i).hinban)
        If sHin <> "G" And sHin <> "Z" Then
            
            For iLp2 = 1 To j
                If (sHin = pHinSpec(iLp2).HIN.hinban) And _
                   (pHinDsn(i).OPCOND = pHinSpec(iLp2).HIN.opecond) And _
                   (pHinDsn(i).REVNUM = pHinSpec(iLp2).HIN.mnorevno) And _
                   (pHinDsn(i).FACT = pHinSpec(iLp2).HIN.factory) Then
                    Exit For
                End If
            Next iLp2
            
            If (iLp2 > j) Then
                sql = "select "
                sql = sql & "HSXTYPE, HSXCDIR, HSXD1CEN, HSXDOP, HSXDPDIR, HSXDDMIN, HSXDDMAX, HSXCTCEN, HSXCYCEN"
                sql = sql & " ,NVL(TOPREG,0) TOPREG, NVL(TAILREG,0) TAILREG, NVL(BTMSPRT,0) BTMSPRT "
                sql = sql & " from TBCME018 E018,TBCME036 E036"
                sql = sql & " where E018.HINBAN='" & pHinDsn(i).hinban & "'"
                sql = sql & " and E018.MNOREVNO=" & pHinDsn(i).REVNUM
                sql = sql & " and E018.FACTORY='" & pHinDsn(i).FACT & "'"
                sql = sql & " and E018.OPECOND='" & pHinDsn(i).OPCOND & "'"
                sql = sql & " and E036.HINBAN='" & pHinDsn(i).hinban & "'"
                sql = sql & " and E036.MNOREVNO=" & pHinDsn(i).REVNUM
                sql = sql & " and E036.FACTORY='" & pHinDsn(i).FACT & "'"
                sql = sql & " and E036.OPECOND='" & pHinDsn(i).OPCOND & "'"
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    rs.Close
                    sErrMsg = GetMsgStr("EGET2", sDbName)
                    DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
                j = j + 1
                With pHinSpec(j)
                    .HIN.hinban = pHinDsn(i).hinban
                    .HIN.mnorevno = pHinDsn(i).REVNUM
                    .HIN.factory = pHinDsn(i).FACT
                    .HIN.opecond = pHinDsn(i).OPCOND
                    .HSXTYPE = rs("HSXTYPE")    ' �^�C�v
                    .HSXCDIR = rs("HSXCDIR")    ' ����
                    .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))  ' ���a ��NULL�Ή�
                    .HSXDOP = rs("HSXDOP")      ' �����h�[�v
                    .HSXDPDIR = rs("HSXDPDIR")          ' �i�r�w�a�ʒu����
                    .HSXDDMIN = fncNullCheck(rs("HSXDDMIN"))          ' �i�r�w�a�[���� ��NULL�Ή�
                    .HSXDDMAX = fncNullCheck(rs("HSXDDMAX"))          ' �i�r�w�a�[��� ��NULL�Ή�
                    ctcen = Abs(CDbl(rs("HSXCTCEN")))
                    cycen = Abs(CDbl(rs("HSXCYCEN")))
                    .TOPREG = fncNullCheck(rs("TOPREG"))              ' TOP�K��
                    .TAILREG = fncNullCheck(rs("TAILREG"))            ' TAIL�K��
                    .BTMSPRT = fncNullCheck(rs("BTMSPRT"))            ' �{�g���͏o�K��
                    If ((ctcen = 2.83) And (cycen = 2.83)) _
                    Or ((ctcen = 4) And (cycen = 0)) _
                    Or ((ctcen = 0) And (cycen = 4)) Then
                        .HSXSDSLP = 4
                    Else
                        .HSXSDSLP = 0
                    End If
                End With
                rs.Close
            End If
        End If
    Next i
    ReDim Preserve pHinSpec(j)
    
    ReDim pCutInd(1) '1�������Ȃ��̂ŌŒ�ł悢
    
    'For i = 1 To recCnt
    '    With pCutInd(i)
    '        .INGOTPOS = rs("INGOTPOS")      ' �J�b�g�ʒu
    '        .LENGTH = rs("LENGTH")          ' ����
    '    End With
    '    rs.MoveNext
    'Next i
    'rs.Close
    
    ' LoadData���I��������_�ő������
    'pCutInd(1).INGOTPOS = '�J�b�g�ʒu(�S������)
    'pCutInd(1).LENGTH = '����(�S������)

    
    ' �����T���v���Ǘ��̎擾
    sDbName = "E043"
    sql = " where CRYNUMCS='" & sCrynum & "'"
    If DBDRV_GetTBCME043(pCryOld(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    
    DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("EGET2", sDbName)
    DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'�T�v      :�e�[�u���uTBCME037�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCME037 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcF_TBCME037_SQL.bas���ړ�)
Public Function DBDRV_GetTBCME037(records() As typ_TBCME037, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUM, DELCLS, KRPROCCD, PROCCD, LPKRPROCCD, LASTPASS, RPHINBAN, RPREVNUM, RPFACT, RPOPCOND, PRODCOND," & _
              " PGID, UPLENGTH, TOPLENG, BODYLENG, BOTLENG, FREELENG, DIAMETER, CHARGE, SEED, ADDDPCLS, ADDDPPOS, ADDDPVAL," & _
              " REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME037"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME037 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' �����ԍ�
            .DELCLS = rs("DELCLS")           ' �폜�敪
            .KRPROCCD = rs("KRPROCCD")       ' �Ǘ��H���R�[�h
            .PROCCD = rs("PROCCD")           ' �H���R�[�h
            .LPKRPROCCD = rs("LPKRPROCCD")   ' �ŏI�ʉߊǗ��H��
            .LASTPASS = rs("LASTPASS")       ' �ŏI�ʉߍH��
            .RPHINBAN = rs("RPHINBAN")       ' �˂炢�i��
            .RPREVNUM = rs("RPREVNUM")       ' �˂炢�i�Ԑ��i�ԍ������ԍ�
            .RPFACT = rs("RPFACT")           ' �˂炢�i�ԍH��
            .RPOPCOND = rs("RPOPCOND")       ' �˂炢�i�ԑ��Ə���
            .PRODCOND = rs("PRODCOND")       ' �������
            .PGID = rs("PGID")               ' �o�f�|�h�c
            .UPLENGTH = rs("UPLENGTH")       ' ���グ����
            .TOPLENG = rs("TOPLENG")         ' �s�n�o����
            .BODYLENG = rs("BODYLENG")       ' ��������
            .BOTLENG = rs("BOTLENG")         ' �a�n�s����
            .FREELENG = rs("FREELENG")       ' �t���[��
            .DIAMETER = rs("DIAMETER")       ' ���a
            .CHARGE = rs("CHARGE")           ' �`���[�W��
            .SEED = rs("SEED")               ' �V�[�h
            .ADDDPCLS = rs("ADDDPCLS")       ' �ǉ��h�[�v���
            .ADDDPPOS = rs("ADDDPPOS")       ' �ǉ��h�[�v�ʒu
            .ADDDPVAL = rs("ADDDPVAL")       ' �ǉ��h�[�v��
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME037 = FUNCTION_RETURN_SUCCESS
End Function

'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uXSDCS�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_XSDCS    ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺 (2002/07 s_cmzcTBCME043_SQL.bas���ړ�)
Public Function DBDRV_GetTBCME043(records() As typ_XSDCS, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long
    ''SQL��g�ݗ��Ă�
    sqlBase = "Select CRYNUMCS, SMPKBNCS, TBKBNCS, REPSMPLIDCS, XTALCS, INPOSCS, HINBCS, REVNUMCS, FACTORYCS, OPECS, KTKBNCS, " & _
              " BLKKTFLAGCS, CRYSMPLIDRSCS, CRYSMPLIDRS1CS, CRYSMPLIDRS2CS, CRYINDRSCS, CRYRESRS1CS, CRYRESRS2CS," & _
              " CRYSMPLIDOICS, CRYINDOICS, CRYRESOICS, CRYSMPLIDB1CS, CRYINDB1CS, CRYRESB1CS, CRYSMPLIDB2CS, CRYINDB2CS, " & _
              " CRYRESB2CS, CRYSMPLIDB3CS, CRYINDB3CS, CRYRESB3CS, CRYSMPLIDL1CS, CRYINDL1CS, CRYRESL1CS, CRYSMPLIDL2CS, " & _
              " CRYINDL2CS, CRYRESL2CS, CRYSMPLIDL3CS, CRYINDL3CS, CRYRESL3CS, CRYSMPLIDL4CS, CRYINDL4CS, CRYRESL4CS, " & _
              " CRYSMPLIDCSCS, CRYINDCSCS, CRYRESCSCS, CRYSMPLIDGDCS, CRYINDGDCS, CRYRESGDCS, CRYSMPLIDTCS, CRYINDTCS, " & _
              " CRYRESTCS, CRYSMPLIDEPCS, CRYINDEPCS,CRYRESEPCS, SMPLNUMCS, SMPLPATCS, TSTAFFCS, TDAYCS, KSTAFFCS, " & _
              " KDAYCS, SNDKCS, SNDDAYCS ,LIVKCS "
    sqlBase = sqlBase & "From XSDCS"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    Debug.Print sql
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME043 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUMCS = rs("CRYNUMCS")          ' �����ԍ�
            .SMPKBNCS = rs("SMPKBNCS")           ' �T���v���敪
            .REPSMPLIDCS = rs("REPSMPLIDCS")           ' �T���v��No
            .XTALCS = rs("CRYNUMCS")           ' �����ԍ�
            .INPOSCS = rs("INPOSCS")       ' �������ʒu
            .HINBCS = rs("HINBCS")           ' �i��
            .REVNUMCS = rs("REVNUMCS")           ' ���i�ԍ������ԍ�
            .FACTORYCS = rs("FACTORYCS")         ' �H��
            .OPECS = rs("OPECS")         ' ���Ə���
            .KTKBNCS = rs("KTKBNCS")             ' �m��敪
            .CRYINDRSCS = rs("CRYINDRSCS")       ' ���FLG�iRs)
            .CRYRESRS1CS = rs("CRYRESRS1CS")       ' ����FLG1�iRs)
            .CRYINDOICS = rs("CRYINDOICS")       ' ���FLG�iOi)
            .CRYRESOICS = rs("CRYRESOICS")       ' ����FLG�iOi)
            .CRYINDB1CS = rs("CRYINDB1CS")       ' ���FLG�iB1)
            .CRYRESB1CS = rs("CRYRESB1CS")       ' ����FLG�iB1)
            .CRYINDB2CS = rs("CRYINDB2CS")       ' ���FLG�iB2�j
            .CRYRESB2CS = rs("CRYRESB2CS")       ' ����FLG�iB2�j
            .CRYINDB3CS = rs("CRYINDB3CS")       ' ���FLG�iB3)
            .CRYRESB3CS = rs("CRYRESB3CS")       ' ����FLG�iB3)
            .CRYINDL1CS = rs("CRYINDL1CS")       ' ���FLG�iL1)
            .CRYRESL1CS = rs("CRYRESL1CS")       ' ����FLG�iL1)
            .CRYINDL2CS = rs("CRYINDL2CS")       ' ���FLG�iL2)
            .CRYRESL2CS = rs("CRYRESL2CS")       ' ����FLG�iL2)
            .CRYINDL3CS = rs("CRYINDL3CS")       ' ���FLG�iL3)
            .CRYRESL3CS = rs("CRYRESL3CS")       ' ����FLG�iL3)
            .CRYINDL4CS = rs("CRYINDL4CS")       ' ���FLG�iL4)
            .CRYRESL4CS = rs("CRYRESL4CS")       ' ����FLG�iL4)
            .CRYINDCSCS = rs("CRYINDCSCS")       ' ���FLG�iCs)
            .CRYRESCSCS = rs("CRYRESCSCS")       ' ����FLG�iCs)
            .CRYINDGDCS = rs("CRYINDGDCS")       ' ���FLG�iGD)
            .CRYRESGDCS = rs("CRYRESGDCS")       ' ����FLG�iGD)
            .CRYINDTCS = rs("CRYINDTCS")         ' ���FLG�iT)
            .CRYRESTCS = rs("CRYRESTCS")         ' ����FLG�iT)
            .CRYINDEPCS = rs("CRYINDEPCS")       ' ���FLG�iEPD)
            .CRYRESEPCS = rs("CRYRESEPCS")       ' ����FLG�iEPD)
            .SMPLNUMCS = rs("SMPLNUMCS")         ' �T���v������
            .SMPLPATCS = rs("SMPLPATCS")         ' �T���v���p�^�[��
            .TDAYCS = rs("TDAYCS")         ' �o�^���t
            .KDAYCS = rs("KDAYCS")         ' �X�V���t
            .SNDKCS = rs("SNDKCS")       ' ���M�t���O
            .SNDDAYCS = rs("SNDDAYCS")       ' ���M���t
            
            .BLKKTFLAGCS = rs("BLKKTFLAGCS")
            .CRYRESRS2CS = rs("CRYRESRS2CS")
            .LIVKCS = rs("LIVKCS")
            .TBKBNCS = rs("TBKBNCS")
            .KSTAFFCS = rs("KSTAFFCS")
            .TSTAFFCS = rs("TSTAFFCS")
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME043 = FUNCTION_RETURN_SUCCESS
End Function

'�ۗ� �{���́As_cmzcDBdriverCOM�ɓo�^����B�Ǘ����X����Ȃ̂ő҂�
'�T�v      :�����T���v���Ǘ��̍폜
'���Ұ��@�@:�ϐ���         ,IO ,�^                  ,����
'      �@�@:CrySmpOld�@�@�@,I  ,typ_XSDCS   �@      ,�V�T���v���Ǘ��i�u���b�N�j�i���j
'      �@�@:CrySmpNew�@�@�@,I  ,typ_XSDCS   �@      ,�V�T���v���Ǘ��i�u���b�N�j�i�V�j
'      �@�@:�߂�l         ,O  ,FUNCTION_RETURN�@   ,�������݂̐���
'����      :�󂯓���������̏ꍇ
'����      :2003/09/25  �쐬 ��n
Public Function DBDRV_CrySmp_Del(CrySmpOld() As typ_XSDCS) As FUNCTION_RETURN

    Dim sql As String
    Dim lFlg As Boolean
    Dim i As Long
    Dim j As Long

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_CrySmp_Del"

    DBDRV_CrySmp_Del = FUNCTION_RETURN_SUCCESS

    For i = 1 To UBound(CrySmpOld)
        With CrySmpOld(i)
            sql = "Delete XSDCS where "
            sql = sql & "CRYNUMCS = '" & .CRYNUMCS & "' and "
            sql = sql & "TBKBNCS = '" & .TBKBNCS & "'"
            '' WriteDBLog sql
            If OraDB.ExecuteSQL(sql) <= 0 Then
                DBDRV_CrySmp_Del = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        End With
    Next i

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_CrySmp_Del = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

Public Function DBDRV_TBCMI002_Exec(record As PURCHASE_CRYSTAL, UpDateFlag As Boolean) As FUNCTION_RETURN

    Dim sql As String
    Dim fullHinban As tFullHinban
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc037_SQL.bas -- Function DBDRV_KCryTbl_Exec"

    DBDRV_TBCMI002_Exec = FUNCTION_RETURN_SUCCESS

    '12���i�Ԃ����߂�
    If GetLastHinban(record.hinban, fullHinban) = FUNCTION_RETURN_FAILURE Then
        DBDRV_TBCMI002_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
'    If Not UpDateFlag Then
        
        'TBCMI002 �֒ǉ�
        
        sql = "insert into TBCMI002 ( "
        sql = sql & "CRYNUM,"
        sql = sql & "INGOTPOS,"
        sql = sql & "LENGTH,"
        sql = sql & "TRANCNT,"
        sql = sql & "KRPROCCD,"
        sql = sql & "PROCCODE,"
        sql = sql & "DMTOP1,"
        sql = sql & "DMTOP2,"
        sql = sql & "DMTAIL1,"
        sql = sql & "DMTAIL2,"
        sql = sql & "NCHPOS,"
        sql = sql & "NCHDPTH,"
        sql = sql & "NCHWIDTH,"
        sql = sql & "BDLNTOP,"
        sql = sql & "BDCDTOP,"
        sql = sql & "BDLNTAIL,"
        sql = sql & "BDCDTAIL,"
        sql = sql & "TSTAFFID,"
        sql = sql & "REGDATE,"
        sql = sql & "KSTAFFID,"
        sql = sql & "UPDDATE,"
        sql = sql & "SENDFLAG,"
        sql = sql & "SENDDATE) "
        With record
            sql = sql & " select "
            sql = sql & "'" & Left(.blkID, 9) & "000" & "',"
            sql = sql & .SXLPOS & ","
            sql = sql & .BlkLen & ","
            sql = sql & "nvl(max(TRANCNT), 0) + 1 ,"
            sql = sql & "'" & .KRPROCCD & "',"
            sql = sql & "'" & .PROCCODE & "',"
            sql = sql & .DMTOP(0) & ","
            sql = sql & .DMTOP(1) & ","
            sql = sql & .DMTAIL(0) & ","
            sql = sql & .DMTAIL(1) & ","
            sql = sql & "'" & .NCHPOS & "',"
            sql = sql & .NCHDPTH(0) & ","
            sql = sql & .NCHWIDTH(0) & ","
            sql = sql & "0,"
            sql = sql & "' ',"
            sql = sql & "0,"
            sql = sql & "' ',"
            sql = sql & "'" & .TSTAFFID & "',"
            sql = sql & "sysdate ,"
            sql = sql & "'" & .TSTAFFID & "',"
            sql = sql & "sysdate,"
            sql = sql & "'0',"
            sql = sql & "sysdate  "
            sql = sql & " from TBCMI002 "
            sql = sql & " where CRYNUM='" & .blkID & "' "
        End With
'    Else
'        '�X�V��
'
'        With record
'            sql = "UPDATE TBCMI002 SET "
'            sql = sql & "INGOTPOS= " & .SXLPOS & ","
'            sql = sql & "LENGTH= " & .BlkLen & ","
'            sql = sql & "KRPROCCD= '" & .KRPROCCD & "',"
'            sql = sql & "PROCCODE= '" & .PROCCODE & "',"
'            sql = sql & "DMTOP1= " & .DMTOP(0) & ","
'            sql = sql & "DMTOP2= " & .DMTOP(1) & ","
'            sql = sql & "DMTAIL1= " & .DMTAIL(0) & ","
'            sql = sql & "DMTAIL2= " & .DMTAIL(1) & ","
'            sql = sql & "NCHPOS= '" & .NCHPOS & "',"
'            sql = sql & "NCHDPTH= " & .NCHDPTH(0) & ","
'            sql = sql & "NCHWIDTH= " & .NCHWIDTH(0) & ","
'            sql = sql & "TSTAFFID= '" & .TSTAFFID & "',"
'            sql = sql & "REGDATE= sysdate ,"
'            sql = sql & "KSTAFFID= '" & .TSTAFFID & "',"
'            sql = sql & "UPDDATE= sysdate,"
'            sql = sql & "SENDFLAG= '0',"
'            sql = sql & "SENDDATE= sysdate  "
'            sql = sql & "WHERE CRYNUM='" & Left(.blkID, 9) & "000" & "'"
'        End With
'    End If
    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_TBCMI002_Exec = FUNCTION_RETURN_FAILURE
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_TBCMI002_Exec = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'�T�v      :�����T���v���Ǘ��̑}���^�X�V
'���Ұ��@�@:�ϐ���         ,IO ,�^                  ,����
'      �@�@:CrySmpOld�@�@�@,I  ,typ_XSDCS   �@      ,�V�T���v���Ǘ��i�u���b�N�j�i���j
'      �@�@:CrySmpNew�@�@�@,I  ,typ_XSDCS   �@      ,�V�T���v���Ǘ��i�u���b�N�j�i�V�j
'      �@�@:�߂�l         ,O  ,FUNCTION_RETURN�@   ,�������݂̐���
'����      :�Â����R�[�h���݂čX�V���}�����𔻕ʂ���
'����      :2001/07/12  �쐬 ���{
Public Function DBDRV_CrySmp_UpdIns037Only(CrySmpOld() As typ_XSDCS, CrySmpNew() As typ_XSDCS) As FUNCTION_RETURN

    Dim sql As String
    Dim lFlg As Boolean
    Dim i As Long
    Dim j As Long
    Dim result As String

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_CrySmp_UpdIns"

    DBDRV_CrySmp_UpdIns037Only = FUNCTION_RETURN_SUCCESS
    
    For i = 1 To UBound(CrySmpNew)
        With CrySmpNew(i)
            lFlg = False
''            For j = 1 To UBound(CrySmpOld)
''                If CrySmpOld(j).XTALCS = .XTALCS And _
''                   CrySmpOld(j).SMPKBNCS = .SMPKBNCS Then
''                    sql = "update XSDCS set "
''                    sql = sql & "HINBCS='" & .HINBCS & "', "                ' �i��
''                    sql = sql & "REVNUMCS=" & .REVNUMCS & ", "              ' ���i�ԍ������ԍ�
''                    sql = sql & "FACTORYCS='" & .FACTORYCS & "', "          ' �H��
''                    sql = sql & "OPECS='" & .OPECS & "', "                  ' ���Ə���
''                    sql = sql & "KTKBNCS='" & .KTKBNCS & "', "              ' �m��敪
''                    sql = sql & "REPSMPLIDCS='" & Abs(.REPSMPLIDCS) & "', " ' �T���v���m��
''                    If .CRYINDRSCS = "2" Then
''                        .CRYINDRSCS = "1"
''                    End If
''                    sql = sql & "CRYINDRSCS='" & .CRYINDRSCS & "', "        ' ���FLG�iRs)
''                    If .CRYINDOICS = "2" Then
''                        .CRYINDOICS = "1"
''                    End If
''                    sql = sql & "CRYINDOICS='" & .CRYINDOICS & "', "        ' ���FLG�iOi)
''                    If .CRYINDB1CS = "2" Then
''                        .CRYINDB1CS = "1"
''                    End If
''                    sql = sql & "CRYINDB1CS='" & .CRYINDB1CS & "', "        ' ���FLG�iB1)
''                    If .CRYINDB2CS = "2" Then
''                        .CRYINDB2CS = "1"
''                    End If
''                    sql = sql & "CRYINDB2CS='" & .CRYINDB2CS & "', "        ' ���FLG�iB2)
''                    If .CRYINDB3CS = "2" Then
''                        .CRYINDB3CS = "1"
''                    End If
''                    sql = sql & "CRYINDB3CS='" & .CRYINDB3CS & "', "        ' ���FLG�iB3)
''                    If .CRYINDL1CS = "2" Then
''                        .CRYINDL1CS = "1"
''                    End If
''                    sql = sql & "CRYINDL1CS='" & .CRYINDL1CS & "', "        ' ���FLG�iL1)
''                    If .CRYINDL2CS = "2" Then
''                        .CRYINDL2CS = "1"
''                    End If
''                    sql = sql & "CRYINDL2CS='" & .CRYINDL2CS & "', "        ' ���FLG�iL2)
''                    If .CRYINDL3CS = "2" Then
''                        .CRYINDL3CS = "1"
''                    End If
''                    sql = sql & "CRYINDL3CS='" & .CRYINDL3CS & "', "        ' ���FLG�iL3)
''                    If .CRYINDL4CS = "2" Then
''                        .CRYINDL4CS = "1"
''                    End If
''                    sql = sql & "CRYINDL4CS='" & .CRYINDL4CS & "', "        ' ���FLG�iL4)
''                    If .CRYINDCSCS = "2" Then
''                        .CRYINDCSCS = "1"
''                    End If
''                    sql = sql & "CRYINDCSCS='" & .CRYINDCSCS & "', "        ' ���FLG�iCs)
''                    If .CRYINDGDCS = "2" Then
''                        .CRYINDGDCS = "1"
''                    End If
''                    sql = sql & "CRYINDGDCS='" & .CRYINDGDCS & "', "        ' ���FLG�iGD)
''                    If .CRYINDTCS = "2" Then
''                        .CRYINDTCS = "1"
''                    End If
''                    sql = sql & "CRYINDTCS='" & .CRYINDTCS & "', "          ' ���FLG�iT)
''                    If .CRYINDEPCS = "2" Then
''                        .CRYINDEPCS = "1"
''                    End If
''                    sql = sql & "CRYINDEPCS='" & .CRYINDEPCS & "', "        ' ���FLG�iEPD)
''
''                    sql = sql & "CRYRESRS1CS='" & .CRYRESRS1CS & "', "      ' ����FLG1�iRs)
''                    sql = sql & "CRYRESOICS='" & .CRYRESOICS & "', "        ' ����FLG�iOi)
''                    sql = sql & "CRYRESB1CS='" & .CRYRESB1CS & "', "        ' ����FLG�iB1)
''                    sql = sql & "CRYRESB2CS='" & .CRYRESB2CS & "', "        ' ����FLG�iB2)
''                    sql = sql & "CRYRESB3CS='" & .CRYRESB3CS & "', "        ' ����FLG�iB3)
''                    sql = sql & "CRYRESL1CS='" & .CRYRESL1CS & "', "        ' ����FLG�iL1)
''                    sql = sql & "CRYRESL2CS='" & .CRYRESL2CS & "', "        ' ����FLG�iL2)
''                    sql = sql & "CRYRESL3CS='" & .CRYRESL3CS & "', "        ' ����FLG�iL3)
''                    sql = sql & "CRYRESL4CS='" & .CRYRESL4CS & "', "        ' ����FLG�iL4)
''                    sql = sql & "CRYRESCSCS='" & .CRYRESCSCS & "', "        ' ����FLG�iCs)
''                    sql = sql & "CRYRESGDCS='" & .CRYRESGDCS & "', "        ' ����FLG�iGD)
''                    sql = sql & "CRYRESTCS='" & .CRYRESTCS & "', "          ' ����FLG�iT)
''                    sql = sql & "CRYRESEPCS='" & .CRYRESEPCS & "', "        ' ����FLG�iEPD)
''                    sql = sql & "SMPLNUMCS=" & .SMPLNUMCS & ", "            ' �T���v������
''                    sql = sql & "SMPLPATCS='" & .SMPLPATCS & "', "          ' �T���v���p�^�[��
''                    sql = sql & "KDAYCS=sysdate, "                          ' �X�V���t
''                    sql = sql & "SNDKCS='0' "                               ' ���M�t���O
''                    sql = sql & " where XTALCS='" & .XTALCS & "'"
''                    sql = sql & " and TBKBNCS='" & .TBKBNCS & "'"
''
''                    WriteDBLog sql
''                    Debug.Print sql
''                    If OraDB.ExecuteSQL(sql) <= 0 Then
''                        DBDRV_CrySmp_UpdIns037Only = FUNCTION_RETURN_FAILURE
''                        GoTo proc_exit
''                    End If
''                    lFlg = True
''                    Exit For
''                End If
''            Next j

            If lFlg <> True Then
                sql = "insert into XSDCS ("
                sql = sql & "CRYNUMCS,"         '�u���b�NID
                sql = sql & "SMPKBNCS,"         '�T���v���敪
                sql = sql & "TBKBNCS,"          'T/B�敪
                sql = sql & "REPSMPLIDCS,"      '��\�T���v��ID
                sql = sql & "XTALCS,"           '�����ԍ�
                sql = sql & "INPOSCS,"          '�������ʒu
                sql = sql & "HINBCS,"           '�i��
                sql = sql & "REVNUMCS,"         '���i�ԍ������ԍ�
                sql = sql & "FACTORYCS,"        '�H��
                sql = sql & "OPECS,"            '���Ɣԍ�
                sql = sql & "KTKBNCS,"          '�m��敪
                sql = sql & "BLKKTFLAGCS,"      '�u���b�N�m��t���O
                sql = sql & "CRYSMPLIDRSCS,"    '�T���v��ID(Rs)
                sql = sql & "CRYSMPLIDRS1CS,"   '����T���v��ID1�iRs�j
                sql = sql & "CRYSMPLIDRS2CS,"   '����T���v��ID2�iRs�j
                sql = sql & "CRYINDRSCS,"       '���FLG(Rs)
                sql = sql & "CRYRESRS1CS,"      '����FLG1(Rs)
                sql = sql & "CRYRESRS2CS,"      '����FLG2(Rs)
                sql = sql & "CRYSMPLIDOICS,"    '�T���v��ID�iOi�j
                sql = sql & "CRYINDOICS,"       '���FLG�iOi�j
                sql = sql & "CRYRESOICS,"       '����FLG�iOi�j
                sql = sql & "CRYSMPLIDB1CS,"    '�T���v��ID�iB1�j
                sql = sql & "CRYINDB1CS,"       '���FLG�iB1�j
                sql = sql & "CRYRESB1CS,"       '����FLG�iB1�j
                sql = sql & "CRYSMPLIDB2CS,"    '�T���v��ID�iB2�j
                sql = sql & "CRYINDB2CS,"       '���FLG�iB2�j
                sql = sql & "CRYRESB2CS,"       '����FLG�iB2�j
                sql = sql & "CRYSMPLIDB3CS,"    '�T���v��ID�iB3�j
                sql = sql & "CRYINDB3CS,"       '���FLG�iB3�j
                sql = sql & "CRYRESB3CS,"       '����FLG�iB3�j
                sql = sql & "CRYSMPLIDL1CS,"    '�T���v��ID�iL1�j
                sql = sql & "CRYINDL1CS,"       '���FLG�iL1�j
                sql = sql & "CRYRESL1CS,"       '����FLG�iL1�j
                sql = sql & "CRYSMPLIDL2CS,"    '�T���v��ID�iL2�j
                sql = sql & "CRYINDL2CS,"       '���FLG�iL2�j
                sql = sql & "CRYRESL2CS,"       '����FLG�iL2�j
                sql = sql & "CRYSMPLIDL3CS,"    '�T���v��ID�iL3�j
                sql = sql & "CRYINDL3CS,"       '���FLG�iL3�j
                sql = sql & "CRYRESL3CS,"       '����FLG�iL3�j
                sql = sql & "CRYSMPLIDL4CS,"    '�T���v��ID�iL4�j
                sql = sql & "CRYINDL4CS,"       '���FLG�iL4�j
                sql = sql & "CRYRESL4CS,"       '����FLG�iL4�j
                sql = sql & "CRYSMPLIDCSCS,"    '�T���v��ID�iCS�j
                sql = sql & "CRYINDCSCS,"       '���FLG�iCS�j
                sql = sql & "CRYRESCSCS,"       '����FLG�iCS�j
                sql = sql & "CRYSMPLIDGDCS,"    '�T���v��ID�iGD�j
                sql = sql & "CRYINDGDCS,"       '���FLG�iGD�j
                sql = sql & "CRYRESGDCS,"       '����FLG�iGD�j
                sql = sql & "CRYSMPLIDTCS,"     '�T���v��ID�iT�j
                sql = sql & "CRYINDTCS,"        '���FLG�iT�j
                sql = sql & "CRYRESTCS,"        '����FLG�iT�j
                sql = sql & "CRYSMPLIDEPCS,"    '�T���v��ID�iEPD�j
                sql = sql & "CRYINDEPCS,"       '���FLG�iEPD�j
                sql = sql & "CRYRESEPCS,"       '����FLG�iEPD�j
                sql = sql & "SMPLNUMCS,"        '�T���v������
                sql = sql & "SMPLPATCS,"        '�T���v���p�^�[��
                sql = sql & "TSTAFFCS,"         '�o�^�Ј�ID
                sql = sql & "TDAYCS,"           '�o�^���t
                sql = sql & "KSTAFFCS,"         '�X�V�Ј�ID
                sql = sql & "KDAYCS,"           '�X�V���t
                sql = sql & "SNDKCS,"           '���M�t���O
                sql = sql & "SNDDAYCS,"         '���M���t
                sql = sql & "LIVKCS)"           '�����敪
                sql = sql & " values ('"
                sql = sql & .CRYNUMCS & "', '"          '�u���b�NID
                sql = sql & .SMPKBNCS & "', '"          '�T���v���敪
                sql = sql & .TBKBNCS & "', "            'T/B�敪
                sql = sql & .REPSMPLIDCS & ", '"        '��\�T���v��ID
                sql = sql & .XTALCS & "', "             '�����ԍ�
                sql = sql & .INPOSCS & ", '"            '�������ʒu
                sql = sql & .HINBCS & "', "             '�i��
                sql = sql & .REVNUMCS & ", '"           '���i�ԍ������ԍ�
                sql = sql & .FACTORYCS & "', '"         '�H��
                sql = sql & .OPECS & "', '"             '���Ə���
                sql = sql & .KTKBNCS & "', '"           '�m��敪
                sql = sql & .BLKKTFLAGCS & "', "        '�u���b�N�m��t���O
                
                If .CRYINDRSCS = "2" Then
                    .CRYINDRSCS = "1"
                End If
                If .CRYINDRSCS = "1" Then
                    .CRYSMPLIDRSCS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDRSCS = 0
                End If
                sql = sql & .CRYSMPLIDRSCS & ", "       '�T���v��ID�iRs�j
                sql = sql & .CRYSMPLIDRS1CS & ", "      '����T���v��ID1�iRs�j
                sql = sql & .CRYSMPLIDRS2CS & ", '"     '����T���v��ID2�iRs�j
                sql = sql & .CRYINDRSCS & "', '"        '���FLG�iRs�j
                sql = sql & .CRYRESRS1CS & "', '"       '����FLG1�iRs�j
                sql = sql & .CRYRESRS2CS & "', "        '����FLG2�iRs�j
                
                If .CRYINDOICS = "2" Then
                    .CRYINDOICS = "1"
                End If
                If .CRYINDOICS = "1" Then
                    .CRYSMPLIDOICS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDOICS = 0
                End If
                sql = sql & .CRYSMPLIDOICS & ", '"      '�T���v��ID�iOi�j
                sql = sql & .CRYINDOICS & "', '"        '���FLG�iOi�j
                sql = sql & .CRYRESOICS & "', "         '����FLG�iOi�j
                
                If .CRYINDB1CS = "2" Then
                    .CRYINDB1CS = "1"
                End If
                If .CRYINDB1CS = "1" Then
                    .CRYSMPLIDB1CS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDB1CS = 0
                End If
                sql = sql & .CRYSMPLIDB1CS & ", '"      '�T���v��ID�iB1�j
                sql = sql & .CRYINDB1CS & "', '"        '���FLG�iB1�j
                sql = sql & .CRYRESB1CS & "', "         '����FLG�iB1�j
                
                
                If .CRYINDB2CS = "2" Then
                    .CRYINDB2CS = "1"
                End If
                If .CRYINDB2CS = "1" Then
                    .CRYSMPLIDB2CS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDB2CS = 0
                End If
                sql = sql & .CRYSMPLIDB2CS & ", '"      '�T���v��ID�iB2�j
                sql = sql & .CRYINDB2CS & "', '"        '���FLG�iB2�j
                sql = sql & .CRYRESB2CS & "', "         '����FLG�iB2�j
                
                If .CRYINDB3CS = "2" Then
                    .CRYINDB3CS = "1"
                End If
                If .CRYINDB3CS = "1" Then
                    .CRYSMPLIDB3CS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDB3CS = 0
                End If
                sql = sql & .CRYSMPLIDB3CS & ", '"      '�T���v��ID�iB3�j
                sql = sql & .CRYINDB3CS & "', '"        '���FLG�iB3�j
                sql = sql & .CRYRESB3CS & "', "         '����FLG�iB3�j
                
                If .CRYINDL1CS = "2" Then
                    .CRYINDL1CS = "1"
                End If
                If .CRYINDL1CS = "1" Then
                    .CRYSMPLIDL1CS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDL1CS = 0
                End If
                sql = sql & .CRYSMPLIDL1CS & ", '"      '�T���v��ID�iL1�j
                sql = sql & .CRYINDL1CS & "', '"        '���FLG�iL1�j
                sql = sql & .CRYRESL1CS & "', "         '����FLG�iL1�j
                
                If .CRYINDL2CS = "2" Then
                    .CRYINDL2CS = "1"
                End If
                If .CRYINDL2CS = "1" Then
                    .CRYSMPLIDL2CS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDL2CS = 0
                End If
                sql = sql & .CRYSMPLIDL2CS & ", '"      '�T���v��ID�iL2�j
                sql = sql & .CRYINDL2CS & "', '"        '���FLG�iL2�j
                sql = sql & .CRYRESL2CS & "', "         '����FLG�iL2�j
                
                If .CRYINDL3CS = "2" Then
                    .CRYINDL3CS = "1"
                End If
                If .CRYINDL3CS = "1" Then
                    .CRYSMPLIDL3CS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDL3CS = 0
                End If
                sql = sql & .CRYSMPLIDL3CS & ", '"      '�T���v��ID�iL3�j
                sql = sql & .CRYINDL3CS & "', '"        '���FLG�iL3�j
                sql = sql & .CRYRESL3CS & "', "         '����FLG�iL3�j
                
                If .CRYINDL4CS = "2" Then
                    .CRYINDL4CS = "1"
                End If
                If .CRYINDL4CS = "1" Then
                    .CRYSMPLIDL4CS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDL4CS = 0
                End If
                sql = sql & .CRYSMPLIDL4CS & ", '"      '�T���v��ID�iL4�j
                sql = sql & .CRYINDL4CS & "', '"        '���FLG�iL4�j
                sql = sql & .CRYRESL4CS & "', "         '����FLG�iL4�j
                
                If .CRYINDCSCS = "2" Then
                    .CRYINDCSCS = "1"
                End If
                If .CRYINDCSCS = "1" Then
                    .CRYSMPLIDCSCS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDCSCS = 0
                End If
                sql = sql & .CRYSMPLIDCSCS & ", '"      '�T���v��ID�iCS�j
                sql = sql & .CRYINDCSCS & "', '"        '���FLG�iCS�j
                sql = sql & .CRYRESCSCS & "', "         '����FLG�iCS�j
                
                If .CRYINDGDCS = "2" Then
                    .CRYINDGDCS = "1"
                End If
                If .CRYINDGDCS = "1" Then
                    .CRYSMPLIDGDCS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDGDCS = 0
                End If
                sql = sql & .CRYSMPLIDGDCS & ", '"      '�T���v��ID�iGD�j
                sql = sql & .CRYINDGDCS & "', '"        '���FLG�iGD�j
                sql = sql & .CRYRESGDCS & "', "         '����FLG�iGD�j
                
                If .CRYINDTCS = "2" Then
                    .CRYINDTCS = "1"
                End If
                If .CRYINDTCS = "1" Then
                    .CRYSMPLIDTCS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDTCS = 0
                End If
                sql = sql & .CRYSMPLIDTCS & ", '"       '�T���v��ID�iT�j
                sql = sql & .CRYINDTCS & "', '"         '���FLG�iT�j
                sql = sql & .CRYRESTCS & "', "          '����FLG�iT�j
                
                If .CRYINDEPCS = "2" Then
                    .CRYINDEPCS = "1"
                End If
                If .CRYINDEPCS = "1" Then
                    .CRYSMPLIDEPCS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDEPCS = 0
                End If
                sql = sql & .CRYSMPLIDEPCS & ", '"      '�T���v��ID�iEPD�j
                sql = sql & .CRYINDEPCS & "', '"        '���FLG�iEPD�j
                sql = sql & .CRYRESEPCS & "', "         '����FLG�iEPD�j
                sql = sql & .SMPLNUMCS & ", "           '�T���v������
                sql = sql & "' ', '"                    '�T���v���p�^�[��
                sql = sql & .TSTAFFCS & "', "           '�o�^�Ј�ID
                sql = sql & "sysdate, '"                '�o�^���t
                sql = sql & .KSTAFFCS & "', "           '�X�V�Ј�ID
                sql = sql & "sysdate, "                 '�X�V���t
                sql = sql & "'0', "                     '���M�t���O
                sql = sql & "sysdate,"                  '���M���t
                sql = sql & "'0')"                      '�����敪
                
                '' WriteDBLog sql
                If OraDB.ExecuteSQL(sql) <= 0 Then
                    DBDRV_CrySmp_UpdIns037Only = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            End If
        End With
    Next i

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_CrySmp_UpdIns037Only = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function


'�T�v      :�e�[�u���̍폜����
'���Ұ��@�@:�ϐ���         ,IO ,�^                  ,����
'      �@�@:sql�@�@      �@,I  ,String      �@      ,�폜SQL��
'          :sTable        ,I  ,String              ,�폜�e�[�u��
'      �@�@:�߂�l         ,O  ,FUNCTION_RETURN�@   ,�������݂̐���
'����      :�w���P�������/������A���ɑ��݂���f�[�^���폜����
'����      :2003/10/31 ooba

Public Function DBDRV_DeleteTable(sql As String, sTable As String) As FUNCTION_RETURN


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc037_SQL.bas -- Function DBDRV_DeleteTable"

    DBDRV_DeleteTable = FUNCTION_RETURN_FAILURE
    
Debug.Print sql
    
    If OraDB.ExecuteSQL(sql) < 1 Then
        Debug.Print "<" & sTable & "> �폜�f�[�^����"
    Else
        '' WriteDBLog sql
    End If
    
    DBDRV_DeleteTable = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
    
End Function


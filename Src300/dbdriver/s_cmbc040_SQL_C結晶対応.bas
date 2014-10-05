Attribute VB_Name = "s_cmbc040_SQL"
Option Explicit

Public Type typ_cmlc001e_Disp
    ' SXL�Ǘ�
    CRYNUM As String * 12           ' �����ԍ�
    INGOTPOS As Integer             ' �������J�n�ʒu
    SXLID As String * 13            ' ��SXLID
    hinban As String * 8            ' �i��
    LENGTH As Integer               ' ����
    COUNT As Integer                ' �\�薇��
    ENDDATE As Date                 ' �������t�iSXL�Ǘ�.�o�^���t�j
    ' WF�z�[���h�i�����j����
    HLDCLASSOLD As String * 1       ' ���z�[���h�����敪 (0:�z�[���h���������A1:�z�[���h����)
    HLDCLASS As String * 1          ' �z�[���h�����敪 (0:�z�[���h���������A1:�z�[���h����)
    HLDDATE As Date                 ' �z�[���h���t(.�o�^���t)
    HLDSTAFFNAME As String          ' �z�[���h�S����(.�o�^�Ј�ID)
    HLDCAUSE As String * 2          ' �z�[���h���R ('SC','17')
    HLDCMNT As String               ' �z�[���h�R�����g
    MUKESAKI As String              ' ���� 2007/09/04 SPK Tsutsumi Add
    AGRSTATUS As String             ' ���F�m�F�敪      add SETkimizuka
    STOP    As String               ' ��~ add SETkimizuka
    CAUSE   As String               ' ��~���R add SETkimizuka
    PRINTNO As String               ' ��s�]�� add SETkimizuka
    '��EDI����ݸ�Ή� 2009/12/4 Add Strat SPK habuki������
    EDIFLG As String                ' EDI�׸�(��:�S��null�AOK:���M�ΏۗL�ANG:���M�Ώۖ�)
    '��EDI����ݸ�Ή� 2009/12/4 Add End   SPK habuki������
End Type

'2002/09/05 ADD hitec)N.MATSUMOTO Start

'�u���b�N�Ǘ�
Public Type typ_cmkc001f_Block
    'E040 �u���b�N�Ǘ�
    INGOTPOS As Integer         ' �������J�n�ʒu
    LENGTH As Integer           ' ����
    REALLEN As Integer          ' ������
    KRPROCCD As String * 5      ' ���݊Ǘ��H��
    NOWPROC As String * 5       ' ���ݍH��
    LPKRPROCCD As String * 5    ' �ŏI�ʉߊǗ��H��
    LASTPASS As String * 5      ' �ŏI�ʉߍH��
    DELCLS As String * 1        ' �폜�敪
    RSTATCLS As String * 1      ' ������ԋ敪
    LSTATCLS As String * 1      ' �ŏI��ԋ敪 */
    'E037 �������Ǘ�
    SEED As String              'SEED
End Type


'�d�l�擾�p
Public Type typ_cmkc001f_Disp
    '�i�ԊǗ�
    hinban As String * 8              ' �i��
    INGOTPOS As Integer               ' �������J�n�ʒu
    REVNUM As Integer                 ' ���i�ԍ������ԍ�
    factory As String * 1             ' �H��
    opecond As String * 1             ' ���Ə���
    LENGTH As Integer                 ' ����
    '���i�d�lSXL�f�[�^
    HSXD1CEN As Double                ' �i�r�w���a�P���S
    HSXRMIN As Double                 ' �i�r�w���R����
    HSXRMAX As Double                 ' �i�r�w���R���
    HSXRMBNP As Double                ' �i�r�w���R�ʓ����z
    HSXRHWYS As String * 1            ' �i�r�w���R�ۏؕ��@�Q��
    HSXONMIN As Double                ' �i�r�w�_�f�Z�x����
    HSXONMAX As Double                ' �i�r�w�_�f�Z�x���
    HSXONMBP As Double                ' �i�r�w�_�f�Z�x�ʓ����z
    HSXONHWS As String * 1            ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
    HSXCNMIN As Double                ' �i�r�w�Y�f�Z�x����
    HSXCNMAX As Double                ' �i�r�w�Y�f�Z�x���
    HSXCNHWS As String * 1            ' �i�r�w�Y�f�Z�x�ۏؕ��@�Q��
    HSXTMMAX As Double                ' �i�r�w�]�ʖ��x���          ���ڒǉ��C�C���Ή� 2003.05.20 yakimura
    HSXBMnAN(1 To 3) As Double        ' �i�r�w�a�l�cn ���ω���
    HSXBMnAX(1 To 3) As Double        ' �i�r�w�a�l�cn ���Ϗ��
    HSXBMnHS(1 To 3) As String * 1    ' �i�r�w�a�l�cn �ۏؕ��@�Q��
    HSXOFnAX(1 To 4) As Double        ' �i�r�w�n�r�en���Ϗ��
    HSXOFnMX(1 To 4) As Double        ' �i�r�w�n�r�en���
    HSXOFnHS(1 To 4) As String * 1    ' �i�r�w�n�r�en �ۏؕ��@�Q��
    HSXDENMX As Integer               ' �i�r�w�c�������
    HSXDENMN As Integer               ' �i�r�w�c��������
    HSXDENHS As String * 1            ' �i�r�w�c�����ۏؕ��@�Q��
    HSXDVDMX As Integer               ' �i�r�w�c�u�c�Q���
    HSXDVDMN As Integer               ' �i�r�w�c�u�c�Q����
    HSXDVDHS As String * 1            ' �i�r�w�c�u�c�Q�ۏؕ��@�Q��
    HSXLDLMX As Integer               ' �i�r�w�k�^�c�k���
    HSXLDLMN As Integer               ' �i�r�w�k�^�c�k����
    HSXLDLHS As String * 1            ' �i�r�w�k�^�c�k�ۏؕ��@�Q��
    HSXLTMIN As Integer               ' �i�r�w�k�^�C������
    HSXLTMAX As Integer               ' �i�r�w�k�^�C�����
    HSXLTHWS As String * 1            ' �i�r�w�k�^�C���ۏؕ��@�Q��
    HSXDPDIR As String * 2            ' �i�r�w�a�ʒu����
    HSXDPDRC As String * 1            ' �i�r�w�a�ʒu����
    HSXDWMIN As Double                ' �i�r�w�a�Љ���
    HSXDWMAX As Double                ' �i�r�w�a�Џ��
    HSXDDMIN As Double                ' �i�r�w�a�[����
    HSXDDMAX As Double                ' �i�r�w�a�[���
    HSXD1MIN As Double                ' �i�r�w���a�P����
    HSXD1MAX As Double                ' �i�r�w���a�P���
    HSXCTCEN As Double                ' �i�r�w�����ʌX�c���S
    HSXCYCEN As Double                ' �i�r�w�����ʌX�����S
    EPDUP As Integer                  ' ���������Ǘ� EPD�@���
End Type

Private Type tCsData
    'Cs����
    SXL_CS_SMPPOS As Integer        ' ����ʒu
    SXLCS_CSMEAS As Double          ' Cs�l
    SXLCS_70PPRE As Double          ' 70%����l
End Type

Public strBlockID()    As String
Public sProSXLID       As String    '����SXLID�@06/03/24 ooba

Public Const PROCD_WFC_SAINUKISI = "CW760"              ' WF�Z���^�[�Ĕ���
Public Const PROCD_SXL_MAP = "TX860"                    ' �V���O���}�b�v
'2002/09/05 ADD hitec)N.MATSUMOTO end

''***********************************************************
'Micron�Ή� 2011/01/14 Add start tkimura
Public Type typ_Y011
    LOTID As String          '' �u���b�NID
    BLOCKSEQ As Integer      '' �u���b�N���A��
    RITOP_POS As String       '' ����ʒu
End Type
'Micron�Ή� 2011/01/14 Add end tkimura
''***********************************************************

Public Function DBDRV_fcmlc001e_Disp(records() As typ_cmlc001e_Disp) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset
    Dim i As Integer
    Dim j As Integer    ' 2007/09/13 SPK Tsutsumi Add
    
    '2004/07/15 koyama
    Dim sWFHOLDDATE   As String
    Dim sUSER_ID As String
    
    Dim sOldID      As String   '09/03/17 add SETkimizuka
    Dim iCnt        As Integer  '09/03/17 add SETkimizuka

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function DBDRV_fcmlc001e_Disp"

''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
    sql = ""
    sql = sql & " SELECT"
    sql = sql & "   WFS.XTALCB as CRYNUM"                               ''�����ԍ�
    sql = sql & "  ,WFS.INPOSCB as INGOTPOS"                            ''�������J�n�ʒu
    sql = sql & "  ,WFS.SXLIDCB as SXLID"                               ''SXLID
    sql = sql & "  ,WFS.HINBCB as HINBAN"                               ''�i��
    sql = sql & "  ,WFS.RLENCB as LENGTH"                               ''���_����
    sql = sql & "  ,WFS.maicb as COUNT"                                 ''������
    ' ����ΰ��ނ𗬓��Ď��f�[�^�֒u������ del SETkimizuka Start  09/03/17
    'sql = sql & "  ,NVL(HLD.HLDCLASS,'0') as HLDCLASS"
    'sql = sql & "  ,NVL(HLD.REGDATE,SYSDATE) as HLDDATE"
    'sql = sql & "  ,NVL(HLD.STAFFNAME,' ') as HLDSTAFFNAME"
    'sql = sql & "  ,NVL(HLD.HLDCAUSE,' ') as HLDCAUSE"
    'sql = sql & "  ,NVL(HLD.HLDCMNT,' ') as HLDCMNT"
    ' ����ΰ��ނ𗬓��Ď��f�[�^�֒u������ del SETkimizuka End  09/03/17
    sql = sql & "  ,NVL(PASS.REGDATE,SYSDATE) as ENDDATE"
    sql = sql & "  ,WFS.PLANTCATCB as PLANTCAT"                         ''���� 2007/09/04 SPK Tsutsumi Add
    ' ����ΰ��ނ𗬓��Ď��f�[�^�֒u������ add SETkimizuka Start  09/03/17
    ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/30
    sql = sql & " , NVL(TO_CHAR(Y4.AGRSTATUS),' ') as AGRSTATUSY4"
    sql = sql & " , NVL(TO_CHAR(Y4.STOP),'0') as STOP "
    sql = sql & " , NVL(Y4.CAUSE,' ') as CAUSEY4"
    sql = sql & " , NVL(Y4.PRINTKIND || Y4.PRINTNO,' ') as PRINTNOY4"
    sql = sql & " , NVL(Y4.WKKTY4,'0') as WKKTY4 "
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/06/30
    sql = sql & "  ,NVL(Y4_DISPHLD.STOPY4,'0') as HLDCLASS"
    sql = sql & "  ,NVL(Y4_DISPHLD.SETTDAYY4,SYSDATE) as HLDDATE"
    sql = sql & "  ,NVL(Y4_DISPHLD.STAFFNAME,' ') as HLDSTAFFNAME"
    sql = sql & "  ,NVL(Y4_DISPHLD.CAUSEY4,' ') as HLDCAUSE"
    '����EDI����ݸ�Ή� 2009/12/4 Add Strat SPK habuki������
    sql = sql & "  ,case NVL(EDI.EDIFLG,'@')"
    sql = sql & "     when '2' then 'OK'"
    sql = sql & "     when '1' then 'NG'"
    sql = sql & "     when '0' then 'NG'"
    '2010/1/22 Null�����b�N����悤���C SPK Hitomi
    sql = sql & "     when '@' then 'NG'"
    sql = sql & "     else '  '"
    sql = sql & "   end  as EDIFLG "
    '����EDI����ݸ�Ή� 2009/12/4 Add End   SPK habuki������
    ' ����ΰ��ނ𗬓��Ď��f�[�^�֒u������ add SETkimizuka End  09/03/17
    sql = sql & " FROM"
    sql = sql & "   XSDCB WFS"
    ' ����ΰ��ނ𗬓��Ď��f�[�^�֒u������ del SETkimizuka Start  09/03/17
    'sql = sql & "  ,("
    'sql = sql & "    SELECT"
    'sql = sql & "      DAT.SNGLID"
    'sql = sql & "     ,DAT.HLDCLASS"
    'sql = sql & "     ,DAT.REGDATE"
    'sql = sql & "     ,DAT.TSTAFFID"
    'sql = sql & "     ,DAT.HLDCAUSE"
    'sql = sql & "     ,DAT.HLDCMNT"
    'sql = sql & "     ,rtrim(STAFF.JFMLNAME)||rtrim(STAFF.JFSTNAME) as STAFFNAME"
    'sql = sql & "    FROM"
    'sql = sql & "      TBCMW008 DAT"
    'sql = sql & "     ,("
    'sql = sql & "       SELECT"
    'sql = sql & "         SNGLID"
    'sql = sql & "        ,MAX(TRANCNT) as MAX_TRANCNT"
    'sql = sql & "       FROM"
    'sql = sql & "         TBCMW008"
    'sql = sql & "       GROUP BY"
    'sql = sql & "         SNGLID"
    'sql = sql & "      ) W008"
    'sql = sql & "     ,TBCMB001 STAFF"
    'sql = sql & "    WHERE DAT.SNGLID   = W008.SNGLID"
    'sql = sql & "      AND DAT.TRANCNT  = W008.MAX_TRANCNT"
    'sql = sql & "      AND DAT.TSTAFFID = STAFF.STAFFID"
    'sql = sql & "   ) HLD"
    ' ����ΰ��ނ𗬓��Ď��f�[�^�֒u������ del SETkimizuka End  09/03/17
    ' �������x���P�Ή�  Add Y.Hitomi 09/12/4
    sql = sql & "  ,("
    sql = sql & "    SELECT"
    sql = sql & "      *"
    sql = sql & "    FROM"
    sql = sql & "    ("
    sql = sql & "        SELECT"
    sql = sql & "             SXLID,"
    sql = sql & "             TRANCNT,"
    sql = sql & "             REGDATE,"
    sql = sql & "             rank() over(partition by SXLID order by TRANCNT desc ) as RANK"
    sql = sql & "         FROM"
    sql = sql & "             TBCMW005"
    sql = sql & "    )"
    sql = sql & "    WHERE"
    sql = sql & "         RANK = 1"
    sql = sql & "   ) PASS"
'�������x���P�Ή�  del Y.Hitomi 09/12/4
'    sql = sql & "    SELECT"
'    sql = sql & "      DAT.SXLID"
'    sql = sql & "     ,DAT.REGDATE"
'    sql = sql & "    FROM"
'    sql = sql & "      TBCMW005 DAT"
'    sql = sql & "     ,("
'    sql = sql & "       SELECT"
'    sql = sql & "         SXLID"
'    sql = sql & "        ,max(TRANCNT) as MAX_TRANCNT"
'    sql = sql & "       FROM"
'    sql = sql & "         TBCMW005"
'    sql = sql & "       GROUP BY"
'    sql = sql & "         SXLID"
'    sql = sql & "      ) W005"
'    sql = sql & "    WHERE DAT.SXLID    = W005.SXLID"
'    sql = sql & "      AND DAT.TRANCNT  = W005.MAX_TRANCNT"
'    sql = sql & "   ) PASS"

    ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/30
    ' ������~���ڒǉ� add SETkimizuka Start  09/03/17
'    sql = sql & "    ,( "
'    sql = sql & "      SELECT SXLIDY3 as SXLIDY4 ,MIN(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)) as AGRSTATUS  "
'    sql = sql & "      ,MAX(STOPY4) as STOP,DECODE(TRIM(CAUSEY4),'',TRIM(CAUSEY4),TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSE ,Y5.PRINTNO,Y5.PRINTKIND "
'    sql = sql & "      FROM XSDCB,XODY3  "
'    sql = sql & "           LEFT OUTER JOIN XODY4 on ( XTALNOY3 = XTALNOY4 AND STOPY4 <> '2' AND LIVKY4 = '0' AND WKKTY4 ='CW000')"
''    sql = sql & "           LEFT OUTER JOIN XODY4 on ( XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND STOPY4 <> '2' AND LIVKY4 = '0' AND WKKTY4 in " & CreateWkktSQL(WATCH_PROCCD) & ") "
'    sql = sql & "           LEFT OUTER JOIN KODA9 on ( SYSCA9 = 'X' AND SHUCA9 = '30' AND CAUSEY4 = CODEA9 ) "
'    sql = sql & "           LEFT OUTER JOIN (SELECT XTALNOY4 as XTALNO,SXLIDY4 as SXLID,PRINTNOY5 as PRINTNO,PRINTKINDY5 as PRINTKIND "
'    sql = sql & "                FROM XODY3,XODY4,XODY5 "
'    sql = sql & "              WHERE XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND LIVKY3 = '0' "
'    sql = sql & "                AND PRINTKINDY4 = PRINTKINDY5 AND PRINTNOY4 = PRINTNOY5  "
'    sql = sql & "                AND HKBNY5 ='0' GROUP BY XTALNOY4,SXLIDY4,PRINTNOY5,PRINTKINDY5) Y5 ON (XTALNOY3 = XTALNO AND SXLIDY3 = SXLID ) "
'    sql = sql & "      WHERE LIVKCB     <>'1' AND SXLIDCB = SXLIDY3 AND GNWKNTCB    = 'CW800' AND  "
'    sql = sql & "       LIVKY3    = '0'  "
'    sql = sql & "       GROUP BY SXLIDY3,AGRSTATUSY4,CAUSEY4,Y5.PRINTNO,Y5.PRINTKIND,NAMEJA9 "
'
'    sql = sql & "      UNION SELECT SXLIDY3 as SXLIDY4 ,MIN(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)) as AGRSTATUS  "
'    sql = sql & "      ,MAX(STOPY4) as STOP,DECODE(TRIM(CAUSEY4),'',TRIM(CAUSEY4),TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSE ,Y5.PRINTNO,Y5.PRINTKIND "
'    sql = sql & "      FROM XSDCB,XODY3  "
'    sql = sql & "           LEFT OUTER JOIN XODY4 on ( XTALNOY3 = XTALNOY4 AND SXLIDY3 = SXLIDY4 AND STOPY4 <> '2' AND LIVKY4 = '0' AND WKKTY4 = 'CW800')"
'    sql = sql & "           LEFT OUTER JOIN KODA9 on ( SYSCA9 = 'X' AND SHUCA9 = '30' AND CAUSEY4 = CODEA9 ) "
'    sql = sql & "           LEFT OUTER JOIN (SELECT XTALNOY4 as XTALNO,SXLIDY4 as SXLID,PRINTNOY5 as PRINTNO,PRINTKINDY5 as PRINTKIND "
'    sql = sql & "                FROM XODY3,XODY4,XODY5 "
'    sql = sql & "              WHERE XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND LIVKY3 = '0' "
'    sql = sql & "                AND PRINTKINDY4 = PRINTKINDY5 AND PRINTNOY4 = PRINTNOY5  "
'    sql = sql & "                AND HKBNY5 ='0' GROUP BY XTALNOY4,SXLIDY4,PRINTNOY5,PRINTKINDY5) Y5 ON (XTALNOY3 = XTALNO AND SXLIDY3 = SXLID ) "
'    sql = sql & "      WHERE LIVKCB     <>'1' AND SXLIDCB = SXLIDY3 AND GNWKNTCB    = 'CW800' AND  "
'    sql = sql & "       LIVKY3    = '0'  "
'    sql = sql & "       GROUP BY SXLIDY3,AGRSTATUSY4,CAUSEY4,Y5.PRINTNO,Y5.PRINTKIND,NAMEJA9 "
'    sql = sql & "       ) Y4 "
    sql = sql & "    ,( "
    sql = sql & "      SELECT SXLIDY3 as SXLIDY4 ,DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4) as AGRSTATUS  "
    sql = sql & "      ,STOPY4 as STOP,DECODE(TRIM(CAUSEY4),'',TRIM(CAUSEY4),TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSE ,Y4.PRINTNOY4 as PRINTNO,Y4.PRINTKINDY4 as PRINTKIND"
    sql = sql & "      ,WKKTY4 "
    sql = sql & "      FROM XSDCB,XODY3 Y3,XODY4 Y4,KODA9 A9 "
    sql = sql & "      WHERE LIVKCB     <>'1' AND SXLIDCB = Y3.SXLIDY3 AND GNWKNTCB    = 'CW800' "
    sql = sql & "       AND Y3.XTALNOY3 = Y4.XTALNOY4(+) "
    sql = sql & "       AND Y3.SXLIDY3 = Y4.SXLIDY4(+) "
    sql = sql & "       AND Y3.LIVKY3(+) = '0' "
    sql = sql & "       AND Y4.LIVKY4(+) = '0' "
    sql = sql & "       AND A9.SYSCA9(+) = 'X' AND A9.SHUCA9(+) = '30' AND Y4.CAUSEY4 = A9.CODEA9(+) "
    sql = sql & " UNION SELECT SXLIDY3 as SXLIDY4 ,DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4) as AGRSTATUS  "
    sql = sql & "      ,STOPY4 as STOP,DECODE(TRIM(CAUSEY4),'',TRIM(CAUSEY4),TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSE ,Y4.PRINTNOY4 as PRINTNO,Y4.PRINTKINDY4 as PRINTKIND"
    sql = sql & "      ,WKKTY4 "
    sql = sql & "      FROM XSDCB,XODY3 Y3,XODY4 Y4,KODA9 A9 "
    sql = sql & "      WHERE LIVKCB     <>'1' AND SXLIDCB = Y3.SXLIDY3 AND GNWKNTCB    = 'CW800' "
    sql = sql & "       AND Y3.XTALNOY3 = Y4.XTALNOY4(+) "
    sql = sql & "       AND Y4.WKKTY4(+) = 'CW000' "
    sql = sql & "       AND Y3.LIVKY3(+) = '0' "
    sql = sql & "       AND Y4.LIVKY4(+) = '0' "
    sql = sql & "       AND A9.SYSCA9(+) = 'X' AND A9.SHUCA9(+) = '30' AND Y4.CAUSEY4 = A9.CODEA9(+) "
    sql = sql & "       ) Y4 "
    ' ������~���ڒǉ� add SETkimizuka End  09/03/17
    sql = sql & "    ,(SELECT SXLIDY4,SETTDAYY4,NAMEJA9 as STAFFNAME,CAUSEY4,STOPY4 FROM XODY4,KODA9 "
    sql = sql & "       WHERE STOPY4 <> '2' AND WKKTY4 = '" & DISP_HOLD & "'"
    sql = sql & "         AND SYSCA9(+) = 'K' AND SHUCA9(+) = '55' AND CODEA9(+) = SETSTAFFIDY4 ) Y4_DISPHLD "
    ' �����Ď�SQL�C�� upd SETkimizuka End  09/06/30
    
    '����EDI����ݸ�Ή� 2009/12/4 Add Strat SPK habuki������
    '/* EDI�׸ނ̐ݒ�󋵊m�F */
    sql = sql & "    ,( "
    sql = sql & "      select "
    sql = sql & "          substr(c61.XTALC6,1,7) XTAL"
    sql = sql & "         ,max(c61.EDIFLGC6)      EDIFLG"
    sql = sql & "      from "
    sql = sql & "          XODC6_1  c61"
    sql = sql & "         ,TBCMH001 t01"
    sql = sql & "      where"
    sql = sql & "           substr(c61.XTALC6,1,7)||substr(c61.XTALC6,9,1) = substr(t01.UPINDNO,1,7)||substr(t01.UPINDNO,9,1)"
    sql = sql & "       and t01.CODE > '4'"
    sql = sql & "      group by"
    sql = sql & "         substr(c61.XTALC6,1,7)"
    sql = sql & "     ) EDI "
    '����EDI����ݸ�Ή� 2009/12/4 Add End   SPK habuki������

    sql = sql & " WHERE WFS.LIVKCB     <>'1'"
    sql = sql & "   AND WFS.GNWKNTCB    = 'CW800'"
    'sql = sql & "   AND WFS.SXLIDCB     = HLD.SNGLID(+)"           'del 09/03/17 SETkimizuka
    sql = sql & "   AND WFS.SXLIDCB     = PASS.SXLID(+)"
    sql = sql & "   AND WFS.SXLIDCB     = Y4.SXLIDY4(+)"            'add 09/03/17 SETkimizuka
    sql = sql & "   AND WFS.SXLIDCB     = Y4_DISPHLD.SXLIDY4(+)"    'add 09/03/17 SETkimizuka
    
    '����EDI����ݸ�Ή� 2009/12/4 Add Strat SPK habuki������
    sql = sql & "   AND substr(WFS.SXLIDCB,1,7) = EDI.XTAL(+)"
    '����EDI����ݸ�Ή� 2009/12/4 Add End   SPK habuki������
    
    ' ���� 2007/09/04 SPK Tsutsumi Add Start
    If sCmbMukesaki <> "ALL" Then
        sql = sql & "   AND WFS.PLANTCATCB      = '" & sCmbMukesaki & "'"
    End If
    ' 2007/09/04 SPK Tsutsumi Add End
    
    ' ����ΰ��ނ𗬓��Ď��f�[�^�֒u������ del SETkimizuka Start  09/03/17
    sql = sql & "   ORDER BY NVL(Y4_DISPHLD.STOPY4,0),WFS.SXLIDCB"
    'sql = sql & "   ORDER BY HLD.HLDCLASS,WFS.SXLIDCB"
    ' ����ΰ��ނ𗬓��Ď��f�[�^�֒u������ del SETkimizuka End  09/03/17
''���ǉ�END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
''���폜START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
'    sql = sql & "select  SXL.CRYNUM, SXL.INGOTPOS, SXL.SXLID, SXL.HINBAN, SXL.LENGTH, WFS.MAICB as COUNT , "
'    sql = sql & "  nvl(HLD.HLDCLASS,'0') as HLDCLASS, nvl(HLD.REGDATE,SYSDATE) as HLDDATE, "
'    sql = sql & "  nvl(HLD.STAFFNAME,' ') as HLDSTAFFNAME, nvl(HLD.HLDCAUSE,' ') as HLDCAUSE, "
'    sql = sql & "  nvl(HLD.HLDCMNT,' ') as HLDCMNT,  nvl(PASS.REGDATE,SYSDATE) as ENDDATE  "
'    sql = sql & "from TBCME042 SXL,XSDCB WFS, "
'    sql = sql & "  (select DAT.SNGLID, DAT.HLDCLASS, DAT.REGDATE, DAT.TSTAFFID, DAT.HLDCAUSE, DAT.HLDCMNT, "
'    sql = sql & "     rtrim(STAFF.JFMLNAME)||rtrim(STAFF.JFSTNAME) as STAFFNAME "
'    sql = sql & "   from TBCMW008 DAT,  "
'    sql = sql & "     (select SNGLID, MAX(TRANCNT) as MAX_TRANCNT from TBCMW008 group by SNGLID) W008, "
'    sql = sql & "     TBCMB001 STAFF "
'    sql = sql & "   Where (DAT.SNGLID = W008.SNGLID) and (DAT.TRANCNT = W008.MAX_TRANCNT) and (DAT.TSTAFFID = STAFF.STAFFID) "
'    sql = sql & "  ) HLD, "
'    sql = sql & "  (select  DAT.SXLID, DAT.REGDATE "
'    sql = sql & "   from  TBCMW005 DAT, "
'    sql = sql & "     (select SXLID, max(TRANCNT) as MAX_TRANCNT from TBCMW005 group by SXLID) W005 "
'    sql = sql & "   where (DAT.SXLID = W005.SXLID) and (DAT.TRANCNT = W005.MAX_TRANCNT) "
'    sql = sql & "  ) PASS "
'    sql = sql & " where (SXL.DELCLS<>'1') "
'    sql = sql & " and (SXL.NOWPROC='CW800') "
'    sql = sql & " and (SXL.SXLID = HLD.SNGLID(+)) "
'    sql = sql & " and (SXL.SXLID = PASS.SXLID(+)) "
'    sql = sql & " and (SXL.SXLID = WFS.SXLIDCB(+)) "
''���폜END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
    
    ReDim records(0)
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'If rs.RecordCount = 0 Then
    '    DBDRV_fcmlc001e_Disp = FUNCTION_RETURN_FAILURE
    '    rs.Close
    '    GoTo PROC_EXIT
    'End If
    
    'ReDim records(rs.RecordCount)  'del 09/03/17 SETkimizuka
    iCnt = 0
    For i = 1 To rs.RecordCount
        If sOldID <> rs("SXLID") Then  'add 09/03/17 SETkimizuka
            iCnt = iCnt + 1        'add 09/03/17 SETkimizuka
            ReDim Preserve records(iCnt)
            With records(iCnt)
                ' SXL�Ǘ�
                .CRYNUM = rs("CRYNUM")          ' �����ԍ�
                .INGOTPOS = rs("INGOTPOS")      ' �������ʒu
                .SXLID = rs("SXLID")            ' ��SXLID
                .hinban = rs("HINBAN")          ' �i��
                .LENGTH = rs("LENGTH")          ' ����
                .COUNT = rs("COUNT")            ' �\�薇��
                .ENDDATE = rs("ENDDATE")        ' �������t�iSXL�Ǘ�.�o�^���t�j
                
                ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/30
                '.AGRSTATUS = rs("AGRSTATUSY4")               ' ������~�敪  add 09/03/17 SETkimizuka
                '.STOP = rs("STOP")               ' ������~�敪  add 09/03/17 SETkimizuka
                'If Trim(rs("CAUSEY4")) <> "" Then
                '    .CAUSE = rs("CAUSEY4") & vbTab      ' ������~���R  add 09/03/17 SETkimizuka
                'End If
                If rs("STOP") <> "2" And (rs("WKKTY4") = "CW800" Or rs("WKKTY4") = "CW000") Then
                    .AGRSTATUS = rs("AGRSTATUSY4")               ' ������~�敪
                    .STOP = rs("STOP")               ' ������~�敪
                    If Trim(rs("CAUSEY4")) <> "" Then
                        .CAUSE = rs("CAUSEY4") & vbTab      ' ������~���R  add 09/03/17 SETkimizuka
                    End If
                Else
                    .STOP = "0"               ' ������~�敪
                End If
                ' �����Ď�SQL�C�� upd SETkimizuka End  09/06/30
                
                If Trim(rs("PRINTNOY4")) <> "" Then
                    .PRINTNO = rs("PRINTNOY4") & vbTab  ' ��s�]��No    add 09/03/17 SETkimizuka
                End If
                
                ' ����ΰ��ނ𗬓��Ď��f�[�^�֒u������ add SETkimizuka Start  09/03/17
                .HLDCLASSOLD = rs("HLDCLASS")
                .HLDCLASS = rs("HLDCLASS")
                .HLDDATE = rs("HLDDATE")
                .HLDCAUSE = rs("HLDCAUSE")
                .HLDSTAFFNAME = rs("HLDSTAFFNAME")
                ' ����ΰ��ނ𗬓��Ď��f�[�^�֒u������ add SETkimizuka End  09/03/17
                
                '����EDI����ݸ�Ή� 2009/12/4 Add Strat SPK habuki������
                .EDIFLG = rs("EDIFLG")
                '����EDI����ݸ�Ή� 2009/12/4 Add End   SPK habuki������
                
                ' ����ΰ��ނ𗬓��Ď��f�[�^�֒u������ del SETkimizuka Start  09/03/17
                '' WF�z�[���h�i�����j����
                'If (CStr(rs("HLDCLASS")) = "1") Then
                '    .HLDCLASSOLD = rs("HLDCLASS")       ' ���z�[���h�����敪 (0:�z�[���h���������A1:�z�[���h����)
                '    .HLDCLASS = rs("HLDCLASS")          ' �z�[���h�����敪 (0:�z�[���h���������A1:�z�[���h����)
                '    .HLDDATE = rs("HLDDATE")            ' �z�[���h���t(.�o�^���t)
                '    .HLDSTAFFNAME = rs("HLDSTAFFNAME")  ' �z�[���h�S���Җ�(.�o�^�Ј�ID)
                '    .HLDCAUSE = rs("HLDCAUSE")          ' �z�[���h���R ('SC','17')
                '    .HLDCMNT = rs("HLDCMNT")            ' �z�[���h�R�����g
                'Else
                '
                '    'WFΰ��ވȊO�̕\�������@2004/07/15 koyama
                '    If DBDRV_s_cmbc040_SQL_Y019XSDCB(rs("SXLID"), rs("CRYNUM"), rs("HINBAN"), _
                                                     rs("INGOTPOS"), sWFHOLDDATE, sUSER_ID _
                                                    ) = FUNCTION_RETURN_FAILURE Then
    
    '                   .HLDCLASSOLD = vbNullString         ' ���z�[���h�����敪 (0:�z�[���h���������A1:�z�[���h����)
                '        .HLDCLASSOLD = "0"                  ' ���z�[���h�����敪 (0:�z�[���h���������A1:�z�[���h����)
                '        .HLDCLASS = vbNullString            ' �z�[���h�����敪 (0:�z�[���h���������A1:�z�[���h����)
    '                    .HLDDATE = " "                      ' �z�[���h������
                '        .HLDSTAFFNAME = vbNullString        ' �z�[���h�S���Җ�(.�o�^�Ј�ID)
                '        .HLDCAUSE = vbNullString            ' �z�[���h���R ('SC','17')
                '        .HLDCMNT = vbNullString             ' �z�[���h�R�����g
    
                '    Else
                '
                '        If sWFHOLDDATE <> "" And IsNull(sWFHOLDDATE) = False Then
                '            .HLDDATE = sWFHOLDDATE      ' �z�[���h���t(.�o�^���t)
                '       End If
                '        .HLDSTAFFNAME = sUSER_ID    ' �z�[���h�S���Җ�(.�o�^�Ј�ID)
                '
                '        '2004/07/21 koyama
                '        .HLDCLASSOLD = "0"                  ' ���z�[���h�����敪 (0:�z�[���h���������A1:�z�[���h����)
                '        .HLDCLASS = vbNullString            ' �z�[���h�����敪 (0:�z�[���h���������A1:�z�[���h����)
                '        .HLDCAUSE = vbNullString            ' �z�[���h���R ('SC','17')
                '       .HLDCMNT = vbNullString             ' �z�[���h�R�����g
                '    End If
                '
                'End If
                ' ����ΰ��ނ𗬓��Ď��f�[�^�֒u������ del SETkimizuka End  09/03/17
                
                ' 2007/09/13 SPK Tsutsumi Add Start
                If IsNull(rs("PLANTCAT")) = False Then
                    For j = 0 To UBound(s_MukesakiBase)
                        If s_MukesakiBase(j).sMukeCode = rs("PLANTCAT") Then
                            .MUKESAKI = s_MukesakiBase(j).sMukeName
                        End If
                    Next j
                End If
                ' 2007/09/13 SPK Tsutsumi Add End
                
                sOldID = rs("SXLID")   'add 09/03/17 SETkimizuka
                rs.MoveNext
            End With
        Else
            ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/30
            '' ������~���R  add 09/03/17 SETkimizuka
            'If InStr(records(iCnt).CAUSE, rs("CAUSEY4")) = 0 And Trim(rs("CAUSEY4")) <> "" Then
            '    records(iCnt).CAUSE = records(iCnt).CAUSE & rs("CAUSEY4") & vbTab
            'End If
            If rs("STOP") <> "2" And (rs("WKKTY4") = "CW800" Or rs("WKKTY4") = "CW000") Then
                If Trim(records(iCnt).AGRSTATUS) = "" Or (rs("AGRSTATUSY4") < records(iCnt).AGRSTATUS) Then
                    records(iCnt).AGRSTATUS = rs("AGRSTATUSY4")               ' ���F�m�F�敪
                    records(iCnt).STOP = rs("STOP")               ' ������~�敪
                End If
                If InStr(records(iCnt).CAUSE, rs("CAUSEY4")) = 0 And Trim(rs("CAUSEY4")) <> "" Then
                    records(iCnt).CAUSE = records(iCnt).CAUSE & rs("CAUSEY4") & vbTab
                End If
            End If
            ' �����Ď�SQL�C�� upd SETkimizuka Start  09/06/30
            ' ��s�]��No    add 09/03/17 SETkimizuka
            If InStr(records(iCnt).PRINTNO, rs("PRINTNOY4")) = 0 And Trim(rs("PRINTNOY4")) <> "" Then
                records(iCnt).PRINTNO = records(iCnt).PRINTNO & rs("PRINTNOY4") & vbTab
            End If
            rs.MoveNext
        End If
    Next
    rs.Close
    
    DBDRV_fcmlc001e_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_fcmlc001e_Disp = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'Public Function DBDRV_fcmlc001e_Exec(records() As typ_cmlc001e_Disp, StaffID$, errTbl$) As FUNCTION_RETURN
'records()��record�@06/10/20 ooba
Public Function DBDRV_fcmlc001e_Exec(record As typ_cmlc001e_Disp, StaffID$, errTbl$) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset
Dim i As Integer
Dim smpId(2) As String
Dim blkID As String
Dim errmsg As String
Dim typXODY3()  As typ_XODY3    'add 09/03/17 SETkimizuka
Dim typXODY4()  As typ_XODY4    'add 09/03/17 SETkimizuka
Dim tXODY4      As typ_XODY4    'add 09/03/17 SETkimizuka
Dim IRow        As Integer      'add 09/03/17 SETkimizuka

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function DBDRV_fcmlc001e_Exec"

    ''��ʂ̍s���ɏ���
'    For i = 1 To UBound(records)
'        With records(i)
        With record     '06/10/20 ooba
        
''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
    ''XSDCB.�z�[���h�敪�̍X�V���s�Ȃ��B
            sql = " UPDATE XSDCB SET" & _
                  "  KDAYCB = SYSDATE"
            If .HLDCLASS = "0" Then
                sql = sql & " ,SHOLDCLSCB = '0' "
            Else
                sql = sql & " ,SHOLDCLSCB = '1' "
            End If
            sql = sql & " WHERE SXLIDCB = '" & .SXLID & "'"
            If 0 >= OraDB.ExecuteSQL(sql) Then
                DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
                errTbl = "XSDCB"
                GoTo proc_exit
            End If
''���ǉ�END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
''���폜START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
'            ''SXL�Ǘ����̍X�V
'            sql = "update TBCME042 set " & _
'                  "KRPROCCD = '" & MGPRCD_SXL_KAKUTEI & "', " & _
'                  "NOWPROC = '" & PROCD_SXL_KAKUTEI & "', " & _
'                  "LPKRPROCCD = '" & MGPRCD_SXL_KAKUTEI & "', " & _
'                  "LASTPASS = '" & PROCD_SXL_KAKUTEI & "', " & _
'                  "UPDDATE = SYSDATE, "
'            If .HLDCLASS = "0" Then     'SXL�m��
'                sql = sql & "HOLDCLS = '0', " & _
'                    "LSTATCLS='S', " & _
'                    "SENDFLAG='3', " & _
'                    "DELCLS='1' "
'            Else                        'WF�z�[���h
'                sql = sql & "HOLDCLS = '1', " & _
'                    "SENDFLAG = '0' "
'            End If
'            sql = sql & "where (SXLID='" & .SXLID & "')"
'            If 0 >= OraDB.ExecuteSQL(sql) Then
'                DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
'                errTbl = "E042"
'                GoTo proc_exit
'            End If
''���폜END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
            
            
            ' ����ΰ��ނ𗬓��Ď��f�[�^�֒u������ del SETkimizuka Start  09/03/17
            '''WF�z�[���h���т̒ǉ�
            'If .HLDCLASS <> .HLDCLASSOLD Then
            '    sql = "insert into TBCMW008 (" & _
            '          " CRYNUM, INGOTPOS, TRANCNT," & _
            '          " CRYLEN, KRPROCCD, PROCCODE," & _
            '          " SNGLID, HLDCLASS, HLDCAUSE, HLDCMNT," & _
            '          " TSTAFFID , REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE" & _
            '          ") select " & _
            '          "'" & .CRYNUM & "', " & .INGOTPOS & ", nvl(max(TRANCNT),0)+1, " & _
            '          .LENGTH & ", '" & MGPRCD_SXL_KAKUTEI & "', '" & PROCD_SXL_KAKUTEI & "', " & _
            '          "'" & .SXLID & "', '" & .HLDCLASS & "', " & NoNullStr(.HLDCAUSE) & ", " & NoNullStr(.HLDCMNT) & ", " & _
            '          "'" & STAFFID & "', SYSDATE, ' ', SYSDATE, '0', SYSDATE " & _
            '          "From TBCMW008 " & _
            '          "where (CRYNUM='" & .CRYNUM & "') and (INGOTPOS=" & .INGOTPOS & ") "
            '    If 0 >= OraDB.ExecuteSQL(sql) Then
            '        DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
            '        errTbl = "W008"
            '        GoTo proc_exit
            '    End If
            'End If
            ' ����ΰ��ނ𗬓��Ď��f�[�^�֒u������ del SETkimizuka End  09/03/17
            
            ' ����ΰ��ނ𗬓��Ď��f�[�^�֒u������ add SETkimizuka Start  09/03/17
            If .HLDCLASS <> .HLDCLASSOLD Then
                Call GetSysdate
                ReDim typXODY4(1)
                If .HLDCLASS = "1" Then
                    If GetXODY3(typXODY3, "WHERE SXLIDY3 ='" & .SXLID & "' AND LIVKY3 = '0' ") = True Then
                        For IRow = 1 To UBound(typXODY3)
                            typXODY4(1).AGRSTATUSY4 = VB_KBN
                            typXODY4(1).CAUSEY4 = .HLDCAUSE
                            typXODY4(1).LIVKY4 = "0"
                            typXODY4(1).SETSTAFFIDY4 = StaffID
                            typXODY4(1).WKKTY4 = DISP_HOLD
                            typXODY4(1).STOPY4 = STOP_KBN
                            typXODY4(1).XTALNOY4 = typXODY3(IRow).XTALNOY3
                            typXODY4(1).RCNTY4 = typXODY3(IRow).RCNTY3
                            typXODY4(1).SXLIDY4 = typXODY3(IRow).SXLIDY3
                            If InsertXODY4(typXODY4) = False Then
                                DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
                                errTbl = "XODY4"
                                GoTo proc_exit
                            End If
                        Next
                    End If
                Else
                    tXODY4.STOPY4 = KAIJO_KBN
                    tXODY4.KSTAFFIDY4 = StaffID
                    tXODY4.KDAYY4 = gsSysdate
                    If UpdateXODY4(tXODY4, _
                        "WHERE SXLIDY4 ='" & .SXLID & "' AND STOPY4 = '" & STOP_KBN & "' " _
                        & "AND WKKTY4 = '" & DISP_HOLD & "'") = False Then
                        DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
                        errTbl = "XODY4"
                        GoTo proc_exit
                    End If
                End If
                
            End If
            ' ����ΰ��ނ𗬓��Ď��f�[�^�֒u������ add SETkimizuka End  09/03/17
            
            ''SXL�m����т̒ǉ�
            If .HLDCLASS = "0" Then
                smpId(1) = vbNullString
                smpId(2) = vbNullString
                blkID = vbNullString
                
                '�T���v��ID�𓾂�
 '               sql = "select E044SMPLID from VECME011 where (E042SXLID='" & .SXLID & "') order by E044INGOTPOS"
 '�@�@�@�@�@�@�@ �T���v���Ǘ��Ƃ��ăT���v���w�����S�ĂȂ����̊m��X�̂��̂͏���
                ''�@�u���b�NID�Z�b�g 2003/09/17 Motegi ===========================================> START
                '' �T���v��ID(From)�A�T���v��ID(To)��������Ȃ���΁A�u���b�NID���Z�b�g���A
                '' �T���v��ID(From)�A�T���v��ID(To)��������΁A�u���b�NID���Z�b�g���Ȃ��B
 
'                sql = "select E044SMPLID from VECME011 where (E042SXLID='" & .SXLID & "' and E044KTKBN != '9') order by E044INGOTPOS"
'                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'                If rs.RecordCount = 2 Then
'                    smpId(1) = rs("E044SMPLID")
'                    rs.MoveNext
'                    smpId(2) = rs("E044SMPLID")
'                    rs.Close
'                    Set rs = Nothing
'                Else
'                    rs.Close
'
'                    '�u���b�NID�𓾂�
'                    sql = "select BLOCKID from TBCME040 " & _
'                          "where (crynum='" & .CRYNUM & "') and (INGOTPOS<=" & .INGOTPOS & ") and (" & .INGOTPOS & "<INGOTPOS+LENGTH)"
'                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'                    If rs.RecordCount = 1 Then
'                        blkID = rs("BLOCKID")
'                    End If
'                    rs.Close
'                    Set rs = Nothing
'                End If
                ''-----------------------------------------------------------------------------
                sql = "select REPSMPLIDCW from XSDCW where SXLIDCW='" & .SXLID & "' and KTKBNCW != '9' order by INPOSCW"
                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 2 Then
                    smpId(1) = rs("REPSMPLIDCW")
                    rs.MoveNext
                    smpId(2) = rs("REPSMPLIDCW")
                    rs.Close
                    Set rs = Nothing
                
                    '���L�T���v���`�F�b�N����
                    If chkComSAMPL(.SXLID, smpId(1), smpId(1)) Then
                        DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
                        errTbl = "DBDRV_fcmlc001e_Exec:���L�����ID�擾(From)"
                        GoTo proc_exit
                    End If
                    If chkComSAMPL(.SXLID, smpId(2), smpId(2)) Then
                        DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
                        errTbl = "DBDRV_fcmlc001e_Exec:���L�����ID�擾(To)"
                        GoTo proc_exit
                    End If
                Else
                    rs.Close

                    '�u���b�NID�𓾂�
                    sql = "select BLOCKID from TBCME040 " & _
                          "where (crynum='" & .CRYNUM & "') and (INGOTPOS<=" & .INGOTPOS & ") and (" & .INGOTPOS & "<INGOTPOS+LENGTH)"
                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                    If rs.RecordCount = 1 Then
                        blkID = rs("BLOCKID")
                    End If
                    rs.Close
                    Set rs = Nothing
                End If
                ''�@�u���b�NID�Z�b�g 2003/09/17 Motegi ===========================================> END
                
'                '�m����т���������
' 2007/09/04 SPK Tsutsumi Add Start
                sql = "insert into TBCMW007 (" & _
                      "CRYNUM, INGOTPOS, " & _
                      "CRYLEN, KRPROCCD, PROCCODE, " & _
                      "SXLID, SAMPLE_FROM, SAMPLE_TO, BLOCKID, " & _
                      "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE" & _
                      ",PLANTCAT  " & _
                      ") values (" & _
                      NoNullStr(.CRYNUM) & ", " & .INGOTPOS & ", " & _
                      .LENGTH & ", '" & MGPRCD_SXL_KAKUTEI & "', '" & PROCD_SXL_KAKUTEI & "', " & _
                      NoNullStr(.SXLID) & ", " & NoNullStr(smpId(1)) & ", " & NoNullStr(smpId(2)) & ", " & NoNullStr(blkID) & ", " & _
                      NoNullStr(StaffID) & ", SYSDATE, ' ', SYSDATE, '0', SYSDATE, " & sCmbMukesaki & " " & _
                      ")"
'                sql = "insert into TBCMW007 (" & _
'                      "CRYNUM, INGOTPOS, " & _
'                      "CRYLEN, KRPROCCD, PROCCODE, " & _
'                      "SXLID, SAMPLE_FROM, SAMPLE_TO, BLOCKID, " & _
'                      "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE" & _
'                      ",PLANTCAT  " & _
'                      ") values (" & _
'                      NoNullStr(.CRYNUM) & ", " & .INGOTPOS & ", " & _
'                      .LENGTH & ", '" & MGPRCD_SXL_KAKUTEI & "', '" & PROCD_SXL_KAKUTEI & "', " & _
'                      NoNullStr(.SXLID) & ", " & NoNullStr(smpId(1)) & ", " & NoNullStr(smpId(2)) & ", " & NoNullStr(BlkId) & ", " & _
'                      NoNullStr(StaffID) & ", SYSDATE, ' ', SYSDATE, '0', SYSDATE " & _
'                      ")"
' 2007/09/04 SPK Tsutsumi Add End
                If 0 >= OraDB.ExecuteSQL(sql) Then
                    DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
                    errTbl = "W007"
                    GoTo proc_exit
                End If
                
                'SXL�������Ƃ��̊֘A�f�[�^����������
                If WriteX00n(.SXLID, .COUNT, errmsg) = FUNCTION_RETURN_FAILURE Then
                    DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
                    errTbl = errmsg
                    GoTo proc_exit
                End If
            End If
        End With
'    Next
        
    
    DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�S�ʕύX 2003/10/17 SystemBrain
'����    :2008/03/31 �� �@���sSXL�m����s������SXL�}�b�v��M���ɑ��M���ݸޕύX����B
'                         �A�T���v��ID�𔽉f���̃T���v��ID(�����T���v��ID�܂�)�ł͂Ȃ��A��\�T���v��ID�ɕύX����B
'                         �BGB7/GB8/GB9��SXL�m����t����v������B
Private Function WriteX00n(ByVal SXLID$, ByVal WfCnt%, errmsg$) As FUNCTION_RETURN
    Dim recX001(1 To 2)     As c_cmzcrec
    Dim recX002(1 To 2)     As c_cmzcrec
    Dim recX003(1 To 2)     As c_cmzcrec        'GD��������_�f�[�^ 2005/02/15 ffc)tanabe
    Dim recX004(1 To 2)     As c_cmzcrec        'EP�������@06/08/10 ooba
    Dim recX005(1 To 2)     As c_cmzcrec        'EP����_�ް��@06/08/10 ooba
    Dim recX006(1 To 2)     As c_cmzcrec        'CuDeco�Ή� 2011/02/14 tkimura
    Dim recX007()           As c_cmzcrec        'GBG�Ή� 2011/06/23 Marushita
    Dim i                   As Integer
    Dim j                   As Integer
    Dim rs                  As OraDynaset
    Dim sql                 As String
    Dim XlSmpPos(1 To 2)    As Integer
    Dim CRYNUM              As String
    Dim blkID               As String
    Dim sBlkId(1 To 2)      As String       'XSDCW BLOCKID�i�[
    Dim smpId(2)            As String
    Dim HIN                 As tFullHinban
'2003/10/19 ۰�ٕϐ��ǉ� SystemBrain ==================================��
    Dim recXSDCS(1 To 2)    As c_cmzcrec        '�V����يǗ�(��ۯ�)
    Dim recXSDCW(1 To 2)    As c_cmzcrec        '�V����يǗ�(SXL)
    Dim recXSDCW_1()        As c_cmzcrec        '�V����يǗ�(SXL���Ԕ���) GBG�Ή� 2011/06/23 Marushita
    Dim recE037             As c_cmzcrec        '�������
    Dim recH001             As c_cmzcrec        '����w������ 08/12/01
    Dim recXSDC1            As c_cmzcrec        '��������
'2003/10/19 ۰�ٕϐ��ǉ� SystemBrain ==================================��

'2003/10/19 ۰�ٕϐ��폜 SystemBrain ==================================��
'    Dim recW009(1 To 2)     As c_cmzcrec
'    Dim recJ014(1 To 2)     As c_cmzcrec
'    Dim recY013(1 To 2)     As c_cmzcrec                '����]������
'    Dim fld As c_cmzcfld
'    Dim SXLPOS(1 To 2) As Integer
'    Dim fldNo As Integer
'    Dim fldCnt As Integer
'    Dim FldName As String
'    Dim specName As Variant
'    Dim CSDATA(1 To 2) As tCsData
'    Dim dMin(3)         As Double
'    Dim dMeas(9)        As Double
'    Dim strMeasPos      As String
'    Dim iRet            As Integer
'2003/10/19 ۰�ٕϐ��폜 SystemBrain ==================================��

    Dim RsHIN       As tFullHinban  '���R(Rs)�d�l�擾�i�ԁ@04/02/12 ooba
    Dim sRsData(10) As String       '���R(Rs)�ް��@04/02/12 ooba
'    Dim sRsPtn      As String       '���R�ް��擾����݁@04/02/12 ooba
    Dim sRsPtn(2)   As String       '���R�ް��擾����݁@04/04/15 ooba
    Dim sPos        As String       'SXL�ʒu(TOP/BOT)�@04/04/15 ooba
    Dim gSmpID(2)   As String       'TBCMX003�p�T���v��ID   2005/02/15 ffc)tanabe
    Dim sErrMsg     As String       '�װү���ށ@06/04/20 ooba
    Dim nowtime     As Date  ':2008/03/31 �� �BGB7/GB8/GB9��SXL�m����t����v������B
    
    '����EDI����ݸ�Ή� 2009/12/4 Add Start SPK habuki�@������
    Dim flgEDI      As Boolean      'EDI���L������p(True:�L�AFalse�F��)
    Dim dbPN        As Double       '�s�����Z�x(P:��)
    Dim dbBN        As Double       '�s�����Z�x(B:����)
    Dim dbASN       As Double       '�s�����Z�x(AS:��f)
    Dim dbCN        As Double       '�s�����Z�x(C:�Y�f)
    '����EDI����ݸ�Ή� 2009/12/4 Add Start SPK habuki�@������
    
    ''***********************************************************
    'Micron�Ή� 2011/01/14 Add start tkimura
    Dim dd          As type_Coefficient_new2    ''�����R,������㗦�v�Z�\����
    Dim sRsPos(2)   As String                   '���R(Rs)�ʒu[TOP/BOT]
    Dim data        As String                   '�V���O�����Œ�R�l���擾���Ă���Ƃ���SXLID,����ȊO�̂Ƃ���BLOCKID���擾����B
    ''2011/01/14 Add end tkimura
    ''***********************************************************
    'Add Start 2011/05/31 Y.Hitomi
    Dim sUP_RATIO(2) As String                  'SXL������/���グ��[TOP/BOT]
    'Add End   2011/05/31 Y.Hitomi
    '>>>>> 2011/06/24 SETsw)Marushita
    Dim iMCUTUNIT As Integer                    '���Ԕ����P��
    Dim sMSMPFLG  As Integer                    '���Ԕ����t���O
    '<<<<< 2011/06/24 SETsw)Marushita
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function WriteX00n"
    
    WriteX00n = FUNCTION_RETURN_FAILURE
    
    ':2008/03/31 �� �BGB7/GB8/GB9��SXL�m����t����v������B
    nowtime = getSvrTime()    '�T�[�o�[�̎��Ԃ��擾


    ''SXL�̕i�Ԃ��擾����
''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
    sql = ""
    sql = sql & " SELECT"
    sql = sql & "   HINBCB as HINBAN"           ''�i��
    sql = sql & "  ,REVNUMCB as REVNUM"         ''���i�ԍ������ԍ�
    sql = sql & "  ,FACTORYCB as FACTORY"       ''�H��
    sql = sql & "  ,OPECB as OPECOND"           ''���Ə���
    sql = sql & "  ,PLANTCATCB as PLANTCAT"     ''����  2007/09/04 SPK Tsutsumi Add
    sql = sql & " FROM"
    sql = sql & "   XSDCB"
    sql = sql & " WHERE SXLIDCB = '" & SXLID & "'"
''���ǉ�END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
''���폜START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
'    sql = "select HINBAN, REVNUM, FACTORY, OPECOND from TBCME042 where SXLID = '" & SXLID & "'"
''���폜END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount < 1 Then
''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
        errmsg = "XSDCB:" & rs.RecordCount
''���ǉ�END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
''���폜START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
'        errmsg = "E042:" & rs.RecordCount
''���폜END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
        rs.Close
        GoTo proc_exit
    End If
    HIN.hinban = rs!hinban
    HIN.mnorevno = rs!REVNUM
    HIN.factory = rs!factory
    HIN.opecond = rs!opecond
    HIN.sMukesaki = rs!PLANTCAT
    Set rs = Nothing
    
    ''�c���_�f�d�l�`�F�b�N�ǉ��@03/12/19 ooba START =====================>
    iChkAoi = ChkAoiSiyou(HIN)
    If iChkAoi < 0 Then
        errmsg = "�c���_�f(AOi)�d�l�G���[" & "  (" & HIN.hinban & Format(HIN.mnorevno, "00") & _
                                                    HIN.factory & HIN.opecond & ")"
        WriteX00n = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    ''�c���_�f�d�l�`�F�b�N�ǉ��@03/12/19 ooba END =======================>
        
    '-------------------- XSDCW�̓ǂݍ��� ----------------------------------------
    sql = "select * from XSDCW where SXLIDCW = '" & SXLID & "' and LIVKCW = '0' order by INPOSCW"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 2 Then
        errmsg = "XSDCW:" & rs.RecordCount
        Set rs = Nothing
        GoTo proc_exit
    End If
    Set recXSDCW(1) = New c_cmzcrec
    recXSDCW(1).CopyFromRs "XSDCW", rs
    rs.MoveNext
    Set recXSDCW(2) = New c_cmzcrec
    recXSDCW(2).CopyFromRs "XSDCW", rs
    Set rs = Nothing
    
    CRYNUM = left$(SXLID, 9) & "000"        '�����ԍ�
    
    '-------------------- �����ID����ۯ�ID�̎擾 ----------------------------------------
    ' �����ID(From)�A�����ID(To)��������Ȃ���΁A��ۯ�ID��Ă��A
    ' �����ID(From)�A�����ID(To)��������΁A��ۯ�ID��Ă��Ȃ��B
    smpId(1) = ""       '�����ID(From)������
    smpId(2) = ""       '�����ID(To)������
    blkID = ""          '��ۯ�ID������
    gSmpID(1) = ""      'TBCMX003�p�����ID(From)������  2005/02/18 ffc)tanabe
    gSmpID(2) = ""      'TBCMX003�p�����ID(To)������    2005/02/18 ffc)tanabe
    
    sql = "select REPSMPLIDCW,WFSMPLIDGDCW from XSDCW where SXLIDCW = '" & SXLID & "' and KTKBNCW != '9' and LIVKCW = '0' order by INPOSCW"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    '��ۯ�ID��K���Z�b�g����悤�ɕύX�iXSDCW�j 2005/3/22 TUKU START ------------------------------------------
    sBlkId(1) = Trim$(recXSDCW(1)("SMCRYNUMCW").Value)      'TOP BLOCKID
    sBlkId(2) = Trim$(recXSDCW(2)("SMCRYNUMCW").Value)      'BOT BLOCKID

    If rs.RecordCount = 2 Then
        smpId(1) = rs("REPSMPLIDCW")
        ':2008/03/31 �� �A�T���v��ID�𔽉f���̃T���v��ID(�����T���v��ID�܂�)�ł͂Ȃ��A��\�T���v��ID�ɕύX����B
        ''gSmpID(1) = rs("WFSMPLIDGDCW")      '�ǉ� 2005/02/18 ffc)tanabe
        gSmpID(1) = rs("REPSMPLIDCW")
        rs.MoveNext
        
        smpId(2) = rs("REPSMPLIDCW")
        ':2008/03/31 �� �A�T���v��ID�𔽉f���̃T���v��ID(�����T���v��ID�܂�)�ł͂Ȃ��A��\�T���v��ID�ɕύX����B
        ''gSmpID(2) = rs("WFSMPLIDGDCW")      '�ǉ� 2005/02/18 ffc)tanabe
        gSmpID(2) = rs("REPSMPLIDCW")
        Set rs = Nothing
    
        '���L�T���v���`�F�b�N����
        If chkComSAMPL(SXLID, smpId(1), smpId(1)) Then
            errmsg = "WriteX00n:���L�����ID�擾(From)"
            GoTo proc_exit
        End If
        If chkComSAMPL(SXLID, smpId(2), smpId(2)) Then
            errmsg = "WriteX00n:���L�����ID�擾(To)"
            GoTo proc_exit
        End If
    Else
        Set rs = Nothing
        
        
        ':2008/03/31 �� �A�T���v��ID�𔽉f���̃T���v��ID(�����T���v��ID�܂�)�ł͂Ȃ��A��\�T���v��ID�ɕύX����B
''''        '�m��敪=�9��Ō���GD�����p���ł���ꍇ�̑Ή��@05/10/24 ooba START ================>
''''        sql = "select WFSMPLIDGDCW from XSDCW "
''''        sql = sql & "where SXLIDCW = '" & SXLID & "' "
''''        sql = sql & "and (KTKBNCW != '9' "
''''        sql = sql & "or (KTKBNCW = '9' and WFINDGDCW <> '0' and WFRESGDCW <> '0')) "
''''        sql = sql & "and LIVKCW = '0' "
''''        sql = sql & "order by INPOSCW"
''''        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''''        If rs.RecordCount = 2 Then
''''            gSmpID(1) = rs("WFSMPLIDGDCW")
''''            rs.MoveNext
''''            gSmpID(2) = rs("WFSMPLIDGDCW")
''''            Set rs = Nothing
''''        Else
''''            Set rs = Nothing
''''        End If
''''        '�m��敪=�9��Ō���GD�����p���ł���ꍇ�̑Ή��@05/10/24 ooba END ==================>
        
        
        
        ''��ۯ�ID�͕K���擾����̂ŃR�����g�� 2005/3/22 TUKU
        '�u���b�NID�𓾂�
''''        sql = "select BLOCKID from TBCME040 "
''''        sql = sql & "where crynum = '" & CRYNUM & "' and "
''''        sql = sql & "INGOTPOS <= " & recXSDCW(1)("INPOSCW").Value & " and "
''''        sql = sql & "INGOTPOS + LENGTH > " & recXSDCW(1)("INPOSCW").Value
''''        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''''        If rs.RecordCount = 1 Then
''''            BlkId = rs("BLOCKID")
''''        End If
''''        Set rs = Nothing
    End If
    '��ۯ�ID��K���Z�b�g����悤�ɕύX�iXSDCW�j 2005/3/22 TUKU END ------------------------------------------
    
    '-------------------- XSDCS�̓ǂݍ��� ----------------------------------------
    For j = 1 To 2
        If j = 1 Then
            '�߂�XL����ʒu(FROM)�����߂�
            sql = "select * from XSDCS where CRYNUMCS = '" & Trim$(recXSDCW(j)("SMCRYNUMCW").Value) & "' and "
            sql = sql & "TBKBNCS = 'T' and LIVKCS = '0'"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount <> 1 Then
                errmsg = "XSDCS:From"
                Set rs = Nothing
                GoTo proc_exit
            End If
            Set recXSDCS(1) = New c_cmzcrec
            recXSDCS(1).CopyFromRs "XSDCS", rs
            Set rs = Nothing
            XlSmpPos(1) = recXSDCS(1)("INPOSCS").Value
        ElseIf j = 2 Then
            '�߂�XL����ʒu(TO)�����߂�
            sql = "select * from XSDCS where CRYNUMCS = '" & Trim$(recXSDCW(j)("SMCRYNUMCW").Value) & "' and "
            sql = sql & "TBKBNCS = 'B' and LIVKCS = '0'"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount <> 1 Then
                errmsg = "XSDCS:To"
                Set rs = Nothing
                GoTo proc_exit
            End If
            Set recXSDCS(2) = New c_cmzcrec
            recXSDCS(2).CopyFromRs "XSDCS", rs
            Set rs = Nothing
            XlSmpPos(2) = recXSDCS(2)("INPOSCS").Value
        End If
    Next j

    '-------------------- TBCME037�̓ǂݍ��� ----------------------------------------
    sql = "select * from TBCME037 where (CRYNUM='" & CRYNUM & "')"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        errmsg = "TBCME037"
        Set rs = Nothing
        GoTo proc_exit
    End If
    Set recE037 = New c_cmzcrec
    recE037.CopyFromRs "TBCME037", rs
    Set rs = Nothing

    '-------------------- TBCMH001�̓ǂݍ��� ---------------------------------------- (08/12/01)
    sql = "select * from TBCMH001 where UPINDNO = '"
      '8���ڂ�0�Ƃ���B 2009/12/23 Change Y.Hitomi
    sql = sql & Mid(CRYNUM, 1, 7) & "0" & Mid(CRYNUM, 9, 1) & "'"
'    Del 2009/12/23 Y.Hitomi
'    If Mid(CRYNUM, 9, 1) = "A" Or Mid(CRYNUM, 9, 1) = "B" Or Mid(CRYNUM, 9, 1) = "C" Then
'        '�c�ʈ���(8����9���ڂ�0)
'        sql = sql & Mid(CRYNUM, 1, 7) & "00'"
'    Else
'        '�ʏ�i,������(8���ڂ�0)
'        sql = sql & Mid(CRYNUM, 1, 7) & "0" & Mid(CRYNUM, 9, 1) & "'"
'    End If
    

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        errmsg = "TBCMH001"
        Set rs = Nothing
        GoTo proc_exit
    End If
    Set recH001 = New c_cmzcrec
    recH001.CopyFromRs "TBCMH001", rs
    Set rs = Nothing
    
    '-------------------- XSDC1�̓ǂݍ��� ----------------------------------------
    sql = "select * from XSDC1 where (XTALC1='" & CRYNUM & "')"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        errmsg = "XSDC1"
        Set rs = Nothing
        GoTo proc_exit
    End If
    Set recXSDC1 = New c_cmzcrec
    recXSDC1.CopyFromRs "XSDC1", rs
    Set rs = Nothing

    '����EDI����ݸ�Ή� 2009/12/4 Add Start SPK habuki�@������
    'EDI���擾
    If Not fncGetEdiInfo(SXLID, dbPN, dbBN, dbASN, dbCN, flgEDI) Then
        errmsg = "XODC6_1"
        GoTo proc_exit
    End If
    '����EDI����ݸ�Ή� 2009/12/4 Add End   SPK habuki�@������
    
    '==============================================
    '�@�e����уf�[�^�̎擾�E�ݒ�
    '==============================================
    For i = 1 To 2
        '-------------------- TBCMX001�Œ���f�[�^�ݒ� ----------------------------------------
        Set recX001(i) = New c_cmzcrec
        recX001(i).TABLENAME = "TBCMX001"
        recX001(i).SetRecDefault
        
        With recX001(i)
            .Fields("SXLID").Value = SXLID                                                  'SXLID
            .Fields("FROMTOKBN").Value = IIf(recXSDCW(i)("TBKBNCW").Value = "T", "1", "2")  'FROMTO�敪
            .Fields("SAMPLE_FROM").Value = smpId(1)                                         '�T���v��ID(From)
            .Fields("SAMPLE_TO").Value = smpId(2)                                           '�T���v��ID(To)
            'XODCW���K��BOLCKID�ݒ肷��悤�ɕύX 2005/03/22 TUKU
            '.Fields("BLOCKID").Value = BlkId                                                '�u���b�NID
            .Fields("BLOCKID").Value = sBlkId(i)                                                '�u���b�NID
            .Fields("CRYNUM").Value = CRYNUM                                                '�����ԍ�
            
            ':2008/03/31 �� �BGB7/GB8/GB9��SXL�m����t����v������B
            '' .Fields("SXLDECDATE").Value = "SYSDATE"                                         'SXL-ID�m����t
            .Fields("SXLDECDATE").Value = nowtime                                           'SXL-ID�m����t
            .nowtime = nowtime
            
            .Fields("PLUPDATE").Value = recXSDC1("TDAYC1").Value                            '������t
            .Fields("INGOTPOS").Value = recXSDCW(i)("INPOSCW").Value                        '�������J�n�ʒu
            .Fields("HINBAN").Value = HIN.hinban                                            '�i��
            .Fields("REVNUM").Value = HIN.mnorevno                                          '���i�ԍ������ԍ�
            .Fields("FACTORY").Value = HIN.factory                                          '�H��
            .Fields("OPECOND").Value = HIN.opecond                                          '���Ə���
            .Fields("PRODCOND").Value = recE037("PRODCOND").Value                           '�������
            .Fields("PGID").Value = Mid(recE037("PGID"), 1, 8)                              'PG-ID
            .Fields("UPLENGTH").Value = recE037("UPLENGTH").Value                           '���グ����
            .Fields("SXLPOS").Value = 0                                                     'SXL�ʒu
            .Fields("SXLLENGTH").Value = XlSmpPos(2) - XlSmpPos(1)                          'SXL-ID�m�蒷��
            .Fields("SXLWAFERCNT").Value = WfCnt                                            'SXL-ID�m�莞��WF����
            .Fields("FREELENG").Value = recE037("FREELENG").Value                           '�t���[��
            .Fields("DIAMETER").Value = recE037("DIAMETER").Value                           '���a
'            .Fields("CHARGE").Value = recE037("CHARGE").Value                               '�`���[�W��
            .Fields("SEED").Value = recE037("SEED").Value                                   '�V�[�h
            If i = 1 Then                                                                   '�����ID
                .Fields("SAMPID").Value = .Fields("SAMPLE_FROM").Value                      'TOP���̒l
            Else
                .Fields("SAMPID").Value = .Fields("SAMPLE_TO").Value                        'TAIL���̒l
            End If
            .Fields("PLANTCAT").Value = HIN.sMukesaki                                       '���� 2007/09/04 SPK Tsutsumi Add
            .Fields("CHARGE").Value = recXSDC1("PUCHAGC1").Value                            '����ޗ� 08/12/01
            .Fields("ROCHARGE").Value = recH001("CHARGE").Value                             '�F��F�d���� 08/12/01
            
            '����EDI����ݸ�Ή� 2009/12/4 Add Start SPK habuki�@������
            If flgEDI Then
                .Fields("PXL_BORON").Value = dbBN                                           '�s�����Z�x�iB:���݁j
                .Fields("PXL_PHOSPHOR").Value = dbPN                                        '�s�����Z�x�iP:�݁j
                .Fields("PXL_CARBON").Value = dbCN                                          '�s�����Z�x�iC:�Y�f�j
                .Fields("PXL_ARSENIC").Value = dbASN                                        '�s�����Z�x�iAS:��f�j
            End If
            '����EDI����ݸ�Ή� 2009/12/4 Add End   SPK habuki�@������
            
            'Add Start 2011/05/30 Y.Hitomi �}���`�Ή�
            '�}���`�t���O
            If Int(recH001("SIJICNT").Value) = 1 Then
                .Fields("MULTI_FLG") = "A"
            ElseIf Int(recH001("SIJICNT").Value) >= 2 Then
                .Fields("MULTI_FLG") = "M"
            End If
            '�c�ʈ����t���O
            If IsNumeric(Mid(CRYNUM, 9, 1)) = True Then
                .Fields("ZANRYO_FLG") = Mid(CRYNUM, 9, 1)
            ElseIf Mid(CRYNUM, 9, 1) = "A" Then
                .Fields("ZANRYO_FLG") = "7"
            ElseIf Mid(CRYNUM, 9, 1) = "B" Then
                .Fields("ZANRYO_FLG") = "8"
            ElseIf Mid(CRYNUM, 9, 1) = "C" Then
                .Fields("ZANRYO_FLG") = "9"
            End If
            'Add End   2011/05/30 Y.Hitomi
        
        End With
        
        '-------------------- TBCMX002�Œ���f�[�^�ݒ� ----------------------------------------
        Set recX002(i) = New c_cmzcrec
        recX002(i).TABLENAME = "TBCMX002"
        recX002(i).SetRecDefault
        
        With recX002(i)
            .Fields("SXLID").Value = SXLID                                                  'SXLID
            .Fields("FROMTOKBN").Value = IIf(recXSDCW(i)("TBKBNCW").Value = "T", "1", "2")  'FROMTO�敪
            .Fields("SAMPLE_FROM").Value = smpId(1)                                         '�T���v��ID(From)
            .Fields("SAMPLE_TO").Value = smpId(2)                                           '�T���v��ID(To)
            'XODCW���K��BOLCKID�ݒ肷��悤�ɕύX 2005/03/22 TUKU
            '.Fields("BLOCKID").Value = BlkId                                                '�u���b�NID
            .Fields("BLOCKID").Value = sBlkId(i)                                                '�u���b�NID
            .Fields("CRYNUM").Value = CRYNUM                                                '�����ԍ�
            
            ':2008/03/31 �� �BGB7/GB8/GB9��SXL�m����t����v������B
            '' .Fields("SXLDECDATE").Value = "SYSDATE"                                         'SXL-ID�m����t
            .Fields("SXLDECDATE").Value = nowtime                                           'SXL-ID�m����t
            .nowtime = nowtime

            
            .Fields("PLUPDATE").Value = recXSDC1("TDAYC1").Value                            '������t
            .Fields("INGOTPOS").Value = recXSDCW(i)("INPOSCW").Value                        '�������J�n�ʒu
            .Fields("HINBAN").Value = HIN.hinban                                            '�i��
            .Fields("REVNUM").Value = HIN.mnorevno                                          '���i�ԍ������ԍ�
            .Fields("FACTORY").Value = HIN.factory                                          '�H��
            .Fields("UPLENGTH").Value = recE037("UPLENGTH").Value                           '���グ����
            .Fields("SXLPOS").Value = 0                                                     'SXL�ʒu
            .Fields("SXLLENGTH").Value = XlSmpPos(2) - XlSmpPos(1)                          'SXL-ID�m�蒷��
            .Fields("SXLWAFERCNT").Value = WfCnt                                            'SXL-ID�m�莞��WF����
            .Fields("FREELENG").Value = recE037("FREELENG").Value                           '�t���[��
            If i = 1 Then                                                                   '�����ID
                .Fields("SAMPID_1").Value = .Fields("SAMPLE_FROM").Value                    'TOP���̒l
            Else
                .Fields("SAMPID_1").Value = .Fields("SAMPLE_TO").Value                      'TAIL���̒l
            End If
            .Fields("PLANTCAT").Value = HIN.sMukesaki                                       '���� 2007/09/04 SPK Tsutsumi Add
        End With
                
        If i = 1 Then sPos = "TOP" Else sPos = "BOT"    '04/04/15 ooba
        
        '-------------------- (����Rs)������R����(TBCMJ002)�f�[�^�擾�ݒ� ----------------------------------------
        If getTBCMJ002(CRYNUM, recXSDCS(), i, HIN, recX001(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "J002:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '-------------------- (����Oi)����Oi����(TBCMJ003)�f�[�^�擾�ݒ� ----------------------------------------
        If getTBCMJ003(CRYNUM, recXSDCS(i), HIN, recX001(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "J003:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '-------------------- (Cs)Cs����(TBCMJ004)�f�[�^�擾�ݒ� ----------------------------------------
'        If getTBCMJ004(CRYNUM, recXSDCS(i), recX001(i)) = FUNCTION_RETURN_FAILURE Then
        '�i�ԁ��װү���ޒǉ��@06/04/20 ooba
        If getTBCMJ004(CRYNUM, recXSDCS(i), HIN, recX001(i), sErrMsg) = FUNCTION_RETURN_FAILURE Then
            If sErrMsg = "" Then
                errmsg = "J004:" & XlSmpPos(i)
            Else
                errmsg = sErrMsg
            End If
            GoTo proc_exit
        End If

        '-------------------- (����OSF1�`4)����OSF����(TBCMJ005)�f�[�^�擾�ݒ� ----------------------------------------
        For j = 1 To 4
            If getTBCMJ005(CRYNUM, recXSDCS(i), j, recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
                errmsg = "J005-" & j & ":" & XlSmpPos(i)
                GoTo proc_exit
            End If
        Next

        '-------------------- (����BMD1�`3)����BMD����(TBCMJ008)�f�[�^�擾�ݒ� ----------------------------------------
        For j = 1 To 3
            If getTBCMJ008(CRYNUM, recXSDCS(i), j, recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
                errmsg = "J008-" & j & ":" & XlSmpPos(i)
                GoTo proc_exit
            End If
        Next

        '-------------------- (GD)GD����(TBCMJ006)�f�[�^�擾�ݒ� ----------------------------------------
        If getTBCMJ006(CRYNUM, recXSDCS(i), recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "J006:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '-------------------- (LT)LT����(TBCMJ007)�f�[�^�擾�ݒ� ----------------------------------------
'        If getTBCMJ007(CRYNUM, recXSDCS(i), i, recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
        If getTBCMJ007(CRYNUM, recXSDCS(i), HIN, i, recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then  '05/12/05 ooba
            errmsg = "J007:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '-------------------- (WFOi)WFOi����(TBCMY013)�f�[�^�擾�ݒ� ----------------------------------------
        If getTBCMY013WFOi(recXSDCW(i), HIN, sPos, recX001(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "Y013-Oi:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '-------------------- (WFRs)WFRs����(TBCMY013)�f�[�^�擾�ݒ� ----------------------------------------
        If getTBCMY013WFRs(recXSDCW(i), HIN, sPos, recX001(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "Y013-Rs:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '-------------------- (WFDOi1�`3)WFDOi����(TBCMY013)�f�[�^�擾�ݒ� ----------------------------------------
        For j = 1 To 3
            If getTBCMY013WFDOi(recXSDCW(i), j, recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
                errmsg = "Y013-DOi" & j & ":" & XlSmpPos(i)
                GoTo proc_exit
            End If
        Next

        '-------------------- (WFOSF1�`4)WFOSF����(TBCMY013)�f�[�^�擾�ݒ� ----------------------------------------
'        For j = 1 To 4  Change  2010/04/19 SIRD�Ή�
        For j = 1 To 3
            If getTBCMY013WFOSF(recXSDCW(i), j, HIN, sPos, recX001(i), recX002(i), recX001(i).TABLENAME) = FUNCTION_RETURN_FAILURE Then
            'If getTBCMY013WFOSF(recXSDCW(i), j, HIN, sPos, recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
                errmsg = "Y013-OSF" & j & ":" & XlSmpPos(i)
                GoTo proc_exit
            End If
        Next

        '-------------------- (WFBMD1�`3)WFBMD����(TBCMY013)�f�[�^�擾�ݒ� ----------------------------------------
        For j = 1 To 3
            If getTBCMY013WFBMD(recXSDCW(i), j, HIN, sPos, recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
                errmsg = "Y013-BMD" & j & ":" & XlSmpPos(i)
                GoTo proc_exit
            End If
        Next

        '-------------------- (WFDSOD)WFDSOD����(TBCMY013)�f�[�^�擾�ݒ� ----------------------------------------
        If getTBCMY013WFDSOD(recXSDCW(i), recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "Y013-DSOD:" & XlSmpPos(i)
            GoTo proc_exit
        End If

    ''Upd start 2005/06/23 (TCS)T.Terauchi      SPV9�_�Ή�
'        '-------------------- (WFSPV)WFSPV����(TBCMY013)�f�[�^�擾�ݒ� ----------------------------------------
'        If getTBCMY013WFSPV(recXSDCW(i), recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
'            errmsg = "Y013-SPV:" & XlSmpPos(i)
'            GoTo proc_exit
'        End If
        '-------------------- (WFSPV)WFSPV����(TBCMJ016)�f�[�^�擾�ݒ� ----------------------------------------
        If getTBCMJ016WFSPV(CRYNUM, recXSDCW(i), HIN, recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "J016-SPV:" & XlSmpPos(i)
            GoTo proc_exit
        End If
        
        '-------------------- �W������(TBCMY018)���Warp���ѐݒ� ----------------------------------------------
        If getTBCMY018WARP(sBlkId(i), recXSDCW(i), recX001(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "Y018-WARP:" & XlSmpPos(i)
            GoTo proc_exit
        End If
        
    ''Upd end   2005/06/23 (TCS)T.Terauchi      SPV9�_�Ή�
    
        '-------------------- (WFDZ)WFDZ����(TBCMY013)�f�[�^�擾�ݒ� ----------------------------------------
        '�i�Ԓǉ��@06/09/06 ooba
        If getTBCMY013WFDZ(recXSDCW(i), HIN, recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
'        If getTBCMY013WFDZ(recXSDCW(i), recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "Y013-DZ:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        ''�c���_�f���ю擾�ǉ��@03/12/19 ooba START ================================================>
        '-------------------- (WFAOi)WFAOi����(TBCMY013)�f�[�^�擾�ݒ� ----------------------------------------
        If getTBCMY013WFAOi(recXSDCW(i), recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "Y013-AOi:" & XlSmpPos(i)
            GoTo proc_exit
        End If
        ''�c���_�f���ю擾�ǉ��@03/12/19 ooba END ==================================================>
        
        ''��SIRD�]�����ю擾�ǉ��@10/04/19 Y.Hitomi
        If getTBCMJ022SIRD(CRYNUM, recXSDCW(i), recX001(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "J022-SIRD:" & XlSmpPos(i)
            GoTo proc_exit
        End If
        ''��SIRD�]�����ю擾�ǉ��@10/04/19 Y.Hitomi
        

        '==============================================
        '�@TBCMX001 �ɏ�������
        '==============================================
        With recX001(i)
            .Fields("REGDATE").Value = "SYSDATE"                                    '�o�^���t
''            .Fields("SENDFLAG").Value = "3"                                       '���M�t���O
''            .Fields("SENDFLAG").Value = "0"                 '���M���ݸޕύX�@05/11/25 ooba
            .Fields("SENDFLAG").Value = "3"  ':2008/03/31 �� �@���sSXL�m����s������SXL�}�b�v��M���ɑ��M���ݸޕύX����B
            .Fields("SENDDATE").Value = "SYSDATE"                                   '���M���t
            sql = .SqlInsert
            If OraDB.ExecuteSQL(sql) < 1 Then
                errmsg = "X001-" & i
                GoTo proc_exit
            End If
        End With

        '==============================================
        '�@TBCMX002 �ɏ�������
        '==============================================
        With recX002(i)
            .Fields("REGDATE").Value = "SYSDATE"                                    '�o�^���t
''            .Fields("SENDFLAG").Value = "3"                                       '���M�t���O
''            .Fields("SENDFLAG").Value = "0"                 '���M���ݸޕύX�@05/11/25 ooba
            .Fields("SENDFLAG").Value = "3"  ':2008/03/31 �� �@���sSXL�m����s������SXL�}�b�v��M���ɑ��M���ݸޕύX����B
            .Fields("SENDDATE").Value = "SYSDATE"                                   '���M���t

            sql = .SqlInsert
            If OraDB.ExecuteSQL(sql) < 1 Then
                errmsg = "X002-" & i
                GoTo proc_exit
            End If
        End With

        ''GD��������_�f�[�^(TBCMX003)�̒ǉ� 2005/02/15 ffc)tanabe START ===========================>
        ''XSDCW��GD��s�]���̏�ԃt���O=1�����уt���O=1�̏ꍇTBCMX003�ɓo�^����B
        If (recXSDCW(i)("WFINDGDCW").Value <> "0") And (recXSDCW(i)("WFRESGDCW").Value <> "0") Then
            
            '-------------------- TBCMX003�Œ���f�[�^�ݒ� ----------------------------------------
            Set recX003(i) = New c_cmzcrec
            recX003(i).TABLENAME = "TBCMX003"
            recX003(i).SetRecDefault
            
            With recX003(i)
                .Fields("SXLID").Value = SXLID                                                  'SXLID
                .Fields("FROMTOKBN").Value = IIf(recXSDCW(i)("TBKBNCW").Value = "T", "1", "2")  'FROMTO�敪
                .Fields("SAMPLE_FROM").Value = gSmpID(1)                                        '�T���v��ID(From)
                .Fields("SAMPLE_TO").Value = gSmpID(2)                                          '�T���v��ID(To)
                .Fields("BLOCKID").Value = sBlkId(i)                                            '�u���b�NID
                .Fields("CRYNUM").Value = CRYNUM                                                '�����ԍ�
            
                ':2008/03/31 �� �BGB7/GB8/GB9��SXL�m����t����v������B
                '' .Fields("SXLDECDATE").Value = "SYSDATE"                                         'SXL-ID�m����t
                .Fields("SXLDECDATE").Value = nowtime                                           'SXL-ID�m����t
                .nowtime = nowtime
                
                .Fields("PLUPDATE").Value = recXSDC1("TDAYC1").Value                            '������t
                .Fields("INGOTPOS").Value = recXSDCW(i)("INPOSCW").Value                        '�������J�n�ʒu
                .Fields("HINBAN").Value = HIN.hinban                                            '�i��
                .Fields("REVNUM").Value = HIN.mnorevno                                          '���i�ԍ������ԍ�
                .Fields("FACTORY").Value = HIN.factory                                          '�H��
                .Fields("UPLENGTH").Value = recE037("UPLENGTH").Value                           '���グ����
                .Fields("SXLPOS").Value = 0                                                     'SXL�ʒu
                .Fields("SXLLENGTH").Value = XlSmpPos(2) - XlSmpPos(1)                          'SXL-ID�m�蒷��
                .Fields("SXLWAFERCNT").Value = WfCnt                                            'SXL-ID�m�莞��WF����
                .Fields("FREELENG").Value = recE037("FREELENG").Value                           '�t���[��
                If i = 1 Then                                                                   '�����ID
                    .Fields("SAMPID_1").Value = .Fields("SAMPLE_FROM").Value                    'TOP���̒l
                Else
                    .Fields("SAMPID_1").Value = .Fields("SAMPLE_TO").Value                      'TAIL���̒l
                End If
                .Fields("PLANTCAT").Value = HIN.sMukesaki                                       '���� 2007/09/04 SPK Tsutsumi Add
            End With
            
            '�ۏ؃t���O=1(����GD���т�ۏ�)�̏ꍇ
            If recXSDCW(i).Fields("WFHSGDCW") = "1" Then
                If getTBCMJ006GD(CRYNUM, recXSDCW(i), recX003(i)) = FUNCTION_RETURN_FAILURE Then
                    errmsg = "J006-GD:" & XlSmpPos(i)
                    GoTo proc_exit
                End If
            '�ۏ؃t���O=0(WFGD���т�ۏ�)�̏ꍇ
            Else
                If getTBCMJ015WFGD(CRYNUM, recXSDCW(i), recX003(i), recX003(i).TABLENAME) = FUNCTION_RETURN_FAILURE Then
                'If getTBCMJ015WFGD(CRYNUM, recXSDCW(i), recX003(i)) = FUNCTION_RETURN_FAILURE Then
                    errmsg = "J015-GD:" & XlSmpPos(i)
                    GoTo proc_exit
                End If
            End If
        
            '==============================================
            '�@TBCMX003 �ɏ�������
            '==============================================
            With recX003(i)
                .Fields("REGDATE").Value = "SYSDATE"                                    '�o�^���t
''                .Fields("SENDFLAG").Value = "3"                                       '���M�t���O
''                .Fields("SENDFLAG").Value = "0"             '���M���ݸޕύX�@05/11/25 ooba
                .Fields("SENDFLAG").Value = "3"  ':2008/03/31 �� �@���sSXL�m����s������SXL�}�b�v��M���ɑ��M���ݸޕύX����B
                .Fields("SENDDATE").Value = "SYSDATE"                                   '���M���t
                
                sql = .SqlInsert
                If OraDB.ExecuteSQL(sql) < 1 Then
                    errmsg = "X003-" & i
                    GoTo proc_exit
                End If
            End With
        End If
        
        ''GD��������_�f�[�^(TBCMX003)�̒ǉ� 2005/02/15 ffc)tanabe END ============================>
        
        
        'EP������(X004)/EP����_�ް�(X005)�쐬�@06/08/10 ooba START ===========================>
        If ((recXSDCW(i)("EPINDB1CW").Value <> "0") And (recXSDCW(i)("EPRESB1CW").Value <> "0")) Or _
            ((recXSDCW(i)("EPINDB2CW").Value <> "0") And (recXSDCW(i)("EPRESB2CW").Value <> "0")) Or _
            ((recXSDCW(i)("EPINDB3CW").Value <> "0") And (recXSDCW(i)("EPRESB3CW").Value <> "0")) Or _
            ((recXSDCW(i)("EPINDL1CW").Value <> "0") And (recXSDCW(i)("EPRESL1CW").Value <> "0")) Or _
            ((recXSDCW(i)("EPINDL2CW").Value <> "0") And (recXSDCW(i)("EPRESL2CW").Value <> "0")) Or _
            ((recXSDCW(i)("EPINDL3CW").Value <> "0") And (recXSDCW(i)("EPRESL3CW").Value <> "0")) Then
        
            '-------------------- TBCMX004�Œ����ް��ݒ� ---------------------------------------
            Set recX004(i) = New c_cmzcrec
            recX004(i).TABLENAME = "TBCMX004"
            recX004(i).SetRecDefault
            
            With recX004(i)
                .Fields("SXLID").Value = SXLID                                                  'SXLID
                .Fields("FROMTOKBN").Value = IIf(recXSDCW(i)("TBKBNCW").Value = "T", "1", "2")  'FROMTO�敪
                .Fields("SAMPLE_FROM").Value = smpId(1)                                         '�����ID(From)
                .Fields("SAMPLE_TO").Value = smpId(2)                                           '�����ID(To)
                .Fields("BLOCKID").Value = sBlkId(i)                                            '��ۯ�ID
                .Fields("CRYNUM").Value = CRYNUM                                                '�����ԍ�
                
                ':2008/03/31 �� �BGB7/GB8/GB9��SXL�m����t����v������B
                '' .Fields("SXLDECDATE").Value = "SYSDATE"                                         'SXL-ID�m����t
                .Fields("SXLDECDATE").Value = nowtime                                           'SXL-ID�m����t
                .nowtime = nowtime
                
                .Fields("PLUPDATE").Value = recXSDC1("TDAYC1").Value                            '������t
                .Fields("INGOTPOS").Value = recXSDCW(i)("INPOSCW").Value                        '�������J�n�ʒu
                .Fields("HINBAN").Value = HIN.hinban                                            '�i��
                .Fields("REVNUM").Value = HIN.mnorevno                                          '���i�ԍ������ԍ�
                .Fields("FACTORY").Value = HIN.factory                                          '�H��
                .Fields("UPLENGTH").Value = recE037("UPLENGTH").Value                           '���グ����
                .Fields("SXLPOS").Value = 0                                                     'SXL�ʒu
                .Fields("SXLLENGTH").Value = XlSmpPos(2) - XlSmpPos(1)                          'SXL-ID�m�蒷��
                .Fields("SXLWAFERCNT").Value = WfCnt                                            'SXL-ID�m�莞��WF����
                .Fields("FREELENG").Value = recE037("FREELENG").Value                           '�ذ��
                If i = 1 Then                                                                   '�����ID
                    .Fields("SAMPID").Value = .Fields("SAMPLE_FROM").Value                      'TOP���̒l
                Else
                    .Fields("SAMPID").Value = .Fields("SAMPLE_TO").Value                        'TAIL���̒l
                End If
                .Fields("PLANTCAT").Value = HIN.sMukesaki                                       '���� 2007/09/04 SPK Tsutsumi Add
            End With
            
            '-------------------- TBCMX005�Œ����ް��ݒ� ---------------------------------------
            Set recX005(i) = New c_cmzcrec
            recX005(i).TABLENAME = "TBCMX005"
            recX005(i).SetRecDefault
            
            With recX005(i)
                .Fields("SXLID").Value = SXLID                                                  'SXLID
                .Fields("FROMTOKBN").Value = IIf(recXSDCW(i)("TBKBNCW").Value = "T", "1", "2")  'FROMTO�敪
                .Fields("SAMPLE_FROM").Value = smpId(1)                                         '�����ID(From)
                .Fields("SAMPLE_TO").Value = smpId(2)                                           '�����ID(To)
                .Fields("BLOCKID").Value = sBlkId(i)                                            '��ۯ�ID
                .Fields("CRYNUM").Value = CRYNUM                                                '�����ԍ�
                
                ':2008/03/31 �� �BGB7/GB8/GB9��SXL�m����t����v������B
                '' .Fields("SXLDECDATE").Value = "SYSDATE"                                         'SXL-ID�m����t
                .Fields("SXLDECDATE").Value = nowtime                                           'SXL-ID�m����t
                .nowtime = nowtime
                
                .Fields("PLUPDATE").Value = recXSDC1("TDAYC1").Value                            '������t
                .Fields("INGOTPOS").Value = recXSDCW(i)("INPOSCW").Value                        '�������J�n�ʒu
                .Fields("HINBAN").Value = HIN.hinban                                            '�i��
                .Fields("REVNUM").Value = HIN.mnorevno                                          '���i�ԍ������ԍ�
                .Fields("FACTORY").Value = HIN.factory                                          '�H��
                .Fields("UPLENGTH").Value = recE037("UPLENGTH").Value                           '���グ����
                .Fields("SXLPOS").Value = 0                                                     'SXL�ʒu
                .Fields("SXLLENGTH").Value = XlSmpPos(2) - XlSmpPos(1)                          'SXL-ID�m�蒷��
                .Fields("SXLWAFERCNT").Value = WfCnt                                            'SXL-ID�m�莞��WF����
                .Fields("FREELENG").Value = recE037("FREELENG").Value                           '�ذ��
                If i = 1 Then                                                                   '�����ID
                    .Fields("SAMPID").Value = .Fields("SAMPLE_FROM").Value                      'TOP���̒l
                Else
                    .Fields("SAMPID").Value = .Fields("SAMPLE_TO").Value                        'TAIL���̒l
                End If
                .Fields("PLANTCAT").Value = HIN.sMukesaki                                       '���� 2007/09/04 SPK Tsutsumi Add
            End With
            
            '-------------------- ���OSF1�`3����(TBCMY022)�ް��擾�ݒ� ---------------------------
            For j = 1 To 3
                If getTBCMY022EPOSF(recXSDCW(i), j, HIN, sPos, recX004(i), recX005(i), recX004(i).TABLENAME) = FUNCTION_RETURN_FAILURE Then
                'If getTBCMY022EPOSF(recXSDCW(i), j, HIN, sPos, recX004(i), recX005(i)) = FUNCTION_RETURN_FAILURE Then
                    errmsg = "Y022-EPOSF" & j & ":" & XlSmpPos(i)
                    GoTo proc_exit
                End If
            Next j
            
            '-------------------- ���BMD1�`3����(TBCMY022)�ް��擾�ݒ� ---------------------------
            For j = 1 To 3
                If getTBCMY022EPBMD(recXSDCW(i), j, HIN, sPos, recX004(i), recX005(i)) = FUNCTION_RETURN_FAILURE Then
                    errmsg = "Y022-EPBMD" & j & ":" & XlSmpPos(i)
                    GoTo proc_exit
                End If
            Next j
            
            '==============================================
            '�@TBCMX004 �ɏ�������
            '==============================================
            With recX004(i)
                .Fields("REGDATE").Value = "SYSDATE"                                    '�o�^���t
''                .Fields("SENDFLAG").Value = "0"                                       '���M�׸�
                .Fields("SENDFLAG").Value = "3"  ':2008/03/31 �� �@���sSXL�m����s������SXL�}�b�v��M���ɑ��M���ݸޕύX����B
                .Fields("SENDDATE").Value = "SYSDATE"                                   '���M���t
    
                sql = .SqlInsert
                If OraDB.ExecuteSQL(sql) < 1 Then
                    errmsg = "X004-" & i
                    GoTo proc_exit
                End If
            End With
            
            '==============================================
            '�@TBCMX005 �ɏ�������
            '==============================================
            With recX005(i)
                .Fields("REGDATE").Value = "SYSDATE"                                    '�o�^���t
''                .Fields("SENDFLAG").Value = "0"                                       '���M�׸�
                .Fields("SENDFLAG").Value = "3"  ':2008/03/31 �� �@���sSXL�m����s������SXL�}�b�v��M���ɑ��M���ݸޕύX����B
                .Fields("SENDDATE").Value = "SYSDATE"                                   '���M���t
    
                sql = .SqlInsert
                If OraDB.ExecuteSQL(sql) < 1 Then
                    errmsg = "X005-" & i
                    GoTo proc_exit
                End If
            End With
        End If
        'EP������(X004)/EP����_�ް�(X005)�쐬�@06/08/10 ooba END =============================>
        
'=================================================================================
' 2011/02/14 tkimura ADD START
        '-------------------- TBCMX006�Œ���f�[�^�ݒ� ----------------------------------------
        'Add 2011/03/01 Y.Hitomi C-OSF3�w���L��ł���΁AX006���쐬����B
        If recXSDCS(i)("CRYINDL4CS").Value = "1" Or recXSDCS(i)("CRYINDL4CS").Value = "2" Then
        
            Set recX006(i) = New c_cmzcrec
            recX006(i).TABLENAME = "TBCMX006"
            recX006(i).SetRecDefault
            
            With recX006(i)
                .Fields("PLANTCAT").Value = HIN.sMukesaki
                .Fields("SXLID").Value = SXLID                                                  'SXLID
                .Fields("FROMTOKBN").Value = IIf(recXSDCW(i)("TBKBNCW").Value = "T", "1", "2")  'FROMTO�敪
                'XODCW���K��BOLCKID�ݒ肷��悤�ɕύX 2005/03/22 TUKU
                '.Fields("BLOCKID").Value = BlkId                                               '�u���b�NID
                .Fields("BLOCKID").Value = sBlkId(i)                                            '�u���b�NID
                .Fields("CRYNUM").Value = CRYNUM                                                '�����ԍ�
                
                ':2011/02/14 tkimura �BGB7/GB8/GB9/GBF��SXL�m����t����v������B
                .Fields("SXLDECDATE").Value = nowtime                                           'SXL-ID�m����t
                .nowtime = nowtime
                
                .Fields("HINBAN").Value = HIN.hinban                                            '�i��
                .Fields("REVNUM").Value = HIN.mnorevno                                          '���i�ԍ������ԍ�
                .Fields("FACTORY").Value = HIN.factory                                          '�H��
                .Fields("OPECOND").Value = HIN.opecond                                          '���Ə���
                
                .Fields("SXL_SMPLICHI").Value = recXSDCS(i)("INPOSCS").Value                    '�������ʒu
                .Fields("SXL_SMPLNO").Value = recXSDCS(i)("CRYSMPLIDL4CS").Value                'C-OSF3�T���v��No
    
            End With
    
            '-------------------- (����OSF3)����OSF����(TBCMJ005)�f�[�^�擾�ݒ� ----------------------------------------
            If getTBCMJ005CuDeco(CRYNUM, recXSDCS(i), recX006(i)) = FUNCTION_RETURN_FAILURE Then
                errmsg = "J005-" & "CuDeco" & ":" & XlSmpPos(i)
                GoTo proc_exit
            End If
            
            '-------------------- (CLESTA)CLESTA�]������(TBCMJ023)�f�[�^�擾�ݒ� ----------------------------------------
            '���͂̍ۂɕK�v�Ȃ��̂͌����ԍ�+�T���v��ID
            If getTBCMJ023(CRYNUM, recXSDCS(i), recX006(i)) = FUNCTION_RETURN_FAILURE Then
                errmsg = "J023-:" & XlSmpPos(i)
                GoTo proc_exit
            End If
            
            '==============================================
            '�@TBCMX006 �ɏ�������
            '==============================================
            With recX006(i)
                .Fields("REGDATE").Value = "SYSDATE"                                    '�o�^���t
                .Fields("SENDFLAG").Value = "3"                                         ':2008/03/31 �� �@���sSXL�m����s������SXL�}�b�v��M���ɑ��M���ݸޕύX����
                .Fields("SENDDATE").Value = "SYSDATE"                                   '���M���t
                sql = .SqlInsert
                Debug.Print (sql)
                If OraDB.ExecuteSQL(sql) < 1 Then
                    errmsg = "X006-" & i
                    GoTo proc_exit
                End If
            End With
        End If
' 2011/02/14 tkimura ADD END
'=================================================================================
        
        ''TOP/BOT�ʂɔ��R�ް��擾�@04/04/15 ooba START ======================================>
        '���R�ް��擾����݂����߂�B
        If SxlRsPattern(HIN, sPos, sRsPtn(i)) = FUNCTION_RETURN_FAILURE Then
            '�擾�ް��Ȃ�
            sRsPtn(i) = "C"
        End If
        '���R�ް����擾����B
        If cmbc040_GetSxlRsData(SXLID, sPos, sRsPtn(i), sRsData()) = FUNCTION_RETURN_FAILURE Then
            If sRsPtn(i) = "A" Then errmsg = "WF"
            If sRsPtn(i) = "B" Then errmsg = "����"
            errmsg = errmsg & "���R�����ް����擾�ł��܂���(Y007)"
            GoTo proc_exit
        End If
        ''TOP/BOT�ʂɔ��R�ް��擾�@04/04/15 ooba END ========================================>
        
'=================================================================================
' 2011/01/14 tkimura ADD START
''�@��R��Top�ʒu�擾���@,��R��Bot�ʒu�擾���@
        '�p�^�[��A�̂Ƃ��̓V���O��ID,B�̂Ƃ��̓u���b�NID
        If sRsPtn(i) = "A" Then
            data = SXLID
        ElseIf sRsPtn(i) = "B" Then
            data = sBlkId(i)
        Else
            data = ""
        End If

        '�֐���:cmbc040_GetSxlRsPos(�V���O��IDor�u���b�NID,sPos,sRsPtn(i),
        'sRsPos(i))+��\�T���v��ID[smpId(1),smpId(2)]
        'sRsPos(i)�Ƀσg�b�v�ʒu�ƃσ{�g���ʒu���i�[����B
        If cmbc040_GetSxlRsPos(data, _
                               sPos, _
                               sRsPtn(i), _
                               sRsPos()) = FUNCTION_RETURN_FAILURE Then
            If sRsPtn(i) = "A" Then errmsg = "WF"
            If sRsPtn(i) = "B" Then errmsg = "����"
            errmsg = errmsg & "���R�ʒu�ް����擾�ł��܂���"
            GoTo proc_exit
        End If
' 2011/01/14 tkimura ADD END
'=================================================================================
'        If i = 1 Then
        If i = 2 Then   '04/04/15 ooba
            If smpId(1) = vbNullString Then
                '' �u���b�NID�擾�@2003/09/16 Motegi ==================================> START
'                    blkID = Trim$(recW009(1)("BLOCKID").Value)
                blkID = Trim$(recXSDCW(1)("SMCRYNUMCW").Value)
                '' �u���b�NID�擾�@2003/09/16 Motegi ==================================> END
            Else
                blkID = vbNullString
            End If
            
'''''            ''���R�ް��擾�@04/02/12 ooba START ===========================================>
'''''            If SXLID <> vbNullString Then
'''''                'SXL�̕i�Ԃ��擾
'''''                RsHIN.HINBAN = ""
'''''                sql = "select HINBAN, REVNUM, FACTORY, OPECOND from TBCME042 "
'''''                sql = sql & "where SXLID = '" & SXLID & "' "
'''''                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                If rs.RecordCount = 1 Then
'''''                    RsHIN.HINBAN = rs("HINBAN")
'''''                    RsHIN.mnorevno = rs("REVNUM")
'''''                    RsHIN.factory = rs("FACTORY")
'''''                    RsHIN.opecond = rs("OPECOND")
'''''                End If
'''''                Set rs = Nothing
'''''                '���R�ް��擾����݂����߂�B
'''''                If SxlRsPattern(RsHIN, sPos, sRsPtn) = FUNCTION_RETURN_FAILURE Then
'''''                    '�擾�ް��Ȃ�
'''''                    sRsPtn = "C"
'''''                End If
'''''                '���R�ް����擾����B
'''''                If cmbc040_GetSxlRsData(SXLID, sRsPtn, sRsData()) = FUNCTION_RETURN_FAILURE Then
'''''                    If sRsPtn = "A" Then errmsg = "WF"
'''''                    If sRsPtn = "B" Then errmsg = "����"
'''''                    errmsg = errmsg & "���R�����ް����擾�ł��܂���(Y007)"
'''''                    GoTo proc_exit
'''''                End If
'''''            End If
'''''            ''���R�ް��擾�@04/02/12 ooba END =============================================>
            
            ''TBCMY007
            ''���R�ް��o�^�ǉ��@04/02/12 ooba

'' 2007/09/04 SPK Tsutsumi Add Start
''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
            sql = "Insert into TBCMY007 " & _
                  "(SXL_ID,SAMPLE_FROM,SAMPLE_TO,BLOCKID,HINBAN,KUBUN,TXID,REGDATE,SUMMITSENDFLAG,SENDFLAG,SENDDATE, " & _
                  "MESDATA1TOP, MESDATA2TOP, MESDATA3TOP, MESDATA4TOP, MESDATA5TOP, " & _
                  "MESDATA1BOT, MESDATA2BOT, MESDATA3BOT, MESDATA4BOT, MESDATA5BOT, PLANTCAT) " & _
                  "values (" & _
                  NoNullStr(SXLID) & ", " & _
                  NoNullStr(smpId(1)) & ", " & _
                  NoNullStr(smpId(2)) & ", " & _
                  NoNullStr(blkID) & ", " & _
                  "(select distinct HINBCB||to_char(REVNUMCB,'FM00') from XSDCB where SXLIDCB=" & NoNullStr(SXLID) & "), " & _
                  "'00', " & _
                  "'TX853I', " & _
                  "SYSDATE, '0', '0', SYSDATE, " & _
                  NoNullStr(sRsData(1)) & ", " & _
                  NoNullStr(sRsData(2)) & ", " & _
                  NoNullStr(sRsData(3)) & ", " & _
                  NoNullStr(sRsData(4)) & ", " & _
                  NoNullStr(sRsData(5)) & ", " & _
                  NoNullStr(sRsData(6)) & ", " & _
                  NoNullStr(sRsData(7)) & ", " & _
                  NoNullStr(sRsData(8)) & ", " & _
                  NoNullStr(sRsData(9)) & ", " & _
                  NoNullStr(sRsData(10)) & "," & _
                  "'" & sCmbMukesaki & "'" & _
                  ")"
''���ǉ�END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
'' 2007/09/04 SPK Tsutsumi Add End
            
'''���ǉ�START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
'            sql = "Insert into TBCMY007 " & _
'                  "(SXL_ID,SAMPLE_FROM,SAMPLE_TO,BLOCKID,HINBAN,KUBUN,TXID,REGDATE,SUMMITSENDFLAG,SENDFLAG,SENDDATE, " & _
'                  "MESDATA1TOP, MESDATA2TOP, MESDATA3TOP, MESDATA4TOP, MESDATA5TOP, " & _
'                  "MESDATA1BOT, MESDATA2BOT, MESDATA3BOT, MESDATA4BOT, MESDATA5BOT) " & _
'                  "values (" & _
'                  NoNullStr(SXLID) & ", " & _
'                  NoNullStr(smpId(1)) & ", " & _
'                  NoNullStr(smpId(2)) & ", " & _
'                  NoNullStr(BlkId) & ", " & _
'                  "(select distinct HINBCB||to_char(REVNUMCB,'FM00') from XSDCB where SXLIDCB=" & NoNullStr(SXLID) & "), " & _
'                  "'00', " & _
'                  "'TX853I', " & _
'                  "SYSDATE, '0', '0', SYSDATE, " & _
'                  NoNullStr(sRsData(1)) & ", " & _
'                  NoNullStr(sRsData(2)) & ", " & _
'                  NoNullStr(sRsData(3)) & ", " & _
'                  NoNullStr(sRsData(4)) & ", " & _
'                  NoNullStr(sRsData(5)) & ", " & _
'                  NoNullStr(sRsData(6)) & ", " & _
'                  NoNullStr(sRsData(7)) & ", " & _
'                  NoNullStr(sRsData(8)) & ", " & _
'                  NoNullStr(sRsData(9)) & ", " & _
'                  NoNullStr(sRsData(10)) & _
'                  ")"
'''���ǉ�END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
''���폜START SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
'            sql = "Insert into TBCMY007 " & _
'                  "(SXL_ID,SAMPLE_FROM,SAMPLE_TO,BLOCKID,HINBAN,KUBUN,TXID,REGDATE,SUMMITSENDFLAG,SENDFLAG,SENDDATE, " & _
'                  "MESDATA1TOP, MESDATA2TOP, MESDATA3TOP, MESDATA4TOP, MESDATA5TOP, " & _
'                  "MESDATA1BOT, MESDATA2BOT, MESDATA3BOT, MESDATA4BOT, MESDATA5BOT) " & _
'                  "values (" & _
'                  NoNullStr(SXLID) & ", " & _
'                  NoNullStr(smpId(1)) & ", " & _
'                  NoNullStr(smpId(2)) & ", " & _
'                  NoNullStr(BlkId) & ", " & _
'                  "(select distinct HINBAN||to_char(REVNUM,'FM00') from TBCME042 where SXLID=" & NoNullStr(SXLID) & "), " & _
'                  "'00', " & _
'                  "'TX853I', " & _
'                  "SYSDATE, '0', '0', SYSDATE, " & _
'                  NoNullStr(sRsData(1)) & ", " & _
'                  NoNullStr(sRsData(2)) & ", " & _
'                  NoNullStr(sRsData(3)) & ", " & _
'                  NoNullStr(sRsData(4)) & ", " & _
'                  NoNullStr(sRsData(5)) & ", " & _
'                  NoNullStr(sRsData(6)) & ", " & _
'                  NoNullStr(sRsData(7)) & ", " & _
'                  NoNullStr(sRsData(8)) & ", " & _
'                  NoNullStr(sRsData(9)) & ", " & _
'                  NoNullStr(sRsData(10)) & _
'                  ")"
''���폜END   SXL�Ǘ��iE042�j��XSDCB�@�\�ڍs '05/12/21 SMP���{
''''                  "(SXL_ID,SAMPLE_FROM,SAMPLE_TO,BLOCKID,HINBAN,KUBUN,TXID,REGDATE,SUMMITSENDFLAG,SENDFLAG,SENDDATE) " &
''''                  "SYSDATE, '0', '0', SYSDATE" &
            If OraDB.ExecuteSQL(sql) < 1 Then
                errmsg = "Y007"
                GoTo proc_exit
            End If

            '=================================================================================
            ' 2011/01/14 tkimura ADD START
            ''�ATBCMY007�擾��A��Top�ʒu���㗦,��Bot�ʒu���㗦,�����ΐ�,���R�l�����Ƃ߂�B
            If GetStandardPosRes(CRYNUM, _
                                 sRsData(), _
                                 sRsPos(), _
                                 dd) = FUNCTION_RETURN_FAILURE Then
                errmsg = errmsg & "���R�l���擾�ł��܂���"
                GoTo proc_exit
            End If
  
            '�B����Ώۈ��㗦,����ʒu���R�l�����߂�B
            '�����R�f�[�^�v�Z�A�X�V(Y011)    SuiteiResDataCalculation(SXLID[input�̂�],dd[input�̂�])

            If SuiteiResDataCalculation(SXLID, _
                                        dd, sUP_RATIO) = FUNCTION_RETURN_FAILURE Then
                errmsg = errmsg & "�����R�l�̌v�Z�Ɏ��s���܂����B"
                GoTo proc_exit
            End If
            
            '�DTBCMY011�e�[�u���̕i���V�X�e�����M�t���O���X�V����B
            If UpdateTBCMY011SendFlag(SXLID, HIN) = FUNCTION_RETURN_FAILURE Then
                errmsg = "Y011:" & XlSmpPos(i)
                GoTo proc_exit
            End If
            
        End If
' 2011/01/14 tkimura ADD END
'=================================================================================

    Next
'Add Start 2011/05/31 Y.Hitomi�@���グ���X�V
    sql = ""
    sql = sql & "UPDATE" & vbCrLf
    sql = sql & " TBCMX001" & vbCrLf
    sql = sql & "SET" & vbCrLf
    sql = sql & " UP_RATIO='" & sUP_RATIO(1) & "'" & vbCrLf  'SXL��BOT�ʒu�̈����グ��
    sql = sql & "WHERE" & vbCrLf
    sql = sql & " SXLID='" & SXLID & "'" & vbCrLf           'SXLID
        
    If OraDB.ExecuteSQL(sql) < 1 Then
        errmsg = "X001-" & i
        GoTo proc_exit
    End If
'Add End   2011/05/31 Y.Hitomi

'>>>>> 2011/06/00 SETsw)Marushita ���Ԕ����T���v�����M�Ή�
    Dim recCnt As Integer
    Dim iPosC2 As Integer
    '-------------------- XSDCW_1�̓ǂݍ��� ----------------------------------------
    sql = "select * from XSDCW_1 where SXLIDCW = '" & SXLID & "' and LIVKCW = '0' order by INPOSCW"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    ReDim recXSDCW_1(recCnt)
    ReDim recX007(recCnt)
    If recCnt > 0 Then
        For i = 1 To recCnt
            Set recXSDCW_1(i) = New c_cmzcrec
            recXSDCW_1(i).CopyFromRs "XSDCW_1", rs
            rs.MoveNext
        Next
    End If
    Set rs = Nothing
    
    If getTBCME036(HIN, iMCUTUNIT, sMSMPFLG) = FUNCTION_RETURN_FAILURE Then
        errmsg = "X007-Get_TBCME036"
        GoTo proc_exit
    End If
    
    If recCnt > 0 And sMSMPFLG = 1 Then
        For i = 1 To recCnt
            '-------------------- TBCMX007�Œ���f�[�^�ݒ� ----------------------------------------
            Set recX007(i) = New c_cmzcrec
            recX007(i).TABLENAME = "TBCMX007"
            recX007(i).SetRecDefault
            '�_�~�[������
            Set recX002(1) = New c_cmzcrec
            recX002(1).TABLENAME = "TBCMX002"
            recX002(1).SetRecDefault
            Set recX005(1) = New c_cmzcrec
            recX005(1).TABLENAME = "TBCMX005"
            recX005(1).SetRecDefault
        
            With recX007(i)
                .Fields("PLANTCAT").Value = HIN.sMukesaki                                       '���� 2007/09/04 SPK Tsutsumi Add
                .Fields("SXLID").Value = SXLID                                                  'SXLID
                .Fields("FROMTOKBN").Value = "C"                                                'FROMTO�敪(C�Œ�)
                .Fields("SAMPLE_ID").Value = recXSDCW_1(i)("REPSMPLIDCW").Value                 '�T���v��ID(��\)
                .Fields("BLOCKID").Value = Trim$(recXSDCW_1(i)("SMCRYNUMCW").Value)             '�u���b�NID
                .Fields("CRYNUM").Value = CRYNUM                                                '�����ԍ�
                .Fields("SXLDECDATE").Value = nowtime                                           'SXL-ID�m����t
                .nowtime = nowtime
            
                'Cng Start 2011/07/11 Y.Hitomi �u���b�N���ʒu�́AXSDC2�̈ʒu���擾���A�������ʒu�������
                iPosC2 = getXSDC2Pos(Trim$(recXSDCW_1(i)("SMCRYNUMCW").Value))
                .Fields("BLOCKPOS").Value = recXSDCW_1(i)("INPOSCW").Value - iPosC2             '�u���b�N�������ʒu
                '.Fields("BLOCKPOS").Value = recXSDCW_1(i)("INPOSCW").Value                      '�u���b�N�������ʒu
                'Cng End   2011/07/11 Y.Hitomi
                
                .Fields("INGOTPOS").Value = recXSDCW_1(i)("INPOSCW").Value                      '�������J�n�ʒu
                .Fields("HINBAN").Value = HIN.hinban                                            '�i��
                .Fields("REVNUM").Value = HIN.mnorevno                                          '���i�ԍ������ԍ�
                .Fields("FACTORY").Value = HIN.factory                                          '�H��
                .Fields("OPECOND").Value = HIN.opecond                                          '���Ə���
                .Fields("MCUTUNIT").Value = iMCUTUNIT                                           '���Ԕ����P��
                
                '-------------------- TBCME036�f�[�^�擾�ݒ� ----------------------------------------
'                .Fields("PRODCOND").Value = recE037("PRODCOND").Value                           '�������
'                .Fields("PGID").Value = Mid(recE037("PGID"), 1, 8)                              'PG-ID
'                .Fields("UPLENGTH").Value = recE037("UPLENGTH").Value                           '���グ����
'                .Fields("SXLPOS").Value = 0                                                     'SXL�ʒu
'                .Fields("SXLLENGTH").Value = XlSmpPos(2) - XlSmpPos(1)                          'SXL-ID�m�蒷��
'                .Fields("SXLWAFERCNT").Value = WfCnt                                            'SXL-ID�m�莞��WF����
'                .Fields("FREELENG").Value = recE037("FREELENG").Value                           '�t���[��
'                .Fields("DIAMETER").Value = recE037("DIAMETER").Value                           '���a
'                .Fields("SEED").Value = recE037("SEED").Value                                   '�V�[�h
                '-------------------- (WFOi)WFOi����(TBCMY013)�f�[�^�擾�ݒ� ----------------------------------------
                If getTBCMY013WFOi(recXSDCW_1(i), HIN, sPos, recX007(i)) = FUNCTION_RETURN_FAILURE Then
                    errmsg = "Y013-Oi:" & XlSmpPos(i)
                    GoTo proc_exit
                End If
                '-------------------- (WFRs)WFRs����(TBCMY013)�f�[�^�擾�ݒ� ----------------------------------------
                If getTBCMY013WFRs(recXSDCW_1(i), HIN, sPos, recX007(i)) = FUNCTION_RETURN_FAILURE Then
                    errmsg = "Y013-Rs:" & XlSmpPos(i)
                    GoTo proc_exit
                End If
                '-------------------- (WFDOi1�`3)WFDOi����(TBCMY013)�f�[�^�擾�ݒ� ----------------------------------------
                For j = 1 To 3
                    If getTBCMY013WFDOi(recXSDCW_1(i), j, recX007(i), recX002(1)) = FUNCTION_RETURN_FAILURE Then
                        errmsg = "Y013-DOi" & j & ":" & XlSmpPos(i)
                        GoTo proc_exit
                    End If
                Next
                '-------------------- (WFOSF1�`4)WFOSF����(TBCMY013)�f�[�^�擾�ݒ� ----------------------------------------
                For j = 1 To 3
                    If getTBCMY013WFOSF(recXSDCW_1(i), j, HIN, sPos, recX007(i), recX002(1), recX007(i).TABLENAME) = FUNCTION_RETURN_FAILURE Then
                        errmsg = "Y013-OSF" & j & ":" & XlSmpPos(i)
                        GoTo proc_exit
                    End If
                Next
                '-------------------- (WFBMD1�`3)WFBMD����(TBCMY013)�f�[�^�擾�ݒ� ----------------------------------------
                For j = 1 To 3
                    If getTBCMY013WFBMD(recXSDCW_1(i), j, HIN, sPos, recX007(i), recX002(1)) = FUNCTION_RETURN_FAILURE Then
                        errmsg = "Y013-BMD" & j & ":" & XlSmpPos(i)
                        GoTo proc_exit
                    End If
                Next
                '-------------------- (WFDSOD)WFDSOD����(TBCMY013)�f�[�^�擾�ݒ� ----------------------------------------
                If getTBCMY013WFDSOD(recXSDCW_1(i), recX007(i), recX002(1)) = FUNCTION_RETURN_FAILURE Then
                    errmsg = "Y013-DSOD:" & XlSmpPos(i)
                    GoTo proc_exit
                End If
                '-------------------- (WFDZ)WFDZ����(TBCMY013)�f�[�^�擾�ݒ� ----------------------------------------
                If getTBCMY013WFDZ(recXSDCW_1(i), HIN, recX007(i), recX002(1)) = FUNCTION_RETURN_FAILURE Then
                    errmsg = "Y013-DZ:" & XlSmpPos(i)
                    GoTo proc_exit
                End If
                ''�c���_�f���ю擾�ǉ��@03/12/19 ooba START ================================================>
                '-------------------- (WFAOi)WFAOi����(TBCMY013)�f�[�^�擾�ݒ� ----------------------------------------
                If getTBCMY013WFAOi(recXSDCW_1(i), recX007(i), recX002(1)) = FUNCTION_RETURN_FAILURE Then
                    errmsg = "Y013-AOi:" & XlSmpPos(i)
                    GoTo proc_exit
                End If
                ''XSDCW��GD��s�]���̏�ԃt���O=1�����уt���O=1�̏ꍇTBCMX007�ɓo�^����B
                If (recXSDCW_1(i)("WFINDGDCW").Value <> "0") And (recXSDCW_1(i)("WFRESGDCW").Value <> "0") Then
                    '-------------------- (GD)GD����(TBCMJ015)�f�[�^�擾�ݒ� ----------------------------------------
                    If getTBCMJ015WFGD(CRYNUM, recXSDCW_1(i), recX007(i), recX007(i).TABLENAME) = FUNCTION_RETURN_FAILURE Then
                        errmsg = "J015-GD:" & XlSmpPos(i)
                        GoTo proc_exit
                    End If
                End If
                '-------------------- ���OSF1�`3����(TBCMY022)�ް��擾�ݒ� ---------------------------
                For j = 1 To 3
                    If getTBCMY022EPOSF(recXSDCW_1(i), j, HIN, sPos, recX007(i), recX005(1), recX007(i).TABLENAME) = FUNCTION_RETURN_FAILURE Then
                        errmsg = "Y022-EPOSF" & j & ":" & XlSmpPos(i)
                        GoTo proc_exit
                    End If
                Next j
                '-------------------- ���BMD1�`3����(TBCMY022)�ް��擾�ݒ� ---------------------------
                For j = 1 To 3
                    If getTBCMY022EPBMD(recXSDCW_1(i), j, HIN, sPos, recX007(i), recX005(1)) = FUNCTION_RETURN_FAILURE Then
                        errmsg = "Y022-EPBMD" & j & ":" & XlSmpPos(i)
                        GoTo proc_exit
                    End If
                Next j
                '==============================================
                '�@TBCMX007 �ɏ�������
                '==============================================
                .Fields("REGDATE").Value = "SYSDATE"                                    '�o�^���t
                .Fields("SENDFLAG").Value = "3"  ':2008/03/31 �� �@���sSXL�m����s������SXL�}�b�v��M���ɑ��M���ݸޕύX����B
                .Fields("SENDDATE").Value = "SYSDATE"                                   '���M���t
                sql = .SqlInsert
                If OraDB.ExecuteSQL(sql) < 1 Then
                    errmsg = "X007-" & i
                    GoTo proc_exit
                End If
            End With
        Next
    End If
    
    WriteX00n = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    WriteX00n = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

Private Function NoNullStr(s$) As String
    If s = vbNullString Then
        NoNullStr = "' '"
    Else
        NoNullStr = "'" & s & "'"
    End If
End Function

'���������i�u���b�N�j�O�H�����ю擾���\���̍쐬 2002/09/03 ADD hitec)N.MATSUMOTO
Public Function cmbc040_CreateXSDC2(ByVal iBlockCnt As Integer, ByRef bNoData As Boolean) As FUNCTION_RETURN

    Dim intLoopCnt  As Integer
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim intProcNo   As Integer
    Dim iRtn    As Integer
    Dim dblDiameter As Double
    Dim intNum  As Integer

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    intProcNo = 0
    bNoData = False
    '�u���b�NID�𓾂�
    sql = " SELECT * FROM XSDC2" '
    sql = sql & " WHERE CRYNUMC2='" & strBlockID(iBlockCnt) & "'"
''''    sql = sql & "   AND NEWKNTC2='" & BeforeProc & "'"
    sql = sql & "   AND LIVKC2= '0'"   '�����敪

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
        rs.Close
        bNoData = True
        cmbc040_CreateXSDC2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    If rs.EOF = False Then
        '�O�H���擾
        With BlkOld
            If IsNull(rs.Fields("CRYNUMC2")) = False Then .CRYNUMC2 = rs.Fields("CRYNUMC2")
            If IsNull(rs.Fields("KCNTC2")) = False Then .KCNTC2 = rs.Fields("KCNTC2")       '�H���A��
            If IsNull(rs.Fields("XTALC2")) = False Then .XTALC2 = rs.Fields("XTALC2")
            If IsNull(rs.Fields("INPOSC2")) = False Then .INPOSC2 = rs.Fields("INPOSC2")
            If IsNull(rs.Fields("NEKKNTC2")) = False Then .NEKKNTC2 = rs.Fields("NEKKNTC2")
            If IsNull(rs.Fields("NEWKNTC2")) = False Then .NEWKNTC2 = rs.Fields("NEWKNTC2")
            If IsNull(rs.Fields("NEWKKBC2")) = False Then .NEWKKBC2 = rs.Fields("NEWKKBC2")
            If IsNull(rs.Fields("NEMACOC2")) = False Then .NEMACOC2 = rs.Fields("NEMACOC2")
            If IsNull(rs.Fields("GNKKNTC2")) = False Then .GNKKNTC2 = rs.Fields("GNKKNTC2")
            If IsNull(rs.Fields("GNWKNTC2")) = False Then .GNWKNTC2 = rs.Fields("GNWKNTC2")
            If IsNull(rs.Fields("GNWKKBC2")) = False Then .GNWKKBC2 = rs.Fields("GNWKKBC2")
            If IsNull(rs.Fields("GNMACOC2")) = False Then .GNMACOC2 = rs.Fields("GNMACOC2")
            If IsNull(rs.Fields("GNDAYC2")) = False Then .GNDAYC2 = rs.Fields("GNDAYC2")
            If IsNull(rs.Fields("GNLC2")) = False Then .GNLC2 = rs.Fields("GNLC2")          '���ݒ���
            If IsNull(rs.Fields("GNWC2")) = False Then .GNWC2 = rs.Fields("GNWC2")          '���ݏd��
            If IsNull(rs.Fields("GNMC2")) = False Then .GNMC2 = rs.Fields("GNMC2")          '���ݖ���
            If IsNull(rs.Fields("SUMITLC2")) = False Then .SUMITLC2 = rs.Fields("SUMITLC2")
            If IsNull(rs.Fields("SUMITWC2")) = False Then .SUMITWC2 = rs.Fields("SUMITWC2")
            If IsNull(rs.Fields("SUMITMC2")) = False Then .SUMITMC2 = rs.Fields("SUMITMC2")
            If IsNull(rs.Fields("CHGC2")) = False Then .CHGC2 = rs.Fields("CHGC2")
            If IsNull(rs.Fields("KAKOUBC2")) = False Then .KAKOUBC2 = rs.Fields("KAKOUBC2")
            If IsNull(rs.Fields("KEIDAYC2")) = False Then .KEIDAYC2 = rs.Fields("KEIDAYC2")
            If IsNull(rs.Fields("GNTKUBC2")) = False Then .GNTKUBC2 = rs.Fields("GNTKUBC2")
            If IsNull(rs.Fields("GNTNOC2")) = False Then .GNTNOC2 = rs.Fields("GNTNOC2")
            If IsNull(rs.Fields("XTWORKC2")) = False Then .XTWORKC2 = rs.Fields("XTWORKC2")
            If IsNull(rs.Fields("WFWORKC2")) = False Then .WFWORKC2 = rs.Fields("WFWORKC2")
            If IsNull(rs.Fields("LSTATBC2")) = False Then .LSTATBC2 = rs.Fields("LSTATBC2")
            If IsNull(rs.Fields("RSTATBC2")) = False Then .RSTATBC2 = rs.Fields("RSTATBC2")
            If IsNull(rs.Fields("LUFRCC2")) = False Then .LUFRCC2 = rs.Fields("LUFRCC2")
            If IsNull(rs.Fields("LUFRBC2")) = False Then .LUFRBC2 = rs.Fields("LUFRBC2")
            If IsNull(rs.Fields("LDFRCC2")) = False Then .LDFRCC2 = rs.Fields("LDFRCC2")
            If IsNull(rs.Fields("LDFRBC2")) = False Then .LDFRBC2 = rs.Fields("LDFRBC2")
            If IsNull(rs.Fields("HOLDCC2")) = False Then .HOLDCC2 = rs.Fields("HOLDCC2")
            If IsNull(rs.Fields("HOLDBC2")) = False Then .HOLDBC2 = rs.Fields("HOLDBC2")
            If IsNull(rs.Fields("EXKUBC2")) = False Then .EXKUBC2 = rs.Fields("EXKUBC2")
            If IsNull(rs.Fields("HENPKC2")) = False Then .HENPKC2 = rs.Fields("HENPKC2")
            If IsNull(rs.Fields("LIVKC2")) = False Then .LIVKC2 = rs.Fields("LIVKC2")
            If IsNull(rs.Fields("KANKC2")) = False Then .KANKC2 = rs.Fields("KANKC2")
            If IsNull(rs.Fields("NFC2")) = False Then .NFC2 = rs.Fields("NFC2")
            If IsNull(rs.Fields("SAKJC2")) = False Then .SAKJC2 = rs.Fields("SAKJC2")
            If IsNull(rs.Fields("TDAYC2")) = False Then .TDAYC2 = rs.Fields("TDAYC2")
            If IsNull(rs.Fields("KDAYC2")) = False Then .KDAYC2 = rs.Fields("KDAYC2")
            If IsNull(rs.Fields("SUMITBC2")) = False Then .SUMITBC2 = rs.Fields("SUMITBC2")
            If IsNull(rs.Fields("SNDKC2")) = False Then .SNDKC2 = rs.Fields("SNDKC2")
            If IsNull(rs.Fields("SNDDAYC2")) = False Then .SNDDAYC2 = rs.Fields("SNDDAYC2")
            If IsNull(rs.Fields("PLANTCATC2")) = False Then .PLANTCATC2 = rs.Fields("PLANTCATC2")   ' 2007/09/04 SPK Tsutsumi Add
        End With

        '���������i��ۯ��j�̑O�H�����A���������i��ۯ��j�̌��ݍH���փR�s�[
        BlkNow = BlkOld

        '���������i��ۯ��j���ݍH���̕ҏW
        With BlkNow
            .KCNTC2 = CInt(.KCNTC2) + 1     '�H���A��]
            'Cng Start 2010/09/02 Y.Hitomi
            ''�u���b�N��SXL��1�ł��������Ă����ꍇ�A�H���R�[�h���X�V���Ȃ��悤�ɂ���
            If .GNWKNTC2 <> "     " Then
                .NEWKNTC2 = Kihon.NOWPROC        '�O�H��
                .GNWKNTC2 = "CW800"              '���ݍH��
            End If
            
'            .NEWKNTC2 = Kihon.NOWPROC           '�O�H��
'            .GNWKNTC2 = Kihon.NEWPROC           '���ݍH��
            'Cng End 2010/09/02 Y.Hitomi

            '���ݏd�ʂ����߂�
            If GetDiameter(strBlockID(iBlockCnt), dblDiameter) = FUNCTION_RETURN_FAILURE Then  '���a�����߂�
                dblDiameter = 0
''''                GoTo proc_wxit
            End If
            Kihon.DIAMETER = dblDiameter
            '�擾�������a�����ɏd�ʂ����߂�
''''            .GNWC2 = CStr(WeightOfCylinder(dblDiameter, CDbl(.GNLC2)))
''''            '���ݖ��������߂�
''''            If WfCount(strBlockID(iBlockCnt), CInt(.GNLC2), intNum) = FUNCTION_RETURN_FAILURE Then
''''                .GNMC2 = 0
''''''''                GNMCA proc_wxit
''''            Else
''''                .GNMC2 = intNum
''''            End If
            .SUMITBC2 = "0"
            .SUMITLC2 = "0"
            .SUMITMC2 = "0"
            .SUMITWC2 = "0"
        End With

    End If
    
    rs.Close

    cmbc040_CreateXSDC2 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    cmbc040_CreateXSDC2 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2002/09/03 ADD hitec)N.MATSUMOTO

'���������i�i�ԁj�O�H�����ю擾���\���̍쐬 2002/09/03 ADD hitec)N.MATSUMOTO
'Cng Start 2010/10/14 Y.Hitomi
'Public Function cmbc040_CreateXSDCA(ByVal iBlockCnt As Integer, ByRef bNoData As Boolean) As FUNCTION_RETURN
Public Function cmbc040_CreateXSDCA(ByVal iBlockCnt As Integer, ByRef bNoData As Boolean, strSxlId As String) As FUNCTION_RETURN
'Cng End   2010/10/14 Y.Hitomi
    Dim iLoopCnt    As Integer
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim intProcNo   As Integer
    Dim dblDiameter As Double
    Dim intNum  As Integer

    '' �G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    intProcNo = 0
    bNoData = False
    '�u���b�NID�𓾂�
    sql = "SELECT * from XSDCA"
    sql = sql & " WHERE CRYNUMCA='" & strBlockID(iBlockCnt) & "'"
''''    sql = sql & "   AND NEWKNTCA='" & BeforeProc & "'"
    sql = sql & "   AND LIVKCA= '0'"   '�����敪

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        bNoData = True
        cmbc040_CreateXSDCA = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    rs.MoveFirst
    iLoopCnt = 0
    
    Do While Not rs.EOF
        ReDim Preserve HinOld(iLoopCnt)
        ReDim Preserve HinNow(iLoopCnt)
        With HinOld(iLoopCnt)
            If IsNull(rs.Fields("CRYNUMCA")) = False Then .CRYNUMCA = rs.Fields("CRYNUMCA")
            If IsNull(rs.Fields("HINBCA")) = False Then .HINBCA = rs.Fields("HINBCA")
            If IsNull(rs.Fields("INPOSCA")) = False Then .INPOSCA = rs.Fields("INPOSCA")
            If IsNull(rs.Fields("REVNUMCA")) = False Then .REVNUMCA = rs.Fields("REVNUMCA")
            If IsNull(rs.Fields("FACTORYCA")) = False Then .FACTORYCA = rs.Fields("FACTORYCA")
            If IsNull(rs.Fields("OPECA")) = False Then .OPECA = rs.Fields("OPECA")
            If IsNull(rs.Fields("KCKNTCA")) = False Then .KCKNTCA = rs.Fields("KCKNTCA")
            If IsNull(rs.Fields("SXLIDCA")) = False Then .SXLIDCA = rs.Fields("SXLIDCA")
            If IsNull(rs.Fields("XTALCA")) = False Then .XTALCA = rs.Fields("XTALCA")
            If IsNull(rs.Fields("NEKKNTCA")) = False Then .NEKKNTCA = rs.Fields("NEKKNTCA")
            If IsNull(rs.Fields("NEWKNTCA")) = False Then .NEWKNTCA = rs.Fields("NEWKNTCA")
            If IsNull(rs.Fields("NEWKKBCA")) = False Then .NEWKKBCA = rs.Fields("NEWKKBCA")
            If IsNull(rs.Fields("NEMACOCA")) = False Then .NEMACOCA = rs.Fields("NEMACOCA")
            If IsNull(rs.Fields("GNKKNTCA")) = False Then .GNKKNTCA = rs.Fields("GNKKNTCA")
            If IsNull(rs.Fields("GNWKNTCA")) = False Then .GNWKNTCA = rs.Fields("GNWKNTCA")
            If IsNull(rs.Fields("GNWKKBCA")) = False Then .GNWKKBCA = rs.Fields("GNWKKBCA")
            If IsNull(rs.Fields("GNMACOCA")) = False Then .GNMACOCA = rs.Fields("GNMACOCA")
            If IsNull(rs.Fields("GNDAYCA")) = False Then .GNDAYCA = rs.Fields("GNDAYCA")
            If IsNull(rs.Fields("GNLCA")) = False Then .GNLCA = rs.Fields("GNLCA")
            If IsNull(rs.Fields("GNWCA")) = False Then .GNWCA = rs.Fields("GNWCA")
            If IsNull(rs.Fields("GNMCA")) = False Then .GNMCA = rs.Fields("GNMCA")
            If IsNull(rs.Fields("SUMITLCA")) = False Then .SUMITLCA = rs.Fields("SUMITLCA")
            If IsNull(rs.Fields("SUMITWCA")) = False Then .SUMITWCA = rs.Fields("SUMITWCA")
            If IsNull(rs.Fields("SUMITMCA")) = False Then .SUMITMCA = rs.Fields("SUMITMCA")
            If IsNull(rs.Fields("CHGCA")) = False Then .CHGCA = rs.Fields("CHGCA")
            If IsNull(rs.Fields("KAKOUBCA")) = False Then .KAKOUBCA = rs.Fields("KAKOUBCA")
            If IsNull(rs.Fields("KEIDAYCA")) = False Then .KEIDAYCA = rs.Fields("KEIDAYCA")
            If IsNull(rs.Fields("GNTKUBCA")) = False Then .GNTKUBCA = rs.Fields("GNTKUBCA")
            If IsNull(rs.Fields("GNTNOCA")) = False Then .GNTNOCA = rs.Fields("GNTNOCA")
            If IsNull(rs.Fields("XTWORKCA")) = False Then .XTWORKCA = rs.Fields("XTWORKCA")
            If IsNull(rs.Fields("WFWORKCA")) = False Then .WFWORKCA = rs.Fields("WFWORKCA")
            If IsNull(rs.Fields("LSTATBCA")) = False Then .LSTATBCA = rs.Fields("LSTATBCA")
            If IsNull(rs.Fields("RSTATBCA")) = False Then .RSTATBCA = rs.Fields("RSTATBCA")
            If IsNull(rs.Fields("LUFRCCA")) = False Then .LUFRCCA = rs.Fields("LUFRCCA")
            If IsNull(rs.Fields("LUFRBCA")) = False Then .LUFRBCA = rs.Fields("LUFRBCA")
            If IsNull(rs.Fields("LDFRCCA")) = False Then .LDFRCCA = rs.Fields("LDFRCCA")
            If IsNull(rs.Fields("LDFRBCA")) = False Then .LDFRBCA = rs.Fields("LDFRBCA")
            If IsNull(rs.Fields("HOLDCCA")) = False Then .HOLDCCA = rs.Fields("HOLDCCA")
            If IsNull(rs.Fields("HOLDBCA")) = False Then .HOLDBCA = rs.Fields("HOLDBCA")
            If IsNull(rs.Fields("EXKUBCA")) = False Then .EXKUBCA = rs.Fields("EXKUBCA")
            If IsNull(rs.Fields("HENPKCA")) = False Then .HENPKCA = rs.Fields("HENPKCA")
            If IsNull(rs.Fields("LIVKCA")) = False Then .LIVKCA = rs.Fields("LIVKCA")
            If IsNull(rs.Fields("KANKCA")) = False Then .KANKCA = rs.Fields("KANKCA")
            If IsNull(rs.Fields("NFCA")) = False Then .NFCA = rs.Fields("NFCA")
            If IsNull(rs.Fields("SAKJCA")) = False Then .SAKJCA = rs.Fields("SAKJCA")
            If IsNull(rs.Fields("TDAYCA")) = False Then .TDAYCA = rs.Fields("TDAYCA")
            If IsNull(rs.Fields("KDAYCA")) = False Then .KDAYCA = rs.Fields("KDAYCA")
            If IsNull(rs.Fields("SUMITBCA")) = False Then .SUMITBCA = rs.Fields("SUMITBCA")
            If IsNull(rs.Fields("SNDKCA")) = False Then .SNDKCA = rs.Fields("SNDKCA")
            If IsNull(rs.Fields("SNDDAYCA")) = False Then .SNDDAYCA = rs.Fields("SNDDAYCA")
            If IsNull(rs.Fields("PLANTCATCA")) = False Then .PLANTCATCA = rs.Fields("PLANTCATCA")   ' 2007/09/04 SPK Tsutsumi Add
        End With
        
        '�O�H���̍\���̂����ݍH���̍\���̂փR�s�[
        HinNow(iLoopCnt) = HinOld(iLoopCnt)
        
        '���ݍH���\���̂̍H���A�Ԃ̕ҏW
        With HinNow(iLoopCnt)
            .KCKNTCA = BlkNow.KCNTC2
            'Cng Start 2010/10/14 Y.Hitomi
            '���s�w��SXLID�̂ݍH���R�[�h��ύX���A����ȊO�́A�O�H���������p��
            If strSxlId = .SXLIDCA Then
                .NEWKNTCA = Kihon.NOWPROC             '�O�H���R�[�h���ŏI�ʉߍH���ɃZ�b�g
                .GNWKNTCA = Kihon.NEWPROC             '���ݍH���R�[�h�����ݍH���փZ�b�g
            Else
                .NEWKNTCA = rs.Fields("NEWKNTCA")     '�O�H���R�[�h���ŏI�ʉߍH���ɃZ�b�g
                .GNWKNTCA = rs.Fields("GNWKNTCA")     '���ݍH���R�[�h�����ݍH���փZ�b�g
            End If
'            .NEWKNTCA = Kihon.NOWPROC       '���ݍH��
'            .GNWKNTCA = Kihon.NEWPROC       '���H��
            'Cng End   2010/10/14 Y.Hitomi
            
            '���ݏd�ʂ����߂�
            If GetDiameter(strBlockID(iBlockCnt), dblDiameter) = FUNCTION_RETURN_FAILURE Then  '���a�����߂�
                dblDiameter = 0
''''                GoTo proc_wxit
            End If
            '�擾�������a�����ɏd�ʂ����߂�
''''            HinNow(iLoopCnt).GNWCA = CStr(WeightOfCylinder(dblDiameter, CDbl(.GNLCA)))
'''''            '���ݖ��������߂�
'''''            If WfCount(strBlockID(iBlockCnt), CInt(.GNLCA), intNum) = FUNCTION_RETURN_FAILURE Then
'''''                .GNMCA = 0
'''''''''                GNMCA proc_wxit
'''''            Else
'''''                .GNMCA = intNum
'''''            End If
            .SUMITBCA = "0"
            .SUMITLCA = HinOld(iLoopCnt).SUMITLCA    ''03/05/13 �㓡
            .SUMITMCA = HinOld(iLoopCnt).SUMITMCA    ''03/05/14 �㓡
            .SUMITWCA = HinOld(iLoopCnt).SUMITWCA    ''03/05/13 �㓡
        End With
        
        iLoopCnt = iLoopCnt + 1
        rs.MoveNext
    Loop
    
    With Kihon  '��{���@�i�ԍ\���̂̃J�E���g���Z�b�g
        .CNTHINOLD = iLoopCnt
        .CNTHINNOW = iLoopCnt
    End With
    
    rs.Close
    cmbc040_CreateXSDCA = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    cmbc040_CreateXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2002/09/03 ADD hitec)N.MATSUMOTO

'�\���̍쐬���� 2002/09/03 ADD hitec)N.MATSUMOTO
Public Function CreateTable(ByVal strSxlId As String, ByRef strErrMsg As String) As FUNCTION_RETURN

    Dim sql     As String
    Dim rsMain  As OraDynaset
    Dim iBlockCnt   As Integer
    Dim strDBName   As String
    Dim bNoData     As Boolean
    Dim sTmpSxl() As String     '�d�|�H���������pSXLID�@06/03/14 ooba

    On Error GoTo proc_err

    bNoData = False
    '�u���b�NID�擾
    strDBName = "XSDCA"
    sql = "select DISTINCT(CRYNUMCA) from XSDCA " & _
          "where SXLIDCA='" & strSxlId & "' " & _
          "  and LIVKCA= '0'"
    Set rsMain = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rsMain.RecordCount = 0 Then
        Debug.Print "XSDC2�F�O�H�����і���"
        Debug.Print sql
        rsMain.Close
        CreateTable = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If

    sProSXLID = strSxlId    '����SXLID��ā@06/03/24 ooba
    
    '�u���b�NID�擾
''''    bNoData = False
''''    strDBName = "E040"
''''    sql = "select BLOCKID from TBCME040 " & _
''''          "where (crynum='" & strCrynum & "') and (INGOTPOS<=" & intIngotpos & ") and (" & intIngotpos & "<INGOTPOS+LENGTH)"
''''    Set rsMain = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''''    If rsMain.RecordCount = 0 Then
''''        rsMain.Close
''''        strErrMsg = GetMsgStr("EAPLY") & strDBName
''''        GoTo proc_exit
''''    End If
    
    '�d�|�H���ă`�F�b�N�@�\�ǉ��@06/03/14 ooba
    ReDim sTmpSxl(1)
    sTmpSxl(1) = strSxlId
    If DBDRV_CheckCodeXSDCB(sTmpSxl, PROCD_SXL_KAKUTEI, strErrMsg) = FUNCTION_RETURN_FAILURE Then
        CreateTable = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    iBlockCnt = 0
    
    Do While Not rsMain.EOF
        iBlockCnt = iBlockCnt + 1
        ReDim strBlockID(iBlockCnt)
        strBlockID(iBlockCnt) = rsMain("CRYNUMCA")
        With Kihon
            .StaffID = Trim(f_cmbc040_1.txtStaffID.text)    '�S���҃R�[�h
''''            BeforeProc = PROCD_WFC_SOUGOUHANTEI '�O�H��
            .NEWPROC = PROCD_SXL_MAP
            .NOWPROC = PROCD_SXL_KAKUTEI
            .DIAMETER = 0   '------------------�ۗ�
            .ALLSCRAP = "N"     '�S���X�N���b�v����
            .FURYOUMU = "N"       '�s�ǖ���
        End With
        
        '���������i�u���b�N�j����O�H�����ю擾
        strDBName = "XSDC2"
        If cmbc040_CreateXSDC2(iBlockCnt, bNoData) = FUNCTION_RETURN_FAILURE Then
            If bNoData = True Then
'                CreateTable = FUNCTION_RETURN_SUCCESS
                CreateTable = FUNCTION_RETURN_FAILURE       '07/02/06 ooba
                strErrMsg = GetMsgStr("EAPLY") & strDBName  '07/02/06 ooba
                Debug.Print "cmbc040_CreateXSDC2(" & iBlockCnt & "," & bNoData & ")�FXSDC2�O�H�����і���"
                GoTo proc_exit
            Else
                CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EAPLY") & strDBName
                Debug.Print "cmbc040_CreateXSDC2(" & iBlockCnt & "," & bNoData & ")�FXSDC2�O�H�����ѓǂݍ��݃G���["
                GoTo proc_exit
            End If
        End If
        
        '���������i�i�ԁj����O�H�����ю擾
        strDBName = "XSDCA"
        'Cng Start 2010/10/14 Y.Hitomi
'        If cmbc040_CreateXSDCA(iBlockCnt, bNoData) = FUNCTION_RETURN_FAILURE Then
        If cmbc040_CreateXSDCA(iBlockCnt, bNoData, strSxlId) = FUNCTION_RETURN_FAILURE Then
        'Cng End   2010/10/14 Y.Hitomi
            If bNoData = True Then
'                CreateTable = FUNCTION_RETURN_SUCCESS
                CreateTable = FUNCTION_RETURN_FAILURE       '07/02/06 ooba
                strErrMsg = GetMsgStr("EAPLY") & strDBName  '07/02/06 ooba
                Debug.Print "cmbc040_CreateXSDCA(" & iBlockCnt & "," & bNoData & ")�FXSDCA�O�H�����і��� "
                GoTo proc_exit
            Else
                CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EAPLY") & strDBName
                Debug.Print "cmbc040_CreateXSDCA(" & iBlockCnt & "," & bNoData & ")�FXSDCA�O�H�����ѓǂݍ��݃G���["
                GoTo proc_exit
            End If
        End If
        
        '��{����
        If KihonProc = FUNCTION_RETURN_FAILURE Then
            CreateTable = FUNCTION_RETURN_FAILURE           '08/04/04 ooba
            strErrMsg = GetMsgStr("EAPLY")
            Debug.Print "KihonProc()�F��{�����Ɏ��s���܂���"
            Exit Function
        End If
        
        rsMain.MoveNext
    Loop
    rsMain.Close
                
proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    CreateTable = FUNCTION_RETURN_FAILURE
    Resume proc_exit
                
End Function
'2002/09/03 ADD hitec)N.MATSUMOTO

'###  WF�ʐM�����i��{�����p�����[�^�쐬�j ### '2002/09/03 ADD hitec)N.MATSUMOTO    Start
'Public Function MakeParameter() As FUNCTION_RETURN
'f_cmbc040_1.frm��s_cmbc040_SQL.bas�֐��ړ�,record�ǉ��@06/10/20 ooba
Public Function MakeParameter(record As typ_cmlc001e_Disp) As FUNCTION_RETURN
    Dim lng     As Long
    Dim dat     As Variant
    Dim lRowCnt As Long
    Dim rsMain      As OraDynaset
    Dim sql     As String
    Dim iBlockCnt   As Integer
    Dim sErrMsg As String

'    For lRowCnt = 1 To f_cmbc040_1.sprList.MaxRows
'        With rec(lRowCnt)
        With record     '06/10/20 ooba
            If .HLDCLASS = "0" Then     '�z�[���h�`�F�b�N��OFF�̏ꍇ
                strSxlData = .SXLID     '03/05/01 Add.�㓡
                If CreateTable(.SXLID, sErrMsg) = FUNCTION_RETURN_FAILURE Then
                    MakeParameter = FUNCTION_RETURN_FAILURE
                    f_cmbc040_1.lblMsg.Caption = sErrMsg
                    GoTo proc_exit
                End If
            End If
        End With
'    Next
    MakeParameter = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

End Function
'2002/09/03 ADD hitec)N.MATSUMOTO    End

'2003/10/19 �g���ĂȂ��̂ō폜 SystemBrain ==========================================================��
''''''�T�v      :Cs���уf�[�^�̎擾�h���C�o
''''''���Ұ��@�@:�ϐ���          , IO , �^               , ����
''''''          :sxlid           , I  ,String            , SXLID
''''''          :Cs()            , I  ,tCsData           , ����Cs���茋��
''''''      �@�@:�߂�l          , O  , FUNCTION_RETURN�@, �ǂݍ��݂̐���
''''''����      :Cs�̏㉺���т��擾����
''''''����      :2002/10/03 �쑺 �쐬
'''''Private Function getSXLCs(SXLID$, Cs() As tCsData) As FUNCTION_RETURN
'''''    Dim rs As OraDynaset
'''''    Dim sql As String
'''''    Dim CRYNUM As String
'''''    Dim sxlFrom As Integer
'''''    Dim sxlLen As Integer
'''''    Dim SpecCsMin As Double
'''''    Dim specCsH As String
'''''
'''''    '' �G���[�n���h���̐ݒ�
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmzcF_cmlc001e_SQL.bas -- Function getSXLCs"
'''''    getSXLCs = FUNCTION_RETURN_FAILURE
'''''
'''''    '���я�����
'''''    With Cs(1)
'''''        .SXL_CS_SMPPOS = -1
'''''        .SXLCS_CSMEAS = -1
'''''        .SXLCS_70PPRE = -1
'''''    End With
'''''    With Cs(2)
'''''        .SXL_CS_SMPPOS = -1
'''''        .SXLCS_CSMEAS = -1
'''''        .SXLCS_70PPRE = -1
'''''    End With
'''''
'''''    '�����ԍ�,SXL�͈�,SXL�i�Ԃ�Cs�d�l(�����l)���擾
'''''    sql = "select CRYNUM, INGOTPOS, LENGTH, HSXCNMIN, HSXCNHWS "
'''''    sql = sql & "from TBCME019 SPEC, TBCME042 SXL "
'''''    sql = sql & "where SXL.SXLID='" & SXLID & "'"
'''''    sql = sql & "  and SPEC.HINBAN=SXL.HINBAN and SPEC.MNOREVNO=SXL.REVNUM and SPEC.FACTORY=SXL.FACTORY and SPEC.OPECOND=SXL.OPECOND"
'''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    If rs.RecordCount > 0 Then
'''''        CRYNUM = rs("CRYNUM")
'''''        sxlFrom = rs("INGOTPOS")
'''''        sxlLen = rs("LENGTH")
'''''        SpecCsMin = rs("HSXCNMIN")
'''''        specCsH = rs("HSXCNHWS")
'''''    Else
'''''        GoTo proc_exit  'SXL�������͕i�Ԏd�l�Ȃ�
'''''    End If
'''''    rs.Close
'''''    '�d�l����iH�@OR�@S�j�̏ꍇ�̂ݎ��т���������
'''''    If specCsH = "H" Or specCsH = "S" Then
''''''        If Left(CRYNUM, 1) <> "8" Then                 '2003/10/18 �폜 SystemBrain
'''''            '���㌋���̎��ю擾
'''''            If SpecCsMin > 0 Then
'''''                'FromTo�d�l�̏ꍇ�́A�u���b�N��Top/Bot����l����������(���p�s��)
'''''                'Top��
'''''                sql = vbNullString
'''''                sql = sql & "select J.POSITION, J.CSMEAS, J.PRE70P "
'''''                sql = sql & "from TBCME040 B, TBCMJ004 J "
'''''                sql = sql & "where B.CRYNUM='" & CRYNUM & "'"
'''''                sql = sql & "  and B.INGOTPOS<=" & sxlFrom
'''''                sql = sql & "  and " & sxlFrom & "<B.INGOTPOS+B.LENGTH"
'''''                sql = sql & "  and J.CRYNUM=B.CRYNUM and J.POSITION=B.INGOTPOS "
'''''                sql = sql & "order by TRANCNT desc"
'''''                sql = "select * from (" & sql & ") where rownum=1"
'''''                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                If rs.RecordCount > 0 Then
'''''                    With Cs(1)
'''''                        .SXL_CS_SMPPOS = rs("POSITION")
'''''                        .SXLCS_CSMEAS = rs("CSMEAS")
'''''                        .SXLCS_70PPRE = rs("PRE70P")
'''''                    End With
'''''                ElseIf specCsH = "H" Then
'''''                    GoTo proc_exit       'Top�����тȂ�(SXL�m��s��)
'''''                End If
'''''                'Bot��
'''''                sql = vbNullString
'''''                sql = sql & "select J.POSITION, J.CSMEAS, J.PRE70P "
'''''                sql = sql & "from TBCME040 B, TBCMJ004 J "
'''''                sql = sql & "where B.CRYNUM='" & CRYNUM & "'"
'''''                sql = sql & "  and B.INGOTPOS<" & sxlFrom + sxlLen
'''''                sql = sql & "  and " & sxlFrom + sxlLen & "<=B.INGOTPOS+B.LENGTH"
'''''                sql = sql & "  and J.CRYNUM=B.CRYNUM and J.POSITION=B.INGOTPOS+B.LENGTH "
'''''                sql = sql & "order by TRANCNT desc"
'''''                sql = "select * from (" & sql & ") where rownum=1"
'''''                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                If rs.RecordCount > 0 Then
'''''                    With Cs(2)
'''''                        .SXL_CS_SMPPOS = rs("POSITION")
'''''                        .SXLCS_CSMEAS = rs("CSMEAS")
'''''                        .SXLCS_70PPRE = rs("PRE70P")
'''''                    End With
'''''                ElseIf specCsH = "H" Then
'''''                    GoTo proc_exit       'Tail�����тȂ�(SXL�m��s��)
'''''                End If
'''''            Else
'''''                'FromTo�d�l�łȂ���΁A�Ȃ�ׂ��߂��������猟������
'''''                sql = vbNullString
'''''                sql = sql & "select * from ("
'''''                sql = sql & "  select POSITION, CSMEAS, PRE70P"
'''''                sql = sql & "  from TBCMJ004 J"
'''''                sql = sql & "  where CRYNUM='" & CRYNUM & "'"
'''''                sql = sql & "    and POSITION>=" & sxlFrom + sxlLen
'''''                sql = sql & "  order by POSITION, TRANCOND, SMPKBN, TRANCNT desc"
'''''                sql = sql & ") where rownum=1"
'''''                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                If rs.RecordCount > 0 Then
'''''                    With Cs(2)
'''''                        .SXL_CS_SMPPOS = rs("POSITION")
'''''                        .SXLCS_CSMEAS = rs("CSMEAS")
'''''                        .SXLCS_70PPRE = rs("PRE70P")
'''''                    End With
'''''                ElseIf specCsH = "H" Then
'''''                    GoTo proc_exit       'Tail�����тȂ�(SXL�m��s��)
'''''                End If
'''''                rs.Close
'''''            End If
''''''2003/10/18 �폜 SystemBrain -------------------------------------------��
''''''        Else
''''''            '�w���P�����̎��ю擾
''''''            sql = vbNullString
''''''            sql = sql & "select * from ("
''''''            sql = sql & " select B.INGOTPOS, B.LENGTH, XL.CSTOP, XL.CSTAIL"
''''''            sql = sql & " from TBCMG002 XL, TBCME040 B "
''''''            sql = sql & " where B.CRYNUM='" & CRYNUM & "' and B.INGOTPOS<=" & sxlFrom & " and " & sxlFrom + sxlLen & "<=B.INGOTPOS+B.LENGTH"
''''''            sql = sql & "   and XL.CRYNUM=B.BLOCKID"
''''''            sql = sql & " order by TRANCNT desc"
''''''            sql = sql & ") where rownum=1"
''''''            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''''''            If rs.RecordCount > 0 Then
''''''                If rs("CSTOP") >= 0 Then    'Top���̒l�������Ă�����
''''''                    With Cs(1)
''''''                        .SXL_CS_SMPPOS = rs("INGOTPOS")
''''''                        .SXLCS_CSMEAS = rs("CSTOP")
''''''                        .SXLCS_70PPRE = 0
''''''                    End With
''''''                ElseIf (specCsH = "H") And (SpecCsMin > 0) Then 'FromTo�ۏ�
''''''                    GoTo proc_exit      'Top�����тȂ�(SXL�m��s��)
''''''                End If
''''''                If rs("CSTAIL") >= 0 Then    'Tail���̒l�������Ă�����
''''''                    With Cs(2)
''''''                        .SXL_CS_SMPPOS = Val(rs("INGOTPOS")) + Val(rs("LENGTH"))
''''''                        .SXLCS_CSMEAS = rs("CSTAIL")
''''''                        .SXLCS_70PPRE = 0
''''''                    End With
''''''                ElseIf specCsH = "H" Then
''''''                    GoTo proc_exit      'Tail�����тȂ�(SXL�m��s��)
''''''                End If
''''''            ElseIf specCsH = "H" Then
''''''                GoTo proc_exit          'Tail�����тȂ�(SXL�m��s��)
''''''            End If
''''''            rs.Close
''''''        End If
''''''2003/10/18 �폜 SystemBrain -------------------------------------------��
'''''    End If
'''''    getSXLCs = FUNCTION_RETURN_SUCCESS
'''''
'''''proc_exit:
'''''    '' �I��
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '' �G���[�n���h��
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    gErr.HandleError
'''''    getSXLCs = FUNCTION_RETURN_FAILURE
'''''    Resume proc_exit
'''''
'''''End Function
'2003/10/19 �g���ĂȂ��̂ō폜 SystemBrain ==========================================================��

'�T�v      :BMD���т�Min�l���v�Z����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :dMin          ,O   ,Double    ,Min�l
'          :strMeasPos    ,I   ,String    ,�������ב���ʒu�R�[�h�i3byte�j
'          :dMeas()       ,I   ,Double    ,����ʒu�z��
'          :�߂�l        ,O   ,Integer     ,���������FFUNCTION_RETURN_SUCCESS�@�������s�FFUNCTION_RETURN_FAILURE
'����      :
Private Function getSXLBMDMIN(dMin As Double, strMeasPos As String, dMeas() As Double) As Integer
    Dim dConv       As Double
    Dim iMeasNum    As Integer
    Dim Index       As Integer
    Dim dForMin()   As Double
    Dim strParam    As String

    On Error GoTo Err
    getSXLBMDMIN = FUNCTION_RETURN_FAILURE

    If strMeasPos = "" Then
        dMin = -1
        getSXLBMDMIN = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If

    '' �������ב���ʒu�i������@�j��芷�Z�W�����擾
    strParam = GetCodeField("GP", "01", Mid(strMeasPos, 1, 1), "INFO8")
    If strParam = vbNullString Then strParam = "1"
    dConv = val(strParam)

    '' �������ב���ʒu�i����_�j�̎擾
    iMeasNum = GetMeasureNum(Mid(strMeasPos, 2, 1), 1)
    If iMeasNum < 1 Then Exit Function

    '' Min�l�v�Z
    ReDim dForMin(iMeasNum - 1)
    For Index = 0 To UBound(dForMin)
        dForMin(Index) = dMeas(Index)
    Next Index
    dMin = GetMin(dForMin) * dConv / 10000

    getSXLBMDMIN = FUNCTION_RETURN_SUCCESS
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
End Function

Private Function NtoS(strWk As String) As String
    If Mid(strWk, 1, 1) = Chr(0) Then
        NtoS = " "
        Exit Function
    End If
    NtoS = strWk
End Function

Private Function NtoZ2(strWk As String) As Double
    If Trim(strWk) = "" Then
        NtoZ2 = -1
        Exit Function
    End If
    NtoZ2 = CDbl(strWk)
End Function

Private Function CryRES_Judg(CRs() As Double, GarRes As Guarantee) As Double
    Dim pt As Integer

    ''RRG����
    Select Case GarRes.cPos
      Case "B", "C", "D", "E", "F", "K", "S", "Y"
          Select Case GarRes.cBunp
          Case "A", "B", "C", "M"
             ''RRG�v�Z
             CryRES_Judg = MENNAI_Cal(RES_JUDG, CRs(), GarRes, GarRes.cBunp)

          Case "", " "                                          '��߰��ǉ��@05/07/05 ooba
             ''�v�Z�敪���X�y�[�X�̏ꍇ�́A�v�Z�C������s��Ȃ�
'             If GarRes.cBunp = "" Or GarRes.cBunp = " " Then   '�����ĉ��@05/07/05 ooba
'                    GoTo Cal_Escp
                CryRES_Judg = -1
                Exit Function
'             End If                                            '�����ĉ��@05/07/05 ooba

          Case Else
             ''RRG�v�Z�@�@�@�R�[�h "A" �ɂČv�Z
             If Trim(GarRes.cCount) = "" Then
                pt = 3
             Else
                pt = val(GarRes.cCount)
             End If
             CryRES_Judg = RoundUp((RGCal(CRs(), pt)), 4)

         End Select

      Case Else
         Select Case GarRes.cBunp
         Case "A", "B", "C", "D", "E", "M", "N"
             ''RRG�v�Z
             CryRES_Judg = MENNAI_Cal(RES_JUDG, CRs(), GarRes, GarRes.cBunp)

         Case "", " "                                           '��߰��ǉ��@05/07/05 ooba
             ''�v�Z�敪���X�y�[�X�̏ꍇ�́A�v�Z�C������s��Ȃ�
'             If GarRes.cBunp = "" Or GarRes.cBunp = " " Then   '�����ĉ��@05/07/05 ooba
'                    GoTo Cal_Escp
                CryRES_Judg = -1
                Exit Function
'             End If                                            '�����ĉ��@05/07/05 ooba

         Case Else
             ''RRG�v�Z�@�@�@�R�[�h "A" �ɂČv�Z
             If Trim(GarRes.cCount) = "" Then
                pt = 3
             Else
                pt = val(GarRes.cCount)
             End If
             CryRES_Judg = RoundUp((RGCal(CRs(), pt)), 4)

         End Select
    End Select
Cal_Escp:
        
End Function

Private Function CryOi_Judg(COi() As Double, GarOi As Guarantee) As Double
    Dim pt As Integer
    ReDim JData(UBound(COi())) As Double
    
    ''ORG����
    
    Select Case GarOi.cPos
      Case "B", "C", "D", "E", "F", "K", "Y"
          Select Case GarOi.cBunp
          Case "A", "B", "C"
             ''ORG�v�Z
             CryOi_Judg = MENNAI_Cal(OI_JUDG, COi(), GarOi, GarOi.cBunp)

          Case "", " "                                              '��߰��ǉ��@05/07/05 ooba
             ''�v�Z�敪���X�y�[�X�̏ꍇ�́A�v�Z�C������s��Ȃ�
'             If GarOi.cBunp = "" Or GarOi.cBunp = " " Then         '�����ĉ��@05/07/05 ooba
'                    GoTo Cal_Escp
                CryOi_Judg = -1
                Exit Function
'             End If                                                '�����ĉ��@05/07/05 ooba

          Case Else
             ''ORG�v�Z�@�@�@�R�[�h "A" �ɂČv�Z
             If Trim(GarOi.cCount) = "" Then
                pt = 3
             Else
                pt = val(GarOi.cCount)
             End If
             CryOi_Judg = RoundUp((RGCal(COi(), pt)), 4)

         End Select

      Case Else

         Select Case GarOi.cBunp
         Case "A", "B", "C", "D", "E", "N"
             ''ORG�v�Z
             CryOi_Judg = MENNAI_Cal(OI_JUDG, COi(), GarOi, GarOi.cBunp)

         Case "", " "                                               '��߰��ǉ��@05/07/05 ooba
             ''�v�Z�敪���X�y�[�X�̏ꍇ�́A�v�Z�C������s��Ȃ�
'             If GarOi.cBunp = "" Or GarOi.cBunp = " " Then         '�����ĉ��@05/07/05 ooba
'                    GoTo Cal_Escp
                CryOi_Judg = -1
                Exit Function
'             End If                                                '�����ĉ��@05/07/05 ooba

         Case Else
             ''ORG�v�Z�@�@�@�R�[�h "A" �ɂČv�Z
             If Trim(GarOi.cCount) = "" Then
                pt = 3
             Else
                pt = val(GarOi.cCount)
             End If
             CryOi_Judg = RoundUp((RGCal(COi(), pt)), 4)

         End Select
    End Select
Cal_Escp:

End Function

'�T�v      :������R����(TBCMJ002)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCS()      , I  ,c_cmzcrec         , XSDCS�\����   (�V����يǗ�(��ۯ�))
'          :i               , I  ,Integer           , Top/Bot���(1:Top, 2:Bot)
'          :hin             , I  ,tFullHinban       , �i��(�S�i�ԍ\����)
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :������R����(TBCMJ002)�����ް����擾���ASXL�������\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
Private Function getTBCMJ002(CRYNUM As String, recXSDCS() As c_cmzcrec, i As Integer, HIN As tFullHinban, recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim k           As Integer
    Dim wMeas1(2)   As Double
    Dim wgtCharge   As Long                 '�ΐ͌v�Z�p�p�����[�^
    Dim wgtCharge1  As Long                 '�ΐ͌v�Z�p�p�����[�^   '' 2008/11/26 SXL�������`���[�W�ʒǉ� ADD By Systech
    Dim wgtCharge2  As Long                 '�ΐ͌v�Z�p�p�����[�^   '' 2008/11/26 SXL�������`���[�W�ʒǉ� ADD By Systech
    Dim wgtTop      As Double               '�ΐ͌v�Z�p�p�����[�^
    Dim wgtTopCut   As Double               '�ΐ͌v�Z�p�p�����[�^
    Dim DM          As Double               '�ΐ͌v�Z�p�p�����[�^
    Dim cc          As type_Coefficient
    Dim CRes        As C_RES                '����RS����\����
    Dim wComp       As Double
    Dim wHSXRHWYS   As String               '�ۏؕ��@�Q��
    Dim RET As FUNCTION_RETURN
    Dim wStaff      As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ002"
    
    getTBCMJ002 = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    With recX001
        .Fields("SXL_RS_SMPPOS").Value = -1                 'SXLRS�T���v������ʒu(SXL������)
        .Fields("SXLRS_MEAS1").Value = -1                   'SXLRS_����l1
        .Fields("SXLRS_MEAS2").Value = -1                   'SXLRS_����l2
        .Fields("SXLRS_MEAS3").Value = -1                   'SXLRS_����l3
        .Fields("SXLRS_MEAS4").Value = -1                   'SXLRS_����l4
        .Fields("SXLRS_MEAS5").Value = -1                   'SXLRS_����l5
        .Fields("SXLRS_EFEHS").Value = -1                   'SXLRS_�����ΐ�
        .Fields("SXLRS_RRG").Value = -1                     'SXLRS_RRG
    
        '-------------------- TBCMJ002�̓ǂݍ���(Rs) ----------------------------------------
        If (recXSDCS(i)("CRYINDRSCS").Value <> "0") And (recXSDCS(i)("CRYRESRS1CS").Value <> "0") Then
            '�����ΐ͎Z�o�ׁ̈ATop/Bot�̗������擾
            For k = 1 To 2
                sql = "select * from TBCMJ002 "
                sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
                sql = sql & "      SMPLNO = " & recXSDCS(k)("CRYSMPLIDRSCS").Value & " "
                sql = sql & "order by TRANCNT desc"
                sql = "select * from (" & sql & ") where rownum = 1"
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                If k = i Then
                    .Fields("SXL_RS_SMPPOS").Value = rs("POSITION")             'SXLRS�T���v������ʒu(SXL������)
                    .Fields("SXLRS_MEAS1").Value = rs("MEAS1")                  'SXLRS_����l1
                    .Fields("SXLRS_MEAS2").Value = rs("MEAS2")                  'SXLRS_����l2
                    .Fields("SXLRS_MEAS3").Value = rs("MEAS3")                  'SXLRS_����l3
                    .Fields("SXLRS_MEAS4").Value = rs("MEAS4")                  'SXLRS_����l4
                    .Fields("SXLRS_MEAS5").Value = rs("MEAS5")                  'SXLRS_����l5
                    wStaff = rs("KSTAFFID")                                     '---TEST2004/10
                End If
                wMeas1(k) = rs("MEAS1")                             '�����ΐ͎Z�o�p
                Set rs = Nothing
            Next k
            
            'SXLRS_EFEHS
            '�}���`����Ή� �֐��Q�Ɛ�ύX 2008/05/26 SETsw Nakada
            If GetCoeffParams_new(CRYNUM, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'            If GetCoeffParams(CRYNUM, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then GoTo PROC_EXIT

'' 2008/11/26 SXL�������`���[�W�ʒǉ� DEL By Systech Start
''            .Fields("CHARGE").Value = wgtCharge '�`���[�W�� �擾��ύX 2008/05/26 SETsw Nakada
'' 2008/11/26 SXL�������`���[�W�ʒǉ� DEL By Systech End
            
            cc.DUNMENSEKI = AreaOfCircle(DM)
            cc.TOPSMPLPOS = recXSDCS(1)("INPOSCS").Value
            cc.BOTSMPLPOS = recXSDCS(2)("INPOSCS").Value
            cc.CHARGEWEIGHT = wgtCharge
            cc.TOPWEIGHT = wgtTop + wgtTopCut
            cc.TOPRES = wMeas1(1)
            cc.BOTRES = wMeas1(2)
            wComp = CoefficientCalculation(cc)
        
            If wComp = -9999 Then
                wComp = 0                                       'SXLRS_�����ΐ�
            End If
            .Fields("SXLRS_EFEHS").Value = wComp                'SXLRS_�����ΐ�
            
'''' 2008/11/26 SXL�������`���[�W�ʒǉ� UPD By Systech Start
''            '�`���[�W��
''            sql = " SELECT C1.SUICHARGE, C1.PUCHAGC1 "
''            sql = sql & " FROM XSDC1 C1 "
''            sql = sql & " WHERE C1.XTALC1 = '" & CRYNUM & "'"
''            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''            If rs.RecordCount = 0 Then
''                wgtCharge1 = 0                      ''����`���[�W
''                wgtCharge2 = 0                      ''�`���[�W��
''            Else
''                wgtCharge1 = rs("SUICHARGE")        ''����`���[�W
''                wgtCharge2 = rs("PUCHAGC1")         ''�`���[�W��
''            End If
''            Set rs = Nothing
''
''            .Fields("CHARGE").Value = wgtCharge2    '�`���[�W��
''            .Fields("ROCHARGE").Value = wgtCharge1  '����`���[�W
'''' 2008/11/26 SXL�������`���[�W�ʒǉ� UPD By Systech End

            'SXLRS_RRG
            sql = "select HSXRHWYS, HSXRSPOH, HSXRSPOT, HSXRSPOI from TBCME018 where "
            sql = sql & " HINBAN = '" & HIN.hinban & "' and "
            sql = sql & " MNOREVNO = " & HIN.mnorevno & " and "
            sql = sql & " FACTORY = '" & HIN.factory & "' and "
            sql = sql & " OPECOND = '" & HIN.opecond & "' "
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
                
            CRes.GuaranteeRes.cBunp = rs("HSXRSPOH")                    ' �i�r�w���R����ʒu�Q��
            CRes.GuaranteeRes.cCount = rs("HSXRSPOT")                   ' �i�r�w���R����ʒu�Q�_
            CRes.GuaranteeRes.cPos = rs("HSXRSPOI")                     ' �i�r�w���R����ʒu�Q��
            wHSXRHWYS = rs("HSXRHWYS")                                  ' �i�r�w���R�ۏؕ��@�Q��
            Set rs = Nothing
            
            CRes.Res(0) = NtoZ2(.Fields("SXLRS_MEAS1").Value)           'Rs����l1
            CRes.Res(1) = NtoZ2(.Fields("SXLRS_MEAS2").Value)           'Rs����l2
            CRes.Res(2) = NtoZ2(.Fields("SXLRS_MEAS3").Value)           'Rs����l3
            CRes.Res(3) = NtoZ2(.Fields("SXLRS_MEAS4").Value)           'Rs����l4
            CRes.Res(4) = NtoZ2(.Fields("SXLRS_MEAS5").Value)           'Rs����l5
            
            ''------TEST2004/10 -> 2004/12 ����f�[�^���ɍX�V�̂��ߍ폜
            ''-----> 2006/06 ����ʒu�ɂ��v�Z�͕K�v�Ȃ��߃R�����g���O�����菇�Ƀf�[�^��߂�������ǉ�����
            If Trim(wStaff) <> KSTAFF_J002 Then   '�V����f�[�^�̏ꍇ������������
                RET = Set_Rs_Ichi(CRes.GuaranteeRes.cCount, CRes.GuaranteeRes.cPos, CRes.Res(0), CRes.Res(1), CRes.Res(2), _
                               CRes.Res(3), CRes.Res(4))
            End If
            
            .Fields("SXLRS_RRG").Value = CryRES_Judg(CRes.Res(), CRes.GuaranteeRes)     'SXLRS_RRG
        
            '2006/06 �ǉ�----
            CRes.Res(0) = NtoZ2(.Fields("SXLRS_MEAS1").Value)           'Rs����l1
            CRes.Res(1) = NtoZ2(.Fields("SXLRS_MEAS2").Value)           'Rs����l2
            CRes.Res(2) = NtoZ2(.Fields("SXLRS_MEAS3").Value)           'Rs����l3
            CRes.Res(3) = NtoZ2(.Fields("SXLRS_MEAS4").Value)           'Rs����l4
            CRes.Res(4) = NtoZ2(.Fields("SXLRS_MEAS5").Value)           'Rs����l5
            '--------
            
            '�ۏؕ��@="H"�A���ASXLRS_RRG�v�Z���ʂ�-1�̏ꍇ�A�G���[�Ƃ���B2003/11/21 SystemBrain
            If (wHSXRHWYS = "H") And (.Fields("SXLRS_RRG").Value = -1) Then GoTo proc_exit
        
        End If
    End With

    getTBCMJ002 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ002 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :����Oi����(TBCMJ003)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS�\����   (�V����يǗ�(��ۯ�))
'          :hin             , I  ,tFullHinban       , �i��(�S�i�ԍ\����)
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :����Oi����(TBCMJ003)�����ް����擾���ASXL�������\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
Private Function getTBCMJ003(CRYNUM As String, recXSDCS As c_cmzcrec, HIN As tFullHinban, recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim COi         As C_Oi                 '����Oi����\����
    Dim wHSXONHWS   As String               '�ۏؕ��@�Q��
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ003"
    
    getTBCMJ003 = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    With recX001
        .Fields("SXL_OI_SMPPOS").Value = -1                 'SXLOI�T���v������ʒu(SXL������)
        .Fields("SXLOI_OIMEAS1").Value = -1                 'SXLOI_Oi����l1
        .Fields("SXLOI_OIMEAS2").Value = -1                 'SXLOI_Oi����l2
        .Fields("SXLOI_OIMEAS3").Value = -1                 'SXLOI_Oi����l3
        .Fields("SXLOI_OIMEAS4").Value = -1                 'SXLOI_Oi����l4
        .Fields("SXLOI_OIMEAS5").Value = -1                 'SXLOI_Oi����l5
        .Fields("SXLOI_ORGRES").Value = -1                  'SXLOI_ORG����
        .Fields("SXLOI_INSPECTWAY").Value = -1              'SXLOI�������@
    
        '-------------------- TBCMJ003�̓ǂݍ���(Oi) ----------------------------------------
        If (recXSDCS("CRYINDOICS").Value <> "0") And (recXSDCS("CRYRESOICS").Value <> "0") Then
            sql = "select * from TBCMJ003 "
            sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
            sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDOICS").Value & " "
            sql = sql & "  and TRANCOND = 0 "   'GFA��FTIR���Z�l�擾�ُ�Ή� 2011/02/28 SETsw kubota
            sql = sql & "order by TRANCNT desc"
            sql = "select * from (" & sql & ") where rownum = 1"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            .Fields("SXL_OI_SMPPOS").Value = rs("POSITION")             'SXLOI�T���v������ʒu(SXL������)
''''            .Fields("SXLOI_OIMEAS1").Value = rs("OIMEAS1")              'SXLOI_Oi����l1
''''            .Fields("SXLOI_OIMEAS2").Value = rs("OIMEAS2")              'SXLOI_Oi����l2
''''            .Fields("SXLOI_OIMEAS3").Value = rs("OIMEAS3")              'SXLOI_Oi����l3
''''            .Fields("SXLOI_OIMEAS4").Value = rs("OIMEAS4")              'SXLOI_Oi����l4
''''            .Fields("SXLOI_OIMEAS5").Value = rs("OIMEAS5")              'SXLOI_Oi����l5
            'OI_NULL�Ή��@2005/03/08 TUKU START --------------------------------------------------
            If IsNull(rs("OIMEAS1")) = False Then .Fields("SXLOI_OIMEAS1").Value = rs("OIMEAS1") Else .Fields("SXLOI_OIMEAS1").Value = -1  'SXLOI_Oi����l1
            If IsNull(rs("OIMEAS2")) = False Then .Fields("SXLOI_OIMEAS2").Value = rs("OIMEAS2") Else .Fields("SXLOI_OIMEAS2").Value = -1  'SXLOI_Oi����l2
            If IsNull(rs("OIMEAS3")) = False Then .Fields("SXLOI_OIMEAS3").Value = rs("OIMEAS3") Else .Fields("SXLOI_OIMEAS3").Value = -1  'SXLOI_Oi����l3
            If IsNull(rs("OIMEAS4")) = False Then .Fields("SXLOI_OIMEAS4").Value = rs("OIMEAS4") Else .Fields("SXLOI_OIMEAS4").Value = -1  'SXLOI_Oi����l4
            If IsNull(rs("OIMEAS5")) = False Then .Fields("SXLOI_OIMEAS5").Value = rs("OIMEAS5") Else .Fields("SXLOI_OIMEAS5").Value = -1  'SXLOI_Oi����l5
            'OI_NULL�Ή��@2005/03/08 TUKU END   --------------------------------------------------
            .Fields("SXLOI_INSPECTWAY").Value = rs("INSPECTWAY")        'SXLOI�������@
            Set rs = Nothing
        
            'SXLOI_ORG
            sql = "select HSXONHWS, HSXONSPH, HSXONSPT, HSXONSPI from TBCME019 where "
            sql = sql & " HINBAN = '" & HIN.hinban & "' and "
            sql = sql & " MNOREVNO = " & HIN.mnorevno & " and "
            sql = sql & " FACTORY = '" & HIN.factory & "' and "
            sql = sql & " OPECOND = '" & HIN.opecond & "' "
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            
            ReDim COi.Oi(4) As Double
            COi.GuaranteeOi.cBunp = rs("HSXONSPH")                      ' �i�r�w�_�f�Z�x����ʒu�Q��
            COi.GuaranteeOi.cCount = rs("HSXONSPT")                     ' �i�r�w�_�f�Z�x����ʒu�Q�_
            COi.GuaranteeOi.cPos = rs("HSXONSPI")                       ' �i�r�w�_�f�Z�x����ʒu�Q��
            wHSXONHWS = rs("HSXONHWS")                                  ' �i�r�w�_�f�Z�x�ۏؕ��@�Q��
            Set rs = Nothing

            COi.Oi(0) = NtoZ2(.Fields("SXLOI_OIMEAS1").Value)           'Oi����l1
            COi.Oi(1) = NtoZ2(.Fields("SXLOI_OIMEAS2").Value)           'Oi����l2
            COi.Oi(2) = NtoZ2(.Fields("SXLOI_OIMEAS3").Value)           'Oi����l3
            COi.Oi(3) = NtoZ2(.Fields("SXLOI_OIMEAS4").Value)           'Oi����l4
            COi.Oi(4) = NtoZ2(.Fields("SXLOI_OIMEAS5").Value)           'Oi����l5
            
            .Fields("SXLOI_ORGRES").Value = CryOi_Judg(COi.Oi(), COi.GuaranteeOi)       'SXLOI_ORG����
            
            '�ۏؕ��@="H"�A���ASXLOI_ORG�v�Z���ʂ�-1�̏ꍇ�A�G���[�Ƃ���B2003/11/21 SystemBrain
            If (wHSXONHWS = "H") And (.Fields("SXLOI_ORGRES").Value = -1) Then GoTo proc_exit
        
        End If
    End With

    getTBCMJ003 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ003 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :Cs����(TBCMJ004)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS�\����   (�V����يǗ�(��ۯ�))
'          :HIN             , I  ,tFullHinban       , �i�ԁ@06/04/20 ooba
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'          :sErrMsg         , O  ,String            , �װү���ށ@06/04/20 ooba
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :Cs����(TBCMJ004)�����ް����擾���ASXL�������\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
Private Function getTBCMJ004(CRYNUM As String, recXSDCS As c_cmzcrec, HIN As tFullHinban, _
                             recX001 As c_cmzcrec, sErrMsg As String) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    
    '06/04/20 ooba START =========================================>
    Dim rs2         As OraDynaset
    Dim dCmax       As Double           '�d�l(����l)
    Dim dCmin       As Double           '�d�l(�����l)
    Dim iSmpNo      As Long             '���茳�����No      'Integer��Long �T���v����6���Ή� 2007/05/28 SETsw kubota
    Dim tCsSuitei   As CS_SUITEI_TYPE   'CS����v�Z�p�\����
    Dim dCsSuitei   As Double           'Cs����l
    '06/04/20 ooba END ===========================================>
    
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ004"
    
    getTBCMJ004 = FUNCTION_RETURN_FAILURE

    sErrMsg = ""        '06/04/20 ooba
    
    '-------------------- �����ر ----------------------------------------
    With recX001
        .Fields("SXL_CS_SMPPOS").Value = -1                 'SXLCS�T���v������ʒu(SXL������)
        .Fields("SXLCS_CSMEAS").Value = -1                  'SXLCS_Cs�����l
        .Fields("SXLCS_70PPRE").Value = -1                  'SXLCS_70%����l
        .Fields("SXLCS_BSUIMEAS").Value = -1                'SXLCS_Cs��ۯ�����l�@06/04/20 ooba
    
        '-------------------- TBCMJ004�̓ǂݍ���(Cs) ----------------------------------------
        If (recXSDCS("CRYINDCSCS").Value <> "0") And (recXSDCS("CRYRESCSCS").Value <> "0") Then
            sql = "select * from TBCMJ004 "
            sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
            sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDCSCS").Value & " "
            sql = sql & "order by TRANCNT desc"
            sql = "select * from (" & sql & ") where rownum = 1"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            .Fields("SXL_CS_SMPPOS").Value = rs("POSITION")             'SXLCS�T���v������ʒu(SXL������)
''''            .Fields("SXLCS_CSMEAS").Value = rs("CSMEAS")                'SXLCS_Cs�����l
''''            .Fields("SXLCS_70PPRE").Value = rs("PRE70P")                'SXLCS_70%����l
            'OI_NULL�Ή��@2005/03/08 TUKU START --------------------------------------------------
            If IsNull(rs("CSMEAS")) = False Then .Fields("SXLCS_CSMEAS").Value = rs("CSMEAS") Else .Fields("SXLCS_CSMEAS").Value = -1  'SXLCS_Cs�����l
            If IsNull(rs("PRE70P")) = False Then .Fields("SXLCS_70PPRE").Value = rs("PRE70P") Else .Fields("SXLCS_70PPRE").Value = -1  'SXLCS_70%����l
            'OI_NULL�Ή��@2005/03/08 TUKU START --------------------------------------------------
            
            Set rs = Nothing
            
            ''Cs��ۯ�����l�v�Z�Ή��@06/04/20 ooba START ======================================>
        
            '�����̏ꍇ�͢��ۯ�����l�������l�
            If recXSDCS("CRYINDCSCS").Value = "1" Then
                .Fields("SXLCS_BSUIMEAS").Value = .Fields("SXLCS_CSMEAS").Value
            Else
                '�@����ʒu
                tCsSuitei.sInfPos = CStr(recXSDCS("INPOSCS").Value)
                
                '�A����وʒu
                '�B����ّ���l
                '���茳�����No�擾
                iSmpNo = recXSDCS("CRYSMPLIDCSCS").Value
                
'''                If recXSDCS("CRYINDCSCS").Value <> "0" Then
'''                    iSmpNo = recXSDCS("CRYSMPLIDCSCS").Value
'''                Else
'''                    '�d�l�l�擾
'''                    sql = "select HSXCNMAX, HSXCNMIN from TBCME019 "
'''                    sql = sql & "where HINBAN = '" & HIN.HINBAN & "' "
'''                    sql = sql & "and MNOREVNO = " & HIN.mnorevno & " "
'''                    sql = sql & "and FACTORY = '" & HIN.factory & "' "
'''                    sql = sql & "and OPECOND = '" & HIN.opecond & "' "
'''                    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''                    If rs.RecordCount = 0 Then
'''                        sErrMsg = GetMsgStr("ECLC3") & " <�����No> "
'''                        Set rs = Nothing
'''                        GoTo PROC_EXIT
'''                    End If
'''                    dCmax = fncNullCheck(rs("HSXCNMAX"))    '�iSX�Y�f�Z�x���
'''                    dCmin = fncNullCheck(rs("HSXCNMIN"))    '�iSX�Y�f�Z�x����
'''                    Set rs = Nothing
'''
'''                    'Cs���fٰقɂ�萄�茳�����No�擾
'''                    '���f�����1
'''                    If dCmax > 0 And dCmin > 0 Then
'''                        sql = "select CRYSMPLIDCSCS from XSDCS "
'''                        'TOP������
'''                        If recXSDCS("TBKBNCS").Value = "T" Then
'''                            sql = sql & "where tbkbncs = 'T' and "
'''                            sql = sql & "      xtalcs = '" & CRYNUM & "' and "
'''                            sql = sql & "      inposcs <= " & recXSDCS("INPOSCS").Value & " and "
'''                            sql = sql & "      (CRYINDCSCS = '1' or CRYINDCSCS = '2') and "
'''                            sql = sql & "      CRYRESCSCS <> '0' "
'''                            sql = sql & "order by inposcs desc"
'''                        'BOT������
'''                        Else
'''                            sql = sql & "where tbkbncs = 'B' and "
'''                            sql = sql & "      xtalcs = '" & CRYNUM & "' and "
'''                            sql = sql & "      inposcs >= " & recXSDCS("INPOSCS").Value & " and "
'''                            sql = sql & "      (CRYINDCSCS = '1' or CRYINDCSCS = '2') and "
'''                            sql = sql & "      CRYRESCSCS <> '0' "
'''                            sql = sql & "order by inposcs asc"
'''                        End If
'''                    '���f�����2
'''                    Else
'''                        sql = "select CRYSMPLIDCSCS from XSDCS "
'''                        sql = sql & "where xtalcs = '" & CRYNUM & "' and "
'''                        sql = sql & "      inposcs >= " & recXSDCS("INPOSCS").Value & " and "
'''                        sql = sql & "      (CRYINDCSCS = '1' or CRYINDCSCS = '2') and "
'''                        sql = sql & "      CRYRESCSCS <> '0' "
'''                        sql = sql & "order by inposcs asc"
'''                    End If
'''
'''                    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''                    If rs.RecordCount > 0 Then
'''                        iSmpNo = rs("CRYSMPLIDCSCS")
'''                    '�擾�o���Ȃ������ꍇ��TOP������
'''                    Else
'''                        sql = "select CRYSMPLIDCSCS from XSDCS "
'''                        sql = sql & "where xtalcs = '" & CRYNUM & "' and "
'''                        sql = sql & "      inposcs <= " & recXSDCS("INPOSCS").Value & " and "
'''                        sql = sql & "      (CRYINDCSCS = '1' or CRYINDCSCS = '2') and "
'''                        sql = sql & "      CRYRESCSCS <> '0' "
'''                        sql = sql & "order by inposcs desc"
'''
'''                        Set rs2 = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''                        If rs2.RecordCount = 0 Then
'''                            sErrMsg = GetMsgStr("ECLC3") & " <�����No> "
'''                            Set rs = Nothing
'''                            Set rs2 = Nothing
'''                            GoTo PROC_EXIT
'''                        End If
'''                        iSmpNo = rs2("CRYSMPLIDCSCS")
'''                        Set rs2 = Nothing
'''                    End If
'''                    Set rs = Nothing
'''                End If
                
                '����وʒu������ّ���l�擾
                sql = "select POSITION, CSMEAS from TBCMJ004 "
                sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
                sql = sql & "      SMPLNO = " & iSmpNo & " "
                sql = sql & "order by TRANCNT desc"
                
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    sErrMsg = GetMsgStr("ECLC3") & " <����ّ���l> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                tCsSuitei.sSamplePos = rs("POSITION")       '����وʒu
                tCsSuitei.sResCs = rs("CSMEAS")             '����ّ���l
                Set rs = Nothing
                
                '�C����ޗ�
                '�DTOP�d��
                sql = "select SUICHARGE, WGHTTOC1, PUTCUTWC1 from XSDC1 "
                sql = sql & "where XTALC1 = '" & CRYNUM & "' "
                
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    sErrMsg = GetMsgStr("ECLC3") & " <����ޗ�> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                '�ް��s��
                If (IsNull(rs("SUICHARGE")) Or IsNull(rs("WGHTTOC1")) Or IsNull(rs("PUTCUTWC1"))) Then
                    sErrMsg = GetMsgStr("ECLC3") & " <����ޗ�> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                
                tCsSuitei.sSiWeight = rs("SUICHARGE")       '��������ޗ�
                tCsSuitei.sTopWT = CLng(rs("WGHTTOC1")) + CLng(rs("PUTCUTWC1"))     'TOP�d��
                Set rs = Nothing
                '���������ޗ�=0�or���������ޗʁ�TOP�d�ʣ�̏ꍇ�ʹװ�Ƃ���
                If CLng(tCsSuitei.sSiWeight) = 0 Or _
                   (CLng(tCsSuitei.sSiWeight) <= CLng(tCsSuitei.sTopWT)) Then
                    sErrMsg = GetMsgStr("ECLC3") & " <����ޗ�> "
                    GoTo proc_exit
                End If
                
                '�E���a
                sql = "select HSXD1CEN from TBCME018 "
                sql = sql & "where HINBAN = '" & HIN.hinban & "' "
                sql = sql & "and MNOREVNO = " & HIN.mnorevno & " "
                sql = sql & "and FACTORY = '" & HIN.factory & "' "
                sql = sql & "and OPECOND = '" & HIN.opecond & "' "
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    sErrMsg = GetMsgStr("ECLC3") & " <���a> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                tCsSuitei.sUpDm = rs("HSXD1CEN")            '�iSX���a1���S
                
                '�F����ݕΐ͌W��
                sql = "select CTR01A9 from KODA9 "
                sql = sql & "where SYSCA9 = 'K' "
                sql = sql & "and SHUCA9 = 'AP' "
                sql = sql & "and CODEA9 = '1' "
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    sErrMsg = GetMsgStr("ECLC3") & " <����ݕΐ͌W��> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                tCsSuitei.sCsHenseki = rs("CTR01A9")        '����ݕΐ͌W��
                
                '�GCs��ۯ�����l�v�Z
                If Not GetCsSuiteiMain(tCsSuitei, dCsSuitei) Then
                    sErrMsg = GetMsgStr("ECLC3")
                    GoTo proc_exit
                End If
                .Fields("SXLCS_BSUIMEAS").Value = dCsSuitei
            End If
            ''Cs��ۯ�����l�v�Z�Ή��@06/04/20 ooba END ========================================>
        End If
    End With

    getTBCMJ004 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ004 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :����OSF����(TBCMJ005)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS�\����   (�V����يǗ�(��ۯ�))
'          :j               , I  ,Integer           , OSF No
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002�\����(����_�ް�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :����OSF����(TBCMJ005)�����ް����擾���ASXL�������E����_�ް��\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
Private Function getTBCMJ005(CRYNUM As String, recXSDCS As c_cmzcrec, j As Integer, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ005"
    
    getTBCMJ005 = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    'TBCMX001
    With recX001
        If j = 1 Then
            .Fields("SXLOSF_SMPPOS").Value = -1             'OSF�T���v������ʒu(SXL������)
        End If
        .Fields("SXLOSF" & j & "_KKSP").Value = ""          'OSFx�������ב���ʒu
        .Fields("SXLOSF" & j & "_NETU").Value = ""          'OSFx�M�����@
        .Fields("SXLOSF" & j & "_KKSET").Value = ""         'OSFx�������ב������+�I��ET��
        .Fields("SXLOSF" & j & "_CALCMAX").Value = -1       'OSFxSXL�v�Z���� Max_x
        .Fields("SXLOSF" & j & "_CALCAVE").Value = -1       'OSFxSXL�v�Z���� Ave_x
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
        If j = 1 Then
            .Fields("SXLOSF1_PTNJUDGRES").Value = ""            'OSF1�p�^�[�����茋��
        End If
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
    End With
        
    'TBCMX002
    With recX002
        If j = 1 Then
            .Fields("SXLOSF1_SMPPOS").Value = -1            'SXLOSF�T���v������ʒu(SXL�ʒu���)
        End If
        .Fields("SXLOSF" & j & "_KKSP").Value = ""          'SXLOSFx�������׊m��ʒu
        .Fields("SXLOSF" & j & "_NETU").Value = ""          'SXLOSFx�M�����@
        .Fields("SXLOSF" & j & "_KKSET").Value = ""         'SXLOSFx�������ב������+�I��ET��
        .Fields("SXLOSF" & j & "_MEAS1").Value = -1         'SXLOSFx����_1
        .Fields("SXLOSF" & j & "_MEAS2").Value = -1         'SXLOSFx����_2
        .Fields("SXLOSF" & j & "_MEAS3").Value = -1         'SXLOSFx����_3
        .Fields("SXLOSF" & j & "_MEAS4").Value = -1         'SXLOSFx����_4
        .Fields("SXLOSF" & j & "_MEAS5").Value = -1         'SXLOSFx����_5
        .Fields("SXLOSF" & j & "_MEAS6").Value = -1         'SXLOSFx����_6
        .Fields("SXLOSF" & j & "_MEAS7").Value = -1         'SXLOSFx����_7
        .Fields("SXLOSF" & j & "_MEAS8").Value = -1         'SXLOSFx����_8
        .Fields("SXLOSF" & j & "_MEAS9").Value = -1         'SXLOSFx����_9
        .Fields("SXLOSF" & j & "_MEAS10").Value = -1        'SXLOSFx����_10
        .Fields("SXLOSF" & j & "_MEAS11").Value = -1        'SXLOSFx����_11
        .Fields("SXLOSF" & j & "_MEAS12").Value = -1        'SXLOSFx����_12
        .Fields("SXLOSF" & j & "_MEAS13").Value = -1        'SXLOSFx����_13
        .Fields("SXLOSF" & j & "_MEAS14").Value = -1        'SXLOSFx����_14
        .Fields("SXLOSF" & j & "_MEAS15").Value = -1        'SXLOSFx����_15
        .Fields("SXLOSF" & j & "_MEAS16").Value = -1        'SXLOSFx����_16
        .Fields("SXLOSF" & j & "_MEAS17").Value = -1        'SXLOSFx����_17
        .Fields("SXLOSF" & j & "_MEAS18").Value = -1        'SXLOSFx����_18
        .Fields("SXLOSF" & j & "_MEAS19").Value = -1        'SXLOSFx����_19
        .Fields("SXLOSF" & j & "_MEAS20").Value = -1        'SXLOSFx����_20
    End With
    
    '-------------------- TBCMJ005�̓ǂݍ���(OSF1�`4) ----------------------------------------
    If (recXSDCS("CRYINDL" & j & "CS").Value <> "0") And (recXSDCS("CRYRESL" & j & "CS").Value <> "0") Then
        sql = "select * from TBCMJ005 "
        sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
        sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDL" & j & "CS").Value & " and "
        sql = sql & "      TRANCOND = '" & j & "' "
        sql = sql & "order by TRANCNT desc"
        sql = "select * from (" & sql & ") where rownum = 1"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
    
        'TBCMX001
        With recX001
            If .Fields("SXLOSF_SMPPOS").Value = -1 Then
                .Fields("SXLOSF_SMPPOS").Value = rs("POSITION")         'OSF�T���v������ʒu(SXL������)
            End If
            .Fields("SXLOSF" & j & "_KKSP").Value = rs("KKSP")          'OSFx�������ב���ʒu
            .Fields("SXLOSF" & j & "_NETU").Value = rs("HTPRC")         'OSFx�M�����@
            .Fields("SXLOSF" & j & "_KKSET").Value = rs("KKSET")        'OSFx�������ב������+�I��ET��
            .Fields("SXLOSF" & j & "_CALCMAX").Value = rs("CALCMAX")    'OSFxSXL�v�Z���� Max_x
            .Fields("SXLOSF" & j & "_CALCAVE").Value = rs("CALCAVE")    'OSFxSXL�v�Z���� Ave_x
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
            'OSF1�p�^�[�����茋��
            If j = 1 Then
                If IsNull(rs("PTNJUDGRES")) = True Then
                    .Fields("SXLOSF1_PTNJUDGRES").Value = " "
                Else
                    .Fields("SXLOSF1_PTNJUDGRES").Value = rs("PTNJUDGRES")
                End If
            End If
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
        
        End With
            
        'TBCMX002
        With recX002
            If .Fields("SXLOSF1_SMPPOS").Value = -1 Then
                .Fields("SXLOSF1_SMPPOS").Value = rs("POSITION")        'SXLOSF�T���v������ʒu(SXL�ʒu���)
            End If
            .Fields("SXLOSF" & j & "_KKSP").Value = rs("KKSP")          'SXLOSFx�������׊m��ʒu
            .Fields("SXLOSF" & j & "_NETU").Value = rs("HTPRC")         'SXLOSFx�M�����@
            .Fields("SXLOSF" & j & "_KKSET").Value = rs("KKSET")        'SXLOSFx�������ב������+�I��ET��
            .Fields("SXLOSF" & j & "_MEAS1").Value = rs("MEAS1")        'SXLOSFx����_1
            .Fields("SXLOSF" & j & "_MEAS2").Value = rs("MEAS2")        'SXLOSFx����_2
            .Fields("SXLOSF" & j & "_MEAS3").Value = rs("MEAS3")        'SXLOSFx����_3
            .Fields("SXLOSF" & j & "_MEAS4").Value = rs("MEAS4")        'SXLOSFx����_4
            .Fields("SXLOSF" & j & "_MEAS5").Value = rs("MEAS5")        'SXLOSFx����_5
            .Fields("SXLOSF" & j & "_MEAS6").Value = rs("MEAS6")        'SXLOSFx����_6
            .Fields("SXLOSF" & j & "_MEAS7").Value = rs("MEAS7")        'SXLOSFx����_7
            .Fields("SXLOSF" & j & "_MEAS8").Value = rs("MEAS8")        'SXLOSFx����_8
            .Fields("SXLOSF" & j & "_MEAS9").Value = rs("MEAS9")        'SXLOSFx����_9
            .Fields("SXLOSF" & j & "_MEAS10").Value = rs("MEAS10")      'SXLOSFx����_10
            .Fields("SXLOSF" & j & "_MEAS11").Value = rs("MEAS11")      'SXLOSFx����_11
            .Fields("SXLOSF" & j & "_MEAS12").Value = rs("MEAS12")      'SXLOSFx����_12
            .Fields("SXLOSF" & j & "_MEAS13").Value = rs("MEAS13")      'SXLOSFx����_13
            .Fields("SXLOSF" & j & "_MEAS14").Value = rs("MEAS14")      'SXLOSFx����_14
            .Fields("SXLOSF" & j & "_MEAS15").Value = rs("MEAS15")      'SXLOSFx����_15
            .Fields("SXLOSF" & j & "_MEAS16").Value = rs("MEAS16")      'SXLOSFx����_16
            .Fields("SXLOSF" & j & "_MEAS17").Value = rs("MEAS17")      'SXLOSFx����_17
            .Fields("SXLOSF" & j & "_MEAS18").Value = rs("MEAS18")      'SXLOSFx����_18
            .Fields("SXLOSF" & j & "_MEAS19").Value = rs("MEAS19")      'SXLOSFx����_19
            .Fields("SXLOSF" & j & "_MEAS20").Value = rs("MEAS20")      'SXLOSFx����_20
        End With
        Set rs = Nothing
    End If

    getTBCMJ005 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ005 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :����BMD����(TBCMJ008)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS�\����   (�V����يǗ�(��ۯ�))
'          :j               , I  ,Integer           , BMD No
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002�\����(����_�ް�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :����BMD����(TBCMJ008)�����ް����擾���ASXL�������E����_�ް��\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
Private Function getTBCMJ008(CRYNUM As String, recXSDCS As c_cmzcrec, j As Integer, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim dMeas(9)    As Double
    Dim strMeasPos  As String
    Dim iRet        As Integer
    Dim wComp       As Double
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ008"
    
    getTBCMJ008 = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    'TBCMX001
    With recX001
        If j = 1 Then
            .Fields("SXLBMD_SMPPOS").Value = -1             'BMD�T���v������ʒu(SXL�ʒu���)
        End If
        .Fields("SXLBMD" & j & "_KKSP").Value = ""          'BMDx�������ב���ʒu
        .Fields("SXLBMD" & j & "_NETU").Value = ""          'BMDx�M�����@
        .Fields("SXLBMD" & j & "_KKSET").Value = ""         'BMDx�������ב�������{�I��ET��
        .Fields("SXLBMD" & j & "_CALCMAX").Value = -1       'BMDxSXL�v�Z���� Max
        .Fields("SXLBMD" & j & "_CALCAVE").Value = -1       'BMDxSXL�v�Z���� Ave
        .Fields("SXLBMD" & j & "_CALCMIN").Value = -1       'BMDxSXL�v�Z���� Min
        .Fields("SXLBMD" & j & "_CALCMB").Value = -1        'BMDxSXL�v�Z���� �ʓ����z
    End With
        
    'TBCMX002
    With recX002
        If j = 1 Then
            .Fields("SXLBMD_SMPPOS").Value = -1             'SXLBMD�T���v������ʒu(SXL�ʒu���)
        End If
        .Fields("SXLBMD" & j & "_KKSP").Value = ""          'SXLBMD1�������ב���ʒu
        .Fields("SXLBMD" & j & "_NETU").Value = ""          'SXLBMD1�M�����@
        .Fields("SXLBMD" & j & "_KKSET").Value = ""         'SXLBMD1�������ב������+�I��ET��
        .Fields("SXLBMD" & j & "_MEAS1").Value = -1         'SXLBMD1����_1
        .Fields("SXLBMD" & j & "_MEAS2").Value = -1         'SXLBMD1����_2
        .Fields("SXLBMD" & j & "_MEAS3").Value = -1         'SXLBMD1����_3
        .Fields("SXLBMD" & j & "_MEAS4").Value = -1         'SXLBMD1����_4
        .Fields("SXLBMD" & j & "_MEAS5").Value = -1         'SXLBMD1����_5
    End With
    
    '-------------------- TBCMJ008�̓ǂݍ���(BMD1�`3) ----------------------------------------
    If (recXSDCS("CRYINDB" & j & "CS").Value <> "0") And (recXSDCS("CRYRESB" & j & "CS").Value <> "0") Then
        sql = "select * from TBCMJ008 "
        sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
        sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDB" & j & "CS").Value & " and "
        sql = sql & "      TRANCOND = '" & j & "' "
        sql = sql & "order by TRANCNT desc"
        sql = "select * from (" & sql & ") where rownum = 1"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            If .Fields("SXLBMD_SMPPOS").Value = -1 Then
                .Fields("SXLBMD_SMPPOS").Value = rs("POSITION")         'BMD�T���v������ʒu(SXL�ʒu���)
            End If
            .Fields("SXLBMD" & j & "_KKSP").Value = rs("KKSP")          'BMDx�������ב���ʒu
            .Fields("SXLBMD" & j & "_NETU").Value = rs("HTPRC")         'BMDx�M�����@
            .Fields("SXLBMD" & j & "_KKSET").Value = rs("KKSET")        'BMDx�������ב�������{�I��ET��
            .Fields("SXLBMD" & j & "_CALCMAX").Value = rs("MEASMAX")    'BMDxSXL�v�Z���� Max
            .Fields("SXLBMD" & j & "_CALCAVE").Value = rs("MEASAVE")    'BMDxSXL�v�Z���� Ave
'            .Fields("SXLBMD" & j & "_CALCMB").Value = rs("BMDMNBUNP")   'BMDxSXL�v�Z���� �ʓ����z
            If IsNull(rs("BMDMNBUNP")) = False Then .Fields("SXLBMD" & j & "_CALCMB").Value = rs("BMDMNBUNP")   'BMDxSXL�v�Z���� �ʓ����z
        End With
            
        'TBCMX002
        With recX002
            If .Fields("SXLBMD_SMPPOS").Value = -1 Then
                .Fields("SXLBMD_SMPPOS").Value = rs("POSITION")         'SXLBMD�T���v������ʒu(SXL�ʒu���)
            End If
            .Fields("SXLBMD" & j & "_KKSP").Value = rs("KKSP")          'SXLBMDx�������ב���ʒu
            .Fields("SXLBMD" & j & "_NETU").Value = rs("HTPRC")         'SXLBMDx�M�����@
            .Fields("SXLBMD" & j & "_KKSET").Value = rs("KKSET")        'SXLBMDx�������ב������+�I��ET��
            .Fields("SXLBMD" & j & "_MEAS1").Value = rs("MEAS1")        'SXLBMDx����_1
            .Fields("SXLBMD" & j & "_MEAS2").Value = rs("MEAS2")        'SXLBMDx����_2
            .Fields("SXLBMD" & j & "_MEAS3").Value = rs("MEAS3")        'SXLBMDx����_3
            .Fields("SXLBMD" & j & "_MEAS4").Value = rs("MEAS4")        'SXLBMDx����_4
            .Fields("SXLBMD" & j & "_MEAS5").Value = rs("MEAS5")        'SXLBMDx����_5
        End With
        Set rs = Nothing
    
        'BMD�ŏ��l�̎擾 2003/05/31 tuku                START
        dMeas(0) = recX002.Fields("SXLBMD" & j & "_MEAS1").Value
        dMeas(1) = recX002.Fields("SXLBMD" & j & "_MEAS2").Value
        dMeas(2) = recX002.Fields("SXLBMD" & j & "_MEAS3").Value
        dMeas(3) = recX002.Fields("SXLBMD" & j & "_MEAS4").Value
        dMeas(4) = recX002.Fields("SXLBMD" & j & "_MEAS5").Value
        ''�������ב���ʒu�R�[�h
        strMeasPos = Trim(recX002.Fields("SXLBMD" & j & "_KKSP").Value)
        ''�ŏ��l���v�Z����B
        iRet = getSXLBMDMIN(wComp, strMeasPos, dMeas)
        ''�v�Z���ʂ��i�[����
        recX001.Fields("SXLBMD" & j & "_CALCMIN").Value = wComp
    End If

    getTBCMJ008 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ008 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :GD����(TBCMJ006)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS�\����   (�V����يǗ�(��ۯ�))
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002�\����(����_�ް�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :GD����(TBCMJ006)�����ް����擾���ASXL�������E����_�ް��\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
Private Function getTBCMJ006(CRYNUM As String, recXSDCS As c_cmzcrec, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ006"
    
    getTBCMJ006 = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("SXLGD_SMPPOS").Value = -1                  'GD�T���v������ʒu(SXL�ʒu���)
        .Fields("SXLGD_MSRSDEN").Value = -1                 'SXLGD_���茋�� Den
        .Fields("SXLGD_MSRSLDL").Value = -1                 'SXLGD_���茋�� L/DL
        .Fields("SXLGD_MSRSDVD2").Value = -1                'SXLGD_���茋�� DVD2
    End With
        
    'TBCMX002
    With recX002
        .Fields("SXLGD_SMPPOS").Value = -1                                  'SXLGD�T���v������ʒu(SXL�ʒu���)
        For i = 1 To 15
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL1").Value = -1       'SXLGD_����lxx L/DL1
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL2").Value = -1       'SXLGD_����lxx L/DL2
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL3").Value = -1       'SXLGD_����lxx L/DL3
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL4").Value = -1       'SXLGD_����lxx L/DL4
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL5").Value = -1       'SXLGD_����lxx L/DL5
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN1").Value = -1       'SXLGD_����lxx Den1
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN2").Value = -1       'SXLGD_����lxx Den2
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN3").Value = -1       'SXLGD_����lxx Den3
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN4").Value = -1       'SXLGD_����lxx Den4
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN5").Value = -1       'SXLGD_����lxx Den5
        Next
    End With
        
    '-------------------- TBCMJ006�̓ǂݍ���(GD) ----------------------------------------
    If (recXSDCS("CRYINDGDCS").Value <> "0") And (recXSDCS("CRYRESGDCS").Value <> "0") Then
        sql = "select * from TBCMJ006 "
        sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
        sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDGDCS").Value & " "
        sql = sql & "order by TRANCNT desc"
        sql = "select * from (" & sql & ") where rownum = 1"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            .Fields("SXLGD_SMPPOS").Value = rs("POSITION")              'GD�T���v������ʒu(SXL�ʒu���)
            .Fields("SXLGD_MSRSDEN").Value = rs("MSRSDEN")              'SXLGD_���茋�� Den
            .Fields("SXLGD_MSRSLDL").Value = rs("MSRSLDL")              'SXLGD_���茋�� L/DL
            .Fields("SXLGD_MSRSDVD2").Value = rs("MSRSDVD2")            'SXLGD_���茋�� DVD2
        End With
            
        'TBCMX002
        With recX002
            .Fields("SXLGD_SMPPOS").Value = rs("POSITION")                                                      'SXLGD�T���v������ʒu(SXL�ʒu���)
            For i = 1 To 15
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL1").Value = rs("MS" & Format(i, "00") & "LDL1")      'SXLGD_����lxx L/DL1
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL2").Value = rs("MS" & Format(i, "00") & "LDL2")      'SXLGD_����lxx L/DL2
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL3").Value = rs("MS" & Format(i, "00") & "LDL3")      'SXLGD_����lxx L/DL3
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL4").Value = rs("MS" & Format(i, "00") & "LDL4")      'SXLGD_����lxx L/DL4
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL5").Value = rs("MS" & Format(i, "00") & "LDL5")      'SXLGD_����lxx L/DL5
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN1").Value = rs("MS" & Format(i, "00") & "DEN1")      'SXLGD_����lxx Den1
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN2").Value = rs("MS" & Format(i, "00") & "DEN2")      'SXLGD_����lxx Den2
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN3").Value = rs("MS" & Format(i, "00") & "DEN3")      'SXLGD_����lxx Den3
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN4").Value = rs("MS" & Format(i, "00") & "DEN4")      'SXLGD_����lxx Den4
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN5").Value = rs("MS" & Format(i, "00") & "DEN5")      'SXLGD_����lxx Den5
            Next
        End With
        Set rs = Nothing
    End If

    getTBCMJ006 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ006 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :LT����(TBCMJ007)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS�\����   (�V����يǗ�(��ۯ�))
'          :ChkHin          , I  ,tFullHinban       , LT�d�l�擾�p�i�ԁ@05/12/05 ooba
'          :i               , I  ,Integer           , BMD No
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002�\����(����_�ް�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :LT����(TBCMJ007)�����ް����擾���ASXL�������E����_�ް��\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
Private Function getTBCMJ007(CRYNUM As String, recXSDCS As c_cmzcrec, ChkHin As tFullHinban, i As Integer, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim j           As Integer      '                       '05/12/05 ooba START =======>
    Dim rs2         As OraDynaset   '
    Dim sql2        As String       '
    Dim iRet        As Integer      '
    Dim iTmpMes(9)  As Integer      'LT�����ް�(1�`10)
    Dim iCalcMeas   As Integer      'LT�v�Z����
    Dim sIchi       As String       '�iSXL��ё���ʒu_��
    Dim iOldFlg     As Integer      '���ް��׸�             '05/12/05 ooba END =========>
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ007"
    
    getTBCMJ007 = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("SXLLT_SMPPOS").Value = -1                  'LT�T���v������ʒu(SXL�ʒu���)
        .Fields("SXLLT_MEASPEAK").Value = -1                'SXLLT_����l �s�[�N�l
        .Fields("SXLLT_CALCMEAS").Value = -1                'SXLLT_�v�Z����
    End With
        
    'TBCMX002
    With recX002
        .Fields("SXLT_SMPPOS").Value = -1                   'SXLLT�T���v������ʒu(SXL�ʒu���)
        .Fields("SXLLT_MEASPEAK").Value = -1                'SXLLT_����l �s�[�N�l
        .Fields("SXLLT_MEAS1").Value = -1                   'SXLLT_����l1
        .Fields("SXLLT_MEAS2").Value = -1                   'SXLLT_����l2
        .Fields("SXLLT_MEAS3").Value = -1                   'SXLLT_����l3
        .Fields("SXLLT_MEAS4").Value = -1                   'SXLLT_����l4
        .Fields("SXLLT_MEAS5").Value = -1                   'SXLLT_����l5
    End With
        
    'BOT���̂��ް��擾
    If i <> 1 Then
        '-------------------- TBCMJ007�̓ǂݍ���(LT) ----------------------------------------
        If (recXSDCS("CRYINDTCS").Value <> "0") And (recXSDCS("CRYRESTCS").Value <> "0") Then
            sql = "select * from TBCMJ007 "
            sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
            sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDTCS").Value & " "
            sql = sql & "order by TRANCNT desc"
            sql = "select * from (" & sql & ") where rownum = 1"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            
            ''LT���ё��M�ް��ύX�@05/12/05 ooba START ==================================>
            If IsNull(rs("LTSPIFLG")) Then iOldFlg = 1 Else iOldFlg = 0
            
            '������
            iCalcMeas = -1
            For j = 0 To 9
                iTmpMes(j) = -1
            Next j
            
            If Not IsNull(rs("MEAS1")) Then iTmpMes(0) = rs("MEAS1")
            If Not IsNull(rs("MEAS2")) Then iTmpMes(1) = rs("MEAS2")
            If Not IsNull(rs("MEAS3")) Then iTmpMes(2) = rs("MEAS3")
            If Not IsNull(rs("MEAS4")) Then iTmpMes(3) = rs("MEAS4")
            If Not IsNull(rs("MEAS5")) Then iTmpMes(4) = rs("MEAS5")
            If Not IsNull(rs("MEAS6")) Then iTmpMes(5) = rs("MEAS6")
            If Not IsNull(rs("MEAS7")) Then iTmpMes(6) = rs("MEAS7")
            If Not IsNull(rs("MEAS8")) Then iTmpMes(7) = rs("MEAS8")
            If Not IsNull(rs("MEAS9")) Then iTmpMes(8) = rs("MEAS9")
            If Not IsNull(rs("MEAS10")) Then iTmpMes(9) = rs("MEAS10")
            
            '10�_����̏ꍇ
            If iOldFlg = 0 Then
                sql2 = "select HSXLTSPI from TBCME019"
                sql2 = sql2 & " where HINBAN = '" & ChkHin.hinban & "'"
                sql2 = sql2 & " and MNOREVNO = " & ChkHin.mnorevno
                sql2 = sql2 & " and FACTORY = '" & ChkHin.factory & "'"
                sql2 = sql2 & " and OPECOND = '" & ChkHin.opecond & "'"
                Set rs2 = OraDB.DBCreateDynaset(sql2, ORADYN_NO_BLANKSTRIP)
                If rs2.RecordCount = 0 Then
                    Set rs2 = Nothing
                    GoTo proc_exit
                End If
                If Not IsNull(rs2("HSXLTSPI")) Then sIchi = rs2("HSXLTSPI") Else sIchi = ""
                Set rs2 = Nothing
            End If
            
            '�v�Z���ʎ擾
            iRet = KNS_CalculateMeasResult_LT(iCalcMeas, iTmpMes(), sIchi, iOldFlg)
            ''LT���ё��M�ް��ύX�@05/12/05 ooba END ====================================>
            
            'TBCMX001
            With recX001
                .Fields("SXLLT_SMPPOS").Value = rs("POSITION")          'LT�T���v������ʒu(SXL�ʒu���)
                .Fields("SXLLT_MEASPEAK").Value = rs("MEASPEAK")        'SXLLT_����l �s�[�N�l
'                .Fields("SXLLT_CALCMEAS").Value = rs("CALCMEAS")        'SXLLT_�v�Z����
                .Fields("SXLLT_CALCMEAS").Value = iCalcMeas             'SXLLT_�v�Z���ʁ@05/12/05 ooba
            End With
                
            'TBCMX002
            With recX002
                .Fields("SXLT_SMPPOS").Value = rs("POSITION")           'SXLLT�T���v������ʒu(SXL�ʒu���)
                .Fields("SXLLT_MEASPEAK").Value = rs("MEASPEAK")        'SXLLT_����l �s�[�N�l
'                .Fields("SXLLT_MEAS1").Value = rs("MEAS1")              'SXLLT_����l1
'                .Fields("SXLLT_MEAS2").Value = rs("MEAS2")              'SXLLT_����l2
'                .Fields("SXLLT_MEAS3").Value = rs("MEAS3")              'SXLLT_����l3
'                .Fields("SXLLT_MEAS4").Value = rs("MEAS4")              'SXLLT_����l4
'                .Fields("SXLLT_MEAS5").Value = rs("MEAS5")              'SXLLT_����l5

                ''LT���ѓo�^�ύX�@05/12/05 ooba START =====================================>
                '���ް�
                If iOldFlg = 1 Then
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_����l1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(1)           'SXLLT_����l2
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(2)           'SXLLT_����l3
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(3)           'SXLLT_����l4
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(4)           'SXLLT_����l5
                '3:CE,Inside3mm
                ElseIf sIchi = "3" Then
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_����l1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(0)           'SXLLT_����l1
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(7)           'SXLLT_����l8
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(8)           'SXLLT_����l9
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(9)           'SXLLT_����l10
                '5:CE,Inside5mm
                ElseIf sIchi = "5" Then
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_����l1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(0)           'SXLLT_����l1
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(4)           'SXLLT_����l5
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(5)           'SXLLT_����l6
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(6)           'SXLLT_����l7
                'A:CE,Inside10mm
                ElseIf sIchi = "A" Then
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_����l1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(0)           'SXLLT_����l1
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(1)           'SXLLT_����l2
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(2)           'SXLLT_����l3
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(3)           'SXLLT_����l4
                '���̑�
                Else
'                    Set rs = Nothing
'                    GoTo proc_exit
                    '���̑��̏ꍇ�͢A:CE,Inside10mm��Ƃ���@05/12/21 ooba
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_����l1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(0)           'SXLLT_����l1
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(1)           'SXLLT_����l2
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(2)           'SXLLT_����l3
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(3)           'SXLLT_����l4
                End If
                ''LT���ѓo�^�ύX�@05/12/05 ooba END =======================================>
            End With
            Set rs = Nothing
        End If
    End If

    getTBCMJ007 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ007 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :WFOi����(TBCMY013)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW�\����   (�V����يǗ�(SXL))
'          :hin             , I  ,tFullHinban       , �i��(�S�i�ԍ\����)
'�@�@      :sPos  �@�@�@    ,I   ,String �@         ,SXL�ʒu(TOP/BOT)   04/04/15 ooba
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :WFOi����(TBCMY013)�����ް����擾���ASXL�������\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
Private Function getTBCMY013WFOi(recXSDCW As c_cmzcrec, HIN As tFullHinban, sPos As String, recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim WOi     As W_OI                     'WFOI�\����
    Dim HWFONKHN As String                  '�i�v�e�_�f�Z�x�����p�x�Q���@04/04/15 ooba
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY013WFOi"
    
    getTBCMY013WFOi = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    With recX001
        .Fields("WFOI_SMPPOS").Value = -1                   'WFOI�����-ID����ʒu(SXL�ʒu���)
        .Fields("WFOI_NETSU").Value = ""                    'WFOI_�M��������
        .Fields("WFOI_ET").Value = ""                       'WFOI_�G�b�`���O����
        .Fields("WFOI_MES").Value = ""                      'WFOI_�v�����@
        .Fields("WFOI_MESDATA1").Value = -1                 'WFOI_����f�[�^���̂P
        .Fields("WFOI_MESDATA2").Value = -1                 'WFOI_����f�[�^���̂Q
        .Fields("WFOI_MESDATA3").Value = -1                 'WFOI_����f�[�^���̂R
        .Fields("WFOI_MESDATA4").Value = -1                 'WFOI_����f�[�^���̂S
        .Fields("WFOI_MESDATA5").Value = -1                 'WFOI_����f�[�^���̂T
        .Fields("WFOI_MESDATA6").Value = -1                 'WFOI_����f�[�^���̂U
        .Fields("WFOI_MESDATA7").Value = -1                 'WFOI_����f�[�^���̂V
        .Fields("WFOI_MESDATA8").Value = -1                 'WFOI_����f�[�^���̂W
        .Fields("WFOI_MESDATA9").Value = -1                 'WFOI_����f�[�^���̂X
        .Fields("WFOI_MESDATA10").Value = -1                'WFOI_����f�[�^���̂P�O
        .Fields("WFOI_ORG").Value = -1                      'WFOI_ORG�v�Z����
    
        '-------------------- TBCMY013�̓ǂݍ���(WFOi) ----------------------------------------
        If (recXSDCW("WFINDOICW").Value <> "0") And (recXSDCW("WFRESOICW").Value <> "0") Then
            sql = "select * from TBCMY013 "
            sql = sql & "where SAMPLEID = '" & recXSDCW("WFSMPLIDOICW").Value & "' and "
            sql = sql & "      SPEC = 'OI'"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            .Fields("WFOI_SMPPOS").Value = recXSDCW("INPOSCW").Value        'WFOI�����-ID����ʒu(SXL�ʒu���)
            .Fields("WFOI_NETSU").Value = rs("NETSU")                       'WFOI_�M��������
            .Fields("WFOI_ET").Value = rs("ET")                             'WFOI_�G�b�`���O����
            .Fields("WFOI_MES").Value = rs("MES")                           'WFOI_�v�����@
            .Fields("WFOI_MESDATA1").Value = rs("MESDATA1")                 'WFOI_����f�[�^���̂P
            .Fields("WFOI_MESDATA2").Value = rs("MESDATA2")                 'WFOI_����f�[�^���̂Q
            .Fields("WFOI_MESDATA3").Value = rs("MESDATA3")                 'WFOI_����f�[�^���̂R
            .Fields("WFOI_MESDATA4").Value = rs("MESDATA4")                 'WFOI_����f�[�^���̂S
            .Fields("WFOI_MESDATA5").Value = rs("MESDATA5")                 'WFOI_����f�[�^���̂T
            .Fields("WFOI_MESDATA6").Value = rs("MESDATA6")                 'WFOI_����f�[�^���̂U
            .Fields("WFOI_MESDATA7").Value = rs("MESDATA7")                 'WFOI_����f�[�^���̂V
            .Fields("WFOI_MESDATA8").Value = rs("MESDATA8")                 'WFOI_����f�[�^���̂W
            .Fields("WFOI_MESDATA9").Value = rs("MESDATA9")                 'WFOI_����f�[�^���̂X
            .Fields("WFOI_MESDATA10").Value = rs("MESDATA10")               'WFOI_����f�[�^���̂P�O
            Set rs = Nothing
            
            'WFOi_ORG
            sql = "select E025.HWFONSPH, E019.HSXONSPT, E019.HSXONSPI, E025.HWFONHWT, E025.HWFONHWS, E025.HWFONMCL, E025.HWFONKHN "
            sql = sql & "from TBCME025 E025, TBCME019 E019 where "
            sql = sql & " E025.HINBAN = '" & HIN.hinban & "' and "
            sql = sql & " E025.MNOREVNO = " & HIN.mnorevno & " and "
            sql = sql & " E025.FACTORY = '" & HIN.factory & "' and "
            sql = sql & " E025.OPECOND = '" & HIN.opecond & "' and "
            sql = sql & " E019.HINBAN = E025.HINBAN and "
            sql = sql & " E019.MNOREVNO = E025.MNOREVNO and "
            sql = sql & " E019.FACTORY = E025.FACTORY and "
            sql = sql & " E019.OPECOND = E025.OPECOND"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            
            WOi.GuaranteeOi.cMeth = rs("HWFONSPH")                      '�i�v�e�_�f�Z�x����ʒu�Q��
            WOi.GuaranteeOi.cCount = rs("HSXONSPT")                     '�i�r�w�_�f�Z�x����ʒu�Q�_
            WOi.GuaranteeOi.cPos = rs("HSXONSPI")                       '�i�r�w�_�f�Z�x����ʒu�Q��
            WOi.GuaranteeOi.cObj = rs("HWFONHWT")                       '�i�v�e�_�f�Z�x�ۏؕ��@�Q��
            WOi.GuaranteeOi.cJudg = rs("HWFONHWS")                      '�i�v�e�_�f�Z�x�ۏؕ��@�Q��
            WOi.GuaranteeCal = rs("HWFONMCL")                           '�i�v�e�_�f�Z�x�ʓ��v�Z
            If IsNull(rs("HWFONKHN")) = False Then HWFONKHN = rs("HWFONKHN")    '�i�v�e�_�f�Z�x�����p�x�Q���@04/04/15 ooba
            Set rs = Nothing
                
            WOi.Oi(0) = NtoZ2(.Fields("WFOI_MESDATA1").Value)           'Oi����l1
            WOi.Oi(1) = NtoZ2(.Fields("WFOI_MESDATA2").Value)           'Oi����l2
            WOi.Oi(2) = NtoZ2(.Fields("WFOI_MESDATA3").Value)           'Oi����l3
            WOi.Oi(3) = NtoZ2(.Fields("WFOI_MESDATA4").Value)           'Oi����l4
            WOi.Oi(4) = NtoZ2(.Fields("WFOI_MESDATA5").Value)           'Oi����l5
            WOi.Oi(5) = NtoZ2(.Fields("WFOI_MESDATA6").Value)           'Oi����l6
            WOi.Oi(6) = NtoZ2(.Fields("WFOI_MESDATA7").Value)           'Oi����l7
            WOi.Oi(7) = NtoZ2(.Fields("WFOI_MESDATA8").Value)           'Oi����l8
            WOi.Oi(8) = NtoZ2(.Fields("WFOI_MESDATA9").Value)           'Oi����l9
            WOi.Oi(9) = NtoZ2(.Fields("WFOI_MESDATA10").Value)          'Oi����l10
                
            .Fields("WFOI_ORG").Value = WFCORGCal(WOi.Oi(), WOi.GuaranteeOi, WOi.GuaranteeCal)      'WFOI_ORG�v�Z����
            
            '�ۏؕ��@="H"�A���AWFOI_ORG�v�Z���ʂ�-1�̏ꍇ�A�G���[�Ƃ���B2003/11/21 SystemBrain
'            If (WOi.GuaranteeOi.cJudg = "H") And (.Fields("WFOI_ORG").Value = -1) Then GoTo proc_exit
            '�ۏؕ��@�����̒ǉ��@04/04/15 ooba
            If ((WOi.GuaranteeOi.cJudg = "H") And CheckKHN(HWFONKHN, 2, sPos)) _
                And (.Fields("WFOI_ORG").Value = -1) Then GoTo proc_exit
        
        End If
    End With

    getTBCMY013WFOi = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY013WFOi = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :WFRs����(TBCMY013)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW�\����   (�V����يǗ�(SXL))
'          :hin             , I  ,tFullHinban       , �i��(�S�i�ԍ\����)
'�@�@      :sPos  �@�@�@    ,I   ,String �@         ,SXL�ʒu(TOP/BOT)   04/04/15 ooba
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :WFRs����(TBCMY013)�����ް����擾���ASXL�������\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
Private Function getTBCMY013WFRs(recXSDCW As c_cmzcrec, HIN As tFullHinban, sPos As String, recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim WRs     As W_RES                    'WFRs�\����
    Dim HWFRKHNN As String                  '�i�v�e���R�����p�x�Q��   04/04/15 ooba
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY013WFRs"
    
    getTBCMY013WFRs = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    With recX001
        .Fields("WFRS_SMPPOS").Value = -1                   'WFRS�����-ID����ʒu(SXL�ʒu���)
        .Fields("WFRS_NETSU").Value = ""                    'WFRS_�M��������
        .Fields("WFRS_ET").Value = ""                       'WFRS_�G�b�`���O����
        .Fields("WFRS_MES").Value = ""                      'WFRS_�v�����@
        .Fields("WFRS_MESDATA1").Value = -1                 'WFRS_����f�[�^���̂P
        .Fields("WFRS_MESDATA2").Value = -1                 'WFRS_����f�[�^���̂Q
        .Fields("WFRS_MESDATA3").Value = -1                 'WFRS_����f�[�^���̂R
        .Fields("WFRS_MESDATA4").Value = -1                 'WFRS_����f�[�^���̂S
        .Fields("WFRS_MESDATA5").Value = -1                 'WFRS_����f�[�^���̂T
        .Fields("WFRS_RRG").Value = -1                      'WFRS_RRG�v�Z����
    
        '-------------------- TBCMY013�̓ǂݍ���(WFRs) ----------------------------------------
        If (recXSDCW("WFINDRSCW").Value <> "0") And (recXSDCW("WFRESRS1CW").Value <> "0") Then
            sql = "select * from TBCMY013 "
            sql = sql & "where SAMPLEID = '" & recXSDCW("WFSMPLIDRSCW").Value & "' and "
            sql = sql & "      SPEC = 'RES'"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            .Fields("WFRS_SMPPOS").Value = recXSDCW("INPOSCW").Value        'WFRS�����-ID����ʒu(SXL�ʒu���)
            .Fields("WFRS_NETSU").Value = rs("NETSU")                       'WFRS_�M��������
            .Fields("WFRS_ET").Value = rs("ET")                             'WFRS_�G�b�`���O����
            .Fields("WFRS_MES").Value = rs("MES")                           'WFRS_�v�����@
            .Fields("WFRS_MESDATA1").Value = rs("MESDATA1")                 'WFRS_����f�[�^���̂P
            .Fields("WFRS_MESDATA2").Value = rs("MESDATA2")                 'WFRS_����f�[�^���̂Q
            .Fields("WFRS_MESDATA3").Value = rs("MESDATA3")                 'WFRS_����f�[�^���̂R
            .Fields("WFRS_MESDATA4").Value = rs("MESDATA4")                 'WFRS_����f�[�^���̂S
            .Fields("WFRS_MESDATA5").Value = rs("MESDATA5")                 'WFRS_����f�[�^���̂T
            Set rs = Nothing
                
            'WFRs_RRG
            sql = "select E021.HWFRSPOH, E018.HSXRSPOT, E018.HSXRSPOI, E021.HWFRHWYT, E021.HWFRHWYS, E021.HWFRMCAL, E021.HWFRKHNN "
            sql = sql & "from TBCME021 E021, TBCME018 E018 where "
            sql = sql & " E021.HINBAN = '" & HIN.hinban & "' and "
            sql = sql & " E021.MNOREVNO = " & HIN.mnorevno & " and "
            sql = sql & " E021.FACTORY = '" & HIN.factory & "' and "
            sql = sql & " E021.OPECOND = '" & HIN.opecond & "' and "
            sql = sql & " E018.HINBAN = E021.HINBAN and "
            sql = sql & " E018.MNOREVNO = E021.MNOREVNO and "
            sql = sql & " E018.FACTORY = E021.FACTORY and "
            sql = sql & " E018.OPECOND = E021.OPECOND"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
                
            WRs.GuaranteeRes.cMeth = rs("HWFRSPOH")                     ' �i�v�e���R����ʒu�Q��
            WRs.GuaranteeRes.cCount = rs("HSXRSPOT")                    ' �i�r�w���R����ʒu�Q�_
            WRs.GuaranteeRes.cPos = rs("HSXRSPOI")                      ' �i�r�w���R����ʒu�Q��
            WRs.GuaranteeRes.cObj = rs("HWFRHWYT")                      ' �i�v�e���R�ۏؕ��@�Q��
            WRs.GuaranteeRes.cJudg = rs("HWFRHWYS")                     ' �i�v�e���R�ۏؕ��@�Q��
            WRs.GuaranteeCal = rs("HWFRMCAL")                           ' �i�v�e���R�ʓ��v�Z
            If IsNull(rs("HWFRKHNN")) = False Then HWFRKHNN = rs("HWFRKHNN")    ' �i�v�e���R�����p�x�Q���@04/04/15 ooba
            Set rs = Nothing
                
            WRs.Res(0) = NtoZ2(.Fields("WFRS_MESDATA1").Value)          'Rs����l1
            WRs.Res(1) = NtoZ2(.Fields("WFRS_MESDATA2").Value)          'Rs����l2
            WRs.Res(2) = NtoZ2(.Fields("WFRS_MESDATA3").Value)          'Rs����l3
            WRs.Res(3) = NtoZ2(.Fields("WFRS_MESDATA4").Value)          'Rs����l4
            WRs.Res(4) = NtoZ2(.Fields("WFRS_MESDATA5").Value)          'Rs����l5
                
            .Fields("WFRS_RRG").Value = WFCRRGCal(WRs.Res(), WRs.GuaranteeRes, WRs.GuaranteeCal)        'WFRS_RRG�v�Z����
            
            '�ۏؕ��@="H"�A���AWFRS_RRG�v�Z���ʂ�-1�̏ꍇ�A�G���[�Ƃ���B2003/11/21 SystemBrain
'            If (WRs.GuaranteeRes.cJudg = "H") And (.Fields("WFRS_RRG").Value = -1) Then GoTo proc_exit
            '�ۏؕ��@�����̒ǉ��@04/04/15 ooba
            If ((WRs.GuaranteeRes.cJudg = "H") And CheckKHN(HWFRKHNN, 1, sPos)) _
                And (.Fields("WFRS_RRG").Value = -1) Then GoTo proc_exit
        
        End If
    End With

    getTBCMY013WFRs = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY013WFRs = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :WFDOi����(TBCMY013)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW�\����   (�V����يǗ�(SXL))
'          :j               , I  ,Integer           , DOi No
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002�\����(����_�ް�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :WFDOi����(TBCMY013)�����ް����擾���ASXL�������E����_�ް��\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
Private Function getTBCMY013WFDOi(recXSDCW As c_cmzcrec, j As Integer, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY013WFDOi"
    
    getTBCMY013WFDOi = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    'TBCMX001
    With recX001
        If j = 1 Then
            .Fields("WFDOI_SMPPOS").Value = -1              'WFDOI�����-ID����ʒu(SXL�ʒu���)
        End If
        .Fields("WFDOI_NETU_" & j).Value = ""               'WFDOI_�M��������_x
        .Fields("WFDOI_MES_" & j).Value = ""                'WFDOI_�v�����@_x
        .Fields("WFDOI_MESDATA1_" & j).Value = -1           'WFDOI_(�Ƽ��Oi-AfterOi)1_x
        .Fields("WFDOI_MESDATA2_" & j).Value = -1           'WFDOI_(�Ƽ��Oi-AfterOi)2_x
        .Fields("WFDOI_MESDATA3_" & j).Value = -1           'WFDOI_(�Ƽ��Oi-AfterOi)3_x
    End With
                
    'TBCMX002
    With recX002
        If j = 1 Then
            .Fields("WFDOI_SMPPOS").Value = -1              'WFDOI�����-ID����ʒu(SXL�ʒu���)
        End If
        .Fields("WFDOI" & j & "_NETSU").Value = " "         'WFDOI-x_�M��������
        .Fields("WFDOI" & j & "_MES").Value = " "           'WFDOI-x_�v�����@
        .Fields("WFDOI" & j & "_MESDATA1").Value = " "      'WFDOI-x_����l1
        .Fields("WFDOI" & j & "_MESDATA2").Value = " "      'WFDOI-x_����l2
        .Fields("WFDOI" & j & "_MESDATA3").Value = " "      'WFDOI-x_����l3
        .Fields("WFDOI" & j & "_MESDATA4").Value = " "      'WFDOI-x_����l4
        .Fields("WFDOI" & j & "_MESDATA5").Value = " "      'WFDOI-x_����l5
        .Fields("WFDOI" & j & "_MESDATA6").Value = " "      'WFDOI-x_����l6
        .Fields("WFDOI" & j & "_MESDATA7").Value = " "      'WFDOI-x_����l7
        .Fields("WFDOI" & j & "_MESDATA8").Value = " "      'WFDOI-x_����l8
        .Fields("WFDOI" & j & "_MESDATA9").Value = " "      'WFDOI-x_����l9
        .Fields("WFDOI" & j & "_MESDATA10").Value = " "     'WFDOI-x_����l10
        .Fields("WFDOI" & j & "_MESDATA11").Value = " "     'WFDOI-x_����l11
        .Fields("WFDOI" & j & "_MESDATA12").Value = " "     'WFDOI-x_����l12
        .Fields("WFDOI" & j & "_MESDATA13").Value = " "     'WFDOI-x_����l13
        .Fields("WFDOI" & j & "_MESDATA14").Value = " "     'WFDOI-x_����l14
        .Fields("WFDOI" & j & "_MESDATA15").Value = " "     'WFDOI-x_����l15
    End With
    
    '-------------------- TBCMY013�̓ǂݍ���(WFDOi) ----------------------------------------
    If (recXSDCW("WFINDDO" & j & "CW").Value <> "0") And (recXSDCW("WFRESDO" & j & "CW").Value <> "0") Then
        sql = "select * from TBCMY013 "
        sql = sql & "where SAMPLEID = '" & recXSDCW("WFSMPLIDDO" & j & "CW").Value & "' and "
        sql = sql & "      SPEC = 'DOI" & j & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            If .Fields("WFDOI_SMPPOS").Value = -1 Then
                .Fields("WFDOI_SMPPOS").Value = recXSDCW("INPOSCW").Value               'WFDOI�����-ID����ʒu(SXL�ʒu���)
            End If
            .Fields("WFDOI_NETU_" & j).Value = rs("NETSU")                              'WFDOI_�M��������_x
            .Fields("WFDOI_MES_" & j).Value = rs("MES")                                 'WFDOI_�v�����@_x
            .Fields("WFDOI_MESDATA1_" & j).Value = rs("MESDATA1") - rs("MESDATA4")      'WFDOI_(�Ƽ��Oi-AfterOi)1_x
            .Fields("WFDOI_MESDATA2_" & j).Value = rs("MESDATA2") - rs("MESDATA5")      'WFDOI_(�Ƽ��Oi-AfterOi)2_x
            .Fields("WFDOI_MESDATA3_" & j).Value = rs("MESDATA3") - rs("MESDATA6")      'WFDOI_(�Ƽ��Oi-AfterOi)3_x
        End With
            
        'TBCMX002
        With recX002
            If .Fields("WFDOI_SMPPOS").Value = -1 Then
                .Fields("WFDOI_SMPPOS").Value = recXSDCW("INPOSCW").Value               'WFDOI�����-ID����ʒu(SXL�ʒu���)
            End If
            .Fields("WFDOI" & j & "_NETSU").Value = rs("NETSU")                         'WFDOI-x_�M��������
            .Fields("WFDOI" & j & "_MES").Value = rs("MES")                             'WFDOI-x_�v�����@
            .Fields("WFDOI" & j & "_MESDATA1").Value = rs("MESDATA1")                   'WFDOI-x_����l1
            .Fields("WFDOI" & j & "_MESDATA2").Value = rs("MESDATA2")                   'WFDOI-x_����l2
            .Fields("WFDOI" & j & "_MESDATA3").Value = rs("MESDATA3")                   'WFDOI-x_����l3
            .Fields("WFDOI" & j & "_MESDATA4").Value = rs("MESDATA4")                   'WFDOI-x_����l4
            .Fields("WFDOI" & j & "_MESDATA5").Value = rs("MESDATA5")                   'WFDOI-x_����l5
            .Fields("WFDOI" & j & "_MESDATA6").Value = rs("MESDATA6")                   'WFDOI-x_����l6
            .Fields("WFDOI" & j & "_MESDATA7").Value = rs("MESDATA7")                   'WFDOI-x_����l7
            .Fields("WFDOI" & j & "_MESDATA8").Value = rs("MESDATA8")                   'WFDOI-x_����l8
            .Fields("WFDOI" & j & "_MESDATA9").Value = rs("MESDATA9")                   'WFDOI-x_����l9
            .Fields("WFDOI" & j & "_MESDATA10").Value = rs("MESDATA10")                 'WFDOI-x_����l10
            .Fields("WFDOI" & j & "_MESDATA11").Value = rs("MESDATA11")                 'WFDOI-x_����l11
            .Fields("WFDOI" & j & "_MESDATA12").Value = rs("MESDATA12")                 'WFDOI-x_����l12
            .Fields("WFDOI" & j & "_MESDATA13").Value = rs("MESDATA13")                 'WFDOI-x_����l13
            .Fields("WFDOI" & j & "_MESDATA14").Value = rs("MESDATA14")                 'WFDOI-x_����l14
            .Fields("WFDOI" & j & "_MESDATA15").Value = rs("MESDATA15")                 'WFDOI-x_����l15
        End With
        Set rs = Nothing
    End If

    getTBCMY013WFDOi = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY013WFDOi = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :WFOSF����(TBCMY013)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW�\����   (�V����يǗ�(SXL))
'          :j               , I  ,Integer           , OSF No
'          :hin             , I  ,tFullHinban       , �i��(�S�i�ԍ\����)
'�@�@      :sPos  �@�@�@    , I  ,String �@         , SXL�ʒu(TOP/BOT)   04/04/15 ooba
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002�\����(����_�ް�)
'�@�@      :sTblName �@     , I  ,String �@         , �e�[�u����   11/06/24 Marushita�@MIN�l�ǉ��Ή�
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :WFOSF����(TBCMY013)�����ް����擾���ASXL�������E����_�ް��\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
'Private Function getTBCMY013WFOSF(recXSDCW As c_cmzcrec, j As Integer, HIN As tFullHinban, sPos As String, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
Private Function getTBCMY013WFOSF(recXSDCW As c_cmzcrec, j As Integer, HIN As tFullHinban, sPos As String, recX001 As c_cmzcrec, recX002 As c_cmzcrec, sTblName As String) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim wos         As W_OSF                    'OSF�\����
    Dim keisu       As Double
    Dim k           As Integer
    Dim HWFOSFKN    As String                   '�i�v�e�n�r�e�����p�x�Q��   04/04/15 ooba
    Dim nFlg        As Integer                  'MIN�l�Z�b�g����p
    
    Const keisu1 As Double = 1.8248175
    Const keisu2 As Double = 1.8518519
    Const keisu3 As Double = 1.9230769
    Const keisu4 As Double = 3.649635
    Const keisu5 As Double = 3.7037037
    Const keisu6 As Double = 3.8461538
    '' 2006/09/25 SMP)kondoh Add -s-
    Const keisu7 As Double = 7.6923077
    '' 2006/09/25 SMP)kondoh Add -e-
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY013WFOSF"
    
    getTBCMY013WFOSF = FUNCTION_RETURN_FAILURE

    '>>>>> 2011/06/27 SETsw)Marushita WFOSFx_���莞��MIN�l_x�̃Z�b�g�Ή�
    'MIN�l�̍��ڂ����݂���ꍇ�̂݃Z�b�g
    If FieldCheck(sTblName, "WFOSF" & j & "_MIN") = FUNCTION_RETURN_SUCCESS Then
        nFlg = 1
    Else
        nFlg = 0
    End If
    '<<<<< 2011/06/27 SETsw)Marushita WFOSFx_���莞��MIN�l_x�̃Z�b�g�Ή�
    
    '-------------------- �����ر ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("WFOSF" & j & "_SMPPOS").Value = -1         'WFOSFx�����-ID����ʒu(SXL�ʒu���)
        .Fields("WFOSF" & j & "_NETSU").Value = ""          'WFOSFx_�M��������
        .Fields("WFOSF" & j & "_ET").Value = ""             'WFOSFx_�G�b�`���O����
        .Fields("WFOSF" & j & "_MES").Value = ""            'WFOSFx_�v�����@
        .Fields("WFOSF" & j & "_MAX").Value = -1            'WFOSFx_���莞��MAX�l_x
        .Fields("WFOSF" & j & "_AVE").Value = -1            'WFOSFx_���莞��AVE�l_x
        If nFlg = 1 Then
            .Fields("WFOSF" & j & "_MIN").Value = -1        'WFOSFx_���莞��MIN�l_x
            .Fields("WFOSF4_MIN").Value = -1                'WFOSF4_���莞��MIN�l_4
        End If
        '��Add SIRD�]���Ή��@2010/04/19 Y.Hitomi
        .Fields("WFOSF4_SMPPOS").Value = -1                 'WFOSF4�����-ID����ʒu(SXL�ʒu���)
        .Fields("WFOSF4_NETSU").Value = ""                  'WFOSF4_�M��������
        .Fields("WFOSF4_ET").Value = ""                     'WFOSF4_�G�b�`���O����
        .Fields("WFOSF4_MES").Value = ""                    'WFOSF4_�v�����@
        .Fields("WFOSF4_MAX").Value = -1                    'WFOSF4_���莞��MAX�l_4
        .Fields("WFOSF4_AVE").Value = -1                    'WFOSF4_���莞��AVE�l_4
        '��Add SIRD�]���Ή��@2010/04/19 Y.Hitomi
    End With
                
    'TBCMX002
    With recX002
        .Fields("WFOSF" & j & "_SMPPOS").Value = -1         'WFOSFx�����-ID����ʒu(SXL�ʒu���)
        .Fields("WFOSF" & j & "_NETSU").Value = " "         'WFOSFx_�M��������
        .Fields("WFOSF" & j & "_ET").Value = " "            'WFOSFx_�G�b�`���O����
        .Fields("WFOSF" & j & "_MES").Value = " "           'WFOSFx_�v�����@
        .Fields("WFOSF" & j & "_DKAN").Value = " "          'WFOSFx_�c�j�A�j�[������
        .Fields("WFOSF" & j & "_MESDATA1").Value = " "      'WFOSFx����_1
        .Fields("WFOSF" & j & "_MESDATA2").Value = " "      'WFOSFx����_2
        .Fields("WFOSF" & j & "_MESDATA3").Value = " "      'WFOSFx����_3
        .Fields("WFOSF" & j & "_MESDATA4").Value = " "      'WFOSFx����_4
        .Fields("WFOSF" & j & "_MESDATA5").Value = " "      'WFOSFx����_5
        .Fields("WFOSF" & j & "_MESDATA6").Value = " "      'WFOSFx����_6
        .Fields("WFOSF" & j & "_MESDATA7").Value = " "      'WFOSFx����_7
        .Fields("WFOSF" & j & "_MESDATA8").Value = " "      'WFOSFx����_8
        .Fields("WFOSF" & j & "_MESDATA9").Value = " "      'WFOSFx����_9
        .Fields("WFOSF" & j & "_MESDATA10").Value = " "     'WFOSFx����_10
        .Fields("WFOSF" & j & "_MESDATA11").Value = " "     'WFOSFx����_11
        .Fields("WFOSF" & j & "_MESDATA12").Value = " "     'WFOSFx����_12
        .Fields("WFOSF" & j & "_MESDATA13").Value = " "     'WFOSFx����_13
        .Fields("WFOSF" & j & "_MESDATA14").Value = " "     'WFOSFx����_14
        .Fields("WFOSF" & j & "_MESDATA15").Value = " "     'WFOSFx����_15
        
        '��Add SIRD�]���Ή��@2010/04/19 Y.Hitomi
        .Fields("WFOSF4_SMPPOS").Value = -1         'WFOSF4�����-ID����ʒu(SXL�ʒu���)
        .Fields("WFOSF4_NETSU").Value = " "         'WFOSF4_�M��������
        .Fields("WFOSF4_ET").Value = " "            'WFOSF4_�G�b�`���O����
        .Fields("WFOSF4_MES").Value = " "           'WFOSF4_�v�����@
        .Fields("WFOSF4_DKAN").Value = " "          'WFOSF4_�c�j�A�j�[������
        .Fields("WFOSF4_MESDATA1").Value = " "      'WFOSF4����_1
        .Fields("WFOSF4_MESDATA2").Value = " "      'WFOSF4����_2
        .Fields("WFOSF4_MESDATA3").Value = " "      'WFOSF4����_3
        .Fields("WFOSF4_MESDATA4").Value = " "      'WFOSF4����_4
        .Fields("WFOSF4_MESDATA5").Value = " "      'WFOSF4����_5
        .Fields("WFOSF4_MESDATA6").Value = " "      'WFOSF4����_6
        .Fields("WFOSF4_MESDATA7").Value = " "      'WFOSF4����_7
        .Fields("WFOSF4_MESDATA8").Value = " "      'WFOSF4����_8
        .Fields("WFOSF4_MESDATA9").Value = " "      'WFOSF4����_9
        .Fields("WFOSF4_MESDATA10").Value = " "     'WFOSF4����_10
        .Fields("WFOSF4_MESDATA11").Value = " "     'WFOSF4����_11
        .Fields("WFOSF4_MESDATA12").Value = " "     'WFOSF4����_12
        .Fields("WFOSF4_MESDATA13").Value = " "     'WFOSF4����_13
        .Fields("WFOSF4_MESDATA14").Value = " "     'WFOSF4����_14
        .Fields("WFOSF4_MESDATA15").Value = " "     'WFOSF4����_15
        '��Add SIRD�]���Ή��@2010/04/19 Y.Hitomi
    End With
    
    '-------------------- TBCMY013�̓ǂݍ���(WFOSF) ----------------------------------------
    If (recXSDCW("WFINDL" & j & "CW").Value <> "0") And (recXSDCW("WFRESL" & j & "CW").Value <> "0") Then
        sql = "select * from TBCMY013 "
        sql = sql & "where SAMPLEID = '" & recXSDCW("WFSMPLIDL" & j & "CW").Value & "' and "
        sql = sql & "      SPEC = 'OSF" & j & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            .Fields("WFOSF" & j & "_SMPPOS").Value = recXSDCW("INPOSCW").Value          'WFOSFx�����-ID����ʒu(SXL�ʒu���)
            .Fields("WFOSF" & j & "_NETSU").Value = rs("NETSU")                         'WFOSFx_�M��������
            .Fields("WFOSF" & j & "_ET").Value = rs("ET")                               'WFOSFx_�G�b�`���O����
            .Fields("WFOSF" & j & "_MES").Value = rs("MES")                             'WFOSFx_�v�����@
        End With
            
        'TBCMX002
        With recX002
            .Fields("WFOSF" & j & "_SMPPOS").Value = recXSDCW("INPOSCW").Value          'WFOSFx�����-ID����ʒu(SXL�ʒu���)
            .Fields("WFOSF" & j & "_NETSU").Value = rs("NETSU")                         'WFOSFx_�M��������
            .Fields("WFOSF" & j & "_ET").Value = rs("ET")                               'WFOSFx_�G�b�`���O����
            .Fields("WFOSF" & j & "_MES").Value = rs("MES")                             'WFOSFx_�v�����@
            .Fields("WFOSF" & j & "_DKAN").Value = rs("DKAN")                           'WFOSFx_�c�j�A�j�[������
            .Fields("WFOSF" & j & "_MESDATA1").Value = rs("MESDATA1")                   'WFOSFx����_1
            .Fields("WFOSF" & j & "_MESDATA2").Value = rs("MESDATA2")                   'WFOSFx����_2
            .Fields("WFOSF" & j & "_MESDATA3").Value = rs("MESDATA3")                   'WFOSFx����_3
            .Fields("WFOSF" & j & "_MESDATA4").Value = rs("MESDATA4")                   'WFOSFx����_4
            .Fields("WFOSF" & j & "_MESDATA5").Value = rs("MESDATA5")                   'WFOSFx����_5
            .Fields("WFOSF" & j & "_MESDATA6").Value = rs("MESDATA6")                   'WFOSFx����_6
            .Fields("WFOSF" & j & "_MESDATA7").Value = rs("MESDATA7")                   'WFOSFx����_7
            .Fields("WFOSF" & j & "_MESDATA8").Value = rs("MESDATA8")                   'WFOSFx����_8
            .Fields("WFOSF" & j & "_MESDATA9").Value = rs("MESDATA9")                   'WFOSFx����_9
            .Fields("WFOSF" & j & "_MESDATA10").Value = rs("MESDATA10")                 'WFOSFx����_10
            .Fields("WFOSF" & j & "_MESDATA11").Value = rs("MESDATA11")                 'WFOSFx����_11
            .Fields("WFOSF" & j & "_MESDATA12").Value = rs("MESDATA12")                 'WFOSFx����_12
            .Fields("WFOSF" & j & "_MESDATA13").Value = rs("MESDATA13")                 'WFOSFx����_13
            .Fields("WFOSF" & j & "_MESDATA14").Value = rs("MESDATA14")                 'WFOSFx����_14
            .Fields("WFOSF" & j & "_MESDATA15").Value = rs("MESDATA15")                 'WFOSFx����_15
        End With
        Set rs = Nothing
        
        'WFOSF_MAX,AVE
        sql = "select HWFOF" & j & "SH, HWFOF" & j & "ST, HWFOF" & j & "SR, HWFOF" & j & "HT, "
        sql = sql & "HWFOF" & j & "HS, HWFOF" & j & "AX, HWFOF" & j & "MX, HWFOSF" & j & "PTK, "
        sql = sql & "HWFOF" & j & "KN "
        sql = sql & "from TBCME029 where "
        sql = sql & "HINBAN = '" & HIN.hinban & "' and "
        sql = sql & "MNOREVNO = " & HIN.mnorevno & " and "
        sql = sql & "FACTORY = '" & HIN.factory & "' and "
        sql = sql & "OPECOND = '" & HIN.opecond & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
            
        If IsNull(rs("HWFOF" & j & "SH")) = False Then wos.GuaranteeOsf.cMeth = rs("HWFOF" & j & "SH")      '�i�v�e�n�r�ex����ʒu�Q��
        If IsNull(rs("HWFOF" & j & "ST")) = False Then wos.GuaranteeOsf.cCount = rs("HWFOF" & j & "ST")     '�i�v�e�n�r�ex����ʒu�Q�_
        If IsNull(rs("HWFOF" & j & "SR")) = False Then wos.GuaranteeOsf.cPos = rs("HWFOF" & j & "SR")       '�i�v�e�n�r�ex����ʒu�Q��
        If IsNull(rs("HWFOF" & j & "HT")) = False Then wos.GuaranteeOsf.cObj = rs("HWFOF" & j & "HT")       '�i�v�e�n�r�ex�ۏؕ��@�Q��
        If IsNull(rs("HWFOF" & j & "HS")) = False Then wos.GuaranteeOsf.cJudg = rs("HWFOF" & j & "HS")      '�i�v�e�n�r�ex�ۏؕ��@�Q��
        If IsNull(rs("HWFOF" & j & "AX")) = False Then wos.SpecOsfAveMax = rs("HWFOF" & j & "AX")           '�i�v�e�n�r�ex���Ϗ��
        If IsNull(rs("HWFOF" & j & "MX")) = False Then wos.SpecOsfMax = rs("HWFOF" & j & "MX")              '�i�v�e�n�r�ex���
        If IsNull(rs("HWFOSF" & j & "PTK")) = False Then wos.JudgDataPTK = rs("HWFOSF" & j & "PTK")         '�i�v�e�n�r�ex�p�^���敪
        If IsNull(rs("HWFOF" & j & "KN")) = False Then HWFOSFKN = rs("HWFOF" & j & "KN")                    '�i�v�e�n�r�e�����p�x�Q���@04/04/15 ooba
        Set rs = Nothing
            
        If wos.GuaranteeOsf.cMeth = "5" And wos.GuaranteeOsf.cCount = "5" And wos.GuaranteeOsf.cPos = "3" Then
            keisu = keisu1
        ElseIf wos.GuaranteeOsf.cMeth = "5" And wos.GuaranteeOsf.cCount = "5" And wos.GuaranteeOsf.cPos = "5" Then
            keisu = keisu2
        ElseIf wos.GuaranteeOsf.cMeth = "5" And wos.GuaranteeOsf.cCount = "5" And wos.GuaranteeOsf.cPos = "A" Then
            keisu = keisu3
        ElseIf wos.GuaranteeOsf.cMeth = "6" And wos.GuaranteeOsf.cCount = "5" And wos.GuaranteeOsf.cPos = "3" Then
            keisu = keisu4
        ElseIf wos.GuaranteeOsf.cMeth = "6" And wos.GuaranteeOsf.cCount = "5" And wos.GuaranteeOsf.cPos = "5" Then
            keisu = keisu5
        ElseIf wos.GuaranteeOsf.cMeth = "6" And wos.GuaranteeOsf.cCount = "5" And wos.GuaranteeOsf.cPos = "A" Then
            keisu = keisu6
        '' 2006/09/25 SMP)kondoh Add -s-
        ElseIf wos.GuaranteeOsf.cMeth = "E" And wos.GuaranteeOsf.cCount = "5" And wos.GuaranteeOsf.cPos = "A" Then
            keisu = keisu7
        '' 2006/09/25 SMP)kondoh Add -e-
        Else
            keisu = -1
'            GoTo proc_exit
        End If
            
        If keisu <> -1 Then
            With recX002
                wos.OSF(0) = NtoZ2(.Fields("WFOSF" & j & "_MESDATA1").Value)                   'OSF����l1
                wos.OSF(1) = NtoZ2(.Fields("WFOSF" & j & "_MESDATA2").Value)                   'OSF����l2
                wos.OSF(2) = NtoZ2(.Fields("WFOSF" & j & "_MESDATA3").Value)                   'OSF����l3
                wos.OSF(3) = NtoZ2(.Fields("WFOSF" & j & "_MESDATA4").Value)                   'OSF����l4
                wos.OSF(4) = NtoZ2(.Fields("WFOSF" & j & "_MESDATA5").Value)                   'OSF����l5
                For k = 0 To 4
                    wos.OSF(k) = IIf(wos.OSF(k) <> -1, wos.OSF(k) * keisu, -1)
                Next
            End With
            
            recX001.Fields("WFOSF" & j & "_MAX").Value = JudgMax(wos.OSF())     'WFOSFx_���莞��MAX�l_x
            recX001.Fields("WFOSF" & j & "_AVE").Value = JudgAve(wos.OSF())     'WFOSFx_���莞��AVE�l_x
            '>>>>> 2011/06/24 SETsw)Marushita WFOSFx_���莞��MIN�l_x�̃Z�b�g�Ή�
            'MIN�l�̍��ڂ����݂���ꍇ�̂݃Z�b�g
            If nFlg = 1 Then
                recX001.Fields("WFOSF" & j & "_MIN").Value = JudgMin(wos.OSF())     'WFOSFx_���莞��MIN�l_x
                recX001.Fields("WFOSF4_MIN").Value = -1     'WFOSFx_���莞��MIN�l_4
            End If
            '<<<<< 2011/06/24 SETsw)Marushita WFOSFx_���莞��MIN�l_x�̃Z�b�g�Ή�
        End If
            
        '�ۏؕ��@="H"�A���AWFOSF��MAX,AVE�l��-1�̏ꍇ�A�G���[�Ƃ���B2003/11/21 SystemBrain
        '�ۏؕ��@�����̒ǉ��@04/04/15 ooba
        ''�ۏؕ��@="H"�A���AWFRS_RRG�v�Z���ʂ�-1�̏ꍇ�A�G���[�Ƃ���B2003/11/21 SystemBrain
'        If (wos.GuaranteeOsf.cJudg = "H") And
        If ((wos.GuaranteeOsf.cJudg = "H") And CheckKHN(HWFOSFKN, j + 2, sPos)) And _
           (recX001.Fields("WFOSF" & j & "_MAX").Value = -1 Or _
            recX001.Fields("WFOSF" & j & "_AVE").Value = -1) Then GoTo proc_exit
        '�H�H�H�H�HMIN�l���`�F�b�N�H�H�H�H�H
    End If

    getTBCMY013WFOSF = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY013WFOSF = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :WFBMD����(TBCMY013)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW�\����   (�V����يǗ�(SXL))
'          :j               , I  ,Integer           , BMD No
'          :hin             , I  ,tFullHinban       , �i��(�S�i�ԍ\����)
'�@�@      :sPos  �@�@�@    ,I   ,String �@         ,SXL�ʒu(TOP/BOT)   04/04/15 ooba
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002�\����(����_�ް�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :WFBMD����(TBCMY013)�����ް����擾���ASXL�������E����_�ް��\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
Private Function getTBCMY013WFBMD(recXSDCW As c_cmzcrec, j As Integer, HIN As tFullHinban, sPos As String, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim bm          As W_BMD                    'BMD�\����
    Dim k           As Integer
    Dim HWFBMDKN    As String                   '�i�v�e�a�l�c�����p�x�Q��   04/04/15 ooba
    Dim JData(4)    As Double                   'MAX�l�Z�o�p�@06/09/06 ooba
    
    '' 2006/09/25 SMP)kondoh Del -s-
''    Const keisu As Double = 1        'BMD�ׂ��搔�ύX�Ή��@2003/05/19 osawa
    '' 2006/09/25 SMP)kondoh Del -e-
    '' 2006/09/25 SMP)kondoh Add -s-
    Dim keisu As Double
    Const keisu1 As Double = 10000
    Const keisu2 As Double = 10000
    Const keisu3 As Double = 10000
    Const keisu4 As Double = 10000
    Const keisu5 As Double = 10000
    Const keisu6 As Double = 333000
    Const keisu7 As Double = 10000
    Const keisu8 As Double = 10000
    '' 2006/09/25 SMP)kondoh Add -e-

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY013WFBMD"
    
    getTBCMY013WFBMD = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("WFBMD" & j & "_SMPPOS").Value = -1         'WFBMDx�����-ID����ʒu(SXL�ʒu���)
        .Fields("WFBMD" & j & "_NETSU").Value = ""          'WFBMDx_�M��������
        .Fields("WFBMD" & j & "_ET").Value = ""             'WFBMDx_�G�b�`���O����
        .Fields("WFBMD" & j & "_MES").Value = ""            'WFBMDx_�v�����@
        .Fields("WFBMD" & j & "_MAX").Value = -1            'WFBMDx_���莞��MAX�l_x
        .Fields("WFBMD" & j & "_AVE").Value = -1            'WFBMDx_���莞��AVE�l_x
        .Fields("WFBMD" & j & "_MIN").Value = -1            'WFBMDx_���莞��MIN�l_x
        .Fields("WFBMD" & j & "_MB").Value = -1             'WFBMDx_���莞�̖ʓ����z
    End With
                
    'TBCMX002
    With recX002
        .Fields("WFBMD" & j & "_SMPPOS").Value = -1         'WFBMDx�����-ID����ʒu(SXL�ʒu���)
        .Fields("WFBMD" & j & "_NETSU").Value = " "         'WFBMDx_�M��������
        .Fields("WFBMD" & j & "_ET").Value = " "            'WFBMDx_�G�b�`���O����
        .Fields("WFBMD" & j & "_MES").Value = " "           'WFBMDx_�v�����@
        .Fields("WFBMD" & j & "_DKAN").Value = " "          'WFBMDx_�c�j�A�j�[������
        .Fields("WFBMD" & j & "_MESDATA1").Value = " "      'WFBMDx����_1
        .Fields("WFBMD" & j & "_MESDATA2").Value = " "      'WFBMDx����_2
        .Fields("WFBMD" & j & "_MESDATA3").Value = " "      'WFBMDx����_3
        .Fields("WFBMD" & j & "_MESDATA4").Value = " "      'WFBMDx����_4
        .Fields("WFBMD" & j & "_MESDATA5").Value = " "      'WFBMDx����_5
        .Fields("WFBMD" & j & "_MESDATA6").Value = " "      'WFBMDx����_6
        .Fields("WFBMD" & j & "_MESDATA7").Value = " "      'WFBMDx����_7
        .Fields("WFBMD" & j & "_MESDATA8").Value = " "      'WFBMDx����_8
        .Fields("WFBMD" & j & "_MESDATA9").Value = " "      'WFBMDx����_9
        .Fields("WFBMD" & j & "_MESDATA10").Value = " "     'WFBMDx����_10
        .Fields("WFBMD" & j & "_MESDATA11").Value = " "     'WFBMDx����_11
        .Fields("WFBMD" & j & "_MESDATA12").Value = " "     'WFBMDx����_12
        .Fields("WFBMD" & j & "_MESDATA13").Value = " "     'WFBMDx����_13
        .Fields("WFBMD" & j & "_MESDATA14").Value = " "     'WFBMDx����_14
        .Fields("WFBMD" & j & "_MESDATA15").Value = " "     'WFBMDx����_15
    End With
    
    '-------------------- TBCMY013�̓ǂݍ���(WFBMD) ----------------------------------------
    If (recXSDCW("WFINDB" & j & "CW").Value <> "0") And (recXSDCW("WFRESB" & j & "CW").Value <> "0") Then
        sql = "select * from TBCMY013 "
        sql = sql & "where SAMPLEID = '" & recXSDCW("WFSMPLIDB" & j & "CW").Value & "' and "
        sql = sql & "      SPEC = 'BMD" & j & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            .Fields("WFBMD" & j & "_SMPPOS").Value = recXSDCW("INPOSCW").Value          'WFBMDx�����-ID����ʒu(SXL�ʒu���)
            .Fields("WFBMD" & j & "_NETSU").Value = rs("NETSU")                         'WFBMDx_�M��������
            .Fields("WFBMD" & j & "_ET").Value = rs("ET")                               'WFBMDx_�G�b�`���O����
            .Fields("WFBMD" & j & "_MES").Value = rs("MES")                             'WFBMDx_�v�����@
        End With
            
        'TBCMX002
        With recX002
            .Fields("WFBMD" & j & "_SMPPOS").Value = recXSDCW("INPOSCW").Value          'WFBMDx�����-ID����ʒu(SXL�ʒu���)
            .Fields("WFBMD" & j & "_NETSU").Value = rs("NETSU")                         'WFBMDx_�M��������
            .Fields("WFBMD" & j & "_ET").Value = rs("ET")                               'WFBMDx_�G�b�`���O����
            .Fields("WFBMD" & j & "_MES").Value = rs("MES")                             'WFBMDx_�v�����@
            .Fields("WFBMD" & j & "_DKAN").Value = rs("DKAN")                           'WFBMDx_�c�j�A�j�[������
            .Fields("WFBMD" & j & "_MESDATA1").Value = rs("MESDATA1")                   'WFBMDx����_1
            .Fields("WFBMD" & j & "_MESDATA2").Value = rs("MESDATA2")                   'WFBMDx����_2
            .Fields("WFBMD" & j & "_MESDATA3").Value = rs("MESDATA3")                   'WFBMDx����_3
            .Fields("WFBMD" & j & "_MESDATA4").Value = rs("MESDATA4")                   'WFBMDx����_4
            .Fields("WFBMD" & j & "_MESDATA5").Value = rs("MESDATA5")                   'WFBMDx����_5
            .Fields("WFBMD" & j & "_MESDATA6").Value = rs("MESDATA6")                   'WFBMDx����_6
            .Fields("WFBMD" & j & "_MESDATA7").Value = rs("MESDATA7")                   'WFBMDx����_7
            .Fields("WFBMD" & j & "_MESDATA8").Value = rs("MESDATA8")                   'WFBMDx����_8
            .Fields("WFBMD" & j & "_MESDATA9").Value = rs("MESDATA9")                   'WFBMDx����_9
            .Fields("WFBMD" & j & "_MESDATA10").Value = rs("MESDATA10")                 'WFBMDx����_10
            .Fields("WFBMD" & j & "_MESDATA11").Value = rs("MESDATA11")                 'WFBMDx����_11
            .Fields("WFBMD" & j & "_MESDATA12").Value = rs("MESDATA12")                 'WFBMDx����_12
            .Fields("WFBMD" & j & "_MESDATA13").Value = rs("MESDATA13")                 'WFBMDx����_13
            .Fields("WFBMD" & j & "_MESDATA14").Value = rs("MESDATA14")                 'WFBMDx����_14
            .Fields("WFBMD" & j & "_MESDATA15").Value = rs("MESDATA15")                 'WFBMDx����_15
        End With
        Set rs = Nothing
                    
        'WFBMD_MAX,MIN,AVE,MBP
        sql = "select HWFBM" & j & "SH, HWFBM" & j & "ST, HWFBM" & j & "SR, HWFBM" & j & "HT, "
        sql = sql & "HWFBM" & j & "HS, HWFBM" & j & "AN, HWFBM" & j & "AX, HWFBM" & j & "MBP, HWFBM" & j & "MCL, "
        sql = sql & "HWFBM" & j & "KN "
        sql = sql & "from TBCME029 where "
        sql = sql & "HINBAN = '" & HIN.hinban & "' and "
        sql = sql & "MNOREVNO = " & HIN.mnorevno & " and "
        sql = sql & "FACTORY = '" & HIN.factory & "' and "
        sql = sql & "OPECOND = '" & HIN.opecond & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
                    
        If IsNull(rs("HWFBM" & j & "SH")) = False Then bm.GuaranteeBmd.cMeth = rs("HWFBM" & j & "SH")       '�i�v�e�a�l�cx����ʒu�Q��
        If IsNull(rs("HWFBM" & j & "ST")) = False Then bm.GuaranteeBmd.cCount = rs("HWFBM" & j & "ST")      '�i�v�e�a�l�cx����ʒu�Q�_
        If IsNull(rs("HWFBM" & j & "SR")) = False Then bm.GuaranteeBmd.cPos = rs("HWFBM" & j & "SR")        '�i�v�e�a�l�cx����ʒu�Q��
        If IsNull(rs("HWFBM" & j & "HT")) = False Then bm.GuaranteeBmd.cObj = rs("HWFBM" & j & "HT")        '�i�v�e�a�l�cx�ۏؕ��@�Q��
        If IsNull(rs("HWFBM" & j & "HS")) = False Then bm.GuaranteeBmd.cJudg = rs("HWFBM" & j & "HS")       '�i�v�e�a�l�cx�ۏؕ��@�Q��
        If IsNull(rs("HWFBM" & j & "AN")) = False Then bm.SpecBmdAveMin = rs("HWFBM" & j & "AN")            '�i�v�e�a�l�cx���ω���
        If IsNull(rs("HWFBM" & j & "AX")) = False Then bm.SpecBmdAveMax = rs("HWFBM" & j & "AX")            '�i�v�e�a�l�cx���Ϗ��
        If IsNull(rs("HWFBM" & j & "MBP")) = False Then bm.SpecBmdMBP = rs("HWFBM" & j & "MBP")             '�i�v�e�a�l�cx�ʓ����z
        If IsNull(rs("HWFBM" & j & "MCL")) = False Then bm.SpecBmdMCL = rs("HWFBM" & j & "MCL")             '�i�v�e�a�l�cx�ʓ��v�Z
        If IsNull(rs("HWFBM" & j & "KN")) = False Then HWFBMDKN = rs("HWFBM" & j & "KN")                    '�i�v�e�a�l�c�����p�x�Q���@04/04/15 ooba
        Set rs = Nothing

        '' 2006/09/25 SMP)kondoh Add -s-
        If bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "4" And bm.GuaranteeBmd.cPos = "H" Then
            keisu = keisu1
        ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "3" And bm.GuaranteeBmd.cPos = "H" Then
            keisu = keisu2
        ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "4" And bm.GuaranteeBmd.cPos = "8" Then
            keisu = keisu3
        ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "3" And bm.GuaranteeBmd.cPos = "8" Then
            keisu = keisu4
        ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "5" And bm.GuaranteeBmd.cPos = "A" Then
            keisu = keisu5
        ElseIf bm.GuaranteeBmd.cMeth = "G" And bm.GuaranteeBmd.cCount = "3" And bm.GuaranteeBmd.cPos = "8" Then
            keisu = keisu6
        ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "5" And bm.GuaranteeBmd.cPos = "8" Then
            keisu = keisu7
        ElseIf bm.GuaranteeBmd.cMeth = "8" And bm.GuaranteeBmd.cCount = "4" And bm.GuaranteeBmd.cPos = "8" Then
            keisu = keisu8
        Else
            keisu = -1
        End If
        '' 2006/09/25 SMP)kondoh Add -e-

        '' 2006/09/25 SMP)kondoh Add -s-
        If keisu <> -1 Then
        '' 2006/09/25 SMP)kondoh Add -e-
                    
            With recX002
                bm.BMD(0) = NtoZ2(.Fields("WFBMD" & j & "_MESDATA1").Value)     'BMD����l1
                bm.BMD(1) = NtoZ2(.Fields("WFBMD" & j & "_MESDATA2").Value)     'BMD����l2
                bm.BMD(2) = NtoZ2(.Fields("WFBMD" & j & "_MESDATA3").Value)     'BMD����l3
                bm.BMD(3) = NtoZ2(.Fields("WFBMD" & j & "_MESDATA4").Value)     'BMD����l4
                bm.BMD(4) = NtoZ2(.Fields("WFBMD" & j & "_MESDATA5").Value)     'BMD����l5
        
                For k = 0 To 4                                      ' 2003/05/20 ooba
                '   ' 2006/09/25 SMP)kondoh Add -s-
''                    bm.BMD(k) = IIf(bm.BMD(k) <> -1, bm.BMD(k) * keisu, -1)
                    bm.BMD(k) = IIf(bm.BMD(k) <> -1, bm.BMD(k) * CDbl(keisu / 10000), -1)
                    '' 2006/09/25 SMP)kondoh Add -e-
                Next
            End With
            
            ''06/09/06 ooba START =============================================================>
            '���躰�ނ� "F"�FMAX(2,4�_��)�C"G"�FMAX(2,3,4�_��) �̏ꍇ��MAX�l�ɂ��̒l���
            If bm.GuaranteeBmd.cJudg = JudgCodeW01 And _
                (bm.GuaranteeBmd.cObj = ObjCode10 Or bm.GuaranteeBmd.cObj = ObjCode11) Then
            
                If GetWfJudgData(WFBMD_JUDG, bm.GuaranteeBmd, bm.BMD(), JData()) = FUNCTION_RETURN_FAILURE Then
                    GoTo proc_exit
                End If
                recX001.Fields("WFBMD" & j & "_MAX").Value = JData(0)
            Else
                recX001.Fields("WFBMD" & j & "_MAX").Value = JudgMax(bm.BMD())
            End If
            ''06/09/06 ooba END ===============================================================>
            
    '        recX001.Fields("WFBMD" & j & "_MAX").Value = JudgMax(bm.BMD())          'WFBMDx_���莞��MAX�l_x
            recX001.Fields("WFBMD" & j & "_AVE").Value = JudgAve(bm.BMD())          'WFBMDx_���莞��AVE�l_x
            recX001.Fields("WFBMD" & j & "_MIN").Value = JudgMin(bm.BMD())          'WFBMDx_���莞��MIN�l_x
            If bm.SpecBmdMCL = "P " Then
                recX001.Fields("WFBMD" & j & "_MB").Value = JudgBmdMBP(bm.BMD())    'WFBMDx_���莞�̖ʓ����z
            Else
                recX001.Fields("WFBMD" & j & "_MB").Value = 0                       '�ʓ����z��"P"�ȊO�̎��͌v�Z���ʂ�0�Ƃ���@2003/06/06 ooba
            End If
            
        '' 2006/09/25 SMP)kondoh Add -s-
        End If
        '' 2006/09/25 SMP)kondoh Add -e-
            
        '�ۏؕ��@="H"�A���AWFRS_RRG�v�Z���ʂ�-1�̏ꍇ�A�G���[�Ƃ���B2003/11/21 SystemBrain
'        If (bm.GuaranteeBmd.cJudg = "H") And (recX001.Fields("WFBMD" & j & "_MB").Value = -1) Then GoTo proc_exit
        '�ۏؕ��@�����̒ǉ��@04/04/15 ooba
        If ((bm.GuaranteeBmd.cJudg = "H") And CheckKHN(HWFBMDKN, j + 6, sPos)) _
            And (recX001.Fields("WFBMD" & j & "_MB").Value = -1) Then GoTo proc_exit
        
    End If

    getTBCMY013WFBMD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY013WFBMD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :WFDSOD����(TBCMY013)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW�\����   (�V����يǗ�(SXL))
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002�\����(����_�ް�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :WFDSOD����(TBCMY013)�����ް����擾���ASXL�������E����_�ް��\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
Private Function getTBCMY013WFDSOD(recXSDCW As c_cmzcrec, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY013WFDSOD"
    
    getTBCMY013WFDSOD = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("WFDSOD_SMPPOS").Value = -1         'WFDSOD�����-ID����ʒu(SXL�ʒu���)
        .Fields("WFDSOD_NETSU").Value = ""          'WFDSOD_�M��������
        .Fields("WFDSOD_ET").Value = ""             'WFDSOD_�G�b�`���O����
        .Fields("WFDSOD_MES").Value = ""            'WFDSOD_�v�����@
        .Fields("WFDSOD_TOTAL").Value = -1          'WFDSOD_���莞��TOTAL�l
    End With
                
    'TBCMX002
    With recX002
        .Fields("WFDSOD_SMPPOS").Value = -1         'WFDSOD�����-ID����ʒu(SXL�ʒu���)
        .Fields("WFDSOD_NETSU").Value = " "         'WFDSOD_�M��������
        .Fields("WFDSOD_ET").Value = " "            'WFDSOD_�G�b�`���O����
        .Fields("WFDSOD_MES").Value = " "           'WFDSOD_�v�����@
        .Fields("WFDSOD_DKAN").Value = " "          'WFDSOD_�c�j�A�j�[������
        .Fields("WFDSOD_MESDATA1").Value = " "      'WFDSOD����_1
        .Fields("WFDSOD_MESDATA2").Value = " "      'WFDSOD����_2
        .Fields("WFDSOD_MESDATA3").Value = " "      'WFDSOD����_3
        .Fields("WFDSOD_MESDATA4").Value = " "      'WFDSOD����_4
        .Fields("WFDSOD_MESDATA5").Value = " "      'WFDSOD����_5
        .Fields("WFDSOD_MESDATA6").Value = " "      'WFDSOD����_6
        .Fields("WFDSOD_MESDATA7").Value = " "      'WFDSOD����_7
        .Fields("WFDSOD_MESDATA8").Value = " "      'WFDSOD����_8
        .Fields("WFDSOD_MESDATA9").Value = " "      'WFDSOD����_9
        .Fields("WFDSOD_MESDATA10").Value = " "     'WFDSOD����_10
        .Fields("WFDSOD_MESDATA11").Value = " "     'WFDSOD����_11
        .Fields("WFDSOD_MESDATA12").Value = " "     'WFDSOD����_12
        .Fields("WFDSOD_MESDATA13").Value = " "     'WFDSOD����_13
        .Fields("WFDSOD_MESDATA14").Value = " "     'WFDSOD����_14
        .Fields("WFDSOD_MESDATA15").Value = " "     'WFDSOD����_15
    End With
    
    '-------------------- TBCMY013�̓ǂݍ���(WFDSOD) ----------------------------------------
    If (recXSDCW("WFINDDSCW").Value <> "0") And (recXSDCW("WFRESDSCW").Value <> "0") Then
        sql = "select * from TBCMY013 "
        sql = sql & "where SAMPLEID = '" & recXSDCW("WFSMPLIDDSCW").Value & "' and "
        sql = sql & "      SPEC = 'DSOD'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            .Fields("WFDSOD_SMPPOS").Value = recXSDCW("INPOSCW").Value          'WFDSOD�����-ID����ʒu(SXL�ʒu���)
            .Fields("WFDSOD_NETSU").Value = rs("NETSU")                         'WFDSOD_�M��������
            .Fields("WFDSOD_ET").Value = rs("ET")                               'WFDSOD_�G�b�`���O����
            .Fields("WFDSOD_MES").Value = rs("MES")                             'WFDSOD_�v�����@
            .Fields("WFDSOD_TOTAL").Value = rs("MESDATA1")                      'WFDSOD_���莞��TOTAL�l
        End With
            
        'TBCMX002
        With recX002
            .Fields("WFDSOD_SMPPOS").Value = recXSDCW("INPOSCW").Value          'WFDSOD�����-ID����ʒu(SXL�ʒu���)
            .Fields("WFDSOD_NETSU").Value = rs("NETSU")                         'WFDSOD_�M��������
            .Fields("WFDSOD_ET").Value = rs("ET")                               'WFDSOD_�G�b�`���O����
            .Fields("WFDSOD_MES").Value = rs("MES")                             'WFDSOD_�v�����@
            .Fields("WFDSOD_DKAN").Value = rs("DKAN")                           'WFDSOD_�c�j�A�j�[������
            .Fields("WFDSOD_MESDATA1").Value = rs("MESDATA1")                   'WFDSOD����_1
            .Fields("WFDSOD_MESDATA2").Value = rs("MESDATA2")                   'WFDSOD����_2
            .Fields("WFDSOD_MESDATA3").Value = rs("MESDATA3")                   'WFDSOD����_3
            .Fields("WFDSOD_MESDATA4").Value = rs("MESDATA4")                   'WFDSOD����_4
            .Fields("WFDSOD_MESDATA5").Value = rs("MESDATA5")                   'WFDSOD����_5
            .Fields("WFDSOD_MESDATA6").Value = rs("MESDATA6")                   'WFDSOD����_6
            .Fields("WFDSOD_MESDATA7").Value = rs("MESDATA7")                   'WFDSOD����_7
            .Fields("WFDSOD_MESDATA8").Value = rs("MESDATA8")                   'WFDSOD����_8
            .Fields("WFDSOD_MESDATA9").Value = rs("MESDATA9")                   'WFDSOD����_9
            .Fields("WFDSOD_MESDATA10").Value = rs("MESDATA10")                 'WFDSOD����_10
            .Fields("WFDSOD_MESDATA11").Value = rs("MESDATA11")                 'WFDSOD����_11
            .Fields("WFDSOD_MESDATA12").Value = rs("MESDATA12")                 'WFDSOD����_12
            .Fields("WFDSOD_MESDATA13").Value = rs("MESDATA13")                 'WFDSOD����_13
            .Fields("WFDSOD_MESDATA14").Value = rs("MESDATA14")                 'WFDSOD����_14
            .Fields("WFDSOD_MESDATA15").Value = rs("MESDATA15")                 'WFDSOD����_15
        End With
        Set rs = Nothing
    End If

    getTBCMY013WFDSOD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY013WFDSOD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :WFSPV����(TBCMY013)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW�\����   (�V����يǗ�(SXL))
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002�\����(����_�ް�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :WFSPV����(TBCMY013)�����ް����擾���ASXL�������E����_�ް��\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
Private Function getTBCMY013WFSPV(recXSDCW As c_cmzcrec, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY013WFSPV"
    
    getTBCMY013WFSPV = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("WFSPV_SMPPOS").Value = -1          'WFSPV�����-ID����ʒu(SXL�ʒu���)
        .Fields("WFSPV_NETSU").Value = ""           'WFSPV_�M��������
        .Fields("WFSPV_ET").Value = ""              'WFSPV_�G�b�`���O����
        .Fields("WFSPV_MES").Value = ""             'WFSPV_�v�����@
        .Fields("WFSPV_KST_MAX").Value = -1         'WFSPV_�g�U�����莞��MAX�l
        .Fields("WFSPV_KST_AVE").Value = -1         'WFSPV_�g�U�����莞��AVE�l
        .Fields("WFSPV_KST_MIN").Value = -1         'WFSPV_�g�U�����莞��MIN�l
        .Fields("WFSPV_FE_MAX").Value = -1          'WFSPV_Fe�Z�x���莞��MAX�l
        .Fields("WFSPV_FE_AVE").Value = -1          'WFSPV_Fe�Z�x���莞��AVE�l
        .Fields("WFSPV_FE_MIN").Value = -1          'WFSPV_Fe�Z�x���莞��MIN�l
    End With
                
    'TBCMX002
    With recX002
        .Fields("WFSPV_SMPPOS").Value = -1         'WFSPV�����-ID����ʒu(SXL�ʒu���)
        .Fields("WFSPV_NETSU").Value = " "         'WFSPV_�M��������
        .Fields("WFSPV_ET").Value = " "            'WFSPV_�G�b�`���O����
        .Fields("WFSPV_MES").Value = " "           'WFSPV_�v�����@
        .Fields("WFSPV_DKAN").Value = " "          'WFSPV_�c�j�A�j�[������
        .Fields("WFSPV_MESDATA1").Value = " "      'WFSPV����_1
        .Fields("WFSPV_MESDATA2").Value = " "      'WFSPV����_2
        .Fields("WFSPV_MESDATA3").Value = " "      'WFSPV����_3
        .Fields("WFSPV_MESDATA4").Value = " "      'WFSPV����_4
        .Fields("WFSPV_MESDATA5").Value = " "      'WFSPV����_5
        .Fields("WFSPV_MESDATA6").Value = " "      'WFSPV����_6
        .Fields("WFSPV_MESDATA7").Value = " "      'WFSPV����_7
        .Fields("WFSPV_MESDATA8").Value = " "      'WFSPV����_8
        .Fields("WFSPV_MESDATA9").Value = " "      'WFSPV����_9
        .Fields("WFSPV_MESDATA10").Value = " "     'WFSPV����_10
        .Fields("WFSPV_MESDATA11").Value = " "     'WFSPV����_11
        .Fields("WFSPV_MESDATA12").Value = " "     'WFSPV����_12
        .Fields("WFSPV_MESDATA13").Value = " "     'WFSPV����_13
        .Fields("WFSPV_MESDATA14").Value = " "     'WFSPV����_14
        .Fields("WFSPV_MESDATA15").Value = " "     'WFSPV����_15
    End With
    
    '-------------------- TBCMY013�̓ǂݍ���(WFSPV) ----------------------------------------
    If (recXSDCW("WFINDSPCW").Value <> "0") And (recXSDCW("WFRESSPCW").Value <> "0") Then
        sql = "select * from TBCMY013 "
        sql = sql & "where SAMPLEID = '" & recXSDCW("WFSMPLIDSPCW").Value & "' and "
        sql = sql & "      SPEC = 'SPV'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            .Fields("WFSPV_SMPPOS").Value = recXSDCW("INPOSCW").Value       'WFSPV�����-ID����ʒu(SXL�ʒu���)
            .Fields("WFSPV_NETSU").Value = rs("NETSU")                      'WFSPV_�M��������
            .Fields("WFSPV_ET").Value = rs("ET")                            'WFSPV_�G�b�`���O����
            .Fields("WFSPV_MES").Value = rs("MES")                          'WFSPV_�v�����@
            .Fields("WFSPV_KST_MAX").Value = rs("MESDATA1")                 'WFSPV_�g�U�����莞��MAX�l
            .Fields("WFSPV_KST_AVE").Value = rs("MESDATA2")                 'WFSPV_�g�U�����莞��AVE�l
            .Fields("WFSPV_KST_MIN").Value = rs("MESDATA3")                 'WFSPV_�g�U�����莞��MIN�l
            .Fields("WFSPV_FE_MAX").Value = rs("MESDATA4")                  'WFSPV_Fe�Z�x���莞��MAX�l
            .Fields("WFSPV_FE_AVE").Value = rs("MESDATA5")                  'WFSPV_Fe�Z�x���莞��AVE�l
            .Fields("WFSPV_FE_MIN").Value = rs("MESDATA6")                  'WFSPV_Fe�Z�x���莞��MIN�l
        End With
            
        'TBCMX002
        With recX002
            .Fields("WFSPV_SMPPOS").Value = recXSDCW("INPOSCW").Value       'WFSPV�����-ID����ʒu(SXL�ʒu���)
            .Fields("WFSPV_NETSU").Value = rs("NETSU")                      'WFSPV_�M��������
            .Fields("WFSPV_ET").Value = rs("ET")                            'WFSPV_�G�b�`���O����
            .Fields("WFSPV_MES").Value = rs("MES")                          'WFSPV_�v�����@
            .Fields("WFSPV_DKAN").Value = rs("DKAN")                        'WFSPV_�c�j�A�j�[������
            .Fields("WFSPV_MESDATA1").Value = rs("MESDATA1")                'WFSPV����_1
            .Fields("WFSPV_MESDATA2").Value = rs("MESDATA2")                'WFSPV����_2
            .Fields("WFSPV_MESDATA3").Value = rs("MESDATA3")                'WFSPV����_3
            .Fields("WFSPV_MESDATA4").Value = rs("MESDATA4")                'WFSPV����_4
            .Fields("WFSPV_MESDATA5").Value = rs("MESDATA5")                'WFSPV����_5
            .Fields("WFSPV_MESDATA6").Value = rs("MESDATA6")                'WFSPV����_6
            .Fields("WFSPV_MESDATA7").Value = rs("MESDATA7")                'WFSPV����_7
            .Fields("WFSPV_MESDATA8").Value = rs("MESDATA8")                'WFSPV����_8
            .Fields("WFSPV_MESDATA9").Value = rs("MESDATA9")                'WFSPV����_9
            .Fields("WFSPV_MESDATA10").Value = rs("MESDATA10")              'WFSPV����_10
            .Fields("WFSPV_MESDATA11").Value = rs("MESDATA11")              'WFSPV����_11
            .Fields("WFSPV_MESDATA12").Value = rs("MESDATA12")              'WFSPV����_12
            .Fields("WFSPV_MESDATA13").Value = rs("MESDATA13")              'WFSPV����_13
            .Fields("WFSPV_MESDATA14").Value = rs("MESDATA14")              'WFSPV����_14
            .Fields("WFSPV_MESDATA15").Value = rs("MESDATA15")              'WFSPV����_15
        End With
        Set rs = Nothing
    End If

    getTBCMY013WFSPV = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY013WFSPV = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :WFDZ����(TBCMY013)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW�\����   (�V����يǗ�(SXL))
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002�\����(����_�ް�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :WFDZ����(TBCMY013)�����ް����擾���ASXL�������E����_�ް��\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
Private Function getTBCMY013WFDZ(recXSDCW As c_cmzcrec, HIN As tFullHinban, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim DZ          As W_DZ                     'DZ�\����
    Dim JData(3)    As Double                   'MAX�l�Z�o�p�@06/09/06 ooba
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY013WFDZ"
    
    getTBCMY013WFDZ = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("WFDZ_SMPPOS").Value = -1           'WFDZ�����-ID����ʒu(SXL�ʒu���)
        .Fields("WFDZ_NETSU").Value = ""            'WFDZ_�M��������
        .Fields("WFDZ_ET").Value = ""               'WFDZ_�G�b�`���O����
        .Fields("WFDZ_MES").Value = ""              'WFDZ_�v�����@
        .Fields("WFDZ_MAX").Value = -1              'WFDZ_���莞��MAX�l
        .Fields("WFDZ_AVE").Value = -1              'WFDZ_���莞��AVE�l
        .Fields("WFDZ_MIN").Value = -1              'WFDZ_���莞��MIN�l
    End With

    'TBCMX002
    With recX002
        .Fields("WFDZ_SMPPOS").Value = -1           'WFDZ�����-ID����ʒu(SXL�ʒu���)
        .Fields("WFDZ_NETSU").Value = " "           'WFDZ_�M��������
        .Fields("WFDZ_ET").Value = " "              'WFDZ_�G�b�`���O����
        .Fields("WFDZ_MES").Value = " "             'WFDZ_�v�����@
        .Fields("WFDZ_DKAN").Value = " "            'WFDZ_�c�j�A�j�[������
        .Fields("WFDZ_MESDATA1").Value = " "        'WFDZ����_1
        .Fields("WFDZ_MESDATA2").Value = " "        'WFDZ����_2
        .Fields("WFDZ_MESDATA3").Value = " "        'WFDZ����_3
        .Fields("WFDZ_MESDATA4").Value = " "        'WFDZ����_4
        .Fields("WFDZ_MESDATA5").Value = " "        'WFDZ����_5
        .Fields("WFDZ_MESDATA6").Value = " "        'WFDZ����_6
        .Fields("WFDZ_MESDATA7").Value = " "        'WFDZ����_7
        .Fields("WFDZ_MESDATA8").Value = " "        'WFDZ����_8
        .Fields("WFDZ_MESDATA9").Value = " "        'WFDZ����_9
        .Fields("WFDZ_MESDATA10").Value = " "       'WFDZ����_10
        .Fields("WFDZ_MESDATA11").Value = " "       'WFDZ����_11
        .Fields("WFDZ_MESDATA12").Value = " "       'WFDZ����_12
        .Fields("WFDZ_MESDATA13").Value = " "       'WFDZ����_13
        .Fields("WFDZ_MESDATA14").Value = " "       'WFDZ����_14
        .Fields("WFDZ_MESDATA15").Value = " "       'WFDZ����_15
    End With
    
    '-------------------- TBCMY013�̓ǂݍ���(WFDZ) ----------------------------------------
    If (recXSDCW("WFINDDZCW").Value <> "0") And (recXSDCW("WFRESDZCW").Value <> "0") Then
        sql = "select * from TBCMY013 "
        sql = sql & "where SAMPLEID = '" & recXSDCW("WFSMPLIDDZCW").Value & "' and "
        sql = sql & "      SPEC = 'DZ'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            .Fields("WFDZ_SMPPOS").Value = recXSDCW("INPOSCW").Value       'WFDZ�����-ID����ʒu(SXL�ʒu���)
            .Fields("WFDZ_NETSU").Value = rs("NETSU")                      'WFDZ_�M��������
            .Fields("WFDZ_ET").Value = rs("ET")                            'WFDZ_�G�b�`���O����
            .Fields("WFDZ_MES").Value = rs("MES")                          'WFDZ_�v�����@
        End With
            
        'TBCMX002
        With recX002
            .Fields("WFDZ_SMPPOS").Value = recXSDCW("INPOSCW").Value       'WFDZ�����-ID����ʒu(SXL�ʒu���)
            .Fields("WFDZ_NETSU").Value = rs("NETSU")                      'WFDZ_�M��������
            .Fields("WFDZ_ET").Value = rs("ET")                            'WFDZ_�G�b�`���O����
            .Fields("WFDZ_MES").Value = rs("MES")                          'WFDZ_�v�����@
            .Fields("WFDZ_DKAN").Value = rs("DKAN")                        'WFDZ_�c�j�A�j�[������
            .Fields("WFDZ_MESDATA1").Value = rs("MESDATA1")                'WFDZ����_1
            .Fields("WFDZ_MESDATA2").Value = rs("MESDATA2")                'WFDZ����_2
            .Fields("WFDZ_MESDATA3").Value = rs("MESDATA3")                'WFDZ����_3
            .Fields("WFDZ_MESDATA4").Value = rs("MESDATA4")                'WFDZ����_4
            .Fields("WFDZ_MESDATA5").Value = rs("MESDATA5")                'WFDZ����_5
            .Fields("WFDZ_MESDATA6").Value = rs("MESDATA6")                'WFDZ����_6
            .Fields("WFDZ_MESDATA7").Value = rs("MESDATA7")                'WFDZ����_7
            .Fields("WFDZ_MESDATA8").Value = rs("MESDATA8")                'WFDZ����_8
            .Fields("WFDZ_MESDATA9").Value = rs("MESDATA9")                'WFDZ����_9
            .Fields("WFDZ_MESDATA10").Value = rs("MESDATA10")              'WFDZ����_10
            .Fields("WFDZ_MESDATA11").Value = rs("MESDATA11")              'WFDZ����_11
            .Fields("WFDZ_MESDATA12").Value = rs("MESDATA12")              'WFDZ����_12
            .Fields("WFDZ_MESDATA13").Value = rs("MESDATA13")              'WFDZ����_13
            .Fields("WFDZ_MESDATA14").Value = rs("MESDATA14")              'WFDZ����_14
            .Fields("WFDZ_MESDATA15").Value = rs("MESDATA15")              'WFDZ����_15
        End With
        Set rs = Nothing
                
        'WFDZ_MAX,MIN,AVE
        With recX002
            DZ.DZ(0) = NtoZ2(.Fields("WFDZ_MESDATA1").Value)               'DZ����l1
            DZ.DZ(1) = NtoZ2(.Fields("WFDZ_MESDATA2").Value)               'DZ����l2
            DZ.DZ(2) = NtoZ2(.Fields("WFDZ_MESDATA3").Value)               'DZ����l3
            DZ.DZ(3) = NtoZ2(.Fields("WFDZ_MESDATA4").Value)               'DZ����l4
        End With
        
        ''06/09/06 ooba START =============================================================>
        'DZ�d�l�擾
        sql = "select HWFMKSPH, HWFMKSPT, HWFMKSPR, HWFMKHWT, HWFMKHWS, HWFMKMIN, HWFMKMAX "
        sql = sql & "from TBCME024 where "
        sql = sql & "HINBAN = '" & HIN.hinban & "' and "
        sql = sql & "MNOREVNO = " & HIN.mnorevno & " and "
        sql = sql & "FACTORY = '" & HIN.factory & "' and "
        sql = sql & "OPECOND = '" & HIN.opecond & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        If IsNull(rs("HWFMKSPH")) = False Then DZ.GuaranteeDz.cMeth = rs("HWFMKSPH")    '�i�v�e�����בw����ʒu�Q��
        If IsNull(rs("HWFMKSPT")) = False Then DZ.GuaranteeDz.cCount = rs("HWFMKSPT")   '�i�v�e�����בw����ʒu�Q�_
        If IsNull(rs("HWFMKSPR")) = False Then DZ.GuaranteeDz.cPos = rs("HWFMKSPR")     '�i�v�e�����בw����ʒu�Q��
        If IsNull(rs("HWFMKHWT")) = False Then DZ.GuaranteeDz.cObj = rs("HWFMKHWT")     '�i�v�e�����בw�ۏؕ��@�Q��
        If IsNull(rs("HWFMKHWS")) = False Then DZ.GuaranteeDz.cJudg = rs("HWFMKHWS")    '�i�v�e�����בw�ۏؕ��@�Q��
        If IsNull(rs("HWFMKMIN")) = False Then DZ.SpecDzMin = rs("HWFMKMIN")            '�i�v�e�����בw����
        If IsNull(rs("HWFMKMAX")) = False Then DZ.SpecDzMax = rs("HWFMKMAX")            '�i�v�e�����בw���
        
        Set rs = Nothing
        
        '���躰�ނ� "F"�FMAX(2,4�_��)�C"G"�FMAX(2,3,4�_��) �̏ꍇ��MAX�l�ɂ��̒l���
        If DZ.GuaranteeDz.cJudg = JudgCodeW01 And _
            (DZ.GuaranteeDz.cObj = ObjCode10 Or DZ.GuaranteeDz.cObj = ObjCode11) Then
        
            If GetWfJudgData(WFDZ_JUDG, DZ.GuaranteeDz, DZ.DZ(), JData()) = FUNCTION_RETURN_FAILURE Then
                GoTo proc_exit
            End If
            recX001.Fields("WFDZ_MAX").Value = JData(0)
        Else
            recX001.Fields("WFDZ_MAX").Value = JudgMax(DZ.DZ())
        End If
        ''06/09/06 ooba END ===============================================================>
        
'        recX001.Fields("WFDZ_MAX").Value = JudgMax(DZ.DZ())             'WFDZ_���莞��MAX�l
        recX001.Fields("WFDZ_AVE").Value = JudgAve(DZ.DZ())             'WFDZ_���莞��AVE�l
        recX001.Fields("WFDZ_MIN").Value = JudgMin(DZ.DZ())             'WFDZ_���莞��MIN�l
    End If

    getTBCMY013WFDZ = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY013WFDZ = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :WFAOi����(TBCMY013)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW�\����   (�V����يǗ�(SXL))
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002�\����(����_�ް�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :WFAOi����(TBCMY013)�����ް����擾���ASXL�������E����_�ް��\���̂ɾ�Ă���
'����      :03/12/19 ooba
Private Function getTBCMY013WFAOi(recXSDCW As c_cmzcrec, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim DZ          As W_DZ                     'DZ�\����
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY013WFAOi"
    
    getTBCMY013WFAOi = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("WFDOI_NETU_3").Value = ""            'WFDOI_�M��������_�R
        .Fields("WFDOI_MES_3").Value = ""             'WFDOI_�v�����@_�R
        .Fields("WFDOI_MESDATA1_3").Value = -1        'WFDOI_(�Ƽ��Oi-AfterOi)�P_�R
        .Fields("WFDOI_MESDATA2_3").Value = -1        'WFDOI_(�Ƽ��Oi-AfterOi)�Q_�R
        .Fields("WFDOI_MESDATA3_3").Value = -1        'WFDOI_(�Ƽ��Oi-AfterOi)�R_�R
        .Fields("ZOIFLG").Value = ""                  '�c���_�f�����׸�
    End With

    'TBCMX002
    With recX002
        .Fields("WFDOI3_NETSU").Value = " "           'WFDOI-3_�M��������
        .Fields("WFDOI3_MES").Value = " "             'WFDOI-3_�v�����@
        .Fields("WFDOI3_MESDATA1").Value = " "        'WFDOI-3_����_1
        .Fields("WFDOI3_MESDATA2").Value = " "        'WFDOI-3_����_2
        .Fields("WFDOI3_MESDATA3").Value = " "        'WFDOI-3_����_3
        .Fields("WFDOI3_MESDATA4").Value = " "        'WFDOI-3_����_4
        .Fields("WFDOI3_MESDATA5").Value = " "        'WFDOI-3_����_5
        .Fields("WFDOI3_MESDATA6").Value = " "        'WFDOI-3_����_6
        .Fields("WFDOI3_MESDATA7").Value = " "        'WFDOI-3_����_7
        .Fields("WFDOI3_MESDATA8").Value = " "        'WFDOI-3_����_8
        .Fields("WFDOI3_MESDATA9").Value = " "        'WFDOI-3_����_9
        .Fields("WFDOI3_MESDATA10").Value = " "       'WFDOI-3_����_10
        .Fields("WFDOI3_MESDATA11").Value = " "       'WFDOI-3_����_11
        .Fields("WFDOI3_MESDATA12").Value = " "       'WFDOI-3_����_12
        .Fields("WFDOI3_MESDATA13").Value = " "       'WFDOI-3_����_13
        .Fields("WFDOI3_MESDATA14").Value = " "       'WFDOI-3_����_14
        .Fields("WFDOI3_MESDATA15").Value = " "       'WFDOI-3_����_15
        .Fields("ZOIFLG").Value = " "                 '�c���_�f�����׸�
    End With
    
    '-------------------- TBCMY013�̓ǂݍ���(WFAOi) ----------------------------------------
    If (recXSDCW("WFINDAOICW").Value <> "0") And (recXSDCW("WFRESAOICW").Value <> "0") Then
        sql = "select * from TBCMY013 "
        sql = sql & "where SAMPLEID = '" & recXSDCW("WFSMPLIDAOICW").Value & "' and "
        sql = sql & "      SPEC = 'AOI'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            .Fields("WFDOI_NETU_3").Value = rs("NETSU")                'WFDOI_�M��������_�R
            .Fields("WFDOI_MES_3").Value = rs("MES")                   'WFDOI_�v�����@_�R
            .Fields("WFDOI_MESDATA1_3").Value = rs("MESDATA4")         'WFDOI_(�Ƽ��Oi-AfterOi)�P_�R
            .Fields("WFDOI_MESDATA2_3").Value = rs("MESDATA5")         'WFDOI_(�Ƽ��Oi-AfterOi)�Q_�R
            .Fields("WFDOI_MESDATA3_3").Value = rs("MESDATA6")         'WFDOI_(�Ƽ��Oi-AfterOi)�R_�R
            .Fields("ZOIFLG").Value = "1"                              '�c���_�f�����׸�
        End With
            
        'TBCMX002
        With recX002
            .Fields("WFDOI3_NETSU").Value = rs("NETSU")                      'WFDOI-3_�M��������
            .Fields("WFDOI3_MES").Value = rs("MES")                          'WFDOI-3_�v�����@
            .Fields("WFDOI3_MESDATA1").Value = rs("MESDATA1")                'WFDOI-3_����_1
            .Fields("WFDOI3_MESDATA2").Value = rs("MESDATA2")                'WFDOI-3_����_2
            .Fields("WFDOI3_MESDATA3").Value = rs("MESDATA3")                'WFDOI-3_����_3
            .Fields("WFDOI3_MESDATA4").Value = rs("MESDATA4")                'WFDOI-3_����_4
            .Fields("WFDOI3_MESDATA5").Value = rs("MESDATA5")                'WFDOI-3_����_5
            .Fields("WFDOI3_MESDATA6").Value = rs("MESDATA6")                'WFDOI-3_����_6
            .Fields("WFDOI3_MESDATA7").Value = rs("MESDATA7")                'WFDOI-3_����_7
            .Fields("WFDOI3_MESDATA8").Value = rs("MESDATA8")                'WFDOI-3_����_8
            .Fields("WFDOI3_MESDATA9").Value = rs("MESDATA9")                'WFDOI-3_����_9
            .Fields("WFDOI3_MESDATA10").Value = rs("MESDATA10")              'WFDOI-3_����_10
            .Fields("WFDOI3_MESDATA11").Value = rs("MESDATA11")              'WFDOI-3_����_11
            .Fields("WFDOI3_MESDATA12").Value = rs("MESDATA12")              'WFDOI-3_����_12
            .Fields("WFDOI3_MESDATA13").Value = rs("MESDATA13")              'WFDOI-3_����_13
            .Fields("WFDOI3_MESDATA14").Value = rs("MESDATA14")              'WFDOI-3_����_14
            .Fields("WFDOI3_MESDATA15").Value = rs("MESDATA15")              'WFDOI-3_����_15
            .Fields("ZOIFLG").Value = "1"                                    '�c���_�f�����׸�
        End With
        Set rs = Nothing
    End If

    getTBCMY013WFAOi = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY013WFAOi = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :���OSF1�`3����(TBCMY022)�ް��擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :recXSDCW        , I  ,c_cmzcrec         , �V����يǗ�(SXL)
'          :j               , I  ,Integer           , ���OSF No
'          :HIN             , I  ,tFullHinban       , �i��
'�@�@      :sPos  �@�@�@    , I  ,String �@         , SXL�ʒu(TOP/BOT)
'          :recX004         , O  ,c_cmzcrec         , EP������
'          :recX005         , O  ,c_cmzcrec         , EP����_�ް�
'�@�@      :sTblName�@�@    , I  ,String �@         , �e�[�u�����@11/06/24 Marushita  MIN�l�ǉ��Ή�
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :EP��s�]������(TBCMY022)��������ް����擾���AEP������/EP����_�ް��\���̂ɾ�Ă���
'����      :06/08/10 ooba
'Private Function getTBCMY022EPOSF(recXSDCW As c_cmzcrec, j As Integer, HIN As tFullHinban, sPos As String, recX004 As c_cmzcrec, recX005 As c_cmzcrec) As FUNCTION_RETURN
Private Function getTBCMY022EPOSF(recXSDCW As c_cmzcrec, j As Integer, _
                                  HIN As tFullHinban, sPos As String, _
                                  recX004 As c_cmzcrec, recX005 As c_cmzcrec, sTblName As String) As FUNCTION_RETURN
                                  
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim eosf        As W_OSF                    '���OSF�\����
    Dim keisu       As Double
    Dim k           As Integer
    Dim HEPOSFKN    As String                   '�����p�x_��
    
    Const keisu1 As Double = 1.8248175
    Const keisu2 As Double = 1.8518519
    Const keisu3 As Double = 1.9230769
    Const keisu4 As Double = 3.649635
    Const keisu5 As Double = 3.7037037
    Const keisu6 As Double = 3.8461538
    Const keisu7 As Double = 7.6923077
    
    '�װ����ׂ̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY022EPOSF"
    
    getTBCMY022EPOSF = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    'TBCMX004
    With recX004
        .Fields("EPOSF" & j & "_SMPPOS").Value = vbNullString       'EPOSFx�����-ID����ʒu(SXL�ʒu���)
        .Fields("EPOSF" & j & "_NETSU").Value = vbNullString        'EPOSFx_�M��������
        .Fields("EPOSF" & j & "_ET").Value = vbNullString           'EPOSFx_���ݸޏ���
        .Fields("EPOSF" & j & "_MES").Value = vbNullString          'EPOSFx_�v�����@
        .Fields("EPOSF" & j & "_MAX").Value = vbNullString          'EPOSFx_���莞��MAX�l_x
        .Fields("EPOSF" & j & "_AVE").Value = vbNullString          'EPOSFx_���莞��AVE�l_x
    End With
                
    'TBCMX005
    With recX005
        .Fields("EPOSF" & j & "_SMPPOS").Value = vbNullString       'EPOSFx�����-ID����ʒu(SXL�ʒu���)
        .Fields("EPOSF" & j & "_NETSU").Value = vbNullString        'EPOSFx_�M��������
        .Fields("EPOSF" & j & "_ET").Value = vbNullString           'EPOSFx_���ݸޏ���
        .Fields("EPOSF" & j & "_MES").Value = vbNullString          'EPOSFx_�v�����@
        .Fields("EPOSF" & j & "_DKAN").Value = vbNullString         'EPOSFx_DK�ưُ���
        .Fields("EPOSF" & j & "_MESDATA1").Value = vbNullString     'EPOSFx����_1
        .Fields("EPOSF" & j & "_MESDATA2").Value = vbNullString     'EPOSFx����_2
        .Fields("EPOSF" & j & "_MESDATA3").Value = vbNullString     'EPOSFx����_3
        .Fields("EPOSF" & j & "_MESDATA4").Value = vbNullString     'EPOSFx����_4
        .Fields("EPOSF" & j & "_MESDATA5").Value = vbNullString     'EPOSFx����_5
        .Fields("EPOSF" & j & "_MESDATA6").Value = vbNullString     'EPOSFx����_6
        .Fields("EPOSF" & j & "_MESDATA7").Value = vbNullString     'EPOSFx����_7
        .Fields("EPOSF" & j & "_MESDATA8").Value = vbNullString     'EPOSFx����_8
        .Fields("EPOSF" & j & "_MESDATA9").Value = vbNullString     'EPOSFx����_9
        .Fields("EPOSF" & j & "_MESDATA10").Value = vbNullString    'EPOSFx����_10
        .Fields("EPOSF" & j & "_MESDATA11").Value = vbNullString    'EPOSFx����_11
        .Fields("EPOSF" & j & "_MESDATA12").Value = vbNullString    'EPOSFx����_12
        .Fields("EPOSF" & j & "_MESDATA13").Value = vbNullString    'EPOSFx����_13
        .Fields("EPOSF" & j & "_MESDATA14").Value = vbNullString    'EPOSFx����_14
        .Fields("EPOSF" & j & "_MESDATA15").Value = vbNullString    'EPOSFx����_15
    End With
    
    '-------------------- TBCMY022�̓ǂݍ���(EPOSF) ----------------------------------------
    If (recXSDCW("EPINDL" & j & "CW").Value <> "0") And _
        (recXSDCW("EPRESL" & j & "CW").Value <> "0") Then
        
        sql = "select "
        sql = sql & "SAMPLEID "
        sql = sql & ",OSITEM "
        sql = sql & ",MAISU "
        sql = sql & ",SPEC "
        sql = sql & ",NETSU "
        sql = sql & ",ET "
        sql = sql & ",MES "
        sql = sql & ",DKAN "
        sql = sql & ",MESDATA1 "
        sql = sql & ",MESDATA2 "
        sql = sql & ",MESDATA3 "
        sql = sql & ",MESDATA4 "
        sql = sql & ",MESDATA5 "
        sql = sql & ",MESDATA6 "
        sql = sql & ",MESDATA7 "
        sql = sql & ",MESDATA8 "
        sql = sql & ",MESDATA9 "
        sql = sql & ",MESDATA10 "
        sql = sql & ",MESDATA11 "
        sql = sql & ",MESDATA12 "
        sql = sql & ",MESDATA13 "
        sql = sql & ",MESDATA14 "
        sql = sql & ",MESDATA15 "
        sql = sql & ",TXID "
        sql = sql & ",REGDATE "
        sql = sql & ",SENDFLAG "
        sql = sql & ",SENDDATE "
        sql = sql & "from TBCMY022 "
        sql = sql & "where SAMPLEID = '" & recXSDCW("EPSMPLIDL" & j & "CW").Value & "' and "
        sql = sql & "      SPEC = 'OSF" & j & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX004
        With recX004
            .Fields("EPOSF" & j & "_SMPPOS").Value = recXSDCW("INPOSCW").Value  'EPOSFx�����-ID����ʒu(SXL�ʒu���)
            .Fields("EPOSF" & j & "_NETSU").Value = rs("NETSU")                 'EPOSFx_�M��������
            .Fields("EPOSF" & j & "_ET").Value = rs("ET")                       'EPOSFx_���ݸޏ���
            .Fields("EPOSF" & j & "_MES").Value = rs("MES")                     'EPOSFx_�v�����@
        End With
            
        'TBCMX005
        With recX005
            .Fields("EPOSF" & j & "_SMPPOS").Value = recXSDCW("INPOSCW").Value  'EPOSFx�����-ID����ʒu(SXL�ʒu���)
            .Fields("EPOSF" & j & "_NETSU").Value = rs("NETSU")                 'EPOSFx_�M��������
            .Fields("EPOSF" & j & "_ET").Value = rs("ET")                       'EPOSFx_���ݸޏ���
            .Fields("EPOSF" & j & "_MES").Value = rs("MES")                     'EPOSFx_�v�����@
            .Fields("EPOSF" & j & "_DKAN").Value = rs("DKAN")                   'EPOSFx_DK�ưُ���
            .Fields("EPOSF" & j & "_MESDATA1").Value = rs("MESDATA1")           'EPOSFx����_1
            .Fields("EPOSF" & j & "_MESDATA2").Value = rs("MESDATA2")           'EPOSFx����_2
            .Fields("EPOSF" & j & "_MESDATA3").Value = rs("MESDATA3")           'EPOSFx����_3
            .Fields("EPOSF" & j & "_MESDATA4").Value = rs("MESDATA4")           'EPOSFx����_4
            .Fields("EPOSF" & j & "_MESDATA5").Value = rs("MESDATA5")           'EPOSFx����_5
            .Fields("EPOSF" & j & "_MESDATA6").Value = rs("MESDATA6")           'EPOSFx����_6
            .Fields("EPOSF" & j & "_MESDATA7").Value = rs("MESDATA7")           'EPOSFx����_7
            .Fields("EPOSF" & j & "_MESDATA8").Value = rs("MESDATA8")           'EPOSFx����_8
            .Fields("EPOSF" & j & "_MESDATA9").Value = rs("MESDATA9")           'EPOSFx����_9
            .Fields("EPOSF" & j & "_MESDATA10").Value = rs("MESDATA10")         'EPOSFx����_10
            .Fields("EPOSF" & j & "_MESDATA11").Value = rs("MESDATA11")         'EPOSFx����_11
            .Fields("EPOSF" & j & "_MESDATA12").Value = rs("MESDATA12")         'EPOSFx����_12
            .Fields("EPOSF" & j & "_MESDATA13").Value = rs("MESDATA13")         'EPOSFx����_13
            .Fields("EPOSF" & j & "_MESDATA14").Value = rs("MESDATA14")         'EPOSFx����_14
            .Fields("EPOSF" & j & "_MESDATA15").Value = rs("MESDATA15")         'EPOSFx����_15
        End With
        Set rs = Nothing
        
        'EPOSF_MAX,AVE
        sql = "select "
        sql = sql & "HEPOF" & j & "SH "
        sql = sql & ",HEPOF" & j & "ST "
        sql = sql & ",HEPOF" & j & "SR "
        sql = sql & ",HEPOF" & j & "HT "
        sql = sql & ",HEPOF" & j & "HS "
        sql = sql & ",HEPOF" & j & "KN "
        sql = sql & ",HEPOF" & j & "AX "
        sql = sql & ",HEPOF" & j & "MX "
        sql = sql & ",HEPOSF" & j & "PTK "
        sql = sql & "from TBCME050 where "
        sql = sql & "HINBAN = '" & HIN.hinban & "' and "
        sql = sql & "MNOREVNO = " & HIN.mnorevno & " and "
        sql = sql & "FACTORY = '" & HIN.factory & "' and "
        sql = sql & "OPECOND = '" & HIN.opecond & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        If IsNull(rs("HEPOF" & j & "SH")) = False Then eosf.GuaranteeOsf.cMeth = rs("HEPOF" & j & "SH")     '�iEPOSFx����ʒu_��
        If IsNull(rs("HEPOF" & j & "ST")) = False Then eosf.GuaranteeOsf.cCount = rs("HEPOF" & j & "ST")    '�iEPOSFx����ʒu_�_
        If IsNull(rs("HEPOF" & j & "SR")) = False Then eosf.GuaranteeOsf.cPos = rs("HEPOF" & j & "SR")      '�iEPOSFx����ʒu_��
        If IsNull(rs("HEPOF" & j & "HT")) = False Then eosf.GuaranteeOsf.cObj = rs("HEPOF" & j & "HT")      '�iEPOSFx�ۏؕ��@_��
        If IsNull(rs("HEPOF" & j & "HS")) = False Then eosf.GuaranteeOsf.cJudg = rs("HEPOF" & j & "HS")     '�iEPOSFx�ۏؕ��@_��
        If IsNull(rs("HEPOF" & j & "KN")) = False Then HEPOSFKN = rs("HEPOF" & j & "KN")                    '�iEPOSFx�����p�x_��
        If IsNull(rs("HEPOF" & j & "AX")) = False Then eosf.SpecOsfAveMax = rs("HEPOF" & j & "AX")          '�iEPOSFx���Ϗ��
        If IsNull(rs("HEPOF" & j & "MX")) = False Then eosf.SpecOsfMax = rs("HEPOF" & j & "MX")             '�iEPOSFx���
        If IsNull(rs("HEPOSF" & j & "PTK")) = False Then eosf.JudgDataPTK = rs("HEPOSF" & j & "PTK")        '�iEPOSFx���݋敪
        Set rs = Nothing
            
        If eosf.GuaranteeOsf.cMeth = "5" And eosf.GuaranteeOsf.cCount = "5" And eosf.GuaranteeOsf.cPos = "3" Then
            keisu = keisu1
        ElseIf eosf.GuaranteeOsf.cMeth = "5" And eosf.GuaranteeOsf.cCount = "5" And eosf.GuaranteeOsf.cPos = "5" Then
            keisu = keisu2
        ElseIf eosf.GuaranteeOsf.cMeth = "5" And eosf.GuaranteeOsf.cCount = "5" And eosf.GuaranteeOsf.cPos = "A" Then
            keisu = keisu3
        ElseIf eosf.GuaranteeOsf.cMeth = "6" And eosf.GuaranteeOsf.cCount = "5" And eosf.GuaranteeOsf.cPos = "3" Then
            keisu = keisu4
        ElseIf eosf.GuaranteeOsf.cMeth = "6" And eosf.GuaranteeOsf.cCount = "5" And eosf.GuaranteeOsf.cPos = "5" Then
            keisu = keisu5
        ElseIf eosf.GuaranteeOsf.cMeth = "6" And eosf.GuaranteeOsf.cCount = "5" And eosf.GuaranteeOsf.cPos = "A" Then
            keisu = keisu6
        ElseIf eosf.GuaranteeOsf.cMeth = "E" And eosf.GuaranteeOsf.cCount = "5" And eosf.GuaranteeOsf.cPos = "A" Then
            keisu = keisu7
        Else
            keisu = -1
        End If
            
        If keisu <> -1 Then
            With recX005
                If IsNull(.Fields("EPOSF" & j & "_MESDATA1").Value) = False Then
                    eosf.OSF(0) = NtoZ2(.Fields("EPOSF" & j & "_MESDATA1").Value)   'OSF����l1
                Else
                    eosf.OSF(0) = -1
                End If
                If IsNull(.Fields("EPOSF" & j & "_MESDATA2").Value) = False Then
                    eosf.OSF(1) = NtoZ2(.Fields("EPOSF" & j & "_MESDATA2").Value)   'OSF����l2
                Else
                    eosf.OSF(1) = -1
                End If
                If IsNull(.Fields("EPOSF" & j & "_MESDATA3").Value) = False Then
                    eosf.OSF(2) = NtoZ2(.Fields("EPOSF" & j & "_MESDATA3").Value)   'OSF����l3
                Else
                    eosf.OSF(2) = -1
                End If
                If IsNull(.Fields("EPOSF" & j & "_MESDATA4").Value) = False Then
                    eosf.OSF(3) = NtoZ2(.Fields("EPOSF" & j & "_MESDATA4").Value)   'OSF����l4
                Else
                    eosf.OSF(3) = -1
                End If
                If IsNull(.Fields("EPOSF" & j & "_MESDATA5").Value) = False Then
                    eosf.OSF(4) = NtoZ2(.Fields("EPOSF" & j & "_MESDATA5").Value)   'OSF����l5
                Else
                    eosf.OSF(4) = -1
                End If
                For k = 0 To 4
                    eosf.OSF(k) = IIf(eosf.OSF(k) <> -1, eosf.OSF(k) * keisu, -1)
                Next
            End With
            
            recX004.Fields("EPOSF" & j & "_MAX").Value = JudgMax(eosf.OSF())        'EPOSFx_���莞��MAX�l_x
            recX004.Fields("EPOSF" & j & "_AVE").Value = JudgAve(eosf.OSF())        'EPOSFx_���莞��AVE�l_x
            '>>>>> 2011/06/24 SETsw)Marushita MIN�l�Z�b�g�ǉ��Ή�
            If FieldCheck(sTblName, "EPOSF" & j & "_MIN") = FUNCTION_RETURN_SUCCESS Then
                recX004.Fields("EPOSF" & j & "_MIN").Value = JudgMin(eosf.OSF())        'EPOSFx_���莞��MIN�l_x
            End If
            '<<<<< 2011/06/24 SETsw)Marushita MIN�l�Z�b�g�ǉ��Ή�
        End If
            
        '�ۏؕ��@="H"����MAX�l/AVE�l��-1�̏ꍇ�װ�Ƃ���
        If ((eosf.GuaranteeOsf.cJudg = "H") And CheckKHN_EP(HEPOSFKN, j + 3, sPos)) And _
           (recX004.Fields("EPOSF" & j & "_MAX").Value = -1 Or _
            recX004.Fields("EPOSF" & j & "_AVE").Value = -1) Then GoTo proc_exit
        '�H�H�H�H�HMIN�l���`�F�b�N�H�H�H�H�H
    End If

    getTBCMY022EPOSF = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�װ�����
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY022EPOSF = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :���BMD1�`3����(TBCMY022)�ް��擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :recXSDCW        , I  ,c_cmzcrec         , �V����يǗ�(SXL)
'          :j               , I  ,Integer           , ���BMD No
'          :HIN             , I  ,tFullHinban       , �i��
'�@�@      :sPos  �@�@�@    , I  ,String �@         , SXL�ʒu(TOP/BOT)
'          :recX004         , O  ,c_cmzcrec         , EP������
'          :recX005         , O  ,c_cmzcrec         , EP����_�ް�
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :EP��s�]������(TBCMY022)��������ް����擾���AEP������/EP����_�ް��\���̂ɾ�Ă���
'����      :06/08/10 ooba
Private Function getTBCMY022EPBMD(recXSDCW As c_cmzcrec, j As Integer, _
                                  HIN As tFullHinban, sPos As String, _
                                  recX004 As c_cmzcrec, recX005 As c_cmzcrec) As FUNCTION_RETURN
                                  
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim ebmd        As W_BMD                    '���BMD�\����
    Dim k           As Integer
    Dim HEPBMDKN    As String                   '�����p�x_��
    Dim JData(4)    As Double
    
    Dim keisu As Double
    Const keisu1 As Double = 10000
    Const keisu2 As Double = 10000
    Const keisu3 As Double = 10000
    Const keisu4 As Double = 10000
    Const keisu5 As Double = 10000
    Const keisu6 As Double = 333000
    Const keisu7 As Double = 10000
    Const keisu8 As Double = 10000

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY022EPBMD"
    
    getTBCMY022EPBMD = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    'TBCMX004
    With recX004
        .Fields("EPBMD" & j & "_SMPPOS").Value = vbNullString       'EPBMDx�����-ID����ʒu(SXL�ʒu���)
        .Fields("EPBMD" & j & "_NETSU").Value = vbNullString        'EPBMDx_�M��������
        .Fields("EPBMD" & j & "_ET").Value = vbNullString           'EPBMDx_���ݸޏ���
        .Fields("EPBMD" & j & "_MES").Value = vbNullString          'EPBMDx_�v�����@
        .Fields("EPBMD" & j & "_MAX").Value = vbNullString          'EPBMDx_���莞��MAX�l_x
        .Fields("EPBMD" & j & "_AVE").Value = vbNullString          'EPBMDx_���莞��AVE�l_x
        .Fields("EPBMD" & j & "_MIN").Value = vbNullString          'EPBMDx_���莞��MIN�l_x
        .Fields("EPBMD" & j & "_MBP").Value = vbNullString          'EPBMDx_���莞�̖ʓ����z
    End With
                
    'TBCMX005
    With recX005
        .Fields("EPBMD" & j & "_SMPPOS").Value = vbNullString       'EPBMDx�����-ID����ʒu(SXL�ʒu���)
        .Fields("EPBMD" & j & "_NETSU").Value = vbNullString        'EPBMDx_�M��������
        .Fields("EPBMD" & j & "_ET").Value = vbNullString           'EPBMDx_���ݸޏ���
        .Fields("EPBMD" & j & "_MES").Value = vbNullString          'EPBMDx_�v�����@
        .Fields("EPBMD" & j & "_DKAN").Value = vbNullString         'EPBMDx_DK�ưُ���
        .Fields("EPBMD" & j & "_MESDATA1").Value = vbNullString     'EPBMDx����_1
        .Fields("EPBMD" & j & "_MESDATA2").Value = vbNullString     'EPBMDx����_2
        .Fields("EPBMD" & j & "_MESDATA3").Value = vbNullString     'EPBMDx����_3
        .Fields("EPBMD" & j & "_MESDATA4").Value = vbNullString     'EPBMDx����_4
        .Fields("EPBMD" & j & "_MESDATA5").Value = vbNullString     'EPBMDx����_5
        .Fields("EPBMD" & j & "_MESDATA6").Value = vbNullString     'EPBMDx����_6
        .Fields("EPBMD" & j & "_MESDATA7").Value = vbNullString     'EPBMDx����_7
        .Fields("EPBMD" & j & "_MESDATA8").Value = vbNullString     'EPBMDx����_8
        .Fields("EPBMD" & j & "_MESDATA9").Value = vbNullString     'EPBMDx����_9
        .Fields("EPBMD" & j & "_MESDATA10").Value = vbNullString    'EPBMDx����_10
        .Fields("EPBMD" & j & "_MESDATA11").Value = vbNullString    'EPBMDx����_11
        .Fields("EPBMD" & j & "_MESDATA12").Value = vbNullString    'EPBMDx����_12
        .Fields("EPBMD" & j & "_MESDATA13").Value = vbNullString    'EPBMDx����_13
        .Fields("EPBMD" & j & "_MESDATA14").Value = vbNullString    'EPBMDx����_14
        .Fields("EPBMD" & j & "_MESDATA15").Value = vbNullString    'EPBMDx����_15
    End With
    
    '-------------------- TBCMY022�̓ǂݍ���(EPBMD) ----------------------------------------
    If (recXSDCW("EPINDB" & j & "CW").Value <> "0") And _
    (recXSDCW("EPRESB" & j & "CW").Value <> "0") Then
    
        sql = "select "
        sql = sql & "SAMPLEID "
        sql = sql & ",OSITEM "
        sql = sql & ",MAISU "
        sql = sql & ",SPEC "
        sql = sql & ",NETSU "
        sql = sql & ",ET "
        sql = sql & ",MES "
        sql = sql & ",DKAN "
        sql = sql & ",MESDATA1 "
        sql = sql & ",MESDATA2 "
        sql = sql & ",MESDATA3 "
        sql = sql & ",MESDATA4 "
        sql = sql & ",MESDATA5 "
        sql = sql & ",MESDATA6 "
        sql = sql & ",MESDATA7 "
        sql = sql & ",MESDATA8 "
        sql = sql & ",MESDATA9 "
        sql = sql & ",MESDATA10 "
        sql = sql & ",MESDATA11 "
        sql = sql & ",MESDATA12 "
        sql = sql & ",MESDATA13 "
        sql = sql & ",MESDATA14 "
        sql = sql & ",MESDATA15 "
        sql = sql & ",TXID "
        sql = sql & ",REGDATE "
        sql = sql & ",SENDFLAG "
        sql = sql & ",SENDDATE "
        sql = sql & "from TBCMY022 "
        sql = sql & "where SAMPLEID = '" & recXSDCW("EPSMPLIDB" & j & "CW").Value & "' and "
        sql = sql & "      SPEC = 'BMD" & j & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX004
        With recX004
            .Fields("EPBMD" & j & "_SMPPOS").Value = recXSDCW("INPOSCW").Value  'EPBMDx�����-ID����ʒu(SXL�ʒu���)
            .Fields("EPBMD" & j & "_NETSU").Value = rs("NETSU")                 'EPBMDx_�M��������
            .Fields("EPBMD" & j & "_ET").Value = rs("ET")                       'EPBMDx_���ݸޏ���
            .Fields("EPBMD" & j & "_MES").Value = rs("MES")                     'EPBMDx_�v�����@
        End With
            
        'TBCMX005
        With recX005
            .Fields("EPBMD" & j & "_SMPPOS").Value = recXSDCW("INPOSCW").Value  'EPBMDx�����-ID����ʒu(SXL�ʒu���)
            .Fields("EPBMD" & j & "_NETSU").Value = rs("NETSU")                 'EPBMDx_�M��������
            .Fields("EPBMD" & j & "_ET").Value = rs("ET")                       'EPBMDx_���ݸޏ���
            .Fields("EPBMD" & j & "_MES").Value = rs("MES")                     'EPBMDx_�v�����@
            .Fields("EPBMD" & j & "_DKAN").Value = rs("DKAN")                   'EPBMDx_DK�ưُ���
            .Fields("EPBMD" & j & "_MESDATA1").Value = rs("MESDATA1")           'EPBMDx����_1
            .Fields("EPBMD" & j & "_MESDATA2").Value = rs("MESDATA2")           'EPBMDx����_2
            .Fields("EPBMD" & j & "_MESDATA3").Value = rs("MESDATA3")           'EPBMDx����_3
            .Fields("EPBMD" & j & "_MESDATA4").Value = rs("MESDATA4")           'EPBMDx����_4
            .Fields("EPBMD" & j & "_MESDATA5").Value = rs("MESDATA5")           'EPBMDx����_5
            .Fields("EPBMD" & j & "_MESDATA6").Value = rs("MESDATA6")           'EPBMDx����_6
            .Fields("EPBMD" & j & "_MESDATA7").Value = rs("MESDATA7")           'EPBMDx����_7
            .Fields("EPBMD" & j & "_MESDATA8").Value = rs("MESDATA8")           'EPBMDx����_8
            .Fields("EPBMD" & j & "_MESDATA9").Value = rs("MESDATA9")           'EPBMDx����_9
            .Fields("EPBMD" & j & "_MESDATA10").Value = rs("MESDATA10")         'EPBMDx����_10
            .Fields("EPBMD" & j & "_MESDATA11").Value = rs("MESDATA11")         'EPBMDx����_11
            .Fields("EPBMD" & j & "_MESDATA12").Value = rs("MESDATA12")         'EPBMDx����_12
            .Fields("EPBMD" & j & "_MESDATA13").Value = rs("MESDATA13")         'EPBMDx����_13
            .Fields("EPBMD" & j & "_MESDATA14").Value = rs("MESDATA14")         'EPBMDx����_14
            .Fields("EPBMD" & j & "_MESDATA15").Value = rs("MESDATA15")         'EPBMDx����_15
        End With
        Set rs = Nothing
                    
        'EPBMD_MAX,MIN,AVE,MBP
        sql = "select "
        sql = sql & "HEPBM" & j & "SH "
        sql = sql & ",HEPBM" & j & "ST "
        sql = sql & ",HEPBM" & j & "SR "
        sql = sql & ",HEPBM" & j & "HT "
        sql = sql & ",HEPBM" & j & "HS "
        sql = sql & ",HEPBM" & j & "KN "
        sql = sql & ",HEPBM" & j & "AN "
        sql = sql & ",HEPBM" & j & "AX "
        sql = sql & ",HEPBM" & j & "MBP "
        sql = sql & ",HEPBM" & j & "MCL "
        sql = sql & "from TBCME050 where "
        sql = sql & "HINBAN = '" & HIN.hinban & "' and "
        sql = sql & "MNOREVNO = " & HIN.mnorevno & " and "
        sql = sql & "FACTORY = '" & HIN.factory & "' and "
        sql = sql & "OPECOND = '" & HIN.opecond & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
                    
        If IsNull(rs("HEPBM" & j & "SH")) = False Then ebmd.GuaranteeBmd.cMeth = rs("HEPBM" & j & "SH")     '�iEPBMDx����ʒu_��
        If IsNull(rs("HEPBM" & j & "ST")) = False Then ebmd.GuaranteeBmd.cCount = rs("HEPBM" & j & "ST")    '�iEPBMDx����ʒu_�_
        If IsNull(rs("HEPBM" & j & "SR")) = False Then ebmd.GuaranteeBmd.cPos = rs("HEPBM" & j & "SR")      '�iEPBMDx����ʒu_��
        If IsNull(rs("HEPBM" & j & "HT")) = False Then ebmd.GuaranteeBmd.cObj = rs("HEPBM" & j & "HT")      '�iEPBMDx�ۏؕ��@_��
        If IsNull(rs("HEPBM" & j & "HS")) = False Then ebmd.GuaranteeBmd.cJudg = rs("HEPBM" & j & "HS")     '�iEPBMDx�ۏؕ��@_��
        If IsNull(rs("HEPBM" & j & "KN")) = False Then HEPBMDKN = rs("HEPBM" & j & "KN")                    '�iEPBMDx�����p�x_��
        If IsNull(rs("HEPBM" & j & "AN")) = False Then ebmd.SpecBmdAveMin = rs("HEPBM" & j & "AN")          '�iEPBMDx���ω���
        If IsNull(rs("HEPBM" & j & "AX")) = False Then ebmd.SpecBmdAveMax = rs("HEPBM" & j & "AX")          '�iEPBMDx���Ϗ��
        If IsNull(rs("HEPBM" & j & "MBP")) = False Then ebmd.SpecBmdMBP = rs("HEPBM" & j & "MBP")           '�iEPBMDx�ʓ����z
        If IsNull(rs("HEPBM" & j & "MCL")) = False Then ebmd.SpecBmdMCL = rs("HEPBM" & j & "MCL")           '�iEPBMDx�ʓ��v�Z
        Set rs = Nothing

        If ebmd.GuaranteeBmd.cMeth = "2" And ebmd.GuaranteeBmd.cCount = "4" And ebmd.GuaranteeBmd.cPos = "H" Then
            keisu = keisu1
        ElseIf ebmd.GuaranteeBmd.cMeth = "2" And ebmd.GuaranteeBmd.cCount = "3" And ebmd.GuaranteeBmd.cPos = "H" Then
            keisu = keisu2
        ElseIf ebmd.GuaranteeBmd.cMeth = "2" And ebmd.GuaranteeBmd.cCount = "4" And ebmd.GuaranteeBmd.cPos = "8" Then
            keisu = keisu3
        ElseIf ebmd.GuaranteeBmd.cMeth = "2" And ebmd.GuaranteeBmd.cCount = "3" And ebmd.GuaranteeBmd.cPos = "8" Then
            keisu = keisu4
        ElseIf ebmd.GuaranteeBmd.cMeth = "2" And ebmd.GuaranteeBmd.cCount = "5" And ebmd.GuaranteeBmd.cPos = "A" Then
            keisu = keisu5
        ElseIf ebmd.GuaranteeBmd.cMeth = "G" And ebmd.GuaranteeBmd.cCount = "3" And ebmd.GuaranteeBmd.cPos = "8" Then
            keisu = keisu6
        ElseIf ebmd.GuaranteeBmd.cMeth = "2" And ebmd.GuaranteeBmd.cCount = "5" And ebmd.GuaranteeBmd.cPos = "8" Then
            keisu = keisu7
        ElseIf ebmd.GuaranteeBmd.cMeth = "8" And ebmd.GuaranteeBmd.cCount = "4" And ebmd.GuaranteeBmd.cPos = "8" Then
            keisu = keisu8
        Else
            keisu = -1
        End If

        If keisu <> -1 Then

            With recX005
                If IsNull(.Fields("EPBMD" & j & "_MESDATA1").Value) = False Then
                    ebmd.BMD(0) = NtoZ2(.Fields("EPBMD" & j & "_MESDATA1").Value)   'BMD����l1
                Else
                    ebmd.BMD(0) = -1
                End If
                If IsNull(.Fields("EPBMD" & j & "_MESDATA2").Value) = False Then
                    ebmd.BMD(1) = NtoZ2(.Fields("EPBMD" & j & "_MESDATA2").Value)   'BMD����l2
                Else
                    ebmd.BMD(1) = -1
                End If
                If IsNull(.Fields("EPBMD" & j & "_MESDATA3").Value) = False Then
                    ebmd.BMD(2) = NtoZ2(.Fields("EPBMD" & j & "_MESDATA3").Value)   'BMD����l3
                Else
                    ebmd.BMD(2) = -1
                End If
                If IsNull(.Fields("EPBMD" & j & "_MESDATA4").Value) = False Then
                    ebmd.BMD(3) = NtoZ2(.Fields("EPBMD" & j & "_MESDATA4").Value)   'BMD����l4
                Else
                    ebmd.BMD(3) = -1
                End If
                If IsNull(.Fields("EPBMD" & j & "_MESDATA5").Value) = False Then
                    ebmd.BMD(4) = NtoZ2(.Fields("EPBMD" & j & "_MESDATA5").Value)   'BMD����l5
                Else
                    ebmd.BMD(4) = -1
                End If
                For k = 0 To 4
                    ebmd.BMD(k) = IIf(ebmd.BMD(k) <> -1, ebmd.BMD(k) * CDbl(keisu / 10000), -1)
                Next
            End With
                
            'EPBMDx_���莞��MAX�l_x
            '���躰�ނ� "F"�FMAX(2,4�_��)�C"G"�FMAX(2,3,4�_��) �̏ꍇ��MAX�l�ɂ��̒l���
            If ebmd.GuaranteeBmd.cJudg = JudgCodeW01 And _
                (ebmd.GuaranteeBmd.cObj = ObjCode10 Or ebmd.GuaranteeBmd.cObj = ObjCode11) Then
            
                If GetWfJudgData(WFBMD_JUDG, ebmd.GuaranteeBmd, ebmd.BMD(), JData()) = FUNCTION_RETURN_FAILURE Then
                    GoTo proc_exit
                End If
                recX004.Fields("EPBMD" & j & "_MAX").Value = JData(0)
            Else
                recX004.Fields("EPBMD" & j & "_MAX").Value = JudgMax(ebmd.BMD())
            End If
            recX004.Fields("EPBMD" & j & "_AVE").Value = JudgAve(ebmd.BMD())        'EPBMDx_���莞��AVE�l_x
            recX004.Fields("EPBMD" & j & "_MIN").Value = JudgMin(ebmd.BMD())        'EPBMDx_���莞��MIN�l_x
            If ebmd.SpecBmdMCL = "P " Then
                recX004.Fields("EPBMD" & j & "_MBP").Value = JudgBmdMBP(ebmd.BMD()) 'EPBMDx_���莞�̖ʓ����z
            Else
                recX004.Fields("EPBMD" & j & "_MBP").Value = 0                      '�ʓ����z��"P"�ȊO�̎��͌v�Z���ʂ�0�Ƃ���
            End If
            
'' 2008/10/20 BMD�]��,�O��1�_�ۏ؋@�\�ǉ� ADD By Systech Start
            If ebmd.GuaranteeBmd.cObj = ObjCode18 Then
                recX004.Fields("EPBMD" & j & "_MBP").Value = ebmd.BMD(0)
            End If
'' 2008/10/20 BMD�]��,�O��1�_�ۏ؋@�\�ǉ� ADD By Systech End
            
        End If
        
        '�ۏؕ��@="H"���ʓ����z��-1�̏ꍇ�װ�Ƃ���
        If ((ebmd.GuaranteeBmd.cJudg = "H") And CheckKHN_EP(HEPBMDKN, j, sPos)) _
            And (recX004.Fields("EPBMD" & j & "_MBP").Value = -1) Then GoTo proc_exit
        
    End If

    getTBCMY022EPBMD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY022EPBMD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :���L�T���v���`�F�b�N����
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :inSXLID         , I  ,String            , SXL-ID
'          :inSMPLID        , I  ,String            , �����ID
'          :outSMPLID       , O  ,String            , ���L�����ID(���L�łȂ��ꍇ�AinSMPLID��Ԃ�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :�w�肳�ꂽ�����ID���S���L���ǂ������������A�S���L�̏ꍇ�A���L�����ID���擾���Ԃ�
'����      :2003/11/19 SystemBrain �V�K�쐬
Private Function chkComSAMPL(inSXLID As String, inSMPLID As String, outSMPLID As String) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim wXTALCW     As String
    Dim wINPOSCW    As Integer
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function chkComSAMPL"
    
    '-------------------- �����ر ----------------------------------------
    chkComSAMPL = FUNCTION_RETURN_SUCCESS
    outSMPLID = inSMPLID
    
    '-------------------- �S���L�m�F(XSDCW) ----------------------------------------
    sql = "select XTALCW, INPOSCW from XSDCW "
    sql = sql & "where SXLIDCW = '" & inSXLID & "' and "
    sql = sql & "      REPSMPLIDCW = '" & inSMPLID & "' and "
    sql = sql & "      (WFINDRSCW = '2' or WFINDRSCW = '0' or WFINDRSCW = ' ' or WFINDRSCW is null) and "
    sql = sql & "      (WFINDOICW = '2' or WFINDOICW = '0' or WFINDOICW = ' ' or WFINDOICW is null) and "
    sql = sql & "      (WFINDB1CW = '2' or WFINDB1CW = '0' or WFINDB1CW = ' ' or WFINDB1CW is null) and "
    sql = sql & "      (WFINDB2CW = '2' or WFINDB2CW = '0' or WFINDB2CW = ' ' or WFINDB2CW is null) and "
    sql = sql & "      (WFINDB2CW = '2' or WFINDB3CW = '0' or WFINDB3CW = ' ' or WFINDB3CW is null) and "
    sql = sql & "      (WFINDL1CW = '2' or WFINDL1CW = '0' or WFINDL1CW = ' ' or WFINDL1CW is null) and "
    sql = sql & "      (WFINDL2CW = '2' or WFINDL2CW = '0' or WFINDL2CW = ' ' or WFINDL2CW is null) and "
    sql = sql & "      (WFINDL3CW = '2' or WFINDL3CW = '0' or WFINDL3CW = ' ' or WFINDL3CW is null) and "
    sql = sql & "      (WFINDL4CW = '2' or WFINDL4CW = '0' or WFINDL4CW = ' ' or WFINDL4CW is null) and "
    sql = sql & "      (WFINDDSCW = '2' or WFINDDSCW = '0' or WFINDDSCW = ' ' or WFINDDSCW is null) and "
    sql = sql & "      (WFINDDZCW = '2' or WFINDDZCW = '0' or WFINDDZCW = ' ' or WFINDDZCW is null) and "
    sql = sql & "      (WFINDSPCW = '2' or WFINDSPCW = '0' or WFINDSPCW = ' ' or WFINDSPCW is null) and "
    sql = sql & "      (WFINDDO1CW = '2' or WFINDDO1CW = '0' or WFINDDO1CW = ' ' or WFINDDO1CW is null) and "
    sql = sql & "      (WFINDDO2CW = '2' or WFINDDO2CW = '0' or WFINDDO2CW = ' ' or WFINDDO2CW is null) and "
    sql = sql & "      (WFINDDO3CW = '2' or WFINDDO3CW = '0' or WFINDDO3CW = ' ' or WFINDDO3CW is null) and "
    sql = sql & "      (WFINDOT1CW = '2' or WFINDOT1CW = '0' or WFINDOT1CW = ' ' or WFINDOT1CW is null) and "
    sql = sql & "      (WFINDOT2CW = '2' or WFINDOT2CW = '0' or WFINDOT2CW = ' ' or WFINDOT2CW is null) and "
    ''�c���_�f�ǉ��@03/12/19 ooba
    sql = sql & "      (WFINDAOICW = '2' or WFINDAOICW = '0' or WFINDAOICW = ' ' or WFINDAOICW is null) and "
    ''GD�ǉ��@2005/02/17 ffc)tanabe
    sql = sql & "      (((WFINDGDCW = '2' or WFINDGDCW = '0' or WFINDGDCW = ' ' or WFINDGDCW is null) and WFHSGDCW = '0') or WFHSGDCW = '1') "
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    sql = sql & "  and (EPINDB1CW = '2' or EPINDB1CW = '0' or EPINDB1CW = ' ' or EPINDB1CW is null) and "
    sql = sql & "      (EPINDB2CW = '2' or EPINDB2CW = '0' or EPINDB2CW = ' ' or EPINDB2CW is null) and "
    sql = sql & "      (EPINDB3CW = '2' or EPINDB3CW = '0' or EPINDB3CW = ' ' or EPINDB3CW is null) and "
    sql = sql & "      (EPINDL1CW = '2' or EPINDL1CW = '0' or EPINDL1CW = ' ' or EPINDL1CW is null) and "
    sql = sql & "      (EPINDL2CW = '2' or EPINDL2CW = '0' or EPINDL2CW = ' ' or EPINDL2CW is null) and "
    sql = sql & "      (EPINDL3CW = '2' or EPINDL3CW = '0' or EPINDL3CW = ' ' or EPINDL3CW is null) "
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    wXTALCW = rs("XTALCW")      '�����ԍ�
    wINPOSCW = rs("INPOSCW")    '�������ʒu
    Set rs = Nothing
    
    '-------------------- ���L�����ID�̎擾(XSDCW) ----------------------------------------
    sql = "select REPSMPLIDCW from XSDCW "
    sql = sql & "where SXLIDCW like '" & left(wXTALCW, 9) & "%' and "       '09/05/26 ooba
    sql = sql & "      XTALCW = '" & wXTALCW & "' and "
    sql = sql & "      INPOSCW = '" & wINPOSCW & "' and "
    sql = sql & "      NUKISIFLGCW = '1' and "                              '09/05/26 ooba
    sql = sql & "      SXLIDCW != '" & inSXLID & "' and "
    sql = sql & "      REPSMPLIDCW != '" & inSMPLID & "' "
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    outSMPLID = rs("REPSMPLIDCW")       '��\�����ID(���L)
    Set rs = Nothing

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    chkComSAMPL = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :SXL�m��w��(TBCMY007)ð��قɾ�Ă���SXL�̔��R�ް����擾����B
'���Ұ�    :�ϐ���          ,IO  ,�^                :����
'          :SXLID          ,I   ,String            ,SXLID
'�@�@      :sPos  �@�@�@    ,I   ,String �@         ,SXL�ʒu(TOP/BOT)   04/04/15 ooba
'          :sPattern       ,I   ,String            ,���R�ް��擾�����
'                                                   �������A : WF�����ް��擾
'                                                   �������B : ���������ް��擾
'                                                   �������C : �擾�ް��Ȃ�
'          :mesdata()      ,O   ,String            ,���R�ް�
'          :�߂�l          ,O   ,FUNCTION_RETURN   ,���o�̐���
'����      :
'����      :04/02/12 ooba�@�쐬
Public Function cmbc040_GetSxlRsData(SXLID As String, sPos As String, sPattern As String, mesdata() As String) As FUNCTION_RETURN
    
    Dim sTBkbn As String        'T/B�敪
    Dim i As Integer
    Dim j As Integer
    Dim sSql As String
    Dim rs As OraDynaset
    Dim dTmpData(10) As Double   '���R(Rs)�ް�
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function cmbc040_GetSxlRsData"
    cmbc040_GetSxlRsData = FUNCTION_RETURN_FAILURE
    
    If sPos = "TOP" Then sTBkbn = "T" Else sTBkbn = "B"  '04/04/15 ooba
    
    '���R�ް��擾����݂��wA�x�̏ꍇ�AWF�����ް�(TBCMY013)���擾����B
    If sPattern = "A" Then
'''        For i = 1 To 2
'''            If i = 1 Then sTBkbn = "T" Else sTBkbn = "B"
        '�Y��SXL���A�V����يǗ�-WF<XSDCW>�̻����ID_Rs���擾�B
        '�����ID_Rs����A����]������<TBCMY013>�̔��R�����ް�(TOP��/BOT��)���擾����B
        sSql = "select MESDATA1, MESDATA2, MESDATA3, MESDATA4, MESDATA5 "
        sSql = sSql & "from TBCMY013 "
        sSql = sSql & "where OSITEM = 'RES' "
        sSql = sSql & "and SAMPLEID in ( "
        sSql = sSql & "         select WFSMPLIDRSCW from XSDCW "
        sSql = sSql & "         where TBKBNCW = '" & sTBkbn & "' "
        sSql = sSql & "         and SXLIDCW = '" & SXLID & "') "
        
        Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
        
        If rs.RecordCount > 0 Then
            'TOP�������ް�
            If sTBkbn = "T" Then
                mesdata(1) = rs("MESDATA1")
                mesdata(2) = rs("MESDATA2")
                mesdata(3) = rs("MESDATA3")
                mesdata(4) = rs("MESDATA4")
                mesdata(5) = rs("MESDATA5")
            'BOT�������ް�
            ElseIf sTBkbn = "B" Then
                mesdata(6) = rs("MESDATA1")
                mesdata(7) = rs("MESDATA2")
                mesdata(8) = rs("MESDATA3")
                mesdata(9) = rs("MESDATA4")
                mesdata(10) = rs("MESDATA5")
            End If
        Else
            '�����ް����Ȃ��ꍇ�ʹװ
            Set rs = Nothing
            GoTo proc_exit
        End If
        Set rs = Nothing
'''        Next
    '���R�ް��擾����݂��wB�x�̏ꍇ�A���������ް�(TBCMJ002)���擾����B
    ElseIf sPattern = "B" Then
'''        For i = 1 To 2
'''            If i = 1 Then sTBkbn = "T" Else sTBkbn = "B"
        '�Y��SXL���A�V����يǗ�-WF<XSDCW>��T/B�敪�A�������ۯ�ID���擾�B
        'T/B�敪�A�������ۯ�ID����A�V����يǗ�-��ۯ�<XSDCS>�̌����ԍ��A�����ID_Rs���擾�B
        '�����ԍ��A�����ID_Rs����A������R����<TBCMJ002>�̔��R�����ް�(TOP��/BOT��)���擾����B
        sSql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 "
        sSql = sSql & "from TBCMJ002 "
        sSql = sSql & "where (CRYNUM, SMPLNO) in ( "
        sSql = sSql & "         select XTALCS, CRYSMPLIDRSCS "
        sSql = sSql & "         from XSDCS "
        sSql = sSql & "         where (TBKBNCS, CRYNUMCS) in ( "
        sSql = sSql & "                  select TBKBNCW, SMCRYNUMCW "
        sSql = sSql & "                  from XSDCW "
        sSql = sSql & "                  where TBKBNCW = '" & sTBkbn & "' "
        sSql = sSql & "                  and SXLIDCW = '" & SXLID & "')) "
        sSql = sSql & "and TRANCNT = ( "
        sSql = sSql & "         select max(TRANCNT) "
        sSql = sSql & "         from TBCMJ002 "
        sSql = sSql & "         where (CRYNUM, SMPLNO) in ( "
        sSql = sSql & "                  select XTALCS, CRYSMPLIDRSCS "
        sSql = sSql & "                  from XSDCS "
        sSql = sSql & "                  where (TBKBNCS, CRYNUMCS) in ( "
        sSql = sSql & "                           select TBKBNCW, SMCRYNUMCW "
        sSql = sSql & "                           from XSDCW "
        sSql = sSql & "                           where TBKBNCW = '" & sTBkbn & "' "
        sSql = sSql & "                           and SXLIDCW = '" & SXLID & "'))) "
    
        Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
    
        If rs.RecordCount > 0 Then
            'TOP�������ް�
            If sTBkbn = "T" Then
                dTmpData(1) = rs("MEAS1")
                dTmpData(2) = rs("MEAS2")
                dTmpData(3) = rs("MEAS3")
                dTmpData(4) = rs("MEAS4")
                dTmpData(5) = rs("MEAS5")
                '�^�ϊ�
                For j = 1 To 5
                    mesdata(j) = CStr(dTmpData(j))
                Next
            'BOT�������ް�
            ElseIf sTBkbn = "B" Then
                dTmpData(6) = rs("MEAS1")
                dTmpData(7) = rs("MEAS2")
                dTmpData(8) = rs("MEAS3")
                dTmpData(9) = rs("MEAS4")
                dTmpData(10) = rs("MEAS5")
                '�^�ϊ�
                For j = 6 To 10
                    mesdata(j) = CStr(dTmpData(j))
                Next
            End If
        Else
            '�����ް����Ȃ��ꍇ�ʹװ
            Set rs = Nothing
            GoTo proc_exit
        End If
        Set rs = Nothing
'''        Next
    '���R�ް��擾����݂��wC�x�̏ꍇ�A�擾�����ް��Ȃ��B
    ElseIf sPattern = "C" Then
    
    End If
    
    '�擾�ް�����/-1/NULL�̎��ͽ�߰���Ă���B
'''    For i = 1 To 10
'''        If mesdata(i) = "" Or mesdata(i) = "-1" Or mesdata(i) = vbNullString Then
'''            mesdata(i) = " "
'''        End If
'''    Next
    For i = 1 To 5
        If sTBkbn = "T" Then j = i Else j = i + 5
        If mesdata(j) = "" Or mesdata(j) = "-1" Or mesdata(j) = vbNullString Then
            mesdata(j) = " "
        End If
    Next
    
    cmbc040_GetSxlRsData = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    cmbc040_GetSxlRsData = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�_�f�͏o�Ǝc���_�f�̎d�l�`�F�b�N
'���Ұ��@�@:�ϐ���        ,IO ,�^              ,����
'      �@�@:pHin�@�@    �@,I  ,tFullHinban   �@,�i��
'      �@�@:�߂�l        ,O  ,Integer       �@,�d�l�`�F�b�N����(-1:�װ�C0:AOi�d�l���C1:AOi�d�l�L)
'����      :�_�f�͏o(��oi)�Ǝc���_�f�̗����Ɏd�l�������Ă����ꍇ�G���[��Ԃ�
'          :��s_cmzcF_cmkc001WF.bas����֐��Ɠ��l
'����      :03/12/19 ooba

Public Function ChkAoiSiyou(pHIN As tFullHinban) As Integer

    Dim sSql As String
    Dim rs As OraDynaset
    Dim sDoiSiyou(2) As String  '�����L��(DOi1�`3)
    Dim sAoiSiyou As String     '�����L��(AOi)
    Dim iCnt As Integer
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001WF.bas -- Function ChkAoiSiyou"

    sSql = "select HWFOS1HS, HWFOS2HS, HWFOS3HS, HWFZOHWS from TBCME025 "
    sSql = sSql & "where HINBAN = '" & pHIN.hinban & "' "
    sSql = sSql & "and MNOREVNO = " & pHIN.mnorevno & " "
    sSql = sSql & "and FACTORY = '" & pHIN.factory & "' "
    sSql = sSql & "and OPECOND = '" & pHIN.opecond & "' "

    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount = 0 Then
        rs.Close
        ChkAoiSiyou = -1
        GoTo proc_exit
    End If
    
    If IsNull(rs("HWFOS1HS")) = False Then sDoiSiyou(0) = rs("HWFOS1HS") '�iWF�_�f�͏o1�ۏؕ��@_��
    If IsNull(rs("HWFOS2HS")) = False Then sDoiSiyou(1) = rs("HWFOS2HS") '�iWF�_�f�͏o2�ۏؕ��@_��
    If IsNull(rs("HWFOS3HS")) = False Then sDoiSiyou(2) = rs("HWFOS3HS") '�iWF�_�f�͏o3�ۏؕ��@_��
    If IsNull(rs("HWFZOHWS")) = False Then sAoiSiyou = rs("HWFZOHWS")    '�iWF�c���_�f�ۏؕ��@_��
    
    '�_�f�͏o�Ǝc���_�f�̎d�l�`�F�b�N
    ChkAoiSiyou = 0
    For iCnt = 0 To 2
        If sDoiSiyou(iCnt) = "H" Or sDoiSiyou(iCnt) = "S" Then
            '�_�f�͏o(��oi)�Ǝc���_�f�̗����Ɏd�l�������Ă����ꍇ�̓G���[
            If sAoiSiyou = "H" Or sAoiSiyou = "S" Then
                ChkAoiSiyou = -1
                Exit For
            End If
        Else
            If sAoiSiyou = "H" Or sAoiSiyou = "S" Then
                ChkAoiSiyou = 1
            End If
        End If
    Next
    
proc_exit:
    '' �I��
    gErr.Pop
    Exit Function
    
proc_err:
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    ChkAoiSiyou = -1
    Resume proc_exit
    
End Function


'------------------------------------------------------------------------------------
'-�T�v    :�v�e�z�[���h�R�[�h�����R�擾����
'-���Ұ�  :�ϐ���        ,IO  ,�^                                     ,����
'-        :sSXLID        ,I   ,string                                ,�V���O��ID
'-        :sBLOCKID      ,I   ,string                                ,�u���b�NID
'-        :sHINBAN       ,I   ,string                                ,�i��
'-        :sINGOTPOS     ,I   ,string                                ,�����ʒu
'-        :sWFHOLDDATE   ,O   ,string                                ,�z�[���h���t
'-        :sUSER_ID      ,O   ,string                                ,WF�z�[���h������ID
'-        :��ؒl         ,O   ,FUNCTION_RETURN                       ,�ǂݍ��ݐ���
'-����    :TBCMY019[KEY:BLOCKID,TRANCNT]�f�[�^���擾����B
'-         TBCMY019����擾�����R�[�h������KODA9[KEY:]���痝�R(���{��)���擾����B
'-����    :�c�a�X�V�ǉ��@2004/07/16 KOYAMA
'------------------------------------------------------------------------------------
Public Function DBDRV_s_cmbc040_SQL_Y019XSDCB(sSXLID As String, sBlockId As String, sHINBAN As String, _
                                                 sINGOTPOS As String, sWFHOLDDATE As String, _
                                                 sUSER_ID As String) As FUNCTION_RETURN



    Dim cbrs As OraDynaset         'XSDCB�����p�J�[�\��
    Dim wfrs As OraDynaset         'TBCMY019�����p�J�[�\��
    Dim sql As String
    Dim ksql As String
    Dim recCnt As Long
    Dim swfkbn As String           'WF�z�[���h�敪(1:WF�z�[���h�A0:�z�[���h����)

    '�ϐ�������
    sql = ""
    ksql = ""
    recCnt = 0

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- DBDRV_s_cmbc040_SQL_Y019XSDCB"

    'XSDCB����(WF�z�[���h�敪����)
    sql = ""
    sql = "select "
    sql = sql & " WFHOLDFLGCB "                  ' WF�z�[���h�敪
    sql = sql & " from XSDCB "
    sql = sql & " where "
    sql = sql & " SXLIDCB = '" & sSXLID & "'"
    sql = sql & " and XTALCB = '" & sBlockId & "'"
    
    
 '   Debug.Print sql
    Set cbrs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    '���R�[�h0�����͐���
        
    If cbrs Is Nothing Then
        DBDRV_s_cmbc040_SQL_Y019XSDCB = FUNCTION_RETURN_FAILURE
        swfkbn = ""
        cbrs.Close
        GoTo proc_exit
    End If

    recCnt = cbrs.RecordCount
    '���R�[�h0�����͏������t�AWF�z�[���h������ID���X�y�[�X�ŕԂ��B
    If recCnt = 0 Then
        swfkbn = ""
        DBDRV_s_cmbc040_SQL_Y019XSDCB = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    Else
        swfkbn = IIf(IsNull(cbrs("WFHOLDFLGCB")), "", cbrs("WFHOLDFLGCB"))  ' WF�z�[���h�敪
        cbrs.Close
    End If

    
    '**** WF�z�[���h��Ԃł���ꍇ(WF�z�[���h�敪=1)
    If swfkbn = "1" Then
    
    
        'TBCMY019����(�v�e�z�[���h�������A�z�[���h������ID�擾)
        sql = ""
        sql = "select "
        sql = sql & " HOLDDT, "                  ' �z�[���h���t
        sql = sql & " USER_ID "                  ' WF�z�[���h������ID
        sql = sql & " from TBCMY019 "
        sql = sql & " where "
        sql = sql & " TRANCNT = any(select MAX(TRANCNT)"
        sql = sql & " from TBCMY019 where BLOCKID ='" & sBlockId & "')"
        sql = sql & " and BLOCKID ='" & sBlockId & "'"
        
        Debug.Print sql
        Set wfrs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        '���R�[�h0�����͐���
            
        If wfrs Is Nothing Then
            DBDRV_s_cmbc040_SQL_Y019XSDCB = FUNCTION_RETURN_FAILURE
            wfrs.Close
            GoTo proc_exit
        End If
    
        recCnt = wfrs.RecordCount
        '���R�[�h0�����͏������t�AWF�z�[���h������ID���X�y�[�X�ŕԂ��B
        If recCnt = 0 Then
            sWFHOLDDATE = ""
            sUSER_ID = ""
        Else
            sWFHOLDDATE = IIf(IsNull(wfrs("HOLDDT")), "", wfrs("HOLDDT"))    ' �z�[���h���t
            sUSER_ID = IIf(IsNull(wfrs("USER_ID")), "", wfrs("USER_ID"))     ' WF�z�[���h������ID
            wfrs.Close
        End If
    
    
    '**** WF�z�[���h��ԈȊO�ꍇ(WF�z�[���h�敪=0)
    Else
        DBDRV_s_cmbc040_SQL_Y019XSDCB = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    DBDRV_s_cmbc040_SQL_Y019XSDCB = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_s_cmbc040_SQL_Y019XSDCB = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
    
End Function

'�T�v      :WFGD����(TBCMJ015)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW�\����   (�V����يǗ�(SXL))
'          :recX003         , O  ,c_cmzcrec         , TBCMX003�\����(GD��������_�ް�)
'          :sTblName        , I  ,String            , �e�[�u�����@2011/06/23 Marushita GBG���M�Ή�
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :WFGD����(TBCMJ015)�����ް����擾���AGD��������_�ް��\���̂ɾ�Ă���
'����      :2005/02/15 ffc)tanabe
'Private Function getTBCMJ015WFGD(CRYNUM As String, recXSDCW As c_cmzcrec, recX003 As c_cmzcrec) As FUNCTION_RETURN
Private Function getTBCMJ015WFGD(CRYNUM As String, recXSDCW As c_cmzcrec, recX003 As c_cmzcrec, sTblName As String) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    Dim nFlg    As Integer
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ015WFGD"
    
    getTBCMJ015WFGD = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    '>>>>> 2011/06/24 SETsw)Marushita WFGD_���莞�̃Z�b�g�Ή�
    'GBG�̃`�F�b�N
    If FieldCheck(sTblName, "SXLGD_HSFLG") = FUNCTION_RETURN_FAILURE Then
        nFlg = 1
    Else
        nFlg = 0
    End If
    '<<<<< 2011/06/24 SETsw)Marushita WFGD_���莞�̃Z�b�g�Ή�
    'TBCMX003
    With recX003
        .Fields("WFGD_SMPPOS").Value = vbNullString                              'WFGD�T���v������ʒu(SXL�ʒu���)
        .Fields("WFGD_MSRSDEN").Value = vbNullString                             'WFGD_���茋�� Den
        .Fields("WFGD_MSRSLDL").Value = vbNullString                             'WFGD_���茋�� L/DL
        .Fields("WFGD_MSRSDVD2").Value = vbNullString                            'WFGD_���茋�� DVD2
        '>>>>> 2011/06/24 SETsw)Marushita WFGD_���莞�̃Z�b�g�Ή�
        If nFlg = 1 Then
        Else
        '<<<<< 2011/06/24 SETsw)Marushita WFGD_���莞�̃Z�b�g�Ή�
            .Fields("SXLGD_HSFLG").Value = vbNullString                              'SXLGDGD���茋�ʕۏ؃t���O
            .Fields("SXLGD_SMPPOS").Value = vbNullString                             'SXLGDGD�T���v������ʒu(SXL�ʒu���)
            .Fields("SXLGD_MSRSDEN").Value = vbNullString                            'SXLGDGD_���茋�� Den
            .Fields("SXLGD_MSRSLDL").Value = vbNullString                            'SXLGDGD_���茋�� L/DL
            .Fields("SXLGD_MSRSDVD2").Value = vbNullString                           'SXLGDGD_���茋�� DVD2
            .Fields("WFGD_HSFLG").Value = vbNullString                               'WFGD���茋�ʕۏ؃t���O
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
            .Fields("GD_PTNJUDGRES").Value = vbNullString                            'GD�p�^�[�����茋��
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
            
            For i = 1 To 15
                .Fields("WFGD_MS" & Format(i, "00") & "LDL1").Value = vbNullString       'WFGD_����lxx L/DL1
                .Fields("WFGD_MS" & Format(i, "00") & "LDL2").Value = vbNullString       'WFGD_����lxx L/DL2
                .Fields("WFGD_MS" & Format(i, "00") & "LDL3").Value = vbNullString       'WFGD_����lxx L/DL3
                .Fields("WFGD_MS" & Format(i, "00") & "LDL4").Value = vbNullString       'WFGD_����lxx L/DL4
                .Fields("WFGD_MS" & Format(i, "00") & "LDL5").Value = vbNullString       'WFGD_����lxx L/DL5
                .Fields("WFGD_MS" & Format(i, "00") & "DEN1").Value = vbNullString       'WFGD_����lxx Den1
                .Fields("WFGD_MS" & Format(i, "00") & "DEN2").Value = vbNullString       'WFGD_����lxx Den2
                .Fields("WFGD_MS" & Format(i, "00") & "DEN3").Value = vbNullString       'WFGD_����lxx Den3
                .Fields("WFGD_MS" & Format(i, "00") & "DEN4").Value = vbNullString       'WFGD_����lxx Den4
                .Fields("WFGD_MS" & Format(i, "00") & "DEN5").Value = vbNullString       'WFGD_����lxx Den5
            Next
            
            For i = 1 To 5
                .Fields("WFGD_MS01DVD2" & i).Value = vbNullString                        'WFGD_����lxx DVD2
            Next
        End If
    End With
        
    '-------------------- TBCMJ015�̓ǂݍ���(GD) ----------------------------------------
    sql = "select * from TBCMJ015 "
    sql = sql & " where CRYNUM = '" & CRYNUM & "'"
    sql = sql & " and   SMPLNO = '" & recXSDCW("WFSMPLIDGDCW").Value & "'"
    sql = sql & " and   HSFLG = '1'"
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        '>>>>> 2011/06/24 SETsw)Marushita WFGD_���莞�̃Z�b�g�Ή�
        'GD���т��Ȃ��ꍇ�̓G���[�ɂ��Ȃ��H
        If nFlg = 1 Then
            getTBCMJ015WFGD = FUNCTION_RETURN_SUCCESS
        End If
        '<<<<< 2011/06/24 SETsw)Marushita WFGD_���莞�̃Z�b�g�Ή�
        GoTo proc_exit
    End If
    
    'TBCMX003
    With recX003
        .Fields("WFGD_SMPPOS").Value = rs("POSITION")                                                     'WFGD�T���v������ʒu(SXL�ʒu���)
        .Fields("WFGD_MSRSDEN").Value = rs("MSRSDEN")                                                     'WFGD_���茋�� Den
        .Fields("WFGD_MSRSLDL").Value = rs("MSRSLDL")                                                     'WFGD_���茋�� L/DL
        .Fields("WFGD_MSRSDVD2").Value = rs("MSRSDVD2")                                                   'WFGD_���茋�� DVD2
        If nFlg = 1 Then
        Else
            .Fields("WFGD_HSFLG").Value = "1"                                                                 'WFGD���茋�ʕۏ؃t���O
            
            For i = 1 To 15
                .Fields("WFGD_MS" & Format(i, "00") & "DEN1").Value = rs("MS" & Format(i, "00") & "DEN1")      'WFGD_����lxx Den1
                .Fields("WFGD_MS" & Format(i, "00") & "DEN2").Value = rs("MS" & Format(i, "00") & "DEN2")      'WFGD_����lxx Den2
                .Fields("WFGD_MS" & Format(i, "00") & "DEN3").Value = rs("MS" & Format(i, "00") & "DEN3")      'WFGD_����lxx Den3
                .Fields("WFGD_MS" & Format(i, "00") & "DEN4").Value = rs("MS" & Format(i, "00") & "DEN4")      'WFGD_����lxx Den4
                .Fields("WFGD_MS" & Format(i, "00") & "DEN5").Value = rs("MS" & Format(i, "00") & "DEN5")      'WFGD_����lxx Den5
                .Fields("WFGD_MS" & Format(i, "00") & "LDL1").Value = rs("MS" & Format(i, "00") & "LDL1")      'WFGD_����lxx L/DL1
                .Fields("WFGD_MS" & Format(i, "00") & "LDL2").Value = rs("MS" & Format(i, "00") & "LDL2")      'WFGD_����lxx L/DL2
                .Fields("WFGD_MS" & Format(i, "00") & "LDL3").Value = rs("MS" & Format(i, "00") & "LDL3")      'WFGD_����lxx L/DL3
                .Fields("WFGD_MS" & Format(i, "00") & "LDL4").Value = rs("MS" & Format(i, "00") & "LDL4")      'WFGD_����lxx L/DL4
                .Fields("WFGD_MS" & Format(i, "00") & "LDL5").Value = rs("MS" & Format(i, "00") & "LDL5")      'WFGD_����lxx L/DL5
            Next
            
            For i = 1 To 5
                .Fields("WFGD_MS01DVD2" & i).Value = rs("MS0" & i & "DVD2")                                    'WFGD_����lxx DVD2
            Next
        
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
            'GD�p�^�[�����茋��
            If IsNull(.Fields("PTNJUDGRES")) = True Then
                .Fields("GD_PTNJUDGRES").Value = " "
            Else
                .Fields("GD_PTNJUDGRES").Value = rs("PTNJUDGRES")
            End If
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
        End If
    End With
    Set rs = Nothing

    getTBCMJ015WFGD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ015WFGD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :����GD����(TBCMJ006)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS�\����   (�V����يǗ�(��ۯ�))
'          :recX003         , O  ,c_cmzcrec         , TBCMX003�\����(GD��������_�ް�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :����GD����(TBCMJ006)�����ް����擾���AGD��������_�ް��\���̂ɾ�Ă���
'          :����GD����(TBCMJ006)�̑���f�[�^�̏����l�ł���-1��NULL�ɕύX����TBCMX003�ɓo�^����B
'����      :2005/02/15 ffc)tanabe
Private Function getTBCMJ006GD(CRYNUM As String, recXSDCW As c_cmzcrec, recX003 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ006GD"
    
    getTBCMJ006GD = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
        
    'TBCMX003
    With recX003
            .Fields("SXLGD_HSFLG").Value = vbNullString                              'SXLGDGD���茋�ʕۏ؃t���O
            .Fields("SXLGD_SMPPOS").Value = vbNullString                             'SXLGDGD�T���v������ʒu(SXL�ʒu���)
            .Fields("SXLGD_MSRSDEN").Value = vbNullString                            'SXLGDGD_���茋�� Den
            .Fields("SXLGD_MSRSLDL").Value = vbNullString                            'SXLGDGD_���茋�� L/DL
            .Fields("SXLGD_MSRSDVD2").Value = vbNullString                           'SXLGDGD_���茋�� DVD2
            .Fields("WFGD_HSFLG").Value = vbNullString                               'WFGD���茋�ʕۏ؃t���O
            .Fields("WFGD_SMPPOS").Value = vbNullString                              'WFGD�T���v������ʒu(SXL�ʒu���)
            .Fields("WFGD_MSRSDEN").Value = vbNullString                             'WFGD_���茋�� Den
            .Fields("WFGD_MSRSLDL").Value = vbNullString                             'WFGD_���茋�� L/DL
            .Fields("WFGD_MSRSDVD2").Value = vbNullString                            'WFGD_���茋�� DVD2
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
            .Fields("GD_PTNJUDGRES").Value = vbNullString                            'GD�p�^�[�����茋��
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End
            
        For i = 1 To 15
            .Fields("WFGD_MS" & Format(i, "00") & "LDL1").Value = vbNullString       'WFGD_����lxx L/DL1
            .Fields("WFGD_MS" & Format(i, "00") & "LDL2").Value = vbNullString       'WFGD_����lxx L/DL2
            .Fields("WFGD_MS" & Format(i, "00") & "LDL3").Value = vbNullString       'WFGD_����lxx L/DL3
            .Fields("WFGD_MS" & Format(i, "00") & "LDL4").Value = vbNullString       'WFGD_����lxx L/DL4
            .Fields("WFGD_MS" & Format(i, "00") & "LDL5").Value = vbNullString       'WFGD_����lxx L/DL5
            .Fields("WFGD_MS" & Format(i, "00") & "DEN1").Value = vbNullString       'WFGD_����lxx Den1
            .Fields("WFGD_MS" & Format(i, "00") & "DEN2").Value = vbNullString       'WFGD_����lxx Den2
            .Fields("WFGD_MS" & Format(i, "00") & "DEN3").Value = vbNullString       'WFGD_����lxx Den3
            .Fields("WFGD_MS" & Format(i, "00") & "DEN4").Value = vbNullString       'WFGD_����lxx Den4
            .Fields("WFGD_MS" & Format(i, "00") & "DEN5").Value = vbNullString       'WFGD_����lxx Den5
        Next
        
        For i = 1 To 5
            .Fields("WFGD_MS01DVD2" & i).Value = vbNullString                        'WFGD_����lxx DVD2
        Next
        
    End With
        
    '-------------------- TBCMJ006�̓ǂݍ���(GD) ----------------------------------------
    sql = "select * from TBCMJ006 "
    sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
    sql = sql & "      SMPLNO = " & Trim(recXSDCW("WFSMPLIDGDCW").Value)
    sql = sql & " order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    
    'TBCMX003
    With recX003
        .Fields("SXLGD_HSFLG").Value = "1"                          'SXLGD���茋�ʕۏ؃t���O
        .Fields("SXLGD_SMPPOS").Value = rs("POSITION")              'SXLGD�T���v������ʒu(SXL�ʒu���)
        If rs("MSRSDEN") <> -1 Then
            .Fields("SXLGD_MSRSDEN").Value = rs("MSRSDEN")          'SXLGD_���茋�� Den
        End If
        If rs("MSRSLDL") <> -1 Then
            .Fields("SXLGD_MSRSLDL").Value = rs("MSRSLDL")          'SXLGD_���茋�� L/DL
        End If
        If rs("MSRSDVD2") <> -1 Then
            .Fields("SXLGD_MSRSDVD2").Value = rs("MSRSDVD2")        'SXLGD_���茋�� DVD2
        End If
        
        For i = 1 To 15
            If rs("MS" & Format(i, "00") & "DEN1") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN1").Value = rs("MS" & Format(i, "00") & "DEN1")      'SXLGD_����lxx Den1
            End If
            If rs("MS" & Format(i, "00") & "DEN2") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN2").Value = rs("MS" & Format(i, "00") & "DEN2")      'SXLGD_����lxx Den2
            End If
            If rs("MS" & Format(i, "00") & "DEN3") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN3").Value = rs("MS" & Format(i, "00") & "DEN3")      'SXLGD_����lxx Den3
            End If
            If rs("MS" & Format(i, "00") & "DEN4") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN4").Value = rs("MS" & Format(i, "00") & "DEN4")      'SXLGD_����lxx Den4
            End If
            If rs("MS" & Format(i, "00") & "DEN5") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN5").Value = rs("MS" & Format(i, "00") & "DEN5")      'SXLGD_����lxx Den5
            End If
            If rs("MS" & Format(i, "00") & "LDL1") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL1").Value = rs("MS" & Format(i, "00") & "LDL1")      'SXLGD_����lxx L/DL1
            End If
            If rs("MS" & Format(i, "00") & "LDL2") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL2").Value = rs("MS" & Format(i, "00") & "LDL2")      'SXLGD_����lxx L/DL2
            End If
            If rs("MS" & Format(i, "00") & "LDL3") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL3").Value = rs("MS" & Format(i, "00") & "LDL3")      'SXLGD_����lxx L/DL3
            End If
            If rs("MS" & Format(i, "00") & "LDL4") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL4").Value = rs("MS" & Format(i, "00") & "LDL4")      'SXLGD_����lxx L/DL4
            End If
            If rs("MS" & Format(i, "00") & "LDL5") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL5").Value = rs("MS" & Format(i, "00") & "LDL5")      'SXLGD_����lxx L/DL5
            End If
        Next
        
        For i = 1 To 5
            If rs("MS0" & i & "DVD2") <> -1 Then
                .Fields("WFGD_MS01DVD2" & i).Value = rs("MS0" & i & "DVD2")         'SXLGD_����lxx DVD2
            End If
        Next
        
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech Start
        'GD�p�^�[�����茋��
        If IsNull(rs("PTNJUDGRES")) = True Then
            .Fields("GD_PTNJUDGRES").Value = " "
        Else
            .Fields("GD_PTNJUDGRES").Value = rs("PTNJUDGRES")
        End If
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� ADD By Systech End

    End With
    Set rs = Nothing

    getTBCMJ006GD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ006GD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

''Upd start 2005/06/23 (TCS)t.terauchi  SPV9�_�Ή�
'�T�v      :WFSPV����(TBCMJ016)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW�\����   (�V����يǗ�(SXL))
'          :HIN             , I  ,tFullHinban       , �i��(�S�i�ԍ\����)
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002�\����(����_�ް�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :WFSPV����(TBCMJ016)�����ް����擾���ASXL�������E����_�ް��\���̂ɾ�Ă���
'����      :2005/06/23  �V�K�쐬�@(TCS)t.terauchi
Private Function getTBCMJ016WFSPV(CRYNUM As String, recXSDCW As c_cmzcrec, HIN As tFullHinban, _
                                  recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset   'TBCMJ016�p
    Dim rs2         As OraDynaset   'TBCME028�p
    Dim sql         As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ016WFSPV"
    
    getTBCMJ016WFSPV = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("WFSPV_SMPPOS").Value = -1          'WFSPV�����-ID����ʒu(SXL�ʒu���)
        .Fields("WFSPV_NETSU").Value = ""           'WFSPV_�M��������
        .Fields("WFSPV_ET").Value = ""              'WFSPV_�G�b�`���O����
        .Fields("WFSPV_MES").Value = ""             'WFSPV_�v�����@
        .Fields("WFSPV_KST_MAX").Value = -1         'WFSPV_�g�U�����莞��MAX�l
        .Fields("WFSPV_KST_AVE").Value = -1         'WFSPV_�g�U�����莞��AVE�l
        .Fields("WFSPV_KST_MIN").Value = -1         'WFSPV_�g�U�����莞��MIN�l
        .Fields("WFSPV_FE_MAX").Value = -1          'WFSPV_Fe�Z�x���莞��MAX�l
        .Fields("WFSPV_FE_AVE").Value = -1          'WFSPV_Fe�Z�x���莞��AVE�l
        .Fields("WFSPV_FE_MIN").Value = -1          'WFSPV_Fe�Z�x���莞��MIN�l
        
        ''>>=====SPV����@20060529 SMP����
        .Fields("WFSPV_FE_PUA").Value = -1           ''SPV_Fe PUA�l
        .Fields("WFSPV_FE_PUAP").Value = -1          ''SPV_Fe PUA���l
        .Fields("WFSPV_FE_STD").Value = -1           ''SPV_Fe STD
        .Fields("WFSPV_DIFF_PUA").Value = -1         ''SPV_�g�U�� PUA�l
        .Fields("WFSPV_DIFF_PUAP").Value = -1        ''SPV_�g�U�� PUA���l
        .Fields("WFSPV_NR_MAX").Value = -1           ''SPV_OtherRecords_MAX
        .Fields("WFSPV_NR_AVE").Value = -1           ''SPV_OtherRecords_AVE
        .Fields("WFSPV_NR_STD").Value = -1           ''SPV_OtherRecords_STD
        .Fields("WFSPV_NR_PUA").Value = -1           ''SPV_OtherRecords_PUA�l
        .Fields("WFSPV_NR_PUAP").Value = -1          ''SPV_OtherRecords_PUA���l
        ''==============================<<
    End With
                
    'TBCMX002
    With recX002
        .Fields("WFSPV_SMPPOS").Value = -1          'WFSPV�����-ID����ʒu(SXL�ʒu���)
        .Fields("WFSPV_NETSU").Value = " "          'WFSPV_�M��������
        .Fields("WFSPV_ET").Value = " "             'WFSPV_�G�b�`���O����
        .Fields("WFSPV_MES").Value = " "            'WFSPV_�v�����@
        .Fields("WFSPV_DKAN").Value = " "           'WFSPV_�c�j�A�j�[������
        .Fields("WFSPV_MESDATA1").Value = " "       'WFSPV����_1
        .Fields("WFSPV_MESDATA2").Value = " "       'WFSPV����_2
        .Fields("WFSPV_MESDATA3").Value = " "       'WFSPV����_3
        .Fields("WFSPV_MESDATA4").Value = " "       'WFSPV����_4
        .Fields("WFSPV_MESDATA5").Value = " "       'WFSPV����_5
        .Fields("WFSPV_MESDATA6").Value = " "       'WFSPV����_6
        .Fields("WFSPV_MESDATA7").Value = " "       'WFSPV����_7
        .Fields("WFSPV_MESDATA8").Value = " "       'WFSPV����_8
        .Fields("WFSPV_MESDATA9").Value = " "       'WFSPV����_9
        .Fields("WFSPV_MESDATA10").Value = " "      'WFSPV����_10
        .Fields("WFSPV_MESDATA11").Value = " "      'WFSPV����_11
        .Fields("WFSPV_MESDATA12").Value = " "      'WFSPV����_12
        .Fields("WFSPV_MESDATA13").Value = " "      'WFSPV����_13
        .Fields("WFSPV_MESDATA14").Value = " "      'WFSPV����_14
        .Fields("WFSPV_MESDATA15").Value = " "      'WFSPV����_15
            
'        .Fields("WFSPV_SMPPOS2").Value = -1         'WFSPV�����-ID����ʒu(SXL�ʒu���)
'        .Fields("WFSPV_NETSU2").Value = " "         'WFSPV_�M��������
'        .Fields("WFSPV_ET2").Value = " "            'WFSPV_�G�b�`���O����
'        .Fields("WFSPV_MES2").Value = " "           'WFSPV_�v�����@
'        .Fields("WFSPV_DKAN2").Value = " "          'WFSPV_�c�j�A�j�[������
'
'        .Fields("WFSPV_FE_MAX").Value = " "         'WFSPV_Fe_MAX
'        .Fields("WFSPV_FE_AVE").Value = " "         'WFSPV_Fe_AVE
'        .Fields("WFSPV_FE_MIN").Value = " "         'WFSPV_Fe_MIN
'        .Fields("WFSPV_F_MESDATA1").Value = " "     'WFSPV����_1   SPV_Fe
'        .Fields("WFSPV_F_MESDATA2").Value = " "     'WFSPV����_2   SPV_Fe
'        .Fields("WFSPV_F_MESDATA3").Value = " "     'WFSPV����_3   SPV_Fe
'        .Fields("WFSPV_F_MESDATA4").Value = " "     'WFSPV����_4   SPV_Fe
'        .Fields("WFSPV_F_MESDATA5").Value = " "     'WFSPV����_5   SPV_Fe
'        .Fields("WFSPV_F_MESDATA6").Value = " "     'WFSPV����_6   SPV_Fe
'        .Fields("WFSPV_F_MESDATA7").Value = " "     'WFSPV����_7   SPV_Fe
'        .Fields("WFSPV_F_MESDATA8").Value = " "     'WFSPV����_8   SPV_Fe
'        .Fields("WFSPV_F_MESDATA9").Value = " "     'WFSPV����_9   SPV_Fe
'
'        .Fields("WFSPV_DIFF_MAX").Value = " "       'WFSPV_�g�U��_MAX
'        .Fields("WFSPV_DIFF_AVE").Value = " "       'WFSPV_�g�U��_AVE
'        .Fields("WFSPV_DIFF_MIN").Value = " "       'WFSPV_�g�U��_MIN
'        .Fields("WFSPV_D_MESDATA1").Value = " "     'WFSPV����_1   SPV_�g�U��
'        .Fields("WFSPV_D_MESDATA2").Value = " "     'WFSPV����_2   SPV_�g�U��
'        .Fields("WFSPV_D_MESDATA3").Value = " "     'WFSPV����_3   SPV_�g�U��
'        .Fields("WFSPV_D_MESDATA4").Value = " "     'WFSPV����_4   SPV_�g�U��
'        .Fields("WFSPV_D_MESDATA5").Value = " "     'WFSPV����_5   SPV_�g�U��
'        .Fields("WFSPV_D_MESDATA6").Value = " "     'WFSPV����_6   SPV_�g�U��
'        .Fields("WFSPV_D_MESDATA7").Value = " "     'WFSPV����_7   SPV_�g�U��
'        .Fields("WFSPV_D_MESDATA8").Value = " "     'WFSPV����_8   SPV_�g�U��
'        .Fields("WFSPV_D_MESDATA9").Value = " "     'WFSPV����_9   SPV_�g�U��
    
''        ''>>=====SPV����@20060529 SMP����
''        .Fields("WFSPV_FE_PUA").Value = -1           ''SPV_Fe PUA�l
''        .Fields("WFSPV_FE_PUAP").Value = -1          ''SPV_Fe PUA���l
''        .Fields("WFSPV_FE_STD").Value = -1           ''SPV_Fe STD
''        .Fields("WFSPV_DIFF_PUA").Value = -1         ''SPV_�g�U�� PUA�l
''        .Fields("WFSPV_DIFF_PUAP").Value = -1        ''SPV_�g�U�� PUA���l
''        .Fields("WFSPV_NR_MAX").Value = -1           ''SPV_OtherRecords_MAX
''        .Fields("WFSPV_NR_AVE").Value = -1           ''SPV_OtherRecords_AVE
''        .Fields("WFSPV_NR_STD").Value = -1           ''SPV_OtherRecords_STD
''        .Fields("WFSPV_NR_PUA").Value = -1           ''SPV_OtherRecords_PUA�l
''        .Fields("WFSPV_NR_PUAP").Value = -1          ''SPV_OtherRecords_PUA���l
''        ''==============================<<
    End With
    
    If (recXSDCW("WFINDSPCW").Value <> "0") And (recXSDCW("WFRESSPCW").Value <> "0") Then
        
    '-------------------- TBCMJ016�̓ǂݍ���(WFSPV) ----------------------------------------
        sql = ""
        sql = sql & " select *"
        sql = sql & " from   tbcmj016 " & vbLf
        sql = sql & " where  crynum = '" & CRYNUM & "'" & vbLf
        sql = sql & " and    smplno = '" & recXSDCW("WFSMPLIDSPCW").Value & "'" & vbLf
        sql = sql & " and    hsflg = '1'" & vbLf
        sql = sql & " and    trancnt = ( select   max(trancnt) from tbcmj016 " & vbLf
        sql = sql & "                    where    crynum = '" & CRYNUM & "'" & vbLf
        sql = sql & "                    and      smplno = '" & recXSDCW("WFSMPLIDSPCW").Value & "'" & vbLf
        sql = sql & "                    and      hsflg = '1'" & vbLf
        sql = sql & "                   )" & vbLf
                
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            GoTo proc_exit
        End If
                
    '-------------------- TBCME028�̓ǂݍ���(SPV�d�l) ----------------------------------------
        sql = ""
        sql = sql & " select HWFSPVSH,HWFSPVST,HWFSPVSI,HWFDLSPH,HWFDLSPT,HWFDLSPI"
        sql = sql & " from   TBCME028"
        sql = sql & " where  HINBAN = '" & HIN.hinban & "'"
        sql = sql & " and    MNOREVNO = " & HIN.mnorevno
        sql = sql & " and    FACTORY = '" & HIN.factory & "'"
        sql = sql & " and    OPECOND = '" & HIN.opecond & "'"
        
        Set rs2 = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        
        If rs2.RecordCount = 0 Then
            GoTo proc_exit
        End If
                
        'TBCMX001
        With recX001
            If Not IsNull(recXSDCW("INPOSCW").Value) Then .Fields("WFSPV_SMPPOS").Value = recXSDCW("INPOSCW").Value     'WFSPV�����-ID����ʒu(SXL�ʒu���)
            If Not IsNull(rs("NETSU")) Then .Fields("WFSPV_NETSU").Value = rs("NETSU")                                  'WFSPV_�M��������
            If Not IsNull(rs("ET")) Then .Fields("WFSPV_ET").Value = rs("ET")                                           'WFSPV_�G�b�`���O����
            If Not IsNull(rs("MES")) Then .Fields("WFSPV_MES").Value = rs("MES")                                        'WFSPV_�v�����@
            
            'MAP����̏ꍇ
            If rs2.Fields("HWFSPVSH") & rs2.Fields("HWFSPVST") & rs2.Fields("HWFSPVSI") = "AMX" Then
                If Not IsNull(rs("SPV_FE_MAX")) Then .Fields("WFSPV_FE_MAX").Value = rs("SPV_FE_MAX")                   'WFSPV_Fe�Z�x���莞��MAX�l
                If Not IsNull(rs("SPV_FE_AVE")) Then .Fields("WFSPV_FE_AVE").Value = rs("SPV_FE_AVE")                   'WFSPV_Fe�Z�x���莞��AVE�l
                If Not IsNull(rs("SPV_FE_MIN")) Then .Fields("WFSPV_FE_MIN").Value = rs("SPV_FE_MIN")                   'WFSPV_Fe�Z�x���莞��MIN�l
                
            '9�_����̏ꍇ
            ElseIf rs2.Fields("HWFSPVSH") & rs2.Fields("HWFSPVST") & rs2.Fields("HWFSPVSI") = "V9T" Then
            
                ''Fe�Z�x��MAX,MIN,AVE���擾
                If getTBCMJ016WFSPV_Fe(CRYNUM, rs("SMPLNO"), rs("TRANCNT"), _
                                         recX001) = FUNCTION_RETURN_FAILURE Then
                    GoTo proc_exit
                End If
                
            '������ł��Ȃ��ꍇ�͊g�U���̑�����@�Ŕ��f
            ElseIf rs2.Fields("HWFDLSPH") & rs2.Fields("HWFDLSPT") & rs2.Fields("HWFDLSPI") = "AMX" Then
                If Not IsNull(rs("SPV_FE_MAX")) Then .Fields("WFSPV_FE_MAX").Value = rs("SPV_FE_MAX")                   'WFSPV_Fe�Z�x���莞��MAX�l
                If Not IsNull(rs("SPV_FE_AVE")) Then .Fields("WFSPV_FE_AVE").Value = rs("SPV_FE_AVE")                   'WFSPV_Fe�Z�x���莞��AVE�l
                If Not IsNull(rs("SPV_FE_MIN")) Then .Fields("WFSPV_FE_MIN").Value = rs("SPV_FE_MIN")                   'WFSPV_Fe�Z�x���莞��MIN�l
            
            ElseIf rs2.Fields("HWFDLSPH") & rs2.Fields("HWFDLSPT") & rs2.Fields("HWFDLSPI") = "V9T" And _
                    IsNull(rs("MS01_SPV_FE")) = False Then
                
                ''Fe�Z�x��MAX,MIN,AVE���擾
                If getTBCMJ016WFSPV_Fe(CRYNUM, rs("SMPLNO"), rs("TRANCNT"), _
                                         recX001) = FUNCTION_RETURN_FAILURE Then
                    GoTo proc_exit
                End If
                
            End If
            
            'MAP����̏ꍇ
            If rs2.Fields("HWFDLSPH") & rs2.Fields("HWFDLSPT") & rs2.Fields("HWFDLSPI") = "AMX" Then
                If Not IsNull(rs("SPV_DIFF_MAX")) Then .Fields("WFSPV_KST_MAX").Value = rs("SPV_DIFF_MAX")              'WFSPV_�g�U�����莞��MAX�l
                If Not IsNull(rs("SPV_DIFF_AVE")) Then .Fields("WFSPV_KST_AVE").Value = rs("SPV_DIFF_AVE")              'WFSPV_�g�U�����莞��AVE�l
                If Not IsNull(rs("SPV_DIFF_MIN")) Then .Fields("WFSPV_KST_MIN").Value = rs("SPV_DIFF_MIN")              'WFSPV_�g�U�����莞��MIN�l
            
            '9�_����̏ꍇ
            ElseIf rs2.Fields("HWFDLSPH") & rs2.Fields("HWFDLSPT") & rs2.Fields("HWFDLSPI") = "V9T" Then
                
                ''�g�U����MAX,MIN,AVE���擾
                If getTBCMJ016WFSPV_Diff(CRYNUM, rs("SMPLNO"), rs("TRANCNT"), _
                                         recX001) = FUNCTION_RETURN_FAILURE Then
                    GoTo proc_exit
                End If
            
            '������ł��Ȃ��ꍇ��Fe�Z�x�̑�����@�Ŕ��f
            ElseIf rs2.Fields("HWFSPVSH") & rs2.Fields("HWFSPVST") & rs2.Fields("HWFSPVSI") = "AMX" Then
                If Not IsNull(rs("SPV_DIFF_MAX")) Then .Fields("WFSPV_KST_MAX").Value = rs("SPV_DIFF_MAX")              'WFSPV_�g�U�����莞��MAX�l
                If Not IsNull(rs("SPV_DIFF_AVE")) Then .Fields("WFSPV_KST_AVE").Value = rs("SPV_DIFF_AVE")              'WFSPV_�g�U�����莞��AVE�l
                If Not IsNull(rs("SPV_DIFF_MIN")) Then .Fields("WFSPV_KST_MIN").Value = rs("SPV_DIFF_MIN")              'WFSPV_�g�U�����莞��MIN�l
            
            ElseIf rs2.Fields("HWFSPVSH") & rs2.Fields("HWFSPVST") & rs2.Fields("HWFSPVSI") = "V9T" And _
                    IsNull(rs("MS01_SPV_DIFF")) = False Then
                
                ''�g�U����MAX,MIN,AVE���擾
                If getTBCMJ016WFSPV_Diff(CRYNUM, rs("SMPLNO"), rs("TRANCNT"), _
                                         recX001) = FUNCTION_RETURN_FAILURE Then
                    GoTo proc_exit
                End If
                
            End If
            ''>>>===SPV����@20060529 SMP���� ==
            If Not IsNull(rs("SPV_Fe_PUA")) Then .Fields("WFSPV_FE_PUA").Value = rs("SPV_Fe_PUA")            ''SPV_Fe PUA�l
            If Not IsNull(rs("SPV_Fe_PUAP")) Then .Fields("WFSPV_FE_PUAP").Value = rs("SPV_Fe_PUAP")         ''SPV_Fe PUA���l
            If Not IsNull(rs("SPV_Fe_STD")) Then .Fields("WFSPV_FE_STD").Value = rs("SPV_Fe_STD")          ''SPV_Fe STD
            If Not IsNull(rs("SPV_Diff_PUA")) Then .Fields("WFSPV_DIFF_PUA").Value = rs("SPV_Diff_PUA")      ''SPV_�g�U�� PUA�l
            If Not IsNull(rs("SPV_Diff_PUAP")) Then .Fields("WFSPV_DIFF_PUAP").Value = rs("SPV_Diff_PUAP")     ''SPV_�g�U�� PUA���l
            If Not IsNull(rs("SPV_Nr_MAX")) Then .Fields("WFSPV_NR_MAX").Value = rs("SPV_Nr_MAX")           ''SPV_OtherRecords_MAX
            If Not IsNull(rs("SPV_Nr_AVE")) Then .Fields("WFSPV_NR_AVE").Value = rs("SPV_Nr_AVE")           ''SPV_OtherRecords_AVE
            If Not IsNull(rs("SPV_Nr_STD")) Then .Fields("WFSPV_NR_STD").Value = rs("SPV_Nr_STD")          ''SPV_OtherRecords_STD
            If Not IsNull(rs("SPV_Nr_PUA")) Then .Fields("WFSPV_NR_PUA").Value = rs("SPV_Nr_PUA")          ''SPV_OtherRecords_PUA�l
            If Not IsNull(rs("SPV_Nr_PUAP")) Then .Fields("WFSPV_NR_PUAP").Value = rs("SPV_Nr_PUAP")         ''SPV_OtherRecords_PUA���l
            ''==================================<<
        End With
            
        'TBCMX002
        With recX002
            
            If Not IsNull(recXSDCW("INPOSCW").Value) Then .Fields("WFSPV_SMPPOS").Value = recXSDCW("INPOSCW").Value     'WFSPV�����-ID����ʒu(SXL�ʒu���)
            If Not IsNull(rs("NETSU")) Then .Fields("WFSPV_NETSU").Value = rs("NETSU")                                  'WFSPV_�M��������
            If Not IsNull(rs("ET")) Then .Fields("WFSPV_ET").Value = rs("ET")                                           'WFSPV_�G�b�`���O����
            If Not IsNull(rs("MES")) Then .Fields("WFSPV_MES").Value = rs("MES")                                        'WFSPV_�v�����@
            If Not IsNull(rs("DKAN")) Then .Fields("WFSPV_DKAN").Value = rs("DKAN")                                     'WFSPV_�c�j�A�j�[������

            .Fields("WFSPV_SMPPOS2").Value = recXSDCW("INPOSCW").Value          'WFSPV�����-ID����ʒu(SXL�ʒu���)
            .Fields("WFSPV_NETSU2").Value = rs("NETSU")                         'WFSPV_�M��������
            .Fields("WFSPV_ET2").Value = rs("ET")                               'WFSPV_�G�b�`���O����
            .Fields("WFSPV_MES2").Value = rs("MES")                             'WFSPV_�v�����@
            .Fields("WFSPV_DKAN2").Value = rs("DKAN")                           'WFSPV_�c�j�A�j�[������

            .Fields("WFSPV_FE_MAX").Value = rs("SPV_FE_MAX")                    'WFSPV_Fe_MAX
            .Fields("WFSPV_FE_AVE").Value = rs("SPV_FE_AVE")                    'WFSPV_Fe_AVE
            .Fields("WFSPV_FE_MIN").Value = rs("SPV_FE_MIN")                    'WFSPV_Fe_MIN
            .Fields("WFSPV_F_MESDATA1").Value = rs("MS01_SPV_FE")               'WFSPV����_1   SPV_Fe
            .Fields("WFSPV_F_MESDATA2").Value = rs("MS02_SPV_FE")               'WFSPV����_2   SPV_Fe
            .Fields("WFSPV_F_MESDATA3").Value = rs("MS03_SPV_FE")               'WFSPV����_3   SPV_Fe
            .Fields("WFSPV_F_MESDATA4").Value = rs("MS04_SPV_FE")               'WFSPV����_4   SPV_Fe
            .Fields("WFSPV_F_MESDATA5").Value = rs("MS05_SPV_FE")               'WFSPV����_5   SPV_Fe
            .Fields("WFSPV_F_MESDATA6").Value = rs("MS06_SPV_FE")               'WFSPV����_6   SPV_Fe
            .Fields("WFSPV_F_MESDATA7").Value = rs("MS07_SPV_FE")               'WFSPV����_7   SPV_Fe
            .Fields("WFSPV_F_MESDATA8").Value = rs("MS08_SPV_FE")               'WFSPV����_8   SPV_Fe
            .Fields("WFSPV_F_MESDATA9").Value = rs("MS09_SPV_FE")               'WFSPV����_9   SPV_Fe
            .Fields("WFSPV_DIFF_MAX").Value = rs("SPV_DIFF_MAX")                'WFSPV_�g�U��_MAX
            .Fields("WFSPV_DIFF_AVE").Value = rs("SPV_DIFF_AVE")                'WFSPV_�g�U��_AVE
            .Fields("WFSPV_DIFF_MIN").Value = rs("SPV_DIFF_MIN")                'WFSPV_�g�U��_MIN
            .Fields("WFSPV_D_MESDATA1").Value = rs("MS01_SPV_DIFF")             'WFSPV����_1   SPV_�g�U��
            .Fields("WFSPV_D_MESDATA2").Value = rs("MS02_SPV_DIFF")             'WFSPV����_2   SPV_�g�U��
            .Fields("WFSPV_D_MESDATA3").Value = rs("MS03_SPV_DIFF")             'WFSPV����_3   SPV_�g�U��
            .Fields("WFSPV_D_MESDATA4").Value = rs("MS04_SPV_DIFF")             'WFSPV����_4   SPV_�g�U��
            .Fields("WFSPV_D_MESDATA5").Value = rs("MS05_SPV_DIFF")             'WFSPV����_5   SPV_�g�U��
            .Fields("WFSPV_D_MESDATA6").Value = rs("MS06_SPV_DIFF")             'WFSPV����_6   SPV_�g�U��
            .Fields("WFSPV_D_MESDATA7").Value = rs("MS07_SPV_DIFF")             'WFSPV����_7   SPV_�g�U��
            .Fields("WFSPV_D_MESDATA8").Value = rs("MS08_SPV_DIFF")             'WFSPV����_8   SPV_�g�U��
            .Fields("WFSPV_D_MESDATA9").Value = rs("MS09_SPV_DIFF")             'WFSPV����_9   SPV_�g�U��
            
            
''            ''>>>===SPV����@20060529 SMP���� ==
''            If Not IsNull(rs("SPV_Fe_PUA")) Then .Fields("WFSPV_FE_PUA").Value = rs("SPV_Fe_PUA")            ''SPV_Fe PUA�l
''            If Not IsNull(rs("SPV_Fe_PUAP")) Then .Fields("WFSPV_FE_PUAP").Value = rs("SPV_Fe_PUAP")         ''SPV_Fe PUA���l
''            If Not IsNull(rs("SPV_Fe_STD")) Then .Fields("WFSPV_FE_STD").Value = rs("SPV_Fe_STD")          ''SPV_Fe STD
''            If Not IsNull(rs("SPV_Diff_PUA")) Then .Fields("WFSPV_DIFF_PUA").Value = rs("SPV_Diff_PUA")      ''SPV_�g�U�� PUA�l
''            If Not IsNull(rs("SPV_Diff_PUAP")) Then .Fields("WFSPV_DIFF_PUAP").Value = rs("SPV_Diff_PUAP")     ''SPV_�g�U�� PUA���l
''            If Not IsNull(rs("SPV_Nr_MAX")) Then .Fields("WFSPV_NR_MAX").Value = rs("SPV_Nr_MAX")           ''SPV_OtherRecords_MAX
''            If Not IsNull(rs("SPV_Nr_AVE")) Then .Fields("WFSPV_NR_AVE").Value = rs("SPV_Nr_AVE")           ''SPV_OtherRecords_AVE
''            If Not IsNull(rs("SPV_Nr_STD")) Then .Fields("WFSPV_NR_STD").Value = rs("SPV_Nr_STD")          ''SPV_OtherRecords_STD
''            If Not IsNull(rs("SPV_Nr_PUA")) Then .Fields("WFSPV_NR_PUA").Value = rs("SPV_Nr_PUA")          ''SPV_OtherRecords_PUA�l
''            If Not IsNull(rs("SPV_Nr_PUAP")) Then .Fields("WFSPV_NR_PUAP").Value = rs("SPV_Nr_PUAP")         ''SPV_OtherRecords_PUA���l
''            ''==================================<<

        End With
        
        Set rs = Nothing
        Set rs2 = Nothing
    
    End If

    getTBCMJ016WFSPV = FUNCTION_RETURN_SUCCESS

proc_exit:
    
    Set rs = Nothing
    Set rs2 = Nothing
    
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    
    Set rs = Nothing
    Set rs2 = Nothing
    
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ016WFSPV = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :WFSPV����(TBCMJ016) Fe�Z�x��MAX/AVE/MIN�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :sSmplID         , I  ,String            , �����ID
'          :iTranCnt        , I  ,Integer           , ������
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :WFSPV����(TBCMJ016)����Fe�Z�x��MAX/AVE/MIN�ް����擾���ASXL�������\���̂ɾ�Ă���
'����      :2005/06/23  �V�K�쐬�@(TCS)t.terauchi
Private Function getTBCMJ016WFSPV_Fe(CRYNUM As String, sSmplID As String, iTrancnt As Integer, _
                                    recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ016WFSPV_Fe"
    
    getTBCMJ016WFSPV_Fe = FUNCTION_RETURN_FAILURE
                    
    '-------------------- TBCMJ016�̓ǂݍ���(WFSPV) ----------------------------------------
        sql = ""
        sql = sql & " SELECT  MAX(SPV_FE) AS MAX_FE,MIN(SPV_FE) AS MIN_FE,AVG(SPV_FE) AS AVE_FE" & vbLf
        sql = sql & " FROM   (SELECT  CRYNUM,SMPLNO,TRANCNT,ms01_SPV_FE AS SPV_FE" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms02_SPV_FE AS SPV_FE" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms03_SPV_FE AS SPV_FE" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms04_SPV_FE AS SPV_FE" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms05_SPV_FE AS SPV_FE" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms06_SPV_FE AS SPV_FE" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms07_SPV_FE AS SPV_FE" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms08_SPV_FE AS SPV_FE" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms09_SPV_FE AS SPV_FE" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "        )" & vbLf
                
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            If Not IsNull(rs("MAX_FE")) Then .Fields("WFSPV_FE_MAX").Value = rs("MAX_FE")   'WFSPV_Fe�Z�x���莞��MAX�l
            If Not IsNull(rs("MAX_FE")) Then .Fields("WFSPV_FE_AVE").Value = rs("AVE_FE")   'WFSPV_Fe�Z�x���莞��AVE�l
            If Not IsNull(rs("MAX_FE")) Then .Fields("WFSPV_FE_MIN").Value = rs("MIN_FE")   'WFSPV_Fe�Z�x���莞��MIN�l
        End With
        Set rs = Nothing

    getTBCMJ016WFSPV_Fe = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ016WFSPV_Fe = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :WFSPV����(TBCMJ016) �g�U����MAX/AVE/MIN�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :sSmplID         , I  ,String            , �����ID
'          :iTranCnt        , I  ,Integer           , ������
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :WFSPV����(TBCMJ016)����g�U����MAX/AVE/MIN�ް����擾���ASXL�������\���̂ɾ�Ă���
'����      :2005/06/23  �V�K�쐬�@(TCS)t.terauchi
Private Function getTBCMJ016WFSPV_Diff(CRYNUM As String, sSmplID As String, iTrancnt As Integer, _
                                    recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ016WFSPV_Diff"
    
    getTBCMJ016WFSPV_Diff = FUNCTION_RETURN_FAILURE
                    
    '-------------------- TBCMJ016�̓ǂݍ���(WFSPV) ----------------------------------------
        sql = ""
        sql = sql & " SELECT  MAX(SPV_DIFF) AS MAX_DIFF,MIN(SPV_DIFF) AS MIN_DIFF,AVG(SPV_DIFF) AS AVE_DIFF" & vbLf
        sql = sql & " FROM   (SELECT  CRYNUM,SMPLNO,TRANCNT,ms01_SPV_DIFF AS SPV_DIFF" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms02_SPV_DIFF AS SPV_DIFF" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms03_SPV_DIFF AS SPV_DIFF" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms04_SPV_DIFF AS SPV_DIFF" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms05_SPV_DIFF AS SPV_DIFF" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms06_SPV_DIFF AS SPV_DIFF" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms07_SPV_DIFF AS SPV_DIFF" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms08_SPV_DIFF AS SPV_DIFF" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms09_SPV_DIFF AS SPV_DIFF" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "        )" & vbLf
               
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            If Not IsNull(rs("MAX_DIFF")) Then .Fields("WFSPV_KST_MAX").Value = rs("MAX_DIFF")  'WFSPV_�g�U�����莞��MAX�l
            If Not IsNull(rs("AVE_DIFF")) Then .Fields("WFSPV_KST_AVE").Value = rs("AVE_DIFF")  'WFSPV_�g�U�����莞��AVE�l
            If Not IsNull(rs("MIN_DIFF")) Then .Fields("WFSPV_KST_MIN").Value = rs("MIN_DIFF")  'WFSPV_�g�U�����莞��MIN�l
        End With
        Set rs = Nothing

    getTBCMJ016WFSPV_Diff = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ016WFSPV_Diff = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function



'�T�v      :�W������(TBCMY018)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :sBlkId          , I  ,String            , �������ۯ�ID
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW�\����   (�V����يǗ�(SXL))
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SXL������)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :�W������(TBCMY018)����Warp���т��擾���ASXL�������E����_�ް��\���̂ɾ�Ă���
'����      :2005/06/23  �V�K�쐬  (TCS)T.terauchi
Private Function getTBCMY018WARP(sBlkId As String, recXSDCW As c_cmzcrec, recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY018WARP"
    
    getTBCMY018WARP = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    'TBCMX001
'    With recX001
'        .Fields("WARP_1").Value = vbNullString            '����Warp-1
'        .Fields("WARP_2").Value = vbNullString            '����Warp-2
'        .Fields("WARP_3").Value = vbNullString            '����Warp-3
'    End With
                    
    '-------------------- TBCMY018�̓ǂݍ��� ----------------------------------------
    sql = ""
    sql = sql & " select max(to_number(measdata)) as warp_1" & vbLf
    sql = sql & "        ,avg(to_number(measdata)) as warp_2" & vbLf
    sql = sql & "        ,min(to_number(measdata)) as warp_3" & vbLf
    sql = sql & " from   tbcmy018" & vbLf
    sql = sql & " where  sublotid = '" & sBlkId & "'" & vbLf
    sql = sql & " and    measitem = 'MSL04WARPU'" & vbLf
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        getTBCMY018WARP = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    
    'TBCMX001
    With recX001
        .Fields("WARP_1").Value = rs("warp_1")          '����Warp-1(�ő�l)
        .Fields("WARP_2").Value = rs("warp_2")          '����Warp-2(����)
        .Fields("WARP_3").Value = rs("warp_3")          '����Warp-3(�ŏ��l)
    End With
        
    Set rs = Nothing
    
    getTBCMY018WARP = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY018WARP = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
''Upd end   2005/06/23 (TCS)T.terauchi      SPV9�_�Ή�

'�T�v      :SXLID�𷰂�XSDCA������ۯ�ID���擾����
'���Ұ�    :�ϐ���        ,IO ,�^                       :����
'          :sSXLID        ,I  ,String                   :SXLID
'          :�߂�l        ,O  ,FUNCTION_RETURN          :���o�̐���
'����      :
'����      :06/01/20 ooba
Public Function GetCaBlockID(sSXLID As String) As FUNCTION_RETURN

    Dim i, m        As Integer
    Dim sql         As String           'SQL�S��
    Dim rs          As OraDynaset       'RecordSet

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "f_cmec060_1.frm -- Function GetCaBlockID"
    
    
    sql = "select CRYNUMCA from XSDCA "
    sql = sql & "where SXLIDCA = '" & sSXLID & "' "
    sql = sql & "group by CRYNUMCA "
    sql = sql & "order by CRYNUMCA"
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount < 1 Then
        GetCaBlockID = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    m = rs.RecordCount
    ReDim sReBlkID(m)
    ''���o���ʂ��i�[����
    For i = 1 To m
        sReBlkID(i) = rs("CRYNUMCA")        '��ۯ�ID
        rs.MoveNext
    Next i
        
    Set rs = Nothing
    
    GetCaBlockID = FUNCTION_RETURN_SUCCESS
  

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    GetCaBlockID = FUNCTION_RETURN_FAILURE
    Resume proc_exit
    
End Function

'����EDI����ݸ�Ή� 2009/12/4 Add Start SPK habuki�@������
'---------------------------------------------------------------------------
'�T�v      :SXLID��V�����AXODC6_1���������A���M�Ώۂ�EDI����Ԃ�
'---------------------------------------------------------------------------
'���Ұ�    :�ϐ���      ,IO     ,�^                     ,����
'          :pSXLID      ,I  �@�@,String                 ,SXLID
'          :pPN         ,O  �@�@,Double                 ,�s�����Z�x(P:��)
'          :pBN         ,O  �@�@,Double                 ,�s�����Z�x(B:����)
'          :pASN        ,O  �@�@,Double                 ,�s�����Z�x(AS:��f)
'          :pCN         ,O  �@�@,Double                 ,�s�����Z�x(C:�Y�f)
'          :pflgEDI     ,O  �@�@,Boolean                ,EDI���L������p(True:�L�AFalse�F��)
'          :�߂�l      ,O      ,Boolean                ,[True:OK�^False:NG]
'---------------------------------------------------------------------------
Public Function fncGetEdiInfo(ByVal pSXLID As String, _
                              ByRef pPN As Double, _
                              ByRef pBN As Double, _
                              ByRef pASN As Double, _
                              ByRef pCN As Double, _
                              ByRef pflgEDI As Boolean _
                             ) As Boolean
    Dim i, m        As Integer
    Dim sql         As String           'SQL�S��
    Dim rs          As OraDynaset       'RecordSet
    
    '--�װ����ׂ̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function fncGetEdiInfo"
    
    '--������
    pPN = 0: pBN = 0: pASN = 0: pCN = 0
    fncGetEdiInfo = False: pflgEDI = False
    Set rs = Nothing      'Oracle RecordSet Free

    '--SQL������
    '<�����d�ʗD���>
''    sql = "select" & vbCrLf
''    sql = sql & "    c61.PNC6         PN" & vbCrLf                      '�s�����Z�x(P:��)
''    sql = sql & "   ,c61.BNC6         BN" & vbCrLf                      '�s�����Z�x(B:����)
''    sql = sql & "   ,c61.ASNC6        ASN" & vbCrLf                     '�s�����Z�x(AS:��f)
''    sql = sql & "   ,c61.CNC6         CN" & vbCrLf                      '�s�����Z�x(C:�Y�f)
''    sql = sql & "from" & vbCrLf
''    sql = sql & "   (" & vbCrLf
''    sql = sql & "     select" & vbCrLf
''    sql = sql & "         max(MATERYOC6)    MATERYOC6" & vbCrLf         '�d�ʁ^����
''    sql = sql & "     from" & vbCrLf
''    sql = sql & "         XODC6_1" & vbCrLf                             '���ޗ�
''    sql = sql & "     where" & vbCrLf
''    sql = sql & "          substr(XTALC6,1,7) = '" & left(pSXLID, 7) & "'" & vbCrLf    '���㌋���ԍ�=SXLID(7��)
''    sql = sql & "      and MATEKC6  = '1'" & vbCrLf                     '�敪(��ؼغ�)
''    sql = sql & "      and EDIFLGC6 = '2'" & vbCrLf                     'EDI�׸�(���M�Ώ�)
''    sql = sql & "   ) tkey" & vbCrLf
''    sql = sql & "   ,XODC6_1  c61" & vbCrLf                             '���ޗ�
''    sql = sql & "   ,TBCMH001 t01" & vbCrLf                             '����w������
''    sql = sql & "where" & vbCrLf
''    sql = sql & "     substr(c61.XTALC6,1,7) = '" & left(pSXLID, 7) & "'" & vbCrLf      '���㌋���ԍ�=SXLID(7��)
''    sql = sql & " and c61.MATEKC6   = '1'" & vbCrLf                     '�敪(��ؼغ�)
''    sql = sql & " and c61.EDIFLGC6  = '2'" & vbCrLf                     'EDI�׸�(���M�Ώ�)
''    sql = sql & " and c61.MATERYOC6 = tkey.MATERYOC6" & vbCrLf          '�d�ʁ^����
''    sql = sql & " and substr(c61.XTALC6,1,7)||substr(c61.XTALC6,9,1) = substr(t01.UPINDNO,1,7)||substr(t01.UPINDNO,9,1)" & vbCrLf          '���㌋���ԍ�7���{9����
''    sql = sql & " and t01.CODE > '4'" & vbCrLf                          '�d�|����
''    sql = sql & " and rownum = 1" & vbCrLf                              '1ں��ޖ�
''    sql = sql & "order by" & vbCrLf
''    sql = sql & "    substr(c61.XTALC6,9,1)" & vbCrLf                   '���㌋���ԍ�9����(0:�ʏ�A1�`4�F�����ށAA�`C�FAB���)
''    sql = sql & "   ,c61.MATESYUC6" & vbCrLf                            '���(��������)

    '<�����ԍ��D���>
    sql = "select" & vbCrLf
    sql = sql & "    c61.PNC6         PN" & vbCrLf                      '�s�����Z�x(P:��)
    sql = sql & "   ,c61.BNC6         BN" & vbCrLf                      '�s�����Z�x(B:����)
    sql = sql & "   ,c61.ASNC6        ASN" & vbCrLf                     '�s�����Z�x(AS:��f)
    sql = sql & "   ,c61.CNC6         CN" & vbCrLf                      '�s�����Z�x(C:�Y�f)
    sql = sql & "from" & vbCrLf
    sql = sql & "    XODC6_1  c61" & vbCrLf                             '���ޗ�
    sql = sql & "   ,TBCMH001 t01" & vbCrLf                             '����w������
    sql = sql & "where" & vbCrLf
    sql = sql & "     substr(c61.XTALC6,1,7) = '" & left(pSXLID, 7) & "'" & vbCrLf      '���㌋���ԍ�=SXLID(7��)
    sql = sql & " and c61.MATEKC6   = '1'" & vbCrLf                     '�敪(��ؼغ�)
    sql = sql & " and c61.EDIFLGC6  = '2'" & vbCrLf                     'EDI�׸�(���M�Ώ�)
    sql = sql & " and substr(c61.XTALC6,1,7)||substr(c61.XTALC6,9,1) = substr(t01.UPINDNO,1,7)||substr(t01.UPINDNO,9,1)" & vbCrLf          '���㌋���ԍ�7���{9����
    sql = sql & " and t01.CODE > '4'" & vbCrLf                          '�d�|����
    sql = sql & " and rownum = 1" & vbCrLf                              '1ں��ޖ�
    sql = sql & "order by" & vbCrLf
    sql = sql & "    substr(c61.XTALC6,9,1)" & vbCrLf                   '���㌋���ԍ�9����(0:�ʏ�A1�`4�F�����ށAA�`C�FAB���)
    sql = sql & "   ,c61.MATESYUC6" & vbCrLf                            '���(��������)

    '--�ް��𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY Or ORADYN_NOCACHE)
    If rs Is Nothing Then
        GoTo proc_exit
    End If
    
    '--���o���ʎQ��
    If rs.EOF Then
        '<< �ް����� >>
        fncGetEdiInfo = True
        GoTo proc_exit
    Else
        '<< �ް��L�� >>
        rs.MoveFirst
        pPN = IIf(IsNull(rs("PN")), 0, rs("PN"))                '�s�����Z�x(P:��)
        pBN = IIf(IsNull(rs("BN")), 0, rs("BN"))                '�s�����Z�x(B:����)
        pASN = IIf(IsNull(rs("ASN")), 0, rs("ASN"))             '�s�����Z�x(AS:��f)
        pCN = IIf(IsNull(rs("CN")), 0, rs("CN"))                '�s�����Z�x(C:�Y�f)
        pflgEDI = True                                          'EDI���L������p(True:�L�AFalse�F��)
    End If
    
    fncGetEdiInfo = True

proc_exit:
    '<< �I�� >>
    'Oracle RecordSet Free
    If Not (rs Is Nothing) Then
        rs.Close
    End If
    Set rs = Nothing
    
    gErr.Pop
    Exit Function

proc_err:
    '<< �װ����� >>
    'Oracle RecordSet Free
    If Not (rs Is Nothing) Then
        rs.Close
    End If
    Set rs = Nothing

    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    
    gErr.HandleError
    Resume proc_exit
    
End Function
'����EDI����ݸ�Ή� 2009/12/4 Add End   SPK habuki�@������
    
'�T�v      :WFSIRD����(TBCMJ022)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW�\����   (�V����يǗ�(SXL))
'          :recX001         , O  ,c_cmzcrec         , TBCMX001�\����(SIRD�]�������ް�)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :SIRD�]������(TBCMJ022)�����ް����擾���A�\���̂ɾ�Ă���
'����      :2010/04/19 Y.Hitomi
Private Function getTBCMJ022SIRD(CRYNUM As String, recXSDCW As c_cmzcrec, recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ022SIRD"
    
    getTBCMJ022SIRD = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
        
    'TBCMX001
    With recX001
            .Fields("SIRD_POS").Value = vbNullString           'FSIRD�T���v������ʒu(SXL�ʒu���)
            .Fields("SIRD_TOTAL").Value = vbNullString         'WFSIRD_���莞��TOTAL�l
    End With
        
    '-------------------- TBCMJ022�̓ǂݍ���(SIRD) ----------------------------------------
    sql = "select * from TBCMJ022 "
    sql = sql & " where CRYNUM = '" & CRYNUM & "'"
'DEL 2010/05/20 Y.Hitomi
'    sql = sql & " and   SMPLNO = '" & recXSDCW("WFSMPLIDL4CW").Value & "'"
    sql = sql & " and   TRANCNT = 0"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        getTBCMJ022SIRD = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    
    'TBCMX001
    With recX001
        .Fields("SIRD_POS").Value = rs("POSITION")              'WFSIRD�T���v������ʒu(SXL�ʒu���)
        .Fields("SIRD_TOTAL").Value = rs("SIRDCNT")             'WFSIRD_���莞��TOTAL�l
        
    End With
    Set rs = Nothing

    getTBCMJ022SIRD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ022SIRD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'=================================================================================
' 2011/01/17 tkimura ADD START
'�T�v      :��Top�ʒu,��Bot�ʒu���擾����B
'���Ұ�    :�ϐ���          ,IO  ,�^                :����
'          :SXLID           ,I   ,String            ,SXLID OR �u���b�NID
'�@�@      :sPos  �@�@�@    ,I   ,String �@         ,SXL�ʒu(TOP/BOT)
'          :sPattern        ,I   ,String            ,���R�ް��擾�����
'                                                   �������A : WF�����ް��擾
'                                                   �������B : ���������ް��擾
'                                                   �������C : �擾�ް��Ȃ�
'          :sRsPos()        ,O   ,String            ,���R�ʒu(TOP/BOT)
'          :�߂�l          ,O   ,FUNCTION_RETURN   ,���o�̐���
'����      :
'����      :2011/01/17 tkimura �쐬
''Public Function cmbc040_GetSxlRsPos(ByVal data As String, _
''                                    ByVal sPos As String, _
''                                    ByVal sSmpId As String, _
''                                    ByVal sPattern As String, _
''                                    ByRef sRsPos() As String) As FUNCTION_RETURN
Public Function cmbc040_GetSxlRsPos(ByVal data As String, _
                                    ByVal sPos As String, _
                                    ByVal sPattern As String, _
                                    ByRef sRsPos() As String) As FUNCTION_RETURN
    
    Dim sTBkbn As String        'T/B�敪
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer           'Top�̂Ƃ�1,Bot�̂Ƃ�2��������B
    Dim sSql As String
    Dim rs As OraDynaset
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function cmbc040_GetSxlRsPos"
    cmbc040_GetSxlRsPos = FUNCTION_RETURN_FAILURE
    
    If sPos = "TOP" Then sTBkbn = "T" Else sTBkbn = "B"  '04/04/15 ooba
    
    '���R�ް��擾����݂��wA�x�̏ꍇ�AXSDCW[�V�T���v���Ǘ�(SXL)]���ʒu���擾����B
    If sPattern = "A" Then
        sSql = ""
        sSql = sSql & "SELECT" & vbCrLf
        sSql = sSql & " INPOSCW " & vbCrLf                      '�������ʒu
        sSql = sSql & "FROM" & vbCrLf
        sSql = sSql & " XSDCW " & vbCrLf
        sSql = sSql & "WHERE" & vbCrLf
        sSql = sSql & " SXLIDCW = '" & data & "' AND" & vbCrLf  'SXLID
        sSql = sSql & " TBKBNCW = '" & sTBkbn & "'" & vbCrLf    'TB�敪
        ''sSQL = sSQL & " REPSMPLIDCW = '" & sSmpId & "'"
        Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
        
        If rs.RecordCount > 0 Then
            'TOP���ʒu
            If sTBkbn = "T" Then
                k = 1
            'BOT���ʒu
            ElseIf sTBkbn = "B" Then
                k = 2
            End If
            sRsPos(k) = rs("INPOSCW")       '�������ʒu(TOP�܂���BOT)
        Else
            '�����ް����Ȃ��ꍇ�ʹװ
            Set rs = Nothing
            GoTo proc_exit
        End If
        Set rs = Nothing
    '���R�ް��擾����݂��wB�x�̏ꍇ�AXSDCS[�V�T���v���Ǘ�(�u���b�N)]���ʒu���擾����B
    '�T���v��ID���Ȃ��Ƃ��ɂ��̃p�^�[��������̂�sql���ɃT���v��ID�������ɓ���Ȃ��B
    ElseIf sPattern = "B" Then
        sSql = ""
        sSql = sSql & "SELECT" & vbCrLf
        sSql = sSql & " INPOSCS" & vbCrLf                        '�������ʒu
        sSql = sSql & "FROM" & vbCrLf
        sSql = sSql & " XSDCS" & vbCrLf
        sSql = sSql & "WHERE" & vbCrLf
        sSql = sSql & " CRYNUMCS = '" & data & "' AND" & vbCrLf  '�u���b�NID
        sSql = sSql & " TBKBNCS = '" & sTBkbn & "'" & vbCrLf     'TB�敪
        'sSQL = sSQL & " REPSMPLIDCS = '" & sSmpId & "'"
        
        Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
    
        If rs.RecordCount > 0 Then
            'TOP���ʒu
            If sTBkbn = "T" Then
                k = 1
            'BOT���ʒu
            ElseIf sTBkbn = "B" Then
                k = 2
            End If
            sRsPos(k) = rs("INPOSCS")       '�������ʒu(TOP�܂���BOT)
        Else
            '�����ް����Ȃ��ꍇ�ʹװ
            Set rs = Nothing
            GoTo proc_exit
        End If
        Set rs = Nothing
    '���R�ް��擾����݂��wC�x�̏ꍇ�A�擾�����ް��Ȃ��B
    ElseIf sPattern = "C" Then
    
    End If
    
'''    '�擾�ް�����/-1/NULL�̎��ͽ�߰���Ă���B
    If sRsPos(k) = "" Or sRsPos(k) = "-1" Or sRsPos(k) = vbNullString Then
        sRsPos(k) = " "
    End If
        
    cmbc040_GetSxlRsPos = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    cmbc040_GetSxlRsPos = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/17 tkimura ADD START
'�T�v      :�E�F�n�|�Z���^�|���ɏ��(TBCMY011)�̑��M�t���O���X�V����B
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :SXLID           , I  ,String            , �V���O��ID
'          :hin             , I  ,tFullHinban       , �i��(�S�i�ԍ\����)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :TBCME018.HSXRKHNM�̒l�𒲂ׁA���̒l��0�̏ꍇ��TBCMEY011.QA1SNDFLG�̒l��3�Ƃ���B����ȊO�̂Ƃ���2�Ƃ���B
'����      :2011/01/17 tkimura
Private Function UpdateTBCMY011SendFlag(ByVal SXLID As String, _
                                        ByRef HIN As tFullHinban) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim hindo   As String       '�����p�x
    Dim sndFlg  As String       '���M�t���O
    Dim i       As Integer
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function UpdateTBCMY011SendFlag"
    
    UpdateTBCMY011SendFlag = FUNCTION_RETURN_FAILURE

'    Del Start 2011/07/05 Y.Hitomi
'    '���i�d�lSXL�f�[�^�P���猟���p�x���擾����B
'    Set rs = Nothing
'    sql = ""
'    sql = sql & "SELECT" & vbCrLf
'    sql = sql & " HSXRKHNM" & vbCrLf        '�����p�x
'    sql = sql & "FROM" & vbCrLf
'    sql = sql & " TBCME018" & vbCrLf
'    sql = sql & "WHERE" & vbCrLf
'    sql = sql & " HINBAN ='" & HIN.hinban & "' AND" & vbCrLf
'    sql = sql & " MNOREVNO =" & HIN.mnorevno & " AND" & vbCrLf
'    sql = sql & " FACTORY ='" & HIN.factory & "' AND" & vbCrLf
'    sql = sql & " OPECOND ='" & HIN.opecond & "'" & vbCrLf
'
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'    If rs.RecordCount = 0 Then
'        Set rs = Nothing
'        UpdateTBCMY011SendFlag = FUNCTION_RETURN_SUCCESS
'        GoTo proc_exit
'    End If
'
'    hindo = rs("HSXRKHNM")
'
'    '�����p�x���u0�v�̏ꍇ�A�i���V�X�e�����M�t���O�iG52)='3'[���M�\��]
'    '����ȊO�ł͕i���V�X�e�����M�t���O�iG52)='2'[���M�ς�]�Ƃ���B
'    If hindo = "0" Then sndFlg = "3" Else sndFlg = "2"
'    Del End 2011/07/05 Y.Hitomi
    
'Add Start 2011/07/05 Y.Hitomi �S�����M�Ή�
    sndFlg = "3"
'Add End   2011/07/05 Y.Hitomi
    
    sql = ""
    sql = sql & "UPDATE" & vbCrLf
    sql = sql & " TBCMY011" & vbCrLf
    sql = sql & "SET" & vbCrLf
    sql = sql & " TBCMY011.QA1SNDFLG='" & sndFlg & "'," & vbCrLf                         '�i���V�X�e�����M�t���O�iG52)
    sql = sql & " UPDPROC='CW800'," & vbCrLf                                             '�X�V�H���@2011/01/31 tkimura
    sql = sql & " UPDDATE=to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')" & vbCrLf '�X�V���t�@2011/01/31 tkimura
    sql = sql & "WHERE" & vbCrLf
    sql = sql & " MSXLID='" & SXLID & "'" & vbCrLf     'SXLID
        
    If 0 >= OraDB.ExecuteSQL(sql) Then
        UpdateTBCMY011SendFlag = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    Set rs = Nothing
    
    UpdateTBCMY011SendFlag = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    UpdateTBCMY011SendFlag = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/17 tkimura ADD START
'�T�v      :�E�F�n�|�Z���^�|���ɏ��(TBCMY011)��
'           �C���S�b�g�ʒu���㗦,���t�����R�l(Center)���X�V����B
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :blockId         , I  ,String            , �u���b�NID
'          :blockSeq        , I  ,String            , �u���b�N���A��
'          :up_Ratio        , I  ,String            , �C���S�b�g�ʒu���㗦
'          :rs_Meas         , I  ,String            , ���t�����R�l(Center)
'          :dSXL_Pos        , I  ,Double            , SXL�ʒu(Intel)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :
'����      :2011/01/17 tkimura
'       �@�@2011/04/25 Marushita ����SXL�ʒu�ǉ��iMicron�ݺޯĈʒu�Ǘ��ǉ��Ή��j
'Private Function UpdateTBCMY011SuiteiResData(ByVal BLOCKID As String, _
'                                             ByVal BLOCKSEQ As Integer, _
'                                             ByVal up_Ratio As String, _
'                                             ByVal rs_Meas As String) As FUNCTION_RETURN
        
Private Function UpdateTBCMY011SuiteiResData(ByVal BLOCKID As String, _
                                             ByVal BLOCKSEQ As Integer, _
                                             ByVal up_Ratio As String, _
                                             ByVal rs_Meas As String, _
                                             ByVal dSXL_Pos As Double) As FUNCTION_RETURN
    Dim sql     As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function UpdateTBCMY011SuiteiResData"
    
    UpdateTBCMY011SuiteiResData = FUNCTION_RETURN_FAILURE
        
    sql = ""
    sql = sql & "UPDATE" & vbCrLf
    sql = sql & " TBCMY011" & vbCrLf
    sql = sql & "SET" & vbCrLf
    sql = sql & " UP_RATIO='" & up_Ratio & "'," & vbCrLf    '�C���S�b�g�ʒu���㗦
    sql = sql & " RS_MEAS='" & rs_Meas & "'" & vbCrLf       '���t�����R�l(Center)
    sql = sql & ",HTOP_POS='" & dSXL_Pos & "'" & vbCrLf     '�␳������(SXL�ʒu(Intel)) 2011/04/25 ADD Marushita
    sql = sql & "WHERE" & vbCrLf
    sql = sql & " LOTID='" & BLOCKID & "' AND" & vbCrLf     '�u���b�NID
    sql = sql & " BLOCKSEQ=" & BLOCKSEQ & "" & vbCrLf       '�u���b�N���A��
    
    If 0 >= OraDB.ExecuteSQL(sql) Then
        UpdateTBCMY011SuiteiResData = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
''    Debug.Print "BLOCK ", BLOCKID, " SEQ ", BLOCKSEQ
''    Debug.Print "UP_RATIO ", up_Ratio, " RS_MEAS ", rs_Meas, " HTOP_POS ", dSXL_Pos
            
    UpdateTBCMY011SuiteiResData = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    UpdateTBCMY011SuiteiResData = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/18 tkimura ADD START
'�T�v      :��Top�ʒu���㗦,��Bot�ʒu���㗦,�����ΐ�,���R�l�����Ƃ߂�B
'���Ұ��@�@:�ϐ���          , IO , �^                       , ����
'          :CRYNUM          , I  ,String                    , �����ԍ�
'          :sRsData         , I  ,String                    , ��R�f�[�^(Top/Bot)
'          :sRsPos          , I  ,String                    , ��R�ʒu�f�[�^(Top/Bot)
'          :d               , O  ,type_Coefficient_new2     , �����R,������㗦�v�Z�\����
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@         , ����
'����      :
'����      :2011/01/18 tkimura
Private Function GetStandardPosRes(ByVal CRYNUM As String, _
                                   ByRef sRsData() As String, _
                                   ByRef sRsPos() As String, _
                                   ByRef d As type_Coefficient_new2) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    
    Dim wgtCharge   As Long                 '�ΐ͌v�Z�p�p�����[�^(�`���[�W��)
    Dim wgtChargeA  As Long                 '�ΐ͌v�Z�p�p�����[�^(START�`���[�W��)
    Dim wgtTop      As Double               '�ΐ͌v�Z�p�p�����[�^(Top���)
    Dim wgtTopCut   As Double               '�ΐ͌v�Z�p�p�����[�^(���d��)
    Dim DM          As Double               '�ΐ͌v�Z�p�p�����[�^(���a����)
    Dim HIKIFLG     As Integer              '���グ�t���O(1=�ʏ�A2=BC����)
    
    '>>>>> 2011/04/25 ADD Marushita ��Micron�ݺޯĈʒu�Ǘ��ǉ��Ή�
    Dim p_CRYNUM    As String               '�O�����ԍ��擾�p
    Dim p_wgtTop    As Double               '�O������擾�p(�OTop���)
    Dim p_DM        As Double               '�O������擾�p(�O���a����)
    Dim p_wgtTA     As Double               '�O������擾�p(�O�e�C���d��)
    Dim p_LENTK     As Long                 '�O������擾�p(�O���㒷)
    Dim dMaeBatLen  As Double               '�O�o�b�`�������@Add 2011/09/27
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function GetStandardPosRes"
    
    GetStandardPosRes = FUNCTION_RETURN_FAILURE
        
    If GetCoeffParams_new2(CRYNUM, wgtCharge, wgtChargeA, wgtTop, _
                           wgtTopCut, DM, HIKIFLG) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
        
    d.DUNMENSEKI = AreaOfCircle(DM)     '�f�ʐ�(���a���ς��v�Z)
    d.TOPSMPLPOS = sRsPos(1)            'TOP�ʒu
    d.BOTSMPLPOS = sRsPos(2)            'BOT�ʒu
    d.CHARGEWEIGHT = wgtCharge          '�`���[�W��
    d.TOPWEIGHT = wgtTop + wgtTopCut    '�g�b�v�d��(Top��ʁ{���d��)
    d.TOPRES = sRsData(1)               'TOP�ʒu��R�l
    d.BOTRES = sRsData(6)               'BOT�ʒu��R�l
    d.CHARGEWEIGHTA = wgtChargeA        'START�`���[�W��
    d.HIKIFLG = HIKIFLG                 '����t���O(1=�ʏ�A2=BC����)
    
    '��Top�ʒu���グ�������߂�B
    d.SMPLPOS = d.TOPSMPLPOS
    d.GT = HikiageCalculation(d)
    '��Bot�ʒu���グ�������߂�B
    d.SMPLPOS = d.BOTSMPLPOS
    d.GB = HikiageCalculation(d)

    '���s�ΐ͂����߂�B
    d.Henseki = CoefficientCalculation_new2(d)
    If d.Henseki = -9999 Then
        d.Henseki = 0       'SXLRS_�����ΐ�
    End If
    
    '���R�l�����߂�B
    d.KIJUNTEIKOU = StandardResCalculation(d)
    If d.KIJUNTEIKOU = -9999 Then
        d.KIJUNTEIKOU = 0
    End If
    
    '>>>>> 2011/04/25 ADD Marushita ��Micron�ݺޯĈʒu�Ǘ��ǉ��Ή�
    d.HOSEICHO = 0
    '�␳�����������߂�
    If HIKIFLG = "1" Then       '�ʏ����̏ꍇ
        'Top���/(�f�ʐ�*0.00233)
        d.HOSEICHO = wgtTop / (d.DUNMENSEKI * HIJU_SILICONE)
    ElseIf HIKIFLG = "2" Then   'B�AC����(�c�ʈ�)�̏ꍇ
        '>>>>> 2011/09/27 ADD Marushita C�����ȏ�Ή�
        '�O�o�b�`���������擾����
        dMaeBatLen = GetMaeBatLen(CRYNUM)
        '���݌����̕␳(Top���/(�f�ʐ�*0.00233))�𑫂��ĕ␳�������߂�
        d.HOSEICHO = dMaeBatLen + wgtTop / (d.DUNMENSEKI * HIJU_SILICONE)
'        '�O����̌����ԍ����擾(�����ԍ�)
'        If GetPreCrynum(CRYNUM, p_CRYNUM) = FUNCTION_RETURN_FAILURE Then
'        Else
'            '�O����̌��������擾(�OTop��ʁA�O���a���ρA�O���㒷�A�O�e�C���d��)
'            If GetPreXSDC1(p_CRYNUM, p_wgtTop, p_DM, p_LENTK, p_wgtTA) = FUNCTION_RETURN_FAILURE Then
'            Else
'                '�␳�������̌v�Z(�O�o�b�`�̌������𑫂�)
'                '�␳������ = (�OTop���/(�O�f�ʐ�*0.00233))+�O���㒷+((�O�e�C���d��+Top���)/(�f�ʐ�*0.00233))
'                d.HOSEICHO = (p_wgtTop / (AreaOfCircle(p_DM) * HIJU_SILICONE)) + p_LENTK + _
'                       ((p_wgtTA + wgtTop) / (d.DUNMENSEKI * HIJU_SILICONE))
'            End If
'        End If
        '<<<<< 2011/09/27 ADD Marushita C�����ȏ�Ή�
    End If
    '<<<<< 2011/04/25 ADD Marushita ��Micron�ݺޯĈʒu�Ǘ��ǉ��Ή�
    
''    Debug.Print "Top���グ��", d.GT
''    Debug.Print "Bot���グ��", d.GB
''    Debug.Print "���a", DM
''    Debug.Print "�f�ʐ�", d.DUNMENSEKI
''    Debug.Print "�g�b�v�J�b�g�d��", wgtTopCut
''    Debug.Print "�g�b�v�d��", wgtTop
''    Debug.Print "����`���[�W��", d.CHARGEWEIGHT
''    Debug.Print "����`���[�W��A", d.CHARGEWEIGHTA
''    Debug.Print "�g�b�v��R�l", d.TOPRES
''    Debug.Print "�{�g����R�l", d.BOTRES
''    Debug.Print "�g�b�v�ʒu", d.TOPSMPLPOS
''    Debug.Print "�{�g���ʒu", d.BOTSMPLPOS
''    Debug.Print "�����ΐ�", d.Henseki
''    Debug.Print "���R�l", d.KIJUNTEIKOU
''    Debug.Print "�␳������", d.HOSEICHO
    
    '2011/01/19 kimura �f�o�b�O���擾(��ŏ���)
''    Debug.Print "Top���グ��"
''    Debug.Print d.GT
''    Debug.Print "Bot���グ��"
''    Debug.Print d.GB
''    Debug.Print "���a"
''    Debug.Print DM
''    Debug.Print "�f�ʐ�"
''    Debug.Print d.DUNMENSEKI
''    Debug.Print "�g�b�v�J�b�g�d��"
''    Debug.Print wgtTopCut
''    Debug.Print "�g�b�v�d��"
''    Debug.Print wgtTop
''    Debug.Print "����`���[�W��(�O�o�b�`���܂�)"
''    Debug.Print d.CHARGEWEIGHT
''    Debug.Print "����`���[�W��A"
''    Debug.Print d.CHARGEWEIGHTA
''    Debug.Print "�g�b�v��R�l"
''    Debug.Print d.TOPRES
''    Debug.Print "�{�g����R�l"
''    Debug.Print d.BOTRES
''    Debug.Print "�g�b�v�ʒu"
''    Debug.Print d.TOPSMPLPOS
''    Debug.Print "�{�g���ʒu"
''    Debug.Print d.BOTSMPLPOS
''    Debug.Print "�����ΐ�"
''    Debug.Print d.Henseki
''    Debug.Print "���R�l"
''    Debug.Print d.KIJUNTEIKOU
    
    GetStandardPosRes = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    GetStandardPosRes = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
' 2011/01/18 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/18 tkimura ADD START
'�T�v      :SXLID�̖��t���Ƃ�
'           �C���S�b�g�ʒu���㗦,���t�����R�l(Center)���v�Z����B
'���Ұ��@�@:�ϐ���          , IO , �^                       , ����
'          :SXLID           , I  ,String                    , SXLID
'          :d               , I  ,type_Coefficient_new2     , �����R,������㗦�v�Z�\����
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@         , ����
'����      :
'����      :2011/01/18 tkimura
Private Function SuiteiResDataCalculation(ByVal SXLID As String, _
                                          ByRef d As type_Coefficient_new2, sUP_RATIO() As String) As FUNCTION_RETURN
'Private Function SuiteiResDataCalculation(ByVal SXLID As String, _
                                          ByRef d As type_Coefficient_new2) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    Dim recY011() As typ_Y011               '' Y011�̃f�[�^���܂Ƃ߂��\����
    Dim y011Cnt As Integer
    Dim Index   As Integer
    Dim suiteiHiki As String                '����Ώۈ��グ��
    Dim suiteiTei  As String                '����ʒu���R�l
    Dim dSXL_Pos   As Double                'SXL�ʒu(Intel)
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function SuiteiResDataCalculation"
    
    SuiteiResDataCalculation = FUNCTION_RETURN_FAILURE
        
    '����ʒu���擾����B
    Set rs = Nothing
    sql = ""
    sql = sql & "SELECT " & vbCrLf
    sql = sql & " LOTID," & vbCrLf                  '�u���b�NID
    sql = sql & " BLOCKSEQ," & vbCrLf               '�u���b�N���A��
    sql = sql & " RITOP_POS " & vbCrLf              '���_�������ʒu
    sql = sql & "FROM " & vbCrLf
    sql = sql & " TBCMY011 " & vbCrLf
    sql = sql & "WHERE " & vbCrLf
    sql = sql & " MSXLID='" & SXLID & "'" & vbCrLf  'SXLID
    sql = sql & "ORDER BY " & vbCrLf
    sql = sql & " LOTID," & vbCrLf                  '�u���b�NID
    sql = sql & " BLOCKSEQ" & vbCrLf                '�u���b�N���A��
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            
    ''Debug.Print "���[�v����Ƃ���(Y011)��SQL��"
    ''Debug.Print sql
    
    y011Cnt = rs.RecordCount - 1
    ReDim recY011(y011Cnt)      '���z��ɃZ�b�g����Ӗ��́H
    Index = 0
    Do While Not rs.EOF
        '��ۯ�ID
        If IsNull(rs("LOTID")) Then     '��NULL�͂��肦�Ȃ��H
            recY011(Index).LOTID = ""
        Else
            recY011(Index).LOTID = rs("LOTID")
        End If
        '��ۯ����A��
        If IsNull(rs("BLOCKSEQ")) Then  '��NULL�͂��肦�Ȃ��H
            recY011(Index).BLOCKSEQ = 0
        Else
            recY011(Index).BLOCKSEQ = CInt(rs("BLOCKSEQ"))
        End If
        '����ʒu(���_�������ʒu)
        'If IsNull(rs.Fields("RTOP_POS")) Then
        If IsNull(rs.Fields("RITOP_POS")) Then
            recY011(Index).RITOP_POS = vbNullString     '��NULL�̏ꍇ��NULL���Z�b�g(�s�v�H)
        Else
            'recY011(Index).RTOP_POS = rs.Fields("RTOP_POS")
            recY011(Index).RITOP_POS = rs.Fields("RITOP_POS")
        End If
                
        '����Ώۈ��グ�������߂�B
        d.SMPLPOS = recY011(Index).RITOP_POS        '������ʒu(���_�������ʒu)�̃Z�b�g
        d.SUITEIHIKIRITU = HikiageCalculation(d)
                
        '����ʒu���R�l�����߂�B
        d.SUITEITEIKOU = SuiteiResCalculation(d)

        
'�CTBCMY011�e�[�u���̃C���S�b�g�ʒu���㗦,���t�����R�l���X�V����B
        '2011/01/19 tkimura �����_2������4���ɑ��₵���B
        suiteiHiki = RoundDown(d.SUITEIHIKIRITU, 4)
        suiteiTei = RoundDown(d.SUITEITEIKOU, 5)
        '2011/04/25 ADD Marushita ��Micron�ݺޯĈʒu�Ǘ��ǉ��Ή�
        '��SXL�ʒu(Intel)���v�Z�������_2���Ő؂�̂�(�␳������+���_�������ʒu)
        dSXL_Pos = RoundDown(d.HOSEICHO + recY011(Index).RITOP_POS, 2)
        
        'Add Start 2011/05/31 Y.Hitomi
        If Index = 0 Then
            sUP_RATIO(0) = suiteiHiki
        End If
        'Add End   2011/05/31 Y.Hitomi
        
        '2011/04/25 MOD Marushita ��Micron�ݺޯĈʒu�Ǘ��ǉ��Ή�
        '��SXL�ʒu(Intel)�������ɒǉ�
'        If UpdateTBCMY011SuiteiResData(recY011(Index).LOTID, _
'                                       recY011(Index).BLOCKSEQ, _
'                                       suiteiHiki, _
'                                       suiteiTei) = FUNCTION_RETURN_FAILURE Then
        If UpdateTBCMY011SuiteiResData(recY011(Index).LOTID, _
                                       recY011(Index).BLOCKSEQ, _
                                       suiteiHiki, suiteiTei, dSXL_Pos) _
                                       = FUNCTION_RETURN_FAILURE Then
            GoTo proc_exit
        End If
            
        Index = Index + 1
        rs.MoveNext
    Loop
    
    'Add Start 2011/05/31 Y.Hitomi
    sUP_RATIO(1) = suiteiHiki
    'Add End   2011/05/31 Y.Hitomi
    
    SuiteiResDataCalculation = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    SuiteiResDataCalculation = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
' 2011/01/18 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/02/14 tkimura ADD START
'�T�v      :CLESTA�]������(TBCMJ023)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS�\����   (�V����يǗ�(��ۯ�))
'          :recX006         , O  ,c_cmzcrec         , TBCMX006�\����(Cu-Deco�\����)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :CLESTA�]������(TBCMJ023)�����ް����擾���ACu-Deco�\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
Private Function getTBCMJ023(CRYNUM As String, recXSDCS As c_cmzcrec, recX006 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ023"
    
    getTBCMJ023 = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    With recX006
        '.Fields("SXLC_HOUDATTIM").Value =                 'C-������t
        .Fields("SXLC_SGYCD").Value = " "                  'C-�����Ǝ�
        .Fields("SXLC_PTN").Value = " "                    'C-�����
        .Fields("SXLC_DHANKEI").Value = -1                 'C-Disk�a(���a)
        .Fields("SXLC_RNAIKEI").Value = -1                 'C-Ring���a
        .Fields("SXLC_RGAIKEI").Value = -1                 'C-Ring�O�a
        .Fields("SXLC_SZ").Value = " "                     'C-�������
        '.Fields("SXLCJ_HOUDATTIM").Value =                 'CJ-������t
        .Fields("SXLCJ_SGYCD").Value = " "                 'CJ-�����Ǝ�
        .Fields("SXLCJ_PTN").Value = " "                   'CJ-�����
        .Fields("SXLCJ_DHANKEI").Value = -1                'CJ-Disk�a(���a)
        .Fields("SXLCJ_RNAIKEI").Value = -1                'CJ-Ring���a
        .Fields("SXLCJ_RGAIKEI").Value = -1                'CJ-Ring�O�a
        .Fields("SXLCJ_BNAIKEI").Value = -1                'CJ-Band���a
        .Fields("SXLCJ_BGAIKEI").Value = -1                'CJ-Band�O�a
        .Fields("SXLCJ_RHABA").Value = -1                  'CJ-Ring��
        .Fields("SXLCJ_PIHABA").Value = -1                 'CJ-Pi��
        .Fields("SXLCJ_NETU").Value = " "                  'CJ-�M�����@
        .Fields("SXLCJ_JUDGE").Value = " "                 'CJ-���ʕʔ��茋��
        '.Fields("SXLCJLT_HOUDATTIM").Value =               'CJLT-������t
        .Fields("SXLCJLT_SGYCD").Value = " "               'CJLT-�����Ǝ�
        .Fields("SXLCJLT_PTN").Value = " "                 'CJLT-�p�^�[��
        .Fields("SXLCJLT_DHANKEI").Value = -1              'CJLT-Disk�a�i���a�j
        .Fields("SXLCJLT_RNAIKEI").Value = -1              'CJLT-Ring���a
        .Fields("SXLCJLT_RGAIKEI").Value = -1              'CJLT-Ring�O�a
        .Fields("SXLCJLT_BNAIKEI").Value = -1              'CJLT-Band���a
        .Fields("SXLCJLT_BGAIKEI").Value = -1              'CJLT-Band�O�a
        .Fields("SXLCJLT_RHABA").Value = -1                'CJLT-Ring��
        .Fields("SXLCJLT_PIHABA").Value = -1               'CJLT-Pi ��
        .Fields("SXLCJLT_BHABA").Value = -1                'CJLT-Band��
        .Fields("SXLCJLT_NETU").Value = " "                'CJLT-�M�����@
        '.Fields("SXLCJ2_HOUDATTIM").Value =                'CJ2-������t
        .Fields("SXLCJ2_SGYCD").Value = " "                'CJ2-�����Ǝ�
        .Fields("SXLCJ2_PTN").Value = " "                  'CJ2-�����
        .Fields("SXLCJ2_DHANKEI").Value = -1               'CJ2-Disk�a(���a)
        .Fields("SXLCJ2_RNAIKEI").Value = -1               'CJ2-Ring���a
        .Fields("SXLCJ2_RGAIKEI").Value = -1               'CJ2-Ring�O�a
        .Fields("SXLCJ2_PIHABA").Value = -1                'CJ2-Pi��
        .Fields("SXLCJ2_NETU").Value = " "                 'CJ2-�M�����@
        .Fields("SXLCJ2_JUDGE").Value = " "                'CJ2-���ʕʔ��茋��

        '-------------------- TBCMJ023�̓ǂݍ���(CJ) ----------------------------------------
        sql = "select * from TBCMJ023 "
        sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
        sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDL4CS").Value
        sql = sql & "order by TRANCNT desc"
        'sql = "select J023.*,to_char(J023.REGDATEC,'YYYY/MM/DD HH24:MI:SS') AS REGDATE from (" & sql & ") J023 where rownum = 1"
        sql = "select * from (" & sql & ") where rownum = 1"
        Debug.Print (sql)
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            getTBCMJ023 = FUNCTION_RETURN_SUCCESS                     '2011/02/28 tkimura
            GoTo proc_exit
        End If

        'TBCMX006
        .Fields("SXLC_HOUDATTIM").Value = rs("REGDATEC")                'C-������t
        '.Fields("SXLC_HOUDATTIM").Value = "to_date( " & rs("REGDATEC") & ",'yyyy/mm/dd hh24:mi:ss')"                'C-������t"
        .Fields("SXLC_SGYCD").Value = rs("TSTAFFIDC")                    'C-�����Ǝ�
        .Fields("SXLC_PTN").Value = rs("CPTNJSK")                        'C-�����
        If IsNull(rs("CDISKJSK").Value) = False Then
            .Fields("SXLC_DHANKEI").Value = rs("CDISKJSK")               'C-Disk�a(���a)
        End If
        If IsNull(rs("CRINGNKJSK").Value) = False Then
            .Fields("SXLC_RNAIKEI").Value = rs("CRINGNKJSK")             'C-Ring���a
        End If
        If IsNull(rs("CRINGGKJSK").Value) = False Then
            .Fields("SXLC_RGAIKEI").Value = rs("CRINGGKJSK")             'C-Ring�O�a
        End If
        .Fields("SXLC_SZ").Value = rs("C_SZ")                            'C-�������
        .Fields("SXLCJ_HOUDATTIM").Value = rs("REGDATECJ")               'CJ-������t
        'Format(Now, "yyyy/mm/dd hh:mm:ss")
        '.Fields("SXLCJ_HOUDATTIM").Value = Format(rs("REGDATECJ"), "yyyy/mm/dd hh:mm:ss")              'CJ-������t
        .Fields("SXLCJ_SGYCD").Value = rs("TSTAFFIDCJ")                  'CJ-�����Ǝ�
        .Fields("SXLCJ_PTN").Value = rs("CJPTNJSK")                      'CJ-�����
        If IsNull(rs("CJDISKJSK").Value) = False Then
            .Fields("SXLCJ_DHANKEI").Value = rs("CJDISKJSK")             'CJ-Disk�a(���a)
        End If
        If IsNull(rs("CJRINGNKJSK").Value) = False Then
            .Fields("SXLCJ_RNAIKEI").Value = rs("CJRINGNKJSK")           'CJ-Ring���a
        End If
        If IsNull(rs("CJRINGGKJSK").Value) = False Then
            .Fields("SXLCJ_RGAIKEI").Value = rs("CJRINGGKJSK")           'CJ-Ring�O�a
        End If
        If IsNull(rs("CJBANDNKJSK").Value) = False Then
            .Fields("SXLCJ_BNAIKEI").Value = rs("CJBANDNKJSK")           'CJ-Band���a
        End If
        If IsNull(rs("CJBANDGKJSK").Value) = False Then
            .Fields("SXLCJ_BGAIKEI").Value = rs("CJBANDGKJSK")           'CJ-Band�O�a
        End If
        If IsNull(rs("CJRINGCALC").Value) = False Then
            .Fields("SXLCJ_RHABA").Value = rs("CJRINGCALC")              'CJ-Ring��
        End If
        If IsNull(rs("CJPICALC").Value) = False Then
            .Fields("SXLCJ_PIHABA").Value = rs("CJPICALC")               'CJ-Pi��
        End If
        .Fields("SXLCJ_NETU").Value = rs("CJ_NETU")                      'CJ-�M�����@
        .Fields("SXLCJ_JUDGE").Value = rs("CJHANTEI")                    'CJ-���ʕʔ��茋��
        .Fields("SXLCJLT_HOUDATTIM").Value = rs("REGDATECJLT")           'CJLT-������t
        .Fields("SXLCJLT_SGYCD").Value = rs("TSTAFFIDCJLT")              'CJLT-�����Ǝ�
        .Fields("SXLCJLT_PTN").Value = rs("CJLTPTNJSK")                  'CJLT-�p�^�[��
        If IsNull(rs("CJLTDISKJSK").Value) = False Then
            .Fields("SXLCJLT_DHANKEI").Value = rs("CJLTDISKJSK")         'CJLT-Disk�a�i���a�j
        End If
        If IsNull(rs("CJLTRINGNKJSK").Value) = False Then
            .Fields("SXLCJLT_RNAIKEI").Value = rs("CJLTRINGNKJSK")       'CJLT-Ring���a
        End If
        If IsNull(rs("CJLTRINGGKJSK").Value) = False Then
            .Fields("SXLCJLT_RGAIKEI").Value = rs("CJLTRINGGKJSK")       'CJLT-Ring�O�a
        End If
        If IsNull(rs("CJLTBANDNKJSK").Value) = False Then
            .Fields("SXLCJLT_BNAIKEI").Value = rs("CJLTBANDNKJSK")       'CJLT-Band���a
        End If
        If IsNull(rs("CJLTBANDGKJSK").Value) = False Then
            .Fields("SXLCJLT_BGAIKEI").Value = rs("CJLTBANDGKJSK")       'CJLT-Band�O�a
        End If
        If IsNull(rs("CJLTRINGCALC").Value) = False Then
            .Fields("SXLCJLT_RHABA").Value = rs("CJLTRINGCALC")          'CJLT-Ring��
        End If
        If IsNull(rs("CJLTPICALC").Value) = False Then
            .Fields("SXLCJLT_PIHABA").Value = rs("CJLTPICALC")           'CJLT-Pi ��
        End If
        If IsNull(rs("CJLTPICALC").Value) = False Then
            .Fields("SXLCJLT_BHABA").Value = rs("CJLTPICALC")            'CJLT-Band��
        End If
        .Fields("SXLCJLT_NETU").Value = rs("CJLT_NETU")                  'CJLT-�M�����@
        .Fields("SXLCJ2_HOUDATTIM").Value = rs("REGDATECJ2")             'CJ2-������t
        '.Fields("SXLCJ2_HOUDATTIM").Value = CDate(rs("REGDATECJ2"))             'CJ2-������t
        .Fields("SXLCJ2_SGYCD").Value = rs("TSTAFFIDCJ2")                'CJ2-�����Ǝ�
        .Fields("SXLCJ2_PTN").Value = rs("CJ2PTNJSK")                    'CJ2-�����
        If IsNull(rs("CJ2DISKJSK").Value) = False Then
            .Fields("SXLCJ2_DHANKEI").Value = rs("CJ2DISKJSK")           'CJ2-Disk�a(���a)
        End If
        If IsNull(rs("CJ2RINGNKJSK").Value) = False Then
            .Fields("SXLCJ2_RNAIKEI").Value = rs("CJ2RINGNKJSK")         'CJ2-Ring���a
        End If
        If IsNull(rs("CJ2RINGGKJSK").Value) = False Then
            .Fields("SXLCJ2_RGAIKEI").Value = rs("CJ2RINGGKJSK")         'CJ2-Ring�O�a
        End If
        If IsNull(rs("CJ2PICALC").Value) = False Then
            .Fields("SXLCJ2_PIHABA").Value = rs("CJ2PICALC")             'CJ2-Pi��
        End If
        .Fields("SXLCJ2_NETU").Value = rs("CJ2_NETU")                    'CJ2-�M�����@
        .Fields("SXLCJ2_JUDGE").Value = rs("CJ2HANTEI")                  'CJ2-���ʕʔ��茋��

        Set rs = Nothing
    End With

    getTBCMJ023 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ023 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
' 2011/02/14 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/02/14 tkimura ADD START
'�T�v      :����OSF���уf�[�^�擾�ݒ�(TBCMJ005)�f�[�^�擾�ݒ�
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :CRYNUM          , I  ,String            , �����ԍ�
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS�\����   (�V����يǗ�(��ۯ�))
'          :recX006         , O  ,c_cmzcrec         , TBCMX006�\����(Cu-Deco�\����)
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :����OSF����(TBCMJ005)�����ް����擾���ACu-Deco�\���̂ɾ�Ă���
'����      :2003/10/18 SystemBrain �V�K�쐬
Private Function getTBCMJ005CuDeco(CRYNUM As String, recXSDCS As c_cmzcrec, recX006 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ005CuDeco"
    
    getTBCMJ005CuDeco = FUNCTION_RETURN_FAILURE

    '-------------------- �����ر ----------------------------------------
    With recX006
        '.Fields("SXLCOSF3_HOUDATTIM").Value =                  'C-OSF3������t
        .Fields("SXLCOSF3_SGYCD").Value = " "                   'C-OSF3�����Ǝ�
        .Fields("SXLCOSF3_PTN").Value = " "                     'C-OSF3�����
        .Fields("SXLCOSF3_DHANKEI").Value = -1                  'C-OSF3Disk�a(���a)
        .Fields("SXLCOSF3_RNAIKEI").Value = -1                  'C-OSF3Ring���a
        .Fields("SXLCOSF3_RGAIKEI").Value = -1                  'C-OSF3Ring�O�a
        .Fields("SXLCOSF3_RHABA").Value = -1                    'C-OSF3Ring��
        .Fields("SXLCOSF3_NETU").Value = " "                    'C-OSF3�M�����@
        .Fields("SXLCOSF3_JUDGE").Value = " "                   'C-OSF3���ʕʔ��茋��

        '-------------------- TBCMJ005�̓ǂݍ���(OSF1,2) ----------------------------------------
        sql = "select"
        sql = sql & "    REGDATE,"
        sql = sql & "    TSTAFFID,"
        sql = sql & "    SXLCOSF3_PTN,"
        sql = sql & "    SXLCOSF3_DHANKEI,"
        sql = sql & "    SXLCOSF3_RNAIKEI,"
        sql = sql & "    SXLCOSF3_RGAIKEI,"
        sql = sql & "    SXLCOSF3_RHABA,"
        sql = sql & "    HTPRC,"
        sql = sql & "    PTNJUDGRES"
        sql = sql & " from"
        sql = sql & "    (select "
        sql = sql & "        TSTAFFID,"
        sql = sql & "        REGDATE,"
        sql = sql & "        case when(OSFRD1 = '-' and OSFRD2 = '-') then '0' "
        sql = sql & "             when(OSFRD1 = 'D' and OSFRD2 = '-') then '2' "
        sql = sql & "             when(OSFRD1 = 'R' and OSFRD2 = '-') then '1' "
        sql = sql & "             when(OSFRD1 = 'D' and OSFRD2 = 'R' or OSFRD1 = 'R' and OSFRD2 = 'D') then '3' "
'Cng Start 2011/04/12 Y.Hitomi
'        sql = sql & "             else '-1' "
        sql = sql & "             when(OSFRD1 = 'R' and OSFRD2 = 'R') then '1' "
        sql = sql & "             else ' ' "
'Cng Start 2011/04/12 Y.Hitomi
        sql = sql & "        end SXLCOSF3_PTN,"
        sql = sql & "        case when(OSFRD1 = 'D') then OSFWID1 "
        sql = sql & "             when(OSFRD2 = 'D') then OSFWID2 "
        sql = sql & "        end SXLCOSF3_DHANKEI,"
        sql = sql & "        case when(OSFRD1 = 'R') then (150-OSFPOS1-OSFWID1) "
        sql = sql & "             when(OSFRD2 = 'R') then (150-OSFPOS2-OSFWID2) "
        sql = sql & "        end SXLCOSF3_RNAIKEI,"
        sql = sql & "        case when(OSFRD1 = 'R') then (150-OSFPOS1) "
        sql = sql & "             when(OSFRD2 = 'R') then (150-OSFPOS2) "
        sql = sql & "        end SXLCOSF3_RGAIKEI,"
        sql = sql & "        case when(OSFRD1 = 'R') then OSFWID1 "
        sql = sql & "             when(OSFRD2 = 'R') then OSFWID2  "
        sql = sql & "        end SXLCOSF3_RHABA,"
        sql = sql & "        HTPRC,"
        sql = sql & "        PTNJUDGRES "
        sql = sql & "    from "
        sql = sql & "        TBCMJ005"
        sql = sql & "    where "
        sql = sql & "        CRYNUM = '" & CRYNUM & "' and "
        sql = sql & "        SMPLNO = " & recXSDCS("CRYSMPLIDL4CS").Value
        sql = sql & "        order by TRANCNT desc"
        sql = sql & "    ) "
        sql = sql & " where rownum = 1"
        Debug.Print (sql)
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            getTBCMJ005CuDeco = FUNCTION_RETURN_SUCCESS                     '2011/02/28 tkimura
            GoTo proc_exit
        End If

        'TBCMX006(���l��NULL�����邱�Ƃ�h���ł����B)
        .Fields("SXLCOSF3_HOUDATTIM").Value = rs("REGDATE")                  'C-OSF3������t
        .Fields("SXLCOSF3_SGYCD").Value = rs("TSTAFFID")                     'C-OSF3�����Ǝ�
        .Fields("SXLCOSF3_PTN").Value = rs("SXLCOSF3_PTN")                   'C-OSF3�����
        If IsNull(rs("SXLCOSF3_DHANKEI").Value) = False Then
            .Fields("SXLCOSF3_DHANKEI").Value = rs("SXLCOSF3_DHANKEI")       'C-OSF3Disk�a(���a)
        End If
        If IsNull(rs("SXLCOSF3_RNAIKEI").Value) = False Then
            .Fields("SXLCOSF3_RNAIKEI").Value = rs("SXLCOSF3_RNAIKEI")       'C-OSF3Ring���a
        End If
        If IsNull(rs("SXLCOSF3_RGAIKEI").Value) = False Then
            .Fields("SXLCOSF3_RGAIKEI").Value = rs("SXLCOSF3_RGAIKEI")       'C-OSF3Ring�O�a
        End If
        If IsNull(rs("SXLCOSF3_RHABA").Value) = False Then
            .Fields("SXLCOSF3_RHABA").Value = rs("SXLCOSF3_RHABA")           'C-OSF3Ring��
        End If
        .Fields("SXLCOSF3_NETU").Value = rs("HTPRC")                         'C-OSF3�M�����@
        .Fields("SXLCOSF3_JUDGE").Value = rs("PTNJUDGRES")                   'C-OSF3���ʕʔ��茋��

        Set rs = Nothing
    End With

    getTBCMJ005CuDeco = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ005CuDeco = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
' 2011/02/14 tkimura ADD END
'=================================================================================

''2011/01/17 tkimura ADD START ==========================================================>
'�T�v      :�ΐ͌v�Z�ɕK�v�Ȋe���v�d�ʎ��т��擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :CRYNUM        ,I  ,String    ,�����ԍ�
'          :wgtCharge     ,O  ,Long      ,�F����
'          :wgtChargeA    ,O  ,Long      ,A�����̘F����
'          :wgtTop        ,O  ,Double    ,�g�b�v�d�ʎ��ђl
'          :wgtTopCut     ,O  ,Double    ,�g�b�v�J�b�g�d�ʎ��ђl
'          :DM            ,O  ,Double    ,���a�P�`�R�̕���
'          :hikiFlg       ,O  ,Integer   ,���グ�t���O(1=�ʏ�A2=BC����)
'          :�߂�l        ,O  ,FUNCTION_RETURN,
'����      :�y�}���`����Ή��z �S�ʈ�����c�ʈ����RC�����ɂ��킹�Ď��уf�[�^���擾����
'����      :2008/04/21 �쐬  SETsw Nakada
'          :2011/01/17 �Q�ƍ쐬  tkimura
'          :2011/04/28 Marushita �i\cmmc001\s_cmmc001z.bas ����ړ�
Public Function GetCoeffParams_new2(ByVal CRYNUM$, _
                                    wgtCharge As Long, _
                                    wgtChargeA As Long, _
                                    wgtTop As Double, _
                                    wgtTopCut As Double, _
                                    DM As Double, _
                                    HIKIFLG As Integer) As FUNCTION_RETURN
    Dim sql As String
    Dim rs As OraDynaset
    Dim cryNumA As String       'BC���������ł�A�������i�[����B

    On Error GoTo Err
    GetCoeffParams_new2 = FUNCTION_RETURN_FAILURE
    wgtCharge = 0
    wgtChargeA = 0
    wgtTop = 0#
    wgtTopCut = 0#
    DM = 0#
    
    '' ����`���[�W�A�d�ʁiTOP�j�A�g�b�v�J�b�g�d�ʁA�������a�̕��ϒl �擾
    sql = " SELECT C1.SUICHARGE, C1.WGHTTOC1, C1.PUTCUTWC1, "
    sql = sql & " (C1.DIA1C1 + C1.DIA2C1 + C1.DIA3C1) / 3.0 AS DM "
    sql = sql & " FROM XSDC1 C1 "
    sql = sql & " WHERE C1.XTALC1 = '" & CRYNUM & "'"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount > 0 Then
        wgtCharge = rs("SUICHARGE")       ''����`���[�W
        wgtTop = rs("WGHTTOC1")           ''�d�ʁiTOP�j
        wgtTopCut = rs("PUTCUTWC1")       ''�g�b�v�J�b�g�d��
        DM = rs("DM")                     ''�������a(���ϒl)
    End If
    rs.Close
    
    '�����ԍ���9����BorC�Ȃ��BC�����ƂȂ�B
    If Mid(CRYNUM, 9, 1) = "B" Or Mid(CRYNUM, 9, 1) = "C" Then
        HIKIFLG = "2"       'BC����
    Else
        HIKIFLG = "1"       '�ʏ�
    End If
    
    '���̂��Ƃ�wgtChargeA�����߂�K�v������B(HIKIFLG="2"�̂Ƃ��̂�)
    If HIKIFLG = "2" Then
        cryNumA = Mid(CRYNUM, 1, 8) & "A" & Mid(CRYNUM, 10, 3)      '�����ԍ���9���ڂ�A�ɂ���B
        sql = " SELECT C1.SUICHARGE "
        sql = sql & " FROM XSDC1 C1 "
        sql = sql & " WHERE C1.XTALC1 = '" & cryNumA & "'"

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        If rs.RecordCount > 0 Then
            wgtChargeA = rs("SUICHARGE")       ''����`���[�W
        End If
        rs.Close
    End If
    
    Set rs = Nothing
    GetCoeffParams_new2 = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    On Error GoTo 0
    Exit Function

Err:
    Resume proc_exit
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/17 tkimura ADD START
'�T�v      :�ΐ͌W�������߂�B
'���Ұ�    :�ϐ���        ,IO ,�^                       ,����
'          :d             ,I ,type_Coefficient_new2     ,�����R,������㗦�v�Z�\����
'          :�߂�l        ,O  ,Double                   ,�ΐ͌W��
'����      :
'����      :2001/06/23�@���� �M�Ɓ@�쐬
'          :2011/01/17  �Q�ƍ쐬  tkimura
'          :2011/04/28  Marushita �i\cmmc001\s_cmmc001z.bas ����ړ�
Public Function CoefficientCalculation_new2(ByRef d As type_Coefficient_new2) As Double
    
    On Error GoTo Err
    
    CoefficientCalculation_new2 = Log(d.BOTRES / (d.TOPRES * 1)) / Log((1 - d.GT) / (1 - d.GB)) + 1
    
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    CoefficientCalculation_new2 = -9999
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/17 tkimura ADD START
'�T�v       :���グ�����v�Z����B
'���Ұ�     :�ϐ���         ,IO ,�^                         ,����
'           :d              ,I  ,type_Coefficient_new2      ,�����R,������㗦�v�Z�\����
'           :�߂�l         ,O  ,Double                     ,�ʒu���㗦
'����       :
'����       :2011/01/17 tkimura
'           :2011/04/28 Marushita �i\cmmc001\s_cmmc001z.bas ����ړ�
Public Function HikiageCalculation(ByRef d As type_Coefficient_new2) As Double
    Dim result As Double

    '�ʏ�
    If d.HIKIFLG = "1" Then
        result = (d.DUNMENSEKI * d.SMPLPOS * HIJU_SILICONE + d.TOPWEIGHT) / (d.CHARGEWEIGHT)
    'BC����
    Else
        result = (d.DUNMENSEKI * d.SMPLPOS * HIJU_SILICONE + d.TOPWEIGHT + d.CHARGEWEIGHTA - d.CHARGEWEIGHT) / (d.CHARGEWEIGHTA)
    End If
    
    HikiageCalculation = result
    
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/17 tkimura ADD START
'�T�v       :���R�l���v�Z����B
'���Ұ�     :�ϐ���         ,IO ,�^                         ,����
'           :d              ,I  ,type_Coefficient_new2      ,�����R,������㗦�v�Z�\����
'           :�߂�l         ,O ,Double                      ,���R�l
'����       :
'����       :2011/01/17 tkimura
'           :2011/04/28 Marushita �i\cmmc001\s_cmmc001z.bas ����ړ�
Public Function StandardResCalculation(d As type_Coefficient_new2) As Double
    
    On Error GoTo Err
        
    StandardResCalculation = d.TOPRES * (1 - d.GT) ^ (d.Henseki - 1)
        
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    StandardResCalculation = -9999
    
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/17 tkimura ADD START
'�T�v       :����ʒu���R�l���v�Z����B
'���Ұ�     :�ϐ���         ,IO ,�^                         ,����
'           :d              ,I  ,type_Coefficient_new2      ,�����R,������㗦�v�Z�\����
'           :�߂�l         ,O ,Double                      ,����ʒu���R�l
'����       :
'����       :2011/01/17 tkimura
'           :2011/04/28 Marushita �i\cmmc001\s_cmmc001z.bas ����ړ�
Public Function SuiteiResCalculation(ByRef d As type_Coefficient_new2) As Double
    
    On Error GoTo Err
        
    SuiteiResCalculation = d.KIJUNTEIKOU / (1 - d.SUITEIHIKIRITU) ^ (d.Henseki - 1)
        
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    SuiteiResCalculation = -9999
    
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
'�T�v       :�O����̌����ԍ����擾����B
'���Ұ�     :�ϐ���         ,IO ,�^                         ,����
'           :sCrynum        ,I  ,String             �@      ,���݌����ԍ�
'           :sP_Crynum      ,O  ,String             �@      ,�O����̌����ԍ�
'           :�߂�l         ,O  ,FUNCTION_RETURN            ,
'����       :
'����       :2011/04/25 Marushita
Public Function GetPreCrynum(ByVal sCryNum As String, ByRef sP_Crynum As String) As FUNCTION_RETURN
    
    On Error GoTo Err
    GetPreCrynum = FUNCTION_RETURN_FAILURE
                
    Dim sCrynum9 As String
    Dim iCrynum9 As Integer
            
    sP_Crynum = ""
    
    'sCrynum��9���ڂ��Z�b�g
    sCrynum9 = Mid(sCryNum, 9, 1)
    
    '"B"��菬�����L���̓G���[�ŕԂ�
    If sCrynum9 < "B" Then
        Exit Function
    Else
        iCrynum9 = Asc(sCrynum9)
        sP_Crynum = Mid(sCryNum, 1, 8) & Chr(iCrynum9 - 1) & Mid(sCryNum, 10, 3)
    End If
    
    GetPreCrynum = FUNCTION_RETURN_SUCCESS
    
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    
End Function
'=================================================================================

'=================================================================================
'�T�v       :�O����̌��������擾����B
'���Ұ�     :�ϐ���         ,IO ,�^                         ,����
'           :sCRYNUM        ,I  ,String    ,�O���㌋���ԍ�
'           :dWgtTop        ,O  ,Double    ,�g�b�v�d�ʎ��ђl
'           :dDM            ,O  ,Double    ,���a�P�`�R�̕���
'           :lLentk         ,O  ,Long      ,���㒷
'           :dwgtTA         ,O  ,Double    ,�e�C���d�ʎ��ђl
'           :�߂�l         ,O  ,FUNCTION_RETURN            ,
'����       :
'����       :2011/04/25 Marushita
Public Function GetPreXSDC1(ByVal sCryNum As String, _
                            ByRef dwgtTop As Double, _
                            ByRef dDM As Double, _
                            ByRef lLenTK As Long, _
                            ByRef dwgtTA As Double) As FUNCTION_RETURN
    
    On Error GoTo proc_err
    
    Dim sql As String
    Dim rs As OraDynaset
        
    GetPreXSDC1 = FUNCTION_RETURN_FAILURE
    
    sql = " SELECT WGHTTOC1, LENTKC1, WGHTTAC1, "
    sql = sql & " (DIA1C1 + DIA2C1 + DIA3C1) / 3.0 AS DM "
    sql = sql & " FROM XSDC1 "
    sql = sql & " WHERE XTALC1 = '" & sCryNum & "'"
        
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount > 0 Then
        dwgtTop = rs("WGHTTOC1")           ''�d�ʁiTOP�j
        dDM = rs("DM")                     ''�������a(���ϒl)
        lLenTK = rs("LENTKC1")             ''���㒷
        dwgtTA = rs("WGHTTAC1")            ''�d�ʁiTAIL�j
    End If
    rs.Close
    
    GetPreXSDC1 = FUNCTION_RETURN_SUCCESS
    
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
    GetPreXSDC1 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
    
End Function
'=================================================================================

'=================================================================================
'�T�v       :�e�[�u���ɍ��ږ������݂��邩�`�F�b�N����B
'���Ұ�     :�ϐ���         ,IO ,�^                         ,����
'           :sTblName       ,I  ,String             �@      ,���݌����ԍ�
'           :sFldName       ,I  ,String             �@      ,�O����̌����ԍ�
'           :�߂�l         ,O  ,FUNCTION_RETURN            ,
'����       :
'����       :2011/04/25 Marushita
Public Function FieldCheck(ByVal sTblName As String, ByVal sFldName As String) As FUNCTION_RETURN
Dim rs As OraDynaset
Dim fld As OraField

    FieldCheck = FUNCTION_RETURN_FAILURE
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err

    Set rs = OraDB.CreateDynaset("select * from " & sTblName, ORADYN_NO_BLANKSTRIP)
    For Each fld In rs.Fields
        ''���̃t�B�[���h�����o�^�Ȃ�A����l�œo�^����
        If fld.Name = sFldName Then
            FieldCheck = FUNCTION_RETURN_SUCCESS
            rs.Close
            Exit Function
        End If
    Next
    rs.Close

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    FieldCheck = FUNCTION_RETURN_FAILURE
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit

End Function

'=================================================================================
'�T�v      :�i�ԃ}�X�^���(TBCME036)�̒��Ԕ����P�ʂ��擾����B
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :HIN             , I  ,tFullHinban       , �i��(�S�i�ԍ\����)
'          :iUnit           , I  ,Integer           , ���Ԕ����P��
'      �@�@:�߂�l          , O  ,FUNCTION_RETURN�@ , ����
'����      :TBCME036.MCUTUNIT�̒l���擾����B
'����      :2011/06/24 Marushita
Private Function getTBCME036(ByRef HIN As tFullHinban, _
                                        ByRef iUnit As Integer, ByRef sFlg As Integer) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim sndFlg  As String       '���M�t���O
    Dim i       As Integer
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCME036"
    
    getTBCME036 = FUNCTION_RETURN_FAILURE

    iUnit = 0
    '���璆�Ԕ����P�ʂ��擾����B
    Set rs = Nothing
    sql = ""
    sql = sql & "SELECT" & vbCrLf
    sql = sql & " NVL(MSMPTANIMAI,0) as MSMPTANIMAI," & vbCrLf     '���Ԕ����P��
    sql = sql & " NVL(MSMPFLG,'0')   as MSMPFLG" & vbCrLf          '���Ԕ����t���O
    sql = sql & "FROM" & vbCrLf
    sql = sql & " TBCME036" & vbCrLf
    sql = sql & "WHERE" & vbCrLf
    sql = sql & " HINBAN ='" & HIN.hinban & "' AND" & vbCrLf
    sql = sql & " MNOREVNO =" & HIN.mnorevno & " AND" & vbCrLf
    sql = sql & " FACTORY ='" & HIN.factory & "' AND" & vbCrLf
    sql = sql & " OPECOND ='" & HIN.opecond & "'" & vbCrLf
            
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        getTBCME036 = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    
    iUnit = rs("MSMPTANIMAI")
    sFlg = rs("MSMPFLG")
        
    getTBCME036 = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCME036 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'=================================================================================
'�T�v      :XSDC2�̃u���b�N�J�n�ʒu���擾����B
'���Ұ��@�@:�ϐ���          , IO , �^               , ����
'          :sCrynum         , I  ,String            , �i��(�S�i�ԍ\����)
'      �@�@:�߂�l          , O  ,Integer        �@ , �u���b�N�J�n�ʒu
'����      :XSDC2.INPOS�̒l���擾����B
'����      :2011/07/11 Marushita
Private Function getXSDC2Pos(ByVal sCryNum As String) As Integer
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim sndFlg  As String       '���M�t���O
    Dim i       As Integer
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getXSDC2Pos"
    
    getXSDC2Pos = 0

    'XSDC2����INPOS���擾����B
    Set rs = Nothing
    sql = ""
    sql = sql & "SELECT" & vbCrLf
    sql = sql & "NVL(INPOSC2,0) INPOSC2" & vbCrLf        '���Ԕ����P��
    sql = sql & "FROM" & vbCrLf
    sql = sql & " XSDC2" & vbCrLf
    sql = sql & "WHERE" & vbCrLf
    sql = sql & " CRYNUMC2 ='" & sCryNum & "' AND" & vbCrLf
    sql = sql & " LIVKC2 <> '1' " & vbCrLf
            
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    
    getXSDC2Pos = CInt(rs.Fields("INPOSC2"))
    
proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getXSDC2Pos = 0
    gErr.HandleError
    Resume proc_exit
End Function


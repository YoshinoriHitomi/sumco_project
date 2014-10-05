Attribute VB_Name = "s_cmbc055_SQL"
Option Explicit
'
'================================================
' DB�A�N�Z�X�֐�
' ��`���e: TBCMB019 (FRS�Z�����)
' �Q�Ɓ@�@: 060200_�S�e�[�u��
'================================================

'------------------------------------------------
' ���[�U��`�^�̐錾
'------------------------------------------------
Public Type typ_cmjc001j_Disp
    GOUKI       As String * 3       ' ���@
    INPDATE     As Date             ' ���t
    FTIROIL     As Double           ' FTIR�iOi��)
    FTIROIM     As Double           ' FTIR�iOi���j
    FTIROIH     As Double           ' FTIR�iOi���j
    MS1OIL      As Double           ' ����T���v��1�iOi��)
    MS1OIM      As Double           ' ����T���v��1�iOi���j
    MS1OIH      As Double           ' ����T���v��1�iOi���j
    MS2OIL      As Double           ' ����T���v��2�iOi��)
    MS2OIM      As Double           ' ����T���v��2�iOi���j
    MS2OIH      As Double           ' ����T���v��2�iOi���j
    MS3OIL      As Double           ' ����T���v��3�iOi��)
    MS3OIM      As Double           ' ����T���v��3�iOi���j
    MS3OIH      As Double           ' ����T���v��3�iOi���j
    MS4OIL      As Double           ' ����T���v��4�iOi��)
    MS4OIM      As Double           ' ����T���v��4�iOi���j
    MS4OIH      As Double           ' ����T���v��4�iOi���j
    MS5OIL      As Double           ' ����T���v��5�iOi��)
    MS5OIM      As Double           ' ����T���v��5�iOi���j
    MS5OIH      As Double           ' ����T���v��5�iOi���j
    MSAVEOIL    As Double           ' ���蕽�ρiOi��)
    MSAVEOIM    As Double           ' ���蕽�ρiOi���j
    MSAVEOIH    As Double           ' ���蕽�ρiOi���j
    MSSGOIL     As Double           ' ����ЁiOi��)
    MSSGOIM     As Double           ' ����ЁiOi���j
    MSSGOIH     As Double           ' ����ЁiOi���j
    MSPSGOIL    As Double           ' ����AVE+�ЁiOi��)
    MSPSGOIM    As Double           ' ����AVE+�ЁiOi���j
    MSPSGOIH    As Double           ' ����AVE+�ЁiOi���j
    MSNSGOIL    As Double           ' ����AVE-�ЁiOi��)
    MSNSGOIM    As Double           ' ����AVE-�ЁiOi���j
    MSNSGOIH    As Double           ' ����AVE-�ЁiOi���j
    MINOIL      As Double           ' MIN�iOi��)
    MINOIM      As Double           ' MIN�iOi���j
    MINOIH      As Double           ' MIN�iOi���j
    MAXOIL      As Double           ' MAX�iOi��)
    MAXOIM      As Double           ' MAX�iOi���j
    MAXOIH      As Double           ' MAX�iOi���j
    SGCK1OIL    As Double           ' ��ck�T���v��1�iOi��)
    SGCK1OIM    As Double           ' ��ck�T���v��1�iOi���j
    SGCK1OIH    As Double           ' ��ck�T���v��1�iOi���j
    SGCK2OIL    As Double           ' ��ck�T���v��2�iOi��)
    SGCK2OIM    As Double           ' ��ck�T���v��2�iOi���j
    SGCK2OIH    As Double           ' ��ck�T���v��2�iOi���j
    SGCK3OIL    As Double           ' ��ck�T���v��3�iOi��)
    SGCK3OIM    As Double           ' ��ck�T���v��3�iOi���j
    SGCK3OIH    As Double           ' ��ck�T���v��3�iOi���j
    SGCK4OIL    As Double           ' ��ck�T���v��4�iOi��)
    SGCK4OIM    As Double           ' ��ck�T���v��4�iOi���j
    SGCK4OIH    As Double           ' ��ck�T���v��4�iOi���j
    SGCK5OIL    As Double           ' ��ck�T���v��5�iOi��)
    SGCK5OIM    As Double           ' ��ck�T���v��5�iOi���j
    SGCK5OIH    As Double           ' ��ck�T���v��5�iOi���j
    SGCKDOIL    As Double           ' ��ck�f�[�^���iOi��)
    SGCKDOIM    As Double           ' ��ck�f�[�^���iOi���j
    SGCKDOIH    As Double           ' ��ck�f�[�^���iOi���j
    SGCKAOIL    As Double           ' ��ck���ρiOi��)
    SGCKAAOIM   As Double           ' ��ck���ρiOi���j
    SGCKAOIH    As Double           ' ��ck���ρiOi���j
    SGNOIL      As Double           ' ��ck�ЁiOi��)
    SGNOIM      As Double           ' ��ck�ЁiOi���j
    SGNOIH      As Double           ' ��ck�ЁiOi���j
    FTIRKOIL    As Double           ' FTIR���Z�iOi��)
    FTIRKOIM    As Double           ' FTIR���Z�iOi���j
    FTIRKOIH    As Double           ' FTIR���Z�iOi���j
    EFFECTTM    As Integer          ' �L������
    YCOEF       As Double           ' �e�s�h�q���Z���i�x�ؕЁj
    XCOEF       As Double           ' �e�s�h�q���Z���i�w�W���j
    RSQUARE     As Double           ' �q�Q��
    SGCKST      As Double           ' �Д���
    SGCKOIL     As String * 1       ' �Д���iOi��)
    SGCKOIM     As String * 1       ' �Д���iOi���j
    SGCKOIH     As String * 1       ' �Д���iOi���j
    FTIRCKST    As Double           ' FTIR���Z����
    FTIRCKOIL   As String * 1       ' FTIR���Z����iOi��)
    FTIRCKOIM   As String * 1       ' FTIR���Z����iOi���j
    FTIRCKOIH   As String * 1       ' FTIR���Z����iOi���j
    MS6OIL      As Double           ' ����T���v��6�iOi��)
    MS6OIM      As Double           ' ����T���v��6�iOi���j
    MS6OIH      As Double           ' ����T���v��6�iOi���j
    SGCK6OIL    As Double           ' ��ck�T���v��6�iOi��)
    SGCK6OIM    As Double           ' ��ck�T���v��6�iOi���j
    SGCK6OIH    As Double           ' ��ck�T���v��6�iOi���j
    CVOIL       As Double           ' CV(%)�iOi��)
    CVOIM       As Double           ' CV(%)�iOi���j
    CVOIH       As Double           ' CV(%)�iOi���j
  '  TSTAFFID As String * 8          ' �o�^�Ј�ID
  '  REGDATE As Date                 ' �o�^���t
  '  KSTAFFID As String * 8          ' �X�V�Ј�ID
  '  UPDDATE As Date                 ' �X�V���t
  '  SENDFLAG As String * 1          ' ���M�t���O
  '  SENDDATE As Date                ' ���M���t
End Type

'''''------------------------------------------------
''''' DB�A�N�Z�X�֐�
'''''------------------------------------------------
''''
'''''�T�v      :�e�[�u���uTBCMB019�v��������ɂ��������R�[�h�𒊏o����
'''''���Ұ�    :�ϐ���        ,IO ,�^           ,����
'''''          :record        ,O  ,typ_cmjc001j_Disp ,���o���R�[�h
'''''          :GOUK          ,I  ,String       ,�u���@�v(SQL�̒��o����)
'''''          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'''''����      :�u���@�v=�����ŁA���u���t�v���ŐV�̃f�[�^�𒊏o����
'''''����      :2001/06/20�쐬�@����
''''Public Function DBDRV_Getcmjc001j_Disp(record As typ_cmjc001j_Disp, GOUK$) As FUNCTION_RETURN
''''Dim sql As String       'SQL�S��
''''Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
''''Dim sqlWhere As String  'SQL��WHERE����
''''Dim sqlGroup As String  'SQL��GROUP����
''''Dim rs As OraDynaset    'RecordSet
''''Dim recCnt As Long      '���R�[�h��
''''Dim i As Long
''''
''''    DBDRV_Getcmjc001j_Disp = FUNCTION_RETURN_FAILURE
''''
''''    ''SQL��g�ݗ��Ă�
''''
''''    '�G���[�n���h���̐ݒ�
''''    On Error GoTo proc_err
''''    gErr.Push "s_cmzcF_cmjc001j_SQL.bas -- Function DBDRV_Getcmjc001j_Disp"
''''
''''    sqlBase = "Select GOUKI, MAX(INPDATE) ""INPDATE"", FTIRFZI, FTIRCZH, FTIRCZC, MS1FZ, MS1CZ1, MS1CZ2, MS2FZ, MS2CZ1, MS2CZ2, MS3FZ, MS3CZ1," & _
''''              " MS3CZ2, MS4FZ, MS4CZ1, MS4CZ2, MS5FZ, MS5CZ1, MS5CZ2, MSAVEFZ, MSAVECZ1, MSAVECZ2, MSSGFZ, MSSGCZ1, MSSGCZ2," & _
''''              " MSPSGFZ, MSPSGCZ1, MSPSGCZ2, MSNSGFZ, MSNSGCZ1, MSNSGCZ2, MINFZ, MINCZ1, MINCZ2, MAXFZ, MAXCZ1, MAXCZ2, SGCK1FZ," & _
''''              " SGCK1CZ1, SGCK1CZ2, SGCK2FZ, SGCK2CZ1, SGCK2CZ2, SGCK3FZ, SGCK3CZ1, SGCK3CZ2, SGCK4FZ, SGCK4CZ1, SGCK4CZ2," & _
''''              " SGCK5FZ, SGCK5CZ1, SGCK5CZ2, SGCKDFZ, SGCKDCZ1, SGCKDCZ2, SGCKAFZ, SGCKAACZ1, SGCKACZ2, SGNFZ, SGNCZ1, SGNCZ2," & _
''''              " FTIRFZ, FTIRCZ1, FTIRCZ2, EFFECTTM, YCOEF, XCOEF, RSQUARE "
''''    sqlBase = sqlBase & "From TBCMB019"
''''    ''���o����(�����NO)�̎��o��
''''    sqlWhere = "WHERE(GOUKI=" & GOUK & ") "
''''    sqlGroup = "GROUP BY GOUKI"
''''    sql = sqlBase & sqlWhere & sqlGroup
''''
''''    ''�f�[�^�𒊏o����
''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''''    If rs Is Nothing Then
''''        ReDim records(0)
''''        DBDRV_Getcmjc001j_Disp = FUNCTION_RETURN_FAILURE
''''        GoTo proc_exit
''''    End If
''''
''''    ''���o���ʂ��i�[����
''''    With record
''''        .GOUKI = rs("GOUKI")             ' ���@
''''        .INPDATE = rs("INPDATE")         ' ���t
''''        .FTIRFZI = rs("FTIRFZI")         ' FTIR�iFZ)
''''        .FTIRCZH = rs("FTIRCZH")         ' FTIR�iCZ���j
''''        .FTIRCZC = rs("FTIRCZC")         ' FTIR�iCZ���j
''''        .MS1FZ = rs("MS1FZ")             ' ����T���v��1�iFZ)
''''        .MS1CZ1 = rs("MS1CZ1")           ' ����T���v��1�iCZ-1)
''''        .MS1CZ2 = rs("MS1CZ2")           ' ����T���v��1�iCZ-2)
''''        .MS2FZ = rs("MS2FZ")             ' ����T���v��2�iFZ)
''''        .MS2CZ1 = rs("MS2CZ1")           ' ����T���v��2�iCZ-1)
''''        .MS2CZ2 = rs("MS2CZ2")           ' ����T���v��2�iCZ-2)
''''        .MS3FZ = rs("MS3FZ")             ' ����T���v��3�iFZ)
''''        .MS3CZ1 = rs("MS3CZ1")           ' ����T���v��3�iCZ-1)
''''        .MS3CZ2 = rs("MS3CZ2")           ' ����T���v��3�iCZ-2)
''''        .MS4FZ = rs("MS4FZ")             ' ����T���v��4�iFZ)
''''        .MS4CZ1 = rs("MS4CZ1")           ' ����T���v��4�iCZ-1)
''''        .MS4CZ2 = rs("MS4CZ2")           ' ����T���v��4�iCZ-2)
''''        .MS5FZ = rs("MS5FZ")             ' ����T���v��5�iFZ)
''''        .MS5CZ1 = rs("MS5CZ1")           ' ����T���v��5�iCZ-1)
''''        .MS5CZ2 = rs("MS5CZ2")           ' ����T���v��5�iCZ-2)
''''        .MSAVEFZ = rs("MSAVEFZ")         ' ���蕽�ρiFZ�j
''''        .MSAVECZ1 = rs("MSAVECZ1")       ' ���蕽�ρiCZ-1�j
''''        .MSAVECZ2 = rs("MSAVECZ2")       ' ���蕽�ρiCZ-2�j
''''        .MSSGFZ = rs("MSSGFZ")           ' ����ЁiFZ�j
''''        .MSSGCZ1 = rs("MSSGCZ1")         ' ����ЁiCZ-1�j
''''        .MSSGCZ2 = rs("MSSGCZ2")         ' ����ЁiCZ-2�j
''''        .MSPSGFZ = rs("MSPSGFZ")         ' ����AVE+�ЁiFZ�j
''''        .MSPSGCZ1 = rs("MSPSGCZ1")       ' ����AVE+�ЁiCZ-1�j
''''        .MSPSGCZ2 = rs("MSPSGCZ2")       ' ����AVE+�ЁiCZ-2�j
''''        .MSNSGFZ = rs("MSNSGFZ")         ' ����AVE-�ЁiFZ�j
''''        .MSNSGCZ1 = rs("MSNSGCZ1")       ' ����AVE-�ЁiCZ-1�j
''''        .MSNSGCZ2 = rs("MSNSGCZ2")       ' ����AVE-�ЁiCZ-2�j
''''        .MINFZ = rs("MINFZ")             ' MIN�iFZ�j
''''        .MINCZ1 = rs("MINCZ1")           ' MIN�iCZ-1�j
''''        .MINCZ2 = rs("MINCZ2")           ' MIN�iCZ-2�j
''''        .MAXFZ = rs("MAXFZ")             ' MAX�iFZ�j
''''        .MAXCZ1 = rs("MAXCZ1")           ' MAX�iCZ-1�j
''''        .MAXCZ2 = rs("MAXCZ2")           ' MAX�iCZ-2�j
''''        .SGCK1FZ = rs("SGCK1FZ")         ' ��ck�T���v��1�iFZ)
''''        .SGCK1CZ1 = rs("SGCK1CZ1")       ' ��ck�T���v��1�iCZ-1)
''''        .SGCK1CZ2 = rs("SGCK1CZ2")       ' ��ck�T���v��1�iCZ-2)
''''        .SGCK2FZ = rs("SGCK2FZ")         ' ��ck�T���v��2�iFZ)
''''        .SGCK2CZ1 = rs("SGCK2CZ1")       ' ��ck�T���v��2�iCZ-1)
''''        .SGCK2CZ2 = rs("SGCK2CZ2")       ' ��ck�T���v��2�iCZ-2)
''''        .SGCK3FZ = rs("SGCK3FZ")         ' ��ck�T���v��3�iFZ)
''''        .SGCK3CZ1 = rs("SGCK3CZ1")       ' ��ck�T���v��3�iCZ-1)
''''        .SGCK3CZ2 = rs("SGCK3CZ2")       ' ��ck�T���v��3�iCZ-2)
''''        .SGCK4FZ = rs("SGCK4FZ")         ' ��ck�T���v��4�iFZ)
''''        .SGCK4CZ1 = rs("SGCK4CZ1")       ' ��ck�T���v��4�iCZ-1)
''''        .SGCK4CZ2 = rs("SGCK4CZ2")       ' ��ck�T���v��4�iCZ-2)
''''        .SGCK5FZ = rs("SGCK5FZ")         ' ��ck�T���v��5�iFZ)
''''        .SGCK5CZ1 = rs("SGCK5CZ1")       ' ��ck�T���v��5�iCZ-1)
''''        .SGCK5CZ2 = rs("SGCK5CZ2")       ' ��ck�T���v��5�iCZ-2)
''''        .SGCKDFZ = rs("SGCKDFZ")         ' ��ck�f�[�^���iFZ�j
''''        .SGCKDCZ1 = rs("SGCKDCZ1")       ' ��ck�f�[�^���iCZ-1�j
''''        .SGCKDCZ2 = rs("SGCKDCZ2")       ' ��ck�f�[�^���iCZ-2�j
''''        .SGCKAFZ = rs("SGCKAFZ")         ' ��ck���ρiFZ�j
''''        .SGCKAACZ1 = rs("SGCKAACZ1")     ' ��ck���ρiCZ-1�j
''''        .SGCKACZ2 = rs("SGCKACZ2")       ' ��ck���ρiCZ-2�j
''''        .SGNFZ = rs("SGNFZ")             ' ��ck�ЁiFZ�j
''''        .SGNCZ1 = rs("SGNCZ1")           ' ��ck�� CZ-1�j
''''        .SGNCZ2 = rs("SGNCZ2")           ' ��ck�ЁiCZ-2�j
''''        .FTIRFZ = rs("FTIRFZ")           ' FTIR���Z�iFZ�j
''''        .FTIRCZ1 = rs("FTIRCZ1")         ' FTIR���Z�iCZ-1�j
''''        .FTIRCZ2 = rs("FTIRCZ2")         ' FTIR���Z�iCZ-2�j
''''        .EFFECTTM = rs("EFFECTTM")       ' �L������
''''        .YCOEF = rs("YCOEF")             ' �e�s�h�q���Z���i�x�ؕЁj
''''        .XCOEF = rs("XCOEF")             ' �e�s�h�q���Z���i�w�W���j
''''        .RSQUARE = rs("RSQUARE")         ' �q�Q��
''''    End With
''''    rs.Close
''''
''''    DBDRV_Getcmjc001j_Disp = FUNCTION_RETURN_SUCCESS
''''
''''proc_exit:
''''    '�I��
''''    gErr.Pop
''''    Exit Function
''''
''''proc_err:
''''    '�G���[�n���h��
''''    Debug.Print "====== Error SQL ======"
''''    Debug.Print sql
''''    gErr.HandleError
''''    Resume proc_exit
''''End Function


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�����œn���ꂽ���R�[�h��TBCMB019�ɒǉ�����
'���Ұ�    :�ϐ���        ,IO ,�^            ,����
'          :record        ,I  ,typ_cmjc001j_Disp ,���o���R�[�h
'          :TSTAFFID      ,I  ,String       ,�o�^�Ј�ID
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :
Public Function DBDRV_Getcmjc001j_Exec(record As typ_cmjc001j_Disp, TSTAFFID$) As FUNCTION_RETURN
    Dim sql As String           'SQL�S��
    Dim SetDate  As Variant     '���͓��t

    DBDRV_Getcmjc001j_Exec = FUNCTION_RETURN_FAILURE
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001j_SQL.bas -- Function DBDRV_Getcmjc001j_Exec"

    SetDate = Format$(record.INPDATE, "yyyy-mm-dd hh:mm:ss")
  
    ''SQL��g�ݗ��Ă�
    sql = "Insert into TBCMB019 ("
    sql = sql & "  GOUKI"                   '' ���@
    sql = sql & ", INPDATE"                 '' ���t
    sql = sql & ", FTIROIL"                 '' FTIR�iOi��)
    sql = sql & ", FTIROIM"                 '' FTIR�iOi��)
    sql = sql & ", FTIROIH"                 '' FTIR�iOi��)
    sql = sql & ", MS1OIL"                  '' ����T���v��1�iOi��)
    sql = sql & ", MS1OIM"                  '' ����T���v��1�iOi��)
    sql = sql & ", MS1OIH"                  '' ����T���v��1�iOi��)
    sql = sql & ", MS2OIL"                  '' ����T���v��2�iOi��)
    sql = sql & ", MS2OIM"                  '' ����T���v��2�iOi��)
    sql = sql & ", MS2OIH"                  '' ����T���v��2�iOi��)
    sql = sql & ", MS3OIL"                  '' ����T���v��3�iOi��)
    sql = sql & ", MS3OIM"                  '' ����T���v��3�iOi��)
    sql = sql & ", MS3OIH"                  '' ����T���v��3�iOi��)
    sql = sql & ", MS4OIL"                  '' ����T���v��4�iOi��)
    sql = sql & ", MS4OIM"                  '' ����T���v��4�iOi��)
    sql = sql & ", MS4OIH"                  '' ����T���v��4�iOi��)
    sql = sql & ", MS5OIL"                  '' ����T���v��5�iOi��)
    sql = sql & ", MS5OIM"                  '' ����T���v��5�iOi��)
    sql = sql & ", MS5OIH"                  '' ����T���v��5�iOi��)
    sql = sql & ", MSAVEOIL"                '' ���蕽�ρiOi��)
    sql = sql & ", MSAVEOIM"                '' ���蕽�ρiOi��)
    sql = sql & ", MSAVEOIH"                '' ���蕽�ρiOi��)
    sql = sql & ", MSSGOIL"                 '' ����ЁiOi��)
    sql = sql & ", MSSGOIM"                 '' ����ЁiOi��)
    sql = sql & ", MSSGOIH"                 '' ����ЁiOi��)
    sql = sql & ", MSPSGOIL"                '' ����AVE+�ЁiOi��)
    sql = sql & ", MSPSGOIM"                '' ����AVE+�ЁiOi��)
    sql = sql & ", MSPSGOIH"                '' ����AVE+�ЁiOi��)
    sql = sql & ", MSNSGOIL"                '' ����AVE-�ЁiOi��)
    sql = sql & ", MSNSGOIM"                '' ����AVE-�ЁiOi��)
    sql = sql & ", MSNSGOIH"                '' ����AVE-�ЁiOi��)
    sql = sql & ", MINOIL"                  '' MIN�iOi��)
    sql = sql & ", MINOIM"                  '' MIN�iOi��)
    sql = sql & ", MINOIH"                  '' MIN�iOi��)
    sql = sql & ", MAXOIL"                  '' MAX�iOi��)
    sql = sql & ", MAXOIM"                  '' MAX�iOi��)
    sql = sql & ", MAXOIH"                  '' MAX�iOi��)
    sql = sql & ", SGCK1OIL"                '' ��ck�T���v��1�iOi��)
    sql = sql & ", SGCK1OIM"                '' ��ck�T���v��1�iOi��)
    sql = sql & ", SGCK1OIH"                '' ��ck�T���v��1�iOi��)
    sql = sql & ", SGCK2OIL"                '' ��ck�T���v��2�iOi��)
    sql = sql & ", SGCK2OIM"                '' ��ck�T���v��2�iOi��)
    sql = sql & ", SGCK2OIH"                '' ��ck�T���v��2�iOi��)
    sql = sql & ", SGCK3OIL"                '' ��ck�T���v��3�iOi��)
    sql = sql & ", SGCK3OIM"                '' ��ck�T���v��3�iOi��)
    sql = sql & ", SGCK3OIH"                '' ��ck�T���v��3�iOi��)
    sql = sql & ", SGCK4OIL"                '' ��ck�T���v��4�iOi��)
    sql = sql & ", SGCK4OIM"                '' ��ck�T���v��4�iOi��)
    sql = sql & ", SGCK4OIH"                '' ��ck�T���v��4�iOi��)
    sql = sql & ", SGCK5OIL"                '' ��ck�T���v��5�iOi��)
    sql = sql & ", SGCK5OIM"                '' ��ck�T���v��5�iOi��)
    sql = sql & ", SGCK5OIH"                '' ��ck�T���v��5�iOi��)
    sql = sql & ", SGCKDOIL"                '' ��ck�f�[�^���iOi��)
    sql = sql & ", SGCKDOIM"                '' ��ck�f�[�^���iOi��)
    sql = sql & ", SGCKDOIH"                '' ��ck�f�[�^���iOi��)
    sql = sql & ", SGCKAOIL"                '' ��ck���ρiOi��)
    sql = sql & ", SGCKAAOIM"               '' ��ck���ρiOi��)
    sql = sql & ", SGCKAOIH"                '' ��ck���ρiOi��)
    sql = sql & ", SGNOIL"                  '' ��ck�ЁiOi��)
    sql = sql & ", SGNOIM"                  '' ��ck�ЁiOi��)
    sql = sql & ", SGNOIH"                  '' ��ck�ЁiOi��)
    sql = sql & ", FTIRKOIL"                '' FTIR���Z�iOi��)
    sql = sql & ", FTIRKOIM"                '' FTIR���Z�iOi��)
    sql = sql & ", FTIRKOIH"                '' FTIR���Z�iOi��)
    sql = sql & ", EFFECTTM"                '' �L������
    sql = sql & ", YCOEF"                   '' �e�s�h�q���Z���i�x�ؕЁj
    sql = sql & ", XCOEF"                   '' �e�s�h�q���Z���i�w�W���j
    sql = sql & ", RSQUARE"                 '' �q�Q��
    sql = sql & ", SGCKST"                  '' �Д���
    sql = sql & ", SGCKOIL"                 '' �Д���iOi��)
    sql = sql & ", SGCKOIM"                 '' �Д���iOi��)
    sql = sql & ", SGCKOIH"                 '' �Д���iOi��)
    sql = sql & ", FTIRCKST"                '' FTIR���Z����
    sql = sql & ", FTIRCKOIL"               '' FTIR���Z����iOi��)
    sql = sql & ", FTIRCKOIM"               '' FTIR���Z����iOi��)
    sql = sql & ", FTIRCKOIH"               '' FTIR���Z����iOi��)
    sql = sql & ", MS6OIL"                  '' ����T���v��6�iOi��)
    sql = sql & ", MS6OIM"                  '' ����T���v��6�iOi��)
    sql = sql & ", MS6OIH"                  '' ����T���v��6�iOi��)
    sql = sql & ", SGCK6OIL"                '' ��ck�T���v��6�iOi��)
    sql = sql & ", SGCK6OIM"                '' ��ck�T���v��6�iOi��)
    sql = sql & ", SGCK6OIH"                '' ��ck�T���v��6�iOi��)
    sql = sql & ", CVOIL"                   '' CV�iOi��)
    sql = sql & ", CVOIM"                   '' CV�iOi��)
    sql = sql & ", CVOIH"                   '' CV�iOi��)
    sql = sql & ", TSTAFFID"                '' �o�^�Ј�ID
    sql = sql & ", REGDATE"                 '' �o�^���t
    sql = sql & ", KSTAFFID"                '' �X�V�Ј�ID
    sql = sql & ", UPDDATE"                 '' �X�V���t
    sql = sql & ", SENDFLAG"                '' ���M�t���O
    sql = sql & ", SENDDATE"                '' ���M���t
    sql = sql & ")"
    
    sql = sql & "Values("
    sql = sql & "'" & record.GOUKI & "'"                                        '' ���@
    sql = sql & ", " & "TO_DATE('" & SetDate & "','YYYY-MM-DD hh24:mi:ss')"     '' ���t
    sql = sql & ", " & record.FTIROIL                                           '' FTIR�iOi��)
    sql = sql & ", " & record.FTIROIM                                           '' FTIR�iOi��)
    sql = sql & ", " & record.FTIROIH                                           '' FTIR�iOi��)
    sql = sql & ", " & record.MS1OIL                                            '' ����T���v��1�iOi��)
    sql = sql & ", " & record.MS1OIM                                            '' ����T���v��1�iOi��)
    sql = sql & ", " & record.MS1OIH                                            '' ����T���v��1�iOi��)
    sql = sql & ", " & record.MS2OIL                                            '' ����T���v��2�iOi��)
    sql = sql & ", " & record.MS2OIM                                            '' ����T���v��2�iOi��)
    sql = sql & ", " & record.MS2OIH                                            '' ����T���v��2�iOi��)
    sql = sql & ", " & record.MS3OIL                                            '' ����T���v��3�iOi��)
    sql = sql & ", " & record.MS3OIM                                            '' ����T���v��3�iOi��)
    sql = sql & ", " & record.MS3OIH                                            '' ����T���v��3�iOi��)
    sql = sql & ", " & record.MS4OIL                                            '' ����T���v��4�iOi��)
    sql = sql & ", " & record.MS4OIM                                            '' ����T���v��4�iOi��)
    sql = sql & ", " & record.MS4OIH                                            '' ����T���v��4�iOi��)
    sql = sql & ", " & record.MS5OIL                                            '' ����T���v��5�iOi��)
    sql = sql & ", " & record.MS5OIM                                            '' ����T���v��5�iOi��)
    sql = sql & ", " & record.MS5OIH                                            '' ����T���v��5�iOi��)
    sql = sql & ", " & record.MSAVEOIL                                          '' ���蕽�ρiOi��)
    sql = sql & ", " & record.MSAVEOIM                                          '' ���蕽�ρiOi��)
    sql = sql & ", " & record.MSAVEOIH                                          '' ���蕽�ρiOi��)
    sql = sql & ", " & record.MSSGOIL                                           '' ����ЁiOi��)
    sql = sql & ", " & record.MSSGOIM                                           '' ����ЁiOi��)
    sql = sql & ", " & record.MSSGOIH                                           '' ����ЁiOi��)
    sql = sql & ", " & record.MSPSGOIL                                          '' ����AVE+�ЁiOi��)
    sql = sql & ", " & record.MSPSGOIM                                          '' ����AVE+�ЁiOi��)
    sql = sql & ", " & record.MSPSGOIH                                          '' ����AVE+�ЁiOi��)
    sql = sql & ", " & record.MSNSGOIL                                          '' ����AVE-�ЁiOi��)
    sql = sql & ", " & record.MSNSGOIM                                          '' ����AVE-�ЁiOi��)
    sql = sql & ", " & record.MSNSGOIH                                          '' ����AVE-�ЁiOi��)
    sql = sql & ", " & record.MINOIL                                            '' MIN�iOi��)
    sql = sql & ", " & record.MINOIM                                            '' MIN�iOi��)
    sql = sql & ", " & record.MINOIH                                            '' MIN�iOi��)
    sql = sql & ", " & record.MAXOIL                                            '' MAX�iOi��)
    sql = sql & ", " & record.MAXOIM                                            '' MAX�iOi��)
    sql = sql & ", " & record.MAXOIH                                            '' MAX�iOi��)
    sql = sql & ", " & record.SGCK1OIL                                          '' ��ck�T���v��1�iOi��)
    sql = sql & ", " & record.SGCK1OIM                                          '' ��ck�T���v��1�iOi��)
    sql = sql & ", " & record.SGCK1OIH                                          '' ��ck�T���v��1�iOi��)
    sql = sql & ", " & record.SGCK2OIL                                          '' ��ck�T���v��2�iOi��)
    sql = sql & ", " & record.SGCK2OIM                                          '' ��ck�T���v��2�iOi��)
    sql = sql & ", " & record.SGCK2OIH                                          '' ��ck�T���v��2�iOi��)
    sql = sql & ", " & record.SGCK3OIL                                          '' ��ck�T���v��3�iOi��)
    sql = sql & ", " & record.SGCK3OIM                                          '' ��ck�T���v��3�iOi��)
    sql = sql & ", " & record.SGCK3OIH                                          '' ��ck�T���v��3�iOi��)
    sql = sql & ", " & record.SGCK4OIL                                          '' ��ck�T���v��4�iOi��)
    sql = sql & ", " & record.SGCK4OIM                                          '' ��ck�T���v��4�iOi��)
    sql = sql & ", " & record.SGCK4OIH                                          '' ��ck�T���v��4�iOi��)
    sql = sql & ", " & record.SGCK5OIL                                          '' ��ck�T���v��5�iOi��)
    sql = sql & ", " & record.SGCK5OIM                                          '' ��ck�T���v��5�iOi��)
    sql = sql & ", " & record.SGCK5OIH                                          '' ��ck�T���v��5�iOi��)
    sql = sql & ", " & record.SGCKDOIL                                          '' ��ck�f�[�^���iOi��)
    sql = sql & ", " & record.SGCKDOIM                                          '' ��ck�f�[�^���iOi��)
    sql = sql & ", " & record.SGCKDOIH                                          '' ��ck�f�[�^���iOi��)
    sql = sql & ", " & record.SGCKAOIL                                          '' ��ck���ρiOi��)
    sql = sql & ", " & record.SGCKAAOIM                                         '' ��ck���ρiOi��)
    sql = sql & ", " & record.SGCKAOIH                                          '' ��ck���ρiOi��)
    sql = sql & ", " & record.SGNOIL                                            '' ��ck�ЁiOi��)
    sql = sql & ", " & record.SGNOIM                                            '' ��ck�ЁiOi��)
    sql = sql & ", " & record.SGNOIH                                            '' ��ck�ЁiOi��)
    sql = sql & ", " & record.FTIRKOIL                                          '' FTIR���Z�iOi��)
    sql = sql & ", " & record.FTIRKOIM                                          '' FTIR���Z�iOi��)
    sql = sql & ", " & record.FTIRKOIH                                          '' FTIR���Z�iOi��)
    sql = sql & ", " & record.EFFECTTM                                          '' �L������
    sql = sql & ", " & record.YCOEF                                             '' �e�s�h�q���Z���i�x�ؕЁj
    sql = sql & ", " & record.XCOEF                                             '' �e�s�h�q���Z���i�w�W���j
    sql = sql & ", " & record.RSQUARE                                           '' �q�Q��
    sql = sql & ", " & record.SGCKST                                            '' �Д���
    sql = sql & ", '" & record.SGCKOIL & "'"                                    '' �Д���iOi��)
    sql = sql & ", '" & record.SGCKOIM & "'"                                    '' �Д���iOi��)
    sql = sql & ", '" & record.SGCKOIH & "'"                                    '' �Д���iOi��)
    sql = sql & ", " & record.FTIRCKST                                          '' FTIR���Z����
    sql = sql & ", '" & record.FTIRCKOIL & "'"                                  '' FTIR���Z����iOi��)
    sql = sql & ", '" & record.FTIRCKOIM & "'"                                  '' FTIR���Z����iOi��)
    sql = sql & ", '" & record.FTIRCKOIH & "'"                                  '' FTIR���Z����iOi��)
    sql = sql & ", " & record.MS6OIL                                            '' ����T���v��6�iOi��)
    sql = sql & ", " & record.MS6OIM                                            '' ����T���v��6�iOi��)
    sql = sql & ", " & record.MS6OIH                                            '' ����T���v��6�iOi��)
    sql = sql & ", " & record.SGCK6OIL                                          '' ��ck�T���v��6�iOi��)
    sql = sql & ", " & record.SGCK6OIM                                          '' ��ck�T���v��6�iOi��)
    sql = sql & ", " & record.SGCK6OIH                                          '' ��ck�T���v��6�iOi��)
    sql = sql & ", " & record.CVOIL                                             '' CV�iOi��)
    sql = sql & ", " & record.CVOIM                                             '' CV�iOi��)
    sql = sql & ", " & record.CVOIH                                             '' CV�iOi��)
    sql = sql & ", '" & TSTAFFID & "'"                                          '' �o�^�Ј�ID
    sql = sql & ", SYSDATE"                                                     '' �o�^���t
    sql = sql & ", ' '"                                                         '' �X�V�Ј�ID
    sql = sql & ", SYSDATE"                                                     '' �X�V���t
    sql = sql & ", '0'"                                                         '' ���M�t���O
    sql = sql & ", SYSDATE"                                                     '' ���M���t
    sql = sql & ")"
  
    '' ��SQL�̎��s
    OraDB.ExecuteSQL (sql)

    DBDRV_Getcmjc001j_Exec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�f�[�^�ϊ����s��
'���Ұ�    :�ϐ���        ,IO ,�^                ,����
'          :tblLeft       ,IO   ,typ_TBCMB019      ,�e�[�u���f�[�^�P
'          :tblRight      ,IO   ,typ_cmjc001j_Disp ,�e�[�u���f�[�^�Q
'          :bFlg          ,I   ,Boolean           ,TRUE:�����P�f�[�^�������Q�f�[�^�ւ̕ϊ�  FALSE:�����P�f�[�^�������Q�f�[�^�ւ̕ϊ�
'����      :
Public Sub ConvDate_F_cmjc001j_a(tblLeft As typ_TBCMB019, tblRight As typ_cmjc001j_Disp, bFlg As Boolean)
    
    If bFlg = True Then
        With tblRight
            .GOUKI = tblLeft.GOUKI
            .INPDATE = tblLeft.INPDATE
            .FTIROIL = tblLeft.FTIROIL
            .FTIROIM = tblLeft.FTIROIM
            .FTIROIH = tblLeft.FTIROIH
            .MS1OIL = tblLeft.MS1OIL
            .MS1OIM = tblLeft.MS1OIM
            .MS1OIH = tblLeft.MS1OIH
            .MS2OIL = tblLeft.MS2OIL
            .MS2OIM = tblLeft.MS2OIM
            .MS2OIH = tblLeft.MS2OIH
            .MS3OIL = tblLeft.MS3OIL
            .MS3OIM = tblLeft.MS3OIM
            .MS3OIH = tblLeft.MS3OIH
            .MS4OIL = tblLeft.MS4OIL
            .MS4OIM = tblLeft.MS4OIM
            .MS4OIH = tblLeft.MS4OIH
            .MS5OIL = tblLeft.MS5OIL
            .MS5OIM = tblLeft.MS5OIM
            .MS5OIH = tblLeft.MS5OIH
            .MSAVEOIL = tblLeft.MSAVEOIL
            .MSAVEOIM = tblLeft.MSAVEOIM
            .MSAVEOIH = tblLeft.MSAVEOIH
            .MSSGOIL = tblLeft.MSSGOIL
            .MSSGOIM = tblLeft.MSSGOIM
            .MSSGOIH = tblLeft.MSSGOIH
            .MSPSGOIL = tblLeft.MSPSGOIL
            .MSPSGOIM = tblLeft.MSPSGOIM
            .MSPSGOIH = tblLeft.MSPSGOIH
            .MSNSGOIL = tblLeft.MSNSGOIL
            .MSNSGOIM = tblLeft.MSNSGOIM
            .MSNSGOIH = tblLeft.MSNSGOIH
            .MINOIL = tblLeft.MINOIL
            .MINOIM = tblLeft.MINOIM
            .MINOIH = tblLeft.MINOIH
            .MAXOIL = tblLeft.MAXOIL
            .MAXOIM = tblLeft.MAXOIM
            .MAXOIH = tblLeft.MAXOIH
            .SGCK1OIL = tblLeft.SGCK1OIL
            .SGCK1OIM = tblLeft.SGCK1OIM
            .SGCK1OIH = tblLeft.SGCK1OIH
            .SGCK2OIL = tblLeft.SGCK2OIL
            .SGCK2OIM = tblLeft.SGCK2OIM
            .SGCK2OIH = tblLeft.SGCK2OIH
            .SGCK3OIL = tblLeft.SGCK3OIL
            .SGCK3OIM = tblLeft.SGCK3OIM
            .SGCK3OIH = tblLeft.SGCK3OIH
            .SGCK4OIL = tblLeft.SGCK4OIL
            .SGCK4OIM = tblLeft.SGCK4OIM
            .SGCK4OIH = tblLeft.SGCK4OIH
            .SGCK5OIL = tblLeft.SGCK5OIL
            .SGCK5OIM = tblLeft.SGCK5OIM
            .SGCK5OIH = tblLeft.SGCK5OIH
            .SGCKDOIL = tblLeft.SGCKDOIL
            .SGCKDOIM = tblLeft.SGCKDOIM
            .SGCKDOIH = tblLeft.SGCKDOIH
            .SGCKAOIL = tblLeft.SGCKAOIL
            .SGCKAAOIM = tblLeft.SGCKAAOIM
            .SGCKAOIH = tblLeft.SGCKAOIH
            .SGNOIL = tblLeft.SGNOIL
            .SGNOIM = tblLeft.SGNOIM
            .SGNOIH = tblLeft.SGNOIH
            .FTIRKOIL = tblLeft.FTIRKOIL
            .FTIRKOIM = tblLeft.FTIRKOIM
            .FTIRKOIH = tblLeft.FTIRKOIH
            .EFFECTTM = tblLeft.EFFECTTM
            .YCOEF = tblLeft.YCOEF
            .XCOEF = tblLeft.XCOEF
            .RSQUARE = tblLeft.RSQUARE
            .SGCKST = tblLeft.SGCKST
            .SGCKOIL = tblLeft.SGCKOIL
            .SGCKOIM = tblLeft.SGCKOIM
            .SGCKOIH = tblLeft.SGCKOIH
            .FTIRCKST = tblLeft.FTIRCKST
            .FTIRCKOIL = tblLeft.FTIRCKOIL
            .FTIRCKOIM = tblLeft.FTIRCKOIM
            .FTIRCKOIH = tblLeft.FTIRCKOIH
            .MS6OIL = tblLeft.MS6OIL
            .MS6OIM = tblLeft.MS6OIM
            .MS6OIH = tblLeft.MS6OIH
            .SGCK6OIL = tblLeft.SGCK6OIL
            .SGCK6OIM = tblLeft.SGCK6OIM
            .SGCK6OIH = tblLeft.SGCK6OIH
            .CVOIL = tblLeft.CVOIL
            .CVOIM = tblLeft.CVOIM
            .CVOIH = tblLeft.CVOIH
        
        End With
    Else
        With tblLeft
            .GOUKI = tblRight.GOUKI
            .INPDATE = tblRight.INPDATE
            .FTIROIL = tblRight.FTIROIL
            .FTIROIM = tblRight.FTIROIM
            .FTIROIH = tblRight.FTIROIH
            .MS1OIL = tblRight.MS1OIL
            .MS1OIM = tblRight.MS1OIM
            .MS1OIH = tblRight.MS1OIH
            .MS2OIL = tblRight.MS2OIL
            .MS2OIM = tblRight.MS2OIM
            .MS2OIH = tblRight.MS2OIH
            .MS3OIL = tblRight.MS3OIL
            .MS3OIM = tblRight.MS3OIM
            .MS3OIH = tblRight.MS3OIH
            .MS4OIL = tblRight.MS4OIL
            .MS4OIM = tblRight.MS4OIM
            .MS4OIH = tblRight.MS4OIH
            .MS5OIL = tblRight.MS5OIL
            .MS5OIM = tblRight.MS5OIM
            .MS5OIH = tblRight.MS5OIH
            .MSAVEOIL = tblRight.MSAVEOIL
            .MSAVEOIM = tblRight.MSAVEOIM
            .MSAVEOIH = tblRight.MSAVEOIH
            .MSSGOIL = tblRight.MSSGOIL
            .MSSGOIM = tblRight.MSSGOIM
            .MSSGOIH = tblRight.MSSGOIH
            .MSPSGOIL = tblRight.MSPSGOIL
            .MSPSGOIM = tblRight.MSPSGOIM
            .MSPSGOIH = tblRight.MSPSGOIH
            .MSNSGOIL = tblRight.MSNSGOIL
            .MSNSGOIM = tblRight.MSNSGOIM
            .MSNSGOIH = tblRight.MSNSGOIH
            .MINOIL = tblRight.MINOIL
            .MINOIM = tblRight.MINOIM
            .MINOIH = tblRight.MINOIH
            .MAXOIL = tblRight.MAXOIL
            .MAXOIM = tblRight.MAXOIM
            .MAXOIH = tblRight.MAXOIH
            .SGCK1OIL = tblRight.SGCK1OIL
            .SGCK1OIM = tblRight.SGCK1OIM
            .SGCK1OIH = tblRight.SGCK1OIH
            .SGCK2OIL = tblRight.SGCK2OIL
            .SGCK2OIM = tblRight.SGCK2OIM
            .SGCK2OIH = tblRight.SGCK2OIH
            .SGCK3OIL = tblRight.SGCK3OIL
            .SGCK3OIM = tblRight.SGCK3OIM
            .SGCK3OIH = tblRight.SGCK3OIH
            .SGCK4OIL = tblRight.SGCK4OIL
            .SGCK4OIM = tblRight.SGCK4OIM
            .SGCK4OIH = tblRight.SGCK4OIH
            .SGCK5OIL = tblRight.SGCK5OIL
            .SGCK5OIM = tblRight.SGCK5OIM
            .SGCK5OIH = tblRight.SGCK5OIH
            .SGCKDOIL = tblRight.SGCKDOIL
            .SGCKDOIM = tblRight.SGCKDOIM
            .SGCKDOIH = tblRight.SGCKDOIH
            .SGCKAOIL = tblRight.SGCKAOIL
            .SGCKAAOIM = tblRight.SGCKAAOIM
            .SGCKAOIH = tblRight.SGCKAOIH
            .SGNOIL = tblRight.SGNOIL
            .SGNOIM = tblRight.SGNOIM
            .SGNOIH = tblRight.SGNOIH
            .FTIRKOIL = tblRight.FTIRKOIL
            .FTIRKOIM = tblRight.FTIRKOIM
            .FTIRKOIH = tblRight.FTIRKOIH
            .EFFECTTM = tblRight.EFFECTTM
            .YCOEF = tblRight.YCOEF
            .XCOEF = tblRight.XCOEF
            .RSQUARE = tblRight.RSQUARE
            .SGCKST = tblRight.SGCKST
            .SGCKOIL = tblRight.SGCKOIL
            .SGCKOIM = tblRight.SGCKOIM
            .SGCKOIH = tblRight.SGCKOIH
            .FTIRCKST = tblRight.FTIRCKST
            .FTIRCKOIL = tblRight.FTIRCKOIL
            .FTIRCKOIM = tblRight.FTIRCKOIM
            .FTIRCKOIH = tblRight.FTIRCKOIH
            .MS6OIL = tblRight.MS6OIL
            .MS6OIM = tblRight.MS6OIM
            .MS6OIH = tblRight.MS6OIH
            .SGCK6OIL = tblRight.SGCK6OIL
            .SGCK6OIM = tblRight.SGCK6OIM
            .SGCK6OIH = tblRight.SGCK6OIH
            .CVOIL = tblRight.CVOIL
            .CVOIM = tblRight.CVOIM
            .CVOIH = tblRight.CVOIH
        
        End With
    End If

End Sub

'''''------------------------------------------------
''''' DB�A�N�Z�X�֐�
'''''------------------------------------------------
''''
'''''�T�v      :�e�[�u���uTBCMB019�v��������ɂ��������R�[�h�𒊏o����
'''''���Ұ�    :�ϐ���        ,IO ,�^           ,����
'''''          :records()     ,O  ,typ_TBCMB019 ,���o���R�[�h
'''''          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'''''          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'''''          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'''''����      :
'''''����      :2001/08/24�쐬�@�쑺
''''Public Function DBDRV_GetTBCMB019(records() As typ_TBCMB019, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
''''Dim sql As String       'SQL�S��
''''Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
''''Dim rs As OraDynaset    'RecordSet
''''Dim recCnt As Long      '���R�[�h��
''''Dim i As Long
''''
''''    ''SQL��g�ݗ��Ă�
''''    sqlBase = "Select GOUKI, INPDATE, FTIRFZI, FTIRCZH, FTIRCZC, MS1FZ, MS1CZ1, MS1CZ2, MS2FZ, MS2CZ1, MS2CZ2, MS3FZ, MS3CZ1," & _
''''              " MS3CZ2, MS4FZ, MS4CZ1, MS4CZ2, MS5FZ, MS5CZ1, MS5CZ2, MSAVEFZ, MSAVECZ1, MSAVECZ2, MSSGFZ, MSSGCZ1, MSSGCZ2," & _
''''              " MSPSGFZ, MSPSGCZ1, MSPSGCZ2, MSNSGFZ, MSNSGCZ1, MSNSGCZ2, MINFZ, MINCZ1, MINCZ2, MAXFZ, MAXCZ1, MAXCZ2, SGCK1FZ," & _
''''              " SGCK1CZ1, SGCK1CZ2, SGCK2FZ, SGCK2CZ1, SGCK2CZ2, SGCK3FZ, SGCK3CZ1, SGCK3CZ2, SGCK4FZ, SGCK4CZ1, SGCK4CZ2," & _
''''              " SGCK5FZ, SGCK5CZ1, SGCK5CZ2, SGCKDFZ, SGCKDCZ1, SGCKDCZ2, SGCKAFZ, SGCKAACZ1, SGCKACZ2, SGNFZ, SGNCZ1, SGNCZ2," & _
''''              " FTIRFZ, FTIRCZ1, FTIRCZ2, EFFECTTM, YCOEF, XCOEF, RSQUARE, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG," & _
''''              " SENDDATE "
''''    sqlBase = sqlBase & "From TBCMB019"
''''    sql = sqlBase
''''    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
''''        sql = sql & " " & sqlWhere & " " & sqlOrder
''''    End If
''''
''''    ''�f�[�^�𒊏o����
''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
''''    If rs Is Nothing Then
''''        ReDim records(0)
''''        DBDRV_GetTBCMB019 = FUNCTION_RETURN_FAILURE
''''        Exit Function
''''    End If
''''
''''    ''���o���ʂ��i�[����
''''    recCnt = rs.RecordCount
''''    ReDim records(recCnt)
''''    For i = 1 To recCnt
''''        With records(i)
''''            .GOUKI = rs("GOUKI")             ' ���@
''''            .INPDATE = rs("INPDATE")         ' ���t
''''            .FTIRFZI = rs("FTIRFZI")         ' FTIR�iFZ)
''''            .FTIRCZH = rs("FTIRCZH")         ' FTIR�iCZ���j
''''            .FTIRCZC = rs("FTIRCZC")         ' FTIR�iCZ���j
''''            .MS1FZ = rs("MS1FZ")             ' ����T���v��1�iFZ)
''''            .MS1CZ1 = rs("MS1CZ1")           ' ����T���v��1�iCZ-1)
''''            .MS1CZ2 = rs("MS1CZ2")           ' ����T���v��1�iCZ-2)
''''            .MS2FZ = rs("MS2FZ")             ' ����T���v��2�iFZ)
''''            .MS2CZ1 = rs("MS2CZ1")           ' ����T���v��2�iCZ-1)
''''            .MS2CZ2 = rs("MS2CZ2")           ' ����T���v��2�iCZ-2)
''''            .MS3FZ = rs("MS3FZ")             ' ����T���v��3�iFZ)
''''            .MS3CZ1 = rs("MS3CZ1")           ' ����T���v��3�iCZ-1)
''''            .MS3CZ2 = rs("MS3CZ2")           ' ����T���v��3�iCZ-2)
''''            .MS4FZ = rs("MS4FZ")             ' ����T���v��4�iFZ)
''''            .MS4CZ1 = rs("MS4CZ1")           ' ����T���v��4�iCZ-1)
''''            .MS4CZ2 = rs("MS4CZ2")           ' ����T���v��4�iCZ-2)
''''            .MS5FZ = rs("MS5FZ")             ' ����T���v��5�iFZ)
''''            .MS5CZ1 = rs("MS5CZ1")           ' ����T���v��5�iCZ-1)
''''            .MS5CZ2 = rs("MS5CZ2")           ' ����T���v��5�iCZ-2)
''''            .MSAVEFZ = rs("MSAVEFZ")         ' ���蕽�ρiFZ�j
''''            .MSAVECZ1 = rs("MSAVECZ1")       ' ���蕽�ρiCZ-1�j
''''            .MSAVECZ2 = rs("MSAVECZ2")       ' ���蕽�ρiCZ-2�j
''''            .MSSGFZ = rs("MSSGFZ")           ' ����ЁiFZ�j
''''            .MSSGCZ1 = rs("MSSGCZ1")         ' ����ЁiCZ-1�j
''''            .MSSGCZ2 = rs("MSSGCZ2")         ' ����ЁiCZ-2�j
''''            .MSPSGFZ = rs("MSPSGFZ")         ' ����AVE+�ЁiFZ�j
''''            .MSPSGCZ1 = rs("MSPSGCZ1")       ' ����AVE+�ЁiCZ-1�j
''''            .MSPSGCZ2 = rs("MSPSGCZ2")       ' ����AVE+�ЁiCZ-2�j
''''            .MSNSGFZ = rs("MSNSGFZ")         ' ����AVE-�ЁiFZ�j
''''            .MSNSGCZ1 = rs("MSNSGCZ1")       ' ����AVE-�ЁiCZ-1�j
''''            .MSNSGCZ2 = rs("MSNSGCZ2")       ' ����AVE-�ЁiCZ-2�j
''''            .MINFZ = rs("MINFZ")             ' MIN�iFZ�j
''''            .MINCZ1 = rs("MINCZ1")           ' MIN�iCZ-1�j
''''            .MINCZ2 = rs("MINCZ2")           ' MIN�iCZ-2�j
''''            .MAXFZ = rs("MAXFZ")             ' MAX�iFZ�j
''''            .MAXCZ1 = rs("MAXCZ1")           ' MAX�iCZ-1�j
''''            .MAXCZ2 = rs("MAXCZ2")           ' MAX�iCZ-2�j
''''            .SGCK1FZ = rs("SGCK1FZ")         ' ��ck�T���v��1�iFZ)
''''            .SGCK1CZ1 = rs("SGCK1CZ1")       ' ��ck�T���v��1�iCZ-1)
''''            .SGCK1CZ2 = rs("SGCK1CZ2")       ' ��ck�T���v��1�iCZ-2)
''''            .SGCK2FZ = rs("SGCK2FZ")         ' ��ck�T���v��2�iFZ)
''''            .SGCK2CZ1 = rs("SGCK2CZ1")       ' ��ck�T���v��2�iCZ-1)
''''            .SGCK2CZ2 = rs("SGCK2CZ2")       ' ��ck�T���v��2�iCZ-2)
''''            .SGCK3FZ = rs("SGCK3FZ")         ' ��ck�T���v��3�iFZ)
''''            .SGCK3CZ1 = rs("SGCK3CZ1")       ' ��ck�T���v��3�iCZ-1)
''''            .SGCK3CZ2 = rs("SGCK3CZ2")       ' ��ck�T���v��3�iCZ-2)
''''            .SGCK4FZ = rs("SGCK4FZ")         ' ��ck�T���v��4�iFZ)
''''            .SGCK4CZ1 = rs("SGCK4CZ1")       ' ��ck�T���v��4�iCZ-1)
''''            .SGCK4CZ2 = rs("SGCK4CZ2")       ' ��ck�T���v��4�iCZ-2)
''''            .SGCK5FZ = rs("SGCK5FZ")         ' ��ck�T���v��5�iFZ)
''''            .SGCK5CZ1 = rs("SGCK5CZ1")       ' ��ck�T���v��5�iCZ-1)
''''            .SGCK5CZ2 = rs("SGCK5CZ2")       ' ��ck�T���v��5�iCZ-2)
''''            .SGCKDFZ = rs("SGCKDFZ")         ' ��ck�f�[�^���iFZ�j
''''            .SGCKDCZ1 = rs("SGCKDCZ1")       ' ��ck�f�[�^���iCZ-1�j
''''            .SGCKDCZ2 = rs("SGCKDCZ2")       ' ��ck�f�[�^���iCZ-2�j
''''            .SGCKAFZ = rs("SGCKAFZ")         ' ��ck���ρiFZ�j
''''            .SGCKAACZ1 = rs("SGCKAACZ1")     ' ��ck���ρiCZ-1�j
''''            .SGCKACZ2 = rs("SGCKACZ2")       ' ��ck���ρiCZ-2�j
''''            .SGNFZ = rs("SGNFZ")             ' ��ck�ЁiFZ�j
''''            .SGNCZ1 = rs("SGNCZ1")           ' ��ck�� CZ-1�j
''''            .SGNCZ2 = rs("SGNCZ2")           ' ��ck�ЁiCZ-2�j
''''            .FTIRFZ = rs("FTIRFZ")           ' FTIR���Z�iFZ�j
''''            .FTIRCZ1 = rs("FTIRCZ1")         ' FTIR���Z�iCZ-1�j
''''            .FTIRCZ2 = rs("FTIRCZ2")         ' FTIR���Z�iCZ-2�j
''''            .EFFECTTM = rs("EFFECTTM")       ' �L������
''''            .YCOEF = rs("YCOEF")             ' �e�s�h�q���Z���i�x�ؕЁj
''''            .XCOEF = rs("XCOEF")             ' �e�s�h�q���Z���i�w�W���j
''''            .RSQUARE = rs("RSQUARE")         ' �q�Q��
''''            .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
''''            .REGDATE = rs("REGDATE")         ' �o�^���t
''''            .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
''''            .UPDDATE = rs("UPDDATE")         ' �X�V���t
''''            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
''''            .SENDDATE = rs("SENDDATE")       ' ���M���t
''''        End With
''''        rs.MoveNext
''''    Next
''''    rs.Close
''''
''''    DBDRV_GetTBCMB019 = FUNCTION_RETURN_SUCCESS
''''End Function

'///////////////////////////////////////////////////
' @(f)
' �@�\    : �Д����擾
'
' �Ԃ�l  : True  - ����
' �@�@�@    False - ���s
'
' ������  : sSigCode  - �Д���
' �@�@�@  : sFtirCode - FTIR���Z����
' �@�@�@  : sR2Code   - R2�攻��
'
' �@�\����:
'///////////////////////////////////////////////////
Public Function GetSigChkCode(Optional ByRef sSigCode As String _
                            , Optional ByRef sFtirCode As String _
                            , Optional ByRef sR2Code As String _
                            ) As Boolean
    Dim dbIsMine    As Boolean
    Dim sSql        As String
    Dim objRs       As Object
    
    GetSigChkCode = False
    sSigCode = ""
    sFtirCode = ""
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzc055_SQL.bas -- Function GetSigChkCode"

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''�r�p�k���쐬
    sSql = ""
    sSql = sSql & "SELECT NVL(kcode01a9, ' ')"   '0:�Д���
    sSql = sSql & "      ,NVL(kcode02a9, ' ')"   '1:FTIR���Z����
    sSql = sSql & "      ,NVL(kcode03a9, ' ')"   '2:R2�攻��
    sSql = sSql & "  FROM koda9"
    sSql = sSql & " WHERE sysca9 = 'X'"
    sSql = sSql & "   AND shuca9 = '19'"
    sSql = sSql & "   AND codea9 = 'FRS'"
    
    Set objRs = OraDB.CreateDynaset(sSql, ORADYN_DEFAULT)
    
    If objRs.EOF Then
        Call MsgOut(0, "�Д����̃R�[�h���o�^����Ă��܂���", ERR_DISP)
        Exit Function
    End If

    sSigCode = objRs(0)     ''�Д���
    sFtirCode = objRs(1)    ''FTIR���Z����
    sR2Code = objRs(2)      ''R2�攻��
    
    objRs.Close
    
    ''�Д���
    If IsNumeric(sSigCode) = False Then
        Call MsgOut(0, "�Д����̃R�[�h������������܂���", ERR_DISP)
        Exit Function
    End If
    ' -10~100�łȂ��C�܂��͏����_��O�ʈȍ~�̓��͂�����ꍇ�̓G���[
    If Not (-10# < CDbl(sSigCode) And CDbl(sSigCode) < 100#) Then
        Call MsgOut(0, "�Д����̃R�[�h������������܂���", ERR_DISP)
        Exit Function
    End If
    If InStr(1, sSigCode, ".", vbTextCompare) >= 1 Then
        If Len(sSigCode) - InStr(1, sSigCode, ".", vbTextCompare) >= 3 Then
            Call MsgOut(0, "�Д����̃R�[�h������������܂���", ERR_DISP)
            Exit Function
        End If
    End If
    
    ''FTIR���Z����
    If IsNumeric(sFtirCode) = False Then
        Call MsgOut(0, "FTIR���Z�����̃R�[�h������������܂���", ERR_DISP)
        Exit Function
    End If
    ' -10~100�łȂ��C�܂��͏����_��O�ʈȍ~�̓��͂�����ꍇ�̓G���[
    If Not (-10# < CDbl(sFtirCode) And CDbl(sFtirCode) < 100#) Then
        Call MsgOut(0, "FTIR���Z�����̃R�[�h������������܂���", ERR_DISP)
        Exit Function
    End If
    If InStr(1, sFtirCode, ".", vbTextCompare) >= 1 Then
        If Len(sFtirCode) - InStr(1, sFtirCode, ".", vbTextCompare) >= 3 Then
            Call MsgOut(0, "FTIR���Z�����̃R�[�h������������܂���", ERR_DISP)
            Exit Function
        End If
    End If
    
    ''R2�攻��
    If IsNumeric(sR2Code) = False Then
        Call MsgOut(0, "�q2�攻���̃R�[�h������������܂���", ERR_DISP)
        Exit Function
    End If
    
    GetSigChkCode = True        ''����������Ԃ�

proc_exit:
    If dbIsMine Then
        OraDBClose
    End If
    
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
    
End Function

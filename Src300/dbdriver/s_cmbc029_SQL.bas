Attribute VB_Name = "s_cmbc029_SQL"
Option Explicit
'                                     2001/06/20
'================================================
' DB�A�N�Z�X�֐�
' ��`���e: TBCMB014 (GFA�Z�����)
' �Q�Ɓ@�@: 060200_�S�e�[�u��
'================================================

'------------------------------------------------
' ���[�U��`�^�̐錾
'------------------------------------------------
Public Type typ_cmjc001j_Disp
    GOUKI As String * 3             ' ���@
    INPDATE As Date                 ' ���t
    FTIRFZI As Double               ' FTIR�iFZ)
    FTIRCZH As Double               ' FTIR�iCZ���j
    FTIRCZC As Double               ' FTIR�iCZ���j
    MS1FZ As Double                 ' ����T���v��1�iFZ)
    MS1CZ1 As Double                ' ����T���v��1�iCZ-1)
    MS1CZ2 As Double                ' ����T���v��1�iCZ-2)
    MS2FZ As Double                 ' ����T���v��2�iFZ)
    MS2CZ1 As Double                ' ����T���v��2�iCZ-1)
    MS2CZ2 As Double                ' ����T���v��2�iCZ-2)
    MS3FZ As Double                 ' ����T���v��3�iFZ)
    MS3CZ1 As Double                ' ����T���v��3�iCZ-1)
    MS3CZ2 As Double                ' ����T���v��3�iCZ-2)
    MS4FZ As Double                 ' ����T���v��4�iFZ)
    MS4CZ1 As Double                ' ����T���v��4�iCZ-1)
    MS4CZ2 As Double                ' ����T���v��4�iCZ-2)
    MS5FZ As Double                 ' ����T���v��5�iFZ)
    MS5CZ1 As Double                ' ����T���v��5�iCZ-1)
    MS5CZ2 As Double                ' ����T���v��5�iCZ-2)
    MSAVEFZ As Double               ' ���蕽�ρiFZ�j
    MSAVECZ1 As Double              ' ���蕽�ρiCZ-1�j
    MSAVECZ2 As Double              ' ���蕽�ρiCZ-2�j
    MSSGFZ As Double                ' ����ЁiFZ�j
    MSSGCZ1 As Double               ' ����ЁiCZ-1�j
    MSSGCZ2 As Double               ' ����ЁiCZ-2�j
    MSPSGFZ As Double               ' ����AVE+�ЁiFZ�j
    MSPSGCZ1 As Double              ' ����AVE+�ЁiCZ-1�j
    MSPSGCZ2 As Double              ' ����AVE+�ЁiCZ-2�j
    MSNSGFZ As Double               ' ����AVE-�ЁiFZ�j
    MSNSGCZ1 As Double              ' ����AVE-�ЁiCZ-1�j
    MSNSGCZ2 As Double              ' ����AVE-�ЁiCZ-2�j
    MINFZ As Double                 ' MIN�iFZ�j
    MINCZ1 As Double                ' MIN�iCZ-1�j
    MINCZ2 As Double                ' MIN�iCZ-2�j
    MAXFZ As Double                 ' MAX�iFZ�j
    MAXCZ1 As Double                ' MAX�iCZ-1�j
    MAXCZ2 As Double                ' MAX�iCZ-2�j
    SGCK1FZ As Double               ' ��ck�T���v��1�iFZ)
    SGCK1CZ1 As Double              ' ��ck�T���v��1�iCZ-1)
    SGCK1CZ2 As Double              ' ��ck�T���v��1�iCZ-2)
    SGCK2FZ As Double               ' ��ck�T���v��2�iFZ)
    SGCK2CZ1 As Double              ' ��ck�T���v��2�iCZ-1)
    SGCK2CZ2 As Double              ' ��ck�T���v��2�iCZ-2)
    SGCK3FZ As Double               ' ��ck�T���v��3�iFZ)
    SGCK3CZ1 As Double              ' ��ck�T���v��3�iCZ-1)
    SGCK3CZ2 As Double              ' ��ck�T���v��3�iCZ-2)
    SGCK4FZ As Double               ' ��ck�T���v��4�iFZ)
    SGCK4CZ1 As Double              ' ��ck�T���v��4�iCZ-1)
    SGCK4CZ2 As Double              ' ��ck�T���v��4�iCZ-2)
    SGCK5FZ As Double               ' ��ck�T���v��5�iFZ)
    SGCK5CZ1 As Double              ' ��ck�T���v��5�iCZ-1)
    SGCK5CZ2 As Double              ' ��ck�T���v��5�iCZ-2)
    SGCKDFZ As Double               ' ��ck�f�[�^���iFZ�j
    SGCKDCZ1 As Double              ' ��ck�f�[�^���iCZ-1�j
    SGCKDCZ2 As Double              ' ��ck�f�[�^���iCZ-2�j
    SGCKAFZ As Double               ' ��ck���ρiFZ�j
    SGCKAACZ1 As Double             ' ��ck���ρiCZ-1�j
    SGCKACZ2 As Double              ' ��ck���ρiCZ-2�j
    SGNFZ As Double                 ' ��ck�ЁiFZ�j
    SGNCZ1 As Double                ' ��ck�� CZ-1�j
    SGNCZ2 As Double                ' ��ck�ЁiCZ-2�j
    FTIRFZ As Double                ' FTIR���Z�iFZ�j
    FTIRCZ1 As Double               ' FTIR���Z�iCZ-1�j
    FTIRCZ2 As Double               ' FTIR���Z�iCZ-2�j
    EFFECTTM As Integer             ' �L������
    YCOEF As Double                 ' �e�s�h�q���Z���i�x�ؕЁj
    XCOEF As Double                 ' �e�s�h�q���Z���i�w�W���j
    RSQUARE As Double               ' �q�Q��
  '  TSTAFFID As String * 8          ' �o�^�Ј�ID
  '  REGDATE As Date                 ' �o�^���t
  '  KSTAFFID As String * 8          ' �X�V�Ј�ID
  '  UPDDATE As Date                 ' �X�V���t
  '  SENDFLAG As String * 1          ' ���M�t���O
  '  SENDDATE As Date                ' ���M���t

'2006/05/22�ǉ�
    SGCKST      As Double           ' �Д���
    SGCKFZ      As String * 1       ' �Д���(FZ)
    SGCKCZ1     As String * 1       ' �Д���(CZ-1)
    SGCKCZ2     As String * 1       ' �Д���(CZ-2)
    FTIRCKST    As Double           ' FTIR���Z����
    FTIRCKFZ    As String * 1       ' FTIR���Z����(FZ)
    FTIRCKCZ1   As String * 1       ' FTIR���Z����(CZ-1)
    FTIRCKCZ2   As String * 1       ' FTIR���Z����(CZ-2)

'2010/03/26�ǉ� SETsw kubota
    MS6FZ       As Double           ' ����T���v��6�iFZ)
    MS6CZ1      As Double           ' ����T���v��6�iCZ-1)
    MS6CZ2      As Double           ' ����T���v��6�iCZ-2)
    SGCK6FZ     As Double           ' ��ck�T���v��6�iFZ)
    SGCK6CZ1    As Double           ' ��ck�T���v��6�iCZ-1)
    SGCK6CZ2    As Double           ' ��ck�T���v��6�iCZ-2)
    CVFZ        As Double           ' CV(%)�iFZ�j
    CVCZ1       As Double           ' CV(%)�iCZ-1�j
    CVCZ2       As Double           ' CV(%)�iCZ-2�j

End Type

'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMB014�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :record        ,O  ,typ_cmjc001j_Disp ,���o���R�[�h
'          :GOUK          ,I  ,String       ,�u���@�v(SQL�̒��o����)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :�u���@�v=�����ŁA���u���t�v���ŐV�̃f�[�^�𒊏o����
'����      :2001/06/20�쐬�@����
Public Function DBDRV_Getcmjc001j_Disp(record As typ_cmjc001j_Disp, GOUK$) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim sqlWhere As String  'SQL��WHERE����
Dim sqlGroup As String  'SQL��GROUP����
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    DBDRV_Getcmjc001j_Disp = FUNCTION_RETURN_FAILURE
    
    ''SQL��g�ݗ��Ă�

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001j_SQL.bas -- Function DBDRV_Getcmjc001j_Disp"

    sqlBase = "Select GOUKI, MAX(INPDATE) ""INPDATE"", FTIRFZI, FTIRCZH, FTIRCZC, MS1FZ, MS1CZ1, MS1CZ2, MS2FZ, MS2CZ1, MS2CZ2, MS3FZ, MS3CZ1," & _
              " MS3CZ2, MS4FZ, MS4CZ1, MS4CZ2, MS5FZ, MS5CZ1, MS5CZ2, MSAVEFZ, MSAVECZ1, MSAVECZ2, MSSGFZ, MSSGCZ1, MSSGCZ2," & _
              " MSPSGFZ, MSPSGCZ1, MSPSGCZ2, MSNSGFZ, MSNSGCZ1, MSNSGCZ2, MINFZ, MINCZ1, MINCZ2, MAXFZ, MAXCZ1, MAXCZ2, SGCK1FZ," & _
              " SGCK1CZ1, SGCK1CZ2, SGCK2FZ, SGCK2CZ1, SGCK2CZ2, SGCK3FZ, SGCK3CZ1, SGCK3CZ2, SGCK4FZ, SGCK4CZ1, SGCK4CZ2," & _
              " SGCK5FZ, SGCK5CZ1, SGCK5CZ2, SGCKDFZ, SGCKDCZ1, SGCKDCZ2, SGCKAFZ, SGCKAACZ1, SGCKACZ2, SGNFZ, SGNCZ1, SGNCZ2," & _
              " FTIRFZ, FTIRCZ1, FTIRCZ2, EFFECTTM, YCOEF, XCOEF, RSQUARE "
    sqlBase = sqlBase & "From TBCMB014"
    ''���o����(�����NO)�̎��o��
    sqlWhere = "WHERE(GOUKI=" & GOUK & ") "
    sqlGroup = "GROUP BY GOUKI"
    sql = sqlBase & sqlWhere & sqlGroup
    
    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_Getcmjc001j_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
    With record
        .GOUKI = rs("GOUKI")             ' ���@
        .INPDATE = rs("INPDATE")         ' ���t
        .FTIRFZI = rs("FTIRFZI")         ' FTIR�iFZ)
        .FTIRCZH = rs("FTIRCZH")         ' FTIR�iCZ���j
        .FTIRCZC = rs("FTIRCZC")         ' FTIR�iCZ���j
        .MS1FZ = rs("MS1FZ")             ' ����T���v��1�iFZ)
        .MS1CZ1 = rs("MS1CZ1")           ' ����T���v��1�iCZ-1)
        .MS1CZ2 = rs("MS1CZ2")           ' ����T���v��1�iCZ-2)
        .MS2FZ = rs("MS2FZ")             ' ����T���v��2�iFZ)
        .MS2CZ1 = rs("MS2CZ1")           ' ����T���v��2�iCZ-1)
        .MS2CZ2 = rs("MS2CZ2")           ' ����T���v��2�iCZ-2)
        .MS3FZ = rs("MS3FZ")             ' ����T���v��3�iFZ)
        .MS3CZ1 = rs("MS3CZ1")           ' ����T���v��3�iCZ-1)
        .MS3CZ2 = rs("MS3CZ2")           ' ����T���v��3�iCZ-2)
        .MS4FZ = rs("MS4FZ")             ' ����T���v��4�iFZ)
        .MS4CZ1 = rs("MS4CZ1")           ' ����T���v��4�iCZ-1)
        .MS4CZ2 = rs("MS4CZ2")           ' ����T���v��4�iCZ-2)
        .MS5FZ = rs("MS5FZ")             ' ����T���v��5�iFZ)
        .MS5CZ1 = rs("MS5CZ1")           ' ����T���v��5�iCZ-1)
        .MS5CZ2 = rs("MS5CZ2")           ' ����T���v��5�iCZ-2)
        .MSAVEFZ = rs("MSAVEFZ")         ' ���蕽�ρiFZ�j
        .MSAVECZ1 = rs("MSAVECZ1")       ' ���蕽�ρiCZ-1�j
        .MSAVECZ2 = rs("MSAVECZ2")       ' ���蕽�ρiCZ-2�j
        .MSSGFZ = rs("MSSGFZ")           ' ����ЁiFZ�j
        .MSSGCZ1 = rs("MSSGCZ1")         ' ����ЁiCZ-1�j
        .MSSGCZ2 = rs("MSSGCZ2")         ' ����ЁiCZ-2�j
        .MSPSGFZ = rs("MSPSGFZ")         ' ����AVE+�ЁiFZ�j
        .MSPSGCZ1 = rs("MSPSGCZ1")       ' ����AVE+�ЁiCZ-1�j
        .MSPSGCZ2 = rs("MSPSGCZ2")       ' ����AVE+�ЁiCZ-2�j
        .MSNSGFZ = rs("MSNSGFZ")         ' ����AVE-�ЁiFZ�j
        .MSNSGCZ1 = rs("MSNSGCZ1")       ' ����AVE-�ЁiCZ-1�j
        .MSNSGCZ2 = rs("MSNSGCZ2")       ' ����AVE-�ЁiCZ-2�j
        .MINFZ = rs("MINFZ")             ' MIN�iFZ�j
        .MINCZ1 = rs("MINCZ1")           ' MIN�iCZ-1�j
        .MINCZ2 = rs("MINCZ2")           ' MIN�iCZ-2�j
        .MAXFZ = rs("MAXFZ")             ' MAX�iFZ�j
        .MAXCZ1 = rs("MAXCZ1")           ' MAX�iCZ-1�j
        .MAXCZ2 = rs("MAXCZ2")           ' MAX�iCZ-2�j
        .SGCK1FZ = rs("SGCK1FZ")         ' ��ck�T���v��1�iFZ)
        .SGCK1CZ1 = rs("SGCK1CZ1")       ' ��ck�T���v��1�iCZ-1)
        .SGCK1CZ2 = rs("SGCK1CZ2")       ' ��ck�T���v��1�iCZ-2)
        .SGCK2FZ = rs("SGCK2FZ")         ' ��ck�T���v��2�iFZ)
        .SGCK2CZ1 = rs("SGCK2CZ1")       ' ��ck�T���v��2�iCZ-1)
        .SGCK2CZ2 = rs("SGCK2CZ2")       ' ��ck�T���v��2�iCZ-2)
        .SGCK3FZ = rs("SGCK3FZ")         ' ��ck�T���v��3�iFZ)
        .SGCK3CZ1 = rs("SGCK3CZ1")       ' ��ck�T���v��3�iCZ-1)
        .SGCK3CZ2 = rs("SGCK3CZ2")       ' ��ck�T���v��3�iCZ-2)
        .SGCK4FZ = rs("SGCK4FZ")         ' ��ck�T���v��4�iFZ)
        .SGCK4CZ1 = rs("SGCK4CZ1")       ' ��ck�T���v��4�iCZ-1)
        .SGCK4CZ2 = rs("SGCK4CZ2")       ' ��ck�T���v��4�iCZ-2)
        .SGCK5FZ = rs("SGCK5FZ")         ' ��ck�T���v��5�iFZ)
        .SGCK5CZ1 = rs("SGCK5CZ1")       ' ��ck�T���v��5�iCZ-1)
        .SGCK5CZ2 = rs("SGCK5CZ2")       ' ��ck�T���v��5�iCZ-2)
        .SGCKDFZ = rs("SGCKDFZ")         ' ��ck�f�[�^���iFZ�j
        .SGCKDCZ1 = rs("SGCKDCZ1")       ' ��ck�f�[�^���iCZ-1�j
        .SGCKDCZ2 = rs("SGCKDCZ2")       ' ��ck�f�[�^���iCZ-2�j
        .SGCKAFZ = rs("SGCKAFZ")         ' ��ck���ρiFZ�j
        .SGCKAACZ1 = rs("SGCKAACZ1")     ' ��ck���ρiCZ-1�j
        .SGCKACZ2 = rs("SGCKACZ2")       ' ��ck���ρiCZ-2�j
        .SGNFZ = rs("SGNFZ")             ' ��ck�ЁiFZ�j
        .SGNCZ1 = rs("SGNCZ1")           ' ��ck�� CZ-1�j
        .SGNCZ2 = rs("SGNCZ2")           ' ��ck�ЁiCZ-2�j
        .FTIRFZ = rs("FTIRFZ")           ' FTIR���Z�iFZ�j
        .FTIRCZ1 = rs("FTIRCZ1")         ' FTIR���Z�iCZ-1�j
        .FTIRCZ2 = rs("FTIRCZ2")         ' FTIR���Z�iCZ-2�j
        .EFFECTTM = rs("EFFECTTM")       ' �L������
        .YCOEF = rs("YCOEF")             ' �e�s�h�q���Z���i�x�ؕЁj
        .XCOEF = rs("XCOEF")             ' �e�s�h�q���Z���i�w�W���j
        .RSQUARE = rs("RSQUARE")         ' �q�Q��
    End With
    rs.Close

    DBDRV_Getcmjc001j_Disp = FUNCTION_RETURN_SUCCESS

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


'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�����œn���ꂽ���R�[�h��TBCMB014�ɒǉ�����
'���Ұ�    :�ϐ���        ,IO ,�^            ,����
'          :record        ,I  ,typ_cmjc001j_Disp ,���o���R�[�h
'          :TSTAFFID      ,I  ,String       ,�o�^�Ј�ID
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/06/22(Fri)�쐬�@����

Public Function DBDRV_Getcmjc001j_Exec(record As typ_cmjc001j_Disp, TSTAFFID$) As FUNCTION_RETURN

Dim sql As String           'SQL�S��
Dim SetDate  As Variant     '���͓��t

    DBDRV_Getcmjc001j_Exec = FUNCTION_RETURN_FAILURE
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmjc001j_SQL.bas -- Function DBDRV_Getcmjc001j_Exec"

    SetDate = Format$(record.INPDATE, "yyyy-mm-dd hh:mm:ss")
  
  ''SQL��g�ݗ��Ă�
    sql = "Insert into TBCMB014 (GOUKI, INPDATE, FTIRFZI, FTIRCZH, FTIRCZC, MS1FZ, MS1CZ1, MS1CZ2, MS2FZ, MS2CZ1, MS2CZ2, " & _
          "MS3FZ, MS3CZ1, MS3CZ2, MS4FZ, MS4CZ1, MS4CZ2, MS5FZ, MS5CZ1, MS5CZ2, MSAVEFZ, MSAVECZ1, MSAVECZ2, " & _
          "MSSGFZ, MSSGCZ1, MSSGCZ2, MSPSGFZ, MSPSGCZ1, MSPSGCZ2, MSNSGFZ, MSNSGCZ1, MSNSGCZ2, MINFZ, MINCZ1, MINCZ2, " & _
          "MAXFZ, MAXCZ1, MAXCZ2, SGCK1FZ, SGCK1CZ1, SGCK1CZ2, SGCK2FZ, SGCK2CZ1, SGCK2CZ2, SGCK3FZ, SGCK3CZ1, SGCK3CZ2, " & _
          "SGCK4FZ, SGCK4CZ1, SGCK4CZ2, SGCK5FZ, SGCK5CZ1, SGCK5CZ2, SGCKDFZ, SGCKDCZ1, SGCKDCZ2, SGCKAFZ, SGCKAACZ1, SGCKACZ2, " & _
          "SGNFZ, SGNCZ1, SGNCZ2, FTIRFZ, FTIRCZ1, FTIRCZ2, EFFECTTM, YCOEF, XCOEF, RSQUARE, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE"
    sql = sql & ",SGCKST,SGCKFZ,SGCKCZ1,SGCKCZ2,FTIRCKST,FTIRCKFZ,FTIRCKCZ1,FTIRCKCZ2"  '2006/05/23�ǉ� kubota
    sql = sql & ",MS6FZ,MS6CZ1,MS6CZ2,SGCK6FZ,SGCK6CZ1,SGCK6CZ2,CVFZ,CVCZ1,CVCZ2"       '2010/03/26�ǉ� kubota
    sql = sql & ")"
    
    sql = sql & "Values('" & record.GOUKI & "', " & "TO_DATE('" & SetDate & "','YYYY-MM-DD hh24:mi:ss'), " & record.FTIRFZI & ", " & _
          record.FTIRCZH & ", " & record.FTIRCZC & ", " & record.MS1FZ & ", " & record.MS1CZ1 & ", " & record.MS1CZ2 & ", " & _
          record.MS2FZ & ", " & record.MS2CZ1 & ", " & record.MS2CZ2 & ", " & record.MS3FZ & ", " & record.MS3CZ1 & ", " & _
          record.MS3CZ2 & ", " & record.MS4FZ & ", " & record.MS4CZ1 & ", " & record.MS4CZ2 & ", " & record.MS5FZ & ", " & _
          record.MS5CZ1 & ", " & record.MS5CZ2 & ", " & record.MSAVEFZ & ", " & record.MSAVECZ1 & ", " & record.MSAVECZ2 & ", " & _
          record.MSSGFZ & ", " & record.MSSGCZ1 & ", " & record.MSSGCZ2 & ", " & record.MSPSGFZ & ", " & record.MSPSGCZ1 & ", " & _
          record.MSPSGCZ2 & ", " & record.MSNSGFZ & ", " & record.MSNSGCZ1 & ", " & record.MSNSGCZ2 & ", " & record.MINFZ & ", " & _
          record.MINCZ1 & ", " & record.MINCZ2 & ", " & record.MAXFZ & ", " & record.MAXCZ1 & ", " & record.MAXCZ2 & ", " & record.SGCK1FZ & ", " & _
          record.SGCK1CZ1 & ", " & record.SGCK1CZ2 & ", " & record.SGCK2FZ & ", " & record.SGCK2CZ1 & ", " & record.SGCK2CZ2 & ", " & _
          record.SGCK3FZ & ", " & record.SGCK3CZ1 & ", " & record.SGCK3CZ2 & ", " & record.SGCK4FZ & ", " & record.SGCK4CZ1 & ", " & _
          record.SGCK4CZ2 & ", " & record.SGCK5FZ & ", " & record.SGCK5CZ1 & ", " & record.SGCK5CZ2 & ", " & record.SGCKDFZ & ", " & _
          record.SGCKDCZ1 & ", " & record.SGCKDCZ2 & ", " & record.SGCKAFZ & ", " & record.SGCKAACZ1 & ", " & record.SGCKACZ2 & ", " & _
          record.SGNFZ & ", " & record.SGNCZ1 & ", " & record.SGNCZ2 & ", " & record.FTIRFZ & ", " & record.FTIRCZ1 & ", " & _
          record.FTIRCZ2 & ", " & record.EFFECTTM & ", " & record.YCOEF & ", " & record.XCOEF & ", " & record.RSQUARE & ", '" & _
          TSTAFFID & "', SYSDATE, ' ', SYSDATE, '0', SYSDATE"
    sql = sql & "," & record.SGCKST & ",'" & record.SGCKFZ & "','" & record.SGCKCZ1 & "','" & record.SGCKCZ2 & "'"          '2006/05/23�ǉ� kubota
    sql = sql & "," & record.FTIRCKST & ",'" & record.FTIRCKFZ & "','" & record.FTIRCKCZ1 & "','" & record.FTIRCKCZ2 & "'"  '2006/05/23�ǉ� kubota
    sql = sql & "," & record.MS6FZ & ",'" & record.MS6CZ1 & "','" & record.MS6CZ2 & "'"                                     '2010/03/26�ǉ� kubota
    sql = sql & "," & record.SGCK6FZ & ",'" & record.SGCK6CZ1 & "','" & record.SGCK6CZ2 & "'"                               '2010/03/26�ǉ� kubota
    sql = sql & "," & record.CVFZ & ",'" & record.CVCZ1 & "','" & record.CVCZ2 & "'"                                        '2010/03/26�ǉ� kubota
    sql = sql & ")"
  
  ''SQL�̎��s
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
'          :tblLeft       ,IO   ,typ_TBCMB014      ,�e�[�u���f�[�^�P
'          :tblRight      ,IO   ,typ_cmjc001j_Disp ,�e�[�u���f�[�^�Q
'          :bFlg          ,I   ,Boolean           ,TRUE:�����P�f�[�^�������Q�f�[�^�ւ̕ϊ�  FALSE:�����P�f�[�^�������Q�f�[�^�ւ̕ϊ�
'����      :
Public Sub ConvDate_F_cmjc001j_a(tblLeft As typ_TBCMB014, tblRight As typ_cmjc001j_Disp, bFlg As Boolean)
    If bFlg = True Then
        With tblRight
            .GOUKI = tblLeft.GOUKI
            .INPDATE = tblLeft.INPDATE
            .FTIRFZI = tblLeft.FTIRFZI
            .FTIRCZH = tblLeft.FTIRCZH
            .FTIRCZC = tblLeft.FTIRCZC
            .MS1FZ = tblLeft.MS1FZ
            .MS1CZ1 = tblLeft.MS1CZ1
            .MS1CZ2 = tblLeft.MS1CZ2
            .MS2FZ = tblLeft.MS2FZ
            .MS2CZ1 = tblLeft.MS2CZ1
            .MS2CZ2 = tblLeft.MS2CZ2
            .MS3FZ = tblLeft.MS3FZ
            .MS3CZ1 = tblLeft.MS3CZ1
            .MS3CZ2 = tblLeft.MS3CZ2
            .MS4FZ = tblLeft.MS4FZ
            .MS4CZ1 = tblLeft.MS4CZ1
            .MS4CZ2 = tblLeft.MS4CZ2
            .MS5FZ = tblLeft.MS5FZ
            .MS5CZ1 = tblLeft.MS5CZ1
            .MS5CZ2 = tblLeft.MS5CZ2
            .MSAVEFZ = tblLeft.MSAVEFZ
            .MSAVECZ1 = tblLeft.MSAVECZ1
            .MSAVECZ2 = tblLeft.MSAVECZ2
            .MSSGFZ = tblLeft.MSSGFZ
            .MSSGCZ1 = tblLeft.MSSGCZ1
            .MSSGCZ2 = tblLeft.MSSGCZ2
            .MSPSGFZ = tblLeft.MSPSGFZ
            .MSPSGCZ1 = tblLeft.MSPSGCZ1
            .MSPSGCZ2 = tblLeft.MSPSGCZ2
            .MSNSGFZ = tblLeft.MSNSGFZ
            .MSNSGCZ1 = tblLeft.MSNSGCZ1
            .MSNSGCZ2 = tblLeft.MSNSGCZ2
            .MINFZ = tblLeft.MINFZ
            .MINCZ1 = tblLeft.MINCZ1
            .MINCZ2 = tblLeft.MINCZ2
            .MAXFZ = tblLeft.MAXFZ
            .MAXCZ1 = tblLeft.MAXCZ1
            .MAXCZ2 = tblLeft.MAXCZ2
            .SGCK1FZ = tblLeft.SGCK1FZ
            .SGCK1CZ1 = tblLeft.SGCK1CZ1
            .SGCK1CZ2 = tblLeft.SGCK1CZ2
            .SGCK2FZ = tblLeft.SGCK2FZ
            .SGCK2CZ1 = tblLeft.SGCK2CZ1
            .SGCK2CZ2 = tblLeft.SGCK2CZ2
            .SGCK3FZ = tblLeft.SGCK3FZ
            .SGCK3CZ1 = tblLeft.SGCK3CZ1
            .SGCK3CZ2 = tblLeft.SGCK3CZ2
            .SGCK4FZ = tblLeft.SGCK4FZ
            .SGCK4CZ1 = tblLeft.SGCK4CZ1
            .SGCK4CZ2 = tblLeft.SGCK4CZ2
            .SGCK5FZ = tblLeft.SGCK5FZ
            .SGCK5CZ1 = tblLeft.SGCK5CZ1
            .SGCK5CZ2 = tblLeft.SGCK5CZ2
            .SGCKDFZ = tblLeft.SGCKDFZ
            .SGCKDCZ1 = tblLeft.SGCKDCZ1
            .SGCKDCZ2 = tblLeft.SGCKDCZ2
            .SGCKAFZ = tblLeft.SGCKAFZ
            .SGCKAACZ1 = tblLeft.SGCKAACZ1
            .SGCKACZ2 = tblLeft.SGCKACZ2
            .SGNFZ = tblLeft.SGNFZ
            .SGNCZ1 = tblLeft.SGNCZ1
            .SGNCZ2 = tblLeft.SGNCZ2
            .FTIRFZ = tblLeft.FTIRFZ
            .FTIRCZ1 = tblLeft.FTIRCZ1
            .FTIRCZ2 = tblLeft.FTIRCZ2
            .EFFECTTM = tblLeft.EFFECTTM
            .YCOEF = tblLeft.YCOEF
            .XCOEF = tblLeft.XCOEF
            .RSQUARE = tblLeft.RSQUARE
            
'2006/05/22�ǉ� kubota
            .SGCKST = tblLeft.SGCKST
            .SGCKFZ = tblLeft.SGCKFZ
            .SGCKCZ1 = tblLeft.SGCKCZ1
            .SGCKCZ2 = tblLeft.SGCKCZ2
            .FTIRCKST = tblLeft.FTIRCKST
            .FTIRCKFZ = tblLeft.FTIRCKFZ
            .FTIRCKCZ1 = tblLeft.FTIRCKCZ1
            .FTIRCKCZ2 = tblLeft.FTIRCKCZ2
        
'2010/03/26�ǉ� SETsw kubota
            .MS6FZ = tblLeft.MS6FZ
            .MS6CZ1 = tblLeft.MS6CZ1
            .MS6CZ2 = tblLeft.MS6CZ2
            .SGCK6FZ = tblLeft.SGCK6FZ
            .SGCK6CZ1 = tblLeft.SGCK6CZ1
            .SGCK6CZ2 = tblLeft.SGCK6CZ2
            .CVFZ = tblLeft.CVFZ
            .CVCZ1 = tblLeft.CVCZ1
            .CVCZ2 = tblLeft.CVCZ2
        
        End With
    Else
        With tblLeft
            .GOUKI = tblRight.GOUKI
            .INPDATE = tblRight.INPDATE
            .FTIRFZI = tblRight.FTIRFZI
            .FTIRCZH = tblRight.FTIRCZH
            .FTIRCZC = tblRight.FTIRCZC
            .MS1FZ = tblRight.MS1FZ
            .MS1CZ1 = tblRight.MS1CZ1
            .MS1CZ2 = tblRight.MS1CZ2
            .MS2FZ = tblRight.MS2FZ
            .MS2CZ1 = tblRight.MS2CZ1
            .MS2CZ2 = tblRight.MS2CZ2
            .MS3FZ = tblRight.MS3FZ
            .MS3CZ1 = tblRight.MS3CZ1
            .MS3CZ2 = tblRight.MS3CZ2
            .MS4FZ = tblRight.MS4FZ
            .MS4CZ1 = tblRight.MS4CZ1
            .MS4CZ2 = tblRight.MS4CZ2
            .MS5FZ = tblRight.MS5FZ
            .MS5CZ1 = tblRight.MS5CZ1
            .MS5CZ2 = tblRight.MS5CZ2
            .MSAVEFZ = tblRight.MSAVEFZ
            .MSAVECZ1 = tblRight.MSAVECZ1
            .MSAVECZ2 = tblRight.MSAVECZ2
            .MSSGFZ = tblRight.MSSGFZ
            .MSSGCZ1 = tblRight.MSSGCZ1
            .MSSGCZ2 = tblRight.MSSGCZ2
            .MSPSGFZ = tblRight.MSPSGFZ
            .MSPSGCZ1 = tblRight.MSPSGCZ1
            .MSPSGCZ2 = tblRight.MSPSGCZ2
            .MSNSGFZ = tblRight.MSNSGFZ
            .MSNSGCZ1 = tblRight.MSNSGCZ1
            .MSNSGCZ2 = tblRight.MSNSGCZ2
            .MINFZ = tblRight.MINFZ
            .MINCZ1 = tblRight.MINCZ1
            .MINCZ2 = tblRight.MINCZ2
            .MAXFZ = tblRight.MAXFZ
            .MAXCZ1 = tblRight.MAXCZ1
            .MAXCZ2 = tblRight.MAXCZ2
            .SGCK1FZ = tblRight.SGCK1FZ
            .SGCK1CZ1 = tblRight.SGCK1CZ1
            .SGCK1CZ2 = tblRight.SGCK1CZ2
            .SGCK2FZ = tblRight.SGCK2FZ
            .SGCK2CZ1 = tblRight.SGCK2CZ1
            .SGCK2CZ2 = tblRight.SGCK2CZ2
            .SGCK3FZ = tblRight.SGCK3FZ
            .SGCK3CZ1 = tblRight.SGCK3CZ1
            .SGCK3CZ2 = tblRight.SGCK3CZ2
            .SGCK4FZ = tblRight.SGCK4FZ
            .SGCK4CZ1 = tblRight.SGCK4CZ1
            .SGCK4CZ2 = tblRight.SGCK4CZ2
            .SGCK5FZ = tblRight.SGCK5FZ
            .SGCK5CZ1 = tblRight.SGCK5CZ1
            .SGCK5CZ2 = tblRight.SGCK5CZ2
            .SGCKDFZ = tblRight.SGCKDFZ
            .SGCKDCZ1 = tblRight.SGCKDCZ1
            .SGCKDCZ2 = tblRight.SGCKDCZ2
            .SGCKAFZ = tblRight.SGCKAFZ
            .SGCKAACZ1 = tblRight.SGCKAACZ1
            .SGCKACZ2 = tblRight.SGCKACZ2
            .SGNFZ = tblRight.SGNFZ
            .SGNCZ1 = tblRight.SGNCZ1
            .SGNCZ2 = tblRight.SGNCZ2
            .FTIRFZ = tblRight.FTIRFZ
            .FTIRCZ1 = tblRight.FTIRCZ1
            .FTIRCZ2 = tblRight.FTIRCZ2
            .EFFECTTM = tblRight.EFFECTTM
            .YCOEF = tblRight.YCOEF
            .XCOEF = tblRight.XCOEF
            .RSQUARE = tblRight.RSQUARE
        
'2006/05/22�ǉ� kubota
            .SGCKST = tblRight.SGCKST
            .SGCKFZ = tblRight.SGCKFZ
            .SGCKCZ1 = tblRight.SGCKCZ1
            .SGCKCZ2 = tblRight.SGCKCZ2
            .FTIRCKST = tblRight.FTIRCKST
            .FTIRCKFZ = tblRight.FTIRCKFZ
            .FTIRCKCZ1 = tblRight.FTIRCKCZ1
            .FTIRCKCZ2 = tblRight.FTIRCKCZ2

'2010/03/26�ǉ� SETsw kubota
            .MS6FZ = tblRight.MS6FZ
            .MS6CZ1 = tblRight.MS6CZ1
            .MS6CZ2 = tblRight.MS6CZ2
            .SGCK6FZ = tblRight.SGCK6FZ
            .SGCK6CZ1 = tblRight.SGCK6CZ1
            .SGCK6CZ2 = tblRight.SGCK6CZ2
            .CVFZ = tblRight.CVFZ
            .CVCZ1 = tblRight.CVCZ1
            .CVCZ2 = tblRight.CVCZ2
        
        End With
    End If

End Sub

'------------------------------------------------
' DB�A�N�Z�X�֐�
'------------------------------------------------

'�T�v      :�e�[�u���uTBCMB014�v��������ɂ��������R�[�h�𒊏o����
'���Ұ�    :�ϐ���        ,IO ,�^           ,����
'          :records()     ,O  ,typ_TBCMB014 ,���o���R�[�h
'          :sqlWhere      ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'          :sqlOrder      ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'          :�߂�l        ,O  ,FUNCTION_RETURN ,���o�̐���
'����      :
'����      :2001/08/24�쐬�@�쑺
Public Function DBDRV_GetTBCMB014(records() As typ_TBCMB014, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL�S��
Dim sqlBase As String   'SQL��{��(WHERE�߂̑O�܂�)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      '���R�[�h��
Dim i As Long

    ''SQL��g�ݗ��Ă�
    sqlBase = "Select GOUKI, INPDATE, FTIRFZI, FTIRCZH, FTIRCZC, MS1FZ, MS1CZ1, MS1CZ2, MS2FZ, MS2CZ1, MS2CZ2, MS3FZ, MS3CZ1," & _
              " MS3CZ2, MS4FZ, MS4CZ1, MS4CZ2, MS5FZ, MS5CZ1, MS5CZ2, MSAVEFZ, MSAVECZ1, MSAVECZ2, MSSGFZ, MSSGCZ1, MSSGCZ2," & _
              " MSPSGFZ, MSPSGCZ1, MSPSGCZ2, MSNSGFZ, MSNSGCZ1, MSNSGCZ2, MINFZ, MINCZ1, MINCZ2, MAXFZ, MAXCZ1, MAXCZ2, SGCK1FZ," & _
              " SGCK1CZ1, SGCK1CZ2, SGCK2FZ, SGCK2CZ1, SGCK2CZ2, SGCK3FZ, SGCK3CZ1, SGCK3CZ2, SGCK4FZ, SGCK4CZ1, SGCK4CZ2," & _
              " SGCK5FZ, SGCK5CZ1, SGCK5CZ2, SGCKDFZ, SGCKDCZ1, SGCKDCZ2, SGCKAFZ, SGCKAACZ1, SGCKACZ2, SGNFZ, SGNCZ1, SGNCZ2," & _
              " FTIRFZ, FTIRCZ1, FTIRCZ2, EFFECTTM, YCOEF, XCOEF, RSQUARE, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG," & _
              " SENDDATE "
    sqlBase = sqlBase & "From TBCMB014"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB014 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''���o���ʂ��i�[����
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .GOUKI = rs("GOUKI")             ' ���@
            .INPDATE = rs("INPDATE")         ' ���t
            .FTIRFZI = rs("FTIRFZI")         ' FTIR�iFZ)
            .FTIRCZH = rs("FTIRCZH")         ' FTIR�iCZ���j
            .FTIRCZC = rs("FTIRCZC")         ' FTIR�iCZ���j
            .MS1FZ = rs("MS1FZ")             ' ����T���v��1�iFZ)
            .MS1CZ1 = rs("MS1CZ1")           ' ����T���v��1�iCZ-1)
            .MS1CZ2 = rs("MS1CZ2")           ' ����T���v��1�iCZ-2)
            .MS2FZ = rs("MS2FZ")             ' ����T���v��2�iFZ)
            .MS2CZ1 = rs("MS2CZ1")           ' ����T���v��2�iCZ-1)
            .MS2CZ2 = rs("MS2CZ2")           ' ����T���v��2�iCZ-2)
            .MS3FZ = rs("MS3FZ")             ' ����T���v��3�iFZ)
            .MS3CZ1 = rs("MS3CZ1")           ' ����T���v��3�iCZ-1)
            .MS3CZ2 = rs("MS3CZ2")           ' ����T���v��3�iCZ-2)
            .MS4FZ = rs("MS4FZ")             ' ����T���v��4�iFZ)
            .MS4CZ1 = rs("MS4CZ1")           ' ����T���v��4�iCZ-1)
            .MS4CZ2 = rs("MS4CZ2")           ' ����T���v��4�iCZ-2)
            .MS5FZ = rs("MS5FZ")             ' ����T���v��5�iFZ)
            .MS5CZ1 = rs("MS5CZ1")           ' ����T���v��5�iCZ-1)
            .MS5CZ2 = rs("MS5CZ2")           ' ����T���v��5�iCZ-2)
            .MSAVEFZ = rs("MSAVEFZ")         ' ���蕽�ρiFZ�j
            .MSAVECZ1 = rs("MSAVECZ1")       ' ���蕽�ρiCZ-1�j
            .MSAVECZ2 = rs("MSAVECZ2")       ' ���蕽�ρiCZ-2�j
            .MSSGFZ = rs("MSSGFZ")           ' ����ЁiFZ�j
            .MSSGCZ1 = rs("MSSGCZ1")         ' ����ЁiCZ-1�j
            .MSSGCZ2 = rs("MSSGCZ2")         ' ����ЁiCZ-2�j
            .MSPSGFZ = rs("MSPSGFZ")         ' ����AVE+�ЁiFZ�j
            .MSPSGCZ1 = rs("MSPSGCZ1")       ' ����AVE+�ЁiCZ-1�j
            .MSPSGCZ2 = rs("MSPSGCZ2")       ' ����AVE+�ЁiCZ-2�j
            .MSNSGFZ = rs("MSNSGFZ")         ' ����AVE-�ЁiFZ�j
            .MSNSGCZ1 = rs("MSNSGCZ1")       ' ����AVE-�ЁiCZ-1�j
            .MSNSGCZ2 = rs("MSNSGCZ2")       ' ����AVE-�ЁiCZ-2�j
            .MINFZ = rs("MINFZ")             ' MIN�iFZ�j
            .MINCZ1 = rs("MINCZ1")           ' MIN�iCZ-1�j
            .MINCZ2 = rs("MINCZ2")           ' MIN�iCZ-2�j
            .MAXFZ = rs("MAXFZ")             ' MAX�iFZ�j
            .MAXCZ1 = rs("MAXCZ1")           ' MAX�iCZ-1�j
            .MAXCZ2 = rs("MAXCZ2")           ' MAX�iCZ-2�j
            .SGCK1FZ = rs("SGCK1FZ")         ' ��ck�T���v��1�iFZ)
            .SGCK1CZ1 = rs("SGCK1CZ1")       ' ��ck�T���v��1�iCZ-1)
            .SGCK1CZ2 = rs("SGCK1CZ2")       ' ��ck�T���v��1�iCZ-2)
            .SGCK2FZ = rs("SGCK2FZ")         ' ��ck�T���v��2�iFZ)
            .SGCK2CZ1 = rs("SGCK2CZ1")       ' ��ck�T���v��2�iCZ-1)
            .SGCK2CZ2 = rs("SGCK2CZ2")       ' ��ck�T���v��2�iCZ-2)
            .SGCK3FZ = rs("SGCK3FZ")         ' ��ck�T���v��3�iFZ)
            .SGCK3CZ1 = rs("SGCK3CZ1")       ' ��ck�T���v��3�iCZ-1)
            .SGCK3CZ2 = rs("SGCK3CZ2")       ' ��ck�T���v��3�iCZ-2)
            .SGCK4FZ = rs("SGCK4FZ")         ' ��ck�T���v��4�iFZ)
            .SGCK4CZ1 = rs("SGCK4CZ1")       ' ��ck�T���v��4�iCZ-1)
            .SGCK4CZ2 = rs("SGCK4CZ2")       ' ��ck�T���v��4�iCZ-2)
            .SGCK5FZ = rs("SGCK5FZ")         ' ��ck�T���v��5�iFZ)
            .SGCK5CZ1 = rs("SGCK5CZ1")       ' ��ck�T���v��5�iCZ-1)
            .SGCK5CZ2 = rs("SGCK5CZ2")       ' ��ck�T���v��5�iCZ-2)
            .SGCKDFZ = rs("SGCKDFZ")         ' ��ck�f�[�^���iFZ�j
            .SGCKDCZ1 = rs("SGCKDCZ1")       ' ��ck�f�[�^���iCZ-1�j
            .SGCKDCZ2 = rs("SGCKDCZ2")       ' ��ck�f�[�^���iCZ-2�j
            .SGCKAFZ = rs("SGCKAFZ")         ' ��ck���ρiFZ�j
            .SGCKAACZ1 = rs("SGCKAACZ1")     ' ��ck���ρiCZ-1�j
            .SGCKACZ2 = rs("SGCKACZ2")       ' ��ck���ρiCZ-2�j
            .SGNFZ = rs("SGNFZ")             ' ��ck�ЁiFZ�j
            .SGNCZ1 = rs("SGNCZ1")           ' ��ck�� CZ-1�j
            .SGNCZ2 = rs("SGNCZ2")           ' ��ck�ЁiCZ-2�j
            .FTIRFZ = rs("FTIRFZ")           ' FTIR���Z�iFZ�j
            .FTIRCZ1 = rs("FTIRCZ1")         ' FTIR���Z�iCZ-1�j
            .FTIRCZ2 = rs("FTIRCZ2")         ' FTIR���Z�iCZ-2�j
            .EFFECTTM = rs("EFFECTTM")       ' �L������
            .YCOEF = rs("YCOEF")             ' �e�s�h�q���Z���i�x�ؕЁj
            .XCOEF = rs("XCOEF")             ' �e�s�h�q���Z���i�w�W���j
            .RSQUARE = rs("RSQUARE")         ' �q�Q��
            .TSTAFFID = rs("TSTAFFID")       ' �o�^�Ј�ID
            .REGDATE = rs("REGDATE")         ' �o�^���t
            .KSTAFFID = rs("KSTAFFID")       ' �X�V�Ј�ID
            .UPDDATE = rs("UPDDATE")         ' �X�V���t
            .SENDFLAG = rs("SENDFLAG")       ' ���M�t���O
            .SENDDATE = rs("SENDDATE")       ' ���M���t
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMB014 = FUNCTION_RETURN_SUCCESS
End Function

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
'2006/05/22�ǉ�
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
    gErr.Push "s_cmzc029_SQL.bas -- Function GetSigChkCode"

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
    sSql = sSql & "   AND codea9 = 'GFA'"
    
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


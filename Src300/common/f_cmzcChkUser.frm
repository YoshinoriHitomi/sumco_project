VERSION 5.00
Begin VB.Form f_cmzcChkUser 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "۸޵�"
   ClientHeight    =   1245
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735.587
   ScaleMode       =   0  'հ�ް
   ScaleWidth      =   3549.215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "��ݾ�"
      Height          =   390
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  '�̌Œ�
      Left            =   1290
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "�߽ܰ��(&P):"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "f_cmzcChkUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private InpPass As String

Private Sub cmdCancel_Click()
    InpPass = ""
    txtPassword.Text = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    InpPass = txtPassword.Text
    txtPassword.Text = ""
    Me.Hide
End Sub

Private Sub Form_Load()
    InpPass = ""
End Sub

Public Function GetInpPass() As String
    GetInpPass = InpPass
End Function

'�T�v      :�����Ǘ��}�X�^�[�ƎЈ��}�X�^�[��茠���`�F�b�N�������Ȃ��B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :StaffID       ,I  ,String           ,�Ј�ID
'          :ProcID        ,I  ,String           ,�H��ID
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,����E�ُ�
'����      :������Ȃ��ꍇ��VbNullString��Ԃ�
'����      :2001/08/16 �l��
'          :2009/08 Sumco �H��

Private Function ChkUser(ByVal StaffID As String, ByVal ProcID As String) As FUNCTION_RETURN
    
    Dim dbIsMine As Boolean
    Dim rs As OraDynaset
    Dim sql As String
    Dim Pass As String
    Dim Pwchk As String
    Dim InpPwd As String

    On Error GoTo proc_err
    gErr.Push "s_cmzcChkUser.bas -- Function ChkUser"

    ChkUser = FUNCTION_RETURN_SUCCESS
    
    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    '' �����Ŏw�肳�ꂽ�Ј�ID�ƍH���R�[�h�Ō������s��
    sql = " select T1.PASSWD as PASS, T4.PWCHECK as PWCHK from TBCMB001 T1, TBCMB004 T4 "
    sql = sql & " Where T1.EXECODE = T4.AUTHCODE "
    sql = sql & " and T1.STAFFID = '" & StaffID & "' "
    sql = sql & " and T4.TRANID = '" & ProcID & "' "
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    
    ''�݂���Ȃ�������G���[
    If rs.RecordCount = 0 Then
        ChkUser = FUNCTION_RETURN_FAILURE
    
    Else ''����������A�p�X���[�h�`�F�b�N�������Ȃ�
        Pass = rs("PASS")
        Pwchk = rs("PWCHK")
        
        ' �p�X���[�h�`�F�b�N����
        If Trim(Pwchk) = "1" Then
            'Load f_cmzcChkUser
            f_cmzcChkUser.Show 1
            InpPwd = f_cmzcChkUser.GetInpPass()
            Unload f_cmzcChkUser
            If InpPwd <> Pass Then
                ChkUser = FUNCTION_RETURN_FAILURE
            End If
        End If
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If
proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    gErr.HandleError
    ChkUser = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'�T�v      :�����`�F�b�N�֐����ĂԑO�̃t�H�[�����[���H���R�[�h�ϊ��֐�
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :StaffID        ,I  ,String          ,�Ј�ID
'          :ProcID          ,I  ,String         ,�H��ID
'          :�߂�l        ,O  ,String           ,����E�ُ�
'����      :������Ȃ��ꍇ��""�󕶎���Ԃ�
'����      :2001/08/16 �l��
'          :2002/07/29 �}�@�i�V�t�H�[�����Ή��j
'          :2009/08 SUMCO �H���@(�w��������ѓ���:f_cmbc053_1)�ǉ�

Private Function CnvFrm2Prc(ByVal FrmN As String) As String

    On Error GoTo proc_err
    gErr.Push "s_cmzcChkUser.bas -- Function CnvFrm2Prc"

    Select Case (FrmN)
        Case "f_cmac001_1":  CnvFrm2Prc = vbNullString                   '���C�����j���[���
        Case "f_cmac001_2":  CnvFrm2Prc = vbNullString                   '�d�l������j���[
        Case "f_cmac001_3":  CnvFrm2Prc = vbNullString                   '�������j���[
        Case "f_cmac001_4":  CnvFrm2Prc = vbNullString                   '�������チ�j���[
        Case "f_cmac001_5":  CnvFrm2Prc = vbNullString                   '�������H���j���[
        Case "f_cmac001_6":  CnvFrm2Prc = vbNullString                   '�����������j���[
        Case "f_cmac001_7":  CnvFrm2Prc = vbNullString                   '�������o���j���[
        Case "f_cmac001_8":  CnvFrm2Prc = vbNullString                   'WF���H���j���[
        Case "f_cmac001_9":  CnvFrm2Prc = vbNullString                   '���̑����j���[
        Case "f_cmac001_10":  CnvFrm2Prc = vbNullString                   '�o�[�R�[�h���x���Ĕ��s���j���[
        
        
        Case "f_cmbc001_1":  CnvFrm2Prc = PROCD_SIYOU_UKEIRE             '�d�l����d�|�ꗗ
        Case "f_cmbc001_2":  CnvFrm2Prc = PROCD_SIYOU_UKEIRE             '���i�d�l���
        Case "f_cmbc003_1":  CnvFrm2Prc = PROCD_SIYOU_NYUURYOKU          '���i�d�l����
        
        
        Case "f_cmbc004_1":  CnvFrm2Prc = PROCD_TAKESSYOU_UKEIRE         '����������I��
        Case "f_cmbc005_1":  CnvFrm2Prc = PROCD_RIMERUTO_UKEIRE          '�������g����E�ؒf
        Case "f_cmbc006_1":  CnvFrm2Prc = PROCD_RIMERUTO_HARAIDASI       '�������g���E���o
        Case "f_cmbc007_1":  CnvFrm2Prc = PROCD_GENRYOU_ZAIKO_SYUUSEI    '�����݌ɏC��
        Case "f_cmbc008_1":  CnvFrm2Prc = PROCD_KAKUAGE                  '�N���X�^���J�^���O�����i��
        
        
        Case "f_cmbc008_3":  CnvFrm2Prc = vbNullString                   '���@�ꗗ
        Case "f_cmbc009_2":  CnvFrm2Prc = vbNullString                   '�w���ꗗ
        Case "f_cmbc009_3":  CnvFrm2Prc = PROCD_HIKIAGE_SIJI             '�w�����e����
        Case "f_cmbc009_4":  CnvFrm2Prc = vbNullString                   '�����ꗗ
        Case "f_cmbc009_5":  CnvFrm2Prc = vbNullString                   'PG-ID�ꗗ
        Case "f_cmbc009_6":  CnvFrm2Prc = vbNullString                   'PG - ID�ڍ�
        Case "f_cmbc009_7":  CnvFrm2Prc = vbNullString                   '�����u���b�N�g����
        Case "f_cmbc013_1":  CnvFrm2Prc = PROCD_HIKIAGE_TOUNYUU          '���㓊�����я������
        Case "f_cmbc013_2":  CnvFrm2Prc = PROCD_HIKIAGE_TOUNYUU          '���㓊������
        Case "f_cmbc013_3":  CnvFrm2Prc = vbNullString                   '�����������шꗗ
        Case "f_cmbc014_1":  CnvFrm2Prc = PROCD_HIKIAGE_SYUURYOU         '����I�����ѓ���
        Case "f_cmbc015_1":  CnvFrm2Prc = PROCD_TEIKO_HENSEKI_KEISSAN    '��R�ΐ͌v�Z
        
        
        Case "f_cmbc016_1":  CnvFrm2Prc = PROCD_KAKOU_HARAIDASI          '�������H���o
        Case "f_cmbc017_1":  CnvFrm2Prc = PROCD_KENNSAKU_KAKOU           '����������H���ѓ���
        Case "f_cmbc018_1":  CnvFrm2Prc = vbNullString                   '�ؒf�҂��ꗗ
        Case "f_cmbc018_2":  CnvFrm2Prc = PROCD_SETUDAN                  '�ؒf
        Case "f_cmbc019_1":  CnvFrm2Prc = PROCD_KESSYOU_HOLD             '�����z�[���h
        Case "f_cmbc020_1":  CnvFrm2Prc = PROCD_KESSYOU_HOLD_KAIJO       '�����z�[���h����
        
        
        Case "f_cmbc021_1":  CnvFrm2Prc = PROCD_FTIR                     'FTIR(Oi,Cs)���ѓ���
        Case "f_cmbc022_1":  CnvFrm2Prc = PROCD_GFA                      'GFA(Oi)���ѓ���
        Case "f_cmbc023_1":  CnvFrm2Prc = PROCD_TEIKOU                   '��R���ѓ���
        Case "f_cmbc024_1":  CnvFrm2Prc = PROCD_BMD                      'BMD���ѓ���
        Case "f_cmbc025_1":  CnvFrm2Prc = PROCD_OSF                      'OSF���ѓ���
        Case "f_cmbc026_1":  CnvFrm2Prc = PROCD_GD                       'GD���ѓ���
        Case "f_cmbc027_1":  CnvFrm2Prc = PROCD_LIFETIME                 '���C�t�^�C�����ѓ���
        Case "f_cmbc028_1":  CnvFrm2Prc = PROCD_EPD                      'EPD���ѓ���
        Case "f_cmbc029_1":  CnvFrm2Prc = PROCD_GFA_KOUSEIJOHO           'GFA�Z�����ݒ�
        '2009/08 SUMCO Akizuki
        Case "f_cmbc053_1":  CnvFrm2Prc = PROCD_X                        'X�����ѓ���

                
'' 09/01/28 FAE)akiyama start
        Case "f_cmbc030_1":  CnvFrm2Prc = PROCD_KESSYOU_SOUGOUHANTEI     '�҂��ꗗ(��������j
'' 09/01/28 FAE)akiyama start
        Case "f_cmbc030_2":  CnvFrm2Prc = PROCD_KESSYOU_SOUGOUHANTEI     '��������
        Case "f_cmbc031_1":  CnvFrm2Prc = PROCD_KESSYOU_SOUGOUHANTEI     '�������� - �ĕ���
        
        Case "f_cmbc032_1":  CnvFrm2Prc = vbNullString                   '�҂��ꗗ�i�����ŏI���o�j
        Case "f_cmbc032_2":  CnvFrm2Prc = PROCD_KESSYOU_SAISYUU_HARAIDASI ' �����ŏI���o����
        Case "f_cmbc033_1":  CnvFrm2Prc = vbNullString                   '�҂��ꗗ�i�����w���j
        Case "f_cmbc033_2":  CnvFrm2Prc = PROCD_NUKISI_SIJI              '�����w������
        Case "f_cmbc034_1":  CnvFrm2Prc = PROCD_WFC_HARAIDASI            'WF�Z���^�[���o
        Case "f_cmbc035_1":  CnvFrm2Prc = PROCD_KESSYOU_SIYOUJOUHOU_HENKOU ' �������ύX
        Case "f_cmbc036_1":  CnvFrm2Prc = vbNullString                   '�����u���b�N�ꗗ
        Case "f_cmbc036_2":  CnvFrm2Prc = PROCD_NUKISI_HENKOU            '�����w���ύX����
        Case "f_cmbc036_3":  CnvFrm2Prc = PROCD_NUKISI_HENKOU            '�֘A��ۯ��Ǘ��@08/01/28 ooba
        Case "f_cmbc037_1":  CnvFrm2Prc = PROCD_KOUNYU_TAN_KESSYOU       '�w���P�������
        Case "f_cmbc038_1":  CnvFrm2Prc = PROCD_BLOCK                    '�u���b�N���x�����s 4/16  Yam
        
        
        Case "f_cmbc039_1":  CnvFrm2Prc = vbNullString                   'WF�������� - ����҂��ꗗ
        Case "f_cmbc039_2":  CnvFrm2Prc = PROCD_WFC_SOUGOUHANTEI         'WF�������� - ��������
        Case "f_cmbc039_3":  CnvFrm2Prc = PROCD_WFC_SOUGOUHANTEI         'WF�������� - �Ĕ����w��
        Case "f_cmbc040_1":  CnvFrm2Prc = PROCD_SXL_KAKUTEI              '�V���O���m��

''Add Start 2011/04/06 SMPK Nakamura
        Case "f_cmbc055_1":  CnvFrm2Prc = PROCD_FRS_KOUSEIJOHO           'FRS�Z�����ݒ�
''Add End 2011/04/06 SMPK Nakamura
        
        Case "f_cmcc001_1":  CnvFrm2Prc = PROCD_PGID_MNT                 'PG-ID�ꗗ
        Case "f_cmcc001_2":  CnvFrm2Prc = PROCD_PGID_MNT                 'PG-ID�ڍ�
        Case "f_cmcc002_1":  CnvFrm2Prc = PROCD_SEISAKUJOUKEN_MNT        '������������e�i���X
        
        Case "f_cmcc003_1":  CnvFrm2Prc = PROCD_SYAIN_MNT                '�Ј��}�X�^�����e�i���X'02/3/30 Yam
        Case "f_cmcc004_1":  CnvFrm2Prc = PROCD_KENGEN_MNT               '�����}�X�^�����e�i���X'02/3/30 Yam
        Case "f_cmcc005_1":  CnvFrm2Prc = PROCD_CODE_MNT                 '�R�[�h�}�X�^�����e�i���X'02/3/30 Yam
        Case "f_cmcc006_1":  CnvFrm2Prc = PROCD_PRINTER_MNT              '���x���v�����^�}�X�^�����e�i���X'02/3/30 Yam
        ' ���̗p��ǉ�  2003/09/12 SystemBrain ===================> START
        Case "TOKUSAI":      CnvFrm2Prc = PROCD_TOKUSAI_KENGEN           '���̌���
        ' ���̗p��ǉ�  2003/09/12 SystemBrain ===================> END
        Case Else:          CnvFrm2Prc = vbNullString
    End Select
    
proc_exit:
    '' �I��
    gErr.Pop
    Exit Function

proc_err:
    '' �G���[�n���h��
    gErr.HandleError
    CnvFrm2Prc = ""
    Resume proc_exit
End Function


'�T�v      :�Ј�ID�Ɖ�ʖ�����A���s�����𓾂�
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :frmName       ,   ,String    ,
'          :staffID       ,   ,String    ,
'          :�߂�l        ,O  ,FUNCTION_RETURN,
'����      :
'����      :2001/8/17 �쐬  �쑺
Public Function CanExec(ByVal frmName$, ByVal StaffID$) As Boolean
    If ChkUser(StaffID, CnvFrm2Prc(frmName)) = FUNCTION_RETURN_SUCCESS Then
        CanExec = True
        XSDC3_StaffID = StaffID
    Else
        CanExec = False
    End If
End Function

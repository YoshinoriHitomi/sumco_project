VERSION 5.00
Begin VB.Form f_cmzc003a 
   Caption         =   "�����} (123456789012)"
   ClientHeight    =   5988
   ClientLeft      =   5112
   ClientTop       =   3060
   ClientWidth     =   5340
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   5988
   ScaleWidth      =   5340
   Begin cmbc039.o_cmzc002a PicXL 
      Height          =   5535
      Left            =   0
      Top             =   360
      Width           =   5295
      _ExtentX        =   9335
      _ExtentY        =   9758
   End
   Begin VB.Label lblDspTime 
      Alignment       =   1  '�E����
      Caption         =   "mm/dd hh:mm"
      Height          =   195
      Left            =   3360
      TabIndex        =   0
      Top             =   60
      Width           =   1935
   End
End
Attribute VB_Name = "f_cmzc003a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'======================================================
' �����}�E�B���h�E
' �T�v    : �^����ꂽ�����N���X�̓��e��}������
' �Q��    : �����R���g���[��(o_cmzc002a.ctl)
'         : �����N���X      (c_cmzcXl.ctl)
'======================================================

'���W�X�g���ۑ��Ɋւ���萔��`
Private Const SYSTEM_NAME = "300mm�������ƃV�X�e��"
Private Const CATEGORY = "�E�B���h�E�ʒu"
'���W�X�g���ۑ��ʒu�́@HKEY_CURRENT_USER\Software\VB and VBA Program Settings\<SYSTEM_NAME>\<CATEGORY>

'*******************************************************************************
'*    �֐���        : Form_KeyUp
'*
'*    �����T�v      : 1.ESC�L�[�������ꂽ��A�E�B���h�E�����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    KeyCode     ,I  ,Integer�@,�L�[�R�[�h
'*                    Shift       ,I  ,Integer�@,Shift�L�[�̏��
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "f_cmzc003a.frm -- Sub Form_KeyUp"

    '' ESC�L�[�������ꂽ��A�E�B���h�E�����
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub

'*******************************************************************************
'*    �֐���        : Form_Load
'*
'*    �����T�v      : 1.�����}�R���g���[���̕\���ʒu��O��ɍ��킹��
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub Form_Load()
    Dim lngOldTop   As Long
    Dim lngOldLeft  As Long

    '' �����}�R���g���[���̕\���ʒu��O��ɍ��킹��

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "f_cmzc003a.frm -- Sub Form_Load"

    lngOldTop = CLng(GetSetting(SYSTEM_NAME, CATEGORY, Me.Name & "-Top", "-1"))    ''Top�ʒu�̕���
    lngOldLeft = CLng(GetSetting(SYSTEM_NAME, CATEGORY, Me.Name & "-Left", "-1"))   ''Left�ʒu�̕���
    If (lngOldTop > 0) And (lngOldLeft > 0) Then
        Me.Move lngOldLeft, lngOldTop
    End If

proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub

'*******************************************************************************
'*    �֐���        : Form_Resize
'*
'*    �����T�v      : 1.�R���g���[���̈ʒu�E�傫����A������
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub Form_Resize()
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "f_cmzc003a.frm -- Sub Form_Resize"

    On Error Resume Next
    
    '' �\���������̑傫����A������
    lblDspTime.left = Width - lblDspTime.Width - 100
    
    '' �����}�R���g���[���̑傫����A������
    PicXL.Width = Width - PicXL.left - 100      ''����A������
    PicXL.Height = Height - PicXL.top - 400     ''������A������

proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub

'*******************************************************************************
'*    �֐���        : Form_Unload
'*
'*    �����T�v      : 1.�����}�E�B���h�E�̕\���ʒu���L������
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    Cancel        ,I  ,Integer
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "f_cmzc003a.frm -- Sub Form_Unload"

    ''��ʂ����Ƃ��ɁA�\���ʒu���L�����Ă���
    SaveSetting SYSTEM_NAME, CATEGORY, Me.Name & "-Top", top     ''Top�ʒu��ۑ�
    SaveSetting SYSTEM_NAME, CATEGORY, Me.Name & "-Left", left   ''Left�ʒu��ۑ�

proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub

'*******************************************************************************
'*    �֐���        : Clear
'*
'*    �����T�v      : 1.�����}�E�B���h�E������������
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Public Sub Clear()
    '' �t�H�[���̃L���v�V����������������

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "f_cmzc003a.frm -- Sub Clear"

    Caption = "�����} (____________)"
    
    '' �\������������������
    lblDspTime.Caption = Format$(Now, "m/d  h:m")
    
    '' �����}������������
    PicXL.Clear

proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub

'*******************************************************************************
'*    �֐���        : Draw
'*
'*    �����T�v      : 1.�����}�E�B���h�E��`�悷��
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    xl�@�@�@�@    ,I  ,c_cmzcXl ,�������
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Public Sub Draw(Xl As c_cmzcXl)

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "f_cmzc003a.frm -- Sub Draw"

    '' �t�H�[���̃L���v�V������ݒ肷��
    Caption = "�����} (" & Xl.CRYNUM & ")"
    
    '' �\��������ݒ肷��
    lblDspTime.Caption = Format$(Now, "m/d  h:m")
    
    '' �����}��`�悷��
    PicXL.Draw Xl

proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub

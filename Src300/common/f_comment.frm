VERSION 5.00
Begin VB.Form f_comment 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�R�����g����"
   ClientHeight    =   1785
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6825
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1054.637
   ScaleMode       =   0  'հ�ް
   ScaleWidth      =   6408.306
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2663
      TabIndex        =   2
      Top             =   1185
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "�L�����Z��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6923
      TabIndex        =   3
      Top             =   1185
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtComment 
      Height          =   345
      IMEMode         =   1  '��
      Left            =   1065
      MaxLength       =   60
      TabIndex        =   1
      Text            =   "�R�����g���̓��͉\�ő吔�͑S�p�łR�O�����܂łƂȂ��Ă��܂�"
      Top             =   508
      Width           =   5538
   End
   Begin VB.Label Label1 
      Caption         =   "�R�����g����͂��ĉ������B�����s���ɉ�ʂ𒆉��ɖ߂��܂��B"
      Height          =   255
      Left            =   213
      TabIndex        =   4
      Top             =   120
      Width           =   5112
   End
   Begin VB.Label lblLabels 
      Caption         =   "�R�����g"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   213
      TabIndex        =   0
      Top             =   641
      Width           =   851
   End
End
Attribute VB_Name = "f_comment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Ans As VbMsgBoxResult

Private Sub btnCancel_Click()
    Ans = vbCancel
    Me.Hide
End Sub

Private Sub btnOK_Click()
    If ChkString(txtComment.text, 60, 0) = CHK_NG Then
        MsgBox "���͉\���𒴂��Ă��܂��B�S�p�R�O�����ȓ��œ��͂��ĉ������B"
        Exit Sub
    End If
    Ans = vbOK
    Me.Hide
End Sub


'�T�v      :���F�@�\Web�V�X�e���p�R�����g����͂���
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :COMMENT$      ,O  ,String    ,�R�����g
'          :�߂�l        ,O  ,VbMsgBoxResult,�����ꂽ�{�^��(vbOk/vbCancel)
'����      :
'����      :2007/09/21 �쐬  �{��
Public Function GetComment(COMMENT$) As VbMsgBoxResult

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "f_comment.frm -- Function GetComment"
    
    txtComment.text = vbNullString
'    btnOK.Enabled = False
    
    Me.Show 1
    If Ans = vbOK Then
        COMMENT = txtComment.text
    Else
        COMMENT = vbNullString
    End If
    GetComment = Ans
    
    Unload Me

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function

'Private Sub txtComment_Change()
'    If Trim(txtComment.text) <> "" And LenB(Trim(txtComment.text)) > 0 Then
'        btnOK.Enabled = True
'    Else
'        btnOK.Enabled = False
'    End If
'End Sub

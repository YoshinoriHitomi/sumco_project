VERSION 5.00
Begin VB.Form WFCJudgDialog 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "WFC Judg Message"
   ClientHeight    =   1380
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6012
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   6012
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1128
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   4305
   End
End
Attribute VB_Name = "WFCJudgDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'*******************************************************************************
'*    �֐���        : WFCErrorMessage
'*
'*    �����T�v      : 1.ErrMsg��\������
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*�@�@�@�@�@�@�@�@�@�@ErrMsg      ,I  ,String   ,ErrMsg
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Public Sub WFCErrorMessage(ErrMsg As String)
    List1.AddItem ErrMsg, List1.ListCount
End Sub

'*******************************************************************************
'*    �֐���        : OKButton_Click
'*
'*    �����T�v      : 1.��ʂ��B��
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*�@�@�@�@�@�@�@�@�@�@�Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub OKButton_Click()
    Me.Visible = False
End Sub

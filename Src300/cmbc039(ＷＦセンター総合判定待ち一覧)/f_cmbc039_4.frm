VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form f_cmbc039_4 
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   " f_cmbc039_4"
   ClientHeight    =   10875
   ClientLeft      =   1575
   ClientTop       =   1680
   ClientWidth     =   15270
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   725
   ScaleMode       =   3  '�߸��
   ScaleWidth      =   1018
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.ComboBox cmbSprChg 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      ItemData        =   "f_cmbc039_4.frx":0000
      Left            =   1710
      List            =   "f_cmbc039_4.frx":0010
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   10
      Top             =   1035
      Width           =   1425
   End
   Begin VB.TextBox txtSxlId 
      BackColor       =   &H0080FF80&
      Height          =   264
      Left            =   165
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1050
      Width           =   1335
   End
   Begin FPSpread.vaSpread sprWfmapView 
      Height          =   7575
      Left            =   165
      TabIndex        =   11
      Top             =   1560
      Width           =   14910
      _Version        =   196608
      _ExtentX        =   26300
      _ExtentY        =   13361
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      BackColorStyle  =   1
      ColsFrozen      =   6
      MaxCols         =   32
      MaxRows         =   1
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "f_cmbc039_4.frx":0030
      UserResize      =   0
      VisibleCols     =   12
   End
   Begin VB.Frame fraHead 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   15225
      Begin VB.Label lblTime 
         Height          =   150
         Left            =   13740
         TabIndex        =   3
         Top             =   240
         Width           =   1450
      End
      Begin VB.Label lblMsg 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3810
         TabIndex        =   2
         Top             =   240
         Width           =   4650
      End
      Begin VB.Label lblTitle 
         Caption         =   "WF�}�b�v��ԕ\��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   210
         TabIndex        =   1
         Top             =   270
         Width           =   4575
      End
   End
   Begin VB.Frame fraF 
      Height          =   1095
      Left            =   30
      TabIndex        =   12
      Top             =   9540
      Width           =   15195
      Begin VB.CommandButton cmdF 
         Caption         =   "[F12]�@�@�@ ����"
         Height          =   735
         Index           =   12
         Left            =   13920
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�R]�@�@�@������"
         Height          =   735
         Index           =   3
         Left            =   2824
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�Q]�@�@�@������"
         Height          =   735
         Index           =   2
         Left            =   1592
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�P]�@�@�@������"
         Height          =   735
         Index           =   0
         Left            =   360
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�X]�@�@�@������"
         Height          =   735
         Index           =   9
         Left            =   10216
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�S]�@�@�@������"
         Height          =   735
         Index           =   4
         Left            =   4056
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�T]�@�@�@������"
         Height          =   735
         Index           =   5
         Left            =   5288
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�U]�@�@�@������"
         Height          =   735
         Index           =   6
         Left            =   6520
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�V]�@�@�@������"
         Height          =   735
         Index           =   7
         Left            =   7752
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�W]�@�@�@������"
         Height          =   735
         Index           =   8
         Left            =   8984
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F10]�@�@�@������"
         Height          =   735
         Index           =   10
         Left            =   11448
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F11]�@�@�@������"
         Height          =   735
         Index           =   11
         Left            =   12680
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  '�s����
      BorderStyle     =   3  '�_��
      Height          =   285
      Left            =   7200
      Top             =   9240
      Width           =   1155
   End
   Begin VB.Label Label4 
      Caption         =   "���Ԕ���WF"
      Height          =   255
      Left            =   8400
      TabIndex        =   25
      Top             =   9270
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "SXLID"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "���o����"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1755
      TabIndex        =   8
      Top             =   840
      Width           =   945
   End
   Begin VB.Label Label3 
      Caption         =   "����WF"
      Height          =   255
      Left            =   6075
      TabIndex        =   6
      Top             =   9270
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "����WF"
      Height          =   255
      Left            =   3765
      TabIndex        =   5
      Top             =   9270
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "�ʏ�WF"
      Height          =   255
      Left            =   1605
      TabIndex        =   4
      Top             =   9270
      Width           =   705
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  '�s����
      BorderStyle     =   3  '�_��
      Height          =   285
      Left            =   4875
      Top             =   9240
      Width           =   1155
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  '�s����
      BorderStyle     =   3  '�_��
      Height          =   285
      Left            =   2565
      Top             =   9240
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  '�s����
      BorderStyle     =   3  '�_��
      Height          =   285
      Left            =   405
      Top             =   9240
      Width           =   1155
   End
End
Attribute VB_Name = "f_cmbc039_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*******************************************************************************
'*    �֐���        : cmbSprChg_Click
'*
'*    �����T�v      : 1.���o�����ɂ��AWFϯ�߈ꗗ����ʖ��̕\���̐؂�ւ����s��
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub cmbSprChg_Click()
    Dim intLoopCnt  As Integer
    Dim intSprSta   As Integer
    Dim sSprSta     As String
    Dim vSprSta     As Variant
    Dim intRowNo    As Integer
    
    intRowNo = 0
    
'Chg Start 2011/03/11 SMPK Miyata
'    With sprExamine
    With sprWfmapView
'Chg End   2011/03/11 SMPK Miyata
        .ReDraw = False
        For intLoopCnt = 1 To .MaxRows
            Select Case cmbSprChg.ListIndex
                Case intConSprChg_0  '�S���w��
                    .row = intLoopCnt
                    .RowHidden = False
                    intRowNo = intRowNo + 1
                    .RowHidden = False
                    .row = intLoopCnt
                    .col = 0
                    .text = intRowNo
                Case intConSprChg_1  '�Ǖi�w��
                    .GetText 30, intLoopCnt, vSprSta
                    If vSprSta <> intConSprChg_1 Then  '�Ǖi�ȊO��������A��\��
                        .row = intLoopCnt
                        .RowHidden = True
                    Else
                        .row = intLoopCnt
                        intRowNo = intRowNo + 1
                        .RowHidden = False
                        .row = intLoopCnt
                        .col = 0
                        .text = intRowNo
                    End If
                Case intConSprChg_2  '�T���v���w��
                    .GetText 30, intLoopCnt, vSprSta
                    If vSprSta <> intConSprChg_2 Then  '�T���v���ȊO��������A��\��
                        .row = intLoopCnt
                        .RowHidden = True
                    Else
                        .row = intLoopCnt
                        intRowNo = intRowNo + 1
                        .RowHidden = False
                        .row = intLoopCnt
                        .col = 0
                        .text = intRowNo
                    End If
                Case intConSprChg_3  '�s�ǎw��
                    .GetText 30, intLoopCnt, vSprSta
                    If vSprSta <> intConSprChg_3 Then  '�s�ǈȊO��������A��\��
                        .row = intLoopCnt
                        .RowHidden = True
                    Else
                        .row = intLoopCnt
                        intRowNo = intRowNo + 1
                        .RowHidden = False
                        .row = intLoopCnt
                        .col = 0
                        .text = intRowNo
                    End If
            End Select
        Next
        
        If .MaxRows > 0 Then
            .col = 1
            .row = 1
            .Action = ActionActiveCell
        End If
        
        .ReDraw = True
    End With
End Sub

'*******************************************************************************
'*    �֐���        : cmdF_Click
'*
'*    �����T�v      : 1.�t�@���N�V�����{�^�����N���b�N���ꂽ��A�e�����ɕ��򂷂�
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    Index       ,I  ,Integer�@,�R���g���[���z��̓Y��
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub cmdF_Click(Index As Integer)
    '' ��������
    Select Case Index
    Case 12       '' �e12�L�[�i���s�j
        Me.Visible = False
        Unload Me
    End Select
End Sub

'*******************************************************************************
'*    �֐���        : Form_Activate
'*
'*    �����T�v      : 1.Form_Activate����
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub Form_Activate()
    cmbSprChg.ListIndex = 0
'Chg Start 2011/03/11 SMPK Miyata
'    With sprExamine
    With sprWfmapView
'Chg End   2011/03/11 SMPK Miyata
        .col = 30
        .ColHidden = True
    End With
End Sub

'*******************************************************************************
'*    �֐���        : Form_KeyDown
'*
'*    �����T�v      : 1.�L�[�{�[�h��������
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    KeyCode     ,I  ,Integer�@,�L�[�R�[�h
'*                    Shift       ,I  ,Integer�@,Shift�L�[�̏��
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intIndex As Integer

    '' �t�@���N�V�����L�[���L���Ȃ�
    If KeyCode >= vbKeyF1 And KeyCode <= vbKeyF12 Then
        '' ��ʕ\�����b�Z�[�W�N���A
        lblMsg.Caption = ""
        
        intIndex = KeyCode - (vbKeyF1 - 1)
        If cmdF(intIndex).Visible = True And cmdF(intIndex).Enabled = True Then
            '' �t�@���N�V�����L�[�������������s����
            If KeyCode <> vbKeyF7 And KeyCode <> vbKeyF8 Then
                Call cmdF_Click(intIndex)
            End If
        End If
    End If
End Sub

'*******************************************************************************
'*    �֐���        : Form_Load
'*
'*    �����T�v      : 1.Form_Load����
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub Form_Load()
    ' ���ݓ����̕\��
    SetPresentTime lblTime
End Sub


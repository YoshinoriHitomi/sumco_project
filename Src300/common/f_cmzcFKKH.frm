VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form f_cmzcFKKH 
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "�U�։\���i��"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.CommandButton cmdKettei 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�L�����Z��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdKouho 
      Caption         =   "���i�ԕ\��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   1920
      Width           =   1455
   End
   Begin FPSpread.vaSpread sprHinban 
      Height          =   2295
      Left            =   840
      TabIndex        =   1
      Top             =   2760
      Width           =   2295
      _Version        =   196608
      _ExtentX        =   4048
      _ExtentY        =   4048
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      MaxCols         =   1
      ScrollBars      =   2
      SelectBlockOptions=   2
      ShadowColor     =   14215660
      ShadowDark      =   10070188
      ShadowText      =   0
      SpreadDesigner  =   "f_cmzcFKKH.frx":0000
      VisibleCols     =   1
      VisibleRows     =   500
   End
   Begin VB.TextBox txtMotoHinban 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "INS0017A00Y1"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  '��������
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      Caption         =   "�U�֌��i��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      Caption         =   "�U�։\���i��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "f_cmzcFKKH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'                                     2003/09/01
'======================================================
' �U�։\���i��
' �T�v    : �U�֌��i�Ԃ��U�։\���i�Ԃ��ꗗ�\�����A
'           �U�֐�i�ԂƂ��Č��肷��B
' �Q��    :
'======================================================

'�T�v      :Form Load������
'����      :�����\��
'����      :2003/09/01  ���c �쐬
Private Sub Form_Load()
    '' �U�֌��i�Ԃ�ݒ肷��
    txtMotoHinban.text = FKKH_MotoHinban
    
    '' �����ݒ�
    '''  �i��
    With sprHinban
        .MaxRows = 0
        .col = -1
        .row = -1
        .Lock = True
        .RowHeight(-1) = 12.27
    End With
    '''  ����{�^��
    cmdKettei.Enabled = False

    '' �\���ʒu
    Me.Move 9000, 3540
End Sub

'�T�v      :���i�ԕ\���{�^������������
'����      :�U�։\�ȕi�Ԃ��ꗗ�ɕ\������
'����      :2003/09/01  ���c �쐬
Private Sub cmdKouho_Click()
    Dim RET As Integer
    Dim ErrCode As Integer
    Dim ErrMsg As String
    
    ' �}�E�X�|�C���^�������v�ɕύX
    Screen.MousePointer = vbHourglass

    '' �U�֌��i�Ԏ擾(�d�l�`�F�b�N)���ʊ֐�
    RET = fncGetKouhoHinbanShiyou(FKKH_Proccd, FKKH_Crynum, FKKH_MotoHinban, KouhoHinban(), ErrCode, ErrMsg)
    
    If RET <> 0 Then
        Screen.MousePointer = vbDefault
        Call MsgBox(ErrMsg, vbOKOnly, "�U�։\���i��")
        Exit Sub
    End If
    
    '' �U�֌��i�Ԃ��ꗗ�ɕ\������
    Call FurikaeKouhoSet

    ' �}�E�X�|�C���^�����ɖ߂�
    Screen.MousePointer = vbDefault
    
    '' ����{�^��
    cmdKettei.Enabled = True

End Sub

'�T�v      :�ꗗ�\��
'����      :�U�֌��i�Ԃ��ꗗ�ɕ\������
'����      :2003/09/01  ���c �쐬
Private Sub FurikaeKouhoSet()
    Dim tblCnt As Long
    Dim cnt As Long
    
    With sprHinban
        .ReDraw = False
        .MaxRows = 0
        
        tblCnt = UBound(KouhoHinban)
        .MaxRows = tblCnt + 1
                
        For cnt = 0 To tblCnt
            .row = cnt + 1
            
            '�U�֌��i��
            .col = 1
            .text = KouhoHinban(cnt).GETHINBAN
        Next
        .ReDraw = True
    End With
End Sub

'�T�v      :����{�^������������
'����      :�U�֐�i�ԂƂ��Č��肵�A�ďo����ʂɖ߂�
'����      :2003/09/01  ���c �쐬
Private Sub cmdKettei_Click()
    '' �U�֐�i�Ԃ�ݒ肷��
    With sprHinban
        .row = .ActiveRow
        .col = 1
        FKKH_SakiHinban = .text
    End With
    
    Unload Me
End Sub

'�T�v      :�L�����Z���{�^������������
'����      :�ďo����ʂɖ߂�
'����      :2003/09/01  ���c �쐬
Private Sub cmdCancel_Click()
    '' �U�֐�i�Ԃ��N���A����
    FKKH_SakiHinban = ""
    
    Unload Me
End Sub

' @(f)
'
' �@�\      : �X�v���b�h�V�[�g�N���b�N
'
' �Ԃ�l    : �Ȃ�
'
' ������    :
'
' �@�\����  : �C�x���g�֐�
'
' ���l      : �X�v���b�h�V�[�g�̃\�[�g����  2008/05/28 �ǉ�:Kameda
'
Private Sub sprHinban_click(ByVal col As Long, ByVal row As Long)
    
    '�X�v���b�h�V�[�g�̕\�����X�V���Ȃ�
    sprHinban.ReDraw = False
    Select Case row
        'P1 ��^�C�g�������������ꍇ�A�������ꂽ������Ƀ\�[�g
        Case 0
            'Call sprSort(sprHinban, col)
            With sprHinban
                .BlockMode = True                               '  �Z���u���b�N��L��
                .col = 1                                        '  ���ݒ�
                .col2 = .MaxCols                                '  �ŏI���ݒ�
                .row = 1                                        '  �s��ݒ�
                .row2 = .MaxRows                                '  �ŏI�s��ݒ�
                .SortBy = SortByRow                             '  �s�P�ʂɕ��ёւ�
                .SortKey(1) = col                               '  ���ёւ��̃L�[��ݒ�
                
                If .SortKey(1) = col And .SortKeyOrder(1) = SortKeyOrderAscending Then
                    .SortKeyOrder(1) = SortKeyOrderDescending   '  �~���ɕ��ёւ���ݒ�
                Else
                    .SortKeyOrder(1) = SortKeyOrderAscending    '  �����ɕ��ёւ���ݒ�
                End If
                
                .Action = ActionSort                            '  ���ёւ������s
                .BlockMode = False                              '  �Z���u���b�N�𖳌�
            End With
    End Select

    '�X�v���b�h�V�[�g�̕\�����X�V����
    sprHinban.ReDraw = True

End Sub



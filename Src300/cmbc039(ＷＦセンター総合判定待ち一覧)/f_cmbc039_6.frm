VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form f_cmbc039_6 
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   " f_cmbc039_6"
   ClientHeight    =   10875
   ClientLeft      =   1575
   ClientTop       =   1680
   ClientWidth     =   15270
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   725
   ScaleMode       =   3  '�߸��
   ScaleWidth      =   1018
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox txtDateT 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  '�̌Œ�
      Left            =   4200
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "20120808"
      Top             =   960
      Width           =   1140
   End
   Begin VB.TextBox txtDateF 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  '�̌Œ�
      Left            =   2400
      MaxLength       =   8
      TabIndex        =   0
      Text            =   "20120808"
      Top             =   960
      Width           =   1140
   End
   Begin VB.Frame fraHead 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15225
      Begin VB.Label lblvers 
         Height          =   195
         Left            =   13680
         TabIndex        =   24
         Top             =   480
         Width           =   1440
      End
      Begin VB.Label lblTime 
         Height          =   150
         Left            =   13680
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblMsg 
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3810
         TabIndex        =   4
         Top             =   240
         Width           =   7050
      End
      Begin VB.Label lblTitle 
         Caption         =   "�O�����b�g�ꗗ�\��"
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
         TabIndex        =   3
         Top             =   270
         Width           =   3495
      End
   End
   Begin VB.Frame fraF 
      Height          =   1095
      Left            =   30
      TabIndex        =   7
      Top             =   9540
      Width           =   15195
      Begin VB.CommandButton cmdF 
         Caption         =   "[F12]�@�@�@ ����"
         Height          =   735
         Index           =   12
         Left            =   13920
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�R]�@�@�@������"
         Height          =   735
         Index           =   3
         Left            =   2824
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�X]�@�@�@���o"
         Height          =   735
         Index           =   9
         Left            =   10216
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�S]�@�@�@������"
         Height          =   735
         Index           =   4
         Left            =   4056
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
   End
   Begin FPSpread.vaSpread spdZeroView 
      Height          =   7725
      Left            =   240
      TabIndex        =   20
      Top             =   1680
      Width           =   14655
      _Version        =   196608
      _ExtentX        =   25850
      _ExtentY        =   13626
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      ColsFrozen      =   4
      MaxCols         =   15
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "f_cmbc039_6.frx":0000
      UserResize      =   0
      VisibleCols     =   15
      VisibleRows     =   1
   End
   Begin VB.Label Label3 
      Caption         =   "(YYYYMMDD)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4140
      TabIndex        =   23
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label Label2 
      Caption         =   "(YYYYMMDD)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2340
      TabIndex        =   22
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "WF���o���F"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   990
      Width           =   1305
   End
   Begin VB.Label Label5 
      Caption         =   "�`"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   6
      Top             =   960
      Width           =   375
   End
End
Attribute VB_Name = "f_cmbc039_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===============================================================================
' �O�����b�g�ꗗ�\�����
' 2012/09/07 SETsw Marushita
' �T�v    :�@�T���v�������0�����b�g�ƂȂ����f�[�^���ꗗ�\������(10���ȉ����b�g�����Ή�)
'===============================================================================

'*******************************************************************************
'*    �֐���        : DispSpdZeroView
'*
'*    �����T�v      : 1.���o�����ɂ��A0�����b�g�ꗗ���擾���\������
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub DispSpdZeroView()
    Dim intLoopCnt  As Integer
    Dim sSprSta     As String
    Dim intRowNo    As Integer
    Dim sDateF      As String
    Dim sDateT      As String
    Dim sMukesaki   As String
    Dim smpId(2)    As String
    Dim sErrMsg     As String
    
    intRowNo = 0
    
    '�X�v���b�h�R���g���[���̏���������
    SpCtrlInit f_cmbc039_6.spdZeroView, 0

    '��ʂ��璊�o�������擾����
    If Trim(txtDateF.text) = "" Then
    Else
        '���͓��t�`�F�b�N
        If DateCheck(Trim(txtDateF.text), 0) = False Then
            lblMsg.Caption = "���������t����͂��Ă��������B"
            txtDateF.SetFocus
            Exit Sub
        End If
    End If
    If Trim(txtDateT.text) = "" Then
    Else
        '���͓��t�`�F�b�N
        If DateCheck(Trim(txtDateT.text), 0) = False Then
            lblMsg.Caption = "���������t����͂��Ă��������B"
            txtDateT.SetFocus
            Exit Sub
        End If
        '���͓��t�召�`�F�b�N
        If Trim(txtDateF.text) > Trim(txtDateT.text) Then
            lblMsg.Caption = "���������t�͈͂���͂��Ă��������B"
            txtDateF.SetFocus
            Exit Sub
        End If
    End If
    
    sDateF = txtDateF.text
    'TO���t�w�莞�͎��Ԃ�t��
    If Trim(txtDateT.text) = "" Then
        sDateT = txtDateT.text
    Else
        sDateT = txtDateT.text & " 23:59:59"
    End If
    
    '0�����b�g�ꗗ���擾
    lblMsg.Caption = GetMsgStr(PWAIT)
    DoEvents
    
    '���o�������w�肵�đΏۃf�[�^���擾����B
    If DBDRV_fcmbc039_6_Disp(gsMukeCd, sDateF, sDateT, typ_zero(), sErrMsg) = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = "�Ώۃf�[�^�擾�G���[�ł��B"
        Exit Sub
    End If
        
    If UBound(typ_zero) = 0 Then
        lblMsg.Caption = "�Ώۃf�[�^������܂���B"
        Exit Sub
    Else
        lblMsg.Caption = ""
    End If
    
    '�Ώۃf�[�^���ꗗ�\������
    With spdZeroView
        .ReDraw = False
        For intLoopCnt = 1 To UBound(typ_zero)
            intRowNo = intRowNo + 1
    
            sSprSta = ""

            If UBound(typ_zero(intLoopCnt).WFSMP) >= 1 Then
                smpId(1) = Trim(typ_zero(intLoopCnt).WFSMP(1).REPSMPLIDCW)
            Else
                smpId(1) = vbNullString
            End If
            If UBound(typ_zero(intLoopCnt).WFSMP) >= 2 Then
                smpId(2) = Trim(typ_zero(intLoopCnt).WFSMP(UBound(typ_zero(intLoopCnt).WFSMP)).REPSMPLIDCW)
            Else
                smpId(2) = vbNullString
            End If
                    
            .MaxRows = intRowNo
            .row = intRowNo
    
            '����
            .col = 1
            .SetText 1, intRowNo, typ_zero(intLoopCnt).PLANTCAT
    
            '��SXL-ID
            .col = 2
            .SetText 2, intRowNo, typ_zero(intLoopCnt).SXLIDCA
    
            '�i��
            .col = 3
            .SetText 3, intRowNo, typ_zero(intLoopCnt).HINBCA
            
            '����
            .col = 4
            .SetText 4, intRowNo, typ_zero(intLoopCnt).GNLCA

            '����
            .col = 5
            .SetText 5, intRowNo, typ_zero(intLoopCnt).MAICB
    
            'WF���o��
            .col = 6
            .SetText 6, intRowNo, Format(typ_zero(intLoopCnt).TDAYCB, "yyyy/mm/dd")
    
            '�T���v��ID(�㑤)
            .col = 7
            .SetText 7, intRowNo, smpId(1)
    
            '�T���v��ID(����)
            .col = 8
            .SetText 8, intRowNo, smpId(2)
    
            '�ŏI��M��(�㑤)
            If Not (smpId(1) = "" Or _
                left(smpId(1), 1) = vbNullChar) Then
                If UBound(typ_zero(intLoopCnt).WFSMP) >= 1 Then
                    .col = 9
                    .SetText 9, intRowNo, Format(typ_zero(intLoopCnt).WFSMP(1).KDAYCW, "yyyy/mm/dd")
                End If
            End If
    
            '�ŏI��M��(����)
            If Not (smpId(2) = "" Or _
                left(smpId(2), 1) = vbNullChar) Then
                If UBound(typ_zero(intLoopCnt).WFSMP) >= 2 Then
                    .col = 10
                    .SetText 10, intRowNo, Format(typ_zero(intLoopCnt).WFSMP(UBound(typ_zero(intLoopCnt).WFSMP)).KDAYCW, "yyyy/mm/dd")
                End If
            End If

            '���ݍH��(TEST)
            .col = 11
            .SetText 11, intRowNo, typ_zero(intLoopCnt).NOWPROC

'            '��(�㑤)
'            If typ_zero(intLoopCnt).KETURAKU = True Then
'                sSprSta = sSprSta & "�L" & Chr$(13) & Chr$(10)
'            Else
'                sSprSta = sSprSta & "��" & Chr$(13) & Chr$(10)
'            End If
                    
            '�f�[�^�\��
            '1�s�f�[�^�Z�b�g
            '.Clip = sSprSta
            
'            bRc = gFnc_SS_RecordSet(.spdWait, intRow, strRecord, udt_ww, i)
'            '�װ�����������ꍇ
'            If bRc = False Then
'                '�װ����
'                .lblMsg.Caption = "�\���G���["
'                Exit Sub
'            End If
            
        Next
        .MaxRows = intRowNo
        .ReDraw = True
        
        If .MaxRows > 0 Then
            .col = 1
            .row = 1
            .Action = ActionActiveCell
        End If
        
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
    Case 9        '' �e9�L�[�i���o�j
        Call DispSpdZeroView
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
    
    ' �o�[�W�������̕\��
    lblvers.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    ' ���o���t�̏����l�Z�b�g
    txtDateF.text = Format(DateAdd("m", -1, Date), "yyyymmdd")
    txtDateT.text = Format(Date, "yyyymmdd")

    ' �O�����b�g�ꗗ��ʂ̕\��
    Call DispSpdZeroView

End Sub


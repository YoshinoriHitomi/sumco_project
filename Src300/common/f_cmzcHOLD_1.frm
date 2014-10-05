VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form f_cmzcHOLD_1 
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "f_cmzcHOLD_1"
   ClientHeight    =   5370
   ClientLeft      =   1875
   ClientTop       =   2820
   ClientWidth     =   9285
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   358
   ScaleMode       =   3  '�߸��
   ScaleWidth      =   619
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton cmdEnd 
      Caption         =   "����"
      Height          =   375
      Left            =   7560
      TabIndex        =   4
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Frame FraTitle 
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.Label lblTitle 
         Caption         =   "�z�[���h�����Q��"
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
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
   End
   Begin FPSpread.vaSpread spdDisp 
      Height          =   3375
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   8655
      _Version        =   196608
      _ExtentX        =   15266
      _ExtentY        =   5953
      _StockProps     =   64
      BorderStyle     =   0
      ColsFrozen      =   3
      DisplayRowHeaders=   0   'False
      MaxCols         =   11
      OperationMode   =   1
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "f_cmzcHOLD_1.frx":0000
      VisibleCols     =   6
      VisibleRows     =   500
   End
   Begin VB.Label lblCrynum 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  '����
      Height          =   255
      Left            =   6000
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "�u���b�N�h�c"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      TabIndex        =   1
      Top             =   480
      Width           =   945
   End
End
Attribute VB_Name = "f_cmzcHOLD_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEnd_Click()
    Unload Me
End Sub

'Form�ݒ�l
Private Sub Form_Load()
    Dim sTblDispData() As typ_TBCMJ012
    Dim sCrynum As String
    ' �o�[�W�������̕\��
    
    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "f_cmzcHOLD_1.frm -- Sub Form_Load"
    Clear
    lblCrynum.Caption = f_cmbc030_2.txtBlockID.Text

    sCrynum = Left(Trim(lblCrynum.Caption), 9) & "000"
    If DBDRV_SELECT_HOLD(sTblDispData, sCrynum) = FUNCTION_RETURN_SUCCESS Then
        If UBound(sTblDispData) > 0 Then
            spdDispSet sTblDispData
        End If
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
Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "f_cmzcHOLD_1.frm -- Sub Form_Unload"

    Unload Me
proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub

Private Sub Clear()
Dim i As Integer

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "f_cmzcHOLD_1.frm -- Sub Clear"
    
    lblCrynum.Caption = vbNullString
    spdDisp.MaxRows = 0
    '���͍ς̃`�F�b�N

proc_exit:
    '�I��
    gErr.Pop
    Exit Sub

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Sub

'�ꗗ�\�Ɋ����̃f�[�^���Z�b�g����
Private Sub spdDispSet(gTblDispData() As typ_TBCMJ012)
    Dim i As Integer
    Dim sCount As Integer
    Dim sData As String
    Dim RET As FUNCTION_RETURN
    
        spdDisp.ReDraw = False
        For sCount = 1 To UBound(gTblDispData)
            sData = ""
            With spdDisp
                '�����ԍ�
                sData = gTblDispData(sCount).CRYNUM
                '�C���S�b�g���ʒu
                sData = sData & vbTab & gTblDispData(sCount).INGOTPOS
                '������
                sData = sData & vbTab & gTblDispData(sCount).TRANCNT
                '����
                sData = sData & vbTab & gTblDispData(sCount).LENGTH
                '�H���R�[�h
                sData = sData & vbTab & GetGPCodeDspStr(gTblDispData(sCount).PROCCODE, GetCodeFieldA9("K", "16", gTblDispData(sCount).PROCCODE, "NAMEJA9"))
                '�����敪
                If gTblDispData(sCount).HLDTRCLS = "0" Then
                    sData = sData & vbTab & "0:�z�[���h��������"
                Else
                     sData = sData & vbTab & GetGPCodeDspStr(gTblDispData(sCount).HLDTRCLS, GetCodeFieldA9("X", "31", gTblDispData(sCount).HLDTRCLS, "NAMEJA9"))
                End If
                '�������t
                sData = sData & vbTab & Format(gTblDispData(sCount).UPDDATE, "yy/mm/dd hh:nn")
                '�z�[���h�S����
                sData = sData & vbTab & GetStaffName(gTblDispData(sCount).KSTAFFID)

                '�z�[���h���R
                sData = sData & vbTab & GetGPCodeDspStr(gTblDispData(sCount).HLDCAUSE, GetCodeFieldA9("X", "30", gTblDispData(sCount).HLDCAUSE, "NAMEJA9"))
                '�z�[���h�R�����g
                sData = sData & vbTab & gTblDispData(sCount).HLDCMNT
                '�z�[���h�H���R�[�h
                sData = sData & vbTab & GetGPCodeDspStr(gTblDispData(sCount).HOLDKT, GetCodeFieldA9("K", "16", gTblDispData(sCount).HOLDKT, "NAMEJA9"))
                
                sData = sData & vbCr
                
                .MaxRows = sCount
                .Row = sCount
                .row2 = sCount
                .col = 1
                .col2 = .MaxCols
                .ClipValue = sData
            End With
        Next
        spdDisp.ReDraw = True
End Sub
Private Function DBDRV_SELECT_HOLD(gTblDispData() As typ_TBCMJ012, pCrynum As String) As FUNCTION_RETURN

    Dim sql As String       'SQL�S��
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      '���R�[�h��
    Dim i As Long

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_XSDC1_SQL.bas -- Function DBDRV_SELECT_HOLD"
    
     ''SQL��g�ݗ��Ă�
     'sql = "SELECT PROCCODE, HLDTRCLS, HLDCAUSE, HLDCMNT, UPDDATE, KSTAFFID, HOLDKT FROM TBCMJ012,XSDC2 "
     sql = "SELECT CRYNUM, INGOTPOS, TRANCNT, LENGTH, PROCCODE, HLDTRCLS, HLDCAUSE, HLDCMNT, UPDDATE, KSTAFFID, HOLDKT "
     sql = sql & "FROM TBCMJ012 "
     'sql = sql & " WHERE CRYNUMC2 = '" & pCrynum & "'"
     sql = sql & " WHERE CRYNUM = '" & pCrynum & "'"
     'sql = sql & " AND   CRYNUM = XTALC2 "
     'sql = sql & " AND   INGOTPOS = INPOSC2 "
    
    ''�f�[�^�𒊏o����
     Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
     
     If rs Is Nothing Then
         DBDRV_SELECT_HOLD = FUNCTION_RETURN_FAILURE
         Exit Function
     End If
     If rs.RecordCount > 0 Then
        recCnt = rs.RecordCount
        ReDim gTblDispData(recCnt)
        For i = 1 To recCnt
            With gTblDispData(i)
                .CRYNUM = rs("CRYNUM")
                .INGOTPOS = rs("INGOTPOS")
                .TRANCNT = rs("TRANCNT")
                .LENGTH = rs("LENGTH")
                If IsNull(rs("HLDCAUSE")) = False Then .HLDCAUSE = rs("HLDCAUSE")
                If IsNull(rs("HLDCMNT")) = False Then .HLDCMNT = rs("HLDCMNT")
                If IsNull(rs("HOLDKT")) = False Then
                   .HOLDKT = rs("HOLDKT")
                Else
                   .HOLDKT = Space(5)
                End If
                If IsNull(rs("HLDTRCLS")) = False Then .HLDTRCLS = rs("HLDTRCLS")
                If IsNull(rs("PROCCODE")) = False Then .PROCCODE = rs("PROCCODE")
                If IsNull(rs("UPDDATE")) = False Then .UPDDATE = rs("UPDDATE")
                If IsNull(rs("KSTAFFID")) = False Then .KSTAFFID = rs("KSTAFFID")
            End With
            rs.MoveNext
        Next
     Else
        ReDim gTblDispData(0)
     End If
    rs.Close

    DBDRV_SELECT_HOLD = FUNCTION_RETURN_SUCCESS
    

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    DBDRV_SELECT_HOLD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function



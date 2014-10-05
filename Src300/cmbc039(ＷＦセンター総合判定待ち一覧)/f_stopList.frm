VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "spr32x30.ocx"
Begin VB.Form f_stopList 
   Caption         =   "流動停止一覧　"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10710
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton btnClose 
      Caption         =   "閉じる"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Frame FraTitle 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      Begin VB.Label lblTitle 
         Caption         =   "流動停止一覧　"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblMsg 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3600
         TabIndex        =   1
         Top             =   240
         Width           =   2550
      End
   End
   Begin FPSpread.vaSpread spdList 
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   10455
      _Version        =   196608
      _ExtentX        =   18441
      _ExtentY        =   5741
      _StockProps     =   64
      ColsFrozen      =   2
      DisplayRowHeaders=   0   'False
      MaxCols         =   9
      MaxRows         =   6
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "f_stopList.frx":0000
      VisibleCols     =   1
      VisibleRows     =   2
   End
   Begin VB.Label lblSxl 
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "仮ＳＸＬ－ＩＤ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   900
      Width           =   1095
   End
End
Attribute VB_Name = "f_stopList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objIE   As Object

Private Sub btnClose_Click()
    Unload Me
End Sub

'概要      :流動停止一覧画面の表示
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :sProc         ,I  ,String    ,表示対象工程
'          :sBlockID      ,I  ,String    ,ブロックID
'          :sSxlID        ,I  ,String    ,SXL-ID
'          :戻り値        なし
'説明      :
Public Sub ShowStopList(ByVal sProc As String, ByVal sBlockID As String _
                    , ByVal sSxlID As String)

    Dim tXODY4()    As typ_XODY4
    Dim sWhere      As String
    Dim sOrder      As String
    Dim sSplit()    As String
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_stopList.frm -- Function ShowStopList"
    
    lblSxl.Caption = sSxlID
    
    '一覧取得
    sWhere = "WHERE "
    sWhere = sWhere & "Y4.STOPY4 <> '2' AND "
    If Trim(sBlockID) <> "" Then
        sWhere = sWhere & "Y4.XTALNOY4 = '" & sBlockID & "' AND "
    End If
    If Trim(sSxlID) <> "" Then
        sWhere = sWhere & "Y4.SXLIDY4 = '" & sSxlID & "' AND "
    End If
    sSplit = Split(sProc, "/")
    If UBound(sSplit) > 0 Then
        sWhere = sWhere & "Y4.WKKTY4 in ("
        For i = 0 To UBound(sSplit) - 1
            sWhere = sWhere & "'" & sSplit(i) & "'"
            If i <> (UBound(sSplit) - 1) Then
                sWhere = sWhere & ","
            End If
        Next
        sWhere = sWhere & ")"
    End If
    sOrder = "ORDER BY NVL(A9_4.CTR02A9,0)"
    If GetXODY4LeftJoin(tXODY4, sWhere, sOrder) = False Then
        GoTo proc_exit
    End If
    
    '一覧表示
    Call DrawList(tXODY4)
    
    Me.Show 1

proc_exit:
    '終了
    Unload Me
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

Private Sub DrawList(tData() As typ_XODY4)
    Dim iRow    As Integer
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_stopList.frm -- Function DrawList"
    
    
    spdList.MaxRows = UBound(tData)
    For iRow = 1 To UBound(tData)
        With spdList
            .SetText 1, iRow, tData(iRow).WKKTY4 & ":" & tData(iRow).WKKTNAME
            .SetText 2, iRow, SetStopName(tData(iRow).STOPY4)
            .SetText 3, iRow, SetAgrStatusName(tData(iRow).AGRSTATUSY4)
            .SetText 4, iRow, Format(tData(iRow).SDAYY4, "yy/mm/dd")
            .SetText 5, iRow, tData(iRow).SETSTAFFNAME
            .SetText 6, iRow, tData(iRow).SETMEMOY4
            .SetText 7, iRow, tData(iRow).CAUSEY4 & ":" & tData(iRow).CAUSENAME
            If Trim(tData(iRow).PRINTNOY4) <> "" Then
                .Col = 8
                .Row = iRow
                .CellType = CellTypeButton
                .TypeButtonText = tData(iRow).PRINTNOY4
                .Lock = False
            Else
                .SetText 8, iRow, tData(iRow).PRINTNOY4
            End If
            .SetText 9, iRow, tData(iRow).PRINTKINDY4
        End With
        
    Next
    
proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit

End Sub

Private Sub spdList_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim sPrintNo    As Variant
    Dim sPrintKind  As Variant
    If Col = 8 And Row > 0 Then
        Call CloseIE
        spdList.Col = 8
        spdList.Row = Row
        sPrintNo = spdList.TypeButtonText
        spdList.GetText 9, Row, sPrintKind
        Set objIE = CreateObject("InternetExplorer.application")
        sVal = GetSWSUrl(CStr(sPrintKind), CStr(sPrintNo))
        objIE.Navigate sVal
        objIE.AddressBar = False
        objIE.MenuBar = False
        objIE.StatusBar = False
        objIE.ToolBar = False
        objIE.Visible = True
    End If

End Sub

Private Sub CloseIE()
    On Error Resume Next
    Dim iCnt As Integer
    If Not objIE Is Nothing Then
        objIE.Quit
    End If
End Sub


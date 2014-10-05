VERSION 5.00
Begin VB.Form WFCJudgDialog 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
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
'*    関数名        : WFCErrorMessage
'*
'*    処理概要      : 1.ErrMsgを表示する
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*　　　　　　　　　　ErrMsg      ,I  ,String   ,ErrMsg
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Public Sub WFCErrorMessage(ErrMsg As String)
    List1.AddItem ErrMsg, List1.ListCount
End Sub

'*******************************************************************************
'*    関数名        : OKButton_Click
'*
'*    処理概要      : 1.画面を隠す
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*　　　　　　　　　　なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub OKButton_Click()
    Me.Visible = False
End Sub

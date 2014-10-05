VERSION 5.00
Begin VB.Form f_cmzc003a 
   Caption         =   "結晶図 (123456789012)"
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
      Alignment       =   1  '右揃え
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
' 結晶図ウィンドウ
' 概要    : 与えられた結晶クラスの内容を図示する
' 参照    : 結晶コントロール(o_cmzc002a.ctl)
'         : 結晶クラス      (c_cmzcXl.ctl)
'======================================================

'レジストリ保存に関する定数定義
Private Const SYSTEM_NAME = "300mm結晶操業システム"
Private Const CATEGORY = "ウィンドウ位置"
'レジストリ保存位置は　HKEY_CURRENT_USER\Software\VB and VBA Program Settings\<SYSTEM_NAME>\<CATEGORY>

'*******************************************************************************
'*    関数名        : Form_KeyUp
'*
'*    処理概要      : 1.ESCキーが押されたら、ウィンドウを閉じる
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    KeyCode     ,I  ,Integer　,キーコード
'*                    Shift       ,I  ,Integer　,Shiftキーの状態
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmzc003a.frm -- Sub Form_KeyUp"

    '' ESCキーが押されたら、ウィンドウを閉じる
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'*******************************************************************************
'*    関数名        : Form_Load
'*
'*    処理概要      : 1.結晶図コントロールの表示位置を前回に合わせる
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub Form_Load()
    Dim lngOldTop   As Long
    Dim lngOldLeft  As Long

    '' 結晶図コントロールの表示位置を前回に合わせる

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmzc003a.frm -- Sub Form_Load"

    lngOldTop = CLng(GetSetting(SYSTEM_NAME, CATEGORY, Me.Name & "-Top", "-1"))    ''Top位置の復元
    lngOldLeft = CLng(GetSetting(SYSTEM_NAME, CATEGORY, Me.Name & "-Left", "-1"))   ''Left位置の復元
    If (lngOldTop > 0) And (lngOldLeft > 0) Then
        Me.Move lngOldLeft, lngOldTop
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'*******************************************************************************
'*    関数名        : Form_Resize
'*
'*    処理概要      : 1.コントロールの位置・大きさを連動する
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub Form_Resize()
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmzc003a.frm -- Sub Form_Resize"

    On Error Resume Next
    
    '' 表示時刻をの大きさを連動する
    lblDspTime.left = Width - lblDspTime.Width - 100
    
    '' 結晶図コントロールの大きさを連動する
    PicXL.Width = Width - PicXL.left - 100      ''幅を連動する
    PicXL.Height = Height - PicXL.top - 400     ''高さを連動する

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'*******************************************************************************
'*    関数名        : Form_Unload
'*
'*    処理概要      : 1.結晶図ウィンドウの表示位置を記憶する
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    Cancel        ,I  ,Integer
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmzc003a.frm -- Sub Form_Unload"

    ''画面を閉じるときに、表示位置を記憶しておく
    SaveSetting SYSTEM_NAME, CATEGORY, Me.Name & "-Top", top     ''Top位置を保存
    SaveSetting SYSTEM_NAME, CATEGORY, Me.Name & "-Left", left   ''Left位置を保存

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'*******************************************************************************
'*    関数名        : Clear
'*
'*    処理概要      : 1.結晶図ウィンドウを初期化する
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Public Sub Clear()
    '' フォームのキャプションを初期化する

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmzc003a.frm -- Sub Clear"

    Caption = "結晶図 (____________)"
    
    '' 表示時刻を初期化する
    lblDspTime.Caption = Format$(Now, "m/d  h:m")
    
    '' 結晶図を初期化する
    PicXL.Clear

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'*******************************************************************************
'*    関数名        : Draw
'*
'*    処理概要      : 1.結晶図ウィンドウを描画する
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    xl　　　　    ,I  ,c_cmzcXl ,結晶情報
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Public Sub Draw(Xl As c_cmzcXl)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmzc003a.frm -- Sub Draw"

    '' フォームのキャプションを設定する
    Caption = "結晶図 (" & Xl.CRYNUM & ")"
    
    '' 表示時刻を設定する
    lblDspTime.Caption = Format$(Now, "m/d  h:m")
    
    '' 結晶図を描画する
    PicXL.Draw Xl

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

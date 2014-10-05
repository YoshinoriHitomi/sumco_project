VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_cmzcRej"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'                                     2001/06/23
'================================================
' TBL情報クラスの宣言
' 定義内容: cmzcRej -- 欠落情報
' 参照    : 060207_結晶管理
'================================================

' 単純公開
Public LOTID As String              ' ブロックID
Public ALLSCRAP As String           ' 全数スクラップ
Public LENFROM As Integer           ' 長さ　FROM
Public LENTO As Integer             ' 長さ　TO


'概要      :自己複製
'説明      :自オブジェクトと同内容のオブジェクトを生成する
'履歴      :2001/08/17 作成  野村
Public Function Clone() As c_cmzcRej
Dim CloneTo As New c_cmzcRej   '複製先

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcRej.cls -- Function Clone"

    '' 新規オブジェクトに単純公開プロパティを複写する
    With CloneTo
        .LOTID = LOTID          ' ブロックID
        .ALLSCRAP = ALLSCRAP    ' 全数スクラップ
        .LENFROM = LENFROM      ' 長さ　FROM
        .LENTO = LENTO          ' 長さ　TO
    End With

    '' オブジェクトを返す
    Set Clone = CloneTo

PROC_EXIT:
    '終了
    Set CloneTo = Nothing
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Function

'概要      :指定位置は欠落しているか
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,Boolean   ,
'説明      :
'履歴      :2001/08/17 作成  野村
Public Function IsInArea(ByVal pos%) As Boolean

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcRej.cls -- Function IsInArea"
    
    If Trim$(ALLSCRAP) = vbNullString Then
        IsInArea = True
    ElseIf (pos >= LENFROM) And (pos <= LENTO) Then
        IsInArea = True
    Else
        IsInArea = False
    End If

PROC_EXIT:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Function


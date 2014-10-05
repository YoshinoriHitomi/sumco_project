VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_cmzcBlkPlan"
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
' 定義内容: cmzcBlkPlan -- ブロック設計
' 参照    : 060207_結晶管理
'================================================

' 単純公開
Public CRYNUM As String             ' 結晶番号
Public IngotPos As Integer          ' 結晶内開始位置
Public LENGTH As Integer            ' 長さ
Public USECLASS As String           ' 使用区分
'Public REGDATE As Date              ' 登録日付
'Public UPDDATE As Date              ' 更新日付
'Public SENDFLAG As String           ' 送信フラグ
'Public SENDDATE As Date             ' 送信日付


'概要      :自己複製
'説明      :自オブジェクトと同内容のオブジェクトを生成する
'履歴      :2001/06/23 作成  野村
Public Function Clone() As c_cmzcBlkPlan
Attribute Clone.VB_Description = "; 自オブジェクトと同内容のオブジェクトを生成します; "
Dim CloneTo As New c_cmzcBlkPlan   '複製先

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcBlkPlan.cls -- Function Clone"

    '' 新規オブジェクトに単純公開プロパティを複写する
    With CloneTo
        .CRYNUM = CRYNUM            ' 結晶番号
        .IngotPos = IngotPos        ' 結晶内開始位置
        .LENGTH = LENGTH            ' 長さ
        .USECLASS = USECLASS        ' 使用区分
'        .REGDATE = REGDATE          ' 登録日付
'        .UPDDATE = UPDDATE          ' 更新日付
'        .SENDFLAG = SENDFLAG        ' 送信フラグ
'        .SENDDATE = SENDDATE        ' 送信日付
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

'概要      :指定位置を含んでいるか
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,Boolean   ,
'説明      :
'履歴      :2001/06/23 作成  野村
Public Function IsInArea(ByVal pos%) As Boolean

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcBlkPlan.cls -- Function IsInArea"

    If (pos >= IngotPos) And (pos < IngotPos + LENGTH) Then
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


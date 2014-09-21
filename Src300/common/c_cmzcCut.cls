VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_cmzcCut"
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
' 定義内容: cmzcCut -- 切断指示
' 参照    : 060207_結晶管理
'================================================

' 単純公開
Public CRYNUM As String             ' 結晶番号
Public IngotPos As Integer          ' 結晶内開始位置
Public TRANCNT As Integer           ' 処理回数
Public LENGTH As Integer            ' 長さ
Public KRPROCCD As String           ' 管理工程コード
Public PROCCODE As String           ' 工程コード
Public STAFFID As String            ' 社員ID
Public hinban As String             ' 上品番
Public REVNUM As Integer            ' 上品番製品番号改訂番号
Public factory As String            ' 上品番工場
Public opecond As String            ' 上品番操業条件
Public BDCAUS As String             ' 区分コード
Public STATCLS As String            ' 状態区分
Public BLOCKID As String            ' ブロックID
Public CRYINDRS As String           ' 結晶検査指示（Rs)
Public CRYINDOI As String           ' 結晶検査指示（Oi)
Public CRYINDB1 As String           ' 結晶検査指示（B1)
Public CRYINDB2 As String           ' 結晶検査指示（B2）
Public CRYINDB3 As String           ' 結晶検査指示（B3)
Public CRYINDL1 As String           ' 結晶検査指示（L1)
Public CRYINDL2 As String           ' 結晶検査指示（L2)
Public CRYINDL3 As String           ' 結晶検査指示（L3)
Public CRYINDL4 As String           ' 結晶検査指示（L4)
Public CRYINDCS As String           ' 結晶検査指示（Cs)
Public CRYINDGD As String           ' 結晶検査指示（GD)
Public CRYINDT As String            ' 結晶検査指示（T)
Public CRYINDEP As String           ' 結晶検査指示（EPD)
Public PRIORITY As String           ' 優先度
'Public REGDATE As Date              ' 登録日付
'Public UPDDATE As Date              ' 更新日付
'Public SENDFLAG As String           ' 送信フラグ
'Public SENDDATE As Date             ' 送信日付


'概要      :自己複製
'説明      :自オブジェクトと同内容のオブジェクトを生成する
'履歴      :2001/06/23 作成  野村
Public Function Clone() As c_cmzcCut
Attribute Clone.VB_Description = "; 自オブジェクトと同内容のオブジェクトを生成します; "
Dim CloneTo As New c_cmzcCut   '複製先
    
    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcCut.cls -- Function Clone"

    '' 新規オブジェクトに単純公開プロパティを複写する
    With CloneTo
        .CRYNUM = CRYNUM            ' 結晶番号
        .IngotPos = IngotPos        ' 結晶内開始位置
        .TRANCNT = TRANCNT          ' 処理回数
        .LENGTH = LENGTH            ' 長さ
        .KRPROCCD = KRPROCCD        ' 管理工程コード
        .PROCCODE = PROCCODE        ' 工程コード
        .STAFFID = STAFFID          ' 社員ID
        .hinban = hinban            ' 上品番
        .REVNUM = REVNUM            ' 上品番製品番号改訂番号
        .factory = factory          ' 上品番工場
        .opecond = opecond          ' 上品番操業条件
        .BDCAUS = BDCAUS            ' 区分コード
        .STATCLS = STATCLS          ' 状態区分
        .BLOCKID = BLOCKID          ' ブロックID
        .CRYINDRS = CRYINDRS        ' 結晶検査指示（Rs)
        .CRYINDOI = CRYINDOI        ' 結晶検査指示（Oi)
        .CRYINDB1 = CRYINDB1        ' 結晶検査指示（B1)
        .CRYINDB2 = CRYINDB2        ' 結晶検査指示（B2）
        .CRYINDB3 = CRYINDB3        ' 結晶検査指示（B3)
        .CRYINDL1 = CRYINDL1        ' 結晶検査指示（L1)
        .CRYINDL2 = CRYINDL2        ' 結晶検査指示（L2)
        .CRYINDL3 = CRYINDL3        ' 結晶検査指示（L3)
        .CRYINDL4 = CRYINDL4        ' 結晶検査指示（L4)
        .CRYINDCS = CRYINDCS        ' 結晶検査指示（Cs)
        .CRYINDGD = CRYINDGD        ' 結晶検査指示（GD)
        .CRYINDT = CRYINDT          ' 結晶検査指示（T)
        .CRYINDEP = CRYINDEP        ' 結晶検査指示（EPD)
        .PRIORITY = PRIORITY        ' 優先度
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
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcCut.cls -- Function IsInArea"

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


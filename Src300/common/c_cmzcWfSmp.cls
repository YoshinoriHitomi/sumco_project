VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_cmzcWfSmp"
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
' 定義内容: cmzcWfSmp -- WFサンプル管理
' 参照    : 060207_結晶管理
'================================================

' 単純公開
Public CRYNUM As String             ' 結晶番号
Public INGOTPOS As Integer          ' 結晶内位置
Public SMPKBN As String             ' サンプル区分
Public SMPLID As String             ' サンプルID
Public hinban As String             ' 品番
Public REVNUM As Integer            ' 製品番号改訂番号
Public factory As String            ' 工場
Public opecond As String            ' 操業条件
Public KTKBN As String              ' 確定区分
Public WFINDRS As String            ' WF検査指示（Rs)
Public WFINDOI As String            ' WF検査指示（Oi)
Public WFINDB1 As String            ' WF検査指示（B1)
Public WFINDB2 As String            ' WF検査指示（B2）
Public WFINDB3 As String            ' WF検査指示（B3)
Public WFINDL1 As String            ' WF検査指示（L1)
Public WFINDL2 As String            ' WF検査指示（L2)
Public WFINDL3 As String            ' WF検査指示（L3)
Public WFINDL4 As String            ' WF検査指示（L4)
Public WFINDDS As String            ' WF検査指示（DS)
Public WFINDDZ As String            ' WF検査指示（DZ)
Public WFINDSP As String            ' WF検査指示（SP)
Public WFINDDO1 As String           ' WF検査指示（DO1)
Public WFINDDO2 As String           ' WF検査指示（DO2)
Public WFINDDO3 As String           ' WF検査指示（DO3)
Public WFINDAOI As String           ' WF検査指示 (AO)　追加　03/12/05 ooba
Public WFINDGD As String            ' WF検査指示 (GD)　追加　05/01/17 ooba
Public WFRESRS As String            ' WF検査実績（Rs)
Public WFRESOI As String            ' WF検査実績（Oi)
Public WFRESB1 As String            ' WF検査実績（B1)
Public WFRESB2 As String            ' WF検査実績（B2）
Public WFRESB3 As String            ' WF検査実績（B3)
Public WFRESL1 As String            ' WF検査実績（L1)
Public WFRESL2 As String            ' WF検査実績（L2)
Public WFRESL3 As String            ' WF検査実績（L3)
Public WFRESL4 As String            ' WF検査実績（L4)
Public WFRESDS As String            ' WF検査実績（DS)
Public WFRESDZ As String            ' WF検査実績（DZ)
Public WFRESSP As String            ' WF検査実績（SP)
Public WFRESDO1 As String           ' WF検査実績（DO1)
Public WFRESDO2 As String           ' WF検査実績（DO2)
Public WFRESDO3 As String           ' WF検査実績（DO3)
Public WFRESAOI As String           ' WF検査実績 (AO)　追加　03/12/05 ooba
Public WFRESGD As String            ' WF検査実績 (GD)　追加　05/01/17 ooba
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
Public EPINDB1 As String            ' WF検査指示 (BMD1E)
Public EPRESB1 As String            ' WF検査実績 (BMD1E)
Public EPINDB2 As String            ' WF検査指示 (BMD2E)
Public EPRESB2 As String            ' WF検査実績 (BMD2E)
Public EPINDB3 As String            ' WF検査指示 (BMD3E)
Public EPRESB3 As String            ' WF検査実績 (BMD3E)
Public EPINDL1 As String            ' WF検査指示 (OSF1E)
Public EPRESL1 As String            ' WF検査実績 (OSF1E)
Public EPINDL2 As String            ' WF検査指示 (OSF2E)
Public EPRESL2 As String            ' WF検査実績 (OSF2E)
Public EPINDL3 As String            ' WF検査指示 (OSF3E)
Public EPRESL3 As String            ' WF検査実績 (OSF3E)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-


'Public REGDATE As Date              ' 登録日付
'Public UPDDATE As Date              ' 更新日付
'Public SENDFLAG As String           ' 送信フラグ
'Public SENDDATE As Date             ' 送信日付


'概要      :自己複製
'説明      :自オブジェクトと同内容のオブジェクトを生成する
'履歴      :2001/06/23 作成  野村
Public Function Clone() As c_cmzcWfSmp
Attribute Clone.VB_Description = "; 自オブジェクトと同内容のオブジェクトを生成します; "
Dim CloneTo As New c_cmzcWfSmp   '複製先

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcWfSmp.cls -- Function Clone"

    '' 新規オブジェクトに単純公開プロパティを複写する
    With CloneTo
        .CRYNUM = CRYNUM            ' 結晶番号
        .INGOTPOS = INGOTPOS        ' 結晶内位置
        .SMPKBN = SMPKBN            ' サンプル区分
        .SMPLID = SMPLID            ' サンプルID
        .hinban = hinban            ' 品番
        .REVNUM = REVNUM            ' 製品番号改訂番号
        .factory = factory          ' 工場
        .opecond = opecond          ' 操業条件
        .KTKBN = KTKBN              ' 確定区分
        .WFINDRS = WFINDRS          ' WF検査指示（Rs)
        .WFINDOI = WFINDOI          ' WF検査指示（Oi)
        .WFINDB1 = WFINDB1          ' WF検査指示（B1)
        .WFINDB2 = WFINDB2          ' WF検査指示（B2）
        .WFINDB3 = WFINDB3          ' WF検査指示（B3)
        .WFINDL1 = WFINDL1          ' WF検査指示（L1)
        .WFINDL2 = WFINDL2          ' WF検査指示（L2)
        .WFINDL3 = WFINDL3          ' WF検査指示（L3)
        .WFINDL4 = WFINDL4          ' WF検査指示（L4)
        .WFINDDS = WFINDDS          ' WF検査指示（DS)
        .WFINDDZ = WFINDDZ          ' WF検査指示（DZ)
        .WFINDSP = WFINDSP          ' WF検査指示（SP)
        .WFINDDO1 = WFINDDO1        ' WF検査指示（DO1)
        .WFINDDO2 = WFINDDO2        ' WF検査指示（DO2)
        .WFINDDO3 = WFINDDO3        ' WF検査指示（DO3)
        .WFINDAOI = WFINDAOI        ' WF検査指示 (AO)　追加　03/12/05 ooba
        .WFINDGD = WFINDGD          ' WF検査指示 (GD)　追加　05/01/18 ooba
        .WFRESRS = WFRESRS          ' WF検査実績（Rs)
        .WFRESOI = WFRESOI          ' WF検査実績（Oi)
        .WFRESB1 = WFRESB1          ' WF検査実績（B1)
        .WFRESB2 = WFRESB2          ' WF検査実績（B2）
        .WFRESB3 = WFRESB3          ' WF検査実績（B3)
        .WFRESL1 = WFRESL1          ' WF検査実績（L1)
        .WFRESL2 = WFRESL2          ' WF検査実績（L2)
        .WFRESL3 = WFRESL3          ' WF検査実績（L3)
        .WFRESL4 = WFRESL4          ' WF検査実績（L4)
        .WFRESDS = WFRESDS          ' WF検査実績（DS)
        .WFRESDZ = WFRESDZ          ' WF検査実績（DZ)
        .WFRESSP = WFRESSP          ' WF検査実績（SP)
        .WFRESDO1 = WFRESDO1        ' WF検査実績（DO1)
        .WFRESDO2 = WFRESDO2        ' WF検査実績（DO2)
        .WFRESDO3 = WFRESDO3        ' WF検査実績（DO3)
        .WFRESAOI = WFRESAOI        ' WF検査実績 (AO)　追加　03/12/05 ooba
        .WFRESGD = WFRESGD          ' WF検査実績 (GD)　追加　05/01/18 ooba
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
        .EPINDB1 = EPINDB1          ' WF検査指示 (BMD1E)
        .EPRESB1 = EPRESB1          ' WF検査実績 (BMD1E)
        .EPINDB2 = EPINDB2          ' WF検査指示 (BMD2E)
        .EPRESB2 = EPRESB2          ' WF検査実績 (BMD2E)
        .EPINDB3 = EPINDB3          ' WF検査指示 (BMD3E)
        .EPRESB3 = EPRESB3          ' WF検査実績 (BMD3E)
        .EPINDL1 = EPINDL1          ' WF検査指示 (OSF1E)
        .EPRESL1 = EPRESL1          ' WF検査実績 (OSF1E)
        .EPINDL2 = EPINDL2          ' WF検査指示 (OSF2E)
        .EPRESL2 = EPRESL2          ' WF検査実績 (OSF2E)
        .EPINDL3 = EPINDL3          ' WF検査指示 (OSF3E)
        .EPRESL3 = EPRESL3          ' WF検査実績 (OSF3E)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
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

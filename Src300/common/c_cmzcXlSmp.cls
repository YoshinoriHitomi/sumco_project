VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_cmzcXlSmp"
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
' 定義内容: cmzcXlSmp -- 結晶サンプル管理
' 参照    : 060207_結晶管理
'================================================

' 単純公開
Public CRYNUM As String             ' 結晶番号
Public INGOTPOS As Integer          ' 結晶内位置
Public SMPKBN As String             ' サンプル区分
Public SMPLNO As Long               ' サンプルNo        Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
Public hinban As String             ' 品番
Public REVNUM As Integer            ' 製品番号改訂番号
Public factory As String            ' 工場
Public opecond As String            ' 操業条件
Public KTKBN As String              ' 確定区分
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
Public CRYRESRS As String           ' 結晶検査実績（Rs)
Public CRYRESOI As String           ' 結晶検査実績（Oi)
Public CRYRESB1 As String           ' 結晶検査実績（B1)
Public CRYRESB2 As String           ' 結晶検査実績（B2）
Public CRYRESB3 As String           ' 結晶検査実績（B3)
Public CRYRESL1 As String           ' 結晶検査実績（L1)
Public CRYRESL2 As String           ' 結晶検査実績（L2)
Public CRYRESL3 As String           ' 結晶検査実績（L3)
Public CRYRESL4 As String           ' 結晶検査実績（L4)
Public CRYRESCS As String           ' 結晶検査実績（Cs)
Public CRYRESGD As String           ' 結晶検査実績（GD)
Public CRYREST As String            ' 結晶検査実績（T)
Public CRYRESEP As String           ' 結晶検査実績（EPD)
Public SMPLNUM As Integer           ' サンプル枚数
Public SMPLPAT As String            ' サンプルパターン
'Public REGDATE As Date              ' 登録日付
'Public UPDDATE As Date              ' 更新日付
'Public SENDFLAG As String           ' 送信フラグ
'Public SENDDATE As Date             ' 送信日付


'概要      :自己複製
'説明      :自オブジェクトと同内容のオブジェクトを生成する
'履歴      :2001/06/23 作成  野村
Public Function Clone() As c_cmzcXlSmp
Attribute Clone.VB_Description = "; 自オブジェクトと同内容のオブジェクトを生成します; "
Dim CloneTo As New c_cmzcXlSmp   '複製先

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXlSmp.cls -- Function Clone"

    '' 新規オブジェクトに単純公開プロパティを複写する
    With CloneTo
        .CRYNUM = CRYNUM            ' 結晶番号
        .INGOTPOS = INGOTPOS        ' 結晶内位置
        .SMPKBN = SMPKBN            ' サンプル区分
        .SMPLNO = SMPLNO            ' サンプルNo
        .hinban = hinban            ' 品番
        .REVNUM = REVNUM            ' 製品番号改訂番号
        .factory = factory          ' 工場
        .opecond = opecond          ' 操業条件
        .KTKBN = KTKBN              ' 確定区分
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
        .CRYRESRS = CRYRESRS        ' 結晶検査実績（Rs)
        .CRYRESOI = CRYRESOI        ' 結晶検査実績（Oi)
        .CRYRESB1 = CRYRESB1        ' 結晶検査実績（B1)
        .CRYRESB2 = CRYRESB2        ' 結晶検査実績（B2）
        .CRYRESB3 = CRYRESB3        ' 結晶検査実績（B3)
        .CRYRESL1 = CRYRESL1        ' 結晶検査実績（L1)
        .CRYRESL2 = CRYRESL2        ' 結晶検査実績（L2)
        .CRYRESL3 = CRYRESL3        ' 結晶検査実績（L3)
        .CRYRESL4 = CRYRESL4        ' 結晶検査実績（L4)
        .CRYRESCS = CRYRESCS        ' 結晶検査実績（Cs)
        .CRYRESGD = CRYRESGD        ' 結晶検査実績（GD)
        .CRYREST = CRYREST          ' 結晶検査実績（T)
        .CRYRESEP = CRYRESEP        ' 結晶検査実績（EPD)
        .SMPLNUM = SMPLNUM          ' サンプル枚数
        .SMPLPAT = SMPLPAT          ' サンプルパターン
'        .REGDATE = REGDATE          ' 登録日付
'        .UPDDATE = UPDDATE          ' 更新日付
'        .SENDFLAG = SENDFLAG        ' 送信フラグ
'        .SENDDATE = SENDDATE        ' 送信日付
    End With

    '' オブジェクトを返す
    Set Clone = CloneTo

proc_exit:
    '終了
    Set CloneTo = Nothing
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Function

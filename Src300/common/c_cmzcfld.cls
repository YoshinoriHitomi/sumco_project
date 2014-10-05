VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_cmzcfld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "フィールドデータの内、必要な情報を保持する"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarFieldName As String 'ﾛｰｶﾙ ｺﾋﾟｰ
Private mvarValue As Variant 'ﾛｰｶﾙ ｺﾋﾟｰ
Private mvarFieldType As Long 'ﾛｰｶﾙ ｺﾋﾟｰ
Public FldSize As Integer       '全桁数
Public FldSizeN As Integer       '全桁数(Number)
Public FldScale As Integer  '小数点ありのときの小数点下桁数

'概要      :FieldTypeプロパティ(Let)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                         ,説明
'          :vData         ,I  ,OracleInProcServer.vbTypes ,フィールド型
'説明      :
'履歴      :2001/06/08 作成  野村
Public Property Let FieldType(ByVal vData As OracleInProcServer.vbTypes)
Attribute FieldType.VB_Description = "フィールドの型"
'ﾌﾟﾛﾊﾟﾃｨに値を代入するときに、代入式の左辺で使用します。
'Syntax: X.FieldType = 5

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcfld.cls -- Property Let FieldType"

    mvarFieldType = vData

proc_exit:
    '終了
    gErr.Pop
    Exit Property

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Property


'概要      :FieldTypeプロパティ(Get)
'説明      :
'履歴      :2001/06/08 作成  野村
Public Property Get FieldType() As OracleInProcServer.vbTypes
'ﾌﾟﾛﾊﾟﾃｨの値を取得するときに、代入式の右辺で使用します。
'Syntax: Debug.Print X.FieldType

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcfld.cls -- Property Get FieldType"

    FieldType = mvarFieldType

proc_exit:
    '終了
    gErr.Pop
    Exit Property

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Property



'概要      :Valueプロパティ(Let)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :vData         ,   ,Variant   ,
'説明      :
'履歴      :2001/06/08 作成  野村
Public Property Let Value(ByVal vData As Variant)
'ﾌﾟﾛﾊﾟﾃｨに値を代入するときに、代入式の左辺で使用します。
'Syntax: X.Value = 5

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcfld.cls -- Property Let Value"

    mvarValue = vData

proc_exit:
    '終了
    gErr.Pop
    Exit Property

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Property


'概要      :FieldTypeプロパティ(Get)
'説明      :
'履歴      :2001/06/08 作成  野村
Public Property Get Value() As Variant
Attribute Value.VB_Description = "フィールドの値"
Attribute Value.VB_UserMemId = 0
'ﾌﾟﾛﾊﾟﾃｨの値を取得するときに、代入式の右辺で使用します。
'Syntax: Debug.Print X.Value

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcfld.cls -- Property Get Value"

    If IsObject(mvarValue) Then
        Set Value = mvarValue
    Else
        Value = mvarValue
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Property

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Property



'概要      :FieldNameプロパティ(Let)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :vData         ,   ,String    ,
'説明      :
'履歴      :2001/06/08 作成  野村
Public Property Let FieldName(ByVal vData As String)
Attribute FieldName.VB_Description = "フィールド名"
'ﾌﾟﾛﾊﾟﾃｨに値を代入するときに、代入式の左辺で使用します。
'Syntax: X.FieldName = 5

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcfld.cls -- Property Let FieldName"

    mvarFieldName = Trim$(vData)

proc_exit:
    '終了
    gErr.Pop
    Exit Property

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Property


'概要      :FieldNameプロパティ(Get)
'説明      :
'履歴      :2001/06/08 作成  野村
Public Property Get FieldName() As String
'ﾌﾟﾛﾊﾟﾃｨの値を取得するときに、代入式の右辺で使用します。
'Syntax: Debug.Print X.FieldName

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcfld.cls -- Property Get FieldName"

    FieldName = mvarFieldName

proc_exit:
    '終了
    gErr.Pop
    Exit Property

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Property





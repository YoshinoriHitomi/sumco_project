VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_cmzcflds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "フィールドコレクション"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"c_cmzcfld"
Attribute VB_Ext_KEY = "Member0" ,"c_cmzcfld"
Option Explicit

'ｺﾚｸｼｮﾝを保持するﾛｰｶﾙ変数
Private mCol As Collection


'概要      :フィールドが登録済ならその値、未登録なら既定値を返す
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :FieldName     ,I  ,String    ,フィールド名
'          :DefaultValue  ,I  ,Variant   ,既定値
'説明      :
'履歴      :2001/06/13 作成  野村
Public Function GetValueOrDefault(ByVal FieldName As String, ByVal DefaultValue As Variant) As Variant
Attribute GetValueOrDefault.VB_Description = "フィールドが登録済ならその値、未登録なら既定値を返す"

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcflds.cls -- Function GetValueOrDefault"

    If Exist(FieldName) Then
        GetValueOrDefault = Item(FieldName).Value
    Else
        GetValueOrDefault = DefaultValue
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function



'概要      :Addメソッド
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :FieldName     ,   ,String    ,
'          :Value         ,   ,Variant   ,
'          :FieldType     ,   ,Long      ,
'          :戻り値        ,O  ,c_cmzcfld ,
'説明      :
'履歴      :2001/06/08 作成  野村
'Public Function Add(FieldName As String, Value As Variant, FieldType As Long) As c_cmzcfld
Public Function Add(FieldName As String, Value As Variant, FieldType As Long, Optional FldSize% = -1, Optional FldScale% = -1, Optional FldSizeN% = -1) As c_cmzcfld
    '新規ｵﾌﾞｼﾞｪｸﾄを作成します。
    Dim objNewMember As c_cmzcfld

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcflds.cls -- Function Add"

    Set objNewMember = New c_cmzcfld


    'ﾒｿｯﾄﾞに渡すﾌﾟﾛﾊﾟﾃｨを設定します。
    objNewMember.FieldName = FieldName
    objNewMember.Value = Value
    objNewMember.FieldType = FieldType
    objNewMember.FldSize = FldSize
    objNewMember.FldScale = FldScale
    objNewMember.FldSizeN = FldSizeN
    If Exist(FieldName) Then
        mCol.Remove FieldName
    End If
    mCol.Add objNewMember, FieldName


    '作成されたｵﾌﾞｼﾞｪｸﾄを返します。
    Set Add = objNewMember
    Set objNewMember = Nothing



proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :Itemプロパティ(Get)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :vntIndexKey   ,   ,Variant   ,
'          :戻り値        ,O  ,c_cmzcfld ,
'説明      :
'履歴      :2001/06/08 作成  野村
Public Property Get Item(vntIndexKey As Variant) As c_cmzcfld
Attribute Item.VB_UserMemId = 0
    'ｺﾚｸｼｮﾝの要素を参照するときに使用します。
    'vntIndexKey は ｲﾝﾃﾞｯｸｽまたはｷｰのどちらかを
    '保持するために Variant で宣言されています。
    '構文: Set foo = x.Item(xyz) または Set foo = x.Item(5)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcflds.cls -- Property Get Item"

  If Exist(vntIndexKey) Then
      Set Item = mCol(vntIndexKey)
  Else
      Set Item = Nothing
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



'概要      :指定のフィールドが登録済であるかを返す
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :vntIndexKey   ,   ,Variant   ,
'          :戻り値        ,O  ,Boolean   ,
'説明      :
'履歴      :
Public Function Exist(ByVal vntIndexKey As Variant) As Boolean
Dim itm As c_cmzcfld

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcflds.cls -- Function Exist"

    On Error Resume Next
    Err.Clear
    Set itm = mCol(vntIndexKey)
    If Err = 5 Then
      Exist = False
    Else
      Exist = True
    End If
    On Error GoTo 0

proc_exit:
    '終了
    Set itm = Nothing
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function



'概要      :Countプロパティ(Get)
'説明      :
'履歴      :2001/06/08 作成  野村
Public Property Get COUNT() As Long
    'ｺﾚｸｼｮﾝの要素数を取得するときに使用します。
    '構文: Debug.Print x.Count

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcflds.cls -- Property Get Count"

    COUNT = mCol.COUNT

proc_exit:
    '終了
    gErr.Pop
    Exit Property

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Property


'概要      :Removeメソッド
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :vntIndexKey   ,   ,Variant   ,
'説明      :
'履歴      :2001/06/08 作成  野村
Public Sub Remove(vntIndexKey As Variant)
    'ｺﾚｸｼｮﾝから要素を削除するときに使用します。
    'vntIndexKey は ｲﾝﾃﾞｯｸｽまたはｷｰのどちらかを
    '保持するために Variant で宣言されています。
    '構文: x.Remove(xyz)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcflds.cls -- Sub Remove"

    mCol.Remove vntIndexKey

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub


'概要      :NewEnumプロパティ(Get)
'説明      :For...Each 構文に対応
'履歴      :2001/06/08 作成  野村
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'このﾌﾟﾛﾊﾟﾃｨは、For...Each 構文を使用して
    'ｺﾚｸｼｮﾝを列挙できるようにします。

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcflds.cls -- Property Get NewEnum"

    Set NewEnum = mCol.[_NewEnum]

proc_exit:
    '終了
    gErr.Pop
    Exit Property

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Property


'概要      :Initialize時処理
'説明      :
'履歴      :2001/06/08 作成  野村
Private Sub Class_Initialize()
    'このｸﾗｽが作成されたときに、ｺﾚｸｼｮﾝを作成します。

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcflds.cls -- Sub Class_Initialize"

    Set mCol = New Collection

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub


'概要      :Terminate時処理
'説明      :
'履歴      :2001/06/08 作成  野村
Private Sub Class_Terminate()
    'このｸﾗｽが終了するときに、ｺﾚｸｼｮﾝを破棄します。

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcflds.cls -- Sub Class_Terminate"

    Set mCol = Nothing

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub


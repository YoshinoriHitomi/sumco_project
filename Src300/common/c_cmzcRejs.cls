VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_cmzcRejs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"c_cmzcRej"
Attribute VB_Ext_KEY = "Member0" ,"c_cmzcRej"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'                                     2001/08/18
'================================================
' 欠落情報保持コレクション
' 定義内容: 欠落情報
'================================================

'ｺﾚｸｼｮﾝを保持するﾛｰｶﾙ変数
Private mCol As Collection

Public Function Add(Rej As c_cmzcRej) As c_cmzcRej
Dim i As Long

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcSxls.cls -- Function Add"

    ''全数欠落は他を消して登録
    If Trim$(Rej.ALLSCRAP) = "Y" Then
        For i = mCol.COUNT To 1 Step -1
            If Left$(mCol(i).LOTID, 12) = Rej.LOTID Then
                mCol.Remove i
            End If
        Next
        mCol.Add Rej, Rej.LOTID & "-000"
        GoTo proc_exit
    End If
    
    ''そこが欠落開始位置でなければ登録する
    If Exist(Rej.LOTID, Rej.LENFROM) Then
        GoTo proc_exit
    Else
        mCol.Add Rej, Rej.LOTID & "-" & Format$(Rej.LENFROM, "000")
    End If

    '作成されたｵﾌﾞｼﾞｪｸﾄを返します。
    Set Add = Rej

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then
        gErr.DisplayMsgOnError = False
        gErr.HandleError
        Debug.Print gErr.ProcName & " : Error " & gErr.ErrorNumber & " " & gErr.ErrorDescription
        gErr.DisplayMsgOnError = True
    End If
    Resume proc_exit
End Function

Public Property Get Item(vntIndexKey As Variant) As c_cmzcRej
Attribute Item.VB_UserMemId = 0
    'ｺﾚｸｼｮﾝの要素を参照するときに使用します。
'vntIndexKey は ｲﾝﾃﾞｯｸｽまたはｷｰのどちらかを
    '保持するために Variant で宣言されています。
    '構文: Set foo = x.Item(xyz) または Set foo = x.Item(5)

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcRejs.cls -- Property Get Item"

  Set Item = mCol(vntIndexKey)

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Property

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Property



Public Property Get COUNT() As Long
    'ｺﾚｸｼｮﾝの要素数を取得するときに使用します。
    '構文: Debug.Print x.Count

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcRejs.cls -- Property Get Count"

    COUNT = mCol.COUNT

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Property

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Property


Public Sub Remove(vntIndexKey As Variant)
    'ｺﾚｸｼｮﾝから要素を削除するときに使用します。
    'vntIndexKey は ｲﾝﾃﾞｯｸｽまたはｷｰのどちらかを
    '保持するために Variant で宣言されています。
    '構文: x.Remove(xyz)

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcRejs.cls -- Sub Remove"

    ''アイテムを削除する
    mCol.Remove vntIndexKey

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'このﾌﾟﾛﾊﾟﾃｨは、For...Each 構文を使用して
    'ｺﾚｸｼｮﾝを列挙できるようにします。

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcRejs.cls -- Property Get NewEnum"

    Set NewEnum = mCol.[_NewEnum]

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Property

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Property


Private Sub Class_Initialize()
    'このｸﾗｽが作成されたときに、ｺﾚｸｼｮﾝを作成します。

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcRejs.cls -- Sub Class_Initialize"

    Set mCol = New Collection

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Sub


Private Sub Class_Terminate()
    'このｸﾗｽが終了するときに、ｺﾚｸｼｮﾝを破棄します。

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcRejs.cls -- Sub Class_Terminate"

    Set mCol = Nothing

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Sub



'概要      :指定位置に欠落開始位置があるかを調べる
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,Boolean   ,True:ある  False:ない
'説明      :
'履歴      :2001/06/23 作成  野村
Public Function Exist(ByVal LOTID$, ByVal pos%) As Boolean
Dim Rej As c_cmzcRej

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcRejs.cls -- Function Exist"

    On Error Resume Next
    Set Rej = mCol(LOTID & "-" & Format$(pos, "000"))
    If Err.Number = 5 Then
        Exist = False
    Else
        Exist = True
    End If
    On Error GoTo 0

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Function


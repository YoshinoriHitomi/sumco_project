VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_cmzcCuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "切断指示コレクション"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"c_cmzcCut"
Attribute VB_Ext_KEY = "Member0" ,"c_cmzcCut"
Option Explicit
'                                     2001/05/17
'================================================
' TBL情報保持コレクション
' 定義内容: 切断指示
' 参照    : 060207_結晶管理
'================================================

'ｺﾚｸｼｮﾝを保持するﾛｰｶﾙ変数
Private mCol As Collection

'概要      :直胴長/フリー長など、指定の位置までに調整する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :INGOTPOS      ,I  ,Integer   ,指定位置
'説明      :LENGTH超過を切り詰める / INGOTPOS超過を削除する
'           INGOTPOS<0 のメンバを捨てる
'           LENGTH<0 のメンバを捨てる
'履歴      :2001/07/09 作成  野村
'           2001/07/28 修正  野村
Public Sub LimitByIngotPos(ByVal INGOTPOS As Integer)
Dim cnt As Integer
Dim i As Integer
Dim del As Boolean

    cnt = mCol.COUNT
    For i = cnt To 1 Step -1
        del = False
        With mCol(i)
            If .INGOTPOS > INGOTPOS Then
                del = True
            ElseIf .INGOTPOS + .LENGTH > INGOTPOS Then
                .LENGTH = INGOTPOS - .INGOTPOS
            End If
            If (.INGOTPOS < 0) Or (.LENGTH < 0) Then
                del = True
            End If
            If del = True Then
                mCol.Remove i
            End If
        End With
    Next
End Sub



'概要      :指定位置の上側にある切断指示位置を検索する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,Integer   ,切断指示位置(見つからないときは -9999)
'説明      :指定位置そのものとなる品番設計区切位置は、下側と見なす
'履歴      :2001/06/23 作成  野村
Public Function UpperPos(ByVal pos As Integer) As Integer
Dim i As Integer
Dim cnt As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcCuts.cls -- Function UpperPos"

    ''指定位置の上側にある切断指示位置を検索する
    pos = pos - 1
    UpperPos = -9999
    cnt = mCol.COUNT
    For i = cnt To 1 Step -1
        If mCol(i).INGOTPOS < pos Then
            UpperPos = mCol(i).INGOTPOS
            Exit For
        End If
    Next

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Function


'概要      :指定位置の下側にある切断指示位置を検索する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,Integer   ,品番設計区切位置(見つからないときは 9999)
'説明      :指定位置そのものとなる品番設計区切位置は、下側と見なす
'履歴      :2001/05/17 作成  野村
Public Function LowerPos(ByVal pos As Integer) As Integer
Dim i As Integer
Dim cnt As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcCuts.cls -- Function LowerPos"

    ''指定位置の下側にある切断指示位置を検索する
    LowerPos = 9999
    cnt = mCol.COUNT
    For i = 1 To cnt
        If mCol(i).INGOTPOS >= pos Then
            LowerPos = mCol(i).INGOTPOS
            Exit For
        End If
    Next

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Function

'概要      :指定位置の上側にある切断指示位置を検索する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,Integer   ,切断指示位置(見つからないときは -9999)
'説明      :指定位置そのものとなる切断指示位置は、下側と見なす
'履歴      :2001/06/23 作成  野村
Public Function UpperArea(ByVal pos As Integer) As Integer
Dim i As Integer
Dim cnt As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcCuts.cls -- Function UpperArea"

    ''(指定位置-1mm)を含む切断指示を検索する
    pos = pos - 1
    UpperArea = -9999
    cnt = mCol.COUNT
    For i = 1 To cnt
        If mCol(i).IsInArea(pos) Then
            UpperArea = mCol(i).INGOTPOS
            Exit For
        End If
    Next

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Function


'概要      :指定位置の下側にある切断指示位置を検索する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,Integer   ,切断指示位置(見つからないときは 9999)
'説明      :指定位置そのものとなる切断位置は、下側と見なす
'履歴      :2001/06/23 作成  野村
Public Function LowerArea(ByVal pos As Integer) As Integer
Dim i As Integer
Dim cnt As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcCuts.cls -- Function LowerArea"

    ''指定位置を含むブロックを検索する
    LowerArea = 9999
    cnt = mCol.COUNT
    For i = cnt To 1 Step -1
        If mCol(i).IsInArea(pos) Then
            LowerArea = mCol(i).INGOTPOS
            Exit For
        ElseIf mCol(i).INGOTPOS >= pos Then
            LowerArea = mCol(i).INGOTPOS
        End If
    Next

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Function


'概要      :切断指示を追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型         ,説明
'          :CutLine       ,I  ,c_cmzcCut ,追加する切断指示
'          :戻り値        ,O  ,c_cmzcCut ,追加されたオブジェクト
'説明      :インゴット内位置順に並ぶように追加する
'履歴      :2001/05/17 作成  野村
Public Function Add(Cut As c_cmzcCut) As c_cmzcCut
Dim pos As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcCuts.cls -- Function Add"

    ''切断指示は同位置に２レコード存在するケースがある
    ''このとき、LENGTH=0のレコードは下品番を持たないので無視する
    If Exist(Cut.INGOTPOS) Then
        If Cut.LENGTH = 0 Then
            Set Add = Nothing
            GoTo proc_exit
        Else
            '既にあるレコードがいらない(LENGTH=0)はずなので削除する
            Remove CStr(Cut.INGOTPOS)
        End If
    End If
    
    ''指定位置を含む切断指示の長さを調整する
    pos = LowerArea(Cut.INGOTPOS)
    If (pos <= Cut.INGOTPOS) Then
        Cut.LENGTH = pos + Item(CStr(pos)).LENGTH - Cut.INGOTPOS
        Item(CStr(pos)).LENGTH = Cut.INGOTPOS - pos
    End If
        
    ''次切断指示にかかっていたら、長さを調整する
    pos = LowerArea(Cut.INGOTPOS)
    If (pos < 9999) And (Cut.INGOTPOS < pos) Then
        If (Cut.LENGTH > pos - Cut.INGOTPOS) Then
            Cut.LENGTH = pos - Cut.INGOTPOS
        End If
    End If
    
    '' インゴット内位置をキーとして、与えられたクラスを追加する
    If (pos = 9999) Then
        mCol.Add Cut, CStr(Cut.INGOTPOS)
    Else
        mCol.Add Cut, CStr(Cut.INGOTPOS), CStr(pos)
    End If

    ''作成されたｵﾌﾞｼﾞｪｸﾄを返します。
    Set Add = Cut

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


'概要      :アイテムを得る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :vntIndexKey   ,I  ,Variant   ,要素番号 または インゴット内位置(文字列)
'          :戻り値        ,O  ,c_cmzcBlk,アイテム
'説明      :
'履歴      :2001/05/17 作成  野村
Public Property Get Item(vntIndexKey As Variant) As c_cmzcCut
Attribute Item.VB_UserMemId = 0
    'ｺﾚｸｼｮﾝの要素を参照するときに使用します。
    'vntIndexKey は ｲﾝﾃﾞｯｸｽまたはｷｰのどちらかを
    '保持するために Variant で宣言されています。
    '構文: Set foo = x.Item(xyz) または Set foo = x.Item(5)

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcCuts.cls -- Property Get Item"

    On Error Resume Next
    Set Item = mCol(vntIndexKey)
    If Err Then
        If VarType(vntIndexKey) = vbString Then
            For Each Item In mCol
                If Item.INGOTPOS = vntIndexKey Then
                    Exit For
                End If
            Next
        End If
    End If

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Property

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Property



'概要      :要素数を求める
'説明      :
'履歴      :2001/05/17 作成  野村
Public Property Get COUNT() As Long
    'ｺﾚｸｼｮﾝの要素数を取得するときに使用します。
    '構文: Debug.Print x.Count

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcCuts.cls -- Property Get Count"

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


'概要      :アイテムを削除する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :vntIndexKey   ,I  ,Variant   ,要素番号 または インゴット内位置(文字列)
'説明      :
'履歴      :2001/05/17 作成  野村
Public Sub Remove(vntIndexKey As Variant)
Dim pos As Integer
Dim upPos As Integer
    'ｺﾚｸｼｮﾝから要素を削除するときに使用します。
    'vntIndexKey は ｲﾝﾃﾞｯｸｽまたはｷｰのどちらかを
    '保持するために Variant で宣言されています。
    '構文: x.Remove(xyz)

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcCuts.cls -- Sub Remove"

    ''前のアイテムの長さを延ばす
    If VarType(vntIndexKey) = vbInteger Then
        pos = Item(vntIndexKey).INGOTPOS
    Else
        pos = Val(vntIndexKey)
    End If
    upPos = UpperPos(pos)
    If upPos >= 0 Then
        Item(CStr(upPos)).LENGTH = Item(CStr(upPos)).LENGTH + Item(CStr(pos)).LENGTH
    End If
    
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


'概要      :For...Each で次に得られる要素番号を得る
'説明      :
'履歴      :2001/05/17 作成  野村
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'このﾌﾟﾛﾊﾟﾃｨは、For...Each 構文を使用して
    'ｺﾚｸｼｮﾝを列挙できるようにします。

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcCuts.cls -- Property Get NewEnum"

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


'概要      :クラスのInitialize時処理
'説明      :
'履歴      :2001/05/17 作成  野村
Private Sub Class_Initialize()
    'このｸﾗｽが作成されたときに、ｺﾚｸｼｮﾝを作成します。

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcCuts.cls -- Sub Class_Initialize"

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


'概要      :クラスのTerminate時処理
'説明      :
'履歴      :2001/05/17 作成  野村
Private Sub Class_Terminate()
    'このｸﾗｽが終了するときに、ｺﾚｸｼｮﾝを破棄します。

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcCuts.cls -- Sub Class_Terminate"

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


'概要      :指定位置に切断位置があるかを調べる
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,Boolean   ,True:ある  False:ない
'説明      :
'履歴      :2001/06/23 作成  野村
Public Function Exist(ByVal pos%) As Boolean
Dim Cut As c_cmzcCut


    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcCuts.cls -- Function Exist"

    On Error Resume Next
    Set Cut = Item(CStr(pos))
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

'概要      :指定範囲内に切断指示があるかを調べる
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,範囲開始位置
'          :length        ,I  ,Integer   ,範囲長さ
'          :戻り値        ,O  ,Boolean   ,True:ある  False:ない
'説明      :範囲開始位置にある場合も、範囲内とみなす。
'履歴      :2001/06/23 作成  野村
Public Function ExistInArea(ByVal pos%, ByVal LENGTH%) As Boolean
Dim Cut As c_cmzcCut
Dim pos2 As Integer


    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcCuts.cls -- Function ExistInArea"

    pos2 = pos + LENGTH
    ExistInArea = False
    For Each Cut In mCol
        If (pos <= Cut.INGOTPOS) And (Cut.INGOTPOS < pos2) Then
            ExistInArea = True
            Exit For
        End If
    Next

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Function

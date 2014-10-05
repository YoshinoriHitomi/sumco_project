VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_cmzcHins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "品番区切位置コレクション"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"c_cmzcHin"
Attribute VB_Ext_KEY = "Member0" ,"c_cmzcHin"
Option Explicit
'                                     2001/05/17
'================================================
' TBL情報保持コレクション
' 定義内容: 品番管理
' 参照    : 060207_結晶管理
'================================================

'ｺﾚｸｼｮﾝを保持するﾛｰｶﾙ変数
Private mCol As Collection

'概要      :直胴長/フリー長など、指定の位置までに調整する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :INGOTPOS      ,I  ,Integer   ,指定位置
'説明      :LENGTH超過を切り詰める / INGOTPOS超過を削除する
'           INGOTPOS<0 のメンバを捨てる
'           LENGTH<=0 のメンバを捨てる
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
            If (.INGOTPOS < 0) Or (.LENGTH <= 0) Then
                del = True
            End If
            If del = True Then
                mCol.Remove i
            End If
        End With
    Next
End Sub


'概要      :最TOP切断位置～最TAIL切断位置など、指定の範囲までに調整する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :INGOTPOS      ,I  ,Integer   ,指定位置
'説明      :
'履歴      :2001/10/04 作成  野村
Public Sub LimitByIngotArea(ByVal areaFrom As Integer, ByVal areaTo As Integer)
Dim cnt As Integer
Dim i As Integer
Dim del As Boolean

    cnt = mCol.COUNT
    For i = cnt To 1 Step -1
        del = False
        With mCol(i)
            If .INGOTPOS >= areaTo Then
                del = True
            ElseIf .INGOTPOS + .LENGTH > areaTo Then
                .LENGTH = areaTo - .INGOTPOS
            End If
            If .INGOTPOS + .LENGTH <= areaFrom Then
                del = True
            ElseIf .INGOTPOS < areaFrom Then
                .LENGTH = .INGOTPOS + .LENGTH - areaFrom
                .INGOTPOS = areaFrom
            End If
            If (.LENGTH <= 0) Then
                del = True
            End If
            If del = True Then
                mCol.Remove i
            End If
        End With
    Next
End Sub


'概要      :指定位置の上側にある品番開始位置を検索する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,Integer   ,品番区切位置(見つからないときは -9999)
'説明      :指定位置そのものとなる品番設計区切位置は、下側と見なす
'履歴      :2001/06/23 作成  野村
Public Function UpperPos(ByVal pos As Integer) As Integer
Dim i As Integer
Dim cnt As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcHins.cls -- Function UpperPos"

    ''指定位置の上側にある品番開始位置を検索する
    'pos = pos - 1
    UpperPos = -9999
    cnt = mCol.COUNT
    For i = cnt To 1 Step -1
        If mCol(i).INGOTPOS < (pos - 0.1) Then
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


'概要      :指定位置の下側にある品番開始位置を検索する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,Integer   ,品番区切位置(見つからないときは 9999)
'説明      :指定位置そのものとなる品番設計区切位置は、下側と見なす
'履歴      :2001/05/17 作成  野村
Public Function LowerPos(ByVal pos As Integer) As Integer
Dim i As Integer
Dim cnt As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcHins.cls -- Function LowerPos"

    ''指定位置の下側にある品番開始位置を検索する
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

'概要      :指定位置の上側にある品番開始位置を検索する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,Integer   ,品番区切位置(見つからないときは -9999)
'説明      :指定位置そのものとなる品番区切位置は、下側と見なす
'履歴      :2001/05/17 作成  野村
Public Function UpperArea(ByVal pos As Integer) As Integer
Dim i As Integer
Dim cnt As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcHins.cls -- Function UpperArea"

    ''(指定位置-1mm)を含む品番を検索する
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


'概要      :指定位置の下側にある品番開始位置を検索する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,Integer   ,品番区切位置(見つからないときは 9999)
'説明      :指定位置そのものとなる品番区切位置は、下側と見なす
'履歴      :2001/05/17 作成  野村
Public Function LowerArea(ByVal pos As Integer) As Integer
Dim i As Integer
Dim cnt As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcHins.cls -- Function LowerArea"

    ''指定位置を含む品番を検索する
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

'概要      :品番を追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型         ,説明
'          :Hin           ,I  ,c_cmzcHin ,追加する品番
'          :戻り値        ,O  ,c_cmzcHin ,追加されたオブジェクト
'説明      :インゴット内位置順に並ぶように追加する
'履歴      :2001/05/17 作成  野村
Public Function Add(HIN As c_cmzcHin) As c_cmzcHin
Dim pos As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcHins.cls -- Function Add"

    ''指定位置を含む品番の長さを調整する
    pos = LowerArea(HIN.INGOTPOS)
    If (pos <= HIN.INGOTPOS) Then
        HIN.LENGTH = pos + Item(CStr(pos)).LENGTH - HIN.INGOTPOS
        Item(CStr(pos)).LENGTH = HIN.INGOTPOS - pos
    End If
    
    ''次品番にかかっていたら、長さを調整する
    pos = LowerArea(HIN.INGOTPOS)
    If (pos < 9999) And (HIN.INGOTPOS < pos) Then
        If (HIN.LENGTH > pos - HIN.INGOTPOS) Then
            HIN.LENGTH = pos - HIN.INGOTPOS
        End If
    End If
    
    '' インゴット内位置をキーとして、与えられたクラスを追加する
    If (pos = 9999) Then
        mCol.Add HIN, CStr(HIN.INGOTPOS)
    Else
        mCol.Add HIN, CStr(HIN.INGOTPOS), CStr(pos)
    End If

    ''作成されたｵﾌﾞｼﾞｪｸﾄを返します。
    Set Add = HIN

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
'          :戻り値        ,O  ,c_cmzcHin,アイテム
'説明      :
'履歴      :2001/05/17 作成  野村
Public Property Get Item(vntIndexKey As Variant) As c_cmzcHin
Attribute Item.VB_UserMemId = 0
    'ｺﾚｸｼｮﾝの要素を参照するときに使用します。
    'vntIndexKey は ｲﾝﾃﾞｯｸｽまたはｷｰのどちらかを
    '保持するために Variant で宣言されています。
    '構文: Set foo = x.Item(xyz) または Set foo = x.Item(5)

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcHins.cls -- Property Get Item"

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
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcHins.cls -- Property Get Count"

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
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcHins.cls -- Sub Remove"

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
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcHins.cls -- Property Get NewEnum"

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
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcHins.cls -- Sub Class_Initialize"

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
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcHins.cls -- Sub Class_Terminate"

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


'概要      :品番区切を追加する(品番構成は維持する)
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,追加位置
'説明      :
'履歴      :2001/06/26 作成  野村
Public Function AddLine(ByVal pos%) As c_cmzcHin
Dim pos1 As Integer
Dim HIN As c_cmzcHin

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcHins.cls -- Function AddLine"

    pos1 = LowerPos(pos)
    If pos1 = pos Then
        '既に品番区切があれば何もする必要はない
        Set AddLine = Item(CStr(pos))
    Else
        pos1 = UpperPos(pos)
        If pos1 >= 0 Then
            'その位置の品番として追加する
            Set HIN = Item(CStr(pos1)).Clone
            HIN.INGOTPOS = pos
            Set AddLine = Add(HIN)
            Set HIN = Nothing
        Else
            'その位置の品番がなければ、空品番のまま追加する
            Set HIN = New c_cmzcHin
            HIN.INGOTPOS = pos
            Set AddLine = Add(HIN)
            Set HIN = Nothing
        End If
    End If

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Function

Public Sub CompactLines()
Dim i As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcHins.cls -- Sub CompactLines"

    For i = mCol.COUNT To 2 Step -1
        With mCol(i - 1)
            If (.HINBAN = mCol(i).HINBAN) _
            And (.REVNUM = mCol(i).REVNUM) _
            And (.FACTORY = mCol(i).FACTORY) _
            And (.OPECOND = mCol(i).OPECOND) Then   'if(品番が前と同じ)
                Remove i
            End If
        End With
    Next

proc_exit:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume proc_exit
End Sub

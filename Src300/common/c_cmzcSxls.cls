VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_cmzcSxls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "SXL管理レコードコレクション"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"c_cmzcSxl"
Attribute VB_Ext_KEY = "Member0" ,"c_cmzcSxl"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'                                     2001/06/28
'================================================
' TBL情報保持コレクション
' 定義内容: SXL管理
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


'概要      :指定位置の上側にあるＳＸＬ開始位置を検索する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,Integer   ,ＳＸＬ開始位置(見つからないときは -9999)
'説明      :指定位置そのものとなる品番設計区切位置は、下側と見なす
'履歴      :2001/06/23 作成  野村
Public Function UpperPos(ByVal pos As Integer) As Integer
Dim i As Integer
Dim cnt As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcSxls.cls -- Function UpperPos"

    ''指定位置の上側にあるＳＸＬ開始位置を検索する
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


'概要      :指定位置の下側にあるＳＸＬ開始位置を検索する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,Integer   ,ＳＸＬ開始位置(見つからないときは 9999)
'説明      :指定位置そのものとなる品番設計区切位置は、下側と見なす
'履歴      :2001/05/17 作成  野村
Public Function LowerPos(ByVal pos As Integer) As Integer
Dim i As Integer
Dim cnt As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcSxls.cls -- Function LowerPos"

    ''指定位置の下側にあるＳＸＬ開始位置を検索する
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


'概要      :指定位置の上側にあるＳＸＬ開始位置を検索する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,Integer   ,ＳＸＬ開始位置(見つからないときは -9999)
'説明      :指定位置から始まるＳＸＬは下側とみなす
'履歴      :2001/06/23 作成  野村
Public Function UpperArea(ByVal pos As Integer) As Integer
Dim i As Integer
Dim cnt As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcSxls.cls -- Function UpperArea"

    ''(指定位置-1mm)を含むＳＸＬを検索する
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


'概要      :指定位置の下側にあるＳＸＬ開始位置を検索する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,Integer   ,ＳＸＬ開始位置(見つからないときは 9999)
'説明      :指定位置から始まるＳＸＬは下側とみなす
'履歴      :2001/06/23 作成  野村
Public Function LowerArea(ByVal pos As Integer) As Integer
Dim i As Integer
Dim cnt As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcSxls.cls -- Function LowerArea"

    ''指定位置を含むＳＸＬを検索する
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



'概要      :ＳＸＬを追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型         ,説明
'          :CutLine       ,I  ,c_cmzcBlk ,追加するＳＸＬ
'          :戻り値        ,O  ,c_cmzcBlk ,追加されたオブジェクト
'説明      :インゴット内位置順に並ぶように追加する
'履歴      :2001/05/17 作成  野村
Public Function Add(SXL As c_cmzcSxl) As c_cmzcSxl
Dim pos As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcSxls.cls -- Function Add"

    ''指定位置を含むＳＸＬの長さを調整する
    pos = LowerArea(SXL.INGOTPOS)
    If (pos <= SXL.INGOTPOS) Then
        SXL.LENGTH = pos + Item(CStr(pos)).LENGTH - SXL.INGOTPOS
        Item(CStr(pos)).LENGTH = SXL.INGOTPOS - pos
    End If
    
    ''次ＳＸＬにかかっていたら、長さを調整する
    pos = LowerArea(SXL.INGOTPOS)
    If (pos < 9999) And (SXL.INGOTPOS < pos) Then
        If (SXL.LENGTH > pos - SXL.INGOTPOS) Then
            SXL.LENGTH = pos - SXL.INGOTPOS
        End If
    End If
    
    '' インゴット内位置をキーとして、与えられたクラスを追加する
    If (pos = 9999) Then
        mCol.Add SXL, CStr(SXL.INGOTPOS)
    Else
        mCol.Add SXL, CStr(SXL.INGOTPOS), CStr(pos)
    End If

    ''作成されたｵﾌﾞｼﾞｪｸﾄを返します。
    Set Add = SXL

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
Public Property Get Item(vntIndexKey As Variant) As c_cmzcSxl
Attribute Item.VB_UserMemId = 0
    'ｺﾚｸｼｮﾝの要素を参照するときに使用します。
    'vntIndexKey は ｲﾝﾃﾞｯｸｽまたはｷｰのどちらかを
    '保持するために Variant で宣言されています。
    '構文: Set foo = x.Item(xyz) または Set foo = x.Item(5)

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcSxls.cls -- Property Get Item"

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
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcSxls.cls -- Property Get Count"

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
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcSxls.cls -- Sub Remove"

    ''前のアイテムの長さを延ばす
    If VarType(vntIndexKey) = vbInteger Then
        pos = Item(vntIndexKey).INGOTPOS
    Else
        pos = Val(vntIndexKey)
    End If
    upPos = UpperPos(pos)
    If upPos >= 0 Then
        Item(CStr(upPos)).LENGTH = mCol(CStr(upPos)).LENGTH + mCol(CStr(pos)).LENGTH
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
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcSxls.cls -- Property Get NewEnum"

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
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcSxls.cls -- Sub Class_Initialize"

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
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcSxls.cls -- Sub Class_Terminate"

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



'概要      :IDを検索し、開始位置を得る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :ID            ,I  ,String    ,ＩＤ
'          :戻り値        ,O  ,Integer   ,開始位置（見つからない場合はマイナス値）
'説明      :
'履歴      :2001/06/29 作成  野村
Public Function GetPosByID(ByVal ID As String) As Integer
Dim SXL As c_cmzcSxl

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcSxls.cls -- Function GetPosByID"

    GetPosByID = -9999
    For Each SXL In mCol
        If SXL.SXLID = ID Then
            GetPosByID = SXL.INGOTPOS
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



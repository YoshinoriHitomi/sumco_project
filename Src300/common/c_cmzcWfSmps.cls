VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_cmzcWfSmps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "WFサンプルコレクション"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"c_cmzcWfSmp"
Attribute VB_Ext_KEY = "Member0" ,"c_cmzcWfSmp"
Option Explicit
'                                     2001/05/17
'================================================
' TBL情報保持コレクション
' 定義内容: WFサンプル管理
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
            End If
            If (.INGOTPOS < 0) Then
                del = True
            End If
            If del = True Then
                mCol.Remove i
            End If
        End With
    Next
End Sub


'概要      :WFサンプル位置を追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型         ,説明
'          :CutLine       ,I  ,c_cmzcWfSmp ,追加するWFサンプル位置
'          :戻り値        ,O  ,c_cmzcWfSmp ,追加されたオブジェクト
'説明      :インゴット内位置順に並ぶように追加する
'履歴      :2001/05/17 作成  野村
Public Function Add(WfSample As c_cmzcWfSmp) As c_cmzcWfSmp
Dim i As Integer
Dim cnt As Integer
Dim pos As Integer
Dim kbn As String
Dim keyBefore As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcWfSmps.cls -- Function Add"

    '' インゴット内位置をキーとして、与えられたクラスを追加する
    pos = WfSample.INGOTPOS
    kbn = WfSample.SMPKBN
    cnt = mCol.COUNT
    For i = 1 To cnt
        If (pos < mCol(i).INGOTPOS) Then
            keyBefore = CStr(mCol(i).INGOTPOS) & mCol(i).SMPKBN
            mCol.Add WfSample, CStr(pos) & kbn, keyBefore   ''１つ後となるアイテムが見つかったら、その前に挿入する
            Set Add = WfSample
            GoTo proc_exit
        End If
    Next
    
    mCol.Add WfSample, CStr(pos) & kbn  ''見つからなかったら、最後に追加する

    ''作成されたｵﾌﾞｼﾞｪｸﾄを返します。
    Set Add = WfSample

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


'概要      :指定位置の下側にあるWFサンプル指示位置を検索する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,Integer   ,WFサンプル指示位置(見つからないときは 9999)
'説明      :指定位置そのものとなる品番設計区切位置は、下側と見なす
'履歴      :2001/06/27 作成  野村
Public Function LowerPos(ByVal pos As Integer) As Integer
Dim i As Integer
Dim cnt As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcWfSmps.cls -- Function LowerPos"

    ''指定位置の下側にあるWFサンプル指示位置を検索する
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


'概要      :アイテムを得る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :vntIndexKey   ,I  ,Variant   ,要素番号 または インゴット内位置(文字列)
'          :戻り値        ,O  ,c_cmzcWfSmp,アイテム
'説明      :
'履歴      :2001/05/17 作成  野村
Public Property Get Item(vntIndexKey As Variant) As c_cmzcWfSmp
Attribute Item.VB_UserMemId = 0
    'ｺﾚｸｼｮﾝの要素を参照するときに使用します。
    'vntIndexKey は ｲﾝﾃﾞｯｸｽまたはｷｰのどちらかを
    '保持するために Variant で宣言されています。
    '構文: Set foo = x.Item(xyz) または Set foo = x.Item(5)

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcWfSmps.cls -- Property Get Item"

    On Error Resume Next
    Set Item = mCol(vntIndexKey)
    If Err Then
        If VarType(vntIndexKey) = vbString Then
            For Each Item In mCol
                If Item.INGOTPOS & Item.SMPKBN = vntIndexKey Then
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
    If Not (gErr Is Nothing) Then
        gErr.DisplayMsgOnError = False
        gErr.HandleError
        Debug.Print gErr.ProcName & " : Error " & gErr.ErrorNumber & " " & gErr.ErrorDescription
        gErr.DisplayMsgOnError = True
    End If
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
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcWfSmps.cls -- Property Get Count"

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
    'ｺﾚｸｼｮﾝから要素を削除するときに使用します。
    'vntIndexKey は ｲﾝﾃﾞｯｸｽまたはｷｰのどちらかを
    '保持するために Variant で宣言されています。
    '構文: x.Remove(xyz)

    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcWfSmps.cls -- Sub Remove"

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
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcWfSmps.cls -- Property Get NewEnum"

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
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcWfSmps.cls -- Sub Class_Initialize"

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
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcWfSmps.cls -- Sub Class_Terminate"

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
Dim WFSMP As c_cmzcWfSmp


    'エラーハンドラの設定
    On Error GoTo proc_err
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcWfSmps.cls -- Function GetPosByID"

    GetPosByID = -9999
    For Each WFSMP In mCol
        If WFSMP.SMPLID = ID Then
            GetPosByID = WFSMP.INGOTPOS
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


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_cmzcXl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "結晶クラス"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"BadBlocks"
Attribute VB_Ext_KEY = "Member1" ,"CutLines"
Attribute VB_Ext_KEY = "Member2" ,"CutOrders"
Attribute VB_Ext_KEY = "Member3" ,"ProductLines"
Attribute VB_Ext_KEY = "Member4" ,"WfSamples"
Attribute VB_Ext_KEY = "Member5" ,"XlSamples"
Option Explicit
'                                     2001/05/17
'======================================================
' 結晶クラス
' 概要    : 結晶に関する各レコード項目を一括保持する
' 参照    : 結晶レコード保持クラス      (c_cmzcBlk.cls～c_cmzc001g.cls)
'         : 結晶レコードコレクション    (c_cmzcBlks.cls～c_cmzc001m.cls)
'======================================================

' 単純公開
Public CRYNUM As String             ' 結晶番号
Public DELCLS As String             ' 削除区分
Public KRPROCCD As String           ' 管理工程コード
Public PROCCD As String             ' 工程コード
Public LPKRPROCCD As String         ' 最終通過管理工程
Public LASTPASS As String           ' 最終通過工程
Public RPHINBAN As String           ' ねらい品番
Public RPREVNUM As Integer          ' ねらい品番製品番号改訂番号
Public RPFACT As String             ' ねらい品番工場
Public RPOPCOND As String           ' ねらい品番操業条件
Public PRODCOND As String           ' 製作条件
Public PGID As String               ' ＰＧ－ＩＤ
Public UPLENGTH As Integer          ' 引上げ長さ
Public TOPLENG As Integer           ' ＴＯＰ長さ
Public BODYLENG As Integer          ' 直胴長さ
Public BOTLENG As Integer           ' ＢＯＴ長さ
Public FREELENG As Integer          ' フリー長
Public DIAMETER As Integer          ' 直径
Public CHARGE As Double             ' チャージ量
Public SEED As String               ' シード
Public ADDDPCLS As String           ' 追加ドープ種類
Public ADDDPPOS As Integer          ' 追加ドープ位置
Public ADDDPVAL As Double           ' 追加ドープ量
'Public REGDATE As Date              ' 登録日付
'Public UPDDATE As Date              ' 更新日付
'Public SENDFLAG As String           ' 送信フラグ
'Public SENDDATE As Date             ' 送信日付
Public CurrentBlock As String   ' CurrentBlock プロパティ

'ﾌﾟﾛﾊﾟﾃｨ値を保持するためのﾛｰｶﾙ変数。
Private mvarBlks As c_cmzcBlks      '実体 Blks プロパティ
Private mvarBlkPlans As c_cmzcBlkPlans 'ﾛｰｶﾙ ｺﾋﾟｰ
Private mvarCuts As c_cmzcCuts     '実体 Cuts プロパティ
Private mvarHins As c_cmzcHins  '実体 Hins プロパティ
Private mvarHinPlans As c_cmzcHinPlans 'ﾛｰｶﾙ ｺﾋﾟｰ
Private mvarXlSmps As c_cmzcXlSmps   '実体 XlSmps プロパティ
Private mvarWfSmps As c_cmzcWfSmps   '実体 WfSmps プロパティ
Private mvarSxls As c_cmzcSxls 'ﾛｰｶﾙ ｺﾋﾟｰ
Private mvarRejs As c_cmzcRejs     '実体 Rejs プロパティ


Public Property Get Sxls() As c_cmzcSxls
Attribute Sxls.VB_Description = "ＳＸＬコレクション"
'ﾌﾟﾛﾊﾟﾃｨの値を取得するときに、代入式の右辺で使用します。
'Syntax: Debug.Print X.Sxls

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Property Get Sxls"

    Set Sxls = mvarSxls

PROC_EXIT:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Property

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Property

Public Property Get HinPlans() As c_cmzcHinPlans
Attribute HinPlans.VB_Description = "品番設計コレクション"
'ﾌﾟﾛﾊﾟﾃｨの値を取得するときに、代入式の右辺で使用します。
'Syntax: Debug.Print X.HinPlans

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Property Get HinPlans"

    Set HinPlans = mvarHinPlans

PROC_EXIT:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Property

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Property


Public Property Get BlkPlans() As c_cmzcBlkPlans
Attribute BlkPlans.VB_Description = "ブロック設計コレクション"
'ﾌﾟﾛﾊﾟﾃｨの値を取得するときに、代入式の右辺で使用します。
'Syntax: Debug.Print X.BlkPlans

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Property Get BlkPlans"

    Set BlkPlans = mvarBlkPlans

PROC_EXIT:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Property

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Property


'概要      :Rejs プロパティ (Get)
'説明      :内部コレクションの公開
'履歴      :2001/05/17 作成  野村
Public Property Get Rejs() As c_cmzcRejs
    
    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Property Get Rejs"

    '' オブジェクトがまだない場合、Newする
    If mvarRejs Is Nothing Then
        Set mvarRejs = New c_cmzcRejs
    End If

    '' 内部コレクションを返す
    Set Rejs = mvarRejs

PROC_EXIT:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Property

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Property


'概要      :WfSmps プロパティ (Get)
'説明      :内部コレクションの公開
'履歴      :2001/05/17 作成  野村
Public Property Get WfSmps() As c_cmzcWfSmps
    
    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Property Get WfSmps"

    '' オブジェクトがまだない場合、Newする
    If mvarWfSmps Is Nothing Then
        Set mvarWfSmps = New c_cmzcWfSmps
    End If

    '' 内部コレクションを返す
    Set WfSmps = mvarWfSmps

PROC_EXIT:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Property

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Property


'概要      :XlSmps プロパティ (Get)
'説明      :内部コレクションの公開
'履歴      :2001/05/17 作成  野村
Public Property Get XlSmps() As c_cmzcXlSmps

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Property Get XlSmps"

    '' オブジェクトがまだない場合、Newする
    If mvarXlSmps Is Nothing Then
        Set mvarXlSmps = New c_cmzcXlSmps
    End If

    '' 内部コレクションを返す
    Set XlSmps = mvarXlSmps

PROC_EXIT:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Property

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Property


'概要      :Hins プロパティ (Get)
'説明      :内部コレクションの公開
'履歴      :2001/05/17 作成  野村
Public Property Get Hins() As c_cmzcHins

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Property Get Hins"

    '' オブジェクトがまだない場合、Newする
    If mvarHins Is Nothing Then
        Set mvarHins = New c_cmzcHins
    End If

    '' 内部コレクションを返す
    Set Hins = mvarHins

PROC_EXIT:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Property

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Property


'概要      :Cuts プロパティ (Get)
'説明      :内部コレクションの公開
'履歴      :2001/05/17 作成  野村
Public Property Get Cuts() As c_cmzcCuts

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Property Get Cuts"

    '' オブジェクトがまだない場合、Newする
    If mvarCuts Is Nothing Then
        Set mvarCuts = New c_cmzcCuts
    End If

    '' 内部コレクションを返す
    Set Cuts = mvarCuts

PROC_EXIT:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Property

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Property


'概要      :Blks プロパティ (Get)
'説明      :内部コレクションの公開
'履歴      :2001/05/17 作成  野村
Public Property Get Blks() As c_cmzcBlks

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Property Get Blks"

    '' オブジェクトがまだない場合、Newする
    If mvarBlks Is Nothing Then
        Set mvarBlks = New c_cmzcBlks
    End If

    '' 内部コレクションを返す
    Set Blks = mvarBlks

PROC_EXIT:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Property

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Property


'概要      :クラスInitialize時処理
'説明      :
'履歴      :2001/05/17 作成  野村
Private Sub Class_Initialize()

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Sub Class_Initialize"

    ''内部データを初期化する
    Clear

PROC_EXIT:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Sub


'概要      :クラスTerminate時処理
'説明      :
'履歴      :2001/05/17 作成  野村
Private Sub Class_Terminate()
    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Sub Class_Terminate"

    ''各内部クラスを破棄する
    Set mvarRejs = Nothing
    Set mvarWfSmps = Nothing
    Set mvarXlSmps = Nothing
    Set mvarHins = Nothing
    Set mvarHinPlans = Nothing
    Set mvarCuts = Nothing
    Set mvarBlks = Nothing
    Set mvarBlkPlans = Nothing
    Set mvarSxls = Nothing

PROC_EXIT:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Sub




'概要      :指定位置の下側の品番を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,String    ,品番
'説明      :
'履歴      :2001/05/17 作成  野村
Public Function GetDownProduct(ByVal pos As Integer) As String
Attribute GetDownProduct.VB_Description = "指定位置の下側の品番を取得する"
Dim prodPos As Integer   '下側を含む品番区切位置

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Function GetDownProduct"

    ''区切位置を求める
    prodPos = mvarHins.LowerArea(pos)               '指定位置以降の最初の区切位置を求める
    
    ''区切位置の品番を返す
    If (prodPos = 9999) Then
        GetDownProduct = vbNullString               '' 区切位置がないときは、VbNullString
    Else
        GetDownProduct = mvarHins(prodPos).hinban   '' 区切位置があるときは、そこの品番
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


'概要      :指定位置の上側の品番を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,String    ,品番
'説明      :
'履歴      :2001/05/17 作成  野村
Public Function GetUpProduct(ByVal pos As Integer) As String
Attribute GetUpProduct.VB_Description = "指定位置の上側の品番を取得する"
Dim prodPos As Integer   '上側を含む品番区切位置

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Function GetUpProduct"

    ''区切位置を求める
    prodPos = mvarHins.UpperArea(pos)               '指定位置以降の最初の区切位置を求める
    
    ''区切位置の品番を返す
    If (prodPos < 0) Then
        GetUpProduct = vbNullString                 '' 区切位置がないときは、VbNullString
    Else
        GetUpProduct = mvarHins(prodPos).hinban     '' 区切位置があるときは、そこの品番
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


'概要      :指定位置の下側のブロックIDを取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,String    ,ブロックID
'説明      :
'履歴      :2001/05/17 作成  野村
Public Function GetDownBlock(ByVal pos As Integer) As String
Attribute GetDownBlock.VB_Description = "指定位置の下側のブロックIDを取得する"
Dim CutPos As Integer   '下側を含む切断位置

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Function GetDownBlock"

    ''切断位置を求める
    CutPos = mvarBlks.LowerArea(pos)               '指定位置以降の最初の切断位置を求める
    
    ''切断位置のブロックIDを返す
    If (CutPos = 9999) Then
        GetDownBlock = vbNullString                     '' ブロックがないときは、VbNullString
    Else
        GetDownBlock = mvarBlks(CStr(CutPos)).BLOCKID     '' ブロックがあるときは、そのBlockID
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


'概要      :指定位置の上のブロックIDを取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :戻り値        ,O  ,String    ,ブロックID
'説明      :
'履歴      :2001/05/17 作成  野村
Public Function GetUpBlock(ByVal pos As Integer) As String
Attribute GetUpBlock.VB_Description = "指定位置の上のブロックIDを取得する"
Dim CutPos As Integer   '上側を含む切断位置

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Function GetUpBlock"

    ''切断位置を求める
    CutPos = mvarBlks.UpperArea(pos)               '指定位置以降の最初の切断位置を求める
    
    ''切断位置のブロックIDを返す
    If (CutPos < 0) Then
        GetUpBlock = vbNullString                       '' ブロックがないときは、VbNullString
    Else
        GetUpBlock = mvarBlks(CStr(CutPos)).BLOCKID       '' ブロックがあるときは、そのBlockID
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


'概要      :自己複製
'説明      :自オブジェクトと同内容のオブジェクトを生成する
'履歴      :2001/05/17 作成  野村
Public Function Clone() As c_cmzcXl
Dim CloneTo As c_cmzcXl

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Function Clone"

    Set CloneTo = New c_cmzcXl
    CopyData CloneTo
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


'概要      :他の結晶クラスオブジェクトに内部データをコピーする
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :copyTo        ,IO ,c_cmz001a ,複写先の結晶クラスオブジェクト
'説明      :
'履歴      :2001/05/17 作成  野村
Public Sub CopyData(ByRef CopyTo As c_cmzcXl)
Attribute CopyData.VB_Description = "他の結晶クラスオブジェクトに内部データをコピーする"
Dim Blk As c_cmzcBlk            '描画対象の Blk
Dim blkPlan As c_cmzcBlkPlan    '描画対象の Blkplan
Dim Cut As c_cmzcCut            '描画対象の Cut
Dim HIN As c_cmzcHin            '描画対象の Hin
Dim hinPlan As c_cmzcHinPlan    '描画対象の Hinplan
Dim XlSample As c_cmzcXlSmp     '描画対象の XlSample
Dim WfSample As c_cmzcWfSmp     '描画対象の WfSample
Dim sxl As c_cmzcSxl            '描画対象の Sxl
Dim Rej As c_cmzcRej            '描画対象の Rej

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Sub CopyData"

    With CopyTo
        ''複写先を初期化する
        .Clear
        
        '' 各単純公開プロパティを複写する
        .CRYNUM = CRYNUM            ' 結晶番号
        .DELCLS = DELCLS            ' 削除区分
        .KRPROCCD = KRPROCCD        ' 管理工程コード
        .PROCCD = PROCCD            ' 工程コード
        .LPKRPROCCD = LPKRPROCCD    ' 最終通過管理工程
        .LASTPASS = LASTPASS        ' 最終通過工程
        .RPHINBAN = RPHINBAN        ' ねらい品番
        .RPREVNUM = RPREVNUM        ' ねらい品番製品番号改訂番号
        .RPFACT = RPFACT            ' ねらい品番工場
        .RPOPCOND = RPOPCOND        ' ねらい品番操業条件
        .PRODCOND = PRODCOND        ' 製作条件
        .PGID = PGID                ' ＰＧ－ＩＤ
        .UPLENGTH = UPLENGTH        ' 引上げ長さ
        .TOPLENG = TOPLENG          ' ＴＯＰ長さ
        .BODYLENG = BODYLENG        ' 直胴長さ
        .BOTLENG = BOTLENG          ' ＢＯＴ長さ
        .FREELENG = FREELENG        ' フリー長
        .DIAMETER = DIAMETER        ' 直径
        .CHARGE = CHARGE            ' チャージ量
        .SEED = SEED                ' シード
        .ADDDPCLS = ADDDPCLS        ' 追加ドープ種類
        .ADDDPPOS = ADDDPPOS        ' 追加ドープ位置
        .ADDDPVAL = ADDDPVAL        ' 追加ドープ量
'        .REGDATE = REGDATE          ' 登録日付
'        .UPDDATE = UPDDATE          ' 更新日付
'        .SENDFLAG = SENDFLAG        ' 送信フラグ
'        .SENDDATE = SENDDATE        ' 送信日付
        .CurrentBlock = CurrentBlock    ' 操作中のブロックID
        
        '' 各コレクションプロパティを複写する
        For Each Blk In mvarBlks            '' Blks を複写する
            .Blks.Add Blk.Clone
        Next
        For Each blkPlan In mvarBlkPlans        '' BlkPlans を複写する
            .BlkPlans.Add blkPlan.Clone
        Next
        For Each Cut In mvarCuts                '' Cuts を複写する
            .Cuts.Add Cut.Clone
        Next
        For Each HIN In mvarHins                '' Hins を複写する
            .Hins.Add HIN.Clone
        Next
        For Each hinPlan In mvarHinPlans        '' HinPlans を複写する
            .HinPlans.Add hinPlan.Clone
        Next
        For Each XlSample In mvarXlSmps         '' XlSmps を複写する
            .XlSmps.Add XlSample.Clone
        Next
        For Each WfSample In mvarWfSmps         '' WfSmps を複写する
            .WfSmps.Add WfSample.Clone
        Next
        For Each sxl In mvarSxls                '' Sxls を複写する
            .Sxls.Add sxl.Clone
        Next
        For Each Rej In mvarRejs                '' Rejs を複写する
            .Rejs.Add Rej.Clone
        Next
    End With

PROC_EXIT:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Sub


'概要      :欠落情報を設定する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :LOTID         ,I  ,String    ,ブロックID
'          :ALLSCRAP      ,I  ,String    ,全数欠落フラグ
'          :LENFROM       ,I  ,Integer   ,欠落長さ　FROM
'          :LENTO         ,I  ,Integer   ,欠落長さ　TO
'説明      :対象となるブロックが未設定の場合、描画できない
'履歴      :2001/05/17 作成  野村
'          :2001/08/18 修正  野村
Public Sub AddReject(ByVal LOTID As String, ByVal ALLSCRAP$, ByVal LENFROM As Integer, ByVal LENTO As Integer)
Attribute AddReject.VB_Description = "欠落情報を設定する"
Dim Rej As New c_cmzcRej      '欠落情報

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Sub AddRej"

    ''欠落情報クラスをコレクションに追加する
    With Rej
        .LOTID = LOTID
        .ALLSCRAP = ALLSCRAP
        .LENFROM = LENFROM
        .LENTO = LENTO
    End With
    Rejs.Add Rej

PROC_EXIT:
    '終了
    Set Rej = Nothing
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Sub


'概要      :WFサンプル位置を指定する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :[ProdID]      ,I  ,String    ,品番
'説明      :
'履歴      :2001/05/17 作成  野村
Public Sub AddWfSample(ByVal pos As Integer, Optional ByVal ProdID$ = vbNullString, Optional SMPKBN$ = "D")
Attribute AddWfSample.VB_Description = "WFサンプル位置を指定する"
Dim WfSample As New c_cmzcWfSmp      'WFサンプル

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Sub AddWfSample"

    ''WFサンプル位置クラスをコレクションに追加する
    WfSample.IngotPos = pos
    WfSample.hinban = ProdID
    WfSample.SMPKBN = Left$(SMPKBN, 1)
    WfSmps.Add WfSample

PROC_EXIT:
    '終了
    Set WfSample = Nothing
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Sub


'概要      :結晶サンプル位置を指定する
'          :pos           ,I  ,Integer   ,指定位置
'説明      :
'履歴      :2001/05/17 作成  野村
Public Sub AddXlSample(ByVal pos As Integer)
Attribute AddXlSample.VB_Description = "結晶サンプル位置を指定する"
Dim XlSample As New c_cmzcXlSmp      '結晶サンプル

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Sub AddXlSample"

    ''結晶サンプル位置クラスをコレクションに追加する
    XlSample.IngotPos = pos
    XlSmps.Add XlSample

PROC_EXIT:
    '終了
    Set XlSample = Nothing
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Sub


'概要      :品番を追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :[length]      ,I  ,Integer   ,長さ
'          :[ProdID]      ,I  ,String    ,品番
'説明      :
'履歴      :2001/05/17 作成  野村
Public Sub AddHin(ByVal pos%, Optional ByVal LENGTH% = 9999, Optional ByVal ProdID$ = vbNullString)
Attribute AddHin.VB_Description = "品番区切位置を設定する"
Dim HIN As New c_cmzcHin      '品番

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Sub AddHin"

    ''品番区切位置クラスをコレクションに追加する
    HIN.IngotPos = pos
    HIN.LENGTH = LENGTH
    HIN.hinban = ProdID
    Hins.Add HIN

PROC_EXIT:
    '終了
    Set HIN = Nothing
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Sub


'概要      :品番設計を追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :[length]      ,I  ,Integer   ,長さ
'          :[ProdID]      ,I  ,String    ,品番
'説明      :
'履歴      :2001/05/17 作成  野村
Public Sub AddHinPlan(ByVal pos%, Optional ByVal LENGTH% = 9999, Optional ByVal ProdID$ = vbNullString)
Dim hinPlan As New c_cmzcHinPlan      '品番設計

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Sub AddHinPlan"

    ''品番区切位置クラスをコレクションに追加する
    hinPlan.IngotPos = pos
    hinPlan.LENGTH = LENGTH
    hinPlan.hinban = ProdID
    HinPlans.Add hinPlan

PROC_EXIT:
    '終了
    Set hinPlan = Nothing
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Sub


'概要      :切断指示を追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :[length]      ,I  ,Integer   ,長さ
'          :[BlockID]     ,I  ,String    ,ブロックID
'          :[hinID]       ,I  ,String    ,品番
'説明      :
'履歴      :2001/05/17 作成  野村
Public Sub AddCut(ByVal pos%, Optional ByVal LENGTH% = 9999, Optional ByVal BLOCKID$ = vbNullString, _
                  Optional ByVal hinban$ = vbNullString, Optional ByVal mnorevno% = 0, Optional ByVal factory$ = vbNullString, Optional ByVal opecond$ = vbNullString)
Dim Cut As New c_cmzcCut      '切断指示
Dim blkPos As Integer     'ブロックの開始位置

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Sub AddCut"

    ''次ブロックにかかるようなら、長さを調整する
    ''※他の切断位置との調整は切断指示コレクションが行う
    blkPos = mvarBlks.LowerArea(pos)
    If pos = blkPos Then
        '調整不要
    ElseIf pos > blkPos Then
        LENGTH = (blkPos + mvarBlks(CStr(blkPos)).LENGTH) - pos
    ElseIf pos + LENGTH > blkPos Then
        LENGTH = blkPos - pos
    End If

    ''切断指示クラスをコレクションに追加する
    Cut.IngotPos = pos
    Cut.LENGTH = LENGTH
    Cut.BLOCKID = BLOCKID
    Cut.hinban = hinban
    Cut.REVNUM = mnorevno
    Cut.factory = factory
    Cut.opecond = opecond
    mvarCuts.Add Cut

PROC_EXIT:
    '終了
    Set Cut = Nothing
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Sub


'概要      :ブロックを追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :[length]      ,I  ,Integer   ,長さ
'          :[BlockID]     ,I  ,String    ,ブロックID
'          :[ProcCode]    ,I  ,String    ,工程コード
'説明      :
'履歴      :2001/05/17 作成  野村
Public Sub AddBlk(ByVal pos%, Optional ByVal LENGTH% = 9999, Optional ByVal BLOCKID$ = vbNullString, Optional ByVal PROCCODE$ = vbNullString)
Attribute AddBlk.VB_Description = "切断位置を設定する"
Dim Blk As New c_cmzcBlk      'ブロック

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Sub AddBlk"

    ''ブロッククラスをコレクションに追加する
    Blk.IngotPos = pos
    Blk.LENGTH = LENGTH
    Blk.BLOCKID = BLOCKID
    Blk.NOWPROC = PROCCODE
    mvarBlks.Add Blk

PROC_EXIT:
    '終了
    Set Blk = Nothing
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Sub


'概要      :ブロック設計を追加する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :pos           ,I  ,Integer   ,指定位置
'          :[length]      ,I  ,Integer   ,長さ
'説明      :
'履歴      :2001/05/17 作成  野村
Public Sub AddBlkPlan(ByVal pos%, Optional ByVal LENGTH% = 9999)
Dim blkPlan As New c_cmzcBlkPlan      'ブロック設計

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Sub AddBlkPlan"

    ''ブロッククラスをコレクションに追加する
    blkPlan.IngotPos = pos
    blkPlan.LENGTH = LENGTH
    mvarBlkPlans.Add blkPlan

PROC_EXIT:
    '終了
    Set blkPlan = Nothing
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Sub


'概要      :内部データを初期化する
'説明      :
'履歴      :2001/05/17 作成  野村
Public Sub Clear()
Attribute Clear.VB_Description = "内部データを初期化する"

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Sub Clear"

    ''各プロパティの初期値を設定する
    CRYNUM = vbNullString           ' 結晶番号
    KRPROCCD = vbNullString         ' 管理工程コード
    PROCCD = vbNullString           ' 工程コード
    RPHINBAN = vbNullString         ' ねらい品番
    RPREVNUM = 0                    ' ねらい品番製品番号改訂番号
    RPFACT = vbNullString           ' ねらい品番工場
    RPOPCOND = vbNullString         ' ねらい品番操業条件
    PRODCOND = vbNullString         ' 製作条件
    PGID = vbNullString             ' ＰＧ－ＩＤ
    UPLENGTH = 0                    ' 引上げ長さ
    TOPLENG = 0                     ' ＴＯＰ長さ
    BODYLENG = 0                    ' 直胴長さ
    BOTLENG = 0                     ' ＢＯＴ長さ
    FREELENG = 0                    ' フリー長
    DIAMETER = 0                    ' 直径
    CHARGE = 0                      ' チャージ量
    SEED = vbNullString             ' シード
    ADDDPCLS = vbNullString         ' 追加ドープ種類
    ADDDPPOS = 0                    ' 追加ドープ位置
    ADDDPVAL = 0                    ' 追加ドープ量
    CurrentBlock = vbNullString     ' CurrentBlock プロパティ
    
    ''内部コレクションを初期化する
    Set mvarBlks = New c_cmzcBlks           ''ブロックコレクション
    Set mvarBlkPlans = New c_cmzcBlkPlans   ''ブロック設計コレクション
    Set mvarCuts = New c_cmzcCuts           ''切断指示コレクション
    Set mvarHins = New c_cmzcHins           ''品番コレクション
    Set mvarHinPlans = New c_cmzcHinPlans   ''品番設計コレクション
    Set mvarXlSmps = New c_cmzcXlSmps       ''結晶サンプルコレクション
    Set mvarWfSmps = New c_cmzcWfSmps       ''WFサンプルコレクション
    Set mvarSxls = New c_cmzcSxls           ''ＳＸＬコレクション
    Set mvarRejs = New c_cmzcRejs           ''欠落情報コレクション

PROC_EXIT:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Sub


'概要      :品番設計/品番データと切断指示データ/WFサンプル指示を元に、品番データを生成する
'説明      :
'履歴      :2001/06/26 作成  野村
Public Sub GenerateHin()
Dim hinPlan As c_cmzcHinPlan
Dim HIN As c_cmzcHin
Dim Cut As c_cmzcCut
Dim WFSMP As c_cmzcWfSmp
Dim lowerHin As c_cmzcHin
Dim pos As Integer
Dim NextPos As Integer
Dim UpperPos As Integer
Dim pos1 As Integer
Dim pos2 As Integer
Dim i As Integer

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Sub GenerateHin"

    ''品番設計データがあるとき（加工払出時）は、品番データにコピーする
    If mvarHinPlans.COUNT > 0 Then
        Set mvarHins = New c_cmzcHins   'まずクリア
        For Each hinPlan In mvarHinPlans
            Set HIN = New c_cmzcHin
            With HIN
                .CRYNUM = hinPlan.CRYNUM        ' 結晶番号
                .IngotPos = hinPlan.IngotPos    ' 結晶内開始位置
                .hinban = hinPlan.hinban        ' 品番
                .REVNUM = hinPlan.REVNUM        ' 製品番号改訂番号
                .factory = hinPlan.FACT         ' 工場
                .opecond = hinPlan.OPCOND       ' 操業条件
                .LENGTH = hinPlan.LENGTH        ' 長さ
            End With
            mvarHins.Add HIN
            Set HIN = Nothing
        Next
    End If
    
    ''切断指示位置を追加する(位置のみ、品番は維持)
    For Each Cut In mvarCuts
        Hins.AddLine Cut.IngotPos
    Next
    
    ''切断指示にしたがい、品番データを調整する
    For Each Cut In mvarCuts
        pos = Cut.IngotPos
        Set HIN = Hins(CStr(pos))
        If (Cut.hinban <> HIN.hinban) _
        Or (Cut.REVNUM <> HIN.REVNUM) _
        Or (Cut.factory <> HIN.factory) _
        Or (Cut.opecond <> HIN.opecond) Then   'if(現品番以外が指定されていた)
            ''次の切断指示・切断位置を求める
            pos1 = Blks.LowerPos(pos + 1)
            pos2 = Cuts.LowerPos(pos + 1)
            If (pos1 = 9999) And (pos2 = 9999) Then
                '以降に区切はない
                NextPos = BODYLENG
            ElseIf (pos1 < pos2) Then
                '切断位置までが変更対象
                Hins.AddLine pos1   '切断位置に品番区切を付加
                NextPos = pos1
            Else
                '切断指示位置までが変更対象
                NextPos = pos2
            End If
            
            ''次の切断指示・切断位置までにある品番境界を削除する
            For i = Hins.COUNT To 1 Step -1
                If (pos < Hins(i).IngotPos) And (Hins(i).IngotPos < NextPos) Then
                    Hins.Remove i
                End If
            Next
            
            ''品番を変更する
            With Hins(CStr(pos))
                .hinban = Cut.hinban
                .REVNUM = Cut.REVNUM
                .factory = Cut.factory
                .opecond = Cut.opecond
            End With
        End If
    Next
    
    ''不要な品番境界があれば整理する
    Hins.CompactLines

PROC_EXIT:
    '終了
    Set hinPlan = Nothing
    Set HIN = Nothing
    Set Cut = Nothing
    Set WFSMP = Nothing
    Set lowerHin = Nothing
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Sub

'概要      :ブロック設計/ブロックデータと切断指示データを元に、ブロックデータを生成する
'説明      :
'履歴      :2001/06/26 作成  野村
Public Sub GenerateBlk()
Dim blkPlan As c_cmzcBlkPlan
Dim Cut As c_cmzcCut

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Sub GenerateBlk"

    For Each blkPlan In BlkPlans
        With blkPlan
            If Blks.LowerPos(.IngotPos) = .IngotPos Then
                ''既にあった場合は何もしない(こないはずである)
                Debug.Print "GenerateBlk : 切断済(" & .IngotPos & ")"
            Else
                ''ブロックを追加する
                AddBlk .IngotPos, .LENGTH
            End If
        End With
    Next

PROC_EXIT:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Sub



'概要      :SXL管理データとWFサンプル指示を元に、SXL管理データを調整する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :genFrom       ,I  ,Integer   ,生成範囲(From)
'          :genTo         ,I  ,Integer   ,生成範囲(To)
'説明      :
'履歴      :2001/06/29 作成  野村
'           2001/10/05 引数追加  野村
Public Sub GenerateSxl(ByVal genFrom%, ByVal genTo%)
Dim sxl As c_cmzcSxl
Dim WFSMP As c_cmzcWfSmp
Dim pos As Integer
Dim pos1 As Integer
Dim pos2 As Integer
Dim posLimit As Integer
Dim i As Integer
Dim RET As FUNCTION_RETURN
Dim blkID As String
Dim blkPos As Integer
Dim SXLID As String
Dim HIN As c_cmzcHin
Dim isNotTailWFSMP As Boolean

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Sub GenerateSxl"

    ''処理前時点での最終位置を調べる
    pos1 = 0
    If (Sxls.COUNT) Then
        pos1 = Sxls(Sxls.COUNT).IngotPos
    End If
    pos2 = 0
    If (WfSmps.COUNT) Then
        pos2 = WfSmps(WfSmps.COUNT).IngotPos
    End If
    If (pos1 > pos2) Then
        posLimit = pos1
    Else
        posLimit = pos2
    End If

    ''WFサンプル指示にしたがい、SXLデータを調整する
'    For Each WfSmp In mvarWfSmps
'    For i = 1 To mvarWfSmps.COUNT - 1
    For i = mvarWfSmps.COUNT To 1 Step -1
        Set WFSMP = mvarWfSmps(i)
        pos = WFSMP.IngotPos
        If (pos < genFrom) Or (pos > genTo) Then
            '対象範囲外なので何もしない
        ElseIf (isNotTailWFSMP = False) Then
            '最後のWFサンプル以降はSXLとしない
            isNotTailWFSMP = True
        ElseIf (pos <> Sxls.LowerPos(pos)) Then
            'そこがSXL開始位置でなければ追加する
            RET = GetBlkPos(WFSMP.IngotPos, blkID, blkPos)
            If RET = FUNCTION_RETURN_SUCCESS Then
                SXLID = Left$(blkID, 10) & Format$(blkPos, "000")
            Else
                Debug.Print "ブロックが登録されていないので正しいSXL-IDが作成できない(" & WFSMP.IngotPos & ")"
                SXLID = WFSMP.CRYNUM & "?" & Format$(blkPos, "000")
            End If
            Set sxl = New c_cmzcSxl
            With sxl
                .CRYNUM = WFSMP.CRYNUM             ' 結晶番号
                .IngotPos = WFSMP.IngotPos          ' 結晶内開始位置
                '.LENGTH = 9999            ' 長さ(自動調整されるので長くてよい)
                .LENGTH = genTo - pos       '長さ(生成する最大長さは「生成範囲(To)」
                .SXLID = SXLID              ' SXLID
                '.KRPROCCD = String           ' 管理工程(呼出側でしか判断できない)
                '.NOWPROC = String            ' 現在工程(呼出側でしか判断できない)
                '.LPKRPROCCD = String         ' 最終通過管理工程(呼出側でしか判断できない)
                '.LASTPASS = String           ' 最終通過工程(呼出側でしか判断できない)
                .DELCLS = DELCLS_NORMAL       ' 削除区分
                .LSTATCLS = LSTATCLS_NORMAL   ' 最終状態区分
                .HOLDCLS = HOLDCLS_NORMAL     ' ホールド区分
                If ((WFSMP.SMPKBN = "D") Or (WFSMP.SMPKBN = "T") Or (WFSMP.SMPKBN = "")) And (Trim$(WFSMP.hinban) <> vbNullString) Then
                    '下品番
                    .hinban = WFSMP.hinban        ' 品番
                    .REVNUM = WFSMP.REVNUM        ' 製品番号改訂番号
                    .factory = WFSMP.factory      ' 工場
                    .opecond = WFSMP.opecond      ' 操業条件
                Else
                    '下品番を得る
                    pos1 = Hins.LowerArea(.IngotPos)
                    If pos1 <> 9999 Then
                        Set HIN = Hins(CStr(pos1))
                        .hinban = HIN.hinban
                        .REVNUM = HIN.REVNUM
                        .factory = HIN.factory
                        .opecond = HIN.opecond
                        Set HIN = Nothing
                    Else
                        .hinban = vbNullString
                        .REVNUM = 0
                        .factory = vbNullString
                        .opecond = vbNullString
                    End If
                End If
                '.BDCAUS = " "                 ' 不良理由
                '.Count = 0                    ' 枚数(呼出側で計算するのでは？)
            End With
            Sxls.Add sxl
            Set sxl = Nothing
        Else
            'そこがSXL開始位置の場合は、品番変更に対応する(廃棄のみ可だが、本クラスでは'Z'かどうかを意識しない)
            Set sxl = Sxls(CStr(pos))
            If (WFSMP.SMPKBN = "D") Or (WFSMP.SMPKBN = "T") _
            And ((WFSMP.hinban <> sxl.hinban) _
                 Or (WFSMP.REVNUM <> sxl.REVNUM) _
                 Or (WFSMP.factory <> sxl.factory) _
                 Or (WFSMP.opecond <> sxl.opecond) _
            ) Then   'if(現品番以外が指定されていた)
                ''品番を変更する
                With Sxls(CStr(pos))
                    .hinban = WFSMP.hinban
                    .REVNUM = WFSMP.REVNUM
                    .factory = WFSMP.factory
                    .opecond = WFSMP.opecond
                End With
            End If
            Set sxl = Nothing
        End If
    Next
    
    ''上限を制限する
    pos1 = 0
    If (Sxls.COUNT) Then
        With Sxls(Sxls.COUNT)
            If .IngotPos + .LENGTH >= 9999 Then
                Sxls.LimitByIngotPos posLimit
            End If
        End With
    End If
        

PROC_EXIT:
    '終了
    Set sxl = Nothing
    Set WFSMP = Nothing
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Sub




'概要      :サンプルＩＤの存在するべきインゴット内位置を得る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :SampleID      ,I  ,String    ,サンプルＩＤ
'          :戻り値        ,O  ,Integer   ,インゴット内位置（無効の場合はマイナス値）
'説明      :存在しないサンプルＩＤでも値を返す（ブロックＩＤ部分が無効の場合は-9999）
'履歴      :2001/06/29 作成  野村
Public Function GetIngotPosBySampleID(SAMPLEID As String) As Integer
Dim pos As Integer

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Function GetIngotPosBySampleID"

    ''形式チェック
    If Len(SAMPLEID) <> 16 Then
        GetIngotPosBySampleID = -9999
        GoTo PROC_EXIT
    End If
    
    ''ブロックＩＤを検索し、見つかればその開始位置＋ブロック内位置を返す
    pos = Blks.GetPosByID(Left$(SAMPLEID, 12))
    If pos < 0 Then
        GetIngotPosBySampleID = -9999
    Else
        GetIngotPosBySampleID = pos + Val(Mid$(SAMPLEID, 13, 3))
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


'概要      :ブロックＩＤとブロック内位置から、インゴット内位置を得る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :BlockID       ,I  ,String    ,ブロックＩＤ
'          :BlockPos      ,I  ,Integer   ,ブロック内位置
'          :IngotPos      ,O  ,Integer   ,インゴット内位置
'          :戻り値        ,O  ,FUNCTION_RETURN, 成否
'説明      :
'履歴      :2001/06/29 作成  野村
Public Function GetIngotPos(BLOCKID As String, BlockPos As Integer, IngotPos As Integer) As FUNCTION_RETURN

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Function GetIngotPos"

    IngotPos = Blks.GetPosByID(BLOCKID)
    If IngotPos < 0 Then
        IngotPos = 0
        GetIngotPos = FUNCTION_RETURN_FAILURE
    Else
        IngotPos = IngotPos + BlockPos
        GetIngotPos = FUNCTION_RETURN_SUCCESS
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


'概要      :インゴット内位置から、ブロックＩＤとブロック内位置を得る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :IngotPos      ,I  ,Integer   ,インゴット内位置
'          :BlockID       ,O  ,String    ,ブロックＩＤ
'          :BlockPos      ,O  ,Integer   ,ブロック内位置
'          :戻り値        ,O  ,FUNCTION_RETURN, 成否
'説明      :
'履歴      :2001/06/29 作成  野村
Public Function GetBlkPos(IngotPos As Integer, BLOCKID As String, BlockPos As Integer) As FUNCTION_RETURN
Dim pos As Integer
Dim Blk As c_cmzcBlk


    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Function GetBlkPos"

'2001/08/31 S.Sano    pos = Blks.LowerPos(IngotPos)
    pos = Blks.UpperPos(IngotPos + 1) '2001/08/31 S.Sano
    If pos >= 0 And pos < 9999 Then
        Set Blk = Blks(CStr(pos))
        BLOCKID = Blk.BLOCKID
        BlockPos = IngotPos - Blk.IngotPos
        GetBlkPos = FUNCTION_RETURN_SUCCESS
    Else
        BLOCKID = vbNullString
        BlockPos = -9999
        GetBlkPos = FUNCTION_RETURN_FAILURE
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

Public Sub Dump()
Dim i As Integer
Dim cnt As Integer

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    If Not (gErr Is Nothing) Then gErr.Push "c_cmzcXl.cls -- Sub Dump"

    Debug.Print "========= c_cmzcXl Dump(" & CRYNUM & ") =========="
    Debug.Print "CRYNUM = " & CRYNUM
    Debug.Print " = " & CRYNUM
    Debug.Print "UPLENGTH=" & UPLENGTH          ' 引上げ長さ"
    Debug.Print "TOPLENG=" & TOPLENG           ' ＴＯＰ長さ"
    Debug.Print "BODYLENG=" & BODYLENG          ' 直胴長さ"
    Debug.Print "BOTLENG=" & BOTLENG           ' ＢＯＴ長さ"
    Debug.Print "FREELENG=" & FREELENG          ' フリー長"
    
    If BlkPlans Is Nothing Then
        Debug.Print "BlkPlans is Nothing"
    Else
        With BlkPlans
            cnt = .COUNT
            Debug.Print "BlkPlans.Count = " & .COUNT
            For i = 1 To cnt
                Debug.Print "(" & i & ") :";
                Debug.Print " INGOTPOS=" & .Item(i).IngotPos;
                Debug.Print " LENGTH=" & .Item(i).LENGTH
            Next
        End With
    End If
    
    If HinPlans Is Nothing Then
        Debug.Print "HinPlans is Nothing"
    Else
        With HinPlans
            cnt = .COUNT
            Debug.Print "HinPlans.Count = " & .COUNT
            For i = 1 To cnt
                Debug.Print "(" & i & ") :";
                Debug.Print " INGOTPOS=" & .Item(i).IngotPos;
                Debug.Print " LENGTH=" & .Item(i).LENGTH;
                Debug.Print " HINBAN12=" & .Item(i).hinban & "-" & Format$(.Item(i).REVNUM, "00") & .Item(i).FACT & .Item(i).OPCOND
            Next
        End With
    End If

    If Blks Is Nothing Then
        Debug.Print "Blks is Nothing"
    Else
        With Blks
            cnt = .COUNT
            Debug.Print "Blks.Count = " & .COUNT
            For i = 1 To cnt
                Debug.Print "(" & i & ") :";
                Debug.Print " INGOTPOS=" & .Item(i).IngotPos;
                Debug.Print " LENGTH=" & .Item(i).LENGTH;
                Debug.Print " BLOCKID=" & .Item(i).BLOCKID
            Next
        End With
    End If
    
    If Hins Is Nothing Then
        Debug.Print "Hins is Nothing"
    Else
        With Hins
            cnt = .COUNT
            Debug.Print "Hins.Count = " & .COUNT
            For i = 1 To cnt
                Debug.Print "(" & i & ") :";
                Debug.Print " INGOTPOS=" & .Item(i).IngotPos;
                Debug.Print " LENGTH=" & .Item(i).LENGTH;
                Debug.Print " HINBAN12=" & .Item(i).hinban & "-" & Format$(.Item(i).REVNUM, "00") & .Item(i).factory & .Item(i).opecond
            Next
        End With
    End If
    
    If Cuts Is Nothing Then
        Debug.Print "Cuts is Nothing"
    Else
        With Cuts
            cnt = .COUNT
            Debug.Print "Cuts.Count = " & .COUNT
            For i = 1 To cnt
                Debug.Print "(" & i & ") :";
                Debug.Print " INGOTPOS=" & .Item(i).IngotPos;
                Debug.Print " LENGTH=" & .Item(i).LENGTH;
                Debug.Print " HINBAN12=" & .Item(i).hinban & "-" & Format$(.Item(i).REVNUM, "00") & .Item(i).factory & .Item(i).opecond;
                Debug.Print " BLOCKID=" & .Item(i).BLOCKID
            Next
        End With
    End If
    
    If XlSmps Is Nothing Then
        Debug.Print "XlSmps is Nothing"
    Else
        With XlSmps
            cnt = .COUNT
            Debug.Print "XlSmps.Count = " & .COUNT
            For i = 1 To cnt
                Debug.Print "(" & i & ") :";
                Debug.Print " INGOTPOS=" & .Item(i).IngotPos;
                Debug.Print " SMPKBN=" & .Item(i).SMPKBN;
                Debug.Print " HINBAN12=" & .Item(i).hinban & "-" & Format$(.Item(i).REVNUM, "00") & .Item(i).factory & .Item(i).opecond
            Next
        End With
    End If
    
    If WfSmps Is Nothing Then
        Debug.Print "WfSmps is Nothing"
    Else
        With WfSmps
            cnt = .COUNT
            Debug.Print "WfSmps.Count = " & .COUNT
            For i = 1 To cnt
                Debug.Print "(" & i & ") :";
                Debug.Print " INGOTPOS=" & .Item(i).IngotPos;
                Debug.Print " SMPKBN=" & .Item(i).SMPKBN;
                Debug.Print " HINBAN12=" & .Item(i).hinban & "-" & Format$(.Item(i).REVNUM, "00") & .Item(i).factory & .Item(i).opecond
            Next
        End With
    End If
    
    If Sxls Is Nothing Then
        Debug.Print "Sxls is Nothing"
    Else
        With Sxls
            cnt = .COUNT
            Debug.Print "Sxls.Count = " & .COUNT
            For i = 1 To cnt
                Debug.Print "(" & i & ") :";
                Debug.Print " INGOTPOS=" & .Item(i).IngotPos;
                Debug.Print " LENGTH=" & .Item(i).LENGTH;
                Debug.Print " HINBAN12=" & .Item(i).hinban & "-" & Format$(.Item(i).REVNUM, "00") & .Item(i).factory & .Item(i).opecond
            Next
        End With
    End If

    If Rejs Is Nothing Then
        Debug.Print "Rejs is Nothing"
    Else
        With Rejs
            cnt = .COUNT
            Debug.Print "Rejs.Count = " & .COUNT
            For i = 1 To cnt
                Debug.Print "(" & i & ") :";
                Debug.Print " LOTID=" & .Item(i).LOTID;
                Debug.Print " ALLSCRAP=" & .Item(i).ALLSCRAP;
                Debug.Print " LENFROM=" & .Item(i).LENFROM;
                Debug.Print " LENTO=" & .Item(i).LENTO
            Next
        End With
    End If

PROC_EXIT:
    '終了
    If Not (gErr Is Nothing) Then gErr.Pop
    Exit Sub

PROC_ERR:
    'エラーハンドラ
    If Not (gErr Is Nothing) Then gErr.HandleError
    Resume PROC_EXIT
End Sub


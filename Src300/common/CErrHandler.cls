VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CErrHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' クラス  : CErrHandler
' 解説    : グローバルなエラーハンドラ
' 作成者 : Total VB SourceBook 6
'
' イベント
' ------------------------------------------------------
' エラーがトラップされたときに発生します。CancelパラメータをTrue
' に設定することで、それ以降の処理をバイパスすることができます。
Public Event BeforeHandlerCalled(Cancel As Boolean)

' イベントがハンドリングされたあとで発生し、必要に応じて追加の
' 処理を行うことができます。例えば、フォームでエラーコードを表
' 示したり、エラーのログを作成することができます。
Public Event AfterHandlerCalled()

' プロパティの値を格納するローカル変数
' ------------------------------------------------------
Private m_intMaxProcStackItems As Integer
Private m_fIncludeExpandedInfo As Boolean
Private m_fDisplayMsgOnError As Boolean
Private m_lngErrorNumber As Long
Private m_lngErrorLine As Long
Private m_strErrorDescription As String
Private m_strProcName As String
Private m_strAppTitle As String
Private m_varDestination As Variant
Private m_strCurrentOperation As String
Private m_fTraceExecution As Boolean

' クラスで使用するプライベート変数
' ------------------------------------------------------

' プロシージャ名のスタックを格納するための配列。スタックのサイ
' ズは、プロパティの設定によって決定され、デフォルトは20となっ
' ています。
Private mastrErrorStack() As String

' スタックの現在位置を示すポインタ
Private mintCurStackPointer As Integer

' 独自のハンドリングのために２つ追加します。
Private mintUseItems As Integer

Private mstrLogFile As String
Private mintLogFile As Integer
Private Const mcstrTraceFile = "trace.log"

Private Sub Class_Initialize()
    ' 解説        : プロパティの設定値で上書きされることの
    '                ある値にデフォルトの初期値を設定します。
    ' 作成者    : Total VB SourceBook 6

    m_intMaxProcStackItems = 100
    mintUseItems = m_intMaxProcStackItems + 2
    
    m_fIncludeExpandedInfo = True
    m_fDisplayMsgOnError = True
    m_fTraceExecution = False
    m_varDestination = "C:\ERRORS.TXT"
    
    ReDim mastrErrorStack(0 To mintUseItems)
    mintCurStackPointer = LBound(mastrErrorStack)
    
End Sub

Private Sub Class_Terminate()
    ' 解説        : もしもオープンされている場合は、
    '                ログのトレースをクローズし、Popの
    '                呼び出しとPushの呼び出しがバラ
    '                ンスしていることを確認します。
    ' 作成者    : Total VB SourceBook 6
    
    If m_fTraceExecution Then
        Close mintLogFile
    End If
    
    Debug.Assert mintCurStackPointer = 0
    
End Sub

Friend Property Get AppTitle() As String
    ' 戻り値     : 現在のAppTitleの値
    ' 作成者    : Total VB SourceBook 6

    AppTitle = m_strAppTitle

End Property

Friend Property Let AppTitle(ByVal strValue As String)
    ' strValue: エラーハンドラのメッセージボックスで使われるタイトル。
    '             app.titleプロパティの値となっている標準的なタイトルを
    '             変更したい場合に使います。
    ' 作成者  : Total VB SourceBook 6

m_strAppTitle = strValue

End Property

Friend Property Get CurrentOperation() As String
    ' 戻り値     : CurrentOperationの現在の値
    ' 作成者    : Total VB SourceBook 6

    CurrentOperation = m_strCurrentOperation
    
End Property

Friend Property Let CurrentOperation(ByVal strValue As String)
    ' strValue: 現在の処理の名前を含んだ文字列。これは、エラー
    '             の発生源を示すユーザー定義のチェックポイントです。
    ' 作成者  : Total VB SourceBook 6
    
    m_strCurrentOperation = strValue
    
End Property

Public Property Get Destination() As Variant
    ' 解説       : このプロパティにエラーログのパス名の
    '               文字列を設定した場合は、文字列が返さ
    '               れます。このプロパティにレコードセットオ
    '               ブジェクトを設定した場合は、レコードセッ
    '               トオブジェクトへのポインタが返されます。
    ' 戻り値     : Destinationプロパティの現在の値
    ' 作成者    : Total VB SourceBook 6
    '
    If IsObject(m_varDestination) Then
        Set Destination = m_varDestination
    Else
        Destination = m_varDestination
    End If

End Property

Public Property Let Destination(ByVal varValue As Variant)
    ' strValue: エラーの出力先を指定します。これは、テキスト
    '             ファイルのパス名を含んだ文字列、またはレコー
    '             ドセットのオブジェクト変数となります。
    ' 作成者 : Total VB SourceBook 6
    '
    Select Case VarType(varValue)
        Case vbString
            m_varDestination = varValue
        Case vbObject
                Err.Raise _
                    vbObjectError + 1000, , _
                    "Invalid object type: " & VarType(varValue)
        Case Else
            Err.Raise _
                vbObjectError + 1000, , "Invalid object type: " & VarType(varValue)
    End Select
    
End Property

Public Property Get DisplayMsgOnError() As Boolean
    ' 戻り値     : DisplayMsgOnErrorの現在の値
    ' 作成者    : Total VB SourceBook 6
    
    DisplayMsgOnError = m_fDisplayMsgOnError

End Property

Public Property Let DisplayMsgOnError(ByVal fValue As Boolean)
    ' fValue: True（デフォルト）に設定すると、エラーがハンドリングされ
    '           たときに基本的な情報を含んだメッセージボックスを表示
    '           します。Falseに設定するとメッセージボックスを表示しま
    '           せん。
    ' 作成者: Total VB SourceBook 6

    m_fDisplayMsgOnError = fValue
    
End Property

Friend Property Get ErrorDescription() As String
    ' 戻り値     : ErrorDescriptionの現在の値。この値には、
    '                トリガされるエラーハンドラが発生するエラ
    '                ーの説明が設定されます。AppSpecificErrorHandler
    '                プロシージャまたはAfterHandlerCalledイ
    '                ベントに応答するコードで、このプロパティ
    '                を参照してください。
    ' 作成者    : Total VB SourceBook 6
    
    ErrorDescription = m_strErrorDescription
    
End Property

Friend Property Get ErrorLine() As Long
    ' 戻り値     : ErrorLineの現在の値。この値には、エラーハンドラ
    '                がトリガされたときにエラーを発生しているプロシー
    '                ジャの行番号が設定されます。AppSpecificErrorHandler
    '                プロシージャまたはAfterHandlerCalledイベントに
    '                応答するコードで、このプロパティを参照してください。
    ' 作成者    : Total VB SourceBook 6
    '
    ErrorLine = m_lngErrorLine

End Property

Friend Property Get ErrorNumber() As Long
    ' 戻り値     : ErrorNumberプロパティの現在の値。この値には、
    '                エラーハンドラがトリガされたときにエラーを発生
    '                しているプロシージャのエラー番号が設定されま
    '                す。AppSpecificErrorHandlerプロシージャまたは
    '                AfterHandlerCalledイベントに応答するコードで、
    '                このプロパティを参照してください。
    ' 作成者    : Total VB SourceBook 6
    '
    ErrorNumber = m_lngErrorNumber

End Property

Public Property Get IncludeExpandedInfo() As Boolean
    ' 戻り値     : IncludeExpandedInfoの現在の値
    ' 作成者    : Total VB SourceBook 6
        
    IncludeExpandedInfo = m_fIncludeExpandedInfo

End Property

Public Property Let IncludeExpandedInfo(ByVal fValue As Boolean)
    ' fValue: True（デフォルト）に設定すると、エラーログにユーザーの
    '           マシン環境に関する追加情報を含めます。Falseに設定す
    '           ると、エラーログに拡張情報を含めません。
    ' 作成者: Total VB SourceBook 6
    '
    m_fIncludeExpandedInfo = fValue

End Property

Public Property Get MaxProcStackItems() As Integer
    ' 戻り値     : MaxProcStackItemsプロパティの現在の値
    ' 作成者    : Total VB SourceBook 6
    '
    MaxProcStackItems = m_intMaxProcStackItems

End Property

Public Property Let MaxProcStackItems(ByVal intValue As Integer)
    ' intValue: デフォルトの20以外の値にプロシージャスタックのサイズ
    '             を設定します。かなり深くネストしたプロシージャ呼び出し
    '             を使っていない限り、この値を大きくする必要はありません。
    ' 作成者 : Total VB SourceBook 6
    
    m_intMaxProcStackItems = intValue
    mintUseItems = m_intMaxProcStackItems + 2
    ReDim mastrErrorStack(0 To mintUseItems)
    
End Property

Friend Property Get ProcName() As String
    ' 戻り値     : ProcNameプロパティの現在の値。この値には、
    '                エラーハンドラをトリガしたエラーを含んだプロ
    '                シージャの名前が設定されます。
    '                AppSpecificErrorHandlerプロシージャまたは
    '                AfterHandlerCalledイベントに応答するコードで、
    '                このプロパティを参照してください。
    ' 作成者    : Total VB SourceBook 6
    '
    ProcName = m_strProcName

End Property

Public Property Get TraceExecution() As Boolean
    ' 戻り値     : TraceExecutionの現在の値
    ' 作成者    : Total VB SourceBook 6

    TraceExecution = m_fTraceExecution

End Property

Public Property Let TraceExecution(ByVal fValue As Boolean)
    ' fValue: Trueに設定すると、プログラムのトレースを開始します。
    '           Falseに設定すると、停止します。この値をTrueに設定
    '           すると、ログファイルが削除されてから再び作成されます。
    ' 作成者: Total VB SourceBook 6
    '
    On Error GoTo PROC_ERR
    
    m_fTraceExecution = fValue
    
    If m_fTraceExecution Then
        ' ログを開始します。
        mstrLogFile = App.Path & "\" & mcstrTraceFile
    
        On Error Resume Next
        Kill mstrLogFile
        
        On Error GoTo PROC_ERR
        mintLogFile = FreeFile
        Open mstrLogFile For Binary As mintLogFile
    
    End If

proc_exit:
    Exit Property

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "TraceExecution"
    Resume proc_exit
    
End Property

Public Function ClearLog() As Boolean
    ' 解説        : エラーログ（またはDestinationプロパティ
    '                で指定されたテーブル）を消去します。
    ' パラメータ: なし
    ' 戻り値     : 成功ならTrue、それ以外はFalse
    ' 作成者    : Total VB SourceBook 6
    '
    On Error GoTo PROC_ERR

    ' ログは、テキストファイルまたはレコードセットです。
    ' Destinationプロパティがオブジェクト（レコードセット）
    ' または文字列のどちらを指しているかをチェックします。
    If IsObject(m_varDestination) Then
        
        ' Destinationはレコードセットです。
    
    Else
        ' Destinationはログファイルのパス名です。
        If VarType(m_varDestination) = vbString Then
            ' ファイルが存在しない場合もあるため、エラーハンドリングが
            ' ランタイムエラーを無視するようにします。
            On Error Resume Next
            Kill m_varDestination
            On Error GoTo PROC_ERR
        End If
    End If
        
    ClearLog = True

proc_exit:
    Exit Function

PROC_ERR:
    ClearLog = False
    Resume proc_exit

End Function

Public Sub HandleError()
    ' 解説        : マスターエラーハンドラ。このメソッドは、このメ
    '                ソッドを呼び出すコードへランタイムエラーが分
    '                岐する場合に呼び出されます。このメソッドは、
    '                エラーが発生したあとだけに呼び出されるように
    '                します。詳細はリファレンスやこのメソッドの呼
    '                び出し方を示すサンプルを参照してください。
    ' パラメータ: なし
    ' 戻り値     : なし
    ' 作成者    : Total VB SourceBook 6
    '
    Dim strTemp As String
    Dim fCancel As Boolean
    
    ' 現在のエラー情報を取得し保存します。
    m_lngErrorNumber = Err.Number
    m_lngErrorLine = Erl
    m_strErrorDescription = Err.Description
    '' 2003/11/10
''    m_strProcName = mastrErrorStack(mintCurStackPointer)
    If mintCurStackPointer <= mintUseItems Then
        m_strProcName = mastrErrorStack(mintCurStackPointer)
    End If
    
    ' 特定のタイトルが指定されていない場合は、App.titleプロパ
    ' ティを使います。
    If m_strAppTitle = "" Then
        m_strAppTitle = App.Title
    End If
    
    ' エラーハンドラがトリガされたことを示すイベントを発生させ
    ' ます。fCancel変数の値が発生させたイベントで変更された
    ' 場合は、これ以降の処理をバイパスします。詳細は、この
    ' イベントのリファレンスを参照してください。
    fCancel = False
    RaiseEvent BeforeHandlerCalled(fCancel)
    
    ' fCancelの値が変更されていない場合は、処理を継続します。
    If Not fCancel Then
    
        ' 現在どのような状態にあるかが特定できないので、エラー
        ' トラップを無効にします。
        On Error Resume Next
        
        ' Destinationにレコードセットが設定されている場合は、その
        ' レコードセットの行に情報を記録するルーチンを呼び出しま
        ' す。ディスクファイルの名前を含んだ文字列の場合は、ディ
        ' スクファイルに情報を書き込むルーチンを呼び出します。
        If IsObject(m_varDestination) Then
            ' オブジェクトはレコードセットまたは未定義です。
        Else
            ' オブジェクトはログファイルの名前を含んだ文字列です。
            Call LogErrorToFile
        End If

        ' エラーハンドリングを無効にします。
        On Error GoTo 0

        ' DisplayMsgOnErroフラグがTrueに設定されている場合は、
        ' メッセージボックスにエラー情報を表示します。
        If m_fDisplayMsgOnError Then
            
            Beep
            
            ' メッセージの文字列を作成します。
            '' 2003/11/10
''            strTemp = "The following error has occurred in procedure: " & _
''                vbCrLf & _
''                mastrErrorStack(mintCurStackPointer) & _
''                " on line " & _
''                IIf(IsNull(m_lngErrorLine), "<unknown>", m_lngErrorLine) & _
''                vbCrLf & _
''                "Error (" & m_lngErrorNumber & ") " & m_strErrorDescription
            If mintCurStackPointer <= mintUseItems Then
                strTemp = "The following error has occurred in procedure: " & _
                    vbCrLf & _
                    mastrErrorStack(mintCurStackPointer) & _
                    " on line " & _
                    IIf(IsNull(m_lngErrorLine), "<unknown>", m_lngErrorLine) & _
                    vbCrLf & _
                    "Error (" & m_lngErrorNumber & ") " & m_strErrorDescription
            End If

            If m_strCurrentOperation <> "" Then
                strTemp = strTemp & vbCrLf & _
                    "Current Operation: " & m_strCurrentOperation
            End If
            
            MsgBox strTemp, vbCritical, m_strAppTitle
        End If
    
        ' エラーハンドラが処理を終了したことを示すイベントを発生
        ' します。ここは、アプリケーションでエラー情報のメッセー
        ' ジボックスを表示したり、ユーザーにその対処方法を問い
        ' 合わせるのによい場所です。
        RaiseEvent AfterHandlerCalled
        
    End If
    
    ' このプロシージャの解説を参照してください。
    Call AppSpecificErrHandler
        
End Sub

Public Sub Pop()
    ' 解説        : エラーハンドリングスタックから現在のプロ
    '                シージャ名をポップします。エラーなしでコ
    '                ードが正常終了するときに、このメソッドを
    '                呼び出してください。
    ' パラメータ: なし
    ' 戻り値     : なし
    ' 作成者    : Total VB SourceBook 6
    '
'Debug.Print String(mintCurStackPointer, ".") & " pop"
    ' スタック上の最後の項目を削除します。
    If mintCurStackPointer <= m_intMaxProcStackItems Then
        mastrErrorStack(mintCurStackPointer) = ""
    End If

    ' １つ前の項目を指すように現在のスタックポインタをデ
    ' クリメントします。
    mintCurStackPointer = mintCurStackPointer - 1
    
    If mintCurStackPointer < 0 Then
        mintCurStackPointer = 0
    End If

End Sub

Public Sub Push(strProc As String)
    ' 解説        : 与えられたプロシージャの名前をエラーハンドリング
    '                スタックにプッシュします。プロシージャの開始位置
    '                でこのメソッドを呼び出してください。
    ' パラメータ: strProc - 現在実行されているプロシージャ名
    ' 戻り値     : なし
    ' 作成者    : Total VB SourceBook 6
    '
    ' 次の空き位置を指すように現在のスタックポインタをインクリメントします。
    mintCurStackPointer = mintCurStackPointer + 1

    If mintCurStackPointer <= m_intMaxProcStackItems Then
        mastrErrorStack(mintCurStackPointer) = strProc
    Else
        '' 2003/11/10
''        mastrErrorStack(mintCurStackPointer + 2) = strProc
        If mintCurStackPointer <= mintUseItems - 2 Then
            mastrErrorStack(mintCurStackPointer + 2) = strProc
        End If
    End If
    
    ' ログのトレースが有効の場合は、ログファイルに格納します。
    If m_fTraceExecution Then
        Put mintLogFile, , strProc & vbCrLf
    End If

'Debug.Print String(mintCurStackPointer, ".") & " push " & strProc
End Sub

Private Sub AppSpecificErrHandler()
            ' theBeforeHandlerCalledまたはAfterHandlerCalledイベント
            ' に応答するためのエラー処理を行わずに、エラー情報をメ
            ' ッセージボックスに表示するだけのエラーハンドラをカスタ
            ' マイズしたい場合は、AppSpecificErrHandlerプロシージャ
            ' にアプリケーション独自のエラーハンドリングロジックを記
            ' 述することができます。これは、このクラスで変更する必
            ' 要のある唯一のプロシージャです。
        
            ' アプリケーション独自のロジックを記述します。
    
End Sub

Private Sub LogErrorToFile()
    ' 解説        : ディスクファイルに最新のエラー情報を書き込
    '                みます。ファイルの名前は、ファイルのパス名
    '                を含んだ文字列を含んだエラーハンドラオブジ
    '                ェクトのDestinationプロパティによって指定され
    '                たものです。
    ' パラメータ: なし
    ' 戻り値     : なし
    ' 作成者    : Total VB SourceBook 6
    '
    Dim intFile As Integer
    Dim intCounter As Integer
    Dim varTemp As Variant
    
    '-*-*- 人見　ログファイルを１Mまで残すように修正
    '-*-*- 野村　画面表示しないエラーはログにも残さないように変更
    
    On Error Resume Next
    If m_fDisplayMsgOnError = False Then Exit Sub
    
    Dim bakfile As String
    bakfile = m_varDestination & ".bak"     ' バックアップファイル名
    
    If Dir(m_varDestination) <> "" Then
        If FileLen(m_varDestination) >= 1048575 Then
            If Dir(bakfile) <> "" Then
                Kill bakfile
            End If
            Name m_varDestination As bakfile
        End If
    End If
    
    ' 下の２行は元のソース
    ' 書き込み先のファイルが既に存在する場合は削除します。
    ' Kill m_varDestination
    
    ' エラーハンドリングを無効にします。
    On Error GoTo 0

    intFile = FreeFile
    Open m_varDestination For Append As intFile
    
    Print #intFile, "///////////////////////////////////////////"  ' 区切り線追加  -*-*- 人見
    
    Print #intFile, m_strAppTitle & " ｱﾌﾟﾘｹｰｼｮﾝ ｴﾗｰ情報"
    Print #intFile, "===================================="
    Print #intFile, "発生時間  : " & Now
    Print #intFile, "ｴﾗｰ文字列  : "; Left$(m_strErrorDescription, 255)
    Print #intFile, "ｴﾗｰ番号  : " & m_lngErrorNumber
    Print #intFile, "発生行  : " & m_lngErrorLine
    Print #intFile, "ｴﾗｰﾌﾟﾛｼｰｼﾞｬ : " & _
        mastrErrorStack(mintCurStackPointer)
    Print #intFile,
    Print #intFile, "ﾌﾟﾛｼｰｼﾞｬ ｽﾀｯｸ"
    Print #intFile, "---------------"
        
    For intCounter = LBound(mastrErrorStack) To UBound(mastrErrorStack)
        If mastrErrorStack(intCounter) <> "" Then
            varTemp = varTemp & _
                Format(intCounter, "00") & Space$(intCounter * 2) & _
                mastrErrorStack(intCounter) & vbCrLf
        End If
    Next intCounter
    
    Print #intFile, varTemp

    If m_strCurrentOperation <> "" Then
        Print #intFile, "Current Operation: " & m_strCurrentOperation
    End If
    
    If m_fIncludeExpandedInfo Then
        Print #intFile,
        Print #intFile, "追加情報"
        Print #intFile, "----------------------"

        varTemp = "ｶﾚﾝﾄﾃﾞｨﾚｸﾄﾘ: " & CurDir & vbCrLf
        varTemp = varTemp & "ﾀｲﾄﾙ : " & App.Title & vbCrLf
        varTemp = varTemp & "会社名 : " & App.CompanyName & vbCrLf
        varTemp = varTemp & "ﾒｼﾞｬｰﾊﾞｰｼﾞｮﾝ r: " & App.Major & vbCrLf
        varTemp = varTemp & "ﾏｲﾅｰﾊﾞｰｼﾞｮﾝ : " & App.Minor & vbCrLf
        varTemp = varTemp & "ﾘﾋﾞｼﾞｮﾝ : " & App.Revision & vbCrLf
        varTemp = varTemp & "ｱﾌﾟﾘｹｰｼｮﾝのﾊﾟｽ : " & App.Path & vbCrLf
        varTemp = varTemp & "製品名 : " & App.ProductName
        
        ' 必要ならば、VB SourceBookが提供する関数を使って、
        ' Windowsディレクトリ、ディスクの空き容量などのさまざ
        ' まな項目を取得することができます。

        Print #intFile, varTemp

    End If

    Print #intFile, "///////////////////////////////////////////"  ' 区切り線追加  -*-*- 人見


    ' オブジェクトをクローズします。
    Close #intFile
    
End Sub

Private Sub LogErrorToTable()
    ' 解説        : テーブルに最新のエラー情報を書き込みます。
    '                テーブルは、アプリケーションで作成されたレ
    '                コードセットオブジェクトに対するエラーハンドラ
    '                オブジェクトのDestinationプロパティで指定さ
    '                れるものです。
    ' パラメータ: なし
    ' 戻り値     : なし
    ' 作成者    : Total VB SourceBook 6
    '
    Dim intCounter As Integer
    Dim varTemp As Variant

    On Error GoTo PROC_ERR

    If m_varDestination Is Nothing Then
        Beep
        MsgBox "Error handler cannot write to the [" & _
            m_varDestination & _
            "] errors table because the Destination recordset " & _
            "variable has not been initialized.", , m_strAppTitle
    Else
        With m_varDestination
            .AddNew
                ![ErrorDate] = Now
                ![ErrorString] = Left$(m_strErrorDescription, 255)
                ![ErrorNumber] = m_lngErrorNumber
                ![ErrorLine] = m_lngErrorLine
                ![ErrorProc] = mastrErrorStack(mintCurStackPointer)
            
                For intCounter = LBound(mastrErrorStack) To _
                    UBound(mastrErrorStack)
                    If mastrErrorStack(intCounter) <> "" Then
                        varTemp = varTemp & Format(intCounter, "00") & _
                            Space$(intCounter * 2) & _
                            mastrErrorStack(intCounter) & vbCrLf
                    End If
                Next intCounter
            
                If m_strCurrentOperation <> "" Then
                    varTemp = varTemp & vbCrLf & "Current Operation: " & _
                        m_strCurrentOperation
                End If
                
                ![CallStack] = varTemp
        
                If m_fIncludeExpandedInfo Then
                    varTemp = "Current Directory: " & CurDir & vbCrLf
                    varTemp = varTemp & "App Title: " & App.Title & vbCrLf
                    varTemp = varTemp & "App CompanyName: " & App.CompanyName & _
                    vbCrLf & _
                    varTemp = varTemp & "App Major: " & App.Major & vbCrLf
                    varTemp = varTemp & "App Minor: " & App.Minor & vbCrLf
                    varTemp = varTemp & "App Revision: " & App.Revision & vbCrLf
                    varTemp = varTemp & "App Path: " & App.Path & vbCrLf
                    varTemp = varTemp & "App ProductName: " & App.ProductName
                    
                    ' 必要ならば、VB SourceBookが提供する関数を使って、
                    ' Windowsディレクトリ、ディスクの空き容量などのさまざ
                    ' まな項目を取得することができます。

                    ![AdditionalInfo] = varTemp
        
                End If
            .Update
        End With
        
    End If
    
proc_exit:
    Exit Sub

PROC_ERR:
    Err.Raise vbObjectError + 1000, "CErr", "Error opening recordset: "
    Resume proc_exit
    
End Sub



Public Property Get fncName() As String
    fncName = mastrErrorStack(mintCurStackPointer)
End Property

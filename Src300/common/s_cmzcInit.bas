Attribute VB_Name = "s_cmzcInit"
Option Explicit

Public Const SW_SHOW = 5
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public gErr As CErrHandler  'エラーハンドラ(from VB SourceBook)

'概要      :プログラム起動時の初期化処理
'説明      :
Public Function InitExe() As FUNCTION_RETURN
    
    '' プログラム起動時の初期化処理
    DoEvents
    
    '' パラメータ初期化
    InitExe = FUNCTION_RETURN_SUCCESS
   
    '' エラー出力オブジェクト作成
    Init_ErrHandler
    
    ''コマンドライン引数取得
    If GetCmdLine() = False Then
        ''コマンドライン引数無し
        Call MsgOut(64, "", ERR_DISP_LOG)
        Exit Function
    End If
    
       ''実行ファイル名の取得
    If GetEXEName = "" Then
        ''コマンドライン引数無し
        Call MsgOut(62, "", ERR_DISP_LOG)
        Exit Function
    End If
 
    '' 多重起動チェック
    If App.PrevInstance = True Then
        '' 多重起動している場合
        '' エラーメッセージ＆ログ出力
        MsgBox "すでにプログラムが起動されています。", vbOKOnly + vbInformation
        InitExe = FUNCTION_RETURN_FAILURE
        End
    End If
    
    '' データベース接続
    OraDBOpen
    
    '' 処理終了

End Function

Private Sub Init_ErrHandler()
    Set gErr = New CErrHandler
    With gErr
        .AppTitle = App.Title
        .Destination = App.Path & "\Err.log"
        .DisplayMsgOnError = True
        .MaxProcStackItems = 20
        .IncludeExpandedInfo = False
    End With
End Sub

Private Sub TerminateHandler()
    On Error Resume Next
    Set gErr = Nothing
    On Error GoTo 0
End Sub


'///////////////////////////////////////////////////
' @(f)
' 機能    : メインメニューに遷移する
'
' 返り値  :
'
' 引き数  :
'
' 機能説明:
'
'///////////////////////////////////////////////////
Public Sub GotoMainMenu()
    Dim sCallCd As String
    sCallCd = "0000000"
    If gbFTPFlg = True Then             ''FTP起動フラグが立っていたら
        sCallCd = UCase(App.EXENAME)    ''自モジュール名を渡す
    End If
    Call ExitExe(sCallCd) ''メインメニューを起動し、終了
End Sub

'///////////////////////////////////////////////////
' @(f)
' 機能    : サブメニューに遷移する
'
' 返り値  :
'
' 引き数  :
'
' 機能説明:
'
'///////////////////////////////////////////////////
Public Sub GotoSubMenu()
    Dim sCallCd As String
    sCallCd = gsCallCd    ''受け取った呼出区分を渡す
    Call ExitExe(sCallCd) ''サブメニューを起動し、終了
End Sub


'///////////////////////////////////////////////////
' @(f)
' 機能    : 終了入力
'
' 返り値  :
'
' 引き数  : 呼出区分（省略した場合、呼出元メニューを起動する）
'
' 機能説明: 終了入力
'
'///////////////////////////////////////////////////
Public Sub ExitExe(Optional sCallCd As String = "0000000")
    Dim sExeName As String          ''実行ファイル名
'    On Error GoTo Er
    On Error GoTo proc_err
    gErr.Push "s_cmzcCtl.bas -- Function ExitExe"
    sExeName = "XMAIN"
    
    DoEvents                            ''メニューがまだ終了してない可能性があるのでちょっと待つ
   ''メニュー遷移許可なら
 '   If mbMenuRet = True Then
        ''コマンドライン取得 （起動時の工場コードでメニューに戻る）
        gsFactryCd = Left(Command, 2)    ''工場コード(2桁)
        myFactryCd = Mid(Command, 24, 2)   ''工場コード(2桁)
        gsHinban = Mid(Command, 12, 11)
        If Len(gsHinban) <> 11 Then
            gsHinban = "00000000000"
        End If
        
        ''起動
        If 0 = Shell(sExeName & " " & gsFactryCd & " " & sCallCd & " " & gsHinban & " " & myFactryCd, vbNormalFocus) Then
            ''０が戻ってきたら異常
            Call MsgOut(65, sExeName, ERR_DISP_LOG)
        End If
  '  End If
    
   
    ''メニュー遷移許可なら
'    If mbMenuRet = True Then
'        ''起動
'        If 0 = Shell(sExeName & " " & sCallCd, vbNormalFocus) Then
'            ''０が戻ってきたら異常
'            Call MsgOut(65, sExeName, ERR_DISP_LOG)
'
'            WriteDBLog " ", "メニューの起動に失敗しました。"
'        End If
'    End If

    
'Er: On Error Resume Next
'    ''オラクルディスコネクト
'    Call OraDisConn
'    On Error GoTo 0
'    End ''終了

proc_exit:
    gErr.Pop

    '' データベース接続終了
    OraDBClose
    '' エラー出力オブジェクト破棄
    TerminateHandler
    

    End ''終了
    
    
proc_err:
    gErr.HandleError
    Resume proc_exit


End Sub


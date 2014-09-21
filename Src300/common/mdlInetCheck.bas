Attribute VB_Name = "mdlInetCheck"
Option Explicit
'///////////////////////////////////////////////////
' @(S)
'       Inetコントロール使用時の多重起動チェック処理
'
' @(h)  mdlGetFile.bas ver 1.0      ( 2004.12.02 窪田　拓 )
'
'///////////////////////////////////////////////////


'クラス名又はキャプション名を与えてウインドウのハンドルを取得
Private Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long


'多重起動チェックを続ける時間(秒)   この秒数を待って終了しない場合、エラーで終了
Private Const MULTIBOOT_CHECKTIME As Long = 7


'**********************************************************************
' @(f)
'
' 機能　　 : 同一ウインドウ名起動の判定
'
' 返り値　 : True  … 同一ウインドウ名が存在する
' 　　　　   False … 同一ウインドウ名が存在しない
'
' 引き数　 : 調査したいウインドウ名
'
' 機能説明 : 同一ウインドウ名起動の判定
'
' 備考　　 :
'**********************************************************************
Private Function CheckWindowName(ByVal strWindowName As String) As Boolean
    
    Dim lnghwnd As Long
    
    'ウインドウ名を与えてハンドルを取得する
    lnghwnd = FindWindow(vbNullString, strWindowName)
    If lnghwnd = 0 Then
        '同一名のウインドウなし
        CheckWindowName = False
    Else
        '同一名のウインドウが開いている
        CheckWindowName = True
    End If

End Function


'**********************************************************************
' @(f)
'
' 機能　　 : 多重起動の判定
'
' 返り値　 : True  … 多重起動していない(正常)
' 　　　　   False … 多重起動している　(異常)
'
' 引き数　 : 起動フォーム
'
' 機能説明 : 多重起動の判定
'
' 備考　　 :
'**********************************************************************
Private Function CheckMultipleBoot(ByRef frmMain As Form) As Boolean
    
    Dim strWindowName   As String       '' exeウインドウ名
    Dim strFormCaption  As String       '' フォームのウインドウ名
    Dim lngCount        As Long         '' ループカウンタ
    Dim blnResult       As Boolean      '' ウインドウ名チェック結果

    '' 返り値初期化
    CheckMultipleBoot = False

    '' ウインドウ名を保持＆変更
    strWindowName = App.Title
    App.Title = App.Title & "_Check"
    strFormCaption = frmMain.Caption
    frmMain.Caption = frmMain.Caption & " 多重起動チェック中"
    
    '' チェック中メッセージ表示
    Call MsgOut(0, "多重起動チェック中です", NORMAL_MSG)
    
    '' フォームのキャプション名でチェック(窓が残っているかチェック)
    If CheckWindowName(strFormCaption) = True Then
        '' 多重起動(異常終了)
        Exit Function
    End If
    
    '' 前回のexeが終了するまで待つ
    For lngCount = 1 To MULTIBOOT_CHECKTIME
        
        '' 1秒待つ
        Sleep (1000)
        
        '' 再度多重起動のチェック
        blnResult = CheckWindowName(strWindowName)
    
        If blnResult = False Then
            '' 終了していたらループから抜ける
            Exit For
        End If
    
    Next lngCount
    
    '' 終了したかどうか判定
    If blnResult = True Then
        '' 多重起動とみなす(異常終了)
        Exit Function
    End If

    'ウインドウ名を元に戻す
    App.Title = strWindowName
    frmMain.Caption = strFormCaption

    '' チェック中メッセージクリア
    Call MsgOut(0, "", NORMAL_MSG)

    '' 多重起動なし(正常終了)
    CheckMultipleBoot = True

End Function


'**********************************************************************
' @(f)
'
' 機能　　 : プログラム起動時の初期化処理(Inetコントロール使用画面用)
'
' 返り値　 : なし
'
' 引き数　 : 起動フォーム
'
' 機能説明 : Inetコントロール使用画面用のInitExe
'
' 備考　　 : Inetコントロール使用時にexeの終了に時間がかかり、二重起動エラーになる件の対応版
'**********************************************************************
Public Function InitExe_Inet(ByRef frmMain As Form) As Integer
    
    DoEvents
    
    ''初期処理失敗：メインメニュー起動指定
    InitExe_Inet = MAINMENU_RET
    mbMenuRet = False       ''メニュー遷移不許可
    
    ''ログ初期化
    If LogInit() = False Then
        ''ログ初期化失敗
        Call MsgOut(61, "", ERR_DISP_LOG)
        Exit Function
    End If
    
    ''コマンドライン引数取得
    If GetCmdLine_Hikiage() = False Then
        ''コマンドライン引数無し
        Call MsgOut(64, "", ERR_DISP_LOG)
        Exit Function
    End If
    
    ''実行ファイル名取得
    If GetEXEName = "" Then
        ''実行ファイル名取得失敗
        Call MsgOut(62, "", ERR_DISP_LOG)
        Exit Function
    End If
    
    
'↓ InitExe() との相違点  *****************************************
    
    '' 多重起動チェック
    If App.PrevInstance = True Then
        
        '' Inetコントロール使用の場合、前回のexeが終了していない可能性がある為、違う方法で再度チェック
        If CheckMultipleBoot(frmMain) = False Then
            ''多重起動した
            Call MsgOut(63, "", ERR_DISP_LOG)
            Exit Function
        End If
        
    End If
    
'↑ InitExe() との相違点  *****************************************
    
    
    ''初期処理失敗：F1以外操作不可指定
    InitExe_Inet = EXITSUB_RET
    mbMenuRet = True       ''メニュー遷移許可
    
    ''オラクル接続
    If OraConn() = False Then
        ''オラクル接続エラー
        Call MsgOut(100, "", ERR_DISP_LOG)
        Call CtrlCancel(Screen.ActiveForm)      ''ﾒｲﾝﾒﾆｭｰ以外のｺﾝﾄﾛｰﾙを使えなくする
        Exit Function
    End If
    
    ''初期処理完了
    InitExe_Inet = NORMAL_RET
    
End Function


'**********************************************************************
' @(f)
'
' 機能　　 : プログラム起動時の初期化処理(Inetコントロール使用画面用)
'
' 返り値　 : なし
'
' 引き数　 : 起動フォーム
'
' 機能説明 : Inetコントロール使用画面用のInitExe
'
' 備考　　 : Inetコントロール使用時にexeの終了に時間がかかり、二重起動エラーになる件の対応版
'**********************************************************************
Public Function InitExe_Re_Inet(ByRef frmMain As Form) As Integer
    
    DoEvents
    
    ''初期処理失敗：メインメニュー起動指定
    InitExe_Re_Inet = MAINMENU_RET
    mbMenuRet = False       ''メニュー遷移不許可
    
    ''ログ初期化
    If LogInit() = False Then
        ''ログ初期化失敗
        Call MsgOut(61, "", ERR_DISP_LOG)
        Exit Function
    End If
    
    ''コマンドライン引数取得
    If GetCmdLine_Re() = False Then
        ''コマンドライン引数無し
        Call MsgOut(64, "", ERR_DISP_LOG)
        Exit Function
    End If
    
    ''実行ファイル名取得
    If GetEXEName = "" Then
        ''実行ファイル名取得失敗
        Call MsgOut(62, "", ERR_DISP_LOG)
        Exit Function
    End If
    
    
'↓ InitExe() との相違点  *****************************************
    
    '' 多重起動チェック
    If App.PrevInstance = True Then
        
        '' Inetコントロール使用の場合、前回のexeが終了していない可能性がある為、違う方法で再度チェック
        If CheckMultipleBoot(frmMain) = False Then
            ''多重起動した
            Call MsgOut(63, "", ERR_DISP_LOG)
            Exit Function
        End If
        
    End If
    
'↑ InitExe() との相違点  *****************************************
    
    
    ''初期処理失敗：F1以外操作不可指定
    InitExe_Re_Inet = EXITSUB_RET
    mbMenuRet = True       ''メニュー遷移許可
    
    ''オラクル接続
    If OraConn() = False Then
        ''オラクル接続エラー
        Call MsgOut(100, "", ERR_DISP_LOG)
        Call CtrlCancel(Screen.ActiveForm)      ''ﾒｲﾝﾒﾆｭｰ以外のｺﾝﾄﾛｰﾙを使えなくする
        Exit Function
    End If
    
    ''初期処理完了
    InitExe_Re_Inet = NORMAL_RET
    
End Function


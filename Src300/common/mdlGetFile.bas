Attribute VB_Name = "mdlGetFile"
Option Explicit
'///////////////////////////////////////////////////
' @(S)
'       ファイル取得処理
'
' @(h)  mdlGetFile.bas ver 1.0      ( 2004.12.02 窪田　拓 )
'
'///////////////////////////////////////////////////

''定数-------------------------------------------

''バージョン処理関係
Const FtpTimeOut = 20               ''タイムアウト
Const ExtBak = "bk"                 ''拡張子bak

Const INIFILENAME = "DownLoad2.ini" ''前回ダウンロード日付INIファイル
Const INIFILESECTION = "OTHER"      ''前回ダウンロード日付INIファイルセクション

''定義
''--------------API--------------------------------
''INIファイル関係
Declare Function GetPrivateProfileString Lib "kernel32" _
     Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
     ByVal lpKeyName As Any, _
     ByVal lpDefault As String, _
     ByVal lpReturnedString As String, _
     ByVal nSize As Long, _
     ByVal lpFileName As String) _
     As Long                        ''INIファイル読込み

Declare Function WritePrivateProfileString Lib "kernel32" _
     Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
     ByVal lpKeyName As Any, _
     ByVal lpString As Any, _
     ByVal lpFileName As String) _
     As Long                        ''INIファイル書込み

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


' ファイルに関するバージョン情報を取得する関数
Public Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" ( _
    ByVal lptstrFilename As String, _
    lpdwHandle As Long _
    ) As Long

' ファイルに関するバージョン情報を取得する関数
Public Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" ( _
    ByVal lptstrFilename As String, _
    ByVal dwHandle As Long, _
    ByVal dwLen As Long, _
    lpData As Any _
    ) As Long

' バージョン情報リソースから選択されたバージョン情報を取得する関数
Public Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" ( _
    pBlock As Any, _
    ByVal lpSubBlock As String, _
    lplpBuffer As Any, _
    puLen As Long _
    ) As Long
    
' ある位置から別の位置にメモリブロックを移動する関数
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    ByVal Souce As Long, _
    ByVal Length As Long _
    )
    

Public Type VS_FIXEDFILEINFO
    dwSignature         As Long
    dwStrucVersion      As Long
    dwFileVersionMS     As Long
    dwFileVersionLS     As Long
    dwProductVersionMS  As Long
    dwProductVersionLS  As Long
    deFileFlagsMask     As Long
    dwFileFlags         As Long
    dwFileOS            As Long
    dwFileType          As Long
    dwFileDateMS        As Long
    dwFileDateLS        As Long
End Type


'///////////////////////////////////////////////////
' @(f)
' 機能    :INIファイル取得
' 返り値  : 文字列
' 引き数  : ARG1 - セクション名
'           ARG2 - キー名
'           ARG2 - ファイル名
' 機能説明:INIファイル取得
'///////////////////////////////////////////////////
Function GetIni(sec As String, Key As String) As String
    Dim strbuf As String * 256
    Dim strLen As Long
    Dim sIniName As String
    
    sIniName = App.Path & "\" & INIFILENAME
    
    strbuf = ""
    strLen = GetPrivateProfileString(sec, Key, "", strbuf, 256, sIniName)
    GetIni = Left(strbuf, strLen)
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    :INIファイル書き込み
'
' 返り値  : 正否
'
' 引き数  : ARG1 - セクション名
'           ARG2 - キー名
' 機能説明:INIファイル書き込み
'
'///////////////////////////////////////////////////
Function SetIni(sec As String, Key As String) As Boolean
    
    Dim sData As String
    Dim sIniName As String
    
    sData = """" & Format$(Now(), "yyyy/mm/dd hh:nn:ss") & """"
    sIniName = App.Path & "\" & INIFILENAME
    
    SetIni = WritePrivateProfileString(sec, Key, sData, sIniName)

End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : ダウンロード処理
' 返り値  : True  - 正常終了
' 　　　　  False - 異常終了
' 引き数  : sDownLoadFile - ダウンロードファイル名
' 　　　　  sExt          - 拡張子
' 　　　　  frmMain       - 呼出し元フォーム
' 　　　　  objInet       - Inetコントロール
' 機能説明: ダウンロード処理
'///////////////////////////////////////////////////
Public Function ActDownLoad(ByVal sDownLoadFile As String _
                          , ByVal sExt As String _
                          , ByRef frmMain As Form _
                          , ByRef objInet As Inet _
                          ) As Boolean
    
    Dim sLastDLDay      As String   '前回ダウンロード日付
    Dim bDownloadRes    As Boolean  'ダウンロード結果
    Dim bCheckRes       As Boolean  'ダウンロードチェック結果
    Dim iResult         As Integer
    
    ActDownLoad = False
    
    ''ダウンロードの要不要を判定
    If Dir(App.Path & "\" & sDownLoadFile & sExt) = "" Then
        ''ファイルが無ければダウロードが必要
        bCheckRes = True
    Else
'>>>>> exeダウンロード対応 2008/07/01 SETsw kubota -------------
'        ''前回ダウンロード日付のINIファイルを参照する
'        sLastDLDay = GetIni(INIFILESECTION, sDownLoadFile)
'        ''前回ダウンロード日付以降に更新されているかを判定
'        If ChkDownload(sLastDLDay, sDownLoadFile, bCheckRes) = False Then
'            Exit Function
'        End If
        If UCase(sExt) = ".EXE" Then
            If ChkDownload_EXE(sDownLoadFile, bCheckRes, sExt) = False Then
                Exit Function
            End If
        Else
            ''前回ダウンロード日付のINIファイルを参照する
            sLastDLDay = GetIni(INIFILESECTION, sDownLoadFile)
            ''前回ダウンロード日付以降に更新されているかを判定
            If ChkDownload(sLastDLDay, sDownLoadFile, bCheckRes) = False Then
                Exit Function
            End If
        End If
'<<<<< exeダウンロード対応 2008/07/01 SETsw kubota -------------
    
    End If
    
    If bCheckRes = False Then
        'ダウンロードする必要がない場合、正常終了
        ActDownLoad = True
        Exit Function
    End If
    
    ''メッセージ表示
    Call MsgOut(0, "ダウンロード開始", NORMAL_MSG)
    
    ''拡張子をバックアップ用に変更する
    Call ReNameFiles(sDownLoadFile, sExt, sExt & ExtBak)
        
    ''ＦＴＰダウンロード
    iResult = FtpGetFiles(objInet, sDownLoadFile & sExt, bDownloadRes)
        
    ''ダウンロード後ファイル処理
    ''  失敗：バックアップファイル復帰
    ''  成功：バックアップファイル削除
    Call ReNameOrDeleteFiles(sDownLoadFile, bDownloadRes, sExt & ExtBak, sExt)
        
    ''ダウンロード失敗なら
    If iResult < 0 Then
        Call MsgOut(0, "ダウンロード失敗", ERR_DISP)
        Exit Function
    ''ダウンロード成功なら
    Else
        If UCase(sExt) <> ".EXE" Then
            ''前回ダウンロード日付INIファイル書込み
            If SetIni(INIFILESECTION, sDownLoadFile) = False Then
                Call MsgOut(0, "ダウンロード日付INIファイル書込み失敗", ERR_DISP)
                Exit Function
            End If
        End If
    End If
        
    ActDownLoad = True

End Function


'////////////////////////////////////////////////////
' @(f)
' 機能    : 前回ダウンロード日付以降に更新されているかを判定
'
' 返り値  : -1:失敗
'          >=0:取得件数
'
' 引き数  : sLastDLDay - 前回ダウンロード日付
' 　　　　  sFileName  - ダウンロード対象モジュール名
' 　　　　  bDownFlg   - ダウンロード要不要フラグ   True - 要  False - 不要
'
' 機能説明: 前回ダウンロード日付以降に更新されているかを判定
'///////////////////////////////////////////////////
Function ChkDownload(ByVal sLastDLDay As String _
                   , ByVal sFileName As String _
                   , ByRef bDownFlg As Boolean _
                   ) As Boolean
    
    Dim sSQL As String
    Dim objOraDyn As Object
    
    ChkDownload = False
    bDownFlg = False
    
    sSQL = "       SELECT codea9                    "   ''ロードモジュール名
    sSQL = sSQL & "FROM   koda9                     "   ''管理コードテーブル
    sSQL = sSQL & "WHERE  sysca9 = 'K'              "   ''
    sSQL = sSQL & "AND    shuca9 = '01'             "   ''バージョン情報
    sSQL = sSQL & "AND    codea9 = '" & sFileName & "' "
    If Trim$(sLastDLDay) <> "" Then ''日付が指定されたら条件に入れる
        sSQL = sSQL & "AND    tdaya9 > TO_DATE(        '" _
                & Format(sLastDLDay, "yyyymmddhhnnss") _
                & "','yyyymmddhh24miss')            "         ''登録日付が前回ダウンロード日付より新しいもの
    End If
    
    ''ダイナセット作成
'>>>>> 300mmDynSet2対応　2008/11/21　SET.Marushita
    'If DynSet(objOraDyn, sSQL) = False Then
    If DynSet2(objOraDyn, sSQL) = False Then
'<<<<< 300mmDynSet2対応　2008/11/21　SET.Marushita
        ''ダイナセット作成失敗
        Call MsgOut(100, sSQL, ERR_DISP_LOG, "kodea9")
        Exit Function
    End If
    
    If objOraDyn.EOF = False Then
        bDownFlg = True
    End If
    
    ChkDownload = True
    
End Function

'////////////////////////////////////////////////////
' @(f)
' 機能    : ＤＢとローカルファイルのバージョンを比較しダウンロード要不要を判定
'
' 返り値  : False - 異常
'           True  - 正常
'
' 引き数  : sFileName  - ダウンロード対象モジュール名
' 　　　　  bDownFlg   - ダウンロード要不要フラグ   True - 要  False - 不要
'
' 機能説明:
'///////////////////////////////////////////////////
Function ChkDownload_EXE(ByVal sFileName As String _
                       , ByRef bDownFlg As Boolean _
                       , ByVal sExt As String _
                       ) As Boolean
    
    Dim sSQL        As String
    Dim objOraDyn   As Object
    Dim sMajor      As String
    Dim sMinor      As String
    Dim sRevision   As String
    Dim tKoda9      As typKoda9Data

    bDownFlg = False
    
    '対象ファイルのローカルファイルVer取得
    If GetFileVer(sFileName & sExt, sMajor, sMinor, sRevision) = False Then
        Exit Function
    End If
    
    '対象ファイルのＤＢ登録Ver取得
    If GetKanriCode("K", "01", sFileName, tKoda9) = False Then
        Exit Function
    End If

    'ﾒｼﾞｬｰ、ﾏｲﾅｰ、ﾘﾋﾞｼﾞｮﾝを比較して一つでも違う場合、ダウンロード要
    If val(sMajor) <> val(tKoda9.sCTR01A9) _
    Or val(sMinor) <> val(tKoda9.sCTR02A9) _
    Or val(sRevision) <> val(tKoda9.sCTR03A9) Then
        bDownFlg = True
    End If
    
    ChkDownload_EXE = True
    
End Function



'///////////////////////////////////////////////////
' @(f)
' 機能    : ＦＴＰダウンロード処理
' 返り値  : -1:異常
'            0:正常
' 引き数  : sFileName - ダウンロードファイル名
' 　　　　  bResult   - ダウンロード結果
' 機能説明: ＦＴＰダウンロード処理
'///////////////////////////////////////////////////
Private Function FtpGetFiles(ByRef objInet As Inet, ByVal sFileName As String, ByRef bResult As Boolean) As Integer
    Dim sHost     As String ''ホスト
    Dim sUserId   As String ''ユーザー
    Dim sPassword As String ''パスワード
    Dim sHostPath As String ''ホストロードモジュールパス
    On Error GoTo Er
    
    bResult = False
    
    Select Case gsFactryCd
    Case "10"               ''野田工場
        sHost = "CLB0"
        sUserId = "oracle"
        sPassword = "oracle"
        sHostPath = "/home2/mist/download/"
    Case "30"               ''生野工場
        sHost = "CLD0"
        sUserId = "oracle"
        sPassword = "oracle"
        sHostPath = "/home2/mist/download/"
    Case "AM"               ''尼崎工場
        sHost = "133.0.0.47"
        sUserId = "oracle"
        sPassword = "oracle"
        sHostPath = "/home2/mist/download/"
    Case "40"               ''米沢工場
        sHost = "CLE0"
        sUserId = "oracle"
        sPassword = "oracle"
        sHostPath = "/home2/mist/download/"
    Case "42"               ''３００ｍｍ
        sHost = "172.20.128.2"
        sUserId = "mqm"
        sPassword = "mqm0001"
        sHostPath = "/home2/cm1/tool/newvb/"
    Case "43"               ''３００ｍｍ試作
        sHost = "172.20.104.24"
        sUserId = "mqm"
        sPassword = "manager"
        sHostPath = "/home2/cm1/tool/newvb/"
    Case "90"               ''テスト
        sHost = "CLA0"
        sUserId = "oracle"
        sPassword = "oracle"
        sHostPath = "/home2/mist/download/"
    Case "91"               ''新テスト
        sHost = "172.20.104.24"
        sUserId = "oracle"
        sPassword = "oracle"
        sHostPath = "/home2/mist/download/"
    Case Else               ''外販
        sHost = "CLB0"
        sUserId = "oracle"
        sPassword = "oracle"
        sHostPath = "/home2/mist/download/"
    End Select
    
    With objInet
        .URL = sHost
        .UserName = sUserId
        .Password = sPassword
        .RequestTimeout = FtpTimeOut
            
        Call MsgOut(0, "ＦＴＰ取得→" & sFileName, DEBUG_DISP_LOG)
        .Execute , "GET " & sHostPath & sFileName _
                  & " """ & App.Path & "\" & sFileName & """"        ''ＦＴＰ取得
        Do While .StillExecuting = True ''終了待ち
            DoEvents
            If .ResponseCode Then   ''エラー
                Exit Do ''ループを抜ける
            End If
        Loop
        If .ResponseCode Then   ''エラー
            bResult = False   ''失敗
            '画面表示
            Call MsgOut(0, "ダウンロード失敗：" & sFileName, DEBUG_DISP_LOG)
            'ログ出力
            Call MsgOut(0, "ＦＴＰエラー:エラーコード(" & .ResponseCode & ")" & _
                                                        .ResponseInfo, ERR_LOG)
            '既にエラー出力済み
            FtpGetFiles = -1
        Else
            bResult = True    ''成功
        End If
        On Error Resume Next
        .Execute , " CLOSE"  '' 接続を閉じる
    End With
    On Error GoTo 0
    
    Call MsgOut(0, "", NORMAL_MSG)
    Exit Function
Er:
    With objInet
        ''画面表示
        Call MsgOut(0, "ダウンロード失敗：" & sFileName, DEBUG_DISP_LOG)
        ''ログ出力
        Call MsgOut(0, "ＦＴＰエラー:エラーコード(" & .ResponseCode & ")" & _
                                                    .ResponseInfo, ERR_LOG)
        FtpGetFiles = -1
    End With
    Resume Next
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : 拡張子変更処理
' 返り値  : -1:異常
'           >0:処理件数
' 引き数  : ファイル名配列
'           変更前拡張子
'           変更後拡張子
' 機能説明: 拡張子変更処理
'///////////////////////////////////////////////////
Private Function ReNameFiles(sFileName, sSExt As String, sDExt As String) As Integer
    On Error GoTo Er
    ''ファイルの存在チェック
    If Dir(App.Path & "\" & sFileName & sSExt) <> "" Then   ''そのファイルが在れば
        ''ファイル名の拡張子を変更
        Name App.Path & "\" & sFileName & sSExt As App.Path & "\" & sFileName & sDExt
        ReNameFiles = ReNameFiles + 1
    End If
    Exit Function
Er:
    Call MsgOut(0, "ﾌｧｲﾙ拡張子変更失敗 " & sFileName & sSExt & "→" & sDExt, ERR_DISP_LOG)
    Resume Next
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : ファイル削除処理
' 返り値  : -1:異常
'           >0:処理件数
' 引き数  : ファイル名
' 機能説明: ファイル削除処理
'///////////////////////////////////////////////////
Private Function DeleteFiles(sFileName)
    On Error GoTo Er
    ''ファイルの存在チェック
    If Dir(App.Path & "\" & sFileName) <> "" Then    ''そのファイルが在れば
        ''削除
        Kill App.Path & "\" & sFileName
        DeleteFiles = DeleteFiles + 1
    End If
    Exit Function
Er:
    Call MsgOut(0, "ﾌｧｲﾙ削除失敗 " & sFileName, ERR_DISP_LOG)
    Resume Next
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : ダウンロード後ファイル処理
' 返り値  :
' 引き数  : ファイル名配列
'           バックアップファイル拡張子
'           実行ファイル拡張子
' 機能説明: ダウンロード結果フラグによりバックアップファイル削除／exe復帰する
' 備考    :     成功：バックアップファイル削除
'               失敗：バックアップファイルexe復帰
'///////////////////////////////////////////////////
Private Sub ReNameOrDeleteFiles(ByVal sFileName As String _
                              , ByVal bDownloadRes As Boolean _
                              , ByVal sBakExt As String _
                              , ByVal sExeExt As String)
    On Error GoTo Er
    ''ファイルの存在チェック
    If Dir(App.Path & "\" & sFileName & sBakExt) <> "" Then    ''そのファイルが在れば
        ''ダウンロード成功なら
        If bDownloadRes = True Then
            ''ファイル削除
            Kill App.Path & "\" & sFileName & sBakExt
        ''ダウンロード失敗なら
        Else
            On Error Resume Next
            ''失敗したダウンロード途中の残骸ファイルを削除
            Kill App.Path & "\" & sFileName & sExeExt
            On Error GoTo 0
            ''ファイル名の拡張子をバックアップからＥＸＥファイルに変更
            Name App.Path & "\" & sFileName & sBakExt As App.Path & "\" & sFileName & sExeExt
        End If
    End If
    Exit Sub
Er:
    Call MsgOut(0, "ﾀﾞｳﾝﾛｰﾄﾞ後ﾌｧｲﾙ処理失敗 " & sFileName, ERR_DISP_LOG)
    Resume Next
End Sub

'///////////////////////////////////////////////////
' @(f)
' 機能    : ファイルバージョン取得
' 返り値  : True  - 正常
' 　　　　  False - 異常
' 引き数  : sFileName  - Ver取得対象ファイル名
' 　　　　  sMajor     - メジャーバージョン
' 　　　　  sMinor     - マイナーバージョン
' 　　　　  sRevision  - 枝番
' 機能説明:
' 備考    : 2008/07/01 EXEダウンロード対応
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetFileVer(ByVal sFileName As String _
                         , ByRef sMajor As String _
                         , ByRef sMinor As String _
                         , ByRef sRevision As String _
                         ) As Boolean
    
    Dim lngSizeOfVersionInfo    As Long
    Dim lngDummyHandle          As Long
    Dim bytDummyVersionInfo()   As Byte
    Dim lngPointerversionInfo   As Long
    Dim lngLengthVersioninfo    As Long
    Dim udtVSFixedFileInfo      As VS_FIXEDFILEINFO
    Dim lngWin32apiResultCode   As Long
    
    lngSizeOfVersionInfo = GetFileVersionInfoSize(sFileName, lngDummyHandle)
    If lngSizeOfVersionInfo > 0 Then
        ReDim bytDummyVersionInfo(lngSizeOfVersionInfo - 1)
        lngWin32apiResultCode = GetFileVersionInfo(sFileName, _
                                                    0, _
                                                    lngSizeOfVersionInfo, _
                                                    bytDummyVersionInfo(0))
        lngWin32apiResultCode = VerQueryValue(bytDummyVersionInfo(0), _
                                            "\", _
                                            lngPointerversionInfo, _
                                            lngLengthVersioninfo)
        
        Call MoveMemory(udtVSFixedFileInfo, lngPointerversionInfo, Len(udtVSFixedFileInfo))
        
        With udtVSFixedFileInfo
            sMajor = (.dwProductVersionMS \ (2 ^ 16)) And &HFFFF&
            sMinor = Format(.dwProductVersionMS And &HFFFF&, "#0")
            sRevision = Format(.dwProductVersionLS)
        End With
    
    Else
        sMajor = ""
        sMinor = ""
        sRevision = ""
        Exit Function
    End If

    GetFileVer = True

End Function





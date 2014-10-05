Attribute VB_Name = "mdlVBX9000"
'///////////////////////////////////////////////////
' @(S)
'       帳票出力依頼処理
'
' @(h)  mdlVBX9000.bas ver 0.1 ( 2000.04.20  )
'
'///////////////////////////////////////////////////
Option Explicit

Public Enum gPRINTCODE  ''  帳票出力指令区分型
    PULCRYSPRINT        ''  引上指示書出力指令
    BUNKATUPRINT        ''  加工検査票出力指令
End Enum
Type PRINTERINFODATA            ''  帳票出力情報テーブル管理データ格納型
    sPrinterName    As String   ''  プリンター名
    sPrintDatName   As String   ''  出力帳票名称
    sInfoDat01      As String   ''  帳票出力キーデータ01
    sInfoDat02      As String   ''  帳票出力キーデータ02
    sInfoDat03      As String   ''  帳票出力キーデータ03
    sInfoDat04      As String   ''  帳票出力キーデータ04
    sInfoDat05      As String   ''  帳票出力キーデータ05
    lMaxCnt         As Long     ''  最大レコード
    lReadCnt        As Long     ''  リードポインター
    lWriteCnt       As Long     ''  ライトポインター
End Type
Type KANRIKODEDAT
    sSysCode    As String       ''  システム区分
    sShuCode    As String       ''  種別コード
    sCode       As String       ''  個別コード
    sShortName  As String       ''  名称(短縮)
    sCodeName   As String       ''  名称
    sKanrenCode As String       ''  関連コード
    sDat01      As String       ''  データ１
    sDat02      As String       ''  データ２
    sDat03      As String       ''  データ３
    sDat04      As String       ''  データ４
    sDat05      As String       ''  データ５
    lCnt01      As Long         ''  カウンター１
    lCnt02      As Long         ''  カウンター２
    lCnt03      As Long         ''  カウンター３
    lCnt04      As Long         ''  カウンター４
    lCnt05      As Long         ''  カウンター５
End Type
Private Const mUNTENNISSI = "CRX1000"
Private Const mHIKIAGESIJISYO = "CRX1010"
Private Const mKAKOUKENSAHYO = "CRX1020"
Private Const mSAIKAKUDUKEUKAGAISYO = "CRX1030"

' @(f)
' 機能      :   帳票出力依頼処理
'
' 返り値    :   True    -   正常
'               False   -   異常
'
' 引き数    :   sWorkCode   -   工場コード
'               pPrintKbn   -   帳票コード
'                               １：運転日誌
'                               ２：引上指示書
'                               ３：加工検査票
'                               ４：再格付伺い指示書
'
' 機能説明  :   帳票出力依頼情報を「依頼情報テーブル」に書き込む
'
' 備考      :   sInfCode(1-4) - sPrintKbn(1)時    引上結晶番号、原料番号、工程連番、号機
'               sInfCode(1-2) - sPrintKbn(2)時    引上指示№、号機
'               sInfCode(1)   - sPrintKbn(3)時    分割結晶番号
'               sInfCode(1-3) - sPrintKbn(4)時    分割結晶番号、転用元品番、転用先品番
'
Public Function SetPrinterManager(sWorkCode As String, _
                                 pPrintKbn As gPRINTCODE, _
                                 Optional sInfCode1 As String = "", _
                                 Optional sInfCode2 As String = "", _
                                 Optional sInfCode3 As String = "", _
                                 Optional sInfCode4 As String = "", _
                                 Optional sInfCode5 As String = "") As Boolean
Dim KanriDat        As KANRIKODEDAT
Dim PrintDat        As PRINTERINFODATA
    
    
    SetPrinterManager = False
    
    PrintDat.sInfoDat01 = sInfCode1
    PrintDat.sInfoDat02 = sInfCode2
    PrintDat.sInfoDat03 = sInfCode3
    PrintDat.sInfoDat04 = sInfCode4
    PrintDat.sInfoDat05 = sInfCode5
    ''  管理コードパラメーター設定
    KanriDat.sSysCode = "X"
    KanriDat.sShuCode = "98"
    KanriDat.sCode = "X"
    ''  管理コードテーブル検索処理
    ''  帳票コードからプリンター名称を取得する
    If SelectKanriTable(KanriDat) = False Then
        Exit Function
    End If
    ''  プリンター名称の取得
    ''  出力指令区分による帳票コードの設定
    Select Case pPrintKbn
    Case PULCRYSPRINT        ''  引上指示書出力指令
        PrintDat.sPrintDatName = "r_cmdc001a"
        PrintDat.sPrinterName = KanriDat.sDat02
    Case BUNKATUPRINT        ''  加工検査票出力指令
        PrintDat.sPrintDatName = "r_cmdc001b"
        PrintDat.sPrinterName = KanriDat.sDat03
    End Select
    ''  管理コードパラメーター設定
    KanriDat.sSysCode = "X"
    KanriDat.sShuCode = "99"
    KanriDat.sCode = "1"
    ''  管理コードテーブル検索処理
    ''  帳票コードから管理データ(ＭＡＸレコード数、読込ポインタ、書き込みポインタ)を取得する
    If SelectKanriTable(KanriDat) = False Then
        Exit Function
    End If
    ''  ＭＡＸレコード、読込ポインタ、書込ポインタ取得
    PrintDat.lMaxCnt = Val(KanriDat.sDat01)
    PrintDat.lReadCnt = Val(KanriDat.sDat02)
    PrintDat.lWriteCnt = Val(KanriDat.sDat03)
    ''  書込ポインターインクリメント
    If PrintDat.lMaxCnt < PrintDat.lWriteCnt + 1 Then
        PrintDat.lWriteCnt = 1
    Else
        PrintDat.lWriteCnt = PrintDat.lWriteCnt + 1
    End If
    ''  読込ポインターと同じ場合はエラーとする
    If PrintDat.lReadCnt = PrintDat.lWriteCnt Then
        Call MsgOut(0, "帳票出力命令エラー", ERR_LOG)
        Exit Function
    End If
    
    ''コンピュータ名取得   1/9 Yam
    If GetCompName() = "" Then
        ''コンピュータ名取得失敗
        Call MsgOut(60, "", ERR_DISP_LOG)
        Debug.Print "コンピュータ名取得失敗"
        Exit Function
    End If
    
    If UpdatePrinterInfo(PrintDat) = False Then
        Exit Function
    End If
    If UpdateKanriTable(2, CStr(PrintDat.lWriteCnt), "") = False Then
        Exit Function
    End If
    
    '2004.09.06 Y.Katabami
    '指示書発行ＦＬＧ、運転日誌再発行ＦＬＧの更新を行う。
    '今回出力要求したＦＬＧを "2" 発行済みとする
    If funcPrintOutFlgUpdata(pPrintKbn, sInfCode1) = False Then
        Exit Function
    End If
    
    SetPrinterManager = True
End Function

'出力登録完了情報をTBCMH001の各出力フラグを２（出力済み）に変更する
'2004.09.06 Y.Katabami
Function funcPrintOutFlgUpdata(pPrintKbn, sInfCode1) As Boolean
    Dim sSQL        As String
    Dim lRowCount   As Long
    Dim sSijiNo As String
    
    funcPrintOutFlgUpdata = False
        
    If (pPrintKbn = PULCRYSPRINT) Then        ''  引上指示書出力指令
        'sInfCode1は指示Ｎｏなのでそのまま紐付け
        sSQL = "UPDATE  TBCMH001  SET "
        sSQL = sSQL & " SIJISYOFLG = '2' "
        sSQL = sSQL & " WHERE UPINDNO = '" & sInfCode1 & "'"
        lRowCount = SqlExec2(sSQL)
        If lRowCount < 0 Then
            Call MsgOut(100, sSQL, ERR_LOG, "koda9")
            Exit Function
        End If
    ElseIf (pPrintKbn = BUNKATUPRINT) Then       ''  加工検査票出力指令
        'sInfCode1はチャージＮｏから指示Ｎｏを作成し紐付け
        sSijiNo = Mid(sInfCode1, 1, 7) & "0" & Mid(sInfCode1, 9, 1)
        
        sSQL = "UPDATE  TBCMH001  SET "
        sSQL = sSQL & " UNTENFLG = '2' "
        sSQL = sSQL & " WHERE UPINDNO = '" & sSijiNo & "'"
        lRowCount = SqlExec2(sSQL)
        If lRowCount < 0 Then
            Call MsgOut(100, sSQL, ERR_LOG, "koda9")
            Exit Function
        End If
    End If
    
    funcPrintOutFlgUpdata = True

End Function

' @(f)
' 機能      :   帳票依頼情報更新処理
' 返り値    :   True    -   正常
'               False   -   異常
' 引き数    :   PrintDat    -   帳票出力情報格納変数
' 機能説明  :   帳票出力依頼情報を「依頼情報テーブル」に書き込む
' 備考      :   更新結果0件の場合は登録を行う
'
Private Function UpdatePrinterInfo(PrintDat As PRINTERINFODATA) As Boolean
    Dim sSQL        As String
    Dim lRowCount   As Long
    UpdatePrinterInfo = False
    sSQL = "UPDATE  kodz5   "
    sSQL = sSQL & " SET crclientz5  =   '" & gsCompName & "'            , "
    sSQL = sSQL & "     crcodez5    =   '" & PrintDat.sPrintDatName & "', "
    sSQL = sSQL & "     crprintz5   =   NULL                            , "
    sSQL = sSQL & "     crymdz5     =   sysdate                         , "
    sSQL = sSQL & "     sdayz5     =   sysdate                         , "
    sSQL = sSQL & "     sndkz5     =   ' '                         , "
    If PrintDat.sInfoDat01 = "" Then
        sSQL = sSQL & " crkey01z5   = NULL                              , "
    Else
        sSQL = sSQL & " crkey01z5   = '" & PrintDat.sInfoDat01 & "'     , "
    End If
    If PrintDat.sInfoDat02 = "" Then
        sSQL = sSQL & " crkey02z5   = NULL                              , "
    Else
        sSQL = sSQL & " crkey02z5   = '" & PrintDat.sInfoDat02 & "'     , "
    End If
    If PrintDat.sInfoDat03 = "" Then
        sSQL = sSQL & " crkey03z5   = NULL                              , "
    Else
        sSQL = sSQL & " crkey03z5   = '" & PrintDat.sInfoDat03 & "'     , "
    End If
    If PrintDat.sInfoDat04 = "" Then
        sSQL = sSQL & " crkey04z5   = NULL                              , "
    Else
        sSQL = sSQL & " crkey04z5   = '" & PrintDat.sInfoDat04 & "'     , "
    End If
    If PrintDat.sInfoDat05 = "" Then
        sSQL = sSQL & " crkey05z5   = NULL                                "
    Else
        sSQL = sSQL & " crkey05z5   = '" & PrintDat.sInfoDat05 & "'       "
    End If
    sSQL = sSQL & "WHERE    crseqz5 = " & CStr(PrintDat.lWriteCnt)
    lRowCount = SqlExec2(sSQL)
    If lRowCount < 0 Then
        Call MsgOut(100, sSQL, ERR_LOG, "koda9")
        Exit Function
    ElseIf lRowCount = 0 Then
        sSQL = "INSERT INTO kodz5"
        sSQL = sSQL & "( crseqz5, crclientz5, crcodez5, croutz5, crprintz5, crymdz5,      "
        sSQL = sSQL & " sdayz5, sndkZ5, crkey01z5, crkey02z5, crkey03z5, crkey04z5, crkey05z5   )"
        
        sSQL = sSQL & " VALUES "
        sSQL = sSQL & "(" & CStr(PrintDat.lWriteCnt) & ","
        sSQL = sSQL & "'" & gsCompName & "'           ,"
        sSQL = sSQL & "'" & PrintDat.sPrintDatName & "', "
        sSQL = sSQL & " '1'                           ,"
        sSQL = sSQL & "NULL                           ,"
        sSQL = sSQL & "     sysdate,"
        sSQL = sSQL & "     sysdate,"
        sSQL = sSQL & "     ' ',"
        If PrintDat.sInfoDat01 = "" Then
            sSQL = sSQL & " NULL,"
        Else
            sSQL = sSQL & " '" & PrintDat.sInfoDat01 & "',"
        End If
        If PrintDat.sInfoDat02 = "" Then
            sSQL = sSQL & " NULL,"
        Else
            sSQL = sSQL & " '" & PrintDat.sInfoDat02 & "',"
        End If
        If PrintDat.sInfoDat03 = "" Then
            sSQL = sSQL & "  NULL,"
        Else
            sSQL = sSQL & "  '" & PrintDat.sInfoDat03 & "',"
        End If
        If PrintDat.sInfoDat04 = "" Then
            sSQL = sSQL & "  NULL,"
        Else
            sSQL = sSQL & "  '" & PrintDat.sInfoDat04 & "',"
        End If
        If PrintDat.sInfoDat05 = "" Then
            sSQL = sSQL & "  NULL)"
        Else
            sSQL = sSQL & "  '" & PrintDat.sInfoDat05 & "')"
        End If
        lRowCount = SqlExec2(sSQL)
        If lRowCount <> 1 Then
            Call MsgOut(100, sSQL, ERR_LOG, "kodz5")
            Exit Function
        End If
    End If
    UpdatePrinterInfo = True
End Function

' @(f)
' 機能      :   管理コードＴＢＬ読出し処理
' 返り値    :   TRUE    -   正常
'               FALSE   -   異常
' 引き数    :   KanriDat    -   管理コードデータ格納データ
'
' 機能説明  :   管理コードテーブルの指定されたシステム区分、種別コード、個別コードでデータを検索する
' 備考      :
' 修正      :   2000.04.27  管理コードテーブルの帳票処理連番レコードは他の帳票出力モジュールからも同時に
'                           参照する可能性があるので処理連番が重複しないようにレコードをロックする
'
Public Function SelectKanriTable(KanriDat As KANRIKODEDAT) As Boolean
    Dim sSQL        As String
    Dim objDS       As Object
    SelectKanriTable = False
    sSQL = "SELECT  NVL(kcode01a9,' '), "
    sSQL = sSQL & " NVL(kcode02a9,' '), "
    sSQL = sSQL & " NVL(kcode03a9,' '), "
    sSQL = sSQL & " NVL(kcode04a9,' '), "
    sSQL = sSQL & " NVL(kcode05a9,' '), "
    sSQL = sSQL & " NVL(ctr01a9,0),     "
    sSQL = sSQL & " NVL(ctr02a9,0),     "
    sSQL = sSQL & " NVL(ctr03a9,0),     "
    sSQL = sSQL & " NVL(ctr04a9,0),     "
    sSQL = sSQL & " NVL(ctr05a9,0)      "
    sSQL = sSQL & "FROM koda9           "
    sSQL = sSQL & "WHERE    codea9  =   '" & KanriDat.sCode & "'"
    sSQL = sSQL & "  AND    shuca9  =   '" & KanriDat.sShuCode & "'"
    sSQL = sSQL & "  AND    sysca9  =   '" & KanriDat.sSysCode & "'"
    sSQL = sSQL & " FOR UPDATE           "   ''  管理コードテーブルレコードロック
    If DynSet2(objDS, sSQL) = False Then
        Call MsgOut(100, sSQL, ERR_LOG, "koda9")
        Exit Function
    End If
    If objDS.EOF = False Then
        Do Until objDS.EOF
            KanriDat.sDat01 = objDS(0)
            KanriDat.sDat02 = objDS(1)
            KanriDat.sDat03 = objDS(2)
            KanriDat.sDat04 = objDS(3)
            KanriDat.sDat05 = objDS(4)
            KanriDat.lCnt01 = objDS(5)
            KanriDat.lCnt02 = objDS(6)
            KanriDat.lCnt03 = objDS(7)
            KanriDat.lCnt04 = objDS(8)
            KanriDat.lCnt05 = objDS(9)
            objDS.MoveNext
        Loop
    Else
        Call MsgOut(55, sSQL, ERR_LOG, "koda9")
    End If
    SelectKanriTable = True
    Set objDS = Nothing
End Function

' @(f)
' 機能      :   管理コードＴＢＬ更新（RWポインタ更新）
' 返り値    :   1     :正常
'               1以外：異常
'
' 引き数    :   iUpFlg          -   更新ポインターフラグ
'               sWritePointer   -   新ﾗｲﾄﾎﾟｲﾝﾀ
'               sReadPointer    -   新ﾘｰﾄﾞﾎﾟｲﾝﾀ
'
' 機能説明  : 依頼情報ﾃｰﾌﾞﾙで使用している処理連番をインクリメントする。
'
' 備考      :   iUpFlg      ０－＞RWポインタ更新
'                           １－＞Rポインタのみ更新
'                           ２－＞Wポインタのみ更新
'
Public Function UpdateKanriTable(iUpFlg As Integer, sWritePointer As String, sReadPointer As String) As Boolean
    Dim sSQL        As String
    Dim lRowCount   As Long
    UpdateKanriTable = False
    sSQL = "UPDATE  koda9"
    sSQL = sSQL & " SET "
    Select Case iUpFlg
    Case 0  ''  RWポインタ更新
        sSQL = sSQL & " kcode02a9 = " & sReadPointer & ","
        sSQL = sSQL & " kcode03a9 = " & sWritePointer & ","
    Case 1  ''  Rポインタのみ更新
        sSQL = sSQL & " kcode02a9 = " & sReadPointer & ","
    Case 2  ''  Wポインタのみ更新
        sSQL = sSQL & " kcode03a9 = " & sWritePointer & ","
    End Select
    sSQL = sSQL & " sdaya9 =  sysdate,"
    sSQL = sSQL & " sndka9 =    ' '"
    sSQL = sSQL & " WHERE   codea9  =   '1' "
    sSQL = sSQL & "   AND   shuca9  =   '99'"
    sSQL = sSQL & "   AND   sysca9  =   'X' "
    'SQL実行
    lRowCount = SqlExec2(sSQL)
    Select Case lRowCount
    Case -1
        Call MsgOut(100, sSQL, ERR_LOG, "koda9")
        Exit Function
    Case 0
        Call MsgOut(72, "koda9", ERR_LOG)
        Exit Function
    Case 1
    Case Else
        Call MsgOut(0, "複数件更新した:" & sSQL, ERR_LOG)
        Exit Function
    End Select
    UpdateKanriTable = True
End Function

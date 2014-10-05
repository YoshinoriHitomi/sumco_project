Attribute VB_Name = "s_cmbc039_Wfmap"
Option Explicit

'表示条件
Public Const intConSprChg_0   As Integer = 0    '全体
Public Const intConSprChg_1   As Integer = 1    '良品
Public Const intConSprChg_2   As Integer = 2    'サンプル
Public Const intConSprChg_3   As Integer = 3    '不良

Public wfmapview As Boolean     ''使ってない
Public sampleon As Boolean      ''使ってない

Public Type SprData
    joutai As String
    SXLID As String
    LOTID As String
    hinban  As String
    BLOCKSEQ As String
    blockp  As String
    wfnum   As String
    flg     As String
    keturakucd  As String
    keturakuriyu    As String
    kotei   As String
    update  As String
    KANKBN  As String
    nukishishiji    As String
    nukishikekka    As String
    shijireturn     As String
    kasetto      As String
End Type

'*******************************************************************************
'*    関数名        : SelWFmap
'*
'*    処理概要      : 1.WFﾏｯﾌﾟ管理ﾃｰﾌﾞﾙ（TBCMY011）からﾃﾞｰﾀを取得
'*
'*    パラメータ    : 変数名        ,IO ,型           ,説明
'*                   records()      ,O  ,typ_TBCME037 ,抽出レコード
'*                   sqlWhere       ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'*                   sqlOrder       ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function SelWFmap(ByVal sBlkId As String, ByVal sSXLID As String, ByRef sErrMsg As String) As FUNCTION_RETURN

    Dim sSQL        As String
    Dim blBlkOn     As Boolean          ' ﾌﾞﾛｯｸが存在していたら、SXLIDの条件にANDを入れる
    Dim rs          As OraDynaset       ' RecordSet
    Dim intDataCnt  As Integer

    On Error GoTo PROC_ERR

    SelWFmap = FUNCTION_RETURN_FAILURE
    blBlkOn = False
    intDataCnt = 0
    
    sSQL = vbNullString
    sSQL = sSQL & "SELECT * FROM TBCMY011"
    sSQL = sSQL & " WHERE"
    
    'NULLでなければ条件に使用
    If sBlkId <> vbNullString Then
        sSQL = sSQL & " LOTID = '" & sBlkId & "'"
        blBlkOn = True
    End If
    If sSXLID <> vbNullString Then
        If blBlkOn = True Then
            sSQL = sSQL & " AND MSXLID = '" & sSXLID & "'"
        Else
            sSQL = sSQL & " MSXLID = '" & sSXLID & "'"
        End If
    End If
    sSQL = sSQL & " ORDER BY LOTID, BLOCKSEQ"

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        SelWFmap = FUNCTION_RETURN_FAILURE
        sErrMsg = GetMsgStr("SET46")
        Exit Function
    End If
    
    Do While Not rs.EOF
        ReDim Preserve gtWFmap(intDataCnt) As typeWFmap
        With gtWFmap(intDataCnt)
            .LOTID = CStr(rs.Fields("LOTID"))
            If IsNull(rs.Fields("BLOCKSEQ")) = True Then
                .BLOCKSEQ = 0
            Else
                .BLOCKSEQ = CInt(rs.Fields("BLOCKSEQ"))
            End If
            If IsNull(rs.Fields("INDTM")) = True Then
                .INDTM = vbNullString
            Else
                .INDTM = CStr(rs.Fields("INDTM"))
            End If
            If IsNull(rs.Fields("BASKETID")) = True Then
                .BASKETID = vbNullString
            Else
                .BASKETID = CStr(rs.Fields("BASKETID"))
            End If
            If IsNull(rs.Fields("SLOTNO")) = True Then
                .SLOTNO = 0
            Else
                .SLOTNO = CInt(rs.Fields("SLOTNO"))
            End If
            If IsNull(rs.Fields("CURRWPCS")) = True Then
                .CURRWPCS = 0
            Else
                .CURRWPCS = CInt(rs.Fields("CURRWPCS"))
            End If
            If IsNull(rs.Fields("EXISTFLG")) = True Then
                .EXISTFLG = vbNullString
            Else
                .EXISTFLG = CStr(rs.Fields("EXISTFLG"))
            End If
            If IsNull(rs.Fields("TOP_POS")) = True Then
                .TOP_POS = 0
            Else
                .TOP_POS = rs.Fields("TOP_POS")
            End If
            If IsNull(rs.Fields("REJCAT")) = True Then
                .REJCAT = vbNullString
            Else
                .REJCAT = CStr(rs.Fields("REJCAT"))
            End If
            If IsNull(rs.Fields("TXID")) = True Then
                .TXID = vbNullString
            Else
                .TXID = CStr(rs.Fields("TXID"))
            End If
            If IsNull(rs.Fields("REGDATE")) = True Then
                .REGDATE = vbNullString
            Else
                .REGDATE = CStr(rs.Fields("REGDATE"))
            End If
            If IsNull(rs.Fields("SUMMITSENDFLAG")) = True Then
                .SUMMITSENDFLG = vbNullString
            Else
                .SUMMITSENDFLG = CStr(rs.Fields("SUMMITSENDFLAG"))
            End If
            If IsNull(rs.Fields("SENDFLAG")) = True Then
                .SENDFLG = vbNullString
            Else
                .SENDFLG = CStr(rs.Fields("SENDFLAG"))
            End If
            If IsNull(rs.Fields("SENDDATE")) = True Then
                .SENDDATE = vbNullString
            Else
                .SENDDATE = CStr(rs.Fields("SENDDATE"))
            End If
            If IsNull(rs.Fields("WFSTA")) = True Then
                .WFSTA = vbNullString
            Else
                .WFSTA = CStr(rs.Fields("WFSTA"))
            End If
            If IsNull(rs.Fields("HREJCODE")) = True Then
                .HREJCODE = vbNullString
            Else
                .HREJCODE = CStr(rs.Fields("HREJCODE"))
            End If
            If IsNull(rs.Fields("UPDPROC")) = True Then
                .UPDPROC = vbNullString
            Else
                .UPDPROC = CStr(rs.Fields("UPDPROC"))
            End If
            If IsNull(rs.Fields("UPDDATE")) = True Then
                .UPDDATE = vbNullString
            Else
                .UPDDATE = CStr(rs.Fields("UPDDATE"))
            End If
            If IsNull(rs.Fields("MSXLID")) = True Then
                .SXLID = vbNullString
            Else
                .SXLID = CStr(rs.Fields("MSXLID"))
            End If
            If IsNull(rs.Fields("MHINBAN")) = True Then
                .hinban = vbNullString
            Else
                .hinban = CStr(rs.Fields("MHINBAN"))
            End If
            If IsNull(rs.Fields("MREVNUM")) = True Then
                .REVNUM = 0
            Else
                .REVNUM = CInt(rs.Fields("MREVNUM"))
            End If
            If IsNull(rs.Fields("MFACTORY")) = True Then
                .factory = vbNullString
            Else
                .factory = CStr(rs.Fields("MFACTORY"))
            End If
            If IsNull(rs.Fields("MOPECOND")) = True Then
                .opecond = vbNullString
            Else
                .opecond = CStr(rs.Fields("MOPECOND"))
            End If
            If IsNull(rs.Fields("KANKBN")) = True Then
                .KANKBN = vbNullString
            Else
                .KANKBN = CStr(rs.Fields("KANKBN"))
            End If
            If IsNull(rs.Fields("MSMPLEID")) = True Then
                .SMPLEID = vbNullString
            Else
                .SMPLEID = CStr(rs.Fields("MSMPLEID"))
            End If
            If IsNull(rs.Fields("NREJCODE")) = True Then
                .NREJCODE = vbNullString
            Else
                .NREJCODE = CStr(rs.Fields("NREJCODE"))
            End If
            If IsNull(rs.Fields("SHAFLAG")) = True Then
                .SMPLEFLG = vbNullString
            Else
                .SMPLEFLG = CStr(rs.Fields("SHAFLAG"))
            End If
            If IsNull(rs.Fields("RTOP_POS")) = True Then
                .RTOP_POS = vbNullString
            Else
                .RTOP_POS = rs.Fields("RTOP_POS")
            End If
            If IsNull(rs.Fields("RITOP_POS")) = True Then
                .RITOP_POS = vbNullString
            Else
                .RITOP_POS = rs.Fields("RITOP_POS")
            End If
        End With
        intDataCnt = intDataCnt + 1
        rs.MoveNext
    Loop
    If intDataCnt = 0 Then
        ReDim records(0)
        SelWFmap = FUNCTION_RETURN_FAILURE
        sErrMsg = GetMsgStr("EWFM7")
        Exit Function
    End If

    SelWFmap = FUNCTION_RETURN_SUCCESS
    Exit Function
    
PROC_ERR:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    SelWFmap = FUNCTION_RETURN_FAILURE
    sErrMsg = GetMsgStr("SET47")
End Function

'*******************************************************************************
'*    関数名        : SetWFmapData
'*
'*    処理概要      : 1.WFマップ状態表示のSpreadデータ表示処理
'*
'*    パラメータ    : 変数名        ,IO ,型           ,説明
'*                    なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function SetWFmapData() As FUNCTION_RETURN

    Dim intLoopCnt    As Integer
    Dim dblTopPos     As Double
    Dim dblRTopPos    As Double
    Dim dblRITopPos   As Double
    Dim intWarpPoint  As Integer

'Chg Start 2011/03/11 SMPK Miyata
'    With f_cmbc039_4.sprExamine
    With f_cmbc039_4.sprWfmapView
'Chg End   2011/03/11 SMPK Miyata
        .MaxRows = 0
        intWarpPoint = 1
        For intLoopCnt = 0 To UBound(gtWFmap)
            .MaxRows = .MaxRows + 1
            .SetText 3, intLoopCnt + 1, gtWFmap(intLoopCnt).LOTID
            .SetText 5, intLoopCnt + 1, gtWFmap(intLoopCnt).BLOCKSEQ
            .SetText 17, intLoopCnt + 1, gtWFmap(intLoopCnt).INDTM
            .SetText 18, intLoopCnt + 1, gtWFmap(intLoopCnt).BASKETID
            .SetText 19, intLoopCnt + 1, gtWFmap(intLoopCnt).SLOTNO
            .SetText 9, intLoopCnt + 1, gtWFmap(intLoopCnt).CURRWPCS
            .SetText 10, intLoopCnt + 1, gtWFmap(intLoopCnt).EXISTFLG
            If Fix(gtWFmap(intLoopCnt).TOP_POS / 10) = 0 Then
                .SetText 6, intLoopCnt + 1, 0
            Else
                dblTopPos = gtWFmap(intLoopCnt).TOP_POS / 10
                dblTopPos = dblTopPos
                .SetText 6, intLoopCnt + 1, dblTopPos
            End If
            .SetText 11, intLoopCnt + 1, gtWFmap(intLoopCnt).REJCAT
            .SetText 26, intLoopCnt + 1, gtWFmap(intLoopCnt).TXID
            .SetText 25, intLoopCnt + 1, Format(CVar(gtWFmap(intLoopCnt).REGDATE), "yyyy/mm/dd")
            .SetText 27, intLoopCnt + 1, gtWFmap(intLoopCnt).SUMMITSENDFLG
            .SetText 28, intLoopCnt + 1, gtWFmap(intLoopCnt).SENDFLG
            .SetText 29, intLoopCnt + 1, Format(CVar(gtWFmap(intLoopCnt).SENDDATE), "yyyy/mm/dd")
            
            'WF状態判定
            Select Case gtWFmap(intLoopCnt).WFSTA
                Case gsWF_STA_0   '通常
                    'サンプルフラグ判定
                    Select Case gtWFmap(intLoopCnt).SMPLEFLG
                        Case gsWF_SMPL_1    '指示待ち
                            .SetText 1, intLoopCnt + 1, gsWF_STA_SIJI & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF状態フラグの表示を追加
                            .SetText 30, intLoopCnt + 1, intConSprChg_2
                            .col = 1
                            .col2 = 32          'Warp判定対応
                            .row = intLoopCnt + 1
                            .row2 = intLoopCnt + 1
                            .BlockMode = True
'Chg Start 2011/03/10 SMPK Miyata
'                            .backColor = vbYellow
                            '中間抜試WFか？
                            '>>>>> mod start 2011/06/30 Marushita
                            'SAMPLEIDでXSDCW_1を検索し存在するかチェック
                            If ChkXSDCW_1(gtWFmap(intLoopCnt).SMPLEID) = FUNCTION_RETURN_SUCCESS Then
                            'If Right(gtWFmap(intLoopCnt).SMPLEID, 1) = "C" Or _
                               Right(gtWFmap(intLoopCnt).SMPLEID, 1) = "N" Then
                            '<<<<< mod end 2011/06/30 Marushita
                                .backColor = vbCyan
                            Else
                                .backColor = vbYellow
                            End If
'Chg End   2011/03/10 SMPK Miyata
                            .BlockMode = False
                        Case gsWF_SMPL_2    '指示OK
                            .SetText 1, intLoopCnt + 1, gsWF_STA_SIJI_OK & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF状態フラグの表示を追加
                            .SetText 30, intLoopCnt + 1, intConSprChg_2
                            .col = 1
                            .col2 = 32          'Warp判定対応
                            .row = intLoopCnt + 1
                            .row2 = intLoopCnt + 1
                            .BlockMode = True
'Chg Start 2011/03/10 SMPK Miyata
'                            .backColor = vbYellow
                            '中間抜試WFか？
                            '>>>>> mod start 2011/06/30 Marushita
                            'SAMPLEIDでXSDCW_1を検索し存在するかチェック
                            If ChkXSDCW_1(gtWFmap(intLoopCnt).SMPLEID) = FUNCTION_RETURN_SUCCESS Then
                            'If Right(gtWFmap(intLoopCnt).SMPLEID, 1) = "C" Or _
                               Right(gtWFmap(intLoopCnt).SMPLEID, 1) = "N" Then
                            '<<<<< mod end 2011/06/30 Marushita
                                .backColor = vbCyan
                            Else
                                .backColor = vbYellow
                            End If
'Chg End   2011/03/10 SMPK Miyata
                            .BlockMode = False
                        Case gsWF_SMPL_3    '指示NG
                            .SetText 1, intLoopCnt + 1, gsWF_STA_SIJI_NG & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF状態フラグの表示を追加
                            .SetText 30, intLoopCnt + 1, intConSprChg_2
                            .col = 1
                            .col2 = 32          'Warp判定対応
                            .row = intLoopCnt + 1
                            .row2 = intLoopCnt + 1
                            .BlockMode = True
'Chg Start 2011/03/10 SMPK Miyata
'                            .backColor = vbYellow
                            '中間抜試WFか？
                            '>>>>> mod start 2011/06/30 Marushita
                            'SAMPLEIDでXSDCW_1を検索し存在するかチェック
                            If ChkXSDCW_1(gtWFmap(intLoopCnt).SMPLEID) = FUNCTION_RETURN_SUCCESS Then
                            'If Right(gtWFmap(intLoopCnt).SMPLEID, 1) = "C" Or _
                               Right(gtWFmap(intLoopCnt).SMPLEID, 1) = "N" Then
                            '<<<<< mod end 2011/06/30 Marushita
                                .backColor = vbCyan
                            Else
                                .backColor = vbYellow
                            End If
'Chg End   2011/03/10 SMPK Miyata
                            .BlockMode = False
                        Case Else
                            .SetText 1, intLoopCnt + 1, gsWF_STA_NORMAL & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF状態フラグの表示を追加
                            .SetText 30, intLoopCnt + 1, intConSprChg_1
                            .col = 1
                            .col2 = 32          'Warp判定対応
                            .row = intLoopCnt + 1
                            .row2 = intLoopCnt + 1
                            .BlockMode = True
                            .backColor = &H80FF80
                            .BlockMode = False
                    End Select
                Case gsWF_STA_1   '共有
                    .SetText 1, intLoopCnt + 1, gsWF_STA_NORMAL & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF状態フラグの表示を追加
                    .SetText 30, intLoopCnt + 1, intConSprChg_1
                    .col = 1
                    .col2 = 32          'Warp判定対応
                    .row = intLoopCnt + 1
                    .row2 = intLoopCnt + 1
                    .BlockMode = True
                    .backColor = &H80FF80
                    .BlockMode = False
                Case gsWF_STA_4   '欠落
                    'サンプルフラグ判定
                    Select Case gtWFmap(intLoopCnt).SMPLEFLG
                        Case gsWF_SMPL_4    'サンプルの結果以外はすべて欠落と判断する
                            .SetText 1, intLoopCnt + 1, gsWF_STA_SIJI_KEKKA & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF状態フラグの表示を追加
                            .SetText 30, intLoopCnt + 1, intConSprChg_2
                            .col = 1
                            .col2 = 32          'Warp判定対応
                            .row = intLoopCnt + 1
                            .row2 = intLoopCnt + 1
                            .BlockMode = True
'Chg Start 2011/03/10 SMPK Miyata
'                            .backColor = vbYellow
                            '中間抜試WFか？
                            '>>>>> mod start 2011/06/30 Marushita
                            'SAMPLEIDでXSDCW_1を検索し存在するかチェック
                            If ChkXSDCW_1(gtWFmap(intLoopCnt).SMPLEID) = FUNCTION_RETURN_SUCCESS Then
                            'If Right(gtWFmap(intLoopCnt).SMPLEID, 1) = "C" Or _
                               Right(gtWFmap(intLoopCnt).SMPLEID, 1) = "N" Then
                            '<<<<< mod end 2011/06/30 Marushita
                                .backColor = vbCyan
                            Else
                                .backColor = vbYellow
                            End If
'Chg End   2011/03/10 SMPK Miyata
                            .BlockMode = False
                        Case Else
                            .SetText 1, intLoopCnt + 1, gsWF_STA_STA_K & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF状態フラグの表示を追加
                            .SetText 30, intLoopCnt + 1, intConSprChg_3
                            .col = 1
                            .col2 = 32          'Warp判定対応
                            .row = intLoopCnt + 1
                            .row2 = intLoopCnt + 1
                            .BlockMode = True
                            .backColor = vbRed
                            .BlockMode = False
                    End Select
            End Select
            
            .SetText 14, intLoopCnt + 1, gtWFmap(intLoopCnt).SMPLEID    ' 無条件に表示
            .SetText 12, intLoopCnt + 1, gtWFmap(intLoopCnt).HREJCODE
            .SetText 23, intLoopCnt + 1, gtWFmap(intLoopCnt).UPDPROC
            .SetText 24, intLoopCnt + 1, gtWFmap(intLoopCnt).UPDDATE
            .SetText 2, intLoopCnt + 1, gtWFmap(intLoopCnt).SXLID
            .SetText 4, intLoopCnt + 1, gtWFmap(intLoopCnt).hinban
            .SetText 20, intLoopCnt + 1, gtWFmap(intLoopCnt).REVNUM
            .SetText 21, intLoopCnt + 1, gtWFmap(intLoopCnt).factory
            .SetText 22, intLoopCnt + 1, gtWFmap(intLoopCnt).opecond
            .SetText 13, intLoopCnt + 1, gtWFmap(intLoopCnt).KANKBN
            .SetText 15, intLoopCnt + 1, gtWFmap(intLoopCnt).NREJCODE
            .SetText 16, intLoopCnt + 1, gtWFmap(intLoopCnt).SMPLEFLG
            
            If gtWFmap(intLoopCnt).RTOP_POS = 0 Then
                .SetText 7, intLoopCnt + 1, 0
            Else
                dblRTopPos = gtWFmap(intLoopCnt).RTOP_POS
                dblRTopPos = dblRTopPos
                .SetText 7, intLoopCnt + 1, dblRTopPos
            End If
            If gtWFmap(intLoopCnt).RITOP_POS = 0 Then
                .SetText 8, intLoopCnt + 1, 0
            Else
                dblRITopPos = gtWFmap(intLoopCnt).RITOP_POS
                dblRITopPos = dblRITopPos
                .SetText 8, intLoopCnt + 1, dblRITopPos
            End If
            
            'Warp情報表示追加
            If UBound(tWarpMeasG) >= intWarpPoint Then
                'ﾌﾞﾛｯｸIDとﾌﾞﾛｯｸ内連番で紐付け
                If tWarpMeasG(intWarpPoint).BLOCKID = gtWFmap(intLoopCnt).LOTID And _
                   tWarpMeasG(intWarpPoint).WAFID = gtWFmap(intLoopCnt).BLOCKSEQ Then
                    '実ﾃﾞｰﾀが無い場合は表示しない
                    If tWarpMeasG(intWarpPoint).EXISTFLG > 0 Then
                        'Warp値
                        .SetText 31, intLoopCnt + 1, CStr(DBData2DispData_nl(tWarpMeasG(intWarpPoint).MEASDATA))
                        '判定
                        .SetText 32, intLoopCnt + 1, IIf(tWarpMeasG(intWarpPoint).Judg, "OK", "NG")
                    End If
                    intWarpPoint = intWarpPoint + 1
                End If
            End If
        Next
    End With

    'ｽﾌﾟﾚｯﾄﾞﾃﾞｰﾀソート
'Chg Start 2011/03/11 SMPK Miyata
'    With f_cmbc039_4.sprExamine
    With f_cmbc039_4.sprWfmapView
'Chg End   2011/03/11 SMPK Miyata
        .BlockMode = True
        .row = 1
        .col = 1
        .row2 = .MaxRows
        .col2 = .MaxCols
        .SortBy = SortByRow
        .SortKey(1) = 8
        .SortKeyOrder(1) = SortKeyOrderAscending
        .Action = ActionSort
        .BlockMode = False
    End With
End Function

'*******************************************************************************
'*    関数名        : ChkXSDCW_1
'*
'*    処理概要      : 対象サンプルIDがXSDCW_1に存在するかチェックする
'*
'*    パラメータ    : 変数名        ,IO ,型           ,説明
'*                    sSXLID        ,I  ,STRING       ,サンプルID
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function ChkXSDCW_1(ByVal sSXLID As String) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sSQL        As String

    'エラーハンドラの設定
    On Error GoTo PROC_ERR

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    ChkXSDCW_1 = FUNCTION_RETURN_FAILURE

    sSQL = "select REPSMPLIDCW from XSDCW_1 where REPSMPLIDCW = '" & sSXLID & "' and LIVKCW = '0'"
    Set rs = OraDB.CreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount = 0 Then
        Exit Function
    End If
    
    Set rs = Nothing
    ChkXSDCW_1 = FUNCTION_RETURN_SUCCESS

proc_exit:

    Exit Function

PROC_ERR:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    ChkXSDCW_1 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

Attribute VB_Name = "s_cmbc039"
Option Explicit

Public typ_Param001b As DBDRV_scmzc_fcmlc001b_SXL039
Public Const MAXREC As Integer = 256
Private intChkPos As Integer                                    ' チェック位置

''Public MaxLine As Integer
Public SelectSxlID039 As String
Public typ_ww() As DBDRV_scmzc_fcmlc001b_SXL039   '待ち一覧情報
Public WFJudgExecOkFlag() As Boolean    'WF総合判定実行可能フラグ

'判定フローはこんな感じですか？
' ＜判定フロー＞
' 仕様保証方法＿処 --+--なし --実績（該当位置）--あってもなくても判定OK
'　　　　　　　　　　|
'                   +--あり --実績（該当位置) --+--あり -- 判定チェック --+-- OK
'                                              |                        |
'                                              |                        +-- MG
'                                              |
'                                              +--なし --+-- 検査指示５・６以外の場合 ---NG
'　　　　　　　　　　　　　　　　　　　　　　　　（検査指示は、指示を立てる側が正常に立てていると考えている）


''
'' 定数定義
''
Public lStfMst As Long
Public intEnCmd As Integer
Public Const MAXCNT As Integer = 16                             ' 最大件数
Public Const SxlTop039 As Integer = 1                           ' TOP側
Public Const SxlTail039 As Integer = 2                          ' TAIL側
Public Const SxlMidl039 As Integer = 3                          ' MIDLE側    'Add 2011/03/07 SMPK Miyata

Public Const KSYSCLASS As String = "GP"                         ' システム区分
Public Const MSYSCLASS As String = "NM"                         ' システム区分
Public Const KCLASS As String = "01"                            ' クラス
Public Const KCODE As String = "1"                              ' コード

Private Const cnEnableColor As Long = &H80FF80                  ' 有効カラー
Private Const cnEnableColor2 As Long = vbWindowBackground       ' 有効カラー
Private Const cnDisenableColor As Long = &H80FF80               ' 無効カラー
Private Const cnDisenableGrayColor As Long = vbButtonFace       ' 無効カラー（灰色）
Private Const cnWarningColor As Long = &H8080FF                 ' 警告カラー

Public Const WFRES039 As Integer = 0
Public Const WFOI As Integer = 1
Public Const WFBMD1 As Integer = 2
Public Const WFBMD2 As Integer = 3
Public Const WFBMD3 As Integer = 4
Public Const WFOSF1 As Integer = 5
Public Const WFOSF2 As Integer = 6
Public Const WFOSF3 As Integer = 7
Public Const WFOSF4 As Integer = 8
Public Const WFDS As Integer = 9
Public Const WFDZ As Integer = 10
Public Const WFSP As Integer = 11
Public Const WFDOI1 As Integer = 12
Public Const WFDOI2 As Integer = 13
Public Const WFDOI3 As Integer = 14

Public Const OSWFRES039 As String = "RES"
Public Const OSWFOI039 As String = "OI"
Public Const OSWFBMD1039 As String = "BMD1"
Public Const OSWFBMD2039 As String = "BMD2"
Public Const OSWFBMD3039 As String = "BMD3"
Public Const OSWFOSF1039 As String = "OSF1"
Public Const OSWFOSF2039 As String = "OSF2"
Public Const OSWFOSF3039 As String = "OSF3"
Public Const OSWFOSF4039 As String = "OSF4"
Public Const OSWFDS039 As String = "DSOD"
Public Const OSWFDZ039 As String = "DZ"
Public Const OSWFSP039 As String = "SPV"
Public Const OSWFDOI1039 As String = "DOI1"
Public Const OSWFDOI2039 As String = "DOI2"
Public Const OSWFDOI3039 As String = "DOI3"
Public Const OSWFAOI039 As String = "AOI"                       ' 残存酸素追加

' コードマスター
Public Type typ_CodeMaster
    SYSCLASS As String * 2                                      ' システム区分
    Class As String * 2                                         ' 区分
    CODE As String * 5                                          ' コード
    INFO1 As String                                             ' 情報１
    INFO2 As String                                             ' 情報２
    INFO3 As String                                             ' 情報３
    INFO4 As String                                             ' 情報４
    INFO5 As String                                             ' 情報５
    INFO6 As String                                             ' 情報６
    INFO7 As String                                             ' 情報７
    INFO8 As String                                             ' 情報８
    INFO9 As String                                             ' 情報９
    NOTE As String                                              ' 備考
    TSTAFFID As String * 8                                      ' 登録社員ID
    REGDATE As Date                                             ' 登録日付
    KSTAFFID As String * 8                                      ' 更新社員ID
    UPDDATE As Date                                             ' 更新日付
End Type

'各実績情報
Public Type typ_ALLRSLT039
    pos As Integer                                              ' 結晶内開始位置
    NAIYO As String                                             ' 内容
    INFO1 As String                                             ' 情報１
    INFO2 As String                                             ' 情報２
    INFO3 As String                                             ' 情報３
    INFO4 As String                                             ' 情報４
    OKNG  As String                                             ' 判定結果
    SMPLID As String                                            ' サンプルＮｏ
End Type

'全情報構造体
Public Type typ_AllTypes
    StrStaffId As String                                        ' スタッフID
    strStaffName As String                                      ' スタッフ名
    dblScut(2) As Double                                        ' 再カット位置
    bOKNG(2) As Boolean                                         ' 比抵抗判定
    COEF(2) As Double                                           ' 偏析係数
    JudgRes(2) As Boolean                                       ' 比抵抗判定
    JudgRrg(2) As Boolean                                       ' RRG判定
    typ_Param As DBDRV_scmzc_fcmlc001b_SXL039                   ' SXL管理（待ち一覧から）
    typ_si As type_DBDRV_scmzc_fcmlc001c_Siyou039               ' 製品仕様
    typ_y013top() As typ_TBCMY013                               ' 測定結果(TOP)
    typ_y013tail() As typ_TBCMY013                              ' 測定結果(TAIL)
    typ_y013(2, MAXCNT) As typ_TBCMY013                         ' 測定結果
    typ_hage(MAXCNT) As typ_TBCMH004                            ' 引上げ終了実績
    typ_rslt(2, MAXCNT) As typ_ALLRSLT039                       ' 各実績情報
End Type

'仕様検査支持構造体
Type Judg_Spec_Wf
        rs As Boolean
        Oi As Boolean
        B1 As Boolean
        B2 As Boolean
        B3 As Boolean
        L1 As Boolean
        L2 As Boolean
        L3 As Boolean
        L4 As Boolean
        Dsod As Boolean
        sp As Boolean
        DZ As Boolean
        Doi1 As Boolean
        Doi2 As Boolean
        Doi3 As Boolean
End Type

Public typ_AType As typ_AllTypes                                ' 全情報構造体-----typ_AllTypes→typ_AllTypesC
Public TotalJudg039 As Boolean                                  ' トータル判定
Public EPSiyouSansyouFlg As Boolean                             ' エピ仕様参照済みフラグ
Public HErrMsg As String
Public bPPlus As Boolean                                        ' P+Flag
Public bNPlus As Boolean                                        ' N+Flag
Public JiltusekiUmu(2, MAXCNT) As Boolean                       ' 実績有無情報
Public MeasFlag(2) As Judg_Spec_Wf                              ' 仕様検査支持構造体
Public TmpOsfData(1, 2, MAXCNT) As String                       ' OSF平均/最大値
Public TmpOsfMBNP(2, 2, MAXCNT) As String * 1                   ' OSF面内分布

Public iChkAoi      As Integer                                  ' 残存酸素仕様チェック
Public sKanrenFlg   As String                                   ' 関連ﾌﾞﾛｯｸ有無　08/01/31 ooba

Public iMode                As Integer                          ' (0:一覧表示,1:一覧表示済み)
Public sCmbMukesaki         As String                           ' 選択した向先コード
Public sCmbMukeName         As String                           ' 選択した向先名
Public bMukesakiChgFlg      As Boolean                          ' 向先変更フラグ(True:向先変更有)

Public Type typ_Mukesaki
    sMukeCode As String                                         ' 向先コード
    sMukeName As String                                         ' 向先名
End Type

Public s_MukesakiBase() As typ_Mukesaki
Public sBaseMukesaki        As String                           ' Baseの向先

' SXLの対象となるﾌﾞﾛｯｸ保存用構造体
Public Type typ_IntoBlock
    SORTID As String
    FULLID As String
End Type

' ブロック情報
Public Type typ_BlkInf
    BLOCKID As String * 12                                      ' ブロックID
    LENGTH As Integer                                           ' 長さ
    REALLEN As Integer                                          ' 実長さ
    KRPROCCD As String * 5                                      ' 現在管理工程
    NOWPROC As String * 5                                       ' 現在工程
    LPKRPROCCD As String * 5                                    ' 最終通過管理工程
    LASTPASS As String * 5                                      ' 最終通過工程
    RSTATCLS As String * 1                                      ' 流動状態区分
    SEED As String * 4                                          ' シード
    COF As type_Coefficient                                     ' 偏析係数計算
    SAMPFLAG As Boolean                                         ' サンプル取得フラグ
End Type

'カット位置用構造体
Public Type typ_CMKC001C
    CRYNUM As String * 12                                       ' 結晶番号
    INGOTPOS As Integer                                         ' 結晶内開始位置
    LENGTH As Integer                                           ' 長さ
End Type

' ブロック情報
Public Type typ_BlkInf3
    BLOCKID As String * 12                                      ' ブロックID
    LENGTH As Integer                                           ' 長さ
    REALLEN As Integer                                          ' 実長さ
    NOWPROC As String * 5                                       ' 現在工程
    DELFLG As String * 1                                        ' 削除区分
    COF As type_Coefficient                                     ' 偏析係数計算
End Type

Public tblHinMng() As typ_TBCME041                              ' 品番管理
Public tblWafSmp() As typ_XSDCW                                 ' 新サンプル管理（SXL）
Public tblBlkInf() As typ_BlkInf                                ' ブロック情報テーブル
Public tblTotal As typ_AllTypesC                                ' 前画面からの情報保持構造体----typ_AllTypes→typ_AllTypesCに変更
Public tblWfSxlMng() As typ_TBCME042                            ' SXL管理構造体
Public tblWfSxlMngS() As typ_TBCME042                           ' 測定評価指示用SXL管理構造体
Public tblWfSample() As typ_WfSampleGr                          ' WFサンプル管理
Public SxlIntoBlock() As typ_IntoBlock                          ' SXLの対象となるﾌﾞﾛｯｸ構造体
Public tblPrcList() As typ_TBCMB005                             ' 区分用コードマスター構造体
Public tblHinbanRs() As type_DBDRV_scmzc_fcmlc001d_In           ' 品番情報保持構造体
Public tblsiyou() As type_DBDRV_scmzc_fcmlc001d_WfSiyou         ' 仕様情報構造体(表示用)
Public tblsmp() As type_DBDRV_scmzc_fcmlc001d_WfSmp             ' サンプル情報構造体(表示用)
Public tblWfHantei As typ_TBCMW005                              ' WF総合判定実績
Public tblHuriHai() As typ_TBCMW006                             ' 振替廃棄実績
Public tblSokuSizi() As typ_TBCMY003                            ' 測定評価方法指示構造体
Public tblSxlKSiji() As typ_TBCMY007                            ' Ｓｘｌ確定指示
Public NoTestHinList() As tFullHinban                           ' 抜試の発生しない品番

'Warp判定対応
Public bMapWarpFlg      As Boolean                              ' WFﾏｯﾌﾟとWarp実績の紐付けﾌﾗｸﾞ
Public tMapHin()        As typ_MapHinData                       ' WFﾏｯﾌﾟ上の品番ﾃﾞｰﾀ
Public sWrpLOTID()      As String                               ' ﾌﾞﾛｯｸID(Warp実績紐付け用)
Public iWrpBLOCKSEQ()   As Integer                              ' ﾌﾞﾛｯｸ内連番(Warp実績紐付け用)

' 抜試指示
Public Type typ_WafInd
    BLOCKID As String * 12                                      ' ブロックID
    BlockPos As Integer                                         ' ブロックＰ
    SAMPLEID    As Variant                                      ' add 2003/03/28 hitec)matsumoto サンプルIDを取得
    SAMPLEID2   As Variant                                      ' add 2003/03/28 hitec)matsumoto サンプルID2を取得
    INGOTPOS As Integer                                         ' 結晶Ｐ
    BkIngotPos  As Integer
    LENGTH As Integer                                           ' 長さ
    HINUP As tFullHinban                                        ' 上品番
    HINDN As tFullHinban                                        ' 下品番
    SMP As typ_WFSample                                         ' 検査項目
    HINFLG As Boolean                                           ' 品番区切りフラグ
    SMPFLG As Boolean                                           ' WFサンプル区切りフラグ
    ERRDNFLG As Boolean                                         ' 下品番エラーフラグ
    SMPLKBN1 As String * 1                                      ' サンプル区分１
    SMPLKBN2 As String * 1                                      ' サンプル区分２
End Type
Public tblWafInd() As typ_WafInd                                ' 抜試指示テーブル
Public tblNukishi() As typ_XSDCW                                ' 抜試データ構造体作成用

' 欠落ウェハー
Public Type typ_LackMap
    BLOCKID As String * 12                                      ' ブロックID
    LACKPOSS As Double                                          ' 欠落位置(From)
    LACKPOSE As Double                                          ' 欠落位置(To)
    REJCAT As String * 1                                        ' 欠落理由
    LACKCNTS As Integer                                         ' 欠落枚目(From)
    LACKCNTE As Integer                                         ' 欠落枚目(To)
End Type
Public tblLackMap() As typ_LackMap                              ' 欠落ウェハーテーブル


' SXLサンプル情報
Public Type typ_SxlSmp
    strCryNum As String * 12                                    ' 結晶番号
    intIngotpos As Integer                                      ' 結晶内開始位置
    intLength As Integer                                        ' 長さ
    strSXLID As String * 13                                     ' SXLID
    StrHinban As String * 12                                    ' 品番
    strSMPLID As String * 16                                    ' サンプルID
    intCount As Integer                                         ' 枚数
    strSMPLUMU As String * 1                                    ' サンプル有無区分
    datREGDATE As Date                                          ' 登録日付
    datUPDDATE As Date                                          ' 更新日付
    strWFINDRS As String * 1                                    ' WF検査指示（Rs)
    strWFINDOI As String * 1                                    ' WF検査指示（Oi)
    strWFINDB1 As String * 1                                    ' WF検査指示（B1)
    strWFINDB2 As String * 1                                    ' WF検査指示（B2）
    strWFINDB3 As String * 1                                    ' WF検査指示（B3)
    strWFINDL1 As String * 1                                    ' WF検査指示（L1)
    strWFINDL2 As String * 1                                    ' WF検査指示（L2)
    strWFINDL3 As String * 1                                    ' WF検査指示（L3)
    strWFINDL4 As String * 1                                    ' WF検査指示（L4)
    strWFINDDS As String * 1                                    ' WF検査指示（DS)
    strWFINDDZ As String * 1                                    ' WF検査指示（DZ)
    strWFINDSP As String * 1                                    ' WF検査指示（SP)
    strWFINDDO1 As String * 1                                   ' WF検査指示（DO1)
    strWFINDDO2 As String * 1                                   ' WF検査指示（DO2)
    strWFINDDO3 As String * 1                                   ' WF検査指示（DO3)
    strWFRESRS As String * 1                                    ' WF検査実績（Rs)
    strWFRESOI As String * 1                                    ' WF検査実績（Oi)
    strWFRESB1 As String * 1                                    ' WF検査実績（B1)
    strWFRESB2 As String * 1                                    ' WF検査実績（B2）
    strWFRESB3 As String * 1                                    ' WF検査実績（B3)
    strWFRESL1 As String * 1                                    ' WF検査実績（L1)
    strWFRESL2 As String * 1                                    ' WF検査実績（L2)
    strWFRESL3 As String * 1                                    ' WF検査実績（L3)
    strWFRESL4 As String * 1                                    ' WF検査実績（L4)
    strWFRESDS As String * 1                                    ' WF検査実績（DS)
    strWFRESDZ As String * 1                                    ' WF検査実績（DZ)
    strWFRESSP As String * 1                                    ' WF検査実績（SP)
    strWFRESDO1 As String * 1                                   ' WF検査実績（DO1)
    strWFRESDO2 As String * 1                                   ' WF検査実績（DO2)
    strWFRESDO3 As String * 1                                   ' WF検査実績（DO3)
End Type

Public Const WATCH_PROCCD           As String = "CW750/CW760/CW000/"            ' 流動監視工程一覧用(複数工程の場合は"/"で区切ること。)  add 09/03/16 SETkimizuka
Public Const WATCH_PROCCD_ENT       As String = "CW750/CW000/"                  ' 流動監視工程実行チェック用(複数工程の場合は"/"で区切ること。)  add 09/03/16 SETkimizuka
Public Const WATCH_PROCCD_NUKISI    As String = "CW760/CW000/"                  ' 流動監視工程再抜試チェック用(複数工程の場合は"/"で区切ること。)  add 09/03/16 SETkimizuka
Public Const CAP_FNAME              As String = "\CMBC039CAP.BMP"               ' キャプチャファイル名 add 2011/07/13 Marushita
Public gsMukeCd                     As String                                   ' 向先(待ち一覧->0枚ロット一覧受渡用) add 2012/09/07 Marushita

'*******************************************************************************
'*    関数名        : WfWaitSetAllData
'*
'*    処理概要      : 1.画面情報を情報構造体に設定する
'*
'*    パラメータ    : 変数名        ,IO ,型                            ,説明
'*                    udt_ww        ,O  ,DBDRV_scmzc_fcmlc001b_SXL039  ,SXL管理
'*
'*    戻り値        : 正常終了時はTrue, エラー終了時はFalse
'*
'*******************************************************************************
Public Function WfWaitSetAllData(udt_ww() As DBDRV_scmzc_fcmlc001b_SXL039) As Boolean
    Dim sErrMsg As String

    DoEvents
    OraDB.BeginTrans

    If DBDRV_scmzc_fcmlc001b_Disp("CW750", udt_ww(), sErrMsg) = FUNCTION_RETURN_FAILURE Then
        WfWaitSetAllData = True
        OraDB.Rollback
        Exit Function
    End If

    OraDB.CommitTrans

    ' データ無し
    If UBound(udt_ww) = 0 Then
        WfWaitSetAllData = False
    Else
        'add 09/03/16 SETkimizuka Start
        If f_cmbc039_1.chkY4Disp.Value = 1 Then
            f_cmbc039_1.spdWait.col = 13
            f_cmbc039_1.spdWait.ColHidden = False
            f_cmbc039_1.spdWait.col = 14
            f_cmbc039_1.spdWait.ColHidden = False
            f_cmbc039_1.spdWait.col = 15
            f_cmbc039_1.spdWait.ColHidden = False
            'Debug.Print "9 " & Now & " XODY4取得開始"
            Call DBDRV_XODY4GET(udt_ww())
            'Debug.Print "A " & Now & " XODY4取得終了"
        Else
            f_cmbc039_1.spdWait.col = 13
            f_cmbc039_1.spdWait.ColHidden = True
            f_cmbc039_1.spdWait.col = 14
            f_cmbc039_1.spdWait.ColHidden = True
            f_cmbc039_1.spdWait.col = 15
            f_cmbc039_1.spdWait.ColHidden = True
        End If
        'add 09/03/16 SETkimizuka End
        WfWaitSetAllData = True
    End If
End Function

'*******************************************************************************
'*    関数名        : HoldLot_Get
'*
'*    処理概要      : 1.ホールドロット検索処理（使用していない）
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    sXtal          ,I  ,String   ,結晶番号
'*                    sHOLDBCB       ,O  ,String   ,ホールド区分
'*                    WFHOLDFLGCB   ,O  ,String   ,ホールド区分(WF)
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function HoldLot_Get(sXtal As String, sHOLDBCB As String, sWFHOLDFLGCB As String) As FUNCTION_RETURN
    Dim sSql        As String
    Dim rs          As OraDynaset
    Dim intSXLCnt   As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc039_SQL.bas -- Function HoldLot_Get"

    HoldLot_Get = FUNCTION_RETURN_SUCCESS

    sSql = "select holdbcb, WFHOLDFLGCB "    'ﾎｰﾙﾄﾞ区分(WF)追加
    sSql = sSql & " from xsdcb "
    sSql = sSql & " where sxlidcb='" & sXtal & "' "
    Set rs = OraDB.CreateDynaset(sSql, ORADYN_DEFAULT)

    If rs.EOF = False Then
        sHOLDBCB = rs("holdbcb")
        'ﾎｰﾙﾄﾞ区分(WF)追加
        If IsNull(rs("WFHOLDFLGCB")) = False Then
            sWFHOLDFLGCB = rs("WFHOLDFLGCB")
        Else
            sWFHOLDFLGCB = " "
        End If
    Else
        HoldLot_Get = FUNCTION_RETURN_FAILURE
    End If
    rs.Close
    Set rs = Nothing

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    HoldLot_Get = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        : InitHensu
'*
'*    処理概要      : 1.変数初期化
'*
'*    パラメータ    : 変数名        ,IO ,型           ,説明
'*                    udt_A         ,IO ,typ_AllTypes ,各情報構造体
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Public Sub InitHensu(udt_A As typ_AllTypes)
    Dim i As Integer, j As Integer

    For i = 1 To 2
        For j = 0 To MAXCNT
            With udt_A
                .typ_y013(i, j).SAMPLEID = "0"         ' サンプルID
                .typ_y013(i, j).MESDATA1 = "0"         ' 測定データその１
                .typ_y013(i, j).MESDATA2 = "0"         ' 測定データその２
                .typ_y013(i, j).MESDATA3 = "0"         ' 測定データその３
                .typ_y013(i, j).MESDATA4 = "0"         ' 測定データその４
                .typ_y013(i, j).MESDATA5 = "0"         ' 測定データその５
                .typ_y013(i, j).MESDATA6 = "0"         ' 測定データその６
                .typ_y013(i, j).MESDATA7 = "0"         ' 測定データその７
                .typ_y013(i, j).MESDATA8 = "0"         ' 測定データその８
                .typ_y013(i, j).MESDATA9 = "0"         ' 測定データその９
                .typ_y013(i, j).MESDATA10 = "0"        ' 測定データその１０
                .typ_y013(i, j).MESDATA11 = "0"        ' 測定データその１１
                .typ_y013(i, j).MESDATA12 = "0"        ' 測定データその１２
                .typ_y013(i, j).MESDATA13 = "0"        ' 測定データその１３
                .typ_y013(i, j).MESDATA14 = "0"        ' 測定データその１４
                .typ_y013(i, j).MESDATA15 = "0"        ' 測定データその１５
            End With
        Next
    Next
End Sub

'*******************************************************************************
'*    関数名        : Sub_S_SetParamData
'*
'*    処理概要      : 1.前画面からの引数を設定する
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Public Sub Sub_S_SetParamData()
    typ_AType.typ_Param = typ_Param001b
End Sub

'*******************************************************************************
'*    関数名        : SetMERInd
'*
'*    処理概要      : 1.測定評価結果配列にDB検索したレコードを整列する
'*
'*    パラメータ    : 変数名        ,IO ,型           ,説明
'*                    typ_a         ,IO ,typ_AllTypes ,各情報構造体
'*                    typ_y013()    ,I  ,typ_TBCMY013 ,測定評価結果情報構造体
'*                    tt            ,I  ,Integer      ,TOP・TAIL
'*
'*    戻り値        : 正常終了時はTrue, エラー終了時は False
'*
'*******************************************************************************
Public Function SetMERInd(udt_A As typ_AllTypes, _
                          udt_y013() As typ_TBCMY013, _
                          tt As Integer) As Boolean
    Dim i As Integer

    With udt_A
        For i = 1 To UBound(udt_y013)
            Select Case Trim(udt_y013(i).Spec)
                Case OSWFRES ' RES
                    .typ_y013(tt, WFRES) = udt_y013(i)
                Case OSWFOI ' OI
                    .typ_y013(tt, WFOI) = udt_y013(i)
                Case OSWFBMD1 ' BMD1
                    .typ_y013(tt, WFBMD1) = udt_y013(i)
                Case OSWFBMD2 ' BMD2
                    .typ_y013(tt, WFBMD2) = udt_y013(i)
                Case OSWFBMD3 ' BMD3
                    .typ_y013(tt, WFBMD3) = udt_y013(i)
                Case OSWFOSF1 ' OSF1
                    .typ_y013(tt, WFOSF1) = udt_y013(i)
                Case OSWFOSF2 ' OSF2
                    .typ_y013(tt, WFOSF2) = udt_y013(i)
                Case OSWFOSF3 ' OSF3
                    .typ_y013(tt, WFOSF3) = udt_y013(i)
                Case OSWFOSF4 ' OSF4
                    .typ_y013(tt, WFOSF4) = udt_y013(i)
                Case OSWFDS ' DSOD
                    .typ_y013(tt, WFDS) = udt_y013(i)
                Case OSWFDZ ' DZ
                    .typ_y013(tt, WFDZ) = udt_y013(i)
                Case OSWFSP ' SPV
                    .typ_y013(tt, WFSP) = udt_y013(i)
                Case OSWFDOI1 ' DOI1
                    .typ_y013(tt, WFDOI1) = udt_y013(i)
                Case OSWFDOI2 ' DOI2
                    .typ_y013(tt, WFDOI2) = udt_y013(i)
                Case OSWFDOI3 ' DOI3
                    .typ_y013(tt, WFDOI3) = udt_y013(i)
            End Select
        Next
    End With
    SetMERInd = True
End Function

'*******************************************************************************
'*    関数名        : RegWfSogoRsltOK
'*
'*    処理概要      : 1.総合判定実績挿入
'*                    2.WF_GD実績(TBCMJ015)更新処理
'*                    3.SXL管理更新
'*                    4.WFサンプル管理更新
'*
'*    パラメータ    : 変数名        ,IO ,型           ,説明
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function RegWfSogoRsltOK() As FUNCTION_RETURN
    Dim udt_soz         As typ_TBCMW005                             ' WF総合判定実績
    Dim udt_sxl         As type_DBDRV_scmzc_fcmlc001c_UpdSXL1       ' SXL管理
    Dim udt_WFSmp(2)    As type_DBDRV_scmzc_fcmlc001c_UpdWfCrySmp
    Dim i               As Long
    Dim intCnt          As Integer

    'WF総合判定実績
    With udt_soz
        .CRYNUM = typ_CType.typ_Param.CRYNUM                                ' 結晶番号
        .INGOTPOS = typ_CType.typ_Param.INGOTPOS                            ' インゴット位置
        .CRYLEN = typ_CType.typ_Param.LENGTH                                ' 長さ
        .KRPROCCD = MGPRCD_WFC_SOUGOUHANTEI                                 ' 管理工程コード
        .PROCCODE = PROCD_WFC_SOUGOUHANTEI                                  ' 工程コード
        .SXLID = NtoS(typ_CType.typ_Param.SXLID)                                  ' SXLID
        .CODE = "0"                                                         ' 区分コード
        .TSTAFFID = typ_CType.StrStaffId                                    ' 登録社員ID
    End With

    'WF総合判定実績挿入
    If DBDRV_scmzc_fcmlc001c_InsWfSougou(udt_soz) <> FUNCTION_RETURN_SUCCESS Then
        f_cmbc039_2.lblMsg.Caption = GetMsgStr("EGET2", "W005")
        RegWfSogoRsltOK = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    '' WF_GD実績(TBCMJ015)更新処理
    If UBound(typ_J015_WFGDUpd) > 0 Then
        'ﾃﾞｰﾀ数分UPDATE
        For intCnt = 1 To UBound(typ_J015_WFGDUpd)
            If DBDRV_scmzc_fcmlc001c_UpdGDdata(typ_J015_WFGDUpd(intCnt), typ_CType.StrStaffId) _
                                        <> FUNCTION_RETURN_SUCCESS Then
                f_cmbc039_2.lblMsg.Caption = GetMsgStr("EAPLY") & "J015"
                RegWfSogoRsltOK = FUNCTION_RETURN_FAILURE
                Exit Function
            End If
        Next
    End If

    'SXL管理
    With udt_sxl
        .CRYNUM = NtoS(typ_AType.typ_Param.CRYNUMCA)                        ' 結晶番号
        .INGOTPOS = typ_CType.typ_Param.INGOTPOS                            ' 結晶内開始位置
        .NOWPROC = PROCD_SXL_KAKUTEI                                        ' 現在工程
        .LASTPASS = PROCD_WFC_SOUGOUHANTEI                                  ' 最終通過工程
    End With

    'SXL管理更新
    If DBDRV_scmzc_fcmlc001c_UpdSXL1(udt_sxl) <> FUNCTION_RETURN_SUCCESS Then
        f_cmbc039_2.lblMsg.Caption = GetMsgStr("EGET2", "E042")
        RegWfSogoRsltOK = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    'WFサンプル管理が存在する場合は確定区分コードに1を立てる
    'エピ先行評価追加対応
    If (UBound(typ_CType.typ_y013top) <> 0 Or UBound(typ_CType_EP.typ_y022top) <> 0) _
        And (UBound(typ_CType.typ_y013tail) <> 0 Or UBound(typ_CType_EP.typ_y022tail) <> 0) Then

        'WFサンプル管理
        udt_WFSmp(1).CRYNUM = NtoS(typ_CType.typ_Param.CRYNUM)                  ' 結晶番号
        udt_WFSmp(1).INGOTPOS = typ_CType.typ_Param.WFSMP(SxlTop039).INPOSCW    ' 結晶内開始位置
        udt_WFSmp(1).SMPKBN = typ_CType.typ_Param.WFSMP(SxlTop039).SMPKBNCW     ' サンプル区分
        udt_WFSmp(2).CRYNUM = NtoS(typ_CType.typ_Param.CRYNUM)                  ' 結晶番号
        udt_WFSmp(2).INGOTPOS = typ_CType.typ_Param.WFSMP(SxlTail039).INPOSCW   ' 結晶内開始位置
        udt_WFSmp(2).SMPKBN = typ_CType.typ_Param.WFSMP(SxlTail039).SMPKBNCW    ' サンプル区分

        'WFサンプル管理更新
'        If DBDRV_scmzc_fcmlc001c_UpdWfCrySmp(udt_WFSmp) <> FUNCTION_RETURN_SUCCESS Then
        '引数変更 09/05/25 ooba
        If DBDRV_scmzc_fcmlc001c_UpdWfCrySmp(NtoS(typ_CType.typ_Param.SXLID)) <> FUNCTION_RETURN_SUCCESS Then
            f_cmbc039_2.lblMsg.Caption = GetMsgStr("EGET2", "E044")
            RegWfSogoRsltOK = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
    End If
    
End Function

'*******************************************************************************
'*    関数名        : NtoS
'*
'*    処理概要      : 1.入力された値がNullならスペースを返す
'*                    2. 1以外ならString型として返す
'*
'*    パラメータ    : 変数名        ,IO ,型           ,説明
'*                    sWk         ,I  ,String       ,Work
'*
'*    戻り値        : String
'*
'*******************************************************************************
Public Function NtoS(sWk As String) As String
    If Mid(sWk, 1, 1) = Chr(0) Then
        NtoS = " "
        Exit Function
    End If
    NtoS = sWk
End Function

'*******************************************************************************
'*    関数名        : NtoZ2
'*
'*    処理概要      : 1.入力された値が-1を返す
'*                    2. 1以外ならDouble型として返す
'*
'*    パラメータ    : 変数名        ,IO ,型           ,説明
'*                    sWk         ,I  ,String       ,Work
'*
'*    戻り値        : Double
'*
'*******************************************************************************
Public Function NtoZ2(sWk As String) As Double
    If Trim(sWk) = "" Then
        NtoZ2 = -1
        Exit Function
    End If
    NtoZ2 = CDbl(sWk)
End Function

'*******************************************************************************
'*    関数名        : StaffIDCheck
'*
'*    処理概要      : 1.担当者コードより担当者名を取得
'*
'*    パラメータ    : 変数名        ,IO ,型           ,説明
'*                    txtStaffID    ,IO ,TextBox      ,担当者コードTextBox
'*                    txtJfName     ,IO ,TextBox      ,担当者名TextBox
'*                    lblMsg        ,O  ,Label        ,メッセージ出力Label
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function StaffIDCheck(txtStaffID As TextBox, txtJfName As TextBox, lblMsg As Label) As FUNCTION_RETURN
    Dim sStaffName    As String

    '' 担当者コードの長さを確認
    If ChkTextBox(txtStaffID, CHK_STRING, 7, 7) = FUNCTION_RETURN_SUCCESS Then
        '' 担当者コードより担当者名を取得
        BeginProcess '' プロセス開始
        lblMsg.Caption = GetMsgStr(PWAIT)
        sStaffName = GetStaffName(Trim(txtStaffID.text))

        If sStaffName <> vbNullString Then
            '' 担当者名を表示
            STAFFIDBUFF = Trim(txtStaffID.text)
            lblMsg.Caption = ""             'クリア位置変更
            txtJfName.text = sStaffName
'            lblMsg.Caption = ""            '流動停止メッセージ表示対応 コメントして上へ移動
            StaffIDCheck = FUNCTION_RETURN_SUCCESS
        Else
            '' 該当担当者コードが見つからない場合、エラーメッセージ表示
            lblMsg.Caption = GetMsgStr(ESTAF)
            StaffIDCheck = FUNCTION_RETURN_FAILURE
        End If
    Else
        ''担当者コード異常
        BeginProcess '' プロセス開始
        lblMsg.Caption = GetMsgStr(ESTAF)
        StaffIDCheck = FUNCTION_RETURN_FAILURE
    End If

    EndProcess '' プロセス終了
End Function

'*******************************************************************************
'*    関数名        : GetSampleID
'*
'*    処理概要      : 1.抜試指示サンプルのサンプルＩＤを取得する
'*
'*    パラメータ    : 変数名       ,IO ,型      ,説明
'*                    intWafPos    ,I  ,Integer ,抜試指示テーブル位置
'*                    sSampID1     ,I  ,String  ,サンプルＩＤ１
'*                    sSampID2     ,I  ,String  ,サンプルＩＤ２
'*
'*    戻り値        : 選択ありの場合はTrue, 選択なしの場合はFalse
'*
'*******************************************************************************
Public Function GetSampleID(intWafPos As Integer, sSampID1 As String, sSampID2 As String, _
                                                     Optional intKubun As Integer) As Boolean
    Dim blBot           As Boolean
    Dim blTop           As Boolean
    Dim blBlk           As Boolean
    Dim intTargetBlkPos As Integer
    Dim p               As Integer
    Dim m               As Integer
    Dim i               As Integer
    Dim intHinbanRow    As Integer
    Dim vUpHinban       As Variant

    blBot = False
    blTop = False
    blBlk = False
    p = intWafPos

    With tblWafInd(intWafPos)
        m = UBound(tblBlkInf)

        For i = 1 To UBound(tblBlkInf)
            If tblWafInd(p).BLOCKID = tblBlkInf(i).BLOCKID Then
                intTargetBlkPos = i
                Exit For
            End If
        Next

        blBot = False
        blTop = False
        Call GetSampleBT(.SMP.CRYINDRS, blTop, blBot)
        Call GetSampleBT(.SMP.CRYINDOI, blTop, blBot)
        Call GetSampleBT(.SMP.CRYINDB1, blTop, blBot)
        Call GetSampleBT(.SMP.CRYINDB2, blTop, blBot)
        Call GetSampleBT(.SMP.CRYINDB3, blTop, blBot)
        Call GetSampleBT(.SMP.CRYINDL1, blTop, blBot)
        Call GetSampleBT(.SMP.CRYINDL2, blTop, blBot)
        Call GetSampleBT(.SMP.CRYINDL3, blTop, blBot)
        Call GetSampleBT(.SMP.CRYINDL4, blTop, blBot)
        Call GetSampleBT(.SMP.CRYINDDS, blTop, blBot)
        Call GetSampleBT(.SMP.CRYINDDZ, blTop, blBot)
        Call GetSampleBT(.SMP.CRYINDSP, blTop, blBot)
        Call GetSampleBT(.SMP.CRYINDD1, blTop, blBot)
        Call GetSampleBT(.SMP.CRYINDD2, blTop, blBot)
        Call GetSampleBT(.SMP.CRYINDD3, blTop, blBot)
        Call GetSampleBT(.SMP.CRYINDAO, blTop, blBot)     '残存酸素追加
        Call GetSampleBT(.SMP.CRYOTHER1, blTop, blBot)    'その他ｻﾝﾌﾟﾙ1追加
        Call GetSampleBT(.SMP.CRYOTHER2, blTop, blBot)    'その他ｻﾝﾌﾟﾙ2追加
        Call GetSampleBT(.SMP.CRYINDGD, blTop, blBot)     'GD追加

        'エピ先行評価追加対応
        Call GetSampleBT(.SMP.EPIINDB1, blTop, blBot)
        Call GetSampleBT(.SMP.EPIINDB2, blTop, blBot)
        Call GetSampleBT(.SMP.EPIINDB3, blTop, blBot)
        Call GetSampleBT(.SMP.EPIINDL1, blTop, blBot)
        Call GetSampleBT(.SMP.EPIINDL2, blTop, blBot)
        Call GetSampleBT(.SMP.EPIINDL3, blTop, blBot)

        '上下品番がZ
        If intWafPos >= 1 Then
            If Trim(tblWafInd(intWafPos).HINDN.hinban) = "Z" Or _
               Trim(tblWafInd(intWafPos).HINUP.hinban) = "Z" Or _
               intKubun = 3 Then      'ブロックが変わる初期表示行
                    blTop = True
                    blBot = True
                    blBlk = False

                    'チェックボックス追加によりサンプル切替の判定を追加 (intWafPos - 1を使用するため区分３のときのみ判定する) 2003/06/01 okazaki
                    If Trim(tblWafInd(intWafPos).HINDN.hinban) <> "Z" And tblWafInd(intWafPos).HINDN.hinban = tblWafInd(intWafPos - 1).HINDN.hinban Then
                        blTop = False
                        blBot = False
                        blBlk = False
                    End If
            End If
        End If

        '' 上方向／下方向サンプル（別）
        If blTop = True And blBot = True Then
            If blBlk = True Then
                If .BlockPos = 0 Then
                    If tblBlkInf(intTargetBlkPos - 1).NOWPROC = PROCD_WFC_SOUGOUHANTEI Then
                        sSampID1 = Right(tblBlkInf(intTargetBlkPos - 1).BLOCKID, 3) & "-" & GetWafPos(tblBlkInf(intTargetBlkPos - 1).LENGTH) & "B"
                        sSampID2 = Right(.BLOCKID, 3) & "-000T"
                    Else
                        sSampID1 = Right(.BLOCKID, 3) & "-000T"
                        sSampID2 = ""
                    End If
                Else
                    sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "B"
                    sSampID2 = Right(tblBlkInf(intTargetBlkPos + 1).BLOCKID, 3) & "-000T"
                End If
            Else
                sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "U"
                sSampID2 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "D"
            End If
            GetSampleID = False
        '' 下方向サンプル
        ElseIf blTop = True And blBot = False Then
            If blBlk = True Then
                sSampID1 = Right(.BLOCKID, 3) & "-000T"
            Else
                sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "D"
            End If
            sSampID2 = ""
            GetSampleID = False
        '' 上方向サンプル
        ElseIf blTop = False And blBot = True Then
            If blBlk = True Then
                If .BlockPos = 0 Then
                    sSampID1 = Right(tblBlkInf(intTargetBlkPos - 1).BLOCKID, 3) & "-" & GetWafPos(tblBlkInf(intTargetBlkPos - 1).LENGTH) & "B"
                Else
                    sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "B"
                End If
            Else
                sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "U"
            End If
            sSampID2 = ""
            GetSampleID = False

        '' 上方向／下方向サンプル（共通）
        ElseIf blTop = False And blBot = False Then
            If blBlk = True Then
                If .BlockPos = 0 Then
                    If tblBlkInf(intTargetBlkPos).NOWPROC = PROCD_WFC_SOUGOUHANTEI Then
                        sSampID1 = Right(tblBlkInf(intTargetBlkPos).BLOCKID, 3) & "-" & GetWafPos(tblBlkInf(intTargetBlkPos - 1).LENGTH) & "B"
                        sSampID2 = Right(.BLOCKID, 3) & "-000T"
                    Else
                        sSampID1 = Right(.BLOCKID, 3) & "-000T"
                        sSampID2 = ""
                        GetSampleID = False
                        Exit Function
                    End If
                Else
                    If tblBlkInf(intTargetBlkPos + 1).NOWPROC = PROCD_WFC_SOUGOUHANTEI Then
                        sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "B"
                        sSampID2 = Right(tblBlkInf(intTargetBlkPos + 1).BLOCKID, 3) & "-000T"
                    Else
                        sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "B"
                        sSampID2 = ""
                        GetSampleID = False
                        Exit Function
                    End If
                End If
            Else
                sSampID1 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "U"
                sSampID2 = Right(.BLOCKID, 3) & "-" & GetWafPos(.BlockPos) & "D"
            End If
            GetSampleID = True
        End If
    End With
End Function

'*******************************************************************************
'*    関数名        : GetSampleBT
'*
'*    処理概要      : 1.サンプルのトップ側／ボトム側区分の取得
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*　　　　　　　　　　sSample       ,I  ,String 　,サンプル
'*                    blTop         ,O  ,Boolean　,トップ側区分の有無
'*                    blBot         ,O  ,Boolean　,ボトム側区分の有無
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Sub GetSampleBT(ByVal sSample As String, blTop As Boolean, blBot As Boolean)
    Select Case sSample
        Case "1"
            blTop = True
        Case "2"
            blBot = True
        Case "4"
            blTop = True
            blBot = True
    End Select
End Sub

'*******************************************************************************
'*    関数名        : LackMapMake
'*
'*    処理概要      : 1.サンプルのトップ側／ボトム側区分の取得
'*
'*    パラメータ    : 変数名        ,IO ,型              ,説明
'*　　　　　　　　　　udtBlkInf     ,I  ,typ_BlkInf3   　,ブロック管理構造体
'*                    udtTmpLackWaf ,I  ,typ_LackWaf　   ,欠落情報
'*                    intBlkInfPos  ,I  ,Integer         ,結晶内全体のブロック数に対する対象ブロックの開始位置
'*                    intBlkCnt     ,I  ,Integer    　   ,対象ブロック数
'*                    RftblLackMap  ,O  ,typ_LackMap     ,ウェハーテーブル構造体
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function LackMapMake(udtBlkInf() As typ_BlkInf3, udtTmpLackWaf() As typ_LackWaf, intBlkInfPos As Integer, intBlkCnt As Integer) As FUNCTION_RETURN
    Dim blFlag  As Boolean
    Dim p       As Integer
    Dim m       As Integer
    Dim n       As Integer
    Dim i       As Integer
    Dim j       As Integer
    Dim k       As Integer

    '' 欠落ウェハーテーブルの作成
    k = 0
    m = intBlkCnt + intBlkInfPos - 1
    n = UBound(udtTmpLackWaf)
    ReDim tblLackMap(n)

    '' ブロックの始まりから
    For i = intBlkInfPos To m
        DoEvents
        For j = 1 To n
            DoEvents
            If udtBlkInf(i).BLOCKID = udtTmpLackWaf(j).BLOCKID Then
                If blFlag = False Then
                    k = k + 1
                    tblLackMap(k).BLOCKID = udtTmpLackWaf(j).BLOCKID
                    p = udtTmpLackWaf(j).WAFERNO
                    If p = -1 Then
                        tblLackMap(k).LACKPOSS = 0
                        tblLackMap(k).LACKCNTS = -1
                        tblLackMap(k).LACKPOSE = udtBlkInf(i).REALLEN
                        tblLackMap(k).LACKCNTE = -1
                        Exit For
                    End If
                    tblLackMap(k).LACKPOSS = udtTmpLackWaf(j).TOP_POS
                    tblLackMap(k).LACKCNTS = udtTmpLackWaf(j).WAFERNO
                    blFlag = True
                Else
                    If udtTmpLackWaf(j).WAFERNO = p + 1 Then
                        p = p + 1
                        If blFlag = True And j = n Then
                            tblLackMap(k).LACKPOSE = udtTmpLackWaf(j).TAIL_POS
                            tblLackMap(k).LACKCNTE = udtTmpLackWaf(j).WAFERNO
                        End If
                    Else
                        tblLackMap(k).LACKPOSE = udtTmpLackWaf(j - 1).TAIL_POS
                        tblLackMap(k).LACKCNTE = udtTmpLackWaf(j - 1).WAFERNO
                        k = k + 1
                        tblLackMap(k).BLOCKID = udtTmpLackWaf(j).BLOCKID
                        tblLackMap(k).LACKPOSS = udtTmpLackWaf(j).TOP_POS
                        tblLackMap(k).LACKCNTS = udtTmpLackWaf(j).WAFERNO
                        p = udtTmpLackWaf(j).WAFERNO
                    End If
                End If
            Else
                If blFlag = True Then
                    tblLackMap(k).LACKPOSE = udtTmpLackWaf(j - 1).TAIL_POS
                    tblLackMap(k).LACKCNTE = udtTmpLackWaf(j - 1).WAFERNO
                    blFlag = False
                    Exit For
                End If
            End If
        Next j
    Next i
    ReDim Preserve tblLackMap(k)

    For i = 1 To k
        With tblLackMap(i)
            If .LACKPOSS > 0 And .LACKPOSE = 0 Then
                .LACKPOSE = .LACKPOSS
            End If
            If .LACKCNTS > 0 And .LACKCNTE = 0 Then
                .LACKCNTE = .LACKCNTS
            End If
        End With
    Next
End Function

'*******************************************************************************
'*    関数名        : NoTestCheck
'*
'*    処理概要      : 使ってないからコメント無し
'*
'*    パラメータ    : 変数名        ,IO ,型              ,説明
'*
'*    戻り値        :
'*
'*******************************************************************************
Public Function NoTestCheck(lblMsg As Label) As FUNCTION_RETURN
    Dim c0          As Long
    Dim udtHin(1)   As tFullHinban
    Dim udtInf(1)      As NoTest_Info

    NoTestCheck = FUNCTION_RETURN_FAILURE

    For c0 = 1 To 2
        '元品番セット
        udtHin(0).factory = tblTotal.typ_Param.factory
        udtHin(0).hinban = tblTotal.typ_Param.hinban
        udtHin(0).mnorevno = tblTotal.typ_Param.REVNUM
        udtHin(0).opecond = tblTotal.typ_Param.opecond
        '振替先品番セット
        If c0 = 1 Then
            udtHin(1) = tblWafInd(1).HINDN
        Else
            udtHin(1) = tblWafInd(UBound(tblWafInd())).HINUP
        End If
        If Trim(udtHin(1).hinban) = "Z" Then
            Exit For
        End If
        If DBDRV_GetNoTestHinInfo(udtHin(), udtInf()) = FUNCTION_RETURN_FAILURE Then
            Exit Function
        End If

        If (tblTotal.typ_Param.WFSMP(c0).WFRESRS1CW <> "1") Then
        '実績無
            'ＷＦサンプル処理変更
            If (udtInf(1).Res.HWFRHWYS = "H") Or (udtInf(1).Res.HWFRHWYS = "S") Then

            '検査有り
                lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " RES実績無")
                Exit Function
            End If
        End If

        If (tblTotal.typ_Param.WFSMP(c0).WFRESOICW <> "1") Then
        '実績無
            'ＷＦサンプル処理変更
            If (udtInf(1).Oi.HWFONHWS = "H") Or (udtInf(1).Oi.HWFONHWS = "S") Then

            '検査有り
                lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " OI実績無")  '03/06/06 後藤
                Exit Function
            End If
        End If
        If (tblTotal.typ_Param.WFSMP(c0).WFRESB1CW = "1") Then
            '実績有り
            'ＷＦサンプル処理変更
            If (udtInf(1).BMD(0).HWFBMxHS = "H") Or (udtInf(1).BMD(0).HWFBMxHS = "S") Then
                '検査有り
                If udtInf(0).BMD(0).HWFBMxET <> udtInf(1).BMD(0).HWFBMxET Or _
                   udtInf(0).BMD(0).HWFBMxNS <> udtInf(1).BMD(0).HWFBMxNS Or _
                   udtInf(0).BMD(0).HWFBMxSH <> udtInf(1).BMD(0).HWFBMxSH Or _
                   udtInf(0).BMD(0).HWFBMxSR <> udtInf(1).BMD(0).HWFBMxSR Or _
                   udtInf(0).BMD(0).HWFBMxST <> udtInf(1).BMD(0).HWFBMxST Or _
                   udtInf(0).BMD(0).HWFBMxSZ <> udtInf(1).BMD(0).HWFBMxSZ Then
                    lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " BMD1")  '03/06/06
                    Exit Function
                End If
            End If
        Else
            '実績無
            'ＷＦサンプル処理変更
            If (udtInf(1).BMD(0).HWFBMxHS = "H") Or (udtInf(1).BMD(0).HWFBMxHS = "S") Then
                '検査有り
                lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " BMD1実績無")  '03/06/06 後藤
                Exit Function
            End If
        End If

        If (tblTotal.typ_Param.WFSMP(c0).WFRESB2CW = "1") Then
            '実績有り
            'ＷＦサンプル処理変更
            If (udtInf(1).BMD(1).HWFBMxHS = "H") Or (udtInf(1).BMD(1).HWFBMxHS = "S") Then
                '検査有り
                If udtInf(0).BMD(1).HWFBMxET <> udtInf(1).BMD(1).HWFBMxET Or _
                   udtInf(0).BMD(1).HWFBMxNS <> udtInf(1).BMD(1).HWFBMxNS Or _
                   udtInf(0).BMD(1).HWFBMxSH <> udtInf(1).BMD(1).HWFBMxSH Or _
                   udtInf(0).BMD(1).HWFBMxSR <> udtInf(1).BMD(1).HWFBMxSR Or _
                   udtInf(0).BMD(1).HWFBMxST <> udtInf(1).BMD(1).HWFBMxST Or _
                   udtInf(0).BMD(1).HWFBMxSZ <> udtInf(1).BMD(1).HWFBMxSZ Then
                    lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " BMD2")
                    Exit Function
                End If
            End If
        Else
            'ＷＦサンプル処理変更
            If (udtInf(1).BMD(1).HWFBMxHS = "H") Or (udtInf(1).BMD(1).HWFBMxHS = "S") Then
                '検査有り
                lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " BMD2実績無")
                Exit Function
            End If
        End If

        If (tblTotal.typ_Param.WFSMP(c0).WFRESB3CW = "1") Then
            '実績有り
            'ＷＦサンプル処理変更
            If (udtInf(1).BMD(2).HWFBMxHS = "H") Or (udtInf(1).BMD(2).HWFBMxHS = "S") Then
                '検査有り
                If udtInf(0).BMD(2).HWFBMxET <> udtInf(1).BMD(2).HWFBMxET Or _
                   udtInf(0).BMD(2).HWFBMxNS <> udtInf(1).BMD(2).HWFBMxNS Or _
                   udtInf(0).BMD(2).HWFBMxSH <> udtInf(1).BMD(2).HWFBMxSH Or _
                   udtInf(0).BMD(2).HWFBMxSR <> udtInf(1).BMD(2).HWFBMxSR Or _
                   udtInf(0).BMD(2).HWFBMxST <> udtInf(1).BMD(2).HWFBMxST Or _
                   udtInf(0).BMD(2).HWFBMxSZ <> udtInf(1).BMD(2).HWFBMxSZ Then
                    lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " BMD3")
                    Exit Function
                End If
            End If
        Else
            'ＷＦサンプル処理変更
            If (udtInf(1).BMD(2).HWFBMxHS = "H") Or (udtInf(1).BMD(2).HWFBMxHS = "S") Then
                '検査有り
                lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " BMD3実績無")
                Exit Function
            End If
        End If
        If (tblTotal.typ_Param.WFSMP(c0).WFRESL1CW = "1") Then
            '実績有り
            'ＷＦサンプル処理変更
            If (udtInf(1).OSF(0).HWFOFxHS = "H") Or (udtInf(1).OSF(0).HWFOFxHS = "S") Then
                '検査有り
                If udtInf(0).OSF(0).HWFOFxET <> udtInf(1).OSF(0).HWFOFxET Or _
                   udtInf(0).OSF(0).HWFOFxNS <> udtInf(1).OSF(0).HWFOFxNS Or _
                   udtInf(0).OSF(0).HWFOFxSH <> udtInf(1).OSF(0).HWFOFxSH Or _
                   udtInf(0).OSF(0).HWFOFxSR <> udtInf(1).OSF(0).HWFOFxSR Or _
                   udtInf(0).OSF(0).HWFOFxST <> udtInf(1).OSF(0).HWFOFxST Or _
                   udtInf(0).OSF(0).HWFOFxSZ <> udtInf(1).OSF(0).HWFOFxSZ Then
                    lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " OSF1")
                    Exit Function
                End If
            End If
        Else
            'ＷＦサンプル処理変更
            If (udtInf(1).OSF(0).HWFOFxHS = "H") Or (udtInf(1).OSF(0).HWFOFxHS = "S") Then
                '検査有り
                lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " OSF1実績無")
                Exit Function
            End If
        End If
        If (tblTotal.typ_Param.WFSMP(c0).WFRESL2CW = "1") Then
            '実績有り
            'ＷＦサンプル処理変更
            If (udtInf(1).OSF(1).HWFOFxHS = "H") Or (udtInf(1).OSF(1).HWFOFxHS = "S") Then
                '検査有り
                If udtInf(0).OSF(1).HWFOFxET <> udtInf(1).OSF(1).HWFOFxET Or _
                   udtInf(0).OSF(1).HWFOFxNS <> udtInf(1).OSF(1).HWFOFxNS Or _
                   udtInf(0).OSF(1).HWFOFxSH <> udtInf(1).OSF(1).HWFOFxSH Or _
                   udtInf(0).OSF(1).HWFOFxSR <> udtInf(1).OSF(1).HWFOFxSR Or _
                   udtInf(0).OSF(1).HWFOFxST <> udtInf(1).OSF(1).HWFOFxST Or _
                   udtInf(0).OSF(1).HWFOFxSZ <> udtInf(1).OSF(1).HWFOFxSZ Then
                    lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " OSF2")
                    Exit Function
                End If
            End If
        Else
            'ＷＦサンプル処理変更
            If (udtInf(1).OSF(1).HWFOFxHS = "H") Or (udtInf(1).OSF(1).HWFOFxHS = "S") Then
                '検査有り
                lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " OSF2実績無")
                Exit Function
            End If
        End If

        If (tblTotal.typ_Param.WFSMP(c0).WFRESL3CW = "1") Then
            '実績有り
            'ＷＦサンプル処理変更
            If (udtInf(1).OSF(2).HWFOFxHS = "H") Or (udtInf(1).OSF(2).HWFOFxHS = "S") Then
                '検査有り
                If udtInf(0).OSF(2).HWFOFxET <> udtInf(1).OSF(2).HWFOFxET Or _
                   udtInf(0).OSF(2).HWFOFxNS <> udtInf(1).OSF(2).HWFOFxNS Or _
                   udtInf(0).OSF(2).HWFOFxSH <> udtInf(1).OSF(2).HWFOFxSH Or _
                   udtInf(0).OSF(2).HWFOFxSR <> udtInf(1).OSF(2).HWFOFxSR Or _
                   udtInf(0).OSF(2).HWFOFxST <> udtInf(1).OSF(2).HWFOFxST Or _
                   udtInf(0).OSF(2).HWFOFxSZ <> udtInf(1).OSF(2).HWFOFxSZ Then
                    lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " OSF3")
                    Exit Function
                End If
            End If
        Else
            'ＷＦサンプル処理変更
            If (udtInf(1).OSF(2).HWFOFxHS = "H") Or (udtInf(1).OSF(2).HWFOFxHS = "S") Then
                '検査有り
                lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " OSF3実績無")
                Exit Function
            End If
        End If
        If (tblTotal.typ_Param.WFSMP(c0).WFRESL4CW = "1") Then
            '実績有り
            'ＷＦサンプル処理変更
            If (udtInf(1).OSF(3).HWFOFxHS = "H") Or (udtInf(1).OSF(3).HWFOFxHS = "S") Then
                '検査有り
                If udtInf(0).OSF(3).HWFOFxET <> udtInf(1).OSF(3).HWFOFxET Or _
                   udtInf(0).OSF(3).HWFOFxNS <> udtInf(1).OSF(3).HWFOFxNS Or _
                   udtInf(0).OSF(3).HWFOFxSH <> udtInf(1).OSF(3).HWFOFxSH Or _
                   udtInf(0).OSF(3).HWFOFxSR <> udtInf(1).OSF(3).HWFOFxSR Or _
                   udtInf(0).OSF(3).HWFOFxST <> udtInf(1).OSF(3).HWFOFxST Or _
                   udtInf(0).OSF(3).HWFOFxSZ <> udtInf(1).OSF(3).HWFOFxSZ Then
                    lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " OSF4")
                    Exit Function
                End If
            End If
        Else
            'ＷＦサンプル処理変更
            If (udtInf(1).OSF(3).HWFOFxHS = "H") Or (udtInf(1).OSF(3).HWFOFxHS = "S") Then
                '検査有り
                lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " OSF4実績無")
                Exit Function
            End If
        End If
        If (tblTotal.typ_Param.WFSMP(c0).WFRESDSCW = "1") Then
            '実績有り
            'ＷＦサンプル処理変更
            If (udtInf(1).Dsod.HWFDSOHS = "H") Or (udtInf(1).Dsod.HWFDSOHS = "S") Then
                '検査有り
                If udtInf(0).Dsod.HWFDSOKE <> udtInf(1).Dsod.HWFDSOKE Then
                    lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " DSOD")
                    Exit Function
                End If
            End If
        Else
            'ＷＦサンプル処理変更
            If (udtInf(1).Dsod.HWFDSOHS = "H") Or (udtInf(1).Dsod.HWFDSOHS = "S") Then
                '検査有り
                lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " DSOD実績無")
                Exit Function
            End If
        End If

        If (tblTotal.typ_Param.WFSMP(c0).WFRESDZCW = "1") Then
            '実績有り
            'ＷＦサンプル処理変更
            If (udtInf(1).DZ.HWFMKHWS = "H") Or (udtInf(1).DZ.HWFMKHWS = "S") Then
                'ＷＦサンプル処理変更
                '検査有り
                If udtInf(0).DZ.HWFMKSPH <> udtInf(1).DZ.HWFMKSPH Or _
                   udtInf(0).DZ.HWFMKSPR <> udtInf(1).DZ.HWFMKSPR Or _
                   udtInf(0).DZ.HWFMKSPT <> udtInf(1).DZ.HWFMKSPT Or _
                   udtInf(0).DZ.HWFMKSZY <> udtInf(1).DZ.HWFMKSZY Then
                    lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " DZ")
                    Exit Function
                End If
            End If
        Else
            'ＷＦサンプル処理変更
            If (udtInf(1).DZ.HWFMKHWS = "H") Or (udtInf(1).DZ.HWFMKHWS = "S") Then
                '検査有り
                lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " DZ実績無")
                Exit Function
            End If
        End If
        If (tblTotal.typ_Param.WFSMP(c0).WFRESSPCW = "1") Then
            '実績有り
            'ＷＦサンプル処理変更
            If (udtInf(1).SpvFe.HWFSPVHS = "H") Or (udtInf(1).SpvFe.HWFSPVHS = "S") Then
                '検査有り
                If udtInf(0).SpvFe.HWFSPVSH <> udtInf(1).SpvFe.HWFSPVSH Or _
                   udtInf(0).SpvFe.HWFSPVSI <> udtInf(1).SpvFe.HWFSPVSI Or _
                   udtInf(0).SpvFe.HWFSPVST <> udtInf(1).SpvFe.HWFSPVST Then
                    lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " SPVFE")
                    Exit Function
                End If
            End If
        Else
            'ＷＦサンプル処理変更
            If (udtInf(1).SpvFe.HWFSPVHS = "H") Or (udtInf(1).SpvFe.HWFSPVHS = "S") Then
                '検査有り
                lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " SPVFE実績無")
                Exit Function
            End If
        End If
        If (tblTotal.typ_Param.WFSMP(c0).WFRESSPCW = "1") Then
            '実績有り
            'ＷＦサンプル処理変更
            If (udtInf(1).Spv.HWFDLHWS = "H") Or (udtInf(1).Spv.HWFDLHWS = "S") Then
                '検査有り
                If udtInf(0).Spv.HWFDLSPH <> udtInf(1).Spv.HWFDLSPH Or _
                   udtInf(0).Spv.HWFDLSPI <> udtInf(1).Spv.HWFDLSPI Or _
                   udtInf(0).Spv.HWFDLSPT <> udtInf(1).Spv.HWFDLSPT Then
                    lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " SPV拡散長")
                    Exit Function
                End If
            End If
        Else
            'ＷＦサンプル処理変更
            If (udtInf(1).Spv.HWFDLHWS = "H") Or (udtInf(1).Spv.HWFDLHWS = "S") Then
                '検査有り
                lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " SPV拡散長実績無")
                Exit Function
            End If
        End If

        If (tblTotal.typ_Param.WFSMP(c0).WFRESDO1CW = "1") Then
            '実績有り
            'ＷＦサンプル処理変更
            If (udtInf(1).Doi(0).HWFOSxHS = "H") Or (udtInf(1).Doi(0).HWFOSxHS = "S") Then
                'ＷＦサンプル処理変更
                '検査有り
                If udtInf(0).Doi(0).HWFOSxNS <> udtInf(1).Doi(0).HWFOSxNS Or _
                   udtInf(0).Doi(0).HWFOSxSH <> udtInf(1).Doi(0).HWFOSxSH Or _
                   udtInf(0).Doi(0).HWFOSxSI <> udtInf(1).Doi(0).HWFOSxSI Or _
                   udtInf(0).Doi(0).HWFOSxST <> udtInf(1).Doi(0).HWFOSxST Then
                    lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " �儖i1")
                    Exit Function
                End If
            End If
        Else
            'ＷＦサンプル処理変更
            If (udtInf(1).Doi(0).HWFOSxHS = "H") Or (udtInf(1).Doi(0).HWFOSxHS = "S") Then
                '検査有り
                lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " �儖i1実績無")
                Exit Function
            End If
        End If
        If (tblTotal.typ_Param.WFSMP(c0).WFRESDO2CW = "1") Then
            '実績有り
            'ＷＦサンプル処理変更
            If (udtInf(1).Doi(1).HWFOSxHS = "H") Or (udtInf(1).Doi(1).HWFOSxHS = "S") Then
                '検査有り
                If udtInf(0).Doi(0).HWFOSxNS <> udtInf(1).Doi(0).HWFOSxNS Or _
                   udtInf(0).Doi(0).HWFOSxSH <> udtInf(1).Doi(0).HWFOSxSH Or _
                   udtInf(0).Doi(0).HWFOSxSI <> udtInf(1).Doi(0).HWFOSxSI Or _
                   udtInf(0).Doi(0).HWFOSxST <> udtInf(1).Doi(0).HWFOSxST Then
                    lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " �儖i2")
                    Exit Function
                End If
            End If
        Else
            'ＷＦサンプル処理変更
            If (udtInf(1).Doi(1).HWFOSxHS = "H") Or (udtInf(1).Doi(1).HWFOSxHS = "S") Then
                '検査有り
                lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " �儖i2実績無")
                Exit Function
            End If
        End If
        If (tblTotal.typ_Param.WFSMP(c0).WFRESDO3CW = "1") Then
            '実績有り
            'ＷＦサンプル処理変更
            If (udtInf(1).Doi(2).HWFOSxHS = "H") Or (udtInf(1).Doi(2).HWFOSxHS = "S") Then
                '検査有り
                If udtInf(0).Doi(0).HWFOSxNS <> udtInf(1).Doi(0).HWFOSxNS Or _
                   udtInf(0).Doi(0).HWFOSxSH <> udtInf(1).Doi(0).HWFOSxSH Or _
                   udtInf(0).Doi(0).HWFOSxSI <> udtInf(1).Doi(0).HWFOSxSI Or _
                   udtInf(0).Doi(0).HWFOSxST <> udtInf(1).Doi(0).HWFOSxST Then
                    lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " �儖i3")
                    Exit Function
                End If
            End If
        Else
            'ＷＦサンプル処理変更
            If (udtInf(1).Doi(2).HWFOSxHS = "H") Or (udtInf(1).Doi(2).HWFOSxHS = "S") Then
                '検査有り
                lblMsg.Caption = GetMsgStr("EHINC", Trim(udtHin(1).hinban) & " " & " �儖i3実績無")
                Exit Function
            End If
        End If
    Next

    NoTestCheck = FUNCTION_RETURN_SUCCESS
End Function

'*******************************************************************************
'*    関数名        : LackMapMake
'*
'*    処理概要      : 1.SXL_IDの取得
'*
'*    パラメータ    : 変数名        ,IO ,型              ,説明
'*　　　　　　　　　　sBlockID 　　 ,I  ,String 　       ,ブロックID
'*                    intIngotpos   ,I  ,Integer    　   ,結晶内開始位置
'*
'*    戻り値        : String(SXL_ID)
'*
'*******************************************************************************
Public Function GetSXLID(sBlockId As String, intIngotpos As Integer) As String
    GetSXLID = left(sBlockId, 10) & GetWafPos(intIngotpos)
End Function

'*******************************************************************************
'*    関数名        : GetWafPos
'*
'*    処理概要      : 1.抜試位置文字列の取得
'*
'*    パラメータ    : 変数名        ,IO ,型              ,説明
'*　　　　　　　　　　intIngotpos   ,I  ,Integer 　      ,結晶内開始位置
'*
'*    戻り値        : String(抜試位置文字列)
'*
'*******************************************************************************
Public Function GetWafPos(intIngotpos As Integer) As String
    Dim i As Integer
    Dim j As Integer

    If intIngotpos >= 1000 Then
        i = Int(intIngotpos / 100)
        j = intIngotpos Mod 100
        GetWafPos = Chr$(i - 10 + Asc("A")) & Format(j, "00")
    Else
        GetWafPos = Format(intIngotpos, "000")
    End If
End Function

'*******************************************************************************
'*    関数名        : MakeMesIndTbl
'*
'*    処理概要      : 1.測定評価方法指示テーブルの作成
'*
'*    パラメータ    : 変数名        ,IO ,型              ,説明
'*　　　　　　　　　　udtPSXLMng    ,I  ,typ_TBCME042 　 ,SXL管理
'*                    udtPWafSmp    ,I  ,typ_XSDCW  　   ,新サンプル管理（SXL）
'*                    udtPMesInd    ,I  ,typ_TBCMY003  　,測定評価方法指示
'*                    udtPEpMesInd  ,I  ,typ_TBCMY020    ,EP測定評価指示
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function MakeMesIndTbl(udtPSXLMng() As typ_TBCME042, udtPWafSmp() As typ_XSDCW, _
                        udtPMesInd() As typ_TBCMY003, udtPEpMesInd() As typ_TBCMY020) As FUNCTION_RETURN
    Dim udtTmpSpWFSamp()    As typ_SpWFSamp
    Dim sHin                As String
    Dim sDKAN               As String
    Dim m                   As Integer
    Dim n                   As Integer
    Dim i                   As Integer
    Dim j                   As Integer
    Dim k                   As Integer
    Dim sGdSpec             As String       '規格値(GD)

    'エピ先行評価追加対応
    '' エピ先行評価項目用のDKアニール条件
    Dim sDKAN_EP            As String
    Dim l                   As Integer

    '' 測定評価方法指示用の製品仕様を取得
    j = 0
    m = UBound(udtPSXLMng)
    ReDim udtTmpSpWFSamp(m)
    For i = 1 To m
            sHin = RTrim$(udtPSXLMng(i).hinban)
        If (sHin <> "" And sHin <> "G" And sHin <> "Z") Then
            j = j + 1
            udtTmpSpWFSamp(j).HIN.hinban = udtPSXLMng(i).hinban
            udtTmpSpWFSamp(j).HIN.mnorevno = udtPSXLMng(i).REVNUM
            udtTmpSpWFSamp(j).HIN.factory = udtPSXLMng(i).factory
            udtTmpSpWFSamp(j).HIN.opecond = udtPSXLMng(i).opecond
            If scmzc_getWF(udtTmpSpWFSamp(j)) = FUNCTION_RETURN_FAILURE Then
                MakeMesIndTbl = FUNCTION_RETURN_FAILURE
                Exit Function
            End If
        End If
    Next i
    ReDim Preserve udtTmpSpWFSamp(j)

    '' 測定評価方法指示テーブルの作成
    k = 0
    m = UBound(udtPWafSmp)
    n = UBound(udtTmpSpWFSamp)

    ReDim udtPMesInd(m * 18)   'OTH2を削除 エピ先行評価追加対応

    'エピ先行評価追加対応
    l = 0
    ReDim udtPEpMesInd(m * 7)  ' OTH2、OSF1E〜OSF3E、BMD1E〜BMD3E

    For i = 1 To m
        For j = 1 To n
            If udtTmpSpWFSamp(j).HIN.hinban = udtPWafSmp(i).HINBCW Then
                Exit For
            End If
        Next j
        If j <= n Then
            With udtTmpSpWFSamp(j)
                '◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
                'エピ先行評価追加対応
'''                If .HWFMKCET = -1 Or _
'''                   .HWFANTNP = -1 Or .HWFANTIM = -1 Or _
'''                   .HWFBM1ET = -1 Or .HWFBM2ET = -1 Or .HWFBM3ET = -1 Or _
'''                   .HWFOF1ET = -1 Or .HWFOF2ET = -1 Or .HWFOF3ET = -1 Or .HWFOF4ET = -1 Or _
'''                   .HEPBM1ET = -1 Or .HEPBM2ET = -1 Or .HEPBM3ET = -1 Or _
'''                   .HEPOF1ET = -1 Or .HEPOF2ET = -1 Or .HEPOF3ET = -1 Then
'''                    MakeMesIndTbl = FUNCTION_RETURN_FAILURE
'''                    Exit Function
'''                End If
                
                If .HWFMKCET = -1 Or _
                   .HWFANTNP = -1 Or .HWFANTIM = -1 Or _
                   .HWFBM1ET = -1 Or .HWFBM2ET = -1 Or .HWFBM3ET = -1 Or _
                   .HWFOF1ET = -1 Or .HWFOF2ET = -1 Or .HWFOF3ET = -1 Or _
                   .HEPBM1ET = -1 Or .HEPBM2ET = -1 Or .HEPBM3ET = -1 Or _
                   .HEPOF1ET = -1 Or .HEPOF2ET = -1 Or .HEPOF3ET = -1 Then
                    MakeMesIndTbl = FUNCTION_RETURN_FAILURE
                    Exit Function
                End If
                '◆--- 2010/01/20 SIRD対応 SPK habuki REP  END (OSF4->SIRD)

                'DKｱﾆｰﾙ条件変更(IG区分"4"→"R",ガス種追加)
                sDKAN = IIf(.HWFIGKBN = "3" Or .HWFIGKBN = "4", "R ", "V ") & Format(.HWFANTNP, "@@@@") & " " & .HWFANGZY

                'エピ先行評価追加対応
                ' エピ先行評価項目用のDKアニール条件
                ' (1桁目：品EPIG区分,3〜6桁目：品EPAN温度,8桁目：品EP高温ANガス条件,10桁目：品E1厚中心の整数部1の位)
                sDKAN_EP = IIf(.HEPIGKBN = "3" Or .HEPIGKBN = "4", "R", "V") & " " & _
                            IIf(.HEPANTNP >= 0, Format(.HEPANTNP, "@@@@"), Space(4)) & " " & _
                            .HEPANGZY & " " & _
                            IIf(.HEPACEN >= 0, Mid(Format(.HEPACEN, "000.00"), 3, 1), Space(1))

                If udtPWafSmp(i).WFINDRSCW <> "0" And udtPWafSmp(i).WFINDRSCW <> "2" Then
                    k = k + 1
                    udtPMesInd(k).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPMesInd(k).OSITEM = "RES"
                    udtPMesInd(k).SAMPLEKB = "A"
                    udtPMesInd(k).Spec = "RES"
                    udtPMesInd(k).NETSU = ""
                    udtPMesInd(k).ET = ""
                    udtPMesInd(k).MES = .HWFRSPOH & .HWFRSPOT & .HWFRSPOI
                    udtPMesInd(k).DKAN = sDKAN
                    udtPMesInd(k).MAISU = "1"
                End If

                If udtPWafSmp(i).WFINDOICW <> "0" And udtPWafSmp(i).WFINDOICW <> "2" Then
                    k = k + 1
                    udtPMesInd(k).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPMesInd(k).OSITEM = "OI"
                    udtPMesInd(k).SAMPLEKB = "A"
                    udtPMesInd(k).Spec = "OI"
                    udtPMesInd(k).NETSU = ""
                    udtPMesInd(k).ET = ""
                    udtPMesInd(k).MES = .HWFONSPH & .HWFONSPT & .HWFONSPI
                    udtPMesInd(k).DKAN = sDKAN
                    udtPMesInd(k).MAISU = "1"
                End If

                If udtPWafSmp(i).WFINDB1CW <> "0" And udtPWafSmp(i).WFINDB1CW <> "2" Then
                    k = k + 1
                    udtPMesInd(k).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPMesInd(k).OSITEM = "BMD"
                    udtPMesInd(k).SAMPLEKB = "A"
                    udtPMesInd(k).Spec = "BMD1"
                    udtPMesInd(k).NETSU = .HWFBM1NS
                    udtPMesInd(k).ET = .HWFBM1SZ & Format(.HWFBM1ET, "00")
                    udtPMesInd(k).MES = .HWFBM1SH & .HWFBM1ST & .HWFBM1SR
                    udtPMesInd(k).DKAN = sDKAN
                    udtPMesInd(k).MAISU = "1"
                End If

                If udtPWafSmp(i).WFINDB2CW <> "0" And udtPWafSmp(i).WFINDB2CW <> "2" Then
                    k = k + 1
                    udtPMesInd(k).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPMesInd(k).OSITEM = "BMD"
                    udtPMesInd(k).SAMPLEKB = "A"
                    udtPMesInd(k).Spec = "BMD2"
                    udtPMesInd(k).NETSU = .HWFBM2NS
                    udtPMesInd(k).ET = .HWFBM2SZ & Format(.HWFBM2ET, "00")
                    udtPMesInd(k).MES = .HWFBM2SH & .HWFBM2ST & .HWFBM2SR
                    udtPMesInd(k).DKAN = sDKAN
                    udtPMesInd(k).MAISU = "1"
                End If

                If udtPWafSmp(i).WFINDB3CW <> "0" And udtPWafSmp(i).WFINDB3CW <> "2" Then
                    k = k + 1
                    udtPMesInd(k).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPMesInd(k).OSITEM = "BMD"
                    udtPMesInd(k).SAMPLEKB = "A"
                    udtPMesInd(k).Spec = "BMD3"
                    udtPMesInd(k).NETSU = .HWFBM3NS
                    udtPMesInd(k).ET = .HWFBM3SZ & Format(.HWFBM3ET, "00")
                    udtPMesInd(k).MES = .HWFBM3SH & .HWFBM3ST & .HWFBM3SR
                    udtPMesInd(k).DKAN = sDKAN
                    udtPMesInd(k).MAISU = "1"
                End If

                If udtPWafSmp(i).WFINDL1CW <> "0" And udtPWafSmp(i).WFINDL1CW <> "2" Then
                    k = k + 1
                    udtPMesInd(k).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPMesInd(k).OSITEM = "OSF"
                    udtPMesInd(k).SAMPLEKB = "A"
                    udtPMesInd(k).Spec = "OSF1"
                    udtPMesInd(k).NETSU = .HWFOF1NS
                    udtPMesInd(k).ET = .HWFOF1SZ & Format(.HWFOF1ET, "00")
                    udtPMesInd(k).MES = .HWFOF1SH & .HWFOF1ST & .HWFOF1SR
                    udtPMesInd(k).DKAN = sDKAN
                    udtPMesInd(k).MAISU = "1"
                End If

                If udtPWafSmp(i).WFINDL2CW <> "0" And udtPWafSmp(i).WFINDL2CW <> "2" Then
                    k = k + 1
                    udtPMesInd(k).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPMesInd(k).OSITEM = "OSF"
                    udtPMesInd(k).SAMPLEKB = "A"
                    udtPMesInd(k).Spec = "OSF2"
                    udtPMesInd(k).NETSU = .HWFOF2NS
                    udtPMesInd(k).ET = .HWFOF2SZ & Format(.HWFOF2ET, "00")
                    udtPMesInd(k).MES = .HWFOF2SH & .HWFOF2ST & .HWFOF2SR
                    udtPMesInd(k).DKAN = sDKAN
                    udtPMesInd(k).MAISU = "1"
                End If

                If udtPWafSmp(i).WFINDL3CW <> "0" And udtPWafSmp(i).WFINDL3CW <> "2" Then
                    k = k + 1
                    udtPMesInd(k).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPMesInd(k).OSITEM = "OSF"
                    udtPMesInd(k).SAMPLEKB = "A"
                    udtPMesInd(k).Spec = "OSF3"
                    udtPMesInd(k).NETSU = .HWFOF3NS
                    udtPMesInd(k).ET = .HWFOF3SZ & Format(.HWFOF3ET, "00")
                    udtPMesInd(k).MES = .HWFOF3SH & .HWFOF3ST & .HWFOF3SR
                    udtPMesInd(k).DKAN = sDKAN
                    udtPMesInd(k).MAISU = "1"
                End If

'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)
'''                If udtPWafSmp(i).WFINDL4CW <> "0" And udtPWafSmp(i).WFINDL4CW <> "2" Then
'''                    k = k + 1
'''                    udtPMesInd(k).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
'''                    udtPMesInd(k).OSITEM = "OSF"
'''                    udtPMesInd(k).SAMPLEKB = "A"
'''                    udtPMesInd(k).SPEC = "OSF4"
'''                    udtPMesInd(k).NETSU = .HWFOF4NS
'''                    udtPMesInd(k).ET = .HWFOF4SZ & Format(.HWFOF4ET, "00")
'''                    udtPMesInd(k).MES = .HWFOF4SH & .HWFOF4ST & .HWFOF4SR
'''                    udtPMesInd(k).DKAN = sDKAN
'''                    udtPMesInd(k).MAISU = "1"
'''                End If

                If udtPWafSmp(i).WFINDL4CW <> "0" And udtPWafSmp(i).WFINDL4CW <> "2" Then
                    k = k + 1
                    udtPMesInd(k).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
'                    udtPMesInd(k).OSITEM = "SIRD"  2010/05/19 REP Y.Hitomi
                    udtPMesInd(k).OSITEM = "TENI"
                    udtPMesInd(k).SAMPLEKB = "A"
'                    udtPMesInd(k).Spec = "SIRD"    2010/05/19 REP Y.Hitomi
                    udtPMesInd(k).Spec = "TENI"
                    udtPMesInd(k).NETSU = ""
                    udtPMesInd(k).ET = .HWFSIRDSZ '''& Format(.HWFOF4ET, "00")
                    udtPMesInd(k).MES = ""
                    udtPMesInd(k).DKAN = sDKAN
                    udtPMesInd(k).MAISU = "1"
                End If
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START(OSF4->SIRD)

                If udtPWafSmp(i).WFINDDSCW <> "0" And udtPWafSmp(i).WFINDDSCW <> "2" Then
                    k = k + 1
                    udtPMesInd(k).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPMesInd(k).OSITEM = "DSOD"
                    udtPMesInd(k).SAMPLEKB = "A"
                    udtPMesInd(k).Spec = "DSOD"
                    udtPMesInd(k).NETSU = "G0"
                    udtPMesInd(k).ET = ""
                    udtPMesInd(k).MES = ""
                    udtPMesInd(k).DKAN = sDKAN
                    udtPMesInd(k).MAISU = "1"
                End If

                If udtPWafSmp(i).WFINDDZCW <> "0" And udtPWafSmp(i).WFINDDZCW <> "2" Then
                    k = k + 1
                    udtPMesInd(k).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPMesInd(k).OSITEM = "DZ"
                    udtPMesInd(k).SAMPLEKB = "A"
                    udtPMesInd(k).Spec = "DZ"
                    udtPMesInd(k).NETSU = .HWFMKNSW
                    udtPMesInd(k).ET = .HWFMKSZY & Format(.HWFMKCET, "00")
                    udtPMesInd(k).MES = .HWFMKSPH & .HWFMKSPT & .HWFMKSPR
                    udtPMesInd(k).DKAN = sDKAN
                    udtPMesInd(k).MAISU = "1"
                End If

                If udtPWafSmp(i).WFINDSPCW <> "0" And udtPWafSmp(i).WFINDSPCW <> "2" Then
                    k = k + 1
                    udtPMesInd(k).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPMesInd(k).OSITEM = "SPV"
                    udtPMesInd(k).SAMPLEKB = "A"
                    udtPMesInd(k).Spec = "SPV"
                    udtPMesInd(k).NETSU = ""
                    udtPMesInd(k).ET = ""

                    If .HWFSPVHS = "H" Or .HWFSPVHS = "S" Then
                        udtPMesInd(k).MES = .HWFSPVSH & .HWFSPVST & .HWFSPVSI
                    ElseIf .HWFDLHWS = "H" Or .HWFDLHWS = "S" Then
                        udtPMesInd(k).MES = .HWFDLSPH & .HWFDLSPT & .HWFDLSPI
                    Else    'Nr濃度追加
                        udtPMesInd(k).MES = .HWFNRSH & .HWFNRST & .HWFNRSI
                    End If

                    udtPMesInd(k).DKAN = sDKAN
                    udtPMesInd(k).MAISU = "1"

                    udtPMesInd(k).FEPUA = .HWFSPVPUG           'SPV_Fe_PUA値
                    udtPMesInd(k).FEPUAPCT = .HWFSPVPUR        'SPV_Fe_PUA％値
                    udtPMesInd(k).FESTD = .HWFSPVSTD           'SPV_Fe_STD
                    udtPMesInd(k).DIFFPUA = .HWFDLPUG          'SPV_拡散長_PUA値
                    udtPMesInd(k).DIFFPUAPCT = .HWFDLPUR       'SPV_拡散長_PUA％値
                    udtPMesInd(k).NRPUA = .HWFNRPUG            'SPV_NR_PUA値
                    udtPMesInd(k).NRPUAPCT = .HWFNRPUR         'SPV_NR_PUA%値
                    udtPMesInd(k).NRSTD = .HWFNRSTD            'SPV_NR_STD
                End If

                If udtPWafSmp(i).WFINDDO1CW <> "0" And udtPWafSmp(i).WFINDDO1CW <> "2" Then
                    k = k + 1
                    udtPMesInd(k).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPMesInd(k).OSITEM = "DOI"
                    udtPMesInd(k).SAMPLEKB = "A"
                    udtPMesInd(k).Spec = "DOI1"
                    udtPMesInd(k).NETSU = .HWFOS1NS
                    udtPMesInd(k).ET = ""
                    udtPMesInd(k).MES = .HWFOS1SH & .HWFOS1ST & .HWFOS1SI
                    udtPMesInd(k).DKAN = sDKAN
                    udtPMesInd(k).MAISU = "1"
                End If

                If udtPWafSmp(i).WFINDDO2CW <> "0" And udtPWafSmp(i).WFINDDO2CW <> "2" Then
                    k = k + 1
                    udtPMesInd(k).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPMesInd(k).OSITEM = "DOI"
                    udtPMesInd(k).SAMPLEKB = "A"
                    udtPMesInd(k).Spec = "DOI2"
                    udtPMesInd(k).NETSU = .HWFOS2NS
                    udtPMesInd(k).ET = ""
                    udtPMesInd(k).MES = .HWFOS2SH & .HWFOS2ST & .HWFOS2SI
                    udtPMesInd(k).DKAN = sDKAN
                    udtPMesInd(k).MAISU = "1"
                End If

                If udtPWafSmp(i).WFINDDO3CW <> "0" And udtPWafSmp(i).WFINDDO3CW <> "2" Then
                    k = k + 1
                    udtPMesInd(k).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPMesInd(k).OSITEM = "DOI"
                    udtPMesInd(k).SAMPLEKB = "A"
                    udtPMesInd(k).Spec = "DOI3"
                    udtPMesInd(k).NETSU = .HWFOS3NS
                    udtPMesInd(k).ET = ""
                    udtPMesInd(k).MES = .HWFOS3SH & .HWFOS3ST & .HWFOS3SI
                    udtPMesInd(k).DKAN = sDKAN
                    udtPMesInd(k).MAISU = "1"
                End If

                If udtPWafSmp(i).WFINDOT1CW <> "0" And udtPWafSmp(i).WFINDOT1CW <> "2" Then
                    k = k + 1
                    udtPMesInd(k).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPMesInd(k).OSITEM = "OTH1"
                    udtPMesInd(k).SAMPLEKB = "A"
                    udtPMesInd(k).Spec = "OTHER1"
                    udtPMesInd(k).NETSU = vbNullString
                    udtPMesInd(k).ET = vbNullString
                    udtPMesInd(k).MES = vbNullString
                    udtPMesInd(k).DKAN = sDKAN
                    udtPMesInd(k).MAISU = .HWOTHER1MAI
                End If

                If udtPWafSmp(i).WFINDOT2CW <> "0" And udtPWafSmp(i).WFINDOT2CW <> "2" Then
                    'エピ先行評価追加対応
                    l = l + 1
                    udtPEpMesInd(l).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPEpMesInd(l).OSITEM = "OTH2"
                    udtPEpMesInd(l).SAMPLEKB = "A"
                    udtPEpMesInd(l).Spec = "OTHER2"
                    udtPEpMesInd(l).NETSU = vbNullString
                    udtPEpMesInd(l).ET = vbNullString
                    udtPEpMesInd(l).MES = vbNullString
                    udtPEpMesInd(l).DKAN = sDKAN_EP
                    udtPEpMesInd(l).MAISU = .HWOTHER2MAI
                End If

                '' 残存酸素追加
                If udtPWafSmp(i).WFINDAOICW = "1" Then
                    k = k + 1
                    udtPMesInd(k).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPMesInd(k).OSITEM = "AOI"
                    udtPMesInd(k).SAMPLEKB = "A"
                    udtPMesInd(k).Spec = "AOI"
                    udtPMesInd(k).NETSU = .HWFZONSW
                    udtPMesInd(k).ET = ""
                    udtPMesInd(k).MES = .HWFZOSPH & .HWFZOSPT & .HWFZOSPI
                    udtPMesInd(k).DKAN = sDKAN
                    udtPMesInd(k).MAISU = "1"
                End If

                '' GD追加
                If udtPWafSmp(i).WFINDGDCW = "1" And udtPWafSmp(i).WFHSGDCW = "0" Then
                    k = k + 1
                    udtPMesInd(k).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW

                    '' 抜試指示4.5ﾗｲﾝ対応
                    If Trim(.HWFGDLINE) = "3" Then
                        udtPMesInd(k).OSITEM = "GD"
                    ElseIf Trim(.HWFGDLINE) = "4.5" Then
                        udtPMesInd(k).OSITEM = "GD45"
                    ElseIf Trim(.HWFGDLINE) = "5" Then
                        udtPMesInd(k).OSITEM = "GD50"
                    Else
                        udtPMesInd(k).OSITEM = "GD"
                    End If

                    udtPMesInd(k).SAMPLEKB = "A"

                    '規格値(SPEC) 1桁目:DVD2
                    If .HWFDVDHS = "H" Or .HWFDVDHS = "S" Then sGdSpec = "V" Else sGdSpec = Space(1)
                    sGdSpec = sGdSpec & Space(1)

                    '規格値(SPEC) 3桁目:L/DL
                    If .HWFLDLHS = "H" Or .HWFLDLHS = "S" Then sGdSpec = sGdSpec & "L" Else sGdSpec = sGdSpec & Space(1)
                    sGdSpec = sGdSpec & Space(1)

                    '規格値(SPEC) 5桁目:Den
                    If .HWFDENHS = "H" Or .HWFDENHS = "S" Then sGdSpec = sGdSpec & "D" Else sGdSpec = sGdSpec & Space(1)

                    udtPMesInd(k).Spec = sGdSpec
                    udtPMesInd(k).NETSU = ""
                    udtPMesInd(k).ET = ""
                    udtPMesInd(k).MES = .HWFGDSPH & .HWFGDSPT & .HWFGDZAR
                    udtPMesInd(k).DKAN = sDKAN
                    udtPMesInd(k).MAISU = "1"
                End If

                ' エピ先行評価追加対応
                '' OSF1E
                If udtPWafSmp(i).EPINDL1CW = "1" Then
                    l = l + 1
                    udtPEpMesInd(l).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPEpMesInd(l).OSITEM = "OSF"
                    udtPEpMesInd(l).SAMPLEKB = "A"
                    udtPEpMesInd(l).Spec = "OSF1"
                    udtPEpMesInd(l).NETSU = .HEPOF1NS
                    udtPEpMesInd(l).ET = .HEPOF1SZ & IIf(.HEPOF1ET >= 0, Format(.HEPOF1ET, "00"), Space(2))
                    udtPEpMesInd(l).MES = .HEPOF1SH & .HEPOF1ST & .HEPOF1SR
                    udtPEpMesInd(l).DKAN = sDKAN_EP
                    udtPEpMesInd(l).MAISU = "1"
                End If

                '' OSF2E
                If udtPWafSmp(i).EPINDL2CW = "1" Then
                    l = l + 1
                    udtPEpMesInd(l).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPEpMesInd(l).OSITEM = "OSF"
                    udtPEpMesInd(l).SAMPLEKB = "A"
                    udtPEpMesInd(l).Spec = "OSF2"
                    udtPEpMesInd(l).NETSU = .HEPOF2NS
                    udtPEpMesInd(l).ET = .HEPOF2SZ & IIf(.HEPOF2ET >= 0, Format(.HEPOF2ET, "00"), Space(2))
                    udtPEpMesInd(l).MES = .HEPOF2SH & .HEPOF2ST & .HEPOF2SR
                    udtPEpMesInd(l).DKAN = sDKAN_EP
                    udtPEpMesInd(l).MAISU = "1"
                End If

                '' OSF3E
                If udtPWafSmp(i).EPINDL3CW = "1" Then
                    l = l + 1
                    udtPEpMesInd(l).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPEpMesInd(l).OSITEM = "OSF"
                    udtPEpMesInd(l).SAMPLEKB = "A"
                    udtPEpMesInd(l).Spec = "OSF3"
                    udtPEpMesInd(l).NETSU = .HEPOF3NS
                    udtPEpMesInd(l).ET = .HEPOF3SZ & IIf(.HEPOF3ET >= 0, Format(.HEPOF3ET, "00"), Space(2))
                    udtPEpMesInd(l).MES = .HEPOF3SH & .HEPOF3ST & .HEPOF3SR
                    udtPEpMesInd(l).DKAN = sDKAN_EP
                    udtPEpMesInd(l).MAISU = "1"
                End If

                '' BMD1E
                If udtPWafSmp(i).EPINDB1CW = "1" Then
                    l = l + 1
                    udtPEpMesInd(l).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPEpMesInd(l).OSITEM = "BMD"
                    udtPEpMesInd(l).SAMPLEKB = "A"
                    udtPEpMesInd(l).Spec = "BMD1"
                    udtPEpMesInd(l).NETSU = .HEPBM1NS
                    udtPEpMesInd(l).ET = .HEPBM1SZ & IIf(.HEPBM1ET >= 0, Format(.HEPBM1ET, "00"), Space(2))
                    udtPEpMesInd(l).MES = .HEPBM1SH & .HEPBM1ST & .HEPBM1SR
                    udtPEpMesInd(l).DKAN = sDKAN_EP
                    udtPEpMesInd(l).MAISU = "1"
                End If

                '' BMD2E
                If udtPWafSmp(i).EPINDB2CW = "1" Then
                    l = l + 1
                    udtPEpMesInd(l).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPEpMesInd(l).OSITEM = "BMD"
                    udtPEpMesInd(l).SAMPLEKB = "A"
                    udtPEpMesInd(l).Spec = "BMD2"
                    udtPEpMesInd(l).NETSU = .HEPBM2NS
                    udtPEpMesInd(l).ET = .HEPBM2SZ & IIf(.HEPBM2ET >= 0, Format(.HEPBM2ET, "00"), Space(2))
                    udtPEpMesInd(l).MES = .HEPBM2SH & .HEPBM2ST & .HEPBM2SR
                    udtPEpMesInd(l).DKAN = sDKAN_EP
                    udtPEpMesInd(l).MAISU = "1"
                End If

                '' BMD3E
                If udtPWafSmp(i).EPINDB3CW = "1" Then
                    l = l + 1
                    udtPEpMesInd(l).SAMPLEID = udtPWafSmp(i).REPSMPLIDCW
                    udtPEpMesInd(l).OSITEM = "BMD"
                    udtPEpMesInd(l).SAMPLEKB = "A"
                    udtPEpMesInd(l).Spec = "BMD3"
                    udtPEpMesInd(l).NETSU = .HEPBM3NS
                    udtPEpMesInd(l).ET = .HEPBM3SZ & IIf(.HEPBM3ET >= 0, Format(.HEPBM3ET, "00"), Space(2))
                    udtPEpMesInd(l).MES = .HEPBM3SH & .HEPBM3ST & .HEPBM3SR
                    udtPEpMesInd(l).DKAN = sDKAN_EP
                    udtPEpMesInd(l).MAISU = "1"
                End If
                'エピ先行評価追加対応
            End With
        End If
    Next i

    ReDim Preserve udtPMesInd(k)

    ' エピ先行評価追加対応
    ReDim Preserve udtPEpMesInd(l)

    MakeMesIndTbl = FUNCTION_RETURN_SUCCESS
End Function

Public Sub SeparateUD()
    'Step3.2にて、機能廃止
End Sub

'*******************************************************************************
'*    関数名        : WarpKakuDisp
'*
'*    処理概要      : 1.Warp/合成角度の表示
'*
'*    パラメータ    : 変数名        ,IO ,型              ,説明
'*　　　　　　　　　　frmWarpForm   ,I  ,Form         　 ,Warp表示ﾌｫｰﾑ
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Public Sub WarpKakuDisp(frmWarpForm As Form)
    Dim i, j, k, m, n   As Integer
    Dim blChkHin        As Boolean      '同じﾌﾞﾛｯｸID/品番ﾃﾞｰﾀの存在有無

    With frmWarpForm.sprWarp
        .MaxRows = 0
        '■最新の合成角度情報表示
        m = UBound(tKakuMeasG)
        k = 0
        For i = 1 To m
            '同じﾌﾞﾛｯｸID/品番ﾃﾞｰﾀの存在ﾁｪｯｸ
            blChkHin = False
            If i <> 1 Then
                For j = i - 1 To 1 Step -1
                    If tKakuMeasG(j).BLOCKID = tKakuMeasG(i).BLOCKID And _
                       tKakuMeasG(j).HIN.hinban = tKakuMeasG(i).HIN.hinban And _
                       tKakuMeasG(j).HIN.mnorevno = tKakuMeasG(i).HIN.mnorevno And _
                       tKakuMeasG(j).HIN.factory = tKakuMeasG(i).HIN.factory And _
                       tKakuMeasG(j).HIN.opecond = tKakuMeasG(i).HIN.opecond Then

                        blChkHin = True
                        Exit For
                    End If
                Next j
            End If

            '既に同じﾌﾞﾛｯｸID/品番のﾃﾞｰﾀが存在する場合は表示しない
            If Not blChkHin Then
                k = k + 1
                .MaxRows = k
                
                ' 2008/02/15 SPK Tsutsumi Add Start
                If frmWarpForm.Name = "f_cmbc039_3" Then
                    .SetText 1, k, Right(tKakuMeasG(i).BLOCKID, 3)                          'ブロックID
                    .SetText 2, k, "合成角"                                                 '項目
                    .SetText 3, k, "-"                                                      '位置
                    .SetText 4, k, CStr(DBData2DispData_nl(tKakuMeasG(i).Min)) & " - " & _
                                   CStr(DBData2DispData_nl(tKakuMeasG(i).max))              '仕様値
                    .SetText 5, k, CStr(DBData2DispData_nl(tKakuMeasG(i).MEASDATA))         '測定値
                    .SetText 6, k, IIf(tKakuMeasG(i).Judg, "OK", "NG")                      '判定
                    
                    '背景色設定
                    If Not tKakuMeasG(i).Judg Then
                        SpCtrlBlockEnabled frmWarpForm.sprWarp, 1, k, 6, k, CTRL_DISABLE_RED
                    End If
                Else
                ' 2008/02/15 SPK Tsutsumi Add End
                    .SetText 1, k, "合成角"                                                 '項目
                    .SetText 2, k, "-"                                                      '位置
                    .SetText 3, k, CStr(DBData2DispData_nl(tKakuMeasG(i).Min)) & " - " & _
                                   CStr(DBData2DispData_nl(tKakuMeasG(i).max))              '仕様値
                    .SetText 4, k, CStr(DBData2DispData_nl(tKakuMeasG(i).MEASDATA))         '測定値
                    .SetText 5, k, IIf(tKakuMeasG(i).Judg, "OK", "NG")                      '判定
    
                    '背景色設定
                    If Not tKakuMeasG(i).Judg Then
                        SpCtrlBlockEnabled frmWarpForm.sprWarp, 1, k, 5, k, CTRL_DISABLE_RED
                    End If
                ' 2008/02/15 SPK Tsutsumi Add Start
                End If
                ' 2008/02/15 SPK Tsutsumi Add End
            End If
        Next i

'Add Start 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応
        '■最新の横(X)角度情報表示
        n = UBound(tKakuXMeasG)
        m = .MaxRows
        k = 0
        For i = 1 To n
            '同じﾌﾞﾛｯｸID/品番ﾃﾞｰﾀの存在ﾁｪｯｸ
            blChkHin = False
            If i <> 1 Then
                For j = i - 1 To 1 Step -1
                    If tKakuXMeasG(j).BLOCKID = tKakuXMeasG(i).BLOCKID And _
                       tKakuXMeasG(j).HIN.hinban = tKakuXMeasG(i).HIN.hinban And _
                       tKakuXMeasG(j).HIN.mnorevno = tKakuXMeasG(i).HIN.mnorevno And _
                       tKakuXMeasG(j).HIN.factory = tKakuXMeasG(i).HIN.factory And _
                       tKakuXMeasG(j).HIN.opecond = tKakuXMeasG(i).HIN.opecond Then

                        blChkHin = True
                        Exit For
                    End If
                Next j
            End If

            '既に同じﾌﾞﾛｯｸID/品番のﾃﾞｰﾀが存在する場合は表示しない
            If Not blChkHin Then
                k = k + 1
                .MaxRows = m + k
                
                If frmWarpForm.Name = "f_cmbc039_3" Then
                    .SetText 1, m + k, Right(tKakuXMeasG(i).BLOCKID, 3)                          'ブロックID
                    .SetText 2, m + k, "横角Ｘ"                                                 '項目
                    .SetText 3, m + k, "-"                                                      '位置
                    .SetText 4, m + k, CStr(DBData2DispData_nl(tKakuXMeasG(i).Min)) & " - " & _
                                       CStr(DBData2DispData_nl(tKakuXMeasG(i).max))              '仕様値
                    .SetText 5, m + k, CStr(DBData2DispData_nl(tKakuXMeasG(i).MEASDATA))         '測定値
                    .SetText 6, m + k, IIf(tKakuXMeasG(i).Judg, "OK", "NG")                      '判定
                    
                    '背景色設定
                    If Not tKakuXMeasG(i).Judg Then
                        SpCtrlBlockEnabled frmWarpForm.sprWarp, 1, m + k, 6, m + k, CTRL_DISABLE_RED
                    End If
                Else
                    .SetText 1, m + k, "横角Ｘ"                                                 '項目
                    .SetText 2, m + k, "-"                                                      '位置
                    .SetText 3, m + k, CStr(DBData2DispData_nl(tKakuXMeasG(i).Min)) & " - " & _
                                       CStr(DBData2DispData_nl(tKakuXMeasG(i).max))              '仕様値
                    .SetText 4, m + k, CStr(DBData2DispData_nl(tKakuXMeasG(i).MEASDATA))         '測定値
                    .SetText 5, m + k, IIf(tKakuXMeasG(i).Judg, "OK", "NG")                      '判定
    
                    '背景色設定
                    If Not tKakuXMeasG(i).Judg Then
                        SpCtrlBlockEnabled frmWarpForm.sprWarp, 1, m + k, 5, m + k, CTRL_DISABLE_RED
                    End If
                End If
            End If
        Next i

        '■最新の縦(Y)角度情報表示
        n = UBound(tKakuYMeasG)
        m = .MaxRows
        k = 0
        For i = 1 To n
            '同じﾌﾞﾛｯｸID/品番ﾃﾞｰﾀの存在ﾁｪｯｸ
            blChkHin = False
            If i <> 1 Then
                For j = i - 1 To 1 Step -1
                    If tKakuYMeasG(j).BLOCKID = tKakuYMeasG(i).BLOCKID And _
                       tKakuYMeasG(j).HIN.hinban = tKakuYMeasG(i).HIN.hinban And _
                       tKakuYMeasG(j).HIN.mnorevno = tKakuYMeasG(i).HIN.mnorevno And _
                       tKakuYMeasG(j).HIN.factory = tKakuYMeasG(i).HIN.factory And _
                       tKakuYMeasG(j).HIN.opecond = tKakuYMeasG(i).HIN.opecond Then

                        blChkHin = True
                        Exit For
                    End If
                Next j
            End If

            '既に同じﾌﾞﾛｯｸID/品番のﾃﾞｰﾀが存在する場合は表示しない
            If Not blChkHin Then
                k = k + 1
                .MaxRows = m + k
                
                If frmWarpForm.Name = "f_cmbc039_3" Then
                    .SetText 1, m + k, Right(tKakuYMeasG(i).BLOCKID, 3)                          'ブロックID
                    .SetText 2, m + k, "縦角Ｙ"                                                  '項目
                    .SetText 3, m + k, "-"                                                       '位置
                    .SetText 4, m + k, CStr(DBData2DispData_nl(tKakuYMeasG(i).Min)) & " - " & _
                                       CStr(DBData2DispData_nl(tKakuYMeasG(i).max))              '仕様値
                    .SetText 5, m + k, CStr(DBData2DispData_nl(tKakuYMeasG(i).MEASDATA))         '測定値
                    .SetText 6, m + k, IIf(tKakuYMeasG(i).Judg, "OK", "NG")                      '判定
                    
                    '背景色設定
                    If Not tKakuYMeasG(i).Judg Then
                        SpCtrlBlockEnabled frmWarpForm.sprWarp, 1, m + k, 6, m + k, CTRL_DISABLE_RED
                    End If
                Else
                    .SetText 1, m + k, "縦角Ｙ"                                                  '項目
                    .SetText 2, m + k, "-"                                                       '位置
                    .SetText 3, m + k, CStr(DBData2DispData_nl(tKakuYMeasG(i).Min)) & " - " & _
                                       CStr(DBData2DispData_nl(tKakuYMeasG(i).max))              '仕様値
                    .SetText 4, m + k, CStr(DBData2DispData_nl(tKakuYMeasG(i).MEASDATA))         '測定値
                    .SetText 5, m + k, IIf(tKakuYMeasG(i).Judg, "OK", "NG")                      '判定
    
                    '背景色設定
                    If Not tKakuYMeasG(i).Judg Then
                        SpCtrlBlockEnabled frmWarpForm.sprWarp, 1, m + k, 5, m + k, CTRL_DISABLE_RED
                    End If
                End If
            End If
        Next i
'Add End 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応
        
        '■最新のWarp情報表示
        n = UBound(tWarpMeasG)
        m = .MaxRows
        k = 0
        For j = 1 To n
            '実ﾃﾞｰﾀが無い場合は表示しない
            If tWarpMeasG(j).EXISTFLG > 0 Then
                k = k + 1
                .MaxRows = m + k
                
                ' 2008/02/15 SPK Tsutsumi Add Start
                If frmWarpForm.Name = "f_cmbc039_3" Then
                    .SetText 1, m + k, Right(tWarpMeasG(j).BLOCKID, 3)                      'ブロックID
                    .SetText 2, m + k, "Warp"                                               '項目
                    .SetText 3, m + k, CStr(tWarpMeasG(j).WAFID)                            '位置
                    .SetText 4, m + k, CStr(DBData2DispData_nl(tWarpMeasG(j).max))          '仕様値
                    .SetText 5, m + k, CStr(DBData2DispData_nl(tWarpMeasG(j).MEASDATA))     '測定値
                    .SetText 6, m + k, IIf(tWarpMeasG(j).Judg, "OK", "NG")                  '判定
                    
                    '背景色設定
                    If Not tWarpMeasG(j).Judg Then
                        SpCtrlBlockEnabled frmWarpForm.sprWarp, 1, m + k, 6, m + k, CTRL_DISABLE_RED
                    End If
                Else
                ' 2008/02/15 SPK Tsutsumi Add End
                    .SetText 1, m + k, "Warp"                                               '項目
                    .SetText 2, m + k, CStr(tWarpMeasG(j).WAFID)                            '位置
                    .SetText 3, m + k, CStr(DBData2DispData_nl(tWarpMeasG(j).max))          '仕様値
                    .SetText 4, m + k, CStr(DBData2DispData_nl(tWarpMeasG(j).MEASDATA))     '測定値
                    .SetText 5, m + k, IIf(tWarpMeasG(j).Judg, "OK", "NG")                  '判定
                    '背景色設定
                    If Not tWarpMeasG(j).Judg Then
                        SpCtrlBlockEnabled frmWarpForm.sprWarp, 1, m + k, 5, m + k, CTRL_DISABLE_RED
                    End If
                ' 2008/02/15 SPK Tsutsumi Add Start
                End If
                ' 2008/02/15 SPK Tsutsumi Add End
            End If
        Next j
    End With
End Sub

'*******************************************************************************************
'*    関数名        : GetMukeCode
'*
'*    処理概要      : 1.向先コード・向先名を取得し、向先名を表示
'*
'*    パラメータ    : 変数名        ,IO ,型              ,説明
'*　　　　　　　　　　なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Public Function GetMukeCode() As FUNCTION_RETURN
    Dim sSql        As String
    Dim rs          As OraDynaset
    Dim intRecCnt   As Long      'レコード数
    Dim i           As Long

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmbc016_0.frm -- Function Getstaffauthority"

    GetMukeCode = FUNCTION_RETURN_FAILURE

    sSql = "Select CODEA9,NAMEJA9 "
    sSql = sSql & "from KODA9 "
    sSql = sSql & "where SYSCA9 = 'X' "
    sSql = sSql & "and SHUCA9 = '20' "
    sSql = sSql & "and (CODEA9 = '14' "
    sSql = sSql & "or CODEA9 = '15' "
    sSql = sSql & "or CODEA9 = '16' "
    sSql = sSql & "or CODEA9 = 'ALL') "
    sSql = sSql & "order by CODEA9 "    '向先不具合対応 2010/01/05 SETsw kubota

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    intRecCnt = rs.RecordCount
    ReDim s_MukesakiBase(intRecCnt)

    If intRecCnt = 0 Then
        Exit Function
    End If

    For i = 1 To intRecCnt
        With s_MukesakiBase(i)
            If IsNull(rs.Fields("CODEA9")) = False Then
                .sMukeCode = rs.Fields("CODEA9")    ' 向先コード
            End If

            If IsNull(rs.Fields("NAMEJA9")) = False Then
                .sMukeName = rs.Fields("NAMEJA9")  ' 向先名
                f_cmbc039_1.cmbMukesaki.AddItem .sMukeName
            End If
        End With
        rs.MoveNext
    Next
    rs.Close

    GetMukeCode = FUNCTION_RETURN_SUCCESS
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    Resume proc_exit
End Function

'***********************************************************************************
'*    関数名        : ChkMukesaki_E001
'*
'*    処理概要      : 1.入力された品番に向先が指定されているかTBCME001から確認
'*
'*    パラメータ    : 変数名        ,IO ,型              ,説明
'*　　　　　　　　　　sHinban　　　 ,I  ,String          ,品番
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'***********************************************************************************
Public Function ChkMukesaki_E001(sHinban As String) As FUNCTION_RETURN
    Dim lngLp   As Long
    Dim sBuf    As String
    Dim rs      As OraDynaset
    Dim sSql    As String
    Dim sMuke4  As String
    Dim sMuke5  As String
    Dim sMuke6  As String

    sBuf = ""

    sSql = "Select hinban,MAX(MNOREVNO), SUM(NVL(TRIM(E1.KFCTFLAG1),'')) FLAG1, SUM(NVL(TRIM(E1.KFCTFLAG2),'')) FLAG2, SUM(NVL(TRIM(E1.KFCTFLAG3),'')) FLAG3 "
    sSql = sSql & "from TBCME001 E1 "
    sSql = sSql & "where E1.HINBAN = '" & Trim(sHinban) & "' "
    sSql = sSql & "and E1.OPECOND = '1' "
    sSql = sSql & "group by hinban"

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        f_cmbc039_2.lblMsg.Caption = "向先取得エラー TBCME001"
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    If IsNull(rs("FLAG1")) = False Then sMuke4 = CStr(rs("FLAG1"))   '向先４棟
    If IsNull(rs("FLAG2")) = False Then sMuke5 = CStr(rs("FLAG2"))   '向先５棟
    If IsNull(rs("FLAG3")) = False Then sMuke6 = CStr(rs("FLAG3"))   '向先６棟
    rs.Close

    For lngLp = 1 To UBound(s_MukesakiBase)
        Select Case lngLp
            Case 1
                If sMuke4 <> "" Then
                    sBuf = s_MukesakiBase(lngLp).sMukeName
                    If sBuf = sBaseMukesaki Then
                        ChkMukesaki_E001 = FUNCTION_RETURN_SUCCESS
                        GoTo proc_exit
                    End If

                End If
            Case 2
                If sMuke5 <> "" Then
                    sBuf = s_MukesakiBase(lngLp).sMukeName
                    If sBuf = sBaseMukesaki Then
                        ChkMukesaki_E001 = FUNCTION_RETURN_SUCCESS
                        GoTo proc_exit
                    End If
                End If
            Case 3
                If sMuke6 <> "" Then
                    sBuf = s_MukesakiBase(lngLp).sMukeName
                    If sBuf = sBaseMukesaki Then
                        ChkMukesaki_E001 = FUNCTION_RETURN_SUCCESS
                        GoTo proc_exit
                    End If

                End If
        End Select
    Next lngLp

    If sBuf = "" Then
        f_cmbc039_2.lblMsg.Caption = "向先取得エラー TBCME001"
    Else
        ChkMukesaki_E001 = FUNCTION_RETURN_FAILURE
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    Resume proc_exit
End Function

'***********************************************************************************
'*    関数名        : GetMukesaki
'*
'*    処理概要      : 1.入力された品番に対する向先を表示
'*
'*    パラメータ    : 変数名        ,IO ,型              ,説明
'*　　　　　　　　　　sHinban　　　 ,I  ,String          ,品番
'*
'*    戻り値        : String(向先)
'*
'***********************************************************************************
Public Function GetMukesaki(sHinban As String) As String
    Dim lngLp   As Long
    Dim sBuf    As String
    Dim rs      As OraDynaset
    Dim sSql    As String
    Dim sMuke4  As String
    Dim sMuke5  As String
    Dim sMuke6  As String

    GetMukesaki = ""
    sBuf = ""

    sSql = "Select hinban,MAX(MNOREVNO), SUM(NVL(TRIM(E1.KFCTFLAG1),'')) FLAG1, SUM(NVL(TRIM(E1.KFCTFLAG2),'')) FLAG2, SUM(NVL(TRIM(E1.KFCTFLAG3),'')) FLAG3 "
    sSql = sSql & "from TBCME001 E1 "
    sSql = sSql & "where E1.HINBAN = '" & Trim(sHinban) & "' "
    sSql = sSql & "and E1.OPECOND = '1' "
    sSql = sSql & "group by hinban"

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        f_cmbc039_2.lblMsg.Caption = "向先取得エラー TBCME001"
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    If IsNull(rs("FLAG1")) = False Then sMuke4 = CStr(rs("FLAG1"))   '向先４棟
    If IsNull(rs("FLAG2")) = False Then sMuke5 = CStr(rs("FLAG2"))   '向先５棟
    If IsNull(rs("FLAG3")) = False Then sMuke6 = CStr(rs("FLAG3"))   '向先６棟
    rs.Close

    For lngLp = 1 To UBound(s_MukesakiBase)
        Select Case lngLp
            Case 1
                If sMuke4 <> "" Then
                    sBuf = left(s_MukesakiBase(lngLp).sMukeName, 1)
                End If
            Case 2
                If sMuke5 <> "" Then
                    If sBuf = "" Then
                        sBuf = left(s_MukesakiBase(lngLp).sMukeName, 1)
                    Else
                        sBuf = sBuf & "," & left(s_MukesakiBase(lngLp).sMukeName, 1)
                    End If
                End If
            Case 3
                If sMuke6 <> "" Then
                    If sBuf = "" Then
                        sBuf = left(s_MukesakiBase(lngLp).sMukeName, 1)
                    Else
                        sBuf = sBuf & "," & left(s_MukesakiBase(lngLp).sMukeName, 1)
                    End If
                End If
        End Select
    Next lngLp

    If sBuf = "" Then
        f_cmbc039_2.lblMsg.Caption = "向先取得エラー TBCME001"
    Else
        GetMukesaki = sBuf
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    Resume proc_exit
End Function

'******************************************************************************************
'* Function Name    :   gFnc_SS_RecordSet
'*
'* Function         :   ｽﾌﾟﾚｯﾄｼｰﾄにﾃﾞｰﾀを表示(1行)
'*
'* Parameter(lngRow)     :   objSpdSeet    ｽﾌﾟﾚｯﾄｼｰﾄｵﾌﾞｼﾞｪｸﾄ
'*                      lngRow           セットするﾚｺｰﾄﾞ
'*                      strRecord        ｽﾌﾟﾚｯﾄｼｰﾄに表示するﾚｺｰﾄﾞﾃﾞｰﾀ
'*
'* Parameter(o)     :   なし
'*
'* Return Code      :   正常終了時は gstrconReturn_OK
'*                      ｴﾗｰ終了時は ｴﾗｰｺｰﾄﾞ + ﾒｯｾｰｼﾞ
'*************************************************************************************
Public Function gFnc_SS_RecordSet(objSpdSeet As Control, lngRow As Integer, strRecord As String, udt_ww() As DBDRV_scmzc_fcmlc001b_SXL039, i As Integer) As Boolean
    '変数の宣言
    Dim strRc       As String    'ﾘﾀｰﾝｺｰﾄﾞ
    Dim strErrCd    As String   'ｴﾗｰｺｰﾄﾞ
    Dim strSplt()   As String   '停止理由 add SETkimizuka 09/03/16
    
    'ｴﾗｰﾄﾗｯﾌﾟ
    On Error GoTo ErrorHandler
    
    '戻り値の初期化
    gFnc_SS_RecordSet = True

    With objSpdSeet
        'ﾃﾞｰﾀ格納ﾌﾞﾛｯｸを指定
        .row = lngRow
        .col = 1
        .row2 = lngRow
        .col2 = objSpdSeet.MaxCols
        
        '1行データセット
        .Clip = strRecord
        
        SpCtrlBlockEnabled f_cmbc039_1.spdWait, 2, lngRow, 11, lngRow, CTRL_DISABLE
    
        If Not WFJudgExecOkFlag(i) Then
' 2009/03/17 SETkimizuka upd Start
' 2007/10/17 SET miyatake Add Start
'                    SpCtrlBlockEnabled f_cmbc039_1.spdWait, 1, lngRow, 11, lngRow, CTRL_DISABLE_GRAY
'            SpCtrlBlockEnabled f_cmbc039_1.spdWait, 1, lngRow, 12, lngRow, CTRL_DISABLE_GRAY
' 2007/10/17 SET miyatake Add End
            SpCtrlBlockEnabled f_cmbc039_1.spdWait, 1, lngRow, f_cmbc039_1.spdWait.MaxCols, lngRow, CTRL_DISABLE_GRAY
' 2009/03/17 SETkimizuka upd End
        End If

        'ホールドロット（0=通常，1=流動停止）
' 2009/03/17 結晶ﾎｰﾙﾄﾞを流動監視に置き換え SETkimizuka upd Start
'        If udt_ww(i).WFHOLDFLGCB = "1" Then
''        If udt_ww(i).HOLDBCB = "1" Or udt_ww(i).WFHOLDFLGCB = "1" Then
'
'            'ﾎｰﾙﾄﾞ区分またはﾎｰﾙﾄﾞ区分(WF)が「1」のﾛｯﾄは選択不可とする
'' 2007/10/17 SET miyatake Add Start
''                    SpCtrlBlockEnabled f_cmbc039_1.spdWait, 1, lngRow, 11, lngRow, CTRL_DISABLE_RED
'            SpCtrlBlockEnabled f_cmbc039_1.spdWait, 1, lngRow, 12, lngRow, CTRL_DISABLE_RED
'' 2007/10/17 SET miyatake Add End
''''                End If
'        End If
' 2009/03/17 結晶ﾎｰﾙﾄﾞを流動監視に置き換え SETkimizuka upd End
        
    
        '流動停止項目追加 add SETkimizuka Start  09/03/16
        If udt_ww(i).STOP <> "0" And udt_ww(i).STOP <> "" Then
            SpCtrlBlockEnabled f_cmbc039_1.spdWait, 2, lngRow, f_cmbc039_1.spdWait.MaxCols, lngRow, CTRL_DISABLE_RED
        End If
        If f_cmbc039_1.chkY4Disp.Value = 1 Then
            .col = 13
            .text = SetSortAgrStatusName(udt_ww(i).AGRSTATUS)
            .col = 14
            strSplt = Split(udt_ww(i).CAUSE, Chr(9))
            If UBound(strSplt) > 1 Then
                .CellType = CellTypeComboBox
                .TypeComboBoxList = udt_ww(i).CAUSE
                .Lock = False
                .TypeComboBoxCurSel = 0
            Else
                .text = udt_ww(i).CAUSE
            End If
            .col = 15
            strSplt = Split(udt_ww(i).PRINTNO, Chr(9))
            If UBound(strSplt) >= 1 Then
                .Lock = False
                .CellType = CellTypeButton
                .TypeButtonText = "確認"
            End If
        End If
        '流動停止項目追加 add SETkimizuka End  09/03/16
    
    End With
            
    Exit Function
    
ErrorHandler:
    gFnc_SS_RecordSet = False
End Function

'概要    :結晶位置取得
'ﾊﾟﾗﾒｰﾀ  :変数名        ,IO  ,型                                    ,説明
'        :sCrynum       ,I   ,string                                ,ブロックＩＤ
'        :戻ﾘ値         ,O   ,結晶位置                              ,成否
'説明    :
'履歴    :2010/03/25 Kameda　新規作成
Public Function GetXtalPos(sSXLID As String) As String
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim recCnt      As Integer
    Dim iBuiS As Integer
    Dim iNewL As Integer
    Dim iMai  As Integer 'Add 2010/08/25 Y.Hitomi
    
    GetXtalPos = ""
    
    sql = "SELECT SXLIDCB, "
    sql = sql & "INPOSCB , "
    sql = sql & "RLENCB, "
    'Add Start 2010/08/25 Y.Hitomi
    sql = sql & "MAICB "
    'Add End   2010/08/25 Y.Hitomi
    sql = sql & "FROM XSDCB "
    sql = sql & "WHERE SXLIDCB = '" & sSXLID & "' "

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt > 0 Then
        iBuiS = fncNullCheck(rs("INPOSCB"))
        iNewL = fncNullCheck(rs("RLENCB"))
        'Add Start 2010/08/25 Y.Hitomi
        iMai = fncNullCheck(rs("MAICB"))
        'Add End   2010/08/25 Y.Hitomi
        
        GetXtalPos = iBuiS & "-" & (iBuiS + iNewL) & "/" & iNewL & "(" & iMai & ")"
    End If
    
    
    rs.Close
End Function



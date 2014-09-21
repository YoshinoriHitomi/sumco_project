Attribute VB_Name = "SB_Com"
Option Explicit

'WFｻﾝﾌﾟﾙ構造体
Public Type typ_Wf_Smpl
    SXLIDCW     As String * 13      'SXL-ID
    TBKBNCW     As String * 1       'T/B区分
    XTALCW      As String * 12      '結晶番号
    INPOSCW     As Integer          '結晶内位置
    HINBCW      As String * 8       '品番
    REVNUMCW    As Integer          '製品番号改訂番号
    FACTORYCW   As String * 1       '工場
    OPECW       As String * 1       '操業条件
End Type

'結晶ｻﾝﾌﾟﾙ構造体
Public Type typ_Cry_Smpl
    CRYNUMCS    As String * 12      'ﾌﾞﾛｯｸID
    SMPKBNCS    As String * 1       'ｻﾝﾌﾟﾙ区分
    TBKBNCS     As String * 1       'T/B区分
    REPSMPLIDCS As Long             '代表ｻﾝﾌﾟﾙID    Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    XTALCS      As String * 12      '結晶番号
    INPOSCS     As Integer          '結晶内位置
    HINBCS      As String * 8       '品番
    REVNUMCS    As Integer          '製品番号改訂番号
    FACTORYCS   As String * 1       '工場
    OPECS       As String * 1       '操業条件
End Type

Public CrySampleID  As typ_CpyJisseki     ' 結晶実績引継ぎﾃﾞｰﾀ　05/06/13 ooba

'結晶実績引継ぎﾃﾞｰﾀ　05/06/13 ooba
Public Type typ_CpyJisseki
    TsmplidGD   As String * 16          'TOP_ｻﾝﾌﾟﾙID(GD)
    TindGD      As String * 1           'TOP_状態FLG(GD)
    BsmplidGD   As String * 16          'BOT_ｻﾝﾌﾟﾙID(GD)
    BindGD      As String * 1           'BOT_状態FLG(GD)
End Type

'Warp/合成角度測定値表示ﾃﾞｰﾀ　05/12/19 ooba
Public Type typ_WarpKakuData
    BLOCKID     As String * 12              'ﾌﾞﾛｯｸID
    HIN         As tFullHinban              '品番
    WAFID       As Double                   'ｳｪﾊｰID
    Min         As Double                   '仕様Min値
    max         As Double                   '仕様Max値
    MEASDATA    As Double                   '測定値
    Judg        As Boolean                  '判定(True:判定OK,False:判定NG)
    EXISTFLG    As Integer                  '存在ﾌﾗｸﾞ(1:実ﾃﾞｰﾀ有,0:実ﾃﾞｰﾀ無,-1:WFﾏｯﾌﾟ紐付け無)
End Type

'WFﾏｯﾌﾟ上の品番ﾃﾞｰﾀ　05/12/19 ooba
Public Type typ_MapHinData
    BLOCKID     As String * 12              'ﾌﾞﾛｯｸID
    HIN         As tFullHinban              '品番
    BLKSEQ_S    As Integer                  'ﾌﾞﾛｯｸ内連番(Start)
    BLKSEQ_E    As Integer                  'ﾌﾞﾛｯｸ内連番(End)
    WARPFLG     As Boolean                  'Warp振替ﾁｪｯｸﾌﾗｸﾞ
    KAKUFLG     As Boolean                  '合成角度振替ﾁｪｯｸﾌﾗｸﾞ
    'Add Start 2011/04/25 SMPK Miyata
    XTALCS      As String * 12              '結晶番号
    INPOSCS_S   As Integer                  '結晶内位置(Start)
    INPOSCS_E   As Integer                  '結晶内位置(End)
    'Add End   2011/04/25 SMPK Miyata
End Type

'C－OSF3判定機能追加 2007/04/23 M.Kaga STRAT ---
Public Type typ_COSF3ID
    
    C_XTALC1    As String          ' 結晶番号
    C_JDGEIDC1  As String          ' C-OSF3判定ID
    C_SYNFLAGC5 As String          ' 承認ﾌﾗｸﾞ
    C_YMKFLAGC5 As String          ' 削除ﾌﾗｸﾞ
    C_strChkR   As String          ' 判定用ﾊﾟﾀｰﾝ区分
    C_strChkD   As String          ' 判定用ﾊﾟﾀｰﾝ区分
    C_POSC5     As String          ' ｻﾝﾌﾟﾙ位置
    C_DMAXC5    As String          ' Dのみ上限
    C_RMAXC5    As String          ' Rのみ上限
    C_DRDMAXC5  As String          ' D共存上限
    C_DRRMAXC5  As String          ' R共存上限
         
'C－OSF3判定機能追加 2007/04/23 M.Kaga END ---
    
End Type

'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : Cu-deco(CJ, CJ2)の部位別規格値構造体(製作条件xodc5_osf31)
Public Type typ_SB_com_xodb5_osf31_Cudeco
    
    JDGEIDC5            As String * 4       ' 判定D(C-OSF3, Cu-Deco)
    POSC5               As Long             ' 部位
    
    CJDMAXPIC5          As Integer          ' CJ Diskのみパターン Pi幅上限
    CJRMAXPIC5          As Integer          ' CJ Ringのみパターン Pi幅上限
    CJDRMAXPIC5         As Integer          ' CJ DiskRingパターン Pi幅上限
    CJALLMAXDIC5        As Integer          ' CJ 共通Disk半径上限
    CJALLMINRINC5       As Integer          ' CJ 共通Ring内径下限
    CJALLMAXRIGC5       As Integer          ' CJ 共通Ring外径上限
    
    CJ2DMAXPIC5         As Integer          ' CJ2 Diskのみパターン Pi幅下限(MAXだが下限です)
    CJ2RMAXPIC5         As Integer          ' CJ2 Ringのみパターン Pi幅下限(MAXだが下限です)
    CJ2RMINRINC5        As Integer          ' CJ2 Ringのみパターン Ring内径下限
    CJ2RMAXRIGC5        As Integer          ' CJ2 Ringのみパターン Ring外径上限
    CJ2DRMAXPIC5        As Integer          ' CJ2 DiskRingパターン Pi幅下限(MAXだが下限です)
    CJ2DRMINRINC5       As Integer          ' CJ2 DiskRingパターン Ring内径下限
    CJ2DRMAXRIGC5       As Integer          ' CJ2 DiskRingパターン Ring外径上限

End Type
''Add End   2011/01/17 SMPK A.Nagamine

Public JudgKoutei           As String       '工程(結晶実績ﾁｪｯｸ用)　08/04/15 ooba

'--------------- 2008/07/25 INSERT START  By Systech ---------------
Public gsTbcmy028ErrCode    As String           ' 振替チェックエラーコード
'--------------- 2008/07/25 INSERT  END   By Systech ---------------

Public tWarpInitG() As typ_WarpKakuData     ' Warpﾃﾞｰﾀ(TBCMY018)      '05/12/18 ooba START ===>
Public tKakuInitG() As typ_WarpKakuData     ' 合成角度ﾃﾞｰﾀ(TBCMY018)
Public tWarpMeasG() As typ_WarpKakuData     ' Warpﾃﾞｰﾀ(表示/判定用)
Public tKakuMeasG() As typ_WarpKakuData     ' 合成角度ﾃﾞｰﾀ(表示/判定用)
Public tMapHinG     As typ_MapHinData       ' WFﾏｯﾌﾟ上の品番ﾃﾞｰﾀ       '05/12/18 ooba END =====>

'------------------------------------------------
' コードＤＢ取得共通関数
'------------------------------------------------

'概要      :指定された項目をキーに、コードマスター(TBCMB005)から該当するデータを取得する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :sSysclass     ,I  ,String       :ｼｽﾃﾑ区分('SB'固定)
'          :sClass        ,I  ,String       :区分
'          :sCode         ,I  ,String       :ｺｰﾄﾞ
'          :iForm         ,I  ,Integer      :取得形式(0:50ﾊﾞｲﾄﾃﾞｰﾀ, 1:1ﾊﾞｲﾄﾃﾞｰﾀ)
'          :sSubCode      ,I  ,String       :ｻﾌﾞｺｰﾄﾞ(取得形式=1のみ有効)
'          :sResult       ,O  ,String       :取得ﾃﾞｰﾀ
'          :戻り値        ,O  ,Integer      :取得の成否(0:正常取得, -1:取得ｴﾗｰ)
'説明      :
'履歴      :2003/09/04 新規作成　システムブレイン

Public Function funCodeDBGet(sSysclass As String, sClass As String, sCode As String, iForm As Integer, sSubCode As String, sResult As String) As Integer
    Dim sql As String       'SQL全体
    Dim rs  As OraDynaset   'RecordSet

    'パラメータチェック
    If sSysclass = "" Or sSysclass = vbNullString Then GoTo CodeDBGetErr
    If sClass = "" Or sClass = vbNullString Then GoTo CodeDBGetErr
    If sCode = "" Or sCode = vbNullString Then GoTo CodeDBGetErr
    If iForm <> 0 And iForm <> 1 Then GoTo CodeDBGetErr
    If sSubCode = "" Or sSubCode = vbNullString Then GoTo CodeDBGetErr
    
    '取得形式 = 0(50ﾊﾞｲﾄﾃﾞｰﾀ)の場合
    If iForm = 0 Then
        sql = "select info1 from tbcmb005 where sysclass = '" & sSysclass & "' and class = '" & sClass & "' and code = '" & sCode & "'"
    
    '取得形式 = 1(1ﾊﾞｲﾄﾃﾞｰﾀ)の場合
    Else
        sql = "select substr(a1.info1, a2.info2, 1) as info1 from tbcmb005 a1, "
        sql = sql & "(select to_number(info2) as info2 from tbcmb005 "
        sql = sql & " where sysclass = '" & sSysclass & "' and class = '" & sClass & "' and code = '" & sSubCode & "') a2 "
        sql = sql & "where a1.sysclass = '" & sSysclass & "' and a1.class = '" & sClass & "' and a1.code = '" & sCode & "'"
    End If
    
    'SQL文の実行
Debug.Print sql
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        GoTo CodeDBGetErr
    End If
    
    '取得データセット
    sResult = rs("info1")
    Set rs = Nothing

    funCodeDBGet = 0
    
    Exit Function

CodeDBGetErr:
    funCodeDBGet = -1
    Set rs = Nothing
End Function

'概要      :指定された項目をキーに、マトリックスからOK/NGを返す
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :sSysclass     ,I  ,String       :ｼｽﾃﾑ区分('SB'固定)
'          :sClass        ,I  ,String       :区分
'          :sCode1        ,I  ,String       :ｺｰﾄﾞ1(マトリックス縦軸)
'          :sCode2        ,I  ,String       :ｺｰﾄﾞ2(マトリックス横軸)
'          :戻り値        ,O  ,Integer      :取得の成否(1:OK(正常取),0:NG(正常取), -1:取得ｴﾗｰ)
'説明      :コードDBに登録されているマトリックスのCodeを取得し、
'           そのコードに無い値を指定した場合はスペースに置き換えマトリックスからOK/NGを取得する
'履歴      :2006/02/10 新規作成　SMP石川
Public Function funCodeDBGetMatrixReturn(sSysclass As String, sClass As String, sCode1 As String, sCode2 As String) As Integer
    Dim liRet           As Integer
    Dim sResult         As String       'コードＤＢ取得関数の取得変数
    Dim lsCodeList()    As String       'コードDBのCode一覧
    Dim llCnt           As Long
    Dim lsCode(1)       As String
    Dim liLoopCnt       As Integer
    
    funCodeDBGetMatrixReturn = -1
    
    lsCode(0) = Trim(sCode1)
    lsCode(1) = Trim(sCode2)
    
    '' コードマスタのコードの一覧を取得
    liRet = funCodeDBGetCodeList(sSysclass, sClass, lsCodeList)
    If liRet <> 0 Then
        funCodeDBGetMatrixReturn = -1
        Exit Function
    Else
        ''コードマスタに登録されていないコードはスペースに変換する
        For liLoopCnt = 0 To 1
            liRet = 0
            For llCnt = 1 To UBound(lsCodeList)
                If Trim(lsCodeList(llCnt)) = Trim(lsCode(liLoopCnt)) Then
                    liRet = 1
                    Exit For
                End If
            Next llCnt
            If liRet = 0 Or Trim(lsCode(liLoopCnt)) = "" Then
                lsCode(liLoopCnt) = "     "
            End If
        Next liLoopCnt
        
        liRet = funCodeDBGet(sSysclass, sClass, lsCode(0), 1, lsCode(1), sResult)
        
        If liRet <> 0 Then
            funCodeDBGetMatrixReturn = -1
            Exit Function
        End If
        
        If sResult = 0 Then
            funCodeDBGetMatrixReturn = 0
            Exit Function
        End If
    End If
    
    funCodeDBGetMatrixReturn = 1
End Function



'概要      :指定された項目をキーに、コードマスター(TBCMB005)からCODEの一覧を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :sSysclass     ,I  ,String       :ｼｽﾃﾑ区分('SB'固定)
'          :sClass        ,I  ,String       :区分
'          :sCode()       ,O  ,String       :ｺｰﾄﾞの一覧
'          :戻り値        ,O  ,Integer      :取得の成否(0:正常取得, -1:取得ｴﾗｰ)
'説明      :
'履歴      :2006/02/10 新規作成　SMP石川

Public Function funCodeDBGetCodeList(sSysclass As String, sClass As String, sCode() As String) As Integer
    Dim sql As String       'SQL全体
    Dim rs  As OraDynaset   'RecordSet

    'パラメータチェック
    If sSysclass = "" Or sSysclass = vbNullString Then GoTo CodeDBGetErr
    If sClass = "" Or sClass = vbNullString Then GoTo CodeDBGetErr
    
    '初期化
    ReDim sCode(0) As String
    
    sql = ""
    sql = sql & " SELECT"
    sql = sql & "   code"
    sql = sql & " FROM"
    sql = sql & "   tbcmb005"
    sql = sql & " WHERE sysclass = '" & sSysclass & "'"
    sql = sql & "   AND class    = '" & sClass & "'"
    
    'SQL文の実行
Debug.Print sql
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    Do Until rs.EOF 'データがなくなるまで取得
        ReDim Preserve sCode(UBound(sCode) + 1) As String
        '取得データセット
        sCode(UBound(sCode)) = Trim(rs("code"))
        rs.MoveNext
    Loop
    
    'データ無しの場合
    If UBound(sCode) = 0 Then
        GoTo CodeDBGetErr
    End If
    
    Set rs = Nothing

    funCodeDBGetCodeList = 0
    
    Exit Function

CodeDBGetErr:
    funCodeDBGetCodeList = -1
    Set rs = Nothing
End Function

'------------------------------------------------
' 結晶ｻﾝﾌﾟﾙとＷＦｻﾝﾌﾟﾙ紐付け共通関数
'------------------------------------------------

'概要      :指定されたWFｻﾝﾌﾟﾙ情報から、対応する結晶ｻﾝﾌﾟﾙを検索し返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               :説明
'          :tWfSmpl       ,I  ,typ_Wf_Smpl      :WFｻﾝﾌﾟﾙ構造体
'          :tCrySmpl      ,O  ,typ_Cry_Smpl     :結晶ｻﾝﾌﾟﾙ構造体
'          :戻り値        ,O  ,Integer          :取得の成否(0:正常取得, -1:該当ﾌﾞﾛｯｸなし)
'説明      :
'履歴      :2003/09/04 新規作成　システムブレイン

Public Function funConSxl_Wf_Sampl(tWfSmpl As typ_Wf_Smpl, tCrySmpl As typ_Cry_Smpl) As Integer
    Dim sql     As String       'SQL全体
    Dim rs      As OraDynaset   'RecordSet

    'SQL文の編集
    'ﾌﾞﾛｯｸID,生死区分条件追加 08/08/11 ooba
    sql = "select CRYNUMCS, SMPKBNCS, TBKBNCS, REPSMPLIDCS, XTALCS, INPOSCS, HINBCS, REVNUMCS, FACTORYCS, OPECS from XSDCS "

    'TOP位置(T/B区分='T')の検索
    If tWfSmpl.TBKBNCW = "T" Then
        sql = sql & "where crynumcs like '" & Mid(tWfSmpl.XTALCW, 1, 9) & "%' and "
        sql = sql & "      tbkbncs = '" & tWfSmpl.TBKBNCW & "' and "
        sql = sql & "      xtalcs = '" & tWfSmpl.XTALCW & "' and "
        sql = sql & "      livkcs = '0' and "
        sql = sql & "      inposcs = (select max(inposcs) from xsdcs "
        sql = sql & "                 where  crynumcs like '" & Mid(tWfSmpl.XTALCW, 1, 9) & "%' and "
        sql = sql & "                        tbkbncs = '" & tWfSmpl.TBKBNCW & "' and "
        sql = sql & "                        xtalcs = '" & tWfSmpl.XTALCW & "' and "
        sql = sql & "                        livkcs = '0' and "
        sql = sql & "                        inposcs <= '" & tWfSmpl.INPOSCW & "')"
    
    'BOT位置(T/B区分='B')の検索
    ElseIf tWfSmpl.TBKBNCW = "B" Then
        sql = sql & "where crynumcs like '" & Mid(tWfSmpl.XTALCW, 1, 9) & "%' and "
        sql = sql & "      tbkbncs = '" & tWfSmpl.TBKBNCW & "' and "
        sql = sql & "      xtalcs = '" & tWfSmpl.XTALCW & "' and "
        sql = sql & "      livkcs = '0' and "
        sql = sql & "      inposcs = (select min(inposcs) from xsdcs "
        sql = sql & "                 where  crynumcs like '" & Mid(tWfSmpl.XTALCW, 1, 9) & "%' and "
        sql = sql & "                        tbkbncs = '" & tWfSmpl.TBKBNCW & "' and "
        sql = sql & "                        xtalcs = '" & tWfSmpl.XTALCW & "' and "
        sql = sql & "                        livkcs = '0' and "
        sql = sql & "                        inposcs >= '" & tWfSmpl.INPOSCW & "')"
    Else
        funConSxl_Wf_Sampl = -1
        Exit Function
    End If
    
    'SQL文の実行
Debug.Print sql
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    '該当データなし
    If rs.EOF Or rs.RecordCount > 1 Then
        funConSxl_Wf_Sampl = -1
        Exit Function
    End If
    
    '取得データセット
    With tCrySmpl
        .CRYNUMCS = rs("CRYNUMCS")          ' ﾌﾞﾛｯｸID
        .SMPKBNCS = rs("SMPKBNCS")          ' ｻﾝﾌﾟﾙ区分
        .TBKBNCS = rs("TBKBNCS")            ' T/B区分
        .REPSMPLIDCS = rs("REPSMPLIDCS")    ' 代表ｻﾝﾌﾟﾙID
        .XTALCS = rs("XTALCS")              ' 結晶番号
        .INPOSCS = rs("INPOSCS")            ' 結晶内位置
        .HINBCS = rs("HINBCS")              ' 品番
        .REVNUMCS = rs("REVNUMCS")          ' 製品番号改訂番号
        .FACTORYCS = rs("FACTORYCS")        ' 工場
        .OPECS = rs("OPECS")                ' 操業条件
    End With
    Set rs = Nothing

    funConSxl_Wf_Sampl = 0

End Function

'概要      :工程実績から振替時の結晶内位置/長さ/結晶番号を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO  ,型                :説明
'          :sLotid         ,I   ,String            :ﾌﾞﾛｯｸID or SXL_ID
'          :iKcnt          ,I   ,Integer           :工程連番
'          :iIngotpos       ,O   ,Integer          :結晶内位置
'          :iLength         ,O   ,Integer          :長さ
'          :sCrynum         ,O   ,String           :結晶番号
'          :戻り値          ,O   ,FUNCTION_RETURN   :抽出の成否
'説明      :
'履歴      :2003/11/07 ooba
Public Function GET_hurikaeC3(sLotid As String, iKcnt As Integer, iIngotPos As Integer, _
                                iLength As Integer, sCryNum As String) As FUNCTION_RETURN

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmec064.bas -- Function GET_hurikaeC3"
    GET_hurikaeC3 = FUNCTION_RETURN_FAILURE

    Dim sSql As String
    Dim rs As OraDynaset
    
    GET_hurikaeC3 = FUNCTION_RETURN_FAILURE
    
    sSql = "select min(INPOSC3), sum(LENC3), XTALC3 "
    sSql = sSql & "from XSDC3 "
    If Len(sLotid) = 12 Then
        sSql = sSql & "where CRYNUMC3 = '" & sLotid & "' "
    ElseIf Len(sLotid) = 13 Then
        sSql = sSql & "where SXLIDC3 = '" & sLotid & "' "
    Else
        GET_hurikaeC3 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    sSql = sSql & "and KCNTC3 = " & iKcnt
    sSql = sSql & "and substr(KNKTC3, 5, 1) = '3' "
    sSql = sSql & "group by XTALC3 "
    
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    
    If rs.RecordCount <> 1 Then
        Set rs = Nothing
        GET_hurikaeC3 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    Else
        If IsNull(rs("min(INPOSC3)")) = False Then iIngotPos = rs("min(INPOSC3)")
        If IsNull(rs("sum(LENC3)")) = False Then iLength = rs("sum(LENC3)")
        If IsNull(rs("XTALC3")) = False Then sCryNum = rs("XTALC3")
    End If

    Set rs = Nothing
    
    GET_hurikaeC3 = FUNCTION_RETURN_SUCCESS


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

'概要      :SXL確定指示(TBCMY007)ﾃｰﾌﾞﾙにｾｯﾄするSXLの比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝを求める。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO  ,型                :説明
'          :HIN            ,I   ,tFullHinban       ,12桁品番
'　　      :sPos  　　　    ,I   ,String 　         ,SXL位置(TOP/BOT)   04/04/15 ooba
'          :sPattern       ,O   ,String            ,比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝ
'                                                   ●ﾊﾟﾀｰﾝA : WF実績ﾃﾞｰﾀ取得
'                                                   ●ﾊﾟﾀｰﾝB : 結晶実績ﾃﾞｰﾀ取得
'                                                   ●ﾊﾟﾀｰﾝC : 取得ﾃﾞｰﾀなし
'          :戻り値          ,O   ,FUNCTION_RETURN   ,抽出の成否
'説明      :
'履歴      :04/02/12 ooba　作成
Public Function SxlRsPattern(HIN As tFullHinban, sPos As String, sPattern As String) As FUNCTION_RETURN

    Dim HSXRHWYS As String      '品ＳＸ比抵抗保証方法＿処
    Dim HWFRHWYS As String      '品ＷＦ比抵抗保証方法＿処
    Dim HWFRKHNN As String      '品ＷＦ比抵抗検査頻度＿抜　04/04/15 ooba
    Dim sSql As String
    Dim rs As OraDynaset
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "SB_com.bas -- Function SxlRsPattern"
    SxlRsPattern = FUNCTION_RETURN_FAILURE
    
    sPattern = "C"
    
    If Trim(HIN.hinban) <> "" And Trim(HIN.hinban) <> "Z" Then
        '該当品番の比抵抗(Rs)仕様を取得
        sSql = "select HSXRHWYS, HWFRHWYS, HWFRKHNN "
        sSql = sSql & "from TBCME018, TBCME021 "
        sSql = sSql & "where TBCME018.HINBAN = TBCME021.HINBAN "
        sSql = sSql & "and TBCME018.MNOREVNO = TBCME021.MNOREVNO "
        sSql = sSql & "and TBCME018.FACTORY = TBCME021.FACTORY "
        sSql = sSql & "and TBCME018.OPECOND = TBCME021.OPECOND "
        sSql = sSql & "and TBCME018.HINBAN = '" & HIN.hinban & "' "
        sSql = sSql & "and TBCME018.MNOREVNO = " & HIN.mnorevno & " "
        sSql = sSql & "and TBCME018.FACTORY = '" & HIN.factory & "' "
        sSql = sSql & "and TBCME018.OPECOND = '" & HIN.opecond & "' "
        
        Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
        
        If rs.RecordCount <> 1 Then
            Set rs = Nothing
            GoTo proc_exit
        Else
            If IsNull(rs("HSXRHWYS")) = False Then HSXRHWYS = rs("HSXRHWYS")   '品ＳＸ比抵抗保証方法＿処
            If IsNull(rs("HWFRHWYS")) = False Then HWFRHWYS = rs("HWFRHWYS")   '品ＷＦ比抵抗保証方法＿処
            If IsNull(rs("HWFRKHNN")) = False Then HWFRKHNN = rs("HWFRKHNN")   '品ＷＦ比抵抗検査頻度＿抜　04/04/15 ooba
        End If
        
        Set rs = Nothing
    Else
        GoTo proc_exit
    End If
    
    '保証方法ﾁｪｯｸ追加　04/04/15 ooba
'    If HWFRHWYS = "H" Then
    If HWFRHWYS = "H" And CheckKHN(HWFRKHNN, 1, sPos) Then
        'WF仕様『H』の場合
        sPattern = "A"
    ElseIf HWFRHWYS = "S" And CheckKHN(HWFRKHNN, 1, sPos) Then
        If HSXRHWYS = "H" Then
            'WF仕様『S』で結晶仕様『H』の場合
            sPattern = "B"
        Else
            'WF仕様『S』で結晶仕様『H』以外の場合
            sPattern = "A"
        End If
    Else
        If HSXRHWYS = "H" Or HSXRHWYS = "S" Then
            'WF仕様なしで結晶仕様『H』『S』の場合
            sPattern = "B"
        End If
    End If
    
    Set rs = Nothing
    SxlRsPattern = FUNCTION_RETURN_SUCCESS
    
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

'概要      :保証方法(検査頻度＿抜)のチェック
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型       ,説明
'　　      :sKHN  　　　,I  ,String 　,保証方法(検査頻度＿抜)
'　　      :iItemNo     ,I  ,Integer  ,検査項目　1:Rs  2:Oi  3:OSF1  4:OSF2  5:OSF3  6:OSF4
'                                               7:BMD1  8:BMD2  9:BMD3  10:Doi1  11:Doi2
'                                               12:Doi3 13:Dsod  14:DZ  15:SPVFE  16:SPV拡
'                                               17:Aoi 18:GD 19:SPVNR
'　　      :sPos  　　　,I  ,String 　,SXL位置(TOP/BOT/MID)
'      　　:戻り値      ,O  ,Boolean　,検査の有無
'説明      :保証方法(検査頻度＿抜)をチェックして検査の有無を返す
'履歴      :04/04/07 ooba
'          :GD追加　05/01/20 ooba
'          :SPVNR追加　06/06/08 ooba
Public Function CheckKHN(sKHN As String, iItemNo As Integer, sPos As String) As Boolean
    Dim RET     As Integer
    Dim sChkPtn As String
    Dim sResult As String
    Dim iChk    As Integer
    
'    CheckKHN = False
    CheckKHN = True '04/05/26 ooba
    sChkPtn = ""
    sResult = ""
    If sPos <> "TOP" And sPos <> "BOT" Then Exit Function
    
    'ｺｰﾄﾞDBより保証方法ﾁｪｯｸの有無情報を取得する
    RET = funCodeDBGet("SB", "HO", "PTN", 0, " ", sChkPtn)
    If RET <> 0 Then Exit Function
    
    If Mid(sChkPtn, iItemNo, 1) = "1" Then
        'ｺｰﾄﾞDBより保証方法ﾁｪｯｸﾊﾟﾀｰﾝを取得する
        RET = funCodeDBGet("SB", "HO", sPos, 0, " ", sResult)
        If RET <> 0 Then Exit Function
        
        '取得ﾊﾟﾀｰﾝより検査の有無を判断する
        Select Case sKHN
        Case "3"    'TOP保証
            iChk = 1
        Case "4"    'BOT保証
            iChk = 2
        Case "6"    'T/B保証
            iChk = 3
        Case Else   'なし(NULL,ｽﾍﾟｰｽ,346以外)
            iChk = 4
        End Select
        
        If Mid(sResult, iChk, 1) = "1" Then CheckKHN = True Else CheckKHN = False
    Else
        CheckKHN = True
    End If
    
End Function

'概要      :保証方法(検査頻度＿抜)のチェック(エピ用)
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型       ,説明
'　　      :sKHN  　　　,I  ,String 　,保証方法(検査頻度＿抜)
'　　      :iItemNo     ,I  ,Integer  ,検査項目　1:BMD1E  2:BMD2E  3:BMD3E  4:OSF1E  5:OSF2E  6:OSF3E
'　　      :sPos  　　　,I  ,String 　,SXL位置(TOP/BOT)
'      　　:戻り値      ,O  ,Boolean　,検査の有無
'説明      :保証方法(検査頻度＿抜)をチェックして検査の有無を返す
'履歴      :06/08/15 SMP)kondoh 新規作成
Public Function CheckKHN_EP(sKHN As String, iItemNo As Integer, sPos As String) As Boolean
    Dim RET     As Integer
    Dim sChkPtn As String
    Dim sResult As String
    Dim iChk    As Integer
    
    CheckKHN_EP = True
    sChkPtn = ""
    sResult = ""
    If sPos <> "TOP" And sPos <> "BOT" Then Exit Function
    
    'ｺｰﾄﾞDBより保証方法ﾁｪｯｸの有無情報を取得する
    RET = funCodeDBGet("SB", "HO", "PTNE", 0, " ", sChkPtn)
    If RET <> 0 Then Exit Function
    
    If Mid(sChkPtn, iItemNo, 1) = "1" Then
        'ｺｰﾄﾞDBより保証方法ﾁｪｯｸﾊﾟﾀｰﾝを取得する
        RET = funCodeDBGet("SB", "HO", sPos, 0, " ", sResult)
        If RET <> 0 Then Exit Function
        
        '取得ﾊﾟﾀｰﾝより検査の有無を判断する
        Select Case sKHN
        Case "3"    'TOP保証
            iChk = 1
        Case "4"    'BOT保証
            iChk = 2
        Case "6"    'T/B保証
            iChk = 3
        Case Else   'なし(NULL,ｽﾍﾟｰｽ,346以外)
            iChk = 4
        End Select
        
        If Mid(sResult, iChk, 1) = "1" Then CheckKHN_EP = True Else CheckKHN_EP = False
    Else
        CheckKHN_EP = True
    End If
    
End Function

'概要      :ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞの取得
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO   ,型                ,説明
'　　      :HIN  　　 　,I    ,tFullHinban 　    ,12桁品番
'　　      :sBflg  　　 ,O    ,String 　         ,ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞ
'      　　:戻り値      ,O    ,FUNCTION_RETURN   ,抽出の成否
'説明      :
'履歴      :05/01/20 ooba
Public Function chkBlkTanFlg(HIN As tFullHinban, sBflg As String) As FUNCTION_RETURN

    Dim sSql As String
    Dim rs As OraDynaset
    
    chkBlkTanFlg = FUNCTION_RETURN_FAILURE
        
    sBflg = ""
    If Trim(HIN.hinban) = "Z" Or Trim(HIN.hinban) = "G" Or Trim(HIN.hinban) = "" Then
        Exit Function
    End If
    
    sSql = "SELECT BLOCKHFLAG"
    sSql = sSql & " FROM TBCME036"
    sSql = sSql & " WHERE"
    sSql = sSql & " HINBAN = '" & HIN.hinban & "'"
    sSql = sSql & " AND MNOREVNO = " & HIN.mnorevno
    sSql = sSql & " AND FACTORY = '" & HIN.factory & "'"
    sSql = sSql & " AND OPECOND = '" & HIN.opecond & "'"
    
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)

    If rs.RecordCount > 0 Then
        If IsNull(rs.Fields("BLOCKHFLAG")) = False Then sBflg = rs.Fields("BLOCKHFLAG")
        chkBlkTanFlg = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
End Function

'概要      :WFｶｯﾄ単位の取得
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO   ,型                ,説明
'　　      :HIN  　　 　,I    ,tFullHinban 　    ,12桁品番
'　　      :iWCtani  　 ,O    ,Integer 　        ,WFｶｯﾄ単位
'      　　:戻り値      ,O    ,FUNCTION_RETURN   ,抽出の成否
'説明      :
'履歴      :05/04/19 ooba
Public Function getWFCUTT(HIN As tFullHinban, iWCtani As Integer) As FUNCTION_RETURN

    Dim sSql As String
    Dim rs As OraDynaset
    
    getWFCUTT = FUNCTION_RETURN_FAILURE
        
    iWCtani = -1
    If Trim(HIN.hinban) = "Z" Or Trim(HIN.hinban) = "G" Or Trim(HIN.hinban) = "" Then
        Exit Function
    End If
    
    sSql = "SELECT WFCUTT"
    sSql = sSql & " FROM TBCME036"
    sSql = sSql & " WHERE"
    sSql = sSql & " HINBAN = '" & HIN.hinban & "'"
    sSql = sSql & " AND MNOREVNO = " & HIN.mnorevno
    sSql = sSql & " AND FACTORY = '" & HIN.factory & "'"
    sSql = sSql & " AND OPECOND = '" & HIN.opecond & "'"
    
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)

    If rs.RecordCount > 0 Then
        If IsNull(rs.Fields("WFCUTT")) = False Then iWCtani = rs.Fields("WFCUTT") Else iWCtani = -1
        getWFCUTT = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
End Function
'概要      :SIRD評価情報の取得
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO   ,型                ,説明
'　　      :HIN  　　 　,I    ,tFullHinban 　    ,12桁品番
'　　      :sSIRDFLG 　 ,O    ,String 　   　    ,SIRD評価フラグ
'      　　:戻り値      ,O    ,FUNCTION_RETURN   ,抽出の成否
'説明      :
'履歴      :2010/01/18 Y.Hitomi
Public Function getSDFlg(HIN As tFullHinban, sSirdFlg As String) As FUNCTION_RETURN

    Dim sSql As String
    Dim rs As OraDynaset
    
    getSDFlg = FUNCTION_RETURN_FAILURE
        
    sSirdFlg = " "
    If Trim(HIN.hinban) = "Z" Or Trim(HIN.hinban) = "G" Or Trim(HIN.hinban) = "" Then
        Exit Function
    End If
    
    sSql = "SELECT HWFSIRDHS"
    sSql = sSql & " FROM TBCME048"
    sSql = sSql & " WHERE"
    sSql = sSql & " HINBAN = '" & HIN.hinban & "'"
    sSql = sSql & " AND MNOREVNO = " & HIN.mnorevno
    sSql = sSql & " AND FACTORY = '" & HIN.factory & "'"
    sSql = sSql & " AND OPECOND = '" & HIN.opecond & "'"
    
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)

    If rs.RecordCount > 0 Then
        If IsNull(rs.Fields("HWFSIRDHS")) = False Then sSirdFlg = rs.Fields("HWFSIRDHS") Else sSirdFlg = " "
        getSDFlg = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
End Function

'概要      :結晶検査実績の引継ぎ処理
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO   ,型                ,説明
'　　      :sBlockID  　,I    ,String      　    ,ﾌﾞﾛｯｸID
'　　      :sTBkbn  　　,I    ,String      　    ,TB区分
'　　      :iInPos  　　,I    ,Integer　　 　    ,結晶内位置
'　　      :HIN  　　 　,I    ,tFullHinban 　    ,12桁品番
'　　      :iItemNo   　,I    ,Integer 　        ,検査項目　(1:GD)
'　　      :sSampleID   ,O    ,String 　         ,結晶ｻﾝﾌﾟﾙID
'　　      :sCryind     ,O    ,String 　         ,結晶状態FLG
'      　　:戻り値      ,O    ,FUNCTION_RETURN   ,抽出の成否
'説明      :ﾌﾞﾛｯｸID／品番／TB区分／結晶内位置を元にｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ)から結晶ｻﾝﾌﾟﾙID／結晶状態FLGを取得する
'履歴      :05/01/20 ooba
Public Function funBlkSmpDataGet(sBlockId As String, sTBkbn As String, _
                                    iInpos As Integer, HIN As tFullHinban, _
                                    iItemNo As Integer, sSampleid As String, _
                                    sCryind As String) As FUNCTION_RETURN
                                                          
    Dim sKensa As String        '検査項目名
    Dim sSql As String
    Dim rs As OraDynaset
    
    funBlkSmpDataGet = FUNCTION_RETURN_FAILURE
        
    sSampleid = ""
    sCryind = ""
    
    '検査項目をｾｯﾄ
    Select Case iItemNo
    Case 1  'GD
        sKensa = "GD"
    Case Else
        Exit Function
    End Select
    
    '結晶ｻﾝﾌﾟﾙIDの取得
    sSql = "SELECT"
    sSql = sSql & " CRYSMPLID" & sKensa & "CS, "                    'ｻﾝﾌﾟﾙID
    sSql = sSql & " CRYIND" & sKensa & "CS"                         '状態FLG
    sSql = sSql & " FROM XSDCS"
    sSql = sSql & " WHERE"
    sSql = sSql & " CRYNUMCS = '" & sBlockId & "'"                  'ﾌﾞﾛｯｸID
    sSql = sSql & " AND TBKBNCS = '" & sTBkbn & "'"                 'T/B区分
    If sTBkbn = "T" Then                                            '結晶内位置
        sSql = sSql & " AND INPOSCS <= " & iInpos
    ElseIf sTBkbn = "B" Then
        sSql = sSql & " AND INPOSCS >= " & iInpos
    End If
    If Trim(HIN.hinban) <> "" Then      'CW740/CW760の場合は品番の条件無し　05/06/13 ooba
        sSql = sSql & " AND HINBCS = '" & HIN.hinban & "'"          '品番
        sSql = sSql & " AND REVNUMCS = " & HIN.mnorevno             '製品番号改訂番号
        sSql = sSql & " AND FACTORYCS = '" & HIN.factory & "'"      '工場
        sSql = sSql & " AND OPECS = '" & HIN.opecond & "'"          '操業条件
    End If
    sSql = sSql & " AND CRYIND" & sKensa & "CS IN ('1', '2')"       '状態FLG
'    sSql = sSql & " AND CRYRES" & sKensa & "CS = '1'"               '実績FLG
    '条件変更　05/07/20 ooba
    sSql = sSql & " AND CRYRES" & sKensa & "CS IN ('1', '2')"       '実績FLG
    
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    
    '結晶ｻﾝﾌﾟﾙID／結晶状態FLGをｾｯﾄ
    If rs.RecordCount > 0 Then
        If IsNull(rs.Fields(0)) = False Then sSampleid = CStr(rs.Fields(0))
        If IsNull(rs.Fields(1)) = False Then sCryind = CStr(rs.Fields(1))
        funBlkSmpDataGet = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close
    
End Function

'------------------------------------------------
' Nullﾁｪｯｸ共通関数
'------------------------------------------------

'概要      :指定された値がNullなら-1を返し、Null以外なら指定された値を返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :vrnN          ,I  ,Variant      :指定値
'          :戻り値        ,O  ,Double       :指定値、または、-1
'説明      :
'履歴      :2003/12/08 新規作成　システムブレイン

Public Function fncNullCheck(vrnN As Variant) As Double 'Nullのチェックをする
    If IsNull(vrnN) = False Then
        fncNullCheck = vrnN 'NULLじゃないときはそのまま
    Else
        fncNullCheck = -1  'NULLのときは-1を入れる
    End If
End Function

'------------------------------------------------
' Null対応表示共通関数
'------------------------------------------------

'概要      :指定された値が-1(Null値の代わり)ならvbNullStringを返し、-1以外なら指定値を返す。ﾌｫｰﾏｯﾄ指定がある場合には、指定ﾌｫｰﾏｯﾄで返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :data          ,I  ,Variant      :指定値
'          :Formatstr     ,I  ,String       :ﾌｫｰﾏｯﾄ形式(省略可)
'          :戻り値        ,O  ,Variant      :戻り値
'説明      :
'履歴      :2003/12/09 新規作成　システムブレイン

Public Function DBData2DispData_nl(data As Variant, Optional Formatstr As String) As Variant   'NULL対応用 2003/12/9
    If data = -1 Then
'        DBData2DispData = ""
        DBData2DispData_nl = vbNullString
    Else
        If Formatstr = "" Then
            DBData2DispData_nl = data
        Else
            DBData2DispData_nl = Format(data, Formatstr)
        End If
    End If
End Function

'''■■■■s_cmzcjudg.bas に移動            2003/12/18 tuku
'''概要      :測定値がNULLだった場合の範囲判定を行う。
'''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'''          :JudgData      ,I  ,double    ,測定値
'''          :SpecMin       ,I  ,double    ,下限値
'''          :SpecMax       ,I  ,double    ,上限値
'''          :戻り値        ,O  ,Boolean   ,JUDG_OK or JUDG_NG
'''説明      :
'''履歴      :2003/12/11 新規作成 システムブレイン
''Public Function RangeDecision_nl(JudgData As Double, SpecMin As Double, SpecMax As Double) As Boolean
''    RangeDecision_nl = False
''    If (JudgData >= SpecMin) Or (SpecMin = -1) Then
''        If (JudgData <= SpecMax) Or (SpecMax = -1) Then
''            RangeDecision_nl = True
''        End If
''    End If
'''    RangeDecision = ((JudgData >= SpecMin) And (JudgData <= SpecMax))
''End Function



'------------------------------------------------
' Null対応実績入力判定共通関数
'------------------------------------------------

'概要      :保証方法が'H'または'S'で、仕様値配列に-1が存在した場合、Falseを返す。それ以外の場合、Trueを返す。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           :説明
'          :Hosyo         ,I  ,String       :保証方法_対象
'          :Shiyo()       ,I  ,Double       :仕様値配列
'          :戻り値        ,O  ,Boolean      : True:OK, False:NG
'説明      :
'履歴      :2003/12/11 新規作成　システムブレイン

Public Function fncJissekiHantei_nl(Hosyo As String, Shiyo() As Double) As Boolean
    Dim cnt As Integer
    
    fncJissekiHantei_nl = True
'    If Hosyo = "H" Or Hosyo = "S" Then '保証方法Sはチェックしない　2003/12/19　tuku
    If Hosyo = "H" Then
        For cnt = 1 To UBound(Shiyo)
            If Shiyo(cnt) = -1 Then
                fncJissekiHantei_nl = False
                Exit For
            End If
        Next
    End If
End Function

'概要      :狙い品番を取得する。
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO   ,型                ,説明
'　　      :sBlockID  　,I    ,String      　    ,結晶番号
'　　      :HIN  　　 　,O    ,tFullHinban 　    ,12桁狙い品番
'      　　:戻り値      ,O    ,FUNCTION_RETURN   ,抽出の成否
'説明      :
'履歴      :06/04/25 ooba
Public Function funNeraiHinGet(sCryNum As String, tHIN As tFullHinban) As FUNCTION_RETURN

    Dim sSql As String
    Dim rs As OraDynaset
    
    funNeraiHinGet = FUNCTION_RETURN_FAILURE
        
    tHIN.hinban = ""
    tHIN.mnorevno = 0
    tHIN.factory = ""
    tHIN.opecond = ""
    tHIN.Hinkubun = ""
    
    sSql = "SELECT PUHINBC1, PUREVNUMC1, PUFACTORYC1, PUOPEC1 "
    sSql = sSql & "FROM XSDC1 "
    sSql = sSql & "WHERE XTALC1 = '" & sCryNum & "' "
    
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)

    If rs.RecordCount = 1 Then
        If Not IsNull(rs.Fields("PUHINBC1")) Then tHIN.hinban = rs.Fields("PUHINBC1")
        If Not IsNull(rs.Fields("PUREVNUMC1")) Then tHIN.mnorevno = rs.Fields("PUREVNUMC1")
        If Not IsNull(rs.Fields("PUFACTORYC1")) Then tHIN.factory = rs.Fields("PUFACTORYC1")
        If Not IsNull(rs.Fields("PUOPEC1")) Then tHIN.opecond = rs.Fields("PUOPEC1")
        funNeraiHinGet = FUNCTION_RETURN_SUCCESS
    End If
    rs.Close

End Function


'概要      :結晶引上(XSDC1)よりC－OSF3判定IDの獲得を行う
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :XTALC1        ,IO ,String           ,結晶番号
'          :JDGEIDC1      ,O  ,String           ,C－OSF3判定ID
'説明      :
'履歴      :2007/04/23 作成  加賀
Public Function GetCOSF3ID(C_JDGEIDC1 As String, C_XTALC1 As String) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String

    '初期化
    GetCOSF3ID = FUNCTION_RETURN_FAILURE

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''結晶引上(XSDC1)よりC－OSF3判定IDの獲得を行う
    sql = ""
    sql = "select XTALC1,JDGEIDC1 from XSDC1 where (trim(XTALC1)='" & Trim$(C_XTALC1) & "')"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    
    'ﾚｺｰﾄﾞ自体が存在しない場合
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GetCOSF3ID = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    Else
        '結晶番号
        C_XTALC1 = rs("XTALC1")
        'C－OSF3判定IDがNULLの場合
        If Trim(rs("JDGEIDC1")) = "" Or IsNull(rs("JDGEIDC1")) Then
            C_JDGEIDC1 = vbNullString
        Else
            C_JDGEIDC1 = rs("JDGEIDC1")
        End If
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If

    '正常終了
    GetCOSF3ID = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

End Function


'概要      :引上条件(XODC5_OSF30)より承認ﾌﾗｸﾞの獲得を行う
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :JDGEIDC5      ,IO ,String           ,判定ID
'          :SYNFLAGC5     ,O  ,String           ,承認ﾌﾗｸﾞ
'          :YNKFLAGC5     ,O  ,String           ,削除ﾌﾗｸﾞ
'説明      :
'履歴      :2007/04/23 作成  加賀
Public Function GetSYNFLAGC5(C_SYNFLAGC5 As String, C_YMKFLAGC5 As String, C_JDGEIDC1 As String) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String

    '初期化
    GetSYNFLAGC5 = FUNCTION_RETURN_FAILURE

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    ''結晶引上(XSDC1)よりC－OSF3判定IDの獲得を行う
    sql = ""
    sql = "select SYNFLAGC5,YMKFLAGC5 from XODC5_OSF30 where (trim(JDGEIDC5)='" & Trim$(C_JDGEIDC1) & "') and YMKFLAGC5 = '0'"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    
    'ﾚｺｰﾄﾞ自体が存在しない場合
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GetSYNFLAGC5 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    Else
        '承認ﾌﾗｸﾞがNULLの場合
        If Trim(rs("SYNFLAGC5")) = "" Or IsNull(rs("SYNFLAGC5")) Then
            C_SYNFLAGC5 = vbNullString
        Else
            C_SYNFLAGC5 = rs("SYNFLAGC5")
        End If
        '削除ﾌﾗｸﾞがNULLの場合
        If Trim(rs("YMKFLAGC5")) = "" Or IsNull(rs("YMKFLAGC5")) Then
            C_YMKFLAGC5 = vbNullString
        Else
            C_YMKFLAGC5 = rs("YMKFLAGC5")
        End If

    End If
    
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If
    
    '正常終了
    GetSYNFLAGC5 = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

End Function

'概要      :判定条件(XODC5_OSF31)より判定データの獲得を行う
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :strChkR       ,I  ,String           ,ﾊﾟﾀｰﾝ区分
'          :strChkD       ,I  ,String           ,ﾊﾟﾀｰﾝ区分
'          :JDGEIDC5      ,IO ,String           ,判定ID
'          :POSC5         ,IO ,Long             ,ｻﾝﾌﾟﾙ位置
'          :DMAXC5        ,O  ,Long             ,Dのみ上限
'          :RMAXC5        ,O  ,Long             ,Rのみ上限
'          :DRDMAXC5      ,O  ,Long             ,共存D上限
'          :DRRMAXC5      ,O  ,Long             ,共存R上限
'説明      :
'履歴      :2007/04/23 作成  加賀

Public Function GetCOSF3PTN(C_JDGEIDC5 As String, C_POSC5 As Long, C_strChkR As String, C_strChkD As String, C_RMAXC5 As String, C_DMAXC5 As String, C_DRRMAXC5 As String, C_DRDMAXC5 As String) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String


    '初期化
    GetCOSF3PTN = FUNCTION_RETURN_FAILURE

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    '判定条件(XODC5_OSF31)より判定データの獲得を行う
    'ﾊﾟﾗﾒｰﾀのﾊﾟﾀｰﾝ区分によって処理分岐
    'Rのみの場合
    If C_strChkR = "R" And C_strChkD = "-" Then
        'SQL編集
        sql = ""
        sql = "SELECT TO_CHAR(RMAXC5) as RMAXC5 FROM(select MIN(POSC5) as W_POSC5 from(select POSC5 from XODC5_OSF31 where (trim(JDGEIDC5)='" & Trim$(C_JDGEIDC5) & "') and POSC5 >='" & Trim$(C_POSC5) & "')  ),XODC5_OSF31 WHERE POSC5 = W_POSC5 AND (trim(JDGEIDC5)='" & Trim$(C_JDGEIDC5) & "')"
    'Dのみの場合
    ElseIf C_strChkR = "D" Then
        'SQL編集
        sql = ""
        sql = "SELECT TO_CHAR(DMAXC5) as DMAXC5 FROM(select MIN(POSC5) as W_POSC5 from(select POSC5 from XODC5_OSF31 where (trim(JDGEIDC5)='" & Trim$(C_JDGEIDC5) & "') and POSC5 >='" & Trim$(C_POSC5) & "')  ),XODC5_OSF31 WHERE POSC5 = W_POSC5 AND (trim(JDGEIDC5)='" & Trim$(C_JDGEIDC5) & "')"
    'R&Dの場合
    ElseIf C_strChkR = "R" And C_strChkD = "D" Then
        'SQL編集
        sql = ""
        sql = "SELECT TO_CHAR(DRRMAXC5) as DRRMAXC5,TO_CHAR(DRDMAXC5) as DRDMAXC5 FROM(select MIN(POSC5) as W_POSC5 from(select POSC5 from XODC5_OSF31 where (trim(JDGEIDC5)='" & Trim$(C_JDGEIDC5) & "') and POSC5 >='" & Trim$(C_POSC5) & "')  ),XODC5_OSF31 WHERE POSC5 = W_POSC5 AND (trim(JDGEIDC5)='" & Trim$(C_JDGEIDC5) & "')"
    End If
    
    'SQL実行
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
        
   'レコード自体が存在しない場合
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GetCOSF3PTN = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    Else
        '判定実績値がNULLの場合
        If C_strChkR = "R" And C_strChkD = "-" Then
            If Trim(rs("RMAXC5")) = "" Or IsNull(rs("RMAXC5")) Then
                C_RMAXC5 = vbNullString
            Else
                C_RMAXC5 = rs("RMAXC5")
            End If
        ElseIf C_strChkR = "D" Then
            If Trim(rs("DMAXC5")) = "" Or IsNull(rs("DMAXC5")) Then
                C_DMAXC5 = vbNullString
            Else
                C_DMAXC5 = rs("DMAXC5")
            End If
        ElseIf C_strChkR = "R" And C_strChkD = "D" Then
            If Trim(rs("DRRMAXC5")) = "" Or IsNull(rs("DRRMAXC5")) Then
                C_DRRMAXC5 = vbNullString
            Else
                C_DRRMAXC5 = rs("DRRMAXC5")
            End If
            If Trim(rs("DRDMAXC5")) = "" Or IsNull(rs("DRDMAXC5")) Then
                C_DRDMAXC5 = vbNullString
            Else
                C_DRDMAXC5 = rs("DRDMAXC5")
            End If
        End If
    End If
    
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If
    
    '正常終了
    GetCOSF3PTN = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function
    
End Function


'Add Start 2011/01/17 SMPK A.Nagamine   : CLESTA評価対応(Cu-deco) : Cu-deco(CJ, CJ2)の部位別規格値構造体(製作条件xodc5_osf31)取得関数
Public Function GetOsf31_CuDeco(pstrC_JDGEIDC5 As String, plngC_POSC5 As Long, ptyp_Ret As typ_SB_com_xodb5_osf31_Cudeco) As FUNCTION_RETURN
Dim dbIsMine As Boolean
Dim rs As OraDynaset
Dim sql As String


    '初期化
    GetOsf31_CuDeco = FUNCTION_RETURN_FAILURE

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    '判定条件(XODC5_OSF31)より判定データの獲得を行う
    'ﾊﾟﾗﾒｰﾀのﾊﾟﾀｰﾝ区分によって処理分岐
    sql = " select CJDMAXPIC5, CJRMAXPIC5, CJDRMAXPIC5, CJALLMAXDIC5, CJALLMINRINC5"
    sql = sql & ", CJALLMAXRIGC5, CJ2DMAXPIC5, CJ2RMAXPIC5, CJ2RMINRINC5, CJ2RMAXRIGC5"
    sql = sql & ", CJ2DRMAXPIC5, CJ2DRMINRINC5, CJ2DRMAXRIGC5, JDGEIDC5, POSC5"

    sql = sql & " FROM (select MIN(POSC5) as W_POSC5 from (select POSC5 from XODC5_OSF31 where (trim(JDGEIDC5)='" & Trim$(pstrC_JDGEIDC5) & "') and POSC5 >='" & Trim$(plngC_POSC5) & "')  ), XODC5_OSF31"
    sql = sql & " WHERE POSC5 = W_POSC5 AND (trim(JDGEIDC5)='" & Trim$(pstrC_JDGEIDC5) & "')"
    
    'SQL実行
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
        
   'レコード自体が存在しない場合
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GetOsf31_CuDeco = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    Else
        ptyp_Ret.JDGEIDC5 = rs("JDGEIDC5")                                  ' 判定ID
        ptyp_Ret.POSC5 = CInt(fncNullCheck(rs("POSC5")))                    ' 部位
        
        '判定実績値がNULLの場合
        ptyp_Ret.CJDMAXPIC5 = CInt(fncNullCheck(rs("CJDMAXPIC5")))          ' CJ Diskのみパターン Pi幅上限
        ptyp_Ret.CJRMAXPIC5 = CInt(fncNullCheck(rs("CJRMAXPIC5")))          ' CJ Ringのみパターン Pi幅上限
        ptyp_Ret.CJDRMAXPIC5 = CInt(fncNullCheck(rs("CJDRMAXPIC5")))        ' CJ DiskRingパターン Pi幅上限
        ptyp_Ret.CJALLMAXDIC5 = CInt(fncNullCheck(rs("CJALLMAXDIC5")))      ' CJ 共通Disk半径上限
        ptyp_Ret.CJALLMINRINC5 = CInt(fncNullCheck(rs("CJALLMINRINC5")))    ' CJ 共通Ring内径下限
        ptyp_Ret.CJALLMAXRIGC5 = CInt(fncNullCheck(rs("CJALLMAXRIGC5")))    ' CJ 共通Ring外径上限
        
        ptyp_Ret.CJ2DMAXPIC5 = CInt(fncNullCheck(rs("CJ2DMAXPIC5")))        ' CJ2 Diskのみパターン Pi幅下限(MAXだが下限です)
        ptyp_Ret.CJ2RMAXPIC5 = CInt(fncNullCheck(rs("CJ2RMAXPIC5")))        ' CJ2 Ringのみパターン Pi幅下限(MAXだが下限です)
        ptyp_Ret.CJ2RMINRINC5 = CInt(fncNullCheck(rs("CJ2RMINRINC5")))      ' CJ2 Ringのみパターン Ring内径下限
        ptyp_Ret.CJ2RMAXRIGC5 = CInt(fncNullCheck(rs("CJ2RMAXRIGC5")))      ' CJ2 Ringのみパターン Ring外径上限
        ptyp_Ret.CJ2DRMAXPIC5 = CInt(fncNullCheck(rs("CJ2DRMAXPIC5")))      ' CJ2 DiskRingパターン Pi幅下限(MAXだが下限です)
        ptyp_Ret.CJ2DRMINRINC5 = CInt(fncNullCheck(rs("CJ2DRMINRINC5")))    ' CJ2 DiskRingパターン Ring内径下限
        ptyp_Ret.CJ2DRMAXRIGC5 = CInt(fncNullCheck(rs("CJ2DRMAXRIGC5")))    ' CJ2 DiskRingパターン Ring外径上限
        
    End If
    
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If
    
    '正常終了
    GetOsf31_CuDeco = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function
    
End Function
''Add End   2011/01/17 SMPK A.Nagamine



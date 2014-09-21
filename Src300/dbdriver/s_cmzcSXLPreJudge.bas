Attribute VB_Name = "s_cmzcSXLPreJudge"
'品番の振替チェックは共通関数を使用するので削除する-------start iida 2003/09/03
'Option Explicit
'
''WF仕様取得構造体定義
'Public Type typ_Wfsiyou
'    HWFTYPE As String * 1             ' 品ＷＦタイプ
'    HWFCDIR As String * 1             ' 品ＷＦ結晶面方
'    HWFCDOP As String * 1             ' 品ＷＦ結晶ドープ
'    HWFDOP As String * 1              ' 品ＷＦドーパント
'End Type
'
''概要      :抜試以降での品番振替時に、タイプ等の判定を行う
''パラメータ    :変数名        ,IO ,型          ,説明
''          :crynum        ,I  ,String      ,結晶番号
''          :ingotpos      ,I  ,Integer     ,対象範囲の開始位置
''          :length        ,I  ,Integer     ,対象範囲の長さ
''          :hin           ,I  ,tFullHinban ,振替先の品番
''          :judge_ok      ,O  ,Boolean     ,判定結果
''          :itemNG        ,O  ,String      ,判定NGとなった項目
''          :戻り値        ,O  ,FUNCTION_RETURN, 判定の合否
''          :                   FUNCTION_RETURN_SUCCESS: 振替可
''          :                   FUNCTION_RETURN_FAILURE: 振替不可
''          :                                もしくは仕様エラー
''説明      :タイプ・方位・ドーパント について判定する
''履歴      :2002/03/26 筑 作成
''          :itemNGエラー内容
''               TYPE :タイプエラー
''               CDIR :方位エラー
''               CDOP :結晶ドープエラー
''               DOP  :ドーパントーエラー
''               E021   :DBエラー(E021,E022,E023仕様取得)
''               E042   :DBエラー(E042)
'Public Function SXLPreJudge(CRYNUM$, IngotPos%, LENGTH%, HIN As tFullHinban, judge_ok As Boolean, itemNG$) As FUNCTION_RETURN
'Dim dbIsMine As Boolean
'Dim rs As OraDynaset
'Dim sql As String
'Dim mHIN As tFullHinban               '振替前品番用構造体
'
'Dim Wsi  As typ_Wfsiyou            'WF仕様取得構造体
'Dim mWsi As typ_Wfsiyou            'WF仕様取得構造体(振替前品番用）
'
'    'エラーハンドラの設定
'    On Error GoTo PROC_ERR
'    gErr.Push "SXLPreJudge.bas -- Function SXLPreJudge"
'
'    If OraDB Is Nothing Then
'        dbIsMine = True
'        OraDBOpen
'    End If
'
'    SXLPreJudge = FUNCTION_RETURN_FAILURE
'
'    '振替後品番がZ、G品番の場合は、無条件でOK
'    If Trim(HIN.HINBAN) = "Z" Or Trim(HIN.HINBAN) = "G" Then
'        judge_ok = True
'        itemNG = "OK"
'        SXLPreJudge = FUNCTION_RETURN_SUCCESS
'        GoTo PROC_EXIT
'    End If
'
'    '振替前品番取得（SXL管理より）
'    sql = "select "
'    sql = sql & " E042.HINBAN, "
'    sql = sql & " E042.REVNUM, "
'    sql = sql & " E042.FACTORY, "
'    sql = sql & " E042.OPECOND "
'    sql = sql & " from "
'    sql = sql & " TBCME042 E042 "
'    sql = sql & " where  "
'    sql = sql & " E042.CRYNUM ='" & CRYNUM & "' and "
'    sql = sql & " E042.INGOTPOS >= "
'    sql = sql & "   (select MAX(INGOTPOS) "
'    sql = sql & "   from TBCME042 "
'    sql = sql & "   where  "
'    sql = sql & "   CRYNUM ='" & CRYNUM & "' and "
'    sql = sql & "   INGOTPOS <= " & IngotPos
'    sql = sql & "   GROUP BY CRYNUM  ) and "
'    sql = sql & " E042.INGOTPOS <" & IngotPos + LENGTH
'
'    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'
'    If rs.RecordCount <> 1 Then
'    '見つからなかったら、FUNCTION_RETURN_FAILUREを返す。
'        judge_ok = False
'        itemNG = "E042"
'        rs.Close
'        SXLPreJudge = FUNCTION_RETURN_FAILURE
'        GoTo PROC_EXIT
'    Else
'        With mHIN
'            .HINBAN = rs("HINBAN")
'            .mnorevno = rs("REVNUM")
'            .factory = rs("FACTORY")
'            .opecond = rs("OPECOND")
'        End With
'        judge_ok = True
'    End If
'    rs.Close
'
'Debug.Print "1 変更後品番(取得)'" & HIN.HINBAN & "' '" & HIN.mnorevno & "' '" & HIN.factory & "' '" & HIN.opecond & "'"
'Debug.Print "2 振替前品番'" & mHIN.HINBAN & "' '" & mHIN.mnorevno & "' '" & mHIN.factory & "' '" & mHIN.opecond & "'"
'
'
''振替前品番仕様取得
'    sql = "select "
'    sql = sql & " E021.HWFTYPE HWFTYPE, "
'    sql = sql & " E022.HWFCDIR HWFCDIR, "
'    sql = sql & " E021.HWFDOP HWFDOP, "
'    sql = sql & " E023.HWFCDOP HWFCDOP "
'    sql = sql & " from "
'    sql = sql & " TBCME021 E021, TBCME022 E022, TBCME023 E023 "
'    sql = sql & " where "
'    sql = sql & " E021.HINBAN='" & mHIN.HINBAN & "' and "
'    sql = sql & " E021.MNOREVNO=" & mHIN.mnorevno & " and "
'    sql = sql & " E021.FACTORY='" & mHIN.factory & "' and "
'    sql = sql & " E021.OPECOND='" & mHIN.opecond & "' and "
'    sql = sql & " E022.HINBAN='" & mHIN.HINBAN & "' and "
'    sql = sql & " E022.MNOREVNO=" & mHIN.mnorevno & " and "
'    sql = sql & " E022.FACTORY='" & mHIN.factory & "' and "
'    sql = sql & " E022.OPECOND='" & mHIN.opecond & "' and "
'    sql = sql & " E023.HINBAN='" & mHIN.HINBAN & "' and "
'    sql = sql & " E023.MNOREVNO=" & mHIN.mnorevno & " and "
'    sql = sql & " E023.FACTORY='" & mHIN.factory & "' and "
'    sql = sql & " E023.OPECOND='" & mHIN.opecond & "'"
'
'    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'
'    If rs.RecordCount <> 1 Then
'        '見つからなかったら、FUNCTION_RETURN_FAILUREを返す。
'        judge_ok = False
'        itemNG = "E042"
'        rs.Close
'        SXLPreJudge = FUNCTION_RETURN_FAILURE
'        GoTo PROC_EXIT
'    Else
'        ''見つかったら、最新の品番情報をセットして FUNCTION_RETURN_SUCCESSを返す
'        With mWsi
'            .HWFCDIR = rs("HWFCDIR")
'            .HWFCDOP = rs("HWFCDOP")
'            .HWFTYPE = rs("HWFTYPE")
'            .HWFDOP = rs("HWFDOP")
'        End With
'        judge_ok = True
'        rs.Close
'    End If
'
''振替後品番仕様取得
'    sql = "select "
'    sql = sql & " E021.HWFTYPE HWFTYPE, "
'    sql = sql & " E022.HWFCDIR HWFCDIR, "
'    sql = sql & " E021.HWFDOP HWFDOP, "
'    sql = sql & " E023.HWFCDOP HWFCDOP "
'    sql = sql & " from "
'    sql = sql & " TBCME021 E021, TBCME022 E022, TBCME023 E023 "
'    sql = sql & " where "
'    sql = sql & " E021.HINBAN='" & HIN.HINBAN & "' and "
'    sql = sql & " E021.MNOREVNO=" & HIN.mnorevno & " and "
'    sql = sql & " E021.FACTORY='" & HIN.factory & "' and "
'    sql = sql & " E021.OPECOND='" & HIN.opecond & "' and "
'    sql = sql & " E022.HINBAN='" & HIN.HINBAN & "' and "
'    sql = sql & " E022.MNOREVNO=" & HIN.mnorevno & " and "
'    sql = sql & " E022.FACTORY='" & HIN.factory & "' and "
'    sql = sql & " E022.OPECOND='" & HIN.opecond & "' and "
'    sql = sql & " E023.HINBAN='" & HIN.HINBAN & "' and "
'    sql = sql & " E023.MNOREVNO=" & HIN.mnorevno & " and "
'    sql = sql & " E023.FACTORY='" & HIN.factory & "' and "
'    sql = sql & " E023.OPECOND='" & HIN.opecond & "'"
'
'    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'
'    If rs.RecordCount <> 1 Then
'        '見つからなかったら、FUNCTION_RETURN_FAILUREを返す。
'        judge_ok = False
'        itemNG = "E021"
'        rs.Close
'        SXLPreJudge = FUNCTION_RETURN_FAILURE
'        GoTo PROC_EXIT
'    Else
'        ''見つかったら、最新の品番情報をセットして FUNCTION_RETURN_SUCCESSを返す
'        With Wsi
'            .HWFCDIR = rs("HWFCDIR")
'            .HWFCDOP = rs("HWFCDOP")
'            .HWFTYPE = rs("HWFTYPE")
'            .HWFDOP = rs("HWFDOP")
'        End With
'        rs.Close
'        judge_ok = True
'    End If
'
'Debug.Print "3 振替前仕様'" & mWsi.HWFTYPE & "' '" & mWsi.HWFCDIR & "' '" & mWsi.HWFCDOP & "' '" & mWsi.HWFDOP & "'"
'Debug.Print "4 振替後仕様'" & Wsi.HWFTYPE & "' '" & Wsi.HWFCDIR & "' '" & Wsi.HWFCDOP & "' '" & Wsi.HWFDOP & "'"
'
''・仕様比較（OR）
'    If mWsi.HWFTYPE <> Wsi.HWFTYPE Then
''       SXLPreJudge = FUNCTION_RETURN_FAILURE
'        judge_ok = False
'        itemNG = "タイプ"
'        rs.Close
''        GoTo proc_exit
'    ElseIf mWsi.HWFCDIR <> Wsi.HWFCDIR Then
''        SXLPreJudge = FUNCTION_RETURN_FAILURE
'        judge_ok = False
'        itemNG = "方位"
''        GoTo proc_exit
'    ElseIf mWsi.HWFCDOP <> Wsi.HWFCDOP Then
''        SXLPreJudge = FUNCTION_RETURN_FAILURE
'    If Tokusai = "1" Then Else judge_ok = False
'        itemNG = "結晶ドープ"
''        GoTo proc_exit
'    ElseIf mWsi.HWFDOP <> Wsi.HWFDOP Then
''        SXLPreJudge = FUNCTION_RETURN_FAILURE
'        judge_ok = False
'        itemNG = "ドーパント"
''        GoTo proc_exit
'    Else
'        itemNG = "OK"
'        judge_ok = True
''        SXLPreJudge = FUNCTION_RETURN_SUCCESS
'    End If
'
'Debug.Print "5 判定 '" & judge_ok & "' 項目 '" & itemNG & "'"
'
'    SXLPreJudge = FUNCTION_RETURN_SUCCESS
'
'PROC_EXIT:
'    '終了
'    gErr.Pop
'    Exit Function
'
'PROC_ERR:
'    'エラーハンドラ
'    Debug.Print "====== Error SQL ======"
'    Debug.Print sql
'    gErr.HandleError
'    Resume PROC_EXIT
'
'End Function
'品番の振替チェックは共通関数を使用するので削除する-------end iida 2003/09/03

